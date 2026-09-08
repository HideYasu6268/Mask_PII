# -*- coding: utf-8 -*-
"""
outlook_client.py

ローカルにインストール・起動済みのOutlookデスクトップアプリを、pywin32(win32com)
経由でCOM操作し、受信トレイ(および顧客ごとの仕分けフォルダ等、その配下の
すべてのサブフォルダ)の未読メールを取得する。Outlook自体が行う送受信を
除けば、このモジュール自身は外部通信を一切行わない(Windows専用)。
"""

from __future__ import annotations

import re
import threading

# Exchangeアカウントだと SenderEmailAddress や Recipient.Address が素のSMTP
# アドレスではなく内部形式(Exchange DN、"/O=..."等)を返すことがあるため、
# その場合のみ PropertyAccessor で実際のSMTPアドレスを取り直す。
_PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"

# 「メール取得」の相手先へ過去に送った最新メールを、送信済みフォルダの
# 新しい順に何件まで遡って探すか。多くしすぎるとCOM経由の逐次アクセスで
# 時間がかかるため、上限を設けて「見つからなければ諦める」設計にしている。
DEFAULT_SENT_SCAN_LIMIT = 200

# 過去の送信済みメールから拾う「宛名・挨拶」として扱う先頭の行数。
GREETING_LINE_COUNT = 10


class OutlookError(RuntimeError):
    pass


def _get_sender_smtp_address(mail_item) -> str:
    """MailItemの差出人の実際のSMTPアドレスを取り出す(ベストエフォート)。"""
    addr = str(getattr(mail_item, "SenderEmailAddress", "") or "")
    if "@" in addr:
        return addr
    try:
        smtp = mail_item.Sender.PropertyAccessor.GetProperty(_PR_SMTP_ADDRESS)
        return str(smtp or "")
    except Exception:  # noqa: BLE001
        return addr


def _get_recipient_smtp_address(recipient) -> str:
    """Recipientの実際のSMTPアドレスを取り出す(ベストエフォート)。"""
    addr = str(getattr(recipient, "Address", "") or "")
    if "@" in addr:
        return addr
    try:
        smtp = recipient.PropertyAccessor.GetProperty(_PR_SMTP_ADDRESS)
        return str(smtp or "")
    except Exception:  # noqa: BLE001
        return addr


def _get_current_user_smtp_address(namespace) -> str:
    """現在ログイン中の自分自身のSMTPアドレスを取り出す(ベストエフォート)。"""
    try:
        return _get_recipient_smtp_address(namespace.CurrentUser).strip().lower()
    except Exception:  # noqa: BLE001
        return ""


def _scan_folder_for_sent_item(folder, target: str, my_address: str, max_scan: int):
    """folder(単一フォルダ、非再帰)を送信日時の新しい順に最大max_scan件まで見て、
    target宛てに自分から送った最新のメールを探す。(item, SentOn) のタプル、
    見つからなければNoneを返す。

    my_addressを指定した場合、差出人が自分自身でないアイテムは除外する
    (受信メールと自分の送信控えが混在するフォルダを検索する際、受信メールを
    誤って「送った相手」として拾わないようにするため)。
    """
    try:
        items = folder.Items
        items.Sort("[SentOn]", True)  # 新しい順
    except Exception:  # noqa: BLE001
        return None

    scanned = 0
    item = items.GetFirst()
    while item is not None and scanned < max_scan:
        scanned += 1
        try:
            if my_address and _get_sender_smtp_address(item).strip().lower() != my_address:
                item = items.GetNext()
                continue
            recipients = item.Recipients
            matched = any(
                _get_recipient_smtp_address(recipients.Item(i)).strip().lower() == target
                for i in range(1, recipients.Count + 1)
            )
            if matched:
                return item, getattr(item, "SentOn", None)
        except Exception:  # noqa: BLE001
            pass  # 1件の読み取りに失敗しても、全体を諦めずに次へ進む
        item = items.GetNext()
    return None


def _find_latest_sent_greeting(
    namespace, sender_email: str, source_folder=None, max_scan: int = DEFAULT_SENT_SCAN_LIMIT
) -> str | None:
    """sender_email宛てに自分から送った最新のメールを探し、本文の先頭
    GREETING_LINE_COUNT行(宛名・挨拶部分だと想定)を返す。見つからなければNone
    (呼び出し元は「宛名・挨拶なし」にフォールバックすること)。

    既定の送信済みフォルダに加えて、source_folder(通常はメール取得元の
    フォルダ)も検索対象にする。受信トレイ配下に顧客ごとの仕分けフォルダを
    作り、そこに受信メールだけでなく自分の送信控えも一緒に移して管理する
    運用があるため。両方で見つかった場合はより新しい方を採用する。
    """
    target = sender_email.strip().lower()
    if not target:
        return None

    sent_folder = namespace.GetDefaultFolder(5)  # 5 = olFolderSentMail
    candidates = []

    found = _scan_folder_for_sent_item(sent_folder, target, my_address="", max_scan=max_scan)
    if found is not None:
        candidates.append(found)

    if source_folder is not None:
        try:
            same_folder = source_folder.EntryID == sent_folder.EntryID
        except Exception:  # noqa: BLE001
            same_folder = False
        if not same_folder:
            my_address = _get_current_user_smtp_address(namespace)
            found = _scan_folder_for_sent_item(
                source_folder, target, my_address=my_address, max_scan=max_scan
            )
            if found is not None:
                candidates.append(found)

    if not candidates:
        return None

    best_item, best_sent_on = candidates[0]
    for item, sent_on in candidates[1:]:
        if best_sent_on is None or (sent_on is not None and sent_on > best_sent_on):
            best_item, best_sent_on = item, sent_on

    body = str(best_item.Body or "")
    # OutlookのプレーンテキストBodyは、HTML本文の段落と段落の間に
    # (中身が無い)空の段落を挟んでいることが多く、それが変換時に
    # 「半角スペース1文字だけの行」として出力される。単純に行末の
    # 空白を削るだけ(rstrip)ではこの行自体は残ってしまうため、
    # 前後の空白を取ってから中身が空になった行はまるごと除外する。
    lines = [line.strip() for line in body.splitlines()[:GREETING_LINE_COUNT]]
    lines = [line for line in lines if line]
    greeting = "\n".join(lines).strip()
    return greeting or None


def _insert_reply_body(reply, reply_body: str) -> None:
    """Reply()で作成した返信アイテムの本文の先頭に、reply_body(匿名化解除後の
    返信文、宛名・挨拶・署名込み)を差し込む。

    reply.Body(プレーンテキスト)に直接書き込むと、返信が既定のHTML形式の場合
    Outlookがこちらの差し込み分だけ独自のデフォルト書式で包んでしまい、
    Reply()が自動生成した引用部分(元のHTML書式)と見た目が食い違う
    (先頭とそれ以外で書式が変わって見える)。そのため、HTML形式の場合は
    reply.HTMLBody の <body> タグの直後に差し込み、引用部分と同じ既定書式を
    継承させる。HTML形式でない場合(プレーンテキスト/リッチテキスト)は
    フォーマットの継承を気にする必要が無いため、reply.Body への追記のままでよい。

    OutlookがWord編集エンジンで作成したHTML本文は、<body>タグ自体には
    フォント指定が無く、各段落に class="MsoNormal" を付与することでスタイル
    (游ゴシック等)を当てる作りになっている。素のテキストを<body>直後にそのまま
    差し込むだけだとこのクラスが付かず、ブラウザ既定フォント(Times New Roman相当)
    になって残りの部分と食い違うため、MsoNormalクラスが定義されている場合は
    そのクラスを使って同じ見た目に揃える。

    改行の表現方法も重要: 当初は全行を1つの段落(<p>/<div>)にまとめ、行内改行
    (<br>、Wordの書式記号では"↓")で区切っていたが、実機で確認したところ、
    このWordスタイルでは行内改行(↓)の行間が、本物の段落区切り(Enterによる
    ¶、Wordの書式記号では"↵")の行間より明らかに広く、引用部分(本物の段落の
    連続)と比べて間延びして見えることが分かった。そのため、行ごとに独立した
    <p class="MsoNormal">(本物の段落)として差し込み、¶と同じ行間になるようにする。

    また、Word文書内の実際の段落は、クラス自体のmargin指定(0mm)とは別に、
    style属性で margin-bottom(通常12.0pt等)を個別に指定していることが多い。
    差し込む各段落にもこれが無いと段落間隔が周囲と食い違うため、元の文書から
    実際に使われているmargin-bottom値を拾えれば、それを各段落にも適用する。

    空行は <p class="MsoNormal">&nbsp;</p> として表現する(空の<p></p>は
    平文変換時に消えてしまい、行そのものが無かったことになるため)。これは
    Word自身が空の段落をHTMLとして書き出す際の標準的な表現方法でもある。
    """
    try:
        body_format = reply.BodyFormat
    except Exception:  # noqa: BLE001
        body_format = None

    if body_format != 2:  # 2 = olFormatHTML以外は従来通りプレーンテキストで追記
        reply.Body = f"{reply_body}\n\n{reply.Body}"
        return

    escaped = (
        reply_body.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )

    original_html = reply.HTMLBody
    if "MsoNormal" in original_html:
        # 実際の段落(<p class=MsoNormal ... style='...margin-bottom:12.0pt...'>)から
        # margin-bottomの値を拾えれば、差し込む各段落にも同じ値を指定して段落間隔を
        # 揃える(見つからなければ指定なしのまま)。
        margin_match = re.search(
            r'<p\b[^>]*\bclass=(?:"MsoNormal"|MsoNormal)\b[^>]*\bstyle=(["\'])'
            r'[^"\']*?margin-bottom:\s*([^;"\']+)',
            original_html,
        )
        margin_style = f"margin-bottom:{margin_match.group(2)};" if margin_match else ""
        style_attr = f' style="{margin_style}font-size:11.0pt"'
        lines = escaped.split("\n")
        html_fragment = "".join(
            f'<p class="MsoNormal"{style_attr}>{line or "&nbsp;"}</p>' for line in lines
        )
        html_fragment += f'<p class="MsoNormal"{style_attr}>&nbsp;</p>'
    else:
        # MsoNormalクラスが無い場合、フォント指定をしないとブラウザ既定フォント
        # (Times New Roman相当)になり、Outlookの本文フォント設定(游ゴシック等)と
        # 食い違って見えるため、明示的にフォントを指定しておく。
        # 単に font-family:游ゴシック とだけ書くと、Outlook上では「游ゴシック」
        # という固定フォント指定として扱われ、フォントリボンには「游ゴシック
        # (本文のフォント)」ではなく素の「游ゴシック」と表示されてしまう
        # (テーマのフォントを変更してもこの部分だけ追従しない)。
        # Wordは mso-*-theme-font:minor-latin/minor-fareast 等の指定がある
        # 要素を「本文のフォント(テーマ追従)」として扱うため、これを併記する。
        theme_font_style = (
            "font-size:11.0pt;"
            "font-family:'游ゴシック',sans-serif;"
            "mso-ascii-font-family:游ゴシック;mso-ascii-theme-font:minor-latin;"
            "mso-fareast-font-family:游ゴシック;mso-fareast-theme-font:minor-fareast;"
            "mso-hansi-font-family:游ゴシック;mso-hansi-theme-font:minor-latin;"
            "mso-bidi-font-family:游ゴシック;mso-bidi-theme-font:minor-bidi;"
        )
        html_fragment = (
            f'<div style="{theme_font_style}">'
            + escaped.replace("\n", "<br>\n")
            + "<br><br>\n</div>"
        )
    match_pos = original_html.lower().find("<body")
    if match_pos == -1:
        # <body>タグが見つからない異常なケースへのフォールバック。
        reply.Body = f"{reply_body}\n\n{reply.Body}"
        return

    tag_end = original_html.find(">", match_pos) + 1
    reply.HTMLBody = original_html[:tag_end] + html_fragment + original_html[tag_end:]


def _iter_unread_mail_items(folder):
    """folder自身と、その配下のすべてのサブフォルダ(再帰的)から、未読アイテムを
    yieldする。受信トレイ配下にOutlookのルール等で作られた顧客ごとの仕分け
    フォルダに振り分けられたメールも拾えるようにするため。

    フォルダ単位で失敗しても(検索フォルダ等の特殊フォルダで起こりうる)、
    そのフォルダだけスキップして探索全体は続行する。
    """
    try:
        unread_items = folder.Items.Restrict("[Unread] = true")
        for i in range(1, unread_items.Count + 1):
            yield unread_items.Item(i)
    except Exception:  # noqa: BLE001
        pass

    try:
        subfolders = folder.Folders
        for i in range(1, subfolders.Count + 1):
            yield from _iter_unread_mail_items(subfolders.Item(i))
    except Exception:  # noqa: BLE001
        pass


def _find_latest_unread_mail(folder):
    """folder以下(自身+再帰的にすべてのサブフォルダ)の未読アイテムのうち、
    ReceivedTimeが最も新しいものを1件返す。1件も無ければNone。
    """
    latest = None
    latest_time = None
    for item in _iter_unread_mail_items(folder):
        try:
            received = item.ReceivedTime
        except Exception:  # noqa: BLE001
            continue  # 1件の読み取りに失敗しても、全体を諦めずに次へ進む
        if latest_time is None or received > latest_time:
            latest = item
            latest_time = received
    return latest


def get_latest_unread_email() -> dict | None:
    """Outlookの受信トレイ(および、顧客ごとの仕分けフォルダ等その配下の
    すべてのサブフォルダ)から、受信日時が最も新しい未読メール1件を取得する。
    取得したメールは既読に更新する(同じメールを繰り返し取得しないようにするため)。

    戻り値: {"subject": str, "sender": str, "sender_email": str, "body": str,
             "received": str, "entry_id": str, "store_id": str,
             "greeting": str | None}
    entry_id/store_idは、後から create_quoted_reply() でこの同じメールを
    Outlook側から再度特定し、Outlook標準の引用返信を作成するために使う。
    greetingは、差出人(sender_email)へ過去に送った最新のメール(既定の送信済み
    フォルダ、および取得元フォルダ自体に送信控えが移されている場合はそちらも
    対象)の先頭GREETING_LINE_COUNT行(宛名・挨拶だと想定)。見つからない場合はNone
    (ベストエフォートのため、取得できなくてもメール取得自体は失敗させない)。
    未読メールが1件も無い場合は None を返す。
    Outlookが未インストール/未起動、またはCOM操作に失敗した場合は OutlookError。
    """
    try:
        import pythoncom
        import win32com.client
    except ImportError as e:
        raise OutlookError(
            "pywin32 がインストールされていません。 `pip install pywin32` を実行してください。"
        ) from e

    # COMはこの呼び出しスレッド単位で初期化が必要(別スレッドから呼ぶ前提のため)。
    pythoncom.CoInitialize()
    try:
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
        except Exception as e:  # noqa: BLE001
            raise OutlookError(
                f"Outlookに接続できませんでした。Outlookが起動しているか確認してください: {e}"
            ) from e

        try:
            inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
            latest = _find_latest_unread_mail(inbox)
            if latest is None:
                return None
            sender_email = _get_sender_smtp_address(latest)
            result = {
                "subject": str(latest.Subject or ""),
                "sender": str(getattr(latest, "SenderName", "") or ""),
                "sender_email": sender_email,
                "body": str(latest.Body or ""),
                # ReceivedTimeはCOM経由の独自の日時型なのでstr()でそのまま文字列化する。
                "received": str(getattr(latest, "ReceivedTime", "") or ""),
                # 後で create_quoted_reply() がこのメールをOutlook側から
                # 再度特定するためのID(Parent = このメールが入っているフォルダ)。
                "entry_id": str(latest.EntryID),
                "store_id": str(latest.Parent.StoreID),
            }

            # 過去にこの相手(sender_email)へ送った直近のメールから、宛名・挨拶
            # (先頭GREETING_LINE_COUNT行)を拾えれば、返信作成時に再利用する(main.py側)。
            # あくまで補助情報なので、取得に失敗してもメール取得自体は失敗させない。
            try:
                result["greeting"] = (
                    _find_latest_sent_greeting(namespace, sender_email, source_folder=latest.Parent)
                    if sender_email
                    else None
                )
            except Exception:  # noqa: BLE001
                result["greeting"] = None

            # 取得したメールは未読のまま残さず、既読に更新しておく
            # (次回「メール取得」を押したときに同じメールを再取得しないようにするため)。
            latest.UnRead = False
            latest.Save()
            return result
        except Exception as e:  # noqa: BLE001
            raise OutlookError(f"未読メールの取得に失敗しました: {e}") from e
    finally:
        pythoncom.CoUninitialize()


def create_quoted_reply(entry_id: str, store_id: str, reply_body: str) -> None:
    """指定したメール(entry_id/store_idで特定)に対して、Outlook標準の「返信」
    (MailItem.Reply())を作成する。Reply()自体がOutlookのいつもの引用形式
    (差出人/送信日時/件名 + 原文本文)を自動生成するので、その本文の先頭に
    reply_body(匿名化解除後のAI返信案)を差し込んだ状態でOutlookの作成画面を開く。

    ここでは Display() のみ行い、送信は一切しない(内容の確認・編集・送信は
    必ずユーザー自身がOutlook上で行う)。
    """
    try:
        import pythoncom
        import win32com.client
    except ImportError as e:
        raise OutlookError(
            "pywin32 がインストールされていません。 `pip install pywin32` を実行してください。"
        ) from e

    pythoncom.CoInitialize()
    try:
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
        except Exception as e:  # noqa: BLE001
            raise OutlookError(
                f"Outlookに接続できませんでした。Outlookが起動しているか確認してください: {e}"
            ) from e

        try:
            mail_item = namespace.GetItemFromID(entry_id, store_id)
        except Exception as e:  # noqa: BLE001
            raise OutlookError(
                f"元のメールが見つかりませんでした(削除・移動された可能性があります): {e}"
            ) from e

        try:
            reply = mail_item.Reply()
            _insert_reply_body(reply, reply_body)
            reply.Display()
        except Exception as e:  # noqa: BLE001
            raise OutlookError(f"引用返信の作成に失敗しました: {e}") from e
    finally:
        pythoncom.CoUninitialize()


def create_quoted_reply_async(entry_id: str, store_id: str, reply_body: str, on_done, on_error):
    """GUIから呼ぶ用のヘルパー。別スレッドでOutlookを操作し、完了をコールバックで通知する。

    on_done()      : 成功時(Outlookの作成画面を開けた時点)に引数なしで呼ばれる
    on_error(msg)  : 失敗時にエラーメッセージ文字列を渡して呼ばれる
    呼び出し元(GUI)は on_done/on_error の中で self.after(0, ...) を使って
    UIスレッドに処理を戻すこと(他の関数と同じ規約)。
    """
    def worker():
        try:
            create_quoted_reply(entry_id, store_id, reply_body)
            on_done()
        except OutlookError as e:
            on_error(str(e))

    threading.Thread(target=worker, daemon=True).start()


def fetch_latest_unread_email_async(on_done, on_error):
    """GUIから呼ぶ用のヘルパー。別スレッドでOutlookを操作し、完了をコールバックで通知する。

    on_done(mail: dict | None) : 成功時に呼ばれる(未読が無ければ mail は None)
    on_error(msg: str)         : 失敗時にエラーメッセージ文字列を渡して呼ばれる
    呼び出し元(GUI)は on_done/on_error の中で self.after(0, ...) を使って
    UIスレッドに処理を戻すこと(local_llm.preload_model_asyncと同じ規約)。
    """
    def worker():
        try:
            mail = get_latest_unread_email()
            on_done(mail)
        except OutlookError as e:
            on_error(str(e))

    threading.Thread(target=worker, daemon=True).start()
