# -*- coding: utf-8 -*-
"""
outlook_client.py

ローカルにインストール・起動済みのOutlookデスクトップアプリを、pywin32(win32com)
経由でCOM操作し、受信トレイの未読メールを取得する。Outlook自体が行う送受信を
除けば、このモジュール自身は外部通信を一切行わない(Windows専用)。
"""

from __future__ import annotations

import threading

# Exchangeアカウントだと SenderEmailAddress や Recipient.Address が素のSMTP
# アドレスではなく内部形式(Exchange DN、"/O=..."等)を返すことがあるため、
# その場合のみ PropertyAccessor で実際のSMTPアドレスを取り直す。
_PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"

# 「①メール取得」の相手先へ過去に送った最新メールを、送信済みフォルダの
# 新しい順に何件まで遡って探すか。多くしすぎるとCOM経由の逐次アクセスで
# 時間がかかるため、上限を設けて「見つからなければ諦める」設計にしている。
DEFAULT_SENT_SCAN_LIMIT = 200

# 過去の送信済みメールから拾う「宛名・挨拶」として扱う先頭の行数。
GREETING_LINE_COUNT = 7


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


def _find_latest_sent_greeting(
    namespace, sender_email: str, max_scan: int = DEFAULT_SENT_SCAN_LIMIT
) -> str | None:
    """送信済みメールを送信日時の新しい順に最大max_scan件まで見て、
    sender_email宛てに送った最新のメールを探し、本文の先頭GREETING_LINE_COUNT行
    (宛名・挨拶部分だと想定)を返す。見つからなければNone
    (呼び出し元は「宛名・挨拶なし」にフォールバックすること)。
    """
    target = sender_email.strip().lower()
    if not target:
        return None

    sent_folder = namespace.GetDefaultFolder(5)  # 5 = olFolderSentMail
    items = sent_folder.Items
    items.Sort("[SentOn]", True)  # 新しい順

    scanned = 0
    item = items.GetFirst()
    while item is not None and scanned < max_scan:
        scanned += 1
        try:
            recipients = item.Recipients
            matched = any(
                _get_recipient_smtp_address(recipients.Item(i)).strip().lower() == target
                for i in range(1, recipients.Count + 1)
            )
            if matched:
                body = str(item.Body or "")
                greeting = "\n".join(body.splitlines()[:GREETING_LINE_COUNT]).strip()
                return greeting or None
        except Exception:  # noqa: BLE001
            pass  # 1件の読み取りに失敗しても、全体を諦めずに次へ進む
        item = items.GetNext()
    return None


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
    html_fragment = escaped.replace("\n", "<br>\n") + "<br><br>\n"

    original_html = reply.HTMLBody
    match_pos = original_html.lower().find("<body")
    if match_pos == -1:
        # <body>タグが見つからない異常なケースへのフォールバック。
        reply.Body = f"{reply_body}\n\n{reply.Body}"
        return

    tag_end = original_html.find(">", match_pos) + 1
    reply.HTMLBody = original_html[:tag_end] + html_fragment + original_html[tag_end:]


def get_latest_unread_email() -> dict | None:
    """Outlookの受信トレイから、受信日時が最も新しい未読メール1件を取得する。
    取得したメールは既読に更新する(同じメールを繰り返し取得しないようにするため)。

    戻り値: {"subject": str, "sender": str, "sender_email": str, "body": str,
             "received": str, "entry_id": str, "store_id": str,
             "greeting": str | None}
    entry_id/store_idは、後から create_quoted_reply() でこの同じメールを
    Outlook側から再度特定し、Outlook標準の引用返信を作成するために使う。
    greetingは、差出人(sender_email)へ過去に送った送信済みメールのうち
    最新のものの先頭GREETING_LINE_COUNT行(宛名・挨拶だと想定)。見つからない場合はNone
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
            items = inbox.Items
            items.Sort("[ReceivedTime]", True)  # 新しい順に並べ替え
            unread_items = items.Restrict("[Unread] = true")
            if unread_items.Count == 0:
                return None
            latest = unread_items.GetFirst()
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
                    _find_latest_sent_greeting(namespace, sender_email) if sender_email else None
                )
            except Exception:  # noqa: BLE001
                result["greeting"] = None

            # 取得したメールは未読のまま残さず、既読に更新しておく
            # (次回「①メール取得」を押したときに同じメールを再取得しないようにするため)。
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
