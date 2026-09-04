# -*- coding: utf-8 -*-
"""
outlook_client.py

ローカルにインストール・起動済みのOutlookデスクトップアプリを、pywin32(win32com)
経由でCOM操作し、受信トレイの未読メールを取得する。Outlook自体が行う送受信を
除けば、このモジュール自身は外部通信を一切行わない(Windows専用)。
"""

from __future__ import annotations

import threading


class OutlookError(RuntimeError):
    pass


def get_latest_unread_email() -> dict | None:
    """Outlookの受信トレイから、受信日時が最も新しい未読メール1件を取得する。
    取得したメールは既読に更新する(同じメールを繰り返し取得しないようにするため)。

    戻り値: {"subject": str, "sender": str, "body": str, "received": str,
             "entry_id": str, "store_id": str}
    entry_id/store_idは、後から create_quoted_reply() でこの同じメールを
    Outlook側から再度特定し、Outlook標準の引用返信を作成するために使う。
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
            result = {
                "subject": str(latest.Subject or ""),
                "sender": str(getattr(latest, "SenderName", "") or ""),
                "body": str(latest.Body or ""),
                # ReceivedTimeはCOM経由の独自の日時型なのでstr()でそのまま文字列化する。
                "received": str(getattr(latest, "ReceivedTime", "") or ""),
                # 後で create_quoted_reply() がこのメールをOutlook側から
                # 再度特定するためのID(Parent = このメールが入っているフォルダ)。
                "entry_id": str(latest.EntryID),
                "store_id": str(latest.Parent.StoreID),
            }
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
            reply.Body = f"{reply_body}\n\n{reply.Body}"
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
