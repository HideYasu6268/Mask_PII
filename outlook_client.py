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

    戻り値: {"subject": str, "sender": str, "body": str}
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
