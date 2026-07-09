#!/usr/bin/env python3
"""一度だけローカルで実行するセットアップスクリプト。

面談ログ用の新しいGoogleスプレッドシートをサービスアカウントの権限で作成し、
指定したGoogleアカウントに編集者として共有する。

事前に GOOGLE_SERVICE_ACCOUNT_JSON 環境変数（サービスアカウントのJSON鍵の中身）を
設定してから実行すること。
"""

import json
import os
import sys
import urllib.error
import urllib.request

from google_auth import DRIVE_SCOPE, SHEETS_SCOPE, get_access_token

SHARE_WITH_EMAIL = os.environ.get("MEETING_LOG_OWNER_EMAIL", "a.suehiro@le-gr.co.jp")
SHEET_TITLE = "面談ログ"
HEADER_ROW = ["日付", "担当社員", "取引先名", "面談者名", "役職名", "面談時刻", "議事録内容", "calendar_event_id"]


def api_call(url, token, *, method="GET", payload=None):
    data = json.dumps(payload).encode("utf-8") if payload is not None else None
    req = urllib.request.Request(
        url,
        data=data,
        method=method,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return resp.getcode(), json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8")
        return e.code, (json.loads(body) if body else {})


def main():
    token = get_access_token([SHEETS_SCOPE, DRIVE_SCOPE])

    code, body = api_call(
        "https://sheets.googleapis.com/v4/spreadsheets",
        token,
        method="POST",
        payload={"properties": {"title": SHEET_TITLE}},
    )
    if code >= 300:
        print(f"スプレッドシート作成に失敗: HTTP {code} {body}")
        sys.exit(1)

    spreadsheet_id = body["spreadsheetId"]
    spreadsheet_url = body["spreadsheetUrl"]

    code, body = api_call(
        f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}/values/Sheet1!A1:H1?valueInputOption=RAW",
        token,
        method="PUT",
        payload={"values": [HEADER_ROW]},
    )
    if code >= 300:
        print(f"ヘッダー行の書き込みに失敗: HTTP {code} {body}")
        sys.exit(1)

    code, body = api_call(
        f"https://www.googleapis.com/drive/v3/files/{spreadsheet_id}/permissions",
        token,
        method="POST",
        payload={"type": "user", "role": "writer", "emailAddress": SHARE_WITH_EMAIL},
    )
    if code >= 300:
        print(f"共有設定に失敗: HTTP {code} {body}")
        sys.exit(1)

    print(f"スプレッドシートを作成しました: {spreadsheet_url}")
    print(f"{SHARE_WITH_EMAIL} に編集者として共有しました。")
    print()
    print(f"この値をGitHub Secretsに SPREADSHEET_ID として登録してください: {spreadsheet_id}")


if __name__ == "__main__":
    main()
