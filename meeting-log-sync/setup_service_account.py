#!/usr/bin/env python3
"""一度だけローカルで実行するセットアップスクリプト。

面談ログ用の新しいGoogleスプレッドシートを、SPREADSHEET_OWNER_EMAIL（employees.py）に
サービスアカウントがなりすまして（ドメイン全体委任）作成する。作成者本人のアカウントで
直接作られるため、後から共有設定をする必要がない。

事前に GOOGLE_SERVICE_ACCOUNT_JSON 環境変数（サービスアカウントのJSON鍵の中身）を設定し、
ドメイン全体委任のスコープに https://www.googleapis.com/auth/spreadsheets を追加しておくこと。
"""

import json
import sys
import urllib.error
import urllib.request

from employees import SPREADSHEET_OWNER_EMAIL
from google_auth import SHEETS_SCOPE, get_access_token

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
    token = get_access_token([SHEETS_SCOPE], subject=SPREADSHEET_OWNER_EMAIL)

    code, body = api_call(
        "https://sheets.googleapis.com/v4/spreadsheets",
        token,
        method="POST",
        # sheets[].properties.titleを明示しないと、日本語ロケールのアカウントでは
        # 既定のシート名が「シート1」になり、コード側で決め打ちの"Sheet1"と一致しなくなる。
        payload={"properties": {"title": SHEET_TITLE}, "sheets": [{"properties": {"title": "Sheet1"}}]},
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

    print(f"スプレッドシートを作成しました（{SPREADSHEET_OWNER_EMAIL}の所有）: {spreadsheet_url}")
    print()
    print(f"この値をGitHub Secretsに SPREADSHEET_ID として登録してください: {spreadsheet_id}")


if __name__ == "__main__":
    main()
