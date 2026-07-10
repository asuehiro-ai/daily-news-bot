#!/usr/bin/env python3
"""面談ログスプレッドシートに、閲覧用タブ「面談ログ（閲覧用）」を追加する一回限りのセットアップスクリプト。

Sheet1（生データ、sync_meeting_log.pyが書き込み続ける）はそのまま維持し、
QUERY関数でcalendar_event_id列を除いた内容を日付降順で表示する閲覧用タブを追加する。
既に存在する場合は何もしない（再実行しても安全）。

事前に GOOGLE_SERVICE_ACCOUNT_JSON・SPREADSHEET_ID 環境変数を設定しておくこと。
"""

import json
import os
import sys
import urllib.error
import urllib.parse
import urllib.request

from employees import SPREADSHEET_OWNER_EMAIL
from google_auth import SHEETS_SCOPE, get_access_token

VIEWER_SHEET_TITLE = "面談ログ（閲覧用）"
HEADER_ROW = ["日付", "担当社員", "取引先名", "面談者名", "役職名", "面談時刻", "議事録内容"]
VIEWER_FORMULA = '=QUERY(Sheet1!A2:H, "select A,B,C,D,E,F,G order by A desc", 0)'


def api_call(url, token, *, method="GET", payload=None):
    data = json.dumps(payload).encode("utf-8") if payload is not None else None
    req = urllib.request.Request(
        url, data=data, method=method, headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return resp.getcode(), json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8")
        return e.code, (json.loads(body) if body else {})


def main():
    if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
        sys.stdout.reconfigure(encoding="utf-8")

    spreadsheet_id = os.environ.get("SPREADSHEET_ID")
    if not spreadsheet_id:
        print("SPREADSHEET_ID が未設定です")
        sys.exit(1)

    token = get_access_token([SHEETS_SCOPE], subject=SPREADSHEET_OWNER_EMAIL)

    code, body = api_call(f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}", token)
    if code != 200:
        print(f"スプレッドシート取得エラー: HTTP {code} {body}")
        sys.exit(1)

    existing_sheet = next(
        (s for s in body.get("sheets", []) if s["properties"]["title"] == VIEWER_SHEET_TITLE), None
    )

    if existing_sheet:
        print(f"「{VIEWER_SHEET_TITLE}」は既に存在します（sheetId={existing_sheet['properties']['sheetId']}）。何もしません。")
        return

    code, body = api_call(
        f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}:batchUpdate",
        token,
        method="POST",
        payload={
            "requests": [
                {
                    "addSheet": {
                        "properties": {
                            "title": VIEWER_SHEET_TITLE,
                            "index": 0,
                            "gridProperties": {"frozenRowCount": 1},
                        }
                    }
                }
            ]
        },
    )
    if code >= 300:
        print(f"シート追加エラー: HTTP {code} {body}")
        sys.exit(1)

    sheet_id = body["replies"][0]["addSheet"]["properties"]["sheetId"]
    print(f"「{VIEWER_SHEET_TITLE}」を追加しました（sheetId={sheet_id}）")

    range_name = urllib.parse.quote(f"'{VIEWER_SHEET_TITLE}'!A1:G2")
    code, body = api_call(
        f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}/values/{range_name}?valueInputOption=USER_ENTERED",
        token,
        method="PUT",
        payload={"values": [HEADER_ROW, [VIEWER_FORMULA]]},
    )
    if code >= 300:
        print(f"ヘッダー・数式の書き込みエラー: HTTP {code} {body}")
        sys.exit(1)

    code, body = api_call(
        f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}:batchUpdate",
        token,
        method="POST",
        payload={
            "requests": [
                {
                    "repeatCell": {
                        "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1},
                        "cell": {"userEnteredFormat": {"textFormat": {"bold": True}}},
                        "fields": "userEnteredFormat.textFormat.bold",
                    }
                }
            ]
        },
    )
    if code >= 300:
        print(f"書式設定エラー: HTTP {code} {body}")
        sys.exit(1)

    print("ヘッダー・QUERY数式・書式設定が完了しました")


if __name__ == "__main__":
    main()
