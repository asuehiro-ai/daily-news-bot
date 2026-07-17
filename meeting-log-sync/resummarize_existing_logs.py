#!/usr/bin/env python3
"""既に記録済みの行のうち、指定日付以前の議事録内容（G列）を、新しい要約ロジック
（summarize_for_log）で書き直す一回限りの修復スクリプト。

PLAUDへの再アクセスは不要。既存のG列テキスト（新ロジック導入前に記録された、
PLAUDの生っぽい要約）をそのまま新しい要約プロンプトに通して上書きする。

使い方: python meeting-log-sync/resummarize_existing_logs.py --until 2026-06-30 [--dry-run]

必須環境変数: GOOGLE_SERVICE_ACCOUNT_JSON, SPREADSHEET_ID, ANTHROPIC_API_KEY
"""

import argparse
import os
import sys

from employees import SPREADSHEET_OWNER_EMAIL
from google_auth import SHEETS_SCOPE, get_access_token
from sync_meeting_log import sheets_api_call, summarize_for_log


def fetch_rows(token, spreadsheet_id):
    code, body = sheets_api_call(
        f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}/values/Sheet1!A2:H", token
    )
    if code != 200:
        print(f"既存データ取得エラー: HTTP {code} {body}")
        return []
    return body.get("values", [])


def batch_update_transcripts(token, spreadsheet_id, updates):
    if not updates:
        return True
    data = [{"range": f"Sheet1!G{row_number}", "values": [[transcript]]} for row_number, transcript in updates]
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}/values:batchUpdate"
    code, body = sheets_api_call(
        url, token, method="POST", payload={"valueInputOption": "USER_ENTERED", "data": data}
    )
    if code >= 300:
        print(f"更新エラー: HTTP {code} {body}")
        return False
    return True


def main():
    if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
        sys.stdout.reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser()
    parser.add_argument("--until", dest="until_date", required=True, help="この日付以前（YYYY-MM-DD、当日を含む）を対象にする")
    parser.add_argument("--dry-run", action="store_true", help="更新をせずログ出力のみ行う")
    args = parser.parse_args()

    spreadsheet_id = os.environ.get("SPREADSHEET_ID")
    anthropic_key = os.environ.get("ANTHROPIC_API_KEY")

    if not spreadsheet_id:
        print("SPREADSHEET_ID が未設定です")
        sys.exit(1)
    if not anthropic_key:
        print("ANTHROPIC_API_KEY が未設定です")
        sys.exit(1)

    print(f"対象: {args.until_date} 以前の全行")

    sheets_token = get_access_token([SHEETS_SCOPE], subject=SPREADSHEET_OWNER_EMAIL)
    rows = fetch_rows(sheets_token, spreadsheet_id)
    print(f"シート行数: {len(rows)}件")

    updates = []
    for i, row in enumerate(rows):
        row_number = i + 2
        date_str = row[0] if len(row) > 0 else ""
        transcript = row[6] if len(row) > 6 else ""
        if not date_str or date_str > args.until_date:
            continue
        if not transcript:
            continue

        new_summary = summarize_for_log(transcript, anthropic_key)
        if new_summary and new_summary != transcript:
            updates.append((row_number, new_summary))
            print(f"{date_str} 行{row_number}: 要約を更新予定")

    print(f"更新対象: {len(updates)}件")

    if args.dry_run:
        for row_number, new_summary in updates[:5]:
            print("---")
            print(f"行{row_number}: {new_summary}")
        print(f"[dry-run] 合計{len(updates)}件を反映予定（実際には書き込みません）")
        return

    if batch_update_transcripts(sheets_token, spreadsheet_id, updates):
        print(f"{len(updates)}件の議事録内容を新しい要約に更新しました")


if __name__ == "__main__":
    main()
