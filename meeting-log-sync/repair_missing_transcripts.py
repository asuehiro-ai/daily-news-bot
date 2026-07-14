#!/usr/bin/env python3
"""既に記録済みの行のうち、議事録内容（G列）が空欄のものを再マッチングして埋め直す一回限りの修復スクリプト。

sync_meeting_log.py・backfill_meeting_log.pyがPLAUDマッチングを末廣さん限定にしていたバグの修正
（2026-07-13）に伴い、その制限のせいで空欄のまま記録されてしまった行（主に大熊さんの分）を対象に、
同じ突き合わせロジックを再実行して議事録内容を反映する。新規行の追加はせず、G列の更新のみ行う。

使い方: python meeting-log-sync/repair_missing_transcripts.py --from 2026-06-01 --to 2026-07-13 [--dry-run]

必須環境変数: GOOGLE_SERVICE_ACCOUNT_JSON, SPREADSHEET_ID, PLAUD_TOKEN
任意環境変数: PLAUD_BASE_URL
"""

import argparse
import datetime
import os
import sys

from employees import EMPLOYEES, SPREADSHEET_OWNER_EMAIL
from google_auth import CALENDAR_READONLY_SCOPE, SHEETS_SCOPE, get_access_token
from plaud_client import (
    DEFAULT_PLAUD_BASE_URL,
    JST,
    PlaudTokenExpiredError,
    clean_summary,
    fetch_all_summary_items,
    fetch_recordings_in_range,
    to_jst_datetime,
)
from sync_meeting_log import (
    fetch_calendar_events,
    match_plaud_summary,
    parse_event_title,
    sheets_api_call,
)


def daterange(start_str, end_str):
    d = datetime.date.fromisoformat(start_str)
    end = datetime.date.fromisoformat(end_str)
    while d <= end:
        yield d.strftime("%Y-%m-%d")
        d += datetime.timedelta(days=1)


def fetch_existing_rows(token, spreadsheet_id):
    """event_id -> (row_number, current_transcript) の辞書を返す。"""
    code, body = sheets_api_call(
        f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}/values/Sheet1!A2:H", token
    )
    if code != 200:
        print(f"既存データ取得エラー: HTTP {code} {body}")
        return {}
    result = {}
    for i, row in enumerate(body.get("values", [])):
        row_number = i + 2
        event_id = row[7] if len(row) > 7 else None
        transcript = row[6] if len(row) > 6 else ""
        if event_id:
            result[event_id] = (row_number, transcript)
    return result


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
    parser.add_argument("--from", dest="from_date", required=True, help="開始日 YYYY-MM-DD")
    parser.add_argument("--to", dest="to_date", default=None, help="終了日 YYYY-MM-DD（省略時は前日）")
    parser.add_argument("--dry-run", action="store_true", help="更新をせずログ出力のみ行う")
    args = parser.parse_args()

    to_date = args.to_date or (datetime.datetime.now(JST) - datetime.timedelta(days=1)).strftime("%Y-%m-%d")

    spreadsheet_id = os.environ.get("SPREADSHEET_ID")
    plaud_token = os.environ.get("PLAUD_TOKEN")
    plaud_base_url = os.environ.get("PLAUD_BASE_URL") or DEFAULT_PLAUD_BASE_URL

    if not spreadsheet_id:
        print("SPREADSHEET_ID が未設定です")
        sys.exit(1)
    if not plaud_token:
        print("PLAUD_TOKEN が未設定です")
        sys.exit(1)

    print(f"対象期間(JST): {args.from_date} 〜 {to_date}")

    sheets_token = get_access_token([SHEETS_SCOPE], subject=SPREADSHEET_OWNER_EMAIL)
    existing_rows = fetch_existing_rows(sheets_token, spreadsheet_id)
    print(f"既存行数: {len(existing_rows)}件")

    try:
        recordings = fetch_recordings_in_range(plaud_token, plaud_base_url, args.from_date, to_date)
        plaud_items = fetch_all_summary_items(plaud_token, plaud_base_url, recordings)
        print(f"PLAUD録音件数: {len(plaud_items)}件")
    except PlaudTokenExpiredError as e:
        print(str(e))
        sys.exit(1)

    # 元の書き込み時と同じ順序（日付昇順→EMPLOYEES順→カレンダー順）で走査し、
    # 消費済みの録音プールを正しく再現しながら、空欄の行だけ更新対象にする。
    updates = []
    for date_str in daterange(args.from_date, to_date):
        for employee in EMPLOYEES:
            cal_token = get_access_token([CALENDAR_READONLY_SCOPE], subject=employee["email"])
            events = fetch_calendar_events(cal_token, employee["email"], date_str)

            for event in events:
                event_id = event.get("id")
                if not event_id or event_id not in existing_rows:
                    continue
                start = event.get("start", {})
                if "dateTime" not in start:
                    continue
                parsed = parse_event_title(event.get("summary", ""))
                if not parsed:
                    continue

                meeting_dt = to_jst_datetime(start["dateTime"])
                row_number, current_transcript = existing_rows[event_id]

                matched = None
                if plaud_items and meeting_dt:
                    matched = match_plaud_summary(meeting_dt, plaud_items)

                if not current_transcript and matched:
                    transcript = clean_summary(matched["summary"])
                    updates.append((row_number, transcript))
                    print(f"{date_str} {employee['name']} {parsed['company']}: 行{row_number}に議事録を反映")

    if not updates:
        print("反映対象なし")
        return

    if args.dry_run:
        print(f"[dry-run] {len(updates)}件を反映予定（実際には書き込みません）")
        return

    if batch_update_transcripts(sheets_token, spreadsheet_id, updates):
        print(f"{len(updates)}件の議事録内容を反映しました")


if __name__ == "__main__":
    main()
