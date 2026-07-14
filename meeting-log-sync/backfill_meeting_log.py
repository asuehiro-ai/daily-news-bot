#!/usr/bin/env python3
"""指定期間のGoogleカレンダーとPLAUD録音を突き合わせ、面談ログスプレッドシートに過去分をまとめて記録する。

sync_meeting_log.py（毎朝の前日分自動実行）とは別に、過去分を一括反映したいときに手動実行する。
Slackダイジェスト投稿はしない（過去分をまとめてSlackに流すとノイズになるため）。

使い方: python meeting-log-sync/backfill_meeting_log.py --from 2026-06-01 [--to 2026-07-13] [--dry-run]

必須環境変数: GOOGLE_SERVICE_ACCOUNT_JSON, SPREADSHEET_ID
任意環境変数: GEMINI_API_KEY（未設定時は氏名・役職の分解をせず生テキストのまま記録）,
             PLAUD_TOKEN, PLAUD_BASE_URL（未設定時は議事録内容の突き合わせをスキップ）
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
    append_rows,
    fetch_calendar_events,
    fetch_existing_event_ids,
    match_plaud_summary,
    parse_event_title,
    split_person_title,
)


def daterange(start_str, end_str):
    d = datetime.date.fromisoformat(start_str)
    end = datetime.date.fromisoformat(end_str)
    while d <= end:
        yield d.strftime("%Y-%m-%d")
        d += datetime.timedelta(days=1)


def main():
    if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
        sys.stdout.reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser()
    parser.add_argument("--from", dest="from_date", required=True, help="開始日 YYYY-MM-DD")
    parser.add_argument("--to", dest="to_date", default=None, help="終了日 YYYY-MM-DD（省略時は前日）")
    parser.add_argument("--dry-run", action="store_true", help="スプレッドシートへの書き込みをせずログ出力のみ行う")
    args = parser.parse_args()

    to_date = args.to_date or (datetime.datetime.now(JST) - datetime.timedelta(days=1)).strftime("%Y-%m-%d")

    spreadsheet_id = os.environ.get("SPREADSHEET_ID")
    gemini_key = os.environ.get("GEMINI_API_KEY")
    plaud_token = os.environ.get("PLAUD_TOKEN")
    plaud_base_url = os.environ.get("PLAUD_BASE_URL") or DEFAULT_PLAUD_BASE_URL

    if not spreadsheet_id:
        print("SPREADSHEET_ID が未設定です")
        sys.exit(1)

    print(f"対象期間(JST): {args.from_date} 〜 {to_date}")

    sheets_token = get_access_token([SHEETS_SCOPE], subject=SPREADSHEET_OWNER_EMAIL)
    existing_ids = fetch_existing_event_ids(sheets_token, spreadsheet_id)
    print(f"記録済みイベント数: {len(existing_ids)}件（重複はスキップします）")

    plaud_items = []
    if plaud_token:
        try:
            recordings = fetch_recordings_in_range(plaud_token, plaud_base_url, args.from_date, to_date)
            plaud_items = fetch_all_summary_items(plaud_token, plaud_base_url, recordings)
            print(f"PLAUD録音件数: {len(plaud_items)}件")
        except PlaudTokenExpiredError as e:
            print(str(e))
    else:
        print("PLAUD_TOKEN 未設定 - 議事録の突き合わせをスキップします")

    new_rows = []
    for date_str in daterange(args.from_date, to_date):
        day_count = 0
        for employee in EMPLOYEES:
            cal_token = get_access_token([CALENDAR_READONLY_SCOPE], subject=employee["email"])
            events = fetch_calendar_events(cal_token, employee["email"], date_str)

            for event in events:
                event_id = event.get("id")
                if not event_id or event_id in existing_ids:
                    continue
                start = event.get("start", {})
                if "dateTime" not in start:
                    continue  # 終日予定はスキップ

                parsed = parse_event_title(event.get("summary", ""))
                if not parsed:
                    continue

                meeting_dt = to_jst_datetime(start["dateTime"])
                person_name, title = split_person_title(parsed["person_raw"], gemini_key)

                transcript = ""
                if plaud_items and meeting_dt:
                    matched = match_plaud_summary(meeting_dt, plaud_items)
                    if matched:
                        transcript = clean_summary(matched["summary"])

                new_rows.append(
                    [
                        date_str,
                        employee["name"],
                        parsed["company"],
                        person_name,
                        title,
                        meeting_dt.strftime("%H:%M") if meeting_dt else "",
                        transcript,
                        event_id,
                    ]
                )
                existing_ids.add(event_id)
                day_count += 1

        if day_count:
            print(f"{date_str}: {day_count}件")

    if not new_rows:
        print("記録対象の面談なし")
        return

    if args.dry_run:
        for row in new_rows:
            print("---")
            print(f"{row[0]} {row[2]} / {row[3]}（{row[4]}） {row[5]} 担当:{row[1]}")
            print(f"議事録: {row[6][:80]}{'...' if len(row[6]) > 80 else ''}")
        print(f"[dry-run] 合計{len(new_rows)}件を書き込み予定（実際には書き込みません）")
        return

    if append_rows(sheets_token, spreadsheet_id, new_rows):
        print(f"合計{len(new_rows)}件をスプレッドシートに記録しました")


if __name__ == "__main__":
    main()
