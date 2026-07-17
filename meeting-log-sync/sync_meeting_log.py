#!/usr/bin/env python3
"""前日の面談（Googleカレンダー）とPLAUD録音を突き合わせ、面談ログスプレッドシートに記録する。
PLAUDの自動要約は話者区別が不正確で長くなりがちなため、Claudeで「会議に参加していない人にも
伝わるレベル」の要約（相手企業の状況・LEG側の説明・決定事項・次回アクション）に書き直して記録する。

必須環境変数: GOOGLE_SERVICE_ACCOUNT_JSON, SPREADSHEET_ID
任意環境変数: GEMINI_API_KEY（未設定時は氏名・役職の分解をせず生テキストのまま記録）,
             PLAUD_TOKEN, PLAUD_BASE_URL（未設定時は議事録内容の突き合わせをスキップ）,
             ANTHROPIC_API_KEY（未設定時は要約をローカル抽出のみで代用）,
             SLACK_WEBHOOK_URL（未設定時はエラー通知をスキップ。PLAUDトークン期限切れの通知にのみ使用）

対象社員・PLAUD突き合わせ対象は employees.py を参照。
予定タイトルの書式は meeting-briefing-bot.gs と同じ「【Web】●●株式会社 山田様」形式。
"""

import argparse
import json
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request

from employees import EMPLOYEES, SPREADSHEET_OWNER_EMAIL
from google_auth import CALENDAR_READONLY_SCOPE, SHEETS_SCOPE, get_access_token
from plaud_client import (
    DEFAULT_PLAUD_BASE_URL,
    PlaudTokenExpiredError,
    TOKEN_EXPIRED_MESSAGE,
    alert_slack,
    build_log_summary_prompt,
    claude_summarize,
    clean_summary,
    extract_key_points,
    fetch_all_summary_items,
    fetch_recordings_by_date,
    gemini_generate,
    get_yesterday_jst_str,
    to_jst_datetime,
)

TITLE_PATTERN = re.compile(r"^【([^】]*)】\s*([^\s（(]+)\s*(.*)$")
MAX_MATCH_DIFF_MINUTES = 90


# ── Sheets API ────────────────────────────────────────────────


def sheets_api_call(url, token, *, method="GET", payload=None):
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


def fetch_existing_event_ids(token, spreadsheet_id):
    code, body = sheets_api_call(
        f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}/values/Sheet1!H2:H", token
    )
    if code != 200:
        print(f"既存データ取得エラー: HTTP {code} {body}")
        return set()
    return {row[0] for row in body.get("values", []) if row}


def append_rows(token, spreadsheet_id, rows):
    url = (
        f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}"
        "/values/Sheet1!A:H:append?valueInputOption=USER_ENTERED"
    )
    code, body = sheets_api_call(url, token, method="POST", payload={"values": rows})
    if code >= 300:
        print(f"スプレッドシート書き込みエラー: HTTP {code} {body}")
        return False
    return True


# ── Calendar API ─────────────────────────────────────────────


def fetch_calendar_events(token, calendar_id, date_str):
    time_min = urllib.parse.quote(f"{date_str}T00:00:00+09:00")
    time_max = urllib.parse.quote(f"{date_str}T23:59:59+09:00")
    encoded_id = urllib.parse.quote(calendar_id, safe="")
    url = (
        f"https://www.googleapis.com/calendar/v3/calendars/{encoded_id}/events"
        f"?timeMin={time_min}&timeMax={time_max}&singleEvents=true&orderBy=startTime"
    )
    code, body = sheets_api_call(url, token)
    if code != 200:
        print(f"カレンダー取得エラー({calendar_id}): HTTP {code} {body}")
        return []
    return body.get("items", [])


def parse_event_title(title):
    m = TITLE_PATTERN.match(title or "")
    if not m:
        return None
    return {"tag": m.group(1).strip(), "company": m.group(2).strip(), "person_raw": m.group(3).strip()}


# ── 氏名・役職の分解（Gemini） ────────────────────────────────


def split_person_title(person_raw, api_key):
    if not person_raw:
        return "", ""
    if not api_key:
        return person_raw, ""

    prompt = (
        "次のテキストは、面談相手の氏名と役職・敬称が含まれた短い文字列です。"
        "氏名部分と役職・敬称部分に分解し、次のJSON形式で1行だけ出力してください（説明文は不要）。\n"
        '{"name": "氏名", "title": "役職・敬称（不明なら空文字）"}\n\n'
        f"テキスト: {person_raw}"
    )
    text = gemini_generate(prompt, api_key, max_output_tokens=200, temperature=0.0)
    if not text:
        return person_raw, ""

    cleaned = re.sub(r"^```(?:json)?|```$", "", text.strip(), flags=re.MULTILINE).strip()
    try:
        data = json.loads(cleaned)
        return data.get("name") or person_raw, data.get("title") or ""
    except json.JSONDecodeError:
        return person_raw, ""


# ── PLAUDとのマッチング ───────────────────────────────────────


def match_plaud_summary(meeting_dt, plaud_items):
    """開始時刻が最も近い録音を取り出して返す（一致したものはリストから取り除く）。"""
    best_index = None
    best_diff = None
    for i, item in enumerate(plaud_items):
        item_dt = to_jst_datetime(item.get("createTime"))
        if not item_dt:
            continue
        diff_minutes = abs((item_dt - meeting_dt).total_seconds()) / 60
        if diff_minutes <= MAX_MATCH_DIFF_MINUTES and (best_diff is None or diff_minutes < best_diff):
            best_index = i
            best_diff = diff_minutes

    return plaud_items.pop(best_index) if best_index is not None else None


# ── 面談ログ要約 ────────────────────────────────────────────────


def summarize_for_log(cleaned_summary, anthropic_api_key):
    if not cleaned_summary:
        return cleaned_summary
    if anthropic_api_key:
        summary = claude_summarize(build_log_summary_prompt(cleaned_summary), anthropic_api_key)
        if summary:
            # プロンプトで禁止していても「# 要約」等のタイトル行が付くことがあるため機械的に除去する
            return re.sub(r"^#{1,3}\s*\S.*\n+", "", summary.strip(), count=1)
    return extract_key_points(cleaned_summary)


# ── メイン処理 ────────────────────────────────────────────────


def main():
    if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
        sys.stdout.reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true", help="スプレッドシートへの書き込みをせずログ出力のみ行う")
    args = parser.parse_args()

    spreadsheet_id = os.environ.get("SPREADSHEET_ID")
    gemini_key = os.environ.get("GEMINI_API_KEY")
    anthropic_key = os.environ.get("ANTHROPIC_API_KEY")
    plaud_token = os.environ.get("PLAUD_TOKEN")
    plaud_base_url = os.environ.get("PLAUD_BASE_URL") or DEFAULT_PLAUD_BASE_URL
    slack_webhook_url = os.environ.get("SLACK_WEBHOOK_URL")

    if not spreadsheet_id:
        print("SPREADSHEET_ID が未設定です")
        sys.exit(1)

    date_str = get_yesterday_jst_str()
    print(f"対象日(JST): {date_str}")

    sheets_token = get_access_token([SHEETS_SCOPE], subject=SPREADSHEET_OWNER_EMAIL)
    existing_ids = fetch_existing_event_ids(sheets_token, spreadsheet_id)
    print(f"記録済みイベント数: {len(existing_ids)}件")

    plaud_items = []
    if plaud_token:
        try:
            recordings = fetch_recordings_by_date(plaud_token, plaud_base_url, date_str)
            plaud_items = fetch_all_summary_items(plaud_token, plaud_base_url, recordings)
            print(f"PLAUD録音件数: {len(plaud_items)}件")
        except PlaudTokenExpiredError as e:
            alert_slack(slack_webhook_url, TOKEN_EXPIRED_MESSAGE)
            print(str(e))
    else:
        print("PLAUD_TOKEN 未設定 - 議事録の突き合わせをスキップします")

    new_rows = []
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
                    transcript = summarize_for_log(clean_summary(matched["summary"]), anthropic_key)

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

    if not new_rows:
        print("記録対象の面談なし")
        return

    if args.dry_run:
        for row in new_rows:
            print("---")
            print(f"取引先: {row[2]} / 面談者: {row[3]}（{row[4]}） / 時刻: {row[5]} / 担当: {row[1]}")
            print(f"議事録: {row[6][:80]}{'...' if len(row[6]) > 80 else ''}")
        print(f"[dry-run] 合計{len(new_rows)}件を書き込み予定（実際には書き込みません）")
        return

    if append_rows(sheets_token, spreadsheet_id, new_rows):
        print(f"{len(new_rows)}件をスプレッドシートに記録しました")


if __name__ == "__main__":
    main()
