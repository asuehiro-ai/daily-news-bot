#!/usr/bin/env python3
"""PLAUD NOTE の前日録音をGeminiで要約しSlackへ投稿する。

必須環境変数: PLAUD_TOKEN, SLACK_WEBHOOK_URL
任意環境変数: GEMINI_API_KEY（未設定時は要約なしでそのまま投稿）,
             PLAUD_BASE_URL（日本アカウントで既定URLが使えない場合に上書き）

録音取得・要約整形の共通ロジックは meeting-log-sync/plaud_client.py に切り出してある
（meeting-log-sync/sync_meeting_log.py と共用するため）。
"""

import argparse
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "meeting-log-sync"))

from plaud_client import (  # noqa: E402
    DEFAULT_PLAUD_BASE_URL,
    PlaudTokenExpiredError,
    TOKEN_EXPIRED_MESSAGE,
    alert_slack,
    extract_key_points,
    fetch_all_summary_items,
    fetch_recordings_by_date,
    get_yesterday_jst_str,
    post_json,
    summarize_items_with_gemini,
    to_jst_time_str,
)


def post_to_slack(webhook_url, yesterday_str, items):
    date_label = yesterday_str.replace("-", "年", 1).replace("-", "月", 1) + "日"
    blocks = [
        {"type": "header", "text": {"type": "plain_text", "text": f"📋 {date_label}の録音ダイジェスト（{len(items)}件）", "emoji": True}}
    ]

    for item in items:
        time_str = to_jst_time_str(item["createTime"]) if item.get("createTime") else ""
        title_line = f"*{item['title']}*" + (f"  _({time_str})_" if time_str else "")
        preview = extract_key_points(item.get("summary") or "")
        blocks.append({"type": "section", "text": {"type": "mrkdwn", "text": f"{title_line}\n{preview}"}})

    post_json(webhook_url, {"blocks": blocks})


def main():
    if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
        sys.stdout.reconfigure(encoding="utf-8")  # Windowsのcp932コンソールで文字化けするのを防ぐ

    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true", help="Slack投稿せずログ出力のみ行う")
    args = parser.parse_args()

    token = os.environ.get("PLAUD_TOKEN")
    webhook_url = os.environ.get("SLACK_WEBHOOK_URL")
    gemini_key = os.environ.get("GEMINI_API_KEY")
    base_url = os.environ.get("PLAUD_BASE_URL") or DEFAULT_PLAUD_BASE_URL

    if not token:
        alert_slack(webhook_url, "⚠️ PLAUD_TOKEN が未設定です")
        sys.exit(1)
    if not webhook_url:
        print("SLACK_WEBHOOK_URL が未設定です")
        sys.exit(1)

    yesterday_str = get_yesterday_jst_str()
    print(f"対象日(JST): {yesterday_str}")

    try:
        recordings = fetch_recordings_by_date(token, base_url, yesterday_str)
        if not recordings:
            print("前日の録音なし - 通知スキップ")
            return
        print(f"録音件数: {len(recordings)}件")

        items = fetch_all_summary_items(token, base_url, recordings)
    except PlaudTokenExpiredError as e:
        alert_slack(webhook_url, TOKEN_EXPIRED_MESSAGE)
        print(str(e))
        sys.exit(1)

    if not items:
        print("取得できた要約なし")
        return

    final_items = summarize_items_with_gemini(items, gemini_key) if gemini_key else items

    if args.dry_run:
        for item in final_items:
            print("---")
            print(f"タイトル: {item['title']}")
            print(f"要約冒頭: {item['summary'][:100]}...")
        return

    post_to_slack(webhook_url, yesterday_str, final_items)
    print("Slack投稿完了")


if __name__ == "__main__":
    main()
