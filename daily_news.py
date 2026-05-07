#!/usr/bin/env python3
"""Daily News Bot - RSSからニュースを収集してSlackに投稿"""

import feedparser
import urllib.request
import json
import datetime
import os
import re

RSS_FEEDS = {
    "経済": [
        "https://www3.nhk.or.jp/rss/news/cat4.xml",
        "https://feeds.reuters.com/reuters/JPbusinessNews",
    ],
    "政治": [
        "https://www3.nhk.or.jp/rss/news/cat6.xml",
        "https://feeds.reuters.com/reuters/JPpoliticsNews",
    ],
    "国際": [
        "https://www3.nhk.or.jp/rss/news/cat7.xml",
        "https://feeds.reuters.com/reuters/JPworldNews",
    ],
    "M&A": [
        "https://feeds.reuters.com/reuters/JPmergers",
        "https://www3.nhk.or.jp/rss/news/cat4.xml",
    ],
}


def clean(text):
    return re.sub(r"<[^>]+>", "", text or "").strip()


def fetch_news(feeds, max_items=3):
    items = []
    for url in feeds:
        try:
            feed = feedparser.parse(url, request_headers={"User-Agent": "Mozilla/5.0"})
            for entry in feed.entries:
                title = clean(entry.get("title", ""))
                summary = clean(entry.get("summary", entry.get("description", "")))[:150]
                if title:
                    items.append((title, summary))
                if len(items) >= max_items:
                    return items
        except Exception as e:
            print(f"Warning: {url} の取得失敗: {e}")
    return items[:max_items]


def post_to_slack(webhook_url, text):
    data = json.dumps({"text": text}).encode("utf-8")
    req = urllib.request.Request(
        webhook_url, data=data,
        headers={"Content-Type": "application/json"}
    )
    with urllib.request.urlopen(req) as r:
        return r.read().decode()


def main():
    webhook_url = os.environ["SLACK_WEBHOOK_URL"]

    jst = datetime.datetime.utcnow() + datetime.timedelta(hours=9)
    date_str = jst.strftime("%Y年%m月%d日")

    print(f"{date_str} のニュース収集開始...")

    parts = [f"おはようございます。本日（{date_str}）の主要ニュースです。"]

    for category, feeds in RSS_FEEDS.items():
        print(f"{category} 取得中...")
        items = fetch_news(feeds)

        parts.append("")
        parts.append("━━━━━━━━━━━━━━━━")
        parts.append(f"【{category}】")
        parts.append("━━━━━━━━━━━━━━━━")

        if items:
            for i, (title, summary) in enumerate(items, 1):
                parts.append(f"{i}. {title}")
                if summary:
                    parts.append(f"  {summary}")
                parts.append("")
        else:
            parts.append("ニュースを取得できませんでした。")

    message = "\n".join(parts)
    result = post_to_slack(webhook_url, message)
    print(f"Slack投稿完了: {result}")


if __name__ == "__main__":
    main()
