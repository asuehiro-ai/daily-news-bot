#!/usr/bin/env python3
"""Daily News Bot - RSSからニュースを収集しAIで要約してSlackに投稿"""

import feedparser
import anthropic
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


def fetch_news(feeds, max_items=6):
    items = []
    for url in feeds:
        try:
            feed = feedparser.parse(url, request_headers={"User-Agent": "Mozilla/5.0"})
            for entry in feed.entries:
                title = clean(entry.get("title", ""))
                summary = clean(entry.get("summary", entry.get("description", "")))[:300]
                if title:
                    items.append(f"・{title}：{summary}")
                if len(items) >= max_items:
                    return items
        except Exception as e:
            print(f"Warning: {url} の取得失敗: {e}")
    return items[:max_items]


def summarize(client, category, news_items):
    if not news_items:
        return "本日のニュースを取得できませんでした。"

    news_text = "\n".join(news_items)

    response = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=800,
        messages=[{
            "role": "user",
            "content": f"""以下の「{category}」ニュースから重要な3件を選び、日本語で簡潔にまとめてください。

形式（この形式を厳守）：
1. [タイトル]
  概要：[2文で要点を説明]
  ポイント：[重要な論点・背景・注目すべき点を1〜2文で]

2. [タイトル]
  概要：[2文で要点を説明]
  ポイント：[重要な論点・背景・注目すべき点を1〜2文で]

3. [タイトル]
  概要：[2文で要点を説明]
  ポイント：[重要な論点・背景・注目すべき点を1〜2文で]

ニュース一覧：
{news_text}

注意：マークダウン記号（**や##）は使わず、プレーンテキストで出力してください。"""
        }]
    )
    return response.content[0].text.strip()


def post_to_slack(webhook_url, text):
    data = json.dumps({"text": text}).encode("utf-8")
    req = urllib.request.Request(
        webhook_url, data=data,
        headers={"Content-Type": "application/json"}
    )
    with urllib.request.urlopen(req) as r:
        return r.read().decode()


def main():
    api_key = os.environ["ANTHROPIC_API_KEY"]
    webhook_url = os.environ["SLACK_WEBHOOK_URL"]

    client = anthropic.Anthropic(api_key=api_key)

    jst = datetime.datetime.utcnow() + datetime.timedelta(hours=9)
    date_str = jst.strftime("%Y年%m月%d日")

    print(f"{date_str} のニュース収集開始...")

    parts = [f"おはようございます。本日（{date_str}）の主要ニュースです。"]

    for category, feeds in RSS_FEEDS.items():
        print(f"{category} 取得・要約中...")
        news_items = fetch_news(feeds)
        summary = summarize(client, category, news_items)

        parts.append("")
        parts.append("━━━━━━━━━━━━━━━━")
        parts.append(f"【{category}】")
        parts.append("━━━━━━━━━━━━━━━━")
        parts.append(summary)

    message = "\n".join(parts)
    result = post_to_slack(webhook_url, message)
    print(f"Slack投稿完了: {result}")


if __name__ == "__main__":
    main()
