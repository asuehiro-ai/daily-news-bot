#!/usr/bin/env python3
"""Daily News Bot - 毎朝ニュースを収集してSlackに投稿するスクリプト"""

import feedparser
import anthropic
import urllib.request
import json
import datetime
import os
import re

# RSSフィードのURL
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


def fetch_news_items(feeds: list, max_per_feed: int = 8) -> list:
    """RSSフィードからニュースを取得する"""
    items = []
    for url in feeds:
        try:
            feed = feedparser.parse(url, request_headers={"User-Agent": "DailyNewsBot/1.0"})
            for entry in feed.entries[:max_per_feed]:
                title = entry.get("title", "").strip()
                summary = entry.get("summary", entry.get("description", "")).strip()
                summary = re.sub(r"<[^>]+>", "", summary)[:300]
                if title:
                    items.append(f"タイトル: {title}\n概要: {summary}")
        except Exception as e:
            print(f"Warning: {url} の取得に失敗: {e}")
    return items


def get_category_summary(client, category: str, news_items: list) -> str:
    """Claudeを使ってカテゴリーのニュースを要約する"""
    if not news_items:
        return "本日のニュースを取得できませんでした。"

    news_text = "\n\n".join(news_items[:15])

    prompt = f"""以下は「{category}」カテゴリーの最新ニュース記事です。
最も重要な3件を選んで、以下の形式で日本語でまとめてください。

出力形式：
1. [タイトル]
  概要：[2〜3文で内容を説明]
  ポイント：[重要な論点・背景・注目すべき点]

2. [タイトル]
  概要：[2〜3文で内容を説明]
  ポイント：[重要な論点・背景・注目すべき点]

3. [タイトル]
  概要：[2〜3文で内容を説明]
  ポイント：[重要な論点・背景・注目すべき点]

ニュース記事：
{news_text}

注意：マークダウン記号（**、##など）は使わず、プレーンテキストで出力してください。"""

    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1000,
        messages=[{"role": "user", "content": prompt}]
    )
    return response.content[0].text.strip()


def post_to_slack(webhook_url: str, text: str) -> str:
    """SlackのWebhook URLにメッセージを投稿する"""
    data = json.dumps({"text": text}).encode("utf-8")
    req = urllib.request.Request(
        webhook_url,
        data=data,
        headers={"Content-Type": "application/json"}
    )
    with urllib.request.urlopen(req) as resp:
        return resp.read().decode()


def main():
    api_key = os.environ["ANTHROPIC_API_KEY"]
    webhook_url = os.environ["SLACK_WEBHOOK_URL"]

    client = anthropic.Anthropic(api_key=api_key)

    jst = datetime.datetime.utcnow() + datetime.timedelta(hours=9)
    date_str = jst.strftime("%Y年%m月%d日")

    print(f"{date_str} のニュース収集を開始します...")

    message_parts = [f"おはようございます。本日（{date_str}）の主要ニュースです。"]

    for category in ["経済", "政治", "国際", "M&A"]:
        print(f"{category} のニュースを取得中...")
        news_items = fetch_news_items(RSS_FEEDS[category])
        summary = get_category_summary(client, category, news_items)

        message_parts.append("")
        message_parts.append("━━━━━━━━━━━━━━━━")
        message_parts.append(f"【{category}】")
        message_parts.append("━━━━━━━━━━━━━━━━")
        message_parts.append(summary)

    full_message = "\n".join(message_parts)
    result = post_to_slack(webhook_url, full_message)
    print(f"Slackへの投稿完了: {result}")


if __name__ == "__main__":
    main()
