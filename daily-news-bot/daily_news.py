#!/usr/bin/env python3
"""Daily News Bot - RSSからニュースを収集しAIで要約してSlackに投稿"""

import feedparser
import anthropic
import urllib.request
import json
import datetime
import os
import re

# Google News RSS（キーワード検索型・安定稼働）
BASE = "https://news.google.com/rss/search?hl=ja&gl=JP&ceid=JP:ja&q="
RSS_FEEDS = {
    "経済": [
        BASE + "%E6%97%A5%E6%9C%AC%E7%B5%8C%E6%B8%88+%E6%99%AF%E6%B0%97",        # 日本経済 景気
        "https://www3.nhk.or.jp/rss/news/cat4.xml",
    ],
    "ビジネス": [
        BASE + "%E4%BC%81%E6%A5%AD%E3%83%8B%E3%83%A5%E3%83%BC%E3%82%B9+%E6%97%A5%E6%9C%AC",  # 企業ニュース 日本
        "https://www3.nhk.or.jp/rss/news/cat4.xml",
    ],
    "マーケット": [
        BASE + "%E6%A0%AA%E5%BC%8F+%E7%82%BA%E6%9B%BF+%E5%B8%82%E5%A0%B4",       # 株式 為替 市場
    ],
    "国内政治": [
        BASE + "%E5%9B%BD%E5%86%85%E6%94%BF%E6%B2%BB+%E5%9B%BD%E4%BC%9A",        # 国内政治 国会
        "https://www3.nhk.or.jp/rss/news/cat6.xml",
    ],
    "M&A": [
        BASE + "%E8%B3%87%E6%9C%AC%E6%8F%90%E6%90%BA+%E8%B3%87%E6%9C%AC%E6%A5%AD%E5%8B%99%E6%8F%90%E6%90%BA", # 資本提携 資本業務提携
        BASE + "M%26A+%E4%BC%81%E6%A5%AD%E8%B2%B7%E5%8F%8E+%E5%90%88%E4%BD%B5",  # M&A 企業買収 合併
        BASE + "%E8%B2%B7%E5%8F%8E+%E5%A3%B2%E5%8D%B4+%E4%BA%8B%E6%A5%AD%E8%AD%B2%E6%B8%A1", # 買収 売却 事業譲渡
    ],
}

EXCLUDE_SPORTS = set()
CUTOFF_HOURS = {"M&A": 72}
DEFAULT_CUTOFF_HOURS = 48


def clean(text):
    return re.sub(r"<[^>]+>", "", text or "").strip()


def fetch_news(feeds, max_items=8, used_titles=None, cutoff_hours=DEFAULT_CUTOFF_HOURS, category=""):
    """RSSフィードからニュースを取得（重複除外）"""
    if used_titles is None:
        used_titles = set()
    cutoff = datetime.datetime.utcnow() - datetime.timedelta(hours=cutoff_hours)
    items = []
    for url in feeds:
        short = url.split("?")[0][-50:]  # ログ用に短縮
        try:
            feed = feedparser.parse(url, request_headers={"User-Agent": "Mozilla/5.0"})
            total   = len(feed.entries)
            n_pass  = 0  # 日付OK
            n_new   = 0  # 重複なし（新規）
            for entry in feed.entries:
                title = clean(entry.get("title", ""))
                summary = clean(entry.get("summary", entry.get("description", "")))[:300]

                published = entry.get("published_parsed")
                if published:
                    pub_dt = datetime.datetime(*published[:6])
                    if pub_dt < cutoff:
                        continue
                n_pass += 1

                if title and title not in used_titles:
                    items.append(f"・{title}：{summary}")
                    used_titles.add(title)
                    n_new += 1
                if len(items) >= max_items:
                    break

            print(f"  [{category}] {short} → 取得:{total}件 / 日付OK:{n_pass}件 / 新規追加:{n_new}件")
        except Exception as e:
            print(f"  [{category}] NG {short} → エラー: {e}")
        if len(items) >= max_items:
            break
    print(f"  [{category}] 合計 {len(items)}件 収集")
    return items


def summarize(client, category, news_items):
    if not news_items:
        return "本日のニュースを取得できませんでした。"

    news_text = "\n".join(news_items)
    if category in EXCLUDE_SPORTS:
        exclude_note = "スポーツ関連のニュースは除外してください。"
    elif category == "M&A":
        exclude_note = "M&A・企業買収・合併・事業譲渡・提携・出資に関するニュースを優先して選んでください。該当するニュースがない場合は「本日はM&A関連の目立った報道はありませんでした。」と1行で返してください。"
    else:
        exclude_note = ""

    response = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=800,
        messages=[{
            "role": "user",
            "content": f"""以下の「{category}」ニュースから重要な3件を選び、日本語で簡潔にまとめてください。{exclude_note}

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

    # カテゴリーをまたいで重複タイトルを管理
    used_titles = set()

    for category, feeds in RSS_FEEDS.items():
        print(f"{category} 取得・要約中...")
        hours = CUTOFF_HOURS.get(category, DEFAULT_CUTOFF_HOURS)
        news_items = fetch_news(feeds, used_titles=used_titles, cutoff_hours=hours, category=category)
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
