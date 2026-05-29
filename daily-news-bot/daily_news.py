#!/usr/bin/env python3
"""Daily News Bot - RSSからニュースを収集しAIで要約してSlackに投稿"""

import feedparser
import anthropic
import urllib.request
import json
import datetime
import sys
import os
import re

# ── RSS ソース定義 ────────────────────────────────────────────
# 同じRSSを複数カテゴリで使うと used_titles で弾かれるため、
# 各フィードは原則1カテゴリのみに割り当てる。
GNEWS = "https://news.google.com/rss/search?hl=ja&gl=JP&ceid=JP:ja&q="
YAHOO = "https://news.yahoo.co.jp/rss/topics/"

RSS_FEEDS = {
    "経済": [
        GNEWS + "%E6%97%A5%E6%9C%AC%E7%B5%8C%E6%B8%88+%E6%99%AF%E6%B0%97",   # 日本経済 景気
        "https://www3.nhk.or.jp/rss/news/cat4.xml",                           # NHK経済
        YAHOO + "business.xml",                                                # Yahoo!ビジネス
    ],
    "ビジネス": [
        GNEWS + "%E4%BC%81%E6%A5%AD%E3%83%8B%E3%83%A5%E3%83%BC%E3%82%B9+%E6%97%A5%E6%9C%AC",  # 企業ニュース 日本
        YAHOO + "top-picks.xml",                                               # Yahoo!トップ（補完）
    ],
    "マーケット": [
        GNEWS + "%E6%A0%AA%E5%BC%8F+%E7%82%BA%E6%9B%BF+%E5%B8%82%E5%A0%B4",  # 株式 為替 市場
    ],
    "国内政治": [
        GNEWS + "%E5%9B%BD%E5%86%85%E6%94%BF%E6%B2%BB+%E5%9B%BD%E4%BC%9A",   # 国内政治 国会
        "https://www3.nhk.or.jp/rss/news/cat6.xml",                           # NHK政治
        YAHOO + "domestic.xml",                                                # Yahoo!国内
    ],
    "M&A": [
        GNEWS + "%E8%B3%87%E6%9C%AC%E6%8F%90%E6%90%BA+%E8%B3%87%E6%9C%AC%E6%A5%AD%E5%8B%99%E6%8F%90%E6%90%BA",  # 資本提携 資本業務提携
        GNEWS + "M%26A+%E4%BC%81%E6%A5%AD%E8%B2%B7%E5%8F%8E+%E5%90%88%E4%BD%B5",                                  # M&A 企業買収 合併
        GNEWS + "%E8%B2%B7%E5%8F%8E+%E5%A3%B2%E5%8D%B4+%E4%BA%8B%E6%A5%AD%E8%AD%B2%E6%B8%A1",                    # 買収 売却 事業譲渡
    ],
}

CUTOFF_HOURS = {"M&A": 72}
DEFAULT_CUTOFF_HOURS = 48


def clean(text):
    return re.sub(r"<[^>]+>", "", text or "").strip()


# ── ステップ②③ RSSフィード取得 ─────────────────────────────
def fetch_news(feeds, max_items=8, used_titles=None,
               cutoff_hours=DEFAULT_CUTOFF_HOURS, category=""):
    """
    RSSを取得して記事を返す。
    戻り値: (items, diag)
      diag = {category, total, feeds:[], errors:[]}
    """
    if used_titles is None:
        used_titles = set()

    cutoff = datetime.datetime.utcnow() - datetime.timedelta(hours=cutoff_hours)
    items       = []
    feed_lines  = []
    error_lines = []

    for url in feeds:
        label = url.split("?")[0][-60:]
        try:
            # ── ステップ② RSS取得 ──
            feed = feedparser.parse(url, request_headers={"User-Agent": "Mozilla/5.0"})

            # HTTP エラー検出
            status = getattr(feed, "status", None)
            if status and status >= 400:
                error_lines.append(f"  ❌ {label}\n     → HTTP {status} エラー（アクセス拒否 or URL不正）")
                print(f"  [{category}] HTTP {status}: {label}")
                continue

            # RSS 解析エラー検出（feedparser の bozo フラグ）
            if getattr(feed, "bozo", False):
                bozo_msg = str(getattr(feed, "bozo_exception", "不明"))
                error_lines.append(f"  ⚠️ {label}\n     → RSS解析エラー: {bozo_msg}")
                print(f"  [{category}] RSS解析エラー ({bozo_msg}): {label}")
                # 解析エラーでも記事が取れることがあるので続行

            total  = len(feed.entries)
            n_pass = 0  # ステップ③ 日付フィルタ通過数
            n_dup  = 0  # 重複でスキップ数
            n_new  = 0  # 新規追加数

            # ── ステップ③ 48h フィルタ＆重複除外 ──
            for entry in feed.entries:
                title   = clean(entry.get("title", ""))
                summary = clean(entry.get("summary", entry.get("description", "")))[:300]

                # 日付フィルタ
                published = entry.get("published_parsed")
                if published:
                    pub_dt = datetime.datetime(*published[:6])
                    if pub_dt < cutoff:
                        continue
                n_pass += 1

                # 重複チェック
                if title and title not in used_titles:
                    items.append(f"・{title}：{summary}")
                    used_titles.add(title)
                    n_new += 1
                else:
                    n_dup += 1

                if len(items) >= max_items:
                    break

            # ステップ③ 原因別エラー記録
            if total == 0:
                error_lines.append(f"  ❌ {label}\n     → フィードが空（記事0件）")
            elif n_pass == 0:
                error_lines.append(
                    f"  ⚠️ {label}\n"
                    f"     → 全{total}件が{cutoff_hours}時間超で日付フィルタに除外"
                )
            elif n_new == 0 and n_dup > 0:
                error_lines.append(
                    f"  ⚠️ {label}\n"
                    f"     → {n_pass}件取得できたが全て他カテゴリと重複してスキップ"
                )

            status_icon = "✅" if n_new > 0 else ("⚠️" if total > 0 else "❌")
            line = (f"  {status_icon} {label}\n"
                    f"     取得:{total}件 → 日付OK:{n_pass}件 → 重複:{n_dup}件 → 新規追加:{n_new}件")
            feed_lines.append(line)
            print(f"  [{category}] 取得:{total} 日付OK:{n_pass} 重複:{n_dup} 新規:{n_new} | {label}")

        except Exception as e:
            error_lines.append(f"  ❌ {label}\n     → 例外エラー: {e}")
            print(f"  [{category}] 例外: {e} | {label}")

        if len(items) >= max_items:
            break

    diag = {
        "category": category,
        "total":    len(items),
        "feeds":    feed_lines,
        "errors":   error_lines,
    }
    print(f"  [{category}] 合計 {len(items)}件収集")
    return items, diag


# ── Claude AI 要約 ─────────────────────────────────────────
def summarize(client, category, news_items):
    if not news_items:
        return "本日のニュースを取得できませんでした。"

    news_text = "\n".join(news_items)
    extra = ""
    if category == "M&A":
        extra = ("M&A・企業買収・合併・事業譲渡・資本提携・出資に関するニュースを優先して選んでください。"
                 "該当するニュースがない場合は「本日はM&A関連の目立った報道はありませんでした。」と1行で返してください。")

    response = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=800,
        messages=[{
            "role": "user",
            "content": (
                f"以下の「{category}」ニュースから重要な3件を選び、日本語で簡潔にまとめてください。{extra}\n\n"
                "形式（この形式を厳守）：\n"
                "1. [タイトル]\n  概要：[2文で要点を説明]\n  ポイント：[重要な論点・背景・注目すべき点を1〜2文で]\n\n"
                "2. [タイトル]\n  概要：[2文で要点を説明]\n  ポイント：[重要な論点・背景・注目すべき点を1〜2文で]\n\n"
                "3. [タイトル]\n  概要：[2文で要点を説明]\n  ポイント：[重要な論点・背景・注目すべき点を1〜2文で]\n\n"
                f"ニュース一覧：\n{news_text}\n\n"
                "注意：マークダウン記号（**や##）は使わず、プレーンテキストで出力してください。"
            )
        }]
    )
    return response.content[0].text.strip()


# ── 診断レポート生成 ──────────────────────────────────────
def build_diag_report(all_diag, date_str):
    """
    全カテゴリの診断結果をまとめる。
    エラーまたは0件取得があった場合のみ Slack に追記する。
    """
    has_problem = any(d["total"] == 0 or d["errors"] for d in all_diag)
    if not has_problem:
        return None

    lines = [
        "",
        "━━━━━━━━━━━━━━━━",
        f"【取得エラー診断】{date_str}",
        "━━━━━━━━━━━━━━━━",
        "▼ カテゴリ別 取得件数",
    ]
    for d in all_diag:
        icon = "✅" if d["total"] > 0 and not d["errors"] else "❌"
        lines.append(f"  {icon} {d['category']}：{d['total']}件")

    lines.append("")
    lines.append("▼ フィード詳細")
    for d in all_diag:
        lines.append(f"【{d['category']}】")
        for fl in d["feeds"]:
            lines.append(fl)
        if not d["feeds"]:
            lines.append("  （フィードなし）")

    error_all = [e for d in all_diag for e in d["errors"]]
    if error_all:
        lines.append("")
        lines.append("▼ エラー一覧")
        lines.extend(error_all)

    return "\n".join(lines)


def post_to_slack(webhook_url, text):
    data = json.dumps({"text": text}).encode("utf-8")
    req  = urllib.request.Request(
        webhook_url, data=data,
        headers={"Content-Type": "application/json"}
    )
    with urllib.request.urlopen(req) as r:
        return r.read().decode()


# ── メイン処理 ────────────────────────────────────────────
def main():
    # ── ステップ① 起動確認 ──
    jst      = datetime.datetime.utcnow() + datetime.timedelta(hours=9)
    date_str = jst.strftime("%Y年%m月%d日")
    print(f"=== {date_str} ニュース収集開始 ===")
    print(f"Python: {sys.version}")
    print(f"UTC: {datetime.datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')}")

    api_key     = os.environ["ANTHROPIC_API_KEY"]
    webhook_url = os.environ["SLACK_WEBHOOK_URL"]
    client      = anthropic.Anthropic(api_key=api_key)

    parts       = [f"おはようございます。本日（{date_str}）の主要ニュースです。"]
    all_diag    = []
    used_titles = set()

    for category, feeds in RSS_FEEDS.items():
        print(f"\n--- {category} ---")
        hours = CUTOFF_HOURS.get(category, DEFAULT_CUTOFF_HOURS)
        news_items, diag = fetch_news(
            feeds, used_titles=used_titles,
            cutoff_hours=hours, category=category
        )
        all_diag.append(diag)
        summary = summarize(client, category, news_items)

        parts += ["", "━━━━━━━━━━━━━━━━", f"【{category}】", "━━━━━━━━━━━━━━━━", summary]

    # 問題があれば診断レポートを末尾に追加
    report = build_diag_report(all_diag, date_str)
    if report:
        parts.append(report)

    message = "\n".join(parts)
    result  = post_to_slack(webhook_url, message)
    print(f"\n=== Slack投稿完了: {result} ===")


if __name__ == "__main__":
    main()
