#!/usr/bin/env python3
"""
RSS フィード疎通テスト
GitHub Actions の手動実行（workflow_dispatch）で動かし、
どのフィードが使えるか・なぜ失敗するかをログで確認する。
ANTHROPIC_API_KEY / SLACK_WEBHOOK_URL 不要。
"""

import feedparser
import datetime
import sys

GNEWS = "https://news.google.com/rss/search?hl=ja&gl=JP&ceid=JP:ja&q="
YAHOO = "https://news.yahoo.co.jp/rss/topics/"

TARGETS = {
    "NHK 経済 (cat4)":     "https://www3.nhk.or.jp/rss/news/cat4.xml",
    "NHK 政治 (cat6)":     "https://www3.nhk.or.jp/rss/news/cat6.xml",
    "NHK 国際 (cat5)":     "https://www3.nhk.or.jp/rss/news/cat5.xml",
    "Yahoo ビジネス":       YAHOO + "business.xml",
    "Yahoo トップ":         YAHOO + "top-picks.xml",
    "Yahoo 国内":           YAHOO + "domestic.xml",
    "Google 日本経済":      GNEWS + "%E6%97%A5%E6%9C%AC%E7%B5%8C%E6%B8%88+%E6%99%AF%E6%B0%97",
    "Google 企業ニュース":  GNEWS + "%E4%BC%81%E6%A5%AD%E3%83%8B%E3%83%A5%E3%83%BC%E3%82%B9+%E6%97%A5%E6%9C%AC",
    "Google 株式市場":      GNEWS + "%E6%A0%AA%E5%BC%8F+%E7%82%BA%E6%9B%BF+%E5%B8%82%E5%A0%B4",
    "Google 国内政治":      GNEWS + "%E5%9B%BD%E5%86%85%E6%94%BF%E6%B2%BB+%E5%9B%BD%E4%BC%9A",
    "Google 資本提携":      GNEWS + "%E8%B3%87%E6%9C%AC%E6%8F%90%E6%90%BA+%E8%B3%87%E6%9C%AC%E6%A5%AD%E5%8B%99%E6%8F%90%E6%90%BA",
    "Google M&A":           GNEWS + "M%26A+%E4%BC%81%E6%A5%AD%E8%B2%B7%E5%8F%8E+%E5%90%88%E4%BD%B5",
    "Google 買収売却":      GNEWS + "%E8%B2%B7%E5%8F%8E+%E5%A3%B2%E5%8D%B4+%E4%BA%8B%E6%A5%AD%E8%AD%B2%E6%B8%A1",
}

CUTOFF = datetime.datetime.utcnow() - datetime.timedelta(hours=48)


def test(name, url):
    print(f"\n{'='*60}")
    print(f"【{name}】")
    print(f"URL: {url}")
    try:
        feed = feedparser.parse(url, request_headers={"User-Agent": "Mozilla/5.0"})

        status = getattr(feed, "status", "N/A")
        print(f"HTTPステータス: {status}")

        if isinstance(status, int) and status >= 400:
            print(f"❌ HTTPエラー → アクセス拒否またはURL不正")
            return

        bozo = getattr(feed, "bozo", False)
        if bozo:
            print(f"⚠️  RSS解析エラー: {getattr(feed, 'bozo_exception', '不明')}")

        total = len(feed.entries)
        print(f"総記事数: {total}")

        if total == 0:
            print("❌ 記事0件 → フィードが空またはブロックされている可能性")
            return

        # 最新記事の日付確認
        newest = None
        n_recent = 0
        for e in feed.entries:
            pub = e.get("published_parsed")
            if pub:
                dt = datetime.datetime(*pub[:6])
                if newest is None or dt > newest:
                    newest = dt
                if dt >= CUTOFF:
                    n_recent += 1

        if newest:
            age_h = (datetime.datetime.utcnow() - newest).total_seconds() / 3600
            print(f"最新記事日時: {newest.strftime('%Y-%m-%d %H:%M')} UTC（約{age_h:.1f}時間前）")
        else:
            print("最新記事日時: 日付情報なし（全件フィルタ通過扱い）")

        print(f"48時間以内の記事: {n_recent}/{total} 件")

        if n_recent == 0 and newest:
            print("⚠️  全記事が48時間超 → 日付フィルタで全除外される")
        elif n_recent > 0:
            print(f"✅ 正常 → {n_recent}件が取得可能")

        # サンプル記事表示（最新3件）
        print("--- サンプル記事（最新3件）---")
        for e in feed.entries[:3]:
            title = e.get("title", "(タイトルなし)")
            pub   = e.get("published", "日付不明")
            print(f"  ・{title}  [{pub}]")

    except Exception as ex:
        print(f"❌ 例外エラー: {ex}")


if __name__ == "__main__":
    print(f"RSS疎通テスト開始: {datetime.datetime.utcnow()} UTC")
    print(f"Python: {sys.version}")
    print(f"48時間カットオフ: {CUTOFF} UTC")

    for name, url in TARGETS.items():
        test(name, url)

    print(f"\n{'='*60}")
    print("テスト完了")
