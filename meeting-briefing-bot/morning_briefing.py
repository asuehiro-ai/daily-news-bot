#!/usr/bin/env python3
"""朝の面談ブリーフィングをSlackへ自動投稿する（meeting_briefing_bot.gsのPython移植）。

対象日のGoogleカレンダー予定から「【Web】●●株式会社 山田様」形式のタイトルを抽出し、
会社ごとにGemini（Google検索グラウンディング）でリサーチしてSlackに投稿する。

必須環境変数: GOOGLE_SERVICE_ACCOUNT_JSON, GEMINI_API_KEY, SLACK_WEBHOOK_URL
"""

import argparse
import datetime
import json
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
from zoneinfo import ZoneInfo

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "meeting-log-sync"))

from google_auth import CALENDAR_READONLY_SCOPE, get_access_token  # noqa: E402
from plaud_client import post_json  # noqa: E402

CALENDAR_OWNER_EMAIL = "a.suehiro@le-gr.co.jp"
GEMINI_MODEL = "gemini-2.5-flash"  # gemini-2.0-flashは2026/6/1に提供終了済み
TITLE_PATTERN = re.compile(r"^【([^】]*)】\s*([^\s（(]+)\s*(.*)$")
JST = ZoneInfo("Asia/Tokyo")


# ── カレンダー読み込み・タイトル解析 ─────────────────────────


def fetch_todays_events(token, calendar_id, date_str):
    time_min = urllib.parse.quote(f"{date_str}T00:00:00+09:00")
    time_max = urllib.parse.quote(f"{date_str}T23:59:59+09:00")
    encoded_id = urllib.parse.quote(calendar_id, safe="")
    url = (
        f"https://www.googleapis.com/calendar/v3/calendars/{encoded_id}/events"
        f"?timeMin={time_min}&timeMax={time_max}&singleEvents=true&orderBy=startTime"
    )
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"})
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            body = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        print(f"カレンダー取得エラー: HTTP {e.code} {e.read().decode('utf-8')}")
        return []
    return body.get("items", [])


def parse_event_title(title):
    m = TITLE_PATTERN.match(title or "")
    if not m:
        return None
    return {"tag": m.group(1).strip(), "company": m.group(2).strip(), "person": m.group(3).strip()}


def group_meetings_by_company(events):
    """会社名をキーに { meetings: [{time, tag, person}] } を積み上げて返す。"""
    companies = {}
    for event in events:
        if "dateTime" not in event.get("start", {}):
            continue  # 終日予定はスキップ
        parsed = parse_event_title(event.get("summary", ""))
        if not parsed:
            continue

        start_dt = datetime.datetime.fromisoformat(event["start"]["dateTime"]).astimezone(JST)
        companies.setdefault(parsed["company"], {"meetings": []})
        companies[parsed["company"]]["meetings"].append(
            {"time": start_dt.strftime("%H:%M"), "tag": parsed["tag"], "person": parsed["person"]}
        )
    return companies


# ── Gemini リサーチ（Google検索グラウンディング） ────────────


def build_research_prompt(company_name):
    return (
        f"「{company_name}」という企業についてWeb検索を行い、以下の2点を日本語でまとめてください。\n\n"
        "① 会社概要（事業内容・規模・沿革・最近のニュースやトピック）\n"
        "② 業界動向（この会社が属する業界の直近の動向・トレンド）\n\n"
        "出力ルール:\n"
        "- ①②それぞれ見出しをそのまま使い、箇条書き（- で始める）で3〜5項目ずつ\n"
        "- 各項目は1〜2文で簡潔に、固有名詞・数値は省略しない\n"
        "- 前置きや結びの言葉は不要、見出しと箇条書きのみ出力\n"
        "- 検索してもこの名称の企業が特定できない場合、①には社名の文字列（業種を示す語・カタカナ英語表記など）"
        "から推測できる事業内容の仮説を1〜2項目で記載する\n"
        "- ②は企業を特定できなかった場合でも空欄にせず、社名や①の仮説から推測される業界の一般的な直近動向を"
        "必ず記載する。その場合は②の先頭に「（企業を特定できなかったため、社名から推測した業界の動向です）」"
        "と一言添える"
    )


def call_gemini(prompt, api_key, *, with_search):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent?key={api_key}"
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": 0.2,
            "maxOutputTokens": 4096,
            "thinkingConfig": {"thinkingBudget": 0},  # 内部思考トークンで出力が尻切れになるのを防ぐ
        },
    }
    if with_search:
        payload["tools"] = [{"google_search": {}}]

    req = urllib.request.Request(
        url, data=json.dumps(payload).encode("utf-8"), method="POST", headers={"Content-Type": "application/json"}
    )
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            result = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        print(f"Gemini エラー({'検索あり' if with_search else '検索なし'}): HTTP {e.code} {e.read().decode('utf-8')[:300]}")
        return None
    except urllib.error.URLError as e:
        print(f"Gemini 呼び出し例外: {e}")
        return None

    candidate = (result.get("candidates") or [None])[0]
    if not candidate:
        return None
    if candidate.get("finishReason") == "MAX_TOKENS":
        print("Gemini 警告: maxOutputTokensに到達し出力が途中で切れています")

    parts = (candidate.get("content") or {}).get("parts") or []
    text = parts[0].get("text") if parts else None
    if not text:
        return None

    chunks = (candidate.get("groundingMetadata") or {}).get("groundingChunks") or []
    seen = set()
    references = []
    for c in chunks:
        web = c.get("web") or {}
        uri = web.get("uri")
        if uri and uri not in seen:
            seen.add(uri)
            references.append({"title": web.get("title") or uri, "uri": uri})

    return {"text": text.strip(), "references": references, "grounded": with_search}


def research_company(company_name, api_key):
    prompt = build_research_prompt(company_name)
    result = call_gemini(prompt, api_key, with_search=True)
    if not result:
        print(f"{company_name}: グラウンディング検索に失敗、フォールバックします")
        result = call_gemini(prompt, api_key, with_search=False)
        if result:
            result["grounded"] = False
    if not result:
        return {"text": "情報を取得できませんでした。", "references": [], "grounded": False}
    return result


def research_companies(companies, api_key):
    results = {}
    for company, data in companies.items():
        print(f"リサーチ中: {company}")
        research = research_company(company, api_key)
        results[company] = {**data, **research}
    return results


# ── Slack 投稿 ────────────────────────────────────────────────


def truncate_for_slack(text):
    return text[:2500] + "…" if len(text) > 2500 else text


def build_briefing_blocks(date_str, results):
    date_label = datetime.datetime.strptime(date_str, "%Y-%m-%d").strftime("%Y/%m/%d")
    blocks = [
        {
            "type": "header",
            "text": {"type": "plain_text", "text": f"📅 本日の面談ブリーフィング（{date_label}・{len(results)}社）", "emoji": True},
        }
    ]

    first = True
    for company, data in results.items():
        if not first:
            blocks.append({"type": "divider"})
        first = False

        meeting_lines = "\n".join(
            f"・{m['time']}　{m['tag']}　{m['person'] or '(相手名不明)'}" for m in data["meetings"]
        )
        body = f"*{company}*\n{meeting_lines}\n\n{truncate_for_slack(data['text'])}"
        if data.get("grounded") is False:
            body += "\n_※検索グラウンディングに失敗したため一般知識ベースの内容です_"
        blocks.append({"type": "section", "text": {"type": "mrkdwn", "text": body}})

        references = data.get("references") or []
        if references:
            ref_text = "参照: " + " / ".join(f"<{r['uri']}|{r['title']}>" for r in references[:5])
            blocks.append({"type": "context", "elements": [{"type": "mrkdwn", "text": ref_text}]})

    return blocks


# ── メイン処理 ────────────────────────────────────────────────


def main():
    if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
        sys.stdout.reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true", help="Slack投稿をせずログ出力のみ行う")
    parser.add_argument("--date", help="対象日(YYYY-MM-DD、JST)。省略時は本日")
    args = parser.parse_args()

    gemini_key = os.environ.get("GEMINI_API_KEY")
    webhook_url = os.environ.get("SLACK_WEBHOOK_URL")

    if not gemini_key:
        print("GEMINI_API_KEY が未設定です")
        sys.exit(1)
    if not webhook_url and not args.dry_run:
        print("SLACK_WEBHOOK_URL が未設定です")
        sys.exit(1)

    date_str = args.date or datetime.datetime.now(JST).strftime("%Y-%m-%d")
    print(f"対象日(JST): {date_str}")

    token = get_access_token([CALENDAR_READONLY_SCOPE], subject=CALENDAR_OWNER_EMAIL)
    events = fetch_todays_events(token, CALENDAR_OWNER_EMAIL, date_str)
    companies = group_meetings_by_company(events)

    if not companies:
        print("対象フォーマットの面談なし - 投稿スキップ")
        return

    print(f"本日の面談先: {len(companies)}社")
    results = research_companies(companies, gemini_key)

    if args.dry_run:
        for company, data in results.items():
            print("---")
            print(f"{company}（グラウンディング: {data.get('grounded')}）")
            print(f"面談: {data['meetings']}")
            print(data["text"])
            print(f"参照: {data.get('references')}")
        return

    blocks = build_briefing_blocks(date_str, results)
    post_json(webhook_url, {"blocks": blocks})
    print("Slack投稿完了")


if __name__ == "__main__":
    main()
