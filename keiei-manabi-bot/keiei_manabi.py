#!/usr/bin/env python3
"""経営者の学びBot - 経営者が学ぶべき一般テーマを日替わりでSlackに配信"""

import anthropic
import urllib.request
import json
import datetime
import os

THEMES = [
    "財務・会計の基礎",
    "マーケティング",
    "組織・リーダーシップ",
    "事業承継",
    "M&A",
    "法務・コンプライアンス",
    "歴史・偉人の経営判断から学ぶ教訓",
    "人事・採用",
    "戦略・競争優位",
    "交渉術",
    "危機管理・リスクマネジメント",
    "イノベーション",
    "顧客理解・ブランディング",
]

MODEL = "claude-sonnet-4-6"


def post_to_slack(webhook_url, text):
    data = json.dumps({"text": text}).encode("utf-8")
    req = urllib.request.Request(
        webhook_url, data=data,
        headers={"Content-Type": "application/json"}
    )
    with urllib.request.urlopen(req) as r:
        return r.read().decode()


def generate_content(client, theme, date_str):
    prompt = f"""あなたは中小企業経営者向けの経営教育コンテンツの執筆者です。
本日のテーマ「{theme}」について、経営者が毎朝15分程度で読める学習コンテンツを1本作成してください。

読者像：中小企業のオーナー経営者。事業承継やM&Aにも関心がある。

出力形式（この形式を厳守。マークダウン記号（**や##）は使わずプレーンテキストで）：

[今日のタイトル（テーマ「{theme}」の中の具体的な切り口を反映したもの）]

[導入：なぜ今日この話題を取り上げるか、経営者にとっての関心の入り口を2〜3文で]

ポイント1：[小見出し]
[説明を3〜4文で]

ポイント2：[小見出し]
[説明を3〜4文で]

ポイント3：[小見出し]
[説明を3〜4文で]

（必要であればポイント4・5を同形式で追加。合計3〜5個）

きょうの実務への活かし方：
[今日学んだ内容を自社の経営にどう活かせるか、具体的な問いかけや次の一歩を2〜3文で]

制約：
・全体で日本語2500〜3500字程度（15分程度で読める分量）
・「絶対」「必ず」「100%」等の断定的な保証表現は使わない
・実在の企業名・個人名を挙げて批判・比較しない
・具体的な財務数値を断定しない（「必ず〇億円になる」等はNG）
・比喩やたとえ話を1つ以上使い、経営者が実感を持てるように書く
・本日の日付（{date_str}）や曜日ネタには触れず、テーマそのものの内容に集中する
"""
    response = client.messages.create(
        model=MODEL,
        max_tokens=3000,
        messages=[{"role": "user", "content": prompt}],
    )
    return response.content[0].text.strip()


def main():
    api_key = os.environ["ANTHROPIC_API_KEY"]
    webhook_url = os.environ["SLACK_WEBHOOK_URL"]

    client = anthropic.Anthropic(api_key=api_key)

    jst = datetime.datetime.utcnow() + datetime.timedelta(hours=9)
    date_str = jst.strftime("%Y年%m月%d日")
    day_of_year = jst.timetuple().tm_yday
    theme = THEMES[day_of_year % len(THEMES)]

    print(f"{date_str} のテーマ: {theme}")
    content = generate_content(client, theme, date_str)

    message = (
        f"おはようございます。本日（{date_str}）の経営者の学びです。\n\n"
        f"テーマ：{theme}\n\n"
        f"━━━━━━━━━━━━━━━━\n\n"
        f"{content}"
    )

    result = post_to_slack(webhook_url, message)
    print(f"Slack投稿完了: {result}")


if __name__ == "__main__":
    main()
