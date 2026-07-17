"""PLAUD NOTE API・Gemini API・Slack通知（エラーアラート用）の共通ロジック。

meeting-log-sync配下の各スクリプト（sync_meeting_log.py・backfill_meeting_log.py・
repair_missing_transcripts.py）から使われる。
"""

import datetime
import gzip
import json
import re
import time
import urllib.error
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from zoneinfo import ZoneInfo

import anthropic

DEFAULT_PLAUD_BASE_URL = "https://api.plaud.ai"
GEMINI_MODEL = "gemini-2.5-flash"  # gemini-2.0-flashは2026/6/1に提供終了済み
CLAUDE_MODEL = "claude-haiku-4-5-20251001"
JST = ZoneInfo("Asia/Tokyo")
MAX_PAGES = 3
PARALLEL_WORKERS = 8

# PLAUDのAPIはPython標準のUser-Agent（Python-urllib/3.x）だと403で弾くため、
# ブラウザ経由のアクセスに見せかける（GASのUrlFetchAppはこの問題が起きない）。
BROWSER_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

TOKEN_EXPIRED_MESSAGE = (
    "⚠️ *PLAUDトークン期限切れ*\n"
    "web.plaud.ai → DevTools → Network → Authorization ヘッダーから新しいトークンを取得し、"
    "GitHub Secretsの PLAUD_TOKEN を更新してください。"
)


class PlaudTokenExpiredError(Exception):
    pass


# ── HTTP ヘルパー ────────────────────────────────────────────


def http_request(url, *, method="GET", headers=None, payload=None, timeout=30):
    """(status_code, body_bytes, response_headers) を返す。接続自体の失敗時は (None, None, None)。"""
    data = json.dumps(payload).encode("utf-8") if payload is not None else None
    merged_headers = {"User-Agent": BROWSER_USER_AGENT, **(headers or {})}
    req = urllib.request.Request(url, data=data, headers=merged_headers, method=method)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return resp.getcode(), resp.read(), resp.headers
    except urllib.error.HTTPError as e:
        return e.code, e.read(), e.headers
    except urllib.error.URLError as e:
        print(f"Fetch エラー: {e}")
        return None, None, None


def plaud_get(token, url):
    code, body, _ = http_request(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        },
    )
    if code is None:
        return None
    if code == 401:
        raise PlaudTokenExpiredError(f"PLAUD_TOKEN が期限切れです: {url}")
    if code != 200:
        print(f"API エラー {code}: {url}")
        return None
    return json.loads(body.decode("utf-8"))


def gemini_generate(prompt, api_key, *, max_output_tokens=800, temperature=0.2):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent?key={api_key}"
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": temperature, "maxOutputTokens": max_output_tokens},
    }
    code, body, _ = http_request(
        url, method="POST", headers={"Content-Type": "application/json"}, payload=payload, timeout=60
    )
    if code is None:
        return None
    if code != 200:
        print(f"Gemini エラー: HTTP {code} {body[:300] if body else ''}")
        return None
    try:
        result = json.loads(body.decode("utf-8"))
        return result["candidates"][0]["content"]["parts"][0]["text"].strip()
    except (KeyError, IndexError, TypeError):
        return None


def post_json(url, payload):
    code, _, _ = http_request(url, method="POST", headers={"Content-Type": "application/json"}, payload=payload)
    if code is None or code >= 300:
        print(f"投稿エラー: HTTP {code}")


def alert_slack(webhook_url, message):
    if not webhook_url:
        print(message)
        return
    post_json(webhook_url, {"text": message})


# ── 日時ユーティリティ ────────────────────────────────────────


def get_yesterday_jst_str():
    return (datetime.datetime.now(JST) - datetime.timedelta(days=1)).strftime("%Y-%m-%d")


def _parse_time_value(value):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        seconds = value / 1000 if value > 1e11 else value
        return datetime.datetime.fromtimestamp(seconds, tz=datetime.timezone.utc)
    try:
        return datetime.datetime.fromisoformat(str(value).replace("Z", "+00:00"))
    except ValueError:
        return None


def to_jst_date_str(value):
    d = _parse_time_value(value)
    return d.astimezone(JST).strftime("%Y-%m-%d") if d else None


def to_jst_time_str(value):
    d = _parse_time_value(value)
    return d.astimezone(JST).strftime("%H:%M") if d else ""


def to_jst_datetime(value):
    d = _parse_time_value(value)
    return d.astimezone(JST) if d else None


# ── PLAUD API ────────────────────────────────────────────────


def to_array(data):
    if not data:
        return []
    if isinstance(data, list):
        return data
    for key in ("data_file_list", "data", "list", "items", "result"):
        val = data.get(key)
        if isinstance(val, list):
            return val
    inner = data.get("data")
    if isinstance(inner, dict) and isinstance(inner.get("list"), list):
        return inner["list"]
    return []


def fetch_recordings_by_date(token, base_url, date_str):
    """PLAUD側の対象日(JST)の録音一覧を返す。dateStrは"YYYY-MM-DD"。"""
    recordings = []
    seen_ids = set()

    for page in range(1, MAX_PAGES + 1):
        data = plaud_get(token, f"{base_url}/file/simple/web?page={page}&page_size=50")
        if not data:
            break
        files = to_array(data)
        if not files:
            break

        any_recent_file = False
        for f in files:
            file_date_str = to_jst_date_str(f.get("start_time") or f.get("create_time") or f.get("created_at"))
            if not file_date_str:
                continue
            if file_date_str >= date_str:
                any_recent_file = True
            if file_date_str == date_str and f.get("id") not in seen_ids:
                seen_ids.add(f.get("id"))
                recordings.append(f)

        if not any_recent_file or len(files) < 50:
            break
        if page < MAX_PAGES:
            time.sleep(0.3)

    return recordings


def fetch_recordings_in_range(token, base_url, start_date_str, end_date_str, max_pages=60):
    """PLAUD側の[start_date_str, end_date_str]（JST、両端含む）の録音一覧を返す。過去分の一括バックフィル用。"""
    recordings = []
    seen_ids = set()

    for page in range(1, max_pages + 1):
        data = plaud_get(token, f"{base_url}/file/simple/web?page={page}&page_size=50")
        if not data:
            break
        files = to_array(data)
        if not files:
            break

        any_in_or_after_start = False
        for f in files:
            file_date_str = to_jst_date_str(f.get("start_time") or f.get("create_time") or f.get("created_at"))
            if not file_date_str:
                continue
            if file_date_str >= start_date_str:
                any_in_or_after_start = True
            if start_date_str <= file_date_str <= end_date_str and f.get("id") not in seen_ids:
                seen_ids.add(f.get("id"))
                recordings.append(f)

        if not any_in_or_after_start or len(files) < 50:
            break
        if page < max_pages:
            time.sleep(0.3)

    return recordings


def fetch_all_summary_items(token, base_url, recordings):
    if not recordings:
        return []

    with ThreadPoolExecutor(max_workers=PARALLEL_WORKERS) as ex:
        details = list(ex.map(lambda rec: plaud_get(token, f"{base_url}/file/detail/{rec['id']}"), recordings))

    s3_jobs = []  # (recording_index, s3_url)
    for i, detail in enumerate(details):
        if not detail:
            continue
        content_list = (detail.get("data") or {}).get("content_list") or []
        item = next(
            (c for c in content_list if c.get("data_type") == "auto_sum_note" and c.get("task_status") == 1 and c.get("data_link")),
            None,
        )
        if item:
            s3_jobs.append((i, item["data_link"]))

    if not s3_jobs:
        return []

    with ThreadPoolExecutor(max_workers=PARALLEL_WORKERS) as ex:
        texts = list(ex.map(lambda job: fetch_s3_summary(job[1]), s3_jobs))

    items = []
    for (rec_index, _), text in zip(s3_jobs, texts):
        if not text or not text.strip():
            continue
        rec = recordings[rec_index]
        items.append(
            {
                "title": rec.get("filename") or rec.get("title") or rec.get("name") or "無題",
                "createTime": rec.get("start_time") or rec.get("create_time"),
                "summary": text,
            }
        )
    return items


def fetch_s3_summary(url):
    code, body, headers = http_request(url)
    if code is None or code != 200:
        return None
    if headers and headers.get("Content-Encoding") == "gzip":
        try:
            body = gzip.decompress(body)
        except OSError:
            pass
    try:
        return body.decode("utf-8")
    except UnicodeDecodeError:
        return None


# ── Claude Haiku による面談ログ要約（非参加者にも伝わるレベル） ─────


def build_log_summary_prompt(cleaned_text):
    return (
        "以下は商談・会議の録音を自動要約したテキストです。話者ラベルが不正確な場合があるため、"
        "発言者の呼称ではなく内容から「相手企業側の発言・状況」と「LEG側の説明・提案」を判断してください。"
        "この会議に参加していない社内メンバーが読んでも状況を理解できるレベルで、"
        "全体を300〜500字程度に要約してください。\n"
        "【盛り込む内容】\n"
        "・相手企業の状況・課題・要望\n"
        "・LEG側の説明・提案の要点\n"
        "・決定事項\n"
        "・次回アクション（誰が・何を・いつまでに。読み取れない場合は「要確認」と明記）\n"
        "・見出し番号や記号は付けず、自然な文章でまとめる\n"
        "・固有名詞・数値・期日は省略しない\n"
        "・元のテキストにない事実を創作しない\n"
        "・前置きや後書きは不要。要約本文のみ出力\n\n" + cleaned_text
    )


def claude_summarize(text, api_key, *, max_tokens=800):
    try:
        client = anthropic.Anthropic(api_key=api_key)
        response = client.messages.create(
            model=CLAUDE_MODEL, max_tokens=max_tokens, messages=[{"role": "user", "content": text}]
        )
        return response.content[0].text.strip()
    except Exception as e:
        print(f"Claude エラー: {e}")
        return None


# ── テキスト整形 ────────────────────────────────────────────


def clean_summary(raw_text):
    text = re.sub(r"!\[[^\]]*\]\([^)]*\)", "", raw_text)
    text = re.sub(r"\[[^\]]*(?:挿入|Insert|Speaker\s*\d*)[^\]]*\]", "", text, flags=re.IGNORECASE)
    text = re.sub(r"^#{1,4}\s*", "", text, flags=re.MULTILINE)
    text = re.sub(r"^>\s*", "", text, flags=re.MULTILINE)
    text = re.sub(r"^[-*]{3,}$", "", text, flags=re.MULTILINE)
    text = text.replace("**", "")

    lines = []
    for line in text.split("\n"):
        line = line.strip()
        if not line:
            continue
        if re.match(r"^(?:日時|日付と時刻|場所|顧客|参加者|顧客名)\s*[:：]", line):
            continue
        if re.match(r"^(?:概要|会議情報|会議メモ)\s*[:：]?\s*$", line):
            continue
        lines.append(line)

    result = "\n".join(lines)
    return result[:2500] + "…" if len(result) > 2500 else result


def extract_key_points(raw_text):
    cleaned = clean_summary(raw_text)
    lines = [l.strip() for l in cleaned.split("\n") if l.strip()]

    bullets = [l for l in lines if re.match(r"^[-•・]", l) or re.match(r"^\d+[.．]", l)]
    if len(bullets) >= 3:
        return "\n".join(bullets[:10])

    meaningful = [l for l in lines if len(l) > 15]
    return "\n".join(meaningful[:10])
