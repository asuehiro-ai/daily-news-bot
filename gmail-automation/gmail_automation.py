#!/usr/bin/env python3
"""請求書メール添付のDrive保存・未返信/添付付きメールの通知（gmail_automation.gsのPython移植）。

必須環境変数: GOOGLE_SERVICE_ACCOUNT_JSON

GAS版はPropertiesServiceで処理済みメッセージIDを保持していたが、実際の重複防止は
Drive側のファイル名存在チェックだけで完結していたためID管理は廃止した（振る舞いは変えない）。
未返信・添付付きメール通知は元々ID管理をしておらず、条件に合致する限り毎回通知する
（答えるまで催促し続ける）仕様をそのまま踏襲している。
"""

import argparse
import base64
import datetime
import email.utils
import json
import mimetypes
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
from email.header import Header
from email.mime.text import MIMEText
from zoneinfo import ZoneInfo

from config import FOLDER_ID, OWNER_EMAIL, REPLY_THRESHOLD_HOURS

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "meeting-log-sync"))

from google_auth import DRIVE_SCOPE, GMAIL_READONLY_SCOPE, GMAIL_SEND_SCOPE, get_access_token  # noqa: E402

JST = ZoneInfo("Asia/Tokyo")
INVOICE_EXT_PATTERN = re.compile(r"\.(pdf|xlsx?|csv|png|jpe?g)$", re.IGNORECASE)

AUTOMATED_PATTERNS = [
    re.compile(p, re.IGNORECASE)
    for p in [
        r"no.?reply",
        r"noreply",
        r"donotreply",
        r"newsletter",
        r"notification",
        r"メルマガ",
        r"配信停止",
        r"unsubscribe",
        r"auto.?reply",
        r"自動返信",
        r"bounce",
        r"mailer-daemon",
        r"info@",
        r"news@",
        r"support@",
        r"system@",
    ]
]


# ── Google API 共通ヘルパー ──────────────────────────────────


def google_get(token, url):
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"})
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        print(f"APIエラー: HTTP {e.code} {e.read().decode('utf-8')[:300]}")
        return None


# ── Gmail API ─────────────────────────────────────────────────


def gmail_search_threads(token, query, max_pages=3):
    threads = []
    page_token = None
    for _ in range(max_pages):
        url = f"https://gmail.googleapis.com/gmail/v1/users/me/threads?q={urllib.parse.quote(query)}"
        if page_token:
            url += f"&pageToken={page_token}"
        data = google_get(token, url)
        if not data:
            break
        threads.extend(data.get("threads") or [])
        page_token = data.get("nextPageToken")
        if not page_token:
            break
    return threads


def gmail_get_thread(token, thread_id):
    url = f"https://gmail.googleapis.com/gmail/v1/users/me/threads/{thread_id}?format=full"
    return google_get(token, url)


def fetch_attachment_data(token, message_id, attachment_id):
    url = f"https://gmail.googleapis.com/gmail/v1/users/me/messages/{message_id}/attachments/{attachment_id}"
    data = google_get(token, url)
    if not data or "data" not in data:
        return None
    raw = data["data"]
    return base64.urlsafe_b64decode(raw + "=" * (-len(raw) % 4))


def send_email(token, to_email, from_email, subject, body):
    msg = MIMEText(body, _charset="utf-8")
    msg["To"] = to_email
    msg["From"] = from_email
    msg["Subject"] = str(Header(subject, "utf-8"))
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")

    req = urllib.request.Request(
        "https://gmail.googleapis.com/gmail/v1/users/me/messages/send",
        data=json.dumps({"raw": raw}).encode("utf-8"),
        method="POST",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            resp.read()
    except urllib.error.HTTPError as e:
        print(f"メール送信エラー: HTTP {e.code} {e.read().decode('utf-8')[:300]}")


# ── メッセージ解析ヘルパー ───────────────────────────────────


def get_header(message, name):
    headers = (message.get("payload") or {}).get("headers", [])
    for h in headers:
        if h.get("name", "").lower() == name.lower():
            return h.get("value", "")
    return ""


def get_message_datetime(message):
    date_header = get_header(message, "Date")
    if date_header:
        try:
            dt = email.utils.parsedate_to_datetime(date_header)
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=datetime.timezone.utc)
            return dt.astimezone(JST)
        except (TypeError, ValueError):
            pass
    internal_date = message.get("internalDate")
    if internal_date:
        return datetime.datetime.fromtimestamp(int(internal_date) / 1000, tz=datetime.timezone.utc).astimezone(JST)
    return None


def walk_attachment_parts(payload):
    for part in payload.get("parts") or []:
        filename = part.get("filename")
        body = part.get("body") or {}
        if filename and body.get("attachmentId"):
            yield part
        if part.get("parts"):
            yield from walk_attachment_parts(part)


def is_automated_email(message):
    from_header = get_header(message, "From")
    subject = get_header(message, "Subject")
    return any(p.search(from_header) or p.search(subject) for p in AUTOMATED_PATTERNS)


# ── Drive API ─────────────────────────────────────────────────


def drive_file_exists(token, filename):
    escaped = filename.replace("\\", "\\\\").replace("'", "\\'")
    query = f"name = '{escaped}' and '{FOLDER_ID}' in parents and trashed = false"
    url = (
        f"https://www.googleapis.com/drive/v3/files?q={urllib.parse.quote(query)}"
        "&fields=files(id)&supportsAllDrives=true&includeItemsFromAllDrives=true"
    )
    data = google_get(token, url)
    return bool(data and data.get("files"))


def drive_upload(token, filename, mime_type, file_bytes):
    boundary = "gmailautomationboundary"
    metadata = json.dumps({"name": filename, "parents": [FOLDER_ID]})
    body = (
        f"--{boundary}\r\n"
        "Content-Type: application/json; charset=UTF-8\r\n\r\n"
        f"{metadata}\r\n"
        f"--{boundary}\r\n"
        f"Content-Type: {mime_type}\r\n\r\n"
    ).encode("utf-8") + file_bytes + f"\r\n--{boundary}--".encode("utf-8")

    req = urllib.request.Request(
        "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true",
        data=body,
        method="POST",
        headers={"Authorization": f"Bearer {token}", "Content-Type": f"multipart/related; boundary={boundary}"},
    )
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            resp.read()
    except urllib.error.HTTPError as e:
        print(f"Driveアップロードエラー: HTTP {e.code} {e.read().decode('utf-8')[:300]}")


# ── 請求書メール添付の保存 ────────────────────────────────────


def save_invoices_to_drive(token, dry_run):
    query = "(subject:請求書 OR subject:invoice) newer_than:30d"
    threads = gmail_search_threads(token, query)
    saved_count = 0

    for t in threads:
        full = gmail_get_thread(token, t["id"])
        if not full:
            continue

        for message in full.get("messages", []):
            msg_date = get_message_datetime(message)
            date_prefix = msg_date.strftime("%Y%m%d") if msg_date else "00000000"

            for part in walk_attachment_parts(message.get("payload", {})):
                name = part.get("filename") or ""
                if not INVOICE_EXT_PATTERN.search(name):
                    continue
                safe_name = f"{date_prefix}_{name}"

                if drive_file_exists(token, safe_name):
                    continue

                if dry_run:
                    print(f"[dry-run] 保存対象: {safe_name}")
                    saved_count += 1
                    continue

                data = fetch_attachment_data(token, message["id"], part["body"]["attachmentId"])
                if not data:
                    continue
                mime_type = part.get("mimeType") or mimetypes.guess_type(name)[0] or "application/octet-stream"
                drive_upload(token, safe_name, mime_type, data)
                saved_count += 1
                print(f"保存完了: {safe_name}")

    print(f"請求書保存: {saved_count}件")


# ── 未返信・添付付きメールの通知 ─────────────────────────────


def get_unanswered_emails(token, my_email):
    cutoff = datetime.datetime.now(JST) - datetime.timedelta(hours=REPLY_THRESHOLD_HOURS)
    query = "in:inbox -category:promotions -category:social -category:updates newer_than:7d"
    threads = gmail_search_threads(token, query)

    unanswered = []
    for t in threads:
        full = gmail_get_thread(token, t["id"])
        messages = full.get("messages") if full else None
        if not messages:
            continue

        last_msg = messages[-1]
        from_header = get_header(last_msg, "From")
        if my_email.lower() in from_header.lower():
            continue
        if is_automated_email(last_msg):
            continue

        received = get_message_datetime(last_msg)
        if not received or received >= cutoff:
            continue

        unanswered.append(
            {
                "subject": get_header(messages[0], "Subject"),
                "from": from_header,
                "received": received,
                "url": f"https://mail.google.com/mail/u/0/#inbox/{t['id']}",
            }
        )
    return unanswered


def get_emails_with_attachments(token):
    query = "in:inbox has:attachment -category:promotions -category:social -category:updates newer_than:7d"
    threads = gmail_search_threads(token, query)

    result = []
    for t in threads:
        full = gmail_get_thread(token, t["id"])
        messages = full.get("messages") if full else None
        if not messages:
            continue

        last_msg = messages[-1]
        if is_automated_email(last_msg):
            continue

        attachment_names = [
            part.get("filename")
            for message in messages
            for part in walk_attachment_parts(message.get("payload", {}))
        ]

        if attachment_names:
            result.append(
                {
                    "subject": get_header(messages[0], "Subject"),
                    "from": get_header(last_msg, "From"),
                    "received": get_message_datetime(last_msg),
                    "attachments": attachment_names,
                    "url": f"https://mail.google.com/mail/u/0/#inbox/{t['id']}",
                }
            )
    return result


def build_digest_body(unanswered, with_attachments):
    body = ""
    if unanswered:
        body += f"■ {REPLY_THRESHOLD_HOURS}時間以上返信していないメール（{len(unanswered)}件）\n"
        body += "─" * 50 + "\n\n"
        for i, item in enumerate(unanswered, 1):
            received_str = item["received"].strftime("%Y/%m/%d %H:%M") if item["received"] else ""
            body += f"【{i}】{item['subject']}\n"
            body += f"　送信者 : {item['from']}\n"
            body += f"　受信日時: {received_str}\n"
            body += f"　リンク  : {item['url']}\n\n"

    if with_attachments:
        if body:
            body += "\n"
        body += f"■ 添付ファイル付きメール（{len(with_attachments)}件）\n"
        body += "─" * 50 + "\n\n"
        for i, item in enumerate(with_attachments, 1):
            received_str = item["received"].strftime("%Y/%m/%d %H:%M") if item["received"] else ""
            body += f"【{i}】{item['subject']}\n"
            body += f"　送信者 : {item['from']}\n"
            body += f"　受信日時: {received_str}\n"
            body += f"　添付   : {'、'.join(item['attachments'])}\n"
            body += f"　リンク  : {item['url']}\n\n"

    return body


def check_and_notify(token, dry_run):
    unanswered = get_unanswered_emails(token, OWNER_EMAIL)
    with_attachments = get_emails_with_attachments(token)

    if not unanswered and not with_attachments:
        print("通知対象なし")
        return

    if dry_run:
        print(f"[dry-run] 未返信{len(unanswered)}件、添付付き{len(with_attachments)}件（メール送信スキップ）")
        return

    now = datetime.datetime.now(JST).strftime("%Y/%m/%d %H:%M")
    sections = []
    if unanswered:
        sections.append(f"未返信{len(unanswered)}件")
    if with_attachments:
        sections.append(f"添付付き{len(with_attachments)}件")
    subject = f"【要確認】{'・'.join(sections)} ({now})"
    body = build_digest_body(unanswered, with_attachments)

    send_email(token, OWNER_EMAIL, OWNER_EMAIL, subject, body)
    print(f"通知送信: 未返信{len(unanswered)}件、添付付き{len(with_attachments)}件")


# ── メイン処理 ────────────────────────────────────────────────


def main():
    if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
        sys.stdout.reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true", help="Drive保存・メール送信をせずログ出力のみ行う")
    args = parser.parse_args()

    token = get_access_token([GMAIL_READONLY_SCOPE, GMAIL_SEND_SCOPE, DRIVE_SCOPE], subject=OWNER_EMAIL)

    save_invoices_to_drive(token, args.dry_run)
    check_and_notify(token, args.dry_run)


if __name__ == "__main__":
    main()
