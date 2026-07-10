"""サービスアカウント（Google Workspaceドメイン全体委任）によるGoogle API認証。

GOOGLE_SERVICE_ACCOUNT_JSON 環境変数にサービスアカウントのJSON鍵の中身をそのまま設定しておく。
"""

import json
import os

import google.auth.transport.requests
from google.oauth2 import service_account

CALENDAR_READONLY_SCOPE = "https://www.googleapis.com/auth/calendar.readonly"
SHEETS_SCOPE = "https://www.googleapis.com/auth/spreadsheets"
DRIVE_SCOPE = "https://www.googleapis.com/auth/drive"
GMAIL_READONLY_SCOPE = "https://www.googleapis.com/auth/gmail.readonly"
GMAIL_SEND_SCOPE = "https://www.googleapis.com/auth/gmail.send"


def _load_service_account_info():
    raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not raw:
        raise RuntimeError("GOOGLE_SERVICE_ACCOUNT_JSON が未設定です")
    return json.loads(raw)


def get_access_token(scopes, subject=None):
    """サービスアカウントのアクセストークンを返す。

    subjectを指定すると、ドメイン全体委任でその社員（メールアドレス）になりすまして
    Calendar等にアクセスするトークンになる。省略時はサービスアカウント自身のトークン
    （スプレッドシートの作成・書き込みに使用）。
    """
    info = _load_service_account_info()
    credentials = service_account.Credentials.from_service_account_info(info, scopes=scopes)
    if subject:
        credentials = credentials.with_subject(subject)
    credentials.refresh(google.auth.transport.requests.Request())
    return credentials.token
