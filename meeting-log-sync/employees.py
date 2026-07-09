"""面談ログ連携の対象社員リスト。

新しい社員を追加する場合は、EMPLOYEES に {"name": ..., "email": ...} を追記してください。
ドメイン全体委任で各社員のカレンダーを読み取るため、事前にGoogle Workspace管理者コンソールで
サービスアカウントへの委任設定が必要（README/プラン参照）。
"""

EMPLOYEES = [
    {"name": "末廣哲彦", "email": "a.suehiro@le-gr.co.jp"},
    {"name": "大熊克也", "email": "k.okuma@le-gr.co.jp"},
    # {"name": "（社員名）", "email": "（メールアドレス）"},
]

# PLAUD録音を紐付ける対象社員のメールアドレス（現状はPLAUDアカウントを持つ末廣さんのみ）
PLAUD_OWNER_EMAIL = "a.suehiro@le-gr.co.jp"

# スプレッドシートの所有者（サービスアカウントがこの人になりすまして作成・書き込みする）
SPREADSHEET_OWNER_EMAIL = "a.suehiro@le-gr.co.jp"
