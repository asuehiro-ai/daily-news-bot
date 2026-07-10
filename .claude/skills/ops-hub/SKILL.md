---
name: ops-hub
description: このリポジトリで動いている自動化システム（daily-news-bot・meeting-log-syncなど）の状態確認・手動実行・トラブルシューティングを一元的に行う。「〇〇を実行して」「今日のバッチ動いた？」「エラー出てないか確認して」「新しい社員を追加して」など、システム運用に関する依頼があったときに使う。
---

# SKILL: 自動化システム一元管理（ops-hub）

## このSkillの目的

末廣さんの「システムをクロードコードに集約して一元制御したい」という方針のもと、GASに散らばっていた自動化をPython + GitHub Actionsに移行してきた。このSkillは、それらを一箇所から状態確認・手動実行・トラブル対応できるようにするための運用ハブ。

---

## 管理対象システム一覧

| システム | 状態 | 実行方式 | スケジュール | 主要ファイル |
|---|---|---|---|---|
| daily-news-bot | 稼働中 | GitHub Actions | 毎朝6:00 JST | `daily_news.py`, `.github/workflows/daily_news.yml` |
| meeting-log-sync | 稼働中（検証中）。面談ログのスプレッドシート記録＋Slackダイジェスト投稿（顧客発言のみClaude Haikuで要約）を兼ねる | GitHub Actions | 毎朝7:30 JST | `meeting-log-sync/`, `.github/workflows/meeting_log_sync.yml` |
| meeting-briefing-bot | 稼働中（2026-07-10手動実行で動作確認済み） | GitHub Actions | 毎朝6:30 JST | `meeting-briefing-bot/morning_briefing.py`, `.github/workflows/morning_briefing.yml` |
| gmail-automation | 稼働中（2026-07-10手動実行で動作確認済み） | GitHub Actions | 毎日0:00・12:00 JST | `gmail-automation/`, `.github/workflows/gmail_automation.yml` |

`plaud-slack-bot`は2026-07-10にmeeting-log-syncへ完全統合し削除済み（独自のPLAUD再取得・Gemini再要約が二重処理になっていたため）。

面談ログのスプレッドシートには、自動化が書き込む生データの「Sheet1」に加えて、閲覧しやすいように`calendar_event_id`列を除いて日付降順に整形した「面談ログ（閲覧用）」タブがある（QUERY関数で自動追従、`meeting-log-sync/setup_viewer_sheet.py`で追加済み・2026-07-11）。

GAS版（`meeting-briefing-bot/meeting_briefing_bot.gs`, `gmail-automation/gmail_automation.gs`）はまだ残置。末廣さん自身にGASのトリガー停止を依頼すること（Apps ScriptエディタでのUI操作はClaude Codeから実行不可）。

gmail-automationの稼働に必要だった外部作業（完了済み・記録として残す）:
- Google Workspace管理者コンソール → セキュリティ → APIの制御 → ドメイン全体の委任 で、サービスアカウント（`meeting-log-sync@praxis-tractor-461301-v0.iam.gserviceaccount.com`）に`gmail.readonly`・`gmail.send`・`drive`スコープを追加登録
- GCPコンソールでGmail APIを有効化（未有効化だと`Gmail API has not been used in project ...`という403エラーになる）
- 保存先Driveフォルダが共有ドライブ上にあり、Drive APIに`supportsAllDrives=true`が無いと404 File not foundになる問題を`gmail_automation.py`側で修正済み

---

## 使えるコマンド（GitHub CLI）

**重要**: このPCでは`gh`をインストール直後のため、新しいシェルセッションでPATHに反映されていないことがある。`gh`が見つからないときは、フルパス `"C:\Program Files\GitHub CLI\gh.exe"` を使うか、PowerShellなら以下でPATHを更新してから実行する:

```powershell
$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
```

認証済みアカウント: `asuehiro-ai`。デフォルトリポジトリ: `asuehiro-ai/daily-news-bot`（このリポジトリ全体がここにpushされている）。

- ワークフロー一覧: `gh workflow list`
- 直近の実行履歴: `gh run list --limit 10` / 特定ワークフローだけ `gh run list --workflow=meeting_log_sync.yml`
- 実行結果の詳細ログ: `gh run view <run-id> --log`（失敗分だけなら `--log-failed`）
- 手動実行: `gh workflow run <workflow-file>`（例: `gh workflow run meeting_log_sync.yml`）
- 実行の進捗を待つ: `gh run watch <run-id>`
- 登録済みSecret名の一覧（値は取得不可）: `gh secret list`

## このSkillが呼ばれたときの動き方

1. 「状態確認して」系の依頼 → `gh run list --limit 10` で全体をまず俯瞰し、失敗（`failure`）や異常に長い実行がないか確認。あれば `gh run view <id> --log-failed` で原因を特定してから、末廣さんに平易な言葉で要約する（スタックトレースをそのまま貼らない）
2. 「〇〇を今すぐ実行して」系の依頼 → 該当する`gh workflow run`を実行し、`gh run watch`等で結果を確認してから報告する
3. 新しい問題を解決したら、下記「トラブルシューティング履歴」に追記しておく（次回以降のセッションで同じ調査を繰り返さないため）

---

## トラブルシューティング履歴（過去に遭遇した問題と解決策）

- **PLAUD APIが403を返す**: PythonのデフォルトUser-Agent（`Python-urllib/3.x`）だとWAFに弾かれる。ブラウザ相当のUser-Agentを付与する（`meeting-log-sync/plaud_client.py`の`BROWSER_USER_AGENT`）
- **Geminiが404 "model not found"を返す**: `gemini-2.0-flash`は2026/6/1に提供終了。`gemini-2.5-flash`を使う
- **サービスアカウントのJSON鍵が作成できない**: 組織ポリシー`iam.disableServiceAccountKeyCreation`がブロックしている。GCPコンソール「IAMと管理」→「組織のポリシー」で対象プロジェクトへの適用を「オーバーライド」し、強制をオフにする
- **サービスアカウント自身の権限でSheets/Driveファイルを作ると403**: Workspace環境では、サービスアカウント自身の識別子でのファイル作成が権限不足になりやすい。ドメイン全体委任で対象ユーザー（例: 末廣さん）に**なりすまして**作成する方式にする（`get_access_token(scopes, subject=...)`）
- **Sheets APIで「Unable to parse range: Sheet1!...」エラー**: 日本語ロケールのGoogleアカウントでは新規スプレッドシートの既定シート名が「シート1」になり、コード側の決め打ち`"Sheet1"`と一致しない。スプレッドシート作成時に`sheets[].properties.title`で明示的に`"Sheet1"`を指定する
- **GitHub Actionsのジョブがrunner割り当て待ちで数分止まる**: こちらの設定の問題ではなく、GitHub側の障害であることが多い。https://www.githubstatus.com でActionsの状態を確認する
- **Gmail APIが403「has not been used in project ... or it is disabled」を返す**: ドメイン全体委任のスコープ追加とは別に、GCPコンソールでそのAPI自体を有効化する必要がある。エラーメッセージ中のURL（`console.developers.google.com/apis/api/gmail.googleapis.com/overview?project=...`）にアクセスして有効化する
- **Drive APIが404「File not found」を返す（フォルダIDは合っているはず）**: 対象フォルダが共有ドライブ（Shared Drive）上にあり、リクエストに`supportsAllDrives=true`（検索時は`includeItemsFromAllDrives=true`も）が付いていないと、権限があってもマイドライブ扱いの検索で見つからず404になる。`files.list`・`files.create`両方にこのパラメータを付ける
- **エラーが起きてもGitHub Actionsが「成功」表示になる**: スクリプト側がAPIエラーを例外にせず`print`でログ出力するだけの設計だと、exit codeは0のままになる。`gh run list`の成功表示だけで判断せず、`--log`の中身（特に「エラー」「件」等の実行結果サマリ行）を必ず確認すること
- **Gemini再要約でノイズ・コストが増える**: `plaud-slack-bot`が録音1件ごとにGemini APIを呼んでいたが、プロンプトに「顧客発言のみ抽出」の指示がなく自社側発話が混入していた。`meeting-log-sync/plaud_client.py`の`claude_summarize`＋`build_digest_prompt`（Claude Haiku、顧客発言限定プロンプト）に切り替え、`sync_meeting_log.py`に統合して解決（2026-07-10）。`ANTHROPIC_API_KEY`は`daily_news.py`用に既に登録済みのSecretを流用でき、新規登録は不要だった

---

## 認証・設定の状況

- **GitHub CLI**: このPC（末廣さんのWindows PC）にインストール済み・`asuehiro-ai`として認証済み
- **Googleサービスアカウント**: `meeting-log-sync@praxis-tractor-461301-v0.iam.gserviceaccount.com`。ドメイン全体委任で`calendar.readonly`・`spreadsheets`スコープを委任済み。JSON鍵はGitHub Secretsの`GOOGLE_SERVICE_ACCOUNT_JSON`に登録済み（ローカルにも末廣さんのDownloadsフォルダに残っている可能性があるが、扱いは慎重に）
- **対象社員リスト**: `meeting-log-sync/employees.py`（現在: 末廣哲彦、大熊克也）
- **面談ログのスプレッドシート**: 末廣さん所有、SPREADSHEET_IDはGitHub Secretsの`SPREADSHEET_ID`に登録済み

---

## 新しい社員を追加する手順

1. Google Workspace管理者コンソールで、既にサービスアカウントへのドメイン全体委任は全社共通設定なので追加作業は不要（委任はクライアントID単位であり、社員個別の許可は不要）
2. `meeting-log-sync/employees.py`の`EMPLOYEES`リストに`{"name": "氏名", "email": "メールアドレス"}`を追記
3. コミット・プッシュ（末廣さんの確認を取ってから）

## 新しいGAS移行（meeting-briefing-bot / gmail-automation）に着手する手順

1. 既存のサービスアカウント（`meeting-log-sync@praxis-tractor-461301-v0.iam.gserviceaccount.com`）のドメイン全体委任スコープに、必要なGmail系スコープ（`gmail.readonly`, `gmail.send`等）を追加
2. GASロジックをPythonに移植（`meeting_briefing_bot.gs`のGemini検索呼び出し、`gmail_automation.gs`のGmail検索・添付取得ロジックなど）
3. 既存の`meeting-log-sync/google_auth.py`をそのまま流用してアクセストークンを取得
4. 新しいGitHub Actionsワークフローを追加し、GAS側のトリガーは動作確認後に停止する
5. 詳細設計は着手時に改めてプランを立てる

---

## 禁止事項

- Secretsやサービスアカウント鍵の中身を画面に表示・出力しない
- 末廣さんの明示的な確認なしにcommit/push・ワークフロー削除・Secret削除をしない
- GAS側のトリガー停止など外部UI操作が必要なものは、末廣さん自身に依頼する（Claude Codeからは実行できない）
