// ================================================================
// 朝の面談ブリーフィング → Slack 自動投稿
// ================================================================
//
// 【予定タイトルのフォーマット】
//   【Web】●●株式会社 山田様
//   【神田】株式会社●● 佐藤取締役
//   【大阪】●●有限会社 鈴木社長（田中様ご紹介）
//
//   ・【】＝ 面談方法／場所（Web・地名など）
//   ・】の直後 ＝ 社名（半角スペースの前まで）
//   ・半角スペースの後 ＝ 面談相手の氏名・役職
//
// 【セットアップ手順】
//
// ■ Step 1: Slack Incoming Webhook URLを取得
//   Slackアプリ管理画面 → Incoming Webhooks → 通知したいチャンネルを選んでURLをコピー
//
// ■ Step 2: スクリプトプロパティを設定
//   GASエディタ → 左サイドバー「プロジェクトの設定」→ スクリプトプロパティ
//   以下を追加:
//     SLACK_WEBHOOK_URL : Step1のURL
//     GEMINI_API_KEY    : Gemini APIキー
//
// ■ Step 3: 動作テスト
//   1. testDryRun()  を実行 → Slack投稿なしでログに抽出結果・リサーチ内容を出力
//   2. testWithDate('2026-07-08') を実行 → 指定日の予定で実際にSlackへテスト投稿
//   3. 問題なければ sendMorningBriefing() を手動実行して最終確認
//
// ■ Step 4: トリガーを設定
//   GASエディタ → 左サイドバー「トリガー」→ 追加
//   関数: sendMorningBriefing / イベント: 時間主導型 / 毎日 / 午前6〜7時
//
// ■ 注意（Google検索グラウンディングについて）
//   GeminiのGoogle検索グラウンディング機能はAPIキーのプランによっては使えない場合があります。
//   その場合は自動的に検索なしのプロンプトにフォールバックし、
//   本文に「一般知識ベースの内容」である旨の注記を付けます。
// ================================================================

const GEMINI_MODEL = 'gemini-2.5-flash'; // gemini-2.0-flashは2026/6/1に提供終了済み

// ================================================================
// メイン処理
// ================================================================

function sendMorningBriefing() {
  const props = PropertiesService.getScriptProperties();
  const webhookUrl = props.getProperty('SLACK_WEBHOOK_URL');
  const geminiKey  = props.getProperty('GEMINI_API_KEY');

  if (!webhookUrl) { Logger.log('SLACK_WEBHOOK_URL が未設定です'); return; }
  if (!geminiKey)  { alertSlack_(webhookUrl, '⚠️ GEMINI_API_KEY が未設定です'); return; }

  const today = new Date();
  const companies = getTodaysMeetings_(today);

  if (companies.size === 0) {
    console.log('本日は対象フォーマットの面談なし - 投稿スキップ');
    return;
  }

  console.log('本日の面談先: ' + companies.size + '社');

  const results = researchCompanies_(companies, geminiKey);
  postBriefingToSlack_(webhookUrl, today, results);
  console.log('Slack投稿完了');
}

// ================================================================
// カレンダー読み込み・タイトル解析
// ================================================================

/**
 * 今日の予定から対象フォーマットのものを抽出し、会社名ごとにまとめて返す。
 * 戻り値: Map<会社名, { meetings: [{ time, tag, person }] }>
 */
function getTodaysMeetings_(date) {
  const events = CalendarApp.getDefaultCalendar().getEventsForDay(date);
  const companies = new Map();

  for (const event of events) {
    if (event.isAllDayEvent()) continue;

    const parsed = parseEventTitle_(event.getTitle());
    if (!parsed) continue;

    if (!companies.has(parsed.company)) {
      companies.set(parsed.company, { meetings: [] });
    }
    companies.get(parsed.company).meetings.push({
      time:   Utilities.formatDate(event.getStartTime(), 'Asia/Tokyo', 'HH:mm'),
      tag:    parsed.tag,
      person: parsed.person
    });
  }

  return companies;
}

/**
 * 「【Web】●●株式会社 山田様」形式のタイトルを解析する。
 * フォーマットに一致しない場合は null を返す。
 */
function parseEventTitle_(title) {
  // 社名は「】」の直後から、半角スペースまたは全角括弧（紹介者情報など）の手前までとする
  const m = String(title || '').match(/^【([^】]*)】\s*([^\s（(]+)\s*(.*)$/);
  if (!m) return null;
  return { tag: m[1].trim(), company: m[2].trim(), person: m[3].trim() };
}

// ================================================================
// Gemini リサーチ（Google検索グラウンディング）
// ================================================================

function researchCompanies_(companies, apiKey) {
  const results = new Map();
  for (const [company, data] of companies) {
    console.log('リサーチ中: ' + company);
    const research = researchCompany_(company, apiKey);
    results.set(company, Object.assign({}, data, research));
  }
  return results;
}

function researchCompany_(companyName, apiKey) {
  const prompt =
    '「' + companyName + '」という企業についてWeb検索を行い、以下の2点を日本語でまとめてください。\n\n' +
    '① 会社概要（事業内容・規模・沿革・最近のニュースやトピック）\n' +
    '② 業界動向（この会社が属する業界の直近の動向・トレンド）\n\n' +
    '出力ルール:\n' +
    '- ①②それぞれ見出しをそのまま使い、箇条書き（- で始める）で3〜5項目ずつ\n' +
    '- 各項目は1〜2文で簡潔に、固有名詞・数値は省略しない\n' +
    '- 前置きや結びの言葉は不要、見出しと箇条書きのみ出力\n' +
    '- 検索してもこの名称の企業が特定できない場合、①には社名の文字列（業種を示す語・カタカナ英語表記など）から推測できる事業内容の仮説を1〜2項目で記載する\n' +
    '- ②は企業を特定できなかった場合でも空欄にせず、社名や①の仮説から推測される業界の一般的な直近動向を必ず記載する。その場合は②の先頭に「（企業を特定できなかったため、社名から推測した業界の動向です）」と一言添える';

  let result = callGemini_(prompt, apiKey, true);
  if (!result) {
    console.log(companyName + ': グラウンディング検索に失敗、フォールバックします');
    result = callGemini_(prompt, apiKey, false);
    if (result) result.grounded = false;
  }

  if (!result) {
    return { text: '情報を取得できませんでした。', references: [], grounded: false };
  }
  return result;
}

/**
 * Gemini generateContent を呼び出す。
 * withSearch=true の場合は google_search ツールを付けてグラウンディングする。
 * 失敗時は null を返す（呼び出し側でフォールバック判断）。
 */
function callGemini_(prompt, apiKey, withSearch) {
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent?key=' + apiKey;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.2,
      maxOutputTokens: 4096,
      thinkingConfig: { thinkingBudget: 0 } // 内部思考トークンで出力が尻切れになるのを防ぐ
    }
  };
  if (withSearch) payload.tools = [{ google_search: {} }];

  try {
    const resp = UrlFetchApp.fetch(url, {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (resp.getResponseCode() !== 200) {
      console.log('Gemini エラー(' + (withSearch ? '検索あり' : '検索なし') + '): HTTP ' + resp.getResponseCode() + ' ' + resp.getContentText().substring(0, 300));
      return null;
    }

    const result = JSON.parse(resp.getContentText());
    const candidate = result.candidates && result.candidates[0];
    if (candidate && candidate.finishReason === 'MAX_TOKENS') {
      console.log('Gemini 警告: maxOutputTokensに到達し出力が途中で切れています');
    }
    const text = candidate && candidate.content && candidate.content.parts &&
                 candidate.content.parts[0] && candidate.content.parts[0].text;
    if (!text) return null;

    const chunks = (candidate.groundingMetadata && candidate.groundingMetadata.groundingChunks) || [];
    const seen = new Set();
    const references = [];
    for (const c of chunks) {
      if (c.web && c.web.uri && !seen.has(c.web.uri)) {
        seen.add(c.web.uri);
        references.push({ title: c.web.title || c.web.uri, uri: c.web.uri });
      }
    }

    return { text: text.trim(), references: references, grounded: withSearch };

  } catch (e) {
    console.log('Gemini 呼び出し例外: ' + e.message);
    return null;
  }
}

// ================================================================
// Slack 投稿
// ================================================================

function postBriefingToSlack_(webhookUrl, date, results) {
  const dateLabel = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd');
  const blocks = [
    {
      type: 'header',
      text: { type: 'plain_text', text: '📅 本日の面談ブリーフィング（' + dateLabel + '・' + results.size + '社）', emoji: true }
    }
  ];

  let first = true;
  for (const [company, data] of results) {
    if (!first) blocks.push({ type: 'divider' });
    first = false;

    const meetingLines = data.meetings.map(function(m) {
      return '・' + m.time + '　' + m.tag + '　' + (m.person || '(相手名不明)');
    }).join('\n');

    let body = '*' + company + '*\n' + meetingLines + '\n\n' + truncateForSlack_(data.text);
    if (data.grounded === false) {
      body += '\n_※検索グラウンディングに失敗したため一般知識ベースの内容です_';
    }

    blocks.push({ type: 'section', text: { type: 'mrkdwn', text: body } });

    if (data.references && data.references.length > 0) {
      const refText = '参照: ' + data.references.slice(0, 5).map(function(r) {
        return '<' + r.uri + '|' + r.title + '>';
      }).join(' / ');
      blocks.push({ type: 'context', elements: [{ type: 'mrkdwn', text: refText }] });
    }
  }

  UrlFetchApp.fetch(webhookUrl, {
    method: 'POST',
    contentType: 'application/json',
    payload: JSON.stringify({ blocks: blocks }),
    muteHttpExceptions: true
  });
}

function alertSlack_(webhookUrl, message) {
  if (!webhookUrl) { console.log(message); return; }
  UrlFetchApp.fetch(webhookUrl, {
    method: 'POST',
    contentType: 'application/json',
    payload: JSON.stringify({ text: message }),
    muteHttpExceptions: true
  });
}

function truncateForSlack_(text) {
  return text.length > 2500 ? text.substring(0, 2500) + '…' : text;
}

// ================================================================
// テスト・デバッグ用
// ================================================================

/**
 * Slack投稿なしで、今日の予定から抽出した会社・リサーチ結果をログ出力のみ行う。
 * 本番実行前の内容確認に使用。
 */
function testDryRun() {
  const props = PropertiesService.getScriptProperties();
  const geminiKey = props.getProperty('GEMINI_API_KEY');
  if (!geminiKey) { Logger.log('❌ GEMINI_API_KEY が未設定です'); return; }

  const today = new Date();
  const companies = getTodaysMeetings_(today);
  Logger.log('抽出できた会社数: ' + companies.size);
  for (const [company, data] of companies) {
    Logger.log('--- ' + company + ' ---');
    Logger.log('面談: ' + JSON.stringify(data.meetings));
  }

  if (companies.size === 0) return;

  const results = researchCompanies_(companies, geminiKey);
  for (const [company, data] of results) {
    Logger.log('=== ' + company + '（グラウンディング: ' + data.grounded + '） ===');
    Logger.log(data.text);
    Logger.log('参照: ' + JSON.stringify(data.references));
  }
}

/**
 * 指定した日付（'yyyy-MM-dd'）の予定で実際にSlackへテスト投稿する。
 */
function testWithDate(dateStr) {
  const props = PropertiesService.getScriptProperties();
  const webhookUrl = props.getProperty('SLACK_WEBHOOK_URL');
  const geminiKey  = props.getProperty('GEMINI_API_KEY');
  if (!webhookUrl) { Logger.log('❌ SLACK_WEBHOOK_URL が未設定です'); return; }
  if (!geminiKey)  { Logger.log('❌ GEMINI_API_KEY が未設定です'); return; }

  const targetDate = new Date(dateStr + 'T00:00:00+09:00');
  const companies = getTodaysMeetings_(targetDate);
  Logger.log('対象日: ' + dateStr + ' / 会社数: ' + companies.size);
  if (companies.size === 0) { Logger.log('対象フォーマットの面談なし'); return; }

  const results = researchCompanies_(companies, geminiKey);
  postBriefingToSlack_(webhookUrl, targetDate, results);
  Logger.log('Slack投稿完了');
}
