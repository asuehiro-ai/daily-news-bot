// ================================================================
// PLAUD NOTE → Slack 毎朝ダイジェスト
// ================================================================
//
// ⚠️ Python版（plaud_slack_bot.py）に移行済み。
// 本番運用はGitHub Actions（.github/workflows/plaud_digest.yml）に切り替え、
// このGASファイルは参考用として残しています。
// Python版の稼働確認後、GASエディタの時間主導型トリガーを停止してください。
//
// 【セットアップ手順】
//
// ■ Step 1: PLAUDトークンを取得
//   1. PCのブラウザで https://web.plaud.ai にログイン
//   2. F12でDevToolsを開く → Networkタブ
//   3. フィルターに「api.plaud」と入力
//   4. ページをリロードするか録音一覧を操作
//   5. api.plaud.ai への任意のリクエストをクリック
//   6. Request Headers の「Authorization」の値を確認
//      例: Bearer eyJhbGc... → "eyJhbGc..." の部分だけコピー
//
// ■ Step 2: Slack Incoming Webhook URLを取得
//   Slackアプリ管理画面 → Incoming Webhooks → URLをコピー
//
// ■ Step 3: スクリプトプロパティを設定
//   GASエディタ → 左サイドバー「プロジェクトの設定」→ スクリプトプロパティ
//   以下を追加:
//     PLAUD_TOKEN       : Step1でコピーしたトークン（Bearerを除いた部分）
//     SLACK_WEBHOOK_URL : Step2のURL
//     GEMINI_API_KEY    : 既存のGemini APIキー
//   ※日本アカウントでapi.plaud.aiが使えない場合は追加:
//     PLAUD_BASE_URL    : https://api-apse1.plaud.ai
//
// ■ Step 4: 動作テスト
//   1. testConnection()  を実行 → ログに録音一覧が出ればOK
//   2. testDryRun()      を実行 → Slack投稿なしで内容を確認
//   3. sendDailyPlaudDigest() を実行 → Slackに実際に投稿
//
// ■ Step 5: トリガーを設定
//   GASエディタ → 左サイドバー「トリガー」→ 追加
//   関数: sendDailyPlaudDigest / イベント: 時間主導型 / 毎日 / 午前8〜9時
//
// ■ トークン更新について
//   トークンが切れるとSlackにアラートが届きます。
//   その場合はStep1の手順でトークンを取り直してプロパティを更新してください。
//   公式APIが解放されたら恒久的な認証に切り替えます。
// ================================================================

const DEFAULT_PLAUD_BASE_URL = 'https://api.plaud.ai';

// ================================================================
// メイン処理
// ================================================================

function sendDailyPlaudDigest() {
  const props = PropertiesService.getScriptProperties();
  const token       = props.getProperty('PLAUD_TOKEN');
  const webhookUrl  = props.getProperty('SLACK_WEBHOOK_URL');
  const geminiKey   = props.getProperty('GEMINI_API_KEY');
  const baseUrl     = props.getProperty('PLAUD_BASE_URL') || DEFAULT_PLAUD_BASE_URL;

  if (!token)      { alertSlack_(webhookUrl, '⚠️ PLAUD_TOKEN が未設定です'); return; }
  if (!webhookUrl) { Logger.log('SLACK_WEBHOOK_URL が未設定です'); return; }

  const yesterdayStr = getYesterdayJstStr_();
  console.log('対象日(JST): ' + yesterdayStr);

  // 前日の録音一覧を取得
  const recordings = fetchRecordingsByDate_(token, baseUrl, yesterdayStr);

  if (recordings.length === 0) {
    console.log('前日の録音なし - 通知スキップ');
    return;
  }

  console.log('録音件数: ' + recordings.length + '件');

  // 全録音の要約を並列取得
  const items = fetchAllSummaryItems_(token, baseUrl, recordings);

  if (items.length === 0) {
    console.log('取得できた要約なし');
    return;
  }

  // 各録音をGeminiで個別に500〜600字へ要約（APIキーがある場合）
  const finalItems = geminiKey ? summarizeItemsWithGemini_(items, geminiKey) : items;

  // Slackに投稿
  postToSlack_(webhookUrl, yesterdayStr, finalItems, null);
  console.log('Slack投稿完了');
}

// ================================================================
// PLAUD API
// ================================================================

function fetchRecordingsByDate_(token, baseUrl, yesterdayStr) {
  const recordings = [];
  const seenIds = new Set(); // ページまたぎ重複を除外
  const MAX_PAGES = 3; // 最大3ページ（150件）で打ち切り

  for (let page = 1; page <= MAX_PAGES; page++) {
    const url = baseUrl + '/file/simple/web?page=' + page + '&page_size=50';
    const data = plaudGet_(token, url);
    if (!data) break;

    const files = toArray_(data);
    if (files.length === 0) break;

    // APIはedit_time順（処理日時順）のため start_time がバラバラに混在する。
    // ページ内を全件スキャンして前日分を収集する。
    let anyRecentFile = false;
    for (const file of files) {
      const dateStr = toJstDateStr_(file.start_time || file.create_time || file.created_at);
      if (!dateStr) continue;
      if (dateStr >= yesterdayStr) anyRecentFile = true;
      if (dateStr === yesterdayStr && !seenIds.has(file.id)) {
        seenIds.add(file.id);
        recordings.push(file);
      }
    }

    // このページに前日以降のファイルが1件もなければ終了
    if (!anyRecentFile) break;
    if (files.length < 50) break;
    if (page < MAX_PAGES) Utilities.sleep(300);
  }

  return recordings;
}

function fetchAllSummaryItems_(token, baseUrl, recordings) {
  if (recordings.length === 0) return [];

  const authHeader = { 'Authorization': 'Bearer ' + token, 'Accept': 'application/json' };

  // ① 全録音のfile/detailを並列取得
  const detailReqs = recordings.map(function(rec) {
    return { url: baseUrl + '/file/detail/' + rec.id, method: 'GET', headers: authHeader, muteHttpExceptions: true };
  });
  const detailResps = UrlFetchApp.fetchAll(detailReqs);

  // ② S3 URLを収集
  const s3Jobs = []; // { recIndex, url }
  detailResps.forEach(function(resp, i) {
    if (resp.getResponseCode() !== 200) return;
    try {
      const contentList = (JSON.parse(resp.getContentText()).data || {}).content_list || [];
      const item = contentList.find(function(c) {
        return c.data_type === 'auto_sum_note' && c.task_status === 1 && c.data_link;
      });
      if (item) s3Jobs.push({ recIndex: i, url: item.data_link });
    } catch(e) { console.log('detail parse error: ' + e.message); }
  });

  if (s3Jobs.length === 0) return [];

  // ③ S3の要約ファイルを並列ダウンロード
  const s3Resps = UrlFetchApp.fetchAll(s3Jobs.map(function(j) {
    return { url: j.url, muteHttpExceptions: true };
  }));

  const items = [];
  s3Resps.forEach(function(resp, k) {
    if (resp.getResponseCode() !== 200) return;
    try {
      // GASがContent-Encoding:gzipを自動解凍するため、そのままテキストで読む
      const text = resp.getContentText('UTF-8');
      if (!text || text.trim().length === 0) return;
      const rec  = recordings[s3Jobs[k].recIndex];
      items.push({
        title:      rec.filename || rec.title || rec.name || '無題',
        createTime: rec.start_time || rec.create_time || null,
        summary:    text
      });
    } catch(e) { console.log('読み取りエラー: ' + e.message); }
  });

  return items;
}

function plaudGet_(token, url) {
  try {
    const resp = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type':  'application/json',
        'Accept':        'application/json'
      },
      muteHttpExceptions: true
    });

    const code = resp.getResponseCode();

    if (code === 401) {
      throw new Error('PLAUD_TOKEN が期限切れです。トークンを更新してください。\n取得方法: web.plaud.ai → DevTools → Network → Authorization ヘッダー');
    }
    if (code !== 200) {
      console.log('API エラー ' + code + ': ' + url);
      return null;
    }

    return JSON.parse(resp.getContentText());

  } catch (e) {
    if (e.message.includes('期限切れ')) {
      const webhookUrl = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');
      alertSlack_(webhookUrl, '⚠️ *PLAUDトークン期限切れ*\nweb.plaud.ai → DevTools → Network → Authorization ヘッダーから新しいトークンを取得し、スクリプトプロパティの PLAUD_TOKEN を更新してください。');
    }
    console.log('Fetch エラー: ' + e.message);
    return null;
  }
}

// ================================================================
// Gemini 再要約
// ================================================================

/**
 * 各録音をGeminiで500〜600字に個別要約。fetchAllで並列処理。
 * Gemini失敗時はその録音の元データをそのまま使用（フォールバック）。
 */
function summarizeItemsWithGemini_(items, apiKey) {
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' + apiKey;

  const reqs = items.map(function(item) {
    const cleanedText = cleanSummary_(item.summary);
    const prompt =
      '以下は商談・会議の録音要約です。重要ポイントを箇条書き（- で始める）で10個以内にまとめてください。\n' +
      '・各ポイントは1〜2文、簡潔に\n' +
      '・固有名詞・数値・期日は省略しない\n' +
      '・決定事項やアクションアイテムを優先\n' +
      '・前置きや後書きは不要。箇条書きのみ出力\n\n' +
      cleanedText;

    return {
      url: url,
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.2, maxOutputTokens: 800 }
      }),
      muteHttpExceptions: true
    };
  });

  console.log('Gemini 個別要約: ' + items.length + '件を並列処理');
  const resps = UrlFetchApp.fetchAll(reqs);

  return items.map(function(item, i) {
    try {
      const code = resps[i].getResponseCode();
      if (code !== 200) {
        console.log('Gemini エラー [' + i + ']: HTTP ' + code);
        return item;
      }
      const result = JSON.parse(resps[i].getContentText());
      const text = result.candidates &&
                   result.candidates[0] &&
                   result.candidates[0].content &&
                   result.candidates[0].content.parts &&
                   result.candidates[0].content.parts[0] &&
                   result.candidates[0].content.parts[0].text;
      if (text) return { title: item.title, createTime: item.createTime, summary: text.trim() };
    } catch(e) {
      console.log('Gemini 要約エラー [' + i + ']: ' + e.message);
    }
    return item; // 失敗時は元データをそのまま使用
  });
}

function generateDigest_(items, yesterdayStr, apiKey) {
  const dateLabel = yesterdayStr.replace(/-/g, '/');

  const body = items.map(function(item, i) {
    const timeStr = item.createTime ? toJstTimeStr_(item.createTime) : '';
    // 各要約は先頭400文字に制限（プロンプトが長くなりすぎないよう）
    const summaryTrunc = item.summary.substring(0, 400);
    return '【' + (i + 1) + '. ' + item.title + (timeStr ? ' (' + timeStr + ')' : '') + '】\n' + summaryTrunc;
  }).join('\n\n');

  const prompt =
    '以下は' + dateLabel + 'の録音要約（' + items.length + '件）です。\n' +
    '全体を通じて重要なポイント・決定事項・アクションアイテムを5項目以内の箇条書きにまとめてください。\n\n' +
    body + '\n\n' +
    '# 出力ルール\n' +
    '- 箇条書き（• で始める）\n' +
    '- 各項目1〜2文、できるだけ50文字以内\n' +
    '- 数字・固有名詞・期日を優先して含める\n' +
    '- 前置き・結び言葉は不要、箇条書きのみ出力';

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' + apiKey;

  try {
    const resp = UrlFetchApp.fetch(url, {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.3, maxOutputTokens: 800 }
      }),
      muteHttpExceptions: true
    });

    console.log('Gemini HTTP status: ' + resp.getResponseCode());
    const rawText = resp.getContentText();
    console.log('Gemini response: ' + rawText.substring(0, 500));

    const result = JSON.parse(rawText);
    const candidate = result && result.candidates && result.candidates[0];
    if (!candidate) {
      console.log('Gemini: candidates なし。error=' + JSON.stringify(result.error || ''));
      return null;
    }
    if (candidate.finishReason && candidate.finishReason !== 'STOP') {
      console.log('Gemini finishReason: ' + candidate.finishReason);
    }
    const text = candidate.content &&
                 candidate.content.parts &&
                 candidate.content.parts[0] &&
                 candidate.content.parts[0].text;
    return text ? text.trim() : null;

  } catch (e) {
    console.log('Gemini エラー: ' + e.message);
    return null;
  }
}

// ================================================================
// Slack 投稿
// ================================================================

function postToSlack_(webhookUrl, yesterdayStr, items, digest) {
  const dateLabel = yesterdayStr.replace('-', '年').replace('-', '月') + '日';
  const header    = '📋 ' + dateLabel + 'の録音ダイジェスト（' + items.length + '件）';

  const blocks = [
    {
      type: 'header',
      text: { type: 'plain_text', text: header, emoji: true }
    }
  ];

  // ダイジェスト（まとめ）
  if (digest) {
    blocks.push({
      type: 'section',
      text: { type: 'mrkdwn', text: '*✅ 全体まとめ*\n' + digest }
    });
    blocks.push({ type: 'divider' });
  }

  // 個別の録音
  for (const item of items) {
    const timeStr  = item.createTime ? toJstTimeStr_(item.createTime) : '';
    const titleLine = '*' + item.title + '*' + (timeStr ? '  _(' + timeStr + ')_' : '');
    const preview  = extractKeyPoints_(item.summary || '');

    blocks.push({
      type: 'section',
      text: { type: 'mrkdwn', text: titleLine + '\n' + preview }
    });
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

// ================================================================
// ユーティリティ
// ================================================================

function getYesterdayJstStr_() {
  // GASのFormatDateはタイムゾーンを正確に扱える
  const now = new Date();
  // 24時間前のDateを取得してJSTで日付文字列化
  const yesterday = new Date(now.getTime() - 24 * 60 * 60 * 1000);
  return Utilities.formatDate(yesterday, 'Asia/Tokyo', 'yyyy-MM-dd');
}

function toJstDateStr_(value) {
  if (!value) return null;
  let d;
  if (typeof value === 'number') {
    d = value > 1e11 ? new Date(value) : new Date(value * 1000); // 13桁以上→ms、10桁→秒
  } else {
    d = new Date(value);
  }
  if (isNaN(d.getTime())) return null;
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
}

function toJstTimeStr_(value) {
  if (!value) return '';
  let d;
  if (typeof value === 'number') {
    d = value > 1e11 ? new Date(value) : new Date(value * 1000);
  } else {
    d = new Date(value);
  }
  if (isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, 'Asia/Tokyo', 'HH:mm');
}

/**
 * テキストから箇条書きポイントを最大10個抽出。
 * Gemini失敗時のフォールバックとしても使用。
 * PLAUDの要約には - や数字リストが含まれるため、それを優先して抽出する。
 */
function extractKeyPoints_(rawText) {
  const cleaned = cleanSummary_(rawText);
  const lines = cleaned.split('\n').map(function(l) { return l.trim(); }).filter(function(l) { return l; });

  // 箇条書き行（- / • / ・ / 数字. で始まる）を抽出
  const bullets = lines.filter(function(l) {
    return /^[-•・]/.test(l) || /^\d+[\.．]/.test(l);
  });

  if (bullets.length >= 3) {
    // 十分な箇条書きがある場合は最大10個
    return bullets.slice(0, 10).join('\n');
  }

  // 箇条書きが少ない場合：意味のある行を最大10行
  const meaningful = lines.filter(function(l) { return l.length > 15; });
  return meaningful.slice(0, 10).join('\n');
}

function cleanSummary_(rawText) {
  const lines = rawText
    // 画像リンク除去
    .replace(/!\[[^\]]*\]\([^)]*\)/g, '')
    // テンプレートプレースホルダー（値）除去
    .replace(/\[[^\]]*(?:挿入|Insert|Speaker\s*\d*)[^\]]*\]/gi, '')
    // Markdownヘッダー記号除去（テキストは残す）
    .replace(/^#{1,4}\s*/gm, '')
    // 引用記号除去
    .replace(/^>\s*/gm, '')
    // 水平線除去
    .replace(/^[-*]{3,}$/gm, '')
    // 太字記号除去
    .replace(/\*\*/g, '')
    .split('\n')
    .map(function(l) { return l.trim(); })
    .filter(function(l) {
      if (!l) return false;
      // テンプレートフィールド行を除去（場所: / 顧客: / 日時: / 参加者: など）
      if (/^(?:日時|日付と時刻|場所|顧客|参加者|顧客名)\s*[:：]/.test(l)) return false;
      // セクションラベルのみの行を除去
      if (/^(?:概要|会議情報|会議メモ)\s*[:：]?\s*$/.test(l)) return false;
      return true;
    });

  const text = lines.join('\n');
  // Slackブロックの3000文字制限への安全対策のみ
  return text.length > 2500 ? text.substring(0, 2500) + '…' : text;
}

function toArray_(data) {
  if (!data) return [];
  if (Array.isArray(data))                        return data;
  if (Array.isArray(data.data_file_list))         return data.data_file_list; // PLAUD API
  if (Array.isArray(data.data))                   return data.data;
  if (Array.isArray(data.list))                   return data.list;
  if (Array.isArray(data.items))                  return data.items;
  if (Array.isArray(data.result))                 return data.result;
  if (data.data && Array.isArray(data.data.list)) return data.data.list;
  return [];
}

// ================================================================
// テスト・デバッグ用
// ================================================================

/**
 * PLAUDへの接続確認。ログに直近3件の録音情報が出ればOK。
 * フィールド名の確認にも使用（要約フィールド名が異なる場合に対応）。
 */
function testConnection() {
  const props   = PropertiesService.getScriptProperties();
  const token   = props.getProperty('PLAUD_TOKEN');
  const baseUrl = props.getProperty('PLAUD_BASE_URL') || DEFAULT_PLAUD_BASE_URL;

  if (!token) { Logger.log('❌ PLAUD_TOKEN が未設定です'); return; }

  const data = plaudGet_(token, baseUrl + '/file/simple/web?page=1&page_size=3');
  if (!data) { Logger.log('❌ 接続失敗。トークンまたはBASE_URLを確認してください'); return; }

  // 生のレスポンス構造を確認（デバッグ用）
  Logger.log('=== RAW レスポンス（構造確認） ===');
  Logger.log(JSON.stringify(data, null, 2).substring(0, 2000));

  const files = toArray_(data);
  Logger.log('✅ 接続成功 - ' + files.length + '件取得');

  if (files.length > 0) {
    Logger.log('--- 最新録音のフィールド確認 ---');
    Logger.log(JSON.stringify(files[0], null, 2));

    // 最初の1件の詳細も取得してsummaryフィールドを確認
    const detail = plaudGet_(token, baseUrl + '/file/detail/' + files[0].id);
    Logger.log('--- file/detail レスポンス ---');
    Logger.log(JSON.stringify(detail, null, 2));
  }
}

/**
 * 指定した日付でSlackに実際に投稿するテスト用関数。
 * testDate を変更して実行する。
 */
function testWithDate() {
  const testDate = '2026-06-26'; // ← テストしたい日付に変更

  const props      = PropertiesService.getScriptProperties();
  const token      = props.getProperty('PLAUD_TOKEN');
  const webhookUrl = props.getProperty('SLACK_WEBHOOK_URL');
  const geminiKey  = props.getProperty('GEMINI_API_KEY');
  const baseUrl    = props.getProperty('PLAUD_BASE_URL') || DEFAULT_PLAUD_BASE_URL;

  if (!token)      { Logger.log('❌ PLAUD_TOKEN が未設定です'); return; }
  if (!webhookUrl) { Logger.log('❌ SLACK_WEBHOOK_URL が未設定です'); return; }

  Logger.log('対象日(テスト): ' + testDate);

  const recordings = fetchRecordingsByDate_(token, baseUrl, testDate);
  Logger.log('録音件数: ' + recordings.length + '件');
  if (recordings.length === 0) { Logger.log('録音なし'); return; }

  const items = fetchAllSummaryItems_(token, baseUrl, recordings);
  Logger.log('要約取得: ' + items.length + '件');
  if (items.length === 0) { Logger.log('要約なし'); return; }

  const finalItems = geminiKey ? summarizeItemsWithGemini_(items, geminiKey) : items;
  postToSlack_(webhookUrl, testDate, finalItems, null);
  Logger.log('Slack投稿完了');
}

/**
 * Slack投稿なしでダイジェストの内容だけをログに出力。
 * 本番実行前の内容確認に使用。
 */
function testDryRun() {
  const props   = PropertiesService.getScriptProperties();
  const token   = props.getProperty('PLAUD_TOKEN');
  const geminiKey = props.getProperty('GEMINI_API_KEY');
  const baseUrl = props.getProperty('PLAUD_BASE_URL') || DEFAULT_PLAUD_BASE_URL;

  if (!token) { Logger.log('❌ PLAUD_TOKEN が未設定です'); return; }

  const yesterdayStr = getYesterdayJstStr_();
  Logger.log('対象日(JST): ' + yesterdayStr);

  const recordings = fetchRecordingsByDate_(token, baseUrl, yesterdayStr);
  Logger.log('録音件数: ' + recordings.length + '件');

  // 並列取得
  const items = fetchAllSummaryItems_(token, baseUrl, recordings);
  items.forEach(function(item, i) {
    Logger.log('---');
    Logger.log('タイトル: ' + item.title);
    Logger.log('要約冒頭: ' + item.summary.substring(0, 100) + '...');
  });

  if (items.length > 0 && geminiKey) {
    const digest = generateDigest_(items, yesterdayStr, geminiKey);
    Logger.log('=== Geminiダイジェスト ===');
    Logger.log(digest);
  }
}
