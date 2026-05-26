/**
 * ════════════════════════════════════════════════════════════
 *  AI財務分析・企業価値算定ツール
 *  Google Apps Script + Claude API (Anthropic)
 *
 *  【セットアップ手順】
 *  1. Google スプレッドシートを新規作成
 *  2. メニュー「拡張機能」→「Apps Script」を開く
 *  3. このファイルの内容をすべてコピーして貼り付け（既存コードを削除）
 *  4. 保存（Ctrl+S）してスプレッドシートに戻る
 *  5. メニュー「AI財務分析」→「① 初期設定（シート作成）」を実行
 *  6. 「② APIキーを設定」からAnthropicのAPIキーを入力
 *  7. 設定シートに会社情報・業種を入力
 *  8. BS・PL・販管費・製造原価に決算数値を入力
 *  9. 「▶ AI分析を実行」でレポートを自動生成
 * ════════════════════════════════════════════════════════════
 */

'use strict';

// ─── Claude APIの設定 ─────────────────────────────────────
const CLAUDE_MODEL   = 'claude-sonnet-4-6';
const CLAUDE_API_URL = 'https://api.anthropic.com/v1/messages';
const CLAUDE_VERSION = '2023-06-01';
const MAX_TOKENS     = 8000;

// ─── シート名 ─────────────────────────────────────────────
const SH = {
  SETTINGS  : '設定',
  BS        : '貸借対照表',
  PL        : '損益計算書',
  SGA       : '販売費及び一般管理費',
  COGS      : '製造原価明細',
  INDUSTRY  : '業界比較・アラート',
  VAL_BEFORE: '企業価値（修正前）',
  VAL_AFTER : '企業価値（修正後）',
};

// ─── カラーパレット ───────────────────────────────────────
const C = {
  NAVY:'#1F3864', BLUE:'#2E75B6', LBLUE:'#BDD7EE', LLBLUE:'#DEEAF1',
  WHITE:'#FFFFFF', LGRAY:'#F8F9FA', GRAY:'#F2F2F2', DGRAY:'#595959',
  GOLD:'#FFF2CC', DGOLD:'#7D4706',
  RED:'#FFC7CE',  DRED:'#9C0006',
  YEL:'#FFEB9C',  DYEL:'#7D4706',
  GRN:'#C6EFCE',  DGRN:'#375623',
  INPUT:'#FFFDE7',
};

// ─── 東証33業種データ ──────────────────────────────────────
// 財務省法人企業統計・中小企業庁データ等をベースにした参考値
// gm=売上総利益率, om=営業利益率, nm=経常利益率, np=純利益率
// eq=自己資本比率, cr=流動比率, fr=固定比率, lc=人件費率
// roa=ROA, roe=ROE, evE=EV/EBITDA, evS=EV/売上高, pbr=PBR, per=PER
const TSE33 = {
  '0050':{ name:'水産・農林業',    gm:.25, om:.03, nm:.04, np:.02, eq:.40, cr:1.20, fr:.80, lc:.20, roa:.03, roe:.07, evE: 6.0, evS: .6, pbr:1.0, per:12 },
  '1050':{ name:'鉱業',            gm:.30, om:.05, nm:.06, np:.03, eq:.45, cr:1.30, fr:1.2, lc:.15, roa:.04, roe:.08, evE: 7.0, evS:1.0, pbr:1.2, per:12 },
  '2050':{ name:'建設業',          gm:.22, om:.03, nm:.04, np:.02, eq:.38, cr:1.30, fr:.60, lc:.20, roa:.03, roe:.08, evE: 5.0, evS: .4, pbr:1.0, per:10 },
  '3050':{ name:'食料品',          gm:.35, om:.04, nm:.05, np:.03, eq:.42, cr:1.25, fr:.75, lc:.18, roa:.04, roe:.08, evE: 8.0, evS: .8, pbr:1.3, per:14 },
  '3100':{ name:'繊維製品',        gm:.28, om:.03, nm:.04, np:.02, eq:.40, cr:1.20, fr:.80, lc:.22, roa:.03, roe:.06, evE: 6.0, evS: .5, pbr: .9, per:10 },
  '3150':{ name:'パルプ・紙',      gm:.20, om:.03, nm:.04, np:.02, eq:.40, cr:1.20, fr:1.1, lc:.15, roa:.03, roe:.06, evE: 6.0, evS: .4, pbr: .9, per:10 },
  '3200':{ name:'化学',            gm:.30, om:.06, nm:.07, np:.04, eq:.48, cr:1.40, fr:.85, lc:.18, roa:.05, roe:.09, evE: 8.0, evS:1.0, pbr:1.4, per:14 },
  '3250':{ name:'医薬品',          gm:.65, om:.15, nm:.16, np:.11, eq:.60, cr:2.00, fr:.50, lc:.25, roa:.10, roe:.14, evE:15.0, evS:3.0, pbr:2.5, per:25 },
  '3300':{ name:'石油・石炭製品',  gm:.08, om:.02, nm:.03, np:.01, eq:.30, cr:1.10, fr:1.2, lc:.05, roa:.02, roe:.06, evE: 5.0, evS: .2, pbr: .8, per:10 },
  '3350':{ name:'ゴム製品',        gm:.28, om:.05, nm:.06, np:.03, eq:.45, cr:1.30, fr:.90, lc:.20, roa:.04, roe:.08, evE: 7.0, evS: .7, pbr:1.1, per:12 },
  '3400':{ name:'ガラス・土石製品',gm:.26, om:.04, nm:.05, np:.02, eq:.42, cr:1.25, fr:1.0, lc:.22, roa:.03, roe:.07, evE: 6.0, evS: .6, pbr:1.0, per:11 },
  '3450':{ name:'鉄鋼',            gm:.15, om:.03, nm:.04, np:.02, eq:.35, cr:1.15, fr:1.2, lc:.12, roa:.03, roe:.07, evE: 5.0, evS: .3, pbr: .9, per:10 },
  '3500':{ name:'非鉄金属',        gm:.18, om:.04, nm:.05, np:.02, eq:.38, cr:1.20, fr:1.0, lc:.14, roa:.03, roe:.07, evE: 6.0, evS: .4, pbr:1.0, per:11 },
  '3550':{ name:'金属製品',        gm:.24, om:.04, nm:.05, np:.02, eq:.42, cr:1.25, fr:.85, lc:.22, roa:.03, roe:.07, evE: 6.0, evS: .5, pbr:1.0, per:11 },
  '3600':{ name:'機械',            gm:.28, om:.06, nm:.07, np:.04, eq:.48, cr:1.40, fr:.80, lc:.20, roa:.05, roe:.09, evE: 8.0, evS: .9, pbr:1.3, per:14 },
  '3650':{ name:'電気機器',        gm:.32, om:.07, nm:.08, np:.05, eq:.50, cr:1.45, fr:.75, lc:.22, roa:.06, roe:.10, evE:10.0, evS:1.2, pbr:1.8, per:18 },
  '3700':{ name:'輸送用機器',      gm:.20, om:.05, nm:.06, np:.03, eq:.42, cr:1.25, fr:.95, lc:.18, roa:.04, roe:.08, evE: 7.0, evS: .6, pbr:1.1, per:12 },
  '3750':{ name:'精密機器',        gm:.45, om:.10, nm:.11, np:.07, eq:.55, cr:1.60, fr:.65, lc:.28, roa:.07, roe:.12, evE:12.0, evS:1.8, pbr:2.0, per:20 },
  '3800':{ name:'その他製品',      gm:.35, om:.05, nm:.06, np:.03, eq:.45, cr:1.30, fr:.80, lc:.22, roa:.04, roe:.08, evE: 8.0, evS: .8, pbr:1.2, per:13 },
  '4050':{ name:'電気・ガス業',    gm:.25, om:.06, nm:.06, np:.03, eq:.25, cr:.90,  fr:2.5, lc:.08, roa:.02, roe:.07, evE: 8.0, evS: .8, pbr:1.0, per:12 },
  '5050':{ name:'陸運業',          gm:.30, om:.04, nm:.05, np:.02, eq:.28, cr:.95,  fr:1.8, lc:.35, roa:.03, roe:.07, evE: 7.0, evS: .6, pbr: .9, per:12 },
  '5100':{ name:'海運業',          gm:.20, om:.08, nm:.09, np:.05, eq:.35, cr:1.10, fr:1.5, lc:.15, roa:.04, roe:.10, evE: 5.0, evS: .5, pbr: .9, per: 8 },
  '5150':{ name:'空運業',          gm:.25, om:.04, nm:.04, np:.02, eq:.20, cr:.85,  fr:2.0, lc:.30, roa:.02, roe:.08, evE: 7.0, evS: .5, pbr:1.0, per:12 },
  '5200':{ name:'倉庫・運輸関連業',gm:.28, om:.05, nm:.06, np:.03, eq:.40, cr:1.20, fr:1.2, lc:.28, roa:.03, roe:.07, evE: 7.0, evS: .7, pbr:1.0, per:12 },
  '5250':{ name:'情報・通信業',    gm:.55, om:.10, nm:.11, np:.07, eq:.50, cr:1.60, fr:.45, lc:.40, roa:.07, roe:.13, evE:14.0, evS:2.5, pbr:2.8, per:23 },
  '6050':{ name:'卸売業',          gm:.18, om:.02, nm:.03, np:.02, eq:.32, cr:1.20, fr:.60, lc:.10, roa:.02, roe:.06, evE: 6.0, evS: .3, pbr:1.0, per:10 },
  '6100':{ name:'小売業',          gm:.32, om:.03, nm:.04, np:.02, eq:.30, cr:1.10, fr:.90, lc:.18, roa:.03, roe:.08, evE: 7.0, evS: .5, pbr:1.1, per:13 },
  '7050':{ name:'銀行業',          gm:null,om:.15, nm:.15, np:.10, eq:.04, cr:null, fr:null,lc:.35, roa:.003,roe:.07, evE:null, evS:null,pbr: .5, per:10 },
  '7100':{ name:'証券・商品先物',  gm:null,om:.15, nm:.15, np:.10, eq:.12, cr:null, fr:null,lc:.40, roa:.02, roe:.10, evE:null, evS:null,pbr: .8, per:12 },
  '7150':{ name:'保険業',          gm:null,om:.08, nm:.08, np:.05, eq:.08, cr:null, fr:null,lc:.20, roa:.01, roe:.08, evE:null, evS:null,pbr: .8, per:12 },
  '7200':{ name:'その他金融業',    gm:null,om:.20, nm:.20, np:.12, eq:.10, cr:null, fr:null,lc:.25, roa:.03, roe:.12, evE:null, evS:null,pbr:1.0, per:12 },
  '8050':{ name:'不動産業',        gm:.40, om:.12, nm:.12, np:.07, eq:.30, cr:1.10, fr:1.8, lc:.08, roa:.04, roe:.10, evE:12.0, evS:2.0, pbr:1.5, per:15 },
  '9050':{ name:'サービス業',      gm:.55, om:.06, nm:.07, np:.04, eq:.38, cr:1.35, fr:.65, lc:.42, roa:.05, roe:.10, evE: 9.0, evS:1.5, pbr:1.6, per:16 },
};

// ════════════════════════════════════════════════════════════
// メニュー
// ════════════════════════════════════════════════════════════
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('AI財務分析')
    .addItem('① 初期設定（シート作成）', 'initializeSpreadsheet')
    .addSeparator()
    .addItem('② APIキーを設定', 'showApiKeyDialog')
    .addSeparator()
    .addItem('▶ AI分析を実行', 'runAnalysis')
    .addToUi();
}

// ════════════════════════════════════════════════════════════
// 初期設定：シート作成
// ════════════════════════════════════════════════════════════
function initializeSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  createSettingsSheet(ss);
  createInputSheet(ss, SH.BS,   '貸借対照表');
  createInputSheet(ss, SH.PL,   '損益計算書');
  createInputSheet(ss, SH.SGA,  '販売費及び一般管理費');
  createInputSheet(ss, SH.COGS, '製造原価明細（製造業のみ・任意）');

  // シート順を整理
  const order = [SH.SETTINGS, SH.BS, SH.PL, SH.SGA, SH.COGS];
  order.forEach((name, i) => {
    const ws = ss.getSheetByName(name);
    if (ws) ss.setActiveSheet(ws), ss.moveActiveSheet(i + 1);
  });
  ss.setActiveSheet(ss.getSheetByName(SH.SETTINGS));

  ui.alert(
    '初期設定完了',
    '入力シートを作成しました。\n\n次のステップ:\n' +
    '① メニュー「APIキーを設定」でAnthropicのAPIキーを入力\n' +
    '② 設定シートに会社情報・業種コードを入力\n' +
    '③ 貸借対照表・損益計算書・販管費に決算数値を入力\n' +
    '④ 製造業の場合は製造原価明細も入力\n' +
    '⑤「AI分析を実行」をクリック',
    ui.ButtonSet.OK
  );
}

function createSettingsSheet(ss) {
  let ws = ss.getSheetByName(SH.SETTINGS);
  if (!ws) ws = ss.insertSheet(SH.SETTINGS);
  ws.clear();

  ws.setColumnWidth(1, 30); ws.setColumnWidth(2, 220); ws.setColumnWidth(3, 320);

  // タイトル
  titleRow(ws, 1, 'AI財務分析ツール　設定シート', 3, C.NAVY, C.WHITE, 30);

  const rows = [
    ['会社名',              '',    '分析対象の会社名を入力してください'],
    ['業種コード（東証33）','',    '右の業種コード一覧から4桁のコードを入力（例: 5250）'],
    ['製造原価明細',        'なし','「あり」または「なし」を入力（製造業以外はなし）'],
    ['分析対象期',          '最新期','「最新期」と入力（または「3」など期番号）'],
    ['単位',               '千円', '「千円」「百万円」「円」のいずれかを入力'],
    ['',                   '',    ''],
    ['年倍法　係数（低）',  '2',   '保守的な係数（デフォルト2年）'],
    ['年倍法　係数（高）',  '4',   '楽観的な係数（デフォルト4年）'],
  ];

  rows.forEach((r, i) => {
    const row = i + 2;
    const labelCell = ws.getRange(row, 1);
    const inputCell = ws.getRange(row, 2);
    const noteCell  = ws.getRange(row, 3);

    labelCell.setValue(r[0]).setBackground(C.LLBLUE).setFontWeight('bold')
      .setVerticalAlignment('middle').setBorder(true,true,true,true,false,false,'#BFBFBF',SpreadsheetApp.BorderStyle.SOLID);

    if (r[0] && !r[0].startsWith('年')) {
      inputCell.setValue(r[1]).setBackground(C.INPUT).setFontColor('#1F497D')
        .setFontWeight('bold').setHorizontalAlignment('left');
    } else {
      inputCell.setValue(r[1]).setBackground(r[0] ? C.INPUT : C.GRAY)
        .setFontColor('#1F497D').setFontWeight('bold');
    }
    inputCell.setVerticalAlignment('middle')
      .setBorder(true,true,true,true,false,false,'#BFBFBF',SpreadsheetApp.BorderStyle.SOLID);

    noteCell.setValue(r[2]).setFontColor(C.DGRAY).setFontSize(9)
      .setVerticalAlignment('middle')
      .setBorder(true,true,true,true,false,false,'#BFBFBF',SpreadsheetApp.BorderStyle.SOLID);
    ws.setRowHeight(row, 22);
  });

  // 業種コード一覧
  const listStart = rows.length + 3;
  titleRow(ws, listStart, '東証33業種コード一覧（設定シートのB3に4桁コードを入力）', 3, C.BLUE, C.WHITE, 22);

  Object.entries(TSE33).forEach(([code, d], i) => {
    const r = listStart + 1 + i;
    ws.getRange(r, 1).setValue(code).setHorizontalAlignment('center');
    ws.getRange(r, 2).setValue(d.name);
    ws.getRange(r, 3).setValue(`売総利益率${pct(d.gm)} / 営業利益率${pct(d.om)} / 経常利益率${pct(d.nm)} / 自己資本比率${pct(d.eq)}`).setFontSize(8).setFontColor(C.DGRAY);
    const bg = i % 2 === 0 ? C.WHITE : C.LGRAY;
    ws.getRange(r, 1, 1, 3).setBackground(bg)
      .setBorder(true,true,true,true,false,false,'#BFBFBF',SpreadsheetApp.BorderStyle.SOLID);
    ws.setRowHeight(r, 18);
  });

  // 注記
  const noteRow = listStart + Object.keys(TSE33).length + 2;
  ws.getRange(noteRow, 1, 1, 3).merge()
    .setValue('【注意】APIキーはメニュー「APIキーを設定」から入力してください（セキュリティのためこのシートには表示されません）。\n業界平均値は参考値です。実際のM&A・事業承継における価格はデューデリジェンス結果や当事者間交渉により決定されます。')
    .setBackground(C.LGRAY).setFontColor(C.DGRAY).setFontSize(8).setWrap(true).setVerticalAlignment('top');
  ws.setRowHeight(noteRow, 50);
}

function createInputSheet(ss, shName, label) {
  if (ss.getSheetByName(shName)) return;
  const ws = ss.insertSheet(shName);

  ws.setColumnWidth(1, 230);
  for (let c = 2; c <= 10; c++) ws.setColumnWidth(c, 130);

  titleRow(ws, 1, label + '　（行・列の追加・削除・科目名変更は自由です）', 10, C.NAVY, C.WHITE, 28);

  ws.getRange(2, 1, 1, 10).merge()
    .setValue('【入力方法】ヘッダー行（3行目）に期の名前を記入し、4行目以降に科目名と数値を入力してください。期数は何期でも追加できます。製造原価明細は製造業のみ入力（他業種は空白でもOK）。')
    .setBackground(C.GOLD).setFontColor(C.DGOLD).setFontSize(9).setWrap(true).setVerticalAlignment('middle');
  ws.setRowHeight(2, 40);

  // ヘッダー行テンプレート
  const headers = ['科目名', '第1期\n（20XX年〇月期）', '第2期\n（20XX年〇月期）', '第3期\n（20XX年〇月期）'];
  headers.forEach((h, i) => {
    ws.getRange(3, i + 1).setValue(h)
      .setBackground(C.NAVY).setFontColor(C.WHITE).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  });
  ws.getRange(3, 5, 1, 6).setBackground(C.BLUE).setFontColor(C.WHITE)
    .setValue('← 期を追加する場合はここに列を挿入').setHorizontalAlignment('center');
  ws.setRowHeight(3, 42);

  ws.getRange(4, 1).setValue('ここから科目と数値を入力 →')
    .setFontColor(C.DGRAY).setFontStyle('italic').setFontSize(9);
  ws.setFrozenRows(3);
}

// ════════════════════════════════════════════════════════════
// APIキー管理
// ════════════════════════════════════════════════════════════
function showApiKeyDialog() {
  const ui  = SpreadsheetApp.getUi();
  const res = ui.prompt(
    'Anthropic APIキーを設定',
    'APIキーを入力してください（sk-ant-api03-... で始まる文字列）\n' +
    '取得先: https://console.anthropic.com/\n' +
    '※スクリプトプロパティに安全に保存されます。シートには表示されません。',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;

  const key = res.getResponseText().trim();
  if (!key.startsWith('sk-ant-')) {
    ui.alert('エラー', 'APIキーが正しくありません。「sk-ant-」で始まる文字列を入力してください。', ui.ButtonSet.OK);
    return;
  }
  PropertiesService.getScriptProperties().setProperty('ANTHROPIC_API_KEY', key);
  ui.alert('設定完了', 'APIキーを保存しました。', ui.ButtonSet.OK);
}

function getApiKey_() {
  const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!key) throw new Error('APIキーが未設定です。メニュー「② APIキーを設定」から入力してください。');
  return key;
}

// ════════════════════════════════════════════════════════════
// メイン分析処理
// ════════════════════════════════════════════════════════════
function runAnalysis() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    const apiKey   = getApiKey_();
    const settings = readSettings_(ss);
    validateSettings_(settings, ui);

    toast_('財務データを読み込み中...', ss);
    const sheetData = readAllSheets_(ss, settings);

    toast_('Claude AIが分析中です（30秒〜1分かかります）...', ss);
    const result = callClaude_(sheetData, settings, apiKey);

    toast_('分析結果をシートに書き込み中...', ss);
    writeIndustrySheet_(ss, result, settings);
    writeValuationSheet_(ss, result, settings, false);
    writeValuationSheet_(ss, result, settings, true);

    // 出力シートを先頭付近に移動
    [SH.INDUSTRY, SH.VAL_BEFORE, SH.VAL_AFTER].forEach((name, i) => {
      const ws = ss.getSheetByName(name);
      if (ws) ss.setActiveSheet(ws), ss.moveActiveSheet(6 + i);
    });

    ss.setActiveSheet(ss.getSheetByName(SH.INDUSTRY));
    toast_('分析完了！', ss);
    ui.alert('分析完了', '3枚のレポートシートを生成しました。\n・業界比較・アラート\n・企業価値（修正前）\n・企業価値（修正後）', ui.ButtonSet.OK);

  } catch (e) {
    toast_('エラーが発生しました', ss);
    ui.alert('エラー', e.message, ui.ButtonSet.OK);
    console.error(e.stack || e.message);
  }
}

// ════════════════════════════════════════════════════════════
// 設定・データ読み込み
// ════════════════════════════════════════════════════════════
function readSettings_(ss) {
  const ws = ss.getSheetByName(SH.SETTINGS);
  if (!ws) throw new Error('設定シートが見つかりません。「① 初期設定」を実行してください。');

  const v = ws.getRange(2, 2, 8, 1).getValues().map(r => String(r[0]).trim());
  return {
    company        : v[0],
    industryCode   : v[1],
    hasCogs        : v[2].includes('あり'),
    targetPeriod   : v[3] || '最新期',
    unit           : v[4] || '千円',
    yearLow        : parseFloat(v[6]) || 2,
    yearHigh       : parseFloat(v[7]) || 4,
    industryData   : TSE33[v[1]],
    industryName   : TSE33[v[1]] ? TSE33[v[1]].name : '',
  };
}

function validateSettings_(s, ui) {
  if (!s.company)        throw new Error('設定シートの「会社名」を入力してください。');
  if (!s.industryData)   throw new Error(`業種コード「${s.industryCode}」が見つかりません。東証33業種コード一覧を確認してください。`);
}

function readAllSheets_(ss, settings) {
  const result = {};
  const targets = [SH.BS, SH.PL, SH.SGA];
  if (settings.hasCogs) targets.push(SH.COGS);

  targets.forEach(name => {
    const ws = ss.getSheetByName(name);
    if (!ws || ws.getLastRow() < 3) { result[name] = '（データなし）'; return; }

    const maxRow = Math.min(ws.getLastRow(), 200);
    const maxCol = Math.min(ws.getLastColumn(), 20);
    const vals   = ws.getRange(1, 1, maxRow, maxCol).getValues();

    const lines = vals
      .filter(row => row.some(v => v !== '' && v !== null))
      .map(row => row.map(v => {
        if (v === '' || v === null) return '';
        if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy/MM/dd');
        if (typeof v === 'number') return v.toLocaleString();
        return String(v);
      }).join('\t'));

    result[name] = lines.join('\n');
  });

  return result;
}

// ════════════════════════════════════════════════════════════
// Claude API 呼び出し
// ════════════════════════════════════════════════════════════
function callClaude_(sheetData, settings, apiKey) {
  const prompt = buildPrompt_(sheetData, settings);

  const payload = JSON.stringify({
    model      : CLAUDE_MODEL,
    max_tokens : MAX_TOKENS,
    system     : buildSystem_(),
    messages   : [{ role: 'user', content: prompt }],
  });

  const res = UrlFetchApp.fetch(CLAUDE_API_URL, {
    method          : 'post',
    contentType     : 'application/json',
    headers         : { 'x-api-key': apiKey, 'anthropic-version': CLAUDE_VERSION },
    payload         : payload,
    muteHttpExceptions: true,
  });

  const code = res.getResponseCode();
  const body = JSON.parse(res.getContentText());

  if (code !== 200) {
    const msg = (body.error || {}).message || JSON.stringify(body);
    throw new Error(`Claude API エラー (HTTP ${code}): ${msg}`);
  }

  const text = body.content[0].text;
  // JSONブロックを抽出
  const m = text.match(/```json\s*([\s\S]*?)\s*```/) || text.match(/(\{[\s\S]*\})/);
  if (!m) throw new Error('AIの応答からJSONを取得できませんでした。応答内容:\n' + text.substring(0, 500));

  try {
    return JSON.parse(m[1] || m[0]);
  } catch(e) {
    throw new Error('JSONのパースに失敗しました: ' + e.message + '\n内容: ' + (m[1] || m[0]).substring(0, 300));
  }
}

function buildSystem_() {
  return `あなたはM&A・事業承継専門の財務アドバイザーです。
提供された財務諸表データを分析し、指定されたJSON形式で結果のみを返してください。
- 科目名は異なる表記でも意味を解釈して正しく識別してください（例:「受取利息配当金」→営業外収益）
- 数値は元の単位のまま使用してください
- 分析コメントはすべて日本語で記述してください
- レスポンスはJSONコードブロック（\`\`\`json ... \`\`\`）のみとし、前後の説明文は不要です`;
}

function buildPrompt_(sheetData, s) {
  const ind = s.industryData;
  const indBenchmark = JSON.stringify({
    売上総利益率: ind.gm, 営業利益率: ind.om, 経常利益率: ind.nm, 純利益率: ind.np,
    自己資本比率: ind.eq, 流動比率: ind.cr, 固定比率: ind.fr, 人件費率: ind.lc,
    ROA: ind.roa, ROE: ind.roe,
    'EV/EBITDA倍率': ind.evE, 'EV/売上高倍率': ind.evS, PBR: ind.pbr, PER: ind.per,
  }, null, 2);

  return `## 財務分析依頼

### 会社情報
- 会社名: ${s.company}
- 業種: ${s.industryName}（東証33業種コード: ${s.industryCode}）
- 製造原価明細: ${s.hasCogs ? 'あり（製造業）' : 'なし'}
- 単位: ${s.unit}
- 年倍法係数: ${s.yearLow}年（低）〜${s.yearHigh}年（高）

---

### 財務データ

#### 貸借対照表
\`\`\`
${sheetData[SH.BS]}
\`\`\`

#### 損益計算書
\`\`\`
${sheetData[SH.PL]}
\`\`\`

#### 販売費及び一般管理費
\`\`\`
${sheetData[SH.SGA]}
\`\`\`
${s.hasCogs ? `
#### 製造原価明細
\`\`\`
${sheetData[SH.COGS]}
\`\`\`
` : ''}

---

### 業界平均財務指標（${s.industryName}・東証上場会社中央値ベース）
\`\`\`json
${indBenchmark}
\`\`\`

---

## 出力指示

以下のJSON形式で分析してください。**期数は入力データに合わせて動的に対応**してください（2期でも5期でも可）。
数値がない場合はnull、計算不能な場合は文字列"N/A"を使用してください。

\`\`\`json
{
  "meta": {
    "company": "会社名",
    "industry": "業種名",
    "industry_code": "コード",
    "unit": "単位",
    "periods": ["第1期（20XX年X月期）", "第2期（20XX年X月期）", ...],
    "n_periods": 期数(整数),
    "valuation_period_index": 企業価値算定に使用した期のインデックス(0始まり),
    "valuation_period": "算定対象期の名称"
  },
  "financials": {
    "revenue":               [期1, 期2, ...],
    "gross_profit":          [期1, 期2, ...],
    "sga_total":             [期1, 期2, ...],
    "operating_profit":      [期1, 期2, ...],
    "non_op_income":         [期1, 期2, ...],
    "non_op_expense":        [期1, 期2, ...],
    "ordinary_profit":       [期1, 期2, ...],
    "net_profit":            [期1, 期2, ...],
    "depreciation_sga":      [期1, 期2, ...],
    "depreciation_cogs":     [期1, 期2, ...],
    "total_assets":          [期1, 期2, ...],
    "current_assets":        [期1, 期2, ...],
    "fixed_assets":          [期1, 期2, ...],
    "total_liabilities":     [期1, 期2, ...],
    "current_liabilities":   [期1, 期2, ...],
    "fixed_liabilities":     [期1, 期2, ...],
    "net_assets":            [期1, 期2, ...],
    "labor_cost":            [期1, 期2, ...],
    "interest_bearing_debt": [期1, 期2, ...]
  },
  "ebitda": [期1, 期2, ...],
  "ratios": {
    "gross_margin":      [期1, 期2, ...],
    "operating_margin":  [期1, 期2, ...],
    "ordinary_margin":   [期1, 期2, ...],
    "net_margin":        [期1, 期2, ...],
    "equity_ratio":      [期1, 期2, ...],
    "current_ratio":     [期1, 期2, ...],
    "fixed_ratio":       [期1, 期2, ...],
    "labor_cost_ratio":  [期1, 期2, ...],
    "roa":               [期1, 期2, ...],
    "roe":               [期1, 期2, ...],
    "revenue_growth":    [null, 期2比期1成長率, ...]
  },
  "industry_comparison": [
    {
      "metric": "売上総利益率",
      "values": [期1, 期2, ...],
      "industry_avg": 業界平均値,
      "latest_diff":  最新期-業界平均,
      "alert": "red|yellow|green|na",
      "comment": "具体的な分析コメント（問題点・改善方向を含む）",
      "high_is_good": true
    },
    { "metric": "営業利益率",   ... },
    { "metric": "経常利益率",   ... },
    { "metric": "純利益率",     ... },
    { "metric": "人件費率",     "high_is_good": false, ... },
    { "metric": "自己資本比率", ... },
    { "metric": "流動比率",     ... },
    { "metric": "固定比率",     "high_is_good": false, ... },
    { "metric": "ROA",          ... },
    { "metric": "ROE",          ... },
    { "metric": "売上高成長率", ... }
  ],
  "valuation_before": {
    "ordinary_profit":      算定基準の経常利益,
    "net_assets":           算定基準の純資産,
    "ebitda":               算定基準のEBITDA,
    "revenue":              算定基準の売上高,
    "net_profit":           算定基準の当期純利益,
    "interest_bearing_debt":算定基準の有利子負債,
    "methods": {
      "year_low":   { "ev": 値, "equity": 値, "formula": "計算式の説明", "applicable": true },
      "year_high":  { "ev": 値, "equity": 値, "formula": "計算式の説明", "applicable": true },
      "ev_ebitda":  { "ev": 値, "equity": 値, "formula": "計算式の説明", "applicable": true/false },
      "ev_sales":   { "ev": 値, "equity": 値, "formula": "計算式の説明", "applicable": true },
      "pbr":        { "ev": 値, "equity": 値, "formula": "計算式の説明", "applicable": true },
      "per":        { "ev": 値, "equity": 値, "formula": "計算式の説明", "applicable": true/false },
      "net_assets": { "ev": 値, "equity": 値, "formula": "計算式の説明", "applicable": true }
    },
    "range_min": 株主価値の最小値,
    "range_max": 株主価値の最大値,
    "assessment": "修正前評価の要約コメント（200字程度）"
  },
  "adjustments": [
    {
      "no": 1,
      "item":   "調整項目名（例: 役員報酬の正常化）",
      "amount": 調整額（利益増加/資産増加はプラス、減少はマイナス）,
      "type":   "profit（利益調整）またはasset（純資産調整）",
      "reason": "財務データに基づく具体的な推奨理由"
    }
  ],
  "valuation_after": {
    "adj_ordinary_profit": 修正後経常利益,
    "adj_net_assets":      修正後純資産,
    "adj_ebitda":          修正後EBITDA,
    "methods": { ... 同じ構造 ... },
    "range_min": 修正後株主価値の最小値,
    "range_max": 修正後株主価値の最大値,
    "assessment": "修正後評価の要約コメント（200字程度）"
  },
  "overall_assessment": "総合評価・財務上の特記事項・事業承継における留意点（400字程度）"
}
\`\`\``;
}

// ════════════════════════════════════════════════════════════
// 業界比較シート 書き込み
// ════════════════════════════════════════════════════════════
function writeIndustrySheet_(ss, result, settings) {
  let ws = ss.getSheetByName(SH.INDUSTRY);
  if (ws) ss.deleteSheet(ws);
  ws = ss.insertSheet(SH.INDUSTRY);

  const meta    = result.meta || {};
  const periods = meta.periods || [];
  const n       = periods.length;
  const ind     = settings.industryData;

  // 列レイアウト
  const cLabel  = 1;
  const cDesc   = 2;
  const cPStart = 3;
  const cAvg    = cPStart + n;
  const cDiff   = cAvg + 1;
  const cJudge  = cDiff + 1;
  const cNote   = cJudge + 1;
  const totalCols = cNote;

  ws.setColumnWidth(cLabel, 190);
  ws.setColumnWidth(cDesc,  160);
  for (let i = 0; i < n; i++) ws.setColumnWidth(cPStart + i, 110);
  ws.setColumnWidth(cAvg,   110);
  ws.setColumnWidth(cDiff,  100);
  ws.setColumnWidth(cJudge,  65);
  ws.setColumnWidth(cNote,  300);

  let r = 1;

  // タイトル
  titleRow(ws, r, `業界比較・アラート　${settings.company}　業種: ${settings.industryName}（${settings.industryCode}）`, totalCols);
  r++;
  ws.getRange(r, 1, 1, totalCols).merge()
    .setValue(`生成: ${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')}　　単位: ${settings.unit}　　注: 業界平均は東証上場会社中央値ベースの参考値`)
    .setBackground(C.LGRAY).setFontColor(C.DGRAY).setFontSize(8).setHorizontalAlignment('right');
  ws.setRowHeight(r++, 14);

  // 凡例
  r++;
  ws.getRange(r, cLabel).setValue('判定凡例 →').setFontWeight('bold').setFontSize(9).setBackground(C.LGRAY);
  [['● 要注意（20%超乖離）', C.RED, C.DRED], ['△ 注意（10%超乖離）', C.YEL, C.DYEL], ['○ 正常範囲', C.GRN, C.DGRN], ['─ 算定不可', C.GRAY, C.DGRAY]]
    .forEach(([label, bg, fc], i) => {
      cell_(ws, r, cDesc + i, label, {bg, fc, bold:true, align:'center', border:true});
    });
  ws.setRowHeight(r++, 20);
  r++;

  // ヘッダー行
  ws.setRowHeight(r, 40);
  cell_(ws, r, cLabel, '指標名',  {bg:C.NAVY, fc:C.WHITE, bold:true, align:'center', wrap:true});
  cell_(ws, r, cDesc,  '計算式',  {bg:C.NAVY, fc:C.WHITE, bold:true, align:'center', wrap:true});
  periods.forEach((p, i) => cell_(ws, r, cPStart + i, p, {bg:C.NAVY, fc:C.WHITE, bold:true, align:'center', wrap:true}));
  cell_(ws, r, cAvg,   `業界平均\n${settings.industryName}`, {bg:C.BLUE, fc:C.WHITE, bold:true, align:'center', wrap:true});
  cell_(ws, r, cDiff,  '差異\n(最新期−平均)',               {bg:C.BLUE, fc:C.WHITE, bold:true, align:'center', wrap:true});
  cell_(ws, r, cJudge, '判定',                              {bg:C.BLUE, fc:C.WHITE, bold:true, align:'center', wrap:true});
  cell_(ws, r, cNote,  'AIによる分析コメント',               {bg:C.NAVY, fc:C.WHITE, bold:true, align:'center', wrap:true});
  r++;

  // 比較データ
  const comparisons = result.industry_comparison || [];
  const sectionMap  = { '売上総利益率':'▶ 収益性分析', '自己資本比率':'▶ 安全性分析', 'ROA':'▶ 効率性分析', '売上高成長率':'▶ 成長性分析' };

  comparisons.forEach(item => {
    if (sectionMap[item.metric]) {
      titleRow(ws, r, sectionMap[item.metric], totalCols, C.LBLUE, C.NAVY, 20);
      r++;
    }

    ws.setRowHeight(r, 36);
    cell_(ws, r, cLabel, item.metric, {align:'left'});
    cell_(ws, r, cDesc,  metricDesc_(item.metric), {bg:C.LGRAY, align:'left', fsize:8});

    // 各期実績
    for (let i = 0; i < n; i++) {
      const v = (item.values || [])[i];
      const c = ws.getRange(r, cPStart + i);
      v !== null && v !== undefined
        ? c.setValue(v).setNumberFormat('0.0%').setHorizontalAlignment('right')
        : c.setValue('─').setHorizontalAlignment('center').setBackground(C.LGRAY);
      c.setBorder(true,true,true,true,false,false,'#BFBFBF',SpreadsheetApp.BorderStyle.SOLID);
    }

    // 業界平均
    const avgC = ws.getRange(r, cAvg);
    item.industry_avg !== null && item.industry_avg !== undefined
      ? avgC.setValue(item.industry_avg).setNumberFormat('0.0%').setBackground(C.LLBLUE)
      : avgC.setValue('─').setBackground(C.LGRAY).setHorizontalAlignment('center');
    avgC.setBorder(true,true,true,true,false,false,'#BFBFBF',SpreadsheetApp.BorderStyle.SOLID);

    // 差異
    const diffC = ws.getRange(r, cDiff);
    if (item.latest_diff !== null && item.latest_diff !== undefined) {
      diffC.setValue(item.latest_diff).setNumberFormat('0.0%').setHorizontalAlignment('right');
      const isGood = (item.high_is_good !== false) ? item.latest_diff >= 0 : item.latest_diff <= 0;
      diffC.setBackground(isGood ? '#E2EFDA' : '#FCE4D6').setFontColor(isGood ? C.DGRN : C.DRED).setFontWeight('bold');
    } else {
      diffC.setValue('─').setBackground(C.LGRAY).setHorizontalAlignment('center');
    }
    diffC.setBorder(true,true,true,true,false,false,'#BFBFBF',SpreadsheetApp.BorderStyle.SOLID);

    // 判定
    const alertMap = { red:['●',C.RED,C.DRED], yellow:['△',C.YEL,C.DYEL], green:['○',C.GRN,C.DGRN], na:['─',C.GRAY,C.DGRAY] };
    const [sym, bg, fc] = alertMap[item.alert] || alertMap.na;
    cell_(ws, r, cJudge, sym, {bg, fc, bold:true, align:'center', border:true});

    // コメント
    cell_(ws, r, cNote, item.comment || '', {align:'left', wrap:true, fsize:9});
    r++;
  });

  // 絶対数値
  r++;
  titleRow(ws, r, '▶ 主要財務数値（絶対額）', totalCols, C.LBLUE, C.NAVY, 20);
  r++;

  ws.setRowHeight(r, 30);
  cell_(ws, r, cLabel, '項目', {bg:C.NAVY, fc:C.WHITE, bold:true, align:'center'});
  periods.forEach((p, i) => cell_(ws, r, cPStart + i, p, {bg:C.NAVY, fc:C.WHITE, bold:true, align:'center', wrap:true}));
  ws.getRange(r, cDesc, 1, cNote - cDesc + 1).merge().setBackground(C.NAVY);
  r++;

  const absItems = [
    ['売上高',    'revenue'],       ['売上総利益', 'gross_profit'],
    ['営業利益',  'operating_profit'],['経常利益',  'ordinary_profit'],
    ['当期純利益','net_profit'],     ['EBITDA',    '__ebitda__'],
    ['総資産',    'total_assets'],  ['純資産',     'net_assets'],
    ['有利子負債','interest_bearing_debt'],
  ];

  absItems.forEach(([label, key]) => {
    ws.setRowHeight(r, 18);
    cell_(ws, r, cLabel, label, {bg:C.LLBLUE, bold:true, align:'left'});
    const vals = key === '__ebitda__' ? (result.ebitda || []) : ((result.financials || {})[key] || []);
    for (let i = 0; i < n; i++) {
      const v = vals[i];
      const c = ws.getRange(r, cPStart + i);
      v !== null && v !== undefined
        ? c.setValue(v).setNumberFormat('#,##0').setHorizontalAlignment('right')
        : c.setValue('─').setBackground(C.LGRAY).setHorizontalAlignment('center');
      c.setBackground(v !== null && v !== undefined ? C.WHITE : C.LGRAY)
        .setBorder(true,true,true,true,false,false,'#BFBFBF',SpreadsheetApp.BorderStyle.SOLID);
    }
    ws.getRange(r, cDesc, 1, cNote - cDesc + 1).merge().setBackground(C.LGRAY)
      .setBorder(true,true,true,true,false,false,'#BFBFBF',SpreadsheetApp.BorderStyle.SOLID);
    r++;
  });

  // 総合評価
  r += 2;
  titleRow(ws, r, '▶ AI総合評価・事業承継における留意点', totalCols, C.GOLD, C.NAVY, 22);
  r++;
  ws.getRange(r, 1, 1, totalCols).merge()
    .setValue(result.overall_assessment || '')
    .setBackground(C.GOLD).setFontColor('#333333').setFontSize(9).setWrap(true).setVerticalAlignment('top');
  ws.setRowHeight(r, 120);

  ws.setFrozenRows(6);
}

// ════════════════════════════════════════════════════════════
// 企業価値算定シート 書き込み
// ════════════════════════════════════════════════════════════
function writeValuationSheet_(ss, result, settings, isAfter) {
  const shName = isAfter ? SH.VAL_AFTER : SH.VAL_BEFORE;
  let ws = ss.getSheetByName(shName);
  if (ws) ss.deleteSheet(ws);
  ws = ss.insertSheet(shName);

  const val    = isAfter ? result.valuation_after  : result.valuation_before;
  const meta   = result.meta || {};
  const unit   = settings.unit;
  const suffix = isAfter ? '（修正後）' : '（修正前・帳簿値）';

  ws.setColumnWidth(1, 250); ws.setColumnWidth(2, 220);
  ws.setColumnWidth(3, 140); ws.setColumnWidth(4, 140); ws.setColumnWidth(5, 240);

  let r = 1;

  // タイトル
  titleRow(ws, r, `企業価値算定${suffix}　${settings.company}　基準期: ${meta.valuation_period || '最新期'}`, 5);
  r++;
  ws.getRange(r, 1, 1, 5).merge()
    .setValue(`生成: ${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')}　　単位: ${unit}`)
    .setBackground(C.LGRAY).setFontColor(C.DGRAY).setFontSize(8).setHorizontalAlignment('right');
  ws.setRowHeight(r++, 14);
  r++;

  // ── 基礎財務データ ──
  titleRow(ws, r, isAfter ? '【修正後 基礎財務データ】' : '【基礎財務データ（帳簿値）】', 5, C.LBLUE, C.NAVY, 22);
  r++;

  const baseItems = isAfter ? [
    ['修正後 経常利益',         val ? val.adj_ordinary_profit : null],
    ['修正後 純資産',           val ? val.adj_net_assets      : null],
    ['修正後 EBITDA',          val ? val.adj_ebitda           : null],
    ['売上高（修正なし）',      result.valuation_before ? result.valuation_before.revenue : null],
    ['当期純利益（参考）',      result.valuation_before ? result.valuation_before.net_profit : null],
    ['有利子負債',              result.valuation_before ? result.valuation_before.interest_bearing_debt : null],
  ] : [
    ['経常利益',               val ? val.ordinary_profit       : null],
    ['純資産（簿価）',         val ? val.net_assets            : null],
    ['EBITDA',                val ? val.ebitda                 : null],
    ['売上高',                 val ? val.revenue               : null],
    ['当期純利益',             val ? val.net_profit            : null],
    ['有利子負債',             val ? val.interest_bearing_debt : null],
  ];

  baseItems.forEach(([label, value]) => {
    ws.setRowHeight(r, 20);
    cell_(ws, r, 1, label, {bg:C.LLBLUE, bold:true, align:'left'});
    const c = ws.getRange(r, 2);
    value !== null && value !== undefined
      ? c.setValue(value).setNumberFormat('#,##0').setHorizontalAlignment('right').setFontWeight('bold')
      : c.setValue('─').setHorizontalAlignment('center');
    c.setBackground(C.LLBLUE).setBorder(true,true,true,true,false,false,'#BFBFBF',SpreadsheetApp.BorderStyle.SOLID);
    ws.getRange(r, 3, 1, 3).merge().setValue(unit).setFontColor(C.DGRAY).setFontSize(8)
      .setBackground(C.LGRAY).setBorder(true,true,true,true,false,false,'#BFBFBF',SpreadsheetApp.BorderStyle.SOLID);
    r++;
  });

  // ── 修正項目（修正後のみ）──
  if (isAfter && Array.isArray(result.adjustments) && result.adjustments.length > 0) {
    r++;
    titleRow(ws, r, '【正常化・修正項目（AIによる推奨）】', 5, C.GOLD, C.NAVY, 22);
    r++;

    ws.setRowHeight(r, 30);
    ['No.', '修正項目', `調整額（${unit}）`, '種別', 'AIによる推奨理由'].forEach((h, i) => {
      cell_(ws, r, i + 1, h, {bg:C.BLUE, fc:C.WHITE, bold:true, align:'center', wrap:true, border:true});
    });
    r++;

    result.adjustments.forEach(adj => {
      ws.setRowHeight(r, 45);
      cell_(ws, r, 1, adj.no,   {align:'center'});
      cell_(ws, r, 2, adj.item, {align:'left'});

      const amt = adj.amount || 0;
      const amtC = ws.getRange(r, 3);
      amtC.setValue(amt).setNumberFormat('#,##0').setHorizontalAlignment('right')
        .setBackground(C.INPUT).setFontWeight('bold')
        .setFontColor(amt > 0 ? C.DGRN : amt < 0 ? C.DRED : C.DGRAY)
        .setBorder(true,true,true,true,false,false,'#BFBFBF',SpreadsheetApp.BorderStyle.SOLID);

      cell_(ws, r, 4, adj.type === 'profit' ? '利益調整' : '純資産調整', {bg:C.LGRAY, align:'center', fsize:8});
      cell_(ws, r, 5, adj.reason || '', {align:'left', wrap:true, fsize:9});
      r++;
    });
  }

  // ── 各評価手法 ──
  r++;
  titleRow(ws, r, '【各評価手法による企業価値算定】', 5, C.NAVY, C.WHITE, 22);
  r++;

  ws.setRowHeight(r, 36);
  [`評価手法`, `算出根拠・計算式`, `事業価値(EV)\n${unit}`, `株主価値(EQ)\n${unit}`, `備考・適用可否`].forEach((h, i) => {
    cell_(ws, r, i + 1, h, {bg:C.BLUE, fc:C.WHITE, bold:true, align:'center', wrap:true, border:true});
  });
  r++;

  const methods = val ? (val.methods || {}) : {};
  const methodDefs = [
    ['year_low',  `①年倍法（${settings.yearLow}年・保守的）`],
    ['year_high', `②年倍法（${settings.yearHigh}年・楽観的）`],
    ['ev_ebitda', '③EV/EBITDA法'],
    ['ev_sales',  '④EV/売上高法'],
    ['pbr',       '⑤PBR法（株主価値ベース）'],
    ['per',       '⑥PER法'],
    ['net_assets','⑦純資産法（簿価）'],
  ];

  methodDefs.forEach(([key, label]) => {
    const m = methods[key];
    if (!m) return;
    ws.setRowHeight(r, 24);
    cell_(ws, r, 1, label, {bold:true, align:'left'});
    cell_(ws, r, 2, m.formula || '', {bg:C.LGRAY, align:'left', fsize:8});

    const ok = m.applicable !== false;
    const vBg = ok ? C.GOLD : C.GRAY;

    const evC = ws.getRange(r, 3);
    ok && m.ev !== null && m.ev !== undefined
      ? evC.setValue(m.ev).setNumberFormat('#,##0').setHorizontalAlignment('right').setFontWeight('bold')
      : evC.setValue(ok ? '─' : '算定不可').setHorizontalAlignment('center').setFontColor(C.DGRAY);
    evC.setBackground(vBg).setBorder(true,true,true,true,false,false,'#BFBFBF',SpreadsheetApp.BorderStyle.SOLID);

    const eqC = ws.getRange(r, 4);
    ok && m.equity !== null && m.equity !== undefined
      ? eqC.setValue(m.equity).setNumberFormat('#,##0').setHorizontalAlignment('right').setFontWeight('bold')
      : eqC.setValue(ok ? '─' : '算定不可').setHorizontalAlignment('center').setFontColor(C.DGRAY);
    eqC.setBackground(vBg).setBorder(true,true,true,true,false,false,'#BFBFBF',SpreadsheetApp.BorderStyle.SOLID);

    cell_(ws, r, 5, ok ? '' : '赤字・マイナス純資産のため参考外', {bg:C.LGRAY, fsize:8, fc:C.DGRAY, align:'left'});
    r++;
  });

  // ── 評価レンジ ──
  r += 2;
  titleRow(ws, r, '【評価レンジまとめ（株主価値ベース）】', 5, C.GOLD, C.NAVY, 22);
  r++;

  [['最小値（保守的）', val ? val.range_min : null], ['最大値（楽観的）', val ? val.range_max : null]].forEach(([label, v]) => {
    ws.setRowHeight(r, 20);
    cell_(ws, r, 1, label, {bg:C.LLBLUE, bold:true, align:'left'});
    const c = ws.getRange(r, 2);
    v !== null && v !== undefined ? c.setValue(v).setNumberFormat('#,##0').setHorizontalAlignment('right').setFontWeight('bold') : c.setValue('─').setHorizontalAlignment('center');
    c.setBackground(C.LLBLUE).setBorder(true,true,true,true,false,false,'#BFBFBF',SpreadsheetApp.BorderStyle.SOLID);
    ws.getRange(r, 3, 1, 3).merge().setBackground(C.LGRAY);
    r++;
  });

  ws.setRowHeight(r, 32);
  cell_(ws, r, 1, '推定評価レンジ', {bg:C.GOLD, bold:true, align:'left', fsize:11});
  const minV = val ? val.range_min : null;
  const maxV = val ? val.range_max : null;
  const rangeStr = (minV !== null && maxV !== null)
    ? `${Number(minV).toLocaleString()}  〜  ${Number(maxV).toLocaleString()}　${unit}`
    : 'データ不足のため算定できません';
  ws.getRange(r, 2, 1, 4).merge()
    .setValue(rangeStr).setBackground(C.GOLD).setFontWeight('bold').setFontSize(13)
    .setFontColor(C.NAVY).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true,true,true,true,false,false,'#999999',SpreadsheetApp.BorderStyle.SOLID);
  r += 2;

  // AIコメント
  titleRow(ws, r, '【AIによる評価コメント】', 5, C.GOLD, C.NAVY, 20);
  r++;
  ws.getRange(r, 1, 1, 5).merge()
    .setValue(val ? (val.assessment || '') : '')
    .setBackground(C.GOLD).setFontColor('#333333').setFontSize(9).setWrap(true).setVerticalAlignment('top');
  ws.setRowHeight(r++, 90);

  // 注記
  r++;
  ws.getRange(r, 1, 1, 5).merge()
    .setValue(
      '【注記】\n' +
      '・本算定は参考値です。実際の取引価格は当事者間の交渉・デューデリジェンス結果により決定されます。\n' +
      '・非上場中小企業のM&A・事業承継では、上場会社マルチプルから30〜50%程度の割引が一般的です。\n' +
      '・有利子負債は全負債で代替しています。実際の算定では借入金・社債等の有利子負債のみで計算してください。\n' +
      '・業界平均倍率は東証上場会社の中央値ベースの参考値です。実際の類似企業データで検証することを推奨します。'
    )
    .setBackground(C.LGRAY).setFontColor(C.DGRAY).setFontSize(8).setWrap(true).setVerticalAlignment('top');
  ws.setRowHeight(r, 75);

  ws.setFrozenRows(2);
}

// ════════════════════════════════════════════════════════════
// ユーティリティ
// ════════════════════════════════════════════════════════════
function titleRow(ws, r, text, cols, bg, fc, h) {
  ws.setRowHeight(r, h || 28);
  ws.getRange(r, 1, 1, cols).merge()
    .setValue(text)
    .setBackground(bg || C.NAVY).setFontColor(fc || C.WHITE)
    .setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setBorder(true,true,true,true,false,false,'#888888',SpreadsheetApp.BorderStyle.SOLID);
}

function cell_(ws, r, col, value, opts) {
  opts = opts || {};
  const c = ws.getRange(r, col);
  if (value !== undefined) c.setValue(value);
  if (opts.bg)    c.setBackground(opts.bg);
  if (opts.fc)    c.setFontColor(opts.fc);
  if (opts.bold !== undefined) c.setFontWeight(opts.bold ? 'bold' : 'normal');
  if (opts.fsize) c.setFontSize(opts.fsize);
  if (opts.align) c.setHorizontalAlignment(opts.align);
  if (opts.wrap)  c.setWrap(true);
  c.setVerticalAlignment('middle');
  if (opts.border !== false)
    c.setBorder(true,true,true,true,false,false,'#BFBFBF',SpreadsheetApp.BorderStyle.SOLID);
  return c;
}

function toast_(msg, ss) {
  (ss || SpreadsheetApp.getActiveSpreadsheet()).toast(msg, 'AI財務分析', -1);
}

function pct(v) {
  return v !== null && v !== undefined ? (v * 100).toFixed(0) + '%' : 'N/A';
}

function metricDesc_(metric) {
  return ({
    '売上総利益率': '売上総利益 ÷ 売上高',
    '営業利益率':   '営業利益 ÷ 売上高',
    '経常利益率':   '経常利益 ÷ 売上高',
    '純利益率':     '当期純利益 ÷ 売上高',
    '人件費率':     '人件費 ÷ 売上高',
    '自己資本比率': '純資産 ÷ 総資産',
    '流動比率':     '流動資産 ÷ 流動負債',
    '固定比率':     '固定資産 ÷ 純資産',
    'ROA':         '経常利益 ÷ 総資産',
    'ROE':         '当期純利益 ÷ 純資産',
    '売上高成長率': '(当期−前期売上) ÷ 前期売上',
  })[metric] || '';
}
