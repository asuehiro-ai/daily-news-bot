// ============================================================
// 事業承継 簡易診断システム
// Google Apps Script — Googleスプレッドシート連携 → スライド自動生成
//
// 【セットアップ手順】
// 1. Googleフォームと連携しているスプレッドシートを開く
// 2. メニュー「拡張機能」→「Apps Script」
// 3. このコードを貼り付けて保存（Ctrl+S）
// 4. 「実行」→「onOpen」を一度実行して権限を付与
// 5. スプレッドシートに戻ると「📊 事業承継診断」メニューが出現
// ============================================================

// ── 設定 ─────────────────────────────────────────────────────
const CONFIG = {
  // 回答シート名（Google Formsと連携しているシート名に合わせて変更）
  SHEET_NAME: 'フォームの回答 1',

  // 作成したスライドの保存先フォルダ名（Googleドライブに自動作成）
  OUTPUT_FOLDER: '事業承継診断レポート',

  // M&A 純資産倍率（純資産 × この倍率で評価）
  MA_MULTIPLE: 3,

  // 資本金（取得価額）デフォルト値 — フォームに資本金項目がない場合の仮定値
  // 実際の資本金が分かる場合はフォームに追加するか、ここを調整してください
  DEFAULT_CAPITAL: 10_000_000,  // 1,000万円

  // ブランドカラー（株式会社LEG）
  C: {
    PRIMARY_DARK:  '#002e3a',
    PRIMARY:       '#005060',
    PRIMARY_LIGHT: '#006d86',
    ACCENT:        '#c07228',
    ACCENT_LIGHT:  '#e8a85a',
    WHITE:         '#ffffff',
    LIGHT_BG:      '#eef3f5',
    TEXT_DARK:     '#1a2533',
    TEXT_MUTED:    '#6b7c93',
    GREEN_BG:      '#e8f5e9',
    GREEN_TEXT:    '#1b5e20',
    AMBER_BG:      '#fff8e1',
    AMBER_TEXT:    '#c07228',
    ORANGE_BG:     '#fff3e0',
    ORANGE_TEXT:   '#7c3a00',
    ROW_ODD:       '#f5f8fa',
  },
};

// ── メニュー追加 ─────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 事業承継診断')
    .addItem('▶ 最新の回答を診断する', 'generateLatest')
    .addSeparator()
    .addItem('全回答を一括診断', 'generateAll')
    .addItem('行番号を指定して診断', 'generateByInput')
    .addToUi();
}

// ── エントリーポイント ────────────────────────────────────────
function generateLatest() {
  const sheet = _getSheet();
  const last  = sheet.getLastRow();
  if (last < 2) { _alert('回答データがありません。'); return; }
  _processRow(sheet, last, true);
}

function generateAll() {
  const sheet = _getSheet();
  const last  = sheet.getLastRow();
  if (last < 2) { _alert('回答データがありません。'); return; }

  const ui = SpreadsheetApp.getUi();
  if (ui.alert(`${last - 1}件の回答を処理します。よろしいですか？`, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  let count = 0;
  for (let r = 2; r <= last; r++) {
    if (_processRow(sheet, r, false)) count++;
  }
  _alert(`✅ ${count}件の診断スライドを作成しました。\nDriveフォルダ「${CONFIG.OUTPUT_FOLDER}」をご確認ください。`);
}

function generateByInput() {
  const ui  = SpreadsheetApp.getUi();
  const res = ui.prompt('診断する行番号を入力（2以上の整数）:');
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const row = parseInt(res.getResponseText());
  if (isNaN(row) || row < 2) { _alert('有効な行番号を入力してください。'); return; }
  const sheet = _getSheet();
  if (row > sheet.getLastRow()) { _alert(`行 ${row} にデータがありません。`); return; }
  _processRow(sheet, row, true);
}

function _getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getSheets()[0];
}

function _alert(msg) {
  SpreadsheetApp.getUi().alert(msg);
}

// ── 1行を処理 ────────────────────────────────────────────────
function _processRow(sheet, rowIndex, showAlert) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data    = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];

  const row = {};
  headers.forEach((h, i) => { row[h.toString().trim()] = data[i]; });

  const diag = _calcDiagnosis(row);
  const url  = _createSlides(row, diag);

  if (showAlert && url) {
    _alert(`✅ 診断スライドを作成しました！\n\n${url}`);
  }
  return url;
}

// ── フィールド取得ヘルパー ────────────────────────────────────
// フォームの質問文が変わっても、キーワードで柔軟にマッチングします
function _field(row, ...keywords) {
  for (const kw of keywords) {
    for (const [k, v] of Object.entries(row)) {
      if (k.includes(kw)) return (v === undefined || v === null) ? '' : v;
    }
  }
  return '';
}

// Yes/No 回答の正規化
// 「保証あり」「している」「はい」→ true、「保証なし」「していない」「いいえ」→ false
function _isYes(str) {
  if (!str || str.toString().trim() === '') return false;
  const s = str.toString();
  // 否定表現を先にチェック（優先させる）
  if (s.includes('していない') || s.includes('いない') || s.includes('なし')
      || s.includes('いいえ')  || s.includes('ない'))  return false;
  // 肯定表現
  return s.includes('している') || s.includes('あり')  || s.includes('いる')
      || s.includes('ある')     || s.includes('はい')  || s.includes('供し');
}

function _num(val) {
  if (val === null || val === undefined || val === '') return null;
  if (typeof val === 'number') return val;
  const n = parseFloat(val.toString().replace(/,|、|円/g, ''));
  return isNaN(n) ? null : n;
}

function _jpy(n) {
  if (n === null || n === undefined || n === '') return '—';
  return `${Math.round(n).toLocaleString()}円`;
}

function _jpyShort(n) {
  if (!n) return '—';
  const abs = Math.abs(n);
  if (abs >= 1e8)  return `${(n / 1e8).toFixed(1)}億円`;
  if (abs >= 1e4)  return `${Math.round(n / 1e4)}万円`;
  return `${n.toLocaleString()}円`;
}

// ── 税額計算 ─────────────────────────────────────────────────

// 贈与税（特例贈与財産）— 基礎控除110万円
function _giftTax(val) {
  const t = Math.max(0, val - 1_100_000);
  if (t <= 0)          return 0;
  if (t <= 2_000_000)  return Math.round(t * 0.10);
  if (t <= 3_000_000)  return Math.round(t * 0.15 - 100_000);
  if (t <= 4_000_000)  return Math.round(t * 0.20 - 250_000);
  if (t <= 6_000_000)  return Math.round(t * 0.30 - 650_000);
  if (t <= 10_000_000) return Math.round(t * 0.40 - 1_250_000);
  if (t <= 15_000_000) return Math.round(t * 0.45 - 1_750_000);
  if (t <= 30_000_000) return Math.round(t * 0.50 - 2_500_000);
  return Math.round(t * 0.55 - 4_000_000);
}

// みなし配当 総合課税（所得税＋住民税10%、簡易計算）
function _deemedDivTax(d) {
  if (d <= 0)            return 0;
  if (d <= 1_950_000)    return Math.round(d * 0.15);
  if (d <= 3_300_000)    return Math.round(d * 0.20 - 97_500);
  if (d <= 6_950_000)    return Math.round(d * 0.30 - 427_500);
  if (d <= 9_000_000)    return Math.round(d * 0.33 - 636_000);
  if (d <= 18_000_000)   return Math.round(d * 0.43 - 1_536_000);
  if (d <= 40_000_000)   return Math.round(d * 0.50 - 2_796_000);
  return Math.round(d * 0.55 - 4_796_000);
}

// ── 診断計算 ─────────────────────────────────────────────────
function _calcDiagnosis(row) {
  const netAssets  = _num(_field(row, '純資産'))   || 0;
  const debt       = _num(_field(row, '借入金'))   || 0;
  const revenue    = _num(_field(row, '売上高'))   || 0;
  const opProfit   = _num(_field(row, '営業利益')) || 0;
  const netIncome  = _num(_field(row, '最終利益', '当期純利益')) || opProfit * 0.6;

  // 資本金（取得価額）= フォームに項目があれば使い、なければデフォルト
  const capitalRaw = _num(_field(row, '資本金'));
  const capital    = capitalRaw
    ? capitalRaw
    : Math.min(CONFIG.DEFAULT_CAPITAL, Math.max(1_000_000, netAssets * 0.05));

  // ① 株式評価（純資産価額方式 簡易）
  const stockVal = netAssets;

  // ② M&A 株式価値（純資産 × 3倍）
  const maEV      = netAssets * CONFIG.MA_MULTIPLE;
  const maLabel   = `純資産×${CONFIG.MA_MULTIPLE}倍`;
  // エクイティ価値 = EV − 有利子負債（最低でも純資産相当）
  const maEquity   = Math.max(maEV - debt, netAssets);
  const maMultiple = `${CONFIG.MA_MULTIPLE}倍`;

  // ③ 廃業 清算価値（10%清算コスト控除）
  const liquidVal = Math.round(netAssets * 0.9);

  // ④ 親族内承継（税制なし）— 後継者の贈与税
  const giftTax = _giftTax(stockVal);

  // ⑤ 社内承継（MBO）— 社長の譲渡所得課税 20.315%
  const mboGain     = Math.max(0, stockVal - capital);
  const mboTax      = Math.round(mboGain * 0.20315);
  const mboTakehome = stockVal - mboTax;

  // ⑥ M&A — 社長の譲渡所得課税 20.315%
  const maGain      = Math.max(0, maEquity - capital);
  const maTax       = Math.round(maGain * 0.20315);
  const maTakehome  = maEquity - maTax;

  // ⑦ 廃業（清算）— みなし配当 総合課税
  const deemedDiv    = Math.max(0, liquidVal - capital);
  const closureTax   = _deemedDivTax(deemedDiv);
  const closureTakeh = Math.max(0, liquidVal - closureTax);

  // 手取りランキング（金額が大きい順）
  const ranking = [
    {
      name:     `M&A（${maMultiple}）`,
      takehome: maTakehome,
      comment:  '創業者利益最大化・税率一定（20.315%）・経営者保証解除期待',
    },
    {
      name:     '社内承継（MBO）',
      takehome: mboTakehome,
      comment:  '購入資金調達が最大ハードル・社内候補の力量次第',
    },
    {
      name:     '廃業（清算）',
      takehome: closureTakeh,
      comment:  '従業員雇用喪失・他方法が実現不可な場合のみ検討',
    },
    {
      name:     '親族内（税制適用）',
      takehome: 0,
      comment:  '2027年期限の特例措置・後継者の贈与税は猶予・免除',
    },
    {
      name:     '親族内（税制なし）',
      takehome: 0,
      comment:  '評価は経営状況により変動するため早期準備が鍵',
    },
  ].sort((a, b) => b.takehome - a.takehome);

  return {
    // 財務データ
    netAssets, debt, revenue, opProfit, netIncome, capital,
    // 評価額
    stockVal, maEquity, maLabel, maMultiple, liquidVal,
    // 親族
    giftTax,
    // MBO
    mboTax, mboTakehome,
    // M&A
    maTax, maTakehome,
    // 廃業
    closureTax, closureTakeh,
    // ランキング
    ranking,
  };
}

// ── スライド生成 ──────────────────────────────────────────────
function _createSlides(row, diag) {
  // 会社名・社長名はキーワードマッチ → なければ列位置（3列目・4列目）でフォールバック
  const vals      = Object.values(row);
  const company   = (_field(row, '会社名', '商号', '法人名', '社名') || vals[2] || '').toString() || '—';
  const president = (_field(row, '社長名', '代表者名', '氏名', 'ご氏名', 'お名前', '代表') || vals[3] || '').toString() || '—';
  const industry  = _field(row, '業種').toString()    || '—';
  const age       = _field(row, '年齢').toString()    || '—';
  const employees = _field(row, '従業員').toString()  || '—';
  const priorities= _field(row, '優先', '重要視').toString();
  const preferred = _field(row, '選択肢', '考えている').toString();

  const familyRaw    = _field(row, '親族').toString();
  const internalRaw  = _field(row, '社内').toString();
  const guaranteeRaw = _field(row, '経営者保証').toString();
  const collateralRaw= _field(row, '担保', '不動産').toString();

  const hasFamily     = _isYes(familyRaw);
  const hasInternal   = _isYes(internalRaw);
  const hasGuarantee  = _isYes(guaranteeRaw);
  const hasCollateral = _isYes(collateralRaw);

  const today   = new Date();
  const dateStr = `${today.getFullYear()}年${today.getMonth()+1}月${today.getDate()}日`;
  const C       = CONFIG.C;

  // プレゼンテーション作成（Googleドライブのマイドライブ直下に作成される）
  const pres = SlidesApp.create(`【LEG】${company} 事業承継簡易診断 ${dateStr}`);
  const W    = pres.getPageWidth();   // pt（通常720）
  const H    = pres.getPageHeight();  // pt（通常405）

  // ─── Slide 1: タイトル ────────────────────────────────────
  const s1 = pres.getSlides()[0];
  _clearSlide(s1);
  s1.getBackground().setSolidFill(C.PRIMARY);

  _rect(s1, 0, 0, W, 6, C.ACCENT);                                       // 上線
  _rect(s1, 0, H - 6, W, 6, C.ACCENT);                                   // 下線
  _text(s1, 'LEG', 44, 20, 100, 28, C.ACCENT, 18, true);
  _text(s1, 'M&A仲介・事業承継支援', 44, 50, 260, 18, C.ACCENT_LIGHT, 8.5, false);
  _text(s1, '事業承継 簡易診断レポート', 60, 100, W - 80, 58, C.WHITE, 30, true);
  _rect(s1, 60, 164, 100, 3, C.ACCENT);
  _text(s1, company,        60, 178, W - 80, 38, C.WHITE,     18, true);
  _text(s1, `${president} 様`, 60, 220, 400, 28, '#aaccdd',  13, false);
  _text(s1, dateStr,        60, 325, 280, 22, '#aaccdd',      10, false);
  _text(s1, '株式会社 LEG', W - 200, 340, 160, 20, '#aaccdd', 10, false);

  // ─── Slide 2: 企業概要 ────────────────────────────────────
  const s2 = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  _header(s2, '企業概要・財務サマリー', company, C, W);

  const ovData = [
    ['会社名',                company,                    '業種',         industry],
    ['社長名',                `${president}（${age}歳）`, '従業員数',     `${employees}名`],
    ['純資産（直近期）',      _jpy(diag.netAssets),       '借入金残高',   _jpy(diag.debt)],
    ['売上高',                _jpy(diag.revenue),         '営業利益',     _jpy(diag.opProfit)],
    ['株式評価（簡易・純資産）', _jpy(diag.stockVal),     'M&A想定価値',  _jpy(diag.maEquity)],
  ];
  const t2 = s2.insertTable(5, 4, 28, 76, W - 56, 278);
  ovData.forEach((rd, r) => rd.forEach((cell, c) => {
    const tc = t2.getCell(r, c);
    tc.getText().setText(cell.toString());
    const ts = tc.getText().getTextStyle();
    ts.setFontFamily('Noto Sans JP').setFontSize(10);
    if (c % 2 === 0) {
      tc.getFill().setSolidFill(C.PRIMARY);
      ts.setForegroundColor(C.WHITE).setBold(true);
    } else {
      tc.getFill().setSolidFill(r % 2 === 0 ? C.LIGHT_BG : C.WHITE);
      ts.setForegroundColor(C.TEXT_DARK);
    }
  }));

  // チェックリスト表示
  const checks = [
    `親族後継者：${hasFamily    ? '✅ あり' : '❌ なし'}`,
    `社内後継者：${hasInternal  ? '✅ あり' : '❌ なし'}`,
    `経営者保証：${hasGuarantee ? '⚠️ あり（要解除）' : '✅ なし'}`,
    `不動産担保：${hasCollateral ? '⚠️ あり' : '✅ なし'}`,
  ].join('    ');
  _text(s2, checks, 28, 366, W - 56, 22, C.TEXT_MUTED, 8.5, false);

  // ─── Slide 3: 社長の手取り比較 ───────────────────────────
  const s3 = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  _header(s3, '《社長の手取り比較》', company, C, W);

  const maName = `M&A（${diag.maMultiple}）`;
  const compData = [
    ['承継方法',        '株式評価',                  '社長の税額',             '社長の手取り',            '税負担者',             '承継時期目安'],
    ['親族内（税制なし）', _jpy(diag.stockVal),      '0',                      '0（後継者負担）',         '後継者が贈与税',        '5年以内'],
    ['親族内（税制適用）', _jpy(diag.stockVal),      '猶予・免除',             '0（後継者負担）',         '後継者が贈与税（猶予）', '5年以内\n※2027年期限'],
    ['社内承継（MBO）', _jpy(diag.stockVal),         _jpy(diag.mboTax),        _jpy(diag.mboTakehome),   '社長',                  '3〜5年'],
    [maName,           _jpy(diag.maEquity),          _jpy(diag.maTax),         _jpy(diag.maTakehome),    '社長',                  '1年前後'],
    ['廃業（清算）',   _jpy(diag.liquidVal),         _jpy(diag.closureTax),    _jpy(diag.closureTakeh),  '社長',                  '1〜3年'],
  ];
  const t3 = s3.insertTable(6, 6, 18, 76, W - 36, 290);
  compData.forEach((rd, r) => rd.forEach((cell, c) => {
    const tc = t3.getCell(r, c);
    tc.getText().setText(cell.toString());
    const ts = tc.getText().getTextStyle();
    ts.setFontFamily('Noto Sans JP');
    if (r === 0) {
      tc.getFill().setSolidFill(C.PRIMARY_DARK);
      ts.setFontSize(8).setForegroundColor(C.WHITE).setBold(true);
    } else if (r === 4 && c === 3) {           // M&A 手取り — 金色ハイライト
      tc.getFill().setSolidFill(C.AMBER_BG);
      ts.setFontSize(9).setForegroundColor(C.AMBER_TEXT).setBold(true);
    } else if (c === 0) {
      tc.getFill().setSolidFill(r % 2 === 0 ? '#dde8ec' : C.LIGHT_BG);
      ts.setFontSize(8).setForegroundColor(C.PRIMARY_DARK).setBold(true);
    } else {
      tc.getFill().setSolidFill(r % 2 === 0 ? C.ROW_ODD : C.WHITE);
      ts.setFontSize(8.5).setForegroundColor(C.TEXT_DARK);
    }
  }));
  _text(s3, '※本試算は社長100%保有と仮定した概算です。実際の税額は必ず税理士・M&Aアドバイザーにご相談ください。',
        18, H - 28, W - 36, 20, '#999999', 7.5, false);

  // ─── Slide 4: 後継者別 負担内訳 ─────────────────────────
  const s4 = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  _header(s4, '《後継者別 負担内訳》', company, C, W);

  const burdData = [
    ['承継方法',          '後継者側',       '負担内容',              '金額',                   '購入資金調達'],
    ['親族内（税制なし）', '後継者（子等）', '贈与税',               _jpy(diag.giftTax),        '不要'],
    ['親族内（税制適用）', '後継者（子等）', '贈与税（猶予・免除）', '猶予・免除',              '不要'],
    ['社内承継（MBO）',   '社内候補者',     '株式購入代金',         _jpy(diag.stockVal),        '要（10年返済目安）'],
    [maName,              '買収企業',       '株式譲渡代金',         _jpy(diag.maEquity),        '不要（買収側が調達）'],
    ['廃業（清算）',      '社長等株主',     'みなし配当税（総合課税）', _jpy(diag.closureTax),  '不要'],
  ];
  const t4 = s4.insertTable(6, 5, 18, 76, W - 36, 290);
  burdData.forEach((rd, r) => rd.forEach((cell, c) => {
    const tc = t4.getCell(r, c);
    tc.getText().setText(cell.toString());
    const ts = tc.getText().getTextStyle();
    ts.setFontFamily('Noto Sans JP');
    if (r === 0) {
      tc.getFill().setSolidFill(C.PRIMARY_DARK);
      ts.setFontSize(9).setForegroundColor(C.WHITE).setBold(true);
    } else if (c === 0) {
      tc.getFill().setSolidFill(C.PRIMARY);
      ts.setFontSize(8.5).setForegroundColor(C.WHITE).setBold(true);
    } else {
      tc.getFill().setSolidFill(r % 2 === 0 ? C.LIGHT_BG : C.WHITE);
      ts.setFontSize(9).setForegroundColor(C.TEXT_DARK);
    }
  }));
  _text(s4, '※後継者側の税・費用の概算です。資金調達条件により変動します。',
        18, H - 28, W - 36, 20, '#999999', 7.5, false);

  // ─── Slide 5: 社長手取りランキング ──────────────────────
  const s5 = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  _header(s5, '《社長手取りランキング》', company, C, W);

  const rankRows = [['順位', '承継方法', '社長手取り', '推奨コメント']].concat(
    diag.ranking.map((opt, i) => [
      `${i + 1}位`,
      opt.name,
      opt.takehome > 0 ? _jpy(opt.takehome) : '—（後継者負担）',
      opt.comment,
    ])
  );
  const t5 = s5.insertTable(6, 4, 18, 76, W - 36, 290);
  rankRows.forEach((rd, r) => rd.forEach((cell, c) => {
    const tc = t5.getCell(r, c);
    tc.getText().setText(cell.toString());
    const ts = tc.getText().getTextStyle();
    ts.setFontFamily('Noto Sans JP');
    if (r === 0) {
      tc.getFill().setSolidFill(C.PRIMARY_DARK);
      ts.setFontSize(10).setForegroundColor(C.WHITE).setBold(true);
    } else if (r === 1) {                      // 1位 — ゴールドハイライト
      tc.getFill().setSolidFill(C.AMBER_BG);
      ts.setFontSize(c === 3 ? 8 : 10).setForegroundColor(C.AMBER_TEXT).setBold(true);
    } else {
      tc.getFill().setSolidFill(r % 2 === 0 ? C.LIGHT_BG : C.WHITE);
      ts.setFontSize(c === 3 ? 8 : 9.5).setForegroundColor(C.TEXT_DARK);
      if (c === 0) ts.setBold(true);
    }
  }));

  // ─── Slide 6: メリット・デメリット ──────────────────────
  const s6 = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  _header(s6, '《主なメリット・デメリット》', company, C, W);

  const mdData = [
    ['承継方法', 'メリット', 'デメリット'],
    [
      '親族内承継',
      '・会社と家族経営の継続\n・従業員・取引先の安心感\n・事業承継税制（2027年期限）で\n　後継者の税負担を大幅軽減可',
      '・税制なしの場合は贈与税負担が重い\n・後継者の資質・意欲・能力が鍵\n・後継者への説得・準備期間が必要',
    ],
    [
      '社内承継（MBO）',
      '・経営理念・文化の継続\n・従業員主体の安定的経営\n・社内への安心感',
      '・株式購入資金の調達が最大ハードル\n・社長の手取りが他手法より少ない\n・候補者の経営能力の見極めが必要',
    ],
    [
      'M&A（第三者承継）',
      '・社長の創業者利益を最大化\n・経営者保証・不動産担保の解除期待\n・買収先とのシナジーで事業成長\n・税率が一定（20.315%）で計算しやすい',
      '・経営継続・従業員離反リスク\n・相手先選定・交渉に1年前後を要する\n・情報漏洩リスクへの配慮が必要',
    ],
    [
      '廃業（清算）',
      '・手続きが比較的シンプル\n　（過大債務がない場合）',
      '・従業員の雇用が失われる\n・取引先・地域への影響が大きい\n・社長の手取りが最も少ない\n・事業価値がゼロになる',
    ],
  ];
  const t6 = s6.insertTable(5, 3, 18, 76, W - 36, 290);
  mdData.forEach((rd, r) => rd.forEach((cell, c) => {
    const tc = t6.getCell(r, c);
    tc.getText().setText(cell.toString());
    const ts = tc.getText().getTextStyle();
    ts.setFontFamily('Noto Sans JP');
    if (r === 0) {
      tc.getFill().setSolidFill(C.PRIMARY_DARK);
      ts.setFontSize(10).setForegroundColor(C.WHITE).setBold(true);
    } else if (c === 0) {
      tc.getFill().setSolidFill(C.PRIMARY);
      ts.setFontSize(9).setForegroundColor(C.WHITE).setBold(true);
    } else if (c === 1) {
      tc.getFill().setSolidFill(r % 2 === 0 ? C.GREEN_BG : '#f1f8f2');
      ts.setFontSize(8).setForegroundColor(C.GREEN_TEXT);
    } else {
      tc.getFill().setSolidFill(r % 2 === 0 ? C.ORANGE_BG : '#fdf8f0');
      ts.setFontSize(8).setForegroundColor(C.ORANGE_TEXT);
    }
  }));

  // ─── Slide 7: 推奨・初回ディスカッションポイント ─────────
  const s7 = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  _header(s7, '初回ディスカッション 推奨ポイント', company, C, W);

  const ageNum = parseInt(age) || 0;
  const rec = _buildRec(diag, hasFamily, hasInternal, hasGuarantee, hasCollateral, priorities, ageNum);

  // メインテキスト
  const recBox = s7.insertTextBox(rec.body, 28, 76, W - 56, 238);
  recBox.getText().getTextStyle()
    .setFontFamily('Noto Sans JP').setFontSize(9.5).setForegroundColor(C.TEXT_DARK);
  recBox.setContentAlignment(SlidesApp.ContentAlignment.TOP);

  // 推奨方法ハイライトボックス
  _rect(s7, 28, 322, W - 56, 48, C.LIGHT_BG);
  _rect(s7, 28, 322, 5, 48, C.ACCENT);
  _text(s7, `📌 推奨承継方法：${rec.topChoice}`, 40, 330, W - 80, 32, C.PRIMARY_DARK, 12, true);

  _text(s7, '※本試算は概算（社長100%保有と仮定）です。実際の税額・手続きは税理士・M&Aアドバイザーに必ずご相談ください。',
        18, H - 24, W - 36, 18, '#aaaaaa', 7, false);

  // ─── Driveの指定フォルダへ移動 ────────────────────────────
  const folder = _getOrMakeFolder(CONFIG.OUTPUT_FOLDER);
  const file   = DriveApp.getFileById(pres.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  return pres.getUrl();
}

// ── 推奨コメント生成 ─────────────────────────────────────────
function _buildRec(diag, hasFamily, hasInternal, hasGuarantee, hasCollateral, priorities, ageNum) {
  const lines = [];

  // ① 年齢・承継時期の緊急度
  if (ageNum > 0) {
    lines.push('【年齢・承継時期について】');
    if (ageNum < 55) {
      lines.push('  現在' + ageNum + '歳で、承継準備に比較的余裕がある段階です。ただし、');
      lines.push('  事業承継税制（特例）の2027年3月末期限を踏まえると、');
      lines.push('  親族内承継をお考えの場合は早めのご検討をお勧めします。');
    } else if (ageNum < 65) {
      lines.push('  現在' + ageNum + '歳で、本格的な検討・準備を始める時期です。');
      lines.push('  5年以内の承継完了を視野に、候補者や方向性の絞り込みを');
      lines.push('  早期に進めることをお勧めします。');
    } else if (ageNum < 70) {
      lines.push('  現在' + ageNum + '歳で、具体的な実行フェーズに入る時期です。');
      lines.push('  3〜5年以内の完了を目標に、専門家を交えた準備を');
      lines.push('  早急に進めることをお勧めします。');
    } else {
      lines.push('  現在' + ageNum + '歳で、早期の具体的な行動が求められる段階です。');
      lines.push('  M&Aであれば通常1年前後での実現が可能なため、');
      lines.push('  迅速な意思決定が重要です。');
    }
    lines.push('');
  }

  // ② 後継者候補の状況
  lines.push('【後継者候補の状況について】');
  if (hasFamily && hasInternal) {
    lines.push('  親族・社内の両方に候補者がいらっしゃいます。それぞれの意向・');
    lines.push('  準備状況を確認した上で、方向性を早期に絞り込むことが重要です。');
  } else if (hasFamily) {
    lines.push('  親族内での承継を基軸に検討を進めることが考えられます。候補者との');
    lines.push('  意思確認と、2027年3月末期限の事業承継税制の活用もご検討ください。');
  } else if (hasInternal) {
    lines.push('  社内承継（MBO）が選択肢の一つとして挙げられます。候補者の');
    lines.push('  経営意欲・資金調達力の確認が次のステップです。M&Aも並行検討可です。');
  } else {
    lines.push('  現時点では内部での後継者候補が不在のため、M&Aによる');
    lines.push('  第三者承継が有力な選択肢となります。早期の情報収集をお勧めします。');
  }

  // ③ 経営者保証・担保の状況
  lines.push('');
  lines.push('【経営者保証・担保の状況について】');
  if (hasGuarantee && hasCollateral) {
    lines.push('  経営者保証・個人担保ともに設定されています。承継の際には');
    lines.push('  金融機関との事前協議が不可欠です。M&Aにより保証・担保が');
    lines.push('  解除されるケースも多く、有効な解決策となり得ます。');
  } else if (hasGuarantee) {
    lines.push('  経営者保証が設定されています。M&Aや社内承継の過程で解除');
    lines.push('  交渉が期待できます。金融機関への早期相談をお勧めします。');
  } else if (hasCollateral) {
    lines.push('  個人不動産が担保設定されています。承継に際して担保の取り扱いを');
    lines.push('  金融機関と事前に協議しておくことが重要です。');
  } else {
    lines.push('  経営者保証・個人担保なし。金融面では承継手続きを');
    lines.push('  進めやすい条件が整っています。');
  }

  // ④ 優先事項との整合
  if (priorities) {
    lines.push('');
    lines.push('【ご重視されている点との整合について】');
    if (priorities.includes('従業員') || priorities.includes('雇用'))
      lines.push('  雇用継続を重視 → 親族内・社内承継が適合。M&Aでも条件交渉が可能です。');
    if (priorities.includes('利益') || priorities.includes('最大') || priorities.includes('手取り'))
      lines.push('  手取り最大化を重視 → M&Aが最も高い可能性。市場環境で変動します。');
    if (priorities.includes('早期') || priorities.includes('速やか'))
      lines.push('  早期実現を重視 → M&A（1年前後）が最も現実的な選択肢です。');
    if (priorities.includes('理念') || priorities.includes('継続') || priorities.includes('文化'))
      lines.push('  理念・文化の継続を重視 → 親族内・社内承継が適合的です。');
    if (priorities.includes('地域') || priorities.includes('社会'))
      lines.push('  地域貢献を重視 → 地域企業へのM&Aや社内承継が適合的です。');
  }

  // 推奨選択肢の決定
  let topChoice = diag.ranking[0].name;
  if (!hasFamily && topChoice.includes('親族'))
    topChoice = diag.ranking.find(o => !o.name.includes('親族'))?.name || 'M&A';
  if (!hasInternal && topChoice.includes('MBO'))
    topChoice = diag.ranking.find(o => !o.name.includes('MBO') && !o.name.includes('親族'))?.name || 'M&A';

  return { body: lines.join('\n'), topChoice };
}

// ── スライドユーティリティ ─────────────────────────────────
function _clearSlide(slide) {
  slide.getPageElements().forEach(e => e.remove());
}

function _header(slide, title, company, C, W) {
  _clearSlide(slide);
  slide.getBackground().setSolidFill(C.WHITE);
  _rect(slide, 0, 0, W, 58, C.PRIMARY);
  _rect(slide, 0, 58, W, 4, C.ACCENT);
  _text(slide, title,   18,    12, W - 210, 36, C.WHITE,   14, true);
  _text(slide, company, W - 195, 16, 175, 26, '#aaccdd',   8.5, false);
}

function _text(slide, str, x, y, w, h, color, size, bold) {
  const box = slide.insertTextBox(str.toString(), x, y, w, h);
  const ts  = box.getText().getTextStyle();
  ts.setFontFamily('Noto Sans JP').setFontSize(size).setForegroundColor(color).setBold(bold);
  return box;
}

function _rect(slide, x, y, w, h, color) {
  const s = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, w, h);
  s.getFill().setSolidFill(color);
  s.getBorder().setTransparent();
  return s;
}

function _getOrMakeFolder(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}
