/***************************************
 * Forecast Agent v1.0
 * 単一メーカー（1クライアント）用 / Google Sheets 実装
 *
 * v1.0（今回反映）
 * - 未確定月補完：月別（同月）トレンドで補完し、補完後に途中実績より下がらない
 * - 未確定月判定：実行日ベースで可変（当月以降=未確定 / 前月まで=確定）
 * - FACTORS/OPINIONS/DEV：必要情報が揃った行のみ計算に使用
 * - 入力異常検出：変な入力があれば実行前にエラー表示して停止
 * - OUTPUT：B=ネガ / C=中立 / D=ポジ、配色も統一（表＆グラフ）
 * - 実行中メッセージ：計算ステップが分かるtoastを追加（読み取り時間も確保）
 ***************************************/

const VERSION = '1.0';
const MENU_NAME = 'Forecast Agent';

const SHEETS = {
  GUIDE: 'GUIDE',
  CONFIG: 'CONFIG',
  SALES: 'SALES',
  FACTORS_PRODUCT: 'FACTORS_PRODUCT',
  FACTORS_CLIENT: 'FACTORS_CLIENT',
  OPINIONS: 'OPINIONS',
  DEV: 'DEV',
  OUTPUT: 'OUTPUT'
};

// 入力セル背景
const COLOR_OBJECTIVE = '#fff2cc'; // 黄色（客観）
const COLOR_SUBJECTIVE = '#cfe2f3'; // 青色（主観）
const COLOR_MIX_LABEL = '#f4cccc'; // 混合ラベル薄赤
const COLOR_OBJ_LABEL = '#cfe2f3'; // 客観ラベル薄青
const COLOR_HEADER = '#eeeeee';
const COLOR_P50_HILITE = '#fff2cc'; // P50強調（薄黄）

// OUTPUTの意味色（表＆グラフ）
const COLOR_NEG = '#f4cccc'; // ネガ薄赤
const COLOR_NEU = '#fff2cc'; // 中立薄黄
const COLOR_POS = '#cfe2f3'; // ポジ薄青
const COLOR_REG = '#5f6368'; // 回帰グレー

const TZ = Session.getScriptTimeZone();

// 外部「実績」集計元スプレッドシート
const EXTERNAL_SS_ID = '1qIAb_y3EhM6uiQrtT5hKCjUDHs3ARYBKdr-aCx0OY0c';
const EXTERNAL_SHEET_PREFIX = '*';
const EXTERNAL_SHEET_SUFFIX = '_actual_value';

// 外部シート列（1-index）
const EXT_COL_CLIENT = 41;        // AO
const EXT_COL_CATEGORY = 50;      // AX（製品名）
const EXT_COL_DATE_PRIMARY = 57;  // BE
const EXT_COL_DATE_SECONDARY = 56;// BD
const EXT_COL_AMOUNT = 66;        // BN

// Monte Carlo
const N_SIM = 1000;

// スパイク（単発外れ）をならすための上限/下限（比率）
const SPIKE_CLIP_MIN = 0.70;
const SPIKE_CLIP_MAX = 1.40;

// 「季節性を潰さない」ための：同月の分布で許容する広さ（MAD倍率）
const SEASONAL_MAD_K = 3.0;

// 未確定月補完に使う係数のクリップ（極端な補完を避ける）
const TREND_FACTOR_MIN = 0.85;
const TREND_FACTOR_MAX = 1.15;

// シートの共通列幅（見切れ防止）
const COL_WIDTHS = {
  W_PERSON: 120,
  W_PRODUCT: 220,
  W_MONTH: 150,
  W_STEP: 160,
  W_CONF: 170,
  W_TEXT: 360,
  W_MONEY: 150
};

// チャートの高さ相当の“余白行”目安（重なり防止）
const CHART_HEIGHT_ROWS = 22;

/** ====== メニュー ====== */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu(MENU_NAME)
    .addItem('① 初期セットアップ', 'setupForecastBook')
    .addItem('② 過去売上データを反映', 'importPastSalesToSalesTab')
    .addItem('③ 製品動向を入力', 'openProductTrendEntryDialog')
    .addItem('④ クライアント動向を入力', 'openClientTrendEntryDialog')
    .addItem('⑤ メーカー担当者意見を入力', 'openOpinionsEntryDialog')
    .addItem('⑥ スポット開発の見込みを入力', 'openDevEntryDialog')
    .addItem('⑦ 予測を出力（再実行した場合は最新情報で上書き）', 'executeForecastUsingConfig')
    .addToUi();
}

/**
 * 【管理者用】GUIDEだけを作成/更新し、GUIDE以外のタブを削除します。
 * - ユーザ配布前に、管理者が1回だけ実行する想定
 * - メニューには出しません（誤操作防止）
 */
function adminSetupGuideOnly() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert(
    '管理者用：GUIDEのみ作成',
    'GUIDEシートを作成/更新し、GUIDE以外のタブシートはすべて削除します。\n※削除したシートは元に戻せません。\n続行しますか？',
    ui.ButtonSet.OK_CANCEL
  );
  if (res !== ui.Button.OK) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  getOrCreateSheet_(ss, SHEETS.GUIDE);
  buildGUIDE_();

  const guide = ss.getSheetByName(SHEETS.GUIDE);
  ss.setActiveSheet(guide);

  ss.getSheets().forEach(sh => {
    if (sh.getName() !== SHEETS.GUIDE) {
      ss.deleteSheet(sh);
    }
  });

  ui.alert('完了', 'GUIDEシートを作成し、他のタブシートを削除しました。', ui.ButtonSet.OK);
}

/**
 * Step列の表示ゆらぎ対策：
 * - ユーザが「10%」「0.1」「-0.3」「+10」などで入力しても
 *   常に「+10%」「-30%」のような表示に正規化する（右寄せ）
 */
function onEdit(e) {
  try {
    const r = e.range;
    const sh = r.getSheet();
    const name = sh.getName();
    const row = r.getRow();
    const col = r.getColumn();
    if (row < 2) return;

    const isStepCell =
      (name === SHEETS.FACTORS_PRODUCT && col === 4) ||
      (name === SHEETS.FACTORS_CLIENT && col === 3) ||
      (name === SHEETS.OPINIONS && col === 3);

    if (!isStepCell) return;

    const v = r.getValue();
    const norm = normalizeStepDisplay_(v);
    if (norm === null) return;

    r.setNumberFormat('@');
    r.setHorizontalAlignment('right');
    r.setValue(norm);
  } catch (err) {
    // noop
  }
}

/** ====== ① 初期セットアップ ====== */
function setupForecastBook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const order = [
    SHEETS.GUIDE,
    SHEETS.CONFIG,
    SHEETS.SALES,
    SHEETS.FACTORS_PRODUCT,
    SHEETS.FACTORS_CLIENT,
    SHEETS.OPINIONS,
    SHEETS.DEV,
    SHEETS.OUTPUT
  ];

  order.forEach((name, idx) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name, idx);
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(idx + 1);
  });

  buildGUIDE_();
  buildCONFIG_();
  buildSALES_();
  buildFACTORS_PRODUCT_();
  buildFACTORS_CLIENT_();
  buildOPINIONS_();
  buildDEV_();
  buildOUTPUT_();

  showInitialSetupDialog_();
}

/** 初期設定ダイアログ（メーカー選択＋予測年度＋担当者） */
function showInitialSetupDialog_() {
  const ui = SpreadsheetApp.getUi();

  const defaultFY = getDefaultFY_();
  const clients = getClientCandidatesForSetup_();

  const esc = s => escapeHtml_(s);
  const optionsHtml = clients.map(c => `<option value="${esc(c)}">${esc(c)}</option>`).join('');

  const html = `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: sans-serif; padding: 14px; }
    h2 { margin: 0 0 10px 0; font-size: 16px; }
    .hint { color: #666; font-size: 12px; margin-bottom: 10px; line-height: 1.5; }
    .block { margin: 12px 0; }
    label { display: block; font-weight: 700; margin-bottom: 6px; }
    select, input { width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box; }
    .grid { display: grid; grid-template-columns: 36px 1fr; gap: 8px; align-items: center; }
    .grid .num { text-align: right; color: #666; font-size: 12px; }
    .btns { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-top: 14px; }
    button { padding: 10px; border: none; border-radius: 4px; font-weight: 700; cursor: pointer; }
    .primary { background: #4CAF50; color: #fff; }
    .secondary { background: #ddd; }
    .status { margin-top: 10px; font-size: 12px; color: #666; }
  </style>
</head>
<body>
  <h2>初期設定</h2>

  <div class="block">
    <label>メーカー名を入力してください。</label>
    <select id="client">
      <option value="" disabled selected>メーカーを選択してください</option>
      ${optionsHtml}
    </select>
    <div class="hint">
      ※ 候補は外部実績シート（最新2年のAO列）から自動抽出しています。
    </div>
  </div>

  <div class="block">
    <label>何年度（FY）を予測しますか？</label>
    <input id="fy" type="number" />
    <div class="hint">※ 空欄の場合はデフォルト年度（${defaultFY}年）を使用します。（決算期：3月末）</div>
  </div>

  <div class="block">
    <label>担当者設定</label>
    <div class="hint">シミュレーションするメーカー担当者の苗字を入力<br>※原則として全員の意見を反映するためです</div>

    <div class="grid">
      ${Array.from({length:10}).map((_,i)=>`
        <div class="num">${i+1}.</div>
        <input id="p${i+1}" type="text" placeholder="例：赤木" />
      `).join('')}
    </div>
    <div class="hint">空欄は無視され、CONFIG!B10 にカンマ区切りで保存されます。</div>
  </div>

  <div class="btns">
    <button class="secondary" onclick="skip()">スキップ</button>
    <button class="primary" onclick="save()">決定</button>
  </div>

  <div class="status" id="status"></div>

<script>
function save(){
  const client = document.getElementById('client').value;
  let fy = document.getElementById('fy').value;
  if(!fy) fy = '${defaultFY}';
  fy = String(fy).trim();

  const people = [];
  for(let i=1;i<=10;i++){
    const v = document.getElementById('p'+i).value;
    if(v && v.trim()) people.push(v.trim());
  }
  const peopleCSV = people.join(',');

  if(!client){
    alert('メーカーを選択してください。');
    return;
  }

  document.getElementById('status').textContent = '反映中…';

  google.script.run
    .withSuccessHandler(function(){
      google.script.host.close();
    })
    .withFailureHandler(function(e){
      document.getElementById('status').textContent = '';
      alert('エラー: ' + e.message);
    })
    .saveInitialSetupSettings(client, fy, peopleCSV);
}

function skip(){
  google.script.host.close();
}
</script>

</body>
</html>`;

  ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(420).setHeight(620), '初期設定');
}

/** 初期設定をCONFIGへ保存 */
function saveInitialSetupSettings(clientName, fyStr, peopleCSV) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getOrCreateSheet_(ss, SHEETS.CONFIG);

  const fy = Number(fyStr);
  if (!fy || !isFinite(fy)) throw new Error('予測年度(FY)が不正です。');

  cfg.getRange('B2').setValue(String(clientName || '').trim());
  cfg.getRange('B3').setValue(fy);
  cfg.getRange('B10').setValue(String(peopleCSV || '').trim());

  // GUIDE更新（更新履歴は保持されます）
  buildGUIDE_();
  ss.setActiveSheet(cfg);
}

/**
 * デフォルトFY：
 * - 現在が 1〜3月なら「今年」(例: 2026/02 → FY2026)
 * - 現在が 4〜12月なら「来年」(例: 2026/10 → FY2027)
 */
function getDefaultFY_() {
  const now = new Date();
  const y = now.getFullYear();
  const m = now.getMonth() + 1;
  return (m >= 4) ? (y + 1) : y;
}

/** 外部SSからメーカー候補（最新2年のAO列）を取得 */
function getClientCandidatesForSetup_() {
  const ext = SpreadsheetApp.openById(EXTERNAL_SS_ID);
  const sheets = ext.getSheets().map(s => s.getName());

  const yearTabs = [];
  sheets.forEach(name => {
    const m = name.match(/^\*(\d{4})_actual_value$/);
    if (m) yearTabs.push({ name, year: Number(m[1]) });
  });
  yearTabs.sort((a,b)=>b.year-a.year);

  const target = yearTabs.slice(0,2).map(o=>o.name);
  const set = new Set();

  target.forEach(tabName => {
    const sh = ext.getSheetByName(tabName);
    if (!sh) return;

    const maxCols = sh.getMaxColumns();
    if (maxCols < EXT_COL_CLIENT) return;

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;

    const vals = sh.getRange(2, EXT_COL_CLIENT, lastRow - 1, 1).getValues();
    vals.forEach(r => {
      const v = r[0];
      if (v && String(v).trim()) set.add(String(v).trim());
    });
  });

  return Array.from(set).sort();
}

/** ====== ② 過去売上データを反映（外部SS→このSSのSALES） ====== */
function importPastSalesToSalesTab() {
  ensureSetupDone_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName(SHEETS.CONFIG);

  const client = String(cfg.getRange('B2').getValue() || '').trim();
  const fy = Number(cfg.getRange('B3').getValue());

  if (!client) {
    SpreadsheetApp.getUi().alert('CONFIG!B2 にメーカー名が未設定です。①初期セットアップを実行してください。');
    return;
  }
  if (!fy || !isFinite(fy)) {
    SpreadsheetApp.getUi().alert('CONFIG!B3 に予測FYが未設定です。①初期セットアップを実行してください。');
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const res = ui.alert(
    '過去売上データを取り込みます',
    `メーカー: ${client}\n予測FY: ${fy}\n\n外部実績シートから過去4年分（48ヶ月）を集計して SALES に反映します。実行しますか？`,
    ui.ButtonSet.OK_CANCEL
  );
  if (res !== ui.Button.OK) return;

  toastProgress_(ss, `STEP: 外部実績を集計（過去4年/48ヶ月）…`);

  const ext = SpreadsheetApp.openById(EXTERNAL_SS_ID);

  // 予測FY=2026なら → 2022,2023,2024,2025
  const years = [fy - 4, fy - 3, fy - 2, fy - 1];
  const tabNames = years.map(y => `${EXTERNAL_SHEET_PREFIX}${y}${EXTERNAL_SHEET_SUFFIX}`);

  const start = new Date(fy - 4, 3, 1); // fy-4/04/01
  const totalMonths = 48;

  const map = new Map(); // productName -> monthly[48]
  let anyTabFound = false;

  tabNames.forEach(tab => {
    const sh = ext.getSheetByName(tab);
    if (!sh) return;
    anyTabFound = true;

    const maxCols = sh.getMaxColumns();
    const need = Math.max(EXT_COL_CLIENT, EXT_COL_CATEGORY, EXT_COL_AMOUNT, EXT_COL_DATE_PRIMARY, EXT_COL_DATE_SECONDARY);
    if (maxCols < need) return;

    const values = sh.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      const row = values[i];

      const c = row[EXT_COL_CLIENT - 1];
      if (String(c).trim() !== client) continue;

      const productName = row[EXT_COL_CATEGORY - 1]; // AX＝製品名
      const amount = row[EXT_COL_AMOUNT - 1];
      if (!productName) continue;
      if (amount === '' || amount === null || isNaN(amount)) continue;

      let dateVal = row[EXT_COL_DATE_PRIMARY - 1];
      if (!dateVal) dateVal = row[EXT_COL_DATE_SECONDARY - 1];

      const dt = toDate_(dateVal);
      if (!dt) continue;

      const idx = monthIndexFromStart_(dt, start);
      if (idx < 0 || idx >= totalMonths) continue;

      const key = String(productName).trim();
      if (!map.has(key)) map.set(key, new Array(totalMonths).fill(0));
      map.get(key)[idx] += Number(amount);
    }
  });

  if (!anyTabFound) {
    ui.alert('外部実績シート側に該当年度のタブが見つかりませんでした。\nタブ名（*YYYY_actual_value）を確認してください。');
    return;
  }
  if (map.size === 0) {
    ui.alert('該当データが見つかりませんでした。\nメーカー名や外部シートの列/タブ名を確認してください。');
    return;
  }

  const sales = ss.getSheetByName(SHEETS.SALES);
  buildSALES_(); // 列数確保

  const headerMonths = [];
  for (let i = 0; i < totalMonths; i++) {
    const d = addMonths_(start, i);
    headerMonths.push(fmtYM_(d)); // yyyy/MM
  }

  sales.getRange(1, 1).setValue('ProductName');
  sales.getRange(1, 2).setValue('(reserved)');
  sales.getRange(1, 3, 1, totalMonths).setValues([headerMonths]);

  const productNames = Array.from(map.keys()).sort();
  const out = productNames.map(name => [name, '', ...map.get(name)]);
  sales.getRange(2, 1, out.length, 2 + totalMonths).setValues(out);

  // 客観（黄色）
  sales.getRange(2, 3, out.length, totalMonths).setBackground(COLOR_OBJECTIVE);

  sales.setFrozenRows(1);
  sales.setFrozenColumns(2);
  sales.autoResizeColumns(1, 2);

  // 取り込み完了後にSALESを開く
  ss.setActiveSheet(sales);

  ui.alert('完了', `SALESに過去4年分（48ヶ月）の売上を反映しました。\nメーカー: ${client}`, ui.ButtonSet.OK);
}

/** ====== ③〜⑥：シート整形＋使い方案内（ポップアップは説明のみ） ====== */
function openProductTrendEntryDialog() {
  ensureSetupDone_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const people = getPeopleListFromConfig_();
  if (people.length === 0) {
    SpreadsheetApp.getUi().alert('CONFIG!B10 に担当者が設定されていません。\n①初期セットアップで担当者を入力してください。');
    return;
  }
  const products = getProductNameListFromSales_();
  if (products.length === 0) {
    SpreadsheetApp.getUi().alert('SALESに製品名がありません。\n②過去売上データを反映 を先に実行してください。');
    return;
  }

  const cfg = ss.getSheetByName(SHEETS.CONFIG);
  const fy = Number(cfg.getRange('B3').getValue()) || getDefaultFY_();
  const defaultDate = new Date(fy, 3, 1);

  const sh = ss.getSheetByName(SHEETS.FACTORS_PRODUCT);
  ensureFactorsProductTemplate_(sh, products, people, defaultDate);

  ss.setActiveSheet(sh);

  showInfoDialog_(
    '③ 製品動向を入力',
    [
      'FACTORS_PRODUCT を入力してください（青色のセルが対象です）。',
      '1) A列：担当者を選択',
      '2) C列：影響が出る日付（この日付以降に反映）',
      '3) D列：増減率（例：-30% = 今後30%減りそう）',
      '4) E列：根拠を短く',
      '',
      '※ B列の製品名はSALESから自動で入っています。',
      '※ Stepは入力ゆらぎが出ないよう自動で「+10%/-30%」形式に整えます。'
    ]
  );
}

function openClientTrendEntryDialog() {
  ensureSetupDone_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const people = getPeopleListFromConfig_();
  if (people.length === 0) {
    SpreadsheetApp.getUi().alert('CONFIG!B10 に担当者が設定されていません。\n①初期セットアップで担当者を入力してください。');
    return;
  }

  const cfg = ss.getSheetByName(SHEETS.CONFIG);
  const fy = Number(cfg.getRange('B3').getValue()) || getDefaultFY_();
  const defaultDate = new Date(fy, 3, 1);

  const sh = ss.getSheetByName(SHEETS.FACTORS_CLIENT);
  ensureFactorsClientTemplate_(sh, people, defaultDate);

  ss.setActiveSheet(sh);

  showInfoDialog_(
    '④ クライアント動向を入力',
    [
      'FACTORS_CLIENT を入力してください（青色のセルが対象です）。',
      '1) A列：担当者を選択',
      '2) B列：影響が出る日付（この日付以降に反映）',
      '3) C列：増減率（例：-10% = 予算圧縮で10%減りそう）',
      '4) D列：根拠を短く',
      '',
      '※ Stepは入力ゆらぎが出ないよう自動で「+10%/-30%」形式に整えます。'
    ]
  );
}

function openOpinionsEntryDialog() {
  ensureSetupDone_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const people = getPeopleListFromConfig_();
  if (people.length === 0) {
    SpreadsheetApp.getUi().alert('CONFIG!B10 に担当者が設定されていません。\n①初期セットアップで担当者を入力してください。');
    return;
  }

  const cfg = ss.getSheetByName(SHEETS.CONFIG);
  const fy = Number(cfg.getRange('B3').getValue()) || getDefaultFY_();
  const defaultDate = new Date(fy, 3, 1);

  const sh = ss.getSheetByName(SHEETS.OPINIONS);
  ensureOpinionsTemplate_(sh, people, defaultDate);

  ss.setActiveSheet(sh);

  showInfoDialog_(
    '⑤ メーカー担当者意見を入力',
    [
      'OPINIONS を入力してください（青色のセルが対象です）。',
      '※原則として担当者全員の入力が必要です（未入力があると⑦が実行できません）。',
      '',
      '入力手順：',
      '1) B列：影響が出る日付（この日付以降に反映）',
      '2) C列：増減率（例：+20% = 今後20%増えそう）',
      '3) D列：信頼度（0..1）',
      '4) E列：所感を短く',
      '',
      '※ 意見はそのまま固定反映されず、シミュレーション内でランダムに活用されます。'
    ]
  );
}

function openDevEntryDialog() {
  ensureSetupDone_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const people = getPeopleListFromConfig_();
  if (people.length === 0) {
    SpreadsheetApp.getUi().alert('CONFIG!B10 に担当者が設定されていません。\n①初期セットアップで担当者を入力してください。');
    return;
  }

  const cfg = ss.getSheetByName(SHEETS.CONFIG);
  const fy = Number(cfg.getRange('B3').getValue()) || getDefaultFY_();
  const defaultDate = new Date(fy, 3, 1);

  const sh = ss.getSheetByName(SHEETS.DEV);
  ensureDevTemplate_(sh, people, defaultDate);

  ss.setActiveSheet(sh);

  showInfoDialog_(
    '⑥ スポット開発（スポットイベント）を入力',
    [
      'DEV を入力してください（青色のセルが対象です）。',
      'スポット開発案件だけでなく、スポットイベント（例：法改定による差し替え等）もここに入力してください。',
      '',
      '入力手順：',
      '1) A列：担当者を選択',
      '2) B列：売上が立つ日付（この日付の月に反映）',
      '3) C列：案件名/イベント名',
      '4) D列：金額（円）',
      '5) E列：確度（0..1）',
      '',
      '※ DEVは「金額×確度」で固定加算されます（運用のシミュレーションには混ぜません）。'
    ]
  );
}

/** 説明だけの統一ポップアップ（キャンセル左／決定右） */
function showInfoDialog_(title, lines) {
  const ui = SpreadsheetApp.getUi();
  const esc = s => escapeHtml_(s);
  const body = lines.map(l => esc(l)).join('<br>');

  const html = `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: sans-serif; padding: 14px; }
    h2 { margin: 0 0 10px 0; font-size: 16px; }
    .box { color:#333; font-size: 12.5px; line-height:1.6; background:#fafafa; border:1px solid #ddd; border-radius:6px; padding: 10px; }
    .btns { display:grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-top: 14px; }
    button { padding:10px; border:none; border-radius:4px; font-weight:700; cursor:pointer; }
    .primary { background:#4CAF50; color:#fff; }
    .secondary { background:#ddd; }
  </style>
</head>
<body>
  <h2>${esc(title)}</h2>
  <div class="box">${body}</div>
  <div class="btns">
    <button class="secondary" onclick="closeIt()">キャンセル</button>
    <button class="primary" onclick="closeIt()">決定</button>
  </div>
<script>
function closeIt(){ google.script.host.close(); }
</script>
</body>
</html>`;
  ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(520).setHeight(360), title);
}

/** ====== ⑦ 予測を出力 ====== */
function executeForecastUsingConfig() {
  ensureSetupDone_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName(SHEETS.CONFIG);

  const client = String(cfg.getRange('B2').getValue() || '').trim();
  const fy = Number(cfg.getRange('B3').getValue());
  if (!client || !fy) {
    SpreadsheetApp.getUi().alert('初期設定が未完了です。①初期セットアップを実行してください。');
    return;
  }

  // 入力異常検出（おかしなデータがあれば止める）
  try {
    validateAllInputsOrThrow_(fy);
  } catch (e) {
    SpreadsheetApp.getUi().alert('入力エラー', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // 全員の意見があるか（必須：有効行があるか）
  const requiredPeople = getPeopleListFromConfig_();
  const missingPeople = findMissingPeopleOpinionsByValidRows_(requiredPeople);
  if (missingPeople.length > 0) {
    SpreadsheetApp.getUi().alert(
      '担当者の意見が不足しています',
      `OPINIONSに全員の意見が必要です。\n未入力: ${missingPeople.join(', ')}\n\n⑤で全員分入力してください。`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const res = ui.alert(
    'シミュレーションを実施しますか？',
    '※すでにシミュレーション実施している場合は上書きされます',
    ui.ButtonSet.OK_CANCEL
  );
  if (res !== ui.Button.OK) return;

  toastProgress_(ss, 'STEP1/6: SALES合算 → 未確定月を月別トレンドで補完（補完後に下がらない）…', 7);
  const result = runForecastFYCore_(fy, client);

  toastProgress_(ss, 'STEP6/6: OUTPUTへ書き出し（表＋グラフ）…', 6);
  writeOutputFY_(result);

  ss.toast('完了：OUTPUTを更新しました', MENU_NAME, 5);
  ss.setActiveSheet(ss.getSheetByName(SHEETS.OUTPUT));
}

/** ====== 予測コア ====== */
function runForecastFYCore_(fy, clientName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sales = ss.getSheetByName(SHEETS.SALES);
  if (!sales) throw new Error('SALESがありません。');

  const salesData = readSales48Months_(sales);

  if (!salesData.isComplete48) {
    const ui = SpreadsheetApp.getUi();
    const res = ui.alert(
      '注意：売上データが48ヶ月揃っていません',
      '48ヶ月（過去4年）揃っていない場合、予測精度が下がる可能性があります。\nこのままシミュレーションを続行しますか？',
      ui.ButtonSet.OK_CANCEL
    );
    if (res !== ui.Button.OK) throw new Error('ユーザーが中断しました。');
  }

  // メーカー合算（48ヶ月）
  const aggY_raw = sumAcrossProducts_(salesData.monthlyByProduct);

  // 未確定月補完（当月以降は未確定扱い／月別トレンド／補完後に途中実績より下がらない）
  const seriesStart = new Date(fy - 4, 3, 1); // fy-4/04/01
  const adj = adjustForUnclosedMonths_(aggY_raw, seriesStart);
  const aggY_adj = adj.series;

  toastProgress_(ss, 'STEP2/6: スパイクをならし（季節性は維持）→ トレンド＋季節性を推定…', 7);

  // スムージング（季節性は守りつつ単発スパイクだけ弱める）
  const smoothY = smoothSeriesSeasonalAware_(aggY_adj);

  // Opsモデル：トレンド＋季節性
  const model = fitOpsModelTrendSeason_(smoothY);

  // 残差%は「確定月のみ」から作る（未確定月の途中実績に依存しにくく）
  const residualPctClosed = [];
  for (let i = 0; i < smoothY.length; i++) {
    const mStart = addMonths_(seriesStart, i);
    if (mStart > adj.lastClosedMonthStart) continue; // 未確定は除外
    const f = model.fitted[i];
    if (f > 0) residualPctClosed.push(smoothY[i] / f - 1);
  }
  const residualPct = residualPctClosed.length ? residualPctClosed : smoothY.map((y, i) => (model.fitted[i] ? (y / model.fitted[i] - 1) : 0));

  const residP10 = percentile_(residualPct, 0.10);
  const residP50 = percentile_(residualPct, 0.50);
  const residP90 = percentile_(residualPct, 0.90);

  // Dev：固定加算（確度で調整）※運用シミュレーションには混ぜない
  const devFixedByMonth = readDevFixed12Months_(fy);

  // 要因（主観係数）※必要情報が揃った行だけ読む
  const factorsProduct = readFactorsProduct_(fy);
  const factorsClient = readFactorsClient_(fy);
  const opinions = readOpinions_(fy);

  // 製品構成比：未確定月を避ける（直近の“確定済み12ヶ月”で重み計算）
  const productWeights = computeProductWeightsClosed12_(
    salesData.productNames,
    salesData.monthlyByProduct,
    seriesStart,
    adj.lastClosedMonthStart
  );

  // 12ヶ月予測対象
  const months = [];
  const start = new Date(fy, 3, 1);
  for (let i = 0; i < 12; i++) months.push(addMonths_(start, i));

  // 線形回帰（参考）予測：季節性込みモデルのトレンド外挿（参考）
  const regTotal = [];
  for (let i = 0; i < 12; i++) {
    const t = 48 + (i + 1);
    const regOps = Math.max(0, (model.intercept + model.slope * t) * model.seasonalIndex[i % 12]);
    regTotal.push(regOps + devFixedByMonth[i]);
  }

  // 「客観のみ」：残差分位点レンジ + Dev固定
  const objOnly = forecastByResidualQuantiles_(model, devFixedByMonth, { p10: residP10, p50: residP50, p90: residP90 });

  toastProgress_(ss, `STEP3/6: 残差からレンジの基礎（P10/P50/P90）を作成…`, 5);
  toastProgress_(ss, `STEP4/6: Dev固定加算 + 主観係数（製品/クライアント/意見）を準備…`, 6);

  toastProgress_(ss, `STEP5/6: Monte Carlo ${N_SIM}回（運用のみを揺らす + Dev固定加算）…`, 8);

  // 「混合」シミュレーション（Opsのみ揺らす）＋係数適用＋Dev固定
  const mixed = forecastMonteCarloMixed_(model, devFixedByMonth, {
    residualPct,
    factorsProduct,
    factorsClient,
    opinions,
    productWeights,
    nSim: N_SIM,
    months
  });

  const opinionsSummaryTop = summarizeOpinionsTop_(opinions);
  const opinionsSummaryByMonth = summarizeOpinionsByMonth_(opinions, months);

  return {
    fy,
    clientName,
    months,
    objOnly,
    mixed,
    regTotal,
    devFixedByMonth,
    opinionsSummaryTop,
    opinionsSummaryByMonth,
    modelInfo: { residP10, residP50, residP90, slope: model.slope, intercept: model.intercept }
  };
}

/** ====== OUTPUT書き込み ====== */
function writeOutputFY_(result) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.OUTPUT);
  if (!sh) throw new Error('OUTPUTがありません。');

  sh.clear({ contentsOnly: true });
  sh.clearFormats();

  // 幅
  sh.setColumnWidth(1, 240);
  sh.setColumnWidth(2, 170);
  sh.setColumnWidth(3, 170);
  sh.setColumnWidth(4, 170);
  sh.setColumnWidth(5, 200);
  sh.setColumnWidth(6, 190);
  for (let c = 7; c <= 12; c++) sh.setColumnWidth(c, 130);

  const fy = result.fy;
  const client = result.clientName;

  const start = new Date(fy, 3, 1);
  const end = new Date(fy + 1, 2, 1);

  sh.getRange(1, 1).setValue(`FY${fy} 売上予測（${client} / ${fmtYM_(start)} 〜 ${fmtYM_(end)}）`);
  sh.getRange(1, 1, 1, 6).merge();
  sh.getRange(1, 1).setFontSize(16).setFontWeight('bold');
  sh.setFrozenRows(2);

  // 上部：担当者所感要約
  sh.getRange(3, 1).setValue('担当者所感（OPINION）');
  sh.getRange(3, 1).setFontWeight('bold');
  sh.getRange(3, 2).setValue(result.opinionsSummaryTop || '（未入力）');
  sh.getRange(3, 2, 1, 5).merge();
  sh.getRange(3, 2).setWrap(true);

  // 既存チャート削除（重なり防止）
  sh.getCharts().forEach(c => sh.removeChart(c));

  let row = 5;

  // ===== セクション1：混合 =====
  row = writeSectionBlock_(sh, row, {
    label: '過去売上（客観）と担当者情報（主観）を混合させたシミュレーション予測',
    labelBg: COLOR_MIX_LABEL,
    months: result.months,
    series: result.mixed,
    regTotal: result.regTotal,
    chartTitle: `混合：FY${fy} 月次予測レンジ（${client} / P10-P50-P90 + 回帰）`
  });

  row += 2;

  // ===== セクション2：客観のみ =====
  row = writeSectionBlock_(sh, row, {
    label: '過去売上のみ（客観）によるシミュレーション予測',
    labelBg: COLOR_OBJ_LABEL,
    months: result.months,
    series: result.objOnly,
    regTotal: result.regTotal,
    chartTitle: `客観のみ：FY${fy} 月次予測レンジ（${client} / P10-P50-P90 + 回帰）`
  });

  row += 2;

  // 参考：内訳（P50比較）
  sh.getRange(row, 1).setValue('（参考）内訳とメモ（P50比較）');
  sh.getRange(row, 1).setFontWeight('bold');
  sh.getRange(row, 1, 1, 6).merge();
  row++;

  sh.getRange(row, 1).setValue('※「運用(Ops)」はトレンド＋季節性から推定し、レンジは残差/シミュレーションで作ります。「開発/イベント(Dev)」は固定額（確度で調整）を加算します。');
  sh.getRange(row, 1, 1, 6).merge();
  sh.getRange(row, 1).setFontColor('#666666').setFontSize(10);
  row++;

  const hdr = ['Month', '運用(Ops)P50（客観のみ）', '運用(Ops)P50（混合）', '開発/イベント（Dev固定）', 'Total P50（客観のみ）', 'Total P50（混合）', 'OPINIONS要約'];
  sh.getRange(row, 1, 1, hdr.length).setValues([hdr]).setBackground(COLOR_HEADER).setFontWeight('bold');
  row++;

  const dev = result.devFixedByMonth;
  const rows = result.months.map((m, i) => {
    const objP50 = result.objOnly.p50[i];
    const mixP50 = result.mixed.p50[i];
    const devVal = dev[i];
    const opsObj = Math.max(0, objP50 - devVal);
    const opsMix = Math.max(0, mixP50 - devVal);
    return [
      fmtYM_(m),
      opsObj,
      opsMix,
      devVal,
      objP50,
      mixP50,
      result.opinionsSummaryByMonth[i] || ''
    ];
  });

  sh.getRange(row, 1, rows.length, hdr.length).setValues(rows);
  sh.getRange(row, 2, rows.length, 5).setNumberFormat('¥#,##0');
  sh.getRange(row, 7, rows.length, 1).setWrap(true);
}

/** セクションブロック（表＋グラフ） */
function writeSectionBlock_(sh, startRow, opt) {
  let r = startRow;

  // ラベル
  sh.getRange(r, 1).setValue(opt.label);
  sh.getRange(r, 1, 1, 6).merge();
  sh.getRange(r, 1).setBackground(opt.labelBg).setFontWeight('bold');
  r++;

  // 年度合計（B=ネガ / C=中立 / D=ポジ）
  const sumPos = sumArr_(opt.series.p90);
  const sumNeu = sumArr_(opt.series.p50);
  const sumNeg = sumArr_(opt.series.p10);
  const sumReg = sumArr_(opt.regTotal);
  const sumRange = sumPos - sumNeg;

  const annualHdr = ['年度合計（シミュレーション予測）', 'ネガ(P10)', '中立(P50)', 'ポジ(P90)', '線形回帰（参考）', 'レンジ(P90-P10)'];
  const annualVal = ['年度合計（予測）', sumNeg, sumNeu, sumPos, sumReg, sumRange];

  sh.getRange(r, 1, 1, annualHdr.length).setValues([annualHdr]).setBackground(COLOR_HEADER).setFontWeight('bold');
  // BCDだけ意味色に
  sh.getRange(r, 2).setBackground(COLOR_NEG);
  sh.getRange(r, 3).setBackground(COLOR_NEU);
  sh.getRange(r, 4).setBackground(COLOR_POS);
  r++;

  sh.getRange(r, 1, 1, annualVal.length).setValues([annualVal]);
  sh.getRange(r, 2, 1, 5).setNumberFormat('¥#,##0');

  // 意味色（値行）
  sh.getRange(r, 2).setBackground(COLOR_NEG);
  sh.getRange(r, 3).setBackground(COLOR_NEU).setFontWeight('bold'); // 中立を強調
  sh.getRange(r, 4).setBackground(COLOR_POS);
  r++;

  // 月次表
  r++;
  const hdr = ['Month', 'ネガ(P10)', '中立(P50)', 'ポジ(P90)', '線形回帰（参考）', 'レンジ(P90-P10)'];
  sh.getRange(r, 1, 1, hdr.length).setValues([hdr]).setBackground(COLOR_HEADER).setFontWeight('bold');
  // BCDだけ意味色に
  sh.getRange(r, 2).setBackground(COLOR_NEG);
  sh.getRange(r, 3).setBackground(COLOR_NEU);
  sh.getRange(r, 4).setBackground(COLOR_POS);
  r++;

  const table = opt.months.map((m, i) => {
    const pos = opt.series.p90[i];
    const neu = opt.series.p50[i];
    const neg = opt.series.p10[i];
    const reg = opt.regTotal[i];
    return [fmtYM_(m), neg, neu, pos, reg, (pos - neg)];
  });

  sh.getRange(r, 1, table.length, hdr.length).setValues(table);
  sh.getRange(r, 2, table.length, 5).setNumberFormat('¥#,##0');
  sh.getRange(r, 1, table.length, 1).setNumberFormat('@');

  // 意味色（列全体）
  sh.getRange(r, 2, table.length, 1).setBackground(COLOR_NEG);
  sh.getRange(r, 3, table.length, 1).setBackground(COLOR_NEU).setFontWeight('bold'); // 中立強調
  sh.getRange(r, 4, table.length, 1).setBackground(COLOR_POS);

  // P10/P50/P90説明（Note）
  sh.getRange(r - 1, 2).setNote('【ネガ(P10)】\nシミュレーション結果の下位10%点（=10パーセンタイル）。\n下振れ側の目安です。');
  sh.getRange(r - 1, 3).setNote('【中立(P50)】\nシミュレーション結果の中央値（=50パーセンタイル）。\n最も参照すべき“中心”の目安です。');
  sh.getRange(r - 1, 4).setNote('【ポジ(P90)】\nシミュレーション結果の上位10%点（=90パーセンタイル）。\n上振れ側の目安です。');
  sh.getRange(r - 1, 5).setNote('【線形回帰（参考）】\n過去売上（ならした推移）に単純な直線を当てて将来を外挿した参考値です。\n季節性も考慮したトレンド外挿を行います。');
  sh.getRange(r - 1, 6).setNote('【レンジ(P90-P10)】\nポジ(P90)からネガ(P10)を引いた幅です。\n不確実性（どれくらいブレうるか）の大きさを表します。');

  // グラフ：Month + Neg + Neu + Pos + Reg（A〜E）
  const chartRange = sh.getRange(r - 1, 1, table.length + 1, 5);

  const chartRow = startRow + 1;
  const chartCol = 8; // H列開始

  // 凡例テキスト（邪魔にならない小さめ）
  sh.getRange(chartRow - 1, chartCol).setValue('凡例：赤=ネガ(P10) / 黄=中立(P50) / 青=ポジ(P90) / 灰=線形回帰（参考）')
    .setFontSize(10).setFontColor('#666666');
  sh.getRange(chartRow - 1, chartCol, 1, 6).merge();

  const chart = sh.newChart()
    .asLineChart()
    .addRange(chartRange)
    .setPosition(chartRow, chartCol, 0, 0)
    .setOption('title', opt.chartTitle)
    .setOption('legend', { position: 'right' })
    .setOption('curveType', 'none')
    .setOption('lineWidth', 2)
    .setOption('pointSize', 0)
    .setOption('hAxis', { slantedText: true, slantedTextAngle: 45, showTextEvery: 1 })
    .setOption('vAxis', { format: '¥#,##0' })
    // 色：ネガ=赤 / 中立=黄 / ポジ=青 / 回帰=灰
    .setOption('colors', ['#ea4335', '#fbbc04', '#1a73e8', COLOR_REG])
    .setOption('series', { 0:{ lineWidth:3 }, 1:{ lineWidth:4 }, 2:{ lineWidth:3 }, 3:{ lineWidth:3 } })
    .setOption('width', 820)
    .setOption('height', 340)
    .build();

  sh.insertChart(chart);

  // チャートが重ならないよう、次の開始行をチャート分だけ下に送る
  const tableBottom = r + table.length + 2;
  const chartBottom = chartRow + CHART_HEIGHT_ROWS;
  return Math.max(tableBottom, chartBottom);
}

/** ====== シート構築 ====== */
function buildGUIDE_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getOrCreateSheet_(ss, SHEETS.GUIDE);

  // 更新履歴（管理者追記）を退避（GUIDE再生成でも保持）
  const LOG_TITLE_ROW = 23;
  const LOG_HEADER_ROW = 24;
  const LOG_ENTRY_START = 25;
  const LOG_ENTRY_ROWS = 20;
  const LOG_COLS = 3;

  let existingLog = [];
  try {
    existingLog = sh.getRange(LOG_ENTRY_START, 1, LOG_ENTRY_ROWS, LOG_COLS).getValues();
  } catch (e) {
    existingLog = [];
  }

  sh.clear({ contentsOnly: true });
  sh.clearFormats();

  sh.setColumnWidth(1, 560);
  sh.setColumnWidth(2, 260);
  sh.setColumnWidth(3, 260);

  sh.getRange(1, 1).setValue(`売上予測ツール 使い方（v${VERSION}）`);
  sh.getRange(1, 1).setFontSize(16).setFontWeight('bold');

  sh.getRange(3, 1).setValue('最短の手順（順番に実行してください）');
  sh.getRange(3, 1).setFontWeight('bold');

  const steps = [
    ['①', '初期セットアップ', 'メーカー・予測FY・担当者を設定（CONFIGに保存）'],
    ['②', '過去売上データを反映', '外部実績から過去4年(48ヶ月)を集計してSALESへ反映'],
    ['③', '製品動向を入力', 'FACTORS_PRODUCTを整形 → シート上で直接入力'],
    ['④', 'クライアント動向を入力', 'FACTORS_CLIENTを整形 → シート上で直接入力'],
    ['⑤', 'メーカー担当者意見を入力', 'OPINIONSを整形（全員分行を作成）→ 直接入力（必須）'],
    ['⑥', 'スポット開発の見込みを入力', 'DEVを整形 → 直接入力（スポットイベントもここ）'],
    ['⑦', '予測を出力', '最新入力でOUTPUTを上書き生成']
  ];

  sh.getRange(4, 1, 1, 3).setValues([['順番', 'Forecast Agent メニュー', '目的']])
    .setBackground(COLOR_HEADER).setFontWeight('bold');
  sh.getRange(5, 1, steps.length, 3).setValues(steps);
  sh.getRange(5, 1, steps.length, 1).setHorizontalAlignment('center');

  sh.getRange(13, 1).setValue('入力の安心ポイント');
  sh.getRange(13, 1).setFontWeight('bold');

  sh.getRange(14, 1).setValue('皆さんの率直なご意見が必要です。\n特定の意見に偏らないよう調整されますので、正解かどうかや結果を気にせず、安心して入力してください。');
  sh.getRange(14, 1).setWrap(true);

  sh.getRange(17, 1).setValue('ポイント');
  sh.getRange(17, 1).setFontWeight('bold');

  const tips = [
    '・③④⑤の「Step(増減率%)」は “その日付以降にどれくらい増減しそうか” の目安です（例：-30% = 今後30%減）。',
    '・⑤の意見は、そのまま固定反映せずシミュレーション内で活用されます（安心して率直に）。',
    '・⑥DEVは「割合」ではなく「固定額で増える」もの（スポット開発/スポットイベント）を入れる場所です。',
    '・未確定月（当月以降）は途中実績の影響を受けないよう、月別トレンドで補完して学習します（補完後に下がりません）。'
  ];
  sh.getRange(18, 1, tips.length, 1).setValues(tips.map(x => [x]));

  // 更新履歴（管理者が追記）
  sh.getRange(LOG_TITLE_ROW, 1).setValue('更新履歴（管理者が追記）').setFontWeight('bold');
  sh.getRange(LOG_HEADER_ROW, 1, 1, LOG_COLS)
    .setValues([['日付', 'Version', '内容']])
    .setBackground(COLOR_HEADER)
    .setFontWeight('bold');

  const restore = (existingLog && existingLog.length)
    ? existingLog
    : Array.from({ length: LOG_ENTRY_ROWS }, () => ['', '', '']);

  sh.getRange(LOG_ENTRY_START, 1, LOG_ENTRY_ROWS, LOG_COLS).setValues(restore);
  sh.getRange(LOG_ENTRY_START, 3, LOG_ENTRY_ROWS, 1).setWrap(true);

  // 目立ちすぎないよう小さめ
  sh.getRange(LOG_TITLE_ROW, 1, 1 + 1 + LOG_ENTRY_ROWS, LOG_COLS).setFontSize(10);

  ss.setActiveSheet(sh);
  ss.moveActiveSheet(1);
}

function buildCONFIG_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getOrCreateSheet_(ss, SHEETS.CONFIG);
  sh.clear({ contentsOnly: true });
  sh.clearFormats();

  sh.setColumnWidth(1, 260);
  sh.setColumnWidth(2, 420);

  // 担当者行はA10/B10
  const rows = [
    ['項目', '値'],
    ['メーカー名（外部集計キー）', ''],
    ['予測年度FY（YYYY）', ''],
    ['（メモ）決算期', '3月末'],
    ['（固定）Monte Carlo試行回数', N_SIM],
    ['（固定）未確定月の扱い', '前月までを確定とみなし、当月以降は同月トレンドで補完して学習（補完後に途中実績より下がらない）'],
    ['（固定）スパイクならし下限比', SPIKE_CLIP_MIN],
    ['（固定）スパイクならし上限比', SPIKE_CLIP_MAX],
    ['（固定）季節性保護（MAD倍率）', SEASONAL_MAD_K],
    ['担当者（カンマ区切り）', '']
  ];

  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.getRange(1, 1, 1, 2).setBackground(COLOR_HEADER).setFontWeight('bold');

  sh.getRange('B2').setBackground(COLOR_OBJECTIVE);
  sh.getRange('B3').setBackground(COLOR_OBJECTIVE);
  sh.getRange('B10').setBackground(COLOR_SUBJECTIVE);

  sh.getRange('A2').setNote('外部実績シート（*YYYY_actual_value）のAO列にあるメーカー名と一致させます。');
  sh.getRange('A3').setNote('例：FY2026 は 2026/04/01〜2027/03/31 の12ヶ月です。');
  sh.getRange('A5').setNote('シミュレーションは1000回試行し、レンジ（P10/P50/P90）を出します。単純な一発計算より「ブレ幅」を扱えるのがメリットです。');
  sh.getRange('A10').setNote('シミュレーションに関与する担当者の苗字をカンマ区切りで記載します。⑤では全員分の意見が必須です。');
}

function buildSALES_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getOrCreateSheet_(ss, SHEETS.SALES);
  sh.clear({ contentsOnly: true });
  sh.clearFormats();

  // 48ヶ月分（C〜AX=48列）を扱うため列数を確保
  ensureSheetHasColumns_(sh, 2 + 48);

  sh.setColumnWidth(1, 240);
  sh.setColumnWidth(2, 120);
  for (let c = 3; c <= 50; c++) sh.setColumnWidth(c, 110);

  sh.getRange(1, 1).setValue('ProductName');
  sh.getRange(1, 2).setValue('(reserved)');
  sh.getRange(1, 1, 1, 2).setBackground(COLOR_HEADER).setFontWeight('bold');

  sh.setFrozenRows(1);
  sh.setFrozenColumns(2);

  sh.getRange(1, 1).setNote('外部実績から取り込まれた「製品名」です。');
  sh.getRange(1, 3).setNote('過去4年（48ヶ月）の月次売上（客観データ）です。');
}

function buildFACTORS_PRODUCT_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getOrCreateSheet_(ss, SHEETS.FACTORS_PRODUCT);
  sh.clear({ contentsOnly: true });
  sh.clearFormats();

  const header = ['Person', 'ProductName', 'Month(yyyy/mm/dd)', 'Step(増減率%)', 'Reason'];
  sh.getRange(1, 1, 1, header.length).setValues([header]).setBackground(COLOR_HEADER).setFontWeight('bold');

  sh.setColumnWidth(1, COL_WIDTHS.W_PERSON);
  sh.setColumnWidth(2, COL_WIDTHS.W_PRODUCT);
  sh.setColumnWidth(3, COL_WIDTHS.W_MONTH);
  sh.setColumnWidth(4, COL_WIDTHS.W_STEP);
  sh.setColumnWidth(5, COL_WIDTHS.W_TEXT);

  sh.getRange(1, 4).setHorizontalAlignment('right');
  sh.getRange('D:D').setNumberFormat('@').setHorizontalAlignment('right');

  sh.getRange(1, 1).setNote('入力者（苗字推奨）。A列はプルダウンで選択します。');
  sh.getRange(1, 2).setNote('SALESにあるProductNameと一致させます（③で自動展開）。');
  sh.getRange(1, 3).setNote('この日付「以降」に影響が出る想定で入力します。');
  sh.getRange(1, 4).setNote('増減率（%）です。例：-30% = 今後30%減りそう。\n※入力値はそのまま直に固定反映せず、シミュレーション内で扱われます。');
  sh.getRange(1, 5).setNote('根拠を短く（例：プロモ終了、競合参入、学会など）。');

  sh.setFrozenRows(1);
}

function buildFACTORS_CLIENT_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getOrCreateSheet_(ss, SHEETS.FACTORS_CLIENT);
  sh.clear({ contentsOnly: true });
  sh.clearFormats();

  const header = ['Person', 'Month(yyyy/mm/dd)', 'Step(増減率%)', 'Reason'];
  sh.getRange(1, 1, 1, header.length).setValues([header]).setBackground(COLOR_HEADER).setFontWeight('bold');

  sh.setColumnWidth(1, COL_WIDTHS.W_PERSON);
  sh.setColumnWidth(2, COL_WIDTHS.W_MONTH);
  sh.setColumnWidth(3, COL_WIDTHS.W_STEP);
  sh.setColumnWidth(4, COL_WIDTHS.W_TEXT);

  sh.getRange('C:C').setNumberFormat('@').setHorizontalAlignment('right');

  sh.getRange(1, 1).setNote('入力者（苗字推奨）。A列はプルダウンで選択します。');
  sh.getRange(1, 2).setNote('この日付「以降」に影響が出る想定で入力します。');
  sh.getRange(1, 3).setNote('増減率（%）です。例：-10% = 予算圧縮で10%減りそう。\n※入力値はそのまま直に固定反映せず、シミュレーション内で扱われます。');
  sh.getRange(1, 4).setNote('根拠を短く（例：予算圧縮、体制変更など）。');

  sh.setFrozenRows(1);
}

function buildOPINIONS_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getOrCreateSheet_(ss, SHEETS.OPINIONS);
  sh.clear({ contentsOnly: true });
  sh.clearFormats();

  const header = ['Person', 'Month(yyyy/mm/dd)', 'Step(増減率%)', 'Confidence(0..1)', 'Note'];
  sh.getRange(1, 1, 1, header.length).setValues([header]).setBackground(COLOR_HEADER).setFontWeight('bold');

  sh.setColumnWidth(1, COL_WIDTHS.W_PERSON);
  sh.setColumnWidth(2, COL_WIDTHS.W_MONTH);
  sh.setColumnWidth(3, COL_WIDTHS.W_STEP);
  sh.setColumnWidth(4, COL_WIDTHS.W_CONF);
  sh.setColumnWidth(5, COL_WIDTHS.W_TEXT);

  sh.getRange('C:C').setNumberFormat('@').setHorizontalAlignment('right');

  sh.getRange(1, 1).setNote('担当者（苗字推奨）。⑤で全員分の行を自動作成します。');
  sh.getRange(1, 2).setNote('この日付「以降」に意見の影響が出る想定で入力します。');
  sh.getRange(1, 3).setNote('増減率（%）です。例：+20% = 今後20%増えそう。\n※意見はそのまま固定反映されず、シミュレーションでランダムに活用されます。');
  sh.getRange(1, 4).setNote('信頼度（0..1）。1に近いほど「この意見を強く信用してよい」として影響が強まります。');
  sh.getRange(1, 5).setNote('所感を短く（例：プロモ減、資材整理、体制変更など）。');

  sh.getRange('D2:D').setDataValidation(SpreadsheetApp.newDataValidation().requireNumberBetween(0, 1).build());
  sh.setFrozenRows(1);
}

function buildDEV_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getOrCreateSheet_(ss, SHEETS.DEV);
  sh.clear({ contentsOnly: true });
  sh.clearFormats();

  const header = ['Person', 'Month(yyyy/mm/dd)', 'Project', 'Amount(JPY)', 'Confidence(0..1)'];
  sh.getRange(1, 1, 1, header.length).setValues([header]).setBackground(COLOR_HEADER).setFontWeight('bold');

  sh.setColumnWidth(1, COL_WIDTHS.W_PERSON);
  sh.setColumnWidth(2, COL_WIDTHS.W_MONTH);
  sh.setColumnWidth(3, 280);
  sh.setColumnWidth(4, COL_WIDTHS.W_MONEY);
  sh.setColumnWidth(5, COL_WIDTHS.W_CONF);

  sh.getRange('D:D').setNumberFormat('¥#,##0');
  sh.getRange(1, 1).setNote('入力者（苗字推奨）。A列はプルダウンで選択します。');
  sh.getRange(1, 2).setNote('この日付の月に固定売上として加算します（スポット開発/スポットイベント）。');
  sh.getRange(1, 3).setNote('案件名（またはスポットイベント名）を短く。');
  sh.getRange(1, 4).setNote('金額（円）。ここは運用(Ops)のシミュレーションには混ぜず、固定額として加算します。');
  sh.getRange(1, 5).setNote('確度（0..1）。金額×確度で固定加算されます（例：1,000,000円×0.9=900,000円）。');

  sh.getRange('E2:E').setDataValidation(SpreadsheetApp.newDataValidation().requireNumberBetween(0, 1).build());
  sh.setFrozenRows(1);
}

function buildOUTPUT_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getOrCreateSheet_(ss, SHEETS.OUTPUT);
  sh.clear({ contentsOnly: true });
  sh.clearFormats();
}

/** ====== テンプレ整形（③〜⑥で呼ぶ） ====== */
function ensureFactorsProductTemplate_(sh, products, people, defaultDate) {
  if (!sh) throw new Error('FACTORS_PRODUCTがありません。');

  const last = sh.getLastRow();
  const existing = new Set();
  if (last >= 2) {
    const vals = sh.getRange(2, 2, last - 1, 1).getValues();
    vals.forEach(r => {
      const v = String(r[0] || '').trim();
      if (v) existing.add(v);
    });
  }

  const toAdd = products.filter(p => !existing.has(p));
  if (toAdd.length > 0) {
    const startRow = sh.getLastRow() + 1;
    const rows = toAdd.map(p => ['', p, defaultDate, '0%', '']);
    sh.getRange(startRow, 1, rows.length, 5).setValues(rows);
  }

  const maxRow = Math.max(sh.getLastRow(), 2);
  sh.getRange(2, 1, maxRow - 1, 5).setBackground(COLOR_SUBJECTIVE);

  const dvPerson = SpreadsheetApp.newDataValidation()
    .requireValueInList(people, true)
    .setAllowInvalid(false)
    .build();
  sh.getRange(2, 1, sh.getMaxRows() - 1, 1).setDataValidation(dvPerson);

  sh.getRange('C2:C').setNumberFormat('yyyy/MM/dd');

  const stepList = buildPercentStepList_();
  const dvStep = SpreadsheetApp.newDataValidation()
    .requireValueInList(stepList, true)
    .setAllowInvalid(true)
    .build();
  sh.getRange(2, 4, sh.getMaxRows() - 1, 1).setDataValidation(dvStep);
  sh.getRange('D:D').setNumberFormat('@').setHorizontalAlignment('right');

  sh.setColumnWidth(1, COL_WIDTHS.W_PERSON);
  sh.setColumnWidth(2, COL_WIDTHS.W_PRODUCT);
  sh.setColumnWidth(3, COL_WIDTHS.W_MONTH);
  sh.setColumnWidth(4, COL_WIDTHS.W_STEP);
  sh.setColumnWidth(5, COL_WIDTHS.W_TEXT);
}

function ensureFactorsClientTemplate_(sh, people, defaultDate) {
  if (!sh) throw new Error('FACTORS_CLIENTがありません。');

  if (sh.getLastRow() < 2) {
    const rows = Array.from({ length: 10 }, () => ['', defaultDate, '0%', '']);
    sh.getRange(2, 1, rows.length, 4).setValues(rows);
  }

  const maxRow = Math.max(sh.getLastRow(), 2);
  sh.getRange(2, 1, maxRow - 1, 4).setBackground(COLOR_SUBJECTIVE);

  const dvPerson = SpreadsheetApp.newDataValidation()
    .requireValueInList(people, true)
    .setAllowInvalid(false)
    .build();
  sh.getRange(2, 1, sh.getMaxRows() - 1, 1).setDataValidation(dvPerson);

  const stepList = buildPercentStepList_();
  const dvStep = SpreadsheetApp.newDataValidation()
    .requireValueInList(stepList, true)
    .setAllowInvalid(true)
    .build();
  sh.getRange(2, 3, sh.getMaxRows() - 1, 1).setDataValidation(dvStep);
  sh.getRange('B2:B').setNumberFormat('yyyy/MM/dd');
  sh.getRange('C:C').setNumberFormat('@').setHorizontalAlignment('right');
}

function ensureOpinionsTemplate_(sh, people, defaultDate) {
  if (!sh) throw new Error('OPINIONSがありません。');

  const last = sh.getLastRow();
  const existing = new Set();
  if (last >= 2) {
    const vals = sh.getRange(2, 1, last - 1, 1).getValues();
    vals.forEach(r => {
      const v = String(r[0] || '').trim();
      if (v) existing.add(v);
    });
  }

  const missing = people.filter(p => !existing.has(p));
  if (missing.length > 0) {
    const startRow = sh.getLastRow() + 1;
    const rows = missing.map(p => [p, defaultDate, '0%', 0.70, '']);
    sh.getRange(startRow, 1, rows.length, 5).setValues(rows);
  }

  const maxRow = Math.max(sh.getLastRow(), 2);
  sh.getRange(2, 1, maxRow - 1, 5).setBackground(COLOR_SUBJECTIVE);

  const stepList = buildPercentStepList_();
  const dvStep = SpreadsheetApp.newDataValidation()
    .requireValueInList(stepList, true)
    .setAllowInvalid(true)
    .build();
  sh.getRange(2, 3, sh.getMaxRows() - 1, 1).setDataValidation(dvStep);
  sh.getRange('B2:B').setNumberFormat('yyyy/MM/dd');
  sh.getRange('C:C').setNumberFormat('@').setHorizontalAlignment('right');

  sh.getRange(2, 4, sh.getMaxRows() - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireNumberBetween(0, 1).setAllowInvalid(true).build()
  );
}

function ensureDevTemplate_(sh, people, defaultDate) {
  if (!sh) throw new Error('DEVがありません。');

  if (sh.getLastRow() < 2) {
    const rows = Array.from({ length: 10 }, () => ['', defaultDate, '', '', 1.0]);
    sh.getRange(2, 1, rows.length, 5).setValues(rows);
  }

  const maxRow = Math.max(sh.getLastRow(), 2);
  sh.getRange(2, 1, maxRow - 1, 5).setBackground(COLOR_SUBJECTIVE);

  const dvPerson = SpreadsheetApp.newDataValidation()
    .requireValueInList(people, true)
    .setAllowInvalid(false)
    .build();
  sh.getRange(2, 1, sh.getMaxRows() - 1, 1).setDataValidation(dvPerson);

  sh.getRange('B2:B').setNumberFormat('yyyy/MM/dd');
  sh.getRange('D:D').setNumberFormat('¥#,##0');

  sh.getRange(2, 5, sh.getMaxRows() - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireNumberBetween(0, 1).setAllowInvalid(true).build()
  );
}

/** ====== 未確定月補完（同月トレンド＋補完後に下がらない） ====== */
function adjustForUnclosedMonths_(y, seriesStart) {
  const lastClosed = getLastClosedMonthStart_(); // 前月まで確定
  const n = y.length;
  const out = y.slice();

  const closedIdx = [];
  for (let i = 0; i < n; i++) {
    const mStart = addMonths_(seriesStart, i);
    if (mStart <= lastClosed) closedIdx.push(i);
  }

  // 同月トレンド（前年同月比）の中央値を月別に算出（極端値はクリップ）
  const monthFactors = computeMonthTrendFactors_(out, closedIdx);

  for (let i = 0; i < n; i++) {
    const mStart = addMonths_(seriesStart, i);
    if (mStart <= lastClosed) continue; // 確定月はそのまま

    const m = i % 12;
    const current = Number(out[i] || 0);

    let base = current;

    if (i - 12 >= 0) {
      const prev = Number(out[i - 12] || 0);
      if (prev > 0) {
        base = prev * monthFactors[m];
      } else {
        base = estimateMonthAverage_(out, closedIdx, m);
      }
    } else {
      base = estimateMonthAverage_(out, closedIdx, m);
    }

    if (!isFinite(base)) base = current;
    base = Math.max(0, base);

    // ★重要：補完後に途中実績より下がらない
    if (base < current) base = current;

    out[i] = base;
  }

  return { series: out, lastClosedMonthStart: lastClosed, monthTrendFactors: monthFactors };
}

/** 実行日ベース：前月まで確定（当月以降は未確定） */
function getLastClosedMonthStart_() {
  const now = new Date();
  const firstThisMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  const prevMonth = new Date(firstThisMonth.getFullYear(), firstThisMonth.getMonth() - 1, 1);
  return prevMonth;
}

function computeMonthTrendFactors_(y, closedIdx) {
  const factors = new Array(12).fill(1);

  for (let m = 0; m < 12; m++) {
    const ratios = [];
    for (let k = 0; k < closedIdx.length; k++) {
      const i = closedIdx[k];
      if (i % 12 !== m) continue;
      if (i - 12 < 0) continue;

      const prev = Number(y[i - 12] || 0);
      const cur = Number(y[i] || 0);
      if (prev > 0 && cur > 0) {
        ratios.push(cur / prev);
      }
    }

    if (ratios.length > 0) {
      let med = percentile_(ratios, 0.50);
      if (!isFinite(med) || med <= 0) med = 1;
      // クリップ
      med = Math.max(TREND_FACTOR_MIN, Math.min(TREND_FACTOR_MAX, med));
      factors[m] = med;
    } else {
      factors[m] = 1;
    }
  }
  return factors;
}

function estimateMonthAverage_(y, closedIdx, monthMod) {
  const arr = [];
  for (let k = 0; k < closedIdx.length; k++) {
    const i = closedIdx[k];
    if (i % 12 !== monthMod) continue;
    const v = Number(y[i] || 0);
    if (isFinite(v) && v > 0) arr.push(v);
  }
  return arr.length ? avg_(arr) : 0;
}

/** ====== 入力異常検出（おかしなデータで止める） ====== */
function validateAllInputsOrThrow_(fy) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // CONFIG
  const cfg = ss.getSheetByName(SHEETS.CONFIG);
  const client = String(cfg.getRange('B2').getValue() || '').trim();
  const fyNum = Number(cfg.getRange('B3').getValue());
  const people = getPeopleListFromConfig_();
  if (!client) throw new Error('CONFIG!B2（メーカー名）が未入力です。');
  if (!isFinite(fyNum) || fyNum <= 2000) throw new Error('CONFIG!B3（予測年度FY）が不正です。');
  if (people.length === 0) throw new Error('CONFIG!B10（担当者）が未入力です。');

  // SALES（数値かどうか）
  const sales = ss.getSheetByName(SHEETS.SALES);
  if (!sales) throw new Error('SALESシートがありません。');

  const lastRow = sales.getLastRow();
  if (lastRow < 2) throw new Error('SALESに製品行がありません。②で取り込み、または手入力してください。');

  const expectedMonths = 48;
  const startCol = 3;
  const endCol = startCol + expectedMonths - 1;
  if (sales.getLastColumn() < endCol) {
    throw new Error('SALESの月次列が48ヶ月分ありません。②過去売上データを反映 を実行してください。');
  }

  const values = sales.getRange(2, 1, lastRow - 1, endCol).getValues();
  for (let r = 0; r < values.length; r++) {
    const pname = String(values[r][0] || '').trim();
    if (!pname) continue;

    for (let c = startCol - 1; c <= endCol - 1; c++) {
      const v = values[r][c];
      if (v === '' || v === null) continue;
      if (typeof v === 'number') {
        if (!isFinite(v)) throw new Error(`SALES: 数値が不正です（${pname} / col ${c + 1}）`);
      } else {
        const n = toNumberSafe_(v);
        if (!isFinite(n)) throw new Error(`SALES: 数値に変換できない値があります（${pname} / col ${c + 1} / "${v}"）`);
      }
    }
  }

  // FACTORS / OPINIONS / DEV：明らかに変な行があれば停止（未完成行は“無視”＝エラーにはしない）
  validateFactorsSheet_(SHEETS.FACTORS_PRODUCT, { cols: 5, mode: 'product' });
  validateFactorsSheet_(SHEETS.FACTORS_CLIENT, { cols: 4, mode: 'client' });
  validateOpinionsSheet_(people);
  validateDevSheet_();
}

function validateFactorsSheet_(sheetName, opt) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return;

  const last = sh.getLastRow();
  if (last < 2) return;

  const vals = sh.getRange(2, 1, last - 1, opt.cols).getValues();

  for (let i = 0; i < vals.length; i++) {
    const rowNum = i + 2;
    const row = vals[i];

    // 行が完全空なら無視
    if (row.every(v => v === '' || v === null)) continue;

    if (opt.mode === 'product') {
      const person = String(row[0] || '').trim();
      const product = String(row[1] || '').trim();
      const monthRaw = row[2];
      const stepRaw = row[3];

      // 未完成なら無視（エラーでは止めない）
      if (!person || !product || !monthRaw || stepRaw === '' || stepRaw === null) continue;

      const dt = toDate_(monthRaw);
      if (!dt) throw new Error(`${sheetName}!C${rowNum} の日付が不正です（yyyy/mm/dd 形式で入力してください）。`);

      const step = parseRate_(stepRaw);
      if (!isFinite(step)) throw new Error(`${sheetName}!D${rowNum} のStepが解釈できません（例：-30% や +10%）。`);

      if (Math.abs(step) > 5) throw new Error(`${sheetName}!D${rowNum} のStepが極端に大きいです（${stepRaw}）。意図した値か確認してください。`);
    }

    if (opt.mode === 'client') {
      const person = String(row[0] || '').trim();
      const monthRaw = row[1];
      const stepRaw = row[2];

      if (!person || !monthRaw || stepRaw === '' || stepRaw === null) continue;

      const dt = toDate_(monthRaw);
      if (!dt) throw new Error(`${sheetName}!B${rowNum} の日付が不正です（yyyy/mm/dd 形式で入力してください）。`);

      const step = parseRate_(stepRaw);
      if (!isFinite(step)) throw new Error(`${sheetName}!C${rowNum} のStepが解釈できません（例：-30% や +10%）。`);

      if (Math.abs(step) > 5) throw new Error(`${sheetName}!C${rowNum} のStepが極端に大きいです（${stepRaw}）。意図した値か確認してください。`);
    }
  }
}

function validateOpinionsSheet_(requiredPeople) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.OPINIONS);
  if (!sh) throw new Error('OPINIONSシートがありません。⑤を実行してください。');

  const last = sh.getLastRow();
  if (last < 2) throw new Error('OPINIONSに入力行がありません。⑤を実行してください。');

  const vals = sh.getRange(2, 1, last - 1, 5).getValues();

  // 有効行：Person + Month + Step + Confidence が揃っている
  const okPeople = new Set();

  for (let i = 0; i < vals.length; i++) {
    const rowNum = i + 2;
    const person = String(vals[i][0] || '').trim();
    const monthRaw = vals[i][1];
    const stepRaw = vals[i][2];
    const confRaw = vals[i][3];

    // 行が完全空なら無視
    if ([person, monthRaw, stepRaw, confRaw, vals[i][4]].every(v => v === '' || v === null)) continue;

    // 途中の未完成行は無視（ただし変な値はエラー）
    if (!person || !monthRaw || stepRaw === '' || stepRaw === null || confRaw === '' || confRaw === null) continue;

    const dt = toDate_(monthRaw);
    if (!dt) throw new Error(`OPINIONS!B${rowNum} の日付が不正です（yyyy/mm/dd 形式で入力してください）。`);

    const step = parseRate_(stepRaw);
    if (!isFinite(step)) throw new Error(`OPINIONS!C${rowNum} のStepが解釈できません（例：-30% や +10%）。`);

    const conf = Number(confRaw);
    if (!isFinite(conf) || conf < 0 || conf > 1) throw new Error(`OPINIONS!D${rowNum} の信頼度が不正です（0..1）。`);

    okPeople.add(person);
  }

  const missing = requiredPeople.filter(p => !okPeople.has(p));
  if (missing.length > 0) {
    throw new Error(`OPINIONSに担当者全員の有効な入力がありません。\n未入力: ${missing.join(', ')}\n⑤で入力してください。`);
  }
}

function validateDevSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.DEV);
  if (!sh) return;

  const last = sh.getLastRow();
  if (last < 2) return;

  const vals = sh.getRange(2, 1, last - 1, 5).getValues();
  for (let i = 0; i < vals.length; i++) {
    const rowNum = i + 2;
    const person = String(vals[i][0] || '').trim();
    const monthRaw = vals[i][1];
    const project = String(vals[i][2] || '').trim();
    const amountRaw = vals[i][3];
    const confRaw = vals[i][4];

    // 完全空行は無視
    if ([person, monthRaw, project, amountRaw, confRaw].every(v => v === '' || v === null)) continue;

    // 未完成行は無視（ただし変な値はエラー）
    if (!monthRaw || amountRaw === '' || amountRaw === null || confRaw === '' || confRaw === null) continue;

    const dt = toDate_(monthRaw);
    if (!dt) throw new Error(`DEV!B${rowNum} の日付が不正です（yyyy/mm/dd 形式で入力してください）。`);

    const amt = toNumberSafe_(amountRaw);
    if (!isFinite(amt)) throw new Error(`DEV!D${rowNum} の金額が数値として不正です（"${amountRaw}"）。`);
    if (amt < 0) throw new Error(`DEV!D${rowNum} の金額が負の値です（${amt}）。`);

    const conf = Number(confRaw);
    if (!isFinite(conf) || conf < 0 || conf > 1) throw new Error(`DEV!E${rowNum} の確度が不正です（0..1）。`);
  }
}

function toNumberSafe_(v) {
  if (v === null || v === undefined) return NaN;
  if (typeof v === 'number') return v;
  const s = String(v).trim();
  if (!s) return NaN;
  const norm = s.replace(/[,\s]/g, '').replace(/¥/g, '').replace(/￥/g, '');
  const n = Number(norm);
  return n;
}

/** ====== 読み取り関数 ====== */
function getPeopleListFromConfig_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName(SHEETS.CONFIG);
  if (!cfg) return [];
  const raw = String(cfg.getRange('B10').getValue() || '');
  return raw.split(',').map(s => s.trim()).filter(s => s.length > 0);
}

function getProductNameListFromSales_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sales = ss.getSheetByName(SHEETS.SALES);
  if (!sales) return [];
  const last = sales.getLastRow();
  if (last < 2) return [];
  const vals = sales.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0] || '').trim()).filter(Boolean);
  return Array.from(new Set(vals)).sort();
}

function findMissingPeopleOpinionsByValidRows_(requiredPeople) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const op = ss.getSheetByName(SHEETS.OPINIONS);
  if (!op) return requiredPeople;

  const last = op.getLastRow();
  if (last < 2) return requiredPeople;

  const vals = op.getRange(2, 1, last - 1, 5).getValues();

  const ok = new Set();
  vals.forEach(r => {
    const person = String(r[0] || '').trim();
    const monthRaw = r[1];
    const stepRaw = r[2];
    const confRaw = r[3];

    if (!person) return;
    if (!monthRaw) return;
    if (stepRaw === '' || stepRaw === null) return;
    if (confRaw === '' || confRaw === null) return;

    const dt = toDate_(monthRaw);
    const step = parseRate_(stepRaw);
    const conf = Number(confRaw);

    if (dt && isFinite(step) && isFinite(conf) && conf >= 0 && conf <= 1) ok.add(person);
  });

  return requiredPeople.filter(p => !ok.has(p));
}

/** SALES読み取り（48ヶ月） */
function readSales48Months_(salesSheet) {
  const lastRow = salesSheet.getLastRow();
  const lastCol = salesSheet.getLastColumn();

  const expectedMonths = 48;
  const startCol = 3; // C列〜
  const endCol = startCol + expectedMonths - 1; // 50（AX）

  const isComplete48 = (lastCol >= endCol);

  const productRows = Math.max(0, lastRow - 1);
  const data = [];

  if (productRows > 0) {
    const width = Math.min(lastCol, endCol);
    const vals = salesSheet.getRange(2, 1, productRows, width).getValues();
    vals.forEach(row => {
      const name = String(row[0] || '').trim();
      if (!name) return;
      const arr = new Array(expectedMonths).fill(0);
      for (let i = 0; i < expectedMonths; i++) {
        const idx = (startCol - 1) + i;
        const v = row[idx];
        if (typeof v === 'number') arr[i] = Number(v) || 0;
        else {
          const n = toNumberSafe_(v);
          arr[i] = isFinite(n) ? n : 0;
        }
      }
      data.push({ productName: name, monthly: arr });
    });
  }

  return {
    isComplete48,
    monthlyByProduct: data.map(x => x.monthly),
    productNames: data.map(x => x.productName)
  };
}

function sumAcrossProducts_(monthlyByProduct) {
  const n = 48;
  const out = new Array(n).fill(0);
  monthlyByProduct.forEach(arr => {
    for (let i = 0; i < n; i++) out[i] += Number(arr[i] || 0);
  });
  return out;
}

/** 製品構成比：直近の“確定済み12ヶ月”で計算 */
function computeProductWeightsClosed12_(productNames, monthlyByProduct, seriesStart, lastClosedMonthStart) {
  const map = new Map();
  if (!productNames || productNames.length === 0) return map;

  let closedEndIdx = -1;
  for (let i = 0; i < 48; i++) {
    const mStart = addMonths_(seriesStart, i);
    if (mStart <= lastClosedMonthStart) closedEndIdx = i;
  }
  if (closedEndIdx < 0) {
    const w = 1 / productNames.length;
    productNames.forEach(n => map.set(n, w));
    return map;
  }

  const startIdx = Math.max(0, closedEndIdx - 11);
  const sums = productNames.map((name, i) => {
    const arr = monthlyByProduct[i] || [];
    let s = 0;
    for (let k = startIdx; k <= closedEndIdx; k++) s += Number(arr[k] || 0);
    return s;
  });
  const total = sums.reduce((a,b)=>a+b,0);

  if (total > 0) {
    productNames.forEach((name, i) => map.set(name, sums[i] / total));
  } else {
    const w = 1 / productNames.length;
    productNames.forEach(name => map.set(name, w));
  }
  return map;
}

/** DEV固定（12ヶ月）※必要情報が揃った行だけ加算 */
function readDevFixed12Months_(fy) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.DEV);
  const out = new Array(12).fill(0);
  if (!sh) return out;

  const last = sh.getLastRow();
  if (last < 2) return out;

  const vals = sh.getRange(2, 1, last - 1, 5).getValues();
  const start = new Date(fy, 3, 1);

  vals.forEach(r => {
    const dt = toDate_(r[1]);
    if (!dt) return;

    const amountRaw = r[3];
    const confRaw = r[4];
    if (amountRaw === '' || amountRaw === null) return;
    if (confRaw === '' || confRaw === null) return;

    const amt = toNumberSafe_(amountRaw);
    if (!isFinite(amt) || amt === 0) return;

    const conf = Number(confRaw);
    if (!isFinite(conf) || conf < 0 || conf > 1) return;

    const idx = monthIndexFromStart_(dt, start);
    if (idx < 0 || idx >= 12) return;

    out[idx] += amt * conf;
  });

  return out;
}

/** FACTORS_PRODUCT ※必要情報が揃った行だけ */
function readFactorsProduct_(fy) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.FACTORS_PRODUCT);
  if (!sh || sh.getLastRow() < 2) return [];

  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
  return vals.map(r => {
    const person = String(r[0] || '').trim();
    const product = String(r[1] || '').trim();
    const monthRaw = r[2];
    const stepRaw = r[3];
    const reason = String(r[4] || '').trim();

    // 必須が揃っていない行は無視
    if (!person || !product || !monthRaw || stepRaw === '' || stepRaw === null) return null;

    const dt = toDate_(monthRaw);
    const step = parseRate_(stepRaw);
    if (!dt || !isFinite(step)) return null;

    return { person, product, month: dt, step, reason };
  }).filter(Boolean);
}

/** FACTORS_CLIENT ※必要情報が揃った行だけ */
function readFactorsClient_(fy) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.FACTORS_CLIENT);
  if (!sh || sh.getLastRow() < 2) return [];

  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
  return vals.map(r => {
    const person = String(r[0] || '').trim();
    const monthRaw = r[1];
    const stepRaw = r[2];
    const reason = String(r[3] || '').trim();

    if (!person || !monthRaw || stepRaw === '' || stepRaw === null) return null;

    const dt = toDate_(monthRaw);
    const step = parseRate_(stepRaw);
    if (!dt || !isFinite(step)) return null;

    return { person, month: dt, step, reason };
  }).filter(Boolean);
}

/** OPINIONS ※必要情報が揃った行だけ */
function readOpinions_(fy) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.OPINIONS);
  if (!sh || sh.getLastRow() < 2) return [];

  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
  return vals.map(r => {
    const person = String(r[0] || '').trim();
    const monthRaw = r[1];
    const stepRaw = r[2];
    const confRaw = r[3];
    const note = String(r[4] || '').trim();

    if (!person || !monthRaw || stepRaw === '' || stepRaw === null || confRaw === '' || confRaw === null) return null;

    const dt = toDate_(monthRaw);
    const step = parseRate_(stepRaw);
    const conf = Number(confRaw);

    if (!dt || !isFinite(step) || !isFinite(conf) || conf < 0 || conf > 1) return null;

    return { person, month: dt, step, confidence: conf, note };
  }).filter(Boolean);
}

/** ====== 予測計算（モデル） ====== */
function fitOpsModelTrendSeason_(y) {
  const n = y.length;
  const x = [];
  for (let i = 0; i < n; i++) x.push(i + 1);

  const slope = slope_(y, x);
  const intercept = intercept_(y, x, slope);

  const ma12 = movingAverage_(y, 12);
  const ratios = y.map((v, i) => (ma12[i] > 0 ? v / ma12[i] : 1));

  const seasonal = new Array(12).fill(1);
  for (let m = 0; m < 12; m++) {
    const arr = [];
    for (let i = 0; i < n; i++) {
      if ((i % 12) === m && isFinite(ratios[i]) && ratios[i] > 0) arr.push(ratios[i]);
    }
    seasonal[m] = arr.length ? avg_(arr) : 1;
  }
  for (let m = 0; m < 12; m++) seasonal[m] = Math.max(0.80, Math.min(1.20, seasonal[m]));

  const fitted = y.map((_, i) => Math.max(0, (intercept + slope * (i + 1)) * seasonal[i % 12]));
  return { slope, intercept, seasonalIndex: seasonal, fitted };
}

function forecastByResidualQuantiles_(model, devFixedByMonth, q) {
  const p10 = [], p50 = [], p90 = [];
  const startT = 48;
  for (let i = 0; i < 12; i++) {
    const t = startT + (i + 1);
    const monthIdx = i % 12;
    const base = Math.max(0, (model.intercept + model.slope * t) * model.seasonalIndex[monthIdx]);
    p10.push(base * (1 + q.p10) + devFixedByMonth[i]);
    p50.push(base * (1 + q.p50) + devFixedByMonth[i]);
    p90.push(base * (1 + q.p90) + devFixedByMonth[i]);
  }
  return { p10, p50, p90 };
}

function forecastMonteCarloMixed_(model, devFixedByMonth, opt) {
  const nSim = opt.nSim || 1000;
  const residualPct = opt.residualPct;
  const factorsProduct = opt.factorsProduct || [];
  const factorsClient = opt.factorsClient || [];
  const opinions = opt.opinions || [];
  const productWeights = opt.productWeights || new Map();
  const months = opt.months || [];

  const kProdByMonth = months.map(m => productFactorsMultiplier_(factorsProduct, m, productWeights));
  const kClientByMonth = months.map(m => clientFactorsMultiplier_(factorsClient, m));

  const startT = 48;
  const simByMonth = Array.from({ length: 12 }, () => []);

  for (let s = 0; s < nSim; s++) {
    for (let i = 0; i < 12; i++) {
      const t = startT + (i + 1);
      const mIdx = i % 12;

      const base = Math.max(0, (model.intercept + model.slope * t) * model.seasonalIndex[mIdx]);
      const e = residualPct[Math.floor(Math.random() * residualPct.length)] || 0;

      let ops = base * (1 + e);
      ops *= kProdByMonth[i];
      ops *= kClientByMonth[i];
      ops *= sampleOpinionMultiplier_(opinions, months[i]);

      const total = Math.max(0, ops) + devFixedByMonth[i];
      simByMonth[i].push(total);
    }
  }

  const p10 = simByMonth.map(arr => percentile_(arr, 0.10));
  const p50 = simByMonth.map(arr => percentile_(arr, 0.50));
  const p90 = simByMonth.map(arr => percentile_(arr, 0.90));

  return { p10, p50, p90 };
}

/** 製品要因：製品別step合算 → 構成比で加重 → 1+加重step */
function productFactorsMultiplier_(factorsProduct, targetMonth, productWeights) {
  if (!factorsProduct || factorsProduct.length === 0) return 1;

  const stepByProduct = new Map();
  factorsProduct.forEach(f => {
    if (!f.month || f.month > targetMonth) return;
    const p = f.product;
    if (!p) return;
    const prev = stepByProduct.get(p) || 0;
    stepByProduct.set(p, prev + (isFinite(f.step) ? f.step : 0));
  });
  if (stepByProduct.size === 0) return 1;

  let aggStep = 0;
  stepByProduct.forEach((step, p) => {
    const w = productWeights.has(p) ? productWeights.get(p) : 0;
    aggStep += w * step;
  });

  const mult = 1 + aggStep;
  return Math.max(0, mult);
}

/** クライアント要因：step合算 → 1+step */
function clientFactorsMultiplier_(factorsClient, targetMonth) {
  if (!factorsClient || factorsClient.length === 0) return 1;

  let step = 0;
  factorsClient.forEach(f => {
    if (!f.month || f.month > targetMonth) return;
    step += (isFinite(f.step) ? f.step : 0);
  });

  const mult = 1 + step;
  return Math.max(0, mult);
}

/** 意見係数：担当者別に最新意見を取り、±5%のランダム揺らしを入れて合成 */
function sampleOpinionMultiplier_(opinions, targetMonth) {
  if (!opinions || opinions.length === 0) return 1;

  const people = new Map();
  opinions.forEach(o => {
    if (!o.month || o.month > targetMonth) return;
    const key = o.person || '';
    if (!key) return;
    const prev = people.get(key);
    if (!prev || prev.month < o.month) people.set(key, o);
  });
  if (people.size === 0) return 1;

  let k = 1;
  people.forEach(o => {
    const baseStep = o.step;
    const conf = isFinite(o.confidence) ? o.confidence : 0.7;

    const jitter = (Math.floor(Math.random() * 3) - 1) * 0.05; // -0.05,0,+0.05
    const stepRand = baseStep + jitter;

    k *= (1 + stepRand * conf);
  });

  return k;
}

/** ====== 意見要約 ====== */
function summarizeOpinionsTop_(opinions) {
  if (!opinions || opinions.length === 0) return '';

  const latest = new Map();
  opinions.forEach(o => {
    const prev = latest.get(o.person);
    if (!prev || (prev.month && o.month && prev.month < o.month)) latest.set(o.person, o);
  });

  const parts = [];
  latest.forEach(o => {
    const pct = Math.round(o.step * 100);
    const sign = pct > 0 ? '+' : '';
    const conf = isFinite(o.confidence) ? o.confidence.toFixed(2) : '0.70';
    const memo = o.note ? `：${o.note}` : '';
    parts.push(`${o.person} ${sign}${pct}%(${conf})${memo}`);
  });
  return parts.join(' / ');
}

function summarizeOpinionsByMonth_(opinions, months) {
  const out = [];
  for (let i = 0; i < months.length; i++) {
    const m = months[i];
    const applicable = opinions.filter(o => o.month && o.month <= m && o.note && String(o.note).trim());
    const latest = new Map();
    applicable.forEach(o => {
      const prev = latest.get(o.person);
      if (!prev || prev.month < o.month) latest.set(o.person, o);
    });
    const parts = [];
    latest.forEach(o => {
      const pct = Math.round(o.step * 100);
      const sign = pct > 0 ? '+' : '';
      const conf = isFinite(o.confidence) ? o.confidence.toFixed(2) : '0.70';
      parts.push(`${o.person}:${sign}${pct}%(${conf})`);
    });
    out.push(parts.join(' / '));
  }
  return out;
}

/** ====== スムージング（季節性を潰しにくい単発スパイクならし） ====== */
function smoothSeriesSeasonalAware_(y) {
  const n = y.length;
  if (n !== 48) return y.slice();

  const base = new Array(n).fill(0);
  for (let i = 0; i < n; i++) {
    const start = Math.max(0, i - 12);
    const end = i;
    const arr = [];
    for (let j = start; j < end; j++) arr.push(y[j]);
    base[i] = arr.length ? avg_(arr) : (y[i] || 0);
  }

  const byM = Array.from({ length: 12 }, () => []);
  for (let i = 0; i < n; i++) {
    const b = base[i] || 0;
    const ratio = b > 0 ? (y[i] / b) : 1;
    if (isFinite(ratio) && ratio > 0) byM[i % 12].push(ratio);
  }

  const mMed = new Array(12).fill(1);
  const mMad = new Array(12).fill(0.1);

  for (let m = 0; m < 12; m++) {
    const arr = byM[m].slice().sort((a,b)=>a-b);
    mMed[m] = arr.length ? percentileSorted_(arr, 0.50) : 1;

    const dev = arr.map(v => Math.abs(v - mMed[m])).sort((a,b)=>a-b);
    mMad[m] = dev.length ? percentileSorted_(dev, 0.50) : 0.1;
    if (mMad[m] === 0) mMad[m] = 0.05;
  }

  const out = y.slice();
  for (let i = 0; i < n; i++) {
    const b = base[i] || 0;
    if (b <= 0) continue;

    const ratio = out[i] / b;
    if (!isFinite(ratio) || ratio <= 0) continue;

    const isSpikeCandidate = (ratio > SPIKE_CLIP_MAX || ratio < SPIKE_CLIP_MIN);
    if (!isSpikeCandidate) continue;

    const m = i % 12;
    const lo = Math.max(0.30, mMed[m] - SEASONAL_MAD_K * mMad[m]);
    const hi = Math.max(lo + 0.05, mMed[m] + SEASONAL_MAD_K * mMad[m]);

    // 季節範囲内なら潰さない
    if (ratio >= lo && ratio <= hi) continue;

    const clipped = Math.max(SPIKE_CLIP_MIN, Math.min(SPIKE_CLIP_MAX, ratio));
    out[i] = b * clipped;
  }
  return out;
}

/** ====== ユーティリティ ====== */
function ensureSetupDone_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(SHEETS.CONFIG)) {
    throw new Error('初期セットアップが必要です。Forecast Agent > ① 初期セットアップ を実行してください。');
  }
}

function ensureSheetHasColumns_(sh, minCols) {
  const cur = sh.getMaxColumns();
  if (cur < minCols) sh.insertColumnsAfter(cur, minCols - cur);
}

function getOrCreateSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function escapeHtml_(s) {
  return String(s || '')
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;')
    .replace(/'/g,'&#039;');
}

function fmtYM_(d) {
  return Utilities.formatDate(d, TZ, 'yyyy/MM');
}

function addMonths_(d, n) {
  const x = new Date(d.getTime());
  x.setMonth(x.getMonth() + n);
  return x;
}

function monthIndexFromStart_(dt, start) {
  return (dt.getFullYear() - start.getFullYear()) * 12 + (dt.getMonth() - start.getMonth());
}

function toDate_(v) {
  if (!v) return null;
  if (v instanceof Date && !isNaN(v.getTime())) return v;

  const s = String(v).trim();
  if (!s) return null;

  const norm = s.replace(/\//g,'-');
  const d = new Date(norm);
  if (!isNaN(d.getTime())) return d;
  return null;
}

/**
 * Stepの解釈（計算用）
 * - "-30%" / "-30％" → -0.30
 * - "-30"  → -0.30（%として扱う）
 * - "-0.3" → -0.30（比率として扱う）
 * - 解釈不能な文字列は NaN を返す（検出用）
 */
function parseRate_(v) {
  if (v === null || v === undefined) return 0;
  if (typeof v === 'number') {
    if (!isFinite(v)) return 0;
    return (Math.abs(v) > 1) ? (v / 100) : v;
  }
  const s0 = String(v).trim();
  if (!s0) return 0;

  const s = s0.replace(/％/g, '%').replace(/[,\s]/g,'').replace(/¥/g,'').replace(/￥/g,'');
  const m = s.match(/^([+-]?\d+(?:\.\d+)?)\s*%$/);
  if (m) return Number(m[1]) / 100;

  const num = Number(s);
  if (isFinite(num)) {
    return (Math.abs(num) > 1) ? (num / 100) : num;
  }
  return NaN;
}

/** 表示の正規化（onEdit用）：常に "+10%" / "-30%" / "0%" にする */
function normalizeStepDisplay_(v) {
  if (v === null || v === undefined) return null;

  if (typeof v === 'string' && v.trim() === '') return '';

  let numPct = null;

  if (typeof v === 'number') {
    if (!isFinite(v)) return null;
    numPct = (Math.abs(v) > 2) ? v : v * 100;
  } else {
    const s0 = String(v).trim();
    if (!s0) return '';
    const s = s0.replace(/％/g, '%').replace(/[,\s]/g,'').replace(/¥/g,'').replace(/￥/g,'');
    const m = s.match(/^([+-]?\d+(?:\.\d+)?)\s*%$/);
    if (m) {
      numPct = Number(m[1]);
    } else {
      const x = Number(s);
      if (isFinite(x)) {
        numPct = (Math.abs(x) > 2) ? x : x * 100;
      } else {
        return s0;
      }
    }
  }

  if (!isFinite(numPct)) return null;

  const rounded = Math.round(numPct * 2) / 2;

  if (rounded === 0) return '0%';
  const sign = rounded > 0 ? '+' : '';
  const txt = (Math.abs(rounded - Math.round(rounded)) < 1e-9) ? String(Math.round(rounded)) : String(rounded);
  return `${sign}${txt}%`;
}

function buildPercentStepList_() {
  const arr = [];
  for (let p = 100; p >= -100; p -= 5) {
    const sign = p > 0 ? '+' : '';
    arr.push(`${sign}${p}%`);
  }
  return arr;
}

function avg_(arr) {
  if (!arr || !arr.length) return 0;
  return arr.reduce((a,b)=>a+b,0) / arr.length;
}

function sumArr_(arr) {
  return (arr || []).reduce((a,b)=>a + (Number(b) || 0), 0);
}

function percentile_(arr, q) {
  if (!arr || !arr.length) return 0;
  const a = arr.slice().sort((x,y)=>x-y);
  return percentileSorted_(a, q);
}

function percentileSorted_(a, q) {
  const n = a.length;
  if (n === 0) return 0;
  const pos = (n - 1) * q;
  const base = Math.floor(pos);
  const rest = pos - base;
  if ((a[base + 1] !== undefined)) {
    return a[base] + rest * (a[base + 1] - a[base]);
  } else {
    return a[base];
  }
}

function movingAverage_(arr, window) {
  const n = arr.length;
  const out = new Array(n).fill(0);
  for (let i = 0; i < n; i++) {
    const start = Math.max(0, i - window + 1);
    const slice = arr.slice(start, i + 1);
    out[i] = avg_(slice);
  }
  return out;
}

// regression helpers
function slope_(y, x) {
  const n = y.length;
  const xbar = avg_(x);
  const ybar = avg_(y);
  let num = 0, den = 0;
  for (let i = 0; i < n; i++) {
    num += (x[i] - xbar) * (y[i] - ybar);
    den += (x[i] - xbar) * (x[i] - xbar);
  }
  return den === 0 ? 0 : num / den;
}

function intercept_(y, x, slope) {
  const xbar = avg_(x);
  const ybar = avg_(y);
  return ybar - slope * xbar;
}

/** ====== toast補助（読み取り時間を確保） ====== */
function toastProgress_(ss, message, seconds) {
  ss.toast(message, MENU_NAME, seconds || 5);
  // 読み取れる程度に少し待つ（スピード最優先ではない方針）
  Utilities.sleep(450);
}
