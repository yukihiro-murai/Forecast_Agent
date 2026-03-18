/***************************************
 * Forecast Agent v1.2
 * 単一メーカー（1クライアント）用 / Google Sheets 実装
 *
 * v1.2（今回反映）
 * - 未確定月補完：月別（同月）トレンドで補完し、補完後に途中実績より下がらない
 * - 未確定月判定：実行日ベースで可変（当月以降=未確定 / 前月まで=確定）
 * - FACTORS/OPINIONS/DEV_SPOT：必要情報が揃った行のみ計算に使用
 * - 入力異常検出：変な入力があれば実行前にエラー表示して停止
 * - OUTPUT：B=ネガ / C=中立 / D=ポジ、配色も統一（表＆グラフ）
 * - 実行中メッセージ：計算ステップが分かるtoastを追加（読み取り時間も確保）
 ***************************************/

const VERSION = '1.2';
const MENU_NAME = 'Forecast Agent';

/***************************************
 * 運用コメント（Phase移行基準・実務ルール）
 *
 * [Phase1 -> Phase2 移行の目安]
 * 1) 最低3か月、月次運用が安定して継続されていること
 *    - A-1/A-2/A-9/A-10/B-2 の実行漏れがなく、PROCESS_STATUSが継続的に success
 * 2) 精度KPIが最低基準を満たすこと
 *    - 全体sMAPE <= 30%（目安）
 *    - 実測がネガ〜ポジ帯に入る割合 >= 70%（目安）
 * 3) データ品質が担保されること
 *    - SALES_INPUT_MONTHLY / ACTUAL_EVAL_MONTHLY の欠損・異常値が許容範囲
 * 4) 現場利用が定着していること
 *    - GUIDEに沿った操作で、担当者が自力運用できる
 *
 * [Phase2で優先的に着手する内容]
 * - 重み更新の高度化（クライアント別最適化）
 * - シミュレーション高度化（分位点回帰との比較導入）
 * - モデル監視（ドリフト検知、エラー分類の定例化）
 *
 * [月次運用ルール（推奨順）]
 * 1) A-2 売上データを取り込み
 * 2) （必要時）A-8 AI調査を取り込む→AI結果貼付
 * 3) A-9 予測実行（単一クライアント）
 * 4) A-10 予測ダッシュボードを更新
 * 5) 実績確定後にB-1検証実績取り込み→B-2予測検証レポート更新
 *
 * [運用時の注意]
 * - 本ツールは「確認→修正→再実行」を前提とする（一発確定しない）
 * - OUTPUTは要点表示、詳細根拠はFORECAST_REPORT/FORECAST_SNAPSHOTで確認
 * - AI結果は補助情報。形式・値域チェックに通らない情報は反映しない
 * - 初期セットアップは全タブ再作成（既存タブ削除）なので本番時は必ず注意喚起
 * - 重大な仕様変更を行った場合は、GUIDEとCHANGELOG（運用記録）を同時更新
 ***************************************/

const SHEETS = {
  GUIDE: 'GUIDE',
  CONFIG: 'CONFIG',
  SALES: 'SALES',
  FACTORS_PRODUCT: 'FACTORS_PRODUCT',
  FACTORS_CLIENT: 'FACTORS_CLIENT',
  OPINIONS: 'OPINIONS',
  DEV_SPOT: 'DEV_SPOT',
  OUTPUT: 'OUTPUT',
  SALES_INPUT_MONTHLY: 'SALES_INPUT_MONTHLY',
  ACTUAL_EVAL_MONTHLY: 'ACTUAL_EVAL_MONTHLY',
  AI_RESEARCH_PROMPT: 'AI_RESEARCH_PROMPT',
  AI_RESEARCH_STRUCTURED: 'AI_RESEARCH_STRUCTURED',
  RUN_LOG: 'RUN_LOG',
  FORECAST_SNAPSHOT: 'FORECAST_SNAPSHOT',
  EVAL_LOG: 'EVAL_LOG',
  EVAL_COMPARE_MONTHLY: 'EVAL_COMPARE_MONTHLY',
  EVAL_INSIGHTS: 'EVAL_INSIGHTS',
  OVERRIDE_LOG: 'OVERRIDE_LOG',
  WEIGHT_UPDATE_LOG: 'WEIGHT_UPDATE_LOG',
  SPIKE_LOG: 'SPIKE_LOG',
  PROCESS_STATUS: 'PROCESS_STATUS',
  CLIENT_PARAMS: 'CLIENT_PARAMS',
  DETERMINISTIC_FACTORS: 'DETERMINISTIC_FACTORS',
  FORECAST_REPORT: 'FORECAST_REPORT',
  DASHBOARD: 'DASHBOARD',
  CHANGELOG: 'CHANGELOG'
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
const EXT_COL_SERVICE_CATEGORY = 46; // AT（サービスカテゴリ）
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

// A-9 実行前の影響度チェック閾値
const STEP_WARN_THRESHOLD = 0.30;   // ±30%
const STEP_STRONG_THRESHOLD = 0.50; // ±50%
const STEP_BLOCK_THRESHOLD = 1.00;  // ±100%
const K_TOTAL_WARN_MIN = 0.70;
const K_TOTAL_WARN_MAX = 1.30;
const K_TOTAL_BLOCK_MIN = 0.50;
const K_TOTAL_BLOCK_MAX = 1.50;

// SPOT背景推定（未知のスポット発生を最低限拾う）
const SPOT_BG_SHRINK = 0.50;      // 履歴同月平均の50%を背景SPOTとして採用
const SPOT_BG_FLOOR_RATE = 0.15;  // 履歴同月平均の15%は最低保証
const SPOT_BG_CAP_RATE = 0.20;    // 背景SPOTの上限（BASE予測P50比）
const AI_WEIGHT_DEFAULT = 0.0005; // AI重み（既定）
const AI_MAX_ABS_EFFECT = 0.05;   // AI係数の絶対上限（±5%）

/** ====== メニュー ====== */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu(MENU_NAME)
    .addItem('A-1 初期セットアップ', 'setupForecastBook')
    .addSeparator()
    .addItem('A-2 売上データを取り込む', 'importSalesInputMonthly')
    .addItem('A-3 予測用に売上データを加工', 'aggregateSalesData')
    .addItem('A-4 製品ごとの動向を入力', 'openProductTrendEntryDialog')
    .addItem('A-5 クライアント動向を入力', 'openClientTrendEntryDialog')
    .addItem('A-6 担当者意見を入力', 'openOpinionsEntryDialog')
    .addItem('A-7 開発/スポット要因を入力', 'openDevEntryDialog')
    .addItem('A-8 AI調査を取り込む', 'generateAIResearchTemplate')
    .addItem('A-9 予測を実行', 'runPhase1Forecast')
    .addItem('A-10 予測ダッシュボードを更新', 'updatePhase1Dashboard')
    .addSeparator()
    .addItem('B-1 検証用に実績データを取り込み', 'importActualEvalMonthly')
    .addItem('B-2 検証レポートを更新', 'updatePhase1EvaluationReport')
    .addItem('B-3 検証インサイトを更新', 'updatePhase1LearningInsights')
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

/** ====== A-1 初期セットアップ ====== */
function setupForecastBook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const res = ui.alert(
    '初期セットアップ（全上書き）',
    '初期セットアップで全て上書きされますがよろしいですか？\n\n※既存のシートタブは削除されます。',
    ui.ButtonSet.OK_CANCEL
  );
  if (res !== ui.Button.OK) return;

  const order = [
    SHEETS.GUIDE,
    SHEETS.OUTPUT,
    SHEETS.CONFIG,
    SHEETS.SALES_INPUT_MONTHLY,
    SHEETS.SALES,
    SHEETS.AI_RESEARCH_PROMPT,
    SHEETS.FACTORS_PRODUCT,
    SHEETS.FACTORS_CLIENT,
    SHEETS.OPINIONS,
    SHEETS.DEV_SPOT,
    SHEETS.AI_RESEARCH_PROMPT,
    SHEETS.OUTPUT,
    SHEETS.FORECAST_REPORT,
    SHEETS.DASHBOARD,
    SHEETS.ACTUAL_EVAL_MONTHLY,
    SHEETS.EVAL_COMPARE_MONTHLY,
    SHEETS.EVAL_LOG,
    SHEETS.EVAL_INSIGHTS,
    SHEETS.AI_RESEARCH_STRUCTURED,
    SHEETS.RUN_LOG,
    SHEETS.FORECAST_SNAPSHOT,
    SHEETS.PROCESS_STATUS
  ];

  try {
    resetWorkbookSheets_(ss, order);

    buildGUIDE_();
    buildCONFIG_();
    buildSALES_();
    buildFACTORS_PRODUCT_();
    buildFACTORS_CLIENT_();
    buildOPINIONS_();
    buildDEV_();
    buildPhase1Sheets_();
    buildOUTPUT_();
    applyTabColors_();
    hideNonUserSheets_();
    const guide = ss.getSheetByName(SHEETS.GUIDE);
    if (guide) ss.setActiveSheet(guide);

    showInitialSetupDialog_();
  } catch (e) {
    ui.alert('初期セットアップでエラー', `${e && e.message ? e.message : e}`);
  }
}

function resetWorkbookSheets_(ss, order) {
  var required = {};
  for (var i = 0; i < order.length; i++) required[order[i]] = true;

  for (var j = 0; j < order.length; j++) {
    if (!ss.getSheetByName(order[j])) ss.insertSheet(order[j]);
  }

  var current = ss.getSheets();
  for (var k = 0; k < current.length; k++) {
    var sh = current[k];
    if (required[sh.getName()]) continue;
    try {
      ss.deleteSheet(sh);
    } catch (e) {
      try { sh.hideSheet(); } catch (ignore) {}
    }
  }

  for (var x = 0; x < order.length; x++) {
    var target = ss.getSheetByName(order[x]);
    if (!target) continue;
    try { target.showSheet(); } catch (e2) {}
    safeMoveSheet_(ss, target, x + 1);
  }
}

function safeMoveSheet_(ss, sh, targetIndex) {
  if (!sh) return;
  try {
    var max = ss.getSheets().length;
    var idx = targetIndex;
    if (idx < 1) idx = 1;
    if (idx > max) idx = max;
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(idx);
  } catch (e) {
    // 並び替え失敗時は継続
  }
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
      ※クライアント名の候補は外部実績シートから自動抽出しています。
    </div>
  </div>

  <div class="block">
    <label>何年度（FY）を予測しますか？</label>
    <input id="fy" type="number" />
    <div class="hint">※ 空欄の場合デフォルト年度（${defaultFY}年）を使用。（決算月：${defaultFY + 1}年3月）</div>
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
  ss.setActiveSheet(ss.getSheetByName(SHEETS.GUIDE));
}

/**
 * デフォルトFY（3月末決算の前後6か月基準）：
 * - 実行日を6か月進めた日付の「年」をFYとして採用
 *   例) 2026/04 実行 → 2026/10 相当 → FY2026
 *   例) 2026/10 実行 → 2027/04 相当 → FY2027
 */
function getDefaultFY_() {
  const now = new Date();
  const shifted = new Date(now.getFullYear(), now.getMonth() + 6, 1);
  return shifted.getFullYear();
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
      if (v && String(v).trim()) set.add(normalizeClientName_(String(v).trim()));
    });
  });

  return Array.from(set).sort();
}

function normalizeClientName_(name) {
  const s = String(name || '').trim();
  if (s === 'ｳﾞｨｱﾄﾘｽ製薬(株)' || s === 'ｳﾞｨｱﾄﾘｽ製薬合同会社') return 'ｳﾞｨｱﾄﾘｽ製薬';
  return s;
}

function isSameClient_(a, b) {
  return normalizeClientName_(a) === normalizeClientName_(b);
}

/** ====== 補助: 過去売上データを反映（外部SS→このSSのSALES） ====== */
function importPastSalesToSalesTab() {
  ensureSetupDone_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName(SHEETS.CONFIG);

  const client = String(cfg.getRange('B2').getValue() || '').trim();
  const fy = Number(cfg.getRange('B3').getValue());

  if (!client) {
    SpreadsheetApp.getUi().alert('CONFIG!B2 にメーカー名が未設定です。A-1 初期セットアップを実行してください。');
    return;
  }
  if (!fy || !isFinite(fy)) {
    SpreadsheetApp.getUi().alert('CONFIG!B3 に予測FYが未設定です。A-1 初期セットアップを実行してください。');
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

  // 予測FY=2026なら → 2022,2023,2024,2025,2026（2026年は1〜3月を使用）
  const years = [fy - 4, fy - 3, fy - 2, fy - 1, fy];
  const tabNames = years.map(y => `${EXTERNAL_SHEET_PREFIX}${y}${EXTERNAL_SHEET_SUFFIX}`);

  const start = new Date(fy - 3, 3, 1); // fy-3/04/01
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
      if (!isSameClient_(c, client)) continue;

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

  sales.getRange(1, 1).setValue('Category');
  sales.getRange(1, 2, 1, totalMonths).setValues([headerMonths]);

  const base = new Array(totalMonths).fill(0);
  const spot = new Array(totalMonths).fill(0);
  map.forEach((arr, name) => {
    const key = String(name || '').toUpperCase();
    if (key.includes('SPOT')) {
      for (let i = 0; i < totalMonths; i++) spot[i] += Number(arr[i] || 0);
    } else {
      for (let i = 0; i < totalMonths; i++) base[i] += Number(arr[i] || 0);
    }
  });
  const out = [['BASE', ...base], ['SPOT', ...spot]];
  sales.getRange(2, 1, out.length, 1 + totalMonths).setValues(out);

  // 客観（黄色）
  sales.getRange(2, 2, out.length, totalMonths).setBackground(COLOR_OBJECTIVE);

  sales.setFrozenRows(1);
  sales.setFrozenColumns(1);
  sales.autoResizeColumns(1, 1);

  // 取り込み完了後にSALESを開く
  ss.setActiveSheet(sales);

  ui.alert('完了', `SALESに過去4年分（48ヶ月）の売上を反映しました。\nメーカー: ${client}`, ui.ButtonSet.OK);
}

/** ====== A-4〜A-7：シート整形＋使い方案内（ポップアップは説明のみ） ====== */
function openProductTrendEntryDialog() {
  ensureSetupDone_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const people = getPeopleListFromConfig_();
  if (people.length === 0) {
    SpreadsheetApp.getUi().alert('CONFIG!B10 に担当者が設定されていません。\nA-1 初期セットアップで担当者を入力してください。');
    return;
  }
  const products = getProductNameListFromSales_();
  if (products.length === 0) {
    SpreadsheetApp.getUi().alert('SALESに製品名がありません。\nA-2〜A-3 を先に実行してください。');
    return;
  }

  const cfg = ss.getSheetByName(SHEETS.CONFIG);
  const fy = Number(cfg.getRange('B3').getValue()) || getDefaultFY_();
  const defaultDate = new Date(fy - 1, 3, 1);

  const sh = ss.getSheetByName(SHEETS.FACTORS_PRODUCT);
  ensureFactorsProductTemplate_(sh, products, people, defaultDate);

  ss.setActiveSheet(sh);

  showInfoDialog_(
    'A-4 製品動向を入力',
    [
      'FACTORS_PRODUCT を入力してください（青色のセルが対象です）。',
      '1) A列：担当者を選択',
      '2) C列：影響が出る日付（この日付以降に反映）',
      '3) D列：増減率（例：-30% = 今後30%減りそう）',
      '4) E列：根拠を短く',
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
    SpreadsheetApp.getUi().alert('CONFIG!B10 に担当者が設定されていません。\nA-1 初期セットアップで担当者を入力してください。');
    return;
  }

  const cfg = ss.getSheetByName(SHEETS.CONFIG);
  const fy = Number(cfg.getRange('B3').getValue()) || getDefaultFY_();
  const defaultDate = new Date(fy - 1, 3, 1);

  const sh = ss.getSheetByName(SHEETS.FACTORS_CLIENT);
  ensureFactorsClientTemplate_(sh, people, defaultDate);

  ss.setActiveSheet(sh);

  showInfoDialog_(
    'A-5 クライアント動向を入力',
    [
      'FACTORS_CLIENT を入力してください（青色のセルが対象です）。',
      '1) A列：担当者を選択',
      '2) B列：影響が出る日付（この日付以降に反映）',
      '3) C列：増減率（例：-10% = 予算圧縮で10%減りそう）',
      '4) D列：根拠を短く',
      '※ Stepは入力ゆらぎが出ないよう自動で「+10%/-30%」形式に整えます。'
    ]
  );
}

function openOpinionsEntryDialog() {
  ensureSetupDone_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const people = getPeopleListFromConfig_();
  if (people.length === 0) {
    SpreadsheetApp.getUi().alert('CONFIG!B10 に担当者が設定されていません。\nA-1 初期セットアップで担当者を入力してください。');
    return;
  }

  const cfg = ss.getSheetByName(SHEETS.CONFIG);
  const fy = Number(cfg.getRange('B3').getValue()) || getDefaultFY_();
  const defaultDate = new Date(fy - 1, 3, 1);

  const sh = ss.getSheetByName(SHEETS.OPINIONS);
  ensureOpinionsTemplate_(sh, people, defaultDate);

  ss.setActiveSheet(sh);

  showInfoDialog_(
    'A-6 メーカー担当者意見を入力',
    [
      'OPINIONS を入力してください（青色のセルが対象です）。',
      '※原則として担当者全員の入力が必要です（未入力があるとA-9が実行できません）。',
      '入力手順：',
      '1) B列：影響が出る日付（この日付以降に反映）',
      '2) C列：増減率（例：+20% = 今後20%増えそう）',
      '3) D列：信頼度（0..1）',
      '4) E列：所感を短く',
      '※ 意見はそのまま固定反映されず、シミュレーション内でランダムに活用されます。'
    ]
  );
}

function openDevEntryDialog() {
  ensureSetupDone_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const people = getPeopleListFromConfig_();
  if (people.length === 0) {
    SpreadsheetApp.getUi().alert('CONFIG!B10 に担当者が設定されていません。\nA-1 初期セットアップで担当者を入力してください。');
    return;
  }

  const cfg = ss.getSheetByName(SHEETS.CONFIG);
  const fy = Number(cfg.getRange('B3').getValue()) || getDefaultFY_();
  const defaultDate = new Date(fy - 1, 3, 1);

  const sh = ss.getSheetByName(SHEETS.DEV_SPOT);
  ensureDevTemplate_(sh, people, defaultDate);

  ss.setActiveSheet(sh);

  showInfoDialog_(
    'A-7 開発/スポット要因を入力',
    [
      'DEV_SPOT を入力してください（青色のセルが対象です）。',
      '開発案件だけでなく、スポット要因（例：法改定による差し替え等）もここに入力してください。',
      '入力手順：',
      '1) A列：担当者を選択',
      '2) B列：売上が立つ日付（この日付の月に反映）',
      '3) C列：案件名/スポット要因名',
      '4) D列：金額（円）',
      '5) E列：確度（0..1）',
      '※ DEV_SPOTは「金額×確度」で固定加算されます（運用のシミュレーションには混ぜません）。'
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

/** ====== A-9 予測を出力 ====== */
function executeForecastUsingConfig() {
  ensureSetupDone_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName(SHEETS.CONFIG);

  const client = String(cfg.getRange('B2').getValue() || '').trim();
  const fy = Number(cfg.getRange('B3').getValue());
  if (!client || !fy) {
    SpreadsheetApp.getUi().alert('初期設定が未完了です。A-1 初期セットアップを実行してください。');
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
      `OPINIONSに全員の意見が必要です。\n未入力: ${missingPeople.join(', ')}\n\nA-6で全員分入力してください。`,
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
  syncSalesFromSalesInput_(fy, client);
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
  const tuning = readModelTuningFromConfig_();
  const runDate = new Date();
  const ctx = getForecastContext_(fy, runDate, salesData.headerMonths || []);

  if (!salesData.isComplete48) {
    const ui = SpreadsheetApp.getUi();
    const res = ui.alert(
      '注意：売上データが48ヶ月揃っていません',
      '48ヶ月（過去4年）揃っていない場合、予測精度が下がる可能性があります。\nこのままシミュレーションを続行しますか？',
      ui.ButtonSet.OK_CANCEL
    );
    if (res !== ui.Button.OK) throw new Error('ユーザーが中断しました。');
  }

  // 予測の土台はBASEのみ（SPOTは別途、背景成分として扱う）
  const aggY_raw = salesData.baseSeries48 && salesData.baseSeries48.length
    ? salesData.baseSeries48.slice()
    : sumAcrossProducts_(salesData.monthlyByProduct);

  const seriesStart = salesData.headerMonths && salesData.headerMonths.length ? salesData.headerMonths[0] : new Date(fy - 4, 3, 1);
  const unclosedAdjusted = adjustForUnclosedMonths_(aggY_raw, seriesStart);
  const aggY_adj = unclosedAdjusted.series.slice();

  toastProgress_(ss, 'STEP2/6: スパイクをならし（季節性は維持）→ トレンド＋季節性を推定…', 7);

  // スムージング（季節性は守りつつ単発スパイクだけ弱める）
  const smoothY = aggY_adj.slice();

  // Opsモデル：トレンド＋季節性
  const model = fitOpsModelTrendSeason_(smoothY);

  // 残差%は「確定月のみ」から作る（未確定月の途中実績に依存しにくく）
  const residualPctClosed = [];
  for (let i = 0; i < smoothY.length; i++) {
    const mStart = addMonths_(seriesStart, i);
    if (mStart > ctx.lastClosedMonthStart) continue; // 未確定は除外
    const f = model.fitted[i];
    if (f > 0) residualPctClosed.push(smoothY[i] / f - 1);
  }
  const residualPct = residualPctClosed.length ? residualPctClosed : smoothY.map((y, i) => (model.fitted[i] ? (y / model.fitted[i] - 1) : 0));

  const residP10 = percentile_(residualPct, 0.10);
  const residP50 = percentile_(residualPct, 0.50);
  const residP90 = percentile_(residualPct, 0.90);

  // Dev：固定加算（確度で調整）※運用シミュレーションには混ぜない
  const devFixedByMonth = readDevFixed12Months_(fy);
  // 背景SPOTの上限を作るため、BASE予測(P50)を先に計算
  const baseOnlyP50 = forecastByResidualQuantiles_(model, new Array(12).fill(0), { p10: residP10, p50: residP50, p90: residP90 }).p50;
  // SPOT背景：履歴SPOTから最低限の未知案件を拾う（上限クリップあり）
  const spotBackgroundByMonth = estimateSpotBackground12Months_(
    salesData.spotSeries48 || [],
    seriesStart,
    ctx.lastClosedMonthStart,
    baseOnlyP50,
    tuning
  );
  const spotFixedByMonth = devFixedByMonth.map((v, i) => Number(v || 0) + Number(spotBackgroundByMonth[i] || 0));

  // 要因（主観係数）※必要情報が揃った行だけ読む
  const factorsProduct = readFactorsProduct_(fy);
  const factorsClient = readFactorsClient_(fy);
  const opinions = readOpinions_(fy);

  // AI調査スコア（topic別 adjusted_score 平均）
  const aiScores = readAIResearchScores_();

  // 製品構成比：未確定月を避ける（直近の“確定済み12ヶ月”で重み計算）
  const productWeights = computeProductWeightsFromSalesInputClosed12_(fy, clientName, ctx);

  // 12ヶ月予測対象（FY開始月〜12ヶ月）
  const months = ctx.forecastMonths;

  // 線形回帰（参考）予測：季節性込みモデルのトレンド外挿（参考）
  const regTotal = [];
  for (let i = 0; i < 12; i++) {
    const t = 48 + (i + 1);
    const regOps = Math.max(0, (model.intercept + model.slope * t) * model.seasonalIndex[i % 12]);
    // 参考線（Linear）は既知案件(DEV)のみ加算し、背景SPOTは含めない
    regTotal.push(regOps + devFixedByMonth[i]);
  }

  // 「客観のみ」：残差分位点レンジ + SPOT固定（背景 + 既知DEV）
  const objOnly = forecastByResidualQuantiles_(model, spotFixedByMonth, { p10: residP10, p50: residP50, p90: residP90 });

  toastProgress_(ss, `STEP3/6: 残差からレンジの基礎（P10/P50/P90）を作成…`, 5);
  toastProgress_(ss, `STEP4/6: Dev固定加算 + 主観係数（製品/クライアント/意見）を準備…`, 6);

  toastProgress_(ss, `STEP5/6: Monte Carlo ${N_SIM}回（運用のみを揺らす + Dev固定加算）…`, 8);

  // 「混合」シミュレーション（Opsのみ揺らす）＋係数適用＋SPOT固定
  const mixed = forecastMonteCarloMixed_(model, spotFixedByMonth, {
    residualPct,
    factorsProduct,
    factorsClient,
    opinions,
    productWeights,
    aiScores,
    nSim: N_SIM,
    months,
    aiWeight: tuning.aiWeight,
    aiMaxAbsEffect: tuning.aiMaxAbsEffect
  });

  const opinionsSummaryTop = summarizeOpinionsTop_(opinions);
  const opinionsSummaryByMonth = summarizeOpinionsByMonth_(opinions, months);

  const totalActual48 = salesData.baseSeries48.map((v, i) => Number(v || 0) + Number((salesData.spotSeries48 || [])[i] || 0));
  const closedOffsets = new Set(ctx.closedForecastMonthOffsets || []);
  const sourceByMonth = months.map((_, i) => closedOffsets.has(i) ? 'actual_closed' : 'forecast_open');
  const actualClosedByMonth = months.map((_, i) => {
    const salesIdx = ctx.forecastMonthIndexesInSales[i];
    return (closedOffsets.has(i) && salesIdx >= 0) ? Number(totalActual48[salesIdx] || 0) : '';
  });

  for (let i = 0; i < months.length; i++) {
    if (sourceByMonth[i] !== 'actual_closed') continue;
    const a = Number(actualClosedByMonth[i] || 0);
    objOnly.p10[i] = a; objOnly.p50[i] = a; objOnly.p90[i] = a;
    mixed.p10[i] = a; mixed.p50[i] = a; mixed.p90[i] = a;
    regTotal[i] = a;
  }

  if (factorsProduct.length > 0 && mixed.diagnostics && mixed.diagnostics.kProdByMonth && mixed.diagnostics.kProdByMonth.every(k => Math.abs(Number(k || 1) - 1) < 1e-9)) {
    throw new Error('FACTORS_PRODUCT に有効行がありますが、kProd が全月1.0です。製品名キーの整合を確認してください。');
  }

  return {
    fy,
    clientName,
    months,
    objOnly,
    mixed,
    mixedDiagnostics: mixed.diagnostics || null,
    regTotal,
    devFixedByMonth,
    spotBackgroundByMonth,
    spotFixedByMonth,
    opinionsSummaryTop,
    opinionsSummaryByMonth,
    sourceByMonth,
    actualClosedByMonth,
    modelInfo: { residP10, residP50, residP90, slope: model.slope, intercept: model.intercept },
    aiScores
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

  const start = new Date(fy - 1, 3, 1);
  const end = new Date(fy, 2, 1);

  sh.getRange(1, 1).setValue(`FY${fy} 売上予測（${client} / ${fmtYM_(start)} 〜 ${fmtYM_(end)}）`);
  sh.getRange(1, 1, 1, 6).merge();
  sh.getRange(1, 1).setFontSize(16).setFontWeight('bold');
  sh.setFrozenRows(2);
  sh.getRange(1, 1, sh.getMaxRows(), 12).setHorizontalAlignment('left');

  // 上部サマリー（要点表示）
  sh.getRange(3, 1).setValue('予測の見方（要点）').setFontWeight('bold');
  sh.getRange(3, 2).setValue('Baseline(P50)=中心値 / Downside(P10)=下振れ目安 / Upside(P90)=上振れ目安。\n予測根拠は「トレンド＋季節性＋シミュレーション」で算出しています。');
  sh.getRange(3, 2, 1, 5).merge();
  sh.getRange(3, 2).setWrap(true);

  // 上部：担当者所感要約
  sh.getRange(4, 1).setValue('担当者所感（OPINION）');
  sh.getRange(4, 1).setFontWeight('bold');
  sh.getRange(4, 2).setValue(result.opinionsSummaryTop || '（未入力）');
  sh.getRange(4, 2, 1, 5).merge();
  sh.getRange(4, 2).setWrap(true);

  // AI調査スコア要約
  const ai = result.aiScores || { Market: 0, Competitor: 0, Channel: 0, DX: 0 };
  const aiSummary = `Market: ${ai.Market} / Competitor: ${ai.Competitor} / Channel: ${ai.Channel} / DX: ${ai.DX}`;
  sh.getRange(5, 1).setValue('AI調査スコア（adjusted）');
  sh.getRange(5, 1).setFontWeight('bold');
  sh.getRange(5, 2).setValue(aiSummary || '（未実施）');
  sh.getRange(5, 2, 1, 5).merge();
  sh.getRange(5, 2).setWrap(true);

  // 既存チャート削除（重なり防止）
  sh.getCharts().forEach(c => sh.removeChart(c));

  let row = 7;

  // ===== セクション1：混合 =====
  row = writeSectionBlock_(sh, row, {
    label: '過去売上（客観）と担当者情報（主観）を混合させたシミュレーション予測',
    labelBg: COLOR_MIX_LABEL,
    months: result.months,
    series: result.mixed,
    regTotal: result.regTotal,
    chartTitle: `混合：FY${fy} 月次予測レンジ（${client} / P10-P50-P90 + 回帰）`,
    spotFixedByMonth: result.spotFixedByMonth,
    devFixedByMonth: result.devFixedByMonth,
    spotBackgroundByMonth: result.spotBackgroundByMonth
  });

  row += 2;

  // ===== セクション2：客観のみ =====
  row = writeSectionBlock_(sh, row, {
    label: '過去売上のみ（客観）によるシミュレーション予測',
    labelBg: COLOR_OBJ_LABEL,
    months: result.months,
    series: result.objOnly,
    regTotal: result.regTotal,
    chartTitle: `客観のみ：FY${fy} 月次予測レンジ（${client} / P10-P50-P90 + 回帰）`,
    spotFixedByMonth: result.spotFixedByMonth,
    devFixedByMonth: result.devFixedByMonth,
    spotBackgroundByMonth: result.spotBackgroundByMonth
  });

  row += 2;

  // 参考：内訳（P50比較）
  sh.getRange(row, 1).setValue('（参考）内訳とメモ（P50比較）');
  sh.getRange(row, 1).setFontWeight('bold');
  sh.getRange(row, 1, 1, 6).merge();
  row++;

  sh.getRange(row, 1).setValue('※「運用(Ops)」はBASEのトレンド＋季節性から推定し、レンジは残差/シミュレーションで作ります。「SPOT固定」は背景SPOT（未知）+ DEV_SPOT（既知）を加算します。');
  sh.getRange(row, 1, 1, 6).merge();
  sh.getRange(row, 1).setFontColor('#666666').setFontSize(10);
  row++;

  const hdr = ['Month', 'ActualClosed', 'ForecastSource', '運用(Ops)P50（客観のみ）', '運用(Ops)P50（混合）', 'SPOT固定（背景+DEV）', 'Total P50（客観のみ）', 'Total P50（混合）', '差分(Mixed-Objective)', 'OPINIONS要約'];
  sh.getRange(row, 1, 1, hdr.length).setValues([hdr]).setBackground(COLOR_HEADER).setFontWeight('bold');
  row++;

  const spotFixed = result.spotFixedByMonth || result.devFixedByMonth;
  const rows = result.months.map((m, i) => {
    const objP50 = result.objOnly.p50[i];
    const mixP50 = result.mixed.p50[i];
    const spotVal = spotFixed[i];
    const opsObj = Math.max(0, objP50 - spotVal);
    const opsMix = Math.max(0, mixP50 - spotVal);
    return [
      fmtYM_(m),
      result.actualClosedByMonth ? (result.actualClosedByMonth[i] || '') : '',
      result.sourceByMonth ? (result.sourceByMonth[i] || 'forecast_open') : 'forecast_open',
      opsObj,
      opsMix,
      spotVal,
      objP50,
      mixP50,
      mixP50 - objP50,
      result.opinionsSummaryByMonth[i] || ''
    ];
  });

  sh.getRange(row, 1, rows.length, hdr.length).setValues(rows);
  sh.getRange(row, 2, rows.length, 1).setNumberFormat('¥#,##0');
  sh.getRange(row, 4, rows.length, 6).setNumberFormat('¥#,##0');
  sh.getRange(row, 10, rows.length, 1).setWrap(true);
  row += rows.length + 2;

  // ===== セクション3：三角測量（手法比較） =====
  sh.getRange(row, 1).setValue('Triangulation View（手法比較）');
  sh.getRange(row, 1, 1, 6).merge();
  sh.getRange(row, 1).setBackground('#d9e1f2').setFontWeight('bold');
  row++;

  const triHdr = ['比較軸', 'Linear Regression', 'Objective-only (P50)', 'Mixed (P50)', 'Mixed-Objective', 'Mixed-Linear'];
  const sumReg = sumArr_(result.regTotal);
  const sumObj = sumArr_(result.objOnly.p50);
  const sumMix = sumArr_(result.mixed.p50);
  const triAnnual = ['年度合計', sumReg, sumObj, sumMix, sumMix - sumObj, sumMix - sumReg];
  sh.getRange(row, 1, 1, triHdr.length).setValues([triHdr]).setBackground(COLOR_HEADER).setFontWeight('bold');
  row++;
  sh.getRange(row, 1, 1, triAnnual.length).setValues([triAnnual]);
  sh.getRange(row, 2, 1, triAnnual.length - 1).setNumberFormat('¥#,##0');
  row += 2;

  const triMonthHdr = ['Month', 'Linear', 'Objective P50', 'Mixed P50', 'Mixed-Objective', 'Mixed-Linear'];
  sh.getRange(row, 1, 1, triMonthHdr.length).setValues([triMonthHdr]).setBackground(COLOR_HEADER).setFontWeight('bold');
  row++;
  const triMonthRows = result.months.map((m, i) => [
    fmtYM_(m),
    result.regTotal[i],
    result.objOnly.p50[i],
    result.mixed.p50[i],
    result.mixed.p50[i] - result.objOnly.p50[i],
    result.mixed.p50[i] - result.regTotal[i]
  ]);
  sh.getRange(row, 1, triMonthRows.length, triMonthHdr.length).setValues(triMonthRows);
  sh.getRange(row, 2, triMonthRows.length, triMonthHdr.length - 1).setNumberFormat('¥#,##0');
  row += triMonthRows.length + 2;

  // ===== セクション4：入力パラメータの影響可視化 =====
  sh.getRange(row, 1).setValue('入力パラメータの影響（目安）');
  sh.getRange(row, 1, 1, 8).merge();
  sh.getRange(row, 1).setBackground('#e2f0d9').setFontWeight('bold');
  row++;

  const d = result.mixedDiagnostics || {};
  const kProd = d.kProdByMonth || new Array(12).fill(1);
  const kClient = d.kClientByMonth || new Array(12).fill(1);
  const kOpinion = d.kOpinionP50ByMonth || new Array(12).fill(1);
  const kAI = d.kAIByMonth || new Array(12).fill(1);
  const opsBase = d.opsBaseByMonth || new Array(12).fill(0);

  const infHdr = ['Month', 'Ops基礎', 'kProd', 'kClient', 'kOpinion(P50)', 'kAI', 'Dev固定', '混合P50', '客観P50', '差分(混合-客観)'];
  sh.getRange(row, 1, 1, infHdr.length).setValues([infHdr]).setBackground(COLOR_HEADER).setFontWeight('bold');
  row++;
  const infRows = result.months.map((m, i) => [
    fmtYM_(m),
    opsBase[i] || 0,
    kProd[i] || 1,
    kClient[i] || 1,
    kOpinion[i] || 1,
    kAI[i] || 1,
    result.devFixedByMonth[i] || 0,
    result.mixed.p50[i] || 0,
    result.objOnly.p50[i] || 0,
    (result.mixed.p50[i] || 0) - (result.objOnly.p50[i] || 0)
  ]);
  sh.getRange(row, 1, infRows.length, infHdr.length).setValues(infRows);
  sh.getRange(row, 2, infRows.length, 1).setNumberFormat('¥#,##0');
  sh.getRange(row, 3, infRows.length, 4).setNumberFormat('0.000');
  sh.getRange(row, 7, infRows.length, 4).setNumberFormat('¥#,##0');
}

/** セクションブロック（表＋グラフ） */
function writeSectionBlock_(sh, startRow, opt) {
  let r = startRow;

  // ラベル
  sh.getRange(r, 1).setValue(opt.label);
  sh.getRange(r, 1, 1, 6).merge();
  sh.getRange(r, 1).setBackground(opt.labelBg).setFontWeight('bold');
  r++;

  // 年度合計（B=Downside / C=Baseline / D=Upside）
  const sumPos = sumArr_(opt.series.p90);
  const sumNeu = sumArr_(opt.series.p50);
  const sumNeg = sumArr_(opt.series.p10);
  const sumReg = sumArr_(opt.regTotal);
  const sumRange = sumPos - sumNeg;

  const annualHdr = ['年度合計（シミュレーション予測）', 'Downside(P10)', 'Baseline(P50)', 'Upside(P90)', 'Linear Regression', 'Range(P90-P10)'];
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
  const hdr = ['Month', 'Downside(P10)', 'Baseline(P50)', 'Upside(P90)', 'Linear Regression', 'Range(P90-P10)'];
  sh.getRange(r, 1, 1, hdr.length).setValues([hdr]).setBackground(COLOR_HEADER).setFontWeight('bold');
  const monthTableHeaderRow = r;
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
  sh.getRange(r - 1, 2).setNote('【Downside(P10)】\nシミュレーション結果の下位10%点（=10パーセンタイル）。\n下振れ側の目安です。');
  sh.getRange(r - 1, 3).setNote('【Baseline(P50)】\nシミュレーション結果の中央値（=50パーセンタイル）。\n最も参照すべき“中心”の目安です。');
  sh.getRange(r - 1, 4).setNote('【Upside(P90)】\nシミュレーション結果の上位10%点（=90パーセンタイル）。\n上振れ側の目安です。');
  sh.getRange(r - 1, 5).setNote('【Linear Regression】\n過去売上（ならした推移）に単純な直線を当てて将来を外挿した参考値です。\n季節性も考慮したトレンド外挿を行います。');
  sh.getRange(r - 1, 6).setNote('【Range(P90-P10)】\nUpside(P90)からDownside(P10)を引いた幅です。\n不確実性（どれくらいブレうるか）の大きさを表します。');

  // BASE/SPOT分離（SPOTは背景SPOT + DEV固定の合算）
  r += table.length + 2;
  sh.getRange(r, 1).setValue('Scenario Split（BASE / SPOT）').setFontWeight('bold');
  sh.getRange(r, 1, 1, 6).merge();
  sh.getRange(r, 1).setBackground('#e2f0d9');
  r++;

  const spotFixed = opt.spotFixedByMonth || opt.devFixedByMonth || new Array(12).fill(0);
  const splitHdr = [
    'Month',
    'Downside_BASE', 'Downside_SPOT',
    'Baseline_BASE', 'Baseline_SPOT',
    'Upside_BASE', 'Upside_SPOT'
  ];
  sh.getRange(r, 1, 1, splitHdr.length).setValues([splitHdr]).setBackground(COLOR_HEADER).setFontWeight('bold');
  r++;

  const splitRows = opt.months.map((m, i) => {
    const spot = Number(spotFixed[i] || 0);
    const neg = Number(opt.series.p10[i] || 0);
    const neu = Number(opt.series.p50[i] || 0);
    const pos = Number(opt.series.p90[i] || 0);
    return [
      fmtYM_(m),
      Math.max(0, neg - spot), spot,
      Math.max(0, neu - spot), spot,
      Math.max(0, pos - spot), spot
    ];
  });
  sh.getRange(r, 1, splitRows.length, splitHdr.length).setValues(splitRows);
  sh.getRange(r, 2, splitRows.length, splitHdr.length - 1).setNumberFormat('¥#,##0');
  sh.getRange(r - 1, 2).setNote('BASEは「シナリオ値 - SPOT固定（背景SPOT + DEV固定）」を表示しています。');
  sh.getRange(r - 1, 3).setNote('SPOTは「背景SPOT + DEV_SPOT（既知案件）」の合算表示です。');

  // グラフ：Month + Neg + Neu + Pos + Reg（A〜E）
  const chartRange = sh.getRange(monthTableHeaderRow, 1, table.length + 1, 5);

  const chartRow = startRow + 1;
  const chartCol = 8; // H列開始

  // 凡例テキスト（邪魔にならない小さめ）
  sh.getRange(chartRow - 1, chartCol).setValue('Legend: red=Downside(P10) / yellow=Baseline(P50) / blue=Upside(P90) / gray=Linear Regression')
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
    // 色：Downside=赤 / Baseline=黄 / Upside=青 / 回帰=灰
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
  sh.clear({ contentsOnly: true });
  sh.clearFormats();
  sh.setColumnWidth(1, 130);
  sh.setColumnWidth(2, 340);
  sh.setColumnWidth(3, 660);

  const C_A = '#d9e8fb';
  const C_B = '#d9ead3';
  const C_AUTO = '#d9e8fb';
  const C_USER = '#fff2cc';
  const C_OUT = '#f4cccc';
  const C_VER = '#d9ead3';

  sh.getRange(1, 1).setValue(`売上予測ツール ガイド（v${VERSION}）`).setFontSize(16).setFontWeight('bold');
  sh.getRange(2, 1, 1, 3).setValues([['分類', 'Forecast Agentボタンの手順', 'ボタン説明']]).setBackground(COLOR_HEADER).setFontWeight('bold');

  const aRows = [
    ['A-予測', 'A-1 初期セットアップ', '初回のみ。クライアント/FY/担当者を設定。'],
    ['A-予測', 'A-2 売上データを取り込む', '案件一覧を SALES_INPUT_MONTHLY へ取り込み。'],
    ['A-予測', 'A-3 予測用に売上データを加工', 'SALES_INPUT_MONTHLY のデータを SALES で48か月横持ち（BASE/SPOT）に集計。'],
    ['A-予測', 'A-4 製品ごとの動向を入力', 'FACTORS_PRODUCT（全製品）へ入力。'],
    ['A-予測', 'A-5 クライアント動向を入力', 'FACTORS_CLIENT へ入力。'],
    ['A-予測', 'A-6 担当者意見を入力', 'OPINIONS へ入力（担当者全員分）。'],
    ['A-予測', 'A-7 開発/スポット要因を入力', 'DEV_SPOT へ入力。'],
    ['A-予測', 'A-8 AI調査を取り込む', '生成されたプロンプトをGemへ貼り付け、返却結果を AI_RESEARCH_PROMPT!D2 に全文貼り付け。'],
    ['A-予測', 'A-9 予測を実行', 'OUTPUT / FORECAST_REPORT を更新（実行前に注意ロジックで1件ずつ確認）。'],
    ['A-予測', 'A-10 予測ダッシュボードを更新', 'DASHBOARD を更新。']
  ];
  sh.getRange(3, 1, aRows.length, 3).setValues(aRows).setBackground(C_A);

  const bRows = [
    ['B-事後検証', 'B-1 検証用に実績データを取り込み', '実績を ACTUAL_EVAL_MONTHLY に取り込み（BASE/SPOT判定つき）。'],
    ['B-事後検証', 'B-2 検証レポートを更新', 'EVAL_LOG と EVAL_COMPARE_MONTHLY を更新。'],
    ['B-事後検証', 'B-3 検証インサイトを更新', 'EVAL_INSIGHTS に外れ要因と次アクションを整理。']
  ];
  sh.getRange(13, 1, bRows.length, 3).setValues(bRows).setBackground(C_B);

  sh.getRange(17, 1, 1, 3).setValues([['シート分類', 'シート名', 'シート説明']]).setBackground(COLOR_HEADER).setFontWeight('bold');
  const links = [
    ['自動入力用', SHEETS.CONFIG, '設定（クライアント/FY/担当者）'],
    ['自動入力用', SHEETS.SALES_INPUT_MONTHLY, '予測入力（月次案件一覧）'],
    ['自動入力用', SHEETS.SALES, '予測用集計（48ヶ月横持ち / BASE・SPOT）'],
    ['自動入力用', SHEETS.AI_RESEARCH_PROMPT, 'AI調査テンプレート兼貼り付け'],
    ['ユーザ入力用', SHEETS.FACTORS_PRODUCT, '製品要因入力'],
    ['ユーザ入力用', SHEETS.FACTORS_CLIENT, 'クライアント要因入力'],
    ['ユーザ入力用', SHEETS.OPINIONS, '担当者意見入力'],
    ['ユーザ入力用', SHEETS.DEV_SPOT, '開発/スポット要因入力'],
    ['出力用', SHEETS.OUTPUT, '予測出力'],
    ['出力用', SHEETS.FORECAST_REPORT, '予測レポート'],
    ['出力用', SHEETS.DASHBOARD, 'ダッシュボード'],
    ['事後検証用', SHEETS.ACTUAL_EVAL_MONTHLY, '検証実績（月次案件一覧）'],
    ['事後検証用', SHEETS.EVAL_COMPARE_MONTHLY, '予測/実績比較（BASE・SPOT）'],
    ['事後検証用', SHEETS.EVAL_LOG, '予測検証ログ'],
    ['事後検証用', SHEETS.EVAL_INSIGHTS, '検証インサイト']
  ];
  setGuideLinkTable_(sh, 18, links);

  const last = 18 + links.length;
  sh.getRange(last + 2, 1).setValue('運用補足').setFontWeight('bold');
  sh.getRange(last + 3, 1, 10, 1).setValues([
    ['・A-予測は「予測作成」、B-事後検証は「外れ理由学習」のための手順です。'],
    ['・織り込める要素: BASE履歴トレンド/季節性、主観入力（製品/クライアント/意見）、AI調査、DEV_SPOT。'],
    ['・SPOTは「背景SPOT（未知）+ DEV_SPOT（既知）」として別枠で加算し、BASEトレンドとは分離します。'],
    ['・A-9 実行時に未入力/型不正/影響過大の入力は、階層アラートで1件ずつ表示します。'],
    ['・対応できない範囲: 突発イベントの完全再現、外部制度変更の即時反映、全案件の網羅。'],
    ['・主なリスク: 人手入力の保守/楽観バイアス、AI情報の鮮度・偏り、外部データ欠損。'],
    ['・予測は意思決定補助であり確定値ではありません。P10/P50/P90レンジで判断してください。'],
    ['・CONFIGの「モデル調整パラメータ」は管理者向けです。むやみに変更しないでください。'],
    ['・検証(B-1〜B-3)を毎月回し、ズレをEVAL_INSIGHTSに蓄積してください。'],
    ['・内部管理シート（RUN_LOG/FORECAST_SNAPSHOT/PROCESS_STATUS など）は初期状態で非表示です。']
  ]);

  ss.setActiveSheet(sh);
  safeMoveSheet_(ss, sh, 1);
}

function buildCONFIG_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getOrCreateSheet_(ss, SHEETS.CONFIG);
  sh.clear({ contentsOnly: true });
  sh.clearFormats();

  sh.setColumnWidth(1, 312);
  sh.setColumnWidth(2, 504);

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
  sh.getRange('B10').setBackground(COLOR_OBJECTIVE);

  sh.getRange('A2').setNote('外部実績シート（*YYYY_actual_value）のAO列にあるメーカー名と一致させます。');
  sh.getRange('A3').setNote('例：FY2026 は 2026/04/01〜2027/03/31 の12ヶ月です（4月開始・3月決算）。');
  sh.getRange('A5').setNote('シミュレーションは1000回試行し、レンジ（P10/P50/P90）を出します。単純な一発計算より「ブレ幅」を扱えるのがメリットです。');
  sh.getRange('A10').setNote('シミュレーションに関与する担当者の苗字をカンマ区切りで記載します。A-6では全員分の意見が必須です。');

  const infoStart = 12;
  const infoHdr = [['入力パラメータ', '計算上の扱い（要点）']];
  const infoRows = [
    ['客観ベース（Ops）', 'SALESのBASE 48ヶ月のみでトレンド+12ヶ月季節性を推定。SPOTは背景成分として別枠で加算します。'],
    ['残差シミュレーション', `過去残差をランダム抽出して ${N_SIM} 回シミュレーション。P10/P50/P90 を算出。`],
    ['製品別要因（FACTORS_PRODUCT）', 'kProd = 1 + Σ(製品構成比×累積step)。月次で乗算。'],
    ['クライアント要因（FACTORS_CLIENT）', 'kClient = 1 + 累積step。月次で乗算。'],
    ['担当者意見（OPINIONS）', '担当者ごとの (1 + step×confidence) を合成（内部では±5%の小さな揺らぎあり）。'],
    ['AI調査（AI_RESEARCH_STRUCTURED）', 'kAI = 1 + 0.001 × (Market+Competitor+Channel+DX)。例: 合計+30 ⇒ +3%。'],
    ['固定額（DEV_SPOT）', 'amount×confidence を月次で固定加算（背景SPOTと合算してSPOT固定成分として扱う）。']
  ];
  sh.getRange(infoStart, 1, 1, 2).setValues(infoHdr).setBackground(COLOR_HEADER).setFontWeight('bold');
  sh.getRange(infoStart + 1, 1, infoRows.length, 2).setValues(infoRows);
  sh.getRange(infoStart + 1, 2, infoRows.length, 1).setWrap(true);

  // A-9 注意ロジック（実装定数と連動）
  const warnStart = infoStart + 1 + infoRows.length + 2;
  const a9Hdr = [['A-9 実行前チェック', '閾値と挙動（定数連動）']];
  const a9Rows = [
    ['Step 警告', `|Step| >= ${Math.round(STEP_WARN_THRESHOLD * 100)}% ：警告表示（OKで続行 / Cancelで中断）`],
    ['Step 強警告', `|Step| >= ${Math.round(STEP_STRONG_THRESHOLD * 100)}% ：強い警告表示（OKで続行 / Cancelで中断）`],
    ['Step 極端値', `|Step| >= ${Math.round(STEP_BLOCK_THRESHOLD * 100)}% ：強い警告表示（OKで続行 / Cancelで中断）`],
    ['合成係数 警告', `kTotal < ${K_TOTAL_WARN_MIN.toFixed(2)} または > ${K_TOTAL_WARN_MAX.toFixed(2)} ：警告表示（OKで続行 / Cancelで中断）`],
    ['合成係数 極端値', `kTotal < ${K_TOTAL_BLOCK_MIN.toFixed(2)} または > ${K_TOTAL_BLOCK_MAX.toFixed(2)} ：強い警告表示（OKで続行 / Cancelで中断）`],
    ['解消手順', '1) 表示された1件を修正 → 2) A-9を再実行 → 3) 次の注意が出たら同様に修正（同時に複数表示しない）']
  ];
  sh.getRange(warnStart, 1, 1, 2).setValues(a9Hdr).setBackground(COLOR_HEADER).setFontWeight('bold');
  sh.getRange(warnStart + 1, 1, a9Rows.length, 2).setValues(a9Rows);
  sh.getRange(warnStart + 1, 2, a9Rows.length, 1).setWrap(true);

  // 管理者が参照しやすいよう固定位置に配置（B32:B36を実値として利用）
  const tuneStart = 31;
  const tuneHdr = [['モデル調整パラメータ', '値（必要時のみ調整）']];
  const tuneRows = [
    ['SPOT_BG_SHRINK（背景SPOT縮小率）', SPOT_BG_SHRINK],
    ['SPOT_BG_FLOOR_RATE（背景SPOT最低保証率）', SPOT_BG_FLOOR_RATE],
    ['SPOT_BG_CAP_RATE（背景SPOT上限/BaseP50比）', SPOT_BG_CAP_RATE],
    ['AI_WEIGHT（AI係数重み）', AI_WEIGHT_DEFAULT],
    ['AI_MAX_ABS_EFFECT（AI係数上限）', AI_MAX_ABS_EFFECT]
  ];
  sh.getRange(tuneStart, 1, 1, 2).setValues(tuneHdr).setBackground(COLOR_HEADER).setFontWeight('bold');
  sh.getRange(tuneStart + 1, 1, tuneRows.length, 2).setValues(tuneRows);
  sh.getRange(tuneStart + 1, 2, tuneRows.length, 1).setNumberFormat('0.0000');
  sh.getRange(tuneStart, 1, 1, 2).setNote('A-9 実行時にこの値を参照します。極端な変更は予測を不安定にします。');
}

function buildSALES_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getOrCreateSheet_(ss, SHEETS.SALES);
  sh.clear({ contentsOnly: true });
  sh.clearFormats();

  // 48ヶ月分（B〜AW=48列）
  ensureSheetHasColumns_(sh, 1 + 48);

  sh.setColumnWidth(1, 180);
  for (let c = 2; c <= 49; c++) sh.setColumnWidth(c, 110);

  sh.getRange(1, 1).setValue('Category');
  sh.getRange(1, 1).setBackground(COLOR_HEADER).setFontWeight('bold');

  sh.setFrozenRows(1);
  sh.setFrozenColumns(1);

  sh.getRange(1, 1).setNote('BASE / SPOT のカテゴリ行です。');
  sh.getRange(1, 2).setNote('過去4年（48ヶ月）の月次売上（客観データ）です。');
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

  sh.getRange(1, 1).setNote('初期設定で入力した表記（CONFIGシートの担当者と同じ表記）を選択してください。');
  sh.getRange(1, 2).setNote('SALES_INPUT_MONTHLYから取得した製品一覧に合わせて自動展開されます。');
  sh.getRange(1, 3).setNote('この日付「以降」に影響が出る想定で入力します。');
  sh.getRange(1, 4).setNote('増減率（%）です。例：-30% = 今後30%減りそう。\n入力は 0%/±5%刻みを推奨。');
  sh.getRange(1, 5).setNote('根拠を短く（例：競合参入、契約更改、規制変更）。\nこの列は予測根拠の説明に使われます。');

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

  sh.getRange(1, 1).setNote('初期設定で入力した表記（CONFIGシートの担当者と同じ表記）を選択してください。');
  sh.getRange(1, 2).setNote('この日付「以降」に影響が出る想定で入力します。');
  sh.getRange(1, 3).setNote('増減率（%）です。例：-10% = 予算圧縮で10%減りそう。\n※入力値はそのまま直に固定反映せず、シミュレーション内で扱われます。');
  sh.getRange(1, 4).setNote('根拠を短く（例：予算圧縮、体制変更など）。\n未入力だと判断根拠が追跡しづらくなります。');

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

  sh.getRange(1, 1).setNote('初期設定で入力した表記（CONFIGシートの担当者と同じ表記）を選択してください。A列はプルダウンです。');
  sh.getRange(1, 2).setNote('この日付「以降」に意見の影響が出る想定で入力します。');
  sh.getRange(1, 3).setNote('増減率（%）です。例：+20% = 今後20%増えそう。\n※意見はそのまま固定反映されず、シミュレーションでランダムに活用されます。');
  sh.getRange(1, 4).setNote('信頼度（0..1）。1に近いほど「この意見を強く信用してよい」として影響が強まります。');
  sh.getRange(1, 5).setNote('所感を短く（例：プロモ減、資材整理、体制変更など）。\nここは必ず入力してください。');

  sh.getRange('D2:D').setDataValidation(SpreadsheetApp.newDataValidation().requireNumberBetween(0, 1).build());
  sh.setFrozenRows(1);
}

function buildDEV_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getOrCreateSheet_(ss, SHEETS.DEV_SPOT);
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
  sh.getRange(1, 1).setNote('初期設定で入力した表記（CONFIGシートの担当者と同じ表記）を選択してください。');
  sh.getRange(1, 2).setNote('この日付の月に固定売上として加算します（開発案件/スポット要因）。');
  sh.getRange(1, 3).setNote('案件名（またはスポット要因名）を短く。');
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

/** ====== テンプレ整形（A-4〜A-7で呼ぶ） ====== */
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
    const rows = missing.map(p => [p, defaultDate, '', '', '']);
    sh.getRange(startRow, 1, rows.length, 5).setValues(rows);
  }

  const maxRow = Math.max(sh.getLastRow(), 2);
  sh.getRange(2, 1, maxRow - 1, 5).setBackground(COLOR_SUBJECTIVE);

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

  sh.getRange(2, 4, sh.getMaxRows() - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireNumberBetween(0, 1).setAllowInvalid(true).build()
  );
}

function ensureDevTemplate_(sh, people, defaultDate) {
  if (!sh) throw new Error('DEV_SPOTがありません。');

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
  if (lastRow < 2) throw new Error('SALESに製品行がありません。A-2で取り込み、または手入力してください。');

  const expectedMonths = 48;
  const startCol = 2;
  const endCol = startCol + expectedMonths - 1;
  if (sales.getLastColumn() < endCol) {
    throw new Error('SALESの月次列が48ヶ月分ありません。A-2 売上データを取り込む を実行してください。');
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

  // FACTORS / OPINIONS / DEV_SPOT：明らかに変な行があれば停止（未完成行は“無視”＝エラーにはしない）
  validateFactorsSheet_(SHEETS.FACTORS_PRODUCT, { cols: 5, mode: 'product' });
  validateFactorsSheet_(SHEETS.FACTORS_CLIENT, { cols: 4, mode: 'client' });
  validateOpinionsSheet_(people);
  validateDevSheet_();
}

function validateRequiredUserInputsOrThrow_() {
  const people = getPeopleListFromConfig_();
  const missingPeople = findMissingPeopleOpinionsByValidRows_(people);
  if (missingPeople.length > 0) {
    throw new Error(`OPINIONSに担当者意見が不足しています。未入力: ${missingPeople.join(', ')}`);
  }

  const fp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.FACTORS_PRODUCT);
  if (!fp || fp.getLastRow() < 2) throw new Error('FACTORS_PRODUCT の入力行がありません。A-4 を実行してください。');

  const hasReason = fp.getRange(2, 5, fp.getLastRow() - 1, 1).getValues().some(r => String(r[0] || '').trim());
  if (!hasReason) throw new Error('FACTORS_PRODUCT のReasonが未入力です。最低1件入力してください。');
}

/**
 * A-9 実行前の階層アラート（1件ずつ解消させる）
 * 1) Stepの極端値
 * 2) 主観/AI合成係数の過大影響
 */
function runHierarchicalA9AlertsOrThrow_(fy) {
  // 閾値は buildCONFIG_ の「A-9 実行前チェック」表示と同一定数を参照
  const issue =
    findFirstExtremeStepIssue_(fy) ||
    findFirstExtremeDevSpotIssue_(fy) ||
    findFirstExtremeMultiplierIssue_(fy);

  if (!issue) return;

  const ui = SpreadsheetApp.getUi();
  const title = issue.level === 'high' ? '注意（影響がかなり大きい入力）' : '注意（影響が大きい入力）';
  const buttons = ui.ButtonSet.OK_CANCEL;
  const res = ui.alert(title, issue.message, buttons);
  if (res !== ui.Button.OK) throw new Error('ユーザーがA-9実行を中断しました（入力内容を見直してください）。');
}

function findFirstExtremeStepIssue_(fy) {
  const factorsProduct = readFactorsProduct_(fy);
  const factorsClient = readFactorsClient_(fy);
  const opinions = readOpinions_(fy);

  const checks = [];
  factorsProduct.forEach(x => checks.push({ src: 'FACTORS_PRODUCT', who: x.person, month: x.month, step: x.step }));
  factorsClient.forEach(x => checks.push({ src: 'FACTORS_CLIENT', who: x.person, month: x.month, step: x.step }));
  opinions.forEach(x => checks.push({ src: 'OPINIONS', who: x.person, month: x.month, step: x.step }));

  for (let i = 0; i < checks.length; i++) {
    const c = checks[i];
    const abs = Math.abs(Number(c.step || 0));
    if (!isFinite(abs)) continue;

    const ym = c.month ? fmtYM_(c.month) : '-';
    const pct = `${Math.round((c.step || 0) * 100)}%`;
    const detail = `シート: ${c.src} / 担当: ${c.who || '-'} / 月: ${ym} / Step: ${pct}`;

    if (abs >= STEP_BLOCK_THRESHOLD) {
      return {
        level: 'high',
        message: `Stepが極端です（±100%以上）。\n\n${detail}\n\nプロモーション終了などで意図した入力ならOKで続行できます。修正する場合はキャンセルしてください。`
      };
    }
    if (abs >= STEP_STRONG_THRESHOLD) {
      return {
        level: 'warn',
        message: `Stepが大きく、予測に強く影響する可能性があります（±50%以上）。\n\n${detail}\n\n修正する場合はキャンセル、続行する場合はOKを押してください。`
      };
    }
    if (abs >= STEP_WARN_THRESHOLD) {
      return {
        level: 'warn',
        message: `Stepがやや大きめです（±30%以上）。\n\n${detail}\n\n修正する場合はキャンセル、続行する場合はOKを押してください。`
      };
    }
  }

  return null;
}

function findFirstExtremeDevSpotIssue_(fy) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sales = ss.getSheetByName(SHEETS.SALES);
  const devFixed = readDevFixed12Months_(fy);
  if (!sales || !devFixed || devFixed.length === 0) return null;

  const salesData = readSales48Months_(sales);
  const base48 = salesData.baseSeries48 || [];
  const baseAvg = base48.length ? (sumArr_(base48) / Math.max(1, base48.length)) : 0;
  if (!isFinite(baseAvg) || baseAvg <= 0) return null;

  const start = new Date(fy, 3, 1);
  for (let i = 0; i < 12; i++) {
    const v = Number(devFixed[i] || 0);
    if (!isFinite(v) || v <= 0) continue;
    const ym = fmtYM_(addMonths_(start, i));
    const ratio = v / baseAvg;
    if (ratio >= 1.2) {
      return {
        level: 'high',
        message: `DEV/SPOT固定が大きい月があります（${ym} / ${Math.round(v).toLocaleString()}円, BASE平均比 ${(ratio * 100).toFixed(1)}%）。\n\n意図した大型案件ならOKで続行、修正する場合はキャンセルしてください。`
      };
    }
    if (ratio >= 0.8) {
      return {
        level: 'warn',
        message: `DEV/SPOT固定がやや大きい月があります（${ym} / ${Math.round(v).toLocaleString()}円, BASE平均比 ${(ratio * 100).toFixed(1)}%）。\n\n意図した入力ならOKで続行、修正する場合はキャンセルしてください。`
      };
    }
  }
  return null;
}

function findFirstExtremeMultiplierIssue_(fy) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sales = ss.getSheetByName(SHEETS.SALES);
  if (!sales) return null;

  const salesData = readSales48Months_(sales);
  const monthlyByProduct = salesData.monthlyByProduct || [];
  if (monthlyByProduct.length === 0) return null;

  const productNames = salesData.productNames || [];
  const totalsByProduct = monthlyByProduct.map(arr => sumArr_(arr));
  const totalAll = sumArr_(totalsByProduct) || 1;
  const weights = new Map();
  for (let i = 0; i < productNames.length; i++) {
    weights.set(productNames[i], (totalsByProduct[i] || 0) / totalAll);
  }

  const months = [];
  const start = new Date(fy, 3, 1);
  for (let i = 0; i < 12; i++) months.push(addMonths_(start, i));

  const factorsProduct = readFactorsProduct_(fy);
  const factorsClient = readFactorsClient_(fy);
  const opinions = readOpinions_(fy);
  const tuning = readModelTuningFromConfig_();
  const ai = readAIResearchScores_();
  const aiTotal = (ai.Market || 0) + (ai.Competitor || 0) + (ai.Channel || 0) + (ai.DX || 0);
  const aiRaw = aiTotal * (isFinite(tuning.aiWeight) ? tuning.aiWeight : AI_WEIGHT_DEFAULT);
  const aiEff = Math.max(-(isFinite(tuning.aiMaxAbsEffect) ? tuning.aiMaxAbsEffect : AI_MAX_ABS_EFFECT), Math.min((isFinite(tuning.aiMaxAbsEffect) ? tuning.aiMaxAbsEffect : AI_MAX_ABS_EFFECT), aiRaw));
  const kAI = 1 + aiEff;

  for (let i = 0; i < 12; i++) {
    const m = months[i];
    const kProd = productFactorsMultiplier_(factorsProduct, m, weights);
    const kClient = clientFactorsMultiplier_(factorsClient, m);
    const kOpinion = opinionExpectedMultiplier_(opinions, m);
    const kTotal = kProd * kClient * kOpinion * kAI;
    const ym = fmtYM_(m);

    if (kTotal < K_TOTAL_BLOCK_MIN || kTotal > K_TOTAL_BLOCK_MAX) {
      return {
        level: 'high',
        message: `主観/AIの合成係数が極端です（${ym} / kTotal=${kTotal.toFixed(3)}）。\n\n意図した戦略変更ならOKで続行できます。修正する場合はキャンセルしてください。`
      };
    }
    if (kTotal < K_TOTAL_WARN_MIN || kTotal > K_TOTAL_WARN_MAX) {
      return {
        level: 'warn',
        message: `主観/AIの合成係数が大きめです（${ym} / kTotal=${kTotal.toFixed(3)}）。\n\n修正する場合はキャンセル、続行する場合はOKを押してください。`
      };
    }
  }

  return null;
}

function opinionExpectedMultiplier_(opinions, targetMonth) {
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
    const baseStep = isFinite(o.step) ? o.step : 0;
    const conf = isFinite(o.confidence) ? o.confidence : 0.7;
    k *= (1 + baseStep * conf);
  });
  return Math.max(0, k);
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
  if (!sh) throw new Error('OPINIONSシートがありません。A-6を実行してください。');

  const last = sh.getLastRow();
  if (last < 2) throw new Error('OPINIONSに入力行がありません。A-6を実行してください。');

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
    throw new Error(`OPINIONSに担当者全員の有効な入力がありません。\n未入力: ${missing.join(', ')}\nA-6で入力してください。`);
  }
}

function validateDevSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.DEV_SPOT);
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
    if (!dt) throw new Error(`DEV_SPOT!B${rowNum} の日付が不正です（yyyy/mm/dd 形式で入力してください）。`);

    const amt = toNumberSafe_(amountRaw);
    if (!isFinite(amt)) throw new Error(`DEV_SPOT!D${rowNum} の金額が数値として不正です（"${amountRaw}"）。`);
    if (amt < 0) throw new Error(`DEV_SPOT!D${rowNum} の金額が負の値です（${amt}）。`);

    const conf = Number(confRaw);
    if (!isFinite(conf) || conf < 0 || conf > 1) throw new Error(`DEV_SPOT!E${rowNum} の確度が不正です（0..1）。`);
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

function readModelTuningFromConfig_() {
  const out = {
    spotBgShrink: SPOT_BG_SHRINK,
    spotBgFloorRate: SPOT_BG_FLOOR_RATE,
    spotBgCapRate: SPOT_BG_CAP_RATE,
    aiWeight: AI_WEIGHT_DEFAULT,
    aiMaxAbsEffect: AI_MAX_ABS_EFFECT
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName(SHEETS.CONFIG);
  if (!cfg) return out;

  const getNum = (aCell, def) => {
    const v = Number(cfg.getRange(aCell).getValue());
    return isFinite(v) ? v : def;
  };

  out.spotBgShrink = Math.max(0, Math.min(1, getNum('B32', out.spotBgShrink)));
  out.spotBgFloorRate = Math.max(0, Math.min(1, getNum('B33', out.spotBgFloorRate)));
  out.spotBgCapRate = Math.max(0, Math.min(1, getNum('B34', out.spotBgCapRate)));
  out.aiWeight = Math.max(0, Math.min(0.01, getNum('B35', out.aiWeight)));
  out.aiMaxAbsEffect = Math.max(0, Math.min(0.50, getNum('B36', out.aiMaxAbsEffect)));
  return out;
}

function getProductNameListFromSales_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const salesInput = ss.getSheetByName(SHEETS.SALES_INPUT_MONTHLY);
  if (!salesInput) return [];
  const last = salesInput.getLastRow();
  if (last < 2) return [];
  const vals = salesInput.getRange(2, 3, last - 1, 1).getValues().map(r => String(r[0] || '').trim()).filter(Boolean);
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
  const startCol = 2; // B列〜
  const endCol = startCol + expectedMonths - 1; // 49

  const isComplete48 = (lastCol >= endCol);

  const productRows = Math.max(0, lastRow - 1);
  const data = [];

  if (productRows > 0) {
    const width = Math.min(lastCol, endCol);
    const vals = salesSheet.getRange(2, 1, productRows, width).getValues();
    vals.forEach(row => {
      const name = String(row[0] || '').trim();
      if (!name) return;
      const category = name.toUpperCase();
      // TOTAL行は表示用のため予測入力には含めない（BASE/SPOTのみを使用）
      if (category !== 'BASE' && category !== 'SPOT') return;
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
    productNames: data.map(x => x.productName),
    baseSeries48: (data.find(x => String(x.productName || '').toUpperCase() === 'BASE') || { monthly: new Array(expectedMonths).fill(0) }).monthly,
    spotSeries48: (data.find(x => String(x.productName || '').toUpperCase() === 'SPOT') || { monthly: new Array(expectedMonths).fill(0) }).monthly,
    headerMonths: readSalesHeaderMonths_(salesSheet, expectedMonths)
  };
}

function readSalesHeaderMonths_(salesSheet, expectedMonths) {
  const vals = salesSheet.getRange(1, 2, 1, expectedMonths).getValues()[0];
  const out = vals.map(v => toMonthStart_(v));
  if (out.some(v => !v)) throw new Error('SALESヘッダ月が解釈できません。A-3を再実行してください。');
  return out;
}

function getForecastContext_(fy, runDate, headerMonths) {
  const forecastStart = new Date(fy, 3, 1);
  const forecastEnd = new Date(fy + 1, 2, 1);
  const currentMonth = new Date(runDate.getFullYear(), runDate.getMonth(), 1);
  const lastClosedMonthStart = addMonths_(currentMonth, -1);

  const forecastMonths = [];
  for (let i = 0; i < 12; i++) forecastMonths.push(addMonths_(forecastStart, i));

  const historyMonthIndexes = [];
  for (let i = 0; i < headerMonths.length; i++) {
    const m = headerMonths[i];
    if (m < forecastStart) historyMonthIndexes.push(i);
  }

  const ymToIndex = new Map();
  for (let i = 0; i < headerMonths.length; i++) ymToIndex.set(fmtYM_(headerMonths[i]), i);

  const forecastMonthIndexesInSales = forecastMonths.map(m => {
    const idx = ymToIndex.get(fmtYM_(m));
    return Number.isInteger(idx) ? idx : -1;
  });

  const closedForecastMonthOffsets = [];
  const openForecastMonthOffsets = [];
  for (let i = 0; i < forecastMonths.length; i++) {
    const salesIdx = forecastMonthIndexesInSales[i];
    if (salesIdx >= 0 && forecastMonths[i] <= lastClosedMonthStart) closedForecastMonthOffsets.push(i);
    else openForecastMonthOffsets.push(i);
  }

  return {
    forecastStart,
    forecastEnd,
    forecastMonths,
    lastClosedMonthStart,
    forecastMonthIndexesInSales,
    historyMonthIndexes,
    closedForecastMonthOffsets,
    openForecastMonthOffsets
  };
}


/** SPOT背景（未知案件）を12ヶ月分推定：履歴同月平均を縮小しつつ最低保証を持たせる */
function estimateSpotBackground12Months_(spotSeries48, seriesStart, lastClosedMonthStart, baseP50ByMonth, tuning) {
  const out = new Array(12).fill(0);
  const src = Array.isArray(spotSeries48) ? spotSeries48 : [];
  const cfg = tuning || {};
  const shrink = isFinite(cfg.spotBgShrink) ? cfg.spotBgShrink : SPOT_BG_SHRINK;
  const floorRate = isFinite(cfg.spotBgFloorRate) ? cfg.spotBgFloorRate : SPOT_BG_FLOOR_RATE;
  const capRate = isFinite(cfg.spotBgCapRate) ? cfg.spotBgCapRate : SPOT_BG_CAP_RATE;
  const baseRef = Array.isArray(baseP50ByMonth) ? baseP50ByMonth : new Array(12).fill(0);
  if (src.length === 0) return out;

  const closedIdx = [];
  for (let i = 0; i < src.length; i++) {
    const mStart = addMonths_(seriesStart, i);
    if (mStart <= lastClosedMonthStart) closedIdx.push(i);
  }
  if (closedIdx.length === 0) return out;

  for (let m = 0; m < 12; m++) {
    const arr = [];
    for (let j = 0; j < closedIdx.length; j++) {
      const idx = closedIdx[j];
      if (idx % 12 !== m) continue;
      const v = Number(src[idx] || 0);
      if (isFinite(v) && v > 0) arr.push(v);
    }
    if (arr.length === 0) {
      out[m] = 0;
      continue;
    }
    const monthAvg = avg_(arr);
    const bg = Math.max(monthAvg * shrink, monthAvg * floorRate);
    const cap = Math.max(0, Number(baseRef[m] || 0) * capRate);
    out[m] = Math.max(0, Math.min(bg, cap));
  }

  return out;
}

function computeProductWeightsFromSalesInputClosed12_(fy, client, ctx) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.SALES_INPUT_MONTHLY);
  const map = new Map();
  if (!sh || sh.getLastRow() < 2) return map;

  const vals = sh.getDataRange().getValues().slice(1);
  const forecastStart = new Date(fy, 3, 1);
  const closedHistStart = addMonths_(forecastStart, -12);
  const closedHistEnd = ctx.lastClosedMonthStart;

  vals.forEach(r => {
    const c = String(r[0] || '').trim();
    const type = String(r[1] || '').trim();
    const product = String(r[2] || '').trim();
    const ym = toMonthStart_(r[3]);
    const amt = toNumberSafe_(r[4]);
    if (!c || !isSameClient_(c, client)) return;
    if (type !== 'BASE' || !product || !ym || !isFinite(amt)) return;
    if (ym < closedHistStart || ym > closedHistEnd) return;
    map.set(product, (map.get(product) || 0) + amt);
  });

  const total = Array.from(map.values()).reduce((a, b) => a + b, 0);
  if (total <= 0) return new Map();

  const out = new Map();
  map.forEach((v, k) => out.set(k, v / total));
  return out;
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
  const sh = ss.getSheetByName(SHEETS.DEV_SPOT);
  const out = new Array(12).fill(0);
  if (!sh) return out;

  const last = sh.getLastRow();
  if (last < 2) return out;

  const vals = sh.getRange(2, 1, last - 1, 5).getValues();
  const start = new Date(fy - 1, 3, 1);

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

/**
 * AI_RESEARCH_STRUCTURED から topic 別の adjusted_score を読み取り、
 * 予測モデル用の係数マップを返す。
 *
 * 返却値: { Market: number, Competitor: number, Channel: number, DX: number }
 * 各値は adjusted_score の topic 別平均（-50〜+50 の範囲）。
 * データがない場合は全て 0（ニュートラル）。
 */
function readAIResearchScores_() {
  const result = { Market: 0, Competitor: 0, Channel: 0, DX: 0 };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.AI_RESEARCH_STRUCTURED);
  if (!sh) return result;

  const last = sh.getLastRow();
  if (last < 2) return result;

  const vals = sh.getDataRange().getValues();
  const header = vals[0];

  // カラムindex特定（ヘッダ名で検索）
  const topicIdx = header.indexOf('topic');
  const adjIdx = header.indexOf('adjusted_score');
  if (topicIdx < 0 || adjIdx < 0) return result;

  const sums = { Market: 0, Competitor: 0, Channel: 0, DX: 0 };
  const counts = { Market: 0, Competitor: 0, Channel: 0, DX: 0 };

  for (let i = 1; i < vals.length; i++) {
    const topic = String(vals[i][topicIdx] || '').trim();
    const score = Number(vals[i][adjIdx] || 0);
    if (!isFinite(score)) continue;
    if (sums[topic] === undefined) continue;

    sums[topic] += score;
    counts[topic]++;
  }

  for (const k in result) {
    result[k] = counts[k] > 0 ? Math.round(sums[k] / counts[k] * 10) / 10 : 0;
  }

  return result;
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

  // AI調査スコア → 係数変換
  // adjusted_score の合計（全topic）を 0.01 倍して係数化
  // 例: 合計 +30 → 1.03（+3%の微調整）
  // 重みを意図的に小さくし、過度な影響を防ぐ
  const aiScores = opt.aiScores || { Market: 0, Competitor: 0, Channel: 0, DX: 0 };
  const aiTotalScore = (aiScores.Market || 0) + (aiScores.Competitor || 0) + (aiScores.Channel || 0) + (aiScores.DX || 0);
  const aiWeight = isFinite(opt.aiWeight) ? opt.aiWeight : AI_WEIGHT_DEFAULT;
  const aiMaxAbsEffect = isFinite(opt.aiMaxAbsEffect) ? opt.aiMaxAbsEffect : AI_MAX_ABS_EFFECT;
  const aiRawEffect = aiTotalScore * aiWeight;
  const aiClampedEffect = Math.max(-aiMaxAbsEffect, Math.min(aiMaxAbsEffect, aiRawEffect));
  const kAI = 1 + aiClampedEffect;

  const startT = 48;
  const simByMonth = Array.from({ length: 12 }, () => []);
  const opinionKByMonth = Array.from({ length: 12 }, () => []);
  const opsBaseByMonth = Array.from({ length: 12 }, (_, i) => {
    const t = startT + (i + 1);
    const mIdx = i % 12;
    return Math.max(0, (model.intercept + model.slope * t) * model.seasonalIndex[mIdx]);
  });

  for (let s = 0; s < nSim; s++) {
    for (let i = 0; i < 12; i++) {
      const t = startT + (i + 1);
      const mIdx = i % 12;

      const base = Math.max(0, (model.intercept + model.slope * t) * model.seasonalIndex[mIdx]);
      const e = residualPct[Math.floor(Math.random() * residualPct.length)] || 0;

      let ops = base * (1 + e);
      ops *= kProdByMonth[i];
      ops *= kClientByMonth[i];
      const kOpinion = sampleOpinionMultiplier_(opinions, months[i]);
      ops *= kOpinion;
      ops *= kAI;
      opinionKByMonth[i].push(kOpinion);

      const total = Math.max(0, ops) + devFixedByMonth[i];
      simByMonth[i].push(total);
    }
  }

  const p10 = simByMonth.map(arr => percentile_(arr, 0.10));
  const p50 = simByMonth.map(arr => percentile_(arr, 0.50));
  const p90 = simByMonth.map(arr => percentile_(arr, 0.90));

  const kOpinionP50ByMonth = opinionKByMonth.map(arr => percentile_(arr, 0.50));
  const kAIByMonth = new Array(12).fill(kAI);

  return {
    p10,
    p50,
    p90,
    diagnostics: {
      opsBaseByMonth,
      kProdByMonth,
      kClientByMonth,
      kOpinionP50ByMonth,
      kAIByMonth,
      aiTotalScore,
      AI_WEIGHT: aiWeight,
      aiRawEffect,
      aiClampedEffect,
      aiMaxAbsEffect
    }
  };
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
    throw new Error('初期セットアップが必要です。Forecast Agent > A-1 初期セットアップ を実行してください。');
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

  // Google Sheetsのシリアル日付に対応
  if (typeof v === 'number' && isFinite(v)) {
    const ms = Math.round((v - 25569) * 86400 * 1000);
    const dNum = new Date(ms);
    if (!isNaN(dNum.getTime())) return dNum;
  }

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


/** ====== v1.1 Phase1実装 ====== */
function buildPhase1Sheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  buildSimpleSheet_(ss, SHEETS.SALES_INPUT_MONTHLY, ['client','service_type','product','target_month','input_amount','status','source_updated_at']);
  buildSimpleSheet_(ss, SHEETS.ACTUAL_EVAL_MONTHLY, ['client','service_type','product','target_month','eval_actual_amount','actual_closed_flag','source_updated_at']);
  buildSimpleSheet_(ss, SHEETS.AI_RESEARCH_PROMPT, ['client','as_of_date','prompt_for_gem','paste_gem_output']);
  ss.getSheetByName(SHEETS.AI_RESEARCH_PROMPT).getRange('D:D').setNumberFormat('@');
  buildSimpleSheet_(ss, SHEETS.AI_RESEARCH_STRUCTURED, ['client','as_of_date','topic','direction','impact_score','confidence','evidence','time_horizon','business_relevance_reason','adjusted_score','report_text']);
  buildSimpleSheet_(ss, SHEETS.RUN_LOG, ['run_id','run_at','run_by','function_name','client','status','count','model_version','parameters_snapshot_json','input_data_hash','execution_duration_sec','error_summary']);
  buildSimpleSheet_(ss, SHEETS.FORECAST_SNAPSHOT, ['snapshot_id','run_date','client','target_month','scenario','linear_pred','robust_pred','regime_pred','simulation_pred','w1','w2','w3','w4','base_pred','subjective_adj','ai_adj','deterministic_adj','final_pred','confidence_interval_lower','confidence_interval_upper','key_factors_json','subjective_input_date']);
  buildSimpleSheet_(ss, SHEETS.EVAL_LOG, ['eval_id','evaluated_at','client','target_month','scenario','pred','actual','ape','was_overridden','error_category']);
  buildSimpleSheet_(ss, SHEETS.EVAL_COMPARE_MONTHLY, ['target_month','forecast_base','forecast_spot','forecast_total','actual_base','actual_spot','actual_total','gap_total']);
  buildSimpleSheet_(ss, SHEETS.EVAL_INSIGHTS, ['evaluated_at','client','target_month','actual_total','pred_p50','diff','error_rate','insight','next_action']);
  buildSimpleSheet_(ss, SHEETS.PROCESS_STATUS, ['step_key','last_run_date','last_run_by','status','target_client','record_count','error_summary']);
  buildSimpleSheet_(ss, SHEETS.FORECAST_REPORT, ['run_date','client','target_month','scenario','final_pred','base_pred','w1','w2','w3','w4','subjective_adj','ai_adj','deterministic_adj','factors_json']);
  buildSimpleSheet_(ss, SHEETS.DASHBOARD, ['metric','value','note']);
  initializeProcessStatus_();
}

function buildSimpleSheet_(ss, name, headers) {
  const sh = getOrCreateSheet_(ss, name);
  sh.clear();
  sh.getRange(1,1,1,headers.length).setValues([headers]).setBackground(COLOR_HEADER).setFontWeight('bold');
  sh.setFrozenRows(1);
}

function initializeProcessStatus_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.PROCESS_STATUS);
  const keys = ['step1_status','step2_status','step3_status','step3a_status','step4_status','step5_status','step6_status','step7_status'];
  const rows = keys.map(k => [k,'','', 'not_run','','','']);
  sh.getRange(2,1,rows.length,7).setValues(rows);
}

function importSalesInputMonthly() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cfg = ss.getSheetByName(SHEETS.CONFIG);
    const fy = Number(cfg.getRange('B3').getValue()) || getDefaultFY_();
    const result = importMonthlyFromExternal_(SHEETS.SALES_INPUT_MONTHLY, true);
    refreshManualInputSheets_(fy);
    const sh = ss.getSheetByName(SHEETS.SALES_INPUT_MONTHLY);
    if (sh) ss.setActiveSheet(sh);
    SpreadsheetApp.getUi().alert('完了', `売上データを取り込みました（${result.count}件 / ${result.range}）。
次は A-3 予測用に売上データを加工 を実行してください。`, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('エラー', e.message || e, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * A-3: SALES_INPUT_MONTHLY のデータを SALES シートに集計（BASE/SPOT × 48ヶ月横持ち）
 */
function aggregateSalesData() {
  try {
    ensureSetupDone_();
    requireStepSuccess_('step1_status', '先にA-2 売上データを取り込む を実行してください。');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cfg = ss.getSheetByName(SHEETS.CONFIG);
    const client = String(cfg.getRange('B2').getValue() || '').trim();
    const fy = Number(cfg.getRange('B3').getValue()) || getDefaultFY_();

    if (!client) throw new Error('CONFIG!B2 にクライアントを設定してください。');

    syncSalesFromSalesInput_(fy, client);

    // 集計結果を確認
    const sales = ss.getSheetByName(SHEETS.SALES);
    const salesData = sales.getDataRange().getValues();
    let nonZeroCount = 0;
    for (let r = 1; r < salesData.length; r++) {
      for (let c = 1; c < salesData[r].length; c++) {
        if (Number(salesData[r][c] || 0) !== 0) nonZeroCount++;
      }
    }

    ss.setActiveSheet(sales);

    if (nonZeroCount === 0) {
      SpreadsheetApp.getUi().alert(
        '警告',
        'SALESシートに集計しましたが、すべての値が0です。\n\n考えられる原因：\n・SALES_INPUT_MONTHLY の service_type（B列）が BASE/SPOT になっていない\n・SALES_INPUT_MONTHLY の target_month（D列）が予測FYの範囲外\n\nSALES_INPUT_MONTHLY の内容を確認してください。',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      SpreadsheetApp.getUi().alert(
        '完了',
        `SALESシートにBASE/SPOT × 48ヶ月の売上データを集計しました（非ゼロセル: ${nonZeroCount}）。
次は A-4〜A-9 を順番に実行してください。`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert('エラー', e.message || e, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function refreshManualInputSheets_(fy) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const people = getPeopleListFromConfig_();
  if (!people.length) return;

  const inSh = ss.getSheetByName(SHEETS.SALES_INPUT_MONTHLY);
  const vals = inSh.getDataRange().getValues().slice(1);
  const products = Array.from(new Set(vals.map(r => String(r[2] || '').trim()).filter(Boolean))).sort();
  if (!products.length) return;

  const defaultDate = new Date((Number(fy) || getDefaultFY_()), 3, 1);
  ensureFactorsProductTemplate_(ss.getSheetByName(SHEETS.FACTORS_PRODUCT), products, people, defaultDate);
  ensureFactorsClientTemplate_(ss.getSheetByName(SHEETS.FACTORS_CLIENT), people, defaultDate);
  ensureOpinionsTemplate_(ss.getSheetByName(SHEETS.OPINIONS), people, defaultDate);
  ensureDevTemplate_(ss.getSheetByName(SHEETS.DEV_SPOT), people, defaultDate);
}

function importActualEvalMonthly() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    importMonthlyFromExternal_(SHEETS.ACTUAL_EVAL_MONTHLY, false);
    const sh = ss.getSheetByName(SHEETS.ACTUAL_EVAL_MONTHLY);
    if (sh) ss.setActiveSheet(sh);
    SpreadsheetApp.getUi().alert('完了', '検証実績を更新しました。次は B-2 予測検証レポート更新 を実行できます。', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('エラー', e.message || e, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function writeRowsInChunks_(sh, startRow, startCol, rows, chunkSize) {
  if (!rows || !rows.length) return;
  const size = Math.max(1, Number(chunkSize) || 2000);
  for (let i = 0; i < rows.length; i += size) {
    const chunk = rows.slice(i, i + size);
    sh.getRange(startRow + i, startCol, chunk.length, chunk[0].length).setValues(chunk);
  }
}

function classifyServiceType_(serviceCategoryRaw) {
  const serviceCategory = String(serviceCategoryRaw || '').trim().toLowerCase();

  // 優先ルール：ベース/スポットの明示文字列を最優先
  if (serviceCategory.includes('ベース')) return 'BASE';
  if (serviceCategory.includes('スポット')) return 'SPOT';

  // 追加マッピング（部分一致）
  const baseKeywords = ['フラグメント', 'テンプレート', '運用更新', '簡便化', '保守サポート'];
  const spotKeywords = ['開発', 'その他', 'myinsights'];

  if (baseKeywords.some(k => serviceCategory.includes(k.toLowerCase()))) return 'BASE';
  if (spotKeywords.some(k => serviceCategory.includes(k.toLowerCase()))) return 'SPOT';

  return 'OTHER';
}

function importMonthlyFromExternal_(targetSheetName, withStatus) {
  const started = new Date();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(targetSheetName);
  const cfg = ss.getSheetByName(SHEETS.CONFIG);
  const targetClient = normalizeClientName_(String(cfg.getRange('B2').getValue() || '').trim());
  const fy = Number(cfg.getRange('B3').getValue()) || getDefaultFY_();
  if (!targetClient) throw new Error('CONFIG!B2 にクライアントを設定してください。');

  const ext = SpreadsheetApp.openById(EXTERNAL_SS_ID);
  const sheets = ext.getSheets().filter(s => s.getName().startsWith(EXTERNAL_SHEET_PREFIX) && s.getName().endsWith(EXTERNAL_SHEET_SUFFIX));

  const isSalesInput = targetSheetName === SHEETS.SALES_INPUT_MONTHLY;
  const start = isSalesInput ? new Date(fy - 4, 3, 1) : new Date(fy - 3, 3, 1);
  const end = isSalesInput ? new Date(fy, 2, 1) : new Date(fy + 1, 2, 1);
  const rows = [];
  const now = new Date();
  const currMonth = new Date(now.getFullYear(), now.getMonth(), 1);

  sheets.forEach(sht => {
    toastProgress_(ss, `取り込み中: ${sht.getName()}（${rows.length}行取得済み）…`, 3);
    const lastRow = sht.getLastRow();
    if (lastRow < 2) return;
    const readCols = Math.max(EXT_COL_AMOUNT, EXT_COL_DATE_PRIMARY, EXT_COL_DATE_SECONDARY, EXT_COL_SERVICE_CATEGORY, EXT_COL_CATEGORY, EXT_COL_CLIENT);
    const vals = sht.getRange(2, 1, lastRow - 1, readCols).getValues();
    for (let i = 0; i < vals.length; i++) {
      const r = vals[i];
      const client = String(r[EXT_COL_CLIENT - 1] || '').trim();
      if (!isSameClient_(client, targetClient)) continue;

      const serviceCategory = String(r[EXT_COL_SERVICE_CATEGORY - 1] || '').trim();
      const serviceType = classifyServiceType_(serviceCategory);
      if (serviceType === 'OTHER') continue;

      let d = r[EXT_COL_DATE_PRIMARY - 1];
      let dt = toDate_(d);
      if (!dt) {
        d = r[EXT_COL_DATE_SECONDARY - 1];
        dt = toDate_(d);
      }
      if (!dt) continue;
      const ym = new Date(dt.getFullYear(), dt.getMonth(), 1);
      if (ym < start || ym > end) continue;

      const product = String(r[EXT_COL_CATEGORY - 1] || '').trim() || serviceType;
      const amount = Number(r[EXT_COL_AMOUNT - 1] || 0);
      if (!isFinite(amount)) continue;

      if (withStatus) {
        const status = ym >= currMonth ? 'open' : 'closed';
        rows.push([normalizeClientName_(client), serviceType, product, fmtYM_(ym), amount, status, new Date()]);
      } else {
        const closed = ym < currMonth ? 1 : 0;
        rows.push([normalizeClientName_(client), serviceType, product, fmtYM_(ym), amount, closed, new Date()]);
      }
    }
  });

  rows.sort((a, b) => (a[0] + a[1] + a[3] + a[2]).localeCompare(b[0] + b[1] + b[3] + b[2]));

  sh.getRange(2, 1, Math.max(1, sh.getMaxRows() - 1), sh.getLastColumn()).clearContent();
  if (rows.length) writeRowsInChunks_(sh, 2, 1, rows, 2000);

  // D列（target_month）をテキスト形式に設定（Sheets自動Date変換を防止）
  if (rows.length) {
    sh.getRange(2, 4, rows.length, 1).setNumberFormat('@');
    sh.getRange(2, 5, rows.length, 1).setNumberFormat('#,##0');
  }

  // 取得データの月範囲をログ
  const ymSet = new Set(rows.map(r => String(r[3] || '')));
  const ymSorted = Array.from(ymSet).sort();
  const expectedMonths = 48;
  const rangeInfo = ymSorted.length
    ? `${ymSorted[0]}〜${ymSorted[ymSorted.length - 1]}（${ymSorted.length}ヶ月 / 想定${expectedMonths}ヶ月）`
    : `データなし（想定${expectedMonths}ヶ月）`;

  const step = (targetSheetName === SHEETS.SALES_INPUT_MONTHLY) ? 'step1_status' : 'step2_status';
  updateProcessStatus_(step, 'success', targetClient, rows.length, '');
  logRun_((targetSheetName === SHEETS.SALES_INPUT_MONTHLY) ? 'importSalesInputMonthly' : 'importActualEvalMonthly', targetClient, 'success', rows.length, started, rangeInfo);

  return { count: rows.length, range: rangeInfo };
}

function generateAIResearchTemplate() {
  requireStepSuccess_('step1_status', '先にA-2 売上データを取り込む を実行してください。');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName(SHEETS.CONFIG);
  const targetClient = String(cfg.getRange('B2').getValue() || '').trim();
  if (!targetClient) throw new Error('CONFIG!B2 にクライアントを設定してください。');
  const shIn = ss.getSheetByName(SHEETS.SALES_INPUT_MONTHLY);
  const shOut = ss.getSheetByName(SHEETS.AI_RESEARCH_PROMPT);
  const vals = shIn.getDataRange().getValues().slice(1);
  const clients = Array.from(new Set(
    vals
      .map(r => String(r[0] || '').trim())
      .filter(c => c && isSameClient_(c, targetClient))
      .map(c => normalizeClientName_(c))
  ));

  const rows=[];
  clients.sort().forEach(c=>{
    const prompt = [
      `Client_Name: ${normalizeClientName_(c)}`,
      `As_of_Date: ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')}`,
      '',
      'カスタム指示に従い、出力してください。'
    ].join('\n');
    rows.push([normalizeClientName_(c),new Date(),prompt]);
  });
  shOut.getRange(2,1,Math.max(1,shOut.getMaxRows()-1),4).clearContent();
  if(rows.length) shOut.getRange(2,1,rows.length,3).setValues(rows);
  shOut.getRange('D1').setValue('paste_gem_output').setBackground('#ffe599').setFontWeight('bold');
  shOut.getRange('D:D').setNumberFormat('@');
  shOut.getRange('D2').setBackground('#fff2cc').setNote('ここにGemの出力を【全文そのまま】貼り付けてください。###REPORT_START### / ###TSV_START### の両方を含んだ状態で貼り付けてOKです（先頭に=は不要）。A-8実行時に自動でパースされます。');
  shOut.setColumnWidth(4, 420);
  updateProcessStatus_('step3_status','success',targetClient,rows.length,'');
  logRun_('generateAIResearchTemplate',targetClient, 'success', rows.length, new Date(), '');
  ss.setActiveSheet(shOut);
  showPromptPreviewDialog_(rows);
}

function parseAIResearchPaste_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.AI_RESEARCH_PROMPT);
  const out = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.AI_RESEARCH_STRUCTURED);
  const raw = String(sh.getRange('D2').getValue() || '').trim();
  if (!raw) return 0;

  // レポート部分を抽出
  const reportMatch = raw.match(/(?:###|===)REPORT_START(?:###|===)([\s\S]*?)(?:###|===)REPORT_END(?:###|===)/);
  const report = reportMatch ? reportMatch[1].trim() : '';

  // TSV部分を抽出
  const tsvMatch = raw.match(/(?:###|===)TSV_START(?:###|===)([\s\S]*?)(?:###|===)TSV_END(?:###|===)/);
  if (!tsvMatch) return 0;

  const tsvLines = tsvMatch[1].trim().split(/\r?\n/).filter(l => l.trim());

  const rows = [];
  tsvLines.forEach((ln, idx) => {
    const cols = ln.split('\t');
    // ヘッダ行スキップ（clientで始まる行）
    if (idx === 0 && cols[0] && cols[0].trim().toLowerCase() === 'client') return;
    if (cols.length < 9) return;

    const impactScore = Number(cols[4] || 0);
    const confidence = Number(cols[5] || 0);
    if (!isFinite(impactScore) || impactScore < 0 || impactScore > 100) return;
    if (!isFinite(confidence) || confidence < 0 || confidence > 1) return;

    const adjustedScore = Math.round((impactScore - 50) * confidence * 10) / 10;

    // 9列のTSVデータ + adjusted_score + report_text
    const row = cols.slice(0, 9);
    row.push(adjustedScore);
    row.push(rows.length === 0 ? report : '');  // レポートは最初の行にだけ格納
    rows.push(row);
  });

  out.getRange(2, 1, Math.max(1, out.getMaxRows() - 1), 11).clearContent();
  if (rows.length) out.getRange(2, 1, rows.length, 11).setValues(rows);

  const cfg = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CONFIG);
  const client = String(cfg.getRange('B2').getValue() || '').trim();
  updateProcessStatus_('step3a_status', 'success', client, rows.length, '');
  return rows.length;
}
function runPhase1Forecast() {
  try {
    requireStepSuccess_('step1_status', '先にA-2 売上データを取り込む を実行してください。');
    const started = new Date();
    const parsed = parseAIResearchPaste_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cfg = ss.getSheetByName(SHEETS.CONFIG);
    const client = String(cfg.getRange('B2').getValue() || '').trim();
    if (!client) throw new Error('CONFIG!B2 にクライアントを設定してください。');
    const fy = Number(cfg.getRange('B3').getValue()) || getDefaultFY_();
    cfg.getRange('B3').setValue(fy);
    validateAllInputsOrThrow_(fy);
    validateRequiredUserInputsOrThrow_();
    runHierarchicalA9AlertsOrThrow_(fy);
    syncSalesFromSalesInput_(fy, client);
    const result = runForecastFYCore_(fy, client);
    writeOutputFY_(result);
    writeForecastArtifacts_(result, client);
    ss.setActiveSheet(ss.getSheetByName(SHEETS.OUTPUT));
    updateProcessStatus_('step4_status','success',client,result.months.length,'');
    logRun_('runPhase1Forecast', client, 'success', result.months.length, started, `ai_rows=${parsed}`);
    SpreadsheetApp.getUi().alert('完了', '予測を更新しました。次は A-10 予測ダッシュボードを更新 を実行してください。', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    updateProcessStatus_('step4_status','error','',0,String(e.message || e));
    SpreadsheetApp.getUi().alert('予測実行エラー', e.message || e, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function writeForecastArtifacts_(result, client) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const snap = ss.getSheetByName(SHEETS.FORECAST_SNAPSHOT);
  const rep = ss.getSheetByName(SHEETS.FORECAST_REPORT);
  const runDate = new Date();
  const sid = Utilities.getUuid();
  const rows=[];
  const scenarios = [
    {name:'nega', arr:result.mixed.p10},
    {name:'neutral', arr:result.mixed.p50},
    {name:'posi', arr:result.mixed.p90}
  ];
  scenarios.forEach(sc=>{
    result.months.forEach((m,i)=>{
      const deterministicAdj = (result.spotFixedByMonth && isFinite(result.spotFixedByMonth[i])) ? result.spotFixedByMonth[i] : (result.devFixedByMonth[i] || 0);
      rows.push([sid,runDate,client,fmtYM_(m),sc.name,'','','','',0.15,0.40,0.25,0.20,result.mixed.p50[i],0,0,deterministicAdj,sc.arr[i],result.mixed.p10[i],result.mixed.p90[i],JSON.stringify({opinion:result.opinionsSummaryByMonth[i]||''}),null]);
    });
  });
  const r0 = snap.getLastRow()+1;
  snap.getRange(r0,1,rows.length,rows[0].length).setValues(rows);

  const repRows = rows.map(r=>[runDate,client,r[3],r[4],r[17],r[13],r[9],r[10],r[11],r[12],r[14],r[15],r[16],r[20]]);
  rep.getRange(rep.getLastRow()+1,1,repRows.length,repRows[0].length).setValues(repRows);
}

/**
 * 実績確定後の検証ステップ。
 * - B-1の後に実行することで EVAL_LOG が更新される
 * - Phase移行判断に使うKPI（sMAPE等）の元データを蓄積する
 */
function updatePhase1EvaluationReport() {
  requireStepSuccess_('step2_status', '先にB-1 検証用に実績データを取り込み を実行してください。');
  requireStepSuccess_('step4_status', '先にA-9 予測実行を実行してください。');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const actual = ss.getSheetByName(SHEETS.ACTUAL_EVAL_MONTHLY).getDataRange().getValues().slice(1);
  const snap = ss.getSheetByName(SHEETS.FORECAST_SNAPSHOT).getDataRange().getValues().slice(1);
  const mapA = new Map();
  actual.forEach(r=>{
    const k = [r[0], r[3]].join('|');
    mapA.set(k, (mapA.get(k) || 0) + Number(r[4] || 0));
  });
  const evalRows=[];
  snap.forEach(r=>{
    const key = [r[2], r[3]].join('|');
    const act = mapA.get(key);
    if (act == null) return;
    const pred = Number(r[17]||0);
    const ape = act ? Math.abs(pred-act)/Math.abs(act) : '';
    evalRows.push([Utilities.getUuid(),new Date(),r[2],r[3],r[4],pred,act,ape,0,'model_limitation']);
  });
  const out = ss.getSheetByName(SHEETS.EVAL_LOG);
  if (evalRows.length) out.getRange(out.getLastRow()+1,1,evalRows.length,evalRows[0].length).setValues(evalRows);

  const compare = ss.getSheetByName(SHEETS.EVAL_COMPARE_MONTHLY);
  writeEvalCompareMonthly_(compare, actual, snap);

  updateProcessStatus_('step5_status','success','',evalRows.length,'');
  logRun_('updatePhase1EvaluationReport','', 'success', evalRows.length, new Date(), '');
  ss.setActiveSheet(compare || out);
}

function writeEvalCompareMonthly_(sh, actualRows, snapRows) {
  if (!sh) return;

  const actualMap = new Map();
  actualRows.forEach(r => {
    const ym = String(r[3] || '');
    const type = String(r[1] || '').trim().toUpperCase();
    const amt = Number(r[4] || 0);
    if (!ym || !isFinite(amt)) return;
    if (!actualMap.has(ym)) actualMap.set(ym, { BASE: 0, SPOT: 0 });
    if (type === 'BASE' || type === 'SPOT') actualMap.get(ym)[type] += amt;
  });

  const neutralMap = new Map();
  snapRows.forEach(r => {
    if (String(r[4] || '') !== 'neutral') return;
    const ym = String(r[3] || '');
    const pred = Number(r[17] || 0);
    if (!ym || !isFinite(pred)) return;
    neutralMap.set(ym, pred);
  });

  const ratio = getBaseSpotRatioFromSales_();
  const months = Array.from(new Set([...actualMap.keys(), ...neutralMap.keys()])).sort();
  const rows = months.map(ym => {
    const predTotal = Number(neutralMap.get(ym) || 0);
    const predBase = predTotal * ratio.base;
    const predSpot = predTotal * ratio.spot;
    const act = actualMap.get(ym) || { BASE: 0, SPOT: 0 };
    const actTotal = act.BASE + act.SPOT;
    return [ym, predBase, predSpot, predTotal, act.BASE, act.SPOT, actTotal, actTotal - predTotal];
  });

  sh.getRange(2, 1, Math.max(1, sh.getMaxRows() - 1), 8).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, 8).setValues(rows);
}

function getBaseSpotRatioFromSales_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SALES);
  if (!sh || sh.getLastRow() < 2) return { base: 0.5, spot: 0.5 };
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, Math.min(sh.getLastColumn(), 50)).getValues();
  let base = 0;
  let spot = 0;
  vals.forEach(r => {
    const t = String(r[0] || '').trim().toUpperCase();
    let s = 0;
    for (let i = 2; i < r.length; i++) s += Number(r[i] || 0);
    if (t === 'BASE') base += s;
    if (t === 'SPOT') spot += s;
  });
  const total = base + spot;
  if (total <= 0) return { base: 0.5, spot: 0.5 };
  return { base: base / total, spot: spot / total };
}

/**
 * 現場閲覧用サマリー更新。
 * - OUTPUTの理解補助（件数・更新時刻・KPI信号）を表示
 * - 詳細分析は FORECAST_REPORT / EVAL_LOG を参照
 */
function updatePhase1Dashboard() {
  requireStepSuccess_('step4_status', '先にA-9 予測実行を実行してください。');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName(SHEETS.DASHBOARD);
  const rep = ss.getSheetByName(SHEETS.FORECAST_REPORT).getDataRange().getValues();
  dash.clear();
  buildSimpleSheet_(ss, SHEETS.DASHBOARD, ['metric','value','note']);
  const total = rep.length - 1;
  dash.getRange(2,1,4,3).setValues([
    ['forecast_rows', total, 'FORECAST_REPORT件数'],
    ['last_updated', new Date(), '更新日時'],
    ['kpi_smape_signal', 'N/A', 'A-9実行後に算出'],
    ['dashboard_status', 'ready', '初期ダッシュボード']
  ]);
  updateProcessStatus_('step6_status','success','',total,'');
  logRun_('updatePhase1Dashboard','', 'success', total, new Date(), '');
  ss.setActiveSheet(dash);
}

function updatePhase1LearningInsights() {
  requireStepSuccess_('step5_status', '先にB-2 検証レポートを更新してください。');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName(SHEETS.CONFIG);
  const client = String(cfg.getRange('B2').getValue() || '').trim();
  const evalSh = ss.getSheetByName(SHEETS.EVAL_LOG);
  const out = ss.getSheetByName(SHEETS.EVAL_INSIGHTS);
  const vals = evalSh.getDataRange().getValues().slice(1).filter(r => String(r[2] || '').trim() === client);

  const byMonth = new Map();
  vals.forEach(r => {
    const month = String(r[3] || '');
    const scenario = String(r[4] || '');
    const pred = Number(r[5] || 0);
    const actual = Number(r[6] || 0);
    if (!month) return;
    if (!byMonth.has(month)) byMonth.set(month, {actual: 0, p50: 0, hasP50: false});
    const obj = byMonth.get(month);
    obj.actual = Math.max(obj.actual, actual);
    if (scenario === 'neutral') {
      obj.p50 = pred;
      obj.hasP50 = true;
    }
  });

  const rows = [];
  Array.from(byMonth.keys()).sort().forEach(month => {
    const v = byMonth.get(month);
    if (!v.hasP50) return;
    const diff = v.actual - v.p50;
    const rate = (v.actual !== 0) ? (diff / Math.abs(v.actual)) : 0;
    const insight = (Math.abs(rate) < 0.1)
      ? '予測精度は概ね良好。継続運用。'
      : (rate > 0
        ? '実績が予測超過。増加要因（スポット案件・大型失注回避等）を追加学習。'
        : '実績が予測未達。失注・延期・単価低下要因を確認。');
    const nextAction = (Math.abs(rate) < 0.1)
      ? '現行手順を継続し、次月も同手順で検証。'
      : 'B-3で要因を記録し、A-3〜A-7入力項目へ反映。';
    rows.push([new Date(), client, month, v.actual, v.p50, diff, rate, insight, nextAction]);
  });

  out.getRange(2,1,Math.max(1,out.getMaxRows()-1),9).clearContent();
  if (rows.length) out.getRange(2,1,rows.length,9).setValues(rows);
  updateProcessStatus_('step7_status', 'success', client, rows.length, '');
  logRun_('updatePhase1LearningInsights', client, 'success', rows.length, new Date(), '');
  ss.setActiveSheet(out);
}

function chooseClientFromSalesInput_() {
  const cfg = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CONFIG);
  return String(cfg.getRange('B2').getValue() || '').trim();
}

function updateProcessStatus_(stepKey, status, targetClient, count, err) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.PROCESS_STATUS);
  const vals = sh.getDataRange().getValues();
  for (let i=1;i<vals.length;i++) {
    if (vals[i][0] === stepKey) {
      sh.getRange(i+1,2,1,6).setValues([[new Date(), Session.getActiveUser().getEmail()||'unknown', status, targetClient||'', count||0, err||'']]);
      return;
    }
  }
  sh.appendRow([stepKey,new Date(),Session.getActiveUser().getEmail()||'unknown',status,targetClient||'',count||0,err||'']);
}

function requireStepSuccess_(stepKey, message) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.PROCESS_STATUS);
  const vals = sh.getDataRange().getValues();
  const row = vals.find(r=>r[0]===stepKey);
  if (!row || row[3] !== 'success') throw new Error(message);
}

function logRun_(fn, client, status, count, startedAt, err) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.RUN_LOG);
  const end = new Date();
  const sec = Math.round((end - startedAt) / 1000);
  const params = JSON.stringify({N_SIM, SPIKE_CLIP_MIN, SPIKE_CLIP_MAX, TREND_FACTOR_MIN, TREND_FACTOR_MAX});
  const hash = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, `${fn}|${client}|${end.toISOString()}`));
  sh.appendRow([Utilities.getUuid(), end, Session.getActiveUser().getEmail()||'unknown', fn, client||'', status, count||0, VERSION, params, hash, sec, err||'']);
}

function parseYM_(s) {
  const m = String(s||'').match(/^(\d{4})\/(\d{2})$/);
  if (!m) return null;
  return new Date(Number(m[1]), Number(m[2])-1, 1);
}

/**
 * Date型・文字列型の両方から「その月の1日」のDateオブジェクトを返す。
 * - Date型: そのまま月初に変換
 * - 文字列 "YYYY/MM": パースして月初に変換
 * - それ以外: null
 */
function toMonthStart_(v) {
  if (!v) return null;

  // Date型の場合（Sheetsが自動変換した場合）
  if (v instanceof Date && !isNaN(v.getTime())) {
    return new Date(v.getFullYear(), v.getMonth(), 1);
  }

  // 文字列の場合
  const s = String(v).trim();
  if (!s) return null;

  // "YYYY/MM" 形式
  const m1 = s.match(/^(\d{4})\/(\d{1,2})$/);
  if (m1) return new Date(Number(m1[1]), Number(m1[2]) - 1, 1);

  // "YYYY-MM" 形式
  const m2 = s.match(/^(\d{4})-(\d{1,2})$/);
  if (m2) return new Date(Number(m2[1]), Number(m2[2]) - 1, 1);

  // "YYYY/MM/DD" 形式（日を無視して月初に）
  const m3 = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-]\d{1,2}/);
  if (m3) return new Date(Number(m3[1]), Number(m3[2]) - 1, 1);

  return null;
}


function applyTabColors_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const colorAuto = '#0b5394';
  const colorManual = '#bf9000';
  const colorOutput = '#990000';
  const colorEval = '#38761d';
  const colorGuide = '#666666';

  const manual = [SHEETS.FACTORS_PRODUCT, SHEETS.FACTORS_CLIENT, SHEETS.OPINIONS, SHEETS.DEV_SPOT];
  const auto = [SHEETS.SALES_INPUT_MONTHLY, SHEETS.SALES, SHEETS.AI_RESEARCH_PROMPT];
  const output = [SHEETS.OUTPUT, SHEETS.FORECAST_REPORT, SHEETS.DASHBOARD];
  const evalSheets = [SHEETS.ACTUAL_EVAL_MONTHLY, SHEETS.EVAL_COMPARE_MONTHLY, SHEETS.EVAL_LOG, SHEETS.EVAL_INSIGHTS];
  const guide = [SHEETS.GUIDE, SHEETS.CONFIG];

  manual.forEach(n => { const sh = ss.getSheetByName(n); if (sh) sh.setTabColor(colorManual); });
  auto.forEach(n => { const sh = ss.getSheetByName(n); if (sh) sh.setTabColor(colorAuto); });
  output.forEach(n => { const sh = ss.getSheetByName(n); if (sh) sh.setTabColor(colorOutput); });
  evalSheets.forEach(n => { const sh = ss.getSheetByName(n); if (sh) sh.setTabColor(colorEval); });
  guide.forEach(n => { const sh = ss.getSheetByName(n); if (sh) sh.setTabColor(colorGuide); });
}

function hideNonUserSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hideTargets = [SHEETS.AI_RESEARCH_STRUCTURED, SHEETS.RUN_LOG, SHEETS.FORECAST_SNAPSHOT, SHEETS.PROCESS_STATUS];
  hideTargets.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) sh.hideSheet();
  });
}

function setGuideLinkTable_(guideSheet, startRow, links) {
  const colorByLabel = {
    '自動入力用': '#d9e8fb',
    'ユーザ入力用': '#fff2cc',
    '出力用': '#f4cccc',
    '事後検証用': '#d9ead3'
  };
  links.forEach((item, i) => {
    const row = startRow + i;
    const target = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(item[1]);
    guideSheet.getRange(row, 1).setValue(item[0]);
    if (target) {
      const formula = `=HYPERLINK("#gid=${target.getSheetId()}", "${item[2]}")`;
      guideSheet.getRange(row, 2).setFormula(formula);
    } else {
      guideSheet.getRange(row, 2).setValue(item[2]);
    }
    guideSheet.getRange(row, 3).setValue(item[1]);
    if (colorByLabel[item[0]]) guideSheet.getRange(row, 1, 1, 3).setBackground(colorByLabel[item[0]]);
  });
}

function showPromptPreviewDialog_(rows) {
  if (!rows || !rows.length) return;
  const prompt = String(rows[0][2] || '');
  const pasteTarget = `${SHEETS.AI_RESEARCH_PROMPT}!D2`; 
  const html = `
  <div style="font-family:sans-serif;padding:12px">
    <h3>AIプロンプト（コピーして利用）</h3>
    <div style="font-size:12px;color:#444;margin-bottom:8px;line-height:1.6;">
      0) <b style="color:#ea4335">Gemの右下のモードを「Pro」に切り替えてください（高速モード等は不可）</b><br>
      1) 下のプロンプトをコピーしてGemに貼り付けて実行してください。<br>
      2) Gemにアクセス（<a href="https://gemini.google.com/gem/1NGUI4UI_tuNF3NvwXV323iuQsqEALB0p?usp=sharing" target="_blank">こちら</a>）し、結果を <b>全文コピー</b> して <b>paste_gem_output（黄色になっている箇所）</b> にペーストしてください。<br>
    </div>
    <textarea id="p" style="width:100%;height:120px">${escapeHtml_(prompt)}</textarea>
    <div style="margin-top:10px">
      <button onclick="document.getElementById('p').select();document.execCommand('copy');">コピー</button>
      <button onclick="google.script.host.close();">閉じる</button>
    </div>
  </div>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(640).setHeight(420), 'AIプロンプト');
}

function syncSalesFromSalesInput_(fy, client) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inSh = ss.getSheetByName(SHEETS.SALES_INPUT_MONTHLY);
  const sales = ss.getSheetByName(SHEETS.SALES);
  if (!inSh || !sales) throw new Error('SALES_INPUT_MONTHLY または SALES がありません。');

  const start = new Date(fy - 4, 3, 1);
  const totalMonths = 48;
  const vals = inSh.getDataRange().getValues().slice(1);
  const map = new Map();
  vals.forEach(r => {
    const c = String(r[0] || '').trim();
    const p = String(r[1] || '').trim();
    const ym = toMonthStart_(r[3]);
    const amt = Number(r[4] || 0);
    if (!c || !p || !ym || !isFinite(amt) || !isSameClient_(c, client)) return;
    if (p !== 'BASE' && p !== 'SPOT') return;
    const idx = monthIndexFromStart_(ym, start);
    if (idx < 0 || idx >= totalMonths) return;
    if (!map.has(p)) map.set(p, new Array(totalMonths).fill(0));
    map.get(p)[idx] += amt;
  });

  buildSALES_();
  const headerMonths = [];
  for (let i = 0; i < totalMonths; i++) headerMonths.push(fmtYM_(addMonths_(start, i)));
  sales.getRange(1, 1).setValue('Category');
  sales.getRange(1, 2, 1, totalMonths).setValues([headerMonths]);

  const names = ['BASE', 'SPOT'];
  const out = names.map(n => [n, ...(map.get(n) || new Array(totalMonths).fill(0))]);
  const totalRow = ['TOTAL', ...new Array(totalMonths).fill(0).map((_, i) => Number((map.get('BASE') || [])[i] || 0) + Number((map.get('SPOT') || [])[i] || 0))];

  sales.getRange(2,1,Math.max(1,sales.getMaxRows()-1),1+totalMonths).clearContent();
  const allRows = [...out, totalRow];
  sales.getRange(2,1,allRows.length,1+totalMonths).setValues(allRows);
  sales.getRange(2,2,allRows.length,totalMonths).setNumberFormat('#,##0');
  sales.getRange(2,2,2,totalMonths).setBackground(COLOR_OBJECTIVE);
  sales.getRange(4,1,1,1+totalMonths).setBackground('#eeeeee').setFontWeight('bold');
}
