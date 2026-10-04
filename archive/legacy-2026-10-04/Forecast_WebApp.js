/**
 * Forecast_WebApp.js — クライアント別売上予測 Webアプリ層
 *
 * このスプレッドシート（クライアント別売上予測）にバインドされたスクリプトの
 * Webアプリ版フロント。既存メニュー（A-1〜C-3）と同じデータ・関数を共有し、
 * スプレッドシートUIを開かずにブラウザから運用できるようにする。
 *
 * 画面構成（Forecast_WebApp.html）:
 *   #home      進捗ダッシュボード（プロセスステップ・次の一手・実行履歴）
 *   #prep      準備（A-1 設定 / A-2 取り込み / A-3 加工 / A-4 AI調査）
 *   #input     入力（A-5 製品 / A-6 クライアント / A-7 見解 / A-8 Spot）
 *   #forecast  予測と予算（A-9 実行 + A-10 予算編集 + OUTPUT 可視化）
 *   #eval      検証（B-1 実績取込 / B-2 レポート / B-3 ダッシュボード / B-4 インサイト）
 *   #quarterly 四半期レビュー（C-1 生成 / 承認入力 / C-3 適用）
 *
 * 公開関数 web* は google.script.run から呼ばれる。内部関数は *_ 接尾辞。
 * getUi() が使えないコンテキストでは Forecast_Agent.js の alertOrThrow_ /
 * forecastConfirmOrAbort_ が例外に変換される。
 */

// ===== doGet =====

function doGet(e) {
  const t = HtmlService.createTemplateFromFile('Forecast_WebAppUI');
  t.bootJson = webJson_(webGetBootstrap_());
  const out = t.evaluate()
    .setTitle('クライアント別売上予測')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return webSetFavicon_(out);
}

/**
 * タブのアイコンを案内キャラ「よみ」にする（下の FORECAST_FAVICON_URL）。
 * ページ内の <link rel="icon"> は Apps Script に無視されるため HtmlOutput.setFaviconUrl で渡す。
 * 失敗しても画面は止めず、既定のアイコンのまま返す。
 */
function webSetFavicon_(out) {
  try {
    out.setFaviconUrl(FORECAST_FAVICON_URL);
  } catch (err) {
    Logger.log('webSetFavicon_: ' + (err && err.message ? err.message : err));
  }
  return out;
}

// ===== タブのアイコン（自動生成: assets/characters/src/build.py。この区間は手で編集しない） =====
// 案内キャラ「よみ」の頭を枠いっぱいに描いた 64x64 の PNG（assets/characters/favicon/yomi_favicon_64.png）。
// 末尾の #favicon.png は消さない。Apps Script は末尾が画像の拡張子でない URL を例外なしで捨てる（# 以降は画像データではない）。
// 別ファイルにしないのは、管理ハブ runtime の許可リスト（VNEXT_ADMIN_RUNTIME_FILE_TYPES_ の 21 ファイル）を変えないため。
const FORECAST_FAVICON_URL = 'data:image/png;base64,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#favicon.png';
// ===== /タブのアイコン =====

function webJson_(obj) {
  return JSON.stringify(obj).replace(/</g, '\\u003c');
}

// ===== bootstrap / parsers =====

function webGetBootstrap() {
  return webGetBootstrap_();
}

function webGetBootstrap_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName(SHEETS.CONFIG);
  const client = cfg ? String(cfg.getRange('B2').getValue() || '').trim() : '';
  const fy = Number(cfg ? cfg.getRange('B3').getValue() : 0) || getDefaultFY_();
  const peopleCsv = cfg ? String(cfg.getRange('B4').getValue() || '').trim() : '';

  const boot = {
    user: Session.getActiveUser().getEmail() || '',
    bookUrl: ss.getUrl(),
    version: VERSION + ' / ' + BUILD_STAGE,
    setup: {
      client: client,
      fy: fy,
      defaultFy: getDefaultFY_(),
      peopleCsv: peopleCsv,
      people: peopleCsv ? peopleCsv.split(',').map(s => s.trim()).filter(Boolean) : [],
      aiEnabled: false,
      geminiReady: false
    },
    steps: webParseSteps_(ss),
    months: webFyMonths_(fy),
    stepOptions: buildPercentStepList_(),
    products: webProductsList_(ss),
    input: webParseInputs_(ss),
    output: webParseOutput_(ss),
    eval: webParseEval_(ss),
    quarterly: webParseQuarterly_(ss),
    runLog: webParseRunLog_(ss),
    learning: webParseLearning_(ss, client)
  };
  // 操作の記録（監査ログ）へのリンクは管理者にだけ渡す
  const actor = webActor_();
  boot.access = { isAdmin: actor.isAdmin, auditUrl: actor.isAdmin ? webAuditLogUrl_() : '' };

  try {
    const vertex = readVertexConfig_();
    boot.setup.aiEnabled = !!vertex.enabled;
    boot.setup.geminiReady = !!vertex.geminiReady;
    boot.setup.ragReady = !!vertex.ragReady;
  } catch (e) { /* 未作成の初期状態では空で返す */ }

  return boot;
}

/** FY の 12ヶ月ラベル（FY=2026 → 2026/04〜2027/03）。value は 'yyyy-MM'。 */
function webFyMonths_(fy) {
  const start = getForecastFYStart_(fy);
  const out = [];
  for (let i = 0; i < 12; i++) {
    const d = addMonths_(start, i);
    out.push({
      value: Utilities.formatDate(d, TZ, 'yyyy-MM'),
      label: Utilities.formatDate(d, TZ, 'yyyy/MM')
    });
  }
  return out;
}

function webParseSteps_(ss) {
  const defs = [
    ['step1_status', 'A-2', '売上データ取り込み'],
    ['step1a_status', 'A-3', '売上データ加工'],
    ['step2_status', 'B-1', '検証実績取り込み'],
    ['step3_status', 'A-4', 'AI調査'],
    ['step3a_status', 'A-4', 'AI調査詳細'],
    ['step4_status', 'A-9', '予測実行'],
    ['step5_status', 'B-2', '検証レポート'],
    ['step6_status', 'B-3', 'ダッシュボード'],
    ['step7_status', 'B-4', '学習インサイト'],
    ['learn_status', 'B-5', '自動学習'],
    ['quarterly_review_status', 'C-1', '四半期レビュー']
  ];
  const sh = ss.getSheetByName(SHEETS.PROCESS_STATUS);
  const map = {};
  if (sh && sh.getLastRow() >= 2) {
    const vals = sh.getRange(1, 1, sh.getLastRow(), 7).getValues();
    for (let i = 1; i < vals.length; i++) {
      const k = String(vals[i][0] || '');
      if (!k) continue;
      map[k] = {
        status: String(vals[i][3] || 'not_run'),
        by: String(vals[i][2] || ''),
        date: vals[i][1] ? Utilities.formatDate(new Date(vals[i][1]), TZ, 'MM/dd') : '',
        count: Number(vals[i][5] || 0) || 0,
        error: String(vals[i][6] || '')
      };
    }
  }
  return defs.map(d => {
    const s = map[d[0]] || { status: 'not_run', by: '', date: '', count: 0, error: '' };
    return { key: d[0], menu: d[1], label: d[2], status: s.status, by: s.by, date: s.date, count: s.count, error: s.error };
  });
}

function webProductsList_(ss) {
  const sh = ss.getSheetByName(SHEETS.PRODUCT);
  const out = [];
  const seen = {};
  if (sh && sh.getLastRow() >= 2) {
    const vals = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues();
    vals.forEach(r => {
      const p = String(r[0] || '').trim();
      if (p && !seen[p]) { seen[p] = 1; out.push(p); }
    });
  }
  return out;
}

/** 入力4シートを現在値で返す（保存時は全面書き換え）。 */
function webParseInputs_(ss) {
  const fy = webCurrentFy_(ss);
  const fyStart = getForecastFYStart_(fy);
  const fyEnd = getForecastFYEnd_(fy);
  const inFy = (d) => d && d >= fyStart && d <= fyEnd;

  const readRows = (name, cols) => {
    const sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return [];
    return sh.getRange(2, 1, sh.getLastRow() - 1, cols).getValues()
      .filter(r => r.some(v => v !== '' && v !== null));
  };
  const ym = (v) => { const d = toDate_(v); return d ? Utilities.formatDate(d, TZ, 'yyyy-MM') : ''; };
  const step = (v) => normalizeStepDisplay_(v);

  return {
    product: readRows(SHEETS.PRODUCT, 5).map(r => ({
      person: String(r[0] || ''), product: String(r[1] || ''), ym: ym(r[2]),
      step: step(r[3]) === null ? String(r[3] || '') : step(r[3]), reason: String(r[4] || '')
    })),
    client: readRows(SHEETS.CLIENT, 4).map(r => ({
      person: String(r[0] || ''), ym: ym(r[1]),
      step: step(r[2]) === null ? String(r[2] || '') : step(r[2]), reason: String(r[3] || '')
    })),
    opinions: readRows(SHEETS.OPINIONS, 5).map(r => ({
      person: String(r[0] || ''), ym: ym(r[1]),
      step: step(r[2]) === null ? String(r[2] || '') : step(r[2]),
      conf: r[3] === '' ? '' : Number(r[3]), note: String(r[4] || '')
    })),
    devspot: readRows(SHEETS.DEV_SPOT, 5).map(r => ({
      person: String(r[0] || ''), ym: ym(r[1]), project: String(r[2] || ''),
      amount: r[3] === '' ? '' : Number(r[3]), conf: r[4] === '' ? '' : Number(r[4])
    }))
  };
}

function webCurrentFy_(ss) {
  const cfg = ss.getSheetByName(SHEETS.CONFIG);
  return Number(cfg ? cfg.getRange('B3').getValue() : 0) || getDefaultFY_();
}

/**
 * OUTPUT シートをアンカー走査でパースする。
 * 行位置は writeSectionBlock_ 等の出力レイアウトに依存するため、固定行番号ではなく
 * 「年度合計（予測）」「Scenario Split」等のラベルを探してセクションを切り出す。
 * 予算セル編集用に月次行の実行番号（1-indexed）も返す。
 */
function webParseOutput_(ss) {
  const sh = ss.getSheetByName(SHEETS.OUTPUT);
  if (!sh || sh.getLastRow() < 10) return null;
  const vals = sh.getDataRange().getValues();
  if (!vals.length) return null;

  const A = (i) => String(vals[i][0] || '');
  const isMonthLabel = (s) => /^20\d\d\/\d{2}$/.test(String(s).trim());
  const num = (v) => (v === '' || v === null || v === undefined) ? null : Number(v);
  const findRow = (needle, from) => {
    for (let i = (from || 0); i < vals.length; i++) if (A(i).indexOf(needle) === 0) return i;
    return -1;
  };
  const findRowContains = (needle, from) => {
    for (let i = (from || 0); i < vals.length; i++) if (A(i).indexOf(needle) !== -1) return i;
    return -1;
  };

  const out = {
    title: A(0),
    policyLines: [],
    engineNote: '',
    composition: [],
    trendInsight: '',
    aiInsight: '',
    aiScores: [],
    aiScoreDetail: '',
    aiScoreNote: '',
    sections: [],
    breakdown: [],
    diagnostics: null
  };

  // 冒頭テキスト（行2〜6）
  for (let i = 1; i <= 5 && i < vals.length; i++) {
    const t = A(i).trim();
    if (t) out.policyLines.push(t.replace(/\n+/g, ' '));
  }
  const noteRow = findRowContains('経過月は実績', 1);
  if (noteRow >= 0) out.engineNote = A(noteRow);

  // 予測構成サマリー（行8-10相当）
  const compLabelRow = findRowContains('予測構成サマリー', 1);
  if (compLabelRow >= 0) {
    for (let i = compLabelRow + 2; i <= compLabelRow + 3 && i < vals.length; i++) {
      const r = vals[i];
      if (!isFinite(Number(r[0]))) break;
      out.composition.push({
        quant: num(r[0]), subjective: num(r[1]), ai: num(r[2]),
        spot: num(r[3]), nonQuantNet: num(r[4]), warn: String(r[5] || '')
      });
    }
  }

  // Insight / AIスコア
  for (let i = 0; i < vals.length && i < 60; i++) {
    const a = A(i);
    if (a.indexOf('数値トレンド Insight') !== -1) out.trendInsight = String(vals[i][1] || '');
    if (a.indexOf('AI調査 Insight') !== -1) out.aiInsight = String(vals[i][1] || '');
    if (['Market', 'Competitor', 'Channel', 'DX'].indexOf(a) >= 0 && isFinite(Number(vals[i][1]))) {
      out.aiScores.push({ name: a, score: Number(vals[i][1]) });
    }
    if (a.indexOf('Market: coverage') !== -1) out.aiScoreDetail = a;
    if (a.indexOf('参考（いずれも正常範囲）') !== -1 || a.indexOf('参考（') === 0) out.aiScoreNote = a;
  }

  // セクション（年度合計（予測）が2箇所: 混合=主要 / 客観=副次）
  const annualIdx = [];
  for (let i = 0; i < vals.length; i++) if (A(i) === '年度合計（予測）') annualIdx.push(i);
  annualIdx.forEach((ai, k) => {
    const labelRow = ai - 2 >= 0 ? A(ai - 2) : '';
    const annual = vals[ai].map(num);
    const months = [];
    // 年間行〜最初の月行の間には空行・月次ヘッダ等が入るため、月ラベルが現れるまで前方に走査する
    let i = ai + 1;
    while (i < vals.length && i <= ai + 6 && !isMonthLabel(A(i))) i++;
    while (i < vals.length && isMonthLabel(A(i))) {
      const r = vals[i];
      months.push({
        row: i + 1, month: String(r[0]).trim(),
        p10: num(r[1]), p50: num(r[2]), p90: num(r[3]),
        reg: num(r[4]), seasonal: num(r[5]), range: num(r[6]),
        adopted: k === 0 ? num(r[7]) : null,
        uplift: k === 0 ? num(r[8]) : null,
        final: k === 0 ? num(r[9]) : null
      });
      i++;
    }
    out.sections.push({
      label: labelRow.replace(/^\s+|\s+$/g, ''),
      annual: { row: ai + 1, p10: annual[1], p50: annual[2], p90: annual[3], reg: annual[4], seasonal: annual[5], range: annual[6], adopted: k === 0 ? annual[7] : null, uplift: k === 0 ? annual[8] : null, final: k === 0 ? annual[9] : null },
      monthly: months
    });
  });

  // Scenario Split（混合セクションの BASE/SPOT 内訳）
  const scRow = findRowContains('Scenario Split', 0);
  if (scRow >= 0) {
    const rows = [];
    for (let i = scRow + 2; i < vals.length; i++) {
      const d = toDate_(vals[i][0]);
      if (!d) break;
      rows.push({
        month: Utilities.formatDate(d, TZ, 'yyyy/MM'),
        p10base: num(vals[i][1]), p10spot: num(vals[i][2]),
        p50base: num(vals[i][3]), p50spot: num(vals[i][4]),
        p90base: num(vals[i][5]), p90spot: num(vals[i][6])
      });
    }
    out.scenario = rows;
  }

  // 内訳表（参考）
  const bdRow = findRowContains('（参考）内訳とメモ', 0);
  if (bdRow >= 0) {
    for (let i = bdRow + 3; i < vals.length; i++) {
      const d = toDate_(vals[i][0]);
      if (!d) break;
      out.breakdown.push({
        month: Utilities.formatDate(d, TZ, 'yyyy/MM'),
        actualClosed: num(vals[i][1]), source: String(vals[i][2] || ''),
        opsQuant: num(vals[i][3]), opsMixed: num(vals[i][4]),
        bgSpot: num(vals[i][5]), knownSpot: num(vals[i][6]),
        totalQuant: num(vals[i][7]), totalMixed: num(vals[i][8]),
        diff: num(vals[i][9]), opinions: String(vals[i][10] || '')
      });
    }
  }

  // Diagnostics（末尾KPI）
  const diagRow = findRow('Diagnostics', 0);
  if (diagRow >= 0 && diagRow + 2 < vals.length) {
    const h = vals[diagRow + 1], v = vals[diagRow + 2];
    out.diagnostics = {};
    for (let j = 0; j < h.length; j++) {
      const key = String(h[j] || '').trim();
      if (key) out.diagnostics[key] = v[j];
    }
    if (diagRow + 4 < vals.length) {
      const h2 = vals[diagRow + 3], v2 = vals[diagRow + 4];
      out.diagnostics.rangeMonthlyAvg = v2[0];
      out.diagnostics.rangeAnnual = v2[1];
      out.diagnostics.rangeWarning = String(v2[2] || '');
    }
  }

  if (!out.sections.length) return null;
  return out;
}

/** DASHBOARD + EVAL_COMPARE_MONTHLY + EVAL_INSIGHTS を返す。 */
function webParseEval_(ss) {
  const res = { metrics: [], compare: [], insights: [], hasActuals: false };

  const dash = ss.getSheetByName(SHEETS.DASHBOARD);
  if (dash && dash.getLastRow() >= 2) {
    dash.getRange(2, 1, dash.getLastRow() - 1, 3).getValues().forEach(r => {
      if (String(r[0] || '').trim()) res.metrics.push({ name: String(r[0]), value: r[1], note: String(r[2] || '') });
    });
  }

  const cmp = ss.getSheetByName(SHEETS.EVAL_COMPARE_MONTHLY);
  if (cmp && cmp.getLastRow() >= 2) {
    res.hasActuals = true;
    cmp.getRange(2, 1, cmp.getLastRow() - 1, 20).getValues().forEach(r => {
      const d = toDate_(r[0]);
      res.compare.push({
        month: d ? Utilities.formatDate(d, TZ, 'yyyy/MM') : String(r[0] || ''),
        forecastTotal: numOrNull_(r[3]), actualTotal: numOrNull_(r[6]),
        gap: numOrNull_(r[7]), p10: numOrNull_(r[8]), p50: numOrNull_(r[9]), p90: numOrNull_(r[10]),
        signedErr: numOrNull_(r[11]), absErr: numOrNull_(r[12]), ape: numOrNull_(r[13]),
        quarter: String(r[14] || ''), half: String(r[15] || ''),
        over: !!r[17], rangeOutside: !!r[19]
      });
    });
  }

  const ins = ss.getSheetByName(SHEETS.EVAL_INSIGHTS);
  if (ins && ins.getLastRow() >= 2) {
    ins.getRange(2, 1, ins.getLastRow() - 1, 24).getValues().forEach((r, i) => {
      if (!String(r[0] || '').trim() && !String(r[2] || '').trim()) return;
      const d = toDate_(r[2]);
      res.insights.push({
        row: i + 2,
        month: d ? Utilities.formatDate(d, TZ, 'yyyy/MM') : String(r[2] || ''),
        actual: numOrNull_(r[3]), pred: numOrNull_(r[4]), errRate: numOrNull_(r[6]),
        insight: String(r[7] || ''), nextAction: String(r[8] || ''),
        annualBreach: !!r[10], halfBreach: !!r[11], overBreach: !!r[12], rangeBreach: !!r[13],
        hypothesis: String(r[14] || ''), causeBucket: String(r[15] || ''),
        actionType: String(r[18] || ''), reflection: String(r[19] || ''),
        owner: String(r[20] || ''), status: String(r[22] || '')
      });
    });
  }
  return res;
}

function numOrNull_(v) { return (v === '' || v === null || v === undefined) ? null : Number(v); }

/** QUARTERLY_REVIEW + QUARTERLY_REVIEW_LOG を返す。 */
function webParseQuarterly_(ss) {
  const res = { title: '', period: '', reviewId: '', proposals: [], applied: false, logRecent: [] };
  const sh = ss.getSheetByName(SHEETS.QUARTERLY_REVIEW);
  if (sh && sh.getLastRow() >= 8) {
    res.title = String(sh.getRange(1, 1).getValue() || '');
    res.period = String(sh.getRange(2, 1).getValue() || '');
    res.reviewId = String(sh.getRange(8, 10).getValue() || '').trim();
    const n = sh.getLastRow() - 7;
    sh.getRange(8, 1, n, 9).getValues().forEach((r, i) => {
      if (!String(r[0] || '').trim()) return;
      res.proposals.push({
        row: i + 8, pid: String(r[0] || ''), target: String(r[1] || ''),
        current: String(r[2] || ''), proposed: String(r[3] || ''),
        conf: numOrNull_(r[4]), rationale: String(r[5] || ''),
        impact: String(r[6] || ''), decision: String(r[7] || ''), rollback: String(r[8] || '')
      });
    });
  }
  const log = ss.getSheetByName(SHEETS.QUARTERLY_REVIEW_LOG);
  if (log && log.getLastRow() >= 2) {
    const vals = log.getRange(1, 1, log.getLastRow(), 21).getValues();
    const idx = {};
    vals[0].forEach((h, i) => { idx[String(h || '').trim()] = i; });
    if (res.reviewId) {
      res.applied = vals.slice(1).some(r => String(r[idx.review_id] || '') === res.reviewId && Number(r[idx.applied] || 0) === 1);
    }
    // 直近のレビュー単位でまとめて返す
    const byReview = {};
    vals.slice(1).forEach(r => {
      const rid = String(r[idx.review_id] || '');
      if (!rid) return;
      if (!byReview[rid]) byReview[rid] = { reviewId: rid, at: r[idx.reviewed_at], client: String(r[idx.client] || ''), quarter: String(r[idx.quarter_label] || ''), n: 0, applied: false };
      byReview[rid].n++;
      if (Number(r[idx.applied] || 0) === 1) byReview[rid].applied = true;
    });
    res.logRecent = Object.keys(byReview).map(k => byReview[k])
      .sort((a, b) => new Date(b.at) - new Date(a.at)).slice(0, 5)
      .map(r => ({ reviewId: r.reviewId.slice(0, 8), quarter: r.quarter, proposals: r.n, applied: r.applied, at: r.at ? Utilities.formatDate(new Date(r.at), TZ, 'MM/dd') : '' }));
  }
  return res;
}

function webParseRunLog_(ss) {
  const sh = ss.getSheetByName(SHEETS.RUN_LOG);
  if (!sh || sh.getLastRow() < 2) return [];
  const n = Math.min(8, sh.getLastRow() - 1);
  const vals = sh.getRange(sh.getLastRow() - n + 1, 1, n, 12).getValues();
  return vals.map(r => ({
    at: r[1] ? Utilities.formatDate(new Date(r[1]), TZ, 'MM/dd HH:mm') : '',
    by: String(r[2] || '').split('@')[0],
    fn: String(r[3] || ''), status: String(r[5] || ''), count: Number(r[6] || 0) || 0
  })).reverse();
}

// ===== setup =====

/** A-1 相当：設定の保存（シートの全消去は Web では行わない）。 */
function webSaveSetup(p) {
  return webAudited_('SETUP.SAVE', () => {
    const client = String(p && p.client || '').trim();
    const fy = Number(p && p.fy);
    const people = String(p && p.peopleCsv || '').trim();
    if (!client) throw new Error('クライアント（メーカー名）を選択してください。');
    if (!fy || !isFinite(fy)) throw new Error('予測年度（FY）を選択してください。');
    saveInitialSetupSettings(client, String(fy), people);
    return webGetBootstrap_();
  }, { detail: { client: p && p.client, fy: p && p.fy, peopleCsv: p && p.peopleCsv },
    before: () => { const cfg = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CONFIG); return cfg ? { client: cfg.getRange('B2').getValue(), fy: cfg.getRange('B3').getValue(), peopleCsv: cfg.getRange('B4').getValue() } : null; } });
}

/** セットアップのクライアント候補（外部実績SSから遅延取得）。 */
function webGetClientCandidates() {
  const list = getClientCandidatesForSetup_();
  return { candidates: list };
}

// ===== 準備アクション =====

/** A-2 相当：外部SSから売上を取り込み、入力シートをFY向けに整える。 */
function webRunImportSales() {
  return webAudited_('IMPORT.SALES', () => {
    ensureSetupDone_();
    const res = importMonthlyFromExternal_(SHEETS.SALES_INPUT, true);
    const fy = webCurrentFy_(SpreadsheetApp.getActiveSpreadsheet());
    refreshManualInputSheets_(fy);
    hideNonUserSheets_();
    return { count: res.count, range: res.range, boot: webGetBootstrap_() };
  }, { after: (res) => ({ count: res.count, range: res.range }) });
}

/** A-3 相当：SALES_INPUT → SALES_MONTHLY への集計。 */
function webRunAggregate() {
  return webAudited_('SALES.AGGREGATE', () => {
    ensureSetupDone_();
    requireStepSuccess_('step1_status', '先に A-2 売上データ取り込みを実行してください。');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cfgSh = ss.getSheetByName(SHEETS.CONFIG);
    const client = normalizeClientName_(String(cfgSh.getRange('B2').getValue() || '').trim());
    if (!client) throw new Error('CONFIG!B2 にクライアントを設定してください。');
    const fy = Number(cfgSh.getRange('B3').getValue()) || getDefaultFY_();
    try {
      const count = syncSalesFromSalesInput_(fy, client);
      hideNonUserSheets_();
      try { updateProcessStatus_('step1a_status', 'success', client, count, ''); } catch (e2) { /* ステータス更新失敗は握り潰す */ }
      return { count: count, boot: webGetBootstrap_() };
    } catch (e) {
      try { updateProcessStatus_('step1a_status', 'error', client, 0, String(e && e.message || e)); } catch (e2) { /* 同上 */ }
      throw e;
    }
  }, { after: (res) => ({ count: res.count }) });
}

/** A-4 相当：Vertex AI 調査（長時間・数分）。 */
function webRunAiResearch() {
  return webAudited_('AI.RESEARCH', () => {
    ensureSetupDone_();
    const before = countAIResearchStructuredRows_();
    runVertexAIResearch(); // 失敗は alertOrThrow_ 経由で例外化
    const after = countAIResearchStructuredRows_();
    return { rows: after, changed: after !== before, boot: webGetBootstrap_() };
  }, { after: (res) => ({ rows: res.rows, changed: res.changed }) });
}

// ===== 入力保存 =====

/** 入力シートを全面書き換え（読み取り側は必須欠落行を自動スキップする既存仕様に準拠）。 */
function webSaveInputs(kind, rows) {
  return webAudited_('INPUT.SAVE', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const spec = {
      product:  { name: SHEETS.PRODUCT,  cols: 5, make: (r) => [r.person || '', r.product || '', webYmToDate_(r.ym), String(r.step || ''), r.reason || ''] },
      client:   { name: SHEETS.CLIENT,   cols: 4, make: (r) => [r.person || '', webYmToDate_(r.ym), String(r.step || ''), r.reason || ''] },
      opinions: { name: SHEETS.OPINIONS, cols: 5, make: (r) => [r.person || '', webYmToDate_(r.ym), String(r.step || ''), r.conf === '' || r.conf === null ? '' : Number(r.conf), r.note || ''] },
      devspot:  { name: SHEETS.DEV_SPOT, cols: 5, make: (r) => [r.person || '', webYmToDate_(r.ym), r.project || '', r.amount === '' || r.amount === null ? '' : Number(r.amount), r.conf === '' || r.conf === null ? '' : Number(r.conf)] }
    }[String(kind || '')];
    if (!spec) throw new Error('不明な入力種別: ' + kind);
    const sh = ss.getSheetByName(spec.name);
    if (!sh) throw new Error(spec.name + ' シートがありません。先に A-2 を実行してください。');

    const clean = (rows || []).filter(r => r && Object.keys(r).some(k => String(r[k] || '').trim() !== ''))
      .map(r => spec.make(r));
    if (clean.length > 500) throw new Error('行数が上限（500）を超えています。');
    // 数値列の検証（空は許容、数値化不能はエラー）
    clean.forEach(r => {
      r.forEach((v, j) => {
        if (typeof v === 'number' && !isFinite(v)) throw new Error('数値でない値が含まれています（' + spec.name + '）。');
      });
    });

    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, spec.cols).clearContent();
    if (clean.length) sh.getRange(2, 1, clean.length, spec.cols).setValues(clean);
    SpreadsheetApp.flush();
    return { saved: clean.length, input: webParseInputs_(ss) };
  }, { detail: { kind: kind, rows: (rows || []).length },
    before: () => (webParseInputs_(SpreadsheetApp.getActiveSpreadsheet())[String(kind || '')] || null),
    after: (res) => ({ saved: res.saved, rows: res.input && res.input[String(kind || '')] }) });
}

function webYmToDate_(ym) {
  const m = /^(\d{4})-(\d{2})$/.exec(String(ym || '').trim());
  if (!m) return '';
  return new Date(Number(m[1]), Number(m[2]) - 1, 1);
}

// ===== 予測 =====

/**
 * A-9 相当：予測実行。確認が必要な入力（48ヶ月未満／極端な入力）があると
 * {needConfirm:{key,title,message}} を返し、クライアント確認後に
 * confirms=[key] を付けて再実行する。
 */
function webRunForecast(confirms) {
  return webAudited_('FORECAST.RUN', () => {
    Object.keys(WEB_UI_CONFIRMS_).forEach(k => { delete WEB_UI_CONFIRMS_[k]; });
    (confirms || []).forEach(k => { WEB_UI_CONFIRMS_[String(k)] = true; });
    try {
      runPhase1Forecast();
    } catch (e) {
      // 確認要否は例外ではなく戻り値で返す（google.script.run は throw の custom props を保持しない）
      if (e && e.webConfirm) return { needConfirm: e.webConfirm };
      throw e;
    } finally {
      Object.keys(WEB_UI_CONFIRMS_).forEach(k => { delete WEB_UI_CONFIRMS_[k]; });
    }
    return { boot: webGetBootstrap_() };
  }, { detail: { confirms: confirms || [] },
    after: (res) => (res && res.needConfirm ? { needConfirm: res.needConfirm.key } : { done: true }) });
}

/** A-10 相当：OUTPUT の予算列（H=Adopted / I=Uplift）を保存。J列（Final）は数式のまま触らない。 */
function webSaveBudget(rows) {
  return webAudited_('BUDGET.SAVE', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEETS.OUTPUT);
    if (!sh) throw new Error('OUTPUT がありません。先に予測を実行してください。');
    if (!Array.isArray(rows) || !rows.length) throw new Error('保存する予算行がありません。');

    // 月次行の正当性をサーバ側でも確認（行番号がずれていたら書き込まない）
    const vals = sh.getDataRange().getValues();
    const writes = [];
    rows.forEach(r => {
      const row = Number(r && r.row);
      if (!row || row < 1 || row > vals.length) return;
      const label = String(vals[row - 1][0] || '').trim();
      if (!/^20\d\d\/\d{2}$/.test(label)) return; // 月次行以外には書かない
      const adopted = r.adopted === '' || r.adopted === null || r.adopted === undefined ? '' : Number(r.adopted);
      const uplift = r.uplift === '' || r.uplift === null || r.uplift === undefined ? '' : Number(r.uplift);
      if ((adopted !== '' && !isFinite(adopted)) || (uplift !== '' && !isFinite(uplift))) throw new Error('予算は数値で入力してください（' + label + '）');
      writes.push({ row: row, adopted: adopted, uplift: uplift });
    });
    if (!writes.length) throw new Error('有効な月次行が見つかりませんでした。予測を再実行してください。');

    writes.forEach(w => {
      sh.getRange(w.row, 8).setValue(w.adopted);
      sh.getRange(w.row, 9).setValue(w.uplift);
    });
    SpreadsheetApp.flush();
    return { saved: writes.length, output: webParseOutput_(ss) };
  }, { detail: { rows: (rows || []).length },
    before: () => webCellsSnapshot_(SHEETS.OUTPUT, rows, [1, 8, 9]),
    after: () => webCellsSnapshot_(SHEETS.OUTPUT, rows, [1, 8, 9]) });
}

// ===== 検証 =====

/** B-1 相当：検証用実績の取り込み。 */
function webRunImportActuals() {
  return webAudited_('IMPORT.ACTUALS', () => {
    ensureSetupDone_();
    requireStepSuccess_('step1_status', '先に A-2 売上データ取り込みを実行してください。');
    const res = importMonthlyFromExternal_(SHEETS.ACTUAL_EVAL_MONTHLY, true);
    hideNonUserSheets_();
    return { count: res.count, range: res.range, boot: webGetBootstrap_() };
  }, { after: (res) => ({ count: res.count, range: res.range }) });
}

/** B-2 相当：検証レポート更新。 */
function webRunEvalReport() {
  return webAudited_('EVAL.REPORT', () => {
    updatePhase1EvaluationReport();
    hideNonUserSheets_();
    return { boot: webGetBootstrap_() };
  });
}

/** B-3 相当：ダッシュボード更新。 */
function webRunDashboard() {
  return webAudited_('EVAL.DASHBOARD', () => {
    updatePhase1Dashboard();
    hideNonUserSheets_();
    return { boot: webGetBootstrap_() };
  });
}

/** B-4 相当：学習インサイト更新。 */
function webRunInsights() {
  return webAudited_('EVAL.INSIGHTS', () => {
    updatePhase1LearningInsights();
    hideNonUserSheets_();
    return { boot: webGetBootstrap_() };
  });
}

/** EVAL_INSIGHTS の記入列を保存（原因仮説・対応など）。 */
function webSaveEvalInsights(rows) {
  return webAudited_('INSIGHT.SAVE', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEETS.EVAL_INSIGHTS);
    if (!sh) throw new Error('EVAL_INSIGHTS がありません。先に B-4 を実行してください。');
    (rows || []).forEach(r => {
      const row = Number(r && r.row);
      if (!row || row < 2 || row > sh.getLastRow()) return;
      // 行がデータ行であることを確認（C-1/B-4再実行で行がずれた場合に誤書き込みしない）
      const head = String(sh.getRange(row, 1).getValue() || '').trim();
      const head2 = String(sh.getRange(row, 3).getValue() || '').trim();
      if (!head && !head2) return;
      sh.getRange(row, 15).setValue(String(r.hypothesis || ''));   // cause_hypothesis
      sh.getRange(row, 19).setValue(String(r.actionType || ''));  // action_type
      sh.getRange(row, 20).setValue(String(r.reflection || ''));  // next_cycle_reflection
      sh.getRange(row, 21).setValue(String(r.owner || ''));       // owner
      sh.getRange(row, 23).setValue(String(r.status || ''));      // status
    });
    SpreadsheetApp.flush();
    return { saved: (rows || []).length };
  }, { detail: { rows: (rows || []).length },
    before: () => webCellsSnapshot_(SHEETS.EVAL_INSIGHTS, rows, [3, 15, 19, 20, 21, 23], 2),
    after: () => webCellsSnapshot_(SHEETS.EVAL_INSIGHTS, rows, [3, 15, 19, 20, 21, 23], 2) });
}

// ===== 四半期レビュー =====

/** C-1 相当：四半期レビュー提案の生成。 */
function webRunQuarterly() {
  return webAudited_('REVIEW.GENERATE', () => {
    runQuarterlyReview();
    return { quarterly: webParseQuarterly_(SpreadsheetApp.getActiveSpreadsheet()) };
  });
}

/** C-2 相当（Web版）：承認列（承認/却下/保留）を保存。 */
function webSaveQuarterlyDecisions(rows) {
  return webAudited_('REVIEW.DECIDE', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEETS.QUARTERLY_REVIEW);
    if (!sh) throw new Error('QUARTERLY_REVIEW がありません。先に C-1 を実行してください。');
    const allowed = { '承認': 1, '却下': 1, '保留': 1 };
    (rows || []).forEach(r => {
      const row = Number(r && r.row);
      if (!row || row < 8 || row > sh.getLastRow()) return;
      // 提案行（pid非空）以外には書き込まない
      if (!String(sh.getRange(row, 1).getValue() || '').trim()) return;
      const d = String(r.decision || '');
      if (d && !allowed[d]) throw new Error('承認列は「承認 / 却下 / 保留」から選択してください。');
      sh.getRange(row, 8).setValue(d);
    });
    SpreadsheetApp.flush();
    return { saved: (rows || []).length };
  }, { detail: { rows: (rows || []).length },
    before: () => webCellsSnapshot_(SHEETS.QUARTERLY_REVIEW, rows, [1, 8], 8),
    after: () => webCellsSnapshot_(SHEETS.QUARTERLY_REVIEW, rows, [1, 8], 8) });
}

/** C-3 相当：承認済み提案を適用。 */
function webApplyQuarterly() {
  return webAudited_('REVIEW.APPLY', () => {
    const res = applyQuarterlyProposals();
    return { result: res || null, quarterly: webParseQuarterly_(SpreadsheetApp.getActiveSpreadsheet()) };
  }, { after: (res) => res && res.result });
}

// ===== 自動学習（B-5）/ Vertexアシスト / 学習バックテスト =====

/** CALIBRATION_STATE + VERTEX_FORECAST_LOG + SOURCE_RELIABILITY を1か所に集約して返す。 */
function webParseLearning_(ss, client) {
  const res = {
    autoUpdate: true,
    biasFactor: 1.0,
    monthBias: {},
    vertex: { enabled: false, hasLog: false, at: '', confidence: 0, rationale: '', adj: [], months: [] },
    vertexReliability: 1.0,
    vertexWeight: 0.5,
    reliabilityNonDefault: 0,
    lastAppliedQuarter: '',
    note: ''
  };
  try {
    const cal = readCalibrationState_(client);
    res.autoUpdate = calibrationNumberOr_(cal.auto_update_enabled, 1) === 1;   // 0 を 1 にしない
    res.biasFactor = isFinite(Number(cal.bias_correction_factor)) ? Number(cal.bias_correction_factor) : 1.0;
    res.monthBias = parseResidualMonthBiasJson_(cal.residual_month_bias_json);
    res.lastAppliedQuarter = String(cal.last_applied_quarter || '');
    res.note = String(cal.note || '');
    res.vertex.enabled = readVertexForecastEnabled_();
    res.vertexWeight = readVertexAssistWeight_(readModelTuningFromConfig_());
    const relMap = readSourceReliability_(client);
    res.reliabilityNonDefault = Array.from(relMap.values()).filter(v => Math.abs(Number(v || 1) - 1) > 1e-9).length;
    res.vertexReliability = getSourceReliability_(relMap, 'vertex_forecast', 'assist');
    const sh = ss.getSheetByName(SHEETS.VERTEX_FORECAST_LOG);
    if (sh && sh.getLastRow() >= 2) {
      const vals = sh.getDataRange().getValues();
      const idx = headerIndexMap_(vals[0] || []);
      for (let i = vals.length - 1; i >= 1; i--) {
        const r = vals[i];
        if (client && !isSameClient_(r[idx.client], client)) continue;
        if (String(r[idx.status] || '') !== 'ok') continue;
        res.vertex.hasLog = true;
        const at = r[idx.run_at];
        res.vertex.at = at instanceof Date ? Utilities.formatDate(at, TZ, 'yyyy-MM-dd') : String(at || '');
        res.vertex.confidence = Number(r[idx.confidence] || 0);
        res.vertex.rationale = String(r[idx.rationale_ja] || '').slice(0, 300);
        try { res.vertex.adj = JSON.parse(String(r[idx.monthly_adj_json] || '[]')); } catch (e) { res.vertex.adj = []; }
        try { res.vertex.months = JSON.parse(String(r[idx.target_months_json] || '[]')); } catch (e) { res.vertex.months = []; }
        break;
      }
    }
  } catch (e) {
    res.error = String(e && e.message || e);
  }
  return res;
}

/** B-5 相当：月次ベイズ自動学習を実行。 */
function webRunMonthlyLearn() {
  return webAudited_('LEARN.MONTHLY', () => {
    ensureSetupDone_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const client = String(ss.getSheetByName(SHEETS.CONFIG).getRange('B2').getValue() || '').trim();
    const res = runMonthlyAutoLearn_(client, {});
    if (res && res.ready) {
      updateProcessStatus_('learn_status', 'success', client, res.n, '');
    } else if (res) {
      updateProcessStatus_('learn_status', 'success', client, res.n || 0, `skipped:${res.skipped}`);
    }
    return { result: res || null, boot: webGetBootstrap_() };
  }, { after: (res) => res && res.result });
}

/** Vertexアシスト単体実行（A-4と同じログを残し、A-9は次回実行時に自動反映）。 */
function webRunVertexAssist() {
  return webAudited_('AI.ASSIST', () => {
    ensureSetupDone_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const client = String(ss.getSheetByName(SHEETS.CONFIG).getRange('B2').getValue() || '').trim();
    const vertex = readVertexConfig_();
    if (!vertex.geminiReady) throw new Error('Vertex の必須設定が未入力です（CONFIG: VERTEX_PROJECT_ID / VERTEX_LOCATION / VERTEX_GEMINI_MODEL）。');
    ensureAIResearchRuntimeSheets_(ss);
    const res = runVertexForecastAssist_(client, ss, vertex, null);
    if (res.skipped === 'disabled') throw new Error('CONFIG の VERTEX_FORECAST_ENABLED を 1 にしてください。');
    if (!res.ok) throw new Error('Vertexアシスト失敗: ' + (res.error || res.skipped || 'unknown'));
    safeLogRun_('runVertexForecastAssist_', client, 'success', 12, new Date(), `manual conf=${res.confidence}`);
    return { result: res, boot: webGetBootstrap_() };
  }, { after: (res) => ({ ok: res.result && res.result.ok, confidence: res.result && res.result.confidence }) });
}

/** カウンターファクト学習バックテスト（EVAL_LOG walk-forward 簡易検算）。 */
function webRunLearningBacktest() {
  ensureSetupDone_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const client = String(ss.getSheetByName(SHEETS.CONFIG).getRange('B2').getValue() || '').trim();
  return { result: runLearningBacktest_(client, {}) };
}

// ===== 操作の記録（監査ログ）と実行者の確認 — 段階0（DESIGN_data_platform_JA.md 10章） =====
// Web からの書き込み・実行は webAudited_ を通す。実行前の「開始」を記録できなければ処理しない（fail-closed）。
// 記録先は年度ごとのログ用スプレッドシート「売上予測 ログ FYyyyy」（Script Property に ID を控え、無ければ初回に作る）。
// 月ごとのシート AUDIT_yyyy_MM に 1 行ずつ追記し、row_hash = SHA-256(prev_hash + 行の内容) の鎖でつなぐ（改ざん・削除の検知用）。
// 管理者 = スクリプトの所有者（デプロイした人）＋ Script Property FORECAST_WEB_ADMIN_EMAILS（カンマ区切り）。
// Web アプリの公開範囲は appsscript.json で「自分のみ」。この確認は公開範囲を広げたときの二重の守り。
const WEB_AUDIT_HEADERS_ = ['audit_id', 'occurred_at', 'actor_email', 'action', 'phase', 'result', 'client', 'fy',
  'detail_json', 'before_json', 'after_json', 'error', 'request_id', 'app_version', 'prev_hash', 'row_hash'];
const WEB_AUDIT_JSON_MAX_ = 40000;
const WEB_AUDIT_PROP_FILES_ = 'FORECAST_LOG_SPREADSHEETS_JSON';  // {"FY2026": "<spreadsheetId>"}
const WEB_AUDIT_PROP_FOLDER_ = 'FORECAST_LOG_FOLDER_ID';          // 任意: ログのファイルを置くフォルダ
const WEB_AUDIT_PROP_LAST_HASH_ = 'FORECAST_AUDIT_LAST_HASH';
const WEB_ADMIN_PROP_ = 'FORECAST_WEB_ADMIN_EMAILS';

/** 操作した人と、管理者かどうか。メールが取れなければ管理者として扱わない。 */
function webActor_() {
  let email = '';
  let owner = '';
  try { email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase(); } catch (e) { email = ''; }
  try { owner = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase(); } catch (e) { owner = ''; }
  const extra = String(PropertiesService.getScriptProperties().getProperty(WEB_ADMIN_PROP_) || '')
    .split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
  const admins = [owner].concat(extra).filter(Boolean);
  return { email: email, isAdmin: !!email && admins.indexOf(email) >= 0 };
}

/** このアプリの年度（4月〜翌3月。始まりの年で呼ぶ: 2026-10 → FY2026。getForecastFYStart_・vNextFiscalYearForDate_ と同じ）。 */
function webAuditFy_(d) {
  return d.getMonth() >= 3 ? d.getFullYear() : d.getFullYear() - 1;
}

const WEB_AUDIT_PROP_FY_NAMING_ = 'FORECAST_LOG_FY_NAMING';  // 'start' = 始まりの年で呼ぶ（ずらし済み）

/**
 * ログのファイルを「始まりの年で呼ぶ年度」にそろえる（一度だけ。webAuditAppend_ のロックの中で呼ぶ）。
 * 2026-10-01 まではログを終わりの年で名付けていた（FY2027 = 2026/04〜2027/03）ので、キーとファイル名を 1 年ずらす。
 */
function webAuditEnsureFyNaming_(props) {
  if (props.getProperty(WEB_AUDIT_PROP_FY_NAMING_) === 'start') return;
  let files = {};
  try { files = JSON.parse(props.getProperty(WEB_AUDIT_PROP_FILES_) || '{}') || {}; } catch (e) { files = {}; }
  const next = {};
  Object.keys(files).forEach(k => {
    const m = /^FY(\d{4})$/.exec(k);
    const nk = m ? 'FY' + (Number(m[1]) - 1) : k;
    next[nk] = files[k];
    try { DriveApp.getFileById(files[k]).setName('売上予測 ログ ' + nk); } catch (e) {
      Logger.log('ログのファイル名を直せません: ' + (e && e.message ? e.message : e));
    }
  });
  props.setProperty(WEB_AUDIT_PROP_FILES_, JSON.stringify(next));
  props.setProperty(WEB_AUDIT_PROP_FY_NAMING_, 'start');
}

/** 長い JSON はハッシュと先頭だけにする（セルの上限 5 万字に対し 4 万字まで）。 */
function webAuditJson_(v) {
  if (v === undefined || v === null || v === '') return '';
  const s = typeof v === 'string' ? v : JSON.stringify(v);
  if (s.length <= WEB_AUDIT_JSON_MAX_) return s;
  return JSON.stringify({ truncated: true, length: s.length, sha256: vNextSha256Hex_(s), head: s.slice(0, 2000) });
}

/** その月の監査シート。ログのファイルと月のシートが無ければ作る。列が想定と違えば止める。 */
function webAuditSheet_(now) {
  const props = PropertiesService.getScriptProperties();
  webAuditEnsureFyNaming_(props);
  const fyKey = 'FY' + webAuditFy_(now);
  let files = {};
  try { files = JSON.parse(props.getProperty(WEB_AUDIT_PROP_FILES_) || '{}') || {}; } catch (e) { files = {}; }
  let ss;
  if (files[fyKey]) {
    ss = SpreadsheetApp.openById(files[fyKey]);  // 開けなければ例外（記録できないので処理も止まる）
  } else {
    ss = SpreadsheetApp.create('売上予測 ログ ' + fyKey);
    const folderId = String(props.getProperty(WEB_AUDIT_PROP_FOLDER_) || '').trim();
    if (folderId) DriveApp.getFileById(ss.getId()).moveTo(DriveApp.getFolderById(folderId));
    files[fyKey] = ss.getId();
    props.setProperty(WEB_AUDIT_PROP_FILES_, JSON.stringify(files));
  }
  const name = 'AUDIT_' + Utilities.formatDate(now, TZ, 'yyyy_MM');
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, WEB_AUDIT_HEADERS_.length).setNumberFormat('@').setValues([WEB_AUDIT_HEADERS_]);
    sh.setFrozenRows(1);
    // 新しいファイルに最初からある空のシートは消す（データシートはヘッダー＋データだけにする）
    ss.getSheets().forEach(s => {
      if (s.getName() !== name && s.getLastRow() === 0 && /^(シート|Sheet)\d+$/.test(s.getName())) ss.deleteSheet(s);
    });
  } else {
    const head = sh.getRange(1, 1, 1, WEB_AUDIT_HEADERS_.length).getValues()[0].map(String);
    if (head.join('|') !== WEB_AUDIT_HEADERS_.join('|')) throw new Error('監査ログの列が想定と違います（' + name + '）。');
  }
  return sh;
}

/** 1 行を文字列のまま書く（日付などへの自動変換を防ぐ）。 */
function webAuditWriteRow_(sh, row) {
  const r = sh.getLastRow() + 1;
  if (r > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 1000);
  sh.getRange(r, 1, 1, row.length).setNumberFormat('@').setValues([row]);
}

/** 監査ログに 1 行足し、ハッシュの鎖を伸ばす。書けなければ例外。 */
function webAuditAppend_(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const now = new Date();
    const sh = webAuditSheet_(now);
    const props = PropertiesService.getScriptProperties();
    const prev = String(props.getProperty(WEB_AUDIT_PROP_LAST_HASH_) || '');
    const body = [
      Utilities.getUuid(),
      Utilities.formatDate(now, TZ, "yyyy-MM-dd'T'HH:mm:ssZ"),
      e.actor || '', e.action || '', e.phase || '', e.result || '',
      e.client || '', e.fy ? String(e.fy) : '',
      webAuditJson_(e.detail), webAuditJson_(e.before), webAuditJson_(e.after),
      String(e.error || '').slice(0, 2000), e.requestId || '', e.appVersion || ''
    ];
    const hash = vNextSha256Hex_(prev + '\n' + JSON.stringify(body));
    webAuditWriteRow_(sh, body.concat([prev, hash]));
    props.setProperty(WEB_AUDIT_PROP_LAST_HASH_, hash);
    SpreadsheetApp.flush();
    return hash;
  } finally {
    lock.releaseLock();
  }
}

/** 記録に添える対象（クライアントと年度）。読めなくても処理は止めない。 */
function webAuditContext_() {
  try {
    const cfg = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CONFIG);
    if (!cfg) return { client: '', fy: '' };
    return { client: String(cfg.getRange('B2').getValue() || '').trim(), fy: String(cfg.getRange('B3').getValue() || '').trim() };
  } catch (e) {
    return { client: '', fy: '' };
  }
}

/** 指定した行・列の今の値（変更前の記録用）。 */
function webCellsSnapshot_(sheetName, rows, cols, minRow) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sh) return null;
  const last = sh.getLastRow();
  return (rows || []).map(r => Number(r && r.row)).filter(n => n >= (minRow || 1) && n <= last).slice(0, 200)
    .map(n => ({ row: n, values: cols.map(c => sh.getRange(n, c).getValue()) }));
}

/**
 * 書き込み・実行を記録つきで行う。管理者でなければ拒否（拒否も記録）。
 * 開始を記録 → 実行 → 終了（OK / NEEDS_CONFIRM / FAILED）を記録。開始を記録できなければ実行しない。
 */
function webAudited_(action, fn, opts) {
  opts = opts || {};
  const actor = webActor_();
  const ctx = webAuditContext_();
  const base = { actor: actor.email || '(unknown)', action: action, requestId: Utilities.getUuid(),
    appVersion: VERSION + ' / ' + BUILD_STAGE, client: ctx.client, fy: ctx.fy };
  if (!actor.isAdmin) {
    try { webAuditAppend_(Object.assign({}, base, { phase: 'DENIED', result: 'DENIED', detail: opts.detail })); }
    catch (e) { Logger.log('webAudited_ DENIED の記録に失敗: ' + (e && e.message ? e.message : e)); }
    throw new Error('この操作は管理者だけが実行できます。');
  }
  let before = '';
  try { before = opts.before ? opts.before() : ''; } catch (e) { before = { error: String(e && e.message || e) }; }
  webAuditAppend_(Object.assign({}, base, { phase: 'START', detail: opts.detail, before: before }));
  try {
    const res = fn();
    let after = '';
    try { after = opts.after ? opts.after(res) : ''; } catch (e) { after = { error: String(e && e.message || e) }; }
    const result = res && res.needConfirm ? 'NEEDS_CONFIRM' : 'OK';
    try { webAuditAppend_(Object.assign({}, base, { phase: 'END', result: result, after: after })); }
    catch (e) { Logger.log('webAudited_ END の記録に失敗: ' + (e && e.message ? e.message : e)); }
    return res;
  } catch (err) {
    try { webAuditAppend_(Object.assign({}, base, { phase: 'END', result: 'FAILED', error: String(err && err.message || err) })); }
    catch (e) { Logger.log('webAudited_ FAILED の記録に失敗: ' + (e && e.message ? e.message : e)); }
    throw err;
  }
}

/** 今の年度のログのファイルの URL（管理者の画面のリンク用。無ければ空）。 */
function webAuditLogUrl_() {
  try {
    const files = JSON.parse(PropertiesService.getScriptProperties().getProperty(WEB_AUDIT_PROP_FILES_) || '{}') || {};
    const id = files['FY' + webAuditFy_(new Date())];
    return id ? 'https://docs.google.com/spreadsheets/d/' + id + '/edit' : '';
  } catch (e) {
    return '';
  }
}
