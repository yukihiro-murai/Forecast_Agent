/**
 * Plan.js — 計画の画面（段階2-3）。旧来の Web アプリ（Forecast_WebApp.js）と同じ内容と操作を、データ本体の上で行う。
 *
 * - 見る: データ本体から読むだけのブック（appStoreBook_）を組み立て、旧来の webGetBootstrap_ をそのまま動かす。
 *         計算用ブックを使わないので速い。結果は入力のハッシュごとに 6 時間覚えておく。
 * - 保存（入力・予算・インサイトの記入・四半期レビューの承認）: 使うシートだけを計算用ブックに組み立て、
 *         旧来の webSave* をそのまま動かし、変わったシートをデータ本体へ戻す（1 つの裏の処理）。
 * - 実行（A-3・B-2〜B-5・C-1・C-3）: 予測の実行と同じく、全部のシートを組み立てて旧来の webRun* を動かし、
 *         計算 → 保存の 2 つの裏の処理で戻す（計算の間にデータ本体が変わっていないことを確かめる）。
 * 記録は PLAN_ACTIONS（1 回 1 行）と監査ログ。旧来の webAudited_（旧来のログ）は使わない（LegacyEngine.js で差し替え）。
 * 外部とつなぐもの（A-2・B-1 の取り込み、A-4 の AI 調査、Vertex のアシスト）と A-1 の設定は、まだここに入れていない。
 */

/** 計画への保存・実行（名前は旧来の Web アプリの操作の記録と同じ） */
const APP_PLAN_ACTIONS = {
  'INPUT.SAVE': { kind: 'edit', minRole: 'PLANNER', fn: 'webSaveInputs', label: '入力の保存',
    sheets: ['CONFIG', 'PRODUCT', 'CLIENT', 'OPINIONS', 'DEV_SPOT'], args: a => [a.kind, a.rows] },
  'BUDGET.SAVE': { kind: 'edit', minRole: 'PLANNER', fn: 'webSaveBudget', label: '予算の保存', sheets: ['OUTPUT'], args: a => [a.rows] },
  'INSIGHT.SAVE': { kind: 'edit', minRole: 'PLANNER', fn: 'webSaveEvalInsights', label: 'インサイトの記入の保存', sheets: ['EVAL_INSIGHTS'], args: a => [a.rows] },
  'REVIEW.DECIDE': { kind: 'edit', minRole: 'APPROVER', fn: 'webSaveQuarterlyDecisions', label: '四半期レビューの承認の保存',
    sheets: ['QUARTERLY_REVIEW'], args: a => [a.rows] },
  // A-1 の担当者（クライアントと年度は計画で決まるので変えない）。旧来の saveInitialSetupSettings と同じく CONFIG!B4 に書く
  'SETUP.PEOPLE': { kind: 'edit', minRole: 'ADMIN', local: 'appPlanSetPeople_', label: '担当者の保存', sheets: ['CONFIG'], args: a => [a.peopleCsv] },
  // A-2・B-1: 設定の「ZAC の実績のスプレッドシート」から読む（旧来と同じ関数）。取り込みは管理者（設計 6 章）
  'IMPORT.SALES': { kind: 'run', minRole: 'ADMIN', fn: 'webRunImportSales', label: 'A-2 売上データの取り込み' },
  'IMPORT.ACTUALS': { kind: 'run', minRole: 'ADMIN', fn: 'webRunImportActuals', label: 'B-1 検証用の実績の取り込み' },
  'SALES.AGGREGATE': { kind: 'run', minRole: 'PLANNER', fn: 'webRunAggregate', label: 'A-3 売上データの加工' },
  'EVAL.REPORT': { kind: 'run', minRole: 'PLANNER', fn: 'webRunEvalReport', label: 'B-2 検証レポートの更新' },
  'EVAL.DASHBOARD': { kind: 'run', minRole: 'PLANNER', fn: 'webRunDashboard', label: 'B-3 ダッシュボードの更新' },
  'EVAL.INSIGHTS': { kind: 'run', minRole: 'PLANNER', fn: 'webRunInsights', label: 'B-4 学習インサイトの更新' },
  'LEARN.MONTHLY': { kind: 'run', minRole: 'PLANNER', fn: 'webRunMonthlyLearn', label: 'B-5 月次の自動学習' },
  'REVIEW.GENERATE': { kind: 'run', minRole: 'PLANNER', fn: 'webRunQuarterly', label: 'C-1 四半期レビューの提案' },
  'REVIEW.APPLY': { kind: 'run', minRole: 'APPROVER', fn: 'webApplyQuarterly', label: 'C-3 承認した提案の適用' }
};

function appPlanAction_(name, kind) {
  const a = APP_PLAN_ACTIONS[String(name || '')];
  if (!a || (kind && a.kind !== kind)) throw new Error('未定義の操作です: ' + name);
  return a;
}

// ---- 読むだけのブック ----

/** 列の文字（A, B, …, AA）を 0 始まりの番号に */
function appColIndex_(letters) {
  let n = 0;
  for (let i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
  return n - 1;
}

/**
 * 旧来の計算が書く数式（OUTPUT の予算の列だけ: =SUM(H5:H16) と =H5+I5）の値。ほかの数式は空として読む。
 * Sheets と同じく、空のセルは 0 として足す。
 */
function appEvalFormula_(f, value) {
  const num = v => (v === '' || v === null || v === undefined ? 0 : v);
  let m = /^=SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)$/.exec(String(f));
  if (m && m[1] === m[3]) {
    const c = appColIndex_(m[1]);
    let s = 0;
    for (let r = Number(m[2]) - 1; r <= Number(m[4]) - 1; r++) { const v = value(r, c); if (typeof v === 'number') s += v; }
    return s;
  }
  m = /^=([A-Z]+)(\d+)\+([A-Z]+)(\d+)$/.exec(String(f));
  if (m) {
    const a = num(value(Number(m[2]) - 1, appColIndex_(m[1])));
    const b = num(value(Number(m[4]) - 1, appColIndex_(m[3])));
    return typeof a === 'number' && typeof b === 'number' ? a + b : '#VALUE!';
  }
  return '';
}

/** 読むだけのもの: 用意していない操作（書き込みなど）を呼ぶと、何をしようとしたかを示して止める */
function appReadOnly_(obj, what) {
  return new Proxy(obj, {
    get(t, k) {
      if (k in t || typeof k === 'symbol') return t[k];
      return () => { throw new Error('画面の表示では使えない操作です（' + what + '.' + String(k) + '）。'); };
    }
  });
}

/** 組み立てたシート 1 枚を、読むだけの Sheet に見せる */
function appViewSheet_(dec, parent) {
  const value = (r, c) => {
    if (r < 0 || c < 0 || r >= dec.lastRow || c >= dec.lastCol) return '';
    const f = dec.formulas[r][c];
    return f ? appEvalFormula_(f, value) : dec.values[r][c];
  };
  const formula = (r, c) => (r < dec.lastRow && c < dec.lastCol ? dec.formulas[r][c] || '' : '');
  const grid = (r0, c0, nr, nc, fn) => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => fn(r0 + i, c0 + j)));
  const display = v => (v === '' || v === null || v === undefined ? '' : appIsDate_(v) ? Utilities.formatDate(v, APP_TZ, 'yyyy/MM/dd') : String(v));
  let sheet = null;
  const range = (r0, c0, nr, nc) => {
    if (!(nr >= 1)) throw new Error('The number of rows in the range must be at least 1.');
    if (!(nc >= 1)) throw new Error('The number of columns in the range must be at least 1.');
    return appReadOnly_({
      getValue: () => value(r0, c0),
      getValues: () => grid(r0, c0, nr, nc, value),
      getDisplayValue: () => display(value(r0, c0)),
      getDisplayValues: () => grid(r0, c0, nr, nc, (r, c) => display(value(r, c))),
      getFormula: () => formula(r0, c0),
      getFormulas: () => grid(r0, c0, nr, nc, formula),
      getRow: () => r0 + 1, getColumn: () => c0 + 1, getNumRows: () => nr, getNumColumns: () => nc,
      getLastRow: () => r0 + nr, getLastColumn: () => c0 + nc,
      getA1Notation: () => appA1_(r0, c0) + (nr > 1 || nc > 1 ? ':' + appA1_(r0 + nr - 1, c0 + nc - 1) : ''),
      getSheet: () => sheet
    }, 'Range');
  };
  const a1 = text => {
    const m = /^([A-Z]+)(\d+)?(?::([A-Z]+)(\d+)?)?$/.exec(String(text).replace(/\$/g, '').toUpperCase());
    if (!m) throw new Error('読めない範囲です: ' + text);
    const r0 = m[2] ? Number(m[2]) - 1 : 0;
    const c0 = appColIndex_(m[1]);
    if (!m[3]) return range(r0, c0, m[2] ? 1 : dec.maxRows, 1);
    const r1 = m[4] ? Number(m[4]) - 1 : dec.maxRows - 1;
    const c1 = appColIndex_(m[3]);
    return range(Math.min(r0, r1), Math.min(c0, c1), Math.abs(r1 - r0) + 1, Math.abs(c1 - c0) + 1);
  };
  sheet = appReadOnly_({
    getName: () => dec.name, getLastRow: () => dec.lastRow, getLastColumn: () => dec.lastCol,
    getMaxRows: () => dec.maxRows, getMaxColumns: () => dec.maxCols,
    getRange: (a, b, c, d) => (typeof a === 'string' ? a1(a) : range(Number(a) - 1, Number(b) - 1, c === undefined ? 1 : Number(c), d === undefined ? 1 : Number(d))),
    getDataRange: () => range(0, 0, Math.max(1, dec.lastRow), Math.max(1, dec.lastCol)),
    getSheetId: () => 0, isSheetHidden: () => false, getFrozenRows: () => 0, getFrozenColumns: () => 0, getParent: () => parent
  }, 'Sheet');
  return sheet;
}

/**
 * データ本体の計画 1 つ分を、読むだけの Spreadsheet に見せる（旧来の画面の読み取りを、計算用ブックなしで動かす）。
 * シートは使われたときに組み立てる。表示形式は読まない。
 */
function appStoreBook_(plan) {
  const planId = plan.plan_id;
  const mine = r => r.plan_id === planId;
  const meta = {};
  appReadTable_('ENG_SHEETS').filter(mine).forEach(r => { meta[r.sheet] = r; });
  const loaded = {};
  let segs = null;
  let book = null;
  const load = name => {
    if (Object.prototype.hasOwnProperty.call(loaded, name)) return loaded[name];
    const m = meta[name];
    if (!m) return (loaded[name] = null);
    if (!segs) {
      segs = {};
      appReadTable_('ENG_ROWS').filter(mine).forEach(r => { (segs[r.sheet] = segs[r.sheet] || []).push(r); });
    }
    const rows = m.mode === 'table' ? appReadTable_('ENG_' + name).filter(mine) : [];
    const dec = appEngDecodeSheet_(Object.assign({}, m, { fmt_columns: '0' }), rows, segs[name] || [], []);
    return (loaded[name] = appViewSheet_(dec, book));
  };
  book = appReadOnly_({
    getSheetByName: n => load(String(n)),
    getSheets: () => Object.keys(APP_ENGINE_SHEETS).filter(n => meta[n]).map(load),
    getId: () => '', getUrl: () => '', getName: () => String(plan.client_label || ''),
    getSpreadsheetTimeZone: () => plan.time_zone || APP_TZ, getSpreadsheetLocale: () => plan.locale || 'ja_JP',
    toast: () => {}
  }, 'Spreadsheet');
  return book;
}

// ---- 見る ----

/** 計画の画面の中身（旧来の Web アプリの webGetBootstrap_ と同じもの）。入力が同じなら覚えておいたものを返す */
function appPlanView_(ctx, planId) {
  const plan = appPlanOf_(planId);
  const clients = {};
  appReadTable_('CLIENTS').forEach(c => { clients[c.client_id] = c.client_name; });
  const inputHash = appPlanInputHash_(plan.plan_id);
  const day = Utilities.formatDate(new Date(), APP_TZ, 'yyyy-MM-dd');   // 旧来の画面は「今日」で変わる（既定の年度など）
  const key = 'VIEW_' + appSha256Hex_([plan.plan_id, inputHash, day, APP_VERSION].join('|')).slice(0, 32);
  const cached = appJobGetResult_(key);
  let view;
  if (cached.found) {
    view = cached.value;
  } else {
    const t0 = new Date().getTime();
    const book = appStoreBook_(plan);
    const call = appLegacyCall_(book, { asOfMs: t0, seed: 'view:' + plan.plan_id, actor: ctx.actor }, 'webGetBootstrap_', []);
    const boot = appSerialize_(call.value);
    delete boot.user; delete boot.bookUrl; delete boot.access;   // 旧ブックの URL・旧来の管理者の判定は出さない
    view = { boot: boot, engine: { version: call.version, sourceSha256: call.sourceSha256, webSha256: call.webSha256 },
      builtMs: new Date().getTime() - t0 };
    // 覚えておけなくても画面は出す（次の表示がまた組み立てになるだけ）
    try { appJobPutResult_(key, view); } catch (e) { Logger.log('画面の中身を覚えておけません: ' + (e && e.message ? e.message : e)); }
    if (view.builtMs > 20000) appRunLog_({ requestId: ctx.requestId, kind: 'PLAN.VIEW', status: 'SLOW', durationMs: view.builtMs, detail: { planId: plan.plan_id } });
  }
  const roles = ctx.roles || [];
  return Object.assign({
    plan: { planId: plan.plan_id, clientName: clients[plan.client_id] || plan.client_label, fy: plan.fy },
    inputHash: inputHash,
    can: {
      plan: appHasRole_(roles, 'PLANNER', plan.client_id),
      approve: appHasRole_(roles, 'APPROVER', plan.client_id),
      admin: appHasRole_(roles, 'ADMIN')
    },
    sourceReady: !!appSettingValue_('source.zac_spreadsheet'),
    actions: Object.keys(APP_PLAN_ACTIONS).map(k => ({ action: k, kind: APP_PLAN_ACTIONS[k].kind, label: APP_PLAN_ACTIONS[k].label,
      minRole: APP_PLAN_ACTIONS[k].minRole })),
    recent: appReadTable_('PLAN_ACTIONS').filter(r => r.plan_id === plan.plan_id).map(appStripRow_)
      .sort((a, b) => (a.finished_at < b.finished_at ? 1 : -1)).slice(0, 10)
      .map(r => ({ action: r.action, label: (APP_PLAN_ACTIONS[r.action] || {}).label || r.action, finishedAt: r.finished_at, actor: r.actor_email,
        changed: appParseJsonList_(r.changed_sheets_json) }))
  }, view);
}

/** 表の *_json の列（文字列で持つ）を一覧に戻す。読めなければ空 */
function appParseJsonList_(v) {
  if (Array.isArray(v)) return v;
  try { const x = JSON.parse(String(v || '[]')); return Array.isArray(x) ? x : []; } catch (e) { return []; }
}

// ---- 保存・実行（裏の処理） ----

/** 新アプリの側で行う保存（旧来に同じ関数が無いもの）。返り値の形は appLegacyCall_ と同じ */
function appPlanLocalCall_(book, name, args) {
  const fns = { appPlanSetPeople_: appPlanSetPeople_ };
  return { value: fns[name].apply(null, [book].concat(args)), version: 'app-' + APP_VERSION, sourceSha256: '', webSha256: '' };
}

/** 担当者（カンマ区切り）を CONFIG!B4 に書く（旧来の saveInitialSetupSettings の担当者の行と同じ） */
function appPlanSetPeople_(book, peopleCsv) {
  const people = String(peopleCsv || '').split(/[,、，]/).map(s => s.trim()).filter(Boolean);
  if (!people.length) throw new Error('担当者を 1 人以上入れてください。');
  if (people.some(p => p.length > 40)) throw new Error('担当者の名前が長すぎます。');
  const cfg = book.getSheetByName('CONFIG');
  if (!cfg) throw new Error('CONFIG がありません。');
  const csv = people.join(',');
  cfg.getRange('B4').setValue(csv);
  return { peopleCsv: csv, people: people };
}

/** 旧来の関数の返り値のうち、記録に残す小さいもの（画面の中身そのものは除く） */
function appPlanResultSummary_(value) {
  if (!value || typeof value !== 'object') return value === undefined ? null : value;
  const out = {};
  Object.keys(value).forEach(k => { if (['boot', 'input', 'output', 'quarterly'].indexOf(k) < 0) out[k] = value[k]; });
  const text = JSON.stringify(appSerialize_(out));
  return text.length > 2000 ? { truncated: true, head: text.slice(0, 2000) } : JSON.parse(text);
}

function appPlanActionRow_(ctx, plan, actionName, x) {
  return {
    action_id: x.actionId, plan_id: plan.plan_id, action: actionName, status: 'DONE',
    engine_version: x.engine.version, engine_sha256: x.engine.sourceSha256, web_sha256: x.engine.webSha256,
    seed: x.seed, as_of: Utilities.formatDate(new Date(x.asOfMs), APP_TZ, "yyyy-MM-dd'T'HH:mm:ssZ"), input_hash: x.inputHash,
    changed_sheets_json: x.changed, result_json: x.result, started_at: x.startedAt, finished_at: appNowIso_(), actor_email: ctx.actor
  };
}

/** 保存: 使うシートだけを組み立てて旧来の webSave* を動かし、変わったシートをデータ本体へ戻す（1 つの処理・ロックの中） */
function appPlanEdit_(ctx, p) {
  const act = appPlanAction_(p.action, 'edit');
  const plan = appPlanOf_(p.planId);
  const args = appJobArgs_(p);
  const t0 = new Date().getTime();
  const actionId = appId_('ACT');
  return appWithLock_(() => {
    const stored = appStoredHashes_(plan.plan_id);
    const inputHash = appPlanInputHash_(plan.plan_id);
    if (p.inputHash && p.inputHash !== inputHash) {
      throw new Error('画面を開いた後に、ほかの人の操作でデータが変わりました。画面を開き直して、もう一度入力してください。');
    }
    const scratch = appParityScratch_(plan);
    const build = appScratchFromStore_(scratch, plan.plan_id, act.sheets);
    const t1 = new Date().getTime();
    const call = act.local ? appPlanLocalCall_(scratch, act.local, act.args(args || {}))
      : appLegacyCall_(scratch, { asOfMs: t0, seed: actionId, actor: ctx.actor }, act.fn, act.args(args || {}));
    const t2 = new Date().getTime();
    const cap = appCaptureChanged_(scratch, plan.plan_id, stored, act.sheets);
    const names = cap.changed.map(e => e.sheetRow.sheet);
    const written = names.length ? appWriteChanged_(ctx, plan.plan_id, cap.changed, actionId) : {};
    const result = appPlanResultSummary_(call.value);
    if (act.local === 'appPlanSetPeople_' && names.length) {
      appUpdateByKey_('PLANS', { plan_id: plan.plan_id }, { people_csv: call.value.peopleCsv }, undefined, ctx.actor);
    }
    appInsertRows_('PLAN_ACTIONS', [appPlanActionRow_(ctx, plan, p.action, { actionId: actionId, engine: call, seed: actionId, asOfMs: t0,
      inputHash: inputHash, changed: names, result: result, startedAt: Utilities.formatDate(new Date(t0), APP_TZ, "yyyy-MM-dd'T'HH:mm:ssZ") })]);
    return { actionId: actionId, planId: plan.plan_id, action: p.action, changed: names, written: written, result: result,
      build: build.filter(x => x.mismatch || x.forcedText || x.formatMismatches).map(x => x.sheet),
      timing: { buildMs: t1 - t0, runMs: t2 - t1, saveMs: new Date().getTime() - t2 },
      audit: { entityId: plan.plan_id, clientId: plan.client_id } };
  });
}

/** 実行（計算）: 全部のシートを組み立てて旧来の webRun* を動かし、変わったシートを控える。保存は続きの処理 */
function appPlanRunCalc_(ctx, p, job) {
  const act = appPlanAction_(p.action, 'run');
  const plan = appPlanOf_(p.planId);
  const t0 = new Date().getTime();
  const actionId = appId_('ACT');
  const inputHash = appPlanInputHash_(plan.plan_id);
  const stored = appStoredHashes_(plan.plan_id);
  return appWithLock_(() => {
    const scratch = appParityScratch_(plan);
    const build = appScratchFromStore_(scratch, plan.plan_id);
    const t1 = new Date().getTime();
    const call = appLegacyCall_(scratch, { asOfMs: t0, seed: actionId, actor: ctx.actor }, act.fn, []);
    const t2 = new Date().getTime();
    const cap = appCaptureChanged_(scratch, plan.plan_id, stored);
    appJobPutResult_(job.id + '_SAVE', { changed: cap.changed });
    const t3 = new Date().getTime();
    return {
      __next: { kind: 'PLAN.RUN_SAVE', payload: {
        planId: plan.plan_id, action: p.action, actionId: actionId, seed: actionId, asOfMs: t0, inputHash: inputHash, parentJobId: job.id,
        engine: { version: call.version, sourceSha256: call.sourceSha256, webSha256: call.webSha256 }, result: appPlanResultSummary_(call.value),
        changed: cap.changed.map(e => e.sheetRow.sheet), unknown: cap.unknown, startedAt: Utilities.formatDate(new Date(t0), APP_TZ, "yyyy-MM-dd'T'HH:mm:ssZ"),
        build: build.filter(x => x.mismatch || x.forcedText || x.formatMismatches).map(x => x.sheet),
        timing: { buildMs: t1 - t0, runMs: t2 - t1, captureMs: t3 - t2 } } },
      audit: { entityId: plan.plan_id, clientId: plan.client_id }
    };
  });
}

/** 実行（保存）: 計算の間にデータ本体が変わっていないことを確かめてから、控えを書き戻す */
function appPlanRunSave_(ctx, p) {
  const plan = appPlanOf_(p.planId);
  const saved = appJobGetResult_(p.parentJobId + '_SAVE');
  if (!saved.found) throw new Error('計算した結果の控えが見つかりません（保存期間が過ぎた）。もう一度実行してください。');
  const changed = saved.value.changed;
  return appWithLock_(() => {
    if (appPlanInputHash_(plan.plan_id) !== p.inputHash) throw new Error('計算している間にデータ本体が変わりました。もう一度実行してください。');
    const names = changed.map(e => e.sheetRow.sheet);
    const written = names.length ? appWriteChanged_(ctx, plan.plan_id, changed, p.actionId) : {};
    appInsertRows_('PLAN_ACTIONS', [appPlanActionRow_(ctx, plan, p.action, Object.assign({}, p, { changed: names }))]);
    return { actionId: p.actionId, planId: plan.plan_id, action: p.action, changed: names, written: written, result: p.result,
      unknown: p.unknown || [], audit: { entityId: p.actionId, clientId: plan.client_id } };
  });
}

// ---- 画面の中身の確認（所有者）: 読むだけのブックと、計算用ブックで、旧来の画面の中身が同じか ----

/** 2 つの値の違う場所（a.b[3].c の形。limit 件まで） */
function appJsonDiff_(a, b, path, out, limit) {
  if (out.length >= limit) return out;
  const ta = Array.isArray(a) ? 'array' : a === null ? 'null' : typeof a;
  const tb = Array.isArray(b) ? 'array' : b === null ? 'null' : typeof b;
  if (ta !== tb) { out.push({ path: path || '(全体)', a: appJsonShort_(a), b: appJsonShort_(b) }); return out; }
  if (ta === 'array') {
    if (a.length !== b.length) out.push({ path: path + '.length', a: a.length, b: b.length });
    for (let i = 0; i < Math.min(a.length, b.length); i++) appJsonDiff_(a[i], b[i], path + '[' + i + ']', out, limit);
  } else if (ta === 'object') {
    Object.keys(Object.assign({}, a, b)).forEach(k => appJsonDiff_(a[k], b[k], path ? path + '.' + k : k, out, limit));
  } else if (a !== b) {
    out.push({ path: path || '(全体)', a: appJsonShort_(a), b: appJsonShort_(b) });
  }
  return out;
}

function appJsonShort_(v) {
  const s = v === undefined ? '(なし)' : JSON.stringify(v);
  return s.length > 120 ? s.slice(0, 120) + '…' : s;
}

/** 計算用ブックに全部のシートを組み立てて旧来の webGetBootstrap_ を動かし、読むだけのブックで動かしたものと比べる（どちらにも書かない） */
function appPlanViewCheck_(ctx, p) {
  const plan = appPlanOf_(p.planId);
  const t0 = new Date().getTime();
  const strip = v => { const x = appSerialize_(v); delete x.user; delete x.bookUrl; delete x.access; return x; };
  return appWithLock_(() => {
    const scratch = appParityScratch_(plan);
    const build = appScratchFromStore_(scratch, plan.plan_id);
    const t1 = new Date().getTime();
    const opts = { asOfMs: t0, seed: 'view:' + plan.plan_id, actor: ctx.actor };
    const real = strip(appLegacyCall_(scratch, opts, 'webGetBootstrap_', []).value);
    const t2 = new Date().getTime();
    const view = strip(appLegacyCall_(appStoreBook_(plan), opts, 'webGetBootstrap_', []).value);
    const t3 = new Date().getTime();
    const diffs = appJsonDiff_(real, view, '', [], 30);
    return { planId: plan.plan_id, same: diffs.length === 0, diffs: diffs,
      build: build.filter(x => x.mismatch || x.forcedText || x.formatMismatches).map(x => x.sheet),
      timing: { buildMs: t1 - t0, realMs: t2 - t1, viewMs: t3 - t2 }, audit: { entityId: plan.plan_id, clientId: plan.client_id } };
  });
}
