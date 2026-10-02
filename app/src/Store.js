/**
 * Store.js — 保存先の差し替え層（スプレッドシート版）。業務のコードはここ以外で SpreadsheetApp を直接触らない。
 * すべてのセルを書式なしテキスト（@）で保存し、読むときに Schema の型へ戻す（日付や数値への自動変換を防ぐ）。
 * 将来ログなどを BigQuery へ移すときは、この層に実装を足す（業務のコードは変えない）。
 */
let APP_STORE_CACHE_ = {};
let APP_LOCK_DEPTH_ = 0;

function appProps_() {
  return PropertiesService.getScriptProperties();
}

function appNowIso_() {
  return Utilities.formatDate(new Date(), APP_TZ, "yyyy-MM-dd'T'HH:mm:ssZ");
}

function appToday_() {
  return Utilities.formatDate(new Date(), APP_TZ, 'yyyy-MM-dd');
}

/**
 * 年度（4月〜翌3月）。始まりの年で呼ぶ: 2026-10 → FY2026（旧来の計算の getForecastFYStart_・vNext の vNextFiscalYearForDate_ と同じ）。
 * 2026-10-01 までは終わりの年で呼んでいた（ログのファイル名。appEnsureFyNaming_ が一度だけ直す）。
 */
function appFy_(d) {
  return d.getMonth() >= 3 ? d.getFullYear() : d.getFullYear() - 1;
}

/** 時刻順に並ぶ ID（例: RL-20261001183000-1A2B3C4D） */
function appId_(prefix) {
  return prefix + '-' + Utilities.formatDate(new Date(), APP_TZ, 'yyyyMMddHHmmss') + '-' +
    Utilities.getUuid().replace(/-/g, '').slice(0, 8).toUpperCase();
}

function appSha256Hex_(text) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text), Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + (b < 0 ? b + 256 : b).toString(16)).slice(-2)).join('');
}

/** 長い JSON はハッシュと先頭だけにする（セルの上限 5 万字に対し 4 万字まで） */
function appJson_(v) {
  if (v === undefined || v === null || v === '') return '';
  const s = typeof v === 'string' ? v : JSON.stringify(v);
  if (s.length <= APP_JSON_MAX) return s;
  return JSON.stringify({ truncated: true, length: s.length, sha256: appSha256Hex_(s), head: s.slice(0, 2000) });
}

/** スクリプト全体のロック。入れ子で呼ばれても 1 回だけ取る */
function appWithLock_(fn) {
  if (APP_LOCK_DEPTH_ > 0) {
    APP_LOCK_DEPTH_++;
    try { return fn(); } finally { APP_LOCK_DEPTH_--; }
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(APP_LOCK_WAIT_MS);
  APP_LOCK_DEPTH_ = 1;
  try {
    return fn();
  } finally {
    APP_LOCK_DEPTH_ = 0;
    lock.releaseLock();
  }
}

function appIsSetUp_() {
  return !!appProps_().getProperty(APP_PROP.dataId);
}

function appDataSpreadsheet_() {
  if (APP_STORE_CACHE_.data) return APP_STORE_CACHE_.data;
  const id = appProps_().getProperty(APP_PROP.dataId);
  if (!id) throw new Error('データ本体がまだありません。管理者が「初期設定」を実行してください。');
  APP_STORE_CACHE_.data = SpreadsheetApp.openById(id);
  return APP_STORE_CACHE_.data;
}

/** 表のシート。create=true なら無いときに作る。列が定義と違えば止める（fail-closed） */
function appTableSheet_(name, create) {
  const def = APP_TABLES[name];
  if (!def) throw new Error('未定義の表: ' + name);
  const ss = appDataSpreadsheet_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    if (!create) throw new Error('表がありません（' + name + '）。管理者が「初期設定」を実行してください。');
    sh = ss.insertSheet(name);
    appFitColumns_(sh, def.columns.length);
    sh.getRange(1, 1, 1, def.columns.length).setNumberFormat('@').setValues([def.columns]);
    sh.setFrozenRows(1);
    return sh;
  }
  const head = sh.getRange(1, 1, 1, def.columns.length).getValues()[0].map(String);
  if (head.join('|') !== def.columns.join('|')) throw new Error('表の列が定義と違います（' + name + '）。書き込みを止めました。');
  return sh;
}

/** シートの列数をちょうど n にする（足りなければ足し、余りは消してセルの上限を節約する） */
function appFitColumns_(sh, n) {
  const max = sh.getMaxColumns();
  if (max < n) sh.insertColumnsAfter(max, n - max);
  else if (max > n) sh.deleteColumns(n + 1, max - n);
}

/** 行を 2 行目から書く（大きいときは分けて書く）。足りない行は足す */
function appWriteBody_(sh, rows, width) {
  if (!rows.length) return;
  const need = rows.length + 1;
  if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  const CHUNK = 5000;
  for (let i = 0; i < rows.length; i += CHUNK) {
    const part = rows.slice(i, i + CHUNK);
    sh.getRange(2 + i, 1, part.length, width).setNumberFormat('@').setValues(part);
  }
}

function appRowToObject_(def, r) {
  const o = {};
  def.columns.forEach((c, j) => {
    const v = r[j] === null || r[j] === undefined ? '' : String(r[j]);
    const t = appColumnType_(c, def);
    o[c] = t === 'bool' ? v === 'TRUE' : t === 'int' ? (v === '' ? 0 : Number(v)) : t === 'num' ? (v === '' ? null : Number(v)) : v;
  });
  return o;
}

function appObjectToRow_(def, o) {
  return def.columns.map(c => {
    const v = o[c];
    const t = appColumnType_(c, def);
    if (v === undefined || v === null) return t === 'bool' ? 'FALSE' : t === 'int' ? '0' : '';
    if (t === 'num' && (typeof v !== 'number' || !isFinite(v))) throw new Error('数値の列に数値でない値: ' + c);
    if (t === 'bool') return v ? 'TRUE' : 'FALSE';
    if (t === 'json' && typeof v !== 'string') return appJson_(v);
    return String(v);
  });
}

/** 表の全行（_row = シートの行番号つき）。キーが空の行は飛ばす */
function appReadTable_(name) {
  const def = APP_TABLES[name];
  const sh = appTableSheet_(name, false);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, def.columns.length).getValues()
    .map((r, i) => { const o = appRowToObject_(def, r); o._row = i + 2; return o; })
    .filter(o => def.key.some(k => o[k] !== ''));
}

function appInsertRows_(name, objs) {
  const def = APP_TABLES[name];
  if (!objs.length) return [];
  // キーの重複を入れない（同じ秒に作った ID がぶつかった場合なども止める）
  const keyOf = o => def.key.map(k => String(o[k] === undefined || o[k] === null ? '' : o[k])).join('\u0001');
  const seen = {};
  appReadTable_(name).forEach(o => { seen[keyOf(o)] = true; });
  objs.forEach(o => {
    const k = keyOf(o);
    if (def.key.every(c => o[c] === undefined || o[c] === null || o[c] === '')) throw new Error('キーが空の行は足せません（' + name + '）。');
    if (seen[k]) throw new Error('同じキーの行がすでにあります（' + name + '）。');
    seen[k] = true;
  });
  const sh = appTableSheet_(name, false);
  const rows = objs.map(o => appObjectToRow_(def, o));
  const start = sh.getLastRow() + 1;
  if (start + rows.length - 1 > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), Math.max(1000, rows.length));
  sh.getRange(start, 1, rows.length, def.columns.length).setNumberFormat('@').setValues(rows);
  return objs;
}

/**
 * 表の中身を入れ替える。keep(row) が true の行は残し、その後ろに objs を足す（計画 1 つ分の入れ替えなどに使う）。
 * 新しい中身を先に書いてから、余った古い行を消す。返り値は { removed, kept, added }
 */
function appReplaceRows_(name, keep, objs) {
  const cur = appReadTable_(name);
  const kept = cur.filter(keep).map(appStripRow_);
  appReplaceWhole_(name, kept.concat(objs));
  return { removed: cur.length - kept.length, kept: kept.length, added: objs.length };
}

/** 表全体を rows にする（2 行目から書き、余った古い行を消す）。何度書いても同じ結果になる（保存の控えの書き直しに使う） */
function appReplaceWhole_(name, rows) {
  const def = APP_TABLES[name];
  const sh = appTableSheet_(name, false);
  const keyOf = o => def.key.map(k => String(o[k] === undefined || o[k] === null ? '' : o[k])).join('\u0001');
  const seen = {};
  rows.forEach(o => {
    const k = keyOf(o);
    if (seen[k]) throw new Error('同じキーの行があります（' + name + '）。');
    seen[k] = true;
  });
  const oldLast = sh.getLastRow();
  appWriteBody_(sh, rows.map(o => appObjectToRow_(def, o)), def.columns.length);
  const newLast = rows.length + 1;
  if (oldLast > newLast) sh.getRange(newLast + 1, 1, oldLast - newLast, def.columns.length).clearContent();
  return { rows: rows.length };
}

/** キーの無い行だけ足す（何度呼んでも同じ結果になる）。返り値は足した行 */
function appEnsureRows_(name, objs) {
  const def = APP_TABLES[name];
  if (!objs.length) return [];
  const keyOf = o => def.key.map(k => String(o[k] === undefined || o[k] === null ? '' : o[k])).join('\u0001');
  const have = {};
  appReadTable_(name).forEach(o => { have[keyOf(o)] = true; });
  return appInsertRows_(name, objs.filter(o => !have[keyOf(o)]));
}

/** キーの行の列を patch の値にする。すでに同じ値なら書かない（何度呼んでも同じ結果になる）。返り値は書いたかどうか */
function appPatchRow_(name, keyVals, patch, actor) {
  const def = APP_TABLES[name];
  const cur = appReadTable_(name).filter(r => def.key.every(k => r[k] === keyVals[k]))[0];
  if (!cur) throw new Error('対象が見つかりません（' + name + '）。');
  if (Object.keys(patch).every(k => String(cur[k]) === String(patch[k]))) return false;
  appUpdateByKey_(name, keyVals, patch, undefined, actor || '');
  return true;
}

/**
 * 計画 1 つ分の行を入れ替える。前の行がそのまま残り、後ろに行が足されただけなら、足された行だけを書く（履歴の表を毎回書き直さない）。
 * 返り値: { appended } か { replaced: true, removed, kept, added }
 */
function appReplaceOrAppend_(name, planId, objs) {
  const def = APP_TABLES[name];
  const key = o => def.columns.map(c => String(o[c] === undefined || o[c] === null ? '' : o[c])).join('\u0001');
  const cur = appReadTable_(name).filter(r => r.plan_id === planId).map(appStripRow_);
  const next = {};
  objs.forEach(o => { next[o.seq] = o; });
  if (cur.every(o => next[o.seq] && key(next[o.seq]) === key(o))) {
    const have = {};
    cur.forEach(o => { have[o.seq] = true; });
    const add = objs.filter(o => !have[o.seq]);
    appInsertRows_(name, add);
    return { appended: add.length };
  }
  return Object.assign({ replaced: true }, appReplaceRows_(name, r => r.plan_id !== planId, objs));
}

function appStripRow_(o) {
  const c = Object.assign({}, o);
  delete c._row;
  return c;
}

/**
 * キーで 1 行を更新する。expectedVersion を渡すと、読み込んだ後にほかの人が更新していたら止める（楽観ロック）。
 * 返り値は { before, after }（監査の変更前後に使う）
 */
function appUpdateByKey_(name, keyVals, patch, expectedVersion, actor) {
  const def = APP_TABLES[name];
  const cur = appReadTable_(name).filter(r => def.key.every(k => r[k] === keyVals[k]))[0];
  if (!cur) throw new Error('対象が見つかりません（' + name + '）。');
  if (expectedVersion !== undefined && expectedVersion !== null && expectedVersion !== '' &&
      Number(expectedVersion) !== Number(cur.row_version)) {
    throw new Error('ほかの人が先に更新しました。画面を読み直してから、もう一度保存してください。');
  }
  const after = Object.assign(appStripRow_(cur), patch, {
    updated_at: appNowIso_(), updated_by: actor, row_version: Number(cur.row_version || 0) + 1
  });
  const sh = appTableSheet_(name, false);
  sh.getRange(cur._row, 1, 1, def.columns.length).setNumberFormat('@').setValues([appObjectToRow_(def, after)]);
  return { before: appStripRow_(cur), after: after };
}
