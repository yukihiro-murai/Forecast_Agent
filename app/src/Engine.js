/**
 * Engine.js — 予測の計算を動かす土台（段階2）。
 * 旧来の計算（Forecast_Agent.js）は書き換えず、新アプリの中の「計算用ブック」の上で動かす。
 * 計算用ブックは「売上予測アプリ（システム）」フォルダに 1 つだけ置き、実行のたびに中身を入れ替える（所有者だけ・人は触らない）。
 * 乱数は実行ごとの種（seed）で固定し、同じ入力と同じ種なら同じ P10/P50/P90 になるようにする（FORECAST_RUNS に種を残す）。
 */

/** 文字列から 128 ビットの種を作る（cyrb128） */
function appSeedFrom_(text) {
  const str = String(text);
  let h1 = 1779033703, h2 = 3144134277, h3 = 1013904242, h4 = 2773480762;
  for (let i = 0; i < str.length; i++) {
    const k = str.charCodeAt(i);
    h1 = h2 ^ Math.imul(h1 ^ k, 597399067);
    h2 = h3 ^ Math.imul(h2 ^ k, 2869860233);
    h3 = h4 ^ Math.imul(h3 ^ k, 951274213);
    h4 = h1 ^ Math.imul(h4 ^ k, 2716044179);
  }
  h1 = Math.imul(h3 ^ (h1 >>> 18), 597399067);
  h2 = Math.imul(h4 ^ (h2 >>> 22), 2869860233);
  h3 = Math.imul(h1 ^ (h3 >>> 17), 951274213);
  h4 = Math.imul(h2 ^ (h4 >>> 19), 2716044179);
  h1 ^= (h2 ^ h3 ^ h4); h2 ^= h1; h3 ^= h1; h4 ^= h1;
  return [h1 >>> 0, h2 >>> 0, h3 >>> 0, h4 >>> 0];
}

/** 種つきの乱数（sfc32。0 以上 1 未満） */
function appSeededRandom_(seedText) {
  let [a, b, c, d] = appSeedFrom_(seedText);
  const next = () => {
    a >>>= 0; b >>>= 0; c >>>= 0; d >>>= 0;
    let t = (a + b) | 0;
    a = b ^ (b >>> 9);
    b = (c + (c << 3)) | 0;
    c = (c << 21) | (c >>> 11);
    d = (d + 1) | 0;
    t = (t + d) | 0;
    c = (c + t) | 0;
    return (t >>> 0) / 4294967296;
  };
  for (let i = 0; i < 15; i++) next();  // 種の偏りを捨てる
  return next;
}

/** fn の間だけ Math.random を種つきの乱数にする（入れ子にしない。終われば必ず元に戻す） */
function appWithSeededRandom_(seedText, fn) {
  if (APP_RANDOM_PATCHED_) throw new Error('乱数の固定が入れ子になっています。');
  const original = Math.random;
  Math.random = appSeededRandom_(seedText);
  APP_RANDOM_PATCHED_ = true;
  try {
    return fn();
  } finally {
    Math.random = original;
    APP_RANDOM_PATCHED_ = false;
  }
}
let APP_RANDOM_PATCHED_ = false;

/**
 * 計算用ブック（無ければ作る）。人が開いて触る前提のファイルではないので、名前で示す。
 * 計算は appWithLock_ の中で行い、同時に 2 つの計算が同じブックを使わないようにする。
 */
function appScratchBook_() {
  const props = appProps_();
  const id = props.getProperty(APP_PROP.scratchId);
  if (id) {
    try {
      const ss = SpreadsheetApp.openById(id);
      if (!DriveApp.getFileById(id).isTrashed()) return ss;
    } catch (e) {
      Logger.log('計算用ブックを開けないので作り直します: ' + (e && e.message ? e.message : e));
    }
  }
  const folderId = props.getProperty(APP_PROP.folderId);
  if (!folderId) throw new Error('初期設定がまだです。');
  const ss = SpreadsheetApp.create(APP_FILES.scratch);
  DriveApp.getFileById(ss.getId()).moveTo(DriveApp.getFolderById(folderId));
  props.setProperty(APP_PROP.scratchId, ss.getId());
  return ss;
}

/**
 * 計算用ブックを空にする（シートを 1 枚の空のシートだけにする）。
 * 旧来の計算は必要なシートを自分で作るか、呼ぶ側が入れたシートだけを読む。
 */
function appScratchReset_(ss, token) {
  // 空にした時点で、前の処理の中身ではなくなる（その処理の続きは appScratchOwnedBy_ で止まる）
  appProps_().setProperty(APP_PROP.scratchOwner, token || appId_('SCR'));
  const keep = ss.insertSheet('_EMPTY_' + Utilities.getUuid().slice(0, 8));
  ss.getSheets().forEach(s => { if (s.getSheetId() !== keep.getSheetId()) ss.deleteSheet(s); });
  return keep;
}

/** 計算用ブックの中身が、token の処理が組み立てたままか */
function appScratchOwnedBy_(token) {
  return !!token && appProps_().getProperty(APP_PROP.scratchOwner) === token;
}

/** 組み立てを打ち切る目安（1 回の実行の上限 6 分のうち、残りでシート 1 枚と後始末ができる時間） */
const APP_BUILD_BUDGET_MS = 4 * 60 * 1000;

/** その実行で、次のシートの組み立てに進んでよい最後の時刻（テストでは小さくして、何回かに分かれる動きを確かめる） */
function appBuildDeadline_(jobStartMs) {
  return jobStartMs + APP_BUILD_BUDGET_MS;
}

/**
 * データ本体から計算用ブックを組み立てる（何回かの実行に分けてよい）。state は前の回の続き（無ければ空にして初めから）。
 * deadlineMs を過ぎたら、次のシートに進まずに返す（続きは次の実行で）。only を渡すとそのシートだけ。
 * 返り値: { state: { token, done, problems }, complete }。problems は書き方を変えても直らなかったシート（小さく保つ）
 */
function appScratchBuildStep_(scratch, planId, only, state, deadlineMs) {
  let st = state;
  if (!st) {
    st = { token: appId_('SCR'), done: [], problems: [] };
    appScratchReset_(scratch, st.token);
  } else if (!appScratchOwnedBy_(st.token)) {
    throw new Error('組み立てている間に、計算用ブックがほかの処理で使われました。もう一度実行してください。');
  }
  const names = Object.keys(APP_ENGINE_SHEETS).filter(n => (!only || only.indexOf(n) >= 0) && st.done.indexOf(n) < 0);
  const sheets = appEngLoadPlanSheets_(planId, names);
  for (let i = 0; i < names.length; i++) {
    if (i > 0 && new Date().getTime() > deadlineMs) break;   // 1 回に 1 枚は必ず進める（同じところで止まり続けない）
    const dec = sheets[names[i]];
    if (dec) {
      const w = appEngWriteSheet_(scratch, dec, null);
      if (w.mismatches || w.formatMismatches) st.problems.push(names[i] + (w.mismatches ? ' 値 ' + w.mismatches : '') + (w.formatMismatches ? ' 表示形式 ' + w.formatMismatches : ''));
    }
    st.done.push(names[i]);
  }
  const complete = Object.keys(APP_ENGINE_SHEETS).every(n => (only && only.indexOf(n) < 0) || st.done.indexOf(n) >= 0);
  if (complete) {
    const all = scratch.getSheets();
    all.forEach(s => { if (/^_EMPTY_/.test(s.getName()) && scratch.getSheets().length > 1) scratch.deleteSheet(s); });
  }
  return { state: st, complete: complete };
}

// ---- 旧来の計算に渡す差し替え（LegacyEngine.js の appLegacyEngine_ に渡す） ----

/** 「今」を asOfMs に固定した Date（引数なしの new Date() と Date.now() だけを固定し、ほかは本物と同じ。instanceof も本物の日付で通る） */
function appFrozenDate_(asOfMs) {
  const RealDate = Date;
  return new Proxy(RealDate, {
    construct(target, args, newTarget) {
      return Reflect.construct(target, args.length ? args : [asOfMs], newTarget);
    },
    apply() {
      return new RealDate(asOfMs).toString();
    },
    get(target, prop) {
      if (prop === 'now') return () => asOfMs;
      return Reflect.get(target, prop, target);
    }
  });
}

/** 種から決まる UUID（旧来の計算が作る ID を、同じ種なら同じにする。Math.random は使わない＝乱数の並びを変えない） */
function appSeededUuidMaker_(seedText) {
  let n = 0;
  return () => {
    const h = appSha256Hex_(String(seedText) + ':uuid:' + (n++));
    return h.slice(0, 8) + '-' + h.slice(8, 12) + '-4' + h.slice(13, 16) + '-a' + h.slice(17, 20) + '-' + h.slice(20, 32);
  };
}

/** 呼ぶと必ず止まるサービス（旧来の計算からの外部への通信や画面表示を止める） */
function appBlockedService_(name) {
  return new Proxy({}, { get: (t, prop) => () => { throw new Error('この実行では ' + name + '.' + String(prop) + ' を使いません。'); } });
}

/**
 * 旧来の計算に渡すサービス一式。book = 計算用ブック（旧来の「開いているスプレッドシート」の代わり）。
 * opts: { asOfMs, seed }
 */
function appLegacyServices_(book, opts) {
  const realSA = SpreadsheetApp;
  const realUtil = Utilities;
  const uuid = appSeededUuidMaker_(opts.seed);
  const pass = (real, overrides) => new Proxy({}, {
    get(t, prop) {
      if (Object.prototype.hasOwnProperty.call(overrides, prop)) return overrides[prop];
      const v = real[prop];
      return typeof v === 'function' ? (...a) => real[prop](...a) : v;
    }
  });
  // 裏の処理（トリガー）には画面がないので、トースト（右下の通知）は Google に止められる
  // （「Cannot call SpreadsheetApp.showNotification() from this context」2026-10-01）。何もしないにする。
  // 表示するシートの切り替えは計算の結果に関係しないので、失敗しても計算を止めない
  const bookView = pass(book, {
    toast: () => {},
    setActiveSheet: sh => { try { return book.setActiveSheet(sh); } catch (e) { return sh; } },
    moveActiveSheet: i => { try { return book.moveActiveSheet(i); } catch (e) { return null; } }
  });
  const props = {
    // 売上・実績の取り込み（A-2・B-1）の元は、設定の「ZAC の実績のスプレッドシート」。未設定のまま旧来の既定に頼らない
    getProperty: k => {
      if (k !== 'FORECAST_SOURCE_SPREADSHEET_ID') return null;
      const id = appSettingValue_('source.zac_spreadsheet');
      if (!id) throw new Error('「設定」の「ZAC の実績のスプレッドシート」を入れてください（売上・実績の取り込みの元）。');
      return id;
    },
    getProperties: () => ({}),
    getKeys: () => [],
    setProperty: () => { throw new Error('旧来の計算からは Script Properties に書きません。'); },
    deleteProperty: () => { throw new Error('旧来の計算からは Script Properties を消しません。'); }
  };
  return {
    SpreadsheetApp: pass(realSA, {
      getActiveSpreadsheet: () => bookView,
      getActive: () => bookView,
      getUi: () => { throw new Error('この実行では画面（SpreadsheetApp.getUi）を使えません。'); }
    }),
    Date: appFrozenDate_(opts.asOfMs),
    // 待ち（トーストの間隔）は結果に関係しないので待たない。ID は種から決める
    Utilities: pass(realUtil, { getUuid: uuid, sleep: () => {} }),
    PropertiesService: {
      getScriptProperties: () => props,
      getUserProperties: () => { throw new Error('旧来の計算からは User Properties を使いません。'); },
      getDocumentProperties: () => { throw new Error('旧来の計算からは Document Properties を使いません。'); }
    },
    // 外への問い合わせは止める。A-4 AI 調査だけ、Vertex AI に限って問い合わせる道具（Ai.js の appAiFetcher_）を渡す
    UrlFetchApp: opts.fetch ? { fetch: opts.fetch } : appBlockedService_('UrlFetchApp'),
    HtmlService: appBlockedService_('HtmlService'),
    // 裏の処理（トリガー）では操作した人のメールが取れないので、頼んだ人を「操作した人」として見せる（PROCESS_STATUS の実行者など）
    Session: pass(Session, { getActiveUser: () => ({ getEmail: () => String(opts.actor || '') }) })
  };
}

/**
 * 計算用ブックの上で旧来の予測（A-9 runPhase1Forecast）を動かす。種と「今」を固定する。
 * opts: { asOfMs, seed, confirms: ['extreme', ...] }。返り値: { ok, needConfirm?, version, sourceSha256 }
 */
function appRunLegacyForecast_(book, opts) {
  const svc = appLegacyServices_(book, opts);
  return appWithSeededRandom_(opts.seed, () => {
    const eng = appLegacyEngine_(svc);
    Object.keys(eng.WEB_UI_CONFIRMS_).forEach(k => { delete eng.WEB_UI_CONFIRMS_[k]; });
    (opts.confirms || []).forEach(k => { eng.WEB_UI_CONFIRMS_[String(k)] = true; });
    try {
      eng.runPhase1Forecast();
      return { ok: true, version: eng.VERSION, sourceSha256: eng.SOURCE_SHA256 };
    } catch (e) {
      if (e && e.webConfirm) return { ok: false, needConfirm: e.webConfirm, version: eng.VERSION, sourceSha256: eng.SOURCE_SHA256 };
      throw e;
    }
  });
}

/**
 * 旧来の関数（計算・Web アプリの保存や実行）を 1 つ、差し替えのもとで呼ぶ。種と「今」を固定する。
 * opts: { asOfMs, seed, actor }。返り値: { value, version, sourceSha256, webSha256 }
 */
/** 計算用ブックの上で、旧来の A-1 初期セットアップ（setupForecastBook の中身。画面の確認とダイアログは除く）と設定の保存を動かす */
function appLegacySetupBook_(book, opts, orderKeys, clientName, fy, peopleCsv) {
  const svc = appLegacyServices_(book, opts);
  return appWithSeededRandom_(opts.seed, () => {
    const eng = appLegacyEngine_(svc);
    const ss = svc.SpreadsheetApp.getActiveSpreadsheet();
    const order = orderKeys.map(k => eng.SHEETS[k]);
    eng.resetWorkbookSheets_(ss, order);
    eng.clearAllNotesOnSheets_(ss, order);
    ['buildGUIDE_', 'buildCONFIG_', 'buildSALES_', 'buildFACTORS_PRODUCT_', 'buildFACTORS_CLIENT_', 'buildOPINIONS_', 'buildDEV_',
      'buildPhase1Sheets_', 'buildOUTPUT_', 'normalizeAllSheetNotes_', 'validateNotesIntegrity_', 'applyDefaultAlignmentForAllSheets_',
      'clearAllTabColors_', 'hideNonUserSheets_'].forEach(fn => eng[fn]());
    eng.saveInitialSetupSettings(clientName, String(fy), peopleCsv);
    return { version: eng.VERSION, sourceSha256: eng.SOURCE_SHA256 };
  });
}

function appLegacyCall_(book, opts, fnName, args) {
  const svc = appLegacyServices_(book, opts);
  return appWithSeededRandom_(opts.seed, () => {
    const eng = appLegacyEngine_(svc);
    if (typeof eng[fnName] !== 'function') throw new Error('旧来の関数がありません: ' + fnName);
    const value = eng[fnName].apply(null, args || []);
    return { value: value, version: eng.VERSION, sourceSha256: eng.SOURCE_SHA256, webSha256: eng.WEB_SOURCE_SHA256 };
  });
}

// ---- 計算用ブックの準備と、計算後の中身の控え ----

/** 計算用ブック（計画の地域・時差に合わせる）。予測・実行・計画の作成で使う */
function appWorkScratch_(plan) {
  const scratch = appScratchBook_();
  try { scratch.setSpreadsheetTimeZone(plan.time_zone || APP_TZ); if (plan.locale) scratch.setSpreadsheetLocale(plan.locale); } catch (e) {
    Logger.log('計算用ブックの地域設定: ' + (e && e.message ? e.message : e));
  }
  return scratch;
}

/** データ本体の計画 1 つ分から、計算用ブックを組み立てる（予測・実行）。only を渡すとそのシートだけ */
function appScratchFromStore_(scratch, planId, only) {
  const sheets = appEngLoadPlanSheets_(planId, only);
  const placeholder = appScratchReset_(scratch);
  const report = [];
  Object.keys(APP_ENGINE_SHEETS).forEach(name => {
    const dec = sheets[name];
    if (!dec) return;
    const w = appEngWriteSheet_(scratch, dec, null);
    report.push({ sheet: name, mismatch: w.mismatches, repaired: w.repaired, forcedText: w.forcedText, samples: w.samples,
      formatMismatches: w.formatMismatches, formatFixed: w.formatFixed, formatSamples: w.formatSamples, blankMethod: w.blankMethod });
  });
  if (scratch.getSheets().length > 1) scratch.deleteSheet(placeholder);
  return report;
}

/**
 * 文字列の速いハッシュ（cyrb53。53 ビットを 14 桁の 16 進で）。行ごとの控えに使う（暗号用ではない）。
 * 行ごとに Utilities.computeDigest を呼ぶと、1 万行で 1 分以上かかったため（2026-10-01）。
 */
function appFastHash_(str) {
  let h1 = 0xdeadbeef ^ 0;
  let h2 = 0x41c6ce57 ^ 0;
  for (let i = 0; i < str.length; i++) {
    const ch = str.charCodeAt(i);
    h1 = Math.imul(h1 ^ ch, 2654435761);
    h2 = Math.imul(h2 ^ ch, 1597334677);
  }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
  h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
  const n = 4294967296 * (2097151 & h2) + (h1 >>> 0);
  return ('00000000000000' + n.toString(16)).slice(-14);
}

/** 予測の主な結果（OUTPUT の年度合計と月ごとの P10/P50/P90。旧来の行の位置のとおり） */
function appForecastHeadline_(book) {
  const sh = book.getSheetByName('OUTPUT');
  if (!sh || sh.getLastRow() < 40) return null;
  const num = v => (typeof v === 'number' && isFinite(v) ? v : null);
  const annual = sh.getRange(26, 2, 1, 3).getValues()[0].map(num);
  const objective = sh.getLastRow() >= 65 ? sh.getRange(65, 2, 1, 3).getValues()[0].map(num) : [null, null, null];
  const monthly = sh.getRange(29, 1, 12, 4).getValues().map(r => ({ month: appIsDate_(r[0]) ? Utilities.formatDate(r[0], APP_TZ, 'yyyy/MM') : String(r[0]),
    p10: num(r[1]), p50: num(r[2]), p90: num(r[3]) }));
  const objectiveMonthly = sh.getLastRow() >= 79 ? sh.getRange(68, 1, 12, 4).getValues().map(r => ({
    month: appIsDate_(r[0]) ? Utilities.formatDate(r[0], APP_TZ, 'yyyy/MM') : String(r[0]), p10: num(r[1]), p50: num(r[2]), p90: num(r[3]) })) : [];
  return { title: String(sh.getRange(1, 1).getValue() || ''), annual: { p10: annual[0], p50: annual[1], p90: annual[2] },
    objective: { p10: objective[0], p50: objective[1], p90: objective[2] }, monthly: monthly, objectiveMonthly: objectiveMonthly };
}
