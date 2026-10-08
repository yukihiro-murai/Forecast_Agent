/**
 * V10.js — 表の版 10 の共通の道具（SCHEMA_PLAN_v10-12_JA.md の 3 章。2026-10-08 村井さん承認「版10はおすすめで」）。
 * - 記録の表（Schema.js の appendOnly）へ足すのは、appLogOps_ で作る控えの書き方 append だけ。本体の保存と同じ控え（appJournalRun_）に入れて一緒に書く。
 *   行は消さず書き換えない（直すときは新しい行を足す）。記録は計算に使わない。新しい自動の処理とメールは足さない（今ある操作の中で書く）。
 * - 見せる範囲はサーバーの応答で絞る（appLogViewer_・appLogVisible_）。人の名前は予算策定担当以上だけ（2026-10-07）。
 * - 一度だけの写し（APP_V10_BACKFILLS）: 表の版がそろった後の操作の初めに、名前の順に動かす（Setup.js の appRunBackfills_）。
 *   中身は、それぞれの記録を作るときに書く（今は仮のもの。{ skipped: true } を返すと、記録せずに次の操作でまた呼ぶ）。
 */

/** 記録の番号の頭（appNewLogId_・appStableLogId_ に渡す） */
const APP_LOG_PREFIX = { INPUT_LOG: 'INL', AI_RESEARCH_LOG: 'AIR', LAYER_EFFECTS: 'LEF', HIT_RECORDS: 'HIT', LEARNING_LOG: 'LRN',
  PERSON_LINKS: 'PLK', BACKTEST: 'BTP' };

/** 測る専用の計画（PLANS.purpose。3-9）と、その数の上限（9 章 9: まず 5 計画） */
const APP_PLAN_MEASURE = 'MEASURE';
const APP_MEASURE_PLAN_MAX = 5;

/** 版 10 の一度だけの写し（名前だけ。関数は動かすときに探す）。それぞれの中身は、その記録を作るときに書く */
const APP_V10_BACKFILLS = ['appV10BackfillInputLog_', 'appV10BackfillAiResearchLog_', 'appV10BackfillHitRecords_'];

// ---- 記録を足す ----

/**
 * 記録の表に行を足す控えの書き方（ops の配列。本体の ops に concat して appJournalRun_ に渡す）。行が無ければ [] を返す。
 * 行はここで確かめて形をそろえる（表に無い列・空のキー・plan_id の無い行は止める。数の列は数か空にし、数でなければ止める。
 * 4 万字を超えるセルはハッシュと先頭だけにしてエラーのログに残す）。止めるのは控えを置く前なので、本体も書かない。
 * キーは控えを作るときに決める（appNewLogId_・appStableLogId_）。書くときに作らない（控えの書き直しで同じ行になるように）。
 * 表の版がまだそろっていない（移行の前・途中）ときは、記録を足さずに [] を返す（本体の保存は止めない。エラーのログに残す）
 */
function appLogOps_(table, rows) {
  const def = APP_TABLES[table];
  if (!def || !def.appendOnly) throw new Error('追記だけの記録の表ではありません: ' + table);
  const list = (rows || []).filter(r => r);
  if (!list.length) return [];
  const out = list.map(r => appLogRow_(table, def, r));
  if (!appLogReady_()) {
    appLogError_('LOG.SKIPPED', new Error('表の版がそろう前なので、記録を足しませんでした（' + table + '・' + out.length + ' 行）。'), null);
    return [];
  }
  return [{ table: table, mode: 'append', rows: out }];
}

/** 記録の表を書ける版か（表の版がそろった印。appEnsureTables_ が付ける） */
function appLogReady_() {
  return appProps_().getProperty(APP_PROP.tablesVersion) === String(APP_SCHEMA_VERSION);
}

/** 記録の 1 行を確かめ、控えに置く形にする（列の型に合わせる） */
function appLogRow_(table, def, r) {
  const extra = Object.keys(r).filter(k => def.columns.indexOf(k) < 0);
  if (extra.length) throw new Error('表に無い列です（' + table + '）: ' + extra.join(', '));
  const empty = v => v === undefined || v === null || v === '';
  if (def.key.some(k => empty(r[k]))) throw new Error('キーが空の行は足せません（' + table + '）。');
  if (def.columns[0] === 'plan_id' && !def.globalRows && empty(r.plan_id)) throw new Error('計画（plan_id）が空の行は足せません（' + table + '）。');
  const o = {};
  Object.keys(r).forEach(c => {
    let v = r[c];
    if (v === undefined) return;
    const t = appColumnType_(c, def);
    if (t === 'num' || t === 'int') {
      if (v === null || v === '') v = null;
      else {
        const n = typeof v === 'number' ? v : Number(v);
        if (typeof v === 'boolean' || !isFinite(n)) throw new Error('数の列に数でない値があります（' + table + ' の ' + c + '）。');
        v = t === 'int' ? Math.round(n) : n;
      }
    } else if (t === 'bool') v = v === true || v === 'TRUE' || v === 1 || v === '1';
    else if (t === 'json') v = v === null || typeof v === 'string' ? v : JSON.stringify(v);
    else if (v !== null) v = String(v);
    if (typeof v === 'string' && v.length > APP_JSON_MAX) {
      appLogError_('LOG.TRUNCATED', new Error('記録のセルが長すぎるので、ハッシュと先頭だけにしました（' + table + ' の ' + c + '・' + v.length + ' 字）。'), null);
      v = appJson_(v);
    }
    o[c] = v;
  });
  return o;
}

/** 記録の番号（時刻順。例: INL-20261008153000-1A2B3C4D5E6F）。控えを作るときに決める */
function appNewLogId_(prefix) {
  return String(prefix) + '-' + Utilities.formatDate(new Date(), APP_TZ, 'yyyyMMddHHmmss') + '-' +
    Utilities.getUuid().replace(/-/g, '').slice(0, 12).toUpperCase();
}

/** 中身から決まる記録の番号（同じ中身なら同じ番号。例: 予測の回と月）。2 回動かしても二重に足さないための印 */
function appStableLogId_(prefix, parts) {
  return String(prefix) + '-' + appSha256Hex_((parts || []).map(p => (p === undefined || p === null ? '' : String(p))).join('\u0001')).slice(0, 24).toUpperCase();
}

/** 表の中の、切り詰めた（ハッシュと先頭だけにした）セルの数（テストで、記録が 4 万字に収まっていることを確かめる） */
function appLogTruncatedCells_(table) {
  let n = 0;
  appRawTable_(table).forEach(r => r.forEach(v => { if (String(v).indexOf('{"truncated":true') === 0) n++; }));
  return n;
}

// ---- 見せる範囲（3-1・3-4: 入力の記録と当たりの記録。2026-10-08 承認）----

/**
 * 記録を見る人: full = 予算策定担当以上（そのメーカー単位の担当を含む。clientId で数える）は全部の行と名前。
 * ほかの人（閲覧・情報提供）は、つなぎ（PERSON_LINKS）で本人と分かる自分の行だけ。ほかの行は種類ごとの件数だけ（名前は送らない）。
 * ymd = つなぎが効く日（省くと今日）
 */
function appLogViewer_(ctx, clientId, ymd) {
  const email = String((ctx && ctx.user && ctx.user.email) || '').trim().toLowerCase();
  const full = appHasRole_(ctx && ctx.roles, 'PLANNER', clientId || undefined);
  return { full: full, email: email, clientId: String(clientId || ''), names: full ? [] : appPersonNamesOf_(email, clientId, ymd) };
}

/**
 * 記録の行を、見る人に合わせて分ける。返り値 { rows: 送る行, counts: 送らない行の種類ごとの件数, hidden: 送らない行の数 }。
 * opts: personOf(r) = 行の担当者の名前（入力の記録は r.person）、emailOf(r) = 行の人のメール（当たりの記録は r.person_email）、
 * typeOf(r) = 件数を数える種類（例: r.kind・r.source_kind）。送る行の中の名前は、そのまま（本人の名前だけ）
 */
function appLogVisible_(viewer, rows, opts) {
  const o = opts || {};
  const list = rows || [];
  if (viewer && viewer.full) return { rows: list.slice(), counts: {}, hidden: 0 };
  const names = {};
  ((viewer && viewer.names) || []).forEach(n => { names[String(n).trim()] = true; });
  const email = viewer && viewer.email ? viewer.email : '';
  const own = r => {
    const e = o.emailOf ? String(o.emailOf(r) || '').trim().toLowerCase() : '';
    if (email && e && e === email) return true;
    const p = o.personOf ? String(o.personOf(r) || '').trim() : '';
    return !!p && !!names[p];
  };
  const out = [], counts = {};
  list.forEach(r => {
    if (own(r)) { out.push(r); return; }
    const k = o.typeOf ? String(o.typeOf(r) || '') : '';
    counts[k] = (counts[k] || 0) + 1;
  });
  return { rows: out, counts: counts, hidden: list.length - out.length };
}

/** 人のつなぎの行（読めなければ空。つなぎが無いときは、本人の行も無いことにする） */
function appPersonLinks_() {
  try { return appReadTable_('PERSON_LINKS'); } catch (e) { return []; }
}

/** つなぎがその日に効くか（有効・期間の中） */
function appPersonLinkLive_(l, ymd) {
  return !!l.is_active && (!l.valid_from || l.valid_from <= ymd) && (!l.valid_to || ymd <= l.valid_to);
}

/**
 * 担当者の名前（入力の「担当者」の欄のまま）がつながるメンバーのメール。無い・決められないときは ''。
 * そのメーカーのつなぎを先に見て、無ければ全部のメーカーのつなぎ（client_id が空）を見る。同じ名前が 2 人につながるときは決めない
 */
function appPersonEmailOf_(name, clientId, ymd, links) {
  const day = ymd || appToday_();
  const n = String(name || '').trim();
  if (!n) return '';
  const live = (links || appPersonLinks_()).filter(l => String(l.person_name || '').trim() === n && appPersonLinkLive_(l, day));
  const own = clientId ? live.filter(l => l.client_id === clientId) : [];
  const pick = own.length ? own : live.filter(l => !l.client_id);
  const emails = pick.map(l => String(l.email || '').trim().toLowerCase()).filter((v, i, a) => v && a.indexOf(v) === i);
  return emails.length === 1 ? emails[0] : '';
}

/** メンバーのメールにつながる担当者の名前（そのメーカーで、その日に効くもの） */
function appPersonNamesOf_(email, clientId, ymd) {
  const e = String(email || '').trim().toLowerCase();
  if (!e) return [];
  const day = ymd || appToday_();
  const links = appPersonLinks_();
  const names = {};
  links.forEach(l => { if (appPersonLinkLive_(l, day) && String(l.email || '').trim().toLowerCase() === e) names[String(l.person_name || '').trim()] = true; });
  return Object.keys(names).filter(n => n && appPersonEmailOf_(n, clientId, day, links) === e);
}

// ---- 測る専用の計画（3-9）と年度を締める条件 ----

/** 測る専用の計画か（PLANS.purpose が MEASURE） */
function appPlanIsMeasure_(plan) {
  return String((plan && plan.purpose) || '') === APP_PLAN_MEASURE;
}

/** 年度を締めるとき、公式版が要る計画（測る専用の計画は外す。YearClose.js の appYearCheck_ から） */
function appYearPlansNeedingVersion_(plans) {
  return (plans || []).filter(p => !appPlanIsMeasure_(p));
}

/**
 * 年度を締める前に、年度の最後の四半期（1〜3 月）の当たりを数え終えたか（9 章 10。締めた年度の行は書けないため）。
 * 数え終えていなければ理由の文、済んでいれば ''。当たりの記録（HIT_RECORDS）を作るときに中身を書く（今は止めない）
 */
function appYearHitsPending_(fy, plans) {
  return '';
}

// ---- 一度だけの写し（版 10 の移行の後に 1 回。Setup.js の appRunBackfills_ が呼ぶ）----
// 返り値: { rows: 足した行の数 }（済み）・{ more: true, rows }（続きがある。次の操作でまた呼ぶ）・{ skipped: true }（仮のもの。記録しない）。
// 自分で appWithLock_ の中で appJournalRun_ を使って書く。キーは中身から決める（appStableLogId_）。1 回は数十秒までに収める。

/** 入力の記録の BASELINE（3-1: 今の 4 つの入力の表の行を 1 回だけ写す）。入力の記録を作るときに中身を書く */
function appV10BackfillInputLog_(ctx) {
  return { skipped: true };
}

/** AI 調査の記録の BASELINE（3-2: 今の AI_RESEARCH_STRUCTURED の 1 回分。中身は AiResearchLog.js） */
function appV10BackfillAiResearchLog_(ctx) {
  return appAiResearchBackfill_(ctx);
}

/** 当たりの記録（3-4: 版 10 の時点で締まっている四半期の分。calc_version に BACKFILL を添える）。当たりの記録を作るときに中身を書く */
function appV10BackfillHitRecords_(ctx) {
  return { skipped: true };
}
