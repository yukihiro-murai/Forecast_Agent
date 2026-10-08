/**
 * InputLog.js — 入力の記録（INPUT_LOG。SCHEMA_PLAN_v10-12_JA.md の 3-1。2026-10-08 村井さん承認「版10はおすすめで」）。
 * - 入力の 4 つの表（ENG_PRODUCT・ENG_CLIENT・ENG_OPINIONS・ENG_DEV_SPOT）が変わる保存と実行のすべて（入力の保存・売上の取り込み・計画を作る ほか）で、
 *   変わった行だけを前と後で残す（Forecast.js の appChangedOps_ から。本体の保存と同じ控えに入れて一緒に書く）。前の行は書く前のデータ本体から読む。
 * - 行の印 row_key = 種類|担当者|製品|月（スポットは案件名）。同じ印が 2 行あれば 2 つ目から #2・#3（上からの順）。
 *   前と後は同じ印どうしで比べ、印の無くなった行と新しい印の行は、同じ位置（行の番号）なら「変えた」とみなす（担当者を入れた・月を変えた）。
 * - 自信（self_conf: 高い・ふつう・低い・空）は製品とメーカー全体の行だけ。入力の画面で選び、この記録にだけ書く（旧来の表には渡さない。計算に使わない）。
 *   行の中身が同じで自信だけ変えたら CONF。入力の保存のほかの操作で変えた・外した行は、その行の一番新しい自信を引き継ぐ（足した行は空）。
 * - 見せる範囲（appInputLogView_）: 予算策定担当以上（そのメーカーの担当を含む）は全部。ほかの人は、人のつなぎ（PERSON_LINKS）で本人と分かる行だけで、
 *   その中でも担当者の名前と保存した人は送らない。ほかの行は種類ごとの件数だけ（V10.js の appLogViewer_・appLogVisible_）。
 * - 版 10 の移行の後に 1 回、今の 4 つの表の行を BASELINE として写す（appInputLogBackfill_。前の保存の履歴は残っていないので作れない）。
 * 行は消さず、書き換えない。記録は計算に使わない。
 */

/** 入力の 4 つの表（旧来のシート → 記録の種類） */
const APP_INPUT_LOG_SHEETS = { PRODUCT: 'product', CLIENT: 'client', OPINIONS: 'opinions', DEV_SPOT: 'devspot' };
/** 種類ごとの行の項目（旧来の見出しの順。画面の入力の行と同じ名前。before_json・after_json の中身） */
const APP_INPUT_LOG_FIELDS = {
  product: ['person', 'product', 'ym', 'step', 'reason'],
  client: ['person', 'ym', 'step', 'reason'],
  opinions: ['person', 'ym', 'step', 'conf', 'note'],
  devspot: ['person', 'ym', 'project', 'amount', 'conf']
};
/** 行の印に使う項目（担当者・製品・月・スポットの名前） */
const APP_INPUT_LOG_KEYS = { product: ['person', 'product', 'ym'], client: ['person', 'ym'], opinions: ['person', 'ym'], devspot: ['person', 'ym', 'project'] };
/** 自信の段階（9 章 7。計算には使わない）と、選べる種類（製品とメーカー全体の行だけ） */
const APP_INPUT_CONF = ['高い', 'ふつう', '低い'];
const APP_INPUT_CONF_KINDS = ['product', 'client'];
/** 保存の理由（任意）の上限 */
const APP_INPUT_REASON_MAX = 200;
/** 画面に送る、1 行あたりの変わった跡の数（新しい順） */
const APP_INPUT_LOG_HIST = 5;
/** 一度だけの写しの 1 回の目安（それを過ぎたら続きは次の操作で） */
const APP_INPUT_LOG_BACKFILL_MS = 20000;
/** 操作を渡さない道（控えの番号の頭 → 操作）: 計画を作る（Portfolio.js）・予測（Forecast.js）・学習の事前分布（Learning.js） */
const APP_INPUT_LOG_BATCH = { NEW: 'PLAN.CREATE', RUN: 'FORECAST.RUN', POOL: 'LEARN.POOL' };

// ---- 行の形と印 ----

/** 種類の旧来のシートの名前 */
function appInputLogSheet_(kind) {
  return Object.keys(APP_INPUT_LOG_SHEETS).filter(s => APP_INPUT_LOG_SHEETS[s] === kind)[0];
}

/**
 * ENG_ の表の 1 行（型つきの文字列）を、記録の行の形にする。値が 1 つも無い行は null（旧来の画面の読み取り webParseInputs_ と同じく数えない）。
 * 日付は月（yyyy-MM。月の欄のほかは yyyy-MM-dd）にする。数は数のまま
 */
function appInputLogRow_(kind, o) {
  const header = APP_ENGINE_SHEETS[appInputLogSheet_(kind)].header;
  const types = String(o._types || '');
  const out = {};
  let any = false;
  APP_INPUT_LOG_FIELDS[kind].forEach((f, j) => {
    const t = types.charAt(j) || 'e';
    const raw = o[header[j]] === undefined || o[header[j]] === null ? '' : o[header[j]];
    let v = t === 'f' ? '=' + String(raw).replace(/^=/, '') : appCellDecode_(t, raw);
    if (appIsDate_(v)) v = Utilities.formatDate(v, APP_TZ, f === 'ym' ? 'yyyy-MM' : 'yyyy-MM-dd');
    if (v !== '' && v !== null) any = true;
    out[f] = v;
  });
  return any ? out : null;
}

/** 行の印（同じ印の 2 つ目からの #n は appInputLogList_ が付ける） */
function appInputLogKey_(kind, row) {
  return kind + '|' + APP_INPUT_LOG_KEYS[kind].map(f => String(row[f] === null || row[f] === undefined ? '' : row[f]).trim()).join('|');
}

/** ENG_ の表の行（計画 1 つ・1 種類）を、行の番号の順に、印つきの記録の形にする（空の行は除く）。返り値 [{ seq, row, key }] */
function appInputLogList_(kind, objs) {
  const seen = {};
  return (objs || []).slice().sort((a, b) => Number(a.seq) - Number(b.seq)).map(o => ({ seq: Number(o.seq), row: appInputLogRow_(kind, o) }))
    .filter(x => x.row).map(x => {
      const k = appInputLogKey_(kind, x.row);
      seen[k] = (seen[k] || 0) + 1;
      return { seq: x.seq, row: x.row, key: seen[k] > 1 ? k + '#' + seen[k] : k };
    });
}

/**
 * 前と後の行を比べる。返り値 [{ change: 'ADD'|'CHANGE'|'REMOVE'|''（同じ）, a: 後の行, b: 前の行 }]。
 * 同じ印どうしで比べる。印の無くなった前の行と新しい印の後の行は、同じ位置（行の番号）なら CHANGE（担当者を入れた・月を変えた）、ほかは REMOVE と ADD
 */
function appInputLogDiff_(before, after) {
  const bk = {}, ak = {};
  before.forEach(b => { bk[b.key] = b; });
  after.forEach(a => { ak[a.key] = a; });
  const gone = {};
  before.forEach(b => { if (!ak[b.key]) gone[b.seq] = b; });
  const paired = {};
  const out = [];
  after.forEach(a => {
    const b = bk[a.key];
    if (b) { out.push({ change: JSON.stringify(a.row) === JSON.stringify(b.row) ? '' : 'CHANGE', a: a, b: b }); return; }
    const p = gone[a.seq];
    if (p && !paired[p.key]) { paired[p.key] = true; out.push({ change: 'CHANGE', a: a, b: p }); return; }
    out.push({ change: 'ADD', a: a, b: null });
  });
  before.forEach(b => { if (!ak[b.key] && !paired[b.key]) out.push({ change: 'REMOVE', a: null, b: b }); });
  return out;
}

// ---- 書く（本体の保存と同じ控えに入れる）----

/**
 * 入力の保存（INPUT.SAVE）の中身から、自信と保存の理由を取り出す（旧来の webSaveInputs には渡さない。旧来の表に列を足さない）。
 * info を埋めて、旧来に渡す中身（写し）を返す: info.reason = 保存の理由、info.conf = { kind, bySeq: { 書く行の番号: 自信（送らない行は null） } }。
 * 書く行の番号は旧来と同じ数え方（値の無い行を飛ばして 2 行目から詰める）。自信は製品とメーカー全体の行だけ。ほかの操作はそのまま返す
 */
function appInputLogTake_(args, info) {
  if (!info || info.action !== 'INPUT.SAVE' || !args || typeof args !== 'object' || Array.isArray(args)) return args;
  const out = Object.assign({}, args);
  const reason = out.reason === undefined || out.reason === null ? '' : String(out.reason).trim();
  if (reason.length > APP_INPUT_REASON_MAX) throw new Error('保存の理由は ' + APP_INPUT_REASON_MAX + ' 字までです。');
  delete out.reason;
  info.reason = reason;
  if (!Array.isArray(out.rows)) return out;
  const kind = String(out.kind || '');
  const useConf = APP_INPUT_CONF_KINDS.indexOf(kind) >= 0;
  const bySeq = {};
  let n = 0;
  out.rows = out.rows.map(r => {
    if (!r || typeof r !== 'object') return r;
    const x = Object.assign({}, r);
    const c = x.selfConf === undefined || x.selfConf === null ? null : String(x.selfConf).trim();   // 送らない（null）= 今の自信のまま。空 = 外した
    delete x.selfConf;
    if (c && APP_INPUT_CONF.indexOf(c) < 0) throw new Error('自信は「高い」「ふつう」「低い」から選んでください。');
    // 旧来の webSaveInputs と同じ見分け方（どの値も空なら書かない）
    if (Object.keys(x).some(k => String(x[k] || '').trim() !== '')) { n++; if (useConf) bySeq[n + 1] = c; }
    return x;
  });
  if (useConf) info.conf = { kind: kind, bySeq: bySeq };
  return out;
}

/**
 * 入力の記録を足す控えの書き方（ops。appChangedOps_ が本体の ops に足す）。changed = 控えの形にしたシート（appCaptureChanged_ の changed）、
 * batchId = 操作の番号（action_id）、info = { action, reason, conf }（appInputLogTake_。無ければ番号の頭から操作を決める）。
 * 入力の表が変わっていなくても、自信だけ変えた行は CONF として足す。前の行と一番新しい自信は、書く前のデータ本体から読む
 */
function appInputLogOps_(ctx, planId, changed, batchId, info) {
  const o = info || {};
  const enc = {};
  (changed || []).forEach(e => {
    const k = APP_INPUT_LOG_SHEETS[e.sheetRow.sheet];
    if (k && e.sheetRow.mode === 'table') enc[k] = e;   // 見出しが変わって行ごとに持つシートは比べない（旧来は見出しを変えない）
  });
  const confKind = o.conf && APP_INPUT_CONF_KINDS.indexOf(o.conf.kind) >= 0 ? o.conf.kind : '';
  if (!Object.keys(enc).length && !confKind) return [];
  const action = o.action || APP_INPUT_LOG_BATCH[String(batchId || '').split('-')[0]] || '';
  const now = appNowIso_();
  let latest = null;
  const confOf = key => { if (!latest) latest = appInputLogLatestConf_(planId); return latest[key] || ''; };
  const rows = [];
  Object.keys(APP_INPUT_LOG_FIELDS).forEach(kind => {
    const asked = kind === confKind ? o.conf.bySeq : null;   // 画面が渡した自信（書く行の番号 → 自信）
    if (!enc[kind] && !asked) return;
    const stored = appInputLogList_(kind, appReadPlanTable_('ENG_' + appInputLogSheet_(kind), planId));
    const next = enc[kind] ? appInputLogList_(kind, enc[kind].tableRows) : stored;
    const withConf = APP_INPUT_CONF_KINDS.indexOf(kind) >= 0;
    appInputLogDiff_(stored, next).forEach(d => {
      let change = d.change;
      let conf = '';
      if (withConf) {
        const said = asked && d.a ? asked[d.a.seq] : null;   // 画面が選んだ自信（送らなければ null）
        if (said !== null && said !== undefined) conf = said;
        else if (d.b && change !== '') conf = confOf(d.b.key);   // 自信を送らない保存・ほかの操作: 変えた・外した行は今の自信を引き継ぐ（足した行は空）
        if (change === '') { if (said === null || said === undefined || conf === confOf(d.a.key)) return; change = 'CONF'; }
      } else if (change === '') return;
      const x = d.a || d.b;
      rows.push({ plan_id: planId, log_id: appNewLogId_(APP_LOG_PREFIX.INPUT_LOG), action_id: String(batchId || ''), action: action, kind: kind,
        change: change, row_key: x.key, person: String(x.row.person === null || x.row.person === undefined ? '' : x.row.person),
        before_json: change === 'ADD' || change === 'CONF' ? '' : d.b.row, after_json: change === 'REMOVE' ? '' : d.a.row,
        self_conf: conf, reason: o.reason || '', signal_id: '', actor_email: ctx.actor, saved_at: now });
    });
  });
  return appLogOps_('INPUT_LOG', rows);
}

// ---- 読む ----

/** 計画の入力の記録（古い順。同じ秒なら BASELINE が先。表が読めなければ空） */
function appInputLogRead_(planId) {
  let rows;
  try { rows = appReadPlanTable_('INPUT_LOG', planId); } catch (e) { return []; }
  const k = r => [r.saved_at, r.change === 'BASELINE' ? '0' : '1', r.log_id].join('\u0001');
  return rows.map(appStripRow_).sort((a, b) => (k(a) < k(b) ? -1 : k(a) > k(b) ? 1 : 0));
}

/** 行の印ごとの一番新しい自信（製品とメーカー全体。外した行の記録は見ない） */
function appInputLogLatestConf_(planId) {
  const out = {};
  appInputLogRead_(planId).forEach(r => {
    if (APP_INPUT_CONF_KINDS.indexOf(r.kind) >= 0 && r.change !== 'REMOVE') out[r.row_key] = String(r.self_conf || '');
  });
  return out;
}

/** 記録の JSON の列を読む（空・読めなければ null） */
function appInputLogJson_(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'object') return v;
  try { return JSON.parse(String(v)); } catch (e) { return null; }
}

/**
 * 計画の画面に付ける入力の記録（apiPlanView。見る人ごと）。rows[種類][i] は画面の入力の行（boot.input[種類][i]）と同じ並び
 * （同じ ENG_ の表から、同じく空の行を除いて作る）。行ごとに { conf: 一番新しい自信, n: 記録の数, hist: 新しい順の跡 }。記録の無い行は null。
 * 予算策定担当以上でない人には、本人と分かる行だけを送り、その中の担当者の名前と保存した人は送らない。ほかは種類ごとの件数（counts）だけ。
 * 版がそろう前・読めないときは null（画面はそのまま出す）
 */
function appInputLogView_(ctx, plan) {
  if (!appLogReady_()) return null;
  try {
    const viewer = appLogViewer_(ctx, plan.client_id);
    const vis = appLogVisible_(viewer, appInputLogRead_(plan.plan_id), { personOf: r => r.person, typeOf: r => r.kind });
    const by = {};
    vis.rows.forEach(r => { (by[r.row_key] = by[r.row_key] || []).unshift(r); });   // 新しい順
    const hideName = j => {
      const x = appInputLogJson_(j);
      if (!x || viewer.full) return x;
      const c = Object.assign({}, x);
      delete c.person;
      return c;
    };
    const rows = {};
    Object.keys(APP_INPUT_LOG_SHEETS).forEach(sheet => {
      const kind = APP_INPUT_LOG_SHEETS[sheet];
      rows[kind] = appInputLogList_(kind, appReadPlanTable_('ENG_' + sheet, plan.plan_id)).map(x => {
        const h = by[x.key];
        if (!h) return null;
        const live = h.filter(r => r.change !== 'REMOVE')[0];
        return { conf: APP_INPUT_CONF_KINDS.indexOf(kind) >= 0 && live ? String(live.self_conf || '') : '', n: h.length,
          hist: h.slice(0, APP_INPUT_LOG_HIST).map(r => ({ at: r.saved_at, change: r.change, action: r.action, by: viewer.full ? r.actor_email : '',
            reason: viewer.full ? r.reason : '', conf: r.self_conf, before: hideName(r.before_json), after: hideName(r.after_json) })) };
      });
    });
    return { full: viewer.full, rows: rows, counts: vis.counts, hidden: vis.hidden };
  } catch (e) {
    appLogError_('INPUT_LOG.VIEW', e, ctx);
    return null;
  }
}

// ---- 一度だけの写し（版 10 の移行の後。V10.js の appV10BackfillInputLog_ から）----

/**
 * 今の 4 つの入力の表の行を、計画ごとに BASELINE として 1 回だけ写す（締めた年度の計画は飛ばす）。キーは計画・印から決める（appStableLogId_。
 * 2 つの操作が同時に動かしても二重にならない）。BASELINE がもうある計画は飛ばす。1 回は約 20 秒までで、残りは次の操作で（{ more: true }）
 */
function appInputLogBackfill_(ctx) {
  if (!appLogReady_()) return { more: true, rows: 0 };
  const t0 = new Date().getTime();
  const done = {};
  let log;
  try { log = appReadTable_('INPUT_LOG'); } catch (e) { return { more: true, rows: 0 }; }   // 表がまだ無い（移行の前）: 次の操作でまた
  log.forEach(r => { if (r.change === 'BASELINE') done[r.plan_id] = true; });
  const plans = appReadTable_('PLANS').filter(p => !done[p.plan_id] && !appYearIsFrozen_(p.fy));
  let rows = 0;
  for (let i = 0; i < plans.length; i++) {
    if (new Date().getTime() - t0 > APP_INPUT_LOG_BACKFILL_MS) return { more: true, rows: rows };
    const plan = plans[i];
    rows += appWithLock_(() => {
      if (appReadPlanTable_('INPUT_LOG', plan.plan_id).some(r => r.change === 'BASELINE')) return 0;   // ほかの操作が先に写した
      const now = appNowIso_();
      const list = [];
      Object.keys(APP_INPUT_LOG_SHEETS).forEach(sheet => {
        const kind = APP_INPUT_LOG_SHEETS[sheet];
        appInputLogList_(kind, appReadPlanTable_('ENG_' + sheet, plan.plan_id)).forEach(x => list.push({
          plan_id: plan.plan_id, log_id: appStableLogId_(APP_LOG_PREFIX.INPUT_LOG, ['BASELINE', plan.plan_id, x.key]), action_id: '', action: 'BASELINE',
          kind: kind, change: 'BASELINE', row_key: x.key, person: String(x.row.person === null || x.row.person === undefined ? '' : x.row.person),
          before_json: '', after_json: x.row, self_conf: '', reason: '', signal_id: '', actor_email: ctx.actor || '', saved_at: now }));
      });
      if (!list.length) return 0;
      appJournalRun_(ctx, '入力の記録の初めの姿（' + plan.plan_id + '）', plan.plan_id, appLogOps_('INPUT_LOG', list));
      return list.length;
    });
  }
  return { rows: rows };
}
