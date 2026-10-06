/**
 * AutoResearch.js — 市場の動き（A-4 AI 調査）を、計画ごとに週 1 回、自動で集める（2026-10-06 村井さん承認「週1回自動」）。
 *
 * - 毎日 1 時ごろの時間主導のトリガー（triggerAutoResearch。所有者として動く）が、計画を 1 つ選んで A-4 を裏の処理（PLAN.RUN）で始める。
 *   選ぶのは、進行中（ACTIVE）・今年度か来年度・年度を締めていない計画のうち、最後に A-4 が終わってから 7 日を過ぎた（一度も無い）もので、いちばん古いもの。
 *   ほかの処理が動いている・待っているときは何もしない（裏の処理は 1 つずつ。Jobs.js）。
 * - 自動で始めた A-4 が終わったら（成功でも失敗でも）、5 時より前なら次に古い計画を始める（Jobs.js の appRunJob_ から appAutoResearchAfterJob_）。
 * - 始め方は画面から始めるのと同じ（appStartJob_）。頼んだ人の役割と年度の締めは appRunJob_ が段ごとに確かめ直し、段ごとに監査ログに残る。
 *   自動で始めた印は処理の中身（画面から渡せる項目だけ残る）ではなく、この状態の控え（始めた処理の ID）に置く（画面から偽れない）。
 * - トリガーは、毎日の手入れ（appDailyMaintenance_）と、所有者が画面を開いたとき（確かめるのは 1 日 1 回）に、無ければ作る。
 * - 止めるとき: Script Properties の AUTO_RESEARCH_ENABLED を 'false' にする（トリガーは残っても何もしない。既定は動かす）。
 * 1 回ごとに実行ログ（RUN、種類 AUTO.RESEARCH）に残す。状態は Script Properties（APP_AUTO_RESEARCH、小さい）に置く。
 * 予測の計算・保存のしかた・権限は変えない（A-4 は画面から始めたときと同じ処理で動く）。
 */
const APP_AUTO_RESEARCH_HANDLER = 'triggerAutoResearch';
const APP_AUTO_RESEARCH_HOUR = 1;               // 毎日 1 時ごろ（3 時ごろのバックアップより前に始める）
const APP_AUTO_RESEARCH_END_HOUR = 5;           // 続けて次の計画を始めてよいのは 5 時より前まで
const APP_AUTO_RESEARCH_STALE_DAYS = 7;         // 最後の A-4 からこれだけたったら集め直す
const APP_AUTO_RESEARCH_RETRY_DAYS = 3;         // 自動で試して失敗した・準備ができていなかった計画は、これだけあけてから試す
const APP_AUTO_RESEARCH_MAX_PER_NIGHT = 12;     // 1 晩に始める計画の上限（Vertex AI の費用の歯止め）
const APP_AUTO_RESEARCH_READY_CHECKS = 5;       // 1 回に準備を確かめる計画の数
const APP_AUTO_RESEARCH_HISTORY = 30;           // 自動で始めた A-4 の結果を残す数（新しい順）
const APP_AUTO_RESEARCH_FAILS = ['FAILED', 'STALLED', 'START_FAILED'];
const APP_AUTO_RESEARCH_PROP = 'APP_AUTO_RESEARCH';              // 状態（APP_JOB_ で始めない。処理の一覧に混ざる）
const APP_AUTO_RESEARCH_CHECKED_PROP = 'APP_AUTO_RESEARCH_CHECKED';   // トリガーを確かめた日（1 日 1 回だけ確かめる）
const APP_AUTO_RESEARCH_SWITCH = 'AUTO_RESEARCH_ENABLED';         // 'false' で止める
const APP_TRIGGER_LIMIT = 20;                   // トリガーは 1 人 1 プロジェクト 20 個まで
const APP_AUTO_RESEARCH_TRIGGER_SPARE = 3;      // 裏の処理（1 回だけのトリガー）のために空けておく数

/** 「今」（テストで差し替える） */
function appAutoResearchNow_() { return new Date(); }

function appAutoResearchIso_(d) {
  return Utilities.formatDate(d, APP_TZ, "yyyy-MM-dd'T'HH:mm:ssZ");
}

/** 止めるスイッチ: Script Properties の AUTO_RESEARCH_ENABLED が 'false'（か '0'）なら動かさない。無ければ動かす */
function appAutoResearchEnabled_() {
  const v = String(appProps_().getProperty(APP_AUTO_RESEARCH_SWITCH) || '').trim().toLowerCase();
  return v !== 'false' && v !== '0';
}

// ---- 状態の控え ----

function appAutoResearchState_() {
  try {
    const st = JSON.parse(appProps_().getProperty(APP_AUTO_RESEARCH_PROP) || '{}');
    return st && typeof st === 'object' ? st : {};
  } catch (e) {
    return {};
  }
}

/** 状態を読み・変えて・書く（ロックの中。fn の返り値を返す） */
function appAutoResearchUpdate_(fn) {
  return appWithLock_(() => {
    const st = appAutoResearchState_();
    const out = fn(st);
    const nowMs = appAutoResearchNow_().getTime();
    st.history = (st.history || []).slice(-APP_AUTO_RESEARCH_HISTORY);
    const tried = st.tried || {};
    Object.keys(tried).forEach(k => { if (!(nowMs - new Date(tried[k].at).getTime() <= 30 * 864e5)) delete tried[k]; });   // 30 日より前の試しは忘れる
    st.tried = tried;
    while (appUtf8Bytes_(JSON.stringify(st)) > 8000 && st.history.length) st.history.shift();   // 1 つの値の上限は約 9KB（バイト数）
    appProps_().setProperty(APP_AUTO_RESEARCH_PROP, JSON.stringify(st));
    return out;
  });
}

/** 計画を試した結果を控える（toHistory が false なら、結果の一覧には入れない） */
function appAutoResearchNote_(st, planId, status, nowMs, reason, toHistory) {
  const at = appAutoResearchIso_(new Date(nowMs));
  (st.tried = st.tried || {})[planId] = { at: at, status: status, reason: reason ? String(reason).slice(0, 120) : '' };
  if (toHistory !== false) (st.history = st.history || []).push({ at: at, planId: planId, status: status });
}

// ---- 対象の計画 ----

/**
 * 自動で集める対象の計画: 進行中・今年度（fy）か来年度・年度を締めていない。
 * 最後に A-4 が終わった時刻（PLAN_ACTIONS の DONE。画面から動かしたものも数える。無ければ 0）の古い順
 */
function appAutoResearchPlans_(fy) {
  const last = {};
  appReadTable_('PLAN_ACTIONS').forEach(r => {
    if (r.action !== 'AI.RESEARCH' || r.status !== 'DONE') return;
    const t = new Date(r.finished_at).getTime();
    if (isFinite(t) && t > (last[r.plan_id] || 0)) last[r.plan_id] = t;
  });
  const names = appClientNameMap_();
  return appReadTable_('PLANS')
    .filter(p => p.state === 'ACTIVE' && (Number(p.fy) === fy || Number(p.fy) === fy + 1) && !appAutoResearchFrozen_(p.fy))
    .map(p => ({ planId: p.plan_id, clientId: p.client_id, clientName: names[p.client_id] || p.client_label, fy: String(p.fy), lastAt: last[p.plan_id] || 0 }))
    .sort((a, b) => a.lastAt - b.lastAt || (a.planId < b.planId ? -1 : a.planId > b.planId ? 1 : 0));
}

/** 年度を締めたか（確かめられなければ締めたとみなす: 自動では書かない） */
function appAutoResearchFrozen_(fy) {
  try { return appYearIsFrozen_(fy); } catch (e) { return true; }
}

/** 最後の A-4 から 7 日を過ぎた（一度も無い）計画 */
function appAutoResearchStale_(plans, nowMs) {
  return plans.filter(p => !p.lastAt || nowMs - p.lastAt >= APP_AUTO_RESEARCH_STALE_DAYS * 864e5);
}

/**
 * A-4 を始められるか。旧来の A-4 が最初に確かめることと同じものを、データ本体から読むだけで確かめる:
 * CONFIG の AI_RESEARCH_ENABLED が 1 以上・VERTEX_PROJECT_ID / VERTEX_LOCATION / VERTEX_GEMINI_MODEL が入っている・A-2 売上の取り込みが成功している。
 * 読めないときは始める（処理そのものが理由つきで止まる）
 */
function appAutoResearchReady_(planId) {
  try {
    const book = appStoreBook_(appPlanOf_(planId));
    const sh = book.getSheetByName('CONFIG');
    if (!sh) return { ok: false, reason: 'NO_SETUP' };
    const cfg = {};
    if (sh.getLastRow() >= 1) {
      sh.getRange(1, 1, sh.getLastRow(), 2).getValues().forEach(r => {
        const s = String(r[0] || '').trim();
        const i = s.search(/[（(]/);   // 旧来の configKeyOf_ と同じ: 「（」「(」より前がキー
        const k = (i >= 0 ? s.slice(0, i) : s).trim();
        if (k) cfg[k] = r[1];
      });
    }
    if (!(Number(cfg.AI_RESEARCH_ENABLED || 0) > 0)) return { ok: false, reason: 'AI_DISABLED' };
    if (!['VERTEX_PROJECT_ID', 'VERTEX_LOCATION', 'VERTEX_GEMINI_MODEL'].every(k => String(cfg[k] || '').trim())) return { ok: false, reason: 'VERTEX_CONFIG' };
    const ps = book.getSheetByName('PROCESS_STATUS');
    const step1 = (ps ? ps.getDataRange().getValues() : []).filter(r => r[0] === 'step1_status')[0];
    if (!step1 || step1[3] !== 'success') return { ok: false, reason: 'NO_SALES_DATA' };
    return { ok: true };
  } catch (e) {
    return { ok: true, note: 'CHECK_FAILED' };
  }
}

/** 動いている・待っている裏の処理（止まったものと、始まらずに 15 分たったものは数えない。appStartJob_ と同じ） */
function appAutoResearchBusy_() {
  const now = new Date().getTime();
  return appJobList_().filter(j => (j.status === 'QUEUED' && now - new Date(j.createdAt).getTime() <= APP_JOB_QUEUE_EXPIRE_MS) ||
    (j.status === 'RUNNING' && !appJobIsStale_(j)))[0] || null;
}

// ---- 始める ----

/** 毎日 1 時ごろ（triggerAutoResearch）: 計画を 1 つ選んで A-4 を始める */
function appAutoResearchTick_(ctx) {
  return appAutoResearchNext_(ctx, 'TICK');
}

/**
 * 次の計画を 1 つ選んで A-4 を始める（毎日のトリガー TICK と、自動で始めた A-4 が終わった後の続き CHAIN の共通）。
 * 始められない理由（止めてある・ほかの処理が動いている・対象が無い・5 時を過ぎた）は例外にせず返す。毎回、実行ログに残す
 */
function appAutoResearchNext_(ctx, via) {
  const t0 = new Date().getTime();
  appStoreForget_();   // 同じ実行の中で書いた表も読み直す（続きのときは、終わった A-4 の記録を数える）
  const now = appAutoResearchNow_();
  const nowMs = now.getTime();
  const nowIso = appAutoResearchIso_(now);
  const night = Utilities.formatDate(now, APP_TZ, 'yyyy-MM-dd');
  const extra = {};
  const finish = (status, reason, more, error) => {
    const res = Object.assign({ via: via, status: status, reason: reason || '' }, extra, more || {});
    try {
      appAutoResearchUpdate_(st => {
        if (via === 'TICK') st.lastTickAt = nowIso;
        st.lastTick = { at: nowIso, via: via, status: status, reason: reason || '', planId: res.planId || '' };
      });
    } catch (e) { Logger.log('自動の AI 調査の状態を控えられません: ' + (e && e.message ? e.message : e)); }
    appRunLog_({ requestId: ctx.requestId, kind: 'AUTO.RESEARCH', status: status, durationMs: new Date().getTime() - t0, detail: res, error: error || '' });
    return res;
  };
  if (!appAutoResearchEnabled_()) return finish('SKIPPED', 'DISABLED');
  if (via === 'CHAIN' && Number(Utilities.formatDate(now, APP_TZ, 'HH')) >= APP_AUTO_RESEARCH_END_HOUR) return finish('SKIPPED', 'AFTER_HOURS');
  const st = appAutoResearchState_();
  const count = st.night && st.night.date === night ? Number(st.night.count || 0) : 0;
  if (count >= APP_AUTO_RESEARCH_MAX_PER_NIGHT) return finish('SKIPPED', 'NIGHT_LIMIT', { count: count });
  const busy = appAutoResearchBusy_();
  if (busy) return finish('SKIPPED', 'BUSY', { busyKind: busy.kind });
  const plans = appAutoResearchPlans_(appFy_(now));
  if (st.chain) {
    // 前に自動で始めた A-4 が、終わりの知らせなしに止まっていた（1 回の上限で打ち切られたなど）。
    // 始めた後に A-4 の記録があれば終わっていた（知らせだけが届かなかった）。無ければ止まったとして控える
    const c = st.chain;
    const p = plans.filter(x => x.planId === c.planId)[0];
    const result = p && p.lastAt >= new Date(c.startedAt).getTime() ? 'DONE' : 'STALLED';
    appAutoResearchUpdate_(s => { if (s.chain && s.chain.jobId === c.jobId) { s.chain = null; appAutoResearchNote_(s, c.planId, result, nowMs); } });
    if (result === 'STALLED') extra.stalled = c.planId;
  }
  const tried = appAutoResearchState_().tried || {};
  const due = appAutoResearchStale_(plans, nowMs).filter(p => {
    const t = tried[p.planId];
    return !(t && nowMs - new Date(t.at).getTime() < APP_AUTO_RESEARCH_RETRY_DAYS * 864e5);
  });
  if (!due.length) return finish('SKIPPED', 'NO_STALE_PLAN');
  let pick = null;
  const notReady = [];
  for (let i = 0; i < due.length && i < APP_AUTO_RESEARCH_READY_CHECKS && !pick; i++) {
    const r = appAutoResearchReady_(due[i].planId);
    if (r.ok) pick = due[i]; else notReady.push({ planId: due[i].planId, reason: r.reason });
  }
  if (notReady.length) {
    extra.notReady = notReady;
    appAutoResearchUpdate_(s => notReady.forEach(n => appAutoResearchNote_(s, n.planId, 'NOT_READY', nowMs, n.reason, false)));
  }
  if (!pick) return finish('SKIPPED', 'NOT_READY');
  const payload = { planId: pick.planId, action: 'AI.RESEARCH' };
  // 画面から始めるときと同じ役割を確かめる（裏の実行でも appRunJob_ がもう一度確かめる）
  if (!appHasRole_(ctx.roles, appJobMinRole_(appJobSpec_('PLAN.RUN'), payload), pick.clientId)) return finish('SKIPPED', 'DENIED', { planId: pick.planId });
  let started;
  try {
    started = appStartJob_(ctx, { kind: 'PLAN.RUN', payload: payload });
  } catch (e) {
    const msg = String(e && e.message ? e.message : e);
    appAutoResearchUpdate_(s => appAutoResearchNote_(s, pick.planId, 'START_FAILED', nowMs, msg));
    return finish('FAILED', 'START_FAILED', { planId: pick.planId }, msg);
  }
  appAutoResearchUpdate_(s => {
    s.chain = { jobId: started.jobId, planId: pick.planId, requestedBy: ctx.actor, startedAt: nowIso };
    s.lastStartedPlan = { planId: pick.planId, clientName: pick.clientName, fy: pick.fy, jobId: started.jobId, at: nowIso, via: via };
    s.night = { date: night, count: count + 1 };
    appAutoResearchNote_(s, pick.planId, 'STARTED', nowMs, '', false);
  });
  return finish('STARTED', '', { planId: pick.planId, fy: pick.fy, jobId: started.jobId,
    lastAt: pick.lastAt ? appAutoResearchIso_(new Date(pick.lastAt)) : '' });
}

/** 続きの処理をたどった、最初の処理の ID */
function appAutoResearchRootOf_(job) {
  let cur = job;
  for (let hop = 0; hop < 100 && cur && cur.parentId; hop++) cur = appJobGet_(cur.parentId);
  return cur ? cur.id : '';
}

/**
 * 裏の処理が続きなしで終わったとき（Jobs.js の appRunJob_ から。DONE か FAILED）。
 * 自動で始めた A-4 なら結果を控え、5 時より前なら次に古い計画を始める（頼んだ人は同じ）。
 * 普通の処理の結果には関係しない（呼ぶ側で例外を握りつぶす）。自動の A-4 でなければ、状態を 1 回読むだけで戻る
 */
function appAutoResearchAfterJob_(job, status) {
  if (!job || job.nextJobId) return null;
  const chain = appAutoResearchState_().chain;
  if (!chain || job.requestedBy !== chain.requestedBy) return null;
  const p = job.payload || {};
  if (p.action !== 'AI.RESEARCH' || p.planId !== chain.planId || appAutoResearchRootOf_(job) !== chain.jobId) return null;
  const now = appAutoResearchNow_();
  const result = status === 'DONE' ? 'DONE' : 'FAILED';
  const error = result === 'FAILED' ? String(job.error || '') : '';
  const mine = appAutoResearchUpdate_(st => {
    if (!st.chain || st.chain.jobId !== chain.jobId) return false;   // ほかの実行が先に控えた
    st.chain = null;
    appAutoResearchNote_(st, chain.planId, result, now.getTime(), error);
    const sp = st.lastStartedPlan && st.lastStartedPlan.planId === chain.planId ? st.lastStartedPlan : {};
    st.lastResult = { at: appAutoResearchIso_(now), planId: chain.planId, clientName: sp.clientName || '', fy: sp.fy || '', status: result,
      jobId: chain.jobId, error: error.slice(0, 200) };
    return true;
  });
  if (!mine) return null;
  appRunLog_({ requestId: job.id, kind: 'AUTO.RESEARCH', status: result, detail: { via: 'JOB', event: 'FINISHED', planId: chain.planId, jobId: chain.jobId, lastJobId: job.id },
    error: error });
  return appAutoResearchNext_(appJobContext_(job), 'CHAIN');
}

// ---- トリガー ----

/**
 * 毎日のトリガー（triggerAutoResearch、1 時ごろ）が無ければ作る（何度呼んでも 1 つ）。確かめるのは 1 日 1 回だけ（日付を控える）。
 * 毎日の手入れと、所有者が画面を開いたときに呼ぶ。止めるスイッチが切ってあるときと、トリガーの上限（20）に余裕が無いときは作らない
 */
function appEnsureAutoResearchTrigger_(ctx) {
  if (!appIsSetUp_()) return { checked: false, reason: 'NOT_SET_UP' };
  const props = appProps_();
  const today = appToday_();
  const prev = (() => { try { return JSON.parse(props.getProperty(APP_AUTO_RESEARCH_CHECKED_PROP) || 'null'); } catch (e) { return null; } })();
  if (prev && prev.date === today) return { checked: false, reason: 'CHECKED_TODAY', installed: !!prev.installed };
  const mark = (installed, reason) => {
    props.setProperty(APP_AUTO_RESEARCH_CHECKED_PROP, JSON.stringify({ date: today, installed: !!installed, reason: reason || '' }));
  };
  mark(prev && prev.installed, 'CHECKING');   // 先に控える（途中で失敗しても、その日のうちは開くたびにやり直さない）
  const all = ScriptApp.getProjectTriggers();
  const mine = all.filter(t => t.getHandlerFunction() === APP_AUTO_RESEARCH_HANDLER);
  mine.slice(1).forEach(t => ScriptApp.deleteTrigger(t));   // 二重にできていたら 1 つにする（同時に確かめた場合）
  if (mine.length) { mark(true, ''); return { checked: true, installed: true, created: false }; }
  if (!appAutoResearchEnabled_()) { mark(false, 'DISABLED'); return { checked: true, installed: false, created: false, reason: 'DISABLED' }; }
  if (all.length + 1 + APP_AUTO_RESEARCH_TRIGGER_SPARE > APP_TRIGGER_LIMIT) {
    mark(false, 'TRIGGER_LIMIT');
    appRunLog_({ requestId: ctx && ctx.requestId, kind: 'AUTO.RESEARCH.TRIGGER', status: 'SKIPPED', detail: { reason: 'TRIGGER_LIMIT', triggers: all.length } });
    return { checked: true, installed: false, created: false, reason: 'TRIGGER_LIMIT' };
  }
  ScriptApp.newTrigger(APP_AUTO_RESEARCH_HANDLER).timeBased().everyDays(1).atHour(APP_AUTO_RESEARCH_HOUR).nearMinute(0).create();
  mark(true, '');
  appRunLog_({ requestId: ctx && ctx.requestId, kind: 'AUTO.RESEARCH.TRIGGER', status: 'CREATED', detail: { hour: APP_AUTO_RESEARCH_HOUR, by: ctx && ctx.actor } });
  return { checked: true, installed: true, created: true };
}

/** 画面を開いたとき（doGet）: 所有者なら、トリガーを確かめる（1 日 1 回だけ。失敗しても画面は開く） */
function appAutoResearchOnOpen_(ctx) {
  if (!ctx || !ctx.user || !ctx.user.isOwner) return;
  try { appEnsureAutoResearchTrigger_(ctx); } catch (e) { Logger.log('自動の AI 調査のトリガー: ' + (e && e.message ? e.message : e)); }
}

// ---- 状態（ホームの仕組みの状態） ----

/**
 * 自動の AI 調査の状態: { enabled, trigger, lastTickAt, lastTick, lastStartedPlan, lastResult, running, failures7d, stalePlans }
 * stalePlans = 最後の A-4 から 7 日を過ぎた対象の計画（古い順）と、自動で最後に試した結果（準備ができていない理由など）
 */
function appAutoResearchStatus_() {
  const st = appAutoResearchState_();
  const now = appAutoResearchNow_();
  const nowMs = now.getTime();
  const fy = appFy_(now);
  let plans = null;
  try { plans = appCachedRead_('AUTO_RESEARCH\u0001' + fy, () => appAutoResearchPlans_(fy)); } catch (e) { plans = null; }
  const tried = st.tried || {};
  const trig = (() => { try { return JSON.parse(appProps_().getProperty(APP_AUTO_RESEARCH_CHECKED_PROP) || 'null'); } catch (e) { return null; } })();
  const chain = st.chain && nowMs - new Date(st.chain.startedAt).getTime() < 2 * 3600 * 1000 ? st.chain : null;   // 2 時間より前のものは止まったとみなす
  return {
    enabled: appAutoResearchEnabled_(),
    trigger: trig ? { installed: !!trig.installed, checkedOn: trig.date || '', reason: trig.reason || '' } : null,
    lastTickAt: st.lastTickAt || '',
    lastTick: st.lastTick || null,
    lastStartedPlan: st.lastStartedPlan || null,
    lastResult: st.lastResult || null,
    running: chain ? { planId: chain.planId, jobId: chain.jobId, startedAt: chain.startedAt } : null,
    failures7d: (st.history || []).filter(h => APP_AUTO_RESEARCH_FAILS.indexOf(h.status) >= 0 && nowMs - new Date(h.at).getTime() <= 7 * 864e5).length,
    stalePlans: plans === null ? null : appAutoResearchStale_(plans, nowMs).map(p => ({ planId: p.planId, clientName: p.clientName, fy: p.fy,
      lastAt: p.lastAt ? appAutoResearchIso_(new Date(p.lastAt)) : '',
      lastTry: tried[p.planId] ? { at: tried[p.planId].at, status: tried[p.planId].status, reason: tried[p.planId].reason || '' } : null }))
  };
}
