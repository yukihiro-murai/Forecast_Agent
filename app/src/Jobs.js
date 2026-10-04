/**
 * Jobs.js — 時間のかかる処理（旧ブックの試し読み・取り込み、今後の予測の実行など）を、画面の通信から切り離して動かす。
 * 画面から始めると、1 回だけ動く時間主導のトリガーを作り、裏で最後まで実行する。画面は数秒ごとに状態を尋ねて結果を受け取る。
 * 画面と 1 分以上つながったままにすると通信が切れる（HTTP 503）ことがあるため（2026-10-01 に旧ブックの試し読みで発生）。
 * 状態は Script Properties（APP_JOB_<id>、小さい）、結果は CacheService（6 時間）に置く。処理そのものの記録は監査ログに残る。
 */
const APP_JOB_PREFIX = 'APP_JOB_';
const APP_JOB_RESULT_PREFIX = 'APP_JOB_RESULT_';
const APP_JOB_QUEUE_EXPIRE_MS = 15 * 60 * 1000;   // 待ち行列に入ってからこれだけたっても始まらなければ、止まったとみなす
const APP_JOB_STALE_MS = 8 * 60 * 1000;   // 1 回の実行の上限（6 分）を超えて「実行中」のままなら止まったとみなす
const APP_JOB_KEEP_MS = 24 * 3600 * 1000;
// 続きの処理を、同じ実行の中で続けて動かしてよい最後の時刻（実行の始まりから）。1 回の実行の上限は 6 分
const APP_JOB_INLINE_LIMIT_MS = 5 * 60 * 1000;
const APP_JOB_STAGE_PROP = 'APP_STAGE_MS';   // 処理の段ごとの、最近かかった時間（続けて動かせるかの見積もり）。APP_JOB_ で始めない（処理の一覧に混ざる）

/** 処理の種類ごとの権限と中身 */
function appJobSpec_(kind) {
  const specs = {
    'MIGRATION.DRYRUN': { ownerOnly: true, label: '旧ブックの試し読み' },
    'MIGRATION.IMPORT': { ownerOnly: true, label: '旧ブックの取り込み' },
    // 予測の実行: 予算策定担当（その計画のクライアントの担当でもよい）以上
    // 組み立て（続きは同じ処理を続けて動かす）→ 計算 → 保存の 3 つ（Forecast.js）
    'FORECAST.RUN': { minRole: 'PLANNER', planScoped: true, label: '予測の実行' },
    'FORECAST.RUN_CALC': { minRole: 'PLANNER', planScoped: true, label: '予測の実行', internal: true },
    'FORECAST.RUN_SAVE': { minRole: 'PLANNER', planScoped: true, label: '予測の実行', internal: true },
    // 計画への保存・実行（Plan.js）: 操作ごとの役割（予算策定担当・承認者）。その計画のクライアントの担当でもよい
    'PLAN.EDIT': { minRoleOf: p => appPlanAction_(p && p.action, 'edit').minRole, planScoped: true, label: '保存' },
    'PLAN.RUN': { minRoleOf: p => appPlanAction_(p && p.action, 'run').minRole, planScoped: true, label: '実行' },
    'PLAN.RUN_CALC': { minRoleOf: p => appPlanAction_(p && p.action, 'run').minRole, planScoped: true, label: '実行', internal: true },
    'PLAN.RUN_SAVE': { minRoleOf: p => appPlanAction_(p && p.action, 'run').minRole, planScoped: true, label: '実行', internal: true },
    // 新しい計画を作る（旧ブックを経ずに。A-1 初期セットアップ → データ本体に保存）: 管理者
    'PLAN.CREATE': { minRole: 'ADMIN', label: '計画の作成' },
    // 全計画の情報源の信頼度から事前分布を作り、各計画の POOL_PRIOR に書く（旧来の C-1 の提案が使う）: 管理者
    'LEARN.POOL': { minRole: 'ADMIN', label: '学習の事前分布' },
    'PLAN.CREATE_SAVE': { minRole: 'ADMIN', label: '計画の作成', internal: true },
    // 途中で止まった保存の続きを、控え（Journal.js）のとおりに書く。予算策定担当以上（控えは、権限を確かめて始めた保存のもの）
    'SYSTEM.RECOVER': { minRole: 'PLANNER', label: '保存の続き' }
  };
  const s = specs[String(kind || '')];
  if (!s) throw new Error('未定義の処理です: ' + kind);
  return s;
}

/** apiStartJob の権限（種類ごと） */
function appJobStartOpts_(input) {
  const spec = appJobSpec_(input && input.kind);
  if (spec.internal) throw new Error('この処理は画面から始められません: ' + input.kind);
  if (spec.ownerOnly) return { ownerOnly: true, audit: false, detail: { kind: input.kind } };
  return { minRole: appJobMinRole_(spec, input.payload), clientId: spec.planScoped ? appJobPlanClient_(input.payload) : undefined, audit: false,
    detail: { kind: input.kind, planId: input.payload && input.payload.planId, action: input.payload && input.payload.action } };
}

/** 処理に要る役割（操作ごとに決まる処理は、頼んだ中身から） */
function appJobMinRole_(spec, payload) {
  return (spec.minRoleOf ? spec.minRoleOf(payload || {}) : spec.minRole) || 'ADMIN';
}

/**
 * 画面から渡された大きい中身（入力の行など）は、処理の記録（Script Properties、8KB まで）ではなく CacheService に置く。
 * 処理は数分以内に動くので、6 時間の保存期間で足りる
 */
function appJobStashArgs_(payload) {
  if (!payload || payload.args === undefined || appUtf8Bytes_(JSON.stringify(payload)) <= 4000) return payload;   // バイト数で見る（日本語は 1 文字 3 バイト）
  const ref = 'ARGS_' + Utilities.getUuid().replace(/-/g, '');
  appJobPutResult_(ref, payload.args);
  const out = Object.assign({}, payload, { argsRef: ref });
  delete out.args;
  return out;
}

function appJobArgs_(payload) {
  if (!payload.argsRef) return payload.args;
  const r = appJobGetResult_(payload.argsRef);
  if (!r.found) throw new Error('保存する内容が見つかりません（保存期間が過ぎた）。もう一度保存してください。');
  return r.value;
}

/** 計画ごとの処理の、その計画のクライアント（クライアント単位の役割で判定するため） */
function appJobPlanClient_(payload) {
  const planId = String(payload && payload.planId || '');
  if (!appIsSetUp_()) return '';
  const plan = appReadTable_('PLANS').filter(x => x.plan_id === planId)[0];
  if (!plan) throw new Error('計画が見つかりません。');
  return plan.client_id;
}

/** 処理を実行する（裏で。ctx は頼んだ人） */
function appJobExecute_(ctx, job) {
  const p = job.payload || {};
  switch (job.kind) {
    case 'MIGRATION.DRYRUN':
      return appAudited_(ctx, 'MIGRATION.DRYRUN', { entityType: 'PLAN', detail: { book: p.bookUrl, jobId: job.id },
        after: res => ({ client: res.client, fy: res.fy, sheets: res.sheets.length, lossless: res.lossless, faithful: res.faithful,
          contentHash: res.contentHash, timing: res.timing }) }, () => appMigrationDryRun_(ctx, p));
    case 'MIGRATION.IMPORT':
      return appAudited_(ctx, 'MIGRATION.IMPORT', { entityType: 'PLAN', detail: { book: p.bookUrl, contentHash: p.contentHash, jobId: job.id },
        after: res => ({ planId: res.planId, unchanged: res.unchanged, written: res.written, verified: res.verified, verify: res.verify }) },
        () => appMigrationImport_(ctx, p));
    case 'FORECAST.RUN':
      return appAudited_(ctx, 'FORECAST.RUN.BUILD', { entityType: 'PLAN', entityId: p.planId,
        detail: { planId: p.planId, confirms: p.confirms || [], runId: p.runId || '', built: p.build ? p.build.done.length : 0, jobId: job.id },
        after: res => ({ runId: res.__next.payload.runId, next: res.__next.kind, built: res.__next.payload.build.done.length, problems: res.__next.payload.build.problems,
          buildMs: res.__next.payload.buildMs }) }, () => appForecastRunBuild_(ctx, p, job));
    case 'FORECAST.RUN_CALC':
      return appAudited_(ctx, 'FORECAST.RUN.CALC', { entityType: 'PLAN', entityId: p.planId, detail: { planId: p.planId, runId: p.runId, confirms: p.confirms || [], jobId: job.id },
        after: res => (res.needConfirm ? { needConfirm: res.needConfirm.key || true } : { runId: res.__next.payload.runId,
          annual: res.__next.payload.headline && res.__next.payload.headline.annual, runMs: res.__next.payload.runMs }) }, () => appForecastRunCalc_(ctx, p));
    case 'FORECAST.RUN_SAVE':
      return appAudited_(ctx, 'FORECAST.RUN.SAVE', { entityType: 'FORECAST_RUN', entityId: p.runId, detail: { planId: p.planId, runId: p.runId, inputHash: p.inputHash, jobId: job.id },
        after: res => ({ runId: res.runId, changed: res.changed, written: res.written, annual: res.headline && res.headline.annual, timing: res.timing }) },
        () => appForecastRunSave_(ctx, p));
    case 'PLAN.EDIT':
      return appAudited_(ctx, 'PLAN.' + p.action, { entityType: 'PLAN', entityId: p.planId,
        detail: { planId: p.planId, action: p.action, args: appJobArgs_(p), jobId: job.id },
        after: res => ({ actionId: res.actionId, changed: res.changed, written: res.written, result: res.result, timing: res.timing }) },
        () => appPlanEdit_(ctx, p));
    case 'PLAN.RUN':
      return appAudited_(ctx, 'PLAN.' + p.action + '.BUILD', { entityType: 'PLAN', entityId: p.planId,
        detail: { planId: p.planId, action: p.action, actionId: p.actionId || '', built: p.build ? p.build.done.length : 0, jobId: job.id },
        after: res => ({ actionId: res.__next.payload.actionId, next: res.__next.kind, built: res.__next.payload.build.done.length, problems: res.__next.payload.build.problems,
          buildMs: res.__next.payload.buildMs }) }, () => appPlanRunBuild_(ctx, p));
    case 'PLAN.RUN_CALC':
      return appAudited_(ctx, 'PLAN.' + p.action + '.CALC', { entityType: 'PLAN', entityId: p.planId, detail: { planId: p.planId, action: p.action, actionId: p.actionId, jobId: job.id },
        after: res => ({ actionId: res.__next.payload.actionId, next: res.__next.kind, result: res.__next.payload.result, runMs: res.__next.payload.runMs,
          aiCalls: res.__next.payload.aiCalls, aiAttempt: res.__next.payload.aiAttempt }) },
        () => appPlanRunCalc_(ctx, p));
    case 'PLAN.RUN_SAVE':
      return appAudited_(ctx, 'PLAN.' + p.action + '.SAVE', { entityType: 'PLAN_ACTION', entityId: p.actionId,
        detail: { planId: p.planId, action: p.action, actionId: p.actionId, inputHash: p.inputHash, jobId: job.id },
        after: res => ({ actionId: res.actionId, changed: res.changed, written: res.written, timing: res.timing }) }, () => appPlanRunSave_(ctx, p));
    case 'LEARN.POOL':
      return appAudited_(ctx, 'LEARN.POOL', { entityType: 'SYSTEM', detail: { jobId: job.id, remaining: p.remaining ? p.remaining.length : null },
        after: res => (res.__next ? { next: res.__next.kind, done: res.__next.payload.done.length, remaining: res.__next.payload.remaining.length } : { written: res.written, plans: res.plans }) },
        () => appPoolApply_(ctx, p));
    case 'PLAN.CREATE':
      return appAudited_(ctx, 'PLAN.CREATE.BUILD', { entityType: 'PLAN', detail: { clientName: p.clientName, fy: p.fy, peopleCsv: p.peopleCsv, jobId: job.id },
        after: res => ({ next: res.__next.kind, buildMs: res.__next.payload.buildMs }) }, () => appPlanCreateBuild_(ctx, p));
    case 'PLAN.CREATE_SAVE':
      return appAudited_(ctx, 'PLAN.CREATE.SAVE', { entityType: 'PLAN', detail: { clientName: p.clientName, fy: p.fy, jobId: job.id },
        after: res => ({ planId: res.planId, sheets: res.sheets, written: res.written, timing: res.timing }) }, () => appPlanCreateSave_(ctx, p));
    case 'SYSTEM.RECOVER':
      return appAudited_(ctx, 'SYSTEM.RECOVER', { entityType: 'SYSTEM', detail: { jobId: job.id, pending: appJournalPending_() },
        after: res => res }, () => appWithLock_(() => appJournalRecover_(ctx) || { nothing: true }));
    default:
      throw new Error('未定義の処理です: ' + job.kind);
  }
}

// ---- 保存（Script Properties と CacheService） ----

function appJobGet_(id) {
  const raw = appProps_().getProperty(APP_JOB_PREFIX + id);
  return raw ? JSON.parse(raw) : null;
}

function appUtf8Bytes_(s) { return encodeURIComponent(s).replace(/%[0-9A-F]{2}/g, 'x').length; }

function appJobSave_(job) {
  appProps_().setProperty(APP_JOB_PREFIX + job.id, JSON.stringify(job));
}

function appJobList_() {
  return appProps_().getKeys().filter(k => k.indexOf(APP_JOB_PREFIX) === 0 && k.indexOf(APP_JOB_RESULT_PREFIX) !== 0)
    .map(k => { try { return JSON.parse(appProps_().getProperty(k)); } catch (e) { return null; } })
    .filter(Boolean)
    .sort((a, b) => (a.createdAt < b.createdAt ? -1 : 1));
}

/** 結果を置く（1 つ 100KB までなので分けて置く。上限はバイト数なので、日本語（1 文字 3 バイト）でも超えない 30,000 文字ずつ） */
function appJobPutResult_(id, result) {
  const text = JSON.stringify(result === undefined ? null : result);
  const CHUNK = 30000;
  const cache = CacheService.getScriptCache();
  const n = Math.max(1, Math.ceil(text.length / CHUNK));
  for (let i = 0; i < n; i++) cache.put(APP_JOB_RESULT_PREFIX + id + '_' + i, text.slice(i * CHUNK, (i + 1) * CHUNK), 21600);
  cache.put(APP_JOB_RESULT_PREFIX + id, String(n), 21600);
}

function appJobGetResult_(id) {
  const cache = CacheService.getScriptCache();
  const n = Number(cache.get(APP_JOB_RESULT_PREFIX + id) || 0);
  if (!n) return { found: false };
  let text = '';
  for (let i = 0; i < n; i++) {
    const part = cache.get(APP_JOB_RESULT_PREFIX + id + '_' + i);
    if (part === null) return { found: false };
    text += part;
  }
  return { found: true, value: JSON.parse(text) };
}

/** 1 日より前の処理の記録と、待っている処理のないトリガーを消す */
function appCleanupJobs_() {
  const now = new Date().getTime();
  const jobs = appJobList_();
  jobs.forEach(j => {
    if (now - new Date(j.createdAt).getTime() > APP_JOB_KEEP_MS) appProps_().deleteProperty(APP_JOB_PREFIX + j.id);
  });
  const waiting = jobs.filter(j => j.status === 'QUEUED').map(j => j.triggerUid);
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'triggerRunJob' && waiting.indexOf(String(t.getUniqueId())) < 0) ScriptApp.deleteTrigger(t);
  });
}

/** トリガーの UID が、待っている処理のために作ったものか（トリガーの一覧から消えた後でも確かめられる） */
function appIsJobTrigger_(uid) {
  return !!uid && appJobList_().some(j => j.status === 'QUEUED' && j.triggerUid === String(uid));
}

/** その人の、まだ終わっていない処理（画面を開き直したときに続きから待つため） */
function appActiveJobOf_(email) {
  const j = appJobList_().filter(x => x.requestedBy === email && (x.status === 'QUEUED' || (x.status === 'RUNNING' && !appJobIsStale_(x)))).pop();
  return j ? { jobId: j.id, kind: j.kind, status: j.status, payload: appJobClientPayload_(j.payload) } : null;
}

// ---- 始める・動かす・状態を返す ----

/** 画面から渡せる項目（ほかは裏の処理どうしの受け渡し専用。画面から渡されたら捨てる） */
const APP_JOB_CLIENT_FIELDS = ['planId', 'confirms', 'action', 'args', 'inputHash', 'bookUrl', 'contentHash', 'clientName', 'fy', 'peopleCsv'];

function appJobClientPayload_(payload) {
  const out = {};
  APP_JOB_CLIENT_FIELDS.forEach(k => { if (payload && payload[k] !== undefined) out[k] = payload[k]; });
  return out;
}

function appStartJob_(ctx, input) {
  const kind = String(input && input.kind || '');
  appJobSpec_(kind);
  if (!appIsSetUp_()) throw new Error('初期設定がまだです。');
  return appWithLock_(() => {
    appCleanupJobs_();
    appExpireQueuedJobs_();
    const busy = appJobList_().filter(j => j.status === 'QUEUED' || (j.status === 'RUNNING' && !appJobIsStale_(j)));
    if (busy.length) throw new Error('ほかの処理（' + appJobSpec_(busy[0].kind).label + '）が終わるまでお待ちください。');
    const job = appEnqueueJob_(kind, appJobStashArgs_(appJobClientPayload_(input && input.payload)), ctx.actor, '');
    return { jobId: job.id, status: job.status };
  });
}

/** 処理を待ち行列に入れ、1 回だけ動くトリガーを作る（ロックの中で呼ぶ） */
function appEnqueueJob_(kind, payload, requestedBy, parentId) {
  const job = { id: appId_('JOB'), kind: kind, payload: payload || {}, status: 'QUEUED', requestedBy: requestedBy, parentId: parentId || '',
    createdAt: new Date().toISOString(), startedAt: '', finishedAt: '', error: '', attempts: 0, nextJobId: '' };
  if (appUtf8Bytes_(JSON.stringify(job)) > 8000) throw new Error('処理に渡す内容が大きすぎます。');   // 1 つの値の上限は約 9KB（文字数ではなくバイト数）
  const trigger = ScriptApp.newTrigger('triggerRunJob').timeBased().after(1000).create();
  job.triggerUid = String(trigger.getUniqueId());
  appJobSave_(job);
  return job;
}

/** 15 分たっても始まらなかった処理は「失敗」にする（トリガーが動かなかった。残すと次の処理を始められない。ロックの中で呼ぶ） */
function appExpireQueuedJobs_() {
  const now = new Date().getTime();
  appJobList_().filter(j => j.status === 'QUEUED' && now - new Date(j.createdAt).getTime() > APP_JOB_QUEUE_EXPIRE_MS).forEach(j => {
    ScriptApp.getProjectTriggers().forEach(t => { if (String(t.getUniqueId()) === j.triggerUid) ScriptApp.deleteTrigger(t); });
    j.status = 'FAILED';
    j.error = '処理が始まりませんでした。';
    j.finishedAt = new Date().toISOString();
    appJobSave_(j);
  });
}

function appJobIsStale_(job) {
  return job.status === 'RUNNING' && new Date().getTime() - new Date(job.startedAt).getTime() > APP_JOB_STALE_MS;
}

/** 頼んだ人として動かすための ctx（裏の実行は所有者のトリガーで動くので、頼んだ人の権限を確かめ直す） */
function appJobContext_(job) {
  const base = appCurrentUser_();
  const email = String(job.requestedBy || '').toLowerCase();
  const user = { email: email, owner: base.owner, domain: base.domain, isOwner: !!email && email === base.owner,
    isInternal: !!email && !!base.domain && email.slice(-(base.domain.length + 1)) === '@' + base.domain };
  const roles = appRolesOf_(user);
  return { user: user, roles: roles, actor: email || '(unknown)', requestId: Utilities.getUuid() };
}

/** トリガーから: 自分のトリガーを消し、待っている処理を古い順に 1 つずつ動かす */
function appRunJobs_(ctx, e) {
  const uid = e && e.triggerUid ? String(e.triggerUid) : '';
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'triggerRunJob' && uid && String(t.getUniqueId()) === uid) ScriptApp.deleteTrigger(t);
  });
  const start = new Date().getTime();
  const ran = [];
  appJobList_().filter(j => j.status === 'QUEUED').forEach(j => {
    let r = appRunJob_(j.id);
    ran.push(r);
    // 続きの処理は、これまでにかかった時間から見て収まるなら、次のトリガーを待たずにこの実行の中で動かす（トリガーを待つだけで数十秒かかる）
    while (r.nextJobId && appJobCanInline_(r.nextJobId, start)) {
      const next = appJobGet_(r.nextJobId);
      r = appRunJob_(r.nextJobId);
      ran.push(Object.assign({ inline: true }, r));
      if (next && next.triggerUid) ScriptApp.getProjectTriggers().forEach(t => { if (String(t.getUniqueId()) === next.triggerUid) ScriptApp.deleteTrigger(t); });
    }
  });
  return { ran: ran };
}

/** 段の名前（予測・実行は操作ごとに時間が違うので、操作も付ける） */
function appJobStageKey_(job) {
  return job.kind + (job.payload && job.payload.action ? ':' + job.payload.action : '');
}

function appJobStageTimes_() {
  try { return JSON.parse(appProps_().getProperty(APP_JOB_STAGE_PROP) || '{}'); } catch (e) { return {}; }
}

/** かかった時間を残す（段ごとに最近 3 回） */
function appJobRecordStage_(job, ms) {
  const t = appJobStageTimes_();
  const k = appJobStageKey_(job);
  t[k] = [Math.round(ms)].concat(t[k] || []).slice(0, 3);
  const keys = Object.keys(t);
  if (keys.length > 60) delete t[keys[0]];
  appProps_().setProperty(APP_JOB_STAGE_PROP, JSON.stringify(t));
}

/** 次の処理を同じ実行の中で動かしてよいか: 最近 3 回のうち一番長い時間の 1.5 倍 + 30 秒が、残りの時間に収まる（初めての段は動かさない） */
function appJobCanInline_(nextId, startMs) {
  const next = appJobGet_(nextId);
  if (!next || next.status !== 'QUEUED') return false;
  const past = appJobStageTimes_()[appJobStageKey_(next)];
  if (!past || !past.length) return false;
  const need = Math.max.apply(null, past) * 1.5 + 30 * 1000;
  return new Date().getTime() - startMs + need <= APP_JOB_INLINE_LIMIT_MS;
}

function appRunJob_(id) {
  const job = appWithLock_(() => {
    const cur = appJobGet_(id);
    if (!cur || cur.status !== 'QUEUED') return null;   // ほかの実行が先に取った
    cur.status = 'RUNNING';
    cur.startedAt = new Date().toISOString();
    cur.attempts = Number(cur.attempts || 0) + 1;
    appJobSave_(cur);
    return cur;
  });
  if (!job) return { id: id, skipped: true };
  const t0 = new Date().getTime();
  let status = 'DONE';
  let error = '';
  let nextJobId = '';
  try {
    APP_STORE_CACHE_ = {};
    const ctx = appJobContext_(job);
    const spec = appJobSpec_(job.kind);
    const clientId = spec.planScoped ? appJobPlanClient_(job.payload) : undefined;
    const allowed = spec.ownerOnly ? ctx.user.isOwner : appHasRole_(ctx.roles, appJobMinRole_(spec, job.payload), clientId);
    if (!allowed) throw new Error('この操作をする権限がありません。');
    const res = appJobExecute_(ctx, job);
    if (res && res.__next) {
      // 続きの処理: 次の処理を待ち行列に入れ、この処理の結果は「続きあり」にする
      // 次の処理を入れるのと同じロックの中で、この処理を「続きあり」で終わりにする（この後で止まっても、画面は次の処理を待てる）
      const next = appWithLock_(() => {
        const n = appEnqueueJob_(res.__next.kind, res.__next.payload, job.requestedBy, job.id);
        const cur = appJobGet_(id) || job;
        appJobSave_(Object.assign(cur, { status: 'DONE', nextJobId: n.id, finishedAt: new Date().toISOString(), payload: appJobClientPayload_(cur.payload) }));
        return n;
      });
      nextJobId = next.id;
      appJobPutResult_(id, { continued: true, nextJobId: next.id, nextKind: next.kind });
    } else {
      appJobPutResult_(id, appSerialize_(res));
    }
  } catch (err) {
    status = 'FAILED';
    error = String(err && err.message ? err.message : err).slice(0, 1000);
    appLogError_(job.kind, err, { requestId: id, actor: job.requestedBy });
  }
  const done = appJobGet_(id) || job;
  done.status = status;
  done.error = error;
  done.nextJobId = nextJobId;
  done.finishedAt = new Date().toISOString();
  done.payload = appJobClientPayload_(done.payload);   // 終わった処理は、画面に要る項目だけ残す（Script Properties の全体の上限 500KB）
  appJobSave_(done);
  appRunLog_({ requestId: id, kind: 'JOB:' + job.kind, startedAt: job.startedAt, durationMs: new Date().getTime() - t0, status: status,
    detail: { jobId: id, requestedBy: job.requestedBy, attempts: done.attempts }, error: error });
  try { if (status === 'DONE') appJobRecordStage_(job, new Date().getTime() - t0); } catch (e) { /* 見積もりの記録は処理の結果に関係しない */ }
  return { id: id, status: status, nextJobId: nextJobId };
}

/** 画面から: 処理の状態（終わっていれば結果も）。頼んだ人と所有者だけが見られる */
function appJobStatus_(ctx, id) {
  let job = appJobGet_(String(id || ''));
  if (!job) throw new Error('処理が見つかりません。もう一度始めてください。');
  if (job.requestedBy !== ctx.actor && !ctx.user.isOwner) throw new Error('この処理の状態を見る権限がありません。');
  // 終わった続きの処理はたどって、いま動いている（または最後の）処理の状態を返す（画面が段ごとに尋ね直さなくてよい）
  for (let hop = 0; hop < 100 && job.status === 'DONE' && job.nextJobId; hop++) {
    const next = appJobGet_(job.nextJobId);
    if (!next || next.requestedBy !== job.requestedBy) break;
    job = next;
  }
  const out = { jobId: job.id, kind: job.kind, status: job.status, createdAt: job.createdAt, startedAt: job.startedAt, finishedAt: job.finishedAt,
    error: job.error, payload: appJobClientPayload_(job.payload) };
  if (job.status === 'QUEUED' && new Date().getTime() - new Date(job.createdAt).getTime() > APP_JOB_QUEUE_EXPIRE_MS) {
    out.status = 'STALLED';
    out.error = '処理が始まりませんでした。もう一度始めてください。';
  } else if (appJobIsStale_(job)) {
    out.status = 'STALLED';
    out.error = '処理が時間内に終わりませんでした（1 回の実行は 6 分まで）。';
  } else if (job.status === 'DONE' && job.nextJobId) {
    const next = appJobGet_(job.nextJobId);
    out.status = 'CONTINUED';
    out.nextJobId = job.nextJobId;
    out.nextKind = next ? next.kind : '';
  } else if (job.status === 'DONE') {
    const r = appJobGetResult_(job.id);
    if (r.found) out.result = r.value; else { out.status = 'LOST'; out.error = '結果の保存期間が過ぎました。もう一度始めてください。'; }
  }
  return out;
}
