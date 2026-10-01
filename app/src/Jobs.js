/**
 * Jobs.js — 時間のかかる処理（旧ブックの試し読み・取り込み、今後の予測の実行など）を、画面の通信から切り離して動かす。
 * 画面から始めると、1 回だけ動く時間主導のトリガーを作り、裏で最後まで実行する。画面は数秒ごとに状態を尋ねて結果を受け取る。
 * 画面と 1 分以上つながったままにすると通信が切れる（HTTP 503）ことがあるため（2026-10-01 に旧ブックの試し読みで発生）。
 * 状態は Script Properties（APP_JOB_<id>、小さい）、結果は CacheService（6 時間）に置く。処理そのものの記録は監査ログに残る。
 */
const APP_JOB_PREFIX = 'APP_JOB_';
const APP_JOB_RESULT_PREFIX = 'APP_JOB_RESULT_';
const APP_JOB_STALE_MS = 8 * 60 * 1000;   // 1 回の実行の上限（6 分）を超えて「実行中」のままなら止まったとみなす
const APP_JOB_KEEP_MS = 24 * 3600 * 1000;

/** 処理の種類ごとの権限と中身 */
function appJobSpec_(kind) {
  const specs = {
    'MIGRATION.DRYRUN': { ownerOnly: true, label: '旧ブックの試し読み' },
    'MIGRATION.IMPORT': { ownerOnly: true, label: '旧ブックの取り込み' }
  };
  const s = specs[String(kind || '')];
  if (!s) throw new Error('未定義の処理です: ' + kind);
  return s;
}

/** apiStartJob の権限（種類ごと） */
function appJobStartOpts_(input) {
  const spec = appJobSpec_(input && input.kind);
  return spec.ownerOnly ? { ownerOnly: true, audit: false, detail: { kind: input.kind } } : { minRole: spec.minRole || 'ADMIN', audit: false, detail: { kind: input.kind } };
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
    default:
      throw new Error('未定義の処理です: ' + job.kind);
  }
}

// ---- 保存（Script Properties と CacheService） ----

function appJobGet_(id) {
  const raw = appProps_().getProperty(APP_JOB_PREFIX + id);
  return raw ? JSON.parse(raw) : null;
}

function appJobSave_(job) {
  appProps_().setProperty(APP_JOB_PREFIX + job.id, JSON.stringify(job));
}

function appJobList_() {
  return appProps_().getKeys().filter(k => k.indexOf(APP_JOB_PREFIX) === 0 && k.indexOf(APP_JOB_RESULT_PREFIX) !== 0)
    .map(k => { try { return JSON.parse(appProps_().getProperty(k)); } catch (e) { return null; } })
    .filter(Boolean)
    .sort((a, b) => (a.createdAt < b.createdAt ? -1 : 1));
}

/** 結果を置く（1 つ 100KB までなので分けて置く） */
function appJobPutResult_(id, result) {
  const text = JSON.stringify(result === undefined ? null : result);
  const CHUNK = 90000;
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
  return j ? { jobId: j.id, kind: j.kind, status: j.status, payload: j.payload } : null;
}

// ---- 始める・動かす・状態を返す ----

function appStartJob_(ctx, input) {
  const kind = String(input && input.kind || '');
  appJobSpec_(kind);
  if (!appIsSetUp_()) throw new Error('初期設定がまだです。');
  return appWithLock_(() => {
    appCleanupJobs_();
    const busy = appJobList_().filter(j => j.status === 'QUEUED' || (j.status === 'RUNNING' && !appJobIsStale_(j)));
    if (busy.length) throw new Error('ほかの処理（' + appJobSpec_(busy[0].kind).label + '）が終わるまでお待ちください。');
    const job = { id: appId_('JOB'), kind: kind, payload: (input && input.payload) || {}, status: 'QUEUED', requestedBy: ctx.actor,
      createdAt: new Date().toISOString(), startedAt: '', finishedAt: '', error: '', attempts: 0 };
    if (JSON.stringify(job).length > 8000) throw new Error('処理に渡す内容が大きすぎます。');
    const trigger = ScriptApp.newTrigger('triggerRunJob').timeBased().after(1000).create();
    job.triggerUid = String(trigger.getUniqueId());
    appJobSave_(job);
    return { jobId: job.id, status: job.status };
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
  const ran = [];
  appJobList_().filter(j => j.status === 'QUEUED').forEach(j => { ran.push(appRunJob_(j.id)); });
  return { ran: ran };
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
  try {
    APP_STORE_CACHE_ = {};
    const ctx = appJobContext_(job);
    const spec = appJobSpec_(job.kind);
    const allowed = spec.ownerOnly ? ctx.user.isOwner : appHasRole_(ctx.roles, spec.minRole || 'ADMIN');
    if (!allowed) throw new Error('この操作をする権限がありません。');
    appJobPutResult_(id, appSerialize_(appJobExecute_(ctx, job)));
  } catch (err) {
    status = 'FAILED';
    error = String(err && err.message ? err.message : err).slice(0, 1000);
    appLogError_(job.kind, err, { requestId: id, actor: job.requestedBy });
  }
  const done = appJobGet_(id) || job;
  done.status = status;
  done.error = error;
  done.finishedAt = new Date().toISOString();
  appJobSave_(done);
  appRunLog_({ requestId: id, kind: 'JOB:' + job.kind, startedAt: job.startedAt, durationMs: new Date().getTime() - t0, status: status,
    detail: { jobId: id, requestedBy: job.requestedBy, attempts: done.attempts }, error: error });
  return { id: id, status: status };
}

/** 画面から: 処理の状態（終わっていれば結果も）。頼んだ人と所有者だけが見られる */
function appJobStatus_(ctx, id) {
  const job = appJobGet_(String(id || ''));
  if (!job) throw new Error('処理が見つかりません。もう一度始めてください。');
  if (job.requestedBy !== ctx.actor && !ctx.user.isOwner) throw new Error('この処理の状態を見る権限がありません。');
  const out = { jobId: job.id, kind: job.kind, status: job.status, createdAt: job.createdAt, startedAt: job.startedAt, finishedAt: job.finishedAt,
    error: job.error, payload: job.payload };
  if (job.status === 'QUEUED' && new Date().getTime() - new Date(job.createdAt).getTime() > 15 * 60 * 1000) {
    out.status = 'STALLED';
    out.error = '処理が始まりませんでした。もう一度始めてください。';
  } else if (appJobIsStale_(job)) {
    out.status = 'STALLED';
    out.error = '処理が時間内に終わりませんでした（1 回の実行は 6 分まで）。';
  } else if (job.status === 'DONE') {
    const r = appJobGetResult_(job.id);
    if (r.found) out.result = r.value; else { out.status = 'LOST'; out.error = '結果の保存期間が過ぎました。もう一度始めてください。'; }
  }
  return out;
}
