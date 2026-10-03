/**
 * Api.js — 画面から呼ぶ入口。公開する関数はこのファイルの doGet・api*・trigger* だけ。
 * どれも api_ を通し、本人確認 → 役割の判定 → （書き込みなら）記録つきで実行 → エラーの記録 を必ず行う。
 * 新しい入口を足したら app/tests/app-contract.test.mjs の一覧にも足す（足さないとテストが止める）。
 */
function doGet(e) {
  const t = HtmlService.createTemplateFromFile('UI');
  t.charsJs = APP_UI_CHARS_JS;
  t.bootJson = JSON.stringify(api_('APP.OPEN', { minRole: 'VIEWER', audit: false, allowAnonymousView: true }, ctx => appBootstrap_(ctx)))
    .replace(/</g, '\\u003c');
  const out = t.evaluate().setTitle(APP_NAME).addMetaTag('viewport', 'width=device-width, initial-scale=1');
  try { out.setFaviconUrl(APP_FAVICON_URL); } catch (err) { Logger.log('favicon: ' + (err && err.message ? err.message : err)); }
  return out;
}

/** 画面の初期データ。役割に応じて出す情報を変える */
function apiBootstrap() {
  return api_('APP.BOOTSTRAP', { minRole: 'VIEWER', audit: false, allowAnonymousView: true }, ctx => appBootstrap_(ctx));
}

function apiSetup() {
  return api_('SETUP.INIT', { ownerOnly: true, audit: false }, ctx => appSetup_(ctx));
}

/**
 * 所有者が Apps Script のエディタから 1 回だけ実行する: A-4 AI 調査の許可（外への問い合わせ・Google Cloud）を出す。
 * 裏の処理は所有者の許可で動くので、許可が無いと Vertex AI に問い合わせられない。データは何も変えない（問い合わせもしない）
 */
function apiAuthorizeAi() {
  return api_('SETUP.AUTHORIZE_AI', { ownerOnly: true, audit: false }, ctx => appAuthorizeAi_(ctx));
}

function apiListDirectory() {
  return api_('DIRECTORY.LIST', { minRole: 'ADMIN', audit: false }, ctx => appListDirectory_(ctx));
}

function apiSaveMember(input) {
  return api_('MEMBER.SAVE', { minRole: 'ADMIN', entityType: 'MEMBER', entityId: input && input.email, detail: input,
    after: res => res.member }, ctx => appSaveMember_(ctx, input));
}

function apiGrantRole(input) {
  return api_('ROLE.GRANT', { minRole: 'ADMIN', entityType: 'ROLE', detail: input, after: res => res.role },
    ctx => appGrantRole_(ctx, input));
}

function apiRevokeRole(input) {
  return api_('ROLE.REVOKE', { minRole: 'ADMIN', entityType: 'ROLE', entityId: input && input.roleId, detail: input,
    after: res => res.role }, ctx => appRevokeRole_(ctx, input));
}

function apiSaveClient(input) {
  return api_('CLIENT.SAVE', { minRole: 'ADMIN', entityType: 'CLIENT', entityId: input && input.clientId, detail: input,
    after: res => res.client }, ctx => appSaveClient_(ctx, input));
}

function apiListSettings() {
  return api_('SETTINGS.LIST', { minRole: 'ADMIN', audit: false }, () => ({ settings: appSettingsCurrent_().map(x => {
    if (x.type !== 'sheet' || !x.value) return x;
    let name = '';
    try { name = DriveApp.getFileById(x.value).getName(); } catch (e) { name = '（開けません）'; }
    return Object.assign({}, x, { display: name, url: 'https://docs.google.com/spreadsheets/d/' + x.value + '/edit' });
  }) }));
}

function apiSaveSetting(input) {
  return api_('SETTING.SAVE', { minRole: 'ADMIN', entityType: 'SETTING', entityId: input && input.key, detail: input,
    after: res => res.saved }, ctx => appSaveSetting_(ctx, input));
}

function apiListAudit(input) {
  return api_('AUDIT.LIST', { minRole: 'ADMIN', audit: false },
    () => ({ rows: appRecentAudit_(Math.min(500, Number(input && input.limit) || 200), input && input.query) }));
}

function apiHealth() {
  return api_('HEALTH.CHECK', { minRole: 'ADMIN', audit: false }, () => appHealth_());
}

function apiEnableBackup() {
  return api_('BACKUP.ENABLE', { minRole: 'ADMIN', entityType: 'SYSTEM', after: res => res }, () => appEnableBackup_());
}

function apiRunBackup() {
  return api_('BACKUP.RUN', { minRole: 'ADMIN', entityType: 'SYSTEM', after: res => res }, ctx => appBackup_(ctx));
}

/** 計画の一覧（取り込み・計算の一致の確認の対象） */
function apiListPlans() {
  return api_('PLANS.LIST', { minRole: 'VIEWER', audit: false }, () => ({ plans: appListPlans_() }));
}

/** 計画の最新の予測（閲覧は社内全員） */
function apiForecastLatest(input) {
  return api_('FORECAST.LATEST', { minRole: 'VIEWER', audit: false }, () => appForecastLatest_(input && input.planId));
}

/** 計画の画面（旧来の Web アプリと同じ中身: 入力・予測と予算・検証・四半期レビュー・進み。閲覧は社内全員） */
function apiPlanView(input) {
  return api_('PLAN.VIEW', { minRole: 'VIEWER', audit: false }, ctx => appPlanView_(ctx, input && input.planId));
}

/** 時間のかかる処理を始める（裏で動かす。種類ごとに権限が違う: 旧ブックの試し読み・取り込みは所有者だけ） */
function apiStartJob(input) {
  return api_('JOB.START', appJobStartOpts_(input), ctx => appStartJob_(ctx, input));
}

/** 処理の状態（終わっていれば結果）。頼んだ人と所有者だけ */
function apiJobStatus(input) {
  return api_('JOB.STATUS', { minRole: 'VIEWER', audit: false }, ctx => appJobStatus_(ctx, input && input.jobId));
}

/** 処理を裏で動かす（apiStartJob が作った 1 回だけのトリガーから） */
function triggerRunJob(e) {
  return api_('JOB.RUN', { minRole: 'ADMIN', audit: false, trigger: { event: e, handler: 'triggerRunJob', alsoAccept: appIsJobTrigger_ } },
    ctx => appRunJobs_(ctx, e));
}

/**
 * 毎日のバックアップ（時間主導のトリガーから）。トリガーは所有者として動く。
 * トリガーの中では操作者のメールが空になることがあるので、そのときはイベントの triggerUid が
 * このプロジェクトのバックアップ用トリガーと一致する場合に限り、所有者として扱う（ブラウザからは UID が分からず偽装できない）。
 */
function triggerDailyBackup(e) {
  return api_('BACKUP.DAILY', { minRole: 'ADMIN', entityType: 'SYSTEM', trigger: { event: e, handler: 'triggerDailyBackup' },
    after: res => res }, ctx => appBackup_(ctx));
}

/**
 * すべての入口の共通処理。opts:
 *   minRole / clientId … 必要な役割（clientId を渡すとクライアント単位の役割も数える）
 *   ownerOnly          … 所有者だけ（初期設定）
 *   audit              … false なら記録しない（読むだけの入口）。既定は記録する
 *   allowAnonymousView … 役割が足りないときも画面の初期データだけは返す（「権限がありません」を画面で出すため）
 *   entityType / entityId / detail / reason / before() / after(res) … 監査の項目
 */
function api_(action, opts, fn) {
  opts = opts || {};
  APP_STORE_CACHE_ = {};
  let user = appCurrentUser_();
  const trigUid = opts.trigger && opts.trigger.event && opts.trigger.event.triggerUid ? String(opts.trigger.event.triggerUid) : '';
  if (opts.trigger && !user.email && (appIsOwnTrigger_(opts.trigger.event, opts.trigger.handler) ||
      (opts.trigger.alsoAccept && trigUid && appIsSetUp_() && opts.trigger.alsoAccept(trigUid)))) {
    user = Object.assign({}, user, { email: user.owner, isOwner: !!user.owner });
  }
  const roles = appRolesOf_(user);
  const ctx = { user: user, roles: roles, actor: user.email || '(unknown)', requestId: Utilities.getUuid() };
  if (roles.length) appAutoEnsureTables_(ctx);   // 社内の人の最初の操作で、版を上げて足した表を作る
  const allowed = opts.ownerOnly ? user.isOwner : appHasRole_(roles, opts.minRole || 'ADMIN', opts.clientId);
  try {
    if (!allowed) {
      if (opts.allowAnonymousView) return appSerialize_(appDeniedBootstrap_(ctx));
      if (appIsSetUp_()) {
        try {
          appAuditAppend_({ actor: ctx.actor, actorRoles: appRoleSummary_(roles), action: action, phase: 'DENIED', result: 'DENIED',
            entityType: opts.entityType, entityId: opts.entityId, clientId: opts.clientId, detail: opts.detail, requestId: ctx.requestId });
        } catch (e) { Logger.log('api_ DENIED の記録に失敗: ' + (e && e.message ? e.message : e)); }
      }
      throw new Error('この操作をする権限がありません。');
    }
    if (opts.audit === false) return appSerialize_(fn(ctx));
    if (!appIsSetUp_()) throw new Error('初期設定がまだです。所有者が「初期設定」を実行してください。');
    return appSerialize_(appAudited_(ctx, action, opts, () => fn(ctx)));
  } catch (err) {
    appLogError_(action, err, ctx);
    throw new Error(String(err && err.message ? err.message : err));
  }
}

/** e.triggerUid が、このプロジェクトの handler 用のトリガーのものか */
function appIsOwnTrigger_(e, handler) {
  const uid = e && e.triggerUid ? String(e.triggerUid) : '';
  if (!uid) return false;
  try {
    return ScriptApp.getProjectTriggers().some(t => t.getUniqueId() === uid && t.getHandlerFunction() === handler);
  } catch (err) {
    return false;
  }
}

/** google.script.run で返せる形にする（Date などを文字列にする） */
function appSerialize_(v) {
  return v === undefined ? null : JSON.parse(JSON.stringify(v));
}

function appBootstrap_(ctx) {
  const isAdmin = appHasRole_(ctx.roles, 'ADMIN');
  return {
    app: { name: APP_NAME, nameJa: APP_NAME_JA, version: APP_VERSION },
    user: { email: ctx.user.email, isOwner: ctx.user.isOwner, isAdmin: isAdmin,
      roles: ctx.roles.map(r => ({ role: r.role, label: APP_ROLE_LABELS[r.role], scopeType: r.scope_type, clientId: r.client_id })) },
    setUp: appIsSetUp_(),
    allowed: true,
    activeJob: (() => { try { return ctx.user.email && appIsSetUp_() ? appActiveJobOf_(ctx.user.email) : null; } catch (e) { return null; } })()
  };
}

function appDeniedBootstrap_(ctx) {
  return { app: { name: APP_NAME, nameJa: APP_NAME_JA, version: APP_VERSION }, user: { email: ctx.user.email, isOwner: false, isAdmin: false, roles: [] },
    setUp: appIsSetUp_(), allowed: false };
}
