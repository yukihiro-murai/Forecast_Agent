/**
 * Api.js — 画面から呼ぶ入口。公開する関数はこのファイルの doGet・api*・trigger* だけ。
 * どれも api_ を通し、本人確認 → 役割の判定 → （書き込みなら）記録つきで実行 → エラーの記録 を必ず行う。
 * 新しい入口を足したら app/tests/app-contract.test.mjs の一覧にも足す（足さないとテストが止める）。
 */
function doGet(e) {
  const t = HtmlService.createTemplateFromFile('UI');
  t.charsJs = APP_UI_CHARS_JS;
  // 所有者が開いたときは、自動の AI 調査のトリガーを確かめる（1 日 1 回だけ。AutoResearch.js）
  t.bootJson = JSON.stringify(api_('APP.OPEN', { minRole: 'VIEWER', audit: false, allowAnonymousView: true }, ctx => { appAutoResearchOnOpen_(ctx); return appBootstrap_(ctx); }))
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

/**
 * 所有者が Apps Script のエディタから管理の操作をする（管理の画面は画面から外す。2026-10-06 村井さん決定）。
 * エディタからは引数を渡せないので、頼む中身はスクリプト プロパティ OWNER_TASK に JSON で置く。結果は OWNER_TASK_RESULT と実行ログ（OwnerTask.js）。
 * 操作ごとに、画面の入口と同じ役割の判定・同じ操作の名前で監査に残すので、この入口そのものは記録しない（断ったことは記録する）
 */
function apiOwnerTask() {
  return api_('OWNER.TASK', { ownerOnly: true, audit: false }, ctx => appOwnerTask_(ctx));
}

function apiListDirectory() {
  return api_('DIRECTORY.LIST', { minRole: 'ADMIN', audit: false }, ctx => appCachedRead_('DIRECTORY', () => appListDirectory_(ctx)));
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

/** 画面に出すクライアントの名前を決める（管理者） */
function apiSaveClientName(input) {
  return api_('CLIENT.NAME', { minRole: 'ADMIN', entityType: 'CLIENT', entityId: input && input.clientId, detail: input,
    after: res => ({ displayName: res.displayName, auto: res.auto }) }, ctx => appSaveClientName_(ctx, input));
}

function apiSaveClient(input) {
  return api_('CLIENT.SAVE', { minRole: 'ADMIN', entityType: 'CLIENT', entityId: input && input.clientId, detail: input,
    after: res => res.client }, ctx => appSaveClient_(ctx, input));
}

function apiListSettings() {
  return api_('SETTINGS.LIST', { minRole: 'ADMIN', audit: false }, () => appCachedRead_('SETTINGS', () => ({ settings: appSettingsCurrent_().map(x => {
    if (x.type !== 'sheet' || !x.value) return x;
    let name = '';
    try { name = DriveApp.getFileById(x.value).getName(); } catch (e) { name = '（開けません）'; }
    return Object.assign({}, x, { display: name, url: 'https://docs.google.com/spreadsheets/d/' + x.value + '/edit' });
  }) })));
}

function apiSaveSetting(input) {
  return api_('SETTING.SAVE', { minRole: 'ADMIN', entityType: 'SETTING', entityId: input && input.key, detail: input,
    after: res => res.saved }, ctx => appSaveSetting_(ctx, input));
}

/** ログ（操作の記録 AUDIT・実行 RUN・エラー ERROR）の、選んだ月の行。月の一覧つき */
function apiListAudit(input) {
  return api_('AUDIT.LIST', { minRole: 'ADMIN', audit: false }, () => appLogList_(input));
}

/** 監査の鎖を全部の月で確かめ、締まった月のハッシュを控える */
function apiVerifyAudit() {
  return api_('AUDIT.VERIFY', { minRole: 'ADMIN', entityType: 'SYSTEM', after: res => ({ ok: res.ok, months: res.months.length, rows: res.rows, anchored: res.anchored }) },
    ctx => { const v = appVerifyAuditAll_(); v.anchored = v.ok ? appWithLock_(() => appAnchorClosedMonths_(ctx, v)) : []; return v; });
}

/** 毎日の手入れを、今すぐ動かす */
function apiRunHousekeeping() {
  return api_('HOUSEKEEPING.RUN', { minRole: 'ADMIN', entityType: 'SYSTEM', after: res => ({ ok: res.ok, problems: res.problems }) }, ctx => appHousekeeping_(ctx));
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
  return api_('PLANS.LIST', { minRole: 'VIEWER', audit: false }, () => appCachedRead_('PLANS', () => ({ plans: appListPlans_() })));
}

/** 全部の計画の要点（閲覧は社内全員） */
function apiPortfolio() {
  return api_('PORTFOLIO.LIST', { minRole: 'VIEWER', audit: false }, () => ({ plans: appPortfolioCached_() }));
}

/** 予測の根拠（月ごとの内訳・入力と AI の押し・補正・AI 調査の根拠・前回からの変化）。閲覧は社内全員 */
function apiForecastBasis(input) {
  return api_('FORECAST.BASIS', { minRole: 'VIEWER', audit: false }, () => appForecastBasis_(input && input.planId));
}

/** 精度の推移と学習の影（月の誤差・偏り・P10〜P90 に入った割合・全計画で縮めた偏りの補正）。閲覧は社内全員 */
function apiLearningView(input) {
  return api_('LEARNING.VIEW', { minRole: 'VIEWER', audit: false }, () => appLearningView_(input && input.planId));
}

/** 全計画で作る、情報源の信頼度の事前分布（書かずに見比べる） */
function apiPoolPreview() {
  return api_('POOL.PREVIEW', { minRole: 'ADMIN', audit: false }, () => appPoolPreview_());
}

/** 分析: 年度のメーカーを横に並べる（予測と予算・予測の改訂・精度・AI の話題と市場のまとめ）。閲覧は社内全員。読むだけ。中身は「今日」でも変わる（既定の年度）ので、控えの鍵に日付も入れる */
function apiCrossMaker(input) {
  return api_('INSIGHT.CROSS', { minRole: 'VIEWER', audit: false }, () => appCachedRead_('CROSS\u0001' + String(input && input.fy || '') + '\u0001' + appToday_(), () => appCrossMaker_(input && input.fy)));
}

/**
 * 学び: 人の学び（外れた月の振り返り・人と話題ごとの当たり・四半期レビューの判断・予算の上乗せの実績）。閲覧は社内全員。読むだけ。
 * 人ごとの当たりと判断した人の名前は予算策定担当以上だけ（ほかの人には情報源の種類ごとのまとめ）。見せ方が違う人に同じ控えを返さないよう、控えの鍵に見せ方も入れる
 */
function apiPeopleLearning() {
  return api_('INSIGHT.PEOPLE', { minRole: 'VIEWER', audit: false }, ctx => appCachedRead_('PEOPLE\u0001' + appInsightDetail_(ctx) + '\u0001' + appToday_(), () => appPeopleLearning_(ctx)));
}

/**
 * 学び: AI の学び（信頼度の学び・全計画で縮めた偏り・精度の推移・補正の変化・承認待ちの提案・気になる点）。閲覧は社内全員。読むだけ。
 * 承認待ちの提案と補正の変化の、人や話題の名前と信頼度の提案の根拠は予算策定担当以上だけ。人の学びと同じく、控えの鍵に見せ方も入れる
 */
function apiAiLearning() {
  return api_('INSIGHT.AI', { minRole: 'VIEWER', audit: false }, ctx => appCachedRead_('AILEARN\u0001' + appInsightDetail_(ctx) + '\u0001' + appToday_(), () => appAiLearning_(ctx)));
}

/** 計画の公式版の一覧と、今の数字 */
function apiVersionList(input) {
  return api_('VERSION.LIST', { minRole: 'VIEWER', audit: false }, ctx => appCachedRead_('VERSIONS\u0001' + ctx.actor + '\u0001' + (input && input.planId), () => appVersionList_(ctx, input)));
}

/** 今の予測と予算を、公式版として出す（予算策定担当。その計画のクライアントの担当でもよい） */
function apiVersionSubmit(input) {
  return api_('VERSION.SUBMIT', { minRole: 'PLANNER', clientId: appJobPlanClient_({ planId: input && input.planId }), entityType: 'PLAN_VERSION',
    detail: { planId: input && input.planId, note: input && input.note }, after: res => ({ versionNo: res.version.no, budgetFinal: res.version.budget.final,
      p50: res.version.annual.p50, withdrawn: res.withdrawn }) }, ctx => appVersionSubmit_(ctx, input));
}

/** 承認待ちの版を承認・却下する（承認者） */
function apiVersionDecide(input) {
  return api_('VERSION.DECIDE', { minRole: 'APPROVER', clientId: appVersionClient_(input), entityType: 'PLAN_VERSION',
    detail: { versionId: input && input.versionId, decision: input && input.decision, note: input && input.note },
    after: res => ({ versionNo: res.version.no, state: res.version.state, superseded: res.superseded, selfApproved: res.selfApproved }) }, ctx => appVersionDecide_(ctx, input));
}

/** 新しい計画のクライアントの候補（ZAC の実績から） */
function apiPlanCandidates(input) {
  return api_('PLAN.CANDIDATES', { minRole: 'ADMIN', audit: false }, ctx => appPlanCandidates_(ctx, input));
}

/** ホーム（最重要のことだけ。よみのセリフの材料）。閲覧は社内全員 */
function apiHome() {
  return api_('HOME', { minRole: 'VIEWER', audit: false }, ctx => appHome_(ctx));
}

/** 計画の最新の予測（閲覧は社内全員） */
function apiForecastLatest(input) {
  return api_('FORECAST.LATEST', { minRole: 'VIEWER', audit: false }, () => appCachedRead_('LATEST\u0001' + (input && input.planId), () => appForecastLatest_(input && input.planId)));
}

/** 計画の画面（入力・予測と予算・検証・四半期レビュー・進み。閲覧は社内全員）。中身は「今日」でも変わる（既定の年度など）ので、控えの鍵に日付も入れる */
function apiPlanView(input) {
  return api_('PLAN.VIEW', { minRole: 'VIEWER', audit: false }, ctx => appCachedRead_('VIEW\u0001' + ctx.actor + '\u0001' + (input && input.planId) + '\u0001' + appToday_(), () => appPlanView_(ctx, input && input.planId)));
}

/** 時間のかかる処理を始める（裏で動かす。種類ごとに権限が違う） */
function apiStartJob(input) {
  return api_('JOB.START', appJobStartOpts_(input), ctx => appStartJobNow_(ctx, input));
}

/** 年度を締める前の確認（管理者。締められるなら控えの指紋だけ返す。まだ何も書かない・記録もしない） */
function apiYearPreview(input) {
  return api_('YEAR.PREVIEW', { minRole: 'ADMIN', audit: false }, ctx => appYearPreview_(ctx, input));
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
    after: res => res }, ctx => appDailyMaintenance_(ctx));
}

/**
 * 市場の動き（A-4 AI 調査）を、計画ごとに週 1 回ずつ自動で集める（毎日 1 時ごろの時間主導のトリガーから。AutoResearch.js）。
 * 本人確認は triggerDailyBackup と同じ（このプロジェクトの、この入口のトリガーの UID のときだけ所有者として扱う）。
 * 始める A-4 は画面から始めるのと同じ処理（頼んだ人の役割と年度の締めは、裏の処理が段ごとに確かめ直す）
 */
function triggerAutoResearch(e) {
  return api_('AUTO.RESEARCH', { minRole: 'ADMIN', entityType: 'SYSTEM', trigger: { event: e, handler: 'triggerAutoResearch' },
    after: res => res }, ctx => appAutoResearchTick_(ctx));
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
    activeJob: (() => { try { return ctx.user.email && appIsSetUp_() ? appActiveJobOf_(ctx.user.email) : null; } catch (e) { return null; } })(),
    // ホームの中身も最初に渡す（開いてすぐ出せる。通信を 1 回減らす）。読めなければ画面が読み直す
    home: (() => { try { return appIsSetUp_() && appHasRole_(ctx.roles, 'VIEWER') ? appHome_(ctx) : null; } catch (e) { Logger.log('ホームの中身: ' + (e && e.message ? e.message : e)); return null; } })()
  };
}

function appDeniedBootstrap_(ctx) {
  return { app: { name: APP_NAME, nameJa: APP_NAME_JA, version: APP_VERSION }, user: { email: ctx.user.email, isOwner: false, isAdmin: false, roles: [] },
    setUp: appIsSetUp_(), allowed: false };
}
