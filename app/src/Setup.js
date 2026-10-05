/**
 * Setup.js — 初期設定（所有者だけ・何度実行しても同じ結果）。
 * マイドライブに「売上予測アプリ（システム）」フォルダを作り、その中に データ本体・今年度のログ・バックアップ・アーカイブを置く。
 * どれも所有者だけのファイルで、メンバーには共有しない（メンバーはアプリ画面だけを使う）。
 */
function appSetup_(ctx) {
  return appWithLock_(() => {
    const props = appProps_();
    const created = [];
    // 1) フォルダ
    let folder = null;
    const folderId = props.getProperty(APP_PROP.folderId);
    if (folderId) folder = DriveApp.getFolderById(folderId);
    if (!folder) {
      folder = DriveApp.createFolder(APP_FILES.folder);
      props.setProperty(APP_PROP.folderId, folder.getId());
      created.push('フォルダ「' + APP_FILES.folder + '」');
    }
    [[APP_PROP.backupFolderId, APP_FILES.backupFolder], [APP_PROP.archiveFolderId, APP_FILES.archiveFolder]].forEach(p => {
      if (!props.getProperty(p[0])) {
        props.setProperty(p[0], folder.createFolder(p[1]).getId());
        created.push('フォルダ「' + p[1] + '」');
      }
    });
    // 2) 今年度のログ（ここから先は記録できる）。以前の名前のログがあれば先に直す
    appEnsureFyNaming_();
    const logExisted = !!JSON.parse(props.getProperty(APP_PROP.logFiles) || '{}')['FY' + appFy_(new Date())];
    const log = appLogSpreadsheet_(new Date(), 'write');
    if (!logExisted) created.push('ログ「' + APP_FILES.logPrefix + 'FY' + appFy_(new Date()) + '」');
    // 3) データ本体
    let dataNew = false;
    if (!props.getProperty(APP_PROP.dataId)) {
      const ss = SpreadsheetApp.create(APP_FILES.data);
      DriveApp.getFileById(ss.getId()).moveTo(folder);
      props.setProperty(APP_PROP.dataId, ss.getId());
      APP_STORE_CACHE_ = {};
      dataNew = true;
      created.push('データ本体「' + APP_FILES.data + '」');
    }
    // 4) 社内ドメイン（所有者のドメイン）
    if (!props.getProperty(APP_PROP.internalDomain)) props.setProperty(APP_PROP.internalDomain, ctx.user.owner.split('@')[1] || '');
    // 5) 表と _SCHEMA（無い表だけ作る。列が違う表があれば止まる）
    const ss = appDataSpreadsheet_();
    const madeTables = appEnsureTables_(ctx);
    // 6) 所有者をメンバーに
    if (!appReadTable_('MEMBERS').some(m => m.email === ctx.user.owner)) {
      const now = appNowIso_();
      appInsertRows_('MEMBERS', [{ email: ctx.user.owner, display_name: ctx.user.owner.split('@')[0], department: '', is_active: true,
        note: '所有者（アプリを動かす人）', created_at: now, created_by: ctx.actor, updated_at: now, updated_by: ctx.actor, row_version: 1 }]);
      created.push('メンバーに所有者');
    }
    const result = {
      created: created, tables: madeTables, dataNew: dataNew,
      folderUrl: folder.getUrl(), dataUrl: ss.getUrl(), logUrl: log.getUrl()
    };
    appAuditAppend_({ actor: ctx.actor, actorRoles: appRoleSummary_(ctx.roles), action: 'SETUP.INIT', phase: 'END', result: 'OK',
      entityType: 'SYSTEM', detail: { created: created, tables: madeTables }, requestId: ctx.requestId });
    return result;
  });
}

/**
 * 無い表を作り、_SCHEMA に記録する（何度実行しても同じ。列が定義と違う表があれば止まる）。作った表の名前を返す。
 * 版を上げて表を足したときは、初期設定をもう一度実行するか、最初の操作（appAutoEnsureTables_）がこれを呼ぶ。使わなくなった表は隠す。
 */
function appEnsureTables_(ctx) {
  return appWithLock_(() => {
    const ss = appDataSpreadsheet_();
    // 年度の凍結の記録（YEAR_CLOSURES）は、一度作られた後に無くなっていたら自動では作り直さない。
    // 空で作り直すと「締めた年度」が全部開いてしまうため、止めて復旧を求める（初めての作成・版の移行では _SCHEMA に記録が無いので普通に作る）。
    if (!ss.getSheetByName('YEAR_CLOSURES') && ss.getSheetByName('_SCHEMA') && appYearRegistryWasInitialized_()) {
      throw new Error('年度の凍結の記録の表（YEAR_CLOSURES）が見つかりません。自動では作り直しません。元の記録を戻すか、管理者に連絡してください。');
    }
    const made = [];
    Object.keys(APP_TABLES).forEach(name => {
      if (!ss.getSheetByName(name)) made.push(name);
      appTableSheet_(name, true);
    });
    appHideRetiredTables_(ss);
    ss.getSheets().forEach(sh => {
      if (!APP_TABLES[sh.getName()] && sh.getLastRow() === 0 && /^(シート|Sheet)\d+$/.test(sh.getName())) ss.deleteSheet(sh);
    });
    const schemaRows = appReadTable_('_SCHEMA');
    const missing = Object.keys(APP_TABLES).filter(n => !schemaRows.some(r => r.table === n));
    appInsertRows_('_SCHEMA', missing.map(n => ({ table: n, schema_version: APP_SCHEMA_VERSION, columns_hash: appColumnsHash_(n),
      migrated_at: appNowIso_(), migrated_by: ctx.actor })));
    appProps_().setProperty(APP_PROP.tablesVersion, String(APP_SCHEMA_VERSION));
    appBumpGen_();
    return made;
  });
}

/**
 * 版を上げて表を足したとき、最初の操作（画面を開く・裏の処理を含む）で足りない表を作る。「初期設定」をもう一度実行しなくてよい。
 * 2026-10-02 に版 3 の表（FORECAST_RUNS）が無くて「予測」の画面が開けなかったため。
 * 作った表は実行ログに残す。失敗しても操作は止めずエラーのログに残す（その表を使う操作がそこで止まる）。
 */
function appAutoEnsureTables_(ctx) {
  try {
    if (!appIsSetUp_()) return;
    if (appProps_().getProperty(APP_PROP.tablesVersion) === String(APP_SCHEMA_VERSION)) return;
    const made = appEnsureTables_(ctx);
    if (made.length) appRunLog_({ requestId: ctx.requestId, kind: 'SCHEMA.ENSURE', status: 'OK', detail: { version: APP_SCHEMA_VERSION, made: made } });
  } catch (e) {
    appLogError_('SCHEMA.ENSURE', e, ctx);
  }
}

/** 状態の点検（表の列・_SCHEMA・監査の鎖・バックアップ・定期処理）。管理画面の「状態」 */
function appHealth_() {
  const out = { setUp: appIsSetUp_(), tables: [], audit: null, backup: null, triggers: [], files: {} };
  if (!out.setUp) return out;
  const props = appProps_();
  const schemaRows = (() => { try { return appReadTable_('_SCHEMA'); } catch (e) { return []; } })();
  const ss = (() => { try { return appDataSpreadsheet_(); } catch (e) { return null; } })();
  Object.keys(APP_TABLES).forEach(name => {
    const t = { name: name, ok: true, rows: 0, note: '', engine: name.indexOf('ENG_') === 0 };
    try {
      if (ss && !ss.getSheetByName(name)) {
        t.ok = false; t.missing = true; t.note = '未作成（「初期設定」をもう一度実行すると作られます）';
      } else {
        t.rows = name === '_SCHEMA' ? schemaRows.length : appReadTable_(name).length;
        const s = schemaRows.filter(r => r.table === name)[0];
        if (!s) { t.ok = false; t.note = '_SCHEMA に記録なし'; }
        else if (s.columns_hash !== appColumnsHash_(name)) { t.ok = false; t.note = '列の定義が _SCHEMA と違う（移行が必要）'; }
      }
    } catch (e) { t.ok = false; t.note = String(e && e.message || e); }
    out.tables.push(t);
  });
  try {
    const now = new Date();
    const ss = appLogSpreadsheet_(now, 'read');
    const sh = ss && ss.getSheetByName('AUDIT_' + Utilities.formatDate(now, APP_TZ, 'yyyy_MM'));
    const v = sh ? appVerifyAuditSheet_(sh) : { rows: 0, ok: true, brokenAt: 0, lastHash: '' };
    v.matchesLatest = !v.lastHash || v.lastHash === String(props.getProperty(APP_PROP.auditLastHash) || '');
    out.audit = v;
  } catch (e) {
    out.audit = { ok: false, note: String(e && e.message || e) };
  }
  try { out.backup = appBackupStatus_(); } catch (e) { out.backup = { note: String(e && e.message || e) }; }
  try { out.triggers = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction()); } catch (e) { out.triggers = []; }
  out.housekeeping = appHousekeepingLast_();
  try { out.journal = appJournalPending_(); } catch (e) { out.journal = null; }
  try { out.jobs = appJobList_().length; } catch (e) { out.jobs = null; }
  try { out.years = appYearStatus_(); } catch (e) { out.years = { error: String(e && e.message || e) }; }   // 年度の一覧（締めの管理。読めなければそのことを出す）
  try {
    out.files.folder = DriveApp.getFolderById(props.getProperty(APP_PROP.folderId)).getUrl();
    out.files.data = appDataSpreadsheet_().getUrl();
    const log = appLogSpreadsheet_(new Date(), 'read');
    out.files.log = log ? log.getUrl() : '';
  } catch (e) { out.files.note = String(e && e.message || e); }
  return out;
}
