/**
 * Backup.js — データ本体のバックアップ。毎日 3 時ごろ「バックアップ」フォルダへ複製し、新しい 14 世代を残す。
 * それより古い世代は「アーカイブ」フォルダへ移す（消さない・ゴミ箱にも送らない。2026-10-05 村井さん承認）。
 * アーカイブのフォルダが決まっていない・開けないときは、複製を作る前に止める（作ってから移せないと、残すはずのものを失う）。
 * スタンドアロンのアプリなので、データ本体を複製してもスクリプトは複製されない。
 */
function appBackupFiles_() {
  const folderId = appProps_().getProperty(APP_PROP.backupFolderId);
  if (!folderId) return [];
  const it = DriveApp.getFolderById(folderId).getFilesByType(MimeType.GOOGLE_SHEETS);
  const files = [];
  while (it.hasNext()) {
    const f = it.next();
    // フォルダの一覧にはゴミ箱へ移したファイルも出るので除く
    if (!f.isTrashed() && f.getName().indexOf(APP_FILES.backupPrefix) === 0) files.push(f);
  }
  return files.sort((a, b) => b.getDateCreated().getTime() - a.getDateCreated().getTime());
}

function appBackup_(ctx) {
  const props = appProps_();
  const dataId = props.getProperty(APP_PROP.dataId);
  const folderId = props.getProperty(APP_PROP.backupFolderId);
  if (!dataId || !folderId) throw new Error('初期設定がまだです。');
  const archiveId = props.getProperty(APP_PROP.archiveFolderId);
  if (!archiveId) throw new Error('アーカイブのフォルダがまだありません。管理者が「初期設定」を実行してください。');
  // 複製を作る前にアーカイブのフォルダを確かめる（無ければここで止まり、複製も古い世代の移動もしない）
  const archive = DriveApp.getFolderById(archiveId);
  const started = new Date();
  const name = APP_FILES.backupPrefix + Utilities.formatDate(started, APP_TZ, 'yyyy-MM-dd HHmm');
  SpreadsheetApp.flush();
  const copy = DriveApp.getFileById(dataId).makeCopy(name, DriveApp.getFolderById(folderId));
  const old = appBackupFiles_().slice(APP_BACKUP_KEEP);
  old.forEach(f => f.moveTo(archive));   // 移せなければ例外。ゴミ箱への代替はしない（「消さない」の約束）
  const res = { backup: copy.getName(), archived: old.map(f => f.getName()), keep: APP_BACKUP_KEEP };
  appRunLog_({ requestId: ctx.requestId, kind: 'BACKUP', startedAt: Utilities.formatDate(started, APP_TZ, "yyyy-MM-dd'T'HH:mm:ssZ"),
    durationMs: new Date().getTime() - started.getTime(), status: 'OK', detail: res });
  return res;
}

function appBackupStatus_() {
  const files = appBackupFiles_();
  const enabled = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'triggerDailyBackup');
  const latestAt = files.length ? files[0].getDateCreated().getTime() : 0;
  return { enabled: enabled, count: files.length, latest: files.length ? files[0].getName() : '',
    ageHours: latestAt ? Math.round((new Date().getTime() - latestAt) / 36e5) : null, keep: APP_BACKUP_KEEP };
}

/** 毎日のバックアップを有効にする（すでにあれば何もしない） */
function appEnableBackup_() {
  const already = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'triggerDailyBackup');
  if (!already) ScriptApp.newTrigger('triggerDailyBackup').timeBased().everyDays(1).atHour(APP_BACKUP_HOUR).create();
  return { enabled: true, already: already };
}
