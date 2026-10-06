/**
 * Backup.js — データ本体のバックアップ。毎日 3 時ごろ「バックアップ」フォルダへ複製し、新しい 14 世代を残す。
 * それより古い世代は「アーカイブ」フォルダへ移す（消さない・ゴミ箱にも送らない。2026-10-05 村井さん承認）。
 * アーカイブのフォルダが決まっていない・開けないときは、複製を作る前に止める（作ってから移せないと、残すはずのものを失う）。
 * スタンドアロンのアプリなので、データ本体を複製してもスクリプトは複製されない。
 * バックアップした後に、フォルダの要点（世代の数・最新の名前と時刻・毎日のバックアップが有効か）を Script Properties に残す。
 * ホームはドライブを数えずにこれを読む（appBackupStatusFast_）。状態の点検（apiHealth）は今までどおりドライブを数える。
 */
const APP_BACKUP_STATUS_PROP = 'APP_BACKUP_STATUS';
const APP_BACKUP_STATUS_TTL_SEC = 300;   // 要点がまだ無いとき、ドライブを数えた結果を控える時間（今までのホームと同じ 5 分）

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
  const files = appBackupFiles_();
  const old = files.slice(APP_BACKUP_KEEP);
  let moved = 0;
  try {
    old.forEach(f => { f.moveTo(archive); moved++; });   // 移せなければ例外。ゴミ箱への代替はしない（「消さない」の約束）
  } finally {
    appBackupRecordStatus_(files.length - moved, files[0] || null);   // 途中で止まっても、作った複製と残っている数を残す
  }
  const res = { backup: copy.getName(), archived: old.map(f => f.getName()), keep: APP_BACKUP_KEEP };
  appRunLog_({ requestId: ctx.requestId, kind: 'BACKUP', startedAt: Utilities.formatDate(started, APP_TZ, "yyyy-MM-dd'T'HH:mm:ssZ"),
    durationMs: new Date().getTime() - started.getTime(), status: 'OK', detail: res });
  return res;
}

/** 毎日のバックアップのトリガーがあるか */
function appBackupEnabled_() {
  return ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'triggerDailyBackup');
}

/** バックアップのフォルダの要点を Script Properties に残す（count = 残っている世代の数、latest = 一番新しいファイル）。失敗してもバックアップは止めない */
function appBackupRecordStatus_(count, latest) {
  try {
    appProps_().setProperty(APP_BACKUP_STATUS_PROP, JSON.stringify({ count: count, latest: latest ? latest.getName() : '',
      latestAt: latest ? latest.getDateCreated().getTime() : 0, enabled: appBackupEnabled_(), at: appNowIso_() }));
  } catch (e) { Logger.log('バックアップの要点を残せません: ' + (e && e.message ? e.message : e)); }
}

/** バックアップの状態（ドライブを数える。状態の点検で使う） */
function appBackupStatus_() {
  const files = appBackupFiles_();
  const enabled = appBackupEnabled_();
  const latestAt = files.length ? files[0].getDateCreated().getTime() : 0;
  return { enabled: enabled, count: files.length, latest: files.length ? files[0].getName() : '',
    ageHours: latestAt ? Math.round((new Date().getTime() - latestAt) / 36e5) : null, keep: APP_BACKUP_KEEP };
}

/**
 * バックアップの状態（appBackupStatus_ と同じ形）を、バックアップしたときに残した要点から返す（ドライブを数えない。ホームで使う）。
 * 要点がまだ無い（この版になってからバックアップしていない）ときだけ、今までどおりドライブを数え、5 分控える
 */
function appBackupStatusFast_() {
  let s = null;
  try { s = JSON.parse(appProps_().getProperty(APP_BACKUP_STATUS_PROP) || 'null'); } catch (e) { s = null; }
  if (s && typeof s === 'object') {
    const latestAt = Number(s.latestAt) || 0;
    return { enabled: !!s.enabled, count: Number(s.count) || 0, latest: String(s.latest || ''),
      ageHours: latestAt ? Math.round((new Date().getTime() - latestAt) / 36e5) : null, keep: APP_BACKUP_KEEP };
  }
  const cache = CacheService.getScriptCache();
  try { const hit = cache.get('HOME_SYS'); if (hit) return JSON.parse(hit); } catch (e) { /* 控えが読めなければ数える */ }
  const st = appBackupStatus_();
  try { cache.put('HOME_SYS', JSON.stringify(st), APP_BACKUP_STATUS_TTL_SEC); } catch (e) { /* 控えは速くするためだけ */ }
  return st;
}

/** 毎日のバックアップを有効にする（すでにあれば何もしない） */
function appEnableBackup_() {
  const already = appBackupEnabled_();
  if (!already) ScriptApp.newTrigger('triggerDailyBackup').timeBased().everyDays(1).atHour(APP_BACKUP_HOUR).create();
  // 残した要点の「有効」も合わせる（ホームが次のバックアップまで「無効」と出さないように）
  try {
    const s = JSON.parse(appProps_().getProperty(APP_BACKUP_STATUS_PROP) || 'null');
    if (s && typeof s === 'object' && !s.enabled) appProps_().setProperty(APP_BACKUP_STATUS_PROP, JSON.stringify(Object.assign(s, { enabled: true })));
  } catch (e) { Logger.log('バックアップの要点を直せません: ' + (e && e.message ? e.message : e)); }
  return { enabled: true, already: already };
}
