/**
 * Config.js — 売上予測アプリの定数（段階1: 土台）。
 * 設計: Forecast_Agent/DESIGN_data_platform_JA.md。正本は「売上予測アプリ データ」、記録は年度ごとの「売上予測アプリ ログ FYyyyy」。
 * 公開する関数は doGet・api*・trigger* だけにし、すべて api_（本人確認・役割・記録）を通す（app/tests が見張る）。
 * それ以外の関数は名前の末尾を _ にする（Web アプリでは _ でない関数をブラウザから呼べるため）。
 */
/** アプリの名前（社内の Web アプリと同じく英語 + 日本語。ブラウザのタブは英語だけ。2026-10-02 村井さん決定） */
const APP_NAME = 'Trends2Targets';
const APP_NAME_JA = '売上予測と予算策定';
const APP_VERSION = '0.8.2';
const APP_TZ = 'Asia/Tokyo';

const APP_FILES = {
  folder: '売上予測アプリ（システム）',
  data: '売上予測アプリ データ',
  logPrefix: '売上予測アプリ ログ ',
  backupFolder: 'バックアップ',
  archiveFolder: 'アーカイブ',
  backupPrefix: '売上予測アプリ データ バックアップ ',
  scratch: '売上予測アプリ 計算用（自動）',
  journalPrefix: '売上予測アプリ 保存の控え（自動） '   // 書き終えたらゴミ箱へ（Journal.js）
};

/** Script Properties のキー（ファイルの ID と鎖の最新ハッシュ。秘密情報は置かない） */
const APP_PROP = {
  folderId: 'APP_FOLDER_ID',
  dataId: 'APP_DATA_SPREADSHEET_ID',
  logFiles: 'APP_LOG_SPREADSHEETS_JSON',
  backupFolderId: 'APP_BACKUP_FOLDER_ID',
  archiveFolderId: 'APP_ARCHIVE_FOLDER_ID',
  auditLastHash: 'APP_AUDIT_LAST_HASH',
  internalDomain: 'APP_INTERNAL_DOMAIN',
  scratchId: 'APP_SCRATCH_SPREADSHEET_ID',
  scratchOwner: 'APP_SCRATCH_OWNER',     // 計算用ブックを今使っている処理の印（組み立て → 計算 → 保存の続きの処理が、同じ中身のままか確かめる）
  fyNaming: 'APP_FY_NAMING',   // 'start' = 年度を始まりの年で呼ぶ（ログのファイル名を直し終えた印）
  tablesVersion: 'APP_TABLES_VERSION'   // データ本体の表をそろえた版（APP_SCHEMA_VERSION と違えば、最初の操作で足りない表を作る）
};

/** 役割（弱い → 強い）。閲覧と情報提供は社内全員が持つ。予測と予算の提出は担当者、承認は承認者（2026-10-01 決定） */
const APP_ROLES = ['VIEWER', 'CONTRIBUTOR', 'PLANNER', 'APPROVER', 'ADMIN'];
/** 管理画面で付与できる役割（VIEWER と CONTRIBUTOR は社内全員に自動で付く） */
const APP_GRANTABLE_ROLES = ['PLANNER', 'APPROVER', 'ADMIN'];
const APP_ROLE_LABELS = { VIEWER: '閲覧', CONTRIBUTOR: '情報提供', PLANNER: '予算策定担当', APPROVER: '承認者', ADMIN: '管理者' };

const APP_BACKUP_KEEP = 14;
const APP_BACKUP_HOUR = 3;
const APP_JSON_MAX = 40000;
const APP_LOCK_WAIT_MS = 30000;
