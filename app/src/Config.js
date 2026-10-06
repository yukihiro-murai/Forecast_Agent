/**
 * Config.js — Trends2Targets の定数（段階1: 土台）。
 * 設計: Trends2Targets/DESIGN_data_platform_JA.md。正本は「Trends2Targets データ」、記録は年度ごとの「Trends2Targets ログ FYyyyy」。
 * 公開する関数は doGet・api*・trigger* だけにし、すべて api_（本人確認・役割・記録）を通す（app/tests が見張る）。
 * それ以外の関数は名前の末尾を _ にする（Web アプリでは _ でない関数をブラウザから呼べるため）。
 */
/** アプリの名前（社内の Web アプリと同じく英語 + 日本語。ブラウザのタブは英語だけ。2026-10-02 村井さん決定） */
const APP_NAME = 'Trends2Targets';
const APP_NAME_JA = '売上予測と予算策定';
const APP_VERSION = '0.27.0';
const APP_TZ = 'Asia/Tokyo';

const APP_FILES = {
  folder: 'Trends2Targets（システム）',
  data: 'Trends2Targets データ',
  logPrefix: 'Trends2Targets ログ ',
  backupFolder: 'バックアップ',
  archiveFolder: 'アーカイブ',
  backupPrefix: 'Trends2Targets データ バックアップ ',
  scratch: 'Trends2Targets 計算用（自動）',
  journalPrefix: 'Trends2Targets 保存の控え（自動） ',   // 書き終えたらゴミ箱へ（Journal.js）
  monthlyBackupPrefix: 'Trends2Targets データ 月次 ',   // 月の最初のバックアップを、アーカイブのフォルダに月ごとに残す（消さない）
  yearPrefix: 'Trends2Targets 年度 FY'   // 年度の締めの控え（アーカイブのフォルダへ。消さない）
};
/**
 * 2026-10-06 に「売上予測アプリ …」から Trends2Targets へ改名する前の名前（今あるファイルは同じ ID のまま改名済み）。
 * ファイルはどれも ID で開くので名前に頼らないが、バックアップの世代と月次の写しは名前で探すため、前の名前のものも数える。
 */
const APP_FILES_LEGACY = { backupPrefix: '売上予測アプリ データ バックアップ ', monthlyBackupPrefix: '売上予測アプリ データ 月次 ' };

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
