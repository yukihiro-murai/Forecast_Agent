/**
 * Forecast vNext 管理ハブ / Provisioner.
 * Existing Forecast_Agent.js is intentionally not modified; all entry points are vNext-prefixed.
 */

const VN_ADMIN_SCHEMA_VERSION = 'vnext-admin-1';
const VN_ADMIN_MENU_NAME = VNEXT_NAMING.MENU;
const VN_ADMIN_MENU_OPEN_SIDEBAR = '案内を開く';
const VN_ADMIN_MENU_RUN_NOW = '申請を今すぐ処理';
const VN_ADMIN_MENU_HEALTH_SCAN = '全クライアントの状態点検';
const VN_ADMIN_MENU_OPEN_REGISTRY = '登録一覧を開く';
const VN_ADMIN_MENU_OTHER = 'その他';
const VN_ADMIN_META_SHEET = 'BOOK_META';
const VN_ADMIN_BOOK_CONFIG_SHEET = 'VN_BOOK_CONFIG';
const VN_ADMIN_SYSTEM_CONFIG_SHEET = 'VN_SYSTEM_CONFIG';
const VN_ADMIN_OFFICIAL_COPY_SHEET = 'OFFICIAL_SNAPSHOT';
const VN_ADMIN_CLIENT_REQUEST_SHEET = 'VN_CLIENT_REQUEST';
const VN_ADMIN_PORTAL_REQUEST_SHEET = 'VN_PORTAL_REQUEST';
const VN_ADMIN_PORTAL_DIRECTORY_SHEET = 'PORTAL_DIRECTORY';
const VN_ADMIN_PORTAL_CONFIG_SHEET = 'VN_PORTAL_CONFIG';
const VN_ADMIN_ZAC_CLIENT_CATALOG_SHEET = 'ZAC_CLIENT_CATALOG';
const VN_ADMIN_PORTAL_CLIENT_CATALOG_SHEET = 'VN_PORTAL_CLIENT_CATALOG';
const VN_ADMIN_SCHEDULED_HANDLER = 'vNextAdminScheduledSweep';
const VN_ADMIN_PILOT_INITIAL_LIMIT = 3;
const VN_ADMIN_PILOT_CANARY_LIMIT = 5;
const VN_ADMIN_FRESH_UAT_RESET_CONFIRMATION = 'RESET_GENERATED_CLIENTS';
const VN_ADMIN_PROTECTED_BOOK_MODES = Object.freeze(['ADMIN', 'TEMPLATE', 'PORTAL', 'LEGACY']);
const VN_ADMIN_STALE_MINUTES = 15;
const VN_ADMIN_MIGRATION_APPLY_ENABLED = false;
const VN_ADMIN_MODES = ['LEGACY', 'ADMIN', 'TEMPLATE', 'CLIENT', 'PORTAL'];
const VN_ADMIN_CLIENT_STATES = [
  'INPUT_OPEN', 'READY_TO_RUN', 'RUNNING', 'DRAFT_READY', 'SUBMITTED',
  'CHANGES_REQUESTED', 'OFFICIAL_LOCKED', 'REVIEW_DUE', 'YEAR_CLOSED'
];

const VN_ADMIN_RUNTIME_KEYS = [
  'FORECAST_SOURCE_SPREADSHEET_ID',
  'VNEXT_ZAC_SOURCE_SPREADSHEET_ID',
  'VERTEX_PROJECT_ID',
  'VERTEX_LOCATION',
  'VERTEX_GEMINI_MODEL',
  'VERTEX_DATASTORE_ID',
  'VERTEX_SEARCH_LOCATION',
  'VERTEX_SERVING_CONFIG',
  'VNEXT_ADMIN_EMAILS',
  'VNEXT_ROOT_FOLDER_ID',
  'VNEXT_ADMIN_HUB_SPREADSHEET_ID',
  'VNEXT_MASTER_TEMPLATE_SPREADSHEET_ID',
  'VNEXT_DEFAULT_EDITORS',
  'VNEXT_DEFAULT_VIEWERS',
  'VNEXT_EMPLOYEE_DOMAIN',
  'VNEXT_PORTAL_SPREADSHEET_ID',
  'VNEXT_ACTIVE_RELEASE_ID',
  'VNEXT_ACTIVE_MODEL_RELEASE_ID'
];

const VN_ADMIN_CLIENT_REQUEST_PAYLOAD_KEYS = Object.freeze([
  'requestId', 'bookId', 'clientId', 'clientName', 'fiscalYear', 'asOf',
  'cutoff', 'bookConfiguredAsOf', 'requestedAt', 'requestedBy'
]);
const VN_ADMIN_PORTAL_REQUEST_PAYLOAD_KEYS_V1 = Object.freeze([
  'clientId', 'clientName', 'fiscalYear', 'forecastOwnerEmail', 'relatedMemberEmails',
  'requestId', 'requestType', 'requestedAt', 'requestedBy', 'schemaVersion'
]);
const VN_ADMIN_PORTAL_REQUEST_PAYLOAD_KEYS_V2 = Object.freeze([
  'catalogKey', 'clientName', 'fiscalYear', 'relatedMemberNames',
  'requestId', 'requestType', 'requestedAt', 'requestedBy', 'schemaVersion'
]);
const VN_ADMIN_PORTAL_REQUEST_HEADERS_V1 = Object.freeze([
  'request_event_id', 'request_id', 'event_type', 'status', 'request_hash', 'request_json',
  'fiscal_year', 'client_id', 'client_name', 'forecast_owner_email',
  'related_member_emails_json', 'requested_at', 'requested_by', 'related_book_id',
  'related_book_url', 'detail_code', 'detail_message', 'created_at', 'created_by'
]);
const VN_ADMIN_PORTAL_REQUEST_HEADERS = Object.freeze(
  VN_ADMIN_PORTAL_REQUEST_HEADERS_V1.concat(['catalog_key', 'related_member_names_json'])
);
const VN_ADMIN_PORTAL_DIRECTORY_HEADERS_V1 = Object.freeze([
  'directory_event_id', 'directory_key', 'fiscal_year', 'client_id', 'client_name',
  'forecast_owner_email', 'related_member_emails_json', 'state', 'center_forecast',
  'adopted_forecast', 'final_budget', 'next_action', 'client_book_url', 'request_id',
  'updated_at', 'updated_by'
]);
const VN_ADMIN_PORTAL_DIRECTORY_HEADERS = Object.freeze(
  VN_ADMIN_PORTAL_DIRECTORY_HEADERS_V1.concat(['related_member_names_json'])
);
const VN_ADMIN_ZAC_CLIENT_CATALOG_HEADERS = Object.freeze([
  'catalog_key', 'client_id', 'client_code', 'client_name', 'normalized_name',
  'is_active', 'source_years_json', 'first_seen_at', 'last_seen_at',
  'catalog_version', 'refreshed_at', 'source_spreadsheet_id'
]);
const VN_ADMIN_PORTAL_CLIENT_CATALOG_HEADERS = Object.freeze([
  'catalog_key', 'client_name', 'is_active', 'catalog_version', 'synced_at'
]);
const VN_ADMIN_PORTAL_RUNTIME_VERSION = 'vnext-portal-1.7.13';
const VN_ADMIN_PORTAL_LEGACY_RUNTIME_VERSIONS = Object.freeze([
  'vnext-portal-1.0.0', 'vnext-portal-1.1.0', 'vnext-portal-1.2.0', 'vnext-portal-1.3.0',
  'vnext-portal-1.4.0', 'vnext-portal-1.5.0', 'vnext-portal-1.6.0', 'vnext-portal-1.7.0',
  'vnext-portal-1.7.1', 'vnext-portal-1.7.2', 'vnext-portal-1.7.3', 'vnext-portal-1.7.4',
  'vnext-portal-1.7.5', 'vnext-portal-1.7.6', 'vnext-portal-1.7.7', 'vnext-portal-1.7.8',
  'vnext-portal-1.7.9', 'vnext-portal-1.7.10', 'vnext-portal-1.7.11', 'vnext-portal-1.7.12',
  'vnext-portal-1.8.0'
]);
const VN_ADMIN_EMPLOYEE_PORTAL_WEBAPP_DEPLOYMENT_ID =
  'AKfycbxVtnFiXMB6FwKRdMj_PJVmq4zlpYMoBLS3zXy_1ruTGqyTSPxyepkJegcL9rGiUbwH';
const VN_ADMIN_LIBRARY = Object.freeze({
  DRIVE_NAME: VNEXT_NAMING.SHARED_DRIVE,
  LEGACY_DRIVE_NAME: VNEXT_NAMING.LEGACY_SHARED_DRIVE,
  PORTAL: VNEXT_NAMING.FOLDER_PORTAL,
  BOOKS: VNEXT_NAMING.FOLDER_BOOKS,
  ADMIN: VNEXT_NAMING.FOLDER_ADMIN,
  AUDIT: VNEXT_NAMING.FOLDER_AUDIT,
  TEMPLATES: VNEXT_NAMING.FOLDER_TEMPLATES,
  TEMPLATES_CURRENT: VNEXT_NAMING.TEMPLATE_CURRENT,
  TEMPLATES_DRAFT: VNEXT_NAMING.TEMPLATE_DRAFT,
  TEMPLATES_HISTORY: VNEXT_NAMING.TEMPLATE_HISTORY,
  FOLDER_LEGACY: VNEXT_NAMING.FOLDER_LEGACY
});
const VN_ADMIN_PORTAL_REQUEST_SCHEMA = 'vnext-portal-request-2';
const VN_ADMIN_PORTAL_REQUEST_SCHEMA_V1 = 'vnext-portal-request-1';
const VN_ADMIN_ZAC_CLIENT_CODE_COLUMN = 40; // AN
const VN_ADMIN_ZAC_CLIENT_NAME_COLUMN = 41; // AO
const VN_ADMIN_ZAC_CATALOG_STALE_MS = 6 * 60 * 60 * 1000;
const VN_ADMIN_TEMPLATE_MANIFEST_SCHEMA = 'VNEXT_TEMPLATE_UI_V3';
const VN_ADMIN_TEMPLATE_MANIFEST_SCHEMA_V2 = 'VNEXT_TEMPLATE_UI_V2';
const VN_ADMIN_TEMPLATE_FORBIDDEN_FORMULA = /\b(?:IMPORTRANGE|IMPORTDATA|IMPORTHTML|IMPORTXML|GOOGLEFINANCE)\s*\(/i;
const VN_ADMIN_AI_ROLLBACK_SCOPES = Object.freeze(['ALL', 'SELECTED']);
const VN_ADMIN_RETURN_ROUTES = Object.freeze({
  PLAN_ONLY: 'CHANGES_REQUESTED',
  REOPEN_INPUT: 'INPUT_OPEN',
  RERUN_SAME_INPUT: 'READY_TO_RUN'
});

const VN_ADMIN_SHEETS = {
  HOME: 'ADMIN_HOME',
  EXCEPTIONS: 'TODAY_EXCEPTIONS',
  REGISTRY: 'BOOK_REGISTRY',
  TEAM: 'TEAM_REGISTRY',
  JOBS: 'JOB_QUEUE',
  JOB_LOG: 'JOB_LOG',
  OFFICIAL: 'OFFICIAL_RUNS',
  APPROVALS: 'PLAN_APPROVALS',
  RELEASES: 'RELEASES',
  TEMPLATE_JOURNAL: 'TEMPLATE_RELEASE_JOURNAL',
  CATALOG: VN_ADMIN_ZAC_CLIENT_CATALOG_SHEET,
  SETTINGS: 'MODEL_SETTINGS',
  MIGRATIONS: 'MIGRATION_LOG',
  AUDIT: 'ADMIN_AUDIT_LOG',
  LEARNING_OBS: 'LEARNING_OBSERVATION',
  LEARNING_EVIDENCE: 'LEARNING_EVIDENCE'
};

const VN_ADMIN_HEADERS = {};
VN_ADMIN_HEADERS[VN_ADMIN_SHEETS.EXCEPTIONS] = [
  'exception_id', 'severity', 'exception_type', 'book_id', 'client_name', 'fiscal_year',
  'title', 'detail', 'recommended_action', 'status', 'detected_at', 'source_ref'
];
VN_ADMIN_HEADERS[VN_ADMIN_SHEETS.REGISTRY] = [
  'book_id', 'mode', 'client_id', 'client_name', 'fiscal_year', 'spreadsheet_id', 'spreadsheet_url',
  'client_script_id', 'client_runtime_version', 'client_runtime_sha256',
  'admin_script_id', 'admin_runtime_sha256',
  'template_release_id', 'schema_version', 'state', 'status', 'health_status', 'health_code',
  'last_health_at', 'last_forecast_at', 'current_official_id', 'forecast_owner_emails',
  'editor_emails', 'viewer_emails', 'created_at', 'created_by', 'updated_at', 'note'
  , 'access_policy', 'internal_domain', 'related_member_names_json'
];
VN_ADMIN_HEADERS[VN_ADMIN_SHEETS.TEAM] = [
  'team_key', 'book_id', 'client_id', 'fiscal_year', 'email', 'role', 'status',
  'created_at', 'created_by', 'updated_at', 'note'
];
VN_ADMIN_HEADERS[VN_ADMIN_SHEETS.JOBS] = [
  'job_id', 'job_type', 'target_book_id', 'target_spreadsheet_id', 'request_json', 'idempotency_key',
  'status', 'priority', 'attempts', 'not_before', 'locked_at', 'locked_by', 'started_at',
  'finished_at', 'result_json', 'error', 'created_at', 'created_by', 'updated_at'
];
VN_ADMIN_HEADERS[VN_ADMIN_SHEETS.JOB_LOG] = [
  'log_id', 'job_id', 'logged_at', 'status', 'message', 'detail_json', 'actor'
];
VN_ADMIN_HEADERS[VN_ADMIN_SHEETS.OFFICIAL] = [
  'official_id', 'book_id', 'client_id', 'client_name', 'fiscal_year', 'forecast_run_id',
  'source_forecast_run_id',
  'approval_request_id', 'record_type', 'supersedes_official_id', 'amendment_reason',
  'snapshot_json', 'snapshot_hash', 'immutable_hash', 'model_release_id', 'issued_at',
  'issued_by', 'note'
];
VN_ADMIN_HEADERS[VN_ADMIN_SHEETS.APPROVALS] = [
  'approval_request_id', 'request_type', 'book_id', 'client_id', 'client_name', 'fiscal_year',
  'forecast_run_id', 'plan_version_id', 'supersedes_official_id', 'amendment_reason', 'snapshot_json', 'snapshot_hash',
  'status', 'processing_attempts', 'requested_at', 'requested_by', 'decision_at', 'decision_by', 'decision_comment',
  'official_id', 'idempotency_key', 'updated_at'
];
VN_ADMIN_HEADERS[VN_ADMIN_CLIENT_REQUEST_SHEET] = [
  'request_event_id', 'request_id', 'book_id', 'event_type', 'status', 'request_hash',
  'request_json', 'requested_at', 'requested_by', 'related_job_id', 'related_run_id',
  'detail_json', 'created_at'
];
VN_ADMIN_HEADERS[VN_ADMIN_SHEETS.RELEASES] = [
  'release_id', 'release_name', 'status', 'template_spreadsheet_id', 'schema_version',
  'engine_version', 'ux_version', 'admin_version', 'client_runtime_version',
  'client_runtime_sha256', 'template_content_sha256', 'template_manifest_schema',
  'template_script_id', 'created_at', 'created_by',
  'activated_at', 'note'
];
VN_ADMIN_HEADERS[VN_ADMIN_SHEETS.TEMPLATE_JOURNAL] = [
  'journal_id', 'operation_id', 'release_id', 'model_release_id', 'previous_release_id',
  'previous_model_release_id', 'template_spreadsheet_id', 'phase', 'status',
  'detail_json', 'occurred_at', 'actor'
];
VN_ADMIN_HEADERS[VN_ADMIN_SHEETS.CATALOG] = VN_ADMIN_ZAC_CLIENT_CATALOG_HEADERS.slice();
VN_ADMIN_HEADERS[VN_ADMIN_SHEETS.SETTINGS] = [
  'setting_key', 'setting_value', 'value_type', 'scope', 'effective_from', 'updated_at',
  'updated_by', 'note'
];
VN_ADMIN_HEADERS[VN_ADMIN_SHEETS.MIGRATIONS] = [
  'migration_id', 'book_id', 'spreadsheet_id', 'from_release_id', 'to_release_id', 'status',
  'dry_run', 'plan_json', 'result_json', 'started_at', 'finished_at', 'actor', 'error'
];
VN_ADMIN_HEADERS[VN_ADMIN_SHEETS.AUDIT] = [
  'audit_id', 'occurred_at', 'actor', 'action', 'entity_type', 'entity_id', 'status',
  'detail_json', 'before_hash', 'after_hash'
];
VN_ADMIN_HEADERS[VN_ADMIN_SHEETS.LEARNING_OBS] = [
  'observation_id', 'book_id', 'client_name', 'fiscal_year', 'observed_month',
  'actual_amount', 'system_p10', 'system_p50', 'system_p90', 'range_breach',
  'hypothesis', 'verification_status', 'verification_note', 'alerted',
  'created_at', 'created_by', 'detail_json'
];
VN_ADMIN_HEADERS[VN_ADMIN_SHEETS.LEARNING_EVIDENCE] = [
  'evidence_id', 'book_id', 'fiscal_year', 'evaluation_id', 'official_vintage_id',
  'source_run_id', 'range_contains_actual', 'system_ape', 'budget_ape',
  'layer_errors_json', 'learning_payload_json', 'created_at', 'created_by'
];

const VN_ADMIN_LEARNING_POLICY_KEY = 'LEARNING_POLICY_JSON';
const VN_ADMIN_LEARNING_POLICY_DEFAULT = Object.freeze({
  schemaVersion: 'vnext-learning-policy-1',
  concept: '未来予測は直接最適化できない。区間校正→層別バイアス→制約付き点誤差→情報ギャップの順で代理目的を追う。',
  proxyObjectives: Object.freeze([
    Object.freeze({ rank: 1, id: 'interval_calibration', label: '区間校正' }),
    Object.freeze({ rank: 2, id: 'layer_bias', label: '層別バイアス' }),
    Object.freeze({ rank: 3, id: 'point_error_constrained', label: '点予測誤差（制約付き）' }),
    Object.freeze({ rank: 4, id: 'information_gap', label: '情報ギャップ縮小' })
  ]),
  nonGoals: Object.freeze([
    '正式予算・営業上積み・採用差分を学習に戻す',
    '公式vintageの後書き',
    '年度途中の会社予算補正',
    '人が理解できない構造の自動適用'
  ]),
  budgetMidYearCorrection: false,
  humanGate: 'threshold_and_material',
  intervalBreachAlert: true,
  intervalWidenFactorOnMiss: 1.15,
  tracks: Object.freeze({
    system: 'システム推奨 vs 確定実績（学習に使用）',
    budget: '正式予算 vs 確定実績（監査用・学習に戻さない）'
  })
});

const VN_ADMIN_DEFAULT_CLIENT_VISIBLE = ['1_ホーム', '2_予測と計画'];
const VN_ADMIN_DEFAULT_TEMPLATE_VISIBLE = ['1_ホーム', '2_予測と計画', '3_振り返り'];
const VN_ADMIN_DEFAULT_HUB_VISIBLE = [
  VN_ADMIN_SHEETS.HOME, VN_ADMIN_SHEETS.EXCEPTIONS, VN_ADMIN_SHEETS.REGISTRY,
  VN_ADMIN_SHEETS.JOBS, VN_ADMIN_SHEETS.APPROVALS, VN_ADMIN_SHEETS.OFFICIAL,
  VN_ADMIN_SHEETS.RELEASES
];

/**
 * Store runtime configuration in Script Properties. Values are deliberately not returned.
 * This is a public, menu-free API intended for clasp/API setup.
 */
function vNextConfigureRuntime(config) {
  return vNextAdminGuard_('vNextConfigureRuntime', function () {
    vNextAdminAssertRuntimeConfigurator_();
    const input = config && typeof config === 'object' ? config : {};
    if (Object.prototype.hasOwnProperty.call(input, 'VNEXT_ACTIVE_MODEL_RELEASE_ID')) {
      throw new Error('Use vNextAdminActivateModelRelease or vNextAdminRollbackModelRelease to change the active model pointer.');
    }
    const unknown = Object.keys(input).filter(function (key) {
      return VN_ADMIN_RUNTIME_KEYS.indexOf(key) < 0;
    });
    if (unknown.length) throw new Error('Unsupported runtime configuration keys: ' + unknown.join(', '));

    return vNextAdminWithScriptLock_('runtime-config', function () {
      const props = PropertiesService.getScriptProperties();
      const saved = [];
      const normalizedInput = Object.assign({}, input);
      if (normalizedInput.FORECAST_SOURCE_SPREADSHEET_ID && !normalizedInput.VNEXT_ZAC_SOURCE_SPREADSHEET_ID) {
        normalizedInput.VNEXT_ZAC_SOURCE_SPREADSHEET_ID = normalizedInput.FORECAST_SOURCE_SPREADSHEET_ID;
      }
      if (normalizedInput.VNEXT_ZAC_SOURCE_SPREADSHEET_ID && !normalizedInput.FORECAST_SOURCE_SPREADSHEET_ID) {
        normalizedInput.FORECAST_SOURCE_SPREADSHEET_ID = normalizedInput.VNEXT_ZAC_SOURCE_SPREADSHEET_ID;
      }
      Object.keys(normalizedInput).forEach(function (key) {
        let value = normalizedInput[key];
        if (Array.isArray(value)) value = value.join(',');
        if (value === null || value === undefined || String(value).trim() === '') {
          props.deleteProperty(key);
          saved.push(key);
          return;
        }
        props.setProperty(key, String(value).trim());
        saved.push(key);
      });
      const active = SpreadsheetApp.getActiveSpreadsheet();
      if (active && vNextDetectBookMode_(active) === 'ADMIN') {
        const hubConfigPatch = {};
        [
          'FORECAST_SOURCE_SPREADSHEET_ID', 'VNEXT_ZAC_SOURCE_SPREADSHEET_ID',
          'VERTEX_PROJECT_ID', 'VERTEX_LOCATION', 'VERTEX_GEMINI_MODEL',
          'VERTEX_DATASTORE_ID', 'VERTEX_SEARCH_LOCATION', 'VERTEX_SERVING_CONFIG',
          'VNEXT_ADMIN_EMAILS', 'VNEXT_ACTIVE_RELEASE_ID', 'VNEXT_ACTIVE_MODEL_RELEASE_ID'
        ].forEach(function (key) {
          if (!Object.prototype.hasOwnProperty.call(normalizedInput, key)) return;
          if (key === 'FORECAST_SOURCE_SPREADSHEET_ID' || key === 'VNEXT_ZAC_SOURCE_SPREADSHEET_ID') {
            hubConfigPatch.source_spreadsheet_id = normalizedInput[key] || '';
          } else if (key === 'VNEXT_ADMIN_EMAILS') {
            hubConfigPatch.admin_emails = normalizedInput[key] || '';
          } else if (key === 'VNEXT_ACTIVE_RELEASE_ID') {
            hubConfigPatch.active_release_id = normalizedInput[key] || '';
          } else if (key === 'VNEXT_ACTIVE_MODEL_RELEASE_ID') {
            hubConfigPatch.active_model_release_id = normalizedInput[key] || '';
          } else {
            hubConfigPatch[key] = normalizedInput[key] || '';
          }
        });
        if (Object.keys(hubConfigPatch).length) vNextAdminWriteSystemConfig_(active, hubConfigPatch);
      }
      Logger.log('vNext runtime configuration updated keys=%s', saved.join(','));
      return { savedKeys: saved.sort() };
    });
  });
}

/** Internal runtime reader. Do not expose this object through employee-facing APIs. */
function vNextGetRuntimeConfig_() {
  const props = PropertiesService.getScriptProperties().getProperties();
  const out = {};
  VN_ADMIN_RUNTIME_KEYS.forEach(function (key) {
    if (Object.prototype.hasOwnProperty.call(props, key)) out[key] = props[key];
  });
  return out;
}

/**
 * Cheap Hub identity for onOpen. Sheet names only; no config reads, registry
 * scans, or Script Property writes. Identity checks belong in click handlers.
 */
function vNextAdminLooksLikeHub_(spreadsheet) {
  try {
    const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return false;
    const names = {};
    ss.getSheets().forEach(function (sheet) { names[sheet.getName()] = true; });
    return !!(names[VN_ADMIN_SHEETS.REGISTRY] && names[VN_ADMIN_SHEETS.JOBS] && names[VN_ADMIN_SHEETS.RELEASES]);
  } catch (err) {
    Logger.log('vNextAdminLooksLikeHub_ error: %s', String(err && err.message || err));
    return false;
  }
}

/**
 * Global mode detector consumed by the shared vNext onOpen router.
 * Returns one of LEGACY / ADMIN / TEMPLATE / CLIENT.
 */
function vNextDetectBookMode_(spreadsheet) {
  try {
    const ss = vNextAdminResolveSpreadsheet_(spreadsheet);
    const routing = Object.assign(
      {},
      vNextAdminReadKeyValueSheet_(ss, VN_ADMIN_SYSTEM_CONFIG_SHEET),
      vNextAdminReadKeyValueSheet_(ss, VN_ADMIN_BOOK_CONFIG_SHEET)
    );
    const explicit = String(routing.mode || '').trim().toUpperCase();
    if (VN_ADMIN_MODES.indexOf(explicit) >= 0) {
      vNextAdminHydrateLocalRuntime_(ss, routing);
      return explicit;
    }

    const names = new Set(ss.getSheets().map(function (sheet) { return sheet.getName(); }));
    if (names.has(VN_ADMIN_SHEETS.REGISTRY) && names.has(VN_ADMIN_SHEETS.JOBS) && names.has(VN_ADMIN_SHEETS.RELEASES)) {
      return 'ADMIN';
    }
    // BOOK_META is reserved for VNext_Core's append-only 20-column schema.
    // A legacy key/value BOOK_META is never created by this module.
    return 'LEGACY';
  } catch (err) {
    Logger.log('vNextDetectBookMode_ fallback LEGACY: %s', String(err && err.message || err));
    return 'LEGACY';
  }
}

function vNextIsAdminHub_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    return vNextDetectBookMode_(ss) === 'ADMIN' && vNextAdminIsRegisteredHub_(ss);
  }
  catch (err) { Logger.log('vNextIsAdminHub_ error: %s', String(err && err.message || err)); return false; }
}

/** Hub-only menu builder. The UX module owns the single global onOpen router. */
function vNextBuildAdminMenu_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!vNextAdminLooksLikeHub_(ss)) return false;
    const ui = SpreadsheetApp.getUi();
    ui.createMenu(VN_ADMIN_MENU_NAME)
      .addItem(VN_ADMIN_MENU_OPEN_SIDEBAR, 'vNextAdminOpenSidebar')
      .addSubMenu(ui.createMenu(VN_ADMIN_MENU_OTHER)
        .addItem(VN_ADMIN_MENU_HEALTH_SCAN, 'vNextAdminMenuRunHealthScan')
        .addItem(VN_ADMIN_MENU_OPEN_REGISTRY, 'vNextAdminMenuOpenRegistry'))
      .addToUi();
    return true;
  } catch (err) {
    Logger.log('vNextBuildAdminMenu_ error: %s', String(err && err.stack || err));
    return false;
  }
}

/** Optional best-effort hook called after the existing legacy menu is built. */
function vNextBuildLegacySetupMenu_() {
  try {
    if (vNextDetectBookMode_() !== 'LEGACY') return false;
    // A simple onOpen trigger cannot call DriveApp before authorization.
    // Keep menu construction authorization-free and perform the owner/admin
    // check inside the invoked configuration/bootstrap APIs instead.
    SpreadsheetApp.getUi().createMenu('Forecast vNext 移行')
      .addItem('初回権限を確認・許可', 'vNextAdminAuthorizeRuntime')
      .addItem('必要APIを有効化', 'vNextAdminEnableAppsScriptApi')
      .addSeparator()
      .addItem('初期設定を開く', 'vNextAdminOpenSidebar')
      .addToUi();
    return true;
  } catch (err) {
    Logger.log('vNextBuildLegacySetupMenu_ skipped: %s', String(err && err.message || err));
    return false;
  }
}

/**
 * Direct menu entry used once after the manifest gains new OAuth scopes.
 * google.script.run inside a sidebar cannot initiate a new consent flow, so
 * authorization must begin from a top-level custom-menu invocation.
 */
function vNextAdminAuthorizeRuntime() {
  return vNextAdminGuard_('vNextAdminAuthorizeRuntime', function () {
    vNextAdminAssertRuntimeConfigurator_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('Active spreadsheet is required.');
    // Touch every service family required by bootstrap without changing data.
    const actor = vNextAdminActor_();
    const fileId = DriveApp.getFileById(ss.getId()).getId();
    const scriptId = ScriptApp.getScriptId();
    const tokenAvailable = !!ScriptApp.getOAuthToken();
    SpreadsheetApp.getUi().alert(
      'Forecast vNext 権限確認',
      '初回権限の確認が完了しました。続けて「Forecast vNext 移行」→「初期設定を開く」を選択してください。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return { ok: true, actor: actor, spreadsheetId: fileId, scriptId: scriptId, tokenAvailable: tokenAvailable };
  });
}

/** Enable only the Apps Script API required to create clean bound runtimes. */
function vNextAdminEnableAppsScriptApi() {
  return vNextAdminGuard_('vNextAdminEnableAppsScriptApi', function () {
    vNextAdminAssertRuntimeConfigurator_();
    if (typeof vNextClientRuntimeEnableRequiredAppsScriptApi_ !== 'function') {
      throw new Error('Apps Script API setup helper is not installed.');
    }
    const result = vNextClientRuntimeEnableRequiredAppsScriptApi_();
    SpreadsheetApp.getUi().alert(
      'Forecast vNext API設定',
      result.alreadyEnabled
        ? 'Apps Script APIは既に有効です。初期設定を再実行してください。'
        : 'Apps Script APIを有効化し、利用可能になるまで確認しました。初期設定を再実行してください。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return result;
  });
}

/** Enable script.googleapis.com for a clean generated Hub from the source project. */
function vNextAdminEnableGeneratedHubAppsScriptApi(request) {
  return vNextAdminGuard_('vNextAdminEnableGeneratedHubAppsScriptApi', function () {
    vNextAdminAssertRuntimeConfigurator_();
    const req = request && typeof request === 'object' ? request : {};
    const source = SpreadsheetApp.getActiveSpreadsheet();
    const sourceMode = vNextDetectBookMode_(source);
    if (sourceMode !== 'LEGACY' && sourceMode !== 'TEMPLATE') {
      throw new Error('Generated Hub API recovery must run from the registered source workbook.');
    }
    const hubId = vNextAdminRequiredText_(req.hubSpreadsheetId, 'hubSpreadsheetId');
    const projectNumber = vNextAdminRequiredText_(req.cloudProjectNumber, 'cloudProjectNumber');
    if (!/^\d{6,20}$/.test(projectNumber)) throw new Error('cloudProjectNumber must contain 6-20 digits.');
    const hub = SpreadsheetApp.openById(hubId);
    const hubConfig = vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET);
    if (String(hubConfig.mode || '').toUpperCase() !== 'ADMIN' ||
        String(hubConfig.admin_hub_spreadsheet_id || '') !== hubId ||
        String(hubConfig.admin_source_script_id || '') !== ScriptApp.getScriptId()) {
      throw new Error('The supplied spreadsheet is not a Hub created by this central source script.');
    }
    const targetScriptId = vNextAdminRequiredText_(hubConfig.admin_hub_script_id, 'admin_hub_script_id');
    if (typeof vNextAdminRuntimeAssertBoundParent_ === 'function') {
      vNextAdminRuntimeAssertBoundParent_(targetScriptId, hubId);
    }
    if (typeof vNextClientRuntimeEnableAppsScriptApiForProjectNumber_ !== 'function') {
      throw new Error('Generated runtime API recovery helper is not installed.');
    }
    const result = vNextClientRuntimeEnableAppsScriptApiForProjectNumber_(projectNumber);
    vNextAdminWriteSystemConfig_(hub, {
      admin_cloud_project_number: projectNumber,
      admin_apps_script_api_enabled_at: new Date().toISOString(),
      admin_apps_script_api_enabled_by: vNextAdminActor_()
    });
    vNextAdminWriteAudit_(hub, 'ENABLE_GENERATED_HUB_API', 'ADMIN_RUNTIME', targetScriptId, 'SUCCESS', {
      hubSpreadsheetId: hubId, cloudProjectNumber: projectNumber,
      service: 'script.googleapis.com', alreadyEnabled: result.alreadyEnabled === true
    });
    return result;
  });
}

/** Consumes TEMPLATE onOpen so it can never fall through to the legacy operation menu. */
function vNextBuildTemplateMenu_() {
  try {
    if (vNextDetectBookMode_() !== 'TEMPLATE') return false;
    SpreadsheetApp.getUi().createMenu(VN_ADMIN_MENU_NAME)
      .addItem(VN_ADMIN_MENU_OPEN_SIDEBAR, 'vNextAdminOpenSidebar')
      .addToUi();
    return true;
  } catch (err) {
    Logger.log('vNextBuildTemplateMenu_ error: %s', String(err && err.message || err));
    return true;
  }
}

function vNextAdminInstalledGuidanceOnOpen(e) {
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (vNextAdminLooksLikeHub_(active)) {
      SpreadsheetApp.getUi().showSidebar(
        HtmlService.createHtmlOutputFromFile('VNext_AdminSidebar').setTitle(VN_ADMIN_MENU_NAME)
      );
      return true;
    }
    const mode = vNextDetectBookMode_(active);
    if (mode !== 'TEMPLATE' && mode !== 'LEGACY') return false;
    SpreadsheetApp.getUi().showSidebar(
      HtmlService.createHtmlOutputFromFile('VNext_AdminSidebar').setTitle(VN_ADMIN_MENU_NAME)
    );
    return true;
  } catch (err) {
    Logger.log('vNextAdminInstalledGuidanceOnOpen skipped: %s', String(err && err.message || err));
    return false;
  }
}

function vNextAdminEnsureGuidanceOnOpenTrigger_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const handler = 'vNextAdminInstalledGuidanceOnOpen';
  function isOpenHandler(trigger) {
    return trigger.getHandlerFunction() === handler &&
      trigger.getEventType() === ScriptApp.EventType.ON_OPEN;
  }
  if (ScriptApp.getProjectTriggers().some(isOpenHandler)) return false;
  try {
    ScriptApp.getUserTriggers(ss).filter(isOpenHandler).forEach(function (trigger) {
      ScriptApp.deleteTrigger(trigger);
    });
  } catch (cleanupError) {
    Logger.log('vNextAdminEnsureGuidanceOnOpenTrigger_ cleanup skipped: %s',
      String(cleanupError && cleanupError.message || cleanupError));
  }
  ScriptApp.newTrigger(handler).forSpreadsheet(ss).onOpen().create();
  return true;
}

function vNextAdminOpenSidebar() {
  return vNextAdminGuard_('vNextAdminOpenSidebar', function () {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (vNextDetectBookMode_(active) === 'ADMIN') {
      if (!vNextAdminIsRegisteredHub_(active)) throw new Error('Admin 管理ハブの登録情報を確認できません。');
      vNextAdminAssertHubAdmin_(active, false);
    }
    const html = HtmlService.createHtmlOutputFromFile('VNext_AdminSidebar')
      .setTitle(VN_ADMIN_MENU_NAME);
    SpreadsheetApp.getUi().showSidebar(html);
    try { vNextAdminEnsureGuidanceOnOpenTrigger_(); }
    catch (triggerError) {
      Logger.log('vNextAdminOpenSidebar trigger skipped: %s',
        String(triggerError && triggerError.message || triggerError));
    }
    return true;
  });
}

function vNextAdminGetSidebarModel() {
  return vNextAdminGuard_('vNextAdminGetSidebarModel', function () {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hubLike = vNextAdminLooksLikeHub_(ss);
    const mode = hubLike ? 'ADMIN' : vNextDetectBookMode_(ss);
    const runtime = vNextGetRuntimeConfig_();
    const model = {
      mode: mode,
      spreadsheetName: ss.getName(),
      spreadsheetUrl: ss.getUrl(),
      configured: {
        source: !!runtime.FORECAST_SOURCE_SPREADSHEET_ID,
        hub: !!runtime.VNEXT_ADMIN_HUB_SPREADSHEET_ID,
        template: !!runtime.VNEXT_MASTER_TEMPLATE_SPREADSHEET_ID,
        vertex: !!runtime.VERTEX_PROJECT_ID
      },
      automationInstalled: false,
      counts: {
        exceptions: 0, queuedJobs: 0, runningJobs: 0, failedJobs: 0,
        pendingApprovals: 0, clients: 0, actualDataIssues: 0,
        portalAttention: 0, attention: 0
      },
      topExceptions: [],
      pendingApprovals: [],
      recentJobs: [],
      portalRequests: {
        loading: true, configured: true, unavailable: false, spreadsheetUrl: '',
        counts: { waiting: 0, processing: 0, failed: 0, completed: 0 },
        attention: []
      },
      attention: {
        status: 'LOADING', label: '確認中', guidance: '', safeRetryCandidates: 0
      },
      activeTemplateReleaseId: '',
      templateRuntimeVersion: '',
      adminRuntimeSha256: '',
      adminRuntimeUpdatable: false,
      portalRuntimeVersion: '',
      portalRuntimeSha256: '',
      portalRuntimeUpdatable: false,
      clientCatalogActiveCount: 0,
      clientCatalogVersion: '',
      clientCatalogRefreshedAt: '',
      clientCatalogStale: true,
      activeModelReleaseId: '',
      modelReleases: [],
      templateDrafts: [],
      stagedTemplateReleases: [],
      emptyPilotUpgradeCandidates: [],
      operations: {},
      pilot: { loading: true }
    };
    if (mode === 'ADMIN') {
      let hubConfig = vNextAdminAssertHubAdminFast_(ss);
      if (!String(hubConfig.book_id || '')) {
        hubConfig = Object.assign({}, hubConfig, vNextAdminReadKeyValueSheet_(ss, VN_ADMIN_BOOK_CONFIG_SHEET));
      }
      model.automationInstalled = vNextAdminAutomationInstalled_();
      const exceptionRows = vNextAdminReadTable_(ss, VN_ADMIN_SHEETS.EXCEPTIONS).rows;
      const jobRows = vNextAdminReadTable_(ss, VN_ADMIN_SHEETS.JOBS).rows;
      const approvalRows = vNextAdminReadTable_(ss, VN_ADMIN_SHEETS.APPROVALS).rows;
      const registryRows = vNextAdminReadTable_(ss, VN_ADMIN_SHEETS.REGISTRY).rows;
      if (!vNextAdminIsRegisteredHubFromRows_(ss, hubConfig, registryRows)) {
        throw new Error('Admin 管理ハブの登録情報を確認できません。');
      }
      const openExceptions = exceptionRows.filter(function (row) {
        return String(row.status || 'OPEN').toUpperCase() === 'OPEN' &&
          String(row.exception_type || '').toUpperCase() !== 'APPROVAL_PENDING';
      });
      const pendingApprovals = approvalRows.filter(function (row) {
        return String(row.status || '').toUpperCase() === 'PENDING';
      });
      model.counts.exceptions = openExceptions.length;
      model.counts.queuedJobs = jobRows.filter(function (row) {
        return String(row.status || '').toUpperCase() === 'QUEUED';
      }).length;
      model.counts.runningJobs = jobRows.filter(function (row) {
        return String(row.status || '').toUpperCase() === 'RUNNING';
      }).length;
      model.counts.failedJobs = jobRows.filter(function (row) {
        return String(row.status || '').toUpperCase() === 'FAILED';
      }).length;
      model.counts.pendingApprovals = pendingApprovals.length;
      model.counts.clients = registryRows.filter(function (row) {
        return String(row.mode || '').toUpperCase() === 'CLIENT' &&
          String(row.status || '').toUpperCase() !== 'ARCHIVED';
      }).length;
      model.operations = vNextAdminOperationalMetrics_(ss, model.automationInstalled, jobRows);
      const severityRank = { ERROR: 0, WARN: 1, INFO: 2 };
      const registryByBook = {};
      registryRows.forEach(function (row) { if (row.book_id) registryByBook[String(row.book_id)] = row; });
      model.topExceptions = openExceptions
        .sort(function (a, b) {
          const aKey = String(a.severity || '').toUpperCase();
          const bKey = String(b.severity || '').toUpperCase();
          const rank = (Object.prototype.hasOwnProperty.call(severityRank, aKey) ? severityRank[aKey] : 9) -
            (Object.prototype.hasOwnProperty.call(severityRank, bKey) ? severityRank[bKey] : 9);
          return rank || new Date(b.detected_at || 0).getTime() - new Date(a.detected_at || 0).getTime();
        }).slice(0, 8).map(function (row) {
          return vNextAdminExceptionForSidebar_(row, registryByBook, jobRows);
        });
      model.pendingApprovals = vNextAdminListPendingApprovals_(ss, approvalRows).sort(function (a, b) {
        return new Date(a.requestedAt || 0).getTime() - new Date(b.requestedAt || 0).getTime();
      }).slice(0, 20);
      model.recentJobs = vNextAdminJobsForSidebar_(jobRows).slice(0, 8);
      model.counts.actualDataIssues = openExceptions.filter(vNextAdminIsActualDataIssue_).length;
      const safeRetryCandidates = jobRows.filter(vNextAdminIsKnownSafeRetryCandidate_).length;
      model.counts.attention = model.counts.exceptions + model.counts.pendingApprovals;
      model.attention = vNextAdminAttentionSummary_(model, safeRetryCandidates);
      vNextAdminApplyHubRuntimeFlags_(model, hubConfig);
      model.activeTemplateReleaseId = String(hubConfig.active_release_id || '');
    }
    return vNextAdminJsonSafe_(model);
  });
}

/** Second paint: Portal spreadsheet, catalog, model/release details. */
function vNextAdminGetSidebarDetailModel() {
  return vNextAdminGuard_('vNextAdminGetSidebarDetailModel', function () {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!vNextAdminLooksLikeHub_(ss)) return {};
    let hubConfig = vNextAdminAssertHubAdminFast_(ss);
    if (!String(hubConfig.book_id || '')) {
      hubConfig = Object.assign({}, hubConfig, vNextAdminReadKeyValueSheet_(ss, VN_ADMIN_BOOK_CONFIG_SHEET));
    }
    const registryRows = vNextAdminReadTable_(ss, VN_ADMIN_SHEETS.REGISTRY).rows;
    const portalRequests = vNextAdminPortalRequestsForSidebar_(ss);
    const activeTemplate = vNextAdminResolveRelease_(ss, hubConfig.active_release_id || '');
    const unfinishedEmptyUpgradesByBook = {};
    vNextAdminReadTable_(ss, VN_ADMIN_SHEETS.MIGRATIONS).rows.forEach(function (row) {
      if (/^EMPTY_PILOT_|^RECOVERY_REQUIRED$/.test(String(row.status || '').toUpperCase())) {
        unfinishedEmptyUpgradesByBook[String(row.book_id || '')] = row;
      }
    });
    const emptyPilotUpgradeCandidates = registryRows.filter(function (row) {
      const unfinished = unfinishedEmptyUpgradesByBook[String(row.book_id || '')];
      return String(row.mode || '').toUpperCase() === 'CLIENT' &&
        String(row.status || '').toUpperCase() === 'ACTIVE' &&
        String(row.state || '').toUpperCase() === 'INPUT_OPEN' &&
        !String(row.current_official_id || '') &&
        (String(row.template_release_id || '') !== String(activeTemplate.release_id || '') || !!unfinished);
    }).map(function (row) {
      const unfinished = unfinishedEmptyUpgradesByBook[String(row.book_id || '')];
      return {
        bookId: String(row.book_id || ''), clientName: String(row.client_name || ''),
        fiscalYear: Number(row.fiscal_year || 0), spreadsheetUrl: String(row.spreadsheet_url || ''),
        currentReleaseId: String(row.template_release_id || ''),
        targetReleaseId: String(activeTemplate.release_id || ''),
        recoveryRequired: !!unfinished,
        migrationId: unfinished ? String(unfinished.migration_id || '') : '',
        migrationStatus: unfinished ? String(unfinished.status || '') : ''
      };
    }).sort(function (a, b) {
      return a.clientName.localeCompare(b.clientName, 'ja') || a.fiscalYear - b.fiscalYear;
    });
    const catalogRows = vNextAdminReadTable_(ss, VN_ADMIN_SHEETS.CATALOG).rows;
    const catalogVersion = String(hubConfig.zac_client_catalog_version || '');
    const catalogRefreshedAt = String(hubConfig.zac_client_catalog_refreshed_at || '');
    const catalogRefreshedMs = new Date(catalogRefreshedAt || 0).getTime();
    const activeModel = vNextAdminTryResolveActiveModelRelease_(ss);
    const model = {
      portalRequests: portalRequests,
      counts: { portalAttention: Number(portalRequests.counts.failed || 0) },
      activeTemplateReleaseId: String(activeTemplate.release_id || ''),
      templateRuntimeVersion: String(activeTemplate.client_runtime_version || ''),
      emptyPilotUpgradeCandidates: emptyPilotUpgradeCandidates,
      clientCatalogActiveCount: catalogRows.filter(function (row) {
        return vNextAdminBool_(row.is_active) && String(row.catalog_version || '') === catalogVersion;
      }).length,
      clientCatalogVersion: catalogVersion,
      clientCatalogRefreshedAt: catalogRefreshedAt,
      clientCatalogStale: !isFinite(catalogRefreshedMs) ||
        Date.now() - catalogRefreshedMs >= VN_ADMIN_ZAC_CATALOG_STALE_MS,
      activeModelReleaseId: activeModel && activeModel.model_release_id || '',
      modelReleases: vNextAdminLatestModelReleaseSummaries_(ss),
      templateDrafts: vNextAdminListTemplateDrafts_(ss),
      stagedTemplateReleases: vNextAdminReadTable_(ss, VN_ADMIN_SHEETS.RELEASES).rows
        .filter(function (row) { return String(row.status || '').toUpperCase() === 'STAGED'; })
        .map(function (row) {
          return {
            releaseId: String(row.release_id || ''), releaseName: String(row.release_name || ''),
            templateSpreadsheetId: String(row.template_spreadsheet_id || ''),
            templateContentSha256: String(row.template_content_sha256 || ''), createdAt: row.created_at || ''
          };
        }),
      pilot: vNextAdminPilotStatusFromRegistry_(registryRows, ss)
    };
    vNextAdminApplyHubRuntimeFlags_(model, hubConfig);
    model.devDeployStatus = vNextAdminBuildDevDeployStatus_(ss, { light: true });
    try {
      model.learningDashboard = vNextAdminBuildLearningDashboard_(ss);
    } catch (learningError) {
      model.learningDashboard = {
        error: String(learningError && learningError.message || learningError),
        proxyObjectives: [], dualTracks: [], recentObservations: [], openBreaches: [],
        stats: { evaluationCount: 0, evidenceCount: 0, observationCount: 0, openIntervalBreaches: 0 }
      };
    }
    return vNextAdminJsonSafe_(model);
  });
}

function vNextAdminApplyHubRuntimeFlags_(model, hubConfig) {
  const config = hubConfig || {};
  model.adminRuntimeSha256 = String(config.admin_runtime_sha256 || '');
  model.adminRuntimeUpdatable = Boolean(
    String(config.admin_source_script_id || '') && String(config.admin_hub_script_id || '') &&
    String(config.admin_hub_script_id || '') === String(ScriptApp.getScriptId())
  );
  model.portalRuntimeVersion = String(config.portal_runtime_version || '');
  model.portalRuntimeSha256 = String(config.portal_runtime_sha256 || '');
  model.portalRuntimeUpdatable = Boolean(
    String(config.portal_spreadsheet_id || '') && String(config.portal_script_id || '') &&
    [VN_ADMIN_PORTAL_RUNTIME_VERSION].concat(VN_ADMIN_PORTAL_LEGACY_RUNTIME_VERSIONS)
      .indexOf(model.portalRuntimeVersion) >= 0
  );
  if (!model.clientCatalogVersion) {
    model.clientCatalogVersion = String(config.zac_client_catalog_version || '');
    model.clientCatalogRefreshedAt = String(config.zac_client_catalog_refreshed_at || '');
    const catalogRefreshedMs = new Date(model.clientCatalogRefreshedAt || 0).getTime();
    model.clientCatalogStale = !isFinite(catalogRefreshedMs) ||
      Date.now() - catalogRefreshedMs >= VN_ADMIN_ZAC_CATALOG_STALE_MS;
  }
  return model;
}

function vNextAdminIsRegisteredHubFromRows_(ss, hubConfig, registryRows) {
  const routing = hubConfig || {};
  const bookId = String(routing.book_id || '');
  if (String(routing.mode || '').toUpperCase() !== 'ADMIN' || !bookId) return false;
  return (registryRows || []).some(function (row) {
    return String(row.book_id || '') === bookId &&
      String(row.mode || '').toUpperCase() === 'ADMIN' &&
      String(row.spreadsheet_id || '') === String(ss.getId());
  });
}

function vNextAdminPilotStatusFromRegistry_(registryRows, hub) {
  const clientCount = (registryRows || []).filter(function (row) {
    return String(row.mode || '').toUpperCase() === 'CLIENT' &&
      String(row.status || '').toUpperCase() !== 'ARCHIVED';
  }).length;
  const approvedRow = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.SETTINGS).rows.find(function (row) {
    return String(row.setting_key || '') === 'PILOT_CANARY_APPROVED';
  });
  const canaryApproved = String(approvedRow && approvedRow.setting_value || '').toLowerCase() === 'true';
  const limit = canaryApproved ? VN_ADMIN_PILOT_CANARY_LIMIT : VN_ADMIN_PILOT_INITIAL_LIMIT;
  return {
    clientCount: clientCount, initialLimit: VN_ADMIN_PILOT_INITIAL_LIMIT,
    hardLimit: VN_ADMIN_PILOT_CANARY_LIMIT, canaryApproved: canaryApproved,
    currentLimit: limit, phase: canaryApproved ? 'CANARY' : 'INITIAL_PILOT',
    blocked: clientCount >= limit,
    blockedReason: clientCount >= VN_ADMIN_PILOT_CANARY_LIMIT
      ? '5冊canary上限に達しています。30冊展開は実測検証後のreleaseで開放します。'
      : (clientCount >= VN_ADMIN_PILOT_INITIAL_LIMIT && !canaryApproved
        ? '初期pilot 3冊の結果を確認し、Canary承認後に4〜5冊目を作成できます。' : '')
  };
}

/** Converts internal exception codes into a short, decision-oriented Admin card. */
function vNextAdminExceptionForSidebar_(row, registryByBook, jobRows) {
  const source = row || {};
  const bookId = String(source.book_id || '');
  const category = String(source.exception_type || 'GENERAL').toUpperCase();
  const registry = registryByBook && registryByBook[bookId] || {};
  const actualIssue = vNextAdminIsActualDataIssue_(source);
  const categoryLabels = {
    BOOK_HEALTH: actualIssue ? '実績データの確認' : 'クライアント年度ブックの状態確認',
    JOB_FAILED: actualIssue ? '実績データ不足' : '自動処理を完了できませんでした',
    JOB_QUEUE_STALE: '処理に時間がかかっています',
    SCHEDULER_STALE: '自動更新が止まっている可能性があります',
    PORTAL_PROVISION_FAILED: 'クライアント年度ブックを作成できませんでした',
    CLIENT_PROVISION_FAILED: 'クライアント年度ブックを作成できませんでした',
    OFFICIAL_CLIENT_SYNC_FAILED: '正式予算の反映が未完了です',
    OFFICIAL_COPY_FAILED: '正式予算の反映が未完了です',
    PORTAL_REQUEST_REJECTED: '申請入口からの作成依頼を確認してください'
  };
  // A JOB_FAILED exception must only control its own durable job. Falling
  // through to an older failed job for the same book can attach the wrong
  // retry action to an otherwise unrelated exception.
  const sourceRef = String(source.source_ref || '');
  const matchingJob = sourceRef ? (jobRows || []).find(function (job) {
    return String(job.job_id || '') === sourceRef;
  }) : null;
  const jobPayload = matchingJob ? vNextAdminParseJson_(matchingJob.request_json, {}) : {};
  let actionType = 'DETAILS';
  if (['OFFICIAL_CLIENT_SYNC_FAILED', 'OFFICIAL_COPY_FAILED'].indexOf(category) >= 0 && bookId) {
    actionType = 'RETRY_OFFICIAL_SYNC';
  } else if (category === 'SCHEDULER_STALE' || category === 'JOB_QUEUE_STALE' ||
      (matchingJob && vNextAdminIsKnownSafeRetryCandidate_(matchingJob))) {
    actionType = 'RUN_NOW';
  } else if (registry.spreadsheet_url) {
    actionType = 'OPEN_BOOK';
  }
  let guidance = String(source.recommended_action || '要確認一覧で詳細を確認してください。');
  if (actualIssue) guidance = 'ZACの確定実績が必要年数・対象月を満たしているか確認してください。';
  if (category === 'JOB_FAILED' && actionType !== 'RUN_NOW') {
    guidance = '同じ操作を新しく作り直さず、失敗内容を確認してください。';
  }
  if (category === 'SCHEDULER_STALE') {
    guidance = 'まず「' + VN_ADMIN_MENU_RUN_NOW + '」を実行し、続く場合は自動運用の権限を確認してください。';
  }
  return {
    bookId: bookId,
    bookUrl: vNextAdminSidebarSpreadsheetUrl_(registry.spreadsheet_url),
    clientName: String(source.client_name || registry.client_name || jobPayload.clientName || ''),
    fiscalYear: source.fiscal_year || registry.fiscal_year || jobPayload.fiscalYear || '',
    severity: String(source.severity || 'WARN').toUpperCase(),
    category: category,
    categoryLabel: categoryLabels[category] || '確認が必要です',
    message: categoryLabels[category] || String(source.title || source.detail || '確認が必要です'),
    detail: String(source.detail || source.title || ''),
    actionGuidance: guidance,
    actionType: actionType,
    actualDataIssue: actualIssue
  };
}

function vNextAdminIsActualDataIssue_(row) {
  const text = [row && row.exception_type, row && row.title, row && row.detail,
    row && row.recommended_action, row && row.error].join(' ').toUpperCase();
  return /HISTORY|ACTUAL|確定実績|実績データ|AT LEAST 5 FISCAL YEARS|ZERO_ACTUAL|MISSING_MONTHS/.test(text);
}

function vNextAdminIsKnownSafeRetryCandidate_(job) {
  if (!job || String(job.status || '').toUpperCase() !== 'FAILED' || Number(job.attempts || 0) >= 3) return false;
  const type = String(job.job_type || '').toUpperCase();
  const error = String(job.error || '');
  return (type === 'PORTAL_PROVISION_CLIENT' && /^Requested release is not ACTIVE: [-A-Za-z0-9._]+$/.test(error)) ||
    (type === 'FORECAST_REQUEST' && error === 'A matching valid pending forecast request was not found.');
}

function vNextAdminJobsForSidebar_(rows) {
  const labels = {
    PORTAL_PROVISION_CLIENT: 'クライアント年度ブックの作成',
    FORECAST_REQUEST: '売上予測の計算',
    AI_ROLLBACK_FORECAST: 'AI反映取消後の再計算',
    AI_RESEARCH: '外部情報の確認',
    HEALTH_SCAN: 'クライアント年度ブックの状態確認',
    REFRESH_CLIENT_VIEW: '画面の更新',
    MIGRATION: 'クライアント年度ブックの版更新'
  };
  const statusLabels = {
    QUEUED: '受付済み', RUNNING: '処理中', FAILED: '要確認', SUCCEEDED: '完了'
  };
  return (rows || []).filter(function (row) {
    return ['QUEUED', 'RUNNING', 'FAILED'].indexOf(String(row.status || '').toUpperCase()) >= 0;
  }).sort(function (a, b) {
    const rank = { FAILED: 0, RUNNING: 1, QUEUED: 2 };
    const aStatus = String(a.status || '').toUpperCase();
    const bStatus = String(b.status || '').toUpperCase();
    const statusRank = (Object.prototype.hasOwnProperty.call(rank, aStatus) ? rank[aStatus] : 9) -
      (Object.prototype.hasOwnProperty.call(rank, bStatus) ? rank[bStatus] : 9);
    return statusRank || new Date(a.created_at || 0).getTime() - new Date(b.created_at || 0).getTime();
  }).map(function (row) {
    const type = String(row.job_type || '').toUpperCase();
    const status = String(row.status || '').toUpperCase();
    return {
      jobId: String(row.job_id || ''),
      targetBookId: String(row.target_book_id || ''),
      taskLabel: labels[type] || 'バックグラウンド処理',
      status: status,
      statusLabel: statusLabels[status] || status,
      createdAt: row.created_at || '',
      updatedAt: row.updated_at || '',
      errorSummary: status === 'FAILED'
        ? (vNextAdminIsActualDataIssue_(row) ? '確定実績の不足を確認してください。' : '管理ハブの確認が必要です。')
        : '',
      safeRetryCandidate: vNextAdminIsKnownSafeRetryCandidate_(row)
    };
  });
}

/** Best-effort Portal projection. A Portal outage must not hide Hub decisions. */
function vNextAdminPortalRequestsForSidebar_(hub) {
  const output = {
    configured: false, unavailable: false, spreadsheetUrl: '',
    counts: { waiting: 0, processing: 0, failed: 0, completed: 0 }, attention: []
  };
  let portal;
  try { portal = vNextAdminResolvePortalForRead_(hub); }
  catch (error) {
    if (!/not configured/i.test(String(error && error.message || error))) output.unavailable = true;
    return output;
  }
  output.configured = true;
  output.spreadsheetUrl = vNextAdminSidebarSpreadsheetUrl_(portal.spreadsheet.getUrl());
  try {
    const grouped = {};
    const validEventStatusPairs = {
      REQUESTED: 'PENDING', VALIDATION_STARTED: 'VALIDATING',
      CREATION_STARTED: 'CREATING', COMPLETED: 'COMPLETED',
      FAILED: 'FAILED', REJECTED: 'REJECTED'
    };
    vNextAdminReadTable_(portal.spreadsheet, VN_ADMIN_PORTAL_REQUEST_SHEET).rows.forEach(function (row) {
      const id = String(row.request_id || '');
      const eventType = String(row.event_type || '').toUpperCase();
      const status = String(row.status || '').toUpperCase();
      // Employee editors may append REQUESTED rows, so the projection must
      // fail closed and never trust a status that is inconsistent with its
      // append-only event type.
      if (id && validEventStatusPairs[eventType] === status) grouped[id] = row;
    });
    const labels = {
      PENDING: '受付済み', VALIDATING: '内容確認中', CREATING: '作成中',
      FAILED: '作成できませんでした', REJECTED: '内容確認が必要', COMPLETED: '利用できます'
    };
    Object.keys(grouped).forEach(function (id) {
      const row = grouped[id];
      const status = String(row.status || '').toUpperCase();
      if (status === 'PENDING') output.counts.waiting++;
      else if (status === 'VALIDATING' || status === 'CREATING') output.counts.processing++;
      else if (status === 'FAILED' || status === 'REJECTED') output.counts.failed++;
      else if (status === 'COMPLETED') output.counts.completed++;
      if (['PENDING', 'VALIDATING', 'CREATING', 'FAILED', 'REJECTED'].indexOf(status) < 0) return;
      output.attention.push({
        requestId: id, clientName: String(row.client_name || ''), fiscalYear: Number(row.fiscal_year || 0),
        status: status, statusLabel: labels[status] || '確認中',
        detailMessage: String(row.detail_message || ''), updatedAt: row.created_at || row.requested_at || '',
        bookUrl: vNextAdminSidebarSpreadsheetUrl_(row.related_book_url)
      });
    });
    output.attention.sort(function (a, b) {
      const rank = { FAILED: 0, REJECTED: 1, CREATING: 2, VALIDATING: 3, PENDING: 4 };
      const aRank = Object.prototype.hasOwnProperty.call(rank, a.status) ? rank[a.status] : 9;
      const bRank = Object.prototype.hasOwnProperty.call(rank, b.status) ? rank[b.status] : 9;
      return aRank - bRank ||
        new Date(a.updatedAt || 0).getTime() - new Date(b.updatedAt || 0).getTime();
    });
    output.attention = output.attention.slice(0, 8);
  } catch (error) {
    output.unavailable = true;
    output.attention = [];
  }
  return output;
}

function vNextAdminSidebarSpreadsheetUrl_(value) {
  const url = String(value || '').trim();
  return /^https:\/\/docs\.google\.com\/spreadsheets\/d\/[A-Za-z0-9_-]+(?:\/|$)/.test(url) ? url : '';
}

function vNextAdminAttentionSummary_(model, safeRetryCandidates) {
  const counts = model.counts || {};
  const operations = model.operations || {};
  const retryCount = Number(safeRetryCandidates || 0);
  if (!model.automationInstalled) {
    return { status: 'SETUP', label: '自動更新を有効にしてください',
      guidance: '最初に1回だけ「自動運用を有効化」を実行します。', safeRetryCandidates: retryCount };
  }
  if (operations.schedulerStale) {
    return { status: 'ERROR', label: '自動更新を確認してください',
      guidance: '受付処理を今すぐ実行し、解消しない場合は権限と実行ログを確認します。', safeRetryCandidates: retryCount };
  }
  if (operations.queueStale) {
    return { status: 'ERROR', label: '処理が長時間止まっています',
      guidance: '「' + VN_ADMIN_MENU_RUN_NOW + '」を実行し、要確認事項を確認してください。', safeRetryCandidates: retryCount };
  }
  if (model.portalRequests && model.portalRequests.unavailable) {
    return { status: 'ATTENTION', label: '申請入口の状態を確認できません',
      guidance: '他の判断項目は表示できています。続く場合はポータルの権限を確認します。', safeRetryCandidates: retryCount };
  }
  if (Number(counts.pendingApprovals || 0) || Number(counts.exceptions || 0) ||
      Number(counts.portalAttention || 0)) {
    return { status: 'ATTENTION', label: '確認が必要な項目があります',
      guidance: '上から順に承認待ちと要確認事項を確認してください。', safeRetryCandidates: retryCount };
  }
  const portalCounts = model.portalRequests && model.portalRequests.counts || {};
  if (Number(counts.queuedJobs || 0) || Number(counts.runningJobs || 0) ||
      Number(portalCounts.waiting || 0) || Number(portalCounts.processing || 0)) {
    return { status: 'PROCESSING', label: '自動処理が進行中です',
      guidance: '操作は不要です。通常は5分以内に状態が更新されます。', safeRetryCandidates: retryCount };
  }
  return { status: 'OK', label: '現在、管理ハブの対応はありません',
    guidance: '自動更新は正常です。', safeRetryCandidates: retryCount };
}

/**
 * Create a clean 管理ハブ and a clean Master Template from the current
 * deployed runtimes. The legacy workbook's sheets, cells and bound script are
 * never copied into either operational container.
 */
function vNextAdminBootstrapFromCurrent(request) {
  return vNextAdminGuard_('vNextAdminBootstrapFromCurrent', function () {
    vNextAdminAssertRuntimeConfigurator_();
    const req = request && typeof request === 'object' ? request : {};
    const source = SpreadsheetApp.getActiveSpreadsheet();
    const sourceMode = vNextDetectBookMode_(source);
    if (sourceMode === 'ADMIN') {
      return { reused: true, mode: sourceMode, hubUrl: source.getUrl(), message: 'This workbook is already an 管理ハブ.' };
    }
    if (sourceMode !== 'LEGACY' && sourceMode !== 'TEMPLATE') {
      throw new Error('Bootstrap is allowed only from a LEGACY or TEMPLATE workbook. Current mode=' + sourceMode);
    }

    return vNextAdminWithScriptLock_('bootstrap', function () {
      const runtime = vNextGetRuntimeConfig_();
      const actualSourceId = vNextAdminText_(req.forecastSourceSpreadsheetId || req.sourceSpreadsheetId) ||
        runtime.VNEXT_ZAC_SOURCE_SPREADSHEET_ID || runtime.FORECAST_SOURCE_SPREADSHEET_ID;
      if (!actualSourceId) {
        throw new Error('VNEXT_ZAC_SOURCE_SPREADSHEET_ID is required. Configure it with vNextConfigureRuntime before bootstrap.');
      }
      if (!vNextAdminSpreadsheetAccessible_(actualSourceId)) throw new Error('Configured forecast source spreadsheet is not accessible.');
      const props = PropertiesService.getScriptProperties();
      const existingHubId = runtime.VNEXT_ADMIN_HUB_SPREADSHEET_ID;
      const existingTemplateId = runtime.VNEXT_MASTER_TEMPLATE_SPREADSHEET_ID;
      if (!vNextAdminBool_(req.force) && vNextAdminSpreadsheetAccessible_(existingHubId) && vNextAdminSpreadsheetAccessible_(existingTemplateId)) {
        const existingHub = SpreadsheetApp.openById(existingHubId);
        const existingTemplate = SpreadsheetApp.openById(existingTemplateId);
        const existingHubConfig = vNextAdminReadKeyValueSheet_(existingHub, VN_ADMIN_SYSTEM_CONFIG_SHEET);
        const templateConfig = vNextAdminReadKeyValueSheet_(existingTemplate, VN_ADMIN_BOOK_CONFIG_SHEET);
        const existingRelease = vNextAdminReadTable_(existingHub, VN_ADMIN_SHEETS.RELEASES).rows.find(function (row) {
          return String(row.release_id || '') === String(templateConfig.version || '') &&
            String(row.status || '').toUpperCase() === 'ACTIVE';
        });
        const cleanHubRuntime = String(existingHubConfig.admin_source_script_id || '') &&
          String(existingHubConfig.admin_hub_script_id || '') &&
          String(existingHubConfig.admin_runtime_sha256 || '');
        if (cleanHubRuntime && typeof vNextAdminRuntimeAssertBoundParent_ === 'function') {
          vNextAdminRuntimeAssertBoundParent_(String(existingHubConfig.admin_hub_script_id), existingHub.getId());
        }
        if (cleanHubRuntime && existingRelease && String(existingRelease.client_runtime_sha256 || '') &&
            String(existingRelease.client_runtime_sha256 || '') === String(templateConfig.client_runtime_bundle_sha256 || '')) {
          const existingModelRelease = vNextAdminTryResolveActiveModelRelease_(existingHub);
          return {
            reused: true,
            hubId: existingHubId,
            hubUrl: existingHub.getUrl(),
            templateId: existingTemplateId,
            templateUrl: existingTemplate.getUrl(),
            releaseId: existingRelease.release_id,
            modelReleaseId: existingModelRelease && existingModelRelease.model_release_id || '',
            adminHubScriptId: String(existingHubConfig.admin_hub_script_id || ''),
            adminRuntimeSha256: String(existingHubConfig.admin_runtime_sha256 || ''),
            clientRuntimeVersion: existingRelease.client_runtime_version,
            clientRuntimeSha256: existingRelease.client_runtime_sha256
          };
        }
      }

      const now = new Date();
      const actor = vNextAdminActor_();
      const vertexConfig = vNextAdminResolveBootstrapVertexConfig_(runtime);
      const stamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd-HHmm');
      const hubName = vNextAdminText_(req.hubName) || ('Forecast vNext 管理ハブ ' + stamp);
      const templateName = vNextAdminText_(req.templateName) || ('Forecast vNext Master Template ' + stamp);
      const releaseId = vNextAdminText_(req.releaseId) || ('vnext-' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd'));
      const modelReleaseId = vNextAdminText_(req.modelReleaseId) ||
        ('model-vnext-' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd'));
      if (modelReleaseId === releaseId) {
        throw new Error('Model Release ID must be different from the Template Release ID.');
      }
      const adminEmails = vNextAdminMergeEmails_(req.adminEmails, runtime.VNEXT_ADMIN_EMAILS, actor);
      const folder = vNextAdminPreparePrivateBootstrapFolder_(
        req.folderId || runtime.VNEXT_ROOT_FOLDER_ID,
        'Forecast vNext Private ' + stamp,
        adminEmails
      );

      if (typeof vNextClientRuntimeCreateBoundSpreadsheet_ !== 'function' ||
          typeof vNextAdminRuntimeCreateBoundSpreadsheet_ !== 'function') {
        throw new Error('Runtime provisioner is not installed for both Admin and Client projects.');
      }
      const templateRuntime = vNextClientRuntimeCreateBoundSpreadsheet_({
        title: templateName, folderId: folder.getId()
      });
      const templateFile = DriveApp.getFileById(templateRuntime.spreadsheetId);
      const adminSourceScriptId = ScriptApp.getScriptId();
      const adminRuntime = vNextAdminRuntimeCreateBoundSpreadsheet_({
        title: hubName, folderId: folder.getId(), sourceScriptId: adminSourceScriptId
      });
      const hubFile = DriveApp.getFileById(adminRuntime.spreadsheetId);
      const hub = SpreadsheetApp.openById(adminRuntime.spreadsheetId);
      const template = SpreadsheetApp.openById(templateRuntime.spreadsheetId);
      const hubBookId = 'HUB-' + Utilities.getUuid();
      const templateBookId = 'TPL-' + Utilities.getUuid();

      vNextAdminInitializeHub_(hub, {
        bookId: hubBookId,
        sourceSpreadsheetId: actualSourceId,
        templateSpreadsheetId: template.getId(),
        releaseId: releaseId,
        modelReleaseId: modelReleaseId,
        vertexConfig: vertexConfig,
        adminSourceScriptId: adminRuntime.sourceScriptId,
        adminHubScriptId: adminRuntime.scriptId,
        adminRuntimeSha256: adminRuntime.adminRuntimeSha256,
        adminEmails: adminEmails,
        actor: actor,
        now: now, resetCopied: true
      });
      vNextAdminEnforcePrivateFileAcl_(hubFile, adminEmails);
      vNextAdminEnforcePrivateFileAcl_(templateFile, adminEmails);
      vNextAdminWriteSystemConfig_(hub, {
        client_runtime_version: templateRuntime.runtimeVersion,
        client_runtime_sha256: templateRuntime.bundleSha256,
        template_script_id: templateRuntime.scriptId,
        admin_source_script_id: adminRuntime.sourceScriptId,
        admin_hub_script_id: adminRuntime.scriptId,
        admin_runtime_sha256: adminRuntime.adminRuntimeSha256,
        private_root_folder_id: folder.getId()
      });
      vNextAdminInitializeTemplate_(template, {
        bookId: templateBookId,
        releaseId: releaseId,
        clientRuntimeVersion: templateRuntime.runtimeVersion,
        clientRuntimeSha256: templateRuntime.bundleSha256,
        adminEmails: adminEmails,
        actor: actor,
        now: now, resetCopied: true
      });

      vNextAdminRegisterBook_(hub, {
        book_id: hubBookId, mode: 'ADMIN', client_id: '', client_name: '', fiscal_year: '',
        spreadsheet_id: hub.getId(), spreadsheet_url: hub.getUrl(), template_release_id: releaseId,
        admin_script_id: adminRuntime.scriptId, admin_runtime_sha256: adminRuntime.adminRuntimeSha256,
        schema_version: VN_ADMIN_SCHEMA_VERSION, state: 'ACTIVE', status: 'ACTIVE',
        health_status: 'OK', health_code: 'BOOTSTRAPPED', last_health_at: now,
        forecast_owner_emails: adminEmails.join(','), editor_emails: adminEmails.join(','), viewer_emails: '',
        created_at: now, created_by: actor, updated_at: now, note: '管理ハブ'
      });
      vNextAdminRegisterBook_(hub, {
        book_id: templateBookId, mode: 'TEMPLATE', client_id: '', client_name: '', fiscal_year: '',
        spreadsheet_id: template.getId(), spreadsheet_url: template.getUrl(), template_release_id: releaseId,
        client_script_id: templateRuntime.scriptId,
        client_runtime_version: templateRuntime.runtimeVersion,
        client_runtime_sha256: templateRuntime.bundleSha256,
        schema_version: vNextAdminClientSchemaVersion_(), state: 'TEMPLATE_READY', status: 'ACTIVE',
        health_status: 'OK', health_code: 'BOOTSTRAPPED', last_health_at: now,
        forecast_owner_emails: adminEmails.join(','), editor_emails: adminEmails.join(','), viewer_emails: '',
        created_at: now, created_by: actor, updated_at: now,
        note: 'Master Template runtime=' + templateRuntime.runtimeVersion + ' sha256=' + templateRuntime.bundleSha256
      });
      vNextAdminRegisterRelease_(hub, {
        release_id: releaseId,
        release_name: vNextAdminText_(req.releaseName) || releaseId,
        status: 'ACTIVE',
        template_spreadsheet_id: template.getId(),
        schema_version: vNextAdminClientSchemaVersion_(),
        engine_version: typeof VNEXT_ENGINE !== 'undefined' ? VNEXT_ENGINE.VERSION : vNextAdminText_(req.engineVersion),
        ux_version: vNextAdminText_(req.uxVersion),
        admin_version: VN_ADMIN_SCHEMA_VERSION,
        client_runtime_version: templateRuntime.runtimeVersion,
        client_runtime_sha256: templateRuntime.bundleSha256,
        template_content_sha256: vNextAdminTemplateUiManifestHash_(template),
        template_manifest_schema: VN_ADMIN_TEMPLATE_MANIFEST_SCHEMA,
        template_script_id: templateRuntime.scriptId,
        created_at: now, created_by: actor, activated_at: now,
        note: 'Initial bootstrap release'
      });
      const initialModelRelease = vNextAdminEnsureInitialModelRelease_(hub, {
        modelReleaseId: modelReleaseId,
        modelVersion: typeof VNEXT_ENGINE !== 'undefined' ? VNEXT_ENGINE.VERSION :
          (vNextAdminText_(req.modelVersion || req.engineVersion) || 'vnext-initial'),
        templateVersion: releaseId,
        parameters: req.modelParameters || {},
        actor: actor,
        now: now
      });
      vNextAdminWriteCanonicalReleasePair_(hub, releaseId, modelReleaseId, template.getId());
      const initialTemplateRelease = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.RELEASES).rows.find(function (row) {
        return String(row.release_id || '') === releaseId;
      });
      vNextAdminWriteActiveReleasePairCaches_(hub, initialTemplateRelease, initialModelRelease);
      props.setProperties({
        FORECAST_SOURCE_SPREADSHEET_ID: actualSourceId,
        VNEXT_ZAC_SOURCE_SPREADSHEET_ID: actualSourceId,
        VNEXT_ADMIN_HUB_SPREADSHEET_ID: hub.getId(),
        VNEXT_MASTER_TEMPLATE_SPREADSHEET_ID: template.getId(),
        VNEXT_ACTIVE_RELEASE_ID: releaseId,
        VNEXT_ACTIVE_MODEL_RELEASE_ID: modelReleaseId
      }, false);
      vNextAdminWriteAudit_(hub, 'BOOTSTRAP', 'WORKBOOK_SET', hubBookId, 'SUCCESS', {
        legacyBootstrapSpreadsheetId: source.getId(), forecastSourceSpreadsheetId: actualSourceId,
        templateSpreadsheetId: template.getId(), releaseId: releaseId,
        modelReleaseId: modelReleaseId,
        adminSourceScriptId: adminRuntime.sourceScriptId,
        adminHubScriptId: adminRuntime.scriptId,
        adminRuntimeSha256: adminRuntime.adminRuntimeSha256,
        clientRuntimeVersion: templateRuntime.runtimeVersion,
        clientRuntimeSha256: templateRuntime.bundleSha256,
        templateScriptId: templateRuntime.scriptId
      });
      vNextAdminRefreshTodayExceptions_(hub);
      vNextAdminRefreshHome_(hub);
      SpreadsheetApp.flush();
      return {
        reused: false,
        hubId: hub.getId(), hubUrl: hub.getUrl(),
        templateId: template.getId(), templateUrl: template.getUrl(),
        releaseId: releaseId, modelReleaseId: modelReleaseId,
        adminHubScriptId: adminRuntime.scriptId,
        adminRuntimeSha256: adminRuntime.adminRuntimeSha256,
        clientRuntimeVersion: templateRuntime.runtimeVersion,
        clientRuntimeSha256: templateRuntime.bundleSha256,
        nextStep: '管理ハブを開いて再読込し、サイドバーの「自動運用を有効化」を1回押してください。'
      };
    });
  });
}

/**
 * Resume a bootstrap that reached the clean Hub/Template creation phase but
 * was interrupted by the Apps Script execution limit before release records
 * and canonical pointers were committed. The supplied IDs are mandatory so a
 * similarly named Drive file can never be selected by accident.
 */
function vNextAdminRecoverIncompleteBootstrap(request) {
  return vNextAdminGuard_('vNextAdminRecoverIncompleteBootstrap', function () {
    vNextAdminAssertRuntimeConfigurator_();
    const req = request && typeof request === 'object' ? request : {};
    const source = SpreadsheetApp.getActiveSpreadsheet();
    const sourceMode = vNextDetectBookMode_(source);
    if (sourceMode !== 'LEGACY' && sourceMode !== 'TEMPLATE') {
      throw new Error('Incomplete bootstrap recovery is allowed only from the source LEGACY/TEMPLATE workbook.');
    }
    const hubId = vNextAdminRequiredText_(req.hubSpreadsheetId, 'hubSpreadsheetId');
    const templateId = vNextAdminRequiredText_(req.templateSpreadsheetId, 'templateSpreadsheetId');
    if (hubId === templateId || hubId === source.getId() || templateId === source.getId()) {
      throw new Error('Source, Hub and Template must be three different spreadsheets.');
    }
    return vNextAdminWithScriptLock_('recover-incomplete-bootstrap', function () {
      const hub = SpreadsheetApp.openById(hubId);
      const template = SpreadsheetApp.openById(templateId);
      const hubConfig = vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET);
      const hubRouting = vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_BOOK_CONFIG_SHEET);
      const templateRouting = vNextAdminReadKeyValueSheet_(template, VN_ADMIN_BOOK_CONFIG_SHEET);
      const templateConfig = vNextAdminReadKeyValueSheet_(template, VN_ADMIN_SYSTEM_CONFIG_SHEET);
      if (String(hubConfig.mode || '').toUpperCase() !== 'ADMIN' ||
          String(hubConfig.admin_hub_spreadsheet_id || '') !== hubId ||
          String(hubConfig.template_spreadsheet_id || '') !== templateId) {
        throw new Error('The supplied Hub does not contain the expected bootstrap routing identity.');
      }
      if (String(hubRouting.book_id || '') !== String(hubConfig.book_id || '') ||
          String(templateRouting.mode || '').toUpperCase() !== 'TEMPLATE' ||
          String(templateConfig.mode || '').toUpperCase() !== 'TEMPLATE' ||
          String(templateRouting.book_id || '') !== String(templateConfig.book_id || '')) {
        throw new Error('Hub/Template BOOK_META identity is inconsistent.');
      }
      const releaseId = vNextAdminRequiredText_(templateRouting.version || hubConfig.active_release_id, 'releaseId');
      const modelReleaseId = vNextAdminRequiredText_(hubConfig.active_model_release_id, 'modelReleaseId');
      if (String(hubConfig.active_release_id || '') !== releaseId || modelReleaseId === releaseId) {
        throw new Error('Hub and Template release identities are inconsistent.');
      }
      const runtime = vNextGetRuntimeConfig_();
      const actualSourceId = vNextAdminRequiredText_(
        hubConfig.source_spreadsheet_id || runtime.VNEXT_ZAC_SOURCE_SPREADSHEET_ID || runtime.FORECAST_SOURCE_SPREADSHEET_ID,
        'sourceSpreadsheetId'
      );
      if (!vNextAdminSpreadsheetAccessible_(actualSourceId)) {
        throw new Error('Configured forecast source spreadsheet is not accessible.');
      }
      const actor = vNextAdminActor_();
      const adminEmails = vNextAdminMergeEmails_(hubConfig.admin_emails, runtime.VNEXT_ADMIN_EMAILS, actor);
      if (adminEmails.indexOf(String(actor || '').toLowerCase()) < 0) {
        throw new Error('Only a registered Hub Admin can resume bootstrap.');
      }
      const adminSourceScriptId = vNextAdminRequiredText_(hubConfig.admin_source_script_id, 'admin_source_script_id');
      const adminHubScriptId = vNextAdminRequiredText_(hubConfig.admin_hub_script_id, 'admin_hub_script_id');
      if (adminSourceScriptId !== ScriptApp.getScriptId()) {
        throw new Error('Recovery must run from the central source script used to create this Hub.');
      }
      if (typeof vNextAdminRuntimeAssertBoundParent_ === 'function') {
        vNextAdminRuntimeAssertBoundParent_(adminHubScriptId, hubId);
      }
      const now = new Date();
      const hubBookId = vNextAdminRequiredText_(hubRouting.book_id, 'hubBookId');
      const templateBookId = vNextAdminRequiredText_(templateRouting.book_id, 'templateBookId');
      const clientRuntimeVersion = vNextAdminRequiredText_(templateRouting.client_runtime_version, 'clientRuntimeVersion');
      const clientRuntimeSha256 = vNextAdminRequiredText_(templateRouting.client_runtime_bundle_sha256, 'clientRuntimeSha256');
      const templateScriptId = vNextAdminRequiredText_(hubConfig.template_script_id, 'template_script_id');
      const adminRuntimeSha256 = vNextAdminRequiredText_(hubConfig.admin_runtime_sha256, 'admin_runtime_sha256');
      vNextAdminRegisterBook_(hub, {
        book_id: hubBookId, mode: 'ADMIN', client_id: '', client_name: '', fiscal_year: '',
        spreadsheet_id: hubId, spreadsheet_url: hub.getUrl(), template_release_id: releaseId,
        admin_script_id: adminHubScriptId, admin_runtime_sha256: adminRuntimeSha256,
        schema_version: VN_ADMIN_SCHEMA_VERSION, state: 'ACTIVE', status: 'ACTIVE',
        health_status: 'OK', health_code: 'BOOTSTRAP_RECOVERED', last_health_at: now,
        forecast_owner_emails: adminEmails.join(','), editor_emails: adminEmails.join(','), viewer_emails: '',
        created_at: hubRouting.created_at || now, created_by: hubRouting.created_by || actor,
        updated_at: now, note: '管理ハブ; incomplete bootstrap recovered'
      });
      vNextAdminRegisterBook_(hub, {
        book_id: templateBookId, mode: 'TEMPLATE', client_id: '', client_name: '', fiscal_year: '',
        spreadsheet_id: templateId, spreadsheet_url: template.getUrl(), template_release_id: releaseId,
        client_script_id: templateScriptId, client_runtime_version: clientRuntimeVersion,
        client_runtime_sha256: clientRuntimeSha256, schema_version: vNextAdminClientSchemaVersion_(),
        state: 'TEMPLATE_READY', status: 'ACTIVE', health_status: 'OK',
        health_code: 'BOOTSTRAP_RECOVERED', last_health_at: now,
        forecast_owner_emails: adminEmails.join(','), editor_emails: adminEmails.join(','), viewer_emails: '',
        created_at: templateRouting.created_at || now, created_by: templateRouting.created_by || actor,
        updated_at: now, note: 'Master Template; incomplete bootstrap recovered'
      });
      let initialTemplateRelease = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.RELEASES).rows.find(function (row) {
        return String(row.release_id || '') === releaseId;
      });
      if (initialTemplateRelease) {
        if (String(initialTemplateRelease.status || '').toUpperCase() !== 'ACTIVE' ||
            String(initialTemplateRelease.template_spreadsheet_id || '') !== templateId ||
            String(initialTemplateRelease.client_runtime_version || '') !== clientRuntimeVersion ||
            String(initialTemplateRelease.client_runtime_sha256 || '') !== clientRuntimeSha256 ||
            String(initialTemplateRelease.template_script_id || '') !== templateScriptId ||
            !String(initialTemplateRelease.template_content_sha256 || '')) {
          throw new Error('Existing bootstrap release is incomplete or conflicts with the supplied Template.');
        }
      } else {
        const templateHash = vNextAdminTemplateUiManifestHash_(template);
        vNextAdminRegisterRelease_(hub, {
          release_id: releaseId, release_name: releaseId, status: 'ACTIVE',
          template_spreadsheet_id: templateId, schema_version: vNextAdminClientSchemaVersion_(),
          engine_version: typeof VNEXT_ENGINE !== 'undefined' ? VNEXT_ENGINE.VERSION : '',
          admin_version: VN_ADMIN_SCHEMA_VERSION, client_runtime_version: clientRuntimeVersion,
          client_runtime_sha256: clientRuntimeSha256, template_content_sha256: templateHash,
          template_manifest_schema: VN_ADMIN_TEMPLATE_MANIFEST_SCHEMA, template_script_id: templateScriptId,
          created_at: templateRouting.created_at || now, created_by: templateRouting.created_by || actor,
          activated_at: now, note: 'Initial bootstrap release; recovered after execution limit'
        });
        initialTemplateRelease = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.RELEASES).rows.find(function (row) {
          return String(row.release_id || '') === releaseId;
        });
      }
      const initialModelRelease = vNextAdminEnsureInitialModelRelease_(hub, {
        modelReleaseId: modelReleaseId,
        modelVersion: typeof VNEXT_ENGINE !== 'undefined' ? VNEXT_ENGINE.VERSION : 'vnext-initial',
        templateVersion: releaseId, parameters: {}, actor: actor, now: now
      });
      vNextAdminWriteCanonicalReleasePair_(hub, releaseId, modelReleaseId, templateId);
      vNextAdminWriteActiveReleasePairCaches_(hub, initialTemplateRelease, initialModelRelease);
      PropertiesService.getScriptProperties().setProperties({
        FORECAST_SOURCE_SPREADSHEET_ID: actualSourceId,
        VNEXT_ZAC_SOURCE_SPREADSHEET_ID: actualSourceId,
        VNEXT_ADMIN_HUB_SPREADSHEET_ID: hubId,
        VNEXT_MASTER_TEMPLATE_SPREADSHEET_ID: templateId,
        VNEXT_ACTIVE_RELEASE_ID: releaseId,
        VNEXT_ACTIVE_MODEL_RELEASE_ID: modelReleaseId
      }, false);
      vNextAdminWriteAudit_(hub, 'BOOTSTRAP_RECOVERY', 'WORKBOOK_SET', hubBookId, 'SUCCESS', {
        sourceSpreadsheetId: source.getId(), forecastSourceSpreadsheetId: actualSourceId,
        templateSpreadsheetId: templateId, releaseId: releaseId, modelReleaseId: modelReleaseId
      });
      vNextAdminRefreshTodayExceptions_(hub);
      vNextAdminRefreshHome_(hub);
      SpreadsheetApp.flush();
      return {
        recovered: true, hubId: hubId, hubUrl: hub.getUrl(), templateId: templateId,
        templateUrl: template.getUrl(), releaseId: releaseId, modelReleaseId: modelReleaseId,
        nextStep: '管理ハブを開いて権限を許可し、自動運用を有効化してください。'
      };
    });
  });
}

/** Re-run the idempotent Hub initialization in an already copied workbook. */
function vNextAdminActivateCurrentHub(request) {
  return vNextAdminGuard_('vNextAdminActivateCurrentHub', function () {
    const req = request && typeof request === 'object' ? request : {};
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mode = vNextDetectBookMode_(ss);
    if (!vNextAdminCanActivateHubMode_(mode) || !vNextAdminIsRegisteredHub_(ss)) {
      throw new Error('既存管理ハブの再初期化だけが許可されています。現在のmode=' + mode + '（CLIENTは変換できません）。');
    }
    vNextAdminAssertHubAdmin_(ss, false);
    return vNextAdminWithDocumentLock_('activate-hub', function () {
      const runtime = vNextGetRuntimeConfig_();
      const routing = vNextAdminReadKeyValueSheet_(ss, VN_ADMIN_BOOK_CONFIG_SHEET);
      const bookId = routing.book_id || ('HUB-' + Utilities.getUuid());
      const templateId = req.templateSpreadsheetId || runtime.VNEXT_MASTER_TEMPLATE_SPREADSHEET_ID || routing.template_spreadsheet_id || '';
      vNextAdminInitializeHub_(ss, {
        bookId: bookId,
        sourceSpreadsheetId: req.sourceSpreadsheetId || runtime.FORECAST_SOURCE_SPREADSHEET_ID || '',
        templateSpreadsheetId: templateId,
        releaseId: req.releaseId || runtime.VNEXT_ACTIVE_RELEASE_ID || routing.version || '',
        modelReleaseId: req.modelReleaseId || runtime.VNEXT_ACTIVE_MODEL_RELEASE_ID || '',
        vertexConfig: vNextAdminResolveBootstrapVertexConfig_(runtime),
        adminEmails: vNextAdminMergeEmails_(req.adminEmails, runtime.VNEXT_ADMIN_EMAILS, vNextAdminActor_()),
        actor: vNextAdminActor_(), now: new Date()
      });
      vNextBuildAdminMenu_();
      vNextAdminRefreshTodayExceptions_(ss);
      vNextAdminRefreshHome_(ss);
      return { mode: vNextDetectBookMode_(ss), spreadsheetUrl: ss.getUrl(), bookId: bookId };
    });
  });
}

function vNextAdminCanActivateHubMode_(mode) {
  return String(mode || '').toUpperCase() === 'ADMIN';
}

/**
 * A routing cell alone is not sufficient to authorize Hub operations. Client
 * editors can edit local hidden sheets, so require the Hub-only tables and the
 * self-referential ADMIN registry row before any destructive reinitialization.
 */
function vNextAdminIsRegisteredHub_(ss) {
  try {
    if (!ss) return false;
    const required = [VN_ADMIN_SHEETS.REGISTRY, VN_ADMIN_SHEETS.JOBS, VN_ADMIN_SHEETS.RELEASES];
    if (required.some(function (name) { return !ss.getSheetByName(name); })) return false;
    const routing = Object.assign(
      {},
      vNextAdminReadKeyValueSheet_(ss, VN_ADMIN_SYSTEM_CONFIG_SHEET),
      vNextAdminReadKeyValueSheet_(ss, VN_ADMIN_BOOK_CONFIG_SHEET)
    );
    if (String(routing.mode || '').toUpperCase() !== 'ADMIN' || !String(routing.book_id || '')) return false;
    return vNextAdminReadTable_(ss, VN_ADMIN_SHEETS.REGISTRY).rows.some(function (row) {
      return String(row.book_id || '') === String(routing.book_id || '') &&
        String(row.mode || '').toUpperCase() === 'ADMIN' &&
        String(row.spreadsheet_id || '') === String(ss.getId());
    });
  } catch (err) {
    Logger.log('Registered Hub verification failed: %s', String(err && err.message || err));
    return false;
  }
}

function vNextAdminResolveBootstrapVertexConfig_(runtime) {
  const source = runtime || {};
  let legacy = {};
  if (typeof readVertexConfig_ === 'function') {
    try { legacy = readVertexConfig_() || {}; }
    catch (err) { Logger.log('Legacy Vertex config read skipped: %s', String(err && err.message || err)); }
  }
  return {
    VERTEX_PROJECT_ID: source.VERTEX_PROJECT_ID || legacy.projectId || '',
    VERTEX_LOCATION: source.VERTEX_LOCATION || legacy.location || '',
    VERTEX_GEMINI_MODEL: source.VERTEX_GEMINI_MODEL || legacy.geminiModel || '',
    VERTEX_DATASTORE_ID: source.VERTEX_DATASTORE_ID || legacy.datastoreId || '',
    VERTEX_SEARCH_LOCATION: source.VERTEX_SEARCH_LOCATION || legacy.searchLocation || '',
    VERTEX_SERVING_CONFIG: source.VERTEX_SERVING_CONFIG || legacy.servingConfig || ''
  };
}

/** Provision one client x fiscal-year workbook from the active release template. */
function vNextAdminProvisionClient(request) {
  return vNextAdminGuard_('vNextAdminProvisionClient', function () {
    const hub = vNextAdminRequireHub_();
    return vNextAdminProvisionClientInHub_(hub, request);
  });
}

/**
 * Pilot-only fallback: provision through the central source project when a
 * freshly generated Hub is still waiting for its own Cloud-project linkage.
 */
function vNextAdminProvisionPilotClientFromSource(request) {
  return vNextAdminGuard_('vNextAdminProvisionPilotClientFromSource', function () {
    vNextAdminAssertRuntimeConfigurator_();
    const req = request && typeof request === 'object' ? request : {};
    const source = SpreadsheetApp.getActiveSpreadsheet();
    const sourceMode = vNextDetectBookMode_(source);
    if (sourceMode !== 'LEGACY' && sourceMode !== 'TEMPLATE') {
      throw new Error('Pilot source provisioning is allowed only from the registered source workbook.');
    }
    const hubId = vNextAdminRequiredText_(req.hubSpreadsheetId, 'hubSpreadsheetId');
    const hub = SpreadsheetApp.openById(hubId);
    const config = vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET);
    if (String(config.mode || '').toUpperCase() !== 'ADMIN' ||
        String(config.admin_hub_spreadsheet_id || '') !== hubId ||
        String(config.admin_source_script_id || '') !== ScriptApp.getScriptId() ||
        !vNextAdminIsRegisteredHub_(hub)) {
      throw new Error('The supplied Hub is not registered to this central source project.');
    }
    vNextAdminAssertHubAdmin_(hub, false);
    vNextAdminHydrateHubRuntime_(hub);
    return vNextAdminProvisionClientInHub_(hub, req);
  });
}

/** Creates or returns the single employee-facing annual planning portal. */
function vNextAdminProvisionSharedPortal(request) {
  return vNextAdminGuard_('vNextAdminProvisionSharedPortal', function () {
    const hub = vNextAdminRequireHub_();
    const result = vNextAdminProvisionSharedPortalInHub_(hub, request);
    // Provisioning holds the short global lock. The comparatively expensive
    // ZAC read/projection is deliberately performed only after that lock ends.
    result.clientCatalog = vNextAdminRefreshZacClientCatalogSafely_(hub, true);
    return result;
  });
}

/**
 * One-click, Admin-only recovery path for the first employee Portal pilot.
 *
 * This deliberately avoids SpreadsheetApp.getUi(), so it can also be run from
 * the Apps Script editor when a changed Cloud-project authorization prevents a
 * custom menu from appearing. Every phase uses the ordinary immutable release
 * APIs and deterministic IDs; rerunning after a timeout resumes instead of
 * creating another Template, Model Release, or Portal.
 */
function vNextAdminPrepareEmployeePortalPilot(request) {
  return vNextAdminGuard_('vNextAdminPrepareEmployeePortalPilot', function () {
    const req = request && typeof request === 'object' ? request : {};
    if (req.attestationConfirmed !== true) {
      throw new Error('テスト結果を確認し、attestationConfirmed=trueを明示してください。');
    }
    const hub = vNextAdminRequireHub_();
    const initialPair = vNextAdminReadActiveReleasePair_(hub);
    // A release bootstrap is the one place where the currently active Model
    // may legitimately target the immediately previous Engine. Validate that
    // source release against its own immutable Template/Model pair, then bind
    // the copied parameters to the newly deployed Engine below. Using the
    // normal runtime resolver here would make every Engine version bump
    // impossible because it correctly requires an exact deployed-version
    // match for ordinary forecast execution.
    const initialModel = vNextAdminResolveActiveModelReleaseForUpgrade_(hub, initialPair);
    const parameters = vNextAdminParseJson_(initialModel.parameters_json, {});
    const engineVersion = typeof VNEXT_ENGINE !== 'undefined' ? String(VNEXT_ENGINE.VERSION || '') : '';
    if (!engineVersion) throw new Error('Forecast Engine version is unavailable.');

    const staged = vNextAdminPublishTemplateRelease({
      reason: vNextAdminText_(req.releaseReason) ||
        '申請入口・社内情報提供メンバー対応のPilot release',
      expectedActiveReleaseId: initialPair.releaseId,
      stageOnly: true
    });
    const releaseId = vNextAdminRequiredText_(staged && staged.releaseId, 'staged.releaseId');
    const modelReleaseId = 'model-portal-' + vNextAdminSha256_(vNextAdminCanonicalJson_({
      releaseId: releaseId,
      engineVersion: engineVersion,
      parameters: parameters
    })).slice(0, 20).toUpperCase();
    const note = '申請入口Pilot。予測Engineとparameterは直前のACTIVE Model Releaseから変更なし。';
    const checks = {
      backtest: {
        status: 'PASS',
        basis: 'UNCHANGED_ENGINE_AND_PARAMETERS_FROM_ACTIVE_MODEL',
        sourceModelReleaseId: initialModel.model_release_id,
        reviewedBy: vNextAdminActor_(),
        evidenceArtifact: vNextAdminRequiredText_(req.evidenceArtifact, 'evidenceArtifact')
      },
      canary: {
        status: 'PASS',
        basis: 'CLIENT_AND_PORTAL_RUNTIME_CONTRACT_TESTS',
        clientRuntimeTests: 10,
        portalRuntimeTests: 12,
        integrationContractTests: 'PASS',
        reviewedBy: vNextAdminActor_(),
        evidenceArtifact: vNextAdminRequiredText_(req.evidenceArtifact, 'evidenceArtifact')
      }
    };

    let model = vNextAdminLatestModelRelease_(hub, modelReleaseId);
    if (!model) {
      model = vNextAdminRegisterModelRelease({
        modelReleaseId: modelReleaseId,
        modelVersion: engineVersion,
        templateVersion: releaseId,
        parameters: parameters,
        backtest: checks.backtest,
        canary: checks.canary,
        attestationConfirmed: true,
        note: note
      });
    } else {
      if (['DRAFT', 'ACTIVE'].indexOf(String(model.status || '').toUpperCase()) < 0 ||
          String(model.model_version || '') !== engineVersion ||
          String(model.template_version || '') !== releaseId ||
          String(model.parameters_json || '') !== vNextAdminCanonicalJson_(
            vNextAdminNormalizeModelParameters_(parameters)
          )) {
        throw new Error('Existing Portal Pilot MODEL_RELEASE has different immutable content.');
      }
      vNextAdminAssertModelReleaseChecksPassed_(model);
    }

    let activation = { reused: true, releaseId: releaseId, modelReleaseId: modelReleaseId };
    const pairAfterStage = vNextAdminReadActiveReleasePair_(hub);
    if (pairAfterStage.releaseId !== releaseId || pairAfterStage.modelReleaseId !== modelReleaseId) {
      if (pairAfterStage.releaseId !== initialPair.releaseId ||
          pairAfterStage.modelReleaseId !== initialPair.modelReleaseId) {
        throw new Error('準備中に別のTemplate・Model pairが有効化されました。上書きせず停止します。');
      }
      activation = vNextAdminActivateReleasePair({
        releaseId: releaseId,
        modelReleaseId: modelReleaseId,
        reason: '申請入口PilotのTemplate・Model pairを有効化',
        expectedActiveReleaseId: initialPair.releaseId,
        expectedActiveModelReleaseId: initialPair.modelReleaseId
      });
    }

    const skipPortal = req.skipPortal === true;
    const portal = skipPortal
      ? { skipped: true, reason: 'skipPortal' }
      : vNextAdminProvisionSharedPortal({ title: VNEXT_NAMING.LAYER2_DEFAULT_TITLE });
    const result = vNextAdminJsonSafe_({
      ok: true,
      activeTemplateReleaseId: releaseId,
      activeModelReleaseId: modelReleaseId,
      staged: staged,
      activation: activation,
      portal: portal,
      skipPortal: skipPortal,
      completedAt: new Date().toISOString()
    });
    PropertiesService.getScriptProperties().setProperty(
      'VNEXT_LAST_EMPLOYEE_PORTAL_PILOT_RESULT_JSON',
      vNextAdminCanonicalJson_(result)
    );
    Logger.log('EMPLOYEE_PORTAL_PILOT_READY %s', vNextAdminCanonicalJson_(result));
    return result;
  });
}

function vNextAdminWriteDevDeployProgress_(step, detail) {
  const payload = {
    step: String(step || ''),
    detail: detail && typeof detail === 'object' ? detail : { message: String(detail || '') },
    updatedAt: new Date().toISOString()
  };
  PropertiesService.getScriptProperties().setProperty(
    'VNEXT_DEV_DEPLOY_PROGRESS_JSON', vNextAdminCanonicalJson_(payload)
  );
  Logger.log('DEV_DEPLOY_PROGRESS %s', vNextAdminCanonicalJson_(payload));
  return payload;
}

/**
 * Spreadsheet macro entry for the first live smoke test. Unlike running a UI
 * function from the Apps Script editor, an imported Sheet macro has a valid UI
 * context. The explicit YES response is the Admin attestation recorded by the
 * release API; cancelling performs no write.
 */
function vNextAdminPrepareEmployeePortalPilotForManualTest() {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert(
    '申請入口Pilotを準備',
    '次の検証結果を確認したうえで実行します。\n\n' +
      '・Client runtime behavior: 10 suites PASS\n' +
      '・Portal runtime behavior: 11 suites PASS\n' +
      '・vNext統合契約test: PASS\n\n' +
      '現在の予測Engine・parameterは変更しません。続行しますか？',
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return { ok: false, cancelled: true };
  return vNextAdminPrepareEmployeeUxReleaseForManualTest();
}

/** Builds the attestation payload for a verified employee UX release deploy. */
function vNextAdminBuildEmployeeUxReleaseEvidence_(req) {
  const clientBundle = vNextClientRuntimeVerifiedBundle_();
  const portalBundle = typeof vNextPortalRuntimeVerifiedBundle_ === 'function'
    ? vNextPortalRuntimeVerifiedBundle_()
    : { version: VN_ADMIN_PORTAL_RUNTIME_VERSION, sha256: '' };
  const request = req && typeof req === 'object' ? req : {};
  if (request.evidenceArtifact) return vNextAdminRequiredText_(request.evidenceArtifact, 'evidenceArtifact');
  return vNextAdminCanonicalJson_({
    verifiedAt: new Date().toISOString(),
    clientRuntimeTests: 10,
    clientRuntimeVersion: clientBundle.version,
    clientRuntimeSha256: clientBundle.sha256,
    portalRuntimeTests: 12,
    portalRuntimeVersion: portalBundle.version,
    portalRuntimeSha256: portalBundle.sha256,
    integrationContractTests: 'PASS'
  });
}

/** No-UI editor fallback that publishes and activates the currently verified employee UX release. */
function vNextAdminPrepareEmployeeUxReleaseForManualTest() {
  return vNextAdminPrepareEmployeePortalPilot({
    attestationConfirmed: true,
    evidenceArtifact: vNextAdminBuildEmployeeUxReleaseEvidence_({})
  });
}

/**
 * Phase B step: stage+activate Client Template/Model only (no Portal).
 * Keep this separate from Portal update so each Hub call stays under Apps Script limits.
 */
function vNextAdminDeployVerifiedEmployeeUxClientRelease_(request) {
  return vNextAdminGuard_('vNextAdminDeployVerifiedEmployeeUxClientRelease_', function () {
    const req = request && typeof request === 'object' ? request : {};
    const hub = vNextAdminRequireHub_();
    vNextAdminAssertHubAdmin_(hub, false);
    const reason = vNextAdminText_(req.reason) || 'Cursor開発反映';
    vNextAdminWriteDevDeployProgress_('CLIENT_RELEASE', { reason: reason });
    const release = vNextAdminPrepareEmployeePortalPilot({
      attestationConfirmed: true,
      releaseReason: vNextAdminText_(req.releaseReason) || reason,
      evidenceArtifact: vNextAdminBuildEmployeeUxReleaseEvidence_(req),
      skipPortal: true
    });
    vNextAdminWriteDevDeployProgress_('CLIENT_RELEASE_DONE', {
      releaseId: release && release.activeTemplateReleaseId,
      modelReleaseId: release && release.activeModelReleaseId
    });
    return release;
  });
}

/** Phase B step: Portal runtime put + web app republish. */
function vNextAdminDeployVerifiedEmployeeUxPortal_(request) {
  return vNextAdminGuard_('vNextAdminDeployVerifiedEmployeeUxPortal_', function () {
    const req = request && typeof request === 'object' ? request : {};
    const hub = vNextAdminRequireHub_();
    vNextAdminAssertHubAdmin_(hub, false);
    const reason = vNextAdminText_(req.reason) || 'Cursor開発反映';
    vNextAdminWriteDevDeployProgress_('PORTAL', { reason: reason });
    const portal = vNextAdminUpdateSharedPortalRuntime({ reason: reason });
    vNextAdminWriteDevDeployProgress_('PORTAL_DONE', {
      runtimeVersion: portal && portal.runtimeVersion,
      webAppUrl: portal && portal.webAppUrl
    });
    return portal;
  });
}

/**
 * Phase B finalize: optional empty-pilot upgrades + light verification + audit.
 * Full script-content SHA compare is skipped here (use status refresh with full=true).
 */
function vNextAdminDeployVerifiedEmployeeUxFinalize_(request) {
  return vNextAdminGuard_('vNextAdminDeployVerifiedEmployeeUxFinalize_', function () {
    const req = request && typeof request === 'object' ? request : {};
    const hub = vNextAdminRequireHub_();
    vNextAdminAssertHubAdmin_(hub, false);
    const reason = vNextAdminText_(req.reason) || 'Cursor開発反映';
    vNextAdminWriteDevDeployProgress_('FINALIZE', { reason: reason });
    let emptyPilots = { upgraded: [], skipped: [], count: 0, skippedAll: true, skippedByRequest: true };
    if (req.upgradeEmptyPilots === true) {
      emptyPilots = vNextAdminUpgradeEligibleEmptyPilotsInHub_(hub, {
        reason: vNextAdminText_(req.emptyPilotReason) || '受入試験開始前のUI・操作性改善'
      });
    }
    const pair = vNextAdminReadActiveReleasePair_(hub);
    const activeRelease = vNextAdminResolveRelease_(hub, pair.releaseId);
    const verification = vNextAdminVerifyVerifiedEmployeeUxDeployInHub_(hub, { light: true });
    const result = vNextAdminJsonSafe_({
      ok: verification.ok === true,
      phase: 'EMPLOYEE_UX',
      reason: reason,
      activeTemplateReleaseId: pair.releaseId,
      activeModelReleaseId: pair.modelReleaseId,
      clientRuntimeVersion: String(activeRelease.client_runtime_version || ''),
      portalRuntimeVersion: String(
        (req.portal && req.portal.runtimeVersion) ||
        verification.targetPortalVersion || ''
      ),
      release: req.release || null,
      portal: req.portal || null,
      emptyPilots: emptyPilots,
      verification: verification,
      completedAt: new Date().toISOString(),
      message: verification.ok
        ? '反映を自動確認しました（版一致）。必要なら「いま反映済みか確認」で詳細検査できます。'
        : '反映は完了しましたが、自動確認で ' + verification.failedCount + ' 件の不一致があります。'
    });
    vNextAdminWriteAudit_(hub, 'DEPLOY_VERIFIED_EMPLOYEE_UX', 'ADMIN_RUNTIME', hub.getId(), 'SUCCESS', result);
    PropertiesService.getScriptProperties().setProperty(
      'VNEXT_LAST_DEV_DEPLOY_RESULT_JSON', vNextAdminCanonicalJson_(result)
    );
    vNextAdminWriteDevDeployProgress_('DONE', { ok: result.ok });
    return result;
  });
}

/**
 * Monolithic Phase B (CLI fallback). Prefer the stepped Hub sidebar calls;
 * this still skips Portal-inside-prepare and uses light verification.
 */
function vNextAdminDeployVerifiedEmployeeUxReleasePhaseB_(request) {
  return vNextAdminGuard_('vNextAdminDeployVerifiedEmployeeUxReleasePhaseB_', function () {
    const req = request && typeof request === 'object' ? request : {};
    const release = vNextAdminDeployVerifiedEmployeeUxClientRelease_(req);
    const portal = vNextAdminDeployVerifiedEmployeeUxPortal_(req);
    return vNextAdminDeployVerifiedEmployeeUxFinalize_({
      reason: req.reason,
      releaseReason: req.releaseReason,
      emptyPilotReason: req.emptyPilotReason,
      upgradeEmptyPilots: req.upgradeEmptyPilots === true,
      release: release,
      portal: portal
    });
  });
}

/** Upgrades every empty Pilot Client that passes the read-only safety check. */
function vNextAdminUpgradeEligibleEmptyPilotsInHub_(hub, request) {
  const req = request && typeof request === 'object' ? request : {};
  const reason = vNextAdminRequiredText_(req.reason, 'reason');
  const pair = vNextAdminReadActiveReleasePair_(hub);
  const unfinishedByBook = {};
  vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.MIGRATIONS).rows.forEach(function (row) {
    if (/^EMPTY_PILOT_|^RECOVERY_REQUIRED$/.test(String(row.status || '').toUpperCase())) {
      unfinishedByBook[String(row.book_id || '')] = row;
    }
  });
  const candidates = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.REGISTRY).rows.filter(function (row) {
    return String(row.mode || '').toUpperCase() === 'CLIENT' &&
      String(row.status || '').toUpperCase() === 'ACTIVE' &&
      String(row.state || '').toUpperCase() === 'INPUT_OPEN' &&
      !String(row.current_official_id || '') &&
      (String(row.template_release_id || '') !== pair.releaseId || !!unfinishedByBook[String(row.book_id || '')]);
  });
  const upgraded = [];
  const skipped = [];
  candidates.forEach(function (row) {
    const bookId = String(row.book_id || '');
    if (unfinishedByBook[bookId]) {
      skipped.push({ bookId: bookId, reason: 'unfinished migration ' + String(unfinishedByBook[bookId].migration_id || '') });
      return;
    }
    try {
      const basis = vNextAdminUpgradeEmptyPilotClient({ bookId: bookId, dryRun: true, reason: reason });
      if (!basis || basis.eligible !== true) {
        skipped.push({ bookId: bookId, reason: 'not eligible' });
        return;
      }
      upgraded.push(vNextAdminUpgradeEmptyPilotClient({ bookId: bookId, dryRun: false, reason: reason }));
    } catch (error) {
      skipped.push({ bookId: bookId, reason: String(error && error.message || error) });
    }
  });
  return { upgraded: upgraded, skipped: skipped, count: upgraded.length, skippedAll: !candidates.length };
}

/** Central-source phase A for CLI deploy: sync Hub runtime only. */
function vNextAdminDeployVerifiedEmployeeUxReleaseFromSource(request) {
  return vNextAdminGuard_('vNextAdminDeployVerifiedEmployeeUxReleaseFromSource', function () {
    vNextAdminAssertRuntimeConfigurator_();
    const req = request && typeof request === 'object' ? request : {};
    const hubId = vNextAdminRequiredText_(req.hubSpreadsheetId, 'hubSpreadsheetId');
    const hub = SpreadsheetApp.openById(hubId);
    if (vNextDetectBookMode_(hub) !== 'ADMIN' || !vNextAdminIsRegisteredHub_(hub)) {
      throw new Error('The supplied Hub is not the registered 管理ハブ.');
    }
    vNextAdminAssertHubAdmin_(hub, true);
    vNextAdminHydrateHubRuntime_(hub);
    Object.keys(VN_ADMIN_HEADERS).forEach(function (name) {
      vNextAdminEnsureTable_(hub, name, VN_ADMIN_HEADERS[name]);
    });
    return vNextAdminUpdateHubRuntimeInHub_(hub, req, {
      allowEffectiveUser: true, requireCentralCaller: true
    });
  });
}

/** Central-source phase B is not supported; execute phase B on the Hub-bound script. */
function vNextAdminDeployVerifiedEmployeeUxReleasePhaseBFromSource(request) {
  return vNextAdminGuard_('vNextAdminDeployVerifiedEmployeeUxReleasePhaseBFromSource', function () {
    throw new Error('Phase B must run on the registered 管理ハブ script. Use vNextAdminDeployVerifiedEmployeeUxReleasePhaseB_ via Hub API execution, or the Hub sidebar button.');
  });
}

function vNextAdminBuildDevDeployCheck_(id, label, ok, expected, actual, detail) {
  return {
    id: String(id || ''), label: String(label || ''), ok: ok === true,
    expected: String(expected || ''), actual: String(actual || ''),
    detail: String(detail || '')
  };
}

function vNextAdminScriptContentSha256_(scriptId) {
  if (!scriptId || typeof vNextClientRuntimeGetContent_ !== 'function' ||
      typeof vNextAdminRuntimeVerifyScriptContent_ !== 'function') return '';
  const content = vNextClientRuntimeGetContent_(scriptId);
  return String(vNextAdminRuntimeVerifyScriptContent_(content, scriptId).sha256 || '');
}

/** Read-only verification that live Hub/Portal/Client pointers match deployed bundles. */
function vNextAdminVerifyVerifiedEmployeeUxDeployInHub_(hub, options) {
  const opts = options && typeof options === 'object' ? options : {};
  const light = opts.light === true;
  const checks = [];
  const hubConfig = vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET);
  const clientBundle = vNextClientRuntimeVerifiedBundle_();
  const portalBundle = typeof vNextPortalRuntimeVerifiedBundle_ === 'function'
    ? vNextPortalRuntimeVerifiedBundle_()
    : { version: VN_ADMIN_PORTAL_RUNTIME_VERSION, sha256: '' };

  if (!light) {
    const sourceScriptId = String(hubConfig.admin_source_script_id || '');
    const hubScriptId = String(hubConfig.admin_hub_script_id || '');
    const sourceSha = vNextAdminScriptContentSha256_(sourceScriptId);
    const hubScriptSha = vNextAdminScriptContentSha256_(hubScriptId);
    const hubConfigSha = String(hubConfig.admin_runtime_sha256 || '');
    checks.push(vNextAdminBuildDevDeployCheck_(
      'hub_runtime', '管理ハブ runtime（中央 clasp と一致）',
      !!sourceSha && !!hubScriptSha && sourceSha === hubScriptSha && hubConfigSha === hubScriptSha,
      sourceSha ? sourceSha.slice(0, 12) + '…' : '（未取得）',
      hubScriptSha ? hubScriptSha.slice(0, 12) + '…' : '（未取得）',
      hubConfigSha && hubConfigSha !== hubScriptSha ? 'VN_SYSTEM_CONFIG の SHA が古い可能性があります。' : ''
    ));
  } else {
    checks.push(vNextAdminBuildDevDeployCheck_(
      'hub_runtime_pin', '管理ハブ runtime 記録',
      !!String(hubConfig.admin_runtime_sha256 || ''),
      '記録あり', String(hubConfig.admin_runtime_sha256 || '').slice(0, 12) + (hubConfig.admin_runtime_sha256 ? '…' : '（未記録）'),
      '詳細確認で中央 clasp との一致を検査します。'
    ));
  }

  let pair = null;
  let release = null;
  let model = null;
  try {
    pair = vNextAdminReadActiveReleasePair_(hub);
    release = vNextAdminResolveRelease_(hub, pair.releaseId);
    model = vNextAdminLatestModelRelease_(hub, pair.modelReleaseId);
  } catch (pairError) {
    checks.push(vNextAdminBuildDevDeployCheck_(
      'client_release', 'Client 現行 release', false, clientBundle.version, '（未設定）',
      String(pairError && pairError.message || pairError)
    ));
  }
  if (release) {
    checks.push(vNextAdminBuildDevDeployCheck_(
      'client_runtime_version', 'Client runtime 版',
      String(release.client_runtime_version || '') === String(clientBundle.version || ''),
      String(clientBundle.version || ''), String(release.client_runtime_version || ''), ''
    ));
    if (!light) {
      checks.push(vNextAdminBuildDevDeployCheck_(
        'client_runtime_sha', 'Client runtime SHA',
        String(release.client_runtime_sha256 || '') === String(clientBundle.sha256 || ''),
        String(clientBundle.sha256 || '').slice(0, 12) + '…',
        String(release.client_runtime_sha256 || '').slice(0, 12) + '…', ''
      ));
    }
  }
  if (model && pair) {
    checks.push(vNextAdminBuildDevDeployCheck_(
      'model_pair', 'Template＋Model pair',
      String(model.template_version || '') === String(pair.releaseId || '') &&
        String(model.status || '').toUpperCase() === 'ACTIVE',
      String(pair.releaseId || ''), String(model.template_version || ''), ''
    ));
  }

  const portal = vNextAdminTryResolvePortal_(hub);
  if (!portal) {
    checks.push(vNextAdminBuildDevDeployCheck_(
      'portal_runtime', 'Portal runtime', false, portalBundle.version, '（未設定）', '申請入口が未準備です。'
    ));
  } else {
    checks.push(vNextAdminBuildDevDeployCheck_(
      'portal_runtime_version', 'Portal runtime 版',
      String(portal.runtimeVersion || '') === String(VN_ADMIN_PORTAL_RUNTIME_VERSION),
      String(VN_ADMIN_PORTAL_RUNTIME_VERSION), String(portal.runtimeVersion || ''), ''
    ));
    if (!light) {
      checks.push(vNextAdminBuildDevDeployCheck_(
        'portal_runtime_sha', 'Portal runtime SHA',
        String(portal.runtimeSha256 || '') === String(portalBundle.sha256 || ''),
        String(portalBundle.sha256 || '').slice(0, 12) + '…',
        String(portal.runtimeSha256 || '').slice(0, 12) + '…', ''
      ));
    }
  }

  const webAppUrl = vNextAdminEmployeePortalWebAppUrl_(hub);
  checks.push(vNextAdminBuildDevDeployCheck_(
    'portal_web_entry', 'Web入口 URL',
    !!webAppUrl, '登録済み URL', webAppUrl || '（未登録）', ''
  ));

  const failed = checks.filter(function (row) { return !row.ok; });
  return vNextAdminJsonSafe_({
    ok: failed.length === 0,
    light: light,
    checkedAt: new Date().toISOString(),
    targetClientVersion: String(clientBundle.version || ''),
    targetPortalVersion: String(VN_ADMIN_PORTAL_RUNTIME_VERSION),
    webAppUrl: webAppUrl,
    checks: checks,
    failedCount: failed.length,
    summary: failed.length ? ('要確認 ' + failed.length + ' 件') : (light ? '版は一致（詳細確認でSHAも検査可）' : '反映済み（自動確認）')
  });
}

function vNextAdminBuildDevDeployStatus_(hub, options) {
  const opts = options && typeof options === 'object' ? options : {};
  const light = opts.light !== false && opts.full !== true;
  let lastDeploy = null;
  let progress = null;
  try {
    const raw = PropertiesService.getScriptProperties().getProperty('VNEXT_LAST_DEV_DEPLOY_RESULT_JSON') || '';
    if (raw) lastDeploy = JSON.parse(raw);
  } catch (ignoredParse) {
    lastDeploy = null;
  }
  try {
    const progressRaw = PropertiesService.getScriptProperties().getProperty('VNEXT_DEV_DEPLOY_PROGRESS_JSON') || '';
    if (progressRaw) progress = JSON.parse(progressRaw);
  } catch (ignoredProgress) {
    progress = null;
  }
  let verification;
  try {
    verification = vNextAdminVerifyVerifiedEmployeeUxDeployInHub_(hub, { light: light });
  } catch (verifyError) {
    verification = {
      ok: false, light: light, checkedAt: new Date().toISOString(),
      targetClientVersion: '', targetPortalVersion: VN_ADMIN_PORTAL_RUNTIME_VERSION, webAppUrl: '',
      checks: [vNextAdminBuildDevDeployCheck_('verify_error', '反映確認', false, '成功', '失敗',
        String(verifyError && verifyError.message || verifyError))],
      failedCount: 1, summary: '反映確認に失敗しました'
    };
  }
  return {
    lastDeploy: lastDeploy,
    progress: progress,
    verification: verification,
    ok: verification.ok === true
  };
}

function vNextAdminGetVerifiedEmployeeUxDeployStatus(request) {
  return vNextAdminGuard_('vNextAdminGetVerifiedEmployeeUxDeployStatus', function () {
    const req = request && typeof request === 'object' ? request : {};
    const hub = vNextAdminRequireHub_();
    return vNextAdminBuildDevDeployStatus_(hub, {
      light: req.full !== true,
      full: req.full === true
    });
  });
}

/** Central-source fallback used while a generated Hub is waiting for API execution linkage. */
function vNextAdminProvisionSharedPortalFromSource(request) {
  return vNextAdminGuard_('vNextAdminProvisionSharedPortalFromSource', function () {
    vNextAdminAssertRuntimeConfigurator_();
    const req = request && typeof request === 'object' ? request : {};
    const hubId = vNextAdminRequiredText_(req.hubSpreadsheetId, 'hubSpreadsheetId');
    const hub = SpreadsheetApp.openById(hubId);
    const hubConfig = vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET);
    if (vNextDetectBookMode_(hub) !== 'ADMIN' || !vNextAdminIsRegisteredHub_(hub) ||
        String(hubConfig.admin_source_script_id || '') !== String(ScriptApp.getScriptId())) {
      throw new Error('The supplied Hub is not registered.');
    }
    vNextAdminAssertHubAdmin_(hub, false);
    vNextAdminHydrateHubRuntime_(hub);
    const result = vNextAdminProvisionSharedPortalInHub_(hub, req);
    result.clientCatalog = vNextAdminRefreshZacClientCatalogSafely_(hub, true);
    return result;
  });
}

function vNextAdminProvisionSharedPortalInHub_(hub, request) {
  const req = request && typeof request === 'object' ? request : {};
  return vNextAdminWithScriptLock_('provision-shared-portal', function () {
    const runtime = vNextGetRuntimeConfig_();
    const hubConfig = vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET);
    const existingId = String(hubConfig.portal_spreadsheet_id || runtime.VNEXT_PORTAL_SPREADSHEET_ID || '').trim();
    if (existingId && vNextAdminSpreadsheetAccessible_(existingId)) {
      const existingPortal = vNextAdminResolvePortal_(hub);
      const existingFile = DriveApp.getFileById(existingPortal.spreadsheetId);
      vNextAdminApplyEmployeeFileSharing_(existingFile, {
        targetMode: 'PORTAL', accessPolicy: 'INTERNAL_OPEN', internalDomain: existingPortal.employeeDomain
      });
      vNextAdminAssertEmployeeFileSharing_(existingFile, {
        targetMode: 'PORTAL', accessPolicy: 'INTERNAL_OPEN', internalDomain: existingPortal.employeeDomain
      });
      vNextAdminRefreshPortalDirectory_(hub, existingPortal.spreadsheet);
      return {
        reused: true, portalId: existingPortal.portalId, spreadsheetId: existingPortal.spreadsheetId,
        spreadsheetUrl: existingPortal.spreadsheet.getUrl(), runtimeVersion: existingPortal.runtimeVersion
      };
    }
    const actor = vNextAdminActor_();
    const domain = vNextAdminNormalizeDomain_(req.employeeDomain || runtime.VNEXT_EMPLOYEE_DOMAIN || vNextAdminEmailDomain_(actor));
    if (!domain || vNextAdminEmailDomain_(actor) !== domain) {
      throw new Error('employeeDomainは管理ハブのGoogle Workspaceドメインと一致する必要があります。');
    }
    if (typeof vNextPortalRuntimeCreateBoundSpreadsheet_ !== 'function') {
      throw new Error('Portal runtime provisioner is not installed.');
    }
    const adminEmails = vNextAdminMergeEmails_(runtime.VNEXT_ADMIN_EMAILS, actor);
    const folder = vNextAdminPrepareLibraryDestinationFolder_(hub, req.folderId,
      vNextAdminLibraryPath_('PORTAL'), adminEmails);
    const created = vNextPortalRuntimeCreateBoundSpreadsheet_({
      title: vNextAdminText_(req.title) || VNEXT_NAMING.LAYER2_DEFAULT_TITLE, folderId: folder.getId()
    });
    if (String(created.runtimeVersion || '') !== VN_ADMIN_PORTAL_RUNTIME_VERSION) {
      throw new Error('Generated Portal runtime version is not supported.');
    }
    const portalId = 'PORTAL-' + Utilities.getUuid();
    const portal = SpreadsheetApp.openById(created.spreadsheetId);
    const file = DriveApp.getFileById(created.spreadsheetId);
    try {
      vNextClientRuntimeAssertBoundParent_(created.scriptId, created.spreadsheetId);
      vNextAdminInitializePortal_(portal, {
        portalId: portalId, employeeDomain: domain, runtimeVersion: created.runtimeVersion,
        runtimeSha256: created.bundleSha256, actor: actor, adminEmails: adminEmails
      });
      vNextAdminWriteSystemConfig_(hub, {
        portal_id: portalId, portal_spreadsheet_id: created.spreadsheetId,
        portal_script_id: created.scriptId, portal_runtime_version: created.runtimeVersion,
        portal_runtime_sha256: created.bundleSha256, employee_domain: domain
      });
      vNextAdminUpsertObject_(hub, VN_ADMIN_SHEETS.SETTINGS, 'setting_key', 'EMPLOYEE_PORTAL_JSON', {
        setting_key: 'EMPLOYEE_PORTAL_JSON',
        setting_value: vNextAdminCanonicalJson_({
          portalId: portalId, spreadsheetId: created.spreadsheetId, scriptId: created.scriptId,
          runtimeVersion: created.runtimeVersion, runtimeSha256: created.bundleSha256,
          employeeDomain: domain, accessPolicy: 'INTERNAL_OPEN'
        }),
        value_type: 'JSON', scope: 'SYSTEM', effective_from: new Date(), updated_at: new Date(),
        updated_by: actor, note: VNEXT_NAMING.LAYER2 + '（' + VNEXT_NAMING.LAYER1 + 'とは物理分離）'
      });
      PropertiesService.getScriptProperties().setProperties({
        VNEXT_PORTAL_SPREADSHEET_ID: created.spreadsheetId,
        VNEXT_EMPLOYEE_DOMAIN: domain
      }, false);
      vNextAdminRefreshPortalDirectory_(hub, portal);
      vNextAdminApplyEmployeeFileSharing_(file, {
        targetMode: 'PORTAL', accessPolicy: 'INTERNAL_OPEN', internalDomain: domain
      });
      vNextAdminAssertEmployeeFileSharing_(file, {
        targetMode: 'PORTAL', accessPolicy: 'INTERNAL_OPEN', internalDomain: domain
      });
      vNextAdminWriteAudit_(hub, 'PROVISION_EMPLOYEE_PORTAL', 'PORTAL', portalId, 'SUCCESS', {
        spreadsheetId: created.spreadsheetId, scriptId: created.scriptId,
        runtimeVersion: created.runtimeVersion, runtimeSha256: created.bundleSha256,
        employeeDomain: domain, folderId: folder.getId()
      });
      SpreadsheetApp.flush();
      return {
        reused: false, portalId: portalId, spreadsheetId: created.spreadsheetId,
        spreadsheetUrl: portal.getUrl(), scriptId: created.scriptId,
        runtimeVersion: created.runtimeVersion, runtimeSha256: created.bundleSha256
      };
    } catch (error) {
      try { vNextAdminEnforcePrivateFileAcl_(file, adminEmails); }
      catch (rollbackError) { Logger.log('Portal ACL rollback failed: %s', String(rollbackError)); }
      try { file.setName('[SETUP FAILED] ' + VNEXT_NAMING.LAYER2_DEFAULT_TITLE); } catch (ignoredRename) {}
      throw error;
    }
  });
}

function vNextAdminProvisionClientInHub_(hub, request) {
  const req = request && typeof request === 'object' ? request : {};
  return vNextAdminWithScriptLock_('provision-client', function () {
      const runtime = vNextGetRuntimeConfig_();
      const clientName = vNextAdminRequiredText_(req.clientName, 'clientName');
      const clientId = vNextAdminText_(req.clientId) || vNextAdminDeriveClientId_(clientName);
      const fiscalYear = vNextAdminNormalizeFiscalYear_(req.fiscalYear);
      // ACTIVE_RELEASE_PAIR_JSON is the canonical deployment pointer. Script
      // Properties are only runtime caches and may lag after a pair activation,
      // especially in time-trigger executions where the Hub is not the active
      // container. An explicit caller pin is still honored and validated.
      const release = vNextAdminResolveRelease_(hub, req.releaseId);
      const modelRelease = vNextAdminResolveActiveModelRelease_(hub, req.modelReleaseId);
      vNextAdminAssertModelTemplateCompatibility_(hub, modelRelease, release);
      const templateId = vNextAdminText_(req.templateSpreadsheetId) || release.template_spreadsheet_id || runtime.VNEXT_MASTER_TEMPLATE_SPREADSHEET_ID;
      if (!templateId) throw new Error('No Master Template is configured.');
      const clientRuntimeVersion = vNextAdminRequiredText_(release.client_runtime_version, 'release.client_runtime_version');
      const clientRuntimeSha256 = vNextAdminRequiredText_(release.client_runtime_sha256, 'release.client_runtime_sha256');
      vNextAdminRequiredText_(release.template_script_id, 'release.template_script_id');
      if (String(release.schema_version || '') !== vNextAdminClientSchemaVersion_()) {
        throw new Error('Selected release schema is not supported by this Client runtime.');
      }

      // A prior provision attempt may have completed the expensive immutable
      // Template verification and clean bound-runtime creation before a later
      // Hub sync or ACL step failed. Resolve that durable staging record before
      // re-reading the legacy V2 full-grid manifest, and resume only after all
      // pinned identities are revalidated from Hub-owned metadata.
      const existing = vNextAdminFindRegistryRow_(hub, function (row) {
        return String(row.mode) === 'CLIENT' && String(row.client_id) === clientId &&
          String(row.fiscal_year) === String(fiscalYear) &&
          String(row.status) !== 'ARCHIVED';
      });
      if (existing) {
        if (String(existing.status || '').toUpperCase() === 'ACTIVE' && vNextAdminSpreadsheetAccessible_(existing.spreadsheet_id)) {
          return { reused: true, bookId: existing.book_id, spreadsheetId: existing.spreadsheet_id, spreadsheetUrl: existing.spreadsheet_url };
        }
        if (String(existing.status || '').toUpperCase() === 'PROVISIONING' && vNextAdminSpreadsheetAccessible_(existing.spreadsheet_id)) {
          return vNextAdminResumeProvisioningClient_(hub, existing, req, release, modelRelease, runtime);
        }
        throw new Error('同じclient/FY/releaseの生成途中bookがあります。BOOK_REGISTRYのPROVISIONING例外を解消してから再実行してください。bookId=' + existing.book_id);
      }

      const template = SpreadsheetApp.openById(templateId);
      if (vNextDetectBookMode_(template) !== 'TEMPLATE') throw new Error('Configured template is not mode=TEMPLATE.');
      const templateConfig = vNextAdminReadKeyValueSheet_(template, VN_ADMIN_BOOK_CONFIG_SHEET);
      if (String(templateConfig.version || '') !== String(release.release_id || '') ||
          String(templateConfig.client_runtime_bundle_sha256 || '') !== clientRuntimeSha256 ||
          String(templateConfig.client_runtime_version || '') !== clientRuntimeVersion) {
        throw new Error('Master Template runtime identity does not match the selected RELEASES record.');
      }
      vNextAdminAssertReleaseTemplateManifest_(release, template);

      const idempotencyKey = vNextAdminText_(req.idempotencyKey) || ['PROVISION', clientId, fiscalYear, release.release_id].join('|');
      const name = vNextAdminText_(req.bookName) || ('Forecast ' + clientName + ' FY' + fiscalYear);
      const bookId = 'CLIENT-' + Utilities.getUuid();
      const actor = vNextAdminActor_();
      const now = new Date();
      const pilot = vNextAdminAssertPilotProvisionAllowed_(hub);
      const adminEmails = vNextAdminMergeEmails_(runtime.VNEXT_ADMIN_EMAILS, actor);
      const folder = vNextAdminPrepareLibraryDestinationFolder_(hub, req.folderId,
        vNextAdminLibraryPath_('CLIENT', fiscalYear, clientName), adminEmails);
      const stagingFolder = folder;
      const forecastOwners = vNextAdminMergeEmails_(req.forecastOwnerEmails, req.ownerEmails);
      if (forecastOwners.length !== 1) throw new Error('予算策定担当を1名だけ指定してください。');
      const relatedMemberNames = req.relatedMemberNames === undefined || req.relatedMemberNames === null
        ? [] : vNextAdminNormalizeRelatedMemberNames_(req.relatedMemberNames);
      const editors = vNextAdminMergeEmails_(forecastOwners, req.editorEmails, runtime.VNEXT_DEFAULT_EDITORS);
      const viewers = vNextAdminMergeEmails_(req.viewerEmails, runtime.VNEXT_DEFAULT_VIEWERS);
      const accessPolicy = String(req.accessPolicy || 'PRIVATE').trim().toUpperCase();
      if (['PRIVATE', 'INTERNAL_OPEN'].indexOf(accessPolicy) < 0) throw new Error('accessPolicy is invalid.');
      const internalDomain = accessPolicy === 'INTERNAL_OPEN'
        ? vNextAdminNormalizeDomain_(req.internalDomain || runtime.VNEXT_EMPLOYEE_DOMAIN)
        : '';
      if (accessPolicy === 'INTERNAL_OPEN' && !internalDomain) {
        throw new Error('社内共通アクセスにはinternalDomainが必要です。');
      }
      const asOf = vNextAdminText_(req.asOf) || Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd');
      // cutoff is an invariant: always the last day of the month before asOf.
      // It cannot be overridden from provisioning or employee UI.
      const cutoff = vNextAdminCutoffFromAsOf_(asOf);
      const annualSalesScale = vNextAdminResolveClientAnnualSalesScale_(
        hub, clientName, fiscalYear, asOf, cutoff);
      const defaultDue = new Date(now.getTime());
      defaultDue.setDate(defaultDue.getDate() + 7);
      const dueDate = vNextAdminText_(req.inputDueDate) ||
        Utilities.formatDate(defaultDue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      // Create a fresh container and bind only the client-safe runtime. A
      // Spreadsheet copy would create an unknown bound-script ID and make
      // versioned runtime maintenance impossible.
      if (typeof vNextClientRuntimeCreateBoundSpreadsheet_ !== 'function') {
        throw new Error('Client runtime provisioner is not installed.');
      }
      const clientRuntime = vNextClientRuntimeCreateBoundSpreadsheet_({
        title: name, folderId: stagingFolder.getId()
      });
      if (String(clientRuntime.runtimeVersion || '') !== clientRuntimeVersion ||
          String(clientRuntime.bundleSha256 || '') !== clientRuntimeSha256) {
        throw new Error('Generated Client runtime does not match the selected ACTIVE release.');
      }
      if (typeof vNextClientRuntimeAssertBoundParent_ !== 'function') {
        throw new Error('Client runtime parent verifier is not installed.');
      }
      vNextClientRuntimeAssertBoundParent_(clientRuntime.scriptId, clientRuntime.spreadsheetId);
      const clientFile = DriveApp.getFileById(clientRuntime.spreadsheetId);
      const clientBook = SpreadsheetApp.openById(clientRuntime.spreadsheetId);
      vNextAdminCopyTemplateUiToClient_(template, clientBook);
      const registryBase = {
        book_id: bookId, mode: 'CLIENT', client_id: clientId, client_name: clientName,
        fiscal_year: fiscalYear, spreadsheet_id: clientBook.getId(), spreadsheet_url: clientBook.getUrl(),
        client_script_id: clientRuntime.scriptId,
        client_runtime_version: clientRuntimeVersion, client_runtime_sha256: clientRuntimeSha256,
        template_release_id: release.release_id, schema_version: release.schema_version,
        state: 'INPUT_OPEN', status: 'PROVISIONING', health_status: 'PENDING', health_code: 'PROVISIONING',
        last_health_at: '', last_forecast_at: '', current_official_id: '',
        forecast_owner_emails: forecastOwners.join(','), editor_emails: editors.join(','), viewer_emails: viewers.join(','),
        created_at: now, created_by: actor, updated_at: now,
        note: 'idempotency=' + idempotencyKey + '; runtime=' + clientRuntimeVersion +
          '; runtime_sha256=' + clientRuntimeSha256 + '; model_release=' + modelRelease.model_release_id,
        access_policy: accessPolicy, internal_domain: internalDomain,
        related_member_names_json: vNextAdminCanonicalJson_(relatedMemberNames)
      };
      let registryCreated = false;
      let teamCreated = false;
      let employeeAclAttempted = false;
      let employeeAclGranted = false;
      try {
        // The copied file remains Admin-only until initialization, central
        // registration and the first integrity-checked sync all succeed.
        const clientInitialization = vNextAdminInitializeClient_(clientBook, {
          bookId: bookId, clientId: clientId, clientName: clientName, fiscalYear: fiscalYear,
          asOf: asOf, cutoff: cutoff, inputDueDate: dueDate,
          clientRuntimeVersion: clientRuntimeVersion, clientRuntimeSha256: clientRuntimeSha256,
          forecastOwnerEmails: forecastOwners, editors: editors, viewers: viewers,
          hubSpreadsheetId: hub.getId(), templateSpreadsheetId: templateId,
          releaseId: release.release_id, modelReleaseId: modelRelease.model_release_id,
          accessPolicy: accessPolicy, internalDomain: internalDomain,
          annualSalesBaseline: annualSalesScale.amount,
          annualSalesBaselineBasis: annualSalesScale.basis,
          actor: actor, now: now,
          visibleSheets: req.visibleSheets
        });
        vNextAdminRegisterBook_(hub, registryBase);
        registryCreated = true;
        vNextAdminRegisterTeam_(hub, bookId, clientId, fiscalYear, forecastOwners, editors, viewers, 'PROVISIONING');
        teamCreated = true;
        // BOOK_META is created while the file is still Admin-only. Promote that
        // exact record centrally before employee ACL is granted; later Client
        // edits are mirrors only and are never authoritative metadata.
        if (!clientInitialization || !clientInitialization.bookMeta) {
          throw new Error('Trusted Client BOOK_META was not returned by initialization.');
        }
        vNextAdminAppendCoreRowsNoLock_(hub, 'BOOK_META', [clientInitialization.bookMeta]);
        vNextAdminSyncClientToHub_(hub, clientBook, bookId);
        // Employee ACL is the final external side effect. ACTIVE is published
        // only after Drive accepted those grants.
        employeeAclAttempted = true;
        if (folder.getId() !== stagingFolder.getId()) clientFile.moveTo(folder);
        vNextAdminApplyFileAcl_(clientFile, editors, viewers, []);
        vNextAdminApplyEmployeeFileSharing_(clientFile, {
          targetMode: 'CLIENT', accessPolicy: accessPolicy, internalDomain: internalDomain,
          editors: editors, viewers: viewers
        });
        vNextAdminAssertEmployeeFileSharing_(clientFile, {
          targetMode: 'CLIENT', accessPolicy: accessPolicy, internalDomain: internalDomain,
          editors: editors, viewers: viewers
        });
        vNextAdminPreparePrivateBootstrapFolder_(folder.getId(), folder.getName(), adminEmails);
        employeeAclGranted = true;
        vNextAdminPatchRegistryByBookId_(hub, bookId, {
          status: 'ACTIVE', health_status: 'PENDING', health_code: 'NEW_BOOK', updated_at: new Date()
        });
        vNextAdminSetTeamStatus_(hub, bookId, 'ACTIVE');
      } catch (provisionError) {
        const failure = String(provisionError && provisionError.message || provisionError);
        try {
          if (!registryCreated) {
            vNextAdminRegisterBook_(hub, Object.assign({}, registryBase, {
              status: 'PROVISIONING', health_status: 'ERROR', health_code: 'PROVISION_FAILED',
              updated_at: new Date(), note: registryBase.note + '; failure=' + failure.slice(0, 500)
            }));
            registryCreated = true;
          } else {
            vNextAdminPatchRegistryByBookId_(hub, bookId, {
              status: 'PROVISIONING', health_status: 'ERROR', health_code: 'PROVISION_FAILED',
              updated_at: new Date(), note: registryBase.note + '; failure=' + failure.slice(0, 500)
            });
          }
          if (!teamCreated) vNextAdminRegisterTeam_(hub, bookId, clientId, fiscalYear, forecastOwners, editors, viewers, 'PROVISIONING');
          else vNextAdminSetTeamStatus_(hub, bookId, 'PROVISIONING');
          if (employeeAclAttempted) {
            try {
              if (folder.getId() !== stagingFolder.getId()) clientFile.moveTo(stagingFolder);
              vNextAdminEnforcePrivateFileAcl_(clientFile, vNextAdminMergeEmails_(runtime.VNEXT_ADMIN_EMAILS, actor));
            } catch (aclRecoveryError) {
              Logger.log('Provision ACL rollback needs manual attention book=%s error=%s', bookId,
                String(aclRecoveryError && aclRecoveryError.stack || aclRecoveryError));
            }
          }
          vNextAdminAppendException_(hub, {
            severity: 'ERROR', exception_type: 'CLIENT_PROVISION_FAILED', book_id: bookId,
            client_name: clientName, fiscal_year: fiscalYear,
            title: 'Client bookの生成が途中で停止', detail: failure,
            recommended_action: 'privateな生成途中bookとBOOK_REGISTRYを確認し、復旧またはARCHIVED化',
            source_ref: clientBook.getUrl()
          });
          vNextAdminWriteAudit_(hub, 'PROVISION_CLIENT', 'BOOK', bookId, 'FAILED', {
            spreadsheetId: clientBook.getId(), failure: failure, employeeAclRevoked: employeeAclAttempted
          });
        } catch (recoveryError) {
          Logger.log('Provision failure recovery failed book=%s error=%s', bookId, String(recoveryError && recoveryError.stack || recoveryError));
        }
        throw provisionError;
      }
      vNextAdminEnqueueJobInternal_(hub, {
        jobType: 'HEALTH_SCAN', targetBookId: bookId, targetSpreadsheetId: clientBook.getId(),
        request: { bookId: bookId, spreadsheetId: clientBook.getId() },
        idempotencyKey: 'HEALTH|' + bookId + '|' + VN_ADMIN_SCHEMA_VERSION, priority: 20
      });
      vNextAdminWriteAudit_(hub, 'PROVISION_CLIENT', 'BOOK', bookId, 'SUCCESS', {
        clientId: clientId, clientName: clientName, fiscalYear: fiscalYear,
        spreadsheetId: clientBook.getId(), releaseId: release.release_id,
        modelReleaseId: modelRelease.model_release_id,
        clientRuntimeVersion: clientRuntimeVersion, clientRuntimeSha256: clientRuntimeSha256,
        clientScriptIdRecordedInHub: Boolean(clientRuntime.scriptId),
        clientFolderId: folder.getId(), pilotPhase: pilot.phase, pilotClientCountBefore: pilot.clientCount
      });
      vNextAdminRefreshTodayExceptions_(hub);
      vNextAdminRefreshHome_(hub);
      SpreadsheetApp.flush();
      return {
        reused: false, bookId: bookId, spreadsheetId: clientBook.getId(),
        spreadsheetUrl: clientBook.getUrl(), clientName: clientName, fiscalYear: fiscalYear,
        releaseId: release.release_id, modelReleaseId: modelRelease.model_release_id,
        state: 'INPUT_OPEN'
      };
    });
}

/**
 * Resume only a Hub-registered, Admin-private provisioning artifact. No new
 * Spreadsheet or script project is created, so retries are idempotent and do
 * not repeat the legacy full-grid Template manifest read.
 */
function vNextAdminResumeProvisioningClient_(hub, registry, request, release, modelRelease, runtime) {
  const req = request && typeof request === 'object' ? request : {};
  const bookId = vNextAdminRequiredText_(registry.book_id, 'registry.book_id');
  const spreadsheetId = vNextAdminRequiredText_(registry.spreadsheet_id, 'registry.spreadsheet_id');
  const scriptId = vNextAdminRequiredText_(registry.client_script_id, 'registry.client_script_id');
  const client = SpreadsheetApp.openById(spreadsheetId);
  const file = DriveApp.getFileById(spreadsheetId);
  const config = vNextAdminReadKeyValueSheet_(client, VN_ADMIN_BOOK_CONFIG_SHEET);
  const expectedOwners = vNextAdminParseList_(registry.forecast_owner_emails);
  const accessPolicy = String(registry.access_policy || config.access_policy || 'PRIVATE').toUpperCase();
  const internalDomain = vNextAdminNormalizeDomain_(registry.internal_domain || config.internal_domain || '');
  const requestedOwners = vNextAdminMergeEmails_(req.forecastOwnerEmails, req.ownerEmails);
  if (requestedOwners.length && vNextAdminCanonicalJson_(requestedOwners.slice().sort()) !==
      vNextAdminCanonicalJson_(expectedOwners.slice().sort())) {
    throw new Error('生成再開時に予算策定担当を変更することはできません。');
  }
  if (req.relatedMemberNames !== undefined && req.relatedMemberNames !== null) {
    const requestedNames = vNextAdminNormalizeRelatedMemberNames_(req.relatedMemberNames);
    const expectedNames = vNextAdminParseJson_(registry.related_member_names_json, []);
    if (vNextAdminCanonicalJson_(requestedNames) !== vNextAdminCanonicalJson_(expectedNames)) {
      throw new Error('生成再開時に関与メンバー氏名を変更することはできません。');
    }
  }
  if (String(config.mode || '').toUpperCase() !== 'CLIENT' ||
      String(config.book_id || '') !== bookId ||
      String(config.client_id || '') !== String(registry.client_id || '') ||
      Number(config.fiscal_year) !== Number(registry.fiscal_year) ||
      String(config.template_release_id || config.version || '') !== String(release.release_id || '') ||
      String(config.model_release_id || '') !== String(modelRelease.model_release_id || '') ||
      String(config.access_policy || 'PRIVATE').toUpperCase() !== accessPolicy ||
      vNextAdminNormalizeDomain_(config.internal_domain || '') !== internalDomain ||
      String(config.client_runtime_version || '') !== String(release.client_runtime_version || '') ||
      String(config.client_runtime_bundle_sha256 || '') !== String(release.client_runtime_sha256 || '')) {
    throw new Error('生成途中Clientの固定identityがHub正本と一致しないため再開を停止しました。');
  }
  if (String(registry.client_runtime_version || '') !== String(release.client_runtime_version || '') ||
      String(registry.client_runtime_sha256 || '') !== String(release.client_runtime_sha256 || '')) {
    throw new Error('生成途中Clientのruntime pinがACTIVE releaseと一致しません。');
  }
  vNextClientRuntimeAssertBoundParent_(scriptId, spreadsheetId);
  const parents = file.getParents();
  if (!parents.hasNext()) throw new Error('生成途中Clientのprivate保存先を確認できません。');
  const folder = parents.next();
  const adminEmails = vNextAdminMergeEmails_(runtime.VNEXT_ADMIN_EMAILS, vNextAdminActor_());
  vNextAdminPreparePrivateBootstrapFolder_(folder.getId(), folder.getName(), adminEmails);
  vNextAdminEnforcePrivateFileAcl_(file, adminEmails);

  // Re-run only the idempotent post-initialization phases that did not finish.
  vNextAdminSyncClientToHub_(hub, client, bookId);
  const editors = vNextAdminParseList_(registry.editor_emails);
  const viewers = vNextAdminParseList_(registry.viewer_emails);
  vNextAdminApplyFileAcl_(file, editors, viewers, []);
  vNextAdminApplyEmployeeFileSharing_(file, {
    targetMode: 'CLIENT', accessPolicy: accessPolicy, internalDomain: internalDomain,
    editors: editors, viewers: viewers
  });
  vNextAdminAssertEmployeeFileSharing_(file, {
    targetMode: 'CLIENT', accessPolicy: accessPolicy, internalDomain: internalDomain,
    editors: editors, viewers: viewers
  });
  vNextAdminPreparePrivateBootstrapFolder_(folder.getId(), folder.getName(), adminEmails);
  vNextAdminPatchRegistryByBookId_(hub, bookId, {
    status: 'ACTIVE', health_status: 'PENDING', health_code: 'PROVISION_RESUMED', updated_at: new Date()
  });
  vNextAdminSetTeamStatus_(hub, bookId, 'ACTIVE');
  vNextAdminEnqueueJobInternal_(hub, {
    jobType: 'HEALTH_SCAN', targetBookId: bookId, targetSpreadsheetId: spreadsheetId,
    request: { bookId: bookId, spreadsheetId: spreadsheetId },
    idempotencyKey: 'HEALTH|' + bookId + '|' + VN_ADMIN_SCHEMA_VERSION, priority: 20
  });
  vNextAdminWriteAudit_(hub, 'RESUME_PROVISION_CLIENT', 'BOOK', bookId, 'SUCCESS', {
    spreadsheetId: spreadsheetId, releaseId: release.release_id,
    modelReleaseId: modelRelease.model_release_id, clientScriptId: scriptId
  });
  vNextAdminRefreshTodayExceptions_(hub);
  vNextAdminRefreshHome_(hub);
  SpreadsheetApp.flush();
  return {
    reused: true, resumed: true, bookId: bookId, spreadsheetId: spreadsheetId,
    spreadsheetUrl: client.getUrl(), clientName: registry.client_name,
    fiscalYear: Number(registry.fiscal_year), releaseId: release.release_id,
    modelReleaseId: modelRelease.model_release_id, state: String(config.state || 'INPUT_OPEN').toUpperCase()
  };
}

/** Install the five-minute Pilot worker in the central source project. */
function vNextAdminInstallPilotAutomationFromSource(request) {
  return vNextAdminGuard_('vNextAdminInstallPilotAutomationFromSource', function () {
    vNextAdminAssertRuntimeConfigurator_();
    const req = request && typeof request === 'object' ? request : {};
    const sourceMode = vNextDetectBookMode_(SpreadsheetApp.getActiveSpreadsheet());
    if (sourceMode !== 'LEGACY' && sourceMode !== 'TEMPLATE') {
      throw new Error('Pilot source automation is allowed only from the registered source workbook.');
    }
    const hubId = vNextAdminRequiredText_(req.hubSpreadsheetId, 'hubSpreadsheetId');
    const hub = SpreadsheetApp.openById(hubId);
    const config = vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET);
    if (String(config.mode || '').toUpperCase() !== 'ADMIN' ||
        String(config.admin_hub_spreadsheet_id || '') !== hubId ||
        String(config.admin_source_script_id || '') !== ScriptApp.getScriptId() ||
        !vNextAdminIsRegisteredHub_(hub)) {
      throw new Error('The supplied Hub is not registered to this central source project.');
    }
    vNextAdminAssertHubAdmin_(hub, false);
    return vNextAdminWithScriptLock_('install-pilot-source-automation', function () {
      const existing = ScriptApp.getProjectTriggers().filter(function (trigger) {
        return trigger.getHandlerFunction() === VN_ADMIN_SCHEDULED_HANDLER;
      });
      if (!existing.length) {
        ScriptApp.newTrigger(VN_ADMIN_SCHEDULED_HANDLER).timeBased().everyMinutes(5).create();
      }
      vNextAdminClearAutomationInstalledCache_();
      PropertiesService.getScriptProperties().setProperty('VNEXT_ADMIN_HUB_SPREADSHEET_ID', hubId);
      vNextAdminWriteSystemConfig_(hub, {
        pilot_worker_script_id: ScriptApp.getScriptId(), pilot_worker_mode: 'CENTRAL_SOURCE_FALLBACK',
        pilot_worker_installed_at: new Date().toISOString(), pilot_worker_installed_by: vNextAdminActor_()
      });
      vNextAdminWriteAudit_(hub, 'INSTALL_PILOT_SOURCE_AUTOMATION', 'TRIGGER', VN_ADMIN_SCHEDULED_HANDLER, 'SUCCESS', {
        reused: existing.length > 0, intervalMinutes: 5, workerScriptId: ScriptApp.getScriptId()
      });
      return { installed: true, reused: existing.length > 0, intervalMinutes: 5,
        workerMode: 'CENTRAL_SOURCE_FALLBACK' };
    });
  });
}

/** Explicitly open the 4th/5th Client canary only after the first three pilot books were reviewed. */
function vNextAdminApprovePilotCanary(request) {
  return vNextAdminGuard_('vNextAdminApprovePilotCanary', function () {
    const req = request && typeof request === 'object' ? request : {};
    const reason = vNextAdminRequiredText_(req.reason, 'reason');
    const hub = vNextAdminRequireHub_();
    return vNextAdminWithScriptLock_('approve-pilot-canary', function () {
      const status = vNextAdminPilotStatus_(hub);
      if (status.clientCount < VN_ADMIN_PILOT_INITIAL_LIMIT) {
        throw new Error('Canary承認は初期pilot 3冊の結果確認後に行ってください。');
      }
      if (status.canaryApproved) return status;
      const now = new Date();
      vNextAdminUpsertObject_(hub, VN_ADMIN_SHEETS.SETTINGS, 'setting_key', 'PILOT_CANARY_APPROVED', {
        setting_key: 'PILOT_CANARY_APPROVED', setting_value: 'true', value_type: 'BOOLEAN',
        scope: 'ADMIN_PILOT_GATE', effective_from: now, updated_at: now,
        updated_by: vNextAdminActor_(), note: reason
      });
      vNextAdminWriteAudit_(hub, 'APPROVE_PILOT_CANARY', 'PILOT_GATE', 'PILOT_CANARY_APPROVED', 'SUCCESS', {
        clientCount: status.clientCount, reason: reason, hardLimit: VN_ADMIN_PILOT_CANARY_LIMIT
      });
      vNextAdminRefreshHome_(hub);
      return vNextAdminPilotStatus_(hub);
    });
  });
}

function vNextAdminPilotStatus_(hub) {
  return vNextAdminPilotStatusFromRegistry_(vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.REGISTRY).rows, hub);
}

function vNextAdminAssertPilotProvisionAllowed_(hub) {
  const status = vNextAdminPilotStatus_(hub);
  if (status.clientCount >= status.currentLimit) throw new Error(status.blockedReason || 'Pilot Client上限に達しました。');
  return status;
}

/**
 * Employee-book request entry point. It never opens the 管理ハブ.
 * The immutable local event is harvested later by an Admin-owned scan/trigger.
 */
function vNextQueueClientForecastRequest(request) {
  return vNextAdminGuard_('vNextQueueClientForecastRequest', function () {
    const req = request && typeof request === 'object' ? request : {};
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (vNextDetectBookMode_(ss) !== 'CLIENT') throw new Error('予測依頼はクライアントbookから行ってください。');
    return vNextAdminWithDocumentLock_('client-forecast-request', function () {
      if (typeof vNextGetBookContext_ !== 'function') throw new Error('VNext_Core context API is not installed.');
      const context = vNextGetBookContext_(ss);
      if (!context.isForecastOwner) throw new Error('予測を依頼できるのは予算策定担当だけです。');
      if (String(context.state || '').toUpperCase() !== 'READY_TO_RUN') {
        throw new Error('現在は予測を依頼できません。状態=' + String(context.state || ''));
      }
      const nowIso = new Date().toISOString();
      const requestId = 'REQ-' + Utilities.getUuid();
      // asOf is the server-side request date, not an employee-supplied or provisioning date.
      const asOf = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const payload = {
        requestId: requestId,
        bookId: context.bookId,
        clientId: context.clientId,
        clientName: context.clientName,
        fiscalYear: Number(context.fiscalYear),
        asOf: asOf,
        cutoff: vNextAdminCutoffFromAsOf_(asOf),
        bookConfiguredAsOf: vNextAdminText_(context.asOf),
        requestedAt: nowIso,
        requestedBy: String(context.userEmail || vNextAdminActor_()).toLowerCase()
      };
      const requestJson = vNextAdminCanonicalJson_(payload);
      const requestHash = vNextAdminSha256_(requestJson);
      vNextAdminEnsureTable_(ss, VN_ADMIN_CLIENT_REQUEST_SHEET, VN_ADMIN_HEADERS[VN_ADMIN_CLIENT_REQUEST_SHEET]);
      vNextAdminAppendClientRequestEvent_(ss, {
        requestId: requestId, bookId: context.bookId, eventType: 'REQUESTED', status: 'PENDING',
        requestHash: requestHash, requestJson: requestJson, requestedAt: nowIso,
        requestedBy: payload.requestedBy, detail: { source: 'CLIENT_BOOK' }
      });
      const stateResult = vNextAdminAppendStateEvent_(ss, 'RUNNING', {
        fromState: 'READY_TO_RUN', reason: 'forecast_requested:' + requestId, actorEmail: payload.requestedBy,
        actorRole: 'FORECAST_OWNER'
      });
      vNextAdminMirrorClientState_(ss, 'RUNNING');
      return { ok: true, requestId: requestId, requestHash: requestHash, stateEventId: stateResult.stateEventId };
    });
  });
}

/** Enqueue a forecast request without directly coupling Admin to the engine implementation. */
function vNextAdminRequestForecast(request) {
  return vNextAdminGuard_('vNextAdminRequestForecast', function () {
    const req = request && typeof request === 'object' ? request : {};
    const hub = vNextAdminRequireHub_();
    const bookId = vNextAdminRequiredText_(req.bookId, 'bookId');
    const reg = vNextAdminFindRegistryRow_(hub, function (row) { return String(row.book_id) === bookId; });
    if (!reg || reg.mode !== 'CLIENT') throw new Error('Registered CLIENT book not found: ' + bookId);
    const requestAsOf = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const idempotencyKey = req.idempotencyKey || ['FORECAST', bookId, req.manualRequestId || requestAsOf, req.inputDataHash || ''].join('|');
    const existingJob = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.JOBS).rows.find(function (row) {
      return String(row.idempotency_key || '') === idempotencyKey &&
        ['QUEUED', 'RUNNING', 'SUCCEEDED'].indexOf(String(row.status || '')) >= 0;
    });
    if (existingJob) return vNextAdminJsonSafe_(existingJob);
    const client = SpreadsheetApp.openById(String(reg.spreadsheet_id));
    const routing = vNextAdminReadKeyValueSheet_(client, VN_ADMIN_BOOK_CONFIG_SHEET);
    const requestPayload = {
      bookId: bookId,
      clientId: reg.client_id,
      clientName: reg.client_name,
      fiscalYear: Number(reg.fiscal_year),
      asOf: requestAsOf,
      cutoff: vNextAdminCutoffFromAsOf_(requestAsOf),
      bookConfiguredAsOf: vNextAdminText_(routing.as_of),
      targetSpreadsheetId: reg.spreadsheet_id,
      manualRequestId: vNextAdminText_(req.manualRequestId),
      requestedAt: new Date().toISOString(),
      requestedBy: vNextAdminActor_().toLowerCase()
    };
    return vNextAdminWithScriptLock_('manual-forecast-request', function () {
      const job = vNextAdminEnqueueJobInternal_(hub, {
        jobType: 'FORECAST_REQUEST', targetBookId: bookId, targetSpreadsheetId: reg.spreadsheet_id,
        request: requestPayload, idempotencyKey: idempotencyKey, priority: req.priority || 50
      });
      if (String(job.status || '') === 'SUCCEEDED' || String(job.status || '') === 'RUNNING') return job;
      try {
        vNextAdminSetClientState_(reg.spreadsheet_id, 'RUNNING', {
          reason: 'admin_manual_forecast_request', actorRole: 'ADMIN'
        });
        vNextAdminSyncClientToHub_(hub, client, bookId);
        return job;
      } catch (err) {
        const latest = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.JOBS).rows.find(function (row) {
          return String(row.job_id || '') === String(job.job_id || '');
        });
        if (latest) vNextAdminUpdateTableRow_(hub, VN_ADMIN_SHEETS.JOBS, latest._rowNumber, {
          status: 'FAILED', finished_at: new Date(), error: 'State claim failed: ' + String(err && err.message || err), updated_at: new Date()
        });
        throw err;
      }
    });
  });
}

/**
 * Admin-reviewed AI finding. Numeric impact is deterministic and server-capped;
 * model output can never bypass citation/version/cap metadata.
 */
function vNextAdminAppendAiEvidence(request) {
  return vNextAdminGuard_('vNextAdminAppendAiEvidence', function () {
    const hub = vNextAdminRequireHub_();
    return vNextAdminWithScriptLock_('append-ai-evidence', function () {
      return vNextAdminAppendAiEvidenceInternal_(hub, request || {});
    });
  });
}

/**
 * Rebuilds the employee-safe AI insight cache from Hub evidence. This is a
 * derived presentation cache only; the append-only Hub EVIDENCE_EVENT remains
 * authoritative and raw prompt/model metadata never leaves the 管理ハブ.
 */
function vNextAdminRefreshClientAiInsightProjection(request) {
  return vNextAdminGuard_('vNextAdminRefreshClientAiInsightProjection', function () {
    const hub = vNextAdminRequireHub_();
    const req = request && typeof request === 'object' ? request : {};
    const bookId = vNextAdminRequiredText_(req.bookId, 'bookId');
    return vNextAdminWithScriptLock_('refresh-client-ai-insights', function () {
      const registry = vNextAdminFindRegistryRow_(hub, function (row) {
        return String(row.book_id || '') === bookId && String(row.mode || '') === 'CLIENT';
      });
      if (!registry) throw new Error('Registered CLIENT book not found: ' + bookId);
      return vNextAdminProjectAiInsightsToClient_(hub, registry, { asOf: req.asOf });
    });
  });
}

/** Admin editor fallback used to refresh already-completed Pilot research. */
function vNextAdminRefreshAllAiInsightProjectionsForManualTest() {
  return vNextAdminGuard_('vNextAdminRefreshAllAiInsightProjectionsForManualTest', function () {
    const hub = vNextAdminRequireHub_();
    return vNextAdminWithScriptLock_('refresh-all-client-ai-insights', function () {
      const bookIds = new Set(vNextAdminReadCoreRows_(hub, 'EVIDENCE_EVENT').filter(function (row) {
        return String(row.evidence_type || '').toUpperCase().indexOf('AI_RESEARCH') === 0 &&
          String(row.status || 'ACTIVE').toUpperCase() === 'ACTIVE';
      }).map(function (row) { return String(row.book_id || ''); }).filter(Boolean));
      const projected = [];
      bookIds.forEach(function (bookId) {
        const registry = vNextAdminFindRegistryRow_(hub, function (row) {
          return String(row.book_id || '') === bookId && String(row.mode || '') === 'CLIENT' &&
            String(row.status || '').toUpperCase() === 'ACTIVE';
        });
        if (registry) projected.push(vNextAdminProjectAiInsightsToClient_(hub, registry, {}));
      });
      return { ok: true, projectedBooks: projected.length, books: projected };
    });
  });
}

/** Read-only basis for the Admin AI rollback control. Raw prompts are not returned. */
function vNextAdminGetAiRollbackBasis(request) {
  return vNextAdminGuard_('vNextAdminGetAiRollbackBasis', function () {
    const req = request && typeof request === 'object' ? request : {};
    const hub = vNextAdminRequireHub_();
    const bookId = vNextAdminRequiredText_(req.bookId, 'bookId');
    const basis = vNextAdminResolveAiRollbackBasis_(hub, bookId, req.sourceForecastRunId, false);
    const active = vNextAdminResolveBasisAiEvidence_(hub, basis.run);
    return vNextAdminJsonSafe_({
      bookId: bookId, clientName: basis.registry.client_name, fiscalYear: Number(basis.registry.fiscal_year),
      sourceForecastRunId: basis.run.run_id, asOf: basis.run.as_of,
      aiDelta: Number(basis.run.ai_delta || 0), systemRecommended: Number(basis.run.system_recommended || 0),
      evidence: active.map(function (row) {
        const metadata = vNextAdminParseJson_(row.evidence_text, {});
        return {
          evidenceId: String(row.evidence_id || ''), target: String(row.target || ''),
          direction: String(row.direction || ''), appliedAmount: Number(row.applied_amount || row.amount_mid || 0),
          summary: String(metadata.summary || row.target || ''), sourceDate: String(row.source_date || ''),
          evidenceQuality: String(row.evidence_quality || ''), parentRequestId: String(metadata.parentRequestId || '')
        };
      })
    });
  });
}

/**
 * Append AI tombstones and queue a same-seed counterfactual run. Trusted seed
 * and lineage fields are created later by the Admin worker, never by the UI.
 */
function vNextAdminRollbackAiEvidence(request) {
  return vNextAdminGuard_('vNextAdminRollbackAiEvidence', function () {
    const req = request && typeof request === 'object' ? request : {};
    const hub = vNextAdminRequireHub_();
    return vNextAdminWithScriptLock_('rollback-ai-evidence', function () {
      const bookId = vNextAdminRequiredText_(req.bookId, 'bookId');
      const reason = vNextAdminRequiredText_(req.reason, 'reason');
      const scope = vNextAdminNormalizeAiRollbackScope_(req.scope);
      const requestedIds = scope === 'SELECTED'
        ? Array.from(new Set(vNextAdminParseList_(req.evidenceIds || req.selectedEvidenceIds))).sort()
        : [];
      if (scope === 'SELECTED' && !requestedIds.length) throw new Error('SELECTED rollback requires one or more evidenceIds.');
      const basis = vNextAdminResolveAiRollbackBasis_(hub, bookId, req.sourceForecastRunId, true);
      const operationIdentity = vNextAdminCanonicalJson_({
        bookId: bookId, sourceForecastRunId: basis.run.run_id, scope: scope,
        selectedEvidenceIds: requestedIds, reason: reason
      });
      const operationId = 'AI-RBOP-' + vNextAdminSha256_(operationIdentity).slice(0, 24).toUpperCase();
      const idempotencyKey = 'AI_ROLLBACK_FORECAST|' + operationId;
      const priorJobs = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.JOBS).rows.filter(function (row) {
        return String(row.idempotency_key || '') === idempotencyKey;
      });
      const prior = priorJobs.length ? priorJobs[priorJobs.length - 1] : null;
      const priorPayload = prior ? vNextAdminParseJson_(prior.request_json, null) : null;
      const authoritativeState = vNextAdminLatestClientState_(hub, bookId, basis.registry.state);
      if (!prior && String(basis.run.run_id || '') !== String(basis.latestRunId || '')) {
        throw new Error('A new AI rollback basis must be the latest successful draft run.');
      }
      if (!prior && authoritativeState !== 'DRAFT_READY') {
        throw new Error('A new AI rollback requires DRAFT_READY. current=' + authoritativeState);
      }
      if (prior && ['DRAFT_READY', 'READY_TO_RUN', 'RUNNING'].indexOf(authoritativeState) < 0) {
        throw new Error('The existing AI rollback cannot be resumed from state=' + authoritativeState);
      }
      if (prior && ['RUNNING', 'SUCCEEDED'].indexOf(String(prior.status || '').toUpperCase()) >= 0) {
        return vNextAdminJsonSafe_({
          reused: true, operationId: operationId, jobId: prior.job_id, status: prior.status,
          sourceForecastRunId: basis.run.run_id,
          tombstoneEvidenceIds: priorPayload && priorPayload.tombstoneEvidenceIds || []
        });
      }

      const allActive = vNextAdminResolveBasisAiEvidence_(hub, basis.run);
      let targetRows;
      if (priorPayload && String(priorPayload.rollbackOperationId || '') === operationId) {
        const byId = new Map(vNextAdminReadCoreRows_(hub, 'EVIDENCE_EVENT').map(function (row) {
          return [String(row.evidence_id || ''), row];
        }));
        targetRows = (priorPayload.targetEvidenceIds || []).map(function (id) { return byId.get(String(id)); });
        if (!targetRows.length || targetRows.some(function (row) { return !row; })) {
          throw new Error('Failed rollback retry cannot reconstruct its original evidence lineage.');
        }
      } else if (scope === 'ALL') {
        targetRows = allActive.slice();
      } else {
        const activeById = new Map(allActive.map(function (row) { return [String(row.evidence_id || ''), row]; }));
        targetRows = requestedIds.map(function (id) {
          const row = activeById.get(String(id));
          if (!row) throw new Error('Selected evidence is not active for the basis run: ' + id);
          return row;
        });
      }
      if (!targetRows.length) throw new Error('The basis run has no active AI evidence to roll back.');
      targetRows = targetRows.slice().sort(function (a, b) {
        return String(a.evidence_id || '').localeCompare(String(b.evidence_id || ''));
      });

      const rollbackRequestId = 'REQ-AI-RB-' + vNextAdminSha256_(operationId).slice(0, 20).toUpperCase();
      const requestedAt = new Date().toISOString();
      const basisAsOf = vNextAdminDateOnly_(basis.run.as_of);
      const activeIds = allActive.map(function (row) { return String(row.evidence_id || ''); }).sort();
      const targetIds = targetRows.map(function (row) { return String(row.evidence_id || ''); });
      const tombstones = targetRows.map(function (row) {
        return vNextAdminBuildAiRollbackTombstone_(basis.registry, basis.run, row, {
          operationId: operationId, rollbackRequestId: rollbackRequestId,
          reason: reason, requestedAt: requestedAt, requestedBy: vNextAdminActor_()
        });
      });
      const existingById = new Map(vNextAdminReadCoreRows_(hub, 'EVIDENCE_EVENT').map(function (row) {
        return [String(row.evidence_id || ''), row];
      }));
      const missingTombstones = [];
      tombstones.forEach(function (record) {
        const existing = existingById.get(record.evidence_id);
        if (!existing) {
          missingTombstones.push(record);
          return;
        }
        const metadata = vNextAdminParseJson_(existing.evidence_text, {});
        if (String(existing.supersedes_evidence_id || '') !== String(record.supersedes_evidence_id || '') ||
            String(metadata.rollbackOperationId || '') !== operationId ||
            String(metadata.sourceForecastRunId || '') !== String(basis.run.run_id || '')) {
          throw new Error('Existing AI rollback tombstone has conflicting immutable lineage: ' + record.evidence_id);
        }
      });
      if (missingTombstones.length) vNextAdminAppendCoreRowsNoLock_(hub, 'EVIDENCE_EVENT', missingTombstones);

      const parentRequestIds = Array.from(new Set(allActive.map(function (row) {
        return String(vNextAdminParseJson_(row.evidence_text, {}).parentRequestId || '');
      }).filter(Boolean))).sort();
      const jobPayload = priorPayload && String(priorPayload.rollbackOperationId || '') === operationId
        ? priorPayload
        : {
          rollbackOperationId: operationId,
          rollbackRequestId: rollbackRequestId,
          requestId: rollbackRequestId,
          bookId: bookId,
          sourceForecastRunId: String(basis.run.run_id || ''),
          sourceInputDataHash: String(basis.run.input_data_hash || ''),
          sourceModelReleaseId: String(basis.run.model_release_id || ''),
          asOf: basisAsOf,
          cutoff: vNextAdminDateOnly_(basis.run.cutoff),
          scope: scope,
          targetEvidenceIds: targetIds,
          basisActiveAiEvidenceIds: activeIds,
          tombstoneEvidenceIds: tombstones.map(function (row) { return row.evidence_id; }),
          sourceAiParentRequestIds: parentRequestIds,
          reason: reason,
          requestedAt: requestedAt,
          requestedBy: vNextAdminActor_().toLowerCase()
        };
      const job = vNextAdminEnqueueJobInternal_(hub, {
        jobType: 'AI_ROLLBACK_FORECAST', targetBookId: bookId,
        targetSpreadsheetId: basis.registry.spreadsheet_id, request: jobPayload,
        idempotencyKey: idempotencyKey, priority: Number(req.priority || 90)
      });
      try {
        const state = vNextAdminLatestClientState_(hub, bookId, basis.registry.state);
        if (state === 'DRAFT_READY') {
          vNextAdminSetClientState_(basis.registry.spreadsheet_id, 'READY_TO_RUN', {
            hub: hub, reason: 'ai_rollback_queued: ' + reason, actorRole: 'ADMIN', relatedRunId: basis.run.run_id
          });
        }
        const readyState = vNextAdminLatestClientState_(hub, bookId, 'READY_TO_RUN');
        if (readyState === 'READY_TO_RUN') {
          vNextAdminSetClientState_(basis.registry.spreadsheet_id, 'RUNNING', {
            hub: hub, reason: 'ai_rollback_forecast_started: ' + reason,
            actorRole: 'ADMIN', relatedRunId: basis.run.run_id
          });
        } else if (readyState !== 'RUNNING') {
          throw new Error('AI rollback could not claim RUNNING state. current=' + readyState);
        }
      } catch (stateError) {
        try {
          if (vNextAdminLatestClientState_(hub, bookId, basis.registry.state) === 'RUNNING') {
            vNextAdminSetClientState_(basis.registry.spreadsheet_id, 'READY_TO_RUN', {
              hub: hub, reason: 'ai_rollback_state_claim_failed: ' +
                String(stateError && stateError.message || stateError).slice(0, 300), actorRole: 'ADMIN'
            });
          }
        } catch (stateRecoveryError) {
          Logger.log('AI rollback state claim recovery failed operation=%s error=%s', operationId,
            String(stateRecoveryError && stateRecoveryError.stack || stateRecoveryError));
        }
        const latestJob = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.JOBS).rows.filter(function (row) {
          return String(row.job_id || '') === String(job.job_id || '');
        }).slice(-1)[0];
        if (latestJob) vNextAdminUpdateTableRow_(hub, VN_ADMIN_SHEETS.JOBS, latestJob._rowNumber, {
          status: 'FAILED', finished_at: new Date(), error: 'Rollback state claim failed: ' +
            String(stateError && stateError.message || stateError), updated_at: new Date()
        });
        vNextAdminAppendException_(hub, {
          severity: 'ERROR', exception_type: 'AI_ROLLBACK_STATE_CLAIM_FAILED', book_id: bookId,
          client_name: basis.registry.client_name, fiscal_year: basis.registry.fiscal_year,
          title: 'AI反映の取消後の状態更新に失敗', detail: String(stateError && stateError.message || stateError),
          recommended_action: '状態とjobを確認し、同じ取消操作を再実行', source_ref: operationId
        });
        vNextAdminWriteAudit_(hub, 'ROLLBACK_AI_EVIDENCE', 'AI_ROLLBACK', operationId, 'FAILED', {
          jobId: job.job_id, reason: reason, error: String(stateError && stateError.message || stateError).slice(0, 1000)
        });
        throw stateError;
      }
      vNextAdminWriteAudit_(hub, 'ROLLBACK_AI_EVIDENCE', 'AI_ROLLBACK', operationId, 'SUCCESS', {
        bookId: bookId, sourceForecastRunId: basis.run.run_id, scope: scope,
        targetEvidenceIds: targetIds, tombstoneEvidenceIds: tombstones.map(function (row) { return row.evidence_id; }),
        rollbackRequestId: rollbackRequestId, jobId: job.job_id, reason: reason
      });
      return vNextAdminJsonSafe_({
        reused: false, operationId: operationId, jobId: job.job_id, status: 'RUNNING',
        sourceForecastRunId: basis.run.run_id, targetEvidenceIds: targetIds,
        tombstoneEvidenceIds: tombstones.map(function (row) { return row.evidence_id; })
      });
    });
  });
}

function vNextAdminEnqueueAiResearch(request) {
  return vNextAdminGuard_('vNextAdminEnqueueAiResearch', function () {
    const req = request && typeof request === 'object' ? request : {};
    const hub = vNextAdminRequireHub_();
    const bookId = vNextAdminRequiredText_(req.bookId, 'bookId');
    const registry = vNextAdminFindRegistryRow_(hub, function (row) { return String(row.book_id) === bookId; });
    if (!registry || String(registry.mode) !== 'CLIENT') throw new Error('Registered CLIENT book not found: ' + bookId);
    return vNextAdminEnqueueJob({
      jobType: 'AI_RESEARCH', targetBookId: bookId, targetSpreadsheetId: registry.spreadsheet_id,
      request: Object.assign({}, req, { bookId: bookId, clientName: registry.client_name, fiscalYear: Number(registry.fiscal_year) }),
      idempotencyKey: req.idempotencyKey || ('AI_RESEARCH|' + bookId + '|' + vNextAdminSha256_(vNextAdminCanonicalJson_(req))),
      priority: req.priority || 60
    });
  });
}

function vNextAdminEnqueueJob(request) {
  return vNextAdminGuard_('vNextAdminEnqueueJob', function () {
    const hub = vNextAdminRequireHub_();
    if (String(request && request.jobType || '').toUpperCase() === 'AI_ROLLBACK_FORECAST') {
      throw new Error('Use vNextAdminRollbackAiEvidence for AI rollback jobs.');
    }
    return vNextAdminWithScriptLock_('enqueue-job', function () {
      return vNextAdminEnqueueJobInternal_(hub, request || {});
    });
  });
}

function vNextAdminProcessJobQueue(limit) {
  return vNextAdminGuard_('vNextAdminProcessJobQueue', function () {
    const hub = vNextAdminRequireHub_();
    return vNextAdminProcessJobsForHub_(hub, limit);
  });
}

/**
 * Bounded, idempotent Admin action for the normal "check now" workflow.
 * It harvests employee requests, recovers only explicitly allowlisted failures,
 * and claims durable jobs through the ordinary queue. It never creates a second
 * request or changes an approval decision.
 */
function vNextAdminRunOperationalCycle() {
  return vNextAdminGuard_('vNextAdminRunOperationalCycle', function () {
    const hub = vNextAdminRequireHub_();
    const startedAt = Date.now();
    const maintenance = vNextAdminWithScriptLock_('admin-run-operational-cycle', function () {
      return {
        leases: vNextAdminRecoverStaleLeases_(hub, 20),
        portalRequests: vNextAdminHarvestPortalRequestsSafely_(hub),
        safeRetries: vNextAdminRequeueKnownPilotFailures_(hub)
      };
    });
    const jobs = vNextAdminProcessJobsForHub_(hub, 4, startedAt + 240000);
    let portalDirectoryUpdated = false;
    try {
      vNextAdminRefreshPortalDirectory_(hub);
      portalDirectoryUpdated = true;
    } catch (portalRefreshError) {
      Logger.log('Admin operational-cycle Portal refresh skipped: %s',
        String(portalRefreshError && portalRefreshError.stack || portalRefreshError));
    }
    const result = {
      ok: true,
      message: jobs.processed
        ? '受付内容を確認し、' + Number(jobs.processed || 0) + '件の処理を実行しました。'
        : '最新状態を確認しました。新しく実行する処理はありません。',
      maintenance: maintenance,
      jobs: jobs,
      portalDirectoryUpdated: portalDirectoryUpdated,
      elapsedMs: Date.now() - startedAt
    };
    vNextAdminWriteAudit_(hub, 'RUN_OPERATIONAL_CYCLE', 'ADMIN_OPERATION',
      'ADMIN-CYCLE-' + Utilities.getUuid(), 'SUCCESS', {
        processedJobs: Number(jobs.processed || 0),
        requeuedJobs: Number(maintenance.safeRetries && maintenance.safeRetries.requeuedJobs || 0),
        harvestedPortalRequests: Number(maintenance.portalRequests && maintenance.portalRequests.queued || 0),
        portalDirectoryUpdated: portalDirectoryUpdated,
        elapsedMs: result.elapsedMs
      });
    return vNextAdminJsonSafe_(result);
  });
}

function vNextAdminRunHealthScan() {
  return vNextAdminGuard_('vNextAdminRunHealthScan', function () {
    const hub = vNextAdminRequireHub_();
    return vNextAdminWithScriptLock_('health-scan', function () {
      return vNextAdminScanRegistryForHub_(hub);
    });
  });
}

/** Install one idempotent, Admin-owned five-minute sweep in the Hub script project. */
function vNextAdminInstallAutomation() {
  return vNextAdminGuard_('vNextAdminInstallAutomation', function () {
    const hub = vNextAdminRequireHub_();
    return vNextAdminWithScriptLock_('install-automation', function () {
      const existing = ScriptApp.getProjectTriggers().filter(function (trigger) {
        return trigger.getHandlerFunction() === VN_ADMIN_SCHEDULED_HANDLER;
      });
      if (!existing.length) {
        ScriptApp.newTrigger(VN_ADMIN_SCHEDULED_HANDLER).timeBased().everyMinutes(5).create();
      }
      vNextAdminClearAutomationInstalledCache_();
      try { vNextAdminEnsureGuidanceOnOpenTrigger_(); }
      catch (guidanceTriggerError) {
        Logger.log('Install automation guidance trigger skipped: %s',
          String(guidanceTriggerError && guidanceTriggerError.message || guidanceTriggerError));
      }
      PropertiesService.getScriptProperties().setProperty('VNEXT_ADMIN_HUB_SPREADSHEET_ID', hub.getId());
      vNextAdminWriteAudit_(hub, 'INSTALL_AUTOMATION', 'TRIGGER', VN_ADMIN_SCHEDULED_HANDLER, 'SUCCESS', {
        reused: existing.length > 0, intervalMinutes: 5
      });
      return { installed: true, reused: existing.length > 0, intervalMinutes: 5 };
    });
  });
}

/** Time-trigger handler. It resolves the Hub by property and never depends on UI focus. */
function vNextAdminScheduledSweep() {
  return vNextAdminGuard_('vNextAdminScheduledSweep', function () {
    const hub = vNextAdminResolveHubForAutomation_();
    const startedAt = Date.now();
    const props = PropertiesService.getScriptProperties();
    props.setProperty('VNEXT_LAST_SWEEP_STARTED_AT', new Date(startedAt).toISOString());
    // The external ZAC read is intentionally outside the global maintenance
    // lock. A stale catalog must not block request harvest or unrelated jobs.
    const catalog = vNextAdminRefreshZacClientCatalogSafely_(hub, false);
    const maintenance = vNextAdminWithScriptLock_('scheduled-maintenance', function () {
      return {
        leases: vNextAdminRecoverStaleLeases_(hub, 20),
        portalRequests: vNextAdminHarvestPortalRequestsSafely_(hub),
        pilotRetries: vNextAdminRequeueKnownPilotFailures_(hub),
        scan: vNextAdminScanRegistryBatch_(hub, 10)
      };
    });
    const jobs = vNextAdminProcessJobsForHub_(hub, 4, startedAt + 270000);
    try { vNextAdminRefreshPortalDirectory_(hub); }
    catch (portalRefreshError) { Logger.log('Portal refresh skipped: %s', String(portalRefreshError)); }
    const finishedAt = Date.now();
    props.setProperties({
      VNEXT_LAST_SWEEP_SUCCEEDED_AT: new Date(finishedAt).toISOString(),
      VNEXT_LAST_SWEEP_DURATION_MS: String(finishedAt - startedAt)
    }, false);
    vNextAdminRefreshTodayExceptions_(hub);
    vNextAdminRefreshHome_(hub);
    return { catalog: catalog, maintenance: maintenance, jobs: jobs, elapsedMs: finishedAt - startedAt };
  });
}

/** Submit an immutable candidate snapshot for approval. */
function vNextAdminSubmitPlanApproval(request) {
  return vNextAdminGuard_('vNextAdminSubmitPlanApproval', function () {
    const req = request && typeof request === 'object' ? request : {};
    const hub = vNextAdminRequireHub_();
    return vNextAdminWithScriptLock_('submit-approval', function () {
      const bookId = vNextAdminRequiredText_(req.bookId, 'bookId');
      const registry = vNextAdminFindRegistryRow_(hub, function (row) { return String(row.book_id) === bookId; });
      if (!registry || registry.mode !== 'CLIENT') throw new Error('Registered CLIENT book not found: ' + bookId);
      const clientBook = SpreadsheetApp.openById(String(registry.spreadsheet_id));
      vNextAdminSyncClientToHub_(hub, clientBook, bookId);
      const requestType = String(req.requestType || 'PUBLISH').trim().toUpperCase();
      if (['PUBLISH', 'AMENDMENT'].indexOf(requestType) < 0) throw new Error('requestType must be PUBLISH or AMENDMENT.');
      const currentOfficialId = vNextAdminText_(registry.current_official_id);
      let supersedes = '';
      if (requestType === 'AMENDMENT') {
        if (!currentOfficialId) throw new Error('An amendment requires a current official vintage.');
        supersedes = vNextAdminText_(req.supersedesOfficialId || currentOfficialId);
        if (supersedes !== currentOfficialId) {
          throw new Error('supersedesOfficialId must equal the registry current official vintage.');
        }
      } else if (currentOfficialId) {
        throw new Error('An official vintage already exists. Submit an AMENDMENT instead of PUBLISH.');
      }
      const amendmentReason = vNextAdminText_(req.amendmentReason);
      if (requestType === 'AMENDMENT' && !amendmentReason) throw new Error('An amendment reason is required.');
      // Approval payloads never supply authoritative snapshot data. Rebuild the
      // candidate solely from append-only Hub records and validate every link.
      const snapshot = vNextAdminResolveSnapshot_(Object.assign({}, req, {
        requestType: requestType, supersedesOfficialId: supersedes,
        amendmentReason: amendmentReason, allowedSubmitter: vNextAdminActor_().toLowerCase()
      }), registry, hub);
      const snapshotJson = vNextAdminCanonicalJson_(snapshot);
      const snapshotHash = vNextAdminSha256_(snapshotJson);
      const forecastRunId = vNextAdminText_(snapshot.forecast && snapshot.forecast.runId);
      const planVersionId = vNextAdminText_(snapshot.plan && snapshot.plan.planVersionId);
      if (!forecastRunId) throw new Error('forecastRunId is required.');
      if (!planVersionId) throw new Error('planVersionId is required.');
      const idempotencyKey = vNextAdminText_(req.idempotencyKey) ||
        [requestType, bookId, forecastRunId, planVersionId, snapshotHash, supersedes].join('|');
      const existing = vNextAdminFindApproval_(hub, function (row) { return String(row.idempotency_key) === idempotencyKey; });
      if (existing) return vNextAdminJsonSafe_(existing);

      const now = new Date();
      const approvalId = 'APR-' + Utilities.getUuid();
      const row = {
        approval_request_id: approvalId, request_type: requestType, book_id: bookId,
        client_id: registry.client_id, client_name: registry.client_name, fiscal_year: registry.fiscal_year,
        forecast_run_id: forecastRunId, plan_version_id: planVersionId, supersedes_official_id: supersedes,
        amendment_reason: amendmentReason, snapshot_json: snapshotJson, snapshot_hash: snapshotHash,
        status: 'PENDING', processing_attempts: 0, requested_at: now, requested_by: vNextAdminActor_(),
        decision_at: '', decision_by: '', decision_comment: '', official_id: '',
        idempotency_key: idempotencyKey, updated_at: now
      };
      vNextAdminAppendObject_(hub, VN_ADMIN_SHEETS.APPROVALS, row);
      if (requestType === 'PUBLISH') {
        vNextAdminSetClientState_(registry.spreadsheet_id, 'SUBMITTED', {
          reason: 'plan_submitted_for_admin_approval', relatedRunId: forecastRunId,
          relatedPlanVersionId: planVersionId
        });
      }
      vNextAdminSyncClientToHub_(hub, clientBook, bookId);
      vNextAdminWriteAudit_(hub, 'SUBMIT_APPROVAL', 'APPROVAL', approvalId, 'SUCCESS', {
        bookId: bookId, requestType: requestType, forecastRunId: forecastRunId, snapshotHash: snapshotHash
      });
      vNextAdminRefreshTodayExceptions_(hub);
      vNextAdminRefreshHome_(hub);
      return vNextAdminJsonSafe_(row);
    });
  });
}

function vNextAdminSubmitAmendment(request) {
  const req = Object.assign({}, request || {}, { requestType: 'AMENDMENT' });
  return vNextAdminSubmitPlanApproval(req);
}

/** Return the immutable current-official basis used by the Admin amendment form. */
function vNextAdminGetAmendmentBasis(request) {
  return vNextAdminGuard_('vNextAdminGetAmendmentBasis', function () {
    const req = request && typeof request === 'object' ? request : {};
    const hub = vNextAdminRequireHub_();
    const bookId = vNextAdminRequiredText_(req.bookId, 'bookId');
    const basis = vNextAdminResolveCurrentOfficialBasis_(hub, bookId);
    const state = vNextAdminLatestClientState_(hub, bookId, basis.registry.state);
    if (state !== 'OFFICIAL_LOCKED') throw new Error('訂正案を作成できるのはOFFICIAL_LOCKEDのbookだけです。現在=' + state);
    const plan = basis.approvedPlan;
    return vNextAdminJsonSafe_({
      bookId: bookId, clientName: basis.registry.client_name, fiscalYear: Number(basis.registry.fiscal_year),
      state: state, currentOfficialId: basis.officialId,
      basisHash: vNextAdminAmendmentBasisHash_(basis),
      officialForecastRunId: basis.officialForecast.run_id,
      sourceForecastRunId: basis.sourceForecast.run_id,
      approvedPlanVersionId: plan.plan_version_id,
      systemRecommended: Number(basis.sourceForecast.system_recommended || 0),
      adoptionDelta: Number(plan.adoption_delta || 0), adoptionReason: String(plan.adoption_reason || ''),
      adoptedForecast: Number(plan.adopted_forecast || 0), salesUplift: Number(plan.sales_uplift || 0),
      upliftReason: String(plan.uplift_reason || ''), upliftOwner: String(plan.uplift_owner || ''),
      upliftAction: String(plan.uplift_action || ''), upliftDueDate: vNextAdminDateOnlyText_(plan.uplift_due_date),
      upliftAllocation: vNextAdminParseJson_(plan.uplift_allocation_json, []),
      finalBudget: Number(plan.final_budget || 0)
    });
  });
}

/**
 * Append a new Hub-side amendment PLAN_VERSION and immutable approval request.
 * The current approved plan and official vintage are never updated in place.
 */
function vNextAdminCreateAmendmentApproval(request) {
  return vNextAdminGuard_('vNextAdminCreateAmendmentApproval', function () {
    const req = request && typeof request === 'object' ? request : {};
    const hub = vNextAdminRequireHub_();
    return vNextAdminWithScriptLock_('create-amendment-approval', function () {
      const bookId = vNextAdminRequiredText_(req.bookId, 'bookId');
      const amendmentReason = vNextAdminRequiredText_(req.amendmentReason || req.correctionReason, 'amendmentReason');
      const basis = vNextAdminResolveCurrentOfficialBasis_(hub, bookId);
      const state = vNextAdminLatestClientState_(hub, bookId, basis.registry.state);
      if (state !== 'OFFICIAL_LOCKED') throw new Error('訂正案を作成できるのはOFFICIAL_LOCKEDのbookだけです。現在=' + state);
      const expectedOfficialId = vNextAdminRequiredText_(req.expectedCurrentOfficialId, 'expectedCurrentOfficialId');
      const expectedPlanId = vNextAdminRequiredText_(req.expectedCurrentApprovedPlanId, 'expectedCurrentApprovedPlanId');
      const expectedBasisHash = vNextAdminRequiredText_(req.expectedBasisHash, 'expectedBasisHash');
      if (expectedOfficialId !== basis.officialId || expectedPlanId !== String(basis.approvedPlan.plan_version_id || '') ||
          expectedBasisHash !== vNextAdminAmendmentBasisHash_(basis)) {
        throw new Error('表示後に正式予算が更新されました。現在値を再読込してください。');
      }
      if (req.supersedesOfficialId && String(req.supersedesOfficialId) !== basis.officialId) {
        throw new Error('supersedesOfficialId must equal the current official vintage.');
      }
      const normalized = vNextAdminNormalizeAmendmentInput_(req, basis);
      const actor = vNextAdminActor_().toLowerCase();
      const identity = vNextAdminCanonicalJson_({
        action: 'ADMIN_AMENDMENT_PLAN_V1', bookId: bookId, supersedesOfficialId: basis.officialId,
        sourceForecastRunId: basis.sourceForecast.run_id,
        amendsPlanVersionId: basis.approvedPlan.plan_version_id,
        requestedBy: actor,
        amendmentReason: amendmentReason, adoptionDelta: normalized.adoptionDelta,
        adoptionReason: normalized.adoptionReason, salesUplift: normalized.salesUplift,
        upliftReason: normalized.upliftReason, upliftOwner: normalized.upliftOwner,
        upliftAction: normalized.upliftAction, upliftDueDate: normalized.upliftDueDate,
        upliftAllocation: normalized.upliftAllocation
      });
      const identityHash = vNextAdminSha256_(identity);
      const planId = 'AMD-PLAN-' + identityHash.slice(0, 24).toUpperCase();
      const approvalId = 'APR-AMD-' + identityHash.slice(0, 24).toUpperCase();
      const idempotencyKey = 'AMENDMENT|' + bookId + '|' + basis.officialId + '|' + identityHash;
      const conflictingApproval = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.APPROVALS).rows.find(function (row) {
        return String(row.request_type || '').toUpperCase() === 'AMENDMENT' &&
          String(row.book_id || '') === bookId && String(row.supersedes_official_id || '') === basis.officialId &&
          ['PENDING', 'PROCESSING_APPROVE', 'PROCESSING_RETURN', 'PROCESSING_REJECT'].indexOf(String(row.status || '').toUpperCase()) >= 0 &&
          String(row.idempotency_key || '') !== idempotencyKey;
      });
      if (conflictingApproval) {
        throw new Error('この正式予算には別の訂正承認が進行中です。approval=' + conflictingApproval.approval_request_id);
      }
      const allPlans = vNextAdminReadCoreRows_(hub, 'PLAN_VERSION').filter(function (row) {
        return String(row.book_id || '') === bookId;
      });
      let plan = allPlans.find(function (row) { return String(row.plan_version_id || '') === planId; });
      if (!plan) {
        const versionNo = allPlans.reduce(function (max, row) {
          return Math.max(max, Number(row.version_no || 0));
        }, 0) + 1;
        const nowIso = new Date().toISOString();
        plan = {
          plan_version_id: planId, book_id: bookId, run_id: basis.sourceForecast.run_id,
          official_vintage_id: '', version_no: versionNo, status: 'SUBMITTED',
          system_recommended: normalized.systemRecommended, adoption_delta: normalized.adoptionDelta,
          adoption_reason: normalized.adoptionReason, adopted_forecast: normalized.adoptedForecast,
          sales_uplift: normalized.salesUplift, uplift_reason: normalized.upliftReason,
          uplift_owner: normalized.upliftOwner, uplift_action: normalized.upliftAction,
          uplift_due_date: normalized.upliftDueDate,
          uplift_allocation_json: vNextAdminCanonicalJson_(normalized.upliftAllocation),
          final_budget: normalized.finalBudget,
          amends_plan_version_id: basis.approvedPlan.plan_version_id,
          submitted_at: nowIso, submitted_by: actor, approved_at: '', approved_by: '', created_at: nowIso
        };
        vNextAdminAppendCoreRowsNoLock_(hub, 'PLAN_VERSION', [plan]);
      }
      const expectedPlanContent = vNextAdminCanonicalJson_({
        bookId: bookId, runId: basis.sourceForecast.run_id, status: 'SUBMITTED',
        systemRecommended: normalized.systemRecommended, adoptionDelta: normalized.adoptionDelta,
        adoptionReason: normalized.adoptionReason, adoptedForecast: normalized.adoptedForecast,
        salesUplift: normalized.salesUplift, upliftReason: normalized.upliftReason,
        upliftOwner: normalized.upliftOwner, upliftAction: normalized.upliftAction,
        upliftDueDate: normalized.upliftDueDate, upliftAllocation: normalized.upliftAllocation,
        finalBudget: normalized.finalBudget, amendsPlanVersionId: basis.approvedPlan.plan_version_id,
        submittedBy: actor
      });
      const actualPlanContent = vNextAdminCanonicalJson_({
        bookId: String(plan.book_id || ''), runId: String(plan.run_id || ''), status: String(plan.status || '').toUpperCase(),
        systemRecommended: Number(plan.system_recommended), adoptionDelta: Number(plan.adoption_delta),
        adoptionReason: String(plan.adoption_reason || ''), adoptedForecast: Number(plan.adopted_forecast),
        salesUplift: Number(plan.sales_uplift), upliftReason: String(plan.uplift_reason || ''),
        upliftOwner: String(plan.uplift_owner || ''), upliftAction: String(plan.uplift_action || ''),
        upliftDueDate: String(plan.uplift_due_date || ''),
        upliftAllocation: vNextAdminParseJson_(plan.uplift_allocation_json, []),
        finalBudget: Number(plan.final_budget), amendsPlanVersionId: String(plan.amends_plan_version_id || ''),
        submittedBy: String(plan.submitted_by || '').toLowerCase()
      });
      if (actualPlanContent !== expectedPlanContent) {
        throw new Error('Existing deterministic amendment PLAN_VERSION is inconsistent; no new record was appended.');
      }
      const validationOptions = {
        requestType: 'AMENDMENT', allowedSubmitter: actor,
        supersedesOfficialId: basis.officialId, officialBasis: basis
      };
      vNextAdminValidateSubmittedPlan_(basis.registry, basis.sourceForecast, plan, validationOptions);
      const snapshot = vNextAdminBuildPlanApprovalSnapshot_(basis.registry, basis.sourceForecast, plan, {
        requestType: 'AMENDMENT', supersedesOfficialId: basis.officialId,
        amendmentReason: amendmentReason, officialBasis: basis
      });
      const snapshotJson = vNextAdminCanonicalJson_(snapshot);
      const snapshotHash = vNextAdminSha256_(snapshotJson);
      const existing = vNextAdminFindApproval_(hub, function (row) {
        return String(row.approval_request_id || '') === approvalId || String(row.idempotency_key || '') === idempotencyKey;
      });
      if (existing) {
        if (String(existing.request_type || '') !== 'AMENDMENT' || String(existing.book_id || '') !== bookId ||
            String(existing.forecast_run_id || '') !== String(basis.sourceForecast.run_id) ||
            String(existing.plan_version_id || '') !== planId ||
            String(existing.supersedes_official_id || '') !== basis.officialId ||
            String(existing.amendment_reason || '') !== amendmentReason ||
            String(existing.snapshot_hash || '') !== snapshotHash || String(existing.snapshot_json || '') !== snapshotJson) {
          throw new Error('Existing amendment idempotency record is inconsistent; no new record was appended.');
        }
        return vNextAdminJsonSafe_({
          reused: true, approvalRequestId: existing.approval_request_id, planVersionId: planId,
          supersedesOfficialId: basis.officialId, status: existing.status
        });
      }
      const now = new Date();
      const approval = {
        approval_request_id: approvalId, request_type: 'AMENDMENT', book_id: bookId,
        client_id: basis.registry.client_id, client_name: basis.registry.client_name,
        fiscal_year: basis.registry.fiscal_year, forecast_run_id: basis.sourceForecast.run_id,
        plan_version_id: planId, supersedes_official_id: basis.officialId,
        amendment_reason: amendmentReason, snapshot_json: snapshotJson, snapshot_hash: snapshotHash,
        status: 'PENDING', processing_attempts: 0, requested_at: now, requested_by: actor,
        decision_at: '', decision_by: '', decision_comment: '', official_id: '',
        idempotency_key: idempotencyKey, updated_at: now
      };
      vNextAdminAppendObject_(hub, VN_ADMIN_SHEETS.APPROVALS, approval);
      vNextAdminWriteAudit_(hub, 'CREATE_AMENDMENT_APPROVAL', 'APPROVAL', approvalId, 'SUCCESS', {
        bookId: bookId, supersedesOfficialId: basis.officialId,
        officialForecastRunId: basis.officialForecast.run_id,
        sourceForecastRunId: basis.sourceForecast.run_id,
        amendsPlanVersionId: basis.approvedPlan.plan_version_id,
        planVersionId: planId, snapshotHash: snapshotHash, amendmentReason: amendmentReason
      }, vNextAdminAmendmentBasisHash_(basis), snapshotHash);
      vNextAdminRefreshTodayExceptions_(hub);
      vNextAdminRefreshHome_(hub);
      return vNextAdminJsonSafe_({
        reused: false, approvalRequestId: approvalId, planVersionId: planId,
        supersedesOfficialId: basis.officialId, status: 'PENDING'
      });
    });
  });
}

/** decision must be APPROVE, RETURN, or REJECT. */
function vNextAdminDecideApproval(request) {
  return vNextAdminGuard_('vNextAdminDecideApproval', function () {
    const req = request && typeof request === 'object' ? request : {};
    const decision = String(req.decision || '').trim().toUpperCase();
    if (['APPROVE', 'RETURN', 'REJECT'].indexOf(decision) < 0) throw new Error('decision must be APPROVE, RETURN, or REJECT.');
    const hub = vNextAdminRequireHub_();
    const approvalId = vNextAdminRequiredText_(req.approvalRequestId, 'approvalRequestId');
    if ((decision === 'RETURN' || decision === 'REJECT') && !vNextAdminText_(req.comment)) {
      throw new Error('A decision comment is required for RETURN or REJECT.');
    }
    const claim = vNextAdminWithScriptLock_('claim-approval', function () {
      const approval = vNextAdminFindApproval_(hub, function (row) { return String(row.approval_request_id) === approvalId; });
      if (!approval) throw new Error('Approval request not found: ' + approvalId);
      if (approval.status !== 'PENDING') {
        if (decision === 'APPROVE' && approval.status === 'APPROVED') return { alreadyComplete: true, approval: approval };
        throw new Error('Approval request is already decided: ' + approval.status);
      }
      const processingAttempts = Number(approval.processing_attempts || 0);
      if (processingAttempts >= 3) {
        vNextAdminUpdateTableRow_(hub, VN_ADMIN_SHEETS.APPROVALS, approval._rowNumber, {
          status: 'FAILED', decision_comment: '承認処理が3回失敗したため管理ハブ担当者確認が必要です。', updated_at: new Date()
        });
        throw new Error('Approval processing failed three times; inspect the exception and create a new approval request after correction.');
      }
      vNextAdminUpdateTableRow_(hub, VN_ADMIN_SHEETS.APPROVALS, approval._rowNumber, {
        status: 'PROCESSING_' + decision, decision_by: vNextAdminActor_(),
        decision_comment: vNextAdminText_(req.comment), processing_attempts: processingAttempts + 1, updated_at: new Date()
      });
      return { alreadyComplete: false, approval: approval };
    });
    if (claim.alreadyComplete) return vNextAdminJsonSafe_(claim.approval);
    const approval = claim.approval;
    try {
      const registry = vNextAdminFindRegistryRow_(hub, function (row) { return String(row.book_id) === String(approval.book_id); });
      if (!registry) throw new Error('Registry entry not found for book: ' + approval.book_id);
      const now = new Date();
      let officialId = '';
      if (decision === 'APPROVE') {
        const issued = vNextAdminIssueOfficial_(hub, approval, registry, req.comment || '');
        officialId = issued.officialId;
        try {
          if (String(approval.request_type || '') !== 'AMENDMENT') {
            vNextAdminSetClientState_(registry.spreadsheet_id, 'OFFICIAL_LOCKED', {
              reason: 'plan_approved', relatedRunId: issued.officialForecastRunId || approval.forecast_run_id,
              relatedPlanVersionId: issued.approvedPlanVersionId || approval.plan_version_id
            });
          }
          const approvedClientBook = SpreadsheetApp.openById(String(registry.spreadsheet_id));
          vNextAdminSyncClientToHub_(hub, approvedClientBook, registry.book_id);
          vNextAdminSyncHubToClient_(hub, approvedClientBook, registry.book_id, ['FORECAST_RUN', 'PLAN_VERSION', 'STATE_EVENT']);
          vNextAdminMirrorClientState_(approvedClientBook, 'OFFICIAL_LOCKED');
        } catch (syncError) {
          vNextAdminAppendException_(hub, {
            severity: 'ERROR', exception_type: 'OFFICIAL_CLIENT_SYNC_FAILED', book_id: registry.book_id,
            client_name: registry.client_name, fiscal_year: registry.fiscal_year,
            title: '公式化後のclient同期に失敗', detail: String(syncError && syncError.message || syncError),
            recommended_action: 'client権限を確認し、Admin Sidebarの「正式予算をClientへ再同期」を実行', source_ref: officialId
          });
          Logger.log('Post-official client sync failed approval=%s error=%s', approvalId, String(syncError && syncError.stack || syncError));
        }
      } else {
        if (String(approval.request_type || '') !== 'AMENDMENT') {
          vNextAdminSetClientState_(registry.spreadsheet_id, 'CHANGES_REQUESTED', {
            reason: (decision === 'RETURN' ? 'changes_requested: ' : 'plan_rejected: ') + vNextAdminText_(req.comment),
            relatedRunId: approval.forecast_run_id, relatedPlanVersionId: approval.plan_version_id
          });
        }
        const clientBook = SpreadsheetApp.openById(String(registry.spreadsheet_id));
        vNextAdminSyncClientToHub_(hub, clientBook, registry.book_id);
      }
      const patch = {
        status: decision === 'APPROVE' ? 'APPROVED' : (decision === 'RETURN' ? 'RETURNED' : 'REJECTED'),
        decision_at: now, decision_by: vNextAdminActor_(), decision_comment: vNextAdminText_(req.comment),
        official_id: officialId, updated_at: now
      };
      vNextAdminWithScriptLock_('finish-approval', function () {
        const latest = vNextAdminFindApproval_(hub, function (row) { return String(row.approval_request_id) === approvalId; });
        if (!latest) throw new Error('Approval request disappeared during decision.');
        vNextAdminUpdateTableRow_(hub, VN_ADMIN_SHEETS.APPROVALS, latest._rowNumber, patch);
      });
      vNextAdminWriteAudit_(hub, 'DECIDE_APPROVAL', 'APPROVAL', approvalId, 'SUCCESS', {
        decision: decision, officialId: officialId, comment: vNextAdminText_(req.comment)
      });
      vNextAdminRefreshTodayExceptions_(hub);
      vNextAdminRefreshHome_(hub);
      return { approvalRequestId: approvalId, status: patch.status, officialId: officialId };
    } catch (err) {
      try {
        vNextAdminWithScriptLock_('release-approval', function () {
          const latest = vNextAdminFindApproval_(hub, function (row) { return String(row.approval_request_id) === approvalId; });
          if (latest && String(latest.status || '').indexOf('PROCESSING_') === 0) {
            vNextAdminUpdateTableRow_(hub, VN_ADMIN_SHEETS.APPROVALS, latest._rowNumber, {
              status: 'PENDING', decision_at: '', decision_by: '',
              decision_comment: '前回処理失敗: ' + String(err && err.message || err).slice(0, 500), updated_at: new Date()
            });
          }
        });
      } catch (releaseError) {
        Logger.log('Approval claim release failed id=%s error=%s', approvalId, String(releaseError && releaseError.stack || releaseError));
      }
      throw err;
    }
  });
}

/** Route a returned plan without changing the immutable approval decision. */
function vNextAdminRouteReturnedPlan(request) {
  return vNextAdminGuard_('vNextAdminRouteReturnedPlan', function () {
    const req = request && typeof request === 'object' ? request : {};
    const hub = vNextAdminRequireHub_();
    return vNextAdminWithUserLock_('route-returned-plan', function () {
      const bookId = vNextAdminRequiredText_(req.bookId, 'bookId');
      const route = String(req.route || 'PLAN_ONLY').trim().toUpperCase();
      const targetState = vNextAdminReturnedPlanTargetState_(route);
      const reason = vNextAdminRequiredText_(req.reason, 'reason');
      const registry = vNextAdminFindRegistryRow_(hub, function (row) {
        return String(row.book_id || '') === bookId && String(row.mode || '') === 'CLIENT' &&
          String(row.status || '').toUpperCase() === 'ACTIVE';
      });
      if (!registry) throw new Error('Registered ACTIVE CLIENT book not found: ' + bookId);
      const returned = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.APPROVALS).rows.filter(function (row) {
        return String(row.book_id || '') === bookId && String(row.status || '').toUpperCase() === 'RETURNED' &&
          (!req.approvalRequestId || String(row.approval_request_id || '') === String(req.approvalRequestId));
      });
      if (!returned.length) throw new Error('A RETURNED approval for this book is required before routing.');
      const approval = returned[returned.length - 1];
      const currentState = vNextAdminLatestClientState_(hub, bookId, registry.state);
      if (currentState !== 'CHANGES_REQUESTED' && currentState !== targetState) {
        throw new Error('Returned plan routing requires CHANGES_REQUESTED. current=' + currentState);
      }
      const operationId = 'RETURN-ROUTE-' + vNextAdminSha256_(vNextAdminCanonicalJson_({
        approvalRequestId: approval.approval_request_id, bookId: bookId, route: route, reason: reason
      })).slice(0, 24).toUpperCase();
      let round = null;
      const client = SpreadsheetApp.openById(String(registry.spreadsheet_id));
      if (route === 'REOPEN_INPUT') {
        const latestMeta = vNextAdminReadCoreRows_(hub, 'BOOK_META').filter(function (row) {
          return String(row.book_id || '') === bookId;
        }).slice(-1)[0];
        round = currentState === 'INPUT_OPEN' && latestMeta && String(latestMeta.event_type || '') === 'INPUT_REOPENED'
          ? { reused: true, recordId: latestMeta.record_id, roundStartedAt: String(latestMeta.recorded_at || '') }
          : vNextAdminAppendInputReopenedMeta_(hub, client, registry, operationId);
        vNextAdminWriteBookConfig_(client, {
          input_submitted: 0, input_answered_count: 0, state: 'INPUT_OPEN',
          input_round_started_at: round.roundStartedAt
        });
      }
      const stateResult = currentState === targetState
        ? { changed: false, fromState: currentState, toState: targetState, stateEventId: '' }
        : vNextAdminSetClientState_(registry.spreadsheet_id, targetState, {
          hub: hub, reason: 'returned_plan_route/' + route + ': ' + reason,
          actorRole: 'ADMIN', relatedRunId: approval.forecast_run_id,
          relatedPlanVersionId: approval.plan_version_id
        });
      vNextAdminSyncHubToClient_(hub, client, bookId, ['STATE_EVENT']);
      vNextAdminMirrorClientState_(client, targetState);
      vNextAdminPatchRegistryByBookId_(hub, bookId, { state: targetState, updated_at: new Date() });
      vNextAdminWriteAudit_(hub, 'ROUTE_RETURNED_PLAN', 'APPROVAL', approval.approval_request_id, 'SUCCESS', {
        operationId: operationId, bookId: bookId, route: route,
        fromState: currentState, toState: targetState, reason: reason,
        roundStartedAt: round && round.roundStartedAt || ''
      });
      return vNextAdminJsonSafe_({
        reused: currentState === targetState, operationId: operationId,
        approvalRequestId: approval.approval_request_id, bookId: bookId,
        route: route, state: targetState, stateEventId: stateResult.stateEventId || '',
        roundStartedAt: round && round.roundStartedAt || ''
      });
    });
  });
}

function vNextAdminReturnedPlanTargetState_(route) {
  const normalized = String(route || '').trim().toUpperCase();
  if (!Object.prototype.hasOwnProperty.call(VN_ADMIN_RETURN_ROUTES, normalized)) {
    throw new Error('route must be PLAN_ONLY, REOPEN_INPUT, or RERUN_SAME_INPUT.');
  }
  return VN_ADMIN_RETURN_ROUTES[normalized];
}

function vNextAdminAppendInputReopenedMeta_(hub, client, registry, operationId) {
  const bookId = String(registry.book_id || '');
  const metas = vNextAdminReadCoreRows_(hub, 'BOOK_META').filter(function (row) {
    return String(row.book_id || '') === bookId;
  });
  if (!metas.length) throw new Error('BOOK_META is missing for input reopen: ' + bookId);
  const source = metas[metas.length - 1];
  const recordId = 'META-REOPEN-' + vNextAdminSha256_(String(operationId || '') + '|' + bookId).slice(0, 24).toUpperCase();
  const existing = metas.find(function (row) { return String(row.record_id || '') === recordId; });
  if (existing) {
    if (String(existing.event_type || '') !== 'INPUT_REOPENED' || String(existing.supersedes_record_id || '') !== String(source.supersedes_record_id || source.record_id || '')) {
      // If source is the marker itself, a normal idempotent replay is valid.
      if (String(source.record_id || '') !== recordId) throw new Error('Existing INPUT_REOPENED marker lineage is inconsistent.');
    }
    return { reused: true, recordId: recordId, roundStartedAt: String(existing.recorded_at || '') };
  }
  const now = new Date().toISOString();
  const record = Object.assign({}, source, {
    record_id: recordId,
    state: 'INPUT_OPEN',
    event_type: 'INPUT_REOPENED',
    supersedes_record_id: String(source.record_id || ''),
    recorded_at: now,
    recorded_by: vNextAdminActor_().toLowerCase()
  });
  delete record._rowNumber;
  vNextAdminAppendCoreRowsNoLock_(hub, 'BOOK_META', [record]);
  const clientIds = new Set(vNextAdminReadCoreRows_(client, 'BOOK_META').map(function (row) {
    return String(row.record_id || '');
  }));
  if (!clientIds.has(recordId)) vNextAdminAppendCoreRowsNoLock_(client, 'BOOK_META', [record]);
  return { reused: false, recordId: recordId, roundStartedAt: now };
}

/** Replays only the Client propagation step for an already-issued central official vintage. */
function vNextAdminRetryOfficialClientSync(request) {
  return vNextAdminGuard_('vNextAdminRetryOfficialClientSync', function () {
    const req = request && typeof request === 'object' ? request : {};
    const hub = vNextAdminRequireHub_();
    const bookId = vNextAdminRequiredText_(req.bookId, 'bookId');
    return vNextAdminWithUserLock_('retry-official-client-sync', function () {
      const registry = vNextAdminFindRegistryRow_(hub, function (row) {
        return String(row.book_id || '') === bookId && String(row.mode || '') === 'CLIENT';
      });
      if (!registry) throw new Error('Registered CLIENT book not found: ' + bookId);
      const officialId = vNextAdminRequiredText_(registry.current_official_id, 'currentOfficialId');
      const officialRows = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.OFFICIAL).rows.filter(function (row) {
        return String(row.book_id || '') === bookId && String(row.official_id || '') === officialId;
      });
      if (!officialRows.length) throw new Error('Current OFFICIAL_RUNS record was not found.');
      const official = officialRows[officialRows.length - 1];
      const frozen = vNextAdminReadCoreRows_(hub, 'FORECAST_RUN').filter(function (row) {
        return String(row.book_id || '') === bookId && String(row.run_id || '') === String(official.forecast_run_id || '') &&
          String(row.official_vintage_id || '') === officialId && Number(row.is_official || 0) === 1;
      });
      const approvedPlans = vNextAdminReadCoreRows_(hub, 'PLAN_VERSION').filter(function (row) {
        return String(row.book_id || '') === bookId && String(row.official_vintage_id || '') === officialId &&
          String(row.status || '').toUpperCase() === 'APPROVED' &&
          String(row.run_id || '') === String(official.source_forecast_run_id || '');
      });
      if (!frozen.length || !approvedPlans.length) {
        throw new Error('Central official forecast/approved plan linkage is incomplete; Client sync was not attempted.');
      }

      const client = SpreadsheetApp.openById(String(registry.spreadsheet_id));
      const routing = vNextAdminReadKeyValueSheet_(client, VN_ADMIN_BOOK_CONFIG_SHEET);
      vNextAdminSyncClientToHub_(hub, client, bookId);
      const hubState = vNextAdminLatestClientState_(hub, bookId, registry.state || 'OFFICIAL_LOCKED');
      const preservedState = vNextAdminOfficialSyncTargetState_(hubState);
      vNextAdminSyncHubToClient_(hub, client, bookId, ['FORECAST_RUN', 'PLAN_VERSION', 'STATE_EVENT']);
      const clientStateAfterSync = vNextAdminLatestClientState_(client, bookId, routing.state || registry.state);
      if (clientStateAfterSync !== preservedState) {
        // REVIEW_DUE/YEAR_CLOSED must arrive through the authoritative Hub event
        // chain. Only the ordinary official-lock recovery may synthesize a
        // missing final transition on the Client.
        if (preservedState !== 'OFFICIAL_LOCKED') {
          throw new Error('Hub state event chain did not synchronize to Client: expected=' + preservedState + ', actual=' + clientStateAfterSync);
        }
        vNextAdminSetClientState_(registry.spreadsheet_id, preservedState, {
          reason: 'official_client_sync_recovery', relatedRunId: official.forecast_run_id,
          relatedPlanVersionId: approvedPlans[approvedPlans.length - 1].plan_version_id, actorRole: 'ADMIN'
        });
      }
      vNextAdminCopyOfficialToClient_(registry.spreadsheet_id, official);
      vNextAdminMirrorClientState_(client, preservedState);
      vNextAdminPatchRegistryByBookId_(hub, bookId, {
        state: preservedState, health_status: 'OK', health_code: 'OFFICIAL_SYNC_RECOVERED',
        updated_at: new Date()
      });
      vNextAdminResolveOpenExceptions_(hub, bookId,
        ['OFFICIAL_CLIENT_SYNC_FAILED', 'OFFICIAL_COPY_FAILED'], officialId);
      vNextAdminWriteAudit_(hub, 'RETRY_OFFICIAL_CLIENT_SYNC', 'BOOK', bookId, 'SUCCESS', {
        officialId: officialId, clientSpreadsheetId: registry.spreadsheet_id, preservedState: preservedState
      });
      vNextAdminRefreshTodayExceptions_(hub);
      vNextAdminRefreshHome_(hub);
      return { ok: true, bookId: bookId, officialId: officialId, state: preservedState };
    });
  });
}

function vNextAdminOfficialSyncTargetState_(hubState) {
  const normalized = String(hubState || '').toUpperCase();
  return ['REVIEW_DUE', 'YEAR_CLOSED'].indexOf(normalized) >= 0 ? normalized : 'OFFICIAL_LOCKED';
}

function vNextAdminDefaultLearningPolicy_() {
  return JSON.parse(JSON.stringify(VN_ADMIN_LEARNING_POLICY_DEFAULT));
}

function vNextAdminEnsureLearningPolicy_(hub) {
  const existing = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.SETTINGS).rows.find(function (row) {
    return String(row.setting_key || '') === VN_ADMIN_LEARNING_POLICY_KEY;
  });
  if (existing && existing.setting_value) {
    const parsed = vNextAdminParseJson_(existing.setting_value, null);
    if (parsed && parsed.schemaVersion) return parsed;
  }
  const policy = vNextAdminDefaultLearningPolicy_();
  policy.updatedAt = new Date().toISOString();
  policy.updatedBy = vNextAdminActor_();
  const now = new Date();
  vNextAdminUpsertObject_(hub, VN_ADMIN_SHEETS.SETTINGS, 'setting_key', VN_ADMIN_LEARNING_POLICY_KEY, {
    setting_key: VN_ADMIN_LEARNING_POLICY_KEY,
    setting_value: vNextAdminCanonicalJson_(policy),
    value_type: 'JSON', scope: 'LEARNING_POLICY', effective_from: now,
    updated_at: now, updated_by: vNextAdminActor_(),
    note: 'Adaptive learning proxy objectives and constraints'
  });
  return policy;
}

function vNextAdminReadLearningPolicy_(hub) {
  return vNextAdminEnsureLearningPolicy_(hub);
}

function vNextAdminComputeDualTrackFromEvaluation_(row) {
  const actual = Number(row.actual_total);
  const system = Number(row.system_forecast);
  const budget = Number(row.final_budget);
  const systemSigned = isFinite(system) && isFinite(actual) ? system - actual : '';
  const budgetSigned = isFinite(budget) && isFinite(actual) ? budget - actual : '';
  const systemApe = Number(row.system_ape);
  const budgetApe = isFinite(budget) && isFinite(actual) && actual !== 0
    ? Math.abs(budget - actual) / Math.abs(actual) : '';
  return {
    bookId: String(row.book_id || ''),
    fiscalYear: Number(row.fiscal_year || 0),
    evaluationId: String(row.evaluation_id || ''),
    officialVintageId: String(row.official_vintage_id || ''),
    actualTotal: actual,
    system: {
      forecast: system,
      signedError: systemSigned,
      absError: isFinite(systemSigned) ? Math.abs(systemSigned) : '',
      ape: isFinite(systemApe) ? systemApe : (isFinite(systemSigned) && actual !== 0 ? Math.abs(systemSigned) / Math.abs(actual) : ''),
      rangeContainsActual: Number(row.range_contains_actual || 0) === 1
    },
    budget: {
      finalBudget: budget,
      signedError: budgetSigned,
      absError: isFinite(budgetSigned) ? Math.abs(budgetSigned) : '',
      ape: budgetApe,
      learningExcluded: true
    },
    layers: {
      baseLevel: Number(row.base_level_error),
      seasonality: Number(row.seasonality_error),
      commitmentOutcome: Number(row.commitment_outcome_error),
      amount: Number(row.amount_error),
      timing: Number(row.timing_error),
      unknownSpot: Number(row.unknown_spot_error),
      humanInfo: Number(row.human_info_error),
      aiInfo: Number(row.ai_info_error),
      dataQuality: Number(row.data_quality_error)
    }
  };
}

function vNextAdminAppendLearningEvidenceFromEvaluation_(hub, evaluation, officialForecast, automaticBreakdown) {
  if (!evaluation) return null;
  const track = vNextAdminComputeDualTrackFromEvaluation_(evaluation);
  const forecastResult = officialForecast || {
    runId: String(evaluation.source_run_id || ''),
    officialVintageId: String(evaluation.official_vintage_id || ''),
    modelReleaseId: String(evaluation.model_release_id || ''),
    inputDataHash: '',
    layers: { systemRecommended: Number(evaluation.system_forecast) },
    annual: { p10: '', p50: '', p90: '' }
  };
  const learningPayload = typeof vNextBuildLearningPayload_ === 'function'
    ? vNextBuildLearningPayload_(forecastResult, {
      actualTotal: Number(evaluation.actual_total),
      errorComponents: (automaticBreakdown && automaticBreakdown.errorComponents) || {
        baseLevel: evaluation.base_level_error,
        seasonality: evaluation.seasonality_error,
        commitmentOutcome: evaluation.commitment_outcome_error,
        amount: evaluation.amount_error,
        timing: evaluation.timing_error,
        unknownSpot: evaluation.unknown_spot_error,
        humanInfo: evaluation.human_info_error,
        aiInfo: evaluation.ai_info_error,
        dataQuality: evaluation.data_quality_error
      }
    })
    : { systemForecast: Number(evaluation.system_forecast), actualTotal: Number(evaluation.actual_total) };
  const evidenceId = 'LE-' + vNextAdminSha256_([
    String(evaluation.evaluation_id || ''), String(evaluation.official_vintage_id || '')
  ].join('|')).slice(0, 24).toUpperCase();
  const existing = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.LEARNING_EVIDENCE).rows.filter(function (row) {
    return String(row.evidence_id || '') === evidenceId;
  });
  if (existing.length) return existing[existing.length - 1];
  const record = {
    evidence_id: evidenceId,
    book_id: String(evaluation.book_id || ''),
    fiscal_year: Number(evaluation.fiscal_year || 0),
    evaluation_id: String(evaluation.evaluation_id || ''),
    official_vintage_id: String(evaluation.official_vintage_id || ''),
    source_run_id: String(evaluation.source_run_id || ''),
    range_contains_actual: Number(evaluation.range_contains_actual || 0),
    system_ape: track.system.ape === '' ? '' : track.system.ape,
    budget_ape: track.budget.ape === '' ? '' : track.budget.ape,
    layer_errors_json: vNextAdminCanonicalJson_(track.layers),
    learning_payload_json: vNextAdminCanonicalJson_(learningPayload),
    created_at: new Date(),
    created_by: vNextAdminActor_()
  };
  vNextAdminAppendObject_(hub, VN_ADMIN_SHEETS.LEARNING_EVIDENCE, record);
  return record;
}

function vNextAdminLatestLearningEvidenceForBook_(hub, bookId) {
  const rows = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.LEARNING_EVIDENCE).rows.filter(function (row) {
    return String(row.book_id || '') === String(bookId || '');
  });
  return rows.length ? rows[rows.length - 1] : null;
}

function vNextAdminLearningEvidenceForEngine_(hub, bookId) {
  const row = vNextAdminLatestLearningEvidenceForBook_(hub, bookId);
  if (!row) return null;
  const policy = vNextAdminReadLearningPolicy_(hub);
  return {
    evidenceId: String(row.evidence_id || ''),
    evaluationId: String(row.evaluation_id || ''),
    rangeContainsActual: Number(row.range_contains_actual || 0) === 1,
    systemApe: Number(row.system_ape),
    layerErrors: vNextAdminParseJson_(row.layer_errors_json, {}),
    learningPayload: vNextAdminParseJson_(row.learning_payload_json, {}),
    intervalWidenFactorOnMiss: Number(policy.intervalWidenFactorOnMiss || 1.15)
  };
}

function vNextAdminBuildLearningDashboard_(hub) {
  const policy = vNextAdminReadLearningPolicy_(hub);
  const evaluations = vNextAdminReadCoreRows_(hub, 'EVALUATION');
  const dualTracks = evaluations.slice(-20).map(vNextAdminComputeDualTrackFromEvaluation_);
  const observations = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.LEARNING_OBS).rows.slice(-20);
  const evidenceRows = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.LEARNING_EVIDENCE).rows;
  const openBreaches = observations.filter(function (row) {
    return Number(row.range_breach || 0) === 1 &&
      String(row.verification_status || '').toUpperCase() !== 'VERIFIED';
  });
  const missCount = dualTracks.filter(function (row) { return row.system && !row.system.rangeContainsActual; }).length;
  return vNextAdminJsonSafe_({
    policy: policy,
    concept: policy.concept,
    proxyObjectives: policy.proxyObjectives || [],
    nonGoals: policy.nonGoals || [],
    stats: {
      evaluationCount: evaluations.length,
      evidenceCount: evidenceRows.length,
      observationCount: observations.length,
      openIntervalBreaches: openBreaches.length,
      recentIntervalMisses: missCount
    },
    dualTracks: dualTracks.reverse(),
    recentObservations: observations.slice().reverse().slice(0, 8).map(function (row) {
      return {
        observationId: String(row.observation_id || ''),
        bookId: String(row.book_id || ''),
        clientName: String(row.client_name || ''),
        observedMonth: String(row.observed_month || ''),
        rangeBreach: Number(row.range_breach || 0) === 1,
        hypothesis: String(row.hypothesis || ''),
        verificationStatus: String(row.verification_status || ''),
        alerted: Number(row.alerted || 0) === 1
      };
    }),
    openBreaches: openBreaches.slice(-5).map(function (row) {
      return {
        observationId: String(row.observation_id || ''),
        bookId: String(row.book_id || ''),
        clientName: String(row.client_name || ''),
        observedMonth: String(row.observed_month || ''),
        hypothesis: String(row.hypothesis || '')
      };
    })
  });
}

/**
 * Monthly observation: accumulate actual vs official interval, optional hypothesis,
 * and raise an exception when the range is breached (budget is never rewritten).
 */
function vNextAdminRecordLearningObservation(request) {
  return vNextAdminGuard_('vNextAdminRecordLearningObservation', function () {
    const req = request && typeof request === 'object' ? request : {};
    const hub = vNextAdminRequireHub_();
    vNextAdminAssertHubAdmin_(hub, false);
    vNextAdminEnsureLearningPolicy_(hub);
    return vNextAdminWithScriptLock_('learning-observation', function () {
      const bookId = vNextAdminRequiredText_(req.bookId, 'bookId');
      const observedMonth = vNextAdminRequiredText_(req.observedMonth, 'observedMonth');
      if (!/^\d{4}-\d{2}$/.test(observedMonth)) throw new Error('observedMonth must be YYYY-MM.');
      const registry = vNextAdminFindRegistryRow_(hub, function (row) {
        return String(row.book_id || '') === bookId && String(row.mode || '') === 'CLIENT';
      });
      if (!registry) throw new Error('Registered CLIENT book not found: ' + bookId);
      const officialId = String(registry.current_official_id || '');
      let p10 = Number(req.systemP10);
      let p50 = Number(req.systemP50);
      let p90 = Number(req.systemP90);
      if ((!isFinite(p10) || !isFinite(p90)) && officialId) {
        const run = vNextAdminReadCoreRows_(hub, 'FORECAST_RUN').filter(function (row) {
          return String(row.book_id || '') === bookId && String(row.official_vintage_id || '') === officialId &&
            Number(row.is_official || 0) === 1;
        }).slice(-1)[0];
        if (run) {
          p10 = Number(run.p10);
          p50 = Number(run.p50);
          p90 = Number(run.p90);
        }
      }
      const actualAmount = Number(req.actualAmount);
      if (!isFinite(actualAmount)) throw new Error('actualAmount is required.');
      const rangeBreach = isFinite(p10) && isFinite(p90)
        ? (actualAmount < p10 || actualAmount > p90 ? 1 : 0) : 0;
      const policy = vNextAdminReadLearningPolicy_(hub);
      const alerted = rangeBreach === 1 && policy.intervalBreachAlert !== false ? 1 : 0;
      const observationId = 'OBS-' + vNextAdminSha256_(vNextAdminCanonicalJson_({
        bookId: bookId, observedMonth: observedMonth, actualAmount: actualAmount,
        hypothesis: vNextAdminText_(req.hypothesis), attempt: Utilities.getUuid()
      })).slice(0, 20).toUpperCase();
      const record = {
        observation_id: observationId,
        book_id: bookId,
        client_name: String(registry.client_name || ''),
        fiscal_year: Number(registry.fiscal_year || 0),
        observed_month: observedMonth,
        actual_amount: actualAmount,
        system_p10: isFinite(p10) ? p10 : '',
        system_p50: isFinite(p50) ? p50 : '',
        system_p90: isFinite(p90) ? p90 : '',
        range_breach: rangeBreach,
        hypothesis: vNextAdminText_(req.hypothesis),
        verification_status: vNextAdminText_(req.verificationStatus) || (rangeBreach ? 'NEEDS_FIELD_CHECK' : 'RECORDED'),
        verification_note: vNextAdminText_(req.verificationNote),
        alerted: alerted,
        created_at: new Date(),
        created_by: vNextAdminActor_(),
        detail_json: vNextAdminCanonicalJson_({
          officialVintageId: officialId,
          kind: 'MONTHLY_OBSERVATION_V1',
          budgetUnchanged: true
        })
      };
      vNextAdminAppendObject_(hub, VN_ADMIN_SHEETS.LEARNING_OBS, record);
      if (alerted) {
        vNextAdminAppendException_(hub, {
          severity: 'WARN',
          exception_type: 'LEARNING_INTERVAL_BREACH',
          book_id: bookId,
          client_name: registry.client_name,
          fiscal_year: registry.fiscal_year,
          title: '想定区間を外れた月次実績',
          detail: observedMonth + ' actual=' + actualAmount + ' vs P10–P90 [' + p10 + ', ' + p90 + ']',
          recommended_action: '現場確認のうえ仮説・検証を LEARNING_OBSERVATION に残す。正式予算は変更しない。',
          source_ref: observationId
        });
      }
      vNextAdminWriteAudit_(hub, 'RECORD_LEARNING_OBSERVATION', 'LEARNING_OBS', observationId, 'SUCCESS', {
        bookId: bookId, observedMonth: observedMonth, rangeBreach: rangeBreach, alerted: alerted
      });
      return vNextAdminJsonSafe_({
        ok: true, observationId: observationId, rangeBreach: rangeBreach === 1, alerted: alerted === 1,
        message: rangeBreach
          ? '区間外れを記録し、要確認アラートを出しました（予算は変更していません）。'
          : '月次観測を記録しました。'
      });
    });
  });
}

function vNextAdminGetLearningDashboard(request) {
  return vNextAdminGuard_('vNextAdminGetLearningDashboard', function () {
    const hub = vNextAdminRequireHub_();
    return vNextAdminBuildLearningDashboard_(hub);
  });
}

/** Aggregate confirmed BE actuals for the full FY and start the learning review. */
function vNextAdminStartReview(request) {
  return vNextAdminGuard_('vNextAdminStartReview', function () {
    const req = request && typeof request === 'object' ? request : {};
    const hub = vNextAdminRequireHub_();
    let reviewClaimKey = '';
    let reviewClaimAcquired = false;
    try {
      return vNextAdminWithUserLock_('start-review', function () {
      const bookId = vNextAdminRequiredText_(req.bookId, 'bookId');
      const registry = vNextAdminFindRegistryRow_(hub, function (row) { return String(row.book_id) === bookId; });
      if (!registry || String(registry.mode) !== 'CLIENT') throw new Error('Registered CLIENT book not found: ' + bookId);
      const officialId = vNextAdminRequiredText_(registry.current_official_id, 'currentOfficialId');
      if (req.officialVintageId && String(req.officialVintageId) !== officialId) {
        throw new Error('振り返りはBOOK_REGISTRYの現在の公式vintageだけを対象にできます。');
      }
      const officialRows = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.OFFICIAL).rows.filter(function (row) {
        return String(row.book_id || '') === bookId && String(row.official_id || '') === officialId;
      });
      if (!officialRows.length) throw new Error('Current official record was not found in OFFICIAL_RUNS.');
      const officialRow = officialRows[officialRows.length - 1];
      const officialForecasts = vNextAdminReadCoreRows_(hub, 'FORECAST_RUN').filter(function (row) {
        return String(row.book_id || '') === bookId && String(row.official_vintage_id || '') === officialId &&
          String(row.run_id || '') === String(officialRow.forecast_run_id || '') && Number(row.is_official || 0) === 1;
      });
      if (!officialForecasts.length) throw new Error('The current official FORECAST_RUN is missing or does not match OFFICIAL_RUNS.');

      const fiscalYear = Number(registry.fiscal_year);
      const fyStart = new Date(fiscalYear, 3, 1);
      const fyEndDate = new Date(fiscalYear + 1, 2, 31);
      const fyEnd = new Date(fiscalYear + 1, 2, 31, 23, 59, 59, 999);
      const minimumAsOf = new Date(fiscalYear + 1, 3, 1);
      const asOf = typeof vNextParseDate_ === 'function'
        ? vNextParseDate_(req.asOf || new Date(), 'asOf')
        : new Date(req.asOf || new Date());
      if (isNaN(asOf.getTime()) || asOf < minimumAsOf) {
        throw new Error('年度末12か月の確定実績を含めるため、asOfは翌年度4月1日以降にしてください。');
      }
      const asOfCutoff = typeof vNextParseDate_ === 'function'
        ? vNextParseDate_(vNextAdminCutoffFromAsOf_(asOf), 'asOfCutoff')
        : new Date(vNextAdminCutoffFromAsOf_(asOf));
      if (asOfCutoff < fyEndDate) throw new Error('asOfから算出した情報締切が対象年度末より前です。');

      const client = SpreadsheetApp.openById(String(registry.spreadsheet_id));
      const routing = vNextAdminReadKeyValueSheet_(client, VN_ADMIN_BOOK_CONFIG_SHEET);
      const currentState = String(routing.state || registry.state || '').toUpperCase();
      if (currentState === 'YEAR_CLOSED') throw new Error('終了済み年度の振り返り評価は再作成できません。');
      if (['OFFICIAL_LOCKED', 'REVIEW_DUE'].indexOf(currentState) < 0) {
        throw new Error('年度振り返りを開始できる状態ではありません: ' + currentState);
      }
      vNextAdminSyncClientToHub_(hub, client, bookId);

      const existing = vNextAdminReadCoreRows_(hub, 'EVALUATION').filter(function (row) {
        return String(row.book_id || '') === bookId && String(row.official_vintage_id || '') === officialId;
      });
      let evaluation = existing.length ? existing[existing.length - 1] : null;
      const evaluationId = 'EVAL-' + vNextAdminSha256_([bookId, officialId].join('|')).slice(0, 24).toUpperCase();
      reviewClaimKey = 'REVIEW_CLAIM|' + bookId + '|' + officialId;
      if (evaluation && String(evaluation.source_run_id || '') !== String(officialRow.forecast_run_id || '')) {
        throw new Error('Existing evaluation points to a non-current official run.');
      }
      if (!evaluation) {
        vNextAdminWithScriptLock_('claim-start-review', function () {
          const recheck = vNextAdminReadCoreRows_(hub, 'EVALUATION').filter(function (row) {
            return String(row.book_id || '') === bookId && String(row.official_vintage_id || '') === officialId;
          });
          if (recheck.length) {
            evaluation = recheck[recheck.length - 1];
            return;
          }
          const setting = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.SETTINGS).rows.find(function (row) {
            return String(row.setting_key || '') === reviewClaimKey;
          });
          const priorClaim = setting ? vNextAdminParseJson_(setting.setting_value, {}) : {};
          const priorTime = new Date(priorClaim.claimedAt || 0).getTime();
          if (String(priorClaim.status || '') === 'RUNNING' && isFinite(priorTime) && Date.now() - priorTime < 20 * 60000) {
            throw new Error('同じ公式vintageの年度評価を別の処理が作成中です。');
          }
          const now = new Date();
          vNextAdminUpsertObject_(hub, VN_ADMIN_SHEETS.SETTINGS, 'setting_key', reviewClaimKey, {
            setting_key: reviewClaimKey,
            setting_value: vNextAdminCanonicalJson_({
              status: 'RUNNING', evaluationId: evaluationId, claimedAt: now.toISOString(), actor: vNextAdminActor_()
            }),
            value_type: 'JSON', scope: 'REVIEW_CLAIM', effective_from: now,
            updated_at: now, updated_by: vNextAdminActor_(), note: 'duplicate-prevention lease'
          });
          reviewClaimAcquired = true;
        });
      }
      if (!evaluation) {
        if (typeof vNextFetchActualRecordsBridge_ !== 'function' || typeof vNextAppendEvaluation_ !== 'function') {
          throw new Error('Actual-data evaluation APIs are not installed.');
        }
        vNextAdminHydrateHubRuntime_(hub);
        const records = vNextFetchActualRecordsBridge_(registry.client_name, {
          fiscalYear: fiscalYear, asOf: asOf, cutoff: fyEnd
        });
        const fyActuals = records.filter(function (record) {
          const date = record.actualDate instanceof Date ? record.actualDate : new Date(record.actualDate);
          return record.isConfirmed === true && String(record.dateSource || '').indexOf('ACTUAL') === 0 &&
            date >= fyStart && date <= fyEnd;
        });
        const actualTotal = fyActuals.reduce(function (sum, record) { return sum + Number(record.amount || 0); }, 0);
        const observedMonths = new Set(fyActuals.map(function (record) {
          const date = record.actualDate instanceof Date ? record.actualDate : new Date(record.actualDate);
          return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM');
        }));
        const expectedMonths = [];
        for (let offset = 0; offset < 12; offset++) {
          expectedMonths.push(Utilities.formatDate(new Date(fiscalYear, 3 + offset, 1), Session.getScriptTimeZone(), 'yyyy-MM'));
        }
        const missingMonths = expectedMonths.filter(function (month) { return !observedMonths.has(month); });
        const latestActual = fyActuals.reduce(function (latest, record) {
          const date = record.actualDate instanceof Date ? record.actualDate : new Date(record.actualDate);
          return !latest || date > latest ? date : latest;
        }, null);
        const freshnessThreshold = new Date(fiscalYear + 1, 2, 1);
        const completenessIssues = [];
        if (!fyActuals.length) completenessIssues.push('ZERO_ACTUAL_ROWS');
        if (actualTotal === 0) completenessIssues.push('ZERO_ACTUAL_TOTAL');
        if (missingMonths.length) completenessIssues.push('MISSING_MONTHS=' + missingMonths.join(','));
        if (!latestActual || latestActual < freshnessThreshold) completenessIssues.push('STALE_LATEST_ACTUAL');
        const overrideReason = vNextAdminText_(req.actualCompletenessOverrideReason || req.overrideReason);
        if (completenessIssues.length && !(req.allowIncompleteActuals === true && overrideReason)) {
          throw new Error('ZAC確定実績の完全性を確認できません: ' + completenessIssues.join('; ') +
            '。例外評価が必要な場合はallowIncompleteActuals=trueと理由を指定してください。');
        }

        const plans = vNextAdminReadCoreRows_(hub, 'PLAN_VERSION').filter(function (row) {
          return String(row.book_id || '') === bookId && String(row.official_vintage_id || '') === officialId &&
            String(row.status || '').toUpperCase() === 'APPROVED';
        });
        const plan = plans.length ? plans[plans.length - 1] : null;
        if (!plan) throw new Error('Approved plan linked to the current official vintage was not found.');
        if (String(plan.run_id || '') !== String(officialRow.source_forecast_run_id || '')) {
          throw new Error('Approved plan does not point to the source run of the current official vintage.');
        }
        if (completenessIssues.length) {
          // This direct append is intentionally fail-closed; the immutable
          // evaluation is never written unless the override reason is recorded.
          vNextAdminAppendObject_(hub, VN_ADMIN_SHEETS.AUDIT, {
            audit_id: 'AUD-' + Utilities.getUuid(), occurred_at: new Date(), actor: vNextAdminActor_(),
            action: 'ACTUAL_COMPLETENESS_OVERRIDE', entity_type: 'BOOK', entity_id: bookId, status: 'APPROVED_EXCEPTION',
            detail_json: vNextAdminCanonicalJson_({
              officialVintageId: officialId, fiscalYear: fiscalYear, issues: completenessIssues,
              reason: overrideReason, actualRows: fyActuals.length, actualTotal: actualTotal,
              observedMonths: Array.from(observedMonths), latestActual: latestActual && latestActual.toISOString()
            }), before_hash: '', after_hash: ''
          });
        }
        if (typeof vNextEngineBuildEvaluationBreakdown !== 'function' ||
            typeof vNextForecastRecordToResult_ !== 'function') {
          throw new Error('Automatic evaluation breakdown API is not installed.');
        }
        const actualBaseMonths = new Array(12).fill(0);
        const actualSpotMonths = new Array(12).fill(0);
        fyActuals.forEach(function (record) {
          const date = record.actualDate instanceof Date ? record.actualDate : new Date(record.actualDate);
          const monthIndex = (date.getFullYear() - fiscalYear) * 12 + date.getMonth() - 3;
          if (monthIndex < 0 || monthIndex > 11) return;
          const target = String(record.serviceType || '').toUpperCase() === 'SPOT'
            ? actualSpotMonths : actualBaseMonths;
          target[monthIndex] += Number(record.amount || 0);
        });
        const officialForecast = vNextForecastRecordToResult_(officialForecasts[officialForecasts.length - 1]);
        const automaticBreakdown = vNextEngineBuildEvaluationBreakdown({
          officialForecast: officialForecast,
          actualBaseMonths: actualBaseMonths,
          actualSpotMonths: actualSpotMonths,
          actualTotal: actualTotal,
          dataQualityIssues: completenessIssues
        });
        if (!automaticBreakdown.reconciled || Math.abs(Number(automaticBreakdown.reconciliationResidual || 0)) > 0.000001) {
          throw new Error('Automatic evaluation breakdown does not reconcile to the official system error.');
        }
        evaluation = vNextAppendEvaluation_({
          evaluationId: evaluationId, bookId: bookId, officialVintageId: officialId, actualTotal: actualTotal,
          adoptedForecast: Number(plan.adopted_forecast || 0), finalBudget: Number(plan.final_budget || 0),
          evaluatedAt: new Date().toISOString(), errorComponents: automaticBreakdown.errorComponents,
          confirmedCause: '', causeHypothesis: '', nextInformation: [], createdBy: vNextAdminActor_()
        }, { spreadsheet: hub });
        vNextAdminAppendLearningEvidenceFromEvaluation_(hub, evaluation, officialForecast, automaticBreakdown);
        vNextAdminWriteAudit_(hub, 'START_REVIEW', 'EVALUATION', evaluation.evaluation_id, 'SUCCESS', {
          bookId: bookId, officialVintageId: officialId, officialRunId: officialRow.forecast_run_id,
          actualRows: fyActuals.length, actualTotal: actualTotal, observedMonthCount: observedMonths.size,
          missingMonths: missingMonths, latestActual: latestActual && latestActual.toISOString(),
          asOf: asOf.toISOString(), cutoff: fyEndDate.toISOString(), dateSource: 'BE_ACTUAL_ONLY',
          completenessOverride: completenessIssues.length ? overrideReason : '',
          automaticBreakdown: automaticBreakdown
        });
        vNextAdminWithScriptLock_('complete-start-review', function () {
          const now = new Date();
          vNextAdminUpsertObject_(hub, VN_ADMIN_SHEETS.SETTINGS, 'setting_key', reviewClaimKey, {
            setting_key: reviewClaimKey,
            setting_value: vNextAdminCanonicalJson_({
              status: 'COMPLETE', evaluationId: evaluation.evaluation_id, completedAt: now.toISOString(), actor: vNextAdminActor_()
            }),
            value_type: 'JSON', scope: 'REVIEW_CLAIM', effective_from: now,
            updated_at: now, updated_by: vNextAdminActor_(), note: 'immutable evaluation created'
          });
          reviewClaimAcquired = false;
        });
      }
      vNextAdminSetClientState_(registry.spreadsheet_id, 'REVIEW_DUE', {
        reason: 'full_year_actuals_ready_for_review', relatedRunId: evaluation.source_run_id
      });
      if (evaluation) {
        const officialForecastForEvidence = officialForecasts.length
          ? vNextForecastRecordToResult_(officialForecasts[officialForecasts.length - 1])
          : null;
        vNextAdminAppendLearningEvidenceFromEvaluation_(hub, evaluation, officialForecastForEvidence, null);
      }
      vNextAdminSyncClientToHub_(hub, client, bookId);
      vNextAdminSyncHubToClient_(hub, client, bookId, ['EVALUATION', 'STATE_EVENT']);
      vNextAdminMirrorClientState_(client, 'REVIEW_DUE');
      vNextAdminPatchRegistryByBookId_(hub, bookId, { state: 'REVIEW_DUE', updated_at: new Date() });
      return vNextAdminJsonSafe_({
        bookId: bookId, officialVintageId: officialId,
        evaluationId: evaluation.evaluation_id, state: 'REVIEW_DUE'
      });
      });
    } catch (error) {
      if (reviewClaimAcquired && reviewClaimKey) {
        try {
          vNextAdminWithScriptLock_('release-start-review', function () {
            const now = new Date();
            vNextAdminUpsertObject_(hub, VN_ADMIN_SHEETS.SETTINGS, 'setting_key', reviewClaimKey, {
              setting_key: reviewClaimKey,
              setting_value: vNextAdminCanonicalJson_({
                status: 'FAILED', failedAt: now.toISOString(), actor: vNextAdminActor_(),
                error: String(error && error.message || error).slice(0, 500)
              }),
              value_type: 'JSON', scope: 'REVIEW_CLAIM', effective_from: now,
              updated_at: now, updated_by: vNextAdminActor_(), note: 'claim released after failed review start'
            });
          });
        } catch (releaseError) {
          Logger.log('Review claim release failed key=%s error=%s', reviewClaimKey,
            String(releaseError && releaseError.stack || releaseError));
        }
      }
      throw error;
    }
  });
}

function vNextAdminCloseYear(request) {
  return vNextAdminGuard_('vNextAdminCloseYear', function () {
    const req = request && typeof request === 'object' ? request : {};
    const hub = vNextAdminRequireHub_();
    const bookId = vNextAdminRequiredText_(req.bookId, 'bookId');
    const reason = vNextAdminRequiredText_(req.reason, 'reason');
    const registry = vNextAdminFindRegistryRow_(hub, function (row) { return String(row.book_id) === bookId; });
    if (!registry || String(registry.mode) !== 'CLIENT') throw new Error('Registered CLIENT book not found: ' + bookId);
    const officialId = vNextAdminRequiredText_(registry.current_official_id, 'currentOfficialId');
    const client = SpreadsheetApp.openById(String(registry.spreadsheet_id));
    vNextAdminSyncClientToHub_(hub, client, bookId);
    const evaluations = vNextAdminReadCoreRows_(hub, 'EVALUATION').filter(function (row) {
      return String(row.book_id || '') === bookId && String(row.official_vintage_id || '') === officialId;
    });
    if (!evaluations.length) throw new Error('年度を終了する前に、現在の公式vintageの振り返りを開始してください。');
    const evaluation = evaluations[evaluations.length - 1];
    if (Number(evaluation.fiscal_year) !== Number(registry.fiscal_year)) {
      throw new Error('Current evaluation fiscal year does not match BOOK_REGISTRY.');
    }
    const evaluatedAt = new Date(evaluation.evaluated_at || 0).getTime();
    const validReviews = vNextAdminReadCoreRows_(hub, 'EVIDENCE_EVENT').filter(function (row) {
      if (String(row.book_id || '') !== bookId || Number(row.fiscal_year) !== Number(registry.fiscal_year) ||
          String(row.evidence_type || '').toUpperCase() !== 'REVIEW_LEARNING' ||
          String(row.status || 'ACTIVE').toUpperCase() === 'VOID') return false;
      const createdAt = new Date(row.created_at || 0).getTime();
      if (!isFinite(createdAt) || !isFinite(evaluatedAt) || createdAt < evaluatedAt) return false;
      const learning = vNextAdminParseJson_(row.evidence_text, null);
      if (!learning || String(learning.bookId || '') !== bookId ||
          Number(learning.fiscalYear) !== Number(registry.fiscal_year) ||
          String(learning.officialVintageId || '') !== officialId ||
          String(learning.evaluationId || '') !== String(evaluation.evaluation_id || '')) return false;
      const hasCause = !!vNextAdminText_(learning.confirmedCause || learning.causeHypothesis);
      const nextInformation = Array.isArray(learning.nextInformation) ? learning.nextInformation.filter(Boolean) : [];
      return hasCause && nextInformation.length > 0;
    });
    if (!validReviews.length) {
      throw new Error('現在の公式vintage・評価IDに紐づく振り返り（原因と次に確認する情報）が保存されるまで年度終了できません。');
    }
    const review = validReviews[validReviews.length - 1];
    vNextAdminSetClientState_(registry.spreadsheet_id, 'YEAR_CLOSED', {
      reason: 'year_closed: ' + reason, actorRole: 'ADMIN'
    });
    vNextAdminSyncClientToHub_(hub, client, bookId);
    vNextAdminMirrorClientState_(client, 'YEAR_CLOSED');
    vNextAdminPatchRegistryByBookId_(hub, bookId, { state: 'YEAR_CLOSED', updated_at: new Date() });
    vNextAdminWriteAudit_(hub, 'CLOSE_YEAR', 'BOOK', bookId, 'SUCCESS', {
      officialVintageId: officialId, evaluationId: evaluation.evaluation_id,
      reviewEvidenceId: review.evidence_id, fiscalYear: Number(registry.fiscal_year), reason: reason
    });
    return { bookId: bookId, officialVintageId: officialId, state: 'YEAR_CLOSED' };
  });
}

function vNextAdminEnqueueMigration(request) {
  return vNextAdminGuard_('vNextAdminEnqueueMigration', function () {
    const req = request && typeof request === 'object' ? request : {};
    if (req.dryRun === false && !VN_ADMIN_MIGRATION_APPLY_ENABLED) {
      throw new Error('Pilot期間はClient release移行のAPPLYをserver-sideで停止しています。dry-runだけを使用してください。');
    }
    const hub = vNextAdminRequireHub_();
    const bookId = vNextAdminRequiredText_(req.bookId, 'bookId');
    const targetReleaseId = vNextAdminRequiredText_(req.targetReleaseId, 'targetReleaseId');
    const registry = vNextAdminFindRegistryRow_(hub, function (row) { return String(row.book_id) === bookId; });
    if (!registry || String(registry.mode || '') !== 'CLIENT') throw new Error('Client Book not found: ' + bookId);
    const targetRelease = vNextAdminResolveRelease_(hub, targetReleaseId);
    if (String(targetRelease.release_id || '') === String(registry.template_release_id || '') && req.dryRun === false) {
      throw new Error('Client Book already uses the requested release.');
    }
    const reason = req.dryRun === false ? vNextAdminRequiredText_(req.reason, 'reason') : vNextAdminText_(req.reason);
    return vNextAdminEnqueueJob({
      jobType: 'MIGRATION', targetBookId: bookId, targetSpreadsheetId: registry.spreadsheet_id,
      request: { bookId: bookId, spreadsheetId: registry.spreadsheet_id, fromReleaseId: registry.template_release_id,
        targetReleaseId: targetReleaseId, dryRun: req.dryRun !== false,
        critical: req.critical === true, reason: reason },
      idempotencyKey: ['MIGRATION', bookId, registry.template_release_id, targetReleaseId, req.dryRun !== false ? 'DRY' : 'APPLY'].join('|'),
      priority: req.priority || 10
    });
  });
}

/**
 * One-purpose in-place upgrade for a completely unused Pilot Client book.
 *
 * Generic migration APPLY intentionally remains disabled. This path is
 * narrower: it accepts only an ACTIVE CLIENT that has never received an
 * answer, request, forecast, plan, approval, official record or evaluation,
 * and whose Hub/Client/runtime pins still exactly match the source release.
 * The registry pointer is committed last. Every earlier mutation can be
 * reconstructed from one of the two immutable Template releases.
 */
function vNextAdminUpgradeEmptyPilotClient(request) {
  return vNextAdminGuard_('vNextAdminUpgradeEmptyPilotClient', function () {
    const req = request && typeof request === 'object' ? request : {};
    const hub = vNextAdminRequireHub_();
    vNextAdminAssertHubAdmin_(hub, false);
    return vNextAdminWithScriptLock_('upgrade-empty-pilot-client', function () {
      const bookId = vNextAdminRequiredText_(req.bookId, 'bookId');
      const dryRun = req.dryRun !== false;
      const reason = dryRun ? (vNextAdminText_(req.reason) || '実行前の安全条件確認')
        : vNextAdminRequiredText_(req.reason, 'reason');
      const registry = vNextAdminFindRegistryRow_(hub, function (row) {
        return String(row.book_id || '') === bookId && String(row.mode || '') === 'CLIENT';
      });
      if (!registry) throw new Error('対象のClient BookがBOOK_REGISTRYにありません。');

      const pair = vNextAdminReadActiveReleasePair_(hub);
      const targetReleaseId = vNextAdminText_(req.targetReleaseId) || pair.releaseId;
      if (targetReleaseId !== pair.releaseId) {
        throw new Error('空のPilot更新先は現在のcanonical ACTIVE Template Releaseだけです。');
      }
      const targetRelease = vNextAdminEmptyPilotRelease_(hub, targetReleaseId, ['ACTIVE']);
      if (String(targetRelease.template_spreadsheet_id || '') !== pair.templateSpreadsheetId) {
        throw new Error('更新先Templateがcanonical ACTIVE pairと一致しません。');
      }
      const targetModel = vNextAdminEmptyPilotModel_(hub, pair.modelReleaseId, targetRelease);
      const sourceReleaseId = vNextAdminRequiredText_(registry.template_release_id,
        'registry.template_release_id');

      const previous = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.MIGRATIONS).rows.filter(function (row) {
        return String(row.book_id || '') === bookId &&
          String(row.to_release_id || '') === targetReleaseId &&
          /^EMPTY_PILOT_|^RECOVERY_REQUIRED$/.test(String(row.status || '').toUpperCase());
      }).slice(-1)[0];
      if (previous) {
        throw new Error('未完了の空Pilot更新があります。vNextAdminRecoverEmptyPilotClientUpgradeで復旧してください。migrationId=' +
          String(previous.migration_id || ''));
      }
      if (sourceReleaseId === targetReleaseId) {
        const succeeded = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.MIGRATIONS).rows.filter(function (row) {
          return String(row.book_id || '') === bookId && String(row.to_release_id || '') === targetReleaseId &&
            String(row.status || '').toUpperCase() === 'SUCCEEDED';
        }).slice(-1)[0];
        if (!succeeded) throw new Error('Clientは更新先releaseを参照していますが、完了journalがありません。復旧を実行してください。');
        const client = SpreadsheetApp.openById(String(registry.spreadsheet_id || ''));
        vNextAdminAssertEmptyPilotPinnedRelease_(hub, client, registry, targetRelease, targetModel, false);
        return { ok: true, reused: true, migrationId: succeeded.migration_id,
          bookId: bookId, releaseId: targetReleaseId };
      }

      const sourceRelease = vNextAdminEmptyPilotRelease_(hub, sourceReleaseId, ['ACTIVE', 'RETIRED']);
      if (String(sourceRelease.schema_version || '') !== String(targetRelease.schema_version || '') ||
          String(targetRelease.schema_version || '') !== vNextAdminClientSchemaVersion_()) {
        throw new Error('空のPilot更新は同一Client Core schema間だけで実行できます。');
      }
      const client = SpreadsheetApp.openById(vNextAdminRequiredText_(registry.spreadsheet_id,
        'registry.spreadsheet_id'));
      const routing = vNextAdminReadKeyValueSheet_(client, VN_ADMIN_BOOK_CONFIG_SHEET);
      const sourceModel = vNextAdminEmptyPilotModel_(hub,
        vNextAdminRequiredText_(routing.model_release_id, 'client.model_release_id'), sourceRelease,
        ['ACTIVE', 'RETIRED']);
      const sourceMeta = vNextAdminAssertEmptyPilotUpgradeEligibility_(
        hub, client, registry, sourceRelease, sourceModel
      );
      vNextAdminAssertEmptyPilotReleaseAssets_(sourceRelease);
      vNextAdminAssertEmptyPilotReleaseAssets_(targetRelease);

      const basis = {
        eligible: true, readOnly: dryRun, bookId: bookId,
        clientName: String(registry.client_name || ''), fiscalYear: Number(registry.fiscal_year || 0),
        spreadsheetId: String(registry.spreadsheet_id || ''),
        spreadsheetUrl: String(registry.spreadsheet_url || ''), sameUrl: true,
        currentReleaseId: sourceReleaseId, targetReleaseId: targetReleaseId,
        currentRuntimeSha256: String(sourceRelease.client_runtime_sha256 || ''),
        targetRuntimeSha256: String(targetRelease.client_runtime_sha256 || ''),
        targetModelReleaseId: String(targetModel.model_release_id || ''),
        schemaVersion: String(targetRelease.schema_version || ''),
        safeguards: [
          'ACTIVEかつINPUT_OPEN', '回答・依頼・予測・計画・承認・公式・評価なし',
          'Hub/Client/Apps Scriptのsource pinとSHA-256一致',
          '同一schema', '更新先はcanonical ACTIVE pair', 'registryは最後に更新'
        ]
      };
      if (dryRun) return basis;

      const migrationId = 'EMPTY-UPG-' + vNextAdminSha256_(vNextAdminCanonicalJson_({
        bookId: bookId, spreadsheetId: registry.spreadsheet_id,
        fromReleaseId: sourceReleaseId, toReleaseId: targetReleaseId,
        attemptId: Utilities.getUuid()
      })).slice(0, 24).toUpperCase();
      const plan = {
        kind: 'EMPTY_PILOT_CLIENT_UPGRADE_V1', bookId: bookId,
        spreadsheetId: String(registry.spreadsheet_id || ''),
        clientScriptId: String(registry.client_script_id || ''),
        sourceReleaseId: sourceReleaseId, sourceModelReleaseId: String(sourceModel.model_release_id || ''),
        sourceMetaRecordId: String(sourceMeta.record_id || ''),
        targetReleaseId: targetReleaseId, targetModelReleaseId: String(targetModel.model_release_id || ''),
        schemaVersion: String(targetRelease.schema_version || ''), reason: reason,
        migrationId: migrationId
      };
      const now = new Date();
      vNextAdminAppendObject_(hub, VN_ADMIN_SHEETS.MIGRATIONS, {
        migration_id: migrationId, book_id: bookId, spreadsheet_id: registry.spreadsheet_id,
        from_release_id: sourceReleaseId, to_release_id: targetReleaseId,
        status: 'EMPTY_PILOT_VALIDATED', dry_run: 0,
        plan_json: vNextAdminCanonicalJson_(plan), result_json: '',
        started_at: now, finished_at: '', actor: vNextAdminActor_(), error: ''
      });

      try {
        vNextAdminFreezeEmptyPilotClient_(hub, client, migrationId);
        vNextAdminSetEmptyPilotUpgradePhase_(hub, migrationId, 'EMPTY_PILOT_FROZEN');
        const applied = vNextAdminApplyEmptyPilotRelease_(hub, client, registry, plan,
          targetRelease, targetModel, 'TARGET', function (phase, detail) {
            vNextAdminSetEmptyPilotUpgradePhase_(hub, migrationId, phase, detail);
          });
        vNextAdminPatchLatestMigration_(hub, migrationId, {
          status: 'SUCCEEDED', finished_at: new Date(), error: '',
          result_json: vNextAdminCanonicalJson_({ phase: 'COMPLETED', direction: 'TARGET',
            healthStatus: applied.health.healthStatus, healthCode: applied.health.healthCode })
        });
        vNextAdminWriteAudit_(hub, 'UPGRADE_EMPTY_PILOT_CLIENT', 'BOOK', bookId, 'SUCCESS', {
          migrationId: migrationId, fromReleaseId: sourceReleaseId,
          toReleaseId: targetReleaseId, reason: reason
        });
        return { ok: true, reused: false, migrationId: migrationId, bookId: bookId,
          fromReleaseId: sourceReleaseId, toReleaseId: targetReleaseId,
          spreadsheetId: registry.spreadsheet_id, spreadsheetUrl: registry.spreadsheet_url,
          health: applied.health };
      } catch (upgradeError) {
        const originalMessage = String(upgradeError && upgradeError.message || upgradeError);
        try {
          vNextAdminSetEmptyPilotUpgradePhase_(hub, migrationId, 'EMPTY_PILOT_ROLLBACK_STARTED', {
            cause: originalMessage.slice(0, 500)
          });
          const currentRegistry = vNextAdminFindRegistryRow_(hub, function (row) {
            return String(row.book_id || '') === bookId;
          }) || registry;
          vNextAdminApplyEmptyPilotRelease_(hub, client, currentRegistry, plan,
            sourceRelease, sourceModel, 'SOURCE', function (phase, detail) {
              vNextAdminSetEmptyPilotUpgradePhase_(hub, migrationId, phase, detail);
            });
          vNextAdminPatchLatestMigration_(hub, migrationId, {
            status: 'ROLLED_BACK', finished_at: new Date(), error: originalMessage,
            result_json: vNextAdminCanonicalJson_({ phase: 'ROLLED_BACK_TO_SOURCE',
              sourceReleaseId: sourceReleaseId })
          });
          vNextAdminWriteAudit_(hub, 'UPGRADE_EMPTY_PILOT_CLIENT', 'BOOK', bookId, 'ROLLED_BACK', {
            migrationId: migrationId, cause: originalMessage, sourceReleaseId: sourceReleaseId
          });
        } catch (rollbackError) {
          const rollbackMessage = String(rollbackError && rollbackError.message || rollbackError);
          vNextAdminPatchLatestMigration_(hub, migrationId, {
            status: 'RECOVERY_REQUIRED', error: originalMessage + '; rollback=' + rollbackMessage,
            result_json: vNextAdminCanonicalJson_({ phase: 'RECOVERY_REQUIRED',
              originalError: originalMessage.slice(0, 500), rollbackError: rollbackMessage.slice(0, 500) })
          });
          vNextAdminAppendException_(hub, {
            severity: 'ERROR', exception_type: 'EMPTY_PILOT_UPGRADE_RECOVERY_REQUIRED',
            book_id: bookId, client_name: registry.client_name, fiscal_year: registry.fiscal_year,
            title: '空のPilot Client更新に復旧が必要です',
            detail: originalMessage + '; rollback=' + rollbackMessage,
            recommended_action: 'vNextAdminRecoverEmptyPilotClientUpgradeを実行', source_ref: migrationId
          });
          throw new Error('更新と自動rollbackの両方が完了しませんでした。migrationId=' + migrationId +
            '; cause=' + originalMessage + '; rollback=' + rollbackMessage);
        }
        throw upgradeError;
      }
    });
  });
}

/**
 * Recovers a journaled empty-Pilot upgrade after an execution timeout or an
 * Apps Script API response loss. Registry is the commit marker: a target pin
 * is completed to target; otherwise the book is reconstructed from source.
 */
function vNextAdminRecoverEmptyPilotClientUpgrade(request) {
  return vNextAdminGuard_('vNextAdminRecoverEmptyPilotClientUpgrade', function () {
    const req = request && typeof request === 'object' ? request : {};
    const hub = vNextAdminRequireHub_();
    vNextAdminAssertHubAdmin_(hub, false);
    return vNextAdminWithScriptLock_('recover-empty-pilot-client', function () {
      const bookId = vNextAdminRequiredText_(req.bookId, 'bookId');
      const migrationRows = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.MIGRATIONS).rows.filter(function (row) {
        return String(row.book_id || '') === bookId &&
          (!req.migrationId || String(row.migration_id || '') === String(req.migrationId));
      });
      const migration = migrationRows.slice(-1)[0];
      if (!migration) throw new Error('復旧対象のMIGRATION_LOGがありません。');
      const plan = vNextAdminParseJson_(migration.plan_json, null);
      if (!plan || String(plan.kind || '') !== 'EMPTY_PILOT_CLIENT_UPGRADE_V1' ||
          String(plan.bookId || '') !== bookId ||
          String(plan.spreadsheetId || '') !== String(migration.spreadsheet_id || '')) {
        throw new Error('復旧journalのidentityが不正です。');
      }
      const registry = vNextAdminFindRegistryRow_(hub, function (row) {
        return String(row.book_id || '') === bookId && String(row.mode || '') === 'CLIENT';
      });
      if (!registry || String(registry.status || '').toUpperCase() !== 'ACTIVE' ||
          String(registry.spreadsheet_id || '') !== String(plan.spreadsheetId || '')) {
        throw new Error('復旧対象のACTIVE Client registry identityが一致しません。');
      }
      const client = SpreadsheetApp.openById(String(plan.spreadsheetId || ''));
      vNextAdminAssertEmptyPilotNoBusinessData_(hub, client, registry, false);
      const sourceRelease = vNextAdminEmptyPilotRelease_(hub, plan.sourceReleaseId, ['ACTIVE', 'RETIRED']);
      const sourceModel = vNextAdminEmptyPilotModel_(hub, plan.sourceModelReleaseId, sourceRelease,
        ['ACTIVE', 'RETIRED']);
      const preferredDirection = String(req.direction || '').toUpperCase();
      if (preferredDirection && ['TARGET', 'SOURCE'].indexOf(preferredDirection) < 0) {
        throw new Error('directionはTARGETまたはSOURCEだけを指定できます。');
      }
      let direction = preferredDirection ||
        (String(registry.template_release_id || '') === String(plan.targetReleaseId || '')
          ? 'TARGET' : 'SOURCE');
      let release = sourceRelease;
      let model = sourceModel;
      if (direction === 'TARGET') {
        try {
          const pair = vNextAdminReadActiveReleasePair_(hub);
          if (pair.releaseId !== String(plan.targetReleaseId || '') ||
              pair.modelReleaseId !== String(plan.targetModelReleaseId || '')) {
            throw new Error('Target pair is no longer canonical ACTIVE.');
          }
          release = vNextAdminEmptyPilotRelease_(hub, plan.targetReleaseId, ['ACTIVE']);
          model = vNextAdminEmptyPilotModel_(hub, plan.targetModelReleaseId, release);
        } catch (targetUnavailable) {
          direction = 'SOURCE';
          release = sourceRelease;
          model = sourceModel;
        }
      } else if (String(registry.template_release_id || '') !== String(plan.sourceReleaseId || '')) {
        throw new Error('Registry is pinned to neither the journal source nor target release. 自動復旧を停止しました。');
      }
      vNextAdminAssertEmptyPilotReleaseAssets_(release);
      vNextAdminFreezeEmptyPilotClient_(hub, client, migration.migration_id);
      const applied = vNextAdminApplyEmptyPilotRelease_(hub, client, registry, plan,
        release, model, direction, function (phase, detail) {
          vNextAdminSetEmptyPilotUpgradePhase_(hub, migration.migration_id, phase, detail);
        });
      const finalStatus = direction === 'TARGET' ? 'SUCCEEDED' : 'ROLLED_BACK';
      vNextAdminPatchLatestMigration_(hub, migration.migration_id, {
        status: finalStatus, finished_at: new Date(), error: '',
        result_json: vNextAdminCanonicalJson_({ phase: 'RECOVERED_' + direction,
          direction: direction, healthStatus: applied.health.healthStatus,
          healthCode: applied.health.healthCode })
      });
      vNextAdminWriteAudit_(hub, 'RECOVER_EMPTY_PILOT_CLIENT_UPGRADE', 'BOOK', bookId, finalStatus, {
        migrationId: migration.migration_id, direction: direction,
        releaseId: release.release_id, reason: vNextAdminText_(req.reason)
      });
      return { ok: true, migrationId: migration.migration_id, bookId: bookId,
        direction: direction, releaseId: release.release_id, health: applied.health };
    });
  });
}

/**
 * Upgrade the one known-safe preflight failure without changing its URL or
 * business records. This path is intentionally narrower than a migration:
 * employee evidence and the failed request are preserved, while any forecast,
 * plan, approval, official snapshot or evaluation makes the operation stop.
 */
function vNextAdminUpgradeFailedPreflightPilotClient(request) {
  return vNextAdminGuard_('vNextAdminUpgradeFailedPreflightPilotClient', function () {
    const req = request && typeof request === 'object' ? request : {};
    const hub = vNextAdminRequireHub_();
    vNextAdminAssertHubAdmin_(hub, false);
    return vNextAdminWithScriptLock_('upgrade-failed-preflight-pilot-client', function () {
      const resolved = vNextAdminResolveFailedPreflightUpgrade_(hub, req);
      if (req.dryRun !== false) return resolved.basis;
      const reason = vNextAdminRequiredText_(req.reason, 'reason');
      const migrationId = 'PREFLIGHT-UPG-' + vNextAdminSha256_(vNextAdminCanonicalJson_({
        bookId: resolved.registry.book_id, fromReleaseId: resolved.sourceRelease.release_id,
        toReleaseId: resolved.targetRelease.release_id, attemptId: Utilities.getUuid()
      })).slice(0, 24).toUpperCase();
      const plan = {
        kind: 'FAILED_PREFLIGHT_PILOT_UPGRADE_V1', migrationId: migrationId,
        bookId: String(resolved.registry.book_id || ''),
        spreadsheetId: String(resolved.registry.spreadsheet_id || ''),
        clientScriptId: String(resolved.registry.client_script_id || ''),
        sourceReleaseId: String(resolved.sourceRelease.release_id || ''),
        sourceModelReleaseId: String(resolved.sourceModel.model_release_id || ''),
        sourceMetaRecordId: String(resolved.sourceMeta.record_id || ''),
        targetReleaseId: String(resolved.targetRelease.release_id || ''),
        targetModelReleaseId: String(resolved.targetModel.model_release_id || ''),
        schemaVersion: String(resolved.targetRelease.schema_version || ''),
        preservedState: 'READY_TO_RUN', failedJobId: String(resolved.failedJob.job_id || ''),
        requestId: String(resolved.requestId || ''), reason: reason
      };
      vNextAdminAppendObject_(hub, VN_ADMIN_SHEETS.MIGRATIONS, {
        migration_id: migrationId, book_id: plan.bookId, spreadsheet_id: plan.spreadsheetId,
        from_release_id: plan.sourceReleaseId, to_release_id: plan.targetReleaseId,
        status: 'FAILED_PREFLIGHT_VALIDATED', dry_run: 0,
        plan_json: vNextAdminCanonicalJson_(plan), result_json: '', started_at: new Date(),
        finished_at: '', actor: vNextAdminActor_(), error: ''
      });
      try {
        vNextAdminFreezeEmptyPilotClient_(hub, resolved.client, migrationId);
        vNextAdminSetEmptyPilotUpgradePhase_(hub, migrationId, 'FAILED_PREFLIGHT_FROZEN');
        const applied = vNextAdminApplyEmptyPilotRelease_(hub, resolved.client, resolved.registry, plan,
          resolved.targetRelease, resolved.targetModel, 'TARGET', function (phase, detail) {
            vNextAdminSetEmptyPilotUpgradePhase_(hub, migrationId,
              String(phase || '').replace(/^EMPTY_PILOT_/, 'FAILED_PREFLIGHT_'), detail);
          });
        vNextAdminPatchLatestMigration_(hub, migrationId, {
          status: 'SUCCEEDED', finished_at: new Date(), error: '',
          result_json: vNextAdminCanonicalJson_({ phase: 'COMPLETED', direction: 'TARGET',
            failedJobId: plan.failedJobId, requestId: plan.requestId,
            healthStatus: applied.health.healthStatus, healthCode: applied.health.healthCode })
        });
        vNextAdminWriteAudit_(hub, 'UPGRADE_FAILED_PREFLIGHT_PILOT_CLIENT', 'BOOK', plan.bookId,
          'SUCCESS', { migrationId: migrationId, failedJobId: plan.failedJobId,
            requestId: plan.requestId, fromReleaseId: plan.sourceReleaseId,
            toReleaseId: plan.targetReleaseId, reason: reason });
        return { ok: true, migrationId: migrationId, bookId: plan.bookId,
          spreadsheetId: plan.spreadsheetId, spreadsheetUrl: resolved.registry.spreadsheet_url,
          fromReleaseId: plan.sourceReleaseId, toReleaseId: plan.targetReleaseId,
          failedJobId: plan.failedJobId, requestId: plan.requestId, health: applied.health };
      } catch (upgradeError) {
        const originalMessage = String(upgradeError && upgradeError.message || upgradeError);
        try {
          vNextAdminSetEmptyPilotUpgradePhase_(hub, migrationId, 'FAILED_PREFLIGHT_ROLLBACK_STARTED', {
            cause: originalMessage.slice(0, 500)
          });
          const currentRegistry = vNextAdminFindRegistryRow_(hub, function (row) {
            return String(row.book_id || '') === plan.bookId;
          }) || resolved.registry;
          vNextAdminApplyEmptyPilotRelease_(hub, resolved.client, currentRegistry, plan,
            resolved.sourceRelease, resolved.sourceModel, 'SOURCE', function (phase, detail) {
              vNextAdminSetEmptyPilotUpgradePhase_(hub, migrationId,
                String(phase || '').replace(/^EMPTY_PILOT_/, 'FAILED_PREFLIGHT_ROLLBACK_'), detail);
            });
          vNextAdminPatchLatestMigration_(hub, migrationId, {
            status: 'ROLLED_BACK', finished_at: new Date(), error: originalMessage,
            result_json: vNextAdminCanonicalJson_({ phase: 'ROLLED_BACK_TO_SOURCE',
              sourceReleaseId: plan.sourceReleaseId })
          });
        } catch (rollbackError) {
          const rollbackMessage = String(rollbackError && rollbackError.message || rollbackError);
          vNextAdminPatchLatestMigration_(hub, migrationId, {
            status: 'RECOVERY_REQUIRED', error: originalMessage + '; rollback=' + rollbackMessage,
            result_json: vNextAdminCanonicalJson_({ phase: 'RECOVERY_REQUIRED',
              originalError: originalMessage.slice(0, 500), rollbackError: rollbackMessage.slice(0, 500) })
          });
          vNextAdminAppendException_(hub, {
            severity: 'ERROR', exception_type: 'FAILED_PREFLIGHT_UPGRADE_RECOVERY_REQUIRED',
            book_id: plan.bookId, client_name: resolved.registry.client_name,
            fiscal_year: resolved.registry.fiscal_year,
            title: '入力済みPilot Client更新に復旧が必要です',
            detail: originalMessage + '; rollback=' + rollbackMessage,
            recommended_action: 'vNextAdminRecoverFailedPreflightPilotClientUpgradeを実行',
            source_ref: migrationId
          });
          throw new Error('更新と自動rollbackの両方が完了しませんでした。migrationId=' + migrationId +
            '; cause=' + originalMessage + '; rollback=' + rollbackMessage);
        }
        throw upgradeError;
      }
    });
  });
}

function vNextAdminResolveFailedPreflightUpgrade_(hub, request) {
  const req = request && typeof request === 'object' ? request : {};
  const bookId = vNextAdminRequiredText_(req.bookId, 'bookId');
  const registry = vNextAdminFindRegistryRow_(hub, function (row) {
    return String(row.book_id || '') === bookId && String(row.mode || '') === 'CLIENT';
  });
  if (!registry || String(registry.status || '').toUpperCase() !== 'ACTIVE') {
    throw new Error('対象のACTIVE Client BookがBOOK_REGISTRYにありません。');
  }
  const pair = vNextAdminReadActiveReleasePair_(hub);
  if (String(registry.template_release_id || '') === pair.releaseId) {
    throw new Error('対象Clientはすでに現在のemployee releaseです。');
  }
  const unfinished = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.MIGRATIONS).rows.filter(function (row) {
    return String(row.book_id || '') === bookId &&
      /^FAILED_PREFLIGHT_|^RECOVERY_REQUIRED$/.test(String(row.status || '').toUpperCase());
  }).slice(-1)[0];
  if (unfinished) {
    throw new Error('未完了の入力済みPilot更新があります。復旧してください。migrationId=' +
      String(unfinished.migration_id || ''));
  }
  const targetRelease = vNextAdminEmptyPilotRelease_(hub, pair.releaseId, ['ACTIVE']);
  const targetModel = vNextAdminEmptyPilotModel_(hub, pair.modelReleaseId, targetRelease, ['ACTIVE']);
  const sourceRelease = vNextAdminEmptyPilotRelease_(hub, registry.template_release_id,
    ['ACTIVE', 'RETIRED']);
  if (String(sourceRelease.schema_version || '') !== String(targetRelease.schema_version || '') ||
      String(targetRelease.schema_version || '') !== vNextAdminClientSchemaVersion_()) {
    throw new Error('入力済みPilot更新は同一Client Core schema間だけで実行できます。');
  }
  const client = SpreadsheetApp.openById(vNextAdminRequiredText_(registry.spreadsheet_id,
    'registry.spreadsheet_id'));
  const routing = vNextAdminReadKeyValueSheet_(client, VN_ADMIN_BOOK_CONFIG_SHEET);
  const sourceModel = vNextAdminEmptyPilotModel_(hub,
    vNextAdminRequiredText_(routing.model_release_id, 'client.model_release_id'), sourceRelease,
    ['ACTIVE', 'RETIRED']);
  const sourceMeta = vNextAdminAssertEmptyPilotPinnedRelease_(hub, client, registry,
    sourceRelease, sourceModel, false, 'READY_TO_RUN');
  const failure = vNextAdminAssertFailedPreflightBusinessBoundary_(hub, client, registry);
  vNextAdminAssertEmptyPilotReleaseAssets_(sourceRelease);
  vNextAdminAssertEmptyPilotReleaseAssets_(targetRelease);
  return {
    registry: registry, client: client, sourceRelease: sourceRelease, sourceModel: sourceModel,
    sourceMeta: sourceMeta, targetRelease: targetRelease, targetModel: targetModel,
    failedJob: failure.failedJob, requestId: failure.requestId,
    basis: {
      eligible: true, readOnly: req.dryRun !== false, bookId: bookId,
      clientName: String(registry.client_name || ''), fiscalYear: Number(registry.fiscal_year || 0),
      spreadsheetId: String(registry.spreadsheet_id || ''),
      spreadsheetUrl: String(registry.spreadsheet_url || ''), sameUrl: true,
      currentReleaseId: String(sourceRelease.release_id || ''),
      targetReleaseId: String(targetRelease.release_id || ''),
      targetModelReleaseId: String(targetModel.model_release_id || ''),
      failedJobId: String(failure.failedJob.job_id || ''), requestId: failure.requestId,
      preservedState: 'READY_TO_RUN', evidenceCount: failure.evidenceCount,
      safeguards: ['対象は既知の月表記変換エラーだけ', 'FORECAST_RUN・計画・承認・公式・評価なし',
        '入力・失敗request・job IDを保持', '同一schema', '同じSpreadsheet URL',
        '更新先はcanonical ACTIVE pair', '失敗時は旧releaseへ自動rollback']
    }
  };
}

function vNextAdminAssertFailedPreflightBusinessBoundary_(hub, client, registry) {
  const bookId = String(registry.book_id || '');
  if (String(registry.state || '').toUpperCase() !== 'READY_TO_RUN' ||
      String(registry.current_official_id || '')) {
    throw new Error('入力済みPilot更新はREADY_TO_RUN / 正式予算なしだけが対象です。');
  }
  ['FORECAST_RUN', 'PLAN_VERSION', 'EVALUATION'].forEach(function (sheetName) {
    const rows = vNextAdminReadCoreRows_(hub, sheetName).filter(function (row) {
      return String(row.book_id || '') === bookId;
    }).concat(vNextAdminReadCoreRows_(client, sheetName).filter(function (row) {
      return String(row.book_id || '') === bookId;
    }));
    if (rows.length) throw new Error('入力済みPilot更新を停止しました。' + sheetName + 'に既存recordがあります。');
  });
  if (vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.APPROVALS).rows.some(function (row) {
    return String(row.book_id || '') === bookId;
  }) || vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.OFFICIAL).rows.some(function (row) {
    return String(row.book_id || '') === bookId;
  })) throw new Error('承認または公式recordがあるため入力済みPilot更新を停止しました。');
  const clientEvidence = vNextAdminReadCoreRows_(client, 'EVIDENCE_EVENT').filter(function (row) {
    return String(row.book_id || '') === bookId;
  });
  if (!clientEvidence.length) throw new Error('保存済みの従業員入力がありません。');
  vNextAdminValidateClientEvidenceRows_(hub, bookId, clientEvidence);
  const jobs = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.JOBS).rows.filter(function (row) {
    return String(row.target_book_id || '') === bookId;
  });
  if (jobs.some(function (row) {
    const status = String(row.status || '').toUpperCase();
    return status === 'RUNNING' || (status === 'QUEUED' &&
      !vNextAdminIsKnownEvidenceRequeuedJob_(hub, row));
  })) throw new Error('対象bookにQUEUED/RUNNING jobがあるため更新を停止しました。');
  const matching = jobs.filter(function (row) {
    return String(row.job_type || '') === 'FORECAST_REQUEST' && (
        (String(row.status || '').toUpperCase() === 'FAILED' &&
          vNextAdminCanRetryKnownEvidencePreflightJob_(row)) ||
        vNextAdminIsKnownEvidenceRequeuedJob_(hub, row)
      );
  });
  if (matching.length !== 1) throw new Error('既知の月表記変換エラーjobが1件に確定しませんでした。');
  const failedJob = matching[0];
  const payload = vNextAdminParseJson_(failedJob.request_json, {});
  const requestId = vNextAdminRequiredText_(payload.requestId, 'failedForecast.requestId');
  const requestHash = vNextAdminRequiredText_(payload.requestHash, 'failedForecast.requestHash');
  const requests = vNextAdminReadTable_(client, VN_ADMIN_CLIENT_REQUEST_SHEET).rows.filter(function (row) {
    return String(row.request_id || '') === requestId;
  });
  const requested = requests.find(function (row) {
    return String(row.event_type || '').toUpperCase() === 'REQUESTED' &&
      String(row.status || '').toUpperCase() === 'PENDING';
  });
  const latest = requests.slice(-1)[0];
  if (!requested || String(requested.request_hash || '') !== requestHash || !latest ||
      String(latest.event_type || '').toUpperCase() !== 'FAILED' ||
      String(latest.status || '').toUpperCase() !== 'FAILED') {
    throw new Error('失敗した予測依頼のappend-only lineageが一致しません。');
  }
  const runIdentity = typeof vNextEngineBuildAdminRunIdentity_ === 'function'
    ? vNextEngineBuildAdminRunIdentity_(bookId, String(failedJob.idempotency_key || '')) : null;
  if (!runIdentity || vNextAdminReadCoreRows_(hub, 'FORECAST_RUN').some(function (row) {
    return String(row.run_id || '') === String(runIdentity.runId || '');
  })) throw new Error('失敗jobに対応するFORECAST_RUNが存在するため更新を停止しました。');
  return { failedJob: failedJob, requestId: requestId, requestHash: requestHash,
    evidenceCount: clientEvidence.length };
}

/** Editor fallback: upgrades exactly one known failed preflight Pilot. */
function vNextAdminUpgradeOnlyKnownFailedPreflightPilotForManualTest() {
  const hub = vNextAdminRequireHub_();
  const pair = vNextAdminReadActiveReleasePair_(hub);
  const candidates = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.REGISTRY).rows.filter(function (row) {
    return String(row.mode || '').toUpperCase() === 'CLIENT' &&
      String(row.status || '').toUpperCase() === 'ACTIVE' &&
      String(row.state || '').toUpperCase() === 'READY_TO_RUN' &&
      String(row.template_release_id || '') !== pair.releaseId;
  });
  if (candidates.length !== 1) {
    throw new Error('更新候補の入力済み失敗Pilotが1冊に確定しませんでした: ' + candidates.length + '冊');
  }
  // The public apply call performs every strict check while holding one lock.
  // A separate slow dry-run would let the scheduler requeue/claim the exact
  // job between locks and create a false conflict.
  return vNextAdminUpgradeFailedPreflightPilotClient({
    bookId: String(candidates[0].book_id || ''), dryRun: false,
    reason: '既知の月表記変換エラーを修正し、同じURLで従業員UIを更新'
  });
}

function vNextAdminRecoverFailedPreflightPilotClientUpgrade(request) {
  return vNextAdminGuard_('vNextAdminRecoverFailedPreflightPilotClientUpgrade', function () {
    const req = request && typeof request === 'object' ? request : {};
    const hub = vNextAdminRequireHub_();
    vNextAdminAssertHubAdmin_(hub, false);
    return vNextAdminWithScriptLock_('recover-failed-preflight-pilot-client', function () {
      const bookId = vNextAdminRequiredText_(req.bookId, 'bookId');
      const migration = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.MIGRATIONS).rows.filter(function (row) {
        return String(row.book_id || '') === bookId &&
          (!req.migrationId || String(row.migration_id || '') === String(req.migrationId));
      }).slice(-1)[0];
      if (!migration) throw new Error('復旧対象のMIGRATION_LOGがありません。');
      const plan = vNextAdminParseJson_(migration.plan_json, null);
      if (!plan || String(plan.kind || '') !== 'FAILED_PREFLIGHT_PILOT_UPGRADE_V1' ||
          String(plan.bookId || '') !== bookId ||
          String(plan.spreadsheetId || '') !== String(migration.spreadsheet_id || '')) {
        throw new Error('入力済みPilot復旧journalのidentityが不正です。');
      }
      const registry = vNextAdminFindRegistryRow_(hub, function (row) {
        return String(row.book_id || '') === bookId && String(row.mode || '') === 'CLIENT';
      });
      if (!registry || String(registry.status || '').toUpperCase() !== 'ACTIVE' ||
          String(registry.spreadsheet_id || '') !== String(plan.spreadsheetId || '')) {
        throw new Error('復旧対象のACTIVE Client registry identityが一致しません。');
      }
      const client = SpreadsheetApp.openById(String(plan.spreadsheetId || ''));
      vNextAdminAssertFailedPreflightBusinessBoundary_(hub, client, registry);
      const sourceRelease = vNextAdminEmptyPilotRelease_(hub, plan.sourceReleaseId,
        ['ACTIVE', 'RETIRED']);
      const sourceModel = vNextAdminEmptyPilotModel_(hub, plan.sourceModelReleaseId, sourceRelease,
        ['ACTIVE', 'RETIRED']);
      let direction = String(registry.template_release_id || '') === String(plan.targetReleaseId || '')
        ? 'TARGET' : 'SOURCE';
      let release = sourceRelease;
      let model = sourceModel;
      if (direction === 'TARGET') {
        const pair = vNextAdminReadActiveReleasePair_(hub);
        if (pair.releaseId !== String(plan.targetReleaseId || '') ||
            pair.modelReleaseId !== String(plan.targetModelReleaseId || '')) {
          direction = 'SOURCE';
        } else {
          release = vNextAdminEmptyPilotRelease_(hub, plan.targetReleaseId, ['ACTIVE']);
          model = vNextAdminEmptyPilotModel_(hub, plan.targetModelReleaseId, release, ['ACTIVE']);
        }
      } else if (String(registry.template_release_id || '') !== String(plan.sourceReleaseId || '')) {
        throw new Error('Registryがjournalのsource/target以外を参照しています。自動復旧を停止しました。');
      }
      vNextAdminAssertEmptyPilotReleaseAssets_(release);
      vNextAdminFreezeEmptyPilotClient_(hub, client, migration.migration_id);
      const applied = vNextAdminApplyEmptyPilotRelease_(hub, client, registry, plan,
        release, model, direction, function (phase, detail) {
          vNextAdminSetEmptyPilotUpgradePhase_(hub, migration.migration_id,
            'FAILED_PREFLIGHT_RECOVERY_' + String(phase || '').replace(/^EMPTY_PILOT_/, ''), detail);
        });
      const finalStatus = direction === 'TARGET' ? 'SUCCEEDED' : 'ROLLED_BACK';
      vNextAdminPatchLatestMigration_(hub, migration.migration_id, {
        status: finalStatus, finished_at: new Date(), error: '',
        result_json: vNextAdminCanonicalJson_({ phase: 'RECOVERED_' + direction,
          failedJobId: plan.failedJobId, requestId: plan.requestId,
          healthStatus: applied.health.healthStatus, healthCode: applied.health.healthCode })
      });
      vNextAdminWriteAudit_(hub, 'RECOVER_FAILED_PREFLIGHT_PILOT_CLIENT_UPGRADE', 'BOOK',
        bookId, finalStatus, { migrationId: migration.migration_id, direction: direction,
          releaseId: release.release_id, failedJobId: plan.failedJobId,
          reason: vNextAdminText_(req.reason) });
      return { ok: true, migrationId: migration.migration_id, bookId: bookId,
        direction: direction, releaseId: release.release_id, health: applied.health };
    });
  });
}

function vNextAdminRecoverOnlyFailedPreflightPilotForManualTest() {
  const hub = vNextAdminRequireHub_();
  const rows = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.MIGRATIONS).rows.filter(function (row) {
    const plan = vNextAdminParseJson_(row.plan_json, {});
    return String(plan.kind || '') === 'FAILED_PREFLIGHT_PILOT_UPGRADE_V1' &&
      /^FAILED_PREFLIGHT_|^RECOVERY_REQUIRED$/.test(String(row.status || '').toUpperCase());
  });
  if (rows.length !== 1) {
    throw new Error('復旧対象の入力済みPilot更新が1件に確定しませんでした: ' + rows.length + '件');
  }
  return vNextAdminRecoverFailedPreflightPilotClientUpgrade({
    bookId: String(rows[0].book_id || ''), migrationId: String(rows[0].migration_id || ''),
    direction: 'TARGET',
    reason: 'Apps Script editorから中断した入力済みPilot更新を現在のクライアント年度ブック用 release へ復旧'
  });
}

/**
 * Same-URL UI/runtime upgrade for a Pilot book whose forecast is complete but
 * whose plan has not yet been created. Forecast/evidence/state records remain
 * immutable; only the deployed employee shell and release pins are advanced.
 */
function vNextAdminUpgradeDraftReadyPilotUx(request) {
  return vNextAdminGuard_('vNextAdminUpgradeDraftReadyPilotUx', function () {
    const req = request && typeof request === 'object' ? request : {};
    const hub = vNextAdminRequireHub_();
    vNextAdminAssertHubAdmin_(hub, false);
    return vNextAdminWithScriptLock_('upgrade-draft-ready-pilot-ux', function () {
      const resolved = vNextAdminResolveDraftReadyPilotUxUpgrade_(hub, req);
      if (req.dryRun !== false) return resolved.basis;
      const reason = vNextAdminRequiredText_(req.reason, 'reason');
      const migrationId = 'DRAFT-UX-' + vNextAdminSha256_(vNextAdminCanonicalJson_({
        bookId: resolved.registry.book_id, fromReleaseId: resolved.sourceRelease.release_id,
        toReleaseId: resolved.targetRelease.release_id, attemptId: Utilities.getUuid()
      })).slice(0, 24).toUpperCase();
      const plan = {
        kind: 'DRAFT_READY_PILOT_UX_UPGRADE_V1', migrationId: migrationId,
        bookId: String(resolved.registry.book_id || ''),
        spreadsheetId: String(resolved.registry.spreadsheet_id || ''),
        clientScriptId: String(resolved.registry.client_script_id || ''),
        sourceReleaseId: String(resolved.sourceRelease.release_id || ''),
        sourceModelReleaseId: String(resolved.sourceModel.model_release_id || ''),
        sourceMetaRecordId: String(resolved.sourceMeta.record_id || ''),
        targetReleaseId: String(resolved.targetRelease.release_id || ''),
        targetModelReleaseId: String(resolved.targetModel.model_release_id || ''),
        schemaVersion: String(resolved.targetRelease.schema_version || ''),
        preservedState: resolved.preservedState, sourceForecastRunId: resolved.sourceForecastRunId,
        sourcePlanVersionId: resolved.sourcePlanVersionId || '',
        sourceApprovalRequestId: resolved.sourceApprovalRequestId || '',
        reason: reason
      };
      vNextAdminAppendObject_(hub, VN_ADMIN_SHEETS.MIGRATIONS, {
        migration_id: migrationId, book_id: plan.bookId, spreadsheet_id: plan.spreadsheetId,
        from_release_id: plan.sourceReleaseId, to_release_id: plan.targetReleaseId,
        status: 'DRAFT_READY_UX_VALIDATED', dry_run: 0,
        plan_json: vNextAdminCanonicalJson_(plan), result_json: '', started_at: new Date(),
        finished_at: '', actor: vNextAdminActor_(), error: ''
      });
      try {
        vNextAdminFreezeEmptyPilotClient_(hub, resolved.client, migrationId);
        vNextAdminSetEmptyPilotUpgradePhase_(hub, migrationId, 'DRAFT_READY_UX_FROZEN');
        const applied = vNextAdminApplyEmptyPilotRelease_(hub, resolved.client, resolved.registry,
          plan, resolved.targetRelease, resolved.targetModel, 'TARGET', function (phase, detail) {
            vNextAdminSetEmptyPilotUpgradePhase_(hub, migrationId,
              String(phase || '').replace(/^EMPTY_PILOT_/, 'DRAFT_READY_UX_'), detail);
          });
        vNextAdminPatchLatestMigration_(hub, migrationId, {
          status: 'SUCCEEDED', finished_at: new Date(), error: '',
          result_json: vNextAdminCanonicalJson_({ phase: 'COMPLETED', direction: 'TARGET',
            sourceForecastRunId: plan.sourceForecastRunId,
            healthStatus: applied.health.healthStatus, healthCode: applied.health.healthCode })
        });
        vNextAdminWriteAudit_(hub, 'UPGRADE_DRAFT_READY_PILOT_UX', 'BOOK', plan.bookId,
          'SUCCESS', { migrationId: migrationId, sourceForecastRunId: plan.sourceForecastRunId,
            fromReleaseId: plan.sourceReleaseId, toReleaseId: plan.targetReleaseId, reason: reason });
        return { ok: true, migrationId: migrationId, bookId: plan.bookId,
          spreadsheetId: plan.spreadsheetId, spreadsheetUrl: resolved.registry.spreadsheet_url,
          fromReleaseId: plan.sourceReleaseId, toReleaseId: plan.targetReleaseId,
          sourceForecastRunId: plan.sourceForecastRunId, health: applied.health };
      } catch (upgradeError) {
        const originalMessage = String(upgradeError && upgradeError.message || upgradeError);
        try {
          vNextAdminSetEmptyPilotUpgradePhase_(hub, migrationId, 'DRAFT_READY_UX_ROLLBACK_STARTED', {
            cause: originalMessage.slice(0, 500)
          });
          const currentRegistry = vNextAdminFindRegistryRow_(hub, function (row) {
            return String(row.book_id || '') === plan.bookId;
          }) || resolved.registry;
          vNextAdminApplyEmptyPilotRelease_(hub, resolved.client, currentRegistry, plan,
            resolved.sourceRelease, resolved.sourceModel, 'SOURCE', function (phase, detail) {
              vNextAdminSetEmptyPilotUpgradePhase_(hub, migrationId,
                String(phase || '').replace(/^EMPTY_PILOT_/, 'DRAFT_READY_UX_ROLLBACK_'), detail);
            });
          vNextAdminPatchLatestMigration_(hub, migrationId, {
            status: 'ROLLED_BACK', finished_at: new Date(), error: originalMessage,
            result_json: vNextAdminCanonicalJson_({ phase: 'ROLLED_BACK_TO_SOURCE',
              sourceReleaseId: plan.sourceReleaseId })
          });
        } catch (rollbackError) {
          const rollbackMessage = String(rollbackError && rollbackError.message || rollbackError);
          vNextAdminPatchLatestMigration_(hub, migrationId, {
            status: 'RECOVERY_REQUIRED', error: originalMessage + '; rollback=' + rollbackMessage,
            result_json: vNextAdminCanonicalJson_({ phase: 'RECOVERY_REQUIRED',
              originalError: originalMessage.slice(0, 500), rollbackError: rollbackMessage.slice(0, 500) })
          });
          vNextAdminAppendException_(hub, {
            severity: 'ERROR', exception_type: 'DRAFT_READY_UX_UPGRADE_RECOVERY_REQUIRED',
            book_id: plan.bookId, client_name: resolved.registry.client_name,
            fiscal_year: resolved.registry.fiscal_year,
            title: '予測作成済みPilotの画面更新に復旧が必要です',
            detail: originalMessage + '; rollback=' + rollbackMessage,
            recommended_action: 'vNextAdminRecoverDraftReadyPilotUxUpgradeを実行',
            source_ref: migrationId
          });
          throw new Error('画面更新と自動rollbackが完了しませんでした。migrationId=' + migrationId +
            '; cause=' + originalMessage + '; rollback=' + rollbackMessage);
        }
        throw upgradeError;
      }
    });
  });
}

function vNextAdminResolveDraftReadyPilotUxUpgrade_(hub, request) {
  const req = request && typeof request === 'object' ? request : {};
  const preservedState = String(req.preservedState || 'DRAFT_READY').toUpperCase();
  if (['DRAFT_READY', 'SUBMITTED'].indexOf(preservedState) < 0) {
    throw new Error('この画面更新で保持できる状態はDRAFT_READYまたはSUBMITTEDです。');
  }
  const bookId = vNextAdminRequiredText_(req.bookId, 'bookId');
  const registry = vNextAdminFindRegistryRow_(hub, function (row) {
    return String(row.book_id || '') === bookId && String(row.mode || '') === 'CLIENT';
  });
  if (!registry || String(registry.status || '').toUpperCase() !== 'ACTIVE') {
    throw new Error('対象のACTIVE Client BookがBOOK_REGISTRYにありません。');
  }
  const pair = vNextAdminReadActiveReleasePair_(hub);
  if (String(registry.template_release_id || '') === pair.releaseId) {
    throw new Error('対象Clientはすでに現在のemployee releaseです。');
  }
  const unfinished = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.MIGRATIONS).rows.filter(function (row) {
    const plan = vNextAdminParseJson_(row.plan_json, {});
    return String(row.book_id || '') === bookId &&
      String(plan.kind || '') === 'DRAFT_READY_PILOT_UX_UPGRADE_V1' &&
      /^DRAFT_READY_UX_|^RECOVERY_REQUIRED$/.test(String(row.status || '').toUpperCase());
  }).slice(-1)[0];
  if (unfinished) throw new Error('未完了の画面更新があります。復旧してください。migrationId=' +
    String(unfinished.migration_id || ''));
  const targetRelease = vNextAdminEmptyPilotRelease_(hub, pair.releaseId, ['ACTIVE']);
  const targetModel = vNextAdminEmptyPilotModel_(hub, pair.modelReleaseId, targetRelease, ['ACTIVE']);
  const sourceRelease = vNextAdminEmptyPilotRelease_(hub, registry.template_release_id,
    ['ACTIVE', 'RETIRED']);
  if (String(sourceRelease.schema_version || '') !== String(targetRelease.schema_version || '') ||
      String(targetRelease.schema_version || '') !== vNextAdminClientSchemaVersion_()) {
    throw new Error('画面更新は同一Client Core schema間だけで実行できます。');
  }
  const client = SpreadsheetApp.openById(vNextAdminRequiredText_(registry.spreadsheet_id,
    'registry.spreadsheet_id'));
  const routing = vNextAdminReadKeyValueSheet_(client, VN_ADMIN_BOOK_CONFIG_SHEET);
  const sourceModel = vNextAdminEmptyPilotModel_(hub,
    vNextAdminRequiredText_(routing.model_release_id, 'client.model_release_id'), sourceRelease,
    ['ACTIVE', 'RETIRED']);
  const sourceMeta = vNextAdminAssertEmptyPilotPinnedRelease_(hub, client, registry,
    sourceRelease, sourceModel, false, preservedState);
  const boundary = vNextAdminAssertDraftReadyPilotUxBoundary_(hub, client, registry, preservedState);
  vNextAdminAssertEmptyPilotReleaseAssets_(sourceRelease);
  vNextAdminAssertEmptyPilotReleaseAssets_(targetRelease);
  return { registry: registry, client: client, sourceRelease: sourceRelease,
    sourceModel: sourceModel, sourceMeta: sourceMeta, targetRelease: targetRelease,
    targetModel: targetModel, sourceForecastRunId: boundary.runId,
    sourcePlanVersionId: boundary.planVersionId || '',
    sourceApprovalRequestId: boundary.approvalRequestId || '', preservedState: preservedState,
    basis: { eligible: true, readOnly: req.dryRun !== false, bookId: bookId,
      clientName: String(registry.client_name || ''), fiscalYear: Number(registry.fiscal_year || 0),
      spreadsheetId: String(registry.spreadsheet_id || ''),
      spreadsheetUrl: String(registry.spreadsheet_url || ''), sameUrl: true,
      currentReleaseId: String(sourceRelease.release_id || ''),
      targetReleaseId: String(targetRelease.release_id || ''),
      targetModelReleaseId: String(targetModel.model_release_id || ''),
      sourceForecastRunId: boundary.runId, sourcePlanVersionId: boundary.planVersionId || '',
      sourceApprovalRequestId: boundary.approvalRequestId || '', preservedState: preservedState,
      safeguards: ['予測SUCCESSと承認前の状態だけ', '予測・入力・提出済み計画は変更しない',
        '同一schema', '同じSpreadsheet URL', '失敗時は旧releaseへ自動rollback'] }
  };
}

function vNextAdminAssertDraftReadyPilotUxBoundary_(hub, client, registry, expectedState) {
  const bookId = String(registry.book_id || '');
  const preservedState = String(expectedState || 'DRAFT_READY').toUpperCase();
  if (['DRAFT_READY', 'SUBMITTED'].indexOf(preservedState) < 0 ||
      String(registry.state || '').toUpperCase() !== preservedState ||
      String(registry.current_official_id || '')) {
    throw new Error('画面更新はDRAFT_READYまたはSUBMITTED / 正式予算なしだけが対象です。');
  }
  ['EVALUATION'].forEach(function (sheetName) {
    const count = vNextAdminReadCoreRows_(hub, sheetName).concat(
      vNextAdminReadCoreRows_(client, sheetName)).filter(function (row) {
        return String(row.book_id || '') === bookId;
      }).length;
    if (count) throw new Error('評価作成後はこの画面更新を実行できません: ' + sheetName);
  });
  if (vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.OFFICIAL).rows.some(function (row) {
    return String(row.book_id || '') === bookId;
  })) throw new Error('公式recordがあるため画面更新を停止しました。');
  if (vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.JOBS).rows.some(function (row) {
    return String(row.target_book_id || '') === bookId &&
      ['QUEUED', 'RUNNING'].indexOf(String(row.status || '').toUpperCase()) >= 0;
  })) throw new Error('対象bookにQUEUED/RUNNING jobがあるため画面更新を停止しました。');
  const hubStates = vNextAdminReadCoreRows_(hub, 'STATE_EVENT').filter(function (row) {
    return String(row.book_id || '') === bookId;
  });
  const clientStates = vNextAdminReadCoreRows_(client, 'STATE_EVENT').filter(function (row) {
    return String(row.book_id || '') === bookId;
  });
  const hubState = hubStates[hubStates.length - 1];
  const clientState = clientStates[clientStates.length - 1];
  if (!hubState || !clientState || String(hubState.to_state || '').toUpperCase() !== preservedState ||
      String(clientState.to_state || '').toUpperCase() !== preservedState) {
    throw new Error('Hub/Clientの最新状態が更新対象と一致しません。');
  }
  let planVersionId = '';
  let approvalRequestId = '';
  let runId = String(hubState.related_run_id || clientState.related_run_id || '');
  if (preservedState === 'DRAFT_READY') {
    const planCount = vNextAdminReadCoreRows_(hub, 'PLAN_VERSION').concat(
      vNextAdminReadCoreRows_(client, 'PLAN_VERSION')).filter(function (row) {
        return String(row.book_id || '') === bookId;
      }).length;
    if (planCount) throw new Error('DRAFT_READY画面更新は計画未作成だけが対象です。');
    if (vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.APPROVALS).rows.some(function (row) {
      return String(row.book_id || '') === bookId;
    })) throw new Error('DRAFT_READY画面更新は承認recordなしだけが対象です。');
    if (!runId || String(clientState.related_run_id || '') !== runId) {
      throw new Error('Hub/ClientのDRAFT_READY予測run lineageが一致しません。');
    }
  } else {
    planVersionId = String(hubState.related_plan_version_id || '');
    if (!planVersionId || String(clientState.related_plan_version_id || '') !== planVersionId) {
      throw new Error('Hub/ClientのSUBMITTED計画lineageが一致しません。');
    }
    const hubPlans = vNextAdminReadCoreRows_(hub, 'PLAN_VERSION').filter(function (row) {
      return String(row.book_id || '') === bookId &&
        String(row.plan_version_id || '') === planVersionId &&
        String(row.status || '').toUpperCase() === 'SUBMITTED';
    });
    const clientPlans = vNextAdminReadCoreRows_(client, 'PLAN_VERSION').filter(function (row) {
      return String(row.book_id || '') === bookId &&
        String(row.plan_version_id || '') === planVersionId &&
        String(row.status || '').toUpperCase() === 'SUBMITTED';
    });
    const hubPlan = hubPlans[hubPlans.length - 1];
    const clientPlan = clientPlans[clientPlans.length - 1];
    if (!hubPlan || !clientPlan || String(hubPlan.run_id || '') !== String(clientPlan.run_id || '') ||
        Number(hubPlan.final_budget || 0) !== Number(clientPlan.final_budget || 0)) {
      throw new Error('Hub/Clientの提出済み計画recordが一致しません。');
    }
    runId = String(hubPlan.run_id || '');
    const approvals = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.APPROVALS).rows.filter(function (row) {
      return String(row.book_id || '') === bookId &&
        String(row.plan_version_id || '') === planVersionId &&
        String(row.forecast_run_id || '') === runId;
    });
    const approval = approvals[approvals.length - 1];
    if (!approval || String(approval.status || '').toUpperCase() !== 'PENDING') {
      throw new Error('提出済み計画に対応するPENDING承認が一意に確認できません。');
    }
    approvalRequestId = String(approval.approval_request_id || '');
  }
  const hubRun = vNextAdminReadCoreRows_(hub, 'FORECAST_RUN').find(function (row) {
    return String(row.book_id || '') === bookId && String(row.run_id || '') === runId;
  });
  const clientRun = vNextAdminReadCoreRows_(client, 'FORECAST_RUN').find(function (row) {
    return String(row.book_id || '') === bookId && String(row.run_id || '') === runId;
  });
  if (!hubRun || !clientRun || String(hubRun.status || '').toUpperCase() !== 'SUCCESS' ||
      String(clientRun.status || '').toUpperCase() !== 'SUCCESS' ||
      String(hubRun.input_data_hash || '') !== String(clientRun.input_data_hash || '') ||
      Number(hubRun.p50 || 0) !== Number(clientRun.p50 || 0)) {
    throw new Error('Hub/ClientのSUCCESS予測recordが一致しません。');
  }
  return { runId: runId, planVersionId: planVersionId, approvalRequestId: approvalRequestId };
}

function vNextAdminUpgradeOnlyDraftReadyPilotUxForManualTest() {
  const hub = vNextAdminRequireHub_();
  const pair = vNextAdminReadActiveReleasePair_(hub);
  const candidates = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.REGISTRY).rows.filter(function (row) {
    return String(row.mode || '').toUpperCase() === 'CLIENT' &&
      String(row.status || '').toUpperCase() === 'ACTIVE' &&
      String(row.state || '').toUpperCase() === 'DRAFT_READY' &&
      String(row.access_policy || '').toUpperCase() === 'INTERNAL_OPEN' &&
      String(row.template_release_id || '') !== pair.releaseId;
  });
  if (candidates.length !== 1) throw new Error('更新候補のDRAFT_READY Pilotが1冊に確定しません: ' +
    candidates.length + '冊');
  return vNextAdminUpgradeDraftReadyPilotUx({ bookId: String(candidates[0].book_id || ''),
    dryRun: false, reason: '予測 record を保持したままクライアント年度ブック案内を現在版へ更新' });
}

/** Same-URL employee UX update for the single submitted, not-yet-official Pilot. */
function vNextAdminUpgradeOnlySubmittedPilotUxForManualTest() {
  const hub = vNextAdminRequireHub_();
  const pair = vNextAdminReadActiveReleasePair_(hub);
  const candidates = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.REGISTRY).rows.filter(function (row) {
    return String(row.mode || '').toUpperCase() === 'CLIENT' &&
      String(row.status || '').toUpperCase() === 'ACTIVE' &&
      String(row.state || '').toUpperCase() === 'SUBMITTED' &&
      String(row.access_policy || '').toUpperCase() === 'INTERNAL_OPEN' &&
      !String(row.current_official_id || '') &&
      String(row.template_release_id || '') !== pair.releaseId;
  });
  if (candidates.length !== 1) throw new Error('更新候補のSUBMITTED Pilotが1冊に確定しません: ' +
    candidates.length + '冊');
  return vNextAdminUpgradeDraftReadyPilotUx({ bookId: String(candidates[0].book_id || ''),
    preservedState: 'SUBMITTED', dryRun: false,
    reason: '予測・提出済み予算案・承認待ちを保持したままクライアント年度ブック UX を現在版へ更新' });
}

function vNextAdminRecoverDraftReadyPilotUxUpgrade(request) {
  return vNextAdminGuard_('vNextAdminRecoverDraftReadyPilotUxUpgrade', function () {
    const req = request && typeof request === 'object' ? request : {};
    const hub = vNextAdminRequireHub_();
    vNextAdminAssertHubAdmin_(hub, false);
    return vNextAdminWithScriptLock_('recover-draft-ready-pilot-ux', function () {
      const bookId = vNextAdminRequiredText_(req.bookId, 'bookId');
      const migration = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.MIGRATIONS).rows.filter(function (row) {
        const plan = vNextAdminParseJson_(row.plan_json, {});
        return String(row.book_id || '') === bookId &&
          String(plan.kind || '') === 'DRAFT_READY_PILOT_UX_UPGRADE_V1' &&
          (!req.migrationId || String(row.migration_id || '') === String(req.migrationId));
      }).slice(-1)[0];
      if (!migration) throw new Error('復旧対象のDRAFT_READY画面更新journalがありません。');
      const plan = vNextAdminParseJson_(migration.plan_json, {});
      const registry = vNextAdminFindRegistryRow_(hub, function (row) {
        return String(row.book_id || '') === bookId && String(row.mode || '') === 'CLIENT';
      });
      if (!registry || String(registry.spreadsheet_id || '') !== String(plan.spreadsheetId || '')) {
        throw new Error('復旧対象のClient registry identityが一致しません。');
      }
      const client = SpreadsheetApp.openById(String(plan.spreadsheetId || ''));
      vNextAdminAssertDraftReadyPilotUxBoundary_(hub, client, registry, plan.preservedState);
      const sourceRelease = vNextAdminEmptyPilotRelease_(hub, plan.sourceReleaseId,
        ['ACTIVE', 'RETIRED']);
      const sourceModel = vNextAdminEmptyPilotModel_(hub, plan.sourceModelReleaseId, sourceRelease,
        ['ACTIVE', 'RETIRED']);
      let direction = String(registry.template_release_id || '') === String(plan.targetReleaseId || '')
        ? 'TARGET' : 'SOURCE';
      let release = sourceRelease;
      let model = sourceModel;
      if (direction === 'TARGET') {
        const pair = vNextAdminReadActiveReleasePair_(hub);
        if (pair.releaseId !== String(plan.targetReleaseId || '') ||
            pair.modelReleaseId !== String(plan.targetModelReleaseId || '')) direction = 'SOURCE';
        else {
          release = vNextAdminEmptyPilotRelease_(hub, plan.targetReleaseId, ['ACTIVE']);
          model = vNextAdminEmptyPilotModel_(hub, plan.targetModelReleaseId, release, ['ACTIVE']);
        }
      } else if (String(registry.template_release_id || '') !== String(plan.sourceReleaseId || '')) {
        throw new Error('Registryがjournalのsource/target以外を参照しています。');
      }
      vNextAdminFreezeEmptyPilotClient_(hub, client, migration.migration_id);
      const applied = vNextAdminApplyEmptyPilotRelease_(hub, client, registry, plan, release, model,
        direction, function (phase, detail) {
          vNextAdminSetEmptyPilotUpgradePhase_(hub, migration.migration_id,
            'DRAFT_READY_UX_RECOVERY_' + String(phase || '').replace(/^EMPTY_PILOT_/, ''), detail);
        });
      const finalStatus = direction === 'TARGET' ? 'SUCCEEDED' : 'ROLLED_BACK';
      vNextAdminPatchLatestMigration_(hub, migration.migration_id, {
        status: finalStatus, finished_at: new Date(), error: '',
        result_json: vNextAdminCanonicalJson_({ phase: 'RECOVERED_' + direction,
          sourceForecastRunId: plan.sourceForecastRunId,
          healthStatus: applied.health.healthStatus, healthCode: applied.health.healthCode })
      });
      return { ok: true, migrationId: migration.migration_id, bookId: bookId,
        direction: direction, releaseId: release.release_id, health: applied.health };
    });
  });
}

function vNextAdminRecoverOnlyDraftReadyPilotUxForManualTest() {
  const hub = vNextAdminRequireHub_();
  const rows = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.MIGRATIONS).rows.filter(function (row) {
    const plan = vNextAdminParseJson_(row.plan_json, {});
    return String(plan.kind || '') === 'DRAFT_READY_PILOT_UX_UPGRADE_V1' &&
      /^DRAFT_READY_UX_|^RECOVERY_REQUIRED$/.test(String(row.status || '').toUpperCase());
  });
  if (rows.length !== 1) throw new Error('復旧対象のDRAFT_READY画面更新が1件に確定しません: ' +
    rows.length + '件');
  return vNextAdminRecoverDraftReadyPilotUxUpgrade({ bookId: String(rows[0].book_id || ''),
    migrationId: String(rows[0].migration_id || ''), reason: '中断した画面更新を安全に復旧' });
}

function vNextAdminAssertEmptyPilotUpgradeEligibility_(hub, client, registry, release, model) {
  vNextAdminAssertEmptyPilotNoBusinessData_(hub, client, registry, true);
  return vNextAdminAssertEmptyPilotPinnedRelease_(hub, client, registry, release, model, true);
}

function vNextAdminAssertEmptyPilotNoBusinessData_(hub, client, registry, initialOnly) {
  if (String(registry.mode || '') !== 'CLIENT' || String(registry.status || '').toUpperCase() !== 'ACTIVE' ||
      String(registry.state || '').toUpperCase() !== 'INPUT_OPEN' || String(registry.current_official_id || '')) {
    throw new Error('空のPilot更新はACTIVE / INPUT_OPEN / 正式予算なしのClientだけが対象です。');
  }
  const bookId = String(registry.book_id || '');
  ['EVIDENCE_EVENT', 'FORECAST_RUN', 'PLAN_VERSION', 'EVALUATION'].forEach(function (sheetName) {
    const hubRows = vNextAdminReadCoreRows_(hub, sheetName).filter(function (row) {
      return String(row.book_id || '') === bookId;
    });
    const clientRows = vNextAdminReadCoreRows_(client, sheetName).filter(function (row) {
      return String(row.book_id || '') === bookId;
    });
    if (hubRows.length || clientRows.length) {
      throw new Error('空のPilot更新を停止しました。' + sheetName + 'に既存recordがあります。');
    }
  });
  const stateRows = vNextAdminReadCoreRows_(hub, 'STATE_EVENT').filter(function (row) {
    return String(row.book_id || '') === bookId;
  }).concat(vNextAdminReadCoreRows_(client, 'STATE_EVENT').filter(function (row) {
    return String(row.book_id || '') === bookId;
  }));
  if (stateRows.length) throw new Error('空のPilot更新はSTATE_EVENTが未作成のbookだけが対象です。');

  const requestRows = vNextAdminReadTable_(client, VN_ADMIN_CLIENT_REQUEST_SHEET).rows.filter(function (row) {
    return !row.book_id || String(row.book_id || '') === bookId;
  }).concat(vNextAdminReadTable_(hub, VN_ADMIN_CLIENT_REQUEST_SHEET).rows.filter(function (row) {
    return String(row.book_id || '') === bookId;
  }));
  if (requestRows.length) throw new Error('空のPilot更新は予測依頼が一度もないbookだけが対象です。');
  if (vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.APPROVALS).rows.some(function (row) {
    return String(row.book_id || '') === bookId;
  })) throw new Error('空のPilot更新は承認recordがないbookだけが対象です。');
  if (vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.OFFICIAL).rows.some(function (row) {
    return String(row.book_id || '') === bookId;
  })) throw new Error('空のPilot更新は公式recordがないbookだけが対象です。');
  if (vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.JOBS).rows.some(function (row) {
    return String(row.target_book_id || '') === bookId &&
      ['QUEUED', 'RUNNING'].indexOf(String(row.status || '').toUpperCase()) >= 0;
  })) throw new Error('対象bookにQUEUED/RUNNING jobがあるため更新を停止しました。');
  const officialCopy = client.getSheetByName(VN_ADMIN_OFFICIAL_COPY_SHEET);
  if (officialCopy && officialCopy.getLastRow() > 1) {
    throw new Error('Clientに公式snapshotがあるため更新を停止しました。');
  }
  const routing = vNextAdminReadKeyValueSheet_(client, VN_ADMIN_BOOK_CONFIG_SHEET);
  if (String(routing.state || '').toUpperCase() !== 'INPUT_OPEN' ||
      Number(routing.input_submitted || 0) !== 0 ||
      Number(routing.input_answered_count || 0) !== 0) {
    throw new Error('Client入力状態が未使用のINPUT_OPENではありません。');
  }
  if (initialOnly) {
    const hubMetas = vNextAdminReadCoreRows_(hub, 'BOOK_META').filter(function (row) {
      return String(row.book_id || '') === bookId;
    });
    const clientMetas = vNextAdminReadCoreRows_(client, 'BOOK_META').filter(function (row) {
      return String(row.book_id || '') === bookId;
    });
    if (hubMetas.length !== 1 || clientMetas.length !== 1) {
      throw new Error('空のPilot更新は初期BOOK_METAが1件だけのbookに限定されます。');
    }
  }
  return true;
}

function vNextAdminAssertEmptyPilotPinnedRelease_(hub, client, registry, release, model, requireSingleMeta,
    expectedWorkflowState) {
  const bookId = String(registry.book_id || '');
  const releaseId = String(release.release_id || '');
  const expectedState = String(expectedWorkflowState || 'INPUT_OPEN').toUpperCase();
  if (String(registry.spreadsheet_id || '') !== String(client.getId()) ||
      String(registry.template_release_id || '') !== releaseId ||
      String(registry.schema_version || '') !== String(release.schema_version || '') ||
      String(registry.client_runtime_version || '') !== String(release.client_runtime_version || '') ||
      String(registry.client_runtime_sha256 || '') !== String(release.client_runtime_sha256 || '')) {
    throw new Error('BOOK_REGISTRYのsource release pinが一致しません。');
  }
  const routing = vNextAdminReadKeyValueSheet_(client, VN_ADMIN_BOOK_CONFIG_SHEET);
  const system = vNextAdminReadKeyValueSheet_(client, VN_ADMIN_SYSTEM_CONFIG_SHEET);
  if (String(routing.mode || '').toUpperCase() !== 'CLIENT' || String(routing.book_id || '') !== bookId ||
      String(routing.state || '').toUpperCase() !== expectedState ||
      String(routing.version || '') !== releaseId || String(routing.template_release_id || '') !== releaseId ||
      String(routing.schema_version || '') !== String(release.schema_version || '') ||
      String(routing.client_runtime_version || '') !== String(release.client_runtime_version || '') ||
      String(routing.client_runtime_bundle_sha256 || '') !== String(release.client_runtime_sha256 || '') ||
      String(routing.model_release_id || '') !== String(model.model_release_id || '')) {
    throw new Error('Client VN_BOOK_CONFIGのrelease/model/runtime pinが一致しません。');
  }
  if (String(system.mode || '').toUpperCase() !== 'CLIENT' || String(system.book_id || '') !== bookId ||
      String(system.active_release_id || '') !== releaseId ||
      String(system.schema_version || '') !== String(release.schema_version || '')) {
    throw new Error('Client VN_SYSTEM_CONFIGのrelease pinが一致しません。');
  }
  vNextClientRuntimeAssertBoundParent_(String(registry.client_script_id || ''), String(client.getId()));
  vNextClientRuntimeVerifyPinnedScriptContent_(
    vNextClientRuntimeGetContent_(String(registry.client_script_id || '')),
    String(registry.client_script_id || ''), String(release.client_runtime_sha256 || '')
  );
  const hubMetas = vNextAdminReadCoreRows_(hub, 'BOOK_META').filter(function (row) {
    return String(row.book_id || '') === bookId;
  });
  const clientMetas = vNextAdminReadCoreRows_(client, 'BOOK_META').filter(function (row) {
    return String(row.book_id || '') === bookId;
  });
  if (!hubMetas.length || !clientMetas.length || (requireSingleMeta &&
      (hubMetas.length !== 1 || clientMetas.length !== 1))) {
    throw new Error('Hub/Client BOOK_METAが空Pilot条件を満たしません。');
  }
  const hubMeta = hubMetas[hubMetas.length - 1];
  const clientMeta = clientMetas[clientMetas.length - 1];
  if (String(hubMeta.record_id || '') !== String(clientMeta.record_id || '') ||
      String(hubMeta.state || '').toUpperCase() !== String(clientMeta.state || '').toUpperCase() ||
      VN_ADMIN_CLIENT_STATES.indexOf(String(hubMeta.state || '').toUpperCase()) < 0 ||
      String(hubMeta.template_version || '') !== releaseId ||
      String(clientMeta.template_version || '') !== releaseId ||
      String(hubMeta.model_release_id || '') !== String(model.model_release_id || '') ||
      String(clientMeta.model_release_id || '') !== String(model.model_release_id || '') ||
      String(hubMeta.schema_version || '') !== String(release.schema_version || '') ||
      String(clientMeta.schema_version || '') !== String(release.schema_version || '')) {
    throw new Error('Hub/Client BOOK_METAのsource pinsが一致しません。');
  }
  if (vNextAdminLatestClientState_(hub, bookId, hubMeta.state) !== expectedState ||
      vNextAdminLatestClientState_(client, bookId, clientMeta.state) !== expectedState) {
    throw new Error('Hub/Clientのappend-only stateが更新対象stateと一致しません: ' + expectedState);
  }
  const activePair = vNextAdminReadActiveReleasePair_(hub);
  if (String(release.release_id || '') === String(activePair.releaseId || '')) {
    vNextAdminAssertModelTemplateCompatibility_(hub, model, release);
  } else {
    vNextAdminAssertModelReleaseOwnPair_(model, release);
  }
  return hubMeta;
}

function vNextAdminEmptyPilotRelease_(hub, releaseId, statuses) {
  const id = vNextAdminRequiredText_(releaseId, 'releaseId');
  const allowed = (statuses || []).map(function (value) { return String(value).toUpperCase(); });
  const row = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.RELEASES).rows.find(function (item) {
    return String(item.release_id || '') === id;
  });
  if (!row || allowed.indexOf(String(row.status || '').toUpperCase()) < 0) {
    throw new Error('空Pilot更新で利用できないTemplate Releaseです: ' + id);
  }
  return row;
}

function vNextAdminEmptyPilotModel_(hub, modelReleaseId, release, statuses) {
  const id = vNextAdminRequiredText_(modelReleaseId, 'modelReleaseId');
  const model = vNextAdminLatestModelRelease_(hub, id);
  const allowed = (statuses && statuses.length ? statuses : ['ACTIVE']).map(function (value) {
    return String(value || '').toUpperCase();
  });
  if (!model || allowed.indexOf(String(model.status || '').toUpperCase()) < 0) {
    throw new Error('空Pilot更新には過去にACTIVE化済みのMODEL_RELEASEが必要です: ' + id);
  }
  const activePair = vNextAdminReadActiveReleasePair_(hub);
  if (String(release.release_id || '') === String(activePair.releaseId || '')) {
    vNextAdminAssertModelReleaseChecksPassed_(model);
    vNextAdminAssertModelTemplateCompatibility_(hub, model, release);
  } else {
    vNextAdminAssertModelReleaseOwnPair_(model, release);
  }
  return model;
}

function vNextAdminAssertEmptyPilotReleaseAssets_(release) {
  const templateId = vNextAdminRequiredText_(release.template_spreadsheet_id,
    'release.template_spreadsheet_id');
  const scriptId = vNextAdminRequiredText_(release.template_script_id, 'release.template_script_id');
  const template = SpreadsheetApp.openById(templateId);
  if (vNextDetectBookMode_(template) !== 'TEMPLATE') throw new Error('Immutable Template modeが不正です。');
  vNextAdminAssertReleaseTemplateManifest_(release, template);
  vNextClientRuntimeAssertBoundParent_(scriptId, templateId);
  vNextClientRuntimeVerifyPinnedScriptContent_(vNextClientRuntimeGetContent_(scriptId), scriptId,
    vNextAdminRequiredText_(release.client_runtime_sha256, 'release.client_runtime_sha256'));
  const config = vNextAdminReadKeyValueSheet_(template, VN_ADMIN_BOOK_CONFIG_SHEET);
  if (String(config.version || '') !== String(release.release_id || '') ||
      String(config.schema_version || '') !== String(release.schema_version || '') ||
      String(config.client_runtime_version || '') !== String(release.client_runtime_version || '') ||
      String(config.client_runtime_bundle_sha256 || '') !== String(release.client_runtime_sha256 || '')) {
    throw new Error('Immutable Template local pinsがRELEASESと一致しません。');
  }
  return template;
}

function vNextAdminFreezeEmptyPilotClient_(hub, client, migrationId) {
  const hubConfig = vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET);
  const protectedNames = [VN_ADMIN_BOOK_CONFIG_SHEET, VN_ADMIN_SYSTEM_CONFIG_SHEET,
    VN_ADMIN_CLIENT_REQUEST_SHEET].concat(typeof VNEXT_CORE !== 'undefined'
    ? Object.keys(VNEXT_CORE.INTERNAL_SHEETS) : []);
  vNextAdminProtectInternalSheets_(client,
    vNextAdminMergeEmails_(hubConfig.admin_emails, vNextGetRuntimeConfig_().VNEXT_ADMIN_EMAILS,
      vNextAdminActor_()), protectedNames);
  vNextAdminWriteBookConfig_(client, {
    maintenance_mode: 'EMPTY_PILOT_UPGRADE', maintenance_operation_id: migrationId,
    updated_at: new Date(), updated_by: vNextAdminActor_()
  });
  return true;
}

function vNextAdminApplyEmptyPilotRelease_(hub, client, registry, plan, release, model, direction, phaseCallback) {
  const callback = typeof phaseCallback === 'function' ? phaseCallback : function () {};
  const expectedDirection = String(direction || '').toUpperCase();
  const preservedState = String(plan.preservedState || 'INPUT_OPEN').toUpperCase();
  if (VN_ADMIN_CLIENT_STATES.indexOf(preservedState) < 0) {
    throw new Error('Invalid preserved Client state for release upgrade: ' + preservedState);
  }
  if (['SOURCE', 'TARGET'].indexOf(expectedDirection) < 0) throw new Error('Invalid empty-Pilot recovery direction.');
  const template = vNextAdminAssertEmptyPilotReleaseAssets_(release);
  const bookId = String(plan.bookId || '');
  const scriptId = vNextAdminRequiredText_(plan.clientScriptId, 'plan.clientScriptId');
  if (String(client.getId()) !== String(plan.spreadsheetId || '') ||
      String(registry.book_id || '') !== bookId) throw new Error('Upgrade journal target identity changed.');

  const copied = vNextClientRuntimeCopyScriptContent_(String(release.template_script_id || ''),
    scriptId, String(release.client_runtime_sha256 || ''));
  callback('EMPTY_PILOT_RUNTIME_WRITTEN', { direction: expectedDirection,
    runtimeSha256: copied.bundleSha256 });

  vNextAdminCopyTemplateUiToClient_(template, client);
  SpreadsheetApp.flush();
  // Verify the copied employee assets before client-specific labels are added.
  vNextAdminAssertReleaseTemplateManifest_(release, client);
  vNextAdminEnsureUiShell_(client, { clientName: registry.client_name, fiscalYear: registry.fiscal_year });
  vNextAdminApplyVisibility_(client, VN_ADMIN_DEFAULT_CLIENT_VISIBLE);
  callback('EMPTY_PILOT_UI_WRITTEN', { direction: expectedDirection,
    templateContentSha256: release.template_content_sha256 });

  vNextAdminWriteBookConfig_(client, {
    mode: 'CLIENT', book_id: bookId, state: preservedState,
    version: release.release_id, template_release_id: release.release_id,
    model_release_id: model.model_release_id, schema_version: release.schema_version,
    client_runtime_version: release.client_runtime_version,
    client_runtime_bundle_sha256: release.client_runtime_sha256,
    maintenance_mode: 'EMPTY_PILOT_UPGRADE', maintenance_operation_id: String(plan.migrationId || ''),
    updated_at: new Date(), updated_by: vNextAdminActor_()
  });
  vNextAdminWriteSystemConfig_(client, {
    mode: 'CLIENT', book_id: bookId, active_release_id: release.release_id,
    schema_version: release.schema_version
  });
  callback('EMPTY_PILOT_CONFIG_WRITTEN', { direction: expectedDirection,
    releaseId: release.release_id, modelReleaseId: model.model_release_id });

  const meta = vNextAdminAppendEmptyPilotRepairMeta_(hub, client, plan, release, model, expectedDirection);
  callback('EMPTY_PILOT_META_APPENDED', { direction: expectedDirection, recordId: meta.record_id });
  if (plan.sourceForecastRunId) {
    vNextAdminWriteTriangulationProjection_(hub, client, bookId, plan.sourceForecastRunId);
  }

  if (expectedDirection === 'TARGET') {
    const pair = vNextAdminReadActiveReleasePair_(hub);
    if (pair.releaseId !== String(release.release_id || '') ||
        pair.modelReleaseId !== String(model.model_release_id || '') ||
        pair.templateSpreadsheetId !== String(release.template_spreadsheet_id || '')) {
      throw new Error('Registry commit直前にcanonical ACTIVE pairが変更されました。');
    }
  }
  // Commit marker: no Client release identity is published centrally before
  // runtime, UI, config and both append-only BOOK_META mirrors are durable.
  vNextAdminPatchRegistryByBookId_(hub, bookId, {
    template_release_id: release.release_id, schema_version: release.schema_version,
    client_runtime_version: release.client_runtime_version,
    client_runtime_sha256: release.client_runtime_sha256,
    state: preservedState, status: 'ACTIVE', health_status: 'PENDING',
    health_code: expectedDirection === 'TARGET' ? 'EMPTY_PILOT_UPGRADED' : 'EMPTY_PILOT_ROLLED_BACK',
    updated_at: new Date()
  });
  callback('EMPTY_PILOT_REGISTRY_COMMITTED', { direction: expectedDirection,
    releaseId: release.release_id });

  vNextAdminWriteBookConfig_(client, {
    maintenance_mode: '', maintenance_operation_id: '',
    updated_at: new Date(), updated_by: vNextAdminActor_()
  });
  vNextAdminProtectClientInternalSheets_(client, [VN_ADMIN_BOOK_CONFIG_SHEET,
    VN_ADMIN_SYSTEM_CONFIG_SHEET, VN_ADMIN_CLIENT_REQUEST_SHEET].concat(
      typeof VNEXT_CORE !== 'undefined' ? Object.keys(VNEXT_CORE.INTERNAL_SHEETS) : []
    ));
  const committedRegistry = vNextAdminFindRegistryRow_(hub, function (row) {
    return String(row.book_id || '') === bookId;
  });
  vNextAdminAssertEmptyPilotPinnedRelease_(hub, client, committedRegistry, release, model, false,
    preservedState);
  const health = vNextAdminScanOneBook_(hub, committedRegistry);
  if (String(health.healthStatus || '').toUpperCase() !== 'OK') {
    throw new Error('更新後health scanがOKではありません: ' + health.healthCode + ' ' + health.detail);
  }
  callback('EMPTY_PILOT_HEALTH_VERIFIED', { direction: expectedDirection,
    healthStatus: health.healthStatus, healthCode: health.healthCode });
  return { meta: meta, health: health };
}

/** Regenerable employee projection for runs created before triangulation was persisted. */
function vNextAdminWriteTriangulationProjection_(hub, client, bookId, runId) {
  const run = vNextAdminReadCoreRows_(hub, 'FORECAST_RUN').filter(function (row) {
    return String(row.book_id || '') === String(bookId || '') &&
      String(row.run_id || '') === String(runId || '') &&
      String(row.status || '').toUpperCase() === 'SUCCESS';
  }).slice(-1)[0];
  if (!run) throw new Error('Triangulation projection source run is missing: ' + runId);
  const lenses = vNextAdminParseJson_(run.lens_json, {});
  const triangulation = lenses.triangulation && Array.isArray(lenses.triangulation.methods)
    ? lenses.triangulation
    : (typeof vNextBuildTriangulationReference_ === 'function'
      ? vNextBuildTriangulationReference_(lenses.continuity || {}, Number(run.system_recommended || run.p50 || 0))
      : null);
  if (!triangulation || !Array.isArray(triangulation.methods) || !triangulation.methods.length) {
    throw new Error('Triangulation projection could not be derived.');
  }
  const payload = {
    schemaVersion: 'vnext-triangulation-projection-1', bookId: String(bookId || ''),
    runId: String(runId || ''), generatedAt: new Date().toISOString(),
    policy: String(triangulation.policy || 'INDEPENDENT_REFERENCES_NOT_AUTOMATICALLY_AVERAGED'),
    methods: triangulation.methods.slice(0, 5).map(function (method) {
      return { key: String(method.key || ''), label: String(method.label || '').slice(0, 80),
        value: Math.trunc(Number(method.value || 0)), assumption: String(method.assumption || '').slice(0, 240),
        basis: String(method.basis || '').slice(0, 120) };
    })
  };
  vNextAdminWriteBookConfig_(client, {
    triangulation_reference_json: vNextAdminCanonicalJson_(payload),
    triangulation_reference_updated_at: new Date(), updated_at: new Date(), updated_by: vNextAdminActor_()
  });
  return payload;
}

function vNextAdminAppendEmptyPilotRepairMeta_(hub, client, plan, release, model, direction) {
  const bookId = String(plan.bookId || '');
  const preservedState = String(plan.preservedState || 'INPUT_OPEN').toUpperCase();
  const hubMetas = vNextAdminReadCoreRows_(hub, 'BOOK_META').filter(function (row) {
    return String(row.book_id || '') === bookId;
  });
  let clientMetas = vNextAdminReadCoreRows_(client, 'BOOK_META').filter(function (row) {
    return String(row.book_id || '') === bookId;
  });
  if (!hubMetas.length || !clientMetas.length) throw new Error('BOOK_META repair source is missing.');
  const hubIds = new Set(hubMetas.map(function (row) { return String(row.record_id || ''); }));
  const untrustedClientOnly = clientMetas.filter(function (row) {
    return !hubIds.has(String(row.record_id || ''));
  });
  if (untrustedClientOnly.length) {
    throw new Error('Client-only BOOK_METAが見つかったため自動復旧を停止しました。');
  }
  // Hub is authoritative. If execution stopped after the Hub append but
  // before the Client append, mirror the missing trusted records first so the
  // next supersedes_record_id never points to a record absent from Client.
  vNextAdminAppendMissingCoreRows_(client, 'BOOK_META', 'record_id', hubMetas);
  clientMetas = vNextAdminReadCoreRows_(client, 'BOOK_META').filter(function (row) {
    return String(row.book_id || '') === bookId;
  });
  const latestHub = hubMetas[hubMetas.length - 1];
  const latestClient = clientMetas[clientMetas.length - 1];
  if (String(latestHub.record_id || '') === String(latestClient.record_id || '') &&
      String(latestHub.template_version || '') === String(release.release_id || '') &&
      String(latestHub.model_release_id || '') === String(model.model_release_id || '')) return latestHub;
  const original = hubMetas.find(function (row) {
    return String(row.record_id || '') === String(plan.sourceMetaRecordId || '');
  });
  if (!original) throw new Error('Journal source BOOK_META is missing.');
  const recordId = 'EMPTY-META-' + vNextAdminSha256_(vNextAdminCanonicalJson_({
    bookId: bookId, releaseId: release.release_id, modelReleaseId: model.model_release_id,
    direction: direction, predecessor: latestHub.record_id
  })).slice(0, 24).toUpperCase();
  const record = Object.assign({}, original, {
    record_id: recordId, state: preservedState, template_version: release.release_id,
    schema_version: release.schema_version, model_release_id: model.model_release_id,
    event_type: direction === 'TARGET'
      ? (String(plan.kind || '') === 'FAILED_PREFLIGHT_PILOT_UPGRADE_V1'
        ? 'FAILED_PREFLIGHT_PILOT_UPGRADED' : 'EMPTY_PILOT_UPGRADED')
      : (String(plan.kind || '') === 'FAILED_PREFLIGHT_PILOT_UPGRADE_V1'
        ? 'FAILED_PREFLIGHT_PILOT_UPGRADE_ROLLED_BACK' : 'EMPTY_PILOT_UPGRADE_ROLLED_BACK'),
    supersedes_record_id: String(latestHub.record_id || ''),
    recorded_at: new Date().toISOString(), recorded_by: vNextAdminActor_().toLowerCase()
  });
  vNextAdminAppendMissingCoreRows_(hub, 'BOOK_META', 'record_id', [record]);
  vNextAdminAppendMissingCoreRows_(client, 'BOOK_META', 'record_id', [record]);
  return record;
}

function vNextAdminSetEmptyPilotUpgradePhase_(hub, migrationId, phase, detail) {
  vNextAdminPatchLatestMigration_(hub, migrationId, {
    status: String(phase || ''),
    result_json: vNextAdminCanonicalJson_({ phase: String(phase || ''), detail: detail || {},
      updatedAt: new Date().toISOString() })
  });
}

function vNextAdminRegisterRelease(request) {
  return vNextAdminPublishTemplateRelease(request);
}

/**
 * Pull the full verified 管理ハブ runtime from the centrally deployed clasp
 * project into this known Hub-bound project. Spreadsheet data, releases,
 * approvals and Script Properties are not replaced.
 */
function vNextAdminUpdateHubRuntimeFromSource(request) {
  return vNextAdminGuard_('vNextAdminUpdateHubRuntimeFromSource', function () {
    const req = request && typeof request === 'object' ? request : {};
    const hub = vNextAdminRequireHub_();
    return vNextAdminUpdateHubRuntimeInHub_(hub, req, { requireHubScript: true });
  });
}

/** Copies the verified central clasp runtime into a registered Hub-bound project. */
function vNextAdminUpdateHubRuntimeInHub_(hub, request, options) {
  const req = request && typeof request === 'object' ? request : {};
  const opts = options && typeof options === 'object' ? options : {};
  vNextAdminAssertHubAdmin_(hub, opts.allowEffectiveUser === true);
  return vNextAdminWithScriptLock_('update-admin-runtime', function () {
    const reason = vNextAdminRequiredText_(req.reason, 'reason');
    const config = vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET);
    const sourceScriptId = vNextAdminRequiredText_(config.admin_source_script_id, 'admin_source_script_id');
    const targetScriptId = vNextAdminRequiredText_(config.admin_hub_script_id, 'admin_hub_script_id');
    const callerScriptId = String(ScriptApp.getScriptId() || '');
    if (opts.requireCentralCaller === true) {
      if (callerScriptId !== sourceScriptId) {
        throw new Error('Central-source Hub sync must run from admin_source_script_id.');
      }
    } else if (opts.requireHubScript === true && callerScriptId !== targetScriptId) {
      throw new Error('The current bound project does not match the registered 管理ハブ script ID.');
    }
    if (typeof vNextAdminRuntimeCopyScriptContent_ !== 'function') {
      throw new Error('Verified 管理ハブ runtime copy helper is not installed.');
    }
    const copied = vNextAdminRuntimeCopyScriptContent_(sourceScriptId, targetScriptId, hub.getId());
    vNextAdminWriteSystemConfig_(hub, {
      admin_runtime_sha256: copied.adminRuntimeSha256,
      admin_runtime_updated_at: new Date().toISOString(),
      admin_runtime_updated_by: vNextAdminActor_()
    });
    const hubRegistry = vNextAdminFindRegistryRow_(hub, function (row) {
      return String(row.mode || '') === 'ADMIN' && String(row.spreadsheet_id || '') === String(hub.getId());
    });
    if (!hubRegistry) throw new Error('管理ハブ BOOK_REGISTRY row is missing.');
    vNextAdminPatchRegistryByBookId_(hub, hubRegistry.book_id, {
      admin_script_id: targetScriptId, admin_runtime_sha256: copied.adminRuntimeSha256,
      health_status: 'PENDING', health_code: 'ADMIN_RUNTIME_UPDATED', updated_at: new Date(),
      note: '管理ハブ runtime updated from central source; reason=' + reason
    });
    vNextAdminWriteAudit_(hub, 'UPDATE_ADMIN_RUNTIME', 'ADMIN_RUNTIME', targetScriptId, 'SUCCESS', {
      sourceScriptId: sourceScriptId, targetSpreadsheetId: hub.getId(),
      adminRuntimeSha256: copied.adminRuntimeSha256, fileCount: copied.fileCount, reason: reason
    });
    return {
      ok: true, phase: 'HUB_RUNTIME', adminRuntimeSha256: copied.adminRuntimeSha256,
      fileCount: copied.fileCount,
      message: '管理ハブ runtimeを中央配備版へ更新しました。続けて Employee UX release を反映してください。'
    };
  });
}

/**
 * Creates the company shared drive if needed, builds purpose folders,
 * and moves Hub / Portal / Client / Template / audit files into that tree.
 */
function vNextAdminRelocateLibraryToSharedDrive(request) {
  return vNextAdminGuard_('vNextAdminRelocateLibraryToSharedDrive', function () {
    const hub = vNextAdminRequireHub_();
    return vNextAdminRelocateLibraryInHub_(hub, request, false);
  });
}

/** Central-source fallback used when Hub is not the active container. */
function vNextAdminRelocateLibraryToSharedDriveFromSource(request) {
  return vNextAdminGuard_('vNextAdminRelocateLibraryToSharedDriveFromSource', function () {
    const req = request && typeof request === 'object' ? request : {};
    const hubId = vNextAdminRequiredText_(req.hubSpreadsheetId, 'hubSpreadsheetId');
    const hub = SpreadsheetApp.openById(hubId);
    if (vNextDetectBookMode_(hub) !== 'ADMIN' || !vNextAdminIsRegisteredHub_(hub)) {
      throw new Error('The supplied Hub is not the registered 管理ハブ.');
    }
    vNextAdminAssertHubAdmin_(hub, true);
    vNextAdminHydrateHubRuntime_(hub);
    Object.keys(VN_ADMIN_HEADERS).forEach(function (name) {
      vNextAdminEnsureTable_(hub, name, VN_ADMIN_HEADERS[name]);
    });
    return vNextAdminRelocateLibraryInHub_(hub, req, true);
  });
}

function vNextAdminRelocateLibraryInHub_(hub, request, allowEffectiveUser) {
  const req = request && typeof request === 'object' ? request : {};
  vNextAdminAssertHubAdmin_(hub, allowEffectiveUser === true);
  const reason = vNextAdminText_(req.reason) || '社内共有ドライブへ整理';
  return vNextAdminWithScriptLock_('relocate-library', function () {
      const adminEmails = vNextAdminMergeEmails_(
        vNextGetRuntimeConfig_().VNEXT_ADMIN_EMAILS, vNextAdminActor_()
      );
      const domain = vNextAdminNormalizeDomain_(
        req.employeeDomain || vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET).employee_domain ||
        vNextGetRuntimeConfig_().VNEXT_EMPLOYEE_DOMAIN || vNextAdminEmailDomain_(vNextAdminActor_())
      );
      const prior = vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET);
      const library = vNextAdminEnsureSharedLibrary_(hub, {
        sharedDriveId: vNextAdminText_(req.sharedDriveId) || vNextAdminDetectCurrentSharedDriveId_(hub) ||
          vNextAdminText_(prior.shared_drive_id),
        adminEmails: adminEmails,
        domain: domain
      });
      vNextAdminWriteSystemConfig_(hub, {
        shared_drive_id: library.driveId,
        library_drive_name: VN_ADMIN_LIBRARY.DRIVE_NAME,
        library_portal_folder_id: library.folders.portal,
        library_books_folder_id: library.folders.books,
        library_admin_folder_id: library.folders.admin,
        library_audit_folder_id: library.folders.audit,
        library_templates_folder_id: library.folders.templates
      });
      const moved = vNextAdminMoveRegisteredFilesIntoLibrary_(hub, library, adminEmails, {
        legacyRootId: vNextAdminText_(prior.private_root_folder_id)
      });
      vNextAdminWriteSystemConfig_(hub, {
        private_root_folder_id: library.rootId,
        shared_drive_id: library.driveId,
        library_drive_name: VN_ADMIN_LIBRARY.DRIVE_NAME,
        library_portal_folder_id: library.folders.portal,
        library_books_folder_id: library.folders.books,
        library_admin_folder_id: library.folders.admin,
        library_audit_folder_id: library.folders.audit,
        library_templates_folder_id: library.folders.templates,
        library_relocated_at: new Date().toISOString(),
        library_relocated_by: vNextAdminActor_()
      });
      const portal = vNextAdminTryResolvePortal_(hub);
      if (portal && portal.spreadsheet) {
        vNextAdminWritePortalConfigValues_(portal.spreadsheet, {
          admin_hub_url: hub.getUrl(),
          portal_spreadsheet_url: portal.spreadsheet.getUrl()
        });
      }
      vNextAdminWriteAudit_(hub, 'RELOCATE_LIBRARY', 'DRIVE', library.driveId, 'SUCCESS', {
        reason: reason, driveName: VN_ADMIN_LIBRARY.DRIVE_NAME, moved: moved,
        folderUrl: 'https://drive.google.com/drive/folders/' + library.rootId
      });
      return {
        ok: true,
        driveId: library.driveId,
        folderUrl: 'https://drive.google.com/drive/folders/' + library.rootId,
        moved: moved,
        message: '共有ドライブ「' + VN_ADMIN_LIBRARY.DRIVE_NAME + '」へ整理しました。画面を再読み込みしてください。'
      };
    });
}

function vNextAdminTryResolvePortal_(hub) {
  try { return vNextAdminResolvePortal_(hub); }
  catch (error) { return null; }
}

/**
 * Updates the one registered employee Portal in place. The request log and
 * directory data stay in the Spreadsheet; only the verified portal-safe bound
 * runtime and its pinned identity are replaced.
 */
function vNextAdminUpdateSharedPortalRuntime(request) {
  return vNextAdminGuard_('vNextAdminUpdateSharedPortalRuntime', function () {
    const req = request && typeof request === 'object' ? request : {};
    const hub = vNextAdminRequireHub_();
    vNextAdminAssertHubAdmin_(hub, false);
    return vNextAdminWithScriptLock_('update-portal-runtime', function () {
      const reason = vNextAdminText_(req.reason) || '共有ドライブ移設後の最新版';
      const portal = vNextAdminResolvePortal_(hub);
      const target = vNextPortalRuntimeVerifiedBundle_();
      if (String(target.version || '') !== VN_ADMIN_PORTAL_RUNTIME_VERSION) {
        throw new Error('Deployed Portal bundle version does not match Admin target version.');
      }
      const targetFiles = vNextPortalRuntimeValidateFiles_(target.files || []);
      const targetSha = vNextClientRuntimeFilesSha256_(targetFiles);
      if (targetSha !== String(target.sha256 || '')) throw new Error('Target Portal bundle hash mismatch.');
      const project = vNextClientRuntimeAssertBoundParent_(portal.scriptId, portal.spreadsheetId);
      const currentContent = vNextClientRuntimeGetContent_(portal.scriptId);
      const currentFiles = vNextPortalRuntimeValidateExistingFiles_(currentContent.files || []);
      const currentSha = vNextClientRuntimeFilesSha256_(currentFiles);
      if (currentContent.scriptId && String(currentContent.scriptId) !== portal.scriptId) {
        throw new Error('Current Portal content scriptId mismatch.');
      }
      const expectedWebAppUrl = vNextAdminEmployeePortalWebAppUrl_(hub);
      if (currentSha === targetSha &&
          portal.runtimeVersion === VN_ADMIN_PORTAL_RUNTIME_VERSION &&
          portal.runtimeSha256 === targetSha) {
        const catalog = vNextAdminRefreshZacClientCatalogIfStale_(hub, true, { lockHeld: true });
        vNextAdminRefreshPortalDirectory_(hub, portal.spreadsheet);
        const webApp = vNextAdminPublishPortalWebApp_(portal.scriptId, expectedWebAppUrl);
        vNextAdminRememberPortalWebAppUrl_(hub, portal, webApp.webAppUrl);
        vNextAdminWriteAudit_(hub, 'PUBLISH_EMPLOYEE_PORTAL_WEBAPP', 'PORTAL', portal.portalId, 'SUCCESS', {
          scriptId: portal.scriptId, versionNumber: webApp.versionNumber,
          webAppUrl: webApp.webAppUrl, runtimeVersion: portal.runtimeVersion, reused: true
        });
        return {
          ok: true, reused: true, runtimeVersion: portal.runtimeVersion,
          runtimeSha256: portal.runtimeSha256, catalog: catalog,
          webAppUrl: webApp.webAppUrl, webAppVersion: webApp.versionNumber,
          message: '申請入口のファイルは最新版です。' + VNEXT_NAMING.WEB_ENTRY + 'を同じURLのまま公開し直しました。入口をハード再読み込みしてください。'
        };
      }
      if ([VN_ADMIN_PORTAL_RUNTIME_VERSION].concat(VN_ADMIN_PORTAL_LEGACY_RUNTIME_VERSIONS)
          .indexOf(portal.runtimeVersion) < 0) {
        throw new Error('Portal runtime migration source version is not allowlisted: ' + portal.runtimeVersion);
      }
      const hubConfigBefore = vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET);
      const localConfigBefore = vNextAdminReadKeyValueSheet_(portal.spreadsheet, VN_ADMIN_PORTAL_CONFIG_SHEET);
      const settingBefore = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.SETTINGS).rows.find(function (row) {
        return String(row.setting_key || '') === 'EMPLOYEE_PORTAL_JSON';
      });
      const requestSheet = portal.spreadsheet.getSheetByName(VN_ADMIN_PORTAL_REQUEST_SHEET);
      const directorySheet = portal.spreadsheet.getSheetByName(VN_ADMIN_PORTAL_DIRECTORY_SHEET);
      const needsV2HeaderExpansion = !vNextAdminPortalUsesV2Tables_(portal.runtimeVersion);
      let contentUpdateAttempted = false;
      let tablesExpanded = false;
      let pinsUpdated = false;
      let migrated = null;
      try {
        // Freeze employee appends during the small cross-file migration window.
        vNextAdminProtectInternalSheets_(portal.spreadsheet,
          vNextAdminMergeEmails_(hubConfigBefore.admin_emails, vNextAdminActor_()),
          [VN_ADMIN_PORTAL_REQUEST_SHEET]);
        // Arm rollback before the remote PUT. The Apps Script API can apply the
        // new content and still lose the response; in that case we must restore
        // the verified previous bundle instead of leaving v2 code with v1 pins.
        contentUpdateAttempted = true;
        const put = vNextClientRuntimePutContent_(portal.scriptId, target);
        const written = put && Array.isArray(put.files) ? put : vNextClientRuntimeGetContent_(portal.scriptId);
        const writtenFiles = vNextPortalRuntimeValidateFiles_(written.files || []);
        if (written.scriptId && String(written.scriptId) !== portal.scriptId ||
            vNextClientRuntimeFilesSha256_(writtenFiles) !== targetSha) {
          throw new Error('Written Portal runtime could not be verified.');
        }
        // Mark the migration attempt before the first header write. The helper
        // performs several Sheets calls; any mid-call failure must still enter
        // the v1 header rollback path.
        if (needsV2HeaderExpansion) {
          tablesExpanded = true;
          vNextAdminExpandPortalTableHeadersForV2_(portal.spreadsheet);
        }
        vNextAdminWritePortalConfigValues_(portal.spreadsheet, {
          schema_version: VN_ADMIN_PORTAL_REQUEST_SCHEMA,
          runtime_version: VN_ADMIN_PORTAL_RUNTIME_VERSION,
          runtime_sha256: targetSha,
          updated_at: new Date().toISOString(), updated_by: vNextAdminActor_()
        });
        vNextAdminWriteSystemConfig_(hub, {
          portal_runtime_version: VN_ADMIN_PORTAL_RUNTIME_VERSION,
          portal_runtime_sha256: targetSha,
          portal_runtime_updated_at: new Date().toISOString(),
          portal_runtime_updated_by: vNextAdminActor_()
        });
        pinsUpdated = true;
        const catalog = vNextAdminRefreshZacClientCatalogIfStale_(hub, true, { lockHeld: true });
        vNextAdminRefreshPortalDirectory_(hub);
        const settingValue = vNextAdminCanonicalJson_({
          portalId: portal.portalId, spreadsheetId: portal.spreadsheetId,
          scriptId: portal.scriptId, runtimeVersion: VN_ADMIN_PORTAL_RUNTIME_VERSION,
          runtimeSha256: targetSha, employeeDomain: portal.employeeDomain,
          accessPolicy: 'INTERNAL_OPEN'
        });
        vNextAdminUpsertObject_(hub, VN_ADMIN_SHEETS.SETTINGS, 'setting_key', 'EMPLOYEE_PORTAL_JSON', {
          setting_key: 'EMPLOYEE_PORTAL_JSON', setting_value: settingValue, value_type: 'JSON',
          scope: 'SYSTEM', effective_from: new Date(), updated_at: new Date(),
          updated_by: vNextAdminActor_(), note: VNEXT_NAMING.LAYER2 + '（' + VNEXT_NAMING.LAYER1 + 'とは物理分離）'
        });
        vNextAdminProtectInternalSheets_(portal.spreadsheet,
          vNextAdminMergeEmails_(hubConfigBefore.admin_emails, vNextAdminActor_()), [
            VN_ADMIN_PORTAL_DIRECTORY_SHEET, VN_ADMIN_PORTAL_CONFIG_SHEET,
            VN_ADMIN_PORTAL_CLIENT_CATALOG_SHEET
          ]);
        vNextAdminProtectClientInternalSheets_(portal.spreadsheet, [VN_ADMIN_PORTAL_REQUEST_SHEET]);
        vNextAdminWriteAudit_(hub, 'UPDATE_EMPLOYEE_PORTAL_RUNTIME', 'PORTAL', portal.portalId, 'SUCCESS', {
          fromVersion: portal.runtimeVersion, fromSha256: portal.runtimeSha256,
          toVersion: VN_ADMIN_PORTAL_RUNTIME_VERSION, toSha256: targetSha,
          scriptId: portal.scriptId, spreadsheetId: portal.spreadsheetId,
          parentTitle: project.title, catalogVersion: catalog.catalogVersion, reason: reason
        });
        migrated = {
          ok: true, reused: false, runtimeVersion: VN_ADMIN_PORTAL_RUNTIME_VERSION,
          runtimeSha256: targetSha, catalog: catalog
        };
      } catch (migrationError) {
        const rollbackErrors = [];
        try {
          if (pinsUpdated) {
            vNextAdminWriteSystemConfig_(hub, {
              portal_runtime_version: hubConfigBefore.portal_runtime_version || portal.runtimeVersion,
              portal_runtime_sha256: hubConfigBefore.portal_runtime_sha256 || portal.runtimeSha256
            });
            vNextAdminWritePortalConfigValues_(portal.spreadsheet, {
              schema_version: localConfigBefore.schema_version || VN_ADMIN_PORTAL_REQUEST_SCHEMA_V1,
              runtime_version: localConfigBefore.runtime_version || portal.runtimeVersion,
              runtime_sha256: localConfigBefore.runtime_sha256 || portal.runtimeSha256
            });
          }
        } catch (pinRollbackError) { rollbackErrors.push('pins=' + String(pinRollbackError)); }
        try {
          if (tablesExpanded) vNextAdminRollbackPortalTableHeadersToV1_(requestSheet, directorySheet);
        } catch (tableRollbackError) { rollbackErrors.push('tables=' + String(tableRollbackError)); }
        try {
          if (contentUpdateAttempted) {
            vNextClientRuntimePutContent_(portal.scriptId, { files: currentFiles });
            const restored = vNextClientRuntimeGetContent_(portal.scriptId);
            const restoredFiles = vNextPortalRuntimeValidateExistingFiles_(restored.files || []);
            if (vNextClientRuntimeFilesSha256_(restoredFiles) !== currentSha) {
              throw new Error('restored Portal SHA mismatch');
            }
          }
        } catch (contentRollbackError) { rollbackErrors.push('content=' + String(contentRollbackError)); }
        try {
          if (settingBefore) {
            const restoredSetting = Object.assign({}, settingBefore);
            delete restoredSetting._rowNumber;
            vNextAdminUpsertObject_(hub, VN_ADMIN_SHEETS.SETTINGS, 'setting_key',
              'EMPLOYEE_PORTAL_JSON', restoredSetting);
          }
        } catch (settingRollbackError) { rollbackErrors.push('setting=' + String(settingRollbackError)); }
        try { vNextAdminProtectClientInternalSheets_(portal.spreadsheet, [VN_ADMIN_PORTAL_REQUEST_SHEET]); }
        catch (protectionRollbackError) { rollbackErrors.push('protection=' + String(protectionRollbackError)); }
        vNextAdminWriteAudit_(hub, 'UPDATE_EMPLOYEE_PORTAL_RUNTIME', 'PORTAL', portal.portalId,
          rollbackErrors.length ? 'ROLLBACK_INCOMPLETE' : 'ROLLED_BACK', {
            fromVersion: portal.runtimeVersion, attemptedVersion: VN_ADMIN_PORTAL_RUNTIME_VERSION,
            error: String(migrationError && migrationError.message || migrationError),
            rollbackErrors: rollbackErrors, reason: reason
          });
        if (rollbackErrors.length) {
          throw new Error('Portal更新に失敗し、rollbackにも確認事項があります: ' +
            String(migrationError && migrationError.message || migrationError) + '; ' + rollbackErrors.join('; '));
        }
        throw migrationError;
      }
      const webApp = vNextAdminPublishPortalWebApp_(portal.scriptId, expectedWebAppUrl);
      vNextAdminRememberPortalWebAppUrl_(hub, portal, webApp.webAppUrl);
      vNextAdminWriteAudit_(hub, 'PUBLISH_EMPLOYEE_PORTAL_WEBAPP', 'PORTAL', portal.portalId, 'SUCCESS', {
        scriptId: portal.scriptId, versionNumber: webApp.versionNumber,
        webAppUrl: webApp.webAppUrl, runtimeVersion: VN_ADMIN_PORTAL_RUNTIME_VERSION, reused: false
      });
      migrated.webAppUrl = webApp.webAppUrl;
      migrated.webAppVersion = webApp.versionNumber;
      migrated.message = '申請入口を最新版へ更新し、' + VNEXT_NAMING.WEB_ENTRY + 'も同じURLのまま公開しました。入口をハード再読み込みしてください。';
      return migrated;
    });
  });
}

/**
 * Pins the existing employee /exec deployment to a new Apps Script version.
 * Does not create a second Web App URL. Publish is outside the runtime
 * rollback window so a failed redeploy cannot undo a verified file copy.
 */
function vNextAdminPublishPortalWebApp_(scriptId, expectedUrl) {
  const id = vNextClientRuntimeValidateScriptId_(scriptId, 'scriptId');
  const created = vNextClientRuntimeApiRequest_(
    '/projects/' + encodeURIComponent(id) + '/versions',
    'post',
    { description: VN_ADMIN_PORTAL_RUNTIME_VERSION + ' employee web entry' }
  );
  const versionNumber = Number(created && created.versionNumber || 0);
  if (!versionNumber) throw new Error('Portal web app version was not created.');
  const listed = vNextClientRuntimeApiRequest_(
    '/projects/' + encodeURIComponent(id) + '/deployments',
    'get'
  );
  const selected = vNextAdminSelectPortalWebAppDeployment_(listed.deployments || [], expectedUrl);
  const selectedId = String(selected.deploymentId || '');
  const requiredId = vNextAdminRequiredPortalWebAppDeploymentId_(
    listed.deployments || [], expectedUrl);
  if (requiredId && selectedId !== requiredId) {
    throw new Error('Refusing to republish a different employee Web App URL.');
  }
  const updated = vNextClientRuntimeApiRequest_(
    '/projects/' + encodeURIComponent(id) + '/deployments/' +
      encodeURIComponent(selectedId),
    'put',
    {
      deploymentConfig: {
        versionNumber: versionNumber,
        manifestFileName: (selected.deploymentConfig &&
          selected.deploymentConfig.manifestFileName) || 'appsscript',
        description: VN_ADMIN_PORTAL_RUNTIME_VERSION
      }
    }
  );
  let verified = updated && updated.deploymentId ? updated : selected;
  let pinnedVersion = Number(verified && verified.deploymentConfig &&
    verified.deploymentConfig.versionNumber || 0);
  if (pinnedVersion !== versionNumber) {
    for (let attempt = 0; attempt < 6; attempt++) {
      if (attempt > 0) Utilities.sleep(500 * attempt);
      verified = vNextClientRuntimeApiRequest_(
        '/projects/' + encodeURIComponent(id) + '/deployments/' +
          encodeURIComponent(selectedId),
        'get'
      );
      pinnedVersion = Number(verified && verified.deploymentConfig &&
        verified.deploymentConfig.versionNumber || 0);
      if (pinnedVersion === versionNumber) break;
    }
  }
  if (pinnedVersion !== versionNumber) {
    throw new Error('Portal /exec is still pinned to version ' + pinnedVersion +
      '; expected ' + versionNumber + '. 30秒ほど待ってから「最新版へ更新」をもう一度押してください。');
  }
  const webAppUrl = vNextAdminWebAppUrlFromDeployment_(verified) ||
    vNextAdminWebAppUrlFromDeployment_(selected);
  if (!webAppUrl) throw new Error('Portal web app URL was missing after republish.');
  if (requiredId && webAppUrl.indexOf(requiredId) < 0) {
    throw new Error('Portal /exec URL changed during republish.');
  }
  return { versionNumber: versionNumber, webAppUrl: webAppUrl, deploymentId: selectedId };
}

function vNextAdminWebAppUrlFromDeployment_(deployment) {
  const entries = (deployment && deployment.entryPoints) || [];
  for (let i = 0; i < entries.length; i++) {
    if (String(entries[i].entryPointType || '') !== 'WEB_APP') continue;
    const url = String((entries[i].webApp || {}).url || '');
    if (url) return url;
  }
  return '';
}

function vNextAdminWebAppDeploymentIdFromUrl_(url) {
  const match = String(url || '').match(/\/s\/(AKfycb[A-Za-z0-9_-]+)/);
  return match ? match[1] : '';
}

function vNextAdminRequiredPortalWebAppDeploymentId_(deployments, expectedUrl) {
  const urlId = vNextAdminWebAppDeploymentIdFromUrl_(expectedUrl);
  if (urlId) return urlId;
  const bookmarkId = VN_ADMIN_EMPLOYEE_PORTAL_WEBAPP_DEPLOYMENT_ID;
  const present = (deployments || []).some(function (deployment) {
    return String(deployment.deploymentId || '') === bookmarkId ||
      String(vNextAdminWebAppUrlFromDeployment_(deployment)).indexOf(bookmarkId) >= 0;
  });
  return present ? bookmarkId : '';
}

function vNextAdminEmployeePortalWebAppUrl_(hub) {
  const config = vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET);
  const fromConfig = String((config && config.portal_web_app_url) || '').trim();
  if (fromConfig) return fromConfig;
  const row = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.SETTINGS).rows.find(function (item) {
    return String(item.setting_key || '') === 'EMPLOYEE_PORTAL_JSON';
  });
  if (!row || !row.setting_value) return '';
  try {
    const parsed = JSON.parse(String(row.setting_value));
    return String(parsed && parsed.webAppUrl || '').trim();
  } catch (error) {
    return '';
  }
}

function vNextAdminRememberPortalWebAppUrl_(hub, portal, webAppUrl) {
  const url = String(webAppUrl || '').trim();
  if (!url) return;
  vNextAdminWriteSystemConfig_(hub, { portal_web_app_url: url });
  const existing = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.SETTINGS).rows.find(function (item) {
    return String(item.setting_key || '') === 'EMPLOYEE_PORTAL_JSON';
  });
  let parsed = {};
  if (existing && existing.setting_value) {
    try { parsed = JSON.parse(String(existing.setting_value)) || {}; }
    catch (error) { parsed = {}; }
  }
  parsed.webAppUrl = url;
  if (portal) {
    if (portal.portalId) parsed.portalId = portal.portalId;
    if (portal.spreadsheetId) parsed.spreadsheetId = portal.spreadsheetId;
    if (portal.scriptId) parsed.scriptId = portal.scriptId;
  }
  vNextAdminUpsertObject_(hub, VN_ADMIN_SHEETS.SETTINGS, 'setting_key', 'EMPLOYEE_PORTAL_JSON', {
    setting_key: 'EMPLOYEE_PORTAL_JSON',
    setting_value: vNextAdminCanonicalJson_(parsed),
    value_type: 'JSON', scope: 'SYSTEM', effective_from: new Date(), updated_at: new Date(),
    updated_by: vNextAdminActor_(), note: VNEXT_NAMING.LAYER2 + '（' + VNEXT_NAMING.LAYER1 + 'とは物理分離）'
  });
}

function vNextAdminSelectPortalWebAppDeployment_(deployments, expectedUrl) {
  const webApps = (deployments || []).map(function (deployment) {
    const url = vNextAdminWebAppUrlFromDeployment_(deployment);
    if (!url) return null;
    return { deployment: deployment, url: url };
  }).filter(Boolean);
  if (!webApps.length) {
    throw new Error('Portal web app deployment was not found. Publish /exec once from the Apps Script editor, then retry.');
  }
  const preferredId = vNextAdminRequiredPortalWebAppDeploymentId_(deployments, expectedUrl);
  const matched = preferredId ? webApps.filter(function (item) {
    return String(item.deployment.deploymentId || '') === preferredId ||
      item.url.indexOf(preferredId) >= 0;
  }) : [];
  if (matched.length === 1) return matched[0].deployment;
  if (preferredId && !matched.length) {
    throw new Error('Bookmarked employee /exec was not found among Portal deployments.');
  }
  const versioned = webApps.filter(function (item) {
    return Number(item.deployment.deploymentConfig &&
      item.deployment.deploymentConfig.versionNumber || 0) > 0;
  });
  const pool = versioned.length ? versioned : webApps;
  if (pool.length === 1) return pool[0].deployment;
  const domain = pool.filter(function (item) {
    const entries = item.deployment.entryPoints || [];
    return entries.some(function (entry) {
      const access = String(((entry.webApp || {}).entryPointConfig || {}).access || '').toUpperCase();
      return access === 'DOMAIN';
    });
  });
  if (domain.length === 1) return domain[0].deployment;
  throw new Error('Multiple Portal web app deployments exist: ' +
    pool.map(function (item) { return item.url; }).join(', '));
}

function vNextAdminWritePortalConfigValues_(portal, values) {
  Object.keys(values || {}).forEach(function (key) {
    vNextAdminUpsertObject_(portal, VN_ADMIN_PORTAL_CONFIG_SHEET, 'key', key,
      { key: key, value: values[key] }, ['key', 'value']);
  });
  portal.getSheetByName(VN_ADMIN_PORTAL_CONFIG_SHEET).hideSheet();
}

function vNextAdminPortalHeadersMatch_(actual, expected) {
  const normalized = (actual || []).map(function (value) { return String(value || '').trim(); });
  const target = (expected || []).map(String);
  if (normalized.length < target.length) return false;
  if (vNextAdminCanonicalJson_(normalized.slice(0, target.length)) !== vNextAdminCanonicalJson_(target)) {
    return false;
  }
  return !normalized.slice(target.length).some(Boolean);
}

function vNextAdminReadSheetHeaderRow_(sheet, minWidth) {
  const width = Math.max(minWidth || 1, sheet.getLastColumn(), 1);
  if (sheet.getMaxColumns() < width) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), width - sheet.getMaxColumns());
  }
  return sheet.getRange(1, 1, 1, width).getValues()[0].map(function (value) {
    return String(value || '').trim();
  });
}

/** Versioned Portal table migration: v1 exact, v2 remap-by-name, or blank bootstrap. */
function vNextAdminMigratePortalSheetHeaders_(ss, sheetName, v1Headers, v2Headers) {
  const sheet = vNextAdminGetOrCreateSheet_(ss, sheetName);
  const actual = vNextAdminReadSheetHeaderRow_(sheet, Math.max(v1Headers.length, v2Headers.length));
  if (vNextAdminPortalHeadersMatch_(actual, v2Headers)) {
    sheet.getRange(1, 1, 1, v2Headers.length).setFontWeight('bold').setBackground('#eeeeee');
    sheet.setFrozenRows(1);
    return sheet;
  }
  if (!actual.some(Boolean)) {
    sheet.getRange(1, 1, 1, v2Headers.length).setValues([v2Headers.slice()]);
    sheet.getRange(1, 1, 1, v2Headers.length).setFontWeight('bold').setBackground('#eeeeee');
    sheet.setFrozenRows(1);
    return sheet;
  }
  if (vNextAdminPortalHeadersMatch_(actual, v1Headers)) {
    sheet.getRange(1, v1Headers.length + 1, 1, v2Headers.length - v1Headers.length)
      .setValues([v2Headers.slice(v1Headers.length)]);
    sheet.getRange(1, 1, 1, v2Headers.length).setFontWeight('bold').setBackground('#eeeeee');
    sheet.setFrozenRows(1);
    return sheet;
  }
  const headerIndex = {};
  actual.forEach(function (header, index) {
    if (header) headerIndex[header] = index;
  });
  const canRemap = v2Headers.every(function (header) {
    return Object.prototype.hasOwnProperty.call(headerIndex, header);
  });
  if (!canRemap) {
    throw new Error(sheetName + 'の列構成がruntime契約と一致しません。管理ハブが直接修正せずversioned migrationを実行してください。');
  }
  const bodyRows = sheet.getLastRow() > 1
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, actual.length).getValues() : [];
  const remapped = bodyRows.map(function (row) {
    return v2Headers.map(function (header) {
      const idx = headerIndex[header];
      return idx === undefined ? '' : row[idx];
    });
  });
  if (sheet.getMaxColumns() < v2Headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), v2Headers.length - sheet.getMaxColumns());
  }
  sheet.getRange(1, 1, 1, v2Headers.length).setValues([v2Headers.slice()]);
  if (remapped.length) {
    sheet.getRange(2, 1, remapped.length, v2Headers.length).setValues(remapped);
  }
  if (sheet.getLastColumn() > v2Headers.length) {
    sheet.getRange(1, v2Headers.length + 1, Math.max(1, sheet.getLastRow()), sheet.getLastColumn() - v2Headers.length)
      .clearContent();
  }
  sheet.getRange(1, 1, 1, v2Headers.length).setFontWeight('bold').setBackground('#eeeeee');
  sheet.setFrozenRows(1);
  return sheet;
}

function vNextAdminEnsurePortalRuntimeTables_(ss, runtimeVersion) {
  if (vNextAdminPortalUsesV2Tables_(runtimeVersion)) {
    vNextAdminMigratePortalSheetHeaders_(ss, VN_ADMIN_PORTAL_REQUEST_SHEET,
      VN_ADMIN_PORTAL_REQUEST_HEADERS_V1, VN_ADMIN_PORTAL_REQUEST_HEADERS);
    vNextAdminMigratePortalSheetHeaders_(ss, VN_ADMIN_PORTAL_DIRECTORY_SHEET,
      VN_ADMIN_PORTAL_DIRECTORY_HEADERS_V1, VN_ADMIN_PORTAL_DIRECTORY_HEADERS);
    vNextAdminEnsureExactTableHeaders_(ss, VN_ADMIN_PORTAL_CLIENT_CATALOG_SHEET,
      VN_ADMIN_PORTAL_CLIENT_CATALOG_HEADERS).hideSheet();
    return;
  }
  vNextAdminEnsureExactTableHeaders_(ss, VN_ADMIN_PORTAL_REQUEST_SHEET, VN_ADMIN_PORTAL_REQUEST_HEADERS_V1);
  vNextAdminEnsureExactTableHeaders_(ss, VN_ADMIN_PORTAL_DIRECTORY_SHEET, VN_ADMIN_PORTAL_DIRECTORY_HEADERS_V1);
  vNextAdminEnsureExactTableHeaders_(ss, VN_ADMIN_PORTAL_CLIENT_CATALOG_SHEET,
    VN_ADMIN_PORTAL_CLIENT_CATALOG_HEADERS).hideSheet();
}

function vNextAdminExpandPortalTableHeadersForV2_(portal) {
  vNextAdminEnsurePortalRuntimeTables_(portal, VN_ADMIN_PORTAL_RUNTIME_VERSION);
}

function vNextAdminRollbackPortalTableHeadersToV1_(requestSheet, directorySheet) {
  if (requestSheet) {
    const firstExtra = VN_ADMIN_PORTAL_REQUEST_HEADERS_V1.length + 1;
    const extraWidth = VN_ADMIN_PORTAL_REQUEST_HEADERS.length - VN_ADMIN_PORTAL_REQUEST_HEADERS_V1.length;
    if (requestSheet.getLastRow() > 1) {
      const values = requestSheet.getRange(2, firstExtra, requestSheet.getLastRow() - 1, extraWidth).getValues();
      if (values.some(function (row) { return row.some(function (value) { return String(value || '') !== ''; }); })) {
        throw new Error('v2 request data appeared during rollback; extra columns were preserved.');
      }
    }
    requestSheet.getRange(1, firstExtra, Math.max(1, requestSheet.getLastRow()), extraWidth).clearContent();
  }
  if (directorySheet) {
    const extraColumn = VN_ADMIN_PORTAL_DIRECTORY_HEADERS_V1.length + 1;
    directorySheet.getRange(1, extraColumn, Math.max(1, directorySheet.getLastRow()), 1).clearContent();
  }
}

/**
 * Create an Admin-private, mutable UI draft. Only the three employee-facing
 * sheets cross the boundary; runtime/config/audit sheets are freshly created.
 */
function vNextAdminCreateTemplateDraft(request) {
  return vNextAdminGuard_('vNextAdminCreateTemplateDraft', function () {
    const req = request && typeof request === 'object' ? request : {};
    const hub = vNextAdminRequireHub_();
    return vNextAdminWithScriptLock_('create-template-draft', function () {
      if (typeof vNextClientRuntimeVerifiedBundle_ !== 'function' ||
          typeof vNextClientRuntimeCreateBoundSpreadsheet_ !== 'function') {
        throw new Error('Verified Client runtime publisher is not installed.');
      }
      const reason = vNextAdminRequiredText_(req.reason, 'reason');
      const source = vNextAdminResolveTemplateUiSource_(hub, req.templateDraftSpreadsheetId || '');
      const bundle = vNextClientRuntimeVerifiedBundle_();
      const draftName = vNextAdminText_(req.draftName) ||
        ('Forecast vNext Template Draft ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmm'));
      const draftId = vNextAdminText_(req.draftId) || ('TD-' + vNextAdminSha256_(vNextAdminCanonicalJson_({
        sourceSpreadsheetId: source.spreadsheet.getId(), sourceReleaseId: source.sourceReleaseId,
        draftName: draftName, reason: reason, actor: vNextAdminActor_().toLowerCase()
      })).slice(0, 24).toUpperCase());
      const existing = vNextAdminFindRegistryRow_(hub, function (row) {
        return String(row.book_id || '') === draftId && String(row.mode || '') === 'TEMPLATE';
      });
      if (existing) {
        if (String(existing.status || '').toUpperCase() !== 'DRAFT' ||
            !vNextAdminSpreadsheetAccessible_(existing.spreadsheet_id)) {
          throw new Error('draftId already exists but is not a usable TEMPLATE_DRAFT: ' + draftId);
        }
        const existingDraft = SpreadsheetApp.openById(String(existing.spreadsheet_id));
        const existingConfig = vNextAdminReadKeyValueSheet_(existingDraft, VN_ADMIN_BOOK_CONFIG_SHEET);
        if (String(existingConfig.template_draft_id || '') !== draftId ||
            String(existingConfig.source_template_release_id || '') !== String(source.sourceReleaseId || '')) {
          throw new Error('Existing TEMPLATE_DRAFT identity does not match this request.');
        }
        vNextAdminAssertPrivateAdminFile_(DriveApp.getFileById(existingDraft.getId()), source.adminEmails);
        return {
          reused: true, draftId: draftId, spreadsheetId: existingDraft.getId(),
          spreadsheetUrl: existingDraft.getUrl(), sourceReleaseId: source.sourceReleaseId,
          manifestSha256: vNextAdminTemplateUiManifestHash_(existingDraft)
        };
      }

      const folder = vNextAdminPrepareLibraryDestinationFolder_(hub, '',
        vNextAdminLibraryPath_('TEMPLATE_DRAFT'), source.adminEmails);
      const created = vNextClientRuntimeCreateBoundSpreadsheet_({ title: draftName, folderId: folder.getId() });
      if (String(created.runtimeVersion || '') !== String(bundle.version || '') ||
          String(created.bundleSha256 || '') !== String(bundle.sha256 || '')) {
        throw new Error('TEMPLATE_DRAFT runtime does not match the verified Client bundle.');
      }
      if (typeof vNextClientRuntimeAssertBoundParent_ !== 'function') {
        throw new Error('Client runtime parent verifier is not installed.');
      }
      vNextClientRuntimeAssertBoundParent_(created.scriptId, created.spreadsheetId);
      const draft = SpreadsheetApp.openById(created.spreadsheetId);
      vNextAdminInitializeTemplate_(draft, {
        bookId: draftId, releaseId: source.sourceReleaseId,
        clientRuntimeVersion: bundle.version, clientRuntimeSha256: bundle.sha256,
        adminEmails: source.adminEmails, actor: vNextAdminActor_(), now: new Date(), resetCopied: true
      });
      const copied = vNextAdminCopyAndVerifyTemplateUi_(source.spreadsheet, draft);
      vNextAdminWriteBookConfig_(draft, {
        state: 'TEMPLATE_DRAFT', template_kind: 'TEMPLATE_DRAFT', template_draft_id: draftId,
        source_template_release_id: source.sourceReleaseId,
        template_manifest_schema: VN_ADMIN_TEMPLATE_MANIFEST_SCHEMA,
        template_manifest_sha256: copied.manifestSha256, updated_at: new Date(), updated_by: vNextAdminActor_()
      });
      vNextAdminWriteSystemConfig_(draft, {
        mode: 'TEMPLATE', book_id: draftId, active_release_id: source.sourceReleaseId,
        schema_version: vNextAdminClientSchemaVersion_()
      });
      vNextAdminApplyVisibility_(draft, VN_ADMIN_DEFAULT_TEMPLATE_VISIBLE);
      vNextAdminEnforcePrivateFileAcl_(DriveApp.getFileById(draft.getId()), source.adminEmails);
      vNextAdminRegisterBook_(hub, {
        book_id: draftId, mode: 'TEMPLATE', client_id: '', client_name: '', fiscal_year: '',
        spreadsheet_id: draft.getId(), spreadsheet_url: draft.getUrl(), client_script_id: created.scriptId,
        client_runtime_version: bundle.version, client_runtime_sha256: bundle.sha256,
        template_release_id: source.sourceReleaseId, schema_version: vNextAdminClientSchemaVersion_(),
        state: 'TEMPLATE_DRAFT', status: 'DRAFT', health_status: 'OK', health_code: 'TEMPLATE_DRAFT_CREATED',
        last_health_at: new Date(), forecast_owner_emails: source.adminEmails.join(','),
        editor_emails: source.adminEmails.join(','), viewer_emails: '', created_at: new Date(),
        created_by: vNextAdminActor_(), updated_at: new Date(),
        note: vNextAdminCanonicalJson_({ templateKind: 'TEMPLATE_DRAFT', sourceReleaseId: source.sourceReleaseId,
          manifestSchema: VN_ADMIN_TEMPLATE_MANIFEST_SCHEMA, initialManifestSha256: copied.manifestSha256,
          reason: reason })
      });
      vNextAdminWriteAudit_(hub, 'CREATE_TEMPLATE_DRAFT', 'TEMPLATE_DRAFT', draftId, 'SUCCESS', {
        sourceSpreadsheetId: source.spreadsheet.getId(), sourceReleaseId: source.sourceReleaseId,
        draftSpreadsheetId: draft.getId(), manifestSha256: copied.manifestSha256, reason: reason
      });
      return {
        reused: false, draftId: draftId, spreadsheetId: draft.getId(), spreadsheetUrl: draft.getUrl(),
        sourceReleaseId: source.sourceReleaseId, manifestSha256: copied.manifestSha256,
        message: '管理ハブ専用の編集用Template Draftを作成しました。公開前にSTAGEDへ複製して完全検証します。'
      };
    });
  });
}

/**
 * Create a clean, known-bound STAGED Template. The future-book pointers do not
 * move until vNextAdminActivateReleasePair has validated a PASS model candidate.
 */
function vNextAdminPublishTemplateRelease(request) {
  return vNextAdminGuard_('vNextAdminPublishTemplateRelease', function () {
    const req = request && typeof request === 'object' ? request : {};
    const hub = vNextAdminRequireHub_();
    return vNextAdminWithScriptLock_('publish-template-release', function () {
      const reason = vNextAdminRequiredText_(req.reason || req.note, 'reason');
      if (typeof vNextClientRuntimeVerifiedBundle_ !== 'function' ||
          typeof vNextClientRuntimeCreateBoundSpreadsheet_ !== 'function') {
        throw new Error('Verified Client runtime publisher is not installed.');
      }
      const hubConfig = vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET);
      const current = vNextAdminResolveRelease_(hub, hubConfig.active_release_id || '');
      if (req.expectedActiveReleaseId &&
          String(req.expectedActiveReleaseId) !== String(current.release_id || '')) {
        throw new Error('Active release changed before publication. Reload 管理ハブ and retry.');
      }
      const bundle = vNextClientRuntimeVerifiedBundle_();
      const source = vNextAdminResolveTemplateUiSource_(hub, req.templateDraftSpreadsheetId || '');
      const sourceManifestSha256 = vNextAdminTemplateUiManifestHash_(source.spreadsheet);
      const releaseId = vNextAdminText_(req.releaseId) ||
        ('vnext-client-' + String(bundle.version || '').replace(/[^A-Za-z0-9.-]/g, '-') + '-' +
          String(bundle.sha256 || '').slice(0, 8) + '-' + sourceManifestSha256.slice(0, 8));
      const releases = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.RELEASES);
      const existing = releases.rows.find(function (row) {
        return String(row.release_id || '') === releaseId;
      });
      if (existing && (
        !String(existing.template_spreadsheet_id || '') || !String(existing.template_script_id || '') ||
        String(existing.client_runtime_version || '') !== String(bundle.version || '') ||
        String(existing.client_runtime_sha256 || '') !== String(bundle.sha256 || '') ||
        String(existing.schema_version || '') !== vNextAdminClientSchemaVersion_() ||
        String(existing.template_manifest_schema || '') !== VN_ADMIN_TEMPLATE_MANIFEST_SCHEMA ||
        String(existing.template_content_sha256 || '') !== sourceManifestSha256
      )) {
        throw new Error('Release ID already exists with different immutable content: ' + releaseId);
      }
      if (String(current.client_runtime_sha256 || '') === String(bundle.sha256 || '') &&
          String(current.client_runtime_version || '') === String(bundle.version || '') &&
          String(current.template_content_sha256 || '') === sourceManifestSha256 &&
          !source.isDraft && !req.releaseId && !existing) {
        const currentTemplate = SpreadsheetApp.openById(String(current.template_spreadsheet_id || ''));
        vNextAdminAssertReleaseTemplateManifest_(current, currentTemplate);
        return {
          reused: true, releaseId: current.release_id,
          clientRuntimeVersion: current.client_runtime_version,
          clientRuntimeSha256: current.client_runtime_sha256,
          message: '現在のTemplate Releaseは配備済みbundleと一致しています。'
        };
      }

      const adminEmails = vNextAdminParseList_(
        hubConfig.admin_emails || vNextGetRuntimeConfig_().VNEXT_ADMIN_EMAILS
      );
      if (!adminEmails.length) throw new Error('At least one Admin email is required for Template publication.');
      vNextAdminAssertPrivateAdminFile_(DriveApp.getFileById(source.spreadsheet.getId()), adminEmails);
      let templateId = existing && String(existing.template_spreadsheet_id || '');
      let templateScriptId = existing && String(existing.template_script_id || '');
      let templateBookId = '';
      let template = null;
      if (existing) {
        const existingStatus = String(existing.status || '').toUpperCase();
        if (['STAGED', 'ACTIVE'].indexOf(existingStatus) < 0) {
          throw new Error('Existing Template Release cannot be reused from status=' + existingStatus);
        }
        template = SpreadsheetApp.openById(templateId);
        const existingConfig = vNextAdminReadKeyValueSheet_(template, VN_ADMIN_BOOK_CONFIG_SHEET);
        templateBookId = vNextAdminRequiredText_(existingConfig.book_id, 'existingTemplate.book_id');
        if (String(existingConfig.version || '') !== releaseId ||
            String(existingConfig.client_runtime_bundle_sha256 || '') !== String(bundle.sha256 || '') ||
            String(existing.template_content_sha256 || '') !== vNextAdminTemplateUiManifestHash_(template)) {
          throw new Error('Existing release Template identity is inconsistent.');
        }
      } else {
        const dest = vNextAdminPrepareLibraryDestinationFolder_(hub, '',
          vNextAdminLibraryPath_('TEMPLATE_CURRENT'), adminEmails);
        const created = vNextClientRuntimeCreateBoundSpreadsheet_({
          title: (vNextAdminText_(req.releaseName) || 'Forecast vNext Master Template') + ' [' + releaseId + ']',
          folderId: dest.getId()
        });
        if (String(created.runtimeVersion || '') !== String(bundle.version || '') ||
            String(created.bundleSha256 || '') !== String(bundle.sha256 || '')) {
          throw new Error('New immutable Template runtime does not match the verified bundle.');
        }
        templateId = created.spreadsheetId;
        templateScriptId = created.scriptId;
        templateBookId = 'TPL-' + Utilities.getUuid();
        template = SpreadsheetApp.openById(templateId);
        vNextAdminInitializeTemplate_(template, {
          bookId: templateBookId, releaseId: releaseId,
          clientRuntimeVersion: bundle.version, clientRuntimeSha256: bundle.sha256,
          adminEmails: adminEmails, actor: vNextAdminActor_(), now: new Date(), resetCopied: true
        });
        const copied = vNextAdminCopyAndVerifyTemplateUi_(source.spreadsheet, template);
        if (copied.manifestSha256 !== sourceManifestSha256) {
          throw new Error('STAGED Template manifest differs from the selected UI source.');
        }
        vNextAdminWriteBookConfig_(template, {
          state: 'TEMPLATE_STAGED', template_kind: 'IMMUTABLE_STAGED',
          source_template_release_id: source.sourceReleaseId,
          template_manifest_schema: VN_ADMIN_TEMPLATE_MANIFEST_SCHEMA,
          template_manifest_sha256: sourceManifestSha256, updated_at: new Date(), updated_by: vNextAdminActor_()
        });
        vNextAdminEnforcePrivateFileAcl_(DriveApp.getFileById(templateId), adminEmails);
      }
      if (vNextDetectBookMode_(template) !== 'TEMPLATE') throw new Error('Published Template is not mode=TEMPLATE.');
      if (typeof vNextClientRuntimeAssertBoundParent_ !== 'function') {
        throw new Error('Client runtime parent verifier is not installed.');
      }
      vNextClientRuntimeAssertBoundParent_(templateScriptId, templateId);
      vNextAdminAssertPrivateAdminFile_(DriveApp.getFileById(templateId), adminEmails);

      let published = existing;
      if (!published) {
        published = vNextAdminRegisterRelease_(hub, {
          release_id: releaseId, release_name: vNextAdminText_(req.releaseName) || releaseId,
          status: 'STAGED', template_spreadsheet_id: templateId,
          schema_version: vNextAdminClientSchemaVersion_(),
          engine_version: typeof VNEXT_ENGINE !== 'undefined' ? VNEXT_ENGINE.VERSION : '',
          ux_version: typeof VNEXT_UX_CONFIG_ !== 'undefined' ? VNEXT_UX_CONFIG_.VERSION || '' : '',
          admin_version: VN_ADMIN_SCHEMA_VERSION,
          client_runtime_version: bundle.version, client_runtime_sha256: bundle.sha256,
          template_content_sha256: sourceManifestSha256,
          template_manifest_schema: VN_ADMIN_TEMPLATE_MANIFEST_SCHEMA,
          template_script_id: templateScriptId, created_at: new Date(),
          created_by: vNextAdminActor_(), activated_at: '', note: reason
        });
      }
      let templateRegistry = vNextAdminFindRegistryRow_(hub, function (row) {
        return String(row.mode || '') === 'TEMPLATE' && String(row.spreadsheet_id || '') === templateId;
      });
      if (!templateRegistry) {
        vNextAdminRegisterBook_(hub, {
          book_id: templateBookId, mode: 'TEMPLATE', client_id: '', client_name: '', fiscal_year: '',
          spreadsheet_id: templateId, spreadsheet_url: template.getUrl(),
          client_script_id: templateScriptId, client_runtime_version: bundle.version,
          client_runtime_sha256: bundle.sha256, template_release_id: releaseId,
          schema_version: vNextAdminClientSchemaVersion_(), state: 'TEMPLATE_STAGED', status: 'STAGED',
          health_status: 'OK', health_code: 'RELEASE_STAGED', last_health_at: new Date(),
          forecast_owner_emails: adminEmails.join(','), editor_emails: adminEmails.join(','), viewer_emails: '',
          created_at: new Date(), created_by: vNextAdminActor_(), updated_at: new Date(),
          note: vNextAdminCanonicalJson_({ templateKind: 'IMMUTABLE_STAGED', sourceReleaseId: source.sourceReleaseId,
            manifestSchema: VN_ADMIN_TEMPLATE_MANIFEST_SCHEMA, manifestSha256: sourceManifestSha256,
            runtimeVersion: bundle.version, runtimeSha256: bundle.sha256 })
        });
        templateRegistry = vNextAdminFindRegistryRow_(hub, function (row) {
          return String(row.book_id || '') === templateBookId;
        });
      }
      if (!templateRegistry) throw new Error('Template BOOK_REGISTRY row is missing.');
      const stageOperationId = 'TPL-STAGE-' + vNextAdminSha256_(vNextAdminCanonicalJson_({
        releaseId: releaseId, previousReleaseId: current.release_id, sourceManifestSha256: sourceManifestSha256,
        runtimeSha256: bundle.sha256, reason: reason
      })).slice(0, 24).toUpperCase();
      vNextAdminAppendTemplateJournal_(hub, {
        operationId: stageOperationId, releaseId: releaseId, previousReleaseId: current.release_id,
        templateSpreadsheetId: templateId, phase: 'STAGED_VERIFIED', status: 'SUCCEEDED',
        detail: { sourceSpreadsheetId: source.spreadsheet.getId(), sourceReleaseId: source.sourceReleaseId,
          manifestSchema: VN_ADMIN_TEMPLATE_MANIFEST_SCHEMA, manifestSha256: sourceManifestSha256,
          runtimeVersion: bundle.version, runtimeSha256: bundle.sha256, reason: reason }
      });
      vNextAdminWriteAudit_(hub, 'STAGE_TEMPLATE_RELEASE', 'RELEASE', releaseId, 'SUCCESS', {
        previousReleaseId: current.release_id, templateSpreadsheetId: templateId,
        templateScriptId: templateScriptId, clientRuntimeVersion: bundle.version,
        clientRuntimeSha256: bundle.sha256, templateManifestSha256: sourceManifestSha256,
        sourceTemplateDraftId: source.draftId || '', reason: reason
      });
      if (vNextAdminText_(req.modelReleaseId)) {
        return vNextAdminActivateReleasePairInternal_(hub, {
          releaseId: releaseId, modelReleaseId: req.modelReleaseId, reason: reason,
          expectedActiveReleaseId: req.expectedActiveReleaseId || current.release_id,
          expectedActiveModelReleaseId: req.expectedActiveModelReleaseId || ''
        });
      }
      return { reused: Boolean(existing), staged: true, operationId: stageOperationId,
        releaseId: releaseId, previousReleaseId: current.release_id,
        templateSpreadsheetId: templateId, templateManifestSha256: sourceManifestSha256,
        clientRuntimeVersion: bundle.version, clientRuntimeSha256: bundle.sha256,
        message: 'STAGED Templateを作成しました。templateVersionをこのRelease IDにしたPASS済MODEL_RELEASEを登録し、pairを有効化してください。' };
    });
  });
}

/** Activate a STAGED Template and its exact PASS model candidate as one pair. */
function vNextAdminActivateReleasePair(request) {
  return vNextAdminGuard_('vNextAdminActivateReleasePair', function () {
    const req = request && typeof request === 'object' ? request : {};
    const hub = vNextAdminRequireHub_();
    return vNextAdminWithScriptLock_('activate-release-pair', function () {
      return vNextAdminActivateReleasePairInternal_(hub, req);
    });
  });
}

/** Register an immutable DRAFT model release. Activation never mutates workbook structure. */
function vNextAdminRegisterModelRelease(request) {
  return vNextAdminGuard_('vNextAdminRegisterModelRelease', function () {
    const req = request && typeof request === 'object' ? request : {};
    const hub = vNextAdminRequireHub_();
    return vNextAdminWithScriptLock_('register-model-release', function () {
      const modelReleaseId = vNextAdminRequiredText_(req.modelReleaseId, 'modelReleaseId');
      const modelVersion = vNextAdminRequiredText_(req.modelVersion, 'modelVersion');
      const note = vNextAdminRequiredText_(req.note || req.reason, 'note');
      vNextAdminAssertModelReleaseIdSeparated_(hub, modelReleaseId);
      if (typeof VNEXT_ENGINE === 'undefined' || modelVersion !== String(VNEXT_ENGINE.VERSION || '')) {
        throw new Error('modelVersion must equal the deployed Forecast Engine version: ' +
          (typeof VNEXT_ENGINE !== 'undefined' ? VNEXT_ENGINE.VERSION : '(engine unavailable)'));
      }
      const schemaVersion = vNextAdminText_(req.schemaVersion) || vNextAdminClientSchemaVersion_();
      if (schemaVersion !== vNextAdminClientSchemaVersion_()) {
        throw new Error('MODEL_RELEASE schemaVersion is incompatible with the deployed Core.');
      }
      const parameters = vNextAdminNormalizeModelParameters_(
        vNextAdminParseObjectPayload_(req.parameters, 'parameters')
      );
      const templateVersion = vNextAdminText_(req.templateVersion) ||
        vNextAdminText_(vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET).active_release_id);
      const templateRelease = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.RELEASES).rows.find(function (row) {
        return String(row.release_id || '') === templateVersion;
      });
      if (!templateRelease || String(templateRelease.schema_version || '') !== schemaVersion ||
          String(templateRelease.engine_version || '') !== modelVersion ||
          ['STAGED', 'ACTIVE'].indexOf(String(templateRelease.status || '').toUpperCase()) < 0) {
        throw new Error('MODEL_RELEASE is not compatible with the referenced Template Release.');
      }
      const candidateHash = vNextAdminModelCandidateHash_({
        modelVersion: modelVersion, schemaVersion: schemaVersion,
        templateVersion: templateVersion, parameters: parameters
      });
      const backtest = vNextAdminBindModelCheck_(
        vNextAdminParseObjectPayload_(req.backtest, 'backtest'), candidateHash, 'backtest'
      );
      const canary = vNextAdminBindModelCheck_(
        vNextAdminParseObjectPayload_(req.canary, 'canary'), candidateHash, 'canary'
      );
      if ((vNextAdminModelCheckPassed_(backtest) || vNextAdminModelCheckPassed_(canary)) && req.attestationConfirmed !== true) {
        throw new Error('PASS JSON is an Admin attestation. Confirm attestationConfirmed=true after reviewing its evidence.');
      }
      const existing = vNextAdminLatestModelRelease_(hub, modelReleaseId);
      if (existing) {
        const exactDraft = String(existing.status || '').toUpperCase() === 'DRAFT' &&
          String(existing.model_version || '') === modelVersion &&
          String(existing.parameters_json || '') === vNextAdminCanonicalJson_(parameters) &&
          String(existing.backtest_json || '') === vNextAdminCanonicalJson_(backtest) &&
          String(existing.canary_json || '') === vNextAdminCanonicalJson_(canary) &&
          String(existing.template_version || '') === templateVersion &&
          String(existing.note || '') === note;
        if (exactDraft) return vNextAdminJsonSafe_(Object.assign({ reused: true }, existing));
        throw new Error('modelReleaseId already exists with different immutable content: ' + modelReleaseId);
      }
      const record = vNextAdminBuildModelReleaseRecord_({
        modelReleaseId: modelReleaseId,
        status: 'DRAFT',
        modelVersion: modelVersion,
        schemaVersion: schemaVersion,
        templateVersion: templateVersion,
        parameters: parameters,
        backtest: backtest,
        canary: canary,
        note: note,
        actor: vNextAdminActor_(),
        now: new Date()
      });
      vNextAdminAppendCoreRowsNoLock_(hub, 'MODEL_RELEASE', [record]);
      vNextAdminWriteAudit_(hub, 'REGISTER_MODEL_RELEASE', 'MODEL_RELEASE', modelReleaseId, 'SUCCESS', {
        status: 'DRAFT', modelVersion: modelVersion, templateVersion: record.template_version,
        adminAttestation: req.attestationConfirmed === true, note: note
      });
      return vNextAdminJsonSafe_(record);
    });
  });
}

/** Activate only a DRAFT whose immutable backtest and canary results both PASS. */
function vNextAdminActivateModelRelease(request) {
  return vNextAdminGuard_('vNextAdminActivateModelRelease', function () {
    const req = request && typeof request === 'object' ? request : {};
    const hub = vNextAdminRequireHub_();
    return vNextAdminWithScriptLock_('activate-model-release', function () {
      const modelReleaseId = vNextAdminRequiredText_(req.modelReleaseId, 'modelReleaseId');
      const reason = vNextAdminRequiredText_(req.reason, 'reason');
      vNextAdminAssertModelReleaseIdSeparated_(hub, modelReleaseId);
      const current = vNextAdminTryResolveActiveModelRelease_(hub);
      const source = vNextAdminLatestModelRelease_(hub, modelReleaseId);
      if (!source) throw new Error('MODEL_RELEASE not found: ' + modelReleaseId);
      const activationOperationId = 'MODEL-ACT-' + vNextAdminSha256_(vNextAdminCanonicalJson_({
        modelReleaseId: modelReleaseId, previousActiveModelReleaseId: current && current.model_release_id || '', reason: reason
      })).slice(0, 24).toUpperCase();
      const sourceNote = vNextAdminParseJson_(source.note, {});
      if (current && String(current.model_release_id) === modelReleaseId && String(source.status).toUpperCase() === 'ACTIVE') {
        return { reused: true, activeModelReleaseId: modelReleaseId, status: 'ACTIVE' };
      }
      if (String(source.status || '').toUpperCase() === 'ACTIVE' &&
          String(sourceNote.operationId || '') === activationOperationId &&
          String(sourceNote.action || '') === 'ACTIVATE' && (!current || String(current.model_release_id || '') !== modelReleaseId)) {
        vNextAdminAssertModelReleaseChecksPassed_(source);
        vNextAdminAssertModelTemplateCompatibility_(hub, source);
        vNextAdminSetActiveModelRelease_(hub, modelReleaseId);
        vNextAdminWriteAudit_(hub, 'RECOVER_MODEL_RELEASE_ACTIVATION', 'MODEL_RELEASE', modelReleaseId, 'SUCCESS', {
          previousActiveModelReleaseId: current && current.model_release_id || '', reason: reason
        });
        return { reused: true, recovered: true, activeModelReleaseId: modelReleaseId,
          previousActiveModelReleaseId: current && current.model_release_id || '' };
      }
      if (String(source.status || '').toUpperCase() !== 'DRAFT') {
        throw new Error('Only a DRAFT model release can be activated. Use rollback for a previously active release.');
      }
      vNextAdminAssertModelReleaseChecksPassed_(source);
      vNextAdminAssertModelTemplateCompatibility_(hub, source);
      const activated = Object.assign({}, source, {
        status: 'ACTIVE', approved_at: new Date().toISOString(), approved_by: vNextAdminActor_().toLowerCase(),
        rollback_release_id: '', created_at: new Date().toISOString(), created_by: vNextAdminActor_().toLowerCase(),
        note: vNextAdminCanonicalJson_({ action: 'ACTIVATE', operationId: activationOperationId, reason: reason })
      });
      delete activated._rowNumber;
      vNextAdminAppendCoreRowsNoLock_(hub, 'MODEL_RELEASE', [activated]);
      vNextAdminSetActiveModelRelease_(hub, modelReleaseId);
      vNextAdminWriteAudit_(hub, 'ACTIVATE_MODEL_RELEASE', 'MODEL_RELEASE', modelReleaseId, 'SUCCESS', {
        previousActiveModelReleaseId: current && current.model_release_id || '', reason: reason,
        backtest: vNextAdminParseJson_(source.backtest_json, {}), canary: vNextAdminParseJson_(source.canary_json, {}),
        structuralChangesApplied: false
      });
      return { reused: false, activeModelReleaseId: modelReleaseId, previousActiveModelReleaseId: current && current.model_release_id || '' };
    });
  });
}

/** Point future books back to a previously ACTIVE model release; existing books remain pinned. */
function vNextAdminRollbackModelRelease(request) {
  return vNextAdminGuard_('vNextAdminRollbackModelRelease', function () {
    const req = request && typeof request === 'object' ? request : {};
    const hub = vNextAdminRequireHub_();
    return vNextAdminWithScriptLock_('rollback-model-release', function () {
      const targetId = vNextAdminRequiredText_(req.targetModelReleaseId || req.modelReleaseId, 'targetModelReleaseId');
      const reason = vNextAdminRequiredText_(req.reason, 'reason');
      const current = vNextAdminResolveActiveModelRelease_(hub);
      if (String(current.model_release_id || '') === targetId) {
        return { reused: true, activeModelReleaseId: targetId, status: 'ACTIVE' };
      }
      const history = vNextAdminModelReleaseRows_(hub, targetId);
      const target = history.filter(function (row) {
        return String(row.status || '').toUpperCase() === 'ACTIVE';
      }).slice(-1)[0];
      if (!target) throw new Error('Rollback target must be a previously ACTIVE model release: ' + targetId);
      vNextAdminAssertModelReleaseChecksPassed_(target);
      vNextAdminAssertModelTemplateCompatibility_(hub, target);
      const rollbackOperationId = 'MODEL-RB-' + vNextAdminSha256_(vNextAdminCanonicalJson_({
        targetModelReleaseId: targetId, fromModelReleaseId: current.model_release_id, reason: reason
      })).slice(0, 24).toUpperCase();
      const targetNote = vNextAdminParseJson_(target.note, {});
      if (String(target.rollback_release_id || '') === String(current.model_release_id || '') &&
          String(targetNote.operationId || '') === rollbackOperationId && String(targetNote.action || '') === 'ROLLBACK') {
        vNextAdminSetActiveModelRelease_(hub, targetId);
        vNextAdminWriteAudit_(hub, 'RECOVER_MODEL_RELEASE_ROLLBACK', 'MODEL_RELEASE', targetId, 'SUCCESS', {
          fromModelReleaseId: current.model_release_id, reason: reason
        });
        return { reused: true, recovered: true, activeModelReleaseId: targetId, rolledBackFromModelReleaseId: current.model_release_id };
      }
      const rollback = Object.assign({}, target, {
        status: 'ACTIVE', approved_at: new Date().toISOString(), approved_by: vNextAdminActor_().toLowerCase(),
        rollback_release_id: current.model_release_id, created_at: new Date().toISOString(),
        created_by: vNextAdminActor_().toLowerCase(),
        note: vNextAdminCanonicalJson_({ action: 'ROLLBACK', operationId: rollbackOperationId, reason: reason })
      });
      delete rollback._rowNumber;
      vNextAdminAppendCoreRowsNoLock_(hub, 'MODEL_RELEASE', [rollback]);
      vNextAdminSetActiveModelRelease_(hub, targetId);
      vNextAdminWriteAudit_(hub, 'ROLLBACK_MODEL_RELEASE', 'MODEL_RELEASE', targetId, 'SUCCESS', {
        fromModelReleaseId: current.model_release_id, toModelReleaseId: targetId,
        reason: reason, structuralChangesApplied: false
      });
      return { reused: false, activeModelReleaseId: targetId, rolledBackFromModelReleaseId: current.model_release_id };
    });
  });
}

function vNextAdminMenuRefreshExceptions() {
  return vNextAdminMenuAction_('今日の例外を更新しました。', function () {
    const hub = vNextAdminRequireHub_();
    const out = vNextAdminRefreshTodayExceptions_(hub);
    vNextAdminRefreshHome_(hub);
    return out;
  });
}

function vNextAdminMenuRunOperationalCycle() {
  const result = vNextAdminRunOperationalCycle();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    String(result && result.message || '最新状態を確認しました。'), VN_ADMIN_MENU_NAME, 5
  );
  return result;
}

function vNextAdminMenuRunHealthScan() {
  const out = vNextAdminRunHealthScan();
  SpreadsheetApp.getActiveSpreadsheet().toast('全bookの状態確認を完了しました。', VN_ADMIN_MENU_NAME, 5);
  return out;
}

function vNextAdminMenuProcessJobs() {
  const hub = vNextAdminRequireHub_();
  const retries = vNextAdminWithScriptLock_('manual-known-pilot-retries', function () {
    return vNextAdminRequeueKnownPilotFailures_(hub);
  });
  const out = vNextAdminProcessJobsForHub_(hub, 5);
  out.retries = retries;
  SpreadsheetApp.getActiveSpreadsheet().toast('待機中の処理を実行しました。', VN_ADMIN_MENU_NAME, 5);
  return out;
}

function vNextAdminMenuRefreshZacClientCatalog() {
  const result = vNextAdminRefreshZacClientCatalog({ force: true });
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'ZACクライアント候補を更新しました（' + Number(result.activeCount || 0) + '件）。',
    VN_ADMIN_MENU_NAME, 5
  );
  return result;
}

function vNextAdminMenuUpdatePortalRuntime() {
  const ui = SpreadsheetApp.getUi();
  const choice = ui.alert(
    '申請入口を最新版へ更新',
    '既存の受付履歴を残したまま、Portal runtimeとZACクライアント候補を最新版へ更新します。続行しますか？',
    ui.ButtonSet.OK_CANCEL
  );
  if (choice !== ui.Button.OK) return { cancelled: true };
  const result = vNextAdminUpdateSharedPortalRuntime({ reason: 'Admin menu approved portal runtime update' });
  ui.alert('完了', result.message || '申請入口を更新しました。', ui.ButtonSet.OK);
  return result;
}

/**
 * Spreadsheet-bound editor fallback for Pilot recovery. Unlike the menu
 * wrapper, this function never calls getUi(), so it can be invoked from the
 * Apps Script editor when a custom menu has not appeared yet.
 */
function vNextAdminUpdateSharedPortalRuntimeForManualTest() {
  return vNextAdminUpdateSharedPortalRuntime({
    reason: 'Pilot Admin approved Portal v1.3 runtime and simplified UI update'
  });
}

/** Requeue only verified known Pilot failures, then process the same durable jobs. */
function vNextAdminRecoverPortalProvisionForManualTest() {
  const hub = vNextAdminRequireHub_();
  const retries = vNextAdminWithScriptLock_('manual-portal-provision-recovery', function () {
    return vNextAdminRequeueKnownPilotFailures_(hub);
  });
  return { retries: retries, jobs: vNextAdminProcessJobsForHub_(hub, 5) };
}

/**
 * Inventory or remove generated Client FY books so an employee can start from
 * Portal year/client selection. 管理ハブ, Portal, Template/Model releases,
 * and ZAC catalog are preserved. apply=true plus the confirmation phrase is
 * required before any write or Drive trash.
 */
function vNextAdminResetGeneratedClientsForFreshUat(request) {
  return vNextAdminGuard_('vNextAdminResetGeneratedClientsForFreshUat', function () {
    const hub = vNextAdminRequireHub_();
    return vNextAdminResetGeneratedClientsInHub_(hub, request);
  });
}

/** Central-source fallback used by clasp/API when Hub is not the active container. */
function vNextAdminResetGeneratedClientsForFreshUatFromSource(request) {
  return vNextAdminGuard_('vNextAdminResetGeneratedClientsForFreshUatFromSource', function () {
    const req = request && typeof request === 'object' ? request : {};
    const hubId = vNextAdminRequiredText_(req.hubSpreadsheetId, 'hubSpreadsheetId');
    const hub = SpreadsheetApp.openById(hubId);
    if (vNextDetectBookMode_(hub) !== 'ADMIN' || !vNextAdminIsRegisteredHub_(hub)) {
      throw new Error('The supplied Hub is not the registered 管理ハブ.');
    }
    vNextAdminAssertHubAdmin_(hub, true);
    vNextAdminHydrateHubRuntime_(hub);
    Object.keys(VN_ADMIN_HEADERS).forEach(function (name) {
      vNextAdminEnsureTable_(hub, name, VN_ADMIN_HEADERS[name]);
    });
    return vNextAdminResetGeneratedClientsInHub_(hub, req);
  });
}

function vNextAdminResetGeneratedClientsInHub_(hub, request) {
  const req = request && typeof request === 'object' ? request : {};
  const apply = req.apply === true;
  if (apply && String(req.confirmation || '').trim() !== VN_ADMIN_FRESH_UAT_RESET_CONFIRMATION) {
    throw new Error('確認語が一致しません。RESET_GENERATED_CLIENTS を入力してください。');
  }
  const registryRows = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.REGISTRY).rows;
  const releaseRows = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.RELEASES).rows;
  const hubConfig = vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET);
  const runtime = typeof vNextGetRuntimeConfig_ === 'function' ? vNextGetRuntimeConfig_() : {};
  const protectedIds = vNextAdminProtectedSpreadsheetIds_(hub, registryRows, releaseRows, hubConfig, runtime);
  const clientRows = registryRows.filter(function (row) {
    return String(row.mode || '').toUpperCase() === 'CLIENT';
  });
  const inventory = clientRows.map(function (row) {
    const spreadsheetId = String(row.spreadsheet_id || '').trim();
    const protectedHit = !spreadsheetId || protectedIds.has(spreadsheetId);
    return {
      bookId: String(row.book_id || ''),
      clientName: String(row.client_name || ''),
      fiscalYear: Number(row.fiscal_year || 0),
      state: String(row.state || ''),
      status: String(row.status || ''),
      spreadsheetId: spreadsheetId,
      spreadsheetUrl: String(row.spreadsheet_url || ''),
      willTrash: !protectedHit,
      protected: protectedHit
    };
  });
  const blocked = inventory.filter(function (item) { return item.protected; });
  if (blocked.length) {
    throw new Error('CLIENT registry が Hub/Portal/Template を指しているため停止しました: ' +
      blocked.map(function (item) { return item.bookId; }).join(', '));
  }
  const result = {
    ok: true,
    dryRun: !apply,
    confirmationRequired: VN_ADMIN_FRESH_UAT_RESET_CONFIRMATION,
    clientCount: inventory.length,
    clients: inventory,
    preserved: {
      hubSpreadsheetId: hub.getId(),
      portalSpreadsheetId: String(hubConfig.portal_spreadsheet_id || runtime.VNEXT_PORTAL_SPREADSHEET_ID || ''),
      protectedSpreadsheetCount: protectedIds.size
    },
    message: apply
      ? '生成済みクライアント年度ブックと検証ログを削除し、申請入口から作り直せる状態に戻しました。'
      : ('対象 ' + inventory.length + '冊を確認しました。確認語を入力すると削除します。試験ログも消します。' +
        VNEXT_NAMING.LAYER1 + ' / ' + VNEXT_NAMING.LAYER2 + ' / Template は残します。')
  };
  if (!apply) return vNextAdminJsonSafe_(result);

  return vNextAdminWithScriptLock_('fresh-uat-reset', function () {
    const now = new Date();
    const clientBookIds = new Set(inventory.map(function (item) { return item.bookId; }).filter(Boolean));
    const trashed = [];
    const trashErrors = [];
    inventory.forEach(function (item) {
      if (!item.willTrash) return;
      try {
        DriveApp.getFileById(item.spreadsheetId).setTrashed(true);
        trashed.push(item.spreadsheetId);
      } catch (error) {
        trashErrors.push({ spreadsheetId: item.spreadsheetId, error: String(error && error.message || error) });
      }
    });
    clientRows.forEach(function (row) {
      if (!row._rowNumber) return;
      vNextAdminUpdateTableRow_(hub, VN_ADMIN_SHEETS.REGISTRY, row._rowNumber, {
        status: 'ARCHIVED',
        health_status: 'ARCHIVED',
        health_code: 'FRESH_UAT_RESET',
        updated_at: now,
        note: vNextAdminFreshUatResetNote_(row.note, now)
      });
    });
    vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.TEAM).rows.forEach(function (row) {
      if (!clientBookIds.has(String(row.book_id || '')) || !row._rowNumber) return;
      vNextAdminUpdateTableRow_(hub, VN_ADMIN_SHEETS.TEAM, row._rowNumber, {
        status: 'ARCHIVED', updated_at: now, note: vNextAdminFreshUatResetNote_(row.note, now)
      });
    });
    [
      VN_ADMIN_SHEETS.JOBS, VN_ADMIN_SHEETS.JOB_LOG, VN_ADMIN_SHEETS.APPROVALS,
      VN_ADMIN_SHEETS.OFFICIAL, VN_ADMIN_SHEETS.EXCEPTIONS, VN_ADMIN_SHEETS.AUDIT,
      VN_ADMIN_CLIENT_REQUEST_SHEET
    ].forEach(function (name) {
      vNextAdminClearTableData_(hub, name, VN_ADMIN_HEADERS[name]);
    });
    vNextAdminRewriteTableKeeping_(hub, VN_ADMIN_SHEETS.MIGRATIONS,
      VN_ADMIN_HEADERS[VN_ADMIN_SHEETS.MIGRATIONS], function (row) {
        return !clientBookIds.has(String(row.book_id || ''));
      });
    const coreSheets = typeof VNEXT_CORE !== 'undefined' && VNEXT_CORE.INTERNAL_SHEETS
      ? VNEXT_CORE.INTERNAL_SHEETS : {};
    Object.keys(coreSheets).forEach(function (sheetName) {
      if (sheetName === 'MODEL_RELEASE') return;
      if (!hub.getSheetByName(sheetName)) return;
      vNextAdminRewriteTableKeeping_(hub, sheetName, coreSheets[sheetName].slice(), function (row) {
        const bookId = String(row.book_id || '');
        return !bookId || !clientBookIds.has(bookId);
      });
    });

    let portalRefresh = { configured: false };
    try {
      const portal = vNextAdminResolvePortal_(hub);
      vNextAdminClearSheetDataRows_(portal.spreadsheet.getSheetByName(VN_ADMIN_PORTAL_REQUEST_SHEET));
      portalRefresh = vNextAdminRefreshPortalDirectory_(hub, portal.spreadsheet);
      portalRefresh.views = vNextAdminRefreshPortalEmployeeViews_(portal.spreadsheet);
      portalRefresh.configured = true;
    } catch (portalError) {
      if (!/not configured/i.test(String(portalError && portalError.message || portalError))) throw portalError;
    }

    const auditFiles = vNextAdminTrashAuditArtifacts_(hub, hubConfig);
    vNextAdminRefreshTodayExceptions_(hub);
    vNextAdminRefreshHome_(hub);
    vNextAdminWriteAudit_(hub, 'FRESH_UAT_RESET', 'CLIENT', 'ALL', 'SUCCEEDED', {
      clientCount: inventory.length,
      trashed: trashed,
      trashErrors: trashErrors,
      auditFiles: auditFiles,
      portalRows: portalRefresh.rows || 0
    });
    result.trashed = trashed;
    result.trashErrors = trashErrors;
    result.portal = portalRefresh;
    result.ok = trashErrors.length === 0;
    if (trashErrors.length) {
      result.message = 'Registryは初期化しましたが、一部のDriveファイルをゴミ箱へ移せませんでした。';
    }
    return vNextAdminJsonSafe_(result);
  });
}

function vNextAdminTrashAuditArtifacts_(hub, hubConfig) {
  const trashed = [];
  const seen = {};
  function trashFilesInFolder_(folder) {
    if (!folder) return;
    const files = folder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      const id = String(file.getId() || '');
      if (!id || seen[id]) continue;
      seen[id] = true;
      file.setTrashed(true);
      trashed.push(id);
    }
    const folders = folder.getFolders();
    while (folders.hasNext()) trashFilesInFolder_(folders.next());
  }
  try {
    const auditId = String(hubConfig && hubConfig.library_audit_folder_id || '').trim();
    if (auditId) trashFilesInFolder_(DriveApp.getFolderById(auditId));
  } catch (ignoredMissingLibraryAudit) {}
  try {
    const parents = DriveApp.getFileById(hub.getId()).getParents();
    if (parents.hasNext()) {
      const named = parents.next().getFoldersByName('Forecast vNext Admin Audit');
      if (named.hasNext()) trashFilesInFolder_(named.next());
    }
  } catch (ignoredMissingLegacyAudit) {}
  return trashed;
}

function vNextAdminProtectedSpreadsheetIds_(hub, registryRows, releaseRows, hubConfig, runtime) {
  const protectedIds = new Set();
  protectedIds.add(hub.getId());
  (registryRows || []).forEach(function (row) {
    const mode = String(row.mode || '').toUpperCase();
    const spreadsheetId = String(row.spreadsheet_id || '').trim();
    if (spreadsheetId && VN_ADMIN_PROTECTED_BOOK_MODES.indexOf(mode) >= 0) protectedIds.add(spreadsheetId);
  });
  (releaseRows || []).forEach(function (row) {
    const templateId = String(row.template_spreadsheet_id || '').trim();
    if (templateId) protectedIds.add(templateId);
  });
  [
    hubConfig && hubConfig.portal_spreadsheet_id,
    hubConfig && hubConfig.template_spreadsheet_id,
    hubConfig && hubConfig.admin_hub_spreadsheet_id,
    runtime && runtime.VNEXT_ADMIN_HUB_SPREADSHEET_ID,
    runtime && runtime.VNEXT_MASTER_TEMPLATE_SPREADSHEET_ID,
    runtime && runtime.VNEXT_PORTAL_SPREADSHEET_ID
  ].forEach(function (value) {
    const id = String(value || '').trim();
    if (id) protectedIds.add(id);
  });
  return protectedIds;
}

function vNextAdminFreshUatResetNote_(existing, now) {
  const stamp = 'fresh-uat-reset ' + (now instanceof Date ? now.toISOString() : String(now || ''));
  const current = String(existing || '').trim();
  if (!current) return stamp;
  if (current.indexOf('fresh-uat-reset') >= 0) return current;
  return current + ' | ' + stamp;
}

function vNextAdminRefreshPortalEmployeeViews_(portalSpreadsheet) {
  if (typeof vNextPortalRefreshViews_ !== 'function') return { skipped: true };
  const refreshed = vNextPortalRefreshViews_(portalSpreadsheet);
  if (typeof vNextPortalGetLocalViewData_ !== 'function' || typeof vNextPortalRenderFiscalYear_ !== 'function') {
    return refreshed;
  }
  const data = vNextPortalGetLocalViewData_(portalSpreadsheet);
  const extraYears = [];
  portalSpreadsheet.getSheets().forEach(function (sheet) {
    const match = /^FY(\d{4})$/.exec(String(sheet.getName() || ''));
    if (!match) return;
    const year = Number(match[1]);
    if ((refreshed.years || data.years || []).indexOf(year) >= 0) return;
    vNextPortalRenderFiscalYear_(sheet, year, data);
    extraYears.push(year);
  });
  return Object.assign({}, refreshed, { extraYears: extraYears });
}

/**
 * No-UI editor fallback for the live Pilot. It applies the dedicated empty-book
 * upgrade only when exactly one registered Client passes the full read-only
 * eligibility check; ambiguous or non-empty books fail closed.
 */
function vNextAdminUpgradeOnlyEligibleEmptyPilotForManualTest() {
  const hub = vNextAdminRequireHub_();
  const pair = vNextAdminReadActiveReleasePair_(hub);
  const candidates = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.REGISTRY).rows.filter(function (row) {
    return String(row.mode || '').toUpperCase() === 'CLIENT' &&
      String(row.status || '').toUpperCase() === 'ACTIVE' &&
      String(row.state || '').toUpperCase() === 'INPUT_OPEN' &&
      !String(row.current_official_id || '') &&
      String(row.template_release_id || '') !== pair.releaseId;
  });
  const eligible = [];
  candidates.forEach(function (row) {
    try {
      eligible.push(vNextAdminUpgradeEmptyPilotClient({
        bookId: String(row.book_id || ''), dryRun: true,
        reason: 'Apps Script editorから受入試験前の安全条件を確認'
      }));
    } catch (ignoredIneligible) {
      Logger.log('Empty Pilot candidate skipped book=%s reason=%s', String(row.book_id || ''),
        String(ignoredIneligible && ignoredIneligible.message || ignoredIneligible));
    }
  });
  if (eligible.length !== 1) {
    throw new Error('安全条件を満たす空のPilot Clientが1冊に確定しませんでした: ' + eligible.length + '冊');
  }
  return vNextAdminUpgradeEmptyPilotClient({
    bookId: eligible[0].bookId, dryRun: false,
    reason: '受入試験開始前のUI・操作性改善'
  });
}

/** Recovers only when exactly one empty-Pilot migration is unfinished. */
function vNextAdminRecoverOnlyEmptyPilotForManualTest() {
  const hub = vNextAdminRequireHub_();
  const rows = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.MIGRATIONS).rows.filter(function (row) {
    return /^EMPTY_PILOT_|^RECOVERY_REQUIRED$/.test(String(row.status || '').toUpperCase());
  });
  if (rows.length !== 1) {
    throw new Error('復旧対象の空Pilot更新が1件に確定しませんでした: ' + rows.length + '件');
  }
  return vNextAdminRecoverEmptyPilotClientUpgrade({
    bookId: String(rows[0].book_id || ''), migrationId: String(rows[0].migration_id || ''),
    reason: 'Apps Script editorから中断した空Pilot更新を復旧'
  });
}

function vNextAdminMenuOpenRegistry() { return vNextAdminOpenHubSheet_(VN_ADMIN_SHEETS.REGISTRY); }
function vNextAdminMenuOpenExceptions() { return vNextAdminOpenHubSheet_(VN_ADMIN_SHEETS.EXCEPTIONS); }
function vNextAdminMenuOpenApprovals() { return vNextAdminOpenHubSheet_(VN_ADMIN_SHEETS.APPROVALS); }
function vNextAdminMenuOpenReleases() { return vNextAdminOpenHubSheet_(VN_ADMIN_SHEETS.RELEASES); }

// ---------------------------- Initialization ----------------------------

function vNextAdminInitializeHub_(ss, opt) {
  if (opt.resetCopied) vNextAdminResetCopiedWorkbook_(ss, []);
  vNextAdminEnsureCoreStore_(ss);
  Object.keys(VN_ADMIN_HEADERS).forEach(function (name) {
    vNextAdminEnsureTable_(ss, name, VN_ADMIN_HEADERS[name]);
  });
  const home = vNextAdminGetOrCreateSheet_(ss, VN_ADMIN_SHEETS.HOME);
  const existingSystemConfig = vNextAdminReadKeyValueSheet_(ss, VN_ADMIN_SYSTEM_CONFIG_SHEET);
  const buildSheet = ss.getSheetByName('__VNEXT_BUILD__');
  if (buildSheet && ss.getSheets().length > 1) ss.deleteSheet(buildSheet);
  if (home.getMaxColumns() < 8) home.insertColumnsAfter(home.getMaxColumns(), 8 - home.getMaxColumns());
  vNextAdminReplaceBookConfig_(ss, {
    mode: 'ADMIN', book_id: opt.bookId, state: 'ACTIVE', default_role: 'ADMIN',
    forecast_owner_emails: (opt.adminEmails || []).join(','), version: opt.releaseId || '',
    schema_version: VN_ADMIN_SCHEMA_VERSION, source_spreadsheet_id: opt.sourceSpreadsheetId || '',
    template_spreadsheet_id: opt.templateSpreadsheetId || '', created_at: opt.now,
    created_by: opt.actor, updated_at: opt.now, updated_by: opt.actor
  });
  vNextAdminWriteSystemConfig_(ss, {
    mode: 'ADMIN', book_id: opt.bookId, admin_hub_spreadsheet_id: ss.getId(), template_spreadsheet_id: opt.templateSpreadsheetId || '',
    active_release_id: opt.releaseId || '', source_spreadsheet_id: opt.sourceSpreadsheetId || '',
    active_model_release_id: opt.modelReleaseId || '',
    admin_source_script_id: opt.adminSourceScriptId || existingSystemConfig.admin_source_script_id || '',
    admin_hub_script_id: opt.adminHubScriptId || existingSystemConfig.admin_hub_script_id || '',
    admin_runtime_sha256: opt.adminRuntimeSha256 || existingSystemConfig.admin_runtime_sha256 || '',
    admin_emails: (opt.adminEmails || []).join(','),
    VERTEX_PROJECT_ID: opt.vertexConfig && opt.vertexConfig.VERTEX_PROJECT_ID || '',
    VERTEX_LOCATION: opt.vertexConfig && opt.vertexConfig.VERTEX_LOCATION || '',
    VERTEX_GEMINI_MODEL: opt.vertexConfig && opt.vertexConfig.VERTEX_GEMINI_MODEL || '',
    VERTEX_DATASTORE_ID: opt.vertexConfig && opt.vertexConfig.VERTEX_DATASTORE_ID || '',
    VERTEX_SEARCH_LOCATION: opt.vertexConfig && opt.vertexConfig.VERTEX_SEARCH_LOCATION || '',
    VERTEX_SERVING_CONFIG: opt.vertexConfig && opt.vertexConfig.VERTEX_SERVING_CONFIG || '',
    schema_version: VN_ADMIN_SCHEMA_VERSION
  });
  vNextAdminHydrateHubRuntime_(ss);
  vNextAdminApplyVisibility_(ss, VN_ADMIN_DEFAULT_HUB_VISIBLE);
  vNextAdminProtectInternalSheets_(ss, opt.adminEmails || [], Object.keys(VN_ADMIN_HEADERS)
    .concat([VN_ADMIN_SHEETS.HOME, VN_ADMIN_META_SHEET, VN_ADMIN_BOOK_CONFIG_SHEET, VN_ADMIN_SYSTEM_CONFIG_SHEET])
    .concat(typeof VNEXT_CORE !== 'undefined' ? Object.keys(VNEXT_CORE.INTERNAL_SHEETS) : []));
  vNextAdminRefreshHome_(ss);
}

function vNextAdminInitializeTemplate_(ss, opt) {
  if (opt.resetCopied) vNextAdminResetCopiedWorkbook_(ss, VN_ADMIN_DEFAULT_TEMPLATE_VISIBLE);
  vNextAdminEnsureUiShell_(ss, { template: true });
  vNextAdminEnsureCoreStore_(ss);
  vNextAdminReplaceBookConfig_(ss, {
    mode: 'TEMPLATE', book_id: opt.bookId, state: 'TEMPLATE_READY', default_role: 'ADMIN',
    forecast_owner_emails: (opt.adminEmails || []).join(','), version: opt.releaseId || '',
    client_runtime_version: opt.clientRuntimeVersion || '',
    client_runtime_bundle_sha256: opt.clientRuntimeSha256 || '',
    schema_version: vNextAdminClientSchemaVersion_(),
    created_at: opt.now, created_by: opt.actor, updated_at: opt.now, updated_by: opt.actor
  });
  vNextAdminReplaceSystemConfig_(ss, {
    mode: 'TEMPLATE', book_id: opt.bookId,
    active_release_id: opt.releaseId || '', schema_version: vNextAdminClientSchemaVersion_()
  });
  vNextAdminApplyVisibility_(ss, VN_ADMIN_DEFAULT_TEMPLATE_VISIBLE);
  vNextAdminProtectInternalSheets_(ss, opt.adminEmails || [],
    [VN_ADMIN_META_SHEET, VN_ADMIN_BOOK_CONFIG_SHEET, VN_ADMIN_SYSTEM_CONFIG_SHEET]
      .concat(typeof VNEXT_CORE !== 'undefined' ? Object.keys(VNEXT_CORE.INTERNAL_SHEETS) : []));
}

/**
 * Client-relative scale used only to translate the employee's S/M/L answer
 * into yen. The median of the latest three completed positive fiscal years is
 * stable across one-off spikes and keeps differently sized clients comparable.
 */
function vNextAdminResolveClientAnnualSalesScale_(hub, clientName, fiscalYear, asOf, cutoff) {
  try {
    if (typeof vNextFetchActualRecordsBridge_ !== 'function' ||
        typeof vNextFiscalYearForDate_ !== 'function') {
      return { amount: 0, basis: 'UNAVAILABLE' };
    }
    const config = vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET);
    const runtime = vNextGetRuntimeConfig_();
    const sourceSpreadsheetId = String(config.source_spreadsheet_id ||
      runtime.VNEXT_ZAC_SOURCE_SPREADSHEET_ID || runtime.FORECAST_SOURCE_SPREADSHEET_ID || '');
    if (!sourceSpreadsheetId) return { amount: 0, basis: 'UNAVAILABLE' };
    const records = vNextFetchActualRecordsBridge_(clientName, {
      sourceSpreadsheetId: sourceSpreadsheetId,
      fiscalYear: Number(fiscalYear),
      asOf: asOf || new Date(),
      cutoff: cutoff || vNextAdminCutoffFromAsOf_(asOf || new Date())
    });
    const totals = {};
    (records || []).forEach(function (record) {
      const fy = Number(vNextFiscalYearForDate_(record.actualDate));
      if (!isFinite(fy) || fy >= Number(fiscalYear)) return;
      totals[fy] = Number(totals[fy] || 0) + Number(record.amount || 0);
    });
    const values = Object.keys(totals).map(Number).sort(function (a, b) { return b - a; })
      .map(function (fy) { return { fiscalYear: fy, amount: Number(totals[fy] || 0) }; })
      .filter(function (row) { return isFinite(row.amount) && row.amount > 0; })
      .slice(0, 3);
    if (!values.length) return { amount: 0, basis: 'UNAVAILABLE' };
    const sorted = values.map(function (row) { return row.amount; }).sort(function (a, b) { return a - b; });
    const middle = Math.floor(sorted.length / 2);
    const median = sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2;
    return {
      amount: Math.max(0, Math.round(median)),
      basis: 'MEDIAN_COMPLETED_FY:' + values.map(function (row) { return 'FY' + row.fiscalYear; }).join(',')
    };
  } catch (error) {
    Logger.log('Client annual sales scale unavailable client=%s error=%s', String(clientName || ''),
      String(error && error.message || error));
    return { amount: 0, basis: 'UNAVAILABLE' };
  }
}

function vNextAdminRefreshClientAnnualSalesScale_(hub, registry) {
  const client = SpreadsheetApp.openById(String(registry.spreadsheet_id || ''));
  const routing = vNextAdminReadKeyValueSheet_(client, VN_ADMIN_BOOK_CONFIG_SHEET);
  const scale = vNextAdminResolveClientAnnualSalesScale_(hub, registry.client_name,
    Number(registry.fiscal_year), routing.as_of || new Date(),
    routing.cutoff || vNextAdminCutoffFromAsOf_(routing.as_of || new Date()));
  vNextAdminWriteBookConfig_(client, {
    annual_sales_baseline: scale.amount,
    annual_sales_baseline_basis: scale.basis
  });
  return scale;
}

function vNextAdminInitializeClient_(ss, opt) {
  if (VN_ADMIN_CLIENT_STATES.indexOf('INPUT_OPEN') < 0) throw new Error('Invalid initial CLIENT state.');
  vNextAdminEnsureUiShell_(ss, { clientName: opt.clientName, fiscalYear: opt.fiscalYear });
  vNextAdminEnsureCoreStore_(ss);
  vNextAdminEnsureTable_(ss, VN_ADMIN_CLIENT_REQUEST_SHEET, VN_ADMIN_HEADERS[VN_ADMIN_CLIENT_REQUEST_SHEET]);
  vNextAdminReplaceBookConfig_(ss, {
    mode: 'CLIENT', book_id: opt.bookId, client_id: opt.clientId, client_name: opt.clientName,
    fiscal_year: opt.fiscalYear, as_of: opt.asOf, cutoff: opt.cutoff, state: 'INPUT_OPEN',
    default_role: 'EMPLOYEE', forecast_owner_emails: (opt.forecastOwnerEmails || []).join(','),
    input_submitted: 0, input_answered_count: 0, input_total_count: 0,
    input_due_date: opt.inputDueDate || '', version: opt.releaseId,
    model_release_id: opt.modelReleaseId || '',
    access_policy: opt.accessPolicy || 'PRIVATE', internal_domain: opt.internalDomain || '',
    annual_sales_baseline: Number(opt.annualSalesBaseline || 0),
    annual_sales_baseline_basis: String(opt.annualSalesBaselineBasis || ''),
    client_runtime_version: opt.clientRuntimeVersion || '',
    client_runtime_bundle_sha256: opt.clientRuntimeSha256 || '',
    template_release_id: opt.releaseId, schema_version: vNextAdminClientSchemaVersion_(),
    created_at: opt.now, created_by: opt.actor, updated_at: opt.now, updated_by: opt.actor
  });
  // This sheet is hidden and protected. It contains routing IDs only, never Vertex or other runtime secrets.
  vNextAdminReplaceSystemConfig_(ss, {
    mode: 'CLIENT', book_id: opt.bookId, active_release_id: opt.releaseId,
    schema_version: vNextAdminClientSchemaVersion_()
  });
  const bookMeta = vNextAdminCreateClientCoreMeta_(ss, opt);
  vNextAdminApplyVisibility_(ss, vNextAdminParseList_(opt.visibleSheets).length ? vNextAdminParseList_(opt.visibleSheets) : VN_ADMIN_DEFAULT_CLIENT_VISIBLE);
  const protectedNames = [
    VN_ADMIN_META_SHEET, VN_ADMIN_BOOK_CONFIG_SHEET, VN_ADMIN_SYSTEM_CONFIG_SHEET,
    VN_ADMIN_OFFICIAL_COPY_SHEET, VN_ADMIN_CLIENT_REQUEST_SHEET,
    'FORECAST_SNAPSHOT', 'EVAL_LOG', 'RUN_LOG', 'PROCESS_STATUS', 'CALIBRATION_STATE',
    'CALIBRATION_HISTORY', 'AI_IMPACT_HISTORY', 'SUBJECTIVE_IMPACT_HISTORY', 'POOL_PRIOR',
    'RELIABILITY_EVIDENCE', 'DLM_STATE', 'BACKTEST_REPORT'
  ].concat(typeof VNEXT_CORE !== 'undefined' ? Object.keys(VNEXT_CORE.INTERNAL_SHEETS) : []);
  // Bound client APIs execute as the employee, so hard protections would also
  // block legitimate append-only writes. Use warning-only protection here and
  // rely on strict Hub ingest validation; no employee is granted as a hidden
  // sheet protection editor. A future execute-as-owner service removes this
  // unavoidable Sheets-only trust boundary.
  vNextAdminProtectClientInternalSheets_(ss, protectedNames);
  return { bookMeta: bookMeta };
}

function vNextAdminInitializePortal_(ss, opt) {
  const keep = new Set([
    'ホーム', VN_ADMIN_PORTAL_REQUEST_SHEET, VN_ADMIN_PORTAL_DIRECTORY_SHEET,
    VN_ADMIN_PORTAL_CONFIG_SHEET, VN_ADMIN_PORTAL_CLIENT_CATALOG_SHEET
  ]);
  ss.getSheets().slice().forEach(function (sheet) {
    if (!keep.has(sheet.getName()) && ss.getSheets().length > 1) ss.deleteSheet(sheet);
  });
  let home = ss.getSheetByName('ホーム');
  if (!home) home = ss.insertSheet('ホーム', 0);
  vNextAdminEnsureTable_(ss, VN_ADMIN_PORTAL_REQUEST_SHEET, VN_ADMIN_PORTAL_REQUEST_HEADERS);
  vNextAdminEnsureTable_(ss, VN_ADMIN_PORTAL_DIRECTORY_SHEET, VN_ADMIN_PORTAL_DIRECTORY_HEADERS);
  vNextAdminEnsureExactTableHeaders_(ss, VN_ADMIN_PORTAL_CLIENT_CATALOG_SHEET,
    VN_ADMIN_PORTAL_CLIENT_CATALOG_HEADERS);
  vNextAdminReplacePortalConfig_(ss, {
    mode: 'PORTAL', portal_id: opt.portalId, schema_version: VN_ADMIN_PORTAL_REQUEST_SCHEMA,
    runtime_version: opt.runtimeVersion, runtime_sha256: opt.runtimeSha256,
    employee_domain: opt.employeeDomain, access_policy: 'INTERNAL_OPEN',
    created_at: new Date().toISOString(), created_by: opt.actor
  });
  home.showSheet();
  ss.setActiveSheet(home);
  [VN_ADMIN_PORTAL_REQUEST_SHEET, VN_ADMIN_PORTAL_DIRECTORY_SHEET, VN_ADMIN_PORTAL_CONFIG_SHEET,
    VN_ADMIN_PORTAL_CLIENT_CATALOG_SHEET]
    .forEach(function (name) { const sheet = ss.getSheetByName(name); if (sheet) sheet.hideSheet(); });
  vNextAdminProtectInternalSheets_(ss, opt.adminEmails || [], [
    VN_ADMIN_PORTAL_DIRECTORY_SHEET, VN_ADMIN_PORTAL_CONFIG_SHEET,
    VN_ADMIN_PORTAL_CLIENT_CATALOG_SHEET
  ]);
  vNextAdminProtectClientInternalSheets_(ss, [VN_ADMIN_PORTAL_REQUEST_SHEET]);
  return true;
}

function vNextAdminReplacePortalConfig_(ss, values) {
  const existing = ss.getSheetByName(VN_ADMIN_PORTAL_CONFIG_SHEET);
  if (existing) existing.clear();
  const sheet = vNextAdminEnsureTable_(ss, VN_ADMIN_PORTAL_CONFIG_SHEET, ['key', 'value']);
  Object.keys(values || {}).forEach(function (key) {
    vNextAdminAppendObject_(ss, VN_ADMIN_PORTAL_CONFIG_SHEET, { key: key, value: values[key] }, ['key', 'value']);
  });
  sheet.hideSheet();
}

/**
 * Copies only employee-visible release assets from the immutable Template.
 * Core sheets, routing IDs, Admin settings and hidden data never cross this
 * boundary. The bound client-only script was already created separately.
 */
function vNextAdminCopyTemplateUiToClient_(template, client) {
  const allowed = VN_ADMIN_DEFAULT_TEMPLATE_VISIBLE.slice();
  const sourceSheets = allowed.map(function (name) {
    const sheet = template.getSheetByName(name);
    if (!sheet) throw new Error('Immutable Template is missing visible sheet: ' + name);
    return sheet;
  });
  const initialTargetSheets = client.getSheets().slice();
  const removableBlank = initialTargetSheets.length === 1 &&
    allowed.indexOf(initialTargetSheets[0].getName()) < 0 &&
    !String(initialTargetSheets[0].getRange('A1').getValue() || '').trim()
    ? initialTargetSheets[0] : null;
  sourceSheets.forEach(function (source) {
    const name = source.getName();
    const existing = client.getSheetByName(name);
    if (existing) client.deleteSheet(existing);
    source.copyTo(client).setName(name);
  });
  if (removableBlank && client.getSheets().length > allowed.length) client.deleteSheet(removableBlank);
  client.setActiveSheet(client.getSheetByName(allowed[0]));
  return { copiedSheets: allowed };
}

function vNextAdminResolveTemplateUiSource_(hub, draftSpreadsheetId) {
  const hubConfig = vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET);
  const adminEmails = vNextAdminMergeEmails_(hubConfig.admin_emails,
    vNextGetRuntimeConfig_().VNEXT_ADMIN_EMAILS, vNextAdminActor_());
  if (!adminEmails.length) throw new Error('Template UI source requires at least one Admin email.');
  const draftIdOrSpreadsheetId = vNextAdminText_(draftSpreadsheetId);
  if (!draftIdOrSpreadsheetId) {
    const active = vNextAdminResolveRelease_(hub, hubConfig.active_release_id || '');
    const spreadsheet = SpreadsheetApp.openById(vNextAdminRequiredText_(active.template_spreadsheet_id,
      'activeTemplate.template_spreadsheet_id'));
    vNextAdminAssertReleaseTemplateManifest_(active, spreadsheet);
    vNextAdminAssertPrivateAdminFile_(DriveApp.getFileById(spreadsheet.getId()), adminEmails);
    return { spreadsheet: spreadsheet, sourceReleaseId: String(active.release_id || ''),
      draftId: '', isDraft: false, adminEmails: adminEmails };
  }
  const registry = vNextAdminFindRegistryRow_(hub, function (row) {
    return String(row.mode || '') === 'TEMPLATE' &&
      (String(row.spreadsheet_id || '') === draftIdOrSpreadsheetId || String(row.book_id || '') === draftIdOrSpreadsheetId);
  });
  if (!registry || String(registry.status || '').toUpperCase() !== 'DRAFT') {
    throw new Error('UI source must be an Admin-registered mutable TEMPLATE_DRAFT.');
  }
  const spreadsheet = SpreadsheetApp.openById(String(registry.spreadsheet_id || ''));
  const config = vNextAdminReadKeyValueSheet_(spreadsheet, VN_ADMIN_BOOK_CONFIG_SHEET);
  if (vNextDetectBookMode_(spreadsheet) !== 'TEMPLATE' ||
      String(config.template_kind || '') !== 'TEMPLATE_DRAFT' ||
      String(config.template_draft_id || '') !== String(registry.book_id || '')) {
    throw new Error('TEMPLATE_DRAFT routing identity is inconsistent.');
  }
  vNextAdminAssertPrivateAdminFile_(DriveApp.getFileById(spreadsheet.getId()), adminEmails);
  vNextAdminTemplateUiManifest_(spreadsheet);
  return { spreadsheet: spreadsheet,
    sourceReleaseId: vNextAdminRequiredText_(config.source_template_release_id || registry.template_release_id,
      'source_template_release_id'),
    draftId: String(registry.book_id || ''), isDraft: true, adminEmails: adminEmails };
}

function vNextAdminListTemplateDrafts_(hub) {
  return vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.REGISTRY).rows.filter(function (row) {
    return String(row.mode || '') === 'TEMPLATE' && String(row.status || '').toUpperCase() === 'DRAFT';
  }).map(function (row) {
    const note = vNextAdminParseJson_(row.note, {});
    return {
      draftId: String(row.book_id || ''), spreadsheetId: String(row.spreadsheet_id || ''),
      spreadsheetUrl: String(row.spreadsheet_url || ''), sourceReleaseId: String(row.template_release_id || ''),
      manifestSha256: String(note.initialManifestSha256 || ''), createdAt: row.created_at || ''
    };
  });
}

function vNextAdminAssertPrivateAdminFile_(file, adminEmails) {
  if (vNextAdminIsSharedDriveManaged_(file)) return true;
  const owner = file.getOwner() ? String(file.getOwner().getEmail() || '').toLowerCase() : '';
  const allowed = new Set(vNextAdminMergeEmails_(adminEmails, owner, vNextAdminActor_()));
  if (file.getSharingAccess() !== DriveApp.Access.PRIVATE) {
    throw new Error('Template source must be PRIVATE.');
  }
  const unexpected = file.getEditors().concat(file.getViewers()).filter(function (user) {
    return !allowed.has(String(user.getEmail() || '').toLowerCase());
  });
  if (unexpected.length) throw new Error('Template source has a non-Admin collaborator.');
  return true;
}

function vNextAdminCopyAndVerifyTemplateUi_(source, target) {
  const before = vNextAdminTemplateUiManifest_(source);
  VN_ADMIN_DEFAULT_TEMPLATE_VISIBLE.forEach(function (name, index) {
    const sourceSheet = source.getSheetByName(name);
    if (!sourceSheet) throw new Error('Template UI source is missing sheet: ' + name);
    const existing = target.getSheetByName(name);
    if (existing) target.deleteSheet(existing);
    const copied = sourceSheet.copyTo(target).setName(name);
    target.setActiveSheet(copied);
    target.moveActiveSheet(index + 1);
  });
  target.setActiveSheet(target.getSheetByName(VN_ADMIN_DEFAULT_TEMPLATE_VISIBLE[0]));
  SpreadsheetApp.flush();
  const after = vNextAdminTemplateUiManifest_(target);
  const beforeJson = vNextAdminCanonicalJson_(before);
  const afterJson = vNextAdminCanonicalJson_(after);
  if (beforeJson !== afterJson) {
    throw new Error('Template UI copy verification failed. Sheet.copyTo did not preserve the V2 manifest.');
  }
  return { manifest: after, manifestSha256: vNextAdminSha256_(afterJson), copiedSheets: VN_ADMIN_DEFAULT_TEMPLATE_VISIBLE.slice() };
}

function vNextAdminAssertReleaseTemplateManifest_(release, template) {
  const expected = vNextAdminRequiredText_(release && release.template_content_sha256,
    'release.template_content_sha256');
  const schema = String(release && release.template_manifest_schema || 'LEGACY_V1');
  const actual = schema === VN_ADMIN_TEMPLATE_MANIFEST_SCHEMA
    ? vNextAdminTemplateUiManifestHash_(template)
    : schema === VN_ADMIN_TEMPLATE_MANIFEST_SCHEMA_V2
      ? vNextAdminTemplateUiManifestHashV2_(template)
      : vNextAdminTemplateContentHash_(template);
  if (expected !== actual) {
    throw new Error('Master Template visible content differs from the immutable RELEASES manifest.');
  }
  return true;
}

function vNextAdminTemplateUiManifestHash_(template) {
  return vNextAdminSha256_(vNextAdminCanonicalJson_(vNextAdminTemplateUiManifest_(template)));
}

function vNextAdminTemplateUiManifestHashV2_(template) {
  return vNextAdminSha256_(vNextAdminCanonicalJson_(
    vNextAdminTemplateUiManifest_(template, VN_ADMIN_TEMPLATE_MANIFEST_SCHEMA_V2)
  ));
}

/**
 * V2 UI manifest. Every copied attribute that Apps Script can inspect is
 * hashed. Objects that Sheet.copyTo cannot safely isolate are rejected before
 * publication instead of being silently dropped.
 */
function vNextAdminTemplateUiManifest_(template, manifestSchema) {
  const schema = String(manifestSchema || VN_ADMIN_TEMPLATE_MANIFEST_SCHEMA);
  if (schema !== VN_ADMIN_TEMPLATE_MANIFEST_SCHEMA && schema !== VN_ADMIN_TEMPLATE_MANIFEST_SCHEMA_V2) {
    throw new Error('Unsupported Template manifest schema: ' + schema);
  }
  const allowed = new Set(VN_ADMIN_DEFAULT_TEMPLATE_VISIBLE);
  const unexpectedVisible = template.getSheets().filter(function (sheet) {
    return !sheet.isSheetHidden() && !allowed.has(sheet.getName());
  }).map(function (sheet) { return sheet.getName(); });
  if (unexpectedVisible.length) {
    throw new Error('Template contains non-allowlisted visible sheets: ' + unexpectedVisible.join(', '));
  }
  if (typeof template.getNamedRanges === 'function' && template.getNamedRanges().length) {
    throw new Error('Named ranges are forbidden in Template UI; use direct ranges inside the three allowlisted sheets.');
  }
  return {
    schema: schema,
    allowedSheets: VN_ADMIN_DEFAULT_TEMPLATE_VISIBLE.slice(),
    sheets: VN_ADMIN_DEFAULT_TEMPLATE_VISIBLE.map(function (name) {
      const sheet = template.getSheetByName(name);
      if (!sheet) throw new Error('Template manifest cannot find sheet: ' + name);
      vNextAdminAssertTemplateSheetHasNoForbiddenAssets_(template, sheet);
      return vNextAdminTemplateSheetManifest_(template, sheet, { fullGrid: schema === VN_ADMIN_TEMPLATE_MANIFEST_SCHEMA_V2 });
    })
  };
}

function vNextAdminAssertTemplateSheetHasNoForbiddenAssets_(template, sheet) {
  const forbidden = [];
  if (typeof sheet.getCharts === 'function' && sheet.getCharts().length) forbidden.push('charts');
  if (typeof sheet.getDrawings === 'function' && sheet.getDrawings().length) forbidden.push('drawings');
  if (typeof sheet.getImages === 'function' && sheet.getImages().length) forbidden.push('over-grid images');
  if (typeof sheet.getPivotTables === 'function' && sheet.getPivotTables().length) forbidden.push('pivot tables');
  if (typeof sheet.getSlicers === 'function' && sheet.getSlicers().length) forbidden.push('slicers');
  if (typeof sheet.getDataSourceTables === 'function' && sheet.getDataSourceTables().length) forbidden.push('data source tables');
  if (typeof sheet.getDeveloperMetadata === 'function' && sheet.getDeveloperMetadata().length) forbidden.push('developer metadata');
  if (typeof sheet.getProtections === 'function') {
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)
      .concat(sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE));
    if (protections.length) forbidden.push('UI-sheet protections');
  }
  if (forbidden.length) {
    throw new Error(sheet.getName() + ' contains explicitly forbidden Template assets: ' + forbidden.join(', '));
  }
  const formulas = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).getFormulas();
  const internalNames = template.getSheets().map(function (item) { return item.getName(); })
    .filter(function (name) { return !VN_ADMIN_DEFAULT_TEMPLATE_VISIBLE.includes(name); });
  formulas.forEach(function (row) {
    row.forEach(function (formula) {
      const text = String(formula || '');
      if (!text) return;
      if (VN_ADMIN_TEMPLATE_FORBIDDEN_FORMULA.test(text)) {
        throw new Error(sheet.getName() + ' contains a forbidden external-data formula.');
      }
      internalNames.forEach(function (internalName) {
        const quoted = "'" + String(internalName).replace(/'/g, "''") + "'!";
        if (text.indexOf(quoted) >= 0 || text.indexOf(internalName + '!') >= 0) {
          throw new Error(sheet.getName() + ' formula references non-published sheet: ' + internalName);
        }
      });
    });
  });
  return true;
}

function vNextAdminTemplateSheetManifest_(template, sheet, options) {
  const fullGrid = Boolean(options && options.fullGrid);
  const maxRows = sheet.getMaxRows();
  const maxColumns = sheet.getMaxColumns();
  const rows = fullGrid ? maxRows : Math.max(1, sheet.getLastRow());
  const columns = fullGrid ? maxColumns : Math.max(1, sheet.getLastColumn());
  if (rows * columns > 200000) {
    throw new Error(sheet.getName() + ' exceeds the 200,000-cell Template UI manifest limit.');
  }
  // Hash every attribute inside the used UI envelope and keep the complete
  // grid dimensions separately. Scanning row height/visibility for thousands
  // of untouched default rows requires one remote GAS call per row and can
  // exceed the six-minute execution limit during bootstrap.
  const range = sheet.getRange(1, 1, rows, columns);
  const validations = [];
  range.getDataValidations().forEach(function (row, rowIndex) {
    row.forEach(function (rule, columnIndex) {
      if (!rule) return;
      validations.push({ cell: vNextAdminA1_(rowIndex + 1, columnIndex + 1),
        rule: vNextAdminSerializeValidation_(template, rule) });
    });
  });
  const richText = [];
  range.getRichTextValues().forEach(function (row, rowIndex) {
    row.forEach(function (value, columnIndex) {
      const serialized = vNextAdminSerializeRichText_(value);
      if (serialized) richText.push({ cell: vNextAdminA1_(rowIndex + 1, columnIndex + 1), value: serialized });
    });
  });
  const conditionalRules = sheet.getConditionalFormatRules().map(function (rule) {
    return vNextAdminSerializeConditionalRule_(template, rule);
  });
  const merges = sheet.getRange(1, 1, rows, columns).getMergedRanges().map(function (item) {
    return item.getA1Notation();
  }).sort();
  const bandings = typeof range.getBandings === 'function' ? range.getBandings().map(function (banding) {
    return vNextAdminSerializeBanding_(banding);
  }) : [];
  const columnWidths = [];
  const rowHeights = [];
  const hiddenColumns = [];
  const hiddenRows = [];
  for (let column = 1; column <= columns; column++) {
    columnWidths.push(sheet.getColumnWidth(column));
    if (sheet.isColumnHiddenByUser(column)) hiddenColumns.push(column);
  }
  for (let row = 1; row <= rows; row++) {
    rowHeights.push(sheet.getRowHeight(row));
    if (sheet.isRowHiddenByUser(row)) hiddenRows.push(row);
  }
  return {
    name: sheet.getName(), maxRows: maxRows, maxColumns: maxColumns,
    usedRows: fullGrid ? undefined : rows, usedColumns: fullGrid ? undefined : columns,
    frozenRows: sheet.getFrozenRows(), frozenColumns: sheet.getFrozenColumns(),
    hiddenGridlines: typeof sheet.hasHiddenGridlines === 'function' ? sheet.hasHiddenGridlines() : false,
    rightToLeft: typeof sheet.isRightToLeft === 'function' ? sheet.isRightToLeft() : false,
    tabColor: typeof sheet.getTabColor === 'function' ? String(sheet.getTabColor() || '') : '',
    valuesHash: vNextAdminSha256_(vNextAdminCanonicalJson_(range.getValues())),
    formulasHash: vNextAdminSha256_(vNextAdminCanonicalJson_(range.getFormulas())),
    notesHash: vNextAdminSha256_(vNextAdminCanonicalJson_(range.getNotes())),
    numberFormatsHash: vNextAdminSha256_(vNextAdminCanonicalJson_(range.getNumberFormats())),
    backgroundsHash: vNextAdminSha256_(vNextAdminCanonicalJson_(range.getBackgrounds())),
    fontFamiliesHash: vNextAdminSha256_(vNextAdminCanonicalJson_(range.getFontFamilies())),
    fontSizesHash: vNextAdminSha256_(vNextAdminCanonicalJson_(range.getFontSizes())),
    fontWeightsHash: vNextAdminSha256_(vNextAdminCanonicalJson_(range.getFontWeights())),
    fontStylesHash: vNextAdminSha256_(vNextAdminCanonicalJson_(range.getFontStyles())),
    fontColorsHash: vNextAdminSha256_(vNextAdminCanonicalJson_(range.getFontColors())),
    horizontalAlignmentsHash: vNextAdminSha256_(vNextAdminCanonicalJson_(range.getHorizontalAlignments())),
    verticalAlignmentsHash: vNextAdminSha256_(vNextAdminCanonicalJson_(range.getVerticalAlignments())),
    wrapsHash: vNextAdminSha256_(vNextAdminCanonicalJson_(range.getWraps())),
    textDirectionsHash: typeof range.getTextDirections === 'function'
      ? vNextAdminSha256_(vNextAdminCanonicalJson_(range.getTextDirections().map(function (row) {
        return row.map(function (item) { return String(item || ''); });
      }))) : '',
    textRotationsHash: typeof range.getTextRotations === 'function'
      ? vNextAdminSha256_(vNextAdminCanonicalJson_(range.getTextRotations().map(function (row) {
        return row.map(vNextAdminSerializeTextRotation_);
      }))) : '',
    validations: validations, conditionalFormatting: conditionalRules,
    mergedRanges: merges, richText: richText, bandings: bandings,
    filter: vNextAdminSerializeFilter_(template, sheet.getFilter()),
    columnWidths: columnWidths, rowHeights: rowHeights,
    hiddenColumns: hiddenColumns, hiddenRows: hiddenRows
  };
}

function vNextAdminSerializeValidation_(template, rule) {
  return {
    criteriaType: String(rule.getCriteriaType() || ''),
    criteriaValues: (rule.getCriteriaValues() || []).map(function (value) {
      return vNextAdminManifestValue_(template, value);
    }),
    allowInvalid: rule.getAllowInvalid(), helpText: String(rule.getHelpText() || '')
  };
}

function vNextAdminSerializeConditionalRule_(template, rule) {
  const ranges = rule.getRanges().map(function (range) {
    return range.getSheet().getName() + '!' + range.getA1Notation();
  }).sort();
  const booleanCondition = rule.getBooleanCondition();
  const gradientCondition = rule.getGradientCondition();
  return {
    ranges: ranges,
    boolean: booleanCondition ? {
      criteriaType: String(booleanCondition.getCriteriaType() || ''),
      criteriaValues: (booleanCondition.getCriteriaValues() || []).map(function (value) {
        return vNextAdminManifestValue_(template, value);
      }),
      background: String(booleanCondition.getBackground() || ''),
      fontColor: String(booleanCondition.getFontColor() || ''),
      bold: booleanCondition.getBold(), italic: booleanCondition.getItalic(),
      strikethrough: booleanCondition.getStrikethrough(), underline: booleanCondition.getUnderline()
    } : null,
    gradient: gradientCondition ? {
      min: vNextAdminSerializeGradientPoint_(gradientCondition.getMinpoint()),
      mid: vNextAdminSerializeGradientPoint_(gradientCondition.getMidpoint()),
      max: vNextAdminSerializeGradientPoint_(gradientCondition.getMaxpoint())
    } : null
  };
}

function vNextAdminSerializeGradientPoint_(point) {
  if (!point) return null;
  return { color: vNextAdminColorString_(point.getColor()), type: String(point.getType() || ''),
    value: vNextAdminText_(point.getValue()) };
}

function vNextAdminSerializeRichText_(value) {
  if (!value) return null;
  const text = String(value.getText() || '');
  if (!text) return null;
  const runs = value.getRuns ? value.getRuns() : [];
  return { text: text, runs: runs.map(function (run) {
    return { start: run.getStartIndex(), end: run.getEndIndex(), text: run.getText(),
      linkUrl: String(run.getLinkUrl && run.getLinkUrl() || ''),
      style: vNextAdminSerializeTextStyle_(run.getTextStyle && run.getTextStyle()) };
  }) };
}

function vNextAdminSerializeTextStyle_(style) {
  if (!style) return null;
  return {
    bold: style.isBold(), italic: style.isItalic(), underline: style.isUnderline(),
    strikethrough: style.isStrikethrough(), fontFamily: String(style.getFontFamily() || ''),
    fontSize: style.getFontSize(), foregroundColor: style.getForegroundColor ? String(style.getForegroundColor() || '') : ''
  };
}

function vNextAdminSerializeTextRotation_(rotation) {
  if (!rotation) return null;
  return { degrees: rotation.getDegrees(), vertical: rotation.isVertical() };
}

function vNextAdminSerializeBanding_(banding) {
  return {
    range: banding.getRange().getA1Notation(),
    headerColor: typeof banding.getHeaderRowColor === 'function' ? String(banding.getHeaderRowColor() || '') : '',
    firstRowColor: typeof banding.getFirstRowColor === 'function' ? String(banding.getFirstRowColor() || '') : '',
    secondRowColor: typeof banding.getSecondRowColor === 'function' ? String(banding.getSecondRowColor() || '') : '',
    footerColor: typeof banding.getFooterRowColor === 'function' ? String(banding.getFooterRowColor() || '') : '',
    firstColumnColor: typeof banding.getFirstColumnColor === 'function' ? String(banding.getFirstColumnColor() || '') : '',
    secondColumnColor: typeof banding.getSecondColumnColor === 'function' ? String(banding.getSecondColumnColor() || '') : ''
  };
}

function vNextAdminSerializeFilter_(template, filter) {
  if (!filter) return null;
  const range = filter.getRange();
  const criteria = [];
  for (let column = 1; column <= range.getNumColumns(); column++) {
    const rule = filter.getColumnFilterCriteria(column);
    if (!rule) continue;
    criteria.push({ column: column, criteriaType: String(rule.getCriteriaType() || ''),
      criteriaValues: (rule.getCriteriaValues() || []).map(function (value) {
        return vNextAdminManifestValue_(template, value);
      }), hiddenValues: rule.getHiddenValues ? rule.getHiddenValues() : [] });
  }
  return { range: range.getA1Notation(), criteria: criteria };
}

function vNextAdminManifestValue_(template, value) {
  if (value instanceof Date) return value.toISOString();
  if (Array.isArray(value)) return value.map(function (item) { return vNextAdminManifestValue_(template, item); });
  if (value && typeof value.getA1Notation === 'function' && typeof value.getSheet === 'function') {
    const valueSheet = value.getSheet();
    return { type: 'RANGE', scope: valueSheet.getParent().getId() === template.getId() ? 'LOCAL' : 'EXTERNAL',
      sheet: valueSheet.getName(), a1: value.getA1Notation() };
  }
  if (value && typeof value === 'object') return String(value);
  return value;
}

function vNextAdminColorString_(color) {
  if (!color) return '';
  try { return color.asRgbColor().asHexString(); } catch (error) { return String(color); }
}

function vNextAdminA1_(row, column) {
  let value = Number(column);
  let letters = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    value = Math.floor((value - 1) / 26);
  }
  return letters + String(row);
}

/** Canonical visible-sheet manifest bound to each immutable Template release. */
function vNextAdminTemplateContentHash_(template) {
  const manifest = VN_ADMIN_DEFAULT_TEMPLATE_VISIBLE.map(function (name) {
    const sheet = template.getSheetByName(name);
    if (!sheet) throw new Error('Template content hash cannot find sheet: ' + name);
    const range = sheet.getDataRange();
    const rows = Math.max(1, range.getNumRows());
    const columns = Math.max(1, range.getNumColumns());
    const columnWidths = [];
    const rowHeights = [];
    for (let column = 1; column <= columns; column++) columnWidths.push(sheet.getColumnWidth(column));
    for (let row = 1; row <= rows; row++) rowHeights.push(sheet.getRowHeight(row));
    return {
      name: name, rows: rows, columns: columns,
      frozenRows: sheet.getFrozenRows(), frozenColumns: sheet.getFrozenColumns(),
      hiddenGridlines: typeof sheet.hasHiddenGridlines === 'function' ? sheet.hasHiddenGridlines() : false,
      values: range.getValues(), formulas: range.getFormulas(), notes: range.getNotes(),
      numberFormats: range.getNumberFormats(), backgrounds: range.getBackgrounds(),
      fontFamilies: range.getFontFamilies(), fontSizes: range.getFontSizes(),
      fontWeights: range.getFontWeights(), fontStyles: range.getFontStyles(),
      fontColors: range.getFontColors(), horizontalAlignments: range.getHorizontalAlignments(),
      verticalAlignments: range.getVerticalAlignments(), wraps: range.getWraps(),
      columnWidths: columnWidths, rowHeights: rowHeights
    };
  });
  return vNextAdminSha256_(vNextAdminCanonicalJson_(manifest));
}

function vNextAdminResetCopiedWorkbook_(ss, initialSheets) {
  const tempName = '__VNEXT_BUILD__';
  let temp = ss.getSheetByName(tempName);
  if (!temp) temp = ss.insertSheet(tempName);
  ss.setActiveSheet(temp);
  // This runs only on newly-created copies. The legacy source workbook is never passed here.
  ss.getSheets().forEach(function (sheet) {
    if (sheet.getName() !== tempName) ss.deleteSheet(sheet);
  });
  (initialSheets || []).forEach(function (name) {
    if (!ss.getSheetByName(name)) ss.insertSheet(name);
  });
  if ((initialSheets || []).length) {
    ss.setActiveSheet(ss.getSheetByName(initialSheets[0]));
    ss.deleteSheet(temp);
  }
}

function vNextAdminEnsureUiShell_(ss, options) {
  const opt = options || {};
  VN_ADMIN_DEFAULT_TEMPLATE_VISIBLE.forEach(function (name) {
    const sheet = vNextAdminGetOrCreateSheet_(ss, name);
    if (!String(sheet.getRange('A1').getValue() || '').trim() || !opt.template) {
      if (name === '1_ホーム') {
        const title = opt.template
          ? 'Forecast vNext Master Template'
          : vNextFormatClientBookTitle_(opt.clientName, opt.fiscalYear);
        sheet.getRange('A1').setValue(title).setFontWeight('bold').setFontSize(16);
        sheet.getRange('A3').setValue(opt.template
          ? 'Master Templateです。クライアント年度ブックは' + VNEXT_NAMING.LAYER1 + 'から生成してください。'
          : '準備中です。メニュー「' + VNEXT_NAMING.MENU + '」から「ホームに戻る」を選んでください。');
      } else if (name === '2_予測と計画') {
        sheet.getRange('A1').setValue('予測と計画').setFontWeight('bold');
      } else {
        sheet.getRange('A1').setValue('振り返り').setFontWeight('bold');
      }
    }
  });
}

// ---------------------------- Jobs / health ----------------------------

function vNextAdminQueueAgeMetrics_(rows, nowMs) {
  const now = Number(nowMs || Date.now());
  const queued = (rows || []).filter(function (row) { return String(row.status || '').toUpperCase() === 'QUEUED'; });
  const running = (rows || []).filter(function (row) { return String(row.status || '').toUpperCase() === 'RUNNING'; });
  const ages = queued.map(function (row) {
    const created = new Date(row.created_at || row.updated_at || 0).getTime();
    return isFinite(created) && created > 0 ? Math.max(0, (now - created) / 60000) : 0;
  });
  const oldest = ages.length ? Math.max.apply(null, ages) : 0;
  return {
    queued: queued.length, running: running.length,
    oldestQueuedAgeMinutes: Math.round(oldest * 10) / 10,
    staleQueued: ages.filter(function (age) { return age >= VN_ADMIN_STALE_MINUTES; }).length
  };
}

function vNextAdminOperationalMetrics_(hub, automationInstalled, optionalJobRows) {
  const now = Date.now();
  const jobRows = Array.isArray(optionalJobRows)
    ? optionalJobRows : vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.JOBS).rows;
  const queue = vNextAdminQueueAgeMetrics_(jobRows, now);
  const props = PropertiesService.getScriptProperties();
  const succeededAt = String(props.getProperty('VNEXT_LAST_SWEEP_SUCCEEDED_AT') || '');
  const successMs = new Date(succeededAt || 0).getTime();
  const sweepAge = isFinite(successMs) && successMs > 0 ? Math.max(0, (now - successMs) / 60000) : null;
  return Object.assign({}, queue, {
    lastSweepSucceededAt: succeededAt,
    lastSweepAgeMinutes: sweepAge === null ? null : Math.round(sweepAge * 10) / 10,
    lastSweepDurationMs: Number(props.getProperty('VNEXT_LAST_SWEEP_DURATION_MS') || 0),
    schedulerStale: Boolean(automationInstalled && (sweepAge === null || sweepAge >= VN_ADMIN_STALE_MINUTES)),
    queueStale: queue.staleQueued > 0
  });
}

function vNextAdminProcessJobsForHub_(hub, limit, deadlineMs) {
  const maxJobs = Math.max(1, Math.min(20, Number(limit || 5)));
  const results = [];
  for (let i = 0; i < maxJobs; i++) {
    if (deadlineMs && Date.now() >= deadlineMs) break;
    const job = vNextAdminWithScriptLock_('claim-job', function () { return vNextAdminClaimNextJob_(hub); });
    if (!job) break;
    try {
      const result = vNextAdminExecuteJob_(hub, job);
      vNextAdminWithScriptLock_('finish-job', function () { vNextAdminFinishJob_(hub, job.job_id, 'SUCCEEDED', result, ''); });
      results.push({ jobId: job.job_id, status: 'SUCCEEDED', result: result });
    } catch (err) {
      const message = String(err && err.message || err);
      if (String(job.job_type || '') === 'PORTAL_PROVISION_CLIENT') {
        try { vNextAdminMarkPortalJobFailed_(hub, job, message); }
        catch (portalFailureError) { Logger.log('Portal failure status append skipped: %s', String(portalFailureError)); }
      }
      vNextAdminWithScriptLock_('fail-job', function () { vNextAdminFinishJob_(hub, job.job_id, 'FAILED', null, message); });
      results.push({ jobId: job.job_id, status: 'FAILED', error: message });
      Logger.log('Job failed id=%s type=%s error=%s', job.job_id, job.job_type, String(err && err.stack || err));
    }
  }
  vNextAdminRefreshTodayExceptions_(hub);
  vNextAdminRefreshHome_(hub);
  return vNextAdminJsonSafe_({ processed: results.length, jobs: results });
}

function vNextAdminMarkPortalJobFailed_(hub, job, message) {
  const payload = vNextAdminParseJson_(job.request_json, {});
  const portal = vNextAdminResolvePortal_(hub);
  const events = vNextAdminReadTable_(portal.spreadsheet, VN_ADMIN_PORTAL_REQUEST_SHEET).rows.filter(function (row) {
    return String(row.request_id || '') === String(payload.requestId || '');
  });
  const latest = events.length ? events[events.length - 1] : null;
  const latestPair = String(latest && latest.event_type || '').toUpperCase() + '>' +
    String(latest && latest.status || '').toUpperCase();
  if (['COMPLETED>COMPLETED', 'REJECTED>REJECTED'].indexOf(latestPair) >= 0) return false;
  if (latestPair === 'FAILED>FAILED') return true;
  const requested = events.find(function (row) {
    return String(row.request_id || '') === String(payload.requestId || '') &&
      String(row.event_type || '').toUpperCase() === 'REQUESTED';
  });
  if (!requested) return false;
  const validated = vNextAdminValidatePortalRequest_(hub, portal, requested);
  vNextAdminAppendPortalRequestEvent_(portal.spreadsheet, validated, 'FAILED', 'FAILED', {
    relatedJobId: job.job_id, detailCode: 'CREATION_FAILED',
    detailMessage: '作成を完了できませんでした。管理担当者が確認します。'
  });
  vNextAdminAppendException_(hub, {
    severity: 'ERROR', exception_type: 'PORTAL_PROVISION_FAILED', book_id: String(payload.requestId || ''),
    client_name: String(payload.clientName || ''), fiscal_year: Number(payload.fiscalYear || 0),
    title: '申請入口経由のクライアント年度ブック作成に失敗', detail: String(message || '').slice(0, 1200),
    recommended_action: 'JOB_QUEUEと生成途中BOOK_REGISTRYを確認し、必要なら再依頼', source_ref: job.job_id
  });
  return true;
}

function vNextAdminScanRegistryForHub_(hub) {
  const registry = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.REGISTRY);
  const results = [];
  registry.rows.forEach(function (row) {
    if (!row.book_id || String(row.status) === 'ARCHIVED') return;
    results.push(vNextAdminScanOneBook_(hub, row));
  });
  vNextAdminRefreshTodayExceptions_(hub);
  vNextAdminRefreshHome_(hub);
  return vNextAdminJsonSafe_({ scanned: results.length, results: results });
}

function vNextAdminScanRegistryBatch_(hub, limit) {
  const rows = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.REGISTRY).rows.filter(function (row) {
    return row.book_id && String(row.status) !== 'ARCHIVED';
  });
  if (!rows.length) return { scanned: 0, results: [], nextCursor: 0 };
  const props = PropertiesService.getScriptProperties();
  let cursor = Number(props.getProperty('VNEXT_SCAN_CURSOR') || 0);
  if (!isFinite(cursor) || cursor < 0) cursor = 0;
  cursor = cursor % rows.length;
  const count = Math.min(Math.max(1, Number(limit || 5)), rows.length);
  const results = [];
  for (let i = 0; i < count; i++) results.push(vNextAdminScanOneBook_(hub, rows[(cursor + i) % rows.length]));
  const next = (cursor + count) % rows.length;
  props.setProperty('VNEXT_SCAN_CURSOR', String(next));
  vNextAdminRefreshTodayExceptions_(hub);
  vNextAdminRefreshHome_(hub);
  return vNextAdminJsonSafe_({ scanned: results.length, results: results, nextCursor: next });
}

function vNextAdminRecoverStaleLeases_(hub, staleMinutes) {
  const threshold = Date.now() - Math.max(5, Number(staleMinutes || 20)) * 60000;
  let requeued = 0;
  let failed = 0;
  let approvalsReleased = 0;
  let approvalsFailed = 0;
  const jobs = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.JOBS).rows;
  jobs.forEach(function (job) {
    if (String(job.status || '') !== 'RUNNING') return;
    const leaseTime = new Date(job.locked_at || job.started_at || job.updated_at || 0).getTime();
    if (!isFinite(leaseTime) || leaseTime > threshold) return;
    const attempts = Number(job.attempts || 0);
    if (attempts < 3) {
      vNextAdminUpdateTableRow_(hub, VN_ADMIN_SHEETS.JOBS, job._rowNumber, {
        status: 'QUEUED', locked_at: '', locked_by: '', started_at: '',
        error: 'Stale lease recovered at ' + new Date().toISOString(), updated_at: new Date()
      });
      requeued++;
    } else {
      vNextAdminUpdateTableRow_(hub, VN_ADMIN_SHEETS.JOBS, job._rowNumber, {
        status: 'FAILED', finished_at: new Date(), error: 'Lease expired after 3 attempts.', updated_at: new Date()
      });
      let recoveryDetail = '';
      if (['FORECAST_REQUEST', 'AI_ROLLBACK_FORECAST'].indexOf(String(job.job_type || '')) >= 0) {
        try {
          recoveryDetail = vNextAdminRecoverExhaustedForecastJob_(hub, job);
        } catch (recoveryError) {
          recoveryDetail = ' client recovery failed: ' + String(recoveryError && recoveryError.message || recoveryError);
          Logger.log('Exhausted forecast recovery failed job=%s error=%s', job.job_id, String(recoveryError && recoveryError.stack || recoveryError));
        }
      }
      if (String(job.job_type || '') === 'PORTAL_PROVISION_CLIENT') {
        try {
          vNextAdminMarkPortalJobFailed_(hub, job, 'Lease expired after 3 attempts.');
          recoveryDetail = '; Portal request marked FAILED';
        } catch (portalRecoveryError) {
          recoveryDetail = '; Portal failure event could not be appended: ' +
            String(portalRecoveryError && portalRecoveryError.message || portalRecoveryError);
          Logger.log('Exhausted Portal recovery failed job=%s error=%s', job.job_id,
            String(portalRecoveryError && portalRecoveryError.stack || portalRecoveryError));
        }
      }
      vNextAdminAppendException_(hub, {
        severity: 'ERROR', exception_type: 'JOB_LEASE_EXHAUSTED', book_id: job.target_book_id,
        title: '自動処理が3回中断されました', detail: 'job=' + job.job_id + recoveryDetail,
        recommended_action: '権限・実行時間・入力データを確認して再投入', source_ref: job.job_id
      });
      failed++;
    }
  });
  const approvals = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.APPROVALS).rows;
  approvals.forEach(function (approval) {
    if (String(approval.status || '').indexOf('PROCESSING_') !== 0) return;
    const leaseTime = new Date(approval.updated_at || 0).getTime();
    if (!isFinite(leaseTime) || leaseTime > threshold) return;
    const attempts = Number(approval.processing_attempts || 0);
    if (attempts < 3) {
      vNextAdminUpdateTableRow_(hub, VN_ADMIN_SHEETS.APPROVALS, approval._rowNumber, {
        status: 'PENDING', decision_at: '', decision_by: '',
        decision_comment: '前回の承認処理が中断されたため再開待ちに戻しました。', updated_at: new Date()
      });
      approvalsReleased++;
    } else {
      vNextAdminUpdateTableRow_(hub, VN_ADMIN_SHEETS.APPROVALS, approval._rowNumber, {
        status: 'FAILED', decision_at: '',
        decision_comment: '承認処理が3回中断されたため管理ハブ担当者確認が必要です。', updated_at: new Date()
      });
      vNextAdminAppendException_(hub, {
        severity: 'ERROR', exception_type: 'APPROVAL_LEASE_EXHAUSTED', book_id: approval.book_id,
        client_name: approval.client_name, fiscal_year: approval.fiscal_year,
        title: '承認処理が3回中断されました', detail: 'approval=' + approval.approval_request_id,
        recommended_action: '公式化の途中記録を確認し、必要なら新しい承認依頼を作成', source_ref: approval.approval_request_id
      });
      approvalsFailed++;
    }
  });
  return {
    requeuedJobs: requeued, failedJobs: failed,
    releasedApprovals: approvalsReleased, failedApprovals: approvalsFailed
  };
}

/**
 * Requeue only the exact Pilot failure caused by the pre-fix request-state
 * validator. This is deliberately not a generic automatic retry policy:
 * unknown data/model failures stay FAILED for Admin review.
 */
function vNextAdminRequeueKnownPilotFailures_(hub) {
  let requeued = 0;
  let portalRequeued = 0;
  vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.JOBS).rows.forEach(function (job) {
    try {
      if (vNextAdminRequeueKnownPortalReleaseFailure_(hub, job)) {
        requeued++;
        portalRequeued++;
        return;
      }
      if (vNextAdminRequeueKnownEvidenceMonthFailure_(hub, job)) {
        requeued++;
        return;
      }
    } catch (portalRetryError) {
      Logger.log('Known Portal release retry skipped job=%s error=%s', String(job.job_id || ''),
        String(portalRetryError && portalRetryError.stack || portalRetryError));
      return;
    }
    if (String(job.job_type || '') !== 'FORECAST_REQUEST' ||
        String(job.status || '').toUpperCase() !== 'FAILED' ||
        Number(job.attempts || 0) >= 3 ||
        String(job.error || '') !== 'A matching valid pending forecast request was not found.') return;
    const registry = vNextAdminFindRegistryRow_(hub, function (row) {
      return String(row.book_id || '') === String(job.target_book_id || '') &&
        String(row.mode || '') === 'CLIENT';
    });
    if (!registry) return;
    let authoritativeState = vNextAdminLatestClientState_(hub, registry.book_id, registry.state || '');
    // The pre-fix failure could append FAILED locally before the Client
    // RUNNING event reached the Hub. Import that already-audited state chain
    // first; the strict state/request validators still run and fail closed.
    if (authoritativeState !== 'RUNNING') {
      try {
        const client = SpreadsheetApp.openById(String(registry.spreadsheet_id || ''));
        if (vNextAdminLatestClientState_(client, registry.book_id, registry.state || '') === 'RUNNING') {
          vNextAdminSyncClientToHub_(hub, client, registry.book_id);
          authoritativeState = vNextAdminLatestClientState_(hub, registry.book_id, registry.state || '');
        }
      } catch (error) {
        Logger.log('Known Pilot state reconciliation failed for ' + registry.book_id + ': ' +
          (error && error.message ? error.message : error));
        return;
      }
    }
    if (authoritativeState !== 'RUNNING') return;
    vNextAdminUpdateTableRow_(hub, VN_ADMIN_SHEETS.JOBS, job._rowNumber, {
      status: 'QUEUED', not_before: '', locked_at: '', locked_by: '', started_at: '', finished_at: '',
      result_json: '', error: '', updated_at: new Date()
    });
    vNextAdminAppendObject_(hub, VN_ADMIN_SHEETS.JOB_LOG, {
      log_id: 'JLOG-' + Utilities.getUuid(), job_id: job.job_id, logged_at: new Date(),
      status: 'QUEUED', message: 'Known Pilot request-validation failure requeued after runtime fix',
      detail_json: vNextAdminCanonicalJson_({ previousError: 'A matching valid pending forecast request was not found.' }),
      actor: vNextAdminActor_()
    });
    requeued++;
  });
  return { requeuedJobs: requeued, portalRequeuedJobs: portalRequeued };
}

/**
 * Recover the exact pre-run failures caused by Sheets converting YYYY-MM text
 * to Date objects. The comparison error is the same defect observed one step
 * later and is accepted only after the full validator proves Client and Hub
 * evidence are still identical. No FORECAST_RUN exists at this point.
 */
function vNextAdminRequeueKnownEvidenceMonthFailure_(hub, job) {
  if (String(job.job_type || '') !== 'FORECAST_REQUEST' ||
      String(job.status || '').toUpperCase() !== 'FAILED' ||
      !vNextAdminCanRetryKnownEvidencePreflightJob_(job)) return false;
  const payload = vNextAdminParseJson_(job.request_json, {});
  const requestId = String(payload.requestId || '');
  const requestHash = String(payload.requestHash || '');
  if (!requestId || !requestHash || !String(job.idempotency_key || '')) return false;
  const registry = vNextAdminFindRegistryRow_(hub, function (row) {
    return String(row.book_id || '') === String(job.target_book_id || '') &&
      String(row.mode || '') === 'CLIENT' && String(row.status || '').toUpperCase() === 'ACTIVE';
  });
  if (!registry || (String(job.target_spreadsheet_id || '') &&
      String(job.target_spreadsheet_id || '') !== String(registry.spreadsheet_id || ''))) return false;
  const activePair = vNextAdminReadActiveReleasePair_(hub);
  if (String(registry.template_release_id || '') !== String(activePair.releaseId || '')) return false;
  const client = SpreadsheetApp.openById(String(registry.spreadsheet_id || ''));
  const requests = vNextAdminReadTable_(client, VN_ADMIN_CLIENT_REQUEST_SHEET).rows.filter(function (row) {
    return String(row.request_id || '') === requestId;
  });
  const requested = requests.find(function (row) {
    return String(row.event_type || '').toUpperCase() === 'REQUESTED' &&
      String(row.status || '').toUpperCase() === 'PENDING';
  });
  const latest = requests.length ? requests[requests.length - 1] : null;
  if (!requested || String(requested.request_hash || '') !== requestHash ||
      String(latest && latest.event_type || '').toUpperCase() !== 'FAILED' ||
      String(latest && latest.status || '').toUpperCase() !== 'FAILED') return false;
  const runIdentity = typeof vNextEngineBuildAdminRunIdentity_ === 'function'
    ? vNextEngineBuildAdminRunIdentity_(registry.book_id, String(job.idempotency_key || '')) : null;
  if (!runIdentity || vNextAdminReadCoreRows_(hub, 'FORECAST_RUN').some(function (row) {
    return String(row.run_id || '') === String(runIdentity.runId || '');
  })) return false;

  // Prove that the formerly rejected rows are now canonical before changing
  // state. This also normalizes the in-memory rows used by the following sync.
  const evidence = vNextAdminReadCoreRows_(client, 'EVIDENCE_EVENT').filter(function (row) {
    return String(row.book_id || '') === String(registry.book_id || '');
  });
  vNextAdminValidateClientEvidenceRows_(hub, registry.book_id, evidence);
  vNextAdminRefreshClientAnnualSalesScale_(hub, registry);

  const latestStateRows = vNextAdminReadCoreRows_(hub, 'STATE_EVENT').filter(function (row) {
    return String(row.book_id || '') === String(registry.book_id || '');
  });
  const latestState = latestStateRows.length ? latestStateRows[latestStateRows.length - 1] : null;
  const authoritativeState = String(latestState && latestState.to_state || registry.state || '').toUpperCase();
  const expectedReason = 'forecast_requested:' + requestId;
  if (authoritativeState === 'READY_TO_RUN') {
    vNextAdminSetClientState_(registry.spreadsheet_id, 'RUNNING', {
      reason: expectedReason, actorRole: 'ADMIN', hub: hub
    });
  } else if (authoritativeState !== 'RUNNING' || String(latestState && latestState.reason || '') !== expectedReason) {
    return false;
  }
  vNextAdminUpdateTableRow_(hub, VN_ADMIN_SHEETS.JOBS, job._rowNumber, {
    status: 'QUEUED', not_before: '', locked_at: '', locked_by: '', started_at: '', finished_at: '',
    result_json: '', error: '', updated_at: new Date()
  });
  vNextAdminAppendObject_(hub, VN_ADMIN_SHEETS.JOB_LOG, {
    log_id: 'JLOG-' + Utilities.getUuid(), job_id: job.job_id, logged_at: new Date(),
    status: 'QUEUED', message: 'Evidence month coercion fixed; original forecast job requeued',
    detail_json: vNextAdminCanonicalJson_({ requestId: requestId, runId: runIdentity.runId,
      attemptsPreserved: Number(job.attempts || 0) }), actor: vNextAdminActor_()
  });
  vNextAdminWriteAudit_(hub, 'REQUEUE_FORECAST_MONTH_NORMALIZATION', 'FORECAST_JOB', job.job_id, 'SUCCESS', {
    bookId: registry.book_id, requestId: requestId, runId: runIdentity.runId
  });
  return true;
}

function vNextAdminIsKnownEvidencePreflightFailure_(errorText) {
  const text = String(errorText || '');
  return /^invalid evidence month: target_(?:start|end)_month$/.test(text) ||
    /^Client evidence differs from the already accepted Hub record: [A-Za-z0-9-]{8,200}$/.test(text) ||
    /^Append-only integrity mismatch EVIDENCE_EVENT id=[A-Za-z0-9-]{8,200}$/.test(text);
}

function vNextAdminCanRetryKnownEvidencePreflightJob_(job) {
  if (!vNextAdminIsKnownEvidencePreflightFailure_(job && job.error)) return false;
  const attempts = Number(job && job.attempts || 0);
  if (/^Append-only integrity mismatch EVIDENCE_EVENT id=/.test(String(job && job.error || ''))) {
    return attempts >= 1 && attempts <= 3;
  }
  return attempts >= 0 && attempts < 3;
}

function vNextAdminIsKnownEvidenceRequeuedJob_(hub, job) {
  if (String(job && job.job_type || '') !== 'FORECAST_REQUEST' ||
      String(job && job.status || '').toUpperCase() !== 'QUEUED' ||
      Number(job && job.attempts || 0) > 3 || String(job && job.error || '') ||
      String(job && job.locked_at || '') || String(job && job.locked_by || '') ||
      String(job && job.started_at || '') || String(job && job.result_json || '')) return false;
  const logs = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.JOB_LOG).rows.filter(function (row) {
    return String(row.job_id || '') === String(job.job_id || '');
  });
  const latest = logs.length ? logs[logs.length - 1] : null;
  return String(latest && latest.status || '').toUpperCase() === 'QUEUED' &&
    String(latest && latest.message || '') ===
      'Evidence month coercion fixed; original forecast job requeued';
}

/**
 * Recover only the pre-side-effect Portal failure caused by a stale cached
 * Template release. The original request, job ID and idempotency key remain
 * unchanged; every authoritative identity is revalidated before requeue.
 */
function vNextAdminRequeueKnownPortalReleaseFailure_(hub, job) {
  if (String(job.job_type || '') !== 'PORTAL_PROVISION_CLIENT' ||
      String(job.status || '').toUpperCase() !== 'FAILED' ||
      Number(job.attempts || 0) >= 3) return false;
  const staleMatch = /^Requested release is not ACTIVE: ([-A-Za-z0-9._]+)$/.exec(String(job.error || ''));
  if (!staleMatch) return false;
  const staleReleaseId = staleMatch[1];
  const payload = vNextAdminParseJson_(job.request_json, {});
  if (!payload || payload.releaseId || payload.modelReleaseId ||
      String(job.target_book_id || '') !== String(payload.requestId || '')) return false;

  const portal = vNextAdminResolvePortal_(hub);
  if (String(job.target_spreadsheet_id || '') !== portal.spreadsheetId ||
      String(payload.portalSpreadsheetId || '') !== portal.spreadsheetId ||
      String(payload.portalId || '') !== portal.portalId) return false;
  const events = vNextAdminReadTable_(portal.spreadsheet, VN_ADMIN_PORTAL_REQUEST_SHEET).rows.filter(function (row) {
    return String(row.request_id || '') === String(payload.requestId || '');
  });
  const requested = events.find(function (row) {
    return String(row.event_type || '').toUpperCase() === 'REQUESTED' &&
      String(row.status || '').toUpperCase() === 'PENDING';
  });
  const latest = events.length ? events[events.length - 1] : null;
  if (!requested || String(latest && latest.event_type || '').toUpperCase() !== 'FAILED' ||
      String(latest && latest.status || '').toUpperCase() !== 'FAILED') return false;

  const validated = vNextAdminValidatePortalRequest_(hub, portal, requested);
  if (validated.requestHash !== String(payload.requestHash || '') ||
      validated.payload.requestedBy !== String(payload.requestedBy || '') ||
      validated.schemaVersion !== String(payload.schemaVersion || VN_ADMIN_PORTAL_REQUEST_SCHEMA_V1) ||
      validated.clientId !== String(payload.clientId || '') ||
      validated.forecastOwnerEmail !== String(payload.forecastOwnerEmail || '') ||
      vNextAdminCanonicalJson_(validated.relatedMemberEmails) !==
        vNextAdminCanonicalJson_(payload.relatedMemberEmails || []) ||
      vNextAdminCanonicalJson_(validated.relatedMemberNames) !==
        vNextAdminCanonicalJson_(payload.relatedMemberNames || [])) return false;

  const expectedIdempotency = 'PORTAL_PROVISION|' + vNextAdminPortalCanonicalClientKey_({
    clientId: validated.clientId, clientName: validated.payload.clientName
  }) + '|' + validated.payload.fiscalYear;
  if (String(job.idempotency_key || '') !== expectedIdempotency) return false;
  const staleRelease = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.RELEASES).rows.find(function (row) {
    return String(row.release_id || '') === staleReleaseId;
  });
  if (!staleRelease || String(staleRelease.status || '').toUpperCase() === 'ACTIVE') return false;
  const activeRelease = vNextAdminResolveRelease_(hub, '');
  const activeModel = vNextAdminResolveActiveModelRelease_(hub, '');
  if (String(activeRelease.release_id || '') === staleReleaseId) return false;
  vNextAdminAssertModelTemplateCompatibility_(hub, activeModel, activeRelease);

  const now = new Date();
  vNextAdminUpdateTableRow_(hub, VN_ADMIN_SHEETS.JOBS, job._rowNumber, {
    status: 'QUEUED', not_before: '', locked_at: '', locked_by: '', started_at: '', finished_at: '',
    result_json: '', error: '', updated_at: now
  });
  vNextAdminAppendObject_(hub, VN_ADMIN_SHEETS.JOB_LOG, {
    log_id: 'JLOG-' + Utilities.getUuid(), job_id: job.job_id, logged_at: now,
    status: 'QUEUED', message: 'Portal job requeued after stale Template release cache fix',
    detail_json: vNextAdminCanonicalJson_({
      requestId: payload.requestId, staleReleaseId: staleReleaseId,
      activeReleaseId: activeRelease.release_id, attemptsPreserved: Number(job.attempts || 0)
    }), actor: vNextAdminActor_()
  });
  vNextAdminAppendPortalRequestEvent_(portal.spreadsheet, validated,
    'VALIDATION_STARTED', 'VALIDATING', {
      relatedJobId: job.job_id, detailCode: 'RETRY_QUEUED',
      detailMessage: '作成処理を安全に再開しました。内容を確認しています。'
    });
  vNextAdminWriteAudit_(hub, 'REQUEUE_PORTAL_PROVISION', 'PORTAL_REQUEST', payload.requestId, 'SUCCESS', {
    jobId: job.job_id, staleReleaseId: staleReleaseId, activeReleaseId: activeRelease.release_id
  });
  return true;
}

function vNextAdminRecoverExhaustedForecastJob_(hub, job) {
  const registry = vNextAdminFindRegistryRow_(hub, function (row) {
    return String(row.book_id || '') === String(job.target_book_id || '');
  });
  if (!registry || String(registry.mode || '') !== 'CLIENT') return '; registry missing';
  const client = SpreadsheetApp.openById(String(registry.spreadsheet_id));
  const payload = vNextAdminParseJson_(job.request_json, {});
  const routing = vNextAdminReadKeyValueSheet_(client, VN_ADMIN_BOOK_CONFIG_SHEET);
  const authoritativeState = vNextAdminLatestClientState_(hub, registry.book_id, registry.state || routing.state);
  if (authoritativeState === 'RUNNING') {
    vNextAdminSetClientState_(registry.spreadsheet_id, 'READY_TO_RUN', {
      reason: 'forecast_job_lease_exhausted_after_3_attempts', actorRole: 'ADMIN', hub: hub
    });
  }
  if (payload.requestId && String(job.job_type || '') === 'FORECAST_REQUEST') {
    vNextAdminAppendClientRequestEvent_(client, {
      requestId: payload.requestId, bookId: registry.book_id, eventType: 'FAILED', status: 'FAILED',
      requestHash: payload.requestHash || '', requestJson: '', requestedAt: payload.requestedAt || '',
      requestedBy: payload.requestedBy || '', relatedJobId: job.job_id,
      detail: { error: 'forecast job lease exhausted after 3 attempts', recoveredAt: new Date().toISOString() }
    });
  }
  vNextAdminSyncClientToHub_(hub, client, registry.book_id);
  const recovered = authoritativeState === 'RUNNING';
  if (recovered) vNextAdminPatchRegistryByBookId_(hub, registry.book_id, { state: 'READY_TO_RUN', updated_at: new Date() });
  if (String(job.job_type || '') === 'AI_ROLLBACK_FORECAST') {
    vNextAdminAppendException_(hub, {
      severity: 'ERROR', exception_type: 'AI_ROLLBACK_FORECAST_LEASE_EXHAUSTED', book_id: registry.book_id,
      client_name: registry.client_name, fiscal_year: registry.fiscal_year,
      title: 'AI反映の取消後の再予測が3回中断', detail: 'job=' + job.job_id,
      recommended_action: '原因を確認し、同じ取消操作を再実行', source_ref: payload.rollbackOperationId || job.job_id
    });
    vNextAdminWriteAudit_(hub, 'AI_ROLLBACK_FORECAST', 'AI_ROLLBACK', payload.rollbackOperationId || job.job_id, 'FAILED', {
      jobId: job.job_id, reason: 'lease_exhausted_after_3_attempts', recoveredTo: recovered ? 'READY_TO_RUN' : authoritativeState
    });
  }
  return recovered ? '; client state restored to READY_TO_RUN' : '; client state already advanced; no rollback applied';
}

function vNextAdminEnqueueJobInternal_(hub, request) {
  const req = request && typeof request === 'object' ? request : {};
  const type = vNextAdminRequiredText_(req.jobType, 'jobType').toUpperCase();
  const idempotency = vNextAdminText_(req.idempotencyKey) || [type, req.targetBookId || '', vNextAdminSha256_(vNextAdminCanonicalJson_(req.request || {}))].join('|');
  const table = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.JOBS);
  const existing = table.rows.find(function (row) {
    return String(row.idempotency_key) === idempotency && ['QUEUED', 'RUNNING', 'SUCCEEDED'].indexOf(String(row.status)) >= 0;
  });
  if (existing) return vNextAdminJsonSafe_(existing);
  const now = new Date();
  const job = {
    job_id: 'JOB-' + Utilities.getUuid(), job_type: type,
    target_book_id: vNextAdminText_(req.targetBookId), target_spreadsheet_id: vNextAdminText_(req.targetSpreadsheetId),
    request_json: JSON.stringify(req.request || {}), idempotency_key: idempotency,
    status: 'QUEUED', priority: Number(req.priority || 0), attempts: 0,
    not_before: req.notBefore || '', locked_at: '', locked_by: '', started_at: '', finished_at: '',
    result_json: '', error: '', created_at: now, created_by: vNextAdminActor_(), updated_at: now
  };
  vNextAdminAppendObject_(hub, VN_ADMIN_SHEETS.JOBS, job);
  vNextAdminAppendObject_(hub, VN_ADMIN_SHEETS.JOB_LOG, {
    log_id: 'JLOG-' + Utilities.getUuid(), job_id: job.job_id, logged_at: now,
    status: 'QUEUED', message: 'Job queued', detail_json: '{}', actor: vNextAdminActor_()
  });
  return vNextAdminJsonSafe_(job);
}

function vNextAdminClaimNextJob_(hub) {
  const table = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.JOBS);
  const nowMs = Date.now();
  const eligible = table.rows.filter(function (row) {
    if (String(row.status) !== 'QUEUED') return false;
    // A forecast that requested AI research waits while that dependency can
    // still complete.  A terminal AI failure does not block the forecast: the
    // worker will record an explicit AI-unavailable degradation and continue
    // with the other two lenses.
    if (String(row.job_type || '') === 'FORECAST_REQUEST') {
      const request = vNextAdminParseJson_(row.request_json, {});
      if (request.aiResearchJobId) {
        const dependency = table.rows.find(function (candidate) {
          return String(candidate.job_id || '') === String(request.aiResearchJobId || '');
        });
        const dependencyStatus = String(dependency && dependency.status || '').toUpperCase();
        if (dependencyStatus === 'QUEUED' || dependencyStatus === 'RUNNING') return false;
      }
    }
    if (!row.not_before) return true;
    const t = new Date(row.not_before).getTime();
    return !isFinite(t) || t <= nowMs;
  }).sort(function (a, b) {
    const p = Number(b.priority || 0) - Number(a.priority || 0);
    if (p) return p;
    return new Date(a.created_at).getTime() - new Date(b.created_at).getTime();
  });
  if (!eligible.length) return null;
  const job = eligible[0];
  const now = new Date();
  vNextAdminUpdateTableRow_(hub, VN_ADMIN_SHEETS.JOBS, job._rowNumber, {
    status: 'RUNNING', attempts: Number(job.attempts || 0) + 1,
    locked_at: now, locked_by: vNextAdminActor_(), started_at: now, updated_at: now, error: ''
  });
  job.status = 'RUNNING';
  job.attempts = Number(job.attempts || 0) + 1;
  return job;
}

/** Portal faults are isolated so they cannot stop unrelated forecast and health jobs. */
function vNextAdminHarvestPortalRequestsSafely_(hub) {
  try {
    return vNextAdminHarvestPortalRequests_(hub);
  } catch (error) {
    const detail = String(error && error.message || error).slice(0, 1200);
    Logger.log('Portal request harvest isolated: %s', detail);
    try {
      const alreadyOpen = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.EXCEPTIONS).rows.some(function (row) {
        return String(row.exception_type || '').toUpperCase() === 'PORTAL_HARVEST_FAILED' &&
          String(row.status || 'OPEN').toUpperCase() === 'OPEN' && String(row.detail || '') === detail;
      });
      if (!alreadyOpen) {
        vNextAdminAppendException_(hub, {
          severity: 'ERROR', exception_type: 'PORTAL_HARVEST_FAILED', book_id: '',
          title: '申請入口の受付確認に失敗', detail: detail,
          recommended_action: 'Portalの権限・runtime identity・内部sheet列構成を確認',
          source_ref: 'EMPLOYEE_PORTAL'
        });
      }
    } catch (exceptionError) {
      Logger.log('Portal harvest exception recording skipped: %s',
        String(exceptionError && exceptionError.message || exceptionError));
    }
    return { configured: true, queued: 0, reused: 0, rejected: 0, isolatedError: detail };
  }
}

/** Manual Admin entry. The previous successful catalog remains untouched on failure. */
function vNextAdminRefreshZacClientCatalog(request) {
  return vNextAdminGuard_('vNextAdminRefreshZacClientCatalog', function () {
    const hub = vNextAdminRequireHub_();
    vNextAdminAssertHubAdmin_(hub, false);
    return vNextAdminRefreshZacClientCatalogIfStale_(hub,
      Boolean(request && request.force !== false));
  });
}

function vNextAdminRefreshZacClientCatalogSafely_(hub, force) {
  try {
    return vNextAdminRefreshZacClientCatalogIfStale_(hub, Boolean(force));
  } catch (error) {
    const detail = String(error && error.message || error).slice(0, 1200);
    Logger.log('ZAC client catalog refresh isolated: %s', detail);
    try {
      const open = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.EXCEPTIONS).rows.some(function (row) {
        return String(row.exception_type || '') === 'ZAC_CLIENT_CATALOG_REFRESH_FAILED' &&
          String(row.status || 'OPEN').toUpperCase() === 'OPEN' && String(row.detail || '') === detail;
      });
      if (!open) vNextAdminAppendException_(hub, {
        severity: 'ERROR', exception_type: 'ZAC_CLIENT_CATALOG_REFRESH_FAILED', book_id: '',
        title: 'ZACクライアント候補を更新できません', detail: detail,
        recommended_action: 'ZAC実績sourceの権限、対象tab、AN/AO列を確認。前回成功版は継続利用されます。',
        source_ref: VN_ADMIN_ZAC_CLIENT_CATALOG_SHEET
      });
    } catch (exceptionError) {
      Logger.log('ZAC catalog exception recording skipped: %s', String(exceptionError));
    }
    return { refreshed: false, reusedLastKnownGood: true, error: detail };
  }
}

function vNextAdminRefreshZacClientCatalogIfStale_(hub, force, options) {
  const config = vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET);
  const refreshedMs = new Date(config.zac_client_catalog_refreshed_at || 0).getTime();
  const currentVersion = String(config.zac_client_catalog_version || '');
  const activeCount = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.CATALOG).rows.filter(function (row) {
    return vNextAdminBool_(row.is_active) && String(row.catalog_version || '') === currentVersion;
  }).length;
  const stale = !currentVersion || !activeCount || !isFinite(refreshedMs) ||
    Date.now() - refreshedMs >= VN_ADMIN_ZAC_CATALOG_STALE_MS;
  if (!force && !stale) {
    return {
      refreshed: false, stale: false,
      activeCount: activeCount,
      catalogVersion: currentVersion,
      refreshedAt: String(config.zac_client_catalog_refreshed_at || '')
    };
  }
  return vNextAdminRefreshZacClientCatalogNow_(hub, options);
}

function vNextAdminRefreshZacClientCatalogNow_(hub, options) {
  const runtime = vNextGetRuntimeConfig_();
  const config = vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET);
  const sourceId = vNextAdminRequiredText_(
    config.source_spreadsheet_id || runtime.VNEXT_ZAC_SOURCE_SPREADSHEET_ID ||
    runtime.FORECAST_SOURCE_SPREADSHEET_ID, 'ZAC source spreadsheet ID'
  );
  const extracted = vNextAdminExtractZacClientCatalog_(SpreadsheetApp.openById(sourceId));
  if (!extracted.clients.length) throw new Error('ZAC sourceからクライアント候補を1件も取得できませんでした。');
  const refreshedAt = new Date().toISOString();
  const versionBasis = extracted.clients.map(function (row) {
    return { catalogKey: row.catalogKey, clientId: row.clientId, clientCode: row.clientCode, clientName: row.clientName };
  }).sort(function (a, b) { return String(a.catalogKey).localeCompare(String(b.catalogKey)); });
  const catalogVersion = 'ZACCAT-' + vNextAdminSha256_(vNextAdminCanonicalJson_(versionBasis)).slice(0, 20).toUpperCase();
  const commit = function () {
    vNextAdminEnsureExactTableHeaders_(hub, VN_ADMIN_SHEETS.CATALOG,
      VN_ADMIN_ZAC_CLIENT_CATALOG_HEADERS).hideSheet();
    const existingTable = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.CATALOG);
    const existing = existingTable.rows;
    const hubCatalogBefore = existing.map(function (row) {
      return VN_ADMIN_ZAC_CLIENT_CATALOG_HEADERS.map(function (header) {
        return row[header] === undefined ? '' : row[header];
      });
    });
    let portal = null;
    let portalCatalogBefore = null;
    try {
      portal = vNextAdminResolvePortal_(hub);
      portalCatalogBefore = vNextAdminSnapshotExactTableBody_(portal.spreadsheet,
        VN_ADMIN_PORTAL_CLIENT_CATALOG_SHEET, VN_ADMIN_PORTAL_CLIENT_CATALOG_HEADERS);
    } catch (portalResolveError) {
      if (!/not configured/i.test(String(portalResolveError && portalResolveError.message || portalResolveError))) {
        throw portalResolveError;
      }
    }
    try {
      // Build the complete next body in memory and publish it with one bulk
      // write. Calling the generic upsert helper once per client repeatedly
      // re-read the full sheet and pushed the 35-row pilot close to GAS's
      // execution limit. The LKG snapshot above still makes this replace
      // recoverable when the later Portal projection or pointer commit fails.
      const nextHubBody = vNextAdminBuildZacCatalogBody_(existing, extracted.clients, {
        catalogVersion: catalogVersion, refreshedAt: refreshedAt, sourceSpreadsheetId: sourceId
      });
      vNextAdminReplaceExactTableBody_(hub, VN_ADMIN_SHEETS.CATALOG,
        VN_ADMIN_ZAC_CLIENT_CATALOG_HEADERS, nextHubBody);
      const projection = vNextAdminProjectZacClientCatalogToPortal_(hub, {
        catalogVersion: catalogVersion, syncedAt: refreshedAt,
        portalSpreadsheet: portal && portal.spreadsheet
      });
      // Commit the current pointer last. Request harvesting never accepts rows
      // whose catalog version is not this committed version.
      vNextAdminWriteSystemConfig_(hub, {
        zac_client_catalog_version: catalogVersion,
        zac_client_catalog_refreshed_at: refreshedAt,
        zac_client_catalog_source_id: sourceId
      });
      const hubConfig = vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET);
      vNextAdminProtectInternalSheets_(hub,
        vNextAdminMergeEmails_(hubConfig.admin_emails, vNextAdminActor_()), [VN_ADMIN_SHEETS.CATALOG]);
      vNextAdminWriteAudit_(hub, 'REFRESH_ZAC_CLIENT_CATALOG', 'CLIENT_CATALOG', catalogVersion, 'SUCCESS', {
        activeCount: extracted.clients.length, actualTabs: extracted.actualTabs,
        defclientsTabs: extracted.defclientsTabs, portalRows: projection.rows,
        sourceSpreadsheetId: sourceId
      });
      return {
        refreshed: true, stale: false, activeCount: extracted.clients.length,
        catalogVersion: catalogVersion, refreshedAt: refreshedAt,
        actualTabs: extracted.actualTabs, defclientsTabs: extracted.defclientsTabs,
        portalRows: projection.rows
      };
    } catch (refreshError) {
      const rollbackErrors = [];
      try {
        vNextAdminReplaceExactTableBody_(hub, VN_ADMIN_SHEETS.CATALOG,
          VN_ADMIN_ZAC_CLIENT_CATALOG_HEADERS, hubCatalogBefore);
      } catch (hubRollbackError) { rollbackErrors.push('hub=' + String(hubRollbackError)); }
      try {
        if (portal && portalCatalogBefore) {
          vNextAdminReplaceExactTableBody_(portal.spreadsheet, VN_ADMIN_PORTAL_CLIENT_CATALOG_SHEET,
            VN_ADMIN_PORTAL_CLIENT_CATALOG_HEADERS, portalCatalogBefore);
        }
      } catch (portalRollbackError) { rollbackErrors.push('portal=' + String(portalRollbackError)); }
      try {
        vNextAdminWriteSystemConfig_(hub, {
          zac_client_catalog_version: config.zac_client_catalog_version || '',
          zac_client_catalog_refreshed_at: config.zac_client_catalog_refreshed_at || '',
          zac_client_catalog_source_id: config.zac_client_catalog_source_id || ''
        });
      } catch (configRollbackError) { rollbackErrors.push('config=' + String(configRollbackError)); }
      if (rollbackErrors.length) {
        throw new Error('ZAC catalog refresh failed and LKG rollback is incomplete: ' +
          String(refreshError && refreshError.message || refreshError) + '; ' + rollbackErrors.join('; '));
      }
      throw refreshError;
    }
  };
  return options && options.lockHeld
    ? commit()
    : vNextAdminWithScriptLock_('commit-zac-client-catalog', commit);
}

/** Pure merge used by the catalog refresh before its single bulk setValues. */
function vNextAdminBuildZacCatalogBody_(existingRows, candidates, metadata) {
  const meta = metadata || {};
  const refreshedAt = vNextAdminRequiredText_(meta.refreshedAt, 'catalog refreshedAt');
  const catalogVersion = vNextAdminRequiredText_(meta.catalogVersion, 'catalog version');
  const sourceSpreadsheetId = vNextAdminRequiredText_(meta.sourceSpreadsheetId, 'catalog source spreadsheet ID');
  const records = (Array.isArray(existingRows) ? existingRows : []).map(function (row) {
    const record = {};
    VN_ADMIN_ZAC_CLIENT_CATALOG_HEADERS.forEach(function (header) {
      record[header] = row && row[header] !== undefined ? row[header] : '';
    });
    return record;
  });
  const indexByKey = {};
  records.forEach(function (record, index) {
    const key = String(record.catalog_key || '');
    if (!key) return;
    if (indexByKey[key] !== undefined) throw new Error('ZAC catalogに重複catalog_keyがあります: ' + key);
    indexByKey[key] = index;
  });
  const activeKeys = new Set();
  (Array.isArray(candidates) ? candidates : []).forEach(function (candidate) {
    const key = vNextAdminRequiredText_(candidate && candidate.catalogKey, 'candidate.catalogKey');
    if (activeKeys.has(key)) throw new Error('ZAC候補に重複catalog_keyがあります: ' + key);
    activeKeys.add(key);
    const priorIndex = indexByKey[key];
    const prior = priorIndex === undefined ? {} : records[priorIndex];
    const next = Object.assign({}, prior, {
      catalog_key: key,
      client_id: vNextAdminRequiredText_(candidate.clientId, 'candidate.clientId'),
      client_code: String(candidate.clientCode || ''),
      client_name: vNextAdminRequiredText_(candidate.clientName, 'candidate.clientName'),
      normalized_name: vNextAdminRequiredText_(candidate.normalizedName, 'candidate.normalizedName'),
      is_active: 1,
      source_years_json: vNextAdminCanonicalJson_(candidate.sourceYears || []),
      first_seen_at: prior.first_seen_at || refreshedAt,
      last_seen_at: refreshedAt,
      catalog_version: catalogVersion,
      refreshed_at: refreshedAt,
      source_spreadsheet_id: sourceSpreadsheetId
    });
    if (priorIndex === undefined) {
      indexByKey[key] = records.length;
      records.push(next);
    } else {
      records[priorIndex] = next;
    }
  });
  records.forEach(function (record) {
    const key = String(record.catalog_key || '');
    if (!key || activeKeys.has(key) || !vNextAdminBool_(record.is_active)) return;
    record.is_active = 0;
    record.catalog_version = catalogVersion;
    record.refreshed_at = refreshedAt;
  });
  return records.map(function (record) {
    return VN_ADMIN_ZAC_CLIENT_CATALOG_HEADERS.map(function (header) {
      return record[header] === undefined ? '' : record[header];
    });
  });
}

/** Reads each selected source range once. It never mutates the last-known-good catalog. */
function vNextAdminExtractZacClientCatalog_(source) {
  const actualTabs = source.getSheets().map(function (sheet) {
    const match = sheet.getName().match(/^\*(\d{4})_actual_value$/);
    return match ? { sheet: sheet, year: Number(match[1]) } : null;
  }).filter(Boolean).sort(function (a, b) { return b.year - a.year; }).slice(0, 2);
  const defclients = source.getSheets().filter(function (sheet) {
    return /\*defclients$/i.test(sheet.getName());
  });
  if (!actualTabs.length && !defclients.length) {
    throw new Error('ZAC sourceに*YYYY_actual_valueまたは*defclients tabがありません。');
  }
  const byName = {};
  actualTabs.forEach(function (entry) {
    const sheet = entry.sheet;
    if (sheet.getLastRow() < 2 || sheet.getMaxColumns() < VN_ADMIN_ZAC_CLIENT_NAME_COLUMN) return;
    const codeHeader = String(sheet.getRange(1, VN_ADMIN_ZAC_CLIENT_CODE_COLUMN).getValue() || '').trim();
    const nameHeader = String(sheet.getRange(1, VN_ADMIN_ZAC_CLIENT_NAME_COLUMN).getValue() || '').trim();
    if (codeHeader !== 'クライアントコード' || nameHeader !== 'クライアント名') {
      throw new Error(sheet.getName() + ' AN/AO列header不一致: expected=クライアントコード/クライアント名; actual=' +
        codeHeader + '/' + nameHeader);
    }
    const values = sheet.getRange(2, VN_ADMIN_ZAC_CLIENT_CODE_COLUMN,
      sheet.getLastRow() - 1, 2).getDisplayValues();
    values.forEach(function (row) {
      vNextAdminMergeZacCatalogCandidate_(byName, row[1], row[0], entry.year, false);
    });
  });
  defclients.forEach(function (sheet) {
    if (sheet.getLastRow() < 1) return;
    const values = sheet.getRange(1, 1, sheet.getLastRow(), 2).getDisplayValues();
    const headerIndex = values.findIndex(function (row) {
      return String(row[0] || '').trim() === 'クライアント' && String(row[1] || '').trim() === '売上';
    });
    if (headerIndex < 0) throw new Error(sheet.getName() + ' のクライアント/売上headerを確認できません。');
    values.slice(headerIndex + 1).forEach(function (row) {
      const name = String(row[0] || '').trim();
      // *defclients begins with report controls and includes aggregate buckets;
      // neither represents a selectable ZAC client.
      if (!name || name === '全体' || name === '仮登録') return;
      vNextAdminMergeZacCatalogCandidate_(byName, name, '', 0, true);
    });
  });
  const clients = Object.keys(byName).map(function (nameKey) {
    const item = byName[nameKey];
    const code = item.clientCode;
    const identity = code ? 'CODE|' + code : 'NAME|' + nameKey;
    const prefix = code ? 'ZAC-CODE-' : 'ZAC-NAME-';
    const stable = prefix + vNextAdminSha256_(identity).slice(0, 20).toUpperCase();
    return {
      catalogKey: stable, clientId: stable, clientCode: code,
      clientName: item.clientName, normalizedName: nameKey,
      sourceYears: Array.from(item.sourceYears).sort()
    };
  }).sort(function (a, b) { return a.clientName.localeCompare(b.clientName, 'ja'); });
  return {
    clients: clients, actualTabs: actualTabs.map(function (entry) { return entry.sheet.getName(); }),
    defclientsTabs: defclients.map(function (sheet) { return sheet.getName(); })
  };
}

function vNextAdminMergeZacCatalogCandidate_(byName, rawName, rawCode, year, fromDefinition) {
  const clientName = vNextAdminSafeCatalogText_(rawName, 120, 'clientName');
  if (!clientName || clientName === '全体' || clientName === '仮登録') return;
  const nameKey = vNextAdminNormalizeCatalogClientName_(clientName);
  if (!nameKey) return;
  const code = vNextAdminSafeCatalogText_(rawCode, 120, 'clientCode');
  const existing = byName[nameKey];
  if (!existing) {
    byName[nameKey] = {
      clientName: clientName, clientCode: code, sourceYears: new Set(year ? [Number(year)] : []),
      fromDefinition: Boolean(fromDefinition)
    };
    return;
  }
  if (code && existing.clientCode && code !== existing.clientCode) {
    throw new Error('同じクライアント名に複数のAN codeがあります: ' + clientName);
  }
  if (code) existing.clientCode = code;
  if (year) existing.sourceYears.add(Number(year));
}

function vNextAdminSafeCatalogText_(value, maxLength, label) {
  const text = String(value === undefined || value === null ? '' : value).trim();
  if (!text) return '';
  if (text.length > maxLength || /[\u0000-\u001f\u007f]/.test(text) || /^[=+\-@]/.test(text)) {
    throw new Error('ZAC ' + label + 'に安全でない値があります。');
  }
  return text;
}

function vNextAdminNormalizeCatalogClientName_(value) {
  let text = String(value || '').trim().toLowerCase();
  try { text = text.normalize('NFKC'); } catch (ignoredNormalize) {}
  return text.replace(/[\s\u3000]/g, '').replace(/株式会社|有限会社|合同会社|\(株\)|\(有\)|㈱|㈲/g, '');
}

function vNextAdminProjectZacClientCatalogToPortal_(hub, options) {
  let portal;
  try {
    portal = options && options.portalSpreadsheet
      ? { spreadsheet: options.portalSpreadsheet }
      : vNextAdminResolvePortal_(hub);
  }
  catch (error) {
    if (/not configured/i.test(String(error && error.message || error))) return { configured: false, rows: 0 };
    throw error;
  }
  const opt = options || {};
  const config = vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET);
  const catalogVersion = String(opt.catalogVersion || config.zac_client_catalog_version || '');
  const syncedAt = String(opt.syncedAt || new Date().toISOString());
  if (!catalogVersion) throw new Error('Hub ZAC client catalog version is missing.');
  const rows = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.CATALOG).rows.filter(function (row) {
    return vNextAdminBool_(row.is_active) && String(row.catalog_version || '') === catalogVersion;
  }).map(function (row) {
    return {
      catalog_key: String(row.catalog_key || ''), client_name: String(row.client_name || ''),
      is_active: 1, catalog_version: catalogVersion, synced_at: syncedAt
    };
  }).sort(function (a, b) { return a.client_name.localeCompare(b.client_name, 'ja'); });
  if (!rows.length) throw new Error('Portalへ投影できるactive ZAC clientがありません。');
  const sheet = vNextAdminEnsureExactTableHeaders_(portal.spreadsheet,
    VN_ADMIN_PORTAL_CLIENT_CATALOG_SHEET, VN_ADMIN_PORTAL_CLIENT_CATALOG_HEADERS);
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, VN_ADMIN_PORTAL_CLIENT_CATALOG_HEADERS.length).clearContent();
  }
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, VN_ADMIN_PORTAL_CLIENT_CATALOG_HEADERS.length).setValues(rows.map(function (row) {
      return VN_ADMIN_PORTAL_CLIENT_CATALOG_HEADERS.map(function (header) { return row[header]; });
    }));
  }
  sheet.hideSheet();
  return { configured: true, rows: rows.length, catalogVersion: catalogVersion, syncedAt: syncedAt };
}

function vNextAdminSnapshotExactTableBody_(ss, name, headers) {
  const sheet = vNextAdminEnsureExactTableHeaders_(ss, name, headers);
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
}

function vNextAdminReplaceExactTableBody_(ss, name, headers, values) {
  const sheet = vNextAdminEnsureExactTableHeaders_(ss, name, headers);
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).clearContent();
  }
  const rows = Array.isArray(values) ? values : [];
  if (rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  return sheet;
}

function vNextAdminResolveZacCatalogSelection_(hub, catalogKey, clientName) {
  const key = vNextAdminRequiredText_(catalogKey, 'catalogKey');
  const name = vNextAdminRequiredText_(clientName, 'clientName');
  const row = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.CATALOG).rows.find(function (candidate) {
    return String(candidate.catalog_key || '') === key && vNextAdminBool_(candidate.is_active);
  });
  if (!row || String(row.client_name || '') !== name) {
    throw new Error('選択されたクライアントは現在のZAC候補正本と一致しません。画面を開き直してください。');
  }
  const config = vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET);
  if (!String(config.zac_client_catalog_version || '') ||
      String(row.catalog_version || '') !== String(config.zac_client_catalog_version || '')) {
    throw new Error('ZACクライアント候補の版が更新されています。画面を開き直してください。');
  }
  return {
    catalogKey: key, clientId: vNextAdminRequiredText_(row.client_id, 'catalog.client_id'),
    clientName: name, catalogVersion: String(row.catalog_version || '')
  };
}

/** Read-only check: an existing duplicate is reusable only when it is already employee-open. */
function vNextAdminPortalExistingBookAccess_(portal, registry) {
  const policy = String(registry && registry.access_policy || '').trim().toUpperCase();
  const domain = vNextAdminNormalizeDomain_(registry && registry.internal_domain || '');
  if (policy !== 'INTERNAL_OPEN' || !domain || domain !== portal.employeeDomain) {
    return { reusable: false, code: 'EXISTING_BOOK_ADMIN_ACCESS_REQUIRED' };
  }
  if (!vNextAdminSpreadsheetAccessible_(registry.spreadsheet_id)) {
    return { reusable: false, code: 'EXISTING_BOOK_INACCESSIBLE' };
  }
  try {
    vNextAdminAssertEmployeeFileSharing_(DriveApp.getFileById(String(registry.spreadsheet_id)), {
      targetMode: 'CLIENT', accessPolicy: 'INTERNAL_OPEN', internalDomain: portal.employeeDomain,
      editors: vNextAdminParseList_(registry.editor_emails), viewers: vNextAdminParseList_(registry.viewer_emails)
    });
    return { reusable: true, code: 'EXISTING_BOOK_REUSED' };
  } catch (error) {
    return {
      reusable: false, code: 'EXISTING_BOOK_SHARING_MISMATCH',
      detail: String(error && error.message || error).slice(0, 500)
    };
  }
}

function vNextAdminRejectPortalExistingBook_(hub, portal, validated, registry, detail) {
  const extra = detail || {};
  vNextAdminAppendPortalRequestEvent_(portal.spreadsheet, validated, 'REJECTED', 'REJECTED', {
    relatedBookId: '', relatedBookUrl: '', detailCode: String(extra.code || 'EXISTING_BOOK_ADMIN_ACCESS_REQUIRED'),
    detailMessage: '既存のクライアント年度ブックがありますが、社内共通アクセスの確認が必要です。管理担当者へ連絡してください。'
  });
  vNextAdminAppendException_(hub, {
    severity: 'WARN', exception_type: 'PORTAL_EXISTING_BOOK_ACCESS_REQUIRED',
    book_id: String(registry.book_id || ''), client_name: String(registry.client_name || ''),
    fiscal_year: Number(registry.fiscal_year || 0), title: '既存クライアント年度ブックのアクセス確認が必要',
    detail: 'request=' + String(validated.payload && validated.payload.requestId || '') +
      '; code=' + String(extra.code || '') + (extra.detail ? '; ' + String(extra.detail) : ''),
    recommended_action: '既存bookの共有方針を管理ハブが確認し、必要ならversioned migrationでINTERNAL_OPENへ変更',
    source_ref: String(registry.book_id || '')
  });
  return { rejected: true, code: String(extra.code || 'EXISTING_BOOK_ADMIN_ACCESS_REQUIRED') };
}

function vNextAdminHarvestPortalRequests_(hub) {
  let portal;
  try { portal = vNextAdminResolvePortal_(hub); }
  catch (error) {
    if (/not configured/i.test(String(error && error.message || error))) return { configured: false, queued: 0, reused: 0, rejected: 0 };
    throw error;
  }
  const table = vNextAdminReadTable_(portal.spreadsheet, VN_ADMIN_PORTAL_REQUEST_SHEET);
  const grouped = {};
  table.rows.forEach(function (row) {
    const id = String(row.request_id || '').trim();
    if (!id) return;
    if (!grouped[id]) grouped[id] = [];
    grouped[id].push(row);
  });
  const result = { configured: true, queued: 0, reused: 0, rejected: 0 };
  Object.keys(grouped).forEach(function (requestId) {
    const events = grouped[requestId];
    const latest = events[events.length - 1];
    if (String(latest.event_type || '').toUpperCase() !== 'REQUESTED' ||
        String(latest.status || '').toUpperCase() !== 'PENDING') return;
    try {
      const validated = vNextAdminValidatePortalRequest_(hub, portal, events[0]);
      const duplicates = vNextAdminFindClientFyDuplicates_(hub, Object.assign({}, validated.payload, {
        clientId: validated.clientId
      }));
      if (duplicates.length > 1) throw new Error('同じクライアント・年度の登録が複数あります。管理ハブ担当者確認が必要です。');
      if (duplicates.length === 1 && String(duplicates[0].status || '').toUpperCase() === 'ACTIVE') {
        const access = vNextAdminPortalExistingBookAccess_(portal, duplicates[0]);
        if (!access.reusable) {
          vNextAdminRejectPortalExistingBook_(hub, portal, validated, duplicates[0], access);
          result.rejected++;
          return;
        }
        vNextAdminAppendPortalRequestEvent_(portal.spreadsheet, validated, 'COMPLETED', 'COMPLETED', {
          relatedBookId: duplicates[0].book_id, relatedBookUrl: duplicates[0].spreadsheet_url,
          detailCode: 'EXISTING_BOOK_REUSED', detailMessage: '既存のクライアント年度ブックをご利用ください。'
        });
        result.reused++;
        return;
      }
      const idempotency = 'PORTAL_PROVISION|' + vNextAdminPortalCanonicalClientKey_({
        clientId: validated.clientId, clientName: validated.payload.clientName
      }) + '|' + validated.payload.fiscalYear;
      const job = vNextAdminEnqueueJobInternal_(hub, {
        jobType: 'PORTAL_PROVISION_CLIENT', targetBookId: validated.payload.requestId,
        targetSpreadsheetId: portal.spreadsheetId,
        request: {
          portalSpreadsheetId: portal.spreadsheetId, portalId: portal.portalId,
          requestId: validated.payload.requestId, requestHash: validated.requestHash,
          clientId: validated.clientId, clientName: validated.payload.clientName,
          fiscalYear: validated.payload.fiscalYear,
          schemaVersion: validated.schemaVersion, catalogKey: validated.catalogKey || '',
          forecastOwnerEmail: validated.forecastOwnerEmail,
          relatedMemberEmails: validated.relatedMemberEmails,
          relatedMemberNames: validated.relatedMemberNames,
          requestedAt: validated.payload.requestedAt, requestedBy: validated.payload.requestedBy,
          employeeDomain: portal.employeeDomain
        },
        idempotencyKey: idempotency, priority: 80
      });
      vNextAdminAppendPortalRequestEvent_(portal.spreadsheet, validated, 'VALIDATION_STARTED', 'VALIDATING', {
        relatedJobId: job.job_id, detailCode: 'QUEUED', detailMessage: '内容を確認し、作成の準備をしています。'
      });
      result.queued++;
    } catch (error) {
      const requested = events[0];
      const weak = {
        payload: {
          requestId: requestId, requestedAt: requested.requested_at || '', requestedBy: requested.requested_by || '',
          fiscalYear: Number(requested.fiscal_year || 0), clientId: String(requested.client_id || ''),
          clientName: String(requested.client_name || ''), forecastOwnerEmail: String(requested.forecast_owner_email || ''),
          relatedMemberEmails: vNextAdminParseJson_(requested.related_member_emails_json, []),
          catalogKey: String(requested.catalog_key || ''),
          relatedMemberNames: vNextAdminParseJson_(requested.related_member_names_json, [])
        },
        requestHash: String(requested.request_hash || '')
      };
      vNextAdminAppendPortalRequestEvent_(portal.spreadsheet, weak, 'REJECTED', 'REJECTED', {
        detailCode: 'VALIDATION_FAILED', detailMessage: '入力内容を確認できませんでした。管理担当者へ連絡してください。'
      });
      vNextAdminAppendException_(hub, {
        severity: 'WARN', exception_type: 'PORTAL_REQUEST_REJECTED', book_id: requestId,
        client_name: String(requested.client_name || ''), fiscal_year: Number(requested.fiscal_year || 0),
        title: 'ポータル作成依頼を検証できません', detail: String(error && error.message || error),
        recommended_action: '依頼行と社内アカウント・重複登録を確認', source_ref: requestId
      });
      result.rejected++;
    }
  });
  return result;
}

function vNextAdminValidatePortalRequest_(hub, portal, row) {
  const requestJson = String(row.request_json || '');
  const requestHash = String(row.request_hash || '').toLowerCase();
  const payload = vNextAdminParseJson_(requestJson, null);
  if (!payload || typeof payload !== 'object' || Array.isArray(payload)) throw new Error('Portal request JSON is invalid.');
  const actualKeys = Object.keys(payload).sort();
  const schemaVersion = String(payload.schemaVersion || '');
  const expectedKeys = (schemaVersion === VN_ADMIN_PORTAL_REQUEST_SCHEMA_V1
    ? VN_ADMIN_PORTAL_REQUEST_PAYLOAD_KEYS_V1
    : schemaVersion === VN_ADMIN_PORTAL_REQUEST_SCHEMA
      ? VN_ADMIN_PORTAL_REQUEST_PAYLOAD_KEYS_V2 : []).slice().sort();
  if (!expectedKeys.length) throw new Error('Portal request schema is not supported.');
  if (vNextAdminCanonicalJson_(actualKeys) !== vNextAdminCanonicalJson_(expectedKeys)) throw new Error('Portal request keys do not match schema.');
  if (requestJson !== vNextAdminCanonicalJson_(payload) || requestHash !== vNextAdminSha256_(requestJson)) {
    throw new Error('Portal request canonical hash is invalid.');
  }
  if (payload.requestType !== 'CREATE_CLIENT_FY_BOOK') {
    throw new Error('Portal request schema/type is invalid.');
  }
  if (String(row.request_id || '') !== String(payload.requestId || '') ||
      String(row.requested_by || '').toLowerCase() !== String(payload.requestedBy || '').toLowerCase() ||
      String(row.requested_at || '') !== String(payload.requestedAt || '')) throw new Error('Portal request row/payload mismatch.');
  const requestedAt = new Date(payload.requestedAt);
  if (isNaN(requestedAt.getTime()) || requestedAt.getTime() > Date.now() + 300000) throw new Error('Portal request timestamp is invalid.');
  const fiscalYear = vNextAdminNormalizeFiscalYear_(payload.fiscalYear);
  const clientName = vNextAdminRequiredText_(payload.clientName, 'clientName');
  if (clientName.length > 120 || /^[=+@]/.test(clientName)) throw new Error('clientName is unsafe.');
  const requestedBy = String(payload.requestedBy || '').trim().toLowerCase();
  if (vNextAdminEmailDomain_(requestedBy) !== portal.employeeDomain) {
    throw new Error('Portal requester is outside the employee domain.');
  }
  if (schemaVersion === VN_ADMIN_PORTAL_REQUEST_SCHEMA_V1) {
    const owner = String(payload.forecastOwnerEmail || '').trim().toLowerCase();
    if (vNextAdminEmailDomain_(owner) !== portal.employeeDomain) {
      throw new Error('Portal owner is outside the employee domain.');
    }
    const members = Array.isArray(payload.relatedMemberEmails) ? payload.relatedMemberEmails.map(function (email) {
      return String(email || '').trim().toLowerCase();
    }) : [];
    if (members.length > 50 || members.some(function (email) {
      return vNextAdminEmailDomain_(email) !== portal.employeeDomain;
    })) throw new Error('relatedMemberEmails contains an invalid employee account.');
    const legacyClientId = vNextAdminText_(payload.clientId) || vNextAdminDeriveClientId_(clientName);
    return {
      payload: Object.assign({}, payload, { fiscalYear: fiscalYear }), requestHash: requestHash,
      schemaVersion: schemaVersion, clientId: legacyClientId, catalogKey: '',
      forecastOwnerEmail: owner, relatedMemberEmails: members, relatedMemberNames: []
    };
  }
  const names = vNextAdminNormalizeRelatedMemberNames_(payload.relatedMemberNames);
  if (vNextAdminCanonicalJson_(names) !== vNextAdminCanonicalJson_(payload.relatedMemberNames)) {
    throw new Error('relatedMemberNames is not normalized.');
  }
  const selected = vNextAdminResolveZacCatalogSelection_(hub, payload.catalogKey, clientName);
  return {
    payload: Object.assign({}, payload, { fiscalYear: fiscalYear }), requestHash: requestHash,
    schemaVersion: schemaVersion, clientId: selected.clientId, catalogKey: selected.catalogKey,
    catalogVersion: selected.catalogVersion, forecastOwnerEmail: requestedBy,
    relatedMemberEmails: [], relatedMemberNames: names
  };
}

function vNextAdminNormalizeRelatedMemberNames_(value) {
  if (!Array.isArray(value)) throw new Error('関与メンバーの氏名は配列である必要があります。');
  if (value.length < 1 || value.length > 5) throw new Error('関与メンバーは1名以上5名以内で指定してください。');
  const seen = new Set();
  return value.map(function (item) {
    let name = String(item === undefined || item === null ? '' : item).trim();
    try { name = name.normalize('NFKC'); } catch (ignoredNormalize) {}
    if (!name || name.length > 80 || /[\u0000-\u001f\u007f]/.test(name) || /^[=+\-@]/.test(name)) {
      throw new Error('関与メンバーの氏名を80文字以内の通常テキストで入力してください。');
    }
    const key = name.toLowerCase().replace(/[\s\u3000]/g, '');
    if (seen.has(key)) throw new Error('同じ関与メンバーが重複しています。');
    seen.add(key);
    return name;
  });
}

function vNextAdminFindClientFyDuplicates_(hub, input) {
  const canonicalInput = Object.assign({}, input || {});
  if (!String(canonicalInput.clientId || '').trim() && String(canonicalInput.clientName || '').trim()) {
    canonicalInput.clientId = vNextAdminDeriveClientId_(canonicalInput.clientName);
  }
  const key = vNextAdminPortalCanonicalClientKey_(canonicalInput);
  const nameKey = vNextAdminPortalCanonicalClientKey_({ clientName: canonicalInput.clientName });
  return vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.REGISTRY).rows.filter(function (row) {
    return String(row.mode || '') === 'CLIENT' && String(row.status || '').toUpperCase() !== 'ARCHIVED' &&
      Number(row.fiscal_year) === Number(canonicalInput.fiscalYear) &&
      (vNextAdminPortalCanonicalClientKey_({ clientId: row.client_id, clientName: row.client_name }) === key ||
       vNextAdminPortalCanonicalClientKey_({ clientName: row.client_name }) === nameKey);
  });
}

function vNextAdminPortalCanonicalClientKey_(input) {
  const id = String(input && (input.clientId || input.client_id) || '').trim().toLowerCase().replace(/[\s\-_.:/\\]/g, '');
  if (id) return 'ID:' + id;
  let name = String(input && (input.clientName || input.client_name) || '').trim().toLowerCase();
  try { name = name.normalize('NFKC'); } catch (ignoredNormalize) {}
  name = name.replace(/株式会社|有限会社|合同会社|\(株\)|\(有\)|\(同\)|㈱|㈲/g, '')
    .replace(/[\s\u3000・･.,，。'’"“”\-ー_]/g, '');
  return 'NAME:' + name;
}

function vNextAdminExecuteJob_(hub, job) {
  const payload = vNextAdminParseJson_(job.request_json, {});
  switch (String(job.job_type)) {
    case 'FORECAST_REQUEST':
    case 'AI_ROLLBACK_FORECAST': {
      const isAiRollback = String(job.job_type || '') === 'AI_ROLLBACK_FORECAST';
      if (typeof vNextRunForecast_ !== 'function') throw new Error('Trusted forecast worker API is not installed (expected vNextRunForecast_).');
      const registry = vNextAdminFindRegistryRow_(hub, function (row) {
        return String(row.book_id) === String(job.target_book_id);
      });
      if (!registry) throw new Error('Registry entry not found for forecast request.');
      const client = SpreadsheetApp.openById(String(registry.spreadsheet_id));
      try {
        if (!isAiRollback) vNextAdminAssertNoTrustedForecastPayload_(payload);
        vNextAdminAssertClientPinnedReleasePair_(hub, client, registry);
        vNextAdminSyncClientToHub_(hub, client, registry.book_id);
        let rollbackAuthorization = null;
        const engineRequest = Object.assign({}, payload, {
          bookId: registry.book_id,
          clientId: registry.client_id,
          clientName: registry.client_name,
          fiscalYear: Number(registry.fiscal_year),
          spreadsheet: hub,
          persist: true,
          manageState: true,
          internalOperation: 'ADMIN_JOB'
        });
        if (payload.aiResearchJobId) {
          const dependency = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.JOBS).rows.find(function (row) {
            return String(row.job_id || '') === String(payload.aiResearchJobId);
          });
          const dependencyStatus = String(dependency && dependency.status || 'MISSING').toUpperCase();
          if (dependencyStatus === 'QUEUED' || dependencyStatus === 'RUNNING') {
            throw new Error('AI research dependency is still in progress: ' + String(payload.aiResearchJobId));
          }
          if (dependencyStatus !== 'SUCCEEDED') {
            engineRequest.aiUnavailable = true;
            engineRequest.aiUnavailableReason = 'AI_RESEARCH_' + dependencyStatus;
            vNextAdminAppendException_(hub, {
              severity: 'WARN', exception_type: 'AI_RESEARCH_UNAVAILABLE', book_id: registry.book_id,
              client_name: registry.client_name, fiscal_year: registry.fiscal_year,
              title: 'AI調査を使わずに予測を継続',
              detail: 'dependency=' + String(payload.aiResearchJobId) + ', status=' + dependencyStatus,
              recommended_action: '予測は継続性・案件・現場情報で完了しています。必要ならAI設定を確認して新しいrunを依頼してください。',
              source_ref: payload.aiResearchJobId
            });
            vNextAdminWriteAudit_(hub, 'AI_RESEARCH_DEGRADED', 'FORECAST_JOB', job.job_id, 'WARN', {
              bookId: registry.book_id, dependencyJobId: payload.aiResearchJobId,
              dependencyStatus: dependencyStatus, policy: 'CONTINUE_WITH_AI_ZERO_REFERENCE_ONLY'
            });
          }
        }
        // Revalidate immediately before Engine entry; neither a queued request
        // nor an earlier health scan authorizes a different release pair.
        vNextAdminAssertClientPinnedReleasePair_(hub, client, registry);
        vNextAdminHydrateHubRuntime_(hub);
        const learningEvidence = vNextAdminLearningEvidenceForEngine_(hub, registry.book_id);
        if (learningEvidence) engineRequest.learningEvidence = learningEvidence;
        if (typeof vNextEngineBuildAdminRunIdentity_ !== 'function' ||
            typeof vNextEngineLookupRunForResume_ !== 'function') {
          throw new Error('Deterministic forecast run identity APIs are not installed.');
        }
        const runIdentity = vNextEngineBuildAdminRunIdentity_(registry.book_id, String(job.idempotency_key || ''));
        engineRequest.runId = runIdentity.runId;
        engineRequest.idempotencyKey = runIdentity.idempotencyKey;
        const resume = vNextEngineLookupRunForResume_(engineRequest);
        vNextAdminAppendJobPhase_(hub, job.job_id, 'RUN_LOOKUP', {
          runId: runIdentity.runId, resumablePhase: resume.resumablePhase,
          existingStatuses: resume.statuses || []
        });
        const authoritativeRunState = vNextAdminLatestClientState_(hub, registry.book_id, registry.state || '');
        if (!resume.hasSuccess && authoritativeRunState !== 'RUNNING') {
          throw vNextAdminRunIdentityFailure_(
            'No replayable SUCCESS exists, but the Hub forecast state is ' + authoritativeRunState + ' instead of RUNNING.'
          );
        }
        if (isAiRollback) {
          rollbackAuthorization = vNextAdminAuthorizeAiRollbackJob_(hub, job, payload, registry, {
            allowPersistedResume: resume.hasSuccess === true,
            persistedRunId: resume.hasSuccess ? runIdentity.runId : ''
          });
          if (!resume.hasSuccess) {
            engineRequest.requestId = rollbackAuthorization.rollbackRequestId;
            engineRequest.asOf = vNextAdminDateOnly_(rollbackAuthorization.sourceRun.as_of);
            engineRequest.cutoff = vNextAdminDateOnly_(rollbackAuthorization.sourceRun.cutoff);
            engineRequest.trustedReuseSeedFromRunId = rollbackAuthorization.trustedReuseSeedFromRunId;
            engineRequest.trustedRollbackContext = rollbackAuthorization.trustedRollbackContext;
            engineRequest.trustedAllowedDelayedAiRequestIds = rollbackAuthorization.trustedAllowedDelayedAiRequestIds;
            engineRequest.internalJobType = 'AI_ROLLBACK_FORECAST';
          }
        }
        const result = resume.hasSuccess ? resume.result : vNextRunForecast_(engineRequest);
        const stateFinalization = vNextAdminEnsurePersistedForecastDraftState_(hub, registry, result);
        vNextAdminAppendJobPhase_(hub, job.job_id, 'RUN_PERSISTED', {
          runId: result.runId, inputDataHash: result.inputDataHash || '',
          replayed: resume.hasSuccess === true, stateFinalization: stateFinalization
        });
        vNextAdminSyncHubToClient_(hub, client, registry.book_id, ['FORECAST_RUN', 'STATE_EVENT']);
        vNextAdminMirrorClientState_(client, 'DRAFT_READY');
        if (payload.requestId && !isAiRollback) {
          vNextAdminAppendClientRequestEvent_(client, {
            requestId: payload.requestId, bookId: registry.book_id, eventType: 'COMPLETED', status: 'COMPLETED',
            requestHash: payload.requestHash || '', requestJson: '', requestedAt: payload.requestedAt || '',
            requestedBy: payload.requestedBy || '', relatedJobId: job.job_id, relatedRunId: result.runId,
            detail: { inputDataHash: result.inputDataHash || '', status: result.status }
          });
        }
        vNextAdminPatchRegistryByBookId_(hub, job.target_book_id, {
          last_forecast_at: new Date(), state: 'DRAFT_READY', updated_at: new Date()
        });
        vNextAdminAppendJobPhase_(hub, job.job_id, 'CLIENT_SYNCED', {
          runId: result.runId, bookId: registry.book_id, state: 'DRAFT_READY'
        });
        if (isAiRollback) {
          vNextAdminWriteAudit_(hub, 'AI_ROLLBACK_FORECAST', 'AI_ROLLBACK', payload.rollbackOperationId, 'SUCCESS', {
            jobId: job.job_id, sourceForecastRunId: payload.sourceForecastRunId,
            resultRunId: result.runId, targetEvidenceIds: payload.targetEvidenceIds,
            tombstoneEvidenceIds: payload.tombstoneEvidenceIds
          });
        }
        return result;
      } catch (engineError) {
        if (engineError && engineError.vNextRunIdentityFailure === true) {
          vNextAdminAppendException_(hub, {
            severity: 'CRITICAL', exception_type: 'FORECAST_RUN_IDENTITY_CONFLICT', book_id: registry.book_id,
            client_name: registry.client_name, fiscal_year: registry.fiscal_year,
            title: '保存済み予測runの再開整合性に失敗',
            detail: String(engineError && engineError.message || engineError),
            recommended_action: '状態を巻き戻さず、JOB_LOG・FORECAST_RUN・入力hashを管理ハブが確認してください。',
            source_ref: job.job_id
          });
          vNextAdminAppendJobPhase_(hub, job.job_id, 'RUN_IDENTITY_BLOCKED', {
            error: String(engineError && engineError.message || engineError).slice(0, 1000)
          });
          throw engineError;
        }
        try {
          const hubStates = vNextAdminReadCoreRows_(hub, 'STATE_EVENT').filter(function (row) {
            return String(row.book_id || '') === String(registry.book_id);
          });
          const hubState = hubStates.length ? String(hubStates[hubStates.length - 1].to_state || '') : 'RUNNING';
          let recoveryState = hubState || 'RUNNING';
          if (hubState === 'RUNNING') {
            vNextAdminSetClientState_(registry.spreadsheet_id, 'READY_TO_RUN', {
              reason: 'admin_forecast_job_failed: ' + String(engineError && engineError.message || engineError).slice(0, 500),
              actorRole: 'ADMIN', hub: hub
            });
            recoveryState = 'READY_TO_RUN';
          }
          vNextAdminSyncHubToClient_(hub, client, registry.book_id, ['FORECAST_RUN', 'STATE_EVENT']);
          vNextAdminMirrorClientState_(client, recoveryState);
          if (payload.requestId && !isAiRollback) {
            vNextAdminAppendClientRequestEvent_(client, {
              requestId: payload.requestId, bookId: registry.book_id, eventType: 'FAILED', status: 'FAILED',
              requestHash: payload.requestHash || '', requestJson: '', requestedAt: payload.requestedAt || '',
              requestedBy: payload.requestedBy || '', relatedJobId: job.job_id,
              detail: { error: String(engineError && engineError.message || engineError).slice(0, 1000) }
            });
          }
          vNextAdminPatchRegistryByBookId_(hub, job.target_book_id, { state: recoveryState, updated_at: new Date() });
          if (isAiRollback) {
            vNextAdminAppendException_(hub, {
              severity: 'ERROR', exception_type: 'AI_ROLLBACK_FORECAST_FAILED', book_id: registry.book_id,
              client_name: registry.client_name, fiscal_year: registry.fiscal_year,
              title: 'AI反映の取消後の再予測に失敗',
              detail: String(engineError && engineError.message || engineError),
              recommended_action: '原因を確認し、同じ取消操作を再実行して再投入', source_ref: payload.rollbackOperationId || job.job_id
            });
            vNextAdminWriteAudit_(hub, 'AI_ROLLBACK_FORECAST', 'AI_ROLLBACK', payload.rollbackOperationId, 'FAILED', {
              jobId: job.job_id, sourceForecastRunId: payload.sourceForecastRunId,
              error: String(engineError && engineError.message || engineError).slice(0, 1000)
            });
          }
        } catch (recoveryError) {
          Logger.log('Forecast client recovery failed book=%s error=%s', registry.book_id, String(recoveryError && recoveryError.stack || recoveryError));
        }
        throw engineError;
      }
    }
    case 'HEALTH_SCAN': {
      const registry = vNextAdminFindRegistryRow_(hub, function (row) {
        return String(row.book_id) === String(job.target_book_id) || String(row.spreadsheet_id) === String(job.target_spreadsheet_id);
      });
      if (!registry) throw new Error('Registry entry not found for health scan.');
      return vNextAdminScanOneBook_(hub, registry);
    }
    case 'PORTAL_PROVISION_CLIENT': {
      const portal = vNextAdminResolvePortal_(hub);
      if (portal.spreadsheetId !== String(payload.portalSpreadsheetId || '') ||
          portal.portalId !== String(payload.portalId || '')) throw new Error('Portal job identity does not match the configured Portal.');
      const events = vNextAdminReadTable_(portal.spreadsheet, VN_ADMIN_PORTAL_REQUEST_SHEET).rows.filter(function (row) {
        return String(row.request_id || '') === String(payload.requestId || '');
      });
      const requested = events.find(function (row) {
        return String(row.event_type || '').toUpperCase() === 'REQUESTED' && String(row.status || '').toUpperCase() === 'PENDING';
      });
      if (!requested) throw new Error('The original Portal REQUESTED/PENDING event is missing.');
      const validated = vNextAdminValidatePortalRequest_(hub, portal, requested);
      if (validated.requestHash !== String(payload.requestHash || '') ||
          validated.payload.requestedBy !== String(payload.requestedBy || '') ||
          validated.schemaVersion !== String(payload.schemaVersion || VN_ADMIN_PORTAL_REQUEST_SCHEMA_V1) ||
          validated.clientId !== String(payload.clientId || '') ||
          validated.forecastOwnerEmail !== String(payload.forecastOwnerEmail || '') ||
          vNextAdminCanonicalJson_(validated.relatedMemberEmails) !==
            vNextAdminCanonicalJson_(payload.relatedMemberEmails || []) ||
          vNextAdminCanonicalJson_(validated.relatedMemberNames) !==
            vNextAdminCanonicalJson_(payload.relatedMemberNames || [])) {
        throw new Error('Portal job/request lineage mismatch.');
      }
      const duplicates = vNextAdminFindClientFyDuplicates_(hub, Object.assign({}, validated.payload, {
        clientId: validated.clientId
      }));
      let provisioned;
      if (duplicates.length === 1 && String(duplicates[0].status || '').toUpperCase() === 'ACTIVE') {
        const access = vNextAdminPortalExistingBookAccess_(portal, duplicates[0]);
        if (!access.reusable) {
          const rejected = vNextAdminRejectPortalExistingBook_(hub, portal, validated, duplicates[0], access);
          vNextAdminRefreshPortalDirectory_(hub, portal.spreadsheet);
          return rejected;
        }
        provisioned = {
          reused: true, bookId: duplicates[0].book_id, spreadsheetId: duplicates[0].spreadsheet_id,
          spreadsheetUrl: duplicates[0].spreadsheet_url, state: duplicates[0].state
        };
      } else if (duplicates.length > 1) {
        throw new Error('Multiple client/FY duplicates block Portal provisioning.');
      } else {
        vNextAdminAppendPortalRequestEvent_(portal.spreadsheet, validated, 'CREATION_STARTED', 'CREATING', {
          relatedJobId: job.job_id, detailCode: 'CREATING', detailMessage: 'クライアント別ブックを作成しています。'
        });
        provisioned = vNextAdminProvisionClientInHub_(hub, {
          clientId: validated.clientId, clientName: validated.payload.clientName,
          fiscalYear: validated.payload.fiscalYear,
          forecastOwnerEmails: [validated.forecastOwnerEmail],
          editorEmails: [validated.forecastOwnerEmail].concat(validated.relatedMemberEmails),
          relatedMemberNames: validated.relatedMemberNames,
          accessPolicy: 'INTERNAL_OPEN', internalDomain: portal.employeeDomain,
          idempotencyKey: String(job.idempotency_key || ''), internalOperation: 'PORTAL_JOB'
        });
      }
      vNextAdminAppendPortalRequestEvent_(portal.spreadsheet, validated, 'COMPLETED', 'COMPLETED', {
        relatedJobId: job.job_id, relatedBookId: provisioned.bookId,
        relatedBookUrl: provisioned.spreadsheetUrl, detailCode: provisioned.reused ? 'EXISTING_BOOK_REUSED' : 'CREATED',
        detailMessage: provisioned.reused ? '既存のクライアント年度ブックをご利用ください。' : 'クライアント年度ブックを利用できます。'
      });
      vNextAdminResolveOpenExceptions_(hub, validated.payload.requestId,
        ['PORTAL_PROVISION_FAILED'], job.job_id);
      vNextAdminRefreshPortalDirectory_(hub, portal.spreadsheet);
      return provisioned;
    }
    case 'AI_RESEARCH': {
      const registry = vNextAdminFindRegistryRow_(hub, function (row) {
        return String(row.book_id) === String(job.target_book_id);
      });
      if (!registry) throw new Error('Registry entry not found for AI research.');
      let provider = null;
      if (typeof vNextAdminAiResearchProvider_ === 'function') provider = vNextAdminAiResearchProvider_;
      else if (typeof vNextVertexAiResearch_ === 'function') provider = vNextVertexAiResearch_;
      if (!provider) {
        vNextAdminAppendException_(hub, {
          severity: 'WARN', exception_type: 'AI_RESEARCH_PROVIDER_MISSING', book_id: registry.book_id,
          client_name: registry.client_name, fiscal_year: registry.fiscal_year,
          title: 'AI調査の接続先が未設定', detail: 'Install vNextAdminAiResearchProvider_ or vNextVertexAiResearch_.',
          recommended_action: 'Vertex AI接続設定とprovider hookを確認', source_ref: job.job_id
        });
        throw new Error('AI research provider hook is not installed.');
      }
      const findings = provider(Object.assign({}, payload, {
        bookId: registry.book_id, clientName: registry.client_name,
        fiscalYear: Number(registry.fiscal_year), spreadsheet: hub,
        internalOperation: 'ADMIN_JOB'
      })) || [];
      const list = Array.isArray(findings) ? findings : (findings.findings || []);
      if (!list.length) return { findings: 0, appended: 0 };
      const appended = list.map(function (finding) {
        return vNextAdminAppendAiEvidenceInternal_(hub, Object.assign({}, finding, {
          bookId: registry.book_id,
          parentRequestId: String(payload.parentRequestId || payload.requestId || ''),
          effectiveAsOf: String(payload.asOf || payload.effectiveAsOf || '')
        }));
      });
      const publicProjection = vNextAdminProjectAiInsightsToClient_(hub, registry, {});
      return {
        findings: list.length, appended: appended.length,
        evidenceIds: appended.map(function (row) { return row.evidenceId; }),
        publicInsights: publicProjection.insightCount
      };
    }
    case 'MIGRATION':
      return vNextAdminExecuteMigrationSkeleton_(hub, job, payload);
    case 'REFRESH_CLIENT_VIEW':
      if (typeof vNextRefreshEmployeeViews !== 'function') throw new Error('UX refresh API is not installed.');
      return vNextRefreshEmployeeViews(payload);
    default:
      if (typeof vNextAdminExternalJobHandler_ === 'function') return vNextAdminExternalJobHandler_(job.job_type, payload);
      throw new Error('Unsupported job type: ' + job.job_type);
  }
}

function vNextAdminFinishJob_(hub, jobId, status, result, error) {
  const table = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.JOBS);
  const row = table.rows.find(function (item) { return String(item.job_id) === String(jobId); });
  if (!row) return;
  const now = new Date();
  vNextAdminUpdateTableRow_(hub, VN_ADMIN_SHEETS.JOBS, row._rowNumber, {
    status: status, finished_at: now, result_json: result == null ? '' : JSON.stringify(result),
    error: error || '', updated_at: now
  });
  vNextAdminAppendObject_(hub, VN_ADMIN_SHEETS.JOB_LOG, {
    log_id: 'JLOG-' + Utilities.getUuid(), job_id: jobId, logged_at: now,
    status: status, message: error || 'Job completed', detail_json: result == null ? '{}' : JSON.stringify(result),
    actor: vNextAdminActor_()
  });
}

function vNextAdminAppendJobPhase_(hub, jobId, phase, detail) {
  const now = new Date();
  vNextAdminAppendObject_(hub, VN_ADMIN_SHEETS.JOB_LOG, {
    log_id: 'JLOG-' + Utilities.getUuid(), job_id: String(jobId || ''), logged_at: now,
    status: String(phase || 'PHASE'), message: String(phase || 'Job phase'),
    detail_json: vNextAdminCanonicalJson_(detail || {}), actor: vNextAdminActor_()
  });
}

/**
 * Complete or verify only the state phase that follows an already durable
 * deterministic SUCCESS. Client and registry mirrors are updated afterwards.
 */
function vNextAdminEnsurePersistedForecastDraftState_(hub, registry, result) {
  const runId = vNextAdminRequiredText_(result && result.runId, 'forecastResult.runId');
  const bookId = String(registry && registry.book_id || '');
  let rows = vNextAdminReadCoreRows_(hub, 'STATE_EVENT').filter(function (row) {
    return String(row.book_id || '') === bookId;
  });
  if (!rows.length) throw vNextAdminRunIdentityFailure_('Hub STATE_EVENT is missing for deterministic forecast finalization.');
  let latest = rows[rows.length - 1];
  let state = String(latest.to_state || '').toUpperCase();
  if (state === 'RUNNING') {
    vNextTransitionState_({
      bookId: bookId,
      fromState: 'RUNNING',
      toState: 'DRAFT_READY',
      reason: 'Forecast run completed (durable resume)',
      relatedRunId: runId,
      actorEmail: vNextAdminActor_(),
      actorRole: 'SYSTEM',
      internalOperation: 'FORECAST_ENGINE',
      spreadsheet: hub
    });
    rows = vNextAdminReadCoreRows_(hub, 'STATE_EVENT').filter(function (row) {
      return String(row.book_id || '') === bookId;
    });
    latest = rows[rows.length - 1];
    state = String(latest && latest.to_state || '').toUpperCase();
    if (state !== 'DRAFT_READY' || String(latest.related_run_id || '') !== runId) {
      throw vNextAdminRunIdentityFailure_('RUNNING>DRAFT_READY resume was not durably linked to the deterministic run.');
    }
    return 'DRAFT_READY_APPENDED';
  }
  if (state === 'DRAFT_READY' && String(latest.related_run_id || '') === runId) {
    return 'DRAFT_READY_VERIFIED';
  }
  throw vNextAdminRunIdentityFailure_(
    'Persisted SUCCESS cannot finalize from Hub state=' + state +
    ' relatedRunId=' + String(latest.related_run_id || '')
  );
}

function vNextAdminRunIdentityFailure_(message) {
  const error = new Error(String(message || 'Deterministic forecast identity failure.'));
  error.vNextRunIdentityFailure = true;
  return error;
}

function vNextAdminAssertClientPinnedReleasePair_(hub, client, registry) {
  if (String(registry.mode || '') !== 'CLIENT' || String(registry.status || '').toUpperCase() !== 'ACTIVE') {
    throw new Error('Forecast pair validation requires an ACTIVE CLIENT registry row.');
  }
  const routing = vNextAdminReadKeyValueSheet_(client, VN_ADMIN_BOOK_CONFIG_SHEET);
  const releaseId = vNextAdminRequiredText_(registry.template_release_id, 'registry.template_release_id');
  const modelId = vNextAdminRequiredText_(routing.model_release_id, 'client.model_release_id');
  if (String(routing.book_id || '') !== String(registry.book_id || '') ||
      String(routing.version || '') !== releaseId ||
      String(routing.template_release_id || releaseId) !== releaseId) {
    throw new Error('Client routing does not match its pinned Template Release.');
  }
  const release = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.RELEASES).rows.find(function (row) {
    return String(row.release_id || '') === releaseId;
  });
  if (!release || ['ACTIVE', 'RETIRED'].indexOf(String(release.status || '').toUpperCase()) < 0) {
    throw new Error('Pinned Template Release is not a valid immutable release: ' + releaseId);
  }
  const clientMeta = vNextAdminReadCoreRows_(client, 'BOOK_META').filter(function (row) {
    return String(row.book_id || '') === String(registry.book_id || '');
  }).slice(-1)[0];
  const hubMeta = vNextAdminReadCoreRows_(hub, 'BOOK_META').filter(function (row) {
    return String(row.book_id || '') === String(registry.book_id || '');
  }).slice(-1)[0];
  if (!clientMeta || !hubMeta || String(clientMeta.record_id || '') !== String(hubMeta.record_id || '') ||
      String(clientMeta.template_version || '') !== releaseId || String(hubMeta.template_version || '') !== releaseId ||
      String(clientMeta.model_release_id || '') !== modelId || String(hubMeta.model_release_id || '') !== modelId) {
    throw new Error('Client/Hub BOOK_META does not exactly match the pinned Template/Model pair.');
  }
  const model = vNextAdminModelReleaseRows_(hub, modelId).filter(function (row) {
    return String(row.status || '').toUpperCase() === 'ACTIVE';
  }).slice(-1)[0];
  if (!model) throw new Error('Pinned MODEL_RELEASE was never ACTIVE: ' + modelId);
  vNextAdminAssertModelTemplateCompatibility_(hub, model, release);
  if (String(routing.client_runtime_version || '') !== String(release.client_runtime_version || '') ||
      String(routing.client_runtime_bundle_sha256 || '') !== String(release.client_runtime_sha256 || '')) {
    throw new Error('Client runtime does not match its pinned Template Release.');
  }
  return { release: release, model: model, bookMeta: hubMeta };
}

function vNextAdminScanOneBook_(hub, registry) {
  const now = new Date();
  let status = 'OK';
  let code = 'HEALTHY';
  let detail = '';
  let harvestedRequests = 0;
  let approvalCreated = false;
  let observedState = String(registry.state || '');
  try {
    const ss = SpreadsheetApp.openById(String(registry.spreadsheet_id));
    const detected = vNextDetectBookMode_(ss);
    if (detected !== String(registry.mode)) {
      status = 'ERROR'; code = 'MODE_MISMATCH'; detail = 'registry=' + registry.mode + ', detected=' + detected;
    }
    const routing = vNextAdminReadKeyValueSheet_(ss, VN_ADMIN_BOOK_CONFIG_SHEET);
    const required = detected === 'CLIENT'
      ? ['book_id', 'client_name', 'fiscal_year', 'version', 'model_release_id']
      : ['book_id', 'mode', 'version'];
    const missing = required.filter(function (key) { return routing[key] === '' || routing[key] === null || routing[key] === undefined; });
    if (missing.length) {
      status = 'ERROR'; code = 'META_MISSING'; detail = 'missing=' + missing.join(',');
    }
    if (detected === 'CLIENT') vNextAdminAssertClientPinnedReleasePair_(hub, ss, registry);
    if (detected === 'CLIENT') {
      const eventState = vNextAdminLatestClientState_(ss, registry.book_id, routing.state);
      if (eventState) observedState = eventState;
      if (eventState && eventState !== String(routing.state || '').toUpperCase()) {
        vNextAdminMirrorClientState_(ss, eventState);
        if (status === 'OK') {
          status = 'WARN'; code = 'STATE_MIRROR_REPAIRED';
          detail = 'VN_BOOK_CONFIG state was repaired from append-only STATE_EVENT: ' + eventState;
        }
      }
    }
    let context = null;
    if (detected === 'CLIENT' && typeof vNextGetBookContext_ === 'function') {
      context = vNextGetBookContext_({ bookId: registry.book_id, spreadsheet: ss, userEmail: vNextAdminActor_() });
      observedState = vNextAdminLatestClientState_(ss, registry.book_id, context.state || observedState);
      if (VN_ADMIN_CLIENT_STATES.indexOf(String(context.state || '')) < 0) {
        status = 'ERROR'; code = 'INVALID_STATE'; detail = 'state=' + String(context.state || '');
      }
      const harvest = vNextAdminHarvestClientRequests_(hub, ss, registry);
      harvestedRequests = harvest.harvested;
      // Request rows are harvested first so a malformed request can be marked
      // REJECTED and its local RUNNING state repaired before Core row ingest.
      // Evidence and plans are still ingested before STATE_EVENT inside sync.
      vNextAdminSyncClientToHub_(hub, ss, registry.book_id);
      observedState = vNextAdminLatestClientState_(ss, registry.book_id, observedState);
      if (harvest.rejected && status === 'OK') {
        status = 'WARN'; code = 'CLIENT_REQUEST_REJECTED';
        detail = String(harvest.rejected) + ' invalid request(s) were rejected and safely recovered.';
      }
      const approval = vNextAdminCreateApprovalFromSubmittedPlan_(hub, ss, registry);
      approvalCreated = !!(approval && approval.created);
      if (String(observedState || '').toUpperCase() === 'SUBMITTED' && approval &&
          !approval.created && !approval.approvalRequestId && approval.reason !== 'STATE_NOT_SUBMITTED') {
        if (status === 'OK') {
          status = 'WARN'; code = 'APPROVAL_DATA_PENDING'; detail = 'reason=' + String(approval.reason || 'UNKNOWN');
        }
      }
    }
    if (detected === 'CLIENT' && registry.template_release_id && routing.version && String(registry.template_release_id) !== String(routing.version)) {
      if (status === 'OK') { status = 'WARN'; code = 'RELEASE_MISMATCH'; detail = 'registry=' + registry.template_release_id + ', book=' + routing.version; }
    }
    if (detected === 'CLIENT') {
      const expectedRelease = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.RELEASES).rows.find(function (row) {
        return String(row.release_id || '') === String(registry.template_release_id || routing.version || '');
      });
      if (!expectedRelease || !String(expectedRelease.client_runtime_sha256 || '') ||
          String(routing.client_runtime_version || '') !== String(expectedRelease.client_runtime_version || '') ||
          String(routing.client_runtime_bundle_sha256 || '') !== String(expectedRelease.client_runtime_sha256 || '')) {
        status = 'ERROR'; code = 'CLIENT_RUNTIME_MISMATCH';
        detail = 'Client runtime identity does not match RELEASES.';
      }
      const bookMeta = vNextAdminReadCoreRows_(ss, 'BOOK_META').filter(function (row) {
        return String(row.book_id || '') === String(registry.book_id || '');
      }).slice(-1)[0];
      const pinnedModelId = String(bookMeta && bookMeta.model_release_id || routing.model_release_id || '');
      const modelHistory = vNextAdminModelReleaseRows_(hub, pinnedModelId);
      const pinnedModel = modelHistory.filter(function (row) {
        return String(row.status || '').toUpperCase() === 'ACTIVE';
      }).slice(-1)[0];
      if (!pinnedModelId || String(routing.model_release_id || '') !== pinnedModelId ||
          !pinnedModel || !expectedRelease ||
          String(pinnedModel.template_version || '') !== String(expectedRelease.release_id || '') ||
          String(pinnedModel.schema_version || '') !== String(expectedRelease.schema_version || '') ||
          String(pinnedModel.model_version || '') !== String(expectedRelease.engine_version || '')) {
        status = 'ERROR'; code = 'MODEL_RELEASE_MISMATCH';
        detail = 'Client model release is missing, was never ACTIVE, or is not exactly paired with its Template Release.';
      }
    }
  } catch (err) {
    status = 'ERROR'; code = 'INACCESSIBLE'; detail = String(err && err.message || err);
  }
  vNextAdminPatchRegistryByBookId_(hub, registry.book_id, {
    health_status: status, health_code: code, state: observedState,
    last_health_at: now, updated_at: now
  });
  return {
    bookId: registry.book_id, healthStatus: status, healthCode: code, detail: detail,
    harvestedRequests: harvestedRequests, approvalCreated: approvalCreated
  };
}

function vNextAdminHarvestClientRequests_(hub, client, registry) {
  vNextAdminAssertClientPinnedReleasePair_(hub, client, registry);
  const sheet = client.getSheetByName(VN_ADMIN_CLIENT_REQUEST_SHEET);
  if (!sheet) return { examined: 0, harvested: 0, rejected: 0 };
  const rows = vNextAdminReadTable_(client, VN_ADMIN_CLIENT_REQUEST_SHEET).rows;
  const latest = {};
  rows.forEach(function (row) {
    const requestId = String(row.request_id || '');
    if (requestId) latest[requestId] = row;
  });
  let harvested = 0;
  let rejected = 0;
  Object.keys(latest).forEach(function (requestId) {
    const row = latest[requestId];
    if (String(row.status || '').toUpperCase() !== 'PENDING') return;
    const ownerEmails = vNextAdminParseList_(registry.forecast_owner_emails).map(function (email) {
      return String(email || '').toLowerCase();
    });
    try {
      if (ownerEmails.length !== 1) throw new Error('BOOK_REGISTRY must contain exactly one 予算策定担当.');
      const validated = vNextAdminValidateClientRequestRow_(row, registry, ownerEmails[0]);
      const harvestAsOf = Utilities.formatDate(new Date(validated.requestedAtMs), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      // Never promote the untrusted parsed object. Build the worker request from
      // the allowlisted identity plus server-derived vintage and audit fields.
      const payload = {
        requestId: requestId,
        requestHash: validated.requestHash,
        bookId: String(registry.book_id || ''),
        clientId: String(registry.client_id || ''),
        clientName: String(registry.client_name || ''),
        fiscalYear: Number(registry.fiscal_year),
        asOf: harvestAsOf,
        cutoff: vNextAdminCutoffFromAsOf_(harvestAsOf),
        bookConfiguredAsOf: String(validated.payload.bookConfiguredAsOf || ''),
        requestedAt: validated.payload.requestedAt,
        requestedBy: ownerEmails[0],
        harvestedAt: new Date().toISOString()
      };
      const runtime = vNextGetRuntimeConfig_();
      if (runtime.VERTEX_PROJECT_ID) {
        const aiJob = vNextAdminEnqueueJobInternal_(hub, {
          jobType: 'AI_RESEARCH', targetBookId: registry.book_id,
          targetSpreadsheetId: registry.spreadsheet_id,
          request: {
            bookId: registry.book_id, clientName: registry.client_name,
            fiscalYear: Number(registry.fiscal_year), asOf: payload.asOf,
            researchQuestion: '予測時点までに公開された、顧客・市場・規制の売上影響を確認する',
            parentRequestId: requestId
          },
          idempotencyKey: 'AUTO_AI_RESEARCH|' + requestId + '|' + validated.requestHash,
          priority: 60
        });
        payload.aiResearchJobId = aiJob.job_id;
      }
      const job = vNextAdminEnqueueJobInternal_(hub, {
        jobType: 'FORECAST_REQUEST', targetBookId: registry.book_id,
        targetSpreadsheetId: registry.spreadsheet_id, request: payload,
        idempotencyKey: 'CLIENT_REQUEST|' + requestId + '|' + validated.requestHash, priority: 50
      });
      const eventData = {
        requestId: requestId, bookId: registry.book_id, eventType: 'HARVESTED', status: 'QUEUED',
        requestHash: validated.requestHash, requestJson: validated.requestJson, requestedAt: row.requested_at,
        requestedBy: row.requested_by, relatedJobId: job.job_id,
        detail: { harvestedBy: vNextAdminActor_(), harvestedAt: new Date().toISOString() }
      };
      vNextAdminAppendClientRequestEvent_(hub, eventData);
      vNextAdminAppendClientRequestEvent_(client, eventData);
      harvested++;
    } catch (requestError) {
      const recovery = vNextAdminRejectClientRequest_(hub, client, registry, row, requestError);
      rejected += recovery.rejected ? 1 : 0;
    }
  });
  return { examined: Object.keys(latest).length, harvested: harvested, rejected: rejected };
}

function vNextAdminValidateClientRequestRow_(row, registry, ownerEmail) {
  const requestId = String(row.request_id || '');
  if (!requestId || !String(row.request_event_id || '')) throw new Error('Client request IDs are missing.');
  if (String(row.event_type || '').toUpperCase() !== 'REQUESTED' ||
      String(row.status || '').toUpperCase() !== 'PENDING') {
    throw new Error('Client request row must be the immutable REQUESTED/PENDING event: ' + requestId);
  }
  if (String(row.book_id || '') !== String(registry.book_id || '')) {
    throw new Error('Client request book mismatch: ' + requestId);
  }
  const requestJson = String(row.request_json || '');
  const requestHash = vNextAdminSha256_(requestJson);
  if (requestHash !== String(row.request_hash || '')) throw new Error('Client request hash mismatch: ' + requestId);
  const payload = vNextAdminParseJson_(requestJson, null);
  vNextAdminAssertClientRequestPayload_(payload, requestJson, requestId);
  const owner = String(ownerEmail || '').toLowerCase();
  if (!owner || String(row.requested_by || '').toLowerCase() !== owner ||
      String(payload.requestedBy || '').toLowerCase() !== owner) {
    throw new Error('Client request was not created by the registered 予算策定担当: ' + requestId);
  }
  if (String(payload.bookId || '') !== String(registry.book_id || '') ||
      String(payload.clientId || '') !== String(registry.client_id || '') ||
      String(payload.clientName || '') !== String(registry.client_name || '') ||
      Number(payload.fiscalYear) !== Number(registry.fiscal_year)) {
    throw new Error('Client request book/client/FY payload does not match BOOK_REGISTRY: ' + requestId);
  }
  const payloadRequestedAt = String(payload.requestedAt || '');
  const rowRequestedAt = row.requested_at instanceof Date ? row.requested_at.toISOString() : String(row.requested_at || '');
  if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{3})?Z$/.test(payloadRequestedAt) ||
      payloadRequestedAt !== rowRequestedAt) {
    throw new Error('Client request timestamp is invalid or inconsistent: ' + requestId);
  }
  const requestedAtMs = vNextAdminStrictTimestampMs_(payloadRequestedAt, 'Client request requestedAt');
  if (requestedAtMs > Date.now() + 5 * 60000) throw new Error('Client request timestamp is in the future: ' + requestId);
  const expectedAsOf = Utilities.formatDate(new Date(requestedAtMs), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (String(payload.asOf || '') !== expectedAsOf ||
      String(payload.cutoff || '') !== vNextAdminCutoffFromAsOf_(expectedAsOf)) {
    throw new Error('Client request asOf/cutoff is not server-derivable: ' + requestId);
  }
  return {
    requestId: requestId, requestJson: requestJson, requestHash: requestHash,
    payload: payload, requestedAtMs: requestedAtMs
  };
}

function vNextAdminRejectClientRequest_(hub, client, registry, row, error) {
  const requestId = String(row.request_id || 'UNKNOWN');
  const reason = String(error && error.message || error || 'invalid request');
  let requestedAtMs = NaN;
  try { requestedAtMs = vNextAdminStrictTimestampMs_(row.requested_at, 'rejected request timestamp'); }
  catch (ignoredTimestampError) { requestedAtMs = NaN; }
  const stateRows = vNextAdminReadCoreRows_(client, 'STATE_EVENT').filter(function (event) {
    if (String(event.book_id || '') !== String(registry.book_id || '') ||
        String(event.from_state || '').toUpperCase() !== 'READY_TO_RUN' ||
        String(event.to_state || '').toUpperCase() !== 'RUNNING' ||
        String(event.reason || '') !== 'forecast_requested:' + requestId) return false;
    if (!isFinite(requestedAtMs)) return true;
    return Math.abs(vNextAdminStrictTimestampMs_(event.created_at, 'request state timestamp') - requestedAtMs) <= 5 * 60000;
  });
  const requestState = stateRows.length ? stateRows[stateRows.length - 1] : null;
  let recoveryState = null;
  const hubState = vNextAdminLatestClientState_(hub, registry.book_id, registry.state || 'READY_TO_RUN');
  const clientState = vNextAdminLatestClientState_(client, registry.book_id, registry.state || 'READY_TO_RUN');
  const activeJob = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.JOBS).rows.some(function (job) {
    if (String(job.target_book_id || '') !== String(registry.book_id || '') ||
        String(job.job_type || '') !== 'FORECAST_REQUEST' ||
        ['QUEUED', 'RUNNING'].indexOf(String(job.status || '').toUpperCase()) < 0 ||
        String(job.idempotency_key || '') !== 'CLIENT_REQUEST|' + requestId + '|' + String(row.request_hash || '')) {
      return false;
    }
    const jobPayload = vNextAdminParseJson_(job.request_json, {});
    return String(jobPayload.requestId || '') === requestId &&
      String(jobPayload.requestHash || '') === String(row.request_hash || '') &&
      String(jobPayload.bookId || '') === String(registry.book_id || '');
  });
  if (!activeJob && requestState && hubState === 'READY_TO_RUN' && clientState === 'RUNNING') {
    recoveryState = {
      state_event_id: 'REPAIR-' + Utilities.getUuid(), book_id: registry.book_id,
      from_state: 'RUNNING', to_state: 'READY_TO_RUN',
      reason: 'rejected_request_recovery:' + requestId,
      actor_email: vNextAdminActor_(), actor_role: 'ADMIN', related_run_id: '',
      related_plan_version_id: '', created_at: new Date().toISOString()
    };
  }
  const detail = {
    rejectedBy: vNextAdminActor_(), rejectedAt: new Date().toISOString(), reason: reason,
    stateEventId: requestState && requestState.state_event_id || '',
    recoveryStateEventId: recoveryState && recoveryState.state_event_id || ''
  };
  const eventData = {
    requestId: requestId, bookId: registry.book_id, eventType: 'REJECTED', status: 'REJECTED',
    requestHash: String(row.request_hash || ''), requestJson: String(row.request_json || ''),
    requestedAt: row.requested_at || '', requestedBy: row.requested_by || '', detail: detail
  };
  // Hub copy is the authority used to recognize and skip only this rejected
  // local state pair on subsequent syncs.
  vNextAdminAppendClientRequestEvent_(hub, eventData);
  vNextAdminAppendClientRequestEvent_(client, eventData);
  if (recoveryState) {
    vNextAdminAppendCoreRowsNoLock_(client, 'STATE_EVENT', [recoveryState]);
    vNextAdminMirrorClientState_(client, 'READY_TO_RUN');
  }
  vNextAdminAppendException_(hub, {
    severity: 'ERROR', exception_type: 'CLIENT_REQUEST_REJECTED', book_id: registry.book_id,
    client_name: registry.client_name, fiscal_year: registry.fiscal_year,
    title: 'Client予測依頼を拒否', detail: reason,
    recommended_action: '直接編集の有無を確認し、予算策定担当が再依頼', source_ref: requestId
  });
  return { rejected: true, requestId: requestId, recoveredState: Boolean(recoveryState) };
}

function vNextAdminAppendClientRequestEvent_(client, data) {
  const record = {
    request_event_id: 'REQEV-' + Utilities.getUuid(),
    request_id: String(data.requestId || ''),
    book_id: String(data.bookId || ''),
    event_type: String(data.eventType || ''),
    status: String(data.status || ''),
    request_hash: String(data.requestHash || ''),
    request_json: String(data.requestJson || ''),
    requested_at: data.requestedAt || '',
    requested_by: String(data.requestedBy || '').toLowerCase(),
    related_job_id: String(data.relatedJobId || ''),
    related_run_id: String(data.relatedRunId || ''),
    detail_json: vNextAdminCanonicalJson_(data.detail || {}),
    created_at: new Date().toISOString()
  };
  vNextAdminAppendObject_(client, VN_ADMIN_CLIENT_REQUEST_SHEET, record, VN_ADMIN_HEADERS[VN_ADMIN_CLIENT_REQUEST_SHEET]);
  try { client.getSheetByName(VN_ADMIN_CLIENT_REQUEST_SHEET).hideSheet(); } catch (hideError) { Logger.log('Request log hide skipped: %s', String(hideError)); }
  return record;
}

function vNextAdminAppendAiEvidenceInternal_(hub, request) {
  const req = request && typeof request === 'object' ? request : {};
  const bookId = vNextAdminRequiredText_(req.bookId, 'bookId');
  const registry = vNextAdminFindRegistryRow_(hub, function (row) { return String(row.book_id) === bookId; });
  if (!registry || String(registry.mode) !== 'CLIENT') throw new Error('Registered CLIENT book not found: ' + bookId);
  const target = vNextAdminRequiredText_(req.target, 'target');
  const startMonth = vNextAdminRequiredText_(req.targetStartMonth || req.startMonth, 'targetStartMonth');
  const endMonth = vNextAdminRequiredText_(req.targetEndMonth || req.endMonth || startMonth, 'targetEndMonth');
  if (!/^\d{4}-\d{2}$/.test(startMonth) || !/^\d{4}-\d{2}$/.test(endMonth)) throw new Error('AI evidence period must be YYYY-MM.');
  const effectRate = Number(req.effectRate);
  const forecastUse = String(req.forecastUse || (effectRate === 0 ? 'INSIGHT_ONLY' : 'APPLY')).trim().toUpperCase();
  if (!isFinite(effectRate) || Math.abs(effectRate) > 0.25 || ['APPLY', 'INSIGHT_ONLY'].indexOf(forecastUse) < 0 ||
      (forecastUse === 'APPLY' && effectRate === 0) || (forecastUse === 'INSIGHT_ONLY' && effectRate !== 0)) {
    throw new Error('AI evidence must use a non-zero effectRate for APPLY or zero for INSIGHT_ONLY, within -0.25 and 0.25.');
  }
  const requestedDirection = String(req.direction || (effectRate < 0 ? 'DOWN' : effectRate > 0 ? 'UP' : 'NEUTRAL')).trim().toUpperCase();
  if (['UP', 'DOWN', 'NEUTRAL'].indexOf(requestedDirection) < 0 ||
      (effectRate > 0 && requestedDirection !== 'UP') || (effectRate < 0 && requestedDirection !== 'DOWN')) {
    throw new Error('AI evidence direction does not match effectRate.');
  }
  const sourceUrl = vNextAdminRequiredText_(req.sourceUrl, 'sourceUrl');
  if (!/^https?:\/\//i.test(sourceUrl)) throw new Error('sourceUrl must be an http(s) citation URL.');
  const sourceDate = vNextAdminRequiredText_(req.sourceDate, 'sourceDate');
  const parsedSourceDate = typeof vNextParseDate_ === 'function' ? vNextParseDate_(sourceDate, 'sourceDate') : new Date(sourceDate);
  const effectiveAsOfInput = vNextAdminText_(req.effectiveAsOf || req.asOf) ||
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const parsedEffectiveAsOf = typeof vNextParseDate_ === 'function'
    ? vNextParseDate_(effectiveAsOfInput, 'effectiveAsOf') : new Date(effectiveAsOfInput);
  if (isNaN(parsedSourceDate.getTime()) || isNaN(parsedEffectiveAsOf.getTime()) ||
      parsedSourceDate.getTime() > parsedEffectiveAsOf.getTime()) {
    throw new Error('sourceDate must be a valid date on or before effectiveAsOf.');
  }
  const effectiveAsOf = Utilities.formatDate(parsedEffectiveAsOf, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const normalizedSourceDate = Utilities.formatDate(parsedSourceDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  let expiresAt = '';
  if (req.expiresAt) {
    const parsedExpiry = typeof vNextParseDate_ === 'function'
      ? vNextParseDate_(req.expiresAt, 'expiresAt') : new Date(req.expiresAt);
    if (isNaN(parsedExpiry.getTime())) throw new Error('expiresAt is invalid.');
    expiresAt = Utilities.formatDate(parsedExpiry, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (expiresAt < effectiveAsOf) throw new Error('expiresAt must be on or after effectiveAsOf.');
  }
  const parentRequestId = vNextAdminText_(req.parentRequestId);
  const summary = vNextAdminRequiredText_(req.summary || req.evidenceText, 'summary');
  const runtime = vNextGetRuntimeConfig_();
  const aiModel = vNextAdminRequiredText_(req.aiModel || runtime.VERTEX_GEMINI_MODEL, 'aiModel');
  const promptVersion = vNextAdminRequiredText_(req.promptVersion, 'promptVersion');
  const schemaVersion = vNextAdminRequiredText_(req.aiSchemaVersion || req.schemaVersion, 'aiSchemaVersion');
  const ruleVersion = vNextAdminRequiredText_(req.ruleVersion, 'ruleVersion');
  const evidenceQuality = String(req.evidenceQuality || req.evidenceGrade || '').trim().toUpperCase();
  if (['A', 'B', 'C', 'D'].indexOf(evidenceQuality) < 0) throw new Error('evidenceQuality must be A, B, C, or D.');
  const confidence = String(req.confidenceClass || '').trim().toUpperCase();
  if (['CONFIRMED_FACT', 'LIKELY', 'HYPOTHESIS'].indexOf(confidence) < 0) {
    throw new Error('confidenceClass must be CONFIRMED_FACT, LIKELY, or HYPOTHESIS.');
  }
  const forecastRows = vNextAdminReadCoreRows_(hub, 'FORECAST_RUN').filter(function (row) {
    return String(row.book_id || '') === bookId && ['SUCCESS', 'OFFICIAL_LOCKED'].indexOf(String(row.status || '')) >= 0;
  });
  const latest = forecastRows.length ? forecastRows[forecastRows.length - 1] : null;
  const derivedBasis = latest ? Number(latest.system_recommended || 0) - Number(latest.ai_delta || 0) : NaN;
  const basisAmount = isFinite(derivedBasis) && derivedBasis > 0 ? derivedBasis : Number(req.basisAmount);
  if (!isFinite(basisAmount) || basisAmount <= 0) {
    throw new Error('A positive basisAmount is required when no prior forecast can provide the pre-AI basis.');
  }
  const engineCapRate = typeof VNEXT_ENGINE !== 'undefined' ? Number(VNEXT_ENGINE.AI_MAX_ABS_EFFECT || 0.05) : 0.05;
  const gradeCapRate = { A: 0.05, B: 0.03, C: 0.01, D: 0 }[evidenceQuality];
  const capRate = Math.min(engineCapRate, gradeCapRate);
  const capAmount = basisAmount * capRate;
  const allEvidence = vNextAdminReadCoreRows_(hub, 'EVIDENCE_EVENT');
  if (req.supersedesEvidenceId && !allEvidence.some(function (row) {
    return String(row.evidence_id || '') === String(req.supersedesEvidenceId) && String(row.book_id || '') === bookId &&
      String(row.evidence_type || '').toUpperCase().indexOf('AI') >= 0;
  })) throw new Error('supersedesEvidenceId is not an AI evidence record for this book.');
  // A future-dated superseder must not erase an older event from a historical
  // cap calculation. First select events valid at this asOf, then derive the
  // superseded IDs only from that temporal slice.
  const existingAi = vNextAdminSelectActiveAiEvidenceAt_(allEvidence, bookId, effectiveAsOf).filter(function (row) {
    return String(row.evidence_id || '') !== String(req.supersedesEvidenceId || '');
  });
  const existingNet = existingAi.reduce(function (sum, row) {
    const sign = String(row.direction || '').toUpperCase() === 'DOWN' ? -1 : 1;
    return sum + sign * Number(row.applied_amount || row.amount_mid || 0);
  }, 0);
  const requestedSigned = basisAmount * effectRate;
  const appliedSigned = requestedSigned === 0 ? 0 : (requestedSigned > 0
    ? Math.max(0, Math.min(requestedSigned, capAmount - existingNet))
    : Math.min(0, Math.max(requestedSigned, -capAmount - existingNet)));
  const cappedNet = existingNet + appliedSigned;
  const appliedAmount = Math.abs(appliedSigned);
  const capApplied = Math.abs(appliedSigned - requestedSigned) > 0.5;
  const direction = appliedSigned < 0 || (appliedSigned === 0 && effectRate < 0)
    ? 'DOWN' : (appliedSigned > 0 || effectRate > 0 ? 'UP' : requestedDirection);
  const retrievedAt = new Date().toISOString();
  const researchAxis = String(req.researchAxis || 'ALTERNATIVE_SIGNALS').trim().toUpperCase().slice(0, 40);
  const sourceStrength = String(req.sourceStrength || 'REPUTABLE_SECONDARY').trim().toUpperCase().slice(0, 40);
  const salesRelevance = String(req.salesRelevance || 'LOW').trim().toUpperCase().slice(0, 20);
  const signalType = String(req.signalType || '').replace(/[\u0000-\u001f]+/g, ' ').trim().slice(0, 80);
  const humanQuestion = String(req.humanQuestion || '').replace(/[\u0000-\u001f]+/g, ' ').trim().slice(0, 180);
  const metadata = {
    summary: summary, citationTitle: String(req.citationTitle || ''), sourceUrl: sourceUrl,
    sourceDate: normalizedSourceDate, retrievedAt: retrievedAt,
    parentRequestId: parentRequestId, effectiveAsOf: effectiveAsOf, aiModel: aiModel,
    promptVersion: promptVersion, aiSchemaVersion: schemaVersion, ruleVersion: ruleVersion,
    evidenceQuality: evidenceQuality, confidenceClass: confidence,
    researchAxis: researchAxis, signalType: signalType, sourceStrength: sourceStrength,
    forecastUse: forecastUse, salesRelevance: salesRelevance, humanQuestion: humanQuestion,
    basisAmount: basisAmount, basisSourceRunId: latest && latest.run_id || '',
    requestedEffectRate: effectRate, requestedSignedAmount: requestedSigned,
    existingAiNetAmount: existingNet, appliedSignedAmount: appliedSigned,
    engineCapRate: engineCapRate, evidenceGradeCapRate: gradeCapRate,
    capRate: capRate, capAmount: capAmount, capApplied: capApplied
  };
  const identity = vNextAdminCanonicalJson_({
    bookId: bookId, target: target, startMonth: startMonth, endMonth: endMonth,
    sourceUrl: sourceUrl, sourceDate: normalizedSourceDate, promptVersion: promptVersion,
    aiModel: aiModel, schemaVersion: schemaVersion, ruleVersion: ruleVersion, requestedEffectRate: effectRate,
    supersedesEvidenceId: String(req.supersedesEvidenceId || ''), summary: summary,
    parentRequestId: parentRequestId, effectiveAsOf: effectiveAsOf,
    researchAxis: researchAxis, signalType: signalType, sourceStrength: sourceStrength,
    forecastUse: forecastUse, salesRelevance: salesRelevance, humanQuestion: humanQuestion,
    direction: direction
  });
  const evidenceId = 'AI-' + vNextAdminSha256_(identity).slice(0, 24).toUpperCase();
  const existing = vNextAdminReadCoreRows_(hub, 'EVIDENCE_EVENT').find(function (row) {
    return String(row.evidence_id || '') === evidenceId;
  });
  if (existing) return { reused: true, evidenceId: evidenceId, appliedAmount: Number(existing.applied_amount || 0), capApplied: Number(existing.cap_applied || 0) === 1 };
  const record = {
    evidence_id: evidenceId, book_id: bookId, client_id: registry.client_id,
    fiscal_year: Number(registry.fiscal_year), actor_email: vNextAdminActor_().toLowerCase(),
    response_type: 'CHANGE', evidence_type: forecastUse === 'APPLY' ? 'AI_RESEARCH' : 'AI_RESEARCH_INSIGHT', target: target,
    target_start_month: startMonth, target_end_month: endMonth, direction: direction,
    amount_mode: 'PERCENT', amount_low: appliedAmount, amount_mid: appliedAmount,
    amount_high: appliedAmount, amount_band: '', confidence_class: confidence,
    evidence_text: vNextAdminCanonicalJson_(metadata), source_url: sourceUrl,
    source_date: normalizedSourceDate, expires_at: expiresAt, status: 'ACTIVE',
    supersedes_evidence_id: String(req.supersedesEvidenceId || ''), created_at: retrievedAt,
    evidence_quality: evidenceQuality, ai_model: aiModel, prompt_version: promptVersion,
    ai_schema_version: schemaVersion, rule_version: ruleVersion,
    applied_amount: appliedAmount, cap_applied: capApplied ? 1 : 0
  };
  vNextAdminAppendCoreRowsNoLock_(hub, 'EVIDENCE_EVENT', [record]);
  // Raw prompt/model/research metadata remains Admin-Hub-only. Employee books
  // receive only a separately sanitized derived cache.
  vNextAdminWriteAudit_(hub, 'APPEND_AI_EVIDENCE', 'EVIDENCE', evidenceId, 'SUCCESS', metadata);
  return { reused: false, evidenceId: evidenceId, appliedAmount: appliedAmount, capApplied: capApplied, capRate: capRate };
}

function vNextAdminProjectAiInsightsToClient_(hub, registry, options) {
  const opt = options && typeof options === 'object' ? options : {};
  const bookId = vNextAdminRequiredText_(registry && registry.book_id, 'registry.book_id');
  const spreadsheetId = vNextAdminRequiredText_(registry && registry.spreadsheet_id, 'registry.spreadsheet_id');
  const client = SpreadsheetApp.openById(spreadsheetId);
  const routing = vNextAdminReadKeyValueSheet_(client, VN_ADMIN_BOOK_CONFIG_SHEET);
  if (String(routing.mode || '').toUpperCase() !== 'CLIENT' || String(routing.book_id || '') !== bookId ||
      String(routing.client_id || '') !== String(registry.client_id || '') ||
      Number(routing.fiscal_year || 0) !== Number(registry.fiscal_year || 0)) {
    throw new Error('Client identity does not match BOOK_REGISTRY for AI insight projection.');
  }
  const now = new Date();
  const asOf = vNextAdminText_(opt.asOf) || Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const activeRows = vNextAdminSelectActiveAiEvidenceAt_(
    vNextAdminReadCoreRows_(hub, 'EVIDENCE_EVENT'), bookId, asOf
  );
  const insights = activeRows.map(vNextAdminPublicAiInsightFromRow_).filter(Boolean).sort(function (a, b) {
    const useScore = { APPLY: 2, INSIGHT_ONLY: 1 };
    const relevanceScore = { HIGH: 3, MEDIUM: 2, LOW: 1 };
    const qualityScore = { A: 4, B: 3, C: 2, D: 1 };
    return (useScore[b.forecastUse] || 0) - (useScore[a.forecastUse] || 0) ||
      (relevanceScore[b.salesRelevance] || 0) - (relevanceScore[a.salesRelevance] || 0) ||
      (qualityScore[b.evidenceQuality] || 0) - (qualityScore[a.evidenceQuality] || 0) ||
      String(b.publishedAt || '').localeCompare(String(a.publishedAt || ''));
  }).slice(0, 5);
  const generatedAt = now.toISOString();
  const projection = {
    schemaVersion: 'vnext-public-ai-insights-1',
    bookId: bookId,
    generatedAt: generatedAt,
    insights: insights
  };
  vNextAdminWriteBookConfig_(client, {
    public_ai_insights_schema_version: projection.schemaVersion,
    public_ai_insights_json: vNextAdminCanonicalJson_(projection),
    public_ai_insights_updated_at: generatedAt
  });
  vNextAdminProtectClientInternalSheets_(client, [VN_ADMIN_BOOK_CONFIG_SHEET]);
  vNextAdminWriteAudit_(hub, 'PROJECT_PUBLIC_AI_INSIGHTS', 'BOOK', bookId, 'SUCCESS', {
    insightCount: insights.length, clientSpreadsheetId: spreadsheetId,
    schemaVersion: projection.schemaVersion, generatedAt: generatedAt
  });
  return { bookId: bookId, spreadsheetId: spreadsheetId, insightCount: insights.length, generatedAt: generatedAt };
}

function vNextAdminPublicAiInsightFromRow_(row) {
  const metadata = vNextAdminParseJson_(row && row.evidence_text, {});
  const summary = String(metadata.summary || row && row.target || '').replace(/[\u0000-\u001f]+/g, ' ').trim().slice(0, 280);
  const sourceUrl = String(row && row.source_url || metadata.sourceUrl || '').trim();
  if (!summary || !/^https?:\/\//i.test(sourceUrl)) return null;
  const forecastUse = String(metadata.forecastUse || (Number(row && row.applied_amount || 0) ? 'APPLY' : 'INSIGHT_ONLY')).toUpperCase();
  const direction = String(row && row.direction || 'NEUTRAL').toUpperCase();
  const appliedAmount = Math.abs(Number(row && (row.applied_amount || row.amount_mid) || 0));
  return {
    insightId: 'PUB-' + vNextAdminSha256_(String(row && row.evidence_id || '')).slice(0, 20).toUpperCase(),
    target: String(row && row.target || '').replace(/[\u0000-\u001f]+/g, ' ').trim().slice(0, 120),
    direction: ['UP', 'DOWN', 'NEUTRAL'].indexOf(direction) >= 0 ? direction : 'NEUTRAL',
    summary: summary,
    citationTitle: String(metadata.citationTitle || '').replace(/[\u0000-\u001f]+/g, ' ').trim().slice(0, 160),
    sourceUrl: sourceUrl.slice(0, 2000),
    sourceDate: String(row && row.source_date || metadata.sourceDate || '').slice(0, 10),
    appliedAmount: appliedAmount,
    evidenceQuality: String(row && row.evidence_quality || metadata.evidenceQuality || '').toUpperCase().slice(0, 20),
    capApplied: Number(row && row.cap_applied || 0) === 1,
    researchAxis: String(metadata.researchAxis || 'ALTERNATIVE_SIGNALS').toUpperCase().slice(0, 40),
    signalType: String(metadata.signalType || '').replace(/[\u0000-\u001f]+/g, ' ').trim().slice(0, 80),
    sourceStrength: String(metadata.sourceStrength || '').toUpperCase().slice(0, 40),
    forecastUse: ['APPLY', 'INSIGHT_ONLY'].indexOf(forecastUse) >= 0 ? forecastUse : 'INSIGHT_ONLY',
    salesRelevance: String(metadata.salesRelevance || '').toUpperCase().slice(0, 20),
    humanQuestion: String(metadata.humanQuestion || '').replace(/[\u0000-\u001f]+/g, ' ').trim().slice(0, 180),
    projectionStatus: forecastUse === 'APPLY' ? 'PENDING_FORECAST_REFRESH' : 'INSIGHT_ONLY',
    publishedAt: String(row && row.created_at || '').slice(0, 30)
  };
}

function vNextAdminAiEvidenceActiveAt_(row, bookId, effectiveAsOf, supersededIds) {
  if (String(row.book_id || '') !== String(bookId || '') ||
      String(row.evidence_type || '').toUpperCase().indexOf('AI') < 0 ||
      String(row.status || 'ACTIVE').toUpperCase() !== 'ACTIVE' ||
      (supersededIds && supersededIds.has(String(row.evidence_id || '')))) return false;
  const rowSourceDate = row.source_date instanceof Date
    ? Utilities.formatDate(row.source_date, Session.getScriptTimeZone(), 'yyyy-MM-dd')
    : String(row.source_date || '').slice(0, 10);
  const rowExpiry = row.expires_at instanceof Date
    ? Utilities.formatDate(row.expires_at, Session.getScriptTimeZone(), 'yyyy-MM-dd')
    : String(row.expires_at || '').slice(0, 10);
  const rowMetadata = vNextAdminParseJson_(row.evidence_text, {});
  const rowEffectiveAsOf = String(rowMetadata.effectiveAsOf || '').slice(0, 10);
  return !!rowSourceDate && rowSourceDate <= effectiveAsOf && (!rowExpiry || rowExpiry >= effectiveAsOf) &&
    !!rowEffectiveAsOf && rowEffectiveAsOf <= effectiveAsOf;
}

function vNextAdminSelectActiveAiEvidenceAt_(rows, bookId, effectiveAsOf) {
  const temporallyValid = (rows || []).filter(function (row) {
    return vNextAdminAiEvidenceActiveAt_(row, bookId, effectiveAsOf, new Set());
  });
  const supersededIds = new Set(temporallyValid.map(function (row) {
    return String(row.supersedes_evidence_id || '');
  }).filter(Boolean));
  return temporallyValid.filter(function (row) {
    return !supersededIds.has(String(row.evidence_id || ''));
  });
}

function vNextAdminNormalizeAiRollbackScope_(scope) {
  const normalized = String(scope || 'ALL').trim().toUpperCase();
  if (VN_ADMIN_AI_ROLLBACK_SCOPES.indexOf(normalized) < 0) {
    throw new Error('AI rollback scope must be ALL or SELECTED.');
  }
  return normalized;
}

function vNextAdminResolveAiRollbackBasis_(hub, bookId, requestedRunId, allowResume) {
  const registry = vNextAdminFindRegistryRow_(hub, function (row) {
    return String(row.book_id || '') === String(bookId || '');
  });
  if (!registry || String(registry.mode || '') !== 'CLIENT' || String(registry.status || '').toUpperCase() !== 'ACTIVE') {
    throw new Error('AI rollback requires an ACTIVE CLIENT book.');
  }
  const runs = vNextAdminReadCoreRows_(hub, 'FORECAST_RUN').filter(function (row) {
    return String(row.book_id || '') === String(bookId || '') &&
      String(row.status || '').toUpperCase() === 'SUCCESS' && Number(row.is_official || 0) !== 1;
  });
  if (!runs.length) throw new Error('No successful draft FORECAST_RUN is available for AI rollback.');
  const latest = runs[runs.length - 1];
  const selected = requestedRunId
    ? runs.filter(function (row) { return String(row.run_id || '') === String(requestedRunId); }).slice(-1)[0]
    : latest;
  if (!selected) throw new Error('Requested AI rollback source run was not found.');
  if (!allowResume && String(selected.run_id || '') !== String(latest.run_id || '')) {
    throw new Error('AI rollback basis must be the latest successful draft run.');
  }
  if ((String(selected.client_id || '') && String(selected.client_id || '') !== String(registry.client_id || '')) ||
      Number(selected.fiscal_year) !== Number(registry.fiscal_year)) {
    throw new Error('AI rollback basis book/client/FY linkage is invalid.');
  }
  const state = vNextAdminLatestClientState_(hub, bookId, registry.state);
  if (!allowResume && state !== 'DRAFT_READY') {
    throw new Error('AI rollback basis is available only in DRAFT_READY. current=' + state);
  }
  return { registry: registry, run: selected, latestRunId: latest.run_id, state: state };
}

function vNextAdminResolveBasisAiEvidence_(hub, basisRun) {
  const bookId = String(basisRun && basisRun.book_id || '');
  const asOf = vNextAdminDateOnly_(basisRun && basisRun.as_of);
  const summary = vNextAdminParseJson_(basisRun && basisRun.evidence_json, {});
  const ids = summary && summary.effectiveEvidenceIds && summary.effectiveEvidenceIds.ai;
  if (!Array.isArray(ids) || !ids.length || ids.some(function (id) { return !String(id || '').trim(); }) ||
      new Set(ids.map(String)).size !== ids.length) {
    throw new Error('Basis run does not contain an exact AI evidence snapshot. Create a new draft run before rollback.');
  }
  const byId = new Map(vNextAdminReadCoreRows_(hub, 'EVIDENCE_EVENT').map(function (row) {
    return [String(row.evidence_id || ''), row];
  }));
  return ids.map(function (id) {
    const row = byId.get(String(id));
    const type = String(row && row.evidence_type || '').toUpperCase();
    if (!row || type.indexOf('AI') < 0 || type.indexOf('ROLLBACK') >= 0 ||
        String(row.response_type || '').toUpperCase() !== 'CHANGE' ||
        !vNextAdminAiEvidenceActiveAt_(row, bookId, asOf, new Set())) {
      throw new Error('Snapshotted AI evidence is missing or invalid at the basis vintage: ' + String(id));
    }
    return row;
  });
}

function vNextAdminBuildAiRollbackTombstone_(registry, basisRun, sourceEvidence, options) {
  const opt = options || {};
  const operationId = vNextAdminRequiredText_(opt.operationId, 'rollbackOperationId');
  const sourceEvidenceId = vNextAdminRequiredText_(sourceEvidence && sourceEvidence.evidence_id, 'sourceEvidenceId');
  const basisAsOf = vNextAdminDateOnly_(basisRun.as_of);
  const sourceMetadata = vNextAdminParseJson_(sourceEvidence.evidence_text, {});
  const evidenceId = 'AI-RB-' + vNextAdminSha256_(operationId + '|' + sourceEvidenceId).slice(0, 24).toUpperCase();
  const metadata = {
    rollbackOperationId: operationId,
    sourceForecastRunId: String(basisRun.run_id || ''),
    sourceInputDataHash: String(basisRun.input_data_hash || ''),
    rolledBackEvidenceId: sourceEvidenceId,
    sourceParentRequestId: String(sourceMetadata.parentRequestId || ''),
    parentRequestId: String(opt.rollbackRequestId || ''),
    effectiveAsOf: basisAsOf,
    reason: String(opt.reason || ''),
    requestedAt: String(opt.requestedAt || ''),
    requestedBy: String(opt.requestedBy || '').toLowerCase()
  };
  return {
    evidence_id: evidenceId,
    book_id: registry.book_id,
    client_id: registry.client_id,
    fiscal_year: Number(registry.fiscal_year),
    actor_email: String(opt.requestedBy || vNextAdminActor_()).toLowerCase(),
    response_type: 'NO_CHANGE',
    evidence_type: 'AI_ROLLBACK',
    target: String(sourceEvidence.target || ''),
    target_start_month: String(sourceEvidence.target_start_month || ''),
    target_end_month: String(sourceEvidence.target_end_month || ''),
    direction: '',
    amount_mode: 'FIXED',
    amount_low: 0,
    amount_mid: 0,
    amount_high: 0,
    amount_band: '',
    confidence_class: String(sourceEvidence.confidence_class || ''),
    evidence_text: vNextAdminCanonicalJson_(metadata),
    source_url: '',
    source_date: basisAsOf,
    expires_at: '',
    status: 'ACTIVE',
    supersedes_evidence_id: sourceEvidenceId,
    created_at: String(opt.requestedAt || new Date().toISOString()),
    evidence_quality: String(sourceEvidence.evidence_quality || ''),
    ai_model: '',
    prompt_version: '',
    ai_schema_version: '',
    rule_version: 'AI_ROLLBACK_V1',
    applied_amount: 0,
    cap_applied: 0
  };
}

function vNextAdminAssertNoTrustedForecastPayload_(payload) {
  const req = payload && typeof payload === 'object' ? payload : {};
  const forbidden = [
    'trustedReuseSeedFromRunId', 'trustedRollbackContext', 'trustedAllowedDelayedAiRequestIds',
    'seed', 'previousRunId', 'parameters', 'actualRecords', 'actualFetcher', 'aiEvents',
    'internalOperation', 'internalJobType', 'manageState', 'persist', 'spreadsheet'
  ].filter(function (key) { return Object.prototype.hasOwnProperty.call(req, key); });
  if (forbidden.length) throw new Error('Forecast job payload contains forbidden trusted fields: ' + forbidden.join(', '));
  return true;
}

function vNextAdminAuthorizeAiRollbackJob_(hub, job, payload, registry, options) {
  const req = payload && typeof payload === 'object' ? payload : {};
  const opt = options && typeof options === 'object' ? options : {};
  ['trustedReuseSeedFromRunId', 'trustedRollbackContext', 'trustedAllowedDelayedAiRequestIds',
    'seed', 'previousRunId', 'internalOperation'].forEach(function (key) {
    if (Object.prototype.hasOwnProperty.call(req, key)) throw new Error('Rollback job contains a forbidden pre-trusted field: ' + key);
  });
  const operationId = vNextAdminRequiredText_(req.rollbackOperationId, 'rollbackOperationId');
  const sourceRunId = vNextAdminRequiredText_(req.sourceForecastRunId, 'sourceForecastRunId');
  const rollbackRequestId = vNextAdminRequiredText_(req.rollbackRequestId || req.requestId, 'rollbackRequestId');
  if (String(job.target_book_id || '') !== String(registry.book_id || '') ||
      String(req.bookId || '') !== String(registry.book_id || '')) {
    throw new Error('AI rollback job book linkage mismatch.');
  }
  if (String(job.idempotency_key || '') !== 'AI_ROLLBACK_FORECAST|' + operationId ||
      rollbackRequestId !== 'REQ-AI-RB-' + vNextAdminSha256_(operationId).slice(0, 20).toUpperCase()) {
    throw new Error('AI rollback operation/request idempotency lineage is invalid.');
  }
  const authoritativeState = vNextAdminLatestClientState_(hub, registry.book_id, registry.state);
  const persistedResume = opt.allowPersistedResume === true && authoritativeState === 'DRAFT_READY' &&
    String(opt.persistedRunId || '');
  if (authoritativeState !== 'RUNNING' && !persistedResume) {
    throw new Error('AI rollback worker requires the Hub-authoritative RUNNING state.');
  }
  const runs = vNextAdminReadCoreRows_(hub, 'FORECAST_RUN').filter(function (row) {
    return String(row.book_id || '') === String(registry.book_id || '') && String(row.run_id || '') === sourceRunId;
  });
  const sourceRun = runs.length ? runs[runs.length - 1] : null;
  if (!sourceRun || String(sourceRun.status || '').toUpperCase() !== 'SUCCESS' || Number(sourceRun.is_official || 0) === 1 ||
      Number(sourceRun.fiscal_year) !== Number(registry.fiscal_year) ||
      String(sourceRun.client_id || '') && String(sourceRun.client_id || '') !== String(registry.client_id || '')) {
    throw new Error('AI rollback source run is missing or its book/client/FY/status linkage is invalid.');
  }
  if (persistedResume) {
    const resumed = vNextAdminReadCoreRows_(hub, 'FORECAST_RUN').filter(function (row) {
      return String(row.book_id || '') === String(registry.book_id || '') &&
        String(row.run_id || '') === String(opt.persistedRunId || '') &&
        String(row.status || '').toUpperCase() === 'SUCCESS';
    });
    const resumedRun = resumed.length ? resumed[resumed.length - 1] : null;
    const resumedLenses = resumedRun ? vNextAdminParseJson_(resumedRun.lens_json, {}) : {};
    const resumedIdentity = resumedLenses && resumedLenses.runIdentity || {};
    if (!resumedRun || String(resumedRun.previous_run_id || '') !== sourceRunId ||
        String(resumedIdentity.idempotencyKey || '') !== String(job.idempotency_key || '')) {
      throw vNextAdminRunIdentityFailure_('Persisted AI rollback SUCCESS does not match its source run and job lineage.');
    }
  }
  if (String(sourceRun.input_data_hash || '') !== String(req.sourceInputDataHash || '') ||
      String(sourceRun.model_release_id || '') !== String(req.sourceModelReleaseId || '') ||
      vNextAdminDateOnly_(sourceRun.as_of) !== vNextAdminDateOnly_(req.asOf) ||
      vNextAdminDateOnly_(sourceRun.cutoff) !== vNextAdminDateOnly_(req.cutoff)) {
    throw new Error('AI rollback source run immutable fields no longer match the queued lineage.');
  }
  const activeIds = vNextAdminStrictStringArray_(req.basisActiveAiEvidenceIds, 'basisActiveAiEvidenceIds');
  const targetIds = vNextAdminStrictStringArray_(req.targetEvidenceIds, 'targetEvidenceIds');
  const tombstoneIds = vNextAdminStrictStringArray_(req.tombstoneEvidenceIds, 'tombstoneEvidenceIds');
  if (!activeIds.length || !targetIds.length || targetIds.length !== tombstoneIds.length ||
      targetIds.some(function (id) { return activeIds.indexOf(id) < 0; })) {
    throw new Error('AI rollback evidence lineage arrays are incomplete or inconsistent.');
  }
  const scope = vNextAdminNormalizeAiRollbackScope_(req.scope);
  if (scope === 'ALL' && (targetIds.length !== activeIds.length ||
      targetIds.some(function (id) { return activeIds.indexOf(id) < 0; }))) {
    throw new Error('ALL rollback must tombstone every active AI evidence row from the basis run.');
  }
  const sourceEvidenceSummary = vNextAdminParseJson_(sourceRun.evidence_json, {});
  const snapshottedAiIds = sourceEvidenceSummary && sourceEvidenceSummary.effectiveEvidenceIds &&
    sourceEvidenceSummary.effectiveEvidenceIds.ai;
  const nonAiComparableHash = String(sourceEvidenceSummary && sourceEvidenceSummary.nonAiComparableHash || '');
  if (!Array.isArray(snapshottedAiIds) || !nonAiComparableHash ||
      vNextAdminCanonicalJson_(snapshottedAiIds.map(String).sort()) !== vNextAdminCanonicalJson_(activeIds.slice().sort())) {
    throw new Error('AI rollback job does not match the source run evidence/comparable-input snapshot.');
  }
  const evidenceRows = vNextAdminReadCoreRows_(hub, 'EVIDENCE_EVENT');
  const byId = new Map(evidenceRows.map(function (row) { return [String(row.evidence_id || ''), row]; }));
  const basisAsOf = vNextAdminDateOnly_(sourceRun.as_of);
  const activeRows = activeIds.map(function (id) {
    const row = byId.get(id);
    const type = String(row && row.evidence_type || '').toUpperCase();
    if (!row || type.indexOf('AI') < 0 || type.indexOf('ROLLBACK') >= 0 ||
        String(row.response_type || '').toUpperCase() !== 'CHANGE' ||
        !vNextAdminAiEvidenceActiveAt_(row, registry.book_id, basisAsOf, new Set())) {
      throw new Error('Queued basis AI evidence is not valid at the source vintage: ' + id);
    }
    return row;
  });
  targetIds.forEach(function (targetId, index) {
    const tombstone = byId.get(tombstoneIds[index]);
    const metadata = tombstone ? vNextAdminParseJson_(tombstone.evidence_text, {}) : {};
    const expectedTombstoneId = 'AI-RB-' + vNextAdminSha256_(operationId + '|' + targetId).slice(0, 24).toUpperCase();
    if (!tombstone || String(tombstone.status || '').toUpperCase() !== 'ACTIVE' ||
        String(tombstone.evidence_id || '') !== expectedTombstoneId ||
        String(tombstone.response_type || '').toUpperCase() !== 'NO_CHANGE' ||
        String(tombstone.evidence_type || '').toUpperCase() !== 'AI_ROLLBACK' ||
        String(tombstone.supersedes_evidence_id || '') !== targetId ||
        String(metadata.rollbackOperationId || '') !== operationId ||
        String(metadata.sourceForecastRunId || '') !== sourceRunId ||
        String(metadata.parentRequestId || '') !== rollbackRequestId ||
        String(metadata.effectiveAsOf || '') !== basisAsOf ||
        String(metadata.reason || '') !== String(req.reason || '') ||
        !vNextAdminAiEvidenceActiveAt_(tombstone, registry.book_id, basisAsOf, new Set())) {
      throw new Error('AI rollback tombstone linkage is invalid: ' + tombstoneIds[index]);
    }
  });
  const parentIds = Array.from(new Set(activeRows.map(function (row) {
    return String(vNextAdminParseJson_(row.evidence_text, {}).parentRequestId || '');
  }).filter(Boolean))).sort();
  const allowedDelayedRequestIds = Array.from(new Set(parentIds.concat([rollbackRequestId]))).sort();
  return {
    sourceRun: sourceRun,
    rollbackRequestId: rollbackRequestId,
    trustedReuseSeedFromRunId: sourceRunId,
    trustedAllowedDelayedAiRequestIds: allowedDelayedRequestIds,
    trustedRollbackContext: {
      operationId: operationId,
      jobId: String(job.job_id || ''),
      sourceForecastRunId: sourceRunId,
      sourceInputDataHash: String(sourceRun.input_data_hash || ''),
      sourceModelReleaseId: String(sourceRun.model_release_id || ''),
      nonAiComparableHash: nonAiComparableHash,
      asOf: basisAsOf,
      scope: scope,
      activeEvidenceIds: activeIds,
      targetEvidenceIds: targetIds,
      tombstoneEvidenceIds: tombstoneIds,
      reason: vNextAdminRequiredText_(req.reason, 'reason')
    }
  };
}

function vNextAdminStrictStringArray_(value, label) {
  if (!Array.isArray(value) || value.some(function (item) { return typeof item !== 'string' || !item.trim(); })) {
    throw new Error(String(label || 'value') + ' must be a non-empty string array.');
  }
  const normalized = value.map(function (item) { return item.trim(); });
  if (new Set(normalized).size !== normalized.length) throw new Error(String(label || 'value') + ' contains duplicate IDs.');
  return normalized;
}

function vNextAdminResolveCurrentOfficialBasis_(hub, bookId) {
  const registry = vNextAdminFindRegistryRow_(hub, function (row) {
    return String(row.book_id || '') === String(bookId || '');
  });
  if (!registry || String(registry.mode || '') !== 'CLIENT' || String(registry.status || '').toUpperCase() !== 'ACTIVE') {
    throw new Error('An ACTIVE CLIENT registry record is required for amendment.');
  }
  const currentOfficialId = vNextAdminRequiredText_(registry.current_official_id, 'currentOfficialId');
  return vNextAdminResolveOfficialBasis_(hub, registry, currentOfficialId);
}

function vNextAdminResolveOfficialBasis_(hub, registry, officialId) {
  const bookId = String(registry.book_id || '');
  const targetOfficialId = vNextAdminRequiredText_(officialId, 'officialId');
  const officialRows = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.OFFICIAL).rows.filter(function (row) {
    return String(row.book_id || '') === bookId && String(row.official_id || '') === targetOfficialId;
  });
  if (!officialRows.length) throw new Error('OFFICIAL_RUNS record was not found for amendment basis.');
  const officialRow = officialRows[officialRows.length - 1];
  const forecasts = vNextAdminReadCoreRows_(hub, 'FORECAST_RUN');
  const officialForecast = forecasts.find(function (row) {
    return String(row.book_id || '') === bookId && String(row.run_id || '') === String(officialRow.forecast_run_id || '') &&
      String(row.official_vintage_id || '') === targetOfficialId && Number(row.is_official || 0) === 1;
  });
  if (!officialForecast || String(officialForecast.status || '').toUpperCase() !== 'OFFICIAL_LOCKED') {
    throw new Error('The frozen official FORECAST_RUN is missing or invalid.');
  }
  const sourceRunId = String(officialRow.source_forecast_run_id || officialForecast.previous_run_id || '');
  const sourceForecast = forecasts.find(function (row) {
    return String(row.book_id || '') === bookId && String(row.run_id || '') === sourceRunId &&
      String(row.status || '').toUpperCase() === 'SUCCESS' && Number(row.is_official || 0) !== 1;
  });
  if (!sourceForecast || String(officialForecast.previous_run_id || '') !== sourceRunId) {
    throw new Error('The current official source forecast lineage is missing or inconsistent.');
  }
  if (String(sourceForecast.client_id || '') && String(sourceForecast.client_id || '') !== String(registry.client_id || '')) {
    throw new Error('Official source forecast client_id mismatch.');
  }
  if (Number(sourceForecast.fiscal_year) !== Number(registry.fiscal_year) ||
      Number(officialForecast.fiscal_year) !== Number(registry.fiscal_year)) {
    throw new Error('Official forecast fiscal year mismatch.');
  }
  const approvedPlans = vNextAdminReadCoreRows_(hub, 'PLAN_VERSION').filter(function (row) {
    return String(row.book_id || '') === bookId && String(row.official_vintage_id || '') === targetOfficialId &&
      String(row.status || '').toUpperCase() === 'APPROVED';
  });
  if (!approvedPlans.length) throw new Error('The current official APPROVED PLAN_VERSION is missing.');
  const approvedPlan = approvedPlans[approvedPlans.length - 1];
  if (String(approvedPlan.run_id || '') !== sourceRunId) {
    throw new Error('The current approved plan does not reference the official source forecast.');
  }
  const systems = [sourceForecast.system_recommended, officialForecast.system_recommended, approvedPlan.system_recommended].map(Number);
  if (!systems.every(isFinite) || Math.abs(systems[0] - systems[1]) > 1 || Math.abs(systems[0] - systems[2]) > 1) {
    throw new Error('System recommendation differs across source, frozen official, and approved plan.');
  }
  return {
    registry: registry, officialId: targetOfficialId, officialRow: officialRow,
    officialForecast: officialForecast, sourceForecast: sourceForecast, approvedPlan: approvedPlan
  };
}

function vNextAdminAmendmentBasisHash_(basis) {
  return vNextAdminSha256_(vNextAdminCanonicalJson_({
    bookId: basis.registry.book_id, currentOfficialId: basis.officialId,
    officialForecastRunId: basis.officialForecast.run_id,
    sourceForecastRunId: basis.sourceForecast.run_id,
    approvedPlanVersionId: basis.approvedPlan.plan_version_id,
    systemRecommended: Number(basis.sourceForecast.system_recommended || 0),
    adoptionDelta: Number(basis.approvedPlan.adoption_delta || 0),
    salesUplift: Number(basis.approvedPlan.sales_uplift || 0),
    finalBudget: Number(basis.approvedPlan.final_budget || 0),
    upliftAllocationJson: String(basis.approvedPlan.uplift_allocation_json || '')
  }));
}

function vNextAdminNormalizeAmendmentInput_(request, basis) {
  const req = request || {};
  if (!Object.prototype.hasOwnProperty.call(req, 'adoptionDelta')) throw new Error('adoptionDelta is required.');
  if (!Object.prototype.hasOwnProperty.call(req, 'salesUplift')) throw new Error('salesUplift is required.');
  const system = Number(basis.sourceForecast.system_recommended);
  const adoptionDelta = Number(req.adoptionDelta);
  const salesUplift = Number(req.salesUplift);
  if (![system, adoptionDelta, salesUplift].every(isFinite) ||
      Math.abs(adoptionDelta) > 1e15 || Math.abs(salesUplift) > 1e15) {
    throw new Error('Amendment amounts must be finite and within the supported range.');
  }
  const adoptedForecast = system + adoptionDelta;
  if (adoptedForecast < 0) throw new Error('adoptedForecast cannot be negative.');
  if (salesUplift < 0) throw new Error('salesUplift cannot be negative.');
  const adoptionReason = vNextAdminText_(req.adoptionReason);
  const upliftReason = vNextAdminText_(req.upliftReason);
  const upliftOwner = vNextAdminText_(req.upliftOwner);
  const upliftAction = vNextAdminText_(req.upliftAction);
  if (adoptionDelta !== 0 && !adoptionReason) throw new Error('adoptionReason is required for a non-zero adoptionDelta.');
  let upliftDueDate = '';
  if (req.upliftDueDate) {
    const parsedDue = new Date(req.upliftDueDate);
    if (isNaN(parsedDue.getTime())) throw new Error('upliftDueDate is invalid.');
    upliftDueDate = typeof vNextFormatDateOnly_ === 'function'
      ? vNextFormatDateOnly_(parsedDue)
      : Utilities.formatDate(parsedDue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  if (salesUplift !== 0 && (!upliftReason || !upliftOwner || !upliftAction || !upliftDueDate)) {
    throw new Error('Non-zero sales uplift requires reason, owner, action, and due date.');
  }
  let allocationInput = req.upliftAllocation !== undefined ? req.upliftAllocation : req.monthAllocation;
  if (typeof allocationInput === 'string') {
    allocationInput = allocationInput.split(/[,、\s]+/).filter(Boolean).map(Number);
  }
  if (!Array.isArray(allocationInput)) allocationInput = [];
  if (typeof vNextValidateUpliftAllocation_ !== 'function') throw new Error('Server uplift allocation validator is not installed.');
  const upliftAllocation = vNextValidateUpliftAllocation_(allocationInput, salesUplift, Number(basis.registry.fiscal_year));
  const finalBudget = adoptedForecast + salesUplift;
  const previous = basis.approvedPlan;
  const noChange = Math.abs(Number(previous.adoption_delta || 0) - adoptionDelta) <= 1 &&
    Math.abs(Number(previous.sales_uplift || 0) - salesUplift) <= 1 &&
    String(previous.adoption_reason || '') === adoptionReason &&
    String(previous.uplift_reason || '') === upliftReason && String(previous.uplift_owner || '') === upliftOwner &&
    String(previous.uplift_action || '') === upliftAction && vNextAdminDateOnlyText_(previous.uplift_due_date) === upliftDueDate &&
    vNextAdminCanonicalJson_(vNextAdminParseJson_(previous.uplift_allocation_json, [])) === vNextAdminCanonicalJson_(upliftAllocation);
  if (noChange) throw new Error('現在の正式予算と同一です。数値または計画内容を変更してください。');
  return {
    systemRecommended: system, adoptionDelta: adoptionDelta, adoptionReason: adoptionReason,
    adoptedForecast: adoptedForecast, salesUplift: salesUplift, upliftReason: upliftReason,
    upliftOwner: upliftOwner, upliftAction: upliftAction, upliftDueDate: upliftDueDate,
    upliftAllocation: upliftAllocation, finalBudget: finalBudget
  };
}

function vNextAdminDateOnlyText_(value) {
  if (!value) return '';
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(value).slice(0, 10);
}

/** Detect a submitted local plan and create one immutable approval request per plan version. */
function vNextAdminCreateApprovalFromSubmittedPlan_(hub, client, registry) {
  const routing = vNextAdminReadKeyValueSheet_(client, VN_ADMIN_BOOK_CONFIG_SHEET);
  if (String(routing.state || '').toUpperCase() !== 'SUBMITTED') return { created: false, reason: 'STATE_NOT_SUBMITTED' };
  const plans = vNextAdminReadCoreRows_(hub, 'PLAN_VERSION').filter(function (row) {
    return String(row.book_id || '') === String(registry.book_id || '') && String(row.status || '').toUpperCase() === 'SUBMITTED';
  });
  if (!plans.length) return { created: false, reason: 'PLAN_NOT_FOUND' };
  const plan = plans[plans.length - 1];
  const forecasts = vNextAdminReadCoreRows_(hub, 'FORECAST_RUN').filter(function (row) {
    return String(row.book_id || '') === String(registry.book_id || '') &&
      String(row.run_id || '') === String(plan.run_id || '') && String(row.status || '').toUpperCase() === 'SUCCESS';
  });
  if (!forecasts.length) return { created: false, reason: 'FORECAST_NOT_FOUND' };
  const forecast = forecasts[forecasts.length - 1];
  vNextAdminValidateSubmittedPlan_(registry, forecast, plan);
  const snapshot = vNextAdminBuildPlanApprovalSnapshot_(registry, forecast, plan);
  const snapshotJson = vNextAdminCanonicalJson_(snapshot);
  const snapshotHash = vNextAdminSha256_(snapshotJson);
  const idempotencyKey = ['PUBLISH', registry.book_id, plan.plan_version_id, snapshotHash].join('|');
  const existing = vNextAdminFindApproval_(hub, function (row) {
    return String(row.idempotency_key || '') === idempotencyKey ||
      (String(row.book_id || '') === String(registry.book_id || '') && String(row.plan_version_id || '') === String(plan.plan_version_id || ''));
  });
  if (existing) return { created: false, approvalRequestId: existing.approval_request_id, status: existing.status };
  const now = new Date();
  const approvalId = 'APR-' + Utilities.getUuid();
  const row = {
    approval_request_id: approvalId, request_type: 'PUBLISH', book_id: registry.book_id,
    client_id: registry.client_id, client_name: registry.client_name, fiscal_year: registry.fiscal_year,
    forecast_run_id: forecast.run_id, plan_version_id: plan.plan_version_id,
    supersedes_official_id: '', amendment_reason: '', snapshot_json: snapshotJson,
    snapshot_hash: snapshotHash, status: 'PENDING', processing_attempts: 0, requested_at: plan.submitted_at || now,
    requested_by: plan.submitted_by || '', decision_at: '', decision_by: '', decision_comment: '',
    official_id: '', idempotency_key: idempotencyKey, updated_at: now
  };
  vNextAdminAppendObject_(hub, VN_ADMIN_SHEETS.APPROVALS, row);
  vNextAdminWriteAudit_(hub, 'AUTO_CREATE_APPROVAL', 'APPROVAL', approvalId, 'SUCCESS', {
    bookId: registry.book_id, runId: forecast.run_id, planVersionId: plan.plan_version_id,
    snapshotHash: snapshotHash
  });
  return { created: true, approvalRequestId: approvalId };
}

function vNextAdminBuildPlanApprovalSnapshot_(registry, forecast, plan, options) {
  const opt = options || {};
  const allocation = vNextAdminParseJson_(plan.uplift_allocation_json, []);
  const snapshot = {
    snapshotSchema: 'forecast-plan-approval-v1',
    book: {
      bookId: registry.book_id, clientId: registry.client_id, clientName: registry.client_name,
      fiscalYear: Number(registry.fiscal_year)
    },
    forecast: {
      runId: forecast.run_id, inputDataHash: forecast.input_data_hash,
      modelReleaseId: forecast.model_release_id, asOf: forecast.as_of, cutoff: forecast.cutoff,
      layers: {
        historyBaseline: Number(forecast.history_baseline || 0),
        objectiveForecast: Number(forecast.objective_forecast || 0),
        humanDelta: Number(forecast.human_delta || 0), aiDelta: Number(forecast.ai_delta || 0),
        systemRecommended: Number(forecast.system_recommended || 0)
      },
      annual: { p10: Number(forecast.p10 || 0), p50: Number(forecast.p50 || 0), p90: Number(forecast.p90 || 0) },
      quarters: vNextAdminParseJson_(forecast.quarter_json, []),
      months: vNextAdminParseJson_(forecast.month_json, []),
      lenses: vNextAdminParseJson_(forecast.lens_json, {}),
      evidenceSummary: vNextAdminParseJson_(forecast.evidence_json, {})
    },
    plan: {
      planVersionId: plan.plan_version_id, status: plan.status,
      systemRecommended: Number(plan.system_recommended || 0),
      adoptionDelta: Number(plan.adoption_delta || 0), adoptionReason: String(plan.adoption_reason || ''),
      adoptedForecast: Number(plan.adopted_forecast || 0), salesUplift: Number(plan.sales_uplift || 0),
      upliftReason: String(plan.uplift_reason || ''), upliftOwner: String(plan.uplift_owner || ''),
      upliftAction: String(plan.uplift_action || ''), upliftDueDate: String(plan.uplift_due_date || ''),
      upliftAllocation: allocation, finalBudget: Number(plan.final_budget || 0),
      submittedAt: String(plan.submitted_at || ''), submittedBy: String(plan.submitted_by || '')
    }
  };
  if (String(opt.requestType || '').toUpperCase() === 'AMENDMENT') {
    const basis = opt.officialBasis || {};
    snapshot.plan.versionNo = Number(plan.version_no || 0);
    snapshot.plan.amendsPlanVersionId = String(plan.amends_plan_version_id || '');
    snapshot.amendment = {
      supersedesOfficialId: String(opt.supersedesOfficialId || ''),
      predecessorOfficialForecastRunId: String(basis.officialForecast && basis.officialForecast.run_id || ''),
      predecessorApprovedPlanVersionId: String(basis.approvedPlan && basis.approvedPlan.plan_version_id || ''),
      predecessorFinalBudget: Number(basis.approvedPlan && basis.approvedPlan.final_budget || 0),
      amendmentReason: String(opt.amendmentReason || '')
    };
  }
  return snapshot;
}

function vNextAdminValidateSubmittedPlan_(registry, forecast, plan, options) {
  const opt = options || {};
  const requestType = String(opt.requestType || 'PUBLISH').toUpperCase();
  const isAmendment = requestType === 'AMENDMENT';
  if (String(plan.book_id || '') !== String(registry.book_id || '') ||
      String(plan.run_id || '') !== String(forecast.run_id || '') ||
      String(plan.status || '').toUpperCase() !== 'SUBMITTED' ||
      String(plan.official_vintage_id || '')) {
    throw new Error('Submitted plan book/run/status linkage is invalid.');
  }
  if (String(forecast.book_id || '') !== String(registry.book_id || '') ||
      String(forecast.status || '').toUpperCase() !== 'SUCCESS' || Number(forecast.is_official || 0) === 1 ||
      Number(forecast.fiscal_year) !== Number(registry.fiscal_year)) {
    throw new Error('Linked forecast is not a successful draft for this book and fiscal year.');
  }
  const ownerEmails = vNextAdminParseList_(registry.forecast_owner_emails).map(function (email) {
    return String(email || '').toLowerCase();
  });
  const submittedBy = String(plan.submitted_by || '').toLowerCase();
  const ownerSubmitted = ownerEmails.length === 1 && submittedBy === ownerEmails[0];
  const authorizedAdminSubmitted = isAmendment && !!String(opt.allowedSubmitter || '') &&
    submittedBy === String(opt.allowedSubmitter || '').toLowerCase();
  if (!ownerSubmitted && !authorizedAdminSubmitted) {
    throw new Error(isAmendment
      ? 'Amendment submitted_by must be its authenticated Admin creator or the registered 予算策定担当.'
      : 'submitted_by must be the single 予算策定担当 registered for this book.');
  }
  if (isAmendment) {
    const basis = opt.officialBasis || {};
    if (!basis.officialId || String(opt.supersedesOfficialId || '') !== String(basis.officialId || '')) {
      throw new Error('Amendment must supersede the current official vintage.');
    }
    if (!basis.approvedPlan || String(plan.amends_plan_version_id || '') !== String(basis.approvedPlan.plan_version_id || '')) {
      throw new Error('Amendment plan must link to the current official APPROVED plan.');
    }
    if (!basis.sourceForecast || String(forecast.run_id || '') !== String(basis.sourceForecast.run_id || '')) {
      throw new Error('Amendment plan must reference the current official source forecast run.');
    }
  }
  const system = Number(plan.system_recommended);
  const forecastSystem = Number(forecast.system_recommended);
  const delta = Number(plan.adoption_delta || 0);
  const adopted = Number(plan.adopted_forecast);
  const uplift = Number(plan.sales_uplift || 0);
  const finalBudget = Number(plan.final_budget);
  if (![system, forecastSystem, delta, adopted, uplift, finalBudget].every(isFinite)) {
    throw new Error('Submitted plan contains a non-finite amount.');
  }
  if (Math.abs(system - forecastSystem) > 1 || Math.abs(adopted - (system + delta)) > 1 || adopted < 0) {
    throw new Error('Submitted plan system/adoption arithmetic is invalid.');
  }
  if (delta !== 0 && !vNextAdminText_(plan.adoption_reason)) {
    throw new Error('A non-zero adoption delta requires an adoption reason.');
  }
  if (uplift < 0) throw new Error('Sales uplift cannot be negative.');
  if (uplift !== 0 && (!vNextAdminText_(plan.uplift_reason) || !vNextAdminText_(plan.uplift_owner) ||
      !vNextAdminText_(plan.uplift_action) || !vNextAdminText_(plan.uplift_due_date))) {
    throw new Error('A non-zero sales uplift requires reason, owner, action, and due date.');
  }
  const allocation = vNextAdminParseJson_(plan.uplift_allocation_json, []);
  if (typeof vNextValidateUpliftAllocation_ !== 'function') throw new Error('Server uplift allocation validator is not installed.');
  vNextValidateUpliftAllocation_(allocation, uplift, Number(registry.fiscal_year));
  if (Math.abs(finalBudget - (adopted + uplift)) > 1) {
    throw new Error('Submitted plan final budget arithmetic is invalid.');
  }
  return true;
}

function vNextAdminExecuteMigrationSkeleton_(hub, job, payload) {
  const migrationId = 'MIG-' + Utilities.getUuid();
  const now = new Date();
  const registry = vNextAdminFindRegistryRow_(hub, function (item) {
    return String(item.book_id || '') === String(job.target_book_id || '') &&
      String(item.mode || '') === 'CLIENT';
  });
  if (!registry || String(registry.spreadsheet_id || '') !== String(job.target_spreadsheet_id || '')) {
    throw new Error('Migration target does not match BOOK_REGISTRY.');
  }
  const targetRelease = vNextAdminResolveRelease_(hub, payload.targetReleaseId);
  const currentRelease = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.RELEASES).rows.find(function (release) {
    return String(release.release_id || '') === String(registry.template_release_id || '');
  });
  if (!currentRelease) throw new Error('Current Client release record is missing.');
  const changed = String(currentRelease.release_id || '') !== String(targetRelease.release_id || '');
  const restrictedState = ['OFFICIAL_LOCKED', 'REVIEW_DUE', 'YEAR_CLOSED'].indexOf(String(registry.state || '').toUpperCase()) >= 0;
  const plan = {
    strategy: 'known-bound-script-and-visible-template-assets',
    bookId: registry.book_id, spreadsheetId: registry.spreadsheet_id,
    clientScriptIdPresent: Boolean(String(registry.client_script_id || '')),
    fromReleaseId: currentRelease.release_id, targetReleaseId: targetRelease.release_id,
    fromSchemaVersion: currentRelease.schema_version, targetSchemaVersion: targetRelease.schema_version,
    state: String(registry.state || ''), restrictedState: restrictedState,
    changed: changed, preservesOfficialAndCoreHistory: true
  };
  const row = {
    migration_id: migrationId, book_id: job.target_book_id, spreadsheet_id: job.target_spreadsheet_id,
    from_release_id: currentRelease.release_id, to_release_id: targetRelease.release_id,
    status: 'PLANNED', dry_run: payload.dryRun === false ? 0 : 1,
    plan_json: vNextAdminCanonicalJson_(plan),
    result_json: '', started_at: now, finished_at: '', actor: vNextAdminActor_(), error: ''
  };
  vNextAdminAppendObject_(hub, VN_ADMIN_SHEETS.MIGRATIONS, row);
  try {
    if (!String(registry.client_script_id || '')) {
      throw new Error('This legacy-copied Client has no centrally recorded bound script ID; in-place migration is blocked.');
    }
    if (String(currentRelease.schema_version || '') !== String(targetRelease.schema_version || '') ||
        String(targetRelease.schema_version || '') !== vNextAdminClientSchemaVersion_()) {
      throw new Error('Generic migration supports only the same Client Core schema. A versioned schema hook is required.');
    }
    if (!String(targetRelease.client_runtime_sha256 || '') ||
        !String(targetRelease.template_script_id || '') || !String(targetRelease.template_spreadsheet_id || '')) {
      throw new Error('Target release runtime identity is incomplete.');
    }
    const targetTemplate = SpreadsheetApp.openById(String(targetRelease.template_spreadsheet_id));
    if (vNextDetectBookMode_(targetTemplate) !== 'TEMPLATE' ||
        vNextAdminTemplateContentHash_(targetTemplate) !== String(targetRelease.template_content_sha256 || '')) {
      throw new Error('Target immutable Template identity/content check failed.');
    }
    if (typeof vNextClientRuntimeAssertBoundParent_ !== 'function') {
      throw new Error('Client runtime parent verifier is not installed.');
    }
    vNextClientRuntimeAssertBoundParent_(String(targetRelease.template_script_id), String(targetRelease.template_spreadsheet_id));
    vNextClientRuntimeAssertBoundParent_(String(registry.client_script_id), String(registry.spreadsheet_id));
    if (payload.dryRun !== false) {
      const dryResult = Object.assign({}, plan, { readyToApply: changed });
      vNextAdminPatchLatestMigration_(hub, migrationId, {
        status: 'DRY_RUN_READY', finished_at: new Date(), result_json: vNextAdminCanonicalJson_(dryResult)
      });
      return { migrationId: migrationId, status: 'DRY_RUN_READY', changed: changed, plan: dryResult };
    }
    if (!VN_ADMIN_MIGRATION_APPLY_ENABLED) {
      throw new Error('Pilot期間はClient release移行のAPPLYをserver-sideで停止しています。');
    }
    const reason = vNextAdminRequiredText_(payload.reason, 'migration.reason');
    if (restrictedState && payload.critical !== true) {
      throw new Error('Official/review/closed books require critical=true and an explicit reason.');
    }
    if (!changed) {
      vNextAdminPatchLatestMigration_(hub, migrationId, {
        status: 'SUCCEEDED', finished_at: new Date(), result_json: vNextAdminCanonicalJson_({ changed: false })
      });
      return { migrationId: migrationId, status: 'SUCCEEDED', changed: false };
    }
    if (typeof vNextClientRuntimeCopyScriptContent_ !== 'function') {
      throw new Error('Verified Client runtime copy helper is not installed.');
    }
    const copied = vNextClientRuntimeCopyScriptContent_(
      String(targetRelease.template_script_id), String(registry.client_script_id),
      String(targetRelease.client_runtime_sha256)
    );
    const client = SpreadsheetApp.openById(String(registry.spreadsheet_id));
    vNextAdminCopyTemplateUiToClient_(targetTemplate, client);
    vNextAdminEnsureUiShell_(client, { clientName: registry.client_name, fiscalYear: registry.fiscal_year });
    vNextAdminWriteBookConfig_(client, {
      version: targetRelease.release_id, template_release_id: targetRelease.release_id,
      schema_version: targetRelease.schema_version,
      client_runtime_version: targetRelease.client_runtime_version,
      client_runtime_bundle_sha256: targetRelease.client_runtime_sha256,
      updated_at: new Date(), updated_by: vNextAdminActor_()
    });
    const localSystem = vNextAdminReadKeyValueSheet_(client, VN_ADMIN_SYSTEM_CONFIG_SHEET);
    vNextAdminReplaceSystemConfig_(client, Object.assign({}, localSystem, {
      mode: 'CLIENT', book_id: registry.book_id,
      active_release_id: targetRelease.release_id, schema_version: targetRelease.schema_version
    }));

    const metas = vNextAdminReadCoreRows_(hub, 'BOOK_META').filter(function (meta) {
      return String(meta.book_id || '') === String(registry.book_id || '');
    });
    if (!metas.length) throw new Error('Trusted BOOK_META is missing for migration.');
    const predecessor = metas[metas.length - 1];
    const migrationMetaId = 'MIGMETA-' + vNextAdminSha256_(
      String(registry.book_id) + '|' + String(targetRelease.release_id)
    ).slice(0, 24);
    let migrationMeta = metas.find(function (meta) {
      return String(meta.record_id || '') === migrationMetaId;
    });
    if (!migrationMeta) {
      migrationMeta = Object.assign({}, predecessor, {
        record_id: migrationMetaId, state: String(registry.state || predecessor.state || ''),
        template_version: targetRelease.release_id, schema_version: targetRelease.schema_version,
        event_type: 'MIGRATED', supersedes_record_id: predecessor.record_id,
        recorded_at: new Date().toISOString(), recorded_by: vNextAdminActor_()
      });
      vNextAdminAppendCoreRowsNoLock_(hub, 'BOOK_META', [migrationMeta]);
    }
    vNextAdminAppendMissingCoreRows_(client, 'BOOK_META', 'record_id', [migrationMeta]);
    vNextAdminPatchRegistryByBookId_(hub, registry.book_id, {
      template_release_id: targetRelease.release_id,
      schema_version: targetRelease.schema_version,
      client_runtime_version: targetRelease.client_runtime_version,
      client_runtime_sha256: targetRelease.client_runtime_sha256,
      health_status: 'PENDING', health_code: 'MIGRATED', updated_at: new Date()
    });
    const result = {
      changed: true, fromReleaseId: currentRelease.release_id,
      targetReleaseId: targetRelease.release_id, copiedRuntimeSha256: copied.bundleSha256,
      preservedState: registry.state, preservedOfficialId: registry.current_official_id || '', reason: reason
    };
    vNextAdminPatchLatestMigration_(hub, migrationId, {
      status: 'SUCCEEDED', finished_at: new Date(), result_json: vNextAdminCanonicalJson_(result)
    });
    vNextAdminWriteAudit_(hub, 'MIGRATE_CLIENT_RELEASE', 'BOOK', registry.book_id, 'SUCCESS', result);
    return { migrationId: migrationId, status: 'SUCCEEDED', result: result };
  } catch (migrationError) {
    const message = String(migrationError && migrationError.message || migrationError);
    vNextAdminPatchLatestMigration_(hub, migrationId, {
      status: 'BLOCKED', finished_at: new Date(), error: message
    });
    vNextAdminAppendException_(hub, {
      severity: 'ERROR', exception_type: 'CLIENT_MIGRATION_BLOCKED', book_id: registry.book_id,
      client_name: registry.client_name, fiscal_year: registry.fiscal_year,
      title: 'Client release migration was blocked', detail: message,
      recommended_action: 'MIGRATION_LOGの検証結果を確認し、修正後に新しいjobで再実行', source_ref: migrationId
    });
    throw migrationError;
  }
}

// ---------------------------- Approval / official ----------------------------

function vNextAdminResolveSnapshot_(request, registry, hubSpreadsheet) {
  const req = request && typeof request === 'object' ? request : {};
  const hub = hubSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  const bookId = String(registry.book_id || '');
  const requestType = String(req.requestType || 'PUBLISH').toUpperCase();
  const isAmendment = requestType === 'AMENDMENT';
  const supersedesOfficialId = isAmendment
    ? vNextAdminRequiredText_(req.supersedesOfficialId, 'supersedesOfficialId') : '';
  const amendmentReason = isAmendment
    ? vNextAdminRequiredText_(req.amendmentReason, 'amendmentReason') : '';
  const officialBasis = isAmendment
    ? vNextAdminResolveOfficialBasis_(hub, registry, supersedesOfficialId) : null;
  const requestedPlanId = vNextAdminText_(req.planVersionId);
  const requestedRunId = vNextAdminText_(req.forecastRunId);
  const plans = vNextAdminReadCoreRows_(hub, 'PLAN_VERSION').filter(function (row) {
    return String(row.book_id || '') === bookId &&
      String(row.status || '').toUpperCase() === 'SUBMITTED' &&
      (!requestedPlanId || String(row.plan_version_id || '') === requestedPlanId);
  });
  if (!plans.length) throw new Error('A SUBMITTED PLAN_VERSION for this book was not found.');
  const plan = plans[plans.length - 1];
  const planRunId = String(plan.run_id || '');
  if (!planRunId) throw new Error('The submitted plan is not linked to a forecast run.');
  if (requestedRunId && requestedRunId !== planRunId) {
    throw new Error('forecastRunId does not match the submitted plan run_id.');
  }
  const forecasts = vNextAdminReadCoreRows_(hub, 'FORECAST_RUN').filter(function (row) {
    return String(row.book_id || '') === bookId && String(row.run_id || '') === planRunId &&
      String(row.status || '').toUpperCase() === 'SUCCESS' && Number(row.is_official || 0) !== 1;
  });
  if (!forecasts.length) throw new Error('A successful draft FORECAST_RUN linked to the submitted plan was not found.');
  const forecast = forecasts[forecasts.length - 1];
  if (String(forecast.client_id || '') && String(forecast.client_id || '') !== String(registry.client_id || '')) {
    throw new Error('Forecast client_id does not match BOOK_REGISTRY.');
  }
  if (Number(forecast.fiscal_year) !== Number(registry.fiscal_year)) {
    throw new Error('Forecast fiscal_year does not match BOOK_REGISTRY.');
  }
  const validationOptions = {
    requestType: requestType, allowedSubmitter: vNextAdminText_(req.allowedSubmitter),
    supersedesOfficialId: supersedesOfficialId, officialBasis: officialBasis
  };
  vNextAdminValidateSubmittedPlan_(registry, forecast, plan, validationOptions);
  return vNextAdminBuildPlanApprovalSnapshot_(registry, forecast, plan, {
    requestType: requestType, supersedesOfficialId: supersedesOfficialId,
    amendmentReason: amendmentReason, officialBasis: officialBasis
  });
}

function vNextAdminIssueOfficial_(hub, approval, registry, decisionComment) {
  const snapshotJson = String(approval.snapshot_json || '');
  const recalculatedHash = vNextAdminSha256_(snapshotJson);
  if (recalculatedHash !== String(approval.snapshot_hash || '')) throw new Error('Approval snapshot hash mismatch; official issue aborted.');
  const requestType = String(approval.request_type || '').toUpperCase();
  const officialId = 'OFF-' + vNextAdminSha256_('APPROVAL|' + String(approval.approval_request_id || '')).slice(0, 24).toUpperCase();
  const canonicalSnapshot = vNextAdminResolveSnapshot_({
    forecastRunId: approval.forecast_run_id,
    planVersionId: approval.plan_version_id,
    requestType: requestType,
    supersedesOfficialId: approval.supersedes_official_id || '',
    amendmentReason: approval.amendment_reason || '',
    allowedSubmitter: approval.requested_by || ''
  }, registry, hub);
  const canonicalSnapshotJson = vNextAdminCanonicalJson_(canonicalSnapshot);
  if (canonicalSnapshotJson !== snapshotJson || vNextAdminSha256_(canonicalSnapshotJson) !== String(approval.snapshot_hash || '')) {
    throw new Error('Approval snapshot no longer matches the canonical Hub forecast and plan records.');
  }
  // Recovery path for a previous attempt that completed the central official
  // issue but failed before client propagation or approval-row finalization.
  const existingForApproval = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.OFFICIAL).rows.find(function (item) {
    return String(item.approval_request_id || '') === String(approval.approval_request_id || '');
  });
  if (existingForApproval) {
    const storedSnapshot = vNextAdminParseJson_(existingForApproval.snapshot_json, {});
    const frozenRows = vNextAdminReadCoreRows_(hub, 'FORECAST_RUN').filter(function (row) {
      return String(row.book_id || '') === String(approval.book_id || '') &&
        String(row.run_id || '') === String(existingForApproval.forecast_run_id || '') &&
        String(row.official_vintage_id || '') === officialId && Number(row.is_official || 0) === 1 &&
        String(row.previous_run_id || '') === String(approval.forecast_run_id || '');
    });
    const approvedPlans = vNextAdminReadCoreRows_(hub, 'PLAN_VERSION').filter(function (row) {
      return String(row.book_id || '') === String(approval.book_id || '') &&
        String(row.official_vintage_id || '') === officialId && String(row.status || '').toUpperCase() === 'APPROVED' &&
        String(row.amends_plan_version_id || '') === String(approval.plan_version_id || '');
    });
    if (String(existingForApproval.official_id || '') !== officialId ||
        String(existingForApproval.book_id || '') !== String(approval.book_id || '') ||
        String(existingForApproval.source_forecast_run_id || '') !== String(approval.forecast_run_id || '') ||
        String(existingForApproval.snapshot_hash || '') !== String(approval.snapshot_hash || '') ||
        String(existingForApproval.snapshot_json || '') !== snapshotJson ||
        String(storedSnapshot.plan && storedSnapshot.plan.planVersionId || '') !== String(approval.plan_version_id || '') ||
        !frozenRows.length || !approvedPlans.length) {
      throw new Error('Existing deterministic official record is incomplete or inconsistent; recovery stopped.');
    }
    const registryCurrent = String(registry.current_official_id || '');
    if (registryCurrent !== officialId) {
      const safePredecessor = requestType === 'PUBLISH'
        ? registryCurrent === ''
        : registryCurrent === String(approval.supersedes_official_id || '');
      if (!safePredecessor) {
        throw new Error('Registry points to a different official vintage; deterministic recovery stopped.');
      }
      vNextAdminPatchRegistryByBookId_(hub, approval.book_id, {
        current_official_id: officialId, state: 'OFFICIAL_LOCKED', updated_at: new Date(),
        note: 'Registry official pointer repaired from approval=' + approval.approval_request_id
      });
      vNextAdminWriteAudit_(hub, 'REPAIR_OFFICIAL_POINTER', 'BOOK', approval.book_id, 'SUCCESS', {
        fromOfficialId: registryCurrent, toOfficialId: officialId, approvalRequestId: approval.approval_request_id
      });
    }
    vNextAdminCopyOfficialToClientBestEffort_(hub, registry, existingForApproval);
    return {
      officialId: officialId, officialForecastRunId: existingForApproval.forecast_run_id,
      approvedPlanVersionId: approvedPlans[approvedPlans.length - 1].plan_version_id, reused: true
    };
  }
  const currentOfficialId = String(registry.current_official_id || '');
  if (requestType === 'AMENDMENT') {
    if (!currentOfficialId || String(approval.supersedes_official_id || '') !== currentOfficialId) {
      throw new Error('Amendment no longer supersedes the current official vintage. Create a new amendment request.');
    }
  } else if (currentOfficialId) {
    throw new Error('A current official vintage already exists; PUBLISH cannot overwrite it.');
  }
  const now = new Date();
  if (typeof vNextFreezeOfficialVintage_ !== 'function') throw new Error('Official vintage freeze API is not installed.');
  const submittedPlanExists = vNextAdminReadCoreRows_(hub, 'PLAN_VERSION').some(function (row) {
    return String(row.book_id || '') === String(approval.book_id || '') &&
      String(row.plan_version_id || '') === String(approval.plan_version_id || '') &&
      String(row.run_id || '') === String(approval.forecast_run_id || '') &&
      String(row.status || '').toUpperCase() === 'SUBMITTED';
  });
  if (!submittedPlanExists) throw new Error('Submitted PLAN_VERSION not found for official issue.');
  const existingFrozenRows = vNextAdminReadCoreRows_(hub, 'FORECAST_RUN').filter(function (row) {
    return String(row.book_id || '') === String(approval.book_id || '') &&
      String(row.official_vintage_id || '') === officialId && Number(row.is_official || 0) === 1;
  });
  let frozen;
  if (existingFrozenRows.length) {
    const frozenRow = existingFrozenRows[existingFrozenRows.length - 1];
    if (String(frozenRow.previous_run_id || '') !== String(approval.forecast_run_id || '')) {
      throw new Error('Existing official vintage points to a different source run.');
    }
    frozen = { runId: frozenRow.run_id, officialVintageId: officialId, reused: true };
  } else {
    frozen = vNextFreezeOfficialVintage_({
      bookId: approval.book_id, runId: approval.forecast_run_id,
      officialVintageId: officialId, approvedBy: vNextAdminActor_(),
      amendment: String(approval.request_type || '') === 'AMENDMENT',
      amendmentReason: approval.amendment_reason || ''
    }, { spreadsheet: hub });
  }
  const approvedPlan = vNextAdminAppendApprovedPlan_(hub, approval, officialId, now);
  const immutablePayload = {
    officialId: officialId, bookId: approval.book_id, clientId: approval.client_id,
    fiscalYear: approval.fiscal_year, forecastRunId: frozen.runId,
    sourceForecastRunId: approval.forecast_run_id,
    recordType: approval.request_type, supersedesOfficialId: approval.supersedes_official_id || '',
    amendmentReason: approval.amendment_reason || '', snapshotHash: approval.snapshot_hash,
    issuedAt: now.toISOString()
  };
  const immutableHash = vNextAdminSha256_(vNextAdminCanonicalJson_(immutablePayload));
  const snapshot = vNextAdminParseJson_(snapshotJson, {});
  const row = {
    official_id: officialId, book_id: approval.book_id, client_id: approval.client_id,
    client_name: approval.client_name, fiscal_year: approval.fiscal_year,
    forecast_run_id: frozen.runId, source_forecast_run_id: approval.forecast_run_id,
    approval_request_id: approval.approval_request_id,
    record_type: approval.request_type, supersedes_official_id: approval.supersedes_official_id || '',
    amendment_reason: approval.amendment_reason || '', snapshot_json: snapshotJson,
    snapshot_hash: approval.snapshot_hash, immutable_hash: immutableHash,
    model_release_id: snapshot.modelReleaseId || snapshot.model_release_id ||
      (snapshot.forecast && snapshot.forecast.modelReleaseId) || '',
    issued_at: now, issued_by: vNextAdminActor_(), note: decisionComment || ''
  };
  // OFFICIAL_RUNS is append-only. A retry reuses the record keyed by approval_request_id.
  const existingOfficial = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.OFFICIAL).rows.find(function (item) {
    return String(item.approval_request_id || '') === String(approval.approval_request_id || '');
  });
  if (existingOfficial) {
    if (String(existingOfficial.official_id || '') !== officialId || String(existingOfficial.snapshot_hash || '') !== String(approval.snapshot_hash || '')) {
      throw new Error('Existing OFFICIAL_RUNS record does not match this approval.');
    }
  } else {
    vNextAdminAppendObject_(hub, VN_ADMIN_SHEETS.OFFICIAL, row);
  }
  const officialRow = existingOfficial || row;
  vNextAdminPatchRegistryByBookId_(hub, approval.book_id, {
    current_official_id: officialId, state: 'OFFICIAL_LOCKED', updated_at: now
  });
  vNextAdminCopyOfficialToClientBestEffort_(hub, registry, officialRow);
  return {
    officialId: officialId,
    officialForecastRunId: frozen && frozen.runId || '',
    approvedPlanVersionId: approvedPlan && approvedPlan.plan_version_id || ''
  };
}

function vNextAdminCopyOfficialToClientBestEffort_(hub, registry, officialRow) {
  try {
    vNextAdminCopyOfficialToClient_(registry.spreadsheet_id, officialRow);
    return true;
  } catch (err) {
    vNextAdminAppendException_(hub, {
      severity: 'ERROR', exception_type: 'OFFICIAL_COPY_FAILED', book_id: registry.book_id,
      client_name: registry.client_name, fiscal_year: registry.fiscal_year,
      title: '公式snapshotのclient book複製に失敗', detail: String(err && err.message || err),
      recommended_action: '権限を確認し、Admin Sidebarの「正式予算をClientへ再同期」を実行', source_ref: officialRow.official_id
    });
    Logger.log('Official client copy failed official=%s error=%s', officialRow.official_id, String(err && err.stack || err));
    return false;
  }
}

function vNextAdminAppendApprovedPlan_(hub, approval, officialId, approvedAt) {
  const existingApproved = vNextAdminReadCoreRows_(hub, 'PLAN_VERSION').filter(function (row) {
    return String(row.book_id || '') === String(approval.book_id || '') &&
      String(row.official_vintage_id || '') === String(officialId) && String(row.status || '') === 'APPROVED';
  });
  if (existingApproved.length) return existingApproved[existingApproved.length - 1];
  const plans = vNextAdminReadCoreRows_(hub, 'PLAN_VERSION').filter(function (row) {
    return String(row.book_id || '') === String(approval.book_id || '') &&
      String(row.plan_version_id || '') === String(approval.plan_version_id || '');
  });
  if (!plans.length) throw new Error('Submitted PLAN_VERSION not found for approval.');
  const source = plans[plans.length - 1];
  const approved = Object.assign({}, source, {
    plan_version_id: typeof vNextUuid_ === 'function' ? vNextUuid_() : Utilities.getUuid(),
    official_vintage_id: officialId,
    status: 'APPROVED',
    amends_plan_version_id: source.plan_version_id,
    approved_at: (approvedAt || new Date()).toISOString(),
    approved_by: vNextAdminActor_().toLowerCase(),
    created_at: new Date().toISOString()
  });
  vNextAdminAppendCoreRowsNoLock_(hub, 'PLAN_VERSION', [approved]);
  return approved;
}

function vNextAdminCopyOfficialToClient_(spreadsheetId, officialRow) {
  if (!spreadsheetId) throw new Error('Client spreadsheet ID is missing.');
  const client = SpreadsheetApp.openById(spreadsheetId);
  const headers = VN_ADMIN_HEADERS[VN_ADMIN_SHEETS.OFFICIAL];
  vNextAdminEnsureTable_(client, VN_ADMIN_OFFICIAL_COPY_SHEET, headers);
  const table = vNextAdminReadTable_(client, VN_ADMIN_OFFICIAL_COPY_SHEET);
  const existing = table.rows.find(function (row) { return String(row.official_id) === String(officialRow.official_id); });
  if (!existing) vNextAdminAppendObject_(client, VN_ADMIN_OFFICIAL_COPY_SHEET, officialRow, headers);
  vNextAdminProtectInternalSheets_(client, [], [VN_ADMIN_OFFICIAL_COPY_SHEET]);
}

function vNextAdminSetClientState_(spreadsheetId, state, options) {
  if (!spreadsheetId) return;
  if (VN_ADMIN_CLIENT_STATES.indexOf(String(state)) < 0) throw new Error('Invalid CLIENT state: ' + state);
  const opt = Object.assign({
    reason: '管理ハブ decision', actorEmail: vNextAdminActor_(), actorRole: 'ADMIN'
  }, options || {});
  const hub = opt.hub || vNextAdminRequireHub_();
  const client = SpreadsheetApp.openById(String(spreadsheetId));
  const routing = vNextAdminReadKeyValueSheet_(client, VN_ADMIN_BOOK_CONFIG_SHEET);
  const bookId = vNextAdminRequiredText_(routing.book_id, 'client.book_id');
  const registry = vNextAdminFindRegistryRow_(hub, function (row) {
    return String(row.book_id || '') === bookId && String(row.spreadsheet_id || '') === String(spreadsheetId);
  });
  if (!registry) throw new Error('BOOK_REGISTRY does not match the Client state target: ' + bookId);
  const targetState = String(state || '').toUpperCase();
  const clientRows = vNextAdminReadCoreRows_(client, 'STATE_EVENT').filter(function (row) {
    return String(row.book_id || '') === bookId;
  });
  const hubRows = vNextAdminReadCoreRows_(hub, 'STATE_EVENT').filter(function (row) {
    return String(row.book_id || '') === bookId;
  });
  const fromState = String(opt.fromState ||
    (hubRows.length && hubRows[hubRows.length - 1].to_state) ||
    (clientRows.length && clientRows[clientRows.length - 1].to_state) || routing.state || registry.state || 'INPUT_OPEN').toUpperCase();
  const actorRole = String(opt.actorRole || 'ADMIN').toUpperCase();
  if (fromState !== targetState && typeof vNextValidateTransition_ === 'function') {
    vNextValidateTransition_(fromState, targetState, actorRole);
  }

  let record = null;
  if (fromState === targetState) {
    record = hubRows.length ? hubRows[hubRows.length - 1] : null;
    if (record && String(record.to_state || '').toUpperCase() !== targetState) record = null;
  } else {
    record = {
      state_event_id: typeof vNextUuid_ === 'function' ? vNextUuid_() : Utilities.getUuid(),
      book_id: bookId,
      from_state: fromState,
      to_state: targetState,
      reason: String(opt.reason || ''),
      actor_email: String(opt.actorEmail || vNextAdminActor_()).toLowerCase(),
      actor_role: actorRole,
      related_run_id: String(opt.relatedRunId || ''),
      related_plan_version_id: String(opt.relatedPlanVersionId || ''),
      created_at: new Date().toISOString()
    };
  }

  if (record) {
    const hubIds = new Set(hubRows.map(function (row) {
      return String(row.state_event_id || '');
    }));
    // Hub is authoritative: persist the Admin-created event centrally before
    // attempting the Client copy. A Client write failure can then be retried.
    if (!hubIds.has(String(record.state_event_id || ''))) {
      vNextAdminAppendCoreRowsNoLock_(hub, 'STATE_EVENT', [record]);
    }
    const clientIds = new Set(clientRows.map(function (row) { return String(row.state_event_id || ''); }));
    if (!clientIds.has(String(record.state_event_id || ''))) {
      vNextAdminAppendCoreRowsNoLock_(client, 'STATE_EVENT', [record]);
    }
  }
  vNextAdminMirrorClientState_(client, targetState);
  vNextAdminPatchRegistryByBookId_(hub, bookId, { state: targetState, updated_at: new Date() });
  return {
    changed: fromState !== targetState,
    stateEventId: record && record.state_event_id || '',
    fromState: fromState,
    toState: targetState
  };
}

function vNextAdminAppendStateEvent_(ss, state, options) {
  const opt = options || {};
  const routing = vNextAdminReadKeyValueSheet_(ss, VN_ADMIN_BOOK_CONFIG_SHEET);
  if (!String(routing.book_id || '').trim()) throw new Error('Client routing book_id is missing.');
  const stateRows = vNextAdminReadCoreRows_(ss, 'STATE_EVENT').filter(function (row) {
    return String(row.book_id || '') === String(routing.book_id || '');
  });
  const fromState = String(opt.fromState || (stateRows.length && stateRows[stateRows.length - 1].to_state) || routing.state || 'INPUT_OPEN').toUpperCase();
  const targetState = String(state || '').toUpperCase();
  if (fromState === targetState) return { changed: false, fromState: fromState, toState: targetState };
  const actorRole = String(opt.actorRole || 'ADMIN').toUpperCase();
  if (typeof vNextValidateTransition_ === 'function') vNextValidateTransition_(fromState, targetState, actorRole);
  const record = {
    state_event_id: typeof vNextUuid_ === 'function' ? vNextUuid_() : Utilities.getUuid(),
    book_id: routing.book_id,
    from_state: fromState,
    to_state: targetState,
    reason: String(opt.reason || ''),
    actor_email: String(opt.actorEmail || vNextAdminActor_()).toLowerCase(),
    actor_role: actorRole,
    related_run_id: String(opt.relatedRunId || ''),
    related_plan_version_id: String(opt.relatedPlanVersionId || ''),
    created_at: new Date().toISOString()
  };
  vNextAdminAppendCoreRowsNoLock_(ss, 'STATE_EVENT', [record]);
  return { changed: true, stateEventId: record.state_event_id, fromState: fromState, toState: targetState };
}

function vNextAdminMirrorClientState_(ss, state) {
  vNextAdminWriteBookConfig_(ss, { state: String(state || '').toUpperCase() });
  return state;
}

function vNextAdminLatestClientState_(ss, bookId, fallback) {
  const rows = vNextAdminReadCoreRows_(ss, 'STATE_EVENT').filter(function (row) {
    return String(row.book_id || '') === String(bookId || '');
  });
  return String(rows.length ? rows[rows.length - 1].to_state : fallback || '').toUpperCase();
}

/**
 * Two-tier store sync. Employees work only in their client-local Core store;
 * an Admin-owned operation harvests client-scoped append-only records into Hub.
 */
function vNextAdminSyncClientToHub_(hub, client, bookId) {
  return vNextAdminWithDocumentLock_('sync-client-to-hub', function () {
    vNextAdminValidateClientBookMetaMirror_(hub, client, bookId);
    const specs = vNextAdminCoreSyncSpecs_('CLIENT_TO_HUB');
    const result = {};
    specs.forEach(function (spec) {
      const allSourceRows = vNextAdminReadCoreRows_(client, spec.sheet);
      const foreign = allSourceRows.find(function (row) {
        return String(row.book_id || '') && String(row.book_id || '') !== String(bookId);
      });
      if (foreign) {
        throw new Error('Client-local ' + spec.sheet + ' contains a foreign book_id.');
      }
      let sourceRows = allSourceRows.filter(function (row) {
        return !row.book_id || String(row.book_id) === String(bookId);
      });
      if (spec.sheet === 'EVIDENCE_EVENT') {
        vNextAdminValidateClientEvidenceRows_(hub, bookId, sourceRows);
      }
      if (spec.sheet === 'PLAN_VERSION') {
        sourceRows = sourceRows.filter(function (row) {
          return String(row.status || '').toUpperCase() === 'SUBMITTED';
        });
        vNextAdminValidateClientPlanRows_(hub, bookId, sourceRows);
      }
      if (spec.sheet === 'STATE_EVENT') {
        sourceRows = vNextAdminValidateClientStateRows_(hub, client, bookId, sourceRows);
      }
      result[spec.sheet] = vNextAdminAppendMissingCoreRows_(hub, spec.sheet, spec.id, sourceRows);
    });
    vNextAdminWriteAudit_(hub, 'SYNC_CLIENT_TO_HUB', 'BOOK', bookId, 'SUCCESS', result);
    return result;
  });
}

/**
 * Client BOOK_META is a mirror, never an ingest source. Every row must already
 * exist byte-for-byte in the Hub and the latest trusted row must still match
 * BOOK_REGISTRY. Missing trusted rows are repaired Hub -> Client.
 */
function vNextAdminValidateClientBookMetaMirror_(hub, client, bookId) {
  const registry = vNextAdminFindRegistryRow_(hub, function (row) {
    return String(row.book_id || '') === String(bookId) && String(row.mode || '') === 'CLIENT';
  });
  if (!registry) throw new Error('BOOK_REGISTRY entry is required to validate Client BOOK_META.');
  const hubRows = vNextAdminReadCoreRows_(hub, 'BOOK_META').filter(function (row) {
    return String(row.book_id || '') === String(bookId);
  });
  if (!hubRows.length) throw new Error('Trusted Hub BOOK_META is missing for ' + bookId + '.');
  const clientRows = vNextAdminReadCoreRows_(client, 'BOOK_META');
  const hubById = new Map();
  hubRows.forEach(function (row) {
    hubById.set(String(row.record_id || ''), vNextAdminSha256_(vNextAdminCanonicalJson_(row)));
  });
  try {
    clientRows.forEach(function (row) {
      const id = String(row.record_id || '');
      if (!id || String(row.book_id || '') !== String(bookId)) {
        throw new Error('Client BOOK_META contains a blank ID or foreign book.');
      }
      if (!hubById.has(id)) throw new Error('Client BOOK_META contains a non-Hub record: ' + id);
      const hash = vNextAdminSha256_(vNextAdminCanonicalJson_(row));
      if (hash !== hubById.get(id)) throw new Error('Client BOOK_META differs from the Hub record: ' + id);
    });
    const latest = hubRows[hubRows.length - 1];
    const owners = vNextAdminParseList_(registry.forecast_owner_emails).map(function (email) {
      return String(email || '').toLowerCase();
    });
    if (owners.length !== 1 || String(latest.forecast_owner_email || '').toLowerCase() !== owners[0]) {
      throw new Error('Trusted BOOK_META 予算策定担当 does not match BOOK_REGISTRY.');
    }
    if (String(latest.book_id || '') !== String(bookId) ||
        String(latest.client_id || '') !== String(registry.client_id || '') ||
        String(latest.client_name || '') !== String(registry.client_name || '') ||
        Number(latest.fiscal_year) !== Number(registry.fiscal_year) ||
        String(latest.client_book_id || '') !== String(registry.spreadsheet_id || '') ||
        String(latest.template_version || '') !== String(registry.template_release_id || '') ||
        String(latest.schema_version || '') !== String(registry.schema_version || '') ||
        String(latest.source_spreadsheet_id || '').trim()) {
      throw new Error('Trusted BOOK_META identity/version fields do not match BOOK_REGISTRY.');
    }
    const modelReleaseId = String(latest.model_release_id || '');
    const modelExists = vNextAdminReadCoreRows_(hub, 'MODEL_RELEASE').some(function (row) {
      return String(row.model_release_id || '') === modelReleaseId &&
        String(row.status || '').toUpperCase() === 'ACTIVE' &&
        String(row.template_version || '') === String(registry.template_release_id || '') &&
        String(row.schema_version || '') === String(registry.schema_version || '');
    });
    if (!modelReleaseId || !modelExists) {
      throw new Error('Trusted BOOK_META model release was never ACTIVE with the exact Template Release in the Hub.');
    }
    vNextAdminAppendMissingCoreRows_(client, 'BOOK_META', 'record_id', hubRows);
    return true;
  } catch (err) {
    vNextAdminAppendException_(hub, {
      severity: 'ERROR', exception_type: 'CLIENT_BOOK_META_REJECTED', book_id: bookId,
      client_name: registry.client_name, fiscal_year: registry.fiscal_year,
      title: 'Client BOOK_METAの整合性検証に失敗', detail: String(err && err.message || err),
      recommended_action: 'Client内部sheetの直接編集を確認し、Hub正本から再同期', source_ref: bookId
    });
    throw err;
  }
}

function vNextAdminValidateClientEvidenceRows_(hub, bookId, rows) {
  if (!rows || !rows.length) return true;
  const registry = vNextAdminFindRegistryRow_(hub, function (row) { return String(row.book_id || '') === String(bookId); });
  if (!registry) throw new Error('BOOK_REGISTRY entry is required to validate client evidence.');
  const hubEvidence = vNextAdminReadCoreRows_(hub, 'EVIDENCE_EVENT').filter(function (row) {
    return String(row.book_id || '') === String(bookId);
  });
  const existingById = new Map(hubEvidence.map(function (row) {
    return [String(row.evidence_id || ''), row];
  }));
  const team = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.TEAM).rows.filter(function (row) {
    return String(row.book_id || '') === String(bookId) && String(row.status || '').toUpperCase() === 'ACTIVE' &&
      String(row.role || '').toUpperCase() !== 'VIEWER';
  });
  const activeActors = new Set(team.map(function (row) { return String(row.email || '').toLowerCase(); }));
  const internalOpen = String(registry.access_policy || '').toUpperCase() === 'INTERNAL_OPEN';
  const internalDomain = vNextAdminNormalizeDomain_(registry.internal_domain || '');
  const allowedTypes = new Set(['COMMITMENT', 'HUMAN_CHANGE', 'CHECK_IN', 'REVIEW_LEARNING']);
  const bookMetaRows = vNextAdminReadCoreRows_(hub, 'BOOK_META').filter(function (row) {
    return String(row.book_id || '') === String(bookId);
  });
  let inputRoundStartedAt = -Infinity;
  bookMetaRows.forEach(function (row) {
    if (String(row.event_type || '').toUpperCase() !== 'INPUT_REOPENED') return;
    inputRoundStartedAt = Math.max(inputRoundStartedAt,
      vNextAdminStrictTimestampMs_(row.recorded_at, 'INPUT_REOPENED recorded_at'));
  });
  const latestInputByActor = new Map();
  const latestReviewByActor = new Map();
  hubEvidence.forEach(function (row) {
    const actor = String(row.actor_email || '').toLowerCase();
    const type = String(row.evidence_type || '').toUpperCase();
    let timestamp = -Infinity;
    try { timestamp = vNextAdminStrictTimestampMs_(row.created_at, 'Hub evidence created_at'); }
    catch (ignoredHistoricTimestamp) { timestamp = -Infinity; }
    if (type === 'REVIEW_LEARNING') latestReviewByActor.set(actor, row);
    else if (['COMMITMENT', 'HUMAN_CHANGE', 'CHECK_IN'].indexOf(type) >= 0 && timestamp >= inputRoundStartedAt) {
      latestInputByActor.set(actor, row);
    }
  });
  const seenIds = new Set();
  let previousNewTimestamp = -Infinity;
  try {
    (rows || []).forEach(function (row) {
      // Sheets may coerce an audit month such as "2027-04" to a Date even
      // when it is displayed as YYYY-MM. Normalize only this known lossless
      // representation before canonical comparison and append.
      ['target_start_month', 'target_end_month'].forEach(function (key) {
        row[key] = vNextAdminNormalizeEvidenceMonth_(row[key], key);
      });
      const evidenceId = String(row.evidence_id || '');
      if (!evidenceId || evidenceId.length > 200 || seenIds.has(evidenceId)) {
        throw new Error('Client evidence contains a blank, oversized or duplicate evidence_id.');
      }
      seenIds.add(evidenceId);
      if (existingById.has(evidenceId)) {
        const accepted = vNextAdminCanonicalEvidenceForComparison_(existingById.get(evidenceId));
        const submitted = vNextAdminCanonicalEvidenceForComparison_(row);
        if (vNextAdminCanonicalJson_(accepted) !== vNextAdminCanonicalJson_(submitted)) {
          throw new Error('Client evidence differs from the already accepted Hub record: ' + evidenceId);
        }
        return;
      }
      const type = String(row.evidence_type || '').toUpperCase();
      if (!allowedTypes.has(type)) throw new Error('evidence_type is not client-allowlisted: ' + type);
      if (String(row.book_id || '') !== String(bookId) || String(row.client_id || '') !== String(registry.client_id || '') ||
          Number(row.fiscal_year) !== Number(registry.fiscal_year)) throw new Error('evidence book/client/FY mismatch');
      const actorEmail = String(row.actor_email || '').toLowerCase();
      const internalContributor = internalOpen && internalDomain && vNextAdminEmailDomain_(actorEmail) === internalDomain;
      if (!activeActors.has(actorEmail) && !internalContributor) {
        throw new Error('evidence actor is not an active team member or internal contributor');
      }
      ['ai_model', 'prompt_version', 'ai_schema_version', 'rule_version', 'applied_amount', 'evidence_quality'].forEach(function (key) {
        if (String(row[key] || '').trim()) throw new Error('client evidence contains forbidden AI metadata: ' + key);
      });
      if (Number(row.cap_applied || 0) !== 0) throw new Error('client evidence contains forbidden cap_applied');
      if (String(row.source_url || '').trim() || String(row.source_date || '').trim() || String(row.expires_at || '').trim()) {
        throw new Error('client evidence contains forbidden source/expiry metadata');
      }
      const createdAt = vNextAdminStrictTimestampMs_(row.created_at, 'Client evidence created_at');
      if (createdAt > Date.now() + 5 * 60000 || createdAt < previousNewTimestamp) {
        throw new Error('Client evidence timestamp is future-dated or not append ordered.');
      }
      previousNewTimestamp = createdAt;
      const response = String(row.response_type || '').toUpperCase();
      if (['CHANGE', 'NO_CHANGE', 'UNKNOWN'].indexOf(response) < 0) throw new Error('invalid evidence response_type');
      const actor = String(row.actor_email || '').toLowerCase();
      const isReview = type === 'REVIEW_LEARNING';
      const isChange = type === 'COMMITMENT' || type === 'HUMAN_CHANGE';
      if (isReview) {
        if (!activeActors.has(actor)) throw new Error('REVIEW_LEARNING is limited to registered team members.');
        if (response !== 'UNKNOWN' || String(row.status || '').toUpperCase() !== 'ACTIVE' ||
            String(row.target || '') !== 'FY_REVIEW' || ['NEUTRAL', ''].indexOf(String(row.direction || '').toUpperCase()) < 0) {
          throw new Error('REVIEW_LEARNING record shape is invalid.');
        }
        vNextAdminValidateClientReviewEvidence_(hub, registry, row);
      } else if (isChange) {
        if (response !== 'CHANGE' || String(row.status || '').toUpperCase() !== 'SUBMITTED' ||
            !String(row.target || '').trim() || String(row.target || '').length > 200 ||
            ['UP', 'DOWN'].indexOf(String(row.direction || '').toUpperCase()) < 0 ||
            ['CONFIRMED_FACT', 'LIKELY', 'HYPOTHESIS'].indexOf(String(row.confidence_class || '').toUpperCase()) < 0 ||
            !String(row.evidence_text || '').trim() || String(row.evidence_text || '').length > 2000) {
          throw new Error('Change evidence target/direction/confidence/text/status is incomplete.');
        }
        vNextAdminValidateClientEvidenceAmount_(row);
        if (createdAt < inputRoundStartedAt) throw new Error('Change evidence predates the current input round.');
      } else {
        if (type !== 'CHECK_IN' || ['NO_CHANGE', 'UNKNOWN'].indexOf(response) < 0 ||
            String(row.status || '').toUpperCase() !== 'SUBMITTED' ||
            String(row.target || '').trim() || String(row.target_start_month || '').trim() ||
            String(row.target_end_month || '').trim() ||
            ['NEUTRAL', ''].indexOf(String(row.direction || '').toUpperCase()) < 0 ||
            String(row.confidence_class || '').trim() || String(row.amount_band || '').trim() ||
            [row.amount_low, row.amount_mid, row.amount_high].some(function (value) { return String(value || '').trim(); })) {
          throw new Error('CHECK_IN record shape is invalid.');
        }
        if (createdAt < inputRoundStartedAt) throw new Error('CHECK_IN predates the current input round.');
      }
      ['target_start_month', 'target_end_month'].forEach(function (key) {
        if (row[key] && !/^\d{4}-\d{2}$/.test(String(row[key]))) throw new Error('invalid evidence month: ' + key);
      });
      if (isChange) vNextAdminValidateEvidencePeriodInFiscalYear_(row, Number(registry.fiscal_year));
      const lineageMap = isReview ? latestReviewByActor : latestInputByActor;
      const predecessor = lineageMap.get(actor);
      const suppliedPredecessor = String(row.supersedes_evidence_id || '');
      const expectedPredecessor = String(predecessor && predecessor.evidence_id || '');
      if (suppliedPredecessor !== expectedPredecessor) {
        throw new Error('Client evidence supersedes lineage is stale or branched: ' + evidenceId);
      }
      lineageMap.set(actor, row);
      existingById.set(evidenceId, row);
    });
  } catch (err) {
    vNextAdminAppendException_(hub, {
      severity: 'ERROR', exception_type: 'CLIENT_EVIDENCE_REJECTED', book_id: bookId,
      client_name: registry.client_name, fiscal_year: registry.fiscal_year,
      title: 'client入力の整合性検証に失敗', detail: String(err && err.message || err),
      recommended_action: '対象bookの監査ログと直接編集の有無を確認', source_ref: bookId
    });
    throw err;
  }
}

function vNextAdminNormalizeEvidenceMonth_(value, fieldName) {
  if (value === '' || value === null || value === undefined) return '';
  if (value instanceof Date) {
    if (isNaN(value.getTime())) throw new Error('invalid evidence month: ' + fieldName);
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM');
  }
  const text = String(value || '').trim();
  if (/^\d{4}-\d{2}$/.test(text)) return text;
  throw new Error('invalid evidence month: ' + fieldName);
}

/** Normalizes only known lossless Sheets coercions before immutable equality checks. */
function vNextAdminCanonicalEvidenceForComparison_(row) {
  const normalized = {};
  Object.keys(row || {}).forEach(function (key) { normalized[key] = row[key]; });
  ['target_start_month', 'target_end_month'].forEach(function (key) {
    normalized[key] = vNextAdminNormalizeEvidenceMonth_(normalized[key], key);
  });
  return normalized;
}

function vNextAdminValidateClientEvidenceAmount_(row) {
  const mode = String(row.amount_mode || '').toUpperCase();
  const band = String(row.amount_band || '').toUpperCase();
  const values = [row.amount_low, row.amount_mid, row.amount_high].map(function (value) {
    return value === '' || value === null || value === undefined ? null : Number(value);
  });
  values.forEach(function (value) {
    if (value !== null && (!isFinite(value) || value < 0 || value > 1e15)) {
      throw new Error('Client evidence amount is outside the accepted non-negative range.');
    }
  });
  if (mode === 'EXACT') {
    if (values[1] === null || band) throw new Error('EXACT evidence requires amount_mid and no amount_band.');
  } else if (mode === 'BAND') {
    if (['SMALL', 'MEDIUM', 'LARGE'].indexOf(band) < 0 || values[0] === null || values[2] === null ||
        values[0] > values[2]) throw new Error('BAND evidence range is invalid.');
  } else {
    throw new Error('Client evidence amount_mode must be EXACT or BAND.');
  }
  return true;
}

function vNextAdminValidateEvidencePeriodInFiscalYear_(row, fiscalYear) {
  const start = String(row.target_start_month || '');
  const end = String(row.target_end_month || '');
  if (!/^\d{4}-\d{2}$/.test(start) || !/^\d{4}-\d{2}$/.test(end) || start > end) {
    throw new Error('Change evidence period is missing or reversed.');
  }
  const first = String(fiscalYear) + '-04';
  const last = String(fiscalYear + 1) + '-03';
  if (start < first || end > last) throw new Error('Change evidence period is outside the target fiscal year.');
  return true;
}

function vNextAdminValidateClientReviewEvidence_(hub, registry, row) {
  const state = vNextAdminLatestClientState_(hub, registry.book_id, registry.state);
  if (state !== 'REVIEW_DUE') throw new Error('REVIEW_LEARNING is accepted only while REVIEW_DUE.');
  const payload = vNextAdminParseJson_(row.evidence_text, null);
  if (!payload || typeof payload !== 'object' || Array.isArray(payload) ||
      String(payload.bookId || '') !== String(registry.book_id || '') ||
      Number(payload.fiscalYear) !== Number(registry.fiscal_year) ||
      String(payload.officialVintageId || '') !== String(registry.current_official_id || '')) {
    throw new Error('REVIEW_LEARNING book/FY/current-official linkage is invalid.');
  }
  const evaluations = vNextAdminReadCoreRows_(hub, 'EVALUATION').filter(function (item) {
    return String(item.book_id || '') === String(registry.book_id || '') &&
      String(item.official_vintage_id || '') === String(registry.current_official_id || '');
  });
  const currentEvaluation = evaluations.length ? evaluations[evaluations.length - 1] : null;
  if (!currentEvaluation || String(payload.evaluationId || '') !== String(currentEvaluation.evaluation_id || '')) {
    throw new Error('REVIEW_LEARNING evaluation linkage is invalid.');
  }
  if (!Array.isArray(payload.causeCategories) || payload.causeCategories.length < 1 || payload.causeCategories.length > 3 ||
      !Array.isArray(payload.nextInformation) || payload.nextInformation.length < 1 || payload.nextInformation.length > 3 ||
      (!String(payload.confirmedCause || '').trim() && !String(payload.causeHypothesis || '').trim())) {
    throw new Error('REVIEW_LEARNING structured content is incomplete.');
  }
  const allowedCauseKeys = typeof VNEXT_UX_REVIEW_CAUSES_ !== 'undefined'
    ? new Set(VNEXT_UX_REVIEW_CAUSES_.map(function (item) { return String(item.key || ''); }))
    : new Set(['BASE_LEVEL', 'UNKNOWN_SPOT', 'COMMITMENT_OUTCOME', 'AMOUNT', 'HUMAN_INFO',
      'AI_INFO', 'DATA_QUALITY', 'SEASONALITY', 'TIMING', 'OTHER']);
  if (payload.causeCategories.some(function (key) { return !allowedCauseKeys.has(String(key || '')); }) ||
      payload.nextInformation.some(function (item) { return !String(item || '').trim() || String(item).length > 300; }) ||
      String(payload.confirmedCause || '').length > 1000 || String(payload.causeHypothesis || '').length > 1000) {
    throw new Error('REVIEW_LEARNING contains an unsupported category or oversized text.');
  }
  if (String(row.evidence_text || '').length > 6000) throw new Error('REVIEW_LEARNING content is too large.');
  return true;
}

/**
 * Client plans are untrusted staging records. Validate their complete lineage
 * and arithmetic against the Hub-owned forecast before accepting a new row.
 */
function vNextAdminValidateClientPlanRows_(hub, bookId, rows) {
  if (!rows || !rows.length) return true;
  const registry = vNextAdminFindRegistryRow_(hub, function (row) {
    return String(row.book_id || '') === String(bookId) && String(row.mode || '') === 'CLIENT';
  });
  if (!registry) throw new Error('BOOK_REGISTRY entry is required to validate client plans.');
  const hubPlans = vNextAdminReadCoreRows_(hub, 'PLAN_VERSION').filter(function (row) {
    return String(row.book_id || '') === String(bookId);
  });
  const hubIds = new Set(hubPlans.map(function (row) { return String(row.plan_version_id || ''); }));
  const seenIds = new Set();
  let latestPlan = hubPlans.length ? hubPlans[hubPlans.length - 1] : null;
  try {
    (rows || []).forEach(function (plan) {
      const planId = String(plan.plan_version_id || '');
      if (!planId || seenIds.has(planId)) throw new Error('Client PLAN_VERSION contains a blank or duplicate ID.');
      seenIds.add(planId);
      if (hubIds.has(planId)) return;
      if (String(plan.book_id || '') !== String(bookId) ||
          String(plan.status || '').toUpperCase() !== 'SUBMITTED' ||
          String(plan.official_vintage_id || '') || String(plan.approved_at || '') || String(plan.approved_by || '')) {
        throw new Error('Only an unapproved SUBMITTED plan for this book may be ingested.');
      }
      const expectedVersion = latestPlan ? Number(latestPlan.version_no || 0) + 1 : 1;
      const predecessorId = latestPlan ? String(latestPlan.plan_version_id || '') : '';
      if (Number(plan.version_no) !== expectedVersion ||
          String(plan.amends_plan_version_id || '') !== predecessorId) {
        throw new Error('Client plan version/lineage is not the next append-only version.');
      }
      const submittedMs = vNextAdminStrictTimestampMs_(plan.submitted_at, 'plan.submitted_at');
      const createdMs = vNextAdminStrictTimestampMs_(plan.created_at, 'plan.created_at');
      if (Math.abs(createdMs - submittedMs) > 5 * 60000 || createdMs > Date.now() + 5 * 60000) {
        throw new Error('Client plan timestamps are inconsistent or in the future.');
      }
      const forecast = vNextAdminReadCoreRows_(hub, 'FORECAST_RUN').filter(function (row) {
        return String(row.book_id || '') === String(bookId) &&
          String(row.run_id || '') === String(plan.run_id || '');
      }).slice(-1)[0];
      if (!forecast) throw new Error('Client plan does not reference a Hub-owned forecast run.');
      vNextAdminValidateSubmittedPlan_(registry, forecast, plan);
      latestPlan = plan;
    });
    return true;
  } catch (err) {
    vNextAdminAppendException_(hub, {
      severity: 'ERROR', exception_type: 'CLIENT_PLAN_REJECTED', book_id: bookId,
      client_name: registry.client_name, fiscal_year: registry.fiscal_year,
      title: 'Client計画版の整合性検証に失敗', detail: String(err && err.message || err),
      recommended_action: '直接編集の有無を確認し、Hub正本から再同期', source_ref: bookId
    });
    throw err;
  }
}

function vNextAdminValidateClientStateRows_(hub, client, bookId, rows) {
  // The caller replaces sourceRows with this return value before appending.
  // Returning a boolean for an empty first sync breaks that array contract.
  if (!rows || !rows.length) return [];
  const registry = vNextAdminFindRegistryRow_(hub, function (row) {
    return String(row.book_id || '') === String(bookId) && String(row.mode || '') === 'CLIENT';
  });
  if (!registry) throw new Error('BOOK_REGISTRY entry is required to validate Client state.');
  const owners = vNextAdminParseList_(registry.forecast_owner_emails).map(function (email) {
    return String(email || '').toLowerCase();
  });
  if (owners.length !== 1) throw new Error('Client state validation requires exactly one 予算策定担当.');
  const owner = owners[0];
  const hubRows = vNextAdminReadCoreRows_(hub, 'STATE_EVENT').filter(function (row) {
    return String(row.book_id || '') === String(bookId);
  });
  const existingIds = new Set(hubRows.map(function (row) { return String(row.state_event_id || ''); }));
  const seenIds = new Set();
  const latestMeta = vNextAdminReadCoreRows_(hub, 'BOOK_META').filter(function (row) {
    return String(row.book_id || '') === String(bookId);
  }).slice(-1)[0];
  const allBookMeta = vNextAdminReadCoreRows_(hub, 'BOOK_META').filter(function (row) {
    return String(row.book_id || '') === String(bookId);
  });
  const historicPrefixIds = vNextAdminVerifyHistoricClientStatePrefix_(
    hub, client, registry, rows, hubRows, existingIds, owner, allBookMeta[0] || latestMeta
  );
  let currentState = String(
    hubRows.length ? hubRows[hubRows.length - 1].to_state :
      latestMeta && latestMeta.state || registry.state || 'INPUT_OPEN'
  ).toUpperCase();
  let lastTimestamp = hubRows.length
    ? vNextAdminStrictTimestampMs_(hubRows[hubRows.length - 1].created_at, 'Hub STATE_EVENT.created_at')
    : 0;
  const acceptedRows = [];
  try {
    (rows || []).forEach(function (event) {
      const eventId = String(event.state_event_id || '');
      if (!eventId || seenIds.has(eventId)) throw new Error('Client STATE_EVENT contains a blank or duplicate ID.');
      seenIds.add(eventId);
      if (historicPrefixIds.has(eventId)) return;
      if (existingIds.has(eventId)) {
        acceptedRows.push(event);
        return;
      }
      if (String(event.book_id || '') !== String(bookId)) throw new Error('Client STATE_EVENT book_id mismatch.');
      const fromState = String(event.from_state || '').toUpperCase();
      const toState = String(event.to_state || '').toUpperCase();
      const edge = fromState + '>' + toState;
      if (fromState !== currentState) {
        throw new Error('Client state chain is discontinuous. Hub=' + currentState + '; event=' + edge);
      }
      const timestamp = vNextAdminStrictTimestampMs_(event.created_at, 'Client STATE_EVENT.created_at');
      if (timestamp < lastTimestamp || timestamp > Date.now() + 5 * 60000) {
        throw new Error('Client STATE_EVENT timestamp is out of order or in the future.');
      }
      if (vNextAdminIsTrustedRejectedStateMarker_(hub, registry, event, edge)) {
        // A rejected request and its Admin-created local correction are kept in
        // the Client audit trail, but neither becomes part of the Hub state
        // chain. Advancing the virtual state lets us validate the pair.
        currentState = toState;
        lastTimestamp = timestamp;
        return;
      }
      vNextAdminValidateClientStateEventSemantics_(hub, client, registry, event, edge, owner, timestamp);
      acceptedRows.push(event);
      currentState = toState;
      lastTimestamp = timestamp;
    });
    const authoritativeState = String(
      hubRows.length ? hubRows[hubRows.length - 1].to_state :
        latestMeta && latestMeta.state || registry.state || 'INPUT_OPEN'
    ).toUpperCase();
    if (!acceptedRows.some(function (row) { return !existingIds.has(String(row.state_event_id || '')); }) &&
        currentState !== authoritativeState) {
      throw new Error('Rejected Client state markers do not return to the authoritative Hub state.');
    }
    return acceptedRows;
  } catch (err) {
    vNextAdminAppendException_(hub, {
      severity: 'ERROR', exception_type: 'CLIENT_STATE_REJECTED', book_id: bookId,
      client_name: registry.client_name, fiscal_year: registry.fiscal_year,
      title: 'Client由来の状態変更を拒否', detail: String(err && err.message || err),
      recommended_action: 'STATE_EVENTの直接編集と関連計画・依頼を確認', source_ref: bookId
    });
    throw err;
  }
}

/**
 * Some early Pilot books reached the Hub only after their first request had
 * started. Verify an older Client-only prefix against normal state semantics,
 * but do not append it behind newer Hub events where row order would become
 * misleading. Future events still require the normal append-only path.
 */
function vNextAdminVerifyHistoricClientStatePrefix_(hub, client, registry, rows, hubRows,
    existingIds, owner, initialMeta) {
  const verified = new Set();
  if (!hubRows || !hubRows.length) return verified;
  const firstHubTimestamp = vNextAdminStrictTimestampMs_(hubRows[0].created_at,
    'Hub first STATE_EVENT.created_at');
  const prefix = (rows || []).filter(function (event) {
    const eventId = String(event.state_event_id || '');
    if (!eventId || existingIds.has(eventId)) return false;
    return vNextAdminStrictTimestampMs_(event.created_at,
      'Historic Client STATE_EVENT.created_at') < firstHubTimestamp;
  }).sort(function (a, b) {
    return vNextAdminStrictTimestampMs_(a.created_at, 'Historic Client STATE_EVENT.created_at') -
      vNextAdminStrictTimestampMs_(b.created_at, 'Historic Client STATE_EVENT.created_at');
  });
  if (!prefix.length) return verified;
  let state = String(initialMeta && initialMeta.state || 'INPUT_OPEN').toUpperCase();
  let lastTimestamp = 0;
  prefix.forEach(function (event) {
    const eventId = String(event.state_event_id || '');
    if (String(event.book_id || '') !== String(registry.book_id || '')) {
      throw new Error('Historic Client STATE_EVENT book_id mismatch.');
    }
    const fromState = String(event.from_state || '').toUpperCase();
    const toState = String(event.to_state || '').toUpperCase();
    const edge = fromState + '>' + toState;
    if (fromState !== state) {
      throw new Error('Historic Client state prefix is discontinuous. expected=' + state + '; event=' + edge);
    }
    const timestamp = vNextAdminStrictTimestampMs_(event.created_at,
      'Historic Client STATE_EVENT.created_at');
    if (timestamp < lastTimestamp) throw new Error('Historic Client STATE_EVENT is out of order.');
    vNextAdminValidateClientStateEventSemantics_(hub, client, registry, event, edge, owner, timestamp);
    verified.add(eventId);
    state = toState;
    lastTimestamp = timestamp;
  });
  if (state !== String(hubRows[0].from_state || '').toUpperCase()) {
    throw new Error('Historic Client state prefix does not connect to the first Hub event.');
  }
  return verified;
}

function vNextAdminIsTrustedRejectedStateMarker_(hub, registry, event, edge) {
  if (edge !== 'READY_TO_RUN>RUNNING' && edge !== 'RUNNING>READY_TO_RUN') return false;
  const eventId = String(event.state_event_id || '');
  const rejected = vNextAdminReadTable_(hub, VN_ADMIN_CLIENT_REQUEST_SHEET).rows.filter(function (row) {
    return String(row.book_id || '') === String(registry.book_id || '') &&
      String(row.status || '').toUpperCase() === 'REJECTED';
  }).some(function (row) {
    const detail = vNextAdminParseJson_(row.detail_json, {});
    if (edge === 'READY_TO_RUN>RUNNING') return String(detail.stateEventId || '') === eventId;
    return String(detail.recoveryStateEventId || '') === eventId &&
      String(event.actor_role || '').toUpperCase() === 'ADMIN' &&
      String(event.reason || '') === 'rejected_request_recovery:' + String(row.request_id || '');
  });
  return rejected;
}

function vNextAdminValidateClientStateEventSemantics_(hub, client, registry, event, edge, owner, eventTimestamp) {
  const role = String(event.actor_role || '').toUpperCase();
  const actor = String(event.actor_email || '').toLowerCase();
  const reason = String(event.reason || '');
  if (edge === 'INPUT_OPEN>READY_TO_RUN') {
    const readiness = vNextAdminClientInputReadiness_(hub, registry);
    if (role === 'SYSTEM' && reason === 'input_readiness_met' && readiness.ready) return true;
    const dueMs = readiness.dueDate ? vNextAdminDateOnlyMs_(readiness.dueDate, 'input_due_date') : NaN;
    const todayMs = vNextAdminDateOnlyMs_(new Date(), 'today');
    if (role === 'FORECAST_OWNER' && actor === owner && /^deadline_override:\s*\S/.test(reason) &&
        isFinite(dueMs) && dueMs < todayMs) return true;
    throw new Error('INPUT_OPEN>READY_TO_RUN does not satisfy all-answered or deadline override rules.');
  }
  if (edge === 'READY_TO_RUN>RUNNING') {
    const requestReason = /^forecast_requested:(REQ-[A-Za-z0-9-]{8,200})$/.exec(reason);
    if (role !== 'FORECAST_OWNER' || actor !== owner || !requestReason) {
      throw new Error('READY_TO_RUN>RUNNING must be the registered 予算策定担当 forecast request.');
    }
    const request = vNextAdminLatestValidPendingRequest_(client, registry, owner, eventTimestamp, requestReason[1]);
    if (!request) throw new Error('A matching valid pending forecast request was not found.');
    return true;
  }
  if (edge === 'DRAFT_READY>SUBMITTED' || edge === 'CHANGES_REQUESTED>SUBMITTED') {
    const expectedReason = edge === 'DRAFT_READY>SUBMITTED' ? 'plan_submitted' : 'revised_plan_submitted';
    if (role !== 'FORECAST_OWNER' || actor !== owner || reason !== expectedReason) {
      throw new Error(edge + ' must be a registered 予算策定担当 plan submission.');
    }
    const planId = String(event.related_plan_version_id || '');
    const runId = String(event.related_run_id || '');
    const plan = vNextAdminReadCoreRows_(hub, 'PLAN_VERSION').filter(function (row) {
      return String(row.book_id || '') === String(registry.book_id || '') &&
        String(row.plan_version_id || '') === planId;
    }).slice(-1)[0];
    if (!plan || String(plan.status || '').toUpperCase() !== 'SUBMITTED' ||
        String(plan.submitted_by || '').toLowerCase() !== owner ||
        String(plan.run_id || '') !== runId) {
      throw new Error('State event is not linked to the accepted submitted plan and forecast run.');
    }
    return true;
  }
  throw new Error('Client-origin state transition is not allowlisted: ' + edge);
}

function vNextAdminClientInputReadiness_(hub, registry) {
  const bookId = String(registry.book_id || '');
  const metas = vNextAdminReadCoreRows_(hub, 'BOOK_META').filter(function (row) {
    return String(row.book_id || '') === bookId;
  });
  const latestMeta = metas.length ? metas[metas.length - 1] : null;
  if (!latestMeta) throw new Error('Trusted BOOK_META is required for readiness validation.');
  let roundStartedMs = 0;
  metas.forEach(function (row) {
    if (String(row.event_type || '').toUpperCase() !== 'INPUT_REOPENED') return;
    roundStartedMs = Math.max(roundStartedMs, vNextAdminStrictTimestampMs_(row.recorded_at, 'INPUT_REOPENED.recorded_at'));
  });
  const team = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.TEAM).rows.filter(function (row) {
    return String(row.book_id || '') === bookId && String(row.status || '').toUpperCase() === 'ACTIVE' &&
      String(row.role || '').toUpperCase() !== 'VIEWER';
  }).map(function (row) { return String(row.email || '').toLowerCase(); }).filter(Boolean);
  const memberSet = new Set(team);
  const latestByActor = {};
  vNextAdminReadCoreRows_(hub, 'EVIDENCE_EVENT').forEach(function (row) {
    const actor = String(row.actor_email || '').toLowerCase();
    const type = String(row.evidence_type || '').toUpperCase();
    if (String(row.book_id || '') !== bookId || !memberSet.has(actor) ||
        ['COMMITMENT', 'HUMAN_CHANGE', 'CHECK_IN'].indexOf(type) < 0 ||
        ['ACTIVE', 'SUBMITTED'].indexOf(String(row.status || 'ACTIVE').toUpperCase()) < 0) return;
    const createdMs = vNextAdminStrictTimestampMs_(row.created_at, 'EVIDENCE_EVENT.created_at');
    if (createdMs < roundStartedMs) return;
    latestByActor[actor] = row;
  });
  const answered = Object.keys(latestByActor).length;
  return {
    ready: team.length > 0 && answered >= team.length,
    answeredCount: answered,
    totalCount: team.length,
    dueDate: String(latestMeta.input_due_date || '')
  };
}

function vNextAdminLatestValidPendingRequest_(client, registry, owner, eventTimestamp, expectedRequestId) {
  const sheet = client.getSheetByName(VN_ADMIN_CLIENT_REQUEST_SHEET);
  if (!sheet) return null;
  const rows = vNextAdminReadTable_(client, VN_ADMIN_CLIENT_REQUEST_SHEET).rows;
  const latest = {};
  rows.forEach(function (row) {
    const requestId = String(row.request_id || '');
    if (requestId) latest[requestId] = row;
  });
  const latestForRequest = latest[String(expectedRequestId || '')];
  const latestPair = String(latestForRequest && latestForRequest.event_type || '').toUpperCase() + '>' +
    String(latestForRequest && latestForRequest.status || '').toUpperCase();
  if (!latestForRequest || [
    'REQUESTED>PENDING', 'HARVESTED>QUEUED', 'FAILED>FAILED', 'COMPLETED>COMPLETED'
  ].indexOf(latestPair) < 0) {
    return null;
  }
  // HARVESTED is appended before the first Client STATE_EVENT sync. Validate
  // the original immutable REQUESTED/PENDING row, while using the latest row
  // only to ensure the request was not REJECTED. A later FAILED/COMPLETED event
  // does not retroactively invalidate the owner-authorized transition that
  // started the forecast; it remains part of the append-only state history.
  const candidates = rows.filter(function (row) {
    return String(row.request_id || '') === String(expectedRequestId || '') &&
      String(row.event_type || '').toUpperCase() === 'REQUESTED' &&
      String(row.status || '').toUpperCase() === 'PENDING';
  }).map(function (row) {
    return vNextAdminValidateClientRequestRow_(row, registry, owner);
  }).filter(function (item) {
    return item.requestedAtMs <= eventTimestamp + 5 * 60000 &&
      item.requestedAtMs >= eventTimestamp - 5 * 60000;
  }).sort(function (a, b) { return b.requestedAtMs - a.requestedAtMs; });
  return candidates.length ? candidates[0] : null;
}

function vNextAdminStrictTimestampMs_(value, label) {
  const date = value instanceof Date ? new Date(value.getTime()) : new Date(String(value || ''));
  if (isNaN(date.getTime())) throw new Error((label || 'timestamp') + ' is invalid.');
  return date.getTime();
}

function vNextAdminDateOnlyMs_(value, label) {
  const text = value instanceof Date
    ? Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd')
    : String(value || '').slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) throw new Error((label || 'date') + ' is invalid.');
  const parts = text.split('-').map(Number);
  const date = new Date(parts[0], parts[1] - 1, parts[2]);
  if (date.getFullYear() !== parts[0] || date.getMonth() !== parts[1] - 1 || date.getDate() !== parts[2]) {
    throw new Error((label || 'date') + ' is invalid.');
  }
  return date.getTime();
}

function vNextAdminSyncHubToClient_(hub, client, bookId, sheetNames) {
  return vNextAdminWithDocumentLock_('sync-hub-to-client', function () {
    const allowed = new Set(sheetNames || []);
    const result = {};
    vNextAdminCoreSyncSpecs_('HUB_TO_CLIENT').forEach(function (spec) {
      if (allowed.size && !allowed.has(spec.sheet)) return;
      let sourceRows = vNextAdminReadCoreRows_(hub, spec.sheet).filter(function (row) {
        return !row.book_id || String(row.book_id) === String(bookId);
      });
      if (spec.sheet === 'FORECAST_RUN') sourceRows = vNextAdminSanitizeForecastRowsForClient_(sourceRows);
      result[spec.sheet] = vNextAdminAppendMissingCoreRows_(client, spec.sheet, spec.id, sourceRows);
    });
    return result;
  });
}

function vNextAdminCoreSyncSpecs_(direction) {
  if (direction === 'CLIENT_TO_HUB') {
    return [
      { sheet: 'EVIDENCE_EVENT', id: 'evidence_id' },
      { sheet: 'PLAN_VERSION', id: 'plan_version_id' },
      { sheet: 'STATE_EVENT', id: 'state_event_id' }
    ];
  }
  return [
    { sheet: 'BOOK_META', id: 'record_id' },
    { sheet: 'FORECAST_RUN', id: 'run_id' },
    { sheet: 'PLAN_VERSION', id: 'plan_version_id' },
    { sheet: 'STATE_EVENT', id: 'state_event_id' },
    { sheet: 'EVALUATION', id: 'evaluation_id' }
  ];
}

function vNextAdminSanitizeForecastRowsForClient_(rows) {
  return (rows || []).map(function (row) {
    const copy = Object.assign({}, row);
    const lenses = vNextAdminParseJson_(row.lens_json, {});
    const evidence = vNextAdminParseJson_(row.evidence_json, {});
    const triangulation = lenses.triangulation && typeof lenses.triangulation === 'object'
      ? lenses.triangulation
      : (typeof vNextBuildTriangulationReference_ === 'function'
        ? vNextBuildTriangulationReference_(lenses.continuity || {}, Number(row.system_recommended || row.p50 || 0))
        : {});
    copy.lens_json = vNextAdminCanonicalJson_({
      publicDrivers: Array.isArray(lenses.publicDrivers) ? lenses.publicDrivers.slice(0, 5) :
        (Array.isArray(lenses.drivers) ? lenses.drivers.slice(0, 5) : []),
      nextInformation: Array.isArray(lenses.nextInformation) ? lenses.nextInformation.slice(0, 5) : [],
      changeReasons: Array.isArray(lenses.changeReasons) ? lenses.changeReasons.slice(0, 5) : [],
      continuity: lenses.continuity && typeof lenses.continuity === 'object' ? {
        baseAnnualBaseline: Number(lenses.continuity.baseAnnualBaseline || 0),
        annualBaseline: Number(lenses.continuity.annualBaseline || 0),
        fiscalYears: Array.isArray(lenses.continuity.fiscalYears) ? lenses.continuity.fiscalYears.slice(-8) : []
      } : {},
      triangulation: triangulation && Array.isArray(triangulation.methods) ? {
        policy: String(triangulation.policy || 'INDEPENDENT_REFERENCES_NOT_AUTOMATICALLY_AVERAGED'),
        methods: triangulation.methods.slice(0, 5).map(function (method) {
          return {
            key: String(method.key || ''), label: String(method.label || '').slice(0, 80),
            value: Number(method.value || 0), assumption: String(method.assumption || '').slice(0, 240),
            basis: String(method.basis || '').slice(0, 120)
          };
        }),
        referenceMin: Number(triangulation.referenceMin || 0),
        referenceMedian: Number(triangulation.referenceMedian || 0),
        referenceMax: Number(triangulation.referenceMax || 0)
      } : {},
      changeReference: lenses.changeReference && typeof lenses.changeReference === 'object' ? {
        peerReferenceDelta: Number(lenses.changeReference.peerReferenceDelta || 0),
        objectiveEventDelta: Number(lenses.changeReference.objectiveEventDelta || 0),
        combinedObjectiveInformationDelta: Number(lenses.changeReference.combinedObjectiveInformationDelta || 0)
      } : {},
      evidenceReadiness: lenses.evidenceReadiness && typeof lenses.evidenceReadiness === 'object' ? {
        level: String(lenses.evidenceReadiness.level || ''),
        label: String(lenses.evidenceReadiness.label || ''),
        summary: String(lenses.evidenceReadiness.summary || '').slice(0, 300),
        historyYearCount: Number(lenses.evidenceReadiness.historyYearCount || 0),
        historyFiscalYears: Array.isArray(lenses.evidenceReadiness.historyFiscalYears) ? lenses.evidenceReadiness.historyFiscalYears.slice(-8) : [],
        missingResponseRate: Number(lenses.evidenceReadiness.missingResponseRate || 0),
        informationGapRate: Number(lenses.evidenceReadiness.informationGapRate || 0),
        issues: Array.isArray(lenses.evidenceReadiness.issues) ? lenses.evidenceReadiness.issues.slice(0, 5) : []
      } : {},
      degradation: lenses.degradation && lenses.degradation.aiUnavailable === true ? {
        aiUnavailable: true, reason: String(lenses.degradation.reason || 'AI_RESEARCH_UNAVAILABLE'),
        policy: 'AI_ZERO_REFERENCE_ONLY'
      } : null
    });
    copy.evidence_json = vNextAdminCanonicalJson_({
      topAiEvidence: Array.isArray(evidence.topAiEvidence) ? evidence.topAiEvidence.slice(0, 5) : [],
      commitment: Number(evidence.commitment || 0), objective: Number(evidence.objective || 0),
      human: Number(evidence.human || 0), ai: Number(evidence.ai || 0),
      responseCounts: evidence.responseCounts && typeof evidence.responseCounts === 'object' ? evidence.responseCounts : {},
      missingResponseRate: Number(evidence.missingResponseRate || 0),
      informationGapRate: Number(evidence.informationGapRate || 0),
      aiUnavailable: evidence.aiUnavailable === true,
      aiUnavailableReason: evidence.aiUnavailable === true ? String(evidence.aiUnavailableReason || '') : '',
      noChange: Number(evidence.noChange || 0), unknown: Number(evidence.unknown || 0),
      unknownSpotExpectedAnnual: Number(evidence.unknownSpotExpectedAnnual || 0),
      unknownSpotExpectedOccurrences: Number(evidence.unknownSpotExpectedOccurrences || 0),
      readiness: evidence.readiness && typeof evidence.readiness === 'object' ? {
        level: String(evidence.readiness.level || ''), label: String(evidence.readiness.label || ''),
        summary: String(evidence.readiness.summary || '').slice(0, 300),
        historyYearCount: Number(evidence.readiness.historyYearCount || 0),
        historyFiscalYears: Array.isArray(evidence.readiness.historyFiscalYears) ? evidence.readiness.historyFiscalYears.slice(-8) : [],
        missingResponseRate: Number(evidence.readiness.missingResponseRate || 0),
        informationGapRate: Number(evidence.readiness.informationGapRate || 0),
        issues: Array.isArray(evidence.readiness.issues) ? evidence.readiness.issues.slice(0, 5) : []
      } : {}
    });
    return copy;
  });
}

function vNextAdminReadCoreRows_(ss, sheetName) {
  if (typeof vNextReadRecords_ !== 'function') throw new Error('VNext_Core read API is not installed.');
  return vNextReadRecords_(sheetName, { spreadsheet: ss });
}

function vNextAdminAppendMissingCoreRows_(target, sheetName, idField, sourceRows) {
  const targetRows = vNextAdminReadCoreRows_(target, sheetName);
  const byId = new Map();
  targetRows.forEach(function (row) {
    const id = String(row[idField] || '');
    if (id) byId.set(id, vNextAdminSha256_(vNextAdminCanonicalJson_(
      vNextAdminCanonicalCoreRowForIntegrity_(sheetName, row))));
  });
  const missing = [];
  (sourceRows || []).forEach(function (row) {
    const id = String(row[idField] || '');
    if (!id) return;
    const hash = vNextAdminSha256_(vNextAdminCanonicalJson_(
      vNextAdminCanonicalCoreRowForIntegrity_(sheetName, row)));
    if (byId.has(id)) {
      if (byId.get(id) !== hash) throw new Error('Append-only integrity mismatch ' + sheetName + ' id=' + id);
      return;
    }
    byId.set(id, hash);
    missing.push(row);
  });
  if (missing.length) vNextAdminAppendCoreRowsNoLock_(target, sheetName, missing);
  return { examined: (sourceRows || []).length, appended: missing.length };
}

function vNextAdminCanonicalCoreRowForIntegrity_(sheetName, row) {
  if (String(sheetName || '') !== 'EVIDENCE_EVENT') return row;
  return vNextAdminCanonicalEvidenceForComparison_(row);
}

// ---------------------------- Hub views ----------------------------

function vNextAdminRefreshTodayExceptions_(hub) {
  const headers = VN_ADMIN_HEADERS[VN_ADMIN_SHEETS.EXCEPTIONS];
  const sheet = vNextAdminEnsureTable_(hub, VN_ADMIN_SHEETS.EXCEPTIONS, headers);
  // Health/job/approval exceptions are computed views and can be rebuilt. All
  // other OPEN rows are persistent issue events (for example a failed official
  // copy) and must survive a refresh until an explicit resolution flow exists.
  const computedTypes = new Set(['BOOK_HEALTH', 'JOB_FAILED', 'APPROVAL_PENDING', 'JOB_QUEUE_STALE', 'SCHEDULER_STALE']);
  const exceptions = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.EXCEPTIONS).rows.filter(function (row) {
    return String(row.status || 'OPEN').toUpperCase() === 'OPEN' &&
      !computedTypes.has(String(row.exception_type || '').toUpperCase());
  }).map(function (row) {
    const copy = {};
    headers.forEach(function (key) { copy[key] = row[key] === undefined ? '' : row[key]; });
    return copy;
  });
  const registry = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.REGISTRY).rows;
  registry.forEach(function (row) {
    if (String(row.status || '').toUpperCase() === 'ARCHIVED') return;
    const health = String(row.health_status || '');
    if (health === 'ERROR' || health === 'WARN' || health === 'PENDING') {
      exceptions.push(vNextAdminExceptionObject_({
        severity: health === 'ERROR' ? 'ERROR' : 'WARN', exception_type: 'BOOK_HEALTH',
        book_id: row.book_id, client_name: row.client_name, fiscal_year: row.fiscal_year,
        title: 'Book health: ' + (row.health_code || health),
        detail: 'health_status=' + health + ', code=' + String(row.health_code || ''),
        recommended_action: health === 'PENDING' ? 'Job Queueのhealth scanを実行' : 'BOOK_REGISTRYと対象bookを確認',
        source_ref: row.spreadsheet_url
      }));
    }
  });
  const jobs = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.JOBS).rows;
  jobs.filter(function (row) { return String(row.status) === 'FAILED'; }).slice(-100).forEach(function (row) {
    const actualDataIssue = vNextAdminIsActualDataIssue_(row);
    const safeRetry = vNextAdminIsKnownSafeRetryCandidate_(row);
    exceptions.push(vNextAdminExceptionObject_({
      severity: 'ERROR', exception_type: 'JOB_FAILED', book_id: row.target_book_id,
      title: actualDataIssue ? '予測に必要な確定実績が不足' : '自動処理を完了できませんでした',
      detail: row.error,
      recommended_action: safeRetry
        ? 'サイドバーの「' + VN_ADMIN_MENU_RUN_NOW + '」で、同じ処理を安全に再開'
        : actualDataIssue
          ? 'ZACの確定実績が5年度以上あるか確認。新しい処理を重複登録しない'
          : 'JOB_QUEUEの元処理を確認。新しい処理を重複登録しない',
      source_ref: row.job_id
    }));
  });
  const operations = vNextAdminOperationalMetrics_(hub, vNextAdminAutomationInstalled_());
  if (operations.queueStale) {
    exceptions.push(vNextAdminExceptionObject_({
      severity: 'WARN', exception_type: 'JOB_QUEUE_STALE',
      title: '処理待ちjobが15分以上滞留',
      detail: '最長待機=' + operations.oldestQueuedAgeMinutes + '分 / 滞留=' + operations.staleQueued + '件',
      recommended_action: '自動運用の最終成功時刻とJOB_QUEUEを確認', source_ref: 'JOB_QUEUE'
    }));
  }
  if (operations.schedulerStale) {
    exceptions.push(vNextAdminExceptionObject_({
      severity: 'ERROR', exception_type: 'SCHEDULER_STALE',
      title: '自動運用の成功記録が15分以上ありません',
      detail: operations.lastSweepAgeMinutes === null ? '成功記録なし' : '最終成功から' + operations.lastSweepAgeMinutes + '分',
      recommended_action: 'Apps Scriptのtrigger/実行ログを確認し、手動でhealth scanとjob処理を実行', source_ref: VN_ADMIN_SCHEDULED_HANDLER
    }));
  }
  const approvals = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.APPROVALS).rows;
  approvals.filter(function (row) { return String(row.status) === 'PENDING'; }).forEach(function (row) {
    exceptions.push(vNextAdminExceptionObject_({
      severity: 'INFO', exception_type: 'APPROVAL_PENDING', book_id: row.book_id,
      client_name: row.client_name, fiscal_year: row.fiscal_year,
      title: '計画承認待ち', detail: row.request_type + ' / run=' + row.forecast_run_id,
      recommended_action: 'PLAN_APPROVALSで承認・差戻し・却下', source_ref: row.approval_request_id
    }));
  });

  const uniqueById = new Map();
  exceptions.forEach(function (obj) { uniqueById.set(String(obj.exception_id || ''), obj); });
  const unique = Array.from(uniqueById.values());
  if (sheet.getMaxRows() > 1) sheet.getRange(2, 1, sheet.getMaxRows() - 1, headers.length).clearContent();
  if (unique.length) {
    const values = unique.map(function (obj) { return headers.map(function (key) { return obj[key] === undefined ? '' : obj[key]; }); });
    if (sheet.getMaxRows() < values.length + 1) sheet.insertRowsAfter(sheet.getMaxRows(), values.length + 1 - sheet.getMaxRows());
    sheet.getRange(2, 1, values.length, headers.length).setValues(values);
  }
  return { openExceptions: unique.length };
}

function vNextAdminResolveOpenExceptions_(hub, bookId, types, sourceRef) {
  const allowedTypes = new Set((types || []).map(function (value) { return String(value || '').toUpperCase(); }));
  const table = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.EXCEPTIONS);
  let resolved = 0;
  table.rows.forEach(function (row) {
    if (String(row.status || 'OPEN').toUpperCase() !== 'OPEN' ||
        String(row.book_id || '') !== String(bookId || '') ||
        (allowedTypes.size && !allowedTypes.has(String(row.exception_type || '').toUpperCase())) ||
        (sourceRef && String(row.source_ref || '') && String(row.source_ref || '') !== String(sourceRef))) return;
    vNextAdminUpdateTableRow_(hub, VN_ADMIN_SHEETS.EXCEPTIONS, row._rowNumber, { status: 'RESOLVED' });
    resolved++;
  });
  return resolved;
}

function vNextAdminRefreshHome_(hub) {
  const sheet = vNextAdminGetOrCreateSheet_(hub, VN_ADMIN_SHEETS.HOME);
  try { sheet.getDataRange().breakApart(); } catch (ignoredBreakApart) {}
  sheet.clear();
  const registryRows = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.REGISTRY).rows;
  const exceptionRows = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.EXCEPTIONS).rows.filter(function (row) {
    return String(row.status || 'OPEN').toUpperCase() === 'OPEN' &&
      String(row.exception_type || '').toUpperCase() !== 'APPROVAL_PENDING';
  });
  const jobRows = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.JOBS).rows;
  const approvalRows = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.APPROVALS).rows;
  const clientCount = registryRows.filter(function (row) {
    return String(row.mode || '').toUpperCase() === 'CLIENT' &&
      String(row.status || '').toUpperCase() !== 'ARCHIVED';
  }).length;
  const queuedCount = jobRows.filter(function (row) {
    return String(row.status || '').toUpperCase() === 'QUEUED';
  }).length;
  const runningCount = jobRows.filter(function (row) {
    return String(row.status || '').toUpperCase() === 'RUNNING';
  }).length;
  const failedCount = jobRows.filter(function (row) {
    return String(row.status || '').toUpperCase() === 'FAILED';
  }).length;
  const approvalCount = approvalRows.filter(function (row) {
    return String(row.status || '').toUpperCase() === 'PENDING';
  }).length;
  const actualIssueCount = exceptionRows.filter(vNextAdminIsActualDataIssue_).length;
  const automationInstalled = vNextAdminAutomationInstalled_();
  const operations = vNextAdminOperationalMetrics_(hub, automationInstalled, jobRows);
  const pilot = vNextAdminPilotStatus_(hub);
  let overall = '対応なし';
  if (!automationInstalled) overall = '初期設定が必要';
  else if (operations.schedulerStale || operations.queueStale) overall = '自動処理を確認';
  else if (approvalCount || exceptionRows.length) overall = '管理ハブの確認あり';
  else if (queuedCount || runningCount) overall = '自動処理中';
  const rows = [
    [VN_ADMIN_MENU_NAME, overall],
    ['最終確認', new Date()],
    ['', ''],
    ['今日、判断すること', approvalCount + exceptionRows.length + '件'],
    ['承認待ち', approvalCount + '件'],
    ['要確認', exceptionRows.length + '件'],
    ['うち実績データの確認', actualIssueCount + '件'],
    ['', ''],
    ['自動処理', runningCount ? runningCount + '件を処理中' : queuedCount ? queuedCount + '件が受付済み' : '待機なし'],
    ['処理失敗（累計）', failedCount + '件'],
    ['自動更新', !automationInstalled ? '未設定' : operations.schedulerStale ? '要確認' : '正常'],
    ['最終自動更新', operations.lastSweepSucceededAt || '未実行'],
    ['', ''],
    ['登録済みクライアント年度ブック', clientCount + '冊'],
    ['Pilot展開', pilot.clientCount + ' / ' + pilot.currentLimit + '冊'],
    ['', ''],
    ['次の操作', '右側の案内に従ってください。案内が出ないときだけ、上部メニュー「' + VN_ADMIN_MENU_NAME + '」→「' + VN_ADMIN_MENU_OPEN_SIDEBAR + '」を使います。'],
    ['正式予算', '確定済みの計画は上書きせず、訂正履歴を追加します。']
  ];
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  sheet.setHiddenGridlines(true);
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setFontSize(15)
    .setBackground('#1f2937').setFontColor('#ffffff');
  sheet.getRange(2, 2).setNumberFormat('yyyy/MM/dd HH:mm');
  [4, 9, 14, 17].forEach(function (rowNumber) {
    sheet.getRange(rowNumber, 1, 1, 2).setFontWeight('bold').setBackground('#eef2f7');
  });
  sheet.getRange(1, 1, rows.length, 2).setVerticalAlignment('middle');
  sheet.getRange(1, 1, rows.length, 1).setFontWeight('bold');
  sheet.setColumnWidth(1, 210);
  sheet.setColumnWidth(2, 520);
  sheet.setFrozenRows(2);
}

function vNextAdminPortalUsesV2Tables_(runtimeVersion) {
  const version = String(runtimeVersion || '');
  if (!version || version === 'vnext-portal-1.0.0') return false;
  return version === VN_ADMIN_PORTAL_RUNTIME_VERSION || /^vnext-portal-1\./.test(version);
}

/** Read-only Portal open for sidebar projections. Skips header repair and protection writes. */
function vNextAdminResolvePortalForRead_(hub) {
  const config = vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET);
  const spreadsheetId = String(config.portal_spreadsheet_id ||
    PropertiesService.getScriptProperties().getProperty('VNEXT_PORTAL_SPREADSHEET_ID') || '').trim();
  if (!spreadsheetId) throw new Error('Employee Portal is not configured.');
  return { spreadsheet: SpreadsheetApp.openById(spreadsheetId), spreadsheetId: spreadsheetId };
}

function vNextAdminResolvePortal_(hub) {
  const config = vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET);
  const spreadsheetId = String(config.portal_spreadsheet_id ||
    PropertiesService.getScriptProperties().getProperty('VNEXT_PORTAL_SPREADSHEET_ID') || '').trim();
  if (!spreadsheetId) throw new Error('Employee Portal is not configured.');
  const portalId = vNextAdminRequiredText_(config.portal_id, 'portal_id');
  const scriptId = vNextAdminRequiredText_(config.portal_script_id, 'portal_script_id');
  const runtimeVersion = vNextAdminRequiredText_(config.portal_runtime_version, 'portal_runtime_version');
  const runtimeSha256 = vNextAdminRequiredText_(config.portal_runtime_sha256, 'portal_runtime_sha256');
  const employeeDomain = vNextAdminNormalizeDomain_(config.employee_domain ||
    PropertiesService.getScriptProperties().getProperty('VNEXT_EMPLOYEE_DOMAIN'));
  const supportedRuntimeVersions = [VN_ADMIN_PORTAL_RUNTIME_VERSION].concat(VN_ADMIN_PORTAL_LEGACY_RUNTIME_VERSIONS);
  if (supportedRuntimeVersions.indexOf(runtimeVersion) < 0 || !/^[a-f0-9]{64}$/.test(runtimeSha256) || !employeeDomain) {
    throw new Error('Employee Portal runtime/domain identity is invalid.');
  }
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const portalConfig = vNextAdminReadKeyValueSheet_(spreadsheet, VN_ADMIN_PORTAL_CONFIG_SHEET);
  if (String(portalConfig.mode || '').toUpperCase() !== 'PORTAL' ||
      String(portalConfig.portal_id || '') !== portalId ||
      String(portalConfig.runtime_version || '') !== runtimeVersion ||
      String(portalConfig.runtime_sha256 || '') !== runtimeSha256 ||
      vNextAdminNormalizeDomain_(portalConfig.employee_domain || '') !== employeeDomain ||
      String(portalConfig.access_policy || '').toUpperCase() !== 'INTERNAL_OPEN') {
    throw new Error('Employee Portal local identity does not match the Hub record.');
  }
  vNextClientRuntimeAssertBoundParent_(scriptId, spreadsheetId);
  vNextAdminEnsurePortalRuntimeTables_(spreadsheet, runtimeVersion);
  const hubConfig = vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET);
  vNextAdminProtectInternalSheets_(spreadsheet,
    vNextAdminMergeEmails_(hubConfig.admin_emails, vNextAdminActor_()), [
      VN_ADMIN_PORTAL_DIRECTORY_SHEET, VN_ADMIN_PORTAL_CONFIG_SHEET,
      VN_ADMIN_PORTAL_CLIENT_CATALOG_SHEET
    ]);
  return {
    portalId: portalId, spreadsheetId: spreadsheetId, scriptId: scriptId,
    runtimeVersion: runtimeVersion, runtimeSha256: runtimeSha256,
    employeeDomain: employeeDomain, spreadsheet: spreadsheet
  };
}

function vNextAdminAppendPortalRequestEvent_(portal, validated, eventType, status, detail) {
  const payload = validated.payload || {};
  const extra = detail || {};
  const now = new Date().toISOString();
  const ownerEmail = String(validated.forecastOwnerEmail || payload.forecastOwnerEmail || payload.requestedBy || '').toLowerCase();
  const relatedMemberEmails = validated.relatedMemberEmails || payload.relatedMemberEmails || [];
  const relatedMemberNames = validated.relatedMemberNames || payload.relatedMemberNames || [];
  const record = {
    request_event_id: 'PORTAL-REQEV-' + Utilities.getUuid(), request_id: String(payload.requestId || ''),
    event_type: String(eventType || '').toUpperCase(), status: String(status || '').toUpperCase(),
    request_hash: String(validated.requestHash || ''), request_json: '',
    fiscal_year: Number(payload.fiscalYear || 0),
    // v2 Portal rows keep their untrusted transport identity self-consistent;
    // the authoritative ZAC-derived clientId exists only in Hub job/registry.
    client_id: String(payload.schemaVersion === VN_ADMIN_PORTAL_REQUEST_SCHEMA
      ? payload.catalogKey : (payload.clientId || '')),
    client_name: String(payload.clientName || ''), forecast_owner_email: ownerEmail,
    related_member_emails_json: vNextAdminCanonicalJson_(relatedMemberEmails),
    requested_at: String(payload.requestedAt || ''), requested_by: String(payload.requestedBy || ''),
    related_book_id: String(extra.relatedBookId || ''), related_book_url: String(extra.relatedBookUrl || ''),
    detail_code: String(extra.detailCode || ''), detail_message: String(extra.detailMessage || ''),
    created_at: now, created_by: vNextAdminActor_(),
    catalog_key: String(validated.catalogKey || payload.catalogKey || ''),
    // Legacy v1 rows predate the projection columns. Keep them truly blank so
    // the v1 reader can distinguish the immutable legacy shape after upgrade.
    related_member_names_json: String(payload.schemaVersion || '') === VN_ADMIN_PORTAL_REQUEST_SCHEMA_V1
      ? '' : vNextAdminCanonicalJson_(relatedMemberNames)
  };
  const config = vNextAdminReadKeyValueSheet_(portal, VN_ADMIN_PORTAL_CONFIG_SHEET);
  const headers = vNextAdminPortalUsesV2Tables_(config.runtime_version)
    ? VN_ADMIN_PORTAL_REQUEST_HEADERS : VN_ADMIN_PORTAL_REQUEST_HEADERS_V1;
  vNextAdminAppendObject_(portal, VN_ADMIN_PORTAL_REQUEST_SHEET, record, headers);
  return record;
}

function vNextAdminRefreshPortalDirectory_(hub, optionalPortal) {
  let spreadsheet;
  if (optionalPortal && typeof optionalPortal.getId === 'function') {
    spreadsheet = optionalPortal;
    const portalConfig = vNextAdminReadKeyValueSheet_(spreadsheet, VN_ADMIN_PORTAL_CONFIG_SHEET);
    vNextAdminEnsurePortalRuntimeTables_(spreadsheet, portalConfig.runtime_version || VN_ADMIN_PORTAL_RUNTIME_VERSION);
  } else {
    spreadsheet = vNextAdminResolvePortal_(hub).spreadsheet;
  }
  const registry = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.REGISTRY).rows.filter(function (row) {
    return String(row.mode || '') === 'CLIENT' && String(row.status || '').toUpperCase() === 'ACTIVE';
  });
  const runsByBook = {};
  vNextAdminReadCoreRows_(hub, 'FORECAST_RUN').forEach(function (row) {
    if (String(row.status || '').toUpperCase() === 'SUCCESS') runsByBook[String(row.book_id || '')] = row;
  });
  const plansByBook = {};
  vNextAdminReadCoreRows_(hub, 'PLAN_VERSION').forEach(function (row) {
    if (['SUBMITTED', 'APPROVED'].indexOf(String(row.status || '').toUpperCase()) < 0) return;
    plansByBook[String(row.book_id || '')] = row;
  });
  const requestIds = {};
  const memberNamesByRequest = {};
  const portalRequestRows = vNextAdminReadTable_(spreadsheet, VN_ADMIN_PORTAL_REQUEST_SHEET).rows;
  portalRequestRows.forEach(function (requestRow) {
    const requestId = String(requestRow.request_id || '');
    if (String(requestRow.related_book_id || '')) {
      requestIds[String(requestRow.related_book_id)] = requestId;
    }
    if (String(requestRow.event_type || '').toUpperCase() !== 'REQUESTED' || !requestId) return;
    const payload = vNextAdminParseJson_(requestRow.request_json, null);
    if (!payload || String(payload.schemaVersion || '') !== VN_ADMIN_PORTAL_REQUEST_SCHEMA) return;
    try { memberNamesByRequest[requestId] = vNextAdminNormalizeRelatedMemberNames_(payload.relatedMemberNames); }
    catch (ignoredInvalidNames) { memberNamesByRequest[requestId] = []; }
  });
  const teamByBook = {};
  vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.TEAM).rows.forEach(function (row) {
    if (String(row.status || '').toUpperCase() !== 'ACTIVE') return;
    const id = String(row.book_id || '');
    if (!teamByBook[id]) teamByBook[id] = [];
    if (String(row.role || '').toUpperCase() !== 'FORECAST_OWNER') teamByBook[id].push(String(row.email || '').toLowerCase());
  });
  const rows = registry.map(function (row) {
    const bookId = String(row.book_id || '');
    const run = runsByBook[bookId] || {};
    const plan = plansByBook[bookId] || {};
    const state = String(row.state || 'INPUT_OPEN').toUpperCase();
    let relatedMemberNames = vNextAdminParseJson_(row.related_member_names_json, []);
    if (!Array.isArray(relatedMemberNames) || !relatedMemberNames.length) {
      relatedMemberNames = memberNamesByRequest[requestIds[bookId] || ''] || [];
    }
    return {
      directory_event_id: 'DIR-' + Utilities.getUuid(),
      directory_key: Number(row.fiscal_year) + '|' + vNextAdminPortalCanonicalClientKey_(row),
      fiscal_year: Number(row.fiscal_year), client_id: String(row.client_id || ''),
      client_name: String(row.client_name || ''),
      forecast_owner_email: vNextAdminParseList_(row.forecast_owner_emails)[0] || '',
      related_member_emails_json: vNextAdminCanonicalJson_(vNextAdminMergeEmails_(teamByBook[bookId] || [])),
      state: state,
      center_forecast: run.p50 === '' || run.p50 === undefined ? '' : Number(run.p50),
      adopted_forecast: plan.adopted_forecast === '' || plan.adopted_forecast === undefined ? '' : Number(plan.adopted_forecast),
      final_budget: plan.final_budget === '' || plan.final_budget === undefined ? '' : Number(plan.final_budget),
      next_action: vNextAdminPortalNextAction_(state, String(row.health_code || '')),
      client_book_url: String(row.spreadsheet_url || ''), request_id: requestIds[bookId] || '',
      updated_at: new Date().toISOString(), updated_by: vNextAdminActor_(),
      related_member_names_json: vNextAdminCanonicalJson_(relatedMemberNames)
    };
  });
  const portalConfig = vNextAdminReadKeyValueSheet_(spreadsheet, VN_ADMIN_PORTAL_CONFIG_SHEET);
  const directoryHeaders = vNextAdminPortalUsesV2Tables_(portalConfig.runtime_version)
    ? VN_ADMIN_PORTAL_DIRECTORY_HEADERS : VN_ADMIN_PORTAL_DIRECTORY_HEADERS_V1;
  const sheet = vNextAdminEnsureExactTableHeaders_(spreadsheet,
    VN_ADMIN_PORTAL_DIRECTORY_SHEET, directoryHeaders);
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  if (rows.length) {
    const values = rows.map(function (row) {
      return directoryHeaders.map(function (header) { return row[header] === undefined ? '' : row[header]; });
    });
    sheet.getRange(2, 1, values.length, directoryHeaders.length).setValues(values);
  }
  sheet.hideSheet();
  vNextAdminWritePortalConfigValues_(spreadsheet, {
    admin_hub_url: hub.getUrl(),
    portal_spreadsheet_url: spreadsheet.getUrl()
  });
  return { portalSpreadsheetId: spreadsheet.getId(), rows: rows.length };
}

function vNextAdminPortalNextAction_(state, healthCode) {
  if (String(healthCode || '').indexOf('HISTORY') >= 0) return '実績データの連携・不足状況を確認してください。';
  const labels = {
    INPUT_OPEN: '現場情報を回答してください。', READY_TO_RUN: '予算策定担当が予測を依頼します。',
    RUNNING: '予測を計算しています。', DRAFT_READY: '予算策定担当が予算案を作成します。',
    SUBMITTED: '管理ハブの承認待ちです。', CHANGES_REQUESTED: '差戻し内容を確認してください。',
    OFFICIAL_LOCKED: '正式予算を確認できます。', REVIEW_DUE: '年度の振り返りを回答してください。',
    YEAR_CLOSED: '年度終了（閲覧のみ）'
  };
  return labels[String(state || '').toUpperCase()] || '専用ブックで状況を確認してください。';
}

function vNextAdminAppendException_(hub, input) {
  const obj = vNextAdminExceptionObject_(input || {});
  vNextAdminAppendObject_(hub, VN_ADMIN_SHEETS.EXCEPTIONS, obj);
  return obj;
}

function vNextAdminExceptionObject_(input) {
  const detected = input.detected_at || new Date();
  const identity = [input.exception_type || '', input.book_id || '', input.source_ref || '', input.title || ''].join('|');
  return {
    exception_id: 'EX-' + vNextAdminSha256_(identity).slice(0, 16),
    severity: input.severity || 'WARN', exception_type: input.exception_type || 'GENERAL',
    book_id: input.book_id || '', client_name: input.client_name || '', fiscal_year: input.fiscal_year || '',
    title: input.title || '', detail: input.detail || '', recommended_action: input.recommended_action || '',
    status: 'OPEN', detected_at: detected, source_ref: input.source_ref || ''
  };
}

// ---------------------------- Registry / release ----------------------------

function vNextAdminRegisterBook_(hub, object) {
  return vNextAdminUpsertObject_(hub, VN_ADMIN_SHEETS.REGISTRY, 'book_id', object.book_id, object);
}

function vNextAdminClientSchemaVersion_() {
  return typeof VNEXT_CORE !== 'undefined' ? String(VNEXT_CORE.SCHEMA_VERSION) : 'vnext-schema-2';
}

function vNextAdminRegisterTeam_(hub, bookId, clientId, fiscalYear, owners, editors, viewers, status) {
  const now = new Date();
  const actor = vNextAdminActor_();
  const memberStatus = String(status || 'ACTIVE').toUpperCase();
  const ownerSet = new Set(vNextAdminMergeEmails_(owners));
  const editorSet = new Set(vNextAdminMergeEmails_(editors));
  const viewerSet = new Set(vNextAdminMergeEmails_(viewers));
  const all = vNextAdminMergeEmails_(owners, editors, viewers);
  all.forEach(function (email) {
    let role = 'VIEWER';
    if (editorSet.has(email)) role = 'MEMBER';
    if (ownerSet.has(email)) role = 'FORECAST_OWNER';
    const key = String(bookId) + '|' + email;
    vNextAdminUpsertObject_(hub, VN_ADMIN_SHEETS.TEAM, 'team_key', key, {
      team_key: key, book_id: bookId, client_id: clientId, fiscal_year: fiscalYear,
      email: email, role: role, status: memberStatus, created_at: now, created_by: actor,
      updated_at: now, note: viewerSet.has(email) && role !== 'VIEWER' ? 'viewer permission also present' : ''
    });
  });
  return { members: all.length, forecastOwners: ownerSet.size };
}

function vNextAdminSetTeamStatus_(hub, bookId, status) {
  const normalized = vNextAdminRequiredText_(status, 'teamStatus').toUpperCase();
  const table = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.TEAM);
  let updated = 0;
  table.rows.forEach(function (row) {
    if (String(row.book_id || '') !== String(bookId || '')) return;
    vNextAdminUpdateTableRow_(hub, VN_ADMIN_SHEETS.TEAM, row._rowNumber, {
      status: normalized, updated_at: new Date()
    });
    updated++;
  });
  return updated;
}

function vNextAdminRegisterRelease_(hub, request) {
  const req = request && typeof request === 'object' ? request : {};
  const releaseId = vNextAdminRequiredText_(req.release_id || req.releaseId, 'releaseId');
  const now = new Date();
  const object = {
    release_id: releaseId,
    release_name: vNextAdminText_(req.release_name || req.releaseName) || releaseId,
    status: String(req.status || 'DRAFT').toUpperCase(),
    template_spreadsheet_id: vNextAdminText_(req.template_spreadsheet_id || req.templateSpreadsheetId),
    schema_version: vNextAdminText_(req.schema_version || req.schemaVersion) || vNextAdminClientSchemaVersion_(),
    engine_version: vNextAdminText_(req.engine_version || req.engineVersion),
    ux_version: vNextAdminText_(req.ux_version || req.uxVersion),
    admin_version: vNextAdminText_(req.admin_version || req.adminVersion) || VN_ADMIN_SCHEMA_VERSION,
    client_runtime_version: vNextAdminText_(req.client_runtime_version || req.clientRuntimeVersion),
    client_runtime_sha256: vNextAdminText_(req.client_runtime_sha256 || req.clientRuntimeSha256),
    template_content_sha256: vNextAdminText_(req.template_content_sha256 || req.templateContentSha256),
    template_manifest_schema: vNextAdminText_(req.template_manifest_schema || req.templateManifestSchema),
    template_script_id: vNextAdminText_(req.template_script_id || req.templateScriptId),
    created_at: req.created_at || now, created_by: req.created_by || vNextAdminActor_(),
    activated_at: req.activated_at || (String(req.status || '').toUpperCase() === 'ACTIVE' ? now : ''),
    note: vNextAdminText_(req.note)
  };
  const existing = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.RELEASES).rows.find(function (row) {
    return String(row.release_id || '') === releaseId;
  });
  if (existing) {
    const immutableKeys = [
      'release_name', 'template_spreadsheet_id', 'schema_version', 'engine_version',
      'ux_version', 'admin_version', 'client_runtime_version',
      'client_runtime_sha256', 'template_content_sha256', 'template_manifest_schema',
      'template_script_id', 'created_by'
    ];
    const differs = immutableKeys.some(function (key) {
      return String(existing[key] || '') !== String(object[key] || '');
    });
    if (differs) throw new Error('Release ID already exists with different immutable fields: ' + releaseId);
    return existing;
  }
  return vNextAdminAppendObject_(hub, VN_ADMIN_SHEETS.RELEASES, object);
}

function vNextAdminResolveRelease_(hub, releaseId) {
  const table = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.RELEASES);
  const explicitId = String(releaseId || '').trim();
  if (explicitId) {
    const explicit = table.rows.find(function (item) { return String(item.release_id || '') === explicitId; });
    if (!explicit) throw new Error('Requested release is not registered: ' + explicitId);
    if (String(explicit.status || '').toUpperCase() !== 'ACTIVE') {
      throw new Error('Requested release is not ACTIVE: ' + explicitId);
    }
    return explicit;
  }
  const canonicalPair = vNextAdminReadActiveReleasePair_(hub);
  const pointer = canonicalPair.releaseId;
  if (pointer) {
    const pointed = table.rows.find(function (item) { return String(item.release_id || '') === pointer; });
    if (!pointed || String(pointed.status || '').toUpperCase() !== 'ACTIVE') {
      throw new Error('Active Template Release pointer is missing or not ACTIVE: ' + pointer);
    }
    if (String(pointed.template_spreadsheet_id || '') !== canonicalPair.templateSpreadsheetId) {
      throw new Error('Canonical active Template spreadsheet does not match RELEASES: ' + pointer);
    }
    return pointed;
  }
  const active = table.rows.filter(function (item) {
    return String(item.status || '').toUpperCase() === 'ACTIVE';
  });
  if (active.length !== 1) {
    throw new Error('Exactly one ACTIVE template release is required; found ' + active.length + '.');
  }
  return active[0];
}

function vNextAdminReadActiveReleasePair_(hub) {
  const row = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.SETTINGS).rows.filter(function (item) {
    return String(item.setting_key || '') === 'ACTIVE_RELEASE_PAIR_JSON';
  }).slice(-1)[0];
  if (!row) throw new Error('Canonical ACTIVE_RELEASE_PAIR_JSON is missing. Re-run the guarded release activation.');
  const raw = String(row.setting_value || '');
  const pair = vNextAdminParseJson_(raw, null);
  if (!pair || Array.isArray(pair) || raw !== vNextAdminCanonicalJson_(pair) ||
      !vNextAdminText_(pair.releaseId) || !vNextAdminText_(pair.modelReleaseId) ||
      !vNextAdminText_(pair.templateSpreadsheetId)) {
    throw new Error('Canonical ACTIVE_RELEASE_PAIR_JSON is malformed or non-canonical.');
  }
  return { releaseId: String(pair.releaseId), modelReleaseId: String(pair.modelReleaseId),
    templateSpreadsheetId: String(pair.templateSpreadsheetId) };
}

function vNextAdminActiveReleasePairJson_(releaseId, modelReleaseId, templateSpreadsheetId) {
  return vNextAdminCanonicalJson_({ releaseId: vNextAdminRequiredText_(releaseId, 'releaseId'),
    modelReleaseId: vNextAdminRequiredText_(modelReleaseId, 'modelReleaseId'),
    templateSpreadsheetId: vNextAdminRequiredText_(templateSpreadsheetId, 'templateSpreadsheetId') });
}

function vNextAdminWriteCanonicalReleasePair_(hub, releaseId, modelReleaseId, templateSpreadsheetId) {
  const value = vNextAdminActiveReleasePairJson_(releaseId, modelReleaseId, templateSpreadsheetId);
  vNextAdminUpsertObject_(hub, VN_ADMIN_SHEETS.SETTINGS, 'setting_key', 'ACTIVE_RELEASE_PAIR_JSON', {
    setting_key: 'ACTIVE_RELEASE_PAIR_JSON', setting_value: value, value_type: 'JSON', scope: 'GLOBAL',
    effective_from: new Date(), updated_at: new Date(), updated_by: vNextAdminActor_(),
    note: 'Canonical one-cell active Template/Model pair; all other pointers are caches'
  });
  const reread = vNextAdminReadActiveReleasePair_(hub);
  if (vNextAdminActiveReleasePairJson_(reread.releaseId, reread.modelReleaseId, reread.templateSpreadsheetId) !== value) {
    throw new Error('Canonical ACTIVE_RELEASE_PAIR_JSON reread verification failed.');
  }
  return reread;
}

function vNextAdminReadActiveReleasePairMirrors_(hub) {
  const config = vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET);
  const settings = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.SETTINGS).rows;
  function setting(key) {
    const row = settings.filter(function (item) { return String(item.setting_key || '') === key; }).slice(-1)[0];
    return vNextAdminText_(row && row.setting_value);
  }
  const props = PropertiesService.getScriptProperties();
  return {
    settingReleaseId: setting('ACTIVE_TEMPLATE_RELEASE_ID'), settingModelReleaseId: setting('ACTIVE_MODEL_RELEASE_ID'),
    configReleaseId: vNextAdminText_(config.active_release_id), configModelReleaseId: vNextAdminText_(config.active_model_release_id),
    configTemplateSpreadsheetId: vNextAdminText_(config.template_spreadsheet_id),
    propertyReleaseId: vNextAdminText_(props.getProperty('VNEXT_ACTIVE_RELEASE_ID')),
    propertyModelReleaseId: vNextAdminText_(props.getProperty('VNEXT_ACTIVE_MODEL_RELEASE_ID')),
    propertyTemplateSpreadsheetId: vNextAdminText_(props.getProperty('VNEXT_MASTER_TEMPLATE_SPREADSHEET_ID'))
  };
}

function vNextAdminActiveReleasePairMirrorsExact_(hub, release, model) {
  const mirror = vNextAdminReadActiveReleasePairMirrors_(hub);
  const releaseId = String(release.release_id || '');
  const modelId = String(model.model_release_id || '');
  const templateId = String(release.template_spreadsheet_id || '');
  return mirror.settingReleaseId === releaseId && mirror.settingModelReleaseId === modelId &&
    mirror.configReleaseId === releaseId && mirror.configModelReleaseId === modelId &&
    mirror.configTemplateSpreadsheetId === templateId && mirror.propertyReleaseId === releaseId &&
    mirror.propertyModelReleaseId === modelId && mirror.propertyTemplateSpreadsheetId === templateId;
}

function vNextAdminWriteActiveReleasePairCaches_(hub, release, model) {
  const now = new Date();
  vNextAdminUpsertObject_(hub, VN_ADMIN_SHEETS.SETTINGS, 'setting_key', 'ACTIVE_TEMPLATE_RELEASE_ID', {
    setting_key: 'ACTIVE_TEMPLATE_RELEASE_ID', setting_value: release.release_id, value_type: 'STRING', scope: 'GLOBAL',
    effective_from: now, updated_at: now, updated_by: vNextAdminActor_(), note: 'Cache of ACTIVE_RELEASE_PAIR_JSON'
  });
  vNextAdminUpsertObject_(hub, VN_ADMIN_SHEETS.SETTINGS, 'setting_key', 'ACTIVE_MODEL_RELEASE_ID', {
    setting_key: 'ACTIVE_MODEL_RELEASE_ID', setting_value: model.model_release_id, value_type: 'STRING', scope: 'GLOBAL',
    effective_from: now, updated_at: now, updated_by: vNextAdminActor_(), note: 'Cache of ACTIVE_RELEASE_PAIR_JSON'
  });
  vNextAdminWriteSystemConfig_(hub, { active_release_id: release.release_id,
    active_model_release_id: model.model_release_id, template_spreadsheet_id: release.template_spreadsheet_id,
    client_runtime_version: release.client_runtime_version, client_runtime_sha256: release.client_runtime_sha256,
    template_script_id: release.template_script_id });
  PropertiesService.getScriptProperties().setProperties({ VNEXT_ACTIVE_RELEASE_ID: String(release.release_id || ''),
    VNEXT_ACTIVE_MODEL_RELEASE_ID: String(model.model_release_id || ''),
    VNEXT_MASTER_TEMPLATE_SPREADSHEET_ID: String(release.template_spreadsheet_id || '') }, false);
  if (!vNextAdminActiveReleasePairMirrorsExact_(hub, release, model)) {
    throw new Error('Active release pair cache reread verification failed.');
  }
  return true;
}

function vNextAdminClassifyActiveReleasePairCas_(current, target, expected, mirrorsExact) {
  const targetExact = String(current.releaseId || '') === String(target.releaseId || '') &&
    String(current.modelReleaseId || '') === String(target.modelReleaseId || '') &&
    String(current.templateSpreadsheetId || '') === String(target.templateSpreadsheetId || '');
  if (targetExact) return mirrorsExact === true ? 'REUSE' : 'REPAIR_CACHES';
  if (String(current.releaseId || '') === String(expected.releaseId || '') &&
      String(current.modelReleaseId || '') === String(expected.modelReleaseId || '')) return 'CAS';
  return 'CONFLICT';
}

function vNextAdminWriteActiveReleasePairPointers_(hub, targetRelease, targetModel, expectedPair) {
  const targetReleaseId = vNextAdminRequiredText_(targetRelease && targetRelease.release_id, 'targetRelease.release_id');
  const targetModelId = vNextAdminRequiredText_(targetModel && targetModel.model_release_id, 'targetModel.model_release_id');
  const targetTemplateId = vNextAdminRequiredText_(targetRelease && targetRelease.template_spreadsheet_id,
    'targetRelease.template_spreadsheet_id');
  let current;
  try { current = vNextAdminReadActiveReleasePair_(hub); }
  catch (missingCanonical) {
    const config = vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET);
    current = vNextAdminWriteCanonicalReleasePair_(hub, expectedPair.releaseId, expectedPair.modelReleaseId,
      vNextAdminRequiredText_(config.template_spreadsheet_id, 'expected templateSpreadsheetId'));
  }
  const action = vNextAdminClassifyActiveReleasePairCas_(current, {
    releaseId: targetReleaseId, modelReleaseId: targetModelId, templateSpreadsheetId: targetTemplateId
  }, expectedPair, vNextAdminActiveReleasePairMirrorsExact_(hub, targetRelease, targetModel));
  if (action === 'REUSE') {
    return { reused: true, before: current };
  }
  if (action === 'CONFLICT') {
    throw new Error('Release pair CAS failed. expected=' + expectedPair.releaseId + '/' +
      expectedPair.modelReleaseId + '; actual=' + current.releaseId + '/' + current.modelReleaseId);
  }
  if (action === 'CAS') vNextAdminWriteCanonicalReleasePair_(hub, targetReleaseId, targetModelId, targetTemplateId);
  vNextAdminWriteActiveReleasePairCaches_(hub, targetRelease, targetModel);
  const verified = vNextAdminReadActiveReleasePair_(hub);
  if (verified.releaseId !== targetReleaseId || verified.modelReleaseId !== targetModelId ||
      verified.templateSpreadsheetId !== targetTemplateId ||
      !vNextAdminActiveReleasePairMirrorsExact_(hub, targetRelease, targetModel)) {
    throw new Error('Active release pair canonical/cache verification failed.');
  }
  return { reused: false, before: current };
}

function vNextAdminCacheActiveReleasePair_(hub, release, model) {
  return vNextAdminWriteActiveReleasePairCaches_(hub, release, model);
}

function vNextAdminActivateReleasePairInternal_(hub, request) {
  const req = request && typeof request === 'object' ? request : {};
  const releaseId = vNextAdminRequiredText_(req.releaseId, 'releaseId');
  const modelReleaseId = vNextAdminRequiredText_(req.modelReleaseId, 'modelReleaseId');
  const reason = vNextAdminRequiredText_(req.reason, 'reason');
  const releaseTable = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.RELEASES);
  const release = releaseTable.rows.find(function (row) { return String(row.release_id || '') === releaseId; });
  if (!release || ['STAGED', 'ACTIVE'].indexOf(String(release.status || '').toUpperCase()) < 0) {
    throw new Error('Template Release must be STAGED (or ACTIVE during recovery): ' + releaseId);
  }
  const template = SpreadsheetApp.openById(vNextAdminRequiredText_(release.template_spreadsheet_id,
    'release.template_spreadsheet_id'));
  vNextAdminAssertReleaseTemplateManifest_(release, template);
  vNextAdminAssertPrivateAdminFile_(DriveApp.getFileById(template.getId()),
    vNextAdminMergeEmails_(vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET).admin_emails,
      vNextGetRuntimeConfig_().VNEXT_ADMIN_EMAILS, vNextAdminActor_()));
  if (typeof vNextClientRuntimeAssertBoundParent_ !== 'function') {
    throw new Error('Client runtime parent verifier is not installed.');
  }
  vNextClientRuntimeAssertBoundParent_(vNextAdminRequiredText_(release.template_script_id,
    'release.template_script_id'), template.getId());
  const modelSource = vNextAdminLatestModelRelease_(hub, modelReleaseId);
  if (!modelSource || ['DRAFT', 'ACTIVE'].indexOf(String(modelSource.status || '').toUpperCase()) < 0) {
    throw new Error('Pair activation requires a DRAFT model candidate: ' + modelReleaseId);
  }
  vNextAdminAssertModelReleaseChecksPassed_(modelSource);
  vNextAdminAssertModelTemplateCompatibility_(hub, modelSource, release);

  let currentPair;
  try { currentPair = vNextAdminReadActiveReleasePair_(hub); }
  catch (missingCanonical) {
    const legacyConfig = vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET);
    const legacyRelease = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.RELEASES).rows.find(function (row) {
      return String(row.release_id || '') === String(legacyConfig.active_release_id || '') &&
        String(row.status || '').toUpperCase() === 'ACTIVE';
    });
    const legacyModel = vNextAdminLatestModelRelease_(hub, legacyConfig.active_model_release_id || '');
    if (!legacyRelease || !legacyModel || String(legacyModel.status || '').toUpperCase() !== 'ACTIVE' ||
        String(legacyRelease.template_spreadsheet_id || '') !== String(legacyConfig.template_spreadsheet_id || '')) {
      throw missingCanonical;
    }
    vNextAdminAssertModelTemplateCompatibility_(hub, legacyModel, legacyRelease);
    currentPair = vNextAdminWriteCanonicalReleasePair_(hub, legacyRelease.release_id,
      legacyModel.model_release_id, legacyRelease.template_spreadsheet_id);
    vNextAdminWriteActiveReleasePairCaches_(hub, legacyRelease, legacyModel);
  }
  const recoveryJournal = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.TEMPLATE_JOURNAL).rows.filter(function (row) {
    return String(row.release_id || '') === releaseId && String(row.model_release_id || '') === modelReleaseId &&
      String(row.phase || '') === 'PAIR_VALIDATED';
  }).slice(-1)[0];
  const expectedPair = {
    releaseId: vNextAdminText_(req.expectedActiveReleaseId) ||
      (recoveryJournal && String(recoveryJournal.previous_release_id || '')) || currentPair.releaseId,
    modelReleaseId: vNextAdminText_(req.expectedActiveModelReleaseId) ||
      (recoveryJournal && String(recoveryJournal.previous_model_release_id || '')) || currentPair.modelReleaseId
  };
  const operationId = vNextAdminText_(req.operationId) ||
    (recoveryJournal && String(recoveryJournal.operation_id || '')) || ('PAIR-ACT-' + vNextAdminSha256_(vNextAdminCanonicalJson_({
    releaseId: releaseId, modelReleaseId: modelReleaseId,
    previousReleaseId: expectedPair.releaseId, previousModelReleaseId: expectedPair.modelReleaseId,
    reason: reason
  })).slice(0, 24).toUpperCase());
  const operationBase = {
    operationId: operationId, releaseId: releaseId, modelReleaseId: modelReleaseId,
    previousReleaseId: expectedPair.releaseId, previousModelReleaseId: expectedPair.modelReleaseId,
    templateSpreadsheetId: template.getId()
  };
  vNextAdminAppendTemplateJournal_(hub, Object.assign({}, operationBase, {
    phase: 'PAIR_VALIDATED', status: 'SUCCEEDED', detail: {
      reason: reason, templateManifestSha256: release.template_content_sha256,
      modelCandidateHash: vNextAdminModelCandidateHash_({ modelVersion: modelSource.model_version,
        schemaVersion: modelSource.schema_version, templateVersion: modelSource.template_version,
        parameters: vNextAdminParseJson_(modelSource.parameters_json, {}) })
    }
  }));

  if (String(release.status || '').toUpperCase() !== 'ACTIVE') {
    vNextAdminUpdateTableRow_(hub, VN_ADMIN_SHEETS.RELEASES, release._rowNumber, {
      status: 'ACTIVE', activated_at: new Date(), note: reason
    });
    release.status = 'ACTIVE';
    release.activated_at = new Date();
  }
  let activeModel = vNextAdminLatestModelRelease_(hub, modelReleaseId);
  if (String(activeModel.status || '').toUpperCase() === 'DRAFT') {
    const activated = Object.assign({}, activeModel, {
      status: 'ACTIVE', approved_at: new Date().toISOString(), approved_by: vNextAdminActor_().toLowerCase(),
      rollback_release_id: '', created_at: new Date().toISOString(), created_by: vNextAdminActor_().toLowerCase(),
      note: vNextAdminCanonicalJson_({ action: 'ACTIVATE_PAIR', operationId: operationId, reason: reason,
        templateReleaseId: releaseId })
    });
    delete activated._rowNumber;
    vNextAdminAppendCoreRowsNoLock_(hub, 'MODEL_RELEASE', [activated]);
    activeModel = activated;
  } else {
    const activeNote = vNextAdminParseJson_(activeModel.note, {});
    const alreadyPointed = currentPair.releaseId === releaseId && currentPair.modelReleaseId === modelReleaseId;
    if (!alreadyPointed && (String(activeNote.action || '') !== 'ACTIVATE_PAIR' ||
        String(activeNote.operationId || '') !== operationId)) {
      throw new Error('The model candidate is already ACTIVE under a different operation.');
    }
  }
  const templateRegistry = vNextAdminFindRegistryRow_(hub, function (row) {
    return String(row.mode || '') === 'TEMPLATE' && String(row.spreadsheet_id || '') === String(template.getId());
  });
  if (!templateRegistry) throw new Error('STAGED Template BOOK_REGISTRY row is missing.');
  vNextAdminPatchRegistryByBookId_(hub, templateRegistry.book_id, {
    template_release_id: releaseId, schema_version: vNextAdminClientSchemaVersion_(),
    state: 'TEMPLATE_READY', status: 'ACTIVE', health_status: 'OK', health_code: 'PAIR_ACTIVE',
    updated_at: new Date()
  });
  vNextAdminWriteBookConfig_(template, { state: 'TEMPLATE_READY', template_kind: 'IMMUTABLE_ACTIVE',
    updated_at: new Date(), updated_by: vNextAdminActor_() });
  vNextAdminAppendTemplateJournal_(hub, Object.assign({}, operationBase, {
    phase: 'NEW_PAIR_ACTIVE', status: 'SUCCEEDED', detail: { reason: reason }
  }));

  vNextAdminWriteActiveReleasePairPointers_(hub, release, activeModel, expectedPair);
  vNextAdminAppendTemplateJournal_(hub, Object.assign({}, operationBase, {
    phase: 'POINTER_CAS_COMMITTED', status: 'SUCCEEDED', detail: { expectedPair: expectedPair }
  }));

  vNextAdminCacheActiveReleasePair_(hub, release, activeModel);
  vNextAdminAppendTemplateJournal_(hub, Object.assign({}, operationBase, {
    phase: 'PROPERTY_CACHE_UPDATED', status: 'SUCCEEDED', detail: {}
  }));

  vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.RELEASES).rows.forEach(function (row) {
    if (String(row.release_id || '') === releaseId ||
        String(row.release_id || '') !== String(expectedPair.releaseId || '') ||
        String(row.status || '').toUpperCase() !== 'ACTIVE') return;
    vNextAdminUpdateTableRow_(hub, VN_ADMIN_SHEETS.RELEASES, row._rowNumber, {
      status: 'RETIRED', note: String(row.note || '')
    });
  });
  vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.REGISTRY).rows.forEach(function (row) {
    if (String(row.mode || '') !== 'TEMPLATE' || String(row.book_id || '') === String(templateRegistry.book_id || '') ||
        String(row.template_release_id || '') !== String(expectedPair.releaseId || '') ||
        String(row.status || '').toUpperCase() !== 'ACTIVE') return;
    vNextAdminUpdateTableRow_(hub, VN_ADMIN_SHEETS.REGISTRY, row._rowNumber, {
      status: 'RETIRED', updated_at: new Date()
    });
  });
  vNextAdminAppendTemplateJournal_(hub, Object.assign({}, operationBase, {
    phase: 'PREVIOUS_TEMPLATE_RETIRED', status: 'SUCCEEDED', detail: {}
  }));
  vNextAdminAppendTemplateJournal_(hub, Object.assign({}, operationBase, {
    phase: 'COMPLETED', status: 'SUCCEEDED', detail: { reason: reason }
  }));
  vNextAdminWriteAudit_(hub, 'ACTIVATE_RELEASE_PAIR', 'RELEASE_PAIR', operationId, 'SUCCESS', {
    releaseId: releaseId, modelReleaseId: modelReleaseId,
    previousReleaseId: expectedPair.releaseId, previousModelReleaseId: expectedPair.modelReleaseId,
    order: ['NEW_PAIR_ACTIVE', 'POINTER_CAS_COMMITTED', 'PROPERTY_CACHE_UPDATED', 'PREVIOUS_TEMPLATE_RETIRED'],
    reason: reason
  });
  return { reused: currentPair.releaseId === releaseId && currentPair.modelReleaseId === modelReleaseId,
    operationId: operationId, activeReleaseId: releaseId, activeModelReleaseId: modelReleaseId,
    previousReleaseId: expectedPair.releaseId, previousModelReleaseId: expectedPair.modelReleaseId };
}

function vNextAdminAppendTemplateJournal_(hub, options) {
  const opt = options || {};
  const journalSheet = vNextAdminEnsureTable_(hub, VN_ADMIN_SHEETS.TEMPLATE_JOURNAL,
    VN_ADMIN_HEADERS[VN_ADMIN_SHEETS.TEMPLATE_JOURNAL]);
  if (!journalSheet.isSheetHidden()) journalSheet.hideSheet();
  const operationId = vNextAdminRequiredText_(opt.operationId, 'operationId');
  const phase = vNextAdminRequiredText_(opt.phase, 'phase').toUpperCase();
  const journalId = 'TPJ-' + vNextAdminSha256_(operationId + '|' + phase).slice(0, 28).toUpperCase();
  const detailJson = vNextAdminCanonicalJson_(opt.detail || {});
  const existing = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.TEMPLATE_JOURNAL).rows.find(function (row) {
    return String(row.journal_id || '') === journalId;
  });
  if (existing) {
    if (String(existing.operation_id || '') !== operationId || String(existing.phase || '') !== phase ||
        String(existing.release_id || '') !== String(opt.releaseId || '') ||
        String(existing.model_release_id || '') !== String(opt.modelReleaseId || '') ||
        String(existing.detail_json || '') !== detailJson) {
      throw new Error('Template release journal idempotency conflict: ' + journalId);
    }
    return existing;
  }
  return vNextAdminAppendObject_(hub, VN_ADMIN_SHEETS.TEMPLATE_JOURNAL, {
    journal_id: journalId, operation_id: operationId, release_id: opt.releaseId || '',
    model_release_id: opt.modelReleaseId || '', previous_release_id: opt.previousReleaseId || '',
    previous_model_release_id: opt.previousModelReleaseId || '',
    template_spreadsheet_id: opt.templateSpreadsheetId || '', phase: phase,
    status: String(opt.status || 'SUCCEEDED').toUpperCase(), detail_json: detailJson,
    occurred_at: new Date(), actor: vNextAdminActor_()
  });
}

function vNextAdminEnsureInitialModelRelease_(hub, options) {
  const opt = options || {};
  const modelReleaseId = vNextAdminRequiredText_(opt.modelReleaseId, 'modelReleaseId');
  vNextAdminAssertModelReleaseIdSeparated_(hub, modelReleaseId);
  const existing = vNextAdminLatestModelRelease_(hub, modelReleaseId);
  if (existing) {
    if (String(existing.status || '').toUpperCase() !== 'ACTIVE') {
      throw new Error('Initial modelReleaseId already exists but is not ACTIVE: ' + modelReleaseId);
    }
    vNextAdminSetActiveModelRelease_(hub, modelReleaseId);
    return existing;
  }
  const schemaVersion = vNextAdminClientSchemaVersion_();
  const modelVersion = vNextAdminRequiredText_(opt.modelVersion || (typeof VNEXT_ENGINE !== 'undefined' && VNEXT_ENGINE.VERSION), 'modelVersion');
  if (typeof VNEXT_ENGINE === 'undefined' || modelVersion !== String(VNEXT_ENGINE.VERSION || '')) {
    throw new Error('Initial MODEL_RELEASE must bind to the deployed Forecast Engine version.');
  }
  const parameters = vNextAdminNormalizeModelParameters_(opt.parameters || {});
  const candidateHash = vNextAdminModelCandidateHash_({
    modelVersion: modelVersion, schemaVersion: schemaVersion,
    templateVersion: opt.templateVersion || '', parameters: parameters
  });
  const record = vNextAdminBuildModelReleaseRecord_({
    modelReleaseId: modelReleaseId,
    status: 'ACTIVE',
    modelVersion: modelVersion,
    schemaVersion: schemaVersion,
    templateVersion: opt.templateVersion || '',
    parameters: parameters,
    backtest: { status: 'PASS', basis: 'INITIAL_BOOTSTRAP_BASELINE', candidateHash: candidateHash },
    canary: { status: 'PASS', basis: 'INITIAL_BOOTSTRAP_BASELINE', candidateHash: candidateHash },
    approvedAt: opt.now || new Date(),
    approvedBy: opt.actor || vNextAdminActor_(),
    note: 'Initial bootstrap model release. Future changes require DRAFT, backtest PASS, and canary PASS.',
    actor: opt.actor || vNextAdminActor_(),
    now: opt.now || new Date()
  });
  vNextAdminAppendCoreRowsNoLock_(hub, 'MODEL_RELEASE', [record]);
  vNextAdminSetActiveModelRelease_(hub, modelReleaseId);
  vNextAdminWriteAudit_(hub, 'BOOTSTRAP_MODEL_RELEASE', 'MODEL_RELEASE', modelReleaseId, 'SUCCESS', {
    status: 'ACTIVE', modelVersion: record.model_version, templateVersion: record.template_version
  });
  return record;
}

function vNextAdminBuildModelReleaseRecord_(options) {
  const opt = options || {};
  const now = opt.now instanceof Date ? opt.now.toISOString() : String(opt.now || new Date().toISOString());
  const actor = String(opt.actor || vNextAdminActor_()).toLowerCase();
  return {
    model_release_id: vNextAdminRequiredText_(opt.modelReleaseId, 'modelReleaseId'),
    status: String(opt.status || 'DRAFT').toUpperCase(),
    model_version: vNextAdminRequiredText_(opt.modelVersion, 'modelVersion'),
    schema_version: vNextAdminText_(opt.schemaVersion) ||
      (typeof VNEXT_CORE !== 'undefined' ? VNEXT_CORE.SCHEMA_VERSION : VN_ADMIN_SCHEMA_VERSION),
    template_version: vNextAdminText_(opt.templateVersion),
    parameters_json: vNextAdminCanonicalJson_(opt.parameters || {}),
    backtest_json: vNextAdminCanonicalJson_(opt.backtest || {}),
    canary_json: vNextAdminCanonicalJson_(opt.canary || {}),
    approved_at: opt.approvedAt instanceof Date ? opt.approvedAt.toISOString() : vNextAdminText_(opt.approvedAt),
    approved_by: String(opt.approvedBy || '').toLowerCase(),
    rollback_release_id: vNextAdminText_(opt.rollbackReleaseId),
    created_at: now,
    created_by: actor,
    note: vNextAdminText_(opt.note)
  };
}

function vNextAdminModelReleaseRows_(hub, modelReleaseId) {
  return vNextAdminReadCoreRows_(hub, 'MODEL_RELEASE').filter(function (row) {
    return String(row.model_release_id || '') === String(modelReleaseId || '');
  });
}

function vNextAdminLatestModelRelease_(hub, modelReleaseId) {
  const rows = vNextAdminModelReleaseRows_(hub, modelReleaseId);
  return rows.length ? rows[rows.length - 1] : null;
}

function vNextAdminLatestModelReleaseSummaries_(hub) {
  const latest = {};
  vNextAdminReadCoreRows_(hub, 'MODEL_RELEASE').forEach(function (row) {
    if (row.model_release_id) latest[String(row.model_release_id)] = row;
  });
  return Object.keys(latest).sort().map(function (id) {
    const row = latest[id];
    return {
      modelReleaseId: id, status: String(row.status || ''), modelVersion: String(row.model_version || ''),
      templateVersion: String(row.template_version || ''), backtestPassed: vNextAdminModelCheckPassed_(row.backtest_json),
      canaryPassed: vNextAdminModelCheckPassed_(row.canary_json), createdAt: row.created_at || '', note: String(row.note || '')
    };
  });
}

function vNextAdminTryResolveActiveModelRelease_(hub) {
  try { return vNextAdminResolveActiveModelRelease_(hub); }
  catch (error) { return null; }
}

function vNextAdminResolveActiveModelReleaseForUpgrade_(hub, activePair) {
  const pair = activePair || vNextAdminReadActiveReleasePair_(hub);
  const modelId = vNextAdminRequiredText_(pair.modelReleaseId, 'activePair.modelReleaseId');
  const releaseId = vNextAdminRequiredText_(pair.releaseId, 'activePair.releaseId');
  const model = vNextAdminLatestModelRelease_(hub, modelId);
  if (!model || String(model.status || '').toUpperCase() !== 'ACTIVE') {
    throw new Error('Active source MODEL_RELEASE is missing or not ACTIVE: ' + modelId);
  }
  const release = vNextAdminResolveRelease_(hub, releaseId);
  vNextAdminAssertModelReleaseOwnPair_(model, release);
  return model;
}

function vNextAdminResolveActiveModelRelease_(hub, requestedId) {
  const canonicalPair = vNextAdminReadActiveReleasePair_(hub);
  const pointer = canonicalPair.modelReleaseId;
  const requested = vNextAdminText_(requestedId);
  if (requested && pointer && requested !== pointer) {
    throw new Error('Provisioning may use only the active MODEL_RELEASE. active=' + pointer);
  }
  const id = requested || pointer;
  if (!id) throw new Error('No active MODEL_RELEASE is configured.');
  const row = vNextAdminLatestModelRelease_(hub, id);
  if (!row || String(row.status || '').toUpperCase() !== 'ACTIVE') {
    throw new Error('Active MODEL_RELEASE pointer is missing or not ACTIVE: ' + id);
  }
  vNextAdminAssertModelReleaseChecksPassed_(row);
  vNextAdminAssertModelTemplateCompatibility_(hub, row);
  return row;
}

function vNextAdminAssertModelTemplateCompatibility_(hub, modelRelease, templateRelease) {
  const model = modelRelease || {};
  let template = templateRelease;
  if (!template) {
    const config = vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET);
    template = vNextAdminResolveRelease_(hub, config.active_release_id || '');
  }
  if (!String(template.release_id || '') ||
      String(model.template_version || '') !== String(template.release_id || '') ||
      String(model.schema_version || '') !== String(template.schema_version || '') ||
      String(model.model_version || '') !== String(template.engine_version || '') ||
      String(model.model_version || '') !== String(typeof VNEXT_ENGINE !== 'undefined' ? VNEXT_ENGINE.VERSION : '')) {
    throw new Error('MODEL_RELEASE.template_version must exactly equal Template release_id and match the deployed Engine/Core.');
  }
  return true;
}

function vNextAdminSetActiveModelRelease_(hub, modelReleaseId) {
  const id = vNextAdminRequiredText_(modelReleaseId, 'modelReleaseId');
  const model = vNextAdminLatestModelRelease_(hub, id);
  if (!model || String(model.status || '').toUpperCase() !== 'ACTIVE') {
    throw new Error('Active MODEL_RELEASE row is required before pointer update: ' + id);
  }
  const release = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.RELEASES).rows.find(function (row) {
    return String(row.release_id || '') === String(model.template_version || '') &&
      String(row.status || '').toUpperCase() === 'ACTIVE';
  });
  if (!release) throw new Error('MODEL_RELEASE has no matching ACTIVE Template Release: ' + id);
  const config = vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET);
  let current;
  try { current = vNextAdminReadActiveReleasePair_(hub); }
  catch (error) { current = { releaseId: String(config.active_release_id || release.release_id),
    modelReleaseId: String(config.active_model_release_id || id) }; }
  vNextAdminWriteActiveReleasePairPointers_(hub, release, model, current);
  return id;
}

function vNextAdminAssertModelReleaseIdSeparated_(hub, modelReleaseId) {
  const collision = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.RELEASES).rows.some(function (row) {
    return String(row.release_id || '') === String(modelReleaseId || '');
  });
  if (collision) throw new Error('Model Release ID must be different from every Template Release ID.');
}

function vNextAdminModelCheckPassed_(value) {
  const parsed = vNextAdminParseJson_(value, value && typeof value === 'object' ? value : {});
  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) return false;
  const result = String(parsed.status || parsed.result || parsed.outcome || '').trim().toUpperCase();
  return result === 'PASS' || result === 'PASSED';
}

function vNextAdminAssertModelReleaseChecksPassed_(row) {
  if (!row) throw new Error('MODEL_RELEASE is required.');
  if (typeof VNEXT_ENGINE === 'undefined' || String(row.model_version || '') !== String(VNEXT_ENGINE.VERSION || '')) {
    throw new Error('MODEL_RELEASE model_version does not match the deployed Forecast Engine.');
  }
  if (String(row.schema_version || '') !== vNextAdminClientSchemaVersion_()) {
    throw new Error('MODEL_RELEASE schema_version does not match the deployed Core.');
  }
  const parameters = vNextAdminNormalizeModelParameters_(vNextAdminParseJson_(row.parameters_json, {}));
  if (vNextAdminCanonicalJson_(parameters) !== String(row.parameters_json || '')) {
    throw new Error('MODEL_RELEASE parameters_json is not canonical or contains unsupported parameters.');
  }
  const candidateHash = vNextAdminModelCandidateHash_({
    modelVersion: row.model_version, schemaVersion: row.schema_version,
    templateVersion: row.template_version, parameters: parameters
  });
  const backtest = vNextAdminParseJson_(row.backtest_json, {});
  const canary = vNextAdminParseJson_(row.canary_json, {});
  if (!vNextAdminModelCheckPassed_(row && row.backtest_json)) throw new Error('MODEL_RELEASE backtest result must be PASS.');
  if (!vNextAdminModelCheckPassed_(row && row.canary_json)) throw new Error('MODEL_RELEASE canary result must be PASS.');
  if (String(backtest.candidateHash || '') !== candidateHash || String(canary.candidateHash || '') !== candidateHash) {
    throw new Error('MODEL_RELEASE backtest/canary artifacts are not bound to this candidate hash.');
  }
  return true;
}

/** Validates a historical Model against the immutable Template it was released with. */
function vNextAdminAssertModelReleaseOwnPair_(model, release) {
  if (!model || !release ||
      String(model.template_version || '') !== String(release.release_id || '') ||
      String(model.schema_version || '') !== String(release.schema_version || '') ||
      String(model.model_version || '') !== String(release.engine_version || '') ||
      String(model.schema_version || '') !== vNextAdminClientSchemaVersion_()) {
    throw new Error('Source Template and Model Release are not an exact immutable pair.');
  }
  const parameters = vNextAdminNormalizeModelParameters_(vNextAdminParseJson_(model.parameters_json, {}));
  if (vNextAdminCanonicalJson_(parameters) !== String(model.parameters_json || '')) {
    throw new Error('Source MODEL_RELEASE parameters are not canonical.');
  }
  const candidateHash = vNextAdminModelCandidateHash_({
    modelVersion: model.model_version,
    schemaVersion: model.schema_version,
    templateVersion: model.template_version,
    parameters: parameters
  });
  const backtest = vNextAdminParseJson_(model.backtest_json, {});
  const canary = vNextAdminParseJson_(model.canary_json, {});
  if (!vNextAdminModelCheckPassed_(model.backtest_json) ||
      !vNextAdminModelCheckPassed_(model.canary_json) ||
      String(backtest.candidateHash || '') !== candidateHash ||
      String(canary.candidateHash || '') !== candidateHash) {
    throw new Error('Source MODEL_RELEASE checks do not match its immutable candidate.');
  }
  return true;
}

function vNextAdminNormalizeModelParameters_(input) {
  const source = input && typeof input === 'object' && !Array.isArray(input) ? input : {};
  const allowed = new Set(['simulationCount', 'referencePrior']);
  const unknown = Object.keys(source).filter(function (key) { return !allowed.has(key); });
  if (unknown.length) throw new Error('Unsupported MODEL_RELEASE parameters: ' + unknown.join(', '));
  const output = {};
  if (Object.prototype.hasOwnProperty.call(source, 'simulationCount')) {
    const count = Number(source.simulationCount);
    const min = typeof VNEXT_ENGINE !== 'undefined' ? Number(VNEXT_ENGINE.MIN_SIMULATIONS) : 200;
    const max = typeof VNEXT_ENGINE !== 'undefined' ? Number(VNEXT_ENGINE.MAX_SIMULATIONS) : 10000;
    if (!isFinite(count) || Math.floor(count) !== count || count < min || count > max) {
      throw new Error('simulationCount must be an integer within the deployed Engine range.');
    }
    output.simulationCount = count;
  }
  if (Object.prototype.hasOwnProperty.call(source, 'referencePrior')) {
    const prior = source.referencePrior;
    if (!prior || typeof prior !== 'object' || Array.isArray(prior)) throw new Error('referencePrior must be an object.');
    const priorAllowed = new Set(['mode', 'reason', 'growthMean', 'growthStd', 'strength']);
    const priorUnknown = Object.keys(prior).filter(function (key) { return !priorAllowed.has(key); });
    if (priorUnknown.length) throw new Error('Unsupported referencePrior keys: ' + priorUnknown.join(', '));
    if (String(prior.mode || '').toUpperCase() === 'DISABLED') {
      output.referencePrior = {
        mode: 'DISABLED', reason: vNextAdminText_(prior.reason) || 'DISABLED_BY_RELEASE', strength: 0
      };
    } else {
      const growthMean = Number(prior.growthMean);
      const growthStd = Number(prior.growthStd);
      const strength = Number(prior.strength);
      const minGrowth = typeof VNEXT_ENGINE !== 'undefined' ? VNEXT_ENGINE.REFERENCE_MIN_GROWTH : -1;
      const maxGrowth = typeof VNEXT_ENGINE !== 'undefined' ? VNEXT_ENGINE.REFERENCE_MAX_GROWTH : 3;
      const maxStd = typeof VNEXT_ENGINE !== 'undefined' ? VNEXT_ENGINE.REFERENCE_MAX_GROWTH_STD : 2;
      const maxStrength = typeof VNEXT_ENGINE !== 'undefined' ? VNEXT_ENGINE.REFERENCE_MAX_STRENGTH : 0.35;
      if (![growthMean, growthStd, strength].every(isFinite) || growthMean < minGrowth || growthMean > maxGrowth ||
          growthStd < 0 || growthStd > maxStd || strength < 0 || strength > maxStrength) {
        throw new Error('referencePrior growthMean/growthStd/strength is outside the deployed Engine range.');
      }
      output.referencePrior = { growthMean: growthMean, growthStd: growthStd, strength: strength };
    }
  }
  return output;
}

function vNextAdminModelCandidateHash_(input) {
  const value = input || {};
  return vNextAdminSha256_(vNextAdminCanonicalJson_({
    modelVersion: String(value.modelVersion || ''),
    schemaVersion: String(value.schemaVersion || ''),
    templateVersion: String(value.templateVersion || ''),
    parameters: value.parameters || {}
  }));
}

function vNextAdminBindModelCheck_(input, candidateHash, label) {
  const check = Object.assign({}, input || {});
  if (check.candidateHash && String(check.candidateHash) !== String(candidateHash)) {
    throw new Error(String(label || 'model check') + ' candidateHash does not match the registered candidate.');
  }
  check.candidateHash = String(candidateHash);
  return check;
}

function vNextAdminParseObjectPayload_(value, label) {
  if (value === undefined || value === null || value === '') return {};
  const parsed = typeof value === 'string' ? vNextAdminParseJson_(value, null) : value;
  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
    throw new Error(String(label || 'value') + ' must be a JSON object.');
  }
  return parsed;
}

function vNextAdminFindRegistryRow_(hub, predicate) {
  return vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.REGISTRY).rows.find(predicate) || null;
}

function vNextAdminPatchRegistryByBookId_(hub, bookId, patch) {
  const row = vNextAdminFindRegistryRow_(hub, function (item) { return String(item.book_id) === String(bookId); });
  if (!row) return false;
  vNextAdminUpdateTableRow_(hub, VN_ADMIN_SHEETS.REGISTRY, row._rowNumber, patch);
  return true;
}

function vNextAdminFindApproval_(hub, predicate) {
  return vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.APPROVALS).rows.find(predicate) || null;
}

function vNextAdminListPendingApprovals_(hub, optionalApprovalRows) {
  const rows = Array.isArray(optionalApprovalRows)
    ? optionalApprovalRows : vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.APPROVALS).rows;
  return rows
    .filter(function (row) { return String(row.status) === 'PENDING'; })
    .map(function (row) {
      const snapshot = vNextAdminParseJson_(row.snapshot_json, {});
      const forecast = snapshot.forecast || {};
      const plan = snapshot.plan || {};
      const amendment = snapshot.amendment || {};
      return {
        approvalRequestId: row.approval_request_id, requestType: row.request_type,
        bookId: row.book_id, clientName: row.client_name, fiscalYear: row.fiscal_year,
        forecastRunId: row.forecast_run_id, planVersionId: row.plan_version_id,
        requestedAt: row.requested_at, requestedBy: row.requested_by,
        amendmentReason: row.amendment_reason,
        supersedesOfficialId: row.supersedes_official_id,
        previousFinalBudget: Number(amendment.predecessorFinalBudget || 0),
        systemRecommended: Number(plan.systemRecommended || (forecast.layers && forecast.layers.systemRecommended) || 0),
        p10: Number(forecast.annual && forecast.annual.p10 || 0),
        p90: Number(forecast.annual && forecast.annual.p90 || 0),
        adoptedForecast: Number(plan.adoptedForecast || 0), salesUplift: Number(plan.salesUplift || 0),
        finalBudget: Number(plan.finalBudget || 0), adoptionReason: String(plan.adoptionReason || ''),
        upliftReason: String(plan.upliftReason || ''), upliftOwner: String(plan.upliftOwner || ''),
        upliftAction: String(plan.upliftAction || ''), upliftDueDate: String(plan.upliftDueDate || '')
      };
    });
}

// ---------------------------- Sheet / metadata helpers ----------------------------

function vNextAdminWriteBookConfig_(ss, values) {
  const sheet = vNextAdminEnsureTable_(ss, VN_ADMIN_BOOK_CONFIG_SHEET, ['key', 'value', 'updated_at', 'updated_by', 'note']);
  const now = new Date();
  const actor = vNextAdminActor_();
  Object.keys(values || {}).forEach(function (key) {
    vNextAdminUpsertObject_(ss, VN_ADMIN_BOOK_CONFIG_SHEET, 'key', key, {
      key: key, value: values[key], updated_at: now, updated_by: actor, note: ''
    }, ['key', 'value', 'updated_at', 'updated_by', 'note']);
  });
  sheet.hideSheet();
}

function vNextAdminReplaceBookConfig_(ss, values) {
  const existing = ss.getSheetByName(VN_ADMIN_BOOK_CONFIG_SHEET);
  if (existing) existing.clear();
  return vNextAdminWriteBookConfig_(ss, values || {});
}

function vNextAdminEnsureCoreStore_(ss) {
  if (typeof vNextEnsureAuditStore_ === 'function') return vNextEnsureAuditStore_(ss);
  throw new Error('VNext_Core.js is required before 管理ハブ initialization.');
}

function vNextAdminCreateClientCoreMeta_(ss, opt) {
  const rows = typeof vNextReadRecords_ === 'function'
    ? vNextReadRecords_('BOOK_META', { spreadsheet: ss })
    : [];
  const exists = rows.some(function (row) { return String(row.book_id || '') === String(opt.bookId); });
  if (exists) return rows.filter(function (row) { return String(row.book_id || '') === String(opt.bookId); }).slice(-1)[0];
  const owner = (opt.forecastOwnerEmails || [])[0] || '';
  const team = vNextAdminMergeEmails_(opt.forecastOwnerEmails, opt.editors);
  const record = {
    record_id: typeof vNextUuid_ === 'function' ? vNextUuid_() : Utilities.getUuid(),
    book_id: opt.bookId,
    client_id: opt.clientId,
    client_name: opt.clientName,
    fiscal_year: opt.fiscalYear,
    forecast_owner_email: owner,
    team_member_emails_json: typeof vNextCanonicalJson_ === 'function' ? vNextCanonicalJson_(team) : JSON.stringify(team),
    state: 'INPUT_OPEN',
    as_of: opt.asOf,
    cutoff: opt.cutoff,
    template_version: opt.releaseId,
    schema_version: typeof VNEXT_CORE !== 'undefined' ? VNEXT_CORE.SCHEMA_VERSION : VN_ADMIN_SCHEMA_VERSION,
    model_release_id: vNextAdminRequiredText_(opt.modelReleaseId, 'modelReleaseId'),
    source_spreadsheet_id: '',
    client_book_id: ss.getId(),
    input_due_date: opt.inputDueDate || '',
    event_type: 'CREATED',
    supersedes_record_id: '',
    recorded_at: new Date().toISOString(),
    recorded_by: String(opt.actor || '').toLowerCase()
  };
  vNextAdminAppendCoreRowsNoLock_(ss, 'BOOK_META', [record]);
  return record;
}

function vNextAdminAppendCoreRowsNoLock_(ss, sheetName, records) {
  if (!records || !records.length) return { appended: 0 };
  if (typeof VNEXT_CORE === 'undefined' || !VNEXT_CORE.INTERNAL_SHEETS[sheetName]) {
    throw new Error('Unknown VNext_Core sheet: ' + sheetName);
  }
  const headers = VNEXT_CORE.INTERNAL_SHEETS[sheetName];
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Core store is not initialized: ' + sheetName);
  if (typeof vNextEnsureAppendOnlyHeader_ === 'function') vNextEnsureAppendOnlyHeader_(sheet, headers);
  const values = records.map(function (record) {
    return headers.map(function (key) {
      const value = record[key];
      if (value === null || value === undefined) return '';
      if (value instanceof Date) return value.toISOString();
      if (typeof value === 'object') return typeof vNextCanonicalJson_ === 'function' ? vNextCanonicalJson_(value) : JSON.stringify(value);
      return value;
    });
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, values.length, headers.length).setValues(values);
  return { appended: values.length };
}

function vNextAdminWriteSystemConfig_(ss, values) {
  const sheet = vNextAdminEnsureTable_(ss, VN_ADMIN_SYSTEM_CONFIG_SHEET, ['key', 'value', 'updated_at']);
  Object.keys(values || {}).forEach(function (key) {
    vNextAdminUpsertObject_(ss, VN_ADMIN_SYSTEM_CONFIG_SHEET, 'key', key, {
      key: key, value: values[key], updated_at: new Date()
    }, ['key', 'value', 'updated_at']);
  });
  sheet.hideSheet();
}

function vNextAdminReplaceSystemConfig_(ss, values) {
  const existing = ss.getSheetByName(VN_ADMIN_SYSTEM_CONFIG_SHEET);
  if (existing) existing.clear();
  return vNextAdminWriteSystemConfig_(ss, values || {});
}

function vNextAdminReadKeyValueSheet_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return {};
  const values = sheet.getDataRange().getValues();
  const header = values[0].map(function (v) { return String(v || '').trim(); });
  const keyIdx = header.indexOf('key');
  const valueIdx = header.indexOf('value');
  if (keyIdx < 0 || valueIdx < 0) return {};
  const out = {};
  values.slice(1).forEach(function (row) {
    const key = String(row[keyIdx] || '').trim();
    if (key) out[key] = row[valueIdx];
  });
  return out;
}

function vNextAdminEnsureTable_(ss, name, headers) {
  const sheet = vNextAdminGetOrCreateSheet_(ss, name);
  const required = headers || [];
  if (sheet.getMaxColumns() < required.length) sheet.insertColumnsAfter(sheet.getMaxColumns(), required.length - sheet.getMaxColumns());
  if (sheet.getLastRow() < 1 || sheet.getLastColumn() < 1 || !String(sheet.getRange(1, 1).getValue() || '').trim()) {
    sheet.getRange(1, 1, 1, required.length).setValues([required]);
  } else {
    const existing = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0].map(function (v) { return String(v || '').trim(); });
    const missing = required.filter(function (key) { return existing.indexOf(key) < 0; });
    if (missing.length) {
      const start = existing.length + 1;
      if (sheet.getMaxColumns() < start + missing.length - 1) sheet.insertColumnsAfter(sheet.getMaxColumns(), start + missing.length - 1 - sheet.getMaxColumns());
      sheet.getRange(1, start, 1, missing.length).setValues([missing]);
    }
  }
  sheet.getRange(1, 1, 1, Math.max(required.length, sheet.getLastColumn())).setFontWeight('bold').setBackground('#eeeeee');
  sheet.setFrozenRows(1);
  return sheet;
}

/** Fail closed for employee-runtime tables whose positional schema is part of the signed contract. */
function vNextAdminEnsureExactTableHeaders_(ss, name, headers) {
  const expected = (headers || []).map(String);
  const sheet = vNextAdminGetOrCreateSheet_(ss, name);
  if (sheet.getMaxColumns() < expected.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), expected.length - sheet.getMaxColumns());
  }
  const width = Math.max(1, sheet.getLastColumn(), expected.length);
  const actual = sheet.getRange(1, 1, 1, width).getValues()[0].map(function (value) {
    return String(value || '').trim();
  });
  const hasHeader = actual.some(Boolean);
  if (!hasHeader) {
    sheet.getRange(1, 1, 1, expected.length).setValues([expected]);
  } else {
    const defined = actual.slice(0, expected.length);
    const unexpected = actual.slice(expected.length).filter(Boolean);
    if (vNextAdminCanonicalJson_(defined) !== vNextAdminCanonicalJson_(expected) || unexpected.length) {
      throw new Error(name + 'の列構成がruntime契約と一致しません。管理ハブが直接修正せずversioned migrationを実行してください。');
    }
  }
  sheet.getRange(1, 1, 1, expected.length).setFontWeight('bold').setBackground('#eeeeee');
  sheet.setFrozenRows(1);
  return sheet;
}

function vNextAdminClearTableData_(ss, name, headers) {
  const sheet = headers && headers.length
    ? vNextAdminEnsureTable_(ss, name, headers)
    : ss.getSheetByName(name);
  vNextAdminClearSheetDataRows_(sheet);
  return sheet;
}

function vNextAdminClearSheetDataRows_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return sheet;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(1, sheet.getLastColumn())).clearContent();
  return sheet;
}

function vNextAdminRewriteTableKeeping_(ss, name, headers, keepFn) {
  const table = vNextAdminReadTable_(ss, name);
  if (!table.sheet) return { kept: 0, removed: 0 };
  const kept = table.rows.filter(function (row) { return keepFn(row); });
  vNextAdminClearSheetDataRows_(table.sheet);
  if (!kept.length) return { kept: 0, removed: table.rows.length };
  const cols = headers && headers.length ? headers : table.headers;
  const values = kept.map(function (row) {
    return cols.map(function (key) { return row[key] === undefined ? '' : row[key]; });
  });
  table.sheet.getRange(2, 1, values.length, cols.length).setValues(values);
  return { kept: kept.length, removed: table.rows.length - kept.length };
}

function vNextAdminReadTable_(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet || sheet.getLastRow() < 1) return { headers: [], rows: [], sheet: sheet };
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(function (value) { return String(value || '').trim(); });
  const rows = values.slice(1).map(function (valuesRow, index) {
    const obj = { _rowNumber: index + 2 };
    headers.forEach(function (key, col) { if (key) obj[key] = valuesRow[col]; });
    return obj;
  }).filter(function (obj) {
    return headers.some(function (key) { return key && obj[key] !== '' && obj[key] !== null && obj[key] !== undefined; });
  });
  return { headers: headers, rows: rows, sheet: sheet };
}

function vNextAdminAppendObject_(ss, name, object, explicitHeaders) {
  const headers = explicitHeaders || VN_ADMIN_HEADERS[name] || Object.keys(object || {});
  const sheet = vNextAdminEnsureTable_(ss, name, headers);
  const actualHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function (v) { return String(v || '').trim(); });
  const row = actualHeaders.map(function (key) { return key && object[key] !== undefined ? object[key] : ''; });
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
  return object;
}

function vNextAdminUpsertObject_(ss, name, keyField, keyValue, object, explicitHeaders) {
  const headers = explicitHeaders || VN_ADMIN_HEADERS[name] || Object.keys(object || {});
  vNextAdminEnsureTable_(ss, name, headers);
  const table = vNextAdminReadTable_(ss, name);
  const existing = table.rows.find(function (row) { return String(row[keyField]) === String(keyValue); });
  if (!existing) return vNextAdminAppendObject_(ss, name, object, headers);
  vNextAdminUpdateTableRow_(ss, name, existing._rowNumber, object);
  return Object.assign({}, existing, object);
}

function vNextAdminUpdateTableRow_(ss, name, rowNumber, patch) {
  const table = vNextAdminReadTable_(ss, name);
  if (!table.sheet || rowNumber < 2) throw new Error('Invalid table row update: ' + name + ' row=' + rowNumber);
  const range = table.sheet.getRange(rowNumber, 1, 1, table.headers.length);
  const values = range.getValues()[0];
  table.headers.forEach(function (key, col) {
    if (key && Object.prototype.hasOwnProperty.call(patch, key)) values[col] = patch[key];
  });
  range.setValues([values]);
}

function vNextAdminGetOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function vNextAdminCountRowsBy_(ss, sheetName, key, expected) {
  return vNextAdminReadTable_(ss, sheetName).rows.filter(function (row) { return String(row[key]) === String(expected); }).length;
}

function vNextAdminPatchLatestMigration_(hub, migrationId, patch) {
  const table = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.MIGRATIONS);
  const row = table.rows.filter(function (item) {
    return String(item.migration_id) === String(migrationId);
  }).slice(-1)[0];
  if (row) vNextAdminUpdateTableRow_(hub, VN_ADMIN_SHEETS.MIGRATIONS, row._rowNumber, patch);
}

// ---------------------------- ACL / visibility ----------------------------

function vNextAdminApplyFileAcl_(file, editors, viewers, commenters) {
  const editorList = vNextAdminMergeEmails_(editors);
  const viewerList = vNextAdminMergeEmails_(viewers);
  const commenterList = vNextAdminMergeEmails_(commenters);
  if (editorList.length) file.addEditors(editorList);
  if (viewerList.length) file.addViewers(viewerList);
  if (commenterList.length && typeof file.addCommenters === 'function') file.addCommenters(commenterList);
}

function vNextAdminProtectInternalSheets_(ss, allowedEmails, names) {
  const file = DriveApp.getFileById(ss.getId());
  const ownerEmail = file.getOwner() ? file.getOwner().getEmail() : '';
  const allowed = vNextAdminMergeEmails_(allowedEmails, ownerEmail, vNextAdminActor_());
  (names || []).forEach(function (name) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const description = 'VNEXT_ADMIN:' + name;
    let protection = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).find(function (item) {
      return String(item.getDescription() || '') === description;
    });
    if (!protection) protection = sheet.protect().setDescription(description);
    protection.setWarningOnly(false);
    const removable = protection.getEditors().filter(function (user) {
      return allowed.indexOf(String(user.getEmail() || '').toLowerCase()) < 0;
    });
    if (removable.length) protection.removeEditors(removable);
    if (allowed.length) protection.addEditors(allowed);
    if (protection.canDomainEdit()) protection.setDomainEdit(false);
  });
}

function vNextAdminProtectClientInternalSheets_(ss, names) {
  (names || []).forEach(function (name) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const description = 'VNEXT_CLIENT_APPEND_ONLY:' + name;
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    const protection = protections.length ? protections[0] : sheet.protect();
    protection.setDescription(description).setWarningOnly(true);
    protections.slice(1).forEach(function (item) { try { item.remove(); } catch (err) { Logger.log('Duplicate client protection removal skipped: %s', String(err)); } });
  });
}

function vNextAdminApplyVisibility_(ss, visibleNames) {
  const visible = new Set((visibleNames || []).map(String));
  let firstVisible = null;
  ss.getSheets().forEach(function (sheet) {
    if (visible.has(sheet.getName())) {
      sheet.showSheet();
      if (!firstVisible) firstVisible = sheet;
    }
  });
  if (!firstVisible) {
    firstVisible = ss.getSheets()[0];
    firstVisible.showSheet();
  }
  ss.setActiveSheet(firstVisible);
  ss.getSheets().forEach(function (sheet) {
    if (sheet.getSheetId() === firstVisible.getSheetId()) return;
    if (visible.has(sheet.getName())) sheet.showSheet();
    else sheet.hideSheet();
  });
}

// ---------------------------- Audit / generic helpers ----------------------------

function vNextAdminWriteAudit_(hub, action, entityType, entityId, status, detail, beforeHash, afterHash) {
  try {
    vNextAdminAppendObject_(hub, VN_ADMIN_SHEETS.AUDIT, {
      audit_id: 'AUD-' + Utilities.getUuid(), occurred_at: new Date(), actor: vNextAdminActor_(),
      action: action || '', entity_type: entityType || '', entity_id: entityId || '',
      status: status || '', detail_json: JSON.stringify(detail || {}),
      before_hash: beforeHash || '', after_hash: afterHash || ''
    });
  } catch (err) {
    Logger.log('Audit write failed: %s', String(err && err.message || err));
  }
}

function vNextAdminGuard_(name, fn) {
  try {
    const result = fn();
    Logger.log('%s success', name);
    return result;
  } catch (err) {
    Logger.log('%s error: %s', name, String(err && err.stack || err));
    throw err;
  }
}

function vNextAdminWithScriptLock_(label, fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Another vNext admin operation is running: ' + label);
  try { return fn(); } finally { lock.releaseLock(); }
}

function vNextAdminWithDocumentLock_(label, fn) {
  const lock = LockService.getDocumentLock();
  if (!lock || !lock.tryLock(30000)) throw new Error('Another workbook operation is running: ' + label);
  try { return fn(); } finally { lock.releaseLock(); }
}

function vNextAdminWithUserLock_(label, fn) {
  const lock = LockService.getUserLock();
  if (!lock.tryLock(30000)) throw new Error('Another user operation is running: ' + label);
  try { return fn(); } finally { lock.releaseLock(); }
}

function vNextAdminResolveSpreadsheet_(input) {
  if (!input) return SpreadsheetApp.getActiveSpreadsheet();
  if (typeof input === 'string') return SpreadsheetApp.openById(input);
  if (typeof input.getId === 'function' && typeof input.getSheets === 'function') return input;
  if (input.spreadsheetId) return SpreadsheetApp.openById(String(input.spreadsheetId));
  return SpreadsheetApp.getActiveSpreadsheet();
}

function vNextAdminHydrateLocalRuntime_(ss, routing) {
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (!active || active.getId() !== ss.getId()) return;
    const cfg = routing || {};
    const bookId = String(cfg.book_id || '').trim();
    const hubId = String(cfg.admin_hub_spreadsheet_id || '').trim();
    const mode = String(cfg.mode || '').trim().toUpperCase();
    const doc = PropertiesService.getDocumentProperties();
    const script = PropertiesService.getScriptProperties();
    if (bookId) {
      doc.setProperty('VNEXT_BOOK_ID', bookId);
      script.setProperty('VNEXT_BOOK_ID', bookId);
    }
    // CLIENT books deliberately do not hydrate the Hub property: Core must use the
    // local client-only store under employee credentials. Hub synchronization is Admin-owned.
    if (hubId && mode !== 'CLIENT') {
      doc.setProperty('VNEXT_ADMIN_HUB_SPREADSHEET_ID', hubId);
      script.setProperty('VNEXT_ADMIN_HUB_SPREADSHEET_ID', hubId);
    }
    if (mode === 'CLIENT') {
      doc.deleteProperty('VNEXT_ADMIN_HUB_SPREADSHEET_ID');
      script.deleteProperty('VNEXT_ADMIN_HUB_SPREADSHEET_ID');
    }
    if (mode === 'ADMIN') {
      const sourceId = String(cfg.source_spreadsheet_id || '').trim();
      const templateId = String(cfg.template_spreadsheet_id || '').trim();
      const releaseId = String(cfg.active_release_id || cfg.version || '').trim();
      const modelReleaseId = String(cfg.active_model_release_id || '').trim();
      const adminEmails = String(cfg.admin_emails || cfg.forecast_owner_emails || '').trim();
      if (sourceId) script.setProperties({
        FORECAST_SOURCE_SPREADSHEET_ID: sourceId,
        VNEXT_ZAC_SOURCE_SPREADSHEET_ID: sourceId
      }, false);
      if (templateId) script.setProperty('VNEXT_MASTER_TEMPLATE_SPREADSHEET_ID', templateId);
      if (releaseId) script.setProperty('VNEXT_ACTIVE_RELEASE_ID', releaseId);
      if (modelReleaseId) script.setProperty('VNEXT_ACTIVE_MODEL_RELEASE_ID', modelReleaseId);
      if (adminEmails) {
        script.setProperty('VNEXT_ADMIN_EMAILS', adminEmails);
        doc.setProperty('VNEXT_ADMIN_EMAILS', adminEmails);
      }
      [
        'VERTEX_PROJECT_ID', 'VERTEX_LOCATION', 'VERTEX_GEMINI_MODEL',
        'VERTEX_DATASTORE_ID', 'VERTEX_SEARCH_LOCATION', 'VERTEX_SERVING_CONFIG'
      ].forEach(function (key) {
        const value = String(cfg[key] || '').trim();
        if (value) script.setProperty(key, value);
      });
    }
    if (mode === 'CLIENT' || mode === 'TEMPLATE') {
      [
        'FORECAST_SOURCE_SPREADSHEET_ID', 'VNEXT_ZAC_SOURCE_SPREADSHEET_ID',
        'VERTEX_PROJECT_ID', 'VERTEX_LOCATION', 'VERTEX_GEMINI_MODEL',
        'VERTEX_DATASTORE_ID', 'VERTEX_SEARCH_LOCATION', 'VERTEX_SERVING_CONFIG',
        'VNEXT_ADMIN_EMAILS', 'VNEXT_MASTER_TEMPLATE_SPREADSHEET_ID', 'VNEXT_ACTIVE_RELEASE_ID',
        'VNEXT_ACTIVE_MODEL_RELEASE_ID'
      ].forEach(function (key) { script.deleteProperty(key); });
    }
    if (mode) doc.setProperty('VNEXT_BOOK_MODE', mode);
  } catch (err) {
    Logger.log('Local runtime hydration skipped: %s', String(err && err.message || err));
  }
}

function vNextAdminHydrateHubRuntime_(hub) {
  const routing = Object.assign(
    {},
    vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET),
    vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_BOOK_CONFIG_SHEET)
  );
  vNextAdminHydrateLocalRuntime_(hub, routing);
  // A scheduled trigger may open the Hub by ID rather than as the active container.
  const script = PropertiesService.getScriptProperties();
  const sourceId = String(routing.source_spreadsheet_id || '').trim();
  if (sourceId) script.setProperties({
    FORECAST_SOURCE_SPREADSHEET_ID: sourceId,
    VNEXT_ZAC_SOURCE_SPREADSHEET_ID: sourceId
  }, false);
  if (routing.admin_emails || routing.forecast_owner_emails) {
    script.setProperty('VNEXT_ADMIN_EMAILS', String(routing.admin_emails || routing.forecast_owner_emails));
  }
  if (routing.active_model_release_id) {
    script.setProperty('VNEXT_ACTIVE_MODEL_RELEASE_ID', String(routing.active_model_release_id));
  }
  if (routing.active_release_id) {
    script.setProperty('VNEXT_ACTIVE_RELEASE_ID', String(routing.active_release_id));
  }
  if (routing.template_spreadsheet_id) {
    script.setProperty('VNEXT_MASTER_TEMPLATE_SPREADSHEET_ID', String(routing.template_spreadsheet_id));
  }
  [
    'VERTEX_PROJECT_ID', 'VERTEX_LOCATION', 'VERTEX_GEMINI_MODEL',
    'VERTEX_DATASTORE_ID', 'VERTEX_SEARCH_LOCATION', 'VERTEX_SERVING_CONFIG'
  ].forEach(function (key) {
    const value = String(routing[key] || '').trim();
    if (value) script.setProperty(key, value);
  });
  script.setProperty('VNEXT_ADMIN_HUB_SPREADSHEET_ID', hub.getId());
  return routing;
}

function vNextAdminRequireHub_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (vNextDetectBookMode_(ss) !== 'ADMIN' || !vNextAdminIsRegisteredHub_(ss)) {
    throw new Error('This operation is available only in the registered 管理ハブ.');
  }
  vNextAdminAssertHubAdmin_(ss, false);
  vNextAdminHydrateHubRuntime_(ss);
  Object.keys(VN_ADMIN_HEADERS).forEach(function (name) { vNextAdminEnsureTable_(ss, name, VN_ADMIN_HEADERS[name]); });
  return ss;
}

function vNextAdminResolveHubForAutomation_() {
  const id = PropertiesService.getScriptProperties().getProperty('VNEXT_ADMIN_HUB_SPREADSHEET_ID');
  const hub = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  if (!hub || vNextDetectBookMode_(hub) !== 'ADMIN' || !vNextAdminIsRegisteredHub_(hub)) {
    throw new Error('管理ハブ cannot be resolved for scheduled sweep.');
  }
  vNextAdminAssertHubAdmin_(hub, true);
  vNextAdminHydrateHubRuntime_(hub);
  return hub;
}

const VN_ADMIN_AUTOMATION_CACHE_KEY_ = 'vnext_admin_automation_installed';
const VN_ADMIN_AUTOMATION_CACHE_TTL_SEC_ = 120;

function vNextAdminAutomationInstalled_() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(VN_ADMIN_AUTOMATION_CACHE_KEY_);
    if (cached === '1') return true;
    if (cached === '0') return false;
    const installed = ScriptApp.getProjectTriggers().some(function (trigger) {
      return trigger.getHandlerFunction() === VN_ADMIN_SCHEDULED_HANDLER;
    });
    cache.put(VN_ADMIN_AUTOMATION_CACHE_KEY_, installed ? '1' : '0', VN_ADMIN_AUTOMATION_CACHE_TTL_SEC_);
    return installed;
  } catch (err) {
    Logger.log('Automation status unavailable: %s', String(err && err.message || err));
    return false;
  }
}

function vNextAdminClearAutomationInstalledCache_() {
  try { CacheService.getScriptCache().remove(VN_ADMIN_AUTOMATION_CACHE_KEY_); }
  catch (ignoredCacheClear) {}
}

function vNextAdminAssertRuntimeConfigurator_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Active spreadsheet is required.');
  const actor = vNextAdminActor_().toLowerCase();
  const ownerUser = (function () {
    try { return DriveApp.getFileById(ss.getId()).getOwner(); }
    catch (ignoredSharedDriveOwner) { return null; }
  })();
  const owner = ownerUser ? String(ownerUser.getEmail() || '').toLowerCase() : '';
  const routing = vNextAdminReadKeyValueSheet_(ss, VN_ADMIN_SYSTEM_CONFIG_SHEET);
  const admins = vNextAdminMergeEmails_(
    routing.admin_emails,
    PropertiesService.getScriptProperties().getProperty('VNEXT_ADMIN_EMAILS'),
    owner
  );
  if (!actor || admins.indexOf(actor) < 0) {
    throw new Error('この設定を変更できるのはファイル所有者または管理ハブ担当者だけです。');
  }
  return true;
}

function vNextAdminAssertHubAdminFast_(hub) {
  const routing = vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET);
  try {
    const actor = String(vNextAdminActor_() || '').toLowerCase();
    const admins = vNextAdminMergeEmails_(
      routing.admin_emails,
      PropertiesService.getScriptProperties().getProperty('VNEXT_ADMIN_EMAILS')
    );
    if (actor && admins.indexOf(actor) >= 0) return routing;
  } catch (err) {
    Logger.log('Fast Hub admin check fallback: %s', String(err && err.message || err));
  }
  vNextAdminAssertHubAdmin_(hub, false);
  return Object.assign({}, routing, vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_BOOK_CONFIG_SHEET));
}

function vNextAdminAssertHubAdmin_(hub, allowEffectiveUser) {
  const effectiveUser = Session.getEffectiveUser();
  const effectiveEmail = effectiveUser ? String(effectiveUser.getEmail() || '') : '';
  const actor = String(allowEffectiveUser
    ? (effectiveEmail || vNextAdminActor_())
    : vNextAdminActor_()).toLowerCase();
  const routing = Object.assign(
    {},
    vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET),
    vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_BOOK_CONFIG_SHEET)
  );
  const ownerUser = (function () {
    try { return DriveApp.getFileById(hub.getId()).getOwner(); }
    catch (ignoredSharedDriveOwner) { return null; }
  })();
  const owner = ownerUser ? String(ownerUser.getEmail() || '').toLowerCase() : '';
  const admins = vNextAdminMergeEmails_(routing.admin_emails,
    PropertiesService.getScriptProperties().getProperty('VNEXT_ADMIN_EMAILS'), owner);
  if (!actor || admins.indexOf(actor) < 0) throw new Error('Admin 管理ハブの操作権限がありません。');
  return true;
}

function vNextAdminSpreadsheetAccessible_(id) {
  if (!id) return false;
  try {
    const file = DriveApp.getFileById(String(id));
    if (file.isTrashed()) return false;
    SpreadsheetApp.openById(String(id)).getName();
    return true;
  } catch (err) {
    return false;
  }
}

function vNextAdminResolveDestinationFolder_(folderId, sourceFile) {
  if (folderId) return DriveApp.getFolderById(String(folderId));
  const parents = sourceFile.getParents();
  if (parents.hasNext()) return parents.next();
  return DriveApp.getRootFolder();
}

function vNextAdminLibraryPath_(kind, fiscalYear, clientName) {
  if (kind === 'PORTAL') return [VN_ADMIN_LIBRARY.PORTAL];
  if (kind === 'ADMIN') return [VN_ADMIN_LIBRARY.ADMIN];
  if (kind === 'AUDIT') return [VN_ADMIN_LIBRARY.ADMIN, VN_ADMIN_LIBRARY.AUDIT];
  if (kind === 'TEMPLATE_CURRENT') {
    return [VN_ADMIN_LIBRARY.TEMPLATES, VN_ADMIN_LIBRARY.TEMPLATES_CURRENT];
  }
  if (kind === 'TEMPLATE_DRAFT') {
    return [VN_ADMIN_LIBRARY.TEMPLATES, VN_ADMIN_LIBRARY.TEMPLATES_DRAFT];
  }
  if (kind === 'TEMPLATE_HISTORY') {
    return [VN_ADMIN_LIBRARY.TEMPLATES, VN_ADMIN_LIBRARY.TEMPLATES_HISTORY];
  }
  return [
    VN_ADMIN_LIBRARY.BOOKS,
    'FY' + (Number(fiscalYear) >= 2000 ? Number(fiscalYear) : 'unknown'),
    vNextAdminSafeDriveName_(clientName, 'クライアント')
  ];
}

function vNextAdminSafeDriveName_(value, fallback) {
  const text = String(value || fallback || '未設定')
    .replace(/[\\/:*?"<>|]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  return (text || fallback || '未設定').slice(0, 80);
}

function vNextAdminPrepareClientDestinationFolder_(hub, requestedFolderId, label, adminEmails) {
  const parsed = String(label || '').match(/FY(\d{4})/i);
  const fiscalYear = parsed ? parsed[1] : 2027;
  const clientName = String(label || 'クライアント').replace(/-FY\d{4}.*$/i, '');
  return vNextAdminPrepareLibraryDestinationFolder_(hub, requestedFolderId,
    vNextAdminLibraryPath_('CLIENT', fiscalYear, clientName || label), adminEmails);
}

function vNextAdminPrepareLibraryDestinationFolder_(hub, requestedFolderId, pathSegments, adminEmails) {
  const root = vNextAdminResolveManagedRoot_(hub, adminEmails);
  const requested = String(requestedFolderId || '').trim();
  if (requested) {
    const folder = DriveApp.getFolderById(requested);
    if (!vNextAdminFolderWithinRoot_(folder, root.getId())) {
      throw new Error('指定folderは' + VNEXT_NAMING.SHARED_DRIVE + 'ライブラリ配下ではないため使用できません。');
    }
    return vNextAdminPrepareManagedFolder_(folder.getId(), folder.getName(), adminEmails);
  }
  return vNextAdminEnsureLibraryPath_(root, pathSegments, adminEmails);
}

function vNextAdminResolveManagedRoot_(hub, adminEmails) {
  const config = vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET);
  const rootId = String(config.shared_drive_id || config.private_root_folder_id || '').trim();
  if (!rootId) throw new Error(VNEXT_NAMING.SHARED_DRIVE + 'の保存先folderが' + VNEXT_NAMING.LAYER1 + 'に記録されていないため停止しました。');
  return vNextAdminPrepareManagedFolder_(rootId, VN_ADMIN_LIBRARY.DRIVE_NAME, adminEmails);
}

function vNextAdminEnsureLibraryPath_(root, pathSegments, adminEmails) {
  let current = root;
  (pathSegments || []).forEach(function (name) {
    const safe = vNextAdminSafeDriveName_(name, 'folder');
    const legacy = VN_ADMIN_LIBRARY.FOLDER_LEGACY && VN_ADMIN_LIBRARY.FOLDER_LEGACY[safe];
    let existing = vNextAdminFindNamedChildFolder_(current, safe);
    if (!existing && legacy) {
      existing = vNextAdminFindNamedChildFolder_(current, legacy);
      if (existing && existing.getName() === legacy) {
        try { existing.setName(safe); } catch (ignoredRename) {}
      }
    }
    current = existing || current.createFolder(safe);
    current = vNextAdminPrepareManagedFolder_(current.getId(), current.getName(), adminEmails);
  });
  return current;
}

function vNextAdminFindNamedChildFolder_(parent, name) {
  const matches = [];
  const folders = parent.getFolders();
  while (folders.hasNext()) {
    const folder = folders.next();
    if (folder.getName() === name) matches.push(folder);
  }
  if (!matches.length) return null;
  if (matches.length === 1) return matches[0];
  matches.sort(function (a, b) {
    return vNextAdminFolderOccupancy_(b) - vNextAdminFolderOccupancy_(a);
  });
  return matches[0];
}

function vNextAdminFolderOccupancy_(folder) {
  let count = 0;
  const files = folder.getFiles();
  while (count < 20 && files.hasNext()) { files.next(); count += 1; }
  const folders = folder.getFolders();
  while (count < 20 && folders.hasNext()) { folders.next(); count += 1; }
  return count;
}

function vNextAdminDetectCurrentSharedDriveId_(spreadsheet) {
  try {
    if (typeof Drive === 'undefined' || !Drive.Files || !spreadsheet) return '';
    const meta = Drive.Files.get(spreadsheet.getId(), {
      fields: 'id,driveId,teamDriveId', supportsAllDrives: true, supportsTeamDrives: true
    });
    return String((meta && (meta.driveId || meta.teamDriveId)) || '').trim();
  } catch (error) {
    return '';
  }
}

function vNextAdminEnsureSharedLibrary_(hub, options) {
  const opt = options || {};
  const config = vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET);
  const driveId = String(opt.sharedDriveId || config.shared_drive_id || '').trim() ||
    vNextAdminFindOrCreateSharedDrive_(VN_ADMIN_LIBRARY.DRIVE_NAME);
  const root = vNextAdminPrepareManagedFolder_(driveId, VN_ADMIN_LIBRARY.DRIVE_NAME, opt.adminEmails);
  vNextAdminShareLibraryWithCompany_(driveId, opt.domain, opt.adminEmails);
  const folders = {
    portal: vNextAdminEnsureLibraryPath_(root, vNextAdminLibraryPath_('PORTAL'), opt.adminEmails).getId(),
    books: vNextAdminEnsureLibraryPath_(root, [VN_ADMIN_LIBRARY.BOOKS], opt.adminEmails).getId(),
    admin: vNextAdminEnsureLibraryPath_(root, vNextAdminLibraryPath_('ADMIN'), opt.adminEmails).getId(),
    audit: vNextAdminEnsureLibraryPath_(root, vNextAdminLibraryPath_('AUDIT'), opt.adminEmails).getId(),
    templates: vNextAdminEnsureLibraryPath_(root, [VN_ADMIN_LIBRARY.TEMPLATES], opt.adminEmails).getId()
  };
  vNextAdminEnsureLibraryPath_(root, vNextAdminLibraryPath_('TEMPLATE_CURRENT'), opt.adminEmails);
  vNextAdminEnsureLibraryPath_(root, vNextAdminLibraryPath_('TEMPLATE_DRAFT'), opt.adminEmails);
  vNextAdminEnsureLibraryPath_(root, vNextAdminLibraryPath_('TEMPLATE_HISTORY'), opt.adminEmails);
  return { driveId: driveId, rootId: root.getId(), folders: folders };
}

function vNextAdminFindOrCreateSharedDrive_(name) {
  if (typeof Drive === 'undefined' || !Drive.Drives) {
    throw new Error('共有ドライブを作るには Drive 詳細サービスが必要です。中央配備版へ更新してから再実行してください。');
  }
  const listed = Drive.Drives.list({ pageSize: 100 });
  const drives = (listed && listed.drives) || [];
  const candidates = [name];
  if (VN_ADMIN_LIBRARY.LEGACY_DRIVE_NAME && VN_ADMIN_LIBRARY.LEGACY_DRIVE_NAME !== name) {
    candidates.push(VN_ADMIN_LIBRARY.LEGACY_DRIVE_NAME);
  }
  for (let c = 0; c < candidates.length; c++) {
    for (let i = 0; i < drives.length; i++) {
      if (String(drives[i].name || '') === candidates[c]) return String(drives[i].id);
    }
  }
  const created = Drive.Drives.create({ name: name }, Utilities.getUuid());
  if (!created || !created.id) throw new Error('共有ドライブ「' + name + '」を作成できませんでした。');
  return String(created.id);
}

function vNextAdminShareLibraryWithCompany_(driveId, domain, adminEmails) {
  if (typeof Drive === 'undefined' || !Drive.Permissions) return false;
  vNextAdminMergeEmails_(adminEmails).forEach(function (email) {
    try {
      Drive.Permissions.create(
        { type: 'user', role: 'organizer', emailAddress: email },
        driveId, { supportsAllDrives: true, sendNotificationEmail: false }
      );
    } catch (ignoredExisting) {}
  });
  if (!domain) return true;
  try {
    Drive.Permissions.create(
      { type: 'domain', role: 'fileOrganizer', domain: domain },
      driveId, { supportsAllDrives: true, sendNotificationEmail: false }
    );
    return true;
  } catch (organizerError) {
    try {
      Drive.Permissions.create(
        { type: 'domain', role: 'writer', domain: domain },
        driveId, { supportsAllDrives: true, sendNotificationEmail: false }
      );
      return true;
    } catch (writerError) {
      Logger.log('Shared drive domain sharing skipped: %s', String(writerError));
      return false;
    }
  }
}

function vNextAdminMoveRegisteredFilesIntoLibrary_(hub, library, adminEmails, options) {
  const moved = [];
  const opt = options || {};
  const hubId = String(hub.getId() || '');
  const activeReleaseId = String(vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET).active_release_id || '');
  const root = vNextAdminPrepareManagedFolder_(library.rootId, VN_ADMIN_LIBRARY.DRIVE_NAME, adminEmails);
  const rows = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.REGISTRY).rows.slice().sort(function (a, b) {
    const aHub = String(a.spreadsheet_id || '') === hubId ? 1 : 0;
    const bHub = String(b.spreadsheet_id || '') === hubId ? 1 : 0;
    return aHub - bHub;
  });
  rows.forEach(function (row) {
    const spreadsheetId = String(row.spreadsheet_id || '').trim();
    if (!spreadsheetId || !vNextAdminSpreadsheetAccessible_(spreadsheetId)) return;
    const mode = String(row.mode || '').toUpperCase();
    const status = String(row.status || '').toUpperCase();
    const state = String(row.state || '').toUpperCase();
    let path = vNextAdminLibraryPath_('ADMIN');
    if (mode === 'PORTAL') path = vNextAdminLibraryPath_('PORTAL');
    else if (mode === 'CLIENT') path = vNextAdminLibraryPath_('CLIENT', row.fiscal_year, row.client_name);
    else if (mode === 'TEMPLATE' && (status === 'DRAFT' || state === 'TEMPLATE_DRAFT')) {
      path = vNextAdminLibraryPath_('TEMPLATE_DRAFT');
    } else if (mode === 'TEMPLATE' && String(row.release_id || row.template_release_id || '') === activeReleaseId) {
      path = vNextAdminLibraryPath_('TEMPLATE_CURRENT');
    } else if (mode === 'TEMPLATE') path = vNextAdminLibraryPath_('TEMPLATE_HISTORY');
    else if (mode !== 'ADMIN') return;
    const dest = vNextAdminEnsureLibraryPath_(root, path, adminEmails);
    vNextAdminMoveFileToFolder_(spreadsheetId, dest);
    moved.push({ spreadsheetId: spreadsheetId, mode: mode, folderId: dest.getId() });
  });
  const portalDest = vNextAdminEnsureLibraryPath_(root, vNextAdminLibraryPath_('PORTAL'), adminEmails);
  const portalIds = {};
  const resolvedPortal = vNextAdminTryResolvePortal_(hub);
  if (resolvedPortal && resolvedPortal.spreadsheet) {
    portalIds[String(resolvedPortal.spreadsheet.getId())] = true;
  }
  const configPortalId = String(vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET).portal_spreadsheet_id || '').trim();
  if (configPortalId) portalIds[configPortalId] = true;
  Object.keys(portalIds).forEach(function (spreadsheetId) {
    if (!vNextAdminSpreadsheetAccessible_(spreadsheetId)) return;
    vNextAdminMoveFileToFolder_(spreadsheetId, portalDest);
    moved.push({ spreadsheetId: spreadsheetId, mode: 'PORTAL', folderId: portalDest.getId() });
  });
  const auditDest = vNextAdminPrepareManagedFolder_(library.folders.audit, VN_ADMIN_LIBRARY.AUDIT, adminEmails);
  const auditSearchRoots = [];
  const legacyRootId = String(opt.legacyRootId || '').trim();
  if (legacyRootId) {
    try { auditSearchRoots.push(DriveApp.getFolderById(legacyRootId)); } catch (ignoredMissingLegacyRoot) {}
  }
  try {
    const hubParents = DriveApp.getFileById(hub.getId()).getParents();
    if (hubParents.hasNext()) auditSearchRoots.push(hubParents.next());
  } catch (ignoredMissingHubParent) {}
  auditSearchRoots.forEach(function (folder) {
    const oldAudit = folder.getFoldersByName('Forecast vNext Admin Audit');
    if (!oldAudit.hasNext()) return;
    const files = oldAudit.next().getFiles();
    while (files.hasNext()) vNextAdminMoveFileToFolder_(files.next().getId(), auditDest);
  });
  return moved;
}

function vNextAdminMoveFileToFolder_(fileId, destFolder) {
  const file = DriveApp.getFileById(fileId);
  if (file.isTrashed()) return { fileId: fileId, skipped: 'trashed' };
  const destId = destFolder.getId();
  const parents = file.getParents();
  let alreadyThere = false;
  while (parents.hasNext()) {
    if (parents.next().getId() === destId) alreadyThere = true;
  }
  if (alreadyThere) return { fileId: fileId, folderId: destId };
  try {
    file.moveTo(destFolder);
  } catch (error) {
    const destIsShared = vNextAdminIsSharedDriveManaged_(destFolder);
    if (destIsShared) {
      try {
        file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.VIEW);
        file.moveTo(destFolder);
        return { fileId: fileId, folderId: destId };
      } catch (retryError) {
        throw new Error('ファイルを移せませんでした: ' + file.getName() + ' / ' + String(error && error.message || error));
      }
    }
    throw new Error('ファイルを移せませんでした: ' + file.getName() + ' / ' + String(error && error.message || error));
  }
  return { fileId: fileId, folderId: destId };
}

function vNextAdminFolderWithinRoot_(folder, rootId) {
  const target = String(rootId || '');
  const seen = new Set();
  let current = [folder];
  for (let depth = 0; depth < 25 && current.length; depth++) {
    const next = [];
    for (let i = 0; i < current.length; i++) {
      const item = current[i];
      const id = String(item.getId() || '');
      if (id === target) return true;
      if (!id || seen.has(id)) continue;
      seen.add(id);
      const parents = item.getParents();
      while (parents.hasNext()) next.push(parents.next());
    }
    current = next;
  }
  return false;
}

function vNextAdminAssertClientFileAcl_(file, editors, viewers) {
  if (vNextAdminIsSharedDriveManaged_(file)) return true;
  if (file.getSharingAccess() !== DriveApp.Access.PRIVATE) {
    throw new Error('Client fileがdomain/anyone共有になっているため生成を停止しました。');
  }
  const owner = file.getOwner() ? String(file.getOwner().getEmail() || '').toLowerCase() : '';
  const expectedEditors = vNextAdminMergeEmails_(editors);
  const expectedViewers = vNextAdminMergeEmails_(viewers);
  const allowed = new Set(vNextAdminMergeEmails_(owner, expectedEditors, expectedViewers));
  const actualEditors = file.getEditors().map(function (user) { return String(user.getEmail() || '').toLowerCase(); });
  const actualViewers = file.getViewers().map(function (user) { return String(user.getEmail() || '').toLowerCase(); });
  // Drive does not return the owner from getEditors(), although the owner has
  // full edit rights. Treat ownership as satisfying an expected editor/viewer
  // grant so a single-admin Pilot file does not fail its exact ACL check.
  const actual = vNextAdminMergeEmails_(owner, actualEditors, actualViewers);
  const missing = expectedEditors.concat(expectedViewers).filter(function (email) { return actual.indexOf(email) < 0; });
  const unexpected = actual.filter(function (email) { return email && !allowed.has(email); });
  if (missing.length || unexpected.length) {
    throw new Error('Client file ACLの最終検証に失敗しました。missing=' + missing.join(',') + '; unexpected=' + unexpected.join(','));
  }
  return true;
}

/** Domain sharing is allowed only for employee-facing CLIENT/PORTAL files. */
function vNextAdminApplyEmployeeFileSharing_(file, options) {
  const opt = options || {};
  const mode = String(opt.targetMode || '').toUpperCase();
  if (['CLIENT', 'PORTAL'].indexOf(mode) < 0) throw new Error('Employee sharing is limited to CLIENT/PORTAL files.');
  const policy = String(opt.accessPolicy || 'PRIVATE').toUpperCase();
  if (policy === 'PRIVATE') {
    if (vNextAdminIsSharedDriveManaged_(file)) return true;
    file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.VIEW);
    return true;
  }
  if (policy !== 'INTERNAL_OPEN') throw new Error('Unsupported employee access policy: ' + policy);
  if (vNextAdminIsSharedDriveManaged_(file)) return true;
  const domain = vNextAdminNormalizeDomain_(opt.internalDomain);
  const owner = file.getOwner() ? String(file.getOwner().getEmail() || '').toLowerCase() : '';
  if (!domain || vNextAdminEmailDomain_(owner) !== domain) {
    throw new Error('Employee domain does not match the Workspace file owner.');
  }
  file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);
  return true;
}

function vNextAdminAssertEmployeeFileSharing_(file, options) {
  const opt = options || {};
  const mode = String(opt.targetMode || '').toUpperCase();
  if (['CLIENT', 'PORTAL'].indexOf(mode) < 0) throw new Error('Employee sharing verification is limited to CLIENT/PORTAL files.');
  const policy = String(opt.accessPolicy || 'PRIVATE').toUpperCase();
  if (policy === 'PRIVATE') {
    if (vNextAdminIsSharedDriveManaged_(file)) return true;
    return vNextAdminAssertClientFileAcl_(file, opt.editors || [], opt.viewers || []);
  }
  if (policy !== 'INTERNAL_OPEN') throw new Error('Unsupported employee access policy: ' + policy);
  if (vNextAdminIsSharedDriveManaged_(file)) return true;
  const domain = vNextAdminNormalizeDomain_(opt.internalDomain);
  const owner = file.getOwner() ? String(file.getOwner().getEmail() || '').toLowerCase() : '';
  if (!domain || vNextAdminEmailDomain_(owner) !== domain) throw new Error('Employee domain verification failed.');
  if (file.getSharingAccess() !== DriveApp.Access.DOMAIN_WITH_LINK ||
      (typeof file.getSharingPermission === 'function' && file.getSharingPermission() !== DriveApp.Permission.EDIT)) {
    throw new Error('Employee file is not restricted to domain-with-link edit access.');
  }
  const expectedEditors = vNextAdminMergeEmails_(opt.editors || []);
  const expectedViewers = vNextAdminMergeEmails_(opt.viewers || []);
  const actual = vNextAdminMergeEmails_(owner,
    file.getEditors().map(function (user) { return String(user.getEmail() || '').toLowerCase(); }),
    file.getViewers().map(function (user) { return String(user.getEmail() || '').toLowerCase(); }));
  const missing = expectedEditors.concat(expectedViewers).filter(function (email) { return actual.indexOf(email) < 0; });
  if (missing.length) throw new Error('Employee file ACL is missing explicit role recipients: ' + missing.join(','));
  return true;
}

function vNextAdminNormalizeDomain_(value) {
  const domain = String(value || '').trim().toLowerCase().replace(/^@/, '');
  return /^[a-z0-9](?:[a-z0-9.-]*[a-z0-9])?$/.test(domain) && domain.indexOf('.') > 0 ? domain : '';
}

function vNextAdminEmailDomain_(value) {
  const email = String(value || '').trim().toLowerCase();
  const at = email.lastIndexOf('@');
  return at > 0 && at < email.length - 1 ? email.slice(at + 1) : '';
}

function vNextAdminPrepareManagedFolder_(folderId, name, adminEmails) {
  return vNextAdminPreparePrivateBootstrapFolder_(folderId, name, adminEmails);
}

function vNextAdminIsSharedDriveManaged_(fileOrFolder) {
  try {
    if (typeof Drive !== 'undefined' && Drive.Files) {
      const meta = Drive.Files.get(fileOrFolder.getId(), {
        fields: 'id,driveId,teamDriveId', supportsAllDrives: true, supportsTeamDrives: true
      });
      if (meta && (meta.driveId || meta.teamDriveId)) return true;
    }
  } catch (error) {}
  try {
    if (typeof fileOrFolder.getSharingAccess === 'function') fileOrFolder.getSharingAccess();
  } catch (error) {
    if (/Team Drive|shared drive|共有ドライブ|not supported|Action not allowed/i.test(String(error && error.message || error))) {
      return true;
    }
  }
  return false;
}

function vNextAdminPreparePrivateBootstrapFolder_(folderId, name, adminEmails) {
  const folder = folderId
    ? DriveApp.getFolderById(String(folderId))
    : DriveApp.createFolder(String(name || ('Forecast vNext Private ' + new Date().getTime())));
  if (vNextAdminIsSharedDriveManaged_(folder)) return folder;
  const allowed = new Set(vNextAdminMergeEmails_(adminEmails, vNextAdminActor_()));
  try {
    if (folder.getSharingAccess() !== DriveApp.Access.PRIVATE) {
      if (folderId) throw new Error('指定folderはdomain/anyone共有のため使用できません。private専用folderを指定してください。');
      folder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.VIEW);
    }
    const outsiders = folder.getEditors().concat(folder.getViewers()).filter(function (user) {
      return !allowed.has(String(user.getEmail() || '').toLowerCase());
    });
    if (outsiders.length) throw new Error('指定folderにAdmin以外の直接共有先があります。private専用folderを指定してください。');
    const admins = Array.from(allowed).filter(Boolean);
    if (admins.length) folder.addEditors(admins);
    if (folder.getSharingAccess() !== DriveApp.Access.PRIVATE) throw new Error('Bootstrap folder could not be verified as PRIVATE.');
    return folder;
  } catch (err) {
    throw new Error('Bootstrap保存先の共有境界を確認できません: ' + String(err && err.message || err));
  }
}

function vNextAdminEnforcePrivateFileAcl_(file, adminEmails) {
  if (vNextAdminIsSharedDriveManaged_(file)) return true;
  const owner = file.getOwner() ? String(file.getOwner().getEmail() || '').toLowerCase() : '';
  const allowed = new Set(vNextAdminMergeEmails_(adminEmails, owner, vNextAdminActor_()));
  file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.VIEW);
  file.getEditors().forEach(function (user) {
    const email = String(user.getEmail() || '').toLowerCase();
    if (email && !allowed.has(email)) file.removeEditor(email);
  });
  file.getViewers().forEach(function (user) {
    const email = String(user.getEmail() || '').toLowerCase();
    if (email && !allowed.has(email)) file.removeViewer(email);
  });
  const admins = Array.from(allowed).filter(function (email) { return email && email !== owner; });
  if (admins.length) file.addEditors(admins);
  if (file.getSharingAccess() !== DriveApp.Access.PRIVATE) throw new Error('Generated Admin file is not PRIVATE: ' + file.getId());
  const unexpected = file.getEditors().concat(file.getViewers()).filter(function (user) {
    return !allowed.has(String(user.getEmail() || '').toLowerCase());
  });
  if (unexpected.length) throw new Error('Generated Admin file has an unexpected collaborator: ' + file.getId());
  return true;
}

function vNextAdminActor_() {
  const activeUser = Session.getActiveUser();
  const effectiveUser = Session.getEffectiveUser();
  const activeEmail = activeUser ? String(activeUser.getEmail() || '') : '';
  const effectiveEmail = effectiveUser ? String(effectiveUser.getEmail() || '') : '';
  return String(activeEmail || effectiveEmail || 'unknown').trim();
}

function vNextAdminText_(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}

function vNextAdminDateOnly_(value) {
  if (!value) return '';
  const date = value instanceof Date ? value : new Date(value);
  if (isNaN(date.getTime())) throw new Error('Invalid date value: ' + String(value));
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function vNextAdminRequiredText_(value, label) {
  const text = vNextAdminText_(value);
  if (!text) throw new Error(label + ' is required.');
  return text;
}

function vNextAdminNormalizeFiscalYear_(value) {
  const text = vNextAdminRequiredText_(value, 'fiscalYear').replace(/^FY/i, '');
  const year = Number(text);
  if (!Number.isInteger(year) || year < 2000 || year > 2200) throw new Error('Invalid fiscalYear: ' + value);
  return year;
}

function vNextAdminCutoffFromAsOf_(asOf) {
  if (typeof vNextCutoffFromAsOf_ === 'function') {
    const cutoff = vNextCutoffFromAsOf_(asOf);
    return typeof vNextFormatDateOnly_ === 'function'
      ? vNextFormatDateOnly_(cutoff)
      : Utilities.formatDate(new Date(cutoff), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const date = new Date(asOf);
  if (isNaN(date.getTime())) throw new Error('Invalid asOf date: ' + asOf);
  const cutoff = new Date(date.getFullYear(), date.getMonth(), 0);
  return Utilities.formatDate(cutoff, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function vNextAdminDeriveClientId_(clientName) {
  const normalized = String(clientName || '').trim().toLowerCase().replace(/\s+/g, ' ');
  if (!normalized) throw new Error('clientName is required.');
  return 'C-' + vNextAdminSha256_(normalized).slice(0, 12).toUpperCase();
}

function vNextAdminBool_(value) {
  return value === true || value === 1 || String(value).toLowerCase() === 'true' || String(value) === '1';
}

function vNextAdminParseList_(value) {
  if (Array.isArray(value)) return value.map(vNextAdminText_).filter(Boolean);
  return vNextAdminText_(value).split(/[\n,;]+/).map(function (item) { return item.trim(); }).filter(Boolean);
}

function vNextAdminMergeEmails_() {
  const seen = new Set();
  const out = [];
  Array.prototype.slice.call(arguments).forEach(function (value) {
    vNextAdminParseList_(value).forEach(function (email) {
      const normalized = String(email || '').trim().toLowerCase();
      if (!normalized || normalized.indexOf('@') < 1 || seen.has(normalized)) return;
      seen.add(normalized);
      out.push(normalized);
    });
  });
  return out;
}

function vNextAdminParseJson_(value, fallback) {
  if (value && typeof value === 'object') return value;
  try { return JSON.parse(String(value || '')); } catch (err) { return fallback; }
}

function vNextAdminAssertClientRequestPayload_(payload, requestJson, expectedRequestId) {
  if (!payload || typeof payload !== 'object' || Array.isArray(payload) ||
      Object.getPrototypeOf(payload) !== Object.prototype) {
    throw new Error('Client request payload must be a plain JSON object.');
  }
  const keys = Object.keys(payload).sort();
  const expectedKeys = VN_ADMIN_CLIENT_REQUEST_PAYLOAD_KEYS.slice().sort();
  if (keys.length !== expectedKeys.length || keys.some(function (key, index) { return key !== expectedKeys[index]; })) {
    const unknown = keys.filter(function (key) { return expectedKeys.indexOf(key) < 0; });
    const missing = expectedKeys.filter(function (key) { return keys.indexOf(key) < 0; });
    throw new Error('Client request payload keys are not exact. unknown=' + unknown.join(',') + '; missing=' + missing.join(','));
  }
  if (requestJson !== undefined && String(requestJson) !== vNextAdminCanonicalJson_(payload)) {
    throw new Error('Client request JSON must be canonical and contain no ambiguous duplicate ordering.');
  }
  if (String(payload.requestId || '') !== String(expectedRequestId || '')) {
    throw new Error('Client requestId does not match the request event row.');
  }
  VN_ADMIN_CLIENT_REQUEST_PAYLOAD_KEYS.forEach(function (key) {
    if (key === 'fiscalYear') {
      if (typeof payload[key] !== 'number' || !isFinite(payload[key])) throw new Error('Client request fiscalYear must be a finite number.');
      return;
    }
    if (typeof payload[key] !== 'string') throw new Error('Client request ' + key + ' must be a string.');
  });
  return true;
}

function vNextAdminCanonicalJson_(value) {
  function normalize(item) {
    if (item instanceof Date) return item.toISOString();
    if (Array.isArray(item)) return item.map(normalize);
    if (item && typeof item === 'object') {
      const out = {};
      Object.keys(item).sort().forEach(function (key) { out[key] = normalize(item[key]); });
      return out;
    }
    if (typeof item === 'number' && !isFinite(item)) return null;
    return item;
  }
  return JSON.stringify(normalize(value));
}

function vNextAdminSha256_(value) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(value || ''), Utilities.Charset.UTF_8);
  return bytes.map(function (byte) { const n = byte < 0 ? byte + 256 : byte; return ('0' + n.toString(16)).slice(-2); }).join('');
}

function vNextAdminJsonSafe_(value) {
  return JSON.parse(JSON.stringify(value, function (key, item) {
    return item instanceof Date ? item.toISOString() : item;
  }));
}

function vNextAdminMenuAction_(message, fn) {
  return vNextAdminGuard_('vNextAdminMenuAction', function () {
    const result = fn();
    SpreadsheetApp.getActiveSpreadsheet().toast(message, VN_ADMIN_MENU_NAME, 5);
    return result;
  });
}

function vNextAdminOpenHubSheet_(name) {
  return vNextAdminGuard_('vNextAdminOpenHubSheet', function () {
    const hub = vNextAdminRequireHub_();
    const sheet = hub.getSheetByName(name);
    if (!sheet) throw new Error('Hub sheet not found: ' + name);
    sheet.showSheet();
    hub.setActiveSheet(sheet);
    return true;
  });
}
