/**
 * Schema.js — 表の定義（スキーマ登録）。表の列はここだけで決め、データ本体の 1 行目と一致しなければ書き込まない。
 * 1 表 = 1 シート、1 行目に英字の列名、2 行目から値だけ（数式・タイトル・説明・色は置かない）。
 * 段階1 はマスタと設定の表。予測・予算・評価などの表は段階2・3 で足す（版を上げ、移行の手順で作る）。
 */
const APP_SCHEMA_VERSION = 1;

const APP_TABLES = {
  _SCHEMA: {
    key: ['table'],
    columns: ['table', 'schema_version', 'columns_hash', 'migrated_at', 'migrated_by']
  },
  MEMBERS: {
    key: ['email'],
    columns: ['email', 'display_name', 'department', 'is_active', 'note',
      'created_at', 'created_by', 'updated_at', 'updated_by', 'row_version']
  },
  ROLES: {
    key: ['role_id'],
    columns: ['role_id', 'email', 'role', 'scope_type', 'client_id', 'valid_from', 'valid_to', 'is_active', 'note',
      'created_at', 'created_by', 'updated_at', 'updated_by', 'row_version']
  },
  CLIENTS: {
    key: ['client_id'],
    columns: ['client_id', 'client_name', 'zac_code', 'normalized_name', 'aliases_json', 'is_active', 'note',
      'created_at', 'created_by', 'updated_at', 'updated_by', 'row_version']
  },
  SETTINGS: {
    // 設定は上書きせず、変更のたびに 1 行足す（effective_from が今日以前で最も新しい行が今の値）
    key: ['setting_id'],
    columns: ['setting_id', 'key', 'value', 'scope', 'scope_id', 'effective_from', 'note', 'created_at', 'created_by']
  }
};

/** ログのファイルの月別シート（AUDIT_yyyy_MM など） */
const APP_LOG_TABLES = {
  AUDIT: ['audit_id', 'occurred_at', 'actor_email', 'actor_roles', 'action', 'phase', 'result', 'entity_type', 'entity_id',
    'client_id', 'plan_id', 'detail_json', 'before_json', 'after_json', 'reason', 'error', 'request_id', 'app_version',
    'prev_hash', 'row_hash'],
  RUN: ['run_id', 'request_id', 'kind', 'started_at', 'finished_at', 'duration_ms', 'status', 'detail_json', 'error'],
  ERROR: ['error_id', 'occurred_at', 'request_id', 'actor_email', 'where', 'message', 'stack_head']
};

/** 列の型（保存はすべて書式なしテキスト。読むときにここで戻す） */
function appColumnType_(col) {
  if (/_json$/.test(col)) return 'json';
  if (col === 'is_active') return 'bool';
  if (col === 'row_version' || col === 'schema_version') return 'int';
  return 'text';
}

/** 表の列の並びのハッシュ（_SCHEMA に控え、定義と実物のずれを見つける） */
function appColumnsHash_(name) {
  return appSha256Hex_(APP_TABLES[name].columns.join('|'));
}
