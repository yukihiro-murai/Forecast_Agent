/**
 * Forecast vNext core: append-only audit store, canonical hashing, roles and state.
 * V8 / pure JavaScript compatible. No legacy sheet is mutated by this module.
 */

var VNEXT_CORE = Object.freeze({
  VERSION: 'vnext-0.2.0',
  SCHEMA_VERSION: 'vnext-schema-2',
  DEFAULT_TEMPLATE_VERSION: 'vnext-template-1',
  HUB_PROPERTY: 'VNEXT_ADMIN_HUB_SPREADSHEET_ID',
  BOOK_PROPERTY: 'VNEXT_BOOK_ID',
  ADMIN_EMAILS_PROPERTY: 'VNEXT_ADMIN_EMAILS',
  INTERNAL_SHEETS: Object.freeze({
    BOOK_META: Object.freeze([
      'record_id', 'book_id', 'client_id', 'client_name', 'fiscal_year',
      'forecast_owner_email', 'team_member_emails_json', 'state', 'as_of',
      'cutoff', 'template_version', 'schema_version', 'model_release_id',
      'source_spreadsheet_id', 'client_book_id', 'input_due_date', 'event_type',
      'supersedes_record_id', 'recorded_at', 'recorded_by'
    ]),
    EVIDENCE_EVENT: Object.freeze([
      'evidence_id', 'book_id', 'client_id', 'fiscal_year', 'actor_email',
      'response_type', 'evidence_type', 'target', 'target_start_month',
      'target_end_month', 'direction', 'amount_mode', 'amount_low', 'amount_mid',
      'amount_high', 'amount_band', 'confidence_class', 'evidence_text',
      'source_url', 'source_date', 'expires_at', 'status',
      'supersedes_evidence_id', 'created_at', 'evidence_quality', 'ai_model',
      'prompt_version', 'ai_schema_version', 'rule_version', 'applied_amount',
      'cap_applied'
    ]),
    FORECAST_RUN: Object.freeze([
      'run_id', 'book_id', 'client_id', 'client_name', 'fiscal_year', 'as_of',
      'cutoff', 'seed', 'input_data_hash', 'model_release_id', 'schema_version',
      'status', 'official_vintage_id', 'is_official', 'history_years',
      'simulation_count', 'history_baseline', 'commitment_delta',
      'reference_delta', 'human_delta', 'ai_delta', 'objective_forecast',
      'system_recommended', 'p10', 'p50', 'p90', 'quarter_json', 'month_json',
      'lens_json', 'evidence_json', 'previous_run_id', 'created_at', 'created_by',
      'error_summary'
    ]),
    PLAN_VERSION: Object.freeze([
      'plan_version_id', 'book_id', 'run_id', 'official_vintage_id',
      'version_no', 'status', 'system_recommended', 'adoption_delta',
      'adoption_reason', 'adopted_forecast', 'sales_uplift', 'uplift_reason',
      'uplift_owner', 'uplift_action', 'uplift_due_date',
      'uplift_allocation_json', 'final_budget', 'amends_plan_version_id',
      'submitted_at', 'submitted_by', 'approved_at', 'approved_by', 'created_at'
    ]),
    STATE_EVENT: Object.freeze([
      'state_event_id', 'book_id', 'from_state', 'to_state', 'reason',
      'actor_email', 'actor_role', 'related_run_id', 'related_plan_version_id',
      'created_at'
    ]),
    EVALUATION: Object.freeze([
      'evaluation_id', 'book_id', 'official_vintage_id', 'source_run_id',
      'fiscal_year', 'evaluated_at', 'actual_total', 'system_forecast',
      'adopted_forecast', 'final_budget', 'system_signed_error',
      'system_abs_error', 'system_ape', 'range_contains_actual',
      'base_level_error', 'seasonality_error', 'commitment_outcome_error',
      'amount_error', 'timing_error', 'unknown_spot_error', 'human_info_error',
      'ai_info_error', 'data_quality_error', 'confirmed_cause',
      'cause_hypothesis', 'next_information_json', 'model_release_id',
      'created_by'
    ]),
    MODEL_RELEASE: Object.freeze([
      'model_release_id', 'status', 'model_version', 'schema_version',
      'template_version', 'parameters_json', 'backtest_json', 'canary_json',
      'approved_at', 'approved_by', 'rollback_release_id', 'created_at',
      'created_by', 'note'
    ])
  }),
  STATES: Object.freeze([
    'INPUT_OPEN', 'READY_TO_RUN', 'RUNNING', 'DRAFT_READY', 'SUBMITTED',
    'CHANGES_REQUESTED', 'OFFICIAL_LOCKED', 'REVIEW_DUE', 'YEAR_CLOSED'
  ]),
  RESPONSE_TYPES: Object.freeze(['CHANGE', 'NO_CHANGE', 'UNKNOWN']),
  CONFIDENCE_CLASSES: Object.freeze(['CONFIRMED_FACT', 'LIKELY', 'HYPOTHESIS']),
  STATE_TRANSITIONS: Object.freeze({
    INPUT_OPEN: Object.freeze(['READY_TO_RUN']),
    READY_TO_RUN: Object.freeze(['RUNNING', 'INPUT_OPEN']),
    RUNNING: Object.freeze(['DRAFT_READY', 'READY_TO_RUN']),
    DRAFT_READY: Object.freeze(['SUBMITTED', 'READY_TO_RUN']),
    SUBMITTED: Object.freeze(['OFFICIAL_LOCKED', 'CHANGES_REQUESTED']),
    CHANGES_REQUESTED: Object.freeze(['INPUT_OPEN', 'READY_TO_RUN', 'SUBMITTED']),
    OFFICIAL_LOCKED: Object.freeze(['REVIEW_DUE']),
    REVIEW_DUE: Object.freeze(['YEAR_CLOSED']),
    YEAR_CLOSED: Object.freeze([])
  })
});

/** Creates missing append-only audit sheets without touching existing data. */
function vNextEnsureAuditStore_(spreadsheet) {
  try {
    var ss = spreadsheet || vNextResolveStoreSpreadsheet_();
    var created = [];
    Object.keys(VNEXT_CORE.INTERNAL_SHEETS).forEach(function (sheetName) {
      var headers = VNEXT_CORE.INTERNAL_SHEETS[sheetName];
      var sh = ss.getSheetByName(sheetName);
      if (!sh) {
        sh = ss.insertSheet(sheetName);
        created.push(sheetName);
      }
      vNextEnsureAppendOnlyHeader_(sh, headers);
      try {
        sh.setFrozenRows(1);
        sh.hideSheet();
      } catch (hideError) {
        vNextLog_('Internal sheet visibility could not be changed: ' + sheetName, hideError);
      }
    });
    return { spreadsheetId: vNextSpreadsheetId_(ss), created: created };
  } catch (error) {
    vNextLog_('vNextEnsureAuditStore_ failed', error);
    throw error;
  }
}

function vNextEnsureAppendOnlyHeader_(sheet, headers) {
  if (!sheet || !headers || !headers.length) throw new Error('Sheet and headers are required.');
  var current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  var blank = current.every(function (value) { return String(value || '').trim() === ''; });
  if (blank) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers.slice()]);
    try {
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#d9eaf7');
    } catch (formatError) {
      vNextLog_('Header formatting skipped: ' + sheet.getName(), formatError);
    }
    return;
  }
  for (var i = 0; i < headers.length; i++) {
    if (String(current[i] || '') !== headers[i]) {
      throw new Error(
        sheet.getName() + ' header mismatch at column ' + (i + 1) +
        '. Existing append-only data was not changed.'
      );
    }
  }
}

/** Bulk append. Existing rows are never updated or cleared. */
function vNextAppendRecords_(sheetName, records, options) {
  var opt = options || {};
  if (!records || !records.length) return { appended: 0, firstRow: 0 };
  var headers = VNEXT_CORE.INTERNAL_SHEETS[sheetName];
  if (!headers) throw new Error('Unknown vNext audit sheet: ' + sheetName);
  return vNextWithScriptLock_(function () {
    var ss = opt.spreadsheet || vNextResolveStoreSpreadsheet_(opt);
    if (opt.ensureStore !== false) vNextEnsureAuditStore_(ss);
    var sh = ss.getSheetByName(sheetName);
    if (!sh) throw new Error(sheetName + ' is not initialized.');
    vNextEnsureAppendOnlyHeader_(sh, headers);
    var rows = records.map(function (record) {
      return headers.map(function (header) {
        return vNextValueForSheet_(record ? record[header] : '');
      });
    });
    var firstRow = sh.getLastRow() + 1;
    sh.getRange(firstRow, 1, rows.length, headers.length).setValues(rows);
    return { appended: rows.length, firstRow: firstRow };
  }, opt.lockTimeoutMs);
}

function vNextAppendRecord_(sheetName, record, options) {
  return vNextAppendRecords_(sheetName, [record], options);
}

function vNextReadRecords_(sheetName, options) {
  var opt = options || {};
  var headers = VNEXT_CORE.INTERNAL_SHEETS[sheetName];
  if (!headers) throw new Error('Unknown vNext audit sheet: ' + sheetName);
  var ss = opt.spreadsheet || vNextResolveStoreSpreadsheet_(opt);
  var sh = ss.getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return [];
  vNextEnsureAppendOnlyHeader_(sh, headers);
  var values = sh.getRange(2, 1, sh.getLastRow() - 1, headers.length).getValues();
  return values.map(function (row) {
    var record = {};
    headers.forEach(function (header, index) { record[header] = row[index]; });
    return record;
  });
}

/** Stable canonical JSON used by input_data_hash. */
function vNextCanonicalJson_(value) {
  var seen = [];
  function encode(current, inArray) {
    if (current === null) return 'null';
    var type = typeof current;
    if (type === 'string') return JSON.stringify(current);
    if (type === 'boolean') return current ? 'true' : 'false';
    if (type === 'number') {
      if (!isFinite(current)) throw new Error('Canonical JSON rejects non-finite numbers.');
      return String(Object.is(current, -0) ? 0 : current);
    }
    if (type === 'undefined') return inArray ? 'null' : undefined;
    if (type === 'function' || type === 'symbol' || type === 'bigint') {
      throw new Error('Canonical JSON rejects type: ' + type);
    }
    if (current instanceof Date) {
      if (isNaN(current.getTime())) throw new Error('Canonical JSON rejects invalid Date.');
      return JSON.stringify(current.toISOString());
    }
    if (seen.indexOf(current) >= 0) throw new Error('Canonical JSON rejects circular values.');
    seen.push(current);
    var encoded;
    if (Array.isArray(current)) {
      encoded = '[' + current.map(function (item) {
        var result = encode(item, true);
        return result === undefined ? 'null' : result;
      }).join(',') + ']';
    } else {
      var keys = Object.keys(current).sort();
      var parts = [];
      keys.forEach(function (key) {
        var item = encode(current[key], false);
        if (item !== undefined) parts.push(JSON.stringify(key) + ':' + item);
      });
      encoded = '{' + parts.join(',') + '}';
    }
    seen.pop();
    return encoded;
  }
  return encode(value, false);
}

function vNextSha256Hex_(value) {
  var input = typeof value === 'string' ? value : vNextCanonicalJson_(value);
  if (typeof Utilities !== 'undefined' && Utilities.computeDigest) {
    var digest = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      input,
      Utilities.Charset.UTF_8
    );
    return digest.map(function (byte) {
      var unsigned = byte < 0 ? byte + 256 : byte;
      return ('0' + unsigned.toString(16)).slice(-2);
    }).join('');
  }
  return vNextSha256PureJs_(input);
}

/** Small SHA-256 fallback for local/pure-JS tests. */
function vNextSha256PureJs_(text) {
  var utf8 = unescape(encodeURIComponent(String(text)));
  var words = [];
  var bitLength = utf8.length * 8;
  for (var i = 0; i < utf8.length; i++) {
    words[i >> 2] = (words[i >> 2] || 0) | (utf8.charCodeAt(i) << (24 - (i % 4) * 8));
  }
  words[bitLength >> 5] = (words[bitLength >> 5] || 0) | (0x80 << (24 - bitLength % 32));
  words[(((bitLength + 64) >> 9) << 4) + 15] = bitLength;
  var constants = [
    0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b,
    0x59f111f1, 0x923f82a4, 0xab1c5ed5, 0xd807aa98, 0x12835b01,
    0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7,
    0xc19bf174, 0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc,
    0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da, 0x983e5152,
    0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147,
    0x06ca6351, 0x14292967, 0x27b70a85, 0x2e1b2138, 0x4d2c6dfc,
    0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
    0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819,
    0xd6990624, 0xf40e3585, 0x106aa070, 0x19a4c116, 0x1e376c08,
    0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f,
    0x682e6ff3, 0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208,
    0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2
  ];
  var hash = [
    0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a,
    0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19
  ];
  for (var offset = 0; offset < words.length; offset += 16) {
    var w = new Array(64);
    for (var j = 0; j < 16; j++) w[j] = words[offset + j] | 0;
    for (j = 16; j < 64; j++) {
      var s0 = vNextRotateRight_(w[j - 15], 7) ^ vNextRotateRight_(w[j - 15], 18) ^ (w[j - 15] >>> 3);
      var s1 = vNextRotateRight_(w[j - 2], 17) ^ vNextRotateRight_(w[j - 2], 19) ^ (w[j - 2] >>> 10);
      w[j] = (w[j - 16] + s0 + w[j - 7] + s1) | 0;
    }
    var a = hash[0], b = hash[1], c = hash[2], d = hash[3];
    var e = hash[4], f = hash[5], g = hash[6], h = hash[7];
    for (j = 0; j < 64; j++) {
      var bigS1 = vNextRotateRight_(e, 6) ^ vNextRotateRight_(e, 11) ^ vNextRotateRight_(e, 25);
      var choice = (e & f) ^ (~e & g);
      var temp1 = (h + bigS1 + choice + constants[j] + w[j]) | 0;
      var bigS0 = vNextRotateRight_(a, 2) ^ vNextRotateRight_(a, 13) ^ vNextRotateRight_(a, 22);
      var majority = (a & b) ^ (a & c) ^ (b & c);
      var temp2 = (bigS0 + majority) | 0;
      h = g; g = f; f = e; e = (d + temp1) | 0;
      d = c; c = b; b = a; a = (temp1 + temp2) | 0;
    }
    hash[0] = (hash[0] + a) | 0; hash[1] = (hash[1] + b) | 0;
    hash[2] = (hash[2] + c) | 0; hash[3] = (hash[3] + d) | 0;
    hash[4] = (hash[4] + e) | 0; hash[5] = (hash[5] + f) | 0;
    hash[6] = (hash[6] + g) | 0; hash[7] = (hash[7] + h) | 0;
  }
  return hash.map(function (part) { return ('00000000' + (part >>> 0).toString(16)).slice(-8); }).join('');
}

function vNextRotateRight_(value, amount) {
  return (value >>> amount) | (value << (32 - amount));
}

/** Mulberry32: deterministic and sufficient for auditable simulation. */
function vNextCreatePrng_(seed) {
  var state = vNextNormalizeSeed_(seed);
  return function () {
    state |= 0;
    state = (state + 0x6D2B79F5) | 0;
    var t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function vNextNormalizeSeed_(seed) {
  if (typeof seed === 'number' && isFinite(seed)) return seed >>> 0;
  var text = String(seed === undefined || seed === null ? '' : seed);
  var hash = 2166136261;
  for (var i = 0; i < text.length; i++) {
    hash ^= text.charCodeAt(i);
    hash = Math.imul(hash, 16777619);
  }
  return hash >>> 0;
}

function vNextCutoffFromAsOf_(asOf) {
  var date = vNextParseDate_(asOf, 'as_of');
  return new Date(date.getFullYear(), date.getMonth(), 0, 23, 59, 59, 999);
}

function vNextParseDate_(value, fieldName) {
  if (value instanceof Date && !isNaN(value.getTime())) return new Date(value.getTime());
  var text = String(value || '').trim();
  var match = text.match(/^(\d{4})[-\/]?(\d{2})[-\/]?(\d{2})/);
  var date = match
    ? new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]))
    : new Date(text);
  if (isNaN(date.getTime())) throw new Error((fieldName || 'date') + ' is invalid: ' + text);
  return date;
}

function vNextFormatDateOnly_(value) {
  var date = vNextParseDate_(value, 'date');
  return [date.getFullYear(), ('0' + (date.getMonth() + 1)).slice(-2), ('0' + date.getDate()).slice(-2)].join('-');
}

function vNextFormatMonth_(value) {
  var date = vNextParseDate_(value, 'month');
  return date.getFullYear() + '-' + ('0' + (date.getMonth() + 1)).slice(-2);
}

function vNextAddMonths_(value, count) {
  var date = vNextParseDate_(value, 'date');
  return new Date(date.getFullYear(), date.getMonth() + Number(count || 0), 1);
}

function vNextFiscalYearForDate_(value) {
  var date = vNextParseDate_(value, 'date');
  return date.getMonth() >= 3 ? date.getFullYear() : date.getFullYear() - 1;
}

function vNextUuid_() {
  if (typeof Utilities !== 'undefined' && Utilities.getUuid) return Utilities.getUuid();
  var random = Math.floor(Math.random() * 0x100000000).toString(16);
  return 'local-' + new Date().getTime().toString(36) + '-' + ('00000000' + random).slice(-8);
}

function vNextNowIso_() {
  return new Date().toISOString();
}

function vNextActiveUserEmail_() {
  try {
    if (typeof Session !== 'undefined' && Session.getActiveUser) {
      var active = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
      if (active) return active;
      if (Session.getEffectiveUser) return String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
    }
  } catch (error) {
    vNextLog_('Active user email unavailable', error);
  }
  return '';
}

function vNextGetBookContext_(options) {
  try {
    var directSpreadsheet = options && typeof options.getSheetByName === 'function' ? options : null;
    var opt = directSpreadsheet ? {} : (options || {});
    var localRouting = directSpreadsheet ? vNextReadLocalRoutingConfig_(directSpreadsheet) : {};
    var activeBookId = opt.bookId || localRouting.book_id || vNextGetProperty_(VNEXT_CORE.BOOK_PROPERTY, true) || (directSpreadsheet ? vNextSpreadsheetId_(directSpreadsheet) : '');
    var directIsStore = directSpreadsheet && vNextIsAuditStoreSpreadsheet_(directSpreadsheet);
    var store;
    var metaRows;
    try {
      store = opt.spreadsheet || (directIsStore ? directSpreadsheet : vNextResolveStoreSpreadsheet_({ clientSpreadsheet: directSpreadsheet }));
      metaRows = vNextReadRecords_('BOOK_META', { spreadsheet: store });
    } catch (storeError) {
      if (String(localRouting.mode || '').toUpperCase() === 'CLIENT') {
        return vNextBuildLocalClientContext_(directSpreadsheet, localRouting, opt, storeError);
      }
      throw storeError;
    }
    var metas = metaRows.filter(function (row) {
      return !activeBookId || String(row.book_id || '') === String(activeBookId);
    });
    if (!metas.length && store) {
      var registryMeta = vNextFindRegistryBookMeta_(store, activeBookId);
      if (registryMeta) metas.push(registryMeta);
    }
    if (!metas.length && directIsStore && !vNextGetProperty_(VNEXT_CORE.BOOK_PROPERTY, true)) {
      var adminEmail = String(opt.userEmail || vNextActiveUserEmail_()).toLowerCase();
      var adminRole = vNextResolveRole_(adminEmail, {}, {});
      if (adminRole !== 'ADMIN') throw new Error('管理ハブへのアクセスは管理ハブ担当者に限定されています。');
      return {
        mode: 'ADMIN_HUB',
        bookId: '',
        clientId: '',
        clientName: '',
        fiscalYear: '',
        asOf: '',
        cutoff: '',
        state: '',
        role: adminRole,
        isForecastOwner: false,
        forecastOwnerEmails: [],
        userEmail: adminEmail,
        inputStatus: { submitted: false, answeredCount: 0, totalCount: 0, dueDate: '' },
        canProceed: false,
        latestOwnEvidence: null,
        version: { core: VNEXT_CORE.VERSION, schema: VNEXT_CORE.SCHEMA_VERSION }
      };
    }
    if (!metas.length) throw new Error('BOOK_META not found for book: ' + (activeBookId || '(unset)'));
    var meta = metas[metas.length - 1];
    var bookId = String(meta.book_id || activeBookId || '');
    var stateRows = vNextReadRecords_('STATE_EVENT', { spreadsheet: store }).filter(function (row) {
      return String(row.book_id || '') === bookId;
    });
    // STATE_EVENT is append-only and authoritative. VN_BOOK_CONFIG is only a
    // routing/display cache because its mirror write can fail after the event
    // has already committed.
    var state = stateRows.length
      ? String(stateRows[stateRows.length - 1].to_state || '')
      : String(localRouting.state || meta.state || 'INPUT_OPEN');
    var userEmail = String(opt.userEmail || vNextActiveUserEmail_()).toLowerCase();
    var role = vNextResolveRole_(userEmail, meta, opt);
    var inputRoundCutoff = vNextLatestInputRoundCutoff_(metas);
    var evidenceRows = vNextReadRecords_('EVIDENCE_EVENT', { spreadsheet: store }).filter(function (row) {
      return String(row.book_id || '') === bookId &&
        ['ACTIVE', 'SUBMITTED'].indexOf(String(row.status || 'ACTIVE').toUpperCase()) >= 0 &&
        ['COMMITMENT', 'HUMAN_CHANGE', 'CHECK_IN'].indexOf(String(row.evidence_type || '').toUpperCase()) >= 0 &&
        vNextEvidenceInInputRound_(row, inputRoundCutoff);
    });
    var latestByActor = {};
    evidenceRows.forEach(function (row) {
      latestByActor[String(row.actor_email || '').toLowerCase()] = row;
    });
    var team = vNextParseJsonArray_(meta.team_member_emails_json).map(function (email) {
      return String(email || '').trim().toLowerCase();
    }).filter(Boolean);
    var owners = String(meta.forecast_owner_email || meta.forecast_owner_emails || '').split(',').map(function (email) {
      return email.trim().toLowerCase();
    }).filter(Boolean);
    owners.forEach(function (ownerEmail) { if (team.indexOf(ownerEmail) < 0) team.push(ownerEmail); });
    var answeredCount = Object.keys(latestByActor).filter(function (email) {
      return team.length ? team.indexOf(email) >= 0 : true;
    }).length;
    var latestResponses = Object.keys(latestByActor).filter(function (email) {
      return team.length ? team.indexOf(email) >= 0 : true;
    }).map(function (email) { return String(latestByActor[email].response_type || '').toUpperCase(); });
    var unknownCount = latestResponses.filter(function (type) { return type === 'UNKNOWN'; }).length;
    var noChangeCount = latestResponses.filter(function (type) { return type === 'NO_CHANGE'; }).length;
    var changeCount = latestResponses.filter(function (type) { return type === 'CHANGE'; }).length;
    var submitted = !!latestByActor[userEmail];
    return {
      mode: 'CLIENT_BOOK',
      bookId: bookId,
      clientId: String(meta.client_id || ''),
      clientName: String(meta.client_name || ''),
      fiscalYear: Number(meta.fiscal_year),
      asOf: vNextCellDateString_(meta.as_of),
      cutoff: vNextCellDateString_(meta.cutoff),
      state: state,
      role: role,
      isForecastOwner: role === 'FORECAST_OWNER' || role === 'ADMIN',
      isTeamMember: team.indexOf(userEmail) >= 0 || role === 'ADMIN',
      forecastOwnerEmails: owners,
      userEmail: userEmail,
      inputStatus: {
        submitted: submitted,
        answeredCount: answeredCount,
        totalCount: team.length,
        unknownCount: unknownCount,
        noChangeCount: noChangeCount,
        changeCount: changeCount,
        informationGapRate: team.length ? unknownCount / team.length : 0,
        dueDate: vNextCellDateString_(meta.input_due_date),
        roundStartedAt: inputRoundCutoff ? inputRoundCutoff.toISOString() : ''
      },
      canProceed: answeredCount >= team.length || role === 'FORECAST_OWNER' || role === 'ADMIN',
      latestOwnEvidence: latestByActor[userEmail]
        ? vNextEvidenceForClientView_(latestByActor[userEmail], Number(meta.fiscal_year))
        : null,
      version: {
        core: VNEXT_CORE.VERSION,
        schema: String(meta.schema_version || VNEXT_CORE.SCHEMA_VERSION),
        template: String(meta.template_version || VNEXT_CORE.DEFAULT_TEMPLATE_VERSION),
        modelReleaseId: String(meta.model_release_id || '')
      }
    };
  } catch (error) {
    vNextLog_('vNextGetBookContext_ failed', error);
    throw error;
  }
}

/** Latest INPUT_REOPENED BOOK_META row starts a fresh answer round. */
function vNextLatestInputRoundCutoff_(metas) {
  var cutoff = null;
  (metas || []).forEach(function (row) {
    if (String(row.event_type || '').toUpperCase() !== 'INPUT_REOPENED') return;
    var value = row.recorded_at;
    var parsed = value instanceof Date ? new Date(value.getTime()) : new Date(String(value || ''));
    if (isNaN(parsed.getTime())) throw new Error('INPUT_REOPENED recorded_at is invalid.');
    if (!cutoff || parsed.getTime() >= cutoff.getTime()) cutoff = parsed;
  });
  return cutoff;
}

function vNextEvidenceInInputRound_(row, cutoff) {
  if (!cutoff) return true;
  var value = row && row.created_at;
  var parsed = value instanceof Date ? new Date(value.getTime()) : new Date(String(value || ''));
  return !isNaN(parsed.getTime()) && parsed.getTime() >= cutoff.getTime();
}

function vNextAppendEvidence_(payload, options) {
  try {
    var opt = options || {};
    var actualUser = vNextActiveUserEmail_();
    var localStore = opt.spreadsheet || vNextResolveActiveLocalAuditStore_(payload && payload.bookId);
    var context = opt.context || vNextGetBookContext_({
      bookId: payload && payload.bookId,
      spreadsheet: localStore,
      userEmail: actualUser
    });
    if (!context.isTeamMember && !context.isForecastOwner) throw new Error('Only registered team members can append evidence.');
    if (['INPUT_OPEN', 'READY_TO_RUN', 'CHANGES_REQUESTED'].indexOf(String(context.state || '').toUpperCase()) < 0) {
      throw new Error('Evidence input is closed in state ' + context.state + '.');
    }
    var normalized = vNextNormalizeEvidencePayload_(payload || {}, context, opt);
    vNextAppendRecord_('EVIDENCE_EVENT', normalized, { spreadsheet: localStore });
    return {
      evidenceId: normalized.evidence_id,
      bookId: normalized.book_id,
      responseType: normalized.response_type,
      createdAt: normalized.created_at
    };
  } catch (error) {
    vNextLog_('vNextAppendEvidence_ failed', error);
    throw error;
  }
}

function vNextNormalizeEvidencePayload_(payload, context, options) {
  var opt = options || {};
  var responseType = vNextNormalizeResponseType_(payload.responseType || payload.response_type);
  var evidenceType = String(payload.evidenceType || payload.evidence_type || (responseType === 'CHANGE' ? 'HUMAN_CHANGE' : 'CHECK_IN')).toUpperCase();
  var confidence = vNextNormalizeConfidence_(payload.confidence || payload.confidenceClass || payload.confidence_class);
  var direction = vNextNormalizeDirection_(payload.direction);
  var normalizedPeriod = vNextEvidencePeriod_(payload.period || {}, context.fiscalYear);
  var periodStart = normalizedPeriod.start;
  var periodEnd = normalizedPeriod.end;
  var amount = payload.amount;
  var amountLow = payload.amountLow;
  var amountHigh = payload.amountHigh;
  if (amount && typeof amount === 'object') {
    amountLow = amount.low;
    amountHigh = amount.high;
    amount = amount.mid !== undefined ? amount.mid : amount.value;
  }
  if (responseType === 'CHANGE') {
    if (!String(payload.target || '').trim()) throw new Error('Change evidence requires target.');
    if (!direction) throw new Error('Change evidence requires direction.');
    if (!String(payload.evidence || payload.evidenceText || '').trim()) throw new Error('Change evidence requires evidence text.');
    if (!confidence) throw new Error('Change evidence requires confidence class.');
    if (!vNextIsFiniteNumber_(amount) && !String(payload.amountBand || payload.amount_band || '').trim()) {
      throw new Error('Change evidence requires amount or amountBand.');
    }
  }
  if (evidenceType.indexOf('AI') >= 0) {
    var requiredAiFields = {
      sourceUrl: payload.sourceUrl || payload.source_url,
      sourceDate: payload.sourceDate || payload.source_date,
      expiresAt: payload.expiresAt || payload.expires_at,
      evidenceQuality: payload.evidenceQuality || payload.evidence_quality,
      aiModel: payload.aiModel || payload.ai_model,
      promptVersion: payload.promptVersion || payload.prompt_version,
      aiSchemaVersion: payload.aiSchemaVersion || payload.ai_schema_version,
      ruleVersion: payload.ruleVersion || payload.rule_version
    };
    Object.keys(requiredAiFields).forEach(function (field) {
      if (!String(requiredAiFields[field] || '').trim()) throw new Error('AI evidence requires ' + field + '.');
    });
  }
  var actor = String(vNextActiveUserEmail_() || context.userEmail || '').trim().toLowerCase();
  if (!actor) throw new Error('Signed-in user email is required to append evidence.');
  var status = String(payload.status || 'ACTIVE').toUpperCase();
  if (['ACTIVE', 'SUBMITTED'].indexOf(status) < 0) throw new Error('Invalid employee evidence status: ' + status);
  return {
    evidence_id: vNextUuid_(),
    book_id: context.bookId,
    client_id: context.clientId || '',
    fiscal_year: context.fiscalYear,
    actor_email: actor,
    response_type: responseType,
    evidence_type: evidenceType,
    target: String(payload.target || '').trim(),
    target_start_month: periodStart ? vNextFormatMonth_(periodStart) : '',
    target_end_month: periodEnd ? vNextFormatMonth_(periodEnd) : '',
    direction: direction,
    amount_mode: String(payload.amountMode || payload.amount_mode || (vNextIsFiniteNumber_(amount) ? 'EXACT' : 'BAND')).toUpperCase(),
    amount_low: vNextNumberOrBlank_(amountLow),
    amount_mid: vNextNumberOrBlank_(amount),
    amount_high: vNextNumberOrBlank_(amountHigh),
    amount_band: String(payload.amountBand || payload.amount_band || '').trim().toUpperCase(),
    confidence_class: confidence,
    evidence_text: String(payload.evidence || payload.evidenceText || '').trim(),
    source_url: String(payload.sourceUrl || payload.source_url || '').trim(),
    source_date: payload.sourceDate || payload.source_date ? vNextFormatDateOnly_(payload.sourceDate || payload.source_date) : '',
    expires_at: payload.expiresAt || payload.expires_at ? vNextFormatDateOnly_(payload.expiresAt || payload.expires_at) : '',
    status: status,
    supersedes_evidence_id: String(payload.supersedesEvidenceId || payload.supersedesEventId || payload.supersedes_evidence_id || ''),
    created_at: vNextNowIso_(),
    evidence_quality: String(payload.evidenceQuality || payload.evidence_quality || ''),
    ai_model: String(payload.aiModel || payload.ai_model || ''),
    prompt_version: String(payload.promptVersion || payload.prompt_version || ''),
    ai_schema_version: String(payload.aiSchemaVersion || payload.ai_schema_version || ''),
    rule_version: String(payload.ruleVersion || payload.rule_version || ''),
    applied_amount: vNextNumberOrBlank_(payload.appliedAmount || payload.applied_amount),
    cap_applied: vNextBoolean_(payload.capApplied !== undefined ? payload.capApplied : payload.cap_applied) ? 1 : 0
  };
}

function vNextTransitionState_(bookIdOrRequest, toState, reason, options) {
  try {
    var request;
    if (bookIdOrRequest && typeof bookIdOrRequest === 'object') {
      request = bookIdOrRequest;
    } else {
      request = options || {};
      request.bookId = bookIdOrRequest || request.bookId;
      request.toState = toState || request.toState;
      request.reason = reason || request.reason;
    }
    var opt = request.options || request;
    var localStore = opt.spreadsheet || vNextResolveActiveLocalAuditStore_(request.bookId);
    var context = request.context || vNextGetBookContext_({
      bookId: request.bookId,
      spreadsheet: localStore,
      userEmail: vNextActiveUserEmail_() || request.actorEmail || opt.userEmail
    });
    var currentState = String(context.state || 'INPUT_OPEN').toUpperCase();
    if (request.fromState && String(request.fromState).toUpperCase() !== currentState) {
      throw new Error('State changed before this action. Expected ' + request.fromState + ' but found ' + currentState + '.');
    }
    var fromState = currentState;
    var targetState = String(request.toState || '').toUpperCase();
    var role = String(context.role || 'MEMBER').toUpperCase();
    var readinessSystemTransition = String(request.actorRole || '').toUpperCase() === 'SYSTEM' &&
      fromState === 'INPUT_OPEN' && targetState === 'READY_TO_RUN' &&
      Number(context.inputStatus && context.inputStatus.totalCount || 0) > 0 &&
      Number(context.inputStatus && context.inputStatus.answeredCount || 0) >= Number(context.inputStatus && context.inputStatus.totalCount || 0);
    var trustedEngineTransition = request.internalOperation === 'FORECAST_ENGINE';
    if (readinessSystemTransition || trustedEngineTransition) role = 'SYSTEM';
    if (fromState === targetState) {
      return { changed: false, bookId: context.bookId, fromState: fromState, toState: targetState };
    }
    vNextValidateTransition_(fromState, targetState, role);
    var event = {
      state_event_id: vNextUuid_(),
      book_id: context.bookId,
      from_state: fromState,
      to_state: targetState,
      reason: String(request.reason || '').trim(),
      actor_email: String(context.userEmail || vNextActiveUserEmail_() || '').toLowerCase(),
      actor_role: role,
      related_run_id: String(request.relatedRunId || ''),
      related_plan_version_id: String(request.relatedPlanVersionId || ''),
      created_at: vNextNowIso_()
    };
    vNextAppendRecord_('STATE_EVENT', event, { spreadsheet: localStore });
    try {
      if (typeof SpreadsheetApp !== 'undefined') {
        var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        var localConfig = vNextReadLocalRoutingConfig_(activeSpreadsheet);
        if (String(localConfig.book_id || '') === String(context.bookId || '')) {
          vNextWriteLocalRoutingValue_(activeSpreadsheet, 'VN_BOOK_CONFIG', 'state', targetState);
          vNextWriteLocalRoutingValue_(activeSpreadsheet, 'BOOK_META', 'state', targetState);
        }
      }
    } catch (localStateError) {
      vNextLog_('Local client state mirror was not updated', localStateError);
    }
    return {
      changed: true,
      stateEventId: event.state_event_id,
      bookId: event.book_id,
      fromState: fromState,
      toState: targetState
    };
  } catch (error) {
    vNextLog_('vNextTransitionState_ failed', error);
    throw error;
  }
}

function vNextValidateTransition_(fromState, toState, actorRole) {
  var from = String(fromState || '').toUpperCase();
  var to = String(toState || '').toUpperCase();
  var role = String(actorRole || 'MEMBER').toUpperCase();
  if (VNEXT_CORE.STATES.indexOf(from) < 0 || VNEXT_CORE.STATES.indexOf(to) < 0) {
    throw new Error('Unknown state transition: ' + from + ' -> ' + to);
  }
  if (VNEXT_CORE.STATE_TRANSITIONS[from].indexOf(to) < 0) {
    throw new Error('State transition is not allowed: ' + from + ' -> ' + to);
  }
  var requiredRoles = {
    'INPUT_OPEN>READY_TO_RUN': ['FORECAST_OWNER', 'ADMIN', 'SYSTEM'],
    'READY_TO_RUN>RUNNING': ['FORECAST_OWNER', 'ADMIN', 'SYSTEM'],
    'READY_TO_RUN>INPUT_OPEN': ['FORECAST_OWNER', 'ADMIN', 'SYSTEM'],
    'RUNNING>DRAFT_READY': ['ADMIN', 'SYSTEM'],
    'RUNNING>READY_TO_RUN': ['ADMIN', 'SYSTEM'],
    'DRAFT_READY>SUBMITTED': ['FORECAST_OWNER', 'ADMIN'],
    'DRAFT_READY>READY_TO_RUN': ['FORECAST_OWNER', 'ADMIN', 'SYSTEM'],
    'SUBMITTED>OFFICIAL_LOCKED': ['ADMIN'],
    'SUBMITTED>CHANGES_REQUESTED': ['ADMIN'],
    'CHANGES_REQUESTED>INPUT_OPEN': ['FORECAST_OWNER', 'ADMIN'],
    'CHANGES_REQUESTED>READY_TO_RUN': ['FORECAST_OWNER', 'ADMIN'],
    'CHANGES_REQUESTED>SUBMITTED': ['FORECAST_OWNER', 'ADMIN'],
    'OFFICIAL_LOCKED>REVIEW_DUE': ['ADMIN', 'SYSTEM'],
    'REVIEW_DUE>YEAR_CLOSED': ['ADMIN']
  };
  var allowedRoles = requiredRoles[from + '>' + to] || ['ADMIN'];
  if (allowedRoles.indexOf(role) < 0) {
    throw new Error('Role ' + role + ' cannot transition ' + from + ' -> ' + to);
  }
  return true;
}

function vNextCreateBookMeta_(payload, options) {
  try {
    var data = payload || {};
    if (!String(data.bookId || data.book_id || '').trim()) throw new Error('bookId is required.');
    if (!String(data.clientName || data.client_name || '').trim()) throw new Error('clientName is required.');
    var fiscalYear = Number(data.fiscalYear || data.fiscal_year);
    if (!isFinite(fiscalYear)) throw new Error('fiscalYear is required.');
    var asOf = vNextParseDate_(data.asOf || data.as_of || new Date(), 'as_of');
    var record = {
      record_id: vNextUuid_(),
      book_id: String(data.bookId || data.book_id),
      client_id: String(data.clientId || data.client_id || ''),
      client_name: String(data.clientName || data.client_name),
      fiscal_year: fiscalYear,
      forecast_owner_email: String(data.forecastOwnerEmail || data.forecast_owner_email || '').toLowerCase(),
      team_member_emails_json: vNextCanonicalJson_(data.teamMemberEmails || data.team_member_emails || []),
      state: String(data.state || 'INPUT_OPEN').toUpperCase(),
      as_of: vNextFormatDateOnly_(asOf),
      cutoff: vNextFormatDateOnly_(vNextCutoffFromAsOf_(asOf)),
      template_version: String(data.templateVersion || VNEXT_CORE.DEFAULT_TEMPLATE_VERSION),
      schema_version: String(data.schemaVersion || VNEXT_CORE.SCHEMA_VERSION),
      model_release_id: String(data.modelReleaseId || ''),
      source_spreadsheet_id: String(data.sourceSpreadsheetId || ''),
      client_book_id: String(data.clientBookId || ''),
      input_due_date: data.inputDueDate ? vNextFormatDateOnly_(data.inputDueDate) : '',
      event_type: String(data.eventType || 'CREATED').toUpperCase(),
      supersedes_record_id: String(data.supersedesRecordId || ''),
      recorded_at: vNextNowIso_(),
      recorded_by: String(data.recordedBy || vNextActiveUserEmail_()).toLowerCase()
    };
    vNextAppendRecord_('BOOK_META', record, options || {});
    return record;
  } catch (error) {
    vNextLog_('vNextCreateBookMeta_ failed', error);
    throw error;
  }
}

function vNextCreateModelRelease_(payload, options) {
  try {
    var data = payload || {};
    var record = {
      model_release_id: String(data.modelReleaseId || vNextUuid_()),
      status: String(data.status || 'DRAFT').toUpperCase(),
      model_version: String(data.modelVersion || 'vnext-engine-1'),
      schema_version: String(data.schemaVersion || VNEXT_CORE.SCHEMA_VERSION),
      template_version: String(data.templateVersion || VNEXT_CORE.DEFAULT_TEMPLATE_VERSION),
      parameters_json: vNextCanonicalJson_(data.parameters || {}),
      backtest_json: vNextCanonicalJson_(data.backtest || {}),
      canary_json: vNextCanonicalJson_(data.canary || {}),
      approved_at: data.approvedAt || '',
      approved_by: String(data.approvedBy || '').toLowerCase(),
      rollback_release_id: String(data.rollbackReleaseId || ''),
      created_at: vNextNowIso_(),
      created_by: String(data.createdBy || vNextActiveUserEmail_()).toLowerCase(),
      note: String(data.note || '')
    };
    vNextAppendRecord_('MODEL_RELEASE', record, options || {});
    return record;
  } catch (error) {
    vNextLog_('vNextCreateModelRelease_ failed', error);
    throw error;
  }
}

function vNextResolveRole_(email, meta, options) {
  var opt = options || {};
  if (opt.trustedActorRole) return String(opt.trustedActorRole).toUpperCase();
  var actor = String(email || '').trim().toLowerCase();
  var adminEmails = String(vNextGetProperty_(VNEXT_CORE.ADMIN_EMAILS_PROPERTY, false) || '')
    .split(',').map(function (item) { return item.trim().toLowerCase(); }).filter(Boolean);
  if (actor && adminEmails.indexOf(actor) >= 0) return 'ADMIN';
  var owners = String(meta.forecast_owner_email || meta.forecast_owner_emails || '')
    .split(',').map(function (item) { return item.trim().toLowerCase(); }).filter(Boolean);
  if (actor && owners.indexOf(actor) >= 0) return 'FORECAST_OWNER';
  return 'MEMBER';
}

function vNextResolveStoreSpreadsheet_(options) {
  var opt = options || {};
  if (opt.spreadsheet) return opt.spreadsheet;
  if (typeof SpreadsheetApp === 'undefined') throw new Error('SpreadsheetApp is unavailable; pass options.spreadsheet.');
  var active = null;
  try { active = opt.clientSpreadsheet || SpreadsheetApp.getActiveSpreadsheet(); } catch (ignore) {}
  var routing = active ? vNextReadLocalRoutingConfig_(active) : {};
  var hubId = opt.hubSpreadsheetId || routing.admin_hub_spreadsheet_id || routing.adminHubSpreadsheetId ||
    vNextGetProperty_(VNEXT_CORE.HUB_PROPERTY, true) || '';
  if (hubId) {
    try {
      return SpreadsheetApp.openById(String(hubId));
    } catch (accessError) {
      throw new Error(
        '管理ハブへアクセスできません。クライアント年度ブック利用者へ管理ハブを共有して解決しないでください。' +
        '管理ハブ担当者権限で実行する Forecast Service / Web App 経路が必要です。hub=' + String(hubId) +
        '; cause=' + String(accessError && accessError.message || accessError)
      );
    }
  }
  if (opt.allowActiveSpreadsheet === false) {
    throw new Error(VNEXT_CORE.HUB_PROPERTY + ' is not configured.');
  }
  return active || SpreadsheetApp.getActiveSpreadsheet();
}

function vNextReadLocalRoutingConfig_(spreadsheet) {
  var output = {};
  if (!spreadsheet || typeof spreadsheet.getSheetByName !== 'function') return output;
  ['VN_SYSTEM_CONFIG', 'VN_BOOK_CONFIG', 'BOOK_META'].forEach(function (sheetName) {
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 1) return;
    var values = sheet.getDataRange().getValues();
    if (!values.length) return;
    var header = values[0].map(function (value) { return String(value || '').trim().toLowerCase(); });
    var keyIndex = header.indexOf('key');
    var valueIndex = header.indexOf('value');
    if (keyIndex >= 0 && valueIndex >= 0) {
      values.slice(1).forEach(function (row) {
        var key = String(row[keyIndex] || '').trim();
        if (key) output[key] = row[valueIndex];
      });
      return;
    }
    if (values[0].length >= 2) {
      values.forEach(function (row) {
        var key = String(row[0] || '').trim();
        if (key && key.toLowerCase() !== 'key') output[key] = row[1];
      });
    }
  });
  return output;
}

function vNextBuildLocalClientContext_(spreadsheet, routing, options, storeError) {
  var opt = options || {};
  var userEmail = String(opt.userEmail || vNextActiveUserEmail_()).toLowerCase();
  var meta = {
    forecast_owner_emails: routing.forecast_owner_emails || '',
    forecast_owner_email: routing.forecast_owner_email || ''
  };
  var role = vNextResolveRole_(userEmail, meta, {});
  var owners = String(meta.forecast_owner_emails || meta.forecast_owner_email || '').split(',').map(function (email) {
    return email.trim().toLowerCase();
  }).filter(Boolean);
  var answered = Number(routing.input_answered_count || 0);
  var total = Number(routing.input_total_count || 0);
  return {
    mode: 'CLIENT_BOOK',
    bookId: String(routing.book_id || vNextSpreadsheetId_(spreadsheet)),
    clientId: String(routing.client_id || ''),
    clientName: String(routing.client_name || ''),
    fiscalYear: Number(routing.fiscal_year),
    asOf: vNextCellDateString_(routing.as_of),
    cutoff: vNextCellDateString_(routing.cutoff),
    state: String(routing.state || 'INPUT_OPEN').toUpperCase(),
    role: role,
    isForecastOwner: role === 'FORECAST_OWNER' || role === 'ADMIN',
    isTeamMember: role === 'FORECAST_OWNER' || role === 'ADMIN',
    forecastOwnerEmails: owners,
    userEmail: userEmail,
      inputStatus: {
        submitted: vNextBoolean_(routing.input_submitted),
        answeredCount: answered,
        totalCount: total,
        unknownCount: Number(routing.input_unknown_count || 0),
        noChangeCount: Number(routing.input_no_change_count || 0),
        changeCount: Number(routing.input_change_count || 0),
        informationGapRate: total > 0 ? Number(routing.input_unknown_count || 0) / total : 0,
        dueDate: vNextCellDateString_(routing.input_due_date)
    },
    canProceed: total > 0 && answered >= total,
    latestOwnEvidence: null,
    serviceUnavailable: true,
    serviceError: String(storeError && storeError.message || storeError || ''),
    version: {
      core: VNEXT_CORE.VERSION,
      schema: String(routing.schema_version || VNEXT_CORE.SCHEMA_VERSION),
      template: String(routing.template_release_id || routing.version || ''),
      modelReleaseId: String(routing.active_release_id || routing.version || '')
    }
  };
}

function vNextResolveActiveLocalAuditStore_(bookId) {
  if (typeof SpreadsheetApp === 'undefined') return undefined;
  try {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (!vNextIsAuditStoreSpreadsheet_(active)) return undefined;
    var routing = vNextReadLocalRoutingConfig_(active);
    var localBookId = String(routing.book_id || '');
    if (bookId && localBookId && String(bookId) !== localBookId) return undefined;
    return active;
  } catch (error) {
    vNextLog_('Active local audit store resolution skipped', error);
    return undefined;
  }
}

function vNextFindRegistryBookMeta_(spreadsheet, bookId) {
  if (!spreadsheet || !bookId) return null;
  var sheet = spreadsheet.getSheetByName('BOOK_REGISTRY');
  if (!sheet || sheet.getLastRow() < 2) return null;
  var values = sheet.getDataRange().getValues();
  var headers = values[0].map(function (value) { return String(value || '').trim(); });
  var index = {};
  headers.forEach(function (header, column) { if (header) index[header] = column; });
  if (index.book_id === undefined) return null;
  for (var i = values.length - 1; i >= 1; i--) {
    if (String(values[i][index.book_id] || '') !== String(bookId)) continue;
    var row = {};
    headers.forEach(function (header, column) { if (header) row[header] = values[i][column]; });
    row.forecast_owner_email = row.forecast_owner_emails || '';
    row.team_member_emails_json = vNextCanonicalJson_(String(row.editor_emails || '').split(',').map(function (email) { return email.trim(); }).filter(Boolean));
    row.as_of = row.as_of || '';
    row.cutoff = row.cutoff || '';
    row.template_version = row.template_release_id || '';
    row.model_release_id = row.template_release_id || '';
    row.input_due_date = row.input_due_date || '';
    return row;
  }
  return null;
}

function vNextWriteLocalRoutingValue_(spreadsheet, sheetName, key, value) {
  if (!spreadsheet || !sheetName || !key) return false;
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 1) return false;
  var values = sheet.getDataRange().getValues();
  var headers = values[0].map(function (item) { return String(item || '').trim().toLowerCase(); });
  var keyIndex = headers.indexOf('key');
  var valueIndex = headers.indexOf('value');
  if (keyIndex < 0 || valueIndex < 0) return false;
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][keyIndex] || '').trim() !== key) continue;
    sheet.getRange(i + 1, valueIndex + 1).setValue(value);
    return true;
  }
  return false;
}

function vNextBoolean_(value) {
  return value === true || value === 1 || String(value || '').toLowerCase() === 'true' || String(value) === '1';
}

function vNextIsAuditStoreSpreadsheet_(spreadsheet) {
  if (!spreadsheet || typeof spreadsheet.getSheetByName !== 'function') return false;
  var bookMeta = spreadsheet.getSheetByName('BOOK_META');
  if (!bookMeta) return false;
  try {
    return String(bookMeta.getRange(1, 1).getValue() || '').trim() === 'record_id';
  } catch (error) {
    return false;
  }
}

function vNextEvidenceForClientView_(row, fiscalYear) {
  var direction = String(row.direction || '').toUpperCase();
  var confidence = String(row.confidence_class || '').toUpperCase();
  return {
    evidenceId: String(row.evidence_id || ''),
    responseType: String(row.response_type || '').toLowerCase(),
    evidenceType: String(row.evidence_type || '').toUpperCase(),
    changeKind: /COMMIT|CONTRACT/.test(String(row.evidence_type || '').toUpperCase()) ? 'contract' : 'other',
    target: String(row.target || ''),
    period: vNextEvidencePeriodLabel_(row.target_start_month, row.target_end_month, fiscalYear),
    direction: direction === 'UP' ? 'increase' : (direction === 'DOWN' ? 'decrease' : ''),
    amountMode: String(row.amount_mode || '').toLowerCase(),
    amount: vNextNumberOrBlank_(row.amount_mid),
    amountBand: String(row.amount_band || '').toLowerCase(),
    evidence: String(row.evidence_text || ''),
    confidence: confidence === 'CONFIRMED_FACT' ? 'confirmed' : (confidence === 'LIKELY' ? 'likely' : (confidence === 'HYPOTHESIS' ? 'hypothesis' : '')),
    createdAt: vNextCellDateString_(row.created_at)
  };
}

function vNextEvidencePeriodLabel_(startValue, endValue, fiscalYear) {
  if (!startValue) return '';
  var start = vNextFormatMonth_(String(startValue).length === 7 ? startValue + '-01' : startValue);
  var end = vNextFormatMonth_(String(endValue || startValue).length === 7 ? String(endValue || startValue) + '-01' : (endValue || startValue));
  var fy = Number(fiscalYear);
  var labels = [
    ['FY通年', fy + '-04', (fy + 1) + '-03'],
    ['Q1', fy + '-04', fy + '-06'],
    ['Q2', fy + '-07', fy + '-09'],
    ['Q3', fy + '-10', fy + '-12'],
    ['Q4', (fy + 1) + '-01', (fy + 1) + '-03']
  ];
  for (var i = 0; i < labels.length; i++) if (labels[i][1] === start && labels[i][2] === end) return labels[i][0];
  return start === end ? start : start + '〜' + end;
}

function vNextGetProperty_(key, includeDocument) {
  try {
    if (typeof PropertiesService === 'undefined') return '';
    if (includeDocument && PropertiesService.getDocumentProperties) {
      var documentValue = PropertiesService.getDocumentProperties().getProperty(key);
      if (documentValue) return documentValue;
    }
    return PropertiesService.getScriptProperties().getProperty(key) || '';
  } catch (error) {
    vNextLog_('Property read failed: ' + key, error);
    return '';
  }
}

function vNextWithScriptLock_(operation, timeoutMs) {
  if (typeof LockService === 'undefined' || !LockService.getScriptLock) return operation();
  var lock = LockService.getScriptLock();
  var acquired = false;
  try {
    lock.waitLock(Math.max(1000, Number(timeoutMs) || 30000));
    acquired = true;
    return operation();
  } finally {
    if (acquired) lock.releaseLock();
  }
}

function vNextValueForSheet_(value) {
  if (value === undefined || value === null) return '';
  if (value instanceof Date) return value.toISOString();
  if (typeof value === 'object') return vNextCanonicalJson_(value);
  return value;
}

function vNextNormalizeResponseType_(value) {
  var text = String(value || '').trim().toUpperCase();
  var map = {
    '変化あり': 'CHANGE', CHANGE: 'CHANGE',
    '確認したが変化なし': 'NO_CHANGE', '変化なし': 'NO_CHANGE', NO_CHANGE: 'NO_CHANGE',
    'わからない・情報不足': 'UNKNOWN', 'わからない': 'UNKNOWN', UNKNOWN: 'UNKNOWN'
  };
  var normalized = map[text] || map[String(value || '').trim()] || '';
  if (VNEXT_CORE.RESPONSE_TYPES.indexOf(normalized) < 0) throw new Error('Invalid responseType: ' + value);
  return normalized;
}

function vNextNormalizeConfidence_(value) {
  var text = String(value || '').trim().toUpperCase();
  var map = {
    '確認済み事実': 'CONFIRMED_FACT', CONFIRMED: 'CONFIRMED_FACT', CONFIRMED_FACT: 'CONFIRMED_FACT',
    '有力情報': 'LIKELY', LIKELY: 'LIKELY',
    '仮説': 'HYPOTHESIS', HYPOTHESIS: 'HYPOTHESIS'
  };
  return map[text] || map[String(value || '').trim()] || '';
}

function vNextNormalizeDirection_(value) {
  var text = String(value || '').trim().toUpperCase();
  var map = { UP: 'UP', INCREASE: 'UP', DOWN: 'DOWN', DECREASE: 'DOWN', NEUTRAL: 'NEUTRAL', '増加': 'UP', '減少': 'DOWN', '中立': 'NEUTRAL' };
  return map[text] || map[String(value || '').trim()] || '';
}

function vNextEvidencePeriod_(value, fiscalYear) {
  if (value && typeof value === 'object') {
    return { start: value.start || value.from || '', end: value.end || value.to || value.start || value.from || '' };
  }
  var text = String(value || '').trim();
  if (!text) return { start: '', end: '' };
  var fy = Number(fiscalYear);
  var labels = {
    'FY通年': [new Date(fy, 3, 1), new Date(fy + 1, 2, 1)],
    'Q1': [new Date(fy, 3, 1), new Date(fy, 5, 1)],
    'Q2': [new Date(fy, 6, 1), new Date(fy, 8, 1)],
    'Q3': [new Date(fy, 9, 1), new Date(fy, 11, 1)],
    'Q4': [new Date(fy + 1, 0, 1), new Date(fy + 1, 2, 1)],
    '4～6月': [new Date(fy, 3, 1), new Date(fy, 5, 1)],
    '7～9月': [new Date(fy, 6, 1), new Date(fy, 8, 1)],
    '10～12月': [new Date(fy, 9, 1), new Date(fy, 11, 1)],
    '1～3月': [new Date(fy + 1, 0, 1), new Date(fy + 1, 2, 1)]
  };
  if (labels[text]) return { start: labels[text][0], end: labels[text][1] };
  return { start: text, end: text };
}

function vNextNumberOrBlank_(value) {
  return vNextIsFiniteNumber_(value) ? Number(value) : '';
}

function vNextIsFiniteNumber_(value) {
  return value !== '' && value !== null && value !== undefined && isFinite(Number(value));
}

function vNextParseJsonArray_(value) {
  if (Array.isArray(value)) return value.slice();
  if (!value) return [];
  try {
    var parsed = JSON.parse(String(value));
    return Array.isArray(parsed) ? parsed : [];
  } catch (error) {
    return String(value).split(',').map(function (item) { return item.trim(); }).filter(Boolean);
  }
}

function vNextCellDateString_(value) {
  if (!value) return '';
  try { return vNextFormatDateOnly_(value); } catch (error) { return String(value); }
}

function vNextSpreadsheetId_(spreadsheet) {
  try { return spreadsheet && spreadsheet.getId ? spreadsheet.getId() : ''; } catch (error) { return ''; }
}

function vNextLog_(message, error) {
  var detail = error ? ': ' + String(error && error.stack ? error.stack : error) : '';
  if (typeof Logger !== 'undefined' && Logger.log) Logger.log('[vNext] ' + message + detail);
  else if (typeof console !== 'undefined' && console.log) console.log('[vNext] ' + message + detail);
}
