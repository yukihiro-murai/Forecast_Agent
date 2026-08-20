/**
 * Forecast vNext Client FY Book local core.
 * No Admin Hub lookup, source-system access, AI call, prompt, or secret is present here.
 */

var VNEXT_CLIENT_CORE = Object.freeze({
  RUNTIME_VERSION: 'vnext-client-1.8.1',
  SCHEMA_VERSION: 'vnext-schema-2',
  CONFIG_SHEET: 'VN_BOOK_CONFIG',
  REQUEST_SHEET: 'VN_CLIENT_REQUEST',
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
  REQUEST_HEADERS: Object.freeze([
    'request_event_id', 'request_id', 'book_id', 'event_type', 'status',
    'request_hash', 'request_json', 'requested_at', 'requested_by',
    'related_job_id', 'related_run_id', 'detail_json', 'created_at'
  ])
});

/** Initializes only missing local audit sheets. Existing append-only rows are untouched. */
function vNextEnsureAuditStore_(spreadsheet) {
  try {
    var ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    Object.keys(VNEXT_CLIENT_CORE.INTERNAL_SHEETS).forEach(function (sheetName) {
      var sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
      vNextEnsureAppendOnlyHeader_(sheet, VNEXT_CLIENT_CORE.INTERNAL_SHEETS[sheetName]);
      sheet.setFrozenRows(1);
      try { sheet.hideSheet(); } catch (hideError) { Logger.log('[vNext Client] hide skipped: ' + sheetName); }
    });
    return { ok: true, spreadsheetId: ss.getId() };
  } catch (error) {
    vNextClientLog_('vNextEnsureAuditStore_ failed', error);
    throw error;
  }
}

function vNextEnsureAppendOnlyHeader_(sheet, headers) {
  var current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  var blank = current.every(function (value) { return String(value || '').trim() === ''; });
  if (blank) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers.slice()]);
    return;
  }
  headers.forEach(function (header, index) {
    if (String(current[index] || '') !== header) {
      throw new Error(sheet.getName() + ' の列構成が正本と一致しません。列 ' + (index + 1));
    }
  });
}

function vNextAppendRecord_(sheetName, record, options) {
  return vNextAppendRecords_(sheetName, [record], options || {});
}

function vNextAppendRecords_(sheetName, records, options) {
  try {
    var opt = options || {};
    if (!records || !records.length) return { appended: 0, firstRow: 0 };
    var headers = VNEXT_CLIENT_CORE.INTERNAL_SHEETS[sheetName];
    if (!headers) throw new Error('未定義の監査テーブルです: ' + sheetName);
    var operation = function () {
      var ss = opt.spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
      vNextEnsureAuditStore_(ss);
      var sheet = ss.getSheetByName(sheetName);
      var rows = records.map(function (item) {
        return headers.map(function (header) { return vNextClientCellValue_(item && item[header]); });
      });
      var firstRow = sheet.getLastRow() + 1;
      // Google Sheets otherwise coerces values such as "2027-04" into Date
      // objects. These two columns are audit identifiers, not dates; keeping
      // them as text preserves the canonical YYYY-MM contract end to end.
      if (sheetName === 'EVIDENCE_EVENT') {
        var startMonthColumn = headers.indexOf('target_start_month') + 1;
        var endMonthColumn = headers.indexOf('target_end_month') + 1;
        if (startMonthColumn > 0 && endMonthColumn === startMonthColumn + 1) {
          sheet.getRange(firstRow, startMonthColumn, rows.length, 2).setNumberFormat('@');
        }
      }
      sheet.getRange(firstRow, 1, rows.length, headers.length).setValues(rows);
      return { appended: rows.length, firstRow: firstRow };
    };
    return opt.skipLock ? operation() : vNextClientWithDocumentLock_(operation);
  } catch (error) {
    vNextClientLog_('vNextAppendRecords_ failed for ' + sheetName, error);
    throw error;
  }
}

function vNextReadRecords_(sheetName, options) {
  try {
    var headers = VNEXT_CLIENT_CORE.INTERNAL_SHEETS[sheetName];
    if (!headers) throw new Error('未定義の監査テーブルです: ' + sheetName);
    var ss = options && options.spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return [];
    vNextEnsureAppendOnlyHeader_(sheet, headers);
    return sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues().map(function (values) {
      var record = {};
      headers.forEach(function (header, index) { record[header] = values[index]; });
      return record;
    });
  } catch (error) {
    vNextClientLog_('vNextReadRecords_ failed for ' + sheetName, error);
    throw error;
  }
}

function vNextGetBookContext_(options) {
  try {
    var ss = options && typeof options.getSheetByName === 'function'
      ? options
      : options && options.spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    var config = vNextClientReadConfig_(ss);
    if (String(config.mode || '').toUpperCase() !== 'CLIENT') throw new Error('このブックはClient FY Bookではありません。');
    var bookId = String(options && options.bookId || config.book_id || ss.getId());
    var metas = vNextReadRecords_('BOOK_META', { spreadsheet: ss }).filter(function (row) {
      return String(row.book_id || '') === bookId;
    });
    if (!metas.length) throw new Error('BOOK_METAが初期化されていません。');
    var meta = metas[metas.length - 1];
    var states = vNextReadRecords_('STATE_EVENT', { spreadsheet: ss }).filter(function (row) {
      return String(row.book_id || '') === bookId;
    });
    // STATE_EVENT is immutable and authoritative. VN_BOOK_CONFIG.state is only a repairable UI cache.
    var latestStateEvent = states.length ? states[states.length - 1] : null;
    var state = String((latestStateEvent && latestStateEvent.to_state) || config.state || meta.state || 'INPUT_OPEN').toUpperCase();
    if (VNEXT_CLIENT_CORE.STATES.indexOf(state) < 0) throw new Error('状態が不正です: ' + state);
    var userEmail = vNextActiveUserEmail_();
    if (!userEmail) throw new Error('ログイン中のメールアドレスを確認できません。会社アカウントで開いてください。');
    var owners = vNextClientEmails_(config.forecast_owner_emails || meta.forecast_owner_email);
    if (owners.length !== 1) throw new Error('予算策定担当は1名である必要があります。');
    var team = vNextClientJsonArray_(meta.team_member_emails_json).map(vNextClientNormalizeEmail_).filter(Boolean);
    owners.forEach(function (owner) { if (team.indexOf(owner) < 0) team.push(owner); });
    var accessPolicy = String(config.access_policy || '').trim().toUpperCase();
    var internalDomain = vNextClientNormalizeDomain_(config.internal_domain);
    var isInternalUser = accessPolicy === 'INTERNAL_OPEN' &&
      Boolean(internalDomain) && vNextClientEmailDomain_(userEmail) === internalDomain;
    var inputRoundCutoff = vNextClientLatestInputRoundCutoff_(metas);
    var evidence = vNextReadRecords_('EVIDENCE_EVENT', { spreadsheet: ss }).filter(function (row) {
      return String(row.book_id || '') === bookId &&
        ['ACTIVE', 'SUBMITTED'].indexOf(String(row.status || 'ACTIVE').toUpperCase()) >= 0 &&
        ['COMMITMENT', 'HUMAN_CHANGE', 'CHECK_IN'].indexOf(String(row.evidence_type || '').toUpperCase()) >= 0 &&
        vNextClientEvidenceInInputRound_(row, inputRoundCutoff);
    });
    var latestByActor = {};
    evidence.forEach(function (row) { latestByActor[vNextClientNormalizeEmail_(row.actor_email)] = row; });
    var answered = team.filter(function (email) { return Boolean(latestByActor[email]); }).length;
    var latestResponses = team.map(function (email) { return latestByActor[email] && String(latestByActor[email].response_type || '').toUpperCase(); }).filter(Boolean);
    var isTeamMember = team.indexOf(userEmail) >= 0;
    var canContribute = isTeamMember || isInternalUser;
    var role = owners.indexOf(userEmail) >= 0
      ? 'FORECAST_OWNER'
      : (isTeamMember ? 'MEMBER' : (isInternalUser ? 'INTERNAL_CONTRIBUTOR' : 'VIEWER'));
    var ownEvidence = latestByActor[userEmail] || null;
    return {
      mode: 'CLIENT_BOOK',
      bookId: bookId,
      clientId: String(meta.client_id || config.client_id || ''),
      clientName: String(meta.client_name || config.client_name || ''),
      fiscalYear: Number(meta.fiscal_year || config.fiscal_year),
      asOf: vNextClientDateOnly_(meta.as_of || config.as_of),
      cutoff: vNextClientDateOnly_(meta.cutoff || config.cutoff),
      state: state,
      stateReason: String(latestStateEvent && latestStateEvent.reason || ''),
      stateChangedAt: String(latestStateEvent && latestStateEvent.created_at || ''),
      relatedRunId: String(latestStateEvent && latestStateEvent.related_run_id || ''),
      role: role,
      isForecastOwner: role === 'FORECAST_OWNER',
      isTeamMember: isTeamMember,
      isInternalUser: isInternalUser,
      canContribute: canContribute,
      forecastOwnerEmails: owners,
      userEmail: userEmail,
      inputStatus: {
        submitted: Boolean(ownEvidence),
        answeredCount: answered,
        totalCount: team.length,
        unknownCount: latestResponses.filter(function (value) { return value === 'UNKNOWN'; }).length,
        noChangeCount: latestResponses.filter(function (value) { return value === 'NO_CHANGE'; }).length,
        changeCount: latestResponses.filter(function (value) { return value === 'CHANGE'; }).length,
        dueDate: vNextClientDateOnly_(meta.input_due_date || config.input_due_date),
        roundStartedAt: inputRoundCutoff ? inputRoundCutoff.toISOString() : ''
      },
      canProceed: team.length > 0 && answered >= team.length,
      latestOwnEvidence: ownEvidence ? vNextClientEvidenceForView_(ownEvidence) : null,
      annualSalesBaseline: Number(config.annual_sales_baseline || 0),
      annualSalesBaselineBasis: String(config.annual_sales_baseline_basis || ''),
      version: {
        runtime: VNEXT_CLIENT_CORE.RUNTIME_VERSION,
        schema: String(meta.schema_version || config.schema_version || VNEXT_CLIENT_CORE.SCHEMA_VERSION),
        template: String(meta.template_version || config.template_release_id || config.version || ''),
        modelReleaseId: String(meta.model_release_id || config.active_release_id || '')
      }
    };
  } catch (error) {
    vNextClientLog_('vNextGetBookContext_ failed', error);
    throw error;
  }
}

/** Latest INPUT_REOPENED BOOK_META row starts a fresh answer round. */
function vNextClientLatestInputRoundCutoff_(metas) {
  var cutoff = null;
  (metas || []).forEach(function (row) {
    if (String(row.event_type || '').toUpperCase() !== 'INPUT_REOPENED') return;
    var value = row.recorded_at;
    var parsed = value instanceof Date ? new Date(value.getTime()) : new Date(String(value || ''));
    if (isNaN(parsed.getTime())) throw new Error('INPUT_REOPENED recorded_atが不正です。');
    if (!cutoff || parsed.getTime() >= cutoff.getTime()) cutoff = parsed;
  });
  return cutoff;
}

function vNextClientEvidenceInInputRound_(row, cutoff) {
  if (!cutoff) return true;
  var value = row && row.created_at;
  var parsed = value instanceof Date ? new Date(value.getTime()) : new Date(String(value || ''));
  return !isNaN(parsed.getTime()) && parsed.getTime() >= cutoff.getTime();
}

function vNextAppendEvidence_(payload, options) {
  try {
    var opt = options || {};
    var ss = opt.spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    var context = vNextGetBookContext_({ spreadsheet: ss, bookId: payload && payload.bookId });
    if (!context.canContribute) throw new Error('登録メンバー、または社内アカウントだけが回答できます。');
    if (['INPUT_OPEN', 'READY_TO_RUN'].indexOf(context.state) < 0) throw new Error('現在は回答期間ではありません。');
    var responseType = vNextClientResponseType_(payload && payload.responseType);
    var direction = responseType === 'CHANGE' ? vNextClientDirection_(payload.direction) : 'NEUTRAL';
    var confidence = responseType === 'CHANGE' ? vNextClientConfidence_(payload.confidence) : '';
    var suppliedType = String(payload && (payload.evidenceType || payload.evidence_type) || '').trim().toUpperCase();
    var changeKind = String(payload && payload.changeKind || '').trim().toLowerCase();
    var evidenceType = 'CHECK_IN';
    if (responseType === 'CHANGE') {
      evidenceType = suppliedType || (changeKind === 'contract' ? 'COMMITMENT' : 'HUMAN_CHANGE');
      if (['COMMITMENT', 'HUMAN_CHANGE'].indexOf(evidenceType) < 0) throw new Error('情報の種類が不正です。');
    } else if (suppliedType && suppliedType !== 'CHECK_IN') {
      throw new Error('変化なし・情報不足はCHECK_INとして保存します。');
    }
    var period = payload && payload.period || {};
    if (responseType === 'CHANGE') {
      if (!String(payload.target || '').trim()) throw new Error('変化の対象を入力してください。');
      if (!String(payload.evidence || '').trim()) throw new Error('根拠を入力してください。');
      if (!direction || direction === 'NEUTRAL') throw new Error('増加または減少を選択してください。');
      if (!confidence) throw new Error('情報確度を選択してください。');
    }
    var record = {
      evidence_id: vNextUuid_(), book_id: context.bookId, client_id: context.clientId,
      fiscal_year: context.fiscalYear, actor_email: context.userEmail,
      response_type: responseType, evidence_type: evidenceType,
      target: String(payload.target || '').trim(),
      target_start_month: period.start ? vNextClientMonth_(period.start) : '',
      target_end_month: period.end ? vNextClientMonth_(period.end) : '',
      direction: direction, amount_mode: String(payload.amountMode || '').toUpperCase(),
      amount_low: vNextClientNumberOrBlank_(payload.amountLow),
      amount_mid: vNextClientNumberOrBlank_(payload.amount),
      amount_high: vNextClientNumberOrBlank_(payload.amountHigh),
      amount_band: String(payload.amountBand || '').toUpperCase(),
      confidence_class: confidence, evidence_text: String(payload.evidence || '').trim(), status: 'SUBMITTED',
      supersedes_evidence_id: String(payload.supersedesEventId || ''),
      source_url: '', source_date: '', expires_at: '', created_at: vNextNowIso_(),
      evidence_quality: '', ai_model: '', prompt_version: '', ai_schema_version: '',
      rule_version: '', applied_amount: '', cap_applied: 0
    };
    vNextAppendRecord_('EVIDENCE_EVENT', record, { spreadsheet: ss });
    return { evidenceId: record.evidence_id, bookId: record.book_id, responseType: record.response_type, createdAt: record.created_at };
  } catch (error) {
    vNextClientLog_('vNextAppendEvidence_ failed', error);
    throw error;
  }
}

/** Employee-side transitions are deliberately narrower than service-side transitions. */
function vNextTransitionState_(request) {
  try {
    var req = request || {};
    var ss = req.spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    var context = vNextGetBookContext_({ spreadsheet: ss, bookId: req.bookId });
    var fromState = context.state;
    var toState = String(req.toState || '').toUpperCase();
    if (req.fromState && String(req.fromState).toUpperCase() !== fromState) throw new Error('状態が更新されています。画面を再読み込みしてください。');
    if (fromState === toState) return { changed: false, fromState: fromState, toState: toState };
    vNextValidateTransition_(fromState, toState, context.role, req, context);
    var isAutomaticReadiness = fromState === 'INPUT_OPEN' && toState === 'READY_TO_RUN' &&
      String(req.reason || '') === 'input_readiness_met' &&
      Number(context.inputStatus.totalCount || 0) > 0 &&
      Number(context.inputStatus.answeredCount || 0) >= Number(context.inputStatus.totalCount || 0);
    var event = {
      state_event_id: vNextUuid_(), book_id: context.bookId, from_state: fromState,
      to_state: toState, reason: String(req.reason || ''), actor_email: context.userEmail,
      actor_role: isAutomaticReadiness ? 'SYSTEM' : context.role,
      related_run_id: String(req.relatedRunId || ''),
      related_plan_version_id: String(req.relatedPlanVersionId || ''), created_at: vNextNowIso_()
    };
    vNextAppendRecord_('STATE_EVENT', event, { spreadsheet: ss, skipLock: Boolean(req.skipLock) });
    vNextWriteLocalRoutingValue_(ss, 'state', toState);
    return { changed: true, stateEventId: event.state_event_id, fromState: fromState, toState: toState };
  } catch (error) {
    vNextClientLog_('vNextTransitionState_ failed', error);
    throw error;
  }
}

function vNextValidateTransition_(fromState, toState, role, request, context) {
  var key = fromState + '>' + toState;
  var req = request || {};
  if (key === 'INPUT_OPEN>READY_TO_RUN') {
    var fullyAnswered = Number(context.inputStatus.totalCount) > 0 && Number(context.inputStatus.answeredCount) >= Number(context.inputStatus.totalCount);
    var deadline = context.inputStatus.dueDate ? new Date(context.inputStatus.dueDate) : null;
    var ownerOverride = role === 'FORECAST_OWNER' && deadline && !isNaN(deadline.getTime()) && deadline.getTime() < new Date().setHours(0, 0, 0, 0) && /^deadline_override:/.test(String(req.reason || ''));
    if (!fullyAnswered && !ownerOverride) throw new Error('全員回答後、または期限後の理由付き進行だけが可能です。');
    return true;
  }
  if (key === 'READY_TO_RUN>RUNNING') {
    if (role !== 'FORECAST_OWNER' || req.internalOperation !== 'CLIENT_QUEUE') throw new Error('予測依頼キューを経由してください。');
    return true;
  }
  if (key === 'DRAFT_READY>SUBMITTED' || key === 'CHANGES_REQUESTED>SUBMITTED') {
    if (role !== 'FORECAST_OWNER' || !String(req.relatedPlanVersionId || '')) throw new Error('予算策定担当の予算案提出だけが可能です。');
    var plan = vNextGetLatestPlanVersion_(context.bookId);
    if (!plan || String(plan.plan_version_id || '') !== String(req.relatedPlanVersionId) || String(plan.status || '') !== 'SUBMITTED') {
      throw new Error('提出済み計画版を確認できません。');
    }
    return true;
  }
  throw new Error('Client Bookから実行できない状態遷移です: ' + key);
}

function vNextCanonicalJson_(value) {
  var stack = [];
  function encode(current, inArray) {
    if (current === null) return 'null';
    if (current instanceof Date) return JSON.stringify(current.toISOString());
    var type = typeof current;
    if (type === 'string') return JSON.stringify(current);
    if (type === 'boolean') return current ? 'true' : 'false';
    if (type === 'number') {
      if (!isFinite(current)) throw new Error('有限でない数値は保存できません。');
      return String(current === 0 ? 0 : current);
    }
    if (type === 'undefined') return inArray ? 'null' : undefined;
    if (type !== 'object') throw new Error('保存できない値です: ' + type);
    if (stack.indexOf(current) >= 0) throw new Error('循環参照は保存できません。');
    stack.push(current);
    var result;
    if (Array.isArray(current)) {
      result = '[' + current.map(function (item) { var part = encode(item, true); return part === undefined ? 'null' : part; }).join(',') + ']';
    } else {
      result = '{' + Object.keys(current).sort().map(function (key) {
        var part = encode(current[key], false);
        return part === undefined ? '' : JSON.stringify(key) + ':' + part;
      }).filter(Boolean).join(',') + '}';
    }
    stack.pop();
    return result;
  }
  return encode(value, false);
}

function vNextSha256Hex_(value) {
  var text = typeof value === 'string' ? value : vNextCanonicalJson_(value);
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return digest.map(function (byte) { var value = byte < 0 ? byte + 256 : byte; return ('0' + value.toString(16)).slice(-2); }).join('');
}

function vNextUuid_() { return Utilities.getUuid(); }
function vNextNowIso_() { return new Date().toISOString(); }

function vNextActiveUserEmail_() {
  try {
    var email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
    if (!email) email = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
    return email;
  } catch (error) {
    vNextClientLog_('user identity unavailable', error);
    return '';
  }
}

function vNextClientReadConfig_(spreadsheet) {
  var sheet = spreadsheet && spreadsheet.getSheetByName(VNEXT_CLIENT_CORE.CONFIG_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return {};
  var values = sheet.getDataRange().getValues();
  var headers = values[0].map(function (value) { return String(value || '').trim().toLowerCase(); });
  var keyColumn = headers.indexOf('key');
  var valueColumn = headers.indexOf('value');
  if (keyColumn < 0 || valueColumn < 0) throw new Error('VN_BOOK_CONFIGの列構成が不正です。');
  var output = {};
  values.slice(1).forEach(function (row) {
    var key = String(row[keyColumn] || '').trim();
    if (key) output[key] = row[valueColumn];
  });
  return output;
}

function vNextWriteLocalRoutingValue_(spreadsheet, key, value) {
  var sheet = spreadsheet.getSheetByName(VNEXT_CLIENT_CORE.CONFIG_SHEET);
  if (!sheet || sheet.getLastRow() < 2) throw new Error('VN_BOOK_CONFIGがありません。');
  var values = sheet.getDataRange().getValues();
  var headers = values[0].map(function (item) { return String(item || '').trim().toLowerCase(); });
  var keyColumn = headers.indexOf('key');
  var valueColumn = headers.indexOf('value');
  for (var index = 1; index < values.length; index++) {
    if (String(values[index][keyColumn] || '').trim() !== key) continue;
    sheet.getRange(index + 1, valueColumn + 1).setValue(value);
    return true;
  }
  throw new Error('VN_BOOK_CONFIGに必要な項目がありません: ' + key);
}

function vNextClientEvidenceForView_(record) {
  var direction = String(record.direction || '').toUpperCase();
  var confidence = String(record.confidence_class || '').toUpperCase();
  return {
    evidenceId: String(record.evidence_id || ''),
    responseType: String(record.response_type || '').toLowerCase(),
    evidenceType: String(record.evidence_type || '').toUpperCase(),
    changeKind: String(record.evidence_type || '').toUpperCase() === 'COMMITMENT' ? 'contract' : 'other',
    target: String(record.target || ''),
    period: vNextClientPeriodLabel_(record.target_start_month, record.target_end_month),
    direction: direction === 'UP' ? 'increase' : direction === 'DOWN' ? 'decrease' : '',
    amountMode: String(record.amount_mode || '').toLowerCase(),
    amount: vNextClientNumberOrBlank_(record.amount_mid),
    amountBand: String(record.amount_band || '').toLowerCase(),
    evidence: String(record.evidence_text || ''),
    confidence: confidence === 'CONFIRMED_FACT' ? 'confirmed' : confidence === 'LIKELY' ? 'likely' : confidence === 'HYPOTHESIS' ? 'hypothesis' : '',
    createdAt: vNextClientDateOnly_(record.created_at)
  };
}

function vNextClientPeriodLabel_(start, end) {
  if (!start) return '';
  return String(start) === String(end || start) ? String(start) : String(start) + '〜' + String(end);
}

function vNextClientResponseType_(value) {
  var map = { change: 'CHANGE', no_change: 'NO_CHANGE', unknown: 'UNKNOWN', CHANGE: 'CHANGE', NO_CHANGE: 'NO_CHANGE', UNKNOWN: 'UNKNOWN' };
  var result = map[String(value || '')];
  if (!result) throw new Error('回答種別を選択してください。');
  return result;
}

function vNextClientDirection_(value) {
  var map = { increase: 'UP', decrease: 'DOWN', UP: 'UP', DOWN: 'DOWN' };
  return map[String(value || '')] || '';
}

function vNextClientConfidence_(value) {
  var map = { confirmed: 'CONFIRMED_FACT', likely: 'LIKELY', hypothesis: 'HYPOTHESIS', CONFIRMED_FACT: 'CONFIRMED_FACT', LIKELY: 'LIKELY', HYPOTHESIS: 'HYPOTHESIS' };
  return map[String(value || '')] || '';
}

function vNextClientMonth_(value) {
  var date = vNextClientDate_(value);
  return date.getFullYear() + '-' + ('0' + (date.getMonth() + 1)).slice(-2);
}

function vNextClientDateOnly_(value) {
  if (!value) return '';
  var date = vNextClientDate_(value);
  return date.getFullYear() + '-' + ('0' + (date.getMonth() + 1)).slice(-2) + '-' + ('0' + date.getDate()).slice(-2);
}

function vNextClientDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return new Date(value.getTime());
  var text = String(value || '');
  var match = text.match(/^(\d{4})[-\/]?(\d{2})[-\/]?(\d{2})/);
  var date = match ? new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3])) : new Date(text);
  if (isNaN(date.getTime())) throw new Error('日付を確認できません: ' + text);
  return date;
}

function vNextClientEmails_(value) {
  return String(value || '').split(',').map(vNextClientNormalizeEmail_).filter(Boolean).filter(function (email, index, all) { return all.indexOf(email) === index; });
}

function vNextClientNormalizeEmail_(value) { return String(value || '').trim().toLowerCase(); }

function vNextClientEmailDomain_(value) {
  var email = vNextClientNormalizeEmail_(value);
  var separator = email.lastIndexOf('@');
  return separator > 0 && separator < email.length - 1 ? email.slice(separator + 1) : '';
}

function vNextClientNormalizeDomain_(value) {
  var domain = String(value || '').trim().toLowerCase().replace(/^@/, '');
  return /^[a-z0-9](?:[a-z0-9.-]*[a-z0-9])?$/.test(domain) && domain.indexOf('.') > 0 ? domain : '';
}

function vNextClientJsonArray_(value) {
  if (Array.isArray(value)) return value.slice();
  if (!value) return [];
  try { var parsed = JSON.parse(String(value)); return Array.isArray(parsed) ? parsed : []; }
  catch (error) { return String(value).split(',').map(function (item) { return item.trim(); }).filter(Boolean); }
}

function vNextClientNumberOrBlank_(value) {
  return value !== '' && value !== null && value !== undefined && isFinite(Number(value)) ? Math.trunc(Number(value)) : '';
}

function vNextClientCellValue_(value) {
  if (value === undefined || value === null) return '';
  if (value instanceof Date) return value.toISOString();
  if (typeof value === 'object') return vNextCanonicalJson_(value);
  return value;
}

function vNextClientWithDocumentLock_(operation) {
  var lock = LockService.getDocumentLock();
  var acquired = false;
  try {
    lock.waitLock(30000);
    acquired = true;
    return operation();
  } finally {
    if (acquired) lock.releaseLock();
  }
}

function vNextClientLog_(message, error) {
  Logger.log('[vNext Client] ' + message + (error ? ': ' + String(error && error.stack || error) : ''));
}
