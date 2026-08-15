/**
 * Forecast vNext engine: strict actual-data bridge and coherent seeded simulation.
 * Depends only on VNext_Core.js; legacy Forecast_Agent.js remains untouched.
 */

var VNEXT_ENGINE = Object.freeze({
  VERSION: 'vnext-engine-0.3.1',
  MIN_HISTORY_YEARS: 5,
  MAX_HISTORY_YEARS: 8,
  DEFAULT_SIMULATIONS: 2000,
  MIN_SIMULATIONS: 200,
  MAX_SIMULATIONS: 10000,
  AI_MAX_ABS_EFFECT: 0.05,
  MISSING_RESPONSE_UNCERTAINTY: 0.30,
  INFORMATION_GAP_UNCERTAINTY: 0.20,
  KNOWN_SPOT_SUPPRESSION_RATE: 0.50,
  DEFAULT_MONTH_COMMON_SHOCK_SIGMA: 0.10,
  MIN_MONTH_COMMON_SHOCK_SIGMA: 0.04,
  MAX_MONTH_COMMON_SHOCK_SIGMA: 0.30,
  REFERENCE_MAX_STRENGTH: 0.35,
  REFERENCE_MIN_GROWTH: -1,
  REFERENCE_MAX_GROWTH: 3,
  REFERENCE_MAX_GROWTH_STD: 2,
  REFERENCE_PRIOR_PILOT_ENABLED: false,
  RUN_IDEMPOTENCY_CONTRACT: 'FORECAST_RUN_IDEMPOTENCY_V1',
  SOURCE_ID_PROPERTY: 'VNEXT_ZAC_SOURCE_SPREADSHEET_ID',
  SOURCE_COLUMNS: Object.freeze({
    client: 41,
    serviceCategory: 46,
    product: 50,
    actualDate: 57,
    amount: 66
  }),
  SOURCE_HEADERS: Object.freeze({
    client: 'クライアント名',
    serviceCategory: 'サービスカテゴリ',
    product: '売上区分',
    actualDate: '売上日',
    amount: '金額'
  })
});

/** Public employee/Admin entry. Client books are queued; only the private worker executes queued jobs. */
function vNextEngineRunForecast(request) {
  var authorized = vNextAuthorizeForecastRequest_(request || {});
  if (vNextShouldQueueClientRun_(authorized) && typeof vNextQueueClientForecastRequest === 'function') {
    return vNextQueueClientForecastRequest(authorized);
  }
  return vNextExecuteForecastRun_(authorized);
}

function vNextShouldQueueClientRun_(request) {
  if (typeof SpreadsheetApp === 'undefined') return false;
  try {
    var spreadsheet = request.spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    var routing = vNextReadLocalRoutingConfig_(spreadsheet);
    return String(routing.mode || '').toUpperCase() === 'CLIENT';
  } catch (error) {
    vNextLog_('Client queue routing detection skipped', error);
    return false;
  }
}

/** Private Admin Hub worker entry. It is not callable through google.script.run. */
function vNextRunForecast_(request) {
  var req = request || {};
  if (String(req.internalOperation || '').toUpperCase() !== 'ADMIN_JOB') {
    throw new Error('vNextRunForecast_ is reserved for an ADMIN_JOB worker.');
  }
  return vNextExecuteForecastRun_(vNextAuthorizeAdminJobRequest_(req));
}

function vNextExecuteForecastRun_(req) {
  var started = new Date();
  var normalized;
  var stateMoved = false;
  var replayResolved = false;
  var deterministicSuccessPersisted = false;
  try {
    normalized = vNextNormalizeRunRequest_(req);
    var existing = normalized.persist && normalized.runIdentity
      ? vNextResolveExistingDeterministicRun_(normalized)
      : null;
    if (existing && existing.result) {
      replayResolved = true;
      var replay = vNextCloneForecastResult_(existing.result);
      replay.idempotent = true;
      replay.idempotentReplay = true;
      if (normalized.persist && normalized.manageState && normalized.initialState === 'RUNNING') {
        vNextTransitionState_({
          bookId: normalized.bookId,
          toState: 'DRAFT_READY',
          reason: 'Forecast run resumed from an existing successful result',
          relatedRunId: replay.runId,
          actorEmail: normalized.createdBy,
          actorRole: 'SYSTEM',
          internalOperation: 'FORECAST_ENGINE',
          spreadsheet: normalized.spreadsheet
        });
      }
      return replay;
    }
    if (normalized.persist && normalized.manageState && normalized.initialState !== 'RUNNING') {
      vNextTransitionState_({
        bookId: normalized.bookId,
        toState: 'RUNNING',
        reason: 'Forecast run started',
        actorEmail: normalized.createdBy,
        actorRole: 'SYSTEM',
        internalOperation: 'FORECAST_ENGINE',
        spreadsheet: normalized.spreadsheet
      });
      stateMoved = true;
    }
    var result = vNextSimulateForecast_(normalized);
    result.durationMs = new Date().getTime() - started.getTime();
    if (normalized.persist) {
      // Re-check immediately before append. Admin normally holds the job lock,
      // but this closes the practical retry window without changing Core locks.
      var lateExisting = normalized.runIdentity
        ? vNextResolveExistingDeterministicRun_(normalized)
        : null;
      if (lateExisting && lateExisting.result) {
        replayResolved = true;
        result = vNextCloneForecastResult_(lateExisting.result);
        result.idempotent = true;
        result.idempotentReplay = true;
      } else {
        vNextPersistForecastRun_(result, normalized);
        deterministicSuccessPersisted = Boolean(normalized.runIdentity);
      }
    }
    if (normalized.persist && normalized.manageState) {
      vNextTransitionState_({
        bookId: normalized.bookId,
        toState: 'DRAFT_READY',
        reason: 'Forecast run completed',
        relatedRunId: result.runId,
        actorEmail: normalized.createdBy,
        actorRole: 'SYSTEM',
        internalOperation: 'FORECAST_ENGINE',
        spreadsheet: normalized.spreadsheet
      });
    }
    return result;
  } catch (error) {
    vNextLog_('vNextEngineRunForecast failed', error);
    var failureContext = normalized || vNextBuildPreflightFailureContext_(req);
    var failureInfo = vNextForecastFailureInfo_(error);
    // A deterministic identity conflict must be side-effect free. Likewise, if
    // SUCCESS was already found/appended, leave RUNNING intact so the same job
    // can resume only the DRAFT_READY phase without adding a false FAILED run.
    var suppressFailureMutation = Boolean(
      error && error.vNextRunIdentityFailure || replayResolved || deterministicSuccessPersisted
    );
    if (!suppressFailureMutation && failureContext && failureContext.persist) {
      try { vNextPersistFailedRun_(failureContext, error, started); } catch (logError) { vNextLog_('Failed run audit failed', logError); }
    }
    if (!suppressFailureMutation && failureContext && failureContext.persist && failureContext.manageState && (stateMoved || failureContext.initialState === 'RUNNING')) {
      try {
        vNextTransitionState_({
          bookId: failureContext.bookId,
          toState: 'READY_TO_RUN',
          reason: 'forecast_failed/' + failureInfo.code + ': ' + failureInfo.userMessage,
          actorEmail: failureContext.createdBy,
          actorRole: 'SYSTEM',
          internalOperation: 'FORECAST_ENGINE',
          spreadsheet: failureContext.spreadsheet
        });
      } catch (stateError) {
        vNextLog_('Run failure state recovery failed', stateError);
      }
    }
    throw error;
  }
}

/**
 * Converts known technical failures into a stable employee-facing code and next action.
 * The original message is retained only for Admin audit/debugging.
 */
function vNextForecastFailureInfo_(error) {
  var technical = String(error && error.message ? error.message : error || 'Unknown forecast error').slice(0, 1000);
  var info = {
    code: 'FORECAST_PROCESSING_FAILED',
    userMessage: '予測を完了できませんでした。入力内容は保存されています。管理者が処理状況を確認します。',
    nextAction: 'しばらくしても状態が変わらない場合は、クライアント名と対象年度を管理者へ連絡してください。',
    retryRecommended: true,
    technicalMessage: technical
  };
  var history = /At least\s+(\d+)\s+fiscal years[^;]*;\s*found\s+(\d+)/i.exec(technical);
  if (history) {
    info.code = 'INSUFFICIENT_CONFIRMED_HISTORY';
    info.userMessage = '予測に必要な確定実績が不足しています（必要' + history[1] + '年度、確認できた実績' + history[2] + '年度）。';
    info.nextAction = 'ZACのクライアント表記と確定実績の登録状況を管理者に確認してください。情報入力はそのまま保持されます。';
    info.retryRecommended = false;
    return info;
  }
  if (/source header mismatch|source schema has blank required headers|planned-date header/i.test(technical)) {
    info.code = 'ACTUAL_SOURCE_SCHEMA_MISMATCH';
    info.userMessage = '確定実績データの列構成を確認できないため、予測を停止しました。';
    info.nextAction = '社員側で再入力する必要はありません。管理者がZAC取込の列名と版を確認します。';
    info.retryRecommended = false;
    return info;
  }
  if (/plannedDate|dateSource is not actual|is not confirmed|after cutoff|missing actualDate/i.test(technical)) {
    info.code = 'ACTUAL_DATA_CONTRACT_FAILED';
    info.userMessage = '予定データまたは締切後データが含まれるため、安全のため予測を停止しました。';
    info.nextAction = '社員側で再入力する必要はありません。管理者が確定実績と情報締切を確認します。';
    info.retryRecommended = false;
    return info;
  }
  if (/permission|not have access|access denied|openById/i.test(technical)) {
    info.code = 'ACTUAL_SOURCE_ACCESS_FAILED';
    info.userMessage = '確定実績データを読み取れませんでした。';
    info.nextAction = '管理者がZAC実績元と実行アカウントの権限を確認します。';
    info.retryRecommended = true;
  }
  return info;
}

function vNextBuildPreflightFailureContext_(request) {
  if (!request) return null;
  var asOf;
  try { asOf = vNextParseDate_(request.asOf || request.as_of || new Date(), 'as_of'); } catch (error) { asOf = new Date(); }
  var cutoff = vNextCutoffFromAsOf_(asOf);
  var seed = request.seed !== undefined && request.seed !== null ? vNextNormalizeSeed_(request.seed) : 0;
  return {
    runId: String(request.runId || vNextUuid_()),
    bookId: String(request.bookId || request.book_id || ''),
    clientId: String(request.clientId || request.client_id || ''),
    clientName: String(request.clientName || request.client_name || ''),
    fiscalYear: Number(request.fiscalYear || request.fiscal_year || 0),
    asOf: asOf,
    cutoff: cutoff,
    seed: seed,
    inputDataHash: vNextSha256Hex_({
      stage: 'PREFLIGHT_FAILED', bookId: String(request.bookId || request.book_id || ''),
      fiscalYear: Number(request.fiscalYear || request.fiscal_year || 0),
      asOf: vNextFormatDateOnly_(asOf), modelReleaseId: String(request.modelReleaseId || ''),
      coreVersion: VNEXT_CORE.VERSION, engineVersion: VNEXT_ENGINE.VERSION
    }),
    modelReleaseId: String(request.modelReleaseId || ''),
    simulationCount: Number(request.simulationCount || VNEXT_ENGINE.DEFAULT_SIMULATIONS),
    previousRunId: String(request.previousRunId || ''),
    runIdentity: request.authorizedRunIdentity || request.runIdentity || null,
    persist: request.persist !== false && typeof SpreadsheetApp !== 'undefined',
    manageState: request.manageState !== false,
    initialState: String(request.initialState || '').toUpperCase(),
    spreadsheet: request.spreadsheet,
    createdBy: String(request.createdBy || request.requestedBy || vNextActiveUserEmail_()).toLowerCase()
  };
}

function vNextAuthorizeForecastRequest_(request) {
  vNextAssertNoEmployeeRunIdentity_(request);
  if (typeof SpreadsheetApp === 'undefined') return request;
  var actor = vNextActiveUserEmail_();
  var context = vNextGetBookContext_({
    bookId: request.bookId || request.book_id,
    spreadsheet: request.spreadsheet,
    userEmail: actor
  });
  if (context.role !== 'FORECAST_OWNER' && context.role !== 'ADMIN') {
    throw new Error('Only the Forecast Owner or administrator can run a forecast.');
  }
  if (String(context.state || '').toUpperCase() !== 'READY_TO_RUN') {
    throw new Error('Forecast can run only from READY_TO_RUN; current state=' + context.state);
  }
  var safe = {};
  Object.keys(request).forEach(function (key) { safe[key] = request[key]; });
  safe.bookId = context.bookId;
  safe.clientId = context.clientId;
  safe.clientName = context.clientName;
  safe.fiscalYear = context.fiscalYear;
  safe.asOf = context.asOf || request.asOf;
  safe.createdBy = actor;
  safe.requestedBy = actor;
  if (context.role !== 'ADMIN' || request.allowAdminOverrides !== true) {
    [
      'actualRecords', 'actualFetcher', 'legacyFetcher', 'legacySourceGuaranteesActualDate',
      'commitmentEvents', 'objectiveEvents', 'changeEvents', 'humanEvents', 'aiEvents',
      'referencePrior', 'parameters', 'sourceSpreadsheetId', 'seed', 'simulationCount',
      'missingResponseRate', 'modelReleaseId', 'runId', 'run_id',
      'idempotencyKey', 'idempotency_key', 'authorizedRunIdentity',
      'serverRunIdentityAuthorized', 'authorizedPreviousRunIdPinned', 'previousRunId'
    ].forEach(function (key) { delete safe[key]; });
  }
  safe.missingResponseRate = vNextMissingResponseRateFromContext_(context);
  safe.informationGapRate = Number(context.inputStatus && context.inputStatus.informationGapRate || 0);
  // Employee callers cannot influence the degradation policy.
  safe.aiUnavailable = false;
  safe.aiUnavailableReason = '';
  safe.evidenceResponseCounts = {
    change: Number(context.inputStatus && context.inputStatus.changeCount || 0),
    unknown: Number(context.inputStatus && context.inputStatus.unknownCount || 0),
    noChange: Number(context.inputStatus && context.inputStatus.noChangeCount || 0)
  };
  safe.modelReleaseId = String(context.version && context.version.modelReleaseId || '');
  safe.schemaVersion = String(context.version && context.version.schema || VNEXT_CORE.SCHEMA_VERSION);
  safe.templateVersion = String(context.version && context.version.template || VNEXT_CORE.DEFAULT_TEMPLATE_VERSION);
  safe.initialState = String(context.state || '').toUpperCase();
  safe.authorizationContext = { role: context.role, state: context.state, actor: actor };
  return safe;
}

function vNextAuthorizeAdminJobRequest_(request) {
  if (typeof SpreadsheetApp === 'undefined') throw new Error('ADMIN_JOB authorization requires Apps Script.');
  var actor = vNextActiveUserEmail_();
  var context = vNextGetBookContext_({
    bookId: request.bookId || request.book_id,
    spreadsheet: request.spreadsheet,
    userEmail: actor
  });
  if (context.role !== 'ADMIN') throw new Error('ADMIN_JOB requires an administrator execution identity.');
  if (String(context.state || '').toUpperCase() !== 'RUNNING') {
    throw new Error('ADMIN_JOB requires a pre-claimed RUNNING state; current state=' + context.state);
  }
  var runIdentity = vNextAuthorizeAdminRunIdentity_(request, context.bookId);
  var safe = {};
  Object.keys(request).forEach(function (key) { safe[key] = request[key]; });
  [
    'actualRecords', 'actualFetcher', 'legacyFetcher', 'legacySourceGuaranteesActualDate',
    'commitmentEvents', 'objectiveEvents', 'changeEvents', 'humanEvents', 'aiEvents',
    'referencePrior', 'parameters', 'sourceSpreadsheetId', 'seed', 'simulationCount',
    'modelReleaseId', 'runId', 'run_id', 'idempotencyKey', 'idempotency_key',
    'authorizedRunIdentity', 'serverRunIdentityAuthorized',
    'authorizedPreviousRunIdPinned', 'previousRunId', 'allowAdminOverrides'
  ].forEach(function (key) { delete safe[key]; });
  safe.bookId = context.bookId;
  safe.clientId = context.clientId;
  safe.clientName = context.clientName;
  safe.fiscalYear = context.fiscalYear;
  // Client request eventはHub harvest時にhash検証済み。provision時のBOOK_META.as_ofより
  // 実際の依頼日を優先し、依頼日までの現場evidenceを取りこぼさない。
  safe.asOf = vNextResolveAdminJobAsOf_(request, context);
  safe.createdBy = actor;
  safe.requestedBy = actor;
  safe.missingResponseRate = vNextMissingResponseRateFromContext_(context);
  safe.informationGapRate = Number(context.inputStatus && context.inputStatus.informationGapRate || 0);
  // AI is an optional evidence layer.  Only the Admin-owned worker may mark a
  // terminal dependency failure and widen uncertainty while continuing with
  // AI=0; employee payloads cannot set this policy flag.
  safe.aiUnavailable = request.aiUnavailable === true;
  safe.aiUnavailableReason = safe.aiUnavailable
    ? String(request.aiUnavailableReason || 'AI_RESEARCH_UNAVAILABLE').slice(0, 200)
    : '';
  if (safe.aiUnavailable) safe.informationGapRate = Math.max(safe.informationGapRate, 0.15);
  safe.evidenceResponseCounts = {
    change: Number(context.inputStatus && context.inputStatus.changeCount || 0),
    unknown: Number(context.inputStatus && context.inputStatus.unknownCount || 0),
    noChange: Number(context.inputStatus && context.inputStatus.noChangeCount || 0)
  };
  safe.modelReleaseId = String(context.version && context.version.modelReleaseId || '');
  safe.schemaVersion = String(context.version && context.version.schema || VNEXT_CORE.SCHEMA_VERSION);
  safe.templateVersion = String(context.version && context.version.template || VNEXT_CORE.DEFAULT_TEMPLATE_VERSION);
  var storedAttempt = runIdentity
    ? vNextInspectStoredDeterministicRun_(runIdentity, { spreadsheet: request.spreadsheet })
    : null;
  var previousRun = storedAttempt && storedAttempt.found
    ? null
    : vNextLoadPreviousRunSummary_(context.bookId, '', request.spreadsheet);
  safe.previousRunId = storedAttempt && storedAttempt.found
    ? String(storedAttempt.previousRunId || '')
    : (previousRun && previousRun.runId || '');
  safe.authorizedPreviousRunIdPinned = Boolean(storedAttempt && storedAttempt.found);
  if (runIdentity) {
    safe.runId = runIdentity.runId;
    safe.idempotencyKey = runIdentity.idempotencyKey;
    safe.authorizedRunIdentity = runIdentity;
    safe.serverRunIdentityAuthorized = true;
    // A deterministic Admin attempt is always an append-only Hub run. Allowing
    // callers to disable persistence would defeat replay verification.
    safe.persist = true;
    safe.manageState = true;
  }
  safe.initialState = 'RUNNING';
  safe.internalOperation = 'ADMIN_JOB';
  safe.authorizationContext = { role: 'ADMIN', state: 'RUNNING', actor: actor };
  return safe;
}

function vNextResolveAdminJobAsOf_(request, context) {
  var req = request || {};
  var ctx = context || {};
  return req.asOf || req.as_of || ctx.asOf || ctx.as_of || vNextFormatDateOnly_(new Date());
}

/** Employee/browser entry points may never choose the immutable run identity. */
function vNextAssertNoEmployeeRunIdentity_(request) {
  var req = request || {};
  var forbidden = [
    'runId', 'run_id', 'idempotencyKey', 'idempotency_key',
    'authorizedRunIdentity', 'serverRunIdentityAuthorized', 'authorizedPreviousRunIdPinned'
  ];
  var supplied = forbidden.filter(function (key) {
    return Object.prototype.hasOwnProperty.call(req, key);
  });
  if (supplied.length) {
    throw new Error('runId/idempotencyKey are server-authorized ADMIN_JOB fields and cannot be supplied by the employee forecast API.');
  }
  return true;
}

/** Deterministic identity builder used by the Admin queue and rollback worker. */
function vNextEngineBuildAdminRunIdentity_(bookId, idempotencyKey) {
  var book = vNextStrictRunIdentityText_(bookId, 'bookId', 1, 200);
  var key = vNextStrictRunIdentityText_(idempotencyKey, 'idempotencyKey', 8, 500);
  var identityHash = vNextSha256Hex_({
    contractVersion: VNEXT_ENGINE.RUN_IDEMPOTENCY_CONTRACT,
    bookId: book,
    idempotencyKey: key
  });
  return {
    contractVersion: VNEXT_ENGINE.RUN_IDEMPOTENCY_CONTRACT,
    runId: 'RUN-' + identityHash.slice(0, 32).toUpperCase(),
    bookId: book,
    idempotencyKey: key,
    identityHash: identityHash
  };
}

function vNextStrictRunIdentityText_(value, fieldName, minLength, maxLength) {
  var raw = String(value === undefined || value === null ? '' : value);
  var text = raw.trim();
  if (raw !== text || text.length < minLength || text.length > maxLength || /[\u0000-\u001f\u007f]/.test(text)) {
    throw new Error(fieldName + ' is invalid for deterministic run identity.');
  }
  return text;
}

function vNextRunIdentityAlias_(request, camel, snake) {
  var req = request || {};
  var hasCamel = Object.prototype.hasOwnProperty.call(req, camel);
  var hasSnake = Object.prototype.hasOwnProperty.call(req, snake);
  if (hasCamel && hasSnake && String(req[camel]) !== String(req[snake])) {
    throw new Error(camel + ' and ' + snake + ' disagree.');
  }
  return { present: hasCamel || hasSnake, value: hasCamel ? req[camel] : req[snake] };
}

function vNextAuthorizeAdminRunIdentity_(request, bookId) {
  var run = vNextRunIdentityAlias_(request, 'runId', 'run_id');
  var key = vNextRunIdentityAlias_(request, 'idempotencyKey', 'idempotency_key');
  if (!run.present && !key.present) return null; // Legacy Admin jobs remain valid during wiring rollout.
  if (!run.present || !key.present) throw new Error('ADMIN_JOB deterministic identity requires both runId and idempotencyKey.');
  var expected = vNextEngineBuildAdminRunIdentity_(bookId, key.value);
  var suppliedRunId = vNextStrictRunIdentityText_(run.value, 'runId', 8, 100);
  if (suppliedRunId !== expected.runId) {
    throw new Error('ADMIN_JOB runId does not match the deterministic book/idempotencyKey identity.');
  }
  return expected;
}

function vNextNormalizeAuthorizedRunIdentity_(request, bookId) {
  var run = vNextRunIdentityAlias_(request, 'runId', 'run_id');
  var key = vNextRunIdentityAlias_(request, 'idempotencyKey', 'idempotency_key');
  var marker = request && request.authorizedRunIdentity;
  if (!run.present && !key.present && !marker) return null;
  if (String(request.internalOperation || '').toUpperCase() !== 'ADMIN_JOB' ||
      request.serverRunIdentityAuthorized !== true || !marker) {
    throw new Error('Deterministic run identity is accepted only after ADMIN_JOB server authorization.');
  }
  var expected = vNextAuthorizeAdminRunIdentity_(request, bookId);
  if (!expected || String(marker.contractVersion || '') !== expected.contractVersion ||
      String(marker.runId || '') !== expected.runId ||
      String(marker.idempotencyKey || '') !== expected.idempotencyKey ||
      String(marker.identityHash || '') !== expected.identityHash ||
      String(marker.bookId || '') !== expected.bookId) {
    throw new Error('Authorized deterministic run identity marker is inconsistent.');
  }
  return expected;
}

function vNextFinalizeRunIdentity_(identity, lineage) {
  if (!identity) return null;
  var output = {};
  Object.keys(identity).forEach(function (key) { output[key] = identity[key]; });
  output.lineageHash = vNextSha256Hex_({
    contractVersion: identity.contractVersion,
    identityHash: identity.identityHash,
    lineage: lineage
  });
  return output;
}

function vNextRunIdentityError_(message) {
  var error = new Error(message);
  error.vNextRunIdentityFailure = true;
  return error;
}

function vNextFindForecastRunRecords_(bookId, runId, options) {
  var opt = options || {};
  return vNextReadRecords_('FORECAST_RUN', { spreadsheet: opt.spreadsheet }).filter(function (row) {
    return String(row.book_id || '') === String(bookId || '') &&
      String(row.run_id || '') === String(runId || '');
  });
}

/** Minimal stored-attempt inspection; exact canonical checks happen after normalize. */
function vNextInspectStoredDeterministicRun_(identity, options) {
  var rows = vNextFindForecastRunRecords_(identity.bookId, identity.runId, options || {});
  if (!rows.length) return { found: false, runId: identity.runId, rows: [] };
  var firstFingerprint = '';
  var successes = [];
  rows.forEach(function (row) {
    var lenses = vNextParseJsonValue_(row.lens_json, {});
    var stored = lenses.runIdentity || {};
    if (Number(row.is_official || 0) === 1 || String(row.official_vintage_id || '') ||
        ['SUCCESS', 'FAILED'].indexOf(String(row.status || '').toUpperCase()) < 0 ||
        String(stored.contractVersion || '') !== identity.contractVersion ||
        String(stored.runId || '') !== identity.runId ||
        String(stored.bookId || '') !== identity.bookId ||
        String(stored.idempotencyKey || '') !== identity.idempotencyKey ||
        String(stored.identityHash || '') !== identity.identityHash) {
      throw vNextRunIdentityError_('Existing FORECAST_RUN conflicts with the deterministic run identity.');
    }
    var fingerprint = vNextCanonicalJson_({
      bookId: String(row.book_id || ''), clientId: String(row.client_id || ''),
      fiscalYear: Number(row.fiscal_year), asOf: String(row.as_of || ''), cutoff: String(row.cutoff || ''),
      inputDataHash: String(row.input_data_hash || ''), modelReleaseId: String(row.model_release_id || ''),
      previousRunId: String(row.previous_run_id || ''), lineageHash: String(stored.lineageHash || '')
    });
    if (!firstFingerprint) firstFingerprint = fingerprint;
    if (fingerprint !== firstFingerprint) {
      throw vNextRunIdentityError_('Existing FORECAST_RUN rows disagree on canonical input or lineage.');
    }
    if (String(row.status || '').toUpperCase() === 'SUCCESS') successes.push(row);
  });
  if (successes.length > 1) throw vNextRunIdentityError_('Multiple SUCCESS rows share one deterministic runId.');
  if (successes.length) vNextAssertReplayableSuccessRecord_(successes[0]);
  var first = rows[0];
  var firstLenses = vNextParseJsonValue_(first.lens_json, {});
  return {
    found: true,
    runId: identity.runId,
    statuses: rows.map(function (row) { return String(row.status || '').toUpperCase(); }),
    inputDataHash: String(first.input_data_hash || ''),
    lineageHash: String(firstLenses.runIdentity && firstLenses.runIdentity.lineageHash || ''),
    previousRunId: String(first.previous_run_id || ''),
    result: successes.length ? vNextForecastRecordToResult_(successes[0]) : null,
    rows: rows
  };
}

function vNextResolveExistingDeterministicRun_(normalized) {
  var inspection;
  try {
    inspection = vNextInspectStoredDeterministicRun_(normalized.runIdentity, { spreadsheet: normalized.spreadsheet });
    if (!inspection.found) return inspection;
    if (!inspection.inputDataHash || inspection.inputDataHash !== normalized.inputDataHash ||
        !inspection.lineageHash || inspection.lineageHash !== normalized.runIdentity.lineageHash) {
      throw vNextRunIdentityError_('Existing runId has a different canonical input hash or immutable lineage.');
    }
    return inspection;
  } catch (error) {
    if (error && error.vNextRunIdentityFailure) throw error;
    throw vNextRunIdentityError_('Deterministic run lookup failed closed: ' + String(error && error.message || error));
  }
}

function vNextAssertReplayableSuccessRecord_(row) {
  var quarters = vNextParseJsonValue_(row.quarter_json, []);
  var months = vNextParseJsonValue_(row.month_json, []);
  var finite = ['seed', 'simulation_count', 'history_baseline', 'commitment_delta', 'reference_delta',
    'human_delta', 'ai_delta', 'objective_forecast', 'system_recommended', 'p10', 'p50', 'p90'];
  if (!/^[a-f0-9]{64}$/i.test(String(row.input_data_hash || '')) ||
      quarters.length !== 4 || months.length !== 12 ||
      finite.some(function (key) { return !vNextIsFiniteNumber_(row[key]); })) {
    throw vNextRunIdentityError_('Existing SUCCESS FORECAST_RUN is incomplete and cannot be replayed.');
  }
  var tolerance = Math.max(0.000001, Math.abs(Number(row.p50 || 0)) * 0.000000001);
  var monthP50 = vNextSum_(months.map(function (item) { return Number(item.p50); }));
  var quarterP50 = vNextSum_(quarters.map(function (item) { return Number(item.p50); }));
  var objective = Number(row.history_baseline) + Number(row.commitment_delta) + Number(row.reference_delta);
  var system = Number(row.objective_forecast) + Number(row.human_delta) + Number(row.ai_delta);
  if (Math.abs(monthP50 - Number(row.p50)) > tolerance ||
      Math.abs(quarterP50 - Number(row.p50)) > tolerance ||
      Math.abs(objective - Number(row.objective_forecast)) > tolerance ||
      Math.abs(system - Number(row.system_recommended)) > tolerance ||
      Math.abs(Number(row.system_recommended) - Number(row.p50)) > tolerance) {
    throw vNextRunIdentityError_('Existing SUCCESS FORECAST_RUN fails stored arithmetic/coherence validation.');
  }
  return true;
}

/** Private Admin API: inspect an immutable attempt before resuming later phases. */
function vNextEngineLookupRunForResume_(request) {
  var req = request || {};
  if (String(req.internalOperation || '').toUpperCase() !== 'ADMIN_JOB') {
    throw new Error('Run resume lookup is reserved for an ADMIN_JOB worker.');
  }
  var bookId = String(req.bookId || req.book_id || '');
  if (!bookId) throw new Error('bookId is required.');
  vNextRequireActualRoleForBook_(bookId, req.spreadsheet, ['ADMIN']);
  var identity = vNextAuthorizeAdminRunIdentity_(req, bookId);
  if (!identity) throw new Error('runId and idempotencyKey are required for run resume lookup.');
  var inspected = vNextInspectStoredDeterministicRun_(identity, { spreadsheet: req.spreadsheet });
  if (inspected.found && req.expectedInputDataHash &&
      String(req.expectedInputDataHash) !== inspected.inputDataHash) {
    throw vNextRunIdentityError_('Stored run input hash differs from the resume expectation.');
  }
  return {
    found: inspected.found,
    runId: identity.runId,
    statuses: inspected.statuses || [],
    inputDataHash: inspected.inputDataHash || '',
    lineageHash: inspected.lineageHash || '',
    previousRunId: inspected.previousRunId || '',
    hasSuccess: Boolean(inspected.result),
    resumablePhase: inspected.result ? 'DRAFT_READY_SYNC' : (inspected.found ? 'FORECAST_RETRY' : 'FORECAST_EXECUTE'),
    result: inspected.result ? vNextCloneForecastResult_(inspected.result) : null
  };
}

function vNextCloneForecastResult_(value) {
  return JSON.parse(JSON.stringify(value));
}

function vNextNormalizeRunRequest_(request) {
  var hasTrustedRollback = Boolean(
    request.trustedReuseSeedFromRunId || request.trustedRollbackContext || request.trustedAllowedDelayedAiRequestIds
  );
  if (hasTrustedRollback && String(request.internalOperation || '').toUpperCase() !== 'ADMIN_JOB') {
    throw new Error('Trusted rollback fields are accepted only from an ADMIN_JOB worker.');
  }
  var allowedDelayedAiRequestIds = hasTrustedRollback
    ? vNextNormalizeTrustedStringArray_(request.trustedAllowedDelayedAiRequestIds, 'trustedAllowedDelayedAiRequestIds')
    : [];
  var context = null;
  if ((!request.fiscalYear && !request.fiscal_year) || (!request.clientName && !request.client_name)) {
    try {
      context = vNextGetBookContext_({
        bookId: request.bookId || request.book_id,
        spreadsheet: request.spreadsheet,
        userEmail: request.requestedBy || request.createdBy
      });
    } catch (contextError) {
      vNextLog_('Run request context hydration failed', contextError);
    }
  }
  var fiscalYear = Number(request.fiscalYear || request.fiscal_year || (context && context.fiscalYear));
  if (!isFinite(fiscalYear)) throw new Error('fiscalYear is required.');
  var asOf = vNextParseDate_(request.asOf || request.as_of || (context && context.asOf) || new Date(), 'as_of');
  var cutoff = vNextCutoffFromAsOf_(asOf);
  var bookId = String(request.bookId || request.book_id || '');
  if (!bookId) throw new Error('bookId is required.');
  var authorizedRunIdentity = vNextNormalizeAuthorizedRunIdentity_(request, bookId);
  var trustedRollbackSource = hasTrustedRollback
    ? vNextResolveTrustedRollbackSource_(request, bookId, fiscalYear, asOf, cutoff)
    : null;
  var exactRollbackEvidenceIds = trustedRollbackSource
    ? vNextTrustedRollbackEvidenceIds_(trustedRollbackSource, request.trustedRollbackContext)
    : [];
  var hydratedEvents = { commitment: [], objective: [], human: [], ai: [], counts: { change: 0, noChange: 0, unknown: 0 } };
  if (request.commitmentEvents === undefined || request.objectiveEvents === undefined || request.humanEvents === undefined || request.aiEvents === undefined) {
    try {
      hydratedEvents = vNextLoadEvidenceEventsForRun_(String(request.bookId || request.book_id || ''), {
        spreadsheet: request.spreadsheet,
        asOf: asOf,
        requestId: request.requestId || request.request_id || '',
        allowedDelayedAiRequestIds: allowedDelayedAiRequestIds,
        allowedEvidenceIds: exactRollbackEvidenceIds
      });
    } catch (evidenceError) {
      if (trustedRollbackSource) throw evidenceError;
      vNextLog_('Evidence hydration failed; explicit request events remain available', evidenceError);
    }
  }
  var commitmentEvents = request.commitmentEvents !== undefined ? request.commitmentEvents : hydratedEvents.commitment;
  var objectiveEvents = request.objectiveEvents !== undefined
    ? request.objectiveEvents
    : (request.changeEvents !== undefined ? request.changeEvents : hydratedEvents.objective);
  var humanEvents = request.humanEvents !== undefined ? request.humanEvents : hydratedEvents.human;
  var aiEvents = request.aiEvents !== undefined ? request.aiEvents : hydratedEvents.ai;
  var evidenceResponseCounts = request.evidenceResponseCounts || hydratedEvents.counts || { change: 0, noChange: 0, unknown: 0 };
  var records = request.actualRecords;
  if (!records) {
    records = vNextFetchActualRecordsBridge_(request.clientName || request.client_name || (context && context.clientName), {
      fiscalYear: fiscalYear,
      asOf: asOf,
      cutoff: cutoff,
      sourceSpreadsheetId: request.sourceSpreadsheetId,
      fetcher: request.actualFetcher,
      legacyFetcher: request.legacyFetcher,
      legacySourceGuaranteesActualDate: request.legacySourceGuaranteesActualDate
    });
  }
  var actualRecords = vNextValidateActualRecords_(records, asOf);
  var previousRunPinned = request.serverRunIdentityAuthorized === true &&
    request.authorizedPreviousRunIdPinned === true;
  var previousRunSummary = trustedRollbackSource || (previousRunPinned && !request.previousRunId
    ? null
    : vNextLoadPreviousRunSummary_(bookId, request.previousRunId || request.previous_run_id, request.spreadsheet));
  var modelReleaseId = String(trustedRollbackSource && trustedRollbackSource.modelReleaseId || request.modelReleaseId || request.model_release_id || (context && context.version && context.version.modelReleaseId) || '');
  var bookSchemaVersion = String(trustedRollbackSource && trustedRollbackSource.versions && trustedRollbackSource.versions.bookSchema || request.schemaVersion || request.schema_version || (context && context.version && context.version.schema) || VNEXT_CORE.SCHEMA_VERSION);
  var templateVersion = String(trustedRollbackSource && trustedRollbackSource.versions && trustedRollbackSource.versions.template || request.templateVersion || request.template_version || (context && context.version && context.version.template) || VNEXT_CORE.DEFAULT_TEMPLATE_VERSION);
  vNextAssertRuntimeReleaseRequest_(modelReleaseId, templateVersion, bookSchemaVersion, {
    requiresBoundRelease: request.persist !== false && typeof SpreadsheetApp !== 'undefined'
  });
  var releaseParameters = vNextLoadModelReleaseParameters_(modelReleaseId, request.spreadsheet, templateVersion);
  var effectiveParameters = Object.keys(request.parameters || {}).length ? request.parameters : releaseParameters;
  var rollbackReferencePrior = trustedRollbackSource && trustedRollbackSource.lenses &&
    trustedRollbackSource.lenses.changeReference && trustedRollbackSource.lenses.changeReference.referencePrior;
  var requestedReferencePrior = rollbackReferencePrior || request.referencePrior || effectiveParameters.referencePrior || {};
  if (!VNEXT_ENGINE.REFERENCE_PRIOR_PILOT_ENABLED &&
      request.persist !== false && typeof SpreadsheetApp !== 'undefined') {
    requestedReferencePrior = {
      mode: 'DISABLED', reason: 'PILOT_REQUIRES_VERSIONED_COHORT_SNAPSHOT', strength: 0
    };
  }
  var referencePrior = vNextNormalizeReferencePrior_(
    requestedReferencePrior,
    String(request.clientId || request.client_id || (context && context.clientId) || '')
  );
  var simulationCount = Math.max(
    VNEXT_ENGINE.MIN_SIMULATIONS,
    Math.min(VNEXT_ENGINE.MAX_SIMULATIONS, Number(
      trustedRollbackSource && trustedRollbackSource.simulationCount || request.simulationCount || effectiveParameters.simulationCount
    ) || VNEXT_ENGINE.DEFAULT_SIMULATIONS)
  );
  var rollbackEvidenceSummary = trustedRollbackSource && trustedRollbackSource.evidenceSummary || {};
  var missingResponseRate = vNextClamp_(
    trustedRollbackSource
      ? Number(rollbackEvidenceSummary.missingResponseRate || 0)
      : request.missingResponseRate !== undefined
      ? Number(request.missingResponseRate)
      : vNextMissingResponseRateFromContext_(context),
    0,
    1
  );
  var informationGapRate = vNextClamp_(Number(
    trustedRollbackSource ? rollbackEvidenceSummary.informationGapRate || 0 : request.informationGapRate || 0
  ), 0, 1);
  if (trustedRollbackSource && rollbackEvidenceSummary.responseCounts) {
    evidenceResponseCounts = rollbackEvidenceSummary.responseCounts;
  }
  var inputForSeed = {
    bookId: bookId,
    clientId: String(request.clientId || request.client_id || (context && context.clientId) || ''),
    fiscalYear: fiscalYear,
    asOf: vNextFormatDateOnly_(asOf),
    cutoff: vNextFormatDateOnly_(cutoff),
    actualRecords: actualRecords.map(vNextActualRecordForHash_).sort(vNextCanonicalSort_),
    commitmentEvents: (commitmentEvents || []).slice().sort(vNextCanonicalSort_),
    objectiveEvents: (objectiveEvents || []).slice().sort(vNextCanonicalSort_),
    humanEvents: (humanEvents || []).slice().sort(vNextCanonicalSort_),
    aiEvents: (aiEvents || []).slice().sort(vNextCanonicalSort_),
    referencePrior: referencePrior,
    modelReleaseId: modelReleaseId,
    parameters: effectiveParameters,
    missingResponseRate: missingResponseRate,
    informationGapRate: informationGapRate,
    aiUnavailable: request.aiUnavailable === true,
    aiUnavailableReason: request.aiUnavailable === true
      ? String(request.aiUnavailableReason || 'AI_RESEARCH_UNAVAILABLE').slice(0, 200)
      : '',
    evidenceResponseCounts: evidenceResponseCounts,
    simulationCount: Math.floor(simulationCount),
    effectivePolicy: {
      coreVersion: VNEXT_CORE.VERSION,
      schemaVersion: VNEXT_CORE.SCHEMA_VERSION,
      bookSchemaVersion: bookSchemaVersion,
      templateVersion: templateVersion,
      engineVersion: VNEXT_ENGINE.VERSION,
      minHistoryYears: VNEXT_ENGINE.MIN_HISTORY_YEARS,
      maxHistoryYears: VNEXT_ENGINE.MAX_HISTORY_YEARS,
      aiMaxAbsEffect: VNEXT_ENGINE.AI_MAX_ABS_EFFECT,
      missingResponseUncertainty: VNEXT_ENGINE.MISSING_RESPONSE_UNCERTAINTY,
      informationGapUncertainty: VNEXT_ENGINE.INFORMATION_GAP_UNCERTAINTY,
      knownSpotSuppressionRate: VNEXT_ENGINE.KNOWN_SPOT_SUPPRESSION_RATE,
      defaultMonthCommonShockSigma: VNEXT_ENGINE.DEFAULT_MONTH_COMMON_SHOCK_SIGMA,
      minMonthCommonShockSigma: VNEXT_ENGINE.MIN_MONTH_COMMON_SHOCK_SIGMA,
      maxMonthCommonShockSigma: VNEXT_ENGINE.MAX_MONTH_COMMON_SHOCK_SIGMA,
      referenceMaxStrength: VNEXT_ENGINE.REFERENCE_MAX_STRENGTH,
      referenceMinGrowth: VNEXT_ENGINE.REFERENCE_MIN_GROWTH,
      referenceMaxGrowth: VNEXT_ENGINE.REFERENCE_MAX_GROWTH,
      referenceMaxGrowthStd: VNEXT_ENGINE.REFERENCE_MAX_GROWTH_STD,
      referencePriorPilotEnabled: VNEXT_ENGINE.REFERENCE_PRIOR_PILOT_ENABLED,
      sourceColumns: VNEXT_ENGINE.SOURCE_COLUMNS,
      sourceHeaders: VNEXT_ENGINE.SOURCE_HEADERS
    },
    requestedSeed: trustedRollbackSource
      ? vNextNormalizeSeed_(trustedRollbackSource.seed)
      : (request.seed !== undefined && request.seed !== null ? vNextNormalizeSeed_(request.seed) : null)
  };
  var nonAiComparableInput = {};
  Object.keys(inputForSeed).forEach(function (key) {
    if (key === 'aiEvents' || key === 'aiUnavailable' || key === 'aiUnavailableReason' || key === 'requestedSeed') return;
    nonAiComparableInput[key] = inputForSeed[key];
  });
  var nonAiComparableHash = vNextSha256Hex_(nonAiComparableInput);
  if (trustedRollbackSource) {
    var sourceComparableHash = String(
      trustedRollbackSource.evidenceSummary && trustedRollbackSource.evidenceSummary.nonAiComparableHash || ''
    );
    if (!sourceComparableHash) {
      throw new Error('AI rollback source predates the comparable-input hash contract. Create a new draft run first.');
    }
    if (sourceComparableHash !== nonAiComparableHash) {
      throw new Error('Non-AI inputs changed after the source run; use a normal re-run instead of an AI-only comparison. source=' +
        sourceComparableHash.slice(0, 12) + ', current=' + nonAiComparableHash.slice(0, 12));
    }
  }
  var effectiveEvidenceIds = vNextEffectiveEvidenceIds_({
    commitment: commitmentEvents || [], objective: objectiveEvents || [],
    human: humanEvents || [], ai: aiEvents || []
  });
  var seedBasisHash = vNextSha256Hex_(inputForSeed);
  var seed = trustedRollbackSource
    ? vNextNormalizeSeed_(trustedRollbackSource.seed)
    : request.seed !== undefined && request.seed !== null
    ? vNextNormalizeSeed_(request.seed)
    : parseInt(seedBasisHash.slice(0, 8), 16) >>> 0;
  var inputForHash = {};
  Object.keys(inputForSeed).forEach(function (key) { inputForHash[key] = inputForSeed[key]; });
  inputForHash.seed = seed;
  var inputDataHash = vNextSha256Hex_(inputForHash);
  var clientId = String(request.clientId || request.client_id || (context && context.clientId) || '');
  var previousRunId = String(previousRunSummary && previousRunSummary.runId || request.previousRunId || '');
  var rollbackContext = trustedRollbackSource
    ? vNextPublicRollbackContext_(request.trustedRollbackContext, trustedRollbackSource)
    : null;
  var runIdentity = vNextFinalizeRunIdentity_(authorizedRunIdentity, {
    bookId: bookId,
    clientId: clientId,
    fiscalYear: fiscalYear,
    asOf: vNextFormatDateOnly_(asOf),
    cutoff: vNextFormatDateOnly_(cutoff),
    modelReleaseId: modelReleaseId,
    bookSchemaVersion: bookSchemaVersion,
    templateVersion: templateVersion,
    previousRunId: previousRunId,
    requestId: String(request.requestId || request.request_id || ''),
    internalJobType: String(request.internalJobType || 'FORECAST_REQUEST').toUpperCase(),
    rollback: rollbackContext
  });
  return {
    runId: runIdentity ? runIdentity.runId : String(vNextUuid_()),
    runIdentity: runIdentity,
    bookId: bookId,
    clientId: clientId,
    clientName: String(request.clientName || request.client_name || (context && context.clientName) || ''),
    fiscalYear: fiscalYear,
    asOf: asOf,
    cutoff: cutoff,
    seed: seed,
    inputDataHash: inputDataHash,
    modelReleaseId: modelReleaseId,
    bookSchemaVersion: bookSchemaVersion,
    templateVersion: templateVersion,
    actualRecords: actualRecords,
    commitmentEvents: commitmentEvents || [],
    objectiveEvents: objectiveEvents || [],
    humanEvents: humanEvents || [],
    aiEvents: aiEvents || [],
    referencePrior: referencePrior,
    parameters: effectiveParameters,
    missingResponseRate: missingResponseRate,
    informationGapRate: informationGapRate,
    aiUnavailable: request.aiUnavailable === true,
    aiUnavailableReason: request.aiUnavailable === true
      ? String(request.aiUnavailableReason || 'AI_RESEARCH_UNAVAILABLE').slice(0, 200)
      : '',
    evidenceResponseCounts: evidenceResponseCounts,
    effectiveEvidenceIds: effectiveEvidenceIds,
    nonAiComparableHash: nonAiComparableHash,
    simulationCount: Math.floor(simulationCount),
    persist: request.persist !== false && typeof SpreadsheetApp !== 'undefined',
    manageState: request.manageState !== false,
    initialState: String(request.initialState || '').toUpperCase(),
    spreadsheet: request.spreadsheet,
    clientSpreadsheetId: String(request.targetSpreadsheetId || request.clientSpreadsheetId || (context && context.clientSpreadsheetId) || ''),
    createdBy: String(request.createdBy || request.requestedBy || vNextActiveUserEmail_()).toLowerCase(),
    previousRunId: previousRunId,
    previousRunSummary: previousRunSummary,
    rollbackContext: rollbackContext
  };
}

function vNextNormalizeTrustedStringArray_(value, fieldName) {
  if (value === undefined || value === null || value === '') return [];
  if (!Array.isArray(value)) throw new Error((fieldName || 'trusted array') + ' must be an array.');
  if (value.length > 100) throw new Error((fieldName || 'trusted array') + ' is too large.');
  return value.map(function (item) { return String(item || '').trim(); })
    .filter(Boolean)
    .filter(function (item, index, all) { return all.indexOf(item) === index; })
    .sort();
}

function vNextResolveTrustedRollbackSource_(request, bookId, fiscalYear, asOf, cutoff) {
  var sourceRunId = String(request.trustedReuseSeedFromRunId || '');
  var rollback = request.trustedRollbackContext;
  if (!sourceRunId || !rollback || typeof rollback !== 'object' || Array.isArray(rollback)) {
    throw new Error('Trusted rollback requires source run ID and rollback context.');
  }
  if (String(rollback.sourceForecastRunId || '') !== sourceRunId) {
    throw new Error('Trusted rollback source lineage is inconsistent.');
  }
  var source = vNextFindForecastByRunId_(bookId, sourceRunId, { spreadsheet: request.spreadsheet });
  if (!source || String(source.status || '').toUpperCase() !== 'SUCCESS' || source.officialVintageId) {
    throw new Error('Trusted rollback source must be a successful draft run.');
  }
  if (String(source.bookId || '') !== String(bookId) || Number(source.fiscalYear) !== Number(fiscalYear)) {
    throw new Error('Trusted rollback source book/FY linkage is invalid.');
  }
  if (vNextFormatDateOnly_(source.asOf) !== vNextFormatDateOnly_(asOf) ||
      vNextFormatDateOnly_(source.cutoff) !== vNextFormatDateOnly_(cutoff)) {
    throw new Error('Trusted rollback must reuse the source information vintage.');
  }
  if (String(rollback.sourceInputDataHash || '') !== String(source.inputDataHash || '') ||
      String(rollback.sourceModelReleaseId || '') !== String(source.modelReleaseId || '')) {
    throw new Error('Trusted rollback source immutable metadata is inconsistent.');
  }
  var sourceComparableHash = String(source.evidenceSummary && source.evidenceSummary.nonAiComparableHash || '');
  if (!sourceComparableHash || String(rollback.nonAiComparableHash || '') !== sourceComparableHash) {
    throw new Error('Trusted rollback comparable-input hash is missing or inconsistent.');
  }
  if (source.versions && source.versions.engine && String(source.versions.engine) !== VNEXT_ENGINE.VERSION) {
    throw new Error('AI rollback cannot compare paths across engine versions. Re-run the draft with the current release first.');
  }
  if (!vNextIsFiniteNumber_(source.seed) || !vNextIsFiniteNumber_(source.simulationCount)) {
    throw new Error('Trusted rollback source seed/simulation count is invalid.');
  }
  var allowed = vNextNormalizeTrustedStringArray_(request.trustedAllowedDelayedAiRequestIds, 'trustedAllowedDelayedAiRequestIds');
  var rollbackRequestId = String(request.requestId || request.request_id || '');
  if (!rollbackRequestId || allowed.indexOf(rollbackRequestId) < 0) {
    throw new Error('Trusted rollback request ID is not in the delayed-AI allowlist.');
  }
  return source;
}

function vNextPublicRollbackContext_(rollback, source) {
  var value = rollback || {};
  return {
    operationId: String(value.operationId || ''),
    sourceForecastRunId: String(source && source.runId || value.sourceForecastRunId || ''),
    sourceInputDataHash: String(source && source.inputDataHash || value.sourceInputDataHash || ''),
    sourceModelReleaseId: String(source && source.modelReleaseId || value.sourceModelReleaseId || ''),
    nonAiComparableHash: String(source && source.evidenceSummary && source.evidenceSummary.nonAiComparableHash || ''),
    scope: String(value.scope || ''),
    targetEvidenceIds: vNextNormalizeTrustedStringArray_(value.targetEvidenceIds || [], 'targetEvidenceIds'),
    tombstoneEvidenceIds: vNextNormalizeTrustedStringArray_(value.tombstoneEvidenceIds || [], 'tombstoneEvidenceIds'),
    sameSeed: true,
    sourceSeed: Number(source && source.seed),
    sourceSimulationCount: Number(source && source.simulationCount)
  };
}

function vNextTrustedRollbackEvidenceIds_(source, rollback) {
  var summary = source && source.evidenceSummary || {};
  var groups = summary.effectiveEvidenceIds;
  if (!groups || typeof groups !== 'object' || Array.isArray(groups)) {
    throw new Error('AI rollback source predates the exact-evidence snapshot contract. Create a new draft run first.');
  }
  var ids = [];
  ['commitment', 'objective', 'human', 'ai'].forEach(function (key) {
    var values = vNextNormalizeTrustedStringArray_(groups[key] || [], 'effectiveEvidenceIds.' + key);
    ids = ids.concat(values);
  });
  var aiIds = vNextNormalizeTrustedStringArray_(groups.ai || [], 'effectiveEvidenceIds.ai');
  if (!aiIds.length) throw new Error('AI rollback source has no snapshotted AI evidence IDs.');
  ids = ids.concat(vNextNormalizeTrustedStringArray_(
    rollback && rollback.tombstoneEvidenceIds || [], 'tombstoneEvidenceIds'
  ));
  return ids.filter(function (id, index, all) { return all.indexOf(id) === index; }).sort();
}

function vNextEffectiveEvidenceIds_(groups) {
  var source = groups || {};
  var output = {};
  ['commitment', 'objective', 'human', 'ai'].forEach(function (key) {
    output[key] = (source[key] || []).map(function (event) {
      return String(event && (event.evidenceId || event.evidence_id) || '').trim();
    }).filter(Boolean).filter(function (id, index, all) {
      return all.indexOf(id) === index;
    }).sort();
  });
  return output;
}

function vNextLoadEvidenceEventsForRun_(bookId, options) {
  if (!bookId) return { commitment: [], objective: [], human: [], ai: [], counts: { change: 0, noChange: 0, unknown: 0 } };
  var opt = options || {};
  var asOf = vNextParseDate_(opt.asOf || new Date(), 'as_of');
  var allowedEvidenceIds = Array.isArray(opt.allowedEvidenceIds) && opt.allowedEvidenceIds.length
    ? new Set(opt.allowedEvidenceIds.map(String))
    : null;
  var rows = vNextReadRecords_('EVIDENCE_EVENT', { spreadsheet: opt.spreadsheet }).filter(function (row) {
    if (String(row.book_id || '') !== bookId) return false;
    if (allowedEvidenceIds && !allowedEvidenceIds.has(String(row.evidence_id || ''))) return false;
    return vNextEvidenceEffectiveForRun_(row, {
      asOf: asOf,
      requestId: opt.requestId || opt.request_id || '',
      allowedDelayedAiRequestIds: opt.allowedDelayedAiRequestIds || opt.allowed_delayed_ai_request_ids || []
    });
  });
  var superseded = {};
  rows.forEach(function (row) {
    var oldId = String(row.supersedes_evidence_id || '');
    if (oldId) superseded[oldId] = true;
  });
  var latestByActor = {};
  rows.forEach(function (row) {
    var evidenceType = String(row.evidence_type || '').toUpperCase();
    // AI/reference records describe evidence, not a team member's response.
    if (evidenceType.indexOf('AI') >= 0 || evidenceType.indexOf('OBJECTIVE') >= 0 || evidenceType.indexOf('REFERENCE') >= 0) return;
    latestByActor[String(row.actor_email || '')] = row;
  });
  var counts = { change: 0, noChange: 0, unknown: 0 };
  Object.keys(latestByActor).forEach(function (actor) {
    var type = String(latestByActor[actor].response_type || '').toUpperCase();
    if (type === 'CHANGE') counts.change++;
    else if (type === 'NO_CHANGE') counts.noChange++;
    else if (type === 'UNKNOWN') counts.unknown++;
  });
  var output = { commitment: [], objective: [], human: [], ai: [], counts: counts };
  rows.forEach(function (row) {
    if (superseded[String(row.evidence_id || '')]) return;
    if (String(row.response_type || '').toUpperCase() !== 'CHANGE') return;
    var event = {
      evidenceId: String(row.evidence_id || ''),
      target: String(row.target || ''),
      direction: String(row.direction || ''),
      amountLow: vNextNumberOrUndefined_(row.amount_low),
      amountMid: vNextNumberOrUndefined_(row.amount_mid),
      amountHigh: vNextNumberOrUndefined_(row.amount_high),
      amountBand: String(row.amount_band || ''),
      confidence: String(row.confidence_class || ''),
      startMonth: String(row.target_start_month || ''),
      endMonth: String(row.target_end_month || ''),
      evidenceText: String(row.evidence_text || ''),
      sourceUrl: String(row.source_url || ''),
      sourceDate: String(row.source_date || ''),
      expiresAt: String(row.expires_at || ''),
      evidenceType: String(row.evidence_type || ''),
      evidenceQuality: String(row.evidence_quality || ''),
      aiModel: String(row.ai_model || ''),
      promptVersion: String(row.prompt_version || ''),
      aiSchemaVersion: String(row.ai_schema_version || ''),
      ruleVersion: String(row.rule_version || ''),
      appliedAmount: vNextNumberOrUndefined_(row.applied_amount),
      capApplied: Number(row.cap_applied || 0) === 1
    };
    var type = String(row.evidence_type || 'HUMAN_CHANGE').toUpperCase();
    if (type.indexOf('COMMIT') >= 0 || type.indexOf('CONTRACT') >= 0) output.commitment.push(event);
    else if (type.indexOf('AI') >= 0) output.ai.push(event);
    else if (type.indexOf('OBJECTIVE') >= 0 || type.indexOf('REFERENCE') >= 0) output.objective.push(event);
    else output.human.push(event);
  });
  return output;
}

/**
 * Information-vintage guard shared by the loader and pure tests.
 *
 * Human evidence must have existed by the run as_of. AI research may finish
 * after midnight, but only the result linked to this exact request can cross
 * that created_at boundary, and only when its source was available by as_of.
 * Expiry is inclusive: evidence expiring on as_of is still valid that day.
 */
function vNextEvidenceEffectiveForRun_(row, options) {
  var item = row || {};
  var opt = options || {};
  if (String(item.status || 'ACTIVE').toUpperCase() === 'VOID') return false;
  var asOf = vNextParseDate_(opt.asOf || new Date(), 'as_of');
  var asOfEnd = vNextAsOfEnd_(asOf);
  var asOfText = vNextFormatDateOnly_(asOf);
  var evidenceType = String(item.evidence_type || '').toUpperCase();
  var isAi = evidenceType.indexOf('AI') >= 0;
  var metadata = {};
  if (isAi && item.evidence_text) {
    try { metadata = vNextParseJsonValue_(item.evidence_text, {}) || {}; }
    catch (ignoreMetadata) { metadata = {}; }
  }
  var createdAfterAsOf = false;
  if (item.created_at) {
    try { createdAfterAsOf = vNextParseDate_(item.created_at, 'created_at') > asOfEnd; }
    catch (invalidCreatedAt) { return false; }
  }
  if (createdAfterAsOf) {
    var requestId = String(opt.requestId || opt.request_id || '');
    var parentRequestId = String(metadata.parentRequestId || metadata.parent_request_id || '');
    var effectiveAsOf = String(metadata.effectiveAsOf || metadata.effective_as_of || '');
    var allowedDelayed = Array.isArray(opt.allowedDelayedAiRequestIds || opt.allowed_delayed_ai_request_ids)
      ? (opt.allowedDelayedAiRequestIds || opt.allowed_delayed_ai_request_ids).map(String)
      : [];
    var linkedToRequest = parentRequestId === requestId || allowedDelayed.indexOf(parentRequestId) >= 0;
    if (!isAi || !requestId || !linkedToRequest || effectiveAsOf !== asOfText) return false;
  }
  if (!isAi) return true;

  // AI evidence without a dated source cannot be proven available at the
  // requested vintage and therefore fails closed.
  if (!item.source_date) return false;
  try {
    if (vNextParseDate_(item.source_date, 'source_date') > asOfEnd) return false;
  } catch (invalidSourceDate) {
    return false;
  }
  if (item.expires_at) {
    try {
      var expiryEnd = vNextAsOfEnd_(vNextParseDate_(item.expires_at, 'expires_at'));
      var asOfStart = vNextParseDate_(asOfText, 'as_of');
      if (expiryEnd < asOfStart) return false;
    } catch (invalidExpiry) {
      return false;
    }
  }
  var metadataAsOf = String(metadata.effectiveAsOf || metadata.effective_as_of || '');
  if (metadataAsOf) {
    try { if (vNextParseDate_(metadataAsOf, 'effective_as_of') > asOfEnd) return false; }
    catch (invalidEffectiveAsOf) { return false; }
  }
  return true;
}

function vNextAssertRuntimeReleaseRequest_(modelReleaseId, templateVersion, bookSchemaVersion, options) {
  var opt = options || {};
  var template = String(templateVersion || '').trim();
  if (!template) throw new Error('templateVersion is required for MODEL_RELEASE binding.');
  if (opt.requiresBoundRelease && !String(modelReleaseId || '').trim()) {
    throw new Error('A bound modelReleaseId is required before forecast simulation.');
  }
  if (opt.requiresBoundRelease && String(bookSchemaVersion || '') !== String(VNEXT_CORE.SCHEMA_VERSION)) {
    throw new Error('Book schema does not match the deployed Core schema.');
  }
  return true;
}

function vNextAssertModelReleaseRuntimeBinding_(release, expectedModelReleaseId, expectedTemplateVersion) {
  var row = release || {};
  var expectedTemplate = String(expectedTemplateVersion || '').trim();
  if (!expectedTemplate) throw new Error('expectedTemplateVersion is required.');
  if (String(row.model_release_id || '') !== String(expectedModelReleaseId || '') ||
      String(row.status || '').toUpperCase() !== 'ACTIVE' ||
      String(row.model_version || '') !== String(VNEXT_ENGINE.VERSION) ||
      String(row.schema_version || '') !== String(VNEXT_CORE.SCHEMA_VERSION) ||
      String(row.template_version || '') !== expectedTemplate) {
    throw new Error('MODEL_RELEASE does not exactly match model/template/Engine/Core runtime binding.');
  }
  return true;
}

function vNextLoadModelReleaseParameters_(modelReleaseId, spreadsheet, expectedTemplateVersion) {
  var expectedTemplate = String(expectedTemplateVersion || '').trim();
  if (!expectedTemplate) throw new Error('expectedTemplateVersion is required for MODEL_RELEASE loading.');
  if (!modelReleaseId || typeof SpreadsheetApp === 'undefined') return {};
  try {
    var rows = vNextReadRecords_('MODEL_RELEASE', { spreadsheet: spreadsheet }).filter(function (row) {
      return String(row.model_release_id || '') === String(modelReleaseId);
    });
    if (!rows.length) throw new Error('Bound MODEL_RELEASE was not found: ' + modelReleaseId);
    var release = rows[rows.length - 1];
    vNextAssertModelReleaseRuntimeBinding_(release, modelReleaseId, expectedTemplate);
    var parameters = vNextValidateModelReleaseParameters_(
      vNextParseJsonValue_(release.parameters_json, {})
    );
    if (vNextCanonicalJson_(parameters) !== String(release.parameters_json || '')) {
      throw new Error('MODEL_RELEASE parameters are not canonical.');
    }
    var candidateHash = vNextSha256Hex_(vNextCanonicalJson_({
      modelVersion: String(release.model_version || ''),
      schemaVersion: String(release.schema_version || ''),
      templateVersion: String(release.template_version || ''),
      parameters: parameters
    }));
    var backtest = vNextParseJsonValue_(release.backtest_json, {});
    var canary = vNextParseJsonValue_(release.canary_json, {});
    if (String(backtest.candidateHash || '') !== candidateHash ||
        String(canary.candidateHash || '') !== candidateHash) {
      throw new Error('MODEL_RELEASE verification artifacts do not match its candidate hash.');
    }
    return parameters;
  } catch (error) {
    vNextLog_('Model release validation failed', error);
    throw error;
  }
}

function vNextValidateModelReleaseParameters_(input) {
  var source = input && typeof input === 'object' && !Array.isArray(input) ? input : {};
  var allowed = { simulationCount: true, referencePrior: true };
  var unknown = Object.keys(source).filter(function (key) { return !allowed[key]; });
  if (unknown.length) throw new Error('Unsupported MODEL_RELEASE parameters: ' + unknown.join(', '));
  var output = {};
  if (Object.prototype.hasOwnProperty.call(source, 'simulationCount')) {
    var count = Number(source.simulationCount);
    if (!isFinite(count) || Math.floor(count) !== count ||
        count < VNEXT_ENGINE.MIN_SIMULATIONS || count > VNEXT_ENGINE.MAX_SIMULATIONS) {
      throw new Error('MODEL_RELEASE simulationCount is outside the deployed Engine range.');
    }
    output.simulationCount = count;
  }
  if (Object.prototype.hasOwnProperty.call(source, 'referencePrior')) {
    var prior = source.referencePrior;
    if (!prior || typeof prior !== 'object' || Array.isArray(prior)) {
      throw new Error('MODEL_RELEASE referencePrior must be an object.');
    }
    var priorAllowed = { mode: true, reason: true, growthMean: true, growthStd: true, strength: true };
    var priorUnknown = Object.keys(prior).filter(function (key) { return !priorAllowed[key]; });
    if (priorUnknown.length) throw new Error('Unsupported MODEL_RELEASE referencePrior keys: ' + priorUnknown.join(', '));
    var normalized = vNextNormalizeReferencePrior_(prior, '');
    if (!VNEXT_ENGINE.REFERENCE_PRIOR_PILOT_ENABLED && String(prior.mode || '').toUpperCase() !== 'DISABLED') {
      throw new Error('MODEL_RELEASE referencePrior is disabled for the initial pilot until a cohort snapshot contract is installed.');
    }
    if (String(prior.mode || '').toUpperCase() === 'DISABLED') {
      if (normalized.mode !== 'DISABLED' || Number(normalized.strength || 0) !== 0) {
        throw new Error('Disabled referencePrior is invalid.');
      }
    } else if (normalized.mode !== 'GROWTH') {
      throw new Error('Only a validated growth referencePrior can be global.');
    }
    output.referencePrior = Object.assign({}, prior);
  }
  return output;
}

function vNextLoadPreviousRunSummary_(bookId, previousRunId, spreadsheet) {
  if (!bookId || typeof SpreadsheetApp === 'undefined') return null;
  try {
    var rows = vNextReadRecords_('FORECAST_RUN', { spreadsheet: spreadsheet }).filter(function (row) {
      return String(row.book_id || '') === String(bookId) &&
        ['SUCCESS', 'OFFICIAL_LOCKED'].indexOf(String(row.status || '').toUpperCase()) >= 0;
    });
    var requestedId = String(previousRunId || '');
    var row = requestedId
      ? rows.filter(function (item) { return String(item.run_id || '') === requestedId; }).slice(-1)[0]
      : rows.slice(-1)[0];
    if (!row) return null;
    return {
      runId: String(row.run_id || ''),
      asOf: String(row.as_of || ''),
      historyBaseline: Number(row.history_baseline || 0),
      commitmentDelta: Number(row.commitment_delta || 0),
      referenceDelta: Number(row.reference_delta || 0),
      humanDelta: Number(row.human_delta || 0),
      aiDelta: Number(row.ai_delta || 0),
      systemRecommended: Number(row.system_recommended || row.p50 || 0)
    };
  } catch (error) {
    vNextLog_('Previous run summary was not available', error);
    return null;
  }
}

function vNextAsOfEnd_(value) {
  var date = vNextParseDate_(value, 'as_of');
  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), 23, 59, 59, 999);
}

function vNextMissingResponseRateFromContext_(context) {
  if (!context || !context.inputStatus) return 0;
  var total = Number(context.inputStatus.totalCount || 0);
  var answered = Number(context.inputStatus.answeredCount || 0);
  return total > 0 ? Math.max(0, total - answered) / total : 0;
}

function vNextNumberOrUndefined_(value) {
  return vNextIsFiniteNumber_(value) ? Number(value) : undefined;
}

/** Strict bridge: a legacy fetcher is usable only with an explicit actual-date guarantee. */
function vNextFetchActualRecordsBridge_(client, options) {
  var opt = options || {};
  var records;
  if (typeof opt.fetcher === 'function') {
    records = opt.fetcher(client, opt);
  } else if (typeof opt.legacyFetcher === 'function') {
    if (opt.legacySourceGuaranteesActualDate !== true) {
      throw new Error('Legacy fetcher rejected: it may fall back to planned dates. Supply an actual-only adapter.');
    }
    records = opt.legacyFetcher(client, opt).map(function (row) {
      var copy = {};
      Object.keys(row || {}).forEach(function (key) { copy[key] = row[key]; });
      copy.actualDate = copy.actualDate || copy.monthStart;
      copy.dateSource = 'ACTUAL_ATTESTED_LEGACY';
      copy.isConfirmed = true;
      return copy;
    });
  } else {
    records = vNextFetchActualRowsFromSpreadsheet_(client, opt);
  }
  return vNextValidateActualRecords_(records || [], opt.asOf || new Date());
}

/** Reads BE actual date only. BD/planned date is intentionally never read. */
function vNextFetchActualRowsFromSpreadsheet_(client, options) {
  try {
    var opt = options || {};
    if (typeof SpreadsheetApp === 'undefined') throw new Error('SpreadsheetApp is unavailable.');
    var sourceId = String(
      opt.sourceSpreadsheetId ||
      vNextGetProperty_(VNEXT_ENGINE.SOURCE_ID_PROPERTY, false) ||
      vNextGetProperty_('FORECAST_SOURCE_SPREADSHEET_ID', false) ||
      ''
    );
    if (!sourceId) throw new Error(VNEXT_ENGINE.SOURCE_ID_PROPERTY + ' or FORECAST_SOURCE_SPREADSHEET_ID is required.');
    var targetClient = String(client || '').trim();
    if (!targetClient) throw new Error('clientName is required.');
    var cutoff = opt.cutoff ? vNextParseDate_(opt.cutoff, 'cutoff') : vNextCutoffFromAsOf_(opt.asOf || new Date());
    var fiscalYear = Number(opt.fiscalYear);
    var historyStart = new Date(fiscalYear - VNEXT_ENGINE.MAX_HISTORY_YEARS, 3, 1);
    var columns = VNEXT_ENGINE.SOURCE_COLUMNS;
    var source = SpreadsheetApp.openById(sourceId);
    var sheets = source.getSheets().filter(function (sheet) {
      var name = sheet.getName();
      var match = name.match(/^\*(\d{4})_actual_value$/);
      if (!match) return false;
      var calendarYear = Number(match[1]);
      return calendarYear >= historyStart.getFullYear() && calendarYear <= cutoff.getFullYear();
    });
    if (!sheets.length) throw new Error('No ZAC actual-value source sheets were found.');
    var output = [];
    sheets.forEach(function (sheet) {
      if (sheet.getLastRow() < 2) return;
      var requiredColumns = [columns.client, columns.serviceCategory, columns.product, columns.actualDate, columns.amount];
      var header = [];
      requiredColumns.forEach(function (column) { header[column - 1] = sheet.getRange(1, column).getValue(); });
      vNextValidateSourceHeader_(header, columns, sheet.getName(), opt.expectedHeaders);
      var rowCount = sheet.getLastRow() - 1;
      var columnValues = {};
      requiredColumns.forEach(function (column) {
        columnValues[column] = sheet.getRange(2, column, rowCount, 1).getValues();
      });
      for (var index = 0; index < rowCount; index++) {
        var rowClient = String(columnValues[columns.client][index][0] || '').trim();
        if (!vNextSameClient_(rowClient, targetClient)) continue;
        var actualRaw = columnValues[columns.actualDate][index][0];
        if (!actualRaw) continue;
        var actualDate;
        try { actualDate = vNextParseDate_(actualRaw, 'actualDate'); } catch (error) { continue; }
        if (actualDate < historyStart || actualDate > cutoff) continue;
        var amount = Number(columnValues[columns.amount][index][0]);
        if (!isFinite(amount)) continue;
        output.push({
          client: rowClient,
          actualDate: actualDate,
          dateSource: 'ACTUAL',
          isConfirmed: true,
          amount: amount,
          serviceType: vNextClassifyServiceType_(columnValues[columns.serviceCategory][index][0]),
          product: String(columnValues[columns.product][index][0] || '').trim(),
          sourceSheet: sheet.getName(),
          sourceRow: index + 2
        });
      }
    });
    return output;
  } catch (error) {
    vNextLog_('vNextFetchActualRowsFromSpreadsheet_ failed', error);
    throw error;
  }
}

function vNextValidateSourceHeader_(header, columns, sheetName, expectedHeaders) {
  var actualHeader = String(header[columns.actualDate - 1] || '').trim();
  var amountHeader = String(header[columns.amount - 1] || '').trim();
  var clientHeader = String(header[columns.client - 1] || '').trim();
  if (!actualHeader || !amountHeader || !clientHeader) throw new Error(sheetName + ' source schema has blank required headers.');
  if (/予定|plan/i.test(actualHeader)) throw new Error(sheetName + ' actual-date column points to a planned-date header: ' + actualHeader);
  var contract = expectedHeaders || VNEXT_ENGINE.SOURCE_HEADERS;
  if (contract) {
    Object.keys(contract).forEach(function (key) {
      var col = columns[key];
      if (!col) return;
      if (String(header[col - 1] || '').trim() !== String(contract[key])) {
        throw new Error(sheetName + ' source header mismatch: ' + key + '; expected=' + contract[key] + '; actual=' + String(header[col - 1] || '').trim());
      }
    });
  }
}

function vNextValidateActualRecords_(records, asOf) {
  var cutoff = vNextCutoffFromAsOf_(asOf);
  return (records || []).map(function (record, index) {
    var actualRaw = record && (record.actualDate || record.actual_date);
    var plannedRaw = record && (record.plannedDate || record.planned_date);
    if (!actualRaw) {
      if (plannedRaw) throw new Error('Record ' + index + ' has only plannedDate; fallback is prohibited.');
      throw new Error('Record ' + index + ' is missing actualDate.');
    }
    var source = String(record.dateSource || record.date_source || 'ACTUAL').toUpperCase();
    if (source !== 'ACTUAL' && source !== 'ACTUAL_ATTESTED_LEGACY') {
      throw new Error('Record ' + index + ' dateSource is not actual: ' + source);
    }
    if (record.isConfirmed === false || record.confirmed === false) {
      throw new Error('Record ' + index + ' is not confirmed.');
    }
    var actualDate = vNextParseDate_(actualRaw, 'actualDate');
    if (actualDate > cutoff) {
      throw new Error('Record ' + index + ' is after cutoff ' + vNextFormatDateOnly_(cutoff) + '.');
    }
    var amount = Number(record.amount);
    if (!isFinite(amount)) throw new Error('Record ' + index + ' amount is invalid.');
    return {
      client: String(record.client || ''),
      actualDate: actualDate,
      dateSource: source,
      isConfirmed: true,
      amount: amount,
      serviceType: String(record.serviceType || record.service_type || 'OTHER').toUpperCase(),
      product: String(record.product || ''),
      sourceSheet: String(record.sourceSheet || ''),
      sourceRow: record.sourceRow || ''
    };
  });
}

function vNextSimulateForecast_(request) {
  var history = vNextBuildContinuityPrior_(request.actualRecords, request.fiscalYear, request.cutoff);
  var rng = vNextCreatePrng_(request.seed);
  var gaussian = vNextCreateGaussianSampler_(rng);
  var n = request.simulationCount;
  var stages = [[], [], [], [], [], []];
  var finalMonths = [];
  var stageMonths = [];
  var aiCapRates = [];
  var aiCapAppliedCount = 0;
  var commitmentDiagnostics = vNextInitializeCommitmentDiagnostics_(
    request.commitmentEvents,
    history.baseAnnualBaseline,
    request.fiscalYear
  );
  var uncertaintyMultiplier = 1 +
    request.missingResponseRate * VNEXT_ENGINE.MISSING_RESPONSE_UNCERTAINTY +
    request.informationGapRate * VNEXT_ENGINE.INFORMATION_GAP_UNCERTAINTY;
  var spotSuppression = vNextCommitmentSpotSuppression_(request.commitmentEvents, request.fiscalYear);
  for (var i = 0; i < n; i++) {
    var annualShockZ = gaussian();
    var baseAnnual = history.baseAnnualBaseline > 0
      ? history.baseAnnualBaseline * Math.exp(annualShockZ * history.logSigma * uncertaintyMultiplier)
      : 0;
    var monthShock = vNextSampleMonthlyCommonShock_(history.seasonalShares, history.monthlyCommonShockSigma, gaussian);
    var pathShares = monthShock.shares;
    var baseMonths = pathShares.map(function (share) { return Math.max(0, baseAnnual * share); });
    var unknownSpot = vNextSampleUnknownSpot_(
      history.unknownSpotModel,
      spotSuppression,
      rng,
      gaussian,
      uncertaintyMultiplier,
      monthShock.multipliers
    );
    var continuity = baseMonths.map(function (value, monthIndex) { return value + unknownSpot[monthIndex]; });
    var continuityAnnual = vNextSum_(continuity);
    var commitment = vNextSampleEventsLayer_(request.commitmentEvents, continuityAnnual, pathShares, request.fiscalYear, rng, gaussian, {
      recognitionDelay: true,
      diagnostics: commitmentDiagnostics
    });
    var afterCommitment = vNextApplyLayerNonNegative_(continuity, commitment);
    var reference = vNextSampleReferenceLayer_(request.referencePrior, afterCommitment, pathShares, gaussian);
    var afterReference = vNextApplyLayerNonNegative_(afterCommitment, reference);
    var objectiveEvents = vNextSampleEventsLayer_(request.objectiveEvents, vNextSum_(afterReference), pathShares, request.fiscalYear, rng, gaussian);
    var objective = vNextApplyLayerNonNegative_(afterReference, objectiveEvents);
    var humanLayer = vNextSampleEventsLayer_(request.humanEvents, vNextSum_(objective), pathShares, request.fiscalYear, rng, gaussian);
    var afterHuman = vNextApplyLayerNonNegative_(objective, humanLayer);
    var aiRaw = vNextSampleEventsLayer_(request.aiEvents, vNextSum_(afterHuman), pathShares, request.fiscalYear, rng, gaussian);
    var aiCap = vNextSum_(afterHuman) * VNEXT_ENGINE.AI_MAX_ABS_EFFECT;
    var aiNet = vNextSum_(aiRaw);
    var scale = Math.abs(aiNet) > aiCap && Math.abs(aiNet) > 1e-9 ? aiCap / Math.abs(aiNet) : 1;
    if (scale < 0.999999) aiCapAppliedCount++;
    var aiLayer = aiRaw.map(function (value) { return value * scale; });
    var final = vNextApplyLayerNonNegative_(afterHuman, aiLayer);
    var pathStages = [continuity, afterCommitment, afterReference, objective, afterHuman, final];
    for (var s = 0; s < pathStages.length; s++) stages[s].push(vNextSum_(pathStages[s]));
    finalMonths.push(final);
    stageMonths.push(pathStages);
    aiCapRates.push(vNextSum_(afterHuman) > 0 ? Math.abs(vNextSum_(final) - vNextSum_(afterHuman)) / vNextSum_(afterHuman) : 0);
  }
  var coherent = vNextBuildCoherentPathSummary_(finalMonths, request.fiscalYear);
  var p50PathIndex = coherent.pathIndexes.p50;
  var centralStages = stageMonths[p50PathIndex].map(vNextSum_);
  var layers = {
    historyBaseline: centralStages[0],
    commitmentDelta: centralStages[1] - centralStages[0],
    peerReferenceDelta: centralStages[2] - centralStages[1],
    objectiveEventDelta: centralStages[3] - centralStages[2],
    // The persisted reference_delta remains the backwards-compatible combined
    // objective-information layer; lens_json keeps its two meanings separate.
    referenceDelta: centralStages[3] - centralStages[1],
    objectiveForecast: centralStages[3],
    humanDelta: centralStages[4] - centralStages[3],
    aiDelta: centralStages[5] - centralStages[4],
    systemRecommended: centralStages[5]
  };
  var aiOffPaths = stageMonths.map(function (path) { return path[4]; });
  var aiOff = vNextBuildCoherentPathSummary_(aiOffPaths, request.fiscalYear);
  var aiCounterfactual = { annual: aiOff.annual, quarters: aiOff.quarters, months: aiOff.months };
  var publicAiEvidence = vNextBuildPublicAiEvidence_(request.aiEvents);
  var publicDrivers = vNextBuildPublicDrivers_(layers);
  var nextInformation = vNextRankNextInformation_(request, history, layers, coherent.annual);
  var changeReasons = vNextBuildChangeReasons_(request.previousRunSummary, layers);
  var evidenceReadiness = vNextBuildPublicEvidenceReadiness_(request, history, coherent.annual);
  return {
    runId: request.runId,
    bookId: request.bookId,
    clientId: request.clientId,
    clientName: request.clientName,
    fiscalYear: request.fiscalYear,
    asOf: vNextFormatDateOnly_(request.asOf),
    cutoff: vNextFormatDateOnly_(request.cutoff),
    seed: request.seed,
    inputDataHash: request.inputDataHash,
    modelReleaseId: request.modelReleaseId,
    previousRunId: request.previousRunId || '',
    versions: {
      core: VNEXT_CORE.VERSION,
      engine: VNEXT_ENGINE.VERSION,
      schema: VNEXT_CORE.SCHEMA_VERSION,
      bookSchema: request.bookSchemaVersion,
      template: request.templateVersion,
      modelReleaseId: request.modelReleaseId
    },
    status: 'SUCCESS',
    officialVintageId: '',
    historyYears: history.fiscalYears,
    simulationCount: n,
    layers: layers,
    annual: coherent.annual,
    quarters: coherent.quarters,
    months: coherent.months,
    lenses: {
      runIdentity: request.runIdentity || null,
      publicDrivers: publicDrivers,
      nextInformation: nextInformation,
      changeReasons: changeReasons,
      evidenceReadiness: evidenceReadiness,
      versions: {
        core: VNEXT_CORE.VERSION,
        engine: VNEXT_ENGINE.VERSION,
        schema: VNEXT_CORE.SCHEMA_VERSION,
        bookSchema: request.bookSchemaVersion,
        template: request.templateVersion,
        modelReleaseId: request.modelReleaseId
      },
      continuity: history,
      commitment: {
        eventCount: request.commitmentEvents.length,
        unknownSpotSuppressionByMonth: spotSuppression,
        delayByEvent: vNextFinalizeCommitmentDiagnostics_(commitmentDiagnostics)
      },
      simulationDesign: {
        seed: request.seed,
        annualCommonShock: {
          distribution: 'LOGNORMAL',
          logSigma: history.logSigma,
          uncertaintyMultiplier: uncertaintyMultiplier
        },
        monthlyCommonShock: {
          distribution: 'LOGNORMAL_NORMALIZED_TO_FY',
          logSigma: history.monthlyCommonShockSigma,
          sharedAcrossLayers: true
        },
        commitmentRecognitionDelay: {
          sampledPerEventPerPath: true,
          outsideFiscalYearPolicy: 'EXCLUDE_FROM_TARGET_FY'
        }
      },
      degradation: request.aiUnavailable === true ? {
        aiUnavailable: true,
        reason: request.aiUnavailableReason || 'AI_RESEARCH_UNAVAILABLE',
        policy: 'AI_ZERO_AND_WIDER_INTERVAL'
      } : null,
      rollback: request.rollbackContext || null,
      changeReference: {
        objectiveEventCount: request.objectiveEvents.length,
        humanEventCount: request.humanEvents.length,
        aiEventCount: request.aiEvents.length,
        referencePrior: request.referencePrior,
        peerReferenceDelta: layers.peerReferenceDelta,
        objectiveEventDelta: layers.objectiveEventDelta,
        combinedObjectiveInformationDelta: layers.referenceDelta,
        maxObservedAiEffectRate: Math.max.apply(null, aiCapRates),
        aiCapApplied: aiCapAppliedCount > 0,
        aiCapAppliedPathCount: aiCapAppliedCount,
        aiCapAppliedPathRate: n ? aiCapAppliedCount / n : 0,
        aiCounterfactual: aiCounterfactual
      }
    },
    evidenceSummary: {
      commitment: request.commitmentEvents.length,
      objective: request.objectiveEvents.length,
      human: request.humanEvents.length,
      ai: request.aiEvents.length,
      topAiEvidence: publicAiEvidence,
      missingResponseRate: request.missingResponseRate,
      informationGapRate: request.informationGapRate,
      aiUnavailable: request.aiUnavailable === true,
      aiUnavailableReason: request.aiUnavailableReason || '',
      effectiveEvidenceIds: request.effectiveEvidenceIds || {
        commitment: [], objective: [], human: [], ai: []
      },
      nonAiComparableHash: request.nonAiComparableHash || '',
      responseCounts: request.evidenceResponseCounts,
      noChange: Number(request.evidenceResponseCounts && request.evidenceResponseCounts.noChange || 0),
      unknown: Number(request.evidenceResponseCounts && request.evidenceResponseCounts.unknown || 0),
      unknownSpotExpectedAnnual: history.unknownSpotModel.expectedAnnual,
      unknownSpotExpectedOccurrences: history.unknownSpotModel.expectedOccurrences,
      unknownSpotMeanOccurrenceRate: history.unknownSpotModel.meanOccurrenceRate,
      readiness: evidenceReadiness
    }
  };
}

function vNextBuildPublicDrivers_(layers) {
  var values = [
    { label: '過去実績から見た継続売上力', amount: Number(layers.historyBaseline || 0), always: true },
    { label: '確認できた契約・案件', amount: Number(layers.commitmentDelta || 0) },
    { label: '比較可能な参照クラス', amount: Number(layers.peerReferenceDelta || 0) },
    { label: '確認できた客観的な変化', amount: Number(layers.objectiveEventDelta || 0) },
    { label: '現場から共有された変化', amount: Number(layers.humanDelta || 0) },
    { label: 'AI調査で確認した外部変化', amount: Number(layers.aiDelta || 0) }
  ];
  return values.filter(function (item) { return item.always || Math.abs(item.amount) >= 1; })
    .sort(function (a, b) { return Math.abs(b.amount) - Math.abs(a.amount); })
    .slice(0, 3)
    .map(function (item) {
      return item.label + ' ' + vNextEngineMoneyText_(item.amount);
    });
}

/**
 * Describes how complete the evidence is without claiming a probability of being correct.
 * This is intentionally separate from employee confidence classes and model calibration.
 */
function vNextBuildPublicEvidenceReadiness_(request, history, annual) {
  var historyYears = history && Array.isArray(history.fiscalYears) ? history.fiscalYears.slice() : [];
  var missingRate = vNextClamp_(Number(request && request.missingResponseRate || 0), 0, 1);
  var gapRate = vNextClamp_(Number(request && request.informationGapRate || 0), 0, 1);
  var responseCounts = request && request.evidenceResponseCounts || {};
  var unknownCount = Math.max(0, Number(responseCounts.unknown || 0));
  var center = Math.abs(Number(annual && annual.p50 || 0));
  var intervalWidth = Math.max(0, Number(annual && annual.p90 || 0) - Number(annual && annual.p10 || 0));
  var intervalWidthRate = center > 0 ? intervalWidth / center : (intervalWidth > 0 ? 1 : 0);
  var issues = [];
  if (missingRate > 0) issues.push('登録メンバーに未回答があります');
  if (gapRate > 0 || unknownCount > 0) issues.push('「情報不足」の回答があります');
  if (historyYears.length === VNEXT_ENGINE.MIN_HISTORY_YEARS) issues.push('利用できる履歴が最低年数です');
  if (intervalWidthRate >= 1) issues.push('通常の振れ幅が中心見込みと同程度以上です');
  else if (intervalWidthRate >= 0.55) issues.push('通常の振れ幅が比較的大きい状態です');
  if (request && request.aiUnavailable === true) issues.push('今回の外部情報調査は未完了です');

  var level = 'READY';
  var label = '根拠が比較的そろっています';
  if (missingRate >= 0.34 || gapRate >= 0.34 || intervalWidthRate >= 1) {
    level = 'NEEDS_ATTENTION';
    label = '追加確認の効果が大きい状態です';
  } else if (issues.length) {
    level = 'REVIEW';
    label = '確認余地があります';
  }
  var summary = level === 'READY'
    ? historyYears.length + '年度の確定実績と現在の回答をもとにしています。'
    : issues.slice(0, 2).join('。') + '。';
  return {
    level: level,
    label: label,
    summary: summary,
    historyYearCount: historyYears.length,
    historyFiscalYears: historyYears,
    missingResponseRate: missingRate,
    informationGapRate: gapRate,
    unknownResponseCount: unknownCount,
    intervalWidth: Math.round(intervalWidth),
    intervalWidthRate: intervalWidthRate,
    issues: issues.slice(0, 3)
  };
}

/** Rank questions by estimated interval reduction × monetary impact ÷ burden. */
function vNextRankNextInformation_(request, history, layers, annual) {
  var width = Math.max(0, Number(annual.p90 || 0) - Number(annual.p10 || 0));
  var scale = Math.max(1, Math.abs(Number(layers.systemRecommended || 0)), Math.abs(Number(layers.historyBaseline || 0)));
  var candidates = [];
  function add(code, text, widthReduction, amountImpact, burden) {
    var reduction = Math.max(0, Number(widthReduction || 0));
    var impact = Math.max(0, Number(amountImpact || 0));
    var cost = Math.max(1, Number(burden || 1));
    if (!reduction || !impact) return;
    candidates.push({
      code: code,
      text: text,
      expectedWidthReduction: Math.round(reduction),
      amountImpact: Math.round(impact),
      confirmationBurden: cost,
      score: reduction * impact / cost
    });
  }
  var missingRate = vNextClamp_(Number(request.missingResponseRate || 0), 0, 1);
  if (missingRate > 0) {
    add('MISSING_MEMBER', '未回答メンバーが把握している変化の有無を確認する', width * missingRate, scale * missingRate, 2);
  }
  var unknownCount = Number(request.evidenceResponseCounts && request.evidenceResponseCounts.unknown || 0);
  var gapRate = vNextClamp_(Number(request.informationGapRate || 0), 0, 1);
  if (unknownCount > 0 || gapRate > 0) {
    add('UNKNOWN_DETAIL', '「情報不足」となった項目の契約時期・金額・確度を確認する', width * Math.max(gapRate, 0.15), scale * Math.max(gapRate, 0.10), 2);
  }
  if (!(request.commitmentEvents || []).length) {
    add('COMMITMENT', '契約更新・確定案件・認識月ずれの有無を確認する', width * 0.35, scale * 0.12, 2);
  }
  var spot = history && history.unknownSpotModel || {};
  if (Number(spot.expectedAnnual || 0) > 0) {
    add('UNKNOWN_SPOT', '過去に突発発生した単発売上が来年度も起きるか確認する', Math.min(width * 0.30, Number(spot.annualStd || spot.expectedAnnual || 0)), Number(spot.expectedAnnual || 0), 3);
  }
  if (!(request.humanEvents || []).length) {
    add('CLIENT_CHANGE', '過去実績に出ない顧客・製品・組織の変化を確認する', width * 0.20, scale * 0.08, 3);
  }
  if (!(request.aiEvents || []).length) {
    add('EXTERNAL_CHANGE', '情報締切時点の公開情報に重要な外部変化がないか確認する', width * 0.12, scale * 0.05, 4);
  }
  return candidates.sort(function (a, b) {
    if (b.score !== a.score) return b.score - a.score;
    return String(a.code).localeCompare(String(b.code));
  }).slice(0, 3);
}

function vNextBuildChangeReasons_(previous, currentLayers) {
  if (!previous || !previous.runId) return ['初回の予測作成です。'];
  var comparisons = [
    { label: '継続売上力', before: previous.historyBaseline, after: currentLayers.historyBaseline },
    { label: '契約・案件', before: previous.commitmentDelta, after: currentLayers.commitmentDelta },
    { label: '参照情報と客観変化', before: previous.referenceDelta, after: currentLayers.referenceDelta },
    { label: '現場情報', before: previous.humanDelta, after: currentLayers.humanDelta },
    { label: 'AI調査', before: previous.aiDelta, after: currentLayers.aiDelta }
  ].map(function (item) {
    item.delta = Number(item.after || 0) - Number(item.before || 0);
    return item;
  }).filter(function (item) { return Math.abs(item.delta) >= 1; })
    .sort(function (a, b) { return Math.abs(b.delta) - Math.abs(a.delta); });
  var reasons = comparisons.slice(0, 3).map(function (item) {
    return item.label + 'が前回から ' + vNextEngineMoneyText_(item.delta) + ' 変わりました。';
  });
  if (!reasons.length) reasons.push('主要な前提と中心見込みに大きな変更はありません。');
  return reasons;
}

function vNextEngineMoneyText_(value) {
  var number = Math.round(Number(value || 0));
  var sign = number > 0 ? '+' : (number < 0 ? '-' : '');
  return sign + '¥' + Math.abs(number).toLocaleString('ja-JP');
}

/**
 * FORECAST_RUN経由で社員へ返すAI根拠は、確認に必要な最小項目だけに限定する。
 * raw response、prompt本文、model内部情報はAdmin Hubから外へ出さない。
 */
function vNextBuildPublicAiEvidence_(events) {
  return (events || []).map(function (event) {
    var metadata = {};
    var rawText = String(event.evidenceText || event.evidence_text || '');
    if (rawText && rawText.charAt(0) === '{') {
      try { metadata = vNextParseJsonValue_(rawText, {}); } catch (ignore) { metadata = {}; }
    }
    var summary = String(metadata.summary || event.summary || event.target || '').replace(/[\u0000-\u001f]+/g, ' ').trim().slice(0, 280);
    var url = String(event.sourceUrl || event.source_url || metadata.sourceUrl || '').trim();
    if (!/^https?:\/\//i.test(url)) url = '';
    return {
      target: String(event.target || '').trim().slice(0, 120),
      direction: vNextNormalizeDirection_(event.direction || 'NEUTRAL'),
      summary: summary,
      sourceUrl: url,
      sourceDate: String(event.sourceDate || event.source_date || metadata.sourceDate || '').slice(0, 10),
      appliedAmount: Math.abs(Number(event.appliedAmount !== undefined ? event.appliedAmount : (event.amountMid !== undefined ? event.amountMid : event.amount_mid)) || 0),
      evidenceQuality: String(event.evidenceQuality || event.evidence_quality || '').toUpperCase().slice(0, 20),
      capApplied: Boolean(event.capApplied || Number(event.cap_applied || 0) === 1),
      researchAxis: String(metadata.researchAxis || 'ALTERNATIVE_SIGNALS').toUpperCase().slice(0, 40),
      signalType: String(metadata.signalType || '').trim().slice(0, 80),
      sourceStrength: String(metadata.sourceStrength || '').toUpperCase().slice(0, 40),
      forecastUse: String(metadata.forecastUse || (Number(event.applied_amount || 0) ? 'APPLY' : 'INSIGHT_ONLY')).toUpperCase().slice(0, 20),
      salesRelevance: String(metadata.salesRelevance || '').toUpperCase().slice(0, 20),
      humanQuestion: String(metadata.humanQuestion || '').replace(/[\u0000-\u001f]+/g, ' ').trim().slice(0, 180)
    };
  }).filter(function (item) {
    return item.summary || item.target || item.sourceUrl;
  }).sort(function (a, b) {
    var useDiff = (b.forecastUse === 'APPLY' ? 1 : 0) - (a.forecastUse === 'APPLY' ? 1 : 0);
    return useDiff || Number(b.appliedAmount || 0) - Number(a.appliedAmount || 0);
  }).slice(0, 5);
}

function vNextBuildContinuityPrior_(records, targetFiscalYear, cutoff) {
  var monthly = {};
  var baseMonthly = {};
  var spotMonthly = {};
  (records || []).forEach(function (record) {
    var date = record.actualDate;
    var fy = vNextFiscalYearForDate_(date);
    if (fy >= targetFiscalYear || fy < targetFiscalYear - VNEXT_ENGINE.MAX_HISTORY_YEARS) return;
    var key = vNextFormatMonth_(date);
    monthly[key] = (monthly[key] || 0) + Number(record.amount || 0);
    if (String(record.serviceType || '').toUpperCase() === 'SPOT') {
      spotMonthly[key] = (spotMonthly[key] || 0) + Number(record.amount || 0);
    } else {
      baseMonthly[key] = (baseMonthly[key] || 0) + Number(record.amount || 0);
    }
  });
  var fiscalYears = Object.keys(monthly).map(function (month) {
    return vNextFiscalYearForDate_(month + '-01');
  }).filter(function (value, index, array) { return array.indexOf(value) === index; })
    .sort(function (a, b) { return a - b; }).slice(-VNEXT_ENGINE.MAX_HISTORY_YEARS);
  if (fiscalYears.length < VNEXT_ENGINE.MIN_HISTORY_YEARS) {
    throw new Error('At least 5 fiscal years of confirmed actual history are required; found ' + fiscalYears.length + '.');
  }
  var fullYearShares = [];
  var annuals = [];
  var aggregateMonths = new Array(12).fill(0);
  fiscalYears.forEach(function (fy) {
    var amounts = [];
    for (var i = 0; i < 12; i++) {
      var month = vNextFormatMonth_(new Date(fy, 3 + i, 1));
      amounts.push(Number(baseMonthly[month] || 0));
      aggregateMonths[i] += Number(baseMonthly[month] || 0);
    }
    var observedMonths = Math.min(12, Math.max(0, (cutoff.getFullYear() - fy) * 12 + cutoff.getMonth() - 2));
    observedMonths = Math.min(12, observedMonths);
    var observedTotal = vNextSum_(amounts.slice(0, observedMonths));
    if (observedMonths === 12) {
      var total = vNextSum_(amounts);
      annuals.push({ fy: fy, total: total, estimated: false });
      if (total > 0) fullYearShares.push(amounts.map(function (value) { return value / total; }));
    } else if (observedMonths > 0) {
      annuals.push({ fy: fy, partialTotal: observedTotal, observedMonths: observedMonths, amounts: amounts, estimated: true });
    }
  });
  var seasonal = new Array(12).fill(0);
  for (var m = 0; m < 12; m++) {
    if (fullYearShares.length) seasonal[m] = vNextMean_(fullYearShares.map(function (shares) { return shares[m]; }));
    else seasonal[m] = aggregateMonths[m] + 1;
    seasonal[m] += 1 / 1200;
  }
  var seasonalTotal = vNextSum_(seasonal);
  seasonal = seasonal.map(function (value) { return value / seasonalTotal; });
  var monthlyShareResiduals = [];
  fullYearShares.forEach(function (shares) {
    shares.forEach(function (share, monthIndex) {
      monthlyShareResiduals.push(Math.log(
        Math.max(0.000001, Number(share || 0)) /
        Math.max(0.000001, Number(seasonal[monthIndex] || 0))
      ));
    });
  });
  var monthlyCommonShockSigma = monthlyShareResiduals.length >= 12
    ? vNextClamp_(
      vNextStdDev_(monthlyShareResiduals),
      VNEXT_ENGINE.MIN_MONTH_COMMON_SHOCK_SIGMA,
      VNEXT_ENGINE.MAX_MONTH_COMMON_SHOCK_SIGMA
    )
    : VNEXT_ENGINE.DEFAULT_MONTH_COMMON_SHOCK_SIGMA;
  annuals.forEach(function (row) {
    if (!row.estimated) return;
    var observedShare = vNextSum_(seasonal.slice(0, row.observedMonths));
    row.total = observedShare > 0 ? row.partialTotal / observedShare : row.partialTotal;
  });
  var positive = annuals.filter(function (row) { return row.total > 0; });
  var logs = positive.map(function (row) { return Math.log(row.total); });
  var slope = logs.length >= 2 ? vNextClamp_(vNextLinearSlope_(logs) * 0.50, -0.18, 0.18) : 0;
  var baseline = 0;
  var sigma = 0.20;
  if (logs.length) {
    var weights = positive.map(function (row, index) { return index + 1; });
    var weightedLog = vNextWeightedMean_(logs, weights);
    var latestLog = logs[logs.length - 1];
    baseline = Math.exp(latestLog * 0.65 + weightedLog * 0.35 + slope);
    var residuals = logs.map(function (value, index) {
      return value - (logs[0] + slope * index);
    });
    sigma = vNextClamp_(vNextStdDev_(residuals), 0.08, 0.45);
  }
  var unknownSpotModel = vNextBuildUnknownSpotModel_(spotMonthly, fiscalYears, cutoff);
  return {
    fiscalYears: fiscalYears,
    annualBaseline: baseline + unknownSpotModel.expectedAnnual,
    baseAnnualBaseline: baseline,
    logTrend: slope,
    logSigma: sigma,
    monthlyCommonShockSigma: monthlyCommonShockSigma,
    seasonalShares: seasonal,
    annualHistory: annuals.map(function (row) { return { fiscalYear: row.fy, baseTotal: row.total, estimated: row.estimated }; }),
    unknownSpotModel: unknownSpotModel
  };
}

function vNextBuildUnknownSpotModel_(spotMonthly, fiscalYears, cutoff) {
  var monthModels = [];
  var totalOccurrences = 0;
  var totalExposures = 0;
  for (var monthIndex = 0; monthIndex < 12; monthIndex++) {
    var amounts = [];
    var exposures = 0;
    fiscalYears.forEach(function (fy) {
      var monthDate = new Date(fy, 3 + monthIndex, 1);
      if (monthDate > cutoff) return;
      exposures++;
      var amount = Math.max(0, Number(spotMonthly[vNextFormatMonth_(monthDate)] || 0));
      if (amount > 0) amounts.push(amount);
    });
    var occurrences = amounts.length;
    var probability = occurrences === 0 || exposures === 0 ? 0 : (occurrences + 0.5) / (exposures + 1);
    var logAmounts = amounts.map(function (amount) { return Math.log(amount); });
    var logMean = logAmounts.length ? vNextMean_(logAmounts) : 0;
    var logSigma = logAmounts.length >= 2 ? vNextClamp_(vNextStdDev_(logAmounts), 0.10, 0.75) : 0.25;
    var expectedConditionalAmount = logAmounts.length ? Math.exp(logMean + 0.5 * logSigma * logSigma) : 0;
    monthModels.push({
      fiscalMonthIndex: monthIndex,
      exposures: exposures,
      occurrences: occurrences,
      probability: probability,
      amountMedian: logAmounts.length ? Math.exp(logMean) : 0,
      logSigma: logSigma,
      expectedAmount: probability * expectedConditionalAmount
    });
    totalOccurrences += occurrences;
    totalExposures += exposures;
  }
  return {
    months: monthModels,
    expectedAnnual: vNextSum_(monthModels.map(function (model) { return model.expectedAmount; })),
    expectedOccurrences: vNextSum_(monthModels.map(function (model) { return model.probability; })),
    meanOccurrenceRate: totalExposures ? totalOccurrences / totalExposures : 0,
    historyOccurrenceCount: totalOccurrences,
    historyExposureMonths: totalExposures
  };
}

/** One seeded month shock is shared by BASE and every evidence layer in a path. */
function vNextSampleMonthlyCommonShock_(seasonalShares, sigma, gaussian) {
  var safeSigma = vNextClamp_(
    Number(sigma || VNEXT_ENGINE.DEFAULT_MONTH_COMMON_SHOCK_SIGMA),
    VNEXT_ENGINE.MIN_MONTH_COMMON_SHOCK_SIGMA,
    VNEXT_ENGINE.MAX_MONTH_COMMON_SHOCK_SIGMA
  );
  var raw = [];
  var multipliers = [];
  (seasonalShares || []).forEach(function (share) {
    var multiplier = Math.exp(gaussian() * safeSigma - 0.5 * safeSigma * safeSigma);
    multipliers.push(multiplier);
    raw.push(Math.max(0.000001, Number(share || 0)) * multiplier);
  });
  var total = vNextSum_(raw);
  return {
    sigma: safeSigma,
    multipliers: multipliers,
    shares: raw.map(function (value) { return total > 0 ? value / total : 1 / 12; })
  };
}

function vNextSampleUnknownSpot_(model, suppressionByMonth, rng, gaussian, uncertaintyMultiplier, monthMultipliers) {
  var output = new Array(12).fill(0);
  if (!model || !model.months) return output;
  model.months.forEach(function (monthModel, index) {
    var suppression = vNextClamp_(Number(suppressionByMonth && suppressionByMonth[index] || 0), 0, 1);
    var probability = vNextClamp_(Number(monthModel.probability || 0) * (1 - suppression), 0, 1);
    if (rng() > probability || Number(monthModel.amountMedian || 0) <= 0) return;
    var multiplier = Math.max(1, Number(uncertaintyMultiplier || 1));
    output[index] = Number(monthModel.amountMedian) *
      Math.exp(gaussian() * Number(monthModel.logSigma || 0) * multiplier) *
      Math.max(0.000001, Number(monthMultipliers && monthMultipliers[index] || 1));
  });
  return output;
}

/** Known upward commitments suppress 50% of duplicate unknown-SPOT probability in their target months. */
function vNextCommitmentSpotSuppression_(events, fiscalYear) {
  var output = new Array(12).fill(0);
  (events || []).forEach(function (event) {
    if (vNextNormalizeDirection_(event.direction || 'UP') === 'DOWN') return;
    var indexes = vNextEventMonthIndexes_(event, fiscalYear);
    indexes.forEach(function (index) { output[index] = Math.max(output[index], VNEXT_ENGINE.KNOWN_SPOT_SUPPRESSION_RATE); });
  });
  return output;
}

/** Selects whole simulation paths, so every reported Q/month sum equals its annual percentile path. */
function vNextBuildCoherentPathSummary_(paths, fiscalYear) {
  if (!paths || !paths.length) throw new Error('At least one simulation path is required.');
  var order = paths.map(function (path, index) { return { value: vNextSum_(path), index: index }; })
    .sort(function (a, b) { return a.value - b.value; });
  function selected(quantile) { return order[Math.floor((order.length - 1) * quantile)]; }
  var picks = { p10: selected(0.10), p50: selected(0.50), p90: selected(0.90) };
  var targetMonths = [];
  for (var monthIndex = 0; monthIndex < 12; monthIndex++) targetMonths.push(new Date(fiscalYear, 3 + monthIndex, 1));
  var months = targetMonths.map(function (month, index) {
    var marginal = paths.map(function (path) { return Number(path[index] || 0); }).sort(function (a, b) { return a - b; });
    return {
      month: vNextFormatMonth_(month),
      p10: Number(paths[picks.p10.index][index] || 0),
      p50: Number(paths[picks.p50.index][index] || 0),
      p90: Number(paths[picks.p90.index][index] || 0),
      marginalP10: vNextPercentileSorted_(marginal, 0.10),
      marginalP50: vNextPercentileSorted_(marginal, 0.50),
      marginalP90: vNextPercentileSorted_(marginal, 0.90)
    };
  });
  var quarters = [0, 1, 2, 3].map(function (quarter) {
    var start = quarter * 3;
    return {
      quarter: 'Q' + (quarter + 1),
      p10: vNextSum_(paths[picks.p10.index].slice(start, start + 3)),
      p50: vNextSum_(paths[picks.p50.index].slice(start, start + 3)),
      p90: vNextSum_(paths[picks.p90.index].slice(start, start + 3))
    };
  });
  return {
    annual: { p10: picks.p10.value, p50: picks.p50.value, p90: picks.p90.value },
    quarters: quarters,
    months: months,
    pathIndexes: { p10: picks.p10.index, p50: picks.p50.index, p90: picks.p90.index }
  };
}

function vNextSampleEventsLayer_(events, baseAnnual, seasonalShares, fiscalYear, rng, gaussian, options) {
  var opt = options || {};
  var layer = new Array(12).fill(0);
  (events || []).forEach(function (event, eventIndex) {
    var diagnostic = opt.diagnostics && opt.diagnostics[eventIndex];
    if (diagnostic) diagnostic.pathCount++;
    var probability = vNextEventProbability_(event);
    if (rng() > probability) return;
    if (diagnostic) diagnostic.occurrenceCount++;
    var amount = vNextSampleEventAmount_(event, baseAnnual, rng, gaussian);
    var direction = vNextNormalizeDirection_(event.direction || 'UP');
    if (direction === 'DOWN') amount = -Math.abs(amount);
    else if (direction === 'NEUTRAL') amount = 0;
    else amount = Math.abs(amount);
    var indexes = vNextEventMonthIndexes_(event, fiscalYear);
    var hasRecognitionMonth = indexes.length > 0;
    if (!indexes.length) indexes = seasonalShares.map(function (value, index) { return index; });
    var originalIndexCount = indexes.length;
    var originalWeights = indexes.map(function (index) { return Math.max(0.000001, seasonalShares[index]); });
    var originalWeightTotal = vNextSum_(originalWeights);
    var delayMonths = opt.recognitionDelay
      ? vNextSampleEventDelayMonths_(event, hasRecognitionMonth, rng)
      : 0;
    if (diagnostic) {
      diagnostic.sampledAmountTotal += Math.abs(amount);
      diagnostic.delayMonthTotal += delayMonths;
      if (delayMonths > 0) diagnostic.delayedOccurrenceCount++;
    }
    var shifted = indexes.map(function (index, position) {
      return { index: index + delayMonths, weight: originalWeights[position] / originalWeightTotal };
    }).filter(function (item) { return item.index >= 0 && item.index <= 11; });
    if (diagnostic && shifted.length < originalIndexCount) diagnostic.partlyOutsideFiscalYearCount++;
    if (!shifted.length) {
      if (diagnostic) diagnostic.fullyOutsideFiscalYearCount++;
      return;
    }
    shifted.forEach(function (item) { layer[item.index] += amount * item.weight; });
    if (diagnostic) {
      diagnostic.recognizedAmountTotal += Math.abs(amount) * vNextSum_(shifted.map(function (item) { return item.weight; }));
    }
  });
  return layer;
}

function vNextInitializeCommitmentDiagnostics_(events, baseAnnual, fiscalYear) {
  return (events || []).map(function (event, index) {
    var indexes = vNextEventMonthIndexes_(event, fiscalYear);
    var distribution = vNextEventDelayDistribution_(event, indexes.length > 0);
    return {
      evidenceId: String(event.evidenceId || event.evidence_id || ('COMMITMENT-' + (index + 1))),
      target: String(event.target || ''),
      targetMonths: indexes.map(function (monthIndex) {
        return vNextFormatMonth_(new Date(fiscalYear, 3 + monthIndex, 1));
      }),
      occurrenceProbability: vNextEventProbability_(event),
      amountRange: vNextEventAmountRange_(event, baseAnnual),
      delayDistribution: distribution,
      expectedDelayMonths: vNextSum_(distribution.map(function (item) {
        return item.months * item.probability;
      })),
      pathCount: 0,
      occurrenceCount: 0,
      delayedOccurrenceCount: 0,
      partlyOutsideFiscalYearCount: 0,
      fullyOutsideFiscalYearCount: 0,
      delayMonthTotal: 0,
      sampledAmountTotal: 0,
      recognizedAmountTotal: 0
    };
  });
}

function vNextFinalizeCommitmentDiagnostics_(diagnostics) {
  return (diagnostics || []).map(function (row) {
    var occurrences = Number(row.occurrenceCount || 0);
    return {
      evidenceId: row.evidenceId,
      target: row.target,
      targetMonths: row.targetMonths,
      occurrenceProbability: row.occurrenceProbability,
      amountRange: row.amountRange,
      delayDistribution: row.delayDistribution,
      expectedDelayMonths: row.expectedDelayMonths,
      sampledOccurrenceRate: row.pathCount ? occurrences / row.pathCount : 0,
      sampledDelayedOccurrenceRate: occurrences ? row.delayedOccurrenceCount / occurrences : 0,
      sampledMeanDelayMonths: occurrences ? row.delayMonthTotal / occurrences : 0,
      partlyOutsideFiscalYearRate: occurrences ? row.partlyOutsideFiscalYearCount / occurrences : 0,
      fullyOutsideFiscalYearRate: occurrences ? row.fullyOutsideFiscalYearCount / occurrences : 0,
      sampledMeanAmount: occurrences ? row.sampledAmountTotal / occurrences : 0,
      recognizedAmountRate: row.sampledAmountTotal > 0 ? row.recognizedAmountTotal / row.sampledAmountTotal : 0
    };
  });
}

function vNextEventMetadata_(event) {
  if (event && event.metadata && typeof event.metadata === 'object') return event.metadata;
  var raw = event && (event.evidenceText || event.evidence_text);
  if (!raw || typeof raw === 'object') return raw && typeof raw === 'object' ? raw : {};
  var parsed = vNextParseJsonValue_(raw, {});
  return parsed && typeof parsed === 'object' && !Array.isArray(parsed) ? parsed : {};
}

function vNextEventDelayDistribution_(event, hasRecognitionMonth) {
  if (!hasRecognitionMonth) return [{ months: 0, probability: 1 }];
  var metadata = vNextEventMetadata_(event);
  var fixed = event.recognitionDelayMonths;
  if (!vNextIsFiniteNumber_(fixed)) fixed = event.delayMonths;
  if (!vNextIsFiniteNumber_(fixed)) fixed = metadata.recognitionDelayMonths;
  if (!vNextIsFiniteNumber_(fixed)) fixed = metadata.delayMonths;
  if (vNextIsFiniteNumber_(fixed)) {
    return [{ months: vNextClamp_(Math.round(Number(fixed)), 0, 12), probability: 1 }];
  }
  var raw = event.delayDistribution || event.delay_distribution || metadata.delayDistribution || metadata.delay_distribution;
  var rows = [];
  if (Array.isArray(raw)) {
    raw.forEach(function (item) {
      if (vNextIsFiniteNumber_(item)) rows.push({ months: Math.round(Number(item)), probability: 1 });
      else if (item && vNextIsFiniteNumber_(item.months) && vNextIsFiniteNumber_(item.probability)) {
        rows.push({ months: Math.round(Number(item.months)), probability: Number(item.probability) });
      }
    });
  } else if (raw && typeof raw === 'object') {
    Object.keys(raw).forEach(function (months) {
      if (vNextIsFiniteNumber_(months) && vNextIsFiniteNumber_(raw[months])) {
        rows.push({ months: Math.round(Number(months)), probability: Number(raw[months]) });
      }
    });
  }
  if (!rows.length) {
    var confidence = vNextNormalizeConfidence_(event.confidence || event.confidenceClass || event.confidence_class);
    if (confidence === 'CONFIRMED_FACT') rows = [{ months: 0, probability: 0.80 }, { months: 1, probability: 0.20 }];
    else if (confidence === 'LIKELY') rows = [{ months: 0, probability: 0.55 }, { months: 1, probability: 0.30 }, { months: 2, probability: 0.15 }];
    else if (confidence === 'HYPOTHESIS') rows = [{ months: 0, probability: 0.35 }, { months: 1, probability: 0.30 }, { months: 2, probability: 0.20 }, { months: 3, probability: 0.15 }];
    else rows = [{ months: 0, probability: 0.60 }, { months: 1, probability: 0.25 }, { months: 2, probability: 0.15 }];
  }
  var combined = {};
  rows.forEach(function (item) {
    var months = vNextClamp_(Math.round(Number(item.months || 0)), 0, 12);
    var probability = Math.max(0, Number(item.probability || 0));
    combined[months] = Number(combined[months] || 0) + probability;
  });
  var total = vNextSum_(Object.keys(combined).map(function (months) { return combined[months]; }));
  if (!(total > 0)) return [{ months: 0, probability: 1 }];
  return Object.keys(combined).map(Number).sort(function (a, b) { return a - b; }).map(function (months) {
    return { months: months, probability: combined[months] / total };
  });
}

function vNextSampleEventDelayMonths_(event, hasRecognitionMonth, rng) {
  var distribution = vNextEventDelayDistribution_(event, hasRecognitionMonth);
  var draw = rng();
  var cumulative = 0;
  for (var i = 0; i < distribution.length; i++) {
    cumulative += distribution[i].probability;
    if (draw <= cumulative || i === distribution.length - 1) return distribution[i].months;
  }
  return 0;
}

function vNextEventAmountRange_(event, baseAnnual) {
  if (vNextIsFiniteNumber_(event.rate || event.percent || event.percentage)) {
    var rate = Number(event.rate || event.percent || event.percentage);
    if (Math.abs(rate) > 1) rate /= 100;
    var rateAmount = Math.abs(Number(baseAnnual || 0) * rate);
    return { low: rateAmount, mid: rateAmount, high: rateAmount, mode: 'RATE' };
  }
  var mid = Number(event.amountMid !== undefined ? event.amountMid : (event.amount !== undefined ? event.amount : event.amount_mid));
  var low = Number(event.amountLow !== undefined ? event.amountLow : event.amount_low);
  var high = Number(event.amountHigh !== undefined ? event.amountHigh : event.amount_high);
  if (!isFinite(mid)) mid = vNextAmountBandMid_(event.amountBand || event.amount_band, baseAnnual);
  if (!isFinite(low)) low = mid * 0.75;
  if (!isFinite(high)) high = mid * 1.25;
  var minimum = Math.min(Math.abs(low), Math.abs(high));
  var maximum = Math.max(Math.abs(low), Math.abs(high), Math.abs(mid));
  return {
    low: minimum,
    mid: vNextClamp_(Math.abs(mid), minimum, maximum),
    high: maximum,
    mode: event.amountBand || event.amount_band ? 'BAND' : 'AMOUNT'
  };
}

/**
 * Normalizes reference-class priors before they enter the input hash.
 * Absolute-yen priors are accepted only when explicitly bound to this client;
 * growth priors are scale-independent and may be shared safely.
 */
function vNextNormalizeReferencePrior_(prior, clientId) {
  if (!prior || typeof prior !== 'object' || Array.isArray(prior) || !Object.keys(prior).length) return {};
  var input = prior;
  if (String(input.mode || '').toUpperCase() === 'DISABLED') {
    return { mode: 'DISABLED', reason: String(input.reason || 'DISABLED_BY_CONFIGURATION'), strength: 0 };
  }
  var hasOwn = Object.prototype.hasOwnProperty;
  var strengthRaw = hasOwn.call(input, 'strength') ? input.strength
    : (hasOwn.call(input, 'weight') ? input.weight : 0.15);
  if (!vNextIsFiniteNumber_(strengthRaw)) return { mode: 'DISABLED', reason: 'INVALID_STRENGTH', strength: 0 };
  var strength = vNextClamp_(Number(strengthRaw), 0, VNEXT_ENGINE.REFERENCE_MAX_STRENGTH);
  var hasGrowthMean = hasOwn.call(input, 'growthMean') || hasOwn.call(input, 'growth_mean');
  var hasGrowthStd = hasOwn.call(input, 'growthStd') || hasOwn.call(input, 'growth_std');
  var hasAnnualMean = hasOwn.call(input, 'annualMean') || hasOwn.call(input, 'mean');
  var hasAnnualStd = hasOwn.call(input, 'annualStd') || hasOwn.call(input, 'std');
  if ((hasGrowthMean || hasGrowthStd) && (hasAnnualMean || hasAnnualStd)) {
    return { mode: 'DISABLED', reason: 'MIXED_ABSOLUTE_AND_GROWTH_PRIOR', strength: 0 };
  }
  if (hasGrowthMean || hasGrowthStd) {
    var growthMeanRaw = hasOwn.call(input, 'growthMean') ? input.growthMean
      : (hasOwn.call(input, 'growth_mean') ? input.growth_mean : 0);
    var growthStdRaw = hasOwn.call(input, 'growthStd') ? input.growthStd
      : (hasOwn.call(input, 'growth_std') ? input.growth_std : 0);
    if (!vNextIsFiniteNumber_(growthMeanRaw) || !vNextIsFiniteNumber_(growthStdRaw)) {
      return { mode: 'DISABLED', reason: 'INVALID_GROWTH_PRIOR', strength: 0 };
    }
    var growthMean = Number(growthMeanRaw);
    var growthStd = Number(growthStdRaw);
    if (growthMean < VNEXT_ENGINE.REFERENCE_MIN_GROWTH || growthMean > VNEXT_ENGINE.REFERENCE_MAX_GROWTH ||
        growthStd < 0 || growthStd > VNEXT_ENGINE.REFERENCE_MAX_GROWTH_STD) {
      return { mode: 'DISABLED', reason: 'GROWTH_PRIOR_OUT_OF_RANGE', strength: 0 };
    }
    return { mode: 'GROWTH', growthMean: growthMean, growthStd: growthStd, strength: strength };
  }
  if (hasAnnualMean || hasAnnualStd) {
    var annualMeanRaw = hasOwn.call(input, 'annualMean') ? input.annualMean : input.mean;
    var annualStdRaw = hasOwn.call(input, 'annualStd') ? input.annualStd : (hasOwn.call(input, 'std') ? input.std : 0);
    if (!vNextIsFiniteNumber_(annualMeanRaw) || !vNextIsFiniteNumber_(annualStdRaw) ||
        Number(annualMeanRaw) < 0 || Number(annualStdRaw) < 0) {
      return { mode: 'DISABLED', reason: 'INVALID_ABSOLUTE_PRIOR', strength: 0 };
    }
    var scope = String(input.scope || input.priorScope || input.prior_scope || '').toUpperCase();
    var boundClientId = String(input.clientId || input.client_id || input.targetClientId || input.target_client_id || '');
    if (scope !== 'CLIENT' || !boundClientId || !clientId || boundClientId !== String(clientId)) {
      return { mode: 'DISABLED', reason: 'GLOBAL_OR_MISMATCHED_ABSOLUTE_PRIOR', strength: 0 };
    }
    return {
      mode: 'ABSOLUTE', scope: 'CLIENT', clientId: boundClientId,
      annualMean: Number(annualMeanRaw), annualStd: Number(annualStdRaw), strength: strength
    };
  }
  return { mode: 'DISABLED', reason: 'UNRECOGNIZED_REFERENCE_PRIOR', strength: 0 };
}

function vNextSampleReferenceLayer_(prior, currentMonths, seasonalShares, gaussian) {
  var zero = new Array(12).fill(0);
  if (!prior) return zero;
  if (!prior.mode) prior = vNextNormalizeReferencePrior_(prior, '');
  if (!prior.mode || prior.mode === 'DISABLED') return zero;
  var strength = vNextIsFiniteNumber_(prior.strength)
    ? vNextClamp_(Number(prior.strength), 0, VNEXT_ENGINE.REFERENCE_MAX_STRENGTH)
    : 0;
  if (strength === 0) return zero;
  var current = vNextSum_(currentMonths);
  var delta = 0;
  if (prior.mode === 'GROWTH' && vNextIsFiniteNumber_(prior.growthMean) && vNextIsFiniteNumber_(prior.growthStd)) {
    delta = current * Number(prior.growthMean) * strength +
      gaussian() * current * Math.max(0, Number(prior.growthStd)) * strength;
  } else if (prior.mode === 'ABSOLUTE' && prior.scope === 'CLIENT' &&
      vNextIsFiniteNumber_(prior.annualMean) && vNextIsFiniteNumber_(prior.annualStd)) {
    delta = (Number(prior.annualMean) - current) * strength +
      gaussian() * Math.max(0, Number(prior.annualStd)) * strength;
  } else return zero;
  return seasonalShares.map(function (share) { return delta * Number(share || 0); });
}

function vNextSampleEventAmount_(event, baseAnnual, rng, gaussian) {
  if (vNextIsFiniteNumber_(event.rate || event.percent || event.percentage)) {
    var rate = Number(event.rate || event.percent || event.percentage);
    if (Math.abs(rate) > 1) rate /= 100;
    return Math.abs(baseAnnual * rate);
  }
  var mid = Number(event.amountMid !== undefined ? event.amountMid : (event.amount !== undefined ? event.amount : event.amount_mid));
  var low = Number(event.amountLow !== undefined ? event.amountLow : event.amount_low);
  var high = Number(event.amountHigh !== undefined ? event.amountHigh : event.amount_high);
  if (!isFinite(mid)) mid = vNextAmountBandMid_(event.amountBand || event.amount_band, baseAnnual);
  if (!isFinite(low)) low = mid * 0.75;
  if (!isFinite(high)) high = mid * 1.25;
  low = Math.min(Math.abs(low), Math.abs(high));
  high = Math.max(Math.abs(low), Math.abs(high), Math.abs(mid));
  mid = vNextClamp_(Math.abs(mid), low, high);
  var u = rng();
  var modePoint = high === low ? 0 : (mid - low) / (high - low);
  if (high === low) return low;
  return u < modePoint
    ? low + Math.sqrt(u * (high - low) * (mid - low))
    : high - Math.sqrt((1 - u) * (high - low) * (high - mid));
}

function vNextEventProbability_(event) {
  if (vNextIsFiniteNumber_(event.probability)) return vNextClamp_(Number(event.probability), 0, 1);
  var confidence = vNextNormalizeConfidence_(event.confidence || event.confidenceClass || event.confidence_class);
  if (confidence === 'CONFIRMED_FACT') return 0.95;
  if (confidence === 'LIKELY') return 0.70;
  if (confidence === 'HYPOTHESIS') return 0.40;
  return 0.60;
}

function vNextEventMonthIndexes_(event, fiscalYear) {
  var startRaw = event.startMonth || event.targetStartMonth || event.target_start_month;
  var endRaw = event.endMonth || event.targetEndMonth || event.target_end_month || startRaw;
  if (!startRaw) return [];
  var start = vNextParseDate_(String(startRaw).length === 7 ? startRaw + '-01' : startRaw, 'event start');
  var end = vNextParseDate_(String(endRaw).length === 7 ? endRaw + '-01' : endRaw, 'event end');
  var fyStart = new Date(fiscalYear, 3, 1);
  var first = (start.getFullYear() - fyStart.getFullYear()) * 12 + start.getMonth() - fyStart.getMonth();
  var last = (end.getFullYear() - fyStart.getFullYear()) * 12 + end.getMonth() - fyStart.getMonth();
  var output = [];
  for (var i = Math.max(0, first); i <= Math.min(11, last); i++) output.push(i);
  return output;
}

function vNextApplyLayerNonNegative_(base, delta) {
  return base.map(function (value, index) { return Math.max(0, Number(value || 0) + Number(delta[index] || 0)); });
}

function vNextPersistForecastRun_(result, request) {
  var record = vNextResultToForecastRecord_(result, request);
  vNextAppendRecord_('FORECAST_RUN', record, { spreadsheet: request.spreadsheet });
  return record;
}

function vNextPersistFailedRun_(request, error, started) {
  var failureInfo = vNextForecastFailureInfo_(error);
  vNextAppendRecord_('FORECAST_RUN', {
    run_id: request.runId,
    book_id: request.bookId,
    client_id: request.clientId,
    client_name: request.clientName,
    fiscal_year: request.fiscalYear,
    as_of: vNextFormatDateOnly_(request.asOf),
    cutoff: vNextFormatDateOnly_(request.cutoff),
    seed: request.seed,
    input_data_hash: request.inputDataHash,
    model_release_id: request.modelReleaseId,
    schema_version: VNEXT_CORE.SCHEMA_VERSION,
    status: 'FAILED',
    is_official: 0,
    simulation_count: request.simulationCount,
    lens_json: vNextCanonicalJson_({
      runIdentity: request.runIdentity || null,
      versions: {
        core: VNEXT_CORE.VERSION,
        engine: VNEXT_ENGINE.VERSION,
        schema: VNEXT_CORE.SCHEMA_VERSION,
        bookSchema: request.bookSchemaVersion || '',
        template: request.templateVersion || '',
        modelReleaseId: request.modelReleaseId || ''
      },
      failure: failureInfo
    }),
    evidence_json: vNextCanonicalJson_({}),
    previous_run_id: request.previousRunId,
    created_at: started.toISOString(),
    created_by: request.createdBy,
    error_summary: String(error && error.message ? error.message : error).slice(0, 1000)
  }, { spreadsheet: request.spreadsheet });
}

function vNextResultToForecastRecord_(result, request) {
  return {
    run_id: result.runId,
    book_id: result.bookId,
    client_id: result.clientId,
    client_name: result.clientName,
    fiscal_year: result.fiscalYear,
    as_of: result.asOf,
    cutoff: result.cutoff,
    seed: result.seed,
    input_data_hash: result.inputDataHash,
    model_release_id: result.modelReleaseId,
    schema_version: VNEXT_CORE.SCHEMA_VERSION,
    status: result.status,
    official_vintage_id: result.officialVintageId || '',
    is_official: result.officialVintageId ? 1 : 0,
    history_years: result.historyYears.length,
    simulation_count: result.simulationCount,
    history_baseline: result.layers.historyBaseline,
    commitment_delta: result.layers.commitmentDelta,
    reference_delta: result.layers.referenceDelta,
    human_delta: result.layers.humanDelta,
    ai_delta: result.layers.aiDelta,
    objective_forecast: result.layers.objectiveForecast,
    system_recommended: result.layers.systemRecommended,
    p10: result.annual.p10,
    p50: result.annual.p50,
    p90: result.annual.p90,
    quarter_json: vNextCanonicalJson_(result.quarters),
    month_json: vNextCanonicalJson_(result.months),
    lens_json: vNextCanonicalJson_(result.lenses),
    evidence_json: vNextCanonicalJson_(result.evidenceSummary),
    previous_run_id: request.previousRunId,
    created_at: vNextNowIso_(),
    created_by: request.createdBy,
    error_summary: ''
  };
}

function vNextGetLatestForecast_(bookIdOrOptions, options) {
  try {
    var directSpreadsheet = options && typeof options.getSheetByName === 'function' ? options : null;
    var opt = typeof bookIdOrOptions === 'object' ? bookIdOrOptions : (directSpreadsheet ? {} : (options || {}));
    var bookId = typeof bookIdOrOptions === 'string' ? bookIdOrOptions : String(opt.bookId || vNextGetProperty_(VNEXT_CORE.BOOK_PROPERTY, true));
    var directIsStore = directSpreadsheet && directSpreadsheet.getSheetByName('FORECAST_RUN');
    var store = opt.spreadsheet || (directIsStore ? directSpreadsheet : undefined);
    var rows = vNextReadRecords_('FORECAST_RUN', { spreadsheet: store }).filter(function (row) {
      if (String(row.book_id || '') !== bookId) return false;
      if (opt.officialOnly && Number(row.is_official || 0) !== 1) return false;
      return String(row.status || '') === 'SUCCESS' || String(row.status || '') === 'OFFICIAL_LOCKED';
    });
    if (!rows.length) return null;
    return vNextForecastRecordToResult_(rows[rows.length - 1]);
  } catch (error) {
    vNextLog_('vNextGetLatestForecast_ failed', error);
    throw error;
  }
}

function vNextFreezeOfficialVintage_(request, options) {
  try {
    var data = typeof request === 'object' ? request : { bookId: request };
    var opt = options || data;
    var bookId = String(data.bookId || '');
    if (!bookId) throw new Error('bookId is required.');
    var store = opt.spreadsheet || vNextResolveActiveLocalAuditStore_(bookId);
    var authorization = vNextRequireActualRoleForBook_(bookId, store, ['ADMIN']);
    var requestedVintageId = String(data.officialVintageId || '');
    var allOfficialRows = vNextReadRecords_('FORECAST_RUN', { spreadsheet: store }).filter(function (row) {
      return Number(row.is_official || 0) === 1 || String(row.status || '') === 'OFFICIAL_LOCKED';
    });
    var sameVintageRows = requestedVintageId ? allOfficialRows.filter(function (row) {
      return String(row.official_vintage_id || '') === requestedVintageId;
    }) : [];
    if (sameVintageRows.length) {
      if (sameVintageRows.some(function (row) { return String(row.book_id || '') !== bookId; })) {
        throw new Error('officialVintageId already belongs to another book.');
      }
      var sameBookRows = sameVintageRows.filter(function (row) { return String(row.book_id || '') === bookId; });
      var existingSame = vNextForecastRecordToResult_(sameBookRows[sameBookRows.length - 1]);
      var retryRunId = String(data.runId || existingSame.previousRunId || '');
      var retrySource = retryRunId ? vNextFindForecastByRunId_(bookId, retryRunId, { spreadsheet: store }) : null;
      vNextAssertOfficialRetryConsistent_(existingSame, retrySource, retryRunId, bookId);
      existingSame.idempotent = true;
      return existingSame;
    }
    var existing = vNextGetLatestForecast_({ bookId: bookId, spreadsheet: store, officialOnly: true });
    if (authorization) {
      vNextValidateOfficialIssueRequest_(authorization.state, data.amendment === true, !!existing, data.amendmentReason);
    } else {
      vNextValidateOfficialIssueRequest_(data.amendment === true ? 'OFFICIAL_LOCKED' : 'SUBMITTED', data.amendment === true, !!existing, data.amendmentReason);
    }
    var source = data.runId
      ? vNextFindForecastByRunId_(bookId, data.runId, { spreadsheet: store })
      : vNextGetLatestForecast_({ bookId: bookId, spreadsheet: store });
    if (!source) throw new Error('Forecast run not found for official freeze.');
    if (source.status !== 'SUCCESS') throw new Error('Only a successful draft run can be frozen as official.');
    var vintageId = requestedVintageId || vNextUuid_();
    var official = {};
    Object.keys(source).forEach(function (key) { official[key] = source[key]; });
    official.runId = vNextUuid_();
    official.status = 'OFFICIAL_LOCKED';
    official.officialVintageId = vintageId;
    var record = vNextResultToForecastRecord_(official, {
      previousRunId: source.runId,
      createdBy: String(vNextActiveUserEmail_() || data.approvedBy || '').toLowerCase()
    });
    record.is_official = 1;
    record.status = 'OFFICIAL_LOCKED';
    record.official_vintage_id = vintageId;
    vNextAppendRecord_('FORECAST_RUN', record, { spreadsheet: store });
    return vNextForecastRecordToResult_(record);
  } catch (error) {
    vNextLog_('vNextFreezeOfficialVintage_ failed', error);
    throw error;
  }
}

function vNextValidateOfficialIssueRequest_(state, amendment, hasExistingOfficial, amendmentReason) {
  var currentState = String(state || '').toUpperCase();
  if (amendment) {
    if (!hasExistingOfficial) throw new Error('An amendment requires an existing official vintage for the same book.');
    if (!String(amendmentReason || '').trim()) throw new Error('amendmentReason is required.');
    if (['OFFICIAL_LOCKED', 'SUBMITTED'].indexOf(currentState) < 0) {
      throw new Error('Official amendment requires OFFICIAL_LOCKED or SUBMITTED state; current state=' + currentState);
    }
    return true;
  }
  if (currentState !== 'SUBMITTED') throw new Error('Official freeze requires SUBMITTED state; current state=' + currentState);
  if (hasExistingOfficial) throw new Error('Official vintage already exists; create an amendment instead of overwriting it.');
  return true;
}

function vNextAssertOfficialRetryConsistent_(official, source, requestedRunId, bookId) {
  if (!official || String(official.bookId || '') !== String(bookId || '')) throw new Error('Existing official vintage book mismatch.');
  if (requestedRunId && String(official.previousRunId || '') !== String(requestedRunId)) {
    throw new Error('Existing official vintage source run mismatch.');
  }
  if (requestedRunId && !source) throw new Error('Requested source run for existing official vintage was not found.');
  if (!source) return true;
  var officialSnapshot = vNextOfficialComparableSnapshot_(official);
  var sourceSnapshot = vNextOfficialComparableSnapshot_(source);
  if (vNextCanonicalJson_(officialSnapshot) !== vNextCanonicalJson_(sourceSnapshot)) {
    throw new Error('Existing official vintage content does not match its source run.');
  }
  return true;
}

function vNextOfficialComparableSnapshot_(forecast) {
  return {
    bookId: String(forecast.bookId || ''),
    clientId: String(forecast.clientId || ''),
    clientName: String(forecast.clientName || ''),
    fiscalYear: Number(forecast.fiscalYear),
    asOf: String(forecast.asOf || ''),
    cutoff: String(forecast.cutoff || ''),
    seed: Number(forecast.seed),
    inputDataHash: String(forecast.inputDataHash || ''),
    modelReleaseId: String(forecast.modelReleaseId || ''),
    simulationCount: Number(forecast.simulationCount),
    layers: forecast.layers || {},
    annual: forecast.annual || {},
    quarters: forecast.quarters || [],
    months: forecast.months || [],
    lenses: forecast.lenses || {},
    evidenceSummary: forecast.evidenceSummary || {}
  };
}

function vNextAppendPlanVersion_(payload, options) {
  try {
    var data = payload || {};
    var opt = options || {};
    var bookId = String(data.bookId || '');
    var runId = String(data.runId || '');
    if (!bookId || !runId) throw new Error('bookId and runId are required.');
    var store = opt.spreadsheet || vNextResolveActiveLocalAuditStore_(bookId);
    var authorization = vNextRequireActualRoleForBook_(bookId, store, ['FORECAST_OWNER', 'ADMIN']);
    var status = String(data.status || 'DRAFT').toUpperCase();
    if (['DRAFT', 'SUBMITTED', 'APPROVED', 'AMENDMENT'].indexOf(status) < 0) throw new Error('Invalid plan status: ' + status);
    if (authorization) {
      var currentState = String(authorization.state || '').toUpperCase();
      if (status === 'SUBMITTED' && ['DRAFT_READY', 'CHANGES_REQUESTED'].indexOf(currentState) < 0) {
        throw new Error('Plan submission requires DRAFT_READY or CHANGES_REQUESTED state; current state=' + currentState);
      }
      if ((status === 'APPROVED' || status === 'AMENDMENT') && authorization.role !== 'ADMIN') {
        throw new Error('Only an administrator can approve or amend a plan.');
      }
    }
    var forecast = typeof SpreadsheetApp !== 'undefined'
      ? vNextFindForecastByRunId_(bookId, runId, { spreadsheet: store })
      : null;
    if (typeof SpreadsheetApp !== 'undefined' && (!forecast || forecast.status !== 'SUCCESS')) {
      throw new Error('The plan must reference a successful forecast run for the same book.');
    }
    var system = Math.trunc(forecast ? Number(forecast.layers.systemRecommended) : Number(data.systemRecommended));
    var adoptionDelta = Math.trunc(Number(data.adoptionDelta || 0));
    var uplift = Math.trunc(Number(data.salesUplift || 0));
    if (!isFinite(system)) throw new Error('systemRecommended is required.');
    if (forecast && vNextIsFiniteNumber_(data.systemRecommended) && Math.abs(Number(data.systemRecommended) - system) > 1) {
      throw new Error('systemRecommended does not match the referenced forecast run.');
    }
    if (!isFinite(adoptionDelta)) throw new Error('adoptionDelta is invalid.');
    if (!isFinite(uplift) || uplift < 0) throw new Error('salesUplift must be a non-negative number.');
    if (adoptionDelta !== 0 && !String(data.adoptionReason || '').trim()) throw new Error('adoptionReason is required for a non-zero adoptionDelta.');
    if (uplift !== 0 && (!String(data.upliftReason || '').trim() || !String(data.upliftOwner || '').trim() || !String(data.upliftAction || '').trim() || !data.upliftDueDate)) {
      throw new Error('Non-zero sales uplift requires reason, owner, action and due date.');
    }
    var adopted = system + adoptionDelta;
    if (adopted < 0) throw new Error('adoptedForecast cannot be negative.');
    var allocation = vNextValidateUpliftAllocation_(data.upliftAllocation || [], uplift, forecast && forecast.fiscalYear);
    var latestPlan = typeof SpreadsheetApp !== 'undefined'
      ? vNextGetLatestPlanVersion_({ bookId: bookId, spreadsheet: store })
      : null;
    var actor = String(vNextActiveUserEmail_() || '').toLowerCase();
    var record = {
      plan_version_id: String(data.planVersionId || vNextUuid_()),
      book_id: bookId,
      run_id: runId,
      official_vintage_id: String(data.officialVintageId || ''),
      version_no: latestPlan ? Number(latestPlan.version_no || 0) + 1 : 1,
      status: status,
      system_recommended: system,
      adoption_delta: adoptionDelta,
      adoption_reason: String(data.adoptionReason || ''),
      adopted_forecast: adopted,
      sales_uplift: uplift,
      uplift_reason: String(data.upliftReason || ''),
      uplift_owner: String(data.upliftOwner || ''),
      uplift_action: String(data.upliftAction || ''),
      uplift_due_date: data.upliftDueDate ? vNextFormatDateOnly_(data.upliftDueDate) : '',
      uplift_allocation_json: vNextCanonicalJson_(allocation),
      final_budget: adopted + uplift,
      amends_plan_version_id: String(data.amendsPlanVersionId || ''),
      submitted_at: status === 'SUBMITTED' ? vNextNowIso_() : '',
      submitted_by: status === 'SUBMITTED' ? actor : '',
      approved_at: status === 'APPROVED' || status === 'AMENDMENT' ? vNextNowIso_() : '',
      approved_by: status === 'APPROVED' || status === 'AMENDMENT' ? actor : '',
      created_at: vNextNowIso_()
    };
    vNextAppendRecord_('PLAN_VERSION', record, { spreadsheet: store });
    return record;
  } catch (error) {
    vNextLog_('vNextAppendPlanVersion_ failed', error);
    throw error;
  }
}

function vNextGetLatestPlanVersion_(bookIdOrOptions, options) {
  var directSpreadsheet = options && typeof options.getSheetByName === 'function' ? options : null;
  var opt = typeof bookIdOrOptions === 'object' ? bookIdOrOptions : (directSpreadsheet ? {} : (options || {}));
  var bookId = typeof bookIdOrOptions === 'string' ? bookIdOrOptions : String(opt.bookId || '');
  var store = opt.spreadsheet || (directSpreadsheet || vNextResolveActiveLocalAuditStore_(bookId));
  var rows = vNextReadRecords_('PLAN_VERSION', { spreadsheet: store }).filter(function (row) {
    return String(row.book_id || '') === bookId;
  });
  return rows.length ? rows[rows.length - 1] : null;
}

function vNextValidateUpliftAllocation_(allocation, uplift, fiscalYear) {
  var source = (allocation || []).slice();
  if (uplift === 0 && !source.length) source = new Array(12).fill(0);
  if (source.length !== 12) {
    throw new Error('upliftAllocation must contain 12 non-negative monthly amounts.');
  }
  var fy = Number(fiscalYear);
  var normalized = source.map(function (item, index) {
    var expectedMonth = isFinite(fy) && fy > 0 ? vNextFormatMonth_(new Date(fy, 3 + index, 1)) : '';
    var isObject = item && typeof item === 'object' && !Array.isArray(item);
    var amount = Math.trunc(Number(isObject ? item.amount : item));
    if (!isFinite(amount) || amount < 0) throw new Error('upliftAllocation must contain 12 non-negative monthly amounts.');
    var suppliedMonth = isObject ? String(item.month || '') : '';
    if (suppliedMonth && expectedMonth && suppliedMonth !== expectedMonth) {
      throw new Error('upliftAllocation month mismatch at index ' + index + ': expected ' + expectedMonth + '.');
    }
    return { month: suppliedMonth || expectedMonth, amount: amount };
  });
  if (Math.abs(vNextSum_(normalized.map(function (item) { return item.amount; })) - uplift) > 1) {
    throw new Error('Monthly uplift allocation must equal salesUplift.');
  }
  return normalized;
}

function vNextRequireActualRoleForBook_(bookId, spreadsheet, allowedRoles) {
  if (typeof SpreadsheetApp === 'undefined') return null;
  var actor = vNextActiveUserEmail_();
  if (!actor) throw new Error('Signed-in execution identity is required.');
  var context = vNextGetBookContext_({ bookId: bookId, spreadsheet: spreadsheet, userEmail: actor });
  if ((allowedRoles || []).indexOf(String(context.role || '').toUpperCase()) < 0) {
    throw new Error('Role ' + context.role + ' is not authorized for this operation.');
  }
  return context;
}

/** Public, side-effect-free API for Admin Hub review preparation. */
function vNextEngineBuildEvaluationBreakdown(payload) {
  return vNextBuildAutomaticEvaluationBreakdown_(payload || {});
}

/**
 * Decomposes official system forecast minus confirmed actuals.
 * The seven annual components reconcile exactly. Seasonality and timing remain
 * explicit monthly diagnostics and therefore contribute zero to the FY total.
 */
function vNextBuildAutomaticEvaluationBreakdown_(payload) {
  var data = payload || {};
  var forecast = data.officialForecast || data.forecast || data.official || {};
  if (!forecast || !forecast.layers) throw new Error('officialForecast with layers is required.');
  if (!String(forecast.officialVintageId || forecast.official_vintage_id || '')) {
    throw new Error('An official vintage is required for automatic evaluation.');
  }
  if (!Array.isArray(data.actualBaseMonths) && !Array.isArray(data.actual_base_months)) {
    throw new Error('actualBaseMonths with twelve fiscal-month values is required.');
  }
  if (!Array.isArray(data.actualSpotMonths) && !Array.isArray(data.actual_spot_months)) {
    throw new Error('actualSpotMonths with twelve fiscal-month values is required.');
  }
  var actualBase = vNextEvaluationMonthlyVector_(data.actualBaseMonths || data.actual_base_months, 'actual BASE');
  var actualSpot = vNextEvaluationMonthlyVector_(data.actualSpotMonths || data.actual_spot_months, 'actual SPOT');
  var forecastMonths = vNextEvaluationMonthlyVector_(forecast.months || [], 'official forecast', 'p50');
  var actualMonths = actualBase.values.map(function (value, index) {
    return value + actualSpot.values[index];
  });
  var monthlyActualTotal = vNextSum_(actualMonths);
  var actualTotal = vNextIsFiniteNumber_(data.actualTotal !== undefined ? data.actualTotal : data.actual_total)
    ? Number(data.actualTotal !== undefined ? data.actualTotal : data.actual_total)
    : monthlyActualTotal;
  var layers = forecast.layers || {};
  var systemForecast = Number(layers.systemRecommended || layers.system_recommended || 0);
  if (!isFinite(systemForecast)) throw new Error('official systemRecommended is invalid.');
  var historyBaseline = Number(layers.historyBaseline || layers.history_baseline || 0);
  var commitmentDelta = Number(layers.commitmentDelta || layers.commitment_delta || 0);
  var referenceDelta = Number(layers.referenceDelta || layers.reference_delta || 0);
  var humanDelta = Number(layers.humanDelta || layers.human_delta || 0);
  var aiDelta = Number(layers.aiDelta || layers.ai_delta || 0);
  var expectedUnknownSpot = vNextEvaluationExpectedUnknownSpot_(forecast);
  expectedUnknownSpot = vNextClamp_(expectedUnknownSpot, 0, Math.max(0, historyBaseline));
  var forecastBase = historyBaseline - expectedUnknownSpot;
  var annual = {
    baseLevel: forecastBase - vNextSum_(actualBase.values),
    seasonality: 0,
    commitmentOutcome: commitmentDelta,
    amount: referenceDelta,
    timing: 0,
    unknownSpot: expectedUnknownSpot - vNextSum_(actualSpot.values),
    humanInfo: humanDelta,
    aiInfo: aiDelta,
    dataQuality: 0
  };
  var signedError = systemForecast - actualTotal;
  var beforeDataQuality = annual.baseLevel + annual.commitmentOutcome + annual.amount +
    annual.unknownSpot + annual.humanInfo + annual.aiInfo;
  annual.dataQuality = signedError - beforeDataQuality;
  var additiveComponents = {
    baseLevel: annual.baseLevel,
    unknownSpot: annual.unknownSpot,
    commitmentOutcome: annual.commitmentOutcome,
    amount: annual.amount,
    humanInfo: annual.humanInfo,
    aiInfo: annual.aiInfo,
    dataQuality: annual.dataQuality
  };
  var componentSum = vNextSum_(Object.keys(additiveComponents).map(function (key) {
    return additiveComponents[key];
  }));
  var shape = vNextEvaluationShapeDiagnostics_(forecastMonths.values, actualMonths);
  var issues = Array.isArray(data.dataQualityIssues || data.data_quality_issues)
    ? (data.dataQualityIssues || data.data_quality_issues).map(String).filter(Boolean)
    : [];
  if (actualBase.missingCount) issues.push('BASE実績の月次値が' + actualBase.missingCount + 'か月不足');
  if (actualSpot.missingCount) issues.push('SPOT実績の月次値が' + actualSpot.missingCount + 'か月不足');
  if (forecastMonths.missingCount) issues.push('公式予測の月次値が' + forecastMonths.missingCount + 'か月不足');
  if (Math.abs(monthlyActualTotal - actualTotal) > 1) issues.push('月次実績合計と確定年度実績が不一致');
  if (Math.abs(vNextSum_(forecastMonths.values) - systemForecast) > 1) issues.push('公式月次予測合計とシステム推奨予測が不一致');
  return {
    officialVintageId: String(forecast.officialVintageId || forecast.official_vintage_id || ''),
    sourceRunId: String(forecast.runId || forecast.run_id || ''),
    systemForecast: systemForecast,
    actualTotal: actualTotal,
    systemSignedError: signedError,
    errorComponents: annual,
    additiveComponents: additiveComponents,
    componentSum: componentSum,
    reconciliationResidual: signedError - componentSum,
    reconciled: Math.abs(signedError - componentSum) <= 1e-6,
    diagnostics: {
      seasonality: {
        annualContribution: 0,
        amountMagnitude: shape.seasonalityMagnitude,
        shareDistance: shape.shareDistance,
        monthlyShareDelta: shape.monthlyShareDelta
      },
      timing: {
        annualContribution: 0,
        amountMagnitude: shape.timingMagnitude,
        maximumCumulativeShareGap: shape.maximumCumulativeShareGap,
        forecastPeakMonthIndex: shape.forecastPeakMonthIndex,
        actualPeakMonthIndex: shape.actualPeakMonthIndex,
        peakShiftMonths: shape.peakShiftMonths
      },
      monthly: {
        forecast: forecastMonths.values,
        actualBase: actualBase.values,
        actualSpot: actualSpot.values,
        actualTotal: actualMonths,
        signedResidual: forecastMonths.values.map(function (value, index) {
          return value - actualMonths[index];
        })
      },
      dataQuality: {
        missingBaseMonths: actualBase.missingCount,
        missingSpotMonths: actualSpot.missingCount,
        missingForecastMonths: forecastMonths.missingCount,
        issues: issues
      }
    },
    notes: [
      'プラスはシステム予測が実績より大きかった方向、マイナスは実績が大きかった方向です。',
      '季節配分と認識月ずれは年度合計を変えないため、年度差への加算額は0円として月次診断に分けています。',
      '採用判断差分・営業上積み・最終予算はモデル誤差へ混ぜていません。'
    ]
  };
}

function vNextEvaluationMonthlyVector_(input, fieldName, preferredField) {
  var values = new Array(12).fill(0);
  var missingCount = 0;
  for (var index = 0; index < 12; index++) {
    var item = Array.isArray(input) ? input[index] : undefined;
    var raw = item;
    if (item && typeof item === 'object') {
      raw = preferredField && item[preferredField] !== undefined ? item[preferredField]
        : (item.amount !== undefined ? item.amount
          : (item.value !== undefined ? item.value
            : (item.actual !== undefined ? item.actual : item.p50)));
    }
    if (raw === '' || raw === null || raw === undefined || !isFinite(Number(raw))) {
      missingCount++;
      values[index] = 0;
    } else values[index] = Number(raw);
  }
  return { field: fieldName, values: values, missingCount: missingCount };
}

function vNextEvaluationExpectedUnknownSpot_(forecast) {
  var lenses = forecast && forecast.lenses || {};
  var continuity = lenses.continuity || {};
  var spot = continuity.unknownSpotModel || continuity.unknown_spot_model || {};
  var evidence = forecast && forecast.evidenceSummary || forecast && forecast.evidence_summary || {};
  var value = spot.expectedAnnual;
  if (!vNextIsFiniteNumber_(value)) value = spot.expected_annual;
  if (!vNextIsFiniteNumber_(value)) value = evidence.unknownSpotExpectedAnnual;
  if (!vNextIsFiniteNumber_(value)) value = evidence.unknown_spot_expected_annual;
  return vNextIsFiniteNumber_(value) ? Math.max(0, Number(value)) : 0;
}

function vNextEvaluationShapeDiagnostics_(forecastMonths, actualMonths) {
  function positiveShares(values) {
    var positive = values.map(function (value) { return Math.max(0, Number(value || 0)); });
    var total = vNextSum_(positive);
    return total > 0 ? positive.map(function (value) { return value / total; }) : new Array(12).fill(0);
  }
  function peakIndex(values) {
    var selected = 0;
    values.forEach(function (value, index) { if (value > values[selected]) selected = index; });
    return selected;
  }
  var forecastShares = positiveShares(forecastMonths);
  var actualShares = positiveShares(actualMonths);
  var shareDelta = forecastShares.map(function (value, index) { return value - actualShares[index]; });
  var shareDistance = 0.5 * vNextSum_(shareDelta.map(Math.abs));
  var cumulative = 0;
  var maximumGap = 0;
  for (var i = 0; i < 11; i++) {
    cumulative += shareDelta[i];
    maximumGap = Math.max(maximumGap, Math.abs(cumulative));
  }
  var scale = Math.min(Math.abs(vNextSum_(forecastMonths)), Math.abs(vNextSum_(actualMonths)));
  var forecastPeak = peakIndex(forecastMonths);
  var actualPeak = peakIndex(actualMonths);
  return {
    monthlyShareDelta: shareDelta,
    shareDistance: shareDistance,
    seasonalityMagnitude: shareDistance * scale,
    maximumCumulativeShareGap: maximumGap,
    timingMagnitude: maximumGap * scale,
    forecastPeakMonthIndex: forecastPeak,
    actualPeakMonthIndex: actualPeak,
    peakShiftMonths: actualPeak - forecastPeak
  };
}

function vNextAppendEvaluation_(payload, options) {
  try {
    var data = payload || {};
    var opt = options || {};
    var bookId = String(data.bookId || '');
    if (!bookId) throw new Error('bookId is required.');
    var store = opt.spreadsheet || vNextResolveActiveLocalAuditStore_(bookId);
    var authorization = vNextRequireActualRoleForBook_(bookId, store, ['ADMIN']);
    if (authorization && ['OFFICIAL_LOCKED', 'REVIEW_DUE', 'YEAR_CLOSED'].indexOf(String(authorization.state || '').toUpperCase()) < 0) {
      throw new Error('Evaluation requires an official/review state; current state=' + authorization.state);
    }
    var vintageId = String(data.officialVintageId || '');
    if (!vintageId) throw new Error('officialVintageId is required; latest/draft runs cannot be evaluated.');
    var official = typeof SpreadsheetApp === 'undefined' && data.officialForecast ? data.officialForecast : vNextGetLatestForecast_({
      bookId: bookId,
      spreadsheet: store,
      officialOnly: true
    });
    if (!official || official.officialVintageId !== vintageId) {
      throw new Error('The specified official vintage was not found.');
    }
    var actual = Number(data.actualTotal);
    if (!isFinite(actual)) throw new Error('actualTotal is required.');
    var system = Number(official.layers.systemRecommended);
    var adopted = vNextIsFiniteNumber_(data.adoptedForecast) ? Number(data.adoptedForecast) : '';
    var finalBudget = vNextIsFiniteNumber_(data.finalBudget) ? Number(data.finalBudget) : '';
    var signed = system - actual;
    var components = data.errorComponents || {};
    var record = {
      evaluation_id: String(data.evaluationId || vNextUuid_()),
      book_id: official.bookId,
      official_vintage_id: vintageId,
      source_run_id: official.runId,
      fiscal_year: official.fiscalYear,
      evaluated_at: data.evaluatedAt || vNextNowIso_(),
      actual_total: actual,
      system_forecast: system,
      adopted_forecast: adopted,
      final_budget: finalBudget,
      system_signed_error: signed,
      system_abs_error: Math.abs(signed),
      system_ape: actual !== 0 ? Math.abs(signed) / Math.abs(actual) : '',
      range_contains_actual: actual >= official.annual.p10 && actual <= official.annual.p90 ? 1 : 0,
      base_level_error: vNextNumberOrBlank_(components.baseLevel),
      seasonality_error: vNextNumberOrBlank_(components.seasonality),
      commitment_outcome_error: vNextNumberOrBlank_(components.commitmentOutcome),
      amount_error: vNextNumberOrBlank_(components.amount),
      timing_error: vNextNumberOrBlank_(components.timing),
      unknown_spot_error: vNextNumberOrBlank_(components.unknownSpot),
      human_info_error: vNextNumberOrBlank_(components.humanInfo),
      ai_info_error: vNextNumberOrBlank_(components.aiInfo),
      data_quality_error: vNextNumberOrBlank_(components.dataQuality),
      confirmed_cause: String(data.confirmedCause || ''),
      cause_hypothesis: String(data.causeHypothesis || ''),
      next_information_json: vNextCanonicalJson_(data.nextInformation || []),
      model_release_id: official.modelReleaseId,
      created_by: String(vNextActiveUserEmail_() || data.createdBy || '').toLowerCase()
    };
    vNextAppendRecord_('EVALUATION', record, { spreadsheet: store });
    return record;
  } catch (error) {
    vNextLog_('vNextAppendEvaluation_ failed', error);
    throw error;
  }
}

/** Learning payload intentionally excludes adoption_delta, uplift and final_budget. */
function vNextBuildLearningPayload_(forecast, evaluation) {
  return {
    runId: forecast.runId,
    officialVintageId: forecast.officialVintageId,
    modelReleaseId: forecast.modelReleaseId,
    inputDataHash: forecast.inputDataHash,
    systemForecast: forecast.layers.systemRecommended,
    interval: { p10: forecast.annual.p10, p50: forecast.annual.p50, p90: forecast.annual.p90 },
    actualTotal: Number(evaluation.actualTotal),
    errorComponents: evaluation.errorComponents || {},
    evidenceSummary: forecast.evidenceSummary || {}
  };
}

function vNextFindForecastByRunId_(bookId, runId, options) {
  var rows = vNextFindForecastRunRecords_(bookId, runId, options || {});
  return rows.length ? vNextForecastRecordToResult_(rows[rows.length - 1]) : null;
}

function vNextForecastRecordToResult_(record) {
  var quarters = vNextParseJsonValue_(record.quarter_json, []);
  var months = vNextParseJsonValue_(record.month_json, []);
  var lenses = vNextParseJsonValue_(record.lens_json, {});
  var evidence = vNextParseJsonValue_(record.evidence_json, {});
  return {
    runId: String(record.run_id || ''),
    bookId: String(record.book_id || ''),
    clientId: String(record.client_id || ''),
    clientName: String(record.client_name || ''),
    fiscalYear: Number(record.fiscal_year),
    asOf: String(record.as_of || ''),
    cutoff: String(record.cutoff || ''),
    seed: Number(record.seed),
    inputDataHash: String(record.input_data_hash || ''),
    modelReleaseId: String(record.model_release_id || ''),
    versions: lenses.versions || {
      core: '', engine: '', schema: String(record.schema_version || ''),
      bookSchema: String(record.schema_version || ''), template: '',
      modelReleaseId: String(record.model_release_id || '')
    },
    status: String(record.status || ''),
    officialVintageId: String(record.official_vintage_id || ''),
    previousRunId: String(record.previous_run_id || ''),
    historyYears: lenses.continuity && lenses.continuity.fiscalYears ? lenses.continuity.fiscalYears : [],
    simulationCount: Number(record.simulation_count),
    layers: {
      historyBaseline: Number(record.history_baseline),
      commitmentDelta: Number(record.commitment_delta),
      peerReferenceDelta: Number(lenses.changeReference && lenses.changeReference.peerReferenceDelta || 0),
      objectiveEventDelta: Number(lenses.changeReference && lenses.changeReference.objectiveEventDelta || 0),
      referenceDelta: Number(record.reference_delta),
      objectiveForecast: Number(record.objective_forecast),
      humanDelta: Number(record.human_delta),
      aiDelta: Number(record.ai_delta),
      systemRecommended: Number(record.system_recommended)
    },
    annual: { p10: Number(record.p10), p50: Number(record.p50), p90: Number(record.p90) },
    quarters: quarters,
    months: months,
    lenses: lenses,
    evidenceSummary: evidence
  };
}

function vNextActualRecordForHash_(record) {
  return {
    client: record.client,
    actualDate: vNextFormatDateOnly_(record.actualDate),
    amount: record.amount,
    serviceType: record.serviceType,
    product: record.product,
    sourceSheet: record.sourceSheet,
    sourceRow: record.sourceRow
  };
}

function vNextCanonicalSort_(a, b) {
  var left = vNextCanonicalJson_(a);
  var right = vNextCanonicalJson_(b);
  return left < right ? -1 : (left > right ? 1 : 0);
}

function vNextCreateGaussianSampler_(rng) {
  var spare = null;
  return function () {
    if (spare !== null) { var value = spare; spare = null; return value; }
    var u = 0, v = 0;
    while (u === 0) u = rng();
    while (v === 0) v = rng();
    var radius = Math.sqrt(-2 * Math.log(u));
    spare = radius * Math.sin(2 * Math.PI * v);
    return radius * Math.cos(2 * Math.PI * v);
  };
}

function vNextAmountBandMid_(band, annual) {
  var key = String(band || '').toUpperCase();
  if (key === 'SMALL' || key === '小') return annual * 0.02;
  if (key === 'LARGE' || key === '大') return annual * 0.10;
  return annual * 0.05;
}

function vNextClassifyServiceType_(value) {
  var text = String(value || '').toLowerCase();
  if (text.indexOf('ベース') >= 0) return 'BASE';
  if (text.indexOf('スポット') >= 0) return 'SPOT';
  if (['フラグメント', 'テンプレート', '運用更新', '簡便化', '保守サポート'].some(function (key) { return text.indexOf(key) >= 0; })) return 'BASE';
  if (['開発', 'その他', 'myinsights'].some(function (key) { return text.indexOf(key) >= 0; })) return 'SPOT';
  return 'OTHER';
}

function vNextSameClient_(a, b) {
  function normalize(value) {
    return String(value || '').replace(/[\s　]/g, '').replace(/（株）|\(株\)|株式会社/g, '').toLowerCase();
  }
  return normalize(a) === normalize(b);
}

function vNextParseJsonValue_(value, fallback) {
  if (value && typeof value === 'object') return value;
  try { return JSON.parse(String(value || '')); } catch (error) { return fallback; }
}

function vNextClamp_(value, minimum, maximum) {
  return Math.max(minimum, Math.min(maximum, value));
}

function vNextSum_(values) {
  return (values || []).reduce(function (sum, value) { return sum + Number(value || 0); }, 0);
}

function vNextMean_(values) {
  return values && values.length ? vNextSum_(values) / values.length : 0;
}

function vNextWeightedMean_(values, weights) {
  var denominator = vNextSum_(weights);
  return denominator ? values.reduce(function (sum, value, index) { return sum + value * weights[index]; }, 0) / denominator : 0;
}

function vNextStdDev_(values) {
  if (!values || values.length < 2) return 0;
  var mean = vNextMean_(values);
  return Math.sqrt(vNextMean_(values.map(function (value) { return Math.pow(value - mean, 2); })));
}

function vNextLinearSlope_(values) {
  if (!values || values.length < 2) return 0;
  var xMean = (values.length - 1) / 2;
  var yMean = vNextMean_(values);
  var numerator = 0, denominator = 0;
  values.forEach(function (value, index) {
    numerator += (index - xMean) * (value - yMean);
    denominator += Math.pow(index - xMean, 2);
  });
  return denominator ? numerator / denominator : 0;
}

function vNextPercentileSorted_(sorted, quantile) {
  if (!sorted.length) return 0;
  var position = (sorted.length - 1) * quantile;
  var base = Math.floor(position);
  var remainder = position - base;
  return sorted[base + 1] === undefined ? sorted[base] : sorted[base] + remainder * (sorted[base + 1] - sorted[base]);
}
