/**
 * Forecast vNext Client FY Book data bridge.
 * Forecast calculation is never performed here; requests are appended locally for the service worker.
 */

function vNextEngineRunForecast(request) {
  try {
    return vNextQueueClientForecastRequest(request || {});
  } catch (error) {
    vNextClientLog_('vNextEngineRunForecast failed', error);
    throw error;
  }
}

function vNextQueueClientForecastRequest(request) {
  try {
    var requestInput = request || {};
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    return vNextClientWithDocumentLock_(function () {
      var context = vNextGetBookContext_({ spreadsheet: ss, bookId: requestInput.bookId });
      if (!context.isForecastOwner) throw new Error('予測を依頼できるのはForecast Ownerだけです。');
      if (context.state !== 'READY_TO_RUN') throw new Error('現在は予測を依頼できません。画面を更新してください。');
      var now = new Date();
      var requestedAt = now.toISOString();
      var asOf = vNextClientDateOnly_(now);
      var cutoffDate = new Date(now.getFullYear(), now.getMonth(), 0);
      var requestId = 'REQ-' + vNextUuid_();
      var payload = {
        requestId: requestId,
        bookId: context.bookId,
        clientId: context.clientId,
        clientName: context.clientName,
        fiscalYear: Number(context.fiscalYear),
        asOf: asOf,
        cutoff: vNextClientDateOnly_(cutoffDate),
        bookConfiguredAsOf: String(context.asOf || ''),
        requestedAt: requestedAt,
        requestedBy: context.userEmail
      };
      var requestJson = vNextCanonicalJson_(payload);
      var requestHash = vNextSha256Hex_(requestJson);
      vNextClientEnsureRequestSheet_(ss);
      vNextClientAppendRequestEvent_(ss, {
        request_event_id: 'REQEV-' + vNextUuid_(), request_id: requestId,
        book_id: context.bookId, event_type: 'REQUESTED', status: 'PENDING',
        request_hash: requestHash, request_json: requestJson, requested_at: requestedAt,
        requested_by: context.userEmail, related_job_id: '', related_run_id: '',
        detail_json: vNextCanonicalJson_({ source: 'CLIENT_BOOK' }), created_at: requestedAt
      });
      var state = vNextTransitionState_({
        spreadsheet: ss, bookId: context.bookId, fromState: 'READY_TO_RUN',
        toState: 'RUNNING', reason: 'forecast_requested:' + requestId, internalOperation: 'CLIENT_QUEUE',
        skipLock: true
      });
      return { ok: true, requestId: requestId, requestHash: requestHash, stateEventId: state.stateEventId };
    });
  } catch (error) {
    vNextClientLog_('vNextQueueClientForecastRequest failed', error);
    throw error;
  }
}

function vNextGetLatestForecast_(bookIdOrOptions, options) {
  try {
    var bookId = typeof bookIdOrOptions === 'string' ? bookIdOrOptions : String(bookIdOrOptions && bookIdOrOptions.bookId || '');
    var ss = options && typeof options.getSheetByName === 'function'
      ? options
      : options && options.spreadsheet || bookIdOrOptions && bookIdOrOptions.spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    if (!bookId) bookId = vNextGetBookContext_(ss).bookId;
    var rows = vNextReadRecords_('FORECAST_RUN', { spreadsheet: ss }).filter(function (row) {
      return String(row.book_id || '') === bookId && ['SUCCESS', 'OFFICIAL_LOCKED'].indexOf(String(row.status || '').toUpperCase()) >= 0;
    });
    if (!rows.length) return null;
    return vNextClientForecastResult_(rows[rows.length - 1]);
  } catch (error) {
    vNextClientLog_('vNextGetLatestForecast_ failed', error);
    throw error;
  }
}

function vNextGetLatestPlanVersion_(bookIdOrOptions, options) {
  try {
    var bookId = typeof bookIdOrOptions === 'string' ? bookIdOrOptions : String(bookIdOrOptions && bookIdOrOptions.bookId || '');
    var ss = options && typeof options.getSheetByName === 'function'
      ? options
      : options && options.spreadsheet || bookIdOrOptions && bookIdOrOptions.spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    var rows = vNextReadRecords_('PLAN_VERSION', { spreadsheet: ss }).filter(function (row) { return String(row.book_id || '') === bookId; });
    return rows.length ? rows[rows.length - 1] : null;
  } catch (error) {
    vNextClientLog_('vNextGetLatestPlanVersion_ failed', error);
    throw error;
  }
}

function vNextAppendPlanVersion_(payload, options) {
  try {
    var data = payload || {};
    var ss = options && options.spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    return vNextClientWithDocumentLock_(function () {
      var context = vNextGetBookContext_({ spreadsheet: ss, bookId: data.bookId });
      if (!context.isForecastOwner) throw new Error('計画案を提出できるのはForecast Ownerだけです。');
      if (['DRAFT_READY', 'CHANGES_REQUESTED'].indexOf(context.state) < 0) throw new Error('現在は計画案を提出できません。');
      var forecast = vNextClientFindForecast_(ss, context.bookId, data.runId);
      if (!forecast || forecast.status !== 'SUCCESS') throw new Error('提出に使用できる予測がありません。');
      var system = Math.trunc(Number(forecast.layers.systemRecommended));
      var adoptionDelta = Math.trunc(Number(data.adoptionDelta || 0));
      var uplift = Math.trunc(Number(data.salesUplift || 0));
      if (!isFinite(adoptionDelta) || !isFinite(uplift) || uplift < 0) throw new Error('計画金額を確認してください。');
      if (adoptionDelta !== 0 && !String(data.adoptionReason || '').trim()) throw new Error('採用判断の理由を入力してください。');
      if (uplift !== 0 && (!String(data.upliftReason || '').trim() || !String(data.upliftOwner || '').trim() || !String(data.upliftAction || '').trim() || !data.upliftDueDate)) {
        throw new Error('営業上積みには理由・責任者・行動・期限が必要です。');
      }
      var allocation = vNextClientValidateAllocation_(data.upliftAllocation || [], uplift, context.fiscalYear);
      var previous = vNextGetLatestPlanVersion_(context.bookId, ss);
      var adopted = system + adoptionDelta;
      if (adopted < 0) throw new Error('採用予測は0円以上にしてください。');
      var now = vNextNowIso_();
      var record = {
        plan_version_id: vNextUuid_(), book_id: context.bookId, run_id: forecast.runId,
        official_vintage_id: '', version_no: previous ? Number(previous.version_no || 0) + 1 : 1,
        status: 'SUBMITTED', system_recommended: system, adoption_delta: adoptionDelta,
        adoption_reason: String(data.adoptionReason || '').trim(), adopted_forecast: adopted,
        sales_uplift: uplift, uplift_reason: String(data.upliftReason || '').trim(),
        uplift_owner: String(data.upliftOwner || '').trim(), uplift_action: String(data.upliftAction || '').trim(),
        uplift_due_date: data.upliftDueDate ? vNextClientDateOnly_(data.upliftDueDate) : '',
        uplift_allocation_json: vNextCanonicalJson_(allocation), final_budget: adopted + uplift,
        amends_plan_version_id: String(data.amendsPlanVersionId || previous && previous.plan_version_id || ''),
        submitted_at: now, submitted_by: context.userEmail, approved_at: '', approved_by: '', created_at: now
      };
      vNextAppendRecord_('PLAN_VERSION', record, { spreadsheet: ss, skipLock: true });
      return record;
    });
  } catch (error) {
    vNextClientLog_('vNextAppendPlanVersion_ failed', error);
    throw error;
  }
}

function vNextClientFindForecast_(spreadsheet, bookId, runId) {
  var rows = vNextReadRecords_('FORECAST_RUN', { spreadsheet: spreadsheet }).filter(function (row) {
    return String(row.book_id || '') === String(bookId || '') && String(row.run_id || '') === String(runId || '');
  });
  return rows.length ? vNextClientForecastResult_(rows[rows.length - 1]) : null;
}

function vNextClientForecastResult_(record) {
  var lenses = vNextClientJsonValue_(record.lens_json, {});
  return {
    runId: String(record.run_id || ''), bookId: String(record.book_id || ''),
    clientId: String(record.client_id || ''), clientName: String(record.client_name || ''),
    fiscalYear: Number(record.fiscal_year), asOf: String(record.as_of || ''), cutoff: String(record.cutoff || ''),
    status: String(record.status || ''), officialVintageId: String(record.official_vintage_id || ''),
    layers: {
      historyBaseline: Number(record.history_baseline || 0), commitmentDelta: Number(record.commitment_delta || 0),
      referenceDelta: Number(record.reference_delta || 0), objectiveForecast: Number(record.objective_forecast || 0),
      humanDelta: Number(record.human_delta || 0), aiDelta: Number(record.ai_delta || 0),
      systemRecommended: Number(record.system_recommended || 0)
    },
    annual: { p10: Number(record.p10 || 0), p50: Number(record.p50 || 0), p90: Number(record.p90 || 0) },
    quarters: vNextClientJsonValue_(record.quarter_json, []),
    months: vNextClientJsonValue_(record.month_json, []),
    lenses: lenses,
    evidenceSummary: vNextClientJsonValue_(record.evidence_json, {}),
    drivers: Array.isArray(lenses.publicDrivers) ? lenses.publicDrivers.slice(0, 3) : [],
    nextInformation: Array.isArray(lenses.nextInformation) ? lenses.nextInformation.slice(0, 3) : [],
    changeReasons: Array.isArray(lenses.changeReasons) ? lenses.changeReasons.slice(0, 3) : []
  };
}

function vNextClientValidateAllocation_(allocation, uplift, fiscalYear) {
  var source = Array.isArray(allocation) ? allocation.slice() : [];
  if (uplift === 0 && source.length === 0) source = new Array(12).fill(0);
  if (source.length !== 12) throw new Error('営業上積みは12か月へ配分してください。');
  var total = 0;
  var output = source.map(function (item, index) {
    var amount = Math.trunc(Number(item && typeof item === 'object' ? item.amount : item));
    if (!isFinite(amount) || amount < 0) throw new Error('月次配分は0円以上にしてください。');
    total += amount;
    var monthDate = new Date(Number(fiscalYear), 3 + index, 1);
    return { month: monthDate.getFullYear() + '-' + ('0' + (monthDate.getMonth() + 1)).slice(-2), amount: amount };
  });
  if (Math.abs(total - uplift) > 1) throw new Error('月次配分の合計を営業上積みと一致させてください。');
  return output;
}

function vNextClientEnsureRequestSheet_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(VNEXT_CLIENT_CORE.REQUEST_SHEET) || spreadsheet.insertSheet(VNEXT_CLIENT_CORE.REQUEST_SHEET);
  vNextEnsureAppendOnlyHeader_(sheet, VNEXT_CLIENT_CORE.REQUEST_HEADERS);
  try { sheet.hideSheet(); } catch (error) { Logger.log('[vNext Client] request sheet hide skipped.'); }
  return sheet;
}

function vNextClientAppendRequestEvent_(spreadsheet, record) {
  var sheet = vNextClientEnsureRequestSheet_(spreadsheet);
  var row = VNEXT_CLIENT_CORE.REQUEST_HEADERS.map(function (header) { return vNextClientCellValue_(record[header]); });
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
  return record;
}

function vNextClientJsonValue_(value, fallback) {
  if (value && typeof value === 'object') return value;
  if (!value) return fallback;
  try { return JSON.parse(String(value)); } catch (error) { return fallback; }
}
