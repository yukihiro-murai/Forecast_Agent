/**
 * Forecast vNext shared employee portal local core.
 * It records creation requests in the bound Spreadsheet and reads only local directory projections.
 */

var VNEXT_PORTAL = Object.freeze({
  RUNTIME_VERSION: 'vnext-portal-1.0.0',
  REQUEST_SCHEMA_VERSION: 'vnext-portal-request-1',
  REQUEST_TYPE: 'CREATE_CLIENT_FY_BOOK',
  HOME_SHEET: 'ホーム',
  DIRECTORY_SHEET: 'PORTAL_DIRECTORY',
  REQUEST_SHEET: 'VN_PORTAL_REQUEST',
  MENU_NAME: '年度計画ポータル',
  PAYLOAD_KEYS: Object.freeze([
    'clientId', 'clientName', 'fiscalYear', 'forecastOwnerEmail', 'relatedMemberEmails',
    'requestId', 'requestType', 'requestedAt', 'requestedBy', 'schemaVersion'
  ]),
  PREVIEW_INPUT_KEYS: Object.freeze([
    'clientId', 'clientName', 'fiscalYear', 'forecastOwnerEmail', 'relatedMembersText'
  ]),
  SUBMIT_INPUT_KEYS: Object.freeze([
    'clientId', 'clientName', 'confirmSimilarDuplicates', 'duplicateCheckHash',
    'fiscalYear', 'forecastOwnerEmail', 'relatedMembersText'
  ]),
  REQUEST_HEADERS: Object.freeze([
    'request_event_id', 'request_id', 'event_type', 'status', 'request_hash', 'request_json',
    'fiscal_year', 'client_id', 'client_name', 'forecast_owner_email',
    'related_member_emails_json', 'requested_at', 'requested_by', 'related_book_id',
    'related_book_url', 'detail_code', 'detail_message', 'created_at', 'created_by'
  ]),
  DIRECTORY_HEADERS: Object.freeze([
    'directory_event_id', 'directory_key', 'fiscal_year', 'client_id', 'client_name',
    'forecast_owner_email', 'related_member_emails_json', 'state', 'center_forecast',
    'adopted_forecast', 'final_budget', 'next_action', 'client_book_url', 'request_id',
    'updated_at', 'updated_by'
  ]),
  ACTIVE_REQUEST_STATUSES: Object.freeze(['PENDING', 'VALIDATING', 'CREATING', 'COMPLETED']),
  REQUEST_STATUS_LABELS: Object.freeze({
    PENDING: '受付済み',
    VALIDATING: '内容確認中',
    CREATING: 'ブック作成中',
    COMPLETED: '利用できます',
    FAILED: '作成できませんでした',
    REJECTED: '確認が必要'
  }),
  STATUS_EVENT_PAIRS: Object.freeze({
    REQUESTED: 'PENDING',
    VALIDATION_STARTED: 'VALIDATING',
    CREATION_STARTED: 'CREATING',
    COMPLETED: 'COMPLETED',
    FAILED: 'FAILED',
    REJECTED: 'REJECTED'
  })
});

/** One-time initializer for a clean portal Spreadsheet. Safe to run repeatedly. */
function vNextPortalInitialize() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    vNextPortalEnsureStructure_(spreadsheet);
    if (typeof vNextPortalRefreshViews_ === 'function') vNextPortalRefreshViews_(spreadsheet);
    return { ok: true, runtimeVersion: VNEXT_PORTAL.RUNTIME_VERSION };
  } catch (error) {
    vNextPortalLog_('vNextPortalInitialize failed', error);
    throw error;
  }
}

/** Returns defaults for the creation sidebar. It does not restrict who may submit. */
function vNextPortalGetCreateModel() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    vNextPortalEnsureStructure_(spreadsheet);
    var fiscalYear = vNextPortalCurrentFiscalYear_(new Date());
    var actor = vNextPortalActiveUserEmail_();
    return {
      ok: true,
      runtimeVersion: VNEXT_PORTAL.RUNTIME_VERSION,
      defaultFiscalYear: fiscalYear + 1,
      fiscalYears: [fiscalYear - 1, fiscalYear, fiscalYear + 1, fiscalYear + 2],
      defaultForecastOwnerEmail: actor,
      instructions: '既存の計画がないか確認してから、作成を依頼します。受付後はホームで進み具合を確認できます。'
    };
  } catch (error) {
    vNextPortalLog_('vNextPortalGetCreateModel failed', error);
    throw new Error('作成画面を準備できませんでした。ポータルを再読み込みしてください。');
  }
}

/** Finds same-FY duplicate candidates from the local directory and pending requests. */
function vNextPortalPreviewCreation(input) {
  try {
    vNextPortalAssertExactKeys_(input, VNEXT_PORTAL.PREVIEW_INPUT_KEYS, '入力内容');
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    vNextPortalEnsureStructure_(spreadsheet);
    var normalized = vNextPortalNormalizeCreationInput_(input);
    var check = vNextPortalBuildDuplicateCheck_(spreadsheet, normalized);
    return {
      ok: true,
      normalized: {
        fiscalYear: normalized.fiscalYear,
        clientName: normalized.clientName,
        clientId: normalized.clientId,
        forecastOwnerEmail: normalized.forecastOwnerEmail,
        relatedMemberEmails: normalized.relatedMemberEmails
      },
      candidates: check.candidates,
      hasExact: check.hasExact,
      hasSimilar: check.hasSimilar,
      canSubmit: !check.hasExact,
      duplicateCheckHash: check.hash,
      message: check.hasExact
        ? '同じ年度の計画または作成依頼が見つかりました。新しく作らず、既存の計画を確認してください。'
        : check.hasSimilar
          ? '似た名前の候補があります。同じクライアントでないことを確認してください。'
          : '同じ年度の重複候補は見つかりませんでした。'
    };
  } catch (error) {
    vNextPortalLog_('vNextPortalPreviewCreation failed', error);
    throw error;
  }
}

/** Appends one immutable REQUESTED/PENDING event after rechecking duplicates under a lock. */
function vNextPortalSubmitCreationRequest(input) {
  try {
    vNextPortalAssertExactKeys_(input, VNEXT_PORTAL.SUBMIT_INPUT_KEYS, '送信内容');
    if (typeof input.confirmSimilarDuplicates !== 'boolean') {
      throw new Error('重複確認の状態が不正です。もう一度候補を確認してください。');
    }
    var duplicateCheckHash = String(input.duplicateCheckHash || '').trim().toLowerCase();
    if (!/^[a-f0-9]{64}$/.test(duplicateCheckHash)) {
      throw new Error('先に「重複候補を確認」を押してください。');
    }
    var normalized = vNextPortalNormalizeCreationInput_({
      fiscalYear: input.fiscalYear,
      clientName: input.clientName,
      clientId: input.clientId,
      forecastOwnerEmail: input.forecastOwnerEmail,
      relatedMembersText: input.relatedMembersText
    });
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    vNextPortalEnsureStructure_(spreadsheet);
    return vNextPortalWithDocumentLock_(function () {
      var check = vNextPortalBuildDuplicateCheck_(spreadsheet, normalized);
      if (check.hash !== duplicateCheckHash) {
        throw new Error('候補一覧が更新されました。もう一度「重複候補を確認」を押してください。');
      }
      if (check.hasExact) {
        throw new Error('同じ年度の計画または作成依頼があります。既存の計画を利用してください。');
      }
      if (check.hasSimilar && input.confirmSimilarDuplicates !== true) {
        throw new Error('似た名前の候補を確認し、「別のクライアントです」にチェックしてください。');
      }

      var actor = vNextPortalActiveUserEmail_();
      if (!actor) throw new Error('ログイン中のGoogleアカウントを確認できません。社内アカウントで開き直してください。');
      var now = new Date().toISOString();
      var requestId = 'PORTAL-REQ-' + Utilities.getUuid();
      var payload = {
        clientId: normalized.clientId,
        clientName: normalized.clientName,
        fiscalYear: normalized.fiscalYear,
        forecastOwnerEmail: normalized.forecastOwnerEmail,
        relatedMemberEmails: normalized.relatedMemberEmails,
        requestId: requestId,
        requestType: VNEXT_PORTAL.REQUEST_TYPE,
        requestedAt: now,
        requestedBy: actor,
        schemaVersion: VNEXT_PORTAL.REQUEST_SCHEMA_VERSION
      };
      vNextPortalValidateRequestPayload_(payload);
      var requestJson = vNextPortalCanonicalJson_(payload);
      var requestHash = vNextPortalSha256Hex_(requestJson);
      vNextPortalAppendRow_(spreadsheet, VNEXT_PORTAL.REQUEST_SHEET, VNEXT_PORTAL.REQUEST_HEADERS, {
        request_event_id: 'PORTAL-REQEV-' + Utilities.getUuid(),
        request_id: requestId,
        event_type: 'REQUESTED',
        status: 'PENDING',
        request_hash: requestHash,
        request_json: requestJson,
        fiscal_year: normalized.fiscalYear,
        client_id: normalized.clientId,
        client_name: normalized.clientName,
        forecast_owner_email: normalized.forecastOwnerEmail,
        related_member_emails_json: vNextPortalCanonicalJson_(normalized.relatedMemberEmails),
        requested_at: now,
        requested_by: actor,
        related_book_id: '',
        related_book_url: '',
        detail_code: '',
        detail_message: '作成依頼を受け付けました。',
        created_at: now,
        created_by: actor
      });
      return {
        ok: true,
        requestId: requestId,
        requestHash: requestHash,
        status: 'PENDING',
        statusLabel: VNEXT_PORTAL.REQUEST_STATUS_LABELS.PENDING,
        message: '作成依頼を受け付けました。ホームで進み具合を確認できます。'
      };
    });
  } catch (error) {
    vNextPortalLog_('vNextPortalSubmitCreationRequest failed', error);
    throw error;
  }
}

function vNextPortalGetLocalViewData_(spreadsheet) {
  spreadsheet = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  vNextPortalEnsureStructure_(spreadsheet);
  var directory = vNextPortalReadDirectory_(spreadsheet);
  var requests = vNextPortalReadRequestModels_(spreadsheet);
  var fiscalYear = vNextPortalCurrentFiscalYear_(new Date());
  var years = [fiscalYear, fiscalYear + 1];
  directory.concat(requests).forEach(function (item) {
    var year = Number(item.fiscalYear || 0);
    if (year && years.indexOf(year) < 0) years.push(year);
  });
  years.sort(function (a, b) { return b - a; });
  return { directory: directory, requests: requests, years: years, currentFiscalYear: fiscalYear };
}

function vNextPortalEnsureStructure_(spreadsheet) {
  var ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var home = ss.getSheetByName(VNEXT_PORTAL.HOME_SHEET);
  if (!home) home = ss.insertSheet(VNEXT_PORTAL.HOME_SHEET, 0);
  home.setTabColor('#174ea6');
  var currentFiscalYear = vNextPortalCurrentFiscalYear_(new Date());
  vNextPortalEnsureFiscalYearSheet_(ss, currentFiscalYear);
  vNextPortalEnsureFiscalYearSheet_(ss, currentFiscalYear + 1);
  var directory = vNextPortalEnsureTable_(ss, VNEXT_PORTAL.DIRECTORY_SHEET, VNEXT_PORTAL.DIRECTORY_HEADERS);
  var requests = vNextPortalEnsureTable_(ss, VNEXT_PORTAL.REQUEST_SHEET, VNEXT_PORTAL.REQUEST_HEADERS);
  vNextPortalHideInternalSheet_(directory);
  vNextPortalHideInternalSheet_(requests);
  vNextPortalHideBlankDefaultSheets_(ss);
  return { home: home, directory: directory, requests: requests };
}

function vNextPortalEnsureFiscalYearSheet_(spreadsheet, fiscalYear) {
  var name = 'FY' + Number(fiscalYear);
  var sheet = spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
  sheet.showSheet();
  sheet.setTabColor('#34a853');
  return sheet;
}

function vNextPortalEnsureTable_(spreadsheet, sheetName, headers) {
  var sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
  vNextPortalEnsureSheetSize_(sheet, 2, headers.length);
  var lastColumn = Math.max(sheet.getLastColumn(), headers.length);
  var current = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  var hasHeader = current.some(function (value) { return String(value || '').trim() !== ''; });
  if (!hasHeader) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers.slice()]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f1f3f4');
  } else {
    var actual = current.slice(0, headers.length).map(function (value) { return String(value || ''); });
    if (vNextPortalCanonicalJson_(actual) !== vNextPortalCanonicalJson_(headers.slice())) {
      throw new Error(sheetName + 'の列構成が正しくありません。管理担当者へ連絡してください。');
    }
    var unexpected = current.slice(headers.length).some(function (value) { return String(value || '').trim() !== ''; });
    if (unexpected) throw new Error(sheetName + 'に未定義の列があります。管理担当者へ連絡してください。');
  }
  // Employee-bound code must append only to the request staging log. The
  // Admin-owned directory projection may be hard-protected by the service and
  // must never be relaxed by an employee onOpen.
  if (sheetName === VNEXT_PORTAL.REQUEST_SHEET) {
    vNextPortalProtectWarningOnly_(sheet, '作成依頼ログ（通常は直接編集しません）');
  }
  return sheet;
}

function vNextPortalAppendRow_(spreadsheet, sheetName, headers, record) {
  vNextPortalAssertExactKeys_(record, headers, sheetName + 'レコード');
  var sheet = vNextPortalEnsureTable_(spreadsheet, sheetName, headers);
  var values = headers.map(function (header) {
    var value = record[header];
    return value === undefined || value === null ? '' : value;
  });
  if (sheet.getLastRow() + 1 > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), 1);
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, values.length).setValues([values]);
  return record;
}

function vNextPortalReadTable_(spreadsheet, sheetName, headers) {
  var sheet = vNextPortalEnsureTable_(spreadsheet, sheetName, headers);
  if (sheet.getLastRow() < 2) return [];
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  return values.filter(function (row) {
    return row.some(function (value) { return value !== '' && value !== null; });
  }).map(function (row) {
    var record = {};
    headers.forEach(function (header, index) { record[header] = row[index]; });
    return record;
  });
}

function vNextPortalReadDirectory_(spreadsheet) {
  var rows = vNextPortalReadTable_(spreadsheet, VNEXT_PORTAL.DIRECTORY_SHEET, VNEXT_PORTAL.DIRECTORY_HEADERS);
  var latestByKey = {};
  rows.forEach(function (row, index) {
    var key = String(row.directory_key || '').trim() ||
      [row.fiscal_year, row.client_id || vNextPortalNormalizeClientName_(row.client_name), row.request_id].join('|');
    latestByKey[key] = { row: row, index: index };
  });
  return Object.keys(latestByKey).map(function (key) {
    var row = latestByKey[key].row;
    return {
      source: 'DIRECTORY',
      directoryKey: key,
      requestId: String(row.request_id || '').trim(),
      fiscalYear: Number(row.fiscal_year || 0),
      clientId: String(row.client_id || '').trim(),
      clientName: String(row.client_name || '').trim(),
      forecastOwnerEmail: String(row.forecast_owner_email || '').trim().toLowerCase(),
      relatedMemberEmails: vNextPortalParseEmailArray_(row.related_member_emails_json),
      state: String(row.state || '').trim().toUpperCase(),
      centerForecast: vNextPortalOptionalNumber_(row.center_forecast),
      adoptedForecast: vNextPortalOptionalNumber_(row.adopted_forecast),
      finalBudget: vNextPortalOptionalNumber_(row.final_budget),
      nextAction: String(row.next_action || '').trim(),
      url: vNextPortalSafeBookUrl_(row.client_book_url),
      updatedAt: vNextPortalIsoText_(row.updated_at)
    };
  }).filter(function (item) { return item.fiscalYear && item.clientName; });
}

function vNextPortalReadRequestModels_(spreadsheet) {
  var rows = vNextPortalReadTable_(spreadsheet, VNEXT_PORTAL.REQUEST_SHEET, VNEXT_PORTAL.REQUEST_HEADERS);
  var grouped = {};
  rows.forEach(function (row, index) {
    var requestId = String(row.request_id || '').trim();
    if (!requestId) return;
    if (!grouped[requestId]) grouped[requestId] = [];
    grouped[requestId].push({ row: row, index: index });
  });
  return Object.keys(grouped).map(function (requestId) {
    var events = grouped[requestId];
    var requested = null;
    events.some(function (entry) {
      if (String(entry.row.event_type || '').toUpperCase() !== 'REQUESTED') return false;
      try {
        var payload = JSON.parse(String(entry.row.request_json || ''));
        vNextPortalValidateRequestPayload_(payload);
        if (payload.requestId !== requestId) throw new Error('requestId mismatch');
        if (vNextPortalSha256Hex_(vNextPortalCanonicalJson_(payload)) !== String(entry.row.request_hash || '').toLowerCase()) {
          throw new Error('request hash mismatch');
        }
        requested = { row: entry.row, payload: payload };
        return true;
      } catch (error) {
        vNextPortalLog_('invalid local request ignored: ' + requestId, error);
        return false;
      }
    });
    if (!requested) return null;
    var latest = requested.row;
    events.forEach(function (entry) {
      if (vNextPortalIsValidStatusEvent_(entry.row, requestId, requested.row.request_hash)) latest = entry.row;
    });
    var status = String(latest.status || requested.row.status || 'PENDING').trim().toUpperCase();
    return {
      source: 'REQUEST',
      requestId: requestId,
      requestHash: String(requested.row.request_hash || '').toLowerCase(),
      fiscalYear: Number(requested.payload.fiscalYear),
      clientId: requested.payload.clientId,
      clientName: requested.payload.clientName,
      forecastOwnerEmail: requested.payload.forecastOwnerEmail,
      relatedMemberEmails: requested.payload.relatedMemberEmails,
      requestedAt: requested.payload.requestedAt,
      requestedBy: requested.payload.requestedBy,
      status: status,
      statusLabel: VNEXT_PORTAL.REQUEST_STATUS_LABELS[status] || '処理状況を確認中',
      detailCode: String(latest.detail_code || '').trim(),
      detailMessage: String(latest.detail_message || '').trim(),
      relatedBookId: String(latest.related_book_id || '').trim(),
      url: vNextPortalSafeBookUrl_(latest.related_book_url),
      updatedAt: vNextPortalIsoText_(latest.created_at || requested.payload.requestedAt)
    };
  }).filter(Boolean);
}

function vNextPortalBuildDuplicateCheck_(spreadsheet, normalized) {
  var directory = vNextPortalReadDirectory_(spreadsheet);
  var requests = vNextPortalReadRequestModels_(spreadsheet).filter(function (request) {
    return VNEXT_PORTAL.ACTIVE_REQUEST_STATUSES.indexOf(request.status) >= 0;
  });
  var directoryRequestIds = {};
  directory.forEach(function (item) { if (item.requestId) directoryRequestIds[item.requestId] = true; });
  var sources = directory.concat(requests.filter(function (item) { return !directoryRequestIds[item.requestId]; }));
  var normalizedName = vNextPortalNormalizeClientName_(normalized.clientName);
  var normalizedId = vNextPortalNormalizeClientId_(normalized.clientId);
  var candidates = [];
  sources.forEach(function (item) {
    if (Number(item.fiscalYear) !== normalized.fiscalYear) return;
    var itemName = vNextPortalNormalizeClientName_(item.clientName);
    var itemId = vNextPortalNormalizeClientId_(item.clientId);
    var level = '';
    var score = 0;
    var reason = '';
    if (normalizedId && itemId && normalizedId === itemId) {
      level = 'EXACT'; score = 100; reason = 'クライアントIDが一致';
    } else if (normalizedName && normalizedName === itemName) {
      level = 'EXACT'; score = 95; reason = 'クライアント名が一致';
    } else {
      var similarity = vNextPortalNameSimilarity_(normalizedName, itemName);
      if (similarity >= 0.56 || (Math.min(normalizedName.length, itemName.length) >= 4 &&
          (normalizedName.indexOf(itemName) >= 0 || itemName.indexOf(normalizedName) >= 0))) {
        level = 'SIMILAR'; score = Math.round(similarity * 100); reason = '名前が似ています';
      }
    }
    if (!level) return;
    candidates.push({
      source: item.source,
      level: level,
      score: score,
      reason: reason,
      fiscalYear: item.fiscalYear,
      clientId: item.clientId,
      clientName: item.clientName,
      status: item.source === 'REQUEST' ? item.statusLabel : vNextPortalDirectoryStateLabel_(item.state),
      requestId: item.requestId || '',
      url: item.url || '',
      updatedAt: item.updatedAt || ''
    });
  });
  candidates.sort(function (a, b) {
    if (a.level !== b.level) return a.level === 'EXACT' ? -1 : 1;
    if (a.score !== b.score) return b.score - a.score;
    return String(a.clientName).localeCompare(String(b.clientName), 'ja');
  });
  var fingerprint = candidates.map(function (candidate) {
    return {
      clientId: candidate.clientId,
      clientName: candidate.clientName,
      fiscalYear: candidate.fiscalYear,
      level: candidate.level,
      requestId: candidate.requestId,
      source: candidate.source,
      status: candidate.status,
      updatedAt: candidate.updatedAt,
      url: candidate.url
    };
  });
  var hashInput = {
    candidates: fingerprint,
    clientId: normalizedId,
    clientName: normalizedName,
    fiscalYear: normalized.fiscalYear,
    forecastOwnerEmail: normalized.forecastOwnerEmail,
    relatedMemberEmails: normalized.relatedMemberEmails,
    schemaVersion: VNEXT_PORTAL.REQUEST_SCHEMA_VERSION
  };
  return {
    candidates: candidates,
    hasExact: candidates.some(function (item) { return item.level === 'EXACT'; }),
    hasSimilar: candidates.some(function (item) { return item.level === 'SIMILAR'; }),
    hash: vNextPortalSha256Hex_(vNextPortalCanonicalJson_(hashInput))
  };
}

function vNextPortalNormalizeCreationInput_(input) {
  var fiscalYear = Number(input.fiscalYear);
  if (!isFinite(fiscalYear) || Math.floor(fiscalYear) !== fiscalYear || fiscalYear < 2000 || fiscalYear > 2100) {
    throw new Error('対象年度を正しく選択してください。');
  }
  var clientName = vNextPortalPlainText_(input.clientName, 120, true, 'クライアント名');
  var clientId = vNextPortalPlainText_(input.clientId, 100, false, 'クライアントID');
  var forecastOwnerEmail = vNextPortalNormalizeEmail_(input.forecastOwnerEmail, true, 'Forecast Owner');
  var relatedMemberEmails = vNextPortalNormalizeEmailList_(input.relatedMembersText, forecastOwnerEmail);
  return {
    fiscalYear: fiscalYear,
    clientName: clientName,
    clientId: clientId,
    forecastOwnerEmail: forecastOwnerEmail,
    relatedMemberEmails: relatedMemberEmails
  };
}

function vNextPortalValidateRequestPayload_(payload) {
  vNextPortalAssertExactKeys_(payload, VNEXT_PORTAL.PAYLOAD_KEYS, 'request payload');
  if (payload.schemaVersion !== VNEXT_PORTAL.REQUEST_SCHEMA_VERSION) throw new Error('request schemaVersionが不正です。');
  if (payload.requestType !== VNEXT_PORTAL.REQUEST_TYPE) throw new Error('requestTypeが不正です。');
  if (!/^PORTAL-REQ-[A-Za-z0-9_-]{8,}$/.test(String(payload.requestId || ''))) throw new Error('requestIdが不正です。');
  if (!/^\d{4}-\d{2}-\d{2}T/.test(String(payload.requestedAt || '')) || isNaN(new Date(payload.requestedAt).getTime())) {
    throw new Error('requestedAtが不正です。');
  }
  var normalized = vNextPortalNormalizeCreationInput_({
    fiscalYear: payload.fiscalYear,
    clientName: payload.clientName,
    clientId: payload.clientId,
    forecastOwnerEmail: payload.forecastOwnerEmail,
    relatedMembersText: Array.isArray(payload.relatedMemberEmails) ? payload.relatedMemberEmails.join(',') : payload.relatedMemberEmails
  });
  if (vNextPortalCanonicalJson_(normalized.relatedMemberEmails) !== vNextPortalCanonicalJson_(payload.relatedMemberEmails)) {
    throw new Error('relatedMemberEmailsが正規化されていません。');
  }
  if (normalized.clientName !== payload.clientName || normalized.clientId !== payload.clientId ||
      normalized.forecastOwnerEmail !== payload.forecastOwnerEmail || normalized.fiscalYear !== payload.fiscalYear) {
    throw new Error('request payloadが正規化されていません。');
  }
  if (vNextPortalNormalizeEmail_(payload.requestedBy, true, 'requestedBy') !== payload.requestedBy) {
    throw new Error('requestedByが正規化されていません。');
  }
  return payload;
}

function vNextPortalIsValidStatusEvent_(row, requestId, requestHash) {
  var eventType = String(row && row.event_type || '').trim().toUpperCase();
  var status = String(row && row.status || '').trim().toUpperCase();
  if (!Object.prototype.hasOwnProperty.call(VNEXT_PORTAL.STATUS_EVENT_PAIRS, eventType)) return false;
  if (VNEXT_PORTAL.STATUS_EVENT_PAIRS[eventType] !== status) return false;
  if (String(row.request_id || '').trim() !== String(requestId || '')) return false;
  if (String(row.request_hash || '').trim().toLowerCase() !== String(requestHash || '').trim().toLowerCase()) return false;
  if (eventType !== 'REQUESTED' && String(row.request_json || '').trim() !== '') return false;
  var rawUrl = String(row.related_book_url || '').trim();
  if (rawUrl && !vNextPortalSafeBookUrl_(rawUrl)) return false;
  return true;
}

function vNextPortalAssertExactKeys_(value, expectedKeys, label) {
  if (!value || Object.prototype.toString.call(value) !== '[object Object]') {
    throw new Error((label || 'データ') + 'が不正です。');
  }
  var actual = Object.keys(value).sort();
  var expected = expectedKeys.slice().sort();
  if (vNextPortalCanonicalJson_(actual) !== vNextPortalCanonicalJson_(expected)) {
    throw new Error((label || 'データ') + 'の項目が契約と一致しません。');
  }
}

function vNextPortalNormalizeEmailList_(value, ownerEmail) {
  var parts = Array.isArray(value) ? value : String(value || '').split(/[\s,;、，]+/);
  var unique = {};
  parts.forEach(function (part) {
    var email = vNextPortalNormalizeEmail_(part, false, '関与メンバー');
    if (email && email !== ownerEmail) unique[email] = true;
  });
  var result = Object.keys(unique).sort();
  if (result.length > 50) throw new Error('関与メンバーは50名以内で入力してください。');
  return result;
}

function vNextPortalNormalizeEmail_(value, required, label) {
  var email = String(value || '').trim().toLowerCase();
  if (!email && required) throw new Error((label || 'メールアドレス') + 'を入力してください。');
  if (!email) return '';
  if (email.length > 254 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    throw new Error((label || 'メールアドレス') + 'をメールアドレス形式で入力してください。');
  }
  return email;
}

function vNextPortalPlainText_(value, maxLength, required, label) {
  var text = String(value === undefined || value === null ? '' : value).trim();
  if (required && !text) throw new Error(label + 'を入力してください。');
  if (text.length > maxLength) throw new Error(label + 'は' + maxLength + '文字以内で入力してください。');
  if (/[\u0000-\u001f\u007f]/.test(text)) throw new Error(label + 'に使用できない文字があります。');
  if (/^[=+\-@]/.test(text)) throw new Error(label + 'の先頭文字を変更してください。');
  return text;
}

function vNextPortalNormalizeClientId_(value) {
  return String(value || '').trim().toLowerCase().replace(/[\s\-_.:/\\]/g, '');
}

function vNextPortalNormalizeClientName_(value) {
  var text = String(value || '').trim().toLowerCase();
  try { text = text.normalize('NFKC'); } catch (error) { /* V8 normally supports normalize. */ }
  return text
    .replace(/株式会社|有限会社|合同会社|\(株\)|\(有\)|\(同\)|㈱|㈲/g, '')
    .replace(/[\s\u3000・･.,，。'’"“”\-ー_]/g, '');
}

function vNextPortalNameSimilarity_(left, right) {
  if (!left || !right) return 0;
  if (left === right) return 1;
  var a = vNextPortalBigrams_(left);
  var b = vNextPortalBigrams_(right);
  if (!a.length || !b.length) return left.indexOf(right) >= 0 || right.indexOf(left) >= 0 ? 0.8 : 0;
  var counts = {};
  a.forEach(function (gram) { counts[gram] = (counts[gram] || 0) + 1; });
  var intersection = 0;
  b.forEach(function (gram) {
    if (counts[gram] > 0) { intersection++; counts[gram]--; }
  });
  return (2 * intersection) / (a.length + b.length);
}

function vNextPortalBigrams_(value) {
  var result = [];
  for (var i = 0; i < value.length - 1; i++) result.push(value.slice(i, i + 2));
  return result;
}

function vNextPortalCanonicalJson_(value) {
  if (value === null) return 'null';
  if (Array.isArray(value)) {
    return '[' + value.map(function (item) {
      return item === undefined || typeof item === 'function' || typeof item === 'symbol'
        ? 'null' : vNextPortalCanonicalJson_(item);
    }).join(',') + ']';
  }
  if (value instanceof Date) return JSON.stringify(value.toISOString());
  if (typeof value === 'object') {
    return '{' + Object.keys(value).sort().filter(function (key) {
      return value[key] !== undefined && typeof value[key] !== 'function' && typeof value[key] !== 'symbol';
    }).map(function (key) {
      return JSON.stringify(key) + ':' + vNextPortalCanonicalJson_(value[key]);
    }).join(',') + '}';
  }
  if (typeof value === 'number' && !isFinite(value)) return 'null';
  return JSON.stringify(value);
}

function vNextPortalSha256Hex_(value) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(value), Utilities.Charset.UTF_8);
  return bytes.map(function (byte) {
    var hex = (byte < 0 ? byte + 256 : byte).toString(16);
    return hex.length === 1 ? '0' + hex : hex;
  }).join('');
}

function vNextPortalWithDocumentLock_(operation) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try { return operation(); } finally { lock.releaseLock(); }
}

function vNextPortalActiveUserEmail_() {
  var active = '';
  try { active = Session.getActiveUser().getEmail(); } catch (error) { active = ''; }
  if (!active) {
    try { active = Session.getEffectiveUser().getEmail(); } catch (error2) { active = ''; }
  }
  return active ? String(active).trim().toLowerCase() : '';
}

function vNextPortalCurrentFiscalYear_(date) {
  var value = date instanceof Date ? date : new Date(date);
  return value.getMonth() >= 3 ? value.getFullYear() : value.getFullYear() - 1;
}

function vNextPortalDirectoryStateLabel_(state) {
  var labels = {
    INPUT_OPEN: '情報入力受付中', READY_TO_RUN: '予測依頼待ち', RUNNING: '予測作成中',
    DRAFT_READY: '予測案完成', SUBMITTED: '承認待ち', CHANGES_REQUESTED: '差戻し',
    OFFICIAL_LOCKED: '正式計画', REVIEW_DUE: '振り返り期間', YEAR_CLOSED: '年度終了'
  };
  return labels[String(state || '').toUpperCase()] || String(state || '準備中');
}

function vNextPortalStatusNextAction_(status, detail, hasUrl) {
  var key = String(status || '').toUpperCase();
  if (key === 'COMPLETED' && hasUrl) return '「開く」から年度計画を開始してください。';
  if (key === 'PENDING') return '受付済みです。そのままお待ちください。';
  if (key === 'VALIDATING') return '内容を確認しています。操作は不要です。';
  if (key === 'CREATING') return '専用ブックを作成しています。操作は不要です。';
  if (key === 'FAILED') return detail || '内容を確認して、必要ならもう一度依頼してください。';
  if (key === 'REJECTED') return detail || '表示された理由を確認してください。';
  return detail || '処理状況を確認しています。';
}

function vNextPortalSafeBookUrl_(value) {
  var url = String(value || '').trim();
  return /^https:\/\/docs\.google\.com\/spreadsheets\/d\/[A-Za-z0-9_-]{20,}(?:\/|$)/.test(url) ? url : '';
}

function vNextPortalParseEmailArray_(value) {
  try {
    if (Array.isArray(value)) return vNextPortalNormalizeEmailList_(value, '');
    if (!value) return [];
    var parsed = JSON.parse(String(value));
    return Array.isArray(parsed) ? vNextPortalNormalizeEmailList_(parsed, '') : [];
  } catch (error) {
    vNextPortalLog_('invalid directory member list ignored', error);
    return [];
  }
}

function vNextPortalOptionalNumber_(value) {
  if (value === '' || value === null || value === undefined) return null;
  var number = Number(value);
  return isFinite(number) ? number : null;
}

function vNextPortalIsoText_(value) {
  if (!value) return '';
  var date = value instanceof Date ? value : new Date(value);
  return isNaN(date.getTime()) ? String(value) : date.toISOString();
}

function vNextPortalHideInternalSheet_(sheet) {
  try { sheet.hideSheet(); } catch (error) { vNextPortalLog_('internal sheet hide skipped', error); }
}

function vNextPortalHideBlankDefaultSheets_(spreadsheet) {
  spreadsheet.getSheets().forEach(function (sheet) {
    if (['Sheet1', 'シート1'].indexOf(sheet.getName()) < 0) return;
    if (sheet.getLastRow() > 1 || sheet.getLastColumn() > 1 || sheet.getRange('A1').getValue() !== '') return;
    try { sheet.hideSheet(); } catch (error) { vNextPortalLog_('blank default sheet hide skipped', error); }
  });
}

function vNextPortalProtectWarningOnly_(sheet, description) {
  try {
    var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    var protection = protections.length ? protections[0] : sheet.protect();
    protection.setDescription(description).setWarningOnly(true);
  } catch (error) {
    vNextPortalLog_('warning protection skipped', error);
  }
}

function vNextPortalEnsureSheetSize_(sheet, rows, columns) {
  if (sheet.getMaxRows() < rows) sheet.insertRowsAfter(sheet.getMaxRows(), rows - sheet.getMaxRows());
  if (sheet.getMaxColumns() < columns) sheet.insertColumnsAfter(sheet.getMaxColumns(), columns - sheet.getMaxColumns());
}

function vNextPortalLog_(message, error) {
  Logger.log('[vNext Portal] ' + message + (error ? ': ' + (error.message || String(error)) : ''));
}
