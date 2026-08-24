/**
 * Forecast vNext shared employee portal local core.
 * It records creation requests in the bound Spreadsheet and reads only local directory projections.
 */

var VNEXT_PORTAL_NAMING = Object.freeze({
  SYSTEM: '年度予算策定',
  MENU: '年度予算策定',
  ADMIN_HUB: '管理ハブ',
  PORTAL: '申請入口',
  CLIENT_BOOK: 'クライアント年度ブック',
  LAYER1_SHORT: '第1層：申請入口',
  LAYER2_SHORT: '第2層：クライアント年度ブック',
  LAYER3_SHORT: '第3層：管理ハブ',
  WEB_ENTRY: '年度予算策定 Web入口',
  /** @deprecated Use PORTAL. */
  LAYER2: '申請入口',
  /** @deprecated Use CLIENT_BOOK. */
  LAYER3: 'クライアント年度ブック'
});

var VNEXT_PORTAL = Object.freeze({
  MENU_NAME: VNEXT_PORTAL_NAMING.MENU,
  RUNTIME_VERSION: 'vnext-portal-1.7.15',
  REQUEST_SCHEMA_VERSION: 'vnext-portal-request-2',
  LEGACY_REQUEST_SCHEMA_VERSION: 'vnext-portal-request-1',
  REQUEST_TYPE: 'CREATE_CLIENT_FY_BOOK',
  HOME_SHEET: 'ホーム',
  DIRECTORY_SHEET: 'PORTAL_DIRECTORY',
  REQUEST_SHEET: 'VN_PORTAL_REQUEST',
  CLIENT_CATALOG_SHEET: 'VN_PORTAL_CLIENT_CATALOG',
  CLIENT_CATALOG_CACHE_KEY: 'vnext-portal-client-catalog-v1',
  CLIENT_CATALOG_CACHE_SECONDS: 300,
  ENTRY_CACHE_KEY: 'vnext-portal-entry-model-v1',
  ENTRY_CACHE_SECONDS: 120,
  PAYLOAD_KEYS: Object.freeze([
    'catalogKey', 'clientName', 'fiscalYear', 'relatedMemberNames',
    'requestId', 'requestType', 'requestedAt', 'requestedBy', 'schemaVersion'
  ]),
  LEGACY_PAYLOAD_KEYS: Object.freeze([
    'clientId', 'clientName', 'fiscalYear', 'forecastOwnerEmail', 'relatedMemberEmails',
    'requestId', 'requestType', 'requestedAt', 'requestedBy', 'schemaVersion'
  ]),
  PREVIEW_INPUT_KEYS: Object.freeze([
    'clientKey', 'fiscalYear', 'relatedMemberNames'
  ]),
  SUBMIT_INPUT_KEYS: Object.freeze([
    'clientKey', 'confirmSimilarDuplicates', 'duplicateCheckHash', 'fiscalYear', 'relatedMemberNames'
  ]),
  REQUEST_HEADERS: Object.freeze([
    'request_event_id', 'request_id', 'event_type', 'status', 'request_hash', 'request_json',
    'fiscal_year', 'client_id', 'client_name', 'forecast_owner_email',
    'related_member_emails_json', 'requested_at', 'requested_by', 'related_book_id',
    'related_book_url', 'detail_code', 'detail_message', 'created_at', 'created_by',
    'catalog_key', 'related_member_names_json'
  ]),
  DIRECTORY_HEADERS: Object.freeze([
    'directory_event_id', 'directory_key', 'fiscal_year', 'client_id', 'client_name',
    'forecast_owner_email', 'related_member_emails_json', 'state', 'center_forecast',
    'adopted_forecast', 'final_budget', 'next_action', 'client_book_url', 'request_id',
    'updated_at', 'updated_by', 'related_member_names_json'
  ]),
  CLIENT_CATALOG_HEADERS: Object.freeze([
    'catalog_key', 'client_name', 'is_active', 'catalog_version', 'synced_at'
  ]),
  ACTIVE_REQUEST_STATUSES: Object.freeze(['PENDING', 'VALIDATING', 'CREATING', 'COMPLETED']),
  REQUEST_STATUS_LABELS: Object.freeze({
    PENDING: '受付済み',
    VALIDATING: '内容確認中',
    CREATING: 'クライアント年度ブック作成中',
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

/** Returns a read-only, cached model for the creation sidebar. */
function vNextPortalGetCreateModel() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var fiscalYear = vNextPortalCurrentFiscalYear_(new Date());
    var actor = vNextPortalActiveUserEmail_();
    if (!actor) throw new Error('ログイン中のGoogleアカウントを確認できません。社内アカウントで開き直してください。');
    var catalog = vNextPortalReadClientCatalog_(spreadsheet, true);
    var fiscalYears = [];
    for (var year = fiscalYear; year <= fiscalYear + 10; year++) fiscalYears.push(year);
    return {
      ok: true,
      runtimeVersion: VNEXT_PORTAL.RUNTIME_VERSION,
      defaultFiscalYear: fiscalYear + 1,
      fiscalYears: fiscalYears,
      requesterEmail: actor,
      clients: catalog.clients,
      catalogVersion: catalog.version,
      catalogSyncedAt: catalog.syncedAt,
      instructions: 'ZACのクライアントから選び、関与メンバーを入力してください。作成担当はログイン中のあなたに自動設定されます。'
    };
  } catch (error) {
    vNextPortalLog_('vNextPortalGetCreateModel failed', error);
    throw new Error(error && error.message ? error.message : '作成画面を準備できませんでした。');
  }
}

/** Finds same-FY duplicate candidates from the local directory and pending requests. */
function vNextPortalPreviewCreation(input) {
  try {
    vNextPortalAssertExactKeys_(input, VNEXT_PORTAL.PREVIEW_INPUT_KEYS, '入力内容');
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var actor = vNextPortalActiveUserEmail_();
    if (!actor) throw new Error('ログイン中のGoogleアカウントを確認できません。社内アカウントで開き直してください。');
    // Preview is a deliberate action, so use the current protected catalog.
    // Only the initial sidebar model uses the short-lived cache.
    var normalized = vNextPortalResolveCreationInput_(input, spreadsheet, actor, false);
    var check = vNextPortalBuildDuplicateCheck_(spreadsheet, normalized);
    return {
      ok: true,
      normalized: {
        fiscalYear: normalized.fiscalYear,
        clientName: normalized.clientName,
        catalogKey: normalized.catalogKey,
        relatedMemberNames: normalized.relatedMemberNames
      },
      candidates: check.candidates,
      hasExact: check.hasExact,
      hasSimilar: check.hasSimilar,
      canSubmit: !check.hasExact,
      duplicateCheckHash: check.hash,
      message: check.hasExact
        ? '同じ年度のクライアント年度ブックまたは作成依頼が見つかりました。新しく作らず、既存のクライアント年度ブックを確認してください。'
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
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    return vNextPortalWithDocumentLock_(function () {
      var actor = vNextPortalActiveUserEmail_();
      if (!actor) throw new Error('ログイン中のGoogleアカウントを確認できません。社内アカウントで開き直してください。');
      // Submission always bypasses the short-lived catalog cache so a stale or
      // browser-modified selection cannot be persisted.
      var normalized = vNextPortalResolveCreationInput_(input, spreadsheet, actor, false);
      var check = vNextPortalBuildDuplicateCheck_(spreadsheet, normalized);
      if (check.hash !== duplicateCheckHash) {
        throw new Error('候補一覧が更新されました。もう一度「重複候補を確認」を押してください。');
      }
      if (check.hasExact) {
        throw new Error('同じ年度のクライアント年度ブックまたは作成依頼があります。既存のクライアント年度ブックを利用してください。');
      }
      if (check.hasSimilar && input.confirmSimilarDuplicates !== true) {
        throw new Error('似た名前の候補を確認し、「別のクライアントです」にチェックしてください。');
      }
      var now = new Date().toISOString();
      var requestId = 'PORTAL-REQ-' + Utilities.getUuid();
      var payload = {
        catalogKey: normalized.catalogKey,
        clientName: normalized.clientName,
        fiscalYear: normalized.fiscalYear,
        relatedMemberNames: normalized.relatedMemberNames,
        requestId: requestId,
        requestType: VNEXT_PORTAL.REQUEST_TYPE,
        requestedAt: now,
        requestedBy: actor,
        schemaVersion: VNEXT_PORTAL.REQUEST_SCHEMA_VERSION
      };
      vNextPortalValidateRequestPayload_(payload);
      var requestJson = vNextPortalCanonicalJson_(payload);
      var requestHash = vNextPortalSha256Hex_(requestJson);
      vNextPortalClearEntryCache_();
      vNextPortalAppendRow_(spreadsheet, VNEXT_PORTAL.REQUEST_SHEET, VNEXT_PORTAL.REQUEST_HEADERS, {
        request_event_id: 'PORTAL-REQEV-' + Utilities.getUuid(),
        request_id: requestId,
        event_type: 'REQUESTED',
        status: 'PENDING',
        request_hash: requestHash,
        request_json: requestJson,
        fiscal_year: normalized.fiscalYear,
        client_id: normalized.catalogKey,
        client_name: normalized.clientName,
        forecast_owner_email: actor,
        related_member_emails_json: '[]',
        requested_at: now,
        requested_by: actor,
        related_book_id: '',
        related_book_url: '',
        detail_code: '',
        detail_message: '作成依頼を受け付けました。',
        created_at: now,
        created_by: actor,
        catalog_key: normalized.catalogKey,
        related_member_names_json: vNextPortalCanonicalJson_(normalized.relatedMemberNames)
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

/**
 * Lightweight request-status endpoint for the creation sidebar.
 * It reads the append-only request log only; it never rebuilds visible sheets.
 */
function vNextPortalGetRequestProgress(requestId) {
  try {
    var id = vNextPortalNormalizeRequestId_(requestId);
    var requests = vNextPortalReadRequestModels_(SpreadsheetApp.getActiveSpreadsheet());
    var request = null;
    requests.some(function (item) {
      if (item.requestId !== id) return false;
      request = item;
      return true;
    });
    if (!request) throw new Error('受付状況を確認できませんでした。ホームを更新して確認してください。');
    return vNextPortalRequestProgressModel_(request);
  } catch (error) {
    vNextPortalLog_('vNextPortalGetRequestProgress failed', error);
    throw error;
  }
}

function vNextPortalRequestProgressModel_(request) {
  var status = String(request && request.status || 'PENDING').trim().toUpperCase();
  var terminal = ['COMPLETED', 'FAILED', 'REJECTED'].indexOf(status) >= 0;
  var phaseByStatus = { PENDING: 1, VALIDATING: 2, REJECTED: 2, CREATING: 3, FAILED: 3, COMPLETED: 4 };
  var updatedAt = vNextPortalIsoText_(request && request.updatedAt || request && request.requestedAt);
  var updatedMs = new Date(updatedAt || '').getTime();
  var stale = !terminal && isFinite(updatedMs) && Date.now() - updatedMs >= 15 * 60 * 1000;
  var nextAction = vNextPortalStatusNextAction_(status, request && request.detailMessage, Boolean(request && request.url));
  if (stale) {
    nextAction = '15分以上状態が変わっていません。受付番号を添えて管理担当者へ連絡してください。';
  }
  var waitMessage = '';
  if (status === 'PENDING') {
    waitMessage = '管理システムは5分ごとに受付を確認します。通常は次回の自動処理で開始します。';
  } else if (status === 'VALIDATING') {
    waitMessage = '入力内容を確認しています。ここで追加操作は必要ありません。';
  } else if (status === 'CREATING') {
    waitMessage = 'クライアント年度ブックを作成しています。通常は数分で完了します。';
  } else if (status === 'COMPLETED') {
    waitMessage = 'クライアント年度ブックを利用できます。';
  } else {
    waitMessage = '自動処理は停止しています。表示された案内を確認してください。';
  }
  return {
    ok: true,
    requestId: String(request.requestId || ''),
    clientName: String(request.clientName || ''),
    fiscalYear: Number(request.fiscalYear || 0),
    status: status,
    statusLabel: VNEXT_PORTAL.REQUEST_STATUS_LABELS[status] || '処理状況を確認中',
    nextAction: nextAction,
    waitMessage: waitMessage,
    updatedAt: updatedAt,
    url: vNextPortalSafeBookUrl_(request.url),
    phase: phaseByStatus[status] || 1,
    phaseCount: 4,
    isTerminal: terminal,
    isSuccess: status === 'COMPLETED',
    isStale: stale
  };
}

function vNextPortalNormalizeRequestId_(value) {
  var requestId = String(value || '').trim();
  if (!/^PORTAL-REQ-[A-Za-z0-9_-]{8,}$/.test(requestId)) throw new Error('受付番号が不正です。');
  return requestId;
}

/** Employee launcher used by the published Web App. */
function vNextPortalGetEntryModel() {
  try {
    vNextPortalPrepareOpenExperienceThrottled_();
    return vNextPortalBuildEntryModelCached_();
  } catch (error) {
    vNextPortalLog_('vNextPortalGetEntryModel failed', error);
    throw new Error(error && error.message ? error.message : '入口を準備できませんでした。');
  }
}

function vNextPortalBuildEntryModelCached_() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var cached = vNextPortalReadEntryCache_();
  if (cached) return cached;
  var data = vNextPortalGetLocalViewData_(spreadsheet);
  var config = vNextPortalReadConfig_(spreadsheet);
  var model = vNextPortalBuildEntryModel_(data, {
    portalUrl: spreadsheet.getUrl(),
    adminHubUrl: vNextPortalSafeSpreadsheetUrl_(config.admin_hub_url)
  });
  vNextPortalWriteEntryCache_(model);
  return model;
}

/** JSON payload for the doGet server-side render. Returns 'null' on failure. */
function vNextPortalEntryModelJsonForTemplate_() {
  try {
    vNextPortalPrepareOpenExperienceThrottled_();
    return JSON.stringify(vNextPortalBuildEntryModelCached_()).replace(/</g, '\\u003c');
  } catch (error) {
    vNextPortalLog_('entry SSR skipped', error);
    return 'null';
  }
}

/**
 * The onOpen-trigger check hits ScriptApp on every call; entry loads are hot,
 * so ensure it at most once per 6 hours.
 */
function vNextPortalPrepareOpenExperienceThrottled_() {
  try {
    var cache = typeof CacheService !== 'undefined' ? CacheService.getScriptCache() : null;
    if (cache && cache.get('vnext-portal-open-prepared')) return;
    vNextPortalPrepareOpenExperience();
    if (cache) cache.put('vnext-portal-open-prepared', '1', 21600);
  } catch (prepareError) {
    vNextPortalLog_('vNextPortalPrepareOpenExperience skipped', prepareError);
  }
}

function vNextPortalBuildEntryModel_(data, options) {
  var opt = options && typeof options === 'object' ? options : {};
  var directory = (data && data.directory) || [];
  var requests = (data && data.requests) || [];
  var seen = {};
  var books = [];
  directory.forEach(function (item) {
    var key = String(item.directoryKey || item.clientName || '') + '|FY' + String(item.fiscalYear || '');
    seen[key] = true;
    books.push(vNextPortalEntryBookFromDirectory_(item));
  });
  requests.forEach(function (item) {
    var key = String(item.clientName || '') + '|FY' + String(item.fiscalYear || '');
    if (seen[key] && item.url) return;
    if (seen[key] && item.status === 'COMPLETED') return;
    if (!seen[key] || !item.url) {
      seen[key] = true;
      books.push(vNextPortalEntryBookFromRequest_(item));
    }
  });
  books.sort(function (a, b) {
    if (Number(a.fiscalYear) !== Number(b.fiscalYear)) return Number(b.fiscalYear) - Number(a.fiscalYear);
    return String(a.clientName).localeCompare(String(b.clientName), 'ja');
  });
  var years = [];
  books.forEach(function (book) {
    var year = Number(book.fiscalYear || 0);
    if (year && years.indexOf(year) < 0) years.push(year);
  });
  years.sort(function (a, b) { return b - a; });
  return {
    ok: true,
    runtimeVersion: VNEXT_PORTAL.RUNTIME_VERSION,
    portalUrl: String(opt.portalUrl || ''),
    adminHubUrl: String(opt.adminHubUrl || ''),
    years: years,
    books: books
  };
}

function vNextPortalReadEntryCache_() {
  try {
    if (typeof CacheService === 'undefined') return null;
    var cache = CacheService.getDocumentCache();
    var raw = cache && cache.get(VNEXT_PORTAL.ENTRY_CACHE_KEY);
    if (!raw) return null;
    var parsed = JSON.parse(raw);
    if (!parsed || parsed.ok !== true || !Array.isArray(parsed.books) || !Array.isArray(parsed.years)) return null;
    return parsed;
  } catch (cacheReadError) {
    return null;
  }
}

function vNextPortalWriteEntryCache_(model) {
  try {
    if (typeof CacheService === 'undefined' || !model) return;
    var cache = CacheService.getDocumentCache();
    if (!cache) return;
    cache.put(VNEXT_PORTAL.ENTRY_CACHE_KEY, JSON.stringify({
      ok: true,
      runtimeVersion: model.runtimeVersion,
      portalUrl: model.portalUrl,
      adminHubUrl: model.adminHubUrl,
      years: model.years || [],
      books: model.books || []
    }), VNEXT_PORTAL.ENTRY_CACHE_SECONDS);
  } catch (cacheWriteError) {
    vNextPortalLog_('entry cache write skipped', cacheWriteError);
  }
}

function vNextPortalClearEntryCache_() {
  try {
    if (typeof CacheService === 'undefined') return;
    var cache = CacheService.getDocumentCache();
    if (cache) cache.remove(VNEXT_PORTAL.ENTRY_CACHE_KEY);
  } catch (cacheClearError) {
    vNextPortalLog_('entry cache clear skipped', cacheClearError);
  }
}

function vNextPortalEntryBookFromDirectory_(item) {
  var state = String(item.state || '').toUpperCase();
  return {
    clientName: String(item.clientName || ''),
    fiscalYear: Number(item.fiscalYear || 0),
    state: state,
    stateLabel: vNextPortalDirectoryStateLabel_(state),
    nextAction: String(item.nextAction || ''),
    url: String(item.url || ''),
    tone: vNextPortalEntryTone_(state, false)
  };
}

function vNextPortalEntryBookFromRequest_(item) {
  var status = String(item.status || '').toUpperCase();
  return {
    clientName: String(item.clientName || ''),
    fiscalYear: Number(item.fiscalYear || 0),
    state: status,
    stateLabel: VNEXT_PORTAL.REQUEST_STATUS_LABELS[status] || status,
    nextAction: vNextPortalStatusNextAction_(status, item.detailMessage, Boolean(item.url)),
    url: String(item.url || ''),
    tone: vNextPortalEntryTone_(status, true)
  };
}

function vNextPortalEntryTone_(state, isRequest) {
  var key = String(state || '').toUpperCase();
  if (key === 'OFFICIAL_LOCKED' || key === 'COMPLETED') return 'good';
  if (key === 'SUBMITTED' || key === 'CHANGES_REQUESTED' || key === 'FAILED' || key === 'REJECTED') return 'warn';
  if (isRequest && (key === 'PENDING' || key === 'VALIDATING' || key === 'CREATING')) return 'warn';
  return '';
}

function vNextPortalReadConfig_(spreadsheet) {
  var sheet = spreadsheet && spreadsheet.getSheetByName('VN_PORTAL_CONFIG');
  var out = {};
  if (!sheet || sheet.getLastRow() < 2) return out;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(function (row) {
    var key = String(row[0] || '').trim();
    if (key) out[key] = row[1];
  });
  return out;
}

function vNextPortalSafeSpreadsheetUrl_(value) {
  var url = String(value || '').trim();
  return /^https:\/\/docs\.google\.com\/spreadsheets\/d\/[A-Za-z0-9_-]{20,}(?:\/|$)/.test(url) ? url : '';
}

function vNextPortalGetLocalViewData_(spreadsheet) {
  spreadsheet = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
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
  var catalog = vNextPortalEnsureTable_(ss, VNEXT_PORTAL.CLIENT_CATALOG_SHEET, VNEXT_PORTAL.CLIENT_CATALOG_HEADERS);
  vNextPortalHideInternalSheet_(directory);
  vNextPortalHideInternalSheet_(requests);
  vNextPortalHideInternalSheet_(catalog);
  vNextPortalHideBlankDefaultSheets_(ss);
  return { home: home, directory: directory, requests: requests };
}

function vNextPortalEnsureFiscalYearSheet_(spreadsheet, fiscalYear) {
  var name = 'FY' + Number(fiscalYear);
  var sheet = spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
  sheet.showSheet();
  sheet.setTabColor(null);
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
  var sheet = vNextPortalRequireTable_(spreadsheet, sheetName, headers);
  var values = headers.map(function (header) {
    var value = record[header];
    return value === undefined || value === null ? '' : value;
  });
  if (sheet.getLastRow() + 1 > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), 1);
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, values.length).setValues([values]);
  return record;
}

function vNextPortalReadTable_(spreadsheet, sheetName, headers) {
  var sheet = vNextPortalRequireTable_(spreadsheet, sheetName, headers);
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

/** Read-path table assertion. It never creates, formats, protects, or hides a sheet. */
function vNextPortalRequireTable_(spreadsheet, sheetName, headers) {
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) throw new Error(sheetName + 'が準備されていません。管理担当者へ連絡してください。');
  var actual = sheet.getRange(1, 1, 1, headers.length).getValues()[0].map(function (value) {
    return String(value || '');
  });
  if (vNextPortalCanonicalJson_(actual) !== vNextPortalCanonicalJson_(headers.slice())) {
    throw new Error(sheetName + 'の列構成が正しくありません。管理担当者へ連絡してください。');
  }
  return sheet;
}

/** Reads the Admin-synchronized ZAC client catalog, with a short read-only cache. */
function vNextPortalReadClientCatalog_(spreadsheet, allowCache) {
  var cache = null;
  if (allowCache && typeof CacheService !== 'undefined') {
    try {
      cache = CacheService.getDocumentCache();
      var cached = cache && cache.get(VNEXT_PORTAL.CLIENT_CATALOG_CACHE_KEY);
      if (cached) return vNextPortalValidateCatalogModel_(JSON.parse(cached));
    } catch (cacheReadError) {
      vNextPortalLog_('client catalog cache read skipped', cacheReadError);
    }
  }

  var rows = vNextPortalReadTable_(
    spreadsheet, VNEXT_PORTAL.CLIENT_CATALOG_SHEET, VNEXT_PORTAL.CLIENT_CATALOG_HEADERS
  );
  var version = '';
  var syncedAt = '';
  var seenKeys = {};
  var clients = [];
  rows.forEach(function (row) {
    var rowVersion = vNextPortalPlainText_(row.catalog_version, 100, true, 'クライアント一覧の版');
    if (!version) version = rowVersion;
    if (version !== rowVersion) throw new Error('クライアント一覧の版が混在しています。管理担当者へ連絡してください。');
    var rowSyncedAt = vNextPortalIsoText_(row.synced_at);
    if (!rowSyncedAt || isNaN(new Date(rowSyncedAt).getTime())) {
      throw new Error('クライアント一覧の更新日時が不正です。管理担当者へ連絡してください。');
    }
    if (!syncedAt || rowSyncedAt > syncedAt) syncedAt = rowSyncedAt;
    var key = vNextPortalPlainText_(row.catalog_key, 100, true, 'クライアント識別子');
    if (seenKeys[key]) throw new Error('クライアント一覧に重複があります。管理担当者へ連絡してください。');
    seenKeys[key] = true;
    var name = vNextPortalPlainText_(row.client_name, 120, true, 'クライアント名');
    if (vNextPortalCatalogActive_(row.is_active)) clients.push({ key: key, name: name });
  });
  if (!version || !clients.length) {
    throw new Error('ZACのクライアント一覧がまだ準備されていません。管理担当者へ連絡してください。');
  }
  clients.sort(function (a, b) { return String(a.name).localeCompare(String(b.name), 'ja'); });
  var model = vNextPortalValidateCatalogModel_({ version: version, syncedAt: syncedAt, clients: clients });
  if (allowCache && cache) {
    try {
      cache.put(VNEXT_PORTAL.CLIENT_CATALOG_CACHE_KEY, JSON.stringify(model), VNEXT_PORTAL.CLIENT_CATALOG_CACHE_SECONDS);
    } catch (cacheWriteError) {
      vNextPortalLog_('client catalog cache write skipped', cacheWriteError);
    }
  }
  return model;
}

function vNextPortalValidateCatalogModel_(model) {
  if (!model || Object.prototype.toString.call(model) !== '[object Object]') throw new Error('クライアント一覧が不正です。');
  var version = vNextPortalPlainText_(model.version, 100, true, 'クライアント一覧の版');
  var syncedAt = vNextPortalIsoText_(model.syncedAt);
  if (!syncedAt || isNaN(new Date(syncedAt).getTime())) throw new Error('クライアント一覧の更新日時が不正です。');
  if (!Array.isArray(model.clients) || !model.clients.length) throw new Error('ZACのクライアント一覧が空です。');
  var seen = {};
  var clients = model.clients.map(function (client) {
    vNextPortalAssertExactKeys_(client, ['key', 'name'], 'クライアント候補');
    var key = vNextPortalPlainText_(client.key, 100, true, 'クライアント識別子');
    var name = vNextPortalPlainText_(client.name, 120, true, 'クライアント名');
    if (seen[key]) throw new Error('クライアント一覧に重複があります。');
    seen[key] = true;
    return { key: key, name: name };
  });
  return { version: version, syncedAt: syncedAt, clients: clients };
}

function vNextPortalCatalogActive_(value) {
  if (value === true || value === 1) return true;
  if (value === false || value === 0) return false;
  var normalized = String(value === undefined || value === null ? '' : value).trim().toUpperCase();
  if (['TRUE', '1', 'ACTIVE', 'YES'].indexOf(normalized) >= 0) return true;
  if (['FALSE', '0', 'INACTIVE', 'NO'].indexOf(normalized) >= 0) return false;
  throw new Error('クライアント一覧の有効フラグが不正です。管理担当者へ連絡してください。');
}

function vNextPortalResolveCreationInput_(input, spreadsheet, actor, allowCatalogCache) {
  var fiscalYear = vNextPortalNormalizeCreationFiscalYear_(input.fiscalYear, new Date());
  var clientKey = vNextPortalPlainText_(input.clientKey, 100, true, 'クライアント');
  var relatedMemberNames = vNextPortalNormalizeMemberNames_(input.relatedMemberNames, true);
  var catalog = vNextPortalReadClientCatalog_(spreadsheet, allowCatalogCache === true);
  var selected = null;
  catalog.clients.some(function (client) {
    if (client.key !== clientKey) return false;
    selected = client;
    return true;
  });
  if (!selected) throw new Error('選択したクライアントは現在のZAC一覧にありません。一覧を開き直してください。');
  return {
    catalogKey: selected.key,
    clientId: selected.key,
    clientName: selected.name,
    fiscalYear: fiscalYear,
    relatedMemberNames: relatedMemberNames,
    requestedBy: vNextPortalNormalizeEmail_(actor, true, '作成担当')
  };
}

function vNextPortalNormalizeCreationFiscalYear_(value, now) {
  var fiscalYear = Number(value);
  var current = vNextPortalCurrentFiscalYear_(now);
  if (!isFinite(fiscalYear) || Math.floor(fiscalYear) !== fiscalYear || fiscalYear < current || fiscalYear > current + 10) {
    throw new Error('対象年度を一覧から正しく選択してください。');
  }
  return fiscalYear;
}

function vNextPortalNormalizeStoredFiscalYear_(value) {
  var fiscalYear = Number(value);
  if (!isFinite(fiscalYear) || Math.floor(fiscalYear) !== fiscalYear || fiscalYear < 2000 || fiscalYear > 2200) {
    throw new Error('対象年度が不正です。');
  }
  return fiscalYear;
}

function vNextPortalNormalizeMemberNames_(value, required) {
  if (!Array.isArray(value)) throw new Error('関与メンバーの入力形式が不正です。');
  if (value.length > 5) throw new Error('関与メンバーは5名以内で入力してください。');
  var seen = {};
  var names = [];
  value.forEach(function (item) {
    var name = vNextPortalPlainText_(item, 80, false, '関与メンバー');
    if (!name) return;
    var key = name;
    try { key = key.normalize('NFKC'); } catch (ignoredNormalize) {}
    key = key.toLowerCase().replace(/[\s\u3000]+/g, '');
    if (seen[key]) throw new Error('同じ関与メンバーが複数の欄に入力されています。');
    seen[key] = true;
    names.push(name);
  });
  if (required && !names.length) throw new Error('関与メンバーを1名以上入力してください。');
  if (names.length > 5) throw new Error('関与メンバーは5名以内で入力してください。');
  return names;
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
      relatedMemberNames: vNextPortalParseMemberNames_(row.related_member_names_json),
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
        vNextPortalAssertRequestRowProjection_(entry.row, payload);
        requested = { row: entry.row, payload: payload, model: vNextPortalRequestPayloadToModel_(payload) };
        return true;
      } catch (error) {
        vNextPortalLog_('invalid local request ignored: ' + requestId, error);
        return false;
      }
    });
    if (!requested) return null;
    var latest = requested.row;
    events.forEach(function (entry) {
      if (vNextPortalIsValidStatusEvent_(
        entry.row, requestId, requested.row.request_hash, requested.payload
      )) latest = entry.row;
    });
    var status = String(latest.status || requested.row.status || 'PENDING').trim().toUpperCase();
    return {
      source: 'REQUEST',
      requestId: requestId,
      requestHash: String(requested.row.request_hash || '').toLowerCase(),
      fiscalYear: requested.model.fiscalYear,
      clientId: requested.model.clientId,
      catalogKey: requested.model.catalogKey,
      clientName: requested.model.clientName,
      forecastOwnerEmail: requested.model.forecastOwnerEmail,
      relatedMemberEmails: requested.model.relatedMemberEmails,
      relatedMemberNames: requested.model.relatedMemberNames,
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
  var normalizedId = vNextPortalNormalizeClientId_(normalized.catalogKey || normalized.clientId);
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
    catalogKey: normalized.catalogKey,
    clientName: normalizedName,
    fiscalYear: normalized.fiscalYear,
    relatedMemberNames: normalized.relatedMemberNames,
    requestedBy: normalized.requestedBy,
    schemaVersion: VNEXT_PORTAL.REQUEST_SCHEMA_VERSION
  };
  return {
    candidates: candidates,
    hasExact: candidates.some(function (item) { return item.level === 'EXACT'; }),
    hasSimilar: candidates.some(function (item) { return item.level === 'SIMILAR'; }),
    hash: vNextPortalSha256Hex_(vNextPortalCanonicalJson_(hashInput))
  };
}

function vNextPortalValidateRequestPayload_(payload) {
  var schemaVersion = String(payload && payload.schemaVersion || '');
  if (schemaVersion === VNEXT_PORTAL.LEGACY_REQUEST_SCHEMA_VERSION) {
    return vNextPortalValidateLegacyRequestPayload_(payload);
  }
  if (schemaVersion !== VNEXT_PORTAL.REQUEST_SCHEMA_VERSION) throw new Error('request schemaVersionが不正です。');
  vNextPortalAssertExactKeys_(payload, VNEXT_PORTAL.PAYLOAD_KEYS, 'request payload');
  if (payload.requestType !== VNEXT_PORTAL.REQUEST_TYPE) throw new Error('requestTypeが不正です。');
  if (!/^PORTAL-REQ-[A-Za-z0-9_-]{8,}$/.test(String(payload.requestId || ''))) throw new Error('requestIdが不正です。');
  if (!/^\d{4}-\d{2}-\d{2}T/.test(String(payload.requestedAt || '')) || isNaN(new Date(payload.requestedAt).getTime())) {
    throw new Error('requestedAtが不正です。');
  }
  var catalogKey = vNextPortalPlainText_(payload.catalogKey, 100, true, 'クライアント識別子');
  var clientName = vNextPortalPlainText_(payload.clientName, 120, true, 'クライアント名');
  var fiscalYear = vNextPortalNormalizeStoredFiscalYear_(payload.fiscalYear);
  var memberNames = vNextPortalNormalizeMemberNames_(payload.relatedMemberNames, true);
  if (vNextPortalCanonicalJson_(memberNames) !== vNextPortalCanonicalJson_(payload.relatedMemberNames) ||
      catalogKey !== payload.catalogKey || clientName !== payload.clientName || fiscalYear !== payload.fiscalYear) {
    throw new Error('request payloadが正規化されていません。');
  }
  if (vNextPortalNormalizeEmail_(payload.requestedBy, true, 'requestedBy') !== payload.requestedBy) {
    throw new Error('requestedByが正規化されていません。');
  }
  return payload;
}

function vNextPortalValidateLegacyRequestPayload_(payload) {
  vNextPortalAssertExactKeys_(payload, VNEXT_PORTAL.LEGACY_PAYLOAD_KEYS, 'legacy request payload');
  if (payload.requestType !== VNEXT_PORTAL.REQUEST_TYPE) throw new Error('requestTypeが不正です。');
  if (!/^PORTAL-REQ-[A-Za-z0-9_-]{8,}$/.test(String(payload.requestId || ''))) throw new Error('requestIdが不正です。');
  if (!/^\d{4}-\d{2}-\d{2}T/.test(String(payload.requestedAt || '')) || isNaN(new Date(payload.requestedAt).getTime())) {
    throw new Error('requestedAtが不正です。');
  }
  var fiscalYear = vNextPortalNormalizeStoredFiscalYear_(payload.fiscalYear);
  var clientName = vNextPortalPlainText_(payload.clientName, 120, true, 'クライアント名');
  var clientId = vNextPortalPlainText_(payload.clientId, 100, false, 'クライアントID');
  var owner = vNextPortalNormalizeEmail_(payload.forecastOwnerEmail, true, '予算策定担当');
  var members = vNextPortalNormalizeEmailList_(payload.relatedMemberEmails, owner);
  if (vNextPortalCanonicalJson_(members) !== vNextPortalCanonicalJson_(payload.relatedMemberEmails) ||
      fiscalYear !== payload.fiscalYear || clientName !== payload.clientName || clientId !== payload.clientId ||
      owner !== payload.forecastOwnerEmail) {
    throw new Error('legacy request payloadが正規化されていません。');
  }
  if (vNextPortalNormalizeEmail_(payload.requestedBy, true, 'requestedBy') !== payload.requestedBy) {
    throw new Error('requestedByが正規化されていません。');
  }
  return payload;
}

function vNextPortalRequestPayloadToModel_(payload) {
  if (payload.schemaVersion === VNEXT_PORTAL.LEGACY_REQUEST_SCHEMA_VERSION) {
    return {
      fiscalYear: Number(payload.fiscalYear), clientId: String(payload.clientId || ''), catalogKey: '',
      clientName: String(payload.clientName || ''), forecastOwnerEmail: String(payload.forecastOwnerEmail || ''),
      relatedMemberEmails: payload.relatedMemberEmails.slice(), relatedMemberNames: []
    };
  }
  return {
    fiscalYear: Number(payload.fiscalYear), clientId: String(payload.catalogKey || ''),
    catalogKey: String(payload.catalogKey || ''), clientName: String(payload.clientName || ''),
    forecastOwnerEmail: String(payload.requestedBy || ''), relatedMemberEmails: [],
    relatedMemberNames: payload.relatedMemberNames.slice()
  };
}

function vNextPortalAssertRequestRowProjection_(row, payload) {
  var expectedClientId = payload.schemaVersion === VNEXT_PORTAL.REQUEST_SCHEMA_VERSION
    ? String(payload.catalogKey || '') : String(payload.clientId || '');
  var expectedOwner = payload.schemaVersion === VNEXT_PORTAL.REQUEST_SCHEMA_VERSION
    ? String(payload.requestedBy || '') : String(payload.forecastOwnerEmail || '');
  var expectedEmails = payload.schemaVersion === VNEXT_PORTAL.REQUEST_SCHEMA_VERSION
    ? '[]' : vNextPortalCanonicalJson_(payload.relatedMemberEmails || []);
  var expectedCatalogKey = payload.schemaVersion === VNEXT_PORTAL.REQUEST_SCHEMA_VERSION
    ? String(payload.catalogKey || '') : '';
  var expectedNames = payload.schemaVersion === VNEXT_PORTAL.REQUEST_SCHEMA_VERSION
    ? vNextPortalCanonicalJson_(payload.relatedMemberNames || []) : '';
  if (Number(row.fiscal_year) !== Number(payload.fiscalYear) ||
      String(row.client_id || '') !== expectedClientId || String(row.client_name || '') !== String(payload.clientName || '') ||
      String(row.forecast_owner_email || '') !== expectedOwner ||
      String(row.related_member_emails_json || '') !== expectedEmails ||
      String(row.requested_at || '') !== String(payload.requestedAt || '') ||
      String(row.requested_by || '') !== String(payload.requestedBy || '') ||
      String(row.catalog_key || '') !== expectedCatalogKey ||
      String(row.related_member_names_json || '') !== expectedNames) {
    throw new Error('request row projection mismatch');
  }
  return true;
}

function vNextPortalIsValidStatusEvent_(row, requestId, requestHash, payload) {
  var eventType = String(row && row.event_type || '').trim().toUpperCase();
  var status = String(row && row.status || '').trim().toUpperCase();
  if (!Object.prototype.hasOwnProperty.call(VNEXT_PORTAL.STATUS_EVENT_PAIRS, eventType)) return false;
  if (VNEXT_PORTAL.STATUS_EVENT_PAIRS[eventType] !== status) return false;
  if (String(row.request_id || '').trim() !== String(requestId || '')) return false;
  if (String(row.request_hash || '').trim().toLowerCase() !== String(requestHash || '').trim().toLowerCase()) return false;
  if (eventType !== 'REQUESTED' && String(row.request_json || '').trim() !== '') return false;
  var rawUrl = String(row.related_book_url || '').trim();
  if (rawUrl && !vNextPortalSafeBookUrl_(rawUrl)) return false;
  if (payload && eventType !== 'REQUESTED') {
    try { vNextPortalAssertRequestRowProjection_(row, payload); }
    catch (projectionError) { return false; }
  }
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
    OFFICIAL_LOCKED: '正式予算', REVIEW_DUE: '振り返り期間', YEAR_CLOSED: '年度終了'
  };
  return labels[String(state || '').toUpperCase()] || String(state || '準備中');
}

function vNextPortalStatusNextAction_(status, detail, hasUrl) {
  var key = String(status || '').toUpperCase();
  if (key === 'COMPLETED' && hasUrl) return '「開く」からクライアント年度ブックで予算作成を開始してください。';
  if (key === 'PENDING') return '受付済みです。自動処理は5分ごとに開始します。';
  if (key === 'VALIDATING') return '内容を確認しています。操作は不要です。';
  if (key === 'CREATING') return 'クライアント年度ブックを作成しています。操作は不要です。';
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

function vNextPortalParseMemberNames_(value) {
  try {
    if (!value) return [];
    var parsed = Array.isArray(value) ? value : JSON.parse(String(value));
    return Array.isArray(parsed) ? vNextPortalNormalizeMemberNames_(parsed, false) : [];
  } catch (error) {
    vNextPortalLog_('invalid directory member names ignored', error);
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
