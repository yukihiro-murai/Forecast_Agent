/**
 * Creates a clean Spreadsheet and binds the generated client-only Apps Script runtime.
 * This file belongs only to the Admin/forecast-service project and is never bundled to clients.
 */

var VNEXT_CLIENT_RUNTIME_API_ROOT_ = 'https://script.googleapis.com/v1';
var VNEXT_CLIENT_RUNTIME_FILE_TYPES_ = Object.freeze({
  Client_Bridge: 'SERVER_JS',
  Client_Core: 'SERVER_JS',
  Client_Entry: 'SERVER_JS',
  VNext_HelpSidebar: 'HTML',
  VNext_InputSidebar: 'HTML',
  VNext_PlanSidebar: 'HTML',
  VNext_ReviewSidebar: 'HTML',
  VNext_UX: 'SERVER_JS',
  appsscript: 'JSON'
});
var VNEXT_CLIENT_RUNTIME_OAUTH_SCOPES_ = Object.freeze([
  'https://www.googleapis.com/auth/script.container.ui',
  'https://www.googleapis.com/auth/spreadsheets.currentonly',
  'https://www.googleapis.com/auth/userinfo.email'
]);
var VNEXT_ADMIN_RUNTIME_FILE_TYPES_ = Object.freeze({
  Forecast_Agent: 'SERVER_JS',
  VNext_AI: 'SERVER_JS',
  VNext_Admin: 'SERVER_JS',
  VNext_AdminSidebar: 'HTML',
  VNext_ClientRuntimeBundle: 'SERVER_JS',
  VNext_ClientRuntimeProvisioning: 'SERVER_JS',
  VNext_Core: 'SERVER_JS',
  VNext_Engine: 'SERVER_JS',
  VNext_HelpSidebar: 'HTML',
  VNext_InputSidebar: 'HTML',
  VNext_PlanSidebar: 'HTML',
  VNext_ReviewSidebar: 'HTML',
  VNext_Tests: 'SERVER_JS',
  VNext_UX: 'SERVER_JS',
  appsscript: 'JSON'
});

/**
 * Creates a new blank Master Template container with a client-only bound script.
 * The caller may then initialize its sheets with vNextAdminInitializeTemplate_.
 */
function vNextClientRuntimeCreateBoundSpreadsheet_(request) {
  try {
    vNextClientRuntimeRequireConfigurator_();
    var req = request || {};
    var title = String(req.title || '').trim();
    if (!title) throw new Error('title is required.');
    var bundle = vNextClientRuntimeVerifiedBundle_();
    var spreadsheet = SpreadsheetApp.create(title);
    var file = DriveApp.getFileById(spreadsheet.getId());
    if (req.folderId) file.moveTo(DriveApp.getFolderById(String(req.folderId)));
    try {
      var project = vNextClientRuntimeApiRequest_('/projects', 'post', {
        title: title + ' Client Runtime',
        parentId: spreadsheet.getId()
      });
      if (!project || !project.scriptId) throw new Error('Apps Script API did not return scriptId.');
      vNextClientRuntimePutContent_(project.scriptId, bundle);
      Logger.log('[vNext Client Runtime] clean template created spreadsheet=%s script=%s bundle=%s', spreadsheet.getId(), project.scriptId, bundle.sha256);
      return {
        spreadsheetId: spreadsheet.getId(),
        spreadsheetUrl: spreadsheet.getUrl(),
        scriptId: project.scriptId,
        runtimeVersion: bundle.version,
        bundleSha256: bundle.sha256
      };
    } catch (installError) {
      try { file.setName('[SETUP FAILED] ' + title); } catch (renameError) { Logger.log('[vNext Client Runtime] incomplete file rename failed.'); }
      throw new Error('Client runtimeのbindingに失敗しました。未完了Spreadsheetは保持されています。spreadsheetId=' + spreadsheet.getId() + '; cause=' + vNextClientRuntimeErrorText_(installError));
    }
  } catch (error) {
    Logger.log('[vNext Client Runtime] create failed: ' + vNextClientRuntimeErrorText_(error));
    throw error;
  }
}

/** Updates an already-known Template script ID for a new stable release. */
function vNextClientRuntimeUpdateScriptContent_(scriptId) {
  try {
    vNextClientRuntimeRequireConfigurator_();
    var id = String(scriptId || '').trim();
    if (!/^[A-Za-z0-9_-]{20,}$/.test(id)) throw new Error('A valid Template scriptId is required.');
    var bundle = vNextClientRuntimeVerifiedBundle_();
    vNextClientRuntimePutContent_(id, bundle);
    Logger.log('[vNext Client Runtime] template updated script=%s bundle=%s', id, bundle.sha256);
    return { ok: true, scriptId: id, runtimeVersion: bundle.version, bundleSha256: bundle.sha256 };
  } catch (error) {
    Logger.log('[vNext Client Runtime] update failed: ' + vNextClientRuntimeErrorText_(error));
    throw error;
  }
}

/**
 * Copies a previously verified client-only runtime between two known Apps
 * Script projects. It never discovers IDs and never accepts content outside
 * the fixed eight-code-file plus manifest contract.
 */
function vNextClientRuntimeCopyScriptContent_(sourceScriptId, targetScriptId, expectedSha256) {
  try {
    vNextClientRuntimeRequireConfigurator_();
    var sourceId = vNextClientRuntimeValidateScriptId_(sourceScriptId, 'sourceScriptId');
    var targetId = vNextClientRuntimeValidateScriptId_(targetScriptId, 'targetScriptId');
    if (sourceId === targetId) throw new Error('sourceScriptId and targetScriptId must be different.');
    var expectedSha = vNextClientRuntimeValidateSha256_(expectedSha256);
    var sourceContent = vNextClientRuntimeGetContent_(sourceId);
    var verifiedSource = vNextClientRuntimeVerifyScriptContent_(sourceContent, sourceId, expectedSha);
    var updateResult = vNextClientRuntimePutContent_(targetId, verifiedSource);
    if (updateResult && updateResult.scriptId && String(updateResult.scriptId) !== targetId) {
      throw new Error('Apps Script update response scriptId does not match targetScriptId.');
    }
    var resultHasContent = updateResult && Array.isArray(updateResult.files);
    var targetContent = resultHasContent ? updateResult : vNextClientRuntimeGetContent_(targetId);
    var verifiedTarget = vNextClientRuntimeVerifyScriptContent_(targetContent, targetId, expectedSha);
    Logger.log('[vNext Client Runtime] content copied source=%s target=%s bundle=%s files=%s',
      sourceId, targetId, expectedSha, verifiedTarget.files.length);
    return {
      ok: true,
      sourceScriptId: sourceId,
      targetScriptId: targetId,
      bundleSha256: verifiedTarget.sha256,
      fileCount: verifiedTarget.files.length,
      updateResult: {
        scriptId: String(updateResult && updateResult.scriptId || targetId),
        fileCount: updateResult && Array.isArray(updateResult.files) ? updateResult.files.length : 0,
        verificationSource: resultHasContent ? 'UPDATE_RESPONSE' : 'POST_WRITE_GET'
      }
    };
  } catch (error) {
    Logger.log('[vNext Client Runtime] copy failed: ' + vNextClientRuntimeErrorText_(error));
    throw error;
  }
}

/** Creates an empty Spreadsheet and installs the complete Admin runtime. */
function vNextAdminRuntimeCreateBoundSpreadsheet_(request) {
  try {
    vNextClientRuntimeRequireConfigurator_();
    var req = request || {};
    var title = String(req.title || '').trim();
    if (!title || title.length > 200) throw new Error('title is required and must be 200 characters or fewer.');
    var folderId = req.folderId
      ? vNextAdminRuntimeValidateSpreadsheetId_(req.folderId, 'folderId')
      : '';
    var sourceId = vNextClientRuntimeValidateScriptId_(
      req.sourceScriptId || ScriptApp.getScriptId(),
      'sourceScriptId'
    );
    var spreadsheet = SpreadsheetApp.create(title);
    var file = DriveApp.getFileById(spreadsheet.getId());
    try {
      if (folderId) file.moveTo(DriveApp.getFolderById(folderId));
      var project = vNextClientRuntimeApiRequest_('/projects', 'post', {
        title: title + ' Admin Runtime',
        parentId: spreadsheet.getId()
      });
      var targetId = vNextClientRuntimeValidateScriptId_(project && project.scriptId, 'targetScriptId');
      var copied = vNextAdminRuntimeCopyScriptContent_(sourceId, targetId, spreadsheet.getId());
      Logger.log('[vNext Admin Runtime] empty Hub created spreadsheet=%s script=%s source=%s bundle=%s',
        spreadsheet.getId(), targetId, sourceId, copied.adminRuntimeSha256);
      return {
        ok: true,
        spreadsheetId: spreadsheet.getId(),
        spreadsheetUrl: spreadsheet.getUrl(),
        scriptId: targetId,
        sourceScriptId: sourceId,
        adminRuntimeSha256: copied.adminRuntimeSha256,
        fileCount: copied.fileCount,
        copyResult: copied.updateResult
      };
    } catch (installError) {
      try { file.setName('[SETUP FAILED] ' + title); }
      catch (renameError) { Logger.log('[vNext Admin Runtime] incomplete file rename failed.'); }
      throw new Error('Admin runtimeのbindingに失敗しました。未完了Spreadsheetは保持されています。spreadsheetId=' +
        spreadsheet.getId() + '; cause=' + vNextClientRuntimeErrorText_(installError));
    }
  } catch (error) {
    Logger.log('[vNext Admin Runtime] create failed: ' + vNextClientRuntimeErrorText_(error));
    throw error;
  }
}

/** Copies the exact full Admin runtime into a known Spreadsheet-bound project. */
function vNextAdminRuntimeCopyScriptContent_(sourceScriptId, targetScriptId, expectedTargetSpreadsheetId) {
  try {
    vNextClientRuntimeRequireConfigurator_();
    var sourceId = vNextClientRuntimeValidateScriptId_(sourceScriptId, 'sourceScriptId');
    var targetId = vNextClientRuntimeValidateScriptId_(targetScriptId, 'targetScriptId');
    if (sourceId === targetId) throw new Error('sourceScriptId and targetScriptId must be different.');
    var spreadsheetId = vNextAdminRuntimeValidateSpreadsheetId_(expectedTargetSpreadsheetId, 'expectedTargetSpreadsheetId');
    var targetProject = vNextAdminRuntimeAssertBoundParent_(targetId, spreadsheetId);
    var sourceContent = vNextClientRuntimeGetContent_(sourceId);
    var verifiedSource = vNextAdminRuntimeVerifyScriptContent_(sourceContent, sourceId);
    var updateResult = vNextClientRuntimePutContent_(targetId, verifiedSource);
    if (updateResult && updateResult.scriptId && String(updateResult.scriptId) !== targetId) {
      throw new Error('Apps Script update response scriptId does not match targetScriptId.');
    }
    var resultHasContent = updateResult && Array.isArray(updateResult.files);
    var targetContent = resultHasContent ? updateResult : vNextClientRuntimeGetContent_(targetId);
    var verifiedTarget = vNextAdminRuntimeVerifyScriptContent_(targetContent, targetId);
    if (verifiedTarget.sha256 !== verifiedSource.sha256) {
      throw new Error('Admin runtime target content does not match the verified source content.');
    }
    Logger.log('[vNext Admin Runtime] content copied source=%s target=%s parent=%s bundle=%s files=%s',
      sourceId, targetId, spreadsheetId, verifiedTarget.sha256, verifiedTarget.files.length);
    return {
      ok: true,
      sourceScriptId: sourceId,
      targetScriptId: targetId,
      targetSpreadsheetId: spreadsheetId,
      targetProject: targetProject,
      adminRuntimeSha256: verifiedTarget.sha256,
      fileCount: verifiedTarget.files.length,
      updateResult: {
        scriptId: String(updateResult && updateResult.scriptId || targetId),
        fileCount: updateResult && Array.isArray(updateResult.files) ? updateResult.files.length : 0,
        verificationSource: resultHasContent ? 'UPDATE_RESPONSE' : 'POST_WRITE_GET'
      }
    };
  } catch (error) {
    Logger.log('[vNext Admin Runtime] copy failed: ' + vNextClientRuntimeErrorText_(error));
    throw error;
  }
}

function vNextAdminRuntimeValidateSpreadsheetId_(value, fieldName) {
  var text = String(value === undefined || value === null ? '' : value);
  if (text !== text.trim() || !/^[A-Za-z0-9_-]{20,200}$/.test(text)) {
    throw new Error((fieldName || 'spreadsheetId') + ' is invalid.');
  }
  return text;
}

function vNextAdminRuntimeAssertBoundParent_(scriptId, spreadsheetId) {
  var project = vNextClientRuntimeApiRequest_('/projects/' + encodeURIComponent(scriptId), 'get');
  if (!project || String(project.scriptId || '') !== String(scriptId) ||
      String(project.parentId || '') !== String(spreadsheetId)) {
    throw new Error('Target Apps Script project is not bound to expectedTargetSpreadsheetId.');
  }
  return {
    scriptId: String(project.scriptId),
    parentId: String(project.parentId),
    title: String(project.title || '')
  };
}

function vNextAdminRuntimeVerifyScriptContent_(content, expectedScriptId) {
  if (!content || typeof content !== 'object' || !Array.isArray(content.files)) {
    throw new Error('Admin Apps Script content response is missing files.');
  }
  if (content.scriptId && String(content.scriptId) !== String(expectedScriptId)) {
    throw new Error('Admin Apps Script content scriptId does not match the requested project.');
  }
  var files = vNextAdminRuntimeValidateFiles_(content.files);
  return {
    scriptId: String(expectedScriptId),
    sha256: vNextClientRuntimeFilesSha256_(files),
    files: files
  };
}

function vNextAdminRuntimeValidateFiles_(files) {
  var expectedNames = Object.keys(VNEXT_ADMIN_RUNTIME_FILE_TYPES_).sort();
  if (!Array.isArray(files) || files.length !== expectedNames.length) {
    throw new Error('Admin runtime must contain exactly the 15 clasp-target files.');
  }
  var byName = {};
  files.forEach(function (file) {
    if (!file || typeof file !== 'object') throw new Error('Admin runtime contains an invalid file record.');
    var name = String(file.name || '');
    if (!Object.prototype.hasOwnProperty.call(VNEXT_ADMIN_RUNTIME_FILE_TYPES_, name) || byName[name]) {
      throw new Error('Admin runtime file allowlist mismatch: ' + name);
    }
    var expectedType = VNEXT_ADMIN_RUNTIME_FILE_TYPES_[name];
    if (String(file.type || '') !== expectedType) throw new Error('Admin runtime file type mismatch: ' + name);
    if (typeof file.source !== 'string') throw new Error('Admin runtime file source is invalid: ' + name);
    byName[name] = { name: name, type: expectedType, source: file.source };
  });
  expectedNames.forEach(function (name) {
    if (!byName[name]) throw new Error('Admin runtime file is missing: ' + name);
  });
  var ordered = expectedNames.map(function (name) { return byName[name]; });
  vNextAdminRuntimeValidateManifest_(byName.appsscript.source);
  return ordered;
}

function vNextAdminRuntimeValidateManifest_(source) {
  var manifest;
  try { manifest = JSON.parse(String(source || '')); }
  catch (error) { throw new Error('Admin runtime manifest is not valid JSON.'); }
  if (!manifest || manifest.runtimeVersion !== 'V8') throw new Error('Admin runtime manifest must use V8.');
  return manifest;
}

function vNextClientRuntimeValidateScriptId_(value, fieldName) {
  var text = String(value === undefined || value === null ? '' : value);
  if (text !== text.trim() || !/^[A-Za-z0-9_-]{20,200}$/.test(text)) {
    throw new Error((fieldName || 'scriptId') + ' is invalid.');
  }
  return text;
}

function vNextClientRuntimeValidateSha256_(value) {
  var text = String(value === undefined || value === null ? '' : value);
  if (text !== text.trim() || !/^[a-fA-F0-9]{64}$/.test(text)) throw new Error('expectedSha256 must be exactly 64 hexadecimal characters.');
  return text.toLowerCase();
}

function vNextClientRuntimeGetContent_(scriptId) {
  return vNextClientRuntimeApiRequest_('/projects/' + encodeURIComponent(scriptId) + '/content', 'get');
}

/** Proves that a stored script ID is bound to the intended Spreadsheet. */
function vNextClientRuntimeAssertBoundParent_(scriptId, spreadsheetId) {
  var id = vNextClientRuntimeValidateScriptId_(scriptId, 'scriptId');
  var parentId = String(spreadsheetId || '').trim();
  if (!/^[A-Za-z0-9_-]{20,200}$/.test(parentId)) throw new Error('spreadsheetId is invalid.');
  var project = vNextClientRuntimeApiRequest_('/projects/' + encodeURIComponent(id), 'get');
  if (!project || String(project.scriptId || '') !== id || String(project.parentId || '') !== parentId) {
    throw new Error('Apps Script project is not bound to the registered Spreadsheet.');
  }
  return { scriptId: id, parentId: parentId, title: String(project.title || '') };
}

function vNextClientRuntimeVerifyScriptContent_(content, expectedScriptId, expectedSha256) {
  if (!content || typeof content !== 'object' || !Array.isArray(content.files)) {
    throw new Error('Apps Script content response is missing files.');
  }
  if (content.scriptId && String(content.scriptId) !== String(expectedScriptId)) {
    throw new Error('Apps Script content scriptId does not match the requested project.');
  }
  var files = vNextClientRuntimeValidateFiles_(content.files);
  var digest = vNextClientRuntimeFilesSha256_(files);
  if (digest !== String(expectedSha256 || '').toLowerCase()) {
    throw new Error('Client runtime content hash does not match expectedSha256.');
  }
  return { scriptId: String(expectedScriptId), sha256: digest, files: files };
}

function vNextClientRuntimeValidateFiles_(files) {
  var expectedNames = Object.keys(VNEXT_CLIENT_RUNTIME_FILE_TYPES_);
  if (!Array.isArray(files) || files.length !== expectedNames.length) {
    throw new Error('Client runtime must contain exactly eight client files and one manifest.');
  }
  var byName = {};
  files.forEach(function (file) {
    if (!file || typeof file !== 'object') throw new Error('Client runtime contains an invalid file record.');
    var name = String(file.name || '');
    if (!Object.prototype.hasOwnProperty.call(VNEXT_CLIENT_RUNTIME_FILE_TYPES_, name) || byName[name]) {
      throw new Error('Client runtime file allowlist mismatch: ' + name);
    }
    var expectedType = VNEXT_CLIENT_RUNTIME_FILE_TYPES_[name];
    if (String(file.type || '') !== expectedType) throw new Error('Client runtime file type mismatch: ' + name);
    if (typeof file.source !== 'string') throw new Error('Client runtime file source is invalid: ' + name);
    byName[name] = { name: name, type: expectedType, source: file.source };
  });
  expectedNames.forEach(function (name) {
    if (!byName[name]) throw new Error('Client runtime file is missing: ' + name);
  });
  var ordered = expectedNames.map(function (name) { return byName[name]; });
  vNextClientRuntimeValidateManifest_(byName.appsscript.source);
  vNextClientRuntimeRejectForbiddenContent_(ordered);
  return ordered;
}

function vNextClientRuntimeValidateManifest_(source) {
  var manifest;
  try { manifest = JSON.parse(String(source || '')); }
  catch (error) { throw new Error('Client runtime manifest is not valid JSON.'); }
  if (!manifest || manifest.runtimeVersion !== 'V8') throw new Error('Client runtime manifest must use V8.');
  var scopes = Array.isArray(manifest.oauthScopes) ? manifest.oauthScopes.map(String).sort() : [];
  var expected = VNEXT_CLIENT_RUNTIME_OAUTH_SCOPES_.slice().sort();
  if (JSON.stringify(scopes) !== JSON.stringify(expected)) {
    throw new Error('Client runtime manifest OAuth scopes do not match the minimal allowlist.');
  }
  return manifest;
}

function vNextClientRuntimeRejectForbiddenContent_(files) {
  var combined = (files || []).map(function (file) { return file.source; }).join('\n');
  var forbidden = [
    /Forecast_Agent\.js|initializeAllSheets_|A-1～A-10/,
    /VNext_Admin|vNextBuildAdminMenu_|vNextAdminBootstrap|vNextIsAdminHub_/,
    /VNext_Engine|vNextRunForecast_|vNextSimulateForecast_/,
    /FORECAST_SOURCE_SPREADSHEET_ID|VNEXT_ZAC_SOURCE_SPREADSHEET_ID/,
    /VERTEX_[A-Z_]+|VertexAI|aiplatform/,
    /DriveApp|UrlFetchApp|SpreadsheetApp\.openById|PropertiesService|ScriptApp/,
    /auth\/cloud-platform|auth\/drive(?:["/])|auth\/script\.projects|auth\/script\.external_request/
  ];
  forbidden.forEach(function (pattern) {
    if (pattern.test(combined)) throw new Error('Client runtime contains a forbidden capability: ' + pattern);
  });
  return true;
}

function vNextClientRuntimeFilesSha256_(files) {
  var joined = (files || []).map(function (file) {
    return file.name + '\u0000' + file.type + '\u0000' + file.source;
  }).join('\u0000');
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, joined, Utilities.Charset.UTF_8)
    .map(function (byte) {
      var value = byte < 0 ? byte + 256 : byte;
      return ('0' + value.toString(16)).slice(-2);
    }).join('');
}

function vNextClientRuntimePutContent_(scriptId, bundle) {
  var body = { files: bundle.files.map(function (file) {
    return { name: file.name, type: file.type, source: file.source };
  }) };
  return vNextClientRuntimeApiRequest_('/projects/' + encodeURIComponent(scriptId) + '/content', 'put', body);
}

function vNextClientRuntimeApiRequest_(path, method, body) {
  var response = null;
  for (var attempt = 0; attempt < 3; attempt++) {
    var requestOptions = {
      method: method,
      contentType: 'application/json; charset=utf-8',
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    };
    if (body !== undefined && body !== null && String(method || '').toLowerCase() !== 'get') {
      requestOptions.payload = JSON.stringify(body);
    }
    response = UrlFetchApp.fetch(VNEXT_CLIENT_RUNTIME_API_ROOT_ + path, requestOptions);
    var status = response.getResponseCode();
    if (status >= 200 && status < 300) {
      var text = response.getContentText();
      return text ? JSON.parse(text) : {};
    }
    if ([429, 500, 502, 503, 504].indexOf(status) < 0 || attempt === 2) {
      throw new Error('Apps Script API request failed status=' + status + '; response=' + response.getContentText().slice(0, 1200));
    }
    Utilities.sleep(500 * Math.pow(2, attempt));
  }
  throw new Error('Apps Script API request failed without a response.');
}

/**
 * Enables only script.googleapis.com for the Cloud project reported by the
 * service-disabled error. The project number is discovered from Google's
 * signed API response and is never accepted from employee/client input.
 */
function vNextClientRuntimeEnableRequiredAppsScriptApi_() {
  vNextClientRuntimeRequireConfigurator_();
  var probePath = '/projects/' + encodeURIComponent(ScriptApp.getScriptId());
  try {
    vNextClientRuntimeApiRequest_(probePath, 'get');
    return { ok: true, alreadyEnabled: true, service: 'script.googleapis.com' };
  } catch (probeError) {
    var errorText = vNextClientRuntimeErrorText_(probeError);
    if (errorText.indexOf('SERVICE_DISABLED') < 0 || errorText.indexOf('script.googleapis.com') < 0) {
      throw probeError;
    }
    var projectMatch = /"consumer"\s*:\s*"projects\/(\d{6,20})"/.exec(errorText) ||
      /[?&]project=(\d{6,20})/.exec(errorText);
    if (!projectMatch) throw new Error('Apps Script APIの所属Cloud project番号を検証できませんでした。');
    var projectNumber = projectMatch[1];
    var token = ScriptApp.getOAuthToken();
    var serviceUrl = 'https://serviceusage.googleapis.com/v1/projects/' + projectNumber +
      '/services/script.googleapis.com:enable';
    var enableResponse = UrlFetchApp.fetch(serviceUrl, {
      method: 'post',
      contentType: 'application/json; charset=utf-8',
      headers: { Authorization: 'Bearer ' + token },
      payload: '{}',
      muteHttpExceptions: true
    });
    var enableStatus = enableResponse.getResponseCode();
    if (enableStatus < 200 || enableStatus >= 300) {
      throw new Error('Apps Script APIの有効化に失敗しました status=' + enableStatus +
        '; response=' + enableResponse.getContentText().slice(0, 800));
    }
    var operation = {};
    try { operation = JSON.parse(enableResponse.getContentText() || '{}'); }
    catch (ignoredJsonError) { operation = {}; }
    if (operation.name && /^[A-Za-z0-9_\/-]{5,300}$/.test(String(operation.name))) {
      var operationUrl = 'https://serviceusage.googleapis.com/v1/' + String(operation.name).replace(/^\//, '');
      for (var operationAttempt = 0; operationAttempt < 15; operationAttempt++) {
        var operationResponse = UrlFetchApp.fetch(operationUrl, {
          method: 'get', headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true
        });
        if (operationResponse.getResponseCode() >= 200 && operationResponse.getResponseCode() < 300) {
          var operationBody = {};
          try { operationBody = JSON.parse(operationResponse.getContentText() || '{}'); }
          catch (ignoredOperationJson) { operationBody = {}; }
          if (operationBody.error) throw new Error('Apps Script API有効化operationが失敗しました。');
          if (operationBody.done === true) break;
        }
        Utilities.sleep(1500);
      }
    }
    for (var probeAttempt = 0; probeAttempt < 15; probeAttempt++) {
      try {
        vNextClientRuntimeApiRequest_(probePath, 'get');
        return {
          ok: true, alreadyEnabled: false, service: 'script.googleapis.com',
          projectNumber: projectNumber
        };
      } catch (retryError) {
        var retryText = vNextClientRuntimeErrorText_(retryError);
        if (retryText.indexOf('SERVICE_DISABLED') < 0) throw retryError;
        Utilities.sleep(2000);
      }
    }
    throw new Error('Apps Script APIの有効化は受け付けられましたが、利用可能になるまで時間がかかっています。1分後に再実行してください。');
  }
}

function vNextClientRuntimeVerifiedBundle_() {
  if (typeof VNEXT_CLIENT_RUNTIME_BUNDLE_ === 'undefined') throw new Error('VNext_ClientRuntimeBundle.js is not deployed.');
  var bundle = VNEXT_CLIENT_RUNTIME_BUNDLE_;
  var files = vNextClientRuntimeValidateFiles_(bundle.files || []);
  var digest = vNextClientRuntimeFilesSha256_(files);
  if (digest !== String(bundle.sha256 || '')) throw new Error('Client runtime bundle hash mismatch.');
  return bundle;
}

function vNextClientRuntimeRequireConfigurator_() {
  if (typeof vNextAdminAssertRuntimeConfigurator_ === 'function') {
    vNextAdminAssertRuntimeConfigurator_();
    return true;
  }
  var active = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  var effective = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  if (!active || !effective || active !== effective) throw new Error('管理者本人の実行が必要です。');
  return true;
}

function vNextClientRuntimeErrorText_(error) {
  return error && error.message ? String(error.message) : String(error || 'unknown error');
}
