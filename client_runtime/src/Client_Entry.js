/**
 * Forecast vNext Client FY Book entry points.
 * This runtime is intentionally limited to employee UI and client-local records.
 */

function onOpen(e) {
  try {
    var mode = vNextClientDetectMode_();
    if (mode === 'TEMPLATE') {
      SpreadsheetApp.getUi().createMenu('年度計画')
        .addItem('案内を開く', 'vNextOpenTemplateGuidance')
        .addToUi();
      return;
    }
    if (mode !== 'CLIENT') {
      Logger.log('[vNext Client] onOpen skipped because mode is not CLIENT.');
      return;
    }
    vNextHandleOnOpen_(e);
  } catch (error) {
    Logger.log('[vNext Client] onOpen failed: ' + vNextClientErrorText_(error));
  }
}

function onInstall(e) {
  try {
    onOpen(e);
  } catch (error) {
    Logger.log('[vNext Client] onInstall failed: ' + vNextClientErrorText_(error));
  }
}

/** A non-sensitive diagnostic which employees can send to support. */
function vNextClientRuntimeInfo() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var context = vNextGetBookContext_(ss);
    var requiredSheets = Object.keys(VNEXT_CLIENT_CORE.INTERNAL_SHEETS).concat(['VN_BOOK_CONFIG', 'VN_CLIENT_REQUEST']);
    var missing = requiredSheets.filter(function (name) { return !ss.getSheetByName(name); });
    return {
      ok: missing.length === 0,
      runtimeVersion: VNEXT_CLIENT_CORE.RUNTIME_VERSION,
      schemaVersion: VNEXT_CLIENT_CORE.SCHEMA_VERSION,
      bookId: context.bookId,
      state: context.state,
      missingSheets: missing
    };
  } catch (error) {
    Logger.log('[vNext Client] runtime info failed: ' + vNextClientErrorText_(error));
    return { ok: false, error: vNextClientErrorText_(error) };
  }
}

function vNextClientDetectMode_() {
  try {
    var config = vNextClientReadConfig_(SpreadsheetApp.getActiveSpreadsheet());
    return String(config.mode || '').trim().toUpperCase();
  } catch (error) {
    Logger.log('[vNext Client] mode detection failed: ' + vNextClientErrorText_(error));
    return '';
  }
}

function vNextClientErrorText_(error) {
  return error && error.message ? String(error.message) : String(error || '不明なエラー');
}
