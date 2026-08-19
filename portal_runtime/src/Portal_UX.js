/**
 * Forecast vNext shared employee portal UI.
 * Daily work stays in one guidance sidebar. The top menu is only a recovery path.
 */

function doGet() {
  try {
    return HtmlService.createHtmlOutputFromFile('Portal_Entry')
      .setTitle(VNEXT_PORTAL_NAMING.SYSTEM)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (error) {
    vNextPortalLog_('doGet failed', error);
    return HtmlService.createHtmlOutput(
      '<p>入口を表示できませんでした。社内アカウントで開き直してください。</p>'
    ).setTitle(VNEXT_PORTAL_NAMING.SYSTEM);
  }
}

function vNextPortalPrepareOpenExperience() {
  try {
    var installed = vNextPortalEnsureGuidanceOnOpenTrigger_();
    return { ok: true, installed: installed };
  } catch (error) {
    vNextPortalLog_('vNextPortalPrepareOpenExperience skipped', error);
    return { ok: false };
  }
}

function onOpen(event) {
  return vNextPortalOnOpen_(event);
}

function vNextPortalOnOpen_() {
  try {
    SpreadsheetApp.getUi().createMenu(VNEXT_PORTAL.MENU_NAME)
      .addItem('案内を開く', 'vNextPortalGoHomeAndShowGuidance')
      .addToUi();
    return true;
  } catch (error) {
    vNextPortalLog_('vNextPortalOnOpen_ failed', error);
    return false;
  }
}

function vNextPortalGoHomeAndShowGuidance() {
  try {
    vNextPortalGoHome();
    vNextPortalOpenGuidanceSidebarQuietly_();
  } catch (error) {
    vNextPortalLog_('vNextPortalGoHomeAndShowGuidance failed', error);
    vNextPortalShowError_('案内を開けませんでした。', error);
  }
}

function vNextPortalInstalledGuidanceOnOpen(e) {
  return vNextPortalOpenGuidanceSidebarQuietly_();
}

function vNextPortalEnsureGuidanceOnOpenTrigger_() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var handler = 'vNextPortalInstalledGuidanceOnOpen';
  function isOpenHandler(trigger) {
    return trigger.getHandlerFunction() === handler &&
      trigger.getEventType() === ScriptApp.EventType.ON_OPEN;
  }
  if (ScriptApp.getProjectTriggers().some(isOpenHandler)) return false;
  try {
    ScriptApp.getUserTriggers(spreadsheet).filter(isOpenHandler).forEach(function (trigger) {
      ScriptApp.deleteTrigger(trigger);
    });
  } catch (cleanupError) {
    vNextPortalLog_('guidance trigger cleanup skipped', cleanupError);
  }
  ScriptApp.newTrigger(handler).forSpreadsheet(spreadsheet).onOpen().create();
  return true;
}

function vNextPortalOpenGuidanceSidebarQuietly_() {
  try {
    var html = HtmlService.createTemplateFromFile('Portal_CreateSidebar').evaluate()
      .setTitle('次にすること')
      .setWidth(430);
    SpreadsheetApp.getUi().showSidebar(html);
    try { vNextPortalEnsureGuidanceOnOpenTrigger_(); }
    catch (triggerError) { vNextPortalLog_('guidance trigger skipped', triggerError); }
    return true;
  } catch (error) {
    vNextPortalLog_('vNextPortalOpenGuidanceSidebarQuietly_ skipped', error);
    return false;
  }
}

function vNextPortalGoHome() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    vNextPortalRefreshViews_(spreadsheet);
    var home = spreadsheet.getSheetByName(VNEXT_PORTAL.HOME_SHEET);
    spreadsheet.setActiveSheet(home);
    home.getRange('A1').activate();
    spreadsheet.toast('最新の受付・申請状況に更新しました。', VNEXT_PORTAL.MENU_NAME, 3);
    return { ok: true };
  } catch (error) {
    vNextPortalLog_('vNextPortalGoHome failed', error);
    vNextPortalShowError_('ホームを更新できませんでした。', error);
  }
}

/** Refreshes only Home after a creation request, avoiding a full FY-tab rebuild. */
function vNextPortalShowRequestOnHome(requestId) {
  try {
    var id = vNextPortalNormalizeRequestId_(requestId);
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var home = spreadsheet.getSheetByName(VNEXT_PORTAL.HOME_SHEET);
    if (!home) throw new Error('ホームが準備されていません。管理担当者へ連絡してください。');
    var data = vNextPortalGetLocalViewData_(spreadsheet);
    var request = null;
    data.requests.some(function (item) {
      if (item.requestId !== id) return false;
      request = item;
      return true;
    });
    if (!request) throw new Error('受付済みの依頼をホームで確認できませんでした。');
    vNextPortalRenderHome_(home, data);
    spreadsheet.setActiveSheet(home);
    home.getRange('A1').activate();
    spreadsheet.toast('受付済みです。ホームに最新状況を表示しました。', VNEXT_PORTAL.MENU_NAME, 4);
    return vNextPortalRequestProgressModel_(request);
  } catch (error) {
    vNextPortalLog_('vNextPortalShowRequestOnHome failed', error);
    throw error;
  }
}

function vNextPortalOpenCreateSidebar() {
  return vNextPortalOpenGuidanceSidebarQuietly_();
}

function vNextPortalOpenHelp() {
  try {
    var html = HtmlService.createHtmlOutput(
      '<!doctype html><html><head><base target="_top"><style>' +
      'body{font-family:Arial,sans-serif;color:#202124;padding:18px;line-height:1.65}' +
      'h2{font-size:18px;margin:0 0 12px}.step{padding:11px 12px;margin:9px 0;background:#f8f9fa;border-radius:8px}' +
      '.note{margin-top:16px;padding:12px;background:#e8f0fe;border-left:4px solid #1a73e8}' +
      '</style></head><body><h2>申請入口（第1層）の使い方</h2>' +
      '<div class="step"><b>1. 開く</b><br>年度予算策定 Web入口の年度ボタンとクライアントボタンから、すでにあるクライアント年度ブックを選びます。</div>' +
      '<div class="step"><b>2. なければ申請</b><br>右側の案内から「新しいクライアント年度ブックを申請する」を選びます。候補確認後に作成依頼を送ります。</div>' +
      '<div class="step"><b>3. 状況を見る</b><br>ホームには受付済み・作成中・完成・確認が必要、の状態が表示されます。</div>' +
      '<div class="note">申請入口ではセルへ直接入力しません。案内が出ないときだけ、上部メニュー「' + VNEXT_PORTAL_NAMING.MENU + '」→「案内を開く」を使います。</div>' +
      '</body></html>'
    ).setTitle('使い方・困ったとき').setWidth(420);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (error) {
    vNextPortalLog_('vNextPortalOpenHelp failed', error);
    vNextPortalShowError_('ヘルプを開けませんでした。', error);
  }
}

function vNextPortalRefreshViews_(spreadsheet) {
  var ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  vNextPortalEnsureStructure_(ss);
  var data = vNextPortalGetLocalViewData_(ss);
  data.years.forEach(function (year) { vNextPortalEnsureFiscalYearSheet_(ss, year); });
  vNextPortalRenderHome_(ss.getSheetByName(VNEXT_PORTAL.HOME_SHEET), data);
  data.years.forEach(function (year) {
    vNextPortalRenderFiscalYear_(ss.getSheetByName('FY' + year), year, data);
  });
  vNextPortalHideInternalSheet_(ss.getSheetByName(VNEXT_PORTAL.DIRECTORY_SHEET));
  vNextPortalHideInternalSheet_(ss.getSheetByName(VNEXT_PORTAL.REQUEST_SHEET));
  vNextPortalHideInternalSheet_(ss.getSheetByName(VNEXT_PORTAL.CLIENT_CATALOG_SHEET));
  return { ok: true, years: data.years, requestCount: data.requests.length, directoryCount: data.directory.length };
}

function vNextPortalRenderHome_(sheet, data) {
  var requests = data.requests.slice().sort(function (a, b) {
    return String(b.updatedAt || b.requestedAt).localeCompare(String(a.updatedAt || a.requestedAt));
  }).slice(0, 20);
  vNextPortalResetViewSheet_(sheet, Math.max(12, requests.length + 6), 6);
  sheet.getRange('A1').setValue(VNEXT_PORTAL_NAMING.PORTAL)
    .setFontSize(16).setFontWeight('bold').setFontColor('#202124');
  sheet.getRange('A2').setValue('既存のクライアント年度ブックは年度予算策定 Web入口から開きます。新規申請は右側の案内から行います。案内が出ないときだけ、上部メニュー「' + VNEXT_PORTAL_NAMING.MENU + '」→「案内を開く」を使います。')
    .setFontSize(10).setFontColor('#5f6368');
  sheet.getRange('A3').setValue('作成依頼の状況')
    .setFontSize(11).setFontWeight('bold').setFontColor('#202124');
  sheet.getRange('E3').setValue('表示更新').setFontColor('#5f6368').setFontSize(9);
  sheet.getRange('F3').setValue(vNextPortalDisplayDateTime_(new Date()))
    .setFontColor('#5f6368').setFontSize(9).setHorizontalAlignment('right');

  var headers = ['状態', 'クライアント', '年度', '次の案内', '更新', '開く'];
  sheet.getRange(4, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setFontColor('#3c4043').setBackground('#f1f3f4')
    .setHorizontalAlignment('left');
  if (!requests.length) {
    sheet.getRange('A5').setValue('作成依頼はまだありません。')
      .setFontColor('#5f6368').setFontStyle('italic');
  } else {
    var rows = requests.map(function (request) {
      var progress = vNextPortalRequestProgressModel_(request);
      return [
        request.statusLabel, vNextPortalCellText_(request.clientName), 'FY' + request.fiscalYear,
        vNextPortalCellText_(progress.nextAction),
        vNextPortalDisplayDateTime_(request.updatedAt), request.url ? '開く' : '準備中'
      ];
    });
    sheet.getRange(5, 1, rows.length, headers.length).setValues(rows).setWrap(true).setVerticalAlignment('middle');
    requests.forEach(function (request, index) {
      var row = 5 + index;
      sheet.getRange(row, 1).setFontColor(vNextPortalStatusColor_(request.status)).setFontWeight('bold');
      if (request.url) vNextPortalWriteLink_(sheet.getRange(row, 6), request.url, '開く');
    });
    try { sheet.getRange(4, 1, rows.length + 1, headers.length).createFilter(); }
    catch (error) { vNextPortalLog_('Home filter creation skipped', error); }
  }
  sheet.getRange(4, 1, Math.max(1, requests.length + 1), headers.length)
    .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
  sheet.setFrozenRows(4);
  sheet.setHiddenGridlines(false);
  [118, 210, 76, 330, 124, 72].forEach(function (width, index) { sheet.setColumnWidth(index + 1, width); });
  sheet.setRowHeight(1, 30);
  sheet.setRowHeight(2, 24);
  vNextPortalProtectWarningOnly_(sheet, '自動生成画面（入力は右側の案内から行います）');
}

function vNextPortalRenderFiscalYear_(sheet, fiscalYear, data) {
  var entries = vNextPortalFiscalYearEntries_(fiscalYear, data);
  vNextPortalResetViewSheet_(sheet, Math.max(12, entries.length + 6), 9);
  sheet.getRange('A1').setValue('FY' + fiscalYear + ' クライアント年度ブック')
    .setFontSize(16).setFontWeight('bold').setFontColor('#202124');
  sheet.getRange('A2').setValue('クライアントを選び、「開く」からクライアント年度ブックへ進んでください。')
    .setFontSize(10).setFontColor('#5f6368');
  sheet.getRange('H2').setValue('表示更新').setFontColor('#5f6368').setFontSize(9);
  sheet.getRange('I2').setValue(vNextPortalDisplayDateTime_(new Date()))
    .setFontColor('#5f6368').setFontSize(9).setHorizontalAlignment('right');
  var headers = ['状態', 'クライアント', '中心見込み', '採用予測', '最終予算', '担当・関与', '次の対応', '更新日', '開く'];
  sheet.getRange(4, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setFontColor('#3c4043').setBackground('#f1f3f4');
  if (!entries.length) {
    sheet.getRange('A5').setValue('この年度のクライアント年度ブックはまだありません。右側の案内から申請できます。')
      .setFontColor('#5f6368').setFontStyle('italic');
  } else {
    var rows = entries.map(function (entry) {
      var people = [];
      if (entry.forecastOwnerEmail) people.push(entry.forecastOwnerEmail);
      var related = entry.relatedMemberNames.length ? entry.relatedMemberNames : entry.relatedMemberEmails;
      if (related.length) people.push(related.join('、'));
      var actualShortage = /(?:実績|データ).{0,8}不足/.test(String(entry.nextAction || ''));
      return [
        entry.statusLabel,
        vNextPortalCellText_(entry.clientName),
        vNextPortalDisplayPlanningValue_(entry.centerForecast, actualShortage ? '実績不足' : '未算出'),
        vNextPortalDisplayPlanningValue_(entry.adoptedForecast, '未設定'),
        vNextPortalDisplayPlanningValue_(entry.finalBudget, '未設定'),
        people.join(' ／ '),
        vNextPortalCellText_(entry.nextAction),
        vNextPortalDisplayDate_(entry.updatedAt),
        entry.url ? '開く' : '準備中'
      ];
    });
    sheet.getRange(5, 1, rows.length, headers.length).setValues(rows).setWrap(true).setVerticalAlignment('middle');
    sheet.getRange(5, 3, rows.length, 3).setNumberFormat('¥#,##0;[Red]-¥#,##0;―');
    entries.forEach(function (entry, index) {
      var row = 5 + index;
      sheet.getRange(row, 1).setFontColor(vNextPortalStatusColor_(entry.status || entry.state || entry.statusLabel)).setFontWeight('bold');
      if (entry.url) vNextPortalWriteLink_(sheet.getRange(row, 9), entry.url, '開く');
    });
    try {
      var existingFilter = sheet.getFilter();
      if (existingFilter) existingFilter.remove();
      sheet.getRange(4, 1, rows.length + 1, headers.length).createFilter();
    } catch (error) {
      vNextPortalLog_('FY filter creation skipped', error);
    }
  }
  sheet.getRange(4, 1, Math.max(1, entries.length + 1), headers.length)
    .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
  sheet.setFrozenRows(4);
  sheet.setHiddenGridlines(false);
  [106, 190, 104, 104, 104, 210, 260, 100, 68].forEach(function (width, index) { sheet.setColumnWidth(index + 1, width); });
  sheet.setRowHeight(1, 30);
  sheet.setRowHeight(2, 24);
  vNextPortalProtectWarningOnly_(sheet, '自動生成画面（入力は各クライアント年度ブックで行います）');
}

function vNextPortalFiscalYearEntries_(fiscalYear, data) {
  var entries = [];
  var representedRequestIds = {};
  data.directory.filter(function (item) { return item.fiscalYear === fiscalYear; }).forEach(function (item) {
    if (item.requestId) representedRequestIds[item.requestId] = true;
    entries.push({
      clientId: item.clientId,
      clientName: item.clientName,
      state: item.state,
      status: item.state,
      statusLabel: vNextPortalDirectoryStateLabel_(item.state),
      centerForecast: item.centerForecast,
      adoptedForecast: item.adoptedForecast,
      finalBudget: item.finalBudget,
      forecastOwnerEmail: item.forecastOwnerEmail,
      relatedMemberEmails: item.relatedMemberEmails,
      relatedMemberNames: item.relatedMemberNames || [],
      nextAction: item.nextAction || 'クライアント年度ブックで次の対応を確認してください。',
      updatedAt: item.updatedAt,
      url: item.url
    });
  });
  data.requests.filter(function (request) {
    return request.fiscalYear === fiscalYear && !representedRequestIds[request.requestId];
  }).forEach(function (request) {
    entries.push({
      clientId: request.clientId,
      clientName: request.clientName,
      status: request.status,
      statusLabel: request.statusLabel,
      centerForecast: null,
      adoptedForecast: null,
      finalBudget: null,
      forecastOwnerEmail: request.forecastOwnerEmail,
      relatedMemberEmails: request.relatedMemberEmails,
      relatedMemberNames: request.relatedMemberNames || [],
      nextAction: vNextPortalStatusNextAction_(request.status, request.detailMessage, Boolean(request.url)),
      updatedAt: request.updatedAt,
      url: request.url
    });
  });
  entries.sort(function (a, b) {
    return String(a.clientName).localeCompare(String(b.clientName), 'ja');
  });
  return entries;
}

function vNextPortalResetViewSheet_(sheet, rows, columns) {
  vNextPortalEnsureSheetSize_(sheet, rows, columns);
  var filter = sheet.getFilter();
  if (filter) filter.remove();
  var dataRange = sheet.getDataRange();
  dataRange.getMergedRanges().forEach(function (range) { range.breakApart(); });
  var clearRows = Math.max(rows, sheet.getLastRow(), 1);
  var clearColumns = Math.max(columns, sheet.getLastColumn(), 1);
  sheet.getRange(1, 1, clearRows, clearColumns).clear().clearNote();
  sheet.getRange(1, 1, rows, columns).setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
}

function vNextPortalStatusColor_(value) {
  var text = String(value || '').toUpperCase();
  if (/FAILED|REJECTED|作成できません|確認が必要/.test(text)) return '#b3261e';
  if (/COMPLETED|利用できます|OFFICIAL|正式/.test(text)) return '#137333';
  if (/PENDING|VALIDATING|CREATING|受付済み|確認中|作成中/.test(text)) return '#1a73e8';
  return '#3c4043';
}

function vNextPortalDisplayPlanningValue_(value, emptyLabel) {
  return typeof value === 'number' && isFinite(value) ? value : String(emptyLabel || '―');
}

function vNextPortalWriteLink_(range, url, label) {
  if (!vNextPortalSafeBookUrl_(url)) return;
  var rich = SpreadsheetApp.newRichTextValue().setText(label).setLinkUrl(url).build();
  range.setRichTextValue(rich).setFontColor('#1a73e8').setFontWeight('bold').setHorizontalAlignment('center');
}

function vNextPortalCellText_(value) {
  var text = String(value === undefined || value === null ? '' : value);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function vNextPortalDisplayDateTime_(value) {
  if (!value) return '';
  var date = value instanceof Date ? value : new Date(value);
  if (isNaN(date.getTime())) return String(value);
  return Utilities.formatDate(date, Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
}

function vNextPortalDisplayDate_(value) {
  if (!value) return '';
  var date = value instanceof Date ? value : new Date(value);
  if (isNaN(date.getTime())) return String(value);
  return Utilities.formatDate(date, Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyy/MM/dd');
}

function vNextPortalShowError_(prefix, error) {
  var message = error && error.message ? error.message : String(error || '不明なエラー');
  SpreadsheetApp.getUi().alert('エラー', prefix + '\n' + message, SpreadsheetApp.getUi().ButtonSet.OK);
}
