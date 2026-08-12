/**
 * Forecast vNext shared employee portal UI.
 * The visible experience is limited to Home, FY tabs, one creation sidebar, and help.
 */

function onOpen(event) {
  return vNextPortalOnOpen_(event);
}

function vNextPortalOnOpen_() {
  try {
    SpreadsheetApp.getUi().createMenu(VNEXT_PORTAL.MENU_NAME)
      .addItem('ホームに戻る', 'vNextPortalGoHome')
      .addItem('新しい年度計画を作る', 'vNextPortalOpenCreateSidebar')
      .addItem('使い方・困ったとき', 'vNextPortalOpenHelp')
      .addToUi();
    return true;
  } catch (error) {
    vNextPortalLog_('vNextPortalOnOpen_ failed', error);
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
    spreadsheet.toast('最新の受付・計画状況に更新しました。', VNEXT_PORTAL.MENU_NAME, 3);
    return { ok: true };
  } catch (error) {
    vNextPortalLog_('vNextPortalGoHome failed', error);
    vNextPortalShowError_('ホームを更新できませんでした。', error);
  }
}

function vNextPortalOpenCreateSidebar() {
  try {
    var html = HtmlService.createTemplateFromFile('Portal_CreateSidebar').evaluate()
      .setTitle('新しい年度計画を作る')
      .setWidth(430);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (error) {
    vNextPortalLog_('vNextPortalOpenCreateSidebar failed', error);
    vNextPortalShowError_('作成画面を開けませんでした。', error);
  }
}

function vNextPortalOpenHelp() {
  try {
    var html = HtmlService.createHtmlOutput(
      '<!doctype html><html><head><base target="_top"><style>' +
      'body{font-family:Arial,sans-serif;color:#202124;padding:18px;line-height:1.65}' +
      'h2{font-size:18px;margin:0 0 12px}.step{padding:11px 12px;margin:9px 0;background:#f8f9fa;border-radius:8px}' +
      '.note{margin-top:16px;padding:12px;background:#e8f0fe;border-left:4px solid #1a73e8}' +
      '</style></head><body><h2>年度計画ポータルの使い方</h2>' +
      '<div class="step"><b>1. 探す</b><br>対象年度のFYタブでクライアント名を検索し、「開く」から専用ブックへ移動します。</div>' +
      '<div class="step"><b>2. なければ作る</b><br>上部メニューから「新しい年度計画を作る」を選びます。候補確認後に作成依頼を送ります。</div>' +
      '<div class="step"><b>3. 状況を見る</b><br>ホームには受付済み・作成中・完成・確認が必要、の状態が表示されます。</div>' +
      '<div class="note">このポータルではセルへ直接入力しません。表示を壊さないため、入力と作成依頼は上部メニューから行ってください。</div>' +
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
  sheet.getRange('A1').setValue('年度計画ポータル')
    .setFontSize(16).setFontWeight('bold').setFontColor('#202124');
  sheet.getRange('A2').setValue('既存の計画はFYタブから確認できます。新規作成と状態更新は上部メニュー「年度計画ポータル」から行います。')
    .setFontSize(10).setFontColor('#5f6368');
  sheet.getRange('A3').setValue('作成依頼の状況')
    .setFontSize(11).setFontWeight('bold').setFontColor('#202124')
    .setNote('依頼後は、受付済み→内容確認中→ブック作成中→利用できます、の順で進みます。');

  var headers = ['状態', 'クライアント', '年度', '次の案内', '更新', '開く'];
  sheet.getRange(4, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setFontColor('#3c4043').setBackground('#f1f3f4')
    .setHorizontalAlignment('left');
  if (!requests.length) {
    sheet.getRange('A5').setValue('作成依頼はまだありません。')
      .setFontColor('#5f6368').setFontStyle('italic');
  } else {
    var rows = requests.map(function (request) {
      return [
        request.statusLabel, vNextPortalCellText_(request.clientName), 'FY' + request.fiscalYear,
        vNextPortalStatusNextAction_(request.status, request.detailMessage, Boolean(request.url)),
        vNextPortalDisplayDateTime_(request.updatedAt), request.url ? '開く' : '準備中'
      ];
    });
    sheet.getRange(5, 1, rows.length, headers.length).setValues(rows).setWrap(true).setVerticalAlignment('middle');
    requests.forEach(function (request, index) {
      var row = 5 + index;
      sheet.getRange(row, 1).setFontColor(vNextPortalStatusColor_(request.status)).setFontWeight('bold');
      sheet.getRange(row, 2).setNote(
        '受付番号: ' + request.requestId + '\n' +
        '作成担当: ' + (request.forecastOwnerEmail || request.requestedBy || '確認中') + '\n' +
        '受付日時: ' + vNextPortalDisplayDateTime_(request.requestedAt)
      );
      if (request.url) vNextPortalWriteLink_(sheet.getRange(row, 6), request.url, '開く');
    });
    sheet.getRange(4, 1, rows.length + 1, headers.length)
      .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
    try { sheet.getRange(4, 1, rows.length + 1, headers.length).createFilter(); }
    catch (error) { vNextPortalLog_('Home filter creation skipped', error); }
  }
  sheet.setFrozenRows(4);
  sheet.setHiddenGridlines(true);
  [118, 210, 76, 330, 124, 72].forEach(function (width, index) { sheet.setColumnWidth(index + 1, width); });
  sheet.setRowHeight(1, 30);
  sheet.setRowHeight(2, 24);
  vNextPortalProtectWarningOnly_(sheet, '自動生成画面（入力は上部メニューから行います）');
}

function vNextPortalRenderFiscalYear_(sheet, fiscalYear, data) {
  var entries = vNextPortalFiscalYearEntries_(fiscalYear, data);
  vNextPortalResetViewSheet_(sheet, Math.max(12, entries.length + 6), 9);
  sheet.getRange('A1').setValue('FY' + fiscalYear + ' 年度計画')
    .setFontSize(16).setFontWeight('bold').setFontColor('#202124');
  sheet.getRange('A2').setValue('クライアントを選び、「開く」から専用ブックへ進んでください。')
    .setFontSize(10).setFontColor('#5f6368');
  var headers = ['状態', 'クライアント', '中心見込み', '採用予測', '最終予算', '担当・関与', '次の対応', '更新日', '開く'];
  sheet.getRange(4, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setFontColor('#3c4043').setBackground('#f1f3f4');
  sheet.getRange('C4').setNote('システムが算出した条件付き予測の中心値です。営業目標ではありません。');
  sheet.getRange('D4').setNote('予測を確認したうえで、計画に採用した金額です。');
  sheet.getRange('E4').setNote('採用予測に営業上積みを加えた正式な予算です。');
  sheet.getRange('F4').setNote('作成担当と、依頼時に登録された関与メンバーです。アクセス権限そのものではありません。');
  if (!entries.length) {
    sheet.getRange('A5').setValue('この年度の計画はまだありません。上部メニューから新しく作成できます。')
      .setFontColor('#5f6368').setFontStyle('italic');
  } else {
    var rows = entries.map(function (entry) {
      var people = [];
      if (entry.forecastOwnerEmail) people.push(entry.forecastOwnerEmail);
      var related = entry.relatedMemberNames.length ? entry.relatedMemberNames : entry.relatedMemberEmails;
      if (related.length) people.push(related.join('、'));
      return [
        entry.statusLabel,
        vNextPortalCellText_(entry.clientName),
        entry.centerForecast,
        entry.adoptedForecast,
        entry.finalBudget,
        people.join(' ／ '),
        entry.nextAction,
        vNextPortalDisplayDate_(entry.updatedAt),
        entry.url ? '開く' : '準備中'
      ];
    });
    sheet.getRange(5, 1, rows.length, headers.length).setValues(rows).setWrap(true).setVerticalAlignment('middle');
    sheet.getRange(5, 3, rows.length, 3).setNumberFormat('¥#,##0;[Red]-¥#,##0;―');
    entries.forEach(function (entry, index) {
      var row = 5 + index;
      sheet.getRange(row, 1).setFontColor(vNextPortalStatusColor_(entry.statusLabel)).setFontWeight('bold');
      if (entry.url) vNextPortalWriteLink_(sheet.getRange(row, 9), entry.url, '開く');
    });
    sheet.getRange(4, 1, rows.length + 1, headers.length)
      .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
    try {
      var existingFilter = sheet.getFilter();
      if (existingFilter) existingFilter.remove();
      sheet.getRange(4, 1, rows.length + 1, headers.length).createFilter();
    } catch (error) {
      vNextPortalLog_('FY filter creation skipped', error);
    }
  }
  sheet.setFrozenRows(4);
  sheet.setHiddenGridlines(true);
  [112, 210, 112, 112, 112, 240, 300, 106, 72].forEach(function (width, index) { sheet.setColumnWidth(index + 1, width); });
  sheet.setRowHeight(1, 30);
  sheet.setRowHeight(2, 24);
  vNextPortalProtectWarningOnly_(sheet, '自動生成画面（各計画の入力は専用ブックで行います）');
}

function vNextPortalFiscalYearEntries_(fiscalYear, data) {
  var entries = [];
  var representedRequestIds = {};
  data.directory.filter(function (item) { return item.fiscalYear === fiscalYear; }).forEach(function (item) {
    if (item.requestId) representedRequestIds[item.requestId] = true;
    entries.push({
      clientId: item.clientId,
      clientName: item.clientName,
      statusLabel: vNextPortalDirectoryStateLabel_(item.state),
      centerForecast: item.centerForecast,
      adoptedForecast: item.adoptedForecast,
      finalBudget: item.finalBudget,
      forecastOwnerEmail: item.forecastOwnerEmail,
      relatedMemberEmails: item.relatedMemberEmails,
      relatedMemberNames: item.relatedMemberNames || [],
      nextAction: item.nextAction || '専用ブックで次の対応を確認してください。',
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
  sheet.getRange(1, 1, clearRows, clearColumns).clear();
  sheet.getRange(1, 1, rows, columns).setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
}

function vNextPortalStatusColor_(value) {
  var text = String(value || '').toUpperCase();
  if (/FAILED|REJECTED|作成できません|確認が必要/.test(text)) return '#b3261e';
  if (/COMPLETED|利用できます|OFFICIAL|正式/.test(text)) return '#137333';
  if (/PENDING|VALIDATING|CREATING|受付済み|確認中|作成中/.test(text)) return '#1a73e8';
  return '#3c4043';
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
