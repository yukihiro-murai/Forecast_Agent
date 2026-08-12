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
    try { vNextPortalRefreshViews_(); }
    catch (refreshError) { vNextPortalLog_('onOpen refresh skipped', refreshError); }
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
  } catch (error) {
    vNextPortalLog_('vNextPortalGoHome failed', error);
    vNextPortalShowError_('ホームを更新できませんでした。', error);
  }
}

function vNextPortalOpenCreateSidebar() {
  try {
    vNextPortalEnsureStructure_(SpreadsheetApp.getActiveSpreadsheet());
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
  return { ok: true, years: data.years, requestCount: data.requests.length, directoryCount: data.directory.length };
}

function vNextPortalRenderHome_(sheet, data) {
  vNextPortalResetViewSheet_(sheet, 44, 10);
  sheet.getRange('A1:J2').merge().setValue('年度計画ポータル')
    .setFontSize(20).setFontWeight('bold').setFontColor('#ffffff').setBackground('#174ea6');
  sheet.getRange('A4:J4').merge().setValue('クライアント別の年度計画を探す・作るための共通入口です。誰でも必要な計画を確認できます。')
    .setFontSize(11).setWrap(true);
  sheet.getRange('A6:J8').merge().setValue('新しい計画が必要ですか？\n上部メニュー「年度計画ポータル」→「新しい年度計画を作る」を選んでください。')
    .setFontSize(13).setFontWeight('bold').setFontColor('#174ea6').setBackground('#e8f0fe')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true)
    .setNote('既存ブックの重複を防ぐため、作成前に同じ年度の候補を自動確認します。');

  sheet.getRange('A10:J10').merge().setValue('作成依頼の状況').setFontWeight('bold').setBackground('#f1f3f4');
  var headers = ['受付日時', '年度', 'クライアント', '状態', '次の案内', 'Forecast Owner', '依頼者', '更新日時', '受付番号', '開く'];
  sheet.getRange(11, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#f8f9fa');
  sheet.getRange('D11').setNote('受付済み→内容確認中→ブック作成中→利用できます、の順で進みます。');
  var requests = data.requests.slice().sort(function (a, b) {
    return String(b.updatedAt || b.requestedAt).localeCompare(String(a.updatedAt || a.requestedAt));
  }).slice(0, 20);
  if (!requests.length) {
    sheet.getRange('A12:J14').merge().setValue('作成依頼はまだありません。既存の計画は下部の年度リンクから確認できます。')
      .setFontColor('#5f6368').setHorizontalAlignment('center').setVerticalAlignment('middle');
  } else {
    var rows = requests.map(function (request) {
      return [
        vNextPortalDisplayDateTime_(request.requestedAt), 'FY' + request.fiscalYear,
        vNextPortalCellText_(request.clientName), request.statusLabel,
        vNextPortalStatusNextAction_(request.status, request.detailMessage, Boolean(request.url)),
        request.forecastOwnerEmail, request.requestedBy, vNextPortalDisplayDateTime_(request.updatedAt),
        request.requestId, request.url ? '開く' : '準備中'
      ];
    });
    sheet.getRange(12, 1, rows.length, headers.length).setValues(rows).setWrap(true).setVerticalAlignment('middle');
    requests.forEach(function (request, index) {
      if (request.url) vNextPortalWriteLink_(sheet.getRange(12 + index, 10), request.url, '開く');
    });
    sheet.getRange(12, 1, rows.length, 1).setNumberFormat('@');
  }

  var sectionRow = Math.max(17, 13 + requests.length);
  sheet.getRange(sectionRow, 1, 1, 10).merge().setValue('年度別の計画一覧').setFontWeight('bold').setBackground('#f1f3f4');
  sheet.getRange(sectionRow + 1, 1, 2, 10).merge().setValue(
    '下部のFYタブを開くと、その年度の全クライアントを確認できます。クライアント名は Ctrl+F（Macは⌘+F）または表のフィルタで探せます。\n' +
    '表示セルは自動更新されます。入力や修正は、各クライアントの専用ブックで行ってください。'
  ).setWrap(true).setVerticalAlignment('middle');
  var yearText = data.years.map(function (year) { return 'FY' + year; }).join(' ／ ');
  sheet.getRange(sectionRow + 4, 1, 1, 10).merge().setValue('現在の年度タブ：' + yearText)
    .setFontWeight('bold').setFontColor('#188038').setHorizontalAlignment('center');
  sheet.setFrozenRows(2);
  sheet.setHiddenGridlines(true);
  [118, 72, 170, 110, 270, 180, 180, 118, 220, 72].forEach(function (width, index) { sheet.setColumnWidth(index + 1, width); });
  vNextPortalProtectWarningOnly_(sheet, '自動生成画面（入力は上部メニューから行います）');
}

function vNextPortalRenderFiscalYear_(sheet, fiscalYear, data) {
  var entries = vNextPortalFiscalYearEntries_(fiscalYear, data);
  vNextPortalResetViewSheet_(sheet, Math.max(30, entries.length + 10), 10);
  sheet.getRange('A1:J2').merge().setValue('FY' + fiscalYear + ' クライアント別年度計画')
    .setFontSize(18).setFontWeight('bold').setFontColor('#ffffff').setBackground('#188038');
  sheet.getRange('A4:J4').merge().setValue('クライアントを探し、「開く」から専用ブックへ進んでください。見つからない場合は上部メニューから新しく作成できます。')
    .setWrap(true);
  var headers = ['状態', 'クライアント', '中心見込み', '採用予測', '最終予算', 'Forecast Owner', '関与メンバー', '次の対応', '更新日', '開く'];
  sheet.getRange(6, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setFontColor('#ffffff').setBackground('#5f6368');
  sheet.getRange('C6').setNote('システムが算出した条件付き予測の中心値です。営業目標ではありません。');
  sheet.getRange('D6').setNote('予測を確認したうえで、計画に採用した金額です。');
  sheet.getRange('E6').setNote('採用予測に営業上積みを加えた正式な予算です。');
  sheet.getRange('F6').setNote('アクセス権限ではなく、計画案をまとめて提出する担当者です。');
  if (!entries.length) {
    sheet.getRange('A7:J10').merge().setValue('この年度の計画はまだありません。\n上部メニュー「年度計画ポータル」→「新しい年度計画を作る」から依頼できます。')
      .setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('middle').setFontColor('#5f6368');
  } else {
    var rows = entries.map(function (entry) {
      return [
        entry.statusLabel,
        vNextPortalCellText_(entry.clientName),
        entry.centerForecast,
        entry.adoptedForecast,
        entry.finalBudget,
        entry.forecastOwnerEmail,
        entry.relatedMemberEmails.join('\n'),
        entry.nextAction,
        vNextPortalDisplayDate_(entry.updatedAt),
        entry.url ? '開く' : '準備中'
      ];
    });
    sheet.getRange(7, 1, rows.length, headers.length).setValues(rows).setWrap(true).setVerticalAlignment('middle');
    sheet.getRange(7, 3, rows.length, 3).setNumberFormat('¥#,##0;[Red]-¥#,##0;―');
    entries.forEach(function (entry, index) {
      if (entry.url) vNextPortalWriteLink_(sheet.getRange(7 + index, 10), entry.url, '開く');
    });
    try {
      var existingFilter = sheet.getFilter();
      if (existingFilter) existingFilter.remove();
      sheet.getRange(6, 1, rows.length + 1, headers.length).createFilter();
    } catch (error) {
      vNextPortalLog_('FY filter creation skipped', error);
    }
  }
  sheet.setFrozenRows(6);
  sheet.setHiddenGridlines(true);
  [118, 190, 118, 118, 118, 190, 210, 270, 108, 72].forEach(function (width, index) { sheet.setColumnWidth(index + 1, width); });
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
  var maxRows = sheet.getMaxRows();
  var maxColumns = sheet.getMaxColumns();
  sheet.getRange(1, 1, maxRows, maxColumns).breakApart();
  var filter = sheet.getFilter();
  if (filter) filter.remove();
  sheet.clear();
  sheet.getRange(1, 1, rows, columns).setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
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
