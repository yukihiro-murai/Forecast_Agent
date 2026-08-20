/**
 * Forecast vNext 従業員UX。
 * Client FY Bookに、状態駆動の最小メニュー・3画面・入力APIを提供する。
 */

var VNEXT_UX_CONFIG_ = Object.freeze({
  MENU_NAME: '年度予算策定',
  HOME_SHEET: '1_ホーム',
  PLAN_SHEET: '2_予測と計画',
  REVIEW_SHEET: '3_振り返り',
  ANALYTICS_SHEET: 'VN_ANALYTICS_FACT',
  META_SHEET: 'BOOK_META',
  CONFIG_SHEET: 'VN_BOOK_CONFIG',
  INPUT_HTML: 'VNext_InputSidebar',
  GUIDANCE_HTML: 'VNext_GuidanceSidebar',
  HELP_HTML: 'VNext_HelpSidebar',
  INPUT_STATES: ['INPUT_OPEN', 'READY_TO_RUN'],
  REVIEW_STATES: ['REVIEW_DUE', 'YEAR_CLOSED'],
  MAX_TEXT: 1000,
  MAX_TARGET: 120,
  MAX_AMOUNT: 1000000000000000
});

var VNEXT_UX_STATE_LABELS_ = Object.freeze({
  INPUT_OPEN: '情報入力中',
  READY_TO_RUN: '予測を実行できます',
  RUNNING: '予測を作成中',
  DRAFT_READY: '予測案ができました',
  SUBMITTED: '管理ハブの確認待ち',
  CHANGES_REQUESTED: '修正依頼があります',
  OFFICIAL_LOCKED: '正式予算',
  REVIEW_DUE: '振り返り期間',
  YEAR_CLOSED: '年度終了'
});

var VNEXT_UX_REVIEW_CAUSES_ = Object.freeze([
  Object.freeze({ key: 'BASE_LEVEL', label: '通常の売上水準', field: 'base_level_error', annual: true }),
  Object.freeze({ key: 'UNKNOWN_SPOT', label: '未把握の単発売上', field: 'unknown_spot_error', annual: true }),
  Object.freeze({ key: 'COMMITMENT_OUTCOME', label: '契約・案件の成否', field: 'commitment_outcome_error', annual: true }),
  Object.freeze({ key: 'AMOUNT', label: '金額見立て・参照情報', field: 'amount_error', annual: true }),
  Object.freeze({ key: 'HUMAN_INFO', label: '現場情報', field: 'human_info_error', annual: true }),
  Object.freeze({ key: 'AI_INFO', label: 'AI・外部情報', field: 'ai_info_error', annual: true }),
  Object.freeze({ key: 'DATA_QUALITY', label: 'データ不足・不備', field: 'data_quality_error', annual: true }),
  Object.freeze({ key: 'SEASONALITY', label: '季節配分', field: 'seasonality_error', annual: false }),
  Object.freeze({ key: 'TIMING', label: '売上認識月のずれ', field: 'timing_error', annual: false }),
  Object.freeze({ key: 'OTHER', label: 'その他', field: '', annual: false })
]);

/** Rebuildable, column-oriented projection for BI tools and human inspection. */
var VNEXT_UX_ANALYTICS_HEADERS_ = Object.freeze([
  'schema_version', 'snapshot_id', 'snapshot_hash', 'book_id', 'client_id',
  'client_name', 'fiscal_year', 'run_id', 'as_of', 'cutoff', 'state',
  'record_type', 'category', 'item_key', 'sequence_no', 'period_type',
  'period_key', 'label', 'description', 'amount_yen', 'p10_yen', 'p50_yen',
  'p90_yen', 'direction', 'forecast_use', 'source_url', 'source_date',
  'evidence_quality', 'related_record_id', 'generated_at', 'source_table'
]);

/** legacy onOpenから呼ぶ安全なrouter。Admin 管理ハブのmenu builderは別moduleへ委譲する。 */
function vNextHandleOnOpen_() {
  try {
    return vNextBuildClientMenu_();
  } catch (error) {
    Logger.log('vNextHandleOnOpen_ error: ' + vNextUxErrorText_(error));
    return false;
  }
}

/** Client FY Bookに復旧用の最小メニューだけを表示する。 */
function vNextBuildClientMenu_() {
  try {
    // Simple onOpen triggers may not expose user identity and cannot open the sidebar.
    // Daily work lives in the guidance sidebar; this menu is only a recovery path.
    SpreadsheetApp.getUi().createMenu(VNEXT_UX_CONFIG_.MENU_NAME)
      .addItem('案内を開く', 'vNextGoHomeAndShowGuidance')
      .addToUi();
    return true;
  } catch (err) {
    Logger.log('vNextBuildClientMenu_ error: ' + vNextUxErrorText_(err));
    return false;
  }
}

/** Master Templateは現場入力の対象ではない。案内だけ出す。 */
function vNextOpenTemplateGuidance() {
  try {
    SpreadsheetApp.getUi().alert(
      '年度予算策定',
      'このクライアント年度ブックは生成用のひな型です。現場のクライアント年度ブックは申請入口から作成してください。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (err) {
    Logger.log('vNextOpenTemplateGuidance error: ' + vNextUxErrorText_(err));
  }
}

/** 非破壊で3画面を用意し、内部sheetを従業員の通常導線から隠す。 */
function vNextSetupClientExperience_() {
  try {
    var context = vNextUxGetBookContext_();
    vNextUxAssertClientBook_(context);
    vNextUxEnsureClientSheets_(context);
    vNextRefreshEmployeeViews();
    return { ok: true };
  } catch (err) {
    Logger.log('vNextSetupClientExperience_ error: ' + vNextUxErrorText_(err));
    throw new Error('クライアント年度ブック画面を準備できませんでした。管理ハブ担当者へ連絡してください。');
  }
}

function vNextGoHome() {
  try {
    vNextRefreshEmployeeViews();
    vNextUxActivateSheet_(VNEXT_UX_CONFIG_.HOME_SHEET, 'A1');
  } catch (err) {
    Logger.log('vNextGoHome error: ' + vNextUxErrorText_(err));
    vNextUxAlertError_('ホームを開けませんでした。', err);
  }
}

function vNextGoHomeAndShowGuidance() {
  try {
    var context = vNextUxGetBookContext_();
    vNextRefreshEmployeeViews(context);
    vNextUxActivateSheet_(VNEXT_UX_CONFIG_.HOME_SHEET, 'A1');
    vNextUxAutoOpenGuidance_(context, true);
  } catch (err) {
    Logger.log('vNextGoHomeAndShowGuidance error: ' + vNextUxErrorText_(err));
    vNextUxAlertError_('ホームと案内を開けませんでした。', err);
  }
}

/** 案内の自動表示を、このクライアント年度ブックのproject triggerとして有効化する。 */
function vNextInstallAutomaticGuidance() {
  try {
    var installed = vNextUxEnsureGuidanceOnOpenTrigger_();
    vNextOpenGuidanceSidebar();
    return { ok: true, installed: installed, message: '次回からブックを開くと案内が自動表示されます。' };
  } catch (error) {
    Logger.log('vNextInstallAutomaticGuidance error: ' + vNextUxErrorText_(error));
    vNextUxAlertError_('自動案内を有効にできませんでした。', error);
    return { ok: false };
  }
}

/** Installable open triggerからだけ呼ぶ。simple onOpenからは呼ばない。 */
function vNextInstalledGuidanceOnOpen(e) {
  return vNextUxOpenGuidanceShellQuietly_();
}

function vNextUxEnsureGuidanceOnOpenTrigger_() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var handler = 'vNextInstalledGuidanceOnOpen';
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
    Logger.log('vNextUxEnsureGuidanceOnOpenTrigger_ cleanup skipped: ' + vNextUxErrorText_(cleanupError));
  }
  ScriptApp.newTrigger(handler).forSpreadsheet(spreadsheet).onOpen().create();
  return true;
}

function vNextOpenGuidanceSidebar() {
  try {
    if (!vNextUxOpenGuidanceShellQuietly_()) throw new Error('案内を表示できませんでした。');
  } catch (err) {
    Logger.log('vNextOpenGuidanceSidebar error: ' + vNextUxErrorText_(err));
    vNextUxAlertError_('案内を開けませんでした。', err);
  }
}

/** サイドバーと同じ安全なview modelを、広いBI風画面で表示する。 */
function vNextOpenForecastDashboard() {
  try {
    var html = HtmlService.createTemplateFromFile(VNEXT_UX_CONFIG_.GUIDANCE_HTML).evaluate()
      .setTitle('予測ダッシュボード')
      .setWidth(860)
      .setHeight(780);
    SpreadsheetApp.getUi().showModelessDialog(html, '予測ダッシュボード');
  } catch (err) {
    Logger.log('vNextOpenForecastDashboard error: ' + vNextUxErrorText_(err));
    vNextUxAlertError_('予測ダッシュボードを開けませんでした。', err);
  }
}

function vNextUxOpenGuidanceShellQuietly_() {
  try {
    var html = HtmlService.createTemplateFromFile(VNEXT_UX_CONFIG_.GUIDANCE_HTML).evaluate()
      .setTitle('次にすること')
      .setWidth(400);
    SpreadsheetApp.getUi().showSidebar(html);
    try { vNextUxEnsureGuidanceOnOpenTrigger_(); }
    catch (triggerError) {
      Logger.log('vNextUxOpenGuidanceShellQuietly_ trigger skipped: ' + vNextUxErrorText_(triggerError));
    }
    return true;
  } catch (error) {
    Logger.log('vNextUxOpenGuidanceShellQuietly_ skipped: ' + vNextUxErrorText_(error));
    return false;
  }
}

function vNextUxAutoOpenGuidance_(context, forceGuidance) {
  try {
    // 最初は常に1つの案内だけを見せ、質問を突然並べない。
    vNextOpenGuidanceSidebar();
    return true;
  } catch (error) {
    Logger.log('vNextUxAutoOpenGuidance_ skipped: ' + vNextUxErrorText_(error));
    return false;
  }
}

function vNextGoForecastPlan() {
  try {
    vNextRefreshEmployeeViews();
    vNextUxActivateSheet_(VNEXT_UX_CONFIG_.PLAN_SHEET, 'A1');
  } catch (err) {
    Logger.log('vNextGoForecastPlan error: ' + vNextUxErrorText_(err));
    vNextUxAlertError_('予測シートを開けませんでした。', err);
  }
}

/** 案内サイドバーの主操作から、状態に応じた予測・計画・振り返り画面へ到達させる。 */
function vNextOpenForecastPlanOrCurrentAction() {
  return vNextOpenForecastDashboard();
}

function vNextGoReview() {
  try {
    vNextRefreshEmployeeViews();
    vNextUxActivateSheet_(VNEXT_UX_CONFIG_.REVIEW_SHEET, 'A1');
  } catch (err) {
    Logger.log('vNextGoReview error: ' + vNextUxErrorText_(err));
    vNextUxAlertError_('振り返りを開けませんでした。', err);
  }
}

function vNextOpenInputSidebar() {
  try {
    var model = vNextGetClientViewModel();
    if (!model.canInput) {
      SpreadsheetApp.getUi().alert('現在は入力期間ではありません', model.inputLockMessage, SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    var html = HtmlService.createTemplateFromFile(VNEXT_UX_CONFIG_.INPUT_HTML).evaluate()
      .setTitle('自分の情報を入力')
      .setWidth(420);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (err) {
    Logger.log('vNextOpenInputSidebar error: ' + vNextUxErrorText_(err));
    vNextUxAlertError_('入力画面を開けませんでした。', err);
  }
}

function vNextOpenHelpSidebar() {
  try {
    var html = HtmlService.createHtmlOutputFromFile(VNEXT_UX_CONFIG_.HELP_HTML)
      .setTitle('使い方・困ったとき')
      .setWidth(420);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (err) {
    Logger.log('vNextOpenHelpSidebar error: ' + vNextUxErrorText_(err));
    vNextUxAlertError_('ヘルプを開けませんでした。', err);
  }
}

/** Master Template上の1個のdrawing buttonへ割り当てる安定した関数名。 */
function vNextOpenCurrentAction() {
  try {
    var model = vNextGetClientViewModel();
    switch (model.primaryAction.key) {
      case 'INPUT': vNextOpenInputSidebar(); return;
      case 'REQUEST_FORECAST': vNextRequestForecast(); return;
      case 'EDIT_PLAN': vNextOpenPlanSidebar(); return;
      case 'REVIEW': vNextOpenReviewSidebar(); return;
      case 'VIEW_PLAN': vNextGoForecastPlan(); return;
      case 'VIEW_REVIEW': vNextGoReview(); return;
      default:
        SpreadsheetApp.getActiveSpreadsheet().toast(model.primaryAction.instruction, VNEXT_UX_CONFIG_.MENU_NAME, 6);
    }
  } catch (err) {
    Logger.log('vNextOpenCurrentAction error: ' + vNextUxErrorText_(err));
    vNextUxAlertError_('操作を開始できませんでした。', err);
  }
}

function vNextOpenReviewSidebar() {
  try {
    vNextUxAssertReviewEditable_(vNextUxGetBookContext_());
    var html = HtmlService.createTemplateFromFile('VNext_ReviewSidebar').evaluate()
      .setTitle('年度の振り返り')
      .setWidth(430);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (err) {
    Logger.log('vNextOpenReviewSidebar error: ' + vNextUxErrorText_(err));
    vNextUxAlertError_('振り返りを開けませんでした。', err);
  }
}

function vNextGetReviewEditorModel() {
  try {
    var context = vNextUxGetBookContext_();
    vNextUxAssertReviewEditable_(context);
    var evaluation = vNextUxGetLatestEvaluation_(context.bookId);
    var previous = vNextUxGetLatestOwnReview_(context);
    return {
      ok: true,
      clientName: context.clientName,
      fiscalYear: context.fiscalYear,
      evaluation: evaluation ? {
        evaluationId: String(evaluation.evaluation_id || ''),
        officialVintageId: String(evaluation.official_vintage_id || ''),
        actualTotal: Number(evaluation.actual_total || 0),
        systemForecast: Number(evaluation.system_forecast || 0),
        adoptedForecast: Number(evaluation.adopted_forecast || 0),
        finalBudget: Number(evaluation.final_budget || 0),
        systemSignedError: Number(evaluation.system_signed_error || 0),
        rangeContainsActual: vNextUxBool_(evaluation.range_contains_actual),
        breakdown: vNextUxEvaluationBreakdown_(evaluation)
      } : null,
      causeCategories: VNEXT_UX_REVIEW_CAUSES_.map(function(item) {
        return { key: item.key, label: item.label, annual: item.annual };
      }),
      previousReview: previous
    };
  } catch (err) {
    Logger.log('vNextGetReviewEditorModel error: ' + vNextUxErrorText_(err));
    throw new Error(vNextUxFriendlyValidationError_(err));
  }
}

function vNextPreviewReview(payload) {
  try {
    var context = vNextUxGetBookContext_();
    vNextUxAssertReviewEditable_(context);
    var review = vNextUxNormalizeReview_(payload, context);
    return {
      ok: true,
      previewHash: vNextUxSha256_(JSON.stringify(review)),
      summary: '原因カテゴリ ' + review.causeCategories.length + '件と、次年度に確認する情報 ' + review.nextInformation.length + '件を保存します。'
    };
  } catch (err) {
    Logger.log('vNextPreviewReview error: ' + vNextUxErrorText_(err));
    throw new Error(vNextUxFriendlyValidationError_(err));
  }
}

function vNextSaveReview(payload) {
  try {
    var context = vNextUxGetBookContext_();
    vNextUxAssertReviewEditable_(context);
    var review = vNextUxNormalizeReview_(payload, context);
    var expectedHash = vNextUxSha256_(JSON.stringify(review));
    if (payload && payload.previewHash && String(payload.previewHash) !== expectedHash) {
      throw new Error('画面の内容が変更されています。入力内容を確認して、もう一度保存してください。');
    }
    if (typeof vNextAppendRecord_ !== 'function') throw new Error('振り返りの保存先が未設定です。管理ハブ担当者へ連絡してください。');
    var previous = vNextUxGetLatestOwnReviewRecord_(context);
    var evidenceId = typeof vNextUuid_ === 'function' ? vNextUuid_() : Utilities.getUuid();
    vNextAppendRecord_('EVIDENCE_EVENT', {
      evidence_id: evidenceId,
      book_id: context.bookId,
      client_id: context.clientId,
      fiscal_year: context.fiscalYear,
      actor_email: context.userEmail,
      response_type: 'UNKNOWN',
      evidence_type: 'REVIEW_LEARNING',
      target: 'FY_REVIEW',
      target_start_month: '', target_end_month: '', direction: 'NEUTRAL',
      amount_mode: '', amount_low: '', amount_mid: '', amount_high: '', amount_band: '', confidence_class: '',
      evidence_text: JSON.stringify(review),
      source_url: '', source_date: '', expires_at: '', status: 'ACTIVE',
      supersedes_evidence_id: previous && previous.evidence_id || '',
      created_at: new Date().toISOString()
    }, { spreadsheet: SpreadsheetApp.getActiveSpreadsheet() });
    return { ok: true, evidenceId: evidenceId, message: '振り返りを保存しました。次年度の情報収集とモデル改善に活用します。' };
  } catch (err) {
    Logger.log('vNextSaveReview error: ' + vNextUxErrorText_(err));
    throw new Error(vNextUxFriendlyValidationError_(err));
  }
}

function vNextOpenPlanSidebar() {
  try {
    vNextUxAssertPlanEditable_(vNextUxGetBookContext_());
    var html = HtmlService.createTemplateFromFile('VNext_PlanSidebar').evaluate()
      .setTitle('予算案を作成・提出')
      .setWidth(440);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (err) {
    Logger.log('vNextOpenPlanSidebar error: ' + vNextUxErrorText_(err));
    vNextUxAlertError_('計画編集を開けませんでした。', err);
  }
}

function vNextGetPlanEditorModel() {
  try {
    var context = vNextUxGetBookContext_();
    vNextUxAssertPlanEditable_(context);
    var rawForecast = vNextUxGetLatestForecast_(context);
    if (!rawForecast) throw new Error('提出に使用できる予測がありません。');
    var forecast = vNextUxForecastForView_(context, rawForecast);
    var plan = vNextUxGetLatestPlan_(context.bookId);
    var allocation = vNextUxParseJsonArray_(plan && plan.uplift_allocation_json);
    var monthValues = vNextUxForecastMonthValues_(forecast.months);
    var totalMonths = monthValues.reduce(function(sum, value) { return sum + value; }, 0);
    var monthWeights = monthValues.map(function(value) { return totalMonths > 0 ? value / totalMonths : 1 / 12; });
    return {
      ok: true,
      clientName: context.clientName,
      fiscalYear: context.fiscalYear,
      state: context.state,
      stateLabel: VNEXT_UX_STATE_LABELS_[context.state] || context.state,
      runId: forecast.runId,
      systemRecommended: forecast.systemRecommended,
      range: { p10: forecast.p10, p90: forecast.p90 },
      evidenceCoverage: forecast.evidenceCoverage,
      layerBreakdown: forecast.layerBreakdown,
      aiEvidence: forecast.aiEvidence,
      warnings: forecast.warnings,
      monthLabels: ['4月','5月','6月','7月','8月','9月','10月','11月','12月','1月','2月','3月'],
      monthWeights: monthWeights,
      changeRequestReason: vNextUxGetLatestChangeRequestReason_(context.bookId),
      previousPlan: plan ? {
        planVersionId: String(plan.plan_version_id || ''),
        adoptionDelta: Number(plan.adoption_delta || 0),
        adoptionReason: String(plan.adoption_reason || ''),
        salesUplift: Number(plan.sales_uplift || 0),
        upliftReason: String(plan.uplift_reason || ''),
        upliftOwner: String(plan.uplift_owner || ''),
        upliftAction: String(plan.uplift_action || ''),
        upliftDueDate: vNextUxDateText_(plan.uplift_due_date),
        monthAllocation: vNextUxMonthlyAllocationValues_(allocation)
      } : null
    };
  } catch (err) {
    Logger.log('vNextGetPlanEditorModel error: ' + vNextUxErrorText_(err));
    throw new Error(vNextUxFriendlyValidationError_(err));
  }
}

function vNextPreviewPlan(payload) {
  try {
    var context = vNextUxGetBookContext_();
    vNextUxAssertPlanEditable_(context);
    var forecast = vNextUxPublicForecast_(vNextUxGetLatestForecast_(context));
    var plan = vNextUxNormalizePlan_(payload, context, forecast);
    return {
      ok: true,
      previewHash: vNextUxSha256_(JSON.stringify(plan.canonical)),
      adoptedForecast: plan.adoptedForecast,
      finalBudget: plan.finalBudget,
      quarterAllocation: plan.quarterAllocation,
      summary: 'システム推奨予測 ' + vNextUxFormatMoney_(forecast.systemRecommended) + '、採用予測 ' + vNextUxFormatMoney_(plan.adoptedForecast) + '、最終予算 ' + vNextUxFormatMoney_(plan.finalBudget) + 'として提出します。'
    };
  } catch (err) {
    Logger.log('vNextPreviewPlan error: ' + vNextUxErrorText_(err));
    throw new Error(vNextUxFriendlyValidationError_(err));
  }
}

function vNextSubmitPlan(payload) {
  try {
    var context = vNextUxGetBookContext_();
    vNextUxAssertPlanEditable_(context);
    var forecast = vNextUxPublicForecast_(vNextUxGetLatestForecast_(context));
    var plan = vNextUxNormalizePlan_(payload, context, forecast);
    var expectedHash = vNextUxSha256_(JSON.stringify(plan.canonical));
    if (payload && payload.previewHash && String(payload.previewHash) !== expectedHash) {
      throw new Error('画面の内容が変更されています。金額と理由を確認して、もう一度提出してください。');
    }
    if (typeof vNextAppendPlanVersion_ !== 'function') throw new Error('予算の保存機能が未設定です。管理ハブ担当者へ連絡してください。');
    var latest = vNextUxGetLatestPlan_(context.bookId);
    var record = vNextAppendPlanVersion_({
      bookId: context.bookId,
      runId: forecast.runId,
      versionNo: latest ? Number(latest.version_no || 0) + 1 : 1,
      status: 'SUBMITTED',
      systemRecommended: forecast.systemRecommended,
      adoptionDelta: plan.adoptionDelta,
      adoptionReason: plan.adoptionReason,
      salesUplift: plan.salesUplift,
      upliftReason: plan.upliftReason,
      upliftOwner: plan.upliftOwner,
      upliftAction: plan.upliftAction,
      upliftDueDate: plan.upliftDueDate,
      upliftAllocation: plan.upliftAllocation,
      amendsPlanVersionId: latest && latest.plan_version_id || '',
      submittedAt: new Date().toISOString(),
      submittedBy: context.userEmail
    }, { spreadsheet: SpreadsheetApp.getActiveSpreadsheet() });
    vNextUxTransition_({
      bookId: context.bookId,
      fromState: context.state,
      toState: 'SUBMITTED',
      reason: context.state === 'CHANGES_REQUESTED' ? 'revised_plan_submitted' : 'plan_submitted',
      actorEmail: context.userEmail,
      actorRole: vNextUxVerifiedOwnerRole_(context),
      relatedRunId: forecast.runId,
      relatedPlanVersionId: record.plan_version_id
    });
    vNextRefreshEmployeeViews();
    return { ok: true, planVersionId: record.plan_version_id, message: '予算案を提出しました。管理ハブの確認をお待ちください。' };
  } catch (err) {
    Logger.log('vNextSubmitPlan error: ' + vNextUxErrorText_(err));
    throw new Error(vNextUxFriendlyValidationError_(err));
  }
}

/** Sidebarとsheet viewが共通利用する、機微情報を含めないview model。 */
function vNextGetClientViewModel() {
  try {
    return vNextBuildClientViewModel_(vNextUxGetBookContext_());
  } catch (err) {
    Logger.log('vNextGetClientViewModel error: ' + vNextUxErrorText_(err));
    throw new Error('画面情報を取得できませんでした。再読み込みしても直らない場合は管理ハブ担当者へ連絡してください。');
  }
}

/** 3画面を再描画してからview modelを1回で返す。Refresh操作のRPC往復を減らす。 */
function vNextRefreshEmployeeViewsAndGetClientViewModel() {
  try {
    var context = vNextUxGetBookContext_();
    vNextRefreshEmployeeViews(context);
    return vNextBuildClientViewModel_(context);
  } catch (err) {
    Logger.log('vNextRefreshEmployeeViewsAndGetClientViewModel error: ' + vNextUxErrorText_(err));
    throw new Error('画面情報を取得できませんでした。再読み込みしても直らない場合は管理ハブ担当者へ連絡してください。');
  }
}

function vNextBuildClientViewModel_(context) {
  vNextUxAssertClientBook_(context);
  var forecast = vNextUxGetLatestForecast_(context);
  var action = vNextUxGetPrimaryAction_(context);
  var bands = vNextUxBuildAmountBands_(forecast, context);
  return {
    ok: true,
    bookId: context.bookId || '',
    clientName: context.clientName || 'クライアント未設定',
    fiscalYear: context.fiscalYear || '',
    asOf: vNextUxDateText_(context.asOf),
    cutoff: vNextUxDateText_(context.cutoff),
    state: context.state,
    stateLabel: VNEXT_UX_STATE_LABELS_[context.state] || context.state,
    roleLabel: vNextUxRoleLabel_(context),
    isForecastOwner: Boolean(context.isForecastOwner),
    isTeamMember: Boolean(context.isTeamMember),
    isInternalUser: Boolean(context.isInternalUser),
    canContribute: Boolean(context.canContribute),
    canInput: Boolean(context.canContribute) && VNEXT_UX_CONFIG_.INPUT_STATES.indexOf(context.state) >= 0,
    inputLockMessage: vNextUxInputLockMessage_(context.state),
    inputStatus: context.inputStatus || {},
    canOverrideInput: vNextUxCanOverrideInput_(context),
    latestOwnEvidence: context.latestOwnEvidence || null,
    amountBands: bands,
    amountBandBasis: bands.length && bands[0].available
      ? 'このクライアントの年間売上基準 ' + vNextUxFormatMoney_(bands[0].basisAmount)
      : '年間売上基準を確認できないため、金額入力を選んでください。',
    primaryAction: action,
    stateNotice: vNextUxStateIssue_(context),
    forecast: vNextUxForecastForView_(context, forecast),
    version: context.version || ''
  };
}

/** 入力内容を保存せず、金額影響と保存内容を確認する。 */
function vNextPreviewEvidence(payload) {
  try {
    var context = vNextUxGetBookContext_();
    vNextUxAssertInputAllowed_(context);
    var normalized = vNextUxNormalizeEvidence_(payload);
    var forecast = vNextUxGetLatestForecast_(context);
    var bands = vNextUxBuildAmountBands_(forecast, context);
    var impact = vNextUxCalculateImpact_(normalized, bands);
    var canonical = vNextUxCanonicalEvidence_(normalized);
    return {
      ok: true,
      previewHash: vNextUxSha256_(JSON.stringify(canonical)),
      summary: vNextUxEvidenceSummary_(normalized, impact),
      impactLow: impact.low,
      impactHigh: impact.high,
      impactText: vNextUxImpactText_(impact.low, impact.high),
      warning: normalized.responseType === 'unknown' ? '情報不足として保存し、予測の振れ幅に反映します。' : '',
      normalized: canonical
    };
  } catch (err) {
    Logger.log('vNextPreviewEvidence error: ' + vNextUxErrorText_(err));
    throw new Error(vNextUxFriendlyValidationError_(err));
  }
}

/** 入力をserver側で検証し、append-only evidenceとして1回で保存する。 */
function vNextSaveEvidence(payload) {
  try {
    var context = vNextUxGetBookContext_();
    vNextUxAssertInputAllowed_(context);
    var normalized = vNextUxNormalizeEvidence_(payload);
    var canonical = vNextUxCanonicalEvidence_(normalized);
    if (typeof vNextAppendEvidence_ !== 'function') throw new Error('保存先が未設定です。管理ハブ担当者へ連絡してください。');
    var bands = vNextUxBuildAmountBands_(vNextUxGetLatestForecast_(context), context);
    var impact = vNextUxCalculateImpact_(normalized, bands);
    var record = vNextUxBuildEvidenceSaveRecord_(canonical, context, impact);
    var saved = vNextAppendEvidence_(record, { spreadsheet: SpreadsheetApp.getActiveSpreadsheet() });
    vNextUxMaybeAdvanceInputState_();
    vNextRefreshEmployeeViews();
    return {
      ok: true,
      evidenceId: saved && (saved.evidenceId || saved.id) || '',
      message: '入力を保存しました。ご協力ありがとうございます。'
    };
  } catch (err) {
    Logger.log('vNextSaveEvidence error: ' + vNextUxErrorText_(err));
    throw new Error(vNextUxFriendlyValidationError_(err));
  }
}

/** 期限後、予算策定担当が未回答を理由付きで締め切る。 */
function vNextCloseInputAndProceed(reason) {
  try {
    var context = vNextUxGetBookContext_();
    if (!vNextUxCanOverrideInput_(context)) throw new Error('この操作は、回答期限後の予算策定担当だけが実行できます。');
    var safeReason = vNextUxSafeText_(reason, 500, true, '進める理由');
    vNextUxTransition_({
      bookId: context.bookId,
      fromState: 'INPUT_OPEN',
      toState: 'READY_TO_RUN',
      reason: 'deadline_override: ' + safeReason,
      actorEmail: context.userEmail
    });
    vNextRefreshEmployeeViews();
    return { ok: true, message: '入力を締め切り、予測を依頼できる状態にしました。' };
  } catch (err) {
    Logger.log('vNextCloseInputAndProceed error: ' + vNextUxErrorText_(err));
    throw new Error(vNextUxFriendlyValidationError_(err));
  }
}

/** 予算策定担当だけがREADY_TO_RUNから予測runを開始できる。 */
function vNextRequestForecast() {
  try {
    var context = vNextUxGetBookContext_();
    if (!context.isForecastOwner) throw new Error('予測の依頼は予算策定担当が行います。');
    if (context.state !== 'READY_TO_RUN') throw new Error('現在は予測を依頼できる状態ではありません。');
    var ui = SpreadsheetApp.getUi();
    var answer = ui.alert('予測を依頼しますか？', '前月末までの確定実績と、保存済みの情報を使って予測を作成します。', ui.ButtonSet.OK_CANCEL);
    if (answer !== ui.Button.OK) return { ok: false, cancelled: true };
    var input = context.inputStatus || {};
    var total = Number(input.totalCount || 0);
    var answered = Number(input.answeredCount || 0);
    var request = {
      bookId: context.bookId,
      clientId: context.clientId,
      clientName: context.clientName,
      fiscalYear: context.fiscalYear,
      asOf: context.asOf,
      createdBy: context.userEmail,
      requestedBy: context.userEmail,
      missingResponseRate: total > 0 ? Math.max(0, total - answered) / total : 0
    };
    request.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    vNextEngineRunForecast(request);
    vNextRefreshEmployeeViews();
    SpreadsheetApp.getActiveSpreadsheet().toast('予測の作成を開始しました。完了後、この画面に結果が表示されます。', VNEXT_UX_CONFIG_.MENU_NAME, 8);
    return { ok: true };
  } catch (err) {
    Logger.log('vNextRequestForecast error: ' + vNextUxErrorText_(err));
    vNextUxAlertError_('予測を依頼できませんでした。', err);
    return { ok: false, error: vNextUxErrorText_(err) };
  }
}

/** 従業員向け3画面を最新の正本レコードから再描画する。 */
function vNextRefreshEmployeeViews(optionalContext) {
  try {
    var context = optionalContext || vNextUxGetBookContext_();
    vNextUxAssertClientBook_(context);
    var forecast = vNextUxGetLatestForecast_(context);
    vNextUxEnsureClientSheets_(context);
    var finalReadOnly = context.state === 'YEAR_CLOSED';
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var home = ss.getSheetByName(VNEXT_UX_CONFIG_.HOME_SHEET);
    var plan = ss.getSheetByName(VNEXT_UX_CONFIG_.PLAN_SHEET);
    var review = ss.getSheetByName(VNEXT_UX_CONFIG_.REVIEW_SHEET);
    // YEAR_CLOSEDの初回表示だけ最終状態を描画し、以後はhard protectionを
    // 解除せず読み取る。これにより閲覧APIは使える一方、確定表示を再生成できない。
    if (!home || !plan || !review) throw new Error('従業員向け表示シートを準備できませんでした。');
    // HTML sidebar callbacks can lose SpreadsheetApp's active-spreadsheet handle while a
    // long server operation is running. Reuse the sheet handles resolved above instead of
    // calling getActiveSpreadsheet() again inside each renderer.
    if (!finalReadOnly || !vNextUxIsHardProtected_(home)) vNextUxRenderHome_(context, home);
    if (!finalReadOnly || !vNextUxIsHardProtected_(plan)) vNextUxRenderPlan_(context, forecast, plan);
    if (!finalReadOnly || !vNextUxIsHardProtected_(review)) vNextUxRenderReview_(context, review);
    vNextUxRefreshAnalyticsFact_(context, forecast);
    return { ok: true };
  } catch (err) {
    Logger.log('vNextRefreshEmployeeViews error: ' + vNextUxErrorText_(err));
    throw err;
  }
}

/**
 * FORECAST_RUNの監査正本は変更せず、外部分析しやすい1行1指標の現在投影を作る。
 * この表は再生成可能であり、監査正本ではない。金額は表示慣習に合わせ整数円。
 */
function vNextUxRefreshAnalyticsFact_(context, rawForecast) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(VNEXT_UX_CONFIG_.ANALYTICS_SHEET) || ss.insertSheet(VNEXT_UX_CONFIG_.ANALYTICS_SHEET);
    var forecast = rawForecast ? vNextUxForecastForView_(context, rawForecast) : null;
    var plan = vNextUxGetLatestPlan_(context.bookId);
    var generatedAt = new Date().toISOString();
    var snapshotSource = {
      schemaVersion: 'vnext-analytics-fact-1', bookId: context.bookId, runId: forecast && forecast.runId || '',
      state: context.state, planVersionId: plan && plan.plan_version_id || '', generatedAt: generatedAt
    };
    var snapshotHash = vNextUxSha256_(JSON.stringify(snapshotSource));
    var base = {
      schema_version: 'vnext-analytics-fact-1', snapshot_id: 'SNAP-' + snapshotHash.slice(0, 20),
      snapshot_hash: snapshotHash, book_id: context.bookId, client_id: context.clientId,
      client_name: context.clientName, fiscal_year: Number(context.fiscalYear || 0),
      run_id: forecast && forecast.runId || '', as_of: vNextUxDateOnlyText_(context.asOf),
      cutoff: vNextUxDateOnlyText_(context.cutoff), state: context.state,
      generated_at: generatedAt
    };
    var facts = [];
    function add(values) { facts.push(Object.assign({}, base, values || {})); }
    if (forecast && forecast.runId) {
      add({ record_type: 'FORECAST_SUMMARY', category: 'ANNUAL_RANGE', item_key: 'SYSTEM_RECOMMENDED',
        sequence_no: 1, period_type: 'FY', period_key: 'FY' + context.fiscalYear,
        label: 'システム推奨予測', description: '通常の振れ幅を伴う年度中心見込み',
        amount_yen: forecast.systemRecommended, p10_yen: forecast.p10, p50_yen: forecast.center,
        p90_yen: forecast.p90, source_table: 'FORECAST_RUN', related_record_id: forecast.runId });
      [
        ['CONTINUITY', '継続性レンズ', '確定実績の水準・成長・季節性', forecast.historyBaseline],
        ['OBJECTIVE', '客観変化レンズ', '契約・参照情報・確認可能な外部情報まで', forecast.objectiveForecast],
        ['INTEGRATED', '統合レンズ', '担当者情報と予測へ採用したAI情報まで統合', forecast.systemRecommended]
      ].forEach(function(item, index) {
        add({ record_type: 'TRIANGULATION', category: 'FORECAST_LENS', item_key: item[0], sequence_no: index + 1,
          period_type: 'FY', period_key: 'FY' + context.fiscalYear, label: item[1], description: item[2],
          amount_yen: vNextUxWholeYen_(item[3]), source_table: 'FORECAST_RUN', related_record_id: forecast.runId });
      });
      (forecast.layerBreakdown && forecast.layerBreakdown.rows || []).forEach(function(item, index) {
        add({ record_type: 'FORECAST_COMPONENT', category: item.kind || 'DELTA', item_key: item.key,
          sequence_no: index + 1, period_type: 'FY', period_key: 'FY' + context.fiscalYear,
          label: item.label, description: item.description, amount_yen: vNextUxWholeYen_(item.amount),
          direction: Number(item.amount || 0) < 0 ? 'DOWN' : Number(item.amount || 0) > 0 ? 'UP' : 'NEUTRAL',
          source_table: 'FORECAST_RUN', related_record_id: forecast.runId });
      });
      (forecast.quarters || []).forEach(function(item, index) {
        add({ record_type: 'FORECAST_PERIOD', category: 'QUARTER', item_key: String(item.quarter || 'Q' + (index + 1)),
          sequence_no: index + 1, period_type: 'QUARTER', period_key: String(item.quarter || 'Q' + (index + 1)),
          label: String(item.quarter || 'Q' + (index + 1)), p10_yen: vNextUxWholeYen_(item.p10),
          p50_yen: vNextUxWholeYen_(item.p50), p90_yen: vNextUxWholeYen_(item.p90),
          amount_yen: vNextUxWholeYen_(item.p50), source_table: 'FORECAST_RUN', related_record_id: forecast.runId });
      });
      (forecast.months || []).forEach(function(item, index) {
        add({ record_type: 'FORECAST_PERIOD', category: 'MONTH', item_key: String(item.month || index + 1),
          sequence_no: index + 1, period_type: 'MONTH', period_key: String(item.month || ''),
          label: String(item.month || ''), p10_yen: vNextUxWholeYen_(item.p10),
          p50_yen: vNextUxWholeYen_(item.p50), p90_yen: vNextUxWholeYen_(item.p90),
          amount_yen: vNextUxWholeYen_(item.p50), source_table: 'FORECAST_RUN', related_record_id: forecast.runId });
      });
      (forecast.aiEvidence || []).forEach(function(item, index) {
        add({ record_type: 'AI_INSIGHT', category: String(item.researchAxis || 'EXTERNAL'),
          item_key: 'AI_INSIGHT_' + (index + 1), sequence_no: index + 1, label: String(item.target || '外部情報'),
          description: String(item.summary || ''), amount_yen: vNextUxWholeYen_(item.appliedAmount),
          direction: String(item.direction || 'NEUTRAL'), forecast_use: String(item.forecastUse || 'INSIGHT_ONLY'),
          source_url: String(item.sourceUrl || ''), source_date: String(item.sourceDate || ''),
          evidence_quality: String(item.evidenceQuality || ''), source_table: 'EVIDENCE_EVENT' });
      });
    }
    if (plan) {
      [
        ['ADOPTED_FORECAST', '採用予測', plan.adopted_forecast],
        ['SALES_UPLIFT', '営業上積み', plan.sales_uplift],
        ['FINAL_BUDGET', '最終予算', plan.final_budget]
      ].forEach(function(item, index) {
        add({ record_type: 'PLAN_SUMMARY', category: 'MANAGEMENT_DECISION', item_key: item[0], sequence_no: index + 1,
          period_type: 'FY', period_key: 'FY' + context.fiscalYear, label: item[1], amount_yen: vNextUxWholeYen_(item[2]),
          source_table: 'PLAN_VERSION', related_record_id: String(plan.plan_version_id || '') });
      });
    }
    add({ record_type: 'AUDIT_REFERENCE', category: 'REPRODUCIBILITY', item_key: 'RUN_AUDIT', sequence_no: 1,
      label: '計算監査と再現性', description: forecast && forecast.runId
        ? '正本には入力hash、実績基準日、seed、使用版、各層、Q/月、根拠IDを保存。同一入力・seed・版から再現します。全simulation pathは保存しません。'
        : '予測実行後に再現情報を記録します。',
      source_table: forecast && forecast.runId ? 'FORECAST_RUN' : 'BOOK_META', related_record_id: forecast && forecast.runId || '' });
    vNextUxWriteAnalyticsFact_(sheet, facts);
    try { sheet.hideSheet(); } catch (hideError) { Logger.log('VN_ANALYTICS_FACT hide warning: ' + vNextUxErrorText_(hideError)); }
    return { ok: true, rowCount: facts.length, snapshotId: base.snapshot_id };
  } catch (error) {
    Logger.log('vNextUxRefreshAnalyticsFact_ warning: ' + vNextUxErrorText_(error));
    return { ok: false, error: vNextUxErrorText_(error) };
  }
}

function vNextUxWriteAnalyticsFact_(sheet, facts) {
  var headers = VNEXT_UX_ANALYTICS_HEADERS_.slice();
  var rows = (facts || []).map(function(record) {
    return headers.map(function(header) {
      var value = record && record[header] !== undefined ? record[header] : '';
      if (typeof value === 'string' && /^[=+\-@]/.test(value)) return "'" + value;
      return value;
    });
  });
  var requiredRows = Math.max(2, rows.length + 1);
  var requiredColumns = headers.length;
  if (sheet.getMaxRows() < requiredRows) sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
  if (sheet.getMaxColumns() < requiredColumns) sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredColumns - sheet.getMaxColumns());
  var existingRows = Math.max(1, sheet.getLastRow());
  var existingColumns = Math.max(requiredColumns, sheet.getLastColumn());
  var filter = sheet.getFilter && sheet.getFilter();
  if (filter) filter.remove();
  sheet.getRange(1, 1, existingRows, existingColumns).clearContent().clearFormat().clearNote();
  sheet.getRange(1, 1, 1, requiredColumns).setValues([headers]).setFontWeight('bold')
    .setBackground('#e8eaed').setFontColor('#202124').setWrap(false);
  sheet.getRange(1, 1).setNote('外部分析向けの再生成可能な現在投影です。監査正本はFORECAST_RUN等の追記型テーブルです。金額は円未満を切り捨てています。');
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, requiredColumns).setNumberFormat('@').setValues(rows);
    [7, 15, 20, 21, 22, 23].forEach(function(column) {
      sheet.getRange(2, column, rows.length, 1).setNumberFormat('#,##0');
    });
    sheet.getRange(1, 1, rows.length + 1, requiredColumns).createFilter();
  }
  sheet.setFrozenRows(1);
  sheet.setHiddenGridlines(false);
  for (var column = 1; column <= requiredColumns; column++) sheet.setColumnWidth(column, 118);
  [6, 18, 19, 26].forEach(function(column) { sheet.setColumnWidth(column, column === 19 ? 360 : 220); });
  try {
    var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    var protection = protections.length ? protections[0] : sheet.protect();
    protection.setDescription('vNext分析投影（自動再生成）').setWarningOnly(true);
  } catch (protectionError) {
    Logger.log('VN_ANALYTICS_FACT protection warning: ' + vNextUxErrorText_(protectionError));
  }
}

function vNextUxDateOnlyText_(value) {
  if (!value) return '';
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyy-MM-dd');
  }
  var text = String(value || '').trim();
  var match = text.match(/^\d{4}-\d{2}-\d{2}/);
  return match ? match[0] : text.slice(0, 10);
}

function vNextUxGetBookContext_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var raw = typeof vNextGetBookContext_ === 'function' ? vNextGetBookContext_(ss) : vNextUxReadMetaContext_(ss);
  if (!raw) return null;
  var rawMode = String(raw.mode || '').toUpperCase();
  var normalizedMode = rawMode === 'CLIENT_BOOK' ? 'CLIENT' : rawMode === 'ADMIN_HUB' ? 'ADMIN' : rawMode;
  var state = String(raw.state || 'INPUT_OPEN').toUpperCase();
  var email = String(raw.userEmail || Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  var role = String(raw.role || raw.defaultRole || raw.default_role || 'EMPLOYEE').trim().toUpperCase();
  var owners = raw.forecastOwnerEmails || raw.forecast_owner_emails || [];
  if (!Array.isArray(owners)) owners = String(owners || '').split(',');
  owners = owners.map(function(value) { return String(value || '').trim().toLowerCase(); }).filter(Boolean);
  var input = raw.inputStatus || {
    submitted: vNextUxBool_(raw.inputSubmitted || raw.input_submitted),
    answeredCount: Number(raw.inputAnsweredCount || raw.input_answered_count || 0),
    totalCount: Number(raw.inputTotalCount || raw.input_total_count || 0),
    dueDate: raw.inputDueDate || raw.input_due_date || ''
  };
  var latestOwnEvidence = raw.latestOwnEvidence || raw.latest_own_evidence || null;
  if (latestOwnEvidence) {
    latestOwnEvidence = Object.assign({}, latestOwnEvidence, {
      evidenceId: latestOwnEvidence.evidenceId || latestOwnEvidence.evidence_id || '',
      responseType: String(latestOwnEvidence.responseType || latestOwnEvidence.response_type || '').toLowerCase()
    });
  }
  var isForecastOwner = raw.isForecastOwner === true || owners.indexOf(email) >= 0;
  var isTeamMember = vNextUxBool_(raw.isTeamMember) || vNextUxBool_(raw.is_team_member) || isForecastOwner;
  var isInternalUser = vNextUxBool_(raw.isInternalUser) || vNextUxBool_(raw.is_internal_user) || role === 'INTERNAL_CONTRIBUTOR';
  var canContribute = isTeamMember || vNextUxBool_(raw.canContribute) || vNextUxBool_(raw.can_contribute) || (isInternalUser && role === 'INTERNAL_CONTRIBUTOR');
  return {
    mode: normalizedMode,
    bookId: raw.bookId || raw.book_id || ss.getId(),
    clientId: raw.clientId || raw.client_id || '',
    clientName: raw.clientName || raw.client_name || '',
    fiscalYear: raw.fiscalYear || raw.fiscal_year || '',
    asOf: raw.asOf || raw.as_of || '',
    cutoff: raw.cutoff || '',
    state: state,
    stateReason: raw.stateReason || raw.state_reason || '',
    stateChangedAt: raw.stateChangedAt || raw.state_changed_at || '',
    relatedRunId: raw.relatedRunId || raw.related_run_id || '',
    role: role,
    userEmail: email,
    isForecastOwner: isForecastOwner,
    isTeamMember: isTeamMember,
    isInternalUser: isInternalUser,
    canContribute: canContribute,
    forecastOwnerEmails: owners,
    inputStatus: input,
    canProceed: raw.canProceed === true || raw.can_proceed === true,
    latestOwnEvidence: latestOwnEvidence,
    annualSalesBaseline: Number(raw.annualSalesBaseline || raw.annual_sales_baseline || 0),
    annualSalesBaselineBasis: raw.annualSalesBaselineBasis || raw.annual_sales_baseline_basis || '',
    version: raw.version || ''
  };
}

function vNextUxReadMetaContext_(ss) {
  var sheet = ss.getSheetByName(VNEXT_UX_CONFIG_.META_SHEET);
  if (!sheet || sheet.getLastRow() < 1) return null;
  var values = sheet.getRange(1, 1, sheet.getLastRow(), Math.min(2, sheet.getLastColumn())).getValues();
  var raw = {};
  values.forEach(function(row) { var key = String(row[0] || '').trim(); if (key) raw[key] = row[1]; });
  return raw;
}

function vNextUxGetLatestForecast_(context) {
  if (typeof vNextGetLatestForecast_ !== 'function') return null;
  var raw = vNextGetLatestForecast_(context.bookId, SpreadsheetApp.getActiveSpreadsheet()) || null;
  if (!raw) return null;
  var plan = vNextUxGetLatestPlan_(context.bookId);
  if (!plan) return raw;
  var merged = Object.assign({}, raw);
  merged.layers = Object.assign({}, raw.layers || {}, {
    adoptionDelta: Number(plan.adoption_delta || 0),
    adoptedForecast: Number(plan.adopted_forecast || 0),
    uplift: Number(plan.sales_uplift || 0),
    finalBudget: Number(plan.final_budget || 0)
  });
  merged.planStatus = String(plan.status || '');
  var planPeriods = vNextUxBuildFinalPlanPeriods_(raw, plan);
  merged.planMonths = planPeriods.months;
  merged.planQuarters = planPeriods.quarters;
  return merged;
}

function vNextUxBuildFinalPlanPeriods_(forecast, plan) {
  var systemMonths = (forecast && forecast.months || []).slice(0, 12).map(function(item) {
    return Number(item && (item.p50 !== undefined ? item.p50 : (item.value !== undefined ? item.value : item.center)) || 0);
  });
  while (systemMonths.length < 12) systemMonths.push(0);
  var system = Number(plan && plan.system_recommended || (forecast && forecast.layers && forecast.layers.systemRecommended) || 0);
  var adopted = Number(plan && plan.adopted_forecast || system);
  var upliftRows = vNextUxParseJsonArray_(plan && plan.uplift_allocation_json);
  var uplift = vNextUxMonthlyAllocationValues_(upliftRows);
  var adoptedMonths;
  if (system > 0) {
    adoptedMonths = systemMonths.map(function(value) { return Math.max(0, value * adopted / system); });
  } else {
    adoptedMonths = new Array(12).fill(Math.max(0, adopted) / 12);
  }
  var rawFinalMonths = adoptedMonths.map(function(value, index) { return value + Number(uplift[index] || 0); });
  var expected = vNextUxWholeYen_(plan && plan.final_budget || adopted + Number(plan && plan.sales_uplift || 0));
  var used = 0;
  var finalMonths = rawFinalMonths.map(function(value, index) {
    var whole = index === 11 ? expected - used : vNextUxWholeYen_(value);
    used += whole;
    return whole;
  });
  var labels = ['04','05','06','07','08','09','10','11','12','01','02','03'];
  var fiscalYear = Number(forecast && forecast.fiscalYear || forecast && forecast.fiscal_year || 0);
  var months = finalMonths.map(function(value, index) {
    var year = index < 9 ? fiscalYear : fiscalYear + 1;
    return { month: (fiscalYear ? year + '-' + labels[index] : labels[index]), p50: value, value: value };
  });
  var quarters = [0, 1, 2, 3].map(function(q) {
    var value = finalMonths.slice(q * 3, q * 3 + 3).reduce(function(sum, item) { return sum + item; }, 0);
    return { quarter: 'Q' + (q + 1), p50: value, value: value };
  });
  return { months: months, quarters: quarters };
}

function vNextUxGetLatestPlan_(bookId) {
  try {
    if (typeof vNextGetLatestPlanVersion_ === 'function') return vNextGetLatestPlanVersion_(bookId);
    if (typeof vNextReadRecords_ !== 'function') return null;
    var rows = vNextReadRecords_('PLAN_VERSION', { spreadsheet: SpreadsheetApp.getActiveSpreadsheet() }).filter(function(row) {
      return String(row.book_id || '') === String(bookId || '');
    });
    return rows.length ? rows[rows.length - 1] : null;
  } catch (err) {
    Logger.log('vNextUxGetLatestPlan_ skipped: ' + vNextUxErrorText_(err));
    return null;
  }
}

function vNextUxTransition_(request) {
  if (typeof vNextTransitionState_ !== 'function') throw new Error('状態更新機能が未設定です。管理ハブ担当者へ連絡してください。');
  request.spreadsheet = request.spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  return vNextTransitionState_(request);
}

function vNextUxMaybeAdvanceInputState_() {
  var context = vNextUxGetBookContext_();
  var status = context.inputStatus || {};
  var ready = status.readyToRun === true || (Number(status.totalCount) > 0 && Number(status.answeredCount) >= Number(status.totalCount));
  if (context.state === 'INPUT_OPEN' && ready) {
    vNextUxTransition_({ bookId: context.bookId, fromState: 'INPUT_OPEN', toState: 'READY_TO_RUN', reason: 'input_readiness_met', actorEmail: context.userEmail, actorRole: 'SYSTEM' });
  }
}

function vNextUxGetPrimaryAction_(context) {
  var owner = Boolean(context.isForecastOwner);
  var teamMember = Boolean(context.isTeamMember);
  var canContribute = Boolean(context.canContribute) || teamMember || owner;
  var submitted = Boolean(context.inputStatus && context.inputStatus.submitted);
  var definitions = {
    INPUT_OPEN: canContribute
      ? { key: 'INPUT', label: '自分の見立てを回答する', instruction: '過去実績だけでは表せない変化について回答してください。' }
      : { key: 'WAIT', label: '閲覧のみです', instruction: 'このクライアント年度ブックの情報提供メンバーには登録されていません。現在、必要な操作はありません。' },
    READY_TO_RUN: owner
      ? { key: 'REQUEST_FORECAST', label: '予測を依頼する', instruction: '入力状況を確認し、予測の作成を依頼してください。' }
      : canContribute && !submitted
        ? { key: 'INPUT', label: '自分の見立てを回答する', instruction: '予測依頼前であれば、あなたの情報を追加できます。' }
        : { key: 'WAIT', label: canContribute ? '予算策定担当の操作待ち' : '閲覧のみです', instruction: canContribute ? 'あなたの入力は保存されています。予算策定担当が予測を依頼します。' : '予算策定担当が予測を依頼します。現在、必要な操作はありません。' },
    RUNNING: { key: 'WAIT', label: '予測の完成をお待ちください', instruction: '通常は5～10分で完了します。15分以上この表示が変わらない場合は右の案内の「最新状態に更新」を使い、それでも変わらなければ管理担当者へ連絡してください。' },
    DRAFT_READY: { key: owner ? 'EDIT_PLAN' : 'VIEW_PLAN', label: owner ? '予測を確認して予算案を作る' : '予測を見る', instruction: owner ? 'システム推奨予測を確認し、採用判断と営業上積みを分けて入力してください。' : '予測の結論と根拠を確認してください。' },
    SUBMITTED: { key: 'WAIT', label: '管理ハブの確認をお待ちください', instruction: '予算案は提出済みです。差戻しまたは正式予算の通知まで、操作は不要です。' },
    CHANGES_REQUESTED: { key: owner ? 'EDIT_PLAN' : 'WAIT', label: owner ? '差戻し内容を確認して再提出する' : '予算策定担当の対応待ち', instruction: owner ? '差戻し理由を確認し、予算案を修正して再提出してください。' : '予算策定担当が差戻しに対応します。' },
    OFFICIAL_LOCKED: { key: 'VIEW_PLAN', label: '正式予算を見る', instruction: '承認済みの正式予算を確認できます。' },
    REVIEW_DUE: teamMember
      ? { key: 'REVIEW', label: '振り返りを回答する', instruction: '予実差と前提差を確認し、次年度に役立つ学びを残してください。' }
      : { key: 'VIEW_REVIEW', label: '振り返りを見る', instruction: '確定実績と予測の差を閲覧できます。' },
    YEAR_CLOSED: { key: 'VIEW_REVIEW', label: '年度の振り返りを見る', instruction: '確定した振り返りを閲覧できます。' }
  };
  var action = definitions[context.state] || { key: 'WAIT', label: '管理ハブ担当者へ確認してください', instruction: '現在の状態を確認できません。' };
  var issue = vNextUxStateIssue_(context);
  if (issue) {
    action = {
      key: 'WAIT',
      label: issue.label,
      instruction: issue.instruction,
      issueKey: issue.key,
      issueLabel: issue.label,
      issueInstruction: issue.instruction
    };
  }
  return action;
}

/** Convert trusted internal state reasons to a small employee-safe vocabulary. */
function vNextUxStateIssue_(context) {
  var reason = String(context && context.stateReason || '').toLowerCase();
  if (!reason || String(context && context.state || '').toUpperCase() !== 'READY_TO_RUN') return null;
  if (/insufficient_confirmed_history|at least 5 fiscal years|confirmed actual history|history.*required|found \d+ fiscal|過去実績.*不足|実績.*5年度/.test(reason)) {
    return {
      key: 'HISTORY_SHORTAGE', label: '実績期間が不足しています',
      instruction: '予測に必要な過去実績が5年度分そろっていません。再依頼せず、管理担当者に実績の対象期間を確認してください。'
    };
  }
  if (/zac_schema|source_schema|source schema|schema.*mismatch|required headers|column|header|列構成|列名/.test(reason)) {
    return {
      key: 'DATA_SCHEMA', label: '実績データの形式を確認中です',
      instruction: '元データの列構成を確認する必要があります。再依頼せず、管理担当者の確認をお待ちください。'
    };
  }
  if (/forecast run failed|admin_forecast_job_failed|lease expired|timeout|timed out|temporar|service unavailable|quota|一時的|処理障害/.test(reason)) {
    return {
      key: 'TEMPORARY_FAILURE', label: '前回の処理を完了できませんでした',
      instruction: '一時的な処理障害の可能性があります。繰り返し依頼せず、管理担当者に現在の状態を確認してください。'
    };
  }
  if (/forecast_failed\//.test(reason)) {
    return {
      key: 'FORECAST_FAILURE', label: '前回の予測作成を完了できませんでした',
      instruction: '同じ依頼を繰り返さず、管理担当者にこの画面の状態を伝えてください。原因確認後に管理側から再開します。'
    };
  }
  return null;
}

function vNextUxActionMenuGuide_(action) {
  if (!action) return '右側の案内に従ってください。案内が出ないときだけ、上部メニュー「年度予算策定」→「案内を開く」を使います。';
  if (action.key === 'INPUT') return '右側の案内のボタンから入力を続けてください。';
  if (['REQUEST_FORECAST', 'EDIT_PLAN', 'REVIEW', 'VIEW_PLAN', 'VIEW_REVIEW'].indexOf(action.key) >= 0) {
    return '右側の案内のボタンから、次の確認画面を開いてください。';
  }
  return '現在、操作は不要です。最新状態を確認する場合は右側の案内を見てください。';
}

function vNextUxAssertClientBook_(context) {
  if (!context || String(context.mode).toUpperCase() !== 'CLIENT') throw new Error('このスプレッドシートはvNext Client FY Bookではありません。');
}

function vNextUxAssertInputAllowed_(context) {
  vNextUxAssertClientBook_(context);
  if (!(context.canContribute || context.isTeamMember || context.isForecastOwner)) {
    throw new Error('登録メンバー、または社内アカウントだけが回答できます。');
  }
  if (VNEXT_UX_CONFIG_.INPUT_STATES.indexOf(context.state) < 0) throw new Error(vNextUxInputLockMessage_(context.state));
}

function vNextUxRoleLabel_(context) {
  if (context && context.isForecastOwner) return '予算策定担当';
  if (context && context.isTeamMember) return '情報提供メンバー';
  if (context && (String(context.role || '').toUpperCase() === 'INTERNAL_CONTRIBUTOR' || (context.isInternalUser && context.canContribute))) {
    return '社内情報提供メンバー';
  }
  return '閲覧メンバー';
}

function vNextUxAssertPlanEditable_(context) {
  vNextUxAssertClientBook_(context);
  if (!context.isForecastOwner) throw new Error('予算案の作成・提出は予算策定担当が行います。');
  if (['DRAFT_READY', 'CHANGES_REQUESTED'].indexOf(context.state) < 0) throw new Error('現在は予算案を編集できる状態ではありません。');
}

function vNextUxVerifiedOwnerRole_(context) {
  if (!context || !context.isForecastOwner) throw new Error('予算策定担当権限を確認できません。');
  return String(context.role || '').toUpperCase() === 'ADMIN' ? 'ADMIN' : 'FORECAST_OWNER';
}

function vNextUxAssertReviewEditable_(context) {
  vNextUxAssertClientBook_(context);
  if (!context.isTeamMember) throw new Error('登録された情報提供メンバーだけが振り返りを保存できます。');
  if (context.state !== 'REVIEW_DUE') throw new Error('現在は振り返りを入力できる状態ではありません。');
}

function vNextUxNormalizeReview_(payload, context) {
  payload = payload || {};
  var evaluation = vNextUxGetLatestEvaluation_(context.bookId);
  if (!evaluation || !String(evaluation.evaluation_id || '').trim() || !String(evaluation.official_vintage_id || '').trim()) {
    throw new Error('今回の正式予算に紐づく評価を確認できません。管理ハブ担当者へ連絡してください。');
  }
  var confirmedCause = vNextUxSafeText_(payload.confirmedCause, 1000, false, '確認できた原因');
  var causeHypothesis = vNextUxSafeText_(payload.causeHypothesis, 1000, false, '原因仮説');
  if (!confirmedCause && !causeHypothesis) throw new Error('確認できた原因または原因仮説を入力してください。');
  var validCauseKeys = VNEXT_UX_REVIEW_CAUSES_.map(function(item) { return item.key; });
  var causeCategories = Array.isArray(payload.causeCategories) ? payload.causeCategories.map(function(value) {
    return String(value || '').toUpperCase();
  }).filter(function(value, index, all) {
    return validCauseKeys.indexOf(value) >= 0 && all.indexOf(value) === index;
  }) : [];
  causeCategories.sort(function(left, right) { return validCauseKeys.indexOf(left) - validCauseKeys.indexOf(right); });
  if (!causeCategories.length) throw new Error('差の主な原因カテゴリを1つ以上選択してください。');
  if (causeCategories.length > 3) throw new Error('差の主な原因カテゴリは3つ以内で選択してください。');
  var nextInformation = Array.isArray(payload.nextInformation) ? payload.nextInformation.map(function(value) {
    return vNextUxSafeText_(value, 300, false, '次に確認する情報');
  }).filter(Boolean) : [];
  if (nextInformation.length > 3) throw new Error('次年度に確認する情報は3件以内で入力してください。');
  if (!nextInformation.length) throw new Error('次年度に確認する情報を1件以上入力してください。');
  return {
    bookId: context.bookId,
    fiscalYear: context.fiscalYear,
    officialVintageId: String(evaluation.official_vintage_id),
    evaluationId: String(evaluation.evaluation_id),
    causeCategories: causeCategories,
    confirmedCause: confirmedCause,
    causeHypothesis: causeHypothesis,
    nextInformation: nextInformation
  };
}

function vNextUxEvaluationBreakdown_(evaluation) {
  var row = evaluation || {};
  return VNEXT_UX_REVIEW_CAUSES_.filter(function(item) { return item.field; }).map(function(item) {
    var raw = row[item.field];
    var amount = raw === '' || raw === null || raw === undefined || !isFinite(Number(raw)) ? 0 : Number(raw);
    return {
      key: item.key,
      label: item.label,
      amount: amount,
      annual: item.annual,
      explanation: item.annual
        ? (amount > 0 ? '予測が実績より大きかった方向' : (amount < 0 ? '実績が予測より大きかった方向' : '年度差への影響なし'))
        : '年度合計ではなく月次のずれとして確認'
    };
  });
}

function vNextUxGetLatestEvaluation_(bookId) {
  try {
    if (typeof vNextReadRecords_ !== 'function') return null;
    var rows = vNextReadRecords_('EVALUATION', { spreadsheet: SpreadsheetApp.getActiveSpreadsheet() }).filter(function(row) { return String(row.book_id || '') === String(bookId || ''); });
    return rows.length ? rows[rows.length - 1] : null;
  } catch (err) {
    Logger.log('vNextUxGetLatestEvaluation_ skipped: ' + vNextUxErrorText_(err));
    return null;
  }
}

function vNextUxGetLatestOwnReviewRecord_(context) {
  try {
    if (typeof vNextReadRecords_ !== 'function') return null;
    var rows = vNextReadRecords_('EVIDENCE_EVENT', { spreadsheet: SpreadsheetApp.getActiveSpreadsheet() }).filter(function(row) {
      return String(row.book_id || '') === String(context.bookId || '') &&
        String(row.actor_email || '').toLowerCase() === String(context.userEmail || '').toLowerCase() &&
        String(row.evidence_type || '').toUpperCase() === 'REVIEW_LEARNING' && String(row.status || 'ACTIVE').toUpperCase() !== 'VOID';
    });
    return rows.length ? rows[rows.length - 1] : null;
  } catch (err) {
    Logger.log('vNextUxGetLatestOwnReviewRecord_ skipped: ' + vNextUxErrorText_(err));
    return null;
  }
}

function vNextUxGetLatestOwnReview_(context) {
  var record = vNextUxGetLatestOwnReviewRecord_(context);
  if (!record || !record.evidence_text) return null;
  try { return JSON.parse(String(record.evidence_text)); } catch (err) { return null; }
}

function vNextUxNormalizePlan_(payload, context, forecast) {
  payload = payload || {};
  if (!forecast || !forecast.runId) throw new Error('提出に使用できる予測がありません。');
  var adoptionDelta = vNextUxNumber_(payload.adoptionDelta, '採用判断の差分');
  var adoptionReason = vNextUxSafeText_(payload.adoptionReason, 1000, adoptionDelta !== 0, '採用判断の理由');
  var adoptedForecast = Number(forecast.systemRecommended) + adoptionDelta;
  if (adoptedForecast < 0) throw new Error('採用予測は0円以上にしてください。');
  var salesUplift = vNextUxNumber_(payload.salesUplift, '営業上積み');
  if (salesUplift < 0) throw new Error('営業上積みは0円以上にしてください。');
  var upliftReason = vNextUxSafeText_(payload.upliftReason, 1000, salesUplift !== 0, '営業上積みの理由');
  var upliftOwner = vNextUxSafeText_(payload.upliftOwner, 200, salesUplift !== 0, '営業上積みの責任者');
  var upliftAction = vNextUxSafeText_(payload.upliftAction, 1000, salesUplift !== 0, '具体的な行動');
  var upliftDueDate = String(payload.upliftDueDate || '').trim();
  if (salesUplift !== 0) {
    if (!upliftDueDate || isNaN(new Date(upliftDueDate).getTime())) throw new Error('営業上積みの期限を入力してください。');
  } else upliftDueDate = '';
  var months = Array.isArray(payload.monthAllocation) ? payload.monthAllocation.map(function(value, index) {
    var number = vNextUxNumber_(value, (index + 1) + '番目の月次配分');
    if (number < 0) throw new Error('月次配分は0円以上にしてください。');
    return number;
  }) : [];
  if (salesUplift !== 0 && months.length !== 12) throw new Error('営業上積みを12か月へ配分してください。');
  if (salesUplift === 0) months = new Array(12).fill(0);
  var monthTotal = months.reduce(function(sum, value) { return sum + value; }, 0);
  if (Math.abs(monthTotal - salesUplift) > 1) throw new Error('月次配分の合計を営業上積みと一致させてください。差額: ' + vNextUxFormatMoney_(salesUplift - monthTotal));
  var labels = ['04','05','06','07','08','09','10','11','12','01','02','03'];
  var upliftAllocation = months.map(function(value, index) {
    var year = index < 9 ? Number(context.fiscalYear) : Number(context.fiscalYear) + 1;
    return { month: year + '-' + labels[index], amount: value };
  });
  var quarterAllocation = [0, 1, 2, 3].map(function(q) { return months.slice(q * 3, q * 3 + 3).reduce(function(sum, value) { return sum + value; }, 0); });
  var canonical = {
    bookId: context.bookId,
    runId: forecast.runId,
    systemRecommended: forecast.systemRecommended,
    adoptionDelta: adoptionDelta,
    adoptionReason: adoptionReason,
    salesUplift: salesUplift,
    upliftReason: upliftReason,
    upliftOwner: upliftOwner,
    upliftAction: upliftAction,
    upliftDueDate: upliftDueDate,
    monthAllocation: months
  };
  return { canonical: canonical, adoptionDelta: adoptionDelta, adoptionReason: adoptionReason, adoptedForecast: adoptedForecast, salesUplift: salesUplift, upliftReason: upliftReason, upliftOwner: upliftOwner, upliftAction: upliftAction, upliftDueDate: upliftDueDate, upliftAllocation: upliftAllocation, quarterAllocation: quarterAllocation, finalBudget: adoptedForecast + salesUplift };
}

function vNextUxGetLatestChangeRequestReason_(bookId) {
  try {
    if (typeof vNextReadRecords_ !== 'function') return '';
    var rows = vNextReadRecords_('STATE_EVENT', { spreadsheet: SpreadsheetApp.getActiveSpreadsheet() }).filter(function(row) {
      return String(row.book_id || '') === String(bookId || '') && String(row.to_state || '').toUpperCase() === 'CHANGES_REQUESTED';
    });
    return rows.length ? String(rows[rows.length - 1].reason || '') : '';
  } catch (err) {
    Logger.log('vNextUxGetLatestChangeRequestReason_ skipped: ' + vNextUxErrorText_(err));
    return '';
  }
}

function vNextUxForecastMonthValues_(months) {
  var output = new Array(12).fill(0);
  (months || []).slice(0, 12).forEach(function(item, index) { output[index] = Number(item && (item.p50 || item.value || item.center) || 0); });
  return output;
}

function vNextUxMonthlyAllocationValues_(allocation) {
  var values = new Array(12).fill(0);
  (allocation || []).slice(0, 12).forEach(function(item, index) { values[index] = Number(item && item.amount !== undefined ? item.amount : item || 0); });
  return values;
}

function vNextUxParseJsonArray_(value) {
  if (Array.isArray(value)) return value;
  if (!value) return [];
  try { var parsed = JSON.parse(String(value)); return Array.isArray(parsed) ? parsed : []; } catch (err) { return []; }
}

function vNextUxNumber_(value, label) {
  if (value === '' || value === null || value === undefined) return 0;
  var number = Number(value);
  if (!isFinite(number) || Math.abs(number) > VNEXT_UX_CONFIG_.MAX_AMOUNT) throw new Error(label + 'を正しい金額で入力してください。');
  return Math.trunc(number);
}

function vNextUxCanOverrideInput_(context) {
  if (!context || !context.isForecastOwner || context.state !== 'INPUT_OPEN') return false;
  var due = context.inputStatus && context.inputStatus.dueDate;
  if (!due) return false;
  var date = new Date(due);
  return !isNaN(date.getTime()) && date.getTime() < new Date().setHours(0, 0, 0, 0);
}

function vNextUxInputLockMessage_(state) {
  var messages = {
    RUNNING: '予測を作成中のため、入力は一時停止しています。',
    DRAFT_READY: '予測案の作成後は入力できません。変更が必要な場合は予算策定担当へ連絡してください。',
    SUBMITTED: '管理ハブの確認中のため入力できません。',
    CHANGES_REQUESTED: '現在は予算策定担当が差戻し内容を修正しています。追加情報が必要な場合は入力期間が再開されます。',
    OFFICIAL_LOCKED: '正式予算は凍結されているため入力できません。',
    REVIEW_DUE: '現在は振り返り期間です。',
    YEAR_CLOSED: '年度が終了しているため入力できません。'
  };
  return messages[state] || '現在は入力できません。';
}

function vNextUxNormalizeEvidence_(payload) {
  payload = payload || {};
  var type = String(payload.responseType || '').toLowerCase();
  if (['change', 'no_change', 'unknown'].indexOf(type) < 0) throw new Error('最初の質問に回答してください。');
  var result = {
    responseType: type,
    target: '', changeKind: '', period: '', direction: '', amountMode: '', amount: '', amountBand: '', evidence: '', confidence: ''
  };
  if (type === 'change') {
    result.target = vNextUxSafeText_(payload.target, VNEXT_UX_CONFIG_.MAX_TARGET, true, '変化の対象');
    result.changeKind = String(payload.changeKind || '').toLowerCase();
    if (['contract', 'other'].indexOf(result.changeKind) < 0) throw new Error('情報の種類を選んでください。');
    result.period = vNextUxSafeText_(payload.period, 40, true, '時期');
    result.direction = String(payload.direction || '').toLowerCase();
    if (['increase', 'decrease'].indexOf(result.direction) < 0) throw new Error('増える／減るを選んでください。');
    result.amountMode = String(payload.amountMode || '').toLowerCase();
    if (['exact', 'band'].indexOf(result.amountMode) < 0) throw new Error('金額または大中小を選んでください。');
    if (result.amountMode === 'exact') {
      result.amount = vNextUxWholeYen_(payload.amount);
      if (!isFinite(result.amount) || result.amount <= 0 || result.amount > VNEXT_UX_CONFIG_.MAX_AMOUNT) throw new Error('金額を正の数で入力してください。');
    } else {
      result.amountBand = String(payload.amountBand || '').toLowerCase();
      if (['small', 'medium', 'large'].indexOf(result.amountBand) < 0) throw new Error('影響の大きさを選んでください。');
    }
    result.evidence = vNextUxSafeText_(payload.evidence, VNEXT_UX_CONFIG_.MAX_TEXT, true, '根拠');
    result.confidence = String(payload.confidence || '').toLowerCase();
    if (['confirmed', 'likely', 'hypothesis'].indexOf(result.confidence) < 0) throw new Error('情報確度を選んでください。');
  } else {
    result.evidence = vNextUxSafeText_(payload.evidence, VNEXT_UX_CONFIG_.MAX_TEXT, false, '補足');
  }
  return result;
}

function vNextUxCanonicalEvidence_(value) {
  return {
    responseType: value.responseType,
    target: value.target,
    changeKind: value.changeKind,
    period: value.period,
    direction: value.direction,
    amountMode: value.amountMode,
    amount: value.amount,
    amountBand: value.amountBand,
    evidence: value.evidence,
    confidence: value.confidence
  };
}

function vNextUxBuildAmountBands_(forecast, context) {
  var publicForecast = vNextUxPublicForecast_(forecast);
  var forecastBase = Math.abs(Number(publicForecast.systemRecommended || publicForecast.center || 0));
  var historyBase = Math.abs(Number(context && context.annualSalesBaseline || 0));
  var base = forecastBase || historyBase;
  var available = isFinite(base) && base > 0;
  return [
    { key: 'small', label: '小（年間売上の0.5〜2%）', low: available ? vNextUxRoundMoney_(base * 0.005) : null, high: available ? vNextUxRoundMoney_(base * 0.02) : null, basisAmount: base, available: available },
    { key: 'medium', label: '中（年間売上の2〜5%）', low: available ? vNextUxRoundMoney_(base * 0.02) : null, high: available ? vNextUxRoundMoney_(base * 0.05) : null, basisAmount: base, available: available },
    { key: 'large', label: '大（年間売上の5〜10%）', low: available ? vNextUxRoundMoney_(base * 0.05) : null, high: available ? vNextUxRoundMoney_(base * 0.10) : null, basisAmount: base, available: available }
  ];
}

function vNextUxCalculateImpact_(evidence, bands) {
  if (evidence.responseType !== 'change') return { low: 0, high: 0 };
  var low;
  var high;
  if (evidence.amountMode === 'exact') low = high = Number(evidence.amount);
  else {
    var band = bands.filter(function(item) { return item.key === evidence.amountBand; })[0];
    if (!band || !band.available || band.low === null || band.high === null) {
      throw new Error('このクライアントの年間売上基準を確認できないため、「金額がわかる」を選んで影響額を入力してください。');
    }
    low = band.low; high = band.high;
  }
  if (evidence.direction === 'decrease') return { low: -high, high: -low };
  return { low: low, high: high };
}

function vNextUxBuildEvidenceSaveRecord_(canonical, context, impact) {
  var isChange = canonical.responseType === 'change';
  var periodRange = isChange
    ? vNextUxPeriodRange_(canonical.period, context.fiscalYear)
    : { start: '', end: '' };
  return Object.assign({}, canonical, {
    bookId: context.bookId,
    actorEmail: context.userEmail,
    evidenceType: canonical.responseType === 'change' ? (canonical.changeKind === 'contract' ? 'COMMITMENT' : 'HUMAN_CHANGE') : 'CHECK_IN',
    period: periodRange,
    amountLow: isChange ? Math.min(Math.abs(impact.low), Math.abs(impact.high)) : '',
    amountHigh: isChange ? Math.max(Math.abs(impact.low), Math.abs(impact.high)) : '',
    impactLow: impact.low,
    impactHigh: impact.high,
    supersedesEventId: context.latestOwnEvidence && context.latestOwnEvidence.evidenceId || '',
    source: 'EMPLOYEE_SIDEBAR',
    status: 'SUBMITTED'
  });
}

function vNextUxPeriodRange_(label, fiscalYear) {
  var fy = Number(fiscalYear);
  if (!isFinite(fy)) throw new Error('対象FYを確認できません。');
  var periods = {
    'FY通年': [[fy, 4], [fy + 1, 3]],
    'Q1': [[fy, 4], [fy, 6]],
    'Q2': [[fy, 7], [fy, 9]],
    'Q3': [[fy, 10], [fy, 12]],
    'Q4': [[fy + 1, 1], [fy + 1, 3]],
    '4～6月': [[fy, 4], [fy, 6]],
    '7～9月': [[fy, 7], [fy, 9]],
    '10～12月': [[fy, 10], [fy, 12]],
    '1～3月': [[fy + 1, 1], [fy + 1, 3]]
  };
  var range = periods[String(label || '')];
  if (!range) throw new Error('時期を確認できません。');
  return {
    start: range[0][0] + '-' + ('0' + range[0][1]).slice(-2) + '-01',
    end: range[1][0] + '-' + ('0' + range[1][1]).slice(-2) + '-01'
  };
}

function vNextUxEvidenceSummary_(value, impact) {
  if (value.responseType === 'no_change') return '「確認したが変化なし」として保存します。';
  if (value.responseType === 'unknown') return '「わからない・情報不足」として保存します。';
  var direction = value.direction === 'increase' ? '増加' : '減少';
  return value.period + 'の「' + value.target + '」について、' + direction + '要因として保存します。概算影響は' + vNextUxImpactText_(impact.low, impact.high) + 'です。';
}

function vNextUxImpactText_(low, high) {
  if (!low && !high) return '金額への直接加算なし';
  if (low === high) return vNextUxFormatMoney_(low);
  return vNextUxFormatMoney_(low) + ' ～ ' + vNextUxFormatMoney_(high);
}

function vNextUxPublicForecast_(raw) {
  raw = raw || {};
  var layers = raw.layers || {};
  var lenses = raw.lenses || {};
  var quantiles = raw.quantiles || raw.annual || {};
  var historyBaseline = Number(raw.historyBaseline || layers.historyBaseline || layers.history_baseline || 0);
  var objectiveForecast = Number(raw.objectiveForecast || layers.objectiveForecast || layers.objective_forecast || 0);
  var explicitObjectiveDelta = raw.objectiveDelta || layers.objectiveDelta || layers.objective_delta;
  var objectiveDelta = explicitObjectiveDelta !== undefined && explicitObjectiveDelta !== ''
    ? Number(explicitObjectiveDelta)
    : Number(layers.commitmentDelta || 0) + Number(layers.referenceDelta || 0);
  if (!objectiveForecast && historyBaseline) objectiveForecast = historyBaseline + objectiveDelta;
  var humanDelta = Number(raw.humanDelta || layers.humanDelta || layers.human_delta || 0);
  var aiDelta = Number(raw.aiDelta || layers.aiDelta || layers.ai_delta || 0);
  var evidenceSummary = raw.evidenceSummary || {};
  var evidenceCoverage = vNextUxEvidenceCoverage_(raw, evidenceSummary, quantiles);
  var derivedDrivers = [];
  if (historyBaseline) derivedDrivers.push('過去実績から見た継続売上力 ' + vNextUxFormatMoney_(historyBaseline));
  if (objectiveDelta) derivedDrivers.push('確認できる案件・客観情報 ' + vNextUxFormatMoney_(objectiveDelta));
  if (humanDelta) derivedDrivers.push('現場から共有された変化 ' + vNextUxFormatMoney_(humanDelta));
  if (aiDelta) derivedDrivers.push('AI調査で確認した外部変化 ' + vNextUxFormatMoney_(aiDelta));
  var derivedNext = [];
  if (Number(evidenceSummary.missingResponseRate || 0) > 0) derivedNext.push('未回答メンバーが持つ情報を確認する');
  if (!Number(evidenceSummary.commitment || 0)) derivedNext.push('契約更新・確定案件の有無を確認する');
  if (!Number(evidenceSummary.human || 0)) derivedNext.push('過去実績に出ない顧客・製品の変化を確認する');
  var providedDrivers = Array.isArray(raw.drivers) && raw.drivers.length
    ? raw.drivers
    : (Array.isArray(raw.topDrivers) && raw.topDrivers.length ? raw.topDrivers : derivedDrivers);
  var providedNextInformation = Array.isArray(raw.nextInformation) && raw.nextInformation.length
    ? raw.nextInformation
    : (Array.isArray(raw.next_information) && raw.next_information.length ? raw.next_information : derivedNext);
  var providedChangeReasons = Array.isArray(raw.changeReasons) && raw.changeReasons.length
    ? raw.changeReasons
    : (Array.isArray(raw.change_reasons) ? raw.change_reasons : []);
  var warnings = [];
  if (evidenceSummary.aiUnavailable === true || evidenceSummary.ai_unavailable === true) {
    warnings.push('AI調査を利用できなかったため、今回はAI差分を0として振れ幅を広げています。');
  }
  var hasPlan = Boolean(String(raw.planStatus || raw.plan_status || '').trim());
  var layerBreakdown = vNextUxLayerBreakdown_(raw, layers, evidenceSummary, historyBaseline, humanDelta, aiDelta);
  var triangulation = vNextUxTriangulation_(raw, lenses, Number(raw.systemRecommended || layers.systemRecommended || layers.system_recommended || raw.p50 || quantiles.p50 || 0));
  var wholePeriods = vNextUxWholeYenPeriods_(raw.months || [], {
    p10: Number(raw.p10 || quantiles.p10 || 0),
    p50: Number(raw.p50 || quantiles.p50 || raw.systemRecommended || layers.systemRecommended || layers.system_recommended || 0),
    p90: Number(raw.p90 || quantiles.p90 || 0)
  }, Number(raw.fiscalYear || raw.fiscal_year || 0));
  return {
    runId: raw.runId || raw.run_id || '',
    status: raw.status || '',
    center: vNextUxWholeYen_(raw.p50 || quantiles.p50 || raw.systemRecommended || layers.systemRecommended || layers.system_recommended || 0),
    p10: vNextUxWholeYen_(raw.p10 || quantiles.p10 || 0),
    p90: vNextUxWholeYen_(raw.p90 || quantiles.p90 || 0),
    historyBaseline: vNextUxWholeYen_(historyBaseline),
    objectiveDelta: vNextUxWholeYen_(objectiveDelta),
    objectiveForecast: vNextUxWholeYen_(objectiveForecast),
    humanDelta: vNextUxWholeYen_(humanDelta),
    aiDelta: vNextUxWholeYen_(aiDelta),
    systemRecommended: vNextUxWholeYen_(raw.systemRecommended || layers.systemRecommended || layers.system_recommended || raw.p50 || quantiles.p50 || 0),
    adoptionDelta: vNextUxWholeYen_(raw.adoptionDelta || layers.adoptionDelta || layers.adoption_delta || 0),
    adoptedForecast: vNextUxWholeYen_(raw.adoptedForecast || layers.adoptedForecast || layers.adopted_forecast || 0),
    uplift: vNextUxWholeYen_(raw.uplift || layers.uplift || 0),
    finalBudget: vNextUxWholeYen_(raw.finalBudget || layers.finalBudget || layers.final_budget || 0),
    hasPlan: hasPlan,
    quarters: wholePeriods.quarters,
    months: wholePeriods.months,
    planQuarters: raw.planQuarters || [],
    planMonths: raw.planMonths || [],
    aiEvidence: (evidenceSummary.topAiEvidence || evidenceSummary.top_ai_evidence || []).slice(0, 5).map(vNextUxNormalizeAiInsight_),
    layerBreakdown: layerBreakdown,
    triangulation: triangulation,
    drivers: providedDrivers.slice(0, 3),
    nextInformation: providedNextInformation.slice(0, 3),
    changeReasons: providedChangeReasons.slice(0, 3),
    evidenceCoverage: evidenceCoverage,
    warnings: warnings
  };
}

/** Independent reference methods; no automatic averaging or forecast adjustment. */
function vNextUxTriangulation_(raw, lenses, systemRecommended) {
  var source = lenses && (lenses.triangulation || lenses.triangulation_reference) || raw && raw.triangulation || {};
  var methods = Array.isArray(source.methods) ? source.methods.slice(0, 5).map(function(method) {
    return {
      key: String(method && method.key || ''),
      label: String(method && method.label || ''),
      value: vNextUxWholeYen_(method && method.value),
      assumption: String(method && method.assumption || ''),
      basis: String(method && method.basis || '')
    };
  }).filter(function(method) { return method.label && isFinite(method.value); }) : [];
  if (!methods.length) {
    methods.push({
      key: 'INTEGRATED_SIMULATION', label: '統合シミュレーション',
      value: vNextUxWholeYen_(systemRecommended),
      assumption: '季節性・案件・現場情報・外部情報と不確実性を統合します。',
      basis: '現在の予測run'
    });
  }
  var values = methods.map(function(method) { return method.value; }).sort(function(a, b) { return a - b; });
  return {
    policy: String(source.policy || 'INDEPENDENT_REFERENCES_NOT_AUTOMATICALLY_AVERAGED'),
    methods: methods,
    referenceMin: values.length ? values[0] : 0,
    referenceMedian: values.length ? values[Math.floor((values.length - 1) / 2)] : 0,
    referenceMax: values.length ? values[values.length - 1] : 0
  };
}

/**
 * Adds the Admin-sanitized external-research cache to the current forecast for
 * presentation only. It never changes aiDelta, p50, or any persisted plan.
 */
function vNextUxForecastForView_(context, rawForecast) {
  var forecast = vNextUxPublicForecast_(rawForecast);
  var triangulationProjection = vNextUxReadTriangulationProjection_(context, forecast.runId);
  if (triangulationProjection.methods.length) {
    forecast.triangulation = vNextUxTriangulation_({ triangulation: triangulationProjection }, {}, forecast.systemRecommended);
  }
  var projection = vNextUxReadPublicAiInsightProjection_(context);
  if (!projection.insights.length) return forecast;
  var projected = projection.insights.map(function(item) {
    var normalized = vNextUxNormalizeAiInsight_(item);
    normalized.publicProjection = true;
    normalized.projectionStatus = String(item.projectionStatus || item.projection_status || 'INSIGHT_ONLY').toUpperCase();
    normalized.publishedAt = String(item.publishedAt || item.published_at || '');
    if (normalized.projectionStatus === 'PENDING_FORECAST_REFRESH' && normalized.forecastUse === 'APPLY') {
      normalized.useLabel = '次回予測で反映予定';
    } else if (normalized.forecastUse === 'INSIGHT_ONLY') {
      normalized.useLabel = '担当者向け参考（予測額には未反映）';
    }
    return normalized;
  });
  var seen = {};
  forecast.aiEvidence = forecast.aiEvidence.concat(projected).filter(function(item) {
    var key = [String(item.sourceUrl || ''), String(item.summary || ''), String(item.target || '')].join('|');
    if (seen[key]) return false;
    seen[key] = true;
    return true;
  }).slice(0, 5);
  var questions = projected.map(function(item) { return item.humanQuestion; }).filter(Boolean);
  forecast.nextInformation = forecast.nextInformation.concat(questions).filter(function(value, index, values) {
    return value && values.indexOf(value) === index;
  }).slice(0, 3);
  if (forecast.layerBreakdown) forecast.layerBreakdown.aiInsightCount = forecast.aiEvidence.length;
  var hadUnavailable = forecast.warnings.some(function(message) {
    return String(message || '').indexOf('AI調査を利用できなかった') >= 0;
  });
  if (hadUnavailable) {
    forecast.warnings = forecast.warnings.filter(function(message) {
      return String(message || '').indexOf('AI調査を利用できなかった') < 0;
    });
    forecast.warnings.push('予測実行後にAI追加調査が完了しました。下の' + projected.length + '件は担当者向け参考で、現在の予測額は変更していません。');
  }
  forecast.aiInsightProjectionUpdatedAt = projection.generatedAt;
  return forecast;
}

function vNextUxReadTriangulationProjection_(context, runId) {
  var empty = { policy: '', methods: [] };
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss && ss.getSheetByName(VNEXT_UX_CONFIG_.CONFIG_SHEET || 'VN_BOOK_CONFIG');
    if (!sheet || sheet.getLastRow() < 2) return empty;
    var values = sheet.getDataRange().getValues();
    var keyColumn = values[0].indexOf('key');
    var valueColumn = values[0].indexOf('value');
    if (keyColumn < 0 || valueColumn < 0) return empty;
    var raw = '';
    values.slice(1).some(function(row) {
      if (String(row[keyColumn] || '').trim() !== 'triangulation_reference_json') return false;
      raw = String(row[valueColumn] || '');
      return true;
    });
    if (!raw) return empty;
    var parsed = JSON.parse(raw);
    if (!parsed || parsed.schemaVersion !== 'vnext-triangulation-projection-1' ||
        String(parsed.bookId || '') !== String(context && context.bookId || '') ||
        String(parsed.runId || '') !== String(runId || '') || !Array.isArray(parsed.methods)) return empty;
    return { policy: String(parsed.policy || ''), methods: parsed.methods.slice(0, 5) };
  } catch (error) {
    Logger.log('vNextUxReadTriangulationProjection_ warning: ' + vNextUxErrorText_(error));
    return empty;
  }
}

function vNextUxReadPublicAiInsightProjection_(context) {
  var empty = { schemaVersion: '', generatedAt: '', insights: [] };
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss && ss.getSheetByName(VNEXT_UX_CONFIG_.CONFIG_SHEET || 'VN_BOOK_CONFIG');
    if (!sheet || sheet.getLastRow() < 2) return empty;
    var values = sheet.getDataRange().getValues();
    var headers = values[0].map(function(value) { return String(value || '').trim().toLowerCase(); });
    var keyColumn = headers.indexOf('key');
    var valueColumn = headers.indexOf('value');
    if (keyColumn < 0 || valueColumn < 0) return empty;
    var raw = '';
    values.slice(1).some(function(row) {
      if (String(row[keyColumn] || '').trim() !== 'public_ai_insights_json') return false;
      raw = String(row[valueColumn] || '');
      return true;
    });
    if (!raw) return empty;
    var parsed = JSON.parse(raw);
    if (!parsed || parsed.schemaVersion !== 'vnext-public-ai-insights-1' ||
        String(parsed.bookId || '') !== String(context && context.bookId || '') ||
        !Array.isArray(parsed.insights)) return empty;
    return {
      schemaVersion: parsed.schemaVersion,
      generatedAt: String(parsed.generatedAt || ''),
      insights: parsed.insights.slice(0, 5).filter(function(item) {
        return item && String(item.summary || '').trim() && /^https?:\/\//i.test(String(item.sourceUrl || ''));
      })
    };
  } catch (error) {
    Logger.log('vNextUxReadPublicAiInsightProjection_ warning: ' + vNextUxErrorText_(error));
    return empty;
  }
}

function vNextUxLayerBreakdown_(raw, layers, evidenceSummary, historyBaseline, humanDelta, aiDelta) {
  raw = raw || {};
  layers = layers || {};
  evidenceSummary = evidenceSummary || {};
  var lenses = raw.lenses || {};
  var continuity = lenses.continuity || {};
  var changeReference = lenses.changeReference || lenses.change_reference || {};
  var unknownSpot = Number(evidenceSummary.unknownSpotExpectedAnnual || evidenceSummary.unknown_spot_expected_annual || 0);
  var baseTrend = Number(continuity.baseAnnualBaseline || continuity.base_annual_baseline || 0);
  if (!baseTrend && historyBaseline) baseTrend = Math.max(0, historyBaseline - unknownSpot);
  if (!unknownSpot && baseTrend && historyBaseline) unknownSpot = historyBaseline - baseTrend;
  var commitment = Number(layers.commitmentDelta || layers.commitment_delta || 0);
  var peer = Number(changeReference.peerReferenceDelta || changeReference.peer_reference_delta || 0);
  var objective = Number(changeReference.objectiveEventDelta || changeReference.objective_event_delta || 0);
  var reference = Number(layers.referenceDelta || layers.reference_delta || 0);
  if (!peer && !objective && reference) objective = reference;
  var rows = [
    { key: 'BASE_TREND', label: '過去売上の継続トレンド', description: '確定実績の水準・成長・季節性', amount: baseTrend, kind: 'BASE' },
    { key: 'UNKNOWN_SPOT', label: '未確認の単発売上', description: '過去の突発売上が再発する可能性', amount: unknownSpot, kind: 'DELTA' },
    { key: 'COMMITMENT', label: '契約・確定案件', description: '契約更新、受注済み案件、認識月ずれ', amount: commitment, kind: 'DELTA' },
    { key: 'PEER_REFERENCE', label: '参照クラス', description: '類似条件の客観的な参照情報', amount: peer, kind: 'DELTA' },
    { key: 'OBJECTIVE_EVENTS', label: '確認できる外部・案件情報', description: '日付と根拠を確認できる客観情報', amount: objective, kind: 'DELTA' },
    { key: 'HUMAN', label: '担当者の見立て', description: '過去実績に表れない現場情報', amount: humanDelta, kind: 'DELTA' },
    { key: 'AI', label: 'AI外部調査', description: '引用できる公開情報のうち予測へ採用した差分', amount: aiDelta, kind: 'DELTA' }
  ];
  var finalAmount = vNextUxWholeYen_(raw.systemRecommended || layers.systemRecommended || layers.system_recommended || 0);
  rows = rows.map(function(row) {
    return Object.assign({}, row, { amount: vNextUxWholeYen_(row.amount) });
  });
  var displayedTotal = rows.reduce(function(sum, row) { return sum + Number(row.amount || 0); }, 0);
  var residual = finalAmount - displayedTotal;
  if (residual) {
    var baseIndex = rows.findIndex(function(row) { return row.key === 'BASE_TREND'; });
    if (baseIndex >= 0) rows[baseIndex].amount += residual;
  }
  return {
    rows: rows,
    historySubtotal: vNextUxWholeYen_(historyBaseline),
    finalAmount: finalAmount,
    checkTotal: rows.reduce(function (sum, row) { return sum + Number(row.amount || 0); }, 0),
    aiInsightCount: Number(evidenceSummary.ai || 0),
    wholeYenPolicy: 'TRUNCATE_TOWARD_ZERO_RECONCILED_TO_FINAL'
  };
}

/** 表示・分析投影では円未満を捨てる。正本runの高精度値は変更しない。 */
function vNextUxWholeYen_(value) {
  var number = Number(value || 0);
  if (!isFinite(number)) return 0;
  return Math.trunc(number);
}

/** 月次を整数円にし、最後の月で年度合計へ合わせ、Qを月次から再集計する。 */
function vNextUxWholeYenPeriods_(months, annual, fiscalYear) {
  var keys = ['p10', 'p50', 'p90'];
  var source = Array.isArray(months) ? months.slice(0, 12) : [];
  while (source.length < 12) source.push({});
  var normalized = source.map(function(item, index) {
    var month = String(item && (item.month || item.period) || '');
    if (!month && fiscalYear) {
      var date = new Date(Number(fiscalYear), 3 + index, 1);
      month = date.getFullYear() + '-' + ('0' + (date.getMonth() + 1)).slice(-2);
    }
    return {
      month: month,
      marginalP10: vNextUxWholeYen_(item && item.marginalP10),
      marginalP50: vNextUxWholeYen_(item && item.marginalP50),
      marginalP90: vNextUxWholeYen_(item && item.marginalP90)
    };
  });
  keys.forEach(function(key) {
    var target = vNextUxWholeYen_(annual && annual[key]);
    var used = 0;
    normalized.forEach(function(item, index) {
      var value = index === 11 ? target - used : vNextUxWholeYen_(source[index] && source[index][key]);
      item[key] = value;
      used += value;
    });
  });
  var quarters = [0, 1, 2, 3].map(function(q) {
    var output = { quarter: 'Q' + (q + 1) };
    keys.forEach(function(key) {
      output[key] = normalized.slice(q * 3, q * 3 + 3).reduce(function(sum, item) {
        return sum + Number(item[key] || 0);
      }, 0);
    });
    return output;
  });
  return { months: normalized, quarters: quarters };
}

function vNextUxNormalizeAiInsight_(item) {
  item = item || {};
  var axis = String(item.researchAxis || item.research_axis || 'ALTERNATIVE_SIGNALS').toUpperCase();
  var axisLabels = {
    FINANCIAL_CAPACITY: '業績・投資余力', DIGITAL_EXECUTION: 'DX・業務変革',
    PRODUCT_MARKET: '製品・市場の勢い', STRATEGY_ORGANIZATION: '戦略・組織',
    REGULATORY_SUPPLY: '規制・供給・調達', ALTERNATIVE_SIGNALS: '先行シグナル'
  };
  var forecastUse = String(item.forecastUse || item.forecast_use || (Number(item.appliedAmount || 0) ? 'APPLY' : 'INSIGHT_ONLY')).toUpperCase();
  return {
    target: String(item.target || ''), direction: String(item.direction || 'NEUTRAL').toUpperCase(),
    summary: String(item.summary || ''), sourceUrl: String(item.sourceUrl || item.source_url || ''),
    citationTitle: String(item.citationTitle || item.citation_title || ''),
    sourceDate: String(item.sourceDate || item.source_date || ''), appliedAmount: Number(item.appliedAmount || item.applied_amount || 0),
    evidenceQuality: String(item.evidenceQuality || item.evidence_quality || ''), capApplied: Boolean(item.capApplied || item.cap_applied),
    researchAxis: axis, axisLabel: axisLabels[axis] || '外部環境', signalType: String(item.signalType || item.signal_type || ''),
    sourceStrength: String(item.sourceStrength || item.source_strength || ''), forecastUse: forecastUse,
    useLabel: forecastUse === 'APPLY' ? '予測へ反映' : '担当者向け参考',
    salesRelevance: String(item.salesRelevance || item.sales_relevance || ''),
    humanQuestion: String(item.humanQuestion || item.human_question || '')
  };
}

/**
 * 根拠coverageの説明。的中確率・モデルの信頼度とは呼ばない。
 * 新runのevidence readinessを優先し、旧runでは同じ公開指標から導出する。
 */
function vNextUxEvidenceCoverage_(raw, evidenceSummary, quantiles) {
  raw = raw || {};
  evidenceSummary = evidenceSummary || {};
  quantiles = quantiles || {};
  var lenses = raw.lenses || {};
  var readiness = evidenceSummary.readiness || lenses.evidenceReadiness || lenses.evidence_readiness || {};
  var level = String(readiness.level || '').toUpperCase();
  var continuity = lenses.continuity || {};
  var fiscalYears = readiness.historyFiscalYears || continuity.fiscalYears || continuity.fiscal_years || [];
  if (!Array.isArray(fiscalYears)) fiscalYears = [];
  var historyYearCount = Number(readiness.historyYearCount);
  if (!isFinite(historyYearCount)) historyYearCount = fiscalYears.length;
  var missingRate = Number(readiness.missingResponseRate);
  if (!isFinite(missingRate)) missingRate = Number(evidenceSummary.missingResponseRate || evidenceSummary.missing_response_rate);
  var gapRate = Number(readiness.informationGapRate);
  if (!isFinite(gapRate)) gapRate = Number(evidenceSummary.informationGapRate || evidenceSummary.information_gap_rate);
  var hasMissingRate = isFinite(missingRate);
  var hasGapRate = isFinite(gapRate);
  if (!level) {
    if ((historyYearCount > 0 && historyYearCount < 5) || (hasMissingRate && missingRate >= 0.5) || (hasGapRate && gapRate >= 0.5)) {
      level = 'NEEDS_ATTENTION';
    } else if (!historyYearCount && !hasMissingRate && !hasGapRate) {
      level = 'REVIEW';
    } else if ((hasMissingRate && missingRate > 0) || (hasGapRate && gapRate > 0)) {
      level = 'REVIEW';
    } else {
      level = 'READY';
    }
  }
  var labels = {
    READY: '情報がそろっています',
    REVIEW: '確認余地があります',
    NEEDS_ATTENTION: '確認余地が大きいです'
  };
  var details = [];
  if (historyYearCount > 0) details.push('確定実績 ' + historyYearCount + '年度分');
  if (hasMissingRate) details.push('メンバー回答 ' + Math.max(0, Math.min(100, Math.round((1 - missingRate) * 100))) + '%');
  if (hasGapRate && gapRate > 0) details.push('情報不足の回答 ' + Math.max(0, Math.min(100, Math.round(gapRate * 100))) + '%');
  var p10 = Number(raw.p10 || quantiles.p10);
  var p90 = Number(raw.p90 || quantiles.p90);
  if (isFinite(p10) && isFinite(p90) && p90 >= p10 && (p10 || p90)) {
    details.push('通常の振れ幅 ' + vNextUxFormatMoney_(p10) + '～' + vNextUxFormatMoney_(p90));
  }
  return {
    level: labels[level] ? level : 'REVIEW',
    label: labels[level] || labels.REVIEW,
    detail: details.length ? details.join('／') : '回答と実績期間を確認しています。'
  };
}

function vNextUxEnsureClientSheets_(context) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var names = [VNEXT_UX_CONFIG_.HOME_SHEET, VNEXT_UX_CONFIG_.PLAN_SHEET, VNEXT_UX_CONFIG_.REVIEW_SHEET];
  names.forEach(function(name) { if (!ss.getSheetByName(name)) ss.insertSheet(name); });
  ss.getSheetByName(VNEXT_UX_CONFIG_.HOME_SHEET).showSheet();
  ss.getSheetByName(VNEXT_UX_CONFIG_.PLAN_SHEET).showSheet();
  var review = ss.getSheetByName(VNEXT_UX_CONFIG_.REVIEW_SHEET);
  if (VNEXT_UX_CONFIG_.REVIEW_STATES.indexOf(context.state) >= 0) review.showSheet(); else review.hideSheet();
  ss.getSheets().forEach(function(sheet) {
    if (names.indexOf(sheet.getName()) < 0 && !sheet.isSheetHidden()) sheet.hideSheet();
  });
}

function vNextUxRenderHome_(context, sheet) {
  sheet = sheet || SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VNEXT_UX_CONFIG_.HOME_SHEET);
  vNextUxResetViewSheet_(sheet, 14, 2);
  vNextUxRemoveLegacyImages_(sheet);
  var action = vNextUxGetPrimaryAction_(context);
  var input = context.inputStatus || {};
  sheet.getRange('A1').setValue((context.clientName || 'クライアント') + '｜FY' + (context.fiscalYear || '') + ' 年度予算').setFontSize(20).setFontWeight('bold');
  sheet.getRange('A2').setValue('シートは見る専用です。作業は右の案内から行います。').setFontColor('#5f6368').setFontSize(12);
  sheet.getRange('A4:B7').setValues([
    ['いまの状態', VNEXT_UX_STATE_LABELS_[context.state] || context.state],
    ['実績基準日', vNextUxDateText_(context.cutoff) || '未設定'],
    ['あなたの役割', vNextUxRoleLabel_(context)],
    ['回答状況', vNextUxInputStatusText_(input)]
  ]);
  sheet.getRange('A4:A7').setFontWeight('bold').setBackground('#f8f9fa');
  sheet.getRange('B4').setFontWeight('bold').setFontSize(13);
  vNextUxSetSectionHeader_(sheet, 9, 1, 2, '次にすること');
  sheet.getRange('A10:B11').setValues([
    ['操作', action.label],
    ['案内', action.instruction]
  ]);
  sheet.getRange('A10:A11').setFontWeight('bold').setBackground('#f8f9fa');
  sheet.getRange('B10').setFontSize(14).setFontWeight('bold')
    .setBackground(action.key === 'WAIT' ? '#f8f9fa' : '#e8f0fe')
    .setFontColor(action.key === 'WAIT' ? '#3c4043' : '#174ea6');
  sheet.getRange('B11').setWrap(true);
  if (action.issueKey) sheet.getRange('A10:B11').setBackground('#fef7e0');
  sheet.getRange('A13').setValue('案内が出ないときだけ、上部メニュー「年度予算策定」→「案内を開く」を使います。')
    .setFontColor('#5f6368').setFontSize(9);
  sheet.setFrozenRows(2);
  sheet.setHiddenGridlines(true);
  sheet.setColumnWidth(1, 170);
  sheet.setColumnWidth(2, 560);
  sheet.setRowHeight(11, 52);
  vNextUxProtectView_(sheet, context.state === 'YEAR_CLOSED');
}

function vNextUxRemoveLegacyImages_(sheet) {
  try {
    var images = sheet.getImages ? sheet.getImages() : [];
    images.forEach(function (image) { try { image.remove(); } catch (removeError) {} });
  } catch (error) {
    Logger.log('vNextUxRemoveLegacyImages_ warning: ' + vNextUxErrorText_(error));
  }
}

function vNextUxRenderPlan_(context, rawForecast, sheet) {
  sheet = sheet || SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VNEXT_UX_CONFIG_.PLAN_SHEET);
  vNextUxResetViewSheet_(sheet, 56, 13);
  var f = vNextUxForecastForView_(context, rawForecast);
  sheet.getRange('A1').setValue('結果').setFontSize(20).setFontWeight('bold');
  sheet.getRange('A2').setValue((context.clientName || 'クライアント') + '｜FY' + (context.fiscalYear || '') + '｜' + (VNEXT_UX_STATE_LABELS_[context.state] || context.state) + '　・　詳しい内訳は右の案内から').setFontColor('#5f6368');
  if (!rawForecast) {
    vNextUxSetSectionHeader_(sheet, 4, 1, 13, 'いま');
    sheet.getRange('A5').setValue(vNextUxEmptyForecastMessage_(context)).setFontSize(13).setWrap(true);
    sheet.setRowHeight(5, 58);
    vNextUxFinishPlanFormat_(sheet, context.state === 'YEAR_CLOSED');
    return;
  }
  var official = context.state === 'OFFICIAL_LOCKED' || context.state === 'REVIEW_DUE' || context.state === 'YEAR_CLOSED';
  var headline = vNextUxHeadlineAmount_(official, f);
  vNextUxSetSectionHeader_(sheet, 4, 1, 13, official ? '決めた予算' : '結論');
  sheet.getRange('A5:B7').setValues([
    [official ? '決めた予算' : '見込み', headline],
    ['振れ幅', vNextUxFormatMoney_(f.p10) + ' ～ ' + vNextUxFormatMoney_(f.p90)],
    [official ? '見込み' : '決めた予算', official ? f.systemRecommended : (f.hasPlan ? f.finalBudget : '未決定')]
  ]);
  sheet.getRange('A5:A7').setFontWeight('bold').setBackground('#f8f9fa');
  sheet.getRange('B5').setNumberFormat('¥#,##0').setFontSize(20).setFontWeight('bold');
  if (official || f.hasPlan) sheet.getRange('B7').setNumberFormat('¥#,##0').setFontWeight('bold');
  else sheet.getRange('B7').setHorizontalAlignment('left').setFontColor('#5f6368').setFontWeight('bold');
  if (f.warnings.length) {
    sheet.getRange('A8').setValue('要確認').setFontWeight('bold').setBackground('#fef7e0');
    sheet.getRange('B8').setValue(f.warnings.join('／')).setBackground('#fef7e0').setWrap(true);
  }
  vNextUxSetSectionHeader_(sheet, 10, 1, 13, '内訳');
  var waterfall = [
    ['履歴の水準', f.historyBaseline], ['現場の情報', f.humanDelta], ['外部情報', f.aiDelta], ['見込み', f.systemRecommended]
  ];
  if (f.hasPlan) waterfall.push(['決めた予算', f.finalBudget]);
  sheet.getRange(11, 1, waterfall.length, 2).setValues(waterfall);
  sheet.getRange(11, 2, waterfall.length, 1).setNumberFormat('¥#,##0');
  var displayedQuarters = official && f.planQuarters.length ? f.planQuarters : f.quarters;
  var displayedMonths = official && f.planMonths.length ? f.planMonths : f.months;
  vNextUxSetSectionHeader_(sheet, 18, 1, 13, official ? '決めた予算の四半期' : '見込みの四半期');
  vNextUxWritePeriodRow_(sheet, 19, displayedQuarters, 4, ['Q1', 'Q2', 'Q3', 'Q4']);
  vNextUxWriteListBlock_(sheet, 'A23:F23', 24, 1, 6, '主な根拠', (f.drivers || []).slice(0, 3), '根拠はまだ登録されていません。');
  vNextUxWriteListBlock_(sheet, 'H23:M23', 24, 8, 6, '次に確認するとよいこと', (f.nextInformation || []).slice(0, 3), '追加確認の提案はありません。');
  vNextUxSetSectionHeader_(sheet, 31, 1, 13, official ? '決めた予算の12か月' : '見込みの12か月');
  vNextUxWritePeriodRow_(sheet, 32, displayedMonths, 12, ['4月','5月','6月','7月','8月','9月','10月','11月','12月','1月','2月','3月']);
  vNextUxWriteListBlock_(sheet, 'A36:M36', 37, 1, 13, '前回から変わった理由', (f.changeReasons || []).slice(0, 3), '初回の予測です。');
  vNextUxWriteAiEvidence_(sheet, f.aiEvidence);
  vNextUxFinishPlanFormat_(sheet, context.state === 'YEAR_CLOSED');
}

function vNextUxRenderReview_(context, sheet) {
  sheet = sheet || SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VNEXT_UX_CONFIG_.REVIEW_SHEET);
  vNextUxResetViewSheet_(sheet, 48, 4);
  sheet.getRange('A1').setValue('年度の振り返り').setFontSize(18).setFontWeight('bold');
  sheet.getRange('A2').setValue((context.clientName || 'クライアント') + '｜FY' + (context.fiscalYear || '')).setFontColor('#5f6368');
  if (VNEXT_UX_CONFIG_.REVIEW_STATES.indexOf(context.state) < 0) {
    sheet.getRange('A5').setValue('振り返り期間になると、この画面に確定実績との比較が表示されます。現在、操作は不要です。').setWrap(true);
  } else {
    var evaluation = vNextUxGetLatestEvaluation_(context.bookId);
    var review = vNextUxGetLatestOwnReview_(context);
    vNextUxSetSectionHeader_(sheet, 4, 1, 4, context.state === 'YEAR_CLOSED' ? '確定した予実結果' : '今回確認する予実結果');
    if (evaluation) {
      var rows = [
        ['確定実績', Number(evaluation.actual_total || 0)],
        ['見込み', Number(evaluation.system_forecast || 0)],
        ['採用予測', Number(evaluation.adopted_forecast || 0)],
        ['決めた予算', Number(evaluation.final_budget || 0)],
        ['システム予測－実績', Number(evaluation.system_signed_error || 0)]
      ];
      sheet.getRange(5, 1, rows.length, 2).setValues(rows);
      sheet.getRange(5, 2, rows.length, 1).setNumberFormat('¥#,##0');
      sheet.getRange('D5').setValue(
        Number(evaluation.range_contains_actual || 0) === 1
          ? '確定実績は、予測時点の「通常の振れ幅」の中でした。'
          : '確定実績は、予測時点の「通常の振れ幅」の外でした。前提差を確認します。'
      ).setWrap(true).setVerticalAlignment('top').setBackground('#f8f9fa');
      sheet.setRowHeight(5, 52);
    } else {
      sheet.getRange('A5').setValue('管理側で確定実績の評価を準備しています。評価が表示されるまで入力は不要です。翌営業日も変わらない場合は管理担当者へ連絡してください。').setWrap(true).setBackground('#fef7e0');
    }
    vNextUxSetSectionHeader_(sheet, 12, 1, 4, '年度差の要因分解（確認の出発点）');
    if (evaluation) {
      var breakdown = vNextUxEvaluationBreakdown_(evaluation).filter(function(item) { return item.annual; });
      sheet.getRange('A13:C13').setValues([['要因', '金額', '説明']]).setFontWeight('bold').setBackground('#f8f9fa');
      breakdown.forEach(function(item, index) {
        var row = 14 + index;
        sheet.getRange(row, 1, 1, 3).setValues([[item.label, item.amount, item.explanation]]);
        sheet.getRange(row, 2).setNumberFormat('¥#,##0');
        sheet.getRange(row, 3).setFontColor('#5f6368').setWrap(true);
      });
      sheet.getRange('A21').setValue('計算上の分解であり、原因を自動断定するものではありません。事実と仮説を分けて確認します。').setFontColor('#5f6368');
    } else {
      sheet.getRange('A13').setValue('要因分解は評価レコード確定後に表示されます。');
    }
    vNextUxSetSectionHeader_(sheet, 23, 1, 4, '月次で別に確認すること');
    sheet.getRange('A24').setValue('季節配分と売上認識月のずれは、年度合計へ二重加算せず月次診断として扱います。').setWrap(false);
    vNextUxSetSectionHeader_(sheet, 27, 1, 4, '振り返りの目的');
    sheet.getRange('A28').setValue('人の的中率ではなく、どの情報・仮説・入力経路が役立ったかを学びます。個人の順位は作りません。').setWrap(false);
    vNextUxSetSectionHeader_(sheet, 31, 1, 4, 'あなたが保存した学び');
    if (review) {
      var learningLines = [];
      var categoryLabels = (review.causeCategories || []).map(function(key) {
        var found = VNEXT_UX_REVIEW_CAUSES_.filter(function(item) { return item.key === key; })[0];
        return found ? found.label : '';
      }).filter(Boolean);
      if (categoryLabels.length) learningLines.push('原因カテゴリ：' + categoryLabels.join('／'));
      if (review.confirmedCause) learningLines.push('確認できた原因：' + review.confirmedCause);
      if (review.causeHypothesis) learningLines.push('原因仮説：' + review.causeHypothesis);
      (review.nextInformation || []).slice(0, 3).forEach(function(item, index) {
        learningLines.push('次に確認する情報' + (index + 1) + '：' + item);
      });
      sheet.getRange('A32').setValue(learningLines.join('\n')).setWrap(true).setVerticalAlignment('top');
      sheet.setRowHeight(32, Math.min(180, 36 + learningLines.length * 24));
    } else {
      sheet.getRange('A32').setValue(context.state === 'REVIEW_DUE' && evaluation ? '右の案内から振り返りを保存してください。' : 'あなたが保存した振り返りはありません。').setWrap(true);
    }
  }
  sheet.setHiddenGridlines(true);
  sheet.setColumnWidth(1, 190);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 360);
  sheet.setColumnWidth(4, 360);
  vNextUxProtectView_(sheet, context.state === 'YEAR_CLOSED');
}

function vNextUxWritePeriodRow_(sheet, row, periods, count, fallbackLabels) {
  var labels = [];
  var values = [];
  for (var i = 0; i < count; i++) {
    var item = Array.isArray(periods) ? periods[i] : periods && (periods[fallbackLabels[i]] || periods[i]);
    labels.push(item && (item.label || item.period) || fallbackLabels[i]);
    values.push(Number(item && (item.p50 || item.value || item.center) || (typeof item === 'number' ? item : 0)));
  }
  sheet.getRange(row, 1, 1, count).setValues([labels]).setFontWeight('bold').setBackground('#f8f9fa').setHorizontalAlignment('center');
  sheet.getRange(row + 1, 1, 1, count).setValues([values]).setNumberFormat('¥#,##0').setHorizontalAlignment('right');
}

function vNextUxHeadlineAmount_(official, forecast) {
  var value = official ? forecast.finalBudget : forecast.systemRecommended;
  return Number(value || 0);
}

function vNextUxWriteListBlock_(sheet, headerRange, startRow, startCol, width, title, items, emptyText) {
  var header = sheet.getRange(headerRange);
  header.setBackground('#f1f3f4');
  sheet.getRange(headerRange.split(':')[0]).setValue(title).setFontWeight('bold');
  var lines = (items || []).map(function(item, index) { return (index + 1) + '. ' + vNextUxItemText_(item); });
  var values = new Array(5).fill('').map(function(_, index) { return [lines[index] || (index === 0 && !lines.length ? emptyText : '')]; });
  sheet.getRange(startRow, startCol, 5, 1).setValues(values).setWrap(false).setVerticalAlignment('middle');
}

function vNextUxWriteAiEvidence_(sheet, items) {
  var evidence = (items || []).slice(0, 3);
  vNextUxSetSectionHeader_(sheet, 47, 1, 13, '外部情報（詳細は右の案内）');
  if (!evidence.length) {
    sheet.getRange('A48').setValue('現在表示できるAI外部情報はありません。');
    return;
  }
  sheet.getRange('A48:G48').setValues([['観点', '対象', '外部情報', '扱い', '出典日', '担当者に確認', '出典']]).setFontWeight('bold').setBackground('#f8f9fa');
  evidence.forEach(function(item, index) {
    var row = 49 + index;
    sheet.getRange(row, 1, 1, 6).setValues([[
      String(item.axisLabel || '外部環境'), String(item.target || ''), String(item.summary || '外部情報'),
      String(item.useLabel || '担当者向け参考'), item.sourceDate || '', String(item.humanQuestion || '')
    ]]);
    sheet.getRange(row, 1, 1, 6).setWrap(true).setVerticalAlignment('top');
    var linkRange = sheet.getRange(row, 7);
    var url = String(item.sourceUrl || '');
    if (/^https?:\/\//i.test(url)) {
      var rich = SpreadsheetApp.newRichTextValue().setText('出典を開く').setLinkUrl(url).build();
      linkRange.setRichTextValue(rich).setFontColor('#1a73e8').setHorizontalAlignment('center');
    } else {
      linkRange.setValue('URLなし').setFontColor('#5f6368').setHorizontalAlignment('center');
    }
  });
}

function vNextUxItemText_(item) {
  if (typeof item === 'string') return item;
  if (!item) return '';
  return String(item.summary || item.title || item.text || item.reason || '');
}

function vNextUxFinishPlanFormat_(sheet, hardProtection) {
  sheet.setFrozenRows(2);
  sheet.setHiddenGridlines(true);
  for (var col = 1; col <= 13; col++) sheet.setColumnWidth(col, 94);
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 130);
  vNextUxProtectView_(sheet, Boolean(hardProtection));
}

function vNextUxResetViewSheet_(sheet, rows, cols) {
  var clearRows = Math.max(1, rows, sheet.getLastRow());
  var clearCols = Math.max(1, cols, sheet.getLastColumn());
  var clearRange = sheet.getRange(1, 1, clearRows, clearCols);
  clearRange.breakApart();
  clearRange.clearContent().clearFormat().clearNote().clearDataValidations();
  sheet.getRange(1, 1, rows, cols).setBackground('#ffffff').setFontFamily('Arial').setFontSize(10).setFontColor('#202124').setVerticalAlignment('middle');
}

function vNextUxSetSectionHeader_(sheet, row, startCol, width, title) {
  sheet.getRange(row, startCol, 1, width).setBackground('#f1f3f4');
  sheet.getRange(row, startCol).setValue(title).setFontWeight('bold');
}

function vNextUxEmptyForecastMessage_(context) {
  var issue = vNextUxStateIssue_(context);
  if (issue) return issue.instruction;
  var state = String(context && context.state || '').toUpperCase();
  if (state === 'RUNNING') return '予測を作成しています。通常は5～10分で完了します。15分以上変わらない場合は右の案内の「最新状態に更新」を使い、それでも変わらなければ管理担当者へ連絡してください。';
  if (state === 'INPUT_OPEN') return 'まだ予測前です。右の案内から、来年度の変化を回答してください。';
  if (state === 'READY_TO_RUN') return context && context.isForecastOwner
    ? '回答状況を確認後、右の案内から予測を依頼してください。'
    : '回答は受け付け済みです。予算策定担当が予測を依頼するまでお待ちください。';
  return 'この年度の予測結果はまだありません。ホームの「次にすること」を確認してください。';
}

function vNextUxProtectView_(sheet, hardProtection) {
  try {
    var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    var protection = protections.length ? protections[0] : sheet.protect();
    protection.setDescription('vNext自動生成表示（編集は入力サイドバーから行います）')
      .setWarningOnly(!hardProtection);
    if (hardProtection && protection.canDomainEdit()) protection.setDomainEdit(false);
  } catch (err) {
    Logger.log('vNextUxProtectView_ warning: ' + vNextUxErrorText_(err));
  }
}

function vNextUxIsHardProtected_(sheet) {
  if (!sheet) return false;
  try {
    return sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).some(function(protection) {
      return !protection.isWarningOnly();
    });
  } catch (err) {
    Logger.log('vNextUxIsHardProtected_ warning: ' + vNextUxErrorText_(err));
    return false;
  }
}

function vNextUxActivateSheet_(name, cell) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error(name + ' がありません。');
  sheet.showSheet();
  ss.setActiveSheet(sheet);
  sheet.getRange(cell || 'A1').activate();
}

function vNextUxInputStatusText_(input) {
  var answered = Number(input.answeredCount || 0);
  var total = Number(input.totalCount || 0);
  var own = input.submitted ? 'あなた：回答済み' : 'あなた：未回答';
  return total ? own + '／チーム：' + answered + ' / ' + total + '名' : own;
}

function vNextUxSafeText_(value, maxLength, required, label) {
  var text = String(value == null ? '' : value).trim();
  if (required && !text) throw new Error(label + 'を入力してください。');
  if (text.length > maxLength) throw new Error(label + 'は' + maxLength + '文字以内で入力してください。');
  if (/^[=+\-@]/.test(text)) text = "'" + text;
  return text;
}

function vNextUxSha256_(text) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return bytes.map(function(value) { var hex = (value < 0 ? value + 256 : value).toString(16); return hex.length === 1 ? '0' + hex : hex; }).join('');
}

function vNextUxRoundMoney_(value) {
  var unit = value >= 100000000 ? 1000000 : value >= 10000000 ? 100000 : 10000;
  return Math.max(unit, Math.round(value / unit) * unit);
}

function vNextUxFormatMoney_(value) {
  var number = vNextUxWholeYen_(value);
  var sign = number < 0 ? '-' : '';
  return sign + '¥' + Math.abs(number).toLocaleString('ja-JP');
}

function vNextUxDateText_(value) {
  if (!value) return '';
  var date = value instanceof Date ? value : new Date(value);
  if (isNaN(date.getTime())) return String(value);
  return Utilities.formatDate(date, Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyy/MM/dd');
}

function vNextUxBool_(value) {
  return value === true || value === 1 || String(value).toLowerCase() === 'true' || String(value) === '1';
}

function vNextUxFriendlyValidationError_(err) {
  var message = vNextUxErrorText_(err);
  return message.indexOf('Exception:') === 0 ? message.slice(10).trim() : message;
}

function vNextUxErrorText_(err) {
  return err && err.message ? err.message : String(err || '不明なエラー');
}

function vNextUxAlertError_(prefix, err) {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.alert('エラー', prefix + '\n' + vNextUxFriendlyValidationError_(err), ui.ButtonSet.OK);
  } catch (uiError) {
    // Spreadsheet UI functions cannot be shown when invoked from the Apps
    // Script editor or a time trigger. Preserve the original error in logs
    // without masking it with a second getUi() exception.
    Logger.log('vNext UI message unavailable outside Spreadsheet context: ' + prefix + ' ' + vNextUxFriendlyValidationError_(err));
  }
}

/** GAS editorから実行する小さな純粋関数test。 */
function testVNextUxEvidenceValidation() {
  try {
    var value = vNextUxNormalizeEvidence_({ responseType: 'change', target: '契約更新', changeKind: 'contract', period: 'Q1', direction: 'increase', amountMode: 'band', amountBand: 'medium', evidence: '顧客確認済み', confidence: 'confirmed' });
    if (value.responseType !== 'change' || value.amountBand !== 'medium' || value.changeKind !== 'contract') throw new Error('正規化に失敗');
    var hash1 = vNextUxSha256_(JSON.stringify(vNextUxCanonicalEvidence_(value)));
    var hash2 = vNextUxSha256_(JSON.stringify(vNextUxCanonicalEvidence_(value)));
    if (hash1 !== hash2) throw new Error('preview hashが再現しません');
    ['no_change', 'unknown'].forEach(function(responseType) {
      var normalized = vNextUxNormalizeEvidence_({ responseType: responseType, evidence: '' });
      var record = vNextUxBuildEvidenceSaveRecord_(
        vNextUxCanonicalEvidence_(normalized),
        { bookId: 'BOOK-TEST', userEmail: 'member@example.com', fiscalYear: 2027, latestOwnEvidence: null },
        { low: 0, high: 0 }
      );
      if (record.evidenceType !== 'CHECK_IN' || record.period.start !== '' || record.period.end !== '') {
        throw new Error(responseType + 'の保存recordが不正です。');
      }
    });
    Logger.log('PASS testVNextUxEvidenceValidation');
  } catch (err) {
    Logger.log('FAIL testVNextUxEvidenceValidation: ' + vNextUxErrorText_(err));
    throw err;
  }
}

function testVNextUxPlanValidation() {
  try {
    var context = { mode: 'CLIENT', bookId: 'BOOK-TEST', fiscalYear: 2027, state: 'DRAFT_READY', isForecastOwner: true };
    var forecast = { runId: 'RUN-TEST', systemRecommended: 1000000 };
    var plan = vNextUxNormalizePlan_({
      adoptionDelta: 100000,
      adoptionReason: '確認済み契約を反映',
      salesUplift: 120000,
      upliftReason: '追加提案',
      upliftOwner: '予算策定担当',
      upliftAction: '顧客へ月次提案',
      upliftDueDate: '2027-09-30',
      monthAllocation: new Array(12).fill(10000)
    }, context, forecast);
    if (plan.adoptedForecast !== 1100000 || plan.finalBudget !== 1220000) throw new Error('計画waterfallが一致しません。');
    var periods = vNextUxBuildFinalPlanPeriods_({
      fiscalYear: 2027,
      months: new Array(12).fill(0).map(function(_, index) { return { p50: index === 11 ? 83.37 : 83.33 }; }),
      layers: { systemRecommended: 1000 }
    }, {
      system_recommended: 1000, adopted_forecast: 1100, sales_uplift: 120,
      final_budget: 1220,
      uplift_allocation_json: JSON.stringify(new Array(12).fill(0).map(function(_, index) { return { month: 'M' + index, amount: 10 }; }))
    });
    var periodTotal = periods.months.reduce(function(sum, item) { return sum + item.p50; }, 0);
    if (Math.abs(periodTotal - 1220) > 0.000001) throw new Error('正式月次配分と最終予算が一致しません。');
    if (plan.quarterAllocation.length !== 4 || plan.quarterAllocation[0] !== 30000) throw new Error('Q配分が一致しません。');
    var unplanned = vNextUxPublicForecast_({ layers: { systemRecommended: 1000 }, annual: { p10: 800, p50: 1000, p90: 1200 } });
    if (unplanned.hasPlan || unplanned.adoptedForecast !== 0 || unplanned.finalBudget !== 0) throw new Error('未作成計画の表示判定が不正です。');
    var zeroOfficial = vNextUxPublicForecast_({
      planStatus: 'APPROVED', layers: { systemRecommended: 1000, adoptionDelta: -1000, adoptedForecast: 0, finalBudget: 0 }
    });
    if (!zeroOfficial.hasPlan || vNextUxHeadlineAmount_(true, zeroOfficial) !== 0) throw new Error('0円の正式予算を識別できません。');
    var viewerInput = vNextUxGetPrimaryAction_({ state: 'INPUT_OPEN', isForecastOwner: false, isTeamMember: false });
    var viewerReview = vNextUxGetPrimaryAction_({ state: 'REVIEW_DUE', isForecastOwner: false, isTeamMember: false });
    if (viewerInput.key !== 'WAIT' || viewerReview.key !== 'VIEW_REVIEW') throw new Error('閲覧メンバーの操作制限が不正です。');
    var internalContributor = { mode: 'CLIENT', state: 'INPUT_OPEN', role: 'INTERNAL_CONTRIBUTOR', isForecastOwner: false, isTeamMember: false, isInternalUser: true, canContribute: true };
    if (vNextUxGetPrimaryAction_(internalContributor).key !== 'INPUT') throw new Error('社内情報提供メンバーの入力導線が不正です。');
    if (vNextUxGetPrimaryAction_(Object.assign({}, internalContributor, { state: 'REVIEW_DUE' })).key !== 'VIEW_REVIEW') throw new Error('社内情報提供メンバーの振り返り導線が不正です。');
    if (vNextUxRoleLabel_(internalContributor) !== '社内情報提供メンバー') throw new Error('社内情報提供メンバーの役割表示が不正です。');
    vNextUxAssertInputAllowed_(internalContributor);
    var internalReviewRejected = false;
    try {
      vNextUxAssertReviewEditable_(Object.assign({}, internalContributor, { state: 'REVIEW_DUE' }));
    } catch (reviewAccessError) {
      internalReviewRejected = true;
    }
    if (!internalReviewRejected) throw new Error('社内情報提供メンバーに登録メンバー限定の振り返り保存が開放されています。');
    if (vNextUxGetPrimaryAction_({ state: 'READY_TO_RUN', isForecastOwner: true, isTeamMember: true, canContribute: true }).key !== 'REQUEST_FORECAST') throw new Error('予算策定担当の予測依頼導線が変更されています。');
    if (vNextUxGetPrimaryAction_({ state: 'READY_TO_RUN', isForecastOwner: false, isTeamMember: false, canContribute: true, inputStatus: { submitted: false } }).key !== 'INPUT') throw new Error('未回答の社内情報提供メンバーが予測依頼前に回答できません。');
    if (vNextUxGetPrimaryAction_({ state: 'READY_TO_RUN', isForecastOwner: false, isTeamMember: false, canContribute: true, inputStatus: { submitted: true } }).key !== 'WAIT') throw new Error('回答済みの社内情報提供メンバーの待機表示が不正です。');
    Logger.log('PASS testVNextUxPlanValidation');
  } catch (err) {
    Logger.log('FAIL testVNextUxPlanValidation: ' + vNextUxErrorText_(err));
    throw err;
  }
}
