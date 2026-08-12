/**
 * Forecast vNext 従業員UX。
 * Client FY Bookに、状態駆動の最小メニュー・3画面・入力APIを提供する。
 */

var VNEXT_UX_CONFIG_ = Object.freeze({
  MENU_NAME: '年度計画',
  HOME_SHEET: '1_ホーム',
  PLAN_SHEET: '2_予測と計画',
  REVIEW_SHEET: '3_振り返り',
  META_SHEET: 'BOOK_META',
  INPUT_HTML: 'VNext_InputSidebar',
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
  SUBMITTED: '管理者の確認待ち',
  CHANGES_REQUESTED: '修正依頼があります',
  OFFICIAL_LOCKED: '正式計画',
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

// Small embedded PNG used only as a clickable OverGridImage. It contains no external URL.
var VNEXT_UX_ACTION_BUTTON_PNG_ = 'iVBORw0KGgoAAAANSUhEUgAAAPAAAAA0CAYAAAC0LLUwAAABzklEQVR4nO3dTU7EMAyA0R6Do3A8bg4rpBGgaUe265i8SG+Zn2byrec4Lo63j/dP4B5XuxQsDCBc+AfEC8OJF4YTMAwmXhhOwDCYgGEwAcNg4oXhBAyDbR9w9rhr7eh3VN5J133vSMCFD6py7eh3VN5J133vSMCFD6py7eh3VN5J133vSMCFD6py7eh3VN5J133vSMCNDyKyd/Tcz+ZXrr3y7zGRgAUs4MEELGABDyZgAQt4MAELWMCDCVjAAh5MwAIW8GACFrCABxNw8ojsvcq5M9eK/h7d72N1Ak4ekb1XOXfmWtHfo/t9rE7AySOy9yrnzlwr+nt0v4/VCTh5RPZe5dyZa0V/j+73sbrtA46KPLiuuWfzz9aujEzArxFwkIAF3EnAQQIWcCcBBwlYwJ0EHLRjwNlnqVprBwIOErCAOwk4SMAC7iTgIAELuNP2AWePyN53zT2bX7l29XftRsDJI7L3XXPP5leuXf1duxFw8ojsfdfcs/mVa1d/124EnDwie98192x+5drV37Wb7QOGyQQMg/mPYBjq+B7dBwFeJ2AYTMAw2PE4ug8DXHf8HN0HAq77FbCIYYY/4xUxrO9pvEKGNV0OV9DQ72qXX+TyyTxx8sbzAAAAAElFTkSuQmCC';
var VNEXT_UX_ACTION_BUTTON_ALT_ = 'VNEXT_PRIMARY_ACTION_BUTTON';

/** legacy onOpenから呼ぶ安全なrouter。Admin Hubのmenu builderは別moduleへ委譲する。 */
function vNextHandleOnOpen_() {
  try {
    return vNextBuildClientMenu_();
  } catch (error) {
    Logger.log('vNextHandleOnOpen_ error: ' + vNextUxErrorText_(error));
    return false;
  }
}

/** Client FY Bookに従業員向け4項目だけを表示する。 */
function vNextBuildClientMenu_() {
  try {
    // Simple onOpen triggers may not expose user identity. Always render the recovery menu first;
    // identity and role are required only after the user invokes an authorized action.
    SpreadsheetApp.getUi().createMenu(VNEXT_UX_CONFIG_.MENU_NAME)
      .addItem('ホームに戻る', 'vNextGoHome')
      .addItem('自分の情報を入力・更新する', 'vNextOpenInputSidebar')
      .addItem('予測と計画を見る', 'vNextOpenForecastPlanOrCurrentAction')
      .addItem('使い方・困ったとき', 'vNextOpenHelpSidebar')
      .addToUi();
    // 初回openでも空白を見せない。描画失敗後もvNext menuは維持し、legacyへ戻さない。
    try {
      var context = vNextUxGetBookContext_();
      vNextUxAssertClientBook_(context);
      vNextRefreshEmployeeViews();
      vNextUxActivateSheet_(VNEXT_UX_CONFIG_.HOME_SHEET, 'A1');
    } catch (refreshError) {
      Logger.log('vNextBuildClientMenu_ initial refresh warning: ' + vNextUxErrorText_(refreshError));
    }
    return true;
  } catch (err) {
    Logger.log('vNextBuildClientMenu_ error: ' + vNextUxErrorText_(err));
    return false;
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
    throw new Error('従業員画面を準備できませんでした。管理者へ連絡してください。');
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

function vNextGoForecastPlan() {
  try {
    vNextRefreshEmployeeViews();
    vNextUxActivateSheet_(VNEXT_UX_CONFIG_.PLAN_SHEET, 'A1');
  } catch (err) {
    Logger.log('vNextGoForecastPlan error: ' + vNextUxErrorText_(err));
    vNextUxAlertError_('予測と計画を開けませんでした。', err);
  }
}

/** 固定4項目menuから、状態に応じた予測・計画・振り返り操作へ到達させる。 */
function vNextOpenForecastPlanOrCurrentAction() {
  try {
    var action = vNextGetClientViewModel().primaryAction;
    if (['REQUEST_FORECAST', 'EDIT_PLAN', 'REVIEW', 'VIEW_REVIEW'].indexOf(action.key) >= 0) {
      vNextOpenCurrentAction();
      return;
    }
    vNextGoForecastPlan();
  } catch (err) {
    Logger.log('vNextOpenForecastPlanOrCurrentAction error: ' + vNextUxErrorText_(err));
    vNextUxAlertError_('予測と計画を開けませんでした。', err);
  }
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
    if (!payload || String(payload.previewHash || '') !== expectedHash) throw new Error('保存前に「内容を確認」を押してください。確認後に内容を変えた場合は、もう一度確認が必要です。');
    if (typeof vNextAppendRecord_ !== 'function') throw new Error('振り返りの保存先が未設定です。管理者へ連絡してください。');
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
      .setTitle('計画案を作成・提出')
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
    var forecast = vNextUxPublicForecast_(rawForecast);
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
    if (!payload || String(payload.previewHash || '') !== expectedHash) throw new Error('提出前に「内容を確認」を押してください。確認後に内容を変えた場合は、もう一度確認が必要です。');
    if (typeof vNextAppendPlanVersion_ !== 'function') throw new Error('計画の保存機能が未設定です。管理者へ連絡してください。');
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
    return { ok: true, planVersionId: record.plan_version_id, message: '計画案を提出しました。管理者の確認をお待ちください。' };
  } catch (err) {
    Logger.log('vNextSubmitPlan error: ' + vNextUxErrorText_(err));
    throw new Error(vNextUxFriendlyValidationError_(err));
  }
}

/** Sidebarとsheet viewが共通利用する、機微情報を含めないview model。 */
function vNextGetClientViewModel() {
  try {
    var context = vNextUxGetBookContext_();
    vNextUxAssertClientBook_(context);
    var forecast = vNextUxGetLatestForecast_(context);
    var action = vNextUxGetPrimaryAction_(context);
    var bands = vNextUxBuildAmountBands_(forecast);
    return {
      ok: true,
      bookId: context.bookId || '',
      clientName: context.clientName || 'クライアント未設定',
      fiscalYear: context.fiscalYear || '',
      asOf: vNextUxDateText_(context.asOf),
      cutoff: vNextUxDateText_(context.cutoff),
      state: context.state,
      stateLabel: VNEXT_UX_STATE_LABELS_[context.state] || context.state,
      roleLabel: context.isForecastOwner ? 'Forecast Owner' : (context.isTeamMember ? '情報提供メンバー' : '閲覧メンバー'),
      isForecastOwner: Boolean(context.isForecastOwner),
      isTeamMember: Boolean(context.isTeamMember),
      canInput: Boolean(context.isTeamMember) && VNEXT_UX_CONFIG_.INPUT_STATES.indexOf(context.state) >= 0,
      inputLockMessage: vNextUxInputLockMessage_(context.state),
      inputStatus: context.inputStatus || {},
      canOverrideInput: vNextUxCanOverrideInput_(context),
      latestOwnEvidence: context.latestOwnEvidence || null,
      amountBands: bands,
      primaryAction: action,
      forecast: vNextUxPublicForecast_(forecast),
      version: context.version || ''
    };
  } catch (err) {
    Logger.log('vNextGetClientViewModel error: ' + vNextUxErrorText_(err));
    throw new Error('画面情報を取得できませんでした。再読み込みしても直らない場合は管理者へ連絡してください。');
  }
}

/** 入力内容を保存せず、金額影響と保存内容を確認する。 */
function vNextPreviewEvidence(payload) {
  try {
    var context = vNextUxGetBookContext_();
    vNextUxAssertInputAllowed_(context);
    var normalized = vNextUxNormalizeEvidence_(payload);
    var forecast = vNextUxGetLatestForecast_(context);
    var bands = vNextUxBuildAmountBands_(forecast);
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

/** preview済みの内容だけをappend-only evidenceとして保存する。 */
function vNextSaveEvidence(payload) {
  try {
    var context = vNextUxGetBookContext_();
    vNextUxAssertInputAllowed_(context);
    var normalized = vNextUxNormalizeEvidence_(payload);
    var canonical = vNextUxCanonicalEvidence_(normalized);
    var expectedHash = vNextUxSha256_(JSON.stringify(canonical));
    if (!payload || String(payload.previewHash || '') !== expectedHash) {
      throw new Error('保存前に「内容を確認」を押してください。確認後に内容を変えた場合は、もう一度確認が必要です。');
    }
    if (typeof vNextAppendEvidence_ !== 'function') throw new Error('保存先が未設定です。管理者へ連絡してください。');
    var bands = vNextUxBuildAmountBands_(vNextUxGetLatestForecast_(context));
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

/** 期限後、Forecast Ownerが未回答を理由付きで締め切る。 */
function vNextCloseInputAndProceed(reason) {
  try {
    var context = vNextUxGetBookContext_();
    if (!vNextUxCanOverrideInput_(context)) throw new Error('この操作は、回答期限後のForecast Ownerだけが実行できます。');
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

/** Forecast OwnerだけがREADY_TO_RUNから予測runを開始できる。 */
function vNextRequestForecast() {
  try {
    var context = vNextUxGetBookContext_();
    if (!context.isForecastOwner) throw new Error('予測の依頼はForecast Ownerが行います。');
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
function vNextRefreshEmployeeViews() {
  try {
    var context = vNextUxGetBookContext_();
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
    return { ok: true };
  } catch (err) {
    Logger.log('vNextRefreshEmployeeViews error: ' + vNextUxErrorText_(err));
    throw err;
  }
}

function vNextUxGetBookContext_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var raw = typeof vNextGetBookContext_ === 'function' ? vNextGetBookContext_(ss) : vNextUxReadMetaContext_(ss);
  if (!raw) return null;
  var rawMode = String(raw.mode || '').toUpperCase();
  var normalizedMode = rawMode === 'CLIENT_BOOK' ? 'CLIENT' : rawMode === 'ADMIN_HUB' ? 'ADMIN' : rawMode;
  var state = String(raw.state || 'INPUT_OPEN').toUpperCase();
  var email = String(raw.userEmail || Session.getActiveUser().getEmail() || '').trim().toLowerCase();
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
  return {
    mode: normalizedMode,
    bookId: raw.bookId || raw.book_id || ss.getId(),
    clientId: raw.clientId || raw.client_id || '',
    clientName: raw.clientName || raw.client_name || '',
    fiscalYear: raw.fiscalYear || raw.fiscal_year || '',
    asOf: raw.asOf || raw.as_of || '',
    cutoff: raw.cutoff || '',
    state: state,
    role: raw.role || raw.defaultRole || raw.default_role || 'EMPLOYEE',
    userEmail: email,
    isForecastOwner: raw.isForecastOwner === true || owners.indexOf(email) >= 0,
    isTeamMember: vNextUxBool_(raw.isTeamMember) || vNextUxBool_(raw.is_team_member) || raw.isForecastOwner === true || owners.indexOf(email) >= 0,
    forecastOwnerEmails: owners,
    inputStatus: input,
    canProceed: raw.canProceed === true || raw.can_proceed === true,
    latestOwnEvidence: latestOwnEvidence,
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
  var finalMonths = adoptedMonths.map(function(value, index) { return value + Number(uplift[index] || 0); });
  var expected = Number(plan && plan.final_budget || adopted + Number(plan && plan.sales_uplift || 0));
  var actual = finalMonths.reduce(function(sum, value) { return sum + value; }, 0);
  finalMonths[11] += expected - actual;
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
  if (typeof vNextTransitionState_ !== 'function') throw new Error('状態更新機能が未設定です。管理者へ連絡してください。');
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
  var definitions = {
    INPUT_OPEN: teamMember
      ? { key: 'INPUT', label: '自分の見立てを回答する', instruction: '過去実績だけでは表せない変化について回答してください。' }
      : { key: 'WAIT', label: '閲覧のみです', instruction: 'このブックの情報提供メンバーには登録されていません。現在、必要な操作はありません。' },
    READY_TO_RUN: owner ? { key: 'REQUEST_FORECAST', label: '予測を依頼する', instruction: '入力状況を確認し、予測の作成を依頼してください。' } : { key: 'WAIT', label: teamMember ? 'Forecast Ownerの操作待ち' : '閲覧のみです', instruction: teamMember ? 'あなたの入力は保存されています。Forecast Ownerが予測を依頼します。' : 'Forecast Ownerが予測を依頼します。現在、必要な操作はありません。' },
    RUNNING: { key: 'WAIT', label: '操作はありません', instruction: '予測を作成しています。完了すると表示が変わります。' },
    DRAFT_READY: { key: owner ? 'EDIT_PLAN' : 'VIEW_PLAN', label: owner ? '予測を確認して計画案を作る' : '予測を見る', instruction: owner ? 'システム推奨予測を確認し、採用判断と営業上積みを分けて入力してください。' : '予測の結論と根拠を確認してください。' },
    SUBMITTED: { key: 'WAIT', label: '操作はありません', instruction: '管理者が計画案を確認しています。' },
    CHANGES_REQUESTED: { key: owner ? 'EDIT_PLAN' : 'WAIT', label: owner ? '差戻し内容を確認して再提出する' : 'Forecast Ownerの対応待ち', instruction: owner ? '差戻し理由を確認し、計画案を修正して再提出してください。' : 'Forecast Ownerが差戻しに対応します。' },
    OFFICIAL_LOCKED: { key: 'VIEW_PLAN', label: '正式計画を見る', instruction: '承認済みの正式計画を確認できます。' },
    REVIEW_DUE: teamMember
      ? { key: 'REVIEW', label: '振り返りを回答する', instruction: '予実差と前提差を確認し、次年度に役立つ学びを残してください。' }
      : { key: 'VIEW_REVIEW', label: '振り返りを見る', instruction: '確定実績と予測の差を閲覧できます。' },
    YEAR_CLOSED: { key: 'VIEW_REVIEW', label: '年度の振り返りを見る', instruction: '確定した振り返りを閲覧できます。' }
  };
  return definitions[context.state] || { key: 'WAIT', label: '管理者へ確認してください', instruction: '現在の状態を確認できません。' };
}

function vNextUxActionMenuGuide_(action) {
  if (!action) return '上部メニュー：年度計画 → ホームに戻る';
  if (action.key === 'INPUT') return '上部メニュー：年度計画 → 自分の情報を入力・更新する';
  if (['REQUEST_FORECAST', 'EDIT_PLAN', 'REVIEW', 'VIEW_PLAN', 'VIEW_REVIEW'].indexOf(action.key) >= 0) {
    return '上部メニュー：年度計画 → 予測と計画を見る';
  }
  return '現在、操作は不要です。最新状態を確認する場合：年度計画 → ホームに戻る';
}

function vNextUxAssertClientBook_(context) {
  if (!context || String(context.mode).toUpperCase() !== 'CLIENT') throw new Error('このスプレッドシートはvNext Client FY Bookではありません。');
}

function vNextUxAssertInputAllowed_(context) {
  vNextUxAssertClientBook_(context);
  if (!context.isTeamMember) throw new Error('登録された情報提供メンバーだけが回答できます。');
  if (VNEXT_UX_CONFIG_.INPUT_STATES.indexOf(context.state) < 0) throw new Error(vNextUxInputLockMessage_(context.state));
}

function vNextUxAssertPlanEditable_(context) {
  vNextUxAssertClientBook_(context);
  if (!context.isForecastOwner) throw new Error('計画案の作成・提出はForecast Ownerが行います。');
  if (['DRAFT_READY', 'CHANGES_REQUESTED'].indexOf(context.state) < 0) throw new Error('現在は計画案を編集できる状態ではありません。');
}

function vNextUxVerifiedOwnerRole_(context) {
  if (!context || !context.isForecastOwner) throw new Error('Forecast Owner権限を確認できません。');
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
    throw new Error('今回の正式計画に紐づく評価を確認できません。管理者へ連絡してください。');
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
  return Math.round(number);
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
    DRAFT_READY: '予測案の作成後は入力できません。変更が必要な場合はForecast Ownerへ連絡してください。',
    SUBMITTED: '管理者の確認中のため入力できません。',
    CHANGES_REQUESTED: '現在はForecast Ownerが差戻し内容を修正しています。追加情報が必要な場合は入力期間が再開されます。',
    OFFICIAL_LOCKED: '正式計画は凍結されているため入力できません。',
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
      result.amount = Number(payload.amount);
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

function vNextUxBuildAmountBands_(forecast) {
  var publicForecast = vNextUxPublicForecast_(forecast);
  var base = Math.abs(Number(publicForecast.systemRecommended || publicForecast.center || 0));
  if (!base) base = 100000000;
  return [
    { key: 'small', label: '小', low: vNextUxRoundMoney_(base * 0.005), high: vNextUxRoundMoney_(base * 0.02) },
    { key: 'medium', label: '中', low: vNextUxRoundMoney_(base * 0.02), high: vNextUxRoundMoney_(base * 0.05) },
    { key: 'large', label: '大', low: vNextUxRoundMoney_(base * 0.05), high: vNextUxRoundMoney_(base * 0.10) }
  ];
}

function vNextUxCalculateImpact_(evidence, bands) {
  if (evidence.responseType !== 'change') return { low: 0, high: 0 };
  var low;
  var high;
  if (evidence.amountMode === 'exact') low = high = Number(evidence.amount);
  else {
    var band = bands.filter(function(item) { return item.key === evidence.amountBand; })[0];
    if (!band) throw new Error('影響の大きさを確認できません。');
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
  return {
    runId: raw.runId || raw.run_id || '',
    status: raw.status || '',
    center: Number(raw.p50 || quantiles.p50 || raw.systemRecommended || layers.systemRecommended || layers.system_recommended || 0),
    p10: Number(raw.p10 || quantiles.p10 || 0),
    p90: Number(raw.p90 || quantiles.p90 || 0),
    historyBaseline: historyBaseline,
    objectiveDelta: objectiveDelta,
    objectiveForecast: objectiveForecast,
    humanDelta: humanDelta,
    aiDelta: aiDelta,
    systemRecommended: Number(raw.systemRecommended || layers.systemRecommended || layers.system_recommended || raw.p50 || quantiles.p50 || 0),
    adoptionDelta: Number(raw.adoptionDelta || layers.adoptionDelta || layers.adoption_delta || 0),
    adoptedForecast: Number(raw.adoptedForecast || layers.adoptedForecast || layers.adopted_forecast || 0),
    uplift: Number(raw.uplift || layers.uplift || 0),
    finalBudget: Number(raw.finalBudget || layers.finalBudget || layers.final_budget || 0),
    hasPlan: hasPlan,
    quarters: raw.quarters || [],
    months: raw.months || [],
    planQuarters: raw.planQuarters || [],
    planMonths: raw.planMonths || [],
    aiEvidence: (evidenceSummary.topAiEvidence || evidenceSummary.top_ai_evidence || []).slice(0, 3),
    drivers: providedDrivers.slice(0, 3),
    nextInformation: providedNextInformation.slice(0, 3),
    changeReasons: providedChangeReasons.slice(0, 3),
    warnings: warnings
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
  vNextUxResetViewSheet_(sheet, 32, 8);
  var action = vNextUxGetPrimaryAction_(context);
  var input = context.inputStatus || {};
  sheet.getRange('A1:H2').merge().setValue((context.clientName || 'クライアント') + '｜FY' + (context.fiscalYear || '') + ' 年度計画').setFontSize(18).setFontWeight('bold').setFontColor('#ffffff').setBackground('#174ea6').setVerticalAlignment('middle');
  sheet.getRange('A4:B4').merge().setValue('現在の状態').setFontWeight('bold');
  sheet.getRange('C4:H4').merge().setValue(VNEXT_UX_STATE_LABELS_[context.state] || context.state).setFontWeight('bold').setFontColor('#174ea6').setBackground('#e8f0fe');
  sheet.getRange('A6:B6').merge().setValue('実績基準日').setFontWeight('bold');
  sheet.getRange('C6:D6').merge().setValue(vNextUxDateText_(context.cutoff) || '未設定');
  sheet.getRange('E6:F6').merge().setValue('あなたの役割').setFontWeight('bold');
  sheet.getRange('G6:H6').merge().setValue(context.isForecastOwner ? 'Forecast Owner' : '情報提供メンバー');
  sheet.getRange('A8:B8').merge().setValue('入力状況').setFontWeight('bold');
  sheet.getRange('C8:H8').merge().setValue(vNextUxInputStatusText_(input));
  sheet.getRange('A11:H11').merge().setValue('次にすること').setFontWeight('bold').setFontSize(13).setFontColor('#ffffff').setBackground('#188038');
  sheet.getRange('A12:H14').merge().setValue(action.label).setFontSize(16).setFontWeight('bold').setFontColor('#137333').setBackground('#e6f4ea').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange('A15:H17').merge().setValue(action.instruction + '\n' + vNextUxActionMenuGuide_(action)).setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange('A20:H20').merge().setValue('このブックで扱う数字').setFontWeight('bold').setBackground('#f1f3f4');
  sheet.getRange('A21:H24').merge().setValue('予測は「現在の情報から見込まれる売上」、採用予測は「責任者が判断した見込み」、最終予算は「採用予測＋営業上積み」です。3つを混ぜずに表示します。').setWrap(true).setVerticalAlignment('middle');
  sheet.getRange('C6').setNote('予測には、この日までに確定した実績だけを使用します。当月・未来・予定日は使用しません。');
  sheet.getRange('A12').setNote('Master Templateの主ボタンには vNextOpenCurrentAction を割り当てます。状態が変わっても、同じボタンが次の操作を開きます。');
  vNextUxEnsureHomeActionButton_(sheet, action);
  sheet.setFrozenRows(2);
  sheet.setHiddenGridlines(true);
  for (var col = 1; col <= 8; col++) sheet.setColumnWidth(col, 118);
  vNextUxProtectView_(sheet, context.state === 'YEAR_CLOSED');
}

function vNextUxEnsureHomeActionButton_(sheet, action) {
  try {
    var images = sheet.getImages ? sheet.getImages() : [];
    var matching = images.filter(function (image) {
      try { return String(image.getAltTextTitle() || '') === VNEXT_UX_ACTION_BUTTON_ALT_; }
      catch (error) { return false; }
    });
    matching.slice(1).forEach(function (image) { try { image.remove(); } catch (removeError) {} });
    if (!action || action.key === 'WAIT') {
      if (matching.length) matching[0].remove();
      return;
    }
    var button = matching.length ? matching[0] : sheet.insertImage(
      Utilities.newBlob(Utilities.base64Decode(VNEXT_UX_ACTION_BUTTON_PNG_), 'image/png', 'vnext-start.png'),
      6,
      12
    );
    button.setAnchorCell(sheet.getRange('F12'))
      .setWidth(200)
      .setHeight(43)
      .setAltTextTitle(VNEXT_UX_ACTION_BUTTON_ALT_)
      .setAltTextDescription(action.label + '。クリックすると現在の主操作を開始します。')
      .assignScript('vNextOpenCurrentAction');
  } catch (error) {
    // The four-item menu remains the accessible fallback when image assignment is unavailable.
    Logger.log('vNextUxEnsureHomeActionButton_ warning: ' + vNextUxErrorText_(error));
  }
}

function vNextUxRenderPlan_(context, rawForecast, sheet) {
  sheet = sheet || SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VNEXT_UX_CONFIG_.PLAN_SHEET);
  vNextUxResetViewSheet_(sheet, 60, 13);
  var f = vNextUxPublicForecast_(rawForecast);
  sheet.getRange('A1:M2').merge().setValue('予測と年度計画').setFontSize(18).setFontWeight('bold').setFontColor('#ffffff').setBackground('#174ea6').setVerticalAlignment('middle');
  if (!rawForecast) {
    sheet.getRange('A5:M9').merge().setValue('まだ予測は作成されていません。ホームに表示された「次にすること」を進めてください。').setFontSize(14).setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#f8f9fa');
    vNextUxFinishPlanFormat_(sheet, context.state === 'YEAR_CLOSED');
    return;
  }
  var official = context.state === 'OFFICIAL_LOCKED' || context.state === 'REVIEW_DUE' || context.state === 'YEAR_CLOSED';
  var headline = vNextUxHeadlineAmount_(official, f);
  sheet.getRange('A4:M4').merge().setValue(official ? '正式な最終予算' : '今回の中心見込み').setFontWeight('bold').setBackground(official ? '#fef7e0' : '#e8f0fe').setHorizontalAlignment('center');
  sheet.getRange('A5:M7').merge().setValue(headline).setNumberFormat('¥#,##0').setFontSize(24).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange('A8:M8').merge().setValue('通常の振れ幅　' + vNextUxFormatMoney_(f.p10) + ' ～ ' + vNextUxFormatMoney_(f.p90)).setHorizontalAlignment('center');
  if (f.warnings.length) {
    sheet.getRange('A9:M9').merge().setValue('要確認：' + f.warnings.join('／')).setBackground('#fef7e0').setFontColor('#8a4b00').setFontWeight('bold').setWrap(true);
  }
  sheet.getRange('A10:M10').merge().setValue('数字ができるまで').setFontWeight('bold').setBackground('#f1f3f4');
  var waterfall = [
    ['履歴基準値', f.historyBaseline], ['客観情報の差分', f.objectiveDelta], ['客観予測', f.objectiveForecast],
    ['現場情報の差分', f.humanDelta], ['AI調査の差分', f.aiDelta], ['システム推奨予測', f.systemRecommended]
  ];
  if (f.hasPlan) {
    waterfall = waterfall.concat([
      ['採用判断の差分', f.adoptionDelta], ['採用予測', f.adoptedForecast],
      ['営業上積み', f.uplift], ['最終予算', f.finalBudget]
    ]);
  }
  sheet.getRange(11, 1, waterfall.length, 2).setValues(waterfall);
  sheet.getRange(11, 2, waterfall.length, 1).setNumberFormat('¥#,##0');
  if (!f.hasPlan) {
    sheet.getRange('A18:B19').setValues([
      ['採用予測', '未決定'], ['最終予算', '未決定']
    ]).setBackground('#f8f9fa');
    sheet.getRange('B18:B19').setHorizontalAlignment('center').setFontColor('#5f6368').setFontWeight('bold');
    sheet.getRange('A18').setNote('採用予測と最終予算は、Forecast Ownerが計画案を作成するまで未決定です。');
  }
  var displayedQuarters = official && f.planQuarters.length ? f.planQuarters : f.quarters;
  var displayedMonths = official && f.planMonths.length ? f.planMonths : f.months;
  sheet.getRange('A22:M22').merge().setValue(official ? '最終予算の四半期配分' : '四半期の展開').setFontWeight('bold').setBackground('#f1f3f4');
  vNextUxWritePeriodRow_(sheet, 23, displayedQuarters, 4, ['Q1', 'Q2', 'Q3', 'Q4']);
  sheet.getRange('A26:M26').merge().setValue(official ? '最終予算の12か月配分' : '12か月の展開').setFontWeight('bold').setBackground('#f1f3f4');
  vNextUxWritePeriodRow_(sheet, 27, displayedMonths, 12, ['4月','5月','6月','7月','8月','9月','10月','11月','12月','1月','2月','3月']);
  vNextUxWriteListBlock_(sheet, 'A31:F31', 32, 1, 6, '主な根拠', f.drivers, '根拠はまだ登録されていません。');
  vNextUxWriteListBlock_(sheet, 'H31:M31', 32, 8, 6, '次に確認すると有効な情報', f.nextInformation, '追加確認の提案はありません。');
  vNextUxWriteListBlock_(sheet, 'A39:M39', 40, 1, 13, '前回から変わった理由', f.changeReasons, '初回の予測です。');
  vNextUxWriteAiEvidence_(sheet, f.aiEvidence);
  sheet.getRange('A5').setNote('中心見込みは、情報締切時点の条件付き予測です。必ず達成する目標を意味しません。');
  sheet.getRange('A8').setNote('P10～P90を平易に表した範囲です。悲観・楽観シナリオではありません。');
  vNextUxFinishPlanFormat_(sheet, context.state === 'YEAR_CLOSED');
}

function vNextUxRenderReview_(context, sheet) {
  sheet = sheet || SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VNEXT_UX_CONFIG_.REVIEW_SHEET);
  vNextUxResetViewSheet_(sheet, 52, 8);
  sheet.getRange('A1:H2').merge().setValue('振り返り').setFontSize(18).setFontWeight('bold').setFontColor('#ffffff').setBackground('#174ea6');
  if (VNEXT_UX_CONFIG_.REVIEW_STATES.indexOf(context.state) < 0) {
    sheet.getRange('A5:H8').merge().setValue('振り返り期間になると、この画面が表示されます。').setWrap(true).setVerticalAlignment('middle');
  } else {
    var evaluation = vNextUxGetLatestEvaluation_(context.bookId);
    var review = vNextUxGetLatestOwnReview_(context);
    sheet.getRange('A4:H4').merge().setValue(context.state === 'YEAR_CLOSED' ? '確定した予実結果' : '今回確認する予実結果').setFontWeight('bold').setBackground('#f1f3f4');
    if (evaluation) {
      var rows = [
        ['確定実績', Number(evaluation.actual_total || 0)],
        ['システム推奨予測', Number(evaluation.system_forecast || 0)],
        ['採用予測', Number(evaluation.adopted_forecast || 0)],
        ['最終予算', Number(evaluation.final_budget || 0)],
        ['システム予測－実績', Number(evaluation.system_signed_error || 0)]
      ];
      sheet.getRange(5, 1, rows.length, 2).setValues(rows);
      sheet.getRange(5, 2, rows.length, 1).setNumberFormat('¥#,##0');
      sheet.getRange('D5:H9').merge().setValue(
        Number(evaluation.range_contains_actual || 0) === 1
          ? '確定実績は、予測時点の「通常の振れ幅」の中でした。'
          : '確定実績は、予測時点の「通常の振れ幅」の外でした。前提差を確認します。'
      ).setWrap(true).setVerticalAlignment('middle').setBackground('#f8f9fa');
    } else {
      sheet.getRange('A5:H9').merge().setValue('評価レコードを確認できません。管理者へ連絡してください。').setBackground('#fce8e6');
    }
    sheet.getRange('A12:H12').merge().setValue('年度差の要因分解（確認の出発点）').setFontWeight('bold').setBackground('#f1f3f4');
    if (evaluation) {
      var breakdown = vNextUxEvaluationBreakdown_(evaluation).filter(function(item) { return item.annual; });
      breakdown.forEach(function(item, index) {
        var row = 13 + index;
        sheet.getRange(row, 1, 1, 3).merge().setValue(item.label);
        sheet.getRange(row, 4, 1, 2).merge().setValue(item.amount).setNumberFormat('¥#,##0');
        sheet.getRange(row, 6, 1, 3).merge().setValue(item.explanation).setFontColor('#5f6368');
      });
      sheet.getRange('A20:H20').merge().setValue('計算上の分解であり、原因を自動断定するものではありません。下の振り返りで事実と仮説を分けて確認します。').setWrap(true).setFontColor('#5f6368');
    } else {
      sheet.getRange('A13:H20').merge().setValue('要因分解は評価レコード確定後に表示されます。').setWrap(true);
    }
    sheet.getRange('A23:H23').merge().setValue('月次で別に確認すること').setFontWeight('bold').setBackground('#f1f3f4');
    sheet.getRange('A24:H26').merge().setValue('季節配分と売上認識月のずれは、年度合計を変えず月ごとの山谷を動かします。そのため上の年度差へ二重加算せず、月次診断として扱います。').setWrap(true).setVerticalAlignment('middle');
    sheet.getRange('A29:H29').merge().setValue('振り返りの目的').setFontWeight('bold').setBackground('#f1f3f4');
    sheet.getRange('A30:H32').merge().setValue('評価するのは人ではなく、どの情報・仮説・入力経路が役に立ったかです。個人の的中率や順位は作りません。').setWrap(true).setVerticalAlignment('middle');
    sheet.getRange('A35:H35').merge().setValue('あなたが保存した学び').setFontWeight('bold').setBackground('#f1f3f4');
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
      sheet.getRange('A36:H48').merge().setValue(learningLines.join('\n')).setWrap(true).setVerticalAlignment('top');
    } else {
      sheet.getRange('A36:H42').merge().setValue(context.state === 'REVIEW_DUE' ? '上部メニュー「年度計画」→「予測と計画を見る」から振り返りを保存してください。' : 'あなたが保存した振り返りはありません。').setWrap(true);
    }
  }
  sheet.setHiddenGridlines(true);
  for (var col = 1; col <= 8; col++) sheet.setColumnWidth(col, 118);
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
  sheet.getRange(headerRange).merge().setValue(title).setFontWeight('bold').setBackground('#f1f3f4');
  var lines = (items || []).map(function(item, index) { return (index + 1) + '. ' + vNextUxItemText_(item); });
  sheet.getRange(startRow, startCol, 5, width).merge().setValue(lines.length ? lines.join('\n') : emptyText).setWrap(true).setVerticalAlignment('top');
}

function vNextUxWriteAiEvidence_(sheet, items) {
  var evidence = (items || []).slice(0, 3);
  sheet.getRange('A47:M47').merge().setValue('AI調査で使用した主な外部情報').setFontWeight('bold').setBackground('#f1f3f4');
  if (!evidence.length) {
    sheet.getRange('A48:M50').merge().setValue('今回、予測へ直接反映したAI外部情報はありません。').setWrap(true).setVerticalAlignment('top');
    return;
  }
  evidence.forEach(function(item, index) {
    var row = 48 + index;
    var direction = String(item.direction || '').toUpperCase() === 'DOWN' ? '減少' : (String(item.direction || '').toUpperCase() === 'UP' ? '増加' : '中立');
    var detail = (index + 1) + '. ' + String(item.summary || item.target || '外部情報') +
      '｜' + direction + '｜適用額 ' + vNextUxFormatMoney_(item.appliedAmount || 0) +
      (item.sourceDate ? '｜出典日 ' + item.sourceDate : '') +
      (item.evidenceQuality ? '｜品質 ' + item.evidenceQuality : '') +
      (item.capApplied ? '｜上限適用' : '');
    sheet.getRange(row, 1, 1, 11).merge().setValue(detail).setWrap(true);
    var linkRange = sheet.getRange(row, 12, 1, 2).merge();
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
  vNextUxProtectView_(sheet, Boolean(hardProtection));
}

function vNextUxResetViewSheet_(sheet, rows, cols) {
  sheet.getRange(1, 1, Math.max(rows, sheet.getMaxRows()), Math.max(cols, sheet.getMaxColumns())).breakApart();
  sheet.clearContents().clearFormats().clearNotes();
  sheet.getRange(1, 1, rows, cols).setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
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
  if (input.submitted) return '回答済み';
  var answered = Number(input.answeredCount || 0);
  var total = Number(input.totalCount || 0);
  return total ? 'チーム ' + answered + ' / ' + total + ' 名が回答済み' : '未回答';
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
  var number = Number(value || 0);
  var sign = number < 0 ? '-' : '';
  return sign + '¥' + Math.abs(Math.round(number)).toLocaleString('ja-JP');
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
  SpreadsheetApp.getUi().alert('エラー', prefix + '\n' + vNextUxFriendlyValidationError_(err), SpreadsheetApp.getUi().ButtonSet.OK);
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
      upliftOwner: 'Forecast Owner',
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
    if (!zeroOfficial.hasPlan || vNextUxHeadlineAmount_(true, zeroOfficial) !== 0) throw new Error('0円の正式計画を識別できません。');
    var viewerInput = vNextUxGetPrimaryAction_({ state: 'INPUT_OPEN', isForecastOwner: false, isTeamMember: false });
    var viewerReview = vNextUxGetPrimaryAction_({ state: 'REVIEW_DUE', isForecastOwner: false, isTeamMember: false });
    if (viewerInput.key !== 'WAIT' || viewerReview.key !== 'VIEW_REVIEW') throw new Error('閲覧メンバーの操作制限が不正です。');
    Logger.log('PASS testVNextUxPlanValidation');
  } catch (err) {
    Logger.log('FAIL testVNextUxPlanValidation: ' + vNextUxErrorText_(err));
    throw err;
  }
}
