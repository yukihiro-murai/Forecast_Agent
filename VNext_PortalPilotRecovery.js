/**
 * Resumable Admin-only recovery for the first employee Portal pilot.
 * One invocation performs at most one durable phase so Apps Script timeouts do not restart the whole setup.
 */

var VNEXT_PORTAL_PILOT_RECOVERY_SCHEMA_ = 'vnext-portal-pilot-recovery-v1';
var VNEXT_PORTAL_PILOT_RECOVERY_PROPERTY_ = 'VNEXT_EMPLOYEE_PORTAL_PILOT_RECOVERY_V1_JSON';
var VNEXT_PORTAL_PILOT_RECOVERY_PHASES_ = Object.freeze([
  'INITIALIZE',
  'TEMPLATE_CREATE',
  'TEMPLATE_INITIALIZE',
  'TEMPLATE_COPY',
  'TEMPLATE_REGISTER',
  'MODEL_REGISTER',
  'PAIR_ACTIVATE',
  'PORTAL_CREATE',
  'PORTAL_REGISTER',
  'VERIFY',
  'COMPLETED'
]);

/**
 * Runs one recovery phase. The first call requires the same explicit Admin
 * attestation as vNextAdminPrepareEmployeePortalPilot; later calls resume from
 * the Script Property journal and do not require the request to be repeated.
 */
function vNextAdminContinueEmployeePortalPilotRecovery(request) {
  return vNextAdminGuard_('vNextAdminContinueEmployeePortalPilotRecovery', function () {
    const hub = vNextAdminRequireHub_();
    const req = request && typeof request === 'object' ? request : {};
    let state = vNextAdminPortalPilotRecoveryLoad_();
    if (!state) {
      state = vNextAdminPortalPilotRecoveryInitialize_(hub, req);
      return vNextAdminPortalPilotRecoveryPublicResult_(state, {
        initialized: true,
        message: '回復journalを作成しました。もう一度実行するとTemplate containerを作成します。'
      });
    }
    vNextAdminPortalPilotRecoveryAssertState_(hub, state);
    if (state.phase === 'COMPLETED') {
      return vNextAdminPortalPilotRecoveryVerifyCompleted_(hub, state);
    }
    try {
      const result = vNextAdminPortalPilotRecoveryRunPhase_(hub, state);
      return vNextAdminPortalPilotRecoveryPublicResult_(result.state, result.detail);
    } catch (error) {
      vNextAdminPortalPilotRecoveryRecordError_(hub, state, error);
      throw error;
    }
  });
}

/** Read-only progress API for the Admin Sidebar or operational monitoring. */
function vNextAdminGetEmployeePortalPilotRecoveryStatus() {
  return vNextAdminGuard_('vNextAdminGetEmployeePortalPilotRecoveryStatus', function () {
    const hub = vNextAdminRequireHub_();
    const state = vNextAdminPortalPilotRecoveryLoad_();
    if (!state) return { exists: false, phase: 'NOT_STARTED', completed: false };
    vNextAdminPortalPilotRecoveryAssertState_(hub, state);
    return vNextAdminPortalPilotRecoveryPublicResult_(state, { statusOnly: true });
  });
}

/**
 * Spreadsheet-menu entry. Do not run this UI wrapper from the Apps Script
 * editor; editor/trigger callers use vNextAdminContinueEmployeePortalPilotRecovery.
 */
function vNextAdminContinueEmployeePortalPilotRecoveryForManualTest() {
  const ui = SpreadsheetApp.getUi();
  const existing = vNextAdminPortalPilotRecoveryLoad_();
  let request = {};
  if (!existing) {
    const answer = ui.alert(
      '申請入口 Pilot（段階実行）',
      'Client runtime 10 suites、Portal runtime 11 suites、統合契約test PASSを確認しましたか？\n\n' +
        '1回につき1段階だけ進めます。完了まで同じメニューを繰り返し選択してください。',
      ui.ButtonSet.YES_NO
    );
    if (answer !== ui.Button.YES) return { ok: false, cancelled: true };
    const clientBundle = vNextClientRuntimeVerifiedBundle_();
    const portalBundle = vNextPortalRuntimeVerifiedBundle_();
    request = {
      attestationConfirmed: true,
      evidenceArtifact: vNextAdminCanonicalJson_({
        verifiedAt: '2026-08-12',
        clientRuntimeTests: 10,
        clientRuntimeVersion: clientBundle.version,
        clientRuntimeSha256: clientBundle.sha256,
        portalRuntimeTests: 11,
        portalRuntimeVersion: portalBundle.version,
        portalRuntimeSha256: portalBundle.sha256,
        integrationContractTests: 'PASS'
      })
    };
  }
  try {
    const result = vNextAdminContinueEmployeePortalPilotRecovery(request);
    const position = result.phaseIndex;
    const progress = result.completed
      ? '準備が完了しました。\n' + String(result.portalSpreadsheetUrl || '')
      : '段階 ' + position + ' / ' + result.phaseCount + ' まで完了しました。\n' +
        '次: ' + vNextAdminPortalPilotRecoveryPhaseLabel_(result.phase) + '\n\n' +
        '同じメニューをもう一度選択してください。';
    ui.alert('申請入口 Pilot', progress, ui.ButtonSet.OK);
    return result;
  } catch (error) {
    Logger.log('Portal Pilot phased menu failed: %s', String(error && error.stack || error));
    ui.alert('申請入口 Pilot を継続できません', String(error && error.message || error), ui.ButtonSet.OK);
    throw error;
  }
}

function vNextAdminPortalPilotRecoveryRunPhase_(hub, state) {
  switch (state.phase) {
    case 'TEMPLATE_CREATE': return vNextAdminPortalPilotRecoveryTemplateCreate_(hub, state);
    case 'TEMPLATE_INITIALIZE': return vNextAdminPortalPilotRecoveryTemplateInitialize_(hub, state);
    case 'TEMPLATE_COPY': return vNextAdminPortalPilotRecoveryTemplateCopy_(hub, state);
    case 'TEMPLATE_REGISTER': return vNextAdminPortalPilotRecoveryTemplateRegister_(hub, state);
    case 'MODEL_REGISTER': return vNextAdminPortalPilotRecoveryModelRegister_(hub, state);
    case 'PAIR_ACTIVATE': return vNextAdminPortalPilotRecoveryPairActivate_(hub, state);
    case 'PORTAL_CREATE': return vNextAdminPortalPilotRecoveryPortalCreate_(hub, state);
    case 'PORTAL_REGISTER': return vNextAdminPortalPilotRecoveryPortalRegister_(hub, state);
    case 'VERIFY': return vNextAdminPortalPilotRecoveryComplete_(hub, state);
    default: throw new Error('Unsupported Portal Pilot recovery phase: ' + String(state.phase || ''));
  }
}

function vNextAdminPortalPilotRecoveryInitialize_(hub, request) {
  const req = request && typeof request === 'object' ? request : {};
  if (req.attestationConfirmed !== true) {
    throw new Error('初回だけattestationConfirmed=trueを明示してください。');
  }
  const evidenceArtifact = vNextAdminRequiredText_(req.evidenceArtifact, 'evidenceArtifact');
  return vNextAdminWithScriptLock_('initialize-portal-pilot-recovery', function () {
    const existing = vNextAdminPortalPilotRecoveryLoad_();
    if (existing) return existing;
    if (typeof vNextClientRuntimeVerifiedBundle_ !== 'function' ||
        typeof vNextClientRuntimeCreateBoundSpreadsheet_ !== 'function' ||
        typeof vNextPortalRuntimeVerifiedBundle_ !== 'function' ||
        typeof vNextPortalRuntimeCreateBoundSpreadsheet_ !== 'function') {
      throw new Error('Client/Portal runtime publisher is not installed.');
    }
    const actor = vNextAdminActor_().toLowerCase();
    const initialPair = vNextAdminReadActiveReleasePair_(hub);
    const initialRelease = vNextAdminResolveRelease_(hub, initialPair.releaseId);
    const initialModel = vNextAdminResolveActiveModelRelease_(hub, initialPair.modelReleaseId);
    const parameters = vNextAdminParseJson_(initialModel.parameters_json, {});
    const engineVersion = typeof VNEXT_ENGINE !== 'undefined' ? String(VNEXT_ENGINE.VERSION || '') : '';
    if (!engineVersion) throw new Error('Forecast Engine version is unavailable.');
    const bundle = vNextClientRuntimeVerifiedBundle_();
    const portalBundle = vNextPortalRuntimeVerifiedBundle_();
    const source = vNextAdminPortalPilotRecoveryResolveTemplateSource_(hub, req);
    const sourceManifestSha256 = source.manifestSha256;
    const calculatedReleaseId = 'vnext-client-' + String(bundle.version || '').replace(/[^A-Za-z0-9.-]/g, '-') + '-' +
      String(bundle.sha256 || '').slice(0, 8) + '-' + sourceManifestSha256.slice(0, 8);
    const calculatedExisting = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.RELEASES).rows.find(function (row) {
      return String(row.release_id || '') === calculatedReleaseId;
    });
    const canReuseInitialRelease = !source.isDraft && !calculatedExisting &&
      String(initialRelease.client_runtime_version || '') === String(bundle.version || '') &&
      String(initialRelease.client_runtime_sha256 || '') === String(bundle.sha256 || '') &&
      String(initialRelease.template_content_sha256 || '') === sourceManifestSha256;
    const releaseId = canReuseInitialRelease ? String(initialRelease.release_id || '') : calculatedReleaseId;
    const modelReleaseId = 'model-portal-' + vNextAdminSha256_(vNextAdminCanonicalJson_({
      releaseId: releaseId, engineVersion: engineVersion, parameters: parameters
    })).slice(0, 20).toUpperCase();
    const hubConfig = vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET);
    const employeeDomain = vNextAdminNormalizeDomain_(req.employeeDomain ||
      hubConfig.employee_domain || vNextGetRuntimeConfig_().VNEXT_EMPLOYEE_DOMAIN || vNextAdminEmailDomain_(actor));
    if (!employeeDomain || vNextAdminEmailDomain_(actor) !== employeeDomain) {
      throw new Error('employeeDomainは管理ハブ担当者の Google Workspace ドメインと一致する必要があります。');
    }
    const adminEmails = vNextAdminMergeEmails_(hubConfig.admin_emails,
      vNextGetRuntimeConfig_().VNEXT_ADMIN_EMAILS, actor);
    const operationId = 'PORTAL-PILOT-' + vNextAdminSha256_(vNextAdminCanonicalJson_({
      hubSpreadsheetId: hub.getId(), initialPair: initialPair, releaseId: releaseId,
      modelReleaseId: modelReleaseId, sourceSpreadsheetId: source.spreadsheet.getId(),
      sourceManifestSha256: sourceManifestSha256, clientRuntimeSha256: bundle.sha256,
      portalRuntimeSha256: portalBundle.sha256
    })).slice(0, 24).toUpperCase();
    const now = new Date().toISOString();
    const currentTemplateFile = DriveApp.getFileById(
      vNextAdminRequiredText_(initialRelease.template_spreadsheet_id, 'activeRelease.template_spreadsheet_id')
    );
    const parents = currentTemplateFile.getParents();
    const templateFolderId = parents.hasNext() ? parents.next().getId() : DriveApp.getRootFolder().getId();
    const state = {
      schemaVersion: VNEXT_PORTAL_PILOT_RECOVERY_SCHEMA_, operationId: operationId,
      hubSpreadsheetId: hub.getId(), phase: 'TEMPLATE_CREATE', status: 'IN_PROGRESS',
      initialReleaseId: initialPair.releaseId, initialModelReleaseId: initialPair.modelReleaseId,
      initialTemplateSpreadsheetId: initialPair.templateSpreadsheetId,
      releaseId: releaseId, modelReleaseId: modelReleaseId,
      sourceSpreadsheetId: source.spreadsheet.getId(), sourceReleaseId: source.sourceReleaseId,
      sourceDraftId: source.draftId || '', sourceManifestSha256: sourceManifestSha256,
      sourceManifestSchema: source.sourceManifestSchema,
      sourceStoredManifestSha256: source.sourceStoredManifestSha256,
      legacyV2BridgeUsed: source.legacyV2BridgeUsed === true,
      clientRuntimeVersion: String(bundle.version || ''), clientRuntimeSha256: String(bundle.sha256 || ''),
      portalRuntimeVersion: String(portalBundle.version || ''), portalRuntimeSha256: String(portalBundle.sha256 || ''),
      engineVersion: engineVersion, parameters: parameters,
      evidenceArtifact: evidenceArtifact,
      reason: '申請入口・社内情報提供メンバー対応の Pilot release',
      modelNote: '申請入口 Pilot。予測 Engine と parameter は直前の ACTIVE Model Release から変更なし。',
      templateFolderId: templateFolderId, templateBookId: 'TPL-' + operationId.slice(-28),
      templateSpreadsheetId: '', templateScriptId: '', templateCreateAttempt: 0,
      templateStagingTitle: '', portalFolderId: '', portalId: 'PORTAL-' + operationId.slice(-28),
      portalSpreadsheetId: '', portalScriptId: '', portalCreateAttempt: 0,
      portalStagingTitle: '', portalFinalTitle: vNextAdminText_(req.portalTitle) || VNEXT_NAMING.LAYER2_DEFAULT_TITLE,
      employeeDomain: employeeDomain, adminEmails: adminEmails,
      actor: actor, startedAt: now, updatedAt: now, completedAt: '',
      errorCount: 0, lastError: '', lastErrorAt: ''
    };
    vNextAdminAppendTemplateJournal_(hub, {
      operationId: operationId, releaseId: releaseId, modelReleaseId: modelReleaseId,
      previousReleaseId: initialPair.releaseId, previousModelReleaseId: initialPair.modelReleaseId,
      templateSpreadsheetId: '', phase: 'PILOT_RECOVERY_INITIALIZED', status: 'SUCCEEDED',
      detail: {
        sourceSpreadsheetId: source.spreadsheet.getId(), sourceReleaseId: source.sourceReleaseId,
        sourceManifestSha256: sourceManifestSha256, clientRuntimeSha256: bundle.sha256,
        portalRuntimeSha256: portalBundle.sha256,
        sourceManifestSchema: source.sourceManifestSchema,
        sourceStoredManifestSha256: source.sourceStoredManifestSha256,
        legacyV2BridgeUsed: source.legacyV2BridgeUsed === true,
        bridgeBasis: source.legacyV2BridgeUsed ? 'ADMIN_ATTESTED_V2_TO_V3_USED_ENVELOPE' : ''
      }
    });
    vNextAdminPortalPilotRecoverySave_(state);
    return state;
  });
}

/**
 * Resolves the UI source without re-hashing a legacy V2 full grid. V2 is the
 * only bridgeable schema and requires an explicit Admin attestation plus exact
 * agreement among the canonical pair, RELEASES, BOOK_REGISTRY, local routing,
 * bound runtime and private ACL. The new V3 used-envelope hash is still always
 * calculated and becomes the immutable identity of the new release.
 */
function vNextAdminPortalPilotRecoveryResolveTemplateSource_(hub, request) {
  const req = request && typeof request === 'object' ? request : {};
  const draftId = vNextAdminText_(req.templateDraftSpreadsheetId);
  if (draftId) {
    const draft = vNextAdminResolveTemplateUiSource_(hub, draftId);
    return Object.assign({}, draft, {
      manifestSha256: vNextAdminTemplateUiManifestHash_(draft.spreadsheet),
      sourceManifestSchema: VN_ADMIN_TEMPLATE_MANIFEST_SCHEMA,
      sourceStoredManifestSha256: '', legacyV2BridgeUsed: false
    });
  }
  const pair = vNextAdminReadActiveReleasePair_(hub);
  const release = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.RELEASES).rows.find(function (row) {
    return String(row.release_id || '') === pair.releaseId;
  });
  if (!release || String(release.status || '').toUpperCase() !== 'ACTIVE' ||
      String(release.template_spreadsheet_id || '') !== pair.templateSpreadsheetId) {
    throw new Error('Canonical ACTIVE pair does not exactly match the ACTIVE Template Release.');
  }
  const activeModel = vNextAdminLatestModelRelease_(hub, pair.modelReleaseId);
  if (!activeModel || String(activeModel.status || '').toUpperCase() !== 'ACTIVE' ||
      String(activeModel.template_version || '') !== pair.releaseId) {
    throw new Error('Canonical ACTIVE pair does not exactly match the ACTIVE Model Release.');
  }
  vNextAdminAssertModelReleaseChecksPassed_(activeModel);
  vNextAdminAssertModelTemplateCompatibility_(hub, activeModel, release);
  if (!vNextAdminActiveReleasePairMirrorsExact_(hub, release, activeModel)) {
    throw new Error('Canonical ACTIVE pair caches do not exactly match RELEASES/MODEL_RELEASE.');
  }
  const schema = String(release.template_manifest_schema || '');
  if ([VN_ADMIN_TEMPLATE_MANIFEST_SCHEMA, VN_ADMIN_TEMPLATE_MANIFEST_SCHEMA_V2].indexOf(schema) < 0) {
    throw new Error('Portal Pilot recovery accepts only V2 or V3 Template UI manifests.');
  }
  const storedManifestSha256 = vNextAdminRequiredText_(release.template_content_sha256,
    'activeRelease.template_content_sha256');
  if (!/^[a-f0-9]{64}$/.test(storedManifestSha256)) {
    throw new Error('ACTIVE Template stored manifest hash is invalid.');
  }
  if (schema === VN_ADMIN_TEMPLATE_MANIFEST_SCHEMA_V2 &&
      (req.attestationConfirmed !== true || !vNextAdminText_(req.evidenceArtifact))) {
    throw new Error('V2→V3 bridge requires the explicit reviewed Admin attestation.');
  }
  const spreadsheetId = vNextAdminRequiredText_(release.template_spreadsheet_id,
    'activeRelease.template_spreadsheet_id');
  const scriptId = vNextAdminRequiredText_(release.template_script_id,
    'activeRelease.template_script_id');
  const runtimeVersion = vNextAdminRequiredText_(release.client_runtime_version,
    'activeRelease.client_runtime_version');
  const runtimeSha256 = vNextAdminRequiredText_(release.client_runtime_sha256,
    'activeRelease.client_runtime_sha256');
  if (!/^[a-f0-9]{64}$/.test(runtimeSha256)) {
    throw new Error('ACTIVE Template runtime hash is invalid.');
  }
  const registry = vNextAdminFindRegistryRow_(hub, function (row) {
    return String(row.mode || '').toUpperCase() === 'TEMPLATE' &&
      String(row.spreadsheet_id || '') === spreadsheetId;
  });
  if (!registry || String(registry.status || '').toUpperCase() !== 'ACTIVE' ||
      String(registry.template_release_id || '') !== pair.releaseId ||
      String(registry.client_script_id || '') !== scriptId ||
      String(registry.client_runtime_version || '') !== runtimeVersion ||
      String(registry.client_runtime_sha256 || '') !== runtimeSha256) {
    throw new Error('ACTIVE Template BOOK_REGISTRY identity does not match RELEASES.');
  }
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const bookConfig = vNextAdminReadKeyValueSheet_(spreadsheet, VN_ADMIN_BOOK_CONFIG_SHEET);
  const systemConfig = vNextAdminReadKeyValueSheet_(spreadsheet, VN_ADMIN_SYSTEM_CONFIG_SHEET);
  if (vNextDetectBookMode_(spreadsheet) !== 'TEMPLATE' ||
      String(bookConfig.book_id || '') !== String(registry.book_id || '') ||
      String(bookConfig.version || '') !== pair.releaseId ||
      String(bookConfig.client_runtime_version || '') !== runtimeVersion ||
      String(bookConfig.client_runtime_bundle_sha256 || '') !== runtimeSha256 ||
      String(systemConfig.mode || '').toUpperCase() !== 'TEMPLATE' ||
      String(systemConfig.book_id || '') !== String(registry.book_id || '') ||
      String(systemConfig.active_release_id || '') !== pair.releaseId) {
    throw new Error('ACTIVE Template local routing identity does not match Hub canonical records.');
  }
  vNextClientRuntimeAssertBoundParent_(scriptId, spreadsheetId);
  const hubConfig = vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET);
  const adminEmails = vNextAdminMergeEmails_(hubConfig.admin_emails,
    vNextGetRuntimeConfig_().VNEXT_ADMIN_EMAILS, vNextAdminActor_());
  vNextAdminAssertPrivateAdminFile_(DriveApp.getFileById(spreadsheetId), adminEmails);
  const v3ManifestSha256 = vNextAdminTemplateUiManifestHash_(spreadsheet);
  if (schema === VN_ADMIN_TEMPLATE_MANIFEST_SCHEMA && storedManifestSha256 !== v3ManifestSha256) {
    throw new Error('ACTIVE V3 Template visible content differs from its immutable manifest.');
  }
  return {
    spreadsheet: spreadsheet, sourceReleaseId: pair.releaseId, draftId: '', isDraft: false,
    adminEmails: adminEmails, manifestSha256: v3ManifestSha256,
    sourceManifestSchema: schema, sourceStoredManifestSha256: storedManifestSha256,
    legacyV2BridgeUsed: schema === VN_ADMIN_TEMPLATE_MANIFEST_SCHEMA_V2
  };
}

function vNextAdminPortalPilotRecoveryTemplateCreate_(hub, priorState) {
  return vNextAdminWithScriptLock_('portal-pilot-template-create', function () {
    const state = vNextAdminPortalPilotRecoveryReloadPhase_(hub, priorState, 'TEMPLATE_CREATE');
    const existingRelease = vNextAdminPortalPilotRecoveryFindRelease_(hub, state);
    if (existingRelease) {
      state.templateSpreadsheetId = String(existingRelease.template_spreadsheet_id || '');
      state.templateScriptId = String(existingRelease.template_script_id || '');
      const registry = vNextAdminFindRegistryRow_(hub, function (row) {
        return String(row.mode || '') === 'TEMPLATE' &&
          String(row.spreadsheet_id || '') === state.templateSpreadsheetId;
      });
      if (registry) state.templateBookId = String(registry.book_id || state.templateBookId);
      if (registry && String(existingRelease.status || '').toUpperCase() === 'ACTIVE') {
        return vNextAdminPortalPilotRecoveryAdvance_(hub, state, 'MODEL_REGISTER',
          'PILOT_TEMPLATE_RELEASE_REUSED', { reused: true, active: true, releaseId: state.releaseId });
      }
      return vNextAdminPortalPilotRecoveryAdvance_(hub, state, 'TEMPLATE_REGISTER',
        'PILOT_TEMPLATE_RELEASE_FOUND', { reused: true, releaseId: state.releaseId });
    }
    if (state.templateSpreadsheetId && state.templateScriptId) {
      vNextAdminPortalPilotRecoveryAssertTemplateContainer_(state);
      return vNextAdminPortalPilotRecoveryAdvance_(hub, state, 'TEMPLATE_INITIALIZE',
        'PILOT_TEMPLATE_CONTAINER_READY', { reused: true, spreadsheetId: state.templateSpreadsheetId });
    }
    vNextAdminPortalPilotRecoveryQuarantineUnknownFiles_(hub, state.templateFolderId,
      state.templateStagingTitle, state.adminEmails, state.operationId, 'TEMPLATE');
    state.templateCreateAttempt = Number(state.templateCreateAttempt || 0) + 1;
    state.templateStagingTitle = 'Forecast vNext Master Template Recovery [' + state.releaseId + '] [' +
      state.operationId.slice(-8) + '-A' + state.templateCreateAttempt + ']';
    vNextAdminPortalPilotRecoverySave_(state); // durable before the external create call
    const created = vNextClientRuntimeCreateBoundSpreadsheet_({
      title: state.templateStagingTitle, folderId: state.templateFolderId
    });
    if (String(created.runtimeVersion || '') !== state.clientRuntimeVersion ||
        String(created.bundleSha256 || '') !== state.clientRuntimeSha256) {
      throw new Error('Created Template runtime does not match the pinned recovery bundle.');
    }
    state.templateSpreadsheetId = String(created.spreadsheetId || '');
    state.templateScriptId = String(created.scriptId || '');
    vNextAdminPortalPilotRecoveryAssertTemplateContainer_(state);
    vNextAdminEnforcePrivateFileAcl_(DriveApp.getFileById(state.templateSpreadsheetId), state.adminEmails);
    return vNextAdminPortalPilotRecoveryAdvance_(hub, state, 'TEMPLATE_INITIALIZE',
      'PILOT_TEMPLATE_CONTAINER_READY', {
        reused: false, spreadsheetId: state.templateSpreadsheetId, scriptId: state.templateScriptId,
        attempt: state.templateCreateAttempt
      });
  });
}

function vNextAdminPortalPilotRecoveryTemplateInitialize_(hub, priorState) {
  return vNextAdminWithScriptLock_('portal-pilot-template-initialize', function () {
    const state = vNextAdminPortalPilotRecoveryReloadPhase_(hub, priorState, 'TEMPLATE_INITIALIZE');
    vNextAdminPortalPilotRecoveryAssertTemplateContainer_(state);
    const template = SpreadsheetApp.openById(state.templateSpreadsheetId);
    vNextAdminInitializeTemplate_(template, {
      bookId: state.templateBookId, releaseId: state.releaseId,
      clientRuntimeVersion: state.clientRuntimeVersion, clientRuntimeSha256: state.clientRuntimeSha256,
      adminEmails: state.adminEmails, actor: state.actor, now: new Date(state.startedAt), resetCopied: true
    });
    vNextAdminEnforcePrivateFileAcl_(DriveApp.getFileById(state.templateSpreadsheetId), state.adminEmails);
    return vNextAdminPortalPilotRecoveryAdvance_(hub, state, 'TEMPLATE_COPY',
      'PILOT_TEMPLATE_INITIALIZED', { spreadsheetId: state.templateSpreadsheetId });
  });
}

function vNextAdminPortalPilotRecoveryTemplateCopy_(hub, priorState) {
  return vNextAdminWithScriptLock_('portal-pilot-template-copy', function () {
    const state = vNextAdminPortalPilotRecoveryReloadPhase_(hub, priorState, 'TEMPLATE_COPY');
    vNextAdminPortalPilotRecoveryAssertSource_(state);
    vNextAdminPortalPilotRecoveryAssertTemplateContainer_(state);
    const source = SpreadsheetApp.openById(state.sourceSpreadsheetId);
    const template = SpreadsheetApp.openById(state.templateSpreadsheetId);
    const copied = vNextAdminCopyAndVerifyTemplateUi_(source, template);
    if (String(copied.manifestSha256 || '') !== state.sourceManifestSha256) {
      throw new Error('Recovery Template manifest differs from the pinned UI source.');
    }
    vNextAdminWriteBookConfig_(template, {
      state: 'TEMPLATE_STAGED', template_kind: 'IMMUTABLE_STAGED',
      source_template_release_id: state.sourceReleaseId,
      template_manifest_schema: VN_ADMIN_TEMPLATE_MANIFEST_SCHEMA,
      template_manifest_sha256: state.sourceManifestSha256,
      updated_at: new Date(), updated_by: state.actor
    });
    vNextAdminEnforcePrivateFileAcl_(DriveApp.getFileById(state.templateSpreadsheetId), state.adminEmails);
    return vNextAdminPortalPilotRecoveryAdvance_(hub, state, 'TEMPLATE_REGISTER',
      'PILOT_TEMPLATE_UI_VERIFIED', {
        spreadsheetId: state.templateSpreadsheetId, manifestSha256: state.sourceManifestSha256
      });
  });
}

function vNextAdminPortalPilotRecoveryTemplateRegister_(hub, priorState) {
  return vNextAdminWithScriptLock_('portal-pilot-template-register', function () {
    const state = vNextAdminPortalPilotRecoveryReloadPhase_(hub, priorState, 'TEMPLATE_REGISTER');
    let release = vNextAdminPortalPilotRecoveryFindRelease_(hub, state);
    if (!release) {
      vNextAdminPortalPilotRecoveryAssertSource_(state);
      vNextAdminPortalPilotRecoveryAssertTemplateReady_(state);
      release = vNextAdminRegisterRelease_(hub, {
        release_id: state.releaseId, release_name: state.releaseId, status: 'STAGED',
        template_spreadsheet_id: state.templateSpreadsheetId,
        schema_version: vNextAdminClientSchemaVersion_(), engine_version: state.engineVersion,
        ux_version: typeof VNEXT_UX_CONFIG_ !== 'undefined' ? VNEXT_UX_CONFIG_.VERSION || '' : '',
        admin_version: VN_ADMIN_SCHEMA_VERSION,
        client_runtime_version: state.clientRuntimeVersion,
        client_runtime_sha256: state.clientRuntimeSha256,
        template_content_sha256: state.sourceManifestSha256,
        template_manifest_schema: VN_ADMIN_TEMPLATE_MANIFEST_SCHEMA,
        template_script_id: state.templateScriptId,
        created_at: new Date(state.startedAt), created_by: state.actor, activated_at: '', note: state.reason
      });
    }
    state.templateSpreadsheetId = String(release.template_spreadsheet_id || state.templateSpreadsheetId);
    state.templateScriptId = String(release.template_script_id || state.templateScriptId);
    vNextAdminPortalPilotRecoveryAssertTemplateReady_(state);
    let registry = vNextAdminFindRegistryRow_(hub, function (row) {
      return String(row.mode || '') === 'TEMPLATE' &&
        String(row.spreadsheet_id || '') === state.templateSpreadsheetId;
    });
    if (!registry) {
      const conflictingBook = vNextAdminFindRegistryRow_(hub, function (row) {
        return String(row.book_id || '') === state.templateBookId;
      });
      if (conflictingBook) throw new Error('Recovery templateBookId is already used by another workbook.');
      const template = SpreadsheetApp.openById(state.templateSpreadsheetId);
      vNextAdminRegisterBook_(hub, {
        book_id: state.templateBookId, mode: 'TEMPLATE', client_id: '', client_name: '', fiscal_year: '',
        spreadsheet_id: state.templateSpreadsheetId, spreadsheet_url: template.getUrl(),
        client_script_id: state.templateScriptId,
        client_runtime_version: state.clientRuntimeVersion,
        client_runtime_sha256: state.clientRuntimeSha256,
        template_release_id: state.releaseId, schema_version: vNextAdminClientSchemaVersion_(),
        state: 'TEMPLATE_STAGED', status: 'STAGED', health_status: 'OK',
        health_code: 'RELEASE_STAGED', last_health_at: new Date(),
        forecast_owner_emails: state.adminEmails.join(','), editor_emails: state.adminEmails.join(','),
        viewer_emails: '', created_at: new Date(state.startedAt), created_by: state.actor,
        updated_at: new Date(), note: vNextAdminCanonicalJson_({
          templateKind: 'IMMUTABLE_STAGED', sourceReleaseId: state.sourceReleaseId,
          manifestSchema: VN_ADMIN_TEMPLATE_MANIFEST_SCHEMA,
          manifestSha256: state.sourceManifestSha256,
          runtimeVersion: state.clientRuntimeVersion, runtimeSha256: state.clientRuntimeSha256,
          recoveryOperationId: state.operationId
        })
      });
    } else {
      state.templateBookId = String(registry.book_id || state.templateBookId);
    }
    const stageOperationId = 'TPL-STAGE-' + vNextAdminSha256_(vNextAdminCanonicalJson_({
      releaseId: state.releaseId, previousReleaseId: state.initialReleaseId,
      sourceManifestSha256: state.sourceManifestSha256,
      runtimeSha256: state.clientRuntimeSha256, reason: state.reason
    })).slice(0, 24).toUpperCase();
    vNextAdminAppendTemplateJournal_(hub, {
      operationId: stageOperationId, releaseId: state.releaseId,
      previousReleaseId: state.initialReleaseId,
      templateSpreadsheetId: state.templateSpreadsheetId,
      phase: 'STAGED_VERIFIED', status: 'SUCCEEDED', detail: {
        sourceSpreadsheetId: state.sourceSpreadsheetId, sourceReleaseId: state.sourceReleaseId,
        manifestSchema: VN_ADMIN_TEMPLATE_MANIFEST_SCHEMA,
        manifestSha256: state.sourceManifestSha256,
        runtimeVersion: state.clientRuntimeVersion, runtimeSha256: state.clientRuntimeSha256,
        reason: state.reason
      }
    });
    vNextAdminWriteAudit_(hub, 'RECOVER_STAGE_TEMPLATE_RELEASE', 'RELEASE', state.releaseId, 'SUCCESS', {
      recoveryOperationId: state.operationId, templateSpreadsheetId: state.templateSpreadsheetId,
      templateScriptId: state.templateScriptId, clientRuntimeSha256: state.clientRuntimeSha256,
      templateManifestSha256: state.sourceManifestSha256
    });
    return vNextAdminPortalPilotRecoveryAdvance_(hub, state, 'MODEL_REGISTER',
      'PILOT_TEMPLATE_REGISTERED', {
        releaseId: state.releaseId, spreadsheetId: state.templateSpreadsheetId
      });
  });
}

function vNextAdminPortalPilotRecoveryModelRegister_(hub, priorState) {
  const state = vNextAdminPortalPilotRecoveryReloadPhase_(hub, priorState, 'MODEL_REGISTER');
  const existing = vNextAdminLatestModelRelease_(hub, state.modelReleaseId);
  if (existing) {
    vNextAdminPortalPilotRecoveryAssertModel_(state, existing);
  } else {
    const checks = vNextAdminPortalPilotRecoveryChecks_(state);
    vNextAdminRegisterModelRelease({
      modelReleaseId: state.modelReleaseId, modelVersion: state.engineVersion,
      templateVersion: state.releaseId, parameters: state.parameters,
      backtest: checks.backtest, canary: checks.canary,
      attestationConfirmed: true, note: state.modelNote
    });
  }
  return vNextAdminPortalPilotRecoveryCommitExternalPhase_(hub, state, 'MODEL_REGISTER',
    'PAIR_ACTIVATE', 'PILOT_MODEL_REGISTERED', { modelReleaseId: state.modelReleaseId });
}

function vNextAdminPortalPilotRecoveryPairActivate_(hub, priorState) {
  const state = vNextAdminPortalPilotRecoveryReloadPhase_(hub, priorState, 'PAIR_ACTIVATE');
  const current = vNextAdminReadActiveReleasePair_(hub);
  const targetActive = current.releaseId === state.releaseId && current.modelReleaseId === state.modelReleaseId;
  if (!targetActive && (current.releaseId !== state.initialReleaseId ||
      current.modelReleaseId !== state.initialModelReleaseId)) {
    throw new Error('準備中に別のTemplate・Model pairが有効化されました。上書きせず停止します。');
  }
  vNextAdminActivateReleasePair({
    releaseId: state.releaseId, modelReleaseId: state.modelReleaseId,
    reason: '申請入口 Pilot の Template・Model pair を有効化',
    expectedActiveReleaseId: state.initialReleaseId,
    expectedActiveModelReleaseId: state.initialModelReleaseId,
    operationId: 'PAIR-ACT-' + vNextAdminSha256_(vNextAdminCanonicalJson_({
      releaseId: state.releaseId, modelReleaseId: state.modelReleaseId,
      previousReleaseId: state.initialReleaseId,
      previousModelReleaseId: state.initialModelReleaseId,
      reason: '社員ポータルPilotのTemplate・Model pairを有効化'
    })).slice(0, 24).toUpperCase()
  });
  return vNextAdminPortalPilotRecoveryCommitExternalPhase_(hub, state, 'PAIR_ACTIVATE',
    'PORTAL_CREATE', 'PILOT_PAIR_ACTIVATED', {
      releaseId: state.releaseId, modelReleaseId: state.modelReleaseId
    });
}

function vNextAdminPortalPilotRecoveryPortalCreate_(hub, priorState) {
  return vNextAdminWithScriptLock_('portal-pilot-portal-create', function () {
    const state = vNextAdminPortalPilotRecoveryReloadPhase_(hub, priorState, 'PORTAL_CREATE');
    const configured = vNextAdminPortalPilotRecoveryConfiguredPortal_(hub);
    if (configured) {
      vNextAdminPortalPilotRecoveryAssertConfiguredPortal_(state, configured);
      state.portalId = configured.portalId;
      state.portalSpreadsheetId = configured.spreadsheetId;
      state.portalScriptId = configured.scriptId;
      return vNextAdminPortalPilotRecoveryAdvance_(hub, state, 'PORTAL_REGISTER',
        'PILOT_PORTAL_FOUND', { reused: true, spreadsheetId: configured.spreadsheetId });
    }
    if (state.portalSpreadsheetId && state.portalScriptId) {
      vNextAdminPortalPilotRecoveryAssertPortalContainer_(state);
      return vNextAdminPortalPilotRecoveryAdvance_(hub, state, 'PORTAL_REGISTER',
        'PILOT_PORTAL_CONTAINER_READY', { reused: true, spreadsheetId: state.portalSpreadsheetId });
    }
    if (!state.portalFolderId) {
      const folder = vNextAdminPrepareClientDestinationFolder_(hub, '',
        'Employee-Portal-' + state.operationId.slice(-8), state.adminEmails);
      state.portalFolderId = folder.getId();
      vNextAdminPortalPilotRecoverySave_(state);
    }
    vNextAdminPortalPilotRecoveryQuarantineUnknownFiles_(hub, state.portalFolderId,
      state.portalStagingTitle, state.adminEmails, state.operationId, 'PORTAL');
    state.portalCreateAttempt = Number(state.portalCreateAttempt || 0) + 1;
    state.portalStagingTitle = state.portalFinalTitle + ' [準備中 ' + state.operationId.slice(-8) +
      '-A' + state.portalCreateAttempt + ']';
    vNextAdminPortalPilotRecoverySave_(state); // durable before the external create call
    const created = vNextPortalRuntimeCreateBoundSpreadsheet_({
      title: state.portalStagingTitle, folderId: state.portalFolderId
    });
    if (String(created.runtimeVersion || '') !== state.portalRuntimeVersion ||
        String(created.bundleSha256 || '') !== state.portalRuntimeSha256) {
      throw new Error('Created Portal runtime does not match the pinned recovery bundle.');
    }
    state.portalSpreadsheetId = String(created.spreadsheetId || '');
    state.portalScriptId = String(created.scriptId || '');
    vNextAdminPortalPilotRecoveryAssertPortalContainer_(state);
    vNextAdminEnforcePrivateFileAcl_(DriveApp.getFileById(state.portalSpreadsheetId), state.adminEmails);
    return vNextAdminPortalPilotRecoveryAdvance_(hub, state, 'PORTAL_REGISTER',
      'PILOT_PORTAL_CONTAINER_READY', {
        reused: false, spreadsheetId: state.portalSpreadsheetId,
        scriptId: state.portalScriptId, attempt: state.portalCreateAttempt
      });
  });
}

function vNextAdminPortalPilotRecoveryPortalRegister_(hub, priorState) {
  return vNextAdminWithScriptLock_('portal-pilot-portal-register', function () {
    const state = vNextAdminPortalPilotRecoveryReloadPhase_(hub, priorState, 'PORTAL_REGISTER');
    const configured = vNextAdminPortalPilotRecoveryConfiguredPortal_(hub);
    if (configured) {
      vNextAdminPortalPilotRecoveryAssertConfiguredPortal_(state, configured);
      state.portalId = configured.portalId;
      state.portalSpreadsheetId = configured.spreadsheetId;
      state.portalScriptId = configured.scriptId;
      const existingFile = DriveApp.getFileById(configured.spreadsheetId);
      vNextAdminApplyEmployeeFileSharing_(existingFile, {
        targetMode: 'PORTAL', accessPolicy: 'INTERNAL_OPEN', internalDomain: state.employeeDomain
      });
      vNextAdminAssertEmployeeFileSharing_(existingFile, {
        targetMode: 'PORTAL', accessPolicy: 'INTERNAL_OPEN', internalDomain: state.employeeDomain
      });
      vNextAdminRefreshPortalDirectory_(hub, configured.spreadsheet);
      return vNextAdminPortalPilotRecoveryAdvance_(hub, state, 'VERIFY',
        'PILOT_PORTAL_REGISTERED', {
          reused: true, portalId: state.portalId, spreadsheetId: state.portalSpreadsheetId,
          runtimeSha256: state.portalRuntimeSha256
        });
    }
    vNextAdminPortalPilotRecoveryAssertPortalContainer_(state);
    const portal = SpreadsheetApp.openById(state.portalSpreadsheetId);
    const file = DriveApp.getFileById(state.portalSpreadsheetId);
    try {
      vNextAdminInitializePortal_(portal, {
        portalId: state.portalId, employeeDomain: state.employeeDomain,
        runtimeVersion: state.portalRuntimeVersion, runtimeSha256: state.portalRuntimeSha256,
        actor: state.actor, adminEmails: state.adminEmails
      });
      vNextAdminWriteSystemConfig_(hub, {
        portal_id: state.portalId, portal_spreadsheet_id: state.portalSpreadsheetId,
        portal_script_id: state.portalScriptId, portal_runtime_version: state.portalRuntimeVersion,
        portal_runtime_sha256: state.portalRuntimeSha256, employee_domain: state.employeeDomain
      });
      vNextAdminUpsertObject_(hub, VN_ADMIN_SHEETS.SETTINGS, 'setting_key', 'EMPLOYEE_PORTAL_JSON', {
        setting_key: 'EMPLOYEE_PORTAL_JSON', setting_value: vNextAdminCanonicalJson_({
          portalId: state.portalId, spreadsheetId: state.portalSpreadsheetId,
          scriptId: state.portalScriptId, runtimeVersion: state.portalRuntimeVersion,
          runtimeSha256: state.portalRuntimeSha256, employeeDomain: state.employeeDomain,
          accessPolicy: 'INTERNAL_OPEN'
        }),
        value_type: 'JSON', scope: 'SYSTEM', effective_from: new Date(), updated_at: new Date(),
        updated_by: state.actor, note: '申請入口（管理ハブとは物理分離）'
      });
      PropertiesService.getScriptProperties().setProperties({
        VNEXT_PORTAL_SPREADSHEET_ID: state.portalSpreadsheetId,
        VNEXT_EMPLOYEE_DOMAIN: state.employeeDomain
      }, false);
      vNextAdminRefreshPortalDirectory_(hub, portal);
      vNextAdminApplyEmployeeFileSharing_(file, {
        targetMode: 'PORTAL', accessPolicy: 'INTERNAL_OPEN', internalDomain: state.employeeDomain
      });
      vNextAdminAssertEmployeeFileSharing_(file, {
        targetMode: 'PORTAL', accessPolicy: 'INTERNAL_OPEN', internalDomain: state.employeeDomain
      });
      file.setName(state.portalFinalTitle);
      vNextAdminWriteAudit_(hub, 'RECOVER_PROVISION_EMPLOYEE_PORTAL', 'PORTAL', state.portalId, 'SUCCESS', {
        recoveryOperationId: state.operationId, spreadsheetId: state.portalSpreadsheetId,
        scriptId: state.portalScriptId, runtimeVersion: state.portalRuntimeVersion,
        runtimeSha256: state.portalRuntimeSha256, employeeDomain: state.employeeDomain,
        folderId: state.portalFolderId
      });
      SpreadsheetApp.flush();
    } catch (error) {
      try { vNextAdminEnforcePrivateFileAcl_(file, state.adminEmails); }
      catch (rollbackError) { Logger.log('Recovery Portal ACL rollback failed: %s', String(rollbackError)); }
      try { file.setName('[SETUP FAILED] ' + state.portalFinalTitle); } catch (ignoredRename) {}
      throw error;
    }
    return vNextAdminPortalPilotRecoveryAdvance_(hub, state, 'VERIFY',
      'PILOT_PORTAL_REGISTERED', {
        portalId: state.portalId, spreadsheetId: state.portalSpreadsheetId,
        runtimeSha256: state.portalRuntimeSha256
      });
  });
}

function vNextAdminPortalPilotRecoveryComplete_(hub, priorState) {
  return vNextAdminWithScriptLock_('portal-pilot-verify', function () {
    const state = vNextAdminPortalPilotRecoveryReloadPhase_(hub, priorState, 'VERIFY');
    const pair = vNextAdminReadActiveReleasePair_(hub);
    if (pair.releaseId !== state.releaseId || pair.modelReleaseId !== state.modelReleaseId) {
      throw new Error('Active Template・Model pair changed before final verification.');
    }
    const portal = vNextAdminResolvePortal_(hub);
    if (portal.spreadsheetId !== state.portalSpreadsheetId || portal.scriptId !== state.portalScriptId ||
        portal.runtimeVersion !== state.portalRuntimeVersion || portal.runtimeSha256 !== state.portalRuntimeSha256 ||
        portal.employeeDomain !== state.employeeDomain) {
      throw new Error('Final Employee Portal identity differs from the recovery journal.');
    }
    const file = DriveApp.getFileById(portal.spreadsheetId);
    vNextAdminAssertEmployeeFileSharing_(file, {
      targetMode: 'PORTAL', accessPolicy: 'INTERNAL_OPEN', internalDomain: state.employeeDomain
    });
    vNextAdminRefreshPortalDirectory_(hub, portal.spreadsheet);
    const completed = Object.assign({}, state, {
      phase: 'COMPLETED', status: 'COMPLETED', completedAt: new Date().toISOString(), lastError: ''
    });
    completed.updatedAt = completed.completedAt;
    const result = vNextAdminPortalPilotRecoveryPublicResult_(completed, {
      portalSpreadsheetUrl: portal.spreadsheet.getUrl(), verified: true
    });
    PropertiesService.getScriptProperties().setProperty(
      'VNEXT_LAST_EMPLOYEE_PORTAL_PILOT_RESULT_JSON', vNextAdminCanonicalJson_(result)
    );
    vNextAdminAppendTemplateJournal_(hub, {
      operationId: completed.operationId, releaseId: completed.releaseId, modelReleaseId: completed.modelReleaseId,
      previousReleaseId: completed.initialReleaseId,
      previousModelReleaseId: completed.initialModelReleaseId,
      templateSpreadsheetId: completed.templateSpreadsheetId,
      phase: 'PILOT_RECOVERY_COMPLETED', status: 'SUCCEEDED', detail: {
        portalId: completed.portalId, portalSpreadsheetId: completed.portalSpreadsheetId,
        clientRuntimeSha256: completed.clientRuntimeSha256,
        portalRuntimeSha256: completed.portalRuntimeSha256
      }
    });
    vNextAdminPortalPilotRecoverySave_(completed);
    Logger.log('EMPLOYEE_PORTAL_PILOT_RECOVERY_READY %s', vNextAdminCanonicalJson_(result));
    return { state: completed, detail: { portalSpreadsheetUrl: portal.spreadsheet.getUrl(), verified: true } };
  });
}

function vNextAdminPortalPilotRecoveryVerifyCompleted_(hub, state) {
  const pair = vNextAdminReadActiveReleasePair_(hub);
  const portal = vNextAdminResolvePortal_(hub);
  if (pair.releaseId !== state.releaseId || pair.modelReleaseId !== state.modelReleaseId ||
      portal.spreadsheetId !== state.portalSpreadsheetId || portal.scriptId !== state.portalScriptId ||
      portal.runtimeSha256 !== state.portalRuntimeSha256) {
    throw new Error('Completed Portal Pilot recovery no longer matches the Hub canonical records.');
  }
  const result = vNextAdminPortalPilotRecoveryPublicResult_(state, {
    reused: true, verified: true, portalSpreadsheetUrl: portal.spreadsheet.getUrl()
  });
  PropertiesService.getScriptProperties().setProperty(
    'VNEXT_LAST_EMPLOYEE_PORTAL_PILOT_RESULT_JSON', vNextAdminCanonicalJson_(result)
  );
  return result;
}

function vNextAdminPortalPilotRecoveryCommitExternalPhase_(hub, priorState, expectedPhase, nextPhase, journalPhase, detail) {
  return vNextAdminWithScriptLock_('commit-' + expectedPhase.toLowerCase(), function () {
    const state = vNextAdminPortalPilotRecoveryLoad_();
    vNextAdminPortalPilotRecoveryAssertState_(hub, state);
    if (state.phase !== expectedPhase) {
      if (vNextAdminPortalPilotRecoveryPhaseIndex_(state.phase) >
          vNextAdminPortalPilotRecoveryPhaseIndex_(expectedPhase)) {
        return { state: state, detail: Object.assign({ reused: true, alreadyAdvanced: true }, detail || {}) };
      }
      throw new Error('Portal Pilot recovery phase CAS failed. expected=' + expectedPhase + '; actual=' + state.phase);
    }
    return vNextAdminPortalPilotRecoveryAdvance_(hub, state, nextPhase, journalPhase, detail);
  });
}

function vNextAdminPortalPilotRecoveryAdvance_(hub, state, nextPhase, journalPhase, detail) {
  const currentIndex = vNextAdminPortalPilotRecoveryPhaseIndex_(state.phase);
  const nextIndex = vNextAdminPortalPilotRecoveryPhaseIndex_(nextPhase);
  if (nextIndex <= currentIndex) throw new Error('Portal Pilot recovery cannot move backward or stay in place.');
  vNextAdminAppendTemplateJournal_(hub, {
    operationId: state.operationId, releaseId: state.releaseId, modelReleaseId: state.modelReleaseId,
    previousReleaseId: state.initialReleaseId,
    previousModelReleaseId: state.initialModelReleaseId,
    templateSpreadsheetId: state.templateSpreadsheetId || '',
    phase: journalPhase, status: 'SUCCEEDED', detail: detail || {}
  });
  state.phase = nextPhase;
  state.status = nextPhase === 'COMPLETED' ? 'COMPLETED' : 'IN_PROGRESS';
  state.updatedAt = new Date().toISOString();
  state.lastError = '';
  vNextAdminPortalPilotRecoverySave_(state);
  return { state: state, detail: detail || {} };
}

function vNextAdminPortalPilotRecoveryFindRelease_(hub, state) {
  const release = vNextAdminReadTable_(hub, VN_ADMIN_SHEETS.RELEASES).rows.find(function (row) {
    return String(row.release_id || '') === state.releaseId;
  });
  if (!release) return null;
  if (['STAGED', 'ACTIVE'].indexOf(String(release.status || '').toUpperCase()) < 0 ||
      String(release.client_runtime_version || '') !== state.clientRuntimeVersion ||
      String(release.client_runtime_sha256 || '') !== state.clientRuntimeSha256 ||
      String(release.template_content_sha256 || '') !== state.sourceManifestSha256 ||
      String(release.template_manifest_schema || '') !== VN_ADMIN_TEMPLATE_MANIFEST_SCHEMA ||
      String(release.schema_version || '') !== vNextAdminClientSchemaVersion_() ||
      String(release.engine_version || '') !== state.engineVersion ||
      !String(release.template_spreadsheet_id || '') || !String(release.template_script_id || '')) {
    throw new Error('Existing recovery Release has different immutable content: ' + state.releaseId);
  }
  const template = SpreadsheetApp.openById(String(release.template_spreadsheet_id));
  if (vNextAdminTemplateUiManifestHash_(template) !== state.sourceManifestSha256) {
    throw new Error('Existing recovery Release Template manifest is inconsistent.');
  }
  vNextClientRuntimeAssertBoundParent_(String(release.template_script_id), template.getId());
  vNextAdminAssertPrivateAdminFile_(DriveApp.getFileById(template.getId()), state.adminEmails);
  return release;
}

function vNextAdminPortalPilotRecoveryAssertTemplateContainer_(state) {
  const spreadsheetId = vNextAdminRequiredText_(state.templateSpreadsheetId, 'templateSpreadsheetId');
  const scriptId = vNextAdminRequiredText_(state.templateScriptId, 'templateScriptId');
  vNextClientRuntimeAssertBoundParent_(scriptId, spreadsheetId);
  vNextAdminAssertPrivateAdminFile_(DriveApp.getFileById(spreadsheetId), state.adminEmails);
  return true;
}

function vNextAdminPortalPilotRecoveryAssertTemplateReady_(state) {
  vNextAdminPortalPilotRecoveryAssertTemplateContainer_(state);
  const template = SpreadsheetApp.openById(state.templateSpreadsheetId);
  const config = vNextAdminReadKeyValueSheet_(template, VN_ADMIN_BOOK_CONFIG_SHEET);
  if (vNextDetectBookMode_(template) !== 'TEMPLATE' || String(config.book_id || '') !== state.templateBookId ||
      String(config.version || '') !== state.releaseId ||
      String(config.client_runtime_version || '') !== state.clientRuntimeVersion ||
      String(config.client_runtime_bundle_sha256 || '') !== state.clientRuntimeSha256 ||
      ['IMMUTABLE_STAGED', 'IMMUTABLE_ACTIVE'].indexOf(String(config.template_kind || '')) < 0 ||
      String(config.template_manifest_sha256 || '') !== state.sourceManifestSha256 ||
      vNextAdminTemplateUiManifestHash_(template) !== state.sourceManifestSha256) {
    throw new Error('Recovery Template identity or manifest is incomplete.');
  }
  return true;
}

function vNextAdminPortalPilotRecoveryAssertSource_(state) {
  const source = SpreadsheetApp.openById(vNextAdminRequiredText_(state.sourceSpreadsheetId, 'sourceSpreadsheetId'));
  if (vNextAdminTemplateUiManifestHash_(source) !== state.sourceManifestSha256) {
    throw new Error('Template UI source changed after recovery initialization. Start a new reviewed release instead.');
  }
  return true;
}

function vNextAdminPortalPilotRecoveryChecks_(state) {
  return {
    backtest: {
      status: 'PASS', basis: 'UNCHANGED_ENGINE_AND_PARAMETERS_FROM_ACTIVE_MODEL',
      sourceModelReleaseId: state.initialModelReleaseId, reviewedBy: state.actor,
      evidenceArtifact: state.evidenceArtifact
    },
    canary: {
      status: 'PASS', basis: 'CLIENT_AND_PORTAL_RUNTIME_CONTRACT_TESTS',
      clientRuntimeTests: 10, portalRuntimeTests: 11, integrationContractTests: 'PASS',
      reviewedBy: state.actor, evidenceArtifact: state.evidenceArtifact
    }
  };
}

function vNextAdminPortalPilotRecoveryAssertModel_(state, model) {
  if (['DRAFT', 'ACTIVE'].indexOf(String(model.status || '').toUpperCase()) < 0 ||
      String(model.model_version || '') !== state.engineVersion ||
      String(model.template_version || '') !== state.releaseId ||
      String(model.parameters_json || '') !== vNextAdminCanonicalJson_(
        vNextAdminNormalizeModelParameters_(state.parameters)
      )) {
    throw new Error('Existing Portal Pilot MODEL_RELEASE has different immutable content.');
  }
  vNextAdminAssertModelReleaseChecksPassed_(model);
  return true;
}

function vNextAdminPortalPilotRecoveryConfiguredPortal_(hub) {
  const config = vNextAdminReadKeyValueSheet_(hub, VN_ADMIN_SYSTEM_CONFIG_SHEET);
  if (!String(config.portal_spreadsheet_id || '').trim()) return null;
  return vNextAdminResolvePortal_(hub);
}

function vNextAdminPortalPilotRecoveryAssertConfiguredPortal_(state, portal) {
  if (portal.runtimeVersion !== state.portalRuntimeVersion ||
      portal.runtimeSha256 !== state.portalRuntimeSha256 ||
      portal.employeeDomain !== state.employeeDomain) {
    throw new Error('Configured Employee Portal differs from the pinned recovery runtime/domain.');
  }
  return true;
}

function vNextAdminPortalPilotRecoveryAssertPortalContainer_(state) {
  const spreadsheetId = vNextAdminRequiredText_(state.portalSpreadsheetId, 'portalSpreadsheetId');
  const scriptId = vNextAdminRequiredText_(state.portalScriptId, 'portalScriptId');
  vNextClientRuntimeAssertBoundParent_(scriptId, spreadsheetId);
  return true;
}

/** Unknown container IDs are never guessed. Keep the file private and visibly quarantine it. */
function vNextAdminPortalPilotRecoveryQuarantineUnknownFiles_(hub, folderId, exactTitle, adminEmails, operationId, kind) {
  if (!folderId || !exactTitle) return 0;
  const folder = DriveApp.getFolderById(String(folderId));
  const files = folder.getFilesByName(String(exactTitle));
  let count = 0;
  while (files.hasNext()) {
    const file = files.next();
    vNextAdminEnforcePrivateFileAcl_(file, adminEmails);
    file.setName('[UNVERIFIED ORPHAN ' + String(kind || 'FILE') + '] ' + exactTitle);
    count++;
    vNextAdminWriteAudit_(hub, 'QUARANTINE_PORTAL_PILOT_ORPHAN', String(kind || 'FILE'), file.getId(), 'SUCCESS', {
      recoveryOperationId: operationId, originalTitle: exactTitle,
      policy: 'PRESERVED_PRIVATE_NOT_REUSED_WITHOUT_SCRIPT_ID'
    });
  }
  return count;
}

function vNextAdminPortalPilotRecoveryRecordError_(hub, priorState, error) {
  try {
    vNextAdminWithScriptLock_('record-portal-pilot-recovery-error', function () {
      const state = vNextAdminPortalPilotRecoveryLoad_() || priorState;
      if (!state || state.operationId !== priorState.operationId) return;
      state.errorCount = Number(state.errorCount || 0) + 1;
      state.lastError = String(error && error.message || error).slice(0, 1500);
      state.lastErrorAt = new Date().toISOString();
      state.updatedAt = state.lastErrorAt;
      vNextAdminPortalPilotRecoverySave_(state);
      vNextAdminWriteAudit_(hub, 'PORTAL_PILOT_RECOVERY_PHASE_FAILED', 'RECOVERY', state.operationId, 'FAILED', {
        phase: state.phase, error: state.lastError, errorCount: state.errorCount
      });
    });
  } catch (recordError) {
    Logger.log('Portal Pilot recovery error journal failed: %s', String(recordError && recordError.stack || recordError));
  }
}

function vNextAdminPortalPilotRecoveryReloadPhase_(hub, priorState, expectedPhase) {
  const state = vNextAdminPortalPilotRecoveryLoad_();
  vNextAdminPortalPilotRecoveryAssertState_(hub, state);
  if (!priorState || state.operationId !== priorState.operationId) {
    throw new Error('Portal Pilot recovery operation changed before phase execution.');
  }
  if (state.phase !== expectedPhase) {
    throw new Error('Portal Pilot recovery phase changed. expected=' + expectedPhase + '; actual=' + state.phase);
  }
  return state;
}

function vNextAdminPortalPilotRecoveryAssertState_(hub, state) {
  if (!state || state.schemaVersion !== VNEXT_PORTAL_PILOT_RECOVERY_SCHEMA_ ||
      state.hubSpreadsheetId !== hub.getId() || !state.operationId ||
      !Array.isArray(state.adminEmails) ||
      vNextAdminPortalPilotRecoveryPhaseIndex_(state.phase) < 0) {
    throw new Error('Portal Pilot recovery journal is missing, malformed, or belongs to another Hub.');
  }
  const actor = vNextAdminActor_().toLowerCase();
  if (state.adminEmails.indexOf(actor) < 0) {
    throw new Error('Only a registered recovery Admin can continue this operation.');
  }
  return true;
}

function vNextAdminPortalPilotRecoveryLoad_() {
  const raw = PropertiesService.getScriptProperties().getProperty(VNEXT_PORTAL_PILOT_RECOVERY_PROPERTY_);
  if (!raw) return null;
  const state = vNextAdminParseJson_(raw, null);
  if (!state || Array.isArray(state) || raw !== vNextAdminCanonicalJson_(state)) {
    throw new Error('Portal Pilot recovery Script Property is not canonical JSON.');
  }
  return state;
}

function vNextAdminPortalPilotRecoverySave_(state) {
  state.updatedAt = state.updatedAt || new Date().toISOString();
  PropertiesService.getScriptProperties().setProperty(
    VNEXT_PORTAL_PILOT_RECOVERY_PROPERTY_, vNextAdminCanonicalJson_(state)
  );
  return state;
}

function vNextAdminPortalPilotRecoveryPublicResult_(state, detail) {
  const extra = detail || {};
  const index = vNextAdminPortalPilotRecoveryPhaseIndex_(state.phase);
  const result = {
    ok: true, exists: true, operationId: state.operationId,
    phase: state.phase, phaseLabel: vNextAdminPortalPilotRecoveryPhaseLabel_(state.phase),
    phaseIndex: index, phaseCount: VNEXT_PORTAL_PILOT_RECOVERY_PHASES_.length - 1,
    completed: state.phase === 'COMPLETED', needsAnotherRun: state.phase !== 'COMPLETED',
    activeTemplateReleaseId: state.releaseId, activeModelReleaseId: state.modelReleaseId,
    templateSpreadsheetId: state.templateSpreadsheetId || '',
    portalId: state.portalId || '', portalSpreadsheetId: state.portalSpreadsheetId || '',
    portalSpreadsheetUrl: extra.portalSpreadsheetUrl || (state.portalSpreadsheetId
      ? 'https://docs.google.com/spreadsheets/d/' + state.portalSpreadsheetId + '/edit' : ''),
    errorCount: Number(state.errorCount || 0), lastError: state.lastError || '',
    updatedAt: state.updatedAt || '', completedAt: state.completedAt || ''
  };
  Object.keys(extra).forEach(function (key) {
    if (!Object.prototype.hasOwnProperty.call(result, key)) result[key] = extra[key];
  });
  return vNextAdminJsonSafe_(result);
}

function vNextAdminPortalPilotRecoveryPhaseIndex_(phase) {
  return VNEXT_PORTAL_PILOT_RECOVERY_PHASES_.indexOf(String(phase || '').toUpperCase());
}

function vNextAdminPortalPilotRecoveryPhaseLabel_(phase) {
  const labels = {
    INITIALIZE: '回復journalの初期化', TEMPLATE_CREATE: 'Template containerの作成',
    TEMPLATE_INITIALIZE: 'Template内部構造の初期化', TEMPLATE_COPY: '従業員画面の複製・検証',
    TEMPLATE_REGISTER: 'Template Releaseの登録', MODEL_REGISTER: 'Model Releaseの登録',
    PAIR_ACTIVATE: 'Template・Model pairの有効化', PORTAL_CREATE: '申請入口 container の作成',
    PORTAL_REGISTER: '申請入口の初期化・共有', VERIFY: '正本・共有境界の最終検証',
    COMPLETED: '完了'
  };
  return labels[String(phase || '').toUpperCase()] || String(phase || '不明');
}
