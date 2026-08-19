#!/usr/bin/env node

import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readdir, readFile } from 'node:fs/promises';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const rootFiles = (await readdir(root)).filter(name => /^VNext_.*\.js$/.test(name)).sort();
let engineSandbox;

await checkJavaScriptSyntax();
await checkHtmlScripts();
await checkGlobalFunctionNames();
await runCoreAndEngineTests();
await checkClientSchemaCompatibility();
await checkClientBundleBoundary();
await checkPortalRuntimeBoundary();
await checkAdminRecoveryContracts();
await checkAdminCoverageContracts();
await checkEmployeeResearchUxContracts();
await checkVertexAiGroundingBudgetContracts();
checkSourceContract();

process.stdout.write('PASS vNext integration contract tests\n');

async function checkJavaScriptSyntax() {
  for (const name of ['0_VNext_Naming.js', 'Forecast_Agent.js', ...rootFiles]) {
    new vm.Script(await readFile(path.join(root, name), 'utf8'), { filename: name });
  }
}

async function checkHtmlScripts() {
  const roots = [root, path.join(root, 'client_runtime', 'src')];
  for (const dir of roots) {
    const names = (await readdir(dir)).filter(name => name.endsWith('.html'));
    for (const name of names) {
      const source = await readFile(path.join(dir, name), 'utf8');
      for (const match of source.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/gi)) {
        new vm.Script(match[1], { filename: path.join(path.basename(dir), name) });
      }
    }
  }
}

async function checkGlobalFunctionNames() {
  const definitions = new Map();
  for (const name of rootFiles) {
    const source = await readFile(path.join(root, name), 'utf8');
    for (const match of source.matchAll(/^function\s+([A-Za-z_$][\w$]*)\s*\(/gm)) {
      const files = definitions.get(match[1]) || [];
      files.push(name);
      definitions.set(match[1], files);
    }
  }
  const duplicates = [...definitions].filter(([, files]) => files.length > 1);
  assert.deepEqual(duplicates, [], `duplicate GAS globals: ${JSON.stringify(duplicates)}`);
}

function gasSandbox() {
  return {
    console,
    Logger: { log() {} },
    Utilities: {
      Charset: { UTF_8: 'UTF_8' },
      DigestAlgorithm: { SHA_256: 'SHA_256' },
      computeDigest(_algorithm, value) {
        return [...createHash('sha256').update(String(value), 'utf8').digest()];
      },
      getUuid() { return 'test-uuid'; },
      formatDate(value) {
        const date = new Date(value);
        return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, '0')}-${String(date.getUTCDate()).padStart(2, '0')}`;
      }
    },
    Session: { getScriptTimeZone: () => 'Asia/Tokyo' }
  };
}

async function runCoreAndEngineTests() {
  const sandbox = gasSandbox();
  vm.createContext(sandbox);
  for (const name of ['VNext_Core.js', 'VNext_Engine.js', 'VNext_AI.js', 'VNext_Tests.js']) {
    vm.runInContext(await readFile(path.join(root, name), 'utf8'), sandbox, { filename: name });
  }
  const result = sandbox.runAllVNextTests();
  assert.equal(result.passed, 29);
  assert.equal(result.failed, 0);
  assert.equal(sandbox.testVNextAiDeterministicMapping(), true);
  engineSandbox = sandbox;
}

async function checkClientSchemaCompatibility() {
  const rootSandbox = gasSandbox();
  const clientSandbox = gasSandbox();
  vm.createContext(rootSandbox);
  vm.createContext(clientSandbox);
  vm.runInContext(await readFile(path.join(root, 'VNext_Core.js'), 'utf8'), rootSandbox);
  vm.runInContext(await readFile(path.join(root, 'client_runtime', 'src', 'Client_Core.js'), 'utf8'), clientSandbox);
  const rootSchemas = JSON.parse(JSON.stringify(rootSandbox.VNEXT_CORE.INTERNAL_SHEETS));
  const clientSchemas = JSON.parse(JSON.stringify(clientSandbox.VNEXT_CLIENT_CORE.INTERNAL_SHEETS));
  assert.deepEqual(clientSchemas, rootSchemas, 'Client/Core append-only schemas must have identical headers and order');
}

async function checkEmployeeResearchUxContracts() {
  const sandbox = gasSandbox();
  vm.createContext(sandbox);
  vm.runInContext(await readFile(path.join(root, 'VNext_UX.js'), 'utf8'), sandbox, { filename: 'VNext_UX.js' });
  const rawForecast = {
    runId: 'RUN-UX',
    layers: {
      historyBaseline: 100, commitmentDelta: 10, referenceDelta: 5,
      objectiveForecast: 115, humanDelta: -4, aiDelta: 1, systemRecommended: 112
    },
    annual: { p10: 90, p50: 112, p90: 130 },
    lenses: {
      continuity: { baseAnnualBaseline: 80, fiscalYears: [2019, 2020, 2021, 2022, 2023] },
      changeReference: { peerReferenceDelta: 2, objectiveEventDelta: 3 },
      triangulation: {
        policy: 'INDEPENDENT_REFERENCES_NOT_AUTOMATICALLY_AVERAGED',
        methods: [
          { key: 'RECENT_WEIGHTED_AVERAGE', label: '直近3年度の加重平均', value: 100, assumption: '水準継続', basis: '3年度' },
          { key: 'LINEAR_REGRESSION', label: '線形回帰トレンド', value: 108, assumption: '一定額増加', basis: '5年度' },
          { key: 'DAMPED_CAGR', label: '減衰CAGR', value: 105, assumption: '成長率減衰', basis: 'CAGR' },
          { key: 'INTEGRATED_SIMULATION', label: '統合シミュレーション', value: 112, assumption: '情報統合', basis: 'seed' }
        ]
      }
    },
    evidenceSummary: {
      unknownSpotExpectedAnnual: 20,
      topAiEvidence: [{
        researchAxis: 'DIGITAL_EXECUTION', forecastUse: 'INSIGHT_ONLY', summary: 'DX投資の実行状況を確認',
        humanQuestion: '予算執行部門を確認する', sourceUrl: 'https://example.com/dx'
      }]
    }
  };
  const forecast = sandbox.vNextUxPublicForecast_(rawForecast);
  assert.equal(forecast.layerBreakdown.rows.length, 7);
  assert.equal(forecast.layerBreakdown.checkTotal, 112,
    'The employee layer view must reconcile exactly to the system recommendation');
  assert.equal(forecast.aiEvidence[0].axisLabel, 'DX・業務変革');
  assert.equal(forecast.aiEvidence[0].useLabel, '担当者向け参考');
  assert.equal(forecast.triangulation.methods.length, 4);
  assert.equal(forecast.triangulation.policy, 'INDEPENDENT_REFERENCES_NOT_AUTOMATICALLY_AVERAGED');

  const projection = {
    schemaVersion: 'vnext-public-ai-insights-1', bookId: 'BOOK-UX', generatedAt: '2026-08-14T10:00:00Z',
    insights: [{
      target: '日本法人', researchAxis: 'FINANCIAL_CAPACITY', forecastUse: 'INSIGHT_ONLY',
      projectionStatus: 'INSIGHT_ONLY', summary: '公開情報から投資余力を確認',
      humanQuestion: '日本向け予算の配分を確認する', sourceUrl: 'https://example.com/public',
      sourceDate: '2026-08-14', evidenceQuality: 'A'
    }]
  };
  sandbox.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({
      getSheetByName: () => ({
        getLastRow: () => 2,
        getDataRange: () => ({ getValues: () => [['key', 'value'], ['public_ai_insights_json', JSON.stringify(projection)]] })
      })
    })
  };
  const projectedForecast = sandbox.vNextUxForecastForView_({ bookId: 'BOOK-UX' }, Object.assign({}, rawForecast, {
    evidenceSummary: Object.assign({}, rawForecast.evidenceSummary, { aiUnavailable: true })
  }));
  assert.equal(projectedForecast.systemRecommended, 112, 'public AI insights must never change the forecast amount');
  assert.equal(projectedForecast.aiEvidence.length, 2);
  assert.equal(projectedForecast.aiEvidence[1].useLabel, '担当者向け参考（予測額には未反映）');
  assert.equal(projectedForecast.nextInformation.includes('日本向け予算の配分を確認する'), true);
  assert.match(projectedForecast.warnings.join(' '), /現在の予測額は変更していません/);

  const guidance = await readFile(path.join(root, 'VNext_GuidanceSidebar.html'), 'utf8');
  assert.match(guidance, /data-panel="layers"/);
  assert.match(guidance, /data-panel="insights"/);
  assert.match(guidance, /担当者への確認/);
  const plan = await readFile(path.join(root, 'VNext_PlanSidebar.html'), 'utf8');
  const review = await readFile(path.join(root, 'VNext_ReviewSidebar.html'), 'utf8');
  const input = await readFile(path.join(root, 'VNext_InputSidebar.html'), 'utf8');
  assert.equal(/内容を確認/.test(plan), false, 'Plan submission must be one-step');
  assert.equal(/内容を確認/.test(review), false, 'Review save must be one-step');
  assert.equal(/内容を確認/.test(input), false, 'Evidence save must be one-step');

  const admin = await readFile(path.join(root, 'VNext_Admin.js'), 'utf8');
  assert.match(admin, /function vNextAdminProjectAiInsightsToClient_/);
  assert.match(admin, /public_ai_insights_json/);
  assert.match(admin, /const publicProjection = vNextAdminProjectAiInsightsToClient_/,
    'a successful standalone AI job must refresh the employee-safe projection');
  const sanitizer = admin.slice(
    admin.indexOf('function vNextAdminPublicAiInsightFromRow_'),
    admin.indexOf('function vNextAdminAiEvidenceActiveAt_')
  );
  assert.equal(/aiModel|promptVersion|ruleVersion|parentRequestId/.test(sanitizer), false,
    'employee projection must not expose model, prompt, rule, or request metadata');
  const adminSandbox = {
    vNextAdminParseJson_: value => JSON.parse(value || '{}'),
    vNextAdminSha256_: value => createHash('sha256').update(String(value), 'utf8').digest('hex')
  };
  vm.createContext(adminSandbox);
  vm.runInContext(sanitizer, adminSandbox);
  const safeInsight = adminSandbox.vNextAdminPublicAiInsightFromRow_({
    evidence_id: 'AI-1', target: '日本法人', direction: 'NEUTRAL', source_url: 'https://example.com/source',
    source_date: '2026-08-14', evidence_quality: 'A', applied_amount: 0, cap_applied: 0,
    created_at: '2026-08-14T10:00:00Z', evidence_text: JSON.stringify({
      summary: '公開情報の要約', citationTitle: '一次情報', researchAxis: 'PRODUCT_MARKET',
      signalType: '製品上市', sourceStrength: 'PRIMARY_OFFICIAL', forecastUse: 'INSIGHT_ONLY',
      salesRelevance: 'MEDIUM', humanQuestion: '日本向け実行予算を確認する', aiModel: 'secret-model',
      promptVersion: 'secret-prompt', parentRequestId: 'secret-request'
    })
  });
  assert.equal(safeInsight.forecastUse, 'INSIGHT_ONLY');
  assert.equal(safeInsight.appliedAmount, 0);
  assert.equal(safeInsight.citationTitle, '一次情報');
  assert.equal('aiModel' in safeInsight || 'promptVersion' in safeInsight || 'parentRequestId' in safeInsight, false);
}

async function checkClientBundleBoundary() {
  const sandbox = {};
  vm.createContext(sandbox);
  vm.runInContext(await readFile(path.join(root, 'VNext_ClientRuntimeBundle.js'), 'utf8'), sandbox);
  const bundle = sandbox.VNEXT_CLIENT_RUNTIME_BUNDLE_;
  const clientSourceDir = path.join(root, 'client_runtime', 'src');
  const generatedFiles = [];
  for (const name of (await readdir(clientSourceDir)).filter(name => /\.(?:js|html|json)$/.test(name)).sort()) {
    const extension = path.extname(name);
    generatedFiles.push({
      name: path.basename(name, extension),
      type: extension === '.html' ? 'HTML' : extension === '.json' ? 'JSON' : 'SERVER_JS',
      source: await readFile(path.join(clientSourceDir, name), 'utf8')
    });
  }
  const generatedBundle = {
    version: 'vnext-client-1.8.0',
    sha256: createHash('sha256').update(
      generatedFiles.map(file => `${file.name}\0${file.type}\0${file.source}`).join('\0')
    ).digest('hex'),
    files: generatedFiles
  };
  assert.deepEqual(
    JSON.parse(JSON.stringify(bundle)),
    generatedBundle,
    'The GAS provisioning bundle must exactly match the verified client runtime build'
  );
  assert.equal(bundle.files.length, 10);
  assert.deepEqual(
    JSON.parse(JSON.stringify(bundle.files.map(file => file.name))).sort(),
    ['Client_Bridge', 'Client_Core', 'Client_Entry', 'VNext_GuidanceSidebar', 'VNext_HelpSidebar', 'VNext_InputSidebar', 'VNext_PlanSidebar', 'VNext_ReviewSidebar', 'VNext_UX', 'appsscript'].sort()
  );
  const joined = bundle.files.map(file => `${file.name}\0${file.type}\0${file.source}`).join('\0');
  assert.equal(createHash('sha256').update(joined).digest('hex'), bundle.sha256);
  const source = bundle.files.map(file => file.source).join('\n');
  const employeeSidebarSource = bundle.files
    .filter(file => file.type === 'HTML')
    .map(file => file.source).join('\n');
  assert.equal(employeeSidebarSource.includes(String.raw`/^https?:\/\//i`), false,
    'Apps Script HTML sidebars must not use a URL regex that is truncated by the HTML sandbox');
  assert.match(employeeSidebarSource, /sourceUrl\.split\(':\'\)\[0\]\.toLowerCase\(\)/,
    'Employee sidebars must validate citation schemes without the sandbox-unsafe URL regex');
  assert.match(source, /reason:\s*'forecast_requested:'\s*\+\s*requestId/,
    'The deployed client runtime must bind READY_TO_RUN>RUNNING to the exact requestId');
  for (const forbidden of [
    /Forecast_Agent/, /vNextAdminBootstrap/, /vNextRunForecast_/, /VERTEX_[A-Z_]+/,
    /FORECAST_SOURCE_SPREADSHEET_ID/, /VNEXT_ZAC_SOURCE_SPREADSHEET_ID/,
    /DriveApp/, /UrlFetchApp/, /SpreadsheetApp\.openById/, /PropertiesService/
  ]) {
    assert.equal(forbidden.test(source), false, `Client bundle contains forbidden capability: ${forbidden}`);
  }
  const manifest = JSON.parse(bundle.files.find(file => file.name === 'appsscript').source);
  assert.deepEqual(manifest.oauthScopes, [
    'https://www.googleapis.com/auth/script.container.ui',
    'https://www.googleapis.com/auth/script.scriptapp',
    'https://www.googleapis.com/auth/spreadsheets.currentonly',
    'https://www.googleapis.com/auth/userinfo.email'
  ]);
}

async function checkPortalRuntimeBoundary() {
  const sandbox = gasSandbox();
  vm.createContext(sandbox);
  vm.runInContext(await readFile(path.join(root, 'VNext_PortalRuntimeBundle.js'), 'utf8'), sandbox);
  const bundle = sandbox.VNEXT_PORTAL_RUNTIME_BUNDLE_;
  assert.equal(bundle.version, 'vnext-portal-1.7.8');
  assert.equal(bundle.files.length, 5);
  assert.deepEqual(
    JSON.parse(JSON.stringify(bundle.files.map(file => file.name))).sort(),
    ['Portal_Core', 'Portal_CreateSidebar', 'Portal_Entry', 'Portal_UX', 'appsscript'].sort()
  );
  const sourceDir = path.join(root, 'portal_runtime', 'src');
  const generated = [];
  for (const name of (await readdir(sourceDir)).filter(name => /\.(?:js|html|json)$/.test(name)).sort()) {
    const extension = path.extname(name);
    generated.push({
      name: path.basename(name, extension),
      type: extension === '.html' ? 'HTML' : extension === '.json' ? 'JSON' : 'SERVER_JS',
      source: await readFile(path.join(sourceDir, name), 'utf8')
    });
  }
  const hash = createHash('sha256').update(
    generated.map(file => `${file.name}\0${file.type}\0${file.source}`).join('\0')
  ).digest('hex');
  assert.deepEqual(JSON.parse(JSON.stringify(bundle.files)), generated,
    'Portal GAS bundle must exactly match isolated portal runtime sources');
  assert.equal(bundle.sha256, hash);
  const manifest = JSON.parse(generated.find(file => file.name === 'appsscript').source);
  assert.deepEqual(manifest.oauthScopes.slice().sort(), [
    'https://www.googleapis.com/auth/script.container.ui',
    'https://www.googleapis.com/auth/script.scriptapp',
    'https://www.googleapis.com/auth/spreadsheets.currentonly',
    'https://www.googleapis.com/auth/userinfo.email'
  ].sort());
  const portalSandbox = gasSandbox();
  vm.createContext(portalSandbox);
  vm.runInContext(await readFile(path.join(sourceDir, 'Portal_Core.js'), 'utf8'), portalSandbox);
  assert.equal(portalSandbox.vNextPortalNormalizeClientName_('株式会社 テスト'), 'テスト');
  assert.equal(portalSandbox.vNextPortalSafeBookUrl_('https://example.com/bad'), '');
  assert.equal(portalSandbox.VNEXT_PORTAL.REQUEST_SCHEMA_VERSION, 'vnext-portal-request-2');
  assert.deepEqual(JSON.parse(JSON.stringify(portalSandbox.VNEXT_PORTAL.PREVIEW_INPUT_KEYS)),
    ['clientKey', 'fiscalYear', 'relatedMemberNames']);
  assert.equal(portalSandbox.vNextPortalNormalizeCreationFiscalYear_(2029, new Date('2026-08-12')), 2029,
    'Future fiscal years beyond FY2028 must be accepted by the employee UI contract');
  assert.equal(portalSandbox.vNextPortalNormalizeCreationFiscalYear_(2036, new Date('2026-08-12')), 2036);
  assert.throws(() => portalSandbox.vNextPortalNormalizeCreationFiscalYear_(2037, new Date('2026-08-12')), /一覧から/);
  assert.deepEqual(JSON.parse(JSON.stringify(portalSandbox.vNextPortalNormalizeMemberNames_(['山田 太郎'], true))),
    ['山田 太郎']);
  assert.throws(() => portalSandbox.vNextPortalNormalizeMemberNames_([], true), /1名以上/);
  assert.throws(() => portalSandbox.vNextPortalNormalizeMemberNames_(['山田太郎', '山田 太郎'], true), /同じ関与メンバー/);
  const portalHtml = await readFile(path.join(sourceDir, 'Portal_CreateSidebar.html'), 'utf8');
  assert.equal(/id=["']clientName["']|id=["']clientId["']|id=["']forecastOwnerEmail["']/.test(portalHtml), false,
    'Client name/ID and Forecast Owner must not be employee input fields');
  for (let memberIndex = 1; memberIndex <= 5; memberIndex++) {
    assert.match(portalHtml, new RegExp(`id=["']memberName${memberIndex}["']`));
  }
  assert.match(portalHtml, /<select id="clientKey"/);
  const legacyPortalPayload = {
    clientId:'', clientName:'Legacy Client', fiscalYear:2027,
    forecastOwnerEmail:'owner@example.com', relatedMemberEmails:[],
    requestId:'PORTAL-REQ-LEGACY01', requestType:'CREATE_CLIENT_FY_BOOK',
    requestedAt:'2026-08-12T00:00:00.000Z', requestedBy:'owner@example.com',
    schemaVersion:'vnext-portal-request-1'
  };
  assert.equal(portalSandbox.vNextPortalValidateRequestPayload_(legacyPortalPayload), legacyPortalPayload);
  assert.equal(portalSandbox.vNextPortalAssertRequestRowProjection_({
    fiscal_year:2027, client_id:'', client_name:'Legacy Client',
    forecast_owner_email:'owner@example.com', related_member_emails_json:'[]',
    requested_at:legacyPortalPayload.requestedAt, requested_by:'owner@example.com',
    catalog_key:'', related_member_names_json:''
  }, legacyPortalPayload), true, 'Legacy v1 rows must remain readable after the v2 table migration');
  const adminSidebar = await readFile(path.join(root, 'VNext_AdminSidebar.html'), 'utf8');
  assert.match(adminSidebar, /申請入口を準備する/);
  assert.match(adminSidebar, /vNextAdminProvisionSharedPortal/);
  assert.match(adminSidebar, /vNextAdminRelocateLibraryToSharedDrive/);
  assert.match(adminSidebar, /共有ドライブ「年度予算策定」へ移す/);
  const adminSource = await readFile(path.join(root, 'VNext_Admin.js'), 'utf8');
  assert.match(adminSource, /VN_ADMIN_PORTAL_REQUEST_SCHEMA\s*=\s*'vnext-portal-request-2'/);
  assert.match(adminSource, /function vNextAdminRefreshZacClientCatalog\(/);
  assert.match(adminSource, /function vNextAdminUpdateSharedPortalRuntime\(/);
  assert.match(adminSource, /VN_ADMIN_ZAC_CLIENT_CODE_COLUMN\s*=\s*40/);
  assert.match(adminSource, /VN_ADMIN_ZAC_CLIENT_NAME_COLUMN\s*=\s*41/);
  assert.match(adminSource, /forecastOwnerEmail:\s*requestedBy/,
    'v2 Portal requests must derive Forecast Owner from the authenticated requester');
  const pilotStart = adminSource.indexOf('function vNextAdminPrepareEmployeePortalPilot(request)');
  const pilotEnd = adminSource.indexOf('/**\n * Spreadsheet macro entry', pilotStart);
  const pilot = adminSource.slice(pilotStart, pilotEnd);
  assert.ok(pilotStart >= 0 &&
    pilot.includes('vNextAdminResolveActiveModelReleaseForUpgrade_(hub, initialPair)') &&
    pilot.includes('vNextAdminPublishTemplateRelease({') &&
    pilot.includes('vNextAdminRegisterModelRelease({') &&
    pilot.includes('vNextAdminActivateReleasePair({') &&
    pilot.includes('vNextAdminProvisionSharedPortal({') &&
    pilot.includes("req.attestationConfirmed !== true") &&
    pilot.includes("vNextAdminRequiredText_(req.evidenceArtifact, 'evidenceArtifact')") &&
    pilot.includes('pairAfterStage.releaseId !== initialPair.releaseId') &&
    !pilot.includes('SpreadsheetApp.getUi'),
  'The first Portal pilot needs an idempotent Admin-only path that also works outside Spreadsheet UI context');
  const upgradeResolverStart = adminSource.indexOf('function vNextAdminResolveActiveModelReleaseForUpgrade_(');
  const upgradeResolverEnd = adminSource.indexOf('function vNextAdminResolveActiveModelRelease_(', upgradeResolverStart);
  const upgradeResolver = adminSource.slice(upgradeResolverStart, upgradeResolverEnd);
  assert.ok(upgradeResolverStart >= 0 &&
    upgradeResolver.includes('vNextAdminAssertModelReleaseOwnPair_(model, release)') &&
    !upgradeResolver.includes('VNEXT_ENGINE.VERSION'),
  'Engine upgrades must validate the active source pair against its own immutable version before registering the new deployed Engine');
  assert.match(adminSource, /function vNextAdminAssertModelReleaseOwnPair_\(model, release\)/);
  assert.match(adminSource, /String\(model\.model_version \|\| ''\) !== String\(release\.engine_version \|\| ''\)/);
  assert.match(adminSource, /String\(backtest\.candidateHash \|\| ''\) !== candidateHash/);
  const emptyModelStart = adminSource.indexOf('function vNextAdminEmptyPilotModel_(');
  const emptyModelEnd = adminSource.indexOf('function vNextAdminAssertEmptyPilotReleaseAssets_(', emptyModelStart);
  const emptyModel = adminSource.slice(emptyModelStart, emptyModelEnd);
  assert.ok(emptyModel.includes("String(release.release_id || '') === String(activePair.releaseId || '')") &&
    emptyModel.includes('vNextAdminAssertModelReleaseOwnPair_(model, release)'),
  'Same-URL upgrades must validate the retired source Model against its own immutable Template while keeping the target pair strict');
  assert.ok(adminSource.includes('function vNextAdminPrepareEmployeePortalPilotForManualTest()') &&
    adminSource.includes("answer !== ui.Button.YES") &&
    adminSource.includes('clientRuntimeTests: 10') &&
    adminSource.includes('portalRuntimeTests: 12'),
  'The manual Sheet-macro entry must require explicit Admin attestation and record the tested runtime identities');
  const adminMenuStart = adminSource.indexOf('function vNextBuildAdminMenu_()');
  const adminMenuEnd = adminSource.indexOf('/** Optional best-effort hook', adminMenuStart);
  const adminMenu = adminSource.slice(adminMenuStart, adminMenuEnd);
  assert.ok(adminMenu.includes('VN_ADMIN_MENU_OPEN_SIDEBAR') &&
    adminMenu.includes('VN_ADMIN_MENU_HEALTH_SCAN') &&
    adminMenu.includes('VN_ADMIN_MENU_OPEN_REGISTRY') &&
    adminMenu.includes('addSubMenu') &&
    !adminMenu.includes('vNextAdminMenuRunOperationalCycle'),
    'The Hub top menu is a recovery path plus nested irregular ops, not the daily run-now action');
  assert.ok(adminSource.includes('VN_ADMIN_MENU_NAME = VNEXT_NAMING.MENU') &&
    adminSource.includes("VN_ADMIN_MENU_OPEN_SIDEBAR = '案内を開く'") &&
    adminSource.includes("VN_ADMIN_MENU_RUN_NOW = '申請を今すぐ処理'") &&
    adminSource.includes("VN_ADMIN_MENU_HEALTH_SCAN = '全クライアントの状態点検'") &&
    adminSource.includes("VN_ADMIN_MENU_OPEN_REGISTRY = '登録一覧を開く'"),
    'Admin Hub menu copy must stay role-neutral and keep daily processing in the sidebar');
  assert.ok(adminSource.includes('function vNextAdminInstalledGuidanceOnOpen(') &&
    adminSource.includes('vNextAdminEnsureGuidanceOnOpenTrigger_'),
    'Hub must auto-open guidance from an installable project trigger, not simple onOpen');
  assert.doesNotMatch(adminMenu, /showSidebar/,
    'Simple Hub onOpen must not call authorized Ui.showSidebar');
  assert.match(adminMenu, /vNextAdminLooksLikeHub_/);
  assert.doesNotMatch(adminMenu, /vNextDetectBookMode_|vNextAdminIsRegisteredHub_|vNextAdminHydrateLocalRuntime_/,
    'Hub menu construction must not wait on config, registry, or property hydration');
  const installedStart = adminSource.indexOf('function vNextAdminInstalledGuidanceOnOpen(');
  const installedEnd = adminSource.indexOf('function vNextAdminEnsureGuidanceOnOpenTrigger_(', installedStart);
  const installed = adminSource.slice(installedStart, installedEnd);
  const likeIdx = installed.indexOf('vNextAdminLooksLikeHub_');
  const showIdx = installed.indexOf('showSidebar');
  const detectIdx = installed.indexOf('vNextDetectBookMode_');
  assert.ok(likeIdx >= 0 && showIdx > likeIdx && detectIdx > showIdx,
    'Hub auto-sidebar must appear before config detection');
  assert.doesNotMatch(installed, /vNextAdminIsRegisteredHub_\(/);
  const getModelStart = adminSource.indexOf('function vNextAdminGetSidebarModel()');
  const getModelEnd = adminSource.indexOf('function vNextAdminGetSidebarDetailModel()', getModelStart);
  const getModel = adminSource.slice(getModelStart, getModelEnd);
  assert.ok(getModelEnd > getModelStart, 'Hub sidebar must split first paint from deferred details');
  assert.doesNotMatch(getModel, /vNextAdminPortalRequestsForSidebar_/);
  assert.doesNotMatch(getModel, /vNextAdminLatestModelReleaseSummaries_/);
  assert.doesNotMatch(getModel, /vNextAdminListTemplateDrafts_/);
  assert.doesNotMatch(getModel, /VN_ADMIN_SHEETS\.CATALOG/);
  assert.match(adminSource, /function vNextAdminGetSidebarDetailModel\(\)[\s\S]*vNextAdminPortalRequestsForSidebar_/);
  assert.match(adminSidebar, /vNextAdminGetSidebarDetailModel/);
  assert.match(adminSidebar, /id="adminPanel">/);
  assert.doesNotMatch(adminSidebar, /id="adminPanel" class="hidden"/);
  assert.ok(!adminMenu.includes('vNextAdminContinueEmployeePortalPilotRecoveryForManualTest'),
    'One-time Portal bootstrap recovery must stay out of the normal Admin menu');
  assert.ok(!adminMenu.includes("'vNextAdminPrepareEmployeePortalPilotForManualTest'"),
    'The long-running legacy one-shot Portal pilot action must not remain in the Admin menu');
  const recoverySource = await readFile(path.join(root, 'VNext_PortalPilotRecovery.js'), 'utf8');
  assert.match(recoverySource,
    /function vNextAdminContinueEmployeePortalPilotRecoveryForManualTest\(\)/);
}

async function checkAdminRecoveryContracts() {
  const source = await readFile(path.join(root, 'VNext_Admin.js'), 'utf8');
  const legacyMenuStart = source.indexOf('function vNextBuildLegacySetupMenu_');
  const legacyMenuEnd = source.indexOf('function vNextAdminAuthorizeRuntime()', legacyMenuStart);
  const legacyMenu = source.slice(legacyMenuStart, legacyMenuEnd);
  assert.ok(legacyMenu.includes("createMenu('Forecast vNext 移行')") &&
    legacyMenu.includes("addItem('初回権限を確認・許可', 'vNextAdminAuthorizeRuntime')") &&
    legacyMenu.includes("addItem('必要APIを有効化', 'vNextAdminEnableAppsScriptApi')") &&
    !legacyMenu.includes('vNextAdminAssertRuntimeConfigurator_(') &&
    !legacyMenu.includes('DriveApp.'),
    'Legacy bootstrap menu construction must remain safe inside a simple onOpen trigger');
  assert.ok(source.includes('function vNextAdminAuthorizeRuntime()') &&
    source.includes("vNextAdminGuard_('vNextAdminAuthorizeRuntime'") &&
    source.includes('vNextAdminAssertRuntimeConfigurator_();'),
    'Manifest scope authorization must be available as a direct custom-menu action');
  assert.ok(source.includes('function vNextAdminEnableAppsScriptApi()') &&
    source.includes('vNextClientRuntimeEnableRequiredAppsScriptApi_()'),
    'The admin menu must expose the fixed-service Apps Script API enablement helper');
  assert.ok(source.includes('function vNextAdminEnableGeneratedHubAppsScriptApi(request)') &&
    source.includes('vNextClientRuntimeEnableAppsScriptApiForProjectNumber_(projectNumber)') &&
    source.includes("service: 'script.googleapis.com'"),
    'A clean generated Hub must have a source-authorized, fixed-service API recovery path');
  const provisioning = await readFile(path.join(root, 'VNext_ClientRuntimeProvisioning.js'), 'utf8');
  const generatedApiStart = provisioning.indexOf('function vNextClientRuntimeEnableAppsScriptApiForProjectNumber_(');
  const generatedApiEnd = provisioning.indexOf('function vNextClientRuntimeVerifiedBundle_', generatedApiStart);
  const generatedApi = provisioning.slice(generatedApiStart, generatedApiEnd);
  assert.ok(generatedApi.includes("'/services/script.googleapis.com'") &&
    generatedApi.includes("String(verifyBody.state || '').toUpperCase() === 'ENABLED'") &&
    !generatedApi.includes('req.service'),
    'Generated runtime recovery must enable and verify only script.googleapis.com');
  const evidenceStart = source.indexOf('function vNextAdminValidateClientEvidenceRows_');
  const evidenceEnd = source.indexOf('function vNextAdminValidateClientStateRows_', evidenceStart);
  const evidenceFunction = source.slice(evidenceStart, evidenceEnd);
  assert.ok(evidenceFunction.indexOf('if (!rows || !rows.length) return true;') >= 0,
    'Empty first-sync evidence must not require a pre-existing registry row');

  const stateStart = source.indexOf('function vNextAdminSetClientState_');
  const stateEnd = source.indexOf('function vNextAdminAppendStateEvent_', stateStart);
  const stateFunction = source.slice(stateStart, stateEnd);
  const hubWrite = stateFunction.indexOf("vNextAdminAppendCoreRowsNoLock_(hub, 'STATE_EVENT'");
  const clientWrite = stateFunction.indexOf("vNextAdminAppendCoreRowsNoLock_(client, 'STATE_EVENT'");
  assert.ok(hubWrite >= 0 && clientWrite > hubWrite,
    'Admin state events must be persisted to the Hub before the Client copy');
  assert.ok(source.includes('function vNextAdminRetryOfficialClientSync('),
    'Official Client propagation must have an explicit recovery API');
  assert.ok(source.includes("status: 'FAILED', failedAt:"),
    'A failed review start must release its duplicate-prevention claim');
  const bootstrapStart = source.indexOf('function vNextAdminBootstrapFromCurrent(');
  const bootstrapEnd = source.indexOf('/** Re-run the idempotent Hub initialization', bootstrapStart);
  const bootstrap = source.slice(bootstrapStart, bootstrapEnd);
  assert.ok(bootstrap.includes('vNextAdminRuntimeCreateBoundSpreadsheet_(') &&
    bootstrap.includes('vNextClientRuntimeCreateBoundSpreadsheet_('),
    'Bootstrap must create clean, known-script Admin and Client runtime containers');
  assert.equal(bootstrap.includes('.makeCopy('), false,
    'Bootstrap must never copy the legacy bound script into the Admin Hub or Template');
  const recoveryStart = source.indexOf('function vNextAdminRecoverIncompleteBootstrap(');
  const recoveryEnd = source.indexOf('/** Re-run the idempotent Hub initialization', recoveryStart);
  const recovery = source.slice(recoveryStart, recoveryEnd);
  assert.ok(recoveryStart >= 0 &&
    recovery.includes("vNextAdminRequiredText_(req.hubSpreadsheetId, 'hubSpreadsheetId')") &&
    recovery.includes("vNextAdminRequiredText_(req.templateSpreadsheetId, 'templateSpreadsheetId')") &&
    recovery.includes('adminSourceScriptId !== ScriptApp.getScriptId()') &&
    recovery.includes('vNextAdminRegisterRelease_(') &&
    recovery.includes('vNextAdminWriteCanonicalReleasePair_('),
    'Interrupted bootstrap recovery must use explicit IDs, validate the central script, and resume immutable release/pointer commits');
  const sidebar = await readFile(path.join(root, 'VNext_AdminSidebar.html'), 'utf8');
  assert.ok(sidebar.includes('onclick="recoverBootstrap()"') &&
    sidebar.includes('vNextAdminRecoverIncompleteBootstrap({hubSpreadsheetId, templateSpreadsheetId})'),
    'Legacy Admin UI must expose the explicit-ID incomplete bootstrap recovery path');
  assert.ok(source.includes('function vNextAdminUpdateHubRuntimeFromSource('),
    'Generated Admin Hubs need a centrally managed runtime update path');
  assert.ok(source.includes('function vNextAdminProvisionPilotClientFromSource(request)') &&
    source.includes('return vNextAdminProvisionClientInHub_(hub, req);') &&
    source.includes('function vNextAdminInstallPilotAutomationFromSource(request)') &&
    source.includes("workerMode: 'CENTRAL_SOURCE_FALLBACK'"),
    'Pilot recovery must support central-source Client provisioning and scheduled processing without duplicating business logic');
  const provisioningStart = source.indexOf('function vNextAdminProvisionClientInHub_');
  const provisioningEnd = source.indexOf('/** Install the five-minute Pilot worker', provisioningStart);
  const provisioningFlow = source.slice(provisioningStart, provisioningEnd);
  assert.ok(provisioningFlow.includes("String(existing.status || '').toUpperCase() === 'PROVISIONING'") &&
    provisioningFlow.includes('vNextAdminResumeProvisioningClient_(') &&
    source.includes("vNextAdminWriteAudit_(hub, 'RESUME_PROVISION_CLIENT'"),
    'A verified PROVISIONING artifact must resume post-initialization phases without creating another Client');
  const stateRowsStart = source.indexOf('function vNextAdminValidateClientStateRows_');
  const stateRowsEnd = source.indexOf('function vNextAdminIsTrustedRejectedStateMarker_', stateRowsStart);
  assert.ok(source.slice(stateRowsStart, stateRowsEnd).includes('if (!rows || !rows.length) return [];'),
    'Empty Client STATE_EVENT validation must preserve the sourceRows array contract');
  const aclStart = source.indexOf('function vNextAdminAssertClientFileAcl_');
  const aclEnd = source.indexOf('function vNextAdminPreparePrivateBootstrapFolder_', aclStart);
  assert.ok(source.slice(aclStart, aclEnd).includes("const actual = vNextAdminMergeEmails_(owner, actualEditors, actualViewers);"),
    'Drive ownership must satisfy the expected editor ACL because getEditors excludes the owner');
  const manifestStart = source.indexOf('function vNextAdminTemplateSheetManifest_(');
  const manifestEnd = source.indexOf('function vNextAdminSerializeValidation_', manifestStart);
  const manifestFunction = source.slice(manifestStart, manifestEnd);
  assert.ok(manifestFunction.includes('const maxRows = sheet.getMaxRows();') &&
    manifestFunction.includes('const rows = fullGrid ? maxRows : Math.max(1, sheet.getLastRow());') &&
    manifestFunction.includes('usedRows: fullGrid ? undefined : rows') &&
    !manifestFunction.includes('const rows = sheet.getMaxRows();'),
    'Template UI manifests must hash the used envelope while retaining full grid dimensions to stay within GAS limits');
  assert.ok(source.includes("const VN_ADMIN_TEMPLATE_MANIFEST_SCHEMA = 'VNEXT_TEMPLATE_UI_V3'") &&
    source.includes("const VN_ADMIN_TEMPLATE_MANIFEST_SCHEMA_V2 = 'VNEXT_TEMPLATE_UI_V2'") &&
    source.includes('vNextAdminTemplateUiManifestHashV2_(template)'),
    'The faster manifest must be versioned while existing V2 releases remain verifiable');
  assert.ok(source.includes('AI_ZERO_REFERENCE_ONLY') &&
    source.includes("exception_type: 'AI_RESEARCH_UNAVAILABLE'"),
    'Terminal AI research failures must degrade explicitly rather than block the forecast');
  const aiSource = await readFile(path.join(root, 'VNext_AI.js'), 'utf8');
  assert.ok(aiSource.includes("String(req.internalOperation || '') === 'ADMIN_JOB' && req.spreadsheet"),
    'Central-source Pilot AI must accept only the trusted worker Hub handle');
  const executeStart = source.indexOf('function vNextAdminExecuteJob_');
  const executeEnd = source.indexOf('function vNextAdminFinishJob_', executeStart);
  const executeJob = source.slice(executeStart, executeEnd);
  assert.ok(executeJob.includes("internalOperation: 'ADMIN_JOB'") && executeJob.includes('spreadsheet: hub'),
    'Admin jobs must pass the explicit registered Hub to providers and Engine calls');
  assert.ok(executeJob.includes('vNextEngineBuildAdminRunIdentity_(') &&
    executeJob.includes('vNextEngineLookupRunForResume_(') &&
    executeJob.includes('vNextAdminEnsurePersistedForecastDraftState_(') &&
    executeJob.includes("'RUN_PERSISTED'") && executeJob.includes("'CLIENT_SYNCED'"),
    'Forecast jobs must use deterministic run identity and resume Client synchronization after durable SUCCESS');
  assert.ok(executeJob.includes('engineError.vNextRunIdentityFailure === true') &&
    executeJob.includes("exception_type: 'FORECAST_RUN_IDENTITY_CONFLICT'"),
    'Run identity conflicts must stop without the generic READY_TO_RUN recovery');
  const queueStart = source.indexOf('function vNextQueueClientForecastRequest(');
  const queueEnd = source.indexOf('/** Enqueue a forecast request', queueStart);
  assert.ok(source.slice(queueStart, queueEnd).includes("reason: 'forecast_requested:' + requestId"),
    'Client request state must carry the exact immutable request ID');
  const stateSemanticStart = source.indexOf('function vNextAdminValidateClientStateEventSemantics_');
  const stateSemanticEnd = source.indexOf('function vNextAdminClientInputReadiness_', stateSemanticStart);
  const stateSemantic = source.slice(stateSemanticStart, stateSemanticEnd);
  assert.ok(stateSemantic.includes('/^forecast_requested:(REQ-') &&
    stateSemantic.includes('requestReason[1]'),
    'Hub state validation must link READY_TO_RUN>RUNNING to the exact request ID');
  const persistedStateStart = source.indexOf('function vNextAdminEnsurePersistedForecastDraftState_');
  const persistedStateEnd = source.indexOf('function vNextAdminRunIdentityFailure_', persistedStateStart);
  const persistedState = source.slice(persistedStateStart, persistedStateEnd);
  assert.ok(persistedState.includes("state === 'RUNNING'") &&
    persistedState.includes("state === 'DRAFT_READY'") &&
    persistedState.includes("String(latest.related_run_id || '') !== runId"),
    'Durable SUCCESS resume must append or verify the exact Hub DRAFT_READY state event before Client sync');
  assert.ok(executeJob.indexOf('vNextEngineLookupRunForResume_(') <
    executeJob.lastIndexOf('vNextAdminAuthorizeAiRollbackJob_('),
    'AI rollback must inspect a persisted deterministic SUCCESS before applying the RUNNING-only authorization path');
  assert.ok(source.includes('function vNextAdminUpgradeEmptyPilotClient(request)') &&
    source.includes('function vNextAdminRecoverEmptyPilotClientUpgrade(request)') &&
    source.includes('const dryRun = req.dryRun !== false;'),
    'A same-URL empty-Pilot upgrade must expose read-only-first apply and durable recovery APIs');
  const emptyApplyStart = source.indexOf('function vNextAdminApplyEmptyPilotRelease_');
  const emptyApplyEnd = source.indexOf('function vNextAdminAppendEmptyPilotRepairMeta_', emptyApplyStart);
  const emptyApply = source.slice(emptyApplyStart, emptyApplyEnd);
  assert.ok(emptyApply.indexOf('vNextClientRuntimeCopyScriptContent_') <
    emptyApply.indexOf('vNextAdminCopyTemplateUiToClient_') &&
    emptyApply.indexOf('vNextAdminAppendEmptyPilotRepairMeta_') <
    emptyApply.indexOf('vNextAdminPatchRegistryByBookId_'),
    'Empty-Pilot registry identity must be the last cross-system release commit');
  assert.ok(sidebar.includes('id="emptyPilotBookSelect"') &&
    sidebar.includes('vNextAdminUpgradeEmptyPilotClient({ bookId, dryRun:true })') &&
    sidebar.includes('vNextAdminRecoverEmptyPilotClientUpgrade({'),
    'Admin Sidebar must select an existing Client without manual Book ID entry, dry-run it, and expose recovery');
}

async function checkAdminCoverageContracts() {
  const source = await readFile(path.join(root, 'VNext_Admin.js'), 'utf8');
  const sidebar = await readFile(path.join(root, 'VNext_AdminSidebar.html'), 'utf8');
  const sandbox = gasSandbox();
  vm.createContext(sandbox);
  vm.runInContext(await readFile(path.join(root, '0_VNext_Naming.js'), 'utf8'), sandbox, { filename: '0_VNext_Naming.js' });
  vm.runInContext(source, sandbox, { filename: 'VNext_Admin.js' });

  const payload = vm.runInContext(`({
    requestId:'REQ-1', bookId:'BOOK-1', clientId:'CLIENT-1', clientName:'Client',
    fiscalYear:2027, asOf:'2026-08-10', cutoff:'2026-07-31',
    bookConfiguredAsOf:'2026-08-01', requestedAt:'2026-08-10T12:00:00.000Z',
    requestedBy:'owner@example.com'
  })`, sandbox);
  const canonical = sandbox.vNextAdminCanonicalJson_(payload);
  assert.equal(sandbox.vNextAdminAssertClientRequestPayload_(payload, canonical, 'REQ-1'), true);
  for (const forbidden of [
    'seed', 'parameters', 'actualRecords', 'previousRunId', 'internalOperation',
    'persist', 'manageState', 'aiResearchJobId', 'trustedReuseSeedFromRunId',
    'trustedRollbackContext', 'trustedAllowedDelayedAiRequestIds', '__proto__'
  ]) {
    const mutated = vm.runInContext(`JSON.parse(${JSON.stringify(JSON.stringify(payload))})`, sandbox);
    Object.defineProperty(mutated, forbidden, { value: forbidden === '__proto__' ? {} : 'forbidden', enumerable: true, configurable: true });
    assert.throws(() => sandbox.vNextAdminAssertClientRequestPayload_(
      mutated, sandbox.vNextAdminCanonicalJson_(mutated), 'REQ-1'
    ), /keys are not exact/);
  }
  const missing = vm.runInContext(`JSON.parse(${JSON.stringify(JSON.stringify(payload))})`, sandbox);
  delete missing.cutoff;
  assert.throws(() => sandbox.vNextAdminAssertClientRequestPayload_(
    missing, sandbox.vNextAdminCanonicalJson_(missing), 'REQ-1'
  ), /keys are not exact/);
  assert.throws(() => sandbox.vNextAdminAssertClientRequestPayload_(payload, canonical, 'REQ-X'), /requestId/);
  assert.throws(() => sandbox.vNextAdminAssertClientRequestPayload_(payload, JSON.stringify(payload), 'REQ-1'), /canonical/);
  const originalReadTable = sandbox.vNextAdminReadTable_;
  const originalValidateClientRequestRow = sandbox.vNextAdminValidateClientRequestRow_;
  const requestRow = {
    request_event_id:'REQEV-1', request_id:'REQ-1', book_id:'BOOK-1',
    event_type:'REQUESTED', status:'PENDING', request_hash:sandbox.vNextAdminSha256_(canonical),
    request_json:canonical, requested_at:payload.requestedAt, requested_by:'owner@example.com'
  };
  const failedRow = {
    request_event_id:'REQEV-2', request_id:'REQ-1', book_id:'BOOK-1',
    event_type:'FAILED', status:'FAILED', request_hash:requestRow.request_hash,
    request_json:canonical, requested_at:payload.requestedAt, requested_by:'owner@example.com'
  };
  sandbox.vNextAdminReadTable_ = () => ({ rows: [requestRow, failedRow] });
  sandbox.vNextAdminValidateClientRequestRow_ = row => ({
    requestId: row.request_id, requestedAtMs: Date.parse(row.requested_at)
  });
  try {
    const stateRequest = sandbox.vNextAdminLatestValidPendingRequest_(
      { getSheetByName: () => ({}) },
      { book_id:'BOOK-1', client_id:'CLIENT-1', client_name:'Client', fiscal_year:2027 },
      'owner@example.com', Date.parse(payload.requestedAt), 'REQ-1'
    );
    assert.equal(stateRequest.requestId, 'REQ-1',
      'A later FAILED event must retain its immutable REQUESTED row for delayed state sync');
    sandbox.vNextAdminReadTable_ = () => ({ rows: [requestRow, {
      ...failedRow, request_event_id:'REQEV-3', event_type:'REJECTED', status:'REJECTED'
    }] });
    assert.equal(sandbox.vNextAdminLatestValidPendingRequest_(
      { getSheetByName: () => ({}) },
      { book_id:'BOOK-1', client_id:'CLIENT-1', client_name:'Client', fiscal_year:2027 },
      'owner@example.com', Date.parse(payload.requestedAt), 'REQ-1'
    ), null, 'A REJECTED request must never authorize a state transition');
  } finally {
    sandbox.vNextAdminReadTable_ = originalReadTable;
    sandbox.vNextAdminValidateClientRequestRow_ = originalValidateClientRequestRow;
  }

  assert.equal(sandbox.vNextAdminNormalizeAiRollbackScope_('all'), 'ALL');
  assert.equal(sandbox.vNextAdminNormalizeAiRollbackScope_('selected'), 'SELECTED');
  assert.throws(() => sandbox.vNextAdminNormalizeAiRollbackScope_('partial'), /ALL or SELECTED/);
  assert.equal(sandbox.vNextAdminReturnedPlanTargetState_('PLAN_ONLY'), 'CHANGES_REQUESTED');
  assert.equal(sandbox.vNextAdminReturnedPlanTargetState_('REOPEN_INPUT'), 'INPUT_OPEN');
  assert.equal(sandbox.vNextAdminReturnedPlanTargetState_('RERUN_SAME_INPUT'), 'READY_TO_RUN');
  assert.equal(sandbox.vNextAdminModelCheckPassed_('{"status":"PASS"}'), true);
  assert.equal(sandbox.vNextAdminModelCheckPassed_('{"status":"FAIL"}'), false);
  assert.equal(sandbox.vNextAdminValidateClientEvidenceAmount_({
    amount_mode:'EXACT', amount_low:100, amount_mid:100, amount_high:100, amount_band:''
  }), true);
  assert.equal(sandbox.vNextAdminValidateClientEvidenceAmount_({
    amount_mode:'BAND', amount_low:100, amount_mid:'', amount_high:200, amount_band:'SMALL'
  }), true);
  assert.throws(() => sandbox.vNextAdminValidateClientEvidenceAmount_({
    amount_mode:'EXACT', amount_mid:-1, amount_band:''
  }), /non-negative range/);
  assert.equal(sandbox.vNextAdminValidateEvidencePeriodInFiscalYear_(
    {target_start_month:'2027-04', target_end_month:'2028-03'}, 2027
  ), true);
  assert.throws(() => sandbox.vNextAdminValidateEvidencePeriodInFiscalYear_(
    {target_start_month:'2027-03', target_end_month:'2028-03'}, 2027
  ), /outside/);
  assert.equal(sandbox.vNextAdminAssertNoTrustedForecastPayload_({ requestId:'REQ-1' }), true);
  assert.throws(() => sandbox.vNextAdminAssertNoTrustedForecastPayload_({ trustedReuseSeedFromRunId:'RUN-1' }), /forbidden trusted fields/);
  const oldPair = {releaseId:'TPL-OLD', modelReleaseId:'MODEL-OLD', templateSpreadsheetId:'SHEET-OLD'};
  const newPair = {releaseId:'TPL-NEW', modelReleaseId:'MODEL-NEW', templateSpreadsheetId:'SHEET-NEW'};
  assert.equal(sandbox.vNextAdminClassifyActiveReleasePairCas_(newPair, newPair, oldPair, true), 'REUSE');
  assert.equal(sandbox.vNextAdminClassifyActiveReleasePairCas_(newPair, newPair, oldPair, false), 'REPAIR_CACHES');
  assert.equal(sandbox.vNextAdminClassifyActiveReleasePairCas_(oldPair, newPair, oldPair, false), 'CAS');
  assert.equal(sandbox.vNextAdminClassifyActiveReleasePairCas_(
    {releaseId:'THIRD', modelReleaseId:'MODEL-OLD', templateSpreadsheetId:'X'}, newPair, oldPair, false
  ), 'CONFLICT');
  const queueMetrics = sandbox.vNextAdminQueueAgeMetrics_([
    {status:'QUEUED', created_at:'2026-08-10T00:00:00.000Z'},
    {status:'RUNNING', created_at:'2026-08-10T00:10:00.000Z'}
  ], Date.parse('2026-08-10T00:20:00.000Z'));
  assert.equal(queueMetrics.queued, 1);
  assert.equal(queueMetrics.running, 1);
  assert.equal(queueMetrics.oldestQueuedAgeMinutes, 20);
  assert.equal(queueMetrics.staleQueued, 1);
  const safePortalRetry = {
    job_type:'PORTAL_PROVISION_CLIENT', status:'FAILED', attempts:1,
    error:'Requested release is not ACTIVE: vnext-pilot-20260812'
  };
  assert.equal(sandbox.vNextAdminIsKnownSafeRetryCandidate_(safePortalRetry), true);
  assert.equal(sandbox.vNextAdminIsKnownSafeRetryCandidate_({...safePortalRetry, attempts:3}), false,
    'Admin UX must never offer an exhausted Portal job as an automatic safe retry');
  assert.equal(sandbox.vNextAdminIsKnownSafeRetryCandidate_({
    job_type:'FORECAST_REQUEST', status:'FAILED', attempts:1,
    error:'At least 5 fiscal years of confirmed actual history are required; found 2.'
  }), false, 'Actual-data failures require a human decision and must not be generically retried');
  assert.equal(sandbox.vNextAdminIsActualDataIssue_({
    error:'At least 5 fiscal years of confirmed actual history are required; found 2.'
  }), true);
  const sidebarJobs = sandbox.vNextAdminJobsForSidebar_([
    {...safePortalRetry, job_id:'J1', created_at:'2026-08-12T00:00:00.000Z'},
    {job_type:'FORECAST_REQUEST',status:'RUNNING',job_id:'J2',created_at:'2026-08-12T00:01:00.000Z'}
  ]);
  assert.equal(sidebarJobs[0].status, 'FAILED');
  assert.equal(sidebarJobs[0].taskLabel, 'クライアント年度ブックの作成');
  assert.equal(sidebarJobs[0].safeRetryCandidate, true);
  const mismatchedException = sandbox.vNextAdminExceptionForSidebar_({
    exception_type:'JOB_FAILED', source_ref:'CURRENT-UNSAFE', book_id:'BOOK-1'
  }, {
    'BOOK-1':{spreadsheet_url:'https://docs.google.com/spreadsheets/d/BOOK_1/edit'}
  }, [
    {...safePortalRetry, job_id:'OLDER-SAFE', target_book_id:'BOOK-1'},
    {job_id:'CURRENT-UNSAFE', target_book_id:'BOOK-1', status:'FAILED', attempts:3,
      job_type:'PORTAL_PROVISION_CLIENT', error:'Unknown failure'}
  ]);
  assert.equal(mismatchedException.actionType, 'OPEN_BOOK',
    'A JOB_FAILED exception must never inherit the safe retry action from another failed job in the same book');
  assert.equal(sandbox.vNextAdminSidebarSpreadsheetUrl_(
    'https://docs.google.com/spreadsheets/d/ABC_123/edit'
  ), 'https://docs.google.com/spreadsheets/d/ABC_123/edit');
  assert.equal(sandbox.vNextAdminSidebarSpreadsheetUrl_('https://example.com/not-allowed'), '');
  const originalResolvePortal = sandbox.vNextAdminResolvePortal_;
  const originalPortalReadTable = sandbox.vNextAdminReadTable_;
  sandbox.vNextAdminResolvePortal_ = () => ({
    spreadsheet:{getUrl:() => 'https://docs.google.com/spreadsheets/d/PORTAL_1/edit'}
  });
  sandbox.vNextAdminReadTable_ = () => ({rows:[
    {request_id:'REQ-1',event_type:'REQUESTED',status:'PENDING',client_name:'Client',fiscal_year:2027},
    {request_id:'REQ-1',event_type:'REQUESTED',status:'COMPLETED',client_name:'Tampered',fiscal_year:2099},
    {request_id:'REQ-2',event_type:'CREATION_STARTED',status:'CREATING',client_name:'Client 2',fiscal_year:2027}
  ]});
  try {
    const portalProjection = sandbox.vNextAdminPortalRequestsForSidebar_({});
    assert.equal(portalProjection.counts.waiting, 1);
    assert.equal(portalProjection.counts.processing, 1);
    assert.equal(portalProjection.counts.completed, 0,
      'Portal status projection must ignore a status that is inconsistent with its append-only event type');
    assert.equal(portalProjection.attention[1].clientName, 'Client');
  } finally {
    sandbox.vNextAdminResolvePortal_ = originalResolvePortal;
    sandbox.vNextAdminReadTable_ = originalPortalReadTable;
  }
  assert.equal(sandbox.vNextAdminAttentionSummary_({
    automationInstalled:true,
    counts:{exceptions:0,pendingApprovals:0,portalAttention:0,queuedJobs:1,runningJobs:0},
    operations:{schedulerStale:false}, portalRequests:{counts:{waiting:0,processing:0}}
  }, 0).status, 'PROCESSING');
  assert.equal(sandbox.vNextAdminAttentionSummary_({
    automationInstalled:true,
    counts:{exceptions:0,pendingApprovals:1,portalAttention:0,queuedJobs:0,runningJobs:0},
    operations:{schedulerStale:false}, portalRequests:{counts:{waiting:0,processing:0}}
  }, 0).status, 'ATTENTION');
  assert.equal(sandbox.vNextAdminAttentionSummary_({
    automationInstalled:true,
    counts:{exceptions:0,pendingApprovals:0,portalAttention:0,queuedJobs:1,runningJobs:0},
    operations:{schedulerStale:false,queueStale:true}, portalRequests:{counts:{waiting:0,processing:0}}
  }, 0).status, 'ERROR', 'A stale queue must never look like normal in-progress work');
  assert.equal(sandbox.vNextAdminAttentionSummary_({
    automationInstalled:true,
    counts:{exceptions:0,pendingApprovals:0,portalAttention:0,queuedJobs:0,runningJobs:0},
    operations:{schedulerStale:false,queueStale:false}, portalRequests:{unavailable:true,counts:{}}
  }, 0).status, 'ATTENTION', 'A configured but unreadable Portal must be visible to the Admin');
  assert.ok(source.includes('function vNextAdminRunOperationalCycle()') &&
    source.includes("vNextAdminWithScriptLock_('admin-run-operational-cycle'") &&
    source.includes('vNextAdminRequeueKnownPilotFailures_(hub)') &&
    source.includes("vNextAdminWriteAudit_(hub, 'RUN_OPERATIONAL_CYCLE'"),
    'The Admin one-click cycle must use durable jobs, narrow retries, a lock, and append-only audit');
  assert.ok(sidebar.includes('<html lang="ja">') && sidebar.includes('id="attentionOverview"') &&
    sidebar.includes('id="portalRequests"') && sidebar.includes('id="recentJobs"') &&
    sidebar.includes('vNextAdminRunOperationalCycle') && sidebar.includes('window.confirm(confirmation)'),
    'The normal Admin view must prioritize decisions, Portal progress, safe processing, and approval confirmation');
  assert.equal(/<div class="metric">待機job/.test(sidebar), false,
    'The primary Admin metrics must not expose internal job terminology');
  assert.equal(sandbox.vNextAdminPortalCanonicalClientKey_({clientName:'株式会社 テスト'}),
    sandbox.vNextAdminPortalCanonicalClientKey_({clientName:'テスト'}));
  assert.notEqual(sandbox.vNextAdminPortalCanonicalClientKey_({client_id:'CLIENT-A', client_name:'A'}),
    sandbox.vNextAdminPortalCanonicalClientKey_({client_id:'CLIENT-B', client_name:'B'}),
    'BOOK_REGISTRY snake_case rows must produce distinct Portal directory keys');
  assert.equal(sandbox.vNextAdminNormalizeDomain_('@Example.COM'), 'example.com');
  assert.equal(sandbox.vNextAdminSafeCatalogText_('ｱｽﾄﾗｾﾞﾈｶ(株)', 120, 'clientName'), 'ｱｽﾄﾗｾﾞﾈｶ(株)',
    'The ZAC display/client name must preserve the exact AO source text used by the forecast engine');
  const fakeActualSheet = {
    getName: () => '*2026_actual_value', getLastRow: () => 3, getMaxColumns: () => 66,
    getRange(row, column, rowCount, columnCount) {
      if (row === 1 && column === 40) return {getValue: () => 'クライアントコード'};
      if (row === 1 && column === 41) return {getValue: () => 'クライアント名'};
      if (row === 2 && column === 40 && rowCount === 2 && columnCount === 2) {
        return {getDisplayValues: () => [['AZ-001', 'ｱｽﾄﾗｾﾞﾈｶ(株)'], ['', '仮登録']]};
      }
      throw new Error(`unexpected actual range ${row},${column},${rowCount},${columnCount}`);
    }
  };
  const fakeDefinitionSheet = {
    getName: () => '*defclients', getLastRow: () => 7,
    getRange: (row,column,rowCount,columnCount) => {
      assert.deepEqual([row,column,rowCount,columnCount],[1,1,7,2]);
      return {getDisplayValues: () => [
        ['年','2026'],['カテゴリ','-'],['クライアント','売上'],['全体','100'],
        ['ｱｽﾄﾗｾﾞﾈｶ(株)','80'],['仮登録','10'],['新規製薬(株)','10']
      ]};
    }
  };
  const extractedCatalog = sandbox.vNextAdminExtractZacClientCatalog_({
    getSheets: () => [fakeActualSheet, fakeDefinitionSheet]
  });
  const astraCatalog = extractedCatalog.clients.find(item => item.clientCode === 'AZ-001');
  assert.equal(astraCatalog.clientName, 'ｱｽﾄﾗｾﾞﾈｶ(株)',
    'The AstraZeneca catalog selection must retain the exact half-width ZAC name for history matching');
  assert.equal(extractedCatalog.clients.some(item => item.clientName === '新規製薬(株)'), true,
    '*defclients-only clients must also be available in the employee picker');
  ['年','カテゴリ','クライアント','全体','仮登録'].forEach(label => {
    assert.equal(extractedCatalog.clients.some(item => item.clientName === label), false,
      `*defclients report label ${label} must never become an employee client option`);
  });

  const catalogHeaders = vm.runInContext('VN_ADMIN_ZAC_CLIENT_CATALOG_HEADERS.slice()', sandbox);
  const existingCatalogRows = [
    {catalog_key:'A',client_id:'A',client_name:'旧A',normalized_name:'a',is_active:1,
      first_seen_at:'2025-01-01',catalog_version:'OLD'},
    {catalog_key:'B',client_id:'B',client_name:'旧B',normalized_name:'b',is_active:1,
      first_seen_at:'2025-01-02',catalog_version:'OLD'},
    {catalog_key:'C',client_id:'C',client_name:'旧C',normalized_name:'c',is_active:0,
      first_seen_at:'2025-01-03',catalog_version:'OLD'}
  ];
  const nextCatalogBody = sandbox.vNextAdminBuildZacCatalogBody_(existingCatalogRows, [
    {catalogKey:'A',clientId:'A',clientCode:'1',clientName:'新A',normalizedName:'a',sourceYears:[2026]},
    {catalogKey:'C',clientId:'C',clientCode:'3',clientName:'新C',normalizedName:'c',sourceYears:[2026]},
    {catalogKey:'D',clientId:'D',clientCode:'4',clientName:'新D',normalizedName:'d',sourceYears:[2026]}
  ], {catalogVersion:'NEW',refreshedAt:'2026-08-12T13:00:00.000Z',sourceSpreadsheetId:'SOURCE'});
  const nextCatalogRows = nextCatalogBody.map(values => Object.fromEntries(
    catalogHeaders.map((header,index) => [header,values[index]])));
  assert.deepEqual(nextCatalogRows.map(row => row.catalog_key), ['A','B','C','D'],
    'Bulk catalog merge must preserve existing row order and append only new clients');
  assert.equal(nextCatalogRows[0].first_seen_at, '2025-01-01');
  assert.equal(nextCatalogRows[1].is_active, 0, 'Missing previously-active client must be deactivated');
  assert.equal(nextCatalogRows[1].catalog_version, 'NEW');
  assert.equal(nextCatalogRows[2].is_active, 1, 'An inactive client seen again must reactivate');
  assert.equal(nextCatalogRows[2].first_seen_at, '2025-01-03');
  assert.equal(nextCatalogRows[3].first_seen_at, '2026-08-12T13:00:00.000Z');
  assert.throws(() => sandbox.vNextAdminBuildZacCatalogBody_([
    {catalog_key:'A'},{catalog_key:'A'}
  ], [], {catalogVersion:'NEW',refreshedAt:'2026-08-12T13:00:00.000Z',sourceSpreadsheetId:'SOURCE'}),
  /重複catalog_key/, 'Bulk catalog merge must fail closed on duplicate persisted identities');

  const originalPortalConfigRead = sandbox.vNextAdminReadKeyValueSheet_;
  const originalPortalAppend = sandbox.vNextAdminAppendObject_;
  const originalPortalActor = sandbox.vNextAdminActor_;
  let legacyStatusRecord = null;
  sandbox.vNextAdminReadKeyValueSheet_ = () => ({runtime_version:'vnext-portal-1.1.0'});
  sandbox.vNextAdminAppendObject_ = (_portal, _sheet, record) => { legacyStatusRecord = record; return record; };
  sandbox.vNextAdminActor_ = () => 'admin@example.com';
  try {
    const legacyPayload = {
      clientId:'', clientName:'Legacy Client', fiscalYear:2027,
      forecastOwnerEmail:'owner@example.com', relatedMemberEmails:[],
      requestId:'PORTAL-REQ-LEGACY01', requestType:'CREATE_CLIENT_FY_BOOK',
      requestedAt:'2026-08-12T00:00:00.000Z', requestedBy:'owner@example.com',
      schemaVersion:'vnext-portal-request-1'
    };
    sandbox.vNextAdminAppendPortalRequestEvent_({}, {
      payload:legacyPayload, requestHash:'a'.repeat(64), clientId:'C-DERIVED',
      forecastOwnerEmail:'owner@example.com', relatedMemberEmails:[], relatedMemberNames:[]
    }, 'VALIDATION_STARTED', 'VALIDATING', {});
    assert.equal(legacyStatusRecord.client_id, '',
      'Legacy v1 status rows must preserve an originally blank clientId projection');
    assert.equal(legacyStatusRecord.catalog_key, '');
    assert.equal(legacyStatusRecord.related_member_names_json, '',
      'Legacy v1 status rows must leave v2-only projection columns blank');
  } finally {
    sandbox.vNextAdminReadKeyValueSheet_ = originalPortalConfigRead;
    sandbox.vNextAdminAppendObject_ = originalPortalAppend;
    sandbox.vNextAdminActor_ = originalPortalActor;
  }

  const portalMigrationStart = source.indexOf('function vNextAdminUpdateSharedPortalRuntime(');
  const portalMigrationEnd = source.indexOf('function vNextAdminWritePortalConfigValues_', portalMigrationStart);
  const portalMigration = source.slice(portalMigrationStart, portalMigrationEnd);
  assert.ok(portalMigration.indexOf('tablesExpanded = true') <
    portalMigration.indexOf('vNextAdminExpandPortalTableHeadersForV2_('),
  'Portal migration must arm the v1 header rollback before its first multi-call table mutation');
  assert.ok(portalMigration.indexOf('contentUpdateAttempted = true') <
    portalMigration.indexOf('vNextClientRuntimePutContent_(portal.scriptId, target)'),
  'Portal migration must arm runtime rollback before the remote content PUT can partially succeed');
  assert.ok(portalMigration.includes('if (contentUpdateAttempted) {'),
    'Portal migration catch path must restore the previous runtime after every attempted PUT');
  assert.ok(portalMigration.includes('const needsV2HeaderExpansion = !vNextAdminPortalUsesV2Tables_') &&
    portalMigration.includes('if (needsV2HeaderExpansion) {'),
  'Portal v1.1 to v1.2 must preserve the already-expanded v2 table headers');
  assert.ok(portalMigration.includes('[VN_ADMIN_PORTAL_RUNTIME_VERSION].concat(VN_ADMIN_PORTAL_LEGACY_RUNTIME_VERSIONS)') &&
    portalMigration.includes('.indexOf(portal.runtimeVersion) < 0'),
  'Portal migration must allow a verified same-version runtime SHA update');
  assert.ok(portalMigration.includes('vNextPortalRuntimeValidateExistingFiles_(currentContent.files'),
    'Portal migration must accept the pre-web-entry four-file runtime when verifying the current pin');
  assert.ok(portalMigration.includes('vNextPortalRuntimeValidateExistingFiles_(restored.files'),
    'Portal migration rollback must re-verify the previous four-file or five-file runtime');
  assert.equal(portalMigration.includes('vNextPortalRuntimeValidateFiles_(currentContent.files'), false,
    'Current Portal content must not be forced through the latest five-file allowlist');
  assert.ok(portalMigration.includes('vNextAdminPublishPortalWebApp_(portal.scriptId, expectedWebAppUrl)'),
    'Portal update must republish the existing employee web app after a verified runtime copy');
  assert.ok(portalMigration.includes("'/deployments/' +") && portalMigration.includes("'put'"),
    'Portal web entry must update the existing deployment instead of creating a new /exec URL');
  assert.equal(portalMigration.includes("'/deployments', 'post'"), false,
    'Portal web entry republish must not create a second Web App URL');
  assert.ok(portalMigration.lastIndexOf('vNextAdminPublishPortalWebApp_(portal.scriptId, expectedWebAppUrl)') >
    portalMigration.lastIndexOf('throw migrationError'),
    'Web app republish must run after the runtime rollback window so a failed /exec pin cannot undo files');
  assert.ok(portalMigration.indexOf('vNextClientRuntimeGetContent_(portal.scriptId)') <
    portalMigration.indexOf('if (currentSha === targetSha &&'),
    'Live Portal files must be hashed before Hub pin reuse can skip the copy');
  assert.equal(portalMigration.includes('does not match its stored SHA-256 pin.'), false,
    'A drifted live Portal must be overwritten with the verified target instead of aborting');
  assert.ok(portalMigration.includes('vNextAdminEmployeePortalWebAppUrl_(hub)'),
    'Portal republish must prefer the bookmarked employee /exec stored on the Hub');
  assert.ok(source.includes('AKfycbxVtnFiXMB6FwKRdMj_PJVmq4zlpYMoBLS3zXy_1ruTGqyTSPxyepkJegcL9rGiUbwH'),
    'Employee /exec republish must pin the bookmarked deployment, not HEAD');
  const bookmarkId = 'AKfycbxVtnFiXMB6FwKRdMj_PJVmq4zlpYMoBLS3zXy_1ruTGqyTSPxyepkJegcL9rGiUbwH';
  const webAppEntry = (id, versionNumber) => ({
    deploymentId: id,
    deploymentConfig: versionNumber ? { versionNumber: versionNumber } : {},
    entryPoints: [{
      entryPointType: 'WEB_APP',
      webApp: { url: 'https://script.google.com/a/macros/example.com/s/' + id + '/exec' }
    }]
  });
  assert.equal(sandbox.vNextAdminWebAppDeploymentIdFromUrl_(
    'https://script.google.com/a/macros/bigm2y.com/s/' + bookmarkId + '/exec'
  ), bookmarkId);
  assert.equal(sandbox.vNextAdminSelectPortalWebAppDeployment_([
    webAppEntry('AKfycbwHEADONLYDEPLOYMENTID0000000000000000000000'),
    webAppEntry(bookmarkId, 8)
  ]).deploymentId, bookmarkId, 'Existing employee /exec must win over HEAD');
  assert.throws(() => sandbox.vNextAdminSelectPortalWebAppDeployment_([
    webAppEntry('AKfycbwHEADONLYDEPLOYMENTID0000000000000000000000', 3)
  ], 'https://script.google.com/macros/s/' + bookmarkId + '/exec'), /Bookmarked employee \/exec/);
  const adminManifest = JSON.parse(await readFile(path.join(root, 'appsscript.json'), 'utf8'));
  assert.ok(adminManifest.oauthScopes.includes('https://www.googleapis.com/auth/script.deployments'),
    'Admin project must request deployment scope to pin the employee /exec version');
  assert.ok(adminManifest.oauthScopes.includes('https://www.googleapis.com/auth/script.webapp.deploy'),
    'Admin project must request webapp.deploy to update the employee Web App');

  const originalAccessible = sandbox.vNextAdminSpreadsheetAccessible_;
  const originalSharingAssert = sandbox.vNextAdminAssertEmployeeFileSharing_;
  sandbox.DriveApp = { getFileById: id => ({ id }) };
  sandbox.vNextAdminSpreadsheetAccessible_ = () => true;
  sandbox.vNextAdminAssertEmployeeFileSharing_ = () => true;
  try {
    assert.equal(sandbox.vNextAdminPortalExistingBookAccess_(
      {employeeDomain:'example.com'},
      {spreadsheet_id:'SHEET-1', access_policy:'PRIVATE', internal_domain:'example.com'}
    ).reusable, false, 'A private existing Client must not be returned as a completed Portal URL');
    assert.equal(sandbox.vNextAdminPortalExistingBookAccess_(
      {employeeDomain:'example.com'},
      {spreadsheet_id:'SHEET-1', access_policy:'INTERNAL_OPEN', internal_domain:'example.com'}
    ).reusable, true, 'An already verified employee-open Client may be reused');
  } finally {
    sandbox.vNextAdminSpreadsheetAccessible_ = originalAccessible;
    sandbox.vNextAdminAssertEmployeeFileSharing_ = originalSharingAssert;
  }

  const originalPortalHarvest = sandbox.vNextAdminHarvestPortalRequests_;
  const originalReadTableForPortal = sandbox.vNextAdminReadTable_;
  const originalAppendExceptionForPortal = sandbox.vNextAdminAppendException_;
  let isolatedPortalExceptions = 0;
  sandbox.vNextAdminHarvestPortalRequests_ = () => { throw new Error('portal-offline'); };
  sandbox.vNextAdminReadTable_ = () => ({rows:[]});
  sandbox.vNextAdminAppendException_ = () => { isolatedPortalExceptions++; };
  try {
    const isolated = sandbox.vNextAdminHarvestPortalRequestsSafely_({});
    assert.equal(isolated.queued, 0);
    assert.match(isolated.isolatedError, /portal-offline/);
    assert.equal(isolatedPortalExceptions, 1,
      'A Portal harvest failure must be isolated and surfaced without throwing out of the scheduler');
  } finally {
    sandbox.vNextAdminHarvestPortalRequests_ = originalPortalHarvest;
    sandbox.vNextAdminReadTable_ = originalReadTableForPortal;
    sandbox.vNextAdminAppendException_ = originalAppendExceptionForPortal;
  }

  const originalReadTableForLease = sandbox.vNextAdminReadTable_;
  const originalUpdateForLease = sandbox.vNextAdminUpdateTableRow_;
  const originalMarkPortalFailed = sandbox.vNextAdminMarkPortalJobFailed_;
  const originalAppendExceptionForLease = sandbox.vNextAdminAppendException_;
  let exhaustedPortalFailures = 0;
  sandbox.vNextAdminReadTable_ = (_hub, name) => ({rows: name === 'JOB_QUEUE' ? [{
    _rowNumber:2, job_id:'JOB-PORTAL', job_type:'PORTAL_PROVISION_CLIENT', target_book_id:'REQ-1',
    status:'RUNNING', attempts:3, locked_at:'2000-01-01T00:00:00.000Z', request_json:'{}'
  }] : []});
  sandbox.vNextAdminUpdateTableRow_ = () => {};
  sandbox.vNextAdminMarkPortalJobFailed_ = () => { exhaustedPortalFailures++; return true; };
  sandbox.vNextAdminAppendException_ = () => {};
  try {
    const recovery = sandbox.vNextAdminRecoverStaleLeases_({}, 20);
    assert.equal(recovery.failedJobs, 1);
    assert.equal(exhaustedPortalFailures, 1,
      'An exhausted Portal lease must append a terminal Portal failure event');
  } finally {
    sandbox.vNextAdminReadTable_ = originalReadTableForLease;
    sandbox.vNextAdminUpdateTableRow_ = originalUpdateForLease;
    sandbox.vNextAdminMarkPortalJobFailed_ = originalMarkPortalFailed;
    sandbox.vNextAdminAppendException_ = originalAppendExceptionForLease;
  }

  for (const contract of [
    'function vNextAdminRollbackAiEvidence(', "case 'AI_ROLLBACK_FORECAST'",
    "evidence_type: 'AI_ROLLBACK'", "response_type: 'NO_CHANGE'",
    'trustedReuseSeedFromRunId', 'trustedRollbackContext', 'trustedAllowedDelayedAiRequestIds',
    'function vNextAdminRegisterModelRelease(', 'function vNextAdminActivateModelRelease(',
    'function vNextAdminRollbackModelRelease(', 'active_model_release_id',
    "event_type: 'INPUT_REOPENED'", 'function vNextAdminRouteReturnedPlan(',
    'effectiveEvidenceIds', 'nonAiComparableHash',
    'function vNextAdminUpdateHubRuntimeFromSource(',
    'function vNextAdminCreateTemplateDraft(', 'function vNextAdminActivateReleasePair(',
    "TEMPLATE_JOURNAL: 'TEMPLATE_RELEASE_JOURNAL'", 'VNEXT_TEMPLATE_UI_V2',
    'function vNextAdminCopyAndVerifyTemplateUi_(', 'function vNextAdminAssertPrivateAdminFile_(',
    'function vNextAdminApprovePilotCanary(', 'function vNextAdminPrepareClientDestinationFolder_('
    , 'function vNextAdminProvisionSharedPortal(', 'function vNextAdminHarvestPortalRequests_('
    , "case 'PORTAL_PROVISION_CLIENT'", 'function vNextAdminRefreshPortalDirectory_('
    , 'function vNextAdminApplyEmployeeFileSharing_(', 'function vNextAdminAssertEmployeeFileSharing_('
    ,     'function vNextAdminResetGeneratedClientsForFreshUat(',
    'function vNextAdminResetGeneratedClientsForFreshUatFromSource(',
    'function vNextAdminRelocateLibraryToSharedDrive(',
    'function vNextAdminRelocateLibraryToSharedDriveFromSource(',
    "VN_ADMIN_FRESH_UAT_RESET_CONFIRMATION = 'RESET_GENERATED_CLIENTS'"
  ]) assert.ok(source.includes(contract), `missing Admin coverage contract: ${contract}`);
  const resetStart = source.indexOf('function vNextAdminResetGeneratedClientsInHub_');
  const resetEnd = source.indexOf('function vNextAdminProtectedSpreadsheetIds_', resetStart);
  const reset = source.slice(resetStart, resetEnd);
  assert.ok(resetStart >= 0 && resetEnd > resetStart, 'Fresh UAT reset must be a dedicated Hub operation');
  assert.ok(reset.includes("req.apply === true") && reset.includes('.setTrashed(true)'),
    'Fresh UAT reset must require apply=true before trashing generated Client files');
  assert.ok(reset.includes("status: 'ARCHIVED'") && reset.includes('vNextAdminRefreshPortalDirectory_') &&
    reset.includes('vNextAdminRefreshPortalEmployeeViews_'),
    'Fresh UAT reset must archive Client registry rows and rebuild Portal views');
  assert.ok(reset.includes('MODEL_RELEASE') && reset.includes("sheetName === 'MODEL_RELEASE'"),
    'Fresh UAT reset must preserve MODEL_RELEASE while clearing client audit rows');
  assert.ok(reset.includes('VN_ADMIN_SHEETS.AUDIT') && reset.includes('vNextAdminTrashAuditArtifacts_'),
    'Fresh UAT reset must clear ADMIN_AUDIT_LOG and trash Drive audit files');
  assert.equal(source.replace(reset, '').includes('.setTrashed('), false,
    'Drive trash must stay inside the confirmed fresh UAT reset path');
  assert.ok(sidebar.includes('vNextAdminResetGeneratedClientsForFreshUat') &&
    sidebar.includes('RESET_GENERATED_CLIENTS') &&
    sidebar.includes('apply:true') &&
    sidebar.includes('管理ハブ監査ログ'),
    'Admin Sidebar must hide the reset behind the exact confirmation phrase');
  assert.ok(sidebar.includes('現場の年度・クライアント指定は、申請入口（第2層）から行います') &&
    sidebar.includes('申請入口') &&
    sidebar.includes('申請を今すぐ処理'),
    'Admin Sidebar must keep daily processing in-panel and point field work to the Portal');
  const provisionStart = source.indexOf('function vNextAdminProvisionClientInHub_');
  const provisionEnd = source.indexOf('function vNextAdminResumeProvisioningClient_', provisionStart);
  const provision = source.slice(provisionStart, provisionEnd);
  assert.ok(provision.includes('vNextAdminResolveRelease_(hub, req.releaseId)') &&
    !provision.includes('req.releaseId || runtime.VNEXT_ACTIVE_RELEASE_ID'),
    'Unpinned Portal provisioning must resolve the canonical active pair instead of a stale Script Property cache');
  const duplicateStart = provision.indexOf('const existing = vNextAdminFindRegistryRow_');
  const duplicateEnd = provision.indexOf('if (existing)', duplicateStart);
  assert.equal(provision.slice(duplicateStart, duplicateEnd).includes('template_release_id'), false,
    'Client/FY duplicate identity must not include Template release');
  assert.ok(source.includes('portalRequests: vNextAdminHarvestPortalRequestsSafely_(hub)') &&
    source.includes("jobType: 'PORTAL_PROVISION_CLIENT'") &&
    source.includes("'PORTAL_PROVISION|' + vNextAdminPortalCanonicalClientKey_"),
    'Portal request harvest must enqueue idempotent async provisioning during scheduled maintenance');
  const portalDirectoryStart = source.indexOf('function vNextAdminRefreshPortalDirectory_');
  const portalDirectoryEnd = source.indexOf('function vNextAdminPortalNextAction_', portalDirectoryStart);
  assert.ok(source.slice(portalDirectoryStart, portalDirectoryEnd)
    .includes("String(row.status || '').toUpperCase() === 'ACTIVE'"),
    'Portal directory must expose only ACTIVE Client books');
  assert.ok(source.includes('function vNextAdminPortalExistingBookAccess_(') &&
    source.includes("detailCode: String(extra.code || 'EXISTING_BOOK_ADMIN_ACCESS_REQUIRED')"),
    'Private or mismatched existing duplicates must require Admin action instead of returning a completed URL');
  assert.ok(source.includes('function vNextAdminUpdateSharedPortalRuntimeForManualTest()') &&
    source.includes('function vNextAdminRecoverPortalProvisionForManualTest()'),
    'The Pilot must retain non-UI editor fallbacks for Portal runtime update and verified job recovery');
  assert.ok(source.includes("String(job.job_type || '') === 'PORTAL_PROVISION_CLIENT'") &&
    source.includes("vNextAdminMarkPortalJobFailed_(hub, job, 'Lease expired after 3 attempts.')"),
    'Exhausted Portal jobs must publish a terminal FAILED event');
  assert.ok(source.includes("targetMode: 'CLIENT', accessPolicy: accessPolicy") &&
    source.includes("targetMode: 'PORTAL', accessPolicy: 'INTERNAL_OPEN'"),
    'Domain sharing must be explicitly limited to employee-facing Client and Portal files');
  assert.ok(source.includes("['FORECAST_REQUEST', 'AI_ROLLBACK_FORECAST']"),
    'Lease exhaustion recovery must include AI rollback forecast jobs');
  const harvestStart = source.indexOf('function vNextAdminHarvestClientRequests_');
  const requestValidation = source.indexOf('vNextAdminValidateClientRequestRow_(row, registry, ownerEmails[0])', harvestStart);
  const forecastEnqueue = source.indexOf("jobType: 'FORECAST_REQUEST'", harvestStart);
  assert.ok(requestValidation >= harvestStart && forecastEnqueue > requestValidation,
    'Client request exact-key validation must run before forecast enqueue');
  const rowValidatorStart = source.indexOf('function vNextAdminValidateClientRequestRow_');
  const rowValidatorEnd = source.indexOf('\nfunction ', rowValidatorStart + 1);
  const rowValidator = source.slice(rowValidatorStart, rowValidatorEnd);
  assert.ok(rowValidator.includes('vNextAdminAssertClientRequestPayload_('),
    'Client request row validation must enforce the exact canonical payload contract');
  assert.ok(rowValidator.includes("event_type || '').toUpperCase() !== 'REQUESTED'") &&
    rowValidator.includes("status || '').toUpperCase() !== 'PENDING'"),
    'Client request row validation must accept only the immutable REQUESTED/PENDING event');
  const rejectStart = source.indexOf('function vNextAdminRejectClientRequest_');
  const rejectEnd = source.indexOf('function vNextAdminAppendClientRequestEvent_', rejectStart);
  const reject = source.slice(rejectStart, rejectEnd);
  assert.ok(reject.includes("String(event.reason || '') !== 'forecast_requested:' + requestId") &&
    reject.includes("'CLIENT_REQUEST|' + requestId + '|' + String(row.request_hash || '')") &&
    reject.includes("String(jobPayload.requestId || '') === requestId"),
    'Rejected-request recovery must consume only the exact state and active job lineage for that request ID/hash');
  assert.ok(source.includes('vNextAdminEnsureInitialModelRelease_(hub') &&
    source.includes("model_release_id: vNextAdminRequiredText_(opt.modelReleaseId"),
    'Bootstrap and Client BOOK_META must bind a distinct MODEL_RELEASE');
  const rollbackApiStart = source.indexOf('function vNextAdminRollbackAiEvidence(');
  const rollbackApiEnd = source.indexOf('function vNextAdminEnqueueAiResearch(', rollbackApiStart);
  const rollbackApi = source.slice(rollbackApiStart, rollbackApiEnd);
  assert.equal(/trustedReuseSeedFromRunId\s*:/.test(rollbackApi), false,
    'Public rollback API must not persist a pre-trusted seed field in the job payload');
  assert.ok(sidebar.includes('vNextAdminRollbackAiEvidence') && sidebar.includes('vNextAdminActivateModelRelease') &&
    sidebar.includes('vNextAdminRouteReturnedPlan') && sidebar.includes('vNextAdminUpdateHubRuntimeFromSource') &&
    sidebar.includes('vNextAdminCreateTemplateDraft') && sidebar.includes('vNextAdminActivateReleasePair') &&
    sidebar.includes('管理ハブ担当者 attestation'),
    'Admin Sidebar must expose the guarded coverage controls');

  const pairStart = source.indexOf('function vNextAdminActivateReleasePairInternal_');
  const pairEnd = source.indexOf('function vNextAdminAppendTemplateJournal_', pairStart);
  const pair = source.slice(pairStart, pairEnd);
  const newActive = pair.indexOf("phase: 'NEW_PAIR_ACTIVE'");
  const cas = pair.indexOf("phase: 'POINTER_CAS_COMMITTED'");
  const cache = pair.indexOf("phase: 'PROPERTY_CACHE_UPDATED'");
  const retired = pair.indexOf("phase: 'PREVIOUS_TEMPLATE_RETIRED'");
  assert.ok(newActive >= 0 && cas > newActive && cache > cas && retired > cache,
    'Template pair publish must journal new ACTIVE -> CAS pointer -> property cache -> old RETIRED');
  assert.ok(pair.includes('vNextAdminAssertModelTemplateCompatibility_(hub, modelSource, release)') &&
    source.includes("String(model.template_version || '') !== String(template.release_id || '')"),
    'MODEL_RELEASE.template_version must exactly match the paired Template release_id');
  const publishStart = source.indexOf('function vNextAdminPublishTemplateRelease(');
  const publishEnd = source.indexOf('/** Activate a STAGED Template', publishStart);
  const publish = source.slice(publishStart, publishEnd);
  assert.ok(publish.includes("status: 'STAGED'") &&
    publish.includes('vNextAdminResolveTemplateUiSource_(hub, req.templateDraftSpreadsheetId') &&
    publish.includes('vNextAdminCopyAndVerifyTemplateUi_('),
    'Publishing must inherit ACTIVE/draft UI into a clean STAGED template');
  assert.ok(source.includes('target.moveActiveSheet(index + 1)') && !source.includes('.setIndex('),
    'Template sheet ordering must use the supported Spreadsheet.moveActiveSheet API');
  assert.equal(publish.indexOf("status: 'RETIRED'"), -1,
    'Staging must not retire the current ACTIVE Template');
  for (const manifestContract of [
    'getNotes()', 'getRichTextValues()', 'getDataValidations()', 'getConditionalFormatRules()',
    'getMergedRanges()', 'getNumberFormats()', 'getBackgrounds()', 'getFontFamilies()',
    'Named ranges are forbidden', 'forbidden external-data formula'
  ]) assert.ok(source.includes(manifestContract), `missing Template manifest/forbidden contract: ${manifestContract}`);
  assert.ok(source.includes("String(pinnedModel.template_version || '') !== String(expectedRelease.release_id || '')") &&
    source.includes("String(row.template_version || '') === String(registry.template_release_id || '')"),
    'Provision/health metadata must preserve the exact Template/Model pair');
  assert.ok(source.includes("setting_key: 'ACTIVE_RELEASE_PAIR_JSON'") &&
    source.includes('Canonical one-cell active Template/Model pair') &&
    source.includes('vNextAdminActiveReleasePairMirrorsExact_'),
    'A canonical one-cell release pair must govern all pointer mirrors');
  const pointerStart = source.indexOf('function vNextAdminWriteActiveReleasePairPointers_');
  const pointerEnd = source.indexOf('function vNextAdminCacheActiveReleasePair_', pointerStart);
  const pointerWriter = source.slice(pointerStart, pointerEnd);
  assert.ok(pointerWriter.includes("action === 'REPAIR_CACHES'") ||
    (pointerWriter.includes("action === 'REUSE'") && pointerWriter.includes('vNextAdminWriteActiveReleasePairCaches_')),
    'Canonical target with stale caches must not early-return before cache repair');
  assert.ok(pointerWriter.includes('vNextAdminReadActiveReleasePair_(hub)') &&
    pointerWriter.includes('vNextAdminActiveReleasePairMirrorsExact_'),
    'Canonical CAS and every mirror must be reread before publication continues');
  const scanStart = source.indexOf('function vNextAdminScanOneBook_');
  const scanEnd = source.indexOf('function vNextAdminHarvestClientRequests_', scanStart);
  const scan = source.slice(scanStart, scanEnd);
  assert.ok(scan.indexOf('vNextAdminAssertClientPinnedReleasePair_') >= 0 &&
    scan.indexOf('vNextAdminAssertClientPinnedReleasePair_') < scan.indexOf('vNextAdminHarvestClientRequests_'),
    'Health must validate the exact pinned pair before harvesting requests');
  const workerStart = source.indexOf('function vNextAdminExecuteJob_');
  const workerEnd = source.indexOf('function vNextAdminFinishJob_', workerStart);
  const worker = source.slice(workerStart, workerEnd);
  assert.ok(worker.lastIndexOf('vNextAdminAssertClientPinnedReleasePair_') < worker.indexOf('vNextRunForecast_(engineRequest)') &&
    worker.lastIndexOf('vNextAdminAssertClientPinnedReleasePair_') >= 0,
    'Worker must revalidate the exact pinned pair immediately before Engine entry');
  const migrationApiStart = source.indexOf('function vNextAdminEnqueueMigration(');
  const migrationApiEnd = source.indexOf('\nfunction ', migrationApiStart + 1);
  assert.ok(source.slice(migrationApiStart, migrationApiEnd).includes('VN_ADMIN_MIGRATION_APPLY_ENABLED'),
    'Migration APPLY must fail closed in the public enqueue API during pilot');
  const migrationWorkerStart = source.indexOf('function vNextAdminExecuteMigrationSkeleton_');
  const migrationWorkerEnd = source.indexOf('\nfunction ', migrationWorkerStart + 1);
  assert.ok(source.slice(migrationWorkerStart, migrationWorkerEnd).includes('VN_ADMIN_MIGRATION_APPLY_ENABLED'),
    'Already queued migration APPLY jobs must fail closed in the worker');
  assert.equal(sidebar.includes('migrationBookId'), false,
    'Migration controls must not be displayed in the pilot Admin Sidebar');
  assert.ok(source.includes('private_root_folder_id: folder.getId()') &&
    source.includes('vNextAdminFolderWithinRoot_') && source.includes('vNextAdminAssertClientFileAcl_') &&
    source.includes('DRIVE_NAME: VNEXT_NAMING.SHARED_DRIVE') &&
    source.includes('function vNextAdminRelocateLibraryToSharedDrive(') &&
    source.includes('vNextAdminPrepareLibraryDestinationFolder_'),
    'Client provisioning must stay inside the recorded library root, with a shared-drive relocate path');
  assert.ok(sidebar.includes('vNextAdminRelocateLibraryToSharedDrive') &&
    !sidebar.includes('処理を完了できませんでした。要確認事項と詳細を確認してください。'),
    'Admin Sidebar must show the actual relocate error instead of a generic placeholder');
  const relocateStart = source.indexOf('function vNextAdminMoveRegisteredFilesIntoLibrary_');
  const relocateEnd = source.indexOf('function vNextAdminFolderWithinRoot_', relocateStart);
  const relocate = source.slice(relocateStart, relocateEnd);
  assert.ok(relocate.includes('file.isTrashed()') && relocate.includes('aHub - bHub') &&
    relocate.includes("mode: 'PORTAL'") && relocate.includes('vNextAdminTryResolvePortal_'),
    'Library relocate must skip trashed files, move the running Hub last, and always move the employee Portal');
  assert.ok(source.includes('vNextAdminFindNamedChildFolder_') &&
    source.includes('vNextAdminDetectCurrentSharedDriveId_'),
    'Shared-drive relocate must reuse existing folders/drives instead of creating duplicates');
  assert.ok(source.includes('VN_ADMIN_PILOT_INITIAL_LIMIT = 3') &&
    source.includes('VN_ADMIN_PILOT_CANARY_LIMIT = 5'),
    'Pilot provisioning must gate at three and hard-stop after five Clients');
  assert.ok(source.includes('vNextAdminScanRegistryBatch_(hub, 10)') &&
    source.includes('vNextAdminProcessJobsForHub_(hub, 4'),
    'Scheduled sweep pilot batches must scan 10 books and process 4 jobs');
  assert.ok(source.includes('pilotRetries: vNextAdminRequeueKnownPilotFailures_(hub)') &&
    source.includes("String(job.error || '') !== 'A matching valid pending forecast request was not found.'") &&
    source.includes('function vNextAdminRequeueKnownPortalReleaseFailure_(') &&
    source.includes('Portal job requeued after stale Template release cache fix'),
    'The scheduler must narrowly recover both known pre-fix Pilot failures without generic retries');
  assert.ok(source.includes("script.setProperty('VNEXT_ACTIVE_RELEASE_ID'") &&
    source.includes("script.setProperty('VNEXT_MASTER_TEMPLATE_SPREADSHEET_ID'"),
    'Time-trigger Hub hydration must refresh both active Template caches');
  assert.ok(worker.includes("actorRole: 'ADMIN', hub: hub"),
    'Central-source forecast recovery must use the explicit Hub instead of UI focus');
}

function checkSourceContract() {
  const expected = {
    client: ['AO', 41, 'クライアント名'],
    serviceCategory: ['AT', 46, 'サービスカテゴリ'],
    product: ['AX', 50, '売上区分'],
    actualDate: ['BE', 57, '売上日'],
    amount: ['BN', 66, '金額']
  };
  const columns = engineSandbox.VNEXT_ENGINE.SOURCE_COLUMNS;
  const headers = engineSandbox.VNEXT_ENGINE.SOURCE_HEADERS;
  for (const [key, [, index, header]] of Object.entries(expected)) {
    assert.equal(columns[key], index, `${key} source column changed`);
    assert.equal(headers[key], header, `${key} source header changed`);
  }
  assert.equal(expected.actualDate[0], 'BE');
  assert.notEqual(expected.actualDate[0], 'BD');
}

async function checkVertexAiGroundingBudgetContracts() {
  const legacy = await readFile(path.join(root, 'Forecast_Agent.js'), 'utf8');
  const groundedStart = legacy.indexOf('function callVertexGeminiGrounded_');
  const groundedEnd = legacy.indexOf('function callVertexSearchRAG_', groundedStart);
  const grounded = legacy.slice(groundedStart, groundedEnd);
  assert.ok(grounded.includes('maxOutputTokens: 8192') &&
    grounded.includes("thinkingLevel: 'LOW'") &&
    grounded.includes("/^gemini-3(?:\\.|-|$)/i"),
    'Gemini 3 grounded research must use bounded LOW thinking and enough output budget to return citations');
  const structuredStart = legacy.indexOf('function callVertexGeminiStructured_');
  const structuredEnd = legacy.indexOf('function buildWebResearchPrompt_', structuredStart);
  assert.ok(legacy.slice(structuredStart, structuredEnd).includes("thinkingLevel: 'LOW'"),
    'Gemini 3 evidence structuring must use the same bounded thinking policy');
  const usageStart = legacy.indexOf('function extractGeminiUsage_');
  const usageEnd = legacy.indexOf('function extractGeminiFinishReason_', usageStart);
  assert.ok(legacy.slice(usageStart, usageEnd).includes('thoughtsTokenCount'),
    'AI audit usage must expose thought tokens for MAX_TOKENS diagnosis');
}
