#!/usr/bin/env node

import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const testDir = path.dirname(fileURLToPath(import.meta.url));
const sourceDir = path.resolve(testDir, '..', 'src');
let uuid = 0;
const sandbox = {
  console,
  Logger: { log() {} },
  SpreadsheetApp: { getActiveSpreadsheet: () => ({ getId: () => 'SHEET-1' }) },
  Session: {
    getActiveUser: () => ({ getEmail: () => 'owner@example.com' }),
    getEffectiveUser: () => ({ getEmail: () => 'owner@example.com' })
  },
  Utilities: {
    getUuid: () => `uuid-${++uuid}`,
    Charset: { UTF_8: 'UTF_8' },
    DigestAlgorithm: { SHA_256: 'SHA_256' },
    computeDigest(algorithm, value) { return [...createHash('sha256').update(value, 'utf8').digest()]; }
  },
  LockService: { getDocumentLock: () => ({ waitLock() {}, releaseLock() {} }) }
};
vm.createContext(sandbox);
for (const name of ['Client_Core.js', 'Client_Bridge.js']) {
  vm.runInContext(await readFile(path.join(sourceDir, name), 'utf8'), sandbox, { filename: name });
}

testCanonicalJson();
testStateEventAuthoritative();
testInternalOpenContributor();
testInputRoundBoundary();
testCommitmentEvidence();
testClientQueue();
testPlanSubmission();
await testStableViewSheetHandles();
await testBlankIdentityStillGetsMenu();
await testEmployeeUxContracts();
process.stdout.write('PASS client runtime behavior tests (10 suites)\n');

function testCanonicalJson() {
  assert.equal(sandbox.vNextCanonicalJson_({ b: 2, a: 1 }), '{"a":1,"b":2}');
  assert.equal(sandbox.vNextCanonicalJson_({ a: undefined, b: [1, undefined] }), '{"b":[1,null]}');
}

function testStateEventAuthoritative() {
  const originalReadConfig = sandbox.vNextClientReadConfig_;
  const originalReadRecords = sandbox.vNextReadRecords_;
  const originalActiveUser = sandbox.vNextActiveUserEmail_;
  sandbox.vNextClientReadConfig_ = () => ({ mode: 'CLIENT', book_id: 'BOOK-1', state: 'INPUT_OPEN' });
  sandbox.vNextActiveUserEmail_ = () => 'owner@example.com';
  sandbox.vNextReadRecords_ = (sheetName) => {
    if (sheetName === 'BOOK_META') return [{
      book_id: 'BOOK-1', client_id: 'CLIENT-1', client_name: 'テスト', fiscal_year: 2027,
      forecast_owner_email: 'owner@example.com', team_member_emails_json: '["owner@example.com"]',
      state: 'INPUT_OPEN', as_of: '2026-08-09', cutoff: '2026-07-31', schema_version: 'vnext-schema-2'
    }];
    if (sheetName === 'STATE_EVENT') return [{
      book_id: 'BOOK-1', from_state: 'RUNNING', to_state: 'READY_TO_RUN',
      reason: 'forecast_failed/INSUFFICIENT_CONFIRMED_HISTORY: 確定実績が不足',
      related_run_id: 'RUN-FAILED-1', created_at: '2026-08-11T03:04:05Z'
    }];
    if (sheetName === 'EVIDENCE_EVENT') return [];
    return [];
  };
  try {
    const context = sandbox.vNextGetBookContext_({ spreadsheet: { getId: () => 'SHEET-1' } });
    assert.equal(context.state, 'READY_TO_RUN');
    assert.equal(context.stateReason, 'forecast_failed/INSUFFICIENT_CONFIRMED_HISTORY: 確定実績が不足');
    assert.equal(context.stateChangedAt, '2026-08-11T03:04:05Z');
    assert.equal(context.relatedRunId, 'RUN-FAILED-1');
    sandbox.vNextActiveUserEmail_ = () => 'viewer@example.com';
    const viewer = sandbox.vNextGetBookContext_({ spreadsheet: { getId: () => 'SHEET-1' } });
    assert.equal(viewer.role, 'VIEWER');
    assert.equal(viewer.isTeamMember, false);
    assert.equal(viewer.isInternalUser, false);
    assert.equal(viewer.canContribute, false);
  } finally {
    sandbox.vNextClientReadConfig_ = originalReadConfig;
    sandbox.vNextReadRecords_ = originalReadRecords;
    sandbox.vNextActiveUserEmail_ = originalActiveUser;
  }
}

function testInternalOpenContributor() {
  const originalReadConfig = sandbox.vNextClientReadConfig_;
  const originalReadRecords = sandbox.vNextReadRecords_;
  const originalActiveUser = sandbox.vNextActiveUserEmail_;
  sandbox.vNextClientReadConfig_ = () => ({
    mode: 'CLIENT', book_id: 'BOOK-1', state: 'INPUT_OPEN',
    access_policy: 'INTERNAL_OPEN', internal_domain: '@example.com'
  });
  sandbox.vNextActiveUserEmail_ = () => 'contributor@example.com';
  sandbox.vNextReadRecords_ = (sheetName) => {
    if (sheetName === 'BOOK_META') return [{
      book_id: 'BOOK-1', client_id: 'CLIENT-1', client_name: 'テスト', fiscal_year: 2027,
      forecast_owner_email: 'owner@example.com',
      team_member_emails_json: '["owner@example.com","member@example.com"]',
      state: 'INPUT_OPEN', as_of: '2026-08-09', cutoff: '2026-07-31', schema_version: 'vnext-schema-2'
    }];
    if (sheetName === 'STATE_EVENT') return [];
    if (sheetName === 'EVIDENCE_EVENT') return [
      {
        book_id: 'BOOK-1', actor_email: 'owner@example.com', evidence_type: 'CHECK_IN',
        response_type: 'NO_CHANGE', status: 'SUBMITTED', created_at: '2026-08-10T00:00:00Z'
      },
      {
        book_id: 'BOOK-1', actor_email: 'contributor@example.com', evidence_type: 'HUMAN_CHANGE',
        response_type: 'CHANGE', status: 'SUBMITTED', created_at: '2026-08-10T00:01:00Z'
      }
    ];
    return [];
  };
  try {
    const contributor = sandbox.vNextGetBookContext_({ spreadsheet: { getId: () => 'SHEET-1' } });
    assert.equal(contributor.role, 'INTERNAL_CONTRIBUTOR');
    assert.equal(contributor.isForecastOwner, false);
    assert.equal(contributor.isTeamMember, false, 'open contributors must not become registered team members');
    assert.equal(contributor.isInternalUser, true);
    assert.equal(contributor.canContribute, true);
    assert.equal(contributor.inputStatus.submitted, true);
    assert.equal(contributor.inputStatus.answeredCount, 1, 'open contributor answers must not advance registered-team readiness');
    assert.equal(contributor.inputStatus.totalCount, 2);
    assert.equal(contributor.inputStatus.changeCount, 0, 'open contributor answers must not alter registered-team response counts');
    assert.equal(contributor.canProceed, false);

    sandbox.vNextActiveUserEmail_ = () => 'external@outside.test';
    const external = sandbox.vNextGetBookContext_({ spreadsheet: { getId: () => 'SHEET-1' } });
    assert.equal(external.role, 'VIEWER');
    assert.equal(external.isTeamMember, false);
    assert.equal(external.isInternalUser, false);
    assert.equal(external.canContribute, false);
  } finally {
    sandbox.vNextClientReadConfig_ = originalReadConfig;
    sandbox.vNextReadRecords_ = originalReadRecords;
    sandbox.vNextActiveUserEmail_ = originalActiveUser;
  }
}

function testInputRoundBoundary() {
  const originalReadConfig = sandbox.vNextClientReadConfig_;
  const originalReadRecords = sandbox.vNextReadRecords_;
  const originalActiveUser = sandbox.vNextActiveUserEmail_;
  sandbox.vNextClientReadConfig_ = () => ({ mode: 'CLIENT', book_id: 'BOOK-1', state: 'INPUT_OPEN' });
  sandbox.vNextActiveUserEmail_ = () => 'owner@example.com';
  sandbox.vNextReadRecords_ = (sheetName) => {
    if (sheetName === 'BOOK_META') return [
      {
        book_id: 'BOOK-1', client_id: 'CLIENT-1', client_name: 'テスト', fiscal_year: 2027,
        forecast_owner_email: 'owner@example.com', team_member_emails_json: '["owner@example.com","member@example.com"]',
        state: 'INPUT_OPEN', as_of: '2026-08-09', cutoff: '2026-07-31', schema_version: 'vnext-schema-2',
        event_type: 'CREATED', recorded_at: '2026-08-01T00:00:00Z'
      },
      {
        book_id: 'BOOK-1', client_id: 'CLIENT-1', client_name: 'テスト', fiscal_year: 2027,
        forecast_owner_email: 'owner@example.com', team_member_emails_json: '["owner@example.com","member@example.com"]',
        state: 'INPUT_OPEN', as_of: '2026-08-09', cutoff: '2026-07-31', schema_version: 'vnext-schema-2',
        event_type: 'INPUT_REOPENED', recorded_at: '2026-08-09T10:00:00Z'
      }
    ];
    if (sheetName === 'STATE_EVENT') return [];
    if (sheetName === 'EVIDENCE_EVENT') return [
      { book_id: 'BOOK-1', actor_email: 'owner@example.com', evidence_type: 'HUMAN_CHANGE', response_type: 'CHANGE', status: 'ACTIVE', created_at: '2026-08-09T09:59:59Z' },
      { book_id: 'BOOK-1', actor_email: 'owner@example.com', evidence_type: 'CHECK_IN', response_type: 'NO_CHANGE', status: 'ACTIVE', created_at: '2026-08-09T10:00:00Z' },
      { book_id: 'BOOK-1', actor_email: 'member@example.com', evidence_type: 'REVIEW_LEARNING', response_type: 'UNKNOWN', status: 'ACTIVE', created_at: '2026-08-09T10:01:00Z' },
      { book_id: 'BOOK-1', actor_email: 'member@example.com', evidence_type: 'AI_RESEARCH', response_type: 'CHANGE', status: 'ACTIVE', created_at: '2026-08-09T10:02:00Z' },
      { book_id: 'BOOK-1', actor_email: 'member@example.com', evidence_type: 'COMMITMENT', response_type: 'CHANGE', status: 'SUBMITTED', created_at: '2026-08-09T10:03:00Z' }
    ];
    return [];
  };
  try {
    const context = sandbox.vNextGetBookContext_({ spreadsheet: { getId: () => 'SHEET-1' } });
    assert.equal(context.inputStatus.roundStartedAt, '2026-08-09T10:00:00.000Z');
    assert.equal(context.inputStatus.answeredCount, 2);
    assert.equal(context.inputStatus.noChangeCount, 1);
    assert.equal(context.inputStatus.changeCount, 1);
    assert.equal(context.inputStatus.unknownCount, 0, 'REVIEW_LEARNING must not count as an answer');
    assert.equal(context.inputStatus.submitted, true);
    assert.equal(context.latestOwnEvidence.responseType, 'no_change');
  } finally {
    sandbox.vNextClientReadConfig_ = originalReadConfig;
    sandbox.vNextReadRecords_ = originalReadRecords;
    sandbox.vNextActiveUserEmail_ = originalActiveUser;
  }
}

function testCommitmentEvidence() {
  let appended = null;
  sandbox.vNextGetBookContext_ = () => ({
    bookId: 'BOOK-1', clientId: 'CLIENT-1', fiscalYear: 2027,
    userEmail: 'owner@example.com', isTeamMember: true, canContribute: true, state: 'INPUT_OPEN'
  });
  sandbox.vNextAppendRecord_ = (sheet, record) => { appended = { sheet, record }; };
  const result = sandbox.vNextAppendEvidence_({
    bookId: 'BOOK-1', responseType: 'change', evidenceType: 'COMMITMENT',
    changeKind: 'contract', target: '契約更新', direction: 'increase',
    confidence: 'confirmed', evidence: '顧客確認済み', amountMode: 'band',
    amountBand: 'medium', amountLow: 100, amountHigh: 200,
    period: { start: '2027-04-01', end: '2027-06-01' }
  }, { spreadsheet: {} });
  assert.equal(result.responseType, 'CHANGE');
  assert.equal(appended.sheet, 'EVIDENCE_EVENT');
  assert.equal(appended.record.evidence_type, 'COMMITMENT');
  assert.equal(appended.record.prompt_version, '');
  for (const responseType of ['no_change', 'unknown']) {
    appended = null;
    const checkIn = sandbox.vNextAppendEvidence_({
      bookId: 'BOOK-1', responseType, evidenceType: 'CHECK_IN', evidence: ''
    }, { spreadsheet: {} });
    assert.equal(checkIn.responseType, responseType.toUpperCase());
    assert.equal(appended.record.evidence_type, 'CHECK_IN');
    assert.equal(appended.record.target_start_month, '');
    assert.equal(appended.record.target_end_month, '');
  }
  assert.throws(() => sandbox.vNextAppendEvidence_({
    bookId: 'BOOK-1', responseType: 'change', evidenceType: 'AI_RESEARCH',
    target: 'x', direction: 'increase', confidence: 'confirmed', evidence: 'x'
  }, { spreadsheet: {} }), /情報の種類が不正/);
  sandbox.vNextGetBookContext_ = () => ({
    bookId: 'BOOK-1', clientId: 'CLIENT-1', fiscalYear: 2027,
    userEmail: 'external@outside.test', isTeamMember: false, canContribute: false, state: 'INPUT_OPEN'
  });
  assert.throws(() => sandbox.vNextAppendEvidence_({
    bookId: 'BOOK-1', responseType: 'no_change'
  }, { spreadsheet: {} }), /社内アカウント/);
}

function testClientQueue() {
  let requestEvent = null;
  let transition = null;
  sandbox.vNextGetBookContext_ = () => ({
    bookId: 'BOOK-1', clientId: 'CLIENT-1', clientName: 'テスト', fiscalYear: 2027,
    asOf: '2026-08-01', userEmail: 'owner@example.com', isForecastOwner: true,
    state: 'READY_TO_RUN'
  });
  sandbox.vNextClientWithDocumentLock_ = (operation) => operation();
  sandbox.vNextClientEnsureRequestSheet_ = () => ({});
  sandbox.vNextClientAppendRequestEvent_ = (ss, record) => { requestEvent = record; };
  sandbox.vNextTransitionState_ = (request) => { transition = request; return { stateEventId: 'STATE-1' }; };
  const result = sandbox.vNextQueueClientForecastRequest({ bookId: 'BOOK-1' });
  assert.equal(result.ok, true);
  assert.equal(requestEvent.status, 'PENDING');
  assert.equal(requestEvent.request_hash, sandbox.vNextSha256Hex_(requestEvent.request_json));
  const payload = JSON.parse(requestEvent.request_json);
  assert.equal(payload.requestedBy, 'owner@example.com');
  assert.equal(transition.internalOperation, 'CLIENT_QUEUE');
  assert.equal(transition.toState, 'RUNNING');
  assert.equal(transition.reason, 'forecast_requested:' + payload.requestId);
}

function testPlanSubmission() {
  let appended = null;
  sandbox.vNextGetBookContext_ = () => ({
    bookId: 'BOOK-1', fiscalYear: 2027, userEmail: 'owner@example.com',
    isForecastOwner: true, state: 'DRAFT_READY'
  });
  sandbox.vNextClientFindForecast_ = () => ({
    runId: 'RUN-1', status: 'SUCCESS', layers: { systemRecommended: 1000 }
  });
  sandbox.vNextGetLatestPlanVersion_ = () => null;
  sandbox.vNextAppendRecord_ = (sheet, record) => { appended = { sheet, record }; };
  const result = sandbox.vNextAppendPlanVersion_({
    bookId: 'BOOK-1', runId: 'RUN-1', adoptionDelta: 100,
    adoptionReason: '契約確認', salesUplift: 120, upliftReason: '追加提案',
    upliftOwner: 'Owner', upliftAction: '提案', upliftDueDate: '2027-09-30',
    upliftAllocation: Array.from({ length: 12 }, (_, index) => ({
      month: `${index < 9 ? 2027 : 2028}-${String(((index + 3) % 12) + 1).padStart(2, '0')}`,
      amount: 10
    }))
  }, { spreadsheet: {} });
  assert.equal(appended.sheet, 'PLAN_VERSION');
  assert.equal(result.adopted_forecast, 1100);
  assert.equal(result.final_budget, 1220);
  assert.equal(JSON.parse(result.uplift_allocation_json).length, 12);
}

async function testStableViewSheetHandles() {
  const uxSource = await readFile(path.join(sourceDir, 'VNext_UX.js'), 'utf8');
  assert.match(uxSource, /vNextUxRenderHome_\(context, home\)/);
  assert.match(uxSource, /vNextUxRenderPlan_\(context, forecast, plan\)/);
  assert.match(uxSource, /vNextUxRenderReview_\(context, review\)/);
  assert.match(uxSource, /function vNextUxRenderHome_\(context, sheet\)/);
  assert.match(uxSource, /function vNextUxRenderPlan_\(context, rawForecast, sheet\)/);
  assert.match(uxSource, /function vNextUxRenderReview_\(context, sheet\)/);
  assert.doesNotMatch(uxSource, /setFrozenRows\([^)]*\)\s*\.setHiddenGridlines/,
    'setFrozenRows returns void in Apps Script and must not be chained');
}

async function testBlankIdentityStillGetsMenu() {
  vm.runInContext(await readFile(path.join(sourceDir, 'VNext_UX.js'), 'utf8'), sandbox, { filename: 'VNext_UX.js' });
  const reviewBreakdown = sandbox.vNextUxEvaluationBreakdown_({
    base_level_error: 100, unknown_spot_error: -20, commitment_outcome_error: 30,
    amount_error: 0, human_info_error: 0, ai_info_error: 0, data_quality_error: -10,
    seasonality_error: 0, timing_error: 0
  });
  assert.equal(reviewBreakdown.filter(item => item.annual).length, 7);
  assert.equal(reviewBreakdown.find(item => item.key === 'BASE_LEVEL').explanation, '予測が実績より大きかった方向');
  sandbox.vNextUxGetLatestEvaluation_ = () => ({ evaluation_id: 'EVAL-1', official_vintage_id: 'OFF-1' });
  const normalizedReview = sandbox.vNextUxNormalizeReview_({
    causeCategories: ['TIMING', 'COMMITMENT_OUTCOME'],
    confirmedCause: '案件開始が遅れた', causeHypothesis: '', nextInformation: ['契約開始月']
  }, { bookId: 'BOOK-1', fiscalYear: 2027 });
  assert.deepEqual(JSON.parse(JSON.stringify(normalizedReview.causeCategories)), ['COMMITMENT_OUTCOME', 'TIMING']);
  assert.throws(() => sandbox.vNextUxNormalizeReview_({
    causeCategories: [], confirmedCause: '原因', nextInformation: ['確認事項']
  }, { bookId: 'BOOK-1', fiscalYear: 2027 }), /原因カテゴリ/);
  const publicForecast = sandbox.vNextUxPublicForecast_({
    layers: { historyBaseline: 1000, commitmentDelta: 100, objectiveForecast: 1100, systemRecommended: 1100 },
    annual: { p10: 900, p50: 1100, p90: 1300 },
    drivers: [], nextInformation: [], changeReasons: [], evidenceSummary: { commitment: 0, human: 0 }
  });
  assert.ok(publicForecast.drivers.length > 0, 'non-zero layers must produce readable drivers');
  assert.ok(publicForecast.nextInformation.length > 0, 'empty service suggestions must use derived next information');
  for (const responseType of ['no_change', 'unknown']) {
    let appended = null;
    sandbox.vNextUxGetBookContext_ = () => ({
      mode: 'CLIENT', bookId: 'BOOK-1', clientId: 'CLIENT-1', fiscalYear: 2027,
      userEmail: 'member@example.com', isTeamMember: true, isForecastOwner: false,
      state: 'INPUT_OPEN', latestOwnEvidence: null
    });
    sandbox.vNextUxGetLatestForecast_ = () => null;
    sandbox.vNextAppendEvidence_ = record => { appended = record; return { evidenceId: `E-${responseType}` }; };
    sandbox.vNextUxMaybeAdvanceInputState_ = () => {};
    sandbox.vNextRefreshEmployeeViews = () => {};
    const result = sandbox.vNextSaveEvidence({ responseType, evidence: '' });
    assert.equal(result.ok, true, `${responseType} must save successfully`);
    assert.equal(appended.evidenceType, 'CHECK_IN');
    assert.deepEqual(JSON.parse(JSON.stringify(appended.period)), { start: '', end: '' });
  }
  let menuAdded = false;
  const menu = {
    addItem() { return this; },
    addToUi() { menuAdded = true; return this; }
  };
  sandbox.SpreadsheetApp = {
    getUi: () => ({ createMenu: () => menu }),
    getActiveSpreadsheet: () => ({ getId: () => 'SHEET-1' })
  };
  sandbox.Session = {
    getActiveUser: () => ({ getEmail: () => '' }),
    getEffectiveUser: () => ({ getEmail: () => '' })
  };
  sandbox.vNextUxGetBookContext_ = () => { throw new Error('identity unavailable'); };
  assert.equal(sandbox.vNextBuildClientMenu_(), true);
  assert.equal(menuAdded, true);
}

async function testEmployeeUxContracts() {
  const uxSource = await readFile(path.join(sourceDir, 'VNext_UX.js'), 'utf8');
  assert.doesNotMatch(uxSource, /\.merge\s*\(/, 'employee sheets must not create merged cells');
  assert.match(uxSource, /sheet\.getLastRow\(\)/, 'view cleanup must be bounded to used rows');
  assert.match(uxSource, /sheet\.getLastColumn\(\)/, 'view cleanup must be bounded to used columns');
  assert.doesNotMatch(uxSource, /sheet\.clearContents\(\)/, 'view cleanup must not clear the full sheet');
  assert.doesNotMatch(uxSource, /insertImage\(|assignScript\(/,
    'employee home must not use an over-grid image as the primary action');
  assert.match(uxSource, /vNextUxAutoOpenGuidance_/,
    'employee open must provide a state-aware sidebar without waiting for a cell click');

  const scaleBands = sandbox.vNextUxBuildAmountBands_(null, { annualSalesBaseline: 200000000 });
  assert.equal(scaleBands.find(item => item.key === 'medium').low, 4000000);
  assert.equal(scaleBands.find(item => item.key === 'medium').high, 10000000);
  assert.match(scaleBands.find(item => item.key === 'medium').label, /2〜5%/);
  assert.equal(sandbox.vNextUxBuildAmountBands_(null, {})[0].available, false,
    'S/M/L must not fall back to a fixed yen amount when client scale is unknown');

  const issue = sandbox.vNextUxStateIssue_({
    state: 'READY_TO_RUN',
    stateReason: 'forecast_failed/INSUFFICIENT_CONFIRMED_HISTORY: raw details'
  });
  assert.equal(issue.key, 'HISTORY_SHORTAGE');
  assert.doesNotMatch(issue.instruction, /raw details|INSUFFICIENT/, 'raw internal failure reason must not reach employees');
  const action = sandbox.vNextUxGetPrimaryAction_({
    state: 'READY_TO_RUN', isForecastOwner: true, isTeamMember: true, canContribute: true,
    stateReason: 'forecast_failed/INSUFFICIENT_CONFIRMED_HISTORY: raw details'
  });
  assert.equal(action.key, 'WAIT', 'known failure must block a duplicate forecast request');
  assert.equal(sandbox.vNextUxGetPrimaryAction_({
    state: 'READY_TO_RUN', isForecastOwner: false, isTeamMember: false,
    canContribute: true, inputStatus: { submitted: false }
  }).key, 'INPUT', 'an unanswered internal contributor must still be able to respond before the owner runs the forecast');
  assert.equal(sandbox.vNextUxGetPrimaryAction_({
    state: 'READY_TO_RUN', isForecastOwner: false, isTeamMember: false,
    canContribute: true, inputStatus: { submitted: true }
  }).key, 'WAIT', 'an answered internal contributor should wait for the Forecast Owner');

  const readiness = sandbox.vNextUxPublicForecast_({
    annual: { p10: 800, p50: 1000, p90: 1300 },
    layers: { systemRecommended: 1000 },
    lenses: { evidenceReadiness: { level: 'NEEDS_ATTENTION', historyYearCount: 4, missingResponseRate: 0.5 } },
    evidenceSummary: {}
  }).evidenceCoverage;
  assert.equal(readiness.label, '確認余地が大きいです');
  assert.match(readiness.detail, /確定実績 4年度分/);

  for (const name of ['VNext_InputSidebar.html', 'VNext_PlanSidebar.html', 'VNext_ReviewSidebar.html']) {
    const html = await readFile(path.join(sourceDir, name), 'utf8');
    assert.match(html, /aria-live=/, `${name} must announce async status changes`);
  }
  const planHtml = await readFile(path.join(sourceDir, 'VNext_PlanSidebar.html'), 'utf8');
  assert.match(planHtml, /id="adoptedAmount"/, 'owner should enter the adopted forecast amount, not calculate a delta');
  assert.doesNotMatch(planHtml, /class="hero"/, 'plan sidebar must avoid decorative hero cards');
  const inputHtml = await readFile(path.join(sourceDir, 'VNext_InputSidebar.html'), 'utf8');
  assert.doesNotMatch(inputHtml, /id="previewButton"|内容を確認/,
    'employee evidence input must save in one server-validated operation');
  assert.match(inputHtml, /vNextSaveEvidence\(payload\(\)\)/);
  assert.match(inputHtml, /data-step="response"/);
  assert.match(inputHtml, /data-step="impact"/);
  assert.match(inputHtml, /function showStep\(key\)/,
    'evidence input should disclose one decision at a time');
  assert.match(planHtml, /data-step="adoption"/);
  assert.match(planHtml, /data-step="uplift"/);
  assert.match(planHtml, /function showStep\(key\)/,
    'plan editing should disclose adoption and uplift as separate decisions');
  const coreSource = await readFile(path.join(sourceDir, 'Client_Core.js'), 'utf8');
  assert.match(coreSource, /target_start_month[\s\S]*setNumberFormat\('@'\)/,
    'evidence month columns must be formatted as text before append');
}
