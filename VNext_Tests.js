/**
 * Forecast vNext pure-JS/GAS tests. Run runAllVNextTests() in Apps Script.
 * These tests do not read or mutate production spreadsheets.
 */

function runAllVNextTests() {
  var tests = [
    vNextTestCanonicalJsonAndSha256_,
    vNextTestSeededPrng_,
    vNextTestCutoff_,
    vNextTestAdminJobAsOfPrecedence_,
    vNextTestActualDateGuard_,
    vNextTestSourceHeaderContract_,
    vNextTestHistoryGuard_,
    vNextTestStateTransitions_,
    vNextTestAsOfEvidenceBoundary_,
    vNextTestInputRoundBoundary_,
    vNextTestAiEvidenceTemporalBoundary_,
    vNextTestExplanationAndInformationRanking_,
    vNextTestUnknownSpotLens_,
    vNextTestCalibratedAnnualUncertainty_,
    vNextTestAnnualSpotCoupling_,
    vNextTestAiFailureDoesNotWidenSalesVolatility_,
    vNextTestTriangulationMethods_,
    vNextTestMonthlyShockAndCommitmentDelay_,
    vNextTestReferencePriorSafety_,
    vNextTestCoherentSimulation_,
    vNextTestPublicAiEvidenceSanitization_,
    vNextTestEffectiveInputHash_,
    vNextTestRunVersionHash_,
    vNextTestInformationGapUncertainty_,
    vNextTestPlanValidation_,
    vNextTestOfficialRetryConsistency_,
    vNextTestLearningSeparation_,
    vNextTestAutomaticEvaluationBreakdown_,
    vNextTestTrustedRollbackContract_
  ];
  var results = [];
  tests.forEach(function (test) {
    var started = new Date();
    try {
      test();
      results.push({ name: test.name, status: 'PASS', durationMs: new Date().getTime() - started.getTime() });
    } catch (error) {
      results.push({
        name: test.name,
        status: 'FAIL',
        durationMs: new Date().getTime() - started.getTime(),
        error: String(error && error.stack ? error.stack : error)
      });
    }
  });
  var failures = results.filter(function (result) { return result.status === 'FAIL'; });
  vNextLog_('vNext tests: ' + vNextCanonicalJson_(results));
  if (failures.length) throw new Error(failures.length + ' vNext test(s) failed: ' + failures.map(function (item) { return item.name; }).join(', '));
  return { passed: results.length, failed: 0, results: results };
}

function runVNextCoreTests() {
  vNextTestCanonicalJsonAndSha256_();
  vNextTestSeededPrng_();
  vNextTestCutoff_();
  vNextTestActualDateGuard_();
  vNextTestSourceHeaderContract_();
  vNextTestStateTransitions_();
  vNextTestAsOfEvidenceBoundary_();
  vNextTestInputRoundBoundary_();
  vNextTestAiEvidenceTemporalBoundary_();
  vNextTestExplanationAndInformationRanking_();
  return { passed: 10, failed: 0 };
}

function runVNextEngineTests() {
  vNextTestAdminJobAsOfPrecedence_();
  vNextTestHistoryGuard_();
  vNextTestUnknownSpotLens_();
  vNextTestCalibratedAnnualUncertainty_();
  vNextTestAnnualSpotCoupling_();
  vNextTestAiFailureDoesNotWidenSalesVolatility_();
  vNextTestTriangulationMethods_();
  vNextTestMonthlyShockAndCommitmentDelay_();
  vNextTestReferencePriorSafety_();
  vNextTestCoherentSimulation_();
  vNextTestPublicAiEvidenceSanitization_();
  vNextTestEffectiveInputHash_();
  vNextTestRunVersionHash_();
  vNextTestInformationGapUncertainty_();
  vNextTestPlanValidation_();
  vNextTestOfficialRetryConsistency_();
  vNextTestLearningSeparation_();
  vNextTestAutomaticEvaluationBreakdown_();
  vNextTestTrustedRollbackContract_();
  return { passed: 19, failed: 0 };
}

function vNextTestCanonicalJsonAndSha256_() {
  var left = vNextCanonicalJson_({ z: 1, a: { y: 2, x: 1 }, list: [3, undefined, 1] });
  var right = vNextCanonicalJson_({ list: [3, null, 1], a: { x: 1, y: 2 }, z: 1 });
  vNextAssertEqual_(left, right, 'Canonical JSON must ignore object insertion order.');
  vNextAssertEqual_(
    vNextSha256Hex_('abc'),
    'ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad',
    'SHA-256 known vector'
  );
}

function vNextTestSeededPrng_() {
  var a = vNextCreatePrng_('book-2027');
  var b = vNextCreatePrng_('book-2027');
  for (var i = 0; i < 20; i++) vNextAssertNear_(a(), b(), 0, 'Seeded PRNG must be reproducible.');
  var c = vNextCreatePrng_('different');
  vNextAssertTrue_(Math.abs(a() - c()) > 1e-12, 'Different seeds should produce different streams.');
}

function vNextTestCutoff_() {
  vNextAssertEqual_(vNextFormatDateOnly_(vNextCutoffFromAsOf_('2027-04-15')), '2027-03-31', 'April cutoff');
  vNextAssertEqual_(vNextFormatDateOnly_(vNextCutoffFromAsOf_('2028-03-01')), '2028-02-29', 'Leap-year cutoff');
  vNextAssertEqual_(vNextFormatDateOnly_(vNextCutoffFromAsOf_('2027-01-01')), '2026-12-31', 'Year-boundary cutoff');
}

function vNextTestAdminJobAsOfPrecedence_() {
  vNextAssertEqual_(
    vNextResolveAdminJobAsOf_({ asOf: '2027-03-15' }, { asOf: '2027-02-01' }),
    '2027-03-15',
    'Verified client request date must take precedence over provision-time context as_of.'
  );
  vNextAssertEqual_(
    vNextFormatDateOnly_(vNextCutoffFromAsOf_(vNextResolveAdminJobAsOf_({ asOf: '2027-03-15' }, { asOf: '2027-02-01' }))),
    '2027-02-28',
    'Admin job cutoff remains the previous month-end of the request date.'
  );
}

function vNextTestActualDateGuard_() {
  vNextAssertThrows_(function () {
    vNextValidateActualRecords_([{ plannedDate: '2026-01-15', amount: 1 }], '2026-03-01');
  }, 'plannedDate', 'Planned-only records must fail.');
  vNextAssertThrows_(function () {
    vNextValidateActualRecords_([{ actualDate: '2026-03-01', amount: 1, dateSource: 'ACTUAL' }], '2026-03-01');
  }, 'after cutoff', 'Current-month actuals must fail.');
  var valid = vNextValidateActualRecords_([{ actualDate: '2026-02-28', amount: 1, dateSource: 'ACTUAL', isConfirmed: true }], '2026-03-01');
  vNextAssertEqual_(valid.length, 1, 'Previous-month confirmed actual is valid.');
}

function vNextTestSourceHeaderContract_() {
  var headers = [];
  Object.keys(VNEXT_ENGINE.SOURCE_COLUMNS).forEach(function (key) {
    headers[VNEXT_ENGINE.SOURCE_COLUMNS[key] - 1] = VNEXT_ENGINE.SOURCE_HEADERS[key];
  });
  vNextValidateSourceHeader_(headers, VNEXT_ENGINE.SOURCE_COLUMNS, '*2026_actual_value');
  headers[VNEXT_ENGINE.SOURCE_COLUMNS.actualDate - 1] = '売上予定日';
  vNextAssertThrows_(function () {
    vNextValidateSourceHeader_(headers, VNEXT_ENGINE.SOURCE_COLUMNS, '*2026_actual_value');
  }, 'planned-date', 'Default source schema must reject planned-date drift.');
}

function vNextTestHistoryGuard_() {
  var tooShort = vNextBuildTestActuals_(2021, 4, '2025-02-01');
  vNextAssertThrows_(function () {
    vNextBuildContinuityPrior_(tooShort, 2025, vNextCutoffFromAsOf_('2025-02-01'));
  }, 'At least 5 fiscal years', 'Four fiscal years must fail.');
  var sufficient = vNextBuildTestActuals_(2019, 6, '2025-02-01');
  var prior = vNextBuildContinuityPrior_(sufficient, 2025, vNextCutoffFromAsOf_('2025-02-01'));
  vNextAssertTrue_(prior.fiscalYears.length >= 5 && prior.fiscalYears.length <= 8, 'History must use 5-8 fiscal years.');
  vNextAssertNear_(vNextSum_(prior.seasonalShares), 1, 1e-9, 'Seasonal shares must sum to one.');
  var failure = vNextForecastFailureInfo_(new Error('At least 5 fiscal years of confirmed actual history are required; found 4.'));
  vNextAssertEqual_(failure.code, 'INSUFFICIENT_CONFIRMED_HISTORY', 'Insufficient history has a stable employee-facing code.');
  vNextAssertTrue_(failure.userMessage.indexOf('必要5年度') >= 0 && failure.userMessage.indexOf('4年度') >= 0,
    'Insufficient history explains the required and observed years in Japanese.');
  vNextAssertTrue_(!failure.retryRecommended, 'Missing confirmed history must not encourage a blind retry.');
}

function vNextTestStateTransitions_() {
  vNextAssertTrue_(vNextValidateTransition_('READY_TO_RUN', 'RUNNING', 'SYSTEM'), 'System can start run.');
  vNextAssertTrue_(vNextValidateTransition_('SUBMITTED', 'OFFICIAL_LOCKED', 'ADMIN'), 'Admin can approve.');
  vNextAssertTrue_(vNextValidateTransition_('CHANGES_REQUESTED', 'SUBMITTED', 'FORECAST_OWNER'), 'Owner can resubmit a corrected plan.');
  vNextAssertThrows_(function () {
    vNextValidateTransition_('SUBMITTED', 'OFFICIAL_LOCKED', 'MEMBER');
  }, 'cannot transition', 'Member cannot approve.');
  vNextAssertThrows_(function () {
    vNextValidateTransition_('OFFICIAL_LOCKED', 'INPUT_OPEN', 'ADMIN');
  }, 'not allowed', 'Official vintage cannot be reopened in place.');
  vNextAssertThrows_(function () {
    vNextValidateTransition_('CHANGES_REQUESTED', 'SUBMITTED', 'MEMBER');
  }, 'cannot transition', 'Member cannot resubmit a plan.');
}

function vNextTestAsOfEvidenceBoundary_() {
  var end = vNextAsOfEnd_('2025-02-01');
  vNextAssertTrue_(new Date('2025-02-01T12:00:00').getTime() <= end.getTime(), 'Evidence created on as_of must be included.');
  vNextAssertTrue_(new Date(2025, 1, 2, 0, 0, 0, 0).getTime() > end.getTime(), 'Evidence after as_of must be excluded.');
}

function vNextTestInputRoundBoundary_() {
  var cutoff = vNextLatestInputRoundCutoff_([
    { event_type: 'CREATED', recorded_at: '2026-01-01T00:00:00Z' },
    { event_type: 'INPUT_REOPENED', recorded_at: '2026-08-09T03:00:00Z' }
  ]);
  vNextAssertEqual_(cutoff.toISOString(), '2026-08-09T03:00:00.000Z', 'Latest reopen timestamp starts the input round.');
  vNextAssertTrue_(!vNextEvidenceInInputRound_({ created_at: '2026-08-09T02:59:59Z' }, cutoff), 'Earlier evidence is excluded from a reopened round.');
  vNextAssertTrue_(vNextEvidenceInInputRound_({ created_at: '2026-08-09T03:00:00Z' }, cutoff), 'Evidence at the cutoff is included.');
  vNextAssertTrue_(vNextEvidenceInInputRound_({ created_at: '2026-08-09T03:00:01Z' }, cutoff), 'New evidence is included.');
}

function vNextTestAiEvidenceTemporalBoundary_() {
  var base = {
    evidence_type: 'AI_RESEARCH', status: 'ACTIVE', source_date: '2025-02-01',
    expires_at: '2025-02-01', created_at: '2025-02-01T23:59:59Z',
    evidence_text: JSON.stringify({ parentRequestId: 'REQ-1', effectiveAsOf: '2025-02-01' })
  };
  vNextAssertTrue_(
    vNextEvidenceEffectiveForRun_(base, { asOf: '2025-02-01', requestId: 'REQ-1' }),
    'AI evidence is valid through its inclusive expiry date.'
  );
  var expired = vNextClonePlain_(base);
  expired.expires_at = '2025-01-31';
  vNextAssertTrue_(
    !vNextEvidenceEffectiveForRun_(expired, { asOf: '2025-02-01', requestId: 'REQ-1' }),
    'AI evidence expired before as_of must be excluded.'
  );
  var futureSource = vNextClonePlain_(base);
  futureSource.source_date = '2025-02-02';
  vNextAssertTrue_(
    !vNextEvidenceEffectiveForRun_(futureSource, { asOf: '2025-02-01', requestId: 'REQ-1' }),
    'AI sources published after as_of must be excluded.'
  );
  var delayed = vNextClonePlain_(base);
  delayed.expires_at = '2025-05-01';
  delayed.created_at = '2025-02-02T01:00:00Z';
  vNextAssertTrue_(
    vNextEvidenceEffectiveForRun_(delayed, { asOf: '2025-02-01', requestId: 'REQ-1' }),
    'Delayed AI completion linked to the exact request remains in that information vintage.'
  );
  vNextAssertTrue_(
    !vNextEvidenceEffectiveForRun_(delayed, { asOf: '2025-02-01', requestId: 'REQ-OTHER' }),
    'Delayed AI evidence cannot leak into another request vintage.'
  );
  vNextAssertTrue_(
    vNextEvidenceEffectiveForRun_(delayed, {
      asOf: '2025-02-01', requestId: 'REQ-ROLLBACK', allowedDelayedAiRequestIds: ['REQ-1', 'REQ-ROLLBACK']
    }),
    'A trusted rollback may reuse delayed AI evidence from its validated source request.'
  );
  var delayedHuman = vNextClonePlain_(delayed);
  delayedHuman.evidence_type = 'HUMAN_CHANGE';
  vNextAssertTrue_(
    !vNextEvidenceEffectiveForRun_(delayedHuman, { asOf: '2025-02-01', requestId: 'REQ-1' }),
    'Human evidence created after as_of is never backdated.'
  );
  vNextAssertThrows_(function () {
    vNextNormalizeRunRequest_({ trustedReuseSeedFromRunId: 'RUN-1' });
  }, 'only from an ADMIN_JOB', 'Client/untrusted callers cannot inject rollback trust fields.');
}

function vNextTestExplanationAndInformationRanking_() {
  var ranked = vNextRankNextInformation_({
    missingResponseRate: 0.5,
    informationGapRate: 0.25,
    evidenceResponseCounts: { unknown: 1 },
    commitmentEvents: [], humanEvents: [], aiEvents: []
  }, {
    unknownSpotModel: { expectedAnnual: 250, annualStd: 150 }
  }, {
    historyBaseline: 1000, systemRecommended: 1100
  }, {
    p10: 700, p50: 1100, p90: 1500
  });
  vNextAssertEqual_(ranked.length, 3, 'Employees receive at most three information requests.');
  vNextAssertTrue_(ranked[0].score >= ranked[1].score && ranked[1].score >= ranked[2].score, 'Information requests are value-ranked.');
  vNextAssertTrue_(ranked.every(function (row) {
    return row.expectedWidthReduction > 0 && row.amountImpact > 0 && row.confirmationBurden >= 1 && row.text;
  }), 'Each recommendation retains value and burden components.');
  var reasons = vNextBuildChangeReasons_({
    runId: 'RUN-OLD', historyBaseline: 1000, commitmentDelta: 0,
    referenceDelta: 0, humanDelta: 0, aiDelta: 0, systemRecommended: 1000
  }, {
    historyBaseline: 1000, commitmentDelta: 300,
    referenceDelta: 0, humanDelta: -100, aiDelta: 50, systemRecommended: 1250
  });
  vNextAssertEqual_(reasons.length, 3, 'The three largest prior-run changes are explained.');
  vNextAssertTrue_(reasons[0].indexOf('契約・案件') >= 0, 'Largest change is explained first.');
  var readiness = vNextBuildPublicEvidenceReadiness_({
    missingResponseRate: 0.5, informationGapRate: 0.25,
    evidenceResponseCounts: { change: 1, noChange: 1, unknown: 1 }, aiUnavailable: false
  }, { fiscalYears: [2019, 2020, 2021, 2022, 2023, 2024] }, { p10: 500, p50: 1000, p90: 1600 });
  vNextAssertEqual_(readiness.level, 'NEEDS_ATTENTION', 'Large evidence gaps must be visible without a false accuracy score.');
  vNextAssertTrue_(readiness.issues.length <= 3 && readiness.historyYearCount === 6,
    'Employee readiness keeps the three most useful facts and history coverage.');
  vNextAssertTrue_(JSON.stringify(readiness).indexOf('confidence') < 0,
    'Evidence readiness must not be mislabeled as statistical confidence.');
}

function vNextTestUnknownSpotLens_() {
  var prior = vNextBuildContinuityPrior_(
    vNextBuildTestActuals_(2019, 6, '2025-02-01'),
    2025,
    vNextCutoffFromAsOf_('2025-02-01')
  );
  vNextAssertTrue_(prior.baseAnnualBaseline > 0, 'BASE continuity must be estimated independently.');
  vNextAssertTrue_(prior.unknownSpotModel.expectedAnnual > 0, 'Unknown SPOT expected value must be explicit.');
  vNextAssertTrue_(prior.unknownSpotModel.historyOccurrenceCount > 0, 'SPOT occurrence history must be retained.');
  var suppression = vNextCommitmentSpotSuppression_([{ direction: 'UP', startMonth: '2025-07', endMonth: '2025-08' }], 2025);
  vNextAssertEqual_(suppression[3], 0.5, 'Known July commitment suppresses duplicate unknown SPOT.');
  vNextAssertEqual_(suppression[4], 0.5, 'Known August commitment suppresses duplicate unknown SPOT.');
}

function vNextTestCalibratedAnnualUncertainty_() {
  var stable = [100, 108, 116.64, 125.9712, 136.049].map(Math.log);
  var stableCalibration = vNextCalibrateAnnualLogSigma_(stable);
  vNextAssertEqual_(stableCalibration.method, 'RECENT_REGIME_DETREND_MAD_SHRINKAGE', 'Annual interval documents its calibration method.');
  vNextAssertTrue_(stableCalibration.logSigma <= 0.13,
    'A stable exponential trend must not be misclassified as high uncertainty.');
  var volatile = [100, 160, 80, 190, 95, 210].map(Math.log);
  var volatileCalibration = vNextCalibrateAnnualLogSigma_(volatile);
  vNextAssertTrue_(volatileCalibration.logSigma > stableCalibration.logSigma,
    'Unexplained annual deviations must widen the calibrated interval.');
  vNextAssertTrue_(volatileCalibration.logSigma <= VNEXT_ENGINE.MAX_ANNUAL_LOG_SIGMA,
    'Small-sample interval calibration remains bounded by the deployed policy.');
  var oldShockWithStableCurrentRegime = [100, 220, 72, 150, 158, 166, 174, 182].map(Math.log);
  var recentRegime = vNextCalibrateAnnualLogSigma_(oldShockWithStableCurrentRegime);
  vNextAssertEqual_(recentRegime.calibrationSampleSize, 5,
    'Ordinary forward uncertainty uses the fixed recent operating regime.');
  vNextAssertTrue_(recentRegime.logSigma < volatileCalibration.logSigma,
    'An old structural shock must not be treated as annually recurring volatility.');
}

function vNextTestAnnualSpotCoupling_() {
  var spotMonthly = {};
  for (var fy = 2019; fy <= 2025; fy++) {
    for (var monthIndex = 0; monthIndex < 12; monthIndex++) {
      spotMonthly[vNextFormatMonth_(new Date(fy, 3 + monthIndex, 1))] = 100 + monthIndex * 5 + (fy - 2019) * 2;
    }
  }
  var years = [2019, 2020, 2021, 2022, 2023, 2024, 2025];
  var model = vNextBuildUnknownSpotModel_(spotMonthly, years, new Date(2026, 2, 31));
  vNextAssertEqual_(model.samplingPolicy, 'ANNUAL_TOTAL_THEN_MONTH_ALLOCATION',
    'Recurring SPOT is sampled as one annual total before month allocation.');
  var rng = vNextCreatePrng_(20260816);
  var gaussian = vNextCreateGaussianSampler_(rng);
  var samples = [];
  for (var i = 0; i < 4000; i++) {
    samples.push(vNextSum_(vNextSampleUnknownSpot_(model, new Array(12).fill(0), rng, gaussian, 1, new Array(12).fill(1))));
  }
  var mean = vNextMean_(samples);
  vNextAssertNear_(mean, model.expectedAnnual, model.expectedAnnual * 0.06,
    'Annual coupling preserves the historical expected SPOT value.');
  vNextAssertTrue_(model.annualAmountCalibration.logSigma <= VNEXT_ENGINE.MAX_SPOT_LOG_SIGMA,
    'Annual SPOT uncertainty is robustly calibrated and bounded.');
}

function vNextTestAiFailureDoesNotWidenSalesVolatility_() {
  var baseline = vNextUncertaintyMultiplier_({ missingResponseRate: 0, informationGapRate: 0 });
  var aiUnavailable = vNextUncertaintyMultiplier_({ missingResponseRate: 0, informationGapRate: 0, aiUnavailable: true });
  vNextAssertNear_(baseline, aiUnavailable, 1e-12,
    'Missing AI research is a reference gap, not evidence of higher sales volatility.');
}

function vNextTestTriangulationMethods_() {
  var triangulation = vNextBuildTriangulationReference_({
    annualHistory: [
      { fiscalYear: 2020, baseTotal: 100 }, { fiscalYear: 2021, baseTotal: 110 },
      { fiscalYear: 2022, baseTotal: 130 }, { fiscalYear: 2023, baseTotal: 145 },
      { fiscalYear: 2024, baseTotal: 170 }
    ],
    unknownSpotModel: { expectedAnnual: 10 }
  }, 190);
  vNextAssertEqual_(triangulation.policy, 'INDEPENDENT_REFERENCES_NOT_AUTOMATICALLY_AVERAGED',
    'Triangulation methods are references, not an automatic ensemble.');
  vNextAssertEqual_(triangulation.methods.length, 4, 'Weighted level, regression, CAGR and simulation are shown.');
  var keys = triangulation.methods.map(function (method) { return method.key; });
  ['RECENT_WEIGHTED_AVERAGE', 'LINEAR_REGRESSION', 'DAMPED_CAGR', 'INTEGRATED_SIMULATION'].forEach(function (key) {
    vNextAssertTrue_(keys.indexOf(key) >= 0, 'Triangulation includes ' + key + '.');
  });
  vNextAssertTrue_(triangulation.methods.every(function (method) {
    return method.value >= 0 && method.assumption && method.basis;
  }), 'Every reference exposes its amount, assumption and basis.');
}

function vNextTestMonthlyShockAndCommitmentDelay_() {
  var shares = new Array(12).fill(1 / 12);
  var firstRng = vNextCreatePrng_(20260809);
  var first = vNextSampleMonthlyCommonShock_(shares, 0.12, vNextCreateGaussianSampler_(firstRng));
  var secondRng = vNextCreatePrng_(20260809);
  var second = vNextSampleMonthlyCommonShock_(shares, 0.12, vNextCreateGaussianSampler_(secondRng));
  vNextAssertEqual_(vNextCanonicalJson_(first), vNextCanonicalJson_(second), 'Month common shock must be seed-reproducible.');
  vNextAssertNear_(vNextSum_(first.shares), 1, 1e-12, 'Shocked monthly shares must preserve the FY total.');
  vNextAssertTrue_(vNextStdDev_(first.multipliers) > 0, 'Month common shock must vary monthly paths.');

  var event = {
    evidenceId: 'COMMIT-1', target: '契約開始', direction: 'UP', probability: 1,
    amountLow: 1200, amountMid: 1200, amountHigh: 1200,
    startMonth: '2025-06', endMonth: '2025-06', recognitionDelayMonths: 2
  };
  var diagnostics = vNextInitializeCommitmentDiagnostics_([event], 10000, 2025);
  var rng = vNextCreatePrng_(1);
  var layer = vNextSampleEventsLayer_([event], 10000, shares, 2025, rng, vNextCreateGaussianSampler_(rng), {
    recognitionDelay: true, diagnostics: diagnostics
  });
  vNextAssertNear_(layer[4], 1200, 1e-9, 'June commitment delayed two months must be recognized in August.');
  vNextAssertNear_(vNextSum_(layer), 1200, 1e-9, 'In-FY recognition delay preserves the event amount.');
  var audit = vNextFinalizeCommitmentDiagnostics_(diagnostics)[0];
  vNextAssertEqual_(audit.evidenceId, 'COMMIT-1', 'Delay audit retains evidence lineage.');
  vNextAssertNear_(audit.sampledMeanDelayMonths, 2, 1e-9, 'Sampled delay is retained per event.');

  var outside = {
    evidenceId: 'COMMIT-OUT', direction: 'UP', probability: 1,
    amountLow: 500, amountMid: 500, amountHigh: 500,
    startMonth: '2026-03', endMonth: '2026-03', recognitionDelayMonths: 1
  };
  var outsideDiagnostics = vNextInitializeCommitmentDiagnostics_([outside], 10000, 2025);
  var outsideRng = vNextCreatePrng_(2);
  var outsideLayer = vNextSampleEventsLayer_([outside], 10000, shares, 2025, outsideRng, vNextCreateGaussianSampler_(outsideRng), {
    recognitionDelay: true, diagnostics: outsideDiagnostics
  });
  vNextAssertNear_(vNextSum_(outsideLayer), 0, 1e-9, 'Revenue delayed beyond March must not be reallocated into the FY.');
  vNextAssertNear_(vNextFinalizeCommitmentDiagnostics_(outsideDiagnostics)[0].fullyOutsideFiscalYearRate, 1, 1e-9, 'Outside-FY delay is auditable.');
}

function vNextTestReferencePriorSafety_() {
  var shares = new Array(12).fill(1 / 12);
  var current = new Array(12).fill(100);
  function mustNotSample() { throw new Error('Gaussian sampler must not run for a disabled prior.'); }
  vNextAssertNear_(vNextSum_(vNextSampleReferenceLayer_({}, current, shares, mustNotSample)), 0, 1e-12, 'Missing reference prior has zero delta.');

  var disabledByStrength = vNextNormalizeReferencePrior_({ growthMean: 0.20, growthStd: 0.10, strength: 0 }, 'CLIENT-1');
  vNextAssertEqual_(disabledByStrength.strength, 0, 'Explicit strength=0 must be retained.');
  vNextAssertNear_(vNextSum_(vNextSampleReferenceLayer_(disabledByStrength, current, shares, mustNotSample)), 0, 1e-12, 'strength=0 disables the reference delta.');

  var unsafeGlobal = vNextNormalizeReferencePrior_({ annualMean: 999999999, annualStd: 0, strength: 0.2 }, 'CLIENT-1');
  vNextAssertEqual_(unsafeGlobal.mode, 'DISABLED', 'Global absolute-yen prior must fail closed.');
  vNextAssertNear_(vNextSum_(vNextSampleReferenceLayer_(unsafeGlobal, current, shares, mustNotSample)), 0, 1e-12, 'Unsafe global absolute prior has zero delta.');
  var mismatched = vNextNormalizeReferencePrior_({
    scope: 'CLIENT', clientId: 'CLIENT-OTHER', annualMean: 1800, annualStd: 0, strength: 0.2
  }, 'CLIENT-1');
  vNextAssertEqual_(mismatched.mode, 'DISABLED', 'Absolute prior for another client must fail closed.');

  var absolute = vNextNormalizeReferencePrior_({
    scope: 'CLIENT', clientId: 'CLIENT-1', annualMean: 1800, annualStd: 0, strength: 0.2
  }, 'CLIENT-1');
  vNextAssertEqual_(absolute.mode, 'ABSOLUTE', 'Client-bound absolute prior is accepted.');
  vNextAssertNear_(vNextSum_(vNextSampleReferenceLayer_(absolute, current, shares, function () { return 0; })), 120, 1e-9, 'Client absolute prior applies a strength-weighted delta.');

  var growth = vNextNormalizeReferencePrior_({ growthMean: 0.10, growthStd: 0.20, strength: 0.25 }, 'CLIENT-1');
  vNextAssertEqual_(growth.mode, 'GROWTH', 'Scale-independent growth prior is accepted.');
  vNextAssertNear_(vNextSum_(vNextSampleReferenceLayer_(growth, current, shares, function () { return 1; })), 90, 1e-9, 'Growth mean/std are rates of the current annual amount.');
  var capped = vNextNormalizeReferencePrior_({ growthMean: 0.10, growthStd: 0, strength: 1 }, 'CLIENT-1');
  vNextAssertNear_(capped.strength, 0.35, 1e-12, 'Reference strength remains capped at 0.35.');
  var invalid = vNextNormalizeReferencePrior_({ growthMean: 4, growthStd: 0.1, strength: 0.2 }, 'CLIENT-1');
  vNextAssertEqual_(invalid.mode, 'DISABLED', 'Out-of-range growth prior fails closed.');
  vNextAssertThrows_(function () {
    vNextValidateModelReleaseParameters_({
      referencePrior: { growthMean: 0.1, growthStd: 0.1, strength: 0.2 }
    });
  }, 'disabled for the initial pilot', 'MODEL_RELEASE alone cannot enable a pilot reference prior.');
  var pilotDisabled = vNextValidateModelReleaseParameters_({
    referencePrior: { mode: 'DISABLED', reason: 'NO_COHORT_SNAPSHOT', strength: 0 }
  });
  vNextAssertEqual_(pilotDisabled.referencePrior.mode, 'DISABLED', 'Explicit disabled reference prior remains valid.');
}

function vNextTestCoherentSimulation_() {
  var actuals = vNextBuildTestActuals_(2019, 6, '2025-02-01');
  var request = {
    bookId: 'BOOK-TEST',
    clientId: 'CLIENT-TEST',
    clientName: 'テスト製薬',
    fiscalYear: 2025,
    asOf: '2025-02-01',
    actualRecords: actuals,
    seed: 123456,
    simulationCount: 400,
    persist: false,
    manageState: false,
    commitmentEvents: [{
      evidenceId: 'COMMIT-COHERENCE', direction: 'UP', amount: 1200000, probability: 0.75,
      startMonth: '2025-07', endMonth: '2025-09'
    }],
    objectiveEvents: [{
      direction: 'DOWN', amountBand: 'SMALL', confidence: 'LIKELY',
      startMonth: '2025-10', endMonth: '2026-03'
    }],
    humanEvents: [{
      direction: 'UP', rate: 0.03, confidence: 'HYPOTHESIS',
      startMonth: '2025-04', endMonth: '2026-03'
    }],
    aiEvents: [{
      direction: 'UP', rate: 0.25, confidence: 'CONFIRMED_FACT',
      startMonth: '2025-04', endMonth: '2026-03'
    }],
    referencePrior: { growthMean: 0.04, growthStd: 0.02, strength: 0.15 }
  };
  request.missingResponseRate = 0.20;
  request.informationGapRate = 0.25;
  request.evidenceResponseCounts = { change: 2, noChange: 3, unknown: 1 };
  var first = vNextEngineRunForecast(request);
  var second = vNextEngineRunForecast(request);
  vNextAssertEqual_(vNextCanonicalJson_(first.annual), vNextCanonicalJson_(second.annual), 'Same input and seed must reproduce annual results.');
  vNextAssertEqual_(vNextCanonicalJson_(first.months), vNextCanonicalJson_(second.months), 'Same input and seed must reproduce monthly results.');
  ['p10', 'p50', 'p90'].forEach(function (key) {
    var monthTotal = vNextSum_(first.months.map(function (row) { return row[key]; }));
    var quarterTotal = vNextSum_(first.quarters.map(function (row) { return row[key]; }));
    vNextAssertNear_(monthTotal, first.annual[key], 1e-6, key + ' month-to-year coherence');
    vNextAssertNear_(quarterTotal, first.annual[key], 1e-6, key + ' quarter-to-year coherence');
  });
  vNextAssertTrue_(first.annual.p10 <= first.annual.p50 && first.annual.p50 <= first.annual.p90, 'Annual quantiles must be ordered.');
  var layers = first.layers;
  vNextAssertNear_(layers.historyBaseline + layers.commitmentDelta + layers.referenceDelta, layers.objectiveForecast, 1e-6, 'Objective waterfall identity');
  vNextAssertNear_(layers.objectiveForecast + layers.humanDelta + layers.aiDelta, layers.systemRecommended, 1e-6, 'Recommended waterfall identity');
  var preAi = layers.objectiveForecast + layers.humanDelta;
  vNextAssertTrue_(Math.abs(layers.aiDelta) <= preAi * 0.05 + 1e-6, 'AI independent delta must stay inside ±5%.');
  vNextAssertTrue_(first.lenses.continuity.unknownSpotModel.expectedAnnual > 0, 'Continuity lens exposes expected unknown SPOT.');
  vNextAssertTrue_(first.lenses.continuity.monthlyCommonShockSigma > 0, 'Continuity lens exposes the monthly common shock.');
  vNextAssertEqual_(first.lenses.commitment.delayByEvent[0].evidenceId, 'COMMIT-COHERENCE', 'Commitment delay audit retains event lineage.');
  vNextAssertTrue_(first.lenses.simulationDesign.monthlyCommonShock.sharedAcrossLayers, 'Monthly common shock is shared across layers.');
  vNextAssertEqual_(first.lenses.simulationDesign.annualCommonShock.calibration.method, 'RECENT_REGIME_DETREND_MAD_SHRINKAGE',
    'Annual interval calibration is auditable.');
  vNextAssertTrue_(first.lenses.triangulation.methods.length >= 4,
    'Independent classical and simulation references are persisted for triangulation.');
  vNextAssertEqual_(first.evidenceSummary.noChange, 3, 'NO_CHANGE remains explicit evidence.');
  vNextAssertEqual_(first.evidenceSummary.unknown, 1, 'UNKNOWN remains an information gap.');
  var counterfactual = first.lenses.changeReference.aiCounterfactual;
  vNextAssertTrue_(!!counterfactual && !!counterfactual.annual, 'AI OFF counterfactual must be recorded.');
  vNextAssertTrue_(first.lenses.changeReference.aiCapApplied, 'AI cap application must be auditable.');
  ['p10', 'p50', 'p90'].forEach(function (key) {
    vNextAssertNear_(vNextSum_(counterfactual.months.map(function (row) { return row[key]; })), counterfactual.annual[key], 1e-6, 'AI OFF month coherence ' + key);
    vNextAssertNear_(vNextSum_(counterfactual.quarters.map(function (row) { return row[key]; })), counterfactual.annual[key], 1e-6, 'AI OFF quarter coherence ' + key);
  });
}

function vNextTestPublicAiEvidenceSanitization_() {
  var rows = vNextBuildPublicAiEvidence_([{
    target: '市場変化', direction: 'UP', amountMid: 300,
    sourceUrl: 'https://example.com/source', sourceDate: '2026-08-01',
    evidenceQuality: 'A', capApplied: true,
    evidenceText: JSON.stringify({
      summary: '公開してよい要約', promptVersion: 'SECRET-PROMPT', aiModel: 'SECRET-MODEL', raw: 'SECRET-RAW'
    })
  }]);
  vNextAssertEqual_(rows.length, 1, 'One public AI citation is retained.');
  vNextAssertEqual_(rows[0].summary, '公開してよい要約', 'Sanitized summary is retained.');
  vNextAssertTrue_(!Object.prototype.hasOwnProperty.call(rows[0], 'promptVersion'), 'Prompt version must not leave Admin Hub in public evidence.');
  vNextAssertTrue_(!Object.prototype.hasOwnProperty.call(rows[0], 'aiModel'), 'AI model must not leave Admin Hub in public evidence.');
  vNextAssertTrue_(JSON.stringify(rows[0]).indexOf('SECRET-RAW') < 0, 'Raw AI payload must not be exposed.');
}

function vNextTestEffectiveInputHash_() {
  var base = vNextBuildMinimalRunRequest_();
  var first = vNextEngineRunForecast(base);
  var changedGap = vNextClonePlain_(base);
  changedGap.informationGapRate = 0.5;
  var second = vNextEngineRunForecast(changedGap);
  vNextAssertTrue_(first.inputDataHash !== second.inputDataHash, 'Information-gap setting must enter input hash.');
  var changedCount = vNextClonePlain_(base);
  changedCount.simulationCount = 500;
  var third = vNextEngineRunForecast(changedCount);
  vNextAssertTrue_(first.inputDataHash !== third.inputDataHash, 'Simulation count must enter input hash.');
  var changedEvidenceVersion = vNextClonePlain_(base);
  changedEvidenceVersion.humanEvents = [{ evidenceId: 'EV-2', ruleVersion: 'RULE-2', direction: 'UP', amount: 1, probability: 1 }];
  var fourth = vNextEngineRunForecast(changedEvidenceVersion);
  vNextAssertTrue_(first.inputDataHash !== fourth.inputDataHash, 'Evidence IDs and rule versions must enter input hash.');
}

function vNextTestRunVersionHash_() {
  var request = vNextBuildMinimalRunRequest_();
  request.modelReleaseId = 'MODEL-1';
  request.schemaVersion = 'BOOK-SCHEMA-1';
  request.templateVersion = 'TEMPLATE-1';
  var first = vNextEngineRunForecast(request);
  vNextAssertEqual_(first.modelReleaseId, 'MODEL-1', 'Authorized model release must survive normalization.');
  vNextAssertEqual_(first.versions.bookSchema, 'BOOK-SCHEMA-1', 'Book schema version must be retained.');
  vNextAssertEqual_(first.versions.template, 'TEMPLATE-1', 'Template version must be retained.');
  var changed = vNextClonePlain_(request);
  changed.modelReleaseId = 'MODEL-2';
  var second = vNextEngineRunForecast(changed);
  vNextAssertTrue_(first.inputDataHash !== second.inputDataHash, 'Model release version must enter input hash.');
  var bound = {
    model_release_id: 'MODEL-1', status: 'ACTIVE', model_version: VNEXT_ENGINE.VERSION,
    schema_version: VNEXT_CORE.SCHEMA_VERSION, template_version: 'TEMPLATE-1'
  };
  vNextAssertTrue_(vNextAssertModelReleaseRuntimeBinding_(bound, 'MODEL-1', 'TEMPLATE-1'),
    'Exact model/template/runtime binding is accepted.');
  var simulationStarted = false;
  var persistenceStarted = false;
  var stateMutationStarted = false;
  vNextAssertThrows_(function () {
    vNextAssertModelReleaseRuntimeBinding_(bound, 'MODEL-1', 'TEMPLATE-OTHER');
    simulationStarted = true;
    persistenceStarted = true;
    stateMutationStarted = true;
  }, 'exactly match', 'Template mismatch fails before downstream run side effects.');
  vNextAssertTrue_(!simulationStarted && !persistenceStarted && !stateMutationStarted,
    'Release mismatch stops before simulation, persistence and state mutation.');
  var wrongEngine = vNextClonePlain_(bound);
  wrongEngine.model_version = 'OTHER-ENGINE';
  vNextAssertThrows_(function () {
    vNextAssertModelReleaseRuntimeBinding_(wrongEngine, 'MODEL-1', 'TEMPLATE-1');
  }, 'runtime binding', 'Engine mismatch fails closed.');
  vNextAssertThrows_(function () {
    vNextAssertRuntimeReleaseRequest_('', 'TEMPLATE-1', VNEXT_CORE.SCHEMA_VERSION, { requiresBoundRelease: true });
  }, 'modelReleaseId', 'Persisted forecast requires a pinned model release.');
  vNextAssertThrows_(function () {
    vNextAssertRuntimeReleaseRequest_('MODEL-1', 'TEMPLATE-1', 'WRONG-SCHEMA', { requiresBoundRelease: true });
  }, 'Book schema', 'Persisted forecast requires the deployed Core schema.');
}

function vNextTestInformationGapUncertainty_() {
  var base = vNextBuildMinimalRunRequest_();
  base.actualRecords = base.actualRecords.map(function (record) {
    var copy = {};
    Object.keys(record).forEach(function (key) { copy[key] = record[key]; });
    copy.serviceType = 'BASE';
    return copy;
  });
  base.seed = 998877;
  base.simulationCount = 800;
  base.informationGapRate = 0;
  var known = vNextEngineRunForecast(base);
  var gapRequest = vNextClonePlain_(base);
  gapRequest.informationGapRate = 1;
  var gap = vNextEngineRunForecast(gapRequest);
  var knownWidth = known.annual.p90 - known.annual.p10;
  var gapWidth = gap.annual.p90 - gap.annual.p10;
  vNextAssertTrue_(gapWidth > knownWidth, 'UNKNOWN information gaps must widen forecast uncertainty.');
}

function vNextTestPlanValidation_() {
  var zero = vNextValidateUpliftAllocation_([], 0, 2025);
  vNextAssertEqual_(zero.length, 12, 'Zero uplift is normalized to twelve coherent months.');
  vNextAssertEqual_(zero[0].month, '2025-04', 'Normalized allocation starts in April.');
  var allocation = new Array(12).fill(100);
  var numeric = vNextValidateUpliftAllocation_(allocation, 1200, 2025);
  vNextAssertEqual_(numeric.length, 12, 'Twelve numeric uplift months are valid.');
  vNextAssertEqual_(numeric[11].month, '2026-03', 'Numeric form is normalized to FY month objects.');
  var uxShape = numeric.map(function (item) { return { month: item.month, amount: item.amount }; });
  vNextAssertEqual_(vNextCanonicalJson_(vNextValidateUpliftAllocation_(uxShape, 1200, 2025)), vNextCanonicalJson_(numeric), 'UX object shape and numeric shape normalize identically.');
  vNextAssertThrows_(function () { vNextValidateUpliftAllocation_(new Array(11).fill(100), 1100, 2025); }, '12 non-negative', 'Eleven months must fail.');
  vNextAssertThrows_(function () { vNextValidateUpliftAllocation_(allocation, 1300, 2025); }, 'must equal', 'Allocation must equal uplift.');
  var wrongMonth = uxShape.map(function (item) { return { month: item.month, amount: item.amount }; });
  wrongMonth[0].month = '2025-05';
  vNextAssertThrows_(function () { vNextValidateUpliftAllocation_(wrongMonth, 1200, 2025); }, 'month mismatch', 'Wrong FY month labels must fail.');
}

function vNextTestOfficialRetryConsistency_() {
  vNextAssertTrue_(vNextValidateOfficialIssueRequest_('SUBMITTED', false, false, ''), 'Initial official issue is valid from SUBMITTED.');
  vNextAssertTrue_(vNextValidateOfficialIssueRequest_('OFFICIAL_LOCKED', true, true, '訂正理由'), 'Versioned amendment is valid from OFFICIAL_LOCKED.');
  vNextAssertThrows_(function () {
    vNextValidateOfficialIssueRequest_('OFFICIAL_LOCKED', true, true, '');
  }, 'amendmentReason', 'Amendment requires a reason.');
  vNextAssertThrows_(function () {
    vNextValidateOfficialIssueRequest_('OFFICIAL_LOCKED', true, false, '訂正理由');
  }, 'existing official', 'Amendment requires an existing vintage.');
  vNextAssertThrows_(function () {
    vNextValidateOfficialIssueRequest_('OFFICIAL_LOCKED', false, true, '');
  }, 'requires SUBMITTED', 'Non-amendment cannot issue from OFFICIAL_LOCKED.');
  var source = {
    runId: 'RUN-1', bookId: 'BOOK-1', clientId: 'CLIENT-1', clientName: 'テスト',
    fiscalYear: 2025, asOf: '2025-02-01', cutoff: '2025-01-31', seed: 123,
    inputDataHash: 'HASH-1', modelReleaseId: 'MODEL-1', simulationCount: 400,
    layers: { systemRecommended: 100 }, annual: { p10: 80, p50: 100, p90: 120 },
    quarters: [], months: [], lenses: {}, evidenceSummary: {}
  };
  var official = vNextClonePlain_(source);
  official.runId = 'OFFICIAL-RUN-1';
  official.previousRunId = 'RUN-1';
  official.officialVintageId = 'OFF-1';
  vNextAssertTrue_(vNextAssertOfficialRetryConsistent_(official, source, 'RUN-1', 'BOOK-1'), 'Identical retry must be idempotent.');
  var corrupted = vNextClonePlain_(official);
  corrupted.inputDataHash = 'CORRUPTED';
  vNextAssertThrows_(function () {
    vNextAssertOfficialRetryConsistent_(corrupted, source, 'RUN-1', 'BOOK-1');
  }, 'content does not match', 'Same vintage ID with different content must fail.');
  vNextAssertThrows_(function () {
    vNextAssertOfficialRetryConsistent_(official, source, 'RUN-OTHER', 'BOOK-1');
  }, 'source run mismatch', 'Same vintage ID with another source run must fail.');
}

function vNextTestLearningSeparation_() {
  var forecast = {
    runId: 'RUN-1', officialVintageId: 'VINTAGE-1', modelReleaseId: 'MODEL-1',
    inputDataHash: 'HASH', layers: { systemRecommended: 100 },
    annual: { p10: 80, p50: 100, p90: 120 }, evidenceSummary: { human: 2 }
  };
  var payload = vNextBuildLearningPayload_(forecast, {
    actualTotal: 90,
    adoptedForecast: 110,
    salesUplift: 20,
    finalBudget: 130,
    errorComponents: { humanInfo: -3 }
  });
  vNextAssertTrue_(!Object.prototype.hasOwnProperty.call(payload, 'adoptedForecast'), 'Adopted forecast must not enter learning payload.');
  vNextAssertTrue_(!Object.prototype.hasOwnProperty.call(payload, 'salesUplift'), 'Sales uplift must not enter learning payload.');
  vNextAssertTrue_(!Object.prototype.hasOwnProperty.call(payload, 'finalBudget'), 'Final budget must not enter learning payload.');
  vNextAssertEqual_(payload.systemForecast, 100, 'System forecast remains learning target.');
}

function vNextTestAutomaticEvaluationBreakdown_() {
  var forecast = {
    runId: 'OFFICIAL-RUN', officialVintageId: 'OFFICIAL-VINTAGE',
    layers: {
      historyBaseline: 1200, commitmentDelta: 100, referenceDelta: -50,
      humanDelta: 20, aiDelta: 10, systemRecommended: 1280
    },
    months: new Array(12).fill(0).map(function (_, index) { return { month: index, p50: 1280 / 12 }; }),
    lenses: { continuity: { unknownSpotModel: { expectedAnnual: 200 } } }
  };
  var actualBase = [960, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
  var actualSpot = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 120];
  var result = vNextEngineBuildEvaluationBreakdown({
    officialForecast: forecast,
    actualBaseMonths: actualBase,
    actualSpotMonths: actualSpot
  });
  vNextAssertNear_(result.systemSignedError, 200, 1e-9, 'Signed FY error is system forecast minus actual.');
  vNextAssertNear_(result.errorComponents.baseLevel, 40, 1e-9, 'BASE error is isolated.');
  vNextAssertNear_(result.errorComponents.unknownSpot, 80, 1e-9, 'Unknown SPOT error is isolated.');
  vNextAssertNear_(result.errorComponents.commitmentOutcome, 100, 1e-9, 'Commitment outcome remains a separate component.');
  vNextAssertNear_(result.errorComponents.amount, -50, 1e-9, 'Reference/amount remains a separate component.');
  vNextAssertNear_(result.componentSum, result.systemSignedError, 1e-9, 'Annual components reconcile additively.');
  vNextAssertTrue_(result.reconciled, 'Reconciliation flag must be true.');
  vNextAssertTrue_(result.diagnostics.seasonality.amountMagnitude > 0, 'Seasonality is diagnosed from monthly shapes.');
  vNextAssertTrue_(result.diagnostics.timing.amountMagnitude > 0, 'Timing is diagnosed from cumulative monthly shapes.');
  vNextAssertEqual_(result.errorComponents.seasonality, 0, 'Seasonality does not double-count the FY total.');
  vNextAssertEqual_(result.errorComponents.timing, 0, 'Timing does not double-count the FY total.');
  var mismatched = vNextEngineBuildEvaluationBreakdown({
    officialForecast: forecast,
    actualBaseMonths: actualBase,
    actualSpotMonths: actualSpot,
    actualTotal: 1100
  });
  vNextAssertNear_(mismatched.errorComponents.dataQuality, -20, 1e-9, 'Actual monthly/annual mismatch is reconciled through data quality.');
  vNextAssertTrue_(mismatched.diagnostics.dataQuality.issues.length > 0, 'Data-quality mismatch is explained.');
}

function vNextTestTrustedRollbackContract_() {
  vNextTestAdminRunIdempotencyContract_();
  var basisRequest = vNextBuildMinimalRunRequest_();
  basisRequest.seed = 314159;
  basisRequest.simulationCount = 400;
  basisRequest.aiEvents = [{
    evidenceId: 'AI-1', direction: 'UP', amount: 1, probability: 1,
    startMonth: '2025-04', endMonth: '2026-03'
  }];
  var source = vNextEngineRunForecast(basisRequest);
  vNextAssertEqual_(source.evidenceSummary.effectiveEvidenceIds.ai[0], 'AI-1',
    'Source run snapshots the exact AI evidence lineage.');
  vNextAssertTrue_(/^[a-f0-9]{64}$/.test(source.evidenceSummary.nonAiComparableHash),
    'Source run stores a non-AI comparable-input hash.');
  var originalFind = vNextFindForecastByRunId_;
  vNextFindForecastByRunId_ = function (bookId, runId) {
    return bookId === source.bookId && runId === source.runId ? source : null;
  };
  try {
    var rollback = vNextClonePlain_(basisRequest);
    // Preserve GAS Date values; JSON cloning would reinterpret midnight across
    // time zones and correctly trip the comparable-input guard.
    rollback.actualRecords = basisRequest.actualRecords;
    delete rollback.seed;
    rollback.internalOperation = 'ADMIN_JOB';
    rollback.requestId = 'REQ-ROLLBACK-1';
    rollback.trustedReuseSeedFromRunId = source.runId;
    rollback.trustedAllowedDelayedAiRequestIds = ['REQ-ROLLBACK-1'];
    rollback.trustedRollbackContext = {
      operationId: 'ROLLBACK-1', sourceForecastRunId: source.runId,
      sourceInputDataHash: source.inputDataHash, sourceModelReleaseId: source.modelReleaseId,
      nonAiComparableHash: source.evidenceSummary.nonAiComparableHash,
      scope: 'ALL', targetEvidenceIds: ['AI-1'], tombstoneEvidenceIds: ['AI-RB-1']
    };
    var rerun = vNextEngineRunForecast(rollback);
    vNextAssertEqual_(rerun.seed, source.seed, 'Trusted rollback must reuse the source seed.');
    vNextAssertEqual_(rerun.simulationCount, source.simulationCount, 'Trusted rollback must reuse the source path count.');
    vNextAssertEqual_(rerun.previousRunId, source.runId, 'Trusted rollback must retain source-run lineage.');
    vNextAssertTrue_(rerun.lenses.rollback.sameSeed, 'Rollback audit declares same-seed comparison.');
    vNextAssertEqual_(rerun.lenses.rollback.sourceForecastRunId, source.runId, 'Rollback lens retains source run ID.');
    vNextAssertEqual_(vNextCanonicalJson_(rerun.annual), vNextCanonicalJson_(source.annual), 'Unchanged evidence with trusted rollback reproduces annual paths.');
    var changedNonAi = vNextClonePlain_(rollback);
    changedNonAi.humanEvents[0].amount = 2;
    vNextAssertThrows_(function () { vNextEngineRunForecast(changedNonAi); },
      'Non-AI inputs changed', 'Non-AI input changes must fail the AI-only comparable run.');
  } finally {
    vNextFindForecastByRunId_ = originalFind;
  }
}

function vNextTestAdminRunIdempotencyContract_() {
  var bookId = 'BOOK-IDEMPOTENCY';
  var key = 'FORECAST|BOOK-IDEMPOTENCY|REQUEST-1';
  var identity = vNextEngineBuildAdminRunIdentity_(bookId, key);
  var repeated = vNextEngineBuildAdminRunIdentity_(bookId, key);
  vNextAssertEqual_(identity.runId, repeated.runId, 'Admin run identity must be deterministic.');
  vNextAssertTrue_(identity.runId !== vNextEngineBuildAdminRunIdentity_(bookId, key + '-OTHER').runId,
    'Different idempotency keys must produce different run IDs.');
  vNextAssertThrows_(function () {
    vNextAuthorizeForecastRequest_({ runId: identity.runId, idempotencyKey: key });
  }, 'server-authorized ADMIN_JOB', 'Employee entry must reject server run identity fields.');
  vNextAssertThrows_(function () {
    vNextNormalizeRunRequest_({
      bookId: bookId, runId: identity.runId, idempotencyKey: key,
      internalOperation: 'ADMIN_JOB'
    });
  }, 'server authorization', 'Raw ADMIN_JOB text without the authorization marker must fail.');

  var request = vNextBuildMinimalRunRequest_();
  request.bookId = bookId;
  request.runId = identity.runId;
  request.idempotencyKey = key;
  request.internalOperation = 'ADMIN_JOB';
  request.serverRunIdentityAuthorized = true;
  request.authorizedRunIdentity = identity;
  var normalized = vNextNormalizeRunRequest_(request);
  vNextAssertEqual_(normalized.runId, identity.runId, 'Authorized deterministic run ID survives normalization.');
  vNextAssertTrue_(/^[a-f0-9]{64}$/.test(normalized.runIdentity.lineageHash),
    'Authorized run stores an immutable lineage hash.');
  var result = vNextSimulateForecast_(normalized);
  var record = vNextResultToForecastRecord_(result, normalized);
  var originalRead = vNextReadRecords_;
  try {
    vNextReadRecords_ = function (sheetName) {
      return sheetName === 'FORECAST_RUN' ? [record] : [];
    };
    var replay = vNextResolveExistingDeterministicRun_(normalized);
    vNextAssertTrue_(!!replay.result, 'Same run ID and canonical input resolves the stored SUCCESS.');
    var resume = vNextEngineLookupRunForResume_({
      internalOperation: 'ADMIN_JOB', bookId: bookId,
      runId: identity.runId, idempotencyKey: key,
      expectedInputDataHash: normalized.inputDataHash
    });
    vNextAssertTrue_(resume.hasSuccess && resume.resumablePhase === 'DRAFT_READY_SYNC',
      'Admin lookup exposes the phase-resume contract.');

    var changed = vNextClonePlain_(record);
    changed.input_data_hash = new Array(65).join('f');
    vNextReadRecords_ = function () { return [changed]; };
    vNextAssertThrows_(function () { vNextResolveExistingDeterministicRun_(normalized); },
      'different canonical input hash', 'Same run ID with different canonical input must fail closed.');

    vNextReadRecords_ = function () { return [record]; };
    var rollbackLineageRequest = vNextClonePlain_(request);
    rollbackLineageRequest.actualRecords = request.actualRecords;
    rollbackLineageRequest.internalJobType = 'AI_ROLLBACK_FORECAST';
    var rollbackLineage = vNextNormalizeRunRequest_(rollbackLineageRequest);
    vNextAssertTrue_(rollbackLineage.runIdentity.lineageHash !== normalized.runIdentity.lineageHash,
      'AI rollback uses the same identity contract with distinct immutable lineage.');
    vNextAssertThrows_(function () { vNextResolveExistingDeterministicRun_(rollbackLineage); },
      'immutable lineage', 'A normal run identity cannot be reused for AI rollback lineage.');
  } finally {
    vNextReadRecords_ = originalRead;
  }
}

function vNextBuildMinimalRunRequest_() {
  return {
    bookId: 'BOOK-HASH',
    clientId: 'CLIENT-HASH',
    clientName: 'テスト製薬',
    fiscalYear: 2025,
    asOf: '2025-02-01',
    actualRecords: vNextBuildTestActuals_(2019, 6, '2025-02-01'),
    seed: 445566,
    simulationCount: 400,
    persist: false,
    manageState: false,
    missingResponseRate: 0,
    informationGapRate: 0,
    evidenceResponseCounts: { change: 1, noChange: 0, unknown: 0 },
    commitmentEvents: [],
    objectiveEvents: [],
    humanEvents: [{ evidenceId: 'EV-1', ruleVersion: 'RULE-1', direction: 'UP', amount: 1, probability: 1 }],
    aiEvents: [],
    referencePrior: {}
  };
}

function vNextClonePlain_(value) {
  return JSON.parse(JSON.stringify(value));
}

function vNextBuildTestActuals_(firstFiscalYear, count, asOf) {
  var cutoff = vNextCutoffFromAsOf_(asOf);
  var records = [];
  for (var fy = firstFiscalYear; fy < firstFiscalYear + count; fy++) {
    for (var monthIndex = 0; monthIndex < 12; monthIndex++) {
      var date = new Date(fy, 3 + monthIndex, 15);
      if (date > cutoff) continue;
      var seasonal = [0.75, 0.80, 0.90, 0.95, 0.85, 1.05, 1.10, 1.00, 1.30, 0.95, 0.90, 1.45][monthIndex];
      records.push({
        client: 'テスト製薬',
        actualDate: date,
        dateSource: 'ACTUAL',
        isConfirmed: true,
        amount: Math.round((1000000 + (fy - firstFiscalYear) * 50000) * seasonal),
        serviceType: monthIndex % 5 === 0 ? 'SPOT' : 'BASE',
        product: 'TEST'
      });
    }
  }
  return records;
}

function vNextAssertTrue_(condition, message) {
  if (!condition) throw new Error(message || 'Expected condition to be true.');
  return true;
}

function vNextAssertEqual_(actual, expected, message) {
  if (actual !== expected) throw new Error((message || 'Values differ.') + ' actual=' + actual + ' expected=' + expected);
}

function vNextAssertNear_(actual, expected, tolerance, message) {
  if (Math.abs(Number(actual) - Number(expected)) > Number(tolerance || 0)) {
    throw new Error((message || 'Values are not near.') + ' actual=' + actual + ' expected=' + expected);
  }
}

function vNextAssertThrows_(operation, expectedText, message) {
  var thrown = null;
  try { operation(); } catch (error) { thrown = error; }
  if (!thrown) throw new Error((message || 'Expected error.') + ' No error was thrown.');
  if (expectedText && String(thrown.message || thrown).indexOf(expectedText) < 0) {
    throw new Error((message || 'Unexpected error.') + ' actual=' + String(thrown.message || thrown));
  }
}
