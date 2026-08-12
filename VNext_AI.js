/**
 * Forecast vNext Admin-only Vertex AI research provider.
 * This file must never be included in the Client Runtime bundle.
 */

var VNEXT_AI = Object.freeze({
  VERSION: 'vnext-ai-0.1.0',
  PROMPT_VERSION: 'vnext-ai-research-ja-1',
  SCHEMA_VERSION: 'vnext-ai-finding-1',
  RULE_VERSION: 'vnext-ai-impact-class-1',
  MAX_FINDINGS: 3,
  IMPACT_RATE_BY_CLASS: Object.freeze({ SMALL: 0.005, MEDIUM: 0.015, LARGE: 0.03 })
});

/** Admin job provider hook consumed by VNext_Admin.js. */
function vNextVertexAiResearch_(request) {
  var req = request && typeof request === 'object' ? request : {};
  // The production Hub runtime uses the active bound spreadsheet. During the
  // guarded pilot, the Admin-owned central source trigger opens the registered
  // Hub by ID, so there is intentionally no active Hub container. Only accept
  // the explicit in-memory Hub handle supplied by the trusted ADMIN_JOB worker.
  var trustedWorker = String(req.internalOperation || '') === 'ADMIN_JOB' && req.spreadsheet;
  if (trustedWorker) {
    if (typeof vNextDetectBookMode_ !== 'function' || vNextDetectBookMode_(req.spreadsheet) !== 'ADMIN' ||
        (typeof vNextAdminIsRegisteredHub_ === 'function' && !vNextAdminIsRegisteredHub_(req.spreadsheet))) {
      throw new Error('Vertex AI research requires the registered Admin Hub.');
    }
    if (typeof vNextAdminAssertHubAdmin_ === 'function') vNextAdminAssertHubAdmin_(req.spreadsheet, true);
    if (typeof vNextAdminHydrateHubRuntime_ === 'function') vNextAdminHydrateHubRuntime_(req.spreadsheet);
  } else {
    if (typeof vNextDetectBookMode_ === 'function' && vNextDetectBookMode_() !== 'ADMIN') {
      throw new Error('Vertex AI research can run only in the Admin Hub.');
    }
    if (typeof vNextAdminRequireHub_ === 'function') vNextAdminRequireHub_();
  }
  var config = vNextAiRuntimeConfig_();
  var clientName = String(req.clientName || '').trim();
  var bookId = String(req.bookId || '').trim();
  var fiscalYear = Number(req.fiscalYear);
  if (!clientName || !bookId || !isFinite(fiscalYear)) throw new Error('AI research requires bookId, clientName and fiscalYear.');

  var asOf = vNextParseDate_(req.asOf || new Date(), 'asOf');
  var asOfText = vNextFormatDateOnly_(asOf);
  var basisAmount = Number(req.basisAmount);
  if (!isFinite(basisAmount) || basisAmount <= 0) {
    var actuals = vNextFetchActualRecordsBridge_(clientName, { fiscalYear: fiscalYear, asOf: asOf });
    var prior = vNextBuildContinuityPrior_(actuals, fiscalYear, vNextCutoffFromAsOf_(asOf));
    basisAmount = Number(prior.annualBaseline || prior.baseAnnualBaseline || 0);
  }
  if (!isFinite(basisAmount) || basisAmount <= 0) throw new Error('AI research could not derive a positive historical basis.');

  var researchPrompt = vNextAiResearchPrompt_(clientName, fiscalYear, asOfText);
  var grounded = callVertexGeminiGrounded_(researchPrompt, { config: config });
  if (!grounded || !grounded.ok || !grounded.text) {
    throw new Error('Vertex grounded research failed: ' + String(grounded && grounded.error || 'empty response'));
  }
  var citations = extractGeminiGroundingCitations_(grounded.raw).filter(function (item) {
    return /^https?:\/\//i.test(String(item.uri || ''));
  }).slice(0, 12);
  if (!citations.length) throw new Error('Vertex research returned no usable citation URL; AI impact was not applied.');

  var structurePrompt = vNextAiStructurePrompt_(clientName, fiscalYear, asOfText, grounded.text, citations);
  var structured = callVertexGeminiStructured_(
    'You convert cited research into conservative Japanese sales-planning evidence. Return JSON only and never invent a citation.',
    structurePrompt,
    { config: config }
  );
  if (!structured || !structured.ok || !structured.json) {
    vNextAiWriteRawAudit_(req, researchPrompt, grounded, structurePrompt, structured, citations, []);
    throw new Error('Vertex finding structure failed: ' + String(structured && structured.error || 'empty response'));
  }

  var rawFindings = Array.isArray(structured.json.findings) ? structured.json.findings : [];
  var findings = vNextAiNormalizeFindings_(rawFindings, citations, {
    bookId: bookId,
    fiscalYear: fiscalYear,
    basisAmount: basisAmount,
    asOf: asOf,
    aiModel: config.geminiModel,
    parentRequestId: String(req.parentRequestId || ''),
    effectiveAsOf: asOfText
  });
  vNextAiWriteRawAudit_(req, researchPrompt, grounded, structurePrompt, structured, citations, findings);
  return findings;
}

function vNextAiRuntimeConfig_() {
  var props = PropertiesService.getScriptProperties();
  var config = {
    projectId: String(props.getProperty('VERTEX_PROJECT_ID') || '').trim(),
    location: String(props.getProperty('VERTEX_LOCATION') || '').trim(),
    geminiModel: String(props.getProperty('VERTEX_GEMINI_MODEL') || '').trim(),
    datastoreId: String(props.getProperty('VERTEX_DATASTORE_ID') || '').trim(),
    searchLocation: String(props.getProperty('VERTEX_SEARCH_LOCATION') || '').trim(),
    servingConfig: String(props.getProperty('VERTEX_SERVING_CONFIG') || 'default_search').trim()
  };
  if (!config.projectId || !config.location || !config.geminiModel) {
    throw new Error('Vertex AI runtime is not configured in Admin Hub Script Properties.');
  }
  config.geminiReady = true;
  config.ragReady = Boolean(config.datastoreId && config.searchLocation);
  return config;
}

function vNextAiResearchPrompt_(clientName, fiscalYear, asOfText) {
  return [
    'あなたは法人向け年度売上計画の外部情報調査担当です。',
    '対象クライアント: ' + clientName,
    '対象年度: FY' + fiscalYear + '（' + fiscalYear + '年4月〜' + (fiscalYear + 1) + '年3月）',
    '情報締切: ' + asOfText,
    '公開情報を検索し、この会社との取引売上を通常状態から変え得る外部要因だけを最大3件報告してください。',
    '対象は、顧客の事業方針・製品/市場・組織・競争・規制などです。一般論や根拠のない推測は除外してください。',
    '各要因について、方向、対象期間、短い根拠、出典を示してください。売上額や率を推測しないでください。'
  ].join('\n');
}

function vNextAiStructurePrompt_(clientName, fiscalYear, asOfText, researchText, citations) {
  var citationText = citations.map(function (item, index) {
    return '[' + index + '] ' + String(item.title || '') + ' | ' + String(item.uri || '');
  }).join('\n');
  return [
    '対象=' + clientName + ', FY=' + fiscalYear + ', as_of=' + asOfText,
    '次の調査文と引用一覧だけを使ってください。',
    'JSON schema:',
    '{"findings":[{"target":"短い対象名","direction":"UP|DOWN","impactClass":"SMALL|MEDIUM|LARGE","confidenceClass":"CONFIRMED_FACT|LIKELY|HYPOTHESIS","evidenceQuality":"A|B|C|D","summary":"日本語140字以内","citationIndex":0,"sourceDate":"YYYY-MM-DD","startMonth":"YYYY-MM","endMonth":"YYYY-MM"}]}',
    'impactClassは売上額ではなく外部変化の重要度。通常はSMALL、明確な全社/主力事業級のみMEDIUM、LARGEは極めて例外的。',
    'citationIndexは必ず一覧に存在する番号。公開日を確認できない項目、根拠が弱い項目は出力しない。最大3件。',
    '対象期間はFY' + fiscalYear + 'の範囲内。',
    '',
    '調査文:',
    String(researchText || '').slice(0, 16000),
    '',
    '引用一覧:',
    citationText
  ].join('\n');
}

function vNextAiNormalizeFindings_(rawFindings, citations, context) {
  var fy = Number(context.fiscalYear);
  var fyStart = fy + '-04';
  var fyEnd = (fy + 1) + '-03';
  var expiry = new Date(context.asOf.getTime());
  expiry.setDate(expiry.getDate() + 90);
  return (rawFindings || []).slice(0, VNEXT_AI.MAX_FINDINGS).map(function (row) {
    var citationIndex = Number(row.citationIndex);
    var citation = isFinite(citationIndex) ? citations[Math.floor(citationIndex)] : null;
    if (!citation || !/^https?:\/\//i.test(String(citation.uri || ''))) return null;
    var direction = String(row.direction || '').toUpperCase();
    if (direction !== 'UP' && direction !== 'DOWN') return null;
    var impactClass = String(row.impactClass || 'SMALL').toUpperCase();
    if (!VNEXT_AI.IMPACT_RATE_BY_CLASS[impactClass]) impactClass = 'SMALL';
    var confidence = String(row.confidenceClass || 'HYPOTHESIS').toUpperCase();
    if (['CONFIRMED_FACT', 'LIKELY', 'HYPOTHESIS'].indexOf(confidence) < 0) confidence = 'HYPOTHESIS';
    var quality = String(row.evidenceQuality || 'C').toUpperCase();
    if (['A', 'B', 'C', 'D'].indexOf(quality) < 0) quality = 'C';
    var startMonth = /^\d{4}-\d{2}$/.test(String(row.startMonth || '')) ? String(row.startMonth) : fyStart;
    var endMonth = /^\d{4}-\d{2}$/.test(String(row.endMonth || '')) ? String(row.endMonth) : fyEnd;
    if (startMonth < fyStart || startMonth > fyEnd) startMonth = fyStart;
    if (endMonth < startMonth || endMonth > fyEnd) endMonth = fyEnd;
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(row.sourceDate || ''))) return null;
    var sourceDate = String(row.sourceDate);
    var parsedSourceDate = new Date(sourceDate + 'T00:00:00');
    if (isNaN(parsedSourceDate.getTime()) || parsedSourceDate.getTime() > context.asOf.getTime()) return null;
    var rate = VNEXT_AI.IMPACT_RATE_BY_CLASS[impactClass] * (direction === 'DOWN' ? -1 : 1);
    return {
      bookId: context.bookId,
      target: String(row.target || '外部環境変化').trim().slice(0, 120),
      targetStartMonth: startMonth,
      targetEndMonth: endMonth,
      effectRate: rate,
      basisAmount: context.basisAmount,
      sourceUrl: String(citation.uri),
      citationTitle: String(citation.title || '').slice(0, 240),
      sourceDate: sourceDate,
      expiresAt: vNextFormatDateOnly_(expiry),
      summary: String(row.summary || row.target || '').replace(/[\u0000-\u001f]+/g, ' ').trim().slice(0, 280),
      confidenceClass: confidence,
      evidenceQuality: quality,
      aiModel: context.aiModel,
      promptVersion: VNEXT_AI.PROMPT_VERSION,
      aiSchemaVersion: VNEXT_AI.SCHEMA_VERSION,
      ruleVersion: VNEXT_AI.RULE_VERSION,
      parentRequestId: String(context.parentRequestId || ''),
      effectiveAsOf: String(context.effectiveAsOf || vNextFormatDateOnly_(context.asOf))
    };
  }).filter(Boolean);
}

function vNextAiWriteRawAudit_(request, researchPrompt, grounded, structurePrompt, structured, citations, findings) {
  try {
    var hub = SpreadsheetApp.getActiveSpreadsheet();
    var hubFile = DriveApp.getFileById(hub.getId());
    var parents = hubFile.getParents();
    var parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
    var folders = parent.getFoldersByName('Forecast vNext Admin Audit');
    var folder = folders.hasNext() ? folders.next() : parent.createFolder('Forecast vNext Admin Audit');
    var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
    var safeBook = String(request.bookId || 'BOOK').replace(/[^A-Za-z0-9_-]/g, '_').slice(0, 80);
    var payload = {
      auditVersion: VNEXT_AI.VERSION,
      createdAt: new Date().toISOString(),
      bookId: String(request.bookId || ''),
      clientName: String(request.clientName || ''),
      fiscalYear: Number(request.fiscalYear || 0),
      researchPrompt: researchPrompt,
      grounded: grounded,
      structurePrompt: structurePrompt,
      structured: structured,
      citations: citations,
      normalizedFindings: findings
    };
    folder.createFile('AI_' + safeBook + '_' + stamp + '_' + Utilities.getUuid() + '.json', JSON.stringify(payload), 'application/json');
  } catch (error) {
    Logger.log('vNext AI raw audit write failed: %s', String(error && error.stack || error));
  }
}

/** Pure deterministic test; no Vertex or spreadsheet access. */
function testVNextAiDeterministicMapping() {
  var findings = vNextAiNormalizeFindings_([{
    target: '制度変更', direction: 'DOWN', impactClass: 'MEDIUM', confidenceClass: 'LIKELY',
    evidenceQuality: 'B', summary: '根拠要約', citationIndex: 0,
    sourceDate: '2027-03-01', startMonth: '2027-04', endMonth: '2028-03'
  }], [{ title: '一次情報', uri: 'https://example.com/source' }], {
    bookId: 'BOOK-1', fiscalYear: 2027, basisAmount: 100000000,
    asOf: new Date('2027-03-15T00:00:00Z'), aiModel: 'model'
  });
  if (findings.length !== 1 || findings[0].effectRate !== -0.015) throw new Error('AI impact class mapping failed.');
  if (Object.prototype.hasOwnProperty.call(findings[0], 'researchPrompt')) throw new Error('Raw prompt leaked into finding.');
  return true;
}
