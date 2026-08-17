/**
 * Forecast vNext Admin-only Vertex AI research provider.
 * This file must never be included in the Client Runtime bundle.
 */

var VNEXT_AI = Object.freeze({
  VERSION: 'vnext-ai-0.2.0',
  PROMPT_VERSION: 'vnext-ai-research-ja-2',
  SCHEMA_VERSION: 'vnext-ai-finding-2',
  RULE_VERSION: 'vnext-ai-evidence-gate-2',
  MAX_FINDINGS: 6,
  IMPACT_RATE_BY_CLASS: Object.freeze({ SMALL: 0.0025, MEDIUM: 0.0075, LARGE: 0.015 }),
  AXES: Object.freeze([
    'FINANCIAL_CAPACITY', 'DIGITAL_EXECUTION', 'PRODUCT_MARKET',
    'STRATEGY_ORGANIZATION', 'REGULATORY_SUPPLY', 'ALTERNATIVE_SIGNALS'
  ]),
  SOURCE_STRENGTHS: Object.freeze([
    'PRIMARY_OFFICIAL', 'PRIMARY_REGISTRY', 'REPUTABLE_SECONDARY', 'ALTERNATIVE_SIGNAL'
  ])
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
    vNextAiWriteRawAudit_(req, researchPrompt, grounded, '', null, [], []);
    throw new Error('Vertex grounded research failed: ' + String(grounded && grounded.error || 'empty response'));
  }
  var citations = vNextAiExtractCitations_(grounded.raw).filter(function (item) {
    return /^https?:\/\//i.test(String(item.uri || ''));
  }).slice(0, 12);
  if (!citations.length && config.ragReady && typeof callVertexSearchRAG_ === 'function') {
    var rag = callVertexSearchRAG_(vNextAiRagQuery_(clientName, fiscalYear, asOfText), { config: config });
    grounded = Object.assign({}, grounded, { ragFallback: rag });
    if (rag && rag.ok) {
      citations = vNextAiDedupeCitations_((rag.citations || []).concat(rag.documents || [])).filter(function (item) {
        return /^https?:\/\//i.test(String(item.uri || ''));
      }).slice(0, 12);
      if (rag.summary) grounded.text += '\n\nVertex Search補足:\n' + String(rag.summary);
    }
  }
  if (!citations.length) {
    vNextAiWriteRawAudit_(req, researchPrompt, grounded, '', null, [], []);
    throw new Error('Vertex research returned no usable citation URL after grounded search and configured search fallback; AI impact was not applied.');
  }

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
  var names = vNextAiClientSearchNames_(clientName);
  return [
    'あなたは法人向け年度売上計画の外部情報リサーチャーです。担当者が過去売上や社内意見だけでは見落としやすい変化を見つけます。',
    '対象クライアント: ' + clientName + '（検索表記候補: ' + names.join(' / ') + '。正式名・英語名・略称も確認）',
    '対象年度: FY' + fiscalYear + '（' + fiscalYear + '年4月〜' + (fiscalYear + 1) + '年3月）',
    '情報締切: ' + asOfText,
    '次の6観点を横断し、重要な事実または先行シグナルを最大6件報告してください。',
    '1. 業績・資金・設備投資・研究開発などの投資余力',
    '2. DX・データ・AI・業務変革が構想でなく実行段階にある証拠',
    '3. 製品・サービス・市場の勢い、上市、採用、成長/失速',
    '4. 中期戦略、提携、買収、経営/組織/購買責任の変化',
    '5. 規制、供給、調達、入札、サプライチェーンの変化',
    '6. 採用増減、特許・治験・認証、施設投資、提携網、公開情報更新頻度などの代替的な先行シグナル',
    '会社公式発表・法定開示・規制/公的登録・調達公告を優先し、信頼できる二次情報と代替シグナルは区別してください。',
    'この取引売上へのつながりが明示できない情報も、担当者が確認すべき参考示唆として残して構いません。',
    '売上額や率を推測しないでください。事実、解釈、担当者への確認質問を分け、各項目に出典URLと公開日を付けてください。'
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
    '{"findings":[{"axis":"FINANCIAL_CAPACITY|DIGITAL_EXECUTION|PRODUCT_MARKET|STRATEGY_ORGANIZATION|REGULATORY_SUPPLY|ALTERNATIVE_SIGNALS","signalType":"短い分類","target":"短い対象名","direction":"UP|DOWN|NEUTRAL","forecastUse":"APPLY|INSIGHT_ONLY","impactClass":"NONE|SMALL|MEDIUM|LARGE","confidenceClass":"CONFIRMED_FACT|LIKELY|HYPOTHESIS","evidenceQuality":"A|B|C|D","sourceStrength":"PRIMARY_OFFICIAL|PRIMARY_REGISTRY|REPUTABLE_SECONDARY|ALTERNATIVE_SIGNAL","salesRelevance":"HIGH|MEDIUM|LOW","summary":"確認できた事実と意味を日本語180字以内","humanQuestion":"担当者が次に確認する質問を日本語100字以内","citationIndex":0,"sourceDate":"YYYY-MM-DD","startMonth":"YYYY-MM","endMonth":"YYYY-MM"}]}',
    'forecastUse=APPLYは、一次情報または公的登録により、対象年度の取引売上との直接的なつながりを説明できる場合だけ。その他はINSIGHT_ONLY。',
    'impactClassは売上額ではなく外部変化の重要度。通常はNONEまたはSMALL。LARGEは例外的で、サーバ側でさらに縮小・棄却されます。',
    '代替シグナルと仮説は担当者向けのINSIGHT_ONLYにし、予測金額へ直接反映しないでください。',
    'citationIndexは必ず一覧に存在する番号。公開日を確認できない項目は出力しない。6観点を優先度順に最大6件。',
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
    var axis = String(row.axis || 'ALTERNATIVE_SIGNALS').toUpperCase();
    if (VNEXT_AI.AXES.indexOf(axis) < 0) axis = 'ALTERNATIVE_SIGNALS';
    var direction = String(row.direction || 'NEUTRAL').toUpperCase();
    if (['UP', 'DOWN', 'NEUTRAL'].indexOf(direction) < 0) direction = 'NEUTRAL';
    var impactClass = String(row.impactClass || 'SMALL').toUpperCase();
    if (impactClass !== 'NONE' && !VNEXT_AI.IMPACT_RATE_BY_CLASS[impactClass]) impactClass = 'SMALL';
    var confidence = String(row.confidenceClass || 'HYPOTHESIS').toUpperCase();
    if (['CONFIRMED_FACT', 'LIKELY', 'HYPOTHESIS'].indexOf(confidence) < 0) confidence = 'HYPOTHESIS';
    var quality = String(row.evidenceQuality || 'C').toUpperCase();
    if (['A', 'B', 'C', 'D'].indexOf(quality) < 0) quality = 'C';
    var sourceStrength = String(row.sourceStrength || 'REPUTABLE_SECONDARY').toUpperCase();
    if (VNEXT_AI.SOURCE_STRENGTHS.indexOf(sourceStrength) < 0) sourceStrength = 'REPUTABLE_SECONDARY';
    var salesRelevance = String(row.salesRelevance || 'LOW').toUpperCase();
    if (['HIGH', 'MEDIUM', 'LOW'].indexOf(salesRelevance) < 0) salesRelevance = 'LOW';
    var requestedUse = String(row.forecastUse || 'INSIGHT_ONLY').toUpperCase();
    var forecastUse = requestedUse === 'APPLY' && impactClass !== 'NONE' && (direction === 'UP' || direction === 'DOWN') &&
      (sourceStrength === 'PRIMARY_OFFICIAL' || sourceStrength === 'PRIMARY_REGISTRY') &&
      (quality === 'A' || quality === 'B') &&
      (confidence === 'CONFIRMED_FACT' || confidence === 'LIKELY') && salesRelevance === 'HIGH'
      ? 'APPLY' : 'INSIGHT_ONLY';
    var startMonth = /^\d{4}-\d{2}$/.test(String(row.startMonth || '')) ? String(row.startMonth) : fyStart;
    var endMonth = /^\d{4}-\d{2}$/.test(String(row.endMonth || '')) ? String(row.endMonth) : fyEnd;
    if (startMonth < fyStart || startMonth > fyEnd) startMonth = fyStart;
    if (endMonth < startMonth || endMonth > fyEnd) endMonth = fyEnd;
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(row.sourceDate || ''))) return null;
    var sourceDate = String(row.sourceDate);
    var parsedSourceDate = new Date(sourceDate + 'T00:00:00');
    if (isNaN(parsedSourceDate.getTime()) || parsedSourceDate.getTime() > context.asOf.getTime()) return null;
    var rate = forecastUse === 'APPLY'
      ? VNEXT_AI.IMPACT_RATE_BY_CLASS[impactClass === 'NONE' ? 'SMALL' : impactClass] * (direction === 'DOWN' ? -1 : 1)
      : 0;
    return {
      bookId: context.bookId,
      target: String(row.target || '外部環境変化').trim().slice(0, 120),
      targetStartMonth: startMonth,
      targetEndMonth: endMonth,
      effectRate: rate,
      direction: direction,
      basisAmount: context.basisAmount,
      sourceUrl: String(citation.uri),
      citationTitle: String(citation.title || '').slice(0, 240),
      sourceDate: sourceDate,
      expiresAt: vNextFormatDateOnly_(expiry),
      summary: String(row.summary || row.target || '').replace(/[\u0000-\u001f]+/g, ' ').trim().slice(0, 280),
      confidenceClass: confidence,
      evidenceQuality: quality,
      researchAxis: axis,
      signalType: String(row.signalType || '').replace(/[\u0000-\u001f]+/g, ' ').trim().slice(0, 80),
      sourceStrength: sourceStrength,
      forecastUse: forecastUse,
      salesRelevance: salesRelevance,
      humanQuestion: String(row.humanQuestion || '').replace(/[\u0000-\u001f]+/g, ' ').trim().slice(0, 180),
      aiModel: context.aiModel,
      promptVersion: VNEXT_AI.PROMPT_VERSION,
      aiSchemaVersion: VNEXT_AI.SCHEMA_VERSION,
      ruleVersion: VNEXT_AI.RULE_VERSION,
      parentRequestId: String(context.parentRequestId || ''),
      effectiveAsOf: String(context.effectiveAsOf || vNextFormatDateOnly_(context.asOf))
    };
  }).filter(Boolean);
}

function vNextAiClientSearchNames_(clientName) {
  var raw = String(clientName || '').trim();
  var normalized = typeof raw.normalize === 'function' ? raw.normalize('NFKC') : raw;
  var bare = normalized.replace(/(?:株式会社|有限会社|合同会社|\(株\)|（株）|\(有\)|（有）)/g, '').trim();
  return [raw, normalized, bare].filter(function (value, index, list) {
    return value && list.indexOf(value) === index;
  });
}

function vNextAiRagQuery_(clientName, fiscalYear, asOfText) {
  return vNextAiClientSearchNames_(clientName).join(' ') + ' FY' + fiscalYear +
    ' 業績 投資 DX 製品 戦略 規制 調達 採用 提携 as of ' + asOfText;
}

function vNextAiExtractCitations_(json) {
  var out = [];
  if (typeof extractGeminiGroundingCitations_ === 'function') {
    out = out.concat(extractGeminiGroundingCitations_(json) || []);
  }
  ((json && json.candidates) || []).forEach(function (candidate) {
    var citationMetadata = candidate && candidate.citationMetadata || {};
    (citationMetadata.citations || []).forEach(function (citation) {
      out.push({ title: String(citation.title || citation.license || '').trim(), uri: String(citation.uri || '').trim() });
    });
    var grounding = candidate && candidate.groundingMetadata || {};
    (grounding.groundingChunks || []).forEach(function (chunk) {
      var source = chunk && (chunk.web || chunk.retrievedContext) || {};
      out.push({ title: String(source.title || '').trim(), uri: String(source.uri || '').trim() });
    });
    var urlContext = candidate && candidate.urlContextMetadata || {};
    (urlContext.urlMetadata || []).forEach(function (item) {
      out.push({ title: String(item.title || '').trim(), uri: String(item.retrievedUrl || item.url || '').trim() });
    });
  });
  return vNextAiDedupeCitations_(out);
}

function vNextAiDedupeCitations_(items) {
  var seen = {};
  return (items || []).map(function (item) {
    return { title: String(item && item.title || '').trim(), uri: String(item && (item.uri || item.link) || '').trim() };
  }).filter(function (item) {
    if (!item.uri || seen[item.uri]) return false;
    seen[item.uri] = true;
    return true;
  });
}

function vNextAiWriteRawAudit_(request, researchPrompt, grounded, structurePrompt, structured, citations, findings) {
  try {
    var hub = request && request.spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    if (!hub || typeof hub.getId !== 'function') throw new Error('Admin Hub spreadsheet is unavailable for AI audit.');
    var destFolder = null;
    try {
      if (typeof vNextAdminReadKeyValueSheet_ === 'function') {
        var config = vNextAdminReadKeyValueSheet_(hub, 'VN_SYSTEM_CONFIG');
        if (config && config.library_audit_folder_id) {
          destFolder = DriveApp.getFolderById(String(config.library_audit_folder_id));
        }
      }
    } catch (configError) {
      destFolder = null;
    }
    if (!destFolder) {
      var hubFile = DriveApp.getFileById(hub.getId());
      var parents = hubFile.getParents();
      var parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
      var folders = parent.getFoldersByName('Forecast vNext Admin Audit');
      destFolder = folders.hasNext() ? folders.next() : parent.createFolder('Forecast vNext Admin Audit');
    }
    var folder = destFolder;
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
    axis: 'REGULATORY_SUPPLY', sourceStrength: 'PRIMARY_REGISTRY', salesRelevance: 'HIGH',
    forecastUse: 'APPLY', target: '制度変更', direction: 'DOWN', impactClass: 'MEDIUM', confidenceClass: 'LIKELY',
    evidenceQuality: 'B', summary: '根拠要約', humanQuestion: '制度対応時期を確認する', citationIndex: 0,
    sourceDate: '2027-03-01', startMonth: '2027-04', endMonth: '2028-03'
  }], [{ title: '一次情報', uri: 'https://example.com/source' }], {
    bookId: 'BOOK-1', fiscalYear: 2027, basisAmount: 100000000,
    asOf: new Date('2027-03-15T00:00:00Z'), aiModel: 'model'
  });
  if (findings.length !== 1 || findings[0].effectRate !== -0.0075 || findings[0].forecastUse !== 'APPLY') {
    throw new Error('AI impact evidence gate failed.');
  }
  var insight = vNextAiNormalizeFindings_([{
    axis: 'ALTERNATIVE_SIGNALS', sourceStrength: 'ALTERNATIVE_SIGNAL', salesRelevance: 'MEDIUM',
    forecastUse: 'APPLY', target: '採用増加', direction: 'UP', impactClass: 'LARGE', confidenceClass: 'HYPOTHESIS',
    evidenceQuality: 'C', summary: '採用情報から投資領域の変化が示唆される', citationIndex: 0,
    sourceDate: '2027-03-01', startMonth: '2027-04', endMonth: '2028-03'
  }], [{ title: '採用情報', uri: 'https://example.com/jobs' }], {
    bookId: 'BOOK-1', fiscalYear: 2027, basisAmount: 100000000,
    asOf: new Date('2027-03-15T00:00:00Z'), aiModel: 'model'
  });
  if (insight.length !== 1 || insight[0].effectRate !== 0 || insight[0].forecastUse !== 'INSIGHT_ONLY') {
    throw new Error('Alternative signal must remain insight-only.');
  }
  var citations = vNextAiExtractCitations_({ candidates: [{
    citationMetadata: { citations: [{ uri: 'https://example.com/citation', title: '引用' }] },
    groundingMetadata: { groundingChunks: [{ retrievedContext: { uri: 'https://example.com/rag', title: 'RAG' } }] }
  }] });
  if (citations.length !== 2) throw new Error('AI citation variants were not extracted.');
  if (Object.prototype.hasOwnProperty.call(findings[0], 'researchPrompt')) throw new Error('Raw prompt leaked into finding.');
  return true;
}
