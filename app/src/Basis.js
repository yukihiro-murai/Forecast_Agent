/**
 * Basis.js — 予測の根拠（予測の画面の「根拠」タブ）。
 *
 * 旧来の予測（A-9）が毎回残す記録を、データ本体から読んで並べる（計算はしない。旧来の計算は変えない）:
 *   - 月ごとの P50: 過去の売上だけの予測（客観）と、入力・AI・スポットを入れた最終（混合）。その差
 *   - 入力（製品・クライアント・見解）・AI の話題・Vertex の補正が、月ごとにどちらへ・どれだけ押したか（SUBJECTIVE_IMPACT_HISTORY）
 *   - AI の倍率と点数（AI_IMPACT_HISTORY）・確定しているスポット（FORECAST_SNAPSHOT の deterministic_adj）
 *   - 補正（CALIBRATION_STATE・FORECAST_SNAPSHOT の calibration_applied_json）
 *   - AI 調査の根拠（AI_RESEARCH_STRUCTURED: 話題・向き・点数・確からしさ・根拠の文）・Vertex の説明
 *   - 前回の予測からの変化（月ごとの P50）と、その間にあった操作
 * 入力のハッシュが同じなら、組み立てた結果を 6 時間控える。
 */
const APP_BASIS_SHEETS = ['FORECAST_SNAPSHOT', 'AI_IMPACT_HISTORY', 'SUBJECTIVE_IMPACT_HISTORY', 'AI_RESEARCH_STRUCTURED', 'AI_SCORE_HISTORY',
  'CALIBRATION_STATE', 'VERTEX_FORECAST_LOG'];
const APP_SOURCE_LABELS = { factor_product: '製品の入力', factor_client: 'クライアント全体の入力', opinion: '見解', ai_topic: 'AI の話題', vertex_forecast: 'Vertex の補正' };

/** データ本体の表の形のシートを、見出し → 値の行にする（{ シート名: [行] }） */
function appEngTableObjects_(planId, names) {
  const sheets = appEngLoadPlanSheets_(planId, names);
  const out = {};
  names.forEach(n => {
    const s = sheets[n];
    if (!s || !s.values.length) { out[n] = []; return; }
    const head = s.values[0].map(String);
    out[n] = s.values.slice(1).filter(r => r.some(v => v !== '' && v !== null)).map(r => {
      const o = {};
      head.forEach((h, j) => { if (h) o[h] = r[j]; });
      return o;
    });
  });
  return out;
}

function appYm_(v) {
  if (appIsDate_(v)) return Utilities.formatDate(v, APP_TZ, 'yyyy/MM');
  const m = /^(\d{4})[-/](\d{1,2})/.exec(String(v || ''));
  return m ? m[1] + '/' + ('0' + m[2]).slice(-2) : String(v || '');
}

function appTimeKey_(v) { return appIsDate_(v) ? v.getTime() : new Date(String(v)).getTime() || 0; }

/** 表の中で、一番新しい回（key の列が最大の値）の行だけ */
function appLatestBatch_(rows, key) {
  let max = null;
  rows.forEach(r => { const t = appTimeKey_(r[key]); if (max === null || t > max) max = t; });
  return max === null ? [] : rows.filter(r => appTimeKey_(r[key]) === max);
}

function appNum_(v) { const n = typeof v === 'number' ? v : Number(v); return v === '' || v === null || !isFinite(n) ? null : n; }

/** 画面: 計画の予測の根拠 */
function appForecastBasis_(planId) {
  const plan = appPlanOf_(planId);
  const inputHash = appPlanInputHash_(plan.plan_id);
  const key = 'BASIS_' + appSha256Hex_([plan.plan_id, inputHash, APP_VERSION].join('|')).slice(0, 32);
  const cached = appJobGetResult_(key);
  if (cached.found) return cached.value;
  const t = appEngTableObjects_(plan.plan_id, APP_BASIS_SHEETS);
  // 新アプリで動かした予測（新しい順）と、その月ごとの P50
  const runs = appReadTable_('FORECAST_RUNS').filter(r => r.plan_id === plan.plan_id && r.status === 'DONE')
    .sort((a, b) => String(b.finished_at).localeCompare(String(a.finished_at)) || b._row - a._row);
  const monthlyOf = runId => {
    const m = {};
    if (runId) appReadTable_('FORECAST_MONTHLY').filter(x => x.run_id === runId).forEach(x => { m[appYm_(x.ym)] = x; });
    return m;
  };
  const latest = runs[0] || null, prev = runs[1] || null;
  const cur = monthlyOf(latest && latest.run_id), before = monthlyOf(prev && prev.run_id);
  const stored = appStoredHeadline_(plan.plan_id);
  const months = (stored ? stored.monthly.map(m => m.month) : Object.keys(cur).sort()).filter(Boolean);

  // 旧来の予測が残した、一番新しい回の記録
  const snap = appLatestBatch_(t.FORECAST_SNAPSHOT, 'run_date').filter(r => String(r.scenario) === 'neutral');
  const ai = appLatestBatch_(t.AI_IMPACT_HISTORY, 'run_at');
  const push = appLatestBatch_(t.SUBJECTIVE_IMPACT_HISTORY, 'run_at');
  const snapBy = {}, aiBy = {}, pushBy = {};
  snap.forEach(r => { snapBy[appYm_(r.target_month)] = r; });
  ai.forEach(r => { aiBy[appYm_(r.target_month)] = r; });
  push.forEach(r => { (pushBy[appYm_(r.target_month)] = pushBy[appYm_(r.target_month)] || []).push(r); });

  const monthly = months.map(ym => {
    const c = cur[ym] || {}, b = before[ym] || {}, s = snapBy[ym] || {}, a = aiBy[ym] || {};
    const sm = stored ? stored.monthly.filter(m => m.month === ym)[0] || {} : {};
    const p50 = appNum_(c.p50) !== null ? appNum_(c.p50) : appNum_(sm.p50);
    const obj = appNum_(c.obj_p50);
    const sources = {};
    (pushBy[ym] || []).forEach(r => {
      const k = String(r.source_type || '');
      const st = appNum_(r.push_step);
      if (st === null) return;
      const o = sources[k] = sources[k] || { type: k, label: APP_SOURCE_LABELS[k] || k, step: 0, n: 0, keys: [] };
      o.step += st; o.n++;
      if (o.keys.length < 5 && r.source_key) o.keys.push(String(r.source_key) + ' ' + (st > 0 ? '+' : '') + Math.round(st * 1000) / 10 + '%' + (appNum_(r.applied_reliability_r) !== null && appNum_(r.applied_reliability_r) !== 1 ? '（信頼度 ' + appNum_(r.applied_reliability_r).toFixed(2) + '）' : ''));
    });
    let opinion = '';
    try { opinion = (JSON.parse(String(s.key_factors_json || '{}')) || {}).opinion || ''; } catch (e) { opinion = ''; }
    return {
      month: ym, p50: p50, objective: obj, diff: p50 !== null && obj !== null ? p50 - obj : null,
      prevP50: appNum_(b.p50), change: p50 !== null && appNum_(b.p50) !== null ? p50 - appNum_(b.p50) : null,
      spot: appNum_(s.deterministic_adj), kAi: appNum_(a.k_ai), aiScore: appNum_(a.ai_total_score), aiDirection: String(a.ai_direction || ''),
      aiNeutralized: String(a.ai_neutralized || '') === '1' || a.ai_neutralized === true, source: String(a.forecast_source || s.forecast_source || ''),
      sources: Object.keys(sources).map(k => sources[k]), opinion: opinion
    };
  });

  // 補正（今の状態と、最後の予測で使ったもの）
  const cal = t.CALIBRATION_STATE[t.CALIBRATION_STATE.length - 1] || null;
  let applied = null;
  try { applied = snap.length ? JSON.parse(String(snap[0].calibration_applied_json || 'null')) : null; } catch (e) { applied = null; }
  // AI 調査の根拠（一番新しい日付の行）
  const research = appLatestBatch_(t.AI_RESEARCH_STRUCTURED, 'as_of_date').map(r => ({
    topic: String(r.topic || ''), rowType: String(r.row_type || ''), direction: String(r.direction || ''), impact: appNum_(r.impact_score),
    confidence: appNum_(r.confidence), score: appNum_(r.blended_score), horizon: String(r.time_horizon || ''),
    evidence: String(r.evidence || '').slice(0, 600), reason: String(r.business_relevance_reason || '').slice(0, 300), asOf: appYm_(r.as_of_date)
  }));
  const vertex = t.VERTEX_FORECAST_LOG.slice().sort((a, b) => appTimeKey_(b.run_at) - appTimeKey_(a.run_at))[0] || null;
  // 前回の予測から今回の予測までにあった操作（入力の保存・取り込み・AI 調査・学習など）
  const between = prev && latest ? appReadTable_('PLAN_ACTIONS').filter(x => x.plan_id === plan.plan_id && x.finished_at > prev.finished_at && x.finished_at <= latest.started_at)
    .map(x => ({ action: x.action, label: (APP_PLAN_ACTIONS[x.action] || {}).label || x.action, at: x.finished_at, actor: x.actor_email, changed: appParseJsonList_(x.changed_sheets_json) })) : [];
  const sum = k => monthly.reduce((s, m) => (m[k] === null ? s : (s || 0) + m[k]), null);
  const out = {
    planId: plan.plan_id, latestRunAt: latest ? latest.finished_at : '', prevRunAt: prev ? prev.finished_at : '',
    annual: { p50: latest ? latest.annual_p50 : stored && stored.annual.p50, objective: latest ? latest.objective_p50 : stored && stored.objective.p50,
      prevP50: prev ? prev.annual_p50 : null, spot: sum('spot') },
    monthly: monthly, between: between, inputChanged: between.length > 0,   // 予測の間に計画への操作があったか（無ければ、変化は乱数の揺れ）
    calibration: cal ? { factor: appNum_(cal.bias_correction_factor), aiWeight: appNum_(cal.ai_weight_override), aiMax: appNum_(cal.ai_max_abs_effect_override),
      monthBias: (() => { try { return JSON.parse(String(cal.residual_month_bias_json || '{}')); } catch (e) { return {}; } })(), updatedAt: String(cal.updated_at || ''),
      quarter: String(cal.last_applied_quarter || '') } : null,
    applied: applied, research: research,
    vertex: vertex ? { at: String(vertex.run_at || ''), confidence: appNum_(vertex.confidence), rationale: String(vertex.rationale_ja || '').slice(0, 1200), status: String(vertex.status || '') } : null
  };
  try { appJobPutResult_(key, out); } catch (e) { /* 控えられなくても返す */ }
  return out;
}
