/**
 * Insights.js — 責任者の 3 つの画面（分析・人の学び・AI の学び）の材料。読むだけ（書かない。旧来の計算・保存・権限は変えない）。
 *
 *   分析（appCrossMaker_）     … 年度のメーカーを横に並べる: 予測と予算・着地の見込み（一覧にあれば）・予測の改訂の道すじ・精度・AI の話題。
 *                                 話題ごとの市場のまとめ（上向き・下向きのメーカーの数と点数の平均）
 *   人の学び（appPeopleLearning_）… 外れた月の振り返り（同じ月の重複を除き、向きは予測と実績から決め直す）・人と話題ごとの当たり
 *                                 （ベータ分布の事後）・四半期レビューの判断の記録・予算の上乗せの実績
 *   AI の学び（appAiLearning_）   … 情報源の種類ごとの信頼度の学び（ベータ分布の四半期ごとの事後）・全計画で縮めた偏り・
 *                                 精度の推移（P10〜P90 に入った割合とその区間）・補正の変化・承認待ちの提案・学びの気になる点
 *
 * 計算用の表（ENG_*）は全計画の分を 1 回ずつ読む（appEngAll_）。計画ごとに組み立て直さない（計画が増えても読む回数は同じ）。
 * 旧来の記録の 2 つの問題を避けて、ここで数え直す（旧来の計算と記録は変えない）:
 *   - SUBJECTIVE_IMPACT_HISTORY の月は日付に変わっていることがあり、旧来の当たりの数え方（RELIABILITY_EVIDENCE）は月が合わず空になる
 *     → 月を appYm_ でそろえ、評価できた月をすべて数える
 *   - EVAL_INSIGHTS の cause_bucket は向きが逆（実績 > 予測を over_forecast と書く）→ 使わず、予測 − 実績の符号で決める
 */
const APP_INSIGHT_REVISIONS = 12;          // 予測の改訂の道すじ（最近の回数）
const APP_INSIGHT_MISS = 0.1;              // 外れた月: 誤差が 10% 以上か P10〜P90 の外（旧来の B-4 の振り分けと同じ）
const APP_INSIGHT_LESSONS = 50;            // 振り返りの最大件数（誤差の大きい順）
const APP_INSIGHT_OPEN_MAX = 100;          // 残っている対応の最大件数
const APP_INSIGHT_DECISIONS = 200;         // 判断の記録の最大件数（新しい順）
const APP_INSIGHT_PRIOR_K = 4;             // 事前分布が無いときの強さ（旧来の RELIABILITY_SHRINKAGE_K と同じ。μ = 0.5 で Beta(2, 2)）
const APP_INSIGHT_Z80 = 1.2815515655446004;   // 標準正規の 90% 点（両側 80% の区間）
const APP_INSIGHT_COVERAGE_TARGET = 0.8;   // P10〜P90 に実績が入る割合の目標
const APP_INSIGHT_COVERAGE_MIN_N = 4;      // 入った割合を気にするのは、これだけの月があるとき
const APP_INSIGHT_FACTOR_MIN = 0.75;       // 旧来の B-5 の偏りの補正の下限
const APP_INSIGHT_FACTOR_MAX = 1.25;       // 同じく上限
const APP_INSIGHT_STATUS = { open: '未着手', in_progress: '対応中', done: '済み', monitoring: '見守り' };
const APP_INSIGHT_ACTION_LABELS = { update: '前提を更新', keep: '継続', add: '追加', remove: '削除' };
const APP_INSIGHT_MACHINE_ACTIONS = ['update', 'keep', 'add', 'remove'];   // 旧来の B-4 が自動で入れる action_type
const APP_INSIGHT_DIRECTION = { over: '予測が高すぎた', under: '予測が低すぎた', exact: '予測どおり' };
const APP_INSIGHT_FIELD_LABELS = { ai_weight_override: 'AI の重み', ai_max_abs_effect_override: 'AI の効きの上限', ai_topic_disable_json: '使わない AI の話題',
  bias_correction_factor: '偏りの補正', residual_month_bias_json: '月ごとの偏りの補正', qual_scale_override: '入力の効きの倍率' };
const APP_INSIGHT_PHASES = { A: 'AI の重み', B: '情報源の信頼度' };
const APP_INSIGHT_HEALTH = {
  shadow_worse: '全計画で縮めた補正を試すと、今より誤差が大きくなる',
  factor_clamp: '偏りの補正が下限か上限（0.75 / 1.25）に張りついている',
  coverage_low: 'P10〜P90 に実績が入る割合が 80% より低い（幅が狭すぎる）',
  coverage_high: 'P10〜P90 に実績が入る割合が 80% より高い（幅が広すぎる）'
};

// ---- 共通: 計算用の表をまとめて読む ----

/**
 * 計算用の表 ENG_<sheet> を 1 回だけ読み、計画ごとの行にする（{ 計画の ID: [見出し → 値] }）。値は型の並び（_types）で元の型に戻す
 * （appEngTableObjects_ と同じ値・同じ並び。数式のセルは空、すべて空の行は除く）。only（計画の ID の一覧）を渡すと、その計画だけ。
 * 見出しが旧来の定義と違い、行ごとに持っている計画（ENG_SHEETS の mode = rows）は、その計画だけ appEngTableObjects_ で読む
 */
function appEngAll_(sheet, only) {
  const reg = APP_ENGINE_SHEETS[sheet];
  if (!reg || reg.mode !== 'table') throw new Error('表の形のシートではありません: ' + sheet);
  if (only && !only.length) return {};
  const want = only ? appInsightSet_(only) : null;
  const mode = {};
  appReadTable_('ENG_SHEETS').forEach(r => { if (r.sheet === sheet && (!want || want[r.plan_id])) mode[r.plan_id] = r.mode; });
  const header = reg.header;
  const by = {};
  appReadTable_('ENG_' + sheet).forEach(o => {
    if (mode[o.plan_id] !== 'table') return;
    const types = String(o._types || '');
    const x = {};
    let any = false;
    header.forEach((h, j) => {
      const t = types.charAt(j) || (o[h] === '' ? 'e' : 's');
      const v = t === 'f' ? '' : appCellDecode_(t, o[h]);
      x[h] = v;
      if (v !== '' && v !== null) any = true;
    });
    if (any) (by[o.plan_id] = by[o.plan_id] || []).push({ seq: Number(o.seq), row: x });
  });
  const out = {};
  Object.keys(mode).forEach(id => {
    out[id] = mode[id] === 'table' ? (by[id] || []).sort((a, b) => a.seq - b.seq).map(x => x.row) : appEngTableObjects_(id, [sheet])[sheet];
  });
  return out;
}

/** 1 回の呼び出しの中で、同じ表を 2 度読まない（表の名前 → appEngAll_ の結果） */
function appInsightTables_(ids) {
  const memo = {};
  return sheet => memo[sheet] || (memo[sheet] = appEngAll_(sheet, ids));
}

function appInsightSet_(xs) {
  const o = {};
  xs.forEach(x => { o[x] = true; });
  return o;
}

/** 閉じていない計画（{ planId, clientName, fy }。年度の新しい順・名前順） */
function appInsightPlans_() {
  const names = appClientNameMap_();
  return appReadTable_('PLANS').filter(p => p.state !== 'ARCHIVED')
    .map(p => ({ planId: p.plan_id, clientName: names[p.client_id] || p.client_label, fy: String(p.fy) }))
    .sort((x, y) => y.fy.localeCompare(x.fy) || String(x.clientName).localeCompare(String(y.clientName), 'ja'));
}

/** 計画ごとの精度（appAccuracyCached_ と同じ控え。控えが無い計画だけ、まとめて読んだ EVAL_LOG から計算する） */
function appInsightAccuracies_(ids, tab) {
  const accs = {};
  ids.forEach(id => {
    try { accs[id] = appAccuracyCached_(id, x => tab('EVAL_LOG')[x] || []); } catch (e) { Logger.log('精度を出せない計画: ' + id + ' ' + (e && e.message ? e.message : e)); }
  });
  return accs;
}

/** 日時を画面に出す形に（日時でなければそのまま） */
function appInsightIso_(v) {
  if (appIsDate_(v)) return isNaN(v.getTime()) ? '' : Utilities.formatDate(v, APP_TZ, "yyyy-MM-dd'T'HH:mm:ssZ");
  return v === null || v === undefined ? '' : String(v);
}

/** 旧来が文字列で残した値（'1.1' など）を、数なら数に。長い文字列は切る */
function appInsightVal_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : null;
  if (appIsDate_(v)) return appInsightIso_(v);
  const s = v === null || v === undefined ? '' : String(v).trim();
  return /^-?\d+(\.\d+)?(e[-+]?\d+)?$/i.test(s) ? Number(s) : s.slice(0, 300);
}

function appInsightMean_(xs) { return xs.length ? xs.reduce((a, b) => a + b, 0) / xs.length : null; }

/** 'yyyy/MM' の四半期（旧来の quarterLabelFromYm_ と同じ表記: FY2026-Q1 = 2026 年 4〜6 月） */
function appInsightQuarter_(ym) {
  const m = /^(\d{4})\/(\d{2})$/.exec(String(ym || ''));
  if (!m) return '';
  const y = Number(m[1]), mo = Number(m[2]);
  return 'FY' + (mo >= 4 ? y : y - 1) + '-Q' + (Math.floor(((mo + 8) % 12) / 3) + 1);
}

/** 'yyyy/MM' の年度（4 月始まり・始まりの年） */
function appInsightFyOfYm_(ym) {
  const y = Number(String(ym).slice(0, 4)), mo = Number(String(ym).slice(5, 7));
  return String(mo >= 4 ? y : y - 1);
}

/** 'yyyy/MM' の k か月前 */
function appInsightPrevYm_(ym, k) {
  const y = Number(String(ym).slice(0, 4)), mo = Number(String(ym).slice(5, 7)) - 1 - k;
  const yy = y + Math.floor(mo / 12), mm = ((mo % 12) + 12) % 12 + 1;
  return yy + '/' + ('0' + mm).slice(-2);
}

/** AI の向きを up / down / neutral に（旧来の normalizeAiDirection_ と同じ見分け方） */
function appInsightDir_(v) {
  const s = String(v === null || v === undefined ? '' : v).trim().toLowerCase();
  if (!s) return '';
  if (/(up|positive|posi|上昇|増)/.test(s)) return 'up';
  if (/(down|negative|nega|低下|減)/.test(s)) return 'down';
  if (/(neutral|flat|中立)/.test(s)) return 'neutral';
  return s;
}

/** 提案・補正の対象の名前（reliability:<種類>:<人や話題> は「見解「鷹野」の信頼度」のように） */
function appInsightTargetLabel_(field) {
  const f = String(field || '');
  if (f.indexOf('reliability:') === 0) {
    const parts = f.split(':');
    const type = parts[1] || '';
    const key = parts.slice(2).join(':');
    return (APP_SOURCE_LABELS[type] || type) + (key ? '「' + key + '」' : '') + 'の信頼度';
  }
  return APP_INSIGHT_FIELD_LABELS[f] || f;
}

// ---- 共通: ベータ分布と二項の区間 ----
// 80% の区間は、正則化不完全ベータ関数 I_x(a, b)（連分数で計算）を二分法で解いた分位点（正規近似は使わない。n が小さくても正しい）

/** log Γ(x)（Lanczos 近似 g = 7。相対誤差は 1e-15 ほど） */
function appLogGamma_(x) {
  if (x < 0.5) return Math.log(Math.PI / Math.sin(Math.PI * x)) - appLogGamma_(1 - x);
  const c = [0.99999999999980993, 676.5203681218851, -1259.1392167224028, 771.32342877765313, -176.61502916214059,
    12.507343278686905, -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7];
  const z = x - 1;
  let a = c[0];
  for (let i = 1; i < 9; i++) a += c[i] / (z + i);
  const t = z + 7.5;
  return 0.5 * Math.log(2 * Math.PI) + (z + 0.5) * Math.log(t) - t + Math.log(a);
}

/** I_x(a, b) の連分数（Lentz 法） */
function appBetaCf_(x, a, b) {
  const TINY = 1e-300;
  let c = 1;
  let d = 1 - (a + b) * x / (a + 1);
  if (Math.abs(d) < TINY) d = TINY;
  d = 1 / d;
  let h = d;
  for (let m = 1; m <= 1000; m++) {
    const m2 = 2 * m;
    let aa = m * (b - m) * x / ((a - 1 + m2) * (a + m2));
    d = 1 + aa * d; if (Math.abs(d) < TINY) d = TINY;
    c = 1 + aa / c; if (Math.abs(c) < TINY) c = TINY;
    d = 1 / d; h *= d * c;
    aa = -(a + m) * (a + b + m) * x / ((a + m2) * (a + 1 + m2));
    d = 1 + aa * d; if (Math.abs(d) < TINY) d = TINY;
    c = 1 + aa / c; if (Math.abs(c) < TINY) c = TINY;
    d = 1 / d;
    const del = d * c;
    h *= del;
    if (Math.abs(del - 1) < 1e-14) break;
  }
  return h;
}

/** 正則化不完全ベータ関数 I_x(a, b) = ベータ分布 Beta(a, b) の累積確率（a, b > 0） */
function appBetaInc_(x, a, b) {
  if (!(x > 0)) return 0;
  if (!(x < 1)) return 1;
  const front = Math.exp(appLogGamma_(a + b) - appLogGamma_(a) - appLogGamma_(b) + a * Math.log(x) + b * Math.log(1 - x));
  return x < (a + 1) / (a + b + 2) ? front * appBetaCf_(x, a, b) / a : 1 - front * appBetaCf_(1 - x, b, a) / b;
}

/** Beta(a, b) の p 分位点（I_x(a, b) = p を二分法で。60 回で幅 1e-18） */
function appBetaQuantile_(p, a, b) {
  let lo = 0, hi = 1;
  for (let i = 0; i < 60; i++) {
    const mid = (lo + hi) / 2;
    if (appBetaInc_(mid, a, b) < p) lo = mid; else hi = mid;
  }
  return (lo + hi) / 2;
}

/** Beta(a, b) の平均と 80% の区間 */
function appBetaSummary_(a, b) {
  return { alpha: a, beta: b, mean: a / (a + b), ci80: [appBetaQuantile_(0.1, a, b), appBetaQuantile_(0.9, a, b)] };
}

/** 二項の割合 k / n の Wilson の区間（z = 1.2816 で 80%）。n = 0 なら null */
function appWilson_(k, n, z) {
  if (!(n > 0)) return null;
  const p = k / n, z2 = z * z;
  const den = 1 + z2 / n;
  const center = (p + z2 / (2 * n)) / den;
  const half = z * Math.sqrt(p * (1 - p) / n + z2 / (4 * n * n)) / den;
  return [Math.max(0, center - half), Math.min(1, center + half)];
}

/** 情報源の種類の事前分布（POOL_PRIOR の pooled_value = 2μ、precision = α + β。無ければ Beta(2, 2)） */
function appInsightPrior_(r) {
  const v = r ? appNum_(r.pooled_value) : null;
  const k = r ? appNum_(r.precision) : null;
  const mu = v === null ? 0.5 : Math.min(0.99, Math.max(0.01, v / 2));
  const precision = k !== null && k > 0 ? k : APP_INSIGHT_PRIOR_K;
  return { mu: mu, precision: precision, alpha0: mu * precision, beta0: (1 - mu) * precision, from: v === null ? 'default' : 'pool' };
}

/** 全計画の POOL_PRIOR から、種類ごとに一番新しい事前分布（{ 種類: appInsightPrior_ }。全部の種類を入れる） */
function appInsightPriors_(poolByPlan) {
  const best = {};
  Object.keys(poolByPlan).forEach(id => poolByPlan[id].forEach((r, i) => {
    const scope = String(r.pool_scope || '').trim();
    if (scope.indexOf('reliability:') !== 0 || String(r.param_key || '').trim() !== 'reliability_r') return;
    const type = scope.slice('reliability:'.length);
    const t = appTimeKey_(r.updated_at);
    if (!best[type] || t > best[type].t) best[type] = { t: t, r: r };
  }));
  const out = {};
  APP_SOURCE_TYPES.concat(Object.keys(best)).forEach(type => { if (!out[type]) out[type] = appInsightPrior_(best[type] ? best[type].r : null); });
  return out;
}

/**
 * 入力と AI の話題が、評価できた月に当たったか（旧来の computeReliabilityHitStats_ と同じ判定。月をそろえ、全部の月で数える）。
 * 当たり = 押した向き（push_direction の符号）が、実績 − 過去の売上だけの予測（AI_IMPACT_HISTORY の pred_p50_quant_only）の符号と同じ。
 * 月・種類・人（話題）ごとに、予測の対象だった月（forecast_open）の一番新しい回だけを使う。実績は EVAL_LOG の neutral・constraint_relevant_flag = 1。
 * 返り値: [{ planId, ym, quarter, type, key, hit(0/1) }]
 */
function appSourceEvidence_(ids, tab) {
  const subj = tab('SUBJECTIVE_IMPACT_HISTORY');
  const impact = tab('AI_IMPACT_HISTORY');
  const evals = tab('EVAL_LOG');
  const newer = (p, t, i) => !p || t > p.t || (t === p.t && i > p.i);
  const out = [];
  ids.forEach(id => {
    const actual = {};
    (evals[id] || []).forEach(r => {
      if (String(r.scenario) !== 'neutral' || String(r.constraint_relevant_flag) !== '1') return;
      const ym = appYm_(r.target_month);
      if (ym) actual[ym] = Number(r.actual || 0);
    });
    const quant = {};
    (impact[id] || []).forEach((r, i) => {
      if (String(r.forecast_source || '').trim() !== 'forecast_open') return;
      const ym = appYm_(r.target_month);
      const t = appTimeKey_(r.run_at);
      if (ym && newer(quant[ym], t, i)) quant[ym] = { v: Number(r.pred_p50_quant_only || 0), t: t, i: i };
    });
    const latest = {};
    (subj[id] || []).forEach((r, i) => {
      if (String(r.forecast_source || '').trim() !== 'forecast_open') return;
      const ym = appYm_(r.target_month);
      const type = String(r.source_type || '').trim();
      const key = String(r.source_key || '').trim();
      if (!ym || !type || !key) return;
      const u = ym + '\u0001' + type + '\u0001' + key;
      const t = appTimeKey_(r.run_at);
      if (newer(latest[u], t, i)) latest[u] = { ym: ym, type: type, key: key, dir: Math.sign(Number(r.push_direction || 0)), t: t, i: i };
    });
    Object.keys(latest).forEach(u => {
      const x = latest[u];
      if (!Object.prototype.hasOwnProperty.call(actual, x.ym) || !quant[x.ym] || !isFinite(actual[x.ym]) || !isFinite(quant[x.ym].v)) return;
      const surprise = Math.sign(actual[x.ym] - quant[x.ym].v);
      if (!surprise || !x.dir) return;
      out.push({ planId: id, ym: x.ym, quarter: appInsightQuarter_(x.ym), type: x.type, key: x.key, hit: x.dir === surprise ? 1 : 0 });
    });
  });
  return out.sort((a, b) => a.ym.localeCompare(b.ym) || a.planId.localeCompare(b.planId) || a.type.localeCompare(b.type) || a.key.localeCompare(b.key));
}

// ---- 分析（メーカーを横に並べる） ----

/**
 * 年度（省くと今の年度。無ければ一番新しい年度）のメーカーごとの要点。
 * 返り値: { fy, fys, plans: [計画の一覧（appPortfolioCached_）の行の項目すべて + revisions, accuracy, topics], totals, market }
 *   revisions: [{ at, p10, p50, p90 }]（新アプリで動かした予測。古い順に最近の 12 回）
 *   accuracy:  { n, leaks, mape, bias, coverage, coverageN, widthScale }（Learning.js の精度。締まった後の予測の月は除く）
 *   topics:    [{ topic, rowType, direction, impact, confidence, score, position, percentile, horizon, asOf }]（AI 調査の一番新しい回）
 *   totals:    { plans, budget, p50, landing, actualYtd, budgetPlans, p50Budgeted, landingBudgeted, ratioP50, ratioLanding }（合計。比は予算のある計画だけで）
 *   market:    [{ topic, makers, up, down, flat, meanScore }]（話題ごと。メーカーの向きは、その話題の行の点数の平均の符号）
 */
function appCrossMaker_(fy) {
  const all = appPortfolioCached_();
  const fys = all.map(p => String(p.fy)).filter((x, i, a) => a.indexOf(x) === i).sort();
  const thisFy = String(appFy_(new Date()));
  const asked = /^\d{4}$/.test(String(fy || '')) ? String(fy) : '';
  const pick = asked || (fys.indexOf(thisFy) >= 0 ? thisFy : (fys.length ? fys[fys.length - 1] : thisFy));
  const plans = all.filter(p => String(p.fy) === pick);
  const ids = plans.map(p => p.planId);
  const want = appInsightSet_(ids);
  const tab = appInsightTables_(ids);
  const runs = {};
  appReadTable_('FORECAST_RUNS').filter(r => r.status === 'DONE' && want[r.plan_id]).forEach(r => { (runs[r.plan_id] = runs[r.plan_id] || []).push(r); });
  const accs = appInsightAccuracies_(ids, tab);
  const research = tab('AI_RESEARCH_STRUCTURED');
  const market = {};
  const rows = plans.map(p => {
    const rs = (runs[p.planId] || []).sort((a, b) => String(a.finished_at).localeCompare(String(b.finished_at)) || a._row - b._row).slice(-APP_INSIGHT_REVISIONS);
    const a = accs[p.planId] || null;
    const topics = appLatestBatch_(research[p.planId] || [], 'as_of_date').map(r => ({
      topic: String(r.topic || '').trim(), rowType: String(r.row_type || ''), direction: appInsightDir_(r.direction), impact: appNum_(r.impact_score),
      confidence: appNum_(r.confidence), score: appNum_(r.blended_score), position: String(r.relative_position_label || ''),
      percentile: appNum_(r.relative_percentile), horizon: String(r.time_horizon || ''), asOf: appYm_(r.as_of_date)
    })).filter(t => t.topic);
    const byTopic = {};
    topics.forEach(t => {
      const o = byTopic[t.topic] = byTopic[t.topic] || { scores: [], up: 0, down: 0 };
      if (t.score !== null) o.scores.push(t.score);
      if (t.direction === 'up') o.up++;
      if (t.direction === 'down') o.down++;
    });
    Object.keys(byTopic).forEach(k => {
      const o = byTopic[k];
      const score = appInsightMean_(o.scores);
      const dir = Math.sign(score !== null ? score : o.up - o.down);
      const m = market[k] = market[k] || { topic: k, makers: 0, up: 0, down: 0, flat: 0, scores: [] };
      m.makers++;
      if (dir > 0) m.up++; else if (dir < 0) m.down++; else m.flat++;
      if (score !== null) m.scores.push(score);
    });
    return Object.assign({}, p, {
      planId: p.planId, clientName: p.clientName, fy: String(p.fy),
      revisions: rs.map(r => ({ at: r.finished_at, p10: r.annual_p10, p50: r.annual_p50, p90: r.annual_p90 })),
      accuracy: a ? { n: a.n, leaks: a.leaks, mape: a.mape, bias: a.bias, coverage: a.coverage, coverageN: a.coverageN, widthScale: a.widthScale } : null,
      topics: topics
    });
  });
  const sum = (xs, k) => xs.reduce((s, p) => (typeof p[k] === 'number' && isFinite(p[k]) ? (s || 0) + p[k] : s), null);
  const budgeted = rows.filter(p => typeof p.budget === 'number' && p.budget > 0);
  const totals = { plans: rows.length, budget: sum(rows, 'budget'), p50: sum(rows, 'p50'), landing: sum(rows, 'landing'), actualYtd: sum(rows, 'actualYtd'),
    budgetPlans: budgeted.length, p50Budgeted: sum(budgeted.filter(p => typeof p.p50 === 'number'), 'p50'), landingBudgeted: sum(budgeted.filter(p => typeof p.landing === 'number'), 'landing') };
  const budgetOf = k => sum(budgeted.filter(p => typeof p[k] === 'number'), 'budget');
  totals.ratioP50 = totals.p50Budgeted !== null && budgetOf('p50') > 0 ? totals.p50Budgeted / budgetOf('p50') : null;
  totals.ratioLanding = totals.landingBudgeted !== null && budgetOf('landing') > 0 ? totals.landingBudgeted / budgetOf('landing') : null;
  return {
    fy: pick, fys: fys, plans: rows, totals: totals,
    market: Object.keys(market).map(k => ({ topic: k, makers: market[k].makers, up: market[k].up, down: market[k].down, flat: market[k].flat,
      meanScore: appInsightMean_(market[k].scores) })).sort((x, y) => y.makers - x.makers || x.topic.localeCompare(y.topic))
  };
}

// ---- 人の学び ----

/**
 * 人の学び。返り値:
 *   lessons:     [{ planId, clientName, fy, ym, actual, pred, err, absErr, direction, directionLabel, range, miss, hypothesis, actionType, actionLabel,
 *                   reflection, owner, status, statusLabel, human, duplicates }]（誤差の大きい順に 50 件。err = (予測 − 実績) / |実績|）
 *   causes:      [{ direction, directionLabel, range, actionType, actionLabel, n }]（外れた月の、向き × 幅の外か × 対応の数）
 *   openActions: [lessons と同じ形]（未着手・対応中。古い月から）
 *   repeats:     [{ clientName, month: 'MM', fys: [...], n }]（同じメーカーで、違う年度の同じ月にまた外れた）
 *   summary:     { months, misses, withNotes, duplicatesRemoved }
 *   scoreboard:  [{ type, label, key, n, hit, hitRate, alpha, beta, postMean, ci80: [下, 上], prior: { alpha0, beta0, from }, appliedR,
 *                   plans: [{ planId, clientName, fy, n, hit, appliedR }] }]（80% の区間の下の端が高い順。少ない数で上位に出ない）
 *   decisions:   [{ planId, clientName, fy, reviewId, proposalId, at, quarter, phase, phaseLabel, target, targetLabel, current, proposed,
 *                   confidence, rationale, status, decidedAt, decidedBy, applied, appliedAt }]（四半期レビューの提案と判断。新しい順）
 *   uplift:      [{ planId, clientName, fy, versionNo, budgetAdopted, budgetUplift, budgetFinal, p50, actual, months, complete, ratio, upliftRealized }]
 *                （公式版の予算と、その年度の実績の合計。complete = 12 か月の実績がそろった）
 */
function appPeopleLearning_() {
  const plans = appInsightPlans_();
  const ids = plans.map(p => p.planId);
  const tab = appInsightTables_(ids);
  const l = appInsightLessons_(plans, tab('EVAL_INSIGHTS'));
  const evidence = appSourceEvidence_(ids, tab);
  return Object.assign(l, {
    scoreboard: appSourceScoreboard_(plans, evidence, tab('POOL_PRIOR'), tab('SOURCE_RELIABILITY')),
    decisions: appInsightDecisions_(plans, tab('QUARTERLY_REVIEW_LOG')),
    uplift: appInsightUplift_(ids)
  });
}

/** 人が書いた跡があるか（旧来の B-4 は action_type・next_cycle_reflection・status を自動で入れるので、それ以外を見る） */
function appInsightHasHuman_(r) {
  const s = k => String(r[k] === null || r[k] === undefined ? '' : r[k]).trim();
  const action = s('action_type');
  return !!(s('cause_hypothesis') || s('owner') || ['in_progress', 'done'].indexOf(s('status')) >= 0 ||
    (action && APP_INSIGHT_MACHINE_ACTIONS.indexOf(action) < 0));
}

/**
 * 外れた月の振り返り。B-4 を動かし直すと同じ月の行が増える（月が日付に変わり、旧来の上書きの鍵が合わない）ので、計画 × 月で 1 つにする:
 * 数字（実績・予測・幅の外か）は一番新しい行から、人が書いた欄（原因・対応・担当・状態）は人が書いた一番新しい行から取る
 */
function appInsightLessons_(plans, insights) {
  const all = [];
  let dup = 0;
  const newer = (p, t, i) => !p || t > p.t || (t === p.t && i > p.i);
  plans.forEach(p => {
    const byYm = {};
    (insights[p.planId] || []).forEach((r, i) => {
      const ym = appYm_(r.target_month);
      if (!/^\d{4}\/\d{2}$/.test(ym)) return;
      const o = byYm[ym] = byYm[ym] || { latest: null, human: null, rows: 0 };
      const t = appTimeKey_(r.evaluated_at);
      o.rows++;
      if (newer(o.latest, t, i)) o.latest = { r: r, t: t, i: i };
      if (appInsightHasHuman_(r) && newer(o.human, t, i)) o.human = { r: r, t: t, i: i };
    });
    Object.keys(byYm).forEach(ym => {
      const x = byYm[ym];
      dup += x.rows - 1;
      const m = x.latest.r;
      const h = x.human ? x.human.r : m;
      const pred = appNum_(m.pred_p50), act = appNum_(m.actual_total);
      const err = pred !== null && act !== null && act !== 0 ? (pred - act) / Math.abs(act) : null;
      const range = String(m.range_breach) === '1' || m.range_breach === true;
      const direction = pred === null || act === null ? '' : pred > act ? 'over' : pred < act ? 'under' : 'exact';
      const status = String(h.status || '').trim();
      const actionType = String(h.action_type || '').trim();
      all.push({ planId: p.planId, clientName: p.clientName, fy: p.fy, ym: ym, actual: act, pred: pred, err: err, absErr: err === null ? null : Math.abs(err),
        direction: direction, directionLabel: APP_INSIGHT_DIRECTION[direction] || '', range: range,
        miss: (err !== null && Math.abs(err) >= APP_INSIGHT_MISS) || range,
        hypothesis: String(h.cause_hypothesis || ''), actionType: actionType, actionLabel: APP_INSIGHT_ACTION_LABELS[actionType] || actionType,
        reflection: String(h.next_cycle_reflection || ''), owner: String(h.owner || ''), status: status, statusLabel: APP_INSIGHT_STATUS[status] || status,
        human: !!x.human, duplicates: x.rows - 1 });
    });
  });
  const byErr = all.slice().sort((a, b) => (b.absErr === null ? -1 : b.absErr) - (a.absErr === null ? -1 : a.absErr) || a.ym.localeCompare(b.ym));
  const misses = all.filter(x => x.miss);
  const causes = {};
  misses.forEach(x => {
    const k = [x.direction, x.range ? 1 : 0, x.actionType].join('\u0001');
    const o = causes[k] = causes[k] || { direction: x.direction, directionLabel: x.directionLabel, range: x.range, actionType: x.actionType, actionLabel: x.actionLabel, n: 0 };
    o.n++;
  });
  const rep = {};
  misses.forEach(x => {
    const k = x.clientName + '\u0001' + x.ym.slice(5, 7);
    const o = rep[k] = rep[k] || { clientName: x.clientName, month: x.ym.slice(5, 7), fys: [], n: 0 };
    o.n++;
    const f = appInsightFyOfYm_(x.ym);
    if (o.fys.indexOf(f) < 0) o.fys.push(f);
  });
  return {
    lessons: byErr.slice(0, APP_INSIGHT_LESSONS),
    causes: Object.keys(causes).map(k => causes[k]).sort((a, b) => b.n - a.n || a.direction.localeCompare(b.direction)),
    openActions: all.filter(x => x.status === 'open' || x.status === 'in_progress')
      .sort((a, b) => a.ym.localeCompare(b.ym) || String(a.clientName).localeCompare(String(b.clientName), 'ja')).slice(0, APP_INSIGHT_OPEN_MAX),
    repeats: Object.keys(rep).map(k => rep[k]).filter(o => o.fys.length >= 2).map(o => Object.assign(o, { fys: o.fys.sort() }))
      .sort((a, b) => b.n - a.n || a.month.localeCompare(b.month)),
    summary: { months: all.length, misses: misses.length, withNotes: all.filter(x => x.human).length, duplicatesRemoved: dup }
  };
}

/** 人と話題ごとの当たり（全計画・全部の評価できた月）と、ベータ分布の事後。事前分布は POOL_PRIOR（無ければ Beta(2, 2)） */
function appSourceScoreboard_(plans, evidence, poolByPlan, relByPlan) {
  const info = {};
  plans.forEach(p => { info[p.planId] = p; });
  const priors = appInsightPriors_(poolByPlan);
  const applied = {};   // 種類 + 人 → { 計画: 今の信頼度 r }（旧来の C-3 が書いた SOURCE_RELIABILITY。無ければ旧来は 1.0 を使う）
  Object.keys(relByPlan).forEach(id => relByPlan[id].forEach(r => {
    const v = appNum_(r.reliability_r);
    if (v === null) return;
    const k = String(r.source_type || '').trim() + '\u0001' + String(r.source_key || '').trim();
    (applied[k] = applied[k] || {})[id] = v;   // 同じ人の行が 2 つあれば、後の行
  }));
  const g = {};
  evidence.forEach(e => {
    const k = e.type + '\u0001' + e.key;
    const o = g[k] = g[k] || { type: e.type, key: e.key, n: 0, hit: 0, by: {} };
    o.n++; o.hit += e.hit;
    const b = o.by[e.planId] = o.by[e.planId] || { n: 0, hit: 0 };
    b.n++; b.hit += e.hit;
  });
  return Object.keys(g).map(k => {
    const o = g[k];
    const pr = priors[o.type] || appInsightPrior_(null);
    const post = appBetaSummary_(pr.alpha0 + o.hit, pr.beta0 + o.n - o.hit);
    const ap = applied[k] || {};
    const apVals = Object.keys(ap).map(id => ap[id]);
    return { type: o.type, label: APP_SOURCE_LABELS[o.type] || o.type, key: o.key, n: o.n, hit: o.hit, hitRate: o.hit / o.n,
      alpha: post.alpha, beta: post.beta, postMean: post.mean, ci80: post.ci80, prior: { alpha0: pr.alpha0, beta0: pr.beta0, from: pr.from },
      appliedR: appInsightMean_(apVals),
      plans: Object.keys(o.by).map(id => ({ planId: id, clientName: (info[id] || {}).clientName || '', fy: (info[id] || {}).fy || '', n: o.by[id].n, hit: o.by[id].hit,
        appliedR: ap[id] === undefined ? null : ap[id] })) };
  }).sort((a, b) => b.ci80[0] - a.ci80[0] || b.n - a.n || a.type.localeCompare(b.type) || a.key.localeCompare(b.key));
}

/** 四半期レビューの提案と、人の判断（QUARTERLY_REVIEW_LOG。新しい順） */
function appInsightDecisions_(plans, logs) {
  const out = [];
  plans.forEach(p => (logs[p.planId] || []).forEach((r, i) => {
    const field = String(r.target_field || '');
    const phase = String(r.phase || '').trim();
    out.push({ t: appTimeKey_(r.reviewed_at), i: i, planId: p.planId, clientName: p.clientName, fy: p.fy,
      reviewId: String(r.review_id || ''), proposalId: String(r.proposal_id || ''), at: appInsightIso_(r.reviewed_at), quarter: String(r.quarter_label || ''),
      phase: phase, phaseLabel: APP_INSIGHT_PHASES[phase] || phase, target: field, targetLabel: appInsightTargetLabel_(field),
      current: appInsightVal_(r.current_value), proposed: appInsightVal_(r.proposed_value), confidence: String(r.confidence || ''),
      rationale: String(r.rationale || '').slice(0, 300), status: String(r.approval_status || ''), decidedAt: appInsightIso_(r.approval_decided_at),
      decidedBy: String(r.approval_decided_by || '').split('@')[0], applied: Number(r.applied || 0) === 1, appliedAt: appInsightIso_(r.applied_at) });
  }));
  return out.sort((a, b) => b.t - a.t || String(a.planId).localeCompare(String(b.planId)) || a.i - b.i).slice(0, APP_INSIGHT_DECISIONS)
    .map(x => { const o = Object.assign({}, x); delete o.t; delete o.i; return o; });
}

/** 公式版（承認された版）の予算と、その年度の実績の合計（計画の一覧の暫定実績 = EVAL_COMPARE_MONTHLY の年度の月の実績） */
function appInsightUplift_(ids) {
  const want = appInsightSet_(ids);
  const port = {};
  appPortfolioCached_().forEach(p => { if (want[p.planId]) port[p.planId] = p; });
  return appVersionTable_().filter(v => v.state === 'APPROVED' && port[v.plan_id]).map(v => {
    const p = port[v.plan_id];
    const actual = typeof p.actualYtd === 'number' ? p.actualYtd : null;
    const months = Number(p.actualMonths || 0);
    return { planId: p.planId, clientName: p.clientName, fy: String(p.fy), versionNo: v.version_no,
      budgetAdopted: v.budget_adopted, budgetUplift: v.budget_uplift, budgetFinal: v.budget_final, p50: v.annual_p50,
      actual: actual, months: months, complete: months >= 12,
      ratio: actual !== null && v.budget_final > 0 ? actual / v.budget_final : null,
      // 上乗せのうち実際に出た割合（(実績 − 採用した予算) / 上乗せ）。年度が終わるまでは途中の値
      upliftRealized: actual !== null && v.budget_uplift && v.budget_adopted !== null ? (actual - v.budget_adopted) / v.budget_uplift : null };
  }).sort((a, b) => b.fy.localeCompare(a.fy) || String(a.clientName).localeCompare(String(b.clientName), 'ja'));
}

// ---- AI の学び ----

/**
 * AI の学び。返り値:
 *   curves:      [{ type, label, prior: { alpha0, beta0, mu, precision, from, mean, ci80 }, n, hit,
 *                   quarters: [{ quarter, n, hit, cumN, cumHit, alpha, beta, mean, ci80 }] }]（四半期ごとに積み上げた事後。画面がベータ分布の曲線を描く）
 *   shrinkage:   { mu, tau, tau2, pooled, plans: [{ planId, clientName, fy, n, bias, se, shrunk, factorShadow, mapeNow, mapeShadow }] }
 *   timeline:    [{ ym, n, mape, coverage, coverageN, coverageCi80: [下, 上], target, rolling3, rolling3N }]（全計画の月ごと。締まった後の予測の月は除く）
 *   calibration: [{ planId, clientName, fy, factorNow, path: [{ at, factor, factorLabel, old, new, source, sourceLabel, quarter }] }]
 *   pending:     [{ planId, clientName, fy, reviewId, at, quarter, proposals: [{ proposalId, phase, phaseLabel, target, targetLabel, current, proposed,
 *                   confidence, rationale, decision }] }]（一番新しい四半期レビューで、まだ適用していないもの。decision = 保存した判断）
 *   health:      [{ planId, clientName, fy, key, label, value }]（気になる点。key = shadow_worse / factor_clamp / coverage_low / coverage_high）
 */
function appAiLearning_() {
  const plans = appInsightPlans_();
  const ids = plans.map(p => p.planId);
  const tab = appInsightTables_(ids);
  const evidence = appSourceEvidence_(ids, tab);
  const priors = appInsightPriors_(tab('POOL_PRIOR'));
  const curves = Object.keys(priors).map(type => {
    const pr = priors[type];
    const byQ = {};
    evidence.forEach(e => {
      if (e.type !== type || !e.quarter) return;
      const q = byQ[e.quarter] = byQ[e.quarter] || { n: 0, hit: 0 };
      q.n++; q.hit += e.hit;
    });
    let a = pr.alpha0, b = pr.beta0, n = 0, hit = 0;
    const quarters = Object.keys(byQ).sort().map(q => {
      n += byQ[q].n; hit += byQ[q].hit;
      a += byQ[q].hit; b += byQ[q].n - byQ[q].hit;
      const s = appBetaSummary_(a, b);
      return { quarter: q, n: byQ[q].n, hit: byQ[q].hit, cumN: n, cumHit: hit, alpha: a, beta: b, mean: s.mean, ci80: s.ci80 };
    });
    const p0 = appBetaSummary_(pr.alpha0, pr.beta0);
    return { type: type, label: APP_SOURCE_LABELS[type] || type, prior: { alpha0: pr.alpha0, beta0: pr.beta0, mu: pr.mu, precision: pr.precision, from: pr.from,
      mean: p0.mean, ci80: p0.ci80 }, n: n, hit: hit, quarters: quarters };
  });
  const info = {};
  plans.forEach(p => { info[p.planId] = p; });
  const accs = appInsightAccuracies_(ids, tab);
  const shadow = appBiasShadow_(accs);
  const shrinkage = { mu: shadow.mu, tau: Math.sqrt(shadow.tau2), tau2: shadow.tau2, pooled: shadow.pooled,
    plans: plans.filter(p => shadow.plans[p.planId]).map(p => {
      const s = shadow.plans[p.planId];
      return { planId: p.planId, clientName: p.clientName, fy: p.fy, n: accs[p.planId].n, bias: s.bias, se: s.se, shrunk: s.shrunk,
        factorShadow: s.factorShadow, mapeNow: s.mapeNow, mapeShadow: s.mapeShadow };
    }) };
  const state = tab('CALIBRATION_STATE');
  const factorNow = {};
  ids.forEach(id => { const c = (state[id] || []).slice(-1)[0]; factorNow[id] = c ? appNum_(c.bias_correction_factor) : null; });
  return {
    curves: curves, shrinkage: shrinkage, timeline: appInsightTimeline_(accs),
    calibration: appInsightCalibration_(plans, tab('CALIBRATION_HISTORY'), factorNow),
    pending: appInsightPending_(plans, tab('QUARTERLY_REVIEW_LOG')),
    health: appInsightHealth_(plans, accs, shadow, factorNow)
  };
}

/** 全計画の月ごとの誤差と、P10〜P90 に入った割合（Wilson の 80% の区間）。rolling3 = その月までの 3 か月の誤差の平均（全計画の月をまとめて） */
function appInsightTimeline_(accs) {
  const by = {};
  Object.keys(accs).forEach(id => (accs[id].months || []).forEach(m => {
    if (m.leak || !/^\d{4}\/\d{2}$/.test(String(m.month))) return;
    const o = by[m.month] = by[m.month] || { apes: [], inside: 0, rangeN: 0 };
    if (typeof m.ape === 'number' && isFinite(m.ape)) o.apes.push(m.ape);
    if (m.inside === true || m.inside === false) { o.rangeN++; if (m.inside) o.inside++; }
  }));
  return Object.keys(by).sort().map(ym => {
    const o = by[ym];
    const win = [].concat.apply([], [ym, appInsightPrevYm_(ym, 1), appInsightPrevYm_(ym, 2)].map(k => (by[k] ? by[k].apes : [])));
    return { ym: ym, n: o.apes.length, mape: appInsightMean_(o.apes), coverage: o.rangeN ? o.inside / o.rangeN : null, coverageN: o.rangeN,
      coverageCi80: appWilson_(o.inside, o.rangeN, APP_INSIGHT_Z80), target: APP_INSIGHT_COVERAGE_TARGET, rolling3: appInsightMean_(win), rolling3N: win.length };
  });
}

/** 補正の変化（CALIBRATION_HISTORY。月次の自動学習 AUTO-MONTHLY と、四半期レビューの適用） */
function appInsightCalibration_(plans, hist, factorNow) {
  return plans.map(p => {
    const rows = (hist[p.planId] || []).map((r, i) => ({ t: appTimeKey_(r.changed_at), i: i, r: r })).sort((a, b) => a.t - b.t || a.i - b.i);
    return { planId: p.planId, clientName: p.clientName, fy: p.fy, factorNow: factorNow[p.planId] === undefined ? null : factorNow[p.planId],
      path: rows.map(x => {
        const r = x.r;
        const rid = String(r.review_id || '').trim();
        const auto = rid === 'AUTO-MONTHLY';
        return { at: appInsightIso_(r.changed_at), factor: String(r.factor_name || ''), factorLabel: appInsightTargetLabel_(r.factor_name),
          old: appInsightVal_(r.old_value), new: appInsightVal_(r.new_value), source: rid, sourceLabel: auto ? '月次の自動学習' : '四半期レビュー',
          quarter: String(r.quarter_label || '') };
      }) };
  }).filter(x => x.path.length || x.factorNow !== null);
}

/**
 * 承認待ちの提案: 計画ごとの一番新しい四半期レビュー（QUARTERLY_REVIEW_LOG の review_id）で、まだ C-3 で処理していないもの
 * （どの行も approval_decided_at が空で applied = 0）。保存した判断（承認・却下・保留）は、その計画の QUARTERLY_REVIEW の画面の行から読む
 */
function appInsightPending_(plans, logs) {
  const out = [];
  plans.forEach(p => {
    const byRid = {};
    (logs[p.planId] || []).forEach((r, i) => {
      const rid = String(r.review_id || '').trim();
      if (!rid) return;
      const o = byRid[rid] = byRid[rid] || { rid: rid, t: -Infinity, i: -1, rows: [] };
      const t = appTimeKey_(r.reviewed_at);
      if (t > o.t || (t === o.t && i > o.i)) { o.t = t; o.i = i; }
      o.rows.push(r);
    });
    const latest = Object.keys(byRid).map(k => byRid[k]).sort((a, b) => b.t - a.t || b.i - a.i)[0];
    if (!latest) return;
    if (latest.rows.some(r => Number(r.applied || 0) === 1 || (r.approval_decided_at !== '' && r.approval_decided_at !== null && r.approval_decided_at !== undefined))) return;
    const saved = appInsightReviewDecisions_(p.planId);
    const decided = saved.reviewId === latest.rid ? saved.byPid : {};
    const r0 = latest.rows[0];
    out.push({ planId: p.planId, clientName: p.clientName, fy: p.fy, reviewId: latest.rid, at: appInsightIso_(r0.reviewed_at), quarter: String(r0.quarter_label || ''),
      proposals: latest.rows.map(r => {
        const field = String(r.target_field || '');
        const phase = String(r.phase || '').trim();
        const pid = String(r.proposal_id || '');
        return { proposalId: pid, phase: phase, phaseLabel: APP_INSIGHT_PHASES[phase] || phase, target: field, targetLabel: appInsightTargetLabel_(field),
          current: appInsightVal_(r.current_value), proposed: appInsightVal_(r.proposed_value), confidence: String(r.confidence || ''),
          rationale: String(r.rationale || '').slice(0, 300), decision: decided[pid] || '' };
      }) });
  });
  return out;
}

/** 計画の QUARTERLY_REVIEW（表の形でないシート）の、提案の行（8 行目から）の保存した判断（8 列目）と review_id（8 行目の 10 列目） */
function appInsightReviewDecisions_(planId) {
  const out = { reviewId: '', byPid: {} };
  appReadPlanTable_('ENG_ROWS', planId).forEach(s => {
    if (s.sheet !== 'QUARTERLY_REVIEW' || Number(s.row_no) < 8) return;
    let cells;
    try { cells = JSON.parse(s.cells_json); } catch (e) { return; }
    const c0 = Number(s.col_from) - 1;
    const at = c => {
      const x = c >= c0 ? cells[c - c0] : undefined;
      if (x === undefined || x === null || x === '') return '';
      const t = String(x).charAt(0);
      return t === 'f' ? '' : String(appCellDecode_(t, String(x).slice(1))).trim();
    };
    if (Number(s.row_no) === 8 && at(9)) out.reviewId = at(9);
    if (at(0)) out.byPid[at(0)] = at(7);
  });
  return out;
}

/** 学びの気になる点（計画ごと） */
function appInsightHealth_(plans, accs, shadow, factorNow) {
  const out = [];
  const push = (p, key, value) => out.push({ planId: p.planId, clientName: p.clientName, fy: p.fy, key: key, label: APP_INSIGHT_HEALTH[key], value: value });
  plans.forEach(p => {
    const s = shadow && shadow.plans ? shadow.plans[p.planId] : null;
    if (s && typeof s.mapeShadow === 'number' && typeof s.mapeNow === 'number' && s.mapeShadow > s.mapeNow + 1e-9) push(p, 'shadow_worse', { now: s.mapeNow, shadow: s.mapeShadow });
    const f = factorNow[p.planId];
    if (typeof f === 'number' && isFinite(f) && (f <= APP_INSIGHT_FACTOR_MIN + 1e-9 || f >= APP_INSIGHT_FACTOR_MAX - 1e-9)) push(p, 'factor_clamp', f);
    const a = accs[p.planId];
    if (a && typeof a.coverage === 'number' && a.coverageN >= APP_INSIGHT_COVERAGE_MIN_N) {
      const w = appWilson_(Math.round(a.coverage * a.coverageN), a.coverageN, APP_INSIGHT_Z80);
      const v = { coverage: a.coverage, n: a.coverageN, ci80: w };
      if (w[1] < APP_INSIGHT_COVERAGE_TARGET) push(p, 'coverage_low', v);
      else if (w[0] > APP_INSIGHT_COVERAGE_TARGET) push(p, 'coverage_high', v);
    }
  });
  return out;
}
