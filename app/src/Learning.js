/**
 * Learning.js — 精度の推移と、ベイジアン的な学習（進化の計画 C）。
 *
 * 旧来の計算は変えない。ここで計算するものは、次の 2 つに分ける:
 *   影（予測に効かない。画面に出すだけ）
 *     - 精度の推移: 月の誤差・偏り・P10〜P90 に実績が入った割合（目標 80%）。締まった後の予測（誤差 0）は除いて数える
 *     - 全計画で縮めた偏りの補正（階層ベイズ / 経験ベイズ）: 計画ごとの偏りを、全計画の平均へ、ばらつきに応じて縮める
 *     - P10〜P90 の幅の較正（コンフォーマル）: 実績の 80% が入る幅にするには、今の幅を何倍にすればよいか
 *   人の判断で入れるもの
 *     - 情報源の信頼度の事前分布（ベータ二項の経験ベイズ）: 全計画の RELIABILITY_EVIDENCE（当たった数 / 数）から作り、
 *       各計画の POOL_PRIOR に入れる。旧来の C-1（四半期レビューの提案）が事前分布として使う（予測に効くのは、提案を人が承認してから）
 */
const APP_SOURCE_TYPES = ['factor_product', 'factor_client', 'opinion', 'ai_topic', 'vertex_forecast'];
const APP_POOL_MIN_PLANS = 2;      // 旧来の POOL_MIN_CLIENTS と同じ
const APP_POOL_MIN_SAMPLES = 5;    // 計画ごとに、これだけの数がある計画だけで作る
const APP_POOL_PRECISION_MIN = 2;
const APP_POOL_PRECISION_MAX = 50;
const APP_BIAS_HALF_LIFE = 4;      // 旧来の B-5 と同じ（月）

/** 計画の EVAL_LOG を、月ごとの P10/P50/P90 と実績にする */
function appEvalMonths_(planId) {
  const rows = appEngTableObjects_(planId, ['EVAL_LOG']).EVAL_LOG;
  const by = {};
  rows.forEach(r => {
    const ym = appYm_(r.target_month);
    const o = by[ym] = by[ym] || { month: ym };
    const pred = appNum_(r.pred), act = appNum_(r.actual);
    if (r.scenario === 'nega') o.p10 = pred;
    if (r.scenario === 'posi') o.p90 = pred;
    if (r.scenario === 'neutral') { o.p50 = pred; o.actual = act; o.evaluatedAt = r.evaluated_at; }
  });
  return Object.keys(by).sort().map(k => by[k]).filter(m => m.p50 !== null && m.p50 !== undefined && m.actual !== null && m.actual !== undefined && m.actual !== 0);
}

/** 重み付きの平均 */
function appWMean_(xs, ws) { const sw = ws.reduce((a, b) => a + b, 0); return sw ? xs.reduce((a, x, i) => a + x * ws[i], 0) / sw : null; }

function appQuantile_(xs, q) {
  if (!xs.length) return null;
  const s = xs.slice().sort((a, b) => a - b);
  const pos = Math.min(s.length - 1, Math.max(0, Math.ceil(q * s.length) - 1));
  return s[pos];
}

/** 計画の精度を、入力のハッシュが同じあいだ 6 時間控える（全計画を見るので、毎回は計算しない） */
function appAccuracyCached_(planId) {
  const key = 'ACC_' + appSha256Hex_([planId, appPlanInputHash_(planId), APP_VERSION].join('|')).slice(0, 32);
  const c = appJobGetResult_(key);
  if (c.found) return c.value;
  const v = appAccuracyOf_(planId);
  try { appJobPutResult_(key, v); } catch (e) { /* 控えられなくても返す */ }
  return v;
}

/** 計画の精度（影）。leak = 予測が実績とまったく同じ月（締まった後に予測し直した月。旧来の検証が拾ってしまう） */
function appAccuracyOf_(planId) {
  const ms = appEvalMonths_(planId).map(m => {
    const err = (m.p50 - m.actual) / Math.abs(m.actual);
    const hasRange = m.p10 !== null && m.p10 !== undefined && m.p90 !== null && m.p90 !== undefined && m.p90 > m.p10;
    return Object.assign({}, m, { err: err, ape: Math.abs(err), leak: Math.abs(m.p50 - m.actual) < 1e-6,
      inside: hasRange ? (m.actual >= m.p10 && m.actual <= m.p90) : null,
      z: hasRange ? Math.abs(m.actual - m.p50) / ((m.p90 - m.p10) / 2) : null });
  });
  const use = ms.filter(m => !m.leak);
  const mean = xs => (xs.length ? xs.reduce((a, b) => a + b, 0) / xs.length : null);
  const inside = use.filter(m => m.inside !== null);
  const zs = use.map(m => m.z).filter(z => z !== null);
  // 偏り（旧来の B-5 と同じく、新しい月ほど重い）
  const w = use.map((m, i) => Math.pow(0.5, (use.length - 1 - i) / APP_BIAS_HALF_LIFE));
  const half = Math.floor(use.length / 2);
  return {
    months: ms, n: use.length, leaks: ms.length - use.length,
    mape: mean(use.map(m => m.ape)), bias: appWMean_(use.map(m => m.err), w), biasW: w.reduce((a, b) => a + b, 0),
    errVar: use.length > 1 ? mean(use.map(m => Math.pow(m.err - mean(use.map(x => x.err)), 2))) : null,
    coverage: inside.length ? inside.filter(m => m.inside).length / inside.length : null, coverageN: inside.length,
    // 実績の 80% が入る幅は、今の幅の何倍か（1 より大きければ今の幅は狭すぎる）
    widthScale: zs.length >= 3 ? appQuantile_(zs, 0.8) : null,
    trend: use.length >= 6 ? { recent: mean(use.slice(-3).map(m => m.ape)), before: mean(use.slice(0, use.length - 3).slice(-3).map(m => m.ape)) }
      : half >= 2 ? { recent: mean(use.slice(half).map(m => m.ape)), before: mean(use.slice(0, half).map(m => m.ape)) } : null
  };
}

/**
 * 全計画の偏りの補正を、全体の平均へ縮める（経験ベイズ・DerSimonian–Laird の τ²）。影: 予測には効かない。
 * 返り値: { mu, tau2, plans: { planId: { bias, shrunk, factorNow, factorShadow, mapeNow, mapeShadow } } }
 */
function appBiasShadow_(accs) {
  const ids = Object.keys(accs).filter(id => accs[id].n >= 3 && accs[id].bias !== null);
  const v = id => { const a = accs[id]; const s2 = a.errVar === null ? 0.04 : Math.max(a.errVar, 1e-4); return s2 / Math.max(1, a.biasW); };
  let mu = 0, tau2 = 0;
  if (ids.length >= 2) {
    const w = ids.map(id => 1 / v(id));
    mu = appWMean_(ids.map(id => accs[id].bias), w);
    const q = ids.reduce((s, id, i) => s + w[i] * Math.pow(accs[id].bias - mu, 2), 0);
    const c = w.reduce((a, b) => a + b, 0) - w.reduce((a, b) => a + b * b, 0) / w.reduce((a, b) => a + b, 0);
    tau2 = Math.max(0, (q - (ids.length - 1)) / (c || 1));
  }
  const plans = {};
  ids.forEach(id => {
    const a = accs[id];
    // 計画が 1 つだけのときは、旧来と同じく 0 へ縮める（事前の精度 3 = 旧来の kG）
    const shrunk = ids.length >= 2 ? (tau2 > 0 ? (a.bias / v(id) + mu / tau2) / (1 / v(id) + 1 / tau2) : mu) : a.bias * a.biasW / (a.biasW + 3);
    const factor = Math.min(1.25, Math.max(0.75, 1 - shrunk));
    const use = a.months.filter(m => !m.leak);
    const mapeAt = f => use.reduce((s, m) => s + Math.abs((m.p50 * f - m.actual) / m.actual), 0) / use.length;
    plans[id] = { bias: a.bias, shrunk: shrunk, factorShadow: factor, mapeNow: mapeAt(1), mapeShadow: mapeAt(factor) };
  });
  return { mu: mu, tau2: tau2, pooled: ids.length >= 2, plans: plans };
}

/** 計画の精度と学習の影（予測の画面の「検証」タブ） */
function appLearningView_(planId) {
  const plan = appPlanOf_(planId);
  const accs = {};
  appReadTable_('PLANS').filter(p => p.state !== 'ARCHIVED').forEach(p => { try { accs[p.plan_id] = appAccuracyCached_(p.plan_id); } catch (e) { /* 読めない計画は除く */ } });
  const mine = accs[plan.plan_id] || appAccuracyOf_(plan.plan_id);
  const shadow = appBiasShadow_(accs);
  const cal = appEngTableObjects_(plan.plan_id, ['CALIBRATION_STATE']).CALIBRATION_STATE.slice(-1)[0] || null;
  return { planId: plan.plan_id, accuracy: Object.assign({}, mine, { months: mine.months.map(m => ({ month: m.month, p10: m.p10, p50: m.p50, p90: m.p90, actual: m.actual,
    err: m.err, inside: m.inside, leak: m.leak })) }),
    shadow: shadow.plans[plan.plan_id] || null, pooledPlans: Object.keys(shadow.plans).length, pooled: shadow.pooled, mu: shadow.mu, tau2: shadow.tau2,
    factorNow: cal ? appNum_(cal.bias_correction_factor) : null };
}

// ---- 情報源の信頼度の事前分布（ベータ二項の経験ベイズ） ----

/** 全計画の当たった数から、情報源ごとの事前分布を作る（書かない。画面で見比べる） */
function appPoolPreview_() {
  const plans = appReadTable_('PLANS').filter(p => p.state !== 'ARCHIVED');
  const names = appClientNameMap_();
  const per = {};   // type → [{ planId, n, hit }]
  const current = {};
  plans.forEach(p => {
    let t;
    try { t = appEngTableObjects_(p.plan_id, ['RELIABILITY_EVIDENCE', 'POOL_PRIOR']); } catch (e) { return; }
    const sum = {};
    t.RELIABILITY_EVIDENCE.forEach(r => {
      const k = String(r.source_type || ''); const n = appNum_(r.n), h = appNum_(r.hit);
      if (!k || n === null || h === null) return;
      const o = sum[k] = sum[k] || { n: 0, hit: 0 }; o.n += n; o.hit += h;
    });
    Object.keys(sum).forEach(k => { (per[k] = per[k] || []).push({ planId: p.plan_id, client: names[p.client_id] || p.client_label, fy: p.fy, n: sum[k].n, hit: sum[k].hit }); });
    current[p.plan_id] = {};
    t.POOL_PRIOR.forEach(r => { current[p.plan_id][String(r.pool_scope)] = { value: appNum_(r.pooled_value), precision: appNum_(r.precision), nClients: appNum_(r.n_clients) }; });
  });
  const types = APP_SOURCE_TYPES.map(type => {
    const xs = (per[type] || []).filter(x => x.n >= APP_POOL_MIN_SAMPLES);
    const N = xs.reduce((s, x) => s + x.n, 0), Hh = xs.reduce((s, x) => s + x.hit, 0);
    const out = { type: type, label: APP_SOURCE_LABELS[type] || type, plans: xs.length, n: N, hit: Hh, ok: xs.length >= APP_POOL_MIN_PLANS && N > 0 };
    if (!out.ok) { out.note = '計画が ' + APP_POOL_MIN_PLANS + ' つ以上（それぞれ ' + APP_POOL_MIN_SAMPLES + ' 件以上）になると作れます'; return out; }
    const mu = Hh / N;
    // 計画どうしのばらつき（二項のばらつきを除いた分）から、事前分布の強さ（α + β）を決める
    const wv = xs.reduce((s, x) => s + x.n * Math.pow(x.hit / x.n - mu, 2), 0) / N;
    const nbar = N / xs.length;
    const between = Math.max(0, wv - mu * (1 - mu) / nbar);
    const precision = between > 0 ? Math.min(APP_POOL_PRECISION_MAX, Math.max(APP_POOL_PRECISION_MIN, mu * (1 - mu) / between - 1)) : APP_POOL_PRECISION_MAX;
    Object.assign(out, { hitRate: mu, alpha: mu * precision, beta: (1 - mu) * precision, precision: Math.round(precision * 10) / 10,
      pooledR: Math.round(Math.min(1.5, Math.max(0, 2 * mu)) * 1000) / 1000, between: between,
      perPlan: xs.map(x => ({ client: x.client, fy: x.fy, n: x.n, hitRate: x.hit / x.n,
        posterior: (x.hit + mu * precision) / (x.n + precision) })) });   // その計画の当たる割合の事後の平均（参考）
    return out;
  });
  return { types: types, plans: plans.length, current: current, ready: types.some(t => t.ok) };
}

/**
 * 全計画の POOL_PRIOR に事前分布を書く（PLAN.RUN と同じく、計算用ブックで旧来の表の形のまま書き、控えを置いてから保存）。
 * 計画が多いときは、1 回の実行の目安（4 分）を過ぎたら残りの計画を続きの処理に回す。値が同じ行は書き直さない（更新日も変えない）
 */
function appPoolApply_(ctx, p) {
  const t0 = new Date().getTime();
  const rowsFor = p && p.rowsFor ? p.rowsFor : appPoolPreview_().types.filter(t => t.ok).map(t => ({ scope: 'reliability:' + t.type, value: t.pooledR, precision: t.precision, nClients: t.plans }));
  if (!rowsFor.length) throw new Error('事前分布を作れる情報源がまだありません（計画が ' + APP_POOL_MIN_PLANS + ' つ以上要ります）。');
  const remaining = p && p.remaining ? p.remaining.slice() : appReadTable_('PLANS').filter(x => x.state !== 'ARCHIVED').map(x => x.plan_id);
  const done = (p && p.done) || [];
  const header = APP_ENGINE_SHEETS.POOL_PRIOR.header;
  const col = h => header.indexOf(h);
  while (remaining.length) {
    if (done.length > (p && p.done ? p.done.length : 0) && new Date().getTime() - t0 > APP_BUILD_BUDGET_MS) break;
    const planId = remaining.shift();
    appWithLock_(() => {
      appJournalRecover_(ctx);
      const plan = appPlanOf_(planId);
      const scratch = appWorkScratch_(plan);
      let st = null;
      do { st = appScratchBuildStep_(scratch, planId, ['POOL_PRIOR'], st && st.state, new Date().getTime() + 60000); } while (!st.complete);
      let sh = scratch.getSheetByName('POOL_PRIOR');
      if (!sh) { sh = scratch.insertSheet('POOL_PRIOR'); sh.getRange(1, 1, 1, header.length).setValues([header]); }
      const last = sh.getLastRow();
      const vals = last >= 2 ? sh.getRange(2, 1, last - 1, header.length).getValues() : [];
      let changed = 0;
      rowsFor.forEach(x => {
        const i = vals.findIndex(v => String(v[col('pool_scope')]) === x.scope && String(v[col('param_key')]) === 'reliability_r');
        const same = i >= 0 && Number(vals[i][col('pooled_value')]) === x.value && Number(vals[i][col('precision')]) === x.precision && Number(vals[i][col('n_clients')]) === x.nClients;
        if (same) return;
        const row = header.map(h => ({ pool_scope: x.scope, param_key: 'reliability_r', pooled_value: x.value, precision: x.precision, n_clients: x.nClients,
          updated_at: new Date(), updated_by: ctx.actor, note: '新アプリ: 全計画のベータ二項の経験ベイズ' })[h]);
        if (i >= 0) vals[i] = row; else vals.push(row);
        changed++;
      });
      if (changed) {
        sh.getRange(2, 1, vals.length, header.length).setValues(vals);
        const cap = appCaptureChanged_(scratch, planId, appStoredHashes_(planId), ['POOL_PRIOR']);
        if (cap.changed.length) appJournalRun_(ctx, '学習の事前分布（' + planId + '）', planId, appChangedOps_(ctx, planId, cap.changed, appId_('POOL')));
        appScratchMarkAfterSave_(scratch, planId, st.state.token, ['POOL_PRIOR'], !!st.state.reused, st.state.scope, st.state.problems);
      }
      done.push({ planId: planId, changed: changed > 0 });
    });
  }
  if (remaining.length) return { __next: { kind: 'LEARN.POOL', payload: { rowsFor: rowsFor, remaining: remaining, done: done } }, audit: { entityId: 'POOL_PRIOR' } };
  return { written: rowsFor, plans: done, audit: { entityId: 'POOL_PRIOR' } };
}
