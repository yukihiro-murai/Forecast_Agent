#!/usr/bin/env node
/**
 * app-landing.test.mjs — 年度の着地見込み（統計。締まった月の実績で今年の水準を学ぶ）と、見通しの空模様（天気）。
 * 計算だけの関数（appLandingSky_ など）を見本の数で確かめ、全部の計画の一覧（apiPortfolio）を通しても確かめる。
 *
 *   node app/tests/app-landing.test.mjs
 */
import assert from 'node:assert/strict';
import { OWNER, makeEnv, STATS, J } from './gas-mock.mjs';

const near = (a, b, tol, msg) => assert.ok(typeof a === 'number' && Math.abs(a - b) <= tol, `${msg || ''}: ${a} と ${b}`);
const yms = Array.from({ length: 12 }, (_, i) => { const m = 4 + i; return (m > 12 ? 2027 : 2026) + '/' + String(m > 12 ? m - 12 : m).padStart(2, '0'); });
const flat = (p10, p50, p90) => yms.map((ym) => ({ ym, p10, p50, p90 }));
const acts = (v) => Object.fromEntries(v.map((x, i) => [yms[i], x]));
const base = { fy: 2026, months: flat(80, 100, 120), budget: 1200 };
const pure = makeEnv();
const sky = (over) => J(pure.run('appLandingSky_(__in)', { __in: Object.assign({}, base, over) }));

// ==== 1. 着地見込みと空模様（見本の表: 月の予測 80/100/120・年間予算 1,200） ====
{
  // A: 締まった月が無い（年度の始め）→ 着地 = 予測の合計。予算どおりなので晴れのち曇り
  const a = sky({ actual: {}, cutoffYm: '2026/04' });
  assert.deepEqual([a.sky, a.skyReason, a.skyDir, a.k, a.theta, a.credibility], ['harenochi', 'ratio', '', 0, 1, 0]);
  near(a.landing, 1200, 1e-9, 'A 着地');
  near(a.landingSd, 187.943, 1e-3, 'A sd');
  near(a.landingP10, 959.141, 1e-3, 'A P10');
  near(a.landingP90, 1440.859, 1e-3, 'A P90');
  near(a.pAbove, 0.5, 1e-9, 'A 予算以上の確率');
  // P / Q: 予算だけが違う
  const p = sky({ actual: {}, cutoffYm: '2026/04', budget: 1440 });
  assert.equal(p.sky, 'kumori');
  near(p.pAbove, 0.1008, 1e-4, 'P 予算以上の確率');
  const q = sky({ actual: {}, cutoffYm: '2026/04', budget: 700 });
  assert.equal(q.sky, 'mousho');
  near(q.ratio, 1.7143, 1e-4, 'Q 着地 / 予算');
  // B: 予測どおりの 6 か月 → 着地は変わらず、ばらつきが縮む
  const b = sky({ actual: acts([100, 100, 100, 100, 100, 100]), cutoffYm: '2026/10' });
  assert.deepEqual([b.sky, b.k, b.actualYtd], ['harenochi', 6, 600]);
  near(b.landing, 1200, 1e-9, 'B 着地');
  near(b.landingSd, 56.261, 1e-3, 'B sd');
  near(b.credibility, 0.7896, 1e-4, 'B 実績の重み');
  // C: 予測の 80% の 6 か月 → 水準 θ が下がり、残りの月にも伸ばす
  const c = sky({ actual: acts([80, 80, 80, 80, 80, 80]), cutoffYm: '2026/10' });
  assert.equal(c.sky, 'kumori');
  near(c.theta, 0.8421, 1e-4, 'C θ');
  near(c.landing, 985.244, 1e-3, 'C 着地');
  near(c.ratio, 0.8210, 1e-4, 'C 着地 / 予算');
  assert.ok(c.landingP10 >= c.actualYtd && c.landingP10 < c.landing && c.landingP90 > c.landing, 'P10 < 着地 < P90');
  // C（τ = 0.25）: 事前分布が緩いと、実績の重みが増す
  const c25 = sky({ actual: acts([80, 80, 80, 80, 80, 80]), cutoffYm: '2026/10', tau: 0.25 });
  near(c25.landing, 970.502, 1e-3, 'C τ=0.25 着地');
  near(c25.credibility, 0.9125, 1e-4, 'C τ=0.25 実績の重み');
  // D / E / F: 180% → 猛暑、130% → 快晴、30% → 雨
  for (const [v, want, r] of [[180, 'mousho', 1.7159], [130, 'kaisei', 1.2684], [30, 'ame', 0.3736]]) {
    const x = sky({ actual: acts(Array(6).fill(v)), cutoffYm: '2026/10' });
    assert.equal(x.sky, want, v + '%');
    near(x.ratio, r, 1e-4, v + '% の着地 / 予算');
  }
  // G: 締まった月に実績が無い（行も無い）→ 雪
  const g = sky({ actual: {}, cutoffYm: '2026/06' });
  assert.deepEqual([g.sky, g.skyReason, g.k, g.actualYtd], ['sekka', 'zero_sales', 2, 0]);
  // G2: 行の無い締まった月は 0 円として数える → 急な落ち込み（天変地異）
  const g2 = sky({ actual: acts([100, 100]), cutoffYm: '2026/07' });
  assert.deepEqual([g2.sky, g2.skyReason, g2.skyDir, g2.k, g2.actualYtd], ['tenpen', 'shock', 'down', 3, 200]);
  near(g2.zLast, -5.53, 0.01, 'G2 z');
  // H: 予算が無い → 霧（着地の数字は出す。着地 / 予算と確率は出さない）
  const h = sky({ budget: null, actual: acts([100]), cutoffYm: '2026/05' });
  assert.deepEqual([h.sky, h.skyReason, h.ratio, h.pAbove], ['mikakunin', 'no_budget', null, null]);
  assert.ok(h.landing > 0, '予測が使えれば着地は出す');
  assert.equal(sky({ budget: 0, actual: {}, cutoffYm: '2026/06' }).skyReason, 'no_budget', '予算 0 は無いのと同じ（雪より先）');
  // I: 月の予測が無い → 霧
  const i = sky({ months: [], actual: acts([100]), cutoffYm: '2026/05' });
  assert.deepEqual([i.sky, i.skyReason, i.landing], ['mikakunin', 'no_forecast', null]);
  const holes = flat(80, 100, 120); holes[7] = { ym: yms[7], p10: null, p50: 100, p90: 120 };
  assert.equal(sky({ months: holes, actual: acts([100]), cutoffYm: '2026/05' }).skyReason, 'no_forecast', 'P10 の無い月がある');
  assert.equal(sky({ months: flat(0, 0, 0), actual: acts([100]), cutoffYm: '2026/05' }).skyReason, 'no_forecast', '予測の合計が 0');
  assert.equal(sky({ months: [], actual: {}, cutoffYm: '2026/06' }).skyReason, 'zero_sales', '雪は予測が無くても先に出す');
  // J / J2: 締まった月の実績が 3 か月以上取り込まれていない → 霧
  const j = sky({ actual: acts([100]), cutoffYm: '2026/05', todayYm: '2026/10' });
  assert.deepEqual([j.sky, j.skyReason], ['mikakunin', 'stale_actuals']);
  assert.equal(sky({ actual: acts([100, 100, 100, 100]), cutoffYm: '2026/08', todayYm: '2026/10' }).sky, 'harenochi', '2 か月の遅れは霧にしない');
  assert.equal(sky({ months: [], actual: acts([100]), cutoffYm: '2026/05', todayYm: '2026/10' }).skyReason, 'no_forecast', '予測が無いのが先');
  assert.equal(sky({ fy: 2027, months: flat(80, 100, 120).map((m) => ({ ...m, ym: String(Number(m.ym.slice(0, 4)) + 1) + m.ym.slice(4) })), actual: {}, cutoffYm: '2026/10', todayYm: '2026/10' }).sky,
    'harenochi', '次の年度の計画は霧にしない（締まった月がまだ無い）');
  // K: 予測どおりの 5 か月のあと、ひと月だけ大きく落ちる → 天変地異（shock）
  const k = sky({ actual: acts([100, 100, 100, 100, 100, 30]), cutoffYm: '2026/10' });
  assert.deepEqual([k.sky, k.skyReason, k.skyDir], ['tenpen', 'shock', 'down']);
  near(k.zLast, -4.075, 1e-3, 'K z');
  near(k.dRatio1, -0.1264, 1e-4, 'K 着地 / 予算の動き');
  assert.equal(sky({ actual: acts([100, 100, 100, 100, 100, 30]), cutoffYm: '2026/10', todayYm: '2027/02' }).skyReason, 'stale_actuals', '霧が天変地異より先');
  // L: 2 か月続けて同じ向きに外れる → 天変地異（shift）
  const l = sky({ actual: acts([100, 100, 100, 100, 140, 140]), cutoffYm: '2026/10' });
  assert.deepEqual([l.sky, l.skyReason, l.skyDir], ['tenpen', 'shift', 'up']);
  near(l.dRatio2, 0.1382, 1e-4, 'L 着地 / 予算の動き（2 か月）');
  // M / M2: 直近 2 回の予測の年間 P50 が 30% 以上違う（31 日以内）→ 天変地異（premise）
  const runs = (age0, p0 = 1700, p1 = 1200) => [{ p50: p0, ageDays: age0 }, { p50: p1, ageDays: age0 + 37 }];
  const m = sky({ actual: acts([100, 100]), cutoffYm: '2026/06', runs: runs(3) });
  assert.deepEqual([m.sky, m.skyReason, m.skyDir], ['tenpen', 'premise', 'up']);
  assert.equal(sky({ actual: acts([100, 100]), cutoffYm: '2026/06', runs: runs(40) }).sky, 'harenochi', '40 日前の変化は拾わない');
  assert.equal(sky({ actual: acts([100, 100]), cutoffYm: '2026/06', runs: runs(3, 800) }).skyDir, 'down', '下がった前提の変化');
  assert.equal(sky({ actual: acts([100, 100]), cutoffYm: '2026/06', runs: runs(3), budget: 10000 }).skyReason, 'ratio', '予算の 10% に満たない変化は拾わない');
  assert.equal(sky({ actual: acts([100, 100]), cutoffYm: '2026/06', runs: runs(3, null) }).skyReason, 'ratio', '年間 P50 の無い回は比べない');
  assert.equal(sky({ actual: {}, cutoffYm: '2026/04', runs: runs(0) }).skyReason, 'premise', '締まった月が無くても前提の変化は拾う');
  // N: 幅の倍率 w = 2 → 実績の重みが下がり、ばらつきが広がる
  const n = sky({ actual: acts([80, 80, 80]), cutoffYm: '2026/07', w: 2 });
  assert.equal(n.sky, 'kumori');
  near(n.credibility, 0.3705, 1e-4, 'N 実績の重み');
  near(n.pAbove, 0.1866, 1e-4, 'N 予算以上の確率');
  // O: 12 か月が締まった（予測が無くても）→ 着地 = 実績の合計。ばらつき 0
  const o = sky({ months: [], actual: acts([...Array(11).fill(100), 150]), cutoffYm: '2027/04' });
  assert.deepEqual([o.sky, o.landing, o.landingSd, o.landingP10, o.landingP90, o.pAbove, o.k], ['harenochi', 1250, 0, 1250, 1250, 1, 12]);
  // R: 取り込んだ月（途中の月）の実績は数えない
  const r = sky({ actual: { ...acts([100, 100, 100, 100, 100, 100]), '2026/10': 5 }, cutoffYm: '2026/10' });
  assert.deepEqual([r.actualYtd, r.landing, r.sky], [600, b.landing, 'harenochi']);
  assert.equal(sky({ actual: acts([100, 100, 100]), cutoffYm: '' }).k, 0, '境目が分からなければ締まった月は 0');
  // θ は 0〜5 に収める
  assert.equal(sky({ actual: acts(Array(6).fill(100000)), cutoffYm: '2026/10' }).theta, 5);
  // 境目（12 か月締まった年度で、着地 / 予算 = 1.5・1.1・1.0999・0.9・0.5・0.5001）
  for (const [total, want] of [[1800, 'mousho'], [1320, 'kaisei'], [1319.88, 'harenochi'], [1080, 'kumori'], [600, 'ame'], [600.12, 'kumori']]) {
    const x = sky({ months: [], actual: acts([...Array(11).fill(100), total - 1100]), cutoffYm: '2027/04' });
    assert.equal(x.sky, want, '着地 ' + total);
  }
}

// ==== 2. 全計画で学ぶ τ（今年の水準のばらつき）と w（P10〜P90 の幅の倍率） ====
{
  const mk = (f, as, p10, p90) => as.map((a) => ({ f, a, p10, p90 }));
  const prior = (plans) => J(pure.run('appLandingPrior_(__in)', { __in: plans }));
  const t = prior([mk(100, [90, 110, 100]), mk(100, [120, 130, 125]), mk(100, [70, 80, 75])]);
  near(t.tau, 0.2466, 1e-4, '水準 1.0・1.25・0.75 の 3 計画');
  assert.deepEqual([t.learned, t.plans, t.w], [true, 3, 1], '幅の倍率は月が 12 未満なら 1');
  assert.deepEqual(prior([mk(100, [90, 110, 100]), mk(100, [120, 130, 125])]).tau, 0.15, '2 計画では学ばない');
  assert.equal(prior([mk(100, [90, 110, 100]), mk(100, [120, 130, 125]), mk(100, [70, 80])]).tau, 0.15, '2 か月の計画は数えない');
  assert.equal(prior([mk(100, [99, 101, 100]), mk(100, [100, 102, 98]), mk(100, [101, 99, 100])]).tau, 0.05, '揃っていれば下限');
  assert.equal(prior([mk(100, [200, 210, 205]), mk(100, [10, 12, 11]), mk(100, [400, 390, 410])]).tau, 0.35, '上限');
  assert.equal(prior([]).tau, 0.15);
  // w: 計画ごとの水準を除いた外れ ÷ 幅の 80% 点
  const wide = [mk(100, [90, 110, 90, 110], 95, 105), mk(100, [90, 110, 90, 110], 95, 105), mk(100, [90, 110, 90, 110], 95, 105)];
  near(prior(wide).w, 2, 1e-9, '幅が半分しかない');
  assert.equal(prior(wide.map((p) => p.map((x) => ({ ...x, p10: 99, p90: 101 })))).w, 3, '上限 3');
  assert.equal(prior(wide.map((p) => p.map((x) => ({ ...x, p10: 60, p90: 140 })))).w, 1, '幅が十分なら 1');
  const noBand = prior(wide.map((p) => p.map((x) => ({ ...x, p10: null, p90: null }))));
  assert.deepEqual([noBand.w, noBand.months, noBand.plans], [1, 0, 3], 'P10/P90 の無い月は幅に使わない（水準には使う）');
}

// ==== 3. 締まった月の境目・月の予測の読み取り・予測の日数 ====
{
  const H = J(pure.run('APP_ENGINE_SHEETS.PROCESS_STATUS.header'));
  assert.equal(H[1], 'last_run_date');
  const st = (key, t, v, status = 'success') => ({ step_key: key, last_run_date: v, status, _types: 's' + t + 'sssns' });
  const cut = (rows) => pure.run('appLandingCutoff_(__in)', { __in: rows });
  assert.equal(cut([st('step2_status', 'd', '2026-10-05T01:00:00.000Z'), st('step5_status', 'd', '2026-10-05T02:00:00.000Z')]), '2026/10', 'B-1 の後に B-2');
  assert.equal(cut([st('step2_status', 'd', '2026-09-30T16:00:00.000Z'), st('step5_status', 'd', '2026-09-30T17:00:00.000Z')]), '2026/10', '日本の暦で 10 月 1 日');
  assert.equal(cut([st('step2_status', 'd', '2026-10-05T01:00:00.000Z'), st('step5_status', 'd', '2026-09-05T02:00:00.000Z')]), '', 'B-2 が古い');
  assert.equal(cut([st('step2_status', 'd', '2026-10-05T01:00:00.000Z'), st('step5_status', 'd', '2026-10-06T02:00:00.000Z', 'error')]), '', 'B-2 が失敗');
  assert.equal(cut([st('step2_status', 'd', '2026-10-05T01:00:00.000Z')]), '', 'B-2 が無い');
  assert.equal(cut([st('step2_status', 's', '2026/10/05 10:00:00'), st('step5_status', 's', '2026/10/05 10:30')]), '2026/10', '文字の日時');
  assert.equal(cut([]), '');
  assert.equal(pure.run(`appLandingAgeDays_('2026-10-03T12:00:00+0900', '2026-10-06')`), 3);
  assert.equal(pure.run(`appLandingAgeDays_('2026-10-05T15:30:00Z', '2026-10-06')`), 0, '日本の暦で同じ日');
  assert.equal(pure.run(`appLandingAgeDays_('', '2026-10-06')`), null);
  const ms = J(pure.run('appLandingMonths_(__in)', { __in: { 29: ['s2026/04', 'n80', 'n100', 'n120'], 30: ['d2026-04-30T15:00:00.000Z', 'n1', 'n2', 'n3'], 31: ['s2026/06', 'e', 'n100', 'f=B31'], 32: ['e'] } }));
  assert.deepEqual(ms, [{ ym: '2026/04', p10: 80, p50: 100, p90: 120 }, { ym: '2026/05', p10: 1, p50: 2, p90: 3 }, { ym: '2026/06', p10: null, p50: 100, p90: null }]);
  near(pure.run('appNormCdf_(1.2815515655446004)'), 0.9, 1e-6, 'Φ');
  near(pure.run('appNormCdf_(-3)'), 0.00135, 1e-5, 'Φ');
}

// ==== 4. 全部の計画の一覧を通して（OUTPUT・検証の表・手順の進みから）。今日の日付が変われば控えも変わる ====
{
  const env = makeEnv();
  env.run(`appToday_ = function () { return '2026-10-06'; }`);
  env.as(OWNER);
  env.call('apiSetup()');
  const HC = J(env.run('APP_ENGINE_SHEETS.EVAL_COMPARE_MONTHLY.header'));
  const HS = J(env.run('APP_ENGINE_SHEETS.PROCESS_STATUS.header'));
  const output = [['FY2026 売上予測']];
  for (let r = 2; r <= 25; r++) output.push([]);
  output.push(['年度合計（予測）', 1000, 1200, 1400]);
  output.push([], ['月', 'P10', 'P50', 'P90', '', '', '', '採用予測', '上乗せ']);
  yms.forEach((ym) => output.push([ym, 80, 100, 120, '', '', '', 100, '']));
  const cmpRow = (ym, act, p50) => [ym, '', '', p50, '', '', act, '', p50 === '' ? '' : 80, p50, p50 === '' ? '' : 120, '', '', p50 === '' ? '' : Math.abs(p50 - act) / act, '', '', '', '', '', p50 === '' ? '' : 0, '', '', ''];
  const book = (client, b2) => env.makeBook(client, {
    CONFIG: { values: [['項目', '値'], ['[必須] メーカー名（外部集計キー）', client], ['[必須] 予測年度FY（YYYY）', 2026], ['[必須] 担当者（カンマ区切り）', '鷹野']] },
    OUTPUT: { values: output },
    // 前の年度の月（予測なし）・締まった 6 か月（予測の 80%）・取り込んだ月の途中の実績
    EVAL_COMPARE_MONTHLY: { values: [HC, cmpRow('2026/03', 999, ''), ...yms.slice(0, 6).map((ym) => cmpRow(ym, 80, 100)), cmpRow('2026/10', 5, 100)] },
    PROCESS_STATUS: { values: [HS, ['step2_status', new Date(2026, 9, 5, 10), 'owner', 'success', client, 10, ''], ['step5_status', b2, 'owner', 'success', '', 6, '']] },
  });
  const idA = env.seedPlan(book('甲製薬', new Date(2026, 9, 5, 11)));
  const idB = env.seedPlan(book('乙製薬', new Date(2026, 8, 5, 11)));   // B-2 が B-1 より古い
  const plan = (id) => env.call('apiPortfolio()').plans.filter((p) => p.planId === id)[0];
  const a = plan(idA);
  assert.deepEqual([a.k, a.actualYtd, a.actualMonths, a.forecastYtd, a.rangeN, a.rangeOut], [6, 480, 6, 600, 6, 0], '締まった 6 か月だけ（途中の月・前の年度は数えない）');
  assert.deepEqual([a.sky, a.skyReason, a.skyDir, a.budget, a.budgetUsed, a.budgetSource], ['kumori', 'ratio', '', 1200, 1200, 'draft']);
  near(a.landing, 985.244, 1e-3, '着地（τ・w は学べないので 0.15・1）');
  near(a.ratio, 0.8210, 1e-4, '着地 / 予算');
  near(a.theta, 0.8421, 1e-4, 'θ');
  assert.ok(a.landingP10 < a.landing && a.landingP90 > a.landing && a.pAbove < 0.01 && a.credibility > 0.7);
  assert.deepEqual([a.p10, a.p50, a.p90], [1000, 1200, 1400], '年間の P10/P50/P90（まだ予測を実行していないので OUTPUT の 26 行）');
  const b = plan(idB);
  assert.deepEqual([b.k, b.actualYtd, b.actualMonths, b.sky, b.skyReason], [0, null, 0, 'mikakunin', 'stale_actuals'], 'B-2 が古ければ締まった月は数えず霧');
  near(b.landing, 1200, 1e-9, '着地は予測の合計');
  // ホームにも同じ数を渡す
  const h = env.call('apiHome()').plans.filter((p) => p.planId === idA)[0];
  for (const key of ['landing', 'landingSd', 'landingP10', 'landingP90', 'pAbove', 'ratio', 'sky', 'skyReason', 'skyDir', 'theta', 'credibility', 'k', 'budgetUsed', 'budgetSource', 'p10', 'p90']) {
    assert.deepEqual(h[key], a[key], 'ホームの ' + key);
  }
  // 承認済みの公式版があれば、その最終予算で分ける
  env.run(`appInsertRows_('PLAN_VERSIONS', [{ version_id: 'VER-T1', plan_id: '${idA}', version_no: 1, state: 'APPROVED', budget_final: 900, submitted_by: '${OWNER}', row_version: 1 }])`);
  const a2 = plan(idA);
  assert.deepEqual([a2.budgetUsed, a2.budgetSource, a2.officialFinal, a2.budget, a2.sky], [900, 'official', 900, 1200, 'harenochi']);
  near(a2.ratio, 985.244 / 900, 1e-4, '公式版の予算で割る');
  // 直近 2 回の予測の年間 P50 が大きく違う（3 日前）→ 天変地異（前提の変化）
  const run = (id, p50, at) => ({ run_id: id, plan_id: idA, status: 'DONE', annual_p10: p50 * 0.8, annual_p50: p50, annual_p90: p50 * 1.2, finished_at: at, actor_email: OWNER });
  env.run(`appInsertRows_('FORECAST_RUNS', __r)`, { __r: [run('RUN-T1', 1200, '2026-09-01T12:00:00+0900'), run('RUN-T2', 1700, '2026-10-03T12:00:00+0900')] });
  const a3 = plan(idA);
  assert.deepEqual([a3.sky, a3.skyReason, a3.skyDir, a3.p50, a3.prevP50], ['tenpen', 'premise', 'up', 1700, 1200]);
  // 同じ日のうちは控えから（表を読まない）。日付が変わると読み直す（締まった月の遅れが 3 か月 → 霧）
  env.call('apiHome()');
  const r0 = STATS.reads;
  assert.equal(plan(idA).sky, 'tenpen');
  env.call('apiHome()');
  assert.equal(STATS.reads - r0, 0, '同じ日は控えから');
  env.run(`appToday_ = function () { return '2027-01-10'; }`);
  const a4 = plan(idA);
  assert.ok(STATS.reads > r0, '日付が変われば読み直す');
  assert.deepEqual([a4.sky, a4.skyReason], ['mikakunin', 'stale_actuals'], '今日の日付で空模様が変わる');
  assert.equal(env.call('apiHome()').plans.filter((p) => p.planId === idA)[0].sky, 'mikakunin', 'ホームの控えも日ごと');
}

console.log('app-landing: all tests passed');
