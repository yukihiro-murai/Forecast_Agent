/**
 * Calibration.js — 所有者が承認した補正の値を、計画の旧来の CALIBRATION_STATE に書く（2026-10-07 村井さん決定 D1）。
 * 自動の学び（B-5・C-3）を止める旗（D2）・補正を無しに戻す（D7）・AI の効きを止める（D8）は、ここで書く。
 * あわせて、締まっていない月から作られた四半期の見直し案を取り下げられる（D3）。
 *
 * - 所有者が apiOwnerTask の setCalibration で頼む（OwnerTask.js）。中身は計画への保存（PLAN.EDIT の CALIBRATION.SET。所有者だけ）と同じ道で、
 *   使うシートだけを計算用ブックに組み立て、旧来の関数（readCalibrationState_・writeCalibrationState_・appendCalibrationHistory_）を
 *   そのまま動かし、変わったシートを控え（Journal.js）を置いてからデータ本体へ戻す。年度の締め・入力のハッシュ・記録は、ほかの保存と同じ。
 * - 変えた項目ごとに CALIBRATION_HISTORY へ 1 行を足す（B-5・C-3 と同じ形。review_id は OWNER-APPROVED）。履歴は消さない。
 * - 取り下げ: 一番新しい四半期レビュー（選び方は Insights.js の appInsightPending_ と同じ）のどの案もまだ反映していなければ、
 *   QUARTERLY_REVIEW_LOG のその行を「取り下げ」にする（判断の日時と判断した人を入れる。applied は 0 のまま。行は消さない）。
 *   判断が保留・却下・承認でも、反映していなければ取り下げる（C-3 は保留・却下の案にも判断の日時を入れ、旗が 0 なら承認した案も反映しない。
 *   どれも、判断し直して C-3 をもう一度動かせば反映できてしまうため）。前の判断と日時は結果（withdrawn.before）に残す。
 *   画面の行（QUARTERLY_REVIEW）の判断も「取り下げ」にし、C-3 が適用する review_id の欄（8 行目の 10 列目）に「取り下げ:」を付ける
 *   （C-3 はその review_id の記録を見つけられず、何も適用しない。後から判断を保存し直しても同じ）。
 * - 同じ頼みをもう一度動かしても、変わる値が無ければ何も書かない（履歴も足さない）。全部の案をもう取り下げていれば、取り下げも何もしない。
 * - 旧来の A-9 は、係数と月ごとの補正を月の P10/P50/P90 にだけ掛け、年度の P10/P50/P90 には掛けない。書いた後（見る: 今）の値で係数が 1 以外か、
 *   月ごとの補正があれば、結果の warnings にそのことを書く（止めない。appCalibrationWarnings_）。
 */

/** 書ける項目（旧来の CALIBRATION_STATE の列の名前） */
const APP_CALIBRATION_FIELDS = ['bias_correction_factor', 'residual_month_bias_json', 'auto_update_enabled', 'ai_weight_override'];
// 範囲は旧来の計算と同じ（app-calibration.test.mjs が Forecast_Agent.js の定数と比べる）
const APP_CALIBRATION_FACTOR_MIN = 0.75;   // AUTOLEARN_GLOBAL_FACTOR_MIN（補正の係数の下限）
const APP_CALIBRATION_FACTOR_MAX = 1.25;   // AUTOLEARN_GLOBAL_FACTOR_MAX（補正の係数の上限）
const APP_CALIBRATION_MONTH_BIAS_MAX = 0.20;   // AUTOLEARN_MONTH_BIAS_CAP（暦月の補正を残すときの上限）
const APP_CALIBRATION_AI_WEIGHT_MAX = 0.01;   // CONFIG の AI_WEIGHT の上限（readModelTuningFromConfig_）
/** 履歴の review_id（月次の自動学習は AUTO-MONTHLY、四半期レビューはその review_id） */
const APP_CALIBRATION_REVIEW_ID = 'OWNER-APPROVED';
const APP_CALIBRATION_ROLLBACK = 'apiOwnerTask の setCalibration で旧値（old_value）を書き直す';
/** 取り下げた見直し案の判断（QUARTERLY_REVIEW_LOG の approval_status と、画面の行の判断） */
const APP_CALIBRATION_WITHDRAWN = '取り下げ';
const APP_CALIBRATION_REASON_MAX = 200;
/** 年度の P10/P50/P90 に補正が掛からないことの注意（旧来の A-9 は月の P10/P50/P90 にだけ掛ける） */
const APP_CALIBRATION_ANNUAL_WARNING = '係数が 1 以外か、月ごとの補正があると、年度の P10/P50/P90（ホームの中心・公式版の中心）には掛からず、月の合計とずれます。';

// ---- 頼みの確かめ ----

/**
 * 頼みの中身（set・withdrawPendingReview・reason）を確かめて、書く値の形にする。返り値 { set: { 項目: 値 }, withdraw, reason }。
 * 知らない項目・範囲の外・形の違う値・理由が無い・何もすることが無い、なら何もせずに止める
 */
function appCalibrationCheck_(input) {
  const i = input || {};
  const set = i.set === undefined || i.set === null ? {} : i.set;
  if (typeof set !== 'object' || Array.isArray(set)) throw new Error('set は {"項目": 値} の形にしてください。');
  const unknown = Object.keys(set).filter(k => APP_CALIBRATION_FIELDS.indexOf(k) < 0);
  if (unknown.length) throw new Error('書けない項目です: ' + unknown.join(', ') + '。書ける項目: ' + APP_CALIBRATION_FIELDS.join(', '));
  const out = {};
  Object.keys(set).forEach(k => { out[k] = appCalibrationValue_(k, set[k]); });
  if (i.withdrawPendingReview !== undefined && typeof i.withdrawPendingReview !== 'boolean') {
    throw new Error('withdrawPendingReview は true か false で入れてください（"" で囲まない）。');
  }
  const withdraw = i.withdrawPendingReview === true;
  const reason = typeof i.reason === 'string' ? i.reason.trim() : '';
  if (!reason) throw new Error('理由（reason）を入れてください（例: "2026-10-07 の決定"）。');
  if (reason.length > APP_CALIBRATION_REASON_MAX) throw new Error('理由（reason）は ' + APP_CALIBRATION_REASON_MAX + ' 字までにしてください。');
  if (!Object.keys(out).length && !withdraw) throw new Error('書く値（set）も、見直し案の取り下げ（withdrawPendingReview）もありません。');
  return { set: out, withdraw: withdraw, reason: reason };
}

/** 項目ごとに値を確かめ、比べる形にする（係数・旗・AI の効きは数、暦月の補正は並べ直した JSON の文字。AI の効きの "" は CONFIG の値に戻す） */
function appCalibrationValue_(field, v) {
  const num = typeof v === 'number' && isFinite(v);
  if (field === 'bias_correction_factor') {
    if (!num || v < APP_CALIBRATION_FACTOR_MIN || v > APP_CALIBRATION_FACTOR_MAX) {
      throw new Error('bias_correction_factor は ' + APP_CALIBRATION_FACTOR_MIN + '〜' + APP_CALIBRATION_FACTOR_MAX + ' の数で入れてください（1 = 補正なし。"" で囲まない）。');
    }
    return v;
  }
  if (field === 'auto_update_enabled') {
    if (v !== 0 && v !== 1) throw new Error('auto_update_enabled は 0（自動の学びを止める）か 1（動かす）で入れてください（"" で囲まない）。');
    return v;
  }
  if (field === 'ai_weight_override') {
    if (v === '') return '';
    if (!num || v < 0 || v > APP_CALIBRATION_AI_WEIGHT_MAX) {
      throw new Error('ai_weight_override は 0〜' + APP_CALIBRATION_AI_WEIGHT_MAX + ' の数で入れてください（0 = AI の効きを止める。"" で CONFIG の値に戻す）。');
    }
    return v;
  }
  return appCalibrationMonthJson_(v);
}

/** 暦月の補正を確かめて並べ直す（{"4":0.05,"9":-0.1}。月は 1〜12、値は ±0.2 まで）。空（"" や {}）は "{}"（補正なし） */
function appCalibrationMonthJson_(v) {
  let o = v;
  if (typeof v === 'string') {
    const t = v.trim();
    if (!t) o = {};
    else {
      try { o = JSON.parse(t); } catch (e) { throw new Error('residual_month_bias_json の JSON が読めません。"{}" か "{\\"4\\":0.05}" の形にしてください。'); }
    }
  }
  if (!o || typeof o !== 'object' || Array.isArray(o)) throw new Error('residual_month_bias_json は {}（補正なし）か {"4":0.05}（月 → 補正）の形にしてください。');
  const keys = Object.keys(o);
  keys.forEach(k => {
    if (!/^(?:[1-9]|1[0-2])$/.test(k)) throw new Error('residual_month_bias_json の月は 1〜12 にしてください: ' + k);
    const x = o[k];
    if (typeof x !== 'number' || !isFinite(x) || Math.abs(x) > APP_CALIBRATION_MONTH_BIAS_MAX) {
      throw new Error('residual_month_bias_json の値は ±' + APP_CALIBRATION_MONTH_BIAS_MAX + ' までの数にしてください（' + k + ' 月）。');
    }
  });
  const parts = {};
  keys.map(Number).sort((a, b) => a - b).forEach(m => { parts[String(m)] = o[String(m)]; });
  return JSON.stringify(parts);
}

// ---- 今の値（旧来の読み方と同じに直す） ----

/**
 * CALIBRATION_STATE のセルを、比べる形にする（旧来の readCalibrationState_・applyCalibrationToTuning_ と同じ読み方）。
 * 係数: 空・読めない・0 以下は 1。旗: 空・読めないは 1、1 でなければ 0。AI の効き: 空・読めないは ""（CONFIG の値）。暦月の補正: 並べ直した JSON
 */
function appCalibrationNorm_(field, raw) {
  const blank = raw === '' || raw === null || raw === undefined;
  const n = blank ? NaN : Number(raw);
  if (field === 'bias_correction_factor') return isFinite(n) && n > 0 ? n : 1;
  if (field === 'auto_update_enabled') return !isFinite(n) || n === 1 ? 1 : 0;
  if (field === 'ai_weight_override') return isFinite(n) ? n : '';
  return appCalibrationStoredMonthJson_(raw);
}

/** 残っている暦月の補正を、旧来の parseResidualMonthBiasJson_ → canonicalMonthBiasJson_ と同じに読んで並べ直す（壊れた JSON は補正なし） */
function appCalibrationStoredMonthJson_(raw) {
  const o = {};
  try {
    const j = JSON.parse(String(raw || ''));
    if (j && typeof j === 'object' && !Array.isArray(j)) {
      Object.keys(j).forEach(k => {
        const m = Number(k); const v = Number(j[k]);
        if (m >= 1 && m <= 12 && isFinite(v)) o[String(m)] = v;
      });
    }
  } catch (e) { /* 空・壊れた JSON は補正なし（旧来の読み方と同じ） */ }
  const parts = {};
  Object.keys(o).filter(k => isFinite(Number(o[k]))).map(Number).sort((a, b) => a - b).forEach(m => { parts[String(m)] = Number(o[String(m)]); });
  return JSON.stringify(parts);
}

/**
 * QUARTERLY_REVIEW_LOG の一番新しい四半期レビュー（選び方は appInsightPending_ と同じ: reviewed_at が一番新しい review_id、同じ時刻なら下の行）。
 * values は見出しつきの行。返り値 { rid, rows: [行の番号（0 始まり・見出しが 0）], statuses: [行ごとの判断], withdrawable, idx: { 列名: 番号 } } か null。
 * withdrawable = どの案も反映していない（applied が 1 の行が無い）・まだ全部は取り下げていない。
 * 判断（保留・却下・承認）と判断の日時では決めない: C-3 は保留・却下の案にも判断の日時を入れるが、反映していない案は判断し直せば反映できるため
 */
function appCalibrationLatestReview_(values) {
  if (!values || values.length < 2) return null;
  const idx = {};
  (values[0] || []).forEach((h, i) => { const k = String(h || '').trim(); if (k && idx[k] === undefined) idx[k] = i; });
  if (idx.review_id === undefined) return null;
  const by = {};
  for (let r = 1; r < values.length; r++) {
    const rid = String(values[r][idx.review_id] || '').trim();
    if (!rid) continue;
    const o = by[rid] = by[rid] || { rid: rid, t: -Infinity, i: -1, rows: [] };
    const t = appTimeKey_(idx.reviewed_at === undefined ? '' : values[r][idx.reviewed_at]);
    if (t > o.t || (t === o.t && r > o.i)) { o.t = t; o.i = r; }
    o.rows.push(r);
  }
  const latest = Object.keys(by).map(k => by[k]).sort((a, b) => b.t - a.t || b.i - a.i)[0];
  if (!latest) return null;
  const cell = (r, k) => (idx[k] === undefined ? '' : values[r][idx[k]]);
  const statuses = latest.rows.map(r => String(cell(r, 'approval_status') || '').trim());
  const applied = latest.rows.some(r => Number(cell(r, 'applied') || 0) === 1);
  return { rid: latest.rid, rows: latest.rows, statuses: statuses, idx: idx,
    withdrawable: !applied && !statuses.every(x => x === APP_CALIBRATION_WITHDRAWN) };
}

/**
 * CALIBRATION_STATE の値（見出しつきの行）から、計画のメーカーの行の値（appCalibrationNorm_ で比べる形）。
 * メーカーの行が無く、行が 1 つだけならその行。無ければ null
 */
function appCalibrationStateValues_(values, client) {
  if (!values || values.length < 2) return null;
  const head = values[0].map(h => String(h || '').trim());
  const rows = values.slice(1).filter(r => r.some(v => v !== '' && v !== null));
  const row = rows.filter(r => String(r[head.indexOf('client')] || '').trim() === client)[0] || (rows.length === 1 ? rows[0] : null);
  if (!row) return null;
  const out = {};
  APP_CALIBRATION_FIELDS.forEach(f => { out[f] = appCalibrationNorm_(f, head.indexOf(f) < 0 ? '' : row[head.indexOf(f)]); });
  return out;
}

/**
 * 補正の値（appCalibrationStateValues_ の形）の注意。係数が 1 以外か、月ごとの補正があれば、年度の P10/P50/P90 と月の合計がずれることを書く。
 * 書くのは止めない（所有者は 0.75〜1.25・±20% を書ける）。どちらも無ければ空
 */
function appCalibrationWarnings_(values) {
  if (!values) return [];
  const f = values.bias_correction_factor;
  const month = values.residual_month_bias_json;
  return (typeof f === 'number' && Math.abs(f - 1) > 1e-12) || (month && month !== '{}') ? [APP_CALIBRATION_ANNUAL_WARNING] : [];
}

function appCalibrationIso_(v) {
  return appIsDate_(v) ? Utilities.formatDate(v, APP_TZ, "yyyy-MM-dd'T'HH:mm:ssZ") : String(v === null || v === undefined ? '' : v);
}

// ---- 書く（計画への保存 PLAN.EDIT の CALIBRATION.SET の中身） ----

/**
 * 組み立てた計算用ブック book の上で、旧来の関数で補正の値を書き、見直し案を取り下げる（データ本体には書かない。保存は appPlanEdit_ が行う）。
 * opts: { asOfMs, seed, actor }（「今」・種・頼んだ人。旧来の関数の updated_at・changed_at・updated_by・changed_by になる）。
 * 返り値は appLegacyCall_ と同じ形（value = { client, changed: [{ field, old, new }], unchanged: [{ field, value }], withdrawn, warnings }。
 * warnings = 書いた後の値の注意（appCalibrationWarnings_）
 */
function appCalibrationSet_(book, input, opts) {
  const c = appCalibrationCheck_(input);
  const fields = APP_CALIBRATION_FIELDS.filter(f => Object.prototype.hasOwnProperty.call(c.set, f));
  // 旧来の履歴の書き込み（appendCalibrationHistory_）は失敗しても黙って続けるので、表が無いときは先に止める
  if (fields.length) ['CALIBRATION_STATE', 'CALIBRATION_HISTORY'].forEach(n => {
    if (!book.getSheetByName(n)) throw new Error('この計画には ' + n + ' がありません。補正の値を書けません（何も書いていません）。');
  });
  return appLegacyWith_(book, opts, eng => {
    const cfg = book.getSheetByName('CONFIG');
    const client = cfg ? String(cfg.getRange('B2').getValue() || '').trim() : '';   // 旧来の B-5（webRunMonthlyLearn）と同じメーカーの名前
    if (!client) throw new Error('CONFIG のメーカー名（B2）が空です。補正の値を書けません。');
    const out = { client: client, changed: [], unchanged: [], withdrawn: null };
    if (fields.length) {
      const cal = eng.readCalibrationState_(client);   // 行が無ければ、旧来の既定の行を足す（旧来と同じ）
      const patch = {};
      fields.forEach(f => {
        const before = appCalibrationNorm_(f, cal[f]);
        const after = c.set[f];
        if (before === after) { out.unchanged.push({ field: f, value: after }); return; }
        patch[f] = f === 'residual_month_bias_json' && after === '{}' ? '' : after;   // 補正なしは空（B-5 と同じ）
        out.changed.push({ field: f, old: before, new: after });
      });
      if (out.changed.length) {
        const at = new Date(opts.asOfMs);
        patch.note = 'owner-approved ' + Utilities.formatDate(at, APP_TZ, 'yyyy-MM-dd HH:mm') + ': ' + c.reason;
        const hist = book.getSheetByName('CALIBRATION_HISTORY');
        const n0 = hist.getLastRow();
        eng.writeCalibrationState_(client, patch);
        const quarter = eng.quarterLabelFromYm_(Utilities.formatDate(at, APP_TZ, 'yyyy/MM'));   // B-5 と同じく、書いた日の四半期
        // 値は文字にして渡す（旧来は String(値 || '') で書くので、数の 0 が空になる）
        out.changed.forEach(x => eng.appendCalibrationHistory_(client, quarter, APP_CALIBRATION_REVIEW_ID, x.field, String(x.old), String(x.new), APP_CALIBRATION_ROLLBACK));
        if (hist.getLastRow() - n0 !== out.changed.length) throw new Error('CALIBRATION_HISTORY に履歴を足せませんでした（何も保存していません）。');
      }
    }
    if (c.withdraw) out.withdrawn = appCalibrationWithdraw_(book, opts);
    const st = book.getSheetByName('CALIBRATION_STATE');
    out.warnings = appCalibrationWarnings_(st && st.getLastRow() >= 2 ? appCalibrationStateValues_(st.getDataRange().getValues(), client) : null);
    return out;
  });
}

/**
 * 一番新しい四半期レビューのどの案もまだ反映していなければ取り下げる（行は消さない・applied は 0 のまま。もう取り下げた行は、その日時と人のまま）。
 * 返り値 { reviewId, quarter, reviewedAt, proposals（今回取り下げた案の数）, before: [{ proposalId, status, decidedAt }]（取り下げる前の判断）,
 * screen（画面の行も直したか） }。取り下げるものが無ければ null
 */
function appCalibrationWithdraw_(book, opts) {
  const log = book.getSheetByName('QUARTERLY_REVIEW_LOG');
  if (!log || log.getLastRow() < 2) return null;
  const values = log.getDataRange().getValues();
  const latest = appCalibrationLatestReview_(values);
  if (!latest || !latest.withdrawable) return null;
  const idx = latest.idx;
  const missing = ['proposal_id', 'approval_status', 'approval_decided_at', 'approval_decided_by'].filter(k => idx[k] === undefined);
  if (missing.length) throw new Error('QUARTERLY_REVIEW_LOG に列がありません（' + missing.join(', ') + '）。取り下げられません。');
  const at = new Date(opts.asOfMs);
  const rows = latest.rows.filter((r, i) => latest.statuses[i] !== APP_CALIBRATION_WITHDRAWN);
  const before = rows.map(r => ({ proposalId: String(values[r][idx.proposal_id] || ''), status: String(values[r][idx.approval_status] || '').trim(),
    decidedAt: appCalibrationIso_(values[r][idx.approval_decided_at]) }));
  rows.forEach(r => {
    log.getRange(r + 1, idx.approval_status + 1).setValue(APP_CALIBRATION_WITHDRAWN);
    log.getRange(r + 1, idx.approval_decided_at + 1).setValue(at);
    log.getRange(r + 1, idx.approval_decided_by + 1).setValue(String(opts.actor || ''));
  });
  // 画面の行: C-3 は 8 行目の 10 列目の review_id のレビューを、8 列目の判断で適用する（旧来の writeQuarterlyReviewSheet_ の置き方）
  let screen = false;
  const sh = book.getSheetByName('QUARTERLY_REVIEW');
  if (sh && sh.getLastRow() >= 8 && sh.getLastColumn() >= 10 && String(sh.getRange(8, 10).getValue() || '').trim() === latest.rid) {
    const pids = {};
    latest.rows.forEach(r => { pids[String(values[r][idx.proposal_id] || '').trim()] = true; });
    sh.getRange(8, 1, sh.getLastRow() - 7, 1).getValues().forEach((row, i) => {
      if (pids[String(row[0] || '').trim()]) sh.getRange(8 + i, 8).setValue(APP_CALIBRATION_WITHDRAWN);
    });
    sh.getRange(8, 10).setValue(APP_CALIBRATION_WITHDRAWN + ':' + latest.rid);
    screen = true;
  }
  const r0 = values[latest.rows[0]];
  return { reviewId: latest.rid, quarter: String(idx.quarter_label === undefined ? '' : r0[idx.quarter_label] || ''),
    reviewedAt: appCalibrationIso_(idx.reviewed_at === undefined ? '' : r0[idx.reviewed_at]), proposals: rows.length, before: before, screen: screen };
}

// ---- 見る（書く前に確かめる。apiOwnerTask の calibrationPreview） ----

/**
 * 計画の今の補正の値と、取り下げられる見直し案（pending: 一番新しい案のどれもまだ反映していないとき。判断が保留・却下・承認でも出す）。
 * warnings = 今の値の注意（appCalibrationWarnings_）。書かない。inputHash は setCalibration にそのまま渡せる
 */
function appCalibrationPreview_(planId) {
  const plan = appPlanOf_(planId);
  const s = appEngLoadPlanSheets_(plan.plan_id, ['CONFIG', 'CALIBRATION_STATE', 'CALIBRATION_HISTORY', 'QUARTERLY_REVIEW_LOG'], true);
  const client = s.CONFIG && s.CONFIG.values[1] ? String(s.CONFIG.values[1][1] || '').trim() : '';
  const st = s.CALIBRATION_STATE;
  const values = st ? appCalibrationStateValues_(st.values, client) : null;
  const lv = s.QUARTERLY_REVIEW_LOG ? s.QUARTERLY_REVIEW_LOG.values : [];
  const latest = appCalibrationLatestReview_(lv);
  let pending = null;
  if (latest && latest.withdrawable) {
    const at = (r, k) => (latest.idx[k] === undefined ? '' : lv[r][latest.idx[k]]);
    // decisions・decidedAt: 今の判断と判断の日時（C-3 を動かした後なら保留・却下・承認と日時が入っている。C-1 の直後は保留で日時は空）
    pending = { reviewId: latest.rid, quarter: String(at(latest.rows[0], 'quarter_label') || ''), reviewedAt: appCalibrationIso_(at(latest.rows[0], 'reviewed_at')),
      proposals: latest.rows.length, targets: latest.rows.map(r => String(at(r, 'target_field') || '')),
      decisions: latest.statuses, decidedAt: latest.rows.map(r => appCalibrationIso_(at(r, 'approval_decided_at'))) };
  }
  return { planId: plan.plan_id, fy: Number(plan.fy), frozen: appYearIsFrozen_(plan.fy), client: client,
    values: values, sheets: { state: !!st, history: !!s.CALIBRATION_HISTORY, reviewLog: !!s.QUARTERLY_REVIEW_LOG },
    pending: pending, warnings: appCalibrationWarnings_(values), inputHash: appPlanInputHash_(plan.plan_id) };
}
