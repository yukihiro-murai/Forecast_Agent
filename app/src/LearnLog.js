/**
 * LearnLog.js — 学びの記録（LEARNING_LOG。SCHEMA_PLAN_v10-12_JA.md の 3-5。2026-10-08 村井さん承認「版10はおすすめで」）。
 * 案（PROPOSE）・判断（DECIDE）・反映（APPLY）を、今ある操作の保存と同じ控えで 1 件ずつ足す（書き換えない。記録は計算に使わない）。
 *   - 所有者の補正の書き込み（setCalibration。CALIBRATION.SET）: 変えた項目ごとに APPLY（origin OWNER）。見直し案の取り下げは DECIDE「取り下げ」
 *   - 四半期の判断の保存（REVIEW.DECIDE）: 判断が変わった案ごとに DECIDE（origin QUARTERLY）
 *   - 承認した案の適用（REVIEW.APPLY・C-3）: C-3 が記録した判断が前の記録と違えば DECIDE、反映した案ごとに APPLY（origin REVIEW_APPLY）。
 *     自動の学びの旗が 0 で C-3 が何も反映しないときは、判断だけ（APPLY は足さない）
 *   - 見直し案を作る（REVIEW.GENERATE・C-1。今は止めている）: 新しい案ごとに PROPOSE（origin QUARTERLY）
 *   - 着地見込みの τ・w の承認（saveSetting の landing.tau・landing.w）: APPLY（origin OWNER・全計画の学びなので plan_id は空）
 * 同じ案の行は proposal_id でつなぐ（四半期の案は「レビューの番号:案の番号」、所有者の書き込みは「OWNER:操作の番号:項目」）。
 * 同じ頼みをもう一度動かしても足さない: 判断は前の記録の判断と同じなら足さず、反映は案ごとに 1 回、設定は中身から決まる番号。
 * 記録を作れなくても本体の保存は止めない（エラーのログに残す）。
 */
const APP_LEARN_DECISIONS = ['承認', '却下', '保留'];   // 四半期の判断（旧来の QUARTERLY_APPROVAL_OPTIONS）。取り下げは APP_CALIBRATION_WITHDRAWN
const APP_LEARN_SETTING_KEYS = ['landing.tau', 'landing.w'];
const APP_LEARN_NOTE_MAX = 200;

/** 計画の学びの記録の、案ごとの今の様子: { proposal_id: { decisions: DECIDE の数, decision: 最後の判断, applied, proposed } } */
function appLearnState_(planId) {
  const out = {};
  appReadPlanTable_('LEARNING_LOG', planId).sort((a, b) => a._row - b._row).forEach(r => {
    const o = out[r.proposal_id] = out[r.proposal_id] || { decisions: 0, decision: '', applied: false, proposed: false };
    if (r.event === 'DECIDE') { o.decisions++; o.decision = String(r.decision || ''); }
    if (r.event === 'APPLY') o.applied = true;
    if (r.event === 'PROPOSE') o.proposed = true;
  });
  return out;
}

/** 学びの記録の 1 行（learn_id は中身から決める。時刻と人はここで入れる） */
function appLearnRow_(ctx, o, idParts) {
  return Object.assign({ learn_id: appStableLogId_(APP_LOG_PREFIX.LEARNING_LOG, idParts), actor_email: ctx.actor, at: appNowIso_() }, o,
    { note: String(o.note || '').slice(0, APP_LEARN_NOTE_MAX) });
}

/** 判断の行（前の記録の判断と同じなら足さない。番号は案と、その案の何回目の判断かで決める） */
function appLearnDecide_(ctx, planId, state, o) {
  const s = state[o.proposal_id] || { decisions: 0, decision: '' };
  if (!o.decision || s.decision === o.decision) return null;
  state[o.proposal_id] = Object.assign({}, s, { decisions: s.decisions + 1, decision: o.decision });
  return appLearnRow_(ctx, Object.assign({ plan_id: planId, event: 'DECIDE' }, o), ['DECIDE', planId, o.proposal_id, s.decisions]);
}

/** 見出しつきの値（getValues）を見出し → 値の行に（同じ見出しが 2 つあれば前の列） */
function appLearnObjects_(values) {
  if (!values || values.length < 2) return [];
  const head = values[0].map(h => String(h || '').trim());
  return values.slice(1).map(r => { const o = {}; head.forEach((h, i) => { if (h && !Object.prototype.hasOwnProperty.call(o, h)) o[h] = r[i]; }); return o; });
}

/** 計算用ブックの表の行（シートが無ければ空） */
function appLearnSheet_(book, name) {
  const sh = book && book.getSheetByName(name);
  return sh && sh.getLastRow() >= 2 ? appLearnObjects_(sh.getDataRange().getValues()) : [];
}

/** 四半期レビューの画面の行（8 行目から: 案の番号・対象・今の値・案の値・確度・根拠・見込み・判断・戻し方。J8 = レビューの番号。旧来の writeQuarterlyReviewSheet_） */
function appLearnReviewScreen_(book) {
  const sh = book && book.getSheetByName('QUARTERLY_REVIEW');
  if (!sh || sh.getLastRow() < 8) return null;
  const v = sh.getDataRange().getValues();
  const rid = String((v[7] || [])[9] || '').trim();
  const quarter = (/(FY\d{4}-Q[1-4])/.exec(String((v[0] || [])[0] || '')) || [])[1] || '';
  return { rid: rid, quarter: quarter, rows: v.slice(7).filter(r => String(r[0] || '').trim())
    .map(r => ({ pid: String(r[0]).trim(), target: String(r[1] || ''), current: String(r[2] === null || r[2] === undefined ? '' : r[2]),
      proposed: String(r[3] === null || r[3] === undefined ? '' : r[3]), decision: String(r[7] || '').trim() })) };
}

/**
 * 計画への保存（Plan.js の appPlanEdit_。同じ控えで書く）で足す学びの行: 所有者の補正の書き込みと、四半期の判断の保存。
 * value = 保存の返り値（CALIBRATION.SET は appCalibrationSet_ の { changed, withdrawn }）、book = 保存の後の計算用ブック、args = 頼みの中身
 */
function appLearnEditOps_(ctx, plan, action, actionId, value, book, args) {
  if ((action !== 'CALIBRATION.SET' && action !== 'REVIEW.DECIDE') || !appLogReady_('LEARNING_LOG')) return [];
  try {
    const state = appLearnState_(plan.plan_id);
    const rows = [];
    if (action === 'CALIBRATION.SET') {
      const reason = String((args && args.reason) || '');
      ((value && value.changed) || []).forEach(x => {
        const pid = 'OWNER:' + actionId + ':' + x.field;
        rows.push(appLearnRow_(ctx, { plan_id: plan.plan_id, proposal_id: pid, event: 'APPLY', origin: 'OWNER', target: x.field,
          current_value: String(x.old), proposed_value: String(x.new), applied_value: String(x.new), note: reason }, ['APPLY', plan.plan_id, pid]));
      });
      const w = value && value.withdrawn;
      if (w && w.reviewId) {
        const log = {};
        appLearnSheet_(book, 'QUARTERLY_REVIEW_LOG').forEach(r => { if (String(r.review_id || '').trim() === w.reviewId) log[String(r.proposal_id || '').trim()] = r; });
        (w.before || []).forEach(b => {
          const r = log[b.proposalId] || {};
          rows.push(appLearnDecide_(ctx, plan.plan_id, state, { proposal_id: w.reviewId + ':' + b.proposalId, origin: 'OWNER', target: String(r.target_field || ''),
            current_value: String(r.current_value === undefined ? '' : r.current_value), proposed_value: String(r.proposed_value === undefined ? '' : r.proposed_value),
            decision: APP_CALIBRATION_WITHDRAWN, review_quarter: String(w.quarter || ''), note: reason }));
        });
      }
    } else {
      const sc = appLearnReviewScreen_(book);
      if (sc && sc.rid && sc.rid.indexOf(APP_CALIBRATION_WITHDRAWN + ':') !== 0) {   // 取り下げたレビューの判断は C-3 が使わない
        sc.rows.filter(x => APP_LEARN_DECISIONS.indexOf(x.decision) >= 0).forEach(x => {
          rows.push(appLearnDecide_(ctx, plan.plan_id, state, { proposal_id: sc.rid + ':' + x.pid, origin: 'QUARTERLY', target: x.target,
            current_value: x.current, proposed_value: x.proposed, decision: x.decision, review_quarter: sc.quarter }));
        });
      }
    }
    return appLogOps_('LEARNING_LOG', rows.filter(Boolean));
  } catch (e) {
    appLogError_('LEARN.LOG', e, ctx);
    return [];
  }
}

/**
 * 実行の保存（Plan.js の appPlanRunSave_）で足す学びの行: C-3（REVIEW.APPLY）の判断と反映、C-1（REVIEW.GENERATE）の案。
 * bookOf() = 実行の後の計算用ブック（この 2 つの操作のときだけ開く）。反映したかは、データ本体（実行の前）と計算用ブック（後）の applied で比べる
 */
function appLearnRunOps_(ctx, plan, action, bookOf) {
  if ((action !== 'REVIEW.APPLY' && action !== 'REVIEW.GENERATE') || !appLogReady_('LEARNING_LOG')) return [];
  try {
    const book = bookOf();
    const sc = appLearnReviewScreen_(book);
    if (!sc || !sc.rid) return [];
    const after = appLearnSheet_(book, 'QUARTERLY_REVIEW_LOG').filter(r => String(r.review_id || '').trim() === sc.rid);
    if (!after.length) return [];
    const state = appLearnState_(plan.plan_id);
    const val = v => String(v === null || v === undefined ? '' : v);
    const rows = [];
    if (action === 'REVIEW.GENERATE') {
      after.forEach(r => {
        const pid = sc.rid + ':' + String(r.proposal_id || '').trim();
        if ((state[pid] || {}).proposed) return;
        let n = null;
        try { const m = JSON.parse(String(r.diagnostic_metrics_json || '{}')); n = typeof m.n === 'number' && isFinite(m.n) ? m.n : null; } catch (e) { n = null; }
        rows.push(appLearnRow_(ctx, { plan_id: plan.plan_id, proposal_id: pid, event: 'PROPOSE', origin: 'QUARTERLY', target: val(r.target_field),
          current_value: val(r.current_value), proposed_value: val(r.proposed_value), evidence_n: n, review_quarter: val(r.quarter_label), note: val(r.rationale) },
          ['PROPOSE', plan.plan_id, pid]));
      });
      return appLogOps_('LEARNING_LOG', rows);
    }
    const before = {};
    appEngTableObjects_(plan.plan_id, ['QUARTERLY_REVIEW_LOG']).QUARTERLY_REVIEW_LOG.forEach(r => {
      if (String(r.review_id || '').trim() === sc.rid) before[String(r.proposal_id || '').trim()] = r;
    });
    after.forEach(r => {
      const id = String(r.proposal_id || '').trim(), pid = sc.rid + ':' + id;
      const base = { proposal_id: pid, origin: 'REVIEW_APPLY', target: val(r.target_field), current_value: val(r.current_value), proposed_value: val(r.proposed_value),
        review_quarter: val(r.quarter_label) };
      const status = String(r.approval_status || '').trim();
      if (APP_LEARN_DECISIONS.indexOf(status) >= 0) rows.push(appLearnDecide_(ctx, plan.plan_id, state, Object.assign({ decision: status }, base)));
      const was = before[id] ? Number(before[id].applied || 0) === 1 : false;
      if (Number(r.applied || 0) === 1 && !was && !(state[pid] || {}).applied) {
        rows.push(appLearnRow_(ctx, Object.assign({ plan_id: plan.plan_id, event: 'APPLY', applied_value: val(r.proposed_value) }, base), ['APPLY', plan.plan_id, pid]));
      }
    });
    return appLogOps_('LEARNING_LOG', rows.filter(Boolean));
  } catch (e) {
    appLogError_('LEARN.LOG', e, ctx);
    return [];
  }
}

/**
 * 業務の設定の保存（Settings.js の appSaveSetting_）で足す学びの行: 着地見込みの τ・w（所有者が承認した値。判断 29）だけ。plan_id は空（年度の控えに入らない）。
 * before = 保存の前の今の値（appSettingsCurrent_ の行）。同じ値をもう一度書いても足さない（効き始めが先の日なら、同じ頼みは同じ番号で 1 行）
 */
function appLearnSettingOps_(ctx, key, before, value, eff, note) {
  if (APP_LEARN_SETTING_KEYS.indexOf(key) < 0 || !appLogReady_('LEARNING_LOG')) return [];
  const cur = before ? String(before.value) : '';
  const isDefault = !before || !!before.isDefault;
  const today = appToday_();
  if (!isDefault && cur === String(value) && eff <= today) return [];
  const id = ['SETTING', key, String(value), eff, cur, isDefault ? 'DEFAULT' : ''];
  const row = appLearnRow_(ctx, { plan_id: '', event: 'APPLY', origin: 'OWNER', target: key, current_value: cur, proposed_value: String(value),
    applied_value: String(value), note: (eff > today ? eff + ' から。' : '') + String(note || '') }, id);
  row.proposal_id = row.learn_id;
  return appLogOps_('LEARNING_LOG', [row]);
}
