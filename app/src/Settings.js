/**
 * Settings.js — 業務の数値の設定表。コードには既定値と範囲だけを持ち、変えた値は SETTINGS に 1 行ずつ足す（上書きしない）。
 * 今の値 = effective_from が今日以前の行のうち最も新しいもの。無ければ既定値。
 * 既定値は旧来の仕組みの値（Forecast_Agent.js の定数・空模様の判定）と同じにしてある。
 */
const APP_SETTING_DEFS = {
  'eval.annual_abs_error_max': { label: '年間誤差率の上限', type: 'number', min: 0, max: 1, def: 0.10, unit: '割合（0.10 = 10%）' },
  'eval.half_wape_max': { label: '半期 WAPE の上限', type: 'number', min: 0, max: 1, def: 0.12, unit: '割合' },
  'eval.overforecast_rate_max': { label: '過大予測率の上限', type: 'number', min: 0, max: 1, def: 0.05, unit: '割合' },
  'weather.min_months': { label: '空模様を判定する最少の月数', type: 'int', min: 1, max: 24, def: 3, unit: 'か月' },
  'weather.tenpen_ape': { label: '天変地異にする誤差率', type: 'number', min: 0.5, max: 10, def: 1.5, unit: '割合（1.5 = 150%）' },
  'weather.taifuu_ape': { label: '台風で数える大きな誤差率', type: 'number', min: 0, max: 5, def: 0.3, unit: '割合' },
  'audit.retention_years': { label: '操作の記録を残す年数', type: 'int', min: 1, max: 20, def: 7, unit: '年' },
  'audit.log_views': { label: '閲覧も記録する', type: 'bool', def: false, unit: 'する / しない' }
};

/** 文字列や数値を設定の型に直し、範囲を確かめる。だめなら例外 */
function appParseSettingValue_(key, raw) {
  const d = APP_SETTING_DEFS[key];
  if (!d) throw new Error('未定義の設定です: ' + key);
  if (d.type === 'bool') {
    if (raw === true || raw === 'true' || raw === 'TRUE' || raw === 1 || raw === '1') return true;
    if (raw === false || raw === 'false' || raw === 'FALSE' || raw === 0 || raw === '0') return false;
    throw new Error(d.label + ' は「する / しない」で指定してください。');
  }
  const n = Number(String(raw).trim());
  if (String(raw).trim() === '' || !isFinite(n)) throw new Error(d.label + ' は数値で入力してください。');
  if (d.type === 'int' && Math.floor(n) !== n) throw new Error(d.label + ' は整数で入力してください。');
  if (n < d.min || n > d.max) throw new Error(d.label + ' は ' + d.min + '〜' + d.max + ' の範囲で入力してください。');
  return n;
}

/** すべての設定の今の値（未設定なら既定値）と、その履歴の件数 */
function appSettingsCurrent_() {
  const rows = appIsSetUp_() ? appReadTable_('SETTINGS') : [];
  const today = appToday_();
  return Object.keys(APP_SETTING_DEFS).map(key => {
    const d = APP_SETTING_DEFS[key];
    // 効き始める日の順。同じ日なら後から足した行（シートの下の行）を新しいとする
    const hist = rows.filter(r => r.key === key && (r.scope || 'GLOBAL') === 'GLOBAL')
      .sort((a, b) => (a.effective_from < b.effective_from ? -1 : a.effective_from > b.effective_from ? 1 : a._row - b._row));
    const live = hist.filter(r => r.effective_from <= today).pop();
    let value = d.def;
    try { if (live) value = appParseSettingValue_(key, live.value); } catch (e) { value = d.def; }
    return {
      key: key, label: d.label, type: d.type, unit: d.unit, min: d.min === undefined ? null : d.min, max: d.max === undefined ? null : d.max,
      def: d.def, value: value, isDefault: !live, effectiveFrom: live ? live.effective_from : '',
      updatedAt: live ? live.created_at : '', updatedBy: live ? live.created_by : '',
      scheduled: hist.filter(r => r.effective_from > today).map(r => ({ value: r.value, effectiveFrom: r.effective_from }))
    };
  });
}

function appSettingValue_(key) {
  const s = appSettingsCurrent_().filter(x => x.key === key)[0];
  if (!s) throw new Error('未定義の設定です: ' + key);
  return s.value;
}

/** 設定を変える（新しい行を足す）。effectiveFrom 省略時は今日から */
function appSaveSetting_(ctx, input) {
  const key = String(input && input.key || '');
  const value = appParseSettingValue_(key, input && input.value);
  const eff = String(input && input.effectiveFrom || '').trim() || appToday_();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(eff)) throw new Error('効き始める日は yyyy-MM-dd で入力してください。');
  const before = appSettingsCurrent_().filter(x => x.key === key)[0];
  const row = {
    setting_id: appId_('ST'), key: key, value: String(value), scope: 'GLOBAL', scope_id: '', effective_from: eff,
    note: String(input && input.note || '').slice(0, 500), created_at: appNowIso_(), created_by: ctx.actor
  };
  appWithLock_(() => appInsertRows_('SETTINGS', [row]));
  return { saved: row, before: before, audit: { entityId: key } };
}
