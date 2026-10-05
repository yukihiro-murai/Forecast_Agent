/**
 * Settings.js — 業務の数値の設定表。コードには既定値と範囲だけを持ち、変えた値は SETTINGS に 1 行ずつ足す（上書きしない）。
 * 今の値 = effective_from が今日以前の行のうち最も新しいもの。無ければ既定値。
 * 既定値は旧来の仕組みの値（Forecast_Agent.js の定数・空模様の判定）と同じにしてある。
 */
const APP_SETTING_DEFS = {
  // 売上・検証用実績の取り込み（A-2・B-1）の元（ZAC の実績のスプレッドシート）。
  // 2026-10-04: 使っていなかった設定（検証の上限・空模様の判定・操作の記録）を外した。数・整数・する/しないの型は、足すときのために残す
  'source.zac_spreadsheet': { label: 'ZAC の実績のスプレッドシート', type: 'sheet', def: '', unit: 'スプレッドシートの URL（売上・実績の取り込みの元）' }
};

/** URL か ID からスプレッドシートの ID を取り出す（取り出せなければ空） */
function appParseBookId_(input) {
  const s = String(input || '').trim();
  const m = /\/spreadsheets\/d\/([a-zA-Z0-9_-]{20,})/.exec(s);
  if (m) return m[1];
  return /^[a-zA-Z0-9_-]{20,}$/.test(s) ? s : '';
}

/** 文字列や数値を設定の型に直し、範囲を確かめる。だめなら例外 */
function appParseSettingValue_(key, raw) {
  const d = APP_SETTING_DEFS[key];
  if (!d) throw new Error('未定義の設定です: ' + key);
  if (d.type === 'sheet') {
    // URL か ID。開けるスプレッドシートであることを確かめ、ID で持つ
    const id = appParseBookId_(raw);
    if (!id) throw new Error(d.label + ' は、スプレッドシートの URL で入力してください。');
    let file;
    try { file = DriveApp.getFileById(id); } catch (e) { throw new Error(d.label + ' を開けません（URL と共有を確かめてください）。'); }
    if (file.getMimeType() !== MimeType.GOOGLE_SHEETS || file.isTrashed()) throw new Error(d.label + ' は、スプレッドシートを指定してください。');
    return id;
  }
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
    // スプレッドシートの指定は、保存のときに確かめた ID をそのまま使う（読むたびに Drive を開かない）
    try { if (live) value = d.type === 'sheet' ? String(live.value) : appParseSettingValue_(key, live.value); } catch (e) { value = d.def; }
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
