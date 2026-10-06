/**
 * YearSnapshot.js - Build a fiscal-year archive from raw stored text, without
 * changing the original tables. Row selection and hashing live in one place.
 */
const APP_YEAR_SNAPSHOT_FORMAT = 1;
const APP_YEAR_SNAPSHOT_MAX_BYTES = 8 * 1024 * 1024;

function appYearNumber_(fy) {
  const year = Number(fy);
  if (!/^\d{4}$/.test(String(fy)) || !Number.isInteger(year) || year < 2000 || year > 2100) {
    throw new Error('年度を確認できません。書き込みを止めました。');
  }
  return year;
}

/** Fixed read selectors shared by preview, freeze checks and the registry. */
function appYearStoredPlans_(fy) {
  const rows = appReadTable_('PLANS');
  rows.forEach(p => appYearNumber_(p.fy));
  return fy === undefined ? rows : rows.filter(p => String(p.fy) === String(appYearNumber_(fy)));
}

function appYearStoredVersions_(planIds) {
  return appReadTable_('PLAN_VERSIONS').filter(v => planIds.indexOf(v.plan_id) >= 0);
}

function appYearStoredClosures_() {
  const rows = appYearRawRows_('YEAR_CLOSURES').map(r => appRowToObject_(APP_TABLES.YEAR_CLOSURES, r));
  if (rows.some(r => !/^\d{4}$/.test(String(r.fy)) || Number(r.fy) < 2000 || Number(r.fy) > 2100)) {
    throw new Error('年度の凍結の記録を確認できません。書き込みを止めました。');
  }
  return rows;
}

function appYearRegistryWasInitialized_() {
  return appReadTable_('_SCHEMA').some(r => r.table === 'YEAR_CLOSURES');
}

function appYearStoredClosure_(fy) {
  const year = appYearNumber_(fy);
  const rows = appYearStoredClosures_().filter(r => String(r.fy) === String(year));
  if (rows.length > 1) throw new Error('年度の凍結の記録が重複しています。書き込みを止めました。');
  return rows[0] || null;
}

/** Keep stored text and column order; reject incomplete or duplicate keys. */
function appYearRawRows_(name) {
  const def = APP_TABLES[name];
  const indices = def.key.map(k => def.columns.indexOf(k));
  const seen = {};
  return appRawTable_(name).filter(r => r.some(v => v !== '')).map(r => {
    if (r.length !== def.columns.length || r.some(v => typeof v !== 'string') ||
        indices.some(i => i < 0 || r[i] === '')) {
      throw new Error('年度の控えを作れません。表のキーか列を確認してください: ' + name);
    }
    const key = JSON.stringify(indices.map(i => r[i]));
    if (Object.prototype.hasOwnProperty.call(seen, key)) {
      throw new Error('年度の控えを作れません。キーが重複しています: ' + name);
    }
    seen[key] = true;
    return r.slice();
  });
}

/** Deterministic key order, independent of physical row positions. */
function appYearSnapshotTable_(name, rows) {
  const def = APP_TABLES[name];
  const indices = def.key.map(k => def.columns.indexOf(k));
  const key = r => JSON.stringify(indices.map(i => r[i]));
  return {
    name: name,
    columns: def.columns.slice(),
    key: def.key.slice(),
    rows: rows.slice().sort((a, b) => {
      const ka = key(a), kb = key(b);
      return ka < kb ? -1 : ka > kb ? 1 : 0;
    })
  };
}

/**
 * Select all plans in the year, and every registered table with plan_id.
 * Client names are contextual copies, not frozen global masters.
 * Global roles, members, settings, logs and closure records are not year data.
 */
function appYearSnapshot_(fy) {
  const year = Number(fy);
  if (!Number.isInteger(year) || year < 2000 || year > 2100) {
    throw new Error('年度を 4 桁の数で選んでください。');
  }
  const planDef = APP_TABLES.PLANS;
  const fyCol = planDef.columns.indexOf('fy');
  const planCol = planDef.columns.indexOf('plan_id');
  const clientCol = planDef.columns.indexOf('client_id');
  const plans = appYearRawRows_('PLANS').filter(r => Number(r[fyCol]) === year);
  if (!plans.length) throw new Error('この年度の計画はありません。');
  if (plans.some(r => !/^\d{4}$/.test(r[fyCol]) || !r[clientCol])) {
    throw new Error('年度かメーカーが正しくない計画があります。');
  }
  const planIds = plans.map(r => r[planCol]);
  const clientIds = plans.map(r => r[clientCol]).filter((v, i, a) => a.indexOf(v) === i);
  const tables = [appYearSnapshotTable_('PLANS', plans)];
  Object.keys(APP_TABLES).filter(name => name !== 'PLANS' &&
      APP_TABLES[name].columns.indexOf('plan_id') >= 0).sort().forEach(name => {
    const col = APP_TABLES[name].columns.indexOf('plan_id');
    const rows = appYearRawRows_(name).filter(r => planIds.indexOf(r[col]) >= 0);
    tables.push(appYearSnapshotTable_(name, rows));
  });
  ['CLIENTS', 'CLIENT_NAMES'].forEach(name => {
    const col = APP_TABLES[name].columns.indexOf('client_id');
    const rows = appYearRawRows_(name).filter(r => clientIds.indexOf(r[col]) >= 0);
    if (name === 'CLIENTS' && clientIds.some(id => !rows.some(r => r[col] === id))) {
      throw new Error('計画に対応するメーカーが見つかりません。');
    }
    tables.push(appYearSnapshotTable_(name, rows));
  });
  tables.sort((a, b) => a.name < b.name ? -1 : a.name > b.name ? 1 : 0);
  const snapshot = {
    format: APP_YEAR_SNAPSHOT_FORMAT,
    fy: year,
    schemaVersion: APP_SCHEMA_VERSION,
    appVersion: APP_VERSION,
    tables: tables
  };
  const text = JSON.stringify(snapshot);
  const bytes = appUtf8Bytes_(text);
  if (bytes > APP_YEAR_SNAPSHOT_MAX_BYTES) {
    throw new Error('年度の控えが 8MB を超えます。元の行は変更せずに止めました。');
  }
  return {
    fy: year, text: text, sha256: appSha256Hex_(text), bytes: bytes,
    planIds: planIds.slice().sort(), planCount: planIds.length,
    tables: tables.map(t => ({ name: t.name, rows: t.rows.length })),
    rowCount: tables.reduce((n, t) => n + t.rows.length, 0)
  };
}

/** Verify the exact saved text, including headers, type tags and formulas. */
function appYearVerifySnapshot_(expected, file) {
  if (file.isTrashed()) throw new Error('年度の控えがゴミ箱にあります。凍結しません。');
  const back = file.getBlob().getDataAsString('UTF-8');
  if (back !== expected.text || appSha256Hex_(back) !== expected.sha256) {
    throw new Error('年度の控えが元のデータと一致しません。凍結しません。');
  }
  const decoded = JSON.parse(back);
  if (decoded.format !== APP_YEAR_SNAPSHOT_FORMAT || decoded.fy !== expected.fy) {
    throw new Error('年度の控えの形式か年度が違います。凍結しません。');
  }
  return true;
}
