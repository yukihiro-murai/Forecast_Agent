/**
 * Directory.js — メンバー（社員名簿とメールの対応）・役割の付与・クライアントのマスタ。
 * 役割は消さずに無効化する（is_active=FALSE・valid_to）。どの変更も監査に変更前後が残る（Api.js から呼ぶ）。
 */
function appNormalizeEmail_(v) {
  return String(v || '').trim().toLowerCase();
}

function appRequireInternalEmail_(ctx, email) {
  if (!/^[^@\s]+@[^@\s]+$/.test(email)) throw new Error('メールアドレスの形式が正しくありません。');
  if (!ctx.user.domain || email.slice(-(ctx.user.domain.length + 1)) !== '@' + ctx.user.domain) {
    throw new Error('社内（@' + ctx.user.domain + '）のメールアドレスだけ登録できます。');
  }
}

function appListDirectory_(ctx) {
  const members = appReadTable_('MEMBERS').map(appStripRow_);
  const roles = appReadTable_('ROLES').map(appStripRow_);
  const clients = appReadTable_('CLIENTS').map(appStripRow_);
  return { members: members, roles: roles, clients: clients, owner: ctx.user.owner, grantable: APP_GRANTABLE_ROLES, roleLabels: APP_ROLE_LABELS };
}

/** メンバーを足す・直す。既存なら rowVersion で楽観ロック */
function appSaveMember_(ctx, input) {
  const email = appNormalizeEmail_(input && input.email);
  appRequireInternalEmail_(ctx, email);
  const name = String(input && input.displayName || '').trim();
  if (!name) throw new Error('名前を入力してください。');
  const patch = { display_name: name.slice(0, 100), department: String(input && input.department || '').trim().slice(0, 100),
    is_active: input && input.isActive === false ? false : true, note: String(input && input.note || '').slice(0, 500) };
  return appWithLock_(() => {
    const cur = appReadTable_('MEMBERS').filter(r => r.email === email)[0];
    if (cur) {
      const r = appUpdateByKey_('MEMBERS', { email: email }, patch, input && input.rowVersion, ctx.actor);
      return { member: r.after, before: r.before, created: false, audit: { entityId: email } };
    }
    const now = appNowIso_();
    const row = Object.assign({ email: email, created_at: now, created_by: ctx.actor, updated_at: now, updated_by: ctx.actor, row_version: 1 }, patch);
    appInsertRows_('MEMBERS', [row]);
    return { member: row, before: null, created: true, audit: { entityId: email } };
  });
}

/** 役割を付ける。メンバーに登録済みの人だけ。PLANNER はクライアント単位にできる */
function appGrantRole_(ctx, input) {
  const email = appNormalizeEmail_(input && input.email);
  const role = String(input && input.role || '');
  const scopeType = String(input && input.scopeType || 'ALL') === 'CLIENT' ? 'CLIENT' : 'ALL';
  const clientId = scopeType === 'CLIENT' ? String(input && input.clientId || '') : '';
  if (APP_GRANTABLE_ROLES.indexOf(role) < 0) throw new Error('付与できる役割は ' + APP_GRANTABLE_ROLES.map(r => APP_ROLE_LABELS[r]).join('・') + ' です。');
  if (scopeType === 'CLIENT' && role !== 'PLANNER') throw new Error('クライアント単位にできるのは予算策定担当だけです。');
  const validFrom = String(input && input.validFrom || '').trim() || appToday_();
  const validTo = String(input && input.validTo || '').trim();
  [validFrom, validTo].forEach(d => { if (d && !/^\d{4}-\d{2}-\d{2}$/.test(d)) throw new Error('日付は yyyy-MM-dd で入力してください。'); });
  if (validTo && validTo < validFrom) throw new Error('終わりの日が始まりの日より前です。');
  return appWithLock_(() => {
    const member = appReadTable_('MEMBERS').filter(r => r.email === email && r.is_active)[0];
    if (!member) throw new Error('先にメンバーとして登録してください（' + email + '）。');
    if (scopeType === 'CLIENT' && !appReadTable_('CLIENTS').some(c => c.client_id === clientId && c.is_active)) {
      throw new Error('クライアントを選んでください。');
    }
    // 同じ役割でも期間が重ならなければ付けられる（期限が切れた付与の後に付け直すなど）
    const end = d => d || '9999-12-31';
    const dup = appReadTable_('ROLES').some(r => r.is_active && r.email === email && r.role === role && r.scope_type === scopeType &&
      r.client_id === clientId && (r.valid_from || '0000-01-01') <= end(validTo) && validFrom <= end(r.valid_to));
    if (dup) throw new Error('同じ役割が同じ期間にすでに付いています。');
    const now = appNowIso_();
    const row = { role_id: appId_('RL'), email: email, role: role, scope_type: scopeType, client_id: clientId, valid_from: validFrom,
      valid_to: validTo, is_active: true, note: String(input && input.note || '').slice(0, 500),
      created_at: now, created_by: ctx.actor, updated_at: now, updated_by: ctx.actor, row_version: 1 };
    appInsertRows_('ROLES', [row]);
    return { role: row, audit: { entityId: row.role_id, clientId: clientId } };
  });
}

/** 役割を外す（行は残して無効にする） */
function appRevokeRole_(ctx, input) {
  const roleId = String(input && input.roleId || '');
  return appWithLock_(() => {
    const cur = appReadTable_('ROLES').filter(r => r.role_id === roleId)[0];
    if (!cur) throw new Error('役割が見つかりません。');
    if (!cur.is_active) throw new Error('この役割はすでに外れています。');
    if (cur.email === ctx.user.owner && cur.role === 'ADMIN') throw new Error('所有者は管理者のままです（外せません）。');
    const r = appUpdateByKey_('ROLES', { role_id: roleId }, { is_active: false, valid_to: appToday_() }, input && input.rowVersion, ctx.actor);
    return { role: r.after, before: r.before, audit: { entityId: roleId, clientId: cur.client_id } };
  });
}

/** 表記ゆれを吸収した比較用の名前（全角英数→半角、会社の種類の語・空白・記号を除く、小文字） */
function appNormalizeName_(s) {
  return String(s || '').normalize('NFKC')
    .replace(/株式会社|有限会社|合同会社|\(株\)|\(有\)/g, '')   // NFKC で ㈱・（株） は (株) になる
    .replace(/[\s・.,、。()「」'"’”\-‐－]/g, '')
    .toLowerCase();
}

/** クライアントを足す・直す */
function appSaveClient_(ctx, input) {
  const name = String(input && input.clientName || '').trim();
  if (!name) throw new Error('クライアント名を入力してください。');
  const normalized = appNormalizeName_(name);
  const patch = { client_name: name.slice(0, 100), zac_code: String(input && input.zacCode || '').trim().slice(0, 50),
    normalized_name: normalized, is_active: input && input.isActive === false ? false : true,
    note: String(input && input.note || '').slice(0, 500) };
  return appWithLock_(() => {
    const all = appReadTable_('CLIENTS');
    const id = String(input && input.clientId || '');
    if (all.some(c => c.client_id !== id && c.normalized_name === normalized)) throw new Error('同じ名前のクライアントがすでにあります。');
    if (id) {
      const r = appUpdateByKey_('CLIENTS', { client_id: id }, patch, input && input.rowVersion, ctx.actor);
      return { client: r.after, before: r.before, created: false, audit: { entityId: id, clientId: id } };
    }
    const now = appNowIso_();
    const row = Object.assign({ client_id: appId_('CL'), aliases_json: '[]', created_at: now, created_by: ctx.actor,
      updated_at: now, updated_by: ctx.actor, row_version: 1 }, patch);
    appInsertRows_('CLIENTS', [row]);
    return { client: row, before: null, created: true, audit: { entityId: row.client_id, clientId: row.client_id } };
  });
}
