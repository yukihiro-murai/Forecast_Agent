#!/usr/bin/env node

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

function makeSheet(initialHeaders, body) {
  const rows = [initialHeaders.slice()].concat((body || []).map(function (row) { return row.slice(); }));
  return {
    maxColumns: Math.max(1, ...rows.map(function (row) { return row.length; })),
    lastColumn: Math.max(1, ...rows.map(function (row) { return row.length; })),
    lastRow: rows.length,
    values: rows,
    getMaxColumns() { return this.maxColumns; },
    getLastColumn() { return this.lastColumn; },
    getLastRow() { return this.lastRow; },
    insertColumnsAfter(count, amount) { this.maxColumns += amount; this.lastColumn += amount; },
    getRange(row, col, numRows, numCols) {
      const self = this;
      const api = {
        getValues() {
          const out = [];
          for (let r = 0; r < numRows; r++) {
            const line = [];
            for (let c = 0; c < numCols; c++) {
              const source = self.values[row - 1 + r] || [];
              line.push(source[col - 1 + c] === undefined ? '' : source[col - 1 + c]);
            }
            out.push(line);
          }
          return out;
        },
        setValues(next) {
          for (let r = 0; r < next.length; r++) {
            const targetRow = row - 1 + r;
            if (!self.values[targetRow]) self.values[targetRow] = [];
            for (let c = 0; c < next[r].length; c++) self.values[targetRow][col - 1 + c] = next[r][c];
          }
          self.lastRow = Math.max(self.lastRow, row - 1 + next.length);
          self.lastColumn = Math.max(self.lastColumn, col - 1 + Math.max(...next.map(function (line) { return line.length; })));
          return api;
        },
        clearContent() {
          for (let r = 0; r < numRows; r++) {
            const targetRow = row - 1 + r;
            if (!self.values[targetRow]) continue;
            for (let c = 0; c < numCols; c++) self.values[targetRow][col - 1 + c] = '';
          }
          return api;
        },
        setFontWeight() { return api; },
        setBackground() { return api; }
      };
      return api;
    },
    setFrozenRows() {}
  };
}

const sandbox = {
  console,
  VN_ADMIN_PORTAL_REQUEST_SHEET: 'VN_PORTAL_REQUEST',
  VN_ADMIN_PORTAL_DIRECTORY_SHEET: 'VN_PORTAL_DIRECTORY',
  VN_ADMIN_PORTAL_CLIENT_CATALOG_SHEET: 'VN_PORTAL_CLIENT_CATALOG',
  VN_ADMIN_PORTAL_REQUEST_HEADERS_V1: Object.freeze([
    'request_event_id', 'request_id', 'event_type', 'status', 'request_hash', 'request_json',
    'fiscal_year', 'client_id', 'client_name', 'forecast_owner_email',
    'related_member_emails_json', 'requested_at', 'requested_by', 'related_book_id',
    'related_book_url', 'detail_code', 'detail_message', 'created_at', 'created_by'
  ]),
  VN_ADMIN_PORTAL_REQUEST_HEADERS: Object.freeze([
    'request_event_id', 'request_id', 'event_type', 'status', 'request_hash', 'request_json',
    'fiscal_year', 'client_id', 'client_name', 'forecast_owner_email',
    'related_member_emails_json', 'requested_at', 'requested_by', 'related_book_id',
    'related_book_url', 'detail_code', 'detail_message', 'created_at', 'created_by',
    'catalog_key', 'related_member_names_json'
  ]),
  VN_ADMIN_PORTAL_DIRECTORY_HEADERS_V1: Object.freeze([
    'directory_event_id', 'directory_key', 'fiscal_year', 'client_id', 'client_name',
    'forecast_owner_email', 'related_member_emails_json', 'state', 'center_forecast',
    'adopted_forecast', 'final_budget', 'next_action', 'client_book_url', 'request_id',
    'updated_at', 'updated_by'
  ]),
  VN_ADMIN_PORTAL_DIRECTORY_HEADERS: Object.freeze([
    'directory_event_id', 'directory_key', 'fiscal_year', 'client_id', 'client_name',
    'forecast_owner_email', 'related_member_emails_json', 'state', 'center_forecast',
    'adopted_forecast', 'final_budget', 'next_action', 'client_book_url', 'request_id',
    'updated_at', 'updated_by', 'related_member_names_json'
  ]),
  VN_ADMIN_PORTAL_CLIENT_CATALOG_HEADERS: Object.freeze([
    'catalog_key', 'client_name', 'is_active', 'catalog_version', 'synced_at'
  ]),
  VN_ADMIN_PORTAL_RUNTIME_VERSION: 'vnext-portal-1.7.17',
  sheets: {},
  vNextAdminCanonicalJson_(value) { return JSON.stringify(value); },
  vNextAdminGetOrCreateSheet_(ss, name) {
    if (!ss.sheets[name]) ss.sheets[name] = makeSheet([], []);
    return ss.sheets[name];
  },
  vNextAdminEnsureExactTableHeaders_(ss, name, headers) {
    const sheet = sandbox.vNextAdminGetOrCreateSheet_(ss, name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers.slice()]);
    return { hideSheet() {} };
  }
};

vm.createContext(sandbox);
const adminSource = await readFile(path.join(root, 'VNext_Admin.js'), 'utf8');
const snippet = adminSource.slice(
  adminSource.indexOf('function vNextAdminPortalHeadersMatch_'),
  adminSource.indexOf('function vNextAdminRollbackPortalTableHeadersToV1_')
) + adminSource.slice(
  adminSource.indexOf('function vNextAdminPortalUsesV2Tables_'),
  adminSource.indexOf('/** Read-only Portal open for sidebar projections.')
);
vm.runInContext(snippet, sandbox, { filename: 'VNext_Admin.portal-migration.js' });

const ss = { sheets: {} };
sandbox.vNextAdminMigratePortalSheetHeaders_(
  ss, sandbox.VN_ADMIN_PORTAL_REQUEST_SHEET,
  sandbox.VN_ADMIN_PORTAL_REQUEST_HEADERS_V1, sandbox.VN_ADMIN_PORTAL_REQUEST_HEADERS
);
assert.equal(
  ss.sheets[sandbox.VN_ADMIN_PORTAL_REQUEST_SHEET].values[0].slice(-2).join(','),
  'catalog_key,related_member_names_json'
);

ss.sheets = {};
const shuffled = sandbox.VN_ADMIN_PORTAL_REQUEST_HEADERS.slice().reverse();
const shuffledSheet = sandbox.vNextAdminGetOrCreateSheet_(ss, sandbox.VN_ADMIN_PORTAL_REQUEST_SHEET);
shuffledSheet.values[0] = shuffled.slice();
shuffledSheet.values[1] = shuffled.map(function (header) { return header === 'client_name' ? 'Astellas' : 'x'; });
shuffledSheet.lastRow = 2;
shuffledSheet.lastColumn = shuffled.length;
sandbox.vNextAdminMigratePortalSheetHeaders_(
  ss, sandbox.VN_ADMIN_PORTAL_REQUEST_SHEET,
  sandbox.VN_ADMIN_PORTAL_REQUEST_HEADERS_V1, sandbox.VN_ADMIN_PORTAL_REQUEST_HEADERS
);
assert.deepEqual(
  ss.sheets[sandbox.VN_ADMIN_PORTAL_REQUEST_SHEET].values[0].slice(0, sandbox.VN_ADMIN_PORTAL_REQUEST_HEADERS.length),
  sandbox.VN_ADMIN_PORTAL_REQUEST_HEADERS.slice()
);
assert.equal(
  ss.sheets[sandbox.VN_ADMIN_PORTAL_REQUEST_SHEET].values[1][
    sandbox.VN_ADMIN_PORTAL_REQUEST_HEADERS.indexOf('client_name')
  ],
  'Astellas'
);

assert.equal(sandbox.vNextAdminPortalUsesV2Tables_('vnext-portal-1.7.9'), true);
assert.equal(sandbox.vNextAdminPortalUsesV2Tables_('vnext-portal-1.0.0'), false);

process.stdout.write('PASS vNext Portal table migration tests\n');
