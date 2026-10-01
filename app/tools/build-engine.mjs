#!/usr/bin/env node
/*
 * build-engine.mjs — 旧来の計算（リポジトリのルートの Forecast_Agent.js）を関数で包み、新アプリの app/src/LegacyEngine.js を作る。
 *
 *   node app/tools/build-engine.mjs           … 作り直す
 *   node app/tools/build-engine.mjs --check   … 作り直さずに、今のファイルが元と一致するかだけ確かめる（違えば終了コード 1）
 *
 * 中身は 1 文字も変えない（前後に包む行を足すだけ）。包むことで旧来の 329 の関数は外（ブラウザ）から呼べなくなり、
 * SpreadsheetApp・Date・Utilities などは呼ぶ側（Engine.js）が差し替えられる。app/tests が元のファイルとの一致を確かめる。
 */
import { createHash } from 'node:crypto';
import { readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const appDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const legacyPath = path.join(appDir, '..', 'Forecast_Agent.js');
const outPath = path.join(appDir, 'src', 'LegacyEngine.js');

/** 呼ぶ側が差し替えるもの（包んだ関数の引数から、同じ名前の定数として渡す） */
export const ENGINE_SERVICES = ['SpreadsheetApp', 'Date', 'Utilities', 'PropertiesService', 'UrlFetchApp', 'HtmlService'];
/** 新アプリから使う旧来の関数と定数 */
export const ENGINE_EXPORTS = [
  'VERSION', 'BUILD_STAGE', 'SHEETS', 'WEB_UI_CONFIRMS_',
  'runPhase1Forecast', 'updatePhase1EvaluationReport', 'updatePhase1Dashboard', 'updatePhase1LearningInsights',
  'runMonthlyAutoLearn_', 'runQuarterlyReview', 'applyQuarterlyProposals', 'syncSalesFromSalesInput_',
  'readCalibrationState_', 'requireStepSuccess_', 'updateProcessStatus_', 'getForecastFYStart_', 'getForecastFYEnd_'
];

export function buildEngine(src) {
  const sha = createHash('sha256').update(src, 'utf8').digest('hex');
  const m = /const VERSION = '([^']+)';/.exec(src);
  if (!m) throw new Error('Forecast_Agent.js に VERSION がありません');
  const header = [
    '/**',
    ' * LegacyEngine.js — 旧来の計算（Forecast_Agent.js）をそのまま関数で包んだもの。自動生成: app/tools/build-engine.mjs（手で編集しない）。',
    ' * 元のファイル: Forecast_Agent.js（VERSION ' + m[1] + '、SHA-256 ' + sha + '）',
    ' * 包んだ中の旧来の関数は外から呼べない。差し替えるもの（' + ENGINE_SERVICES.join('・') + '）は Engine.js の appLegacyServices_ が渡す。',
    ' */',
    'function appLegacyEngine_(__appSvc) {',
    ...ENGINE_SERVICES.map((n) => '  const ' + n + ' = __appSvc.' + n + ';'),
    '  // ===== Forecast_Agent.js ここから（変更しない） =====',
    ''
  ].join('\n');
  const footer = [
    '',
    '  // ===== Forecast_Agent.js ここまで =====',
    '  return {',
    ...ENGINE_EXPORTS.map((n) => "    " + n + ": typeof " + n + " === 'undefined' ? undefined : " + n + ','),
    "    SOURCE_SHA256: '" + sha + "'",
    '  };',
    '}',
    ''
  ].join('\n');
  return { text: header + src + footer, sha, version: m[1], headerLength: header.length, footerLength: footer.length };
}

if (import.meta.url === 'file://' + process.argv[1]) {
  const src = await readFile(legacyPath, 'utf8');
  const built = buildEngine(src);
  const cur = await readFile(outPath, 'utf8').catch(() => null);
  if (process.argv.includes('--check')) {
    if (cur !== built.text) { console.error('LegacyEngine.js が Forecast_Agent.js と一致しません（node app/tools/build-engine.mjs で作り直す）'); process.exit(1); }
    console.log('LegacyEngine.js は Forecast_Agent.js と一致（VERSION ' + built.version + '）');
  } else if (cur !== built.text) {
    await writeFile(outPath, built.text);
    console.log('LegacyEngine.js を作り直しました（VERSION ' + built.version + '、' + built.text.length + ' 文字）');
  } else {
    console.log('LegacyEngine.js は最新です');
  }
}
