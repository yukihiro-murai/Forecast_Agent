#!/usr/bin/env node

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';
import { createHash } from 'node:crypto';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

await checkAdminLearningContracts();
await checkEngineLearningWiden();

process.stdout.write('PASS vNext learning Phase A contracts\n');

async function checkAdminLearningContracts() {
  const admin = await readFile(path.join(root, 'VNext_Admin.js'), 'utf8');
  const sidebar = await readFile(path.join(root, 'VNext_AdminSidebar.html'), 'utf8');
  const engine = await readFile(path.join(root, 'VNext_Engine.js'), 'utf8');
  assert.ok(admin.includes("LEARNING_OBS: 'LEARNING_OBSERVATION'"));
  assert.ok(admin.includes("LEARNING_EVIDENCE: 'LEARNING_EVIDENCE'"));
  assert.ok(admin.includes('VN_ADMIN_LEARNING_POLICY_KEY'));
  assert.ok(admin.includes('function vNextAdminBuildLearningDashboard_'));
  assert.ok(admin.includes('function vNextAdminRecordLearningObservation'));
  assert.ok(admin.includes('function vNextAdminAppendLearningEvidenceFromEvaluation_'));
  assert.ok(admin.includes('engineRequest.learningEvidence = learningEvidence'));
  assert.ok(admin.includes('vNextAdminAppendLearningEvidenceFromEvaluation_(hub, evaluation'));
  assert.ok(sidebar.includes('適応学習（System / Budget）'));
  assert.ok(sidebar.includes('recordLearningObservation()'));
  assert.ok(sidebar.includes('vNextAdminGetLearningDashboard'));
  assert.ok(engine.includes('function vNextEngineApplyLearningEvidence_'));
  assert.ok(engine.includes('learningApplication: learningApplication.meta'));
}

async function checkEngineLearningWiden() {
  const engineSource = await readFile(path.join(root, 'VNext_Engine.js'), 'utf8');
  const sandbox = {
    console,
    Math,
    Object,
    Number,
    String,
    Boolean,
    Array,
    JSON,
    isFinite,
    Logger: { log() {} },
    Utilities: {
      DigestAlgorithm: { SHA_256: 'SHA_256' },
      computeDigest(_algorithm, value) {
        return [...createHash('sha256').update(String(value), 'utf8').digest()];
      }
    }
  };
  vm.createContext(sandbox);
  // Load only the apply helper by extracting and evaling a minimal slice is brittle;
  // instead execute the full Engine after stubbing heavy globals.
  sandbox.VNEXT_CORE = { VERSION: 'test', SCHEMA_VERSION: 'test' };
  sandbox.VNEXT_ENGINE = { VERSION: 'test' };
  sandbox.SpreadsheetApp = undefined;
  sandbox.Session = { getScriptTimeZone: () => 'Asia/Tokyo', getActiveUser: () => ({ getEmail: () => '' }) };
  sandbox.PropertiesService = { getScriptProperties: () => ({ getProperty: () => '', setProperty() {} }) };
  try {
    vm.runInContext(engineSource, sandbox);
  } catch (error) {
    // Engine may reference other GAS symbols at parse/run of top-level; if load fails, test via Function extract.
    const match = engineSource.match(/function vNextEngineApplyLearningEvidence_\([\s\S]*?\n\}/);
    assert.ok(match, 'vNextEngineApplyLearningEvidence_ must exist');
    vm.runInContext(match[0], sandbox);
  }
  assert.equal(typeof sandbox.vNextEngineApplyLearningEvidence_, 'function');
  const base = { logSigma: 0.10, annualSigmaCalibration: { method: 'TEST' } };
  const hit = sandbox.vNextEngineApplyLearningEvidence_(base, {
    evidenceId: 'LE-1', rangeContainsActual: true, intervalWidenFactorOnMiss: 1.15
  });
  assert.equal(hit.meta.applied, false);
  assert.equal(hit.history.logSigma, 0.10);
  const miss = sandbox.vNextEngineApplyLearningEvidence_(base, {
    evidenceId: 'LE-2', rangeContainsActual: false, intervalWidenFactorOnMiss: 1.2
  });
  assert.equal(miss.meta.applied, true);
  assert.ok(Math.abs(miss.history.logSigma - 0.12) < 1e-9);
  assert.equal(miss.meta.reason, 'PREVIOUS_INTERVAL_MISS_WIDEN');
}
