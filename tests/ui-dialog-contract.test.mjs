import { existsSync, readFileSync } from 'node:fs';
import { test } from 'node:test';
import assert from 'node:assert/strict';

const code = readFileSync('Code.gs', 'utf8');

test('settings UI is opened as a modal dialog, not a sidebar', () => {
  assert.match(code, /function openSettingsDialog\(\)/);
  assert.match(code, /createHtmlOutputFromFile\('SettingsDialog'\)/);
  assert.match(code, /\.showModalDialog\(/);
  assert.doesNotMatch(code, /showSidebar\(/);
  assert.equal(existsSync('Sidebar.html'), false);
  assert.equal(existsSync('SettingsDialog.html'), true);
});

test('dialog calls the existing settings endpoints through dialog-named wrappers', () => {
  const html = readFileSync('SettingsDialog.html', 'utf8');

  assert.match(code, /function getSettingsDialogData\(\)/);
  assert.match(code, /function saveSettingsFromDialog\(payload, confirmedReset\)/);
  assert.match(html, /\.getSettingsDialogData\(\)/);
  assert.match(html, /\.fetchHeadersFromSources\(\)/);
  assert.match(html, /\.saveSettingsFromDialog\(payload, confirmedReset\)/);
});

test('dialog has unified busy state, status messages, and destructive save confirmation', () => {
  const html = readFileSync('SettingsDialog.html', 'utf8');

  [
    'statusRegion',
    'confirmPanel',
    'saveConfirmButton',
    'saveCancelButton',
    'fieldsContainer',
    'messageTemplate',
    'setBusy',
    'showStatus',
    'showConfirmPanel'
  ].forEach((token) => {
    assert.match(html, new RegExp(token));
  });

  ['status--success', 'status--warning', 'status--error'].forEach((className) => {
    assert.match(html, new RegExp(className));
  });
});
