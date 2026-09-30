import { readFileSync } from 'node:fs';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';

const code = readFileSync(new URL('../Code.gs', import.meta.url), 'utf8');
const sourceId = 'abcdefghijklmnopqrstuv';
const sourceUrl = `https://docs.google.com/spreadsheets/d/${sourceId}/edit`;
const plain = (value) => JSON.parse(JSON.stringify(value));

function makeSheet(initialRows) {
  const rows = initialRows.map((row) => row.slice());
  const sheet = {
    rows,
    getLastRow() {
      let length = rows.length;
      while (length && rows[length - 1].every((value) => value === '' || value == null)) length--;
      return length;
    },
    getLastColumn() {
      return Math.max(0, ...rows.map((row) => row.length));
    },
    getRange(row, column, height = 1, width = 1) {
      const read = () => Array.from({ length: height }, (_, y) =>
        Array.from({ length: width }, (_, x) => rows[row - 1 + y]?.[column - 1 + x] ?? ''));
      const write = (values) => {
        values.forEach((valuesRow, y) => {
          while (rows.length < row + y) rows.push([]);
          valuesRow.forEach((value, x) => { rows[row - 1 + y][column - 1 + x] = value; });
        });
      };
      return {
        getValues: read,
        getDisplayValues: () => read().map((values) => values.map(String)),
        getValue: () => read()[0][0],
        setValues(values) { write(values); return this; },
        setValue(value) { write([[value]]); return this; },
        clearContent() { write(Array.from({ length: height }, () => Array(width).fill(''))); return this; }
      };
    },
    getDataRange() { return this.getRange(1, 1, Math.max(this.getLastRow(), 1), Math.max(this.getLastColumn(), 1)); },
    deleteRow(row) { rows.splice(row - 1, 1); return this; },
    deleteRows(row, count) { rows.splice(row - 1, count); return this; },
    appendRow(row) { rows.splice(this.getLastRow(), 0, row.slice()); return this; },
    clear() { rows.length = 0; return this; },
    clearContents() { rows.length = 0; return this; },
    setFrozenRows() { return this; },
    autoResizeColumns() { return this; }
  };
  return sheet;
}

function harness(inputRows, cachedOverrides = {}, { cache = true } = {}) {
  const sheets = new Map();
  if (inputRows !== null) sheets.set('設定シート', makeSheet(inputRows));
  const properties = new Map();
  const context = vm.createContext({
    Logger: { log() {} },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({
        getSheetByName: (name) => sheets.get(name) || null,
        insertSheet(name) { const sheet = makeSheet([]); sheets.set(name, sheet); return sheet; }
      })
    },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (key) => properties.get(key) ?? null,
        setProperty(key, value) { properties.set(key, String(value)); },
        deleteProperty(key) { properties.delete(key); }
      })
    }
  });
  vm.runInContext(code, context, { filename: 'Code.gs' });
  const cached = {
    ...plain(context.buildDefaultConfig_()),
    reminderTo: ['previous@example.com'],
    formSheetNamePrefix: ['PreviousPrefix'],
    formSources: [{ id: sourceId, url: sourceUrl, courseId: '' }],
    ...cachedOverrides
  };
  if (cache) properties.set('APP_CONFIG', JSON.stringify(cached));
  else sheets.set('Config', makeSheet([[JSON.stringify(cached)]]));
  return { context, sheets, cached };
}

for (const [key, property] of [
  ['REMINDER_TO', 'reminderTo'],
  ['FORM_SS', 'formSources'],
  ['FORM_SHEET_PREFIX', 'formSheetNamePrefix']
]) {
  test(`present but empty ${key} clears its cached value`, () => {
    const { context } = harness([['KEY', 'VALUE', 'NOTE'], [key, '', 'intentionally cleared']]);
    assert.deepEqual(plain(context.loadConfig_()[property]), []);
  });
}

for (const [label, rows] of [
  ['missing settings sheet', null],
  ['missing setting keys', [['KEY', 'VALUE', 'NOTE'], ['UNRELATED', 'value', 'keep']]]
]) {
  for (const cache of [true, false]) {
    test(`${label} preserves ${cache ? 'ScriptProperties' : 'Config sheet'} migration fallback`, () => {
      const { context, cached } = harness(rows, {}, { cache });
      const loaded = plain(context.loadConfig_());
      for (const property of ['reminderTo', 'formSources', 'formSheetNamePrefix']) {
        assert.deepEqual(loaded[property], cached[property]);
      }
    });
  }
}

test('nonempty setting rows override the cached values and ignore extra blank rows', () => {
  const { context } = harness([
    ['KEY', 'VALUE', 'NOTE'],
    ['REMINDER_TO', 'current@example.com', ''],
    ['REMINDER_TO', '', ''],
    ['FORM_SHEET_PREFIX', 'CurrentPrefix', ''],
    ['FORM_SS', sourceUrl, '']
  ]);
  const loaded = plain(context.loadConfig_());
  assert.deepEqual(loaded.reminderTo, ['current@example.com']);
  assert.deepEqual(loaded.formSheetNamePrefix, ['CurrentPrefix']);
  assert.equal(loaded.formSources[0].id, sourceId);
});

for (const cache of [true, false]) {
  test(`internal save preserves existing ${cache ? 'ScriptProperties' : 'Config'} values`, () => {
    const { context, cached } = harness(null, {
      reminderTo: ['first@example.invalid', 'second@example.invalid'],
      formSheetNamePrefix: ['回答', 'Responses']
    }, { cache });
    context.saveConfig_(context.loadConfig_());
    const loaded = plain(context.loadConfig_());
    for (const property of ['reminderTo', 'formSources', 'formSheetNamePrefix']) {
      assert.deepEqual(loaded[property], cached[property]);
    }
  });
}

test('an intentionally blank prefix never selects arbitrary sheets and cannot be saved', () => {
  const { context } = harness([
    ['KEY', 'VALUE', 'NOTE'],
    ['FORM_SHEET_PREFIX', '', 'intentionally cleared']
  ]);
  assert.equal(context.sheetMatchesPrefixes_('Unrelated confidential sheet', []), false);
  assert.equal(context.sheetMatchesPrefixes_('Unrelated confidential sheet', ['']), false);
  assert.throws(() => context.validateAppConfig_(context.loadConfig_()), /接頭辞|PREFIX|対象シート/);
});

const writeFixtures = [
  ['grouped keys', [
    ['KEY', 'VALUE', 'NOTE'],
    ['REMINDER_TO', 'current@example.com', 'recipient note'],
    ['REMINDER_TO', '', ''],
    ['FORM_SHEET_PREFIX', 'CurrentPrefix', 'prefix note'],
    ['FORM_SS', sourceUrl, 'source note'],
    ['UNRELATED', 'untouched', 'custom note']
  ]],
  ['interleaved keys', [
    ['KEY', 'VALUE', 'NOTE'],
    ['REMINDER_TO', 'current@example.com', 'recipient note'],
    ['UNRELATED', 'untouched', 'custom note'],
    ['FORM_SHEET_PREFIX', 'CurrentPrefix', 'prefix note'],
    ['REMINDER_TO', '', ''],
    ['FORM_SS', sourceUrl, 'source note'],
    ['REMINDER_TO', 'obsolete@example.com', ''],
    ['FORM_SHEET_PREFIX', '', ''],
    ['FORM_SS', '', ''],
    ['ANOTHER_KEY', 'another value', 'another note']
  ]]
];

for (const [label, rows] of writeFixtures) {
  test(`internal save preserves ${label} legacy contents until verified retirement`, () => {
    const { context, sheets } = harness(rows);
    const config = context.loadConfig_();
    context.saveConfig_(config);
    assert.deepEqual(sheets.get('設定シート').rows, rows);
    assert.deepEqual(plain(context.loadConfig_()), plain(config));
    sheets.get('設定シート').rows[1][1]='stale@example.com';
    assert.deepEqual(plain(context.loadConfig_()), plain(config));
  });
}

const fixedSendHeaders = ['元SS_ID', '元シート名', '元行番号', 'No', 'メールアドレス', '名前', 'クラス名', 'コースID', 'studentId', '送信状態'];
for (const replyBodyHeader of fixedSendHeaders) {
  test(`reply body header cannot overwrite fixed send column ${replyBodyHeader}`, () => {
    const { context } = harness(null);
    const config = context.normalizeAppConfig_({
      ...plain(context.buildDefaultConfig_()), formSources: [sourceUrl], replyBodyHeader
    });
    assert.throws(() => context.validateAppConfig_(config), /重複|列名|返信本文/);
  });
}

test('a distinct reply body header remains valid', () => {
  const { context } = harness(null);
  const config = context.normalizeAppConfig_({
    ...plain(context.buildLegacyDefaultConfig_()), formSources: [sourceUrl], replyBodyHeader: '生徒向け本文'
  });
  assert.doesNotThrow(() => context.validateAppConfig_(config));
  assert.ok(plain(context.getConfiguredSendHeaders_(config)).includes('生徒向け本文'));
});
