import { test } from 'node:test';
import assert from 'node:assert/strict';
import { harness, sheet, plain, completeConfig } from './backend-harness.mjs';

// 数式と文字列の区別を保持する限定モック。式の評価や外部アクセスは行わない。
// getValues はエスケープ用の引用符を返さず、getFormulas は実際の式だけを返す。
function formulaAwareSheet(initial = []) {
  const s = sheet(initial), formulas = new Map(), getRange = s.getRange.bind(s);
  s.getRange = (row, col, height = 1, width = 1) => {
    const range = getRange(row, col, height, width), write = range.setValues.bind(range);
    range.getFormulas = () => Array.from({ length: height }, (_, y) =>
      Array.from({ length: width }, (_, x) => formulas.get(`${row + y}:${col + x}`) || ''));
    range.setValues = values => {
      const parsed = values.map((r, y) => r.map((value, x) => {
        const key = `${row + y}:${col + x}`;
        formulas.delete(key);
        if (typeof value === 'string' && value.startsWith("'=")) return value.slice(1);
        if (typeof value === 'string' && value.startsWith('=')) {
          formulas.set(key, value);
          return '#FORMULA!';
        }
        return value;
      }));
      write(parsed);
      return range;
    };
    range.setValue = value => range.setValues([[value]]);
    range.clearContent = () => range.setValues(Array.from({ length: height }, () => Array(width).fill('')));
    return range;
  };
  const clear = s.clearContents.bind(s);
  s.clear = s.clearContents = () => { formulas.clear(); clear(); return s; };
  return s;
}

test('append boundary quotes formula-like text without mutating values or types', () => {
  const { context: c } = harness(), date = new Date('2026-10-09T00:00:00Z');
  const values = [['=1+1', '=A1', '通常の回答', "'=1+1", "'Alice", 0, -2, false, date, '']];
  const before = values[0].slice();
  let written;
  c.appendRows_({ getLastRow: () => 2, getRange: (r, col, h, w) => {
    assert.deepEqual([r, col, h, w], [3, 1, 1, values[0].length]);
    return { setValues: rows => { written = rows; } };
  } }, values);
  assert.equal(written[0][0], "'=1+1");
  assert.equal(written[0][1], "'=A1");
  for (let i = 2; i < before.length; i++) assert.equal(written[0][i], before[i]);
  assert.deepEqual(values[0], before);
});

function preparation({ type = 'text', input = '=1+1', legacy = false, template = '＜payload＞', conversion = false } = {}) {
  const e = harness(), c = e.context, config = plain(completeConfig(c));
  config.fields.push({ key: 'payload', sourceHeader: 'Payload', evalHeader: 'Payload', sendHeader: 'Payload', type,
    ...(type === 'score_grade' ? { gradeScale: [{ from: '5', to: conversion ? '=1+1' : '優' }] } : {}) });
  config.messageTemplate = template;
  e.properties.set('APP_CONFIG', JSON.stringify(config));
  e.sheets.set('生徒一覧', sheet([
    ['No', 'メールアドレス', '名前', 'クラス名', 'コースID', 'studentId'],
    [1, 'student@example.test', '生徒', 'テスト', 'course', 'student']
  ]));
  const source = sheet([['Email', 'Name', 'Status', 'Grade', 'Payload'], ['student@example.test', '生徒', '', 5, input]]);
  source.getName = () => 'Responses 1';
  c.SpreadsheetApp.openById = () => ({ getSheets: () => [source], getSheetByName: () => source });
  if (legacy) e.sheets.set('評価データ', sheet([
    plain(c.getConfiguredEvalHeaders_(config)),
    [config.formSources[0].id, 'Responses 1', 2, 'student@example.test', '生徒', '', '5', input]
  ]));
  const target = formulaAwareSheet([plain(c.getConfiguredSendHeaders_(config))]);
  e.sheets.set('送信シート', target);
  return { ...e, c, config, source, target };
}

for (const example of [
  { input: '=1+1' },
  { input: ' \t\n=1+1' },
  { input: '=1+1', legacy: true },
  { input: '=1+1', type: 'date' },
  { input: '=1+1', type: 'score_grade' },
  { input: 5, type: 'score_grade', conversion: true }
]) test('preparation stores answer and reply as literal text: ' + JSON.stringify(example), () => {
  const e = preparation(example);
  e.c.prepareSendData();
  const headers = e.target.rows[0], row = e.target.rows[1];
  assert.equal(row[headers.indexOf('Payload')], '=1+1');
  assert.equal(row[headers.indexOf('Reply')], '=1+1');
  assert.ok(e.target.getDataRange().getFormulas().flat().every(value => value === ''));
  e.c.prepareSendData();
  assert.equal(e.target.rows.length, 2);
  assert.equal(e.target.rows[1][headers.indexOf('Payload')], '=1+1');
});

test('ordinary preparation retains the fixed reply and numeric score', () => {
  const e = preparation({ input: '授業の振り返り', template: '評価：＜grade＞\n＜payload＞' });
  e.c.prepareSendData();
  const headers = e.target.rows[0], row = e.target.rows[1];
  assert.equal(row[headers.indexOf('Reply')], '評価：5\n授業の振り返り');
  assert.equal(row[headers.indexOf('studentId')], 'student');
  assert.equal(e.source.rows[1][2], '準備〇');
});

const classHeaders = ['クラス名', 'コースID', '同期対象(1)'];
const studentHeaders = ['No', 'メールアドレス', '出席番号（任意）', '名前', 'クラス名', 'コースID', 'studentId', '除外(1)'];

test('course sync keeps a formula-like course name literal across repeated syncs', () => {
  const e = harness(), target = formulaAwareSheet([classHeaders, ['以前のクラス', 'course', 1]]);
  e.sheets.set('クラス一覧', target);
  e.context.Classroom.Courses.list = () => ({ courses: [{ id: 'course', name: '=1+1' }] });
  for (let i = 0; i < 2; i++) {
    e.context.classroomdataUnlocked_();
    assert.equal(target.rows[1][0], '=1+1');
    assert.equal(target.rows[1][2], 1);
    assert.equal(target.getRange(2, 1).getFormulas()[0][0], '');
  }
});

test('student sync protects fetched names and retained excluded names', () => {
  const e = harness(), target = formulaAwareSheet([
    studentHeaders, [1, 'old@example.test', '01', '=A1', 'テスト', 'course', 'old', '1']
  ]);
  e.sheets.set('クラス一覧', sheet([classHeaders, ['テスト', 'course', 1]]));
  e.sheets.set('生徒一覧', target);
  e.context.Classroom.Courses.Students.list = () => ({ students: [
    { profile: { id: 'new', emailAddress: 'new@example.test', name: { fullName: ' \t=1+1' } } }
  ] });
  for (let i = 0; i < 2; i++) {
    e.context.studentdataMultiUnlocked_();
    assert.equal(target.rows.find(row => row[1] === 'new@example.test')[3], '=1+1');
    assert.equal(target.rows.find(row => row[1] === 'old@example.test')[3], '=A1');
    assert.ok(target.getDataRange().getFormulas().flat().every(value => value === ''));
  }
});

test('failed roster replacement restores literal equals text and genuine formulas separately', () => {
  const e = harness(), target = formulaAwareSheet();
  target.getRange(1, 1, 2, 3).setValues([classHeaders, ["'=1+1", 'course', '=1+1']]);
  e.sheets.set('クラス一覧', target);
  target.setFrozenRows = () => { throw Error('freeze failed'); };
  assert.throws(() => e.context.replaceRosterRows_('クラス一覧', classHeaders, [['更新', 'course', 1]]), /freeze failed/);
  assert.equal(target.rows[1][0], '=1+1');
  assert.equal(target.getRange(2, 1).getFormulas()[0][0], '');
  assert.equal(target.getRange(2, 3).getFormulas()[0][0], '=1+1');
});

test('snapshot restore keeps literal values when formula metadata is unavailable', () => {
  const e = harness(), target = formulaAwareSheet([['名前'], ['=1+1']]);
  e.sheets.set('生徒一覧', target);
  const get = target.getRange.bind(target);
  target.getRange = (...args) => { const r = get(...args); delete r.getFormulas; return r; };
  const snapshot = e.context.snapshotSheetContents_(e.ss, '生徒一覧');
  target.getRange = get;
  e.context.restoreSheetContents_(e.ss, snapshot);
  assert.equal(target.rows[1][0], '=1+1');
  assert.equal(target.getRange(2, 1).getFormulas()[0][0], '');
});
