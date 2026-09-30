import { test } from 'node:test';
import assert from 'node:assert/strict';
import { harness, sheet, plain, completeConfig } from './backend-harness.mjs';

function pipeline(count = 500, { uncachedStatusUpdates = false } = {}) {
  const h = harness(), c = h.context, config = plain(completeConfig(c));
  config.fields[0].type = 'score_grade';
  config.gradeScale = [{ from: '5', to: '優' }, { from: '4', to: '良' }];
  config.messageTemplate = '＜student_name＞：＜grade＞';
  h.properties.set('APP_CONFIG', JSON.stringify(config));
  const students = [['No', 'メールアドレス', '名前', 'クラス名', 'コースID', 'studentId']];
  const source = sheet([['Email', 'Name', 'Status', 'Grade']]);
  source.getName = () => 'Responses 1';
  for (let i = 1; i <= count; i++) {
    students.push([i, `student${i}@example.com`, `生徒${i}`, 'Class A', 'course-a', `student-${i}`]);
    source.rows.push([`student${i}@example.com`, `生徒${i}`, '', i % 2 ? 5 : 4]);
  }
  h.sheets.set('生徒一覧', sheet(students));
  const calls = { opens: 0, headers: 0, flushes: 0, posts: [], checkpoints: 0 }, missing = { sheet: false };
  const sourceRange = source.getRange.bind(source);
  source.getRange = (r, col, height = 1, width = 1) => {
    const range = sourceRange(r, col, height, width);
    if (r === 1 && height === 1) {
      const read = range.getDisplayValues;
      range.getDisplayValues = () => { calls.headers++; return read(); };
    }
    return range;
  };
  c.SpreadsheetApp.openById = id => {
    assert.equal(id, config.formSources[0].id); calls.opens++;
    return { getSheets: () => [source], getSheetByName: () => missing.sheet ? null : source };
  };
  c.SpreadsheetApp.flush = () => { calls.flushes++; };
  c.Utilities = { sleep() {} };
  c.Classroom.Courses.Announcements = { create: (data, course) => {
    const index = Number(data.individualStudentsOptions.studentIds[0].split('-')[1]);
    const send = h.sheets.get('送信シート'), statusColumn = send.rows[0].indexOf('送信状態');
    assert.equal(send.rows[index][statusColumn], '送信確認待ち');
    assert.ok(calls.flushes > calls.lastPostFlush);
    calls.lastPostFlush = calls.flushes;
    calls.checkpoints++;
    calls.posts.push({ data: plain(data), course });
    return { id: 'post-' + index };
  } };
  calls.lastPostFlush = 0;
  if (uncachedStatusUpdates) {
    const update = c.updateFormStatus_;
    // Exercise the backward-compatible uncached call path as the I/O baseline.
    c.updateFormStatus_ = (config, id, name, row, status) => update(config, id, name, row, status);
  }
  return { ...h, c, source, config, calls, missing };
}

test('500 rows retain recipients, configured content and every send checkpoint with one source open per phase', t => {
  const e = pipeline();
  e.c.importFromFormsToEval();
  const afterImport = { opens: e.calls.opens, headers: e.calls.headers };
  e.c.evalToSendSheet();
  const afterPrepare = { opens: e.calls.opens, headers: e.calls.headers };
  e.c.sendMessages();
  assert.equal(e.calls.posts.length, 500);
  assert.equal(e.calls.checkpoints, 500);
  e.calls.posts.forEach(({ data, course }, i) => {
    assert.equal(course, 'course-a');
    assert.deepEqual(data.individualStudentsOptions.studentIds, ['student-' + (i + 1)]);
    assert.ok(data.text.endsWith(`生徒${i + 1}：${i % 2 ? '良' : '優'}`));
  });
  const send = e.sheets.get('送信シート'), statusColumn = send.rows[0].indexOf('送信状態');
  assert.ok(send.rows.slice(1).every(row => row[statusColumn] === '済'));
  assert.ok(e.source.rows.slice(1).every(row => row[2] === '済'));
  assert.equal(e.calls.flushes, 1005, '1000 send flushes, 2 batch commits, 3 lock-finally flushes remain');
  assert.deepEqual(afterImport, { opens: 1, headers: 0 });
  assert.deepEqual(afterPrepare, { opens: 2, headers: 1 });
  assert.equal(e.calls.opens, 3);
  assert.equal(e.calls.headers, 2);
  const baseline = pipeline(500, { uncachedStatusUpdates: true });
  baseline.c.importFromFormsToEval(); baseline.c.evalToSendSheet(); baseline.c.sendMessages();
  assert.equal(baseline.calls.opens, 1501);
  assert.equal(baseline.calls.headers, 1500);
  assert.deepEqual(e.calls.posts, baseline.calls.posts, 'cache changes I/O counts without changing content or recipients');
  assert.equal(e.calls.checkpoints, baseline.calls.checkpoints);
  assert.equal(e.calls.flushes, baseline.calls.flushes);
  t.diagnostic('500 rows, one source: cached opens=3/header-only reads=2; uncached opens=1501/header-only reads=1500; both posts=500/checkpoints=500/flushes=1005. No elapsed-time claim.');
  e.c.sendMessages();
  assert.equal(e.calls.posts.length, 500, 'completed rows are not resent');
});

test('missing email during preparation creates an actionable error record', () => {
  const e = pipeline(1); e.c.importFromFormsToEval();
  const evaluation = e.sheets.get('評価データ');
  evaluation.rows[1][evaluation.rows[0].indexOf('メールアドレス')] = '';
  e.c.evalToSendSheet();
  assert.equal(evaluation.rows[1][evaluation.rows[0].indexOf('処理状態')], '準備×');
  assert.ok(e.sheets.get('エラー').rows.some(row => row[5] === 'CONFIG' && /メール/.test(row[6])));
  assert.equal(e.calls.posts.length, 0);
});

for (const failure of ['missing sheet', 'missing status column']) {
  test(failure + ' logs FORM_UPDATE_ERROR after successful posting and keeps sent checkpoint', () => {
    const e = pipeline(1);
    e.c.importFromFormsToEval(); e.c.evalToSendSheet();
    if (failure === 'missing sheet') e.missing.sheet = true;
    else e.source.rows[0][2] = 'Renamed';
    e.c.sendMessages(); e.c.sendMessages();
    const send = e.sheets.get('送信シート');
    assert.equal(send.rows[1][send.rows[0].indexOf('送信状態')], '済');
    assert.equal(e.calls.posts.length, 1);
    const errors = e.sheets.get('エラー').rows;
    assert.ok(errors.some(row => row[5] === 'FORM_UPDATE_ERROR'));
  });
}

test('separate send invocations reread source headers instead of reusing a stale cache', () => {
  const e = pipeline(2);
  e.c.importFromFormsToEval(); e.c.evalToSendSheet();
  const send = e.sheets.get('送信シート'), statusColumn = send.rows[0].indexOf('送信状態');
  send.rows[2][statusColumn] = '保留';
  e.c.sendMessages();
  send.rows[2][statusColumn] = '';
  e.source.rows[0][2] = 'Renamed';
  e.c.sendMessages();
  assert.equal(e.calls.posts.length, 2);
  assert.ok(e.sheets.get('エラー').rows.some(row => row[5] === 'FORM_UPDATE_ERROR'));
});
