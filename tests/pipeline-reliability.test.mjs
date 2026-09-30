import { readFileSync } from 'node:fs';
import vm from 'node:vm';
import { test } from 'node:test';
import assert from 'node:assert/strict';

const code = readFileSync('Code.gs', 'utf8');
const evalHeaders = ['元SS_ID', '元シート名', '元行番号', 'メールアドレス', '名前', '処理状態', '評価'];
const sendHeaders = ['元SS_ID', '元シート名', '元行番号', 'No', 'メールアドレス', '名前', 'クラス名', 'コースID', 'studentId', '送信状態', '評価', '返信本文'];
const sourceId = 'source_12345678901234567890';
const origin = [sourceId, 'フォームの回答 1', 2];
const evalRow = [...origin, 'student@example.invalid', '生徒', '', 'A'];
const sendRow = [...origin, 1, 'student@example.invalid', '生徒', '授業', 'course-1', 'student-1', '', 'A', '評価：A'];

// GAS I/O only is replaced; the application entrypoints and helpers run unchanged.
class Sheet {
  constructor(name, rows = []) { this.name = name; this.rows = rows.map(r => [...r]); }
  getName() { return this.name; }
  getLastRow() { return this.rows.length; }
  getLastColumn() { return Math.max(0, ...this.rows.map(r => r.length)); }
  getDataRange() { return this.getRange(1, 1, this.getLastRow(), this.getLastColumn()); }
  clear() { this.rows = []; }
  setFrozenRows() {}
  getRange(row, col, height = 1, width = 1) {
    const read = () => Array.from({ length: height }, (_, y) => Array.from({ length: width }, (_, x) => this.rows[row - 1 + y]?.[col - 1 + x] ?? ''));
    return {
      getValues: read,
      getDisplayValues: () => read().map(r => r.map(String)),
      getValue: () => read()[0][0],
      clearContent: () => {
        for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) {
          if (this.rows[row - 1 + y]) this.rows[row - 1 + y][col - 1 + x] = '';
        }
      },
      setValue: value => {
        this.beforeWrite?.({ row, col, value });
        this.rows[row - 1] ??= [];
        this.rows[row - 1][col - 1] = value;
      },
      setValues: values => {
        this.beforeBatch?.(values);
        values.forEach((r, y) => r.forEach((value, x) => {
          this.rows[row - 1 + y] ??= [];
          this.rows[row - 1 + y][col - 1 + x] = value;
        }));
        this.afterBatch?.(values);
      }
    };
  }
  appendRow(row) { this.rows.push([...row]); }
}

function environment() {
  const config = {
    reminderTo: [], formSheetNamePrefix: ['フォームの回答'], emailHeader: 'メールアドレス',
    studentNameHeader: '名前', formStatusHeader: '送信状態', scoreSourceHeader: '評価',
    replyBodyHeader: '返信本文', messageTemplate: '評価：＜score＞',
    formSources: [{ id: sourceId, url: `https://docs.google.com/spreadsheets/d/${sourceId}/edit` }],
    fields: [{ key: 'score', sourceHeader: '評価', evalHeader: '評価', sendHeader: '評価', type: 'score_grade' }]
  };
  const sheets = new Map([
    ['評価データ', new Sheet('評価データ', [evalHeaders])],
    ['送信シート', new Sheet('送信シート', [sendHeaders])],
    ['生徒一覧', new Sheet('生徒一覧', [['No', 'メールアドレス', '名前', 'クラス名', 'コースID', 'studentId'], [1, 'student@example.invalid', '生徒', '授業', 'course-1', 'student-1']])]
  ]);
  const source = new Sheet('フォームの回答 1', [['メールアドレス', '名前', '送信状態', '評価'], ['student@example.invalid', '生徒', '', 5]]);
  const posts = [], messages = [];
  let locked = false, flushes = 0;
  const book = {
    getSheetByName: name => sheets.get(name),
    insertSheet: name => { const s = new Sheet(name); sheets.set(name, s); return s; },
    deleteSheet: sheet => sheets.delete(sheet.name)
  };
  const c = vm.createContext({
    console, Logger: { log() {} },
    PropertiesService: { getScriptProperties: () => ({ getProperty: key => key === 'APP_CONFIG' ? JSON.stringify(config) : null }) },
    LockService: { getScriptLock: () => ({ tryLock: () => { if (locked) return false; locked = true; return true; }, releaseLock: () => { locked = false; } }) },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => book,
      openById: id => { assert.equal(id, sourceId); return { getSheets: () => [source], getSheetByName: name => name === source.name ? source : null }; },
      getUi: () => ({ alert: message => { assert.equal(locked, false, 'interactive alert must run outside the lock'); messages.push(message); } }),
      flush: () => { flushes++; }
    },
    Utilities: { sleep() {} },
    Classroom: { Courses: { Announcements: { create: (data, course) => { posts.push({ data, course }); return { id: `post-${posts.length}` }; } } } }
  });
  vm.runInContext(code, c);
  return { c, config, sheets, source, posts, messages, isLocked: () => locked, flushes: () => flushes };
}

test('normal import, prepare and send keep the existing row layout and recipients', () => {
  const e = environment();
  e.c.importFromFormsToEval(); e.c.evalToSendSheet(); e.c.sendMessages();
  assert.equal(e.posts.length, 1);
  assert.equal(e.posts[0].course, 'course-1');
  assert.deepEqual(Array.from(e.posts[0].data.individualStudentsOptions.studentIds), ['student-1']);
  assert.equal(e.posts[0].data.text, '生徒さんへ(個別メッセージ)\n\n評価：A');
  assert.equal(e.sheets.get('送信シート').rows[1][9], '済');
  assert.equal(e.source.rows[1][2], '済');
  assert.deepEqual(e.sheets.get('送信シート').rows[0], sendHeaders);
  assert.equal(e.messages.length, 1);
});

test('retry import after a source status failure does not append the response twice', () => {
  const e = environment();
  e.source.beforeWrite = () => { throw Error('source permission denied'); };
  e.c.importFromFormsToEval(); e.c.importFromFormsToEval();
  assert.equal(e.sheets.get('評価データ').rows.length, 2);
  assert.equal(e.source.rows[1][2], '');
});

test('retry import after an append response is lost reads the committed row before appending', () => {
  const e = environment(), target = e.sheets.get('評価データ');
  target.afterBatch = () => { throw Error('append response lost'); };
  assert.throws(() => e.c.importFromFormsToEval(), /append response lost/);
  target.afterBatch = null;
  e.c.importFromFormsToEval();
  assert.equal(target.rows.length, 2);
  assert.equal(e.isLocked(), false);
});

test('retry prepare after evaluation status failure does not duplicate the send row', () => {
  const e = environment(), evaluation = e.sheets.get('評価データ');
  evaluation.rows.push([...evalRow]);
  evaluation.beforeWrite = () => { throw Error('evaluation write failed'); };
  assert.throws(() => e.c.evalToSendSheet(), /evaluation write failed/);
  evaluation.beforeWrite = null;
  e.c.evalToSendSheet();
  assert.equal(e.sheets.get('送信シート').rows.length, 2);
  assert.equal(evaluation.rows[1][5], '準備〇');
});

test('duplicate evaluation rows in one preparation produce only one send row', () => {
  const e = environment(), evaluation = e.sheets.get('評価データ');
  evaluation.rows.push([...evalRow], [...evalRow]);
  e.c.evalToSendSheet();
  assert.equal(e.sheets.get('送信シート').rows.length, 2);
  assert.deepEqual(evaluation.rows.slice(1).map(r => r[5]), ['準備〇', '準備〇']);
});

test('posted message is not posted again when recording success fails', () => {
  const e = environment(), send = e.sheets.get('送信シート');
  send.rows.push([...sendRow]);
  send.beforeWrite = ({ value }) => { if (value === '済') throw Error('Internal spreadsheet write failure'); };
  e.c.sendMessages();
  send.beforeWrite = null;
  e.c.sendMessages();
  assert.equal(e.posts.length, 1);
  assert.equal(send.rows[1][9], '送信確認待ち');
  assert.ok(e.sheets.get('エラー').rows.some(r => String(r[6]).includes('post-1')));
});

test('ambiguous Classroom failure is not automatically retried or resubmitted', () => {
  const e = environment(); e.sheets.get('送信シート').rows.push([...sendRow]);
  let requests = 0;
  e.c.Classroom.Courses.Announcements.create = () => { requests++; throw Error('Internal error: response lost'); };
  e.c.sendMessages(); e.c.sendMessages();
  assert.equal(requests, 1);
  assert.equal(e.sheets.get('送信シート').rows[1][9], '送信確認待ち');
});

test('failed send checkpoint cannot call Classroom', () => {
  const e = environment(), send = e.sheets.get('送信シート'); send.rows.push([...sendRow]);
  send.beforeWrite = () => { throw Error('checkpoint failed'); };
  assert.throws(() => e.c.sendMessages(), /checkpoint failed/);
  assert.equal(e.posts.length, 0);
  assert.equal(e.isLocked(), false);
});

test('failed checkpoint flush cannot call Classroom and releases the lock', () => {
  const e = environment(); e.sheets.get('送信シート').rows.push([...sendRow]);
  e.c.SpreadsheetApp.flush = () => { throw Error('flush failed'); };
  assert.throws(() => e.c.sendMessages(), /flush failed/);
  assert.equal(e.posts.length, 0);
  assert.equal(e.isLocked(), false);
});

test('another entrypoint cannot mutate sheets during a send', () => {
  const e = environment(); e.sheets.get('送信シート').rows.push([...sendRow]);
  e.c.Classroom.Courses.Announcements.create = () => {
    assert.throws(() => e.c.importFromFormsToEval(), /実行中/);
    return { id: 'post-1' };
  };
  e.c.sendMessages();
  assert.equal(e.sheets.get('評価データ').rows.length, 1);
  assert.equal(e.isLocked(), false);
});

test('a pending duplicate before a later sent row is not posted', () => {
  const e = environment(), send = e.sheets.get('送信シート');
  const done = [...sendRow]; done[9] = '済'; send.rows.push([...sendRow], done);
  e.c.sendMessages();
  assert.equal(e.posts.length, 0);
  assert.equal(send.rows[1][9], '送信確認待ち');
});

test('historical duplicate send rows require confirmation instead of guessing which to post', () => {
  const e = environment(), send = e.sheets.get('送信シート'); send.rows.push([...sendRow], [...sendRow]);
  e.c.sendMessages();
  assert.equal(e.posts.length, 0);
  assert.deepEqual(send.rows.slice(1).map(r => r[9]), ['送信確認待ち', '送信確認待ち']);
});

test('malformed source rows are rejected rather than posted or merged as one identity', () => {
  const e = environment(), send = e.sheets.get('送信シート');
  for (const rowNo of [0, -1, 1.5, 'not-a-row']) { const row = [...sendRow]; row[2] = rowNo; send.rows.push(row); }
  e.c.sendMessages();
  assert.equal(e.posts.length, 0);
  assert.deepEqual(send.rows.slice(1).map(r => r[9]), ['エラー', 'エラー', 'エラー', 'エラー']);
});

for (const entrypoint of ['initializeSheets', 'saveSettingsFromDialog', 'classroomdata',
  'studentdataMulti', 'createMappingSheet', 'importFromFormsToEval', 'evalToSendSheet',
  'sendMessages', 'remindUngradedAndErrors', 'getSettingsDialogData', 'fetchHeadersFromSources']) {
  test(`${entrypoint} cannot enter during another application operation`, () => {
    const e = environment();
    e.c.withAppLock_(() => assert.throws(() => e.c[entrypoint](), /実行中/));
    assert.equal(e.isLocked(), false);
  });
}

test('import commits evaluation rows before advancing source status', () => {
  const e = environment();
  e.c.SpreadsheetApp.flush = () => { throw Error('flush failed'); };
  assert.throws(() => e.c.importFromFormsToEval(), /flush failed/);
  assert.equal(e.source.rows[1][2], '');
});

test('prepare commits send rows before advancing evaluation or source status', () => {
  const e = environment(), evaluation = e.sheets.get('評価データ');
  evaluation.rows.push([...evalRow]);
  e.c.SpreadsheetApp.flush = () => { throw Error('flush failed'); };
  assert.throws(() => e.c.evalToSendSheet(), /flush failed/);
  assert.equal(evaluation.rows[1][5], '');
  assert.equal(e.source.rows[1][2], '');
});

test('an existing review row remains visible in the completion summary without posting', () => {
  const e = environment(), row = [...sendRow]; row[9] = '送信確認待ち';
  e.sheets.get('送信シート').rows.push(row);
  e.c.sendMessages();
  assert.equal(e.posts.length, 0);
  assert.match(e.messages[0], /要確認.*1件/);
});

test('a number in an ambiguous API message is not mistaken for an HTTP rejection', () => {
  const e = environment(); e.sheets.get('送信シート').rows.push([...sendRow]);
  let requests = 0;
  e.c.Classroom.Courses.Announcements.create = () => { requests++; throw Error('Internal failure for request 429'); };
  e.c.sendMessages();
  assert.equal(requests, 1);
  assert.equal(e.sheets.get('送信シート').rows[1][9], '送信確認待ち');
});

test('an explicit HTTP rate rejection keeps the bounded retry and later resend behavior', () => {
  const e = environment(); e.sheets.get('送信シート').rows.push([...sendRow]);
  let requests = 0;
  e.c.Classroom.Courses.Announcements.create = () => { requests++; throw Error('HTTP 429: RESOURCE_EXHAUSTED'); };
  e.c.sendMessages();
  assert.equal(requests, 3);
  assert.equal(e.sheets.get('送信シート').rows[1][9], '未');
  e.c.Classroom.Courses.Announcements.create = () => { requests++; return { id: 'eventual-success' }; };
  e.c.sendMessages();
  assert.equal(requests, 4);
  assert.equal(e.sheets.get('送信シート').rows[1][9], '済');
});

test('permission rejection remains an error and is not automatically resent', () => {
  const e = environment(); e.sheets.get('送信シート').rows.push([...sendRow]);
  let requests = 0;
  e.c.Classroom.Courses.Announcements.create = () => { requests++; throw Error('HTTP 403: PERMISSION_DENIED'); };
  e.c.sendMessages(); e.c.sendMessages();
  assert.equal(requests, 1);
  assert.equal(e.sheets.get('送信シート').rows[1][9], 'エラー');
  assert.equal(e.source.rows[1][2], '送信エラー');
});

test('source update and even its logging failure cannot roll back a successful send', () => {
  const e = environment(); e.sheets.get('送信シート').rows.push([...sendRow]);
  e.source.beforeWrite = () => { throw Error('source failed'); };
  e.c.ensureErrorSheet_().appendRow = () => { throw Error('log failed'); };
  assert.throws(() => e.c.sendMessages(), /log failed/);
  e.c.sendMessages();
  assert.equal(e.posts.length, 1);
  assert.equal(e.sheets.get('送信シート').rows[1][9], '済');
});

test('prepare deduplication preserves a later sent source status', () => {
  const e = environment(), done = [...sendRow]; done[9] = '済';
  e.sheets.get('送信シート').rows.push(done);
  e.sheets.get('評価データ').rows.push([...evalRow]);
  e.source.rows[1][2] = '済';
  e.c.evalToSendSheet();
  assert.equal(e.sheets.get('送信シート').rows.length, 2);
  assert.equal(e.source.rows[1][2], '済');
});

test('different response rows from the same student are still independent sends', () => {
  const e = environment(), second = [...sendRow]; second[2] = 3;
  e.source.rows.push([...e.source.rows[1]]);
  e.sheets.get('送信シート').rows.push([...sendRow], second);
  e.c.sendMessages();
  assert.equal(e.posts.length, 2);
  assert.deepEqual(e.sheets.get('送信シート').rows.slice(1).map(r => r[9]), ['済', '済']);
});

for (const [entrypoint, sheetName, row] of [
  ['clearEvaluationData', '評価データ', evalRow],
  ['clearSendSheet', '送信シート', sendRow],
  ['clearErrorLog', 'エラー', ['old log']]
]) {
  test(`${entrypoint} confirms before locking and blocks a competing writer`, () => {
    const e = environment();
    e.c.ensureErrorSheet_();
    e.sheets.get(sheetName).rows.push([...row]);
    const before = JSON.stringify(e.sheets.get(sheetName).rows);
    let prompts = 0;
    e.c.SpreadsheetApp.getUi = () => ({ ButtonSet: { OK_CANCEL: 'buttons' }, Button: { OK: 'ok' }, alert: (...args) => {
      assert.equal(e.isLocked(), false);
      if (args.length === 3) prompts++;
      return 'ok';
    } });
    const locks = e.c.LockService;
    e.c.LockService = { getScriptLock: () => ({ tryLock: () => false }) };
    assert.throws(() => e.c[entrypoint](), /実行中/);
    assert.equal(JSON.stringify(e.sheets.get(sheetName).rows), before);
    e.c.LockService = locks;
    e.c[entrypoint]();
    assert.equal(prompts, 2);
    assert.equal(e.isLocked(), false);
    if (sheetName === 'エラー') assert.equal(e.sheets.has(sheetName), false);
    else assert.ok(e.sheets.get(sheetName).rows[1].every(v => v === ''));
  });
}

test('cleared reminder recipients cannot send to the cached recipients', () => {
  const e = environment(); e.config.reminderTo = ['old@example.invalid'];
  e.sheets.set('設定シート', new Sheet('設定シート', [['KEY', 'VALUE', 'NOTE'], ['REMINDER_TO', '', '']]));
  let mails = 0;
  e.c.GmailApp = { sendEmail: () => { mails++; } };
  assert.throws(() => e.c.remindUngradedAndErrors(), /送信先が設定されていません/);
  assert.equal(mails, 0);
  assert.equal(e.isLocked(), false);
});

test('failed success flush and failed error log never permit another post', () => {
  const e = environment(); e.sheets.get('送信シート').rows.push([...sendRow]);
  let flushes = 0;
  e.c.SpreadsheetApp.flush = () => { if (++flushes === 2) throw Error('success flush failed'); };
  e.c.ensureErrorSheet_().appendRow = () => { throw Error('log failed'); };
  assert.throws(() => e.c.sendMessages(), /log failed/);
  e.c.sendMessages();
  assert.equal(e.posts.length, 1);
  assert.equal(e.isLocked(), false);
});

test('empty intermediate rows do not block preparation or become send errors', () => {
  const e = environment();
  e.sheets.get('評価データ').rows.push(new Array(evalHeaders.length).fill(''), [...evalRow]);
  e.c.evalToSendSheet();
  const send = e.sheets.get('送信シート');
  send.rows.splice(1, 0, new Array(sendHeaders.length).fill(''));
  e.c.sendMessages();
  assert.equal(e.posts.length, 1);
  assert.ok(send.rows[1].every(v => v === ''));
  assert.equal(send.rows[2][9], '済');
});
