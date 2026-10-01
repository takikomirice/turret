import { test } from 'node:test';
import assert from 'node:assert/strict';
import { harness, sheet, plain, completeConfig } from './backend-harness.mjs';

function setup() {
  const e = harness(), c = e.context, config = plain(completeConfig(c));
  config.fields[0].type = 'score_grade';
  config.gradeScale = [{ from: '5', to: '優' }, { from: '4', to: '良' }];
  config.messageTemplate = '＜student_name＞：＜grade＞';
  e.properties.set('APP_CONFIG', JSON.stringify(config));
  e.sheets.set('生徒一覧', sheet([['No','メールアドレス','名前','クラス名','コースID','studentId'],[1,'student@example.test','生徒','テスト','course','student']]));
  const source = sheet([['Email','Name','Status','Grade'],['student@example.test','生徒','',5]]);
  source.getName = () => 'Responses 1';
  c.SpreadsheetApp.openById = () => ({ getSheets: () => [source], getSheetByName: () => source });
  const legacy = (state = '', grade = '保存済みの優') => {
    const s = sheet([plain(c.getConfiguredEvalHeaders_(config)),[config.formSources[0].id,'Responses 1',2,'student@example.test','生徒',state,grade]]);
    e.sheets.set('評価データ', s); return s;
  };
  const send = () => e.sheets.get('送信シート');
  const value = (header, row = 1) => send().rows[row][send().rows[0].indexOf(header)];
  return { ...e, c, config, source, legacy, send, value };
}

test('direct preparation converts answers into a fixed send body without creating evaluation data', () => {
  const e = setup(); e.c.prepareSendData();
  assert.equal(e.sheets.has('評価データ'), false);
  assert.equal(e.value('Reply'), '生徒：優');
  assert.equal(e.value('studentId'), 'student');
  assert.equal(e.source.rows[1][2], '準備〇');
  e.source.rows[1][3] = 4;
  e.config.messageTemplate = '変更：＜grade＞';
  e.properties.set('APP_CONFIG', JSON.stringify(e.config));
  e.c.prepareSendData();
  assert.equal(e.send().rows.length, 2);
  assert.equal(e.value('Reply'), '生徒：優');
});

for (const sourceState of ['', '反映〇']) test('legacy pending values take priority over source: ' + sourceState, () => {
  const e = setup(), old = e.legacy(); e.source.rows[1][2] = sourceState;
  e.c.prepareSendData(); e.c.prepareSendData();
  assert.equal(e.send().rows.length, 2);
  assert.equal(e.value('Reply'), '生徒：保存済みの優');
  assert.equal(old.rows[1][5], '準備〇');
  assert.equal(old.rows[1][6], '保存済みの優');
});

for (const state of ['準備〇','準備×']) test('legacy ' + state + ' is not regenerated from a blank source status', () => {
  const e = setup(); e.legacy(state); e.c.prepareSendData();
  assert.equal(e.send().rows.length, 1);
});

test('failed source-state write and retry keep exactly one send row', () => {
  const e = setup(), range = e.source.getRange.bind(e.source);
  e.source.getRange = (...args) => { const r = range(...args); if (args[0] > 1) r.setValue = () => { throw Error('source denied'); }; return r; };
  e.c.prepareSendData(); e.c.prepareSendData();
  assert.equal(e.send().rows.length, 2);
  assert.equal(e.source.rows[1][2], '');
});

test('lost append response is recovered using the committed send key', () => {
  const e = setup(); const send = e.c.ensureSendSheet_(e.config), range = send.getRange.bind(send);
  send.getRange = (...args) => { const r = range(...args), write = r.setValues; if (args[0] > 1) r.setValues = values => { write(values); throw Error('lost append response'); }; return r; };
  assert.throws(() => e.c.prepareSendData(), /lost append response/);
  send.getRange = range; e.c.prepareSendData();
  assert.equal(send.rows.length, 2);
});

test('failed batch flush never advances source or legacy state', () => {
  const e = setup(), old = e.legacy();
  e.c.SpreadsheetApp.flush = () => { throw Error('flush failed'); };
  assert.throws(() => e.c.prepareSendData(), /flush failed/);
  assert.equal(e.source.rows[1][2], ''); assert.equal(old.rows[1][5], '');
});

for (const status of ['済','送信確認待ち','エラー']) test('existing ' + status + ' and its body survive preparation', () => {
  const e = setup(); e.c.prepareSendData();
  e.send().rows[1][e.send().rows[0].indexOf('送信状態')] = status;
  e.legacy(); e.source.rows[1][2] = '';
  e.c.prepareSendData();
  assert.equal(e.send().rows.length, 2); assert.equal(e.value('送信状態'), status);
  assert.equal(e.value('Reply'), '生徒：優'); assert.equal(e.source.rows[1][2], '');
});

for (const migrated of [false,true]) test('return guard blocks ' + (migrated ? 'migration' : 'direct preparation'), () => {
  const e = setup(); if (migrated) e.legacy();
  e.c.scoringBuildReturnGuard_ = () => ({ bySource: new Map() });
  e.c.prepareSendData();
  assert.equal(e.send().rows.length, 1);
});

test('compatibility entrypoints also create send rows without evaluation data', () => {
  for (const name of ['importFromFormsToEval','evalToSendSheet']) {
    const e = setup(); e.c[name](); assert.equal(e.sheets.has('評価データ'), false); assert.equal(e.send().rows.length, 2);
  }
});

test('completed legacy history does not block a new field layout', () => {
  const e = setup(), old = e.legacy('準備〇'), before = plain(old.rows);
  const config = plain(e.config); config.fields.push({key:'comment',sourceHeader:'Comment',type:'text'});
  assert.doesNotThrow(() => e.c.validateConfigSheetLayouts_(e.c.normalizeAppConfig_(config)));
  assert.deepEqual(old.rows, before);
});

test('pending legacy values retain their column-layout protection', () => {
  const e = setup(); e.legacy();
  const config = plain(e.config); config.fields[0].evalHeader = 'Changed';
  assert.throws(() => e.c.validateConfigSheetLayouts_(e.c.normalizeAppConfig_(config)), /列|構成/);
});

test('direct fields can reuse a source header with distinct send columns', () => {
  const e = setup();
  e.config.fields.push({key:'raw',sourceHeader:'Grade',sendHeader:'元点数',type:'text'});
  e.config.fields.push({key:'name',sourceHeader:'Name',sendHeader:'回答者名',type:'text',evalHeader:'名前'});
  e.properties.set('APP_CONFIG', JSON.stringify(e.config));
  e.c.prepareSendData();
  assert.equal(e.value('元点数'),'5');assert.equal(e.value('回答者名'),'生徒');
});

test('ambiguous legacy headers cannot select a different recipient', () => {
  const e = setup(), old = e.legacy();
  old.rows[0].push('メールアドレス');old.rows[1].push('other@example.test');
  e.sheets.get('生徒一覧').rows.push([2,'other@example.test','別の生徒','テスト','course','other']);
  assert.throws(() => e.c.prepareSendData(), /旧評価データ.*重複/);
  assert.equal(e.send().rows.length,1);assert.equal(old.rows[1][5],'');
});

test('pending migration rejects ambiguous field mappings even when send headers differ', () => {
  const e = setup();e.legacy();
  e.config.fields.push({key:'raw',sourceHeader:'Grade',sendHeader:'元点数',type:'text'});
  assert.throws(() => e.c.validateConfigSheetLayouts_(e.config), /旧評価データ.*重複/);
  e.properties.set('APP_CONFIG',JSON.stringify(e.config));
  assert.throws(() => e.c.prepareSendData(), /旧評価データ.*重複/);
});

test('extra duplicate recipient column on a populated send sheet is rejected before preparing or sending',()=>{
  const e=setup();e.c.prepareSendData();const send=e.send();
  send.rows[0].push('studentId');send.rows[1].push('wrong-student');
  const before=plain(send.rows);
  assert.throws(()=>e.c.prepareSendData(),/送信シート.*構成/);
  e.c.Classroom={Courses:{Announcements:{create(){assert.fail('must reject before posting');}}}};
  assert.throws(()=>e.c.sendMessages(),/送信シート.*構成/);
  assert.deepEqual(send.rows,before);
});

test('extra headers on an empty send sheet are removed when preparing its schema',()=>{
  const e=setup(),send=e.c.ensureSendSheet_(e.config);send.rows[0].push('余分な列');
  e.c.ensureSendSheet_(e.config);
  assert.deepEqual(send.rows[0],plain(e.c.getConfiguredSendHeaders_(e.config)));
});
