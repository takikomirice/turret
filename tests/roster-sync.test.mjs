import {test} from 'node:test';
import assert from 'node:assert/strict';
import {harness,sheet,plain} from './backend-harness.mjs';
const classes=[['クラス名','コースID','同期対象(1)'],['Old A','a',1],['Old B','b','']];
const students=[['No','メールアドレス','名前','クラス名','コースID','studentId'],[1,'old@example.com','Old','Old A','a','old']];
test('courses fetches all pages and retains flags by ID',()=>{const h=harness();h.sheets.set('クラス一覧',sheet(classes));const calls=[];h.context.Classroom.Courses.list=args=>{calls.push(args);return args.pageToken?{courses:[{id:'b',name:'New B'}]}:{courses:[{id:'a',name:'New A'}],nextPageToken:'next'};};h.context.classroomdataUnlocked_();assert.equal(calls.length,2);assert.deepEqual(h.sheets.get('クラス一覧').rows,[classes[0],['New A','a',1],['New B','b','']]);});
test('failed later class page preserves all old rows',()=>{const h=harness();h.sheets.set('クラス一覧',sheet(classes));h.context.Classroom.Courses.list=args=>{if(args.pageToken)throw Error('offline');return {courses:[{id:'a',name:'New A'}],nextPageToken:'next'};};assert.throws(()=>h.context.classroomdataUnlocked_(),/offline/);assert.deepEqual(h.sheets.get('クラス一覧').rows,classes);});
for(const mode of ['none','failure'])test('student sync preserves roster on '+mode,()=>{const h=harness();h.sheets.set('クラス一覧',sheet(mode==='none'?[classes[0],['A','a','']]:[...classes.slice(0,2),['B','b',1]]));h.sheets.set('生徒一覧',sheet(students));h.context.Classroom.Courses.Students.list=id=>{if(id==='b')throw Error('offline');return {students:[]};};if(mode==='failure')assert.throws(()=>h.context.studentdataMultiUnlocked_(),/offline/);else h.context.studentdataMultiUnlocked_();assert.deepEqual(h.sheets.get('生徒一覧').rows,students);});
test('successful empty roster clears stale rows and records selected IDs',()=>{const h=harness();h.sheets.set('クラス一覧',sheet(classes));h.sheets.set('生徒一覧',sheet(students));h.context.studentdataMultiUnlocked_();assert.equal(h.sheets.get('生徒一覧').getLastRow(),1);assert.equal(h.properties.get('TURRET_ROSTER_SELECTION'),'["a"]');});
test('initialize refuses malformed populated student sheet without erasing it',()=>{const h=harness(),rows=[['Wrong'],['keep']];h.sheets.set('生徒一覧',sheet(rows));assert.throws(()=>h.context.initializeSheetsUnlocked_(),/生徒一覧|構成/);assert.deepEqual(h.sheets.get('生徒一覧').rows,rows);});

test('all student pages from all selected courses are fetched before replacement', () => {
  const h = harness(), calls = [];
  h.sheets.set('クラス一覧', sheet([classes[0], ['A', 'a', 1], ['B', 'b', 1]]));
  h.sheets.set('生徒一覧', sheet(students));
  h.context.Classroom.Courses.Students.list = (id, options) => {
    assert.deepEqual(h.sheets.get('生徒一覧').rows, students);
    calls.push([id, options.pageToken]);
    const n = options.pageToken ? '2' : '1';
    return { students: [{ profile: { id: id + n, emailAddress: id + n + '@example.com', name: { fullName: id + n } } }], nextPageToken: options.pageToken ? '' : 'next' };
  };
  h.context.studentdataMultiUnlocked_();
  assert.deepEqual(calls, [['a', ''], ['a', 'next'], ['b', ''], ['b', 'next']]);
  assert.equal(h.sheets.get('生徒一覧').getLastRow(), 5);
  assert.equal(h.properties.get('TURRET_ROSTER_SELECTION'), '["a","b"]');
});

test('failed final roster write restores old rows', () => {
  const h = harness();
  h.sheets.set('クラス一覧', sheet(classes));
  h.sheets.set('生徒一覧', sheet(students));
  let once = true;
  h.context.SpreadsheetApp.flush = () => { if (once) { once = false; throw Error('flush denied'); } };
  assert.throws(() => h.context.studentdataMultiUnlocked_(), /flush denied/);
  assert.deepEqual(h.sheets.get('生徒一覧').rows, students);
  assert.equal(h.properties.has('TURRET_ROSTER_SELECTION'), false);
});
