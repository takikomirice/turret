import {test} from 'node:test';
import assert from 'node:assert/strict';
import {harness,sheet as baseSheet,plain} from './backend-harness.mjs';
function sheet(rows) {
 const s=baseSheet(rows);
 s.insertColumnBefore=column=>{s.rows.forEach(row=>row.splice(column-1,0,''));return s;};
 s.deleteColumn=column=>{s.rows.forEach(row=>row.splice(column-1,1));return s;};
 return s;
}
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

const numberedHeaders=['No','メールアドレス','出席番号（任意）','名前','クラス名','コースID','studentId'];
test('legacy roster initialization inserts optional number without rewriting existing cells',()=>{
 const h=harness(),rows=[students[0],[1,'old@example.com','=A2','Old A','a','old']],s=sheet(rows);
 const format=['n','email','name-style','class-style','course-style','id-style'];
 const insert=s.insertColumnBefore;s.insertColumnBefore=c=>{format.splice(c-1,0,'');return insert(c);};
 h.sheets.set('生徒一覧',s);h.context.ensureStudentSheetForSync_();
 assert.deepEqual(s.rows,[numberedHeaders.concat('除外(1)'),[1,'old@example.com','','=A2','Old A','a','old']]);
 assert.equal(format[3],'name-style');
 h.context.ensureStudentSheetForSync_();assert.equal(s.rows[0].length,8);
});
test('legacy migration rolls back column insertion when header write fails',()=>{
 const h=harness(),s=sheet(students),get=s.getRange;
 s.getRange=(...args)=>{const r=get(...args);r.setValue=()=>{throw Error('header denied');};return r;};
 h.sheets.set('生徒一覧',s);assert.throws(()=>h.context.ensureStudentSheetForSync_(),/header denied/);
 assert.deepEqual(s.rows,students);
});
test('numbered roster keeps attendance numbers per course and email across sync',()=>{
 const h=harness();h.sheets.set('クラス一覧',sheet([classes[0],['A','a',1],['B','b',1]]));
 h.sheets.set('生徒一覧',sheet([numberedHeaders,[1,'SAME@example.com','001','Old','A','a','s'],[2,'same@example.com','07','Old','B','b','s']]));
 h.context.Classroom.Courses.Students.list=()=>({students:[{profile:{id:'s',emailAddress:'same@example.com',name:{fullName:'New'}}}]});
 h.context.studentdataMultiUnlocked_();
 assert.deepEqual(h.sheets.get('生徒一覧').rows,[numberedHeaders.concat('除外(1)'),[1,'same@example.com','001','New','A','a','s',''],[2,'same@example.com','07','New','B','b','s','']]);
});
test('conflicting stored attendance numbers stop before fetching or changing roster',()=>{
 const h=harness(),rows=[numberedHeaders,[1,'student@example.com','01','S','A','a','s'],[2,'STUDENT@example.com','02','S','A','a','s']];
 h.sheets.set('クラス一覧',sheet(classes));h.sheets.set('生徒一覧',sheet(rows));
 h.context.Classroom.Courses.Students.list=()=>{throw Error('must not fetch');};
 assert.throws(()=>h.context.studentdataMultiUnlocked_(),/出席番号.*(一致|重複|異な)/);assert.deepEqual(h.sheets.get('生徒一覧').rows,rows);
});
test('legacy roster rejects unexpected extra columns without deleting them',()=>{
 const h=harness(),rows=[students[0].concat('独自列'),students[1].concat('keep')];h.sheets.set('生徒一覧',sheet(rows));
 assert.throws(()=>h.context.ensureStudentSheetForSync_(),/構成/);assert.deepEqual(h.sheets.get('生徒一覧').rows,rows);
});
test('both old and new system sheet names stay excluded from scoring',()=>{
 const h=harness();for(const name of ['フォーム管理','システム管理'])assert.equal(h.context.scoringIsInternalSheet_(name),true);
});
test('failed sync after successful legacy migration restores original rows and column layout',()=>{
 const h=harness();h.sheets.set('クラス一覧',sheet(classes));h.sheets.set('生徒一覧',sheet(students));let flushes=0;
 h.context.SpreadsheetApp.flush=()=>{if(++flushes===2)throw Error('final write denied');};
 assert.throws(()=>h.context.studentdataMultiUnlocked_(),/final write denied/);
 assert.deepEqual(h.sheets.get('生徒一覧').rows,students);assert.equal(h.properties.has('TURRET_ROSTER_SELECTION'),false);
});
test('blank attendance number and newly fetched student remain blank',()=>{
 const h=harness();h.sheets.set('クラス一覧',sheet(classes));h.sheets.set('生徒一覧',sheet([numberedHeaders,[1,'old@example.com','','Old','A','a','old']]));
 h.context.Classroom.Courses.Students.list=()=>({students:['old','new'].map(id=>({profile:{id,emailAddress:id+'@example.com',name:{fullName:id}}}))});
 h.context.studentdataMultiUnlocked_();assert.deepEqual(h.sheets.get('生徒一覧').rows.slice(1).map(row=>row[2]),['','']);
});
test('same student duplicated with blank versus populated number is a conflict',()=>{
 const h=harness();h.sheets.set('クラス一覧',sheet(classes));
 const rows=[numberedHeaders,[1,'s@example.com','','S','A','a','s'],[2,'s@example.com','01','S','A','a','s']];h.sheets.set('生徒一覧',sheet(rows));
 assert.throws(()=>h.context.studentdataMultiUnlocked_(),/出席番号.*一致/);assert.deepEqual(h.sheets.get('生徒一覧').rows,rows);
});
