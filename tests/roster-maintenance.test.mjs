import {test} from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import vm from 'node:vm';
import {harness,sheet,completeConfig,plain} from './backend-harness.mjs';
import {adminEnvironment,AdminSheet} from './helpers/admin-environment.mjs';
import {scoringEnvironment,sourceId} from './helpers/scoring-environment.mjs';

const headers=['No','メールアドレス','出席番号（任意）','名前','クラス名','コースID','studentId','除外(1)'];
const classHeaders=['クラス名','コースID','同期対象(1)'];
const student=(email,name,course='a',number='',excluded='')=>[99,email,number,name,'クラス2',course,'id-'+email,excluded];
function syncSetup(rows=[]){const e=harness();e.sheets.set('クラス一覧',sheet([classHeaders,['クラス2','a',1]]));e.sheets.set('生徒一覧',sheet([headers,...rows]));e.context.Classroom.Courses.Students.list=()=>({students:[]});return e;}
const fetched=(email,name)=>({profile:{id:'id-'+email,emailAddress:email,name:{fullName:name}}});

test('class sync uses natural class-name order and preserves selection by course ID',()=>{
 const e=syncSetup();e.context.Classroom.Courses.list=()=>({courses:[{id:'b',name:'クラス10'},{id:'a',name:'クラス2'}]});e.context.classroomdataUnlocked_();
 assert.deepEqual(e.sheets.get('クラス一覧').rows.slice(1),[['クラス2','a',1],['クラス10','b','']]);
});
for(const courses of [[{name:'No ID'}],[{id:'a',name:''}],[{id:'a',name:'A'},{id:'a',name:'B'}]])test('malformed or duplicate course results cannot replace selected classes: '+JSON.stringify(courses),()=>{
 const e=syncSetup(),before=plain(e.sheets.get('クラス一覧').rows);e.context.Classroom.Courses.list=()=>({courses});assert.throws(()=>e.context.classroomdataUnlocked_(),/空欄|重複/);assert.deepEqual(e.sheets.get('クラス一覧').rows,before);
});
test('student sync sorts attendance numerically within class and keeps excluded students last',()=>{
 const e=syncSetup([student('ten@x','Ten','a','10'),student('two@x','Two','a','02'),student('off@x','Off','a','01','1')]);
 e.context.Classroom.Courses.Students.list=()=>({students:[fetched('ten@x','Ten'),fetched('blank@x','Blank'),fetched('off@x','Off'),fetched('two@x','Two')]});e.context.studentdataMultiUnlocked_();
 assert.deepEqual(e.sheets.get('生徒一覧').rows.slice(1).map(r=>[r[0],r[1],r[2],r[7]]),[[1,'two@x','02',''],[2,'ten@x','10',''],[3,'blank@x','',''],[4,'off@x','01','1']]);
});
test('excluded record survives an empty Classroom result so rejoining cannot silently re-enable it',()=>{
 const e=syncSetup([student('off@x','Off','a','01','1')]);e.context.studentdataMultiUnlocked_();assert.equal(e.sheets.get('生徒一覧').rows[1][7],'1');
 e.context.Classroom.Courses.Students.list=()=>({students:[fetched('off@x','New name')]});e.context.studentdataMultiUnlocked_();assert.equal(e.sheets.get('生徒一覧').rows[1][7],'1');
});
test('deselecting and reselecting a course preserves its exclusion records',()=>{
 const e=syncSetup([student('off@x','Off','a','','1')]);e.sheets.get('クラス一覧').rows=[classHeaders,['B','b',1]];
 e.context.studentdataMultiUnlocked_();assert.equal(e.sheets.get('生徒一覧').rows[1][7],'1');
 e.sheets.get('クラス一覧').rows=[classHeaders,['A','a',1]];e.context.Classroom.Courses.Students.list=()=>({students:[fetched('off@x','Off')]});e.context.studentdataMultiUnlocked_();assert.equal(e.sheets.get('生徒一覧').rows[1][7],'1');
});
test('duplicate course and case-insensitive email stops sync before any API or data replacement',()=>{
 const rows=[student('same@x','A'),student('SAME@x','B')],e=syncSetup(rows);let calls=0;e.context.Classroom.Courses.Students.list=()=>{calls++;return {};};
 assert.throws(()=>e.context.studentdataMultiUnlocked_(),/重複/);assert.equal(calls,0);assert.deepEqual(e.sheets.get('生徒一覧').rows,[headers,...rows]);
});
function delivery(){const e=harness(),c=e.context,config=completeConfig(c);e.properties.set('APP_CONFIG',JSON.stringify(config));
 const source=sheet([['Email','Name','Status','Grade'],['s@x','S','','A']]);source.getName=()=> 'Responses 1';
 c.SpreadsheetApp.openById=()=>({getSheets:()=>[source],getSheetByName:()=>source});c.Utilities={sleep(){}};
 const posts=[];c.Classroom.Courses.Announcements={create:(payload,course)=>{posts.push({payload,course});return {id:'post'};}};
 e.sheets.set('生徒一覧',sheet([headers,student('s@x','S')]));
 const sendHeaders=plain(c.getConfiguredSendHeaders_(config)),row=Array(sendHeaders.length).fill(''),values={'元SS_ID':'abcdefghijklmnopqrstuv','元シート名':'Responses 1','元行番号':2,No:1,'メールアドレス':'s@x','名前':'S','クラス名':'クラス2','コースID':'a',studentId:'id-s@x',Reply:'A'};
 sendHeaders.forEach((key,i)=>{row[i]=values[key]??'';});e.sheets.set('送信シート',sheet([sendHeaders,row]));return {...e,c,posts,source};}
for(const mode of ['deleted','excluded','changed ID'])test('queued delivery rechecks current roster: '+mode,()=>{
 const e=delivery(),s=e.sheets.get('生徒一覧');if(mode==='deleted')s.rows.splice(1);if(mode==='excluded')s.rows[1][7]='1';if(mode==='changed ID')s.rows[1][6]='other';
 e.c.sendMessagesUnlocked_();assert.equal(e.posts.length,0);assert.equal(e.source.rows[1][2],'');assert.notEqual(e.sheets.get('送信シート').rows[1].at(e.sheets.get('送信シート').rows[0].indexOf('送信状態')),'済');
});
test('a valid queued recipient still sends after unrelated roster rows are reordered',()=>{
 const e=delivery();e.sheets.get('生徒一覧').rows.splice(1,0,student('other@x','Other'));e.c.sendMessagesUnlocked_();assert.equal(e.posts.length,1);assert.equal(e.posts[0].course,'a');
});
test('duplicate recipient keys stop queued delivery without any post',()=>{
 const e=delivery();e.sheets.get('生徒一覧').rows.push(student('S@x','Other'));assert.throws(()=>e.c.sendMessagesUnlocked_(),/重複/);assert.equal(e.posts.length,0);
});
for(const mappedCourse of ['', 'a'])test('excluding one course never redirects its answer into another active course: '+(mappedCourse||'unmapped'),()=>{
 const e=delivery(),roster=e.sheets.get('生徒一覧');roster.rows[1][7]='1';roster.rows.push(student('s@x','S','b'));e.sheets.get('送信シート').rows.splice(1);
 if(mappedCourse)e.sheets.set('対応表',sheet([['元SS_ID','元スプシ名','元シート名','クラス名','courseId','メモ'],['abcdefghijklmnopqrstuv','回答','Responses 1','クラス2',mappedCourse,'']]));
 e.c.prepareSendDataUnlocked_();assert.equal(e.sheets.get('送信シート').rows.length,1);assert.equal(e.source.rows[1][2],'準備×');
});
test('exclusion edited during the first post blocks a later queued recipient in the same run',()=>{
 const e=delivery(),s=e.sheets.get('生徒一覧');s.rows.push(student('two@x','Two'));e.source.rows.push(['two@x','Two','','A']);
 const send=e.sheets.get('送信シート'),copy=send.rows[1].slice(),h=send.rows[0];copy[h.indexOf('元行番号')]=3;copy[h.indexOf('メールアドレス')]='two@x';copy[h.indexOf('studentId')]='id-two@x';send.rows.push(copy);
 const post=e.c.Classroom.Courses.Announcements.create;e.c.Classroom.Courses.Announcements.create=(...args)=>{s.rows[2][7]='1';return post(...args);};e.c.sendMessagesUnlocked_();assert.equal(e.posts.length,1);assert.equal(e.source.rows[2][2],'');
});
for(const columns of [6,7])test('a roster trimmed to '+columns+' physical columns migrates without losing cells',()=>{
 const e=adminEnvironment(),old=columns===6?['No','メールアドレス','名前','クラス名','コースID','studentId']:headers.slice(0,7),row=columns===6?[1,'s@x','=A2','A','a','s']:[1,'s@x','01','=A2','A','a','s'],s=new AdminSheet('生徒一覧',[old,row]);s.maxColumns=columns;
 const get=s.getRange.bind(s);s.getRange=(r,c,h=1,w=1)=>{if(c+w-1>s.maxColumns)throw Error('range exceeds grid');return get(r,c,h,w);};e.sheets.set('生徒一覧',s);e.c.ensureStudentSheetForSync_();assert.equal(s.maxColumns,8);assert.deepEqual(s.rows[0],headers);assert.equal(s.rows[1][3],'=A2');
});
test('failed exclusion-column cleanup still removes a newly inserted attendance column and reports incomplete recovery',()=>{
 const e=adminEnvironment(),s=new AdminSheet('生徒一覧',[['No','メールアドレス','名前','クラス名','コースID','studentId'],[1,'s@x','S','A','a','s']]);e.sheets.set('生徒一覧',s);let flushes=0;e.c.SpreadsheetApp.flush=()=>{if(++flushes===2)throw Error('migration flush failed');};
 const get=s.getRange.bind(s);s.getRange=(...args)=>{const r=get(...args);if(args[1]===8)r.clearContent=()=>{throw Error('clear failed');};return r;};assert.throws(()=>e.c.ensureStudentSheetForSync_(),/復旧が未完了/);assert.equal(s.rows[1][2],'S');assert.equal(s.rows[0][2],'名前');
});
test('excluded student is absent from name completion',()=>{
 const e=adminEnvironment();e.sheets.set('生徒一覧',new AdminSheet('生徒一覧',[headers,student('on@x','On'),student('off@x','Off','a','','1')]));assert.deepEqual(plain(e.c.managedRoster_('a')),[{email:'on@x',name:'On'}]);
});
test('excluded historical answers are omitted without blocking grade output',()=>{
 const {c}=adminEnvironment(),profile=c.gradeDefault_();profile.base.valueHeader='点';profile.base.leading=['roster:email'];
 const data={classes:[{id:'a',name:'A',students:[{email:'on@x',name:'On'}]}],excludedStudents:[{courseId:'a',email:'off@x'}],responses:[{id:'r',courseId:'a',email:'off@x',date:'2026-09-30',fingerprint:'f',fields:{点:'5'}}],issues:[]};
 const out=plain(c.gradeBuild_(data,profile));assert.equal(out.issues.length,0);assert.deepEqual(out.classes[0].rows,[['メールアドレス'],['on@x']]);
});
test('retired excluded-only class with stale ICS mapping cannot block another active class CSV',()=>{
 const {c}=adminEnvironment(),p=c.gradeDefault_();p.base.mapping={old:'Retired calendar'};
 const data={classes:[{id:'old',name:'Old',students:[]},{id:'active',name:'Active',students:[{email:'on@x',name:'On'}]}],excludedStudents:[{courseId:'old',email:'off@x'}],responses:[],issues:[]};
 const out=plain(c.gradeBuild_(data,p));assert.equal(out.issues.length,0);assert.deepEqual(out.classes.map(x=>x.id),['active']);
 data.responses.push({id:'r',courseId:'old',email:'unknown@x',date:'2026-09-30',fingerprint:'f',fields:{}});assert.ok(c.gradeBuild_(data,p).issues.some(x=>x.type==='roster'));
});
test('excluded-only historical tab needs no obsolete grade column, but unknown answers still require validation',()=>{
 const e=scoringEnvironment(),p=e.c.gradeDefault_();p.base.valueHeader='点数';e.configure({formSources:[{id:sourceId,courseId:'old'}]});
 e.sheets.set('生徒一覧',new AdminSheet('生徒一覧',[headers,student('student@example.com','Off','old','','1'),student('on@x','On','active')]));
 e.answer.rows=[['メール','名前','状態','古い評価'],['student@example.com','Off','','5']];
 assert.equal(e.c.gradeReadData_(p.base,true).issues.length,0);e.answer.rows.push(['unknown@x','Unknown','','5']);assert.ok(e.c.gradeReadData_(p.base,true).issues.some(x=>x.type==='columns'));
});
test('manual sorting uses whole rows, preserves No and cell formulas, and removes its temporary column',()=>{
 const e=adminEnvironment(),s=new AdminSheet('生徒一覧',[headers,student('ten@x','=A3','a','10'),student('two@x','=A2','a','2')]);
 const range=s.getRange.bind(s);s.getRange=(...args)=>{const r=range(...args);r.sort=spec=>{const col=spec.column-1,sorted=s.rows.slice(1).sort((a,b)=>a[col]-b[col]);s.rows=[s.rows[0],...sorted];return r;};return r;};
 const remove=s.deleteColumns.bind(s);s.deleteColumns=(start,count)=>{s.rows.forEach(r=>{r.splice(start-1,count);while(r.length&&(r.at(-1)===''||r.at(-1)==null))r.pop();});return remove(start,count);};e.sheets.set('生徒一覧',s);
 e.c.sortRosterSheet_('生徒一覧');assert.deepEqual(s.rows.slice(1).map(r=>[r[0],r[1],r[3]]),[[99,'two@x','=A2'],[99,'ten@x','=A3']]);assert.equal(s.maxColumns,26);assert.equal(s.getLastColumn(),8);
});
test('failed native sort removes its key column without changing any roster row',()=>{
 const e=adminEnvironment(),s=new AdminSheet('生徒一覧',[headers,student('ten@x','Ten','a','10'),student('two@x','Two','a','2')]),before=plain(s.rows),get=s.getRange.bind(s),del=s.deleteColumns.bind(s);e.sheets.set('生徒一覧',s);
 s.getRange=(...args)=>{const r=get(...args);r.sort=()=>{throw Error('sort denied');};return r;};s.deleteColumns=(start,count)=>{s.rows=s.rows.map(r=>r.slice(0,8));return del(start,count);};
 assert.throws(()=>e.c.sortRosterSheet_('生徒一覧'),/sort denied/);assert.deepEqual(s.rows,before);assert.equal(s.maxColumns,26);
});
test('failed key-column cleanup is recoverable on the next sort without deleting user columns',()=>{
 const e=adminEnvironment(),s=new AdminSheet('生徒一覧',[headers,student('ten@x','Ten','a','10'),student('two@x','Two','a','2')]),get=s.getRange.bind(s),del=s.deleteColumns.bind(s);e.sheets.set('生徒一覧',s);let fail=true;
 s.getRange=(...args)=>{const r=get(...args);r.sort=spec=>{s.rows=[s.rows[0],...s.rows.slice(1).sort((a,b)=>a[spec.column-1]-b[spec.column-1])];return r;};return r;};
 s.deleteColumns=(start,count)=>{if(fail)throw Error('cleanup unavailable');s.rows=s.rows.map(r=>r.slice(0,8));return del(start,count);};assert.throws(()=>e.c.sortRosterSheet_('生徒一覧'),/再実行/);assert.equal(s.maxColumns,27);
 fail=false;e.c.sortRosterSheet_('生徒一覧');assert.equal(s.maxColumns,26);assert.deepEqual(s.rows.slice(1).map(r=>r[1]),['two@x','ten@x']);
});
test('class-and-student screen directs sorting to sheets and keeps exclusion instructions',()=>{
 const script=readFileSync('Setting.html','utf8').match(/<script>([\s\S]*?)<\/script>/)[1],c=vm.createContext({console,structuredClone,setTimeout,clearTimeout});vm.runInContext(script.replace(/boot\(\);\s*$/,''),c);
 vm.runInContext("state.data={counts:{},sheets:[],progress:[]};",c);const markup=vm.runInContext("renderStep('classes')",c);assert.doesNotMatch(markup,/data-action="sortClasses"|data-action="sortStudents"|クラスを整列|生徒を整列/);assert.match(markup,/並べ替えは各シート上/);assert.match(markup,/除外\(1\)/);assert.match(markup,/再取得後も除外/);
});

test('roster preparation creates full-grid filters without reordering existing data',()=>{
 const e=adminEnvironment(),classes=new AdminSheet('クラス一覧',[classHeaders,['クラス10','b',1],['クラス2','a','']]),students=new AdminSheet('生徒一覧',[headers,student('ten@x','Ten'),student('two@x','Two')]);
 e.sheets.set('クラス一覧',classes);e.sheets.set('生徒一覧',students);
 for(const [s,width]of [[classes,3],[students,8]]){
  const before=plain(s.rows);e.c.ensureSheet_(s.name,s.rows[0]);if(s===students)e.c.ensureStudentSheetForSync_();
  assert.deepEqual(s.getFilter(),{row:1,column:1,rows:s.getMaxRows(),columns:width});assert.deepEqual(s.rows,before);
 }
});

test('existing roster filters and their criteria remain untouched when preparing or updating',()=>{
 const e=adminEnvironment(),s=new AdminSheet('クラス一覧',[classHeaders,['クラス2','a',1]]),filter={criteria:{3:'1'}};s.filter=filter;e.sheets.set('クラス一覧',s);
 e.c.ensureSheet_('クラス一覧',classHeaders);e.c.replaceRosterRows_('クラス一覧',classHeaders,[['クラス10','b',1]]);
 assert.equal(s.getFilter(),filter);assert.deepEqual(filter.criteria,{3:'1'});
 assert.equal(e.c.ensureSheet_('エラー',['エラー']).getFilter(),null);
});

test('newly fetched roster receives a filter and filter creation failure restores roster cells',()=>{
 const e=adminEnvironment(),s=new AdminSheet('生徒一覧',[headers,student('old@x','Old')]);e.sheets.set('生徒一覧',s);
 const before=plain(s.rows),get=s.getRange.bind(s);s.getRange=(...args)=>{const r=get(...args);r.createFilter=()=>{throw Error('filter denied');};return r;};
 assert.throws(()=>e.c.replaceRosterRows_('生徒一覧',headers,[student('new@x','New')]),/filter denied/);assert.deepEqual(s.rows,before);
 s.getRange=get;e.c.replaceRosterRows_('生徒一覧',headers,[student('new@x','New')]);assert.equal(s.getFilter().columns,8);
});
