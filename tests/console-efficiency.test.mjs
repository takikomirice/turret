import test from 'node:test';
import assert from 'node:assert/strict';
import {scoringEnvironment,AdminSheet,sourceId} from './helpers/scoring-environment.mjs';
import {scoringWorkload} from './helpers/scoring-workload.mjs';
const plain=value=>JSON.parse(JSON.stringify(value));

test('empty send with 5,000 answers does not open or read answer workbooks',()=>{
 const e=scoringWorkload(5000,5);e.resetMetrics();
 assert.match(e.c.sendMessagesUnlocked_(),/データがありません/);
 assert.equal(e.metrics.opens,0);assert.equal(e.metrics.valueReads,0);assert.equal(e.metrics.cells,0);
});

function queued(){
 const e=scoringEnvironment(),sheet=e.c.ensureSendSheet_(e.c.getConfig_());
 e.sheets.set('生徒一覧',new AdminSheet('生徒一覧',[['メールアドレス','名前','コースID','studentId'],['student@example.com','生徒','100','kid']]));
 const headers=sheet.rows[0];
 const add=(row,status='未')=>sheet.rows.push(headers.map(h=>({'元SS_ID':sourceId,'元シート名':'回答 1','元行番号':row,'メールアドレス':'student@example.com','名前':'生徒','コースID':'100','studentId':'kid','返信本文':'本文','送信状態':status}[h]??'')));
 return {...e,sheet,add};
}
test('sent, review and blank rows do not build the answer guard; review count is preserved',()=>{
 const e=queued();e.add(2,'済');e.add(3,'送信確認待ち');e.sheet.rows.push([]);
 e.c.scoringBuildReturnGuard_=()=>{throw Error('unnecessary answer scan');};
 assert.match(e.c.sendMessagesUnlocked_(),/確認待ち.*1|要確認.*1/);
});
test('eligible queued rows build even a null guard only once and retain send checkpoints',()=>{
 const e=queued();e.add(2);e.add(3);let guards=0,posts=0;
 e.answer.rows.push([...e.answer.rows[1]]);
 e.c.scoringBuildReturnGuard_=()=>{guards++;return null;};
 e.c.Classroom={Courses:{Announcements:{create(){
  assert.equal(e.sheet.rows[posts+1][e.sheet.rows[0].indexOf('送信状態')],'送信確認待ち');
  return {id:'post-'+(++posts)};
 }}}};
 e.c.sendMessagesUnlocked_();assert.equal(guards,1);assert.equal(posts,2);
});
test('guard failure stops before any external send or queued status change',()=>{
 const e=queued();e.add(2);e.c.scoringBuildReturnGuard_=()=>{throw Error('answer unavailable');};
 assert.throws(()=>e.c.sendMessagesUnlocked_(),/answer unavailable/);
 assert.equal(e.sheet.rows[1][e.sheet.rows[0].indexOf('送信状態')],'未');
});

function historyEnvironment(count){
 const e=scoringEnvironment();
 const records=[{id:'form:one',kind:'form',stage:'registered'},{id:'template:one',kind:'template'},...Array.from({length:count},(_,i)=>({id:'post:'+i,kind:'announcement',title:'返却'+i,postId:'p'+i,revision:'revision-'+i,stage:'published'}))];
 for(const r of records)Object.assign(r,{schema:1,ownerSpreadsheetId:e.ss.getId(),ownerScriptId:e.c.ScriptApp.getScriptId()});
 e.sheets.set('フォーム管理',new AdminSheet('フォーム管理',[['種別','ID','内容(JSON)'],...records.map(r=>[r.kind,r.id,JSON.stringify(r)])]));
 return e;
}
test('console transfers 50 history records per page, keeps other records and legacy calls, and preserves all server history',()=>{
 const e=historyEnvironment(103),page=n=>e.c.getFormConsoleData({returnPage:n});
 const first=page(0),second=page(1),last=page(99);
 assert.deepEqual(plain(first.returnHistory),{page:0,pageSize:50,total:103});
 assert.equal(first.records.length,52);assert.deepEqual(plain(first.records.slice(2).map(r=>r.id)),Array.from({length:50},(_,i)=>'post:'+(i+53)));
 assert.deepEqual(plain(second.records.slice(2).map(r=>r.id)),Array.from({length:50},(_,i)=>'post:'+(i+3)));
 assert.equal(last.returnHistory.page,2);assert.deepEqual(plain(last.records.slice(2).map(r=>r.id)),['post:0','post:1','post:2']);
 assert.deepEqual(plain(first.setupProgress),plain(last.setupProgress));
 assert.equal(e.c.getFormConsoleData().records.length,105);assert.equal(e.c.getManagedRecords_().length,105);
 assert.equal(e.c.loadManagedRecord_('post:0','revision-0').postId,'p0');
});
test('empty history clamps to page zero; malformed page requests release the lock without mutations',()=>{
 const e=historyEnvironment(0),before=JSON.stringify(e.sheets.get('フォーム管理').rows);
 assert.deepEqual(plain(e.c.getFormConsoleData({returnPage:9}).returnHistory),{page:0,pageSize:50,total:0});
 for(const page of [-1,0.5,'1',Infinity,NaN])assert.throws(()=>e.c.getFormConsoleData({returnPage:page}),/ページ番号/);
 assert.equal(e.isLocked(),false);assert.equal(JSON.stringify(e.sheets.get('フォーム管理').rows),before);
});
