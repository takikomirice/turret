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

function managedNames(){
 const e=scoringEnvironment();
 const roster=new AdminSheet('生徒一覧',[['メールアドレス','名前','コースID'],['kid@example.com','生徒A','100'],['kid@example.com','生徒B','200']]);
 e.sheets.set(roster.name,roster);
 const sheets=[];
 for(let i=0;i<3;i++){
  const sheet=new AdminSheet('回答'+i,[['メール','長文回答','名前'],...Array.from({length:20},()=>['kid@example.com','回答本文',''])]);
  e.sheets.set(sheet.name,sheet);sheets.push(sheet);
  e.c.saveManagedRecord_({id:'form:'+i,kind:'form',courseId:i===2?'200':'100',responseSheetId:sheet.sheetId,input:{emailHeader:'メール',nameHeader:'名前'}});
 }
 return {...e,roster,answers:sheets};
}
test('name refresh shares roster and sheet list across courses, reads only needed columns and batches each contiguous run',()=>{
 const e=managedNames();let rosterReads=0,lists=0,writes=0;
 const range=e.roster.getRange.bind(e.roster);e.roster.getRange=(...args)=>{const r=range(...args),read=r.getDisplayValues;r.getDisplayValues=()=>{rosterReads++;return read();};return r;};
 const getSheets=e.ss.getSheets;e.ss.getSheets=()=>{lists++;return getSheets();};
 for(const sheet of e.answers){const getRange=sheet.getRange.bind(sheet);sheet.getRange=(r,c,h=1,w=1)=>{
  if(r>1){assert.equal(w,1);assert.notEqual(c,2);}
  return getRange(r,c,h,w);
 };sheet.beforeWrite=()=>writes++;}
 e.c.refreshAllManagedNames_();assert.equal(rosterReads,1);assert.equal(lists,1);assert.equal(writes,3);
 assert.equal(e.answers[0].rows[20][2],'生徒A');assert.equal(e.answers[2].rows[20][2],'生徒B');
 e.roster.rows[1][1]='変更後';e.answers[0].rows[1][2]='';
 e.c.refreshAllManagedNames_();assert.equal(rosterReads,2);assert.equal(e.answers[0].rows[1][2],'変更後');
});
test('name batching leaves manual names and blank-result formulas intact and rejects conflicting roster names',()=>{
 const e=scoringEnvironment();e.sheets.set('生徒一覧',new AdminSheet('生徒一覧',[
  ['メールアドレス','名前','コースID'],['a@example.com','=literal','100'],['dup@example.com','A','100'],['dup@example.com','B','100']
 ]));
 const sheet=new AdminSheet('回答',[['メール','名前','評価'],[' A@example.com ','',5],['a@example.com','手入力',4],['a@example.com','=""',3],['a@example.com','',2],['dup@example.com','',1],['','',0]]);
 const getRange=sheet.getRange.bind(sheet);let writes=0;sheet.beforeWrite=()=>writes++;
 sheet.getRange=(...args)=>{const r=getRange(...args),read=r.getDisplayValues;r.getDisplayValues=()=>read().map(row=>row.map(v=>v==='=""'?'':v));return r;};
 const result=e.c.fillManagedNames_({courseId:'100',input:{emailHeader:'メール',nameHeader:'名前'}},sheet);
 assert.deepEqual(plain(result),{updated:2,unmatched:1,skipped:3});assert.equal(writes,2);
 assert.deepEqual(sheet.rows.slice(1).map(r=>r[1]),["'=literal",'手入力','=""',"'=literal",'','']);
 assert.deepEqual(sheet.rows.slice(1).map(r=>r[2]),[5,4,3,2,1,0]);
});
test('a failed name batch can be retried without touching already populated names',()=>{
 const e=managedNames();let fail=true;e.answers[1].beforeWrite=()=>{if(fail)throw Error('write failed');};
 assert.throws(()=>e.c.refreshAllManagedNames_(),/write failed/);
 assert.equal(e.answers[0].rows[1][2],'生徒A');assert.equal(e.answers[1].rows[1][2],'');
 fail=false;e.c.refreshAllManagedNames_();assert.equal(e.answers[2].rows[20][2],'生徒B');
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
