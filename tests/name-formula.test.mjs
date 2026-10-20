import test from 'node:test';
import assert from 'node:assert/strict';
import {scoringEnvironment,AdminSheet} from './helpers/scoring-environment.mjs';

function environment(rows){
 const e=scoringEnvironment(),sheet=new AdminSheet('回答 名前数式',rows||[
  ['メール','名前','点数','講評'],['KID@EXAMPLE.COM','生徒','5','保持'],['missing@example.com','','0','未登録']
 ]);
 e.sheets.set(sheet.name,sheet);
 e.sheets.set('生徒一覧',new AdminSheet('生徒一覧',[
  ['メールアドレス','名前','コースID','除外(1)'],['kid@example.com','生徒','100',''],['kid@example.com','別クラス','200','']
 ]));
 const record=e.c.saveManagedRecord_({id:'managed-names',kind:'form',courseId:'100',responseSheetId:sheet.sheetId,input:{emailHeader:'メール',nameHeader:'名前'}});
 return {...e,sheet,record,target:{spreadsheetId:e.ss.getId(),sheetId:sheet.sheetId}};
}

test('formula setup migrates only the name column in one write and marks the managed form',()=>{
 const e=environment(),before=e.sheet.rows.map(r=>r.slice()),writes=[];e.sheet.beforeWrite=values=>writes.push(values);
 const result=e.c.scoringRefreshNames(e.target);
 assert.equal(result.formulaEnabled,true);assert.equal(writes.length,1);
 assert.match(e.sheet.rows[1][1],/^=ARRAYFORMULA\(LET\(turret_names_version,1,/);
 assert.equal(e.sheet.rows[2][1],'');
 assert.deepEqual(e.sheet.rows.map(r=>[r[0],...r.slice(2)]),before.map(r=>[r[0],...r.slice(2)]));
 assert.equal(e.c.loadManagedRecord_(e.record.id).nameFormulaVersion,1);
});

test('formula setup rejects hand-edited names and independent formulas without partial changes',()=>{
 for(const name of ['手入力の別名','=IF(TRUE,"","")']){
  const e=environment([['メール','名前','点数'],['kid@example.com',name,5]]),before=structuredClone(e.sheet.rows);
  assert.throws(()=>e.c.scoringRefreshNames(e.target),/2行|2.*行/);
  assert.deepEqual(e.sheet.rows,before);assert.equal(e.c.loadManagedRecord_(e.record.id).nameFormulaVersion,undefined);
 }
});

test('name formulas use exact forward/reverse lookups and an insertion-safe dynamic answer range',()=>{
 const e=environment();e.c.scoringRefreshNames(e.target);const formula=e.sheet.rows[1][1];
 assert.match(formula,/INDEX\(A:A,ROW\(\)\)/);assert.match(formula,/LOOKUP\(2,1\/\(A:A<>""\),ROW\(A:A\)\)/);
 assert.match(formula,/UNIQUE\(FILTER/);assert.match(formula,/TO_TEXT\('生徒一覧'!C2:C\)="100"/);
 assert.match(formula,/'生徒一覧'!D2:D/);assert.match(formula,/XLOOKUP.*"",0,1/);assert.match(formula,/XLOOKUP.*"",0,-1/);
 assert.match(formula,/EXACT\(firstnames,lastnames\)/);
});

test('existing managed formula is left intact by legacy filling and a repeated setup',()=>{
 const e=environment();e.c.scoringRefreshNames(e.target);const formula=e.sheet.rows[1][1];let writes=0;e.sheet.beforeWrite=()=>writes++;
 const result=e.c.scoringRefreshNames(e.target);assert.equal(result.formulaEnabled,true);assert.equal(writes,0);
 e.c.fillManagedNames_(e.record,e.sheet);assert.equal(writes,0);assert.equal(e.sheet.rows[1][1],formula);
});

test('a failed migration write keeps the original names and does not enable the formula mode',()=>{
 const e=environment(),before=structuredClone(e.sheet.rows);e.sheet.beforeWrite=()=>{throw Error('write failed');};
 assert.throws(()=>e.c.scoringRefreshNames(e.target),/write failed/);assert.deepEqual(e.sheet.rows,before);
 assert.equal(e.c.loadManagedRecord_(e.record.id).nameFormulaVersion,undefined);
});

test('the formula installer rejects missing or duplicate roster and response headers before writing',()=>{
 for(const broken of ['roster','duplicate','response']){
  const e=environment();if(broken==='roster')e.sheets.delete('生徒一覧');
  if(broken==='duplicate')e.sheets.get('生徒一覧').rows[0].push('名前');
  if(broken==='response')e.sheet.rows[0][0]='別メール';
  const before=structuredClone(e.sheet.rows);assert.throws(()=>e.c.scoringRefreshNames(e.target),/列|名簿|生徒一覧/);assert.deepEqual(e.sheet.rows,before);
 }
});

function responseEvent(e){
 e.sheet.getParent=()=>e.ss;
 e.props.set('TURRET_RESPONSE_TRIGGER',JSON.stringify({id:'trigger',owner:'owner-token',spreadsheetId:e.ss.getId()}));
 e.userProps.set('APP_AUTOMATION_OWNER','owner-token');
 return {triggerUid:'trigger',source:e.ss,range:{getSheet:()=>e.sheet,getRow:()=>2,getNumRows:()=>1}};
}

test('responses with an installed formula do no writes and need no lock even during another operation',()=>{
 const e=environment();e.c.scoringRefreshNames(e.target);const event=responseEvent(e);let writes=0;e.sheet.beforeWrite=()=>writes++;
 e.c.withAppLock_(()=>e.c.managedFormResponseReceived_(event));assert.equal(writes,0);
});

test('an initial response that moves the formula down restores its anchor with a bounded lock wait',()=>{
 const e=environment();e.c.scoringRefreshNames(e.target);const formula=e.sheet.rows[1][1];
 e.sheet.rows[1][1]='';e.sheet.rows[2][1]=formula;
 const event=responseEvent(e),waits=[],getLock=e.c.LockService.getScriptLock;
 e.c.LockService.getScriptLock=()=>{const lock=getLock(),tryLock=lock.tryLock;lock.tryLock=ms=>{waits.push(ms);return tryLock(ms);};return lock;};
 e.c.managedFormResponseReceived_(event);
 assert.deepEqual(waits,[30000]);assert.equal(e.sheet.rows[1][1],formula);assert.equal(e.sheet.rows[2][1],'');
});

test('an obsolete event or interrupted form update does not repair or write names',()=>{
 for(const blocked of ['trigger','formUpdate','settingsUpdate','columnSetup']){
  const e=environment();e.c.scoringRefreshNames(e.target);e.sheet.rows[1][1]='';const event=responseEvent(e);
  if(blocked==='trigger')event.triggerUid='old';else {const record=e.c.loadManagedRecord_(e.record.id);record[blocked]={};e.c.saveManagedRecord_(record);}
  let writes=0;e.sheet.beforeWrite=()=>writes++;e.c.managedFormResponseReceived_(event);assert.equal(writes,0);
 }
});

test('response repair consolidates array formulas copied by a table across two inserted rows',()=>{
 const e=environment();e.c.scoringRefreshNames(e.target);const formula=e.sheet.rows[1][1];
 e.sheet.rows=[['メール','名前','点数'],['kid@example.com',formula,1],['kid@example.com',formula,2],['kid@example.com',formula,3],['kid@example.com','生徒',4]];
 const getRange=e.sheet.getRange.bind(e.sheet);
 e.sheet.getRange=(...args)=>{const range=getRange(...args),display=range.getDisplayValues;range.getDisplayValues=()=>display().map((row,i)=>row.map(v=>String(v).startsWith('=ARRAYFORMULA(')?(e.sheet.rows[2][1]&&args[0]+i<4?'#REF!':'生徒'):v));return range;};
 e.c.managedFormResponseReceived_(responseEvent(e));
 assert.deepEqual(e.sheet.rows.slice(1).map(r=>r[1]),[formula,'','','']);
 assert.deepEqual(e.sheet.rows.slice(1).map(r=>r[2]),[1,2,3,4]);
});

test('a write response lost after successful storage is verified without clearing or repeating the migration',()=>{
 const e=environment(),getRange=e.sheet.getRange.bind(e.sheet);let writes=0;
 e.sheet.getRange=(...args)=>{const range=getRange(...args),set=range.setValues;range.setValues=values=>{writes++;set(values);throw Error('response lost');};return range;};
 assert.equal(e.c.scoringRefreshNames(e.target).formulaEnabled,true);assert.equal(writes,1);
});

test('a formula calculation failure restores the previous name values and preserves grading',()=>{
 const e=environment(),before=structuredClone(e.sheet.rows),getRange=e.sheet.getRange.bind(e.sheet);
 e.sheet.getRange=(...args)=>{const range=getRange(...args),display=range.getDisplayValues;range.getDisplayValues=()=>display().map(row=>row.map(v=>String(v).startsWith('=ARRAYFORMULA(')?'#REF!':v));return range;};
 assert.throws(()=>e.c.scoringRefreshNames(e.target),/設定できません/);assert.deepEqual(e.sheet.rows,before);
 assert.equal(e.c.loadManagedRecord_(e.record.id).nameFormulaVersion,undefined);
});

test('repair rollback restores a moved array anchor without writing its calculated spill values',()=>{
 const e=environment();e.c.scoringRefreshNames(e.target);const formula=e.sheet.rows[1][1];
 e.sheet.rows=[['メール','名前','点数'],['kid@example.com','',1],['kid@example.com',formula,2],['kid@example.com','生徒',3]];
 const getRange=e.sheet.getRange.bind(e.sheet);
 e.sheet.getRange=(...args)=>{const range=getRange(...args),display=range.getDisplayValues;range.getDisplayValues=()=>display().map((row,i)=>row.map((v,j)=>String(v).startsWith('=ARRAYFORMULA(')?(args[0]+i===2?'#REF!':'生徒'):v));return range;};
 assert.throws(()=>e.c.scoringRefreshNames(e.target),/設定できません/);
 assert.equal(e.sheet.rows[2][1],formula);assert.equal(e.sheet.rows[3][1],'');
 assert.deepEqual(e.sheet.rows.map(r=>r[2]),['点数',1,2,3]);
});
