import test from 'node:test';
import assert from 'node:assert/strict';
import {scoringEnvironment,AdminSheet} from './helpers/scoring-environment.mjs';
const config={nameCol:'B',displayCols:['G'],scoreCols:['D','','','',''],commentCol:'E',ruleOutputSlots:[],startRow:2,headerRow:1};
function setup(){const e=scoringEnvironment();e.c.scoringSetConfig(config);return e;}
const cell=(e,row,header)=>e.answer.rows[row-1][e.answer.rows[0].indexOf(header)];
const states=e=>e.answer.rows.slice(1).map((_,i)=>JSON.parse(cell(e,i+2,'管理')).save.state);
const prepare=e=>e.c.scoringManagementColumn_(e.answer,true);

test('a pending-state failure prevents answer writes and completion failure never reports success',()=>{
 for(const failedState of ['pending','complete']) {
  const e=setup();prepare(e);
  e.answer.beforeWrite=values=>{if(values.some(row=>row.some(v=>typeof v==='string' && v.startsWith('{') && JSON.parse(v).save?.state===failedState)))throw Error('management write failed');};
  const result=e.c.scoringSaveRows({target:e.target,rows:[{rowNumber:2,scores:[0,'','','',''],comment:'saved'}]});
  assert.equal(result.ok,false);assert.equal(result.results[0].ok,false);
  if(failedState==='pending')assert.equal(cell(e,2,'点数'),'');
  else {assert.equal(cell(e,2,'点数'),0);assert.deepEqual(states(e),['pending']);}
 }
});
test('every row is durably pending before any batch answer write and complete only after the answer checkpoint',()=>{
 const e=setup();e.answer.rows.push([...e.answer.rows[1]]);prepare(e);let writes=0,answerCheckpoint=false,pendingCheckpoint=false;
 e.c.SpreadsheetApp.flush=()=>{if(!writes){assert.deepEqual(states(e),['pending','pending']);pendingCheckpoint=true;}else answerCheckpoint=true;};
 e.answer.beforeWrite=values=>{if(values.some(row=>row.includes('saved'))){assert.equal(pendingCheckpoint,true);assert.deepEqual(states(e),['pending','pending']);writes++;}};
 const result=e.c.scoringSaveRows({target:e.target,rows:[2,3].map(rowNumber=>({rowNumber,scores:[0,'','','',''],comment:'saved'}))});
 assert.equal(result.ok,true);assert.equal(answerCheckpoint,true);assert.deepEqual(states(e),['complete','complete']);
});
test('an answer checkpoint failure keeps every affected row pending until retry',()=>{
 const e=setup();e.answer.rows.push([...e.answer.rows[1]]);prepare(e);let checkpoint=0;
 e.c.SpreadsheetApp.flush=()=>{if(++checkpoint===2)throw Error('checkpoint failed');};
 const payload={target:e.target,rows:[2,3].map(rowNumber=>({rowNumber,scores:[3,'','','',''],comment:'saved'}))};
 const result=e.c.scoringSaveRows(payload);assert.equal(result.ok,false);assert.equal(result.results.every(r=>!r.ok),true);
 assert.deepEqual(states(e),['pending','pending']);
 e.c.SpreadsheetApp.flush=()=>{};assert.equal(e.c.scoringSaveRows(payload).ok,true);
});

test('batch writes preserve holes, untouched formulas, zero, and the order of repeated row edits',()=>{
 const e=setup();e.answer.rows.push([...e.answer.rows[1]],[...e.answer.rows[1]]);
 e.answer.rows[1][3]='=1+2';e.answer.rows[2][3]='=2+2';e.answer.rows[3][3]=5;
 const result=e.c.scoringSaveRows({target:e.target,rows:[
  {rowNumber:4,scores:[0,'','','',''],comment:'first'},
  {rowNumber:2,scores:['','','','',''],comment:'kept'},
  {rowNumber:4,scores:['','','','',''],comment:'last'}
 ]});
 assert.equal(result.ok,true);assert.deepEqual(Array.from(result.results,r=>r.rowNumber),[4,2,4]);
 assert.equal(cell(e,2,'点数'),'=1+2');assert.equal(cell(e,3,'点数'),'=2+2');assert.equal(cell(e,4,'点数'),0);assert.equal(cell(e,4,'講評'),'last');
});

test('a failed later column segment leaves only its row pending and never writes intervening answer columns',()=>{
 const e=setup();e.answer.rows.push([...e.answer.rows[1]]);
 e.c.scoringSetConfig({...config,commentCol:'G'});
 e.answer.beforeWrite=values=>{if(values.some(row=>row.includes('fail')))throw Error('write failed');};
 const result=e.c.scoringSaveRows({target:e.target,rows:[
  {rowNumber:2,scores:[0,'','','',''],comment:'ok'},
  {rowNumber:3,scores:[1,'','','',''],comment:'fail'}
 ]});
 assert.deepEqual(Array.from(result.results,r=>r.ok),[true,false]);
 assert.equal(cell(e,2,'授業の日付を入力してください'),'2026-09-29');assert.equal(cell(e,3,'授業の日付を入力してください'),'2026-09-29');
 assert.deepEqual(states(e),['complete','pending']);
});

test('read and save use an authorized ID target, preserve zero and do not send',()=>{
 const e=setup();const data=e.c.scoringGetInitData({target:e.target});
 assert.equal(data.rows[0].displays[0],'回答内容');
 const result=e.c.scoringSaveRows({target:e.target,rows:[{rowNumber:2,scores:[0,'','','',''],comment:'講評'}]});
 assert.equal(result.ok,true);assert.equal(cell(e,2,'点数'),0);assert.equal(cell(e,2,'講評'),'講評');assert.equal(cell(e,2,'状態'),'');
 assert.equal(e.sheets.has('送信シート'),false);assert.equal(e.isLocked(),false);
});

test('blank scores preserve existing values while empty feedback clears it',()=>{
 const e=setup();e.answer.rows[1][3]=3;e.answer.rows[1][4]='以前';
 e.c.scoringSaveRows({target:e.target,rows:[{rowNumber:2,scores:['','','','',''],comment:''}]});
 assert.equal(cell(e,2,'点数'),3);assert.equal(cell(e,2,'講評'),'');
});

test('save refuses headers, fractional rows, out of bounds and unregistered targets',()=>{
 const e=setup();
 for(const rowNumber of [1,2.5,3]){const r=e.c.scoringSaveRows({target:e.target,rows:[{rowNumber,scores:[3,'','','',''],comment:'bad'}]});assert.equal(r.ok,false);}
 assert.equal(cell(e,2,'点数'),'');assert.equal(e.answer.rows.length,2);
 for(const api of ['scoringGetInitData','scoringSaveRows'])assert.throws(()=>e.c[api]({target:{spreadsheetId:'other',sheetId:e.answer.sheetId},rows:[]}),/対象|登録/);
});

test('partial save failures report individual rows and release the common lock',()=>{
 const e=setup();e.answer.rows.push(['other@example.com','別','','','','2026-09-29','別の回答']);
 e.answer.beforeWrite=values=>{if(values.some(row=>row.includes('失敗')))throw Error('sheet unavailable');};
 const r=e.c.scoringSaveRows({target:e.target,rows:[{rowNumber:2,scores:[3,'','','',''],comment:'成功'},{rowNumber:3,scores:[2,'','','',''],comment:'失敗'}]});
 assert.deepEqual(Array.from(r.results,x=>x.ok),[true,false]);assert.equal(r.ok,false);assert.equal(e.isLocked(),false);
});

test('saving while locked or as another viewer never reports success',()=>{
 const e=setup(),payload={target:e.target,rows:[{rowNumber:2,scores:[3,'','','',''],comment:'講評'}]};
 e.c.withAppLock_(()=>assert.throws(()=>e.c.scoringSaveRows(payload),/別の処理/));
 e.c.Session.getActiveUser=()=>({getEmail:()=> 'other@example.com'});assert.throws(()=>e.c.scoringSaveRows(payload),/本人/);
 assert.equal(e.answer.rows[1][3],'');
});

test('feedback beginning with an equals sign is saved as text rather than a sheet formula',()=>{
 const e=setup();e.c.scoringSaveRows({target:e.target,rows:[{rowNumber:2,scores:[3,'','','',''],comment:'=1+1'}]});
 assert.equal(e.answer.getRange(2,e.answer.rows[0].indexOf('講評')+1).getFormulas()[0][0],'');
});

test('legacy scoring journals remain read-only until explicit migration',()=>{
 const e=setup(),journal=new AdminSheet('_scoring_saves',[['key','JSON'],['prior','{"state":"complete"}']]);e.sheets.set(journal.name,journal);
 assert.equal(e.c.scoringSaveRows({target:e.target,rows:[{rowNumber:2,scores:[0,'','','',''],comment:'new'}]}).ok,true);
 assert.deepEqual(journal.rows[0],['key','JSON']);assert.deepEqual(states(e),['complete']);assert.equal(e.c.scoringReadRecords_('_scoring_saves').records.get('prior').data.state,'complete');
});
