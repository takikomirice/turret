import test from 'node:test';
import assert from 'node:assert/strict';
import {adminEnvironment, AdminSheet} from './helpers/admin-environment.mjs';
import {scoringWorkload} from './helpers/scoring-workload.mjs';

test('admin display reads each grid once and action confirmation still rejects later edits',()=>{
  const e=adminEnvironment(),reads={};
  for(const name of ['送信シート','評価データ','エラー']) {
    const sheet=new AdminSheet(name,[['送信状態','本文'],['未','before']]);e.sheets.set(name,sheet);
    const getRange=sheet.getRange.bind(sheet);
    sheet.getRange=(...args)=>{const range=getRange(...args),read=range.getDisplayValues;range.getDisplayValues=()=>{reads[name]=(reads[name]||0)+1;return read();};return range;};
  }
  const data=e.c.getAdminConsoleData();
  assert.deepEqual(reads,{'送信シート':1,'評価データ':1,'エラー':1});
  for(const action of ['send','clearEval','clearSend','clearErrors'])assert.equal(data.actionRevisions[action],e.c.adminActionRevision_(action,data.config));
  e.sheets.get('送信シート').rows[1][1]='after';
  assert.throws(()=>e.c.runAdminAction('send',{confirmed:true,revision:data.actionRevisions.send},data.revision),/変更|確認/);
});

test('import reuses duplicate-check raw grids and observes new answers on the next invocation',()=>{
  const e=scoringWorkload(500,5);let raw=0;
  const roster=new AdminSheet('生徒一覧',[['メールアドレス','名前','クラス名','コースID','studentId']]);
  for(let course=0;course<5;course++)for(let i=0;i<100;i++)roster.rows.push([`student-${i}@example.invalid`,'架空の生徒','クラス'+course,'course-'+course,'student-'+i]);
  e.sheets.set('生徒一覧',roster);
  const getRange=e.answer.getRange.bind(e.answer);
  e.answer.getRange=(r,c,h=1,w=1)=>{const range=getRange(r,c,h,w),read=range.getValues;range.getValues=()=>{if(r===1&&h>1&&w===10)raw++;return read();};return range;};
  e.c.importFromFormsToEvalUnlocked_();
  assert.equal(raw,1);assert.equal(e.sheets.get('送信シート').rows.length,501);
  const added=[...e.answer.rows[1]];added[0]='new@example.invalid';added[2]='';e.answer.rows.push(added);roster.rows.push(['new@example.invalid','架空の生徒','クラス0','course-0','new-student']);
  e.c.importFromFormsToEvalUnlocked_();
  assert.equal(raw,2);assert.equal(e.sheets.get('送信シート').rows.length,502);
});

test('100 contiguous grading rows use bounded writes and checkpoints while preserving all values',()=>{
  const e=scoringWorkload(500,5);e.c.scoringManagementColumn_(e.answer,true);let flushes=0,writes=0;
  e.c.SpreadsheetApp.flush=()=>{flushes++;};
  const getRange=e.answer.getRange.bind(e.answer);
  e.answer.getRange=(...args)=>{const range=getRange(...args),write=range.setValues;range.setValues=values=>{writes++;return write(values);};return range;};
  const result=e.c.scoringSaveRows({target:e.target,rows:Array.from({length:100},(_,i)=>({rowNumber:i+2,scores:[i,'','','',''],comment:'講評 '+i}))});
  assert.equal(result.ok,true);assert.equal(result.results.length,100);
  for(let i=0;i<100;i++){assert.equal(e.answer.rows[i+1][e.answer.rows[0].indexOf('点数')],i);assert.equal(e.answer.rows[i+1][e.answer.rows[0].indexOf('講評')],'講評 '+i);}
  assert.ok(flushes<=4,`flushes: ${flushes}`);assert.ok(writes<=4,`answer writes: ${writes}`);
  assert.equal(e.answer.rows.slice(1,101).every(r=>JSON.parse(r[e.answer.rows[0].indexOf('管理')]).save.state==='complete'),true);assert.equal(e.sheets.has('_scoring_saves'),false);
});
