import {performance} from 'node:perf_hooks';
import assert from 'node:assert/strict';
import {scoringWorkload} from '../tests/helpers/scoring-workload.mjs';
import {AdminSheet} from '../tests/helpers/scoring-environment.mjs';

for(const count of [500,5000]) {
  const e=scoringWorkload(count),start=performance.now();
  const result=e.c.scoringGetInitData({target:e.target,reviewMode:'summary'});
  const load={...e.metrics,bytes:Buffer.byteLength(JSON.stringify(result)),rows:result.rows.length,nodeMs:Math.round(performance.now()-start)};
  e.resetMetrics();const index=e.c.scoringBuildReturnGuard_();
  const guard={...e.metrics,eligible:[...index.responses.values()].filter(r=>r.eligible).length};
  console.log(JSON.stringify({answers:count,sheets:5,load,guard}));
  // 現行の直接送信準備には、回答とクラスを照合する名簿が必要。
  const students=[['No','メールアドレス','出席番号（任意）','名前','クラス名','コースID','studentId']];
  [...e.source.sheets.values()].forEach((sheet,s)=>{
    [...new Set(sheet.rows.slice(1).map(row=>row[0]))].forEach((email,i)=>students.push([students.length,email,String(i+1),'架空の生徒','クラス'+s,'course-'+s,'student-'+s+'-'+i]));
  });
  e.sheets.set('生徒一覧',new AdminSheet('生徒一覧',students));
  e.resetMetrics();e.c.prepareSendDataUnlocked_();
  const prepared=e.sheets.get('送信シート').getLastRow()-1;
  assert.equal(prepared,count);
  assert.equal(e.sheets.has('評価データ'),false);
  console.log(JSON.stringify({answers:count,prepare:{...e.metrics,prepared}}));
}

const saved=scoringWorkload(500,5);let flushes=0,answerWrites=0;
saved.c.SpreadsheetApp.flush=()=>{flushes++;};
const getRange=saved.answer.getRange.bind(saved.answer);
saved.answer.getRange=(...args)=>{const range=getRange(...args),write=range.setValues;range.setValues=values=>{answerWrites++;return write(values);};return range;};
const result=saved.c.scoringSaveRows({target:saved.target,rows:Array.from({length:100},(_,i)=>({rowNumber:i+2,scores:[3,'','','',''],comment:'架空の講評'}))});
console.log(JSON.stringify({save:{rows:100,ok:result.ok,flushes,answerWrites}}));
