import {performance} from 'node:perf_hooks';
import {scoringWorkload} from '../tests/helpers/scoring-workload.mjs';

for(const count of [500,5000]) {
  const e=scoringWorkload(count),start=performance.now();
  const result=e.c.scoringGetInitData({target:e.target,reviewMode:'summary'});
  const load={...e.metrics,bytes:Buffer.byteLength(JSON.stringify(result)),rows:result.rows.length,nodeMs:Math.round(performance.now()-start)};
  e.resetMetrics();const index=e.c.scoringBuildReturnGuard_();
  const guard={...e.metrics,eligible:[...index.responses.values()].filter(r=>r.eligible).length};
  console.log(JSON.stringify({answers:count,sheets:5,load,guard}));
  e.resetMetrics();e.c.importFromFormsToEvalUnlocked_();
  console.log(JSON.stringify({answers:count,import:{...e.metrics,imported:e.sheets.get('評価データ').getLastRow()-1}}));
}

const saved=scoringWorkload(500,5);let flushes=0,answerWrites=0;
saved.c.SpreadsheetApp.flush=()=>{flushes++;};
const getRange=saved.answer.getRange.bind(saved.answer);
saved.answer.getRange=(...args)=>{const range=getRange(...args),write=range.setValues;range.setValues=values=>{answerWrites++;return write(values);};return range;};
const result=saved.c.scoringSaveRows({target:saved.target,rows:Array.from({length:100},(_,i)=>({rowNumber:i+2,scores:[3,'','','',''],comment:'架空の講評'}))});
console.log(JSON.stringify({save:{rows:100,ok:result.ok,flushes,answerWrites}}));
