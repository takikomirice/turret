import test from 'node:test';
import assert from 'node:assert/strict';
import {scoringWorkload} from './helpers/scoring-workload.mjs';

function setup(lesson, timestamp, header='タイムスタンプ') {
 const e=scoringWorkload(1,1);
 e.answer.rows[0].push(header);e.answer.rows[1].push(timestamp);e.answer.rows[1][5]=lesson;
 // GASの日付整形だけを代替し、シートのタイムゾーン指定も検査する。
 e.c.Utilities.formatDate=(date,timezone,pattern)=>{
  assert.equal(timezone,'Asia/Tokyo');assert.equal(pattern,'yyyy-MM-dd HH:mm:ss');
  const parts=Object.fromEntries(new Intl.DateTimeFormat('en-GB',{timeZone:timezone,year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit',second:'2-digit',hourCycle:'h23'}).formatToParts(date).map(p=>[p.type,p.value]));
  return `${parts.year}-${parts.month}-${parts.day} ${parts.hour}:${parts.minute}:${parts.second}`;
 };
 return e;
}

test('summary supplies calendar-day timing in sheet timezone without extra sheet reads',()=>{
 const e=setup('2026-09-29',new Date('2026-09-29T15:00:00Z'));
 const result=e.c.scoringGetInitData({target:e.target,reviewMode:'summary'});
 assert.deepEqual(JSON.parse(JSON.stringify(result.review.rows[2].timing)),{lessonDate:'2026-09-29',submittedAt:'2026-09-30 00:00:00',daysLate:1});
 assert.equal(result.review.rows[2].eligible,true);
 const metrics={...e.metrics};e.answer.rows[1].pop();e.answer.rows[0].pop();e.resetMetrics();
 e.c.scoringGetInitData({target:e.target,reviewMode:'summary'});
 for(const key of ['opens','valueReads','displayReads','timezones'])assert.equal(metrics[key],e.metrics[key],key);
});

test('timing handles same day, month/year boundaries, leap days, future lessons and malformed input',()=>{
 for(const [lesson,timestamp,days] of [
  ['2026-09-29','2026/09/29 23:59:59',0],
  ['2026-09-29','2026-10-01 00:00',2],
  ['2025-12-31','2026-01-01 00:00:00',1],
  ['2024-02-28','2024-03-01 10:00:00',2],
  ['2025-09-29','2026-09-29 10:00:00',365],
  ['2027-09-29','2026-09-29 10:00:00',-365],
  ['2026-02-30','2026-09-29 10:00:00',null],
  ['適当','2026-09-29 10:00:00',null],
  ['','2026-09-29 10:00:00',null],
  ['2026-09-29','',null],
  ['2026-09-29','2026-02-30 10:00:00',null],
  ['2026-09-29','2026-09-29 25:00:00',null],
  ['2026-09-29',new Date(NaN),null]
 ]) {
  const e=setup(lesson,timestamp,'Timestamp');
  const summary=e.c.scoringGetReview({target:e.target,reviewMode:'summary'}).rows[2];
  assert.equal(summary.timing.daysLate,days,`${lesson} / ${timestamp}`);
  const full=e.c.scoringGetReview({target:e.target}).rows[2];
  assert.deepEqual(summary.timing,full.timing);
 }
});

test('updated adoption summaries retain timing and lateness never blocks returns',()=>{
 const e=setup('2025-09-29','2026-09-29 10:00:00');e.answer.rows.push([...e.answer.rows[1]]);
 const group=e.c.scoringGetReview({target:e.target}).groups[0];
 const result=e.c.scoringChooseResponses({target:e.target,groupId:group.id,fingerprint:group.fingerprint,accepted:[group.responses[0].key],reviewMode:'summary'});
 assert.equal(result.review.rows[2].timing.daysLate,365);assert.equal(result.review.rows[2].eligible,true);
});
