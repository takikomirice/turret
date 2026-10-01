import test from 'node:test';
import assert from 'node:assert/strict';
import {scoringWorkload} from './helpers/scoring-workload.mjs';

test('summary loading reads each answer grid once and does not hash every unique response',()=>{
 const e=scoringWorkload(5000),result=e.c.scoringGetInitData({target:e.target,reviewMode:'summary'});
 assert.equal(result.rows.length,1000);assert.equal(result.review.detailsDeferred,true);
 assert.equal(e.metrics.opens,1);assert.ok(e.metrics.timezones<=5);assert.ok(e.metrics.cells<51000);assert.ok(e.metrics.digests<10);
 assert.ok(Buffer.byteLength(JSON.stringify(result))<2000000);
 assert.equal(result.review.rows[2].state,'single');assert.equal(result.review.rows[2].answers,undefined);
});

test('summary review and return guards agree; comparison loads only the chosen group with unchanged fingerprints',()=>{
 const e=scoringWorkload(500);e.answer.rows.push([...e.answer.rows[1]]);
 const full=e.c.scoringGetReview({target:e.target}),summary=e.c.scoringGetReview({target:e.target,reviewMode:'summary'});
 assert.equal(summary.detailsDeferred,true);assert.equal(summary.rows[2].duplicateCount,2);
 const group=full.groups[0],detail=e.c.scoringGetComparison({target:e.target,groupId:group.id});
 assert.equal(detail.fingerprint,group.fingerprint);assert.deepEqual(JSON.parse(JSON.stringify(detail)),JSON.parse(JSON.stringify(group)));
 e.resetMetrics();const guard=e.c.scoringBuildReturnGuard_();
 for(const item of guard.responses.values())if(item.target.sheetId===e.answer.sheetId)assert.equal(item.eligible,summary.rows[item.rowNumber].eligible);
 assert.ok(e.metrics.digests<10);assert.equal(e.metrics.timezones,5);
 assert.throws(()=>e.c.scoringGetComparison({target:e.target,groupId:'other'}),/重複|変更|回答/);
});

test('adoption can return the updated summary in the same request without returning all answer text',()=>{
 const e=scoringWorkload(500);e.answer.rows.push([...e.answer.rows[1]]);
 const group=e.c.scoringGetReview({target:e.target}).groups[0];
 const result=e.c.scoringChooseResponses({target:e.target,groupId:group.id,fingerprint:group.fingerprint,accepted:[group.responses[0].key],reviewMode:'summary'});
 assert.equal(result.review.rows[group.responses[0].rowNumber].state,'adopted');
 assert.equal(result.review.rows[group.responses[1].rowNumber].state,'excluded');
 assert.equal(result.review.rows[2].answers,undefined);
});

test('request-scoped reads never retain answers, eligibility or authorization across requests',()=>{
 const e=scoringWorkload(500);let res=e.c.scoringGetInitData({target:e.target,reviewMode:'summary'});assert.equal(res.rows[0].scores[0],'3');
 e.answer.rows[1][3]=5;e.answer.rows.push([...e.answer.rows[1]]);
 res=e.c.scoringGetInitData({target:e.target,reviewMode:'summary'});assert.equal(res.rows[0].scores[0],5);assert.equal(res.review.rows[2].state,'pending');
 e.c.Session.getActiveUser=()=>({getEmail:()=> 'other@example.invalid'});
 assert.throws(()=>e.c.scoringGetComparison({target:e.target,groupId:res.review.rows[2].groupId}),/本人/);
});

test('comparison batches display reads per sheet even for many duplicate submissions',()=>{
 const e=scoringWorkload(5,1);for(let i=0;i<100;i++)e.answer.rows.push([...e.answer.rows[1]]);const review=e.c.scoringGetReview({target:e.target,reviewMode:'summary'});e.resetMetrics();const detail=e.c.scoringGetComparison({target:e.target,groupId:review.rows[2].groupId});assert.equal(detail.responses.length,101);assert.ok(e.metrics.displayReads<=2);
});
