import test from 'node:test';
import assert from 'node:assert/strict';
import {scoringEnvironment,AdminSheet,sourceId} from './helpers/scoring-environment.mjs';
function setup(){
 const e=scoringEnvironment();
 e.c.scoringSetConfig({nameCol:'B',displayCols:['G'],scoreCols:['D','','','',''],commentCol:'E',lessonDateHeader:'授業の日付を入力してください'});
 e.sheets.set('生徒一覧',new AdminSheet('生徒一覧',[['No','メールアドレス','名前','クラス名','コースID','studentId'],[1,'student@example.com','生徒','授業A','course-a','student-a']]));
 return e;
}
const review=e=>e.c.scoringGetReview({target:e.target});
const duplicate=e=>e.answer.rows.push(e.answer.rows[0].map(h=>({'メール':'student@example.com','名前':'生徒','点数':4,'授業の日付を入力してください':'2026-09-29','回答':'再提出'}[h]??'')));
const save=e=>e.c.scoringSaveRows({target:e.target,rows:[{rowNumber:2,scores:[3,'','','',''],comment:'講評'}]});
const data=(sheet,row=2)=>JSON.parse(sheet.rows[row-1][sheet.rows[0].indexOf('管理')]);
test('adding management moves status to the right and preserves positional grading configuration and old drafts',()=>{
 const e=setup(),before=e.c.scoringGetInitData({target:e.target});
 const r=e.c.scoringSaveRows({target:e.target,rows:[{rowNumber:2,scores:[3,'','','',''],comment:'講評',configRevision:before.configRevision}]});
 assert.equal(r.ok,true);assert.deepEqual(e.answer.rows[0],['メール','名前','点数','講評','授業の日付を入力してください','回答','管理','状態']);
 assert.equal(e.answer.rows[1][2],3);assert.equal(e.answer.rows[1][3],'講評');assert.equal(e.answer.rows[1][7],'');
 const after=e.c.scoringGetInitData({target:e.target});assert.equal(after.config.scoreCols[0],'C');assert.equal(after.config.commentCol,'D');
 assert.equal(save(e).ok,true);assert.equal(e.answer.rows[1][2],3);
});
const choose=(e,indices)=>{const g=review(e).groups[0];return e.c.scoringChooseResponses({target:e.target,groupId:g.id,fingerprint:g.fingerprint,accepted:indices===null?null:indices.map(i=>g.responses[i].key)});};
test('grading and duplicate choices persist in one answer column without creating journal tabs',()=>{
 const e=setup();duplicate(e);choose(e,[1]);assert.equal(save(e).ok,true);
 assert.equal(data(e.answer).save.state,'complete');assert.equal(data(e.answer).duplicate.choice,'excluded');
 assert.equal(review(e).rows[3].state,'adopted');assert.equal(e.sheets.has('_scoring_saves'),false);assert.equal(e.sheets.has('_scoring_decisions'),false);
 assert.equal(e.answer.rows[0].filter(h=>h==='管理').length,1);
 assert.ok(e.c.scoringGetComparison({target:e.target,groupId:review(e).groups[0].id}).responses.every(r=>r.answers.every(a=>a.header!=='管理')));
});
test('a partially replaced duplicate decision across sheets holds the entire group until retry',()=>{
 const e=setup(),other=new AdminSheet('回答 2',e.answer.rows);e.source.sheets.set(other.name,other);
 choose(e,[0]);other.beforeWrite=values=>{if(values.some(r=>r.some(v=>typeof v==='string'&&v.includes('"duplicate"'))))throw Error('write denied');};
 assert.throws(()=>choose(e,[1]),/write denied/);
 assert.ok(review(e).groups[0].responses.every(r=>r.state==='pending'&&!r.eligible));
 other.beforeWrite=null;choose(e,[1]);assert.equal(review(e).groups[0].responses[1].state,'adopted');
});
test('clearing a choice and receiving another answer invalidate the whole choice',()=>{
 const e=setup();duplicate(e);choose(e,[0]);choose(e,null);assert.equal(review(e).rows[2].state,'pending');
 choose(e,[0]);duplicate(e);assert.equal(review(e).rows[2].state,'pending');
});
test('malformed management cells cannot be overwritten or returned',()=>{
 const e=setup();e.answer.rows[0].push('管理');e.answer.rows[1].push('{broken');
 assert.equal(save(e).ok,false);assert.equal(e.answer.rows[1][3],'');assert.throws(()=>review(e),/管理/);
});
test('preparing sheets migrates old journals with readback, keeps choices and removes only migrated tabs',()=>{
 const e=setup();duplicate(e);const g=review(e).groups[0];
 e.sheets.set('_scoring_saves',new AdminSheet('_scoring_saves',[['key','JSON'],[JSON.stringify([sourceId,e.answer.sheetId,2]),'{"state":"pending"}']]));
 e.sheets.set('_scoring_decisions',new AdminSheet('_scoring_decisions',[['key','JSON'],[g.id,JSON.stringify({fingerprint:g.fingerprint,accepted:[g.responses[1].key]})]]));
 e.c.initializeSheetsUnlocked_();assert.equal(e.sheets.has('フォーム管理'),true);
 assert.equal(e.sheets.has('_scoring_saves'),false);assert.equal(e.sheets.has('_scoring_decisions'),false);
 assert.equal(review(e).rows[2].saveFailed,true);assert.equal(review(e).rows[3].state,'adopted');
 assert.equal(save(e).ok,true);assert.equal(review(e).rows[3].state,'adopted');
});
test('migration readback failure retains legacy records and retries without losing newer column data',()=>{
 const e=setup(),key=JSON.stringify([sourceId,e.answer.sheetId,2]);
 e.sheets.set('_scoring_saves',new AdminSheet('_scoring_saves',[['key','JSON'],[key,'{"state":"pending"}']]));
 const range=e.answer.getRange.bind(e.answer);e.answer.getRange=(...args)=>{const r=range(...args),set=r.setValues;r.setValues=values=>values.some(row=>row.some(v=>typeof v==='string'&&v.includes('"save"')))?r:set(values);return r;};
 assert.throws(()=>e.c.initializeSheetsUnlocked_(),/確認|保存/);assert.equal(e.sheets.has('_scoring_saves'),true);
 e.answer.getRange=range;assert.equal(save(e).ok,true);e.c.initializeSheetsUnlocked_();assert.equal(data(e.answer).save.state,'complete');assert.equal(e.sheets.has('_scoring_saves'),false);
});
test('management is reserved for internal data in output columns and reply fields',()=>{
 const e=setup();save(e);
 e.c.scoringSetConfig({scoreCols:['G','','','',''],commentCol:''});assert.throws(()=>save(e),/管理|保存/);
 const cfg=e.c.getConfig_();cfg.fields.push({key:'internal',sourceHeader:'管理',type:'text'});assert.throws(()=>e.c.validateConfigDraft_(cfg),/管理/);
});
test('missing legacy targets and invalid records prevent deleting either journal',()=>{
 const e=setup();e.sheets.set('_scoring_saves',new AdminSheet('_scoring_saves',[['key','JSON'],[JSON.stringify(['missing',123,2]),'{"state":"pending"}']]));
 e.sheets.set('_scoring_decisions',new AdminSheet('_scoring_decisions',[['key','JSON']]));
 assert.throws(()=>e.c.initializeSheetsUnlocked_(),/移行先/);assert.equal(e.sheets.has('_scoring_saves'),true);assert.equal(e.sheets.has('_scoring_decisions'),true);
});
test('management layout resumes after a column move fails without misdirecting old drafts',()=>{
 const e=setup(),before=e.c.scoringGetInitData({target:e.target}),move=e.answer.moveColumns.bind(e.answer);let fail=true;
 e.answer.moveColumns=(...args)=>{if(fail){fail=false;throw Error('move failed');}return move(...args);};
 assert.throws(()=>save(e),/move failed/);
 assert.equal(e.c.scoringSaveRows({target:e.target,rows:[{rowNumber:2,scores:[4,'','','',''],comment:'再保存',configRevision:before.configRevision}]}).ok,true);
 assert.equal(e.answer.rows[1][e.answer.rows[0].indexOf('点数')],4);assert.equal(e.answer.rows[1][e.answer.rows[0].indexOf('講評')],'再保存');
});
test('choices span separate registered workbooks and grading in one preserves the choice',()=>{
 const e=setup(),secondId='source_second_123456789012345',second=e.addBook(secondId),other=new AdminSheet('回答 2',e.answer.rows);second.sheets.set(other.name,other);
 e.configure({formSources:[{id:sourceId},{id:secondId}]});choose(e,[1]);save(e);
 assert.equal(review(e).groups[0].responses.filter(r=>r.state==='adopted').length,1);
 assert.equal(data(other).duplicate.choice,'adopted');assert.equal(e.sheets.has('_scoring_decisions'),false);
});

test('legacy duplicate fingerprints retain blank headers and repeated literal header names during column moves',()=>{
 const e=setup(),headers=['回答','','回答','回答 (2)','状態'],row=['a','blank','b','c',''];
 const expected=['a','blank','b','c',''];
 assert.deepEqual(Array.from(e.c.scoringFingerprintValues_(row,headers,new Set([4]),null)),expected);
 const moved=['回答','','回答','回答 (2)','管理','状態'];
 assert.deepEqual(Array.from(e.c.scoringFingerprintValues_(['a','blank','b','c','{}',''],moved,new Set([5]),{headers})),expected);
});

test('malformed header layout bindings cannot override grading configuration',()=>{
 const e=setup();save(e);const range=e.answer.getRange(1,e.answer.rows[0].indexOf('管理')+1),note=JSON.parse(range.getNote());
 note.bindings.startRow=999;range.setNote(JSON.stringify(note));
 assert.throws(()=>save(e),/メモ/);
});
