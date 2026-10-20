import test from 'node:test';
import assert from 'node:assert/strict';
import {scoringEnvironment,AdminSheet,sourceId} from './helpers/scoring-environment.mjs';
function setup(){
 const e=scoringEnvironment();
 e.c.scoringSetConfig({nameCol:'B',displayCols:['G'],scoreCols:['D','','','',''],commentCol:'E',lessonDateHeader:'授業の日付を入力してください'});
 e.sheets.set('生徒一覧',new AdminSheet('生徒一覧',[['No','メールアドレス','名前','クラス名','コースID','studentId'],[1,'student@example.com','生徒','授業A','course-a','student-a']]));
 e.sheets.set('対応表',new AdminSheet('対応表',[['元SS_ID','元シート名','クラス名','courseId','メモ'],[sourceId,'回答 1','授業A','course-a','']]));
 e.answer.rows[0].push('タイムスタンプ');e.answer.rows[1][3]=3;e.answer.rows[1].push('2026-09-29 10:00');
 e.posts=[];e.c.Classroom={Courses:{Announcements:{create(data,courseId){e.posts.push({data,courseId});return {id:'post-'+e.posts.length};}}}};
 return e;
}
const repeat=e=>e.answer.rows.push(e.answer.rows[0].map(h=>({'メール':' STUDENT@example.com ','名前':'生徒','点数':4,'授業の日付を入力してください':'2026-09-29','回答':'再提出','タイムスタンプ':'2026-09-30 10:00'}[h]??'')));
const review=e=>e.c.scoringGetReview({target:e.target});
function choose(e,indices){const group=review(e).groups[0];return e.c.scoringChooseResponses({target:e.target,groupId:group.id,fingerprint:group.fingerprint,accepted:indices.map(i=>group.responses[i].key)});}
const send=e=>{e.c.importFromFormsToEval();e.c.evalToSendSheet();e.c.sendMessagesUnlocked_();};

function remind(e){
 const messages=[];e.c.GmailApp={sendEmail:(to,subject,body)=>messages.push({to,subject,body})};
 e.c.remindUngradedAndErrors();return messages;
}

test('a returned adopted answer and an ungraded excluded answer need no reminder',()=>{
 const e=setup();repeat(e);e.answer.rows[2][3]='';choose(e,[0]);send(e);
 assert.equal(e.posts.length,1);assert.equal(remind(e).length,0);
});

test('only adopted ungraded answers are counted after choosing among duplicates',()=>{
 const e=setup();repeat(e);e.answer.rows[1][3]='';e.answer.rows[2][3]='';choose(e,[1]);
 const [message]=remind(e);assert.match(message.body,/未採点　：1件/);
 assert.match(message.body,/重複の確認待ち　：0件/);
});

for(const graded of [false,true])test('unresolved duplicates share one reminder and are counted separately from ungraded answers: graded='+graded,()=>{
 const e=setup();repeat(e);if(!graded){e.answer.rows[1][3]='';e.answer.rows[2][3]='';}
 const messages=remind(e);assert.equal(messages.length,1);
 assert.match(messages[0].body,/回答ブック \/ 回答 1：未採点 0件, 未送信 0件, エラー 0件, 重複の確認待ち 2件/);
 assert.match(messages[0].body,/重複の確認待ち　：2件/);
 assert.match(messages[0].body,/採点画面/);
});

test('a new response invalidates the adoption decision and returns all three answers to review',()=>{
 const e=setup();repeat(e);choose(e,[0]);repeat(e);e.answer.rows[3][3]='';
 const [message]=remind(e);assert.match(message.body,/未採点　：0件/);
 assert.match(message.body,/重複の確認待ち　：3件/);
});

test('cross-tab duplicates use the shared adoption decision in reminders',()=>{
 const e=setup();const other=new AdminSheet('回答 2',e.answer.rows);other.rows[1][3]='';
 e.source.sheets.set(other.name,other);
 e.sheets.get('対応表').rows.push([sourceId,other.name,'授業A','course-a','']);
 const before=remind(e);assert.equal(before.length,1);assert.match(before[0].body,/重複の確認待ち　：2件/);
 const group=review(e).groups[0],accepted=group.responses.find(r=>r.target.sheetId===e.answer.sheetId);
 e.c.scoringChooseResponses({target:e.target,groupId:group.id,fingerprint:group.fingerprint,accepted:[accepted.key]});
 assert.equal(remind(e).length,0);
});

test('ordinary ungraded answers and disabled duplicate checks retain their original reminder behavior',()=>{
 const e=setup();e.answer.rows[1][3]='';
 assert.match(remind(e)[0].body,/未採点　：1件/);
 e.c.scoringSetConfig({...e.c.scoringGetConfig(),lessonDateHeader:''});repeat(e);e.answer.rows[2][3]='';
 const [message]=remind(e);assert.match(message.body,/未採点　：2件/);assert.match(message.body,/重複の確認待ち　：0件/);
});

test('accepting both duplicates counts both missing scores but treats zero as graded',()=>{
 const e=setup();repeat(e);e.answer.rows[1][3]='';e.answer.rows[2][3]='';choose(e,[0,1]);
 assert.match(remind(e)[0].body,/未採点　：2件/);
 const scoreCol=e.answer.rows[0].indexOf('点数');
 e.answer.rows[1][scoreCol]=0;assert.match(remind(e)[0].body,/未採点　：1件/);
 e.answer.rows[2][scoreCol]=0;assert.equal(remind(e).length,0);
});

test('pending delivery and send errors remain visible alongside duplicate review',()=>{
 const e=setup();repeat(e);e.answer.rows[1][2]='準備〇';e.answer.rows[2][2]='送信エラー';
 const [message]=remind(e);assert.match(message.body,/未送信　：1件/);assert.match(message.body,/エラー　：1件/);
 assert.match(message.body,/重複の確認待ち　：2件/);
 choose(e,[0]);const [after]=remind(e);assert.match(after.body,/エラー　：1件/);assert.match(after.body,/重複の確認待ち　：0件/);
});

test('legacy forms without lesson dates keep ordinary ungraded reminders',()=>{
 const e=setup();e.props.delete('TURRET_SCORING_CONFIG');e.answer.rows[0][5]='自由記述';e.answer.rows[1][3]='';
 const [message]=remind(e);assert.match(message.body,/未採点　：1件/);assert.match(message.body,/重複の確認待ち　：0件/);
});

test('default lesson dates detect review waiting even before scoring settings are saved',()=>{
 const e=setup();e.props.delete('TURRET_SCORING_CONFIG');repeat(e);
 assert.match(remind(e)[0].body,/重複の確認待ち　：2件/);
});

test('reminders preserve answers and adoption records and stop on unreadable decisions',()=>{
 const e=setup();repeat(e);choose(e,[0]);
 const before=structuredClone(e.answer.rows);remind(e);assert.deepEqual(e.answer.rows,before);
 const column=e.answer.rows[0].indexOf('管理');assert.notEqual(column,-1);
 e.answer.rows[1][column]='broken';const messages=[];
 e.c.GmailApp={sendEmail:(...args)=>messages.push(args)};
 assert.throws(()=>e.c.remindUngradedAndErrors(),/管理列/);
 assert.equal(messages.length,0);assert.equal(e.isLocked(),false);
});

test('same student, class and lesson date are duplicates even when submitted on different days',()=>{
 const e=setup();repeat(e);const r=review(e);assert.equal(r.groups.length,1);assert.equal(r.groups[0].responses.length,2);
 assert.equal(r.rows[2].state,'pending');assert.equal(r.rows[3].state,'pending');send(e);assert.equal(e.posts.length,0);
 choose(e,[1]);send(e);assert.equal(e.posts.length,1);assert.match(e.posts[0].data.text,/B\+/);assert.equal(e.answer.rows.length,3);
});

test('ordinary responses need no adoption step and accepting both sends both once',()=>{
 const e=setup();assert.equal(review(e).rows[2].state,'single');send(e);assert.equal(e.posts.length,1);
 repeat(e);choose(e,[0,1]);send(e);assert.equal(e.posts.length,2);send(e);assert.equal(e.posts.length,2);
});

test('a duplicate arriving after preparation holds the old queued post until reviewed',()=>{
 const e=setup();e.c.importFromFormsToEval();e.c.evalToSendSheet();repeat(e);
 e.c.sendMessagesUnlocked_();assert.equal(e.posts.length,0);choose(e,[1]);send(e);assert.equal(e.posts.length,1);assert.match(e.posts[0].data.text,/B\+/);
 choose(e,[0,1]);send(e);assert.equal(e.posts.length,2);
});

test('accepting a newer response after return adds a post and preserves the old one',()=>{
 const e=setup();send(e);const original=e.posts[0];repeat(e);send(e);assert.equal(e.posts.length,1);
 assert.equal(review(e).groups[0].responses[0].returned,true);choose(e,[1]);send(e);
 assert.equal(e.posts.length,2);assert.equal(e.posts[0],original);
});

test('new third response invalidates an old decision and rejects stale comparison choices',()=>{
 const e=setup();repeat(e);choose(e,[0]);const group=review(e).groups[0];
 repeat(e);
 assert.throws(()=>e.c.scoringChooseResponses({target:e.target,groupId:group.id,fingerprint:group.fingerprint,accepted:[group.responses[0].key]}),/更新|変更/);
 assert.equal(review(e).rows[2].state,'pending');send(e);assert.equal(e.posts.length,0);
});

test('different dates and different classes are independent; selected blank or invalid dates are held',()=>{
 const e=setup();repeat(e);e.answer.rows[2][5]='2026-09-30';assert.equal(review(e).groups.length,0);
 for(const date of ['', '2026-02-30','not a date']){e.answer.rows[2][5]=date;assert.equal(review(e).rows[3].state,'invalid-date');}
 const other=new AdminSheet('回答 2',[...e.answer.rows.map(row=>[...row])]);other.rows=other.rows.slice(0,2);e.source.sheets.set(other.name,other);
 e.sheets.get('対応表').rows.push([sourceId,other.name,'授業B','course-b','']);assert.equal(review(e).groups.length,0);
});

test('without a date mapping duplication checking is explicitly disabled',()=>{
 const e=setup();e.c.scoringSetConfig({...e.c.scoringGetConfig(),lessonDateHeader:''});repeat(e);
 const r=review(e);assert.equal(r.enabled,false);assert.equal(r.rows[2].state,'disabled');send(e);assert.equal(e.posts.length,2);
});

test('a partially failed grading write cannot be returned until a successful retry',()=>{
 const e=setup();e.answer.beforeWrite=values=>{if(values.some(row=>row.includes('失敗')))throw Error('write failed');};
 const p={target:e.target,rows:[{rowNumber:2,scores:[5,'','','',''],comment:'失敗'}]};assert.equal(e.c.scoringSaveRows(p).ok,false);
 send(e);assert.equal(e.posts.length,0);e.answer.beforeWrite=null;p.rows[0].comment='成功';assert.equal(e.c.scoringSaveRows(p).ok,true);
 send(e);assert.equal(e.posts.length,1);
});

test('a managed answer outside the registered return prefix remains held even after adoption',()=>{
 const e=setup();const managed=new AdminSheet('対象外の回答',e.answer.rows);e.sheets.set(managed.name,managed);
 e.c.saveManagedRecord_({id:'outside-form',kind:'form',responseSheetId:managed.sheetId,courseId:'course-b'});
 const index=e.c.scoringBuildReviewIndex_();const item=[...index.responses.values()].find(r=>r.target.sheetId===managed.sheetId);
 assert.equal(item.eligible,false);
});

test('default lesson-date detection also holds duplicates before the first scoring settings save',()=>{
 const e=setup();e.props.delete('TURRET_SCORING_CONFIG');repeat(e);
 assert.equal(review(e).rows[2].state,'pending');send(e);assert.equal(e.posts.length,0);
});

test('the second identically named lesson-date column is resolved and protected from grading writes',()=>{
 const e=setup();e.answer.rows[0].push('授業の日付を入力してください');e.answer.rows[1][5]='bad';e.answer.rows[1].push('2026-09-29');
 e.c.scoringSetConfig({...e.c.scoringGetConfig(),lessonDateHeader:'授業の日付を入力してください (2)'});
 assert.equal(review(e).rows[2].state,'single');
 e.c.scoringSetConfig({...e.c.scoringGetConfig(),scoreCols:['I','','','','']});
 assert.throws(()=>e.c.scoringSaveRows({target:e.target,rows:[{rowNumber:2,scores:[1]}]}),/識別|保存/);
});
