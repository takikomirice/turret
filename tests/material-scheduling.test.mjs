import test from 'node:test';
import assert from 'node:assert/strict';
import {plain} from './helpers/admin-environment.mjs';
import {schedulingEnvironment} from './helpers/scheduling-environment.mjs';

const scheduled={scheduledTime:'2026-10-01T10:00:00+09:00'};

test('scheduled material is a draft with a UTC deadline, while its form is already accepting',()=>{
 const e=schedulingEnvironment(),r=e.run('schedule',scheduled);
 assert.equal(r.stage,'scheduled');assert.equal(r.scheduledTime,'2026-10-01T01:00:00.000Z');
 assert.equal(e.form.accepting,true);assert.equal(e.form.published,true);
 assert.deepEqual(e.calls.find(c=>c[0]==='create')[1],{title:'資料',description:'回答してください',state:'DRAFT',scheduledTime:'2026-10-01T01:00:00.000Z',materials:[{link:{url:e.form.getPublishedUrl()}}]});
 assert.throws(()=>e.run('schedule',scheduled),/投稿|予約|確認/);assert.equal(e.calls.filter(c=>c[0]==='create').length,1);
});

for(const time of ['', 'yesterday', '2026-10-01T10:00', '2026-02-30T10:00:00+09:00', '2026-10-01T08:59:00+09:00'])test('invalid or past schedule is rejected without publication: '+time,()=>{
 const e=schedulingEnvironment();assert.throws(()=>e.run('schedule',{scheduledTime:time}),/日時|未来/);
 assert.equal(e.form.published,false);assert.equal(e.calls.length,0);
});

test('a changed preview time or time that expired during confirmation cannot publish',()=>{
 const e=schedulingEnvironment(),p=e.preview('schedule',scheduled);
 assert.throws(()=>e.c.runManagedFormAction('f1','schedule',p.fingerprint,true,{scheduledTime:'2026-10-01T11:00:00+09:00'}),/確認|更新/);
 e.advance(3600001);assert.throws(()=>e.c.runManagedFormAction('f1','schedule',p.fingerprint,true,scheduled),/未来/);
 assert.equal(e.form.published,false);
});

test('schedule changes and cancellation reuse the draft; cancellation leaves responses open',()=>{
 const e=schedulingEnvironment(),r=e.run('schedule',scheduled);
 e.run('reschedule',{scheduledTime:'2026-10-01T11:00:00+09:00'});
 assert.equal(e.posts.get(r.materialId).scheduledTime,'2026-10-01T02:00:00.000Z');
 const cancelled=e.run('cancel-schedule');assert.equal(cancelled.stage,'draft');assert.equal(e.posts.get(r.materialId).scheduledTime,undefined);
 assert.equal(e.form.accepting,true);assert.equal(e.calls.filter(c=>c[0]==='remove').length,0);
 const published=e.run('publish');assert.equal(published.materialId,r.materialId);assert.equal(published.stage,'published');
 assert.equal(e.calls.filter(c=>c[0]==='create').length,1);
});

test('publication is observed from Classroom, never inferred solely from the clock',()=>{
 const e=schedulingEnvironment(),r=e.run('schedule',scheduled);e.advance(3600001);
 assert.equal(e.c.reconcileManagedPost(r.id,e.read().revision).stage,'scheduled');
 e.posts.get(r.materialId).state='PUBLISHED';e.posts.get(r.materialId).alternateLink='https://classroom.google.com/post';
 const done=e.c.reconcileManagedPost(r.id,e.read().revision);assert.equal(done.stage,'published');assert.equal(done.postUrl,'https://classroom.google.com/post');
 assert.throws(()=>e.run('cancel-schedule'),/公開|予約/);
});

test('a material published since confirmation cannot be unscheduled',()=>{
 const e=schedulingEnvironment(),r=e.run('schedule',scheduled),p=e.preview('cancel-schedule');
 e.posts.get(r.materialId).state='PUBLISHED';
 assert.throws(()=>e.c.runManagedFormAction(r.id,'cancel-schedule',p.fingerprint,true,{}),/公開|確認|予約/);
 assert.equal(e.calls.filter(c=>c[0]==='patch').length,0);
});

test('lost create response is reconciled among drafts without recreating the material',()=>{
 const e=schedulingEnvironment(),create=e.service.create;e.service.create=(...args)=>{create(...args);throw Error('timeout');};
 assert.throws(()=>e.run('schedule',scheduled),/timeout/);assert.equal(e.read().stage,'publish_review');
 assert.throws(()=>e.run('schedule',scheduled),/確認|投稿|予約/);
 const r=e.c.reconcileManagedPost('f1',e.read().revision);assert.equal(r.stage,'scheduled');assert.equal(r.materialId,'m1');
 assert.ok(e.calls.find(c=>c[0]==='list')[2].courseWorkMaterialStates.includes('DRAFT'));
});

test('lost patch response keeps confirmation pending and can be reconciled',()=>{
 const e=schedulingEnvironment();e.run('schedule',scheduled);
 const patch=e.service.patch;e.service.patch=(...args)=>{patch(...args);throw Error('timeout');};
 assert.throws(()=>e.run('cancel-schedule'),/timeout/);assert.equal(e.read().stage,'schedule_review');
 assert.throws(()=>e.run('reschedule',scheduled),/照合|確認/);
 const r=e.c.reconcileManagedPost('f1',e.read().revision);assert.equal(r.stage,'draft');assert.equal(r.scheduledTime,'');
});

test('unknown public state and failed journal writes never trigger another create',()=>{
 const e=schedulingEnvironment();const save=e.c.saveManagedRecord_;
 e.c.saveManagedRecord_=r=>{if(r.stage==='scheduled')throw Error('disk');return save(r);};
 assert.throws(()=>e.run('schedule',scheduled),/disk/);assert.equal(e.read().stage,'publish_review');
 e.c.saveManagedRecord_=save;const r=e.c.reconcileManagedPost('f1',e.read().revision);
 assert.equal(r.stage,'scheduled');assert.equal(e.calls.filter(c=>c[0]==='create').length,1);
});

test('scheduled forms are prepared, and progress distinguishes reservation from publication',()=>{
 const e=schedulingEnvironment();e.run('schedule',scheduled);
 const progress=plain(e.c.getFormSetupProgress_(true,e.c.getManagedRecords_()));
 assert.equal(progress.forms.state,'complete');assert.match(progress.publish.detail,/予約/);assert.match(progress.publish.detail,/0.*投稿済み|投稿済み.*0/);
});
