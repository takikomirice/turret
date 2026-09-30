import test from 'node:test';
import assert from 'node:assert/strict';
import {schedulingEnvironment} from './helpers/scheduling-environment.mjs';

const deadline={closesAt:'2026-10-01T11:00:00+09:00'};
test('deadline creates a separate shared clock without enabling daily return',()=>{
 const e=schedulingEnvironment();e.run('set-close',deadline);e.run('set-close',{closesAt:'2026-10-01T12:00:00+09:00'});
 assert.equal(e.read().closesAt,'2026-10-01T03:00:00.000Z');assert.equal(e.triggers.length,1);
 assert.equal(e.c.getAutomationSummary_().enabled,false);assert.equal(e.form.published,false);
});
test('clock closes only due forms, preserves answers and never calls Classroom',()=>{
 const e=schedulingEnvironment();e.run('publish');e.run('set-close',deadline);const before=e.calls.length;
 e.fire();assert.equal(e.form.accepting,true);e.advance(7200000);e.fire();
 assert.equal(e.form.accepting,false);assert.equal(e.read().closeState,'closed');assert.equal(e.calls.length,before);
 assert.throws(()=>e.run('open'),/終了|期限/);
 e.run('cancel-close');e.run('open');assert.equal(e.form.accepting,true);
});
test('a cancelled deadline never closes the form; changing the deadline uses the latest value',()=>{
 const e=schedulingEnvironment();e.run('publish');e.run('set-close',deadline);
 e.run('set-close',{closesAt:'2026-10-01T12:00:00+09:00'});e.advance(7200000);e.fire();assert.equal(e.form.accepting,true);
 e.run('cancel-close');e.advance(7200000);e.fire();assert.equal(e.form.accepting,true);
});
test('close failure is visible and retried; already closed forms need no second mutation',()=>{
 const e=schedulingEnvironment();e.run('publish');e.run('set-close',deadline);e.advance(7200000);
 const close=e.form.setAcceptingResponses;e.form.setAcceptingResponses=()=>{throw Error('Forms unavailable');};e.fire();
 assert.equal(e.read().closeState,'error');assert.match(e.read().closeError,/Forms unavailable/);
 e.form.setAcceptingResponses=close;e.fire();assert.equal(e.read().closeState,'closed');
 e.form.setAcceptingResponses=()=>{throw Error('should not repeat');};e.fire();assert.equal(e.read().closeState,'closed');
});
test('obsolete or foreign trigger invocations cannot close forms',()=>{
 const e=schedulingEnvironment();e.run('publish');e.run('set-close',deadline);e.advance(7200000);
 e.c.managedFormScheduleTick_({triggerUid:'foreign'});assert.equal(e.form.accepting,true);
 e.userProps.clear();e.fire();assert.equal(e.form.accepting,true);
 assert.throws(()=>e.run('set-close',{closesAt:'2026-10-01T13:00:00+09:00'}),/管理者/);
});
test('intake deadline must follow material publication in either editing direction',()=>{
 const e=schedulingEnvironment();e.run('schedule',{scheduledTime:'2026-10-01T12:00:00+09:00'});
 assert.throws(()=>e.run('set-close',deadline),/公開|投稿/);
 e.run('set-close',{closesAt:'2026-10-01T13:00:00+09:00'});
 assert.throws(()=>e.run('reschedule',{scheduledTime:'2026-10-01T14:00:00+09:00'}),/終了/);
});

test('a failed reschedule cannot substitute its intended time for the actual Classroom publication',()=>{
 const e=schedulingEnvironment();e.run('schedule',{scheduledTime:'2026-10-01T12:00:00+09:00'});
 e.service.patch=()=>{throw Error('request rejected');};
 assert.throws(()=>e.run('reschedule',{scheduledTime:'2026-10-01T10:00:00+09:00'}),/request rejected/);
 assert.throws(()=>e.run('set-close',{closesAt:'2026-10-01T11:00:00+09:00'}),/公開|照合/);
 assert.equal(e.read().closesAt,undefined);assert.equal(e.triggers.length,0);
});
test('failed timer creation leaves the deadline unset',()=>{
 const e=schedulingEnvironment();e.c.ScriptApp.newTrigger=()=>{throw Error('quota');};
 assert.throws(()=>e.run('set-close',deadline),/quota/);assert.equal(e.read().closesAt,undefined);
});
test('invalid and expired close dates cannot create a trigger or modify a form',()=>{
 const e=schedulingEnvironment();for(const closesAt of ['', '2026-02-30T10:00:00+09:00','2026-10-01T08:00:00+09:00'])assert.throws(()=>e.run('set-close',{closesAt}),/日時|未来/);
 assert.equal(e.triggers.length,0);assert.equal(e.form.published,false);
});

test('timer registration recognizes write-then-throw but an uncommitted orphan never runs',()=>{
 const e=schedulingEnvironment(),api=e.c.PropertiesService.getScriptProperties(),set=api.setProperty;
 e.c.PropertiesService.getScriptProperties=()=>({...api,setProperty(key,value){set(key,value);if(key==='TURRET_FORM_SCHEDULE_TRIGGER')throw Error('lost response');}});
 e.run('set-close',deadline);assert.equal(e.triggers.length,1);assert.equal(e.read().closeState,'scheduled');
 const f=schedulingEnvironment(),fapi=f.c.PropertiesService.getScriptProperties();
 f.c.PropertiesService.getScriptProperties=()=>({...fapi,setProperty(key,value){if(key==='TURRET_FORM_SCHEDULE_TRIGGER')throw Error('disk');fapi.setProperty(key,value);}});
 f.c.ScriptApp.deleteTrigger=()=>{throw Error('cleanup failed');};assert.throws(()=>f.run('set-close',deadline),/disk/);
 f.form.accepting=true;f.advance(7200000);f.fire();assert.equal(f.form.accepting,true);assert.equal(f.read().closesAt,undefined);
});

test('clock stops other forms after one failure and retries failed result persistence',()=>{
 const e=schedulingEnvironment();e.run('publish');e.run('set-close',deadline);
 const other=e.c.saveManagedRecord_({...e.read(),id:'f2',formId:'f2',revision:undefined});
 const form2={...e.form};e.c.FormApp.openById=id=>{if(id==='f')throw Error('inaccessible');return form2;};
 e.advance(7200000);e.fire();assert.equal(e.read().closeState,'error');assert.equal(e.c.loadManagedRecord_(other.id).closeState,'closed');assert.equal(form2.accepting,false);
});
