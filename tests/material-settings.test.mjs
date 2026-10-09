import test from 'node:test';
import assert from 'node:assert/strict';
import {schedulingEnvironment} from './helpers/scheduling-environment.mjs';

function environment(){
 const e=schedulingEnvironment(),r=e.read();
 r.label='1A';r.input.titlePattern='{class_label} フォーム';r.input.materialTitlePattern='{class_label} 資料';r.materialTitle='1A 資料';r.materialSettingsSaved=false;e.c.saveManagedRecord_(r);
 e.save=(changes={materialTitlePattern:'{class_label} 資料',description:'回答してください'})=>{
  const p=e.c.previewManagedFormSettings({targets:[{id:r.id,revision:e.read().revision}],changes,saveMaterialSettings:true});
  return e.c.runManagedFormSettings(r.id,changes,p.targets[0].fingerprint,true);
 };return e;
}
test('unsaved posting settings refuse immediate and scheduled publishing before opening the form',()=>{
 const e=environment();
 for(const action of ['publish','schedule'])assert.throws(()=>e.run(action,{scheduledTime:'2026-10-02T01:00:00Z'}),/保存/);
 assert.equal(e.form.published,false);assert.equal(e.calls.length,0);
});
test('explicit save persists unchanged posting settings and permits publishing after reopening',()=>{
 const e=environment();assert.equal(e.save().materialSettingsSaved,true);
 assert.equal(e.read().materialSettingsSaved,true);e.run('publish');assert.equal(e.posts.size,1);
});
test('saving only the posting title and an empty message preserves form settings and scheduled time',()=>{
 const e=environment();e.save();e.run('schedule',{scheduledTime:'2026-10-02T01:00:00Z'});const id=e.read().materialId;
 const r=e.save({materialTitlePattern:'{class_label} 新しい資料',description:''});
 assert.equal(r.materialTitle,'1A 新しい資料');assert.equal(r.input.description,'');assert.equal(r.title,'フォーム');
 assert.equal(e.posts.get(id).title,'1A 新しい資料');assert.equal(e.posts.get(id).description,'');assert.equal(e.posts.get(id).scheduledTime,'2026-10-02T01:00:00.000Z');
});

for(const action of ['publish','schedule'])test('plain posting titles can be saved, posted and updated without changing identity: '+action,()=>{
 const e=environment();e.save({materialTitlePattern:'授業の振り返り'});
 e.run(action,action==='schedule'?{scheduledTime:'2026-10-02T01:00:00Z'}:{});
 const before=e.read(),postBefore={...e.posts.get(before.materialId)};
 assert.equal(postBefore.title,'授業の振り返り');assert.equal(e.calls.find(c=>c[0]==='create')[2],'100');
 const after=e.save({materialTitlePattern:'今週の振り返り'}),post=e.posts.get(after.materialId);
 assert.equal(after.materialTitle,'今週の振り返り');assert.equal(after.materialId,before.materialId);assert.equal(after.title,before.title);
 assert.equal(post.title,'今週の振り返り');assert.equal(post.state,postBefore.state);assert.equal(post.scheduledTime,postBefore.scheduledTime);
 assert.equal(e.calls.filter(c=>c[0]==='create').length,1);
});

test('posting-title validation still rejects empty and oversized titles before writing',()=>{
 const e=environment(),before=JSON.stringify(e.read());
 for(const title of ['', '   ', 'あ'.repeat(151)])assert.throws(()=>e.save({materialTitlePattern:title}),/タイトル|入力/);
 assert.equal(JSON.stringify(e.read()),before);assert.equal(e.calls.length,0);
});
test('failed saves remain unconfirmed and a lost response resumes without creating another post',()=>{
 const e=environment(),sheet=e.sheets.get('システム管理');sheet.beforeWrite=()=>{throw Error('storage failed');};
 assert.throws(()=>e.save(),/storage failed/);sheet.beforeWrite=null;assert.equal(e.read().materialSettingsSaved,false);
 e.save();e.run('publish');const patch=e.service.patch;e.service.patch=(...args)=>{patch(...args);throw Error('response lost');};
 assert.throws(()=>e.save({description:'新本文'}),/response lost/);assert.ok(e.read().settingsUpdate);
 assert.throws(()=>e.run('publish'),/途中|設定/);e.service.patch=patch;
 const r=e.c.resumeManagedFormSettings('f1',e.read().revision);assert.equal(r.materialSettingsSaved,true);assert.equal(r.input.description,'新本文');assert.equal(e.posts.size,1);
});
test('save confirmation cannot be borrowed from another payload or used for form access changes',()=>{
 const e=environment(),changes={description:'本文'};
 const p=e.c.previewManagedFormSettings({targets:[{id:'f1',revision:e.read().revision}],changes});
 assert.throws(()=>e.c.runManagedFormSettings('f1',changes,p.targets[0].fingerprint,true),/確認|更新/);
 assert.throws(()=>e.save({domains:'example.com'}),/投稿|設定/);
});
test('registered common text survives individual saves and cannot borrow ordinary save confirmation',()=>{
 const e=environment(),changes={materialTitlePattern:'{class_label} 共通資料',description:'共通本文'};
 const normal=e.c.previewManagedFormSettings({targets:[{id:'f1',revision:e.read().revision}],changes,saveMaterialSettings:true});
 assert.throws(()=>e.c.runManagedFormSettings('f1',changes,normal.targets[0].fingerprint,true,true),/確認|更新/);
 const common=e.c.previewManagedFormSettings({targets:[{id:'f1',revision:e.read().revision}],changes,saveMaterialSettings:true,registerCommonMaterial:true});
 e.c.runManagedFormSettings('f1',changes,common.targets[0].fingerprint,true,true);
 assert.equal(e.read().materialSettingsBaseline.description,'共通本文');
 e.save({description:'個別本文'});
 assert.equal(e.read().input.description,'個別本文');assert.equal(e.read().materialSettingsBaseline.description,'共通本文');
});
test('common text registration survives interrupted material updates and resume',()=>{
 const e=environment();e.save();e.run('publish');
 const changes={materialTitlePattern:'{class_label} 共通資料',description:'共通本文'};
 const p=e.c.previewManagedFormSettings({targets:[{id:'f1',revision:e.read().revision}],changes,saveMaterialSettings:true,registerCommonMaterial:true});
 const patch=e.service.patch;e.service.patch=(...args)=>{patch(...args);throw Error('response lost');};
 assert.throws(()=>e.c.runManagedFormSettings('f1',changes,p.targets[0].fingerprint,true,true),/response lost/);
 e.service.patch=patch;e.c.resumeManagedFormSettings('f1',e.read().revision);
 assert.equal(e.read().materialSettingsBaseline?.description,'共通本文');
});
