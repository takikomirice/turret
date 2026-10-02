import test from 'node:test';
import assert from 'node:assert/strict';
import {schedulingEnvironment} from './helpers/scheduling-environment.mjs';
import {AdminSheet, plain} from './helpers/admin-environment.mjs';

function environment(){
 const e=schedulingEnvironment(),{c,form}=e;
 form.title='1A 振り返り';form.getTitle=()=>form.title;form.setTitle=value=>{form.title=value;};form.collectsEmail=()=>true;
 let fileName=form.title;c.DriveApp={getFileById:()=>({getName:()=>fileName,setName:value=>{fileName=value;}})};
 const permissions=[{id:'school',view:'published',type:'domain',role:'reader',domain:'example.com'}];
 c.Drive.Permissions={list:()=>({permissions:plain(permissions)}),remove:(id,pid)=>{permissions.splice(permissions.findIndex(p=>p.id===pid),1);},create:(p)=>{permissions.push({...plain(p),id:'new-'+permissions.length});}};
 e.sheets.set('回答 1A',new AdminSheet('回答 1A',[['メール','質問','点数'],['kid@example.com','回答',3]]));
 const r=e.read();Object.assign(r,{label:'1A',sheetName:'回答 1A',title:form.title,materialTitle:'1A 資料',input:{templateId:'template-1234567890',folderId:'folder-1234567890',titlePattern:'{class_label} 振り返り',materialTitlePattern:'{class_label} 資料',description:'回答してください',policy:{domains:['example.com'],emails:[]},labels:{100:'1A'},emailHeader:'メール',nameHeader:'名前',statusHeader:'状態',prefix:'回答',additionalColumns:[]}});c.saveManagedRecord_(r);
 e.previewSettings=changes=>c.previewManagedFormSettings({targets:[{id:'f1',revision:e.read().revision}],changes});
 e.saveSettings=changes=>{const p=e.previewSettings(changes);return c.runManagedFormSettings('f1',changes,p.targets[0].fingerprint);};
 return {...e,permissions,fileName:()=>fileName};
}

test('editing description alone preserves form, answers, access and patches the existing material',()=>{
 const e=environment();e.run('publish');const r=e.read(),before=plain(e.sheets.get('回答 1A').rows);
 const p=e.previewSettings({description:''});assert.match(p.targets[0].changes.join('\n'),/回答してください/);
 const done=e.c.runManagedFormSettings(r.id,{description:''},p.targets[0].fingerprint);
 assert.equal(done.input.description,'');assert.equal(done.title,'1A 振り返り');assert.equal(done.materialId,r.materialId);assert.equal(e.posts.get(r.materialId).description,'');
 assert.equal(e.form.accepting,true);assert.deepEqual(e.sheets.get('回答 1A').rows,before);
 const patch=e.calls.find(c=>c[0]==='patch');assert.deepEqual(patch[1],{description:''});assert.equal(patch[4].updateMask,'description');
 assert.equal(e.calls.filter(c=>c[0]==='create').length,1);
});
test('short class labels change form and material titles but retain response sheet identity',()=>{
 const e=environment(),done=e.saveSettings({labels:{100:'1B'}});
 assert.equal(done.label,'1B');assert.equal(done.input.labels[100],'1B');assert.equal(done.title,'1B 振り返り');assert.equal(done.materialTitle,'1B 資料');
 assert.equal(e.form.title,'1B 振り返り');assert.equal(e.fileName(),'1B 振り返り');assert.equal(done.sheetName,'回答 1A');assert.ok(e.sheets.has('回答 1A'));
});
test('only selected patch fields are applied to targets with different existing settings',()=>{
 const e=environment(),r=plain(e.read());Object.assign(r,{id:'f2',courseId:'200',label:'2A',materialTitle:'2A 独自資料'});r.input.description='独自本文';r.input.policy.emails=['teacher@example.net'];e.c.saveManagedRecord_(r);
 const p=e.c.previewManagedFormSettings({targets:[{id:'f1',revision:e.read().revision},{id:'f2',revision:r.revision}],changes:{description:'新本文'}});
 for(const target of p.targets)e.c.runManagedFormSettings(target.id,{description:'新本文'},target.fingerprint);
 const b=e.c.loadManagedRecord_('f2');assert.equal(b.input.description,'新本文');assert.equal(b.materialTitle,'2A 独自資料');assert.deepEqual(plain(b.input.policy.emails),['teacher@example.net']);
});
test('stale records, external changes, invalid fields and unfinished updates stop before writes',()=>{
 const e=environment(),p=e.previewSettings({titlePattern:'{class_label} 新名称'});e.form.title='手動変更';
 assert.throws(()=>e.c.runManagedFormSettings('f1',{titlePattern:'{class_label} 新名称'},p.targets[0].fingerprint),/確認|更新/);assert.equal(e.form.title,'手動変更');
 for(const changes of [{folderId:'other'},{description:3},{titlePattern:'無置換'},{labels:{100:''}},{domains:'gmail.com'},{emails:'bad'},{description:'x'.repeat(5001)}])assert.throws(()=>e.previewSettings(changes));
 assert.throws(()=>e.c.previewManagedFormSettings({targets:[{id:'f1',revision:'stale'}],changes:{description:'a'}}),/更新/);
 const r=e.read();r.formUpdate={};e.c.saveManagedRecord_(r);assert.throws(()=>e.previewSettings({description:'a'}),/更新/);
});
test('access changes validate the roster, keep unpublished forms unpublished and preserve editor permissions',()=>{
 const e=environment();e.permissions.push({id:'editor',type:'user',role:'writer',emailAddress:'teacher@example.com'});
 assert.throws(()=>e.saveSettings({domains:'other.example'}),/生徒|回答/);
 const r=e.saveSettings({domains:'',emails:'kid@example.com teacher@example.net'});
 assert.equal(e.form.published,false);assert.equal(e.form.accepting,false);assert.deepEqual(plain(r.input.policy),{domains:[],emails:['kid@example.com','teacher@example.net']});
 assert.ok(e.permissions.some(p=>p.id==='editor'));assert.ok(!e.permissions.some(p=>p.id==='school'));
});
test('a lost material patch response is resumable without reposting or losing unrelated settings',()=>{
 const e=environment();e.run('publish');const patch=e.service.patch;let failed=false;
 e.service.patch=(...args)=>{const result=patch(...args);if(!failed){failed=true;throw Error('response lost');}return result;};
 assert.throws(()=>e.saveSettings({description:'新しい本文'}),/response lost/);assert.ok(e.read().settingsUpdate);
 assert.throws(()=>e.run('publish'),/設定|途中/);
 const r=e.c.resumeManagedFormSettings('f1',e.read().revision);assert.equal(r.settingsUpdate,undefined);assert.equal(r.input.description,'新しい本文');assert.equal(e.posts.get(r.materialId).description,'新しい本文');assert.equal(e.calls.filter(c=>c[0]==='create').length,1);
});
test('access failure remains stopped and a deadline reached during recovery cannot reopen the form',()=>{
 const e=environment();e.run('publish');const r=e.read();r.closesAt='2026-10-01T00:01:00Z';e.c.saveManagedRecord_(r);
 const create=e.c.Drive.Permissions.create;let failed=false;e.c.Drive.Permissions.create=(...args)=>{if(!failed){failed=true;throw Error('access failed');}return create(...args);};
 assert.throws(()=>e.saveSettings({domains:'',emails:'kid@example.com'}),/access failed/);assert.equal(e.form.accepting,false);assert.ok(e.read().settingsUpdate);
 e.advance(61000);const done=e.c.resumeManagedFormSettings('f1',e.read().revision);assert.equal(done.settingsUpdate,undefined);assert.equal(e.form.accepting,false);
});
test('a failed journal write makes no external changes and copied records are refused',()=>{
 const e=environment(),sheet=e.sheets.get('システム管理');const p=e.previewSettings({description:'new'});sheet.beforeWrite=()=>{throw Error('storage failed');};
 assert.throws(()=>e.c.runManagedFormSettings('f1',{description:'new'},p.targets[0].fingerprint),/storage failed/);assert.equal(e.form.title,'1A 振り返り');assert.equal(e.read().input.description,'回答してください');
 sheet.beforeWrite=null;e.c.ScriptApp.getScriptId=()=> 'copied';assert.throws(()=>e.previewSettings({description:'a'}),/別の運用/);
});
