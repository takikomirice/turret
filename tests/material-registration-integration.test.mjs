import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import {readFileSync} from 'node:fs';
import {schedulingEnvironment} from './helpers/scheduling-environment.mjs';
import {plain} from './helpers/admin-environment.mjs';
const script=readFileSync('Setting.html','utf8').match(/<script>([\s\S]*?)<\/script>/)[1];

// 実際の画面制御からサーバー処理へ接続し、外部サービスだけを代替する。
function environment(){
 const e=schedulingEnvironment(),r=e.read(),topics=[],topicCalls=[];
 Object.assign(r.input,{titlePattern:'{class_label} フォーム',materialTitlePattern:'{class_label} 資料'});r.label='1A';r.materialSettingsSaved=false;e.c.saveManagedRecord_(r);
 e.c.Classroom.Courses.Topics={list:()=>({topic:plain(topics)}),create:(body,courseId)=>{const topic={...plain(body),courseId,topicId:'topic-one'};topics.push(topic);topicCalls.push(topic);return plain(topic);}};
 const c=vm.createContext({console,structuredClone,Date:e.c.Date});vm.runInContext(script.replace(/boot\(\);\s*$/,''),c);
 c.document={getElementById:()=>null,querySelectorAll:()=>[]};c.savedRecord=plain(e.read());c.backend=async(method,...args)=>plain(e.c[method](...args));
 vm.runInContext(`state.panel='setup';state.step='publish';state.forms={classes:[{courseId:'100'}],defaults:{},records:[savedRecord]};state.batchMaterialDraft={materialTitlePattern:'{class_label} 新しい資料',description:'新しい本文',topicName:'振り返り',topicChange:true};state.scheduleDrafts={'batch:publish':'2026-10-05T10:00','batch:close':'2026-10-06T10:00'};let confirmations=[],messages=[];rpc=backend;requireSaved=()=>true;task=async work=>work();render=()=>{};reloadForms=async()=>{};status=message=>messages.push(message);confirmAction=async(title,summary)=>{confirmations.push(summary);return true;};`,c);
 return {...e,ui:c,topics,topicCalls};
}
test('aggregate registration persists individual dates without posting, opening forms or installing a timer',async()=>{
 const e=environment();await vm.runInContext("formCommand('register-all')",e.ui);
 assert.equal(e.read().input.description,'新しい本文');
 assert.equal(e.read().stage,'registered');assert.equal(e.posts.size,0);assert.equal(e.form.accepting,false);assert.equal(e.form.published,false);assert.equal(e.triggers.length,0);
 assert.deepEqual(plain(e.read().publishingSettings),{scheduledTime:'2026-10-05T01:00:00.000Z',closesAt:'2026-10-06T01:00:00.000Z'});
 assert.equal(vm.runInContext('confirmations.length',e.ui),1);
 e.ui.savedRecord=plain(e.read());vm.runInContext('state.forms.records=[savedRecord];state.scheduleDrafts={};',e.ui);
 assert.equal(vm.runInContext("scheduleDraftValue('publish:f1',publishingSettings(savedRecord).scheduledTime)",e.ui),'2026-10-05T10:00');
});

for(const action of ['schedule-all','set-close-all'])test(`${action} saves only its date and does not activate either reservation`,async()=>{
 const e=environment();await vm.runInContext(`formCommand('${action}')`,e.ui);
 assert.equal(e.posts.size,0);assert.equal(e.form.published,false);assert.equal(e.triggers.length,0);
 assert.deepEqual(plain(e.read().publishingSettings),action==='schedule-all'?{scheduledTime:'2026-10-05T01:00:00.000Z',closesAt:''}:{scheduledTime:'',closesAt:'2026-10-06T01:00:00.000Z'});
});

test('bulk posting uses saved dates rather than unsaved bulk inputs and activates the deadline',async()=>{
 const e=environment();await vm.runInContext("formCommand('register-all')",e.ui);
 assert.equal(e.posts.size,0);
 vm.runInContext("state.scheduleDrafts={'batch:publish':'2026-10-08T10:00','batch:close':'2026-10-09T10:00'};",e.ui);
 await vm.runInContext("formCommand('publish-all')",e.ui);
 assert.equal(e.posts.size,1);const post=[...e.posts.values()][0];assert.equal(post.state,'DRAFT');assert.equal(post.scheduledTime,'2026-10-05T01:00:00.000Z');
 assert.equal(e.read().closesAt,'2026-10-06T01:00:00.000Z');assert.equal(e.read().closeState,'scheduled');assert.equal(e.triggers.length,1);assert.equal(e.form.accepting,true);
 e.advance(6*86400000);e.fire();assert.equal(e.form.accepting,false);
});

test('registration of dates keeps an existing deadline and scheduled Classroom material unchanged',async()=>{
 const e=environment();e.c.saveManagedRecord_({...e.read(),materialSettingsSaved:true});e.run('schedule',{scheduledTime:'2026-10-03T01:00:00Z'});e.run('set-close',{closesAt:'2026-10-04T01:00:00Z'});
 const original=plain([...e.posts.values()][0]);e.ui.savedRecord=plain(e.read());vm.runInContext('state.forms.records=[savedRecord];',e.ui);
 await vm.runInContext("formCommand('set-close-all')",e.ui);await vm.runInContext("formCommand('schedule-all')",e.ui);
 assert.deepEqual(plain([...e.posts.values()][0]),original);assert.equal(e.read().closesAt,'2026-10-04T01:00:00.000Z');assert.equal(e.read().closeState,'scheduled');
 assert.deepEqual(plain(e.read().publishingSettings),{scheduledTime:'2026-10-05T01:00:00.000Z',closesAt:'2026-10-06T01:00:00.000Z'});
});

test('individual date overrides select immediate posting and defer the saved deadline until posting',async()=>{
 const e=environment();await vm.runInContext("formCommand('register-all')",e.ui);
 vm.runInContext("state.scheduleDrafts={'publish:f1':'','close:f1':'2026-10-07T12:00'};",e.ui);
 await vm.runInContext("formCommand('save-close-date','f1')",e.ui);
 assert.equal(e.triggers.length,0);assert.equal(vm.runInContext("state.scheduleDrafts['publish:f1']",e.ui),'');
 await vm.runInContext("formCommand('save-publish-date','f1')",e.ui);
 assert.equal(e.posts.size,0);assert.equal(e.form.published,false);
 await vm.runInContext("formCommand('publish','f1')",e.ui);
 assert.equal([...e.posts.values()][0].state,'PUBLISHED');assert.equal(e.read().closesAt,'2026-10-07T03:00:00.000Z');assert.equal(e.triggers.length,1);
});

test('expired saved publication and deadline are rejected before any posting or timer changes',async()=>{
 for(const elapsed of [5*86400000,6*86400000]){
  const e=environment();await vm.runInContext("formCommand('register-all')",e.ui);e.advance(elapsed);
  await vm.runInContext("formCommand('publish-all')",e.ui);
  assert.equal(e.posts.size,0);assert.equal(e.form.published,false);assert.equal(e.triggers.length,0);assert.equal(e.read().stage,'registered');
  assert.match(vm.runInContext('messages.at(-1)',e.ui),/未来/);
 }
});

test('invalid, conflicting or foreign date batches leave every managed row unchanged',()=>{
 for(const invalid of ['past','reversed','stale','foreign','duplicate','unprepared','wrong-key']){
  const e=environment(),r=e.read(),other={...plain(r),id:'f2'};if(invalid==='foreign')other.ownerScriptId='another';if(invalid==='unprepared')other.stage='creating';
  if(invalid==='foreign'){const saved=e.c.saveManagedRecord_({...other,ownerScriptId:r.ownerScriptId});const sheet=e.sheets.get('システム管理'),values=sheet.getRange(3,1,1,3).getDisplayValues();values[0][2]=JSON.stringify({...saved,ownerScriptId:'another'});sheet.getRange(3,1,1,3).setValues(values);}else e.c.saveManagedRecord_(other);
  const before=plain(e.c.getManagedRecords_()),targets=before.map(x=>({id:x.id,revision:x.revision}));
  if(invalid==='stale')targets[1].revision='stale';if(invalid==='duplicate')targets[1]=targets[0];
  const changes=invalid==='past'?{closesAt:'2026-09-30T01:00:00Z'}:invalid==='reversed'?{scheduledTime:'2026-10-05T01:00:00Z',closesAt:'2026-10-05T01:00:00Z'}:invalid==='wrong-key'?{stage:'published'}:{closesAt:'2026-10-06T01:00:00Z'};
  assert.throws(()=>e.c.saveManagedPublishingSettings(targets,changes));assert.deepEqual(plain(e.c.getManagedRecords_()),before);
  assert.equal(e.posts.size,0);assert.equal(e.form.published,false);assert.equal(e.triggers.length,0);
 }
});

test('date saves preserve unrelated rows and old active deadlines survive after pending settings are saved',()=>{
 const e=environment();e.c.saveManagedRecord_({...e.read(),materialSettingsSaved:true});e.run('publish');e.run('set-close',{closesAt:'2026-10-02T01:00:00Z'});
 const other=e.c.saveManagedRecord_({...plain(e.read()),id:'f2'}),r=e.read();
 e.c.saveManagedPublishingSettings([{id:r.id,revision:r.revision}],{closesAt:'2026-10-06T01:00:00Z'});
 assert.deepEqual(plain(e.c.loadManagedRecord_('f2')),plain(other));assert.equal(e.read().closesAt,'2026-10-02T01:00:00.000Z');
 e.advance(2*86400000);e.fire();assert.equal(e.form.accepting,false);
});

test('cancelling active reservations clears their saved dates so a later post cannot reuse them',()=>{
 const e=environment();e.c.saveManagedRecord_({...e.read(),materialSettingsSaved:true});let r=e.read();
 e.c.saveManagedPublishingSettings([{id:r.id,revision:r.revision}],{scheduledTime:'2026-10-05T01:00:00Z',closesAt:'2026-10-06T01:00:00Z'});
 e.run('schedule',{scheduledTime:'2026-10-05T01:00:00Z',closesAt:'2026-10-06T01:00:00Z'});
 e.run('cancel-schedule');assert.equal(e.read().publishingSettings.scheduledTime,'');
 e.run('cancel-close');assert.equal(e.read().publishingSettings.closesAt,'');assert.equal(e.read().closesAt,undefined);
});

test('a refreshed record cannot silently adopt the revision of a locally edited date',async()=>{
 const e=environment(),field={dataset:{scheduleKey:'close:f1',scheduleSaved:''},value:'2026-10-06T10:00'};
 e.ui.document={getElementById:()=>null,querySelectorAll:selector=>selector==='[data-schedule-key]'?[field]:[]};vm.runInContext('readScheduleDrafts()',e.ui);
 const r=e.c.saveManagedRecord_({...e.read(),publishingSettings:{scheduledTime:'',closesAt:'2026-10-07T01:00:00Z'}});e.ui.savedRecord=plain(r);vm.runInContext('acceptFormConsoleData({...state.forms,records:[savedRecord]});',e.ui);
 await assert.rejects(vm.runInContext("formCommand('save-close-date','f1')",e.ui),/更新|再確認/);
 assert.equal(e.read().publishingSettings.closesAt,'2026-10-07T01:00:00Z');assert.equal(vm.runInContext("state.scheduleDrafts['close:f1']",e.ui),'2026-10-06T10:00');
});
test('cancelled aggregate confirmation does not change backend records, topics, forms or posts',async()=>{
 const e=environment(),before=plain(e.read());vm.runInContext('confirmAction=async()=>false;',e.ui);await vm.runInContext("formCommand('register-all')",e.ui);
 assert.deepEqual(plain(e.read()),before);assert.equal(e.topicCalls.length,0);assert.equal(e.posts.size,0);assert.equal(e.form.published,false);
});
test('aggregate deadline extension works against real server validation of the previous deadline',async()=>{
 const e=environment(),r=e.read();r.closesAt='2026-10-03T01:00:00Z';e.c.saveManagedRecord_(r);e.ui.savedRecord=plain(e.read());vm.runInContext('state.forms.records=[savedRecord];',e.ui);
 await vm.runInContext("formCommand('register-all')",e.ui);
 assert.equal(e.read().publishingSettings.scheduledTime,'2026-10-05T01:00:00.000Z');assert.equal(e.read().publishingSettings.closesAt,'2026-10-06T01:00:00.000Z');assert.equal(e.read().closesAt,'2026-10-03T01:00:00Z');
});

for(const stage of ['publish','schedule'])for(const command of ['save-material-settings','apply-batch-material-settings','register-all']){
 test(`${command} preserves a legacy manual topic while editing title and description on ${stage}`,async()=>{
  const e=environment();e.c.saveManagedRecord_({...e.read(),materialSettingsSaved:true});e.run(stage,stage==='schedule'?{scheduledTime:'2026-10-05T01:00:00Z'}:{});
  const r=e.read(),post=e.posts.get(r.materialId);post.topicId='legacy-manual';const original=plain(post);
  e.ui.savedRecord=plain(r);vm.runInContext(`state.forms.records=[savedRecord];state.batchMaterialDraft=null;state.scheduleDrafts={};Object.assign(materialSettingsDraft(savedRecord),{materialTitlePattern:'{class_label} 編集',description:'編集本文'});Object.assign(batchMaterialSettingsDraft(),{materialTitlePattern:'{class_label} 編集',description:'編集本文'});`,e.ui);
  await vm.runInContext(`formCommand('${command}','f1')`,e.ui);
  assert.equal(post.topicId,'legacy-manual');assert.equal(post.title,'1A 編集');assert.equal(post.description,'編集本文');
  assert.equal(post.state,original.state);assert.equal(post.scheduledTime,original.scheduledTime);assert.equal(post.id,original.id);
  const patch=e.calls.filter(c=>c[0]==='patch').at(-1);assert.equal(patch[4].updateMask.includes('topicId'),false);
  assert.match(vm.runInContext('confirmations[0]',e.ui),/トピック：変更しない/);
 });
}
test('explicit assignment then explicit blank clears only the topic and later edits preserve external reassignment',async()=>{
 const e=environment();e.c.saveManagedRecord_({...e.read(),materialSettingsSaved:true});e.run('publish');const id=e.read().materialId;
 e.ui.savedRecord=plain(e.read());vm.runInContext(`state.forms.records=[savedRecord];state.scheduleDrafts={};Object.assign(materialSettingsDraft(savedRecord),{topicChange:true,topicName:'振り返り'});`,e.ui);
 await vm.runInContext("formCommand('save-material-settings','f1')",e.ui);
 assert.equal(e.posts.get(id).topicId,'topic-one');
 vm.runInContext(`Object.assign(materialSettingsDraft(state.forms.records[0]),{topicChange:true,topicName:''});`,e.ui);
 await vm.runInContext("formCommand('save-material-settings','f1')",e.ui);
 assert.equal(e.posts.get(id).topicId,undefined);assert.equal(e.posts.get(id).state,'PUBLISHED');
 e.posts.get(id).topicId='external-after-clear';
 vm.runInContext(`materialSettingsDraft(state.forms.records[0]).description='本文のみ変更';`,e.ui);
 await vm.runInContext("formCommand('save-material-settings','f1')",e.ui);
 assert.equal(e.posts.get(id).topicId,'external-after-clear');
});
test('bulk preserve keeps different manual topics on multiple existing materials',async()=>{
 const e=environment();e.c.saveManagedRecord_({...e.read(),materialSettingsSaved:true});e.run('publish');const first=e.read();e.posts.get(first.materialId).topicId='manual-a';
 const second=plain(first);Object.assign(second,{id:'f2',materialId:'manual-material-b'});e.c.saveManagedRecord_(second);e.posts.set(second.materialId,{...plain(e.posts.get(first.materialId)),id:second.materialId,topicId:'manual-b'});
 e.ui.savedRecords=[plain(e.read()),plain(e.c.loadManagedRecord_('f2'))];
 vm.runInContext(`state.forms.records=savedRecords;state.batchMaterialDraft=null;state.scheduleDrafts={};batchMaterialSettingsDraft().description='共通本文だけ';`,e.ui);
 await vm.runInContext("formCommand('register-all')",e.ui);
 assert.equal(e.posts.get(first.materialId).topicId,'manual-a');assert.equal(e.posts.get(second.materialId).topicId,'manual-b');
 assert.equal(e.posts.get(first.materialId).description,'共通本文だけ');assert.equal(e.posts.get(second.materialId).description,'共通本文だけ');
});

for(const stage of ['publish','schedule'])test(`bulk registration restores the requested description when the ${stage} material differs from its saved record`,async()=>{
 const e=environment();e.c.saveManagedRecord_({...e.read(),materialSettingsSaved:true});
 e.run(stage,stage==='schedule'?{scheduledTime:'2026-10-05T01:00:00Z'}:{});
 const r=e.read(),post=e.posts.get(r.materialId);post.description='';post.topicId='manual-topic';
 e.ui.savedRecord=plain(r);
 vm.runInContext(`state.forms.records=[savedRecord];state.batchMaterialDraft=null;state.scheduleDrafts={};batchMaterialSettingsDraft().materialTitlePattern='{class_label} 編集';`,e.ui);
 await vm.runInContext("formCommand('register-all')",e.ui);
 assert.equal(post.title,'1A 編集');assert.equal(post.description,r.input.description);assert.equal(post.topicId,'manual-topic');
 assert.equal(post.id,r.materialId);assert.equal(post.state,stage==='publish'?'PUBLISHED':'DRAFT');
});

test('bulk registration reads title, multiline description and explicit topic selection from the input fields',async()=>{
 const e=environment();vm.runInContext('state.batchMaterialDraft=null;state.scheduleDrafts={};',e.ui);
 const fields=[{dataset:{batchMaterialField:'materialTitlePattern'},value:'{class_label} 画面の資料'},
  {dataset:{batchMaterialField:'description'},value:'画面の本文\n二行目'},
  {dataset:{batchMaterialField:'topicChange'},type:'checkbox',checked:true},
  {dataset:{batchMaterialField:'topicName'},value:'振り返り'}];
 e.ui.document={getElementById:()=>null,querySelectorAll:selector=>selector==='[data-batch-material-field]'?fields:[]};
 await vm.runInContext("formCommand('register-all')",e.ui);
 assert.equal(e.read().materialTitle,'1A 画面の資料');assert.equal(e.read().input.description,'画面の本文\n二行目');assert.equal(e.read().topicId,'topic-one');
});

for(const description of ['回答してください',''])test(`unchanged saved text is reconciled against remote text for ${description?'nonempty':'empty'} descriptions`,async()=>{
 const e=environment();const initial=e.read();initial.input.description=description;initial.materialSettingsSaved=true;e.c.saveManagedRecord_(initial);e.run('publish');
 const r=e.read(),post=e.posts.get(r.materialId);post.description='Classroom 側の異なる本文';post.topicId='manual';
 r.materialSettingsBaseline={materialTitlePattern:r.input.materialTitlePattern,description,topicName:''};e.c.saveManagedRecord_(r);
 e.ui.savedRecord=plain(e.read());vm.runInContext('state.forms.records=[savedRecord];state.batchMaterialDraft=null;state.scheduleDrafts={};',e.ui);
 await vm.runInContext("formCommand('register-all')",e.ui);
 assert.equal(post.description,description);assert.equal(post.topicId,'manual');assert.match(vm.runInContext('confirmations[0]',e.ui),/Classroom の資料本文/);
 const patches=e.calls.filter(c=>c[0]==='patch').length;
 await vm.runInContext("formCommand('register-all')",e.ui);
 assert.equal(e.calls.filter(c=>c[0]==='patch').length,patches);
});

test('remote description changes after confirmation reject registration even when saved input is unchanged',async()=>{
 const e=environment();e.c.saveManagedRecord_({...e.read(),materialSettingsSaved:true});e.run('publish');
 const r=e.read(),post=e.posts.get(r.materialId);post.description='確認前の本文';e.ui.savedRecord=plain(r);
 vm.runInContext('state.forms.records=[savedRecord];state.batchMaterialDraft=null;state.scheduleDrafts={};',e.ui);
 e.ui.changeRemote=()=>{post.description='確認後の本文';};vm.runInContext('confirmAction=async()=>{changeRemote();return true;};',e.ui);
 await vm.runInContext("formCommand('register-all')",e.ui);
 assert.equal(post.description,'確認後の本文');assert.equal(e.calls.filter(c=>c[0]==='patch').length,0);
 assert.match(vm.runInContext('messages.at(-1)',e.ui),/再確認/);
});

test('bulk registration updates title, saved description and an explicitly selected topic on an existing material',async()=>{
 const e=environment();e.c.saveManagedRecord_({...e.read(),materialSettingsSaved:true});e.run('publish');
 const r=e.read(),post=e.posts.get(r.materialId);post.description='';e.ui.savedRecord=plain(r);
 vm.runInContext(`state.forms.records=[savedRecord];state.batchMaterialDraft={materialTitlePattern:'{class_label} まとめて更新',description:savedRecord.input.description,topicName:'振り返り',topicChange:true};state.scheduleDrafts={};`,e.ui);
 await vm.runInContext("formCommand('register-all')",e.ui);
 assert.equal(post.title,'1A まとめて更新');assert.equal(post.description,r.input.description);assert.equal(post.topicId,'topic-one');
 assert.equal(e.read().input.topicName,'振り返り');assert.equal(post.id,r.materialId);assert.equal(e.posts.size,1);
});

test('bulk registration updates title, saved description and an explicitly selected topic on an existing material',async()=>{
 const e=environment();e.c.saveManagedRecord_({...e.read(),materialSettingsSaved:true});e.run('publish');
 const r=e.read(),post=e.posts.get(r.materialId);post.description='';e.ui.savedRecord=plain(r);
 vm.runInContext(`state.forms.records=[savedRecord];state.batchMaterialDraft={materialTitlePattern:'{class_label} まとめて更新',description:savedRecord.input.description,topicName:'振り返り',topicChange:true};state.scheduleDrafts={};`,e.ui);
 await vm.runInContext("formCommand('register-all')",e.ui);
 assert.equal(post.title,'1A まとめて更新');assert.equal(post.description,r.input.description);assert.equal(post.topicId,'topic-one');
 assert.equal(e.read().input.topicName,'振り返り');assert.equal(post.id,r.materialId);assert.equal(e.posts.size,1);
});
