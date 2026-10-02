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
test('one aggregate confirmation creates topic, saves settings, schedules material and closes form later',async()=>{
 const e=environment();await vm.runInContext("formCommand('register-all')",e.ui);
 assert.equal(e.topicCalls.length,1);assert.equal(e.read().topicId,'topic-one');assert.equal(e.read().input.description,'新しい本文');assert.equal(e.read().scheduledTime,'2026-10-05T01:00:00.000Z');assert.equal(e.read().closesAt,'2026-10-06T01:00:00.000Z');assert.equal([...e.posts.values()][0].topicId,'topic-one');assert.equal(e.form.accepting,true);assert.equal(vm.runInContext('confirmations.length',e.ui),1);assert.match(vm.runInContext('confirmations[0]',e.ui),/メールアドレスを収集する/);assert.match(vm.runInContext('confirmations[0]',e.ui),/同名トピック.*[\s\S]*作成/);
});
test('cancelled aggregate confirmation does not change backend records, topics, forms or posts',async()=>{
 const e=environment(),before=plain(e.read());vm.runInContext('confirmAction=async()=>false;',e.ui);await vm.runInContext("formCommand('register-all')",e.ui);
 assert.deepEqual(plain(e.read()),before);assert.equal(e.topicCalls.length,0);assert.equal(e.posts.size,0);assert.equal(e.form.published,false);
});
test('aggregate deadline extension works against real server validation of the previous deadline',async()=>{
 const e=environment(),r=e.read();r.closesAt='2026-10-03T01:00:00Z';e.c.saveManagedRecord_(r);e.ui.savedRecord=plain(e.read());vm.runInContext('state.forms.records=[savedRecord];',e.ui);
 await vm.runInContext("formCommand('register-all')",e.ui);
 assert.equal(e.read().scheduledTime,'2026-10-05T01:00:00.000Z');assert.equal(e.read().closesAt,'2026-10-06T01:00:00.000Z');
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
