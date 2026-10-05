import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import {readFileSync} from 'node:fs';
const script=readFileSync('Setting.html','utf8').match(/<script>([\s\S]*?)<\/script>/)[1];
function model(){
 const c=vm.createContext({console,structuredClone});vm.runInContext(script.replace(/boot\(\);\s*$/,''),c);
 c.document={getElementById:()=>null,querySelectorAll:()=>[]};
 vm.runInContext(`state.panel='setup';state.step='publish';state.forms={classes:[],defaults:{},records:['a','b'].map(id=>({id,kind:'form',stage:'registered',label:id,revision:id+'1',materialSettingsSaved:true,input:{materialTitlePattern:'{class_label} 資料',description:'元の本文',policy:{domains:[],emails:[]}}}))};
 let calls=[],messages=[];task=async fn=>fn();requireSaved=()=>true;confirmAction=async()=>true;render=()=>{};reloadForms=async()=>{};status=m=>messages.push(m);
 rpc=async(method,...args)=>{calls.push([method,...args]);if(method==='previewManagedFormSettings')return {targets:args[0].targets.map(t=>({...t,fingerprint:t.id+'fp'}))};if(method==='runManagedFormSettings'){const r=state.forms.records.find(r=>r.id===args[0]);return {...r,revision:r.id+'2',materialSettingsSaved:true,...(args[4]?{materialSettingsBaseline:copy(args[1])}:{}),input:{...r.input,...args[1]}};}if(method==='previewManagedFormAction')return {fingerprint:'fp',title:args[0]};return {...state.forms.records.find(r=>r.id===args[0]),stage:'published'};};`,c);
 return c;
}
test('publishing groups common settings above two individual columns with compact controls',()=>{
 const c=model(),html=vm.runInContext("renderForms('publish')",c);
 assert.match(html,/<details[^>]*id="batchSchedules"[^>]* open/);
 for(const label of ['一括設定','クラスルーム投稿の一括登録','公開日時/受付終了日時の一括登録','各クラスに反映','一括投稿（2クラス）','個別登録','個別操作'])assert.ok(html.includes(label),label);
 assert.doesNotMatch(html,/一番上のクラスから読み込む|各クラスの入力欄に反映|data-form-command="reload-material-settings"/);
 assert.match(html,/data-help-body="[^"]*フォームを準備[^"]*読み方/);
 assert.match(html,/<summary>投稿タイトル・投稿文/);assert.match(html,/class="schedule-row"/);
 assert.match(html,/<select[^>]*data-form-operation="a"/);assert.match(html,/data-form-command="execute-operation" data-record="a"/);
 assert.ok(html.indexOf('data-form-command="publish-all"')<html.indexOf('配布するフォーム'));
 assert.doesNotMatch(html,/<details[^>]*id="materialSettings-a"[^>]* open/);
});
test('registration buttons use save before posting and apply after posting or scheduling',()=>{
 const c=model();
 for(const stage of ['registered','draft','deleted','scheduled','published']){
  vm.runInContext(`state.forms.records[0].stage='${stage}';`,c);
  const html=vm.runInContext('renderMaterialSettings(state.forms.records[0])+renderScheduleControls(state.forms.records[0])',c);
  const label=['scheduled','published'].includes(stage)?'反映':'保存';
  assert.match(html,new RegExp('data-form-command="save-material-settings"[^>]*>'+label+'<'));
  assert.match(html,/data-form-command="save-close-date"[^>]*>保存</);
 }
});
test('common text is previewed once and saved to all editable forms',async()=>{
 const c=model();vm.runInContext("state.batchMaterialDraft={materialTitlePattern:'{class_label} 共通',description:''};state.forms.records.push({...copy(state.forms.records[0]),id:'blocked',formUpdate:{}});",c);
 await vm.runInContext("formCommand('apply-batch-material-settings')",c);
 assert.equal(vm.runInContext("calls.filter(c=>c[0]==='previewManagedFormSettings').length",c),1);
 assert.equal(vm.runInContext("calls.filter(c=>c[0]==='runManagedFormSettings').length",c),2);
 assert.equal(vm.runInContext("state.forms.records.slice(0,2).every(r=>r.input.description===''&&materialSettingsReady(r))",c),true);
 assert.equal(vm.runInContext('state.forms.records[2].input.description',c),'元の本文');
});
test('partial common save retains failed inputs and retry skips successfully saved forms',async()=>{
 const c=model();vm.runInContext("state.batchMaterialDraft={materialTitlePattern:'共通',description:'新本文'};const normalRpc=rpc;let fail=true;rpc=async(method,...args)=>{if(method==='runManagedFormSettings'&&args[0]==='b'&&fail)throw Error('保存失敗');return normalRpc(method,...args);};",c);
 await vm.runInContext("formCommand('apply-batch-material-settings')",c);
 assert.equal(vm.runInContext('materialSettingsReady(state.forms.records[0])',c),true);
 assert.equal(vm.runInContext('materialSettingsReady(state.forms.records[1])',c),false);
 assert.equal(vm.runInContext('state.materialDrafts.b.description',c),'新本文');assert.match(vm.runInContext('messages.at(-1)',c),/保存失敗/);
 vm.runInContext('calls=[];fail=false;',c);await vm.runInContext("formCommand('apply-batch-material-settings')",c);
 assert.equal(vm.runInContext("calls.filter(c=>c[0]==='runManagedFormSettings').map(c=>c[1]).join(',')",c),'b');
});
test('bulk post targets saved unposted forms and retry skips successes',async()=>{
 const c=model();vm.runInContext("state.forms.records.push({...copy(state.forms.records[0]),id:'scheduled',stage:'scheduled'},{...copy(state.forms.records[0]),id:'deleted',stage:'deleted'},{...copy(state.forms.records[0]),id:'blocked',formUpdate:{}});const normalRpc=rpc;rpc=async(method,...args)=>{if(method==='runManagedFormAction'&&args[0]==='b'){calls.push([method,...args]);throw Error('投稿失敗');}return normalRpc(method,...args);};",c);
 await vm.runInContext("formCommand('publish-all')",c);
 assert.equal(vm.runInContext("calls.filter(c=>c[0]==='runManagedFormAction').map(c=>c[1]).join(',')",c),'a,b');
 assert.equal(vm.runInContext('state.forms.records[0].stage',c),'published');assert.match(vm.runInContext('messages.at(-1)',c),/投稿失敗/);
 vm.runInContext('calls=[];',c);await vm.runInContext("formCommand('publish-all')",c);
 assert.equal(vm.runInContext("calls.filter(c=>c[0]==='runManagedFormAction').map(c=>c[1]).join(',')",c),'b');
});
test('individual execution validates available actions and unsaved text',async()=>{
 const c=model();c.document.getElementById=id=>id==='form-operation-a'?{value:'publish'}:null;
 vm.runInContext("materialSettingsDraft(state.forms.records[0]).description='編集中';",c);
 await vm.runInContext("formCommand('execute-operation','a')",c);assert.equal(vm.runInContext('calls.length',c),0);
 vm.runInContext("delete state.materialDrafts.a;state.forms.records[0].stage='scheduled';",c);
 await vm.runInContext("formCommand('execute-operation','a')",c);assert.equal(vm.runInContext('calls.length',c),0);
 vm.runInContext("state.forms.records[0].stage='registered';",c);
 await vm.runInContext("formCommand('execute-operation','a')",c);assert.equal(vm.runInContext("calls.some(c=>c[0]==='runManagedFormAction')",c),true);
});
test('saved individual text is marked without comparing against unregistered batch edits',async()=>{
 const c=model();vm.runInContext("renderForms('publish');materialSettingsDraft(state.forms.records[0]).description='個別の本文';",c);
 await vm.runInContext("formCommand('save-material-settings','a')",c);
 assert.match(vm.runInContext('renderMaterialSettings(state.forms.records[0])',c),/data-individual-material="a">個別設定あり/);
 assert.match(vm.runInContext('renderMaterialSettings(state.forms.records[1])',c),/data-individual-material="b" hidden/);
 vm.runInContext("state.batchMaterialDraft.description='一括欄で編集中';",c);
 assert.match(vm.runInContext('renderMaterialSettings(state.forms.records[1])',c),/data-individual-material="b" hidden/);
});
test('registration details keep both opened and closed states after a save',async()=>{
 const c=model(),details=[{id:'batchSchedules',open:false},{id:'materialSettings-a',open:true},{id:'formClose-a',open:true}];
 c.document.querySelectorAll=selector=>selector==='[data-publish-details]'?details:[];
 await vm.runInContext("formCommand('save-material-settings','a')",c);
 const html=vm.runInContext("renderForms('publish')",c);
 assert.doesNotMatch(html,/<details[^>]*id="batchSchedules"[^>]* open/);
 assert.match(html,/<details[^>]*id="materialSettings-a"[^>]* open/);
 assert.match(html,/<details[^>]*id="formClose-a"[^>]* open/);
});
test('reopening the screen restores common text rather than the latest individually edited class',()=>{
 const c=model();vm.runInContext("state.forms.records.forEach(r=>{r.materialSettingsBaseline={materialTitlePattern:'{class_label} 資料',description:'元の本文'};});state.forms.records[1].input.description='最新クラスの個別本文';",c);
 vm.runInContext("renderForms('publish')",c);
 assert.equal(vm.runInContext('state.batchMaterialDraft.description',c),'元の本文');
 assert.match(vm.runInContext('renderMaterialSettings(state.forms.records[1])',c),/data-individual-material="b">個別設定あり/);
 assert.match(vm.runInContext('renderMaterialSettings(state.forms.records[0])',c),/data-individual-material="a" hidden/);
});
test('common text preview batches at most 100 classes per request',async()=>{
 const c=model();vm.runInContext("state.forms.records=Array.from({length:101},(_,i)=>({...copy(state.forms.records[0]),id:'class-'+i,revision:'r1'}));state.batchMaterialDraft={materialTitlePattern:'{class_label} 共通',description:'共通本文'};const normalRpc=rpc;rpc=async(method,...args)=>{if(method==='previewManagedFormSettings'&&args[0].targets.length>100)throw Error('100件上限');return normalRpc(method,...args);};",c);
 await vm.runInContext("formCommand('apply-batch-material-settings')",c);
 assert.equal(vm.runInContext("calls.filter(c=>c[0]==='previewManagedFormSettings').map(c=>c[1].targets.length).join(',')",c),'100,1');
 assert.equal(vm.runInContext("calls.filter(c=>c[0]==='runManagedFormSettings').length",c),101);
});
test('a newly prepared form does not replace the persisted common text during initial display',()=>{
 const c=model();vm.runInContext("state.forms.records[0].materialSettingsBaseline={materialTitlePattern:'{class_label} 共通資料',description:'登録済み共通本文'};state.forms.records[1].materialSettingsSaved=false;",c);
 assert.equal(vm.runInContext('batchMaterialSettingsDraft().description',c),'登録済み共通本文');
});
