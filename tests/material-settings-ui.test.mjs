import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import {readFileSync} from 'node:fs';
const script=readFileSync('Setting.html','utf8').match(/<script>([\s\S]*?)<\/script>/)[1];
function environment(saved=false){
 const c=vm.createContext({console,structuredClone});vm.runInContext(script.replace(/boot\(\);\s*$/,''),c);
 c.document={getElementById:()=>null,querySelectorAll:()=>[]};
 vm.runInContext(`state.panel='setup';state.step='publish';state.forms={classes:[],defaults:{},records:[{id:'a',kind:'form',stage:'registered',revision:'r1',label:'1A',courseId:'100',materialTitle:'1A 資料',materialSettingsSaved:${saved},input:{titlePattern:'{class_label} フォーム',materialTitlePattern:'{class_label} 資料',description:'保存済み本文',policy:{domains:['example.com'],emails:[]}}}]};
 let calls=[],messages=[];task=async fn=>fn();requireSaved=()=>true;confirmAction=async()=>true;render=()=>{};reloadForms=async()=>{};status=message=>messages.push(message);
 rpc=async(method,...args)=>{calls.push([method,...args]);if(method==='previewManagedFormSettings')return {targets:args[0].targets.map(t=>({...t,label:t.id,changed:true,changes:[],fingerprint:'fp'}))};if(method==='runManagedFormSettings'){const r=state.forms.records.find(r=>r.id===args[0]);return {...r,materialSettingsSaved:true,...(args[4]?{materialSettingsBaseline:copy(args[1])}:{}),revision:r.id+'2',input:{...r.input,...args[1]}};}return {fingerprint:'fp'};};`,c);return c;
}
test('publishing no longer offers or dispatches name enrichment',async()=>{
 const c=environment(true);vm.runInContext('state.forms.records[0].responseSheetId=123;',c);
 assert.doesNotMatch(vm.runInContext("renderForms('publish')",c),/data-form-command="names"|名前を補完/);
 await vm.runInContext("formCommand('names','a')",c);
 assert.equal(vm.runInContext('calls.length',c),0);
});

test('posting fields live in the publishing step and saving unchanged defaults unlocks the post',async()=>{
 const c=environment();let html=vm.runInContext("renderForms('prepare')",c);assert.doesNotMatch(html,/id="form-materialTitlePattern"|id="form-description"/);
 html=vm.runInContext("renderForms('publish')",c);assert.match(html,/保存済み本文/);assert.match(html,/data-form-command="save-material-settings"/);assert.match(html,/<option value="publish" disabled>/);
 await vm.runInContext("formCommand('save-material-settings','a')",c);
 assert.equal(vm.runInContext('state.forms.records[0].materialSettingsSaved',c),true);
 html=vm.runInContext("renderForms('publish')",c);assert.match(html,/<option value="publish">/);
});
test('unsaved edits block immediate, scheduled and batch publishing but preserve deadline operations',async()=>{
 const c=environment(true);vm.runInContext("materialSettingsDraft(state.forms.records[0]).description='入力中';state.scheduleDrafts={'batch:publish':'2027-10-02T10:00'};state.forms.records[0].publishingSettings={scheduledTime:'',closesAt:'2027-10-02T03:00:00Z'};",c);
 for(const action of ['publish','schedule','publish-all'])await vm.runInContext(`formCommand('${action}','a')`,c);
 assert.equal(vm.runInContext('calls.length',c),0);assert.match(vm.runInContext('messages.at(-1)',c),/保存/);
 await vm.runInContext("formCommand('set-close','a')",c);assert.equal(vm.runInContext("calls.some(c=>c[0]==='runManagedFormAction')",c),true);
});
test('input survives refresh and save failures; stale drafts do not adopt a newer revision',async()=>{
 const c=environment(true);vm.runInContext("materialSettingsDraft(state.forms.records[0]).description='入力中';acceptFormConsoleData({...state.forms,records:[{...state.forms.records[0],revision:'external'}]});rpc=async()=>{throw Error('競合');};",c);
 assert.equal(vm.runInContext('materialSettingsDraft(state.forms.records[0]).revision',c),'r1');
 await assert.rejects(vm.runInContext("formCommand('save-material-settings','a')",c),/競合/);
 assert.match(vm.runInContext("renderForms('publish')",c),/入力中/);assert.equal(vm.runInContext('materialSettingsReady(state.forms.records[0])',c),false);
});

test('refresh shows external saved changes when the local inputs were not edited',()=>{
 const c=environment(true);
 const el={dataset:{record:'a',materialField:'description'},value:'保存済み本文'};
 c.document={getElementById:()=>null,querySelectorAll:selector=>selector==='[data-material-field]'?[el]:[]};
 vm.runInContext('materialSettingsDraft(state.forms.records[0])',c);
 vm.runInContext("acceptFormConsoleData({...state.forms,records:[{...state.forms.records[0],revision:'r2',input:{...state.forms.records[0].input,description:'外部保存'}}]});readMaterialSettingsDrafts();",c);
 assert.equal(vm.runInContext('materialSettingsDraft(state.forms.records[0]).description',c),'外部保存');
 assert.equal(vm.runInContext('materialSettingsReady(state.forms.records[0])',c),true);
});

test('typing disables posting buttons immediately while leaving scheduling cancellation usable',()=>{
 const c=environment(true);vm.runInContext("materialSettingsDraft(state.forms.records[0])",c);
 const input={dataset:{record:'a',materialField:'description'},value:'編集中'},post={dataset:{materialRequired:'a'},disabled:false},batch={dataset:{materialRequired:'batch'},disabled:false},info={dataset:{materialSaveStatus:'a'},textContent:''};
 c.document={getElementById:()=>null,querySelectorAll:selector=>selector==='[data-material-field]'?[input]:selector==='[data-material-required]'?[post,batch]:selector==='[data-material-save-status]'?[info]:[]};
 vm.runInContext('readMaterialSettingsDrafts();updateMaterialSettingsControls();',c);
 assert.equal(post.disabled,true);assert.equal(batch.disabled,true);assert.match(info.textContent,/未保存/);
 input.value='保存済み本文';vm.runInContext('readMaterialSettingsDrafts();updateMaterialSettingsControls();',c);assert.equal(post.disabled,false);
});

test('a failed save retains edited content and cannot unlock publishing with older saved values',async()=>{
 const c=environment(true);vm.runInContext("materialSettingsDraft(state.forms.records[0]).description='未保存の本文';const previewRpc=rpc;rpc=async(method,...args)=>{if(method==='runManagedFormSettings')throw Error('保存失敗');return previewRpc(method,...args);};",c);
 await assert.rejects(vm.runInContext("formCommand('save-material-settings','a')",c),/保存失敗/);
 assert.equal(vm.runInContext('materialSettingsDraft(state.forms.records[0]).description',c),'未保存の本文');
 assert.equal(vm.runInContext('materialSettingsReady(state.forms.records[0])',c),false);
 await vm.runInContext("formCommand('publish','a')",c);
 assert.equal(vm.runInContext("calls.some(c=>c[0]==='previewManagedFormAction')",c),false);
});

test('reloading saved settings requires accepting draft discard and resolves a stale draft explicitly',async()=>{
 const c=environment(true);vm.runInContext("materialSettingsDraft(state.forms.records[0]).description='入力中';state.forms.records[0].revision='r2';state.forms.records[0].input.description='新しい保存値';confirmAction=async()=>false;",c);
 await vm.runInContext("formCommand('reload-material-settings','a')",c);
 assert.equal(vm.runInContext('materialSettingsDraft(state.forms.records[0]).description',c),'入力中');
 vm.runInContext('confirmAction=async()=>true;',c);await vm.runInContext("formCommand('reload-material-settings','a')",c);
 assert.equal(vm.runInContext('materialSettingsDraft(state.forms.records[0]).description',c),'新しい保存値');
 assert.equal(vm.runInContext('materialSettingsDraft(state.forms.records[0]).revision',c),'r2');
});

function batchEnvironment(){
 const c=environment(true);
 vm.runInContext(`state.forms.records.push(
  {...copy(state.forms.records[0]),id:'b',label:'1B',revision:'b1',stage:'scheduled',materialId:'post-b',scheduledTime:'2026-10-03T01:00:00Z'},
  {...copy(state.forms.records[0]),id:'updating',formUpdate:{}},
  {...copy(state.forms.records[0]),id:'review',stage:'publish_review'},
  {id:'return',kind:'return',stage:'published'}
 );`,c);
 return c;
}

test('batch posting fields render above class settings and preserve escaped input across refreshes',()=>{
 const c=batchEnvironment();let html=vm.runInContext("renderForms('publish')",c);
 assert.match(html,/data-batch-material-field="materialTitlePattern"/);
 assert.match(html,/data-batch-material-field="description"/);
 assert.ok(html.indexOf('data-batch-material-field')<html.indexOf('id="material-a-'));
 const inputs=[{dataset:{batchMaterialField:'materialTitlePattern'},value:'{class_label} <資料>'},{dataset:{batchMaterialField:'description'},value:'一括入力\n<&>'}];
 c.document={getElementById:()=>null,querySelectorAll:selector=>selector==='[data-batch-material-field]'?inputs:[]};
 vm.runInContext('readCurrent();acceptFormConsoleData(copy(state.forms));',c);
 html=vm.runInContext("renderForms('publish')",c);
 assert.match(html,/\{class_label\} &lt;資料&gt;/);assert.match(html,/一括入力\n&lt;&amp;&gt;/);
 assert.equal(vm.runInContext('materialSettingsReady(state.forms.records[0])',c),true);
 assert.equal(vm.runInContext('calls.length',c),0);
});

test('applying common text saves only editable classes including scheduled materials',async()=>{
 const c=batchEnvironment();vm.runInContext(`renderForms('publish');state.batchMaterialDraft={materialTitlePattern:'{class_label} 共通資料',description:''};
 materialSettingsDraft(state.forms.records[1]).description='編集中';`,c);
 const before=vm.runInContext('JSON.stringify(state.forms.records.slice(2))',c);
 await vm.runInContext("formCommand('apply-batch-material-settings')",c);
 for(const id of ['a','b']){
  assert.equal(vm.runInContext(`state.forms.records.find(r=>r.id==='${id}').input.materialTitlePattern`,c),'{class_label} 共通資料');
  assert.equal(vm.runInContext(`state.forms.records.find(r=>r.id==='${id}').input.description`,c),'');
  assert.equal(vm.runInContext(`materialSettingsReady(state.forms.records.find(r=>r.id==='${id}'))`,c),true);
 }
 assert.equal(vm.runInContext('state.forms.records[1].scheduledTime',c),'2026-10-03T01:00:00Z');
 assert.equal(vm.runInContext('state.forms.records[1].materialId',c),'post-b');
 assert.equal(vm.runInContext('state.materialDrafts.updating',c),undefined);
 assert.equal(vm.runInContext('state.materialDrafts.review',c),undefined);
 assert.equal(vm.runInContext('JSON.stringify(state.forms.records.slice(2))',c),before);
 assert.equal(vm.runInContext("calls.filter(c=>c[0]==='previewManagedFormSettings').length",c),1);
 assert.equal(vm.runInContext("calls.filter(c=>c[0]==='runManagedFormSettings').length",c),2);
 assert.doesNotMatch(vm.runInContext("renderForms('publish')",c),/data-form-command="schedule-all"[^>]* disabled/);
});

test('cancelling batch replacement keeps individually edited text intact',async()=>{
 const c=batchEnvironment();vm.runInContext(`renderForms('publish');materialSettingsDraft(state.forms.records[1]).description='個別入力';
 state.batchMaterialDraft={materialTitlePattern:'一括タイトル',description:'一括本文'};let confirmations=0;confirmAction=async()=>{confirmations++;return false;};`,c);
 await vm.runInContext("formCommand('apply-batch-material-settings')",c);
 assert.equal(vm.runInContext('state.materialDrafts.b.description',c),'個別入力');
 assert.equal(vm.runInContext('materialSettingsReady(state.forms.records[0])',c),true);
 assert.equal(vm.runInContext("calls.filter(c=>c[0]==='runManagedFormSettings').length",c),0);
 assert.equal(vm.runInContext('confirmations',c),1);
});

test('a refresh while confirming batch text aborts replacement when target state or revision changes',async()=>{
 for(const change of ["formUpdate:{}","revision:'external',input:{...state.forms.records[0].input,description:'外部保存'}"]){
  const c=batchEnvironment();vm.runInContext(`renderForms('publish');state.batchMaterialDraft={materialTitlePattern:'一括タイトル',description:'一括本文'};
   confirmAction=async()=>{acceptFormConsoleData({...state.forms,records:state.forms.records.map(r=>r.id==='a'?{...r,${change}}:r)});return true;};`,c);
  await vm.runInContext("formCommand('apply-batch-material-settings')",c);
  assert.equal(vm.runInContext('materialSettingsDraft(state.forms.records[1]).description',c),'保存済み本文');
  assert.equal(vm.runInContext("state.materialDrafts.a?.description==='一括本文'",c),false);
  assert.match(vm.runInContext('messages.at(-1)',c),/更新|再確認/);
  assert.equal(vm.runInContext("calls.filter(c=>c[0]==='runManagedFormSettings').length",c),0);
 }
});

test('batch replacement rejects a blank title and safely handles no editable forms',async()=>{
 const c=batchEnvironment();vm.runInContext("renderForms('publish');state.batchMaterialDraft={materialTitlePattern:'  ',description:'共通'};",c);
 await vm.runInContext("formCommand('apply-batch-material-settings')",c);
 assert.equal(vm.runInContext('materialSettingsReady(state.forms.records[0])',c),true);
 assert.match(vm.runInContext('messages.at(-1)',c),/タイトル/);
 vm.runInContext('state.forms.records=[];',c);
 await vm.runInContext("formCommand('apply-batch-material-settings')",c);
 assert.match(vm.runInContext('messages.at(-1)',c),/対象/);
});

test('reading the first class uses current inputs without changing other class drafts or schedule dates',async()=>{
 const c=batchEnvironment();vm.runInContext("renderForms('publish');state.scheduleDrafts={'batch:publish':'2026-10-03T10:00'};",c);
 const html=vm.runInContext("renderForms('publish')",c);
 assert.ok(html.indexOf('id="material-b-')<html.indexOf('id="material-a-'));
 const fields=[{dataset:{record:'b',materialField:'materialTitlePattern'},value:'{class_label} 先頭の資料'},{dataset:{record:'b',materialField:'description'},value:'先頭で入力中'}];
 c.document={getElementById:()=>null,querySelectorAll:selector=>selector==='[data-material-field]'?fields:[]};
 await vm.runInContext("formCommand('load-first-material-settings')",c);
 assert.equal(vm.runInContext('state.batchMaterialDraft.materialTitlePattern',c),'{class_label} 先頭の資料');
 assert.equal(vm.runInContext('state.batchMaterialDraft.description',c),'先頭で入力中');
 assert.equal(vm.runInContext('state.materialDrafts.a.description',c),'保存済み本文');
 assert.equal(vm.runInContext("state.scheduleDrafts['batch:publish']",c),'2026-10-03T10:00');
 assert.equal(vm.runInContext('calls.length',c),0);
});
