import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import {readFileSync} from 'node:fs';
const script=readFileSync('Setting.html','utf8').match(/<script>([\s\S]*?)<\/script>/)[1];
function environment(){
 const c=vm.createContext({console,structuredClone});vm.runInContext(script.replace(/boot\(\);\s*$/,''),c);
 c.document={getElementById:()=>null,querySelectorAll:()=>[]};
 vm.runInContext(`state.forms={classes:[{courseId:'100',className:'クラス1'},{courseId:'200',className:'クラス2'}],defaults:{},records:[
 {id:'a',kind:'form',stage:'registered',revision:'r1',courseId:'100',label:'1A',sheetName:'回答1',updatedAt:'2026-10-01',input:{templateId:'template-1234567890',folderId:'folder-1234567890',titlePattern:'{class_label} 前の名前',materialTitlePattern:'{class_label} 資料',description:'保存済み本文',policy:{domains:['example.com'],emails:[]},labels:{100:'1A'},emailHeader:'メール',nameHeader:'氏名'}},
 {id:'b',kind:'form',stage:'published',revision:'r2',courseId:'200',label:'2A',sheetName:'回答2',input:{titlePattern:'{class_label} 別の名前',materialTitlePattern:'{class_label} 独自資料',description:'独自本文',policy:{domains:['other.example'],emails:['kid@example.net']}}}
 ]};`,c);return c;
}
test('reopening creation uses persisted settings and each prepared record offers editing',()=>{
 const c=environment(),html=vm.runInContext("renderForms('prepare')",c);
 assert.match(vm.runInContext("renderForms('publish')",c),/保存済み本文/);assert.match(html,/folder-1234567890/);assert.match(html,/data-form-command="edit-settings" data-record="a"/);
 assert.doesNotMatch(html,/id="form-verifiedEmailConfirmed" checked/);
});
test('publishing hides form settings editing while preserving posting fields and preparation drafts',()=>{
 const c=environment();vm.runInContext("startFormSettingsEdit(state.forms.records[0]);state.formSettingsEdit.values.titlePattern='編集中のフォーム名';",c);
 const html=vm.runInContext("renderForms('publish')",c);
 assert.doesNotMatch(html,/data-form-command="edit-settings"|id="formSettingsEditor"|data-form-command="save-settings"/);
 assert.match(html,/data-form-command="save-material-settings"/);
 assert.match(html,/保存済み本文/);
 const prepare=vm.runInContext("renderForms('prepare')",c);
 assert.match(prepare,/data-form-command="edit-settings" data-record="a"/);
 assert.match(prepare,/id="formSettingsEditor"/);
 assert.match(prepare,/編集中のフォーム名/);
});

test('publishing cannot open or save the form settings editor through stale actions',async()=>{
 const c=environment();vm.runInContext(`state.panel='setup';state.step='publish';let calls=[];
 readCurrent=()=>{};requireSaved=()=>true;task=async fn=>fn();render=()=>{};status=()=>{};confirmAction=async()=>true;
 rpc=async(method)=>{calls.push(method);return {targets:[]};};`,c);
 await vm.runInContext("formCommand('edit-settings','a')",c);
 assert.equal(vm.runInContext('state.formSettingsEdit',c),undefined);
 vm.runInContext("startFormSettingsEdit(state.forms.records[0]);state.formSettingsEdit.values.titlePattern='編集中のフォーム名';",c);
 await vm.runInContext("formCommand('save-settings')",c);
 assert.equal(vm.runInContext('calls.length',c),0);
 assert.equal(vm.runInContext('state.formSettingsEdit.values.titlePattern',c),'編集中のフォーム名');
});

test('editor loads all editable values and emits only changed fields including explicit blanks',()=>{
 const c=environment();vm.runInContext("startFormSettingsEdit(state.forms.records[0]);",c);
 let html=vm.runInContext("renderForms('prepare')",c);assert.match(vm.runInContext("renderForms('publish')",c),/保存済み本文/);assert.match(html,/example.com/);assert.match(html,/回答タブ名/);
 vm.runInContext("state.formSettingsEdit.values.description='';state.formSettingsEdit.recordIds.push('b');",c);
 const patch=JSON.parse(vm.runInContext('JSON.stringify(formSettingsPatch())',c));assert.deepEqual(patch,{description:''});
 vm.runInContext("acceptFormConsoleData({...state.forms,defaults:{emailHeader:'changed'}})",c);
 assert.deepEqual(JSON.parse(vm.runInContext('JSON.stringify(formSettingsPatch())',c)),{description:''});
});
test('different target settings stay untouched when only a shared message is changed',async()=>{
 const c=environment();vm.runInContext(`startFormSettingsEdit(state.forms.records[0]);state.formSettingsEdit.values.description='新本文';state.formSettingsEdit.recordIds=['a','b'];
 let calls=[];readCurrent=()=>{};requireSaved=()=>true;task=async fn=>fn();confirmAction=async()=>true;render=()=>{};status=()=>{};reloadForms=async()=>{};
 rpc=async(method,...args)=>{calls.push([method,...args]);if(method==='previewManagedFormSettings')return {targets:[{id:'a',label:'1A',changed:true,changes:['本文'],fingerprint:'pa'},{id:'b',label:'2A',changed:true,changes:['本文'],fingerprint:'pb'}]};return {id:args[0]};};`,c);
 await vm.runInContext("formCommand('save-settings')",c);
 const calls=JSON.parse(vm.runInContext('JSON.stringify(calls)',c));assert.deepEqual(calls[0][1],{targets:[{id:'a',revision:'r1'},{id:'b',revision:'r2'}],changes:{description:'新本文'}});
 assert.deepEqual(calls.filter(c=>c[0]==='runManagedFormSettings').map(c=>c[2]),[{description:'新本文'},{description:'新本文'}]);
 assert.equal(vm.runInContext('state.formSettingsEdit',c),null);
});
test('batch partial failure keeps failed inputs but removes completed targets from retries',async()=>{
 const c=environment();vm.runInContext(`startFormSettingsEdit(state.forms.records[0]);state.formSettingsEdit.values.description='新本文';state.formSettingsEdit.recordIds=['a','b'];
 readCurrent=()=>{};requireSaved=()=>true;task=async fn=>fn();confirmAction=async()=>true;render=()=>{};status=()=>{};reloadForms=async()=>{};
 rpc=async(method,id)=>{if(method==='previewManagedFormSettings')return {targets:[{id:'a',label:'1A',changed:true,changes:['本文'],fingerprint:'pa'},{id:'b',label:'2A',changed:true,changes:['本文'],fingerprint:'pb'}]};if(id==='b')throw Error('失敗');return {};};`,c);
 await vm.runInContext("formCommand('save-settings')",c);
 assert.equal(vm.runInContext('state.formSettingsEdit.values.description',c),'新本文');assert.equal(vm.runInContext('state.formSettingsEdit.recordIds.join()',c),'b');
});
test('refresh reads editor input before rerender and interrupted settings offer resuming',()=>{
 const c=environment();vm.runInContext('startFormSettingsEdit(state.forms.records[0]);',c);
 c.document={getElementById:id=>id==='edit-description'?{value:'入力中'}:null,querySelectorAll:()=>[]};
 vm.runInContext('readCurrent()',c);assert.equal(vm.runInContext('state.formSettingsEdit.values.description',c),'入力中');
 vm.runInContext('state.forms.records[0].settingsUpdate={};',c);
 assert.match(vm.runInContext("renderFormRecords('prepare')",c),/data-form-command="resume-settings"/);
 assert.doesNotMatch(vm.runInContext("renderFormRecords('publish')",c),/data-form-command="update-form" data-record="a"/);
});
test('saved creation settings can be reloaded in the same session without silently replacing a draft',async()=>{
 const c=environment();vm.runInContext(`state.formDraft=defaultFormDraft();state.formDraft.description='新規作成用の下書き';
 state.forms.records[0].input.description='編集後の保存済み本文';let confirmed=0;
 confirmAction=async()=>{confirmed++;return false;};render=()=>{};`,c);
 await vm.runInContext("formCommand('load-creation-settings')",c);assert.equal(vm.runInContext('state.formDraft.description',c),'新規作成用の下書き');
 vm.runInContext('confirmAction=async()=>true;',c);await vm.runInContext("formCommand('load-creation-settings')",c);
 assert.equal(vm.runInContext('state.formDraft.description',c),'編集後の保存済み本文');assert.equal(vm.runInContext('state.formDraft.verifiedEmailConfirmed',c),false);
});
