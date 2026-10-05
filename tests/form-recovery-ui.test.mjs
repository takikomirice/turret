import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import {readFileSync} from 'node:fs';

const script=readFileSync('Setting.html','utf8').match(/<script>([\s\S]*?)<\/script>/)[1];
function environment(){
 const c=vm.createContext({console,structuredClone});
 vm.runInContext(script.replace(/boot\(\);\s*$/,''),c);
 c.document={getElementById:()=>null,querySelectorAll:()=>[]};
 vm.runInContext(`state.forms={defaults:{},records:[
  {id:'failed',kind:'form',stage:'creating',revision:'old',label:'2-2',title:'振り返り 2-2',input:{}},
  {id:'ready',kind:'form',stage:'registered',revision:'ready-rev',label:'2-3',input:{}}
 ]};
 let calls=[],confirmations=[],messages=[],reloads=0;
 readCurrent=()=>{};requireSaved=()=>true;task=async fn=>fn();render=()=>{};
 status=(...args)=>messages.push(args);reloadForms=async()=>{reloads++;};
 confirmAction=async(...args)=>{confirmations.push(args);return true;};
 rpc=async(method,...args)=>{calls.push([method,...args]);return {id:args[0],stage:'pending',revision:'new'};};`,c);
 return c;
}
const read=(c,expr)=>JSON.parse(vm.runInContext('JSON.stringify('+expr+')',c));

test('only a creating record with a completed zero-match check offers individual recreation',()=>{
 const c=environment();
 assert.doesNotMatch(vm.runInContext("renderFormRecords('prepare')",c),/data-form-command="retry-form"/);
 vm.runInContext("state.forms.records.forEach(r=>r.copyNotFound=true)",c);
 const html=vm.runInContext("renderFormRecords('prepare')",c);
 assert.match(html,/data-form-command="retry-form" data-record="failed"/);
 assert.doesNotMatch(html,/data-form-command="retry-form" data-record="ready"/);
});

test('recreation confirms the selected class then prepares only that record with the fresh revision',async()=>{
 const c=environment();vm.runInContext('state.forms.records[0].copyNotFound=true',c);
 await vm.runInContext("formCommand('retry-form','failed')",c);
 assert.deepEqual(read(c,'calls'),[
  ['reconcileManagedForm','failed','old',true],['prepareFormTarget','failed','new']
 ]);
 assert.match(read(c,'confirmations')[0][1],/2-2/);
 assert.match(read(c,'confirmations')[0][1],/保存先フォルダ/);
 assert.equal(read(c,'reloads'),1);
});

test('cancelling recreation and stale buttons cause no server calls',async()=>{
 const c=environment();
 await vm.runInContext("formCommand('retry-form','failed')",c);
 await vm.runInContext("formCommand('retry-form','ready')",c);
 assert.equal(read(c,'calls.length'),0);assert.equal(read(c,'confirmations.length'),0);
 vm.runInContext('state.forms.records[0].copyNotFound=true;confirmAction=async()=>false',c);
 await vm.runInContext("formCommand('retry-form','failed')",c);
 assert.equal(read(c,'calls.length'),0);
});

test('a zero-match check displays recovery instructions without reporting completion or retrying',async()=>{
 const c=environment();
 vm.runInContext("rpc=async(method,...args)=>{calls.push([method,...args]);return {stage:'creating',copyNotFound:true,lastError:'保存先フォルダを確認してください'};}",c);
 await vm.runInContext("formCommand('reconcile-form','failed')",c);
 assert.deepEqual(read(c,'calls'),[['reconcileManagedForm','failed','old']]);
 assert.deepEqual(read(c,'messages')[0],['保存先フォルダを確認してください','error']);
 assert.equal(read(c,'reloads'),1);
});

for(const phase of ['reconcileManagedForm','prepareFormTarget'])test('a failed recreation refreshes the saved state and stops subsequent operations: '+phase,async()=>{
 const c=environment();vm.runInContext('state.forms.records[0].copyNotFound=true',c);
 c.failurePhase=phase;
 vm.runInContext("rpc=async(method,...args)=>{calls.push([method,...args]);if(method===failurePhase)throw Error('failure');return {stage:'created',revision:'new'};}",c);
 await assert.rejects(vm.runInContext("formCommand('retry-form','failed')",c),/failure/);
 assert.equal(read(c,'reloads'),1);
 assert.equal(read(c,'calls.length'),phase==='reconcileManagedForm'?1:2);
});

function previewEnvironment(){
 const c=environment();
 vm.runInContext(`state.formDraft={labels:{},statusHeader:'送信状態'};state.forms.configRevision='config';
 state.forms.records=[
  {id:'ready',kind:'form',stage:'registered',label:'2-3',revision:'ready-rev'},
  {id:'failed',kind:'form',stage:'creating',label:'2-2',revision:'failed-rev'}
 ];
 const plan={targets:[
  {courseId:'100',recordId:'ready',action:'skip',className:'2-3',title:'振り返り 2-3',sheetName:'回答 2-3'},
  {courseId:'200',recordId:'failed',action:'recreate',className:'2-2',title:'振り返り 2-2',sheetName:'回答 2-2'}
 ],configRevision:'config',fingerprint:'plan'};
 rpc=async(method,...args)=>{
  calls.push([method,...args]);
  if(method==='getFormConsoleData')return state.forms;
  if(method==='previewFormSetup')return plan;
  if(method==='beginFormSetup')return state.forms.records;
  if(method==='reconcileManagedForm')return args[2]===true?{id:args[0],label:'2-2',stage:'pending',revision:'reset-rev'}:{id:args[0],label:'2-2',stage:'creating',copyNotFound:true,revision:'checked-rev'};
  return {stage:'registered'};
 };`,c);
 return c;
}

test('creation confirmation shows skips and recreation then only prepares the unfinished record',async()=>{
 const c=previewEnvironment();
 await vm.runInContext("formCommand('preview')",c);
 const calls=read(c,'calls');
 assert.deepEqual(calls.filter(call=>call[0]==='prepareFormTarget'),[['prepareFormTarget','failed','reset-rev']]);
 assert.deepEqual(calls.filter(call=>call[0]==='reconcileManagedForm'),[
  ['reconcileManagedForm','failed','failed-rev'],['reconcileManagedForm','failed','checked-rev',true]
 ]);
 assert.match(read(c,'confirmations')[0][1],/スキップ.*2-3/);
 assert.match(read(c,'confirmations')[0][1],/再作成.*2-2/);
});

test('a skip-only preview reports completion without confirmation or writes',async()=>{
 const c=previewEnvironment();vm.runInContext("plan.targets=plan.targets.filter(t=>t.action==='skip')",c);
 await vm.runInContext("formCommand('preview')",c);
 assert.equal(read(c,'confirmations.length'),0);
 assert.deepEqual(read(c,'calls.map(c=>c[0])'),['getFormConsoleData','previewFormSetup']);
 assert.match(read(c,'messages')[0][0],/準備済み/);
});

test('a newly missing form is not recreated when the confirmation only authorized continuing an existing copy',async()=>{
 const c=previewEnvironment();vm.runInContext("plan.targets[1].action='reconcile'",c);
 await vm.runInContext("formCommand('preview')",c);
 assert.deepEqual(read(c,"calls.filter(c=>c[0]==='reconcileManagedForm')"),[['reconcileManagedForm','failed','failed-rev']]);
 assert.equal(read(c,"calls.filter(c=>c[0]==='prepareFormTarget').length"),0);
 assert.equal(read(c,'messages.at(-1)[1]'),'error');
});
