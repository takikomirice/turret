import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import vm from 'node:vm';
import {adminEnvironment, plain} from './helpers/admin-environment.mjs';

function env(){
 const e=adminEnvironment(),{c,ss}=e;
 c.ScriptApp.getScriptId=()=> 'script-one';
 ss.getName=()=> 'test';
 c.FormApp={openById:id=>({getDestinationId:()=>ss.getId(),getId:()=>id}),openByUrl:()=>({getId:()=> 'form-one'})};
 const sheet=ss.insertSheet('回答 1-1');sheet.getFormUrl=()=> 'form-one';
 sheet.rows=[['メール','質問','名前','管理','状態'],['kid@example.test','回答','生徒','{"v":1}','済']];
 sheet.validations=[];sheet.validationClears=[];
 const range=sheet.getRange.bind(sheet);
 sheet.getRange=(...args)=>Object.assign(range(...args),{setDataValidation(rule){sheet.validations.push({column:args[1],rule});return this;},clearDataValidations(){sheet.validationClears.push(args[1]);return this;}});
 c.SpreadsheetApp.newDataValidation=()=>({requireValueInList(values){this.values=plain(values);return this;},setAllowInvalid(){return this;},build(){return this.values;}});
 const record=c.saveManagedRecord_({id:'r1',kind:'form',stage:'registered',formId:'form-one',responseSheetId:sheet.getSheetId(),sheetName:sheet.name,label:'1-1',input:{emailHeader:'メール',nameHeader:'名前',statusHeader:'状態',additionalColumns:[]},baseHeaders:['メール','質問'],baseColumnCount:2});
 const raw={recordIds:[record.id],additionalColumns:[{header:'点数',choices:['1','2','3']},{header:'コメント',choices:[]}]};
 const preview=()=>c.previewResponseColumns(raw,c.getConfigRevision_(c.getConfig_()));
 const apply=()=>{const p=preview();return c.applyResponseColumns(raw,p.configRevision,p.fingerprint);};
 return {...e,sheet,record,raw,preview,apply};
}

test('second step appends evaluation columns, registers fields and preserves answers and system values',()=>{
 const e=env(),before=JSON.stringify(e.sheet.rows);const p=e.preview();
 assert.equal(JSON.stringify(e.sheet.rows),before);
 assert.deepEqual(plain(p.targets[0].added),['点数','コメント']);
 const old=plain(e.c.getConfig_().fields[0]);const result=e.apply();
 assert.equal(result.completed,1);assert.equal(result.errors.length,0);
 assert.deepEqual(e.sheet.rows[0],['メール','質問','名前','点数','コメント','管理','状態']);
 assert.deepEqual(e.sheet.rows[1],['kid@example.test','回答','生徒','','','{"v":1}','済']);
 assert.deepEqual(plain(e.c.getConfig_().fields[0]),old);
 assert.equal(e.c.getConfig_().scoreSourceHeader,'点数');
 assert.deepEqual(plain(e.c.getConfig_().fields.map(f=>f.sourceHeader)),['点数','コメント']);
 assert.deepEqual(e.sheet.validations[0],{column:4,rule:['1','2','3']});
 assert.equal(e.isLocked(),false);
});

test('name-only setup creates free-input columns and preserves existing validation and recorded choices',()=>{
 const e=env();e.apply();const validations=e.sheet.validations.length;e.sheet.validationClears=[];
 e.raw.columnNamesOnly=true;e.raw.additionalColumns=[{header:'点数'},{header:'観点'}];
 const p=e.preview();assert.equal(p.columnNamesOnly,true);assert.equal(e.apply().completed,1);
 assert.equal(e.sheet.validations.length,validations);
 assert.deepEqual(e.sheet.validationClears,[e.sheet.rows[0].indexOf('観点')+1]);
 const stored=e.c.loadManagedRecord_('r1').input.additionalColumns;
 assert.deepEqual(plain(stored.find(c=>c.header==='点数').choices),['1','2','3']);
 assert.deepEqual(plain(stored.find(c=>c.header==='観点').choices),[]);
});

test('name-only mode is validated and included in confirmation fingerprints',()=>{
 const e=env();e.raw.additionalColumns=[{header:'観点'}];const p=e.preview();
 e.raw.columnNamesOnly=true;assert.throws(()=>e.c.applyResponseColumns(e.raw,p.configRevision,p.fingerprint),/変更|再確認/);
 e.raw.additionalColumns=[{header:'観点',choices:['1']}];assert.throws(e.preview,/選択肢/);
 e.raw.columnNamesOnly='true';assert.throws(e.preview,/形式|真偽値/);
});

test('name-only setup resumes after new headers were written and preserves preexisting rules',()=>{
 const e=env();e.apply();e.sheet.validationClears=[];
 e.raw.columnNamesOnly=true;e.raw.additionalColumns=[{header:'点数'},{header:'観点'}];
 const getRange=e.sheet.getRange;e.sheet.getRange=(...args)=>{const r=getRange(...args);if(args[0]===2)r.clearDataValidations=()=>{throw Error('validation write failed');};return r;};
 const result=e.apply();assert.equal(result.completed,0);assert.match(result.errors[0].message,/validation write failed/);
 const pending=e.c.loadManagedRecord_('r1');assert.equal(pending.columnSetup.columnNamesOnly,true);
 e.sheet.getRange=getRange;e.c.resumeResponseColumns(pending.id,pending.revision);
 assert.deepEqual(e.sheet.validationClears,[e.sheet.rows[0].indexOf('観点')+1]);
 assert.deepEqual(plain(e.c.loadManagedRecord_('r1').input.additionalColumns.find(c=>c.header==='点数').choices),['1','2','3']);
});

test('ordinary field saves reject body-header renames and retain existing aliases',()=>{
 const e=env(),revision=e.c.getConfigRevision_(e.c.getConfig_());
 assert.throws(()=>e.c.saveSetupSection('fields',{replyBodyHeader:'別本文'},revision),/本文|変更/);
 const config=e.c.getConfig_();config.replyBodyHeader='旧本文';e.props.set('APP_CONFIG',JSON.stringify(config));
 e.c.saveSetupSection('fields',{replyBodyHeader:'旧本文'},e.c.getConfigRevision_(e.c.getConfig_()));
 assert.equal(e.c.getConfig_().replyBodyHeader,'旧本文');
});

test('reapplication reuses existing columns and does not erase grades or add duplicates',()=>{
 const e=env();e.apply();e.sheet.rows[1][3]='3';e.sheet.rows[1][4]='よい';
 const p=e.preview();assert.deepEqual(plain(p.targets[0].added),[]);
 e.apply();assert.equal(e.sheet.rows[1][3],'3');assert.equal(e.sheet.rows[1][4],'よい');assert.equal(e.sheet.rows[0].length,7);
 e.raw.additionalColumns=[{header:'観点',choices:[]}];e.apply();
 assert.deepEqual(e.sheet.rows[0],['メール','質問','名前','点数','コメント','観点','管理','状態']);
 assert.deepEqual(plain(e.c.loadManagedRecord_('r1').input.additionalColumns.map(c=>c.header)),['点数','コメント','観点']);
});

test('older prepared response tabs receive a missing management column automatically',()=>{
 const e=env();e.sheet.rows.forEach(row=>row.splice(3,1));
 const result=e.apply();assert.equal(result.completed,1);
 assert.deepEqual(e.sheet.rows[0],['メール','質問','名前','点数','コメント','管理','状態']);
 assert.equal(e.sheet.rows[1][6],'済');assert.equal(e.sheet.rows[1][1],'回答');
});

test('column setup stays within native bounds without creating automatic table headers',()=>{
 for(const missingManagement of [false,true]){
  const e=env();if(missingManagement)e.sheet.rows.forEach(row=>row.splice(3,1));
  e.sheet.maxColumns=e.sheet.getLastColumn();e.raw.columnNamesOnly=true;
  e.raw.additionalColumns=[{header:'観点'}];
  const lastColumn=e.sheet.getLastColumn.bind(e.sheet),getRange=e.sheet.getRange;let pendingWidth;
  e.sheet.getLastColumn=()=>pendingWidth===undefined?lastColumn():pendingWidth;
  e.sheet.getRange=(...args)=>{const r=getRange(...args),write=r.setValues;r.setValues=values=>{if(args.length>2&&args[0]===1&&args[1]>lastColumn())pendingWidth=lastColumn();return write(values);};return r;};
  e.c.SpreadsheetApp.flush=()=>{pendingWidth=undefined;};
  const move=e.sheet.moveColumns.bind(e.sheet);e.sheet.moveColumns=(range,to)=>{if(to>e.sheet.getMaxColumns())throw Error('Those columns are out of bounds.');return move(range,to);};
  const result=e.apply();assert.deepEqual(plain(result.errors),[]);assert.equal(result.completed,1);
  assert.deepEqual(e.sheet.rows[0].slice(-3),['観点','管理','状態']);
  assert.equal(e.sheet.getMaxColumns(),e.sheet.rows[0].length);
 }
});

test('column setup rejects system and question collisions before writing',()=>{
 for(const header of ['管理','状態','名前','メール','質問','タイムスタンプ','Timestamp']){
  const e=env();e.raw.additionalColumns=[{header,choices:[]}];const before=JSON.stringify(e.sheet.rows);
  assert.throws(e.preview,/列|質問/);assert.equal(JSON.stringify(e.sheet.rows),before);
 }
 for(const columns of [[],null,[{header:'同じ'},{header:'同じ'}],[{header:'評価',choices:{}}]]){
  const e=env();e.raw.additionalColumns=columns;assert.throws(e.preview,/列|選択肢/);
 }
});

test('stale sheet headers or settings stop application before mutation',()=>{
 for(const change of ['headers','config','record']){
  const e=env(),p=e.preview();
  if(change==='headers')e.sheet.rows[0][1]='変わった質問';
  if(change==='config')e.props.set('APP_CONFIG',JSON.stringify({...plain(e.c.getConfig_()),replyBodyHeader:'別本文'}));
  if(change==='record')e.c.saveManagedRecord_(e.c.loadManagedRecord_('r1'));
  const before=JSON.stringify(e.sheet.rows);
  assert.throws(()=>e.c.applyResponseColumns(e.raw,p.configRevision,p.fingerprint),/更新|変更/);
  assert.equal(JSON.stringify(e.sheet.rows),before);
 }
});

test('failed configuration save leaves resumable state and preserves grades on retry',()=>{
 const e=env(),save=e.c.applyConfigDraft_;e.c.applyConfigDraft_=()=>{throw Error('storage failed');};
 const result=e.apply();assert.equal(result.completed,0);assert.match(result.errors[0].message,/storage failed/);
 const r=e.c.loadManagedRecord_('r1');assert.ok(r.columnSetup);
 assert.throws(()=>e.c.assertNoManagedFormUpdate_(),/列/);
 e.sheet.rows[1][3]='2';e.c.applyConfigDraft_=save;
 e.c.resumeResponseColumns(r.id,r.revision);
 assert.equal(e.sheet.rows[1][3],'2');assert.equal(e.sheet.rows[0].length,7);
 assert.equal(e.c.loadManagedRecord_('r1').columnSetup,undefined);
 assert.doesNotThrow(()=>e.c.assertNoManagedFormUpdate_());assert.equal(e.isLocked(),false);
});

test('partial header write resumes without losing cells or duplicating columns',()=>{
 const e=env(),getRange=e.sheet.getRange.bind(e.sheet);let failed=false;
 e.sheet.getRange=(...args)=>{const r=getRange(...args),write=r.setValues;r.setValues=values=>{if(args[0]===1&&args[1]===6&&!failed){failed=true;write([values[0].slice(0,1)]);throw Error('connection lost');}return write(values);};return r;};
 const result=e.apply();assert.equal(result.completed,0);
 const pending=e.c.loadManagedRecord_('r1');e.c.resumeResponseColumns(pending.id,pending.revision);
 assert.deepEqual(e.sheet.rows[0],['メール','質問','名前','点数','コメント','管理','状態']);
 assert.deepEqual(e.sheet.rows[1],['kid@example.test','回答','生徒','','','{"v":1}','済']);
});

test('unprepared, updating, copied or internal records cannot be targeted',()=>{
 for(const change of ['pending','update','copied','internal','missing']){
  const e=env(),r=e.c.loadManagedRecord_('r1');
  if(change==='pending')r.stage='linked';
  if(change==='update')r.formUpdate={};
  if(change==='copied')r.ownerScriptId='other';
  if(change==='internal')r.kind='batch';
  if(change==='missing')r.responseSheetId=-1;
  if(change==='copied'){e.c.ScriptApp.getScriptId=()=> 'other-script';}else e.c.saveManagedRecord_(r);
  assert.throws(e.preview);assert.equal(e.sheet.rows[0].length,5);
 }
});

test('normal settings save cannot rename the system status column and legacy names remain usable',()=>{
 const e=env(),revision=e.c.getConfigRevision_(e.c.getConfig_());
 assert.throws(()=>e.c.saveSetupSection('fields',{formStatusHeader:'新状態'},revision),/送信状態|システム/);
 e.c.saveSetupSection('fields',{formStatusHeader:'状態',scoreSourceHeader:'別評価'},revision);
 assert.equal(e.c.getConfig_().formStatusHeader,'状態');
});

test('batch column setup reports partial success and resumes only the failed target',()=>{
 const e=env(),second=e.ss.insertSheet('回答 1-2');second.rows=[['メール','質問','名前','管理','状態']];second.getFormUrl=()=> 'form-two';
 e.c.FormApp.openByUrl=id=>({getId:()=>id});e.sheet.getFormUrl=()=> 'form-one';
 e.c.saveManagedRecord_({...plain(e.record),id:'r2',formId:'form-two',responseSheetId:second.getSheetId(),sheetName:second.name,label:'1-2'});
 e.raw.recordIds=['r1','r2'];e.raw.additionalColumns=[{header:'コメント',choices:[]}];
 let saves=0;const save=e.c.applyConfigDraft_;e.c.applyConfigDraft_=(...args)=>{if(++saves===2)throw Error('second failed');return save(...args);};
 const result=e.apply();assert.equal(result.completed,1);assert.equal(result.errors[0].id,'r2');
 assert.equal(e.c.loadManagedRecord_('r1').columnSetup,undefined);assert.ok(e.c.loadManagedRecord_('r2').columnSetup);
 e.c.applyConfigDraft_=save;const pending=e.c.loadManagedRecord_('r2');e.c.resumeResponseColumns(pending.id,pending.revision);
 assert.deepEqual(second.rows[0],['メール','質問','名前','コメント','管理','状態']);assert.equal(e.sheet.rows[0].length,6);
});

test('all targets are validated before writing and missing system headers fail safely',()=>{
 const e=env();e.raw.recordIds.push('missing');assert.throws(e.apply);assert.equal(e.sheet.rows[0].length,5);
 e.raw.recordIds=['r1'];e.sheet.rows[0][4]='別状態';assert.throws(e.preview,/列/);assert.equal(e.sheet.rows[0].length,5);
});

test('after interrupted setup external header changes are not silently accepted on resume',()=>{
 const e=env();e.c.applyConfigDraft_=()=>{throw Error('storage failed');};e.apply();
 e.sheet.rows[0][1]='質問の変更';const before=JSON.stringify(e.sheet.rows),r=e.c.loadManagedRecord_('r1');
 assert.throws(()=>e.c.resumeResponseColumns(r.id,r.revision),/変更/);assert.equal(JSON.stringify(e.sheet.rows),before);assert.equal(e.isLocked(),false);
});

function ui(){
 const c=vm.createContext({console,structuredClone});
 const script=readFileSync('Setting.html','utf8').match(/<script>([\s\S]*?)<\/script>/)[1];
 vm.runInContext(script.replace(/boot\(\);\s*$/,''),c);
 vm.runInContext("state.forms={classes:[],defaults:{statusHeader:'状態'},records:[{id:'r1',kind:'form',stage:'registered',responseSheetId:1,label:'1-1',sheetName:'回答 1-1'}]};state.data={config:{fields:[],formStatusHeader:'状態',gradeScale:[]}}",c);
 return c;
}

test('form creation and evaluation setup have separate controls and system column names are fixed',()=>{
 const c=ui();const form=vm.runInContext('renderForms()',c);
 assert.doesNotMatch(form,/data-extra-header|data-form-command="add-column"|id="form-statusHeader"/);
 assert.match(form,/管理/);assert.match(form,/状態/);
 const columns=vm.runInContext('renderResponseColumns()',c);
 assert.match(columns,/data-column-target="r1"/);assert.match(columns,/data-form-command="add-column"/);
 const fields=vm.runInContext('renderFields()',c);
 assert.doesNotMatch(fields,/<(?:input|select)[^>]*id="formStatusHeader"/);
 assert.match(fields,/状態/);
});

test('creation request discards evaluation draft so step one cannot add evaluation columns',async()=>{
 const c=ui();
 vm.runInContext("state.formDraft=defaultFormDraft();state.formDraft.additionalColumns=[{header:'評価',choices:[]}];const calls=[];readCurrent=()=>{};requireSaved=()=>true;task=async f=>f();confirmAction=async()=>false;rpc=async(method,...args)=>{calls.push({method,args});if(method==='getFormConsoleData')return state.forms;return {input:{additionalColumns:[]},targets:[{courseId:'100',className:'1組',title:'1組 振り返り',sheetName:'回答 1組',action:'create'}],configRevision:'r',fingerprint:'p'};}",c);
 await vm.runInContext("formCommand('preview')",c);
 assert.deepEqual(JSON.parse(vm.runInContext("JSON.stringify(calls.find(c=>c.method==='previewFormSetup').args[0].additionalColumns)",c)),[]);
});

test('evaluation draft collects selection and names while navigation preserves the form draft',()=>{
 const c=ui();c.document={getElementById:id=>id==='responseColumns'?{}:null,querySelectorAll:selector=>selector==='[data-column-target]'?[{checked:true,dataset:{columnTarget:'r1'}},{checked:false,dataset:{columnTarget:'r2'}}]:selector==='[data-extra-column]'?[{querySelector:sel=>({value:sel==='[data-extra-header]'?'観点':'A,B'})}]:[]};
 vm.runInContext("state.panel='setup';state.step='columns';state.formDraft={description:'保持'};readCurrent();state.step='forms';readCurrent();",c);
 assert.deepEqual(JSON.parse(vm.runInContext('JSON.stringify(state.columnDraft)',c)),{recordIds:['r1'],additionalColumns:[{header:'観点'}]});
 assert.equal(vm.runInContext('state.formDraft.description',c),'保持');
});

test('column progress and target list show saved names for current and legacy records',()=>{
 const c=ui();
 vm.runInContext(`state.forms.records[0].input={additionalColumns:[{header:'点数'},{header:'<コメント>'}]};state.forms.records.push({id:'r2',kind:'form',stage:'registered',responseSheetId:2,sheetName:'回答 1-2',input:{gradeHeader:'旧評価',commentHeader:'旧コメント'}});`,c);
 const progress=vm.runInContext("visibleProgress('columns')",c);
 assert.equal(progress.state,'complete');assert.match(progress.detail,/2 \/ 2件/);
 for(const header of ['点数','<コメント>','旧評価','旧コメント'])assert.ok(progress.detail.includes(header));
 const markup=vm.runInContext('renderResponseColumns()',c);
 assert.match(markup,/設定済み：点数、&lt;コメント&gt;/);assert.match(markup,/設定済み：旧評価、旧コメント/);
 assert.doesNotMatch(markup,/<コメント>/);
});

test('additional column guidance is available through help instead of visible prose',()=>{
 const c=ui(),markup=vm.runInContext('renderResponseColumns()',c);
 assert.match(markup,/data-help-title="追加する列"/);
 assert.match(markup,/data-help-body="[^"]*新しい列は入力した順に追加/);
 assert.doesNotMatch(markup.replace(/<[^>]*>/g,''),/新しい列は入力した順に追加|追加列はメッセージへの取り込みルールにも登録/);
});

function columnConfirmationUi(e,recordIds=['r1'],additionalColumns=[]){
 const c=ui(),region={textContent:'',dataset:{}},calls=[],nodes={statusRegion:region};
 for(const id of ['workspace','rail','content','stepNav','draftStatus'])nodes[id]={classList:{toggle(){}},innerHTML:'',textContent:''};
 c.clearTimeout=clearTimeout;
 c.document={getElementById:id=>nodes[id]||null,querySelectorAll:()=>[]};
 c.rpc=async(method,...args)=>{
  calls.push(method);
  if(method==='getFormConsoleData')return {records:[plain(e.c.loadManagedRecord_('r1'))],defaults:{statusHeader:'状態'},configRevision:e.c.getConfigRevision_(e.c.getConfig_())};
  if(method==='previewResponseColumns'||method==='applyResponseColumns')return e.c[method](...args);
  throw Error('Unexpected RPC: '+method);
 };
 c.draft={recordIds,additionalColumns};
 vm.runInContext("state.panel='setup';state.step='columns';state.columnDraft=draft;readCurrent=()=>{};readColumnDraft=()=>{};requireSaved=()=>true;task=async f=>f();confirmAction=async()=>{throw Error('Read-only confirmation must not request writes');};",c);
 return {c,region,calls,nodes};
}

test('empty addition draft confirms latest saved columns without changing sheets or settings',async()=>{
 const e=env();e.apply();
 const {c,region,calls}=columnConfirmationUi(e),before=JSON.stringify({rows:e.sheet.rows,props:[...e.props]});
 await vm.runInContext("formCommand('preview-columns')",c);
 assert.match(region.textContent,/回答 1-1.*設定済み.*点数、コメント/);
 assert.notEqual(region.dataset.kind,'error');
 assert.deepEqual(calls,['getFormConsoleData']);
 assert.equal(JSON.stringify({rows:e.sheet.rows,props:[...e.props]}),before);
 assert.deepEqual(JSON.parse(vm.runInContext('JSON.stringify(state.columnDraft)',c)),{recordIds:['r1'],additionalColumns:[]});
});

test('empty addition draft explains how to add columns on an unconfigured target',async()=>{
 const e=env(),{c,region}=columnConfirmationUi(e);
 await vm.runInContext("formCommand('preview-columns')",c);
 assert.match(region.textContent,/回答 1-1.*未設定/);assert.match(region.textContent,/＋ 列を追加/);
 assert.notEqual(region.dataset.kind,'error');assert.equal(e.sheet.rows[0].length,5);
});

test('checking saved columns updates visible names and preserves target selection',async()=>{
 const e=env();e.apply();const {c,nodes}=columnConfirmationUi(e);
 vm.runInContext("state.forms.records[0].input={additionalColumns:[{header:'古い表示'}]};render()",c);
 assert.match(nodes.content.innerHTML,/古い表示/);
 await vm.runInContext("formCommand('preview-columns')",c);
 assert.doesNotMatch(nodes.content.innerHTML,/古い表示/);
 assert.match(nodes.content.innerHTML,/設定済み：点数、コメント/);
 assert.match(nodes.content.innerHTML,/data-column-target="r1" checked/);
});

test('checking saved columns rejects stale targets and still validates blank new column names',async()=>{
 for(const [recordIds,columns] of [[['missing'],[]],[[],[]],[['r1','r1'],[]],[['r1'],[{header:' '}]]]){
  const e=env(),{c}=columnConfirmationUi(e,recordIds,columns),before=JSON.stringify(e.sheet.rows);
  await assert.rejects(vm.runInContext("formCommand('preview-columns')",c),/対象|列/);
  assert.equal(JSON.stringify(e.sheet.rows),before);
 }
});
