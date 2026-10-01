import { existsSync, readFileSync } from 'node:fs';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
const html = readFileSync('Setting.html', 'utf8');
const script = html.match(/<script>([\s\S]*?)<\/script>/)[1];

test('managed forms show intake stop while publication result is unknown',()=>{
 const c=model();
 vm.runInContext("state.forms={classes:[],defaults:{},records:[{id:'r1',kind:'form',stage:'publish_review',label:'<script>bad</script>',materialTitle:'資料'}]};",c);
 const rendered=vm.runInContext("renderForms('publish')",c);
 assert.match(rendered,/data-form-command="close"/);
 assert.doesNotMatch(rendered,/data-form-command="open"/);
 assert.match(rendered,/&lt;script&gt;bad&lt;\/script&gt;/);
 assert.match(rendered,/投稿結果を照合/);
});
test('new console entrypoints retain explicit destructive confirmation',()=>{
 assert.match(html,/previewManagedFormAction/);assert.match(html,/runManagedFormAction/);
 assert.match(html,/preview\.fingerprint/);assert.match(html,/投稿ID：/);
 assert.match(html,/getWebConsoleBootstrap/);assert.match(html,/別GASへ5ファイル/);
});
function model() {
  const context = vm.createContext({ console, structuredClone });
  vm.runInContext(script.replace(/boot\(\);\s*$/, ''), context);
  return context;
}

test('verified email attestation is read from the checkbox and invalidated when the template changes',()=>{
 const c=model(),fields={'form-templateId':{value:'template-one'},'form-verifiedEmailConfirmed':{checked:true}};
 c.document={getElementById:id=>fields[id]||null,querySelectorAll:()=>[]};
 vm.runInContext("state.formDraft={templateId:'template-one',labels:{},additionalColumns:[]};readFormDraft()",c);
 assert.equal(vm.runInContext('state.formDraft.verifiedEmailConfirmed',c),true);
 fields['form-templateId'].value='template-two';vm.runInContext('readFormDraft()',c);
 assert.equal(vm.runInContext('state.formDraft.verifiedEmailConfirmed',c),false);assert.equal(fields['form-verifiedEmailConfirmed'].checked,false);
});

test('adding or reordering columns keeps the expanded creation form open after a form already exists',()=>{
 const c=model();c.document={getElementById:id=>id==='form-templateId'?{value:''}:id==='formCreateDetails'?{open:true}:null,querySelectorAll:()=>[]};
 vm.runInContext("state.forms={defaults:{},classes:[],records:[{kind:'form',id:'r',stage:'deleted'}]};readFormDraft()",c);
 assert.match(vm.runInContext("renderForms('prepare')",c),/<details id="formCreateDetails" open>/);
});

test('extra column up/down moves keep dropdown choices with their header and enforce boundaries',async()=>{
 const c=model();
 vm.runInContext("readFormDraft=()=>{};render=()=>{};state.formDraft={additionalColumns:[{header:'評価',choices:'1,2,3'},{header:'コメント',choices:''},{header:'観点',choices:'A,B'}]};",c);
 await vm.runInContext("formCommand('column-up','2')",c);
 assert.equal(vm.runInContext('JSON.stringify(state.formDraft.additionalColumns)',c),JSON.stringify([{header:'評価',choices:'1,2,3'},{header:'観点',choices:'A,B'},{header:'コメント',choices:''}]));
 await vm.runInContext("formCommand('column-down','0')",c);
 await vm.runInContext("formCommand('column-up','0')",c);
 assert.equal(vm.runInContext('state.formDraft.additionalColumns[0].header',c),'観点');
});

test('form update and resume controls distinguish finished forms from interrupted updates',()=>{
 const c=model();
 vm.runInContext("state.forms={classes:[],defaults:{},records:[{id:'r1',kind:'form',stage:'published',label:'テスト',materialId:'m1'}]};",c);
 assert.match(vm.runInContext("renderForms('publish')",c),/data-form-command="update-form"/);
 vm.runInContext("state.forms.records[0].formUpdate={phase:'columns'}",c);
 const markup=vm.runInContext("renderForms('publish')",c);
 assert.match(markup,/data-form-command="resume-update"/);
 assert.doesNotMatch(markup,/data-form-command="update-form"|data-form-command="delete"|data-form-command="open"/);
});

test('web setup sheet controls link directly to the exact tab and disable absent tabs',()=>{
 const c=interactiveModel();
 vm.runInContext("state.bootstrap={web:true};state.data.sheets=[{key:'classes',exists:true,url:'https://docs.google.com/spreadsheets/d/test/edit#gid=0'},{key:'students',exists:true,url:'https://docs.google.com/spreadsheets/d/test/edit#gid=42'}];",c);
 const markup=vm.runInContext("renderStep('classes')",c);
 assert.match(markup,/<a[^>]+href="https:\/\/docs.google.com\/spreadsheets\/d\/test\/edit#gid=0"[^>]*>クラス一覧を開く<\/a>/);
 assert.match(markup,/<a[^>]+href="https:\/\/docs.google.com\/spreadsheets\/d\/test\/edit#gid=42"/);
 assert.doesNotMatch(markup,/data-sheet=/);assert.match(vm.runInContext("sheetButton('mapping','対応表')",c),/disabled/);
 vm.runInContext('state.bootstrap.web=false;',c);assert.match(vm.runInContext("sheetButton('classes','クラス一覧')",c),/href="https:/);
});

function backgroundModel(){
 const c=interactiveModel();
 c.document={getElementById:()=>({scrollTop:0,focus(){},querySelectorAll:()=>[]}),activeElement:null};
 vm.runInContext("const renders=[];render=()=>renders.push(state.panel);readFormDraft=()=>{};let finish,fail;rpc=()=>new Promise((resolve,reject)=>{finish=resolve;fail=reject});state.forms={classes:[],defaults:{},records:[]};",c);
 return c;
}

test('forms navigation renders cached data before the background response and deduplicates requests',async()=>{
 const c=backgroundModel();
 const pending=vm.runInContext("handleClick({target:{closest:()=>({dataset:{step:'forms'}})}})",c);
 assert.equal(vm.runInContext('state.panel',c),'setup');
 assert.equal(vm.runInContext('state.step',c),'forms');
 assert.equal(vm.runInContext('state.busy',c),false);
 assert.match(vm.runInContext('renderForms()',c),/最新情報を取得中/);
 const promise=vm.runInContext('state.formsLoadingPromise',c);
 assert.equal(vm.runInContext('refreshFormsInBackground()',c),promise);
 vm.runInContext("finish({classes:[{className:'新クラス',courseId:'2'}],defaults:{},records:[]})",c);
 await pending;await promise;
 assert.equal(vm.runInContext('state.forms.classes[0].className',c),'新クラス');
});

test('background failures retain cache and expose retry guidance without returning to the old panel',async()=>{
 const c=backgroundModel();vm.runInContext("state.panel='setup';state.step='forms';refreshFormsInBackground()",c);
 const promise=vm.runInContext('state.formsLoadingPromise',c);vm.runInContext("state.panel='maintenance';fail(Error('通信エラー'))",c);
 await promise;
 assert.equal(vm.runInContext('state.panel',c),'maintenance');
 assert.equal(vm.runInContext('state.forms.records.length',c),0);
 assert.match(vm.runInContext('renderForms()',c),/通信エラー/);
 assert.match(vm.runInContext('renderForms()',c),/再取得/);
});

test('mutating tasks wait for background reads instead of contending for the server lock',async()=>{
 const c=backgroundModel();vm.runInContext("state.panel='setup';state.step='forms';refreshFormsInBackground();let ran=false",c);
 const work=vm.runInContext('task(async()=>{ran=true})',c);
 assert.equal(vm.runInContext('ran',c),false);
 vm.runInContext('finish({classes:[],defaults:{},records:[]})',c);await work;
 assert.equal(vm.runInContext('ran',c),true);
});

test('first forms visit stays responsive and offers retry when loading fails',async()=>{
 const c=backgroundModel();vm.runInContext("state.forms=null;state.panel='setup';state.step='forms';refreshFormsInBackground()",c);
 assert.match(vm.runInContext('renderForms()',c),/最新情報を取得中/);
 const promise=vm.runInContext('state.formsLoadingPromise',c);vm.runInContext("fail(Error('offline'))",c);await promise;
 assert.equal(vm.runInContext('state.formsLoadingPromise',c),null);
 assert.match(vm.runInContext('renderForms()',c),/再取得/);
 const retry=vm.runInContext('refreshFormsInBackground()',c);vm.runInContext('finish({classes:[],defaults:{},records:[]})',c);await retry;
 assert.equal(vm.runInContext('state.formsError',c),'');
});

test('background rerender keeps typed input, focus, selection, expanded details and scroll position',()=>{
 const c=interactiveModel();let focused=false,selection=[];
 const afterDetails=[{open:false}],content={scrollTop:120,querySelectorAll:()=>[{open:true}]};
 const input={focus(){focused=true;},setSelectionRange(...args){selection=args;}};
 const copyInputs={copyName:{value:'次の学期'},copyFolder:{value:'folder-url'}};
 c.document={getElementById:id=>id==='content'?content:copyInputs[id]||input,activeElement:{id:'form-description',selectionStart:2,selectionEnd:4}};
 c.afterDetails=afterDetails;
 vm.runInContext("state.panel='setup';state.step='forms';readFormDraft=()=>{state.formDraft={description:'編集中'}};render=()=>{document.getElementById('content').scrollTop=0;document.getElementById('content').querySelectorAll=()=>afterDetails;document.getElementById('copyName').value='';document.getElementById('copyFolder').value=''};renderUpdatedForms()",c);
 assert.equal(vm.runInContext('state.formDraft.description',c),'編集中');assert.equal(content.scrollTop,120);
 assert.equal(afterDetails[0].open,true);assert.equal(focused,true);assert.deepEqual(selection,[2,4]);
 assert.equal(copyInputs.copyName.value,'次の学期');assert.equal(copyInputs.copyFolder.value,'folder-url');
});
test('console is modeless so sheet tables can be edited', () => {
  const code = ['Code.gs', 'Administration.gs'].filter(existsSync).map(file => readFileSync(file, 'utf8')).join('\n');
  assert.match(code, /showModelessDialog\(/);
  assert.doesNotMatch(code.match(/function connectWebScreen_\([^]*?\n\}/)[0], /showSidebar\(/);
  assert.equal(existsSync('Sidebar.html'), false);
});
test('guided console offers every step, local JSON files, and accessible confirmation', () => {
  assert.deepEqual(JSON.parse(vm.runInContext('JSON.stringify(steps.map(s=>s.id))', model())), ['prepare','classes','forms','sources','fields','template','publish','automation']);
  assert.match(html,/data-panel="setup"[^>]*>設定<\/button>/);
  assert.doesNotMatch(html,/data-panel="forms"/);
  assert.match(html, /<dialog[^>]+id="confirmDialog"/);
  assert.match(html, /aria-live="polite"/);
  assert.match(html, /type="file"/);
  assert.match(html, /new Blob/);
  assert.match(html, /URL\.revokeObjectURL/);
  assert.doesNotMatch(html, /saveSettingsFromDialog/);
});
test('empty settings do not acquire fake completed steps', () => {
  const c = model();
  assert.equal(vm.runInContext("progressFor('sources').state", c), 'pending');
  assert.equal(vm.runInContext('automationReady()', c), false);
});
test('refresh keeps unsaved drafts and original concurrency revision', () => {
  const c = model();
  vm.runInContext("state.data={config:{messageTemplate:'old'},revision:'r1',progress:[]}; putDraft('template',{messageTemplate:'mine'}); acceptData({config:{messageTemplate:'external'},revision:'r2',progress:[]});", c);
  assert.equal(vm.runInContext("draftValue('template').messageTemplate", c), 'mine');
  assert.equal(vm.runInContext("draftRevision('template')", c), 'r1');
});
test('own save only rebases drafts from saved revision', () => {
  const c = model();
  vm.runInContext("state.data={config:{},revision:'r1',progress:[]}; putDraft('template',{messageTemplate:'mine'}); state.drafts.sources={payload:{},revision:'old'}; rebaseDrafts('r1','r2');", c);
  assert.equal(vm.runInContext("draftRevision('template')", c), 'r2');
  assert.equal(vm.runInContext("draftRevision('sources')", c), 'old');
});
test('destructive summaries include target counts and recipients', () => {
  const c = model();
  vm.runInContext("state.data={config:{reminderTo:['teacher@example.test']},counts:{pendingSend:7,reviewSend:2,sendRows:12,evalRows:9},progress:[]};", c);
  assert.match(vm.runInContext("actionSummary('send')", c), /7/);
  assert.match(vm.runInContext("actionSummary('send')", c), /2/);
  assert.match(vm.runInContext("actionSummary('clearSend')", c), /12/);
  assert.match(vm.runInContext("actionSummary('remind')", c), /teacher@example.test/);
});
test('automation requires complete server-confirmed preconditions', () => {
  const c = model();
  vm.runInContext("state.data={progress:requiredSteps.map(id=>({id,state:'complete'}))};", c);
  assert.equal(vm.runInContext('automationReady()', c), true);
  vm.runInContext("state.data.progress[3].state='attention'", c);
  assert.equal(vm.runInContext('automationReady()', c), false);
});

test('different object key order and normalized source IDs do not create phantom edits', () => {
 const c=model();
 vm.runInContext("state.data={config:{formSources:[{id:'123',url:'https://docs.google.com/spreadsheets/d/123/edit'}],fields:[]},revision:'r1'}; putDraft('sources',{reminderTo:[],formSheetNamePrefix:[],formSources:[{url:'https://docs.google.com/spreadsheets/d/123/edit'}]});",c);
 assert.equal(vm.runInContext('dirtySections().length',c),0);
});

test('automation is disabled for recovered or corrupt config even if steps look complete', () => {
 const c=model();
 vm.runInContext("state.data={recoveryRequired:true,progress:requiredSteps.concat('forms','publish','automation').map(id=>({id,state:'complete'}))};",c);
 assert.equal(vm.runInContext('automationReady()',c),false);
 vm.runInContext('state.data.recoveryRequired=false;state.data.ready=false;',c);
 assert.equal(vm.runInContext('automationReady()',c),false);
});

function interactiveModel() {
 const c=model();
 vm.runInContext(`
 const calls=[],messages=[];
 state.data={config:{fields:[],messageTemplate:'',reminderTo:['teacher@example.test']},revision:'r1',progress:requiredSteps.concat('forms','publish','automation').map(id=>({id,state:'complete'})),counts:{pendingSend:3},actionRevisions:{send:'snapshot1',remind:'snapshot2',clearSend:'snapshot3'},sheets:[]};
 state.automation={enabled:false,schedule:{importHour:5,deliveryHour:8,reminderHour:16,reminderEnabled:false},revision:'r1'};
 readCurrent=()=>{};render=()=>{};status=(...args)=>messages.push(args);setBusy=value=>state.busy=value;
 rpc=async(method,...args)=>{calls.push({method,args});if(method==='getAdminConsoleData')return state.data;if(method==='getAutomationState')return state.automation;return {message:'ok',revision:'r2',config:state.data.config};};
 confirmAction=async()=>true;
 `,c);
 return c;
}

test('manual return has prepare and send actions without an evaluation step',()=>{
 const c=model();
 vm.runInContext("state.data={config:{},counts:{evalRows:0,pendingSend:0,reviewSend:0},sheets:[]};",c);
 const markup=vm.runInContext('renderManual()',c);
 assert.match(markup,/送信データを準備/);assert.match(markup,/data-action="prepare"/);
 assert.doesNotMatch(markup,/data-action="import"|評価データ/);
});

test('legacy migration is shown only while old evaluation rows exist',()=>{
 const c=model();
 vm.runInContext("state.data={config:{},counts:{evalRows:2,legacyPending:1,pendingSend:0,reviewSend:0},sheets:[]};",c);
 const markup=vm.runInContext('renderManual()',c);
 assert.match(markup,/旧評価データ/);assert.match(markup,/未移行/);
});

test('paged return history renders the server page and actions for only the displayed records',()=>{
 const c=interactiveModel();
 vm.runInContext(`acceptFormConsoleData({classes:[],defaults:{},records:[{id:'last',kind:'announcement',title:'最終履歴',postId:'p',stage:'published'}],returnHistory:{page:2,pageSize:50,total:101}})`,c);
 const rendered=vm.runInContext("renderFormRecords('returns')",c);
 assert.match(rendered,/返却 101〜101 \/ 101件/);assert.match(rendered,/最終履歴/);
 assert.match(rendered,/data-form-command="delete"/);assert.match(rendered,/data-form-command="next-page" disabled/);
 assert.equal(vm.runInContext('state.formPage',c),2);
});
test('history navigation requests a page and commits it only on success while keeping drafts',async()=>{
 const c=interactiveModel();
 vm.runInContext(`
 state.panel='maintenance';state.formDraft={description:'入力保持'};renderUpdatedForms=()=>{};
 acceptFormConsoleData({records:[{id:'old',kind:'announcement'}],returnHistory:{page:0,pageSize:50,total:120}});
 let resolvePage;rpc=(method,...args)=>{calls.push({method,args});return new Promise(resolve=>{resolvePage=resolve;});};
 `,c);
 const pending=vm.runInContext("formCommand('next-page')",c);
 assert.equal(vm.runInContext('state.formPage',c),0);assert.equal(vm.runInContext('state.forms.records[0].id',c),'old');
 assert.deepEqual(JSON.parse(vm.runInContext('JSON.stringify(calls)',c)),[{method:'getFormConsoleData',args:[{returnPage:1}]}]);
 vm.runInContext("resolvePage({records:[{id:'new',kind:'announcement'}],returnHistory:{page:1,pageSize:50,total:120}})",c);await pending;
 assert.equal(vm.runInContext('state.formPage',c),1);assert.equal(vm.runInContext('state.formDraft.description',c),'入力保持');
 vm.runInContext("rpc=async()=>{throw Error('取得失敗');}",c);await vm.runInContext("formCommand('previous-page')",c);
 assert.equal(vm.runInContext('state.formPage',c),1);assert.equal(vm.runInContext('state.forms.records[0].id',c),'new');
 assert.match(vm.runInContext('messages.at(-1)[0]',c),/取得失敗/);assert.equal(vm.runInContext('state.busy',c),false);
});
test('refresh and background loading retain the requested history page without coupling the admin and automation endpoints',async()=>{
 const c=interactiveModel();
 vm.runInContext(`
 state.panel='maintenance';
 renderUpdatedForms=()=>{};
 acceptFormConsoleData({records:[],returnHistory:{page:2,pageSize:50,total:101}});
 rpc=async(method,...args)=>{calls.push({method,args});if(method==='getAdminConsoleData')return state.data;if(method==='getAutomationState')return state.automation;return {records:[],returnHistory:{page:1,pageSize:50,total:100}};};
 `,c);
 await vm.runInContext('refresh(false)',c);
 assert.deepEqual(JSON.parse(vm.runInContext('JSON.stringify(calls)',c)),[
  {method:'getAdminConsoleData',args:[]},{method:'getAutomationState',args:[]},{method:'getFormConsoleData',args:[{returnPage:2}]}
 ]);
 await vm.runInContext('refreshFormsInBackground()',c);
 assert.deepEqual(JSON.parse(vm.runInContext('JSON.stringify(calls.at(-1))',c)),{method:'getFormConsoleData',args:[{returnPage:1}]});
});

test('new template autofills its URL while preserving form inputs and unrelated drafts',async()=>{
 const c=interactiveModel();
 vm.runInContext(`
 state.forms={classes:[],records:[],defaults:{}};
 state.formDraft={...defaultFormDraft(),folderId:'keep-folder',description:'keep-description',additionalColumns:[{header:'観点',choices:'1,2'}]};
 putDraft('template',{messageTemplate:'pending'});
 const input={value:''};document={getElementById:id=>id==='form-templateId'?input:null};
 readFormDraft=()=>{};renderUpdatedForms=()=>{};
 rpc=async(method,...args)=>{calls.push({method,args});return {editUrl:'https://docs.google.com/forms/d/new-template-form/edit'};};
 `,c);
 await vm.runInContext("formCommand('create-template')",c);
 assert.equal(vm.runInContext('state.formDraft.templateId',c),'https://docs.google.com/forms/d/new-template-form/edit');
 assert.equal(vm.runInContext('input.value',c),'https://docs.google.com/forms/d/new-template-form/edit');
 assert.equal(vm.runInContext('state.formDraft.folderId',c),'keep-folder');
 assert.equal(vm.runInContext('state.formDraft.additionalColumns[0].header',c),'観点');
 assert.equal(vm.runInContext("draftValue('template').messageTemplate",c),'pending');
 assert.match(vm.runInContext('renderForms()',c),/ひな形を編集/);
 assert.equal(vm.runInContext('calls[0].method',c),'createTemplateForm');
});

test('template retry keeps its request key and cancelling URL replacement creates nothing',async()=>{
 const c=interactiveModel();
 vm.runInContext(`state.formDraft=defaultFormDraft();document={getElementById:()=>null};readFormDraft=()=>{};
 rpc=async(method,...args)=>{calls.push({method,args});throw Error('connection lost');};`,c);
 await vm.runInContext("formCommand('create-template')",c);
 await vm.runInContext("formCommand('create-template')",c);
 assert.equal(vm.runInContext('calls.length',c),2);
 assert.equal(vm.runInContext('calls[0].args[0]===calls[1].args[0]',c),true);
 assert.equal(vm.runInContext('state.busy',c),false);
 vm.runInContext("state.formDraft.templateId='existing';confirmAction=async()=>false;",c);
 await vm.runInContext("formCommand('create-template')",c);
 assert.equal(vm.runInContext('calls.length',c),2);
});

test('cancelled sending and clearing make no server call',async()=>{
 const c=interactiveModel();
 vm.runInContext('confirmAction=async()=>false;',c);
 await vm.runInContext("runAction('send')",c);
 await vm.runInContext("runAction('clearSend')",c);
 assert.equal(vm.runInContext('calls.length',c),0);
});

test('accepted send carries the displayed action snapshot and config revision',async()=>{
 const c=interactiveModel();
 await vm.runInContext("runAction('send')",c);
 const call=JSON.parse(vm.runInContext('JSON.stringify(calls[0])',c));
 assert.deepEqual(call,{method:'runAdminAction',args:['send',{confirmed:true,revision:'snapshot1'},'r1']});
 assert.equal(vm.runInContext('state.busy',c),false);
});

test('unrelated dirty settings block manual execution without losing data',async()=>{
 const c=interactiveModel();
 vm.runInContext("putDraft('template',{messageTemplate:'pending'});",c);
 await vm.runInContext("runAction('send')",c);
 assert.equal(vm.runInContext('calls.length',c),0);
 assert.equal(vm.runInContext("draftValue('template').messageTemplate",c),'pending');
 assert.match(vm.runInContext('messages[0][0]',c),/未保存/);
});

test('step save sends only selected section and preserves unrelated draft',async()=>{
 const c=interactiveModel();
 vm.runInContext("putDraft('template',{messageTemplate:'pending'});putDraft('sources',{formSources:[],formSheetNamePrefix:['回答'],reminderTo:[]});",c);
 await vm.runInContext("saveSection('sources')",c);
 assert.equal(vm.runInContext("calls[0].method",c),'saveSetupSection');
 assert.equal(vm.runInContext("calls[0].args[0]",c),'sources');
 assert.equal(vm.runInContext("'messageTemplate' in calls[0].args[1]",c),false);
 assert.equal(vm.runInContext("draftValue('template').messageTemplate",c),'pending');
 assert.equal(vm.runInContext("draftRevision('template')",c),'r2');
 assert.equal(vm.runInContext("Boolean(state.drafts.sources)",c),false);
});

test('failed save retains draft, unlocks controls, and reports error',async()=>{
 const c=interactiveModel();
 vm.runInContext("putDraft('template',{messageTemplate:'pending'});rpc=async()=>{throw new Error('競合です');};",c);
 await vm.runInContext("saveSection('template')",c);
 assert.equal(vm.runInContext("draftValue('template').messageTemplate",c),'pending');
 assert.equal(vm.runInContext('state.busy',c),false);
 assert.match(vm.runInContext('messages[0][0]',c),/競合/);
});

test('refresh only reads state, retains draft, never activates automation',async()=>{
 const c=interactiveModel();
 vm.runInContext("putDraft('template',{messageTemplate:'pending'});",c);
 await vm.runInContext('refresh(true)',c);
 assert.deepEqual(JSON.parse(vm.runInContext('JSON.stringify(calls.map(c=>c.method))',c)),['getAdminConsoleData','getAutomationState']);
 assert.equal(vm.runInContext("draftValue('template').messageTemplate",c),'pending');
});

test('stop works with unfinished setup and unrelated edits and explicitly includes legacy',async()=>{
 const c=interactiveModel();
 vm.runInContext("state.data.progress=[];state.automation.enabled=true;state.automation.legacyCount=2;putDraft('template',{messageTemplate:'pending'});",c);
 await vm.runInContext('configureSchedule(true)',c);
 assert.deepEqual(JSON.parse(vm.runInContext('JSON.stringify(calls[0])',c)),{method:'stopAutomation',args:['r1',true]});
 assert.equal(vm.runInContext("draftValue('template').messageTemplate",c),'pending');
});

for(const panel of ['manual','settings','automation'])test('refresh skips cached form details outside their visible panels: '+panel,async()=>{
 const c=interactiveModel();c.testPanel=panel;
 vm.runInContext(`state.panel=testPanel==='automation'?'setup':testPanel;state.step='automation';
 state.data.progress.find(p=>p.id==='forms').state='pending';
 acceptFormConsoleData({records:[],setupProgress:{forms:{id:'forms',state:'complete'}}});state.formsError='古い取得エラー';`,c);
 await vm.runInContext('refresh(false)',c);
 assert.deepEqual(JSON.parse(vm.runInContext('JSON.stringify(calls.map(c=>c.method))',c)),['getAdminConsoleData','getAutomationState']);
 assert.equal(vm.runInContext("visibleProgress('forms').state",c),'pending');
 assert.equal(vm.runInContext('state.formsStale',c),true);
});

for(const step of ['forms','publish'])test('refresh re-fetches cached form details when their step is visible: '+step,async()=>{
 const c=interactiveModel();c.testStep=step;
 vm.runInContext(`state.panel='setup';state.step=testStep;renderUpdatedForms=()=>{};
 acceptFormConsoleData({records:[]});rpc=async(method,...args)=>{calls.push({method,args});return method==='getAdminConsoleData'?state.data:method==='getAutomationState'?state.automation:{records:[],setupProgress:{forms:{state:'complete'}}};};`,c);
 await vm.runInContext('refresh(false)',c);
 assert.deepEqual(JSON.parse(vm.runInContext('JSON.stringify(calls.map(c=>c.method))',c)),['getAdminConsoleData','getAutomationState','getFormConsoleData']);
 assert.equal(vm.runInContext('state.formsStale',c),false);
});

for(const failure of ['getAdminConsoleData','getAutomationState'])test('confirmed stop is retained when subsequent refresh fails: '+failure,async()=>{
 const c=interactiveModel();c.failedMethod=failure;
 vm.runInContext(`state.automation.enabled=true;state.automation.revision='auto-before';
 putDraft('template',{messageTemplate:'入力保持'});putDraft('automation',{importHour:4,deliveryHour:8,reminderHour:16,reminderEnabled:false});
 rpc=async(method,...args)=>{calls.push({method,args});if(method==='stopAutomation')return {...state.automation,enabled:false,revision:'auto-stopped',triggers:[]};if(method===failedMethod)throw Error('読込失敗');return method==='getAdminConsoleData'?state.data:state.automation;};`,c);
 await vm.runInContext('configureSchedule(true)',c);
 assert.equal(vm.runInContext('state.automation.enabled',c),false);
 assert.equal(vm.runInContext('state.automation.revision',c),'auto-stopped');
 assert.equal(vm.runInContext("draftRevision('automation')",c),'auto-stopped');
 assert.equal(vm.runInContext("draftValue('template').messageTemplate",c),'入力保持');
 assert.match(vm.runInContext('messages.at(-1)[0]',c),/自動実行を停止しました。[\s\S]*停止後の画面更新に失敗/);
 assert.doesNotMatch(vm.runInContext('messages.at(-1)[0]',c),/再実行してください/);
 assert.equal(vm.runInContext("calls.filter(c=>c.method==='stopAutomation').length",c),1);
 assert.equal(vm.runInContext('state.busy',c),false);
});

test('failed stop never reports success or starts a general refresh',async()=>{
 const c=interactiveModel();
 vm.runInContext("state.automation.enabled=true;rpc=async(method,...args)=>{calls.push({method,args});throw Error('停止の結果を取得できません');};",c);
 await vm.runInContext('configureSchedule(true)',c);
 assert.equal(vm.runInContext('state.automation.enabled',c),true);
 assert.doesNotMatch(vm.runInContext('JSON.stringify(messages)',c),/停止しました/);
 assert.equal(vm.runInContext('calls.length',c),1);
});

test('stop result is shown before refresh finishes and remaining legacy triggers stay visible',async()=>{
 const c=interactiveModel();
 vm.runInContext(`state.automation.enabled=true;state.automation.legacyCount=1;let finishRefresh;
 rpc=async()=>({...state.automation,enabled:false,revision:'stopped',warnings:['旧トリガーの削除失敗']});
 refresh=()=>new Promise(resolve=>{finishRefresh=resolve;});`,c);
 const pending=vm.runInContext('configureSchedule(true)',c);
 await new Promise(resolve=>setImmediate(resolve));
 assert.equal(vm.runInContext('state.automation.enabled',c),false);
 assert.match(vm.runInContext('messages.at(-1)[0]',c),/以前の処理のトリガーが残っています/);
 vm.runInContext('finishRefresh()',c);await pending;
 assert.equal(vm.runInContext('messages.at(-1)[1]',c),'error');
});

test('automation start needs confirmation and carries editable hours',async()=>{
 const c=interactiveModel();
 vm.runInContext("putDraft('automation',{importHour:3,deliveryHour:7,reminderHour:18,reminderEnabled:true});state.automation.legacyCount=1;confirmAction=async()=>false;",c);
 await vm.runInContext('configureSchedule(false)',c);
 assert.equal(vm.runInContext('calls.length',c),0);
 vm.runInContext('confirmAction=async()=>true;',c);
 await vm.runInContext('configureSchedule(false)',c);
 assert.deepEqual(JSON.parse(vm.runInContext('JSON.stringify(calls[0])',c)),{method:'configureAutomation',args:[{importHour:3,deliveryHour:7,reminderHour:18,reminderEnabled:true},'r1',true,'r1']});
});

test('file inspection is read only and apply cannot discard pending edits',async()=>{
 const c=interactiveModel();
 c.file={size:20,text:async()=>'{"version":1}'};
 await vm.runInContext('inspectFile(file)',c);
 assert.equal(vm.runInContext('calls[0].method',c),'inspectSettingsProfile');
 vm.runInContext("putDraft('template',{messageTemplate:'pending'});",c);
 await vm.runInContext('applyProfile()',c);
 assert.equal(vm.runInContext('calls.length',c),1);
 assert.equal(vm.runInContext("draftValue('template').messageTemplate",c),'pending');
});

test('local file apply is explicit and uses inspection revision',async()=>{
 const c=interactiveModel();
 vm.runInContext("state.profile={json:'{}',inspection:{summary:'確認'},revision:'r1'};confirmAction=async()=>false;",c);
 await vm.runInContext('applyProfile()',c);
 assert.equal(vm.runInContext('calls.length',c),0);
 vm.runInContext('confirmAction=async()=>true;',c);
 await vm.runInContext('applyProfile()',c);
 assert.deepEqual(JSON.parse(vm.runInContext('JSON.stringify(calls[0])',c)),{method:'applySettingsProfile',args:['{}','r1']});
});

test('server-provided strings and template payload are escaped in markup',()=>{
 const c=interactiveModel();
 vm.runInContext("state.data.config.messageTemplate='<img onerror=alert(1)>';state.data.config.fields=[{key:'feedback',sourceHeader:'<script>',type:'text'}];",c);
 const markup=vm.runInContext('renderTemplate()',c);
 assert.doesNotMatch(markup,/<img onerror|<script>/);
 assert.match(markup,/&lt;img/);
});

test('automation and configuration use independent concurrency revisions',async()=>{
 const c=interactiveModel();
 vm.runInContext("state.automation.revision='auto-7';putDraft('automation',{importHour:4,deliveryHour:8,reminderHour:16,reminderEnabled:false});putDraft('template',{messageTemplate:'pending'});rebaseDrafts('r1','r2');",c);
 assert.equal(vm.runInContext("draftRevision('automation')",c),'auto-7');
 assert.equal(vm.runInContext("draftRevision('template')",c),'r2');
 vm.runInContext('delete state.drafts.template;',c);
 await vm.runInContext('configureSchedule(false)',c);
 assert.equal(vm.runInContext('calls[0].args[1]',c),'auto-7');
 assert.equal(vm.runInContext('calls[0].args[3]',c),'r1');
});

test('same-lock state endpoints are never requested concurrently',async()=>{
 const c=interactiveModel();
 vm.runInContext("let requestsInFlight=0,maxInFlight=0;rpc=async method=>{requestsInFlight++;maxInFlight=Math.max(maxInFlight,requestsInFlight);await Promise.resolve();requestsInFlight--;return method==='getAdminConsoleData'?state.data:state.automation;};",c);
 await vm.runInContext('refresh(false)',c);
 assert.equal(vm.runInContext('maxInFlight',c),1);
});

test('editing notification recipients preserves legacy per-source course assignment',()=>{
 const c=model();
 vm.runInContext("state.data={config:{formSources:[{id:'123',url:'https://docs.google.com/spreadsheets/d/123/edit',courseId:'course_456'}],reminderTo:['old@example.test'],formSheetNamePrefix:['回答']},revision:'r1'};const sourceDraft=draftValue('sources');sourceDraft.formSources=sourceDraft.formSources.map(s=>parseSourceLine(formatSourceLine(s)));sourceDraft.reminderTo=['new@example.test'];putDraft('sources',sourceDraft);",c);
 assert.equal(vm.runInContext("draftValue('sources').formSources[0].courseId",c),'course_456');
 assert.equal(vm.runInContext("draftValue('sources').formSources[0].url",c),'https://docs.google.com/spreadsheets/d/123/edit');
 assert.match(vm.runInContext("renderStep('sources')",c),/edit \| course_456/);
});

test('grade conversion draft preserves empty and custom scales, with legacy migration fallback',()=>{
 const c=interactiveModel();
 vm.runInContext('state.data.config.gradeScale=[];',c);
 assert.equal(vm.runInContext("draftValue('fields').gradeScale.length",c),0);
 vm.runInContext("state.data.config.gradeScale=[{from:'合格',to:'よくできました'}];",c);
 assert.equal(vm.runInContext("draftValue('fields').gradeScale[0].to",c),'よくできました');
 vm.runInContext('delete state.data.config.gradeScale;',c);
 assert.equal(vm.runInContext("draftValue('fields').gradeScale.length",c),5);
});

test('grade scale is saved with the fields section and used in token examples',async()=>{
 const c=interactiveModel();
 vm.runInContext("state.data.config.gradeScale=[];const fields=draftValue('fields');fields.gradeScale=[{from:'1',to:'再提出'}];fields.fields=[{key:'grade',sourceHeader:'結果',type:'score_grade'}];putDraft('fields',fields);",c);
 assert.equal(vm.runInContext("tokenItems().find(t=>t.key==='grade').example",c),'再提出');
 await vm.runInContext("saveSection('fields')",c);
 assert.equal(vm.runInContext('calls[0].args[1].gradeScale[0].to',c),'再提出');
});

test('diagnosis requires saved settings and makes only the read-only diagnostic call',async()=>{
 const c=interactiveModel();
 vm.runInContext("putDraft('template',{messageTemplate:'draft'});",c);
 await vm.runInContext('runDiagnostics()',c);
 assert.equal(vm.runInContext('calls.length',c),0);
 vm.runInContext("delete state.drafts.template;rpc=async(method,...args)=>{calls.push({method,args});return {revision:'r1',status:'attention',counts:{eligible:3},checks:[],samples:[]};};",c);
 await vm.runInContext('runDiagnostics()',c);
 assert.deepEqual(JSON.parse(vm.runInContext('JSON.stringify(calls)',c)),[{method:'diagnoseSetup',args:['r1']}]);
 assert.equal(vm.runInContext('state.diagnostic.status',c),'attention');
});

test('diagnostic failures discard previous findings and leave controls unlocked',async()=>{
 const c=interactiveModel();
 vm.runInContext("state.diagnostic={status:'ok'};rpc=async()=>{throw new Error('参照元にアクセスできません');};",c);
 await vm.runInContext('runDiagnostics()',c);
 assert.equal(vm.runInContext('state.diagnostic',c),null);
 assert.equal(vm.runInContext('state.busy',c),false);
 assert.match(vm.runInContext('messages[0][0]',c),/アクセス/);
});

test('diagnostic samples identify actual data and stage, escape content, and show limits',()=>{
 const c=interactiveModel();
 vm.runInContext("state.diagnostic={status:'attention',checkedAt:'2026-09-29',counts:{sources:2,sheets:3,eligible:4,previewed:1,blocked:3,skipped:2},checks:[{severity:'error',title:'要確認',detail:'宛先が未設定',action:'名簿を確認'}],samples:[{stage:'送信待ち',sourceSheet:'回答',sourceRow:3,email:'test@example.test',name:'例',className:'A組',courseId:'course1',body:'<img src=x onerror=alert(1)>'}],truncated:true};",c);
 const markup=vm.runInContext('renderDiagnostics()',c);
 assert.match(markup,/実データ/);assert.match(markup,/送信待ち/);assert.match(markup,/&lt;img/);assert.doesNotMatch(markup,/<img src/);
 assert.match(markup,/上限/);assert.match(markup,/投稿権限/);assert.match(markup,/名簿を確認/);
});

test('refresh and changed settings invalidate diagnostic results',async()=>{
 const c=interactiveModel();
 vm.runInContext("state.diagnostic={status:'ok'};putDraft('template',{messageTemplate:'changed'});",c);
 assert.equal(vm.runInContext('state.diagnostic',c),null);
 vm.runInContext("state.diagnostic={status:'ok'};",c);
 await vm.runInContext('refresh(false)',c);
 assert.equal(vm.runInContext('state.diagnostic',c),null);
});

test('health guidance has only allowlisted sheet links and no automatic recovery action',()=>{
 const c=interactiveModel();
 vm.runInContext("state.data.sheets=[{key:'send',exists:true,url:'https://docs.google.com/spreadsheets/d/test/edit#gid=3'}];state.data.health={status:'attention',items:[{title:'送信確認待ち',count:2,detail:'結果を確認',action:'送信先と実行結果を確認する',sheetKey:'send'},{title:'不正',detail:'<script>',action:'確認',sheetKey:'arbitrary'}]};",c);
 const markup=vm.runInContext('renderHealth()',c);
 assert.match(markup,/href="https:\/\/docs.google.com\/spreadsheets\/d\/test\/edit#gid=3"/);assert.doesNotMatch(markup,/data-sheet="arbitrary"|data-action=|<script>/);
 assert.match(markup,/送信先と実行結果を確認する/);
});

test('last runs label attention and display elapsed time',()=>{
 const c=interactiveModel();
 const markup=vm.runInContext("renderLastRuns([{kind:'delivery',status:'attention',elapsedMs:1250,message:'一部を保留'}])",c);
 assert.match(markup,/要確認/);assert.match(markup,/1.3秒/);assert.match(markup,/一部を保留/);
});

test('conversion inputs are collected with field roles and survive navigation drafts',()=>{
 const c=model();
 const values={emailHeader:'メール',studentNameHeader:'名前',formStatusHeader:'状態',scoreSourceHeader:'点',replyBodyHeader:'本文',draftStatus:''};
 const nodes=Object.fromEntries(Object.entries(values).map(([key,value])=>[key,{value,textContent:''}]));
 const rows=[{querySelector:selector=>({value:selector.includes('from')?' 5 ':' 秀 '})}];
 c.document={getElementById:id=>nodes[id],querySelectorAll:selector=>selector==='[data-grade-index]'?rows:[]};
 vm.runInContext("state.data={config:{fields:[],gradeScale:[]},revision:'r1'};state.step='fields';readCurrent();state.step='template';",c);
 assert.deepEqual(JSON.parse(vm.runInContext("JSON.stringify(draftValue('fields').gradeScale)",c)),[{from:'5',to:'秀'}]);
 assert.equal(vm.runInContext("draftValue('fields').emailHeader",c),'メール');
 assert.equal(vm.runInContext("draftRevision('fields')",c),'r1');
});

test('fictional template preview carries unsaved conversion rules to server renderer',async()=>{
 const c=interactiveModel();
 vm.runInContext("state.data.config.gradeScale=[];const p=draftValue('fields');p.gradeScale=[{from:'PASS',to:'合格'}];putDraft('fields',p);",c);
 await vm.runInContext("command('preview')",c);
 assert.equal(vm.runInContext('calls[0].method',c),'previewReplyTemplate');
 assert.equal(vm.runInContext('calls[0].args[0].gradeScale[0].to',c),'合格');
});

test('diagnostic and recovery cards are reachable from their intended panels',()=>{
 const c=interactiveModel();
 vm.runInContext("state.data.health={status:'ok',items:[]};",c);
 assert.match(vm.runInContext("renderStep('automation')",c),/data-command="diagnose"/);
 assert.match(vm.runInContext('renderManual()',c),/data-command="diagnose"/);
 assert.match(vm.runInContext('renderMaintenance()',c),/healthTitle/);
 assert.match(vm.runInContext('renderFields()',c),/gradeScaleRows/);
});

test('removing the final conversion rule preserves an explicitly empty scale',async()=>{
 const c=interactiveModel();
 vm.runInContext("state.data.config.gradeScale=[{from:'5',to:'A'}];const removeButton={disabled:false,dataset:{removeGrade:'0'}};const event={target:{closest:()=>removeButton}};",c);
 await vm.runInContext('handleClick(event)',c);
 assert.deepEqual(JSON.parse(vm.runInContext("JSON.stringify(draftValue('fields').gradeScale)",c)),[]);
 assert.match(vm.runInContext('renderGradeScale()',c),/変換表は未設定です/);
});

test('stale diagnostic response cannot be displayed as a current result',async()=>{
 const c=interactiveModel();
 vm.runInContext("rpc=async()=>({revision:'old',status:'ok',samples:[]});",c);
 await vm.runInContext('runDiagnostics()',c);
 assert.equal(vm.runInContext('state.diagnostic',c),null);
 assert.match(vm.runInContext('messages[0][0]',c),/設定が変更/);
});

test('step and panel navigation starts content at the top while preserving drafts',async()=>{
 for(const dataset of [{step:'sources'},{panel:'maintenance'}]){
  const c=interactiveModel(),content={scrollTop:420,focus(options){this.focusOptions=options;}};
  vm.runInContext("renderUpdatedForms=()=>{};",c);
  c.document={getElementById:id=>{assert.equal(id,'content');return content;}};
  c.navigationEvent={target:{closest:()=>({disabled:false,dataset})}};
  vm.runInContext("putDraft('template',{messageTemplate:'編集中'});",c);
  await vm.runInContext('handleClick(navigationEvent)',c);
  assert.equal(content.scrollTop,0);
  assert.equal(content.focusOptions.preventScroll,true);
  assert.equal(vm.runInContext("draftValue('template').messageTemplate",c),'編集中');
  assert.deepEqual(JSON.parse(vm.runInContext('JSON.stringify(calls.map(c=>c.method))',c)),dataset.panel==='maintenance'?['getFormConsoleData']:[]);
 }
});


test('roster and later source mapping have independent progress while both gate automation',()=>{
 const c=interactiveModel();
 vm.runInContext("state.data.mappingNeeded=true;state.data.progress.find(p=>p.id==='students').state='pending';",c);
 assert.equal(vm.runInContext("visibleProgress('classes').state",c),'pending');
 assert.equal(vm.runInContext('automationReady()',c),false);
 vm.runInContext("state.data.progress.find(p=>p.id==='students').state='complete';state.data.progress.find(p=>p.id==='mapping').state='attention';",c);
 assert.equal(vm.runInContext("visibleProgress('classes').state",c),'complete');
 assert.equal(vm.runInContext("visibleProgress('sources').state",c),'attention');
 assert.equal(vm.runInContext('automationReady()',c),false);
});

test('mapping follows form creation and source registration rather than roster retrieval',()=>{
 const c=interactiveModel();
 let markup=vm.runInContext("renderStep('classes')",c);
 assert.match(markup,/data-action="students"/);
 assert.doesNotMatch(markup,/回答シートとクラスの対応表/);
 markup=vm.runInContext("renderStep('sources')",c);
 assert.match(markup,/data-command="mapping-sources"/);
 assert.doesNotMatch(markup,/data-action="mapping"/);
 vm.runInContext("state.data.config.formSources=[{url:'test'}];state.data.config.formSheetNamePrefix=['回答'];",c);
 markup=vm.runInContext("renderStep('sources')",c);
 assert.match(markup,/data-action="mapping"/);
 assert.match(markup,/<details class="card"><summary>回答シート/);
 vm.runInContext('state.data.mappingNeeded=true;',c);
 assert.match(vm.runInContext("renderStep('sources')",c),/<details class="card" open>/);
});

test('form preparation and Classroom publishing render distinct actions and retain an existing-source route',()=>{
 const c=interactiveModel();
 vm.runInContext("state.forms={classes:[],defaults:{},records:[{id:'f1',kind:'form',stage:'registered',label:'A',materialTitle:'資料'},{id:'f2',kind:'form',stage:'pending',label:'B'}]};",c);
 const forms=vm.runInContext("renderStep('forms')",c),publish=vm.runInContext("renderStep('publish')",c);
 assert.match(forms,/data-form-command="create-template"/);assert.match(forms,/data-form-command="prepare"/);
 assert.doesNotMatch(forms,/data-form-command="publish"/);assert.match(forms,/既存のフォーム/);
 assert.match(publish,/data-form-command="publish"/);assert.doesNotMatch(publish,/id="form-templateId"/);
 assert.match(publish,/data-step="forms"/);
});

test('form drafts survive leaving the integrated step and publishing navigation only reads',async()=>{
 const c=backgroundModel();
 vm.runInContext("state.panel='setup';state.step='forms';readCurrent=()=>{if(state.step==='forms')state.formDraft={templateId:'keep',folderId:'folder',description:'draft'};};",c);
 await vm.runInContext("handleClick({target:{closest:()=>({dataset:{step:'publish'}})}})",c);
 assert.equal(vm.runInContext('state.formDraft.description',c),'draft');
 assert.equal(vm.runInContext('state.step',c),'publish');
 const pending=vm.runInContext('state.formsLoadingPromise',c);vm.runInContext('finish({classes:[],defaults:{},records:[]})',c);await pending;
 assert.equal(vm.runInContext('state.formDraft.templateId',c),'keep');
});

test('late return history does not collapse the open period-copy section',()=>{
 const c=interactiveModel();
 const after=[{id:'formReturnDetails',open:false},{id:'operationCopyDetails',open:false}];
 const content={scrollTop:12,querySelectorAll:()=>[{id:'operationCopyDetails',open:true}]};
 c.document={getElementById:id=>id==='content'?content:null,activeElement:null};c.after=after;
 vm.runInContext("state.panel='maintenance';readFormDraft=()=>{};render=()=>{document.getElementById('content').querySelectorAll=()=>after;};renderUpdatedForms();",c);
 assert.equal(after[0].open,false);assert.equal(after[1].open,true);
});

test('opening directly on an unfinished form step loads its data and never posts',async()=>{
 const c=interactiveModel();
 vm.runInContext(`state.data.progress.find(p=>p.id==='forms').state='pending';renderUpdatedForms=()=>{};
 rpc=async(method)=>{calls.push({method});if(method==='getAdminConsoleData')return state.data;if(method==='getAutomationState')return state.automation;if(method==='getFormConsoleData')return {records:[],classes:[],defaults:{}};throw Error('unexpected mutation');};`,c);
 await vm.runInContext('refresh(true)',c);
 assert.equal(vm.runInContext('state.step',c),'forms');
 assert.equal(vm.runInContext('state.formsLoading',c),false);
 assert.deepEqual(JSON.parse(vm.runInContext('JSON.stringify(calls.map(c=>c.method))',c)),['getAdminConsoleData','getAutomationState','getFormConsoleData']);
});

test('management keeps return history and period copying outside form preparation',()=>{
 const c=interactiveModel();
 vm.runInContext("state.bootstrap={canCopyOperation:true};state.forms={classes:[],defaults:{},records:[{id:'r',kind:'announcement',stage:'published',postId:'p',title:'返却'}]};",c);
 const management=vm.runInContext('renderMaintenance()',c),forms=vm.runInContext("renderStep('forms')",c);
 assert.match(management,/個別返却の投稿を管理/);assert.match(management,/次の期間用に複製/);assert.match(management,/data-form-command="delete"/);
 assert.doesNotMatch(forms,/次の期間用に複製|個別返却の履歴|data-form-command="delete"/);
});

test('sources use two-row URL entries and single-line email entries',()=>{
 const c=interactiveModel();
 vm.runInContext("state.data.config.formSources=[{url:'first'},{url:'second'}];state.data.config.reminderTo=['one@example.test','two@example.test'];",c);
 const markup=vm.runInContext("renderStep('sources')",c);
 assert.equal((markup.match(/textarea rows="2" data-source-entry="formSources"/g)||[]).length,2);
 assert.equal((markup.match(/input type="email" data-source-entry="reminderTo"/g)||[]).length,2);
 assert.match(markup,/data-add-entry="formSources"/);
 assert.match(markup,/formSheetNamePrefix/);
});

test('top settings panel owns export and import; management keeps only operational controls',()=>{
 const c=interactiveModel();
 assert.match(html,/<div class="header-tools">[\s\S]*?<button type="button" data-panel="settings">設定を保存・出力/);
 const markup=vm.runInContext('renderSettings()',c);
 assert.match(markup,/data-command="save-all"/);assert.match(markup,/data-command="export"/);assert.match(markup,/id="profileFile"/);
 assert.doesNotMatch(vm.runInContext('renderMaintenance()',c),/profileFile|data-command="export"|MAINTENANCE/);
});

test('save all submits configuration atomically and retains independent schedule draft',async()=>{
 const c=interactiveModel();
 vm.runInContext("putDraft('template',{messageTemplate:'new'});putDraft('sources',{formSources:[],formSheetNamePrefix:['回答'],reminderTo:[]});putDraft('automation',{importHour:4,deliveryHour:8,reminderHour:16,reminderEnabled:false});",c);
 await vm.runInContext('saveAllSettings()',c);
 const call=JSON.parse(vm.runInContext('JSON.stringify(calls[0])',c));
 assert.equal(call.method,'saveSetupSections');assert.equal(call.args[1],'r1');assert.deepEqual(Object.keys(call.args[0]).sort(),['sources','template']);
 assert.equal(vm.runInContext('Boolean(state.drafts.template)',c),false);
 assert.equal(vm.runInContext("draftValue('automation').importHour",c),4);
 assert.equal(vm.runInContext("calls.some(c=>c.method==='configureAutomation')",c),false);
 assert.match(vm.runInContext('messages.at(-1)[0]',c),/時間帯の入力は保持/);
});

test('save all rejects mixed revisions and keeps every draft on server failure',async()=>{
 const c=interactiveModel();
 vm.runInContext("putDraft('template',{messageTemplate:'new'});state.drafts.template.revision='old';",c);
 await vm.runInContext('saveAllSettings()',c);
 assert.equal(vm.runInContext('calls.length',c),0);
 assert.equal(vm.runInContext("draftValue('template').messageTemplate",c),'new');
 vm.runInContext("state.drafts.template.revision='r1';rpc=async()=>{throw new Error('保存失敗');};",c);
 await vm.runInContext('saveAllSettings()',c);
 assert.equal(vm.runInContext("draftValue('template').messageTemplate",c),'new');
 assert.match(vm.runInContext('messages.at(-1)[0]',c),/保存失敗/);
});

test('per-field conversion examples override common scale and remain explicit when empty',()=>{
 const c=interactiveModel();
 vm.runInContext("state.data.config.gradeScale=[{from:'1',to:'C'}];state.data.config.fields=[{key:'symbols',sourceHeader:'評価',type:'score_grade',gradeScale:[{from:'1',to:'×'}]},{key:'letters',sourceHeader:'点',type:'score_grade'},{key:'blank',sourceHeader:'別',type:'score_grade',gradeScale:[]}];",c);
 assert.deepEqual(JSON.parse(vm.runInContext("JSON.stringify(tokenItems().filter(t=>['symbols','letters','blank'].includes(t.key)).map(t=>t.example))",c)),['×','C','元の評価値']);
 const markup=vm.runInContext('renderFields()',c);
 assert.match(markup,/1→×/);assert.match(markup,/1→C/);assert.match(markup,/data-scale-mode="0"/);
});

test('removing last field conversion never falls back to common mapping',async()=>{
 const c=interactiveModel();
 vm.runInContext("state.data.config.gradeScale=[{from:'1',to:'C'}];state.data.config.fields=[{key:'grade',sourceHeader:'評価',type:'score_grade',gradeScale:[{from:'1',to:'×'}]}];const event={target:{closest:()=>({disabled:false,dataset:{removeFieldGrade:'0',gradeRow:'0'}})}};",c);
 await vm.runInContext('handleClick(event)',c);
 assert.deepEqual(JSON.parse(vm.runInContext("JSON.stringify(draftValue('fields').fields[0].gradeScale)",c)),[]);
 assert.equal(vm.runInContext("gradeExample(draftValue('fields').fields[0])",c),'元の評価値');
});

test('compact tokens offer keyboard-accessible explanation buttons',()=>{
 const c=interactiveModel();
 const markup=vm.runInContext('renderTemplate()',c);
 assert.match(markup,/data-help-title="生徒の名前"[^>]*aria-label="生徒の名前の説明"/);
 assert.match(markup,/<strong>\{student_name\}<\/strong>/);
 assert.match(html,/<div id="contextTooltip"[^>]*role="tooltip"/);
 assert.doesNotMatch(markup,/<small>.*を挿入/);
});

test('form defaults distinguish short labels and explain domain syntax',()=>{
 const c=model();
 vm.runInContext('state.forms={classes:[],defaults:{},records:[]}',c);
 assert.equal(vm.runInContext('defaultFormDraft().titlePattern',c),'{class_label} 授業の振り返り');
 const markup=vm.runInContext('renderForms()',c);
 assert.match(markup,/\{class_label\}/);assert.match(markup,/@なし/);assert.match(markup,/school\.example/);
});

test('form columns start empty and class guidance disables preparation until selection',()=>{
 const c=model();vm.runInContext("state.forms={classes:[],defaults:{scoreSourceHeader:'総合'},records:[],classWarning:'クラスを選択してください'}",c);
 assert.equal(vm.runInContext('defaultFormDraft().additionalColumns.length',c),0);
 const markup=vm.runInContext('renderForms()',c);
 assert.match(markup,/クラスを選択してください/);assert.match(markup,/data-step="classes"/);
 assert.match(markup,/data-form-command="preview"[^>]*disabled/);
 assert.match(markup,/総合/);assert.match(markup,/data-step="fields"/);
 assert.doesNotMatch(markup,/id="form-gradeHeader"|id="form-commentHeader"|id="form-gradeChoices"/);
});

test('custom columns render escaped names and individual optional choice controls',()=>{
 const c=model();vm.runInContext("state.forms={classes:[{courseId:'1',className:'化学'}],defaults:{},records:[]};state.formDraft=defaultFormDraft();state.formDraft.additionalColumns=[{header:'<観点>',choices:'1,2,3'}]",c);
 const markup=vm.runInContext('renderForms()',c);
 assert.match(markup,/&lt;観点&gt;/);assert.match(markup,/data-extra-header/);assert.match(markup,/data-extra-choices/);
 assert.match(markup,/data-form-command="add-column"/);assert.match(markup,/data-form-command="remove-column"/);
 assert.match(markup,/data-class-label="1"/);assert.doesNotMatch(markup,/data-form-command="preview"[^>]*disabled/);
});

test('add and remove custom columns preserve current inputs without invoking a server mutation',async()=>{
 const c=interactiveModel();
 vm.runInContext("state.forms={defaults:{},classes:[],records:[]};state.formDraft=defaultFormDraft();state.formDraft.additionalColumns=[{header:'理解',choices:'1,2'},{header:'表現',choices:''}];readFormDraft=()=>{};const focused=[];document={getElementById:id=>({focus:()=>focused.push(id)})}",c);
 await vm.runInContext("formCommand('add-column')",c);
 assert.equal(vm.runInContext('state.formDraft.additionalColumns.length',c),3);
 assert.equal(vm.runInContext('focused[0]',c),'extra-header-2');
 await vm.runInContext("formCommand('remove-column','1')",c);
 assert.deepEqual(JSON.parse(vm.runInContext('JSON.stringify(state.formDraft.additionalColumns)',c)),[{header:'理解',choices:'1,2'},{header:'',choices:''}]);
 assert.equal(vm.runInContext('calls.length',c),0);
});

test('form draft collects optional columns and labels before navigation',()=>{
 const c=model();
 const inputs={'form-templateId':{value:'template'},'form-nameHeader':{value:'名前'}};
 c.document={getElementById:id=>inputs[id],querySelectorAll:selector=>selector==='[data-extra-column]'?[{querySelector:sel=>({value:sel==='[data-extra-header]'?'観点':'A,B'})}]:selector==='[data-class-label]'?[{dataset:{classLabel:'1'},value:'1-1'}]:[]};
 vm.runInContext("state.forms={defaults:{}};state.panel='setup';state.step='forms';readCurrent()",c);
 assert.deepEqual(JSON.parse(vm.runInContext('JSON.stringify(state.formDraft.additionalColumns)',c)),[{header:'観点',choices:'A,B'}]);
 assert.equal(vm.runInContext("state.formDraft.labels['1']",c),'1-1');
});


test('select input then change events preserve a hidden individual grade conversion',()=>{
 const c=model(),listeners={};
 const values={emailHeader:'メール',studentNameHeader:'名前',formStatusHeader:'状態',scoreSourceHeader:'点',replyBodyHeader:'本文',draftStatus:''};
 const nodes=Object.fromEntries(Object.entries(values).map(([key,value])=>[key,{value,textContent:''}]));
 nodes.confirmCheck={addEventListener(){}};
 const row={dataset:{fieldIndex:'0',customGradeRendered:'false'},querySelectorAll:selector=>selector==='[data-field-prop]'?Object.entries({key:'score',sourceHeader:'点',evalHeader:'点',sendHeader:'点',type:'score_grade'}).map(([fieldProp,value])=>({dataset:{fieldProp},value})):[]};
 c.document={getElementById:id=>nodes[id],querySelectorAll:selector=>selector==='[data-field-index]'?[row]:[],addEventListener:(name,handler)=>{listeners[name]=handler;}};
 c.window={addEventListener(){}};
 vm.runInContext("state.data={config:{fields:[{key:'score',sourceHeader:'点',evalHeader:'点',sendHeader:'点',type:'text',gradeScale:[{from:'1',to:'×'}]}],gradeScale:[]},revision:'r1'};state.step='fields';task=()=>{};render=()=>{};boot();",c);
 const event={target:{dataset:{fieldProp:'type'},closest:()=>true}};
 listeners.input(event);
 assert.equal(vm.runInContext("draftValue('fields').fields[0].type",c),'score_grade');
 listeners.change(event);
 assert.deepEqual(JSON.parse(vm.runInContext("JSON.stringify(draftValue('fields').fields[0].gradeScale)",c)),[{from:'1',to:'×'}]);
});

test('visible individual grade rules are collected without leaking into the shared scale',()=>{
 const c=model();
 const values={emailHeader:'メール',studentNameHeader:'名前',formStatusHeader:'状態',scoreSourceHeader:'点',replyBodyHeader:'本文',draftStatus:''};
 const nodes=Object.fromEntries(Object.entries(values).map(([key,value])=>[key,{value,textContent:''}]));
 const rule={querySelector:selector=>({value:selector.includes('from')?' 1 ':' ○ '})};
 const row={dataset:{fieldIndex:'0',customGradeRendered:'true'},querySelectorAll:selector=>selector==='[data-field-prop]'?Object.entries({key:'score',sourceHeader:'点',evalHeader:'点',sendHeader:'点',type:'score_grade'}).map(([fieldProp,value])=>({dataset:{fieldProp},value})):selector==='[data-field-grade]'?[rule]:[]};
 c.document={getElementById:id=>nodes[id],querySelectorAll:selector=>selector==='[data-field-index]'?[row]:[]};
 vm.runInContext("state.data={config:{fields:[{key:'score',sourceHeader:'点',type:'score_grade',gradeScale:[{from:'1',to:'×'}]}],gradeScale:[]},revision:'r1'};state.step='fields';readCurrent();",c);
 assert.deepEqual(JSON.parse(vm.runInContext("JSON.stringify(draftValue('fields').fields[0].gradeScale)",c)),[{from:'1',to:'○'}]);
 assert.deepEqual(JSON.parse(vm.runInContext("JSON.stringify(draftValue('fields').gradeScale)",c)),[]);
});


test('adding a URL entry renders a blank input and never serializes its object',async()=>{
 const c=interactiveModel();
 c.document={querySelectorAll:()=>[{focus(){}}]};
 vm.runInContext("state.data.config.formSources=[{url:'https://docs.google.com/spreadsheets/d/123/edit'}];const event={target:{closest:()=>({disabled:false,dataset:{addEntry:'formSources'}})}};",c);
 await vm.runInContext('handleClick(event)',c);
 assert.equal(vm.runInContext("draftValue('sources').formSources.length",c),2);
 assert.equal(vm.runInContext("formatSourceLine(draftValue('sources').formSources[1])",c),'');
 const markup=vm.runInContext("renderStep('sources')",c);
 assert.doesNotMatch(markup,/\[object Object\]/);
 assert.match(markup,/aria-label="URL2"><\/textarea>/);
 assert.equal(vm.runInContext("formatSourceLine('legacy-source')",c),'legacy-source');
 assert.equal(vm.runInContext("formatSourceLine({id:'123',courseId:'456'})",c),'123 | 456');
});

test('closing the web console preserves unsaved-input confirmation and only guides closing the tab',async()=>{
 const c=model();vm.runInContext("readCurrent=()=>{};dirtySections=()=>['fields'];let messages=[];status=m=>messages.push(m);confirmAction=async()=>false;",c);
 await vm.runInContext("command('close')",c);assert.equal(vm.runInContext('messages.length',c),0);
 vm.runInContext('confirmAction=async()=>true',c);await vm.runInContext("command('close')",c);
 assert.match(vm.runInContext('messages[0]',c),/タブを閉じ/);
});

test('maintenance always links to the connected workbook and handles no connection',()=>{
 const c=model();vm.runInContext("state.bootstrap={spreadsheetId:'book'}",c);
 assert.match(vm.runInContext('maintenanceSheetLinks()',c),/https:\/\/docs.google.com\/spreadsheets\/d\/book\/edit/);
 vm.runInContext('state.bootstrap=null',c);assert.match(vm.runInContext('maintenanceSheetLinks()',c),/未接続/);
});
