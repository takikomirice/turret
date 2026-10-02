import {test} from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import vm from 'node:vm';

const html=readFileSync('Setting.html','utf8'),script=html.match(/<script>([\s\S]*?)<\/script>/)[1];
function model(){
 const c=vm.createContext({console,structuredClone});vm.runInContext(script.replace(/boot\(\);\s*$/,''),c);
 vm.runInContext("state.data={config:{emailHeader:'メール',studentNameHeader:'名前',scoreSourceHeader:'評価',formStatusHeader:'状態',replyBodyHeader:'返信本文',fields:[{key:'score',sourceHeader:'評価',type:'text'}],gradeScale:[]},revision:'r1',counts:{},progress:[]};state.headers=['メール','名前','評価','管理','状態','<観点>'];",c);
 return c;
}

test('roles and message source columns are native selects with real headers and saved selections',()=>{
 const c=model(),markup=vm.runInContext('renderFields()',c);
 for(const id of ['emailHeader','studentNameHeader','scoreSourceHeader'])assert.match(markup,new RegExp('<select id="'+id+'"'));
 assert.match(markup,/<select id="field-0-sourceHeader" data-field-prop="sourceHeader">/);
 assert.match(markup,/<option value="評価" selected>評価<\/option>/);
 assert.match(markup,/&lt;観点&gt;/);assert.doesNotMatch(markup,/<datalist|list="headerOptions"|id="replyBodyHeader"/);
 assert.match(markup,/各列のメッセージへの取り込みルール/);assert.match(markup,/本文の保存列：返信本文（固定）/);
 assert.match(markup,/class="grid field-grid"/);
 const options=vm.runInContext("headerOptions('保存済み列')",c);
 assert.match(options,/value="保存済み列" selected/);assert.doesNotMatch(options,/value="管理"|value="状態"/);
 assert.match(options,/value=""/);
});

test('body header is preserved from saved settings when collecting a field draft',()=>{
 const c=model(),nodes={emailHeader:{value:'メール'},studentNameHeader:{value:'名前'},scoreSourceHeader:{value:'評価'},draftStatus:{}};
 c.document={getElementById:id=>nodes[id],querySelectorAll:()=>[]};
 vm.runInContext("state.data.config.replyBodyHeader='旧本文';state.step='fields';readCurrent()",c);
 assert.equal(vm.runInContext("draftValue('fields').replyBodyHeader",c),'旧本文');
});

test('adding evaluation columns offers names only and submits no input domain',async()=>{
 const c=model();vm.runInContext("state.forms={records:[{id:'r',kind:'form',stage:'registered',responseSheetId:1}],classes:[],defaults:{},configRevision:'r1'};state.columnDraft={recordIds:['r'],additionalColumns:[{header:'評価',choices:'1,2,3,4,5'}]};",c);
 const markup=vm.runInContext('renderResponseColumns()',c);
 assert.doesNotMatch(markup,/data-extra-choices|入力用プルダウン|選択肢・/);
 vm.runInContext("readCurrent=()=>{};readColumnDraft=()=>{};requireSaved=()=>true;task=async f=>f();render=()=>{};status=()=>{};const calls=[];rpc=async(method,...args)=>{calls.push({method,args});return method==='getFormConsoleData'?state.forms:{targets:[],additionalColumns:[{header:'評価',choices:[]}],configRevision:'r1'};};confirmAction=async()=>false;",c);
 await vm.runInContext("formCommand('preview-columns')",c);
 const raw=JSON.parse(vm.runInContext("JSON.stringify(calls.find(c=>c.method==='previewResponseColumns').args[0])",c));
 assert.equal(raw.columnNamesOnly,true);assert.deepEqual(raw.additionalColumns,[{header:'評価'}]);
});

test('header loading is deduplicated, waits for form requests and preserves unsaved selections',async()=>{
 const c=model();vm.runInContext("state.step='fields';readCurrent=()=>{};render=()=>{};let releaseForms,releaseHeaders;const calls=[];state.formsLoadingPromise=new Promise(resolve=>releaseForms=resolve);rpc=method=>{calls.push(method);return new Promise(resolve=>releaseHeaders=resolve);};",c);
 const first=vm.runInContext('refreshHeadersInBackground()',c),second=vm.runInContext('refreshHeadersInBackground()',c);
 assert.equal(first,second);assert.equal(vm.runInContext('calls.length',c),0);
 vm.runInContext('releaseForms()',c);await new Promise(resolve=>setImmediate(resolve));
 assert.equal(vm.runInContext('calls.length',c),1);
 vm.runInContext("const draft=draftValue('fields');draft.scoreSourceHeader='未保存の列';putDraft('fields',draft);releaseHeaders({headers:['メール','評価'],warnings:[]});",c);
 await first;assert.equal(vm.runInContext("draftValue('fields').scoreSourceHeader",c),'未保存の列');
 assert.equal(vm.runInContext('state.headers.length',c),2);
 await vm.runInContext('refreshHeadersInBackground()',c);assert.equal(vm.runInContext('calls.length',c),1);
});

test('stale or failed header responses keep saved selections and allow retry',async()=>{
 const c=model();vm.runInContext("readCurrent=()=>{};render=()=>{};let finish;rpc=()=>new Promise(resolve=>finish=resolve)",c);
 const loading=vm.runInContext('refreshHeadersInBackground()',c);await new Promise(resolve=>setImmediate(resolve));
 vm.runInContext("state.data.revision='r2';finish({headers:['古い結果']})",c);await loading;
 assert.equal(vm.runInContext("state.headers.includes('古い結果')",c),false);
 vm.runInContext("rpc=async()=>{throw Error('permission denied')}",c);await vm.runInContext('refreshHeadersInBackground()',c);
 assert.match(vm.runInContext('state.headersError',c),/permission denied/);
 vm.runInContext("rpc=async()=>({headers:['再取得した列']})",c);await vm.runInContext('refreshHeadersInBackground()',c);
 assert.deepEqual(JSON.parse(vm.runInContext('JSON.stringify(state.headers)',c)),['再取得した列']);
});
