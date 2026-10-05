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
 rpc=async(method,...args)=>{calls.push([method,...args]);if(method==='previewManagedFormSettings')return {targets:args[0].targets.map(t=>({...t,label:t.id,changed:true,changes:[],fingerprint:'fp'}))};if(method==='runManagedFormSettings'){const r=state.forms.records.find(r=>r.id===args[0]);return {...r,materialSettingsSaved:true,...(args[4]?{materialSettingsBaseline:copy(args[1])}:{}),revision:r.id+'2',input:{...r.input,...args[1]}};}if(method==='saveManagedPublishingSettings')return {records:args[0].map(t=>({...state.forms.records.find(r=>r.id===t.id),publishingSettings:{...publishingSettings(state.forms.records.find(r=>r.id===t.id)),...args[1]},revision:t.id+'dates'}))};return {fingerprint:'fp'};};`,c);return c;
}

test('topic select has a blank default and uses first-course candidates without inserting HTML',()=>{
 const c=environment();vm.runInContext(`state.materialTopics={courseId:'100',courseName:'1A',topics:[{topicId:'t',name:'<復習>'}]};`,c);
 const html=vm.runInContext('renderBatchMaterialSettings()',c);
 assert.match(html,/<select[^>]+data-batch-material-field="topicName"[^>]*><option value="" selected><\/option>/);
 assert.match(html,/&lt;復習&gt;/);
});
test('topic changes make a saved material dirty and are sent only with material settings',async()=>{
 const c=environment(true);vm.runInContext(`materialSettingsDraft(state.forms.records[0]).topicName='復習';materialSettingsDraft(state.forms.records[0]).topicChange=true;`,c);
 assert.equal(vm.runInContext('materialSettingsReady(state.forms.records[0])',c),false);
 await vm.runInContext("formCommand('save-material-settings','a')",c);
 assert.equal(vm.runInContext('calls.find(c=>c[0]==="runManagedFormSettings")[2].topicName',c),'復習');
});
test('missing topics are confirmed with affected classes and cancellation does not write',async()=>{
 const c=environment(true);vm.runInContext(`state.batchMaterialDraft={materialTitlePattern:'{class_label} 新しい資料',description:'本文',topicName:'復習',topicChange:true};let confirms=[];confirmAction=async(...args)=>{confirms.push(args);return false;};const oldRpc=rpc;rpc=async(method,...args)=>{const result=await oldRpc(method,...args);if(method==='previewManagedFormSettings')result.targets[0].missingTopic={courseId:'100',className:'1A',name:'復習'};return result;};`,c);
 await vm.runInContext("formCommand('apply-batch-material-settings')",c);
 assert.match(vm.runInContext('confirms[0][1]',c),/復習[\s\S]*1A|1A[\s\S]*復習/);
 assert.match(vm.runInContext('confirms[0][1]',c),/作成/);
 assert.equal(vm.runInContext('calls.filter(c=>c[0].startsWith("run")).length',c),0);
});
test('each independent save leaves both other settings as drafts and sends only its own payload',async()=>{
 for(const action of ['apply-batch-material-settings','schedule-all','set-close-all']){
  const c=environment(true);vm.runInContext(`state.batchMaterialDraft={materialTitlePattern:'{class_label} 編集中',description:'編集中',topicName:''};state.scheduleDrafts={'batch:publish':'2027-10-02T10:00','batch:close':'2027-10-03T10:00'};const oldRpc=rpc;rpc=async(method,...args)=>method==='runManagedFormAction'?({...state.forms.records[0],...args[4]}):oldRpc(method,...args);`,c);
  await vm.runInContext(`formCommand('${action}')`,c);
  assert.equal(vm.runInContext("state.scheduleDrafts['batch:publish']",c),'2027-10-02T10:00');
  assert.equal(vm.runInContext("state.scheduleDrafts['batch:close']",c),'2027-10-03T10:00');
  if(action!=='apply-batch-material-settings')assert.equal(vm.runInContext('state.forms.records[0].input.description',c),'保存済み本文');
 }
});
test('bulk registration button precedes bulk posting and saves text, publish date and close date',async()=>{
 const c=environment();vm.runInContext(`state.batchMaterialDraft={materialTitlePattern:'{class_label} 新しい資料',description:'本文',topicName:''};state.scheduleDrafts={'batch:publish':'2027-10-02T10:00','batch:close':'2027-10-03T10:00'};const oldRpc=rpc;rpc=async(method,...args)=>{if(method==='runManagedFormAction'){calls.push([method,...args]);return {...state.forms.records[0],...args[4],stage:args[1]==='schedule'?'scheduled':state.forms.records[0].stage};}return oldRpc(method,...args);};`,c);
 const html=vm.runInContext('renderBatchSchedules()',c);assert.ok(html.includes('一括登録'));assert.ok(html.indexOf('一括登録</button>')<html.indexOf('一括投稿（'));
 await vm.runInContext("formCommand('register-all')",c);
 assert.equal(vm.runInContext('state.forms.records[0].input.description',c),'本文');
 assert.equal(vm.runInContext('state.forms.records[0].publishingSettings.scheduledTime',c),'2027-10-02T01:00:00.000Z');
 assert.equal(vm.runInContext('state.forms.records[0].publishingSettings.closesAt',c),'2027-10-03T01:00:00.000Z');
});
test('bulk registration rejects invalid or reversed dates before any writes',async()=>{
 const c=environment();vm.runInContext(`state.scheduleDrafts={'batch:publish':'2027-10-03T10:00','batch:close':'2027-10-02T10:00'};`,c);
 await vm.runInContext("formCommand('register-all')",c);
 assert.equal(vm.runInContext('calls.filter(c=>c[0].startsWith("run")).length',c),0);
 assert.match(vm.runInContext('messages.at(-1)',c),/終了.*後/);
});

test('confirmed missing-topic plan alone authorizes its topic creation',async()=>{
 const c=environment();vm.runInContext(`state.batchMaterialDraft={materialTitlePattern:'{class_label} 資料',description:'本文',topicName:'復習',topicChange:true};const oldRpc=rpc;rpc=async(method,...args)=>{const result=await oldRpc(method,...args);if(method==='previewManagedFormSettings')result.targets[0].missingTopic={courseId:'100',className:'1A',name:'復習'};return result;};`,c);
 await vm.runInContext("formCommand('apply-batch-material-settings')",c);
 assert.equal(vm.runInContext('calls.find(c=>c[0]==="runManagedFormSettings")[6]',c),true);
});
test('bulk cancellation stops all writes and a failed text save stops both dates',async()=>{
 for(const cancel of [true,false]){
  const c=environment();vm.runInContext(`state.scheduleDrafts={'batch:publish':'2027-10-02T10:00','batch:close':'2027-10-03T10:00'};confirmAction=async()=>${!cancel};const oldRpc=rpc;rpc=async(method,...args)=>{if(method==='runManagedFormSettings')throw Error('保存失敗');return oldRpc(method,...args);};`,c);
  await vm.runInContext("formCommand('register-all')",c);
  assert.equal(vm.runInContext('calls.filter(c=>c[0]==="runManagedFormAction").length',c),0);
  assert.equal(vm.runInContext('state.registrationPending',c),false);
 }
});
test('bulk registration confirms once and skips blank dates without a publish call',async()=>{
 const c=environment();vm.runInContext(`let confirmations=0;confirmAction=async()=>{confirmations++;return true;};state.scheduleDrafts={'batch:publish':'','batch:close':''};`,c);
 await vm.runInContext("formCommand('register-all')",c);
 assert.equal(vm.runInContext('confirmations',c),1);
 assert.equal(vm.runInContext('calls.some(c=>c[0]==="runManagedFormAction")',c),false);
 assert.equal(vm.runInContext('state.forms.records[0].materialSettingsSaved',c),true);
});
test('topic loading uses the first selected course and discards stale responses',async()=>{
 const c=environment();vm.runInContext(`state.forms.classes=[{courseId:'100'},{courseId:'200'}];let resolveTopic;rpc=()=>new Promise(resolve=>resolveTopic=resolve);`,c);
 const pending=vm.runInContext('refreshMaterialTopics()',c);
 vm.runInContext(`state.forms.classes=[{courseId:'200'}];resolveTopic({courseId:'100',topics:[{topicId:'t',name:'古い候補'}]});`,c);await pending;
 assert.equal(vm.runInContext('state.materialTopics',c),undefined);
 vm.runInContext(`rpc=async()=>({courseId:'200',courseName:'新クラス',topics:[{topicId:'u',name:'新候補'}]});`,c);
 await vm.runInContext('refreshMaterialTopics()',c);
 assert.equal(vm.runInContext('state.materialTopics.courseId',c),'200');
});
test('two bulk clicks during confirmation cannot execute the registration twice',async()=>{
 const c=environment();vm.runInContext('let confirmResolve;confirmAction=()=>new Promise(resolve=>confirmResolve=resolve);',c);
 const pending=vm.runInContext("formCommand('register-all')",c);
 await new Promise(resolve=>setImmediate(resolve));
 await vm.runInContext("formCommand('register-all')",c);
 vm.runInContext('confirmResolve(true)',c);await pending;
 assert.equal(vm.runInContext('calls.filter(c=>c[0]==="runManagedFormSettings").length',c),1);
});
test('bulk registration confirmation describes settings without requesting publication confirmation',async()=>{
 const c=environment();vm.runInContext(`state.scheduleDrafts={'batch:publish':'2027-10-02T10:00'};let summaries=[];confirmAction=async(title,summary)=>{summaries.push(summary);return false;};`,c);
 await vm.runInContext("formCommand('register-all')",c);
 const summary=vm.runInContext('summaries[0]',c);
 assert.match(summary,/個別設定へ保存/);assert.doesNotMatch(summary,/確認済み.*実行|公開前でもURL/);
 assert.equal(vm.runInContext('calls.some(c=>c[0].startsWith("run")||c[0]==="saveManagedPublishingSettings")',c),false);
});
test('bulk registration saves both replacement dates together without changing an active deadline',async()=>{
 const c=environment(true);vm.runInContext(`state.forms.records[0].closesAt='2027-10-03T01:00:00Z';state.scheduleDrafts={'batch:publish':'2027-10-05T10:00','batch:close':'2027-10-06T10:00'};`,c);
 await vm.runInContext("formCommand('register-all')",c);
 assert.equal(vm.runInContext('state.forms.records[0].publishingSettings.scheduledTime',c),'2027-10-05T01:00:00.000Z');
 assert.equal(vm.runInContext('state.forms.records[0].publishingSettings.closesAt',c),'2027-10-06T01:00:00.000Z');
 assert.equal(vm.runInContext('state.forms.records[0].closesAt',c),'2027-10-03T01:00:00Z');
 assert.equal(vm.runInContext('calls.some(c=>c[0]==="runManagedFormAction")',c),false);
});

test('registering an individually changed topic updates the durable common-topic baseline',async()=>{
 const c=environment(true);vm.runInContext(`const r=state.forms.records[0];r.input.topicName='新';r.materialSettingsBaseline={materialTitlePattern:r.input.materialTitlePattern,description:r.input.description,topicName:'旧'};state.batchMaterialDraft={...r.materialSettingsBaseline,topicName:'新',topicChange:true};`,c);
 await vm.runInContext("formCommand('apply-batch-material-settings')",c);
 assert.equal(vm.runInContext('state.forms.records[0].materialSettingsBaseline.topicName',c),'新');
 assert.equal(vm.runInContext('individualMaterialSettings(state.forms.records[0])',c),false);
});

test('unchanged topic is omitted from individual and bulk title-only saves',async()=>{
 for(const action of ['save-material-settings','apply-batch-material-settings','register-all']){
  const c=environment(true);
  vm.runInContext(`materialSettingsDraft(state.forms.records[0]).materialTitlePattern='{class_label} 編集';batchMaterialSettingsDraft().materialTitlePattern='{class_label} 編集';`,c);
  await vm.runInContext(`formCommand('${action}','a')`,c);
  assert.equal(vm.runInContext(`Object.hasOwn(calls.find(c=>c[0]==='runManagedFormSettings')[2],'topicName')`,c),false);
  assert.equal(vm.runInContext(`Object.hasOwn(calls.find(c=>c[0]==='runManagedFormSettings')[2],'topicChange')`,c),false);
 }
});
test('topic changes require the explicit checkbox and a checked blank means clear',async()=>{
 const c=environment(true);const html=vm.runInContext('renderMaterialSettings(state.forms.records[0])',c);
 assert.match(html,/type="checkbox"[^>]+data-material-field="topicChange"/);
 assert.match(html,/トピックを変更する/);
 vm.runInContext(`materialSettingsDraft(state.forms.records[0]).topicChange=true;`,c);
 assert.equal(vm.runInContext('materialSettingsReady(state.forms.records[0])',c),false);
 await vm.runInContext("formCommand('save-material-settings','a')",c);
 assert.equal(vm.runInContext('calls.find(c=>c[0]==="runManagedFormSettings")[2].topicChange',c),true);
 assert.equal(vm.runInContext('calls.find(c=>c[0]==="runManagedFormSettings")[2].topicName',c),'');
 assert.equal(vm.runInContext('materialSettingsDraft(state.forms.records[0]).topicChange===true',c),false);
});
test('cancel then uncheck preserves topic and reload resets explicit topic intent',async()=>{
 const c=environment(true);
 vm.runInContext(`const d=materialSettingsDraft(state.forms.records[0]);d.topicChange=true;d.topicName='';confirmAction=async()=>false;`,c);
 await vm.runInContext("formCommand('save-material-settings','a')",c);
 assert.equal(vm.runInContext('calls.filter(c=>c[0]==="runManagedFormSettings").length',c),0);
 vm.runInContext(`d.topicChange=false;d.description='本文のみ';confirmAction=async()=>true;`,c);
 await vm.runInContext("formCommand('save-material-settings','a')",c);
 assert.equal(vm.runInContext('Object.hasOwn(calls.find(c=>c[0]==="runManagedFormSettings")[2],"topicName")',c),false);
 vm.runInContext(`materialSettingsDraft(state.forms.records[0]).topicChange=true;`,c);
 await vm.runInContext("formCommand('reload-material-settings','a')",c);
 assert.equal(vm.runInContext('materialSettingsDraft(state.forms.records[0]).topicChange===true',c),false);
});
test('successful bulk topic change resets intent before the next title-only save',async()=>{
 const c=environment(true);
 vm.runInContext(`Object.assign(batchMaterialSettingsDraft(),{topicChange:true,topicName:''});`,c);
 await vm.runInContext("formCommand('apply-batch-material-settings')",c);
 assert.equal(vm.runInContext('calls.find(c=>c[0]==="runManagedFormSettings")[2].topicChange',c),true);
 assert.equal(vm.runInContext('batchMaterialSettingsDraft().topicChange===true',c),false);
 vm.runInContext(`calls=[];batchMaterialSettingsDraft().description='本文だけ';`,c);
 await vm.runInContext("formCommand('apply-batch-material-settings')",c);
 assert.equal(vm.runInContext('Object.hasOwn(calls.find(c=>c[0]==="runManagedFormSettings")[2],"topicName")',c),false);
});
test('a partial topic batch retries failed records without replaying successful clears',async()=>{
 const c=environment(true);
 vm.runInContext(`const other=copy(state.forms.records[0]);other.id='b';other.revision='b1';state.forms.records.push(other);Object.assign(batchMaterialSettingsDraft(),{topicChange:true,topicName:''});const backendRpc=rpc;let failOnce=true;rpc=async(method,...args)=>{if(method==='runManagedFormSettings'&&args[0]==='b'&&failOnce){failOnce=false;calls.push([method,...args]);throw Error('temporary failure');}return backendRpc(method,...args);};`,c);
 await vm.runInContext("formCommand('apply-batch-material-settings')",c);
 vm.runInContext('calls=[]',c);
 await vm.runInContext("formCommand('apply-batch-material-settings')",c);
 assert.deepEqual(Array.from(vm.runInContext('calls.filter(c=>c[0]==="runManagedFormSettings").map(c=>c[1])',c)),['b']);
});
test('legacy topic recovery explains preservation and cancellation does not resume',async()=>{
 const c=environment(true);
 vm.runInContext(`state.forms.records[0].settingsUpdate={plan:{topic:{name:''}}};let confirmations=[];confirmAction=async(title,summary)=>{confirmations.push(summary);return false;};`,c);
 await vm.runInContext("formCommand('resume-settings','a')",c);
 assert.equal(vm.runInContext('calls.filter(c=>c[0]==="resumeManagedFormSettings").length',c),0);
 assert.match(vm.runInContext('confirmations[0]',c),/トピック.*変更せず/);
 vm.runInContext('confirmAction=async()=>true;',c);
 await vm.runInContext("formCommand('resume-settings','a')",c);
 assert.equal(vm.runInContext('calls.filter(c=>c[0]==="resumeManagedFormSettings").length',c),1);
});
test('bulk retry reports a pending recovery rather than claiming all records saved',async()=>{
 const c=environment(true);
 vm.runInContext(`const other=copy(state.forms.records[0]);other.id='b';other.revision='b1';state.forms.records.push(other);Object.assign(batchMaterialSettingsDraft(),{topicChange:true,topicName:''});const backendRpc=rpc;rpc=async(method,...args)=>{if(method==='runManagedFormSettings'&&args[0]==='b'){state.forms.records.find(r=>r.id==='b').settingsUpdate={plan:{topic:{name:''},topicChange:true}};throw Error('uncertain write');}return backendRpc(method,...args);};`,c);
 await vm.runInContext("formCommand('apply-batch-material-settings')",c);
 const result=await vm.runInContext("formCommand('apply-batch-material-settings')",c);
 assert.equal(result,false);assert.match(vm.runInContext('messages.at(-1)',c),/設定の更新を再開/);
});
test('pending recovery blocks a mixed retry even when another failed record remains editable',async()=>{
 const c=environment(true);
 vm.runInContext(`for(const id of ['b','c']){const r=copy(state.forms.records[0]);r.id=id;r.revision=id+'1';state.forms.records.push(r);}Object.assign(batchMaterialSettingsDraft(),{topicChange:true,topicName:''});const backendRpc=rpc;let firstAttempt=true;rpc=async(method,...args)=>{if(method==='runManagedFormSettings'&&firstAttempt){if(args[0]==='b')state.forms.records.find(r=>r.id==='b').settingsUpdate={plan:{topic:{name:''},topicChange:true}};if(args[0]!=='a')throw Error('failure');}return backendRpc(method,...args);};`,c);
 await vm.runInContext("formCommand('apply-batch-material-settings')",c);
 vm.runInContext('firstAttempt=false;calls=[];',c);
 const result=await vm.runInContext("formCommand('apply-batch-material-settings')",c);
 assert.equal(result,false);assert.match(vm.runInContext('messages.at(-1)',c),/設定の更新を再開/);
 assert.equal(vm.runInContext('calls.filter(c=>c[0]==="runManagedFormSettings").length',c),0);
});
