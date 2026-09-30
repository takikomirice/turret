import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {adminEnvironment, AdminSheet, plain} from './helpers/admin-environment.mjs';

function env(){
 const e=adminEnvironment();
 e.c.ScriptApp.getScriptId=()=> 'script-one';
 e.ss.getUrl=()=> 'https://docs.google.com/spreadsheets/d/test-admin-sheet/edit';
 const folder={getId:()=> 'turret-parent'};
 e.c.DriveApp={getFolderById:()=>folder,getFileById:()=>({getParents:()=>({hasNext:()=>true,next:()=>folder}),moveTo:()=>{}})};
 e.sheets.set('クラス一覧',new AdminSheet('クラス一覧',[['クラス名','コースID','同期対象(1)'],['テストクラス','100','1']]));
 return e;
}

test('blank template creation stays unpublished and retries return the same form',()=>{
 const {c,props}=env();let creates=0;const before=JSON.stringify([...props]);
 c.FormApp={create:(title,published)=>{creates++;assert.match(title,/ひな形/);assert.equal(published,false);return {getId:()=> 'new-template-form-id'};}};
 const result=c.createTemplateForm('request-1234567890');
 assert.equal(result.editUrl,'https://docs.google.com/forms/d/new-template-form-id/edit');
 assert.deepEqual(plain(c.createTemplateForm('request-1234567890')),plain(result));
 assert.equal(creates,1);assert.equal(JSON.stringify([...props]),before);
});

for(const parentId of ['turret-folder','my-drive-root'])test('template is placed beside its workbook: '+parentId,()=>{
 const {c}=env();let moved,creates=0;
 const folder={getId:()=>parentId};
 c.DriveApp={getFolderById:id=>{assert.equal(id,parentId);return folder;},getFileById:id=>id==='test-admin-sheet'
  ?{getParents:()=>({hasNext:()=>true,next:()=>folder})}
  :{getParents:()=>({hasNext:()=>false}),moveTo:target=>{moved=target.getId();}}};
 c.FormApp={create:()=>{creates++;return {getId:()=> 'new-template-id'};}};
 c.createTemplateForm('request-1234567890');
 assert.equal(moved,parentId);assert.equal(creates,1);
});

test('missing workbook parent stops before form creation',()=>{
 const {c}=env();let creates=0;
 c.DriveApp={getFileById:()=>({getParents:()=>({hasNext:()=>false})})};
 c.FormApp={create:()=>{creates++;}};
 assert.throws(()=>c.createTemplateForm('request-1234567890'),/保存先フォルダ/);
 assert.equal(creates,0);assert.equal(c.getManagedRecords_().length,0);
});

for(const responseLost of [false,true])test('placement retry reuses the created form after failure; response lost='+responseLost,()=>{
 const {c}=env();let creates=0,moves=0,inFolder=false;
 const folder={getId:()=> 'turret-folder'};
 c.DriveApp={getFolderById:()=>folder,getFileById:id=>id==='test-admin-sheet'
  ?{getParents:()=>({hasNext:()=>true,next:()=>folder})}
  :{getParents:()=>({hasNext:()=>inFolder,next:()=>folder}),moveTo:()=>{moves++;if(moves===1){inFolder=responseLost;throw Error('move failed');}inFolder=true;}}};
 c.FormApp={create:()=>{creates++;return {getId:()=> 'new-template-id'};}};
 assert.throws(()=>c.createTemplateForm('request-1234567890'),/再試行/);
 assert.equal(c.createTemplateForm('request-1234567890').editUrl,'https://docs.google.com/forms/d/new-template-id/edit');
 assert.equal(creates,1);assert.equal(moves,responseLost?1:2);
});

test('ambiguous template creation cannot be repeated and invalid requests do no I/O',()=>{
 const {c}=env();let creates=0;
 c.FormApp={create:()=>{creates++;throw Error('connection lost');}};
 assert.throws(()=>c.createTemplateForm('../bad'),/識別/);
 assert.equal(creates,0);
 assert.throws(()=>c.createTemplateForm('request-1234567890'),/connection lost/);
 assert.throws(()=>c.createTemplateForm('request-1234567890'),/マイドライブ/);
 assert.equal(creates,1);
});

test('a failed completion checkpoint retries placement without creating another form',()=>{
 const {c}=env();let creates=0,saves=0;
 c.FormApp={create:()=>{creates++;return {getId:()=> 'created-form-id'};}};
 const save=c.saveManagedRecord_;
 c.saveManagedRecord_=r=>{if(++saves===3)throw Error('write failed');return save(r);};
 assert.throws(()=>c.createTemplateForm('request-1234567890'),/再試行/);
 assert.match(c.createTemplateForm('request-1234567890').editUrl,/created-form-id/);
 assert.equal(creates,1);
});

test('template creation refuses copied journals and does not repeat after checkpoint failure',()=>{
 const {c}=env();let creates=0;
 c.FormApp={create:()=>{creates++;return {getId:()=> 'created-form-id'};}};
 const save=c.saveManagedRecord_;let saves=0;
 c.saveManagedRecord_=r=>{if(++saves===2)throw Error('write failed');return save(r);};
 assert.throws(()=>c.createTemplateForm('request-1234567890'),/created-form-id/);
 assert.throws(()=>c.createTemplateForm('request-1234567890'),/マイドライブ/);
 assert.equal(creates,1);
 c.ScriptApp.getScriptId=()=> 'copied-script';
 assert.throws(()=>c.createTemplateForm('request-1234567890'),/別の運用/);
});
test('form authorization setup requests the integration scopes without creating data or posts',()=>{
 const {c,sheets,props,triggers}=env();
 const before=JSON.stringify([...sheets].map(([name,s])=>[name,s.rows]));
 const savedProps=JSON.stringify([...props]);let requested;
 c.ScriptApp.AuthMode={FULL:'FULL'};
 c.ScriptApp.requireScopes=(mode,scopes)=>{requested={mode,scopes:plain(scopes)};};
 assert.equal(c.authorizeFormIntegration().authorized,true);
 assert.deepEqual(requested,{mode:'FULL',scopes:[
  'https://www.googleapis.com/auth/forms',
  'https://www.googleapis.com/auth/drive',
  'https://www.googleapis.com/auth/classroom.courseworkmaterials'
 ]});
 assert.equal(JSON.stringify([...sheets].map(([name,s])=>[name,s.rows])),before);
 assert.equal(JSON.stringify([...props]),savedProps);assert.equal(triggers.length,0);
});
test('form authorization setup never reports success when consent is missing',()=>{
 const {c}=env();c.ScriptApp.AuthMode={FULL:'FULL'};
 c.ScriptApp.requireScopes=()=>{throw Error('consent required');};
 assert.throws(()=>c.authorizeFormIntegration(),/consent required/);
});
test('native authorization does not request Forms REST access',()=>{
 const manifest=JSON.parse(readFileSync('appsscript.json','utf8'));
 assert.ok(!manifest.oauthScopes.includes('https://www.googleapis.com/auth/forms.body.readonly'));
 assert.ok(manifest.oauthScopes.includes('https://www.googleapis.com/auth/forms'));
});

test('journal refuses copied ownership and stale mutations',()=>{
 const {c}=env();
 const record=c.saveManagedRecord_({id:'job:100',kind:'form',courseId:'100',stage:'pending'});
 assert.equal(c.loadManagedRecord_(record.id,record.revision).courseId,'100');
 assert.throws(()=>c.loadManagedRecord_(record.id,'old'),/更新/);
 record.ownerSpreadsheetId='other';
 assert.throws(()=>c.saveManagedRecord_(record),/運用/);
});
test('journal does not overwrite unrelated sheet or corrupted JSON',()=>{
 const {c,sheets}=env();
 sheets.set('フォーム管理',new AdminSheet('フォーム管理',[['手入力'],['残す']]));
 assert.throws(()=>c.getManagedRecords_(),/構成/);
 assert.equal(sheets.get('フォーム管理').rows[1][0],'残す');
});
test('responder policy rejects public, unknown groups, wrong domains and unverified email',()=>{
 const {c}=env(); const policy={domains:['school.example'],emails:[]};
 const person=[{type:'domain',role:'reader',view:'published',domain:'school.example'}];
 assert.equal(c.validateResponderAccess_(person,policy,['kid@school.example'],true),true);
 for(const bad of [{type:'anyone'},{type:'group',emailAddress:'g@school.example'},{type:'domain',domain:'other.example'}]){
  assert.throws(()=>c.validateResponderAccess_([...person,{role:'reader',view:'published',...bad}],policy,['kid@school.example'],true),/回答/);
 }
 assert.throws(()=>c.validateResponderAccess_(person,policy,['kid@school.example'],false),/メール収集/);
 assert.throws(()=>c.validateResponderAccess_([],policy,['kid@school.example'],true),/回答/);
});
test('personal responders explicitly allow only the test student and owner',()=>{
 const {c}=env();
 const policy={domains:[],emails:['kid@gmail.com','teacher@gmail.com']};
 assert.equal(c.validateResponderAccess_([{type:'user',role:'reader',view:'published',emailAddress:'kid@gmail.com'}],policy,['kid@gmail.com'],true),true);
 assert.throws(()=>c.validateResponderAccess_([{type:'user',role:'reader',view:'published',emailAddress:'kid@gmail.com'}],policy,['other@gmail.com'],true),/回答/);
});
test('deletion preview and execution are scoped; successful deletion cannot re-send',()=>{
 const {c,sheets}=env();let removed=0;
 c.Classroom={Courses:{Announcements:{get:()=>({id:'post',text:'本文',state:'PUBLISHED'}),remove:(course,id)=>{assert.equal(course,'100');assert.equal(id,'post');removed++;}}}};
 const send=new AdminSheet('送信シート',[['送信状態'],['済']]);sheets.set('送信シート',send);
 const record=c.saveManagedRecord_({id:'announcement:post',kind:'announcement',courseId:'100',postId:'post',stage:'published',title:'テスト返却'});
 const p=c.previewManagedFormAction(record.id,'delete');
 assert.throws(()=>c.runManagedFormAction(record.id,'delete','wrong'),/確認|更新/);
 const done=c.runManagedFormAction(record.id,'delete',p.fingerprint);
 assert.equal(done.stage,'deleted');assert.equal(removed,1);assert.equal(send.rows[1][0],'済');
 assert.throws(()=>c.runManagedFormAction(record.id,'delete',p.fingerprint),/削除|更新|確認/);
 assert.equal(removed,1);
});
test('unknown delete result keeps tracked ID and can be reconciled without blind deletion',()=>{
 const {c}=env();
 c.Classroom={Courses:{CourseWorkMaterials:{get:()=>({id:'material',state:'PUBLISHED',title:'資料'}),remove:()=>{throw Error('timeout');}}}};
 const r=c.saveManagedRecord_({id:'form:x',kind:'form',courseId:'100',materialId:'material',stage:'published',title:'資料'});
 const p=c.previewManagedFormAction(r.id,'delete');
 assert.throws(()=>c.runManagedFormAction(r.id,'delete',p.fingerprint),/timeout/);
 const after=c.loadManagedRecord_(r.id);assert.equal(after.materialId,'material');assert.equal(after.stage,'delete_review');
});
test('name enrichment matches course and email, preserves grades and manual names',()=>{
 const {c,sheets}=env();
 sheets.set('生徒一覧',new AdminSheet('生徒一覧',[['メールアドレス','名前','コースID'],['KID@EXAMPLE.COM','正しい名前','100'],['kid@example.com','別クラス名','200']]));
 const answer=new AdminSheet('フォーム回答 テスト',[['メール','名前','評価'],['kid@example.com','','5'],['missing@example.com','','3'],['kid@example.com','手入力','4']]);
 const result=c.fillManagedNames_({courseId:'100',input:{emailHeader:'メール',nameHeader:'名前'}},answer);
 assert.deepEqual(plain(result),{updated:1,unmatched:1,skipped:1});
 assert.deepEqual(answer.rows.slice(1),[['kid@example.com','正しい名前','5'],['missing@example.com','','3'],['kid@example.com','手入力','4']]);
});

function provisioningEnv(){
 const e=env(),{c,sheets,ss}=e;let next=0;
 const calls={copies:0,posts:0,removes:0,destination:0,trigger:0};const forms=new Map(),permissions=new Map();
 const templateId='template_12345678901234567890',folderId='folder_12345678901234567890';
 ss.getName=()=> 'turret test';
 for(const sheet of sheets.values())sheet.getSheetId=()=> ++next;
 const originalInsert=ss.insertSheet;ss.insertSheet=name=>{const sheet=originalInsert(name),id=++next;sheet.getSheetId=()=>id;sheet.getFormUrl=()=>null;return sheet;};
 sheets.set('生徒一覧',new AdminSheet('生徒一覧',[['メールアドレス','名前','コースID'],['kid@example.com','生徒','100']]));
 for(const sheet of sheets.values()){const id=++next;sheet.getSheetId=()=>id;sheet.getFormUrl=()=>null;}
 c.DriveApp={getFolderById:()=>({getName:()=> 'folder'}),getFileById:()=>({getLastUpdated:()=>new Date('2026-09-29T00:00:00Z')})};
 function form(id){const state={published:false,accepting:false,destination:null};return {state,getId:()=>id,supportsAdvancedResponderPermissions:()=>true,
  getPublishedUrl:()=> 'https://docs.google.com/forms/d/'+id+'/viewform',getEditUrl:()=> 'https://docs.google.com/forms/d/'+id+'/edit',
  setTitle(){},isPublished:()=>state.published,isAcceptingResponses:()=>state.accepting,
  setPublished(v){state.published=v;state.accepting=v;},setAcceptingResponses(v){state.accepting=v;},getDestinationId:()=>state.destination,
  setDestination(type,destination){calls.destination++;state.destination=destination;const sheet=ss.insertSheet('Form Responses '+id);sheet.rows=[['メール','質問']];sheet.getFormUrl=()=>id;sheet.setName=name=>{sheets.delete(sheet.name);sheet.name=name;sheets.set(name,sheet);};}
 };}
 forms.set(templateId,form(templateId));
 c.FormApp={DestinationType:{SPREADSHEET:'SPREADSHEET'},openById:id=>{if(!forms.has(id))throw Error('missing form');return forms.get(id);},openByUrl:id=>forms.get(id)};
 c.formMetadata_=()=>({settings:{collectsEmail:true},items:[]});
 c.Drive={Files:{copy:body=>{calls.copies++;const id='copy-'+calls.copies;forms.set(id,form(id));permissions.set(id,[]);return {id};}},Permissions:{
  list:id=>({permissions:permissions.get(id)||[]}),create:(p,id)=>{permissions.get(id).push({...p,id:'perm'});},remove:(id,pid)=>permissions.set(id,permissions.get(id).filter(p=>p.id!==pid))
 }};
 c.ensureManagedResponseTrigger_=()=>{calls.trigger++;};
 c.Classroom={Courses:{CourseWorkMaterials:{create:(data,course)=>{calls.posts++;return {id:'material-'+calls.posts,alternateLink:'https://classroom.google.com/c/'+course};},get:()=>({id:'material-1',state:'PUBLISHED'}),remove:()=>{calls.removes++;}}}};
 const raw={templateId,folderId,verifiedEmailConfirmed:true,titlePattern:'振り返り {class}',materialTitlePattern:'資料 {class}',description:'回答してください',prefix:'回答',emailHeader:'メール',nameHeader:'名前',gradeHeader:'点数',commentHeader:'コメント',statusHeader:'状態',domains:'example.com',emails:'',gradeChoices:'',labels:{'100':'1-1'}};
 const preview=()=>c.previewFormSetup(raw,c.getConfigRevision_(c.getConfig_()));
 const begin=()=>{const p=preview();return c.beginFormSetup(raw,p.configRevision,p.fingerprint);};
 return {...e,calls,forms,permissions,raw,preview,begin};
}

test('publication confirmation ignores API object and permission ordering but detects access changes',()=>{
 const e=provisioningEnv(),{c}=e;let [r]=e.begin();r=c.prepareFormTarget(r.id,r.revision);
 const original=e.permissions.get(r.formId)[0];
 const extra={id:'extra',type:'user',role:'reader',view:'published',emailAddress:'teacher@example.com'};
 e.permissions.set(r.formId,[original,extra]);
 const p=c.previewManagedFormAction(r.id,'publish');
 e.permissions.set(r.formId,[extra,Object.fromEntries(Object.entries(original).reverse())]);
 assert.equal(c.previewManagedFormAction(r.id,'publish').fingerprint,p.fingerprint);
 e.permissions.set(r.formId,[original]);
 assert.notEqual(c.previewManagedFormAction(r.id,'publish').fingerprint,p.fingerprint);
});

test('missing Classroom consent stops before publication and keeps the prepared record retryable',()=>{
 const e=provisioningEnv(),{c}=e;let [r]=e.begin();r=c.prepareFormTarget(r.id,r.revision);
 const p=c.previewManagedFormAction(r.id,'publish');
 c.ScriptApp.requireScopes=()=>{throw Error('consent required');};
 assert.throws(()=>c.runManagedFormAction(r.id,'publish',p.fingerprint,true),/consent required/);
 assert.equal(c.loadManagedRecord_(r.id).stage,'registered');assert.equal(e.forms.get(r.formId).state.published,false);assert.equal(e.calls.posts,0);
});
test('canonical class labels render in titles and legacy class tokens remain compatible',()=>{
 const e=provisioningEnv();
 e.raw.titlePattern='振り返り {class_label}';e.raw.materialTitlePattern='資料 {class_label} / {class}';
 let target=e.preview().targets[0];
 assert.equal(target.title,'振り返り 1-1');assert.equal(target.materialTitle,'資料 1-1 / 1-1');
 e.raw.labels={};target=e.preview().targets[0];
 assert.equal(target.title,'振り返り テストクラス');
 e.raw.titlePattern='{class_name}';assert.throws(()=>e.preview(),/class_label/);
});

test('domain input accepts the school email suffix without an at sign',()=>{
 const e=provisioningEnv();
 assert.deepEqual(plain(e.c.normalizeFormSetup_(e.raw).policy.domains),['example.com']);
 e.raw.domains='@example.com';assert.throws(()=>e.c.normalizeFormSetup_(e.raw),/回答者/);
});

function flexibleInput(e, columns=[]){
 delete e.raw.gradeHeader;delete e.raw.commentHeader;delete e.raw.gradeChoices;
 e.raw.additionalColumns=columns;
}

test('optional columns allow no comments or grades without changing the existing completion column',()=>{
 const e=provisioningEnv();flexibleInput(e);
 const [r]=e.begin();const done=e.c.prepareFormTarget(r.id,r.revision);
 assert.equal(done.stage,'registered');
 assert.deepEqual(e.sheets.get('回答 1-1').rows[0],['メール','質問','名前','状態']);
 assert.equal(e.c.getConfig_().scoreSourceHeader,'点数');
 assert.equal(e.c.getConfig_().fields.length,1);
});

test('multiple custom columns use independent optional dropdowns and preserve existing field conversions',()=>{
 const e=provisioningEnv();flexibleInput(e,[{header:'点数',choices:['1','2','3']},{header:'思考',choices:'A,B,C'},{header:'補足',choices:[]}]);
 const validations=[];
 e.c.SpreadsheetApp.newDataValidation=()=>{let values;return {requireValueInList(v){values=plain(v);return this;},setAllowInvalid(v){assert.equal(v,false);return this;},build(){return values;}};};
 const insert=e.ss.insertSheet;e.ss.insertSheet=name=>{const sheet=insert(name),getRange=sheet.getRange.bind(sheet);sheet.getRange=(...args)=>{const range=getRange(...args);range.setDataValidation=rule=>{validations.push({args,rule});return range;};return range;};return sheet;};
 const old=plain(e.c.getConfig_().fields[0]);const [r]=e.begin();e.c.prepareFormTarget(r.id,r.revision);
 assert.deepEqual(e.sheets.get('回答 1-1').rows[0],['メール','質問','名前','点数','思考','補足','状態']);
 assert.deepEqual(validations.map(v=>v.args[1]),[4,5]);
 assert.deepEqual(validations.map(v=>v.rule),[['1','2','3'],['A','B','C']]);
 const config=e.c.getConfig_();assert.deepEqual(plain(config.fields[0]),old);
 assert.deepEqual(plain(config.fields.map(f=>f.sourceHeader)),['点数','思考','補足']);
 assert.equal(config.scoreSourceHeader,'点数');
});

test('new blank setup does not assign a completion column from the first extra column',()=>{
 const e=provisioningEnv();e.props.delete('APP_CONFIG');flexibleInput(e,[{header:'観点A',choices:[]}]);
 const [r]=e.begin();e.c.prepareFormTarget(r.id,r.revision);
 assert.equal(e.c.getConfig_().scoreSourceHeader,'');
 assert.deepEqual(plain(e.c.getConfig_().fields.map(f=>f.sourceHeader)),['観点A']);
});

test('custom columns reject malformed lists, duplicate headers and protected column collisions before copying',()=>{
 for(const columns of [null,{},[{header:''}],[{header:'名前'}],[{header:'メール'}],[{header:'状態'}],[{header:'同じ'},{header:'同じ'}],[{header:'評価',choices:{bad:1}}]]){
  const e=provisioningEnv();flexibleInput(e,columns);
  assert.throws(()=>e.preview(),/列|選択肢/);assert.equal(e.calls.copies,0);
 }
});

test('custom column setup resumes after interrupted configuration save without shifting or erasing grades',()=>{
 const e=provisioningEnv();flexibleInput(e,[{header:'観点A',choices:[]},{header:'観点B',choices:[]}]);
 const [r]=e.begin(),save=e.c.applyConfigDraft_;e.c.applyConfigDraft_=()=>{throw Error('config storage');};
 assert.throws(()=>e.c.prepareFormTarget(r.id,r.revision),/config storage/);
 const sheet=e.sheets.get('回答 1-1');sheet.rows.push(['kid@example.com','回答','生徒','5','4','']);
 const before=plain(sheet.rows);e.c.applyConfigDraft_=save;
 const pending=e.c.loadManagedRecord_(r.id);e.c.prepareFormTarget(r.id,pending.revision);
 assert.deepEqual(sheet.rows,before);assert.equal(e.calls.copies,1);
});

test('legacy pending records keep the original four-column layout and completion mapping',()=>{
 const e=provisioningEnv();const [r]=e.begin();
 // Simulate records persisted before additionalColumns existed.
 const stored=e.c.loadManagedRecord_(r.id);delete stored.input.additionalColumns;e.c.saveManagedRecord_(stored);
 e.c.prepareFormTarget(stored.id,stored.revision);
 assert.deepEqual(e.sheets.get('回答 1-1').rows[0],['メール','質問','名前','点数','コメント','状態']);
 assert.equal(e.c.getConfig_().scoreSourceHeader,'点数');
});

test('missing selected classes return actionable diagnostics instead of an unexplained empty list',()=>{
 const e=provisioningEnv();e.sheets.get('クラス一覧').rows[1][2]='';
 const data=e.c.getFormConsoleData();assert.equal(data.classes.length,0);assert.match(data.classWarning,/同期対象/);
});

test('new field layout conflicts stop before Drive copies when existing evaluation data would be affected',()=>{
 const e=provisioningEnv();flexibleInput(e,[{header:'新しい観点',choices:[]}]);
 const headers=plain(e.c.getConfiguredEvalHeaders_(e.c.getConfig_()));
 e.sheets.set('評価データ',new AdminSheet('評価データ',[headers,headers.map(()=> '保持する')]));
 const before=plain(e.sheets.get('評価データ').rows);
 assert.throws(()=>e.preview(),/既存データ/);assert.equal(e.calls.copies,0);
 assert.deepEqual(e.sheets.get('評価データ').rows,before);
});

test('one class provisions, links, registers and posts only on explicit action; deletion permits explicit repost',()=>{
 const e=provisioningEnv(),{c,calls,forms}=e;
 const p=e.preview();assert.equal(calls.copies,0);
 let [r]=c.beginFormSetup(e.raw,p.configRevision,p.fingerprint);assert.equal(calls.copies,0);
 r=c.prepareFormTarget(r.id,r.revision);assert.equal(r.stage,'registered');assert.equal(calls.copies,1);assert.equal(calls.posts,0);assert.equal(forms.get(r.formId).state.published,false);
 c.prepareFormTarget(r.id,r.revision);assert.equal(calls.copies,1);assert.equal(calls.destination,1);
 let confirm=c.previewManagedFormAction(r.id,'publish');r=c.runManagedFormAction(r.id,'publish',confirm.fingerprint,true);assert.equal(calls.posts,1);
 assert.throws(()=>c.previewManagedFormAction(r.id,'publish'),/投稿済み/);
 confirm=c.previewManagedFormAction(r.id,'delete');r=c.runManagedFormAction(r.id,'delete',confirm.fingerprint);assert.equal(calls.removes,1);
 confirm=c.previewManagedFormAction(r.id,'publish');r=c.runManagedFormAction(r.id,'publish',confirm.fingerprint,true);assert.equal(calls.posts,2);assert.equal(r.deletedPosts[0].id,'material-1');
});
test('copy success followed by response loss never triggers a second copy',()=>{
 const e=provisioningEnv(),{c,calls}=e;const [r]=e.begin(),copy=c.Drive.Files.copy;
 c.Drive.Files.copy=body=>{copy(body);throw Error('response lost');};
 assert.throws(()=>c.prepareFormTarget(r.id,r.revision),/response lost/);
 const saved=c.loadManagedRecord_(r.id);assert.equal(saved.stage,'creating');assert.equal(calls.copies,1);
 assert.throws(()=>c.prepareFormTarget(r.id,saved.revision),/未確認/);assert.equal(calls.copies,1);
});
test('publication verifies updated permissions; no call occurs on stale preview',()=>{
 const e=provisioningEnv(),{c,calls,permissions}=e;let [r]=e.begin();r=c.prepareFormTarget(r.id,r.revision);
 const p=c.previewManagedFormAction(r.id,'publish');permissions.get(r.formId).push({type:'anyone',role:'reader',view:'published'});
 assert.throws(()=>c.runManagedFormAction(r.id,'publish',p.fingerprint,true),/回答/);assert.equal(calls.posts,0);
});
test('name column collision is rejected without overwriting response headers',()=>{
 const e=provisioningEnv(),{c,forms}=e;const [r]=e.begin(),copy=c.Drive.Files.copy;
 c.Drive.Files.copy=body=>{const result=copy(body);const f=forms.get(result.id),set=f.setDestination;f.setDestination=(...args)=>{set(...args);e.sheets.get('Form Responses '+result.id).rows=[['メール','名前']];};return result;};
 assert.throws(()=>c.prepareFormTarget(r.id,r.revision),/重複/);
 assert.deepEqual(e.sheets.get('回答 1-1').rows[0],['メール','名前']);
});
test('a partially persisted batch recovers missing class records before creating forms',()=>{
 const e=provisioningEnv(),{c,sheets}=e;
 sheets.get('クラス一覧').rows.push(['もう一組','200','1']);sheets.get('生徒一覧').rows.push(['second@example.com','生徒2','200']);
 const save=c.saveManagedRecord_;let count=0;
 c.saveManagedRecord_=r=>{if(r.kind==='form'&&++count===2)throw Error('storage');return save(r);};
 const p=e.preview();assert.throws(()=>c.beginFormSetup(e.raw,p.configRevision,p.fingerprint),/storage/);
 c.saveManagedRecord_=save;const records=c.beginFormSetup(e.raw,p.configRevision,p.fingerprint);
 assert.equal(records.length,2);assert.equal(c.getManagedRecords_().filter(r=>r.kind==='form').length,2);
});
test('500 tracked return IDs use one registry read per execution with cache',()=>{
 const {c}=env();const cache={records:[],sheet:null};let reads=0;const read=c.getManagedRecords_;c.getManagedRecords_=()=>{reads++;return read();};
 for(let i=0;i<500;i++)c.recordManagedAnnouncement_({id:'p'+i},{courseId:'100',name:'生徒',sourceKey:'source'+i},cache);
 assert.equal(reads,0);assert.equal(c.getManagedRecords_().length,500);
});

test('deployed send path records IDs and preserves uncertain journal outcomes without resending',()=>{
 for(const fails of [false,true]){
  const {c,sheets,ss}=env();let posts=0;
  const config=c.getConfig_(),headers=plain(c.getConfiguredSendHeaders_(config));
  const row=headers.map(h=>({'元SS_ID':'source_12345678901234567890','元シート名':'回答 1','元行番号':'2','No':'1','メールアドレス':'kid@example.com','名前':'生徒','コースID':'100','studentId':'student-1','送信状態':'未','返信本文':'確認用','点数':'5'}[h]||''));
  const send=new AdminSheet('送信シート',[headers,row]);sheets.set('送信シート',send);
  c.Classroom={Courses:{Announcements:{create:()=>{posts++;return {id:'sent-1'};}}}};
  c.SpreadsheetApp.openById=()=>({getSheetByName:()=>new AdminSheet('回答 1',[['状態'],['']])});
  if(fails){const original=ss.insertSheet;ss.insertSheet=name=>{const sheet=original(name);if(name==='フォーム管理')sheet.beforeWrite=values=>{if(values[0][0]==='announcement')throw Error('registry storage unavailable');};return sheet;};}
  c.sendMessages();assert.equal(posts,1);
  assert.equal(send.rows[1][headers.indexOf('送信状態')],fails?'送信確認待ち':'済');
  if(!fails)assert.equal(c.getManagedRecords_()[0].postId,'sent-1');
  c.sendMessages();assert.equal(posts,1);
 }
});
test('regular import backfills names missed by a busy event without changing grades',()=>{
 const {c,sheets}=env();
 const answer=new AdminSheet('回答 1',[['メール','名前','点数','状態'],['kid@example.com','','5','済']]);answer.getSheetId=()=>99;
 sheets.set('回答 1',answer);sheets.set('生徒一覧',new AdminSheet('生徒一覧',[['メールアドレス','名前','コースID'],['kid@example.com','生徒','100']]));
 c.saveManagedRecord_({id:'form:100',kind:'form',courseId:'100',responseSheetId:99,stage:'published',input:{emailHeader:'メール',nameHeader:'名前'}});
 for(const s of sheets.values())if(!s.getSheetId)s.getSheetId=()=>1;
 c.SpreadsheetApp.openById=()=>({getSheets:()=>[]});
 c.importFromFormsToEval();
 assert.deepEqual(answer.rows[1],['kid@example.com','生徒','5','済']);
});
test('failed material publication can still close response intake',()=>{
 const e=provisioningEnv(),{c,forms}=e;let [r]=e.begin();r=c.prepareFormTarget(r.id,r.revision);
 c.Classroom.Courses.CourseWorkMaterials.create=()=>{throw Error('connection lost');};
 let p=c.previewManagedFormAction(r.id,'publish');assert.throws(()=>c.runManagedFormAction(r.id,'publish',p.fingerprint,true),/connection lost/);
 assert.equal(forms.get(r.formId).state.accepting,true);
 p=c.previewManagedFormAction(r.id,'close');c.runManagedFormAction(r.id,'close',p.fingerprint);
 assert.equal(forms.get(r.formId).state.accepting,false);
});
test('self-source excludes internal tabs but preserves matching names in external workbooks',()=>{
 const {c,ss}=env();
 assert.equal(c.isManagedInternalSheet_('フォーム管理',ss.getId()),true);
 assert.equal(c.isManagedInternalSheet_('評価データ',ss.getId()),true);
 assert.equal(c.isManagedInternalSheet_('回答 1',ss.getId()),false);
 assert.equal(c.isManagedInternalSheet_('評価データ','external-workbook'),false);
});
test('material response without an ID remains uncertain and never auto-posts again',()=>{
 const e=provisioningEnv(),{c}=e;let [r]=e.begin();r=c.prepareFormTarget(r.id,r.revision);
 c.Classroom.Courses.CourseWorkMaterials.create=()=>({});
 const p=c.previewManagedFormAction(r.id,'publish');
 assert.throws(()=>c.runManagedFormAction(r.id,'publish',p.fingerprint,true),/投稿ID/);
 assert.equal(c.loadManagedRecord_(r.id).stage,'publish_review');
 assert.throws(()=>c.previewManagedFormAction(r.id,'publish'),/確認待ち/);
});

test('setup requires explicit human confirmation of the verified email setting',()=>{
 const e=provisioningEnv();delete e.raw.verifiedEmailConfirmed;
 assert.throws(()=>e.preview(),/確認済み/);assert.equal(e.calls.copies,0);
 e.raw.verifiedEmailConfirmed=true;assert.equal(e.preview().input.verifiedEmailConfirmed,true);
 e.c.formMetadata_=()=>({settings:{collectsEmail:false},items:[]});
 assert.throws(()=>e.preview(),/メール収集/);
});

test('new Forms throw for a missing destination; provisioning links once and propagates other errors',()=>{
 const e=provisioningEnv(),{c,forms}=e,copy=c.Drive.Files.copy;
 c.Drive.Files.copy=(...args)=>{const r=copy(...args),f=forms.get(r.id);f.getDestinationId=()=>{if(!f.state.destination)throw Error('The form currently has no response destination.');return f.state.destination;};return r;};
 let [r]=e.begin();r=c.prepareFormTarget(r.id,r.revision);assert.equal(r.stage,'registered');assert.equal(e.calls.destination,1);
 assert.throws(()=>c.managedDestinationId_({getDestinationId(){throw Error('permission denied');}}),/permission denied/);
});
test('publishing requires a fresh human confirmation and closing never needs email verification',()=>{
 const e=provisioningEnv(),{c}=e;let [r]=e.begin();r=c.prepareFormTarget(r.id,r.revision);
 let p=c.previewManagedFormAction(r.id,'publish');
 assert.throws(()=>c.runManagedFormAction(r.id,'publish',p.fingerprint),/確認済み/);assert.equal(e.calls.posts,0);
 c.runManagedFormAction(r.id,'publish',p.fingerprint,true);assert.equal(e.calls.posts,1);
 c.formMetadata_=()=>{throw Error('unavailable metadata');};
 p=c.previewManagedFormAction(r.id,'close');c.runManagedFormAction(r.id,'close',p.fingerprint);
 assert.equal(e.forms.get(r.formId).state.accepting,false);
});
