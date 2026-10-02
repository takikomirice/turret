import test from 'node:test';
import assert from 'node:assert/strict';
import {scoringEnvironment,AdminSheet} from './helpers/scoring-environment.mjs';
const plain=value=>JSON.parse(JSON.stringify(value));
const any=()=>Array.from({length:5},()=>({mode:'any'}));
const rule=(condition='',outputs=['講評','',''],outputVisible=[true,true,true])=>({conditions:[condition,'','','',''],outputs,outputVisible,enabled:true});
const template=(name='化学')=>({name,scorePolicy:[{mode:'list',values:[-1.5,0,2]},...any().slice(1)],rules:[rule('>=0',['A','B','C'],[true,false,true])]});
const save=(e,templates)=>e.c.scoringSaveTemplateTable({revision:e.c.scoringGetTemplateTable().revision,templates});
const policyKeys=e=>[...e.props.keys()].filter(k=>k.startsWith('TURRET_TEMPLATE_POLICY_V2_'));

test('thirteen exact columns and Japanese header notes preserve per-output visibility',()=>{
 const e=scoringEnvironment();let notes;
 const insert=e.ss.insertSheet;e.ss.insertSheet=name=>{const s=insert(name),get=s.getRange.bind(s);s.getRange=(...args)=>{const r=get(...args);r.setNotes=v=>{notes=v;return r;};return r;};return s;};
 const t=template(),saved=save(e,[t]),sheet=e.sheets.get('採点テンプレ');
 assert.deepEqual(sheet.rows[0],['テンプレート名','枠1','枠2','枠3','枠4','枠5','出力1','出力2','出力3','出力1を表示','出力2を表示','出力3を表示','使用する']);
 assert.equal(notes[0].length,13);assert.ok(notes[0].every(s=>/[ぁ-んァ-ヶ一-龠]/.test(s)));
 assert.deepEqual(plain(saved.templates[0].scorePolicy),t.scorePolicy);assert.deepEqual(plain(saved.templates[0].rules[0].outputVisible),[true,false,true]);
 const result=e.c.scoringGetRules({templateId:t.name});assert.deepEqual(plain(result.scorePolicy),t.scorePolicy);assert.deepEqual(plain(result.rules.map(r=>r.visible)),[1,0,1]);
 for(const r of result.rules){assert.deepEqual(plain(r._scorePolicy),t.scorePolicy);assert.deepEqual(plain(r.outputVisible),[true,false,true]);assert.equal(r._rowNumber,2);assert.equal(r._tableDomains,undefined);}
 assert.deepEqual(plain(e.c.scoringPreviewTemplate({template:t,scores:[0]}).outputs),['A','B','C']);
 sheet.rows[1][9]='';sheet.rows[1][10]='';sheet.rows[1][11]='';sheet.rows[1][12]='';
 assert.deepEqual(plain(e.c.scoringGetTemplateTable().templates[0].rules[0].outputVisible),[true,true,true]);
});

test('new and directly entered templates default to unrestricted numeric inputs',()=>{
 const e=scoringEnvironment(),t={name:'新規',rules:[rule('>=-100')]};
 const normalized=e.c.scoringNormalizeTemplates_([t])[0];assert.deepEqual(plain(normalized.scorePolicy),any());
 save(e,[t]);e.sheets.get('採点テンプレ').rows.push(['手入力','','','','','','手入力の出力','','','','','','']);
 assert.deepEqual(plain(e.c.scoringGetTemplateTable().templates[1].scorePolicy),any());
 assert.equal(e.c.scoringPreviewTemplate({template:t,scores:[-1.25]}).outputs[0],'講評');
});

test('policy rejects malformed, empty, duplicate, nonfinite lists and legacy template fields',()=>{
 const e=scoringEnvironment();
 for(const policy of [[{mode:'unused'},...any().slice(1)],[{mode:'list',values:[]},...any().slice(1)],[{mode:'list',values:[0,0]},...any().slice(1)],[{mode:'list',values:[Infinity]},...any().slice(1)],[{mode:'list',values:['1']},...any().slice(1)],any().slice(1)])assert.throws(()=>e.c.scoringNormalizeTemplates_([{...template(),scorePolicy:policy}]),/制限|数値|重複|5|点数/);
 assert.throws(()=>e.c.scoringNormalizeTemplates_([{name:'旧版',domains:[[1],[],[],[],[]],rules:[]}]),/旧形式/);
 assert.throws(()=>e.c.scoringNormalizeTemplates_([{...template(),rules:[{...rule(),visible:false}]}]),/旧形式/);
});

test('preview checks finite numeric membership and preserves blank versus zero',()=>{
 const e=scoringEnvironment(),t=template();t.rules=[rule('0',['zero','','']),rule('',['blank','',''])];
 assert.equal(e.c.scoringPreviewTemplate({template:t,scores:['']}).outputs[0],'blank');assert.equal(e.c.scoringPreviewTemplate({template:t,scores:[0]}).outputs[0],'zero');
 for(const score of [1,Infinity,NaN,'bad',true])assert.throws(()=>e.c.scoringPreviewTemplate({template:t,scores:[score]}),/枠1|数値|点数/);
 t.scorePolicy=any();for(const score of [-100.125,0,1234.567])assert.doesNotThrow(()=>e.c.scoringPreviewTemplate({template:t,scores:[score]}));
 for(const score of [Infinity,NaN,'bad',true])assert.throws(()=>e.c.scoringPreviewTemplate({template:t,scores:[score]}),/枠1|数値/);
 assert.throws(()=>e.c.scoringPreviewTemplate({template:t,scores:{0:0}}),/点数|配列/);
});

test('unrestricted validation reports interval overlap and single-rule shadowing without claiming exhaustive coverage',()=>{
 const e=scoringEnvironment(),t={name:'自由',scorePolicy:any(),rules:[rule('>0'),rule('>=2'),rule('<=0',['','B',''])]};
 const result=e.c.scoringValidateTemplateTable({templates:[t]});assert.equal(result.ok,true);assert.equal(result.complete,false);
 assert.ok(result.issues.some(i=>i.code==='non-exhaustive'));assert.ok(result.issues.some(i=>i.code==='overlap'&&i.rule===2));assert.ok(result.issues.some(i=>i.code==='unreachable'&&i.rule===2));assert.ok(!result.issues.some(i=>i.code==='gap'));
 t.rules=[rule('<0'),rule('>=0')];assert.ok(!e.c.scoringValidateTemplateTable({templates:[t]}).issues.some(i=>i.code==='overlap'));
 t.rules=[rule('<=0'),rule('>=0')];assert.ok(e.c.scoringValidateTemplateTable({templates:[t]}).issues.some(i=>i.code==='overlap'&&i.example[0]===0));
 t.rules=[rule('<0'),rule('>=0'),rule('')];assert.ok(!e.c.scoringValidateTemplateTable({templates:[t]}).issues.some(i=>i.code==='unreachable'&&i.rule===3));
});

test('policy survives full-payload rename and clone, including names with punctuation',()=>{
 const e=scoringEnvironment(),t=template('物理/a:b');save(e,[t]);const renamed=plain(e.c.scoringGetTemplateTable().templates[0]);renamed.name='物理:a/b';
 const clone={...plain(renamed),name:'複製'};save(e,[renamed,clone]);const result=e.c.scoringGetTemplateTable();
 assert.deepEqual(plain(result.templates.map(t=>t.scorePolicy)),[t.scorePolicy,t.scorePolicy]);assert.equal(policyKeys(e).length,2);
});

test('nontrivial policies orphaned by direct sheet rename fail closed and are not silently reassigned',()=>{
 const e=scoringEnvironment(),t=template();save(e,[t]);const before=plain([...e.props]),sheet=e.sheets.get('採点テンプレ');sheet.rows[1][0]='直接改名';
 for(const fn of [()=>e.c.scoringGetTemplateTable(),()=>e.c.scoringGetRules({templateId:'直接改名'}),()=>save(e,[])])assert.throws(fn,/名前.*戻|改名|入力制限/);
 assert.deepEqual(plain([...e.props]),before);assert.equal(sheet.rows[1][0],'直接改名');sheet.rows[1][0]=t.name;assert.deepEqual(plain(e.c.scoringGetTemplateTable().templates[0].scorePolicy),t.scorePolicy);
});

test('score policy is operation scoped and changes invalidate the combined revision',()=>{
 const e=scoringEnvironment(),t=template(),first=save(e,[t]),key=policyKeys(e)[0],record=JSON.parse(e.props.get(key));record.scorePolicy[0].values=[0,2];e.props.set(key,JSON.stringify(record));
 assert.notEqual(e.c.scoringGetTemplateTable().revision,first.revision);assert.throws(()=>e.c.scoringSaveTemplateTable({revision:first.revision,templates:[t]}),/変更|再読込/);
 e.ss.getId=()=> 'other-operation';assert.deepEqual(plain(e.c.scoringGetTemplateTable().templates[0].scorePolicy),any());
});

test('empty templates expose scorePolicy to the scoring engine',()=>{
 const e=scoringEnvironment(),t=template();t.rules=[];save(e,[t]);const result=e.c.scoringGetRules({templateId:t.name});assert.equal(result.rules.length,0);assert.deepEqual(plain(result.scorePolicy),t.scorePolicy);
});

test('direct, management, and bundle version two packs retain score policy and display flags',()=>{
 const source=scoringEnvironment(),t=template();save(source,[t]);const pack=source.c.scoringExportTemplatePack();assert.equal(pack.meta.version,2);
 for(const kind of ['direct','templates','bundle']){
  const target=scoringEnvironment();
  if(kind==='direct')target.c.scoringImportTemplatePack(pack);
  else {const file=source.c.exportConfigurationFile(kind,false),inspection=target.c.inspectConfigurationFile(kind,file.json);target.c.applyConfigurationFile(kind,file.json,inspection.revision);}
  const read=target.c.scoringGetTemplateTable().templates[0];assert.deepEqual(plain(read.scorePolicy),t.scorePolicy);assert.deepEqual(plain(read.rules[0].outputVisible),t.rules[0].outputVisible);
 }
});

test('all import paths reject version one packs before changing sheets or properties',()=>{
 const e=scoringEnvironment(),t=template();save(e,[t]);const pack={meta:{type:'turret-grading-template-pack',version:1},templates:[]};
 const bundle=JSON.parse(e.c.exportConfigurationFile('bundle',false).json);bundle.templates=pack;const rows=plain(e.sheets.get('採点テンプレ').rows),props=plain([...e.props]);
 assert.throws(()=>e.c.scoringImportTemplatePack(pack),/旧形式.*バックアップ|version 2|バージョン2/);
 assert.throws(()=>e.c.inspectConfigurationFile('templates',JSON.stringify(pack)),/旧形式.*バックアップ|version 2|バージョン2/);
 assert.throws(()=>e.c.inspectConfigurationFile('bundle',JSON.stringify(bundle)),/旧形式.*バックアップ|version 2|バージョン2/);
 assert.deepEqual(e.sheets.get('採点テンプレ').rows,rows);assert.deepEqual(plain([...e.props]),props);
});

test('sheet or property writes that succeed then throw roll back both data stores',()=>{
 for(const failure of ['sheet','property']){
  const e=scoringEnvironment(),t=template();save(e,[t]);const before=e.c.scoringGetTemplateTable(),props=Object.fromEntries(e.props),sheet=e.sheets.get('採点テンプレ'),rows=plain(sheet.rows);let once=true;
  if(failure==='sheet'){const get=sheet.getRange.bind(sheet);sheet.getRange=(...args)=>{const r=get(...args),set=r.setValues;r.setValues=v=>{set(v);if(once){once=false;throw Error('partial sheet write');}return r;};return r;};}
  else {const api=e.c.PropertiesService.getScriptProperties(),set=api.setProperty;api.setProperty=(key,value)=>{set(key,value);if(key.startsWith('TURRET_TEMPLATE_POLICY_V2_')&&once){once=false;throw Error('partial policy write');}};e.c.PropertiesService.getScriptProperties=()=>api;}
  const next=template('新しい名前');next.scorePolicy[0].values=[-2,0,5];assert.throws(()=>e.c.scoringSaveTemplateTable({revision:before.revision,templates:[next]}),/partial/);
  assert.deepEqual(sheet.rows,rows);assert.deepEqual(Object.fromEntries(e.props),props);assert.equal(e.c.scoringGetTemplateTable().revision,before.revision);assert.equal(e.isLocked(),false);
 }
});

test('silent property write failure is detected and rolled back',()=>{
 const e=scoringEnvironment(),t=template();save(e,[t]);const before=e.c.scoringGetTemplateTable(),sheet=e.sheets.get('採点テンプレ'),rows=plain(sheet.rows),props=Object.fromEntries(e.props),api=e.c.PropertiesService.getScriptProperties(),set=api.setProperty;let once=true;
 api.setProperty=(key,value)=>{if(key.startsWith('TURRET_TEMPLATE_POLICY_V2_')&&once){once=false;return;}set(key,value);};e.c.PropertiesService.getScriptProperties=()=>api;
 t.scorePolicy[0].values=[2];assert.throws(()=>e.c.scoringSaveTemplateTable({revision:before.revision,templates:[t]}),/保存.*確認|一致/);assert.deepEqual(sheet.rows,rows);assert.deepEqual(Object.fromEntries(e.props),props);
});

test('property quota preflight leaves the current sheet intact',()=>{
 const e=scoringEnvironment(),t=template();save(e,[t]);const before=e.c.scoringGetTemplateTable(),rows=plain(e.sheets.get('採点テンプレ').rows);e.props.set('unrelated','x'.repeat(470000));
 assert.throws(()=>e.c.scoringSaveTemplateTable({revision:before.revision,templates:[template('別名')]}),/容量|バイト/);assert.deepEqual(e.sheets.get('採点テンプレ').rows,rows);
});

test('version two packs require explicit portable policies and output visibility',()=>{
 const e=scoringEnvironment();for(const omit of [t=>delete t.scorePolicy,t=>delete t.rules[0].outputVisible]){
  const t=template();omit(t);const pack={meta:{type:'turret-grading-template-pack',version:2},templates:[t]};
  assert.throws(()=>e.c.scoringImportTemplatePack(pack),/入力制限|表示/);
  assert.throws(()=>e.c.inspectConfigurationFile('templates',JSON.stringify(pack)),/入力制限|表示/);
 }
 assert.equal(e.sheets.size,0);
});

test('property peak quota includes old keys retained during a rename',()=>{
 const e=scoringEnvironment(),t=template();save(e,[t]);
 const api=e.c.PropertiesService.getScriptProperties(),all=api.getProperties(),bytes=e.c.utf8ByteLength_;
 const current=Object.keys(all).reduce((n,k)=>n+bytes(k)+bytes(all[k]),0);
 e.props.set('padding','x'.repeat(449950-current-bytes('padding')));
 const rows=plain(e.sheets.get('採点テンプレ').rows);assert.throws(()=>save(e,[template('新しい制限名')]),/容量|バイト/);assert.deepEqual(e.sheets.get('採点テンプレ').rows,rows);
});

test('actual header notes are written to every one of the thirteen columns',()=>{
 const e=scoringEnvironment();save(e,[template()]);const notes=e.sheets.get('採点テンプレ').getRange(1,1,1,13).getNotes()[0];
 assert.equal(notes.length,13);assert.ok(notes.every(n=>n.length>10));
});

test('deleting a policy that succeeds then throws restores the deleted template and policy',()=>{
 const e=scoringEnvironment();save(e,[template()]);const before=e.c.scoringGetTemplateTable(),props=Object.fromEntries(e.props),rows=plain(e.sheets.get('採点テンプレ').rows),api=e.c.PropertiesService.getScriptProperties(),remove=api.deleteProperty;let once=true;
 api.deleteProperty=k=>{remove(k);if(k.startsWith('TURRET_TEMPLATE_POLICY_V2_')&&once){once=false;throw Error('partial delete');}};e.c.PropertiesService.getScriptProperties=()=>api;
 assert.throws(()=>e.c.scoringSaveTemplateTable({revision:before.revision,templates:[]}),/partial delete/);assert.equal(e.c.scoringGetTemplateTable().revision,before.revision);assert.deepEqual(Object.fromEntries(e.props),props);assert.deepEqual(e.sheets.get('採点テンプレ').rows.slice(0,rows.length),rows);
});

test('failed first policy write removes the newly-created template sheet and restores properties',()=>{
 const e=scoringEnvironment(),before=Object.fromEntries(e.props),api=e.c.PropertiesService.getScriptProperties(),set=api.setProperty;let once=true;
 api.setProperty=(k,v)=>{set(k,v);if(k.startsWith('TURRET_TEMPLATE_POLICY_V2_')&&once){once=false;throw Error('first policy failure');}};e.c.PropertiesService.getScriptProperties=()=>api;
 assert.throws(()=>save(e,[template()]),/first policy failure/);assert.equal(e.sheets.has('採点テンプレ'),false);assert.deepEqual(Object.fromEntries(e.props),before);
});

test('corrupt or mismatched policy records fail closed without falling back to unrestricted',()=>{
 for(const change of [()=>'{broken',value=>JSON.stringify({...JSON.parse(value),name:'別名'}),value=>JSON.stringify({...JSON.parse(value),scorePolicy:undefined})]){
  const e=scoringEnvironment();save(e,[template()]);const key=policyKeys(e)[0];e.props.set(key,change(e.props.get(key)));assert.throws(()=>e.c.scoringGetTemplateTable(),/入力制限/);
 }
});

test('a failed rollback blocks subsequent grading until the stored inconsistency is repaired',()=>{
 const e=scoringEnvironment();save(e,[template()]);const before=e.c.scoringGetTemplateTable(),api=e.c.PropertiesService.getScriptProperties(),set=api.setProperty;
 api.setProperty=(k,v)=>{if(k.startsWith('TURRET_TEMPLATE_POLICY_V2_')){if(JSON.parse(v).scorePolicy[0].values[0]===99){set(k,v);throw Error('saved then failed');}throw Error('restore denied');}set(k,v);};e.c.PropertiesService.getScriptProperties=()=>api;
 const next=template();next.scorePolicy[0].values=[99];assert.throws(()=>e.c.scoringSaveTemplateTable({revision:before.revision,templates:[next]}),/復旧も完了していません/);assert.throws(()=>e.c.scoringGetRules({templateId:next.name}),/復旧が未完了|復旧が.*未完了|復旧が|復旧.*未完了/);assert.equal(e.isLocked(),false);
});

test('preview normalizes numeric spelling exactly as the scoring engine does',()=>{
 const e=scoringEnvironment(),t={name:'表記',scorePolicy:any(),rules:[rule('',['{枠1}/{枠2}/{枠3}/{合計}','',''])]};
 assert.equal(e.c.scoringPreviewTemplate({template:t,scores:['01','-1.50','']}).outputs[0],'1/-1.5//-0.5');
});

test('a concurrent grading read is blocked while sheet and policies are being saved',()=>{
 const e=scoringEnvironment();save(e,[template()]);let checked=false;const sheet=e.sheets.get('採点テンプレ');
 sheet.beforeWrite=()=>{checked=true;assert.throws(()=>e.c.scoringGetTemplateTable(),/別の処理|保存中|保存処理中|復旧/);assert.throws(()=>e.c.scoringGetRules({templateId:'化学'}),/別の処理|保存中|保存処理中|復旧/);};
 const next=template('新しい名前');save(e,[next]);assert.equal(checked,true);assert.deepEqual(plain(e.c.scoringGetRules({templateId:next.name}).scorePolicy),next.scorePolicy);assert.ok(![...e.props.keys()].some(k=>k.startsWith('TURRET_TEMPLATE_RECOVERY_V2_')));
});

test('an interrupted save marker blocks external reads before any policy can default to unrestricted',()=>{
 const e=scoringEnvironment();save(e,[template()]);e.props.set(e.c.scoringTemplateRecoveryKey_(),'保存処理中');e.sheets.get('採点テンプレ').rows[1][0]='まだ制限がない新しい名前';
 assert.throws(()=>e.c.scoringGetRules({templateId:'まだ制限がない新しい名前'}),/別の処理|保存中|保存処理中|復旧/);
});

test('preview accepts already-numeric finite values even when JavaScript prints exponent notation',()=>{
 const e=scoringEnvironment(),t={name:'数値',scorePolicy:any(),rules:[rule('',['{枠1}','',''])]};
 for(const n of [1e21,1e-7,-1e308])assert.equal(e.c.scoringPreviewTemplate({template:t,scores:[n]}).outputs[0],String(n));
});

test('a direct sheet edit made while setting the write marker is retained rather than overwritten or rolled back',()=>{
 const e=scoringEnvironment();save(e,[template()]);const before=e.c.scoringGetTemplateTable(),api=e.c.PropertiesService.getScriptProperties(),set=api.setProperty,sheet=e.sheets.get('採点テンプレ');let once=true;
 api.setProperty=(k,v)=>{set(k,v);if(k===e.c.scoringTemplateRecoveryKey_()&&once){once=false;sheet.rows[1][6]='直前の同時編集';}};e.c.PropertiesService.getScriptProperties=()=>api;
 assert.throws(()=>e.c.scoringSaveTemplateTable({revision:before.revision,templates:[template()]}),/変更|再読込/);assert.equal(sheet.rows[1][6],'直前の同時編集');assert.ok(!e.props.has(e.c.scoringTemplateRecoveryKey_()));
});

test('scientific notation survives policy pack and preview round trips while overflow is rejected',()=>{
 const e=scoringEnvironment(),t=template();t.scorePolicy[0].values=[1e-7,1e21];t.rules=[rule('',['{枠1}','',''])];save(e,[t]);
 const pack=e.c.scoringExportTemplatePack(),target=scoringEnvironment();target.c.scoringImportTemplatePack(pack);const imported=target.c.scoringGetTemplateTable().templates[0];
 assert.equal(target.c.scoringPreviewTemplate({template:imported,scores:['1e-7']}).outputs[0],'1e-7');assert.equal(target.c.scoringPreviewTemplate({template:imported,scores:['1E+21']}).outputs[0],'1e+21');
 for(const s of ['1e999','-1e999','1e','0x10'])assert.throws(()=>target.c.scoringPreviewTemplate({template:imported,scores:[s]}),/有限の数値/);
});

test('every public template read holds the write lock across both sheet and policy reads',()=>{
 for(const endpoint of ['table','rules','list','direct-pack','configuration-pack','bundle']){
  const e=scoringEnvironment(),t=template();save(e,[t]);const before=e.c.scoringGetTemplateTable(),sheet=e.sheets.get('採点テンプレ'),get=sheet.getRange.bind(sheet);let attempted=false;
  sheet.getRange=(...args)=>{const r=get(...args),display=r.getDisplayValues;r.getDisplayValues=()=>{const values=display();if(!attempted){attempted=true;const next=template();next.scorePolicy[0].values=[99];next.rules[0].outputs[0]='new';assert.throws(()=>e.c.scoringSaveTemplateTable({revision:before.revision,templates:[next]}),/処理|実行|ロック/);}return values;};return r;};
  const result=endpoint==='table'?e.c.scoringGetTemplateTable():endpoint==='rules'?e.c.scoringGetRules({templateId:t.name}):endpoint==='list'?e.c.scoringGetRuleTemplates():endpoint==='direct-pack'?e.c.scoringExportTemplatePack():e.c.exportConfigurationFile(endpoint==='bundle'?'bundle':'templates',false);
  assert.equal(attempted,true);assert.ok(result);assert.equal(e.c.scoringGetTemplateTable().revision,before.revision);assert.equal(e.isLocked(),false);
 }
});
