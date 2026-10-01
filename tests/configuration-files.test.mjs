import test from 'node:test';
import assert from 'node:assert/strict';
import {scoringEnvironment, AdminSheet} from './helpers/scoring-environment.mjs';
import {plain} from './helpers/admin-environment.mjs';

const template = name => ({name, domains:[[1,2,3],[],[],[],[]], rules:[{conditions:['>=2','','','',''],outputs:['よくできました','',''],enabled:true,visible:true}]});
function configured() {
  const e=scoringEnvironment();
  e.c.scoringSaveTemplateTable({revision:e.c.scoringGetTemplateTable().revision,templates:[template('化学')]});
  e.c.scoringSetConfig({scoreCols:['D','','','',''],commentCol:'E',ruleTemplateId:'化学'});
  return e;
}
const inspect = (e,kind,json) => e.c.inspectConfigurationFile(kind,json);
const apply = (e,kind,json) => e.c.applyConfigurationFile(kind,json,inspect(e,kind,json).revision);

test('automatic inspection recognizes all four exports and legacy management and unwrapped scoring without writes',()=>{
  const source=configured(),target=scoringEnvironment(),before=plain([...target.props]);
  for(const kind of ['settings','scoring','templates','bundle']) {
    const file=source.c.exportConfigurationFile(kind,false);
    const automatic=inspect(target,'auto',file.json),explicit=inspect(target,kind,file.json);
    assert.equal(automatic.kind,kind);assert.equal(automatic.revision,explicit.revision);
  }
  const scoring=JSON.parse(source.c.exportConfigurationFile('scoring').json);
  assert.equal(inspect(target,'auto',JSON.stringify(scoring.config)).kind,'scoring');
  const management=JSON.parse(source.c.exportConfigurationFile('settings',false).json);
  management.schemaVersion=2;
  assert.equal(inspect(target,'auto',JSON.stringify(management)).kind,'settings');
  assert.deepEqual(plain([...target.props]),before);assert.equal(target.sheets.size,0);
});

test('automatic inspection rejects unknown or conflicting identifiers rather than guessing a configuration type',()=>{
  const e=scoringEnvironment(),source=configured(),p=JSON.parse(source.c.exportConfigurationFile('scoring').json);
  for(const payload of [{},[],{...p,meta:{type:'unrelated'}},{...p,application:'other'},{...p,application:'turret',schemaVersion:3},{...p,type:'unrelated'},{templates:[]}]) {
    assert.throws(()=>inspect(e,'auto',JSON.stringify(payload)),/種類|形式|オブジェクト/);
  }
  assert.throws(()=>inspect(e,'auto','broken'),/JSON/);
  assert.equal(e.props.has('TURRET_SCORING_CONFIG'),false);assert.equal(e.sheets.size,0);
});

test('automatic inspection applies the detected management size limit and rejects incomplete bundles',()=>{
  const e=scoringEnvironment(),source=configured(),p=JSON.parse(source.c.exportConfigurationFile('settings',false).json);
  p.config.messageTemplate='x'.repeat(131072);
  assert.throws(()=>inspect(e,'auto',JSON.stringify(p)),/128/);
  assert.throws(()=>inspect(e,'auto',' '.repeat(1048577)),/1MiB/);
  const bundle=JSON.parse(source.c.exportConfigurationFile('bundle',false).json);delete bundle.templates;
  assert.throws(()=>inspect(e,'auto',JSON.stringify(bundle)));
});

test('all file exports use a safe operation name and Japan time including the date boundary',()=>{
  const e=configured();e.ss.getName=()=> '化学/1学期:<試験>?';
  const date=new Date('2026-09-30T15:30:00Z');
  assert.equal(e.c.configurationFilename_('bundle',date),'turret-bundle-化学_1学期_試験_-20261001-003000.json');
  for(const kind of ['settings','scoring','templates','bundle']) {
    const result=e.c.exportConfigurationFile(kind,false);
    assert.match(result.filename,new RegExp('^turret-'+kind+'-化学_1学期_試験_-\\d{8}-\\d{6}\\.json$'));
    assert.ok(JSON.parse(result.json));
  }
  assert.match(e.c.exportSettingsProfile(false).filename,/^turret-settings-化学_1学期_試験_-/);
});

test('bundle exports only the three configuration sections and round trips onto another operation',()=>{
  const source=configured(),file=source.c.exportConfigurationFile('bundle',false),p=JSON.parse(file.json);
  assert.deepEqual(Object.keys(p).sort(),['application','exportedAt','management','schemaVersion','scoring','templates','type']);
  assert.equal(p.type,'configuration-bundle');assert.equal(p.management.includesConnections,false);
  assert.equal(p.management.config.formSources,undefined);assert.equal(p.management.schedule.enabled,undefined);
  const target=scoringEnvironment();target.c.scoringSaveTemplateTable({revision:target.c.scoringGetTemplateTable().revision,templates:[template('旧テンプレ')]});
  target.configure({reminderTo:['target@example.invalid'],formSources:[]});
  const answers=plain(target.answer.rows),preview=inspect(target,'bundle',file.json);
  assert.match(preview.summary,/置き換|入れ替/);assert.equal(target.c.scoringGetTemplateTable().templates[0].name,'旧テンプレ');
  const result=target.c.applyConfigurationFile('bundle',file.json,preview.revision);
  assert.match(result.message,/適用/);assert.equal(target.c.scoringGetConfig().ruleTemplateId,'化学');
  assert.deepEqual(plain(target.c.scoringGetTemplateTable().templates.map(t=>t.name)),['化学']);
  assert.deepEqual(plain(target.c.getConfig_().formSources),[]);
  assert.deepEqual(plain(target.c.getConfig_().reminderTo),['target@example.invalid']);
  assert.deepEqual(target.answer.rows,answers);assert.equal(target.triggers.length,0);assert.equal(target.isLocked(),false);
});

test('template form URL and notification mail can be exported and applied independently',()=>{
  const source=configured(),target=scoringEnvironment();
  const url='https://docs.google.com/forms/d/template_12345678901234567890/edit';
  source.props.set('TURRET_TEMPLATE_FORM_URL',url);
  source.configure({reminderTo:['source@example.invalid']});
  target.configure({reminderTo:['target@example.invalid']});
  for (const kind of ['settings','bundle']) {
    const template=JSON.parse(source.c.exportConfigurationFile(kind,{templateFormUrl:true,reminderTo:false}).json);
    const mail=JSON.parse(source.c.exportConfigurationFile(kind,{templateFormUrl:false,reminderTo:true}).json);
    const templateProfile=kind==='bundle'?template.management:template;
    const mailProfile=kind==='bundle'?mail.management:mail;
    assert.equal(templateProfile.schemaVersion,4);
    assert.equal(templateProfile.templateFormUrl,url);
    assert.equal(templateProfile.config.reminderTo,undefined);
    assert.equal(templateProfile.config.formSources,undefined);
    assert.equal(mailProfile.templateFormUrl,undefined);
    assert.deepEqual(plain(mailProfile.config.reminderTo),['source@example.invalid']);
  }
  const templateJson=source.c.exportConfigurationFile('settings',{templateFormUrl:true,reminderTo:false}).json;
  const preview=inspect(target,'settings',templateJson);
  assert.match(preview.summary,/ひな形フォームのURL/);
  assert.doesNotMatch(preview.summary,/通知先メールも置き換え/);
  target.c.applyConfigurationFile('settings',templateJson,preview.revision);
  assert.equal(JSON.parse(target.props.get('TURRET_TEMPLATE_FORM_URL')).url,url);
  assert.deepEqual(plain(target.c.getConfig_().reminderTo),['target@example.invalid']);
  const mailJson=source.c.exportConfigurationFile('settings',{templateFormUrl:false,reminderTo:true}).json;
  apply(target,'settings',mailJson);
  assert.equal(JSON.parse(target.props.get('TURRET_TEMPLATE_FORM_URL')).url,url);
  assert.deepEqual(plain(target.c.getConfig_().reminderTo),['source@example.invalid']);
});

test('invalid template URL is rejected before import changes settings',()=>{
  const e=configured(),profile=JSON.parse(e.c.exportConfigurationFile('settings',{templateFormUrl:false,reminderTo:false}).json);
  profile.templateFormUrl='https://example.com/not-a-form';
  const before=plain([...e.props]);
  assert.throws(()=>inspect(e,'settings',JSON.stringify(profile)),/ひな形フォーム/);
  assert.deepEqual(plain([...e.props]),before);
});

test('a new operation exports importable defaults without saving scoring settings or creating template sheets',()=>{
  const source=scoringEnvironment(),before=plain([...source.props]);
  const scoring=source.c.exportConfigurationFile('scoring').json;
  assert.doesNotThrow(()=>inspect(source,'scoring',scoring));
  const bundle=source.c.exportConfigurationFile('bundle',false).json;
  assert.doesNotThrow(()=>inspect(source,'bundle',bundle));
  assert.deepEqual(plain([...source.props]),before);assert.equal(source.sheets.size,0);
  const target=scoringEnvironment();apply(target,'bundle',bundle);
  assert.deepEqual(plain(target.c.scoringGetConfig().displayCols),['A','B','C','','']);
  assert.equal(target.c.scoringGetTemplateTable().templates.length,0);assert.equal(target.triggers.length,0);
});

test('individual management and scoring exports preserve their existing formats and scope',()=>{
  const source=configured(),target=scoringEnvironment();
  const before=target.props.get('APP_CONFIG');
  apply(target,'scoring',source.c.exportConfigurationFile('scoring').json);
  assert.equal(target.c.scoringGetConfig().scoreCols[0],'D');assert.equal(target.props.get('APP_CONFIG'),before);
  const scoringBefore=target.props.get('TURRET_SCORING_CONFIG');
  apply(target,'settings',source.c.exportConfigurationFile('settings',false).json);
  assert.equal(target.props.get('TURRET_SCORING_CONFIG'),scoringBefore);
  const old=JSON.parse(source.c.exportConfigurationFile('scoring').json);delete old.config.lessonDateHeader;old.config.sheetName='旧回答';
  apply(target,'scoring',JSON.stringify(old));assert.equal(target.c.scoringGetConfig().sheetName,'');
});

test('individual packs append, refuse duplicates and preserve grading configuration',()=>{
  const source=configured(),target=scoringEnvironment(),file=source.c.exportConfigurationFile('templates');
  target.c.scoringSaveTemplateTable({revision:target.c.scoringGetTemplateTable().revision,templates:[template('物理')]});
  const before=target.props.get('TURRET_SCORING_CONFIG');apply(target,'templates',file.json);
  assert.deepEqual(plain(target.c.scoringGetTemplateTable().templates.map(t=>t.name)),['物理','化学']);
  assert.equal(target.props.get('TURRET_SCORING_CONFIG'),before);
  assert.throws(()=>inspect(target,'templates',file.json),/同名/);
});

test('inspection is read only and rejects wrong kinds, missing bundle sections, invalid rules and unresolved templates',()=>{
  const e=configured(),file=e.c.exportConfigurationFile('bundle',false),before=plain([...e.props]),rows=plain(e.sheets.get('採点テンプレ').rows);
  assert.throws(()=>inspect(e,'scoring',file.json),/形式|項目|許可/);
  for(const modify of [p=>delete p.scoring,p=>p.schemaVersion=99,p=>p.templates.templates[0].rules[0].conditions[0]='oops',p=>p.scoring.config.ruleTemplateId='不存在',p=>p.scoring.config.scoreCols=['D','D','','','']]) {
    const p=JSON.parse(file.json);modify(p);assert.throws(()=>inspect(e,'bundle',JSON.stringify(p)));
  }
  assert.throws(()=>inspect(e,'bundle','あ'.repeat(350000)),/1MiB|1024KB/);
  assert.deepEqual(plain([...e.props]),before);assert.deepEqual(e.sheets.get('採点テンプレ').rows,rows);
});

test('file inspection is bound to its content and catches changes to all destination sections',()=>{
  for(const change of [e=>e.c.scoringSetConfig({scoreCols:['E','','','','']}),e=>e.sheets.get('採点テンプレ').rows[1][7]='別の編集',e=>e.props.set('APP_CONFIG',JSON.stringify({...e.c.getConfig_(),messageTemplate:'別の編集'})),e=>e.schedule.deliveryHour=9,e=>e.ss.getId=()=> 'other-operation']) {
    const e=configured(),json=e.c.exportConfigurationFile('bundle',false).json,preview=inspect(e,'bundle',json);change(e);
    assert.throws(()=>e.c.applyConfigurationFile('bundle',json,preview.revision),/変更|再.*込/);assert.equal(e.isLocked(),false);
  }
  const e=configured(),json=e.c.exportConfigurationFile('scoring').json,preview=inspect(e,'scoring',json),p=JSON.parse(json);p.config.startRow=10;
  assert.throws(()=>e.c.applyConfigurationFile('scoring',JSON.stringify(p),preview.revision),/変更|再.*込/);
});

test('bundle requires automation stopped and preflights incompatible queues before any changes',()=>{
  const e=configured(),p=JSON.parse(e.c.exportConfigurationFile('bundle',false).json);p.management.config.messageTemplate='新本文';
  const before=plain([...e.props]),templates=plain(e.sheets.get('採点テンプレ').rows);
  e.c.getAutomationStateUnlocked_=()=>({enabled:true,legacyCount:0});
  assert.throws(()=>apply(e,'bundle',JSON.stringify(p)),/停止/);
  e.c.getAutomationStateUnlocked_=()=>({enabled:false,legacyCount:1});assert.throws(()=>apply(e,'bundle',JSON.stringify(p)),/停止/);
  e.c.getAutomationStateUnlocked_=()=>({enabled:false,legacyCount:0});
  e.sheets.set('送信シート',new AdminSheet('送信シート',[['違う構成'],['保持する送信行']]));
  assert.throws(()=>apply(e,'bundle',JSON.stringify(p)),/列構成/);
  assert.deepEqual(plain([...e.props]),before);assert.deepEqual(e.sheets.get('採点テンプレ').rows,templates);
});

test('bundle restores every section when an external write fails, including a write that succeeded then threw',()=>{
  for(const failure of ['APP_CONFIG','TURRET_SCORING_CONFIG','templates']) {
    const e=configured(),p=JSON.parse(e.c.exportConfigurationFile('bundle',true).json);
    p.management.schemaVersion=4;p.management.includesConnections=false;delete p.management.config.formSources;
    p.management.templateFormUrl='https://docs.google.com/forms/d/template_12345678901234567890/edit';
    p.management.config.messageTemplate='新本文';p.management.schedule.deliveryHour=10;p.scoring.config.startRow=9;p.templates.templates=[template('新テンプレ')];p.scoring.config.ruleTemplateId='新テンプレ';
    const before=plain([...e.props]),rows=plain(e.sheets.get('採点テンプレ').rows),schedule=plain(e.schedule);
    let once=true;
    if(failure==='templates') {const sheet=e.sheets.get('採点テンプレ'),get=sheet.getRange.bind(sheet);sheet.getRange=(...args)=>{const range=get(...args),set=range.setValues;range.setValues=values=>{set(values);if(once){once=false;throw Error('injected write failure');}return range;};return range;};}
    else {const props=e.c.PropertiesService.getScriptProperties(),set=props.setProperty;props.setProperty=(key,value)=>{set(key,value);if(key===failure&&once){once=false;throw Error('injected write failure');}};e.c.PropertiesService.getScriptProperties=()=>props;}
    assert.throws(()=>apply(e,'bundle',JSON.stringify(p)),/injected/);
    assert.deepEqual(Object.fromEntries(e.props),Object.fromEntries(before));assert.deepEqual(e.sheets.get('採点テンプレ').rows,rows);assert.deepEqual(e.schedule,schedule);assert.equal(e.isLocked(),false);
  }
});

test('configuration file endpoints enforce the deployed operator identity',()=>{
  const e=configured(),json=e.c.exportConfigurationFile('scoring').json;
  e.c.Session.getActiveUser=()=>({getEmail:()=> 'intruder@example.com'});
  assert.throws(()=>e.c.exportConfigurationFile('scoring'),/本人/);
  assert.throws(()=>inspect(e,'scoring',json),/本人/);
  assert.throws(()=>e.c.applyConfigurationFile('scoring',json,'old'),/本人/);
});

test('failed template import restores trusted sheet formulas as formulas and preserves literal equals text',()=>{
  const e=configured(),sheet=e.sheets.get('採点テンプレ');
  sheet.rows[1][6]='="数式の講評"';sheet.rows[1][7]="'=文字の講評";
  const get=sheet.getRange.bind(sheet);let once=true;
  sheet.getRange=(...args)=>{
    const range=get(...args),display=range.getDisplayValues,set=range.setValues;
    range.getDisplayValues=()=>display().map(row=>row.map(v=>v==='="数式の講評"'?'数式の講評':v));
    range.setValues=values=>{set(values);if(once){once=false;throw Error('injected template write failure');}return range;};
    return range;
  };
  const rows=plain(sheet.rows),json=JSON.stringify({meta:{type:'turret-grading-template-pack',version:1},templates:[template('物理')]});
  assert.throws(()=>apply(e,'templates',json),/injected template write failure/);
  assert.deepEqual(sheet.rows.slice(0,2),rows);assert.equal(e.props.has('APP_CONFIG_SAVE_ERROR'),false);
  assert.equal(e.c.scoringReadTemplateTable_().templates[0].rules[0].outputs[0],'数式の講評');
});

test('a concurrent direct template edit during property saves is rejected and never rolled back',()=>{
  const e=configured(),json=e.c.exportConfigurationFile('bundle',false).json,preview=inspect(e,'bundle',json);
  const before=plain([...e.props]),sheet=e.sheets.get('採点テンプレ'),props=e.c.PropertiesService.getScriptProperties(),set=props.setProperty;
  let edited=false;
  props.setProperty=(key,value)=>{set(key,value);if(key==='TURRET_SCORING_CONFIG'&&!edited){edited=true;sheet.rows[1][6]='同時編集した講評';}};
  e.c.PropertiesService.getScriptProperties=()=>props;
  assert.throws(()=>e.c.applyConfigurationFile('bundle',json,preview.revision),/シートが変更/);
  assert.equal(sheet.rows[1][6],'同時編集した講評');assert.deepEqual(Object.fromEntries(e.props),Object.fromEntries(before));assert.equal(e.isLocked(),false);
});

test('real automation record is preserved byte for byte on failed bundle apply and no triggers start',()=>{
  const e=scoringEnvironment({actualAutomation:true});
  e.userProps.set('APP_AUTOMATION_OWNER','manager-token');
  const oldRecord=JSON.stringify({revision:'original',owner:'manager-token',enabled:false,active:[],activeSchedule:null,schedule:{importHour:5,deliveryHour:8,reminderHour:16,reminderEnabled:false}});
  e.props.set('APP_AUTOMATION_STATE',oldRecord);
  const p=JSON.parse(e.c.exportConfigurationFile('bundle',false).json);p.management.schedule.deliveryHour=10;
  const props=e.c.PropertiesService.getScriptProperties(),set=props.setProperty;let once=true;
  props.setProperty=(key,value)=>{set(key,value);if(key==='TURRET_SCORING_CONFIG'&&once){once=false;throw Error('injected scoring failure');}};
  e.c.PropertiesService.getScriptProperties=()=>props;
  assert.throws(()=>apply(e,'bundle',JSON.stringify(p)),/injected scoring failure/);
  assert.equal(e.props.get('APP_AUTOMATION_STATE'),oldRecord);assert.equal(e.triggers.length,0);assert.equal(e.isLocked(),false);
});

test('bundle cannot modify automation owned by another manager even while stopped',()=>{
  const e=scoringEnvironment({actualAutomation:true});
  e.props.set('APP_AUTOMATION_STATE',JSON.stringify({revision:'original',owner:'different-manager',enabled:false,active:[],activeSchedule:null,schedule:{importHour:5,deliveryHour:8,reminderHour:16,reminderEnabled:false}}));
  const json=e.c.exportConfigurationFile('bundle',false).json,before=plain([...e.props]);
  assert.throws(()=>apply(e,'bundle',json),/管理者/);assert.deepEqual(plain([...e.props]),before);
});

test('unrecoverable partial configuration writes report recovery needed and prevent runtime configuration use',()=>{
  const e=configured(),p=JSON.parse(e.c.exportConfigurationFile('bundle',false).json);p.scoring.config.startRow=9;
  const props=e.c.PropertiesService.getScriptProperties(),set=props.setProperty;
  props.setProperty=(key,value)=>{
    if(key==='TURRET_SCORING_CONFIG') {if(value.includes('"startRow":9')){set(key,value);throw Error('save failed');}throw Error('restore denied');}
    set(key,value);
  };e.c.PropertiesService.getScriptProperties=()=>props;
  assert.throws(()=>apply(e,'bundle',JSON.stringify(p)),/復旧も完了していません/);
  assert.ok(e.props.get('APP_CONFIG_SAVE_ERROR'));assert.throws(()=>e.c.getConfig_(),/復旧が未完了/);assert.equal(e.isLocked(),false);
});
