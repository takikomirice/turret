import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { adminEnvironment, plain } from './helpers/admin-environment.mjs';

test('optional assessment example imports explicitly while new configuration stays empty',()=>{
  const e=adminEnvironment({config:null});
  const initial=plain(e.c.getConfig_());
  assert.equal(initial.messageTemplate,'');assert.deepEqual(initial.fields,[]);
  const profile=e.c.inspectSettingsProfile(readFileSync('examples/profiles/assessment.json','utf8'));
  assert.equal(profile.config.fields.length,5);assert.ok(profile.config.messageTemplate);
  assert.deepEqual(plain(e.c.getConfig_()),initial);
});

test('profile byte limit counts multibyte text',()=>{
  const e=adminEnvironment();
  assert.throws(()=>e.c.inspectSettingsProfile('あ'.repeat(50000)),/128KB/);
});

test('portable export omits connections and always excludes data and trigger activation',()=>{
  const e=adminEnvironment(); const result=e.c.exportSettingsProfile(false),profile=JSON.parse(result.json);
  assert.equal(profile.application,'turret'); assert.equal(profile.schemaVersion,3);
  assert.equal(profile.includesConnections,false);assert.equal(profile.config.formSources,undefined);assert.equal(profile.config.reminderTo,undefined);
  assert.equal(profile.config.messageTemplate,'評価：＜score＞');
  assert.equal(profile.schedule.deliveryHour,8);assert.equal(profile.schedule.enabled,undefined);
  assert.equal(profile.students,undefined);assert.equal(profile.triggers,undefined);
});

test('full export includes connections only when explicitly selected',()=>{
  const e=adminEnvironment();const profile=JSON.parse(e.c.exportSettingsProfile(true).json);
  assert.equal(profile.includesConnections,true); assert.equal(profile.config.formSources.length,1);assert.equal(profile.config.reminderTo.length,1);
});

test('inspection of exported JSON is read-only and validates supported version and field types',()=>{
  const e=adminEnvironment();const json=e.c.exportSettingsProfile(false).json,before=plain(Object.fromEntries(e.props));
  assert.ok(e.c.inspectSettingsProfile(json).summary);
  assert.deepEqual(plain(Object.fromEntries(e.props)),before);
  const p=JSON.parse(json);p.schemaVersion=99; assert.throws(()=>e.c.inspectSettingsProfile(JSON.stringify(p)),/バージョン/);
  p.schemaVersion=2;p.config.fields={};assert.throws(()=>e.c.inspectSettingsProfile(JSON.stringify(p)),/項目|fields|配列/);
  assert.throws(()=>e.c.inspectSettingsProfile('{'),/JSON/);
});

test('common grade scale round trips through v3 and v2 while v1 restores the legacy scale explicitly',()=>{
  const e=adminEnvironment();const config=e.c.getConfig_();config.gradeScale=[{from:'合格',to:'達成'}];e.props.set('APP_CONFIG',JSON.stringify(config));
  const exported=JSON.parse(e.c.exportSettingsProfile(false).json);
  assert.deepEqual(exported.config.gradeScale,[{from:'合格',to:'達成'}]);
  assert.deepEqual(plain(e.c.inspectSettingsProfile(JSON.stringify(exported)).config.gradeScale),exported.config.gradeScale);
  exported.schemaVersion=2;
  assert.deepEqual(plain(e.c.inspectSettingsProfile(JSON.stringify(exported)).config.gradeScale),exported.config.gradeScale);
  exported.schemaVersion=1;delete exported.config.gradeScale;
  const imported=e.c.inspectSettingsProfile(JSON.stringify(exported));
  assert.equal(imported.config.gradeScale.length,5);assert.deepEqual(plain(imported.config.gradeScale[0]),{from:'5',to:'A'});
  exported.config.fields[0].type='text';
  assert.deepEqual(plain(e.c.inspectSettingsProfile(JSON.stringify(exported)).config.gradeScale),[]);
  exported.schemaVersion=2;assert.throws(()=>e.c.inspectSettingsProfile(JSON.stringify(exported)),/gradeScale/);
});

test('v3 profiles preserve separate field conversions and explicit empty overrides',()=>{
  const e=adminEnvironment(),config=e.c.getConfig_();
  config.fields[0].gradeScale=[{from:'1',to:'×'},{from:'2',to:'△'},{from:'3',to:'○'}];
  config.fields.push({...config.fields[0],key:'other',sourceHeader:'別評価',evalHeader:'別評価',sendHeader:'別評価',gradeScale:[]});
  e.props.set('APP_CONFIG',JSON.stringify(config));
  const exported=JSON.parse(e.c.exportSettingsProfile(false).json);
  const inspected=plain(e.c.inspectSettingsProfile(JSON.stringify(exported)).config);
  assert.deepEqual(inspected.fields,exported.config.fields);
  assert.deepEqual(inspected.fields[1].gradeScale,[]);
  exported.schemaVersion=2;
  assert.throws(()=>e.c.inspectSettingsProfile(JSON.stringify(exported)),/未対応.*gradeScale/);
});

test('v3 profile rejects invalid and duplicate normalized field conversions',()=>{
  const e=adminEnvironment(),p=JSON.parse(e.c.exportSettingsProfile(false).json);
  p.config.fields[0].gradeScale=[{from:'01',to:'A'},{from:'1.0',to:'B'}];
  assert.throws(()=>e.c.inspectSettingsProfile(JSON.stringify(p)),/重複/);
  p.config.fields[0].gradeScale={from:'1',to:'A'};
  assert.throws(()=>e.c.inspectSettingsProfile(JSON.stringify(p)),/配列/);
});

test('portable profile preserves target connections on application and never starts automation',()=>{
  const e=adminEnvironment(),p=JSON.parse(e.c.exportSettingsProfile(false).json);p.config.messageTemplate='新しい文面';
  let applied, starts=0; const revision=e.c.getConfigRevision_(e.c.getConfig_());
  e.c.applyConfigDraft_=(config,expected)=>{assert.equal(expected,revision);applied=plain(config);return {ok:true};};
  e.c.configureAutomation=()=>{starts++;};
  e.c.applySettingsProfile(JSON.stringify(p),revision);
  assert.equal(applied.formSources.length,1);assert.equal(applied.reminderTo[0],'teacher@example.invalid');assert.equal(applied.messageTemplate,'新しい文面');assert.equal(starts,0);
});

test('profile application requires automation stopped, including legacy triggers',()=>{
  const e=adminEnvironment(),json=e.c.exportSettingsProfile(false).json;let applied=false;
  e.c.applyConfigDraft_=()=>{applied=true;};
  for(const state of [{enabled:true,legacyCount:0},{enabled:false,legacyCount:2}]){
    e.c.getAutomationStateUnlocked_=()=>state;
    assert.throws(()=>e.c.applySettingsProfile(json,e.c.getConfigRevision_(e.c.getConfig_())),/停止/);
  }
  assert.equal(applied,false);
});

test('failed config apply restores desired schedule without touching triggers',()=>{
  const e=adminEnvironment(),p=JSON.parse(e.c.exportSettingsProfile(false).json);p.schedule.deliveryHour=10;
  e.c.applyConfigDraft_=()=>{throw Error('save failed');};
  assert.throws(()=>e.c.applySettingsProfile(JSON.stringify(p),e.c.getConfigRevision_(e.c.getConfig_())),/save failed/);
  assert.equal(e.schedule.deliveryHour,8);
});

test('profile rejects unknown executable settings and prototype keys',()=>{
  const e=adminEnvironment(),p=JSON.parse(e.c.exportSettingsProfile(false).json);
  p.schedule.enabled=true;assert.throws(()=>e.c.inspectSettingsProfile(JSON.stringify(p)),/対応|項目/);
  delete p.schedule.enabled;p.config.fields[0].key='student_name';assert.throws(()=>e.c.inspectSettingsProfile(JSON.stringify(p)),/予約|固定/);
});

test('real automation integration prevents applying a profile owned by another manager',()=>{
  const e=adminEnvironment({actualAutomation:true});
  e.props.set('APP_AUTOMATION_STATE',JSON.stringify({revision:'old',owner:'other-owner',enabled:false,active:[],activeSchedule:null,schedule:{importHour:5,deliveryHour:8,reminderHour:16,reminderEnabled:false}}));
  const json=e.c.exportSettingsProfile(false).json;let applied=false;e.c.applyConfigDraft_=()=>{applied=true;};
  assert.throws(()=>e.c.applySettingsProfile(json,e.c.getConfigRevision_(e.c.getConfig_())),/管理者/);
  assert.equal(applied,false);
});
