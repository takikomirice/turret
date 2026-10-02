import { test } from 'node:test';
import assert from 'node:assert/strict';
import { adminEnvironment, AdminSheet, plain } from './helpers/admin-environment.mjs';

test('empty recipients remain pending until opting out is explicitly saved',()=>{
 const e=adminEnvironment(),{c}=e,config=c.getConfig_();config.reminderTo=[];e.props.set('APP_CONFIG',JSON.stringify(config));
 assert.equal(c.getAdminConsoleData().progress.find(p=>p.id==='sources').state,'pending');
 let result=c.saveSetupSection('sources',{reminderTo:[],reminderOptOutConfirmed:false},c.getConfigRevision_(c.getConfig_()));
 assert.equal(c.getAdminConsoleData().progress.find(p=>p.id==='sources').state,'pending');
 result=c.saveSetupSection('sources',{reminderTo:[],reminderOptOutConfirmed:true},result.revision);
 assert.equal(result.config.reminderOptOutConfirmed,true);
 assert.equal(c.getAdminConsoleData().progress.find(p=>p.id==='sources').state,'complete');
 assert.equal(c.getConfig_().reminderOptOutConfirmed,true);
 assert.deepEqual(plain(c.getConfig_().formSources),plain(config.formSources));
 assert.deepEqual(plain(c.getConfig_().formSheetNamePrefix),plain(config.formSheetNamePrefix));
});

test('saved recipients complete setup; removing them requires a new opt-out confirmation',()=>{
 const e=adminEnvironment(),{c}=e;
 assert.equal(c.getAdminConsoleData().progress.find(p=>p.id==='sources').state,'complete');
 const result=c.saveSetupSection('sources',{reminderTo:[],reminderOptOutConfirmed:false},c.getConfigRevision_(c.getConfig_()));
 assert.equal(result.config.reminderOptOutConfirmed,false);
 assert.equal(c.getAdminConsoleData().progress.find(p=>p.id==='sources').state,'pending');
});

test('opt-out does not complete setup before form destinations exist or permit malformed values',()=>{
 const e=adminEnvironment({config:null}),{c}=e;
 c.saveSetupSection('sources',{reminderTo:[],reminderOptOutConfirmed:true},c.getConfigRevision_(c.getConfig_()));
 assert.equal(c.getAdminConsoleData().progress.find(p=>p.id==='sources').state,'pending');
 const revision=c.getConfigRevision_(c.getConfig_());
 assert.throws(()=>c.saveSetupSection('sources',{reminderTo:[],reminderOptOutConfirmed:'true'},revision),/確認|真偽値/);
 assert.throws(()=>c.saveSetupSection('sources',{reminderTo:['invalid'],reminderOptOutConfirmed:true},revision),/メール/);
 assert.equal(c.getConfigRevision_(c.getConfig_()),revision);
});

test('legacy JSON recipients do not carry an empty-recipient confirmation across operations',()=>{
 const e=adminEnvironment(),{c}=e;
 c.saveSetupSection('sources',{reminderTo:[],reminderOptOutConfirmed:true},c.getConfigRevision_(c.getConfig_()));
 const profile=c.exportSettingsProfile({reminderTo:true});
 assert.equal(Object.hasOwn(JSON.parse(profile.json).config,'reminderOptOutConfirmed'),false);
 c.applySettingsProfile(profile.json,c.getConfigRevision_(c.getConfig_()));
 assert.equal(c.getConfig_().reminderOptOutConfirmed,false);
 assert.equal(c.getAdminConsoleData().progress.find(p=>p.id==='sources').state,'pending');
});

test('adding recipients clears opt-out and stale saves preserve the current choice',()=>{
 const e=adminEnvironment(),{c}=e;
 const first=c.saveSetupSection('sources',{reminderTo:[],reminderOptOutConfirmed:true},c.getConfigRevision_(c.getConfig_()));
 const saved=c.saveSetupSection('sources',{reminderTo:['teacher@example.invalid'],reminderOptOutConfirmed:true},first.revision);
 assert.equal(saved.config.reminderOptOutConfirmed,false);
 assert.throws(()=>c.saveSetupSection('sources',{reminderTo:[],reminderOptOutConfirmed:true},first.revision),/変更/);
 assert.deepEqual(plain(c.getConfig_().reminderTo),['teacher@example.invalid']);
});

test('automation setup is not ready until the empty-recipient choice is saved',()=>{
 const e=adminEnvironment(),{c}=e;
 c.initializeSheetsUnlocked_();
 e.sheets.set('クラス一覧',new AdminSheet('クラス一覧',[['クラス名','コースID','同期対象(1)'],['授業','course-1',1]]));
 e.sheets.set('生徒一覧',new AdminSheet('生徒一覧',[['No','メールアドレス','名前','クラス名','コースID','studentId'],[1,'student@example.invalid','生徒','授業','course-1','student-1']]));
 e.props.set('TURRET_ROSTER_SELECTION',JSON.stringify(['course-1']));
 c.saveSetupSection('sources',{reminderTo:[],reminderOptOutConfirmed:false},c.getConfigRevision_(c.getConfig_()));
 assert.equal(c.getAdminConsoleData().ready,false);
 c.saveSetupSection('sources',{reminderTo:[],reminderOptOutConfirmed:true},c.getConfigRevision_(c.getConfig_()));
 assert.equal(c.getAdminConsoleData().ready,true);
});
