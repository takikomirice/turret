import { test } from 'node:test';
import assert from 'node:assert/strict';
import { adminEnvironment, AdminSheet, plain, validConfig } from './helpers/admin-environment.mjs';

test('initial setup creates a send sheet before configuration and safely refreshes empty headers',()=>{
  const e=adminEnvironment({config:null});
  let data=e.c.getAdminConsoleData();
  const result=e.c.runAdminAction('initialize',false,data.revision);
  assert.match(result.message,/送信シート/);
  const send=e.sheets.get('送信シート');assert.ok(send);
  assert.deepEqual(send.rows[0],plain(e.c.getConfiguredSendHeaders_(e.c.getConfig_())));
  assert.equal(e.sheets.has('評価データ'),false);
  e.props.set('APP_CONFIG',JSON.stringify(validConfig()));
  e.c.initializeSheetsUnlocked_();
  assert.deepEqual(send.rows[0],plain(e.c.getConfiguredSendHeaders_(e.c.getConfig_())));
  const row=send.rows[0].map((h,i)=>h==='送信状態'?'済':'保存値'+i);send.rows.push(row);
  e.c.initializeSheetsUnlocked_();assert.deepEqual(send.rows[1],row);
  const config=e.c.getConfig_();config.replyBodyHeader='別の本文';e.props.set('APP_CONFIG',JSON.stringify(config));
  const before=plain(send.rows);
  assert.throws(()=>e.c.initializeSheetsUnlocked_(),/送信シート.*構成/);
  assert.deepEqual(send.rows,before);assert.equal(e.sheets.has('評価データ'),false);
});

test('setup completion requires the send sheet to exist',()=>{
  const e=adminEnvironment();e.c.initializeSheetsUnlocked_();
  assert.equal(e.c.getAdminConsoleData().progress.find(p=>p.id==='prepare').state,'complete');
  e.sheets.delete('送信シート');
  assert.equal(e.c.getAdminConsoleData().progress.find(p=>p.id==='prepare').state,'pending');
});

test('new empty setup does not claim any completed business setup',()=>{
  const e=adminEnvironment({config:null}); const data=e.c.getAdminConsoleData();
  assert.equal(data.ready,false);
  assert.equal(data.progress.find(p=>p.id==='sources').state,'pending');
  assert.equal(data.progress.find(p=>p.id==='template').state,'pending');
  assert.equal(data.progress.find(p=>p.id==='mapping').state,'pending');
  assert.equal(data.counts.students,0);
  assert.equal(e.sheets.size,0);
});

test('only allowlisted sheet navigation is accepted',()=>{
  const e=adminEnvironment(); e.sheets.set('生徒一覧',new AdminSheet('生徒一覧',[['名前']]));
  e.c.openAdminSheet('students'); assert.equal(e.activeSheet(),'生徒一覧');
  assert.throws(()=>e.c.openAdminSheet('Config'),/対象|操作/);
  assert.throws(()=>e.c.openAdminSheet('__proto__'),/対象|操作/);
});

test('setup distinguishes connected forms, interrupted preparation, and material publication',()=>{
 const {c}=adminEnvironment();c.ScriptApp.getScriptId=()=> 'script-one';
 let p=c.getFormSetupProgress_(false,[]);
 assert.equal(p.forms.state,'pending');assert.equal(p.publish.state,'pending');
 p=c.getFormSetupProgress_(true,[]);
 assert.equal(p.forms.state,'complete');assert.equal(p.publish.state,'complete');assert.match(p.publish.detail,/既存/);
 const r={kind:'form',ownerSpreadsheetId:'test-admin-sheet',ownerScriptId:'script-one',stage:'registered'};
 p=c.getFormSetupProgress_(true,[r]);assert.equal(p.forms.state,'complete');assert.equal(p.publish.state,'pending');
 p=c.getFormSetupProgress_(true,[{...r,stage:'published'},{...r,stage:'creating'}]);
 assert.equal(p.forms.state,'attention');assert.notEqual(p.publish.state,'complete');
 p=c.getFormSetupProgress_(true,[{...r,stage:'published'}]);assert.equal(p.publish.state,'complete');
 p=c.getFormSetupProgress_(true,[{...r,stage:'publish_review'}]);assert.equal(p.publish.state,'attention');
 p=c.getFormSetupProgress_(true,[{...r,ownerScriptId:'copied-script'}]);assert.match(p.publish.detail,/既存/);
});

test('console provides exact tab URLs including gid zero without activating any sheet',()=>{
  const e=adminEnvironment();const classes=new AdminSheet('クラス一覧',[['クラス名']]);classes.getSheetId=()=>0;
  const students=new AdminSheet('生徒一覧',[['名前']]);students.getSheetId=()=>42;
  e.sheets.set('クラス一覧',classes);e.sheets.set('生徒一覧',students);
  const data=e.c.getAdminConsoleData();
  assert.equal(data.sheets.find(s=>s.key==='classes').url,'https://docs.google.com/spreadsheets/d/test-admin-sheet/edit#gid=0');
  assert.equal(data.sheets.find(s=>s.key==='students').url,'https://docs.google.com/spreadsheets/d/test-admin-sheet/edit#gid=42');
  assert.equal(data.sheets.find(s=>s.key==='mapping').url,'');assert.equal(e.activeSheet(),null);
});

test('send requires current explicit confirmation; unconfirmed and stale confirmations never post',()=>{
  const e=adminEnvironment(); let calls=0; e.c.sendMessagesUnlocked_=()=>{calls++;return 'sent';};
  const data=e.c.getAdminConsoleData();
  assert.throws(()=>e.c.runAdminAction('send',false,data.revision),/確認/);
  assert.throws(()=>e.c.runAdminAction('send',{confirmed:true,revision:'stale'},data.revision),/変更|確認/);
  assert.equal(calls,0);
  e.c.runAdminAction('send',{confirmed:true,revision:data.actionRevisions.send},data.revision);
  assert.equal(calls,1); assert.equal(e.isLocked(),false);
});

test('clear rejects a snapshot after a new row appears and preserves all data',()=>{
  const e=adminEnvironment(),sheet=new AdminSheet('評価データ',[['元SS_ID'],['first']]); e.sheets.set('評価データ',sheet);
  const data=e.c.getAdminConsoleData(); sheet.rows.push(['new']);
  assert.throws(()=>e.c.runAdminAction('clearEval',{confirmed:true,revision:data.actionRevisions.clearEval},data.revision),/変更|確認/);
  assert.equal(sheet.rows.length,3);
});

test('current confirmed clear keeps the header and does not use native alerts',()=>{
  const e=adminEnvironment(),sheet=new AdminSheet('評価データ',[['元SS_ID'],['first']]); e.sheets.set('評価データ',sheet);
  const data=e.c.getAdminConsoleData();
  e.c.runAdminAction('clearEval',{confirmed:true,revision:data.actionRevisions.clearEval},data.revision);
  assert.equal(sheet.rows[0][0],'元SS_ID'); assert.equal(sheet.getLastRow(),1);
});

test('unknown action and changed settings cannot call application workers',()=>{
  const e=adminEnvironment();let calls=0;e.c.classroomdataUnlocked_=()=>{calls++;};
  assert.throws(()=>e.c.runAdminAction('constructor',true,'x'),/操作/);
  assert.throws(()=>e.c.runAdminAction('classes',false,'stale'),/変更|読み/);
  assert.equal(calls,0);
});

test('preview explains field source and uses actual renderer with fictional values only',()=>{
  const e=adminEnvironment(); const preview=e.c.previewReplyTemplate(e.c.getConfig_());
  assert.match(preview.text,/評価：A/);
  const token=preview.tokens.find(t=>t.key==='score');
  assert.match(token.source,/点数/);assert.equal(token.example,'A');
  assert.ok(preview.tokens.find(t=>t.key==='student_name').label);
  assert.equal(e.sheets.size,0);
});

test('a broken save still opens the console in recovery mode so automation can be stopped',()=>{
  const e=adminEnvironment({actualAutomation:true});e.props.set('APP_CONFIG_SAVE_ERROR','rollback failed');
  const data=e.c.getAdminConsoleData();
  assert.equal(data.recoveryRequired,true);assert.equal(data.ready,false);assert.ok(data.warnings.length);
  const trigger={getEventType:()=> 'CLOCK',getHandlerFunction:()=> 'sendMessages',getUniqueId:()=> 'old-send'};e.triggers.push(trigger);
  const state=e.c.getAutomationState();e.c.stopAutomation(state.revision,true);
  assert.equal(e.triggers.length,0);
});

test('server readiness rejects starting automation with valid settings but no roster',()=>{
  const e=adminEnvironment({actualAutomation:true});
  const state=e.c.getAutomationState();
  assert.throws(()=>e.c.configureAutomation({importHour:5,deliveryHour:8,reminderHour:16,reminderEnabled:false},state.revision,false),/設定|準備|生徒/);
  assert.equal(e.triggers.length,0);
});

test('starting automation rejects a changed sending configuration before creating triggers',()=>{
  const e=adminEnvironment({actualAutomation:true});
  const revision=e.c.getConfigRevision_(e.c.getConfig_()),state=e.c.getAutomationState();
  const config=e.c.getConfig_();config.messageTemplate='別の設定';e.props.set('APP_CONFIG',JSON.stringify(config));
  e.c.assertAdminReadyForAutomation_=()=>{};
  assert.throws(()=>e.c.configureAutomation(state.schedule,state.revision,false,revision),/送信設定が変更/);
  assert.equal(e.triggers.length,0);assert.equal(e.userProps.size,0);
});

test('readiness derives actual roster selection and becomes incomplete when classes change',()=>{
  const e=adminEnvironment();
  e.sheets.set('設定シート',new AdminSheet('設定シート',[['KEY','VALUE','NOTE']]));
  e.sheets.set('クラス一覧',new AdminSheet('クラス一覧',[['クラス名','コースID','同期対象(1)'],['授業','course-1',1]]));
  e.sheets.set('生徒一覧',new AdminSheet('生徒一覧',[['No','メールアドレス','名前','クラス名','コースID','studentId'],[1,'student@example.invalid','生徒','授業','course-1','student-1']]));
  e.props.set('TURRET_ROSTER_SELECTION',JSON.stringify(['course-1']));
  e.c.ensureSendSheet_(e.c.getConfig_());
  e.c.managedSheet_(true);
  assert.equal(e.c.getAdminConsoleData().ready,true);
  e.sheets.get('クラス一覧').rows.push(['新しい授業','course-2',1]);
  assert.equal(e.c.getAdminConsoleData().ready,false);
});
