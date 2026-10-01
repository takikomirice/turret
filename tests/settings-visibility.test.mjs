import test from 'node:test';
import assert from 'node:assert/strict';
import {adminEnvironment,AdminSheet,plain} from './helpers/admin-environment.mjs';
function legacy(e) {
 e.sheets.set('Config',new AdminSheet('Config',[[e.props.get('APP_CONFIG')]]));
 e.sheets.set('設定シート',new AdminSheet('設定シート',[['KEY','VALUE'],['REMINDER_TO','new@example.test'],['FORM_SS','']]));
 e.sheets.set('設定(編集不可)',new AdminSheet('設定(編集不可)',[['古い表示']]));
}
test('migration verifies internal settings before retiring all three tabs and preserves explicit blanks',()=>{
 const e=adminEnvironment();legacy(e);e.c.getAdminConsoleData();
 for(const key of ['Config','設定シート','設定(編集不可)'])assert.equal(e.sheets.has(key),false);
 assert.deepEqual(plain(e.c.getConfig_().reminderTo),['new@example.test']);assert.deepEqual(plain(e.c.getConfig_().formSources),[]);
 assert.equal(e.props.get('APP_CONFIG_BACKUP'),e.props.get('APP_CONFIG'));
 e.sheets.set('設定シート',new AdminSheet('設定シート',[['KEY','VALUE'],['REMINDER_TO','stale@example.test']]));
 assert.deepEqual(plain(e.c.getConfig_().reminderTo),['new@example.test']);
});
test('new settings and initialization never generate settings tabs',()=>{
 const e=adminEnvironment({config:null});e.c.initializeSheetsUnlocked_();e.c.saveConfig_(e.c.getConfig_());
 assert.deepEqual([...e.sheets.keys()].sort(),['クラス一覧','生徒一覧','採点テンプレ','対応表','送信シート','エラー','システム管理'].sort());
});
test('migration failure retains every legacy tab and retries safely',()=>{
 const e=adminEnvironment();legacy(e);const save=e.c.saveConfig_;
 e.c.saveConfig_=()=>{throw Error('save denied');};assert.equal(e.c.getAdminConsoleData().recoveryRequired,true);
 for(const key of ['Config','設定シート','設定(編集不可)'])assert.equal(e.sheets.has(key),true);
 e.c.saveConfig_=save;e.c.getAdminConsoleData();assert.equal(e.sheets.has('Config'),false);
});
test('partial tab deletion resumes without reapplying legacy overrides',()=>{
 const e=adminEnvironment();legacy(e);const remove=e.ss.deleteSheet;let fail=true;
 e.ss.deleteSheet=s=>{if(s.name==='設定シート'&&fail){fail=false;throw Error('delete denied');}return remove(s);};
 assert.equal(e.c.getAdminConsoleData().recoveryRequired,true);
 e.sheets.get('設定シート').rows[1][1]='stale@example.test';
 assert.equal(e.c.getAdminConsoleData().recoveryRequired,false);
 assert.deepEqual(plain(e.c.getConfig_().reminderTo),['new@example.test']);assert.equal(e.sheets.has('設定シート'),false);
});
test('corrupt primary and legacy fallback never delete configuration or silently default',()=>{
 const e=adminEnvironment();legacy(e);e.props.set('APP_CONFIG','broken');e.sheets.get('Config').rows[0][0]='broken';
 assert.equal(e.c.getAdminConsoleData().recoveryRequired,true);assert.equal(e.sheets.has('Config'),true);assert.throws(()=>e.c.getConfig_(),/設定/);
});
test('migration reads legacy fallback when primary is corrupt',()=>{
 const e=adminEnvironment();legacy(e);e.props.set('APP_CONFIG','broken');
 assert.equal(e.c.getAdminConsoleData().recoveryRequired,false);assert.equal(e.sheets.has('Config'),false);assert.equal(e.c.getConfig_().messageTemplate,'評価：＜score＞');
});
test('migrated reads need no spreadsheet I/O and backup recovery is visible',()=>{
 const e=adminEnvironment();e.c.getAdminConsoleData();e.c.SpreadsheetApp.getActiveSpreadsheet=()=>{throw Error('unexpected sheet access');};
 assert.equal(e.c.getConfig_().messageTemplate,'評価：＜score＞');e.props.set('APP_CONFIG','broken');
 assert.equal(e.c.loadConfig_().messageTemplate,'評価：＜score＞');assert.equal(e.props.get('APP_CONFIG_PARSE_ERROR'),'1');
});

test('migration keeps a visible non-settings sheet before deleting the final old tab',()=>{
 const e=adminEnvironment();legacy(e);const remove=e.ss.deleteSheet;
 e.ss.deleteSheet=s=>{assert.ok([...e.sheets.values()].some(other=>other!==s&&!other.isSheetHidden()));return remove(s);};
 e.c.getAdminConsoleData();assert.equal(e.sheets.get('クラス一覧').isSheetHidden(),false);
});
test('failed readback leaves properties and legacy settings unchanged',()=>{
 const e=adminEnvironment();legacy(e);const original=e.props.get('APP_CONFIG'),api=e.c.PropertiesService.getScriptProperties();let fail=true;
 e.c.PropertiesService.getScriptProperties=()=>({...api,setProperty(key,value){if(key==='APP_CONFIG_BACKUP'&&fail){fail=false;return;}api.setProperty(key,value);}});
 assert.equal(e.c.getAdminConsoleData().recoveryRequired,true);assert.equal(e.props.get('APP_CONFIG'),original);
 for(const name of ['Config','設定シート','設定(編集不可)'])assert.equal(e.sheets.has(name),true);
});
test('stale console revision cannot overwrite settings saved in another console',()=>{
 const e=adminEnvironment();e.c.getAdminConsoleData();const revision=e.c.getConfigRevision_(e.c.getConfig_());
 e.c.saveSetupSection('template',{messageTemplate:'first'},revision);
 assert.throws(()=>e.c.saveSetupSection('template',{messageTemplate:'stale'},revision),/変更/);
 assert.equal(e.c.getConfig_().messageTemplate,'first');
});


test('prepare sheets creates template, mapping and error headers and preserves their populated rows on repeat',()=>{
 const e=adminEnvironment();e.c.initializeSheetsUnlocked_();
 assert.deepEqual(e.sheets.get('採点テンプレ').rows[0].slice(0,9),['テンプレート名','枠1','枠2','枠3','枠4','枠5','出力1','出力2','出力3']);
 assert.deepEqual(e.sheets.get('対応表').rows[0],['元SS_ID','元スプシ名','元シート名','クラス名','courseId','メモ']);
 assert.equal(e.sheets.get('エラー').rows[0][5],'エラー種別');
 const before={};for(const name of ['クラス一覧','生徒一覧','採点テンプレ','対応表','エラー']){e.sheets.get(name).rows.push(['既存データ']);before[name]=plain(e.sheets.get(name).rows);}
 e.c.runAdminAction('initialize',null,e.c.getConfigRevision_(e.c.getConfig_()));
 for(const name of Object.keys(before))assert.deepEqual(plain(e.sheets.get(name).rows),before[name]);assert.equal(e.isLocked(),false);
});
