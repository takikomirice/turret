import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import fs from 'node:fs';
import {adminEnvironment,AdminSheet} from './helpers/admin-environment.mjs';

test('web opens only the stored operation; bound copy stays on its own workbook',()=>{
 const {c,ss,props}=adminEnvironment();
 props.set('TURRET_WEB_TARGET', 'original');
 assert.equal(c.getAppSpreadsheet_(),ss);
 c.SpreadsheetApp.getActiveSpreadsheet=()=>null;
 c.SpreadsheetApp.openById=id=>{assert.equal(id,'original');return ss;};
 assert.equal(c.getAppSpreadsheet_(),ss);
 props.delete('TURRET_WEB_TARGET');
 assert.throws(()=>c.getAppSpreadsheet_(),/接続/);
});
test('web bootstrap validates viewer before returning application data',()=>{
 const {c}=adminEnvironment();
 c.Session.getActiveUser=()=>({getEmail:()=> 'other@example.com'});
 c.Session.getEffectiveUser=()=>({getEmail:()=> 'owner@example.com'});
 assert.throws(()=>c.doGet(),/本人/);
});
test('new operation never replaces an existing connection',()=>{
 const {c}=adminEnvironment();
 assert.throws(()=>c.createWebOperation('次期'),/接続|既存/);
});
test('standalone bootstrap creates one JST workbook and never starts delivery',()=>{
 const {c,ss,props,sheets,triggers}=adminEnvironment({config:null});let created=0,timezone='';
 c.Session.getActiveUser=c.Session.getEffectiveUser=()=>({getEmail:()=> 'owner@example.com'});
 c.SpreadsheetApp.getActiveSpreadsheet=()=>null;
 c.SpreadsheetApp.create=name=>{assert.equal(name,'新しい運用');created++;return ss;};
 c.SpreadsheetApp.openById=id=>{assert.equal(id,ss.getId());return ss;};
 ss.setSpreadsheetTimeZone=value=>{timezone=value;};
 const result=c.createWebOperation('新しい運用');
 assert.equal(result.connected,true);assert.equal(result.web,true);assert.equal(timezone,'Asia/Tokyo');
 assert.equal(props.get('TURRET_WEB_TARGET'),ss.getId());assert.equal(sheets.has('クラス一覧'),true);assert.equal(triggers.length,0);
 assert.throws(()=>c.createWebOperation('別の運用'),/接続済み/);assert.equal(created,1);
});
test('next-period copy keeps settings and roster but never changes the source or copies queues',()=>{
 const {c,ss,sheets}=adminEnvironment();let copies=0;const copied=new Map();
 c.ScriptApp.getScriptId=()=> 'script-one';
 sheets.set('クラス一覧',new AdminSheet('クラス一覧',[['クラス名','コースID','同期対象(1)'],['テスト','100','1']]));
 sheets.set('送信シート',new AdminSheet('送信シート',[['送信状態'],['未']]));
 sheets.set('回答 1',new AdminSheet('回答 1',[['回答'],['元回答']]));
 c.DriveApp={getFolderById:()=>({getName:()=> 'folder'})};
 c.Drive={Files:{copy:()=>{copies++;for(const [name,s]of sheets)copied.set(name,new AdminSheet(name,s.rows));return {id:'new-copy'};}}};
 c.SpreadsheetApp.openById=id=>{assert.equal(id,'new-copy');return {getSheets:()=>[...copied.values()],getSheetByName:n=>copied.get(n),deleteSheet:s=>copied.delete(s.name),insertSheet:n=>{const s=new AdminSheet(n);copied.set(n,s);return s;}};};
 const p=c.previewNextOperation('次の期間','folder_12345678901234567890');
 const result=c.createNextOperation(p.name,p.folderId,p.fingerprint);
 assert.equal(result.stage,'ready');assert.equal(copies,1);
 assert.equal(copied.has('送信シート'),false);assert.equal(copied.has('回答 1'),false);assert.equal(copied.has('フォーム管理'),false);
 assert.equal(copied.has('クラス一覧'),true);assert.equal(copied.has('Config'),false);assert.equal(copied.has('設定シート'),false);
 const profile=JSON.parse(result.settingsProfile.json);assert.deepEqual(profile.config.formSources,[]);assert.equal(profile.includesConnections,true);assert.equal(profile.config.messageTemplate,'評価：＜score＞');assert.doesNotThrow(()=>c.parseSettingsProfile_(result.settingsProfile.json));
 assert.equal(sheets.get('送信シート').rows[1][0],'未');assert.equal(sheets.get('回答 1').rows[1][0],'元回答');
 c.createNextOperation(p.name,p.folderId,p.fingerprint);assert.equal(copies,1);
 const second=c.previewNextOperation('さらに次の期間',p.folderId),save=c.saveManagedRecord_;
 c.saveManagedRecord_=r=>{if(r.kind==='copy'&&r.stage==='ready')throw Error('final write lost');return save(r);};
 assert.throws(()=>c.createNextOperation(second.name,second.folderId,second.fingerprint),/final write lost/);
 c.saveManagedRecord_=save;
 copied.set('新しい回答',new AdminSheet('新しい回答',[['回答'],['新データ']]));
 assert.throws(()=>c.createNextOperation(second.name,second.folderId,second.fingerprint),/確認/);
 assert.equal(copied.get('新しい回答').rows[1][0],'新データ');
});

for(const url of ['https://script.google.com/macros/s/test/exec',null])test('web link dialog exposes selectable URL or deployment guidance: '+url,()=>{
 const {c,props}=adminEnvironment();let content='',shown=0;
 c.ScriptApp.getService=()=>({getUrl:()=>url});
 c.ScriptApp.getScriptId=()=> 'current-script';c.rememberWebConsoleUrl_();
 c.HtmlService={createHtmlOutput:html=>{content=html;return {setWidth(){return this;},setHeight(){return this;}};}};
 c.SpreadsheetApp.getUi=()=>({showModelessDialog(){shown++;}});
 c.connectWebConsole();assert.equal(shown,1);assert.equal(props.get('TURRET_WEB_TARGET'),'test-admin-sheet');
 if(url){assert.ok(content.includes(url));assert.match(content,/readonly/);assert.match(content,/URLをコピー/);assert.match(content,/target="_blank"/);assert.match(content,/select\(\)/);}
 else {assert.match(content,/デプロイ/);assert.doesNotMatch(content,/href="null"/);}
});
test('spreadsheet menu opens both independent web screens',()=>{
 const {c}=adminEnvironment();const entries=[];const menu={addItem(label,handler){entries.push([label,handler]);return this;},addToUi(){}};
 c.SpreadsheetApp.getUi=()=>({createMenu:()=>menu});c.onOpen();
 assert.deepEqual(entries,[['管理画面を開く','connectWebConsole'],['採点画面を開く','connectScoringConsole'],['採点テンプレ作成画面へ','openScoringRuleEditor']]);
});

for(const [page,file] of [[undefined,'Setting'],['settings','Setting'],['scoring','Scoring'],['templates','Scoring'],['../../Code','Setting']])test('web routing validates page and keeps operator authorization: '+page,()=>{
 const {c}=adminEnvironment();c.Session.getActiveUser=c.Session.getEffectiveUser=()=>({getEmail:()=> 'owner@example.com'});
 c.ScriptApp.getScriptId=()=> 'current';c.ScriptApp.getService=()=>({getUrl:()=>null});
 const html={content:fs.readFileSync(file+'.html','utf8'),getContent(){return this.content;},setContent(value){this.content=value;return this;},setTitle(){return this;},addMetaTag(){return this;}};
 c.HtmlService={createHtmlOutputFromFile:name=>{if(name==='Setting'&&file==='Scoring')return {getContent:()=>fs.readFileSync('Setting.html','utf8')};assert.equal(name,file);return html;}};
 assert.equal(c.doGet({parameter:{page}}),html);
 c.Session.getActiveUser=()=>({getEmail:()=> 'other@example.com'});assert.throws(()=>c.doGet({parameter:{page}}),/本人/);
});

test('scoring and standalone template editor share HTML and palette without a sidebar',()=>{
 const {c}=adminEnvironment();c.ScriptApp.getScriptId=()=> 'current';
 c.Session.getActiveUser=c.Session.getEffectiveUser=()=>({getEmail:()=> 'owner@example.com'});
 c.HtmlService={createHtmlOutputFromFile:name=>({content:fs.readFileSync(name+'.html','utf8'),getContent(){return this.content;},setContent(value){this.content=value;return this;},setTitle(){return this;}})};
 const theme=fs.readFileSync('Setting.html','utf8').match(/<style id="app-theme">[\s\S]*?<\/style>/)[0];
 function check(html){
  const content=html.getContent(),builder=content.match(/<template id="ruleBuilderTemplate">([\s\S]*?)<\/template>/)[1];
  assert.equal(content.split(theme).length-1,2);
  assert.ok(builder.includes(theme));
  assert.match(content,/<head>\s*<style id="app-theme">/);
 }
 check(c.createAppHtmlOutput_('Scoring'));
 const dedicated=c.createAppHtmlOutput_('Templates').getContent();
 assert.equal(dedicated.split(theme).length-1,1);assert.match(dedicated,/採点テンプレ作成/);assert.match(dedicated,/id="templateSelect"/);assert.doesNotMatch(dedicated,/<iframe|id="sheetSelect"|rule-sidebar/);
 assert.doesNotMatch(fs.readFileSync('Administration.gs','utf8'),/showSidebar/);
 c.HtmlService.createHtmlOutputFromFile=()=>({getContent:()=>'<head></head>'});
 assert.throws(()=>c.createAppHtmlOutput_('Scoring'),/共通配色/);
});

test('scoring menu preserves the verified deployment and adds the scoring route',()=>{
 const {c}=adminEnvironment();let html='';c.ScriptApp.getScriptId=()=> 'current';
 c.ScriptApp.getService=()=>({getUrl:()=> 'https://script.google.com/macros/s/test/exec'});c.rememberWebConsoleUrl_();
 c.HtmlService={createHtmlOutput:value=>{html=value;return {setWidth(){return this;},setHeight(){return this;}};}};
 c.SpreadsheetApp.getUi=()=>({showModelessDialog(){}});c.connectScoringConsole();
 assert.match(html,/\/exec\?page=scoring/);assert.match(html,/採点画面/);
});

for(const allowed of [true,false])test('copy URL works or leaves selected text with a keyboard fallback: '+allowed,async()=>{
 const {c}=adminEnvironment();let html='',selected=false;
 c.ScriptApp.getService=()=>({getUrl:()=> 'https://script.google.com/macros/s/test/exec'});
 c.ScriptApp.getScriptId=()=> 'current-script';c.rememberWebConsoleUrl_();
 c.HtmlService={createHtmlOutput:value=>{html=value;return {setWidth(){return this;},setHeight(){return this;}};}};
 c.SpreadsheetApp.getUi=()=>({showModelessDialog(){}});c.connectWebConsole();
 const input={value:'https://script.google.com/macros/s/test/exec',focus(){},select(){selected=true;}},status={};
 const context=vm.createContext({document:{getElementById:id=>id==='webUrl'?input:status,execCommand:()=>false},navigator:{clipboard:{async writeText(value){assert.equal(value,input.value);if(!allowed)throw Error('denied');}}}});
 vm.runInContext(html.match(/<script>([\s\S]*?)<\/script>/)[1],context);await context.copyUrl();
 assert.equal(selected,true);assert.match(status.textContent,allowed?/コピーしました/:/Ctrl\+C/);
});
test('deployment URL is HTML escaped in attributes',()=>{
 const {c}=adminEnvironment();assert.equal(c.escapeWebLink_('a"<>&'), 'a&quot;&lt;&gt;&amp;');
});

test('menu uses the URL captured in web context, not a stale bound getUrl result',()=>{
 const {c}=adminEnvironment();let html='';c.ScriptApp.getScriptId=()=> 'script-one';
 c.ScriptApp.getService=()=>({getUrl:()=> 'https://script.google.com/macros/s/current/exec'});c.rememberWebConsoleUrl_();
 c.ScriptApp.getService=()=>({getUrl:()=> 'https://script.google.com/macros/s/deleted/exec'});
 c.HtmlService={createHtmlOutput:value=>{html=value;return {setWidth(){return this;},setHeight(){return this;}};}};
 c.SpreadsheetApp.getUi=()=>({showModelessDialog(){}});c.connectWebConsole();
 assert.match(html,/\/current\/exec/);assert.doesNotMatch(html,/\/deleted\/exec/);
});

test('development URLs and copied project records never become public menu links',()=>{
 const {c,props}=adminEnvironment();c.ScriptApp.getScriptId=()=> 'script-one';
 c.ScriptApp.getService=()=>({getUrl:()=> 'https://script.google.com/macros/s/current/exec'});c.rememberWebConsoleUrl_();
 c.ScriptApp.getService=()=>({getUrl:()=> 'https://script.google.com/macros/s/dev-only/dev'});c.rememberWebConsoleUrl_();
 assert.match(c.getVerifiedWebConsoleUrl_(),/\/current\/exec$/);
 c.ScriptApp.getScriptId=()=> 'copied-script';assert.equal(c.getVerifiedWebConsoleUrl_(),null);
 props.set('TURRET_WEB_ENTRY','broken');assert.equal(c.getVerifiedWebConsoleUrl_(),null);
});

test('an unverified bound URL is never offered as a working web link',()=>{
 const {c}=adminEnvironment();let html='';c.ScriptApp.getScriptId=()=> 'script-one';
 c.ScriptApp.getService=()=>({getUrl:()=> 'https://script.google.com/macros/s/deleted/exec'});
 c.HtmlService={createHtmlOutput:value=>{html=value;return {setWidth(){return this;},setHeight(){return this;}};}};
 c.SpreadsheetApp.getUi=()=>({showModelessDialog(){}});c.connectWebConsole();
 assert.doesNotMatch(html,/\/deleted\/exec/);assert.match(html,/一度開/);
});

for(const url of [
 'https://script.google.com/macros/s/current/exec',
 'https://script.google.com/a/macros/e.osakamanabi.jp/s/current/exec',
 'https://script.google.com/a/e.osakamanabi.jp/macros/s/current/exec'
])test('authorized published entry registers its URL for all menu screens and rejects other viewers: '+url,()=>{
 const {c,props}=adminEnvironment();c.Session.getActiveUser=c.Session.getEffectiveUser=()=>({getEmail:()=> 'owner@example.com'});
 c.ScriptApp.getScriptId=()=> 'script-one';c.ScriptApp.getService=()=>({getUrl:()=>url});
 const html={getContent:()=>'<script></script>',setContent(){return this;},setTitle(){return this;},addMetaTag(){return this;}};
 c.HtmlService={createHtmlOutputFromFile:name=>{assert.equal(name,'Setting');return html;}};assert.equal(c.doGet(),html);
 assert.equal(c.getVerifiedWebConsoleUrl_(),url);
 assert.deepEqual(JSON.parse(props.get('TURRET_WEB_ENTRY')),{scriptId:'script-one',url});
 let dialog='';c.HtmlService.createHtmlOutput=value=>{dialog=value;return {setWidth(){return this;},setHeight(){return this;}};};
 c.SpreadsheetApp.getUi=()=>({showModelessDialog(){}});
 for(const [entry,suffix] of [['connectWebConsole',''],['connectScoringConsole','?page=scoring'],['openScoringRuleEditor','?page=templates']]){
  c[entry]();assert.ok(dialog.includes('href="'+url+suffix+'"'),entry+' must open the registered deployment');
 }
 props.delete('TURRET_WEB_ENTRY');c.Session.getActiveUser=()=>({getEmail:()=> 'other@example.com'});
 assert.throws(()=>c.doGet(),/本人/);assert.equal(props.has('TURRET_WEB_ENTRY'),false);
});

test('development and untrusted URLs never replace a registered published entry',()=>{
 const {c,props}=adminEnvironment();const url='https://script.google.com/macros/s/current/exec';
 c.ScriptApp.getScriptId=()=> 'script-one';c.ScriptApp.getService=()=>({getUrl:()=>url});c.rememberWebConsoleUrl_();
 const stored=props.get('TURRET_WEB_ENTRY');
 for(const rejected of [
  null,undefined,'',
  'https://script.google.com/macros/s/current/dev',
  'https://script.google.com/a/macros/e.osakamanabi.jp/s/current/dev',
  'https://script.google.com/a/e.osakamanabi.jp/macros/s/current/dev',
  'http://script.google.com/a/e.osakamanabi.jp/macros/s/current/exec',
  'https://example.com/a/e.osakamanabi.jp/macros/s/current/exec',
  'https://script.google.com.example.com/a/e.osakamanabi.jp/macros/s/current/exec',
  'https://script.google.com/a/e.osakamanabi.jp/macros/s/current/exec?next=https://example.com',
  'https://script.google.com/a/e.osakamanabi.jp/macros/s/current/exec/extra'
 ]){
  c.ScriptApp.getService=()=>({getUrl:()=>rejected});c.rememberWebConsoleUrl_();
  assert.equal(c.isPublishedWebConsoleUrl_(rejected),false,String(rejected));
  assert.equal(props.get('TURRET_WEB_ENTRY'),stored,'invalid URL must preserve the registered entry');
  assert.equal(c.getVerifiedWebConsoleUrl_(),url);
 }
});

for(const entry of ['openSettingsDialog','openMaintenanceDialog','openManualActionsDialog'])test('legacy entry redirects to web connection: '+entry,()=>{
 const {c}=adminEnvironment();let calls=0;c.connectWebConsole=()=>{calls++;return 'web';};
 assert.equal(c[entry](),'web');assert.equal(calls,1);
});

for(const outcome of ['opened','blocked','error'])test('menu automatically opens the verified URL and retains fallback when unavailable: '+outcome,()=>{
 const {c}=adminEnvironment();let html='',closed=0,attempts=0;
 c.ScriptApp.getService=()=>({getUrl:()=> 'https://script.google.com/macros/s/test/exec'});
 c.ScriptApp.getScriptId=()=> 'current-script';c.rememberWebConsoleUrl_();
 c.HtmlService={createHtmlOutput:value=>{html=value;return {setWidth(){return this;},setHeight(){return this;}};}};
 c.SpreadsheetApp.getUi=()=>({showModelessDialog(){}});c.connectWebConsole();
 const input={value:'https://script.google.com/macros/s/test/exec'},status={},popup={opener:{},location:{}};
 const context=vm.createContext({document:{getElementById:id=>id==='webUrl'?input:status},window:{open(url,target){attempts++;assert.equal(url,'about:blank');assert.equal(target,'_blank');if(outcome==='error')throw Error('blocked');return outcome==='opened'?popup:null;}},google:{script:{host:{close(){closed++;}}}}});
 vm.runInContext(html.match(/<script>([\s\S]*?)<\/script>/)[1],context);
 assert.equal(attempts,1);
 if(outcome==='opened'){assert.equal(popup.location.href,input.value);assert.equal(popup.opener,null);assert.equal(closed,1);}
 else {assert.equal(closed,0);assert.match(status.textContent,/リンク/);assert.match(html,/target="_blank"/);assert.match(html,/URLをコピー/);}
});


test('template menu launches the verified web editor route with a usable popup fallback',()=>{
 const {c}=adminEnvironment();let html='',label='';c.ScriptApp.getScriptId=()=> 'current';
 c.ScriptApp.getService=()=>({getUrl:()=> 'https://script.google.com/macros/s/test/exec'});c.rememberWebConsoleUrl_();
 c.HtmlService={createHtmlOutput:value=>{html=value;return {setWidth(){return this;},setHeight(){return this;}};}};
 c.SpreadsheetApp.getUi=()=>({showModelessDialog(_,title){label=title;}});c.openScoringRuleEditor();
 assert.equal(label,'採点テンプレ作成画面');assert.match(html,/exec\?page=templates/);assert.doesNotMatch(html,/docs\.google\.com|↗/);
 const status={};vm.runInNewContext(html.match(/<script>([\s\S]*?)<\/script>/)[1],{window:{open:()=>null},document:{getElementById:()=>status}});
 assert.match(status.textContent,/採点テンプレ作成画面/);
});
