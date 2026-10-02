import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import {readFileSync} from 'node:fs';
const script=readFileSync('Setting.html','utf8').match(/<script>([\s\S]*?)<\/script>/)[1];
function model(){const c=vm.createContext({console,structuredClone});vm.runInContext(script.replace(/boot\(\);\s*$/,''),c);return c;}
test('datetime inputs always represent Japan time, independently of browser timezone',()=>{
 const c=model();assert.equal(vm.runInContext("scheduleInputToIso('2026-10-01T10:30')",c),'2026-10-01T01:30:00.000Z');
 assert.equal(vm.runInContext("scheduleIsoToInput('2026-10-01T01:30:00Z')",c),'2026-10-01T10:30');
 for(const bad of ['', '2026-02-30T10:00','2026-10-01T24:30'])assert.throws(()=>c.scheduleInputToIso(bad),/日時/);
});
test('date displays include all Japanese weekdays at Japan day, month and year boundaries',()=>{
 const c=model();
 for(const [iso,expected] of [
  ['2026-10-03T15:00:00Z','2026-10-04（日） 00:00（日本時間）'],
  ['2026-10-04T15:00:00Z','2026-10-05（月） 00:00（日本時間）'],
  ['2026-10-05T15:00:00Z','2026-10-06（火） 00:00（日本時間）'],
  ['2026-10-06T15:00:00Z','2026-10-07（水） 00:00（日本時間）'],
  ['2026-10-07T15:00:00Z','2026-10-08（木） 00:00（日本時間）'],
  ['2026-10-08T15:00:00Z','2026-10-09（金） 00:00（日本時間）'],
  ['2026-10-09T15:00:00Z','2026-10-10（土） 00:00（日本時間）'],
  ['2026-12-31T14:59:00Z','2026-12-31（木） 23:59（日本時間）'],
  ['2026-12-31T15:00:00Z','2027-01-01（金） 00:00（日本時間）'],
  ['2028-02-28T15:00:00Z','2028-02-29（火） 00:00（日本時間）'],
  ['2028-02-29T15:00:00Z','2028-03-01（水） 00:00（日本時間）'],
  ['2026-09-30T08:00:00-07:00','2026-10-01（木） 00:00（日本時間）']
 ])assert.equal(c.scheduleDisplay(iso),expected);
});
test('native datetime inputs retain their value with accessible weekday previews',()=>{
 const c=model();
 for(const key of ['batch:publish','batch:close','publish:r1','close:r1']){
  const html=c.scheduleInput(key,'日時','2026-10-01T01:30:00Z');
  assert.match(html,/type="datetime-local"/);assert.match(html,/value="2026-10-01T10:30"/);
  assert.ok(html.includes('aria-describedby="schedule-preview-'+key+'"'));
  assert.ok(html.includes('id="schedule-preview-'+key+'"'));
  assert.match(html,/aria-live="polite"/);assert.match(html,/aria-atomic="true"/);
  assert.match(html,/2026-10-01（木） 10:30（日本時間）/);
 }
 vm.runInContext("state.scheduleDrafts={'publish:r1':'2027-01-01T00:00'}",c);
 assert.match(c.scheduleInput('publish:r1','日時','2026-10-01T01:30:00Z'),/2027-01-01（金） 00:00（日本時間）/);
 vm.runInContext("delete state.scheduleDrafts['publish:r1']",c);
 assert.match(c.scheduleInput('publish:r1','日時','2027-01-02T00:00:00Z'),/2027-01-02（土） 09:00（日本時間）/);
});
test('blank and invalid dates never show a misleading weekday or throw during rendering',()=>{
 const c=model();assert.equal(c.scheduleDisplay(''),'未設定');
 let display;assert.doesNotThrow(()=>{display=c.scheduleDisplay('invalid');});assert.match(display,/日時/);assert.doesNotMatch(display,/[（(][日月火水木金土][）)]/);
 for(const value of ['', '2026-02-30T10:00','2026-10-01T24:30','invalid']){
  c.value=value;vm.runInContext("state.scheduleDrafts={'publish:r1':value}",c);
  const html=c.scheduleInput('publish:r1','日時');
  assert.doesNotMatch(html,/[（(][日月火水木金土][）)]/);
  assert.match(html,value?/日時が不正/:/未設定/);
 }
 assert.doesNotThrow(()=>c.scheduleInput('publish:r1','日時','invalid'));
});
test('nonexistent saved calendar dates are rejected instead of rolling into a different weekday',()=>{
 const c=model();
 for(const value of ['2026-02-30T10:00:00Z','2026-04-31T10:00:00+09:00']){
  assert.equal(c.scheduleDisplay(value),'日時が不正です。');
  assert.equal(c.scheduleIsoToInput(value),'');
  assert.doesNotMatch(c.scheduleInput('publish:r1','日時',value),/[（(][日月火水木金土][）)]/);
 }
});
test('editing datetime input updates its weekday preview without replacing the focused field',async()=>{
 const c=model(),events={},preview={textContent:'2026-10-01（木） 10:30（日本時間）'};
 const input={id:'schedule-publish:r1',dataset:{scheduleKey:'publish:r1',scheduleSaved:'2026-10-01T10:30'},value:'2026-10-02T10:30',closest:()=>true};
 c.document={activeElement:input,getElementById:id=>id==='schedule-preview-publish:r1'?preview:{addEventListener(){}},querySelectorAll:selector=>selector==='[data-schedule-key]'?[input]:[],addEventListener:(name,fn)=>{(events[name]||=[]).push(fn);}};
 c.window={addEventListener(){}};
 vm.runInContext("initHelp=()=>{};task=()=>{};readCurrent=()=>readScheduleDrafts();updateMaterialSettingsControls=()=>{};render=()=>{throw Error('日時入力中に画面を再描画しない');};boot();",c);
 for(const listener of events.input)await listener({target:input});
 assert.equal(preview.textContent,'2026-10-02（金） 10:30（日本時間）');
 assert.equal(c.document.activeElement,input);assert.equal(input.value,'2026-10-02T10:30');
 assert.equal(vm.runInContext("state.scheduleDrafts['publish:r1']",c),'2026-10-02T10:30');
 input.value='2026-10-03T10:30';for(const listener of events.change)await listener({target:input});
 assert.equal(preview.textContent,'2026-10-03（土） 10:30（日本時間）');
 input.value='';for(const listener of events.input)await listener({target:input});assert.equal(preview.textContent,'未設定');
 input.value='2026-02-30T10:00';for(const listener of events.input)await listener({target:input});assert.match(preview.textContent,/日時が不正/);
});
test('accepted schedule results refresh and clear weekday previews beside the existing input',()=>{
 const c=model(),preview={textContent:'以前の日時'},input={dataset:{scheduleKey:'publish:r1'},value:'2026-10-01T10:00'};
 c.document={getElementById:id=>id==='schedule-publish:r1'?input:id==='schedule-preview-publish:r1'?preview:null};
 c.acceptScheduleResult('r1','reschedule',{scheduledTime:'2026-12-31T15:00:00Z'});
 assert.equal(input.value,'2027-01-01T00:00');assert.equal(preview.textContent,'2027-01-01（金） 00:00（日本時間）');
 c.acceptScheduleResult('r1','cancel-schedule',{});assert.equal(input.value,'');assert.equal(preview.textContent,'未設定');
});
test('publishing date summaries and delete confirmations include Japan weekdays',()=>{
 const c=model();
 vm.runInContext("state.forms={classes:[],defaults:{},records:[{id:'r1',kind:'form',stage:'scheduled',label:'クラス',scheduledTime:'2026-12-31T15:00:00Z',closesAt:'2027-01-02T00:00:00Z'}]}",c);
 const html=vm.runInContext("renderFormRecords('publish')",c);
 assert.match(html,/資料公開予定：2027-01-01（金） 00:00（日本時間）/);assert.match(html,/受付終了：2027-01-02（土） 09:00（日本時間）/);
 const summary=c.managedActionSummary({className:'クラス',title:'資料',scheduledTime:'2026-12-31T15:00:00Z',closesAt:'2027-01-02T00:00:00Z',postedAt:'2026-10-01T01:30:00Z'},'delete');
 assert.match(summary,/資料公開：2027-01-01（金） 00:00（日本時間）/);assert.match(summary,/受付終了：2027-01-02（土） 09:00（日本時間）/);assert.match(summary,/投稿日時：2026-10-01（木） 10:30（日本時間）/);
});
test('scheduled forms offer changing/cancelling and deadline controls with explicit local status',()=>{
 const c=model();vm.runInContext("state.forms={classes:[],defaults:{},records:[{id:'r1',kind:'form',stage:'scheduled',label:'クラス',materialId:'m1',scheduledTime:'2026-10-01T01:00:00Z',closesAt:'2026-10-02T03:00:00Z',closeState:'error',closeError:'接続失敗'}]};",c);
 const html=vm.runInContext("renderForms('publish')",c);
 assert.match(html,/予約済み/);assert.match(html,/2026-10-01T10:00/);assert.match(html,/data-form-command="reschedule"/);
 assert.match(html,/<option value="cancel-schedule">/);assert.match(html,/data-form-command="set-close"/);assert.match(html,/接続失敗/);
 assert.doesNotMatch(html,/data-form-command="publish" data-record="r1"/);
});
test('schedule drafts remain attached to their record across refreshes',()=>{
 const c=model();c.document={getElementById:()=>null,querySelectorAll:()=>[{dataset:{scheduleKey:'publish:r1'},value:'2026-10-03T10:00'}]};
 vm.runInContext('readScheduleDrafts()',c);
 vm.runInContext("state.forms={records:[{id:'r1',kind:'form',stage:'scheduled',scheduledTime:'2026-10-01T01:00:00Z'}],classes:[],defaults:{}}",c);
 assert.match(vm.runInContext("renderForms('publish')",c),/2026-10-03T10:00/);
});
test('batch scheduling retains successes and reports each failure without repeating successful records',async()=>{
 const c=model();c.document={getElementById:()=>null,querySelectorAll:()=>[]};
 vm.runInContext(`state.forms={records:[{id:'a',kind:'form',stage:'registered',label:'A'},{id:'b',kind:'form',stage:'registered',label:'B'}]};
 state.scheduleDrafts={'batch:publish':'2026-10-01T10:00'};let calls=[],messages=[];
 readCurrent=()=>{};requireSaved=()=>true;task=async fn=>fn();confirmAction=async()=>true;reloadForms=async()=>{};status=(message)=>messages.push(message);
 rpc=async(method,id,action,...args)=>{calls.push([method,id,action,...args]);if(method==='previewManagedFormAction')return {fingerprint:'fp',title:id,className:id};if(id==='b')throw Error('拒否');state.forms.records[0].stage='scheduled';return {};};`,c);
 await vm.runInContext("formCommand('schedule-all')",c);
 assert.equal(vm.runInContext("calls.filter(c=>c[0]==='runManagedFormAction').length",c),2);
 assert.match(vm.runInContext('messages.at(-1)',c),/1件/);assert.match(vm.runInContext('messages.at(-1)',c),/B.*拒否/);
 assert.equal(vm.runInContext("calls.find(c=>c[0]==='runManagedFormAction')[5].scheduledTime",c),'2026-10-01T01:00:00.000Z');
 vm.runInContext('calls=[]',c);await vm.runInContext("formCommand('schedule-all')",c);
 assert.equal(vm.runInContext("calls.some(c=>c[1]==='a')",c),false);
});

test('saving a deadline preserves an unrelated unsaved material date and rejects an explicitly blank deadline',async()=>{
 const c=model();c.document={getElementById:()=>null,querySelectorAll:()=>[]};
 vm.runInContext(`state.forms={records:[{id:'r',kind:'form',stage:'registered',closesAt:'2026-10-02T02:00:00Z'}]};
 state.scheduleDrafts={'publish:r':'2026-10-03T10:00','close:r':'2026-10-02T12:00'};
 readCurrent=()=>{};requireSaved=()=>true;task=async fn=>fn();confirmAction=async()=>true;reloadForms=async()=>{};status=()=>{};
 rpc=async(method)=>method==='previewManagedFormAction'?{fingerprint:'fp'}:{id:'r',closesAt:'2026-10-02T03:00:00Z'};`,c);
 await vm.runInContext("formCommand('set-close','r')",c);
 assert.equal(vm.runInContext("state.scheduleDrafts['publish:r']",c),'2026-10-03T10:00');
 vm.runInContext("state.scheduleDrafts['close:r']=''",c);
 await assert.rejects(vm.runInContext("formCommand('set-close','r')",c),/日時/);
});

test('retrying a deadline batch skips records whose requested deadline was already saved',async()=>{
 const c=model();c.document={getElementById:()=>null,querySelectorAll:()=>[]};
 vm.runInContext(`state.forms={records:[{id:'a',kind:'form',stage:'registered',label:'A'},{id:'b',kind:'form',stage:'registered',label:'B'}]};
 state.scheduleDrafts={'batch:close':'2026-10-01T11:00'};let calls=[];
 readCurrent=()=>{};requireSaved=()=>true;task=async fn=>fn();confirmAction=async()=>true;reloadForms=async()=>{};status=()=>{};
 rpc=async(method,id,action,...args)=>{calls.push([method,id]);if(method==='previewManagedFormAction')return {fingerprint:'fp',title:id};if(id==='b')throw Error('拒否');state.forms.records[0].closesAt='2026-10-01T02:00:00.000Z';return state.forms.records[0];};`,c);
 await vm.runInContext("formCommand('set-close-all')",c);vm.runInContext('calls=[]',c);
 await vm.runInContext("formCommand('set-close-all')",c);
 assert.equal(vm.runInContext("calls.some(c=>c[1]==='a')",c),false);
 assert.equal(vm.runInContext("calls.filter(c=>c[0]==='runManagedFormAction'&&c[1]==='b').length",c),1);
});
