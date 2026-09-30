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
test('scheduled forms offer changing/cancelling and deadline controls with explicit local status',()=>{
 const c=model();vm.runInContext("state.forms={classes:[],defaults:{},records:[{id:'r1',kind:'form',stage:'scheduled',label:'クラス',materialId:'m1',scheduledTime:'2026-10-01T01:00:00Z',closesAt:'2026-10-02T03:00:00Z',closeState:'error',closeError:'接続失敗'}]};",c);
 const html=vm.runInContext("renderForms('publish')",c);
 assert.match(html,/予約済み/);assert.match(html,/2026-10-01T10:00/);assert.match(html,/data-form-command="reschedule"/);
 assert.match(html,/data-form-command="cancel-schedule"/);assert.match(html,/data-form-command="set-close"/);assert.match(html,/接続失敗/);
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
