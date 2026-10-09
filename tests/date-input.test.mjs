import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import {readFileSync} from 'node:fs';

const source=readFileSync('Setting.html','utf8').match(/<script>([\s\S]*?)<\/script>/)[1].replace(/boot\(\);\s*$/,'');
function model(type='date',value='2026-10-09'){
 const elements=new Map(),events={};
 const c=vm.createContext({console,structuredClone,setTimeout:()=>0,clearTimeout(){}});vm.runInContext(source,c);
 const native={id:'sample',type,value,dataset:{dateNative:'true'},closest:()=>true,matches:()=>false};
 const display={id:'date-display-sample',value:'',dataset:{dateInput:'sample'},closest:()=>true,setCustomValidity(message){this.validationMessage=message;},reportValidity(){return !this.validationMessage;},focus(){}};
 elements.set(native.id,native);elements.set(display.id,display);
 c.document={getElementById:id=>elements.get(id)||{addEventListener(){}},querySelectorAll:selector=>selector==='[data-date-input]'?[display]:[],addEventListener:(event,handler)=>{(events[event]||=[]).push(handler);}};
 c.window={addEventListener(){}};
 return {c,native,display,elements,events};
}

test('posting inputs show formatted weekdays inside a text field and keep a native calendar with ISO values',()=>{
 const {c}=model();
 for(const key of ['batch:publish','batch:close','publish:r1','close:r1']){
  const html=c.scheduleInput(key,'日時','2026-10-09T07:52:00Z');
  assert.match(html,/type="text"[^>]*value="2026\/10\/09（金） 16:52"/);
  assert.match(html,/type="datetime-local"[^>]*value="2026-10-09T16:52"/);
  assert.ok(html.includes('data-date-input="schedule-'+key+'"'));
 }
});

test('date text parses to canonical values without storing the weekday and normalizes only on commit',()=>{
 const {c,native,display}=model('datetime-local','2026-10-09T16:52');
 display.value='2026/10/15 09:30';assert.equal(c.syncDateInput(display),native);
 assert.equal(native.value,'2026-10-15T09:30');assert.equal(display.value,'2026/10/15 09:30');
 c.syncDateInput(display,true);assert.equal(display.value,'2026/10/15（木） 09:30');
 native.value='2027-01-01T00:00';c.syncDateInput(native);assert.equal(display.value,'2027/01/01（金） 00:00');
 display.value='';c.syncDateInput(display,true);assert.equal(native.value,'');assert.equal(display.value,'');
});

test('invalid typed dates retain the last valid value and prevent saving instead of clearing the date',()=>{
 const {c,native,display}=model();
 for(const bad of ['2026/02/30','2026/13/01','0000/01/01','途中','2026/10/09 24:00']){
  display.value=bad;c.syncDateInput(display,true);
  assert.equal(native.value,'2026-10-09');assert.equal(display.value,bad);
  assert.throws(()=>c.assertDateInputsValid(),/日付|日時/);
 }
 display.value='2028/2/29';c.syncDateInput(display,true);assert.equal(native.value,'2028-02-29');assert.equal(display.value,'2028/02/29（火）');
 assert.doesNotThrow(()=>c.assertDateInputsValid());
});

test('date controls use calendar weekdays independent of browser timezone and never decorate invalid dates',()=>{
 const {c}=model();
 for(const [date,day] of [['2026-10-04','日'],['2026-10-05','月'],['2026-10-06','火'],['2026-10-07','水'],['2026-10-08','木'],['2026-10-09','金'],['2026-10-10','土'],['2027-01-01','金'],['2028-02-29','火']]){
  assert.equal(c.dateInputDisplay(date,'date'),date.replaceAll('-','/')+'（'+day+'）');
 }
 assert.equal(c.dateInputDisplay('','date'),'');assert.equal(c.dateInputDisplay('2026-02-30','date'),'2026-02-30');
});

test('redrawing preserves an unfinished edit until a calendar selection replaces it',()=>{
 const {c,native,display}=model();display.value='2026/';c.syncDateInput(display);
 let html=c.dateInputControl('sample','date',native.value,'日付');assert.match(html,/type="text"[^>]*value="2026\/"/);assert.match(html,/aria-invalid="true"/);
 native.value='2026-10-15';c.syncDateInput(native);html=c.dateInputControl('sample','date',native.value,'日付');
 assert.match(html,/value="2026\/10\/15（木）"/);assert.doesNotMatch(html,/aria-invalid="true"/);assert.doesNotThrow(()=>c.assertDateInputsValid());
});

test('input and change handlers synchronize the visible field before the existing schedule draft is read',async()=>{
 const e=model('datetime-local','2026-10-09T16:52'),{c,native,display,events}=e;
 native.dataset.scheduleKey='publish:r1';native.dataset.scheduleSaved='2026-10-09T16:52';
 const preview={textContent:''};e.elements.set('schedule-preview-publish:r1',preview);
 c.document.querySelectorAll=selector=>selector==='[data-date-input]'?[display]:selector==='[data-schedule-key]'?[native]:[];
 vm.runInContext("initHelp=()=>{};task=()=>{};readCurrent=()=>readScheduleDrafts();updateMaterialSettingsControls=()=>{};render=()=>{throw Error('入力中は再描画しない');};boot();",c);
 display.value='2026/10/15 09:30';for(const fn of events.input)await fn({target:display});
 assert.equal(vm.runInContext("state.scheduleDrafts['publish:r1']",c),'2026-10-15T09:30');
 assert.match(preview.textContent,/2026-10-15（木）/);
 for(const fn of events.change)await fn({target:display});assert.equal(display.value,'2026/10/15（木） 09:30');
});
