import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
function client(){
 const storage=new Map(),calls=[],timers=new Map();let timer=0;
 const localStorage={getItem:k=>storage.get(k)||null,setItem:(k,v)=>storage.set(k,v),removeItem:k=>storage.delete(k),get length(){return storage.size;},key:i=>[...storage.keys()][i]};
 const window={__SCORING_TOOL_ENABLE_TEST_HOOKS__:true};
 const c=vm.createContext({window,document:{readyState:'loading',addEventListener(){}},console,localStorage,setTimeout:fn=>{timers.set(++timer,fn);return timer;},clearTimeout:id=>timers.delete(id),google:{script:{get run(){const run={withSuccessHandler(fn){this.success=fn;return this;},withFailureHandler(fn){this.failure=fn;return this;}};return new Proxy(run,{get:(o,k)=>k in o?o[k]:(payload)=>{calls.push({method:k,payload,success:o.success,failure:o.failure});}});}}}});
 let script=fs.readFileSync('Scoring.html','utf8').match(/<script>([\s\S]*?)<\/script>/)[1];
 script=script.replace('window.__scoringToolTestHooks = {',`setStatus=function(message){window.lastStatus=message;};loadInitData=function(){window.reloadCount=(window.reloadCount||0)+1;};
 window.__scoringToolTestHooks = {state,els,saveCurrentRow,flushPendingEdits,getDraftKey,getSpreadsheetId,savePendingEditToLocalStorage,loadPendingEditFromLocalStorage,markDirty,moveRow,`);
 vm.runInContext(script,c);const h=window.__scoringToolTestHooks;
 h.state.config={scoreCols:['D','','','',''],colChecks:[false,false,false,false,false],sheetName:''};
 h.state.operationId='operation';h.state.currentSheet='["book-a",1]';h.state.targets={[h.state.currentSheet]:{spreadsheetId:'book-a',sheetId:1}};
 h.state.rows=[{rowNumber:2,scores:[1,'','','',''],comment:'old'}];h.state.currentIndex=0;h.state.dirty=true;
 h.els.scoreInputs=Array.from({length:5},(_,i)=>({value:i?'':'3',disabled:i>0}));h.els.ruleOutInputs=[{value:'new'},{value:''},{value:''}];
 return {h,window,storage,calls,timers};
}

test('flush includes the last edited row and pending feedback calculation',()=>{
 const e=client();e.h.state.pendingRuleCalculation=()=>{e.h.els.ruleOutInputs[0].value='latest feedback';};
 e.h.flushPendingEdits(false);
 assert.equal(e.calls.length,1);assert.equal(e.calls[0].method,'scoringSaveRows');
 assert.equal(e.calls[0].payload.rows[0].scores[0],'3');assert.equal(e.calls[0].payload.rows[0].comment,'latest feedback');
});

test('a save response cannot clear newer edits or drafts for a different target',()=>{
 const e=client();e.h.flushPendingEdits(false);
 e.h.els.scoreInputs[0].value='4';e.h.saveCurrentRow();
 e.h.state.currentSheet='["book-b",1]';e.h.state.targets[e.h.state.currentSheet]={spreadsheetId:'book-b',sheetId:1};
 e.h.savePendingEditToLocalStorage(e.h.state.currentSheet,2,[5,'','','',''],'book B',[]);
 e.calls[0].success({ok:true,results:[{rowNumber:2,ok:true}]});
 assert.equal(e.storage.size,2);assert.equal(e.h.loadPendingEditFromLocalStorage('["book-a",1]',2).scores[0],'4');
 assert.equal(e.h.loadPendingEditFromLocalStorage('["book-b",1]',2).comment,'book B');assert.equal(e.window.reloadCount||0,0);
});

test('failed save retains drafts and successful rows are removed individually',()=>{
 const e=client();e.h.savePendingEditToLocalStorage(e.h.state.currentSheet,3,[2,'','','',''],'other',[]);
 e.h.flushPendingEdits(false);e.calls[0].success({ok:false,results:[{rowNumber:2,ok:true},{rowNumber:3,ok:false,error:'failed'}]});
 assert.equal(e.h.loadPendingEditFromLocalStorage(e.h.state.currentSheet,2),null);
 assert.ok(e.h.loadPendingEditFromLocalStorage(e.h.state.currentSheet,3));assert.match(e.window.lastStatus,/失敗/);
});

test('the first edit schedules saving even before a local draft exists',()=>{
 const e=client();e.h.markDirty();assert.ok(e.timers.size>0);
});

test('local persistence failure prevents sending an older draft as the latest edit',()=>{
 const e=client();e.h.saveCurrentRow();e.h.els.scoreInputs[0].value='4';e.h.markDirty();
 e.storage.set=()=>{throw Error('quota exceeded');};e.h.flushPendingEdits(false);assert.equal(e.calls.length,0);
});
