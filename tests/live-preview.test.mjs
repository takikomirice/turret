import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import {readFileSync} from 'node:fs';

function model(){
 const timers=new Map(),calls=[],nodes=Object.fromEntries(['previewText','previewStatus','previewWarnings'].map(id=>[id,{textContent:'',hidden:false,classList:{toggle(){}}}]));
 let nextTimer=0;
 const events={},c=vm.createContext({console,document:{getElementById:id=>nodes[id],addEventListener:(type,fn)=>events[type]=fn},window:{addEventListener(){}},setTimeout:fn=>{timers.set(++nextTimer,fn);return nextTimer;},clearTimeout:id=>timers.delete(id)});
 const script=readFileSync('Setting.html','utf8').match(/<script>([\s\S]*?)<\/script>/)[1];
 vm.runInContext(script.replace(/boot\(\);\s*$/,''),c);
 c.calls=calls;
 vm.runInContext("state.panel='setup';state.step='template';state.data={config:{fields:[],messageTemplate:'初期本文'},revision:'r1'};rpc=(method,payload)=>new Promise((resolve,reject)=>calls.push({method,payload,resolve,reject}));render=()=>{throw Error('preview must not render');};setBusy=()=>{throw Error('preview must not block input');};",c);
 return {c,calls,nodes,events,run:code=>vm.runInContext(code,c),flush:async()=>{const batch=[...timers.values()];timers.clear();batch.forEach(fn=>fn());await Promise.resolve();},settle:async()=>{await Promise.resolve();await Promise.resolve();}};
}

test('live preview debounces edits and includes unsaved conversion rules',async()=>{
 const m=model();
 m.run("scheduleTemplatePreview();putDraft('template',{messageTemplate:'最新本文'});const p=draftValue('fields');p.gradeScale=[{from:'5',to:'優'}];putDraft('fields',p);scheduleTemplatePreview();");
 assert.equal(m.calls.length,0);
 await m.flush();
 assert.equal(m.calls.length,1);
 assert.equal(m.calls[0].payload.messageTemplate,'最新本文');
 assert.equal(m.calls[0].payload.gradeScale[0].to,'優');
 m.calls[0].resolve({text:'最新の仕上がり',warnings:[]});await m.settle();
 assert.equal(m.nodes.previewText.textContent,'最新の仕上がり');
 assert.equal(m.run('state.busy'),false);
});

test('one in-flight preview queues the newest edit and ignores stale success',async()=>{
 const m=model();m.run('scheduleTemplatePreview(true)');
 m.run("putDraft('template',{messageTemplate:'途中'});scheduleTemplatePreview();");await m.flush();
 m.run("putDraft('template',{messageTemplate:'最終'});scheduleTemplatePreview();");await m.flush();
 assert.equal(m.calls.length,1);
 m.calls[0].resolve({text:'古い結果',warnings:['古い警告']});await m.settle();
 assert.equal(m.calls.length,2);
 assert.equal(m.calls[1].payload.messageTemplate,'最終');
 assert.notEqual(m.nodes.previewText.textContent,'古い結果');
 m.calls[1].resolve({text:'<b>最終結果</b>',warnings:['未知の差し込み']});await m.settle();
 assert.equal(m.nodes.previewText.textContent,'<b>最終結果</b>');
 assert.equal(m.nodes.previewWarnings.textContent,'未知の差し込み');
 m.run('scheduleTemplatePreview()');await m.flush();assert.equal(m.calls.length,2);
});

test('leaving the screen cancels pending previews and discards late failures',async()=>{
 const m=model();m.run('scheduleTemplatePreview()');m.run("state.step='fields';cancelTemplatePreview();");await m.flush();assert.equal(m.calls.length,0);
 m.run("state.step='template';scheduleTemplatePreview(true);");
 m.run("state.step='fields';cancelTemplatePreview();");
 m.calls[0].reject(Error('古い失敗'));await m.settle();
 assert.doesNotMatch(m.nodes.previewStatus.textContent,/古い失敗/);
 m.run("state.step='template';scheduleTemplatePreview(true);");assert.equal(m.calls.length,2);
 m.calls[1].reject(Error('通信失敗'));await m.settle();assert.match(m.nodes.previewStatus.textContent,/通信失敗/);
 m.run('scheduleTemplatePreview(true)');assert.equal(m.calls.length,3);
 m.calls[2].resolve({text:'',warnings:[]});await m.settle();assert.match(m.nodes.previewText.textContent,/本文は空/);
});

test('late failure waits for the newest edit debounce and cannot replace it',async()=>{
 const m=model();m.run('scheduleTemplatePreview(true)');
 m.run("putDraft('template',{messageTemplate:'新本文'});scheduleTemplatePreview();");
 m.calls[0].reject(Error('古い失敗'));await m.settle();
 assert.equal(m.calls.length,1);assert.doesNotMatch(m.nodes.previewStatus.textContent,/古い失敗/);
 await m.flush();assert.equal(m.calls.length,2);
 m.calls[1].resolve({text:'新本文',warnings:[]});await m.settle();assert.equal(m.nodes.previewText.textContent,'新本文');
});

test('reentering template uses changed unsaved rules instead of the cached result',async()=>{
 const m=model();m.run('scheduleTemplatePreview(true)');m.calls[0].resolve({text:'以前の評価',warnings:[]});await m.settle();
 m.run("state.step='fields';cancelTemplatePreview();const p=draftValue('fields');p.gradeScale=[{from:'5',to:'優秀'}];putDraft('fields',p);state.step='template';scheduleTemplatePreview();");
 assert.notEqual(m.nodes.previewText.textContent,'以前の評価');await m.flush();assert.equal(m.calls[1].payload.gradeScale[0].to,'優秀');
});

test('IME composition waits until confirmation and preserves the newest text',async()=>{
 const m=model();
 m.nodes.confirmCheck={addEventListener(){}};
 m.run("initHelp=()=>{};task=()=>{};readCurrent=()=>putDraft('template',{messageTemplate:'確定した日本語'});boot();");
 const target={id:'messageTemplate',closest:()=>true};
 m.events.compositionstart({target});m.events.input({target,isComposing:true});await m.flush();assert.equal(m.calls.length,0);
 m.events.compositionend({target});m.events.input({target,isComposing:false});await m.flush();
 assert.equal(m.calls.length,1);assert.equal(m.calls[0].payload.messageTemplate,'確定した日本語');
});

test('token insertion updates preview without replacing the focused input',async()=>{
 const m=model();
 m.nodes.messageTemplate={value:'本文',selectionStart:1,selectionEnd:1,focused:false,setRangeText(token,start,end){this.value=this.value.slice(0,start)+token+this.value.slice(end);this.selectionStart=this.selectionEnd=start+token.length;},focus(){this.focused=true;}};
 m.run("readCurrent=()=>putDraft('template',{messageTemplate:$('messageTemplate').value});");
 const button={dataset:{token:'student_name'},disabled:false};
 m.c.click={target:{closest:()=>button}};
 await m.run('handleClick(click)');await m.flush();
 assert.equal(m.calls[0].payload.messageTemplate,'本{student_name}文');
 m.calls[0].resolve({text:'本山田 太郎（例）文',warnings:[]});await m.settle();
 assert.equal(m.nodes.messageTemplate.focused,true);assert.equal(m.nodes.messageTemplate.selectionStart,15);
});
