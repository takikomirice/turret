import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import {readFileSync} from 'node:fs';
const html=readFileSync('Scoring.html','utf8');
function helpModel(source){
 const handlers={},timers=new Map(),attrs={};let next=0;
 const tip={style:{},hidden:true,offsetWidth:300,offsetHeight:180,setAttribute(){},contains:t=>t===tip};
 const host={appendChild(){}};
 const button={disabled:false,dataset:{helpTitle:'入力の説明',helpBody:'<b>編集しない</b>'},closest:s=>s==='dialog'?host:button,getAttribute:()=>null,setAttribute:(k,v)=>attrs[k]=v,removeAttribute:k=>delete attrs[k],getBoundingClientRect:()=>({left:480,top:460,bottom:478})};
 const document={body:host,createElement:()=>tip,getElementById:()=>null,addEventListener:(n,fn)=>handlers[n]=fn};
 const window={innerWidth:540,innerHeight:540,addEventListener:(n,fn)=>handlers[n]=fn};
 const c=vm.createContext({document,window,setTimeout:fn=>{timers.set(++next,fn);return next;},clearTimeout:id=>timers.delete(id)});
 assert.ok(source,'画面にヘルプ制御がある');vm.runInContext(source+';initHelpPopovers_();',c);
 return {handlers,tip,button,attrs,timers};
}
const controllers=[...html.matchAll(/function initHelpPopovers_\(\) \{[\s\S]*?\n\}/g)].map(m=>m[0]);
for(const [i,label] of ['scoring','template editor'].entries()){
 test(label+' help opens on hover and focus, stays inside viewport, and dismisses with Escape',()=>{
  const p=helpModel(controllers[i]);p.handlers.pointerover({target:p.button});
  assert.equal(p.tip.hidden,false);assert.equal(p.tip.textContent,'入力の説明\n<b>編集しない</b>');
  assert.equal(p.tip.style.left,'228px');assert.equal(p.tip.style.top,'272px');assert.equal(p.attrs['aria-describedby'],'scoringHelpTooltip');
  p.handlers.keydown({key:'Escape'});assert.equal(p.tip.hidden,true);assert.equal(p.attrs['aria-describedby'],undefined);
  p.handlers.focusin({target:p.button});assert.equal(p.tip.hidden,false);
  p.handlers.resize();assert.equal(p.tip.hidden,true);
 });
 test(label+' help pins on tap without activating a label or losing text while crossing the gap',()=>{
  const p=helpModel(controllers[i]);let prevented=0;
  p.handlers.click({target:p.button,preventDefault(){prevented++;},stopPropagation(){}});
  assert.equal(prevented,1);p.handlers.pointerout({target:p.button});assert.equal(p.timers.size,0);
  p.handlers.click({target:p.button,preventDefault(){},stopPropagation(){}});assert.equal(p.tip.hidden,true);
  p.handlers.pointerover({target:p.button});p.handlers.pointerout({target:p.button});assert.equal(p.timers.size,1);
  p.tip.closest=()=>null;p.handlers.pointerover({target:p.tip});assert.equal(p.timers.size,0);
 });
}
