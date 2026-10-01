import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';

function page(){
 const html=fs.readFileSync('Scoring.html','utf8'),visible=html.split('<template id="ruleBuilderTemplate">')[0],elements=new Map(),events={},calls=[],storage=new Map();
 const element=()=>({value:'',textContent:'',innerHTML:'',style:{},dataset:{},children:[],classList:{add(){},remove(){},contains(){return false;}},addEventListener(name,handler){const previous=this['on'+name];this['on'+name]=function(...args){if(previous)previous.apply(this,args);handler.apply(this,args);};},setAttribute(){},getAttribute(){return '';},appendChild(e){this.children.push(e);return e;},replaceChildren(){this.children=[];},get options(){return this.children;},querySelectorAll(){return [];},querySelector(){return null;},focus(){},select(){},showModal(){this.open=true;},close(){this.open=false;}});
 for(const match of visible.matchAll(/\bid="([^"]+)"/g))elements.set(match[1],element());
 const groups={'.score-input':Array.from({length:5},element)};
 const document={body:element(),readyState:'loading',getElementById:id=>elements.get(id)||null,createElement:element,createDocumentFragment:element,createTextNode:text=>({textContent:text}),querySelectorAll:s=>groups[s]||[],addEventListener:(name,fn)=>{events[name]=fn;}};
 const window={__SCORING_TOOL_ENABLE_TEST_HOOKS__:true,addEventListener(name,fn){events["window:"+name]=fn;}};
 const c=vm.createContext({document,window,console,localStorage:{getItem:k=>storage.get(k)||null,setItem:(k,v)=>storage.set(k,v),removeItem:k=>storage.delete(k)},setTimeout:()=>1,clearTimeout(){},google:{script:{get run(){const run={withSuccessHandler(fn){this.success=fn;return this;},withFailureHandler(fn){this.failure=fn;return this;}};return new Proxy(run,{get:(o,k)=>k in o?o[k]:payload=>calls.push({method:k,payload,success:o.success,failure:o.failure})});}}}});
 const script=html.match(/<script>([\s\S]*?)<\/script>/)[1].replace('window.__scoringToolTestHooks = {','window.__scoringToolTestHooks = {state,els,loadRuleTemplates_,resolveRuleOutputSlots_,loadInitData,saveConfig,markDirty,reviewLabel_,comparisonCard_,handleScoresChanged,finishPendingFeedback_,saveCurrentRow,loadPendingEditFromLocalStorage,loadCurrentTemplateRules,renderRulesSummary,openComparison_,adoptComparison_,');
 vm.runInContext(script.replace('function tokenizeWhenExpr(expr) {','function tokenizeWhenExpr(expr) { window.parseCount=(window.parseCount||0)+1;'),c);return {events,calls,elements,window,storage,h:window.__scoringToolTestHooks};
}

function readyPage(){
 const p=page();p.events.DOMContentLoaded();
 Object.assign(p.h.state,{config:{headerRow:1,startRow:2,scoreCols:['D','','','',''],colChecks:[false,true,false,false,false],displayCols:['G'],ruleOutputSlots:[]},colChecks:[false,true,false,false,false],operationId:'operation',currentSheet:'["book",1]',sheetNames:['["book",1]'],configRevision:'revision-1'});
 p.response=()=>({sheetName:p.h.state.currentSheet,config:JSON.parse(JSON.stringify(p.h.state.config)),configRevision:p.h.state.configRevision,rows:[{rowNumber:2,name:'生徒',displays:['回答'],scores:['','','','',''],comment:''}],headers:[],allHeaders:[{col:'D',header:'点数',label:'点数'},{col:'E',header:'新点数',label:'新点数'}],review:{rows:{},groups:[]}});
 p.load=()=>{p.h.loadInitData();p.calls.at(-1).success(p.response());};p.load();return p;
}

test('successful loads restore editable scores while respecting locked columns',()=>{
 const p=readyPage();assert.deepEqual(Array.from(p.h.els.scoreInputs,e=>e.disabled),[false,true,false,false,false]);
 p.load();assert.equal(p.h.els.scoreInputs[0].disabled,false);
});

test('applying start row preserves the final dirty edit before any reload and retains the draft status',()=>{
 const p=readyPage();p.h.els.scoreInputs[0].disabled=false;p.h.els.scoreInputs[0].value='5';p.h.markDirty();
 p.elements.get('applyStartBtn').onclick();p.calls.filter(c=>c.method==='scoringGetInitData').at(-1).success(p.response());
 assert.equal(p.h.els.scoreInputs[0].value,'5');assert.ok([...p.storage.values()].some(v=>v.includes('revision-1')));
 assert.match(p.elements.get('statusText').textContent,/未送信|未反映/);
});

test('a failed local save prevents a reload from discarding edits',()=>{
 const p=readyPage();p.h.els.scoreInputs[0].disabled=false;p.h.els.scoreInputs[0].value='5';p.h.markDirty();
 p.storage.set=()=>{throw Error('quota');};const n=p.calls.length;p.h.loadInitData();
 assert.equal(p.calls.length,n);assert.equal(p.h.els.scoreInputs[0].value,'5');
});

test('changing only a score column reloads the write revision before new grading',()=>{
 const p=readyPage();p.h.els.scoreHeaderSelects[0].value='新点数';p.h.saveConfig();
 const save=p.calls.at(-1);assert.equal(save.method,'scoringSetConfig');save.success({...save.payload});
 assert.equal(p.calls.at(-1).method,'scoringGetInitData');p.calls.at(-1).success({...p.response(),configRevision:'revision-2'});
 assert.equal(p.h.state.configRevision,'revision-2');
});

test('a stale load error cannot overwrite a newer successful view',()=>{
 const p=readyPage();p.h.loadInitData();const stale=p.calls.at(-1);p.load();const status=p.elements.get('statusText').textContent;
 stale.failure(Error('stale'));assert.equal(p.elements.get('statusText').textContent,status);
});

test('the independent grading page initializes without removed standalone connection controls',()=>{
 const p=page();p.events.DOMContentLoaded();assert.equal(p.calls[0].method,'getWebConsoleBootstrap');
 p.calls[0].success({connected:true,spreadsheetId:'operation',settingsUrl:'https://script.google.com/macros/s/test/exec',url:'https://docs.google.com/spreadsheets/d/operation/edit'});
 assert.equal(p.calls[1].method,'scoringGetConfig');assert.match(p.elements.get('scoringSettingsLink').href,/\/exec$/);
});

test('comparison escapes student content and clearly shows held and returned states',()=>{
 const p=page();const card=p.h.comparisonCard_({name:'生徒',target:{sheetName:'回答'},rowNumber:2,state:'pending',returnState:'済',date:'2026-09-29',answers:[{header:'回答',value:'<img src=x onerror=alert(1)>'}]});
 assert.match(card,/返却保留/);assert.match(card,/返却: 済/);assert.match(card,/&lt;img/);assert.doesNotMatch(card,/<img/);
});

function rulesPage(){const p=readyPage();p.h.state.config.ruleTemplateId='test';p.h.state.currentRulesHeaders=['whenExpr','message','target'];p.h.state.currentRules=[{whenExpr:'',message:'auto A',target:'A'},{whenExpr:'',message:'auto B',target:'B'},{whenExpr:'',message:'auto C',target:'C'}];p.h.state.ruleOutputSlotsResolved=[{target:'A'},{target:'B'},{target:'C'}];return p;}
test('condition parsing is reused but results follow scores and the cache expires on template reload',()=>{
 const p=rulesPage(),expr='score1 >= 3';
 assert.equal(p.h.evaluateWhenExprSafe(expr,{score1:3}).result,true);
 assert.equal(p.h.evaluateWhenExprSafe(expr,{score1:2}).result,false);
 assert.equal(p.window.parseCount,1);
 p.h.loadCurrentTemplateRules('next');
 assert.equal(p.h.evaluateWhenExprSafe(expr,{score1:4}).result,true);
 assert.equal(p.window.parseCount,2);
 for(let i=0;i<2;i++)assert.equal(p.h.evaluateWhenExprSafe('score1 == "3',{score1:3}).ok,false);
 assert.equal(p.window.parseCount,3);
});
test('pending re-evaluation is saved under the old target before switching sheets',()=>{
 const p=rulesPage();p.h.handleScoresChanged();p.h.els.sheetSelect.value='["other-book",2]';p.h.els.sheetSelect.onchange();
 assert.equal(p.h.loadPendingEditFromLocalStorage('["book",1]',2).comment,'auto A');assert.equal(p.h.loadPendingEditFromLocalStorage('["other-book",2]',2),null);
});
test('explicit re-evaluation marks its generated feedback dirty',()=>{const p=rulesPage();p.h.handleScoresChanged();p.h.finishPendingFeedback_();assert.equal(p.h.state.dirty,true);});
test('unavailable rules preserve feedback and delay saving until the same row can be evaluated',()=>{
 const p=rulesPage();p.h.els.ruleOutInputs[0].value='existing';p.h.loadCurrentTemplateRules('test');const request=p.calls.at(-1);
 p.h.handleScoresChanged();p.h.finishPendingFeedback_();assert.equal(p.h.els.ruleOutInputs[0].value,'existing');assert.equal(p.h.saveCurrentRow(),false);
 request.failure(Error('offline'));assert.equal(p.h.els.ruleOutInputs[0].value,'existing');assert.equal(p.h.saveCurrentRow(),false);
 p.h.loadCurrentTemplateRules('test');p.calls.at(-1).success({ok:true,headers:['whenExpr','message'],rules:[{whenExpr:'',message:'ready'}]});assert.equal(p.h.els.ruleOutInputs[0].value,'ready');assert.equal(p.h.saveCurrentRow(),true);
});
test('manual feedback cancels pending evaluation including a late rule load',()=>{
 const p=rulesPage();p.h.loadCurrentTemplateRules('test');const request=p.calls.at(-1);p.h.handleScoresChanged();p.h.finishPendingFeedback_();
 p.h.els.ruleOutInputs[0].value='manual';p.h.els.ruleOutInputs[0].oninput();request.success({ok:true,headers:['whenExpr','message'],rules:[{whenExpr:'',message:'automatic'}]});p.h.finishPendingFeedback_();assert.equal(p.h.els.ruleOutInputs[0].value,'manual');assert.equal(p.h.saveCurrentRow(),true);
});
test('reapplying slot one does not replace manually edited slots two and three',()=>{
 const p=rulesPage();p.h.els.ruleOutInputs.forEach((e,i)=>{e.value='manual '+i;});p.h.els.reapplyRulesBtn.onclick();p.h.finishPendingFeedback_();assert.deepEqual(Array.from(p.h.els.ruleOutInputs,e=>e.value),['auto A','manual 1','manual 2']);
});
test('rule evaluation and rule summary share one fetch',()=>{const p=rulesPage();const before=p.calls.length;p.h.loadCurrentTemplateRules('test');p.h.renderRulesSummary('test');assert.equal(p.calls.length-before,1);});
test('grading loads summaries and comparison details are requested only when opened',()=>{
 const p=readyPage();assert.equal(p.calls.filter(c=>c.method==='scoringGetInitData').at(-1).payload.reviewMode,'summary');p.h.openComparison_('group','row');assert.equal(p.calls.at(-1).method,'scoringGetComparison');
 const request=p.calls.at(-1);p.elements.get('comparisonClose').onclick();request.success({id:'group',responses:[]});assert.equal(p.elements.get('comparisonDialog').open,false);assert.equal(p.h.state.comparison,null);
});

test('closing the page warns when rule loading prevents local persistence',()=>{const p=rulesPage();p.h.state.currentRules=null;p.h.handleScoresChanged();let prevented=false;const event={preventDefault(){prevented=true;}};p.events['window:beforeunload'](event);assert.equal(prevented,true);assert.equal(event.returnValue,'');});
test('retrying one slot after rule fetch failure preserves the other manual slots',()=>{
 const p=rulesPage();p.h.state.currentRules=null;p.h.state.currentRulesResult={ok:false};p.h.els.ruleOutInputs.forEach((e,i)=>{e.value='manual '+i;});p.h.els.reapplyRulesBtn.onclick();assert.equal(p.calls.at(-1).method,'scoringGetRules');
 p.calls.at(-1).success({ok:true,headers:['whenExpr','message','target'],rules:[{whenExpr:'',message:'retried',target:'講評・改善点'}]});assert.deepEqual(Array.from(p.h.els.ruleOutInputs,e=>e.value),['retried','manual 1','manual 2']);
});
test('adoption updates the summary in one RPC',()=>{
 const p=readyPage();p.h.state.comparison={group:{id:'group',fingerprint:'fingerprint'},current:'one',other:'two'};p.elements.get('comparisonDialog').open=true;const before=p.calls.length;p.h.adoptComparison_(['one']);const request=p.calls.at(-1);assert.equal(request.payload.reviewMode,'summary');request.success({ok:true,review:{rows:{},groups:[],detailsDeferred:true}});assert.equal(p.calls.length,before+1);assert.equal(p.elements.get('comparisonDialog').open,false);
});

test('comparison rejects a fresh group that no longer contains the displayed response',()=>{const p=readyPage();p.h.openComparison_('group','old-row');p.calls.at(-1).success({id:'group',responses:[{key:'other-row'}]});assert.equal(p.h.state.comparison,null);assert.match(p.elements.get('comparisonMessage').textContent,/読み直/);});

test('rule evaluation skips conditions for filled or unused output targets',()=>{
 const p=rulesPage();let evaluated=0;const skipped=()=>({target:'A',get whenExpr(){evaluated++;return 'score1 > 0';},message:'later'});const rules=[{target:'A',whenExpr:'',message:'first'},...Array.from({length:2000},skipped),{target:'unused',get whenExpr(){evaluated++;return 'score1 > 0';},message:'unused'}];
 const result=p.h.evaluateRulesLocallyMulti_(rules,['target','whenExpr','message'],{score1:5},[{target:'A'}]);assert.equal(result.outputs[0],'first');assert.equal(evaluated,0);
});


test('table outputs keep fixed positions and reject invalid scores without replacing manual feedback',()=>{
 const p=readyPage();p.h.state.config.ruleTemplateId='表';
 const r={target:'出力3',whenExpr:'score1>=0',message:'値{枠1}',enabled:1,_tableConditions:[{op:'>=',value:0},{op:'any'},{op:'any'},{op:'any'},{op:'any'}],_tableDomains:[[0,.5,1],[],[],[],[]]};
 const headers=['target','whenExpr','message','enabled'];p.h.state.currentRules=[r];p.h.state.currentRulesHeaders=headers;
 const slots=p.h.resolveRuleOutputSlots_({ruleOutputSlots:[{col:'E',enabled:true},{col:'F',enabled:true},{col:'G',enabled:true}]},[r],headers);
 assert.deepEqual(Array.from(slots,s=>s.target),['出力1','出力2','出力3']);assert.equal(slots[2].col,'G');
 const evaluate=value=>p.h.evaluateRulesLocallyMulti_([r],headers,p.h.buildEvalVars(p.h.buildRuleVars_([value,'','','',''])),slots);
 assert.deepEqual(Array.from(evaluate('.5').outputs),['','','値0.5']);assert.equal(evaluate('1.').outputs[2],'値1');assert.equal(evaluate('').outputs[2],'');assert.match(evaluate('oops').error,/数値/);
 p.h.state.ruleOutputSlotsResolved=slots;p.h.els.ruleOutInputs[0].value='手直し';p.h.els.scoreInputs[0].value='oops';p.h.handleScoresChanged();p.h.finishPendingFeedback_();assert.equal(p.h.els.ruleOutInputs[0].value,'手直し');
});

test('deleting the selected template clears cached conditions and preserves manual feedback',()=>{
 const p=rulesPage();p.h.els.ruleOutInputs[0].value='手直し';p.h.loadRuleTemplates_('test');p.calls.at(-1).success({ok:true,headers:['templateId','name','enabled'],templates:[]});
 assert.equal(p.h.state.config.ruleTemplateId,'');assert.equal(p.h.state.currentRules,null);
 p.h.handleScoresChanged();p.h.finishPendingFeedback_();assert.equal(p.h.els.ruleOutInputs[0].value,'手直し');
});
