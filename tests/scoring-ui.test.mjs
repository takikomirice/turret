import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';

function page(){
 const html=fs.readFileSync('Scoring.html','utf8'),visible=html.split('<template id="ruleBuilderTemplate">')[0],elements=new Map(),events={},calls=[],storage=new Map();
 const element=()=>({value:'',textContent:'',innerHTML:'',style:{},dataset:{},children:[],classList:{add(){},remove(){},contains(){return false;}},addEventListener(name,handler){const previous=this['on'+name];this['on'+name]=function(...args){if(previous)previous.apply(this,args);handler.apply(this,args);};},setAttribute(){},getAttribute(){return '';},appendChild(e){this.children.push(e);return e;},replaceChildren(){this.children=[];},get options(){return this.children;},querySelectorAll(){return [];},querySelector(){return null;},focus(){},select(){},click(){this.clicked=true;},remove(){},showModal(){this.open=true;},close(){this.open=false;}});
 for(const match of visible.matchAll(/\bid="([^"]+)"/g))elements.set(match[1],element());
 const groups={'.score-input':Array.from({length:5},element)};
 const document={body:element(),readyState:'loading',getElementById:id=>elements.get(id)||null,createElement:element,createDocumentFragment:element,createTextNode:text=>({textContent:text}),querySelectorAll:s=>groups[s]||[],addEventListener:(name,fn)=>{events[name]=fn;}};
 const window={__SCORING_TOOL_ENABLE_TEST_HOOKS__:true,addEventListener(name,fn){events["window:"+name]=fn;}};
 const c=vm.createContext({document,window,console,localStorage:{getItem:k=>storage.get(k)||null,setItem:(k,v)=>storage.set(k,v),removeItem:k=>storage.delete(k)},setTimeout:()=>1,clearTimeout(){},google:{script:{get run(){const run={withSuccessHandler(fn){this.success=fn;return this;},withFailureHandler(fn){this.failure=fn;return this;}};return new Proxy(run,{get:(o,k)=>k in o?o[k]:(...args)=>calls.push({method:k,payload:args[0],args,success:o.success,failure:o.failure})});}}}});
 const script=html.match(/<script>([\s\S]*?)<\/script>/)[1].replace('window.__scoringToolTestHooks = {','window.__scoringToolTestHooks = {state,els,loadRuleTemplates_,resolveRuleOutputSlots_,loadInitData,saveConfig,markDirty,reviewLabel_,comparisonCard_,handleScoresChanged,finishPendingFeedback_,saveCurrentRow,loadPendingEditFromLocalStorage,loadCurrentTemplateRules,renderRulesSummary,openComparison_,adoptComparison_,');
 vm.runInContext(script.replace('function tokenizeWhenExpr(expr) {','function tokenizeWhenExpr(expr) { window.parseCount=(window.parseCount||0)+1;'),c);return {events,calls,elements,window,storage,document,c,h:window.__scoringToolTestHooks};
}

test('scoring settings and template packs download the shared JSON files with server filenames',async()=>{
 const p=page();p.events.DOMContentLoaded();p.h.state.config={};const blobs=[];
 p.c.Blob=Blob;p.c.URL={createObjectURL:blob=>{blobs.push(blob);return 'blob:download';},revokeObjectURL(){}};
 for(const [button,kind] of [['exportConfigBtn','scoring'],['exportTemplatePackBtn','templates']]) {
  p.elements.get(button).onclick();const request=p.calls.at(-1);assert.equal(request.method,'exportConfigurationFile');assert.equal(request.payload,kind);
  request.success({filename:'turret-'+kind+'-化学-20261001-153000.json',json:'{"saved":true}'});
  const link=p.elements.get(kind==='scoring'?'exportResult':'templatePackResult').children.at(-1);assert.equal(link.download,'turret-'+kind+'-化学-20261001-153000.json');assert.equal(link.clicked,true);
 }
 assert.equal(await blobs[0].text(),'{"saved":true}');assert.equal(await blobs[1].text(),'{"saved":true}');
});

test('scoring imports inspect before confirmation and apply only the inspected file revision',()=>{
 const p=page();p.events.DOMContentLoaded();p.h.state.config={};const prompts=[];
 p.c.FileReader=class {readAsText(){this.onload({target:{result:'{"config":{}}'}});}};
 p.window.confirm=text=>{prompts.push(text);return true;};
 const input=p.elements.get('importConfigFileInput');input.files=[{size:13}];input.onchange({target:input});
 const inspection=p.calls.at(-1);assert.equal(inspection.method,'inspectConfigurationFile');assert.deepEqual(inspection.args,['scoring','{"config":{}}']);assert.equal(prompts.length,0);
 inspection.success({summary:'採点設定を置き換えます',warnings:['列を確認してください'],revision:'file-revision'});
 const apply=p.calls.at(-1);assert.equal(apply.method,'applyConfigurationFile');assert.deepEqual(apply.args,['scoring','{"config":{}}','file-revision']);assert.match(prompts[0],/置き換え/);assert.match(prompts[0],/列を確認/);
});

test('template import cancellation and oversized file selection do not apply or discard grading edits',()=>{
 const p=page();p.events.DOMContentLoaded();p.h.state.config={};p.c.FileReader=class {readAsText(){this.onload({target:{result:'{}'}});}};
 p.window.confirm=()=>false;
 const input=p.elements.get('importTemplatePackFileInput');input.files=[{size:2}];input.onchange({target:input});const inspection=p.calls.at(-1);
 assert.equal(inspection.method,'inspectConfigurationFile');assert.equal(inspection.payload,'templates');
 const n=p.calls.length;inspection.success({summary:'追加します',warnings:[],revision:'file-revision'});assert.equal(p.calls.length,n);
 input.files=[{size:1048577}];input.onchange({target:input});assert.equal(p.calls.length,n);assert.match(p.elements.get('templatePackResult').textContent,/1MiB/);
});

function readyPage(){
 const p=page();p.events.DOMContentLoaded();
 Object.assign(p.h.state,{config:{headerRow:1,startRow:2,scoreCols:['D','','','',''],colChecks:[false,true,false,false,false],displayCols:['G'],ruleOutputSlots:[]},colChecks:[false,true,false,false,false],operationId:'operation',currentSheet:'["book",1]',sheetNames:['["book",1]'],configRevision:'revision-1'});
 p.response=()=>({sheetName:p.h.state.currentSheet,config:JSON.parse(JSON.stringify(p.h.state.config)),configRevision:p.h.state.configRevision,rows:[{rowNumber:2,name:'生徒',displays:['回答'],scores:['','','','',''],comment:''}],headers:[],allHeaders:[{col:'D',header:'点数',label:'点数'},{col:'E',header:'新点数',label:'新点数'}],review:{rows:{},groups:[]}});
 p.load=()=>{p.h.loadInitData();p.calls.at(-1).success(p.response());};p.load();return p;
}

function namesPage(){
 const p=readyPage();p.h.state.targets={[p.h.state.currentSheet]:{namesAvailable:true}};p.load();return p;
}

test('scoring name enrichment preserves grading drafts and refreshes the displayed student name',()=>{
 const p=namesPage(),button=p.elements.get('refreshNamesBtn');assert.ok(button);assert.equal(button.disabled,false);
 p.h.els.scoreInputs[0].value='5';p.h.markDirty();button.onclick();
 const request=p.calls.at(-1);assert.equal(request.method,'scoringRefreshNames');assert.equal(request.payload,'["book",1]');
 assert.equal(button.disabled,true);const count=p.calls.length;button.onclick();assert.equal(p.calls.length,count);
 p.h.els.scoreInputs[0].value='6';p.h.markDirty();
 request.success({updated:1,unmatched:2,skipped:0});const reload=p.calls.at(-1);assert.equal(reload.method,'scoringGetInitData');
 const response=p.response();response.rows[0].name='補完した名前';reload.success(response);
 assert.match(p.elements.get('studentName').textContent,/補完した名前/);assert.equal(p.h.els.scoreInputs[0].value,'6');
 assert.match(p.elements.get('nameRefreshStatus').textContent,/補完 1件.*未照合 2件/);assert.equal(button.disabled,false);
});

test('name enrichment failures retain inputs and allow retry; unsupported targets do not send requests',()=>{
 const p=namesPage(),button=p.elements.get('refreshNamesBtn');assert.ok(button);
 p.h.els.scoreInputs[0].value='5';p.h.markDirty();button.onclick();p.calls.at(-1).failure(Error('名簿がありません'));
 assert.equal(button.disabled,false);assert.equal(p.h.els.scoreInputs[0].value,'5');assert.match(p.elements.get('nameRefreshStatus').textContent,/名簿がありません/);
 p.h.state.targets[p.h.state.currentSheet].namesAvailable=false;p.load();assert.equal(button.disabled,true);
 const count=p.calls.length;button.onclick();assert.equal(p.calls.length,count);
});

test('a name enrichment response for the previous target cannot reload a newly selected sheet',()=>{
 const p=namesPage(),button=p.elements.get('refreshNamesBtn');assert.ok(button);button.onclick();const request=p.calls.at(-1);
 p.h.state.currentSheet='["other",2]';p.h.state.targets[p.h.state.currentSheet]={namesAvailable:true};p.load();const count=p.calls.length;
 request.success({updated:1,unmatched:0,skipped:0});assert.equal(p.calls.length,count);assert.equal(p.h.state.currentSheet,'["other",2]');
 assert.equal(p.elements.get('nameRefreshStatus').textContent,'');assert.equal(button.disabled,false);
});

test('loading, grading saves and failed local persistence prevent name enrichment requests',()=>{
 for(const blocked of ['loading','saving','storage']){
  const p=namesPage(),button=p.elements.get('refreshNamesBtn');assert.ok(button);
  if(blocked==='storage'){p.h.els.scoreInputs[0].value='5';p.h.markDirty();p.storage.set=()=>{throw Error('quota');};}
  else p.h.state[blocked]=true;
  const count=p.calls.length;button.onclick();assert.equal(p.calls.length,count);
 }
});

test('lesson timing and duplicate status have independent warning colors across row changes',()=>{
 const p=readyPage();
 const contents=el=>el.textContent+(el.children || []).map(contents).join('');
 for(const [daysLate,status,warnDate,warnReview,label] of [[0,'single',false,false,'当日提出'],[2,'single',true,false,'2日遅れ'],[0,'pending',false,true,'当日提出'],[99,'adopted',true,true,'99日遅れ'],[100,'excluded',true,true,'99日超'],[-365,'single',false,false,'授業日前'],[null,'invalid-date',false,false,'日付不明']]) {
  p.h.loadInitData();const response=p.response();
  response.review.rows[2]={state:status,returnState:'未返却',timing:{lessonDate:daysLate===null?'':'2026-09-29',submittedAt:'2026-10-01 10:00:00',daysLate}};
  p.calls.at(-1).success(response);
  const [date,details]=p.elements.get('duplicateStatus').children;
  assert.equal(date.className,'submission-timing');assert.match(contents(date),new RegExp(label));
  assert.equal(date.children.some(el=>el.className==='status-warning'),warnDate);
  assert.equal(details.children[0].className==='status-warning',warnReview);
  assert.notEqual(details.children[1].className,'status-warning');
  assert.match(date.title,/提出日時: 2026-10-01 10:00:00/);
  if(daysLate===100)assert.match(date.title,/100日遅れ/);
 }
});

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
