import {test} from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import {readFileSync} from 'node:fs';
import {scoringEnvironment} from './helpers/scoring-environment.mjs';
import {plain} from './helpers/admin-environment.mjs';

const source=readFileSync('Setting.html','utf8').match(/<script>([\s\S]*?)<\/script>/)[1].replace(/boot\(\);\s*$/,'');
function client() {
 const backend=scoringEnvironment(),elements=new Map(),selectors=new Map(),timers=new Map(),calls=[],messages=[];let tick=0;
 const c=vm.createContext({console,structuredClone,setTimeout:fn=>{timers.set(++tick,fn);return tick;},clearTimeout:id=>timers.delete(id),document:{getElementById:id=>elements.get(id)||null,querySelectorAll:s=>selectors.get(s)||[]}});
 vm.runInContext(source,c);vm.runInContext('render=()=>{};state.panel="grades";',c);
 c.status=(...args)=>messages.push(args);c.confirmAction=async()=>true;
 const g=vm.runInContext('gradeUI',c),profile=plain(backend.c.gradeDefault_());profile.base.valueHeader='点数';
 Object.assign(g,{loaded:true,profile,store:{revision:'',value:{draft:plain(profile),presets:[]}},data:{headers:['点数','授業日','名前'],classes:[{id:'a',name:'A',students:[]},{id:'b',name:'B',students:[]}]}});
 const server=async(method,...args)=>{
  calls.push({method,args:plain(args)});
  if(method==='gradePreview')return {revision:'preview',issues:[],responses:[],classes:[]};
  if(method==='gradeGetConsole')return {store:plain(backend.c.gradeReadStore_()),data:plain(g.data)};
  return plain(backend.c[method](...plain(args)));
 };
 c.rpc=server;
 const element=(id,value)=>{const el={id,value,dataset:{},matches:()=>true};elements.set(id,el);return el;};
 return {c,g,backend,elements,selectors,timers,calls,messages,server,element};
}

test('column reorder reads current selects and marks the result stale',async()=>{
 const e=client();e.element('gradeDateHeader','授業日');e.element('gradeValueHeader','点数');e.element('gradeSort','number');
 e.selectors.set('[data-grade-leading]',[{value:'roster:email'},{value:'field:名前'}]);e.g.result={revision:'old'};
 await e.c.gradeAction('up',{dataset:{index:'1'}});
 assert.deepEqual(plain(e.g.profile.base.leading),['field:名前','roster:email']);assert.equal(e.g.profile.base.sort,'number');assert.equal(e.g.dirty,true);assert.equal(e.g.result,null);
 assert.equal(e.timers.size,1);
});

test('class selection distinguishes all selected from no classes and captures period',()=>{
 const e=client();e.element('gradeStart','2026-09-01');e.element('gradeEnd','2026-09-30');
 e.selectors.set('[data-grade-class]:checked',[{dataset:{gradeClass:'a'}},{dataset:{gradeClass:'b'}}]);e.c.gradeReadInputs();assert.deepEqual(plain(e.g.profile.base.classIds),[]);
 e.selectors.set('[data-grade-class]:checked',[]);e.c.gradeReadInputs();assert.deepEqual(plain(e.g.profile.base.classIds),['__none__']);assert.equal(e.g.profile.base.start,'2026-09-01');
});

test('failed autosave retains current inputs, dirty state and revision, then retry succeeds',async()=>{
 const e=client();e.g.profile.base.sort='number';e.c.gradeChanged();e.c.rpc=async()=>{throw Error('offline');};
 await assert.rejects(e.c.gradePersist(),/offline/);assert.equal(e.g.dirty,true);assert.equal(e.g.store.revision,'');assert.equal(e.g.profile.base.sort,'number');
 e.c.rpc=e.server;await e.c.gradePersist();assert.equal(e.g.dirty,false);assert.equal(e.backend.c.gradeReadStore_().value.draft.base.sort,'number');assert.ok(e.g.store.revision);
});

test('autosave response cannot mark edits made during a pending request as saved',async()=>{
 const e=client();let finish;e.c.rpc=async(method,...args)=>{await new Promise(resolve=>{finish=resolve;});return e.server(method,...args);};
 e.g.profile.base.sort='number';e.c.gradeChanged();const first=e.c.gradePersist();await Promise.resolve();await Promise.resolve();
 e.g.profile.base.sort='roster';e.c.gradeChanged();finish();await first;
 assert.equal(e.g.dirty,true);assert.equal(e.g.profile.base.sort,'roster');assert.equal(e.backend.c.gradeReadStore_().value.draft.base.sort,'number');
 e.c.rpc=e.server;await e.c.gradePersist();assert.equal(e.g.dirty,false);assert.equal(e.backend.c.gradeReadStore_().value.draft.base.sort,'roster');
});

test('named preset excludes scope by default while preserving the current working scope',async()=>{
 const e=client();e.g.profile.base.classIds=['a'];e.g.profile.base.start='2026-09-01';e.g.profile.base.end='2026-09-30';e.g.profile.base.sort='number';e.c.gradeChanged();e.element('gradePresetName','年間');
 await e.c.gradeAction('save-as',{});const saved=plain(e.backend.c.gradeReadStore_().value);
 assert.equal(saved.presets.length,1);assert.equal(saved.presets[0].name,'年間');assert.deepEqual(saved.presets[0].profile.base.classIds,[]);assert.equal(saved.presets[0].profile.base.start,'');
 assert.deepEqual(saved.draft.base.classIds,['a']);assert.equal(saved.draft.base.start,'2026-09-01');assert.equal(saved.presets[0].profile.base.sort,'number');
 e.g.profile.base.sort='name';e.c.gradeChanged();await e.c.gradePersist();assert.equal(e.backend.c.gradeReadStore_().value.presets[0].profile.base.sort,'number');
});

test('load preset makes an independent draft and saves it without changing the preset',async()=>{
 const e=client();e.element('gradePresetName','年度');e.g.remember=true;e.g.profile.base.classIds=['a'];e.c.gradeChanged();await e.c.gradeAction('save-as',{});
 const id=e.g.preset;e.element('gradePreset',id);e.g.profile.base.sort='number';e.c.gradeChanged();await e.c.gradeAction('load-preset',{});
 assert.equal(e.g.profile.base.sort,'name');assert.deepEqual(plain(e.g.profile.base.classIds),['a']);assert.equal(e.g.step,1);
 e.g.profile.base.sort='roster';e.c.gradeChanged();await e.c.gradePersist();assert.equal(e.backend.c.gradeReadStore_().value.presets[0].profile.base.sort,'name');
});

test('date and per-answer rules replace only their own target and persist before analysis',async()=>{
 const e=client();e.g.result={issues:[{type:'date',courseId:'a',date:'2026-09-25'}],responses:[]};e.element('gradeDate0','2026-09-24');e.element('gradeMode0','planned');
 await e.c.gradeAction('date-rule',{dataset:{index:'0'}});
 assert.deepEqual(plain(e.g.profile.rules.dates),[{courseId:'a',from:'2026-09-25',to:'2026-09-24',mode:'planned'}]);assert.equal(e.calls.at(-2).method,'gradeSave');assert.equal(e.calls.at(-1).method,'gradePreview');
 e.g.result={issues:[],responses:[{id:'r1',fingerprint:'f1'}]};e.element('gradeDater0','');e.element('gradeModer0','extra');await e.c.gradeAction('answer-rule',{dataset:{index:'0'}});
 assert.deepEqual(plain(e.g.profile.rules.answers),[{id:'r1',fingerprint:'f1',to:'',mode:'extra'}]);assert.equal(e.g.profile.rules.dates.length,1);
});

test('ICS confirmation cancel retains schedule and acceptance preserves correction decisions',async()=>{
 const e=client();e.g.profile.calendar=[{id:'old'}];e.g.profile.rules.dates=[{courseId:'a',from:'2026-09-25',to:'2026-09-24',mode:'planned'}];const before=plain(e.g.profile.rules),next=[{id:'new',title:'A',date:'2026-09-28',time:'09:00'}];
 e.c.rpc=async()=>next;e.c.confirmAction=async()=>false;const event={target:{id:'gradeIcs',files:[{size:100,text:async()=> 'ics'}]}};await e.c.gradeChange(event);assert.deepEqual(plain(e.g.profile.calendar),[{id:'old'}]);
 e.c.confirmAction=async()=>true;await e.c.gradeChange(event);assert.deepEqual(plain(e.g.profile.calendar),next);assert.deepEqual(plain(e.g.profile.rules),before);assert.equal(e.g.dirty,true);
});

test('all class mappings are read together and persist as one draft',async()=>{
 const e=client();e.selectors.set('[data-grade-mapping]',[{dataset:{gradeMapping:'a'},value:'授業A'},{dataset:{gradeMapping:'b'},value:'授業B'}]);
 await e.c.gradeChange({target:{id:'gradeMap0',matches:()=>true}});await e.c.gradePersist();
 assert.deepEqual(plain(e.backend.c.gradeReadStore_().value.draft.base.mapping),{a:'授業A',b:'授業B'});
});

test('unselected duplicate choice changes nothing and a chosen adoption persists',async()=>{
 const e=client(),result={issues:[{type:'duplicate',key:'group',fingerprint:'fp',responses:[{id:'a'},{id:'b'}]}],responses:[]};e.g.result=result;
 e.element('gradeAssignment0_0','0');e.element('gradeAssignment0_1','');await assert.rejects(e.c.gradeAction('duplicate-rule',{dataset:{index:'0'}}),/扱い/);
 assert.equal(e.g.profile.rules.duplicates.length,0);assert.equal(e.g.result,result);assert.equal(e.calls.length,0);
 e.element('gradeAssignment0_1','1');await e.c.gradeAction('duplicate-rule',{dataset:{index:'0'}});
 assert.deepEqual(plain(e.backend.c.gradeReadStore_().value.draft.rules.duplicates),[{key:'group',fingerprint:'fp',assignments:{a:0,b:1}}]);
});

test('temporary rule save failure retains the decision and reanalysis retries it',async()=>{
 const e=client();e.g.result={issues:[{type:'date',courseId:'a',date:'2026-09-25'}],responses:[]};e.element('gradeDate0','2026-09-24');e.element('gradeMode0','planned');e.c.rpc=async()=>{throw Error('offline');};
 await assert.rejects(e.c.gradeAction('date-rule',{dataset:{index:'0'}}),/offline/);
 assert.equal(e.g.profile.rules.dates.length,1);assert.equal(e.g.dirty,true);assert.equal(e.g.store.revision,'');
 e.c.rpc=e.server;await e.c.gradeAction('analyze',{});assert.equal(e.g.dirty,false);assert.equal(e.g.result.revision,'preview');assert.equal(e.backend.c.gradeReadStore_().value.draft.rules.dates.length,1);
});

test('removing a saved correction preserves other kinds and persists before refreshing',async()=>{
 const e=client();e.g.profile.rules.dates=[{courseId:'a',from:'2026-09-25',to:'2026-09-24',mode:'planned'}];e.g.profile.rules.answers=[{id:'r',fingerprint:'f',to:'',mode:'extra'}];e.c.gradeChanged();await e.c.gradePersist();
 await e.c.gradeAction('remove-rule',{dataset:{kind:'dates',index:'0'}});const rules=plain(e.backend.c.gradeReadStore_().value.draft.rules);
 assert.equal(rules.dates.length,0);assert.equal(rules.answers.length,1);assert.equal(e.calls.at(-1).method,'gradePreview');
});

test('invalid date correction is rejected before poisoning draft or discarding review result',async()=>{
 const e=client(),result={issues:[{type:'date',courseId:'a',date:''}],responses:[]};e.g.result=result;e.element('gradeDate0','');e.element('gradeMode0','planned');const before=plain(e.g.profile);
 await assert.rejects(e.c.gradeAction('date-rule',{dataset:{index:'0'}}),/日付/);
 assert.deepEqual(plain(e.g.profile),before);assert.equal(e.g.result,result);assert.equal(e.calls.length,0);
});

test('preview arriving after a new edit cannot restore a stale actionable result',async()=>{
 const e=client();let finish,entered;const started=new Promise(resolve=>{entered=resolve;});e.c.rpc=async()=>new Promise(resolve=>{finish=resolve;entered();});const pending=e.c.gradeAnalyze();await started;
 e.g.profile.base.sort='number';e.c.gradeChanged();finish({revision:'old',issues:[{type:'date',courseId:'a',date:'2026-09-24'}],responses:[]});await pending;
 assert.equal(e.g.result,null);
});

test('date candidates prefer nearby lessons from the selected class only',()=>{
 const e=client();e.g.course='a';e.g.profile.base.mapping={a:'授業A',b:'授業B'};e.g.profile.calendar=[{title:'授業B',date:'2026-09-25'},{title:'授業A',date:'2026-09-28'},{title:'授業A',date:'2026-09-24'}];
 assert.deepEqual(plain(e.c.gradeDateCandidates('2026-09-25')),['2026-09-24','2026-09-28']);
});

test('grade period and correction inputs display weekdays while preserving date-only values',()=>{
 const e=client();e.g.profile.base.start='2026-10-09';e.g.profile.base.end='2026-10-15';
 let html=e.c.renderGrades();
 assert.match(html,/type="text"[^>]*value="2026\/10\/09（金）"/);assert.match(html,/type="text"[^>]*value="2026\/10\/15（木）"/);
 assert.match(html,/id="gradeStart"[^>]*value="2026-10-09"/);assert.match(html,/id="gradeEnd"[^>]*value="2026-10-15"/);
 html=e.c.gradeRuleControls('r0','2026-10-09');assert.match(html,/type="text"[^>]*value="2026\/10\/09（金）"/);
});

test('choosing a correction candidate refreshes its displayed weekday without changing the saved date format',async()=>{
 const e=client(),native=e.element('gradeDate0',''),display=e.element('date-display-gradeDate0','');native.type='date';native.dataset.dateNative='true';e.element('gradeMode0','extra');
 await e.c.gradeChange({target:{dataset:{gradeCandidate:'0'},value:'2026-10-09'}});
 assert.equal(native.value,'2026-10-09');assert.equal(display.value,'2026/10/09（金）');
});

test('an invalid visible grade date blocks navigation and cannot save the previous date silently',async()=>{
 const e=client(),native=e.element('gradeStart','2026-10-09'),display=e.element('date-display-gradeStart','2026/02/30');
 native.type='date';display.dataset.dateInput='gradeStart';display.focus=()=>{};
 e.c.syncDateInput(display,true);const before=plain(e.g.profile);
 await assert.rejects(e.c.gradeAction('next',{}),/日付|日時/);
 assert.equal(e.g.step,1);assert.deepEqual(plain(e.g.profile),before);assert.equal(e.calls.length,0);
});
