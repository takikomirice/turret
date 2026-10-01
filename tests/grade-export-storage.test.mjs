import {test} from 'node:test';
import assert from 'node:assert/strict';
import {scoringEnvironment,AdminSheet} from './helpers/scoring-environment.mjs';
import {plain} from './helpers/admin-environment.mjs';

const kinds=['grade-all','grade-layout','grade-rules'];
const headers=['種別','ID','内容(JSON)'];
function setup() {
 const e=scoringEnvironment(),p=plain(e.c.gradeDefault_());
 p.base.valueHeader='点数';p.base.classIds=['course-a'];p.base.mapping={'course-a':'授業A'};
 p.calendar=[{id:'lesson-1',title:'授業A',date:'2026-09-24',time:'09:00'}];
 p.rules.dates=[{courseId:'course-a',from:'2026-09-25',to:'2026-09-24',mode:'planned'}];
 return {e,p};
}
function save(e,p,action='draft',name='年間設定') {
 return e.c.gradeSave({revision:e.c.gradeReadStore_().revision,action,profile:p,name});
}
function apply(e,kind,json,choices) {
 const review=e.c.inspectConfigurationFile('auto',json);
 assert.equal(review.kind,kind);
 return e.c.applyConfigurationFile(kind,json,review.revision,choices);
}
function state(e) {return plain(e.c.gradeReadStore_());}
function addSheet(e,name,rows) {
 const sheet=new AdminSheet(name,rows);
 sheet.setName=newName=>{if(e.sheets.has(newName))throw Error('name exists');e.sheets.delete(sheet.name);sheet.name=newName;e.sheets.set(newName,sheet);return sheet;};
 e.sheets.set(name,sheet);return sheet;
}

for(const kind of kinds)test(kind+' export auto-detects and round trips without including answer data',()=>{
 const {e,p}=setup();save(e,p);const file=e.c.exportConfigurationFile(kind),payload=JSON.parse(file.json),target=scoringEnvironment();
 const answers=plain(target.answer.rows);assert.match(file.filename,new RegExp('^turret-'+kind+'-'));
 assert.equal(payload.application,'turret');assert.equal(payload.type,kind);assert.equal(payload.schemaVersion,1);
 assert.equal(file.json.includes('student@example.com'),false);assert.equal(file.json.includes('回答内容'),false);
 assert.equal('base' in payload,kind!=='grade-rules');assert.equal('rules' in payload,kind!=='grade-layout');assert.equal('calendar' in payload,kind!=='grade-layout');
 const beforeSheets=target.sheets.size;const review=target.c.inspectConfigurationFile('auto',file.json);assert.equal(target.sheets.size,beforeSheets);
 target.c.applyConfigurationFile(review.kind,file.json,review.revision);
 const value=state(target).value;
 if(kind!=='grade-rules')assert.deepEqual(value.draft.base,p.base);
 if(kind!=='grade-layout'){assert.deepEqual(value.draft.rules,p.rules);assert.deepEqual(value.draft.calendar,p.calendar);}
 assert.equal(value.presets.length,kind==='grade-all'?1:0);assert.deepEqual(target.answer.rows,answers);assert.equal(target.triggers.length,0);
 assert.equal(target.isLocked(),false);
});

test('layout import preserves existing correction rules, dates, and named presets',()=>{
 const {e,p}=setup();save(e,p,'preset');const source=scoringEnvironment(),incoming=plain(source.c.gradeDefault_());incoming.base.sort='number';incoming.base.leading=['roster:name'];save(source,incoming);
 apply(e,'grade-layout',source.c.exportConfigurationFile('grade-layout').json);
 const value=state(e).value;assert.deepEqual(value.draft.base,incoming.base);assert.deepEqual(value.draft.rules,p.rules);assert.deepEqual(value.draft.calendar,p.calendar);assert.deepEqual(value.presets[0].profile,p);
});

for(const choice of ['current','incoming'])test('rule merge requires explicit conflicts and honors '+choice,()=>{
 const {e,p}=setup();save(e,p);const source=scoringEnvironment(),incoming=plain(p);incoming.rules.dates[0].to='2026-09-23';incoming.calendar.push({id:'lesson-2',title:'授業A',date:'2026-09-28',time:'09:00'});save(source,incoming);
 const json=source.c.exportConfigurationFile('grade-rules').json,review=e.c.inspectConfigurationFile('auto',json),before=state(e);
 assert.equal(review.conflicts.length,1);assert.throws(()=>e.c.applyConfigurationFile('grade-rules',json,review.revision),/競合/);assert.deepEqual(state(e),before);
 e.c.applyConfigurationFile('grade-rules',json,review.revision,{[review.conflicts[0].id]:choice});
 const draft=state(e).value.draft;assert.equal(draft.rules.dates[0].to,choice==='incoming'?'2026-09-23':'2026-09-24');assert.equal(draft.calendar.length,2);assert.deepEqual(draft.base,p.base);
});

test('stale revision or changed file cannot overwrite a newer draft',()=>{
 const {e,p}=setup();save(e,p);const json=e.c.exportConfigurationFile('grade-layout').json,review=e.c.inspectConfigurationFile('auto',json);
 const newer=plain(p);newer.base.sort='number';save(e,newer);const before=state(e);
 assert.throws(()=>e.c.applyConfigurationFile('grade-layout',json,review.revision),/変更/);assert.deepEqual(state(e),before);
 const valid=e.c.inspectConfigurationFile('auto',json),changed=JSON.parse(json);changed.base.sort='roster';
 assert.throws(()=>e.c.applyConfigurationFile('grade-layout',JSON.stringify(changed),valid.revision),/変更/);assert.deepEqual(state(e),before);
});

test('malformed grade files and schema mismatches do not create storage or alter settings',()=>{
 const {e,p}=setup();save(e,p);const payload=JSON.parse(e.c.exportConfigurationFile('grade-all').json),target=scoringEnvironment(),before=plain([...target.props]);
 const variants=[{...payload,schemaVersion:2},{...payload,application:'other'},{...payload,type:'grade-layout'},{...payload,answers:[['private']]},{...payload,base:{...payload.base,sort:'invalid'}},{...payload,calendar:[{id:'bad',title:'授業A',date:'2026-02-30',time:'09:00'}]}];
 for(const bad of variants)assert.throws(()=>target.c.inspectConfigurationFile('auto',JSON.stringify(bad)));
 assert.throws(()=>target.c.inspectConfigurationFile('auto','broken'),/JSON/);
 assert.throws(()=>target.c.inspectConfigurationFile('auto',' '.repeat(1048577)),/1MiB/);
 assert.deepEqual(plain([...target.props]),before);assert.equal(target.sheets.size,0);assert.equal(target.isLocked(),false);
});

test('failed import write leaves committed draft and presets readable',()=>{
 const {e,p}=setup();save(e,p,'preset');const incoming=plain(p);incoming.base.sort='number';const source=scoringEnvironment();save(source,incoming);const json=source.c.exportConfigurationFile('grade-all').json,review=e.c.inspectConfigurationFile('auto',json),before=state(e);
 let writes=0;e.sheets.get('システム管理').beforeWrite=()=>{if(++writes===2)throw Error('head denied');};
 assert.throws(()=>e.c.applyConfigurationFile('grade-all',json,review.revision),/head denied/);assert.deepEqual(state(e),before);assert.equal(e.isLocked(),false);
});

test('multi-part failure keeps the committed bank and a later retry can replace it',()=>{
 const {e,p}=setup();save(e,p,'preset');const before=state(e),large=plain(p);
 large.rules.answers=Array.from({length:500},(_,i)=>({id:'answer-'+i,fingerprint:'fingerprint-'+i,to:'2026-09-24',mode:'planned'}));
 let writes=0;const sheet=e.sheets.get('システム管理');sheet.beforeWrite=()=>{if(++writes===2)throw Error('second chunk denied');};
 assert.throws(()=>save(e,large),/second chunk denied/);assert.deepEqual(state(e),before);
 sheet.beforeWrite=null;save(e,large);assert.deepEqual(state(e).value.draft,large);assert.equal(state(e).value.presets.length,1);assert.deepEqual(state(e).value.presets[0].profile,p);
});

test('legacy managed sheet read is nonmutating and locked save renames it without losing history',()=>{
 const {e,p}=setup(),record={schema:1,id:'existing',kind:'form',ownerSpreadsheetId:e.ss.getId(),ownerScriptId:e.c.ScriptApp.getScriptId(),revision:'old'};
 const rows=[headers,['form','existing',JSON.stringify(record)]],sheet=addSheet(e,'フォーム管理',rows),id=sheet.sheetId;
 assert.deepEqual(plain(e.c.getManagedRecords_()),[record]);assert.equal(e.sheets.has('システム管理'),false);
 save(e,p);assert.equal(e.sheets.has('フォーム管理'),false);assert.equal(e.sheets.get('システム管理'),sheet);assert.equal(sheet.sheetId,id);assert.deepEqual(sheet.rows.slice(0,2),rows);assert.deepEqual(state(e).value.draft,p);
});

for(const currentRows of [[headers],[headers,['manual','x','{}']]])test('coexisting old and new managed sheets block reads and saves without editing either',()=>{
 const {e,p}=setup(),oldRows=[headers,['form','old','{"schema":1,"id":"old","kind":"form"}']];addSheet(e,'フォーム管理',oldRows);addSheet(e,'システム管理',currentRows);
 const before=plain([...e.sheets].map(([name,s])=>[name,s.rows]));assert.throws(()=>e.c.gradeReadStore_(),/両方/);assert.throws(()=>e.c.gradeSave({revision:'',action:'draft',profile:p}),/両方/);assert.deepEqual(plain([...e.sheets].map(([name,s])=>[name,s.rows])),before);
});

test('malformed legacy managed sheet is never renamed or replaced',()=>{
 const {e,p}=setup(),rows=[['ユーザー列'],['残す']];const old=addSheet(e,'フォーム管理',rows);
 assert.throws(()=>save(e,p),/構成/);assert.equal(e.sheets.get('フォーム管理'),old);assert.equal(e.sheets.has('システム管理'),false);assert.deepEqual(old.rows,rows);
});

test('rescheduled calendar event with the same UID requires a choice instead of adding a second lesson',()=>{
 const {e,p}=setup();save(e,p);const incoming=plain(p);incoming.calendar[0].date='2026-09-28';const source=scoringEnvironment();save(source,incoming);
 const json=source.c.exportConfigurationFile('grade-rules').json,review=e.c.inspectConfigurationFile('auto',json);
 assert.equal(review.conflicts.length,1);
 e.c.applyConfigurationFile('grade-rules',json,review.revision,{[review.conflicts[0].id]:'incoming'});
 const calendar=state(e).value.draft.calendar;assert.equal(calendar.length,1);assert.equal(calendar[0].date,'2026-09-28');
});
