import test from 'node:test';
import assert from 'node:assert/strict';
import {scoringEnvironment,AdminSheet} from './helpers/scoring-environment.mjs';
const plain=value=>JSON.parse(JSON.stringify(value));
const template=(rules=[])=>({name:'化学',enabled:true,domains:[[1,2,3],[1,2,3],[],[],[]],rules});
const rule=(a,b,outputs)=>({conditions:[a,b,'','',''],outputs,enabled:true,visible:true});

test('comparison cells preserve blanks and zero and accept inclusive aliases',()=>{
 const {c}=scoringEnvironment();
 assert.equal(c.scoringParseCondition_('').op,'any');
 assert.equal(c.scoringConditionMatches_(c.scoringParseCondition_('<=1'),null),false);
 assert.equal(c.scoringConditionMatches_(c.scoringParseCondition_('=<1'),0),true);
 assert.equal(c.scoringConditionMatches_(c.scoringParseCondition_('>1'),1),false);
 assert.equal(c.scoringConditionMatches_(c.scoringParseCondition_('3'),3),true);
 assert.throws(()=>c.scoringParseCondition_('>'),/数値|条件/);
 assert.throws(()=>c.scoringParseCondition_('score1+score2'),/条件/);
});

test('first match is per output; a blank output skips it and zero is emitted',()=>{
 const {c}=scoringEnvironment(),t=template([rule('>=2','',['','先行','']),rule('3','3',['0','後続','合計{合計}／{枠1}']),rule('','',['予備','',''])]);
 assert.deepEqual(plain(c.scoringPreviewTemplate({template:t,scores:[3,3,'','','']}).outputs),['0','先行','合計6／3']);
 assert.deepEqual(plain(c.scoringPreviewTemplate({template:t,scores:['',3,'','','']}).outputs),['予備','','']);
});

test('validator identifies per-output overlap, unreachable rules, impossible conditions and gaps',()=>{
 const {c}=scoringEnvironment(),t=template([rule('>=2','',['A','','']),rule('3','',['B','','']),rule('>3','',['C','','']),rule('>=2','',['','別出力',''])]);
 const result=c.scoringValidateTemplateTable({templates:[t]});
 assert.ok(result.issues.some(i=>i.code==='overlap'&&i.output===1&&i.example[0]===3));
 assert.ok(result.issues.some(i=>i.code==='unreachable'&&i.rule===2));
 assert.ok(result.issues.some(i=>i.code==='impossible'&&i.rule===3));
 assert.ok(result.issues.some(i=>i.code==='gap'&&i.output===1));
 assert.ok(!result.issues.some(i=>i.code==='overlap'&&i.output===2));
});

test('sheet and editor round-trip, reject stale revision, retain foreign sheets',()=>{
 const e=scoringEnvironment(),before=e.c.scoringGetTemplateTable();
 const t=template([rule('3','3',['5','よい',''])]);
 const saved=e.c.scoringSaveTemplateTable({revision:before.revision,templates:[t]});
 assert.equal(saved.templates[0].name,'化学');
 const sheet=e.sheets.get('採点テンプレ');assert.equal(sheet.rows[0][1],'枠1');assert.equal(sheet.rows[1][6],'5');
 sheet.rows[1][7]='直接編集';
 assert.equal(e.c.scoringGetTemplateTable().templates[0].rules[0].outputs[1],'直接編集');
 assert.throws(()=>e.c.scoringSaveTemplateTable({revision:saved.revision,templates:[t]}),/変更|再読込/);
 assert.equal(sheet.rows[1][7],'直接編集');
 e.c.Session.getActiveUser=()=>({getEmail:()=> 'intruder'});
 assert.throws(()=>e.c.scoringGetTemplateTable(),/本人/);
 assert.equal(e.isLocked(),false);
});

test('invalid sheet conditions stop grading instead of silently using partial rules',()=>{
 const e=scoringEnvironment(),t=template([rule('3','',['','','コメント'])]);
 const saved=e.c.scoringSaveTemplateTable({revision:e.c.scoringGetTemplateTable().revision,templates:[t]});
 const rules=e.c.scoringGetRules({templateId:'化学'});assert.equal(rules.ok,true);assert.equal(rules.rules[0].target,'出力3');
 e.sheets.get('採点テンプレ').rows[1][1]='oops';
 assert.throws(()=>e.c.scoringGetRules({templateId:'化学'}),/条件/);
});

test('validation reports incomplete analysis rather than claiming complete coverage',()=>{
 const {c}=scoringEnvironment(),t=template([rule('','','fallback'.split())]);
 t.rules[0].outputs=['fallback','',''];t.domains=Array.from({length:5},()=>Array.from({length:20},(_,i)=>i));
 const result=c.scoringValidateTemplateTable({templates:[t]});assert.equal(result.complete,false);assert.ok(result.issues.some(i=>i.code==='analysis-limit'));
});

test('disabled rules, separate outputs, fallback and collective shadowing are checked correctly',()=>{
 const {c}=scoringEnvironment(),t=template([rule('1','',['one','','']),rule('>=2','',['two','','']),rule('','',['fallback','','']),rule('','',['','other',''])]);
 const result=c.scoringValidateTemplateTable({templates:[t]});assert.ok(result.issues.some(i=>i.code==='unreachable'&&i.rule===3));assert.ok(!result.issues.some(i=>i.code==='gap'));
 t.rules[2].enabled=false;assert.ok(!c.scoringValidateTemplateTable({templates:[t]}).issues.some(i=>i.code==='overlap'));
});

test('saving shorter tables clears trailing rules and failures preserve old content and release locks',()=>{
 const e=scoringEnvironment(),t=template([rule('1','',['a','','']),rule('2','',['b','',''])]);
 let result=e.c.scoringSaveTemplateTable({revision:e.c.scoringGetTemplateTable().revision,templates:[t]});
 const sheet=e.sheets.get('採点テンプレ');sheet.beforeWrite=()=>{throw Error('write failed');};
 assert.throws(()=>e.c.scoringSaveTemplateTable({revision:result.revision,templates:[]}),/write failed/);assert.equal(sheet.rows[2][6],'b');assert.equal(e.isLocked(),false);
 delete sheet.beforeWrite;result=e.c.scoringSaveTemplateTable({revision:result.revision,templates:[template([rule('1','',['new','',''])])]});assert.equal(sheet.getLastRow(),2);assert.equal(result.templates[0].rules.length,1);
});

test('direct sheet metadata conflicts and malformed headers never overwrite the sheet',()=>{
 const e=scoringEnvironment(),t=template([rule('1','',['a','','']),rule('2','',['b','',''])]);
 e.c.scoringSaveTemplateTable({revision:e.c.scoringGetTemplateTable().revision,templates:[t]});const sheet=e.sheets.get('採点テンプレ');
 sheet.rows[2][11]='1,2';assert.throws(()=>e.c.scoringGetTemplateTable(),/異なります/);sheet.rows[2][11]='';sheet.rows[0][1]='score1';assert.throws(()=>e.c.scoringGetTemplateTable(),/ヘッダ/);
});

test('all templates are selectable even with legacy disabled flags, and save removes only the obsolete column',()=>{
 const e=scoringEnvironment(),t=template([rule('3','',['文言','',''])]);t.enabled=false;
 let saved=e.c.scoringSaveTemplateTable({revision:e.c.scoringGetTemplateTable().revision,templates:[t]});
 const sheet=e.sheets.get('採点テンプレ');assert.equal(sheet.rows[0].length,16);assert.equal(saved.templates[0].enabled,undefined);
 sheet.rows[0][16]='テンプレートを使用';sheet.rows[1][16]='0';
 const read=e.c.scoringGetTemplateTable();assert.equal(e.c.scoringTableTemplateList_().templates[0].enabled,1);
 assert.equal(e.c.scoringGetRules({templateId:'化学'}).rules[0].message,'文言');
 assert.equal(e.c.scoringPreviewTemplate({template:t,scores:[3]}).outputs[0],'文言');
 sheet.beforeWrite=()=>{throw Error('write failed');};assert.throws(()=>e.c.scoringSaveTemplateTable(read),/write failed/);assert.equal(sheet.rows[0][16],'テンプレートを使用');delete sheet.beforeWrite;
 saved=e.c.scoringSaveTemplateTable(read);assert.equal(sheet.rows[0][16],'');assert.equal(sheet.rows[1][16],'');assert.equal(saved.templates[0].rules[0].outputs[0],'文言');
 sheet.rows[0][16]='別の情報';assert.throws(()=>e.c.scoringGetTemplateTable(),/ヘッダ/);assert.equal(sheet.rows[0][16],'別の情報');
});

test('template flags never bypass validation, but rule use and visibility flags remain effective',()=>{
 const e=scoringEnvironment(),t=template([rule('bad','',['文言','',''])]);t.enabled=false;
 assert.equal(e.c.scoringValidateTemplateTable({templates:[t]}).ok,false);
 t.rules[0].enabled=false;t.rules.push(rule('3','',['有効','','']));t.rules[1].visible=false;
 e.c.scoringSaveTemplateTable({revision:e.c.scoringGetTemplateTable().revision,templates:[t]});
 const rules=e.c.scoringGetRules({templateId:t.name}).rules;assert.equal(rules.length,1);assert.equal(rules[0].visible,0);
});
