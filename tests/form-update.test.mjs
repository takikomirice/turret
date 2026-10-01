import test from 'node:test';
import assert from 'node:assert/strict';
import {adminEnvironment, AdminSheet, plain} from './helpers/admin-environment.mjs';

const question=(id,title)=>({itemId:id,title,questionItem:{question:{questionId:'q'+id,textQuestion:{}}}});
const doc=items=>({revisionId:'rev1',info:{title:'クラス別',description:''},settings:{emailCollectionType:'VERIFIED'},items});
const record=()=>({input:{templateId:'template',emailHeader:'メール',nameHeader:'名前',statusHeader:'状態',additionalColumns:[{header:'コメント',choices:[]}]}});
const headers=['日時','メール','A','B','C','D','名前','コメント','状態'];
test('updated forms keep management JSON and visible status at the right, after archived answers',()=>{
 const {c}=adminEnvironment(),h=['日時','メール','A','B','名前','コメント','管理','状態','手入力'];
 const p=c.buildManagedFormUpdate_(record(),doc([question('A','A')]),doc([question('A','A'),question('B','B')]),h);
 const json='{"v":1,"save":{"state":"pending"}}',s=sheet([h,['t','m','a','b','n','comment',json,'済','keep']]);
 c.arrangeManagedFormColumns_(s,p);c.arrangeManagedFormColumns_(s,p);
 assert.deepEqual(s.rows[0],['日時','メール','A','名前','コメント','手入力','B','管理','状態']);
 assert.equal(s.rows[1].at(-2),json);assert.equal(s.rows[1].at(-1),'済');assert.equal(s.rows[1][6],'b');
});
function sheet(rows){
 const s=new AdminSheet('回答',rows),get=s.getRange.bind(s);
 s.getRange=(r,c,h,w)=>Object.assign(get(r,c,h,w),{getColumn:()=>c,getNumColumns:()=>w||1});
 s.moveColumns=(range,to)=>{const from=range.getColumn()-1;for(const row of s.rows){const [v]=row.splice(from,1);row.splice(to-1-(from<to-1?1:0),0,v??'');}};
 s.deleteColumns=(start,count)=>{s.beforeDelete?.(start);for(const row of s.rows)row.splice(start-1,count);};
 return s;
}

test('template insertion uses stable question IDs and moves E between B and C without replacing answers',()=>{
 const {c}=adminEnvironment();assert.equal(typeof c.buildManagedFormUpdate_,'function');
 const before=doc('ABCD'.split('').map(x=>question(x,x)));
 const after=doc('ABECD'.split('').map(x=>question(x,x)));
 const p=c.buildManagedFormUpdate_(record(),after,before,headers);
 assert.deepEqual(plain(p.items.map(i=>i.title)),['A','B','E','C','D']);
 assert.equal(p.items[3].questionItem.question.questionId,'qC');
 assert.equal(p.requests.filter(r=>r.createItem).length,1);
 assert.equal(p.requests.filter(r=>r.deleteItem).length,0);
 const s=sheet([headers.concat('E'),['t','m','a','b','c','d','n','comment','済','e']]);
 c.arrangeManagedFormColumns_(s,p);
 assert.deepEqual(s.rows[0],['日時','メール','A','B','E','C','D','名前','コメント','状態']);
 assert.deepEqual(s.rows[1],['t','m','a','b','e','c','d','n','comment','済']);
});

for(const value of ['回答',0,false,' ', '=IF(TRUE,"","")'])test('removed question with data is kept at far right: '+String(value),()=>{
 const {c}=adminEnvironment();assert.equal(typeof c.buildManagedFormUpdate_,'function');
 const p=c.buildManagedFormUpdate_(record(),doc('ABD'.split('').map(x=>question(x,x))),doc('ABCD'.split('').map(x=>question(x,x))),headers);
 const s=sheet([headers,['t','m','a','b',value,'d','n','comment','済']]);
 c.arrangeManagedFormColumns_(s,p);
 assert.deepEqual(s.rows[0],['日時','メール','A','B','D','名前','コメント','状態','C']);
 assert.equal(s.rows[1][8],value);
});

test('only empty removed question columns are deleted; empty unmanaged and management columns survive',()=>{
 const {c}=adminEnvironment();assert.equal(typeof c.buildManagedFormUpdate_,'function');
 const h=headers.concat('手入力');
 const p=c.buildManagedFormUpdate_(record(),doc('ABD'.split('').map(x=>question(x,x))),doc('ABCD'.split('').map(x=>question(x,x))),h);
 const s=sheet([h,['t','m','a','b','','d','','','','']]);
 c.arrangeManagedFormColumns_(s,p);c.arrangeManagedFormColumns_(s,p);
 assert.deepEqual(s.rows[0],['日時','メール','A','B','D','名前','コメント','状態','手入力']);
});

test('question renames preserve the configured response header and existing question ID',()=>{
 const {c}=adminEnvironment();assert.equal(typeof c.buildManagedFormUpdate_,'function');
 const before=doc([question('A','A')]),after=doc([question('A','新しい質問文')]);
 const p=c.buildManagedFormUpdate_(record(),after,before,['日時','メール','A','名前','コメント','状態']);
 const s=sheet([['日時','メール','新しい質問文','名前','コメント','状態'],['t','m','回答','n','c','済']]);
 c.arrangeManagedFormColumns_(s,p);
 assert.equal(s.rows[0][2],'A');assert.equal(s.rows[1][2],'回答');
 assert.equal(p.items[0].questionItem.question.questionId,'qA');
});

test('copy mapping survives template rename and reorder when copy IDs differ',()=>{
 const {c}=adminEnvironment();assert.equal(typeof c.buildManagedFormUpdate_,'function');
 const r=record();r.formSync={itemMap:{A:'copyA',B:'copyB'},questionMap:{qA:'qcopyA',qB:'qcopyB'},columns:{qcopyA:'A',qcopyB:'B'},archivedHeaders:[]};
 const p=c.buildManagedFormUpdate_(r,doc([question('B','新B'),question('A','A')]),doc([question('copyA','A'),question('copyB','B')]),['日時','メール','A','B','名前','コメント','状態']);
 assert.deepEqual(plain(p.items.map(i=>i.itemId)),['copyB','copyA']);
 assert.equal(p.requests.filter(r=>r.deleteItem||r.createItem).length,0);
});

test('creation records source-to-copy identities before later template renames',()=>{
 const {c}=adminEnvironment();assert.equal(typeof c.captureManagedFormSync_,'function');
 const r=record();r.templateItems=[question('A','A'),question('B','B')];
 const sync=c.captureManagedFormSync_(r,[question('copyA','A'),question('copyB','B')],['日時','メール','A','B','名前','コメント','状態']);
 assert.equal(sync.itemMap.A,'copyA');assert.equal(sync.questionMap.qB,'qcopyB');assert.equal(sync.columns.qcopyA,'A');
});

test('duplicate headers, collisions, and changed question type stop before mutation',()=>{
 const {c}=adminEnvironment();assert.equal(typeof c.buildManagedFormUpdate_,'function');
 const before=doc([question('A','A')]);
 for(const after of [doc([question('A','A'),question('B','A')]),doc([question('A','名前')]),doc([{itemId:'A',title:'A',questionItem:{question:{questionId:'qA',choiceQuestion:{type:'RADIO',options:[{value:'x'}]}}}}])]){
  assert.throws(()=>c.buildManagedFormUpdate_(record(),after,before,['日時','メール','A','名前','コメント','状態']),/重複|種類|衝突/);
 }
});

test('missing new answer column prevents all column deletions and moves until Google linking catches up',()=>{
 const {c}=adminEnvironment();assert.equal(typeof c.buildManagedFormUpdate_,'function');
 const p=c.buildManagedFormUpdate_(record(),doc([question('E','E')]),doc([question('C','C')]),['日時','メール','C','名前','コメント','状態']);
 const s=sheet([['日時','メール','C','名前','コメント','状態'],['t','m','','n','','']]);
 assert.throws(()=>c.arrangeManagedFormColumns_(s,p),/回答列|反映待ち/);
 assert.deepEqual(s.rows[0],['日時','メール','C','名前','コメント','状態']);
});

test('empty display produced by a formula is retained even when getValues returns an empty string',()=>{
 const {c}=adminEnvironment(),p=c.buildManagedFormUpdate_(record(),doc([]),doc([question('C','C')]),['日時','メール','C','名前','コメント','状態']);
 const s=sheet([['日時','メール','C','名前','コメント','状態'],['t','m','','n','','']]),get=s.getRange;
 s.getRange=(r,col,h,w)=>{const range=get(r,col,h,w);if(r===2&&col===3)range.getFormulas=()=>[['=IF(TRUE,"","")']];return range;};
 c.arrangeManagedFormColumns_(s,p);assert.equal(s.rows[0].at(-1),'C');
});

test('grid row removal and branch destinations retain identities and track each response column',()=>{
 const {c}=adminEnvironment(),r=record();
 const grid={itemId:'grid',title:'表',questionGroupItem:{questions:[{questionId:'r1',rowQuestion:{title:'行1'}},{questionId:'r2',rowQuestion:{title:'行2'}}],grid:{columns:{type:'RADIO',options:[{value:'はい'}]}}}};
 const old=doc([grid]),updated=plain(grid);updated.questionGroupItem.questions.splice(0,1);
 const p=c.buildManagedFormUpdate_(r,doc([updated]),old,['日時','メール','表 [行1]','表 [行2]','名前','コメント','状態']);
 assert.deepEqual(plain(p.archivedHeaders),['表 [行1]']);assert.equal(p.items[0].questionGroupItem.questions[0].questionId,'r2');
 const branch={itemId:'branch',title:'分岐',questionItem:{question:{questionId:'bq',choiceQuestion:{type:'RADIO',options:[{value:'次',goToSectionId:'section'}]}}}};
 const next=c.buildManagedFormUpdate_(r,doc([branch,{itemId:'section',title:'次へ',pageBreakItem:{}}]),doc([]),['日時','メール','名前','コメント','状態']);
 assert.equal(next.items[0].questionItem.question.choiceQuestion.options[0].goToSectionId,next.items[1].itemId);
});

test('a removed column containing only a cell note is retained',()=>{
 const {c}=adminEnvironment(),p=c.buildManagedFormUpdate_(record(),doc([]),doc([question('C','C')]),['日時','メール','C','名前','コメント','状態']);
 const s=sheet([['日時','メール','C','名前','コメント','状態']]),get=s.getRange;
 s.getRange=(r,col,h,w)=>{const range=get(r,col,h,w);if(r===2&&col===3)range.getNotes=()=>[['記録を残す']];return range;};
 c.arrangeManagedFormColumns_(s,p);assert.equal(s.rows[0].at(-1),'C');
});
