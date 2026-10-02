import test from 'node:test';
import assert from 'node:assert/strict';
import {scoringEnvironment,AdminSheet} from './helpers/scoring-environment.mjs';

function environment(patch={}){
 const e=scoringEnvironment(),sheet=new AdminSheet('フォーム回答 1A',[
  ['メール','名前','点数','講評'],
  ['KID@EXAMPLE.COM','','5','採点済み'],
  ['missing@example.com','','3','未照合'],
  ['kid@example.com','手入力','4','保持'],
  ['kid@example.com','=IF(TRUE,"","")','2','数式'],
  ['conflict@example.com','','1','氏名競合']
 ]);
 e.sheets.set(sheet.name,sheet);
 e.sheets.set('生徒一覧',new AdminSheet('生徒一覧',[
  ['メールアドレス','名前','コースID'],
  ['kid@example.com','正しい名前','100'],['kid@example.com','別クラス名','200'],
  ['conflict@example.com','候補1','100'],['conflict@example.com','候補2','100']
 ]));
 e.c.saveManagedRecord_({id:'managed-a',kind:'form',courseId:'100',responseSheetId:sheet.sheetId,input:{emailHeader:'メール',nameHeader:'名前'},...patch});
 return {...e,sheet,managedTarget:{spreadsheetId:e.ss.getId(),sheetId:sheet.sheetId}};
}

test('scoring name enrichment fills only blank names in the selected managed answer tab',()=>{
 const e=environment(),before=structuredClone(e.answer.rows);
 const targets=e.c.scoringGetTargets();
 assert.equal(targets.find(t=>t.sheetId===e.sheet.sheetId).namesAvailable,true);
 assert.equal(targets.find(t=>t.sheetId===e.answer.sheetId).namesAvailable,false);
 assert.deepEqual({...e.c.scoringRefreshNames(JSON.stringify([e.ss.getId(),e.sheet.sheetId]))},{updated:1,unmatched:2,skipped:2});
 assert.deepEqual(e.sheet.rows.slice(1),[
  ['KID@EXAMPLE.COM','正しい名前','5','採点済み'],['missing@example.com','','3','未照合'],
  ['kid@example.com','手入力','4','保持'],['kid@example.com','=IF(TRUE,"","")','2','数式'],
  ['conflict@example.com','','1','氏名競合']
 ]);
 assert.deepEqual(e.answer.rows,before);
});

test('name enrichment rejects external, invalid and non-owner targets before writing',()=>{
 const e=environment(),before=structuredClone(e.sheet.rows);
 assert.throws(()=>e.c.scoringRefreshNames(e.target),/管理.*フォーム/);
 assert.throws(()=>e.c.scoringRefreshNames({spreadsheetId:e.ss.getId(),sheetId:-1}),/対象|登録/);
 e.c.Session.getActiveUser=()=>({getEmail:()=> 'other@example.com'});
 assert.throws(()=>e.c.scoringRefreshNames(e.managedTarget),/本人/);
 assert.deepEqual(e.sheet.rows,before);
});

for(const field of ['formUpdate','settingsUpdate','columnSetup'])test('name enrichment rejects an interrupted '+field+' without changing answers',()=>{
 const e=environment({[field]:{}}),before=structuredClone(e.sheet.rows);
 assert.equal(e.c.scoringGetTargets().find(t=>t.sheetId===e.sheet.sheetId).namesAvailable,false);
 assert.throws(()=>e.c.scoringRefreshNames(e.managedTarget),/更新|設定/);
 assert.deepEqual(e.sheet.rows,before);
});

test('name enrichment reports missing roster or name headers without changing grades',()=>{
 for(const missing of ['roster','name']){
  const e=environment();if(missing==='roster')e.sheets.delete('生徒一覧');else e.sheet.rows[0][1]='別の列';
  const before=structuredClone(e.sheet.rows);
  assert.throws(()=>e.c.scoringRefreshNames(e.managedTarget),/生徒一覧|名前補完/);
  assert.deepEqual(e.sheet.rows,before);
 }
});
