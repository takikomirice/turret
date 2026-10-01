import test from 'node:test';
import assert from 'node:assert/strict';
import {scoringEnvironment,AdminSheet,sourceId} from './helpers/scoring-environment.mjs';

test('internal exclusions preserve the different external-workbook policies of scoring and import',()=>{
 const e=scoringEnvironment(),local=e.ss.getId();
 for(const name of ['クラス一覧','生徒一覧','設定シート','Config','設定(編集不可)','評価データ','送信シート','エラー','対応表','フォーム管理']) {
  assert.equal(e.c.scoringIsInternalSheet_(name),true,name);
  assert.equal(e.c.isManagedInternalSheet_(name,local),true,name);
  assert.equal(e.c.isManagedInternalSheet_(name,sourceId),false,name);
 }
 for(const name of ['_templates','_rules','_scoring_decisions','_scoring_saves']) {
  assert.equal(e.c.scoringIsInternalSheet_(name),true);
  assert.equal(e.c.isManagedInternalSheet_(name,sourceId),true);
 }
 assert.equal(e.c.scoringIsInternalSheet_('回答 1'),false);
 assert.equal(e.c.isManagedInternalSheet_('回答 1',local),false);
});

test('scoring candidates exclude system/rule tabs and distinguish same names by workbook and sheet ID',()=>{
 const e=scoringEnvironment();const otherId='other_12345678901234567890';const other=e.addBook(otherId);other.sheets.set('回答 1',new AdminSheet('回答 1'));
 for(const name of ['生徒一覧','評価データ','送信シート','フォーム管理','_rules','_templates'])e.source.sheets.set(name,new AdminSheet(name));
 e.configure({formSources:[{id:sourceId},{id:otherId},{id:sourceId}],formSheetNamePrefix:['回答','生徒','評価','送信','フォーム','_']});
 const targets=e.c.scoringGetTargets();assert.equal(targets.length,2);assert.notEqual(targets[0].key,targets[1].key);
 assert.equal(e.c.resolveScoringTarget_(targets[0]).sheet,e.answer);
 assert.equal(e.c.resolveScoringTarget_(targets[1]).sheet,other.sheets.get('回答 1'));
});

test('recorded answer tab survives a prefix mismatch but reports that return is disabled',()=>{
 const e=scoringEnvironment();const managed=new AdminSheet('クラスA振り返り');e.sheets.set(managed.name,managed);
 e.c.saveManagedRecord_({kind:'form',id:'form-a',responseSheetId:managed.sheetId,courseId:'class-a'});
 const targets=e.c.scoringGetTargets();const found=targets.find(t=>t.sheetId===managed.sheetId);
 assert.equal(found.spreadsheetId,e.ss.getId());assert.equal(found.returnEnabled,false);assert.match(found.warning,/接頭辞|返却/);
});

test('deleted, unregistered, missing and internal target selectors are rejected without fallback',()=>{
 const e=scoringEnvironment();
 for(const target of [null,{},'回答 1',{spreadsheetId:'other',sheetId:e.answer.sheetId}])assert.throws(()=>e.c.resolveScoringTarget_(target),/対象|登録/);
 e.source.sheets.delete(e.answer.name);assert.throws(()=>e.c.resolveScoringTarget_(e.target),/対象|登録/);
 assert.deepEqual(Array.from(e.c.scoringGetTargets()),[]);
 e.c.Session.getActiveUser=()=>({getEmail:()=> 'other@example.com'});assert.throws(()=>e.c.scoringGetTargets(),/本人/);
});

test('an inaccessible registered workbook is an error, not an empty candidate list',()=>{
 const e=scoringEnvironment();e.c.SpreadsheetApp.openById=()=>{throw Error('permission denied');};
 assert.throws(()=>e.c.scoringGetTargets(),/permission denied/);
});
