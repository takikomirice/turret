import test from 'node:test';
import assert from 'node:assert/strict';
import {adminEnvironment, AdminSheet, plain} from './helpers/admin-environment.mjs';

const headers=['元SS_ID','元スプシ名','元シート名','クラス名','courseId','メモ'];
function env(rows=[]){
 const e=adminEnvironment();e.ss.getName=()=> '運用';
 e.sheets.set('クラス一覧',new AdminSheet('クラス一覧',[['クラス名','コースID','同期対象(1)'],['授業A','100','1'],['授業B','200','1']]));
 e.sheet=new AdminSheet('対応表',[headers,...rows]);e.sheets.set('対応表',e.sheet);
 e.c.collectFormTargetSheets_=()=>[{sourceId:'test-admin-sheet',spreadsheetName:'運用',sheetName:'回答 A'}];
 e.record={sheetName:'回答 A',className:'授業A',courseId:'100',input:{prefix:'回答',emailHeader:'メール',nameHeader:'名前',statusHeader:'状態',additionalColumns:[]}};
 return e;
}

test('form registration installs class choices and a formula immediately',()=>{
 const e=env();e.c.registerManagedSource_(e.record);
 assert.deepEqual(e.sheet.validations[1][3].values,['授業A','授業B']);
 assert.equal(e.sheet.rows[1][4],'=IF(D2="","",IFS(D2="授業A","100",D2="授業B","200",TRUE,""))');
});

test('registration retry preserves a manually selected class, formula and memo',()=>{
 const e=env([['test-admin-sheet','運用','回答 A','授業B','200','先生のメモ']]);
 e.c.registerManagedSource_(e.record);
 const before=plain(e.sheet.rows);e.c.registerManagedSource_(e.record);
 assert.deepEqual(e.sheet.rows,before);assert.equal(e.sheet.rows[1][3],'授業B');assert.equal(e.sheet.rows[1][5],'先生のメモ');
 assert.match(e.sheet.rows[1][4],/^=IF/);
});

test('update upgrades fixed IDs and adds new targets without deleting disconnected rows or extra columns',()=>{
 const e=env([['old-source','旧運用','以前の回答','授業B','200','残す','独自情報']]);
 e.c.createMappingSheetUnlocked_();
 assert.equal(e.sheet.rows.length,3);assert.equal(e.sheet.rows[1][5],'残す');assert.equal(e.sheet.rows[1][6],'独自情報');
 assert.match(e.sheet.rows[1][4],/^=IF/);assert.equal(e.sheet.rows[2][2],'回答 A');
 assert.equal(e.sheet.rows[2][3],'');assert.match(e.sheet.rows[2][4],/^=IF\(D3/);
});

test('same-name classes are distinguishable and existing IDs select the right option',()=>{
 const e=env([['test-admin-sheet','運用','回答 A','授業','200','残す']]);
 e.sheets.get('クラス一覧').rows=[['クラス名','コースID','同期対象(1)'],['授業','100','1'],['授業','200','1']];
 e.c.createMappingSheetUnlocked_();
 assert.equal(e.sheet.rows[1][3],'授業（ID: 200）');
 assert.deepEqual(e.sheet.validations[1][3].values,['授業（ID: 100）','授業（ID: 200）']);
 assert.match(e.sheet.rows[1][4],/D2="授業（ID: 200）","200"/);
});

test('ambiguous legacy names without an ID stop before modifying the sheet',()=>{
 const e=env([['test-admin-sheet','運用','回答 A','授業','','残す']]);
 e.sheets.get('クラス一覧').rows=[['クラス名','コースID','同期対象(1)'],['授業','100','1'],['授業','200','1']];
 const before=plain(e.sheet.rows);assert.throws(()=>e.c.createMappingSheetUnlocked_(),/特定|同名/);assert.deepEqual(e.sheet.rows,before);
});

test('deselected existing assignments remain usable and new classes enter the choices',()=>{
 const e=env([['test-admin-sheet','運用','回答 A','前の授業','300','残す']]);e.c.createMappingSheetUnlocked_();
 assert.equal(e.sheet.rows[1][3],'前の授業');assert.match(e.sheet.rows[1][4],/"前の授業","300"/);
 assert.deepEqual(e.sheet.validations[1][3].values,['授業A','授業B','前の授業']);
});

test('a cleared assignment is not restored from an old fixed ID',()=>{
 const e=env([['test-admin-sheet','運用','回答 A','','100','残す']]);e.c.registerManagedSource_(e.record);
 assert.equal(e.sheet.rows[1][3],'');assert.match(e.sheet.rows[1][4],/^=IF\(D2="",""/);
});

test('renaming a class cannot move its existing assignment to a different class reusing the old name',()=>{
 const e=env([['test-admin-sheet','運用','回答 A','授業A','100','残す']]);
 e.sheets.get('クラス一覧').rows=[['クラス名','コースID','同期対象(1)'],['改名した授業','100','1'],['授業A','200','1']];
 e.c.createMappingSheetUnlocked_();assert.equal(e.sheet.rows[1][3],'改名した授業');
 assert.equal(e.sheet.getRange(2,5).getDisplayValues()[0][0],'100');
});

test('clearing and changing the class uses the generated formula without pressing update',()=>{
 const e=env();e.c.registerManagedSource_(e.record);
 e.sheet.getRange(2,4).setValue('授業B');assert.equal(e.sheet.getRange(2,5).getDisplayValues()[0][0],'200');
 e.sheet.getRange(2,4).setValue('');assert.equal(e.sheet.getRange(2,5).getDisplayValues()[0][0],'');
});

test('duplicate source assignments and malformed headers cannot silently overwrite rows',()=>{
 const row=['test-admin-sheet','運用','回答 A','授業A','100','残す'];
 for(const corrupt of ['duplicate','headers']){
  const e=env(corrupt==='duplicate'?[row,row]:[row]);if(corrupt==='headers')e.sheet.rows[0][4]='別の列';
  const before=plain(e.sheet.rows);assert.throws(()=>e.c.createMappingSheetUnlocked_(),/重複|見出し/);assert.deepEqual(e.sheet.rows,before);
 }
});

test('failed validation update restores values, formulas and existing choices; retry succeeds',()=>{
 const e=env([['old-source','旧運用','旧回答','授業A','100','残す']]);
 e.sheet.getRange(2,4).setDataValidation({values:['旧候補'],allowInvalid:false});
 const before=plain(e.sheet.rows),rules=plain(e.sheet.getRange(1,1,2,6).getDataValidations()),get=e.sheet.getRange.bind(e.sheet);let failed=false;
 e.sheet.getRange=(...args)=>{const range=get(...args),set=range.setDataValidation;range.setDataValidation=rule=>{if(!failed){failed=true;throw Error('validation failed');}return set(rule);};return range;};
 assert.throws(()=>e.c.createMappingSheetUnlocked_(),/validation failed/);
 assert.deepEqual(e.sheet.rows.slice(0,e.sheet.getLastRow()),before);assert.deepEqual(plain(e.sheet.getRange(1,1,2,6).getDataValidations()),rules);
 e.c.createMappingSheetUnlocked_();assert.match(e.sheet.rows[1][4],/^=IF/);
});

test('a failure after writing formulas restores old formulas and rules and removes only appended content',()=>{
 const e=env();e.c.registerManagedSource_(e.record);
 e.sheet.getRange(2,6).setValue('手修正');
 const before=plain(e.sheet.rows),rules=plain(e.sheet.getRange(2,4,1,2).getDataValidations());
 e.sheets.get('クラス一覧').rows.push(['追加クラス','300','1']);
 e.c.collectFormTargetSheets_=()=>[{sourceId:'test-admin-sheet',spreadsheetName:'運用',sheetName:'回答 追加'}];
 let failed=false;e.c.SpreadsheetApp.flush=()=>{if(!failed){failed=true;throw Error('flush failed');}};
 assert.throws(()=>e.c.createMappingSheetUnlocked_(),/flush failed/);
 assert.deepEqual(e.sheet.rows.slice(0,e.sheet.getLastRow()),before);
 assert.deepEqual(plain(e.sheet.getRange(2,4,1,2).getDataValidations()),rules);
 e.c.createMappingSheetUnlocked_();assert.equal(e.sheet.rows[1][5],'手修正');assert.equal(e.sheet.rows[2][2],'回答 追加');
});

test('quoted class names are escaped in formulas and can be selected again on retry',()=>{
 const e=env();e.sheets.get('クラス一覧').rows[1][0]='授業"A';e.record.className='授業"A';
 e.c.registerManagedSource_(e.record);e.c.registerManagedSource_(e.record);
 assert.equal(e.sheet.rows[1][4],'=IF(D2="","",IFS(D2="授業""A","100",D2="授業B","200",TRUE,""))');
 assert.equal(e.sheet.getRange(2,5).getDisplayValues()[0][0],'100');
});
