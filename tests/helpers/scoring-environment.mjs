import {adminEnvironment, AdminSheet, validConfig, sourceId} from './admin-environment.mjs';
export {AdminSheet, sourceId};
export function scoringEnvironment(options={}) {
 const e=adminEnvironment(options);
 e.c.Session.getActiveUser=e.c.Session.getEffectiveUser=()=>({getEmail:()=> 'owner@example.com'});
 e.c.ScriptApp.getScriptId=()=> 'scoring-test';
 const books=new Map();
 e.ss.getName=()=> '運用';books.set(e.ss.getId(),e.ss);
 const addBook=(id,name='回答ブック')=>{const sheets=new Map();const book={getId:()=>id,getName:()=>name,getSheets:()=>[...sheets.values()],getSheetByName:n=>sheets.get(n)||null,getSpreadsheetTimeZone:()=> 'Asia/Tokyo'};books.set(id,book);return {book,sheets};};
 const source=addBook(sourceId);
 e.c.SpreadsheetApp.openById=id=>{if(!books.has(id))throw Error('unknown workbook');return books.get(id);};
 const answer=new AdminSheet('回答 1',[['メール','名前','状態','点数','講評','授業の日付を入力してください','回答'],['student@example.com','生徒','','','','2026-09-29','回答内容']]);
 source.sheets.set(answer.name,answer);
 const target={spreadsheetId:sourceId,sheetId:answer.sheetId};
 const configure=patch=>e.props.set('APP_CONFIG',JSON.stringify({...validConfig(),...patch}));
 return {...e,books,source,answer,target,addBook,configure};
}
