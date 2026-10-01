import { readFileSync, existsSync } from 'node:fs';
import vm from 'node:vm';
import { createHash } from 'node:crypto';

export const plain = value => JSON.parse(JSON.stringify(value));
export const sourceId = 'source_12345678901234567890';
let nextSheetId = 1;
export const validConfig = () => ({
  reminderTo: ['teacher@example.invalid'], formSheetNamePrefix: ['回答'],
  emailHeader: 'メール', studentNameHeader: '名前', formStatusHeader: '状態', scoreSourceHeader: '点数', replyBodyHeader: '返信本文',
  formSources: [{ id: sourceId, url: `https://docs.google.com/spreadsheets/d/${sourceId}/edit`, courseId: '' }],
  fields: [{ key: 'score', sourceHeader: '点数', evalHeader: '点数', sendHeader: '点数', type: 'score_grade' }],
  messageTemplate: '評価：＜score＞'
});

export class AdminSheet {
  constructor(name, rows = []) { this.name = name; this.rows = rows.map(row => [...row]); this.notes=[]; this.maxColumns = 26; this.sheetId = nextSheetId++; }
  getSheetId() { return this.sheetId; }
  getName() { return this.name; }
  getLastRow() { let n = this.rows.length; while(n && this.rows[n - 1].every(v => v === '' || v == null)) n--; return n; }
  getLastColumn() { return Math.max(0, ...this.rows.map(row => row.length)); }
  getMaxColumns() { return this.maxColumns; }
  getMaxRows() { return 1000; }
  getDataRange() { return this.getRange(1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); }
  getRange(row, col, height = 1, width = 1) {
    const read = () => Array.from({length:height}, (_,y) => Array.from({length:width},(_,x)=>this.rows[row+y-1]?.[col+x-1] ?? ''));
    const range = {
      getValues: read, getDisplayValues: () => read().map(r=>r.map(String)), getValue: () => read()[0][0],
      getFormulas: () => read().map(r=>r.map(v=>typeof v==='string' && v.startsWith('=') ? v : '')),
      getColumn:()=>col, getNumColumns:()=>width,
      getNote:()=>this.notes[row-1]?.[col-1] || '',
      setNote:value=>{this.notes[row-1]??=[];this.notes[row-1][col-1]=value;return range;},
      getNotes: () => Array.from({length:height},(_,y)=>Array.from({length:width},(_,x)=>this.notes[row+y-1]?.[col+x-1] || '')), setNotes: () => range,
      setValue: value => range.setValues([[value]]),
      setValues: values => { this.beforeWrite?.(values); values.forEach((r,y)=>r.forEach((value,x)=>{this.rows[row+y-1] ??=[]; this.rows[row+y-1][col+x-1]=value;})); return range; },
      clearContent: () => range.setValues(Array.from({length:height},()=>Array(width).fill(''))),
      setNumberFormat: () => range, setFontWeight: () => range, setBackground: () => range, setWrap: () => range,
      setFormulas: values => range.setValues(values), clearDataValidations: () => range
    }; return range;
  }
  appendRow(row) { this.rows.push([...row]); return this; }
  clear() { this.rows=[]; return this; }
  clearContents() { return this.clear(); }
  deleteRow(row) { this.rows.splice(row-1,1); return this; }
  deleteRows(row,count) { this.rows.splice(row-1,count); return this; }
  deleteColumns(start,count) { this.maxColumns-=count; return this; }
  insertColumnsAfter(start,count) { this.maxColumns+=count; return this; }
  moveColumns(range,to) {const from=range.getColumn()-1,width=this.getLastColumn();for(const rows of [this.rows,this.notes])for(const row of rows){while(row.length<width)row.push('');const [value]=row.splice(from,1);row.splice(to-1-(from<to-1?1:0),0,value);}return this;}
  setFrozenRows() { return this; }
  setColumnWidths() { return this; }
  insertRowsAfter(start,count) { return this; }
  autoResizeColumns() { return this; }
  hideSheet() { this.hidden=true; return this; }
  showSheet() { this.hidden=false; return this; }
  isSheetHidden() { return !!this.hidden; }
}

export function adminEnvironment({config = validConfig(), includeAdmin = true, actualAutomation = false} = {}) {
  const props = new Map(), userProps = new Map(), sheets = new Map();
  if (config) props.set('APP_CONFIG', JSON.stringify(config));
  const propertyApi = map => ({ getProperty:key=>map.get(key)??null, setProperty:(key,value)=>{map.set(key,String(value));}, deleteProperty:key=>map.delete(key), getProperties:()=>Object.fromEntries(map) });
  let locked = false, activeSheet = null, uuid = 0;
  const triggers = [];
  const ss = { getSheetByName:name=>sheets.get(name)||null, getSheets:()=>[...sheets.values()],
    insertSheet:name=>{const sheet=new AdminSheet(name);sheets.set(name,sheet);return sheet;},
    deleteSheet:sheet=>sheets.delete(sheet.name), setActiveSheet:sheet=>{activeSheet=sheet.name;}, getId:()=> 'test-admin-sheet', getSpreadsheetTimeZone:()=> 'Asia/Tokyo' };
  const c=vm.createContext({console,Logger:{log(){}},Date,PropertiesService:{getScriptProperties:()=>propertyApi(props),getUserProperties:()=>propertyApi(userProps)},
    SpreadsheetApp:{getActiveSpreadsheet:()=>ss,flush(){},getUi:()=>({alert(){throw Error('unexpected native alert');}})},
    LockService:{getScriptLock:()=>({tryLock:()=>{if(locked)return false;locked=true;return true;},releaseLock:()=>{locked=false;}})},
    Session:{getScriptTimeZone:()=> 'Asia/Tokyo'},
    Utilities:{DigestAlgorithm:{SHA_256:'sha256'},Charset:{UTF_8:'utf8'},computeDigest:(algorithm,value)=>Array.from(createHash('sha256').update(value).digest()),getUuid:()=> 'test-uuid-'+(++uuid),sleep(){} },
    ScriptApp:{AuthMode:{FULL:'FULL'},requireScopes(){},EventType:{CLOCK:'CLOCK'},getProjectTriggers:()=>triggers,deleteTrigger:trigger=>{const index=triggers.indexOf(trigger);if(index>=0)triggers.splice(index,1);},newTrigger(){throw Error('unexpected real trigger creation boundary');}}
  });
  vm.runInContext(readFileSync('Code.gs','utf8'),c,{filename:'Code.gs'});
  if(includeAdmin && existsSync('Administration.gs')) vm.runInContext(readFileSync('Administration.gs','utf8'),c,{filename:'Administration.gs'});
  const schedule={importHour:5,deliveryHour:8,reminderHour:16,reminderEnabled:true};
  if(!actualAutomation) {
  c.getAutomationSchedule_=()=>({...schedule});
  c.getAutomationSummary_=()=>({schedule:{...schedule},enabled:false,managedByCurrentUser:true,hasManagedOwner:false});
  c.getAutomationStateUnlocked_=()=>({schedule:{...schedule},enabled:false,legacyCount:0,triggers:[],warnings:[],revision:'automation-test'});
  c.saveAutomationScheduleDraft_=value=>Object.assign(schedule,plain(value));
  }
  return {c,props,userProps,sheets,ss,schedule,triggers,activeSheet:()=>activeSheet,isLocked:()=>locked};
}
