import { readFileSync } from 'node:fs';
import vm from 'node:vm';
export const plain = value => JSON.parse(JSON.stringify(value));
export function sheet(initial = []) {
 const rows = initial.map(row => row.slice());
 let filter=null;
 return {rows,getFilter(){return filter;},getMaxRows(){return 1000;}, getLastRow(){let n=rows.length;while(n&&rows[n-1].every(v=>v===''||v==null))n--;return n;},getLastColumn(){return Math.max(0,...rows.map(r=>r.length));},
 getRange(r,c,h=1,w=1){const read=()=>Array.from({length:h},(_,y)=>Array.from({length:w},(_,x)=>rows[r+y-1]?.[c+x-1]??''));const write=v=>v.forEach((a,y)=>{while(rows.length<r+y)rows.push([]);a.forEach((v,x)=>rows[r+y-1][c+x-1]=v);});return {createFilter(){if(filter)throw Error('filter exists');return filter={row:r,column:c,rows:h,columns:w};},getValues:read,getDisplayValues:()=>read().map(a=>a.map(String)),getValue:()=>read()[0][0],setValue(v){write([[v]]);return this;},setValues(v){write(v);return this;},clearContent(){write(Array.from({length:h},()=>Array(w).fill('')));return this;}};},
 getDataRange(){return this.getRange(1,1,Math.max(1,this.getLastRow()),Math.max(1,this.getLastColumn()));},clear(){rows.length=0;return this;},clearContents(){rows.length=0;return this;},appendRow(r){rows.splice(this.getLastRow(),0,r.slice());return this;},deleteRow(r){rows.splice(r-1,1);return this;},setFrozenRows(){return this;},autoResizeColumns(){return this;}};
}
export function harness(){
 const sheets=new Map(), properties=new Map(); const props={getProperty:k=>properties.get(k)??null,setProperty(k,v){properties.set(k,String(v));},deleteProperty(k){properties.delete(k);}};
 const ss={getSheetByName:k=>sheets.get(k)||null,insertSheet(k){const s=sheet();sheets.set(k,s);return s;},deleteSheet(s){for(const[k,v]of sheets)if(v===s)sheets.delete(k);}};
 const context=vm.createContext({Logger:{log(){}},LockService:{getScriptLock:()=>({tryLock:()=>true,releaseLock(){}})},SpreadsheetApp:{getActiveSpreadsheet:()=>ss,flush(){}},PropertiesService:{getScriptProperties:()=>props},Session:{getActiveUser:()=> 'teacher@example.com'},Classroom:{Courses:{list(){return {};},Students:{list(){return {};}}}}});
 vm.runInContext(readFileSync(new URL('../Code.gs',import.meta.url),'utf8'),context);
 return {context,sheets,properties,props,ss};
}
export function completeConfig(c){return c.normalizeAppConfig_({formSources:['abcdefghijklmnopqrstuv'],formSheetNamePrefix:['Responses'],emailHeader:'Email',studentNameHeader:'Name',formStatusHeader:'Status',scoreSourceHeader:'Grade',replyBodyHeader:'Reply',fields:[{key:'grade',sourceHeader:'Grade',type:'text'}],messageTemplate:'＜grade＞',reminderTo:[]});}
