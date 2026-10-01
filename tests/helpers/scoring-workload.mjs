import {scoringEnvironment,AdminSheet,sourceId} from './scoring-environment.mjs';

// Synthetic answers only. Count GAS boundaries; Node elapsed time is not GAS latency.
export function scoringWorkload(total=500, sheetCount=5) {
  const e=scoringEnvironment(), metrics={opens:0,lists:0,valueReads:0,displayReads:0,cells:0,timezones:0,digests:0};
  e.source.sheets.clear();
  e.c.scoringSetConfig({nameCol:'B',displayCols:['G'],scoreCols:['D','','','',''],commentCol:'E'});
  const headers=['メール','名前','状態','点数','講評','授業の日付を入力してください','回答','参考1','参考2','参考3'];
  const text='今日の振り返りをもとに考えた内容を説明します。'.repeat(20);
  const mapping=new AdminSheet('対応表',[['元SS_ID','元シート名','クラス名','courseId','メモ']]);e.sheets.set('対応表',mapping);
  for(let s=0;s<sheetCount;s++) {
    const rows=[headers],count=Math.floor(total/sheetCount)+(s<total%sheetCount?1:0);
    for(let i=0;i<count;i++)rows.push([`student-${i%100}@example.invalid`,'架空の生徒','','3','',`2026-09-${String(1+Math.floor(i/100)%28).padStart(2,'0')}`,text,text,text,text]);
    const sheet=new AdminSheet('回答 '+(s+1),rows);e.source.sheets.set(sheet.name,sheet);
    mapping.rows.push([sourceId,sheet.name,'クラス'+s,'course-'+s,'']);
    if(s===0){e.answer=sheet;e.target={spreadsheetId:sourceId,sheetId:sheet.sheetId};}
  }
  for(const book of e.books.values()) {
    const getSheets=book.getSheets.bind(book);book.getSheets=()=>{metrics.lists++;return getSheets();};
    const timezone=book.getSpreadsheetTimeZone.bind(book);book.getSpreadsheetTimeZone=()=>{metrics.timezones++;return timezone();};
    for(const sheet of getSheets()) {
      const getRange=sheet.getRange.bind(sheet);sheet.getRange=(r,c,h=1,w=1)=>{
        const range=getRange(r,c,h,w),values=range.getValues,display=range.getDisplayValues;
        range.getValues=()=>{metrics.valueReads++;metrics.cells+=h*w;return values();};
        range.getDisplayValues=()=>{metrics.displayReads++;metrics.cells+=h*w;return display();};return range;
      };
    }
  }
  const open=e.c.SpreadsheetApp.openById;e.c.SpreadsheetApp.openById=id=>{metrics.opens++;return open(id);};
  const digest=e.c.Utilities.computeDigest;e.c.Utilities.computeDigest=(...args)=>{metrics.digests++;return digest(...args);};
  return {...e,metrics,resetMetrics(){for(const k in metrics)metrics[k]=0;}};
}
