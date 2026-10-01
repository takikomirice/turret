/**
 * 確認用GASのCode.gsの先頭へ一時追加して実行し、終了後は必ず取り除く。
 * 専用の架空データ2ブックを作成する。既存の設定・回答・トリガーは変更しない。
 * メール・Classroomへの送信は呼ばず、送信シートの準備まで検証する。
 */
function aaVerifyManagementColumns() {
  assertWebOperator_();
  const props=PropertiesService.getScriptProperties(),before=props.getProperties();
  const originalBook=getAppSpreadsheet_();
  const digestBook=function(book){return scoringDigest_(book.getSheets().map(function(s){return [s.getSheetId(),s.getName(),s.getDataRange().getValues()];}));};
  const originalDigest=digestBook(originalBook);
  const oldBook=getAppSpreadsheet_,oldConfig=getConfig_,oldScoring=scoringLegacyApiGetConfig_,oldMigrate=migrateConfigStorageUnlocked_,oldWrite=scoringWriteBlocks_;
  const book=SpreadsheetApp.create('turret 管理列再検証 '+new Date().toISOString());
  const second=SpreadsheetApp.create('turret 管理列再検証 別回答ブック '+new Date().toISOString());
  const answer=book.getSheets()[0].setName('検証回答 A'),other=second.getSheets()[0].setName('検証回答 B');
  const cfg=normalizeAppConfig_({formSources:[{id:book.getId()},{id:second.getId()}],formSheetNamePrefix:['検証回答'],emailHeader:'メール',studentNameHeader:'名前',formStatusHeader:'送信状態',scoreSourceHeader:'点数',replyBodyHeader:'返信本文',fields:[{key:'score',sourceHeader:'点数',type:'text'}],messageTemplate:'確認：＜score＞'});
  const scoring=Object.assign(scoringLegacyGetDefaultConfig_(),{nameCol:'B',displayCols:['G','','','',''],scoreCols:['D','','','',''],commentCol:'E',ruleOutputSlots:[],nameHeader:'',displayHeaders:[],scoreHeaders:[],commentHeader:'',lessonDateHeader:'授業の日付を入力してください'});
  const checks=[];
  const check=function(ok,label){if(!ok)throw new Error('検証失敗: '+label);checks.push(label);console.log('PASS '+checks.length+' '+label);};
  const fail=function(fn,pattern,label){let error;try{fn();}catch(e){error=e;}check(!!error&&pattern.test(String(error.message||error)),label);};
  const target=function(sheet){return {spreadsheetId:sheet.getParent().getId(),sheetId:sheet.getSheetId()};};
  const headers=function(sheet){return sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0];};
  const cell=function(sheet,row,name){return sheet.getRange(row,headers(sheet).indexOf(name)+1);};
  const management=function(sheet,row){return JSON.parse(cell(sheet,row,'管理').getValue());};
  const index=function(){return scoringBuildReviewIndex_(null,{summary:true});};
  const group=function(){return Array.from(index().groups.values())[0];};
  const choose=function(sheet,row){const g=group();return scoringChooseResponses({target:target(answer),groupId:g.id,fingerprint:g.fingerprint,accepted:sheet?[scoringResponseKey_(target(sheet),row)]:null});};
  const save=function(sheet,row,score,comment,revision){return scoringSaveRows({target:target(sheet),rows:[{rowNumber:row,scores:[score,'','','',''],comment:comment,configRevision:revision}]});};
  const prepare=function(){return withAppLock_(initializeSheetsUnlocked_);};
  getAppSpreadsheet_=function(){return book;};getConfig_=function(){return cfg;};
  scoringLegacyApiGetConfig_=function(){return JSON.parse(JSON.stringify(scoring));};migrateConfigStorageUnlocked_=function(){};
  try {
    const originalHeaders=['メール','名前','送信状態','点数','講評','授業の日付を入力してください','回答'];
    answer.getRange(1,1,3,7).setValues([originalHeaders,['sample@example.invalid','架空生徒','',3,'','2026-09-30','最初の回答'],['sample@example.invalid','架空生徒','',4,'','2026-09-30','再提出']]);
    other.getRange(1,1,2,7).setValues([originalHeaders,['sample@example.invalid','架空生徒','',2,'','2026-09-30','別ブックの回答']]);
    prepare();check(!!book.getSheetByName('フォーム管理'),'初期準備でフォーム管理を作成');
    book.getSheetByName('生徒一覧').getRange(2,1,1,6).setValues([[1,'sample@example.invalid','架空生徒','検証クラス','test-course','sample-student']]);
    const initial=scoringGetInitData({target:target(answer),reviewMode:'summary'}),g=group();
    check(g.responses.length===3,'2ブックの3回答を同じ重複グループとして検出');
    book.insertSheet('_scoring_saves').getRange(1,1,2,2).setValues([['key','JSON'],[scoringResponseKey_(target(answer),2),'{"state":"pending"}']]);
    book.insertSheet('_scoring_decisions').getRange(1,1,2,2).setValues([['key','JSON'],[g.id,JSON.stringify({fingerprint:g.fingerprint,accepted:[scoringResponseKey_(target(other),2)]})]]);
    prepare();check(!book.getSheetByName('_scoring_saves')&&!book.getSheetByName('_scoring_decisions'),'旧2シートの移行と削除');
    check([answer,other].every(function(s){return headers(s).slice(-2).join(',')==='管理,送信状態';}),'両ブックの右端が管理・送信状態');
    check(management(answer,2).save.state==='pending'&&management(other,2).duplicate.choice==='adopted','未完了保存と別ブックの採用判断を保持');
    check(save(answer,2,5,'列移動前の下書き',initial.configRevision).ok,'列移動前に読み込んだ下書きを公開保存関数で保存');
    check(cell(answer,2,'点数').getValue()===5&&cell(answer,2,'講評').getValue()==='列移動前の下書き','列移動後も正しい採点列へ保存');
    check(management(answer,2).save.state==='complete'&&management(other,2).duplicate.choice==='adopted','保存完了の更新が採用判断を壊さない');
    choose(answer,2);check(management(answer,2).save.state==='complete'&&management(answer,2).duplicate.choice==='adopted','採用変更が保存完了を壊さない');
    check(scoringGetComparison({target:target(answer),groupId:g.id}).responses.every(function(r){return r.answers.every(function(a){return a.header!=='管理';});}),'比較表示に管理JSONを含めない');
    cell(answer,2,'送信状態').setValue('済');save(answer,2,4,'状態保持');choose(other,2);
    check(cell(answer,2,'送信状態').getValue()==='済','再採点・採用変更でも送信済み表示を保持');
    cell(answer,2,'送信状態').clearContent();
    choose(null);check(group().responses.every(function(r){return r.state==='pending'&&!r.eligible;}),'採用解除で全回答を返却保留');
    choose(other,2);
    other.appendRow(headers(other).map(function(h){return ({'メール':'sample@example.invalid','名前':'架空生徒','点数':3,'授業の日付を入力してください':'2026-09-30','回答':'追加された回答'})[h]||'';}));
    check(group().responses.length===4&&group().responses.every(function(r){return r.state==='pending'&&!r.eligible;}),'新しい回答が増えたら既存の採用判断を無効化');
    fail(function(){scoringChooseResponses({target:target(answer),groupId:g.id,fingerprint:g.fingerprint,accepted:[scoringResponseKey_(target(answer),2)]});},/更新|読み直し/,'古い比較画面からの採用変更を拒否');
    choose(answer,2);
    // 2つ目のブックへの通信失敗を注入し、実シートに部分保存が残った状態を検証する。
    let writes=0;
    scoringWriteBlocks_=function(sheet,blocks,onError){if(!onError&&++writes===2)throw new Error('検証用の書き込み失敗');return oldWrite(sheet,blocks,onError);};
    try{fail(function(){choose(other,2);},/検証用/,'ブック間の採用書き込み途中で例外を検出');}finally{scoringWriteBlocks_=oldWrite;}
    check(group().responses.every(function(r){return r.state==='pending'&&!r.eligible;}),'部分保存が残っても全回答を返却保留');
    choose(answer,2);check(group().responses.filter(function(r){return r.eligible;}).length===1,'採用変更の再試行で1回答だけ返却可能に復旧');
    // 点数書き込みの失敗を注入する。pendingの永続化と読み戻しは実サービスを使う。
    scoringWriteBlocks_=function(sheet,blocks,onError){if(onError){blocks.forEach(function(b){onError(b,new Error('検証用の採点保存失敗'));});return;}return oldWrite(sheet,blocks,onError);};
    try{check(!save(answer,2,1,'失敗する採点').ok,'採点書き込み失敗を保存成功と扱わない');}finally{scoringWriteBlocks_=oldWrite;}
    check(management(answer,2).save.state==='pending'&&management(answer,2).duplicate.choice==='adopted','失敗時は採用判断を残してpendingを永続化');
    check(!scoringMayReturn_(scoringBuildReturnGuard_(),makeResponseKey_(book.getId(),answer.getName(),2)),'採用済みでも未完了保存は返却不可');
    prepareSendData();check(book.getSheetByName('送信シート').getLastRow()===1,'未完了の採点を送信シートへ準備しない');
    check(save(answer,2,5,'再保存成功').ok&&management(answer,2).save.state==='complete','再保存で完了へ復旧');
    const kept=cell(answer,2,'管理').getValue(),scoreBefore=cell(answer,2,'点数').getValue();
    cell(answer,2,'管理').setValue('{broken');
    check(!save(answer,2,1,'上書き禁止').ok&&cell(answer,2,'点数').getValue()===scoreBefore&&cell(answer,2,'管理').getValue()==='{broken','壊れた管理JSONと採点値を上書きしない');
    fail(function(){scoringBuildReturnGuard_();},/管理/,'壊れた管理JSONがあると返却判定を停止');
    cell(answer,2,'管理').setValue(kept);
    const oldSave=book.insertSheet('_scoring_saves');oldSave.getRange(1,1,2,2).setValues([['key','JSON'],[JSON.stringify(['missing',123,2]),'{"state":"pending"}']]);
    book.insertSheet('_scoring_decisions').getRange(1,1,1,2).setValues([['key','JSON']]);
    fail(prepare,/移行先/,'移行先不明の旧記録で処理を停止');
    check(!!book.getSheetByName('_scoring_saves')&&!!book.getSheetByName('_scoring_decisions'),'移行失敗時は旧2シートを保持');
    oldSave.getRange(2,1).setValue(scoringResponseKey_(target(answer),2));prepare();
    check(management(answer,2).save.state==='complete'&&!book.getSheetByName('_scoring_saves'),'移行再試行でも最新completeを旧pendingへ戻さない');
    const beforePrepare=cell(answer,2,'管理').getValue();prepareSendData();
    const send=book.getSheetByName('送信シート'),sendRows=send.getDataRange().getValues(),sendHeaders=sendRows[0];
    check(sendRows.length===2&&sendRows[1][sendHeaders.indexOf('元SS_ID')]===book.getId()&&sendRows[1][sendHeaders.indexOf('元行番号')]===2,'採用かつ保存完了の1回答だけ送信準備');
    check(sendRows[1][sendHeaders.indexOf('返信本文')]==='確認：5','最新の採点値から返信本文を生成');
    check(cell(answer,2,'送信状態').getValue()==='準備〇'&&cell(answer,2,'管理').getValue()===beforePrepare,'送信準備による状態更新が管理JSONを変えない');
    prepareSendData();prepare();check(send.getLastRow()===2&&[book,second].every(function(b){return !b.getSheetByName('_scoring_saves')&&!b.getSheetByName('_scoring_decisions');}),'再実行でも送信行や旧内部シートを増やさない');
    check([answer,other].every(function(s){return headers(s).filter(function(h){return h==='管理';}).length===1;}),'処理を重ねても管理列は1列だけ');
    const after=props.getProperties();check(Object.keys(before).length===Object.keys(after).length&&Object.keys(before).every(function(k){return before[k]===after[k];}),'既存のScriptPropertiesを保持');
    check(digestBook(originalBook)===originalDigest,'既存の運用ブック全シートの値を保持');
    console.log(JSON.stringify({ok:true,checks:checks,books:[book.getUrl(),second.getUrl()],headers:[headers(answer),headers(other)]}));
  } catch(error) {
    console.log(JSON.stringify({ok:false,checks:checks,books:[book.getUrl(),second.getUrl()],error:String(error.message||error)}));throw error;
  } finally {
    getAppSpreadsheet_=oldBook;getConfig_=oldConfig;scoringLegacyApiGetConfig_=oldScoring;migrateConfigStorageUnlocked_=oldMigrate;scoringWriteBlocks_=oldWrite;
  }
}
