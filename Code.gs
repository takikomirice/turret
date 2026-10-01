/******************************************************
 * Google Classroom 評価自動送信 GAS
 *
 * - 管理画面で設定し、ScriptPropertiesへ保存する。
 * - 旧設定シートは初回管理画面表示時に内部保存へ移行して削除する。
 * - 別期間・別プロジェクトへの設定引き継ぎはJSONを使う。
 ******************************************************/

const SETTINGS_SHEET_NAME = '設定(編集不可)';
const INPUT_SETTINGS_SHEET_NAME = '設定シート';
const CONFIG_SHEET_NAME = 'Config';
const EVAL_SHEET_NAME = '評価データ';
const SEND_SHEET_NAME = '送信シート';
const STUDENT_SHEET_NAME = '生徒一覧';
const ERROR_SHEET_NAME = 'エラー';
const MAPPING_SHEET_NAME = '対応表';

const EVAL_BASE_HEADERS = [
  '元SS_ID',
  '元シート名',
  '元行番号',
  'メールアドレス',
  '名前',
  '処理状態'
];

const SEND_BASE_HEADERS = [
  '元SS_ID',
  '元シート名',
  '元行番号',
  'No',
  'メールアドレス',
  '名前',
  'クラス名',
  'コースID',
  'studentId',
  '送信状態'
];

const STUDENT_SHEET_HEADERS = ['No', 'メールアドレス', '名前', 'クラス名', 'コースID', 'studentId'];
const MAPPING_SHEET_HEADERS = ['元SS_ID', '元スプシ名', '元シート名', 'クラス名', 'courseId', 'メモ'];
const SEND_REVIEW_STATUS = '送信確認待ち';
const SCORING_MANAGEMENT_HEADER = '管理';

// === 独立した採点画面で使用できる対象 ===
function internalSheetKind_(name) {
  if (['採点テンプレ', '_templates', '_rules', '_scoring_decisions', '_scoring_saves'].includes(name)) return 'scoring';
  return ['クラス一覧', STUDENT_SHEET_NAME, INPUT_SETTINGS_SHEET_NAME, CONFIG_SHEET_NAME,
    SETTINGS_SHEET_NAME, EVAL_SHEET_NAME, SEND_SHEET_NAME, ERROR_SHEET_NAME, MAPPING_SHEET_NAME,
    'フォーム管理'].includes(name) ? 'operation' : '';
}
function scoringIsInternalSheet_(name) { return !!internalSheetKind_(name); }

/** キャッシュはこの実行中だけ使う。次のリクエストでは新しい回答と設定を読み取る。 */
function scoringReadContext_(books) {
  const operation=getAppSpreadsheet_();
  return {operation:operation,app:getConfig_(),books:books || new Map(),sheetLists:new Map(),snapshots:new Map(),targets:null,
    managed:typeof getManagedRecords_==='function' ? getManagedRecords_() : []};
}
function scoringBook_(read,id) {
  if(!read.books.has(id))read.books.set(id,id===read.operation.getId()?read.operation:SpreadsheetApp.openById(id));
  return read.books.get(id);
}
function scoringSheets_(read,id) {
  if(!read.sheetLists.has(id))read.sheetLists.set(id,scoringBook_(read,id).getSheets());
  return read.sheetLists.get(id);
}
function scoringSnapshot_(read,target) {
  if(!read.snapshots.has(target.key)) {
    const sheet=scoringSheets_(read,target.spreadsheetId).find(function(s){return String(s.getSheetId())===String(target.sheetId);});
    if(!sheet)throw new Error('回答シートが変更されました。読み直してください。');
    const lastRow=sheet.getLastRow(),lastCol=sheet.getLastColumn();
    read.snapshots.set(target.key,{sheet:sheet,lastRow:lastRow,lastCol:lastCol,values:lastRow && lastCol ? sheet.getRange(1,1,lastRow,lastCol).getValues() : [],display:null});
  }
  return read.snapshots.get(target.key);
}
function getScoringTargets_(books,read) {
  read=read || scoringReadContext_(books);
  if(read.targets)return read.targets;
  const config = read.app, operation = read.operation;
  const records = read.managed.filter(function(r) {
    return r.kind === 'form' && r.responseSheetId != null && r.ownerSpreadsheetId === operation.getId() && r.ownerScriptId === ScriptApp.getScriptId();
  });
  const sources = new Map();
  (config.formSources || []).forEach(function(source) { sources.set(source.id, source); });
  if (records.length && !sources.has(operation.getId())) sources.set(operation.getId(), {id:operation.getId()});
  const targets = [];
  sources.forEach(function(source, id) {
    const book=scoringBook_(read,id),bookName=book.getName();
    scoringSheets_(read,id).forEach(function(sheet) {
      const name = sheet.getName();
      if (scoringIsInternalSheet_(name)) return;
      const record = id === operation.getId() && records.find(function(r) { return String(r.responseSheetId) === String(sheet.getSheetId()); });
      const returnEnabled = (config.formSources || []).some(function(s) { return s.id === id; }) && sheetMatchesPrefixes_(name, config.formSheetNamePrefix);
      if (!returnEnabled && !record) return;
      targets.push({key:JSON.stringify([id, sheet.getSheetId()]), spreadsheetId:id, sheetId:sheet.getSheetId(),
        spreadsheetName:bookName, sheetName:name, label:bookName + ' / ' + name,
        courseId:String(record && record.courseId || source.courseId || ''), returnEnabled:returnEnabled,
        warning:returnEnabled ? '' : '返却対象の登録または接頭辞が一致していません。管理画面で確認してください。',
        url:'https://docs.google.com/spreadsheets/d/' + id + '/edit#gid=' + sheet.getSheetId()});
    });
  });
  read.targets=targets;return targets;
}

function resolveScoringTarget_(target,read) {
  read=read || scoringReadContext_();
  if (typeof target === 'string') {
    try { const pair=JSON.parse(target); target={spreadsheetId:pair[0],sheetId:pair[1]}; } catch(error) { target=null; }
  }
  if (!target || !target.spreadsheetId || target.sheetId == null) throw new Error('採点対象の回答シートを選択してください。');
  const found=getScoringTargets_(null,read).find(function(t) {return t.spreadsheetId === target.spreadsheetId && String(t.sheetId) === String(target.sheetId);});
  if (!found) throw new Error('登録された採点対象ではありません。対象一覧を読み直してください。');
  const book=scoringBook_(read,found.spreadsheetId);
  const sheet=scoringSheets_(read,found.spreadsheetId).find(function(s) { return String(s.getSheetId()) === String(found.sheetId); });
  if (!sheet || scoringIsInternalSheet_(sheet.getName())) throw new Error('採点対象が変更されました。一覧を読み直してください。');
  return {target:found,book:book,sheet:sheet};
}

function scoringGetTargets() { assertWebOperator_(); return getScoringTargets_(); }

function scoringGetConfig() { assertWebOperator_(); return scoringLegacyApiGetConfig_(); }
function scoringSetConfig(config) {
  assertWebOperator_();
  return withAppLock_(function() {
    scoringValidateConfigInput_(config);
    const copy=Object.assign({},config,{sheetName:''});
    return scoringLegacyApiSetConfig_(copy);
  });
}

function scoringGetInitData(params) {
  assertWebOperator_();
  const read=scoringReadContext_(),context=resolveScoringTarget_(params && (params.target || params.sheetName),read);
  const snapshot=scoringSnapshot_(read,context.target),config=scoringLegacyApiGetConfig_();
  scoringValidateResolvedConfig_(config,context.sheet,false,snapshot.values[(config.headerRow || 1)-1] || []);
  const data=scoringLegacyApiGetInitData_(Object.assign({},params,{sheetName:context.target}),snapshot,config);
  data.target=context.target;
  data.sheetName=context.target.key;
  data.targetLabel=context.target.label;
  data.configRevision=scoringWriteRevision_(data.config);
  const summary=!!(params && params.reviewMode==='summary');
  data.review=scoringReviewForTarget_(scoringBuildReviewIndex_(null,{read:read,summary:summary}),context.target,summary);
  return data;
}

function scoringSaveRows(payload) {
  assertWebOperator_();
  if (!payload || !Array.isArray(payload.rows) || payload.rows.length > 100) throw new Error('保存する行は100件以内で指定してください。');
  return withAppLock_(function() {
    if (typeof assertNoManagedFormUpdate_ === 'function') assertNoManagedFormUpdate_();
    const context=resolveScoringTarget_(payload.target || payload.sheetName), sheet=context.sheet;
    const config=scoringLegacyApiGetConfig_();
    scoringValidateResolvedConfig_(config,sheet,true);
    const revision=scoringWriteRevision_(config),lastRow=sheet.getLastRow();
    const entries=payload.rows.map(function(row) {
      const result={rowNumber:row && row.rowNumber,ok:false};
      try {
        if(row && row.configRevision && row.configRevision!==revision && !scoringAcceptLayoutRevision_(sheet,row.configRevision))throw new Error('採点の列設定が変更されました。元の列設定に戻して保存するか、下書きの内容と保存先を確認してください。');
        if (!row || !Number.isInteger(row.rowNumber) || row.rowNumber <= config.headerRow || row.rowNumber < config.startRow || row.rowNumber > lastRow) throw new Error('保存する回答行が不正です。');
        const normalized=scoringLegacyNormalizeRowPayload_(row,config);
        return {key:scoringResponseKey_(context.target,row.rowNumber),normalized:normalized,result:result};
      } catch(error) { result.error=String(error.message || error);return {result:result}; }
    });
    if(entries.some(function(entry){return !entry.result.error;})) {
      scoringManagementColumn_(sheet,true);
      scoringValidateResolvedConfig_(config,sheet,true);
      const plan=scoringColumnPlan_(config);
      entries.forEach(function(entry){if(!entry.result.error)entry.cells=scoringRowCells_(plan,entry.normalized);});
    }
    // 同じ行を繰り返し編集しても、空欄の得点を含めて入力順を保つ。
    let batch=[],seen=new Set();
    entries.forEach(function(entry){
      if(entry.result.error)return;
      if(seen.has(entry.key)){scoringSaveBatch_(sheet,batch);batch=[];seen.clear();}
      batch.push(entry);seen.add(entry.key);
    });
    if(batch.length)scoringSaveBatch_(sheet,batch);
    const results=entries.map(function(entry){return entry.result;});
    return {ok:results.every(function(r){return r.ok;}),sheetName:context.target.key,results:results};
  });
}

/** 保留中の記録は回答より先に永続化し、完了記録は回答の保存確認後に更新する。 */
function scoringSaveBatch_(sheet,entries) {
  try {
    const journal=scoringManagementColumn_(sheet,true);
    scoringWriteSaveRecords_(journal,entries,'pending');
    const blocks=[];
    entries.forEach(function(entry){
      let block;
      entry.cells.forEach(function(cell){
        if(!block || cell.col!==block.col+block.values.length){
          block={row:entry.result.rowNumber,col:cell.col,values:[],entry:entry};blocks.push(block);
        }
        block.values.push(scoringSheetLiteral_(cell.value));
      });
    });
    scoringWriteBlocks_(sheet,blocks,function(block,error){block.entry.result.error=String(error.message || error);});
    SpreadsheetApp.flush();
    const completed=entries.filter(function(entry){return !entry.result.error;});
    scoringWriteSaveRecords_(journal,completed,'complete');
    completed.forEach(function(entry){entry.result.ok=true;});
  } catch(error) {
    entries.forEach(function(entry){if(!entry.result.error)entry.result.error=String(error.message || error);});
  }
}

/** 書き込み列が同じ隣接行だけをまとめる。間のセルを既存値で埋めない。 */
function scoringWriteBlocks_(sheet,blocks,onError) {
  blocks.sort(function(a,b){return a.col-b.col || a.values.length-b.values.length || a.row-b.row;});
  for(let i=0;i<blocks.length;) {
    let end=i+1;
    while(end<blocks.length && blocks[end].col===blocks[i].col && blocks[end].values.length===blocks[i].values.length && blocks[end].row===blocks[end-1].row+1)end++;
    const group=blocks.slice(i,end);
    const write=function(rows){sheet.getRange(rows[0].row,rows[0].col,rows.length,rows[0].values.length).setValues(rows.map(function(row){return row.values;}));};
    try { write(group); }
    catch(error) {
      if(!onError)throw error;
      // セル単位の書き込みは再実行しても同じ結果になる。一括書き込み失敗時は行ごとに失敗を切り分ける。
      if(group.length===1)onError(group[0],error);
      else group.forEach(function(block){try{write([block]);}catch(rowError){onError(block,rowError);}});
    }
    i=end;
  }
}

function scoringWriteSaveRecords_(journal,entries,state) {
  if(!entries.length)return;
  scoringPatchManagement_(journal,entries.map(function(entry){return {row:entry.result.rowNumber,patch:{save:{state:state}}};}));
}

function scoringGetTargetSpreadsheetEditUrl(params) { assertWebOperator_(); return resolveScoringTarget_(params && (params.target || params.sheetName)).target.url; }
function scoringGetRulesEditUrl() { assertWebOperator_(); return scoringGetTemplateTable().url; }
function scoringGetRuleTemplates() { assertWebOperator_(); return scoringTableTemplateList_(); }
function scoringGetRules(params) { assertWebOperator_(); return scoringTableRules_(params); }
function scoringExportConfig() { assertWebOperator_(); return scoringLegacyApiExportConfig_(); }
function scoringExportTemplatePack() { assertWebOperator_(); return {meta:{type:'turret-grading-template-pack',version:1},templates:scoringGetTemplateTable().templates}; }
function scoringSetSlotLabels(labels) { assertWebOperator_(); return withAppLock_(function(){return scoringLegacyApiSetSlotLabels_(labels);}); }
function scoringEnsureRuleSheets() { assertWebOperator_(); return withAppLock_(function(){scoringEnsureTemplateTable_();return {ok:true};}); }
function scoringImportTemplatePack(payload) {
  assertWebOperator_();
  return withAppLock_(function(){
    if(!payload || !payload.meta || payload.meta.type!=='turret-grading-template-pack' || payload.meta.version!==1)throw new Error('採点テンプレ形式のパックを指定してください。旧形式は取り込めません。');
    const current=scoringReadTemplateTable_(),incoming=scoringNormalizeTemplates_(payload.templates);
    const names=new Set(current.templates.map(function(t){return t.name;}));
    incoming.forEach(function(t){if(names.has(t.name))throw new Error('同名のテンプレートがあります: '+t.name);});
    scoringWriteTemplateTable_(current.templates.concat(incoming),current.revision);
    return {ok:true,imported:{templates:incoming.length,rules:incoming.reduce(function(n,t){return n+t.rules.length;},0)}};
  });
}
function scoringImportConfig(payload) {
  assertWebOperator_();
  return withAppLock_(function(){
    const copy=JSON.parse(JSON.stringify(payload));
    (copy.config || copy).sheetName='';
    scoringValidateConfigInput_(copy.config || copy);
    return scoringLegacyApiImportConfig_(copy);
  });
}
function scoringRestoreLatestConfigBackup() { assertWebOperator_(); return withAppLock_(scoringLegacyApiRestoreLatestConfigBackup_); }

// 採点テンプレはこの表を唯一の保存先とし、画面と直編集で同じ形式を使う。
function scoringTemplateHeaders_() {
  return ['テンプレート名','枠1','枠2','枠3','枠4','枠5','出力1','出力2','出力3','使用する','採点画面に表示',
    '枠1の点数','枠2の点数','枠3の点数','枠4の点数','枠5の点数'];
}
function scoringParseCondition_(value) {
  const text=String(value==null?'':value).trim().replace(/^=</,'<=').replace(/^=>/,'>=');
  if(!text)return {op:'any'};
  const match=text.match(/^(>=|<=|>|<|=)?\s*(-?(?:\d+(?:\.\d*)?|\.\d+))$/);
  if(!match || !Number.isFinite(Number(match[2])))throw new Error('条件は数値、>1、>=2、<3、<=2 の形式で入力してください: '+text);
  return {op:match[1] || '=',value:Number(match[2])};
}
function scoringConditionMatches_(condition,value) {
  if(condition.op==='any')return true;
  if(value==null || String(value).trim()==='' || !Number.isFinite(Number(value)))return false;
  const n=Number(value),v=condition.value;
  return condition.op==='='?n===v:condition.op==='>'?n>v:condition.op==='>='?n>=v:condition.op==='<'?n<v:n<=v;
}
function scoringTemplateText_(text,scores) {
  return String(text).replace(/\{(枠[1-5]|合計)\}/g,function(_,key){
    if(key==='合計')return String(scores.reduce(function(sum,v){return sum+(v!=null && String(v).trim()!=='' && Number.isFinite(Number(v))?Number(v):0);},0));
    const value=scores[Number(key.slice(1))-1];return value==null?'':String(value);
  });
}
function scoringTemplateFlag_(value,label) {
  if(value==null || String(value).trim()==='')return true;
  if(value===true || /^(1|true|はい)$/i.test(String(value).trim()))return true;
  if(value===false || /^(0|false|いいえ)$/i.test(String(value).trim()))return false;
  throw new Error(label+' は 1（はい）または 0（いいえ）で入力してください。');
}
function scoringNormalizeTemplates_(templates) {
  if(!Array.isArray(templates) || templates.length>100)throw new Error('テンプレートは100件以内で指定してください。');
  const names=new Set();let totalRules=0;
  return templates.map(function(t){
    if(!t || !String(t.name || '').trim())throw new Error('テンプレート名が空欄です。');
    const name=String(t.name).trim();
    if(name.length>100 || names.has(name))throw new Error('テンプレート名が重複、または長すぎます: '+name);
    names.add(name);
    if(!Array.isArray(t.domains) || t.domains.length!==5)throw new Error(name+': 枠1〜5の点数を設定してください。');
    const domains=t.domains.map(function(domain,i){
      if(typeof domain==='string')domain=domain.trim()?domain.split(/[,、\s]+/):[];
      if(!Array.isArray(domain) || domain.length>20)throw new Error(name+': 枠'+(i+1)+'の点数は20種類以内で指定してください。');
      const values=domain.map(function(v){if(!/^-?(?:\d+(?:\.\d*)?|\.\d+)$/.test(String(v).trim()) || !Number.isFinite(Number(v)))throw new Error(name+': 点数には数値を指定してください。');return Number(v);});
      if(new Set(values).size!==values.length)throw new Error(name+': 枠'+(i+1)+'の点数が重複しています。');
      return values.sort(function(a,b){return a-b;});
    });
    if(!Array.isArray(t.rules) || t.rules.length>500)throw new Error(name+': ルールは500行以内で指定してください。');
    totalRules+=t.rules.length;if(totalRules>2000)throw new Error('表全体のルールは2000行以内で指定してください。');
    const rules=t.rules.map(function(r,i){
      if(!r || !Array.isArray(r.conditions) || r.conditions.length!==5 || !Array.isArray(r.outputs) || r.outputs.length!==3)throw new Error(name+': ルール'+(i+1)+'は枠5列・出力3列で指定してください。');
      const conditions=r.conditions.map(function(v){return String(v==null?'':v).trim();});
      const outputs=r.outputs.map(function(v){const text=String(v==null?'':v);if(text.length>10000)throw new Error('出力は10000文字以内で入力してください。');return text;});
      return {conditions:conditions,outputs:outputs,enabled:scoringTemplateFlag_(r.enabled,'使用する'),visible:scoringTemplateFlag_(r.visible,'採点画面に表示'),sheetRow:Number(r.sheetRow)||null};
    }).filter(function(r){return r.conditions.some(Boolean) || r.outputs.some(function(v){return v!=='';});});
    return {name:name,domains:domains,rules:rules};
  });
}
function scoringReadTemplateTable_() {
  const ss=getAppSpreadsheet_(),sheet=ss.getSheetByName('採点テンプレ'),headers=scoringTemplateHeaders_();
  const values=sheet && sheet.getLastRow()?sheet.getRange(1,1,sheet.getLastRow(),Math.max(headers.length,sheet.getLastColumn())).getDisplayValues():[];
  const revision=scoringDigest_(values),templates=[],byName=new Map(),metadata=new Map();
  // 旧版のテンプレート単位の使用列は読み捨て、次の保存で消去する。
  const legacyWidth=values.length && values[0][16]==='テンプレートを使用'?17:16;
  if(values.length && (values[0].slice(0,headers.length).join('\t')!==headers.join('\t') || values.some(function(r){return r.slice(legacyWidth).some(Boolean);})))throw new Error('採点テンプレのヘッダ構成が異なります。列名と余分な列を確認してください。');
  values.slice(1).forEach(function(row,index){
    if(!row.slice(0,16).some(function(v){return String(v).trim()!=='';}))return;
    const name=String(row[0] || '').trim();if(!name)throw new Error('採点テンプレ '+(index+2)+'行: テンプレート名が空欄です。');
    if(!byName.has(name)){const t={name:name,domains:[[],[],[],[],[]],rules:[]};byName.set(name,t);templates.push(t);metadata.set(name,{});}
    const t=byName.get(name),meta=metadata.get(name);
    for(let i=11;i<16;i++) {
      if(String(row[i] || '').trim()==='')continue;
      if(Object.prototype.hasOwnProperty.call(meta,i) && meta[i]!==String(row[i]).trim())throw new Error(name+': '+headers[i]+'が行によって異なります。最初の行だけに記入してください。');
      meta[i]=String(row[i]).trim();t.domains[i-11]=row[i];
    }
    if(row.slice(1,9).some(function(v){return v!=='';}))t.rules.push({conditions:row.slice(1,6),outputs:row.slice(6,9),enabled:scoringTemplateFlag_(row[9],'使用する'),visible:scoringTemplateFlag_(row[10],'採点画面に表示'),sheetRow:index+2});
  });
  return {templates:scoringNormalizeTemplates_(templates),revision:revision,url:'https://docs.google.com/spreadsheets/d/'+ss.getId()+'/edit'+(sheet?'#gid='+sheet.getSheetId():'')};
}
function scoringEnsureTemplateTable_() {
  const ss=getAppSpreadsheet_();let sheet=ss.getSheetByName('採点テンプレ');
  if(!sheet)sheet=ss.insertSheet('採点テンプレ');
  if(!sheet.getLastRow()){
    sheet.getRange(1,1,sheet.getMaxRows(),16).setNumberFormat('@');
    sheet.getRange(1,1,1,16).setValues([scoringTemplateHeaders_()]).setFontWeight('bold').setBackground('#e7f0ea');sheet.setFrozenRows(1);
    sheet.setColumnWidths(1,1,180);sheet.setColumnWidths(2,5,80);sheet.setColumnWidths(7,3,240);sheet.setColumnWidths(10,7,125);
    sheet.getRange(1,1,sheet.getMaxRows(),16).setWrap(true);
  }
  return sheet;
}
function scoringGetTemplateTable() {assertWebOperator_();return scoringReadTemplateTable_();}
function scoringWriteTemplateTable_(templates,revision,onBeforeWrite) {
  const normalized=scoringNormalizeTemplates_(templates),validation=scoringValidateTemplateTable_({templates:normalized});
  const errors=validation.issues.filter(function(i){return i.severity==='error';});
  if(errors.length)throw new Error(errors.slice(0,5).map(function(i){return i.message;}).join('\n'));
  const current=scoringReadTemplateTable_();
  if(typeof revision!=='string' || revision!==current.revision)throw new Error('シートが変更されています。編集内容をコピーしてから再読込してください。');
  const rows=[scoringTemplateHeaders_()];
  normalized.forEach(function(t){
    const rules=t.rules.length?t.rules:[{conditions:['','','','',''],outputs:['','',''],enabled:true,visible:true}];
    rules.forEach(function(r,i){rows.push([t.name].concat(r.conditions,r.outputs,[r.enabled?'1':'0',r.visible?'1':'0'],i===0?t.domains.map(function(d){return d.join(',');}):['','','','','']));});
  });
  // 競合検査を通り、実際の変更に入る時点を一括読み込み側へ通知する。
  if(onBeforeWrite)onBeforeWrite();
  const sheet=scoringEnsureTemplateTable_(),oldRows=sheet.getLastRow();
  if(sheet.getMaxRows()<rows.length)sheet.insertRowsAfter(sheet.getMaxRows(),rows.length-sheet.getMaxRows());
  // 一度の書込みに末尾の消去も含め、先に既存内容を消さない。
  const width=Math.max(16,sheet.getLastColumn());
  while(rows.length<oldRows)rows.push(Array(16).fill(''));
  sheet.getRange(1,1,rows.length,width).setNumberFormat('@').setValues(rows.map(function(row){return row.concat(Array(width-row.length).fill('')).map(scoringSheetLiteral_);}));
  sheet.setFrozenRows(1);
  return scoringReadTemplateTable_();
}
function scoringSaveTemplateTable(payload) {
  assertWebOperator_();return withAppLock_(function(){return scoringWriteTemplateTable_(payload && payload.templates,payload && payload.revision);});
}
function scoringValidateTemplateTable(payload) {assertWebOperator_();return scoringValidateTemplateTable_(payload);}
function scoringValidateTemplateTable_(payload) {
  const templates=scoringNormalizeTemplates_(payload && payload.templates),issues=[];let complete=true;
  templates.forEach(function(t){
    const active=[],usedOutputs=new Set();
    function issue(code,severity,r,output,message,extra){issues.push(Object.assign({code:code,severity:severity,template:t.name,rule:r?r.index+1:null,sheetRow:r?r.rule.sheetRow:null,output:output,message:t.name+': '+message},extra || {}));}
    t.rules.forEach(function(rule,index){
      if(!rule.enabled)return;
      const r={rule:rule,index:index,sets:[],conditions:[]};let invalid=false;
      rule.conditions.forEach(function(cell,i){try{r.conditions.push(scoringParseCondition_(cell));}catch(e){invalid=true;issue('syntax','error',r,null,'ルール'+(index+1)+' 枠'+(i+1)+' '+e.message);}});
      rule.outputs.forEach(function(v,i){if(v!=='')usedOutputs.add(i);});
      if(invalid)return;
      r.sets=t.domains.map(function(domain,i){return (domain.length?domain:[null]).filter(function(v){return scoringConditionMatches_(r.conditions[i],v);});});
      if(r.sets.some(function(s){return !s.length;})){issue('impossible','warning',r,null,'ルール'+(index+1)+'は設定した点数では成立しません。');return;}
      if(!rule.outputs.some(function(v){return v!=='';}))issue('no-output','warning',r,null,'ルール'+(index+1)+'には出力がありません。');
      active.push(r);
    });
    // 組合せを全展開せず、各枠の集合の交差で重複の具体例を求める。
    let overlapCount=0;
    for(let i=0;i<active.length;i++)for(let j=i+1;j<active.length;j++){
      const a=active[i],b=active[j],intersection=a.sets.map(function(s,k){return s.filter(function(v){return b.sets[k].indexOf(v)>=0;});});
      if(intersection.some(function(s){return !s.length;}))continue;
      for(let o=0;o<3;o++)if(a.rule.outputs[o]!=='' && b.rule.outputs[o]!==''){
        if(overlapCount++<200)issue('overlap','warning',b,o+1,'ルール'+(a.index+1)+'と'+(b.index+1)+'の出力'+(o+1)+'が重複します。上の行を優先します。',{otherRule:a.index+1,otherSheetRow:a.rule.sheetRow,example:intersection.map(function(s){return s[0];})});
      }
    }
    if(overlapCount>200)issue('issue-limit','warning',null,null,'重複の表示は最初の200件です（全'+overlapCount+'件）。');
    const domains=t.domains.map(function(d){return d.length?d:[null];}),count=domains.reduce(function(n,d){return n*d.length;},1);
    if(count>50000 || count*Math.max(active.length,1)>5000000){complete=false;issue('analysis-limit','warning',null,null,'組合せが多いため、条件漏れと到達不能の全件検査を省略しました。点数の種類かルール数を減らしてください。');return;}
    const wins=new Set(),gaps={},scores=[];
    function visit(slot){
      if(slot<5){domains[slot].forEach(function(v){scores[slot]=v;visit(slot+1);});return;}
      const matched=[false,false,false];
      active.forEach(function(r){if(!r.sets.every(function(s,i){return s.indexOf(scores[i])>=0;}))return;
        for(let o=0;o<3;o++)if(!matched[o] && r.rule.outputs[o]!==''){matched[o]=true;wins.add(r.index+':'+o);}
      });
      usedOutputs.forEach(function(o){if(!matched[o] && !gaps[o])gaps[o]=scores.slice();});
    }
    visit(0);
    active.forEach(function(r){r.rule.outputs.forEach(function(v,o){if(v!=='' && !wins.has(r.index+':'+o))issue('unreachable','warning',r,o+1,'ルール'+(r.index+1)+'の出力'+(o+1)+'は上のルールに覆われ、一度も使われません。');});});
    Object.keys(gaps).forEach(function(o){issue('gap','warning',null,Number(o)+1,'出力'+(Number(o)+1)+'に当てはまるルールがない点数の組合せがあります。',{example:gaps[o]});});
  });
  return {ok:!issues.some(function(i){return i.severity==='error';}),complete:complete,issues:issues};
}
function scoringPreviewTemplate(payload) {
  assertWebOperator_();const t=scoringNormalizeTemplates_([payload.template])[0],scores=payload.scores || [],outputs=['','',''],matched=[false,false,false];
  const validation=scoringValidateTemplateTable_({templates:[t]});if(!validation.ok)throw new Error(validation.issues.filter(function(i){return i.severity==='error';})[0].message);
  scores.forEach(function(v,i){if(v!=null && String(v).trim()!=='' && (!/^-?(?:\d+(?:\.\d*)?|\.\d+)$/.test(String(v).trim()) || i>4 || !t.domains[i].includes(Number(v))))throw new Error('枠'+(i+1)+'は設定した点数から選んでください。');});
  t.rules.forEach(function(r){if(!r.enabled || !r.conditions.every(function(c,i){return scoringConditionMatches_(scoringParseCondition_(c),scores[i]);}))return;
    r.outputs.forEach(function(v,i){if(!matched[i] && v!==''){outputs[i]=scoringTemplateText_(v,scores);matched[i]=true;}});
  });
  return {outputs:outputs};
}
function scoringTableTemplateList_() {
  return {ok:true,headers:['templateId','name','enabled'],templates:scoringReadTemplateTable_().templates.map(function(t){return {templateId:t.name,name:t.name,enabled:1};})};
}
function scoringTableRules_(params) {
  const name=String(params && params.templateId || ''),t=scoringReadTemplateTable_().templates.find(function(t){return t.name===name;});
  if(!t)throw new Error('採点テンプレが見つかりません。');
  const check=scoringValidateTemplateTable_({templates:[t]});if(!check.ok)throw new Error(check.issues.filter(function(i){return i.severity==='error';})[0].message);
  const rules=[];
  t.rules.forEach(function(r,index){if(!r.enabled)return;const conditions=r.conditions.map(scoringParseCondition_);
    r.outputs.forEach(function(message,o){if(message==='')return;rules.push({target:'出力'+(o+1),whenExpr:conditions.map(function(c,i){return c.op==='any'?'':'score'+(i+1)+(c.op==='='?'==':c.op)+c.value;}).filter(Boolean).join(' && '),
      whenLabel:r.conditions.map(function(c,i){return c?'枠'+(i+1)+' '+c:'';}).filter(Boolean).join('・') || 'すべて',message:message,enabled:1,visible:r.visible?1:0,priority:index+1,_rowNumber:r.sheetRow,_tableConditions:conditions,_tableDomains:t.domains});});
  });
  return {ok:true,templateId:name,tableFormat:true,rules:rules,headers:['target','whenExpr','whenLabel','message','enabled','visible','priority'],skipped:0};
}

function scoringValidateConfigInput_(cfg) {
  if(!cfg || typeof cfg!=='object' || Array.isArray(cfg))throw new Error('採点設定が不正です。');
  ['headerRow','startRow'].forEach(function(key){if(cfg[key]!=null && (!Number.isInteger(cfg[key]) || cfg[key]<1))throw new Error(key+' は1以上の整数で指定してください。');});
  if(cfg.lessonDateHeader!=null && typeof cfg.lessonDateHeader!=='string')throw new Error('授業日の列はヘッダ名で指定してください。');
  const cols=[cfg.nameCol,cfg.commentCol].concat(Array.isArray(cfg.scoreCols)?cfg.scoreCols:[],Array.isArray(cfg.displayCols)?cfg.displayCols:[],(Array.isArray(cfg.ruleOutputSlots)?cfg.ruleOutputSlots:[]).map(function(s){return s.col;}));
  cols.forEach(function(col){if(col!=null && !scoringLegacyIsValidColA1OrEmpty_(col))throw new Error('列は A、AA のような列記号で指定してください。');});
}
function scoringWriteRevision_(cfg) {return scoringDigest_([cfg.headerRow,cfg.startRow,cfg.scoreCols,cfg.commentCol,(cfg.ruleOutputSlots || []).map(function(s){return s.col;})]);}
function scoringSheetLiteral_(value) {return typeof value==='string' && value.charAt(0)==='=' ? "'"+value : value;}
function scoringValidateResolvedConfig_(cfg,sheet,writing,headerValues) {
  const width=sheet.getLastColumn(),headers=headerValues || (width ? sheet.getRange(cfg.headerRow || 1,1,1,width).getDisplayValues()[0] : []);
  scoringApplyLayoutBindings_(cfg,scoringManagementLayout_(sheet,headers));
  const meta=scoringLegacyBuildHeaderMeta_(headers);
  [cfg.nameHeader,cfg.commentHeader].concat(cfg.displayHeaders || [],cfg.scoreHeaders || [],(cfg.ruleOutputSlots || []).map(function(s){return s.header;})).filter(Boolean).forEach(function(header){
    if(!scoringLegacyResolveColIndexByHeaderLabel_(header,meta.map))throw new Error('設定したヘッダがありません: '+header+'。列設定を確認してください。');
  });
  scoringLegacyApplyHeaderResolutionToConfig_(cfg,meta);
  scoringLegacyAssertUniqueScoreCols_(cfg.scoreCols);
  const columns=(cfg.scoreCols || []).concat([cfg.commentCol],(cfg.ruleOutputSlots || []).map(function(s){return s.col;})).filter(Boolean);
  if(!writing)return;
  const app=getConfig_(),protectedHeaders=[app.emailHeader,app.studentNameHeader,app.formStatusHeader,SCORING_MANAGEMENT_HEADER,scoringDateHeader_(cfg,headers),'タイムスタンプ','Timestamp'].filter(Boolean);
  const dateIndex=scoringLegacyResolveColIndexByHeaderLabel_(scoringDateHeader_(cfg,headers),meta.map);
  const scoreCols=new Set(cfg.scoreCols.filter(Boolean)),outputCols=new Set();
  columns.forEach(function(col){
    const index=scoringLegacyColA1ToIndex_(col);
    if(index>sheet.getMaxColumns())throw new Error('採点列がシートの範囲外です: '+col);
    if(index===dateIndex || protectedHeaders.indexOf(headers[index-1])>=0)throw new Error('回答の識別・送信状態の列へ採点値は保存できません: '+headers[index-1]);
  });
  (cfg.ruleOutputSlots || []).forEach(function(slot){if(!slot.col)return;if(scoreCols.has(slot.col) || outputCols.has(slot.col))throw new Error('講評の出力列が重複しています。');outputCols.add(slot.col);});
  if(cfg.commentCol && scoreCols.has(cfg.commentCol))throw new Error('講評列と採点列が重複しています。');
}

/** 旧内部シートは読み取りと移行にだけ使い、新しい記録は回答の管理列へ保存する。 */
function scoringReadRecords_(name) {
  const sheet=getAppSpreadsheet_().getSheetByName(name), records=new Map();
  if (!sheet) return {sheet:null,records:records};
  const values=sheet.getDataRange().getDisplayValues();
  if (![JSON.stringify(['key','JSON']),JSON.stringify(['識別子','内容(JSON)'])].includes(JSON.stringify(values[0]))) throw new Error(name+' の構成が異なります。記録を確認してください。');
  values.slice(1).forEach(function(row,i) {
    if (!row[0] || records.has(row[0])) throw new Error(name+' の記録が重複または破損しています。');
    let data;try{data=JSON.parse(row[1]);}catch(error){throw new Error(name+' の記録を読み取れません。');}
    if (!data || typeof data!=='object') throw new Error(name+' の記録形式が不正です。');
    records.set(row[0],{row:i+2,data:data});
  });
  return {sheet:sheet,records:records,legacyHeaders:values[0][0]==='key'};
}
/** 管理セルはシステム専用。壊れた記録を空欄扱いにして返却しない。 */
function scoringParseManagement_(raw) {
  if(raw==='' || raw==null)return {v:1};
  let data;try{data=JSON.parse(String(raw));}catch(error){throw new Error('管理列のJSONを読み取れません。既存データを確認してください。');}
  if(!data || data.v!==1 || Array.isArray(data) || Object.keys(data).some(function(k){return !['v','save','duplicate'].includes(k);}) ||
    ('save' in data && (!data.save || typeof data.save!=='object' || Array.isArray(data.save) || !['pending','complete'].includes(data.save.state))) ||
    ('duplicate' in data && (!data.duplicate || typeof data.duplicate!=='object' || Array.isArray(data.duplicate) || typeof data.duplicate.group!=='string' || !data.duplicate.group || typeof data.duplicate.fingerprint!=='string' || !data.duplicate.fingerprint ||
      typeof data.duplicate.revision!=='string' || !data.duplicate.revision || !['adopted','excluded','pending'].includes(data.duplicate.choice)))) {
    throw new Error('管理列の記録形式が不正です。既存データを確認してください。');
  }
  return data;
}

function scoringManagementColumn_(sheet,create,statusHeader) {
  const width=sheet.getLastColumn(),headers=width?sheet.getRange(1,1,1,width).getDisplayValues()[0]:[];
  if(headers.filter(function(h){return h===SCORING_MANAGEMENT_HEADER;}).length>1)throw new Error('管理列が重複しています。');
  let column=headers.indexOf(SCORING_MANAGEMENT_HEADER)+1;
  if(!column && create) {
    column=width+1;
    if(sheet.getMaxColumns()<column)sheet.insertColumnsAfter(sheet.getMaxColumns(),column-sheet.getMaxColumns());
    sheet.getRange(1,column).setValue(SCORING_MANAGEMENT_HEADER);
  }
  if(column && create)column=scoringArrangeManagement_(sheet,column,statusHeader);
  const values=column && sheet.getLastRow()>1?sheet.getRange(2,column,sheet.getLastRow()-1,1).getValues():[];
  return {sheet:sheet,column:column,values:values};
}

/** 列移動前の採点列名と照合順を管理ヘッダのメモに残す。追加のシートは使わない。 */
function scoringManagementLayout_(sheet,headers) {
  const index=headers.indexOf(SCORING_MANAGEMENT_HEADER);if(index<0)return null;
  const note=sheet.getRange(1,index+1).getNote();if(!note)return null;
  let data;try{data=JSON.parse(note);}catch(error){throw new Error('管理列ヘッダのメモを確認してください。');}
  if(!data || data.type!=='turret-management-layout' || data.v!==1 || !Array.isArray(data.headers) || data.headers.some(function(h){return typeof h!=='string';}) ||
    !data.bindings || typeof data.bindings!=='object' || Array.isArray(data.bindings) || Object.keys(data.bindings).some(function(k){return !['nameHeader','displayHeaders','scoreHeaders','commentHeader','ruleOutputSlots'].includes(k);}) ||
    typeof data.configKey!=='string' || !data.configKey || typeof data.writeRevision!=='string' || !data.writeRevision)throw new Error('管理列ヘッダのメモの形式が不正です。');
  return data;
}
function scoringLayoutConfigKey_(cfg) {
  return scoringDigest_([cfg.headerRow,cfg.startRow,cfg.nameCol,cfg.nameHeader,cfg.displayCols,cfg.displayHeaders,cfg.scoreCols,cfg.scoreHeaders,cfg.commentCol,cfg.commentHeader,cfg.ruleOutputSlots]);
}
function scoringApplyLayoutBindings_(cfg,layout) {
  if(layout && layout.configKey===scoringLayoutConfigKey_(cfg))Object.assign(cfg,JSON.parse(JSON.stringify(layout.bindings)));
}
function scoringAcceptLayoutRevision_(sheet,revision) {
  const headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0],layout=scoringManagementLayout_(sheet,headers);
  return !!(layout && layout.configKey===scoringLayoutConfigKey_(scoringLegacyApiGetConfig_()) && layout.writeRevision===revision);
}
function scoringArrangeManagement_(sheet,column,statusHeader) {
  let headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0];
  const status=statusHeader || getConfig_().formStatusHeader,statusIndex=headers.indexOf(status);
  if(!status || status===SCORING_MANAGEMENT_HEADER || statusIndex<0 || headers.filter(function(h){return h===status;}).length!==1)throw new Error('送信状態列を一意に確認できません。管理列は送信状態と分けてください。');
  if(headers[headers.length-2]===SCORING_MANAGEMENT_HEADER && headers[headers.length-1]===status)return column;
  if(!scoringManagementLayout_(sheet,headers)) {
    const cfg=scoringLegacyApiGetConfig_(),key=scoringLayoutConfigKey_(cfg);
    scoringLegacyApplyHeaderResolutionToConfig_(cfg,scoringLegacyBuildHeaderMeta_(headers));
    const bindings={};['nameHeader','displayHeaders','scoreHeaders','commentHeader','ruleOutputSlots'].forEach(function(k){bindings[k]=cfg[k];});
    const note=JSON.stringify({type:'turret-management-layout',v:1,headers:headers.filter(function(h){return h!==SCORING_MANAGEMENT_HEADER;}),configKey:key,bindings:bindings,writeRevision:scoringWriteRevision_(cfg)});
    sheet.getRange(1,column).setNote(note);SpreadsheetApp.flush();
    if(sheet.getRange(1,column).getNote()!==note)throw new Error('管理列の移行前情報を保存確認できません。');
  }
  const width=headers.length;
  if(column!==width)sheet.moveColumns(sheet.getRange(1,column,sheet.getMaxRows(),1),width+1);
  headers=sheet.getRange(1,1,1,width).getDisplayValues()[0];
  sheet.moveColumns(sheet.getRange(1,headers.indexOf(status)+1,sheet.getMaxRows(),1),width+1);
  SpreadsheetApp.flush();
  headers=sheet.getRange(1,1,1,width).getDisplayValues()[0];
  if(headers[width-2]!==SCORING_MANAGEMENT_HEADER || headers[width-1]!==status)throw new Error('管理列と送信状態の配置を確認できません。再実行してください。');
  return width-1;
}
function scoringFingerprintValues_(row,headers,ignored,layout) {
  // 空見出しと「回答 (2)」などの実際の見出しも区別し、移行前と同じ全列を照合する。
  function keys(values) {
    const counts=new Map();
    return values.map(function(h){const n=(counts.get(h)||0)+1;counts.set(h,n);return h===SCORING_MANAGEMENT_HEADER?null:JSON.stringify([h,n]);});
  }
  const current=keys(headers),positions=new Map();
  current.forEach(function(key,i){if(key!==null)positions.set(key,i);});
  const order=layout?keys(layout.headers).filter(function(key){return key!==null;}):[];
  current.forEach(function(key){if(key!==null && !order.includes(key))order.push(key);});
  return order.map(function(key){
    const index=positions.get(key);
    return index===undefined?null:ignored.has(index)?'':row[index];
  });
}

/** 同じセルの採用判断と保存状態をそれぞれ保持して更新する。呼び出し元で共通ロックを取る。 */
function scoringPatchManagement_(column,patches) {
  const blocks=patches.map(function(item){
    const data=scoringParseManagement_((column.values[item.row-2] || [''])[0]);
    const json=JSON.stringify(Object.assign(data,item.patch));
    scoringParseManagement_(json);
    return {row:item.row,col:column.column,values:[json]};
  });
  scoringWriteBlocks_(column.sheet,blocks);
  SpreadsheetApp.flush();
  // 書き込みの読み戻しが済むまで、旧シート削除や採点保存の完了を宣言しない。
  const values=column.sheet.getRange(2,column.column,Math.max(1,column.sheet.getLastRow()-1),1).getValues();
  blocks.forEach(function(block){if(!values[block.row-2] || values[block.row-2][0]!==block.values[0])throw new Error('管理列の保存を確認できません。再実行してください。');});
  column.values=values;
}

function scoringWriteDecision_(group,accepted,read,revision) {
  const byTarget=new Map();
  group.responses.forEach(function(item){
    if(!byTarget.has(item.target.key))byTarget.set(item.target.key,[]);
    byTarget.get(item.target.key).push({row:item.rowNumber,patch:{duplicate:{group:group.id,fingerprint:group.fingerprint,
      revision:revision,choice:accepted===null?'pending':accepted.includes(item.key)?'adopted':'excluded'}}});
  });
  byTarget.forEach(function(patches,key){scoringPatchManagement_(scoringManagementColumn_(read.snapshots.get(key).sheet,true),patches);});
}

/** 旧記録は全件の移行先と読み戻しを確認してから削除する。移行途中は旧記録も読み続ける。 */
function scoringMigrateLegacyRecords_() {
  const saves=scoringReadRecords_('_scoring_saves'),decisions=scoringReadRecords_('_scoring_decisions');
  if(!saves.sheet && !decisions.sheet)return;
  const read=scoringReadContext_(),index=scoringBuildReviewIndex_(null,{read:read,summary:true}),byTarget=new Map();
  saves.records.forEach(function(record,key){
    let tuple;try{tuple=JSON.parse(key);}catch(error){throw new Error('旧採点保存記録の回答を特定できません。旧シートを保持します。');}
    const target=index.targets.find(function(t){return t.spreadsheetId===tuple[0] && Number(t.sheetId)===tuple[1];});
    const snapshot=target && scoringSnapshot_(read,target),row=tuple[2];
    if(!snapshot || !Number.isInteger(row) || row<2 || row>snapshot.lastRow || !['pending','complete'].includes(record.data.state))throw new Error('旧採点保存記録の移行先を確認できません。旧シートを保持します。');
    if(!byTarget.has(target.key))byTarget.set(target.key,[]);
    byTarget.get(target.key).push({row:row,patch:{save:record.data}});
  });
  decisions.records.forEach(function(record,id){
    if(!index.groups.has(id))throw new Error('旧重複判断の移行先を確認できません。旧シートを保持します。');
  });
  byTarget.forEach(function(patches,key){
    const column=scoringManagementColumn_(read.snapshots.get(key).sheet,true);
    // 新方式で再保存済みの状態を、古い未完了記録へ戻さない。
    patches=patches.filter(function(p){return !scoringParseManagement_((column.values[p.row-2] || [''])[0]).save;});
    if(patches.length)scoringPatchManagement_(column,patches);
  });
  const migratedColumns=new Map();
  decisions.records.forEach(function(record,id){
    const group=index.groups.get(id),decision=record.data,revision='legacy:'+scoringDigest_(decision);
    if(!Array.isArray(decision.accepted) && decision.accepted!==null)throw new Error('旧重複判断の形式を確認してください。');
    // 部分移行だけなら同じ世代で再開。新しい判断が始まっていれば古い判断は上書きしない。
    const columns=group.responses.map(function(item){
      if(!migratedColumns.has(item.target.key))migratedColumns.set(item.target.key,scoringManagementColumn_(read.snapshots.get(item.target.key).sheet,false));
      return scoringParseManagement_((migratedColumns.get(item.target.key).values[item.rowNumber-2] || [''])[0]).duplicate;
    });
    if(columns.some(function(d){return d && d.revision!==revision;}))return;
    scoringWriteDecision_(Object.assign({},group,{fingerprint:decision.fingerprint}),decision.accepted,read,revision);
  });
  // 削除の一部が失敗しても、管理列の新しい状態を優先して再開できる。
  if(saves.sheet)getAppSpreadsheet_().deleteSheet(saves.sheet);
  if(decisions.sheet)getAppSpreadsheet_().deleteSheet(decisions.sheet);
}
function scoringResponseKey_(target,row) {return JSON.stringify([target.spreadsheetId,Number(target.sheetId),Number(row)]);}
function scoringDigest_(value) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,JSON.stringify(value),Utilities.Charset.UTF_8).map(function(b){return ('0'+((b+256)%256).toString(16)).slice(-2);}).join('');
}
function scoringLessonDate_(value,timezone) {
  if (Object.prototype.toString.call(value)==='[object Date]') {
    if (!isFinite(value.getTime())) return '';
    return Utilities.formatDate(value,timezone,'yyyy-MM-dd');
  }
  const match=String(value == null ? '' : value).trim().match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
  if (!match) return '';
  const y=Number(match[1]),m=Number(match[2]),d=Number(match[3]),date=new Date(Date.UTC(y,m-1,d));
  if (date.getUTCFullYear()!==y || date.getUTCMonth()!==m-1 || date.getUTCDate()!==d) return '';
  return match[1]+'-'+('0'+m).slice(-2)+'-'+('0'+d).slice(-2);
}
function scoringDateHeader_(cfg,headers) {
  if (cfg.lessonDateHeader === '') return '';
  if (cfg.lessonDateHeader) return cfg.lessonDateHeader;
  return headers.indexOf('授業の日付を入力してください')>=0 ? '授業の日付を入力してください' : '';
}
function scoringBuildReviewIndex_(books,options) {
  options=options || {};
  const read=options.read || scoringReadContext_(books),summary=!!options.summary;
  const app=read.app,cfg=scoringLegacyApiGetConfig_(),operation=read.operation;
  const targets=getScoringTargets_(null,read),groups=new Map(),responses=new Map(),bySource=new Map(),enabledTargets=new Set(),ignoredByTarget=new Map();
  const decisions=scoringReadRecords_('_scoring_decisions'),saves=scoringReadRecords_('_scoring_saves');
  const mappings=readMappingEntries_(true),coursesByEmail=new Map(),returnStates=new Map();
  const students=operation.getSheetByName(STUDENT_SHEET_NAME);
  if (students && students.getLastRow()>1) {
    const values=students.getDataRange().getDisplayValues(),h=createHeaderMap_(values[0]);
    values.slice(1).forEach(function(row){const email=String(row[h['メールアドレス']] || '').trim().toLowerCase(),course=String(row[h['コースID']] || '').trim();if(!email || !course)return;const bucket=coursesByEmail.get(email)||new Set();bucket.add(course);coursesByEmail.set(email,bucket);});
  }
  const sent=operation.getSheetByName(SEND_SHEET_NAME);
  if (sent && sent.getLastRow()>1) {
    const values=sent.getDataRange().getDisplayValues(),h=createHeaderMap_(values[0]);
    values.slice(1).forEach(function(row){returnStates.set(makeResponseKey_(row[h['元SS_ID']],row[h['元シート名']],row[h['元行番号']]),String(row[h['送信状態']] || '送信待ち'));});
  }
  const posted=new Set(read.managed.filter(function(r){return r.kind==='announcement' && r.postId;}).map(function(r){return r.sourceKey;}));
  targets.forEach(function(target){
    const snapshot=scoringSnapshot_(read,target),sheet=snapshot.sheet,book=scoringBook_(read,target.spreadsheetId);
    const values=snapshot.values,headers=(values[0] || []).map(String),timezone=book.getSpreadsheetTimeZone();
    snapshot.managementLayout=scoringManagementLayout_(sheet,headers);
    const display=summary ? null : (snapshot.display || (snapshot.display=sheet.getDataRange().getDisplayValues()));
    const dateHeader=scoringDateHeader_(cfg,headers),dateColumn=(scoringLegacyResolveColIndexByHeaderLabel_(dateHeader,scoringLegacyBuildHeaderMeta_(headers).map) || 0)-1,emailColumn=headers.indexOf(app.emailHeader),statusColumn=headers.indexOf(app.formStatusHeader);
    const nameColumn=headers.indexOf(app.studentNameHeader),timeColumn=headers.indexOf('タイムスタンプ'),englishTimeColumn=headers.indexOf('Timestamp');
    if (dateHeader) enabledTargets.add(target.key);
    const resolved=JSON.parse(JSON.stringify(cfg));
    resolved.ruleOutputSlots=scoringLegacyNormalizeRuleOutputSlots_(resolved.ruleOutputSlots);
    scoringApplyLayoutBindings_(resolved,snapshot.managementLayout);
    scoringLegacyApplyHeaderResolutionToConfig_(resolved,scoringLegacyBuildHeaderMeta_(values[(cfg.headerRow || 1)-1] || []));
    const outputCols=(resolved.scoreCols || []).concat([resolved.commentCol],(resolved.ruleOutputSlots || []).map(function(s){return s.col;})).filter(Boolean).map(function(col){return scoringLegacyColA1ToIndex_(col)-1;});
    const managementColumn=headers.indexOf(SCORING_MANAGEMENT_HEADER);
    if(headers.filter(function(h){return h===SCORING_MANAGEMENT_HEADER;}).length>1)throw new Error('管理列が重複しています。');
    const ignored=new Set(outputCols.concat([statusColumn]));
    ignoredByTarget.set(target.key,ignored);
    const mapping=mappings.get(makeMappingLookupKey_(target.spreadsheetId,target.sheetName));
    for(let i=1;i<values.length;i++) {
      const row=values[i],email=String(row[emailColumn] || '').trim().toLowerCase();if(!email)continue;
      const candidates=coursesByEmail.get(email),course=String(mapping && mapping.courseId || target.courseId || (candidates && candidates.size===1 ? Array.from(candidates)[0] : ''));
      const date=dateHeader ? scoringLessonDate_(row[dateColumn],timezone) : '';
      const key=scoringResponseKey_(target,i+1),sourceKey=makeResponseKey_(target.spreadsheetId,target.sheetName,i+1);
      const state=returnStates.get(sourceKey) || String(row[statusColumn] || '未返却');
      const item={key:key,sourceKey:sourceKey,target:target,rowNumber:i+1,email:email,courseId:course,date:date,
        name:String(row[nameColumn] || ''),submittedAt:display ? display[i][timeColumn] || display[i][englishTimeColumn] || '' : '',
        returned:state==='済' || posted.has(sourceKey),returnState:state,
        state:!dateHeader ? 'disabled' : !date ? 'invalid-date' : !course ? 'invalid-class' : 'single',eligible:true};
      if(!summary) {
        item.answers=headers.map(function(header,col){return {header:header,value:display[i][col] || ''};}).filter(function(a){return a.header!==SCORING_MANAGEMENT_HEADER;});
        item.contentFingerprint=scoringDigest_([key,scoringFingerprintValues_(row,headers,ignored,snapshot.managementLayout)]);
      }
      const management=scoringParseManagement_(row[managementColumn]);
      const save=management.save || (saves.records.get(key) || {}).data;item.saveFailed=!!(save && save.state!=='complete');
      item.managementDecision=management.duplicate;
      if(!target.returnEnabled || item.saveFailed || item.state==='invalid-date' || item.state==='invalid-class')item.eligible=false;
      responses.set(key,item);bySource.set(sourceKey,item);
      if(dateHeader && date && course){const tuple=JSON.stringify([email,course,date]);let group=groups.get(tuple);if(!group){group={email:email,courseId:course,date:date,responses:[]};groups.set(tuple,group);}group.responses.push(item);}
    }
  });
  Array.from(groups.entries()).forEach(function(entry){
    const tuple=entry[0],group=entry[1];groups.delete(tuple);
    if(group.responses.length<2)return;
    const id=scoringDigest_([group.email,group.courseId,group.date]);group.id=id;groups.set(id,group);
    group.responses.sort(function(a,b){return a.key.localeCompare(b.key);});
    if(summary)group.responses.forEach(function(item){
      const row=read.snapshots.get(item.target.key).values[item.rowNumber-1],ignored=ignoredByTarget.get(item.target.key);
      const snapshot=read.snapshots.get(item.target.key);
      item.contentFingerprint=scoringDigest_([item.key,scoringFingerprintValues_(row,snapshot.values[0],ignored,snapshot.managementLayout)]);
    });
    group.fingerprint=scoringDigest_(group.responses.map(function(r){return [r.key,r.contentFingerprint];}));
    const saved=decisions.records.get(id);let decision=saved && saved.data;
    const entries=group.responses.map(function(r){return r.managementDecision;}),first=entries.find(Boolean);
    if(first) {
      const consistent=entries.every(function(d){return d && d.group===id && d.fingerprint===group.fingerprint && d.revision===first.revision && d.choice!=='pending';});
      const accepted=group.responses.filter(function(r){return r.managementDecision && r.managementDecision.choice==='adopted';}).map(function(r){return r.key;});
      decision=consistent && accepted.length ? {fingerprint:group.fingerprint,accepted:accepted} : null;
    }
    const valid=decision && decision.fingerprint===group.fingerprint && Array.isArray(decision.accepted);
    group.accepted=valid ? decision.accepted : null;
    group.responses.forEach(function(item){item.groupId=id;item.duplicateCount=group.responses.length;item.state=!valid ? 'pending' : decision.accepted.indexOf(item.key)>=0 ? 'adopted' : 'excluded';item.eligible=item.target.returnEnabled && !item.saveFailed && item.state==='adopted';});
  });
  return {targets:targets,responses:responses,bySource:bySource,groups:groups,enabledTargets:enabledTargets,read:read};
}
function scoringReviewForTarget_(index,target,summary) {
  const rows={};index.responses.forEach(function(item){if(item.target.key===target.key)rows[item.rowNumber]=summary ? {
    key:item.key,rowNumber:item.rowNumber,state:item.state,groupId:item.groupId || '',duplicateCount:item.duplicateCount || 0,
    returned:item.returned,returnState:item.returnState,saveFailed:item.saveFailed,eligible:item.eligible
  } : item;});
  if(summary)return {enabled:index.enabledTargets.has(target.key),rows:rows,groups:[],detailsDeferred:true};
  return {enabled:index.enabledTargets.has(target.key),rows:rows,groups:Array.from(index.groups.values()).filter(function(group){return group.responses.some(function(r){return r.target.key===target.key;});})};
}
function scoringGetReview(params) {
  assertWebOperator_();const read=scoringReadContext_(),context=resolveScoringTarget_(params && params.target,read),summary=!!(params && params.reviewMode==='summary');
  return scoringReviewForTarget_(scoringBuildReviewIndex_(null,{read:read,summary:summary}),context.target,summary);
}
function scoringGetComparison(params) {
  assertWebOperator_();const read=scoringReadContext_(),context=resolveScoringTarget_(params && params.target,read);
  const index=scoringBuildReviewIndex_(null,{read:read,summary:true}),group=index.groups.get(params.groupId);
  if(!group || !group.responses.some(function(r){return r.target.key===context.target.key;}))throw new Error('重複回答が変更されました。採点画面を読み直してください。');
  const spans=new Map();
  group.responses.forEach(function(item){
    const span=spans.get(item.target.key) || {first:item.rowNumber,last:item.rowNumber};
    span.first=Math.min(span.first,item.rowNumber);span.last=Math.max(span.last,item.rowNumber);spans.set(item.target.key,span);
  });
  spans.forEach(function(span,key){const snapshot=read.snapshots.get(key);span.display=snapshot.sheet.getRange(span.first,1,span.last-span.first+1,snapshot.lastCol).getDisplayValues();});
  group.responses.forEach(function(item){
    const snapshot=read.snapshots.get(item.target.key),headers=(snapshot.values[0] || []).map(String),span=spans.get(item.target.key);
    const display=span.display[item.rowNumber-span.first];
    item.answers=headers.map(function(header,col){return {header:header,value:display[col] || ''};}).filter(function(a){return a.header!==SCORING_MANAGEMENT_HEADER;});
    item.submittedAt=display[headers.indexOf('タイムスタンプ')] || display[headers.indexOf('Timestamp')] || '';
  });
  return group;
}
function scoringChooseResponses(payload) {
  assertWebOperator_();
  return withAppLock_(function(){
    const read=scoringReadContext_(),context=resolveScoringTarget_(payload && payload.target,read),index=scoringBuildReviewIndex_(null,{read:read,summary:true}),group=index.groups.get(payload.groupId);
    if(!group || !group.responses.some(function(r){return r.target.key===context.target.key;}) || group.fingerprint!==payload.fingerprint)throw new Error('回答が更新されました。比較画面を読み直してください。');
    if(payload.accepted!==null && (!Array.isArray(payload.accepted) || !payload.accepted.length || new Set(payload.accepted).size!==payload.accepted.length || payload.accepted.some(function(key){return !group.responses.some(function(r){return r.key===key;});})))throw new Error('採用する回答を選択してください。');
    scoringWriteDecision_(group,payload.accepted,read,Utilities.getUuid());
    if(payload.reviewMode==='summary') {
      group.responses.forEach(function(item){item.state=payload.accepted===null?'pending':payload.accepted.indexOf(item.key)>=0?'adopted':'excluded';item.eligible=item.target.returnEnabled && !item.saveFailed && item.state==='adopted';});
      return {ok:true,review:scoringReviewForTarget_(index,context.target,true)};
    }
    return {ok:true};
  });
}
function scoringBuildReturnGuard_(cache) {
  const books=cache ? cache.spreadsheets : new Map();
  if(!PropertiesService.getScriptProperties().getProperty('TURRET_SCORING_CONFIG')) {
    const app=getConfig_();let hasDefaultDate=false,hasManagement=false;
    (app.formSources || []).forEach(function(source){
      if(!books.has(source.id))books.set(source.id,SpreadsheetApp.openById(source.id));
      books.get(source.id).getSheets().forEach(function(sheet){
        if(isManagedInternalSheet_(sheet.getName(),source.id) || !sheetMatchesPrefixes_(sheet.getName(),app.formSheetNamePrefix) || !sheet.getLastColumn())return;
        const headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0];
        if(cache)cache.targets.set(JSON.stringify([source.id,sheet.getName(),app.formStatusHeader]),{sheet:sheet,statusColumn:headers.indexOf(app.formStatusHeader)+1});
        if(headers.indexOf('授業の日付を入力してください')>=0)hasDefaultDate=true;
        if(headers.includes(SCORING_MANAGEMENT_HEADER))hasManagement=true;
      });
    });
    // 日付のない既存運用は従来通り。標準授業日は設定保存前から画面と同じ判定を行う。
    if(!hasDefaultDate && !hasManagement && !getAppSpreadsheet_().getSheetByName('_scoring_saves'))return null;
  }
  return scoringBuildReviewIndex_(books,{summary:true});
}
function scoringMayReturn_(index,sourceKey) {
  if(!index)return true;
  const item=index.bySource.get(sourceKey);
  return !!(item && item.eligible);
}

function scoringPasteToRuleCell(text) {
  assertWebOperator_();
  return withAppLock_(function(){
    const ss=SpreadsheetApp.getActiveSpreadsheet();
    if(!ss)throw new Error('運用スプレッドシートの式作成サイドバーから実行してください。');
    const sheet=ss.getActiveSheet(),cell=sheet && sheet.getActiveCell();
    if(!cell || sheet.getName()!=='_rules' || cell.getRow()<2)throw new Error('_rules の条件式または本文のセルを選択してください。');
    const header=String(sheet.getRange(1,cell.getColumn()).getValue()).trim().toLowerCase();
    if(['whenexpr','when','expr','condition','if','message','text','output','result','comment'].indexOf(header)<0)throw new Error('条件式または本文の列を選択してください。');
    const value=String(text || '');cell.setValue(value.charAt(0)==='=' ? "'"+value : value);
    sheet.setActiveRange(sheet.getRange(cell.getRow()+1,cell.getColumn()));
    return {ok:true,message:'貼り付けました（1行下へ移動）。'};
  });
}

/** コンテナに紐付くコピーは自身を使う。Web・単独トリガーでは明示された対象を使う。 */
function getAppSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const id = PropertiesService.getScriptProperties().getProperty('TURRET_WEB_TARGET');
  if (!id) throw new Error('運用スプレッドシートが未接続です。Web画面で新規作成するか、スプシのメニューから接続してください。');
  return SpreadsheetApp.openById(id);
}

/** 更新入口で共通ロックを使用する。対話ダイアログは呼び出し元で解放後に表示する。 */
function withAppLock_(operation) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) {
    throw new Error('別の処理を実行中です。終了してからもう一度実行してください。');
  }
  try {
    return operation();
  } finally {
    try {
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
  }
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('自動送信システム')
    .addItem('管理画面を開く', 'connectWebConsole')
    .addItem('採点画面を開く', 'connectScoringConsole')
    .addItem('採点テンプレ作成画面へ', 'openScoringRuleEditor')
    .addToUi();
}

function initializeSheets() {
  withAppLock_(initializeSheetsUnlocked_);
  safeAlert_('シートを作成しました（クラス一覧・生徒一覧・採点テンプレ・対応表・送信シート・エラー・フォーム管理）');
}

function initializeSheetsUnlocked_() {
  ensureSheet_('クラス一覧', ['クラス名', 'コースID', '同期対象(1)']);
  ensureStudentSheetForSync_();
  migrateConfigStorageUnlocked_();
  scoringEnsureTemplateTable_();
  ensureSheet_(MAPPING_SHEET_NAME, MAPPING_SHEET_HEADERS);
  // 初期設定前も固定列で作成できる。既存の送信データは消さない。
  ensureSendSheet_(getConfig_());
  ensureErrorSheet_();
  managedSheet_(true);
  scoringMigrateLegacyRecords_();
}

function openSettingsDialog() {
  return connectWebConsole();
}

function getSettingsDialogData() {
  // 設定読み込みと保存を排他する。
  return withAppLock_(getSettingsDialogDataUnlocked_);
}

function getSettingsDialogDataUnlocked_() {
  const config = getConfig_();
  const parseError = PropertiesService.getScriptProperties()
    .getProperty('APP_CONFIG_PARSE_ERROR') === '1';

  return {
    config: config,
    revision: getConfigRevision_(config),
    configParseError: parseError,
    availableFieldTypes: [
      { value: 'text', label: 'text' },
      { value: 'date', label: 'date' },
      { value: 'score_grade', label: 'score_grade' }
    ],
    templateGuide: [
      '差し込みは {score} です。例: 評価：{score}',
      '固定トークンも同じ書き方です。例: {student_name}, {class_name}, {course_id}, {row_no}',
      '値だけの行は、値が空なら自動で非表示になります。',
      '括弧や見出しを付けた行も、対応する値が空ならその行は自動で消えます。'
    ].join('\n')
  };
}

function saveSettingsFromDialog(payload, expectedRevision) {
  return withAppLock_(function() {
    return saveSettingsFromDialogUnlocked_(payload, expectedRevision);
  });
}

function saveSettingsFromDialogUnlocked_(payload, expectedRevision) {
  const existing = getConfig_();
  // 旧クライアントのリセット確認値では、現在は削除を許可しない。
  return applyConfigDraft_(Object.assign({}, existing, payload),
    typeof expectedRevision === 'string' ? expectedRevision : getConfigRevision_(existing));
}

function getConfigRevision_(config) {
  // 旧設定の移行時に上書きされた値も含め、有効な設定から再現可能で欠落のない識別値を作る。
  return JSON.stringify(normalizeAppConfig_(config));
}

function saveSetupSection(section, payload, expectedRevision) {
  const sections = {};
  Object.defineProperty(sections, section, { value: payload, enumerable: true });
  return saveSetupSections(sections, expectedRevision);
}

/** 編集された初期設定項目をまとめて保存する。自動処理の設定は別に行う。 */
function saveSetupSections(payload, expectedRevision) {
  return withAppLock_(function() {
    const sections = {
      sources: ['formSources', 'reminderTo', 'formSheetNamePrefix'],
      fields: ['emailHeader', 'studentNameHeader', 'formStatusHeader', 'scoreSourceHeader', 'replyBodyHeader', 'fields', 'gradeScale'],
      template: ['messageTemplate']
    };
    if (!payload || typeof payload !== 'object' || Array.isArray(payload) || !Object.keys(payload).length) {
      throw new Error('保存する設定項目の形式が不正です。');
    }
    const config = getConfig_();
    Object.keys(payload).forEach(function(section) {
      const values = payload[section];
      if (!Object.prototype.hasOwnProperty.call(sections, section) || !values || typeof values !== 'object' || Array.isArray(values)) {
        throw new Error('保存する設定項目の形式が不正です: ' + section);
      }
      Object.keys(values).forEach(function(key) {
        if (sections[section].indexOf(key) < 0) throw new Error('この手順では保存できない設定です: ' + key);
        config[key] = values[key];
      });
    });
    return applyConfigDraft_(config, expectedRevision);
  });
}

/** 呼び出し元は withAppLock_ を保持する。下書き保存では評価行や送信行を変更しない。 */
function applyConfigDraft_(rawConfig, expectedRevision) {
  const existing = getConfig_();
  if (typeof expectedRevision !== 'string' || expectedRevision !== getConfigRevision_(existing)) {
    throw new Error('画面を開いた後に設定が変更されました。再読込してから保存してください。');
  }
  validateConfigDraft_(rawConfig);
  const config = normalizeAppConfig_(rawConfig);
  validateConfigSheetLayouts_(config);
  saveConfig_(config);
  return { ok: true, message: '設定を保存しました。既存データは保持しています。', config: config, revision: getConfigRevision_(config) };
}

/** 設定保存と管理フォーム準備で共用する、読み取り専用の事前確認。 */
function validateConfigSheetLayouts_(config) {
  const ss = getAppSpreadsheet_();
  [[EVAL_SHEET_NAME, getConfiguredEvalHeaders_(config)], [SEND_SHEET_NAME, getConfiguredSendHeaders_(config)]].forEach(function(entry) {
    const sheet = ss.getSheetByName(entry[0]);
    if (entry[0] === EVAL_SHEET_NAME && sheet && sheet.getLastRow() > 1) {
      const values = sheet.getDataRange().getDisplayValues(), h = createHeaderMap_(values[0]);
      if (values.slice(1).every(function(row) { return !row.some(function(v) { return String(v).trim(); }) || row[h['処理状態']] === '準備〇'; })) return;
      validateLegacyEvaluationFields_(config);
    }
    if (sheet && sheet.getLastRow() > 1 && (sheet.getLastColumn() !== entry[1].length || !hasMatchingHeaders_(sheet, entry[1]))) {
      throw new Error('「' + entry[0] + '」の既存データと列構成が一致しません。管理画面で必要なデータを退避・クリアしてから保存してください。');
    }
  });
}

/**
 * 指定された URL リストとシート名接頭辞から利用可能なヘッダー候補を収集する。
 * @param {string[]} urlList - 未正規化の URL 文字列の配列
 * @param {string} prefix - 対象シート名の接頭辞
 * @returns {{ headers: string[], warnings: string[] }}
 */
function collectAvailableSourceHeaders_(urlList, prefix) {
  const headerSet = {};
  const headers = [];
  const warnings = [];
  let matchedSheets = 0;
  const normalizedPrefixes = Array.isArray(prefix)
    ? prefix.map(function(p) { return String(p || '').trim(); }).filter(Boolean)
    : (String(prefix || '').trim() ? [String(prefix || '').trim()] : []);

  if (!Array.isArray(urlList) || urlList.length === 0) {
    warnings.push('フォーム回答スプレッドシートが未設定です。');
    return { headers: headers, warnings: warnings };
  }
  if (normalizedPrefixes.length === 0) {
    warnings.push('対象シート名の接頭辞が未設定です。設定シートの FORM_SHEET_PREFIX を入力してください。');
    return { headers: headers, warnings: warnings };
  }

  urlList.forEach(function(rawUrl) {
    let sourceId = '';
    try {
      const normalized = normalizeFormSource_(String(rawUrl || '').trim());
      if (!normalized) {
        warnings.push('URLの形式が不正です: ' + trunc_(rawUrl, 60));
        return;
      }
      sourceId = normalized.id;
    } catch (e) {
      warnings.push('URLの解析に失敗しました: ' + trunc_(rawUrl, 60));
      return;
    }

    try {
      const ss = SpreadsheetApp.openById(sourceId);
      const sheets = ss.getSheets().filter(function(sheet) {
        return !isManagedInternalSheet_(sheet.getName(), sourceId) && sheetMatchesPrefixes_(sheet.getName(), normalizedPrefixes);
      });

      if (sheets.length === 0) {
        warnings.push(
          ss.getName() + ' に対象シートがありません。' +
          (normalizedPrefixes.length > 0
            ? ' 接頭辞「' + normalizedPrefixes.join('」「') + '」を確認してください。'
            : '')
        );
        return;
      }

      sheets.forEach(function(sheet) {
        const lastCol = sheet.getLastColumn();
        if (lastCol === 0) return;
        matchedSheets++;

        sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].forEach(function(header) {
          const key = String(header || '').trim();
          if (!key || headerSet[key]) return;
          headerSet[key] = true;
          headers.push(key);
        });
      });
    } catch (error) {
      warnings.push(sourceId + ' のヘッダー取得に失敗しました: ' + trunc_(error && error.message || error, 80));
    }
  });

  if (matchedSheets === 0 && warnings.length === 0) {
    warnings.push('対象シートからヘッダー候補を取得できませんでした。');
  }

  return { headers: headers, warnings: warnings };
}

/**
 * 設定ダイアログの「ヘッダー候補を更新」ボタンから呼ばれる。
 * 設定シートから formSources と formSheetNamePrefix を読み取り、ヘッダー候補を返す（保存は行わない）。
 * @returns {{ headers: string[], warnings: string[] }}
 */
function fetchHeadersFromSources() {
  return withAppLock_(fetchHeadersFromSourcesUnlocked_);
}

function fetchHeadersFromSourcesUnlocked_() {
  const config = getConfig_();
  const urlList = (config.formSources || []).map(function(s) {
    return (s && s.url) ? s.url : String(s || '');
  });
  return collectAvailableSourceHeaders_(urlList, config.formSheetNamePrefix || '');
}

function scoreToGrade(v) {
  const n = Number(String(v == null ? '' : v).trim());
  switch (n) {
    case 5: return 'A';
    case 4: return 'B+';
    case 3: return 'B';
    case 2: return 'B-';
    case 1: return 'C';
    default: return '';
  }
}

function ensureSheet_(name, headers) {
  const ss = getAppSpreadsheet_();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  const lastCol = Math.max(sheet.getLastColumn(), headers.length, 1);
  const firstRow = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const hasHeader = firstRow.some(function(v) {
    return String(v || '').trim() !== '';
  });

  if (!hasHeader) {
    if (sheet.getLastRow() > 1) throw new Error('「' + name + '」の見出しが空です。既存データを確認してください。');
    sheet.clear();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function ensureErrorSheet_() {
  return ensureSheet_(ERROR_SHEET_NAME, [
    'タイムスタンプ',
    'シート名',
    '行番号',
    'メールアドレス',
    '送信状態',
    'エラー種別',
    'エラーメッセージ'
  ]);
}

function safeAlert_(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    Logger.log('[ALERT suppressed] ' + message);
  }
}

function trunc_(s, n) {
  const text = String(s == null ? '' : s).trim();
  return text.length > n ? text.slice(0, n) + '…' : text;
}

function pad2(n) {
  const num = Number(n);
  return (num < 10 ? '0' : '') + num;
}

function toYmdString(s) {
  const str = String(s == null ? '' : s).trim();
  if (!str) return '';

  let m = str.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (m) return m[1] + '/' + pad2(m[2]) + '/' + pad2(m[3]);

  m = str.match(/^\w+\s+([A-Za-z]{3})\s+(\d{1,2})\s+(\d{4})/);
  if (m) {
    const mon = {
      Jan: 1, Feb: 2, Mar: 3, Apr: 4, May: 5, Jun: 6,
      Jul: 7, Aug: 8, Sep: 9, Oct: 10, Nov: 11, Dec: 12
    }[m[1]];
    if (mon) return m[3] + '/' + pad2(mon) + '/' + pad2(m[2]);
  }

  return str;
}

function createHeaderMap_(headers) {
  const map = {};
  headers.forEach(function(header, index) {
    const key = String(header || '').trim();
    if (key) map[key] = index;
  });
  return map;
}

function makeStudentLookupKey_(courseId, email) {
  return String(courseId || '').trim() + '\t' + String(email || '').trim().toLowerCase();
}

function makeMappingLookupKey_(sourceId, sheetName) {
  return String(sourceId || '').trim() + '\t' + String(sheetName || '').trim();
}

/** 既存の元情報で照合する。元シートの物理的な並べ替えには対応しない。 */
function makeResponseKey_(sourceId, sheetName, rowNumber) {
  const id = String(sourceId || '').trim();
  const name = String(sheetName || '').trim();
  const row = Number(rowNumber);
  if (!id || !name || !Number.isInteger(row) || row < 2) return '';
  return JSON.stringify([id, name, row]);
}

function collectResponseKeys_(values, headerMap) {
  const keys = new Set();
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const key = makeResponseKey_(row[headerMap['元SS_ID']], row[headerMap['元シート名']], row[headerMap['元行番号']]);
    if (key) keys.add(key);
  }
  return keys;
}

function buildCourseIdFormulaForRow_(rowNum, classRecords) {
  const clauses = classRecords.map(function(item) {
    const className = String(item.className || '').replace(/"/g, '""');
    const courseId = String(item.courseId || '').replace(/"/g, '""');
    return 'D' + rowNum + '="' + className + '","' + courseId + '"';
  });
  return '=IF(D' + rowNum + '="","",IFS(' + clauses.join(',') + ',TRUE,""))';
}

function normalizeMultiline_(value) {
  return String(value == null ? '' : value)
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n');
}

/**
 * 文字列または配列を正規化して文字列配列を返す。
 * @param {*} value - 文字列、配列、またはその他
 * @param {string[]} [fallback=[]] - 結果が空の場合のデフォルト値
 * @returns {string[]}
 */
function normalizeStringArray_(value, fallback) {
  let arr;
  if (Array.isArray(value)) {
    arr = value.map(function(v) { return String(v || '').trim(); }).filter(Boolean);
  } else {
    const str = String(value != null ? value : '').trim();
    arr = str ? [str] : [];
  }
  return arr.length > 0 ? arr : (fallback || []);
}

/**
 * シート名がいずれかの接頭辞に前方一致するか判定する。
 * 接頭辞が空の場合は対象なしとする（意図せず全シートを取り込まない）。
 * @param {string} sheetName
 * @param {string[]} prefixes
 * @returns {boolean}
 */
function sheetMatchesPrefixes_(sheetName, prefixes) {
  const active = (prefixes || []).filter(function(p) { return !!p; });
  if (active.length === 0) return false;
  return active.some(function(p) { return sheetName.indexOf(p) === 0; });
}

function isManagedInternalSheet_(name, sourceId) {
  const kind=internalSheetKind_(name);
  if (kind==='scoring') return true;
  const ss=getAppSpreadsheet_();
  if(typeof ss.getId!=='function'||sourceId!==ss.getId())return false;
  return kind==='operation';
}

function buildDefaultConfig_() {
  return {
    reminderTo: [], formSheetNamePrefix: [], formSources: [], fields: [], gradeScale: [],
    emailHeader: '', studentNameHeader: '', formStatusHeader: '', scoreSourceHeader: '',
    replyBodyHeader: '返信本文', messageTemplate: ''
  };
}

/** 旧管理画面以前に保存された設定だけ、過去の既定値で欠落項目を補う。 */
function buildLegacyDefaultConfig_() {
  return {
    gradeScale: buildLegacyGradeScale_(),
    reminderTo: [],
    formSheetNamePrefix: ['フォームの回答'],
    emailHeader: 'メールアドレス',
    studentNameHeader: '名前',
    formStatusHeader: '送信状態',
    scoreSourceHeader: '評価',
    scoreFieldKey: 'score',
    replyBodyHeader: '返信本文',
    formSources: [],
    fields: [
      {
        key: 'lesson_date',
        sourceHeader: '授業の日付を入力してください。',
        evalHeader: '授業日付',
        sendHeader: '授業日付',
        type: 'date'
      },
      {
        key: 'submission',
        sourceHeader: '改行除去',
        evalHeader: '提出内容',
        sendHeader: '提出内容',
        type: 'text'
      },
      {
        key: 'score',
        sourceHeader: '評価',
        evalHeader: '評価',
        sendHeader: '評価',
        type: 'score_grade'
      },
      {
        key: 'reason',
        sourceHeader: '評価理由',
        evalHeader: '評価理由',
        sendHeader: '評価理由',
        type: 'text'
      },
      {
        key: 'advice',
        sourceHeader: '改善点',
        evalHeader: '改善点',
        sendHeader: '改善点',
        type: 'text'
      }
    ],
    messageTemplate: normalizeMultiline_(
      '評価：＜score＞\n\n' +
      '【理由】\n＜reason＞\n\n' +
      '【改善】\n＜advice＞\n\n' +
      '【提出内容（抜粋）】\n授業日付：＜lesson_date＞\n＜submission＞'
    )
  };
}

function normalizeFormSource_(item) {
  let raw = '';
  let courseId = '';

  if (item && typeof item === 'object') {
    raw = String(item.url || item.id || item.value || item.raw || '').trim();
    courseId = String(item.courseId || '').trim();
  } else {
    raw = String(item || '').trim();
  }
  if (!raw) return null;

  const match = raw.match(/^(.+?)\s*\|\s*([A-Za-z0-9_-]+)\s*$/);
  if (match) {
    raw = match[1].trim();
    if (!courseId) courseId = match[2].trim();
  }

  const id = extractSpreadsheetId_(raw);
  return {
    url: /^https?:\/\//i.test(raw)
      ? raw
      : ('https://docs.google.com/spreadsheets/d/' + id + '/edit'),
    id: id,
    courseId: courseId
  };
}

function formatFormSourceLine_(source) {
  const normalized = normalizeFormSource_(source);
  if (!normalized) return '';
  return normalized.url + (normalized.courseId ? ' | ' + normalized.courseId : '');
}

function normalizeFieldConfig_(field, fallbackIndex) {
  const normalized = {
    key: String(field && field.key || '').trim(),
    sourceHeader: String(field && field.sourceHeader || '').trim(),
    evalHeader: String(field && field.evalHeader || '').trim(),
    sendHeader: String(field && field.sendHeader || '').trim(),
    type: String(field && field.type || 'text').trim() || 'text'
  };

  if (!normalized.key) {
    normalized.key = 'field_' + (fallbackIndex + 1);
  }
  // 項目がない場合だけ旧式の共通対応表を使う。明示された空配列は未設定のままにする。
  if (field && Object.prototype.hasOwnProperty.call(field, 'gradeScale')) {
    normalized.gradeScale = normalizeGradeScale_(field.gradeScale);
  }

  return normalized;
}

function normalizeGradeScale_(gradeScale) {
  return (Array.isArray(gradeScale) ? gradeScale : []).map(function(entry) {
    return { from: String(entry.from).trim(), to: String(entry.to).trim() };
  });
}

function getFieldGradeScale_(field, config) {
  if (field && Object.prototype.hasOwnProperty.call(field, 'gradeScale')) {
    return Array.isArray(field.gradeScale) ? field.gradeScale : [];
  }
  return config && Array.isArray(config.gradeScale) ? config.gradeScale : [];
}

function getFieldEvalHeader_(field) {
  return String(field && (field.evalHeader || field.sourceHeader) || '').trim();
}

function getFieldSendHeader_(field) {
  return String(field && (field.sendHeader || field.sourceHeader) || '').trim();
}

function normalizeAppConfig_(rawConfig) {
  const defaults = buildDefaultConfig_();
  const config = rawConfig || {};
  const formSources = Array.isArray(config.formSources) ? config.formSources : [];
  const fields = Array.isArray(config.fields) ? config.fields : [];
  const normalizedFields = fields.map(function(field, index) { return normalizeFieldConfig_(field, index); });
  let scoreSourceHeader = String(config.scoreSourceHeader != null ? config.scoreSourceHeader : '').trim();

  if (config.scoreSourceHeader == null && config.scoreFieldKey != null) {
    const legacyField = getFieldByKey_({ fields: normalizedFields }, String(config.scoreFieldKey).trim());
    if (legacyField) scoreSourceHeader = legacyField.sourceHeader;
  }
  if (!scoreSourceHeader) {
    scoreSourceHeader = defaults.scoreSourceHeader;
  }

  return {
    reminderTo: normalizeStringArray_(config.reminderTo != null ? config.reminderTo : defaults.reminderTo),
    formSheetNamePrefix: normalizeStringArray_(config.formSheetNamePrefix != null ? config.formSheetNamePrefix : defaults.formSheetNamePrefix),
    emailHeader: String(config.emailHeader != null ? config.emailHeader : defaults.emailHeader).trim(),
    studentNameHeader: String(config.studentNameHeader != null ? config.studentNameHeader : defaults.studentNameHeader).trim(),
    formStatusHeader: String(config.formStatusHeader != null ? config.formStatusHeader : defaults.formStatusHeader).trim(),
    scoreSourceHeader: scoreSourceHeader,
    replyBodyHeader: String(config.replyBodyHeader != null ? config.replyBodyHeader : defaults.replyBodyHeader).trim(),
    messageTemplate: normalizeMultiline_(config.messageTemplate != null ? config.messageTemplate : defaults.messageTemplate),
    formSources: formSources.map(normalizeFormSource_).filter(function(item) { return !!item; }),
    fields: normalizedFields,
    gradeScale: normalizeGradeScale_(config.gradeScale)
  };
}

function validateAppConfig_(config) {
  validateConfigDraft_(config);
  // reminderTo は任意。リマインダー送信時のみチェックする。
  if (config.formSources.length === 0) {
    throw new Error('フォーム回答スプレッドシートURLを1件以上入力してください。');
  }
  if (config.formSheetNamePrefix.length === 0) {
    throw new Error('管理画面の「フォームと通知先」で対象シート名の接頭辞を入力してください。');
  }
  if (SEND_BASE_HEADERS.indexOf(config.replyBodyHeader) >= 0) {
    throw new Error('返信本文の列名が送信シートの固定列名と重複しています: ' + config.replyBodyHeader);
  }
  if (config.fields.length === 0) {
    throw new Error('取り込み項目を1件以上設定してください。');
  }
  if (!config.messageTemplate.trim()) {
    throw new Error('返信テンプレートを入力してください。');
  }
  config.fields.forEach(function(field) {
    if (field.type === 'score_grade' && !getFieldGradeScale_(field, config).length) {
      throw new Error('「' + field.sourceHeader + '」の評価変換表を1件以上設定してください。');
    }
  });
  ['emailHeader', 'studentNameHeader', 'formStatusHeader', 'scoreSourceHeader', 'replyBodyHeader'].forEach(function(key) {
    if (!config[key]) throw new Error('必要な列名が未設定です: ' + key);
  });

  const allowedTypes = { text: true, date: true, score_grade: true };
  const fieldKeySet = {};
  const sendHeaderSet = createHeaderSet_(SEND_BASE_HEADERS.concat([config.replyBodyHeader]));
  let hasScoreField = false;

  config.fields.forEach(function(field) {
    const sendHeader = getFieldSendHeader_(field);

    if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(field.key)) {
      throw new Error('項目キーは半角英数字とアンダースコアで入力してください: ' + field.key);
    }
    if (!field.sourceHeader) {
      throw new Error('フォーム側ヘッダを入力してください: ' + field.key);
    }
    if (fieldKeySet[field.key]) {
      throw new Error('項目キーが重複しています: ' + field.key);
    }
    if (!allowedTypes[field.type]) {
      throw new Error('未対応の項目タイプです: ' + field.type);
    }
    fieldKeySet[field.key] = true;

    if (sendHeader) {
      if (sendHeaderSet[sendHeader]) {
        throw new Error('送信シートの列名が重複しています: ' + sendHeader);
      }
      sendHeaderSet[sendHeader] = true;
    }
  });

  hasScoreField = !!getFieldBySourceHeader_(config, config.scoreSourceHeader);
  if (!hasScoreField) {
    throw new Error('採点済み判定に使うフォーム側ヘッダが fields に存在しません: ' + config.scoreSourceHeader);
  }
}

/** 正規化前に入力を厳密に確認する。未完成の下書き項目は許容する。 */
function validateConfigDraft_(config) {
  if (!config || typeof config !== 'object' || Array.isArray(config)) throw new Error('設定の形式が不正です。');
  ['emailHeader', 'studentNameHeader', 'formStatusHeader', 'scoreSourceHeader', 'scoreFieldKey', 'replyBodyHeader', 'messageTemplate'].forEach(function(key) {
    if (config[key] !== undefined && typeof config[key] !== 'string') throw new Error('文字列で入力してください: ' + key);
  });
  ['emailHeader','studentNameHeader','formStatusHeader','scoreSourceHeader'].forEach(function(key){if(String(config[key] || '').trim()===SCORING_MANAGEMENT_HEADER)throw new Error('「管理」はシステム専用列です。列の用途には指定できません。');});
  ['reminderTo', 'formSheetNamePrefix', 'formSources', 'fields', 'gradeScale'].forEach(function(key) {
    if (config[key] !== undefined && !Array.isArray(config[key])) throw new Error('配列で入力してください: ' + key);
  });
  if (config.gradeScale !== undefined) validateGradeScaleDraft_(config.gradeScale);
  (config.reminderTo || []).forEach(function(value) {
    if (typeof value !== 'string' || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value.trim())) throw new Error('通知先メールアドレスが不正です。');
  });
  (config.formSheetNamePrefix || []).forEach(function(value) {
    if (typeof value !== 'string' || !value.trim()) throw new Error('対象シート名の接頭辞が不正です。');
  });
  (config.formSources || []).forEach(function(value) {
    if (typeof value !== 'string') {
      if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error('参照元URLの形式が不正です。');
      ['url', 'id', 'value', 'raw', 'courseId'].forEach(function(key) {
        if (value[key] !== undefined && typeof value[key] !== 'string') throw new Error('参照元URLの形式が不正です。');
      });
    }
    if (!normalizeFormSource_(value)) throw new Error('参照元URLが空です。');
  });
  const keys = new Set(['student_name', 'class_name', 'course_id', 'row_no', 'prototype'].concat(Object.getOwnPropertyNames(Object.prototype)));
  const sendHeaders = new Set(SEND_BASE_HEADERS);
  const reply = String(config.replyBodyHeader || '').trim();
  if (reply && sendHeaders.has(reply)) throw new Error('返信本文の列名が固定列名と重複しています: ' + reply);
  if (reply) sendHeaders.add(reply);
  (config.fields || []).forEach(function(field) {
    if (!field || typeof field !== 'object' || Array.isArray(field)) throw new Error('項目の形式が不正です。');
    ['key', 'sourceHeader', 'evalHeader', 'sendHeader', 'type'].forEach(function(key) {
      if (field[key] !== undefined && typeof field[key] !== 'string') throw new Error('項目は文字列で入力してください: ' + key);
    });
    if (Object.prototype.hasOwnProperty.call(field, 'gradeScale')) validateGradeScaleDraft_(field.gradeScale);
    const key = String(field.key || '').trim();
    if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(key)) throw new Error('項目キーが不正です: ' + key);
    if (keys.has(key)) throw new Error('項目キーが重複または予約済みです: ' + key);
    keys.add(key);
    if (['text', 'date', 'score_grade'].indexOf(field.type || 'text') < 0) throw new Error('未対応の項目タイプです: ' + field.type);
    if (!String(field.sourceHeader || '').trim()) throw new Error('フォーム側ヘッダを入力してください: ' + key);
    if(field.sourceHeader.trim()===SCORING_MANAGEMENT_HEADER)throw new Error('「管理」はシステム専用列です。本文の項目には指定できません。');
    const sendHeader = getFieldSendHeader_(field);
    if (sendHeaders.has(sendHeader)) throw new Error('項目の列名が重複しています: ' + key);
    sendHeaders.add(sendHeader);
  });
}

function createHeaderSet_(headers) {
  const set = {};
  headers.forEach(function(header) {
    set[String(header || '').trim()] = true;
  });
  return set;
}

function extractSpreadsheetId_(value) {
  const text = String(value == null ? '' : value).trim();
  if (!text) {
    throw new Error('スプレッドシートURLまたはIDが空です。');
  }

  const match = text.match(/^https:\/\/docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9_-]+)(?:[/?#]|$)/);
  if (match) return match[1];

  if (/^[a-zA-Z0-9-_]{20,}$/.test(text)) {
    return text;
  }

  throw new Error('スプレッドシートURLまたはIDの形式が不正です: ' + trunc_(text, 80));
}

/** アプリ共通ロック中にだけ呼ぶ。旧タブを削除する前に保存と読み戻し確認を行う。 */
function migrateConfigStorageUnlocked_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('APP_CONFIG_STORAGE') !== 'properties-v1') {
    const config = getConfig_();
    validateConfigDraft_(config);
    saveConfig_(config);
  }
  if (props.getProperty('APP_CONFIG_SHEETS_RETIRED') === '1') return;
  // 復旧または読み戻し確認が未完了なら、旧保存先を削除しない。
  const config = getConfig_();
  validateConfigDraft_(config);
  if (props.getProperty('APP_CONFIG_PARSE_ERROR') === '1' ||
      props.getProperty('APP_CONFIG') !== props.getProperty('APP_CONFIG_BACKUP')) {
    throw new Error('内部設定の保存確認ができないため、旧設定シートを保持しています。');
  }
  const ss = getAppSpreadsheet_();
  const names = [CONFIG_SHEET_NAME, INPUT_SETTINGS_SHEET_NAME, SETTINGS_SHEET_NAME];
  const obsolete = names.map(function(name) { return ss.getSheetByName(name); }).filter(Boolean);
  if (obsolete.length) {
    // スプレッドシートには表示可能なタブが最低1枚必要。無関係なデータはすべて残す。
    let visible = ss.getSheets().find(function(sheet) { return names.indexOf(sheet.getName()) < 0 && !sheet.isSheetHidden(); });
    if (!visible) {
      visible = ss.getSheetByName('クラス一覧') || ensureSheet_('クラス一覧', ['クラス名', 'コースID', '同期対象(1)']);
      visible.showSheet();
    }
    obsolete.forEach(function(sheet) { ss.deleteSheet(sheet); });
    SpreadsheetApp.flush();
  }
  props.setProperty('APP_CONFIG_SHEETS_RETIRED', '1');
}

/**
 * Config シートから CONFIG_JSON を読み込む。
 * @returns {Object|null} パース済みオブジェクト、または null
 */
function readConfigSheet_() {
  const ss = getAppSpreadsheet_();
  const sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet || sheet.getLastRow() === 0) return null;
  const json = String(sheet.getRange(1, 1).getValue() || '').trim();
  if (!json) return null;
  try {
    return parseStoredConfig_(json);
  } catch (e) {
    throw new Error('旧Configの設定を読み取れません。旧シートを保持して復旧してください。' + e);
  }
}

/**
 * 設定シートから REMINDER_TO・FORM_SHEET_PREFIX・FORM_SS を読み込む。
 * 存在するキーの空欄は明示的な解除。キー自体がない旧シートのみ保存値にフォールバックする。
 * @returns {Object} 見つかった値だけを含む部分的な config オブジェクト
 */
function readSettingsInputSheet_() {
  const ss = getAppSpreadsheet_();
  const sheet = ss.getSheetByName(INPUT_SETTINGS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return {};

  const values = sheet.getDataRange().getDisplayValues();
  const result = {};
  const reminderTos = [];
  const prefixes = [];
  const formSources = [];
  const presentKeys = new Set();

  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0] || '').trim();
    const value = String(values[i][1] || '').trim();
    if (!key) continue;
    presentKeys.add(key);
    if (!value) continue;

    switch (key) {
      case 'REMINDER_TO':
        reminderTos.push(value);
        break;
      case 'FORM_SHEET_PREFIX':
        prefixes.push(value);
        break;
      case 'FORM_SS': {
        const normalized = normalizeFormSource_(value);
        if (normalized) formSources.push(normalized);
        break;
      }
    }
  }

  if (presentKeys.has('REMINDER_TO')) result.reminderTo = reminderTos;
  if (presentKeys.has('FORM_SHEET_PREFIX')) result.formSheetNamePrefix = prefixes;
  if (presentKeys.has('FORM_SS')) result.formSources = formSources;

  return result;
}

/** 内部保存が初めて成功するまでは、読み取り専用で旧形式に対応する。 */
function loadConfig_() {
  const props = PropertiesService.getScriptProperties();
  const internal = props.getProperty('APP_CONFIG_STORAGE') === 'properties-v1';
  const overrides = internal ? {} : readSettingsInputSheet_();
  const json = props.getProperty('APP_CONFIG');
  if (json) {
    try { return normalizeSavedConfig_(parseStoredConfig_(json), overrides); }
    catch (error) { props.setProperty('APP_CONFIG_PARSE_ERROR', '1'); }
  }
  if (internal) {
    props.setProperty('APP_CONFIG_PARSE_ERROR', '1');
    const backup = props.getProperty('APP_CONFIG_BACKUP');
    if (backup) return normalizeSavedConfig_(parseStoredConfig_(backup), {});
    throw new Error('内部設定を読み取れません。設定JSONまたは内部バックアップから復旧してください。');
  }
  const legacy = readConfigSheet_();
  if (legacy) return normalizeSavedConfig_(legacy, overrides);
  if (json) throw new Error('保存済み設定を読み取れません。旧設定を保持して復旧してください。');
  return normalizeAppConfig_(Object.assign({}, buildDefaultConfig_(), overrides));
}

function parseStoredConfig_(json) {
  const value = JSON.parse(json);
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error('設定の形式が不正です。');
  validateConfigDraft_(value);
  return value;
}

function normalizeSavedConfig_(saved, overrides) {
  const merged = Object.assign({}, buildLegacyDefaultConfig_(), saved, overrides);
  if (saved.gradeScale === undefined && Array.isArray(saved.fields) && !saved.fields.some(function(field) { return field && field.type === 'score_grade'; })) merged.gradeScale = [];
  if (saved.scoreSourceHeader === undefined && saved.scoreFieldKey !== undefined) delete merged.scoreSourceHeader;
  return normalizeAppConfig_(merged);
}

/** 外部サービスを使わずに UTF-8 のバイト数を数える。 */
function utf8ByteLength_(text) {
  let bytes = 0;
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    if (code < 0x80) bytes++;
    else if (code < 0x800) bytes += 2;
    else if (code >= 0xD800 && code <= 0xDBFF && i + 1 < text.length &&
        text.charCodeAt(i + 1) >= 0xDC00 && text.charCodeAt(i + 1) <= 0xDFFF) {
      bytes += 4;
      i++;
    } else {
      // BMP 内の文字と、対応しないサロゲートを置換文字にした場合を扱う。
      bytes += 3;
    }
  }
  return bytes;
}

function validateGradeScaleDraft_(gradeScale) {
  if (!Array.isArray(gradeScale)) throw new Error('評価変換表は配列で入力してください。');
  const gradeKeys = new Set();
  gradeScale.forEach(function(entry) {
    if (!entry || typeof entry !== 'object' || Array.isArray(entry) ||
        typeof entry.from !== 'string' || typeof entry.to !== 'string' || !entry.from.trim() || !entry.to.trim()) {
      throw new Error('評価変換表の変換前・変換後は空でない文字列を入力してください。');
    }
    const key = normalizeGradeScaleKey_(entry.from);
    if (gradeKeys.has(key)) throw new Error('評価変換表の変換前が重複しています: ' + entry.from);
    gradeKeys.add(key);
  });
}

/** 内部設定の本体と復旧用コピー。呼び出し元はアプリ共通ロックを保持する。 */
function saveConfig_(config) {
  const json = JSON.stringify(config), byteLength = utf8ByteLength_(json);
  if (byteLength > 8000) throw new Error('設定データが大きすぎます（' + byteLength + ' バイト）。フォーム回答スプレッドシートの件数や取り込み項目の数を減らしてください。');
  const props = PropertiesService.getScriptProperties();
  const keys = ['APP_CONFIG', 'APP_CONFIG_BACKUP', 'APP_CONFIG_STORAGE', 'APP_CONFIG_PARSE_ERROR', 'APP_CONFIG_SAVE_ERROR'];
  const previous = keys.map(function(key) { return [key, props.getProperty(key)]; });
  try {
    props.setProperty('APP_CONFIG_BACKUP', json);
    props.setProperty('APP_CONFIG', json);
    if (props.getProperty('APP_CONFIG') !== json || props.getProperty('APP_CONFIG_BACKUP') !== json) throw new Error('内部設定の保存を確認できません。');
    props.setProperty('APP_CONFIG_STORAGE', 'properties-v1');
    if (props.getProperty('APP_CONFIG_STORAGE') !== 'properties-v1') throw new Error('内部設定の移行を確認できません。');
    props.deleteProperty('APP_CONFIG_PARSE_ERROR');
    props.deleteProperty('APP_CONFIG_SAVE_ERROR');
  } catch (error) {
    const failures = [];
    previous.forEach(function(entry) {
      try {
        if (entry[1] === null) props.deleteProperty(entry[0]);
        else props.setProperty(entry[0], entry[1]);
        if (props.getProperty(entry[0]) !== entry[1]) throw new Error('復元値が一致しません');
      } catch (rollbackError) { failures.push(entry[0] + ': ' + rollbackError); }
    });
    if (failures.length) {
      try { props.setProperty('APP_CONFIG_SAVE_ERROR', failures.join('\n')); } catch (ignored) { Logger.log(ignored); }
      throw new Error('設定保存に失敗し、復旧も完了していません。自動実行を停止し、内部設定と設定JSONを確認してください。' + error + '\n' + failures.join('\n'));
    }
    throw error;
  }
}

function snapshotSheetContents_(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) return { name: name, exists: false };
  const range = sheet.getDataRange();
  const values = range.getValues();
  const formulas = typeof range.getFormulas === 'function' ? range.getFormulas() : [];
  formulas.forEach(function(row, r) { row.forEach(function(formula, c) { if (formula) values[r][c] = formula; }); });
  return { name: name, exists: true, values: values };
}

function restoreSheetContents_(ss, snapshot) {
  let sheet = ss.getSheetByName(snapshot.name);
  if (!snapshot.exists) {
    if (sheet) ss.deleteSheet(sheet);
    return;
  }
  if (!sheet) sheet = ss.insertSheet(snapshot.name);
  sheet.clearContents();
  if (snapshot.values.length && snapshot.values[0].length) sheet.getRange(1, 1, snapshot.values.length, snapshot.values[0].length).setValues(snapshot.values);
}

function getConfig_() {
  const saveError = PropertiesService.getScriptProperties().getProperty('APP_CONFIG_SAVE_ERROR');
  if (saveError) throw new Error('前回の設定保存からの復旧が未完了です。内部設定と設定JSONを確認してください。' + saveError);
  return loadConfig_();
}

function getConfiguredEvalHeaders_(config) {
  const headers = EVAL_BASE_HEADERS.slice();
  config.fields.forEach(function(field) {
    headers.push(getFieldEvalHeader_(field));
  });
  return uniqueHeaders_(headers);
}

function getConfiguredSendHeaders_(config) {
  const headers = SEND_BASE_HEADERS.slice();
  config.fields.forEach(function(field) {
    headers.push(getFieldSendHeader_(field));
  });
  headers.push(config.replyBodyHeader);
  return uniqueHeaders_(headers);
}

function uniqueHeaders_(headers) {
  const seen = {};
  const result = [];
  headers.forEach(function(header) {
    const key = String(header || '').trim();
    if (!key || seen[key]) return;
    seen[key] = true;
    result.push(key);
  });
  return result;
}

function hasMatchingHeaders_(sheet, expectedHeaders) {
  const lastCol = Math.max(sheet.getLastColumn(), expectedHeaders.length, 1);
  const current = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0]
    .slice(0, expectedHeaders.length)
    .map(function(value) { return String(value || '').trim(); });

  if (current.length < expectedHeaders.length) return false;
  for (let i = 0; i < expectedHeaders.length; i++) {
    if (current[i] !== expectedHeaders[i]) return false;
  }
  return true;
}

function ensureConfiguredSheet_(name, headers) {
  const ss = getAppSpreadsheet_();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return sheet;
  }

  const hasContent = sheet.getLastRow() > 0 || sheet.getLastColumn() > 0;
  if (!hasContent) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return sheet;
  }

  if (sheet.getLastColumn() === headers.length && hasMatchingHeaders_(sheet, headers)) {
    return sheet;
  }

  if (sheet.getLastRow() <= 1) {
    sheet.clear();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return sheet;
  }

  throw new Error('「' + name + '」シートの構成が設定と一致しません。管理画面で必要なデータを退避・クリアしてから再実行してください。');
}

function ensureSendSheet_(config) {
  return ensureConfiguredSheet_(SEND_SHEET_NAME, getConfiguredSendHeaders_(config));
}

function ensureStudentSheetForSync_() {
  const ss = getAppSpreadsheet_();
  let sheet = ss.getSheetByName(STUDENT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(STUDENT_SHEET_NAME);
  }

  if (!hasMatchingHeaders_(sheet, STUDENT_SHEET_HEADERS)) {
    if (sheet.getLastRow() > 0) throw new Error('「生徒一覧」の構成が一致しません。既存データを退避してから見出しを修正してください。');
    sheet.clear();
    sheet.getRange(1, 1, 1, STUDENT_SHEET_HEADERS.length).setValues([STUDENT_SHEET_HEADERS]);
  }

  sheet.setFrozenRows(1);
  return sheet;
}

function getResetImpactSummary_() {
  const ss = getAppSpreadsheet_();
  const evalSheet = ss.getSheetByName(EVAL_SHEET_NAME);
  const sendSheet = ss.getSheetByName(SEND_SHEET_NAME);

  const evalRows = evalSheet ? Math.max(evalSheet.getLastRow() - 1, 0) : 0;
  const sendRows = sendSheet ? Math.max(sendSheet.getLastRow() - 1, 0) : 0;

  let pendingSend = 0;
  let retryableSend = 0;
  let errorSend = 0;

  if (sendSheet && sendRows > 0) {
    const values = sendSheet.getDataRange().getDisplayValues();
    const headers = values[0].map(String);
    const idxStatus = headers.indexOf('送信状態');
    if (idxStatus >= 0) {
      for (let i = 1; i < values.length; i++) {
        const status = String(values[i][idxStatus] || '').trim();
        if (!status) pendingSend++;
        if (status === '未') retryableSend++;
        if (status === 'エラー') errorSend++;
      }
    }
  }

  return {
    evalRows: evalRows,
    sendRows: sendRows,
    pendingSend: pendingSend,
    retryableSend: retryableSend,
    errorSend: errorSend
  };
}

function buildResetSummaryText_(info) {
  const lines = [
    '設定を保存すると、次のシートを再生成します。',
    '・評価データ: ' + info.evalRows + '件を削除',
    '・送信シート: ' + info.sendRows + '件を削除'
  ];

  if (info.pendingSend > 0 || info.retryableSend > 0 || info.errorSend > 0) {
    lines.push('');
    lines.push('送信シートの内訳');
    lines.push('・未送信: ' + info.pendingSend + '件');
    lines.push('・再送待ち: ' + info.retryableSend + '件');
    lines.push('・エラー: ' + info.errorSend + '件');
  }

  lines.push('');
  lines.push('この操作は取り消せません。');
  lines.push('続行する場合は保存を確定してください。');

  return lines.join('\n');
}

function resetConfiguredSheets_(config) {
  resetSheetWithHeaders_(SEND_SHEET_NAME, getConfiguredSendHeaders_(config));
}

function resetSheetWithHeaders_(name, headers) {
  const ss = getAppSpreadsheet_();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  } else {
    sheet.clear();
  }

  if (sheet.getMaxColumns() > headers.length) {
    sheet.deleteColumns(headers.length + 1, sheet.getMaxColumns() - headers.length);
  }
  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
}

function getFieldByKey_(config, key) {
  for (let i = 0; i < config.fields.length; i++) {
    if (config.fields[i].key === key) return config.fields[i];
  }
  return null;
}

function getFieldBySourceHeader_(config, sourceHeader) {
  const target = String(sourceHeader || '').trim();
  for (let i = 0; i < config.fields.length; i++) {
    if (String(config.fields[i].sourceHeader || '').trim() === target) return config.fields[i];
  }
  return null;
}

function resolveScoreField_(config) {
  return getFieldBySourceHeader_(config, config.scoreSourceHeader)
    || getFieldByKey_(config, config.scoreFieldKey);
}

function buildLegacyGradeScale_() {
  return [{ from: '5', to: 'A' }, { from: '4', to: 'B+' }, { from: '3', to: 'B' },
    { from: '2', to: 'B-' }, { from: '1', to: 'C' }];
}

/** 10進数の文字列は Number の丸めを避けて数値として比較し、その他の文字列は大文字と小文字を区別する。 */
function normalizeGradeScaleKey_(value) {
  const text = String(value == null ? '' : value).trim();
  if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$/.test(text)) return 'text:' + text;
  const negative = text.charAt(0) === '-';
  const parts = text.replace(/^[+-]/, '').split('.');
  const integer = (parts[0] || '0').replace(/^0+(?=\d)/, '');
  const fraction = (parts[1] || '').replace(/0+$/, '');
  return 'number:' + (negative && (integer !== '0' || fraction) ? '-' : '') + integer + (fraction ? '.' + fraction : '');
}

function processConfiguredFieldValue_(field, rawValue, displayValue, config) {
  if (field.type === 'date') {
    return toYmdString(displayValue || rawValue);
  }
  if (field.type === 'score_grade') {
    if (!config && !Object.prototype.hasOwnProperty.call(field, 'gradeScale')) return scoreToGrade(rawValue) || String(rawValue == null ? '' : rawValue).trim();
    const key = normalizeGradeScaleKey_(rawValue);
    const entry = getFieldGradeScale_(field, config).find(function(item) { return normalizeGradeScaleKey_(item.from) === key; });
    return entry ? String(entry.to).trim() : String(rawValue == null ? '' : rawValue).trim();
  }
  return String(rawValue == null ? '' : rawValue).trim();
}

function buildTemplateLabelMap_(config) {
  const labels = {
    student_name: '名前',
    class_name: 'クラス名',
    course_id: 'コースID',
    row_no: 'No'
  };

  (config && config.fields || []).forEach(function(field) {
    if (!field || !field.key) return;
    labels[field.key] = String(field.sendHeader || field.evalHeader || field.sourceHeader || field.key).trim();
  });

  return labels;
}

function friendlyPlaceholderPattern_() {
  // 旧形式の山括弧トークンと明示的な {{sections}} を、単独の波括弧と区別する。
  return /[＜<]\s*([A-Za-z_][A-Za-z0-9_]*)\s*[＞>]|(?<!\{)\{\s*([A-Za-z_][A-Za-z0-9_]*)\s*\}(?!\})/g;
}

function friendlyPlaceholderOnlyPattern_() {
  return /^(?:[＜<]\s*[A-Za-z_][A-Za-z0-9_]*\s*[＞>]|\{\s*[A-Za-z_][A-Za-z0-9_]*\s*\})$/;
}

function replaceFriendlyPlaceholdersOnly_(text) {
  return normalizeMultiline_(text || '').replace(friendlyPlaceholderPattern_(), function(_, legacyKey, key) {
    return '{{' + (legacyKey || key) + '}}';
  });
}

function convertFriendlyTemplateToSections_(template, config) {
  const labels = buildTemplateLabelMap_(config);
  const normalized = normalizeMultiline_(template || '').trim();
  if (!normalized) return '';

  return normalized.split(/\n{2,}/).map(function(block) {
    const lines = block.split('\n');
    const lineInfos = lines.map(function(line) {
      const keys = [];
      line.replace(friendlyPlaceholderPattern_(), function(_, legacyKey, key) {
        keys.push(legacyKey || key);
        return _;
      });

      return {
        raw: line,
        keys: keys,
        replaced: replaceFriendlyPlaceholdersOnly_(line),
        placeholderOnly: keys.length === 1 && friendlyPlaceholderOnlyPattern_().test(line.trim())
      };
    });

    const uniqueKeys = {};
    const orderedKeys = [];
    lineInfos.forEach(function(info) {
      info.keys.forEach(function(key) {
        if (uniqueKeys[key]) return;
        uniqueKeys[key] = true;
        orderedKeys.push(key);
      });
    });

    let mainKey = '';
    lineInfos.forEach(function(info) {
      if (info.placeholderOnly) mainKey = info.keys[0];
    });
    if (!mainKey && orderedKeys.length === 1) {
      mainKey = orderedKeys[0];
    }

    const convertedLines = lineInfos.map(function(info) {
      if (info.keys.length === 0) return info.raw;

      if (info.placeholderOnly) {
        const key = info.keys[0];
        if (key === mainKey) {
          return '{{' + key + '}}';
        }
        const label = labels[key];
        const content = label ? (label + '：{{' + key + '}}') : ('{{' + key + '}}');
        return '{{#' + key + '}}' + content + '{{/' + key + '}}';
      }

      if (info.keys.length === 1) {
        const key = info.keys[0];
        if (key === mainKey) {
          return info.replaced;
        }
        return '{{#' + key + '}}' + info.replaced + '{{/' + key + '}}';
      }

      return info.replaced;
    });

    const joined = convertedLines.join('\n');
    if (!mainKey) return joined;
    return '{{#' + mainKey + '}}' + joined + '{{/' + mainKey + '}}';
  }).join('\n\n');
}

function renderTemplateWithSections_(template, context, config) {
  let result = normalizeMultiline_(template || '');
  if (result.indexOf('{{') >= 0) {
    result = replaceFriendlyPlaceholdersOnly_(result);
  } else {
    result = convertFriendlyTemplateToSections_(result, config);
  }
  const sectionPattern = /{{#([A-Za-z_][A-Za-z0-9_]*)}}([\s\S]*?){{\/\1}}/g;

  let guard = 0;
  while (guard < 20) {
    sectionPattern.lastIndex = 0;
    if (!sectionPattern.test(result)) break;
    sectionPattern.lastIndex = 0;
    result = result.replace(sectionPattern, function(_, key, inner) {
      const value = String(context[key] == null ? '' : context[key]).trim();
      return value ? inner : '';
    });
    guard++;
  }

  result = result.replace(/{{([A-Za-z_][A-Za-z0-9_]*)}}/g, function(_, key) {
    return String(context[key] == null ? '' : context[key]);
  });

  return result
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

/** このキャッシュはワーカーの1回の実行中だけ使い、実行間では保存しない。 */
function createFormStatusCache_() {
  return { spreadsheets: new Map(), targets: new Map() };
}

function updateFormStatus_(config, ssId, sheetName, rowIndex, status, invocationCache) {
  if (!ssId || !sheetName || !rowIndex) return;

  try {
    const cache = invocationCache || createFormStatusCache_();
    if (!cache.spreadsheets.has(ssId)) cache.spreadsheets.set(ssId, SpreadsheetApp.openById(ssId));
    const key = JSON.stringify([ssId, sheetName, config.formStatusHeader]);
    if (!cache.targets.has(key)) {
      const sheet = cache.spreadsheets.get(ssId).getSheetByName(sheetName);
      const headers = sheet && sheet.getLastColumn() > 0
        ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(String) : [];
      cache.targets.set(key, { sheet: sheet, statusColumn: headers.indexOf(config.formStatusHeader) + 1 });
    }
    const target = cache.targets.get(key);
    if (!target.sheet) throw new Error('元回答シートが見つかりません: ' + sheetName);
    if (!target.statusColumn) throw new Error('元回答シートの状態列が見つかりません: ' + config.formStatusHeader);
    target.sheet.getRange(rowIndex, target.statusColumn).setValue(status);
  } catch (e) {
    ensureErrorSheet_().appendRow([
      new Date(),
      sheetName || 'Form更新',
      rowIndex,
      '',
      status,
      'FORM_UPDATE_ERROR',
      String(e)
    ]);
  }
}

function getStudentSheet_() {
  const ss = getAppSpreadsheet_();
  const sheet = ss.getSheetByName(STUDENT_SHEET_NAME);
  if (!sheet) {
    throw new Error('生徒一覧シートが見つかりません: ' + STUDENT_SHEET_NAME);
  }
  return sheet;
}

function getSelectedClassRecords_() {
  const classSheet = ensureSheet_('クラス一覧', ['クラス名', 'コースID', '同期対象(1)']);
  const lastRow = classSheet.getLastRow();
  if (lastRow < 2) {
    throw new Error('クラス一覧が空です。設定の「クラスと生徒」で一覧を取得してください。');
  }

  const values = classSheet.getRange(2, 1, lastRow - 1, 3).getDisplayValues();
  const classes = values.map(function(row) {
    return {
      className: String(row[0] || '').trim(),
      courseId: String(row[1] || '').trim(),
      syncFlag: String(row[2] || '').trim()
    };
  }).filter(function(item) {
    return item.courseId && item.syncFlag === '1';
  });

  if (classes.length === 0) {
    throw new Error('クラス一覧で「同期対象(1)」が選択されていません。先に対象クラスへ 1 を入れてください。');
  }

  return classes;
}

function collectFormTargetSheets_(config) {
  const prefixes = Array.isArray(config && config.formSheetNamePrefix) ? config.formSheetNamePrefix : [];
  const formSources = Array.isArray(config && config.formSources) ? config.formSources : [];
  if (formSources.length === 0) {
    throw new Error('管理画面の「フォームと通知先」にフォーム回答スプレッドシートを登録してください。');
  }

  const seen = {};
  const rows = [];

  formSources.forEach(function(source) {
    const sourceId = String(source && source.id || '').trim();
    if (!sourceId) return;

    const formSs = SpreadsheetApp.openById(sourceId);
    const spreadsheetName = formSs.getName();
    formSs.getSheets().forEach(function(sheet) {
      const sheetName = sheet.getName();
      if (isManagedInternalSheet_(sheetName, sourceId) || !sheetMatchesPrefixes_(sheetName, prefixes)) return;

      const key = makeMappingLookupKey_(sourceId, sheetName);
      if (seen[key]) return;
      seen[key] = true;
      rows.push({
        sourceId: sourceId,
        spreadsheetName: spreadsheetName,
        sheetName: sheetName
      });
    });
  });

  if (rows.length === 0) {
    const activePrefixes = prefixes.filter(function(p) { return !!p; });
    throw new Error(
      'FORM_SS に対象シートがありません。' +
      (activePrefixes.length > 0
        ? ' 接頭辞「' + activePrefixes.join('」「') + '」を確認してください。'
        : '')
    );
  }

  return rows;
}

function getMappingSheet_() {
  const ss = getAppSpreadsheet_();
  return ss.getSheetByName(MAPPING_SHEET_NAME);
}

function readMappingEntries_(strict) {
  const sheet = getMappingSheet_();
  const map = new Map();
  if (!sheet || sheet.getLastRow() < 2) {
    return map;
  }

  const values = sheet.getDataRange().getDisplayValues();
  const headers = values[0].map(String);
  const headerMap = createHeaderMap_(headers);
  ['元SS_ID', '元シート名', 'クラス名', 'courseId', 'メモ'].forEach(function(header) {
    if (!(header in headerMap)) {
      if (!strict) return;
      throw new Error('対応表シートのヘッダが不正です。設定の「回答先と通知先」で対応表を確認してください。');
    }
  });
  if (!('元SS_ID' in headerMap) || !('元シート名' in headerMap) || !('クラス名' in headerMap) || !('courseId' in headerMap) || !('メモ' in headerMap)) {
    return map;
  }

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const sourceId = String(row[headerMap['元SS_ID']] || '').trim();
    const sheetName = String(row[headerMap['元シート名']] || '').trim();
    if (!sourceId || !sheetName) continue;

    map.set(makeMappingLookupKey_(sourceId, sheetName), {
      className: String(row[headerMap['クラス名']] || '').trim(),
      courseId: String(row[headerMap['courseId']] || '').trim(),
      note: String(row[headerMap['メモ']] || '').trim()
    });
  }

  return map;
}

function createMappingSheet() {
  const message = withAppLock_(createMappingSheetUnlocked_);
  safeAlert_(message);
}

function createMappingSheetUnlocked_() {
  const config = getConfig_();
  const classRecords = getSelectedClassRecords_();
  const targetSheets = collectFormTargetSheets_(config);
  const existingMap = readMappingEntries_(false);

  const classNameSet = {};
  classRecords.forEach(function(item) {
    const className = String(item.className || '').trim();
    if (!className) return;
    if (classNameSet[className]) {
      throw new Error('同期対象クラスに同名のクラス名があります。対応表のドロップダウンに使うため、クラス名が重複しない状態で実行してください: ' + className);
    }
    classNameSet[className] = true;
  });

  const ss = getAppSpreadsheet_();
  let sheet = ss.getSheetByName(MAPPING_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(MAPPING_SHEET_NAME);
  } else {
    const lastRow = Math.max(sheet.getLastRow(), 1);
    const lastCol = Math.max(sheet.getLastColumn(), MAPPING_SHEET_HEADERS.length);
    sheet.getRange(1, 1, lastRow, lastCol).clearDataValidations();
    sheet.clear();
  }

  const rows = [MAPPING_SHEET_HEADERS];
  targetSheets.forEach(function(item) {
    const existing = existingMap.get(makeMappingLookupKey_(item.sourceId, item.sheetName)) || {};
    rows.push([
      item.sourceId,
      item.spreadsheetName,
      item.sheetName,
      existing.className || '',
      '',
      existing.note || ''
    ]);
  });

  sheet.getRange(1, 1, rows.length, MAPPING_SHEET_HEADERS.length).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, MAPPING_SHEET_HEADERS.length);

  const numRows = rows.length - 1;
  if (numRows > 0) {
    const classNames = classRecords.map(function(item) { return item.className; });
    const validation = SpreadsheetApp.newDataValidation()
      .requireValueInList(classNames, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, 4, numRows, 1).setDataValidation(validation);

    const formulas = [];
    for (let i = 0; i < numRows; i++) {
      const rowNum = i + 2;
      formulas.push([buildCourseIdFormulaForRow_(rowNum, classRecords)]);
    }
    sheet.getRange(2, 5, numRows, 1).setFormulas(formulas);
  }

  return '対応表シートを作成/更新しました（' + targetSheets.length + '件）';
}

function appendRows_(sheet, rows) {
  if (!rows || rows.length === 0) return;
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
}

function classroomdata() {
  return withAppLock_(classroomdataUnlocked_);
}

function classroomdataUnlocked_() {
  const ss = getAppSpreadsheet_();
  const existing = ss.getSheetByName('クラス一覧');
  const flags = new Map();
  if (existing && existing.getLastRow() > 0) {
    if (!hasMatchingHeaders_(existing, ['クラス名', 'コースID', '同期対象(1)'])) throw new Error('クラス一覧の構成が一致しません。見出しを確認してください。');
    existing.getDataRange().getValues().slice(1).forEach(function(row) { flags.set(String(row[1] || '').trim(), row[2]); });
  }
  const courses = [];
  let pageToken = '';
  do {
    const response = Classroom.Courses.list({ teacherId: Session.getActiveUser(), courseStates: 'ACTIVE', pageSize: 100, pageToken: pageToken });
    if (!response || typeof response !== 'object' || (response.courses != null && !Array.isArray(response.courses))) throw new Error('クラス一覧の取得結果が不正です。');
    Array.prototype.push.apply(courses, response.courses || []);
    pageToken = response.nextPageToken || '';
  } while (pageToken);
  replaceRosterRows_('クラス一覧', ['クラス名', 'コースID', '同期対象(1)'], courses.map(function(course) {
    const id = String(course.id);
    return [course.name, id, flags.has(id) ? flags.get(id) : ''];
  }));
  return 'クラス一覧を更新しました（' + courses.length + '件）。同期対象は保持しています。';
}

function replaceRosterRows_(name, headers, rows) {
  const ss = getAppSpreadsheet_();
  const snapshot = snapshotSheetContents_(ss, name);
  try {
    const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
    sheet.clearContents();
    sheet.getRange(1, 1, rows.length + 1, headers.length).setValues([headers].concat(rows));
    sheet.setFrozenRows(1);
    SpreadsheetApp.flush();
  } catch (error) {
    try { restoreSheetContents_(ss, snapshot); SpreadsheetApp.flush(); }
    catch (rollbackError) { throw new Error('名簿更新と復旧に失敗しました。シートを確認してください。' + error + ' / ' + rollbackError); }
    throw error;
  }
}

function studentdataMulti() {
  const message = withAppLock_(studentdataMultiUnlocked_);
  safeAlert_(message);
}

function studentdataMultiUnlocked_() {
  const ss = getAppSpreadsheet_();
  const classSheet = ss.getSheetByName('クラス一覧');
  if (!classSheet || classSheet.getLastRow() < 2) return 'クラス一覧が空です。既存の生徒一覧は保持しています。';
  if (!hasMatchingHeaders_(classSheet, ['クラス名', 'コースID', '同期対象(1)'])) throw new Error('クラス一覧の構成が一致しません。');
  const selected = classSheet.getRange(2, 1, classSheet.getLastRow() - 1, 3).getValues().filter(function(row) {
    return String(row[1] || '').trim() && String(row[2] || '').trim() === '1';
  });
  if (!selected.length) return '同期対象クラスがありません。クラス一覧の同期対象に1を入力してください。既存の生徒一覧は保持しています。';
  const studentSheet = ss.getSheetByName(STUDENT_SHEET_NAME);
  if (studentSheet && studentSheet.getLastRow() && !hasMatchingHeaders_(studentSheet, STUDENT_SHEET_HEADERS)) throw new Error('生徒一覧の構成が一致しません。既存データと見出しを確認してください。');
  const rows = [];
  const courseIds = [];
  selected.forEach(function(item) {
    const courseId = String(item[1]).trim();
    if (courseIds.indexOf(courseId) >= 0) return;
    courseIds.push(courseId);
    getStudentListMax(courseId).forEach(function(student) {
      rows.push([rows.length + 1, student.email, student.name, item[0] || '', courseId, student.studentId]);
    });
  });
  replaceRosterRows_(STUDENT_SHEET_NAME, STUDENT_SHEET_HEADERS, rows);
  PropertiesService.getScriptProperties().setProperty('TURRET_ROSTER_SELECTION', JSON.stringify(courseIds.sort()));
  return '生徒一覧を更新しました（' + rows.length + '名）。';
}

function getStudentListMax(classId) {
  let pageToken = '';
  const all = [];

  do {
    const response = Classroom.Courses.Students.list(classId, { pageToken: pageToken });
    if (!response || typeof response !== 'object' || (response.students != null && !Array.isArray(response.students))) throw new Error('生徒一覧の取得結果が不正です。');
    const list = response.students || [];
    for (let i = 0; i < list.length; i++) {
      const student = list[i];
      all.push({
        email: String(student.profile && student.profile.emailAddress || '').trim(),
        name: String(student.profile && student.profile.name && student.profile.name.fullName || '').trim(),
        studentId: String(student.profile && student.profile.id || '').trim()
      });
    }
    pageToken = response ? response.nextPageToken : '';
  } while (pageToken);

  return all;
}

/** 旧トリガー・手動実行からの入口も、送信データの準備へ統一する。 */
function importFromFormsToEval() { return prepareSendData(); }
function importFromFormsToEvalUnlocked_() { return prepareSendDataUnlocked_(); }
function evalToSendSheet() { return prepareSendData(); }
function evalToSendSheetUnlocked_() { return prepareSendDataUnlocked_(); }

function prepareSendData() { return withAppLock_(prepareSendDataUnlocked_); }

/** 未移行データがある間だけ、旧保存列と項目の一対一の対応を保護する。 */
function validateLegacyEvaluationFields_(config) {
  const headers = new Set(EVAL_BASE_HEADERS);
  config.fields.forEach(function(field) {
    const header = getFieldEvalHeader_(field);
    if (headers.has(header)) throw new Error('旧評価データの項目列が重複しています: ' + header + '。移行を完了するか、旧列の対応を確認してください。');
    headers.add(header);
  });
}

/** 旧評価データは移行元として読むだけで、新しいシート・回答行を作らない。 */
function readLegacyEvaluation_(config) {
  const sheet = getAppSpreadsheet_().getSheetByName(EVAL_SHEET_NAME);
  const values = sheet && sheet.getLastRow() > 1 ? sheet.getDataRange().getDisplayValues() : [];
  const headerMap = createHeaderMap_(values[0] || []);
  const rows = values.slice(1).filter(function(row) { return row.some(function(value) { return String(value).trim(); }); });
  if (rows.length) {
    const headers = values[0].map(function(value) { return String(value).trim(); });
    if (new Set(headers).size !== headers.length || headers.some(function(header) { return !header; })) {
      throw new Error('旧評価データの列名が重複または空欄です。移行前に列を確認してください。');
    }
    const required = EVAL_BASE_HEADERS.slice();
    // 完了済みの履歴は、今後の項目変更を妨げない。未移行・失敗行は旧列を保護する。
    if (rows.some(function(row) { return String(row[headerMap['処理状態']] || '').trim() !== '準備〇'; })) {
      validateLegacyEvaluationFields_(config);
      config.fields.forEach(function(field) { required.push(getFieldEvalHeader_(field)); });
    }
    const missing = required.filter(function(header) { return !(header in headerMap); });
    if (missing.length) throw new Error('旧評価データの移行に必要な列がありません: ' + missing.join(', ') + '。旧設定・列を確認してください。');
  }
  return { sheet: sheet, values: values, headerMap: headerMap, keys: collectResponseKeys_(values, headerMap) };
}

/** 旧取り込み値を優先し、新しい回答は変換結果をメモリ上だけに保持する。 */
function collectSendPreparationCandidates_(config, cache, guard, sendKeys) {
  const legacy = readLegacyEvaluation_(config), candidates = [], lh = legacy.headerMap;
  legacy.values.slice(1).forEach(function(row, index) {
    if (!row.some(function(value) { return String(value).trim(); }) || String(row[lh['処理状態']] || '').trim()) return;
    const item = { sourceId: String(row[lh['元SS_ID']] || '').trim(), sourceSheet: String(row[lh['元シート名']] || '').trim(),
      sourceRow: Number(row[lh['元行番号']] || 0), email: String(row[lh['メールアドレス']] || '').trim(),
      name: String(row[lh['名前']] || '').trim(), legacyRow: index + 2, fields: {} };
    config.fields.forEach(function(field) { item.fields[field.key] = row[lh[getFieldEvalHeader_(field)]]; });
    candidates.push(item);
  });
  const scoreField = resolveScoreField_(config), read = guard && guard.read, seen = new Set(legacy.keys);
  if (!scoreField) throw new Error('scoreSourceHeader に対応する項目が見つかりません。');
  config.formSources.forEach(function(source) {
    if (!cache.spreadsheets.has(source.id)) cache.spreadsheets.set(source.id, SpreadsheetApp.openById(source.id));
    const book = cache.spreadsheets.get(source.id);
    (read ? scoringSheets_(read,source.id) : book.getSheets()).forEach(function(sheet) {
      const name = sheet.getName();
      if (isManagedInternalSheet_(name,source.id) || !sheetMatchesPrefixes_(name,config.formSheetNamePrefix)) return;
      const snapshot = read && read.snapshots.get(JSON.stringify([source.id,sheet.getSheetId()]));
      const lastRow = snapshot ? snapshot.lastRow : sheet.getLastRow(), lastCol = snapshot ? snapshot.lastCol : sheet.getLastColumn();
      if (lastRow <= 1 || !lastCol) return;
      const values = snapshot ? snapshot.values : sheet.getRange(1,1,lastRow,lastCol).getValues();
      const display = sheet.getRange(1,1,lastRow,lastCol).getDisplayValues(), headers = values[0].map(String), h = createHeaderMap_(headers);
      cache.targets.set(JSON.stringify([source.id,name,config.formStatusHeader]), {sheet:sheet,statusColumn:headers.indexOf(config.formStatusHeader)+1});
      const missing = [config.emailHeader,config.studentNameHeader,config.formStatusHeader].concat(config.fields.map(function(field){return field.sourceHeader;}))
        .filter(function(header){return !(header in h);});
      if (missing.length) { ensureErrorSheet_().appendRow([new Date(),name,'-','','','CONFIG','不足ヘッダ: '+missing.join(', ')]); return; }
      values.slice(1).forEach(function(row,index) {
        const sourceRow = index + 2, key = makeResponseKey_(source.id,name,sourceRow), score = row[h[scoreField.sourceHeader]];
        const email = String(row[h[config.emailHeader]] || '').trim();
        if (!email || String(row[h[config.formStatusHeader]] || '').trim() || String(score == null ? '' : score).trim() === '' || seen.has(key) || sendKeys.has(key)) return;
        if (!scoringMayReturn_(guard,key)) return;
        const item = {sourceId:source.id,sourceSheet:name,sourceRow:sourceRow,email:email,name:String(row[h[config.studentNameHeader]] || '').trim(),fields:{}};
        config.fields.forEach(function(field) { item.fields[field.key] = processConfiguredFieldValue_(field,row[h[field.sourceHeader]],display[index+1][h[field.sourceHeader]],config); });
        candidates.push(item); seen.add(key);
      });
    });
  });
  return {legacy:legacy,candidates:candidates};
}

function prepareSendDataUnlocked_() {
  if(typeof assertNoManagedFormUpdate_==='function')assertNoManagedFormUpdate_();
  if(typeof refreshAllManagedNames_==='function')refreshAllManagedNames_();
  const config = getConfig_();
  validateAppConfig_(config);
  const formStatusCache = createFormStatusCache_();
  const scoringGuard=scoringBuildReturnGuard_(formStatusCache);

  const sendSheet = ensureSendSheet_(config);
  const studentSheet = getStudentSheet_();
  const mappingEntries = readMappingEntries_(true);

  const studentValues = studentSheet.getDataRange().getDisplayValues();
  if (studentValues.length <= 1) {
    throw new Error('生徒一覧にデータがありません。');
  }

  const studentHeaders = studentValues[0].map(String);
  const studentHeaderMap = createHeaderMap_(studentHeaders);
  ['メールアドレス', '名前', 'クラス名', 'コースID', 'studentId'].forEach(function(header) {
    if (!(header in studentHeaderMap)) {
      throw new Error('生徒一覧のヘッダに「' + header + '」が必要です。');
    }
  });

  const studentMap = new Map();
  const studentsByEmail = new Map();
  for (let i = 1; i < studentValues.length; i++) {
    const row = studentValues[i];
    const email = String(row[studentHeaderMap['メールアドレス']] || '').trim();
    const courseId = String(row[studentHeaderMap['コースID']] || '').trim();
    const lookupKey = makeStudentLookupKey_(courseId, email);
    if (!email || !courseId) continue;
    const record = {
      name: String(row[studentHeaderMap['名前']] || '').trim(),
      className: String(row[studentHeaderMap['クラス名']] || '').trim(),
      courseId: courseId,
      studentId: String(row[studentHeaderMap['studentId']] || '').trim()
    };
    studentMap.set(lookupKey, record);

    const emailKey = String(email || '').trim().toLowerCase();
    const bucket = studentsByEmail.get(emailKey) || [];
    bucket.push(record);
    studentsByEmail.set(emailKey, bucket);
  }

  const sendHeaders = sendSheet.getRange(1, 1, 1, sendSheet.getLastColumn()).getDisplayValues()[0].map(String);
  const sendHeaderMap = createHeaderMap_(sendHeaders);
  const existingResponses = collectResponseKeys_(sendSheet.getDataRange().getDisplayValues(), sendHeaderMap);
  const input = collectSendPreparationCandidates_(config, formStatusCache, scoringGuard, existingResponses);

  let nextNo = 1;
  if (sendSheet.getLastRow() > 1 && sendHeaderMap.No != null) {
    const lastNo = Number(sendSheet.getRange(sendSheet.getLastRow(), sendHeaderMap.No + 1).getValue() || 0);
    if (!isNaN(lastNo) && lastNo > 0) nextNo = lastNo + 1;
  }

  const sendRows = [];
  const stateUpdates = [];
  const errorSheet = ensureErrorSheet_();

  for (const item of input.candidates) {
    const srcSsId = item.sourceId, srcSheetName = item.sourceSheet, srcRow = item.sourceRow;
    const email = item.email, sourceName = item.name;
    const responseKey = makeResponseKey_(srcSsId, srcSheetName, srcRow);
    if (!responseKey) {
      throw new Error('旧評価データの元回答情報が不正です。行 ' + item.legacyRow + ' の元SS_ID・元シート名・元行番号を確認してください。');
    }
    if (!scoringMayReturn_(scoringGuard,responseKey)) continue;
    if (existingResponses.has(responseKey)) {
      // 既存の送信行は保持し、元フォームの最終状態（済など）を巻き戻さない。
      stateUpdates.push({ legacyRow: item.legacyRow, state: '準備〇' });
      continue;
    }

    if (!email) {
      errorSheet.appendRow([new Date(), item.legacyRow ? EVAL_SHEET_NAME : srcSheetName, item.legacyRow || srcRow, '', '', 'CONFIG', 'メールアドレスが空のため本文を準備できません。元回答または旧評価データを確認してください。']);
      stateUpdates.push({ legacyRow: item.legacyRow, state: '準備×', srcState: '準備×', srcSsId: srcSsId, srcSheetName: srcSheetName, srcRow: srcRow });
      continue;
    }

    const mapping = mappingEntries.get(makeMappingLookupKey_(srcSsId, srcSheetName));
    const mappedCourseId = String(mapping && mapping.courseId || '').trim();
    let student = null;

    if (mappedCourseId) {
      student = studentMap.get(makeStudentLookupKey_(mappedCourseId, email)) || null;
      if (!student) {
        errorSheet.appendRow([
          new Date(),
          SEND_SHEET_NAME,
          '-',
          email,
          '',
          'CONFIG',
          '対応表の courseId と生徒一覧が一致しません: ' + srcSheetName + ' / ' + mappedCourseId
        ]);
        stateUpdates.push({ legacyRow: item.legacyRow, state: '準備×', srcState: '準備×', srcSsId: srcSsId, srcSheetName: srcSheetName, srcRow: srcRow });
        continue;
      }
    }

    const candidates = studentsByEmail.get(String(email).trim().toLowerCase()) || [];
    if (!student && candidates.length === 1) {
      student = candidates[0];
    }

    if (!student && candidates.length > 1) {
      errorSheet.appendRow([
        new Date(),
        SEND_SHEET_NAME,
        '-',
        email,
        '',
        'CONFIG',
        '同じメールアドレスが複数クラスに存在するため送信先クラスを特定できません。対応表シートで元シート名にクラスを割り当ててください'
      ]);
      stateUpdates.push({ legacyRow: item.legacyRow, state: '準備×', srcState: '準備×', srcSsId: srcSsId, srcSheetName: srcSheetName, srcRow: srcRow });
      continue;
    }

    if (!student) {
      errorSheet.appendRow([
        new Date(),
        SEND_SHEET_NAME,
        '-',
        email,
        '',
        'FATAL',
        '生徒一覧にメールアドレスが存在しません'
      ]);
      stateUpdates.push({ legacyRow: item.legacyRow, state: '準備×', srcState: '準備×', srcSsId: srcSsId, srcSheetName: srcSheetName, srcRow: srcRow });
      continue;
    }

    if (!student.courseId) {
      errorSheet.appendRow([
        new Date(),
        SEND_SHEET_NAME,
        '-',
        email,
        '',
        'FATAL',
        '生徒一覧にコースIDがありません'
      ]);
      stateUpdates.push({ legacyRow: item.legacyRow, state: '準備×', srcState: '準備×', srcSsId: srcSsId, srcSheetName: srcSheetName, srcRow: srcRow });
      continue;
    }

    if (!student.studentId) {
      errorSheet.appendRow([
        new Date(),
        SEND_SHEET_NAME,
        '-',
        email,
        '',
        'FATAL',
        '生徒一覧に studentId がありません。生徒一覧取得をやり直してください'
      ]);
      stateUpdates.push({ legacyRow: item.legacyRow, state: '準備×', srcState: '準備×', srcSsId: srcSsId, srcSheetName: srcSheetName, srcRow: srcRow });
      continue;
    }

    const rowNo = nextNo++;
    const context = {
      student_name: student.name || sourceName,
      class_name: student.className,
      course_id: student.courseId,
      row_no: String(rowNo)
    };

    config.fields.forEach(function(field) {
      context[field.key] = String(item.fields[field.key] == null ? '' : item.fields[field.key]).trim();
    });

    const replyBody = renderTemplateWithSections_(config.messageTemplate, context, config);
    if (!replyBody) {
      errorSheet.appendRow([
        new Date(),
        SEND_SHEET_NAME,
        '-',
        email,
        '',
        'CONFIG',
        '返信テンプレートの結果が空になりました'
      ]);
      stateUpdates.push({ legacyRow: item.legacyRow, state: '準備×', srcState: '準備×', srcSsId: srcSsId, srcSheetName: srcSheetName, srcRow: srcRow });
      continue;
    }

    const sendRow = new Array(sendHeaders.length).fill('');
    sendRow[sendHeaderMap['元SS_ID']] = srcSsId;
    sendRow[sendHeaderMap['元シート名']] = srcSheetName;
    sendRow[sendHeaderMap['元行番号']] = srcRow;
    sendRow[sendHeaderMap['No']] = rowNo;
    sendRow[sendHeaderMap['メールアドレス']] = email;
    sendRow[sendHeaderMap['名前']] = student.name || sourceName;
    sendRow[sendHeaderMap['クラス名']] = student.className;
    sendRow[sendHeaderMap['コースID']] = student.courseId;
    sendRow[sendHeaderMap['studentId']] = student.studentId;
    sendRow[sendHeaderMap['送信状態']] = '';
    sendRow[sendHeaderMap[config.replyBodyHeader]] = replyBody;

    config.fields.forEach(function(field) {
      const sendHeader = getFieldSendHeader_(field);
      if (sendHeader && sendHeaderMap[sendHeader] != null) {
        sendRow[sendHeaderMap[sendHeader]] = context[field.key];
      }
    });

    sendRows.push(sendRow);
    existingResponses.add(responseKey);
    stateUpdates.push({ legacyRow: item.legacyRow, state: '準備〇', srcState: '準備〇', srcSsId: srcSsId, srcSheetName: srcSheetName, srcRow: srcRow });
  }

  appendRows_(sendSheet, sendRows);
  if (sendRows.length > 0) SpreadsheetApp.flush();

  stateUpdates.forEach(function(update) {
    if (update.legacyRow) input.legacy.sheet.getRange(update.legacyRow, input.legacy.headerMap['処理状態'] + 1).setValue(update.state);
    if (update.srcSsId && update.srcSheetName && update.srcRow) {
      updateFormStatus_(config, update.srcSsId, update.srcSheetName, update.srcRow, update.srcState, formStatusCache);
    }
  });
}

function postAnnouncementIndividual_(courseId, studentId, name, text) {
  const data = {
    courseId: String(courseId),
    text: name + 'さんへ(個別メッセージ)\n\n' + text,
    assigneeMode: 'INDIVIDUAL_STUDENTS',
    individualStudentsOptions: { studentIds: [String(studentId)] },
    state: 'PUBLISHED'
  };

  const maxRetry = 3;
  for (let t = 0; t < maxRetry; t++) {
    try {
      return Classroom.Courses.Announcements.create(data, String(courseId));
    } catch (e) {
      // 明示的にレート制限で拒否された場合だけ再試行する。内部エラーやタイムアウトは、
      // Classroom が投稿を受け付けた後に起きた可能性がある。
      if (t < maxRetry - 1 && isClassroomRateLimitError_(e)) {
        Utilities.sleep(500 * (t + 1));
        continue;
      }
      throw e;
    }
  }
}

function isClassroomRateLimitError_(error) {
  return /\bHTTP\s+429\b|\bRESOURCE_EXHAUSTED\b|\brateLimitExceeded\b|\buserRateLimitExceeded\b/.test(String(error));
}

function isClassroomRejectedError_(error) {
  return /\bHTTP\s+(?:400|401|403|404)\b|\b(?:INVALID_ARGUMENT|UNAUTHENTICATED|PERMISSION_DENIED|NOT_FOUND)\b/.test(String(error));
}

function sendMessages() {
  const message = withAppLock_(sendMessagesUnlocked_);
  safeAlert_(message);
}

function sendMessagesUnlocked_() {
  if(typeof assertNoManagedFormUpdate_==='function')assertNoManagedFormUpdate_();
  const config = getConfig_();
  validateAppConfig_(config);
  const formStatusCache = createFormStatusCache_();
  let scoringGuard, scoringGuardLoaded=false;
  const sendSheet = ensureSendSheet_(config);
  const lastRow = sendSheet.getLastRow();
  const lastCol = sendSheet.getLastColumn();
  if (lastRow <= 1) {
    return '送信シートにデータがありません。';
  }

  const values = sendSheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const headers = values[0].map(String);
  const headerMap = createHeaderMap_(headers);
  const requiredHeaders = ['元SS_ID', '元シート名', '元行番号', 'メールアドレス', '名前', 'コースID', 'studentId', '送信状態', config.replyBodyHeader];

  requiredHeaders.forEach(function(header) {
    if (!(header in headerMap)) {
      throw new Error('送信シートのヘッダに「' + header + '」が必要です。');
    }
  });

  const errorSheet = ensureErrorSheet_();
  let sentCount = 0;
  let errorCount = 0;
  let reviewCount = 0;
  const postRegistry = typeof getManagedRecords_ === 'function' ? {records:getManagedRecords_(),sheet:null} : null;
  const responseCounts = new Map();
  values.slice(1).forEach(function(row) {
    const key = makeResponseKey_(row[headerMap['元SS_ID']], row[headerMap['元シート名']], row[headerMap['元行番号']]);
    if (key) responseCounts.set(key, (responseCounts.get(key) || 0) + 1);
  });

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const status = String(row[headerMap['送信状態']] || '').trim();
    if (row.every(function(value) { return String(value == null ? '' : value).trim() === ''; })) continue;
    if (status === SEND_REVIEW_STATUS) {
      reviewCount++;
      continue;
    }
    if (status && status !== '未') continue;

    const srcSsId = String(row[headerMap['元SS_ID']] || '').trim();
    const srcSheetName = String(row[headerMap['元シート名']] || '').trim();
    const srcRow = Number(row[headerMap['元行番号']] || 0);
    const email = String(row[headerMap['メールアドレス']] || '').trim();
    const name = String(row[headerMap['名前']] || '').trim();
    const courseId = String(row[headerMap['コースID']] || '').trim();
    const studentId = String(row[headerMap['studentId']] || '').trim();
    const body = String(row[headerMap[config.replyBodyHeader]] || '').trim();
    const responseKey = makeResponseKey_(srcSsId, srcSheetName, srcRow);
    const statusCell = sendSheet.getRange(r + 1, headerMap['送信状態'] + 1);

    const missingFields = [];
    if (!responseKey) missingFields.push('有効な元SS_ID・元シート名・元行番号（2以上の整数）');
    if (!email) missingFields.push('メールアドレス');
    if (!name) missingFields.push('名前');
    if (!courseId) missingFields.push('コースID');
    if (!studentId) missingFields.push('studentId');
    if (!body) missingFields.push(config.replyBodyHeader);
    if (missingFields.length > 0) {
      sendSheet.getRange(r + 1, headerMap['送信状態'] + 1).setValue('エラー');
      errorSheet.appendRow([
        new Date(),
        SEND_SHEET_NAME,
        r + 1,
        email,
        'エラー',
        'CONFIG',
        '送信に必要な値が不足しています: ' + missingFields.join(', ')
      ]);
      if (responseKey) {
        updateFormStatus_(config, srcSsId, srcSheetName, srcRow, '送信エラー', formStatusCache);
      }
      errorCount++;
      continue;
    }

    if (responseCounts.get(responseKey) > 1) {
      statusCell.setValue(SEND_REVIEW_STATUS);
      errorSheet.appendRow([new Date(), SEND_SHEET_NAME, r + 1, email,
        SEND_REVIEW_STATUS, 'DUPLICATE_SOURCE',
        '同じ元回答の送信行が複数あります。Classroomの投稿と各行を確認してください。']);
      reviewCount++;
      continue;
    }

    if (!scoringGuardLoaded) {
      scoringGuard=scoringBuildReturnGuard_(formStatusCache);
      scoringGuardLoaded=true;
    }
    if (!scoringMayReturn_(scoringGuard,responseKey)) { reviewCount++;continue; }

    // 外部 API を呼ぶ前に処理位置を保存する。ここで停止した場合や結果が不明な場合、
    // 次の実行で再投稿してはならない。
    statusCell.setValue(SEND_REVIEW_STATUS);
    SpreadsheetApp.flush();
    let announcement;
    try {
      announcement = postAnnouncementIndividual_(courseId, studentId, name, body);
    } catch (e) {
      const msg = String(e);
      let errType = 'SEND_RESULT_UNKNOWN';
      let newStatus = SEND_REVIEW_STATUS;
      if (isClassroomRateLimitError_(e)) {
        errType = 'RETRYABLE';
        newStatus = '未';
      } else if (isClassroomRejectedError_(e)) {
        errType = 'FATAL';
        newStatus = 'エラー';
      }
      if (newStatus !== SEND_REVIEW_STATUS) statusCell.setValue(newStatus);
      errorSheet.appendRow([
        new Date(),
        SEND_SHEET_NAME,
        r + 1,
        email,
        newStatus,
        errType,
        msg
      ]);
      if (newStatus === SEND_REVIEW_STATUS) {
        reviewCount++;
      } else {
        errorCount++;
        if (newStatus === 'エラー') updateFormStatus_(config, srcSsId, srcSheetName, srcRow, '送信エラー', formStatusCache);
      }
      continue;
    }

    // 投稿の成功と Sheets への書き込み失敗は別の結果として扱う。
    // 投稿済みの行を自動送信の対象状態へ戻さない。
    try {
      if (postRegistry) recordManagedAnnouncement_(announcement, {courseId:courseId,name:name,sourceKey:responseKey},postRegistry);
      statusCell.setValue('済');
      SpreadsheetApp.flush();
    } catch (e) {
      errorSheet.appendRow([new Date(), SEND_SHEET_NAME, r + 1, email,
        SEND_REVIEW_STATUS, 'SEND_RECORD_ERROR',
        'Classroom投稿は成功。投稿ID: ' + String(announcement && announcement.id || '不明') +
        '。送信状態の保存を確認してください。' + String(e)]);
      reviewCount++;
      continue;
    }
    sentCount++;
    updateFormStatus_(config, srcSsId, srcSheetName, srcRow, '済', formStatusCache);
    Utilities.sleep(120);
  }

  return '送信処理が終了しました。\n' +
    '送信成功：' + sentCount + '件\n' +
    'エラー　：' + errorCount + '件（エラーシート参照）\n' +
    '要確認　：' + reviewCount + '件（送信確認待ち・エラーシート参照）';
}

function remindUngradedAndErrors() {
  return withAppLock_(remindUngradedAndErrorsUnlocked_);
}

function remindUngradedAndErrorsUnlocked_() {
  if(typeof assertNoManagedFormUpdate_==='function')assertNoManagedFormUpdate_();
  const config = getConfig_();
  validateAppConfig_(config);

  if (!Array.isArray(config.reminderTo) || config.reminderTo.length === 0) {
    throw new Error('リマインドメール送信先が設定されていません。管理画面の「フォームと通知先」で通知先を入力してください。');
  }

  const submissionSourceHeader = config.emailHeader;
  const scoreField = resolveScoreField_(config);

  if (!submissionSourceHeader || !scoreField) {
    throw new Error('リマインド判定に必要な項目設定が不足しています。');
  }

  const lines = [];
  let totalUngraded = 0;
  let totalPending = 0;
  let totalError = 0;

  config.formSources.forEach(function(source) {
    const formSs = SpreadsheetApp.openById(source.id);
    const fileName = formSs.getName();

    formSs.getSheets().forEach(function(sheet) {
      const sheetName = sheet.getName();
      if (isManagedInternalSheet_(sheetName, source.id) || !sheetMatchesPrefixes_(sheetName, config.formSheetNamePrefix)) return;

      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow <= 1 || lastCol === 0) return;

      const values = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
      const headers = values[0].map(String);
      const headerMap = createHeaderMap_(headers);

      if (!(submissionSourceHeader in headerMap) || !(scoreField.sourceHeader in headerMap)) {
        return;
      }

      const idxSubmission = headerMap[submissionSourceHeader];
      const idxScore = headerMap[scoreField.sourceHeader];
      const idxStatus = headerMap[config.formStatusHeader];

      let ungraded = 0;
      let pending = 0;
      let errorCnt = 0;

      for (let r = 1; r < values.length; r++) {
        const row = values[r];
        const submission = String(row[idxSubmission] || '').trim();
        const score = String(row[idxScore] || '').trim();
        const status = idxStatus != null ? String(row[idxStatus] || '').trim() : '';

        if (submission && !score) ungraded++;
        if (status === '準備〇') pending++;
        if (status === '準備×' || status === '送信エラー') errorCnt++;
      }

      if (ungraded > 0 || pending > 0 || errorCnt > 0) {
        lines.push(
          fileName + ' / ' + sheetName +
          '：未採点 ' + ungraded + '件, ' +
          '未送信 ' + pending + '件, ' +
          'エラー ' + errorCnt + '件'
        );
        totalUngraded += ungraded;
        totalPending += pending;
        totalError += errorCnt;
      }
    });
  });

  if (totalUngraded === 0 && totalPending === 0 && totalError === 0) {
    return;
  }

  let body = '';
  body += '【フォーム別の状況】\n';
  lines.forEach(function(line) {
    body += '・' + line + '\n';
  });
  body += '\n';
  body += '【全体集計】\n';
  body += '・未採点　：' + totalUngraded + '件\n';
  body += '・未送信　：' + totalPending + '件（送信シートは作成済み）\n';
  body += '・エラー　：' + totalError + '件（準備×／送信エラー）\n\n';
  body += '（このメールはGASの時間トリガーで自動送信されています）\n';

  GmailApp.sendEmail(
    (config.reminderTo || []).join(','),
    '授業まとめ／未採点・送信状況のリマインド（Form基準）',
    body
  );
}

function clearEvaluationData() {
  const ss = getAppSpreadsheet_();
  const sheet = ss.getSheetByName(EVAL_SHEET_NAME);
  if (!sheet) {
    safeAlert_('「評価データ」シートが存在しません。');
    return;
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1) {
    safeAlert_('「評価データ」にはクリアする行がありません。');
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    '確認',
    '「評価データ」シートの 2 行目以降をすべて削除します。\nこの操作は取り消しできません。\n実行しますか？',
    ui.ButtonSet.OK_CANCEL
  );
  if (result !== ui.Button.OK) return;

  withAppLock_(function() {
    const current = ss.getSheetByName(EVAL_SHEET_NAME);
    if (current && current.getLastRow() > 1) {
      current.getRange(2, 1, current.getLastRow() - 1, current.getLastColumn()).clearContent();
    }
  });
  safeAlert_('「評価データ」シートをクリアしました。');
}

function clearSendSheet() {
  const ss = getAppSpreadsheet_();
  const sheet = ss.getSheetByName(SEND_SHEET_NAME);
  if (!sheet) {
    safeAlert_('「送信シート」が存在しません。');
    return;
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1) {
    safeAlert_('「送信シート」にはクリアする行がありません。');
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    '確認',
    '「送信シート」の 2 行目以降をすべて削除します。\nこの操作は取り消せません。\nよろしいですか？',
    ui.ButtonSet.OK_CANCEL
  );
  if (result !== ui.Button.OK) return;

  withAppLock_(function() {
    const current = ss.getSheetByName(SEND_SHEET_NAME);
    if (current && current.getLastRow() > 1) {
      current.getRange(2, 1, current.getLastRow() - 1, current.getLastColumn()).clearContent();
    }
  });
  safeAlert_('「送信シート」をクリアしました。');
}

function clearErrorLog() {
  const ss = getAppSpreadsheet_();
  const sheet = ss.getSheetByName(ERROR_SHEET_NAME);
  if (!sheet) {
    safeAlert_('「エラー」シートは存在しません。');
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    '確認',
    '「エラー」シートを削除します。\nこの操作は取り消せません。\n実行しますか？',
    ui.ButtonSet.OK_CANCEL
  );
  if (result !== ui.Button.OK) return;

  withAppLock_(function() {
    const current = ss.getSheetByName(ERROR_SHEET_NAME);
    if (current) ss.deleteSheet(current);
  });
  safeAlert_('「エラー」シートを削除しました。');
}

// === 採点処理。scoring_tool の main 94ef234 (MIT) を元に調整 ===
// 採点用デジタル型紙 GAS サーバーサイド

/** デバッグモード（true のとき _debug をレスポンスに含める） */
const SCORING_DEBUG_MODE = false;
const SCORING_CONFIG_BACKUP_KEY = 'TURRET_SCORING_CONFIG_BACKUP';

/** 対象スプレッドシート ID（必要に応じて書き換え） */

/** 現在の採点対象スプレッドシートIDを取得（プロパティ優先、なければ定数） */

/** 設定用スプレッドシートIDを取得（コンテナ優先、なければScriptProperties） */
function scoringLegacyGetConfigSpreadsheetId_() { return getAppSpreadsheet_().getId(); }

/** URLまたはIDからスプレッドシートIDを抽出 */

/** 設定のデフォルト値 */
function scoringLegacyGetDefaultConfig_() {
  return {
    ruleTemplateId: '',
    sheetName: '', // 接続先はturretの登録を使用
    lessonDateHeader: null, // null: ひな形の授業日を検出、空文字: 無効
    startRow: 2,
    // 表示ヘッダ行（採点画面上段の見出しに使用）。未指定は 1 行目。
    headerRow: 1,
    nameCol: 'B',
    nameHeader: '',
    displayCols: ['A', 'B', 'C'],
    displayHeaders: ['', '', '', '', ''],
    charCountCols: [],
    scoreCols: ['', '', '', '', ''],
    scoreHeaders: ['', '', '', '', ''],
    // 旧互換用。v1.0.0候補の公開仕様では使用しない。
    mergeRules: {},
    // スコア入力ロック（列チェック）
    colChecks: [false, false, false, false, false],
    // 末行から開始
    startFromLastRow: false,
    // UI上部の項目名（スロット1〜5）。未設定なら空文字。
    // ※旧キー scoreLabels は互換性なし（scoringLegacyAssertNoDeprecatedKeys_で拒否）
    slotLabels: ['', '', '', '', ''],
    commentCol: '',
    commentHeader: '',
    /**
     * ルール出力先（講評/改善点など）のスロット（最大3）
     * - target: `_rules.target` のサブグループ名
     * - col: 書き込み先列（A1形式の列記号）
     * - enabled: true のとき読み取り/書き込み/編集対象とする（UI側の「編集有効化」に連動）
     *
     * 後方互換:
     * - 旧configは commentCol のみを保持している想定
     * - UI側で commentCol → slot1 へ移行・反映する
     */
    ruleOutputSlots: [],
    flushShortcut: 'Ctrl+S',
    commentTemplate: {
      type: 'byTotal',
      highThreshold: 24,
      midThreshold: 12
    }
  };
}

/**
 * ruleOutputSlots を正規化（最大3、型ブレ対策）
 * @param {*} slots
 * @return {Array<{target:string,col:string,enabled:boolean}>}
 */
function scoringLegacyNormalizeRuleOutputSlots_(slots) {
  var out = [];
  if (!Array.isArray(slots)) {
    return out;
  }
  for (var i = 0; i < slots.length && out.length < 3; i++) {
    var s = slots[i];
    if (!s || typeof s !== 'object') continue;
    var target = String(s.target == null ? '' : s.target).trim();
    var col = String(s.col == null ? '' : s.col).trim().toUpperCase();
    var header = String(s.header == null ? '' : s.header).trim();
    var enabled = !!s.enabled;
    // 空スロットは除外（enabledだけtrueでtarget/colが空などはUIで警告する想定だが、保存自体は許可）
    if (!target && !col && !enabled && !header) continue;
    out.push({ target: target, col: col, enabled: enabled, header: header });
  }
  return out;
}

/** HTML を返す */

/** ScriptProperties 取得 */
function scoringLegacyGetScriptProps_() {
  return PropertiesService.getScriptProperties();
}

/** 列記号 "AA" → 1始まり列番号 */
function scoringLegacyColA1ToIndex_(col) {
  col = String(col || '').toUpperCase().trim();
  if (!col) throw new Error('列記号が空です');
  var sum = 0;
  for (var i = 0; i < col.length; i++) {
    var c = col.charCodeAt(i);
    if (c < 65 || c > 90) throw new Error('不正な列記号: ' + col);
    sum = sum * 26 + (c - 64);
  }
  return sum;
}

/** 列番号 (1始まり) → 列記号 "AA" */
function scoringLegacyIndexToColA1_(index) {
  var n = Number(index);
  if (!isFinite(n) || n < 1) throw new Error('列番号が不正です: ' + index);
  var s = '';
  while (n > 0) {
    var mod = (n - 1) % 26;
    s = String.fromCharCode(65 + mod) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function scoringLegacyNormalizeStringArrayFixed_(arr, len) {
  var out = [];
  if (Array.isArray(arr)) {
    out = arr.map(function (v) { return v == null ? '' : String(v); });
  } else if (typeof arr === 'string') {
    out = [String(arr)];
  }
  if (len) {
    if (out.length > len) out = out.slice(0, len);
    while (out.length < len) out.push('');
  }
  return out;
}

function scoringLegacyNormalizeColArrayFixed_(arr, len) {
  var out = [];
  if (Array.isArray(arr)) {
    out = arr.map(function (v) { return v == null ? '' : String(v).trim().toUpperCase(); });
  } else if (typeof arr === 'string') {
    out = [String(arr).trim().toUpperCase()];
  }
  if (len) {
    if (out.length > len) out = out.slice(0, len);
    while (out.length < len) out.push('');
  }
  return out;
}

function scoringLegacyParseHeaderLabel_(label) {
  var s = String(label == null ? '' : label).trim();
  if (!s) return null;
  var m = s.match(/^(.*)\s*\((\d+)\)$/);
  if (m) {
    var base = String(m[1] || '').trim();
    var nth = Number(m[2] || 0);
    if (!base || !isFinite(nth) || nth < 1) return null;
    return { base: base, nth: nth };
  }
  return { base: s, nth: 1 };
}

function scoringLegacyBuildHeaderMeta_(headerValues) {
  var map = {};
  var labelByIndex = {};
  var list = [];
  var counts = {};
  var values = Array.isArray(headerValues) ? headerValues : [];
  for (var i = 0; i < values.length; i++) {
    var base = String(values[i] == null ? '' : values[i]).trim();
    if (!base) continue;
    counts[base] = (counts[base] || 0) + 1;
    var nth = counts[base];
    var label = base + (nth > 1 ? ' (' + nth + ')' : '');
    var colIndex = i + 1;
    if (!map[base]) map[base] = [];
    map[base].push(colIndex);
    labelByIndex[colIndex] = label;
    list.push({ col: scoringLegacyIndexToColA1_(colIndex), header: base, label: label });
  }
  return { map: map, labelByIndex: labelByIndex, list: list };
}

function scoringLegacyResolveColIndexByHeaderLabel_(label, headerMap) {
  if (!label || !headerMap) return null;
  var parsed = scoringLegacyParseHeaderLabel_(label);
  if (!parsed) return null;
  var list = headerMap[parsed.base];
  if (!list || !list.length) return null;
  var idx = parsed.nth || 1;
  if (idx < 1 || idx > list.length) return null;
  return list[idx - 1];
}

function scoringLegacyApplyHeaderResolutionToConfig_(cfg, headerMeta) {
  if (!cfg || !headerMeta) return;
  var headerMap = headerMeta.map || {};
  var labelByIndex = headerMeta.labelByIndex || {};

  cfg.displayCols = scoringLegacyNormalizeColArrayFixed_(cfg.displayCols, 5);
  cfg.displayHeaders = scoringLegacyNormalizeStringArrayFixed_(cfg.displayHeaders, 5);
  cfg.scoreCols = scoringLegacyNormalizeColArrayFixed_(cfg.scoreCols, 5);
  cfg.scoreHeaders = scoringLegacyNormalizeStringArrayFixed_(cfg.scoreHeaders, 5);

  cfg.nameHeader = String(cfg.nameHeader || '').trim();
  cfg.commentHeader = String(cfg.commentHeader || '').trim();

  function resolveCol_(label) {
    var idx = scoringLegacyResolveColIndexByHeaderLabel_(label, headerMap);
    return idx ? scoringLegacyIndexToColA1_(idx) : '';
  }

  if (cfg.nameHeader) {
    var nc = resolveCol_(cfg.nameHeader);
    if (nc) cfg.nameCol = nc;
  }
  for (var i = 0; i < cfg.displayHeaders.length; i++) {
    var dc = cfg.displayHeaders[i] ? resolveCol_(cfg.displayHeaders[i]) : '';
    if (dc) cfg.displayCols[i] = dc;
  }
  for (var s = 0; s < cfg.scoreHeaders.length; s++) {
    var sc = cfg.scoreHeaders[s] ? resolveCol_(cfg.scoreHeaders[s]) : '';
    if (sc) cfg.scoreCols[s] = sc;
  }
  if (cfg.commentHeader) {
    var cc = resolveCol_(cfg.commentHeader);
    if (cc) cfg.commentCol = cc;
  }

  if (Array.isArray(cfg.ruleOutputSlots)) {
    for (var r = 0; r < cfg.ruleOutputSlots.length; r++) {
      var slot = cfg.ruleOutputSlots[r];
      if (!slot || typeof slot !== 'object') continue;
      var header = String(slot.header || '').trim();
      if (!header) continue;
      var rc = resolveCol_(header);
      if (rc) slot.col = rc;
    }
  }

  function fillHeaderByCol_(col, current) {
    if (current) return current;
    if (!col) return current;
    try {
      var idx = scoringLegacyColA1ToIndex_(col);
      var label = labelByIndex[idx];
      if (label) return label;
    } catch (e) {}
    return current;
  }

  cfg.nameHeader = fillHeaderByCol_(cfg.nameCol, cfg.nameHeader);
  for (var i2 = 0; i2 < cfg.displayCols.length; i2++) {
    cfg.displayHeaders[i2] = fillHeaderByCol_(cfg.displayCols[i2], cfg.displayHeaders[i2]);
  }
  for (var s2 = 0; s2 < cfg.scoreCols.length; s2++) {
    cfg.scoreHeaders[s2] = fillHeaderByCol_(cfg.scoreCols[s2], cfg.scoreHeaders[s2]);
  }
  cfg.commentHeader = fillHeaderByCol_(cfg.commentCol, cfg.commentHeader);

  if (Array.isArray(cfg.ruleOutputSlots)) {
    for (var r2 = 0; r2 < cfg.ruleOutputSlots.length; r2++) {
      var slot2 = cfg.ruleOutputSlots[r2];
      if (!slot2 || typeof slot2 !== 'object') continue;
      slot2.header = fillHeaderByCol_(slot2.col, String(slot2.header || '').trim());
    }
  }
}

function scoringLegacyApplyHeaderResolutionForSheet_(cfg, sheet) {
  if (!cfg || !sheet) return;
  var headerRow = Number(cfg && cfg.headerRow);
  if (!isFinite(headerRow) || headerRow < 1) headerRow = 1;
  var lastCol = sheet.getLastColumn();
  var headerValuesAll = lastCol > 0 ? sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0] : [];
  var headerMeta = scoringLegacyBuildHeaderMeta_(headerValuesAll);
  cfg.ruleOutputSlots = scoringLegacyNormalizeRuleOutputSlots_(cfg.ruleOutputSlots);
  scoringLegacyApplyHeaderResolutionToConfig_(cfg, headerMeta);
}

/**
 * HTML Service へ返すためのセル値正規化
 * - Date はそのままだと google.script.run の戻り値として不正になるため文字列化
 * - null/undefined は空文字
 */
function scoringLegacyToClientCellValue_(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) {
    var dt = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    return / 00:00:00$/.test(dt) ? dt.replace(/ 00:00:00$/, '') : dt;
  }
  return v;
}

/**
 * 採点列（scoreCols）の重複を禁止する。
 * 空欄は未使用枠として無視する。
 * @param {Array<string>} scoreCols
 */
function scoringLegacyAssertUniqueScoreCols_(scoreCols) {
  var seen = {};
  var dups = [];
  var cols = Array.isArray(scoreCols) ? scoreCols : [];
  for (var i = 0; i < cols.length; i++) {
    var c = String(cols[i] == null ? '' : cols[i]).trim().toUpperCase();
    if (!c) continue;
    if (seen[c]) {
      if (dups.indexOf(c) < 0) dups.push(c);
    } else {
      seen[c] = true;
    }
  }
  if (dups.length) {
    throw new Error('採点値入力列が重複しています: ' + dups.join(', '));
  }
}

/** シート取得（なければエラー） */
function scoringLegacyGetSheetByName_(target) { return resolveScoringTarget_(target).sheet; }

/** 旧設定キー検出（互換性なし） */
function scoringLegacyAssertNoDeprecatedKeys_(obj) {
  if (!obj || typeof obj !== 'object') return;
  var deprecatedKeys = ['ruleInputCols', 'ruleInputSlots', 'enableTotal', 'scoreLabels'];
  var foundDeprecated = [];
  for (var i = 0; i < deprecatedKeys.length; i++) {
    if (obj.hasOwnProperty(deprecatedKeys[i])) {
      foundDeprecated.push(deprecatedKeys[i]);
    }
  }
  if (foundDeprecated.length > 0) {
    throw new Error('旧設定キーが検出されました。再設定が必要です。検出されたキー: ' + foundDeprecated.join(', '));
  }
}

/** 設定取得（デフォルトとマージ） */
function scoringLegacyApiGetConfig_() {
  var props = scoringLegacyGetScriptProps_();
  var json = props.getProperty('TURRET_SCORING_CONFIG');
  var base = scoringLegacyGetDefaultConfig_();
  if (!json) return base;
  try {
    var saved = JSON.parse(json);
    // 浅いマージ（必要十分）
    for (var k in base) {
      if (saved.hasOwnProperty(k)) base[k] = saved[k];
    }
    // ruleTemplateId が存在しない場合は空文字として扱う（未選択状態）
    if (!saved.hasOwnProperty('ruleTemplateId')) {
      base.ruleTemplateId = '';
    }
    // commentTemplate も中身をマージ
    if (saved.commentTemplate) {
      base.commentTemplate = base.commentTemplate || {};
      var t = saved.commentTemplate;
      for (var kk in t) {
        if (t.hasOwnProperty(kk)) base.commentTemplate[kk] = t[kk];
      }
    }
    // 8個設定を5個にトリム（互換性のため）
    if (Array.isArray(base.scoreCols) && base.scoreCols.length === 8) {
      base.scoreCols = base.scoreCols.slice(0, 5);
    }

    // 旧キー検出でエラー（互換性なし）
    scoringLegacyAssertNoDeprecatedKeys_(saved);

    return base;
  } catch (e) {
    throw new Error('保存済みの採点設定を読み取れません。設定JSONを確認・復元してください。');
  }
}

/** 設定保存 */
function scoringLegacyApiSetConfig_(config) {
  if (!config) throw new Error('config が空です');

  // 旧キー検出でエラー（互換性なし）
  scoringLegacyAssertNoDeprecatedKeys_(config);

  // 必要項目の正規化
  var out = scoringLegacyGetDefaultConfig_();
  var ruleTemplateIdRaw = String(config.ruleTemplateId == null ? '' : config.ruleTemplateId).trim();
  // ruleTemplateId が空文字の場合は保存しない（未選択状態）
  if (ruleTemplateIdRaw) {
    out.ruleTemplateId = ruleTemplateIdRaw;
  }
  // 空文字の場合は out.ruleTemplateId を設定しない（デフォルト値も使わない）
  out.sheetName = '';
  out.lessonDateHeader = config.lessonDateHeader == null ? null : String(config.lessonDateHeader).trim();
  out.startRow = Number(config.startRow) || 2;
  // headerRow: 未指定/不正は 1。開始行とは独立。
  out.headerRow = Number(config.headerRow);
  if (!isFinite(out.headerRow) || out.headerRow < 1) out.headerRow = 1;
  out.nameCol = (config.nameCol || '').toString().trim().toUpperCase();
  out.nameHeader = String(config.nameHeader || '').trim();

  // カンマ区切りの入力（半角, 全角， 日本語、）を配列へ正規化
  function splitList_(s) {
    var str = String(s == null ? '' : s);
    if (!str.trim()) return [];
    return str.split(/[,，、]/).map(function (x) { return String(x || '').trim(); }).filter(function (x) { return x; });
  }

  function parseColArray(val, expected) {
    var out = [];
    if (Array.isArray(val)) {
      val.forEach(function (v) {
        out = out.concat(splitList_(v));
      });
      return out;
    }
    if (typeof val === 'string') {
      return splitList_(val);
    }
    return [];
  }

  function parseStringArray(val) {
    var out = [];
    if (Array.isArray(val)) {
      out = val.map(function (v) { return v == null ? '' : String(v); });
      return out;
    }
    if (typeof val === 'string') {
      return splitList_(val);
    }
    return [];
  }

  function parseBoolArray(val, expected) {
    var out = [];
    if (Array.isArray(val)) {
      for (var i = 0; i < val.length; i++) {
        out.push(!!val[i]);
      }
    }
    if (expected) {
      while (out.length < expected) out.push(false);
      if (out.length > expected) out = out.slice(0, expected);
    }
    return out;
  }

  function normalizeMergeRules_(mr) {
    var outMr = {};
    if (!mr || typeof mr !== 'object') return outMr;
    for (var kk in mr) {
      if (!mr.hasOwnProperty(kk)) continue;
      var key = String(kk || '').trim().toUpperCase();
      if (!key) continue;
      var val = mr[kk];
      if (val == null) continue;
      var rule = String(val).trim();
      if (!rule) continue;
      outMr[key] = rule;
    }
    return outMr;
  }

  out.displayCols = parseColArray(config.displayCols, null).map(function (c) { return c.toUpperCase(); });
  out.displayHeaders = parseStringArray(config.displayHeaders);
  // charCountCols: DISPLAY_COLS と同じパーサーでOK（カンマ区切り、空は除外、大文字化）
  // A方式なので、displayColsに無い列を弾く（server側で防御）
  var charCountColsRaw = parseColArray(config.charCountCols, null).map(function (c) { return c.toUpperCase(); });
  out.charCountCols = charCountColsRaw.filter(function (c) { return c && out.displayCols.indexOf(c) >= 0; });
  out.scoreCols = (Array.isArray(config.scoreCols) ? config.scoreCols.map(function(c){return String(c || '').trim();}) : parseColArray(config.scoreCols,5)).map(function(c){return c.toUpperCase();});
  out.scoreHeaders = parseStringArray(config.scoreHeaders);
  out.colChecks = parseBoolArray(config.colChecks, 5);
  out.startFromLastRow = !!config.startFromLastRow;
  // UI上部の項目名（スロット1〜5）
  out.slotLabels = parseStringArray(config.slotLabels);
  out.commentCol = (config.commentCol || '').toString().trim().toUpperCase();
  out.commentHeader = String(config.commentHeader || '').trim();
  out.ruleOutputSlots = scoringLegacyNormalizeRuleOutputSlots_(config.ruleOutputSlots);
  out.flushShortcut = String(config.flushShortcut || out.flushShortcut || 'Ctrl+S').trim();

  // 旧互換用。v1.0.0候補の公開仕様では使用しない。
  out.mergeRules = normalizeMergeRules_(config.mergeRules);

  // displayCols / displayHeaders を5に正規化
  if (out.displayCols.length > 5) {
    out.displayCols = out.displayCols.slice(0, 5);
  }
  while (out.displayCols.length < 5) {
    out.displayCols.push('');
  }
  if (out.displayHeaders.length > 5) {
    out.displayHeaders = out.displayHeaders.slice(0, 5);
  }
  while (out.displayHeaders.length < 5) {
    out.displayHeaders.push('');
  }

  // 正規化：可変長（0..5）を受け入れ、5にパディング
  // 過剰は切り捨て（8個互換含む）
  if (out.scoreCols.length > 5) {
    out.scoreCols = out.scoreCols.slice(0, 5);
  }

  // 不足は空で補完
  while (out.scoreCols.length < 5) {
    out.scoreCols.push('');
  }
  // 統合ルールは廃止。採点列の重複は保存時エラー。
  scoringLegacyAssertUniqueScoreCols_(out.scoreCols);

  // scoreHeaders も 5 に正規化
  if (out.scoreHeaders.length > 5) {
    out.scoreHeaders = out.scoreHeaders.slice(0, 5);
  }
  while (out.scoreHeaders.length < 5) {
    out.scoreHeaders.push('');
  }

  // colChecks も 5 に正規化（不足はfalseで補完、過剰は切り捨て）
  if (out.colChecks.length > 5) {
    out.colChecks = out.colChecks.slice(0, 5);
  }
  while (out.colChecks.length < 5) {
    out.colChecks.push(false);
  }

  // slotLabels も 5 に正規化（不足は空で補完、過剰は切り捨て）
  if (out.slotLabels.length > 5) {
    out.slotLabels = out.slotLabels.slice(0, 5);
  }
  while (out.slotLabels.length < 5) {
    out.slotLabels.push('');
  }

  // commentTemplate（合計点3段階）の検証・デフォルト補完
  var def = scoringLegacyGetDefaultConfig_().commentTemplate;
  var tmpl = config.commentTemplate || {};
  out.commentTemplate = {
    type: tmpl.type || def.type,
    highThreshold: Number(tmpl.highThreshold || def.highThreshold),
    midThreshold: Number(tmpl.midThreshold || def.midThreshold)
  };

  var props = scoringLegacyGetScriptProps_();

  // 未選択も空文字で保持し、エクスポート・バックアップの形式を揃える。

  props.setProperty('TURRET_SCORING_CONFIG', JSON.stringify(out));
  return out;
}

/**
 * slotLabels（UI上部の項目名）だけを更新する
 * - 設定モーダルの未保存変更を巻き込まないための専用API
 * @param {Array<string>|string} slotLabels - 5枠分（配列推奨、文字列の場合はカンマ区切り）
 * @return {Object} 更新後のconfig
 */
function scoringLegacyApiSetSlotLabels_(slotLabels) {
  var props = scoringLegacyGetScriptProps_();
  var existingJson = props.getProperty('TURRET_SCORING_CONFIG');
  var existing = null;
  try {
    existing = existingJson ? JSON.parse(existingJson) : null;
  } catch (e) {
    existing = null;
  }
  if (!existing || typeof existing !== 'object') {
    existing = {};
  }
  // 旧キー検出でエラー（互換性なし）
  scoringLegacyAssertNoDeprecatedKeys_(existing);

  // slotLabels のみ更新（他のキーは一切触らない）
  if (Array.isArray(slotLabels)) {
    existing.slotLabels = slotLabels.map(function (v) { return v == null ? '' : String(v); });
  } else {
    existing.slotLabels = String(slotLabels == null ? '' : slotLabels).split(',').map(function (s) { return String(s).trim(); });
  }
  // 長さ5に正規化
  if (existing.slotLabels.length > 5) existing.slotLabels = existing.slotLabels.slice(0, 5);
  while (existing.slotLabels.length < 5) existing.slotLabels.push('');

  props.setProperty('TURRET_SCORING_CONFIG', JSON.stringify(existing));
  return scoringLegacyApiGetConfig_();
}

/** シート名一覧 */

/** スプレッドシートURL/IDを設定 */

/** 採点対象スプレッドシートIDをクリア（未設定状態に戻す） */

/**
 * 採点対象スプシ（TARGET_SPREADSHEET_ID）の編集URLを返す
 * - 可能なら指定シート（なければ config.sheetName）の gid を付与
 * @param {{sheetName?:string}=} params
 * @return {string}
 */

/**
 * ルール編集用のURLを返す（設定用スプシ）
 * - まずはスプレッドシートトップURL
 * - 可能なら `_rules`（なければ `_templates`）の gid を付与
 */
function scoringLegacyApiGetRulesEditUrl_() {
  var ssId = scoringLegacyGetConfigSpreadsheetId_();

  // scoringLegacyGetConfigSpreadsheetId_()がnullを返した場合、コンテナスプシを直接取得を試みる
  if (!ssId) {
    try {
      var active = SpreadsheetApp.getActiveSpreadsheet();
      if (active) {
        ssId = active.getId();
      }
    } catch (e) {
      // コンテナスプシが取得できない場合はエラー
      throw new Error('設定用スプレッドシートIDが取得できませんでした');
    }
  }

  if (!ssId) {
    throw new Error('設定用スプレッドシートIDが取得できませんでした');
  }

  var baseUrl = 'https://docs.google.com/spreadsheets/d/' + ssId + '/edit';
  try {
    var ss = SpreadsheetApp.openById(ssId);
    var sheet = ss.getSheetByName('_rules') || ss.getSheetByName('_templates');
    if (sheet) {
      return baseUrl + '#gid=' + sheet.getSheetId();
    }
  } catch (e) {
    // 権限や一時エラー時はトップURLにフォールバック
  }
  return baseUrl;
}

/**
 * 初期データ取得
 * @param {{sheetName:string,startRow:number}} params
 */
function scoringLegacyApiGetInitData_(params,snapshot,config) {
  var cfg = config || scoringLegacyApiGetConfig_();
  var sheetName = params && params.sheetName ? params.sheetName : (cfg.sheetName || '');
  var startRow = params && params.startRow ? Number(params.startRow) : cfg.startRow || 2;
  if (startRow < 1) startRow = 1;

  // 表示ヘッダ行（未指定/不正は 1）。開始行とは独立。
  var headerRow = Number(cfg && cfg.headerRow);
  if (!isFinite(headerRow) || headerRow < 1) headerRow = 1;
  // 応答内の config と整合させる（旧設定が headerRow を持たない場合の互換・表示用）
  cfg.headerRow = headerRow;

  var sheet = snapshot ? snapshot.sheet : scoringLegacyGetSheetByName_(sheetName);
  var lastRow = snapshot ? snapshot.lastRow : sheet.getLastRow();
  var lastCol = snapshot ? snapshot.lastCol : sheet.getLastColumn();
  var headerValuesAll = snapshot ? snapshot.values[headerRow-1] || [] : lastCol > 0 ? sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0] : [];
  var headerMeta = scoringLegacyBuildHeaderMeta_(headerValuesAll);

  // ルール出力枠を正規化してから、ヘッダ名→列記号の解決を行う
  cfg.ruleOutputSlots = scoringLegacyNormalizeRuleOutputSlots_(cfg.ruleOutputSlots);
  scoringLegacyApplyHeaderResolutionToConfig_(cfg, headerMeta);

  if (lastRow < startRow) {
    return {
      sheetName: sheet.getName(),
      startRow: startRow,
      headerRow: headerRow,
      rows: [],
      config: cfg,
      headers: [],
      allHeaders: headerMeta && headerMeta.list ? headerMeta.list : []
    };
  }

  // 必要列のインデックス算出
  var indices = [];
  var colMap = { nameCol: null, display: [], score: [], commentCol: null, ruleOutputs: [] };

  if (cfg.nameCol) {
    var idxName = scoringLegacyColA1ToIndex_(cfg.nameCol);
    colMap.nameCol = idxName;
    indices.push(idxName);
  }
  cfg.displayCols.forEach(function (c) {
    if (c && c.trim()) {
      var idx = scoringLegacyColA1ToIndex_(c);
      colMap.display.push(idx);
      indices.push(idx);
    } else {
      // 未使用枠はnullとして記録（位置合わせ用）
      colMap.display.push(null);
    }
  });
  cfg.scoreCols.forEach(function (c) {
    if (c && c.trim()) {
      var idx = scoringLegacyColA1ToIndex_(c);
      colMap.score.push(idx);
      indices.push(idx);
    } else {
      // 未使用枠（空文字）はnullとして記録
      colMap.score.push(null);
    }
  });
  if (cfg.commentCol) {
    var idxCom = scoringLegacyColA1ToIndex_(cfg.commentCol);
    colMap.commentCol = idxCom;
    indices.push(idxCom);
  }

  // ルール出力スロット（最大3）。col 指定があるものを読み取る。
  // - enabled は「編集可否（UI表示）」であり、データの読み取り/書き込み可否とは切り離す
  // ※返却する row.ruleOutputs はスロット数（最大3）に合わせて配列化し、未使用は '' とする。
  var ruleSlots = scoringLegacyNormalizeRuleOutputSlots_(cfg.ruleOutputSlots);
  // cfgに正規化結果を反映（UI側の型ブレ対策）
  cfg.ruleOutputSlots = ruleSlots;
  for (var rs = 0; rs < 3; rs++) {
    if (ruleSlots[rs] && ruleSlots[rs].col) {
      var idxRule = scoringLegacyColA1ToIndex_(ruleSlots[rs].col);
      colMap.ruleOutputs.push(idxRule);
      indices.push(idxRule);
    } else {
      colMap.ruleOutputs.push(null);
    }
  }

  indices.sort(function (a, b) { return a - b; });

  // 空のindicesリストに対応（すべてのスコア列が未使用等）
  if (indices.length === 0) {
    return {
      sheetName: sheet.getName(),
      startRow: startRow,
      headerRow: headerRow,
      rows: [],
      config: cfg,
      headers: [],
      allHeaders: headerMeta && headerMeta.list ? headerMeta.list : []
    };
  }

  var minCol = indices[0];
  var maxCol = indices[indices.length - 1];
  var width = maxCol - minCol + 1;

  var headerValues = headerValuesAll.slice(minCol - 1, minCol - 1 + width);
  var dataValues = snapshot ? snapshot.values.slice(startRow-1).map(function(row){return row.slice(minCol-1,maxCol);}) : sheet.getRange(startRow, minCol, lastRow - startRow + 1, width).getValues();

  // 末尾の空行をトリム（すべて空の行を後ろから削る）
  var end = dataValues.length;
  outer: while (end > 0) {
    var row = dataValues[end - 1];
    for (var i = 0; i < row.length; i++) {
      if (row[i] !== '' && row[i] !== null) break outer;
    }
    end--;
  }
  dataValues = dataValues.slice(0, end);

  function pickCols(row, idxArray) {
    return idxArray.map(function (cIdx) {
      // nullは未使用枠として空文字を返す
      if (cIdx === null) return '';
      // 0 / false を空扱いにしない（null/undefinedのみ空）
      var v = row[cIdx - minCol];
      return scoringLegacyToClientCellValue_(v);
    });
  }

  var effectiveDisplayIdxs = [];
  var effectiveDisplayCols = [];
  for (var d = 0; d < colMap.display.length; d++) {
    if (colMap.display[d] != null) {
      effectiveDisplayIdxs.push(colMap.display[d]);
      effectiveDisplayCols.push(cfg.displayCols[d]);
    }
  }

  var rows = dataValues.map(function (row, i) {
    var rowNumber = startRow + i;
    var name = '';
    if (colMap.nameCol) {
      var vName = row[colMap.nameCol - minCol];
      name = scoringLegacyToClientCellValue_(vName);
    }
    var displays = pickCols(row, effectiveDisplayIdxs);
    var scores = pickCols(row, colMap.score);
    var comment = '';
    if (colMap.commentCol) {
      var vCom = row[colMap.commentCol - minCol];
      comment = scoringLegacyToClientCellValue_(vCom);
    }
    var ruleOutputs = pickCols(row, colMap.ruleOutputs);
    // 後方互換: commentCol 未指定だが ruleOutputs[0] がある場合は comment に反映
    if (!colMap.commentCol && ruleOutputs && ruleOutputs.length) {
      comment = ruleOutputs[0];
    }
    var rowData = {
      rowNumber: rowNumber,
      name: name,
      displays: displays,
      scores: scores,
      comment: comment,
      ruleOutputs: ruleOutputs
    };
    return rowData;
  });

  // 表示用ヘッダ（DISPLAY_COLS 分だけ）
  var headers = effectiveDisplayCols.map(function (c, idx) {
    var colIdx = effectiveDisplayIdxs[idx];
    var headerText = '';
    if (colIdx != null) {
      headerText = scoringLegacyToClientCellValue_(headerValues[colIdx - minCol]);
    }
    return { col: c, header: headerText };
  });

  return {
    sheetName: sheet.getName(),
    startRow: startRow,
    headerRow: headerRow,
    rows: rows,
    config: cfg,
    headers: headers,
    allHeaders: headerMeta && headerMeta.list ? headerMeta.list : []
  };
}

/**
 * 1行保存
 * @param {{sheetName:string,rowNumber:number,scores:*,comment:string}} payload
 */

/**
 * 複数行保存（連続行をまとめて送っても、行単位で安全に処理）
 * @param {{sheetName:string,rows:Array<{rowNumber:number,scores:*,comment?:string,ruleOutputs?:*}>}} payload
 */

/**
 * 1行分のpayloadを正規化
 * @param {{rowNumber:number,scores:*,comment?:string,ruleOutputs?:*}} payload
 * @param {*} cfg
 * @return {{rowNumber:number,scores:Array,comment:string,ruleOutputs:Array<string>}}
 */
function scoringLegacyNormalizeRowPayload_(payload, cfg) {
  if (!payload) throw new Error('row payload が空です');
  var rowNumber = Number(payload.rowNumber);
  if (!rowNumber || rowNumber < 1) {
    throw new Error('rowNumber が不正です: ' + payload.rowNumber);
  }

  var scores = payload.scores || [];
  if (!Array.isArray(scores) || scores.length !== cfg.scoreCols.length) {
    throw new Error('scores の配列長が SCORE_COLS と一致しません');
  }

  // 後方互換: payload.comment は従来通り受け入れる
  // 新仕様: payload.ruleOutputs（最大3）を受け取り、cfg.ruleOutputSlots に従って書き込む
  var ruleOutputs = Array.isArray(payload.ruleOutputs) ? payload.ruleOutputs : null;
  var comment = payload.comment != null ? String(payload.comment) : '';
  if (ruleOutputs && ruleOutputs.length) {
    // payload.comment が未指定でも slot1 を comment として扱う（既存UI互換）
    if (payload.comment == null) {
      comment = ruleOutputs[0] != null ? String(ruleOutputs[0]) : '';
    }
  }

  return {
    rowNumber: rowNumber,
    scores: scores,
    comment: comment,
    ruleOutputs: ruleOutputs
  };
}

/** 保存リクエストごとに、共通の列設定を一度だけ確定する。 */
function scoringColumnPlan_(cfg) {
  scoringLegacyAssertUniqueScoreCols_(cfg.scoreCols);
  return {
    scores:cfg.scoreCols.map(function(col){return col && col.trim() ? scoringLegacyColA1ToIndex_(col) : 0;}),
    comment:cfg.commentCol ? scoringLegacyColA1ToIndex_(cfg.commentCol) : 0,
    outputs:scoringLegacyNormalizeRuleOutputSlots_(cfg.ruleOutputSlots).map(function(slot){return slot.col ? scoringLegacyColA1ToIndex_(slot.col) : 0;})
  };
}

function scoringRowCells_(plan,row) {
  const cells=new Map();
  plan.scores.forEach(function(col,i){
    const value=row.scores[i];
    if(col && value!=='' && value!=null && String(value).trim()!=='')cells.set(col,value);
  });
  if(plan.comment)cells.set(plan.comment,row.comment);
  plan.outputs.forEach(function(col,i){
    if(!col)return;
    const value=row.ruleOutputs && i<row.ruleOutputs.length ? row.ruleOutputs[i] : i===0 ? row.comment : '';
    cells.set(col,value==null ? '' : String(value));
  });
  return Array.from(cells,function(pair){return {col:pair[0],value:pair[1]};}).sort(function(a,b){return a.col-b.col;});
}

function scoringLegacyApiExportConfig_() {
  var cfg = scoringLegacyApiGetConfig_();
  // 未保存時の表示列3件も、読み込み形式の5枠へ空欄で補う。保存先は変更しない。
  cfg.displayCols = cfg.displayCols.slice();
  while (cfg.displayCols.length < 5) cfg.displayCols.push('');
  return {
    meta: {
      version: 'v0.1.0',
      exportedAt: new Date().toISOString(),
      type: 'screen-config'
    },
    config: cfg
  };
}

// -----------------------------
// テンプレパック機能
// -----------------------------

const SCORING_TEMPLATE_PACK_META_TYPE = 'scoring-tool-template-pack';
const SCORING_TEMPLATE_PACK_VERSION = '0.1.0';

function scoringLegacyGetTemplatePackTemplateHeaderCandidates_() {
  return {
    templateId: ['templateid', 'template'],
    name: ['name', 'label', 'title', 'displayname'],
    enabled: ['enabled', 'enable', 'active', 'isenabled', 'isactive']
  };
}

function scoringLegacyGetTemplatePackRuleHeaderCandidates_() {
  return {
    templateId: ['templateid', 'template'],
    target: ['target', 'dest', 'group', 'outputtarget'],
    priority: ['priority', 'prio', 'order', 'rank'],
    whenExpr: ['whenexpr', 'when', 'expr', 'condition', 'if'],
    message: ['message', 'text', 'output', 'result', 'comment'],
    enabled: ['enabled', 'enable', 'active', 'isenabled', 'isactive'],
    visible: ['visible', 'show', 'display']
  };
}

function scoringLegacyTemplatePackRequireHeader_(headers, candidates, sheetName, logicalName) {
  var key = scoringLegacyFindKey_(headers || [], candidates);
  if (!key) {
    throw new Error('テンプレパックをエクスポートできません: ' + sheetName + ' に ' + logicalName + ' 列が見つかりません');
  }
  return key;
}

function scoringLegacyTemplatePackOptionalHeader_(headers, candidates) {
  return scoringLegacyFindKey_(headers || [], candidates);
}

function scoringLegacyTemplatePackBoolToSheetValue_(v) {
  return v ? 1 : 0;
}

function scoringLegacyTemplatePackVisibleToBool_(v, columnExists) {
  if (!columnExists) return true;
  if (v === true) return true;
  if (v === 1) return true;
  if (v === '1') return true;
  var s = String(v || '').trim().toLowerCase();
  return s === 'true';
}

function scoringLegacyBuildTemplatePackFromSheetObjects_(templateRes, ruleRes, exportedAt, name) {
  templateRes = templateRes || { headers: [], rows: [] };
  ruleRes = ruleRes || { headers: [], rows: [] };
  var templateCandidates = scoringLegacyGetTemplatePackTemplateHeaderCandidates_();
  var ruleCandidates = scoringLegacyGetTemplatePackRuleHeaderCandidates_();

  var templateIdKey = scoringLegacyTemplatePackRequireHeader_(templateRes.headers, templateCandidates.templateId, '_templates', 'templateId');
  var templateEnabledKey = scoringLegacyTemplatePackRequireHeader_(templateRes.headers, templateCandidates.enabled, '_templates', 'enabled');
  var templateNameKey = scoringLegacyTemplatePackOptionalHeader_(templateRes.headers, templateCandidates.name);

  var ruleTemplateIdKey = scoringLegacyTemplatePackRequireHeader_(ruleRes.headers, ruleCandidates.templateId, '_rules', 'templateId');
  var rulePriorityKey = scoringLegacyTemplatePackRequireHeader_(ruleRes.headers, ruleCandidates.priority, '_rules', 'priority');
  var ruleWhenKey = scoringLegacyTemplatePackRequireHeader_(ruleRes.headers, ruleCandidates.whenExpr, '_rules', 'whenExpr');
  var ruleMessageKey = scoringLegacyTemplatePackRequireHeader_(ruleRes.headers, ruleCandidates.message, '_rules', 'message');
  var ruleEnabledKey = scoringLegacyTemplatePackRequireHeader_(ruleRes.headers, ruleCandidates.enabled, '_rules', 'enabled');
  var ruleTargetKey = scoringLegacyTemplatePackOptionalHeader_(ruleRes.headers, ruleCandidates.target);
  var ruleVisibleKey = scoringLegacyTemplatePackOptionalHeader_(ruleRes.headers, ruleCandidates.visible);

  var templates = [];
  var templateRows = Array.isArray(templateRes.rows) ? templateRes.rows : [];
  for (var t = 0; t < templateRows.length; t++) {
    var tr = templateRows[t];
    if (!tr || typeof tr !== 'object') continue;
    var tid = String(tr[templateIdKey] == null ? '' : tr[templateIdKey]).trim();
    if (!tid) continue;
    templates.push({
      templateId: tid,
      name: templateNameKey ? String(tr[templateNameKey] == null ? '' : tr[templateNameKey]).trim() : '',
      enabled: scoringLegacyIsEnabled_(tr[templateEnabledKey])
    });
  }

  var rules = [];
  var ruleRows = Array.isArray(ruleRes.rows) ? ruleRes.rows : [];
  for (var r = 0; r < ruleRows.length; r++) {
    var rr = ruleRows[r];
    if (!rr || typeof rr !== 'object') continue;
    var rtid = String(rr[ruleTemplateIdKey] == null ? '' : rr[ruleTemplateIdKey]).trim();
    if (!rtid) continue;
    var priority = scoringLegacyToFiniteNumberOrNaN_(rr[rulePriorityKey]);
    rules.push({
      templateId: rtid,
      target: ruleTargetKey ? String(rr[ruleTargetKey] == null ? '' : rr[ruleTargetKey]).trim() : '',
      priority: isNaN(priority) ? null : priority,
      whenExpr: String(rr[ruleWhenKey] == null ? '' : rr[ruleWhenKey]).trim(),
      message: String(rr[ruleMessageKey] == null ? '' : rr[ruleMessageKey]),
      enabled: scoringLegacyIsEnabled_(rr[ruleEnabledKey]),
      visible: scoringLegacyTemplatePackVisibleToBool_(ruleVisibleKey ? rr[ruleVisibleKey] : null, !!ruleVisibleKey)
    });
  }

  return {
    meta: {
      type: SCORING_TEMPLATE_PACK_META_TYPE,
      version: SCORING_TEMPLATE_PACK_VERSION,
      name: String(name || 'テンプレパック'),
      exportedAt: exportedAt && typeof exportedAt.toISOString === 'function' ? exportedAt.toISOString() : new Date().toISOString()
    },
    templates: templates,
    rules: rules
  };
}

function scoringLegacyEnsureTemplatePackString_(v, key) {
  if (typeof v !== 'string') {
    throw new Error('テンプレパックが不正です: ' + key + ' は文字列で指定してください');
  }
}

function scoringLegacyEnsureTemplatePackBool_(v, key) {
  if (typeof v !== 'boolean') {
    throw new Error('テンプレパックが不正です: ' + key + ' は true/false で指定してください');
  }
}

function scoringLegacyIsPlainObject_(v) {
  return v && typeof v === 'object' && !Array.isArray(v);
}

function scoringLegacyNormalizeTemplatePackPayloadStrict_(payload) {
  if (!scoringLegacyIsPlainObject_(payload)) {
    throw new Error('テンプレパックが不正です: JSONの最上位はオブジェクトで指定してください');
  }
  for (var topKey in payload) {
    if (!payload.hasOwnProperty(topKey)) continue;
    if (topKey !== 'meta' && topKey !== 'templates' && topKey !== 'rules') {
      throw new Error('テンプレパックが不正です: top-level の ' + topKey + ' は許可されていません');
    }
  }
  if (!scoringLegacyIsPlainObject_(payload.meta)) {
    throw new Error('テンプレパックが不正です: meta はオブジェクトで指定してください');
  }
  if (payload.meta.type !== SCORING_TEMPLATE_PACK_META_TYPE) {
    throw new Error('テンプレパックが不正です: meta.type は scoring-tool-template-pack を指定してください');
  }
  scoringLegacyEnsureTemplatePackString_(payload.meta.version, 'meta.version');
  scoringLegacyEnsureTemplatePackString_(payload.meta.name, 'meta.name');
  scoringLegacyEnsureTemplatePackString_(payload.meta.exportedAt, 'meta.exportedAt');
  if (!Array.isArray(payload.templates)) {
    throw new Error('テンプレパックが不正です: templates は配列で指定してください');
  }
  if (!Array.isArray(payload.rules)) {
    throw new Error('テンプレパックが不正です: rules は配列で指定してください');
  }

  var templateIds = {};
  var templates = [];
  for (var i = 0; i < payload.templates.length; i++) {
    var tmpl = payload.templates[i];
    if (!scoringLegacyIsPlainObject_(tmpl)) {
      throw new Error('テンプレパックが不正です: templates[' + i + '] はオブジェクトで指定してください');
    }
    var allowedTemplateKeys = { templateId: true, name: true, enabled: true };
    for (var tk in tmpl) {
      if (tmpl.hasOwnProperty(tk) && !allowedTemplateKeys[tk]) {
        throw new Error('テンプレパックが不正です: templates[' + i + '].' + tk + ' は許可されていません');
      }
    }
    scoringLegacyEnsureTemplatePackString_(tmpl.templateId, 'templates[' + i + '].templateId');
    scoringLegacyEnsureTemplatePackString_(tmpl.name, 'templates[' + i + '].name');
    scoringLegacyEnsureTemplatePackBool_(tmpl.enabled, 'templates[' + i + '].enabled');
    var tid = String(tmpl.templateId || '').trim();
    if (!tid) {
      throw new Error('テンプレパックが不正です: templates[' + i + '].templateId は空にできません');
    }
    var normTid = scoringLegacyNormalizeTemplateId_(tid);
    if (templateIds[normTid]) {
      throw new Error('テンプレパックが不正です: templateId が重複しています: ' + tid);
    }
    templateIds[normTid] = tid;
    templates.push({ templateId: tid, name: String(tmpl.name), enabled: tmpl.enabled });
  }
  if (templates.length === 0) {
    throw new Error('テンプレパックが不正です: templates は1件以上必要です');
  }

  var rules = [];
  for (var r = 0; r < payload.rules.length; r++) {
    var rule = payload.rules[r];
    if (!scoringLegacyIsPlainObject_(rule)) {
      throw new Error('テンプレパックが不正です: rules[' + r + '] はオブジェクトで指定してください');
    }
    var allowedRuleKeys = { templateId: true, target: true, priority: true, whenExpr: true, message: true, enabled: true, visible: true };
    for (var rk in rule) {
      if (rule.hasOwnProperty(rk) && !allowedRuleKeys[rk]) {
        throw new Error('テンプレパックが不正です: rules[' + r + '].' + rk + ' は許可されていません');
      }
    }
    scoringLegacyEnsureTemplatePackString_(rule.templateId, 'rules[' + r + '].templateId');
    scoringLegacyEnsureTemplatePackString_(rule.target, 'rules[' + r + '].target');
    scoringLegacyEnsureTemplatePackString_(rule.whenExpr, 'rules[' + r + '].whenExpr');
    scoringLegacyEnsureTemplatePackString_(rule.message, 'rules[' + r + '].message');
    scoringLegacyEnsureTemplatePackBool_(rule.enabled, 'rules[' + r + '].enabled');
    scoringLegacyEnsureTemplatePackBool_(rule.visible, 'rules[' + r + '].visible');
    var ruleTid = String(rule.templateId || '').trim();
    if (!ruleTid) {
      throw new Error('テンプレパックが不正です: rules[' + r + '].templateId は空にできません');
    }
    if (!templateIds[scoringLegacyNormalizeTemplateId_(ruleTid)]) {
      throw new Error('テンプレパックが不正です: rules[' + r + '].templateId は templates に存在しません: ' + ruleTid);
    }
    var priority = Number(rule.priority);
    if (typeof rule.priority !== 'number' || !isFinite(priority)) {
      throw new Error('テンプレパックが不正です: rules[' + r + '].priority は有限数で指定してください');
    }
    rules.push({
      templateId: ruleTid,
      target: String(rule.target),
      priority: priority,
      whenExpr: String(rule.whenExpr),
      message: String(rule.message),
      enabled: rule.enabled,
      visible: rule.visible
    });
  }

  return {
    meta: {
      type: SCORING_TEMPLATE_PACK_META_TYPE,
      version: payload.meta.version,
      name: payload.meta.name,
      exportedAt: payload.meta.exportedAt
    },
    templates: templates,
    rules: rules
  };
}

function scoringLegacyAssertTemplatePackNoTemplateIdConflicts_(pack, existingTemplateIds) {
  var existing = {};
  var list = Array.isArray(existingTemplateIds) ? existingTemplateIds : [];
  for (var i = 0; i < list.length; i++) {
    var norm = scoringLegacyNormalizeTemplateId_(list[i]);
    if (norm) existing[norm] = true;
  }
  var conflicts = [];
  var templates = pack && Array.isArray(pack.templates) ? pack.templates : [];
  for (var t = 0; t < templates.length; t++) {
    var tid = templates[t].templateId;
    if (existing[scoringLegacyNormalizeTemplateId_(tid)]) conflicts.push(tid);
  }
  if (conflicts.length) {
    throw new Error('既存のtemplateIdと衝突しています: ' + conflicts.join(', ') + '。既存テンプレートを削除するか、templateIdを変更してからインポートしてください');
  }
}

function scoringLegacyGetConfigSpreadsheetForRules_() { return getAppSpreadsheet_(); }

function scoringLegacyResolveTemplatePackImportHeaders_(currentHeaders, candidateMap, canonicalKeys) {
  var headers = Array.isArray(currentHeaders) ? currentHeaders.slice() : [];
  var keyMap = {};
  var keys = canonicalKeys || [];
  for (var i = 0; i < keys.length; i++) {
    var canonical = keys[i];
    var existing = scoringLegacyFindKey_(headers, candidateMap[canonical] || [scoringLegacyNormalizeKey_(canonical)]);
    if (existing) {
      keyMap[canonical] = existing;
      continue;
    }
    headers.push(canonical);
    keyMap[canonical] = canonical;
  }
  return { headers: headers, keyMap: keyMap };
}

function scoringLegacyEnsureSheetWithTemplatePackHeaders_(ss, sheetName, candidateMap, canonicalKeys) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  var required = canonicalKeys || [];
  var lastCol = Math.max(sheet.getLastColumn(), required.length);
  var current = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h || '').trim(); }) : [];
  while (current.length && !current[current.length - 1]) {
    current.pop();
  }
  var hasAnyHeader = false;
  for (var i = 0; i < current.length; i++) {
    if (current[i]) {
      hasAnyHeader = true;
      break;
    }
  }
  if (!hasAnyHeader) {
    sheet.getRange(1, 1, 1, required.length).setValues([required]);
    var initialKeyMap = {};
    for (var k0 = 0; k0 < required.length; k0++) initialKeyMap[required[k0]] = required[k0];
    return { sheet: sheet, headers: required.slice(), keyMap: initialKeyMap };
  }
  var resolved = scoringLegacyResolveTemplatePackImportHeaders_(current, candidateMap || {}, required);
  sheet.getRange(1, 1, 1, resolved.headers.length).setValues([resolved.headers]);
  return { sheet: sheet, headers: resolved.headers, keyMap: resolved.keyMap };
}

function scoringLegacyAppendObjectsByHeaders_(sheet, headers, objects) {
  if (!objects || !objects.length) return;
  var rows = [];
  for (var i = 0; i < objects.length; i++) {
    var obj = objects[i];
    var row = [];
    for (var h = 0; h < headers.length; h++) {
      var key = headers[h];
      row.push(scoringSheetLiteral_(obj.hasOwnProperty(key) ? obj[key] : ''));
    }
    rows.push(row);
  }
  var startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, headers.length).setValues(rows);
}

function scoringLegacyReadExistingTemplateIds_(sheet) {
  if (!sheet) return [];
  var res = scoringLegacyReadSheetObjects_(sheet);
  var idKey = scoringLegacyFindKey_(res.headers, scoringLegacyGetTemplatePackTemplateHeaderCandidates_().templateId);
  if (!idKey) return [];
  var out = [];
  for (var i = 0; i < res.rows.length; i++) {
    var tid = String(res.rows[i][idKey] == null ? '' : res.rows[i][idKey]).trim();
    if (tid) out.push(tid);
  }
  return out;
}

function scoringLegacyApiExportTemplatePack_() {
  var templateSheet = scoringLegacyGetRuleSheetOrNull_('_templates');
  if (!templateSheet) {
    throw new Error('テンプレパックをエクスポートできません: _templates シートが見つかりません');
  }
  var ruleSheet = scoringLegacyGetRuleSheetOrNull_('_rules');
  if (!ruleSheet) {
    throw new Error('テンプレパックをエクスポートできません: _rules シートが見つかりません');
  }
  return scoringLegacyBuildTemplatePackFromSheetObjects_(
    scoringLegacyReadSheetObjects_(templateSheet),
    scoringLegacyReadSheetObjects_(ruleSheet),
    new Date(),
    'テンプレパック'
  );
}

/**
 * ルール用の _templates / _rules シートを作成し、必要ヘッダを補完する。
 * 既存ヘッダに許容別名がある場合は、その列名を維持する。
 * @return {{ok:boolean, spreadsheetId:string, templates:Object, rules:Object}}
 */
function scoringLegacyApiEnsureRuleSheets_() {
  {
    var ss = scoringLegacyGetConfigSpreadsheetForRules_();
    var templateKeys = ['templateId', 'name', 'enabled'];
    var ruleKeys = ['templateId', 'target', 'priority', 'whenExpr', 'message', 'enabled', 'visible'];
    var templateInfo = scoringLegacyEnsureSheetWithTemplatePackHeaders_(ss, '_templates', scoringLegacyGetTemplatePackTemplateHeaderCandidates_(), templateKeys);
    var ruleInfo = scoringLegacyEnsureSheetWithTemplatePackHeaders_(ss, '_rules', scoringLegacyGetTemplatePackRuleHeaderCandidates_(), ruleKeys);

    return {
      ok: true,
      spreadsheetId: ss && typeof ss.getId === 'function' ? ss.getId() : '',
      templates: {
        sheetName: '_templates',
        headers: templateInfo.headers,
        keyMap: templateInfo.keyMap
      },
      rules: {
        sheetName: '_rules',
        headers: ruleInfo.headers,
        keyMap: ruleInfo.keyMap
      }
    };
  }
}

function scoringLegacyApiImportTemplatePack_(payload) {
  var pack = scoringLegacyNormalizeTemplatePackPayloadStrict_(payload);
  {
    var ss = scoringLegacyGetConfigSpreadsheetForRules_();
    var existingTemplateSheet = ss.getSheetByName('_templates');
    scoringLegacyAssertTemplatePackNoTemplateIdConflicts_(pack, scoringLegacyReadExistingTemplateIds_(existingTemplateSheet));

    var templateKeys = ['templateId', 'name', 'enabled'];
    var ruleKeys = ['templateId', 'target', 'priority', 'whenExpr', 'message', 'enabled', 'visible'];
    var templateInfo = scoringLegacyEnsureSheetWithTemplatePackHeaders_(ss, '_templates', scoringLegacyGetTemplatePackTemplateHeaderCandidates_(), templateKeys);
    var ruleInfo = scoringLegacyEnsureSheetWithTemplatePackHeaders_(ss, '_rules', scoringLegacyGetTemplatePackRuleHeaderCandidates_(), ruleKeys);

    var templateRows = pack.templates.map(function (t) {
      var row = {};
      row[templateInfo.keyMap.templateId] = t.templateId;
      row[templateInfo.keyMap.name] = t.name;
      row[templateInfo.keyMap.enabled] = scoringLegacyTemplatePackBoolToSheetValue_(t.enabled);
      return row;
    });
    var ruleRows = pack.rules.map(function (r) {
      var row = {};
      row[ruleInfo.keyMap.templateId] = r.templateId;
      row[ruleInfo.keyMap.target] = r.target;
      row[ruleInfo.keyMap.priority] = r.priority;
      row[ruleInfo.keyMap.whenExpr] = r.whenExpr;
      row[ruleInfo.keyMap.message] = r.message;
      row[ruleInfo.keyMap.enabled] = scoringLegacyTemplatePackBoolToSheetValue_(r.enabled);
      row[ruleInfo.keyMap.visible] = scoringLegacyTemplatePackBoolToSheetValue_(r.visible);
      return row;
    });

    scoringLegacyAppendObjectsByHeaders_(templateInfo.sheet, templateInfo.headers, templateRows);
    scoringLegacyAppendObjectsByHeaders_(ruleInfo.sheet, ruleInfo.headers, ruleRows);
    return {
      ok: true,
      imported: {
        templates: templateRows.length,
        rules: ruleRows.length
      },
      templateIds: pack.templates.map(function (t) { return t.templateId; })
    };
  }
}

function scoringLegacyIsValidColA1OrEmpty_(v) {
  var s = String(v == null ? '' : v).trim();
  return s === '' || /^[A-Za-z]+$/.test(s);
}

function scoringLegacyEnsureString_(v, key) {
  if (typeof v !== 'string') {
    throw new Error('インポートデータが不正です: ' + key + ' は文字列で指定してください');
  }
}

function scoringLegacyEnsureBool_(v, key) {
  if (typeof v !== 'boolean') {
    throw new Error('インポートデータが不正です: ' + key + ' は true/false で指定してください');
  }
}

function scoringLegacyEnsureNumberMin_(v, key, minValue) {
  if (typeof v !== 'number' || !isFinite(v) || v < minValue) {
    throw new Error('インポートデータが不正です: ' + key + ' は ' + minValue + ' 以上の数値で指定してください');
  }
}

function scoringLegacyEnsureStringArray_(arr, key, len) {
  if (!Array.isArray(arr)) {
    throw new Error('インポートデータが不正です: ' + key + ' は配列で指定してください');
  }
  if (typeof len === 'number' && arr.length !== len) {
    throw new Error('インポートデータが不正です: ' + key + ' は長さ ' + len + ' で指定してください');
  }
  for (var i = 0; i < arr.length; i++) {
    if (typeof arr[i] !== 'string') {
      throw new Error('インポートデータが不正です: ' + key + '[' + i + '] は文字列で指定してください');
    }
  }
}

function scoringLegacyEnsureColArray_(arr, key, len) {
  scoringLegacyEnsureStringArray_(arr, key, len);
  for (var i = 0; i < arr.length; i++) {
    if (!scoringLegacyIsValidColA1OrEmpty_(arr[i])) {
      throw new Error('インポートデータが不正です: ' + key + '[' + i + '] は列記号（A, AA など）で指定してください');
    }
  }
}

function scoringLegacyValidateMergeRulesStrict_(mergeRules) {
  if (!scoringLegacyIsPlainObject_(mergeRules)) {
    throw new Error('インポートデータが不正です: mergeRules はオブジェクトで指定してください');
  }
  for (var col in mergeRules) {
    if (!mergeRules.hasOwnProperty(col)) continue;
    if (!scoringLegacyIsValidColA1OrEmpty_(col) || !String(col).trim()) {
      throw new Error('インポートデータが不正です: mergeRules のキーは列記号で指定してください');
    }
    if (typeof mergeRules[col] !== 'string' || !String(mergeRules[col]).trim()) {
      throw new Error('インポートデータが不正です: mergeRules["' + col + '"] は空でない文字列で指定してください');
    }
  }
}

function scoringLegacyValidateRuleOutputSlotsStrict_(slots) {
  if (!Array.isArray(slots)) {
    throw new Error('インポートデータが不正です: ruleOutputSlots は配列で指定してください');
  }
  if (slots.length > 3) {
    throw new Error('インポートデータが不正です: ruleOutputSlots は最大3件です');
  }
  for (var i = 0; i < slots.length; i++) {
    var s = slots[i];
    if (!scoringLegacyIsPlainObject_(s)) {
      throw new Error('インポートデータが不正です: ruleOutputSlots[' + i + '] はオブジェクトで指定してください');
    }
    var allowed = { target: true, col: true, enabled: true, header: true };
    for (var key in s) {
      if (s.hasOwnProperty(key) && !allowed[key]) {
        throw new Error('インポートデータが不正です: ruleOutputSlots[' + i + '].' + key + ' は許可されていません');
      }
    }
    scoringLegacyEnsureString_(s.target, 'ruleOutputSlots[' + i + '].target');
    scoringLegacyEnsureString_(s.col, 'ruleOutputSlots[' + i + '].col');
    scoringLegacyEnsureString_(s.header, 'ruleOutputSlots[' + i + '].header');
    scoringLegacyEnsureBool_(s.enabled, 'ruleOutputSlots[' + i + '].enabled');
    if (!scoringLegacyIsValidColA1OrEmpty_(s.col)) {
      throw new Error('インポートデータが不正です: ruleOutputSlots[' + i + '].col は列記号（A, AA など）で指定してください');
    }
  }
}

function scoringLegacyValidateCommentTemplateStrict_(commentTemplate) {
  if (!scoringLegacyIsPlainObject_(commentTemplate)) {
    throw new Error('インポートデータが不正です: commentTemplate はオブジェクトで指定してください');
  }
  var allowed = { type: true, highThreshold: true, midThreshold: true };
  for (var key in commentTemplate) {
    if (commentTemplate.hasOwnProperty(key) && !allowed[key]) {
      throw new Error('インポートデータが不正です: commentTemplate.' + key + ' は許可されていません');
    }
  }
  scoringLegacyEnsureString_(commentTemplate.type, 'commentTemplate.type');
  scoringLegacyEnsureNumberMin_(commentTemplate.highThreshold, 'commentTemplate.highThreshold', 0);
  scoringLegacyEnsureNumberMin_(commentTemplate.midThreshold, 'commentTemplate.midThreshold', 0);
}

function scoringLegacyValidateImportedConfigStrict_(cfg) {
  if (!scoringLegacyIsPlainObject_(cfg)) {
    throw new Error('インポートデータが不正です: config はオブジェクトで指定してください');
  }
  scoringLegacyAssertNoDeprecatedKeys_(cfg);

  var defaults = scoringLegacyGetDefaultConfig_();
  var allowedKeys = {};
  for (var k in defaults) {
    if (defaults.hasOwnProperty(k)) allowedKeys[k] = true;
  }

  for (var key in cfg) {
    if (cfg.hasOwnProperty(key) && !allowedKeys[key]) {
      throw new Error('インポートデータが不正です: config.' + key + ' は許可されていません');
    }
  }
  for (var req in allowedKeys) {
    if (req !== 'lessonDateHeader' && allowedKeys.hasOwnProperty(req) && !cfg.hasOwnProperty(req)) {
      throw new Error('インポートデータが不正です: config.' + req + ' がありません');
    }
  }

  scoringLegacyEnsureString_(cfg.ruleTemplateId, 'config.ruleTemplateId');
  scoringLegacyEnsureString_(cfg.sheetName, 'config.sheetName');
  scoringLegacyEnsureNumberMin_(cfg.startRow, 'config.startRow', 1);
  scoringLegacyEnsureNumberMin_(cfg.headerRow, 'config.headerRow', 1);
  scoringLegacyEnsureString_(cfg.nameCol, 'config.nameCol');
  scoringLegacyEnsureString_(cfg.nameHeader, 'config.nameHeader');
  scoringLegacyEnsureColArray_(cfg.displayCols, 'config.displayCols', 5);
  scoringLegacyEnsureStringArray_(cfg.displayHeaders, 'config.displayHeaders', 5);
  scoringLegacyEnsureColArray_(cfg.charCountCols, 'config.charCountCols');
  scoringLegacyEnsureColArray_(cfg.scoreCols, 'config.scoreCols', 5);
  scoringLegacyAssertUniqueScoreCols_(cfg.scoreCols);
  scoringLegacyEnsureStringArray_(cfg.scoreHeaders, 'config.scoreHeaders', 5);
  scoringLegacyValidateMergeRulesStrict_(cfg.mergeRules);
  if (!Array.isArray(cfg.colChecks) || cfg.colChecks.length !== 5) {
    throw new Error('インポートデータが不正です: config.colChecks は長さ5の配列で指定してください');
  }
  for (var i = 0; i < cfg.colChecks.length; i++) {
    scoringLegacyEnsureBool_(cfg.colChecks[i], 'config.colChecks[' + i + ']');
  }
  scoringLegacyEnsureBool_(cfg.startFromLastRow, 'config.startFromLastRow');
  scoringLegacyEnsureStringArray_(cfg.slotLabels, 'config.slotLabels', 5);
  scoringLegacyEnsureString_(cfg.commentCol, 'config.commentCol');
  scoringLegacyEnsureString_(cfg.commentHeader, 'config.commentHeader');
  if (!scoringLegacyIsValidColA1OrEmpty_(cfg.nameCol)) {
    throw new Error('インポートデータが不正です: config.nameCol は列記号（A, AA など）で指定してください');
  }
  if (!scoringLegacyIsValidColA1OrEmpty_(cfg.commentCol)) {
    throw new Error('インポートデータが不正です: config.commentCol は列記号（A, AA など）で指定してください');
  }
  scoringLegacyValidateRuleOutputSlotsStrict_(cfg.ruleOutputSlots);
  scoringLegacyEnsureString_(cfg.flushShortcut, 'config.flushShortcut');
  scoringLegacyValidateCommentTemplateStrict_(cfg.commentTemplate);
}

function scoringLegacyBackupCurrentConfig_() {
  var props = scoringLegacyGetScriptProps_();
  var current = props.getProperty('TURRET_SCORING_CONFIG');
  if (current && String(current).trim()) {
    props.setProperty(SCORING_CONFIG_BACKUP_KEY, current);
  }
}

/**
 * 画面上設定をインポート（CONFIGを上書き）
 * @param {{meta?:Object, config:Object}} payload
 * @return {Object} 保存後の設定
 */
function scoringLegacyApiImportConfig_(payload) {
  var imported = scoringParseImportedConfig_(payload);
  scoringLegacyBackupCurrentConfig_();
  var props = scoringLegacyGetScriptProps_();
  props.setProperty('TURRET_SCORING_CONFIG', JSON.stringify(imported));
  return scoringLegacyApiGetConfig_();
}

/** 旧採点JSONの検証を、保存せずに実行する。 */
function scoringParseImportedConfig_(payload) {
  if (!payload || typeof payload !== 'object') {
    throw new Error('インポートデータが不正です');
  }

  var imported = null;
  if (payload.config && typeof payload.config === 'object') {
    if (payload.meta !== undefined && !scoringLegacyIsPlainObject_(payload.meta)) {
      throw new Error('インポートデータが不正です: meta はオブジェクトで指定してください');
    }
    for (var topKey in payload) {
      if (!payload.hasOwnProperty(topKey)) continue;
      if (topKey !== 'meta' && topKey !== 'config') {
        throw new Error('インポートデータが不正です: top-level の ' + topKey + ' は許可されていません');
      }
    }
    imported = payload.config;
  } else {
    // 互換: 旧形式やconfigラッパーなしのオブジェクト
    imported = payload;
  }
  if (!imported || typeof imported !== 'object') {
    throw new Error('config が見つかりません');
  }

  scoringLegacyValidateImportedConfigStrict_(imported);
  if (utf8ByteLength_(JSON.stringify(imported)) > 8000) throw new Error('採点設定はUTF-8で8,000バイト以内にしてください。');
  return imported;
}

/**
 * 直近バックアップから画面上設定を復元
 * @return {Object} 復元後の設定
 */
function scoringLegacyApiRestoreLatestConfigBackup_() {
  var props = scoringLegacyGetScriptProps_();
  var backup = props.getProperty(SCORING_CONFIG_BACKUP_KEY);
  if (!backup || !String(backup).trim()) {
    throw new Error('復元可能なバックアップがありません');
  }
  var parsed = null;
  try {
    parsed = JSON.parse(backup);
  } catch (e) {
    throw new Error('バックアップが破損しています');
  }
  scoringLegacyValidateImportedConfigStrict_(parsed);
  props.setProperty('TURRET_SCORING_CONFIG', JSON.stringify(parsed));
  return scoringLegacyApiGetConfig_();
}

// -----------------------------
// ルールAPI（読み取り専用）
// -----------------------------

/** ルール用シートを取得（設定用スプシから、無ければ null） */
function scoringLegacyGetRuleSheetOrNull_(name) { return getAppSpreadsheet_().getSheetByName(name); }

/** ヘッダ行（1行目）つきテーブルを Object 配列として読み取る */
function scoringLegacyReadSheetObjects_(sheet) {
  var range = sheet.getDataRange();
  var values = range.getValues();
  if (!values || values.length === 0) {
    return { headers: [], rows: [] };
  }

  var headers = values[0].map(function (h) { return String(h || '').trim(); });
  if (values.length < 2) {
    return { headers: headers, rows: [] };
  }

  var rows = [];
  for (var r = 1; r < values.length; r++) {
    var rowVals = values[r];
    var isEmpty = true;
    for (var c0 = 0; c0 < rowVals.length; c0++) {
      if (rowVals[c0] !== '' && rowVals[c0] !== null) {
        isEmpty = false;
        break;
      }
    }
    if (isEmpty) continue;

    var obj = { _rowNumber: r + 1 };
    for (var c = 0; c < headers.length; c++) {
      var key = headers[c];
      if (!key) continue;
      obj[key] = rowVals[c];
    }
    rows.push(obj);
  }
  return { headers: headers, rows: rows };
}

function scoringLegacyNormalizeKey_(key) {
  return String(key || '')
    .toLowerCase()
    .replace(/\s+/g, '')
    .replace(/[_-]/g, '');
}

function scoringLegacyFindKey_(headers, normalizedCandidates) {
  for (var i = 0; i < headers.length; i++) {
    var h = headers[i];
    if (!h) continue;
    var nh = scoringLegacyNormalizeKey_(h);
    for (var j = 0; j < normalizedCandidates.length; j++) {
      if (nh === normalizedCandidates[j]) return h;
    }
  }
  return '';
}

function scoringLegacyFindRuleMessageKey_(headers) {
  var candidates = ['message', 'text', 'output', 'result', 'comment'];
  var list = Array.isArray(headers) ? headers : [];
  for (var c = 0; c < candidates.length; c++) {
    for (var i = 0; i < list.length; i++) {
      var h = list[i];
      if (!h) continue;
      if (scoringLegacyNormalizeKey_(h) === candidates[c]) return h;
    }
  }
  return '';
}

/**
 * enabled の値を true/false に解釈（型ブレ対策）
 * true扱い: 1, "1", true, "true", "TRUE"（大小文字無視）
 * それ以外は false 扱い
 */
function scoringLegacyIsEnabled_(v) {
  if (v === 1) return true;
  if (v === '1') return true;
  if (v === true) return true;
  var s = String(v || '').trim().toLowerCase();
  if (s === 'true') return true;
  return false;
}

/**
 * priority の値を有限数に変換（空欄/非数は NaN）
 * 注意: Number('') === 0 なので、空欄を 0 扱いしないようガードが必要
 */
function scoringLegacyToFiniteNumberOrNaN_(v) {
  if (v === null || v === undefined) return NaN;
  var s = String(v).trim();
  if (s === '') return NaN; // 空欄は invalid
  var n = (typeof v === 'number') ? v : Number(s);
  return (isFinite(n) ? n : NaN);
}

/**
 * templateId を正規化（trim + lowercase）
 * 大小文字を区別しない比較のため
 */
function scoringLegacyNormalizeTemplateId_(v) {
  return String(v || '').trim().toLowerCase();
}

/**
 * ルールテンプレ一覧取得（_templates）
 * @return {{ok:boolean, templates:object[], headers?:string[], error?:string}}
 */
function scoringLegacyApiGetRuleTemplates_() {
  try {
    var sheet = scoringLegacyGetRuleSheetOrNull_('_templates');
    if (!sheet) {
      var result = { ok: false, templates: [], error: 'sheet not found: _templates' };
      if (SCORING_DEBUG_MODE) result._debug = { step: 'scoringLegacyGetRuleSheetOrNull_', sheetFound: false };
      return result;
    }
    var res = scoringLegacyReadSheetObjects_(sheet);
    var templateIdKey = scoringLegacyFindKey_(res.headers, ['templateid', 'template']);
    var enabledKey = scoringLegacyFindKey_(res.headers, ['enabled', 'enable', 'active', 'isenabled', 'isactive']);
    if (!templateIdKey) {
      var result = { ok: false, templates: [], error: 'templateId column not found in _templates' };
      if (SCORING_DEBUG_MODE) result._debug = { step: 'findKey_templateId', headers: res.headers };
      return result;
    }
    if (!enabledKey) {
      var result = { ok: false, templates: [], error: 'enabled column not found in _templates' };
      if (SCORING_DEBUG_MODE) result._debug = { step: 'findKey_enabled', headers: res.headers };
      return result;
    }
    var result = { ok: true, templates: res.rows, headers: res.headers };
    if (SCORING_DEBUG_MODE) result._debug = { step: 'success', templatesCount: res.rows.length, headersCount: res.headers.length };
    return result;
  } catch (e) {
    var result = { ok: false, templates: [], error: String(e && e.message ? e.message : e) };
    if (SCORING_DEBUG_MODE) result._debug = { step: 'exception', errorMessage: e ? e.message : null, errorString: String(e) };
    return result;
  }
}

/**
 * ルール一覧取得（_rules）
 * 表示＝評価対象を保証するため、enabled=true かつ priority有限数のみ返し、priority昇順でソート
 * @param {{templateId:string}} params
 * @return {{ok:boolean, templateId?:string, rules:object[], headers?:string[], skipped?:{disabled:number,invalidPriority:number,invalidRow:number}, error?:string}}
 */
function scoringLegacyApiGetRules_(params) {
  try {
    var templateIdRaw = String(params && params.templateId ? params.templateId : '').trim();
    if (!templateIdRaw) {
      return { ok: false, rules: [], error: 'templateId is required' };
    }
    var templateIdNorm = scoringLegacyNormalizeTemplateId_(templateIdRaw);

    // _templates の存在確認（templateId存在+enabled=true を確認）
    var tmplSheet = scoringLegacyGetRuleSheetOrNull_('_templates');
    if (!tmplSheet) {
      return { ok: false, templateId: templateIdRaw, rules: [], error: 'sheet not found: _templates' };
    }
    var tmplRes = scoringLegacyReadSheetObjects_(tmplSheet);
    var tmplIdKey = scoringLegacyFindKey_(tmplRes.headers, ['templateid', 'template']);
    var tmplEnabledKey = scoringLegacyFindKey_(tmplRes.headers, ['enabled', 'enable', 'active', 'isenabled', 'isactive']);
    if (!tmplIdKey) {
      return { ok: false, templateId: templateIdRaw, rules: [], error: 'templateId column not found in _templates' };
    }
    if (!tmplEnabledKey) {
      return { ok: false, templateId: templateIdRaw, rules: [], error: 'enabled column not found in _templates' };
    }
    var templateExists = false;
    for (var t = 0; t < tmplRes.rows.length; t++) {
      var tr = tmplRes.rows[t];
      if (scoringLegacyNormalizeTemplateId_(tr[tmplIdKey]) !== templateIdNorm) continue;
      if (scoringLegacyIsEnabled_(tr[tmplEnabledKey])) {
        templateExists = true;
        break;
      }
    }
    if (!templateExists) {
      return { ok: false, templateId: templateIdRaw, rules: [], error: 'template not found or disabled: ' + templateIdRaw };
    }

    var sheet = scoringLegacyGetRuleSheetOrNull_('_rules');
    if (!sheet) {
      return { ok: false, templateId: templateIdRaw, rules: [], error: 'sheet not found: _rules' };
    }
    var res = scoringLegacyReadSheetObjects_(sheet);
    var templateIdKey = scoringLegacyFindKey_(res.headers, ['templateid', 'template']);
    var enabledKey = scoringLegacyFindKey_(res.headers, ['enabled', 'enable', 'active', 'isenabled', 'isactive']);
    var priorityKey = scoringLegacyFindKey_(res.headers, ['priority', 'prio', 'order', 'rank']);
    if (!templateIdKey) {
      return { ok: false, templateId: templateIdRaw, rules: [], error: 'templateId column not found in _rules' };
    }
    if (!enabledKey) {
      return { ok: false, templateId: templateIdRaw, rules: [], error: 'enabled column not found in _rules' };
    }
    if (!priorityKey) {
      return { ok: false, templateId: templateIdRaw, rules: [], error: 'priority column not found in _rules' };
    }

    var skipped = { disabled: 0, invalidPriority: 0, invalidRow: 0 };
    var candidates = [];
    for (var i = 0; i < res.rows.length; i++) {
      var row = res.rows[i];
      if (!row || typeof row !== 'object') {
        skipped.invalidRow++;
        continue;
      }
      // templateId 正規化比較
      if (scoringLegacyNormalizeTemplateId_(row[templateIdKey]) !== templateIdNorm) continue;
      // enabled=true のみ採用
      if (!scoringLegacyIsEnabled_(row[enabledKey])) {
        skipped.disabled++;
        continue;
      }
      // priority が有限数のみ採用
      var pr = scoringLegacyToFiniteNumberOrNaN_(row[priorityKey]);
      if (isNaN(pr)) {
        skipped.invalidPriority++;
        continue;
      }
      candidates.push({ row: row, priority: pr, idx: i });
    }

    // priority 昇順でソート
    candidates.sort(function (a, b) {
      if (a.priority !== b.priority) return a.priority - b.priority;
      return a.idx - b.idx;
    });

    var rules = candidates.map(function (c) { return c.row; });

    // 返却rulesが 0 件なら ok:false（テンプレ有効だがルール0件）
    if (rules.length === 0) {
      return { ok: false, templateId: templateIdRaw, rules: [], skipped: skipped, error: 'no valid rules found for template: ' + templateIdRaw };
    }

    return { ok: true, templateId: templateIdRaw, rules: rules, headers: res.headers, skipped: skipped };
  } catch (e) {
    return { ok: false, rules: [], error: String(e && e.message ? e.message : e) };
  }
}
