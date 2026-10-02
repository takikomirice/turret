/** 管理画面の入口、進捗、明示操作、設定プロファイル。外部送信は操作allowlist経由のみ。 */
const ADMIN_SHEETS = [
  { key: 'classes', label: 'クラス一覧' }, { key: 'students', label: '生徒一覧' },
  { key: 'mapping', label: '対応表' },
  { key: 'evaluation', label: '評価データ' }, { key: 'send', label: '送信シート' }, { key: 'errors', label: 'エラー' }
];
const PROFILE_CONFIG_KEYS = ['emailHeader', 'studentNameHeader', 'formStatusHeader', 'scoreSourceHeader',
  'replyBodyHeader', 'formSheetNamePrefix', 'fields', 'messageTemplate', 'gradeScale'];

const MANAGED_FORM_SHEET_ = 'システム管理';
const MANAGED_FORM_HEADERS_ = ['種別', 'ID', '内容(JSON)'];
const TEMPLATE_FORM_URL_KEY_ = 'TURRET_TEMPLATE_FORM_URL';

/** 成績出力。元回答・返却判断から独立した出力用の設定を扱う。 */
function gradeDefault_() {
  return {base:{classIds:[],start:'',end:'',dateHeader:'授業の日付を入力してください',valueHeader:'',leading:['roster:number','roster:name'],sort:'name',mapping:{}},calendar:[],rules:{dates:[],answers:[],duplicates:[],fields:[],lessons:[]}};
}
function gradeDate_(value) { return scoringLessonDate_(value, 'Asia/Tokyo'); }
function gradeKey_(parts) { return JSON.stringify(parts); }
/** RPCがオブジェクトのキー順を変えても同じ内容として照合する。 */
function gradeDigest_(value) {
  function canonical(x){if(Object.prototype.toString.call(x)==='[object Date]')return isFinite(x.getTime())?x.toISOString():'Invalid Date';if(Array.isArray(x))return x.map(canonical);if(x && typeof x==='object'){const out=Object.create(null);Object.keys(x).sort().forEach(function(k){out[k]=canonical(x[k]);});return out;}return x;}
  return scoringDigest_(canonical(value));
}
function gradeText_(value) { return value == null ? '' : String(value); }
function gradeInPeriod_(date, base) { return !date || ((!base.start || date>=base.start) && (!base.end || date<=base.end)); }

/** 学校出力の個別予定を読む。解釈できない繰り返しは黙って欠落させない。 */
function gradeParseIcs_(text) {
  if(typeof text!=='string' || text.length>2000000)throw new Error('ICSは2MB以内で指定してください。');
  const lines=text.replace(/^\uFEFF/,'').replace(/\r\n/g,'\n').replace(/\r/g,'\n').replace(/\n[ \t]/g,'').split('\n'),events=[];
  if(!lines.includes('BEGIN:VCALENDAR') || !lines.includes('END:VCALENDAR'))throw new Error('iCalendar形式ではありません。');
  let event=null,nested=0;
  lines.forEach(function(line){
    if(line==='BEGIN:VEVENT'){if(event)throw new Error('ICSの予定が不正です。');event={};return;}
    if(line==='END:VEVENT'){
      if(!event)throw new Error('ICSの予定が不正です。');
      if(event.STATUS!=='CANCELLED'){
        if(!event.SUMMARY || !event.DTSTART)throw new Error('授業名または開始日時のない予定です。');
        const m=event.DTSTART.match(/^(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})(\d{2})(Z)?)?$/);
        if(!m || !gradeDate_(m[1]+'-'+m[2]+'-'+m[3]) || (m[4] && (+m[4]>23 || +m[5]>59 || +m[6]>59)))throw new Error('ICSの日時が不正です。');
        let date=m[1]+'-'+m[2]+'-'+m[3],time=m[4]?m[4]+':'+m[5]:'';
        if(m[7]){const d=new Date(Date.UTC(+m[1],+m[2]-1,+m[3],+m[4]+9,+m[5],+m[6]));date=d.toISOString().slice(0,10);time=d.toISOString().slice(11,16);}
        events.push({id:event.UID || gradeKey_([event.SUMMARY,date,time]),title:event.SUMMARY.replace(/\\[nN]/g,'\n').replace(/\\([,;\\])/g,'$1'),date:date,time:time});
      }
      event=null;return;
    }
    if(!event)return;
    if(line.startsWith('BEGIN:')){nested++;return;}
    if(line.startsWith('END:')&&nested){nested--;return;}
    if(nested)return;
    const at=line.indexOf(':');if(at<0)return;
    const left=line.slice(0,at),key=left.split(';')[0],value=line.slice(at+1);
    if(['RRULE','RDATE','EXDATE','RECURRENCE-ID'].includes(key))throw new Error('繰り返し予定には未対応です。各授業を個別予定として書き出してください。');
    if(key==='DTSTART' && /TZID=/.test(left) && !/TZID="?(Asia\/Tokyo|Japan)"?(;|$)/.test(left))throw new Error('ICSのタイムゾーンに対応していません。日本時間またはUTCで書き出してください。');
    event[key]=value;
  });
  if(event)throw new Error('ICSの予定が閉じられていません。');
  const ids=new Map();events.forEach(function(e){if(ids.has(e.id) && JSON.stringify(ids.get(e.id))!==JSON.stringify(e))throw new Error('ICSに同じIDの異なる予定があります。');ids.set(e.id,e);});
  return Array.from(ids.values()).sort(function(a,b){return (a.date+a.time).localeCompare(b.date+b.time);});
}
function gradeParseIcs(text) { assertWebOperator_(); return gradeParseIcs_(text); }

/** 集計は純粋関数にして、全員出力と未解決データの非欠落を検証する。 */
function gradeBuild_(data, profile) {
  const base=profile.base,rules=profile.rules,issues=(data.issues || []).slice(),output=[];
  const excluded=new Set((data.excludedStudents||[]).map(function(s){return makeStudentLookupKey_(s.courseId,s.email);}));
  const selected=data.classes.filter(function(c){return !base.classIds.length || base.classIds.includes(c.id);});
  base.classIds.forEach(function(id){if(!data.classes.some(function(c){return c.id===id;}))issues.push({type:'class',message:'対象クラスが見つかりません: '+id});});
  selected.forEach(function(course){
    const calendar=course.students.length ? profile.calendar.filter(function(e){return base.mapping[course.id] && e.title===base.mapping[course.id];}) : [];
    if(course.students.length && base.mapping[course.id] && !calendar.length)issues.push({type:'calendar',courseId:course.id,message:'対応するICS授業がありません。対応を選び直すか「ICSを使わない」を選んでください。'});
    const days=new Map(),groups=new Map(),byEmail=new Map(),dateIssues=new Map();
    course.students.forEach(function(s){if(byEmail.has(s.email))issues.push({type:'roster',courseId:course.id,message:'名簿に同じメールの生徒が重複しています: '+s.email});byEmail.set(s.email,s);});
    calendar.forEach(function(e){if(!gradeInPeriod_(e.date,base))return;const day=days.get(e.date)||{date:e.date,extra:false,slots:0};day.slots++;days.set(e.date,day);});
    (rules.lessons||[]).filter(function(x){return x.courseId===course.id && gradeInPeriod_(x.date,base);}).forEach(function(x){const day=days.get(x.date)||{date:x.date,extra:false,slots:1};day.slots=x.count;days.set(x.date,day);});
    const responses=data.responses.filter(function(r){return r.courseId===course.id;});
    responses.forEach(function(original){
      if(excluded.has(makeStudentLookupKey_(course.id,original.email)))return;
      const r=Object.assign({},original),own=rules.answers.find(function(x){return x.id===r.id && x.fingerprint===r.fingerprint;});
      const stale=rules.answers.find(function(x){return x.id===r.id && x.fingerprint!==r.fingerprint;});
      if(stale && r.date && stale.to && !gradeInPeriod_(r.date,base) && !gradeInPeriod_(stale.to,base))return;
      if(stale){issues.push({type:'date',stale:true,key:r.id,courseId:course.id,date:stale.to,message:'以前の個別補正後に回答が変更されました。この回答の補正を再確認してください。',responses:[original]});return;}
      const rule=own || rules.dates.find(function(x){return x.courseId===course.id && x.from===r.date;});
      r.date=rule ? rule.to : r.date;r.extra=!!(rule && rule.mode==='extra');
      if(!gradeInPeriod_(r.date,base))return;
      if(!byEmail.has(r.email)){issues.push({type:'roster',courseId:course.id,message:'生徒一覧にない回答です。名簿を確認してください: '+r.email,response:r});return;}
      if(!rule && (!r.date || (base.mapping[course.id] && !calendar.some(function(e){return e.date===r.date;}) && !(rules.lessons||[]).some(function(x){return x.courseId===course.id && x.date===r.date;})))){
        const key=gradeKey_([course.id,original.date]);let issue=dateIssues.get(key);
        if(!issue){issue={type:'date',key:key,courseId:course.id,date:original.date,message:original.date?'ICSにない授業日です。扱いを選んでください。':'授業日が不明です。日付を指定するか、予定外として残してください。',responses:[]};dateIssues.set(key,issue);}
        issue.responses.push(original);
        // 空欄も確認対象として残す。
        if(!rule)return;
      }
      const date=r.date || '日付不明',bucket=date+(r.extra?'|extra':''),day=days.get(bucket)||{date:date,extra:r.extra,slots:1};
      days.set(bucket,day);r.date=date;r.bucket=bucket;
      const key=gradeKey_([course.id,r.email,date,r.extra]),group=groups.get(key)||[];group.push(r);groups.set(key,group);
    });
    dateIssues.forEach(function(issue){issues.push(issue);});
    const cells=new Map();
    function groupFingerprint(rows){return gradeDigest_([rows.map(function(r){return [r.id,r.fingerprint];}).sort(),calendar.filter(function(e){return e.date===rows[0].date;}).map(function(e){return [e.title,e.date,e.time];}).sort(),(rules.lessons||[]).filter(function(x){return x.courseId===course.id && x.date===rows[0].date;})]);}
    // 全生徒共通の授業回数を先に確定し、処理順による自動割当の違いを防ぐ。
    groups.forEach(function(rows,key){const decision=rules.duplicates.find(function(d){return d.key===key && d.fingerprint===groupFingerprint(rows);});if(decision){const day=days.get(rows[0].bucket);rows.forEach(function(r){day.slots=Math.max(day.slots,decision.assignments[r.id]||0);});}});
    groups.forEach(function(rows,key){
      const day=days.get(rows[0].bucket),fingerprint=groupFingerprint(rows),decision=rules.duplicates.find(function(d){return d.key===key && d.fingerprint===fingerprint;});
      let assignments={};
      if(rows.length>1 || day.slots>1 || rules.duplicates.some(function(d){return d.key===key;})){
        const used=new Set();let valid=!!decision;
        rows.forEach(function(r){const n=decision && decision.assignments[r.id];if(!Number.isInteger(n) || n<0 || n>20 || (n && used.has(n)))valid=false;if(n)used.add(n);});
        if(!used.size)valid=false;
        if(!valid){issues.push({type:'duplicate',key:key,fingerprint:fingerprint,courseId:course.id,date:rows[0].date,email:rows[0].email,message:'授業への割り当て・採否を選んでください。0は今回の出力から除外です。',responses:rows,slots:day.slots});return;}
        assignments=decision.assignments;
      }else assignments[rows[0].id]=1;
      rows.forEach(function(r){const slot=assignments[r.id];if(!slot)return;day.slots=Math.max(day.slots,slot);cells.set(gradeKey_([r.email,r.bucket,slot]),gradeText_(r.fields[base.valueHeader]));});
    });
    const columns=Array.from(days.values()).sort(function(a,b){return Number(a.extra)-Number(b.extra)||a.date.localeCompare(b.date);}).flatMap(function(d){return Array.from({length:d.slots},function(_,i){return {date:d.date,bucket:d.date+(d.extra?'|extra':''),slot:i+1,label:d.date.replace(/-/g,'/')+(d.slots>1?'（'+(i+1)+'回目）':'')+(d.extra?'（予定外）':'')};});});
    let students=course.students.slice();
    if(base.sort==='name')students.sort(function(a,b){return a.name.localeCompare(b.name,'ja',{numeric:true})||a.email.localeCompare(b.email);});
    if(base.sort==='number')students.sort(function(a,b){return Number(!a.number)-Number(!b.number)||a.number.localeCompare(b.number,'ja',{numeric:true})||a.name.localeCompare(b.name,'ja');});
    const labels={'roster:name':'名前','roster:number':'出席番号','roster:email':'メールアドレス','roster:no':'No'};
    const rows=[base.leading.map(function(h){return labels[h] || h.slice(6);}).concat(columns.map(function(c){return c.label;}))];
    students.forEach(function(s){
      const leading=base.leading.map(function(h){if(h.startsWith('roster:'))return gradeText_(s[h.slice(7)]);
        const header=h.slice(6),values=Array.from(new Set(responses.filter(function(r){return r.email===s.email;}).map(function(r){return gradeText_(r.fields[header]);}).filter(Boolean)));
        if(values.length<=1)return values[0] || '';
        const key=gradeKey_([course.id,s.email,header]),fingerprint=gradeDigest_(values.slice().sort()),decision=rules.fields.find(function(x){return x.key===key && x.fingerprint===fingerprint && values.includes(x.value);});
        if(decision)return decision.value;
        issues.push({type:'field',key:key,fingerprint:fingerprint,courseId:course.id,email:s.email,values:values,message:'先頭列「'+header+'」の値が回答ごとに異なります。出力値を選んでください。'});return '';
      });
      rows.push(leading.concat(columns.map(function(col){return cells.get(gradeKey_([s.email,col.bucket,col.slot])) || '';})));
    });
    if(students.length)output.push({id:course.id,name:course.name,rows:rows});
  });
  if(!output.length)issues.push({type:'class',message:'生徒一覧に対象クラスがありません。生徒一覧を取得してください。'});
  return {classes:output,issues:issues};
}
function gradeCsv_(rows) {
  return '\uFEFF'+rows.map(function(row){return row.map(function(value){
    let s=gradeText_(value);
    // 回答由来の文字列を表計算の式として実行させない。数値の負数は保持する。
    if(/^[\s]*[=+@]/.test(s) || (/^[\s]*-/.test(s) && !/^\s*-\d+(\.\d+)?\s*$/.test(s)))s="'"+s;
    return /[",\r\n]/.test(s)?'"'+s.replace(/"/g,'""')+'"':s;
  }).join(',');}).join('\r\n')+'\r\n';
}

function gradeValidate_(input) {
  profileKeys_(input,['base','calendar','rules'],'成績出力の形式');
  const p=JSON.parse(JSON.stringify(input)),b=p.base,r=p.rules;
  profileKeys_(b,['classIds','start','end','dateHeader','valueHeader','leading','sort','mapping'],'成績出力設定');
  if(!Array.isArray(b.classIds)||b.classIds.some(function(x){return typeof x!=='string';}) || !['name','number','roster'].includes(b.sort))throw new Error('対象クラス・並び順が不正です。');
  ['start','end','dateHeader','valueHeader'].forEach(function(k){if(typeof b[k]!=='string')throw new Error('列または日付の形式が不正です。');});
  if((b.start && gradeDate_(b.start)!==b.start)||(b.end && gradeDate_(b.end)!==b.end)||(b.start && b.end && b.start>b.end))throw new Error('期間の日付が不正です。');
  if(!Array.isArray(b.leading)||!b.leading.length || b.leading.length>30 || b.leading.some(function(x){return typeof x!=='string'||(!['roster:name','roster:number','roster:email','roster:no'].includes(x)&&!/^field:.+/.test(x));}))throw new Error('先頭列が不正です。');
  requireProfileObject_(b.mapping,'クラスの対応');
  Object.keys(b.mapping).forEach(function(k){if(['__proto__','constructor','prototype'].includes(k)||typeof b.mapping[k]!=='string')throw new Error('対応関係の形式が不正です。');});
  if(!Array.isArray(p.calendar)||p.calendar.length>20000)throw new Error('授業日一覧が不正です。');
  p.calendar.forEach(function(e){profileKeys_(e,['id','title','date','time'],'授業日');if(typeof e.id!=='string'||!e.id||typeof e.title!=='string'||!e.title||gradeDate_(e.date)!==e.date||typeof e.time!=='string'||(e.time&&!/^(?:[01]\d|2[0-3]):[0-5]\d$/.test(e.time)))throw new Error('授業日の形式が不正です。');});
  profileKeys_(r,['dates','answers','duplicates','fields','lessons'],'補正ルール');
  if(r.lessons===undefined)r.lessons=[];
  if(!Array.isArray(r.lessons))throw new Error('授業回数が不正です。');
  r.lessons.forEach(function(x){profileKeys_(x,['courseId','date','count'],'授業回数');if(typeof x.courseId!=='string'||!x.courseId||gradeDate_(x.date)!==x.date||!Number.isInteger(x.count)||x.count<1||x.count>20)throw new Error('授業回数・日付が不正です。');});
  ['dates','answers','duplicates','fields'].forEach(function(k){if(!Array.isArray(r[k])||r[k].length>20000)throw new Error('補正ルールが不正です。');});
  r.answers.forEach(function(x){if(typeof x.id!=='string'||!x.id||typeof x.fingerprint!=='string'||!x.fingerprint)throw new Error('個別回答の識別情報が不正です。');});
  r.dates.forEach(function(x){if(typeof x.courseId!=='string'||!x.courseId||typeof x.from!=='string'||x.id!==undefined)throw new Error('クラス日付補正の形式が不正です。');});
  r.dates.concat(r.answers).forEach(function(x){
    profileKeys_(x,['courseId','from','to','mode','id','fingerprint'],'日付補正');
    if(!['planned','extra'].includes(x.mode)||typeof x.to!=='string'||(x.to && gradeDate_(x.to)!==x.to)||(!x.to&&x.mode!=='extra'))throw new Error('補正先の日付が不正です。');
    if(x.id!==undefined){if(typeof x.id!=='string'||typeof x.fingerprint!=='string')throw new Error('回答の形式が不正です。');}
    else if(typeof x.courseId!=='string'||typeof x.from!=='string')throw new Error('補正元が不正です。');
  });
  r.duplicates.forEach(function(x){profileKeys_(x,['key','fingerprint','assignments'],'重複判断');if(typeof x.key!=='string'||typeof x.fingerprint!=='string')throw new Error('重複判断が不正です。');requireProfileObject_(x.assignments,'授業割り当て');Object.keys(x.assignments).forEach(function(k){const n=x.assignments[k];if(['__proto__','constructor','prototype'].includes(k)||!Number.isInteger(n)||n<0||n>20)throw new Error('授業割り当てが不正です。');});});
  r.fields.forEach(function(x){profileKeys_(x,['key','fingerprint','value'],'先頭列の判断');if(['key','fingerprint','value'].some(function(k){return typeof x[k]!=='string';}))throw new Error('先頭列の判断が不正です。');});
  ['dates','answers','duplicates','fields','lessons'].forEach(function(k){const ids=new Set();r[k].forEach(function(x){const id=k==='lessons'?gradeKey_([x.courseId,x.date]):k==='dates'?gradeKey_([x.courseId,x.from]):k==='answers'?x.id:x.key;if(ids.has(id))throw new Error('補正ルールの対象が重複しています。');ids.add(id);});});
  if(JSON.stringify(p).length>1500000)throw new Error('成績出力設定が大きすぎます。期間別のプリセットに分けてください。');
  return p;
}

/** 二つの保存領域を交互に使い、読戻し後に参照を切り替える。途中失敗は旧版を保持する。 */
function gradeReadStore_() {
  const records=getManagedRecords_(),head=records.find(function(r){return r.id==='grade:head';});
  if(!head)return {revision:'',value:{draft:gradeDefault_(),presets:[]}};
  assertManagedOwner_(head);
  let text='';for(let i=0;i<head.count;i++){
    const part=records.find(function(r){return r.id==='grade:part:'+head.bank+':'+i;});
    if(!part || part.generation!==head.generation)throw new Error('成績出力の保存内容が不完全です。JSONから復元してください。');text+=part.text;
  }
  if(gradeDigest_(text)!==head.digest)throw new Error('成績出力の保存内容が一致しません。');
  return {revision:head.generation,bank:head.bank,value:JSON.parse(text)};
}
function gradeWriteStore_(value, expected) {
  const previous=gradeReadStore_();if(previous.revision!==expected)throw new Error('成績出力設定が別の画面で変更されました。再読み込みしてください。');
  const text=JSON.stringify(value);if(text.length>4000000)throw new Error('保存量が上限を超えました。不要なプリセットをJSONに退避して削除してください。');
  const bank=previous.bank==='a'?'b':'a',generation=Utilities.getUuid(),parts=[];
  for(let i=0;i<text.length;i+=24000)parts.push(text.slice(i,i+24000));
  const cache={records:getManagedRecords_(),sheet:managedSheet_(true)};
  parts.forEach(function(part,i){
    const record={schema:1,kind:'grade-part',id:'grade:part:'+bank+':'+i,generation:generation,text:part,ownerSpreadsheetId:getAppSpreadsheet_().getId(),ownerScriptId:ScriptApp.getScriptId(),revision:generation,updatedAt:new Date().toISOString()};
    const at=cache.records.findIndex(function(r){return r.id===record.id;});if(at<0)cache.records.push(record);else{assertManagedOwner_(cache.records[at]);cache.records[at]=record;}
  });
  // 既存の確定参照はそのまま残し、分割内容を一括で書き込む。
  cache.sheet.getRange(2,1,cache.records.length,3).setValues(cache.records.map(function(r){return [r.kind,r.id,JSON.stringify(r)];}));SpreadsheetApp.flush();
  const verify=getManagedRecords_();
  if(parts.some(function(part,i){const r=verify.find(function(x){return x.id==='grade:part:'+bank+':'+i;});return !r || r.text!==part || r.generation!==generation;}))throw new Error('成績出力設定の保存を確認できません。旧設定を保持しています。');
  saveManagedRecord_({kind:'grade-head',id:'grade:head',bank:bank,generation:generation,count:parts.length,digest:gradeDigest_(text)},cache);
  return gradeReadStore_();
}
function gradeReadData_(base, strict) {
  const read=scoringReadContext_(),classes=new Map(),issues=[],headers=new Set(),responses=[],excludedStudents=[],roster=read.operation.getSheetByName(STUDENT_SHEET_NAME);
  if(roster && roster.getLastRow()>1){const values=roster.getDataRange().getDisplayValues(),h=createHeaderMap_(values[0]);
    ['メールアドレス','名前','コースID'].forEach(function(k){if(h[k]==null)throw new Error('生徒一覧に必要な列がありません: '+k);});
    readStudentRosterRecords_(roster).forEach(function(s){const c=classes.get(s.courseId)||{id:s.courseId,name:s.className||s.courseId,students:[]};if(s.excluded)excludedStudents.push({courseId:s.courseId,email:s.email});else c.students.push({email:s.email,name:s.name,number:s.number,no:s.no});classes.set(s.courseId,c);});
  }
  const mappings=readMappingEntries_(true);
  getScoringTargets_(null,read).forEach(function(target){
    const mapping=mappings.get(makeMappingLookupKey_(target.spreadsheetId,target.sheetName)),courseId=gradeText_(mapping && mapping.courseId || target.courseId);
    if(base.classIds.length && courseId && !base.classIds.includes(courseId))return;
    const snapshot=scoringSnapshot_(read,target),values=snapshot.values;if(!values.length)return;
    const names=values[0].map(String),h=createHeaderMap_(names),display=snapshot.sheet.getDataRange().getDisplayValues();names.forEach(function(n){if(n!==SCORING_MANAGEMENT_HEADER)headers.add(n);});
    if(!strict)return;
    if(!courseId){issues.push({type:'mapping',message:target.label+' のクラス対応を設定の「対応表」で指定してください。'});return;}
    if(!classes.has(courseId)){issues.push({type:'roster',message:target.label+' のクラスが生徒一覧にありません。'});return;}
    // 全員除外の回答タブはメールだけで判定し、不要な過去の評価列・日付列で出力を止めない。
    if(!classes.get(courseId).students.length && names.filter(function(n){return n===read.app.emailHeader;}).length===1 &&
      values.slice(1).filter(function(row){return row.some(function(v){return v!=='' && v!=null;});}).every(function(row){return excludedStudents.some(function(s){return s.courseId===courseId && s.email===gradeText_(row[h[read.app.emailHeader]]).trim().toLowerCase();});}))return;
    const required=[read.app.emailHeader,base.dateHeader,base.valueHeader].concat(base.leading.filter(function(k){return k.startsWith('field:');}).map(function(k){return k.slice(6);}));
    if(required.some(function(k){return !k || names.filter(function(n){return n===k;}).length!==1;})){issues.push({type:'columns',courseId:courseId,message:target.label+' に選択した列がないか、列名が重複しています。'});return;}
    const timezone=scoringBook_(read,target.spreadsheetId).getSpreadsheetTimeZone();
    values.slice(1).forEach(function(row,index){if(row.every(function(v){return v==='' || v==null;}))return;
      const fields={};names.forEach(function(n,col){if(n!==SCORING_MANAGEMENT_HEADER && !['__proto__','constructor','prototype'].includes(n))fields[n]=display[index+1][col];});
      const email=gradeText_(row[h[read.app.emailHeader]]).trim().toLowerCase(),date=scoringLessonDate_(row[h[base.dateHeader]],timezone),id=scoringResponseKey_(target,index+2);
      responses.push({id:id,courseId:courseId,email:email,date:date,rawDate:display[index+1][h[base.dateHeader]],submitted:display[index+1][h['タイムスタンプ']]||display[index+1][h.Timestamp]||'',source:target.sheetName+' '+(index+2)+'行',fields:fields,fingerprint:gradeDigest_([id,row.filter(function(_,i){return names[i]!==SCORING_MANAGEMENT_HEADER && names[i]!==read.app.formStatusHeader;})])});
    });
  });
  return {classes:Array.from(classes.values()),headers:Array.from(headers).sort(),responses:responses,issues:issues,excludedStudents:excludedStudents};
}
function gradeGetConsole() {
  assertWebOperator_();return withAppLock_(function(){const store=gradeReadStore_();return {store:store,data:gradeReadData_(store.value.draft.base,false)};});
}
function gradeSave(payload) {
  assertWebOperator_();return withAppLock_(function(){const current=gradeReadStore_();if(payload.revision!==current.revision)throw new Error('設定が変更されました。読み込み直してください。');const value=current.value;
    if(payload.action==='delete'){value.presets=value.presets.filter(function(p){return p.id!==payload.id;});}
    else {const profile=gradeValidate_(payload.profile);value.draft=profile;
      if(payload.action==='preset'){const name=gradeText_(payload.name).trim();if(!name||name.length>100)throw new Error('プリセット名は1〜100文字で入力してください。');let preset=value.presets.find(function(p){return p.id===payload.id;});if(!preset){preset={id:Utilities.getUuid()};value.presets.push(preset);}Object.assign(preset,{name:name,profile:profile});}
      else if(payload.action!=='draft')throw new Error('保存操作が不正です。');
    }
    return gradeWriteStore_(value,current.revision);
  });
}
function gradePreview(profile) {
  assertWebOperator_();const p=gradeValidate_(profile),data=gradeReadData_(p.base,true),result=gradeBuild_(data,p);
  result.revision=gradeDigest_([p,data]);result.responses=data.responses;return result;
}
function gradeDownload(profile, expected) {
  assertWebOperator_();const result=gradePreview(profile);
  if(result.revision!==expected)throw new Error('プレビュー後に回答・名簿・成績が変更されました。もう一度確認してください。');
  if(result.issues.length)throw new Error('未解決の確認事項があります。先に確認してください。');
  const files=result.classes.map(function(c,i){const name=String(c.name).replace(/[<>:"/\\|?*\u0000-\u001f]/g,'_').slice(0,80)||'クラス';return Utilities.newBlob(gradeCsv_(c.rows),'text/csv',(i+1)+'-'+name+'.csv');});
  const file=files.length===1?files[0]:Utilities.zip(files,'成績出力.zip');
  return {filename:file.getName(),mime:files.length===1?'text/csv':'application/zip',base64:Utilities.base64Encode(file.getBytes())};
}
function gradeConfigurationKind_(kind) { return ['grade-all','grade-layout','grade-rules'].includes(kind); }
function gradeExportConfiguration_(kind) {
  const p=gradeReadStore_().value.draft,payload={application:'turret',type:kind,schemaVersion:1,exportedAt:new Date().toISOString()};
  if(kind!=='grade-rules')payload.base=p.base;
  if(kind!=='grade-layout'){payload.calendar=p.calendar;payload.rules=p.rules;}
  const json=JSON.stringify(payload,null,2);if(utf8ByteLength_(json)>1048576)throw new Error('JSONが1MiBを超えました。プリセットを分けてください。');
  return {filename:configurationFilename_(kind),json:json};
}
function gradeParseConfiguration_(kind,json) {
  if(typeof json!=='string'||utf8ByteLength_(json)>1048576)throw new Error('成績出力JSONは1MiB以内で指定してください。');
  const p=JSON.parse(json);profileKeys_(p,['application','type','schemaVersion','exportedAt'].concat(kind!=='grade-rules'?['base']:[],kind!=='grade-layout'?['calendar','rules']:[]),'成績出力JSON');
  if(p.application!=='turret'||p.type!==kind||p.schemaVersion!==1)throw new Error('成績出力JSONの種類または版が不正です。');
  const profile=gradeDefault_();if(kind!=='grade-rules')profile.base=p.base;if(kind!=='grade-layout'){profile.calendar=p.calendar;profile.rules=p.rules;}
  return gradeValidate_(profile);
}
function gradeMergeRules_(current,incoming,choices) {
  const rules=JSON.parse(JSON.stringify(current.rules)),calendar=current.calendar.slice(),conflicts=[];
  function merge(list,extra,idOf,category){extra.forEach(function(row){const key=idOf(row),index=list.findIndex(function(x){return idOf(x)===key;});if(index<0){list.push(row);return;}if(gradeDigest_(list[index])===gradeDigest_(row))return;const id=category+':'+key;conflicts.push({id:id,label:category+' '+key,current:JSON.stringify(list[index]),incoming:JSON.stringify(row)});if(choices && choices[id]==='incoming')list[index]=row;});}
  incoming.calendar.forEach(function(e){const old=calendar.find(function(x){return x.id===e.id;})||calendar.find(function(x){return x.title===e.title && x.date===e.date && x.time===e.time;});if(!old){calendar.push(e);return;}if(old.title===e.title && old.date===e.date && old.time===e.time)return;const id='授業日:'+old.id;conflicts.push({id:id,label:'授業日 '+old.title,current:JSON.stringify(old),incoming:JSON.stringify(e)});if(choices&&choices[id]==='incoming')calendar[calendar.indexOf(old)]=e;});
  ['dates','answers','duplicates','fields','lessons'].forEach(function(k){rules[k]=rules[k]||[];merge(rules[k],incoming.rules[k]||[],function(x){return k==='lessons'?gradeKey_([x.courseId,x.date]):k==='dates'?gradeKey_([x.courseId,x.from]):k==='answers'?x.id:x.key;},k);});
  return {rules:rules,calendar:calendar,conflicts:conflicts};
}
function gradeInspectConfiguration_(kind,json) {
  const incoming=gradeParseConfiguration_(kind,json),store=gradeReadStore_(),merged=kind==='grade-rules'?gradeMergeRules_(store.value.draft,incoming):{conflicts:[]};
  return {kind:kind,revision:gradeDigest_([kind,json,store.revision]),summary:kind==='grade-all'?'成績出力の新しいプリセットとして追加し、作業下書きに読み込みます。':kind==='grade-layout'?'出力列・並び順・クラス対応を置き換えます。補正ルールは保持します。':'授業日と補正ルールを追加・統合します。競合する対象は選択してください。',warnings:['名簿・回答・採点結果は含みません。クラス・列の対応を確認してから出力してください。'],conflicts:merged.conflicts};
}
function gradeApplyConfiguration_(kind,json,revision,choices) {
  return withAppLock_(function(){const incoming=gradeParseConfiguration_(kind,json),inspection=gradeInspectConfiguration_(kind,json),store=gradeReadStore_();
    if(inspection.revision!==revision)throw new Error('確認後に成績出力設定が変更されました。再読み込みしてください。');
    if(kind==='grade-all'){store.value.presets.push({id:Utilities.getUuid(),name:'読み込み '+new Date().toISOString().slice(0,10),profile:incoming});store.value.draft=incoming;}
    else if(kind==='grade-layout')store.value.draft.base=incoming.base;
    else {const merged=gradeMergeRules_(store.value.draft,incoming,choices);if(merged.conflicts.some(function(x){return !choices||!['current','incoming'].includes(choices[x.id]);}))throw new Error('競合する補正ルールの扱いを選んでください。');store.value.draft.rules=merged.rules;store.value.draft.calendar=merged.calendar;}
    gradeValidate_(store.value.draft);gradeWriteStore_(store.value,store.revision);return {message:'成績出力設定を読み込みました。成績出力タブで対応を確認してください。'};
  });
}

/** 最後に準備したひな形を優先し、JSONから読み込んだURLはその後の既定値として使う。 */
function currentTemplateFormUrl_(records) {
  records = records || getManagedRecords_();
  const spreadsheetId = getAppSpreadsheet_().getId(), scriptId = ScriptApp.getScriptId();
  let recordUrl = '', recordTime = 0;
  for (let i = records.length - 1; i >= 0; i--) {
    const record = records[i];
    if (record.ownerSpreadsheetId !== spreadsheetId || record.ownerScriptId !== scriptId) continue;
    const id = record.kind === 'batch' && record.input && record.input.templateId || record.kind === 'template' && record.stage === 'created' && record.formId;
    const time = Date.parse(record.updatedAt || '') || 0;
    if (id && time >= recordTime) { recordUrl = 'https://docs.google.com/forms/d/' + id + '/edit'; recordTime = time; }
  }
  const saved = PropertiesService.getScriptProperties().getProperty(TEMPLATE_FORM_URL_KEY_);
  if (!saved) return recordUrl;
  let preferred;
  try { preferred = JSON.parse(saved); } catch (error) { preferred = { url: saved, savedAt: '' }; }
  return preferred && typeof preferred.url === 'string' && (Date.parse(preferred.savedAt || '') || 0) >= recordTime ? preferred.url : recordUrl;
}

function savedTemplateFormUrl_(url) { return JSON.stringify({ url: url, savedAt: new Date().toISOString() }); }

function normalizeTemplateFormUrl_(value) {
  if (typeof value !== 'string') throw new Error('ひな形フォームのURLが不正です。');
  const match = value.trim().match(/^https:\/\/docs\.google\.com\/forms\/(?:u\/\d+\/)?d\/([\w-]{15,})\/edit\/?(?:[?#].*)?$/);
  if (!match) throw new Error('ひな形フォームの編集URLを指定してください。');
  return 'https://docs.google.com/forms/d/' + match[1] + '/edit';
}

/** Apps Script エディタから実行し、フォーム連携に必要な同意を求める。データは変更しない。 */
function authorizeFormIntegration() {
  ScriptApp.requireScopes(ScriptApp.AuthMode.FULL, [
    'https://www.googleapis.com/auth/forms',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/classroom.courseworkmaterials'
  ]);
  return { authorized: true };
}

function managedSheet_(create) {
  const ss = getAppSpreadsheet_();
  let sheet = ss.getSheetByName(MANAGED_FORM_SHEET_);
  const legacy=ss.getSheetByName('フォーム管理');
  if(sheet && legacy)throw new Error('システム管理とフォーム管理が両方あります。自動統合はできません。内容を確認してください。');
  if(!sheet && legacy) {
    if(legacy.getLastColumn()!==3 || JSON.stringify(legacy.getRange(1,1,1,3).getDisplayValues()[0])!==JSON.stringify(MANAGED_FORM_HEADERS_))throw new Error('フォーム管理シートの構成が異なります。');
    // 読み取りでは旧名も受理。共通ロックを持つ保存・初期化時に改名する。
    sheet=legacy;if(create)sheet.setName(MANAGED_FORM_SHEET_);
  }
  if (!sheet && create) {
    sheet = ss.insertSheet(MANAGED_FORM_SHEET_);
    sheet.getRange(1, 1, 1, 3).setValues([MANAGED_FORM_HEADERS_]);
    sheet.setFrozenRows(1);
  }
  if (sheet && (sheet.getLastColumn() !== 3 || JSON.stringify(sheet.getRange(1, 1, 1, 3).getDisplayValues()[0]) !== JSON.stringify(MANAGED_FORM_HEADERS_))) {
    throw new Error('システム管理シートの構成が異なります。既存データは変更していません。');
  }
  return sheet;
}

function getManagedRecords_() {
  const sheet = managedSheet_(false);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const ids = new Set();
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getDisplayValues().map(function(row) {
    let record;
    try { record = JSON.parse(row[2]); } catch (error) { throw new Error('システム管理の記録を読み取れません。'); }
    if (!record || record.schema !== 1 || !record.id || record.id !== row[1] || record.kind !== row[0] || ids.has(record.id)) {
      throw new Error('システム管理の記録形式・IDが不正です。');
    }
    ids.add(record.id);
    return record;
  });
}

function assertManagedOwner_(record) {
  if (record.ownerSpreadsheetId !== getAppSpreadsheet_().getId() || record.ownerScriptId !== ScriptApp.getScriptId()) {
    throw new Error('別の運用の記録です。複製元のフォーム・投稿は操作できません。');
  }
}

function loadManagedRecord_(id, revision) {
  const record = getManagedRecords_().find(function(item) { return item.id === id; });
  if (!record) throw new Error('管理対象の記録がありません。');
  assertManagedOwner_(record);
  if (revision != null && record.revision !== revision) throw new Error('記録が更新されました。状態を更新してください。');
  return record;
}

/** 呼び出し元はアプリ共通ロックを保持する。削除を含む外部変更より先に保存する。 */
function saveManagedRecord_(record, cache) {
  if (record.ownerSpreadsheetId) assertManagedOwner_(record);
  const records = cache && !cache.failed ? cache.records : getManagedRecords_();
  const at = records.findIndex(function(item) { return item.id === record.id; });
  if (at >= 0) assertManagedOwner_(records[at]);
  Object.assign(record, { schema: 1, ownerSpreadsheetId: getAppSpreadsheet_().getId(), ownerScriptId: ScriptApp.getScriptId(),
    revision: Utilities.getUuid(), updatedAt: new Date().toISOString() });
  const json = JSON.stringify(record);
  if (json.length > 40000) throw new Error('管理情報が大きすぎます。説明文・対象を短くしてください。');
  try {
    const sheet = cache && cache.sheet || managedSheet_(true);
    if(cache)cache.sheet=sheet;
    sheet.getRange(at < 0 ? records.length + 2 : at + 2, 1, 1, 3).setValues([[record.kind, record.id, json]]);
    SpreadsheetApp.flush();
    if(cache){if(at<0)records.push(record);else records[at]=record;cache.records=records;cache.failed=false;}
  } catch(error){if(cache)cache.failed=true;throw error;}
  return record;
}

/** 一定のクライアント要求キーを使い、再試行でひな形が重複作成されるのを防ぐ。 */
function createTemplateForm(requestKey) {
  if (typeof requestKey !== 'string' || !/^[A-Za-z0-9_-]{16,100}$/.test(requestKey)) {
    throw new Error('ひな形作成の識別情報が不正です。画面を開き直してください。');
  }
  return withAppLock_(function() {
    const id = 'template:' + requestKey;
    const existing = getManagedRecords_().find(function(record) { return record.id === id; });
    if (existing) {
      assertManagedOwner_(existing);
      if (existing.kind === 'template' && existing.stage === 'created' && existing.formId) {
        return { editUrl: 'https://docs.google.com/forms/d/' + existing.formId + '/edit' };
      }
      if (existing.kind === 'template' && existing.stage === 'placing' && existing.formId && existing.folderId) {
        return placeTemplateForm_(existing);
      }
      if (existing.kind === 'template' && existing.stage === 'initializing' && existing.formId && existing.folderId) {
        return initializeTemplateForm_(existing);
      }
      throw new Error('ひな形の作成結果を確認できません。二重作成を防ぐため停止しました。マイドライブの「turret ひな形」を確認し、作成済みなら編集URLを入力してください。');
    }
    const parents = DriveApp.getFileById(getAppSpreadsheet_().getId()).getParents();
    if (!parents.hasNext()) throw new Error('turretスプレッドシートの保存先フォルダを取得できません。Driveの配置とアクセス権を確認してください。');
    const record = { id: id, kind: 'template', stage: 'creating', folderId: parents.next().getId() };
    saveManagedRecord_(record);
    // 公開しない。確認済みメールアドレスの収集はフォームエディタで設定する必要がある。
    const form = FormApp.create('turret ひな形', false);
    record.formId = form.getId();
    const editUrl = 'https://docs.google.com/forms/d/' + record.formId + '/edit';
    record.stage = 'initializing';
    try { saveManagedRecord_(record); }
    catch (error) { throw new Error('ひな形は作成済みですが履歴を保存できませんでした。次の編集URLを入力してください：' + editUrl); }
    return initializeTemplateForm_(record, form);
  });
}

/** 呼び出し元はロックを保持する。最初の日付項目の作成が中断された場合、重複させずに復旧する。 */
function initializeTemplateForm_(record, form) {
  const title = '授業の日付を入力してください';
  try {
    form = form || FormApp.openById(record.formId);
    const items = form.getItems();
    let item;
    if (record.lessonDateItemId != null) {
      item = items.find(function(candidate) { return String(candidate.getId()) === String(record.lessonDateItemId); });
    } else if (!items.length) {
      item = form.addDateItem();
    } else if (items.length === 1 && (!items[0].getTitle() || items[0].getTitle() === title)) {
      item = items[0];
    }
    if (!item || String(item.getType()) !== 'DATE' || (item.getTitle() && item.getTitle() !== title)) {
      throw new Error('ひな形の設問が変更されています。編集画面で授業日の項目を確認してください。');
    }
    record.lessonDateItemId = item.getId();
    saveManagedRecord_(record);
    const dateItem = typeof item.asDateItem === 'function' ? item.asDateItem() : item;
    dateItem.setTitle(title).setHelpText('回答する日ではなく、振り返りの対象となる授業の日付を入力してください。')
      .setIncludesYear(true).setRequired(true);
    record.stage = 'placing';
    saveManagedRecord_(record);
  } catch (error) {
    throw new Error('ひな形は作成済みですが初期設定が完了していません。同じ画面から再試行してください。編集URL：https://docs.google.com/forms/d/' + record.formId + '/edit\n' + String(error && error.message || error));
  }
  return placeTemplateForm_(record);
}

/** 呼び出し元はロックを保持する。移動を試みる前にフォーム ID を保存する。 */
function placeTemplateForm_(record) {
  const editUrl = 'https://docs.google.com/forms/d/' + record.formId + '/edit';
  try {
    const file = DriveApp.getFileById(record.formId), parents = file.getParents();
    let placed = false;
    while (parents.hasNext()) {
      if (parents.next().getId() === record.folderId) { placed = true; break; }
    }
    if (!placed) file.moveTo(DriveApp.getFolderById(record.folderId));
    record.stage = 'created';
    saveManagedRecord_(record);
  } catch (error) {
    throw new Error('ひな形は作成済みですが、同じフォルダへの配置完了を確認できません。フォルダのアクセス権を確認し、同じ画面で再試行してください。編集URL：' + editUrl);
  }
  return { editUrl: editUrl };
}

function getFormConsoleData(options) {
  return withAppLock_(function() {
    const config = getConfig_();
    let classes = [], classWarning = '';
    try { classes = getSelectedClassRecords_(); } catch (error) { classWarning = String(error.message || error); }
    const ss = getAppSpreadsheet_();
    const records = getManagedRecords_();
    // 開いたままの旧管理画面からの呼び出しには、従来の応答形式を維持する。
    let visibleRecords=records.filter(function(r){return !r.kind.startsWith('grade-');}), returnHistory;
    if(options && options.returnPage != null) {
      if(!Number.isSafeInteger(options.returnPage) || options.returnPage<0)throw new Error('履歴のページ番号が不正です。');
      const history=records.filter(function(r){return r.kind==='announcement';}).reverse(), pageSize=50;
      const page=Math.min(options.returnPage,Math.max(0,Math.ceil(history.length/pageSize)-1));
      returnHistory={page:page,pageSize:pageSize,total:history.length};
      visibleRecords=records.filter(function(r){return r.kind!=='announcement' && !r.kind.startsWith('grade-');}).concat(history.slice(page*pageSize,(page+1)*pageSize).reverse());
    }
    const result = { records: visibleRecords, scheduleTimer: typeof getManagedFormScheduleSummary_==='function'?getManagedFormScheduleSummary_():null, setupProgress: getFormSetupProgress_(config.formSources.length > 0 && config.formSheetNamePrefix.length > 0, records), classes: classes, classWarning: classWarning, configRevision: getConfigRevision_(config),
      spreadsheetId: ss.getId(), spreadsheetUrl: 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/edit',
      defaults: { emailHeader: config.emailHeader, nameHeader: config.studentNameHeader, gradeHeader: config.scoreSourceHeader, scoreSourceHeader: config.scoreSourceHeader, statusHeader: config.formStatusHeader, templateUrl: currentTemplateFormUrl_(records) } };
    if(returnHistory)result.returnHistory=returnHistory;
    return result;
  });
}

/** 設定の進捗表示では内部記録だけを読み、Forms や Classroom には接続しない。 */
function getFormSetupProgress_(sourceReady, records) {
  const managed = (records || getManagedRecords_()).filter(function(r) {
    return r.kind === 'form' && r.ownerSpreadsheetId === getAppSpreadsheet_().getId() && r.ownerScriptId === ScriptApp.getScriptId();
  });
  const result = function(formsState, formsDetail, publishState, publishDetail) {
    return { forms: { id:'forms', title:'フォームを準備', state:formsState, detail:formsDetail },
      publish: { id:'publish', title:'資料を投稿', state:publishState, detail:publishDetail } };
  };
  if (!managed.length) return result(sourceReady?'complete':'pending',sourceReady?'既存の回答先を登録済み':'ひな形から作成、または既存のフォームを登録します',
    sourceReady?'complete':'pending',sourceReady?'既存フォームの共有はClassroom側で確認してください':'フォームの準備後に投稿します');
  const updating = managed.filter(function(r) {return r.formUpdate||r.settingsUpdate;}).length;
  if (updating) return result('attention',updating+'件のフォームが更新途中です','attention','一覧の更新再開ボタンでフォーム・設定の更新を完了してください');
  const unfinished = managed.filter(function(r) { return ['registered','published','scheduled','draft','schedule_review','deleted','publish_review','delete_review'].indexOf(r.stage) < 0; }).length;
  const published = managed.filter(function(r) { return r.stage === 'published'; }).length;
  const scheduled = managed.filter(function(r) { return r.stage === 'scheduled'; }).length;
  const uncertain = managed.some(function(r) { return ['publish_review','schedule_review','delete_review'].indexOf(r.stage) >= 0; });
  return result(unfinished?'attention':sourceReady?'complete':'pending',unfinished?unfinished+'件のフォームが準備途中です':sourceReady?'フォームと回答先を準備済み':'回答先の登録を確認してください',
    uncertain?'attention':published===managed.length?'complete':'pending',published+' / '+managed.length+'件の資料を投稿済み'+(scheduled?'・'+scheduled+'件を予約済み':'')+(uncertain?'・結果の照合が必要です':''));
}

function managedId_(value) {
  const text = String(value || '').trim();
  const match = text.match(/\/(?:d|folders)\/([\w-]+)/);
  const id = match ? match[1] : text;
  if (!/^[\w-]{15,}$/.test(id)) throw new Error('フォーム・フォルダのURLまたはIDを確認してください。');
  return id;
}

function normalizeFormSetup_(raw) {
  if (!raw || typeof raw !== 'object') throw new Error('作成設定がありません。');
  const input = { templateId: managedId_(raw.templateId), folderId: managedId_(raw.folderId) };
  if (raw.verifiedEmailConfirmed !== true) throw new Error('ひな形の設定画面でメールアドレスの収集が「確認済み」であることを確認し、確認欄にチェックしてください。');
  input.verifiedEmailConfirmed = true;
  // 新しい画面では投稿設定を準備から分離する。旧画面の入力形式も引き続き受け付ける。
  if (raw.deferMaterialSettings === true) {
    input.deferMaterialSettings = true;
    raw = Object.assign({}, raw, {materialTitlePattern:raw.titlePattern,description:''});
  }
  const legacy = !Object.prototype.hasOwnProperty.call(raw, 'additionalColumns') && raw.gradeHeader !== undefined;
  ['titlePattern','materialTitlePattern','description','prefix','emailHeader','nameHeader','statusHeader'].concat(legacy ? ['gradeHeader','commentHeader'] : []).forEach(function(key) {
    input[key] = String(raw[key] || '').trim();
    if ((!input[key] && key !== 'description') || input[key].length > (key === 'description' ? 5000 : 150)) throw new Error('入力を確認してください: ' + key);
  });
  if (![input.titlePattern, input.materialTitlePattern].every(function(title) { return title.includes('{class_label}') || title.includes('{class}'); })) throw new Error('タイトルに {class_label} を含めてください。');
  if (legacy) input.gradeChoices = normalizeManagedChoices_(String(raw.gradeChoices || '').split(/[\s,、]+/));
  else input.additionalColumns = normalizeManagedColumns_(raw.additionalColumns === undefined ? [] : raw.additionalColumns);
  const headers = [input.emailHeader,input.nameHeader,input.statusHeader,SCORING_MANAGEMENT_HEADER].concat(managedCustomColumns_(input).map(function(c) { return c.header; }));
  if (new Set(headers).size !== headers.length) throw new Error('列名が重複しています。');
  function list(value) { return Array.from(new Set(String(value || '').split(/[\s,、]+/).map(function(v) { return v.trim().toLowerCase(); }).filter(Boolean))); }
  input.policy = { domains: list(raw.domains), emails: list(raw.emails) };
  if (!input.policy.domains.length && !input.policy.emails.length) throw new Error('学校ドメインまたは回答を許可するメールを入力してください。');
  if (input.policy.domains.some(function(d) { return !/^[a-z0-9.-]+\.[a-z]{2,}$/.test(d) || ['gmail.com','googlemail.com'].includes(d); }) ||
      input.policy.emails.some(function(e) { return !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e); })) throw new Error('回答者の指定が不正です。個人アカウントはメールで指定してください。');
  input.labels = {};
  // 確認と開始のRPCでキー順が変わっても、同じクラス表記は同じ変更判定値にする。
  Object.keys(raw.labels || {}).sort().forEach(function(k) { input.labels[k] = String(raw.labels[k] || '').trim(); });
  return input;
}

function normalizeManagedChoices_(value) {
  if (typeof value !== 'string' && !Array.isArray(value)) throw new Error('入力用プルダウンの選択肢は文字列または配列で指定してください。');
  const values = typeof value === 'string' ? value.split(/[,、\n]+/) : value;
  if (values.some(function(v) { return typeof v !== 'string'; })) throw new Error('選択肢は文字列で入力してください。');
  const result = Array.from(new Set(values.map(function(v) { return v.trim(); }).filter(Boolean)));
  if (result.length > 30 || result.some(function(v) { return v.length > 150; })) throw new Error('選択肢は30個以内・各150文字以内にしてください。');
  return result;
}

function normalizeManagedColumns_(columns) {
  if (!Array.isArray(columns) || columns.length > 50) throw new Error('追加列は50列以内の配列で指定してください。');
  return columns.map(function(column) {
    if (!column || typeof column.header !== 'string' || !column.header.trim() || column.header.trim().length > 150) throw new Error('追加列の名前を1〜150文字で入力してください。');
    return { header: column.header.trim(), choices: normalizeManagedChoices_(column.choices === undefined ? [] : column.choices) };
  });
}

/** 未完了の処理を再開しても、既存の記録行の順序を変えない。 */
function managedCustomColumns_(input) {
  return Array.isArray(input.additionalColumns) ? input.additionalColumns : [
    {header:input.gradeHeader,choices:input.gradeChoices || []}, {header:input.commentHeader,choices:[]}
  ];
}

function managedSourceConfig_(input, config) {
  const pairs = [['emailHeader','emailHeader'],['studentNameHeader','nameHeader'],['formStatusHeader','statusHeader']];
  const legacy = !Array.isArray(input.additionalColumns);
  if (legacy) pairs.push(['scoreSourceHeader','gradeHeader']);
  pairs.forEach(function(pair) {
    if (config[pair[0]] && config[pair[0]] !== input[pair[1]]) throw new Error('既存の列設定と一致しません: ' + pair[0] + '。「列と項目」で確認してください。');
    config[pair[0]] = input[pair[1]];
  });
  if (!config.replyBodyHeader) config.replyBodyHeader = '返信本文';
  if (legacy && !config.fields.length) config.fields = [
    {key:'score',sourceHeader:input.gradeHeader,evalHeader:input.gradeHeader,sendHeader:input.gradeHeader,type:'text'},
    {key:'comment',sourceHeader:input.commentHeader,evalHeader:input.commentHeader,sendHeader:input.commentHeader,type:'text'}
  ];
  if (!legacy) managedCustomColumns_(input).forEach(function(column) {
    if (config.fields.some(function(field) { return field.sourceHeader === column.header; })) return;
    let n = 1;
    while (config.fields.some(function(field) { return field.key === 'field_' + n; })) n++;
    config.fields.push({key:'field_' + n,sourceHeader:column.header,evalHeader:column.header,sendHeader:column.header,type:'text'});
  });
  validateConfigDraft_(config);
  validateConfigSheetLayouts_(config);
  return config;
}

function formMetadata_(id) {
  const form = FormApp.openById(id);
  const metadata = {engine:'formapp-v1',info:{title:form.getTitle(),description:form.getDescription() || ''},
    settings:{collectsEmail:form.collectsEmail(),quizSettings:{isQuiz:form.isQuiz()}},items:form.getItems().map(nativeFormItem_)};
  metadata.revisionId = adminDigest_(metadata);
  return metadata;
}

const NATIVE_FORM_TYPES_ = {TEXT:'Text',PARAGRAPH_TEXT:'ParagraphText',MULTIPLE_CHOICE:'MultipleChoice',CHECKBOX:'Checkbox',LIST:'List',
  SCALE:'Scale',GRID:'Grid',CHECKBOX_GRID:'CheckboxGrid',DATE:'Date',DATETIME:'DateTime',TIME:'Time',DURATION:'Duration',
  SECTION_HEADER:'SectionHeader',PAGE_BREAK:'PageBreak',IMAGE:'Image',VIDEO:'Video'};

function nativeTypedItem_(item) {
  const suffix = NATIVE_FORM_TYPES_[String(item.getType())];
  return suffix ? item['as' + suffix + 'Item']() : item;
}

/** FormApp が公開する項目だけを読む。collectsEmail だけで「確認済み」と判定しない。 */
function nativeFormItem_(generic) {
  const type = String(generic.getType()), item = nativeTypedItem_(generic);
  const result = {itemId:String(item.getId()),title:item.getTitle() || '',description:item.getHelpText() || '',native:{type:type}};
  const n = result.native;
  if (['TEXT','PARAGRAPH_TEXT','MULTIPLE_CHOICE','CHECKBOX','LIST','SCALE','GRID','CHECKBOX_GRID','DATE','DATETIME','TIME','DURATION'].includes(type)) n.required = item.isRequired();
  if (['MULTIPLE_CHOICE','CHECKBOX','LIST'].includes(type)) {
    n.choices = item.getChoices().map(function(c) {
      const option = {value:c.getValue()};
      if (type === 'MULTIPLE_CHOICE') {
        const page = c.getGotoPage(), action = c.getPageNavigationType();
        if (page) option.goToSectionId = String(page.getId());
        else if (action) option.goToAction = String(action);
      }
      return option;
    });
    if (type !== 'LIST') n.other = item.hasOtherOption();
  }
  if (['GRID','CHECKBOX_GRID'].includes(type)) {n.rows = item.getRows();n.columns = item.getColumns();}
  if (type === 'SCALE') {n.bounds=[item.getLowerBound(),item.getUpperBound()];n.labels=[item.getLeftLabel() || '',item.getRightLabel() || ''];}
  if (['DATE','DATETIME'].includes(type)) n.includesYear = item.includesYear();
  if (type === 'PAGE_BREAK') {
    const page = item.getGoToPage();n.navigation = page ? {goToSectionId:String(page.getId())} : {goToAction:String(item.getPageNavigationType() || 'CONTINUE')};
  }
  if (type === 'IMAGE') {const blob=item.getImage();n.imageHash=blob ? adminDigest_(blob.getBytes()) : '';}
  if (['IMAGE','VIDEO'].includes(type)) {n.width=item.getWidth();n.alignment=String(item.getAlignment());}
  return result;
}

function publishedPermissions_(id) {
  const permissions = []; let token;
  do {
    const options = { includePermissionsForView: 'published', fields: 'nextPageToken,permissions(id,type,role,view,emailAddress,domain,deleted)', pageSize: 100 };
    if (token) options.pageToken = token;
    const page = Drive.Permissions.list(id, options);
    (page.permissions || []).forEach(function(p) { if (!p.deleted && p.view === 'published') permissions.push(p); });
    token = page.nextPageToken;
  } while (token);
  return permissions;
}

function validateResponderAccess_(permissions, policy, students, collectsEmail) {
  if (collectsEmail !== true) throw new Error('フォームのメール収集を有効にしてください。');
  const allowed = permissions.filter(function(p) { return p.view === 'published'; });
  const norm = function(v) { return String(v || '').toLowerCase(); };
  if (!allowed.length || allowed.some(function(p) {
    return p.role !== 'reader' || (p.type === 'domain' ? !policy.domains.includes(norm(p.domain)) : p.type === 'user' ? !policy.emails.includes(norm(p.emailAddress)) && !policy.domains.includes(norm(p.emailAddress).split('@')[1]) : true);
  })) throw new Error('指定外または確認できない回答者権限があります。回答範囲を確認してください。');
  if (!students.length || students.some(function(email) {
    return !allowed.some(function(p) { return p.type === 'domain' ? norm(p.domain) === norm(email).split('@')[1] : norm(p.emailAddress) === norm(email); });
  })) throw new Error('対象クラスの生徒全員に回答権限がありません。生徒一覧と回答範囲を確認してください。');
  return true;
}

function managedRoster_(courseId, context) {
  const values = context && context.rosterValues || adminSheetValues_(getAppSpreadsheet_().getSheetByName(STUDENT_SHEET_NAME));
  if(context)context.rosterValues=values;
  const map = createHeaderMap_(values[0] || []);
  if (['メールアドレス','名前','コースID'].some(function(k) { return !(k in map); })) throw new Error('生徒一覧を取得してください。');
  // 除外者は権限確認・氏名補完の対象にしない。氏名競合は既存の補完処理で空欄のまま残す。
  return values.slice(1).filter(function(row) { return String(row[map['コースID']]) === courseId && String(row[map['除外(1)']]||'').trim()!=='1'; }).map(function(row) {
    return { email: String(row[map['メールアドレス']] || '').trim().toLowerCase(), name: String(row[map['名前']] || '').trim() };
  }).filter(function(student) { return !!student.email; });
}

function previewFormSetup(raw, expectedConfigRevision) {
  return withAppLock_(function() { return previewFormSetupUnlocked_(raw, expectedConfigRevision); });
}

function previewFormSetupUnlocked_(raw, expectedConfigRevision) {
  const config = getConfig_();
  if (expectedConfigRevision !== getConfigRevision_(config)) throw new Error('設定が更新されました。状態を更新してください。');
  const input = normalizeFormSetup_(raw), ss = getAppSpreadsheet_();
  managedSourceConfig_(input, config); // 作成を始める前に列の登録内容を検証する。
  const template = FormApp.openById(input.templateId);
  if (!template.supportsAdvancedResponderPermissions()) throw new Error('ひな形を新しい公開方式のフォームへ更新してください。');
  const metadata = formMetadata_(input.templateId);
  if (!metadata.settings || !metadata.settings.collectsEmail) throw new Error('ひな形のメール収集を有効にしてください。');
  DriveApp.getFolderById(input.folderId).getName();
  const labels = new Set();
  const targets = getSelectedClassRecords_().map(function(c) {
    const label = input.labels[c.courseId] || c.className;
    const sheetName = input.prefix + ' ' + label;
    if (!label || sheetName.length > 100 || /[\[\]:*?\/\\]/.test(sheetName) || labels.has(sheetName) || ss.getSheetByName(sheetName)) throw new Error('回答タブ名が重複または不正です: ' + sheetName);
    labels.add(sheetName);
    const students = managedRoster_(c.courseId);
    if (!students.length) throw new Error(c.className + 'の生徒一覧を取得してください。');
    const outside = students.some(function(s) { return !input.policy.emails.includes(s.email) && !input.policy.domains.includes(s.email.split('@')[1]); });
    if (outside) throw new Error(c.className + 'の生徒を回答許可の範囲に含めてください。');
    return {courseId:c.courseId,className:c.className,label:label,sheetName:sheetName,
      title:input.titlePattern.replace(/\{class(?:_label)?\}/g,function(){return label;}),materialTitle:input.materialTitlePattern.replace(/\{class(?:_label)?\}/g,function(){return label;}), students: students.map(function(s) {return s.email;}).sort()};
  });
  const preview = { input: input, targets: targets, configRevision: expectedConfigRevision,
    templateRevision: DriveApp.getFileById(input.templateId).getLastUpdated().toISOString(), ownerSpreadsheetId:ss.getId() };
  preview.fingerprint = adminDigest_(preview);
  return preview;
}

function beginFormSetup(raw, expectedRevision, fingerprint) {
  return withAppLock_(function() {
    const records=getManagedRecords_();
    let batch=records.find(function(r){return r.kind==='batch'&&r.requestFingerprint===fingerprint;});
    if(batch)assertManagedOwner_(batch);
    else {
      const p = previewFormSetupUnlocked_(raw, expectedRevision);
      if (p.fingerprint !== fingerprint) throw new Error('プレビュー後に設定が更新されました。再確認してください。');
      batch=saveManagedRecord_({id:Utilities.getUuid(),kind:'batch',input:p.input,targets:p.targets,stage:'prepared',requestFingerprint:fingerprint});
    }
    return batch.targets.map(function(t) {
      const id=batch.id+':'+t.courseId,existing=records.find(function(r){return r.id===id;});
      if(existing){assertManagedOwner_(existing);return existing;}
      return saveManagedRecord_(Object.assign({},t,{id:id,kind:'form',jobId:batch.id,input:batch.input,stage:'pending',requestFingerprint:fingerprint},batch.input.deferMaterialSettings?{materialSettingsSaved:false}:{}));
    });
  });
}

function ensureCopiedResponderPolicy_(record) {
  const allowed = record.input.policy;
  const current = publishedPermissions_(record.formId);
  current.forEach(function(p) {
    const keep = p.role === 'reader' && (p.type === 'domain' ? allowed.domains.includes(String(p.domain).toLowerCase()) : p.type === 'user' && allowed.emails.includes(String(p.emailAddress).toLowerCase()));
    if (!keep) Drive.Permissions.remove(record.formId, p.id);
  });
  const desired = allowed.domains.map(function(domain) {return {type:'domain',domain:domain};}).concat(allowed.emails.map(function(email) {return {type:'user',emailAddress:email};}));
  desired.forEach(function(p) {
    if (!current.some(function(old) {return old.role==='reader' && old.type===p.type && (p.type==='domain' ? old.domain===p.domain : old.emailAddress===p.emailAddress);})) {
      Drive.Permissions.create(Object.assign({role:'reader',view:'published'},p),record.formId,{sendNotificationEmail:false});
    }
  });
}

function managedAnswerSheet_(record, context) {
  const ss = getAppSpreadsheet_();
  const sheets = context && context.sheets || ss.getSheets();
  if(context)context.sheets=sheets;
  if (record.responseSheetId != null) {
    const found = sheets.find(function(s) { return String(s.getSheetId()) === String(record.responseSheetId); });
    if (found) return found;
    throw new Error('登録済み回答タブが見つかりません。');
  }
  return sheets.find(function(sheet) {
    const url = sheet.getFormUrl();
    return url && FormApp.openByUrl(url).getId() === record.formId;
  });
}

function managedDestinationId_(form) {
  try {return form.getDestinationId();}
  catch (error) {
    // 複製直後の未接続は正常。GASの英語・日本語環境どちらでも回答先の接続へ進む。
    if (/^(?:Exception: )?(?:The form currently has no response destination\.?|フォームに応答先がありません。?)$/.test(String(error.message || error))) return '';
    throw error;
  }
}

function prepareFormTarget(id, expectedRevision) {
  return withAppLock_(function() {
    const r = loadManagedRecord_(id,expectedRevision);
    if (['registered','published','scheduled','draft','deleted'].includes(r.stage)) return r;
    if (r.stage === 'creating') throw new Error('フォーム作成結果が未確認です。Driveを確認し、作成済みフォームIDを照合してください。');
    if (r.stage === 'pending') {
      r.templateItems = formMetadata_(r.input.templateId).items || [];
      r.stage='creating';saveManagedRecord_(r);
      const copy = Drive.Files.copy({name:r.title,parents:[r.input.folderId],appProperties:{turretRecord:r.id,turretScript:ScriptApp.getScriptId()}},r.input.templateId,{fields:'id'});
      if(!copy||!copy.id)throw new Error('作成したフォームIDを確認できません。作成結果の照合が必要です。');
      r.formId=copy.id;r.stage='created';saveManagedRecord_(r);
    }
    const form=FormApp.openById(r.formId);
    if (!form.supportsAdvancedResponderPermissions()) throw new Error('フォームの公開方式を確認してください。');
    form.setPublished(false);
    if (r.stage==='created') {
      form.setTitle(r.title);ensureCopiedResponderPolicy_(r);
      r.stage='linking';saveManagedRecord_(r);
    }
    if (r.stage==='linking') {
      const destination=managedDestinationId_(form);
      if (destination && destination!==getAppSpreadsheet_().getId()) throw new Error('回答先が別のスプシです。フォームの設定を確認してください。');
      if (!destination) form.setDestination(FormApp.DestinationType.SPREADSHEET,getAppSpreadsheet_().getId());
      r.stage='linked';saveManagedRecord_(r);
    }
    const sheet=managedAnswerSheet_(r);
    if (!sheet) {r.lastError='回答タブの生成待ちです。「準備を続ける」で再確認してください。';return saveManagedRecord_(r);}
    r.responseSheetId=sheet.getSheetId();
    const collision=getAppSpreadsheet_().getSheetByName(r.sheetName);
    if (collision && collision.getSheetId()!==sheet.getSheetId()) throw new Error('同名の回答タブがあります。');
    if (sheet.getName()!==r.sheetName) sheet.setName(r.sheetName);
    const headers=sheet.getRange(1,1,1,Math.max(1,sheet.getLastColumn())).getDisplayValues()[0];
    if (!headers.includes(r.input.emailHeader)) throw new Error('回答タブにメール列「'+r.input.emailHeader+'」がありません。列名を確認してください。');
    if (new Set(headers.filter(Boolean)).size!==headers.filter(Boolean).length) throw new Error('フォームの質問見出しが重複しています。');
    const custom = managedCustomColumns_(r.input);
    const legacyAdditional=[r.input.nameHeader].concat(custom.map(function(c) { return c.header; }), [r.input.statusHeader]);
    const legacyTail=r.baseColumnCount && JSON.stringify(headers.slice(r.baseColumnCount,r.baseColumnCount+legacyAdditional.length))===JSON.stringify(legacyAdditional);
    const additional=legacyTail?legacyAdditional:[r.input.nameHeader].concat(custom.map(function(c) { return c.header; }), [SCORING_MANAGEMENT_HEADER,r.input.statusHeader]);
    if (!r.baseColumnCount) {
      if (additional.some(function(h){return headers.includes(h);})) throw new Error('追加列と質問名が重複しています。');
      r.baseColumnCount=headers.length;r.baseHeaders=headers;saveManagedRecord_(r);
    }
    if(r.baseHeaders && JSON.stringify(headers.slice(0,r.baseColumnCount))!==JSON.stringify(r.baseHeaders))throw new Error('準備中にフォームの質問列が変更されました。回答タブを確認してください。');
    if(sheet.getMaxColumns()<r.baseColumnCount+additional.length)sheet.insertColumnsAfter(sheet.getMaxColumns(),r.baseColumnCount+additional.length-sheet.getMaxColumns());
    const tail=sheet.getRange(1,r.baseColumnCount+1,1,additional.length).getDisplayValues()[0];
    if (tail.some(function(h,i){return h && h!==additional[i];})) throw new Error('追加列の配置が変わっています。');
    sheet.getRange(1,r.baseColumnCount+1,1,additional.length).setValues([additional.map(managedLiteral_)]);
    if(legacyTail)scoringManagementColumn_(sheet,true,r.input.statusHeader);
    custom.forEach(function(column,i) {
      if (!column.choices.length) return;
      const rule=SpreadsheetApp.newDataValidation().requireValueInList(column.choices,true).setAllowInvalid(false).build();
      sheet.getRange(2,r.baseColumnCount+2+i,Math.max(1,sheet.getMaxRows()-1),1).setDataValidation(rule);
    });
    if (r.templateItems && !r.formSync) {
      r.formSync = captureManagedFormSync_(r,formMetadata_(r.formId).items || [],managedSheetHeaders_(sheet));
      delete r.templateItems;saveManagedRecord_(r);
    }
    registerManagedSource_(r);
    ensureManagedResponseTrigger_();
    r.stage='registered';r.lastError='';r.formUrl=form.getPublishedUrl();r.formEditUrl=form.getEditUrl();
    return saveManagedRecord_(r);
  });
}

function reconcileManagedForm(id, expectedRevision) {
  return withAppLock_(function() {
    const r=loadManagedRecord_(id,expectedRevision);
    if (r.stage!=='creating') return r;
    const marker=r.id.replace(/'/g,"\\'");
    const files=Drive.Files.list({q:"trashed = false and appProperties has { key='turretRecord' and value='"+marker+"' }",fields:'files(id,appProperties),nextPageToken',pageSize:100});
    const matches=(files.files||[]).filter(function(f){return f.appProperties && f.appProperties.turretScript===ScriptApp.getScriptId();});
    if(matches.length!==1 || files.nextPageToken) throw new Error('作成済みフォームを一意に確認できません。自動で再作成しません。');
    FormApp.openById(matches[0].id).setPublished(false);
    r.formId=matches[0].id;r.stage='created';return saveManagedRecord_(r);
  });
}

function registerManagedSource_(r) {
  const config=managedSourceConfig_(r.input,getConfig_()), ss=getAppSpreadsheet_();
  if (!config.formSources.some(function(s){return s.id===ss.getId();})) config.formSources.push({id:ss.getId(),url:'https://docs.google.com/spreadsheets/d/'+ss.getId()+'/edit',courseId:''});
  if (!config.formSheetNamePrefix.includes(r.input.prefix)) config.formSheetNamePrefix.push(r.input.prefix);
  applyConfigDraft_(config,getConfigRevision_(getConfig_()));
  const mapping=ensureSheet_(MAPPING_SHEET_NAME,MAPPING_SHEET_HEADERS);
  if (!hasMatchingHeaders_(mapping,MAPPING_SHEET_HEADERS)) throw new Error('対応表の見出しを確認してください。');
  const rows=adminSheetValues_(mapping),found=[];
  rows.slice(1).forEach(function(row,i){if(row[0]===ss.getId()&&row[2]===r.sheetName)found.push(i+2);});
  if(found.length>1)throw new Error('対応表が重複しています。');
  const values=[ss.getId(),ss.getName(),r.sheetName,r.className,r.courseId,'turret作成'];
  mapping.getRange(found[0]||mapping.getLastRow()+1,1,1,6).setValues([values.map(managedLiteral_)]);
}

function managedLiteral_(value) { return /^[=+\-@]/.test(String(value)) ? "'"+value : value; }

/** 既存回答の識別に使うタブ名・列名は固定し、指定された設定だけを変更する。 */
function managedSettingsPlan_(r, changes) {
  if (!changes || typeof changes !== 'object' || Array.isArray(changes)) throw new Error('変更する設定がありません。');
  const keys=Object.keys(changes), allowed=['titlePattern','materialTitlePattern','description','domains','emails','labels'];
  if (!keys.length || keys.some(function(k){return !allowed.includes(k);})) throw new Error('変更できない設定が含まれています。');
  const input=JSON.parse(JSON.stringify(r.input)), result={input:input,label:r.label||r.className,title:r.title,materialTitle:r.materialTitle,changes:[]};
  function change(label,before,after){if(before!==after)result.changes.push(label+'：'+(before||'（空欄）')+' → '+(after||'（空欄）'));}
  keys.filter(function(k){return k!=='labels';}).forEach(function(k){
    if(typeof changes[k]!=='string' || changes[k].length>(k==='description'?5000:k==='domains'||k==='emails'?5000:150))throw new Error('入力を確認してください：'+k);
    if(['titlePattern','materialTitlePattern'].includes(k) && !/\{class(?:_label)?\}/.test(changes[k]))throw new Error('タイトルに {class_label} を含めてください。');
    if(!['domains','emails'].includes(k)){change({titlePattern:'フォーム名',materialTitlePattern:'資料タイトル',description:'資料の本文'}[k],input[k],changes[k].trim());input[k]=changes[k].trim();}
  });
  if(Object.prototype.hasOwnProperty.call(changes,'labels')){
    if(!changes.labels || typeof changes.labels!=='object' || Array.isArray(changes.labels) || Object.values(changes.labels).some(function(v){return typeof v!=='string'||!v.trim()||v.length>100;}))throw new Error('クラスの表記を確認してください。');
    if(Object.prototype.hasOwnProperty.call(changes.labels,r.courseId)){
      result.label=changes.labels[r.courseId].trim();change('クラスの表記',r.label||r.className,result.label);input.labels=Object.assign({},input.labels);input.labels[r.courseId]=result.label;
    }
  }
  function title(pattern){return String(pattern||'').replace(/\{class(?:_label)?\}/g,function(){return result.label;});}
  if(keys.includes('titlePattern') || result.label!==(r.label||r.className))result.title=title(input.titlePattern);
  if(keys.includes('materialTitlePattern') || result.label!==(r.label||r.className))result.materialTitle=title(input.materialTitlePattern);
  if(!result.title || result.title.length>300 || !result.materialTitle || result.materialTitle.length>300)throw new Error('展開後のタイトルを確認してください。');
  if(keys.includes('domains')||keys.includes('emails')){
    function list(k){return Object.prototype.hasOwnProperty.call(changes,k)?Array.from(new Set(changes[k].split(/[\s,、]+/).map(function(v){return v.toLowerCase();}).filter(Boolean))):input.policy[k].slice();}
    const policy={domains:list('domains'),emails:list('emails')};
    if(!policy.domains.length&&!policy.emails.length || policy.domains.some(function(d){return !/^[a-z0-9.-]+\.[a-z]{2,}$/.test(d)||['gmail.com','googlemail.com'].includes(d);}) || policy.emails.some(function(e){return !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e);}))throw new Error('回答者の指定が不正です。学校ドメインまたは個別メールを指定してください。');
    ['domains','emails'].forEach(function(k){change(k==='domains'?'回答許可ドメイン':'回答許可メール',input.policy[k].join('、'),policy[k].join('、'));});input.policy=policy;
  }
  result.formTitleChanged=result.title!==r.title;
  result.policyChanged=JSON.stringify(input.policy)!==JSON.stringify(r.input.policy);
  result.materialPatch={};
  if(result.materialTitle!==r.materialTitle)result.materialPatch.title=result.materialTitle;
  if(input.description!==r.input.description)result.materialPatch.description=input.description;
  return result;
}

function assertManagedSettingsEditable_(r) {
  assertManagedOwner_(r);
  if(r.kind!=='form'||!['registered','published','scheduled','draft','deleted'].includes(r.stage))throw new Error('フォームの準備と投稿結果の照合を完了してください。');
  if(r.formUpdate||r.columnSetup||r.settingsUpdate)throw new Error('更新が途中です。フォーム・列・設定の更新を再開して完了してください。');
}

function managedSettingsAccess_(r, policy, context) {
  const students=managedRoster_(r.courseId,context).map(function(s){return s.email;});
  if(!students.length||students.some(function(email){return !policy.emails.includes(email)&&!policy.domains.includes(email.split('@')[1]);}))throw new Error('対象クラスの生徒全員を回答許可の範囲に含めてください。');
  return students.sort();
}

function managedSettingsSnapshot_(r, plan, context) {
  const snapshot={};
  if(plan.formTitleChanged){snapshot.formTitle=FormApp.openById(r.formId).getTitle();snapshot.fileName=DriveApp.getFileById(r.formId).getName();}
  if(plan.policyChanged){
    const form=FormApp.openById(r.formId);
    if(!form.collectsEmail())throw new Error('フォームのメール収集を有効にしてください。');
    snapshot.students=managedSettingsAccess_(r,plan.input.policy,context);snapshot.accepting=form.isAcceptingResponses();
    snapshot.permissions=publishedPermissions_(r.formId).map(managedCanonicalSnapshot_).sort(function(a,b){return JSON.stringify(a).localeCompare(JSON.stringify(b));});
    if(snapshot.permissions.some(function(p){return p.role!=='reader';}))throw new Error('回答者権限を安全に変更できません。フォーム画面で権限を確認してください。');
  }
  if(r.materialId&&r.stage!=='deleted'&&Object.keys(plan.materialPatch).length){
    const post=Classroom.Courses.CourseWorkMaterials.get(r.courseId,r.materialId);
    if(!post||String(post.id)!==String(r.materialId)||!['DRAFT','PUBLISHED'].includes(post.state))throw new Error('資料の投稿状態を確認してください。');
    snapshot.material={};Object.keys(plan.materialPatch).forEach(function(k){snapshot.material[k]=post[k]||'';});
  }
  return snapshot;
}

function previewManagedFormSettings(raw) {
  return withAppLock_(function(){
    if(!raw||!Array.isArray(raw.targets)||!raw.targets.length||raw.targets.length>100||new Set(raw.targets.map(function(t){return t&&t.id;})).size!==raw.targets.length)throw new Error('対象フォームを1〜100件選択してください。');
    const records=getManagedRecords_(),context={};
    return {targets:raw.targets.map(function(t){
      const r=records.find(function(r){return r.id===t.id;});if(!r)throw new Error('管理対象の記録がありません。');assertManagedSettingsEditable_(r);
      if(typeof t.revision!=='string'||r.revision!==t.revision)throw new Error('記録が更新されました。編集を開き直してください。');
      const plan=managedSettingsPlan_(r,raw.changes),snapshot=managedSettingsSnapshot_(r,plan,context);
      const savingMaterial=raw.saveMaterialSettings===true;
      const registerCommon=raw.registerCommonMaterial===true;
      if(registerCommon&&!savingMaterial)throw new Error('共通文面は投稿設定の保存と同時に登録してください。');
      if(savingMaterial)assertManagedMaterialSettingsPatch_(raw.changes);
      return {id:r.id,label:r.label||r.className,changes:plan.changes,changed:savingMaterial||plan.changes.length>0,fingerprint:adminDigest_(managedCanonicalSnapshot_([r,raw.changes,snapshot].concat(savingMaterial?[true]:[]).concat(registerCommon?[true]:[])))};
    })};
  });
}

function assertManagedMaterialSettingsPatch_(changes) {
  if(Object.keys(changes).some(function(k){return !['materialTitlePattern','description'].includes(k);}))throw new Error('投稿設定では投稿タイトル・投稿文だけを保存してください。');
}

function runManagedFormSettings(id, changes, fingerprint, saveMaterialSettings, registerCommonMaterial) {
  return withAppLock_(function(){
    const r=loadManagedRecord_(id);assertManagedSettingsEditable_(r);
    const plan=managedSettingsPlan_(r,changes),snapshot=managedSettingsSnapshot_(r,plan,{});
    const savingMaterial=saveMaterialSettings===true;
    const registerCommon=registerCommonMaterial===true;
    if(registerCommon&&!savingMaterial)throw new Error('共通文面は投稿設定の保存と同時に登録してください。');
    if(savingMaterial)assertManagedMaterialSettingsPatch_(changes);
    if(adminDigest_(managedCanonicalSnapshot_([r,changes,snapshot].concat(savingMaterial?[true]:[]).concat(registerCommon?[true]:[])))!==fingerprint)throw new Error('確認後に設定・フォーム・資料が更新されました。再確認してください。');
    if(!plan.changes.length&&!savingMaterial)return r;
    if(savingMaterial)plan.materialSettingsSaved=true;
    // 共通文面も更新計画に保持し、途中からの再開と画面の再表示で基準を失わない。
    if(registerCommon)plan.materialSettingsBaseline={materialTitlePattern:plan.input.materialTitlePattern,description:plan.input.description||''};
    if(snapshot.material)ScriptApp.requireScopes(ScriptApp.AuthMode.FULL,['https://www.googleapis.com/auth/classroom.courseworkmaterials']);
    r.settingsUpdate={plan:plan,snapshot:snapshot,accepting:snapshot.accepting};saveManagedRecord_(r);
    return resumeManagedFormSettingsUnlocked_(r);
  });
}

function resumeManagedFormSettings(id, revision) {
  return withAppLock_(function(){return resumeManagedFormSettingsUnlocked_(loadManagedRecord_(id,revision));});
}

/** 書き込み結果が不明でも同じ値の再適用だけで復旧し、手動変更は上書きしない。 */
function resumeManagedFormSettingsUnlocked_(r) {
  if(!r.settingsUpdate)return r;
  const update=r.settingsUpdate,plan=update.plan,before=update.snapshot;
  function unchanged(current,old,desired){if(current!==old&&current!==desired)throw new Error('更新途中に手動変更されています。元の値または確認した値へ戻してから設定の更新を再開してください。');}
  const form=(plan.formTitleChanged||plan.policyChanged)?FormApp.openById(r.formId):null;
  if(plan.formTitleChanged){
    const file=DriveApp.getFileById(r.formId);unchanged(form.getTitle(),before.formTitle,plan.title);unchanged(file.getName(),before.fileName,plan.title);
    form.setTitle(plan.title);file.setName(plan.title);
  }
  if(plan.policyChanged){
    managedSettingsAccess_(r,plan.input.policy,{});
    if(!form.collectsEmail())throw new Error('フォームのメール収集を有効にしてください。');
    const permissions=publishedPermissions_(r.formId),policy=plan.input.policy;
    function identity(p){return [p.type,p.role,p.domain||p.emailAddress||'',p.view].join('|');}
    const original=before.permissions.map(identity),desired=policy.domains.map(function(d){return 'domain|reader|'+d+'|published';}).concat(policy.emails.map(function(e){return 'user|reader|'+e+'|published';}));
    if(permissions.some(function(p){return !original.includes(identity(p))&&!desired.includes(identity(p));}))throw new Error('更新途中に回答者権限が変更されています。フォーム画面で確認してください。');
    form.setAcceptingResponses(false);
    ensureCopiedResponderPolicy_(Object.assign({},r,{input:plan.input}));
    validateResponderAccess_(publishedPermissions_(r.formId),policy,managedSettingsAccess_(r,policy,{}),true);
  }
  if(before.material){
    const service=Classroom.Courses.CourseWorkMaterials,post=service.get(r.courseId,r.materialId);
    if(!post||String(post.id)!==String(r.materialId)||!['DRAFT','PUBLISHED'].includes(post.state))throw new Error('資料の投稿状態を確認してください。');
    Object.keys(plan.materialPatch).forEach(function(k){unchanged(post[k]||'',before.material[k],plan.materialPatch[k]);});
    service.patch(plan.materialPatch,r.courseId,r.materialId,{updateMask:Object.keys(plan.materialPatch).join(',')});
  }
  r.input=plan.input;r.label=plan.label;r.title=plan.title;r.materialTitle=plan.materialTitle;
  if(plan.materialSettingsSaved===true)r.materialSettingsSaved=true;
  if(plan.materialSettingsBaseline)r.materialSettingsBaseline=plan.materialSettingsBaseline;
  saveManagedRecord_(r);
  if(plan.policyChanged){form.setAcceptingResponses(update.accepting===true&&!managedCloseIsDue_(r));r.accepting=form.isAcceptingResponses();}
  delete r.settingsUpdate;r.lastError='';return saveManagedRecord_(r);
}

/** フォーム作成とは別に、準備済みの回答タブへ評価・コメント列を追加する。 */
function previewResponseColumns(raw, expectedRevision) {
  return withAppLock_(function() {return previewResponseColumnsUnlocked_(raw,expectedRevision);});
}

function responseColumnsContext_(record) {
  if (record.settingsUpdate) throw new Error('設定の更新を再開して完了してください。');
  if (record.formUpdate) throw new Error('フォーム更新を完了してから列を設定してください。');
  return managedUpdateContext_(record);
}

function validateResponseColumns_(record, columns, headers) {
  managedUniqueHeaders_(headers);
  const reserved=[record.input.emailHeader,record.input.nameHeader,record.input.statusHeader,SCORING_MANAGEMENT_HEADER,'タイムスタンプ','Timestamp'];
  if (reserved.slice(0,3).some(function(h){return !headers.includes(h);})) throw new Error('名前・送信状態などの必要な列を確認してください。');
  const names=columns.map(function(c){return c.header;});
  if (new Set(names).size!==names.length) throw new Error('追加する列名が重複しています。');
  if (names.some(function(h){return reserved.includes(h) || (record.baseHeaders || []).includes(h);})) throw new Error('システム列やフォームの質問列は追加列に指定できません。');
}

function previewResponseColumnsUnlocked_(raw, expectedRevision) {
  const config=getConfig_();
  if (expectedRevision!==getConfigRevision_(config)) throw new Error('設定が更新されました。状態を更新してください。');
  if (!raw || !Array.isArray(raw.recordIds) || !raw.recordIds.length || raw.recordIds.length>100 || raw.recordIds.some(function(id){return typeof id!=='string' || !id;}) || new Set(raw.recordIds).size!==raw.recordIds.length) throw new Error('対象の回答タブを1〜100件選択してください。');
  const columns=normalizeManagedColumns_(raw.additionalColumns);
  if (raw.columnNamesOnly !== undefined && typeof raw.columnNamesOnly !== 'boolean') throw new Error('列名のみの設定は真偽値で指定してください。');
  if (raw.columnNamesOnly && columns.some(function(c){return c.choices.length;})) throw new Error('列名のみの設定では入力用の選択肢は指定できません。');
  if (!columns.length) throw new Error('追加する列を入力してください。');
  const targets=raw.recordIds.map(function(id) {
    const r=loadManagedRecord_(id);
    if (r.columnSetup) throw new Error('列の設定が途中です。「列の設定を再開」を実行してください。');
    const sheet=responseColumnsContext_(r).sheet,headers=managedSheetHeaders_(sheet);
    validateResponseColumns_(r,columns,headers);
    managedSourceConfig_(Object.assign({},r.input,{additionalColumns:columns}),config);
    return {id:r.id,revision:r.revision,label:r.label,sheetName:sheet.getName(),sheetId:sheet.getSheetId(),headers:headers,added:columns.filter(function(c){return !headers.includes(c.header);}).map(function(c){return c.header;})};
  });
  const preview={additionalColumns:columns,targets:targets,configRevision:expectedRevision};
  if (raw.columnNamesOnly) preview.columnNamesOnly=true;
  preview.fingerprint=adminDigest_(preview);
  return preview;
}

function applyResponseColumns(raw, expectedRevision, fingerprint) {
  return withAppLock_(function() {
    const p=previewResponseColumnsUnlocked_(raw,expectedRevision);
    if (p.fingerprint!==fingerprint) throw new Error('確認後に回答列や設定が変更されました。再確認してください。');
    let completed=0;const errors=[];
    p.targets.forEach(function(target) {
      try {
        const r=loadManagedRecord_(target.id,target.revision);
        // 書込み前に入力と元の見出しを保存し、部分成功でも同じ列設定を再開する。
        r.columnSetup={columns:p.additionalColumns,headers:target.headers};
        if (p.columnNamesOnly) r.columnSetup.columnNamesOnly=true;
        saveManagedRecord_(r);
        finishResponseColumns_(r);completed++;
      } catch(error) {errors.push({id:target.id,label:target.label,message:String(error.message || error)});}
    });
    return {completed:completed,errors:errors};
  });
}

function resumeResponseColumns(id, expectedRevision) {
  return withAppLock_(function() {
    const r=loadManagedRecord_(id,expectedRevision);
    if (!r.columnSetup) return r;
    return finishResponseColumns_(r);
  });
}

function finishResponseColumns_(r) {
  const sheet=responseColumnsContext_(r).sheet,columns=normalizeManagedColumns_(r.columnSetup.columns);
  const namesOnly=r.columnSetup.columnNamesOnly === true;
  let headers=managedSheetHeaders_(sheet);
  validateResponseColumns_(r,columns,headers);
  const original=r.columnSetup.headers;
  if (original.some(function(h){return !headers.includes(h);}) || headers.some(function(h){return h!==SCORING_MANAGEMENT_HEADER && !original.includes(h) && !columns.some(function(c){return c.header===h;});})) throw new Error('列の設定中に回答列が変更されました。見出しを確認してください。');
  const existing=managedCustomColumns_(r.input).filter(function(c){return !!c.header;});
  const merged=existing.map(function(c){return namesOnly ? c : columns.find(function(n){return n.header===c.header;}) || c;});
  columns.forEach(function(c){if(!merged.some(function(old){return old.header===c.header;}))merged.push(c);});
  const input=Object.assign({},r.input,{additionalColumns:merged});
  const config=managedSourceConfig_(input,getConfig_());
  const added=columns.filter(function(c){return !headers.includes(c.header);});
  if (added.length) {
    const width=headers.length;
    if (sheet.getMaxColumns()<width+added.length) sheet.insertColumnsAfter(sheet.getMaxColumns(),width+added.length-sheet.getMaxColumns());
    sheet.getRange(1,width+1,1,added.length).setValues([added.map(function(c){return managedLiteral_(c.header);})]);
    // 書込みを確定し、追加列も含めた列数で移動先を計算する。
    SpreadsheetApp.flush();
  }
  // 全列移動でセルの値・数式・メモを保持し、管理・送信状態を右端へ戻す。
  scoringManagementColumn_(sheet,true,r.input.statusHeader);
  headers=managedSheetHeaders_(sheet);
  columns.forEach(function(column) {
    // 列名だけの設定では既存の入力規則を保持する。新規列は自由入力にする。
    if (namesOnly && original.includes(column.header)) return;
    const range=sheet.getRange(2,headers.indexOf(column.header)+1,Math.max(1,sheet.getMaxRows()-1),1);
    if (!column.choices.length) {range.clearDataValidations();return;}
    const rule=SpreadsheetApp.newDataValidation().requireValueInList(column.choices,true).setAllowInvalid(false).build();
    range.setDataValidation(rule);
  });
  SpreadsheetApp.flush();
  const savedHeaders=managedSheetHeaders_(sheet);
  if (columns.some(function(c){return !savedHeaders.includes(c.header);})) throw new Error('追加列の保存結果を確認できません。列の設定を再開してください。');
  applyConfigDraft_(config,getConfigRevision_(getConfig_()));
  r.input=input;delete r.columnSetup;return saveManagedRecord_(r);
}

function fillManagedNames_(record, sheet, firstRow, numberRows, context) {
  context=context||{};
  if(!context.nameMaps)context.nameMaps=new Map();
  if(!context.nameMaps.has(record.courseId)) {
    const names=new Map();
    managedRoster_(record.courseId,context).forEach(function(s){if(names.has(s.email)&&names.get(s.email)!==s.name)names.set(s.email,'');else if(!names.has(s.email))names.set(s.email,s.name);});
    context.nameMaps.set(record.courseId,names);
  }
  const map=context.nameMaps.get(record.courseId);
  const headers=createHeaderMap_(sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0]);
  if (!(record.input.emailHeader in headers) || !(record.input.nameHeader in headers)) throw new Error('名前補完用の列がありません。');
  const start=firstRow||2, length=numberRows||Math.max(0,sheet.getLastRow()-start+1);
  const result={updated:0,unmatched:0,skipped:0};if(!length)return result;
  const nameCol=headers[record.input.nameHeader]+1, nameRange=sheet.getRange(start,nameCol,length,1);
  const names=nameRange.getDisplayValues(), formulas=nameRange.getFormulas();
  const emails=sheet.getRange(start,headers[record.input.emailHeader]+1,length,1).getDisplayValues(), blocks=[];
  emails.forEach(function(row,i){
    if(String(names[i][0]||'').trim() || formulas[i][0]){result.skipped++;return;}
    const email=String(row[0]||'').trim().toLowerCase(),name=map.get(email);
    if(!email){result.skipped++;return;}
    if(!name){result.unmatched++;return;}
    blocks.push({row:start+i,col:nameCol,values:[managedLiteral_(name)]});result.updated++;
  });
  scoringWriteBlocks_(sheet,blocks);
  return result;
}

function refreshManagedNames(id, revision) {
  return withAppLock_(function(){const r=loadManagedRecord_(id,revision);if(r.formUpdate)throw new Error('フォームの更新を完了してから名前を補完してください。');return fillManagedNames_(r,managedAnswerSheet_(r));});
}

/** 通常の取り込み前に、取りこぼした送信イベントや処理中だったイベントを補完する。 */
function refreshAllManagedNames_() {
  const context={}; // この実行中だけ使う。次回の取り込みでは名簿の変更を反映する。
  getManagedRecords_().filter(function(r){return r.kind==='form'&&r.responseSheetId!=null;}).forEach(function(r){
    assertManagedOwner_(r);fillManagedNames_(r,managedAnswerSheet_(r,context),undefined,undefined,context);
  });
}

function managedPostService_(r) { return r.kind==='announcement' ? Classroom.Courses.Announcements : Classroom.Courses.CourseWorkMaterials; }
function managedPostId_(r) { return r.kind==='announcement' ? r.postId : r.materialId; }

/** フォーム更新では項目と質問の ID を維持し、回答行は書き換えない。 */
function managedItemQuestions_(item) {
  if (item.native) {
    if (['SECTION_HEADER','PAGE_BREAK','IMAGE','VIDEO'].includes(item.native.type)) return [];
    if (['GRID','CHECKBOX_GRID'].includes(item.native.type)) return item.native.rows.map(function(row) {return {questionId:item.itemId + ':' + adminDigest_(row).slice(0,16),rowQuestion:{title:row}};});
    return [{questionId:item.itemId}];
  }
  return item.questionItem ? [item.questionItem.question] : item.questionGroupItem ? item.questionGroupItem.questions || [] : [];
}

function captureManagedFormSync_(record, items, headers) {
  const source = record.templateItems, itemMap = {}, questionMap = {}, columns = {};
  if (source.length !== items.length) throw new Error('複製中にひな形が変更されました。質問の対応を確認してください。');
  source.forEach(function(item,index) {
    const target = items[index];
    if (item.title !== target.title || managedItemKind_(item) !== managedItemKind_(target)) throw new Error('複製した質問の対応を確認できません。');
    itemMap[item.itemId] = target.itemId;
    const questions = managedItemQuestions_(item), copies = managedItemQuestions_(target);
    if (questions.length !== copies.length) throw new Error('複製した質問の対応を確認できません。');
    questions.forEach(function(q,i) {questionMap[q.questionId] = copies[i].questionId;});
  });
  managedQuestionColumns_(items).forEach(function(q) {
    if (!headers.includes(q.header)) throw new Error('回答列の反映待ちです。「準備を続ける」で再確認してください: ' + q.header);
    columns[q.id] = q.header;
  });
  const first = Math.min.apply(null,Object.values(columns).map(function(h) {return headers.indexOf(h);}).concat([record.baseColumnCount || headers.length]));
  return {engine:items.some(function(i) {return i.native;})?'formapp-v1':undefined,itemMap:itemMap,questionMap:questionMap,columns:columns,archivedHeaders:[],fixedHeaders:headers.slice(0,first)};
}

function managedQuestionColumns_(items) {
  const columns = [];
  (items || []).forEach(function(item) {
    managedItemQuestions_(item).forEach(function(q) {
      columns.push({ id: q.questionId, header: q.rowQuestion ? item.title + ' [' + q.rowQuestion.title + ']' : item.title });
    });
  });
  return columns;
}

function managedItemKind_(item) {
  if (item.native) return item.native.type;
  const kind = ['questionItem','questionGroupItem','pageBreakItem','textItem','imageItem','videoItem'].find(function(k) {return item[k] != null;});
  if (!kind) throw new Error('未対応のフォーム項目です。');
  const q = item.questionItem && item.questionItem.question;
  return kind + (q ? ':' + Object.keys(q).filter(function(k) {return /Question$/.test(k);}).sort().join(',') : '');
}

/** 比較用の表現では、オブジェクトのキー順と出力専用のメディア URL を無視する。 */
function managedFormShape_(value) {
  if (Array.isArray(value)) return value.map(managedFormShape_);
  if (!value || typeof value !== 'object') return value;
  const result = {};
  Object.keys(value).sort().forEach(function(key) {
    if (['contentUri','sourceUri'].includes(key) || value[key] === '' || value[key] === false || value[key] == null) return;
    result[key] = managedFormShape_(value[key]);
  });
  return result;
}

function managedFormContentDigest_(metadata) {
  return adminDigest_(managedFormShape_({items:metadata.items || [],description:(metadata.info || {}).description || ''}));
}

function managedWritableItem_(value) {
  if (Array.isArray(value)) return value.map(managedWritableItem_);
  if (!value || typeof value !== 'object') return value;
  const result = {};
  Object.keys(value).forEach(function(key) {
    if (key !== 'contentUri' && key !== '_turretImageHash') result[key] = managedWritableItem_(value[key]);
  });
  if (value.contentUri && !result.sourceUri) result.sourceUri = value.contentUri;
  return result;
}

/** 画像のダウンロード URL は期限切れや読み取り時の変化があるため、URL ではなく内容を比較する。 */
function managedUpdateMetadata_(id) {
  return formMetadata_(id);
}

function managedItemChanged_(before, after) {
  if (JSON.stringify(managedFormShape_(before)) !== JSON.stringify(managedFormShape_(after))) return true;
  // 移行処理など、元のメタデータを渡す呼び出し元は、まだメディアをダウンロードしていない。
  return JSON.stringify(before).includes('contentUri') && !JSON.stringify(before).includes('_turretImageHash') && JSON.stringify(before) !== JSON.stringify(after);
}

function managedUniqueHeaders_(headers) {
  if (headers.some(function(h) {return !String(h || '').trim();}) || new Set(headers).size !== headers.length) {
    throw new Error('回答列の見出しが空欄または重複しています。列名を確認してください。');
  }
}

function buildManagedFormUpdate_(record, template, target, headers) {
  managedUniqueHeaders_(headers);
  const oldItems = target.items || [], sourceItems = template.items || [], previous = record.formSync || {};
  const itemMap = {}, questionMap = {}, used = new Set(), requests = [], changes = [];
  const oldColumns = managedQuestionColumns_(oldItems);
  const oldColumnMap = {};
  oldColumns.forEach(function(q) {
    const header = previous.columns && previous.columns[q.id] || q.header;
    if (!headers.includes(header)) throw new Error('既存の回答列を確認できません: ' + header);
    oldColumnMap[q.id] = header;
  });
  const management = [record.input.nameHeader].concat(managedCustomColumns_(record.input).map(function(c) {return c.header;}), headers.includes(SCORING_MANAGEMENT_HEADER)?[SCORING_MANAGEMENT_HEADER,record.input.statusHeader]:[record.input.statusHeader]);
  management.concat(record.input.emailHeader).forEach(function(h) {
    if (!headers.includes(h)) throw new Error('管理用の回答列がありません: ' + h);
  });
  const items = sourceItems.map(function(source) {
    if (!source.itemId) throw new Error('ひな形の質問IDを取得できません。');
    if (source.native && (!NATIVE_FORM_TYPES_[source.native.type] || source.native.type === 'VIDEO') || managedItemQuestions_(source).some(function(q) {return q.fileUploadQuestion;})) throw new Error('未対応の質問種類があります。動画・ファイルアップロードなどを含むフォームはフォーム画面で編集してください。');
    const mappedId = record.formSync ? previous.itemMap && previous.itemMap[source.itemId] : source.itemId;
    let old = oldItems.find(function(i) {return i.itemId === mappedId;});
    // 旧コピーには ID の記録がない。タイトルと種類から一意に特定できる場合だけ引き継ぐ。
    if (!old && !record.formSync) {
      const matches = oldItems.filter(function(i) {return !used.has(i.itemId) && i.title === source.title && managedItemKind_(i) === managedItemKind_(source);});
      if (matches.length > 1) throw new Error('ひな形と配付先の質問の対応が重複しています。');
      old = matches[0];
    }
    if (old && used.has(old.itemId)) throw new Error('質問の対応が重複しています。');
    if (old && managedItemKind_(old) !== managedItemKind_(source)) throw new Error('質問の種類を変更する場合は、ひな形で別の質問として追加してください: ' + source.title);
    if (source.native && source.native.type === 'IMAGE' && (!old || old.native.imageHash !== source.native.imageHash)) throw new Error('画像の追加・差し替えは自動更新の対象外です。各フォームの画面で設定してください。初回複製では画像も引き継ぎます。');
    const item = JSON.parse(JSON.stringify(source));
    // 作成時に指定した ID により、一括処理の応答を失っても安全に照合できる。
    item.itemId = old ? old.itemId : source.native ? 'new:' + source.itemId : source.itemId;
    if (!old && oldItems.some(function(i) {return i.itemId === item.itemId;})) throw new Error('追加する質問のIDが配付先と衝突しています。');
    itemMap[source.itemId] = item.itemId;
    if (old) used.add(old.itemId);
    const oldQuestions = old ? managedItemQuestions_(old) : [];
    managedItemQuestions_(source.native ? source : item).forEach(function(q, index) {
      const sourceId = q.questionId;
      if (source.native) {questionMap[sourceId] = managedItemQuestions_(item)[index].questionId;return;}
      const mappedQuestionId = record.formSync ? previous.questionMap && previous.questionMap[sourceId] : sourceId;
      let prior = oldQuestions.find(function(p) {return p.questionId === mappedQuestionId;});
      if (!prior && !record.formSync && oldQuestions.length === 1) prior = oldQuestions[0];
      if (!prior && !record.formSync && q.rowQuestion) prior = oldQuestions.find(function(p) {return p.rowQuestion && p.rowQuestion.title === q.rowQuestion.title;});
      q.questionId = prior ? prior.questionId : sourceId;
      if (!prior && oldColumns.some(function(c) {return c.id === q.questionId;})) throw new Error('追加する質問のIDが配付先と衝突しています。');
      questionMap[sourceId] = q.questionId;
    });
    if (!old) changes.push('追加: ' + (item.title || '説明・セクション'));
    else if (managedItemChanged_(old,item)) changes.push('変更: ' + (item.title || '説明・セクション'));
    return item;
  });
  // 分岐先にはコピー先ではなく、元フォームの項目 ID が入っている。
  items.forEach(function(item) {
    if (item.native) {
      (item.native.choices || []).concat(item.native.navigation ? [item.native.navigation] : []).forEach(function(option) {
        if (option.goToSectionId) {if (!itemMap[option.goToSectionId]) throw new Error('分岐先のセクションがありません。');option.goToSectionId = itemMap[option.goToSectionId];}
      });
    }
    managedItemQuestions_(item).forEach(function(q) {
      ((q.choiceQuestion || {}).options || []).forEach(function(option) {
        if (option.goToSectionId) {
          if (!itemMap[option.goToSectionId]) throw new Error('分岐先のセクションがありません。');
          option.goToSectionId = itemMap[option.goToSectionId];
        }
      });
    });
  });
  const desiredColumns = managedQuestionColumns_(items), columns = {}, restoreHeaders = [];
  if (new Set(items.map(function(i) {return i.itemId;})).size !== items.length || new Set(desiredColumns.map(function(q) {return q.id;})).size !== desiredColumns.length) throw new Error('質問IDが重複しています。');
  managedUniqueHeaders_(desiredColumns.map(function(q) {return q.header;}));
  desiredColumns.forEach(function(q) {
    const stable = oldColumnMap[q.id] || q.header;
    if (q.header===SCORING_MANAGEMENT_HEADER || management.includes(q.header) || q.header === record.input.emailHeader || (headers.includes(q.header) && q.header !== oldColumnMap[q.id])) {
      throw new Error('質問名が既存の回答列・管理列と衝突しています: ' + q.header);
    }
    columns[q.id] = stable;
    if (q.header !== stable) restoreHeaders.push({from:q.header,to:stable});
  });
  const archivedHeaders = Array.from(new Set((previous.archivedHeaders || []).concat(oldColumns.filter(function(q) {return !columns[q.id];}).map(function(q) {return oldColumnMap[q.id];}))));
  const fixedHeaders = previous.fixedHeaders || headers.filter(function(h) {return !Object.values(oldColumnMap).includes(h) && !management.includes(h) && !archivedHeaders.includes(h);});
  // 既知のシステム列を先頭に置き、無関係な手動追加列は管理列の後に残す。
  const systemHeaders = fixedHeaders.filter(function(h) {return headers.indexOf(h) < Math.min.apply(null, oldColumns.map(function(q) {return headers.indexOf(oldColumnMap[q.id]);}).concat([headers.length]));});
  const otherHeaders = headers.filter(function(h) {return !systemHeaders.includes(h) && !Object.values(oldColumnMap).includes(h) && !management.includes(h) && !archivedHeaders.includes(h);});
  const order = systemHeaders.concat(desiredColumns.map(function(q) {return columns[q.id];}),management,otherHeaders);
  managedUniqueHeaders_(order);
  const working = oldItems.map(function(i) {return i.itemId;});
  for (let i = working.length - 1; i >= 0; i--) {
    if (!items.some(function(item) {return item.itemId === working[i];})) {
      requests.push({deleteItem:{location:{index:i}}});changes.push('削除: ' + (oldItems[i].title || '説明・セクション'));working.splice(i,1);
    }
  }
  items.forEach(function(item,index) {
    const at = working.indexOf(item.itemId);
    if (at < 0) {requests.push({createItem:{item:managedWritableItem_(item),location:{index:index}}});working.splice(index,0,item.itemId);}
    else {
      if (at !== index) {requests.push({moveItem:{originalLocation:{index:at},newLocation:{index:index}}});working.splice(at,1);working.splice(index,0,item.itemId);}
      const old = oldItems.find(function(i) {return i.itemId === item.itemId;});
      if (managedItemChanged_(old,item)) requests.push({updateItem:{item:managedWritableItem_(item),location:{index:index},updateMask:'*'}});
    }
  });
  const description = (template.info || {}).description || '';
  if (description !== ((target.info || {}).description || '')) {requests.push({updateFormInfo:{info:{description:description},updateMask:'description'}});changes.push('フォームの説明文を変更');}
  return {items:items,description:description,requests:requests,changes:changes,order:order,restoreHeaders:restoreHeaders,statusHeader:record.input.statusHeader,
    sync:{engine:template.engine,itemMap:itemMap,questionMap:questionMap,columns:columns,archivedHeaders:archivedHeaders,fixedHeaders:systemHeaders},
    archivedHeaders:archivedHeaders};
}

function managedSheetHeaders_(sheet) {
  return sheet.getRange(1,1,1,Math.max(1,sheet.getLastColumn())).getDisplayValues()[0];
}

function managedColumnHasData_(sheet, column) {
  const count = sheet.getLastRow() - 1;
  if (count > 0) {
    const range = sheet.getRange(2,column,count,1);
    if (range.getValues().some(function(row) {return row[0] !== '' && row[0] != null;}) || range.getFormulas().some(function(row) {return !!row[0];})) return true;
  }
  return sheet.getMaxRows() > 1 && sheet.getRange(2,column,sheet.getMaxRows()-1,1).getNotes().some(function(row) {return !!row[0];});
}

function arrangeManagedFormColumns_(sheet, plan) {
  let headers = managedSheetHeaders_(sheet);
  managedUniqueHeaders_(headers);
  // 最初の変更前に列配置全体を検証する。Sheets 側の作成は Forms 側より遅れる場合がある。
  const projected = headers.map(function(h) {const rename = plan.restoreHeaders.find(function(r) {return r.from === h;});return rename ? rename.to : h;});
  managedUniqueHeaders_(projected);
  if (plan.order.some(function(h) {return !projected.includes(h);})) throw new Error('回答列の反映待ちです。「更新を再開」で再確認してください。');
  if (projected.some(function(h) {return !plan.order.includes(h) && !plan.archivedHeaders.includes(h);})) throw new Error('更新中に回答列が変更されました。列名を確認してください。');
  plan.restoreHeaders.forEach(function(r) {
    const index = headers.indexOf(r.from);
    if (index >= 0) {sheet.getRange(1,index+1).setValue(managedLiteral_(r.to));headers[index]=r.to;}
  });
  const deleted = [], retained = [];
  plan.archivedHeaders.forEach(function(header) {
    headers = managedSheetHeaders_(sheet);
    const index = headers.indexOf(header);
    if (index < 0) return;
    // 0、false、空白文字、結果が空の数式もデータとして扱う。
    if (managedColumnHasData_(sheet,index+1)) retained.push(header);
    else {sheet.deleteColumns(index+1,1);deleted.push(header);}
  });
  let order = plan.order.concat(retained);
  if(order.includes(SCORING_MANAGEMENT_HEADER))order=order.filter(function(h){return h!==SCORING_MANAGEMENT_HEADER && h!==plan.statusHeader;}).concat([SCORING_MANAGEMENT_HEADER,plan.statusHeader]);
  order.forEach(function(header,index) {
    headers = managedSheetHeaders_(sheet);
    const current = headers.indexOf(header);
    if (current < 0) throw new Error('並べ替え対象の回答列がありません: ' + header);
    if (current !== index) sheet.moveColumns(sheet.getRange(1,current+1, sheet.getMaxRows(),1),index+1);
  });
  SpreadsheetApp.flush();
  if (JSON.stringify(managedSheetHeaders_(sheet)) !== JSON.stringify(order)) throw new Error('回答列の並べ替え結果を確認できません。「更新を再開」で再確認してください。');
  return {deleted:deleted,retained:retained};
}

function assertNoManagedFormUpdate_() {
  const records=getManagedRecords_();
  if (records.some(function(r) {return r.kind==='form' && r.columnSetup;})) throw new Error('回答列の設定が途中です。「評価・コメント列」の「列の設定を再開」を完了してから処理してください。');
  if (records.some(function(r) {return r.kind === 'form' && r.formUpdate;})) throw new Error('フォームの更新が途中です。「資料を投稿」の「更新を再開」を完了してから処理してください。');
}

function managedUpdateContext_(record) {
  if (record.kind !== 'form' || !['registered','published','scheduled','draft','deleted'].includes(record.stage)) throw new Error('フォームの準備・投稿結果の確認を完了してください。');
  const form = FormApp.openById(record.formId), sheet = managedAnswerSheet_(record);
  if (!sheet || form.getDestinationId() !== getAppSpreadsheet_().getId() || !sheet.getFormUrl() || FormApp.openByUrl(sheet.getFormUrl()).getId() !== record.formId) throw new Error('フォームと回答先の対応が変更されています。');
  return {form:form,sheet:sheet};
}

function previewManagedFormUpdate(id) {
  return withAppLock_(function() {return previewManagedFormUpdateUnlocked_(id).preview;});
}

function previewManagedFormUpdateUnlocked_(id) {
  const r = loadManagedRecord_(id);
  if (r.settingsUpdate) throw new Error('設定の更新を再開して完了してください。');
  if (r.columnSetup) throw new Error('回答列の設定を再開し、完了してからフォームを更新してください。');
  if (r.formUpdate) throw new Error('フォームの更新が途中です。「更新を再開」を実行してください。');
  const context = managedUpdateContext_(r), target = managedUpdateMetadata_(r.formId), template = managedUpdateMetadata_(r.input.templateId);
  if (!target.revisionId || !template.revisionId) throw new Error('フォームの更新版を確認できません。');
  if (!target.settings || !target.settings.collectsEmail) throw new Error('配付済みフォームのメール収集を有効にしてください。');
  if (!!(target.settings.quizSettings || {}).isQuiz !== !!((template.settings || {}).quizSettings || {}).isQuiz) throw new Error('ひな形と配付先のテスト設定が異なります。採点履歴を保護するため更新を停止しました。');
  if ((target.settings.quizSettings || {}).isQuiz) throw new Error('テスト形式のフォームは正解・採点設定を完全に取得できないため、自動更新の対象外です。複製・配付は利用できます。');
  if (target.items.some(function(i) {return i.native && (!NATIVE_FORM_TYPES_[i.native.type] || i.native.type === 'VIDEO');})) throw new Error('配付先に未対応の質問種類があります。動画・ファイルアップロードなどを含むフォームはフォーム画面で編集してください。');
  const headers = managedSheetHeaders_(context.sheet), plan = buildManagedFormUpdate_(r,template,target,headers);
  const archived = plan.archivedHeaders.filter(function(h) {return headers.includes(h);}).map(function(h) {return {header:h,hasData:managedColumnHasData_(context.sheet,headers.indexOf(h)+1)};});
  const accepting = context.form.isAcceptingResponses();
  const fingerprint = adminDigest_([r,managedFormShape_(target),managedFormShape_(template),headers,archived,accepting]);
  const preview = {id:id,label:r.label || r.className,title:r.title,changes:plan.changes,
    order:plan.order,retained:archived.filter(function(c) {return c.hasData;}).map(function(c) {return c.header;}),
    removed:archived.filter(function(c) {return !c.hasData;}).map(function(c) {return c.header;}),
    preservedNames:plan.restoreHeaders, warnings:['回答の検証ルール・質問内の画像・選択肢の画像・シャッフル設定は同期対象外です。既存の設定は保持します。追加した質問では必要に応じて各フォームで設定してください。','更新中はひな形・配付先フォーム・回答シートを手動編集しないでください。'],fingerprint:fingerprint};
  return {record:r,context:context,plan:plan,target:target,template:template,accepting:accepting,preview:preview};
}

function nativeFormDigest_(metadata) {
  return adminDigest_(managedFormShape_({items:metadata.items,info:metadata.info,settings:metadata.settings}));
}

function nativeExpectedChange_(metadata, operation) {
  const next = JSON.parse(JSON.stringify(metadata)), index = next.items.findIndex(function(i) {return i.itemId === operation.id;});
  if (operation.kind === 'delete') next.items.splice(index,1);
  if (operation.kind === 'move') next.items.splice(operation.index,0,next.items.splice(index,1)[0]);
  if (operation.kind === 'description') next.info.description = operation.value;
  if (operation.kind === 'field') {
    if (['title','description'].includes(operation.key)) next.items[index][operation.key] = operation.value;
    else next.items[index].native[operation.key] = operation.value;
  }
  return next;
}

function nativeNextChange_(plan, current) {
  const equal = function(a,b) {return JSON.stringify(managedFormShape_(a)) === JSON.stringify(managedFormShape_(b));};
  const added = plan.items.find(function(i) {return i.itemId.indexOf('new:') === 0;});
  if (added) return {kind:'create',id:added.itemId,type:added.native.type};
  for (const desired of plan.items) {
    const item = current.items.find(function(i) {return i.itemId === desired.itemId;});
    if (!item) throw new Error('更新対象の質問がなくなりました。フォームの外部変更を確認してください。');
    for (const key of ['title','description']) if (!equal(item[key],desired[key])) return {kind:'field',id:item.itemId,key:key,value:desired[key]};
    for (const key of Object.keys(desired.native).filter(function(k) {return k !== 'type';})) {
      if (!equal(item.native[key],desired.native[key])) return {kind:'field',id:item.itemId,key:key,value:desired.native[key]};
    }
  }
  const removed = current.items.slice().reverse().find(function(i) {return !plan.items.some(function(d) {return d.itemId === i.itemId;});});
  if (removed) return {kind:'delete',id:removed.itemId};
  for (let index=0;index<plan.items.length;index++) if (current.items[index].itemId !== plan.items[index].itemId) return {kind:'move',id:plan.items[index].itemId,index:index};
  if ((current.info.description || '') !== plan.description) return {kind:'description',value:plan.description};
  return null;
}

function nativeApplyChange_(record, form, operation) {
  if (operation.kind === 'create') {form['add' + NATIVE_FORM_TYPES_[operation.type] + 'Item']();return;}
  if (operation.kind === 'description') {form.setDescription(operation.value);return;}
  const item = nativeTypedItem_(form.getItemById(Number(operation.id)));
  if (operation.kind === 'delete') {form.deleteItem(item.getIndex());return;}
  if (operation.kind === 'move') {form.moveItem(item.getIndex(),operation.index);return;}
  const key = operation.key, value = operation.value;
  const setters = {title:'setTitle',description:'setHelpText',required:'setRequired',other:'showOtherOption',rows:'setRows',columns:'setColumns',includesYear:'setIncludesYear',width:'setWidth',videoUrl:'setVideoUrl'};
  if (setters[key]) {item[setters[key]](value);return;}
  if (key === 'bounds') {item.setBounds(value[0],value[1]);return;}
  if (key === 'labels') {item.setLabels(value[0],value[1]);return;}
  if (key === 'alignment') {item.setAlignment(FormApp.Alignment[value]);return;}
  function navigation(option) {return option.goToSectionId ? form.getItemById(Number(option.goToSectionId)).asPageBreakItem() : FormApp.PageNavigationType[option.goToAction];}
  if (key === 'navigation') {item.setGoToPage(navigation(value));return;}
  if (key === 'choices') {
    item.setChoices(value.map(function(c) {return c.goToSectionId || c.goToAction ? item.createChoice(c.value,navigation(c)) : item.createChoice(c.value);}));return;
  }
  if (key === 'imageHash') {
    const sourceId = Object.keys(record.formUpdate.plan.sync.itemMap).find(function(id) {return record.formUpdate.plan.sync.itemMap[id] === operation.id;});
    const source = FormApp.openById(record.input.templateId).getItemById(Number(sourceId)).asImageItem(), blob = source.getImage();
    if (adminDigest_(blob.getBytes()) !== value) throw new Error('ひな形の画像が変更されました。元の画像に戻してから再開してください。');
    item.setImage(blob);return;
  }
  throw new Error('未対応の更新項目です: ' + key);
}

function nativeAssignCreatedId_(plan, temporaryId, actualId) {
  const item = plan.items.find(function(i) {return i.itemId === temporaryId;}), before = managedItemQuestions_(item);
  item.itemId = actualId;
  const after = managedItemQuestions_(item), replacements = {};
  before.forEach(function(q,i) {replacements[q.questionId] = after[i].questionId;});
  Object.keys(plan.sync.itemMap).forEach(function(k) {if (plan.sync.itemMap[k] === temporaryId) plan.sync.itemMap[k] = actualId;});
  Object.keys(plan.sync.questionMap).forEach(function(k) {if (replacements[plan.sync.questionMap[k]]) plan.sync.questionMap[k] = replacements[plan.sync.questionMap[k]];});
  Object.keys(replacements).forEach(function(k) {plan.sync.columns[replacements[k]] = plan.sync.columns[k];delete plan.sync.columns[k];});
  plan.items.forEach(function(i) {(i.native.choices || []).concat(i.native.navigation ? [i.native.navigation] : []).forEach(function(c) {if (c.goToSectionId === temporaryId) c.goToSectionId = actualId;});});
}

/** FormApp の更新操作は記録を残して1件ずつ行い、応答を失った場合は再試行前に結果を照合する。 */
function applyNativeFormUpdate_(record, form) {
  const update = record.formUpdate, plan = update.plan, started = Date.now();
  if (!update.checkpoint) throw new Error('旧方式の更新が途中です。保存済みの更新計画を確認してください。');
  for (let count=0;count<150;count++) {
    if (Date.now() - started > 200000) break;
    let current = formMetadata_(record.formId), operation = update.pending;
    const before = update.checkpoint, unchanged = nativeFormDigest_(current) === nativeFormDigest_(before);
    if (!operation) {
      if (!unchanged) throw new Error('フォームが更新計画と一致しません。外部での変更を確認してください。');
      operation = nativeNextChange_(plan,current);
      if (!operation) {update.after = managedFormContentDigest_(current);return;}
      update.pending = operation;saveManagedRecord_(record);
    }
    if (unchanged) {nativeApplyChange_(record,form,operation);current = formMetadata_(record.formId);}
    if (operation.kind === 'create') {
      const extra = current.items.filter(function(i) {return !before.items.some(function(b) {return b.itemId === i.itemId;});});
      const without = JSON.parse(JSON.stringify(current));without.items = without.items.filter(function(i) {return !extra.some(function(e) {return e.itemId === i.itemId;});});
      if (extra.length !== 1 || extra[0].native.type !== operation.type || extra[0].title || extra[0].description || current.items[current.items.length-1].itemId !== extra[0].itemId || nativeFormDigest_(without) !== nativeFormDigest_(before)) {
        throw new Error('追加した質問の結果を一意に確認できません。フォームの外部変更を確認してください。');
      }
      nativeAssignCreatedId_(plan,operation.id,extra[0].itemId);
    } else if (nativeFormDigest_(current) !== nativeFormDigest_(nativeExpectedChange_(before,operation))) {
      throw new Error('フォームの更新結果が計画と一致しません。外部での変更を確認してください。');
    }
    update.checkpoint = current;delete update.pending;saveManagedRecord_(record);
  }
  throw new Error('更新を分割して保存しました。「更新を再開」で続きを反映してください。');
}

function runManagedFormUpdate(id, fingerprint) {
  return withAppLock_(function() {
    const p = previewManagedFormUpdateUnlocked_(id);
    if (p.preview.fingerprint !== fingerprint) throw new Error('確認後にフォーム・回答列が更新されました。変更内容を再確認してください。');
    const r = p.record;
    delete p.plan.requests;
    r.formUpdate = {phase:'apply',plan:p.plan,accepting:p.accepting,
      checkpoint:p.target,before:managedFormContentDigest_(p.target),after:managedFormContentDigest_({items:p.plan.items,info:{description:p.plan.description}})};
    // 保存容量やサイズによる失敗は、フォームを閉じたり変更したりする前に検出する。
    saveManagedRecord_(r);
    return resumeManagedFormUpdateUnlocked_(r);
  });
}

function resumeManagedFormUpdate(id, revision) {
  return withAppLock_(function() {return resumeManagedFormUpdateUnlocked_(loadManagedRecord_(id,revision));});
}

function resumeManagedFormUpdateUnlocked_(r) {
  if (!r.formUpdate) return r;
  const context = managedUpdateContext_(r), update = r.formUpdate;
  try {
    // Forms と Sheets の両方を読み戻して確認するまで、回答受付を停止したままにする。
    if (context.form.isAcceptingResponses()) context.form.setAcceptingResponses(false);
    let current = managedUpdateMetadata_(r.formId), digest = managedFormContentDigest_(current);
    if (update.phase === 'apply') {
      applyNativeFormUpdate_(r,context.form);
      current = managedUpdateMetadata_(r.formId);digest = managedFormContentDigest_(current);
      if (digest !== update.after) throw new Error('フォームの更新結果を確認できません。「更新を再開」で照合してください。');
      update.phase = 'columns';saveManagedRecord_(r);
    }
    if (digest !== update.after || !update.checkpoint || nativeFormDigest_(current) !== nativeFormDigest_(update.checkpoint)) throw new Error('フォームが更新計画と一致しません。外部での変更を確認してください。');
    if (update.phase === 'columns') {
      const result = arrangeManagedFormColumns_(context.sheet,update.plan);
      r.formSync = update.plan.sync;
      r.formSync.archivedHeaders = result.retained;
      r.baseHeaders = update.plan.order.filter(function(h) {return ![r.input.nameHeader,r.input.statusHeader,SCORING_MANAGEMENT_HEADER].concat(managedCustomColumns_(r.input).map(function(c) {return c.header;})).includes(h);});
      r.baseColumnCount = r.baseHeaders.length;
      update.phase = 'restore';update.result = result;saveManagedRecord_(r);
    }
    if (update.accepting && !managedCloseIsDue_(r)) context.form.setAcceptingResponses(true);
    r.accepting = context.form.isAcceptingResponses();r.lastError = '';r.formUpdatedAt = new Date().toISOString();
    r.lastFormUpdate = update.result;delete r.formUpdate;
    return saveManagedRecord_(r);
  } catch (error) {
    // エラーや状態の書き込みに失敗しても、保存済みの処理は再開できる。
    const saved = loadManagedRecord_(r.id);
    saved.lastError = String(error.message || error).slice(0,500);
    try {saveManagedRecord_(saved);} catch (ignored) { /* 元の処理記録を正とする。 */ }
    throw error;
  }
}

function managedCanonicalSnapshot_(value) {
  if (Array.isArray(value)) return value.map(managedCanonicalSnapshot_);
  if (!value || typeof value !== 'object') return value;
  const result = {};
  Object.keys(value).sort().forEach(function(key) {result[key] = managedCanonicalSnapshot_(value[key]);});
  return result;
}

/** RFC3339 の入力を厳密に確認し、正規化前に実在しない日付を拒否する。 */
function managedFutureTime_(value) {
  const text=typeof value==='string'?value:'';
  const parts=text.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})(?:\.\d{1,3})?(Z|[+-]\d{2}:\d{2})$/);
  if(!parts)throw new Error('日時はタイムゾーンを含む正しい形式で指定してください。');
  const year=Number(parts[1]),month=Number(parts[2]),day=Number(parts[3]),zone=parts[7];
  if(year<2000||month<1||month>12||day<1||day>new Date(Date.UTC(year,month,0)).getUTCDate()||Number(parts[4])>23||Number(parts[5])>59||Number(parts[6])>59||
    (zone!=='Z'&&(Number(zone.slice(1,3))>23||Number(zone.slice(4))>59)))throw new Error('日時が不正です。');
  const time=Date.parse(text);
  if(!isFinite(time))throw new Error('日時が不正です。');
  if(time<=Date.now())throw new Error('未来の日時を指定してください。');
  return new Date(time).toISOString();
}

function managedCloseIsDue_(record) { return !!record.closesAt && Date.parse(record.closesAt)<=Date.now(); }

function applyManagedMaterialState_(record, post) {
  if(!post||!post.id||!['DRAFT','PUBLISHED','DELETED'].includes(post.state))throw new Error('資料の状態を確認できません。投稿結果を照合してください。');
  if(record.materialId&&String(record.materialId)!==String(post.id))throw new Error('資料IDが一致しません。');
  record.materialId=String(post.id);record.scheduledTime=post.scheduledTime||'';
  record.stage=post.state==='PUBLISHED'?'published':post.state==='DELETED'?'deleted':post.scheduledTime?'scheduled':'draft';
  record.postUrl=post.alternateLink||'';record.postedAt=post.creationTime||record.postedAt||'';
  record.materialCheckedAt=new Date().toISOString();delete record.materialChange;
  return record;
}

function previewManagedFormAction(id, action, options) { return withAppLock_(function(){return previewManagedFormActionUnlocked_(id,action,options);}); }
function previewManagedFormActionUnlocked_(id, action, options) {
  const r=loadManagedRecord_(id),desired={};let snapshot;options=options||{};
  if(r.settingsUpdate&&action!=='close')throw new Error('設定の更新が途中です。「設定の更新を再開」を完了してください。');
  if(r.formUpdate&&action!=='close')throw new Error('フォームの更新が途中です。「更新を再開」を完了してください。');
  if(['schedule','reschedule'].includes(action)) {
    desired.scheduledTime=managedFutureTime_(options.scheduledTime);
    if(r.closesAt&&Date.parse(desired.scheduledTime)>=Date.parse(r.closesAt))throw new Error('資料の公開予定は受付終了日時より前にしてください。');
  }
  if(['set-close','cancel-close'].includes(action)) {
    if(r.kind!=='form'||!['registered','published','scheduled','draft','deleted','publish_review','schedule_review'].includes(r.stage))throw new Error('フォームの準備を完了してください。');
    if(action==='set-close') {
      desired.closesAt=managedFutureTime_(options.closesAt);
      let scheduled=r.scheduledTime;
      if(r.materialId){snapshot=managedPostService_(r).get(r.courseId,r.materialId);scheduled=snapshot.state==='DRAFT'?snapshot.scheduledTime:'';}
      else if(r.materialChange&&r.materialChange.scheduledTime)scheduled=r.materialChange.scheduledTime;
      if(scheduled&&Date.parse(desired.closesAt)<=Date.parse(scheduled))throw new Error('受付終了日時は資料の公開予定より後にしてください。');
    }else if(!r.closesAt)throw new Error('受付終了の予約がありません。');
    snapshot={material:snapshot||null,accepting:FormApp.openById(r.formId).isAcceptingResponses()};
  }else if(action==='delete') {
    if(!managedPostId_(r)||r.stage==='deleted')throw new Error('削除対象の投稿がありません。');
    snapshot=managedPostService_(r).get(r.courseId,managedPostId_(r));
  }else if(['publish','schedule','reschedule','cancel-schedule','open','close'].includes(action)) {
    if(['publish','schedule','reschedule'].includes(action)&&r.materialSettingsSaved===false)throw new Error('投稿タイトルと投稿文の「設定を保存」を押してから投稿・予約してください。');
    if(r.kind!=='form'||!(['registered','published','scheduled','draft','deleted','publish_review','schedule_review'].includes(r.stage)||(action==='close'&&r.stage==='delete_review')))throw new Error('フォームの準備が完了していません。');
    if(['publish','schedule'].includes(action)&&!['registered','deleted','draft'].includes(r.stage))throw new Error('投稿済み・予約済み、または投稿結果の確認待ちです。投稿結果を照合してください。');
    if(['reschedule','cancel-schedule'].includes(action)&&!['scheduled','draft'].includes(r.stage))throw new Error('予約を変更できません。公開状態・投稿結果を照合してください。');
    if(action==='cancel-schedule'&&r.stage!=='scheduled')throw new Error('資料の予約がありません。');
    if(['publish','schedule','open'].includes(action)&&managedCloseIsDue_(r))throw new Error('受付終了日時を過ぎています。終了予約を変更・取消してから受付を開始してください。');
    const form=FormApp.openById(r.formId);
    snapshot={published:form.isPublished(),accepting:form.isAcceptingResponses()};
    if(['publish','schedule','reschedule','cancel-schedule'].includes(action)&&r.materialId&&r.stage!=='deleted') {
      snapshot.material=managedPostService_(r).get(r.courseId,r.materialId);
      if(snapshot.material.state!=='DRAFT')throw new Error('資料はすでに公開または削除されています。投稿結果を照合してください。');
    }
    if(!['close','cancel-schedule'].includes(action)) {
      const permissions=publishedPermissions_(r.formId),meta=formMetadata_(r.formId);
      validateResponderAccess_(permissions,r.input.policy,managedRoster_(r.courseId).map(function(s){return s.email;}),meta.settings&&meta.settings.collectsEmail);
      snapshot.permissions=permissions.map(managedCanonicalSnapshot_).sort(function(a,b){return JSON.stringify(a).localeCompare(JSON.stringify(b));});snapshot.settings=meta.settings;
    }
  }else throw new Error('操作が不正です。');
  return {id:id,action:action,title:r.materialTitle||r.title||'',courseId:r.courseId,className:r.className||'',postId:managedPostId_(r)||'',
    description:action==='delete'?(snapshot.text||snapshot.description||''):(r.input&&r.input.description||''),
    postedAt:snapshot.creationTime||r.postedAt||'',url:r.formUrl||'',scheduledTime:desired.scheduledTime||'',closesAt:desired.closesAt||r.closesAt||'',accepting:snapshot.accepting,
    emailNotice:['publish','schedule','reschedule','open'].includes(action)?'各フォームの「設定 → 回答 → メールアドレスを収集する」が「確認済み」になっていることを確認してください。収集方式の自動判定はできません。':'',
    fingerprint:adminDigest_(managedCanonicalSnapshot_([r,action,snapshot,desired]))};
}

function runManagedFormAction(id, action, fingerprint, verifiedEmailConfirmed, options) {
  return withAppLock_(function(){
    const p=previewManagedFormActionUnlocked_(id,action,options);
    if(p.fingerprint!==fingerprint)throw new Error('確認後に内容が更新されました。再確認してください。');
    if(['publish','schedule','reschedule','open'].includes(action)&&verifiedEmailConfirmed!==true)throw new Error('フォームのメール収集が「確認済み」であることを確認してから実行してください。');
    const r=loadManagedRecord_(id);
    if(['publish','schedule','reschedule','cancel-schedule','delete'].includes(action))ScriptApp.requireScopes(ScriptApp.AuthMode.FULL,
      ['https://www.googleapis.com/auth/'+(r.kind==='announcement'?'classroom.announcements':'classroom.courseworkmaterials')]);
    if(action==='set-close') {
      ensureManagedFormScheduleTrigger_();r.closesAt=p.closesAt;r.closeState='scheduled';r.closeError='';delete r.closedAt;
      return saveManagedRecord_(r);
    }
    if(action==='cancel-close') {
      assertManagedFormScheduleOwner_();delete r.closesAt;r.closeState='cancelled';r.closeError='';return saveManagedRecord_(r);
    }
    if(action==='delete') {
      r.stage='delete_review';saveManagedRecord_(r);
      managedPostService_(r).remove(r.courseId,managedPostId_(r));
      r.stage='deleted';r.deletedAt=new Date().toISOString();return saveManagedRecord_(r);
    }
    const form=FormApp.openById(r.formId);
    if(action==='close'){form.setAcceptingResponses(false);r.accepting=form.isAcceptingResponses();if(r.formUpdate)r.formUpdate.accepting=false;if(r.settingsUpdate)r.settingsUpdate.accepting=false;return saveManagedRecord_(r);}
    if(action==='open'){form.setPublished(true);form.setAcceptingResponses(true);r.accepting=form.isAcceptingResponses();return saveManagedRecord_(r);}
    if(r.stage==='deleted'&&r.materialId){r.deletedPosts=(r.deletedPosts||[]).concat([{id:r.materialId,deletedAt:r.deletedAt}]);delete r.materialId;delete r.postUrl;}
    const changing=!!r.materialId;
    r.stage=changing?'schedule_review':'publish_review';r.publishingStartedAt=new Date().toISOString();
    r.materialChange={action:action,scheduledTime:p.scheduledTime};saveManagedRecord_(r);
    if(['publish','schedule'].includes(action)){form.setPublished(true);form.setAcceptingResponses(true);}
    const service=Classroom.Courses.CourseWorkMaterials;let post;
    if(changing) {
      const body=action==='publish'?{state:'PUBLISHED'}:p.scheduledTime?{scheduledTime:p.scheduledTime}:{};
      post=service.patch(body,r.courseId,r.materialId,{updateMask:action==='publish'?'state,scheduledTime':'scheduledTime'});
    }else {
      const body={title:r.materialTitle,description:r.input.description,state:action==='schedule'?'DRAFT':'PUBLISHED',materials:[{link:{url:form.getPublishedUrl()}}]};
      if(p.scheduledTime)body.scheduledTime=p.scheduledTime;
      post=service.create(body,r.courseId);
    }
    if(!post||!post.id)throw new Error('投稿IDを確認できません。Classroomの投稿結果を照合してください。');
    // 即時作成の従来の応答形式を維持する（ID とリンクだけで十分）。
    if(!changing&&action==='publish'&&!post.state)post.state='PUBLISHED';
    applyManagedMaterialState_(r,post);r.accepting=form.isAcceptingResponses();return saveManagedRecord_(r);
  });
}

/** 結果が不明な場合は投稿を再作成せずに照合する。時刻だけでは公開済みと判断しない。 */
function reconcileManagedPost(id, expectedRevision) {
  return withAppLock_(function(){
    const r=loadManagedRecord_(id,expectedRevision),service=managedPostService_(r);
    if(r.stage==='publish_review'&&!r.materialId){
      const matches=[];let token;
      do{
        const options={pageSize:100,courseWorkMaterialStates:['DRAFT','PUBLISHED']};if(token)options.pageToken=token;
        const page=service.list(r.courseId,options);
        (page.courseWorkMaterial||[]).forEach(function(p){
          const sameLink=(p.materials||[]).some(function(m){return m.link&&m.link.url===r.formUrl||m.form&&m.form.formUrl===r.formUrl;});
          const intended=r.materialChange&&r.materialChange.scheduledTime;
          const stateMatches=intended?(p.state==='PUBLISHED'||p.state==='DRAFT'&&Date.parse(p.scheduledTime)===Date.parse(intended)):p.state==='PUBLISHED';
          if(sameLink&&p.title===r.materialTitle&&(p.description||'')===r.input.description&&stateMatches&&Date.parse(p.creationTime)>=Date.parse(r.publishingStartedAt)-5000)matches.push(p);
        });token=page.nextPageToken;
      }while(token);
      if(matches.length!==1)throw new Error('投稿結果を一意に照合できません。Classroomで確認してください。自動では再投稿しません。');
      applyManagedMaterialState_(r,matches[0]);return saveManagedRecord_(r);
    }
    if(r.stage==='delete_review'){
      const post=service.get(r.courseId,managedPostId_(r));
      // アクセス権を失った場合も 404 になるため、明示的な状態でのみ削除を確認する。
      if(post.state!=='DELETED')throw new Error('投稿が残っています。対象を確認して削除を再実行できます。');
      r.stage='deleted';r.deletedAt=new Date().toISOString();return saveManagedRecord_(r);
    }
    if(r.kind==='form'&&r.materialId&&['scheduled','draft','published','schedule_review'].includes(r.stage)) {
      applyManagedMaterialState_(r,service.get(r.courseId,r.materialId));return saveManagedRecord_(r);
    }
    return r;
  });
}


function recordManagedAnnouncement_(announcement, info, cache) {
  if (!announcement || !announcement.id) throw new Error('投稿IDを取得できませんでした。');
  return saveManagedRecord_({id:'announcement:'+info.courseId+':'+announcement.id,kind:'announcement',postId:String(announcement.id),courseId:info.courseId,
    title:info.name+'さんへの返却',sourceKey:info.sourceKey,stage:'published',postedAt:announcement.creationTime||new Date().toISOString(),postUrl:announcement.alternateLink||''},cache);
}

function assertWebOperator_() {
  const active = Session.getActiveUser().getEmail();
  const effective = Session.getEffectiveUser().getEmail();
  if (!active || !effective || active.toLowerCase() !== effective.toLowerCase()) {
    throw new Error('この管理画面はデプロイした管理者本人として開いてください。');
  }
}

/** 採点テンプレの専用画面と埋め込み画面で同じHTML・配色を使う。配置ファイルは増やさない。 */
function createAppHtmlOutput_(name) {
  const html = HtmlService.createHtmlOutputFromFile(name === 'Templates' ? 'Scoring' : name);
  if (name === 'Templates') {
    const editor=html.getContent().match(/<template id="ruleBuilderTemplate">([\s\S]*?)<\/template>/);
    if(!editor)throw new Error('採点テンプレ作成画面が見つかりません。Scoring.html を更新してください。');
    const url=getVerifiedWebConsoleUrl_();
    const navigation='<nav class="toolbar" aria-label="画面の移動"><strong>採点テンプレ作成</strong>'+
      (url?' <a class="link" href="'+escapeWebLink_(url)+'" target="_blank" rel="noopener">管理画面</a> <a class="link" href="'+escapeWebLink_(url+'?page=scoring')+'" target="_blank" rel="noopener">採点画面</a>':'')+'</nav>';
    html.setContent(editor[1].replace('<body>','<body>'+navigation));
  }
  if (name === 'Scoring' || name === 'Templates') {
    const setting = HtmlService.createHtmlOutputFromFile('Setting').getContent();
    const theme = setting.match(/<style id="app-theme">[\s\S]*?<\/style>/);
    if (!theme) throw new Error('共通配色が見つかりません。Setting.html を更新してください。');
    html.setContent(html.getContent().replace(/<head>/g, '<head>\n' + theme[0]));
  }
  return html;
}

function doGet(e) {
  assertWebOperator_();
  rememberWebConsoleUrl_();
  const page=e && e.parameter && e.parameter.page;
  const name=page==='scoring'?'Scoring':page==='templates'?'Templates':'Setting';
  const html=createAppHtmlOutput_(name);
  return html.setTitle(name==='Scoring'?'turret 採点画面':name==='Templates'?'turret 採点テンプレ作成':'turret 管理画面').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function connectWebConsole() {
  return connectWebScreen_('settings');
}

function connectScoringConsole() {
  return connectWebScreen_('scoring');
}

function openScoringRuleEditor() {
  return connectWebScreen_('templates');
}

function connectWebScreen_(page) {
  const label = page === 'scoring' ? '採点画面' : page === 'templates' ? '採点テンプレ作成画面' : '管理画面';
  const result = withAppLock_(function() {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (!active) throw new Error('接続するスプシのメニューから実行してください。');
    PropertiesService.getScriptProperties().setProperty('TURRET_WEB_TARGET', active.getId());
    PropertiesService.getScriptProperties().setProperty('TURRET_BOUND_TEMPLATE', active.getId());
    const url = getVerifiedWebConsoleUrl_();
    return url && (page === 'scoring' || page === 'templates') ? url + '?page=' + page : url;
  });
  const body = result ?
    '<p id="openStatus" role="status">' + label + 'を別のタブで開いています。開かない場合は下のリンクを押してください。</p>' +
    '<p><a href="' + escapeWebLink_(result) + '" target="_blank" rel="noopener">' + label + 'を開く</a></p>' +
    '<label for="webUrl">WebアプリのURL</label><input id="webUrl" readonly value="' + escapeWebLink_(result) + '" onclick="this.select()">' +
    '<button onclick="copyUrl()">URLをコピー</button><span id="copyStatus" role="status"></span>' +
    '<script>async function copyUrl(){const input=document.getElementById("webUrl");input.focus();input.select();let copied=false;' +
    'try{await navigator.clipboard.writeText(input.value);copied=true;}catch(e){try{copied=document.execCommand("copy");}catch(ignored){}}' +
    'document.getElementById("copyStatus").textContent=copied?"コピーしました":"選択したURLを Ctrl+C / ⌘C でコピーしてください";}' +
    'try{const tab=window.open("about:blank","_blank");if(tab){tab.opener=null;tab.location.href=document.getElementById("webUrl").value;google.script.host.close();}' +
    'else{document.getElementById("openStatus").textContent="自動で開けませんでした。下のリンクから'+label+'を開いてください。";}}' +
    'catch(e){document.getElementById("openStatus").textContent="下のリンクから'+label+'を開いてください。";}</script>' :
    '<p>このスプシを接続しました。公開したWeb管理画面を一度開くと、ここに正しいURLが登録されます。</p>' +
    '<p>Apps Scriptの「デプロイ → デプロイを管理」でWebアプリのURLを開いてください。未公開なら「新しいデプロイ → ウェブアプリ」で作成します。</p>' +
    '<p>実行ユーザー：自分 ／ アクセスできるユーザー：自分のみ</p>';
  const html = HtmlService.createHtmlOutput('<!doctype html><html lang="ja"><head><meta charset="utf-8"><style>' +
    'body{font:14px sans-serif;padding:12px;color:#183536}input{display:block;box-sizing:border-box;width:100%;padding:10px;margin:8px 0 12px}' +
    'a{color:#126d63}button{padding:8px 14px;cursor:pointer}#copyStatus{display:block;margin-top:10px}</style></head><body>' + body + '</body></html>').setWidth(620).setHeight(280);
  SpreadsheetApp.getUi().showModelessDialog(html, label);
}

/** doGet からだけ呼ぶ。コンテナに紐付く getUrl() は、コピー元の古い URL を返す場合がある。 */
function rememberWebConsoleUrl_() {
  const url = ScriptApp.getService().getUrl();
  if (!isPublishedWebConsoleUrl_(url)) return;
  const props = PropertiesService.getScriptProperties();
  const entry = JSON.stringify({ scriptId: ScriptApp.getScriptId(), url: url });
  if (props.getProperty('TURRET_WEB_ENTRY') !== entry) props.setProperty('TURRET_WEB_ENTRY', entry);
}

function isPublishedWebConsoleUrl_(url) {
  // 学校ドメイン付きURLは、デプロイ画面とgetUrl()でドメインの位置が異なる場合がある。
  return typeof url === 'string' && /^https:\/\/script\.google\.com\/(?:macros|a\/macros\/[A-Za-z0-9.-]+|a\/[A-Za-z0-9.-]+\/macros)\/s\/[A-Za-z0-9_-]+\/exec$/.test(url);
}

function getVerifiedWebConsoleUrl_() {
  let entry;
  try { entry = JSON.parse(PropertiesService.getScriptProperties().getProperty('TURRET_WEB_ENTRY') || 'null'); }
  catch (error) { return null; }
  return entry && entry.scriptId === ScriptApp.getScriptId() && isPublishedWebConsoleUrl_(entry.url) ? entry.url : null;
}

function escapeWebLink_(value) {
  return String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function getWebConsoleBootstrap() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) assertWebOperator_();
  const id = active ? active.getId() : PropertiesService.getScriptProperties().getProperty('TURRET_WEB_TARGET');
  return { web: !active, connected: !!id, spreadsheetId: id || '', canCopyOperation:!!active || PropertiesService.getScriptProperties().getProperty('TURRET_BOUND_TEMPLATE')===id,
    url: id ? 'https://docs.google.com/spreadsheets/d/' + id + '/edit' : '',
    settingsUrl: getVerifiedWebConsoleUrl_() || '',
    scoringUrl: getVerifiedWebConsoleUrl_() ? getVerifiedWebConsoleUrl_() + '?page=scoring' : '' };
}

function createWebOperation(name) {
  return withAppLock_(function() {
    if (SpreadsheetApp.getActiveSpreadsheet() || PropertiesService.getScriptProperties().getProperty('TURRET_WEB_TARGET')) {
      throw new Error('既存の運用に接続済みです。別の運用は別のGASプロジェクトで作成してください。');
    }
    assertWebOperator_();
    name = String(name || '').trim();
    if (!name || name.length > 100) throw new Error('運用名を1〜100文字で入力してください。');
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty('TURRET_WEB_CREATING')) throw new Error('前回の作成結果を確認してください。Driveに作成されたスプシを確認し、重複作成を避けてください。');
    props.setProperty('TURRET_WEB_CREATING', name);
    const ss = SpreadsheetApp.create(name);
    props.setProperty('TURRET_WEB_TARGET', ss.getId());
    ss.setSpreadsheetTimeZone('Asia/Tokyo');
    props.deleteProperty('TURRET_WEB_CREATING');
    initializeSheetsUnlocked_();
    return getWebConsoleBootstrap();
  });
}

function previewNextOperation(name, folder) {
  return withAppLock_(function(){return previewNextOperationUnlocked_(name,folder);});
}
function previewNextOperationUnlocked_(name, folder) {
  if(!getWebConsoleBootstrap().canCopyOperation)throw new Error('この運用は独立GASから作成されています。次期間用には別GASへ5ファイルを配置し、設定JSONを再利用してください。');
  name=String(name||'').trim();if(!name||name.length>100)throw new Error('新しい運用名を1〜100文字で入力してください。');
  const folderId=managedId_(folder);DriveApp.getFolderById(folderId).getName();
  const ss=getAppSpreadsheet_(),config=getConfig_();
  const classes=adminSheetValues_(ss.getSheetByName('クラス一覧')),students=adminSheetValues_(ss.getSheetByName(STUDENT_SHEET_NAME));
  return {name:name,folderId:folderId,sourceId:ss.getId(),fingerprint:adminDigest_([name,folderId,ss.getId(),config,classes,students,getAutomationSchedule_()])};
}
function createNextOperation(name, folder, fingerprint) {
  return withAppLock_(function(){
    const p=previewNextOperationUnlocked_(name,folder);if(p.fingerprint!==fingerprint)throw new Error('確認後に設定が更新されました。');
    let record=getManagedRecords_().find(function(r){return r.kind==='copy'&&r.requestFingerprint===fingerprint;});
    if(record){assertManagedOwner_(record);if(record.stage==='ready')return nextOperationResult_(record);
      if(record.stage==='sanitizing')throw new Error('複製先の初期化結果を確認してください。新しい回答を消さないよう、初期化を自動で繰り返しません。複製先: https://docs.google.com/spreadsheets/d/'+record.copyId+'/edit');
      if(!record.copyId)throw new Error('複製結果をDriveで確認してください。自動では再複製しません。');}
    else{
      record=saveManagedRecord_({id:Utilities.getUuid(),kind:'copy',stage:'creating',requestFingerprint:fingerprint,title:p.name});
      const file=Drive.Files.copy({name:p.name,parents:[p.folderId]},p.sourceId,{fields:'id'});
      if(!file||!file.id)throw new Error('複製したスプシIDを確認できません。Driveで作成結果を確認してください。');
      record.copyId=file.id;record.stage='copied';saveManagedRecord_(record);
    }
    if(record.copyId===p.sourceId)throw new Error('複製元は変更できません。');
    const copy=SpreadsheetApp.openById(record.copyId);
    record.stage='sanitizing';saveManagedRecord_(record);
    // 新規作成したコピーだけを初期化し、コピー元の運用データは消さない。
    const keep=['クラス一覧',STUDENT_SHEET_NAME];
    // 設定前に作成したコピーでも、表示可能な名簿タブを必ず残す。
    const classes=copy.getSheetByName('クラス一覧')||copy.insertSheet('クラス一覧');
    if(!classes.getLastRow())classes.getRange(1,1,1,3).setValues([['クラス名','コースID','同期対象(1)']]);
    classes.showSheet();
    copy.getSheets().forEach(function(s){if(!keep.includes(s.getName()))copy.deleteSheet(s);});
    SpreadsheetApp.flush();
    record.stage='ready';record.url='https://docs.google.com/spreadsheets/d/'+record.copyId+'/edit';return nextOperationResult_(saveManagedRecord_(record));
  });
}

function nextOperationResult_(record) {
  const config=normalizeAppConfig_(getConfig_());config.formSources=[];
  return Object.assign({},record,{settingsProfile:buildSettingsProfile_(config,true)});
}

// 既存のスプレッドシート内ボタンとの互換性を保つ。管理画面はすべて Web で開く。
function openMaintenanceDialog() { return connectWebConsole(); }
function openManualActionsDialog() { return connectWebConsole(); }
function openClassesSheet() { return openAdminSheet('classes'); }
function openStudentsSheet() { return openAdminSheet('students'); }
function openMappingSheet() { return openAdminSheet('mapping'); }

function adminSheetValues_(sheet) {
  if (!sheet || !sheet.getLastRow() || !sheet.getLastColumn()) return [];
  return sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getDisplayValues();
}

function adminNonemptyRows_(values) {
  return values.slice(1).filter(function(row) {
    return row.some(function(value) { return String(value == null ? '' : value).trim() !== ''; });
  });
}

function adminDigest_(value) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify(value), Utilities.Charset.UTF_8)
    .map(function(byte) { return ('0' + ((byte + 256) % 256).toString(16)).slice(-2); }).join('');
}

function adminActionRevision_(action, config, snapshots) {
  const names = { send: SEND_SHEET_NAME, clearEval: EVAL_SHEET_NAME, clearSend: SEND_SHEET_NAME, clearErrors: ERROR_SHEET_NAME };
  const name = names[action];
  const values = !name ? [] : snapshots && Object.prototype.hasOwnProperty.call(snapshots,name)
    ? snapshots[name] : adminSheetValues_(getAppSpreadsheet_().getSheetByName(name));
  return adminDigest_([action, getConfigRevision_(config), values]);
}

function getAdminConsoleData() { return withAppLock_(getAdminConsoleDataUnlocked_); }

function getAdminConsoleDataUnlocked_() {
  let config, recoveryRequired = false;
  const warnings = [];
  try { migrateConfigStorageUnlocked_(); config = getConfig_(); } catch (error) {
    recoveryRequired = true;
    warnings.push('設定の復旧が必要です。自動実行を停止し、内部設定と設定JSONを確認してください。' + String(error.message || error));
    // 表示専用の代替処理。書き込みと送信では必ず、不正な設定を拒否する getConfig_ を呼ぶ。
    try { config = loadConfig_(); } catch (readError) {
      config = buildDefaultConfig_();
      warnings.push('現在の設定を表示できないため、入力欄は空です。空欄を保存済み設定とみなさないでください。');
    }
  }
  const ss = getAppSpreadsheet_();
  const values = {}, snapshots = {};
  const sheets = ADMIN_SHEETS.map(function(item) {
    const sheet = ss.getSheetByName(item.label);
    values[item.key] = adminSheetValues_(sheet);
    snapshots[item.label] = values[item.key];
    return { key: item.key, label: item.label, exists: !!sheet, rows: adminNonemptyRows_(values[item.key]).length,
      url: sheet ? adminSheetUrl_(ss, sheet) : '' };
  });
  const classes = adminNonemptyRows_(values.classes);
  const selected = classes.filter(function(row) { return String(row[2] || '').trim() === '1' && String(row[1] || '').trim(); });
  const selectedIds = selected.map(function(row) { return String(row[1]).trim(); }).sort();
  const allStudents=adminNonemptyRows_(values.students),studentMap = createHeaderMap_(values.students[0] || []);
  const students = allStudents.filter(function(row){return String(row[studentMap['除外(1)']]||'').trim()!=='1';});
  const studentCourses = new Set(students.map(function(row) { return String(row[studentMap['コースID']] || '').trim(); }));
  const lastSelection = PropertiesService.getScriptProperties().getProperty('TURRET_ROSTER_SELECTION');
  const rosterCurrent = selectedIds.length > 0 && lastSelection === JSON.stringify(selectedIds);
  const validStudents = students.length > 0 && students.every(function(row) {
    return ['メールアドレス', '名前', 'コースID', 'studentId'].every(function(header) { return String(row[studentMap[header]] || '').trim(); });
  });
  const studentsReady = rosterCurrent && validStudents && selectedIds.every(function(id) { return studentCourses.has(id); });
  const emailCourses = new Map();
  allStudents.forEach(function(row) {
    const email = String(row[studentMap['メールアドレス']] || '').trim().toLowerCase();
    const course = String(row[studentMap['コースID']] || '').trim();
    if (!email || !course) return;
    if (!emailCourses.has(email)) emailCourses.set(email, new Set());
    emailCourses.get(email).add(course);
  });
  const mappingNeeded = Array.from(emailCourses.values()).some(function(courses) { return courses.size > 1; });
  const mappings = adminNonemptyRows_(values.mapping);
  const mappingReady = mappings.length > 0 && mappings.every(function(row) {
    return row[0] && row[2] && row[3] && selectedIds.indexOf(String(row[4] || '').trim()) >= 0;
  });
  const mappingOptional = studentsReady && !mappingNeeded && !mappings.length;
  const sourceReady = config.formSources.length > 0 && config.formSheetNamePrefix.length > 0;
  const notificationsReady = automationRecipientsValid_(config) || config.reminderOptOutConfirmed === true;
  let configError = '';
  try { validateAppConfig_(config); } catch (error) { configError = String(error.message || error); }
  let fieldsReady = false;
  try {
    const check = Object.assign({}, config, { messageTemplate: '確認用' });
    validateAppConfig_(check);
    fieldsReady = true;
  } catch (error) { /* 未完了項目は手順内で案内する。 */ }
  const templateWarnings = adminUnknownTokens_(config);
  const templateReady = !!config.messageTemplate.trim() && templateWarnings.length === 0;
  let automation;
  try { automation = getAutomationSummary_(); } catch (error) {
    automation = { enabled: false };
    recoveryRequired = true;
    warnings.push('自動実行の情報を読み取れません。Apps Scriptのトリガー画面で状態を確認してください。');
  }
  const sendHeaderMap = createHeaderMap_(values.send[0] || []);
  const sendRows = adminNonemptyRows_(values.send);
  const legacyHeader = createHeaderMap_(values.evaluation[0] || []);
  const counts = { classes: classes.length, selectedClasses: selected.length, students: students.length, mappings: mappings.length,
    evalRows: adminNonemptyRows_(values.evaluation).length, sendRows: sendRows.length,
    legacyPending: adminNonemptyRows_(values.evaluation).filter(function(row) { return String(row[legacyHeader['処理状態']] || '').trim() !== '準備〇'; }).length,
    pendingSend: sendRows.filter(function(row) { return ['', '未'].indexOf(String(row[sendHeaderMap['送信状態']] || '').trim()) >= 0; }).length,
    reviewSend: sendRows.filter(function(row) { return String(row[sendHeaderMap['送信状態']] || '').trim() === SEND_REVIEW_STATUS; }).length,
    errorRows: adminNonemptyRows_(values.errors).length };
  const prepared = !!managedSheet_(false) && ['classes', 'students', 'send'].every(function(key) { return sheets.some(function(sheet) { return sheet.key === key && sheet.exists; }); });
  const progress = [
    { id: 'prepare', title: 'シートの準備', state: prepared ? 'complete' : 'pending', detail: prepared ? '必要なシートを作成済み' : '最初に管理用シートを準備します' },
    { id: 'classes', title: 'クラスの選択', state: selected.length ? 'complete' : 'pending', detail: selected.length ? selected.length + 'クラスを選択中' : '一覧を取得し、同期対象に1を入力します' },
    { id: 'students', title: '生徒の取得', state: studentsReady ? 'complete' : (students.length ? 'attention' : 'pending'), detail: studentsReady ? students.length + '名の名簿を取得済み' : '対象クラスの生徒一覧を取得・確認してください' },
    { id: 'sources', title: '回答先と通知先', state: sourceReady && notificationsReady ? 'complete' : 'pending', detail: !sourceReady ? '「フォームを準備」で回答先を自動接続します' : !notificationsReady ? '通知先を登録するか、メール通知を利用しないことを確認して保存してください' : config.reminderOptOutConfirmed ? '回答先を登録済み。メール通知を利用しないことを確認済み' : '回答先と通知先メールを登録済み' },
    { id: 'mapping', title: 'クラスとの対応', state: !studentsReady ? 'pending' : (mappingOptional || mappingReady ? 'complete' : 'attention'), detail: !studentsReady ? '生徒一覧を取得すると、対応表が必要か確認できます' : (mappingOptional ? '任意。同じ生徒が複数クラスにいなければ省略できます' : (mappingReady ? mappings.length + '件を割り当て済み。対象フォームとの対応を確認してください' : '対応表で各フォームの送信先クラスを選んでください')) },
    { id: 'fields', title: '列と取り込み項目', state: fieldsReady ? 'complete' : 'pending', detail: fieldsReady ? config.fields.length + '項目を登録済み' : 'フォームの列と、その用途を指定します' },
    { id: 'template', title: '返信本文', state: templateReady ? 'complete' : (config.messageTemplate.trim() ? 'attention' : 'pending'), detail: templateReady ? '本文を登録済み。プレビューで内容を確認できます' : (templateWarnings[0] || '送る文面を作成します') },
    { id: 'automation', title: '自動実行', state: automation.enabled ? 'complete' : 'pending', detail: automation.enabled ? '管理トリガーを設定済み。自動実行画面で稼働状態を確認してください' : '希望の時間帯を選び、確認して開始します' }
  ];
  try {
    const formProgress = getFormSetupProgress_(sourceReady);
    progress.push(formProgress.forms, formProgress.publish);
  } catch (error) {
    ['forms','publish'].forEach(function(id) { progress.push({id:id,title:id==='forms'?'フォームを準備':'資料を投稿',state:'attention',detail:'フォームの履歴を確認できません。該当手順で再取得してください。'}); });
  }
  const actionRevisions = {};
  ['send', 'remind', 'clearEval', 'clearSend', 'clearErrors'].forEach(function(action) { actionRevisions[action] = adminActionRevision_(action, config, snapshots); });
  const configParseError = PropertiesService.getScriptProperties().getProperty('APP_CONFIG_PARSE_ERROR') === '1';
  if (configParseError) warnings.push('保存済み設定の読み込みに失敗しました。内部バックアップから表示した設定と設定JSONを確認してから保存してください。');
  return { config: config, revision: getConfigRevision_(config), progress: progress, counts: counts, sheets: sheets,
    health: buildAdminHealth_(values, recoveryRequired || configParseError),
    ready: !recoveryRequired && !configParseError && prepared && !configError && notificationsReady && studentsReady && (!mappingNeeded || mappingReady) && templateReady,
    mappingNeeded: mappingNeeded, configParseError: configParseError, recoveryRequired: recoveryRequired, warnings: warnings,
    initialPanel: PropertiesService.getUserProperties().getProperty('TURRET_ADMIN_PANEL') || 'setup', actionRevisions: actionRevisions };
}

/** 読み取り専用の復旧案内。ログは履歴であり、問題が未解決である証拠とは限らない。 */
function buildAdminHealth_(values, recoveryRequired) {
  const items = [];
  const add = function(code, title, count, detail, action, sheetKey) {
    if (count) items.push({ code: code, title: title, count: count, detail: detail, action: action, sheetKey: sheetKey });
  };
  const send = values.send || [], map = createHeaderMap_(send[0] || []);
  const countStatus = function(status) { return adminNonemptyRows_(send).filter(function(row) { return String(row[map['送信状態']] || '').trim() === status; }).length; };
  add('CONFIG_RECOVERY', '設定の復旧が必要', recoveryRequired ? 1 : 0, '保存または読み取りに問題があります。', '自動実行を停止し、バックアップと設定を照合してください。READMEの設定保存の復旧手順を参照してください。', '');
  add('SEND_REVIEW', '投稿結果を確認する', countStatus(SEND_REVIEW_STATUS), '自動再送を止めている行です。', 'Classroomの投稿とエラーの投稿IDを照合。投稿済みなら「済」、未投稿と確認できた場合だけ「未」に戻します。不明なまま再送しないでください。', 'send');
  add('SEND_ERROR', '送信できなかった行', countStatus('エラー'), '宛先や権限などの確認が必要です。', 'エラーシートの原因を解消し、宛先・本文と未投稿を確認してから対象の送信状態を「未」に戻します。', 'send');
  add('SEND_RETRY', '再送待ち', countStatus('未'), '送信を拒否された行で、次回の送信対象です。', '制限や設定を確認し、次回の自動実行を待つか、手動実行から送信内容を確認してください。', 'send');
  const evaluation = values.evaluation || [], evalMap = createHeaderMap_(evaluation[0] || []);
  add('PREPARE_ERROR', '旧評価データの移行で準備できなかった行', adminNonemptyRows_(evaluation).filter(function(row) { return row[evalMap['処理状態']] === '準備×'; }).length,
    '生徒・対応表・本文を確認してください。', '設定診断で原因を確認し、修正後に対象の評価データの処理状態を空欄にして本文を再生成します。既存の送信行がある場合は先に照合してください。', 'evaluation');
  const errors = adminNonemptyRows_(values.errors || []);
  add('FORM_UPDATE_ERROR', '元フォームへの書き戻しの記録', errors.filter(function(row) { return row[5] === 'FORM_UPDATE_ERROR'; }).length,
    '過去の失敗ログです。すでに復旧済みの場合も残ります。', '元フォームのシート名・状態列・アクセス権を確認し、評価データ・送信シートと照合して元回答の状態を補正してください。投稿済みの行を再送しないでください。', 'errors');
  add('ERROR_LOG', 'エラー履歴', errors.length, '過去のログ件数です。現在の未解決件数とは異なります。', '日時と対象行を照合してください。ログのクリアだけでは原因や送信状態は修復されません。', 'errors');
  let lastRuns = [];
  try {
    const runs = JSON.parse(PropertiesService.getScriptProperties().getProperty('APP_AUTOMATION_RUNS') || '[]');
    if (!Array.isArray(runs)) throw new Error('履歴の形式');
    const seen = new Set();
    lastRuns = runs.filter(function(run) {
      if (!run || ['import', 'delivery', 'reminder'].indexOf(run.kind) < 0 || seen.has(run.kind)) return false;
      seen.add(run.kind); return true;
    });
    add('AUTOMATION_RESULT', '直近の自動実行を確認', lastRuns.filter(function(run) { return run.status !== 'success'; }).length,
      '処理ごとの直近の実行に失敗・要確認の記録があります。', '自動実行画面の実行結果とApps Scriptの実行履歴を確認し、対象行とエラーを照合してください。', 'errors');
  } catch (error) {
    add('AUTOMATION_HISTORY', '実行履歴を読み取れません', 1, '履歴の保存内容を確認できません。', 'Apps Scriptの実行履歴から直近の処理結果を確認してください。', 'errors');
  }
  return { status: items.length ? 'attention' : 'ok', items: items, lastRuns: lastRuns };
}

/** この診断では作成・書き込み関数も Classroom 呼び出しも使わない。 */
function diagnoseSetup(expectedRevision) {
  return withAppLock_(function() {
    const config = getConfig_(), revision = getConfigRevision_(config);
    if (revision !== expectedRevision) throw new Error('設定が変更されました。再読込してから診断してください。');
    const result = { checkedAt: new Date().toISOString(), revision: revision, status: 'ok',
      counts: { sources: 0, sheets: 0, eligible: 0, previewed: 0, blocked: 0, skipped: 0 }, checks: [], samples: [], truncated: false };
    let rowBudget = 5000, number = 1;
    const check = function(severity, title, detail, action) {
      if (severity === 'error') result.status = 'error';
      else if (severity === 'warning' && result.status === 'ok') result.status = 'attention';
      if (result.checks.length < 100) result.checks.push({ severity: severity, title: title, detail: String(detail), action: action || '' });
    };
    const read = function(sheet) {
      if (!sheet || !sheet.getLastRow() || !sheet.getLastColumn()) return { raw: [], display: [] };
      const n = Math.min(Math.max(0, sheet.getLastRow() - 1), rowBudget);
      if (n < sheet.getLastRow() - 1) result.truncated = true;
      rowBudget -= n;
      const range = sheet.getRange(1, 1, n + 1, sheet.getLastColumn());
      return { raw: range.getValues(), display: range.getDisplayValues() };
    };
    const requireHeaders = function(values, required, name, ordered) {
      const headers = new Set(values[0] || []), missing = required.filter(function(h) { return !headers.has(h); });
      if (missing.length) { check('error', name + 'の列', '不足している列: ' + missing.join('、'), '列と項目の設定、またはシートの見出しを確認してください。'); return false; }
      const duplicates = (values[0] || []).filter(function(h, i, all) { return h && all.indexOf(h) !== i; });
      if (duplicates.length) { check('error', name + 'の列', '重複した列名: ' + duplicates.join('、'), '同じ名前の列を整理してください。'); return false; }
      if (ordered && ((values[0] || []).length !== required.length || required.some(function(h, i) { return values[0][i] !== h; }))) {
        check('error', name + 'の列順', '保存された設定と列の順序・数が一致しません。', 'データを退避し、設定とシートの見出しを照合してください。'); return false;
      }
      return true;
    };
    if (PropertiesService.getScriptProperties().getProperty('APP_CONFIG_PARSE_ERROR') === '1') {
      check('error', '保存済み設定の読み取り', '設定の読み取りに失敗した記録があります。', '設定とバックアップを確認してから設定を保存してください。'); return result;
    }
    try { validateAppConfig_(config); } catch (error) {
      check('error', '設定の入力', error.message || error, '設定の各手順を保存してください。'); return result;
    }
    const tokenWarnings = adminUnknownTokens_(config);
    if (tokenWarnings.length) { check('error', '本文の差し込み', tokenWarnings.join(' '), '本文の候補から差し込み名を選び直してください。'); return result; }
    const ss = getAppSpreadsheet_();
    const studentValues = read(ss.getSheetByName(STUDENT_SHEET_NAME)).display;
    if (!requireHeaders(studentValues, ['メールアドレス', '名前', 'クラス名', 'コースID', 'studentId'], '生徒一覧')) return result;
    const sh = createHeaderMap_(studentValues[0]), students = new Map();
    studentValues.slice(1).forEach(function(row) {
      const email = String(row[sh['メールアドレス']] || '').trim().toLowerCase();
      if (!email) return;
      const list = students.get(email) || [];
      list.push({ name: String(row[sh['名前']] || '').trim(), className: String(row[sh['クラス名']] || '').trim(),
        courseId: String(row[sh['コースID']] || '').trim(), studentId: String(row[sh.studentId] || '').trim(),excluded:String(row[sh['除外(1)']]||'').trim()==='1' });
      students.set(email, list);
    });
    if (!students.size) check('error', '生徒一覧', '照合できる生徒がいません。', 'クラスを選択して生徒一覧を取得してください。');
    const mappings = new Map(), mappingValues = read(ss.getSheetByName(MAPPING_SHEET_NAME)).display;
    if (mappingValues.length > 1) {
      if (!requireHeaders(mappingValues, ['元SS_ID', '元シート名', 'クラス名', 'courseId', 'メモ'], '対応表')) return result;
      const mh = createHeaderMap_(mappingValues[0]);
      mappingValues.slice(1).forEach(function(row) {
        if (!row.some(function(v) { return String(v).trim(); })) return;
        const key = makeMappingLookupKey_(row[mh['元SS_ID']], row[mh['元シート名']]);
        if (mappings.has(key)) check('error', '対応表', '同じ元シートの割り当てが重複しています。', '対応表を1シートにつき1行に整理してください。');
        mappings.set(key, String(row[mh.courseId] || '').trim());
      });
      if (result.status === 'error') return result;
    }
    const evalValues = read(ss.getSheetByName(EVAL_SHEET_NAME)).display;
    const sendSheet = ss.getSheetByName(SEND_SHEET_NAME);
    const sendValues = read(sendSheet).display;
    const legacyStateColumn = (evalValues[0] || []).indexOf('処理状態');
    const legacyPending = evalValues.slice(1).some(function(row) { return row.some(function(v) { return String(v).trim(); }) && row[legacyStateColumn] !== '準備〇'; });
    let evalValid = evalValues.length <= 1 || requireHeaders(evalValues, legacyPending ? getConfiguredEvalHeaders_(config) : EVAL_BASE_HEADERS, '旧評価データ', false);
    if (legacyPending) {
      try { validateLegacyEvaluationFields_(config); } catch (error) {
        evalValid = false;
        check('error', '旧評価データの項目列', error.message, '旧項目の列対応を確認し、移行を完了してから項目を変更してください。');
      }
    }
    const sendHeaders = getConfiguredSendHeaders_(config);
    // 初期準備の空シートは、送信データ準備時に現在の設定で見出しを整える。
    const emptySend = !sendSheet || sendSheet.getLastRow() <= 1;
    const sendValid = emptySend || requireHeaders(sendValues, sendHeaders, '送信シート', true);
    if (emptySend && sendValues.length && JSON.stringify(sendValues[0]) !== JSON.stringify(sendHeaders)) {
      check('info', '送信シートの列', 'データ行はありません。送信データを準備すると、現在の設定に合わせて見出しを更新します。');
    }
    const eh = createHeaderMap_(evalValues[0] || []), qh = createHeaderMap_(sendValues[0] || []);
    const evalKeys = evalValid ? collectResponseKeys_(evalValues, eh) : new Set();
    const sendKeys = sendValid ? collectResponseKeys_(sendValues, qh) : new Set();
    const sendCounts = new Map();
    if (sendValid) sendValues.slice(1).forEach(function(row) {
      const key = makeResponseKey_(row[qh['元SS_ID']], row[qh['元シート名']], row[qh['元行番号']]);
      if (key) sendCounts.set(key, (sendCounts.get(key) || 0) + 1);
    });
    if (sendValid && sendValues.length > 1) {
      const lastNo = Number(sendValues[sendValues.length - 1][qh.No] || 0);
      if (!isNaN(lastNo) && lastNo > 0) number = lastNo + 1;
    }
    const candidate = function(item) {
      if (result.counts.eligible >= 500) { result.truncated = true; return; }
      result.counts.eligible++;
      const fail = function(detail, action) {
        result.counts.blocked++; check('error', item.stage + ' / ' + item.sourceSheet + ' ' + item.sourceRow + '行', detail, action || '生徒一覧・対応表と対象行を確認してください。');
      };
      if (!makeResponseKey_(item.sourceId, item.sourceSheet, item.sourceRow)) { fail('元回答の情報が不正です。'); return; }
      const list = students.get(item.email.trim().toLowerCase()) || [];
      const courseId = item.courseId || mappings.get(makeMappingLookupKey_(item.sourceId, item.sourceSheet));
      const matches = courseId ? list.filter(function(s) { return s.courseId === courseId; }) : list;
      if (matches.length !== 1) { fail(matches.length ? '送信先の生徒・クラスを一意に特定できません。' : '生徒一覧と宛先が一致しません。'); return; }
      const student = matches[0];
      if(student.excluded){fail('このクラスの生徒は名簿で除外されています。');return;}
      if (!student.studentId || !student.courseId || (item.studentId && item.studentId !== student.studentId)) { fail('studentIdまたはコースIDを確認してください。'); return; }
      let name = item.saved ? item.name : (student.name || item.name);
      let body = item.body;
      if (!item.saved) {
        const context = { student_name: name, class_name: student.className, course_id: student.courseId, row_no: String(number++) };
        config.fields.forEach(function(field) { context[field.key] = item.fields[field.key]; });
        body = renderTemplateWithSections_(config.messageTemplate, context, config);
      }
      if (!name || !String(body || '').trim()) { fail('名前または生成される本文が空です。', '本文・差し込み項目と生徒名を確認してください。'); return; }
      result.counts.previewed++;
      if (result.samples.length < 10) result.samples.push({ stage: item.stage, sourceSheet: item.sourceSheet, sourceRow: item.sourceRow,
        email: item.email, name: name, className: item.saved ? item.className : student.className, courseId: student.courseId,
        body: name + 'さんへ(個別メッセージ)\n\n' + String(body).slice(0, 8000) + (String(body).length > 8000 ? '\n（表示を省略）' : '') });
      return true;
    };
    const rowItem = function(row, headers, stage) {
      return { stage: stage, sourceId: row[headers['元SS_ID']], sourceSheet: row[headers['元シート名']], sourceRow: Number(row[headers['元行番号']]),
        email: String(row[headers['メールアドレス']] || '').trim(), name: String(row[headers['名前']] || '').trim(), fields: {} };
    };
    if (sendValid) sendValues.slice(1).forEach(function(row) {
      if (!row.some(function(v) { return String(v).trim(); })) return;
      const status = String(row[qh['送信状態']] || '').trim();
      if (status && status !== '未') { result.counts.skipped++; return; }
      const item = rowItem(row, qh, '送信待ち');
      const key = makeResponseKey_(item.sourceId, item.sourceSheet, item.sourceRow);
      if (sendCounts.get(key) > 1) { result.counts.blocked++; check('error', '送信シート', '同じ元回答の送信行が重複しています。', '投稿の有無を確認し、重複行を整理してください。'); return; }
      item.saved = true; item.body = row[qh[config.replyBodyHeader]]; item.courseId = String(row[qh['コースID']] || '').trim();
      item.studentId = String(row[qh.studentId] || '').trim(); item.className = row[qh['クラス名']];
      if (!item.courseId || !item.studentId) { result.counts.blocked++; check('error', '送信シート', '保存済みの宛先が不完全です。', '送信シートのコースIDとstudentIdを確認してください。'); return; }
      candidate(item);
    });
    const preparedKeys = new Set(sendKeys);
    if (evalValid) evalValues.slice(1).forEach(function(row) {
      const item = rowItem(row, eh, '旧評価データの移行待ち');
      const key = makeResponseKey_(item.sourceId, item.sourceSheet, item.sourceRow);
      if (!row.some(function(v) { return String(v).trim(); }) || row[eh['処理状態']] || preparedKeys.has(key)) { result.counts.skipped++; return; }
      config.fields.forEach(function(field) { item.fields[field.key] = String(row[eh[getFieldEvalHeader_(field)]] || '').trim(); });
      if (candidate(item)) preparedKeys.add(key);
    });
    const gradeKeys = new Map(), unmapped = new Map();
    config.fields.filter(function(field) { return field.type === 'score_grade'; }).forEach(function(field) {
      gradeKeys.set(field.key, new Set(getFieldGradeScale_(field, config).map(function(entry) { return normalizeGradeScaleKey_(entry.from); })));
    });
    const seenSources = new Set();
    config.formSources.forEach(function(source) {
      if (seenSources.has(source.id)) return;
      seenSources.add(source.id); result.counts.sources++;
      try {
        const sourceSs = SpreadsheetApp.openById(source.id);
        const sheets = sourceSs.getSheets().filter(function(sheet) { return !isManagedInternalSheet_(sheet.getName(), source.id) && sheetMatchesPrefixes_(sheet.getName(), config.formSheetNamePrefix); });
        if (!sheets.length) check('error', 'フォームの対象シート', '設定した先頭文字に一致するシートがありません。', '対象シート名の先頭文字を確認してください。');
        sheets.forEach(function(sheet) {
          result.counts.sheets++;
          const values = read(sheet), headers = values.display[0] || [], h = createHeaderMap_(headers);
          const required = [config.emailHeader, config.studentNameHeader, config.formStatusHeader].concat(config.fields.map(function(f) { return f.sourceHeader; }));
          if (!requireHeaders(values.display, required, sheet.getName())) return;
          check('ok', sheet.getName() + 'の接続と列', '参照でき、必要な列が揃っています。');
          values.raw.slice(1).forEach(function(row, index) {
            const rowNo = index + 2, key = makeResponseKey_(source.id, sheet.getName(), rowNo);
            const email = String(row[h[config.emailHeader]] || '').trim();
            const sourceState = String(row[h[config.formStatusHeader]] || '').trim();
            if (!email || (sourceState && sourceState !== '準備×') || String(row[h[config.scoreSourceHeader]] == null ? '' : row[h[config.scoreSourceHeader]]).trim() === '' || evalKeys.has(key) || sendKeys.has(key)) { result.counts.skipped++; return; }
            if (sourceState === '準備×') check('error', sheet.getName() + ' ' + rowNo + '行の準備エラー', '前回の送信データ準備に失敗しています。', '下の診断結果を確認・修正し、この回答の状態列を空欄にして送信データを準備してください。');
            const item = { stage: '送信データ準備待ち', sourceId: source.id, sourceSheet: sheet.getName(), sourceRow: rowNo, email: email, name: String(row[h[config.studentNameHeader]] || '').trim(), fields: {} };
            config.fields.forEach(function(field) {
              const raw = row[h[field.sourceHeader]];
              item.fields[field.key] = processConfiguredFieldValue_(field, raw, values.display[index + 1][h[field.sourceHeader]], config);
              if (field.type === 'score_grade' && String(raw == null ? '' : raw).trim() && !gradeKeys.get(field.key).has(normalizeGradeScaleKey_(raw))) unmapped.set(field.sourceHeader, (unmapped.get(field.sourceHeader) || 0) + 1);
            });
            candidate(item);
          });
        });
      } catch (error) { check('error', 'フォームの接続・読み取り', source.id + ': ' + String(error.message || error), '参照元URL・アクセス権・列の設定を確認してください。'); }
    });
    unmapped.forEach(function(count, header) { check('warning', '評価変換 / ' + header, count + '件は変換表にないため、そのまま本文に使います。', '元の値を使うか、評価変換表に追加するか確認してください。'); });
    if (result.truncated) check('warning', '診断範囲の上限', '対象候補500件・読取データ行5000行までを確認しました。未確認の行があります。', '対象や過去データを整理し、分けて確認してください。');
    if (!result.counts.eligible && !result.counts.blocked) check('warning', '今回の処理対象', '確認できる未処理の採点済み回答・送信待ち行がありません。', '回答・採点済み判定列・処理状態を確認してください。');
    return result;
  });
}

/** 自動処理が共通ロック中、トリガーを操作する前に呼ぶ。 */
function assertAdminReadyForAutomation_() {
  const data = getAdminConsoleDataUnlocked_();
  if (!data.ready) {
    const incomplete = data.progress.filter(function(step) { return step.id !== 'automation' && step.state !== 'complete'; });
    throw new Error('自動実行を開始する準備ができていません。' +
      (data.recoveryRequired || data.configParseError ? '設定の復旧と確認が必要です。' : incomplete.map(function(step) { return step.title; }).join('・') + 'を確認してください。'));
  }
}

function openAdminSheet(key) {
  const definition = ADMIN_SHEETS.find(function(item) { return item.key === key; });
  if (!definition) throw new Error('対象のシート操作が不正です。');
  const ss = getAppSpreadsheet_(), sheet = ss.getSheetByName(definition.label);
  if (!sheet) throw new Error('「' + definition.label + '」がありません。先に該当手順で作成してください。');
  if (SpreadsheetApp.getActiveSpreadsheet()) ss.setActiveSheet(sheet);
  return { url:adminSheetUrl_(ss,sheet),
    message: '「' + definition.label + '」を開きました。シートを編集後、管理画面で再読込してください。' };
}

function adminSheetUrl_(ss, sheet) {
  return 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/edit#gid=' + sheet.getSheetId();
}

function runAdminAction(action, confirmed, expectedRevision) {
  const allowed = ['initialize', 'classes', 'students', 'sortClasses', 'sortStudents', 'mapping', 'import', 'prepare', 'send', 'remind', 'clearEval', 'clearSend', 'clearErrors'];
  if (allowed.indexOf(action) < 0) throw new Error('未対応の操作です。');
  return withAppLock_(function() {
    const config = getConfig_();
    if (expectedRevision !== getConfigRevision_(config)) throw new Error('画面を開いた後に設定が変更されました。再読込して確認してください。');
    if (['send', 'remind', 'clearEval', 'clearSend', 'clearErrors'].indexOf(action) >= 0) {
      if (!confirmed || confirmed.confirmed !== true) throw new Error('操作内容の確認が必要です。');
      if (confirmed.revision !== adminActionRevision_(action, config)) throw new Error('確認対象のデータが変更されました。再読込して件数と内容を確認してください。');
    }
    let message = '';
    switch (action) {
      case 'initialize': initializeSheetsUnlocked_(); message = '必要なシートを準備しました（クラス一覧・生徒一覧・採点テンプレ・対応表・送信シート・エラー・システム管理）。'; break;
      case 'classes': classroomdataUnlocked_(); message = 'クラス一覧を更新しました。シートで同期対象に1を入力してください。'; break;
      case 'students': message = studentdataMultiUnlocked_(); break;
      case 'sortClasses': sortRosterSheet_('クラス一覧'); message='クラス一覧を整列しました。'; break;
      case 'sortStudents': sortRosterSheet_(STUDENT_SHEET_NAME); message='生徒一覧を整列しました。転退学者は除外(1)に1を入力すると再取得後も除外します。'; break;
      case 'mapping': message = createMappingSheetUnlocked_(); break;
      case 'import':
      case 'prepare': prepareSendDataUnlocked_(); message = '送信データの準備が終了しました。送信シートとエラーを確認してください。'; break;
      case 'send': message = sendMessagesUnlocked_(); break;
      case 'remind': remindUngradedAndErrorsUnlocked_(); message = 'リマインダー処理が終了しました。通知対象がある場合のみメールを送信します。'; break;
      default:
        const names = { clearEval: EVAL_SHEET_NAME, clearSend: SEND_SHEET_NAME, clearErrors: ERROR_SHEET_NAME };
        const sheet = getAppSpreadsheet_().getSheetByName(names[action]);
        if (sheet && sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
        message = '「' + names[action] + '」のデータ行をクリアしました。';
    }
    SpreadsheetApp.flush();
    return { message: message, data: getAdminConsoleDataUnlocked_() };
  });
}

function adminUnknownTokens_(config) {
  const known = new Set(['student_name', 'class_name', 'course_id', 'row_no'].concat(config.fields.map(function(field) { return field.key; })));
  const missing = new Set();
  const template = String(config.messageTemplate || '');
  const record = function(_, key) { if (!known.has(key)) missing.add(key); return _; };
  template.replace(friendlyPlaceholderPattern_(), function(match, legacyKey, key) { return record(match, legacyKey || key); })
    .replace(/{{[#/]?([A-Za-z_][A-Za-z0-9_]*)}}/g, record);
  return Array.from(missing).map(function(key) { return '差し込み「' + key + '」に対応する項目がありません。'; });
}

function previewReplyTemplate(payload) {
  validateConfigDraft_(payload);
  const config = normalizeAppConfig_(payload);
  const tokens = [
    { key: 'student_name', label: '生徒の名前', source: '生徒一覧の名前', example: '山田 太郎（例）' },
    { key: 'class_name', label: 'クラス名', source: '送信先のクラス', example: 'サンプル授業' },
    { key: 'course_id', label: 'ClassroomのコースID', source: 'クラス一覧のコースID', example: 'sample-course' },
    { key: 'row_no', label: '送信データの番号', source: '送信シートのNo', example: '1' }
  ];
  config.fields.forEach(function(field) {
    const scale = field.type === 'score_grade' ? getFieldGradeScale_(field, config) : [];
    tokens.push({ key: field.key, label: getFieldSendHeader_(field) || getFieldEvalHeader_(field),
      source: 'フォームの「' + field.sourceHeader + '」列',
      example: field.type === 'score_grade' ? (scale.length ? scale[0].to : '（変換先未設定）') : (field.type === 'date' ? '2026/04/01' : '「' + field.sourceHeader + '」の回答例') });
  });
  const context = {};
  tokens.forEach(function(token) { context[token.key] = token.example; });
  const body = renderTemplateWithSections_(config.messageTemplate, context, config);
  return { text: body ? context.student_name + 'さんへ(個別メッセージ)\n\n' + body : '', tokens: tokens, warnings: adminUnknownTokens_(config) };
}

function exportSettingsProfile(includeConnections) {
  return withAppLock_(function() { return buildSettingsProfile_(getConfig_(), includeConnections); });
}

function buildSettingsProfile_(source, includeConnections) {
  const config = {};
  PROFILE_CONFIG_KEYS.forEach(function(key) { config[key] = source[key]; });
  const separate = includeConnections && typeof includeConnections === 'object';
  if (includeConnections === true) { config.formSources = source.formSources; config.reminderTo = source.reminderTo; }
  if (separate && includeConnections.reminderTo === true) config.reminderTo = source.reminderTo;
  const profile = { application: 'turret', schemaVersion: separate ? 4 : 3, exportedAt: new Date().toISOString(),
    includesConnections: includeConnections === true, config: config, schedule: getAutomationSchedule_() };
  if (separate && includeConnections.templateFormUrl === true) {
    const url = currentTemplateFormUrl_();
    if (!url) throw new Error('保存済みのひな形フォームURLがありません。「フォームを準備」でひな形からフォームを作成してください。');
    profile.templateFormUrl = normalizeTemplateFormUrl_(url);
  }
  return { filename: configurationFilename_('settings'), json: JSON.stringify(profile, null, 2) };
}

/** 各画面で同じ種類・運用名・日本時間のファイル名を使う。 */
function configurationFilename_(kind, date) {
  const ss = getAppSpreadsheet_();
  const rawName = typeof ss.getName === 'function' ? ss.getName() : '運用';
  const name = Array.from(String(rawName || '運用').replace(/[<>:"/\\|?*\u0000-\u001f]/g, '_').replace(/_+/g, '_').replace(/^[.\s]+|[.\s]+$/g, '')).slice(0, 80).join('') || '運用';
  const japan = new Date((date || new Date()).getTime() + 9 * 60 * 60 * 1000).toISOString();
  const timestamp = japan.slice(0, 10).replace(/-/g, '') + '-' + japan.slice(11, 19).replace(/:/g, '');
  return 'turret-' + kind + '-' + name + '-' + timestamp + '.json';
}

function requireConfigurationKind_(kind) {
  if (!gradeConfigurationKind_(kind) && ['settings', 'scoring', 'templates', 'bundle'].indexOf(kind) < 0) throw new Error('設定ファイルの種類を選んでください。');
}

/** ファイル名ではなく既存JSONの識別情報を使い、不明な形式を推測で適用しない。 */
function detectConfigurationKind_(json) {
  if (typeof json !== 'string' || utf8ByteLength_(json) > configurationFileLimit_('bundle')) throw new Error('設定JSONは1MiB以内にしてください。');
  let payload;
  try { payload = JSON.parse(json); } catch (error) { throw new Error('JSONを読み取れません。書き出したファイルを選んでください。'); }
  requireProfileObject_(payload, '設定JSON');
  const unknown = '設定JSONの種類を判定できません。turretから書き出したファイルを選んでください。';
  if (payload.type !== undefined) {
    if(payload.application==='turret' && gradeConfigurationKind_(payload.type))return payload.type;
    if (payload.type === 'configuration-bundle' && payload.application === 'turret' && payload.meta === undefined) return 'bundle';
    throw new Error(unknown);
  }
  if (payload.application !== undefined) {
    if (payload.application === 'turret' && payload.meta === undefined) return 'settings';
    throw new Error(unknown);
  }
  if (payload.meta && payload.meta.type !== undefined) {
    if (payload.meta.type === 'turret-grading-template-pack') return 'templates';
    if (payload.meta.type === 'screen-config') return 'scoring';
    throw new Error(unknown);
  }
  // 識別情報がない旧採点設定だけは、採点固有の必須キーで判定する。
  const config = payload.config || payload;
  if (config && typeof config === 'object' && ['scoreCols', 'displayCols', 'colChecks'].every(function(key) { return Object.prototype.hasOwnProperty.call(config, key); })) return 'scoring';
  throw new Error(unknown);
}

/** 既存の個別JSONの形式を保ち、一括出力だけ3種類を包む。 */
function exportConfigurationFile(kind, includeConnections) {
  assertWebOperator_();
  requireConfigurationKind_(kind);
  if(gradeConfigurationKind_(kind))return withAppLock_(function(){return gradeExportConfiguration_(kind);});
  return withAppLock_(function() {
    let payload;
    if (kind === 'settings') payload = JSON.parse(buildSettingsProfile_(getConfig_(), includeConnections).json);
    else if (kind === 'scoring') payload = scoringLegacyApiExportConfig_();
    else if (kind === 'templates') payload = scoringExportTemplatePack();
    else payload = { application: 'turret', type: 'configuration-bundle', schemaVersion: 1, exportedAt: new Date().toISOString(),
      management: JSON.parse(buildSettingsProfile_(getConfig_(), includeConnections).json),
      scoring: scoringLegacyApiExportConfig_(), templates: scoringExportTemplatePack() };
    const json = JSON.stringify(payload, null, 2);
    if (utf8ByteLength_(json) > configurationFileLimit_(kind)) throw new Error('設定ファイルが大きすぎます。個別の出力、または内容の整理を行ってください。');
    return { filename: configurationFilename_(kind), json: json };
  });
}

function configurationFileLimit_(kind) { return kind === 'settings' ? 131072 : 1048576; }

function parseConfigurationScoring_(payload) {
  requireProfileObject_(payload, '採点設定');
  const copy = JSON.parse(JSON.stringify(payload));
  const config = copy.config || copy;
  requireProfileObject_(config, '採点設定');
  // 回答タブの指定は運用先で選ぶ。旧採点JSONも同じ検証を使う。
  config.sheetName = '';
  scoringValidateConfigInput_(config);
  return scoringParseImportedConfig_(copy);
}

function parseConfigurationTemplates_(payload) {
  requireProfileObject_(payload, '採点テンプレパック');
  profileKeys_(payload, ['meta', 'templates'], '採点テンプレパック');
  if (!payload.meta || payload.meta.type !== 'turret-grading-template-pack' || payload.meta.version !== 1) throw new Error('採点テンプレ形式のパックを指定してください。旧形式は取り込めません。');
  const templates = scoringNormalizeTemplates_(payload.templates);
  const validation = scoringValidateTemplateTable_({ templates: templates });
  if (!validation.ok) throw new Error(validation.issues.filter(function(i) { return i.severity === 'error'; }).map(function(i) { return i.message; }).slice(0, 5).join('\n'));
  return { templates: templates, warnings: validation.issues.filter(function(i) { return i.severity === 'warning'; }).slice(0, 10).map(function(i) { return i.message; }) };
}

function parseConfigurationFile_(kind, json) {
  requireConfigurationKind_(kind);
  if (typeof json !== 'string' || utf8ByteLength_(json) > configurationFileLimit_(kind)) throw new Error(kind === 'settings' ? '管理設定JSONは128KiB以内にしてください。' : '設定JSONは1MiB以内にしてください。');
  let payload;
  try { payload = JSON.parse(json); } catch (error) { throw new Error('JSONを読み取れません。書き出したファイルを選んでください。'); }
  requireProfileObject_(payload, '設定JSON');
  const parsed = { kind: kind, warnings: [] };
  if (kind === 'bundle') {
    profileKeys_(payload, ['application', 'type', 'schemaVersion', 'exportedAt', 'management', 'scoring', 'templates'], '一括設定JSON');
    if (payload.application !== 'turret' || payload.type !== 'configuration-bundle' || payload.schemaVersion !== 1) throw new Error('未対応の一括設定JSONの形式またはバージョンです。');
    parsed.management = parseSettingsProfile_(JSON.stringify(payload.management));
    parsed.scoring = parseConfigurationScoring_(payload.scoring);
    const pack = parseConfigurationTemplates_(payload.templates);
    parsed.templates = pack.templates; parsed.warnings = pack.warnings;
    if (parsed.scoring.ruleTemplateId && !parsed.templates.some(function(t) { return t.name === parsed.scoring.ruleTemplateId; })) throw new Error('採点設定で選択したテンプレートがパッケージにありません: ' + parsed.scoring.ruleTemplateId);
  } else if (kind === 'settings') parsed.management = parseSettingsProfile_(json);
  else if (kind === 'scoring') parsed.scoring = parseConfigurationScoring_(payload);
  else { const pack = parseConfigurationTemplates_(payload); parsed.templates = pack.templates; parsed.warnings = pack.warnings; }
  return parsed;
}

/** 確認したファイルと、読み込み先の関連設定をまとめて照合する。 */
function configurationFileRevision_(kind, json) {
  const value = { kind: kind, json: json, spreadsheetId: getAppSpreadsheet_().getId() };
  if (kind === 'settings' || kind === 'bundle') {
    value.management = getConfigRevision_(getConfig_()); value.schedule = getAutomationSchedule_();
    const payload = JSON.parse(json), profile = kind === 'bundle' ? payload.management : payload;
    if (profile.templateFormUrl) value.templateFormUrl = currentTemplateFormUrl_();
  }
  if (kind === 'scoring' || kind === 'bundle') value.scoring = PropertiesService.getScriptProperties().getProperty('TURRET_SCORING_CONFIG');
  if (kind !== 'settings') value.templates = scoringReadTemplateTable_().revision;
  return scoringDigest_(value);
}

function configurationIncomingTemplates_(parsed, current) {
  if (parsed.kind === 'bundle') return parsed.templates;
  const names = new Set(current.templates.map(function(t) { return t.name; }));
  parsed.templates.forEach(function(t) { if (names.has(t.name)) throw new Error('同名のテンプレートがあります: ' + t.name); });
  const combined = scoringNormalizeTemplates_(current.templates.concat(parsed.templates));
  return combined;
}

function inspectConfigurationFile(kind, json) {
  assertWebOperator_();
  if (kind === 'auto') kind = detectConfigurationKind_(json);
  if(gradeConfigurationKind_(kind))return withAppLock_(function(){return gradeInspectConfiguration_(kind,json);});
  const parsed = parseConfigurationFile_(kind, json);
  return withAppLock_(function() {
    const lines = [], warnings = parsed.warnings.slice();
    if (parsed.management) {
      const p = parsed.management;
      const changed = [];
      if (p.includesConnections) changed.push('フォーム回答スプレッドシートのURL');
      if (p.templateFormUrl) changed.push('ひな形フォームのURL');
      if (p.includesConnections || Object.prototype.hasOwnProperty.call(p.config, 'reminderTo')) changed.push('通知先メール');
      lines.push('管理設定を置き換えます。取り込み項目 ' + p.config.fields.length + '件／本文 ' + p.config.messageTemplate.length + '文字／' + (changed.length ? changed.join('・') + 'も置き換えます。' : 'フォームのURL・通知先メールは保持します。'));
      warnings.push('自動実行を停止してから適用してください。適用しても自動実行は開始しません。');
    }
    if (parsed.scoring) {
      lines.push('採点設定を置き換えます。テンプレート：' + (parsed.scoring.ruleTemplateId || '未選択'));
      warnings.push('採点画面の未保存の設定を先に保存してください。適用後は採点画面を開き直し、列設定を確認してください。');
      if (kind === 'scoring' && parsed.scoring.ruleTemplateId && !scoringReadTemplateTable_().templates.some(function(t) { return t.name === parsed.scoring.ruleTemplateId; })) warnings.push('選択したテンプレートは読み込み先にありません。テンプレパックも読み込んでください。');
    }
    if (parsed.templates) {
      configurationIncomingTemplates_(parsed, scoringReadTemplateTable_());
      lines.push('採点テンプレート ' + parsed.templates.length + '件を' + (kind === 'bundle' ? '適用し、既存のテンプレートをすべて置き換えます。' : '追加します。同名のテンプレートは追加できません。'));
      warnings.push('テンプレ編集画面の未保存の内容を先に保存してください。適用後はシートから再読込してください。');
    }
    return { kind: kind, summary: lines.join('\n'), warnings: warnings, revision: configurationFileRevision_(kind, json) };
  });
}

/** 復元も一括書込みし、先に既存テンプレートを消去しない。 */
function restoreConfigurationTemplates_(ss, snapshot, revision) {
  const sheet = ss.getSheetByName('採点テンプレ');
  if (!snapshot.exists) { if (sheet) ss.deleteSheet(sheet); }
  else {
    const target = sheet || ss.insertSheet('採点テンプレ');
    const height = Math.max(snapshot.values.length, target.getLastRow(), 1), width = Math.max(snapshot.values[0] ? snapshot.values[0].length : 0, target.getLastColumn(), 16);
    // 既存シートから退避した数式だけは数式として復元する。JSON由来の文字列とは区別する。
    const rows = Array.from({ length: height }, function(_, r) { return Array.from({ length: width }, function(_, c) {
      if (snapshot.formulas[r] && snapshot.formulas[r][c]) return snapshot.formulas[r][c];
      return snapshot.values[r] && snapshot.values[r][c] != null ? scoringSheetLiteral_(snapshot.values[r][c]) : '';
    }); });
    target.getRange(1, 1, height, width).setValues(rows);
  }
  if (scoringReadTemplateTable_().revision !== revision) throw new Error('採点テンプレートの復元値が一致しません。');
}

function applyConfigurationFile(kind, json, expectedRevision, choices) {
  assertWebOperator_();
  if(gradeConfigurationKind_(kind))return gradeApplyConfiguration_(kind,json,expectedRevision,choices);
  const parsed = parseConfigurationFile_(kind, json);
  return withAppLock_(function() {
    if (typeof expectedRevision !== 'string' || expectedRevision !== configurationFileRevision_(kind, json)) throw new Error('確認後にファイルまたは設定が変更されました。再読み込みしてから適用してください。');
    const ss = getAppSpreadsheet_(), props = PropertiesService.getScriptProperties();
    let management, oldSchedule, currentTemplates, templates, snapshot;
    if (parsed.management) {
      const automation = getAutomationStateUnlocked_(), ownership = getAutomationSummary_();
      if (ownership.hasManagedOwner && !ownership.managedByCurrentUser) throw new Error('別の管理者が自動実行を管理しています。同じ管理者アカウントで設定を適用してください。');
      if (automation.enabled || automation.legacyCount > 0) throw new Error('設定JSONを適用する前に、自動実行を停止してください。');
      management = normalizeAppConfig_(Object.assign({}, getConfig_(), parsed.management.config));
      validateConfigDraft_(management); validateConfigSheetLayouts_(management);
      if (utf8ByteLength_(JSON.stringify(management)) > 8000) throw new Error('管理設定はUTF-8で8,000バイト以内にしてください。');
      oldSchedule = getAutomationSchedule_();
    }
    if (parsed.templates) {
      currentTemplates = scoringReadTemplateTable_(); templates = configurationIncomingTemplates_(parsed, currentTemplates);
      const validation = scoringValidateTemplateTable_({ templates: templates });
      if (!validation.ok) throw new Error('採点テンプレートに不正な条件があります。');
      snapshot = snapshotSheetContents_(ss, '採点テンプレ');
      const templateSheet = ss.getSheetByName('採点テンプレ');
      snapshot.formulas = templateSheet ? templateSheet.getDataRange().getFormulas() : [];
    }
    const keys = [];
    if (management) keys.push('APP_CONFIG', 'APP_CONFIG_BACKUP', 'APP_CONFIG_STORAGE', 'APP_CONFIG_PARSE_ERROR', 'APP_CONFIG_SAVE_ERROR', AUTOMATION_STATE_KEY_);
    if (parsed.management && parsed.management.templateFormUrl) keys.push(TEMPLATE_FORM_URL_KEY_);
    if (parsed.scoring) keys.push('TURRET_SCORING_CONFIG', SCORING_CONFIG_BACKUP_KEY);
    const previous = keys.map(function(key) { return [key, props.getProperty(key)]; });
    let templatesAttempted = false, scheduleAttempted = false;
    try {
      if (management) {
        scheduleAttempted = true; saveAutomationScheduleDraft_(parsed.management.schedule); saveConfig_(management);
        if (parsed.management.templateFormUrl) {
          const savedUrl = savedTemplateFormUrl_(parsed.management.templateFormUrl);
          props.setProperty(TEMPLATE_FORM_URL_KEY_, savedUrl);
          if (props.getProperty(TEMPLATE_FORM_URL_KEY_) !== savedUrl) throw new Error('ひな形フォームURLの保存を確認できません。');
        }
      }
      if (parsed.scoring) {
        scoringLegacyBackupCurrentConfig_();
        const serialized = JSON.stringify(parsed.scoring); props.setProperty('TURRET_SCORING_CONFIG', serialized);
        if (props.getProperty('TURRET_SCORING_CONFIG') !== serialized) throw new Error('採点設定の保存を確認できません。');
      }
      if (templates) scoringWriteTemplateTable_(templates, currentTemplates.revision, function() { templatesAttempted = true; });
    } catch (error) {
      const failures = [];
      if (templatesAttempted) { try { restoreConfigurationTemplates_(ss, snapshot, currentTemplates.revision); } catch (restoreError) { failures.push(String(restoreError)); } }
      if (scheduleAttempted) { try { saveAutomationScheduleDraft_(oldSchedule); } catch (restoreError) { failures.push(String(restoreError)); } }
      previous.forEach(function(entry) {
        try { if (entry[1] === null) props.deleteProperty(entry[0]); else props.setProperty(entry[0], entry[1]); if (props.getProperty(entry[0]) !== entry[1]) throw new Error('復元値が一致しません'); }
        catch (restoreError) { failures.push(entry[0] + ': ' + restoreError); }
      });
      if (failures.length) {
        try { props.setProperty('APP_CONFIG_SAVE_ERROR', failures.join('\n')); } catch (markerError) { Logger.log(markerError); }
        throw new Error('設定JSONの適用に失敗し、復旧も完了していません。自動実行を停止し、設定JSONと内部設定を確認してください。' + error + '\n' + failures.join('\n'));
      }
      throw error;
    }
    return { message: '設定JSONを適用しました。' + (management ? '自動実行は開始していません。' : '') + '各画面で設定を再読み込みしてください。' };
  });
}

function requireProfileObject_(value, label) {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error(label + 'はオブジェクトで指定してください。');
}

function profileKeys_(object, keys, label) {
  Object.keys(object).forEach(function(key) {
    if (keys.indexOf(key) < 0) throw new Error(label + 'に未対応の項目があります: ' + key);
  });
}

function parseSettingsProfile_(json) {
  if (typeof json !== 'string' || utf8ByteLength_(json) > 131072) throw new Error('設定JSONは128KB以内のテキストで指定してください。');
  let profile;
  try { profile = JSON.parse(json); } catch (error) { throw new Error('JSONを読み取れません。設定書き出しで作成したファイルを選んでください。'); }
  requireProfileObject_(profile, '設定JSON');
  profileKeys_(profile, ['application', 'schemaVersion', 'exportedAt', 'includesConnections', 'config', 'schedule'].concat(profile.schemaVersion === 4 ? ['templateFormUrl'] : []), '設定JSON');
  if (profile.application !== 'turret' || [1, 2, 3, 4].indexOf(profile.schemaVersion) < 0) throw new Error('未対応の設定ファイル、またはバージョンです。');
  if (typeof profile.includesConnections !== 'boolean') throw new Error('接続先を含む設定か指定されていません。');
  if (profile.schemaVersion === 4 && profile.includesConnections) throw new Error('新しい設定JSONでは接続先を個別に選んでください。');
  if (profile.templateFormUrl !== undefined) profile.templateFormUrl = normalizeTemplateFormUrl_(profile.templateFormUrl);
  requireProfileObject_(profile.config, '設定');
  if (profile.schemaVersion === 1 && !Object.prototype.hasOwnProperty.call(profile.config, 'gradeScale')) {
    profile.config.gradeScale = Array.isArray(profile.config.fields) && profile.config.fields.some(function(field) { return field && field.type === 'score_grade'; }) ? buildLegacyDefaultConfig_().gradeScale : [];
  }
  profileKeys_(profile.config, PROFILE_CONFIG_KEYS.concat(profile.includesConnections ? ['formSources', 'reminderTo'] : profile.schemaVersion === 4 ? ['reminderTo'] : []), '設定');
  PROFILE_CONFIG_KEYS.forEach(function(key) {
    if (!Object.prototype.hasOwnProperty.call(profile.config, key)) throw new Error('設定の項目がありません: ' + key);
  });
  if (!Array.isArray(profile.config.fields)) throw new Error('取り込み項目fieldsは配列で指定してください。');
  profile.config.fields.forEach(function(field) {
    requireProfileObject_(field, '取り込み項目');
    profileKeys_(field, ['key', 'sourceHeader', 'evalHeader', 'sendHeader', 'type'].concat(profile.schemaVersion >= 3 ? ['gradeScale'] : []), '取り込み項目');
  });
  if (profile.includesConnections && (!Array.isArray(profile.config.formSources) || !Array.isArray(profile.config.reminderTo))) {
    throw new Error('フォームURLと通知先は配列で指定してください。');
  }
  if (profile.schemaVersion === 4 && Object.prototype.hasOwnProperty.call(profile.config, 'reminderTo') && !Array.isArray(profile.config.reminderTo)) {
    throw new Error('通知先メールは配列で指定してください。');
  }
  requireProfileObject_(profile.schedule, '自動実行時間');
  profileKeys_(profile.schedule, ['importHour', 'deliveryHour', 'reminderHour', 'reminderEnabled'], '自動実行時間');
  ['importHour', 'deliveryHour', 'reminderHour'].forEach(function(key) {
    if (!Number.isInteger(profile.schedule[key]) || profile.schedule[key] < 0 || profile.schedule[key] > 23) throw new Error('自動実行の時間帯は0〜23時で指定してください。');
  });
  if (typeof profile.schedule.reminderEnabled !== 'boolean') throw new Error('リマインダーの有効設定が不正です。');
  if (profile.schedule.importHour >= profile.schedule.deliveryHour || (profile.schedule.reminderEnabled && profile.schedule.deliveryHour >= profile.schedule.reminderHour)) {
    throw new Error('時間帯は送信データ準備→送信→リマインダーの順に分けてください。');
  }
  validateConfigDraft_(profile.config);
  return profile;
}

function inspectSettingsProfile(json) {
  const profile = parseSettingsProfile_(json);
  const changed = [];
  if (profile.includesConnections) changed.push('フォーム回答スプレッドシートのURL');
  if (profile.templateFormUrl) changed.push('ひな形フォームのURL');
  if (profile.includesConnections || Object.prototype.hasOwnProperty.call(profile.config, 'reminderTo')) changed.push('通知先メール');
  return { summary: '取り込み項目 ' + profile.config.fields.length + '件／本文 ' + profile.config.messageTemplate.length + '文字／' +
    (changed.length ? changed.join('・') + 'も置き換えます' : 'フォームのURL・通知先メールは保持します'),
    config: profile.config, schedule: profile.schedule,
    warnings: ['適用しても自動実行は開始しません。実行中の自動処理がある場合は、先に停止してください。'] };
}

function applySettingsProfile(json, expectedRevision) {
  const profile = parseSettingsProfile_(json);
  return withAppLock_(function() {
    const current = getConfig_();
    if (expectedRevision !== getConfigRevision_(current)) throw new Error('設定が変更されました。再読込してから適用してください。');
    const automation = getAutomationStateUnlocked_();
    const ownership = getAutomationSummary_();
    if (ownership.hasManagedOwner && !ownership.managedByCurrentUser) throw new Error('別の管理者が自動実行を管理しています。同じ管理者アカウントで設定を適用してください。');
    if (automation.enabled || automation.legacyCount > 0) {
      throw new Error('設定JSONを適用する前に、自動実行を停止してください。');
    }
    const config = Object.assign({}, current, profile.config);
    // JSONの通知先を取り込む場合、通知しない選択はこの運用で改めて確認する。
    if (Object.prototype.hasOwnProperty.call(profile.config, 'reminderTo')) config.reminderOptOutConfirmed = false;
    const oldSchedule = getAutomationSchedule_();
    const props = PropertiesService.getScriptProperties();
    const oldTemplateUrl = profile.templateFormUrl ? props.getProperty(TEMPLATE_FORM_URL_KEY_) : null;
    saveAutomationScheduleDraft_(profile.schedule);
    try {
      if (profile.templateFormUrl) {
        const savedUrl = savedTemplateFormUrl_(profile.templateFormUrl);
        props.setProperty(TEMPLATE_FORM_URL_KEY_, savedUrl);
        if (props.getProperty(TEMPLATE_FORM_URL_KEY_) !== savedUrl) throw new Error('ひな形フォームURLの保存を確認できません。');
      }
      applyConfigDraft_(config, expectedRevision);
    } catch (error) {
      if (profile.templateFormUrl) {
        try { if (oldTemplateUrl === null) props.deleteProperty(TEMPLATE_FORM_URL_KEY_); else props.setProperty(TEMPLATE_FORM_URL_KEY_, oldTemplateUrl); }
        catch (restoreError) { throw new Error(String(error.message || error) + '。ひな形フォームURLの復元にも失敗しました。設定を確認してください。'); }
      }
      try { saveAutomationScheduleDraft_(oldSchedule); } catch (restoreError) {
        throw new Error(String(error.message || error) + '。希望時刻の復元にも失敗しました。再読込して自動実行設定を確認してください。');
      }
      throw error;
    }
    return { message: '設定JSONを適用しました。自動実行は開始していません。各手順の設定を確認してください。', data: getAdminConsoleDataUnlocked_() };
  });
}

// === 予約処理（旧 Automation.gs） ===
/** 日次自動処理。古いトリガーや孤立したトリガーが動かないよう、トリガー ID をまとめて確定する。 */
const AUTOMATION_STATE_KEY_ = 'APP_AUTOMATION_STATE';
const AUTOMATION_OWNER_KEY_ = 'APP_AUTOMATION_OWNER';
const AUTOMATION_RUNS_KEY_ = 'APP_AUTOMATION_RUNS';
const AUTOMATION_TIMEZONE_ = 'Asia/Tokyo';
const AUTOMATION_HANDLERS_ = {
  import: 'managedAutomationImport_',
  delivery: 'managedAutomationDelivery_',
  reminder: 'managedAutomationReminder_'
};
const AUTOMATION_LEGACY_HANDLERS_ = ['prepareSendData', 'importFromFormsToEval', 'evalToSendSheet', 'sendMessages', 'remindUngradedAndErrors'];

/** 回答締切は返却処理から独立させる。共通の時刻トリガーを1つ使い、フォームごとのトリガーは作らない。 */
const FORM_SCHEDULE_KEY_ = 'TURRET_FORM_SCHEDULE_TRIGGER';
const FORM_SCHEDULE_HANDLER_ = 'managedFormScheduleTick_';

function readManagedFormSchedule_() {
  const raw=PropertiesService.getScriptProperties().getProperty(FORM_SCHEDULE_KEY_);
  if(!raw)return null;
  let record;try{record=JSON.parse(raw);}catch(error){throw new Error('受付終了タイマーの登録情報を読めません。');}
  if(!record||!record.id||!record.owner||!record.spreadsheetId||!record.scriptId)throw new Error('受付終了タイマーの登録情報が不正です。');
  return record;
}

function assertManagedFormScheduleOwner_() {
  const record=readManagedFormSchedule_();
  if(record&&(record.spreadsheetId!==getAppSpreadsheet_().getId()||record.scriptId!==ScriptApp.getScriptId()))throw new Error('別の運用の受付終了タイマーです。');
  return assertAutomationOwner_(record||{owner:''});
}

/** 呼び出し元はアプリ共通ロックを保持する。確定済みの ID だけ実行を許可する。 */
function ensureManagedFormScheduleTrigger_() {
  const owner=assertManagedFormScheduleOwner_(),old=readManagedFormSchedule_();
  const triggers=ScriptApp.getProjectTriggers().filter(function(t){return t.getHandlerFunction()===FORM_SCHEDULE_HANDLER_&&t.getEventType()===ScriptApp.EventType.CLOCK;});
  if(old&&triggers.some(function(t){return String(t.getUniqueId())===old.id;}))return old;
  if(triggers.length)throw new Error('受付終了タイマーの作成結果が未確認です。Apps Scriptの未登録トリガーを確認してから修復してください。');
  const trigger=ScriptApp.newTrigger(FORM_SCHEDULE_HANDLER_).timeBased().everyMinutes(5).create();
  const record={id:String(trigger.getUniqueId()),owner:owner,spreadsheetId:getAppSpreadsheet_().getId(),scriptId:ScriptApp.getScriptId()};
  const props=PropertiesService.getScriptProperties(),json=JSON.stringify(record);
  try{props.setProperty(FORM_SCHEDULE_KEY_,json);}catch(error){
    // 応答が失敗しても登録は完了している可能性がある。結果を推測で決めない。
    let saved;try{saved=props.getProperty(FORM_SCHEDULE_KEY_);}catch(readError){throw new Error('受付終了タイマーの保存結果が未確認です。状態を更新してください。');}
    if(saved!==json){try{ScriptApp.deleteTrigger(trigger);}catch(ignored){}throw error;}
  }
  return record;
}

function repairFormScheduleTimer() { return withAppLock_(function(){ensureManagedFormScheduleTrigger_();return getManagedFormScheduleSummary_();}); }

function getManagedFormScheduleSummary_() {
  try{
    const record=readManagedFormSchedule_();
    if(!record)return {registered:false,healthy:false,message:'受付終了日時の設定時に専用タイマーを作成します。'};
    const owner=PropertiesService.getUserProperties().getProperty(AUTOMATION_OWNER_KEY_);
    if(record.owner!==owner||record.spreadsheetId!==getAppSpreadsheet_().getId()||record.scriptId!==ScriptApp.getScriptId())return {registered:true,healthy:false,message:'受付終了タイマーを設定した管理者・運用で確認してください。'};
    const healthy=ScriptApp.getProjectTriggers().some(function(t){return String(t.getUniqueId())===record.id&&t.getHandlerFunction()===FORM_SCHEDULE_HANDLER_;});
    return {registered:true,healthy:healthy,message:healthy?'受付終了を約5分間隔で確認します。実行時刻は遅れる場合があります。':'受付終了タイマーが見つかりません。「タイマーを修復」を実行してください。'};
  }catch(error){return {registered:true,healthy:false,message:String(error.message||error)};}
}

function managedFormScheduleTick_(event) {
  return withAppLock_(function(){
    const registration=readManagedFormSchedule_();
    if(!registration||!event||String(event.triggerUid)!==registration.id||registration.owner!==PropertiesService.getUserProperties().getProperty(AUTOMATION_OWNER_KEY_)||
      registration.spreadsheetId!==getAppSpreadsheet_().getId()||registration.scriptId!==ScriptApp.getScriptId())return;
    const started=Date.now(),records=getManagedRecords_(),cache={records:records,sheet:null};
    const due=records.filter(function(r){return r.kind==='form'&&r.ownerSpreadsheetId===registration.spreadsheetId&&r.ownerScriptId===registration.scriptId&&managedCloseIsDue_(r)&&r.closeState!=='closed';});
    due.sort(function(a,b){return String(a.closeCheckedAt||'').localeCompare(String(b.closeCheckedAt||''));});
    due.forEach(function(r){
      if(Date.now()-started>200000)return;
      try{
        const form=FormApp.openById(r.formId);
        if(form.isAcceptingResponses())form.setAcceptingResponses(false);
        r.accepting=form.isAcceptingResponses();
        if(r.accepting)throw new Error('フォームが受付中のままです。');
        r.closeState='closed';r.closedAt=new Date().toISOString();r.closeError='';
        if(r.formUpdate)r.formUpdate.accepting=false;
        if(r.settingsUpdate)r.settingsUpdate.accepting=false;
      }catch(error){r.closeState='error';r.closeError=String(error.message||error).slice(0,500);}
      r.closeCheckedAt=new Date().toISOString();
      try{saveManagedRecord_(r,cache);}catch(error){Logger.log('受付終了の結果記録に失敗しました。次回に再確認します。');}
    });
  });
}

/** 日次返却とは別の処理で、メッセージは送信しない。呼び出し元はアプリ共通ロックを保持する。 */
function ensureManagedResponseTrigger_() {
  const props=PropertiesService.getScriptProperties(),raw=props.getProperty('TURRET_RESPONSE_TRIGGER');
  const record=raw?JSON.parse(raw):null,ss=getAppSpreadsheet_();
  const owner=assertAutomationOwner_(record||{owner:''});
  if(record && record.spreadsheetId!==ss.getId())throw new Error('別の運用の回答トリガーが登録されています。');
  const triggers=ScriptApp.getProjectTriggers().filter(function(t){return t.getHandlerFunction()==='managedFormResponseReceived_';});
  if(record && triggers.some(function(t){return String(t.getUniqueId())===record.id;}))return record;
  if(triggers.length)throw new Error('回答トリガーの作成結果を確認してください。Apps Scriptで残った未登録トリガーを確認・削除してから再実行してください。');
  const trigger=ScriptApp.newTrigger('managedFormResponseReceived_').forSpreadsheet(ss).onFormSubmit().create();
  const next={id:String(trigger.getUniqueId()),owner:owner,spreadsheetId:ss.getId()};
  props.setProperty('TURRET_RESPONSE_TRIGGER',JSON.stringify(next));return next;
}

function managedFormResponseReceived_(event) {
  return withAppLock_(function(){
    const raw=PropertiesService.getScriptProperties().getProperty('TURRET_RESPONSE_TRIGGER');
    if(!raw||!event||!event.triggerUid||!event.range||!event.source)return;
    const registration=JSON.parse(raw), ss=getAppSpreadsheet_();
    if(registration.id!==String(event.triggerUid)||registration.owner!==PropertiesService.getUserProperties().getProperty(AUTOMATION_OWNER_KEY_)||registration.spreadsheetId!==ss.getId()||event.source.getId()!==ss.getId())return;
    const sheet=event.range.getSheet(),start=event.range.getRow();
    if(start<2 || event.range.getNumRows()!==1 || sheet.getParent().getId()!==ss.getId())return;
    const record=getManagedRecords_().find(function(r){return r.kind==='form'&&String(r.responseSheetId)===String(sheet.getSheetId());});
    if(!record)return;
    assertManagedOwner_(record);
    if(record.formUpdate)return;
    const result=fillManagedNames_(record,sheet,start,1);
    if(result.unmatched)ensureErrorSheet_().appendRow([new Date(),sheet.getName(),start,'','','NAME_UNMATCHED','Classroom生徒一覧に照合できません。名簿を更新して採点画面の「名前を反映」を実行してください。']);
  });
}

function automationRecipientsValid_(config) {
  return Array.isArray(config.reminderTo) && config.reminderTo.length > 0 && config.reminderTo.every(function(address) {
    return typeof address === 'string' && /^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(address);
  });
}

function validateAutomationSchedule_(schedule) {
  if (!schedule || typeof schedule !== 'object' || Array.isArray(schedule)) throw new Error('自動実行の時刻設定が不正です。');
  ['importHour', 'deliveryHour', 'reminderHour'].forEach(function(key) {
    if (typeof schedule[key] !== 'number' || !Number.isInteger(schedule[key]) || schedule[key] < 0 || schedule[key] > 23) {
      throw new Error('時刻は0〜23の整数で入力してください。');
    }
  });
  if (typeof schedule.reminderEnabled !== 'boolean') throw new Error('リマインダーの有効・無効を選択してください。');
  if (schedule.importHour >= schedule.deliveryHour || (schedule.reminderEnabled && schedule.deliveryHour >= schedule.reminderHour)) {
    throw new Error('時間帯は送信データ準備、送信、リマインダーの順に重ならないよう設定してください。');
  }
  return { importHour: schedule.importHour, deliveryHour: schedule.deliveryHour, reminderHour: schedule.reminderHour, reminderEnabled: schedule.reminderEnabled };
}

function readAutomationRecord_() {
  const raw = PropertiesService.getScriptProperties().getProperty(AUTOMATION_STATE_KEY_);
  if (!raw) return {
    revision: '0', owner: '', enabled: false, active: [], activeSchedule: null,
    // アプリ設定を読めない場合でも、復旧と停止の操作を可能にする。
    schedule: { importHour: 5, deliveryHour: 8, reminderHour: 16, reminderEnabled: false }
  };
  let record;
  try { record = JSON.parse(raw); } catch (error) { throw new Error('自動実行の保存情報を読めません。管理者に確認してください。'); }
  if (!record || typeof record.revision !== 'string' || typeof record.owner !== 'string' || typeof record.enabled !== 'boolean' || !Array.isArray(record.active)) {
    throw new Error('自動実行の保存情報が不正です。管理者に確認してください。');
  }
  record.schedule = validateAutomationSchedule_(record.schedule);
  return record;
}

function getAutomationSchedule_() {
  return readAutomationRecord_().schedule;
}

function getAutomationSummary_() {
  const record = readAutomationRecord_();
  const token = PropertiesService.getUserProperties().getProperty(AUTOMATION_OWNER_KEY_);
  return { schedule: record.schedule, enabled: record.enabled, hasManagedOwner: !!record.owner, managedByCurrentUser: !!record.owner && record.owner === token };
}

/** 呼び出し元はアプリ共通ロックを保持する。この関数は ScriptApp にアクセスしない。 */
function saveAutomationScheduleDraft_(schedule) {
  const desired = validateAutomationSchedule_(schedule);
  const record = readAutomationRecord_();
  record.schedule = desired;
  writeAutomationRecord_(record);
  return desired;
}

function writeAutomationRecord_(record) {
  record.revision = Utilities.getUuid();
  const serialized = JSON.stringify(record);
  const properties = PropertiesService.getScriptProperties();
  try { properties.setProperty(AUTOMATION_STATE_KEY_, serialized); }
  catch (error) {
    // サービスが保存後に失敗を返す場合がある。巻き戻すか判断する前に読み戻す。
    let current;
    try { current = properties.getProperty(AUTOMATION_STATE_KEY_); }
    catch (readError) {
      const uncertain = new Error('自動実行の保存結果を確認できません。画面を再読込して状態を確認してください。');
      uncertain.commitUncertain = true;
      throw uncertain;
    }
    if (current !== serialized) throw new Error('自動実行の設定を保存できませんでした。画面を再読込して再実行してください。');
  }
}

function assertAutomationRevision_(record, expectedRevision) {
  if (expectedRevision !== record.revision) throw new Error('自動実行の設定が別の操作で変更されました。画面を再読込してください。');
}

function assertAutomationOwner_(record) {
  let token = PropertiesService.getUserProperties().getProperty(AUTOMATION_OWNER_KEY_);
  if (record.owner && token !== record.owner) throw new Error('別の管理者が自動実行を管理しています。同じ管理者アカウントで操作してください。');
  if (!token) {
    token = Utilities.getUuid();
    PropertiesService.getUserProperties().setProperty(AUTOMATION_OWNER_KEY_, token);
  }
  return token;
}

function listAutomationTriggers_() {
  return ScriptApp.getProjectTriggers().filter(function(trigger) {
    return trigger.getEventType() === ScriptApp.EventType.CLOCK;
  }).map(function(trigger) {
    const handler = trigger.getHandlerFunction();
    const kind = Object.keys(AUTOMATION_HANDLERS_).find(function(key) { return AUTOMATION_HANDLERS_[key] === handler; });
    return { trigger: trigger, id: String(trigger.getUniqueId()), handler: handler, kind: kind || 'legacy', managed: !!kind, legacy: AUTOMATION_LEGACY_HANDLERS_.indexOf(handler) >= 0 };
  }).filter(function(row) { return row.managed || row.legacy; });
}

function getAutomationState() {
  return withAppLock_(getAutomationStateUnlocked_);
}

function getAutomationStateUnlocked_() {
  const record = readAutomationRecord_();
  const token = PropertiesService.getUserProperties().getProperty(AUTOMATION_OWNER_KEY_);
  const owned = record.owner && record.owner === token;
  const rows = listAutomationTriggers_();
  const warnings = [
    '実行時刻はAsia/Tokyoの指定した時台です。分単位の定刻は保証されません。',
    '他のアカウントが作成したトリガーは確認できません。管理者を一人に決め、他のアカウントの旧トリガーは各作成者が停止してください。'
  ];
  if (record.owner && !owned) warnings.push('別の管理者の自動実行設定です。開始・変更・停止はその管理者アカウントで行ってください。');
  if (record.enabled && JSON.stringify(record.schedule) !== JSON.stringify(record.activeSchedule)) warnings.push('保存した時間帯と稼働中の時間帯が異なります。「開始・変更」で反映してください。');
  const triggers = rows.map(function(row) {
    const active = !!owned && record.enabled && record.active.some(function(item) { return item.id === row.id && item.kind === row.kind; });
    return { id: row.id, handler: row.handler, kind: row.kind, managed: row.managed, legacy: row.legacy, active: active };
  });
  const legacyCount = rows.filter(function(row) { return row.legacy; }).length;
  if (legacyCount) warnings.push('旧処理のトリガーが' + legacyCount + '件あります。切り替える場合は旧トリガーの置換を確認してください。');
  if (triggers.some(function(row) { return row.managed && !row.active; })) warnings.push('無効な管理用トリガーが残っています。実行は抑止されています。開始・変更または停止で削除を再試行できます。');
  if (owned && record.enabled && record.active.some(function(item) { return !triggers.some(function(row) { return row.active && row.id === item.id; }); })) {
    warnings.push('稼働予定のトリガーが見つかりません。「開始・変更」で作り直してください。');
  }
  let lastRuns = [];
  try { lastRuns = JSON.parse(PropertiesService.getScriptProperties().getProperty(AUTOMATION_RUNS_KEY_) || '[]'); } catch (error) { warnings.push('最近の実行結果を読み取れませんでした。'); }
  return { schedule: record.schedule, timezone: AUTOMATION_TIMEZONE_, enabled: record.enabled, legacyCount: legacyCount,
    hasManagedOwner: !!record.owner, managedByCurrentUser: !!owned,
    triggers: triggers, warnings: warnings, revision: record.revision, lastRuns: Array.isArray(lastRuns) ? lastRuns : [] };
}

function cleanupAutomationTriggers_(rows) {
  const failed = [];
  rows.forEach(function(row) {
    try { ScriptApp.deleteTrigger(row.trigger || row); } catch (error) { failed.push(row); }
  });
  return failed;
}

function configureAutomation(schedule, expectedRevision, replaceLegacy, expectedConfigRevision) {
  return withAppLock_(function() {
    const desired = validateAutomationSchedule_(schedule);
    const record = readAutomationRecord_();
    assertAutomationRevision_(record, expectedRevision);
    if (expectedConfigRevision !== undefined && expectedConfigRevision !== getConfigRevision_(getConfig_())) {
      throw new Error('画面を開いた後に送信設定が変更されました。再読込して内容を確認してください。');
    }
    assertAdminReadyForAutomation_();
    const owner = assertAutomationOwner_(record);
    const config = getConfig_();
    validateAppConfig_(config);
    if (desired.reminderEnabled && !automationRecipientsValid_(config)) throw new Error('リマインダーの通知先を正しいメールアドレスで設定するか、リマインダーを無効にしてください。');
    const existing = listAutomationTriggers_();
    const legacy = existing.filter(function(row) { return row.legacy; });
    if (legacy.length && replaceLegacy !== true) throw new Error('旧トリガーの置換確認が必要です。対象を確認してから切り替えてください。');
    const created = [];
    try {
      ['import', 'delivery'].concat(desired.reminderEnabled ? ['reminder'] : []).forEach(function(kind) {
        const trigger = ScriptApp.newTrigger(AUTOMATION_HANDLERS_[kind]).timeBased().atHour(desired[kind + 'Hour']).everyDays(1).inTimezone(AUTOMATION_TIMEZONE_).create();
        created.push({ trigger: trigger, id: String(trigger.getUniqueId()), kind: kind });
      });
    } catch (error) {
      cleanupAutomationTriggers_(created);
      throw new Error('自動実行トリガーの作成に失敗しました。新しいトリガーは有効化していません。画面を再読込してください。');
    }
    // 旧トリガーには実行制限がないため、新しい世代を有効にする前にすべて削除する。
    if (cleanupAutomationTriggers_(legacy).length) {
      cleanupAutomationTriggers_(created);
      throw new Error('旧トリガーを一部削除できませんでした。新しい自動実行は有効化していません。削除済みの旧時間帯は復元できないため、残った旧トリガーを確認して切り替えを再実行してください。');
    }
    record.schedule = desired;
    record.activeSchedule = desired;
    record.owner = owner;
    record.enabled = true;
    record.active = created.map(function(row) { return { id: row.id, kind: row.kind }; });
    try { writeAutomationRecord_(record); }
    catch (error) {
      if (!error.commitUncertain) cleanupAutomationTriggers_(created);
      if (legacy.length) throw new Error(error.message + ' 旧トリガーは削除済みで時間帯を復元できません。状態を確認して開始を再実行してください。');
      throw error;
    }
    // ここで失敗しても、確定済みの世代しかイベントの許可リストを通れない。
    cleanupAutomationTriggers_(existing.filter(function(row) { return row.managed; }));
    return getAutomationStateUnlocked_();
  });
}

function stopAutomation(expectedRevision, includeLegacy) {
  return withAppLock_(function() {
    const record = readAutomationRecord_();
    assertAutomationRevision_(record, expectedRevision);
    assertAutomationOwner_(record);
    const existing = listAutomationTriggers_();
    record.enabled = false;
    record.active = [];
    record.activeSchedule = null;
    writeAutomationRecord_(record);
    const failed = cleanupAutomationTriggers_(existing.filter(function(row) { return row.managed || (includeLegacy === true && row.legacy); }));
    const state = getAutomationStateUnlocked_();
    if (failed.some(function(row) { return row.legacy; })) state.warnings.push('旧トリガーを一部停止できませんでした。旧処理が動く可能性があります。残った旧トリガーを停止してください。');
    return state;
  });
}

function automationErrorRowCount_() {
  try {
    const sheet = getAppSpreadsheet_().getSheetByName('エラー');
    return sheet ? Math.max(0, sheet.getLastRow() - 1) : 0;
  } catch (error) { return null; }
}

function recordAutomationRun_(kind, status, startedAt, deliveryResult, errorDelta) {
  try {
    const properties = PropertiesService.getScriptProperties();
    let runs;
    try { runs = JSON.parse(properties.getProperty(AUTOMATION_RUNS_KEY_) || '[]'); } catch (error) { runs = []; }
    if (!Array.isArray(runs)) runs = [];
    let message = status === 'success' ? '処理が完了しました。処理対象の状態とエラーシートも確認してください。' : '処理に失敗しました。Apps Scriptの実行履歴とエラーシートを確認してください。';
    // 管理画面に渡すのは既知の数値項目だけ。ワーカーやエラーの任意の文章は保存しない。
    if (status === 'success' && kind === 'delivery' && typeof deliveryResult === 'string') {
      const sent = deliveryResult.match(/送信成功\s*：\s*(\d+)件/);
      const errors = deliveryResult.match(/エラー\s*：\s*(\d+)件/);
      const reviews = deliveryResult.match(/要確認\s*：\s*(\d+)件/);
      if (sent && errors && reviews) {
        message = '送信成功：' + sent[1] + '件、エラー：' + errors[1] + '件、要確認：' + reviews[1] + '件。送信シートとエラーシートで確認してください。';
        if (Number(errors[1]) || Number(reviews[1])) status = 'attention';
      }
    }
    if (status !== 'error' && (errorDelta === null || errorDelta > 0)) {
      status = 'attention';
      message += errorDelta === null ? ' エラー記録を確認できませんでした。保守画面で状態を確認してください。' : ' 新しいエラー記録が' + errorDelta + '件あります。保守画面で確認してください。';
    }
    const finishedAt = new Date().toISOString();
    const elapsedMs = Math.max(0, Date.parse(finishedAt) - Date.parse(startedAt));
    runs.unshift({ kind: kind, status: status, startedAt: startedAt, finishedAt: finishedAt, elapsedMs: isFinite(elapsedMs) ? elapsedMs : 0, message: message });
    properties.setProperty(AUTOMATION_RUNS_KEY_, JSON.stringify(runs.slice(0, 20)));
  } catch (error) { Logger.log('自動実行の結果概要を保存できませんでした。'); }
}

function runManagedAutomation_(kind, event) {
  return withAppLock_(function() {
    const record = readAutomationRecord_();
    const owner = PropertiesService.getUserProperties().getProperty(AUTOMATION_OWNER_KEY_);
    if (!event || !event.triggerUid || !record.enabled || !record.owner || record.owner !== owner || !record.active.some(function(row) { return row.id === String(event.triggerUid) && row.kind === kind; })) return;
    const startedAt = new Date().toISOString();
    const errorsBefore = automationErrorRowCount_();
    try {
      let deliveryResult;
      const config = getConfig_();
      validateAppConfig_(config);
      if (kind === 'import') prepareSendDataUnlocked_();
      if (kind === 'delivery') deliveryResult = sendMessagesUnlocked_();
      if (kind === 'reminder') {
        if (!automationRecipientsValid_(config)) throw new Error('リマインダーの通知先が不正です。設定を確認してください。');
        remindUngradedAndErrorsUnlocked_();
      }
      const errorsAfter = automationErrorRowCount_();
      recordAutomationRun_(kind, 'success', startedAt, deliveryResult,
        errorsBefore === null || errorsAfter === null ? null : Math.max(0, errorsAfter - errorsBefore));
    } catch (error) {
      recordAutomationRun_(kind, 'error', startedAt);
      throw error;
    }
  });
}

function managedAutomationImport_(event) { return runManagedAutomation_('import', event); }
function managedAutomationDelivery_(event) { return runManagedAutomation_('delivery', event); }
function managedAutomationReminder_(event) { return runManagedAutomation_('reminder', event); }
