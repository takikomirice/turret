import { test } from 'node:test';
import assert from 'node:assert/strict';
import { adminEnvironment, AdminSheet, sourceId, plain } from './helpers/admin-environment.mjs';

function environment() {
  const e=adminEnvironment();
  e.sheets.set('生徒一覧',new AdminSheet('生徒一覧',[
    ['No','メールアドレス','名前','クラス名','コースID','studentId'],
    [1,'student@example.invalid','生徒','授業','course-1','student-1']
  ]));
  const source=new AdminSheet('回答1', [['メール','名前','状態','点数'],['student@example.invalid','回答名','',5]]);
  e.c.SpreadsheetApp.openById=id=>{assert.equal(id,sourceId);return {getName:()=> 'フォーム',getSheets:()=>[source]};};
  e.c.Classroom={Courses:{Announcements:{create(){throw Error('diagnostics must not send');}}}};
  const snapshot=()=>JSON.stringify({props:[...e.props],sheets:[...e.sheets].map(([n,s])=>[n,s.rows]),source:source.rows});
  return {...e,source,snapshot,run:()=>plain(e.c.diagnoseSetup(e.c.getConfigRevision_(e.c.getConfig_())))};
}

test('diagnosis resolves actual recipient and template without changing sheets, properties or sending',()=>{
  const e=environment(),before=e.snapshot(),r=e.run();
  assert.equal(r.status,'ok');assert.equal(r.counts.eligible,1);assert.equal(r.counts.previewed,1);
  assert.equal(r.samples[0].email,'student@example.invalid');assert.equal(r.samples[0].courseId,'course-1');
  assert.match(r.samples[0].body,/生徒さんへ\(個別メッセージ\)/);assert.match(r.samples[0].body,/評価：A/);
  assert.equal(e.snapshot(),before);
});

test('diagnosis reports missing columns and inaccessible sources without partial writes',()=>{
  const e=environment();e.source.rows[0][2]='削除した列';const before=e.snapshot();
  const r=e.run();assert.equal(r.status,'error');assert.ok(r.checks.some(c=>c.detail.includes('状態')));assert.equal(e.snapshot(),before);
  e.c.SpreadsheetApp.openById=()=>{throw Error('permission denied');};
  assert.ok(e.run().checks.some(c=>c.title.includes('接続')));
});

test('ambiguous recipients and duplicate mappings are blocked with corrective guidance',()=>{
  const e=environment();e.sheets.get('生徒一覧').rows.push([2,'student@example.invalid','生徒','別授業','course-2','student-1']);
  let r=e.run();assert.equal(r.counts.blocked,1);assert.ok(r.checks.some(c=>/対応表/.test(c.action)));
  const row=[sourceId,'フォーム','回答1','授業','course-1',''];
  e.sheets.set('対応表',new AdminSheet('対応表',[['元SS_ID','元SS名','元シート名','クラス名','courseId','メモ'],row,row]));
  r=e.run();assert.equal(r.status,'error');assert.ok(r.checks.some(c=>/重複/.test(c.detail)));
});

test('pending send queue uses its saved body and does not duplicate source preview',()=>{
  const e=environment();e.source.rows[1][2]='準備〇';
  e.sheets.set('送信シート',new AdminSheet('送信シート',[
    ['元SS_ID','元シート名','元行番号','No','メールアドレス','名前','クラス名','コースID','studentId','送信状態','点数','返信本文'],
    [sourceId,'回答1',2,7,'student@example.invalid','生徒','授業','course-1','student-1','','B','保存済みの文面']
  ]));
  const r=e.run();assert.equal(r.counts.eligible,1);assert.equal(r.samples[0].stage,'送信待ち');
  assert.match(r.samples[0].body,/保存済みの文面/);assert.doesNotMatch(r.samples[0].body,/評価：A/);
});

test('pending evaluation renders stored converted values, not a changed conversion scale',()=>{
  const e=environment();e.source.rows[1][2]='反映〇';
  e.sheets.set('評価データ',new AdminSheet('評価データ',[
    ['元SS_ID','元シート名','元行番号','メールアドレス','名前','処理状態','点数'],
    [sourceId,'回答1',2,'student@example.invalid','生徒','','以前の評価']
  ]));
  const r=e.run();assert.equal(r.counts.eligible,1);assert.equal(r.samples[0].stage,'本文生成待ち');assert.match(r.samples[0].body,/評価：以前の評価/);
});

test('unknown template fields and stale settings cannot be reported as valid',()=>{
  const e=environment();assert.throws(()=>e.c.diagnoseSetup('old'),/変更/);
  const config=e.c.getConfig_();config.messageTemplate='＜not_registered＞';e.props.set('APP_CONFIG',JSON.stringify(config));
  const r=e.run();assert.equal(r.status,'error');assert.ok(r.checks.some(c=>/差し込み/.test(c.detail)));
});

test('diagnosis counts 500 candidates but bounds returned personal samples',()=>{
  const e=environment();for(let i=1;i<500;i++)e.source.rows.push(['student@example.invalid','生徒','',5]);
  const r=e.run();assert.equal(r.counts.eligible,500);assert.equal(r.counts.previewed,500);assert.equal(r.samples.length,10);assert.equal(r.truncated,false);
  e.source.rows.push(['student@example.invalid','生徒','',5]);
  const capped=e.run();assert.equal(capped.truncated,true);assert.notEqual(capped.status,'ok');
});

test('diagnosis enforces runtime column order and all mapping columns',()=>{
  const e=environment(),headers=plain(e.c.getConfiguredEvalHeaders_(e.c.getConfig_()));
  [headers[0],headers[1]]=[headers[1],headers[0]];
  e.sheets.set('評価データ',new AdminSheet('評価データ',[headers]));
  assert.equal(e.run().status,'error');
  e.sheets.delete('評価データ');
  e.sheets.set('対応表',new AdminSheet('対応表',[['元SS_ID','元シート名','courseId'],[sourceId,'回答1','course-1']]));
  assert.equal(e.run().status,'error');
});

test('diagnosis matches last-row numbering and deduplicates pending evaluations',()=>{
  const e=environment(),config=e.c.getConfig_();config.messageTemplate='番号＜row_no＞';e.props.set('APP_CONFIG',JSON.stringify(config));
  const row=[sourceId,'回答1',2,'student@example.invalid','生徒','','A'];
  e.sheets.set('評価データ',new AdminSheet('評価データ',[plain(e.c.getConfiguredEvalHeaders_(config)),row,row]));
  const queue=n=>[sourceId,'回答1',n+10,n,'student@example.invalid','生徒','授業','course-1','student-1','済','A','保存本文'];
  e.sheets.set('送信シート',new AdminSheet('送信シート',[plain(e.c.getConfiguredSendHeaders_(config)),queue(9),queue(3)]));
  const r=e.run();assert.equal(r.counts.previewed,1);assert.match(r.samples[0].body,/番号4/);
});

test('unmapped grades remain literal and diagnostics explain this',()=>{
  const e=environment();e.source.rows[1][3]='未提出';const r=e.run();
  assert.equal(r.status,'attention');assert.match(r.samples[0].body,/評価：未提出/);
  assert.ok(r.checks.some(c=>c.title.includes('評価変換') && c.detail.includes('そのまま')));
});

test('diagnosis and template examples use each field conversion without extra source reads',()=>{
  const e=environment(),config=e.c.getConfig_();
  config.fields[0].gradeScale=[{from:'5',to:'○'}];
  config.fields.push({...config.fields[0],key:'second',sourceHeader:'第二評価',evalHeader:'第二評価',sendHeader:'第二評価',gradeScale:[{from:'5',to:'B'}]});
  config.messageTemplate='一：＜score＞ 二：＜second＞';
  e.props.set('APP_CONFIG',JSON.stringify(config));
  e.source.rows[0].push('第二評価');e.source.rows[1].push(5);
  let opens=0;const open=e.c.SpreadsheetApp.openById;
  e.c.SpreadsheetApp.openById=id=>{opens++;return open(id);};
  const before=e.snapshot(),result=e.run();
  assert.equal(result.status,'ok');assert.match(result.samples[0].body,/一：○ 二：B/);
  assert.equal(opens,1);assert.equal(e.snapshot(),before);
  const preview=plain(e.c.previewReplyTemplate(config));
  assert.equal(preview.tokens.find(t=>t.key==='score').example,'○');
  assert.equal(preview.tokens.find(t=>t.key==='second').example,'B');
  config.fields[1].gradeScale=[{from:'4',to:'C'}];e.props.set('APP_CONFIG',JSON.stringify(config));
  const warnings=e.run().checks.filter(c=>c.title.includes('評価変換'));
  assert.equal(warnings.length,1);assert.match(warnings[0].title,/第二評価/);
});

test('health and diagnostics never call corrupt stored configuration healthy',()=>{
  const e=environment(),config=e.c.getConfig_();
  e.c.getAdminConsoleData();e.props.set('APP_CONFIG','corrupt');
  assert.ok(e.c.getAdminConsoleData().health.items.some(i=>i.code==='CONFIG_RECOVERY'));
  assert.equal(e.run().status,'error');
});

test('health surfaces latest failure per automation type and supersedes recovered runs',()=>{
  const e=environment();e.props.set('APP_AUTOMATION_RUNS',JSON.stringify([
    {kind:'import',status:'success'}, {kind:'delivery',status:'error'}, {kind:'import',status:'error'}
  ]));
  const health=plain(e.c.getAdminConsoleData().health);
  assert.equal(health.items.find(i=>i.code==='AUTOMATION_RESULT').count,1);
});

test('health distinguishes unresolved queue rows from historical logs and supplies recovery actions',()=>{
  const e=environment();e.sheets.set('送信シート',new AdminSheet('送信シート',[
    ['送信状態'],['済'],['送信確認待ち'],['エラー'],['未'],['']
  ]));e.sheets.set('エラー',new AdminSheet('エラー',[['日時','シート','行','メール','状態','種別','内容'],['today','回答1',2,'','','FORM_UPDATE_ERROR','missing column']]));
  const health=plain(e.c.getAdminConsoleData().health);
  assert.equal(health.status,'attention');assert.equal(health.items.find(i=>i.code==='SEND_REVIEW').count,1);
  assert.ok(health.items.find(i=>i.code==='SEND_REVIEW').action.includes('投稿'));
  assert.ok(health.items.find(i=>i.code==='FORM_UPDATE_ERROR').action.includes('列'));
  assert.ok(health.items.find(i=>i.code==='ERROR_LOG').detail.includes('過去'));
});
