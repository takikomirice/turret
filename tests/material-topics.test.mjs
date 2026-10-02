import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {schedulingEnvironment} from './helpers/scheduling-environment.mjs';
import {AdminSheet, plain} from './helpers/admin-environment.mjs';

function environment(){
 const e=schedulingEnvironment(),r=e.read(),topicCalls=[],topics=new Map([['100',[]],['200',[]]]);let serial=0;
 Object.assign(r,{label:'1A',materialTitle:'1A 資料'});Object.assign(r.input,{titlePattern:'{class_label} フォーム',materialTitlePattern:'{class_label} 資料'});e.c.saveManagedRecord_(r);
 e.sheets.set('クラス一覧',new AdminSheet('クラス一覧',[['クラス名','コースID','同期対象(1)'],['対象外','900',''],['1A','100','1'],['2A','200','1']]));
 const topicService={
  list(course,options){assert.equal(e.isLocked(),true);topicCalls.push(['list',course,plain(options)]);const all=topics.get(course)||[],offset=Number(options.pageToken||0),page=all.slice(offset,offset+2);return {topic:plain(page),...(offset+2<all.length?{nextPageToken:String(offset+2)}:{})};},
  create(body,course){assert.equal(e.isLocked(),true);topicCalls.push(['create',plain(body),course]);const topic={topicId:'t'+(++serial),name:body.name,courseId:course,updateTime:'2026-10-01T00:00:00Z'};topics.set(course,[...(topics.get(course)||[]),topic]);return plain(topic);}
 };
 e.c.Classroom.Courses.Topics=topicService;
 const preview=(changes,ids=['f1'],common=false)=>e.c.previewManagedFormSettings({targets:ids.map(id=>({id,revision:e.c.loadManagedRecord_(id).revision})),changes,saveMaterialSettings:true,registerCommonMaterial:common});
 const apply=(changes,p,approve=false,common=false)=>p.targets.map(t=>e.c.runManagedFormSettings(t.id,changes,t.fingerprint,true,common,approve));
 const save=(changes,approve=false,common=false)=>apply(changes,preview(changes,['f1'],common),approve,common)[0];
 const add=(id,courseId='200')=>{const copy=plain(e.read());delete copy.materialId;Object.assign(copy,{id,courseId,className:courseId==='100'?'1A':'2A',label:courseId==='100'?'1A':'2A',stage:'registered'});return e.c.saveManagedRecord_(copy);};
 return {...e,topics,topicCalls,topicService,previewSettings:preview,applySettings:apply,saveSettings:save,add};
}

// 最初の対象クラスと全ページを使わない実装を検出する。
test('topic picker reads every page of the first selected class without creating or saving anything',()=>{
 const e=environment();e.topics.set('100',[{topicId:'a',name:'単元1'},{topicId:'b',name:'単元2'},{topicId:'c',name:'単元3'}]);e.topics.set('200',[{topicId:'other',name:'別クラス'}]);const before=e.read().revision;
 assert.equal(typeof e.c.getManagedMaterialTopics,'function');
 assert.deepEqual(plain(e.c.getManagedMaterialTopics()),{courseId:'100',courseName:'1A',topics:[{topicId:'a',name:'単元1'},{topicId:'b',name:'単元2'},{topicId:'c',name:'単元3'}]});
 assert.deepEqual(e.topicCalls.map(c=>[c[0],c[1],c[2].pageToken||'']),[['list','100',''],['list','100','2']]);assert.equal(e.read().revision,before);
});
test('topic preview maps same names to course-local IDs and explicitly reports missing targets without writes',()=>{
 const e=environment();e.add('f2');e.topics.set('100',[{topicId:'first',name:'振り返り'}]);const before=[e.read().revision,e.c.loadManagedRecord_('f2').revision];
 const p=e.previewSettings({topicName:'振り返り'},['f1','f2']);
 assert.equal(p.targets[0].missingTopic,null);assert.deepEqual(plain(p.targets[1].missingTopic),{courseId:'200',className:'2A',name:'振り返り'});assert.match(p.targets[0].changes.join(' '),/トピック/);
 assert.deepEqual([e.read().revision,e.c.loadManagedRecord_('f2').revision],before);assert.equal(e.topicCalls.filter(c=>c[0]==='create').length,0);
});
test('a missing topic cannot be created without explicit confirmation and cannot borrow another name fingerprint',()=>{
 const e=environment(),changes={topicName:'振り返り'},p=e.previewSettings(changes);
 assert.throws(()=>e.applySettings(changes,p),/トピック.*確認|確認.*トピック/);assert.equal(e.read().settingsUpdate,undefined);assert.equal(e.topics.get('100').length,0);
 assert.throws(()=>e.applySettings({topicName:'別名'},p,true),/更新|確認/);assert.equal(e.topics.get('100').length,0);
 const result=e.applySettings(changes,p,true)[0];assert.equal(result.topicId,'t1');assert.equal(result.input.topicName,'振り返り');assert.equal(result.stage,'registered');assert.equal(e.form.published,false);
});
test('confirmed batch reuses a single newly-created topic across multiple records of one course',()=>{
 const e=environment();e.add('f2','100');const changes={topicName:'振り返り'},p=e.previewSettings(changes,['f1','f2']);
 const results=e.applySettings(changes,p,true);assert.deepEqual(results.map(r=>r.topicId),['t1','t1']);assert.equal(e.topics.get('100').length,1);
});
test('existing same-name topics use distinct local IDs and are included in new material creation',()=>{
 const e=environment();e.add('f2');e.topics.set('100',[{topicId:'local100',name:'振り返り'}]);e.topics.set('200',[{topicId:'local200',name:'振り返り'}]);const changes={topicName:'振り返り'};
 const results=e.applySettings(changes,e.previewSettings(changes,['f1','f2']));assert.deepEqual(results.map(r=>r.topicId),['local100','local200']);e.run('publish');assert.equal([...e.posts.values()][0].topicId,'local100');assert.equal(e.topicCalls.filter(c=>c[0]==='create').length,0);
});
test('topic assignment and clearing patch only the existing material topic and preserve its scheduled state',()=>{
 const e=environment();e.run('schedule',{scheduledTime:'2026-10-02T01:00:00Z'});const original=plain(e.read()),postId=original.materialId;e.topics.set('100',[{topicId:'local100',name:'振り返り'}]);
 const assigned=e.saveSettings({topicName:'振り返り'});assert.equal(assigned.materialId,postId);assert.equal(assigned.stage,'scheduled');assert.equal(e.posts.get(postId).topicId,'local100');
 e.saveSettings({topicName:''});const cleared=e.posts.get(postId),patch=e.calls.filter(c=>c[0]==='patch').at(-1);assert.equal(cleared.topicId,undefined);assert.deepEqual(patch[1],{});assert.equal(patch[4].updateMask,'topicId');assert.equal(cleared.state,'DRAFT');assert.equal(cleared.scheduledTime,'2026-10-02T01:00:00.000Z');assert.equal(e.read().topicId,'');assert.equal(e.read().input.topicName,'');assert.equal(e.posts.size,1);
});
test('blank topics require no Topics access and are omitted from new material bodies',()=>{
 const e=environment();delete e.c.Classroom.Courses.Topics;const saved=e.saveSettings({topicName:''});assert.equal(saved.input.topicName,'');assert.equal(saved.topicId,'');e.run('publish');assert.equal(Object.hasOwn([...e.posts.values()][0],'topicId'),false);
});
test('ambiguous topic names on later pages and changed topic mappings fail before mutations',()=>{
 const e=environment();e.topics.set('100',[{topicId:'a',name:'振り返り'},{topicId:'b',name:'別名'},{topicId:'c',name:'振り返り'}]);
 assert.throws(()=>e.previewSettings({topicName:'振り返り'}),/重複|一意/);assert.equal(e.topicCalls.filter(c=>c[0]==='create').length,0);
 e.topics.set('100',[{topicId:'a',name:'振り返り'}]);const changes={topicName:'振り返り'},p=e.previewSettings(changes);e.topics.set('100',[{topicId:'replaced',name:'振り返り'}]);assert.throws(()=>e.applySettings(changes,p,true),/更新|確認/);assert.equal(e.read().settingsUpdate,undefined);
});
test('uncertain topic creation reconciles by name on resume and never creates twice',()=>{
 const e=environment(),create=e.topicService.create;e.topicService.create=(...args)=>{create(...args);throw Error('response lost');};
 assert.throws(()=>e.saveSettings({topicName:'振り返り'},true),/response lost/);assert.ok(e.read().settingsUpdate);assert.equal(e.form.published,false);
 const result=e.c.resumeManagedFormSettings('f1',e.read().revision);assert.equal(result.topicId,'t1');assert.equal(result.settingsUpdate,undefined);assert.equal(e.topicCalls.filter(c=>c[0]==='create').length,1);
});
test('an unconfirmed creation result remains blocked when listing is empty, including another record',()=>{
 const e=environment();e.add('f2','100');e.topicService.create=(body,course)=>{e.topicCalls.push(['create',plain(body),course]);throw Error('response lost');};
 assert.throws(()=>e.saveSettings({topicName:'振り返り'},true),/response lost/);
 for(let i=0;i<2;i++)assert.throws(()=>e.c.resumeManagedFormSettings('f1',e.read().revision),/未確認|照合|再作成/);
 const changes={topicName:'振り返り'},p=e.previewSettings(changes,['f2']);assert.throws(()=>e.applySettings(changes,p,true),/未確認|照合|再作成/);assert.equal(e.topicCalls.filter(c=>c[0]==='create').length,1);
});
test('durable settings failure happens before any topic creation',()=>{
 const e=environment(),changes={topicName:'振り返り'},p=e.previewSettings(changes);e.sheets.get('システム管理').beforeWrite=()=>{throw Error('storage failed');};
 assert.throws(()=>e.applySettings(changes,p,true),/storage failed/);assert.equal(e.topicCalls.filter(c=>c[0]==='create').length,0);assert.equal(e.read().settingsUpdate,undefined);
});
test('topic name normalization obeys Classroom limits and common baselines preserve it across individual edits',()=>{
 const e=environment();const result=e.saveSettings({topicName:'  Unit   One  '},true,true);assert.equal(result.input.topicName,'Unit One');assert.equal(result.materialSettingsBaseline.topicName,'Unit One');assert.equal(e.topics.get('100')[0].name,'Unit One');
 e.saveSettings({description:'個別本文'});assert.equal(e.read().topicId,'t1');assert.equal(e.read().materialSettingsBaseline.topicName,'Unit One');
 for(const topicName of [3,'x'.repeat(101),null])assert.throws(()=>e.previewSettings({topicName}),/入力|トピック/);
});
test('topic scope is declared for listing and confirmed creation',()=>{
 const manifest=JSON.parse(readFileSync('appsscript.json','utf8'));assert.ok(manifest.oauthScopes.includes('https://www.googleapis.com/auth/classroom.topics'));
});
test('read-only topic loading never creates sheet filters or repairs an empty class sheet',()=>{
 const e=environment(),sheet=e.sheets.get('クラス一覧');assert.equal(sheet.getFilter(),null);e.c.getManagedMaterialTopics();assert.equal(sheet.getFilter(),null);
 sheet.clear();sheet.beforeWrite=()=>{throw Error('unexpected sheet write');};assert.throws(()=>e.c.getManagedMaterialTopics(),/対象クラス|クラス.*空/);assert.equal(sheet.getLastRow(),0);
});
test('uncertain material creation only reconciles posts carrying the selected topic',()=>{
 const e=environment();e.topics.set('100',[{topicId:'wanted',name:'振り返り'}]);e.saveSettings({topicName:'振り返り'});const create=e.service.create;e.service.create=(...args)=>{create(...args);throw Error('response lost');};
 assert.throws(()=>e.run('publish'),/response lost/);const post=[...e.posts.values()][0];post.topicId='other';assert.throws(()=>e.c.reconcileManagedPost('f1',e.read().revision),/一意|照合/);assert.equal(e.read().stage,'publish_review');
 post.topicId='wanted';const result=e.c.reconcileManagedPost('f1',e.read().revision);assert.equal(result.materialId,post.id);assert.equal(result.topicId,'wanted');assert.equal(e.posts.size,1);
});
test('lost topic patch responses resume the existing published material without changing its publication state',()=>{
 const e=environment();e.run('publish');const id=e.read().materialId;e.topics.set('100',[{topicId:'wanted',name:'振り返り'}]);const patch=e.service.patch;e.service.patch=(...args)=>{patch(...args);throw Error('patch response lost');};
 assert.throws(()=>e.saveSettings({topicName:'振り返り'}),/patch response lost/);assert.ok(e.read().settingsUpdate);e.service.patch=patch;
 const result=e.c.resumeManagedFormSettings('f1',e.read().revision);assert.equal(result.settingsUpdate,undefined);assert.equal(result.topicId,'wanted');assert.equal(result.materialId,id);assert.equal(result.stage,'published');assert.equal(e.posts.get(id).state,'PUBLISHED');assert.equal(e.posts.size,1);assert.equal(e.topicCalls.filter(c=>c[0]==='create').length,0);
});
test('manual topic edits during patch recovery are not overwritten',()=>{
 const e=environment();e.run('publish');const id=e.read().materialId;e.topics.set('100',[{topicId:'wanted',name:'振り返り'}]);const patch=e.service.patch;e.service.patch=(...args)=>{patch(...args);throw Error('patch response lost');};
 assert.throws(()=>e.saveSettings({topicName:'振り返り'}),/patch response lost/);e.service.patch=patch;e.posts.get(id).topicId='manual';
 assert.throws(()=>e.c.resumeManagedFormSettings('f1',e.read().revision),/手動変更/);assert.equal(e.posts.get(id).topicId,'manual');assert.ok(e.read().settingsUpdate);
});
test('malformed pages, repeated continuation tokens and permissions failures never trigger topic creation',()=>{
 for(const fixture of [()=>({topic:[{topicId:'a',name:'振り返り',courseId:'other'}]}),()=>({topic:[],nextPageToken:'loop'}),()=>({topic:'bad'}),()=>{throw Error('permission denied');}]){
  const e=environment();e.topicService.list=fixture;assert.throws(()=>e.saveSettings({topicName:'振り返り'},true));assert.equal(e.topicCalls.filter(c=>c[0]==='create').length,0);assert.equal(e.read().settingsUpdate,undefined);
 }
});
test('topic journal persistence failures stop before create and before material changes',()=>{
 const e=environment(),original=e.c.PropertiesService.getScriptProperties;e.c.PropertiesService.getScriptProperties=()=>({...original(),setProperty(key,value){if(key.startsWith('TURRET_MATERIAL_TOPIC_'))throw Error('journal failed');return original().setProperty(key,value);}});
 assert.throws(()=>e.saveSettings({topicName:'振り返り'},true),/journal failed/);assert.equal(e.topicCalls.filter(c=>c[0]==='create').length,0);assert.ok(e.read().settingsUpdate);assert.equal(e.posts.size,0);
});
test('external same-name appearance after preview requires repreview rather than borrowing creation approval',()=>{
 const e=environment(),changes={topicName:'振り返り'},p=e.previewSettings(changes);e.topics.set('100',[{topicId:'external',name:'振り返り'}]);assert.throws(()=>e.applySettings(changes,p,true),/確認|更新/);assert.equal(e.read().settingsUpdate,undefined);
 assert.equal(e.saveSettings(changes).topicId,'external');assert.equal(e.topicCalls.filter(c=>c[0]==='create').length,0);
});
