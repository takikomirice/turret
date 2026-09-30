import test from 'node:test';
import assert from 'node:assert/strict';
import {adminEnvironment,AdminSheet,plain} from './helpers/admin-environment.mjs';
import {NativeForm} from './helpers/native-forms.mjs';

function env(){
 const e=adminEnvironment(),{c,sheets}=e;c.ScriptApp.getScriptId=()=> 'script';c.ScriptApp.getOAuthToken=()=> 'fake';
 const source=new NativeForm('template',['A','B','C','D']),target=new NativeForm('target',['A','B','C','D']);
 // Copied forms may use different IDs.
 target.items.forEach(i=>i.id+=1000);target.next+=1000;
 c.FormApp={openById:id=>id==='template'?source:target,openByUrl:()=>target,PageNavigationType:{CONTINUE:'CONTINUE',SUBMIT:'SUBMIT',RESTART:'RESTART'}};
 c.UrlFetchApp={fetch(){throw Error('Forms REST must not be used');}};
 const s=new AdminSheet('回答',[['日時','メール','A','B','C','D','名前','コメント','状態'],['t','m','a','b','c','','生徒','コメント保持','済']]);
 s.getFormUrl=()=> 'target';const get=s.getRange.bind(s);
 s.getRange=(r,c,h,w)=>Object.assign(get(r,c,h,w),{getColumn:()=>c});
 s.moveColumns=(range,to)=>{const from=range.getColumn()-1;for(const row of s.rows){const [v]=row.splice(from,1);row.splice(to-1-(from<to-1?1:0),0,v??'');}};
 s.deleteColumns=(start,count)=>{for(const row of s.rows)row.splice(start-1,count);};
 target.onTitle=(item,old)=>{if(!item.title)return;const at=s.rows[0].indexOf(old);if(at>=0)s.rows[0][at]=item.title;else{s.rows[0].push(item.title);s.rows[1].push('');}};
 sheets.set(s.name,s);
 const raw={id:'form:test',kind:'form',stage:'published',formId:'target',courseId:'100',label:'テスト',responseSheetId:s.getSheetId(),materialId:'keep',input:{templateId:'template',emailHeader:'メール',nameHeader:'名前',statusHeader:'状態',additionalColumns:[{header:'コメント',choices:[]}]}};
 const setup=()=>{raw.templateItems=c.formMetadata_('template').items;raw.formSync=c.captureManagedFormSync_(raw,c.formMetadata_('target').items,s.rows[0]);delete raw.templateItems;return c.saveManagedRecord_(raw);};
 return {...e,source,target,s,setup};
}

test('native metadata reads ordered items and distinguishes email collection from verified identity without REST',()=>{
 const e=env(),meta=e.c.formMetadata_('template');
 assert.deepEqual(plain(meta.items.map(i=>i.title)),['A','B','C','D']);
 assert.equal(meta.settings.collectsEmail,true);assert.notEqual(meta.settings.emailCollectionType,'VERIFIED');
 e.source.collects=false;assert.equal(e.c.formMetadata_('template').settings.collectsEmail,false);
});

test('a deadline reached during form updating prevents response intake from reopening',()=>{
 const e=env(),{c,source,target}=e,r=e.setup();let now=Date.parse('2026-10-01T00:00:00Z');
 c.Date=class extends Date{constructor(...args){super(...(args.length?args:[now]));}static now(){return now;}};
 r.closesAt='2026-10-01T00:01:00Z';r.closeState='scheduled';c.saveManagedRecord_(r);
 source.addTextItem().setTitle('E');target.onChange=()=>{now=Date.parse('2026-10-01T00:02:00Z');};
 const p=c.previewManagedFormUpdate(r.id);c.runManagedFormUpdate(r.id,p.fingerprint);
 assert.equal(target.accepting,false);assert.equal(c.loadManagedRecord_(r.id).accepting,false);
});

test('a settings change while waiting for sheet columns keeps responses closed',()=>{
 const e=env(),{c,source,target}=e,r=e.setup();source.addTextItem().setTitle('E');
 const arrange=c.arrangeManagedFormColumns_;c.arrangeManagedFormColumns_=()=>{throw Error('wait for sheet');};
 const p=c.previewManagedFormUpdate(r.id);assert.throws(()=>c.runManagedFormUpdate(r.id,p.fingerprint),/wait for sheet/);
 c.arrangeManagedFormColumns_=arrange;target.collects=false;
 assert.throws(()=>c.resumeManagedFormUpdate(r.id,c.loadManagedRecord_(r.id).revision),/一致|変更/);
 assert.equal(target.accepting,false);assert.ok(c.loadManagedRecord_(r.id).formUpdate);
});

test('native insert/delete cycle preserves question identity, answers, custom cells and form URL',()=>{
 const e=env(),{c,source,target,s}=e,r=e.setup(),ids=target.items.map(i=>i.id);
 source.moveItem(source.addTextItem().setTitle('E'),2);
 let p=c.previewManagedFormUpdate(r.id),done=c.runManagedFormUpdate(r.id,p.fingerprint);
 assert.deepEqual(target.items.map(i=>i.title),['A','B','E','C','D']);
 assert.deepEqual(target.items.filter(i=>i.title!=='E').map(i=>i.id),ids);
 assert.deepEqual(s.rows[0],['日時','メール','A','B','E','C','D','名前','コメント','状態']);
 source.getItems().filter(i=>['C','D'].includes(i.title)).forEach(i=>source.deleteItem(i));
 p=c.previewManagedFormUpdate(r.id);done=c.runManagedFormUpdate(r.id,p.fingerprint);
 assert.deepEqual(s.rows[0],['日時','メール','A','B','E','名前','コメント','状態','C']);
 assert.deepEqual(s.rows[1],['t','m','a','b','','生徒','コメント保持','済','c']);
 assert.equal(done.materialId,'keep');assert.equal(done.formId,'target');assert.equal(done.formUpdate,undefined);assert.equal(target.accepting,true);
});

test('FormApp move and delete receive indices rather than specialized question objects',()=>{
 const e=env(),{c,source,target}=e,r=e.setup();source.moveItem(3,0);source.deleteItem(2);
 const move=target.moveItem.bind(target),remove=target.deleteItem.bind(target);
 target.moveItem=(from,to)=>{assert.equal(typeof from,'number');return move(from,to);};
 target.deleteItem=index=>{assert.equal(typeof index,'number');return remove(index);};
 const p=c.previewManagedFormUpdate(r.id);c.runManagedFormUpdate(r.id,p.fingerprint);
 assert.deepEqual(target.items.map(i=>i.title),['D','A','C']);
});

for(const action of ['add','title','required','delete','move'])test('native interrupted '+action+' reconciles without duplicating questions or answers',()=>{
 const e=env(),{c,source,target}=e,r=e.setup();
 source.moveItem(source.addTextItem().setTitle('E').setRequired(true),1);source.deleteItem(source.items.find(i=>i.title==='D'));
 let thrown=false;target.onChange=(op)=>{if(!thrown&&op===action){thrown=true;throw Error('lost response');}};
 const p=c.previewManagedFormUpdate(r.id);assert.throws(()=>c.runManagedFormUpdate(r.id,p.fingerprint),/lost response/);
 assert.equal(target.accepting,false);assert.throws(()=>c.assertNoManagedFormUpdate_(),/更新/);
 target.onChange=null;const done=c.resumeManagedFormUpdate(r.id,c.loadManagedRecord_(r.id).revision);
 assert.equal(done.formUpdate,undefined);assert.deepEqual(target.items.map(i=>i.title),['A','E','B','C']);assert.equal(target.items.filter(i=>i.title==='E').length,1);
 assert.equal(target.items.find(i=>i.title==='E').required,true);assert.equal(target.accepting,true);
});

test('native stale preview and unrelated edits during a failed write cannot be overwritten',()=>{
 const e=env(),{c,source,target}=e,r=e.setup();source.items[0].setTitle('新A');
 const p=c.previewManagedFormUpdate(r.id);assert.throws(()=>c.runManagedFormUpdate(r.id,'bad'),/確認/);assert.equal(target.mutations,0);
 target.onChange=()=>{throw Error('lost response');};assert.throws(()=>c.runManagedFormUpdate(r.id,p.fingerprint),/lost response/);
 target.onChange=null;target.items[1].setTitle('外部編集');
 assert.throws(()=>c.resumeManagedFormUpdate(r.id,c.loadManagedRecord_(r.id).revision),/一致|変更/);assert.equal(target.items[1].title,'外部編集');assert.equal(target.accepting,false);
});

test('native rename retains configured sheet header; later updates and closed intake stay valid',()=>{
 const e=env(),{c,source,target,s}=e,r=e.setup();source.items[0].setTitle('新A');target.accepting=false;
 let p=c.previewManagedFormUpdate(r.id);c.runManagedFormUpdate(r.id,p.fingerprint);
 assert.equal(target.items[0].title,'新A');assert.equal(s.rows[0][2],'A');assert.equal(target.accepting,false);
 source.items[1].setHelpText('新しい説明');p=c.previewManagedFormUpdate(r.id);c.runManagedFormUpdate(r.id,p.fingerprint);
 assert.equal(target.items[1].help,'新しい説明');assert.equal(s.rows[1][2],'a');
});

test('quiz or unsupported update stops before mutating a distributed form',()=>{
 const e=env(),{c,source,target}=e,r=e.setup();source.quiz=target.quiz=true;
 assert.throws(()=>c.previewManagedFormUpdate(r.id),/テスト|採点/);assert.equal(target.mutations,0);
 source.quiz=target.quiz=false;source.add('FILE_UPLOAD').setTitle('添付');
 assert.throws(()=>c.previewManagedFormUpdate(r.id),/未対応|ファイル/);assert.equal(target.mutations,0);
});

test('image additions and replacements stop before native updates; unchanged images remain supported',()=>{
 const e=env(),{c}=e,r=e.setup(),source=c.formMetadata_('template'),target=c.formMetadata_('target');
 source.items.push({itemId:'image',title:'画像',description:'',native:{type:'IMAGE',imageHash:'new',width:400,alignment:'CENTER'}});
 assert.throws(()=>c.buildManagedFormUpdate_(r,source,target,e.s.rows[0]),/画像/);
 target.items.push({...plain(source.items.at(-1)),itemId:'copy-image'});r.formSync.itemMap.image='copy-image';
 assert.doesNotThrow(()=>c.buildManagedFormUpdate_(r,source,target,e.s.rows[0]));
 source.items.at(-1).native.imageHash='changed';
 assert.throws(()=>c.buildManagedFormUpdate_(r,source,target,e.s.rows[0]),/画像/);
});

test('failed sheet movement resumes without repeating native edits; copied ownership is rejected',()=>{
 const e=env(),{c,source,target,s}=e,r=e.setup();source.moveItem(source.addTextItem().setTitle('E'),2);
 const p=c.previewManagedFormUpdate(r.id),move=s.moveColumns;s.moveColumns=()=>{throw Error('sheet busy');};
 assert.throws(()=>c.runManagedFormUpdate(r.id,p.fingerprint),/sheet busy/);const count=target.mutations;
 s.moveColumns=move;c.resumeManagedFormUpdate(r.id,c.loadManagedRecord_(r.id).revision);assert.equal(target.mutations,count);
 c.ScriptApp.getScriptId=()=> 'other';assert.throws(()=>c.previewManagedFormUpdate(r.id),/別の運用/);
});

test('native creation checkpoint failure preserves identity mapping on resume',()=>{
 const e=env(),{c,source,target}=e,r=e.setup();source.addTextItem().setTitle('E');
 const save=c.saveManagedRecord_;let fail=true;
 c.saveManagedRecord_=record=>{if(fail&&record.formUpdate?.checkpoint.items.length===5&&!record.formUpdate.pending){fail=false;throw Error('storage');}return save(record);};
 const p=c.previewManagedFormUpdate(r.id);assert.throws(()=>c.runManagedFormUpdate(r.id,p.fingerprint),/storage/);
 c.saveManagedRecord_=save;c.resumeManagedFormUpdate(r.id,c.loadManagedRecord_(r.id).revision);
 assert.equal(target.items.length,5);assert.equal(target.items.at(-1).title,'E');
});

test('new section and multiple choice navigation use copied destination IDs',()=>{
 const e=env(),{c,source,target,s}=e,r=e.setup();
 const section=source.addPageBreakItem().setTitle('次のページ'),choice=source.addMultipleChoiceItem().setTitle('分岐');
 choice.setChoices([choice.createChoice('次へ',section),choice.createChoice('終了','SUBMIT')]);source.moveItem(choice,0);
 // Only question titles create response columns (layout titles do not).
 const onTitle=target.onTitle;target.onTitle=(i,old)=>{if(i.type!=='PAGE_BREAK')onTitle(i,old);};
 const p=c.previewManagedFormUpdate(r.id);c.runManagedFormUpdate(r.id,p.fingerprint);
 const copied=target.items.find(i=>i.title==='次のページ'),copiedChoice=target.items[0];
 assert.notEqual(copied.id,section.id);assert.equal(copiedChoice.choices[0].getGotoPage().id,copied.id);
 assert.equal(copiedChoice.choices[1].getPageNavigationType(),'SUBMIT');assert.equal(s.rows[0][2],'分岐');
});
