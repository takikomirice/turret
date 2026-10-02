import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import {readFileSync} from 'node:fs';
import {schedulingEnvironment} from './helpers/scheduling-environment.mjs';
import {plain} from './helpers/admin-environment.mjs';
const script=readFileSync('Setting.html','utf8').match(/<script>([\s\S]*?)<\/script>/)[1];

// 実際の画面制御からサーバー処理へ接続し、外部サービスだけを代替する。
function environment(){
 const e=schedulingEnvironment(),r=e.read(),topics=[],topicCalls=[];
 Object.assign(r.input,{titlePattern:'{class_label} フォーム',materialTitlePattern:'{class_label} 資料'});r.label='1A';r.materialSettingsSaved=false;e.c.saveManagedRecord_(r);
 e.c.Classroom.Courses.Topics={list:()=>({topic:plain(topics)}),create:(body,courseId)=>{const topic={...plain(body),courseId,topicId:'topic-one'};topics.push(topic);topicCalls.push(topic);return plain(topic);}};
 const c=vm.createContext({console,structuredClone,Date:e.c.Date});vm.runInContext(script.replace(/boot\(\);\s*$/,''),c);
 c.document={getElementById:()=>null,querySelectorAll:()=>[]};c.savedRecord=plain(e.read());c.backend=async(method,...args)=>plain(e.c[method](...args));
 vm.runInContext(`state.panel='setup';state.step='publish';state.forms={classes:[{courseId:'100'}],defaults:{},records:[savedRecord]};state.batchMaterialDraft={materialTitlePattern:'{class_label} 新しい資料',description:'新しい本文',topicName:'振り返り'};state.scheduleDrafts={'batch:publish':'2026-10-05T10:00','batch:close':'2026-10-06T10:00'};let confirmations=[],messages=[];rpc=backend;requireSaved=()=>true;task=async work=>work();render=()=>{};reloadForms=async()=>{};status=message=>messages.push(message);confirmAction=async(title,summary)=>{confirmations.push(summary);return true;};`,c);
 return {...e,ui:c,topics,topicCalls};
}
test('one aggregate confirmation creates topic, saves settings, schedules material and closes form later',async()=>{
 const e=environment();await vm.runInContext("formCommand('register-all')",e.ui);
 assert.equal(e.topicCalls.length,1);assert.equal(e.read().topicId,'topic-one');assert.equal(e.read().input.description,'新しい本文');assert.equal(e.read().scheduledTime,'2026-10-05T01:00:00.000Z');assert.equal(e.read().closesAt,'2026-10-06T01:00:00.000Z');assert.equal([...e.posts.values()][0].topicId,'topic-one');assert.equal(e.form.accepting,true);assert.equal(vm.runInContext('confirmations.length',e.ui),1);assert.match(vm.runInContext('confirmations[0]',e.ui),/メールアドレスを収集する/);assert.match(vm.runInContext('confirmations[0]',e.ui),/同名トピック.*[\s\S]*作成/);
});
test('cancelled aggregate confirmation does not change backend records, topics, forms or posts',async()=>{
 const e=environment(),before=plain(e.read());vm.runInContext('confirmAction=async()=>false;',e.ui);await vm.runInContext("formCommand('register-all')",e.ui);
 assert.deepEqual(plain(e.read()),before);assert.equal(e.topicCalls.length,0);assert.equal(e.posts.size,0);assert.equal(e.form.published,false);
});
test('aggregate deadline extension works against real server validation of the previous deadline',async()=>{
 const e=environment(),r=e.read();r.closesAt='2026-10-03T01:00:00Z';e.c.saveManagedRecord_(r);e.ui.savedRecord=plain(e.read());vm.runInContext('state.forms.records=[savedRecord];',e.ui);
 await vm.runInContext("formCommand('register-all')",e.ui);
 assert.equal(e.read().scheduledTime,'2026-10-05T01:00:00.000Z');assert.equal(e.read().closesAt,'2026-10-06T01:00:00.000Z');
});
