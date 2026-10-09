import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import {readFileSync} from 'node:fs';
function env(){
 const c=vm.createContext({console});
 vm.runInContext(readFileSync('Setting.html','utf8').match(/<script>([\s\S]*?)<\/script>/)[1].replace(/boot\(\);\s*$/,''),c);
 vm.runInContext(`state.forms={classes:[],records:[],defaults:{},setupProgress:{forms:{id:'forms',state:'complete'}}};state.data={config:{formSources:[{id:'ss'}],formSheetNamePrefix:['回答'],reminderTo:[]},counts:{mappings:1},mappingNeeded:true,sheets:[{key:'mapping',exists:true,url:'https://example.invalid/mapping'}],progress:[{id:'sources',state:'complete'},{id:'mapping',state:'attention',detail:'クラスを選択'}]};`,c);
 return c;
}
test('mapping controls belong to form preparation and notifications remain independently accessible',()=>{
 const c=env(),forms=vm.runInContext("renderStep('forms')",c),sources=vm.runInContext("renderStep('sources')",c);
 assert.match(forms,/対応表シートへ/);assert.match(forms,/data-action="mapping"/);assert.doesNotMatch(forms,/通知先と対応表へ/);
 assert.doesNotMatch(sources,/data-action="mapping"|回答シートとクラスの対応表/);assert.match(sources,/通知先メールアドレス/);
 assert.equal(vm.runInContext("visibleProgress('sources').state",c),'complete');
 assert.equal(vm.runInContext("visibleProgress('forms').state",c),'attention');
});
test('mapping remains available even when the form list cannot load',()=>{
 const c=env();vm.runInContext("state.forms=null;state.formsError='取得失敗';",c);
 const html=vm.runInContext("renderStep('forms')",c);assert.match(html,/対応表シートへ/);assert.match(html,/data-action="mapping"/);
});
