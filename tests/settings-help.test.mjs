import { readFileSync } from 'node:fs';
import vm from 'node:vm';
import { test } from 'node:test';
import assert from 'node:assert/strict';

const html=readFileSync('Setting.html','utf8');
const script=html.match(/<script>([\s\S]*?)<\/script>/)[1];
function model(){
  const c=vm.createContext({console,structuredClone,setTimeout,clearTimeout});
  vm.runInContext(script.replace(/boot\(\);\s*$/,''),c);
  vm.runInContext(`state.data={config:{fields:[],gradeScale:[],messageTemplate:'',formSources:[],formSheetNamePrefix:[],reminderTo:[]},revision:'r',counts:{},sheets:[],progress:[]};state.automation={schedule:{importHour:5,deliveryHour:8,reminderHour:16,reminderEnabled:false}};state.forms={classes:[],defaults:{},records:[]};`,c);
  return c;
}
const visible=markup=>markup.replace(/<[^>]*>/g,'');

test('context help uses a named native button and preserves form label associations',()=>{
  const c=model();
  const markup=vm.runInContext(`headerInput('emailHeader','生徒メールの列','<未保存>','回答者を名簿と照合します。')`,c);
  assert.match(markup,/<label for="emailHeader">生徒メールの列<\/label>/);
  assert.match(markup,/type="button"[^>]*aria-label="生徒メールの列の説明"/);
  assert.doesNotMatch(markup,/aria-haspopup="dialog"/);
  assert.match(markup,/>\?<\/button>/);assert.match(markup,/value="&lt;未保存&gt;"/);
  assert.doesNotMatch(visible(markup),/回答者を名簿と照合/);
});

test('help safely encodes quotes and markup in both titles and bodies',()=>{
  const c=model();c.title='説明"<img>';c.body='\" onclick=\"bad() <script>bad()</script> &';
  const markup=vm.runInContext('helpButton(title,body)',c);
  assert.match(markup,/&quot;/);assert.match(markup,/&lt;script&gt;/);
  assert.doesNotMatch(markup,/<script>|<img>|" onclick=/);
});

function helpDom(c){
  const handlers={},title={},body={},attrs={};
  const tip={hidden:true,style:{},offsetWidth:360,offsetHeight:120,contains:target=>target===tip};
  const b={dataset:{helpTitle:'説明',helpBody:'<b>そのまま表示</b>\n2行目'},setAttribute:(k,v)=>attrs[k]=v,removeAttribute:k=>delete attrs[k],getBoundingClientRect:()=>({left:700,top:500,bottom:518}),closest:()=>b};
  c.document={getElementById:id=>({contextHelpTitle:title,contextHelpBody:body,contextTooltip:tip}[id]),addEventListener:(name,fn)=>handlers[name]=fn};
  c.window={innerWidth:800,innerHeight:600,addEventListener:(name,fn)=>handlers[name]=fn};
  return {handlers,title,body,tip,b,attrs};
}

test('hover and keyboard focus show a nonmodal tooltip without editing or RPC',()=>{
  const c=model(),d=helpDom(c);
  vm.runInContext(`state.drafts={template:{payload:{messageTemplate:'入力途中'},revision:'r'}};readCurrent=render=rpc=()=>{throw Error('help must not change the editor');};initHelp();`,c);
  const before=vm.runInContext('JSON.stringify(state)',c);
  d.handlers.pointerover({target:d.b});
  assert.equal(d.tip.hidden,false);assert.equal(d.title.textContent,'説明');assert.equal(d.body.textContent,'<b>そのまま表示</b>\n2行目');
  assert.equal(d.attrs['aria-describedby'],'contextTooltip');
  assert.equal(d.tip.style.left,'428px');assert.equal(d.tip.style.top,'372px');
  d.handlers.keydown({key:'Escape'});assert.equal(d.tip.hidden,true);assert.equal(d.attrs['aria-describedby'],undefined);
  d.handlers.focusin({target:d.b});assert.equal(d.tip.hidden,false);
  d.handlers.click({target:{closest:()=>null}});assert.equal(d.tip.hidden,true);
  assert.equal(before,vm.runInContext('JSON.stringify(state)',c));
  assert.match(html,/<div id="contextTooltip"[^>]*role="tooltip" hidden/);
  assert.doesNotMatch(html,/tokenInfoDialog/);
});

test('click pins and toggles help without activating a details summary',async()=>{
  const c=model(),d=helpDom(c);let prevented=false;
  c.event={preventDefault(){prevented=true;},target:d.b};
  await vm.runInContext('handleClick(event)',c);
  assert.equal(d.tip.hidden,false);assert.equal(prevented,true);assert.equal(vm.runInContext('helpPinned',c),true);
  await vm.runInContext('handleClick(event)',c);assert.equal(d.tip.hidden,true);
});

test('pointer departure dismisses help and hovering the explanation keeps it readable',()=>{
  const c=model(),d=helpDom(c);let pending;
  c.setTimeout=fn=>{pending=fn;return 1;};c.clearTimeout=()=>{pending=null;};
  vm.runInContext('initHelp()',c);
  d.handlers.pointerover({target:d.b});d.handlers.pointerout({target:d.b});assert.equal(typeof pending,'function');
  d.tip.closest=()=>null;d.handlers.pointerover({target:d.tip});assert.equal(pending,null);assert.equal(d.tip.hidden,false);
  d.handlers.pointerout({target:d.tip});pending();assert.equal(d.tip.hidden,true);
});

test('routine explanations move to help across setup, automation, and settings',()=>{
  const c=model();
  for(const [expression,explanation] of [
    ["renderStep('sources')",'新規作成した回答先は自動登録済み'],
    ['renderFields()','差し込み名は半角英数字'],
    ['renderTemplate()','本文のカーソル位置に挿入'],
    ['renderAutomation()','指定した1時間の間に実行'],
    ['renderSettings()','名簿・送信履歴・自動実行の有効状態'],
    ["renderForms('prepare')",'採点・コメントなど必要な列だけ追加'],
    ["renderForms('publish')",'Classroomの「授業」に資料を即時投稿'],
  ]){
    const markup=vm.runInContext(expression,c);
    assert.ok(markup.includes(explanation),expression+' retains explanation');
    assert.ok(!visible(markup).includes(explanation),expression+' moves explanation behind help');
    assert.match(markup,/data-help-title=/);
  }
});

test('errors, pending drafts, operational counts and destructive warnings stay visible',()=>{
  const c=model();
  vm.runInContext(`state.drafts.automation={payload:{},revision:'r'};state.data.counts={evalRows:2,legacyPending:1,pendingSend:3,reviewSend:1};state.formsError='通信エラー';`,c);
  assert.match(visible(vm.runInContext('renderSettings()',c)),/時間帯の入力は未確定/);
  assert.match(visible(vm.runInContext("renderForms('prepare')",c)),/通信エラー/);
  assert.match(visible(vm.runInContext('renderMaintenance()',c)),/移行前にクリアしないでください/);
  assert.match(visible(vm.runInContext('renderManual()',c)),/未送信|確認待ち/);
  assert.match(html,/id="confirmDialog"/);assert.match(html,/対象と影響を確認しました/);
});
