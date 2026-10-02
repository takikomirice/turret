const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm');
function editor(){
 const elements=new Map(),calls=[],events={};
 function element(tag='div'){return {tagName:tag.toUpperCase(),value:'',children:[],style:{},dataset:{},attributes:{},selectionStart:0,selectionEnd:0,addEventListener(k,f){this['on'+k]=f;},setAttribute(k,v){this.attributes[k]=String(v);},getAttribute(k){return this.attributes[k];},removeAttribute(k){delete this.attributes[k];},replaceChildren(){this.children=[];},appendChild(e){e.parentNode=this;this.children.push(e);},get options(){return this.children;},focus(){document.activeElement=this;},select(){this.selectionStart=0;this.selectionEnd=this.value.length;},setSelectionRange(a,b){this.selectionStart=a;this.selectionEnd=b;},showModal(){this.open=true;},close(){this.open=false;}};}
 const document={getElementById(id){if(!elements.has(id))elements.set(id,element());return elements.get(id);},createElement:element,addEventListener(k,f){events['document:'+k]=f;}};
 const window={__TEMPLATE_EDITOR_TEST__:true,confirm:()=>true,addEventListener(k,f){events[k]=f;}};
 const c=vm.createContext({document,window,parent:{postMessage(){}},navigator:{},console,TextDecoder,google:{script:{get run(){const chain={withSuccessHandler(f){this.success=f;return this;},withFailureHandler(f){this.failure=f;return this;}};return new Proxy(chain,{get:(o,k)=>k in o?o[k]:p=>calls.push({method:k,payload:p,success:o.success,failure:o.failure})});}}}});
 const html=fs.readFileSync('Scoring.html','utf8').match(/<template id="ruleBuilderTemplate">([\s\S]*?)<\/template>/)[1];for(const m of html.matchAll(/id="([^"]+)"/g))document.getElementById(m[1]);vm.runInContext(html.match(/<script>([\s\S]*?)<\/script>/)[1],c);
 calls[0].success({revision:'v1',templates:[],url:'https://docs.google.com/test'});
 return {h:window.templateEditor,e:elements,calls,events};
}
test('editor table round-trips multiline cells and zero, and rejects misaligned AI output',()=>{
 const p=editor(),rules=[{conditions:['>=2','<=3','','',''],outputs:['0','改行\nと"引用"','タブ\t文言']}];
 const result=p.h.parseTable(p.h.tableText({rules}));assert.deepEqual(Array.from(result[0].outputs),rules[0].outputs);
 assert.equal(result[0].conditions[0],'>=2');assert.throws(()=>p.h.parseTable('>=2,>=2,よくできています'),/8列/);
 assert.throws(()=>p.h.parseTable('"閉じていない'),/引用符/);
});
test('new, duplicate and reordered rules preserve entered outputs',()=>{
 const p=editor();p.e.get('newTemplate').onclick();p.e.get('addRule').onclick();
 const t=p.h.getData().templates[0];t.rules[0].outputs[0]='first';t.rules[1].outputs[0]='second';
 const actionCell=p.e.get('rules').children[1].children.at(-1);actionCell.children[0].onclick();assert.equal(t.rules[0].outputs[0],'second');
 p.e.get('duplicateTemplate').onclick();assert.equal(p.h.getData().templates.length,2);assert.equal(p.h.getData().templates[1].rules[0].outputs[0],'second');assert.notEqual(p.h.getData().templates[1].name,t.name);
});
test('conflicting save keeps every draft and unlocks the editor',()=>{
 const p=editor();p.e.get('newTemplate').onclick();p.h.getData().templates[0].rules[0].outputs[0]='手入力';p.e.get('save').onclick();
 assert.equal(p.calls.at(-1).payload.revision,'v1');assert.equal(p.e.get('work').disabled,true);p.calls.at(-1).failure(Error('変更されています'));
 assert.equal(p.h.getData().templates[0].rules[0].outputs[0],'手入力');assert.equal(p.h.isDirty(),true);assert.equal(p.e.get('work').disabled,false);
 let prevented=false;p.events.beforeunload({preventDefault(){prevented=true;}});assert.equal(prevented,true);
});
test('AI help carries current input policy and independent output matching rules',()=>{
 const p=editor();p.e.get('newTemplate').onclick();const t=p.h.getData().templates[0];t.scorePolicy[0]={mode:'list',values:[-1,0,2.5]};const text=p.h.aiText(t);assert.match(text,/カンマ区切り8列のCSV/);assert.match(text,/枠1 = -1,0,2.5/);assert.match(text,/枠2 = 制限なし/);assert.match(text,/出力1〜3それぞれ/);assert.match(text,/未入力/);assert.match(text,/一致しない.*空欄/);for(const field of ['採点・フィードバックの目的','評価対象者','課題・提出物','条件・前提','採点基準','文言を読む相手','口調・長さ','枠1で評価する内容','出力3に入れる内容'])assert.ok(text.includes(field));assert.ok((text.match(/「　　　」/g)||[]).length>=17);
});

const plain=value=>JSON.parse(JSON.stringify(value));
function cell(p,row,col){return p.e.get('rules').children[row].children[col+1].children[0];}
function key(el,key,extra={}){const event={key,prevented:false,preventDefault(){this.prevented=true;},...extra};el.onkeydown(event);return event;}
function type(el,value){el.value=value;el.oninput({target:el});}

test('new templates default to five unrestricted policies and independent output visibility',()=>{
 const p=editor();p.e.get('newTemplate').onclick();const t=p.h.getData().templates[0];
 assert.deepEqual(plain(t.scorePolicy),Array.from({length:5},()=>({mode:'any'})));assert.equal(t.domains,undefined);
 assert.deepEqual(plain(t.rules[0].outputVisible),[true,true,true]);assert.equal(t.rules[0].visible,undefined);
 const tr=p.e.get('rules').children[0];assert.equal(tr.children.length,14);
 const show2=tr.children[10].children[0];show2.checked=false;show2.onchange();assert.deepEqual(plain(t.rules[0].outputVisible),[true,false,true]);assert.equal(t.rules[0].enabled,true);
 p.e.get('duplicateTemplate').onclick();assert.deepEqual(plain(p.h.getData().templates[1].scorePolicy),plain(t.scorePolicy));
});

test('selection arrows and Tab navigate while edit arrows and multiline Enter stay native',()=>{
 const p=editor();p.e.get('newTemplate').onclick();p.e.get('addRule').onclick();p.h.selectCell(0,5);
 assert.equal(key(cell(p,0,5),'ArrowDown').prevented,true);assert.deepEqual(plain(p.h.getSelection()),{row:1,col:5,anchorRow:1,anchorCol:5});
 assert.equal(key(cell(p,1,5),'F2').prevented,true);assert.equal(key(cell(p,1,5),'ArrowLeft').prevented,false);assert.equal(key(cell(p,1,5),'Enter').prevented,false);
 type(cell(p,1,5),'改行\n文言');key(cell(p,1,5),'Tab');assert.equal(p.h.getData().templates[0].rules[1].outputs[0],'改行\n文言');assert.equal(p.h.getSelection().col,6);
 key(cell(p,1,6),'Enter');type(cell(p,1,6),'取消');key(cell(p,1,6),'Escape');assert.equal(p.h.getData().templates[0].rules[1].outputs[1],'');assert.equal(cell(p,1,6).value,'');
});

test('direct typing replaces the selection, one edit has one undo, and IME does not navigate',()=>{
 const p=editor();p.e.get('newTemplate').onclick();p.h.selectCell(0,5);p.h.pasteCells('以前');
 let el=cell(p,0,5);assert.equal(key(el,'新').prevented,false);assert.equal(el.selectionStart,0);assert.equal(el.selectionEnd,2);type(el,'新');key(el,'Tab');
 assert.equal(p.h.getData().templates[0].rules[0].outputs[0],'新');p.h.undo();assert.equal(p.h.getData().templates[0].rules[0].outputs[0],'以前');p.h.redo();assert.equal(p.h.getData().templates[0].rules[0].outputs[0],'新');
 p.h.selectCell(0,5);el=cell(p,0,5);el.oncompositionstart();assert.equal(key(el,'Enter',{isComposing:true,keyCode:229}).prevented,false);assert.equal(key(el,'ArrowRight',{isComposing:true}).prevented,false);type(el,'日本語');el.oncompositionend();key(el,'Tab');assert.equal(p.h.getData().templates[0].rules[0].outputs[0],'日本語');
});

test('rectangular TSV paste preserves blanks, zero and quotes and rejects partial invalid operations',()=>{
 const p=editor();p.e.get('newTemplate').onclick();p.h.selectCell(0,5);
 p.h.pasteCells('"a\tb\n""引用"""\t0\t\nsecond\t\tlast\n');const t=p.h.getData().templates[0];assert.equal(t.rules.length,2);assert.deepEqual(plain(t.rules[0].outputs),['a\tb\n"引用"','0','']);
 const before=JSON.stringify(t.rules);assert.throws(()=>p.h.pasteCells('x\ty\tz\tw'),/8列|右端/);assert.equal(JSON.stringify(t.rules),before);
 p.h.selectCell(0,0);assert.throws(()=>p.h.pasteCells('1\t>=2\n3\t不正'),/条件/);assert.equal(JSON.stringify(t.rules),before);
 assert.throws(()=>p.h.pasteCells('1\t2\n3'),/列数|矩形/);assert.equal(JSON.stringify(t.rules),before);
 p.h.undo();assert.equal(t.rules.length,1);assert.deepEqual(plain(t.rules[0].outputs),['','','']);p.h.redo();assert.equal(t.rules.length,2);
});

test('Shift arrows select a rectangle, copy uses TSV, and range deletion is one undo',()=>{
 const p=editor();p.e.get('newTemplate').onclick();p.h.selectCell(0,5);p.h.pasteCells('one\t0\ntwo\tthree');p.h.selectCell(0,5);
 key(cell(p,0,5),'ArrowRight',{shiftKey:true});key(cell(p,0,6),'ArrowDown',{shiftKey:true});assert.equal(p.h.copyCells(),'one\t0\ntwo\tthree');
 key(cell(p,1,6),'Delete');assert.deepEqual(plain(p.h.getData().templates[0].rules.map(r=>r.outputs)),[['','',''],['','','']]);p.h.undo();assert.equal(p.h.copyCells(),'one\t0\ntwo\tthree');
});

test('row operations retain focus on moved data and undo restores the original selection',()=>{
 const p=editor();p.e.get('newTemplate').onclick();p.h.selectCell(0,5);p.h.pasteCells('first\nsecond');p.h.selectCell(1,5);
 p.e.get('rules').children[1].children.at(-1).children[0].onclick();assert.equal(p.h.getSelection().row,0);assert.equal(p.h.getSelection().col,5);assert.equal(p.h.getData().templates[0].rules[0].outputs[0],'second');
 p.h.undo();assert.equal(p.h.getSelection().row,1);assert.equal(p.h.getData().templates[0].rules[1].outputs[0],'second');
});

test('native clipboard events act on ranges, but remain native while editing',()=>{
 const p=editor();p.e.get('newTemplate').onclick();p.h.selectCell(0,5);
 let el=cell(p,0,5),prevented=false;el.onpaste({clipboardData:{getData:()=> '日本語\t0\n次の行\t"a\nb"'},preventDefault(){prevented=true;}});assert.equal(prevented,true);
 el=cell(p,0,5);let copied;el.oncopy({clipboardData:{setData(type,value){assert.equal(type,'text/plain');copied=value;}},preventDefault(){}});assert.equal(copied,'日本語\t0\n次の行\t"a\nb"');
 const before=JSON.stringify(p.h.getData());el.onpaste({clipboardData:{getData:()=> '1\t2\n3'},preventDefault(){}});assert.equal(JSON.stringify(p.h.getData()),before);assert.match(p.e.get('status').textContent,/列数/);
 key(el,'F2');prevented=false;el.onpaste({clipboardData:{getData:()=> 'native paste'},preventDefault(){prevented=true;}});assert.equal(prevented,false);el.oncopy({clipboardData:{setData(){assert.fail('editing copy must be native');}},preventDefault(){assert.fail('editing copy must be native');}});
});

test('input policy controls accept finite values, block invalid saves and keep duplicate drafts',()=>{
 const p=editor();p.e.get('newTemplate').onclick();let box=p.e.get('scorePolicies').children[0],mode=box.children[0].children[0],values=box.children[1];mode.value='list';mode.onchange();values.value='-1,0,2.5';values.oninput();
 assert.deepEqual(plain(p.h.getData().templates[0].scorePolicy[0]),{mode:'list',values:[-1,0,2.5]});assert.equal(values.hidden,false);
 values.value='1,不正';values.oninput();let count=p.calls.length;p.e.get('save').onclick();assert.equal(p.calls.length,count);assert.match(p.e.get('status').textContent,/数値/);
 p.e.get('duplicateTemplate').onclick();box=p.e.get('scorePolicies').children[0];assert.equal(box.children[1].value,'1,不正');p.e.get('save').onclick();assert.equal(p.calls.length,count);
 for(const text of ['NaN','Infinity','1,1','1,',''])assert.throws(()=>p.h.parsePolicy(text));assert.deepEqual(plain(p.h.parsePolicy('0 -.5 1.5')), [0,-.5,1.5]);
});

test('clipboard round-trip keeps the final blank row in a one-column range',()=>{
 const p=editor();p.e.get('newTemplate').onclick();p.e.get('addRule').onclick();p.h.selectCell(0,5);p.h.pasteCells('first');p.h.selectCell(0,5);key(cell(p,0,5),'ArrowDown',{shiftKey:true});
 const text=p.h.copyCells();assert.deepEqual(plain(p.h.readRectangle(text)),[['first'],['']]);
});

test('checkbox and policy controls keep native keys; row deletion and duplication are one undo',()=>{
 const p=editor();p.e.get('newTemplate').onclick();p.h.selectCell(0,5);p.h.pasteCells('first\nsecond');p.h.selectCell(1,5);
 const tr=p.e.get('rules').children[1];assert.equal(tr.children[9].children[0].onkeydown,undefined);assert.equal(p.e.get('scorePolicies').children[0].children[0].children[0].onkeydown,undefined);
 tr.children.at(-1).children[2].onclick();assert.equal(p.h.getData().templates[0].rules.length,3);assert.equal(p.h.getSelection().row,2);p.h.undo();assert.equal(p.h.getData().templates[0].rules.length,2);assert.equal(p.h.getSelection().row,1);
 p.e.get('rules').children[1].children.at(-1).children[3].onclick();assert.equal(p.h.getData().templates[0].rules.length,1);assert.equal(p.h.getSelection().row,0);p.h.undo();assert.equal(p.h.getData().templates[0].rules[1].outputs[0],'second');assert.equal(p.h.getSelection().row,1);
});

test('IME composition does not consume Escape, undo, arrows or the confirm Enter',()=>{
 const p=editor();p.e.get('newTemplate').onclick();p.h.selectCell(0,5);const el=cell(p,0,5);el.oncompositionstart();
 for(const name of ['Enter','Escape','ArrowLeft','Tab','z'])assert.equal(key(el,name,{ctrlKey:name==='z',isComposing:true}).prevented,false);
 type(el,'漢字');el.oncompositionend();key(el,'Escape');assert.equal(el.value,'');assert.equal(p.h.getData().templates[0].rules[0].outputs[0],'');
});

test('the undo button immediately cancels a pending first edit as one operation',()=>{
 const p=editor();p.e.get('newTemplate').onclick();p.h.selectCell(0,5);assert.equal(p.e.get('undo').disabled,true);
 const el=cell(p,0,5);key(el,'F2');type(el,'pending');assert.equal(p.e.get('undo').disabled,false);p.e.get('undo').onclick();assert.equal(p.h.getData().templates[0].rules[0].outputs[0],'');assert.equal(cell(p,0,5).value,'');assert.equal(p.e.get('redo').disabled,false);
 p.e.get('redo').onclick();assert.equal(p.h.getData().templates[0].rules[0].outputs[0],'pending');
});

test('loaded finite policy numbers in scientific notation save unchanged',()=>{
 const p=editor(),scorePolicy=[{mode:'list',values:[1e-7,1e21]},...Array.from({length:4},()=>({mode:'any'}))];
 p.calls[0].success({revision:'imported',url:'#',templates:[{name:'指数表記',scorePolicy,rules:[]}]});
 assert.equal(p.e.get('scorePolicies').children[0].children[1].value,'1e-7,1e+21');p.e.get('save').onclick();assert.equal(p.calls.at(-1).method,'scoringSaveTemplateTable');assert.deepEqual(plain(p.calls.at(-1).payload.templates[0].scorePolicy),scorePolicy);
 assert.deepEqual(plain(p.h.parsePolicy('-1e-7,2E+3')),[-1e-7,2000]);assert.throws(()=>p.h.parsePolicy('1e309'));
});
test('paste adds only complete rows and preserves zero and empty cells',()=>{
 const p=editor();p.e.get('newTemplate').onclick();p.e.get('pasteText').value='>=2\t>=2\t\t\t\tよくできています\t0\t';p.e.get('previewPaste').onclick();p.e.get('appendPaste').onclick();
 assert.equal(p.h.getData().templates[0].rules.length,1);assert.deepEqual(Array.from(p.h.getData().templates[0].rules[0].outputs),['よくできています','0','']);
 p.e.get('pasteText').value='bad';p.e.get('previewPaste').onclick();p.e.get('appendPaste').onclick();assert.equal(p.h.getData().templates[0].rules.length,1);
});

test('CSV guide is prefilled, excluded on import, and never drops headerless data',()=>{
 const p=editor(),header='枠1,枠2,枠3,枠4,枠5,出力1,出力2,出力3';
 assert.equal(p.e.get('pasteText').value,header+'\n');assert.equal(p.h.parseTable(header).length,0);
 const row='>=2,>=2,,,,よくできています,0,';
 for(const text of [row,header+'\n'+row,'\uFEFF'+header+'\r\n'+header+'\r\n'+row,'"'+header.split(',').join('","')+'"\n'+row]){
  const rules=p.h.parseTable(text);assert.equal(rules.length,1);assert.deepEqual(Array.from(rules[0].outputs),['よくできています','0','']);
 }
 assert.equal(p.h.parseTable(header+'\n'+row.replaceAll(',','\t')).length,1);
 assert.throws(()=>p.h.parseTable(header+'\n'+row+'\n'+header),/3行目.*見出し/);
});

test('CSV handles quoted comma, quotation and CRLF; reports physical error lines',()=>{
 const p=editor(),header=p.e.get('pasteText').value;
 const row='1,2,,,,"a,b\r\n""引用""",0,';
 assert.equal(p.h.parseTable(row)[0].outputs[0],'a,b\r\n"引用"');
 assert.throws(()=>p.h.parseTable(header+row+'\r\nbad'),/4行目は1列/);
 assert.throws(()=>p.h.parseTable('1,2,,,,a"b,0,'),/引用符/);
 assert.throws(()=>p.h.parseTable('1,2,,,,"a"x,0,'),/余分/);
 assert.throws(()=>p.h.parseTable('1,2,,,,a,0,,extra'),/9列/);
});

test('preview is non-destructive and changing input or target prevents stale import',()=>{
 const p=editor();p.e.get('newTemplate').onclick();const t=p.h.getData().templates[0];
 t.rules[0].outputs[0]='既存';p.e.get('pasteText').value='1,2,,,,追加,,';p.e.get('previewPaste').onclick();
 assert.equal(t.rules.length,1);assert.equal(p.e.get('appendPaste').disabled,false);
 p.e.get('pasteText').value='1,2,,,,変更,,';p.e.get('appendPaste').onclick();assert.equal(t.rules.length,1);
 p.e.get('pasteText').oninput();assert.equal(p.e.get('appendPaste').disabled,true);
 p.e.get('previewPaste').onclick();p.e.get('newTemplate').onclick();p.e.get('appendPaste').onclick();assert.equal(t.rules.length,1);
 p.e.get('previewPaste').onclick();p.e.get('appendPaste').onclick();assert.equal(p.h.getData().templates[1].rules[0].outputs[0],'変更');
 assert.match(p.e.get('pasteText').value,/^枠1,枠2/);
});

test('CSV file decoding previews without appending and ignores late file results',async()=>{
 const p=editor();p.e.get('newTemplate').onclick();const bytes=new TextEncoder().encode('1,2,,,,日本語,,');
 p.e.get('csvFile').files=[{size:bytes.length,arrayBuffer:async()=>bytes.buffer}];
 await p.e.get('csvFile').onchange();assert.match(p.e.get('pasteSummary').textContent,/CSV：1行/);assert.equal(p.h.getData().templates[0].rules[0].outputs[0],'');
 let finish;p.e.get('csvFile').files=[{size:1,arrayBuffer:()=>new Promise(r=>finish=r)}];const pending=p.e.get('csvFile').onchange();
 p.e.get('pasteText').value='手入力';p.e.get('pasteText').oninput();finish(bytes.buffer);await pending;assert.equal(p.e.get('pasteText').value,'手入力');
 p.e.get('csvFile').files=[{size:6*1024*1024}];await p.e.get('csvFile').onchange();assert.match(p.e.get('status').textContent,/5MB/);
 const invalid=Uint8Array.from([0xff]);p.e.get('csvFile').files=[{size:1,arrayBuffer:async()=>invalid.buffer}];await p.e.get('csvFile').onchange();assert.match(p.e.get('status').textContent,/文字コード/);
 const sjis=Uint8Array.from([...Buffer.from('1,2,,,,'),0x93,0xfa,0x96,0x7b,0x8c,0xea,44,44]);
 p.e.get('csvFile').files=[{size:sjis.length,arrayBuffer:async()=>sjis.buffer}];p.e.get('csvEncoding').value='shift_jis';await p.e.get('csvEncoding').onchange();assert.match(p.e.get('pasteText').value,/日本語/);
});
