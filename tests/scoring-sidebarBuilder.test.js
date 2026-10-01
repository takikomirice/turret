const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm');
function editor(){
 const elements=new Map(),calls=[],events={};
 function element(){return {value:'',children:[],style:{},addEventListener(k,f){this['on'+k]=f;},setAttribute(){},replaceChildren(){this.children=[];},appendChild(e){this.children.push(e);},get options(){return this.children;},focus(){},select(){},showModal(){this.open=true;},close(){this.open=false;}};}
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
test('AI help carries current score domains and independent output matching rules',()=>{
 const p=editor();p.e.get('newTemplate').onclick();const text=p.h.aiText(p.h.getData().templates[0]);assert.match(text,/カンマ区切り8列のCSV/);assert.match(text,/枠1 = 1,2,3/);assert.match(text,/出力1〜3それぞれ/);assert.match(text,/未入力/);for(const field of ['採点・フィードバックの目的','評価対象者','課題・提出物','条件・前提','採点基準','文言を読む相手','口調・長さ','枠1で評価する内容','出力3に入れる内容'])assert.ok(text.includes(field));assert.ok((text.match(/「　　　」/g)||[]).length>=17);
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
