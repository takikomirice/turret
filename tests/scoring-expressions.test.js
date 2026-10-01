const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

function loadClientHooks() {
  const html = fs.readFileSync(path.join(__dirname, '..', 'Scoring.html'), 'utf8');
  const match = html.match(/<script>([\s\S]*?)<\/script>/);
  assert.ok(match, 'Index.html の script が見つかりません');

  const windowObj = {
    __SCORING_TOOL_ENABLE_TEST_HOOKS__: true,
    setTimeout,
    clearTimeout,
    confirm: () => false,
    open: () => null
  };
  const documentObj = {
    readyState: 'loading',
    addEventListener: () => {},
    getElementById: () => null,
    querySelectorAll: () => []
  };
  const sandbox = {
    console,
    window: windowObj,
    document: documentObj,
    setTimeout,
    clearTimeout
  };
  vm.createContext(sandbox);
  vm.runInContext(match[1], sandbox, { filename: 'Index.html<script>' });

  assert.ok(windowObj.__scoringToolTestHooks, 'クライアント評価器のテストフックが見つかりません');
  return windowObj.__scoringToolTestHooks;
}

// Literal expectations preserve the supported language without a second production evaluator.
test('whenExpr preserves arithmetic, logical operators, precedence and missing variables', () => {
  const client=loadClientHooks();
  const vars=client.buildEvalVars(client.buildRuleVars_(['4','3','3','5','5']));
  for(const [expr,expected] of [
    ['score1 + score2 >= 7',true],
    ['score1 + score2 >= 7 and score3 >= 3',true],
    ['score1 >= 3 or score2 >= 3',true],
    ['not(score1 == 0)',true],['total >= 20',true],
    ['score1 + score2 >= 8',false],['total > 20',false],
    ['score1 >= 3 && score2 >= 3',true],['score1 >= 3 || score2 >= 3',true],
    ['score1 < 3 || score2 >= 3',true],['!(score1 == 0)',true],
    ['missingValue == null',true],['true or false and false',true],
    ['(true or false) and false',false],['score1 != 4',false],
    ['score1 <= 4 and score2 < 4',true]
  ]) {
    const result=client.evaluateWhenExprSafe(expr,vars);
    assert.equal(result.ok,true,expr);assert.equal(result.result,expected,expr);
  }
});

test('non-numeric scores count as zero in total and arithmetic',()=>{
  const client=loadClientHooks(),vars=client.buildEvalVars(client.buildRuleVars_(['3a','','','','']));
  for(const expr of ['score1 + score2 == 0','total == 0'])assert.equal(client.evaluateWhenExprSafe(expr,vars).result,true);
});

test('numeric score normalization keeps empty values absent and computes the sum',()=>{
  const client=loadClientHooks(),vars=client.buildEvalVars(client.buildRuleVars_(['3','3a','',null,2]));
  assert.deepEqual(Array.from({length:5},(_,i)=>vars['score'+(i+1)]),[3,null,null,null,2]);
  assert.equal(client.evaluateWhenExprSafe('total == 5',vars).result,true);
});

test('unterminated string conditions remain errors',()=>{
  const client=loadClientHooks();
  const result=client.evaluateWhenExprSafe('score1 == "3',{score1:3});
  assert.equal(result.ok,false);assert.match(result.error,/string|文字列|unterminated/i);
});

test('rule headers prefer message over the legacy text column',()=>{
  const client=loadClientHooks();assert.equal(client.findRuleMessageKey_(['text','message']),'message');
  assert.equal(client.findRuleMessageKey_(['text']),'text');
});

test('message expressions preserve labels, score sums, maps and literal escaping',()=>{
  const client=loadClientHooks(),vars=client.buildEvalVars(client.buildRuleVars_(['3','2','','','']));
  for(const [message,expected] of [
    ['=label1 & "：" & map(score1,"1=×,2=△,3=○")','観点1：○'],
    ['=score1+score2','5'],['通常メッセージ','通常メッセージ'],["'=式にしない",'=式にしない']
  ])assert.equal(client.evaluateRuleMessageLocal_(message,vars,['観点1','観点2','','','']),expected,message);
});

test('multiple outputs use the first matching rule per target and retain slot order',()=>{
  const client=loadClientHooks();
  const result=client.evaluateRulesLocallyMulti_([
    {target:'改善点',whenExpr:'score1 >= 3',message:'次の課題へ',enabled:1,_rowNumber:2},
    {target:'講評',whenExpr:'score1 >= 3',message:'=map(score1,"3=良好")',enabled:1,_rowNumber:3},
    {target:'講評',whenExpr:'',message:'予備コメント',enabled:1,_rowNumber:4}
  ],['target','whenExpr','message','enabled'],client.buildEvalVars(client.buildRuleVars_(['3','','','',''])),[
    {target:'講評',enabled:true},{target:'改善点',enabled:true},{target:'',enabled:false}
  ]);
  assert.deepEqual(JSON.parse(JSON.stringify(result)),{outputs:['良好','次の課題へ',''],matchedRowNumberForSlot1:3,error:null});
});
