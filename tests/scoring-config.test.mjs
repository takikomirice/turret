import test from 'node:test';
import assert from 'node:assert/strict';
import {scoringEnvironment} from './helpers/scoring-environment.mjs';

test('old scoring JSON migrates columns without changing turret connection and return settings',()=>{
 const e=scoringEnvironment(),before=e.props.get('APP_CONFIG');
 e.c.scoringSetConfig({scoreCols:['D','','','',''],displayCols:['G'],commentCol:'E'});
 const old=JSON.parse(JSON.stringify(e.c.scoringGetConfig()));delete old.lessonDateHeader;old.sheetName='別のスプシの回答';
 const result=e.c.scoringImportConfig({config:old});
 assert.equal(result.sheetName,'');assert.equal(result.scoreCols[0],'D');assert.equal(e.props.get('APP_CONFIG'),before);
 assert.equal(e.props.has('TARGET_SPREADSHEET_ID'),false);
 assert.throws(()=>e.c.scoringImportConfig({config:{...old,scoreCols:['D','D','','','']}}),/重複/);
 assert.equal(e.props.get('APP_CONFIG'),before);
});

test('rule initialization and pack import use one common lock and preserve existing templates',()=>{
 const e=scoringEnvironment();e.c.scoringEnsureRuleSheets();assert.equal(e.sheets.has('採点テンプレ'),true);assert.equal(e.isLocked(),false);
 const pack={meta:{type:'turret-grading-template-pack',version:1},templates:[{name:'テスト',enabled:true,domains:[[1,2,3],[],[],[],[]],rules:[{conditions:['>=3','','','',''],outputs:['よくできました','',''],enabled:true,visible:true}]}]};
 e.c.scoringImportTemplatePack(pack);const exported=e.c.scoringExportTemplatePack();assert.equal(exported.templates[0].rules[0].outputs[0],'よくできました');
 assert.throws(()=>e.c.scoringImportTemplatePack(pack),/同名/);assert.equal(e.sheets.get('採点テンプレ').rows.length,2);assert.equal(e.isLocked(),false);
 assert.throws(()=>e.c.scoringImportTemplatePack({meta:{type:'scoring-tool-template-pack'},templates:[]}),/旧形式/);
});

test('invalid columns and nonintegral row configuration cannot replace a valid setup',()=>{
 const e=scoringEnvironment();e.c.scoringSetConfig({scoreCols:['D','','','','']});const before=e.props.get('TURRET_SCORING_CONFIG');
 for(const patch of [{scoreCols:['A1']},{headerRow:1.5},{startRow:-1},{lessonDateHeader:{bad:true}}])assert.throws(()=>e.c.scoringSetConfig({...e.c.scoringGetConfig(),...patch}));
 assert.equal(e.props.get('TURRET_SCORING_CONFIG'),before);
});

test('missing selected headers and identity/status output columns are rejected before data writes',()=>{
 const e=scoringEnvironment();
 for(const config of [{scoreCols:['A']},{scoreCols:['C']},{scoreCols:['D'],scoreHeaders:['存在しない列']}]){
  e.c.scoringSetConfig(config);assert.throws(()=>e.c.scoringSaveRows({target:e.target,rows:[{rowNumber:2,scores:[3,'','','',''],comment:''}]}),/列|ヘッダ/);
 }
 assert.equal(e.answer.rows[1][0],'student@example.com');assert.equal(e.answer.rows[1][2],'');
});

test('a draft made before output columns change is rejected without writing to the new column',()=>{
 const e=scoringEnvironment();e.c.scoringSetConfig({scoreCols:['D','','','',''],commentCol:'E'});
 const before=e.c.scoringGetInitData({target:e.target});
 e.c.scoringSetConfig({scoreCols:['E','','','',''],commentCol:''});
 const result=e.c.scoringSaveRows({target:e.target,rows:[{rowNumber:2,scores:[3,'','','',''],comment:'',configRevision:before.configRevision}]});
 assert.equal(result.ok,false);assert.equal(e.answer.rows[1][e.answer.rows[0].indexOf('講評')],'');
});

test('sparse score slots preserve their slot number across configuration saves',()=>{
 const e=scoringEnvironment();const cfg=e.c.scoringSetConfig({scoreCols:['','D','','','']});
 assert.deepEqual(Array.from(cfg.scoreCols),['','D','','','']);
});

test('corrupt scoring settings fail closed instead of reverting grading columns silently',()=>{
 const e=scoringEnvironment();e.props.set('TURRET_SCORING_CONFIG','{broken');
 assert.throws(()=>e.c.scoringGetConfig(),/採点設定|読み取/);
});

test('a backup made with no selected template can be restored',()=>{
 const e=scoringEnvironment();e.c.scoringSetConfig({scoreCols:['D','','','','']});
 const config=JSON.parse(JSON.stringify(e.c.scoringGetConfig()));config.scoreCols=['E','','','',''];
 e.c.scoringImportConfig({config});const restored=e.c.scoringRestoreLatestConfigBackup();
 assert.equal(restored.scoreCols[0],'D');assert.equal(restored.ruleTemplateId,'');
});
