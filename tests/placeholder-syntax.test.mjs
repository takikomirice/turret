import test from 'node:test';
import assert from 'node:assert/strict';
import {adminEnvironment, plain} from './helpers/admin-environment.mjs';

test('curly reply placeholders preserve legacy rendering and blank section behavior',()=>{
 const {c}=adminEnvironment(), config=c.getConfig_();
 const template='評価：{score}\n\n【コメント】\n{comment}\n\n{student_name} / {class_name}';
 const context={score:'A',comment:'',student_name:'生徒',class_name:'化学'};
 const legacy=template.replace(/\{(\w+)\}/g,'＜$1＞');
 assert.equal(c.renderTemplateWithSections_(template,context,config),'評価：A\n\n生徒 / 化学');
 assert.equal(c.renderTemplateWithSections_(template,context,config),c.renderTemplateWithSections_(legacy,context,config));
});

test('curly tokens coexist with legacy tokens and explicit double-brace sections',()=>{
 const {c}=adminEnvironment(),config=c.getConfig_();
 const template='{{#score}}{score}/＜score＞/<score>/{{score}}{{/score}}\n{student_name}';
 assert.equal(c.renderTemplateWithSections_(template,{score:'A',student_name:'生徒'},config),'A/A/A/A\n生徒');
 assert.equal(c.renderTemplateWithSections_(template,{score:'',student_name:'生徒'},config),'生徒');
 assert.equal(c.renderTemplateWithSections_('{ score }{score}',{score:'A'},config),'AA');
 assert.equal(c.renderTemplateWithSections_('{"value": 1} / {score＞ / ＜score}',{},config),'{"value": 1} / {score＞ / ＜score}');
});

test('preview reports unknown curly tokens without changing persisted old templates',()=>{
 const {c}=adminEnvironment(),config=c.getConfig_();
 config.messageTemplate='{student_name}：{score}\n{missing} ＜missing＞ {{missing}}';
 const preview=c.previewReplyTemplate(config);
 assert.match(preview.text,/山田 太郎（例）：/);
 assert.deepEqual(plain(preview.warnings),['差し込み「missing」に対応する項目がありません。']);
 assert.equal(c.getConfig_().messageTemplate,'評価：＜score＞');
});
