import { test } from 'node:test';
import assert from 'node:assert/strict';
import { harness, plain, completeConfig } from './backend-harness.mjs';
const field = { key: 'grade', sourceHeader: 'Grade', type: 'score_grade' };

test('vanilla grade scale is empty; legacy saved configurations inherit the historical scale', () => {
  const h = harness(), c = h.context;
  assert.deepEqual(plain(c.buildDefaultConfig_().gradeScale), []);
  const old = plain(completeConfig(c)); delete old.gradeScale;
  old.fields[0].type = 'score_grade';
  h.properties.set('APP_CONFIG', JSON.stringify(old));
  assert.deepEqual(plain(c.getConfig_().gradeScale), [
    { from: '5', to: 'A' }, { from: '4', to: 'B+' }, { from: '3', to: 'B' },
    { from: '2', to: 'B-' }, { from: '1', to: 'C' }
  ]);
  h.properties.set('APP_CONFIG', JSON.stringify({ ...old, gradeScale: [] }));
  assert.deepEqual(plain(c.getConfig_().gradeScale), []);
});

test('configured grade conversion trims exact strings and retains decimal equivalence without rounding', () => {
  const { context: c } = harness();
  const config = { gradeScale: [{ from: '5', to: ' 優 ' }, { from: '未提出', to: '再提出' }, { from: '0.5', to: '半分' }, { from: '9007199254740993', to: '大きい整数' }] };
  for (const input of [5, ' 5 ', '05', '+5.0']) assert.equal(c.processConfiguredFieldValue_(field, input, '', config), '優');
  assert.equal(c.processConfiguredFieldValue_(field, ' 未提出 ', '', config), '再提出');
  assert.equal(c.processConfiguredFieldValue_(field, '.50', '', config), '半分');
  assert.equal(c.processConfiguredFieldValue_(field, '9007199254740993', '', config), '大きい整数');
  for (const value of [' unknown ', '5e0', '0x5', '9007199254740992']) assert.equal(c.processConfiguredFieldValue_(field, value, '', config), value.trim());
  assert.equal(c.processConfiguredFieldValue_(field, null, '', config), '');
  assert.equal(c.processConfiguredFieldValue_(field, 5, '', { gradeScale: [] }), '5');
});

test('saved vanilla configurations without grade fields do not acquire a legacy scale', () => {
  const h = harness(), c = h.context, config = plain(c.buildDefaultConfig_());
  delete config.gradeScale;
  h.properties.set('APP_CONFIG', JSON.stringify(config));
  assert.deepEqual(plain(c.getConfig_().gradeScale), []);
});

test('public legacy scoreToGrade and old three-argument conversion stay compatible', () => {
  const { context: c } = harness();
  assert.equal(c.scoreToGrade(2), 'B-');
  assert.equal(c.scoreToGrade('5.0'), 'A');
  assert.equal(c.scoreToGrade('unknown'), '');
  assert.equal(c.processConfiguredFieldValue_(field, 5, ''), 'A');
  assert.equal(c.processConfiguredFieldValue_({ type: 'text' }, 5, '', { gradeScale: [{ from: '5', to: 'A' }] }), '5');
});

for (const gradeScale of [null, 'bad', [{}], [{ from: 5, to: 'A' }], [{ from: '5', to: 1 }], [{ from: ' ', to: 'A' }], [{ from: '5', to: '' }], [{ from: '5', to: 'A' }, { from: '05.0', to: 'B' }], [{ from: '-0', to: 'A' }, { from: '0.0', to: 'B' }], [{ from: ' text ', to: 'A' }, { from: 'text', to: 'B' }]]) {
  test('grade draft rejects malformed/ambiguous map ' + JSON.stringify(gradeScale), () => {
    const { context: c } = harness();
    assert.throws(() => c.validateConfigDraft_({ gradeScale }), /評価|変換|配列|文字列|重複/);
  });
}

test('empty grade scale is allowed in drafts and runtime text fields but blocks score_grade execution', () => {
  const { context: c } = harness(), config = completeConfig(c);
  config.gradeScale = [];
  assert.doesNotThrow(() => c.validateAppConfig_(config));
  config.fields[0].type = 'score_grade';
  assert.doesNotThrow(() => c.validateConfigDraft_(config));
  assert.throws(() => c.validateAppConfig_(config), /評価|変換/);
  config.gradeScale = [{ from: '5', to: 'A' }];
  assert.doesNotThrow(() => c.validateAppConfig_(config));
});

test('fields setup saves a normalized grade scale and revision covers the mapping', () => {
  const h = harness(), c = h.context, config = completeConfig(c);
  h.properties.set('APP_CONFIG', JSON.stringify(config));
  const revision = c.getConfigRevision_(config);
  const result = c.saveSetupSection('fields', { gradeScale: [{ from: ' 5 ', to: ' A ' }] }, revision);
  assert.deepEqual(plain(result.config.gradeScale), [{ from: '5', to: 'A' }]);
  assert.notEqual(result.revision, revision);
  assert.deepEqual(plain(c.getConfig_().gradeScale), [{ from: '5', to: 'A' }]);
});

test('each grade field overrides or inherits the shared map without changing other fields', () => {
  const { context: c } = harness();
  const config = { gradeScale: [{ from: '1', to: 'C' }] };
  const own = { ...field, gradeScale: [{ from: '1', to: '×' }] };
  assert.equal(c.processConfiguredFieldValue_(own, '01.0', '', config), '×');
  assert.equal(c.processConfiguredFieldValue_(field, 1, '', config), 'C');
  assert.equal(c.processConfiguredFieldValue_({ ...field, gradeScale: [] }, 1, '', config), '1');
  assert.deepEqual(config.gradeScale, [{ from: '1', to: 'C' }]);
});

test('field normalization preserves explicit empty maps and absence, and revisions cover field mappings', () => {
  const { context: c } = harness();
  assert.equal(Object.hasOwn(c.normalizeFieldConfig_(field, 0), 'gradeScale'), false);
  assert.deepEqual(plain(c.normalizeFieldConfig_({ ...field, gradeScale: [] }, 0).gradeScale), []);
  const config = plain(completeConfig(c));
  const before = c.getConfigRevision_(config);
  config.fields[0].gradeScale = [{ from: ' 1 ', to: ' × ' }];
  assert.deepEqual(plain(c.normalizeAppConfig_(config).fields[0].gradeScale), [{ from: '1', to: '×' }]);
  assert.notEqual(c.getConfigRevision_(config), before);
});

test('field scales survive persistence without generating a display-only settings sheet', () => {
  const h = harness(), c = h.context, config = plain(completeConfig(c));
  h.properties.set('APP_CONFIG', JSON.stringify(config));
  const fields = [{ ...config.fields[0], type: 'score_grade', gradeScale: [{ from: ' 1 ', to: ' × ' }] }];
  c.saveSetupSection('fields', { fields, gradeScale: [{ from: '1', to: 'C' }] }, c.getConfigRevision_(config));
  assert.deepEqual(plain(c.getConfig_().fields[0].gradeScale), [{ from: '1', to: '×' }]);
  assert.equal(h.sheets.has('設定(編集不可)'), false);
  assert.deepEqual(JSON.parse(h.properties.get('APP_CONFIG_BACKUP')).fields[0].gradeScale, [{ from:'1',to:'×' }]);
});

test('runtime requires an effective scale per grade field but allows different maps with no common scale', () => {
  const { context: c } = harness(), config = plain(completeConfig(c));
  config.fields[0] = { ...config.fields[0], type: 'score_grade', gradeScale: [{ from: '1', to: '×' }] };
  config.fields.push({ key: 'second', sourceHeader: 'Second', type: 'score_grade', gradeScale: [{ from: '1', to: 'C' }] });
  assert.doesNotThrow(() => c.validateAppConfig_(config));
  config.fields[1].gradeScale = [];
  config.gradeScale = [{ from: '1', to: 'shared' }];
  assert.doesNotThrow(() => c.validateConfigDraft_(config));
  assert.throws(() => c.validateAppConfig_(config), /Second|second/);
  delete config.fields[1].gradeScale;
  assert.doesNotThrow(() => c.validateAppConfig_(config));
});

for (const gradeScale of [null, 'bad', [{}], [{ from: 1, to: 'A' }], [{ from: '1', to: '' }], [{ from: '1', to: 'A' }, { from: '01.0', to: 'B' }]]) {
  test('field grade scale rejects malformed map ' + JSON.stringify(gradeScale), () => {
    const { context: c } = harness();
    assert.throws(() => c.validateConfigDraft_({ fields: [{ ...field, gradeScale }] }), /評価|変換|配列|文字列|重複/);
  });
}
