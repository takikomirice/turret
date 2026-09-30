import {test} from 'node:test';
import assert from 'node:assert/strict';
import {harness,sheet,plain,completeConfig} from './backend-harness.mjs';
test('new configuration is vanilla and explicit empty values stay empty',()=>{const {context:c}=harness();const d=plain(c.buildDefaultConfig_());for(const key of ['emailHeader','studentNameHeader','formStatusHeader','scoreSourceHeader','messageTemplate'])assert.equal(d[key],'');for(const key of ['fields','formSources','reminderTo','formSheetNamePrefix'])assert.deepEqual(d[key],[]);assert.deepEqual(plain(c.normalizeAppConfig_({fields:[]})).fields,[]);});
test('legacy saved config gets missing historical defaults but preserves empty fields',()=>{const {context:c,properties:p}=harness();p.set('APP_CONFIG',JSON.stringify({messageTemplate:'legacy'}));assert.equal(c.getConfig_().emailHeader,'メールアドレス');assert.equal(c.getConfig_().fields.length,5);p.set('APP_CONFIG',JSON.stringify({fields:[],emailHeader:'',messageTemplate:''}));assert.equal(c.getConfig_().fields.length,0);assert.equal(c.getConfig_().emailHeader,'');});
test('body-only save keeps evaluation and send rows and returns revision',()=>{const h=harness(),c=h.context,config=completeConfig(c);h.properties.set('APP_CONFIG',JSON.stringify(config));for(const[name,headers]of [['評価データ',c.getConfiguredEvalHeaders_(config)],['送信シート',c.getConfiguredSendHeaders_(config)]])h.sheets.set(name,sheet([plain(headers),Array(headers.length).fill('keep')]));const before=plain([...h.sheets].map(([n,s])=>[n,s.rows]));const out=c.saveSetupSection('template',{messageTemplate:'New ＜grade＞'},c.getConfigRevision_(config));assert.equal(out.ok,true);assert.ok(out.revision);for(const[n,r]of before)assert.deepEqual(h.sheets.get(n).rows,r);});
test('schema conflict rejects before config writes',()=>{const h=harness(),c=h.context,config=completeConfig(c);h.properties.set('APP_CONFIG',JSON.stringify(config));h.sheets.set('評価データ',sheet([plain(c.getConfiguredEvalHeaders_(config)),['keep']]));const before=h.properties.get('APP_CONFIG');assert.throws(()=>c.saveSetupSection('fields',{fields:[]},c.getConfigRevision_(config)),/保守|構成|列/);assert.equal(h.properties.get('APP_CONFIG'),before);assert.equal(h.sheets.has('Config'),false);});
test('stale revision catches direct settings sheet edits',()=>{const h=harness(),c=h.context,config=completeConfig(c);h.properties.set('APP_CONFIG',JSON.stringify(config));const rev=c.getConfigRevision_(config);h.sheets.set('設定シート',sheet([['KEY','VALUE'],['FORM_SHEET_PREFIX','Changed']]));assert.throws(()=>c.saveSetupSection('template',{messageTemplate:'new'},rev),/変更|再読込/);});
test('incomplete draft saves but runtime refuses it',()=>{const {context:c}=harness();const out=c.saveSetupSection('template',{messageTemplate:''},c.getConfigRevision_(c.getConfig_()));assert.equal(out.ok,true);assert.throws(()=>c.validateAppConfig_(out.config));});
for(const patch of [{fields:[{key:'student_name',sourceHeader:'X',type:'text'}]},{fields:[{key:'x',sourceHeader:'X',type:'invalid'}]},{reminderTo:['broken']},{formSources:['https://evil.invalid/spreadsheets/d/abcdefghijklmnopqrstuv/edit']},{fields:'wrong'},{emailHeader:42}])test('invalid draft rejects '+JSON.stringify(patch),()=>{const {context:c}=harness();assert.throws(()=>c.applyConfigDraft_({...plain(c.buildDefaultConfig_()),...patch},c.getConfigRevision_(c.getConfig_())));});
for(const failure of ['APP_CONFIG_BACKUP','APP_CONFIG','APP_CONFIG_STORAGE'])test('internal save failure rolls all properties back: '+failure,()=>{
 const h=harness(),c=h.context,old=completeConfig(c);c.saveConfig_(old);
 const before=plain([...h.properties]);const set=h.props.setProperty;let once=true;
 h.props.setProperty=(k,v)=>{set(k,v);if(k===failure&&once){once=false;throw Error('injected property failure');}};
 assert.throws(()=>c.saveConfig_({...plain(old),messageTemplate:'new'}),/injected/);
 assert.deepEqual([...h.properties],before);assert.equal(h.sheets.size,0);
});

test('full legacy dialog save preserves populated queues even with old reset confirmation', () => {
  const h = harness(), c = h.context, config = completeConfig(c);
  h.properties.set('APP_CONFIG', JSON.stringify(config));
  const headers = plain(c.getConfiguredSendHeaders_(config));
  const original = [headers, Array(headers.length).fill('keep')];
  h.sheets.set('送信シート', sheet(original));
  assert.equal(c.saveSettingsFromDialog({ messageTemplate: 'new' }, true).ok, true);
  assert.deepEqual(h.sheets.get('送信シート').rows, original);
});

test('failed rollback is explicit and prevents pipeline use of a mixed configuration', () => {
  const h = harness(), c = h.context, config = completeConfig(c);
  h.properties.set('APP_CONFIG', JSON.stringify(config));
  h.sheets.set('Config', sheet([[JSON.stringify(config)]]));
  const set = h.props.setProperty;
  h.props.setProperty = (key,value) => { if(key === 'APP_CONFIG')throw Error('save and rollback denied');set(key,value); };
  assert.throws(() => c.saveConfig_({ ...plain(config), messageTemplate: 'new' }), /復旧も完了/);
  assert.ok(h.properties.has('APP_CONFIG_SAVE_ERROR'));
  assert.throws(() => c.getConfig_(), /復旧が未完了/);
});

test('invalid draft field is rejected without silently filtering or inventing a key', () => {
  const h = harness(), c = h.context;
  for (const field of [{ key: 'x', sourceHeader: '' }, { key: '', sourceHeader: 'X' }, { key: 'constructor', sourceHeader: 'X' }]) {
    assert.throws(() => c.validateConfigDraft_({ fields: [field] }), /キー|ヘッダ/);
  }
});

test('incompatible send schema fails before updating an existing Config-only store', () => {
  const h = harness(), c = h.context, config = completeConfig(c);
  h.sheets.set('Config', sheet([[JSON.stringify(config)]]));
  const headers = plain(c.getConfiguredSendHeaders_(config));
  h.sheets.set('送信シート', sheet([headers, Array(headers.length).fill('keep')]));
  const revision = c.getConfigRevision_(config);
  assert.throws(() => c.applyConfigDraft_({ ...plain(config), replyBodyHeader: 'Different' }, revision), /構成/);
  assert.equal(h.properties.has('APP_CONFIG'), false);
  assert.equal(h.sheets.get('Config').rows[0][0], JSON.stringify(config));
});

test('prototype member field keys cannot be accepted in a draft and rejected only at runtime', () => {
  const { context: c } = harness();
  assert.throws(() => c.validateConfigDraft_({ fields: [{ key: 'toString', sourceHeader: 'Value' }] }), /予約|重複/);
});

test('legacy connection course IDs survive save and returned revision matches the stored effective config', () => {
  const h = harness(), c = h.context, config = plain(completeConfig(c));
  config.formSources[0].courseId = 'course-a';
  h.properties.set('APP_CONFIG', JSON.stringify(config));
  const out = c.saveSetupSection('template', { messageTemplate: 'updated' }, c.getConfigRevision_(config));
  assert.equal(c.getConfig_().formSources[0].courseId, 'course-a');
  assert.equal(out.revision, c.getConfigRevision_(c.getConfig_()));
});

for (const character of ['a', 'あ', '😀']) {
  test('settings byte limit rejects ' + character + ' before accessing any store', () => {
    const h = harness(), c = h.context;
    const config = { ...plain(completeConfig(c)), messageTemplate: character.repeat(8001 / Buffer.byteLength(character) | 0) };
    const byteLength = Buffer.byteLength(JSON.stringify(config), 'utf8');
    assert.ok(byteLength > 8000);
    c.SpreadsheetApp.getActiveSpreadsheet = () => { throw Error('unexpected sheet access'); };
    c.PropertiesService.getScriptProperties = () => { throw Error('unexpected property access'); };
    assert.throws(() => c.saveConfig_(config), new RegExp(byteLength + ' バイト'));
  });
}

test('UTF-8 count handles ASCII, BMP, surrogate pairs and unmatched surrogates', () => {
  const { context: c } = harness();
  for (const value of ['', 'ASCII', '日本語', '😀', '\ud800', '\udc00', 'a😀あ\ud800b']) {
    assert.equal(c.utf8ByteLength_(value), Buffer.byteLength(value, 'utf8'));
  }
});

test('exactly 8000 UTF-8 bytes remain saveable and 8001 are rejected', () => {
  const h = harness(), c = h.context, config = plain(completeConfig(c));
  config.messageTemplate = '';
  const remaining = 8000 - Buffer.byteLength(JSON.stringify(config), 'utf8');
  config.messageTemplate = 'a'.repeat(remaining);
  assert.equal(Buffer.byteLength(JSON.stringify(config), 'utf8'), 8000);
  assert.doesNotThrow(() => c.saveConfig_(config));
  config.messageTemplate += 'a';
  assert.throws(() => c.saveConfig_(config), /8001 バイト/);
});

test('all setup sections save atomically under one lock without touching queues or automation', () => {
  const h = harness(), c = h.context, config = plain(completeConfig(c));
  h.properties.set('APP_CONFIG', JSON.stringify(config));
  h.properties.set('TURRET_AUTOMATION_SCHEDULE', 'unchanged schedule');
  h.properties.set('TURRET_AUTOMATION_OWNER', 'unchanged owner');
  const headers = plain(c.getConfiguredSendHeaders_(config));
  const original = [headers, Array(headers.length).fill('keep')];
  h.sheets.set('送信シート', sheet(original));
  let locks = 0, releases = 0;
  c.LockService.getScriptLock = () => ({ tryLock: () => (++locks, true), releaseLock: () => releases++ });
  const out = c.saveSetupSections({ sources: { reminderTo: [' teacher@example.com '] }, fields: { gradeScale: [{ from: '1', to: 'A' }] }, template: { messageTemplate: 'Updated ＜grade＞' } }, c.getConfigRevision_(config));
  assert.equal(out.config.messageTemplate, 'Updated ＜grade＞');
  assert.deepEqual(plain(out.config.reminderTo), ['teacher@example.com']);
  assert.deepEqual(plain(out.config.gradeScale), [{ from: '1', to: 'A' }]);
  assert.equal(out.revision, c.getConfigRevision_(c.getConfig_()));
  assert.equal(locks, 1); assert.equal(releases, 1);
  assert.deepEqual(h.sheets.get('送信シート').rows, original);
  assert.equal(h.properties.get('TURRET_AUTOMATION_SCHEDULE'), 'unchanged schedule');
  assert.equal(h.properties.get('TURRET_AUTOMATION_OWNER'), 'unchanged owner');
});

for (const patch of [{}, null, [], { automation: {} }, { sources: { messageTemplate: 'wrong section' } }, { sources: {} , template: { messageTemplate: 123 } }, { sources: { reminderTo: ['bad'] }, template: { messageTemplate: 'must not save' } }]) {
  test('atomic settings reject invalid bundles without partial writes: ' + JSON.stringify(patch), () => {
    const h = harness(), c = h.context, config = plain(completeConfig(c));
    h.properties.set('APP_CONFIG', JSON.stringify(config));
    assert.throws(() => c.saveSetupSections(patch, c.getConfigRevision_(config)));
    assert.equal(h.properties.get('APP_CONFIG'), JSON.stringify(config));
    assert.equal(h.sheets.size, 0);
  });
}

test('atomic save rejects a stale revision and rolls every section back if persistence fails', () => {
  const h = harness(), c = h.context, config = plain(completeConfig(c));
  h.properties.set('APP_CONFIG', JSON.stringify(config));
  const changes = { sources: { reminderTo: ['teacher@example.com'] }, template: { messageTemplate: 'new' } };
  assert.throws(() => c.saveSetupSections(changes, 'stale'), /変更|再読込/);
  const original = h.props.setProperty;let once=true;
  h.props.setProperty = (key,value) => { original(key,value);if(key==='APP_CONFIG'&&once){once=false;throw Error('injected failure');} };
  assert.throws(() => c.saveSetupSections(changes, c.getConfigRevision_(config)), /injected failure/);
  assert.equal(h.properties.get('APP_CONFIG'), JSON.stringify(config));
  assert.equal(h.sheets.size, 0);
});
