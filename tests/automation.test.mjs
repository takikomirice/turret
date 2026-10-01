import { existsSync, readFileSync } from 'node:fs';
import vm from 'node:vm';
import { test } from 'node:test';
import assert from 'node:assert/strict';

const administration = readFileSync('Administration.gs', 'utf8');
const automationStart = administration.indexOf('const AUTOMATION_STATE_KEY_');
assert.notEqual(automationStart, -1, 'automation implementation must be present');
const code = administration.slice(automationStart);
const schedule = { importHour: 5, deliveryHour: 8, reminderHour: 16, reminderEnabled: false };
const plain = value => JSON.parse(JSON.stringify(value));

function environment() {
  const props = new Map(), users = new Map(), triggers = [], calls = [];
  let user = 'a', serial = 0, locked = false, createCount = 0;
  const faults = { errorRows: 0 };
  const store = map => ({ getProperty: key => { if (faults.read?.(key)) throw Error('read failed'); return map.get(key) ?? null; }, setProperty: (key, value) => {
    if (faults.property?.(key, value, false)) throw Error('property write failure');
    map.set(key, value);
    if (faults.property?.(key, value, true)) throw Error('post-write failure');
  } });
  const add = (handler, event = 'CLOCK') => {
    const t = { id: 'trigger-' + (++serial), owner: user, handler, event,
      getUniqueId() { return this.id; }, getHandlerFunction() { return this.handler; }, getEventType() { return this.event; } };
    triggers.push(t); return t;
  };
  const config = { reminderTo: [] };
  const worker = kind => () => { assert.equal(locked, true); calls.push(kind); if (faults.logOnWorker === kind) faults.errorRows++; if (faults.worker === kind) throw Error('private student content'); return faults.result?.[kind]; };
  const c = vm.createContext({
    console, Logger: { log() {} }, Utilities: { getUuid: () => 'uuid-' + (++serial) },
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: () => ({ getLastRow: () => { if (faults.errorInventory) throw Error('read error'); return faults.errorRows + 1; } }) }) },
    PropertiesService: { getScriptProperties: () => store(props), getUserProperties: () => {
      if (!users.has(user)) users.set(user, new Map()); return store(users.get(user));
    } },
    ScriptApp: { EventType: { CLOCK: 'CLOCK' }, getProjectTriggers: () => {
      if (faults.inventory) throw Error('inventory unavailable'); return triggers.filter(t => t.owner === user);
    }, newTrigger: handler => {
      const timing = {}; const builder = {
        timeBased() { return this; }, atHour(hour) { timing.hour = hour; return this; },
        everyDays(days) { timing.days = days; return this; }, inTimezone(timezone) { timing.timezone = timezone; return this; },
        create() { assert.equal(locked, true); if (++createCount === faults.createAt) throw Error('create failed'); const t = add(handler); Object.assign(t, timing); return t; }
      }; return builder;
    }, deleteTrigger: t => { assert.equal(locked, true); if (faults.delete?.(t)) throw Error('delete failed'); triggers.splice(triggers.indexOf(t), 1); } },
    withAppLock_: fn => { assert.equal(locked, false); locked = true; try { return fn(); } finally { locked = false; } },
    getConfig_: () => { if (faults.readConfig) throw Error('config recovery required'); return config; },
    assertAdminReadyForAutomation_: () => { assert.equal(locked, true); if (faults.readiness) throw Error('setup incomplete'); },
    validateAppConfig_: () => { calls.push('validate'); if (faults.config) throw Error('incomplete config'); },
    prepareSendDataUnlocked_: worker('prepare'),
    sendMessagesUnlocked_: worker('send'), remindUngradedAndErrorsUnlocked_: worker('reminder')
  });
  c.getAppSpreadsheet_ = () => c.SpreadsheetApp.getActiveSpreadsheet();
  vm.runInContext(code, c);
  return { c, props, triggers, calls, faults, config, add, switchUser: value => { user = value; }, isLocked: () => locked,
    state: () => plain(c.getAutomationState()), start: (s = schedule, replace = false) => c.configureAutomation(s, c.getAutomationState().revision, replace),
    fire: t => c[t.handler]({ triggerUid: t.id }) };
}

test('reading defaults and profile schedule never creates triggers or requires trigger inventory', () => {
  const e = environment(); e.faults.inventory = true;
  assert.deepEqual(plain(e.c.getAutomationSchedule_()), schedule);
  assert.equal(e.c.getAutomationSummary_().enabled, false);
  e.c.saveAutomationScheduleDraft_({ ...schedule, importHour: 4 });
  assert.equal(e.c.getAutomationSchedule_().importHour, 4);
  assert.equal(e.triggers.length, 0);
});

test('explicit start creates only JST daily windows and ordered locked delivery', () => {
  const e = environment(); const unrelated = e.add('otherJob');
  const result = e.start();
  assert.equal(result.enabled, true); assert.equal(result.timezone, 'Asia/Tokyo');
  assert.deepEqual(e.triggers.filter(t => t !== unrelated).map(t => [t.hour, t.days, t.timezone]), [[5, 1, 'Asia/Tokyo'], [8, 1, 'Asia/Tokyo']]);
  e.calls.length = 0; e.fire(e.triggers.find(t => t.hour === 5));
  assert.deepEqual(e.calls, ['validate', 'prepare']);
  const delivery = result.triggers.find(t => t.kind === 'delivery');
  e.calls.length = 0; e.fire(e.triggers.find(t => t.id === delivery.id));
  assert.deepEqual(e.calls, ['validate', 'send']); assert.equal(e.isLocked(), false);
  assert.equal(e.state().lastRuns[0].status, 'success');
});

test('strict hours, ordering, reminder recipients and config are validated before creation', () => {
  const e = environment();
  for (const value of [-1, 24, 5.5, '5', null, true]) assert.throws(() => e.start({ ...schedule, importHour: value }), /時刻/);
  assert.throws(() => e.start({ ...schedule, deliveryHour: 5 }), /順/);
  assert.throws(() => e.start({ ...schedule, reminderEnabled: 'false' }), /有効/);
  assert.throws(() => e.start({ ...schedule, reminderEnabled: true }), /通知先/);
  e.config.reminderTo = ['bad']; assert.throws(() => e.start({ ...schedule, reminderEnabled: true }), /通知先/);
  e.faults.config = true; assert.throws(() => e.start(), /incomplete/);
  assert.equal(e.triggers.length, 0);
});

test('legacy clock triggers require explicit replacement; unrelated and non-clock handlers survive', () => {
  const e = environment(); const legacy = e.add('sendMessages'), other = e.add('otherJob'), edit = e.add('sendMessages', 'EDIT');
  assert.equal(e.state().legacyCount, 1); assert.throws(() => e.start(), /旧.*確認/);
  e.start(schedule, true);
  assert.ok(!e.triggers.includes(legacy)); assert.ok(e.triggers.includes(other)); assert.ok(e.triggers.includes(edit));
});

test('direct preparation clock triggers are detected for managed replacement and stopping',()=>{
  for(const operation of ['replace','stop']) {
    const e=environment(),direct=e.add('prepareSendData'),edit=e.add('prepareSendData','EDIT');
    assert.equal(e.state().legacyCount,1);
    if(operation==='replace') { assert.throws(()=>e.start(),/旧.*確認/);e.start(schedule,true); }
    else e.c.stopAutomation(e.state().revision,true);
    assert.ok(!e.triggers.includes(direct));assert.ok(e.triggers.includes(edit));
  }
});

test('saved trigger generation continues preparation and sending without recreation after code reload',()=>{
  const e=environment();e.start();const before=e.props.get('APP_AUTOMATION_STATE'),active=[...e.triggers];
  // GASの別実行と同様に、保存済みプロパティとトリガーだけを次の実行環境へ渡す。
  const next=environment();for(const [key,value] of e.props)next.props.set(key,value);
  const owner=e.c.PropertiesService.getUserProperties().getProperty('APP_AUTOMATION_OWNER');
  next.c.PropertiesService.getUserProperties().setProperty('APP_AUTOMATION_OWNER',owner);
  next.triggers.push(...active);next.fire(active[0]);next.fire(active[1]);
  assert.deepEqual(next.calls,['validate','prepare','validate','send']);
  assert.equal(next.props.get('APP_AUTOMATION_STATE'),before);assert.deepEqual(next.triggers,active);
});

test('stale revisions and different owners cannot change or execute managed automation', () => {
  const e = environment(); const revision = e.state().revision; e.start(); const oldTriggers = [...e.triggers];
  assert.throws(() => e.c.configureAutomation(schedule, revision, false), /変更/);
  e.switchUser('b'); assert.throws(() => e.start(), /管理者/);
  assert.throws(() => e.c.stopAutomation(e.state().revision, true), /管理者/);
  e.calls.length = 0; e.fire(oldTriggers[0]); assert.deepEqual(e.calls, []);
  assert.equal(e.c.getAutomationSummary_().managedByCurrentUser, false);
});

test('partial creation and failed persistence keep old generation active; orphan ids never run', () => {
  const e = environment(); e.start(); const old = [...e.triggers];
  e.faults.createAt = 4; e.faults.delete = t => !old.includes(t);
  assert.throws(() => e.start({ ...schedule, importHour: 4 }), /作成/);
  const orphan = e.triggers.find(t => !old.includes(t)); e.calls.length = 0; e.fire(orphan); assert.deepEqual(e.calls, []);
  e.faults.createAt = 0; e.faults.property = (key, value, after) => key === 'APP_AUTOMATION_STATE' && !after;
  assert.throws(() => e.start({ ...schedule, importHour: 3 }), /保存/);
  assert.equal(e.state().schedule.importHour, 5);
  e.calls.length = 0; e.fire(old[0]); assert.deepEqual(e.calls, ['validate', 'prepare']);
});

test('write-then-throw commits are recognized and failed old cleanup cannot duplicate execution', () => {
  const e = environment(); e.start(); const old = [...e.triggers];
  e.faults.property = (key, value, after) => key === 'APP_AUTOMATION_STATE' && after;
  e.faults.delete = t => old.includes(t);
  const next = e.start({ ...schedule, importHour: 4 });
  assert.equal(next.schedule.importHour, 4); assert.ok(next.warnings.some(w => /削除|無効/.test(w)));
  e.calls.length = 0; old.forEach(t => e.fire(t)); assert.deepEqual(e.calls, []);
  e.fire(e.triggers.find(t => next.triggers.some(row => row.id === t.id && row.kind === 'delivery' && row.active)));
  assert.deepEqual(e.calls, ['validate', 'send']);
});

test('stop disables allowlist before deleting; cleanup failure cannot execute; failed commit does not delete', () => {
  const e = environment(); e.start(); const old = [...e.triggers];
  e.faults.property = (key, value, after) => key === 'APP_AUTOMATION_STATE' && !after;
  assert.throws(() => e.c.stopAutomation(e.state().revision, false), /保存/); assert.equal(e.triggers.length, 2);
  e.faults.property = null; e.faults.delete = () => true;
  const stopped = e.c.stopAutomation(e.state().revision, false); assert.equal(stopped.enabled, false);
  e.calls.length = 0; old.forEach(t => e.fire(t)); assert.deepEqual(e.calls, []);
});

test('profile draft changes desired schedule without altering actual trigger generation', () => {
  const e = environment(); e.start(); const old = [...e.triggers];
  e.c.saveAutomationScheduleDraft_({ ...schedule, importHour: 4 });
  assert.deepEqual(e.triggers, old); assert.equal(e.state().enabled, true);
  assert.ok(e.state().warnings.some(w => /時間帯.*異|反映/.test(w)));
});

test('preparation failure records an error without sending; invalid config prevents work', () => {
  const e = environment(); e.start(); const delivery = e.triggers.find(t => t.hour === 8);
  e.calls.length = 0; e.faults.worker = 'prepare'; assert.throws(() => e.fire(e.triggers.find(t => t.hour === 5)), /private student/);
  assert.deepEqual(e.calls, ['validate', 'prepare']); assert.equal(e.state().lastRuns[0].status, 'error');
  assert.ok(!JSON.stringify(e.state().lastRuns).includes('student'));
  e.calls.length = 0; e.faults.config = true; assert.throws(() => e.fire(delivery), /incomplete/);
  assert.deepEqual(e.calls, ['validate']);
});

test('partial legacy deletion never activates new managed triggers and reports manual recovery', () => {
  const e = environment(); const first = e.add('importFromFormsToEval'), second = e.add('sendMessages');
  e.faults.delete = t => t === second;
  assert.throws(() => e.start(schedule, true), /一部.*旧|旧.*一部/);
  assert.ok(!e.triggers.includes(first)); assert.ok(e.triggers.includes(second)); assert.equal(e.state().enabled, false);
  e.calls.length = 0; e.triggers.filter(t => t !== second).forEach(t => e.fire(t)); assert.deepEqual(e.calls, []);
});

test('valid reminder gets its own handler; invalid recipients at execution prevent sending', () => {
  const e = environment(); e.config.reminderTo = ['teacher@example.invalid'];
  assert.equal(e.c.getAutomationSchedule_().reminderEnabled, false);
  e.start({ ...schedule, reminderEnabled: true }); const reminder = e.triggers.find(t => t.hour === 16);
  e.calls.length = 0; e.fire(reminder); assert.deepEqual(e.calls, ['validate', 'reminder']);
  e.calls.length = 0; e.config.reminderTo = []; assert.throws(() => e.fire(reminder), /通知先/);
  assert.deepEqual(e.calls, ['validate']);
});

test('missing, mismatched and fabricated trigger UIDs never execute a worker', () => {
  const e = environment(); e.start(); e.calls.length = 0;
  e.c.managedAutomationDelivery_(); e.c.managedAutomationDelivery_({ triggerUid: 'fake' });
  e.c.managedAutomationDelivery_({ triggerUid: e.triggers.find(t => t.hour === 5).id });
  assert.deepEqual(e.calls, []);
});

test('stop preserves legacy without explicit include and reports failed legacy removal', () => {
  const e = environment(); const legacy = e.add('sendMessages');
  e.c.stopAutomation(e.state().revision, false); assert.ok(e.triggers.includes(legacy));
  e.faults.delete = t => t === legacy;
  const state = e.c.stopAutomation(e.state().revision, true);
  assert.equal(state.legacyCount, 1); assert.ok(state.warnings.some(w => /旧処理が動く/.test(w)));
});

test('failed commit after legacy removal explains unrecoverable old schedule', () => {
  const e = environment(); e.add('sendMessages');
  e.faults.property = (key, value, after) => key === 'APP_AUTOMATION_STATE' && !after;
  assert.throws(() => e.start(schedule, true), /旧トリガーは削除済み.*復元/);
  assert.equal(e.state().enabled, false); assert.equal(e.triggers.length, 0);
});

test('run history is bounded and exposes only safe counts from worker results', () => {
  const e = environment(); e.start(); const delivery = e.triggers.find(t => t.hour === 8);
  e.faults.result = { send: '送信処理が終了しました。\n送信成功：2件\nエラー　：1件（エラーシート参照）\n要確認　：3件（送信確認待ち・エラーシート参照）' };
  e.fire(delivery); assert.match(e.state().lastRuns[0].message, /成功.*2.*エラー.*1.*要確認.*3/);
  e.faults.result.send = 'student@example.invalid private';
  for (let i = 0; i < 25; i++) e.fire(delivery);
  assert.equal(e.state().lastRuns.length, 20); assert.ok(!JSON.stringify(e.state().lastRuns).includes('student'));
});

test('completed workers with rejected or uncertain deliveries are attention, with elapsed time',()=>{
  const e=environment();e.start();e.faults.result={send:'送信成功：2件\nエラー　：1件\n要確認　：3件'};
  e.fire(e.triggers.find(t=>t.hour===8));const run=e.state().lastRuns[0];
  assert.equal(run.status,'attention');assert.ok(run.elapsedMs>=0);assert.ok(Number.isFinite(run.elapsedMs));
});

test('new logged row failures mark a nonthrowing import as attention but historical logs do not',()=>{
  const e=environment();e.start();e.faults.errorRows=4;const trigger=e.triggers.find(t=>t.hour===5);
  e.fire(trigger);assert.equal(e.state().lastRuns[0].status,'success');
  e.faults.logOnWorker='prepare';e.fire(trigger);assert.equal(e.state().lastRuns[0].status,'attention');
  assert.match(e.state().lastRuns[0].message,/1件/);
});

test('unavailable result inspection is not reported as fully successful',()=>{
  const e=environment();e.start();e.faults.errorInventory=true;
  e.fire(e.triggers.find(t=>t.hour===5));assert.equal(e.state().lastRuns[0].status,'attention');
});

test('unknown commit outcome preserves triggers for readback and still gates exactly one generation', () => {
  const e = environment(); e.start(); const old = [...e.triggers];
  e.faults.property = (key, value, after) => {
    if (key !== 'APP_AUTOMATION_STATE' || !after) return false;
    e.faults.read = key => key === 'APP_AUTOMATION_STATE'; return true;
  };
  assert.throws(() => e.start({ ...schedule, importHour: 4 }), /保存結果を確認/);
  assert.equal(e.triggers.length, 4);
  e.faults.read = null; e.faults.property = null;
  assert.equal(e.state().schedule.importHour, 4);
  e.calls.length = 0; old.forEach(t => e.fire(t)); assert.deepEqual(e.calls, []);
  e.fire(e.triggers.find(t => t.hour === 4)); assert.deepEqual(e.calls, ['validate', 'prepare']);
});

test('run history write failure does not rerun or hide the pipeline outcome', () => {
  const e = environment(); e.start(); const delivery = e.triggers.find(t => t.hour === 8);
  e.faults.property = key => key === 'APP_AUTOMATION_RUNS'; e.calls.length = 0;
  e.fire(delivery); assert.deepEqual(e.calls, ['validate', 'send']);
  e.faults.worker = 'send'; assert.throws(() => e.fire(delivery), /private student content/);
  assert.equal(e.isLocked(), false);
});

test('long Japanese attention history stays within property byte budget and keeps newest result', () => {
  const e = environment(); e.start();
  const delivery = e.triggers.find(t => t.hour === 8);
  e.faults.result = { send: '送信成功：500件、エラー：500件、要確認：500件' };
  e.faults.logOnWorker = 'send';
  for (let i = 0; i < 30; i++) e.fire(delivery);
  const history = e.props.get('APP_AUTOMATION_RUNS');
  assert.ok(Buffer.byteLength(history, 'utf8') <= 8000);
  assert.ok(JSON.parse(history).length > 0);
  assert.equal(JSON.parse(history)[0].status, 'attention');
});

test('setup readiness is checked under lock before any owner or trigger mutation', () => {
  const e = environment(); e.faults.readiness = true;
  // A mutation would produce a distinct error before readiness validation.
  e.faults.property = () => { throw Error('unexpected property mutation'); };
  assert.throws(() => e.start(), /setup incomplete/);
  assert.equal(e.triggers.length, 0); assert.equal(e.props.size, 0);
});

test('broken configuration cannot block state reads or stopping legacy triggers', () => {
  const e = environment(); const legacy = e.add('sendMessages'); e.faults.readConfig = true;
  assert.deepEqual(plain(e.c.getAutomationSchedule_()), schedule);
  const state = e.state(); assert.equal(state.legacyCount, 1);
  assert.equal(state.hasManagedOwner, false); assert.equal(state.managedByCurrentUser, false);
  const stopped = e.c.stopAutomation(state.revision, true);
  assert.equal(stopped.enabled, false); assert.ok(!e.triggers.includes(legacy));
});

test('state exposes owner flags without exposing the private ownership token', () => {
  const e = environment(); e.start();
  assert.equal(e.state().hasManagedOwner, true); assert.equal(e.state().managedByCurrentUser, true);
  assert.equal(e.state().owner, undefined);
  e.switchUser('b'); assert.equal(e.state().managedByCurrentUser, false);
});
