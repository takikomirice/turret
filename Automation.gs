/** Daily automation. Trigger IDs are committed together so obsolete/orphan triggers cannot run. */
const AUTOMATION_STATE_KEY_ = 'APP_AUTOMATION_STATE';
const AUTOMATION_OWNER_KEY_ = 'APP_AUTOMATION_OWNER';
const AUTOMATION_RUNS_KEY_ = 'APP_AUTOMATION_RUNS';
const AUTOMATION_TIMEZONE_ = 'Asia/Tokyo';
const AUTOMATION_HANDLERS_ = {
  import: 'managedAutomationImport_',
  delivery: 'managedAutomationDelivery_',
  reminder: 'managedAutomationReminder_'
};
const AUTOMATION_LEGACY_HANDLERS_ = ['importFromFormsToEval', 'evalToSendSheet', 'sendMessages', 'remindUngradedAndErrors'];

/** Intake deadlines are independent of return automation. One shared clock, no per-form triggers. */
const FORM_SCHEDULE_KEY_ = 'TURRET_FORM_SCHEDULE_TRIGGER';
const FORM_SCHEDULE_HANDLER_ = 'managedFormScheduleTick_';

function readManagedFormSchedule_() {
  const raw=PropertiesService.getScriptProperties().getProperty(FORM_SCHEDULE_KEY_);
  if(!raw)return null;
  let record;try{record=JSON.parse(raw);}catch(error){throw new Error('受付終了タイマーの登録情報を読めません。');}
  if(!record||!record.id||!record.owner||!record.spreadsheetId||!record.scriptId)throw new Error('受付終了タイマーの登録情報が不正です。');
  return record;
}

function assertManagedFormScheduleOwner_() {
  const record=readManagedFormSchedule_();
  if(record&&(record.spreadsheetId!==getAppSpreadsheet_().getId()||record.scriptId!==ScriptApp.getScriptId()))throw new Error('別の運用の受付終了タイマーです。');
  return assertAutomationOwner_(record||{owner:''});
}

/** Caller holds app lock. Only a committed ID is allowed to run. */
function ensureManagedFormScheduleTrigger_() {
  const owner=assertManagedFormScheduleOwner_(),old=readManagedFormSchedule_();
  const triggers=ScriptApp.getProjectTriggers().filter(function(t){return t.getHandlerFunction()===FORM_SCHEDULE_HANDLER_&&t.getEventType()===ScriptApp.EventType.CLOCK;});
  if(old&&triggers.some(function(t){return String(t.getUniqueId())===old.id;}))return old;
  if(triggers.length)throw new Error('受付終了タイマーの作成結果が未確認です。Apps Scriptの未登録トリガーを確認してから修復してください。');
  const trigger=ScriptApp.newTrigger(FORM_SCHEDULE_HANDLER_).timeBased().everyMinutes(5).create();
  const record={id:String(trigger.getUniqueId()),owner:owner,spreadsheetId:getAppSpreadsheet_().getId(),scriptId:ScriptApp.getScriptId()};
  const props=PropertiesService.getScriptProperties(),json=JSON.stringify(record);
  try{props.setProperty(FORM_SCHEDULE_KEY_,json);}catch(error){
    // The registration may have committed despite a failed response. Never guess its outcome.
    let saved;try{saved=props.getProperty(FORM_SCHEDULE_KEY_);}catch(readError){throw new Error('受付終了タイマーの保存結果が未確認です。状態を更新してください。');}
    if(saved!==json){try{ScriptApp.deleteTrigger(trigger);}catch(ignored){}throw error;}
  }
  return record;
}

function repairFormScheduleTimer() { return withAppLock_(function(){ensureManagedFormScheduleTrigger_();return getManagedFormScheduleSummary_();}); }

function getManagedFormScheduleSummary_() {
  try{
    const record=readManagedFormSchedule_();
    if(!record)return {registered:false,healthy:false,message:'受付終了日時の設定時に専用タイマーを作成します。'};
    const owner=PropertiesService.getUserProperties().getProperty(AUTOMATION_OWNER_KEY_);
    if(record.owner!==owner||record.spreadsheetId!==getAppSpreadsheet_().getId()||record.scriptId!==ScriptApp.getScriptId())return {registered:true,healthy:false,message:'受付終了タイマーを設定した管理者・運用で確認してください。'};
    const healthy=ScriptApp.getProjectTriggers().some(function(t){return String(t.getUniqueId())===record.id&&t.getHandlerFunction()===FORM_SCHEDULE_HANDLER_;});
    return {registered:true,healthy:healthy,message:healthy?'受付終了を約5分間隔で確認します。実行時刻は遅れる場合があります。':'受付終了タイマーが見つかりません。「タイマーを修復」を実行してください。'};
  }catch(error){return {registered:true,healthy:false,message:String(error.message||error)};}
}

function managedFormScheduleTick_(event) {
  return withAppLock_(function(){
    const registration=readManagedFormSchedule_();
    if(!registration||!event||String(event.triggerUid)!==registration.id||registration.owner!==PropertiesService.getUserProperties().getProperty(AUTOMATION_OWNER_KEY_)||
      registration.spreadsheetId!==getAppSpreadsheet_().getId()||registration.scriptId!==ScriptApp.getScriptId())return;
    const started=Date.now(),records=getManagedRecords_(),cache={records:records,sheet:null};
    const due=records.filter(function(r){return r.kind==='form'&&r.ownerSpreadsheetId===registration.spreadsheetId&&r.ownerScriptId===registration.scriptId&&managedCloseIsDue_(r)&&r.closeState!=='closed';});
    due.sort(function(a,b){return String(a.closeCheckedAt||'').localeCompare(String(b.closeCheckedAt||''));});
    due.forEach(function(r){
      if(Date.now()-started>200000)return;
      try{
        const form=FormApp.openById(r.formId);
        if(form.isAcceptingResponses())form.setAcceptingResponses(false);
        r.accepting=form.isAcceptingResponses();
        if(r.accepting)throw new Error('フォームが受付中のままです。');
        r.closeState='closed';r.closedAt=new Date().toISOString();r.closeError='';
        if(r.formUpdate)r.formUpdate.accepting=false;
      }catch(error){r.closeState='error';r.closeError=String(error.message||error).slice(0,500);}
      r.closeCheckedAt=new Date().toISOString();
      try{saveManagedRecord_(r,cache);}catch(error){Logger.log('受付終了の結果記録に失敗しました。次回に再確認します。');}
    });
  });
}

/** Separate from daily delivery: never sends messages. Caller holds application lock. */
function ensureManagedResponseTrigger_() {
  const props=PropertiesService.getScriptProperties(),raw=props.getProperty('TURRET_RESPONSE_TRIGGER');
  const record=raw?JSON.parse(raw):null,ss=getAppSpreadsheet_();
  const owner=assertAutomationOwner_(record||{owner:''});
  if(record && record.spreadsheetId!==ss.getId())throw new Error('別の運用の回答トリガーが登録されています。');
  const triggers=ScriptApp.getProjectTriggers().filter(function(t){return t.getHandlerFunction()==='managedFormResponseReceived_';});
  if(record && triggers.some(function(t){return String(t.getUniqueId())===record.id;}))return record;
  if(triggers.length)throw new Error('回答トリガーの作成結果を確認してください。Apps Scriptで残った未登録トリガーを確認・削除してから再実行してください。');
  const trigger=ScriptApp.newTrigger('managedFormResponseReceived_').forSpreadsheet(ss).onFormSubmit().create();
  const next={id:String(trigger.getUniqueId()),owner:owner,spreadsheetId:ss.getId()};
  props.setProperty('TURRET_RESPONSE_TRIGGER',JSON.stringify(next));return next;
}

function managedFormResponseReceived_(event) {
  return withAppLock_(function(){
    const raw=PropertiesService.getScriptProperties().getProperty('TURRET_RESPONSE_TRIGGER');
    if(!raw||!event||!event.triggerUid||!event.range||!event.source)return;
    const registration=JSON.parse(raw), ss=getAppSpreadsheet_();
    if(registration.id!==String(event.triggerUid)||registration.owner!==PropertiesService.getUserProperties().getProperty(AUTOMATION_OWNER_KEY_)||registration.spreadsheetId!==ss.getId()||event.source.getId()!==ss.getId())return;
    const sheet=event.range.getSheet(),start=event.range.getRow();
    if(start<2 || event.range.getNumRows()!==1 || sheet.getParent().getId()!==ss.getId())return;
    const record=getManagedRecords_().find(function(r){return r.kind==='form'&&String(r.responseSheetId)===String(sheet.getSheetId());});
    if(!record)return;
    assertManagedOwner_(record);
    if(record.formUpdate)return;
    const result=fillManagedNames_(record,sheet,start,1);
    if(result.unmatched)ensureErrorSheet_().appendRow([new Date(),sheet.getName(),start,'','','NAME_UNMATCHED','Classroom生徒一覧に照合できません。名簿を更新してフォーム管理の「名前を補完」を実行してください。']);
  });
}

function automationRecipientsValid_(config) {
  return Array.isArray(config.reminderTo) && config.reminderTo.length > 0 && config.reminderTo.every(function(address) {
    return typeof address === 'string' && /^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(address);
  });
}

function validateAutomationSchedule_(schedule) {
  if (!schedule || typeof schedule !== 'object' || Array.isArray(schedule)) throw new Error('自動実行の時刻設定が不正です。');
  ['importHour', 'deliveryHour', 'reminderHour'].forEach(function(key) {
    if (typeof schedule[key] !== 'number' || !Number.isInteger(schedule[key]) || schedule[key] < 0 || schedule[key] > 23) {
      throw new Error('時刻は0〜23の整数で入力してください。');
    }
  });
  if (typeof schedule.reminderEnabled !== 'boolean') throw new Error('リマインダーの有効・無効を選択してください。');
  if (schedule.importHour >= schedule.deliveryHour || (schedule.reminderEnabled && schedule.deliveryHour >= schedule.reminderHour)) {
    throw new Error('時間帯は取り込み、準備と送信、リマインダーの順に重ならないよう設定してください。');
  }
  return { importHour: schedule.importHour, deliveryHour: schedule.deliveryHour, reminderHour: schedule.reminderHour, reminderEnabled: schedule.reminderEnabled };
}

function readAutomationRecord_() {
  const raw = PropertiesService.getScriptProperties().getProperty(AUTOMATION_STATE_KEY_);
  if (!raw) return {
    revision: '0', owner: '', enabled: false, active: [], activeSchedule: null,
    // Recovery/stop must remain available even when application configuration cannot be read.
    schedule: { importHour: 5, deliveryHour: 8, reminderHour: 16, reminderEnabled: false }
  };
  let record;
  try { record = JSON.parse(raw); } catch (error) { throw new Error('自動実行の保存情報を読めません。管理者に確認してください。'); }
  if (!record || typeof record.revision !== 'string' || typeof record.owner !== 'string' || typeof record.enabled !== 'boolean' || !Array.isArray(record.active)) {
    throw new Error('自動実行の保存情報が不正です。管理者に確認してください。');
  }
  record.schedule = validateAutomationSchedule_(record.schedule);
  return record;
}

function getAutomationSchedule_() {
  return readAutomationRecord_().schedule;
}

function getAutomationSummary_() {
  const record = readAutomationRecord_();
  const token = PropertiesService.getUserProperties().getProperty(AUTOMATION_OWNER_KEY_);
  return { schedule: record.schedule, enabled: record.enabled, hasManagedOwner: !!record.owner, managedByCurrentUser: !!record.owner && record.owner === token };
}

/** Caller holds the common application lock. Never accesses ScriptApp. */
function saveAutomationScheduleDraft_(schedule) {
  const desired = validateAutomationSchedule_(schedule);
  const record = readAutomationRecord_();
  record.schedule = desired;
  writeAutomationRecord_(record);
  return desired;
}

function writeAutomationRecord_(record) {
  record.revision = Utilities.getUuid();
  const serialized = JSON.stringify(record);
  const properties = PropertiesService.getScriptProperties();
  try { properties.setProperty(AUTOMATION_STATE_KEY_, serialized); }
  catch (error) {
    // A service may report failure after persisting. Read back before deciding to roll back.
    let current;
    try { current = properties.getProperty(AUTOMATION_STATE_KEY_); }
    catch (readError) {
      const uncertain = new Error('自動実行の保存結果を確認できません。画面を再読込して状態を確認してください。');
      uncertain.commitUncertain = true;
      throw uncertain;
    }
    if (current !== serialized) throw new Error('自動実行の設定を保存できませんでした。画面を再読込して再実行してください。');
  }
}

function assertAutomationRevision_(record, expectedRevision) {
  if (expectedRevision !== record.revision) throw new Error('自動実行の設定が別の操作で変更されました。画面を再読込してください。');
}

function assertAutomationOwner_(record) {
  let token = PropertiesService.getUserProperties().getProperty(AUTOMATION_OWNER_KEY_);
  if (record.owner && token !== record.owner) throw new Error('別の管理者が自動実行を管理しています。同じ管理者アカウントで操作してください。');
  if (!token) {
    token = Utilities.getUuid();
    PropertiesService.getUserProperties().setProperty(AUTOMATION_OWNER_KEY_, token);
  }
  return token;
}

function listAutomationTriggers_() {
  return ScriptApp.getProjectTriggers().filter(function(trigger) {
    return trigger.getEventType() === ScriptApp.EventType.CLOCK;
  }).map(function(trigger) {
    const handler = trigger.getHandlerFunction();
    const kind = Object.keys(AUTOMATION_HANDLERS_).find(function(key) { return AUTOMATION_HANDLERS_[key] === handler; });
    return { trigger: trigger, id: String(trigger.getUniqueId()), handler: handler, kind: kind || 'legacy', managed: !!kind, legacy: AUTOMATION_LEGACY_HANDLERS_.indexOf(handler) >= 0 };
  }).filter(function(row) { return row.managed || row.legacy; });
}

function getAutomationState() {
  return withAppLock_(getAutomationStateUnlocked_);
}

function getAutomationStateUnlocked_() {
  const record = readAutomationRecord_();
  const token = PropertiesService.getUserProperties().getProperty(AUTOMATION_OWNER_KEY_);
  const owned = record.owner && record.owner === token;
  const rows = listAutomationTriggers_();
  const warnings = [
    '実行時刻はAsia/Tokyoの指定した時台です。分単位の定刻は保証されません。',
    '他のアカウントが作成したトリガーは確認できません。管理者を一人に決め、他のアカウントの旧トリガーは各作成者が停止してください。'
  ];
  if (record.owner && !owned) warnings.push('別の管理者の自動実行設定です。開始・変更・停止はその管理者アカウントで行ってください。');
  if (record.enabled && JSON.stringify(record.schedule) !== JSON.stringify(record.activeSchedule)) warnings.push('保存した時間帯と稼働中の時間帯が異なります。「開始・変更」で反映してください。');
  const triggers = rows.map(function(row) {
    const active = !!owned && record.enabled && record.active.some(function(item) { return item.id === row.id && item.kind === row.kind; });
    return { id: row.id, handler: row.handler, kind: row.kind, managed: row.managed, legacy: row.legacy, active: active };
  });
  const legacyCount = rows.filter(function(row) { return row.legacy; }).length;
  if (legacyCount) warnings.push('旧処理のトリガーが' + legacyCount + '件あります。切り替える場合は旧トリガーの置換を確認してください。');
  if (triggers.some(function(row) { return row.managed && !row.active; })) warnings.push('無効な管理用トリガーが残っています。実行は抑止されています。開始・変更または停止で削除を再試行できます。');
  if (owned && record.enabled && record.active.some(function(item) { return !triggers.some(function(row) { return row.active && row.id === item.id; }); })) {
    warnings.push('稼働予定のトリガーが見つかりません。「開始・変更」で作り直してください。');
  }
  let lastRuns = [];
  try { lastRuns = JSON.parse(PropertiesService.getScriptProperties().getProperty(AUTOMATION_RUNS_KEY_) || '[]'); } catch (error) { warnings.push('最近の実行結果を読み取れませんでした。'); }
  return { schedule: record.schedule, timezone: AUTOMATION_TIMEZONE_, enabled: record.enabled, legacyCount: legacyCount,
    hasManagedOwner: !!record.owner, managedByCurrentUser: !!owned,
    triggers: triggers, warnings: warnings, revision: record.revision, lastRuns: Array.isArray(lastRuns) ? lastRuns : [] };
}

function cleanupAutomationTriggers_(rows) {
  const failed = [];
  rows.forEach(function(row) {
    try { ScriptApp.deleteTrigger(row.trigger || row); } catch (error) { failed.push(row); }
  });
  return failed;
}

function configureAutomation(schedule, expectedRevision, replaceLegacy, expectedConfigRevision) {
  return withAppLock_(function() {
    const desired = validateAutomationSchedule_(schedule);
    const record = readAutomationRecord_();
    assertAutomationRevision_(record, expectedRevision);
    if (expectedConfigRevision !== undefined && expectedConfigRevision !== getConfigRevision_(getConfig_())) {
      throw new Error('画面を開いた後に送信設定が変更されました。再読込して内容を確認してください。');
    }
    assertAdminReadyForAutomation_();
    const owner = assertAutomationOwner_(record);
    const config = getConfig_();
    validateAppConfig_(config);
    if (desired.reminderEnabled && !automationRecipientsValid_(config)) throw new Error('リマインダーの通知先を正しいメールアドレスで設定するか、リマインダーを無効にしてください。');
    const existing = listAutomationTriggers_();
    const legacy = existing.filter(function(row) { return row.legacy; });
    if (legacy.length && replaceLegacy !== true) throw new Error('旧トリガーの置換確認が必要です。対象を確認してから切り替えてください。');
    const created = [];
    try {
      ['import', 'delivery'].concat(desired.reminderEnabled ? ['reminder'] : []).forEach(function(kind) {
        const trigger = ScriptApp.newTrigger(AUTOMATION_HANDLERS_[kind]).timeBased().atHour(desired[kind + 'Hour']).everyDays(1).inTimezone(AUTOMATION_TIMEZONE_).create();
        created.push({ trigger: trigger, id: String(trigger.getUniqueId()), kind: kind });
      });
    } catch (error) {
      cleanupAutomationTriggers_(created);
      throw new Error('自動実行トリガーの作成に失敗しました。新しいトリガーは有効化していません。画面を再読込してください。');
    }
    // Legacy triggers are not gated, so they must all be removed before enabling the new generation.
    if (cleanupAutomationTriggers_(legacy).length) {
      cleanupAutomationTriggers_(created);
      throw new Error('旧トリガーを一部削除できませんでした。新しい自動実行は有効化していません。削除済みの旧時間帯は復元できないため、残った旧トリガーを確認して切り替えを再実行してください。');
    }
    record.schedule = desired;
    record.activeSchedule = desired;
    record.owner = owner;
    record.enabled = true;
    record.active = created.map(function(row) { return { id: row.id, kind: row.kind }; });
    try { writeAutomationRecord_(record); }
    catch (error) {
      if (!error.commitUncertain) cleanupAutomationTriggers_(created);
      if (legacy.length) throw new Error(error.message + ' 旧トリガーは削除済みで時間帯を復元できません。状態を確認して開始を再実行してください。');
      throw error;
    }
    // Failure here is safe: only the committed generation can pass the event allowlist.
    cleanupAutomationTriggers_(existing.filter(function(row) { return row.managed; }));
    return getAutomationStateUnlocked_();
  });
}

function stopAutomation(expectedRevision, includeLegacy) {
  return withAppLock_(function() {
    const record = readAutomationRecord_();
    assertAutomationRevision_(record, expectedRevision);
    assertAutomationOwner_(record);
    const existing = listAutomationTriggers_();
    record.enabled = false;
    record.active = [];
    record.activeSchedule = null;
    writeAutomationRecord_(record);
    const failed = cleanupAutomationTriggers_(existing.filter(function(row) { return row.managed || (includeLegacy === true && row.legacy); }));
    const state = getAutomationStateUnlocked_();
    if (failed.some(function(row) { return row.legacy; })) state.warnings.push('旧トリガーを一部停止できませんでした。旧処理が動く可能性があります。残った旧トリガーを停止してください。');
    return state;
  });
}

function automationErrorRowCount_() {
  try {
    const sheet = getAppSpreadsheet_().getSheetByName('エラー');
    return sheet ? Math.max(0, sheet.getLastRow() - 1) : 0;
  } catch (error) { return null; }
}

function recordAutomationRun_(kind, status, startedAt, deliveryResult, errorDelta) {
  try {
    const properties = PropertiesService.getScriptProperties();
    let runs;
    try { runs = JSON.parse(properties.getProperty(AUTOMATION_RUNS_KEY_) || '[]'); } catch (error) { runs = []; }
    if (!Array.isArray(runs)) runs = [];
    let message = status === 'success' ? '処理が完了しました。処理対象の状態とエラーシートも確認してください。' : '処理に失敗しました。Apps Scriptの実行履歴とエラーシートを確認してください。';
    // Only known numeric fields may enter the console. Never persist arbitrary worker/error text.
    if (status === 'success' && kind === 'delivery' && typeof deliveryResult === 'string') {
      const sent = deliveryResult.match(/送信成功\s*：\s*(\d+)件/);
      const errors = deliveryResult.match(/エラー\s*：\s*(\d+)件/);
      const reviews = deliveryResult.match(/要確認\s*：\s*(\d+)件/);
      if (sent && errors && reviews) {
        message = '送信成功：' + sent[1] + '件、エラー：' + errors[1] + '件、要確認：' + reviews[1] + '件。送信シートとエラーシートで確認してください。';
        if (Number(errors[1]) || Number(reviews[1])) status = 'attention';
      }
    }
    if (status !== 'error' && (errorDelta === null || errorDelta > 0)) {
      status = 'attention';
      message += errorDelta === null ? ' エラー記録を確認できませんでした。保守画面で状態を確認してください。' : ' 新しいエラー記録が' + errorDelta + '件あります。保守画面で確認してください。';
    }
    const finishedAt = new Date().toISOString();
    const elapsedMs = Math.max(0, Date.parse(finishedAt) - Date.parse(startedAt));
    runs.unshift({ kind: kind, status: status, startedAt: startedAt, finishedAt: finishedAt, elapsedMs: isFinite(elapsedMs) ? elapsedMs : 0, message: message });
    properties.setProperty(AUTOMATION_RUNS_KEY_, JSON.stringify(runs.slice(0, 20)));
  } catch (error) { Logger.log('自動実行の結果概要を保存できませんでした。'); }
}

function runManagedAutomation_(kind, event) {
  return withAppLock_(function() {
    const record = readAutomationRecord_();
    const owner = PropertiesService.getUserProperties().getProperty(AUTOMATION_OWNER_KEY_);
    if (!event || !event.triggerUid || !record.enabled || !record.owner || record.owner !== owner || !record.active.some(function(row) { return row.id === String(event.triggerUid) && row.kind === kind; })) return;
    const startedAt = new Date().toISOString();
    const errorsBefore = automationErrorRowCount_();
    try {
      let deliveryResult;
      const config = getConfig_();
      validateAppConfig_(config);
      if (kind === 'import') importFromFormsToEvalUnlocked_();
      if (kind === 'delivery') { evalToSendSheetUnlocked_(); deliveryResult = sendMessagesUnlocked_(); }
      if (kind === 'reminder') {
        if (!automationRecipientsValid_(config)) throw new Error('リマインダーの通知先が不正です。設定を確認してください。');
        remindUngradedAndErrorsUnlocked_();
      }
      const errorsAfter = automationErrorRowCount_();
      recordAutomationRun_(kind, 'success', startedAt, deliveryResult,
        errorsBefore === null || errorsAfter === null ? null : Math.max(0, errorsAfter - errorsBefore));
    } catch (error) {
      recordAutomationRun_(kind, 'error', startedAt);
      throw error;
    }
  });
}

function managedAutomationImport_(event) { return runManagedAutomation_('import', event); }
function managedAutomationDelivery_(event) { return runManagedAutomation_('delivery', event); }
function managedAutomationReminder_(event) { return runManagedAutomation_('reminder', event); }
