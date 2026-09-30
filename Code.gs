/******************************************************
 * Google Classroom 評価自動送信 GAS
 *
 * - 管理画面で設定し、ScriptPropertiesへ保存する。
 * - 旧設定シートは初回管理画面表示時に内部保存へ移行して削除する。
 * - 別期間・別プロジェクトへの設定引き継ぎはJSONを使う。
 ******************************************************/

const SETTINGS_SHEET_NAME = '設定(編集不可)';
const INPUT_SETTINGS_SHEET_NAME = '設定シート';
const CONFIG_SHEET_NAME = 'Config';
const EVAL_SHEET_NAME = '評価データ';
const SEND_SHEET_NAME = '送信シート';
const STUDENT_SHEET_NAME = '生徒一覧';
const ERROR_SHEET_NAME = 'エラー';
const MAPPING_SHEET_NAME = '対応表';

const EVAL_BASE_HEADERS = [
  '元SS_ID',
  '元シート名',
  '元行番号',
  'メールアドレス',
  '名前',
  '処理状態'
];

const SEND_BASE_HEADERS = [
  '元SS_ID',
  '元シート名',
  '元行番号',
  'No',
  'メールアドレス',
  '名前',
  'クラス名',
  'コースID',
  'studentId',
  '送信状態'
];

const STUDENT_SHEET_HEADERS = ['No', 'メールアドレス', '名前', 'クラス名', 'コースID', 'studentId'];
const MAPPING_SHEET_HEADERS = ['元SS_ID', '元スプシ名', '元シート名', 'クラス名', 'courseId', 'メモ'];
const SEND_REVIEW_STATUS = '送信確認待ち';

/** A bound copy always uses itself. Web/standalone triggers use one explicitly bound target. */
function getAppSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const id = PropertiesService.getScriptProperties().getProperty('TURRET_WEB_TARGET');
  if (!id) throw new Error('運用スプレッドシートが未接続です。Web画面で新規作成するか、スプシのメニューから接続してください。');
  return SpreadsheetApp.openById(id);
}

/** 更新入口で共通ロックを使用する。対話ダイアログは呼び出し元で解放後に表示する。 */
function withAppLock_(operation) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) {
    throw new Error('別の処理を実行中です。終了してからもう一度実行してください。');
  }
  try {
    return operation();
  } finally {
    try {
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
  }
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('自動送信システム')
    .addItem('管理画面を開く', 'connectWebConsole')
    .addToUi();
}

function initializeSheets() {
  withAppLock_(initializeSheetsUnlocked_);
  safeAlert_('シートを作成しました（クラス一覧・生徒一覧）');
}

function initializeSheetsUnlocked_() {
  ensureSheet_('クラス一覧', ['クラス名', 'コースID', '同期対象(1)']);
  ensureStudentSheetForSync_();
  migrateConfigStorageUnlocked_();
}

function openSettingsDialog() {
  return connectWebConsole();
}

function getSettingsDialogData() {
  // 設定読み込みと保存を排他する。
  return withAppLock_(getSettingsDialogDataUnlocked_);
}

function getSettingsDialogDataUnlocked_() {
  const config = getConfig_();
  const parseError = PropertiesService.getScriptProperties()
    .getProperty('APP_CONFIG_PARSE_ERROR') === '1';

  return {
    config: config,
    revision: getConfigRevision_(config),
    configParseError: parseError,
    availableFieldTypes: [
      { value: 'text', label: 'text' },
      { value: 'date', label: 'date' },
      { value: 'score_grade', label: 'score_grade' }
    ],
    templateGuide: [
      '差し込みは {score} です。例: 評価：{score}',
      '固定トークンも同じ書き方です。例: {student_name}, {class_name}, {course_id}, {row_no}',
      '値だけの行は、値が空なら自動で非表示になります。',
      '括弧や見出しを付けた行も、対応する値が空ならその行は自動で消えます。'
    ].join('\n')
  };
}

function saveSettingsFromDialog(payload, expectedRevision) {
  return withAppLock_(function() {
    return saveSettingsFromDialogUnlocked_(payload, expectedRevision);
  });
}

function saveSettingsFromDialogUnlocked_(payload, expectedRevision) {
  const existing = getConfig_();
  // Older clients supplied a reset confirmation boolean. It no longer permits deletion.
  return applyConfigDraft_(Object.assign({}, existing, payload),
    typeof expectedRevision === 'string' ? expectedRevision : getConfigRevision_(existing));
}

function getConfigRevision_(config) {
  // Deterministic, lossless fingerprint of the effective configuration, including legacy migration overrides.
  return JSON.stringify(normalizeAppConfig_(config));
}

function saveSetupSection(section, payload, expectedRevision) {
  const sections = {};
  Object.defineProperty(sections, section, { value: payload, enumerable: true });
  return saveSetupSections(sections, expectedRevision);
}

/** Save the edited setup sections together; automation is configured separately. */
function saveSetupSections(payload, expectedRevision) {
  return withAppLock_(function() {
    const sections = {
      sources: ['formSources', 'reminderTo', 'formSheetNamePrefix'],
      fields: ['emailHeader', 'studentNameHeader', 'formStatusHeader', 'scoreSourceHeader', 'replyBodyHeader', 'fields', 'gradeScale'],
      template: ['messageTemplate']
    };
    if (!payload || typeof payload !== 'object' || Array.isArray(payload) || !Object.keys(payload).length) {
      throw new Error('保存する設定項目の形式が不正です。');
    }
    const config = getConfig_();
    Object.keys(payload).forEach(function(section) {
      const values = payload[section];
      if (!Object.prototype.hasOwnProperty.call(sections, section) || !values || typeof values !== 'object' || Array.isArray(values)) {
        throw new Error('保存する設定項目の形式が不正です: ' + section);
      }
      Object.keys(values).forEach(function(key) {
        if (sections[section].indexOf(key) < 0) throw new Error('この手順では保存できない設定です: ' + key);
        config[key] = values[key];
      });
    });
    return applyConfigDraft_(config, expectedRevision);
  });
}

/** Caller must hold withAppLock_. Draft saves never mutate evaluation or send rows. */
function applyConfigDraft_(rawConfig, expectedRevision) {
  const existing = getConfig_();
  if (typeof expectedRevision !== 'string' || expectedRevision !== getConfigRevision_(existing)) {
    throw new Error('画面を開いた後に設定が変更されました。再読込してから保存してください。');
  }
  validateConfigDraft_(rawConfig);
  const config = normalizeAppConfig_(rawConfig);
  validateConfigSheetLayouts_(config);
  saveConfig_(config);
  return { ok: true, message: '設定を保存しました。既存データは保持しています。', config: config, revision: getConfigRevision_(config) };
}

/** Read-only preflight shared by setting saves and managed form preparation. */
function validateConfigSheetLayouts_(config) {
  const ss = getAppSpreadsheet_();
  [[EVAL_SHEET_NAME, getConfiguredEvalHeaders_(config)], [SEND_SHEET_NAME, getConfiguredSendHeaders_(config)]].forEach(function(entry) {
    const sheet = ss.getSheetByName(entry[0]);
    if (sheet && sheet.getLastRow() > 1 && (sheet.getLastColumn() !== entry[1].length || !hasMatchingHeaders_(sheet, entry[1]))) {
      throw new Error('「' + entry[0] + '」の既存データと列構成が一致しません。管理画面で必要なデータを退避・クリアしてから保存してください。');
    }
  });
}

/**
 * 指定された URL リストとシート名接頭辞から利用可能なヘッダー候補を収集する。
 * @param {string[]} urlList - 未正規化の URL 文字列の配列
 * @param {string} prefix - 対象シート名の接頭辞
 * @returns {{ headers: string[], warnings: string[] }}
 */
function collectAvailableSourceHeaders_(urlList, prefix) {
  const headerSet = {};
  const headers = [];
  const warnings = [];
  let matchedSheets = 0;
  const normalizedPrefixes = Array.isArray(prefix)
    ? prefix.map(function(p) { return String(p || '').trim(); }).filter(Boolean)
    : (String(prefix || '').trim() ? [String(prefix || '').trim()] : []);

  if (!Array.isArray(urlList) || urlList.length === 0) {
    warnings.push('フォーム回答スプレッドシートが未設定です。');
    return { headers: headers, warnings: warnings };
  }
  if (normalizedPrefixes.length === 0) {
    warnings.push('対象シート名の接頭辞が未設定です。設定シートの FORM_SHEET_PREFIX を入力してください。');
    return { headers: headers, warnings: warnings };
  }

  urlList.forEach(function(rawUrl) {
    let sourceId = '';
    try {
      const normalized = normalizeFormSource_(String(rawUrl || '').trim());
      if (!normalized) {
        warnings.push('URLの形式が不正です: ' + trunc_(rawUrl, 60));
        return;
      }
      sourceId = normalized.id;
    } catch (e) {
      warnings.push('URLの解析に失敗しました: ' + trunc_(rawUrl, 60));
      return;
    }

    try {
      const ss = SpreadsheetApp.openById(sourceId);
      const sheets = ss.getSheets().filter(function(sheet) {
        return !isManagedInternalSheet_(sheet.getName(), sourceId) && sheetMatchesPrefixes_(sheet.getName(), normalizedPrefixes);
      });

      if (sheets.length === 0) {
        warnings.push(
          ss.getName() + ' に対象シートがありません。' +
          (normalizedPrefixes.length > 0
            ? ' 接頭辞「' + normalizedPrefixes.join('」「') + '」を確認してください。'
            : '')
        );
        return;
      }

      sheets.forEach(function(sheet) {
        const lastCol = sheet.getLastColumn();
        if (lastCol === 0) return;
        matchedSheets++;

        sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].forEach(function(header) {
          const key = String(header || '').trim();
          if (!key || headerSet[key]) return;
          headerSet[key] = true;
          headers.push(key);
        });
      });
    } catch (error) {
      warnings.push(sourceId + ' のヘッダー取得に失敗しました: ' + trunc_(error && error.message || error, 80));
    }
  });

  if (matchedSheets === 0 && warnings.length === 0) {
    warnings.push('対象シートからヘッダー候補を取得できませんでした。');
  }

  return { headers: headers, warnings: warnings };
}

/**
 * 設定ダイアログの「ヘッダー候補を更新」ボタンから呼ばれる。
 * 設定シートから formSources と formSheetNamePrefix を読み取り、ヘッダー候補を返す（保存は行わない）。
 * @returns {{ headers: string[], warnings: string[] }}
 */
function fetchHeadersFromSources() {
  return withAppLock_(fetchHeadersFromSourcesUnlocked_);
}

function fetchHeadersFromSourcesUnlocked_() {
  const config = getConfig_();
  const urlList = (config.formSources || []).map(function(s) {
    return (s && s.url) ? s.url : String(s || '');
  });
  return collectAvailableSourceHeaders_(urlList, config.formSheetNamePrefix || '');
}

function scoreToGrade(v) {
  const n = Number(String(v == null ? '' : v).trim());
  switch (n) {
    case 5: return 'A';
    case 4: return 'B+';
    case 3: return 'B';
    case 2: return 'B-';
    case 1: return 'C';
    default: return '';
  }
}

function ensureSheet_(name, headers) {
  const ss = getAppSpreadsheet_();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  const lastCol = Math.max(sheet.getLastColumn(), headers.length, 1);
  const firstRow = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const hasHeader = firstRow.some(function(v) {
    return String(v || '').trim() !== '';
  });

  if (!hasHeader) {
    if (sheet.getLastRow() > 1) throw new Error('「' + name + '」の見出しが空です。既存データを確認してください。');
    sheet.clear();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function ensureErrorSheet_() {
  return ensureSheet_(ERROR_SHEET_NAME, [
    'タイムスタンプ',
    'シート名',
    '行番号',
    'メールアドレス',
    '送信状態',
    'エラー種別',
    'エラーメッセージ'
  ]);
}

function safeAlert_(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    Logger.log('[ALERT suppressed] ' + message);
  }
}

function trunc_(s, n) {
  const text = String(s == null ? '' : s).trim();
  return text.length > n ? text.slice(0, n) + '…' : text;
}

function pad2(n) {
  const num = Number(n);
  return (num < 10 ? '0' : '') + num;
}

function toYmdString(s) {
  const str = String(s == null ? '' : s).trim();
  if (!str) return '';

  let m = str.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (m) return m[1] + '/' + pad2(m[2]) + '/' + pad2(m[3]);

  m = str.match(/^\w+\s+([A-Za-z]{3})\s+(\d{1,2})\s+(\d{4})/);
  if (m) {
    const mon = {
      Jan: 1, Feb: 2, Mar: 3, Apr: 4, May: 5, Jun: 6,
      Jul: 7, Aug: 8, Sep: 9, Oct: 10, Nov: 11, Dec: 12
    }[m[1]];
    if (mon) return m[3] + '/' + pad2(mon) + '/' + pad2(m[2]);
  }

  return str;
}

function createHeaderMap_(headers) {
  const map = {};
  headers.forEach(function(header, index) {
    const key = String(header || '').trim();
    if (key) map[key] = index;
  });
  return map;
}

function makeStudentLookupKey_(courseId, email) {
  return String(courseId || '').trim() + '\t' + String(email || '').trim().toLowerCase();
}

function makeMappingLookupKey_(sourceId, sheetName) {
  return String(sourceId || '').trim() + '\t' + String(sheetName || '').trim();
}

/** 既存の元情報で照合する。元シートの物理的な並べ替えには対応しない。 */
function makeResponseKey_(sourceId, sheetName, rowNumber) {
  const id = String(sourceId || '').trim();
  const name = String(sheetName || '').trim();
  const row = Number(rowNumber);
  if (!id || !name || !Number.isInteger(row) || row < 2) return '';
  return JSON.stringify([id, name, row]);
}

function collectResponseKeys_(values, headerMap) {
  const keys = new Set();
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const key = makeResponseKey_(row[headerMap['元SS_ID']], row[headerMap['元シート名']], row[headerMap['元行番号']]);
    if (key) keys.add(key);
  }
  return keys;
}

function buildCourseIdFormulaForRow_(rowNum, classRecords) {
  const clauses = classRecords.map(function(item) {
    const className = String(item.className || '').replace(/"/g, '""');
    const courseId = String(item.courseId || '').replace(/"/g, '""');
    return 'D' + rowNum + '="' + className + '","' + courseId + '"';
  });
  return '=IF(D' + rowNum + '="","",IFS(' + clauses.join(',') + ',TRUE,""))';
}

function normalizeMultiline_(value) {
  return String(value == null ? '' : value)
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n');
}

/**
 * 文字列または配列を正規化して文字列配列を返す。
 * @param {*} value - 文字列、配列、またはその他
 * @param {string[]} [fallback=[]] - 結果が空の場合のデフォルト値
 * @returns {string[]}
 */
function normalizeStringArray_(value, fallback) {
  let arr;
  if (Array.isArray(value)) {
    arr = value.map(function(v) { return String(v || '').trim(); }).filter(Boolean);
  } else {
    const str = String(value != null ? value : '').trim();
    arr = str ? [str] : [];
  }
  return arr.length > 0 ? arr : (fallback || []);
}

/**
 * シート名がいずれかの接頭辞に前方一致するか判定する。
 * 接頭辞が空の場合は対象なしとする（意図せず全シートを取り込まない）。
 * @param {string} sheetName
 * @param {string[]} prefixes
 * @returns {boolean}
 */
function sheetMatchesPrefixes_(sheetName, prefixes) {
  const active = (prefixes || []).filter(function(p) { return !!p; });
  if (active.length === 0) return false;
  return active.some(function(p) { return sheetName.indexOf(p) === 0; });
}

function isManagedInternalSheet_(name, sourceId) {
  const ss=getAppSpreadsheet_();
  if(typeof ss.getId!=='function'||sourceId!==ss.getId())return false;
  return ['クラス一覧',STUDENT_SHEET_NAME,INPUT_SETTINGS_SHEET_NAME,CONFIG_SHEET_NAME,SETTINGS_SHEET_NAME,EVAL_SHEET_NAME,SEND_SHEET_NAME,ERROR_SHEET_NAME,MAPPING_SHEET_NAME,'フォーム管理'].includes(name);
}

function buildDefaultConfig_() {
  return {
    reminderTo: [], formSheetNamePrefix: [], formSources: [], fields: [], gradeScale: [],
    emailHeader: '', studentNameHeader: '', formStatusHeader: '', scoreSourceHeader: '',
    replyBodyHeader: '返信本文', messageTemplate: ''
  };
}

/** Only persisted pre-console configurations inherit historical missing values. */
function buildLegacyDefaultConfig_() {
  return {
    gradeScale: buildLegacyGradeScale_(),
    reminderTo: [],
    formSheetNamePrefix: ['フォームの回答'],
    emailHeader: 'メールアドレス',
    studentNameHeader: '名前',
    formStatusHeader: '送信状態',
    scoreSourceHeader: '評価',
    scoreFieldKey: 'score',
    replyBodyHeader: '返信本文',
    formSources: [],
    fields: [
      {
        key: 'lesson_date',
        sourceHeader: '授業の日付を入力してください。',
        evalHeader: '授業日付',
        sendHeader: '授業日付',
        type: 'date'
      },
      {
        key: 'submission',
        sourceHeader: '改行除去',
        evalHeader: '提出内容',
        sendHeader: '提出内容',
        type: 'text'
      },
      {
        key: 'score',
        sourceHeader: '評価',
        evalHeader: '評価',
        sendHeader: '評価',
        type: 'score_grade'
      },
      {
        key: 'reason',
        sourceHeader: '評価理由',
        evalHeader: '評価理由',
        sendHeader: '評価理由',
        type: 'text'
      },
      {
        key: 'advice',
        sourceHeader: '改善点',
        evalHeader: '改善点',
        sendHeader: '改善点',
        type: 'text'
      }
    ],
    messageTemplate: normalizeMultiline_(
      '評価：＜score＞\n\n' +
      '【理由】\n＜reason＞\n\n' +
      '【改善】\n＜advice＞\n\n' +
      '【提出内容（抜粋）】\n授業日付：＜lesson_date＞\n＜submission＞'
    )
  };
}

function normalizeFormSource_(item) {
  let raw = '';
  let courseId = '';

  if (item && typeof item === 'object') {
    raw = String(item.url || item.id || item.value || item.raw || '').trim();
    courseId = String(item.courseId || '').trim();
  } else {
    raw = String(item || '').trim();
  }
  if (!raw) return null;

  const match = raw.match(/^(.+?)\s*\|\s*([A-Za-z0-9_-]+)\s*$/);
  if (match) {
    raw = match[1].trim();
    if (!courseId) courseId = match[2].trim();
  }

  const id = extractSpreadsheetId_(raw);
  return {
    url: /^https?:\/\//i.test(raw)
      ? raw
      : ('https://docs.google.com/spreadsheets/d/' + id + '/edit'),
    id: id,
    courseId: courseId
  };
}

function formatFormSourceLine_(source) {
  const normalized = normalizeFormSource_(source);
  if (!normalized) return '';
  return normalized.url + (normalized.courseId ? ' | ' + normalized.courseId : '');
}

function normalizeFieldConfig_(field, fallbackIndex) {
  const normalized = {
    key: String(field && field.key || '').trim(),
    sourceHeader: String(field && field.sourceHeader || '').trim(),
    evalHeader: String(field && field.evalHeader || '').trim(),
    sendHeader: String(field && field.sendHeader || '').trim(),
    type: String(field && field.type || 'text').trim() || 'text'
  };

  if (!normalized.key) {
    normalized.key = 'field_' + (fallbackIndex + 1);
  }
  // Missing means legacy shared mapping; an explicit empty array stays unconfigured.
  if (field && Object.prototype.hasOwnProperty.call(field, 'gradeScale')) {
    normalized.gradeScale = normalizeGradeScale_(field.gradeScale);
  }

  return normalized;
}

function normalizeGradeScale_(gradeScale) {
  return (Array.isArray(gradeScale) ? gradeScale : []).map(function(entry) {
    return { from: String(entry.from).trim(), to: String(entry.to).trim() };
  });
}

function getFieldGradeScale_(field, config) {
  if (field && Object.prototype.hasOwnProperty.call(field, 'gradeScale')) {
    return Array.isArray(field.gradeScale) ? field.gradeScale : [];
  }
  return config && Array.isArray(config.gradeScale) ? config.gradeScale : [];
}

function getFieldEvalHeader_(field) {
  return String(field && (field.evalHeader || field.sourceHeader) || '').trim();
}

function getFieldSendHeader_(field) {
  return String(field && (field.sendHeader || field.sourceHeader) || '').trim();
}

function normalizeAppConfig_(rawConfig) {
  const defaults = buildDefaultConfig_();
  const config = rawConfig || {};
  const formSources = Array.isArray(config.formSources) ? config.formSources : [];
  const fields = Array.isArray(config.fields) ? config.fields : [];
  const normalizedFields = fields.map(function(field, index) { return normalizeFieldConfig_(field, index); });
  let scoreSourceHeader = String(config.scoreSourceHeader != null ? config.scoreSourceHeader : '').trim();

  if (config.scoreSourceHeader == null && config.scoreFieldKey != null) {
    const legacyField = getFieldByKey_({ fields: normalizedFields }, String(config.scoreFieldKey).trim());
    if (legacyField) scoreSourceHeader = legacyField.sourceHeader;
  }
  if (!scoreSourceHeader) {
    scoreSourceHeader = defaults.scoreSourceHeader;
  }

  return {
    reminderTo: normalizeStringArray_(config.reminderTo != null ? config.reminderTo : defaults.reminderTo),
    formSheetNamePrefix: normalizeStringArray_(config.formSheetNamePrefix != null ? config.formSheetNamePrefix : defaults.formSheetNamePrefix),
    emailHeader: String(config.emailHeader != null ? config.emailHeader : defaults.emailHeader).trim(),
    studentNameHeader: String(config.studentNameHeader != null ? config.studentNameHeader : defaults.studentNameHeader).trim(),
    formStatusHeader: String(config.formStatusHeader != null ? config.formStatusHeader : defaults.formStatusHeader).trim(),
    scoreSourceHeader: scoreSourceHeader,
    replyBodyHeader: String(config.replyBodyHeader != null ? config.replyBodyHeader : defaults.replyBodyHeader).trim(),
    messageTemplate: normalizeMultiline_(config.messageTemplate != null ? config.messageTemplate : defaults.messageTemplate),
    formSources: formSources.map(normalizeFormSource_).filter(function(item) { return !!item; }),
    fields: normalizedFields,
    gradeScale: normalizeGradeScale_(config.gradeScale)
  };
}

function validateAppConfig_(config) {
  validateConfigDraft_(config);
  // reminderTo は任意。リマインダー送信時のみチェックする。
  if (config.formSources.length === 0) {
    throw new Error('フォーム回答スプレッドシートURLを1件以上入力してください。');
  }
  if (config.formSheetNamePrefix.length === 0) {
    throw new Error('管理画面の「フォームと通知先」で対象シート名の接頭辞を入力してください。');
  }
  if (SEND_BASE_HEADERS.indexOf(config.replyBodyHeader) >= 0) {
    throw new Error('返信本文の列名が送信シートの固定列名と重複しています: ' + config.replyBodyHeader);
  }
  if (config.fields.length === 0) {
    throw new Error('取り込み項目を1件以上設定してください。');
  }
  if (!config.messageTemplate.trim()) {
    throw new Error('返信テンプレートを入力してください。');
  }
  config.fields.forEach(function(field) {
    if (field.type === 'score_grade' && !getFieldGradeScale_(field, config).length) {
      throw new Error('「' + field.sourceHeader + '」の評価変換表を1件以上設定してください。');
    }
  });
  ['emailHeader', 'studentNameHeader', 'formStatusHeader', 'scoreSourceHeader', 'replyBodyHeader'].forEach(function(key) {
    if (!config[key]) throw new Error('必要な列名が未設定です: ' + key);
  });

  const allowedTypes = { text: true, date: true, score_grade: true };
  const fieldKeySet = {};
  const evalHeaderSet = createHeaderSet_(EVAL_BASE_HEADERS);
  const sendHeaderSet = createHeaderSet_(SEND_BASE_HEADERS.concat([config.replyBodyHeader]));
  let hasScoreField = false;

  config.fields.forEach(function(field) {
    const evalHeader = getFieldEvalHeader_(field);
    const sendHeader = getFieldSendHeader_(field);

    if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(field.key)) {
      throw new Error('項目キーは半角英数字とアンダースコアで入力してください: ' + field.key);
    }
    if (!field.sourceHeader) {
      throw new Error('フォーム側ヘッダを入力してください: ' + field.key);
    }
    if (fieldKeySet[field.key]) {
      throw new Error('項目キーが重複しています: ' + field.key);
    }
    if (!allowedTypes[field.type]) {
      throw new Error('未対応の項目タイプです: ' + field.type);
    }
    if (evalHeaderSet[evalHeader]) {
      throw new Error('評価データの列名が重複しています: ' + evalHeader);
    }

    fieldKeySet[field.key] = true;
    evalHeaderSet[evalHeader] = true;

    if (sendHeader) {
      if (sendHeaderSet[sendHeader]) {
        throw new Error('送信シートの列名が重複しています: ' + sendHeader);
      }
      sendHeaderSet[sendHeader] = true;
    }
  });

  hasScoreField = !!getFieldBySourceHeader_(config, config.scoreSourceHeader);
  if (!hasScoreField) {
    throw new Error('採点済み判定に使うフォーム側ヘッダが fields に存在しません: ' + config.scoreSourceHeader);
  }
}

/** Strict input check before normalization; incomplete draft sections are allowed. */
function validateConfigDraft_(config) {
  if (!config || typeof config !== 'object' || Array.isArray(config)) throw new Error('設定の形式が不正です。');
  ['emailHeader', 'studentNameHeader', 'formStatusHeader', 'scoreSourceHeader', 'scoreFieldKey', 'replyBodyHeader', 'messageTemplate'].forEach(function(key) {
    if (config[key] !== undefined && typeof config[key] !== 'string') throw new Error('文字列で入力してください: ' + key);
  });
  ['reminderTo', 'formSheetNamePrefix', 'formSources', 'fields', 'gradeScale'].forEach(function(key) {
    if (config[key] !== undefined && !Array.isArray(config[key])) throw new Error('配列で入力してください: ' + key);
  });
  if (config.gradeScale !== undefined) validateGradeScaleDraft_(config.gradeScale);
  (config.reminderTo || []).forEach(function(value) {
    if (typeof value !== 'string' || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value.trim())) throw new Error('通知先メールアドレスが不正です。');
  });
  (config.formSheetNamePrefix || []).forEach(function(value) {
    if (typeof value !== 'string' || !value.trim()) throw new Error('対象シート名の接頭辞が不正です。');
  });
  (config.formSources || []).forEach(function(value) {
    if (typeof value !== 'string') {
      if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error('参照元URLの形式が不正です。');
      ['url', 'id', 'value', 'raw', 'courseId'].forEach(function(key) {
        if (value[key] !== undefined && typeof value[key] !== 'string') throw new Error('参照元URLの形式が不正です。');
      });
    }
    if (!normalizeFormSource_(value)) throw new Error('参照元URLが空です。');
  });
  const keys = new Set(['student_name', 'class_name', 'course_id', 'row_no', 'prototype'].concat(Object.getOwnPropertyNames(Object.prototype)));
  const evalHeaders = new Set(EVAL_BASE_HEADERS);
  const sendHeaders = new Set(SEND_BASE_HEADERS);
  const reply = String(config.replyBodyHeader || '').trim();
  if (reply && sendHeaders.has(reply)) throw new Error('返信本文の列名が固定列名と重複しています: ' + reply);
  if (reply) sendHeaders.add(reply);
  (config.fields || []).forEach(function(field) {
    if (!field || typeof field !== 'object' || Array.isArray(field)) throw new Error('項目の形式が不正です。');
    ['key', 'sourceHeader', 'evalHeader', 'sendHeader', 'type'].forEach(function(key) {
      if (field[key] !== undefined && typeof field[key] !== 'string') throw new Error('項目は文字列で入力してください: ' + key);
    });
    if (Object.prototype.hasOwnProperty.call(field, 'gradeScale')) validateGradeScaleDraft_(field.gradeScale);
    const key = String(field.key || '').trim();
    if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(key)) throw new Error('項目キーが不正です: ' + key);
    if (keys.has(key)) throw new Error('項目キーが重複または予約済みです: ' + key);
    keys.add(key);
    if (['text', 'date', 'score_grade'].indexOf(field.type || 'text') < 0) throw new Error('未対応の項目タイプです: ' + field.type);
    if (!String(field.sourceHeader || '').trim()) throw new Error('フォーム側ヘッダを入力してください: ' + key);
    const evalHeader = getFieldEvalHeader_(field);
    const sendHeader = getFieldSendHeader_(field);
    if (evalHeaders.has(evalHeader) || sendHeaders.has(sendHeader)) throw new Error('項目の列名が重複しています: ' + key);
    evalHeaders.add(evalHeader); sendHeaders.add(sendHeader);
  });
}

function createHeaderSet_(headers) {
  const set = {};
  headers.forEach(function(header) {
    set[String(header || '').trim()] = true;
  });
  return set;
}

function extractSpreadsheetId_(value) {
  const text = String(value == null ? '' : value).trim();
  if (!text) {
    throw new Error('スプレッドシートURLまたはIDが空です。');
  }

  const match = text.match(/^https:\/\/docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9_-]+)(?:[/?#]|$)/);
  if (match) return match[1];

  if (/^[a-zA-Z0-9-_]{20,}$/.test(text)) {
    return text;
  }

  throw new Error('スプレッドシートURLまたはIDの形式が不正です: ' + trunc_(text, 80));
}

/** Called only under the app lock. Persist and verify before deleting legacy tabs. */
function migrateConfigStorageUnlocked_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('APP_CONFIG_STORAGE') !== 'properties-v1') {
    const config = getConfig_();
    validateConfigDraft_(config);
    saveConfig_(config);
  }
  if (props.getProperty('APP_CONFIG_SHEETS_RETIRED') === '1') return;
  // Do not delete the old stores when recovery or verification is incomplete.
  const config = getConfig_();
  validateConfigDraft_(config);
  if (props.getProperty('APP_CONFIG_PARSE_ERROR') === '1' ||
      props.getProperty('APP_CONFIG') !== props.getProperty('APP_CONFIG_BACKUP')) {
    throw new Error('内部設定の保存確認ができないため、旧設定シートを保持しています。');
  }
  const ss = getAppSpreadsheet_();
  const names = [CONFIG_SHEET_NAME, INPUT_SETTINGS_SHEET_NAME, SETTINGS_SHEET_NAME];
  const obsolete = names.map(function(name) { return ss.getSheetByName(name); }).filter(Boolean);
  if (obsolete.length) {
    // Sheets requires at least one visible tab. Preserve all unrelated data.
    let visible = ss.getSheets().find(function(sheet) { return names.indexOf(sheet.getName()) < 0 && !sheet.isSheetHidden(); });
    if (!visible) {
      visible = ss.getSheetByName('クラス一覧') || ensureSheet_('クラス一覧', ['クラス名', 'コースID', '同期対象(1)']);
      visible.showSheet();
    }
    obsolete.forEach(function(sheet) { ss.deleteSheet(sheet); });
    SpreadsheetApp.flush();
  }
  props.setProperty('APP_CONFIG_SHEETS_RETIRED', '1');
}

/**
 * Config シートから CONFIG_JSON を読み込む。
 * @returns {Object|null} パース済みオブジェクト、または null
 */
function readConfigSheet_() {
  const ss = getAppSpreadsheet_();
  const sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet || sheet.getLastRow() === 0) return null;
  const json = String(sheet.getRange(1, 1).getValue() || '').trim();
  if (!json) return null;
  try {
    return parseStoredConfig_(json);
  } catch (e) {
    throw new Error('旧Configの設定を読み取れません。旧シートを保持して復旧してください。' + e);
  }
}

/**
 * 設定シートから REMINDER_TO・FORM_SHEET_PREFIX・FORM_SS を読み込む。
 * 存在するキーの空欄は明示的な解除。キー自体がない旧シートのみ保存値にフォールバックする。
 * @returns {Object} 見つかった値だけを含む部分的な config オブジェクト
 */
function readSettingsInputSheet_() {
  const ss = getAppSpreadsheet_();
  const sheet = ss.getSheetByName(INPUT_SETTINGS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return {};

  const values = sheet.getDataRange().getDisplayValues();
  const result = {};
  const reminderTos = [];
  const prefixes = [];
  const formSources = [];
  const presentKeys = new Set();

  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0] || '').trim();
    const value = String(values[i][1] || '').trim();
    if (!key) continue;
    presentKeys.add(key);
    if (!value) continue;

    switch (key) {
      case 'REMINDER_TO':
        reminderTos.push(value);
        break;
      case 'FORM_SHEET_PREFIX':
        prefixes.push(value);
        break;
      case 'FORM_SS': {
        const normalized = normalizeFormSource_(value);
        if (normalized) formSources.push(normalized);
        break;
      }
    }
  }

  if (presentKeys.has('REMINDER_TO')) result.reminderTo = reminderTos;
  if (presentKeys.has('FORM_SHEET_PREFIX')) result.formSheetNamePrefix = prefixes;
  if (presentKeys.has('FORM_SS')) result.formSources = formSources;

  return result;
}

/** Read-only legacy compatibility until the first successful internal save. */
function loadConfig_() {
  const props = PropertiesService.getScriptProperties();
  const internal = props.getProperty('APP_CONFIG_STORAGE') === 'properties-v1';
  const overrides = internal ? {} : readSettingsInputSheet_();
  const json = props.getProperty('APP_CONFIG');
  if (json) {
    try { return normalizeSavedConfig_(parseStoredConfig_(json), overrides); }
    catch (error) { props.setProperty('APP_CONFIG_PARSE_ERROR', '1'); }
  }
  if (internal) {
    props.setProperty('APP_CONFIG_PARSE_ERROR', '1');
    const backup = props.getProperty('APP_CONFIG_BACKUP');
    if (backup) return normalizeSavedConfig_(parseStoredConfig_(backup), {});
    throw new Error('内部設定を読み取れません。設定JSONまたは内部バックアップから復旧してください。');
  }
  const legacy = readConfigSheet_();
  if (legacy) return normalizeSavedConfig_(legacy, overrides);
  if (json) throw new Error('保存済み設定を読み取れません。旧設定を保持して復旧してください。');
  return normalizeAppConfig_(Object.assign({}, buildDefaultConfig_(), overrides));
}

function parseStoredConfig_(json) {
  const value = JSON.parse(json);
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error('設定の形式が不正です。');
  validateConfigDraft_(value);
  return value;
}

function normalizeSavedConfig_(saved, overrides) {
  const merged = Object.assign({}, buildLegacyDefaultConfig_(), saved, overrides);
  if (saved.gradeScale === undefined && Array.isArray(saved.fields) && !saved.fields.some(function(field) { return field && field.type === 'score_grade'; })) merged.gradeScale = [];
  if (saved.scoreSourceHeader === undefined && saved.scoreFieldKey !== undefined) delete merged.scoreSourceHeader;
  return normalizeAppConfig_(merged);
}

/** Count UTF-8 bytes without requiring external services. */
function utf8ByteLength_(text) {
  let bytes = 0;
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    if (code < 0x80) bytes++;
    else if (code < 0x800) bytes += 2;
    else if (code >= 0xD800 && code <= 0xDBFF && i + 1 < text.length &&
        text.charCodeAt(i + 1) >= 0xDC00 && text.charCodeAt(i + 1) <= 0xDFFF) {
      bytes += 4;
      i++;
    } else {
      // BMP characters and replacement encoding for an unmatched surrogate.
      bytes += 3;
    }
  }
  return bytes;
}

function validateGradeScaleDraft_(gradeScale) {
  if (!Array.isArray(gradeScale)) throw new Error('評価変換表は配列で入力してください。');
  const gradeKeys = new Set();
  gradeScale.forEach(function(entry) {
    if (!entry || typeof entry !== 'object' || Array.isArray(entry) ||
        typeof entry.from !== 'string' || typeof entry.to !== 'string' || !entry.from.trim() || !entry.to.trim()) {
      throw new Error('評価変換表の変換前・変換後は空でない文字列を入力してください。');
    }
    const key = normalizeGradeScaleKey_(entry.from);
    if (gradeKeys.has(key)) throw new Error('評価変換表の変換前が重複しています: ' + entry.from);
    gradeKeys.add(key);
  });
}

/** Internal primary + recovery copy. Callers hold the app lock. */
function saveConfig_(config) {
  const json = JSON.stringify(config), byteLength = utf8ByteLength_(json);
  if (byteLength > 8000) throw new Error('設定データが大きすぎます（' + byteLength + ' バイト）。フォーム回答スプレッドシートの件数や取り込み項目の数を減らしてください。');
  const props = PropertiesService.getScriptProperties();
  const keys = ['APP_CONFIG', 'APP_CONFIG_BACKUP', 'APP_CONFIG_STORAGE', 'APP_CONFIG_PARSE_ERROR', 'APP_CONFIG_SAVE_ERROR'];
  const previous = keys.map(function(key) { return [key, props.getProperty(key)]; });
  try {
    props.setProperty('APP_CONFIG_BACKUP', json);
    props.setProperty('APP_CONFIG', json);
    if (props.getProperty('APP_CONFIG') !== json || props.getProperty('APP_CONFIG_BACKUP') !== json) throw new Error('内部設定の保存を確認できません。');
    props.setProperty('APP_CONFIG_STORAGE', 'properties-v1');
    if (props.getProperty('APP_CONFIG_STORAGE') !== 'properties-v1') throw new Error('内部設定の移行を確認できません。');
    props.deleteProperty('APP_CONFIG_PARSE_ERROR');
    props.deleteProperty('APP_CONFIG_SAVE_ERROR');
  } catch (error) {
    const failures = [];
    previous.forEach(function(entry) {
      try {
        if (entry[1] === null) props.deleteProperty(entry[0]);
        else props.setProperty(entry[0], entry[1]);
        if (props.getProperty(entry[0]) !== entry[1]) throw new Error('復元値が一致しません');
      } catch (rollbackError) { failures.push(entry[0] + ': ' + rollbackError); }
    });
    if (failures.length) {
      try { props.setProperty('APP_CONFIG_SAVE_ERROR', failures.join('\n')); } catch (ignored) { Logger.log(ignored); }
      throw new Error('設定保存に失敗し、復旧も完了していません。自動実行を停止し、内部設定と設定JSONを確認してください。' + error + '\n' + failures.join('\n'));
    }
    throw error;
  }
}

function snapshotSheetContents_(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) return { name: name, exists: false };
  const range = sheet.getDataRange();
  const values = range.getValues();
  const formulas = typeof range.getFormulas === 'function' ? range.getFormulas() : [];
  formulas.forEach(function(row, r) { row.forEach(function(formula, c) { if (formula) values[r][c] = formula; }); });
  return { name: name, exists: true, values: values };
}

function restoreSheetContents_(ss, snapshot) {
  let sheet = ss.getSheetByName(snapshot.name);
  if (!snapshot.exists) {
    if (sheet) ss.deleteSheet(sheet);
    return;
  }
  if (!sheet) sheet = ss.insertSheet(snapshot.name);
  sheet.clearContents();
  if (snapshot.values.length && snapshot.values[0].length) sheet.getRange(1, 1, snapshot.values.length, snapshot.values[0].length).setValues(snapshot.values);
}

function getConfig_() {
  const saveError = PropertiesService.getScriptProperties().getProperty('APP_CONFIG_SAVE_ERROR');
  if (saveError) throw new Error('前回の設定保存からの復旧が未完了です。内部設定と設定JSONを確認してください。' + saveError);
  return loadConfig_();
}

function getConfiguredEvalHeaders_(config) {
  const headers = EVAL_BASE_HEADERS.slice();
  config.fields.forEach(function(field) {
    headers.push(getFieldEvalHeader_(field));
  });
  return uniqueHeaders_(headers);
}

function getConfiguredSendHeaders_(config) {
  const headers = SEND_BASE_HEADERS.slice();
  config.fields.forEach(function(field) {
    headers.push(getFieldSendHeader_(field));
  });
  headers.push(config.replyBodyHeader);
  return uniqueHeaders_(headers);
}

function uniqueHeaders_(headers) {
  const seen = {};
  const result = [];
  headers.forEach(function(header) {
    const key = String(header || '').trim();
    if (!key || seen[key]) return;
    seen[key] = true;
    result.push(key);
  });
  return result;
}

function hasMatchingHeaders_(sheet, expectedHeaders) {
  const lastCol = Math.max(sheet.getLastColumn(), expectedHeaders.length, 1);
  const current = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0]
    .slice(0, expectedHeaders.length)
    .map(function(value) { return String(value || '').trim(); });

  if (current.length < expectedHeaders.length) return false;
  for (let i = 0; i < expectedHeaders.length; i++) {
    if (current[i] !== expectedHeaders[i]) return false;
  }
  return true;
}

function ensureConfiguredSheet_(name, headers) {
  const ss = getAppSpreadsheet_();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return sheet;
  }

  const hasContent = sheet.getLastRow() > 0 || sheet.getLastColumn() > 0;
  if (!hasContent) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return sheet;
  }

  if (hasMatchingHeaders_(sheet, headers)) {
    return sheet;
  }

  if (sheet.getLastRow() <= 1) {
    sheet.clear();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return sheet;
  }

  throw new Error('「' + name + '」シートの構成が設定と一致しません。管理画面で必要なデータを退避・クリアしてから再実行してください。');
}

function ensureEvalSheet_(config) {
  return ensureConfiguredSheet_(EVAL_SHEET_NAME, getConfiguredEvalHeaders_(config));
}

function ensureSendSheet_(config) {
  return ensureConfiguredSheet_(SEND_SHEET_NAME, getConfiguredSendHeaders_(config));
}

function ensureStudentSheetForSync_() {
  const ss = getAppSpreadsheet_();
  let sheet = ss.getSheetByName(STUDENT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(STUDENT_SHEET_NAME);
  }

  if (!hasMatchingHeaders_(sheet, STUDENT_SHEET_HEADERS)) {
    if (sheet.getLastRow() > 0) throw new Error('「生徒一覧」の構成が一致しません。既存データを退避してから見出しを修正してください。');
    sheet.clear();
    sheet.getRange(1, 1, 1, STUDENT_SHEET_HEADERS.length).setValues([STUDENT_SHEET_HEADERS]);
  }

  sheet.setFrozenRows(1);
  return sheet;
}

function getResetImpactSummary_() {
  const ss = getAppSpreadsheet_();
  const evalSheet = ss.getSheetByName(EVAL_SHEET_NAME);
  const sendSheet = ss.getSheetByName(SEND_SHEET_NAME);

  const evalRows = evalSheet ? Math.max(evalSheet.getLastRow() - 1, 0) : 0;
  const sendRows = sendSheet ? Math.max(sendSheet.getLastRow() - 1, 0) : 0;

  let pendingSend = 0;
  let retryableSend = 0;
  let errorSend = 0;

  if (sendSheet && sendRows > 0) {
    const values = sendSheet.getDataRange().getDisplayValues();
    const headers = values[0].map(String);
    const idxStatus = headers.indexOf('送信状態');
    if (idxStatus >= 0) {
      for (let i = 1; i < values.length; i++) {
        const status = String(values[i][idxStatus] || '').trim();
        if (!status) pendingSend++;
        if (status === '未') retryableSend++;
        if (status === 'エラー') errorSend++;
      }
    }
  }

  return {
    evalRows: evalRows,
    sendRows: sendRows,
    pendingSend: pendingSend,
    retryableSend: retryableSend,
    errorSend: errorSend
  };
}

function buildResetSummaryText_(info) {
  const lines = [
    '設定を保存すると、次のシートを再生成します。',
    '・評価データ: ' + info.evalRows + '件を削除',
    '・送信シート: ' + info.sendRows + '件を削除'
  ];

  if (info.pendingSend > 0 || info.retryableSend > 0 || info.errorSend > 0) {
    lines.push('');
    lines.push('送信シートの内訳');
    lines.push('・未送信: ' + info.pendingSend + '件');
    lines.push('・再送待ち: ' + info.retryableSend + '件');
    lines.push('・エラー: ' + info.errorSend + '件');
  }

  lines.push('');
  lines.push('この操作は取り消せません。');
  lines.push('続行する場合は保存を確定してください。');

  return lines.join('\n');
}

function resetConfiguredSheets_(config) {
  resetSheetWithHeaders_(EVAL_SHEET_NAME, getConfiguredEvalHeaders_(config));
  resetSheetWithHeaders_(SEND_SHEET_NAME, getConfiguredSendHeaders_(config));
}

function resetSheetWithHeaders_(name, headers) {
  const ss = getAppSpreadsheet_();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  } else {
    sheet.clear();
  }

  if (sheet.getMaxColumns() > headers.length) {
    sheet.deleteColumns(headers.length + 1, sheet.getMaxColumns() - headers.length);
  }
  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
}

function getFieldByKey_(config, key) {
  for (let i = 0; i < config.fields.length; i++) {
    if (config.fields[i].key === key) return config.fields[i];
  }
  return null;
}

function getFieldBySourceHeader_(config, sourceHeader) {
  const target = String(sourceHeader || '').trim();
  for (let i = 0; i < config.fields.length; i++) {
    if (String(config.fields[i].sourceHeader || '').trim() === target) return config.fields[i];
  }
  return null;
}

function resolveScoreField_(config) {
  return getFieldBySourceHeader_(config, config.scoreSourceHeader)
    || getFieldByKey_(config, config.scoreFieldKey);
}

function buildLegacyGradeScale_() {
  return [{ from: '5', to: 'A' }, { from: '4', to: 'B+' }, { from: '3', to: 'B' },
    { from: '2', to: 'B-' }, { from: '1', to: 'C' }];
}

/** Decimal strings compare numerically without Number rounding; other strings are case-sensitive. */
function normalizeGradeScaleKey_(value) {
  const text = String(value == null ? '' : value).trim();
  if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$/.test(text)) return 'text:' + text;
  const negative = text.charAt(0) === '-';
  const parts = text.replace(/^[+-]/, '').split('.');
  const integer = (parts[0] || '0').replace(/^0+(?=\d)/, '');
  const fraction = (parts[1] || '').replace(/0+$/, '');
  return 'number:' + (negative && (integer !== '0' || fraction) ? '-' : '') + integer + (fraction ? '.' + fraction : '');
}

function processConfiguredFieldValue_(field, rawValue, displayValue, config) {
  if (field.type === 'date') {
    return toYmdString(displayValue || rawValue);
  }
  if (field.type === 'score_grade') {
    if (!config && !Object.prototype.hasOwnProperty.call(field, 'gradeScale')) return scoreToGrade(rawValue) || String(rawValue == null ? '' : rawValue).trim();
    const key = normalizeGradeScaleKey_(rawValue);
    const entry = getFieldGradeScale_(field, config).find(function(item) { return normalizeGradeScaleKey_(item.from) === key; });
    return entry ? String(entry.to).trim() : String(rawValue == null ? '' : rawValue).trim();
  }
  return String(rawValue == null ? '' : rawValue).trim();
}

function buildTemplateLabelMap_(config) {
  const labels = {
    student_name: '名前',
    class_name: 'クラス名',
    course_id: 'コースID',
    row_no: 'No'
  };

  (config && config.fields || []).forEach(function(field) {
    if (!field || !field.key) return;
    labels[field.key] = String(field.sendHeader || field.evalHeader || field.sourceHeader || field.key).trim();
  });

  return labels;
}

function friendlyPlaceholderPattern_() {
  // Keep legacy angle tokens and explicit {{sections}} separate from single braces.
  return /[＜<]\s*([A-Za-z_][A-Za-z0-9_]*)\s*[＞>]|(?<!\{)\{\s*([A-Za-z_][A-Za-z0-9_]*)\s*\}(?!\})/g;
}

function friendlyPlaceholderOnlyPattern_() {
  return /^(?:[＜<]\s*[A-Za-z_][A-Za-z0-9_]*\s*[＞>]|\{\s*[A-Za-z_][A-Za-z0-9_]*\s*\})$/;
}

function replaceFriendlyPlaceholdersOnly_(text) {
  return normalizeMultiline_(text || '').replace(friendlyPlaceholderPattern_(), function(_, legacyKey, key) {
    return '{{' + (legacyKey || key) + '}}';
  });
}

function convertFriendlyTemplateToSections_(template, config) {
  const labels = buildTemplateLabelMap_(config);
  const normalized = normalizeMultiline_(template || '').trim();
  if (!normalized) return '';

  return normalized.split(/\n{2,}/).map(function(block) {
    const lines = block.split('\n');
    const lineInfos = lines.map(function(line) {
      const keys = [];
      line.replace(friendlyPlaceholderPattern_(), function(_, legacyKey, key) {
        keys.push(legacyKey || key);
        return _;
      });

      return {
        raw: line,
        keys: keys,
        replaced: replaceFriendlyPlaceholdersOnly_(line),
        placeholderOnly: keys.length === 1 && friendlyPlaceholderOnlyPattern_().test(line.trim())
      };
    });

    const uniqueKeys = {};
    const orderedKeys = [];
    lineInfos.forEach(function(info) {
      info.keys.forEach(function(key) {
        if (uniqueKeys[key]) return;
        uniqueKeys[key] = true;
        orderedKeys.push(key);
      });
    });

    let mainKey = '';
    lineInfos.forEach(function(info) {
      if (info.placeholderOnly) mainKey = info.keys[0];
    });
    if (!mainKey && orderedKeys.length === 1) {
      mainKey = orderedKeys[0];
    }

    const convertedLines = lineInfos.map(function(info) {
      if (info.keys.length === 0) return info.raw;

      if (info.placeholderOnly) {
        const key = info.keys[0];
        if (key === mainKey) {
          return '{{' + key + '}}';
        }
        const label = labels[key];
        const content = label ? (label + '：{{' + key + '}}') : ('{{' + key + '}}');
        return '{{#' + key + '}}' + content + '{{/' + key + '}}';
      }

      if (info.keys.length === 1) {
        const key = info.keys[0];
        if (key === mainKey) {
          return info.replaced;
        }
        return '{{#' + key + '}}' + info.replaced + '{{/' + key + '}}';
      }

      return info.replaced;
    });

    const joined = convertedLines.join('\n');
    if (!mainKey) return joined;
    return '{{#' + mainKey + '}}' + joined + '{{/' + mainKey + '}}';
  }).join('\n\n');
}

function renderTemplateWithSections_(template, context, config) {
  let result = normalizeMultiline_(template || '');
  if (result.indexOf('{{') >= 0) {
    result = replaceFriendlyPlaceholdersOnly_(result);
  } else {
    result = convertFriendlyTemplateToSections_(result, config);
  }
  const sectionPattern = /{{#([A-Za-z_][A-Za-z0-9_]*)}}([\s\S]*?){{\/\1}}/g;

  let guard = 0;
  while (guard < 20) {
    sectionPattern.lastIndex = 0;
    if (!sectionPattern.test(result)) break;
    sectionPattern.lastIndex = 0;
    result = result.replace(sectionPattern, function(_, key, inner) {
      const value = String(context[key] == null ? '' : context[key]).trim();
      return value ? inner : '';
    });
    guard++;
  }

  result = result.replace(/{{([A-Za-z_][A-Za-z0-9_]*)}}/g, function(_, key) {
    return String(context[key] == null ? '' : context[key]);
  });

  return result
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

/** A worker owns this cache for one invocation; never persist it between runs. */
function createFormStatusCache_() {
  return { spreadsheets: new Map(), targets: new Map() };
}

function updateFormStatus_(config, ssId, sheetName, rowIndex, status, invocationCache) {
  if (!ssId || !sheetName || !rowIndex) return;

  try {
    const cache = invocationCache || createFormStatusCache_();
    if (!cache.spreadsheets.has(ssId)) cache.spreadsheets.set(ssId, SpreadsheetApp.openById(ssId));
    const key = JSON.stringify([ssId, sheetName, config.formStatusHeader]);
    if (!cache.targets.has(key)) {
      const sheet = cache.spreadsheets.get(ssId).getSheetByName(sheetName);
      const headers = sheet && sheet.getLastColumn() > 0
        ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(String) : [];
      cache.targets.set(key, { sheet: sheet, statusColumn: headers.indexOf(config.formStatusHeader) + 1 });
    }
    const target = cache.targets.get(key);
    if (!target.sheet) throw new Error('元回答シートが見つかりません: ' + sheetName);
    if (!target.statusColumn) throw new Error('元回答シートの状態列が見つかりません: ' + config.formStatusHeader);
    target.sheet.getRange(rowIndex, target.statusColumn).setValue(status);
  } catch (e) {
    ensureErrorSheet_().appendRow([
      new Date(),
      sheetName || 'Form更新',
      rowIndex,
      '',
      status,
      'FORM_UPDATE_ERROR',
      String(e)
    ]);
  }
}

function getStudentSheet_() {
  const ss = getAppSpreadsheet_();
  const sheet = ss.getSheetByName(STUDENT_SHEET_NAME);
  if (!sheet) {
    throw new Error('生徒一覧シートが見つかりません: ' + STUDENT_SHEET_NAME);
  }
  return sheet;
}

function getSelectedClassRecords_() {
  const classSheet = ensureSheet_('クラス一覧', ['クラス名', 'コースID', '同期対象(1)']);
  const lastRow = classSheet.getLastRow();
  if (lastRow < 2) {
    throw new Error('クラス一覧が空です。設定の「クラスと生徒」で一覧を取得してください。');
  }

  const values = classSheet.getRange(2, 1, lastRow - 1, 3).getDisplayValues();
  const classes = values.map(function(row) {
    return {
      className: String(row[0] || '').trim(),
      courseId: String(row[1] || '').trim(),
      syncFlag: String(row[2] || '').trim()
    };
  }).filter(function(item) {
    return item.courseId && item.syncFlag === '1';
  });

  if (classes.length === 0) {
    throw new Error('クラス一覧で「同期対象(1)」が選択されていません。先に対象クラスへ 1 を入れてください。');
  }

  return classes;
}

function collectFormTargetSheets_(config) {
  const prefixes = Array.isArray(config && config.formSheetNamePrefix) ? config.formSheetNamePrefix : [];
  const formSources = Array.isArray(config && config.formSources) ? config.formSources : [];
  if (formSources.length === 0) {
    throw new Error('管理画面の「フォームと通知先」にフォーム回答スプレッドシートを登録してください。');
  }

  const seen = {};
  const rows = [];

  formSources.forEach(function(source) {
    const sourceId = String(source && source.id || '').trim();
    if (!sourceId) return;

    const formSs = SpreadsheetApp.openById(sourceId);
    const spreadsheetName = formSs.getName();
    formSs.getSheets().forEach(function(sheet) {
      const sheetName = sheet.getName();
      if (isManagedInternalSheet_(sheetName, sourceId) || !sheetMatchesPrefixes_(sheetName, prefixes)) return;

      const key = makeMappingLookupKey_(sourceId, sheetName);
      if (seen[key]) return;
      seen[key] = true;
      rows.push({
        sourceId: sourceId,
        spreadsheetName: spreadsheetName,
        sheetName: sheetName
      });
    });
  });

  if (rows.length === 0) {
    const activePrefixes = prefixes.filter(function(p) { return !!p; });
    throw new Error(
      'FORM_SS に対象シートがありません。' +
      (activePrefixes.length > 0
        ? ' 接頭辞「' + activePrefixes.join('」「') + '」を確認してください。'
        : '')
    );
  }

  return rows;
}

function getMappingSheet_() {
  const ss = getAppSpreadsheet_();
  return ss.getSheetByName(MAPPING_SHEET_NAME);
}

function readMappingEntries_(strict) {
  const sheet = getMappingSheet_();
  const map = new Map();
  if (!sheet || sheet.getLastRow() < 2) {
    return map;
  }

  const values = sheet.getDataRange().getDisplayValues();
  const headers = values[0].map(String);
  const headerMap = createHeaderMap_(headers);
  ['元SS_ID', '元シート名', 'クラス名', 'courseId', 'メモ'].forEach(function(header) {
    if (!(header in headerMap)) {
      if (!strict) return;
      throw new Error('対応表シートのヘッダが不正です。設定の「回答先と通知先」で対応表を確認してください。');
    }
  });
  if (!('元SS_ID' in headerMap) || !('元シート名' in headerMap) || !('クラス名' in headerMap) || !('courseId' in headerMap) || !('メモ' in headerMap)) {
    return map;
  }

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const sourceId = String(row[headerMap['元SS_ID']] || '').trim();
    const sheetName = String(row[headerMap['元シート名']] || '').trim();
    if (!sourceId || !sheetName) continue;

    map.set(makeMappingLookupKey_(sourceId, sheetName), {
      className: String(row[headerMap['クラス名']] || '').trim(),
      courseId: String(row[headerMap['courseId']] || '').trim(),
      note: String(row[headerMap['メモ']] || '').trim()
    });
  }

  return map;
}

function createMappingSheet() {
  const message = withAppLock_(createMappingSheetUnlocked_);
  safeAlert_(message);
}

function createMappingSheetUnlocked_() {
  const config = getConfig_();
  const classRecords = getSelectedClassRecords_();
  const targetSheets = collectFormTargetSheets_(config);
  const existingMap = readMappingEntries_(false);

  const classNameSet = {};
  classRecords.forEach(function(item) {
    const className = String(item.className || '').trim();
    if (!className) return;
    if (classNameSet[className]) {
      throw new Error('同期対象クラスに同名のクラス名があります。対応表のドロップダウンに使うため、クラス名が重複しない状態で実行してください: ' + className);
    }
    classNameSet[className] = true;
  });

  const ss = getAppSpreadsheet_();
  let sheet = ss.getSheetByName(MAPPING_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(MAPPING_SHEET_NAME);
  } else {
    const lastRow = Math.max(sheet.getLastRow(), 1);
    const lastCol = Math.max(sheet.getLastColumn(), MAPPING_SHEET_HEADERS.length);
    sheet.getRange(1, 1, lastRow, lastCol).clearDataValidations();
    sheet.clear();
  }

  const rows = [MAPPING_SHEET_HEADERS];
  targetSheets.forEach(function(item) {
    const existing = existingMap.get(makeMappingLookupKey_(item.sourceId, item.sheetName)) || {};
    rows.push([
      item.sourceId,
      item.spreadsheetName,
      item.sheetName,
      existing.className || '',
      '',
      existing.note || ''
    ]);
  });

  sheet.getRange(1, 1, rows.length, MAPPING_SHEET_HEADERS.length).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, MAPPING_SHEET_HEADERS.length);

  const numRows = rows.length - 1;
  if (numRows > 0) {
    const classNames = classRecords.map(function(item) { return item.className; });
    const validation = SpreadsheetApp.newDataValidation()
      .requireValueInList(classNames, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, 4, numRows, 1).setDataValidation(validation);

    const formulas = [];
    for (let i = 0; i < numRows; i++) {
      const rowNum = i + 2;
      formulas.push([buildCourseIdFormulaForRow_(rowNum, classRecords)]);
    }
    sheet.getRange(2, 5, numRows, 1).setFormulas(formulas);
  }

  return '対応表シートを作成/更新しました（' + targetSheets.length + '件）';
}

function appendRows_(sheet, rows) {
  if (!rows || rows.length === 0) return;
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
}

function classroomdata() {
  return withAppLock_(classroomdataUnlocked_);
}

function classroomdataUnlocked_() {
  const ss = getAppSpreadsheet_();
  const existing = ss.getSheetByName('クラス一覧');
  const flags = new Map();
  if (existing && existing.getLastRow() > 0) {
    if (!hasMatchingHeaders_(existing, ['クラス名', 'コースID', '同期対象(1)'])) throw new Error('クラス一覧の構成が一致しません。見出しを確認してください。');
    existing.getDataRange().getValues().slice(1).forEach(function(row) { flags.set(String(row[1] || '').trim(), row[2]); });
  }
  const courses = [];
  let pageToken = '';
  do {
    const response = Classroom.Courses.list({ teacherId: Session.getActiveUser(), courseStates: 'ACTIVE', pageSize: 100, pageToken: pageToken });
    if (!response || typeof response !== 'object' || (response.courses != null && !Array.isArray(response.courses))) throw new Error('クラス一覧の取得結果が不正です。');
    Array.prototype.push.apply(courses, response.courses || []);
    pageToken = response.nextPageToken || '';
  } while (pageToken);
  replaceRosterRows_('クラス一覧', ['クラス名', 'コースID', '同期対象(1)'], courses.map(function(course) {
    const id = String(course.id);
    return [course.name, id, flags.has(id) ? flags.get(id) : ''];
  }));
  return 'クラス一覧を更新しました（' + courses.length + '件）。同期対象は保持しています。';
}

function replaceRosterRows_(name, headers, rows) {
  const ss = getAppSpreadsheet_();
  const snapshot = snapshotSheetContents_(ss, name);
  try {
    const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
    sheet.clearContents();
    sheet.getRange(1, 1, rows.length + 1, headers.length).setValues([headers].concat(rows));
    sheet.setFrozenRows(1);
    SpreadsheetApp.flush();
  } catch (error) {
    try { restoreSheetContents_(ss, snapshot); SpreadsheetApp.flush(); }
    catch (rollbackError) { throw new Error('名簿更新と復旧に失敗しました。シートを確認してください。' + error + ' / ' + rollbackError); }
    throw error;
  }
}

function studentdataMulti() {
  const message = withAppLock_(studentdataMultiUnlocked_);
  safeAlert_(message);
}

function studentdataMultiUnlocked_() {
  const ss = getAppSpreadsheet_();
  const classSheet = ss.getSheetByName('クラス一覧');
  if (!classSheet || classSheet.getLastRow() < 2) return 'クラス一覧が空です。既存の生徒一覧は保持しています。';
  if (!hasMatchingHeaders_(classSheet, ['クラス名', 'コースID', '同期対象(1)'])) throw new Error('クラス一覧の構成が一致しません。');
  const selected = classSheet.getRange(2, 1, classSheet.getLastRow() - 1, 3).getValues().filter(function(row) {
    return String(row[1] || '').trim() && String(row[2] || '').trim() === '1';
  });
  if (!selected.length) return '同期対象クラスがありません。クラス一覧の同期対象に1を入力してください。既存の生徒一覧は保持しています。';
  const studentSheet = ss.getSheetByName(STUDENT_SHEET_NAME);
  if (studentSheet && studentSheet.getLastRow() && !hasMatchingHeaders_(studentSheet, STUDENT_SHEET_HEADERS)) throw new Error('生徒一覧の構成が一致しません。既存データと見出しを確認してください。');
  const rows = [];
  const courseIds = [];
  selected.forEach(function(item) {
    const courseId = String(item[1]).trim();
    if (courseIds.indexOf(courseId) >= 0) return;
    courseIds.push(courseId);
    getStudentListMax(courseId).forEach(function(student) {
      rows.push([rows.length + 1, student.email, student.name, item[0] || '', courseId, student.studentId]);
    });
  });
  replaceRosterRows_(STUDENT_SHEET_NAME, STUDENT_SHEET_HEADERS, rows);
  PropertiesService.getScriptProperties().setProperty('TURRET_ROSTER_SELECTION', JSON.stringify(courseIds.sort()));
  return '生徒一覧を更新しました（' + rows.length + '名）。';
}

function getStudentListMax(classId) {
  let pageToken = '';
  const all = [];

  do {
    const response = Classroom.Courses.Students.list(classId, { pageToken: pageToken });
    if (!response || typeof response !== 'object' || (response.students != null && !Array.isArray(response.students))) throw new Error('生徒一覧の取得結果が不正です。');
    const list = response.students || [];
    for (let i = 0; i < list.length; i++) {
      const student = list[i];
      all.push({
        email: String(student.profile && student.profile.emailAddress || '').trim(),
        name: String(student.profile && student.profile.name && student.profile.name.fullName || '').trim(),
        studentId: String(student.profile && student.profile.id || '').trim()
      });
    }
    pageToken = response ? response.nextPageToken : '';
  } while (pageToken);

  return all;
}

function importFromFormsToEval() {
  return withAppLock_(importFromFormsToEvalUnlocked_);
}

function importFromFormsToEvalUnlocked_() {
  if(typeof assertNoManagedFormUpdate_==='function')assertNoManagedFormUpdate_();
  if(typeof refreshAllManagedNames_==='function')refreshAllManagedNames_();
  const config = getConfig_();
  validateAppConfig_(config);
  const formStatusCache = createFormStatusCache_();
  const evalSheet = ensureEvalSheet_(config);
  const evalHeaders = evalSheet.getRange(1, 1, 1, evalSheet.getLastColumn()).getDisplayValues()[0].map(String);
  const evalHeaderMap = createHeaderMap_(evalHeaders);
  const existingResponses = collectResponseKeys_(evalSheet.getDataRange().getDisplayValues(), evalHeaderMap);

  const scoreField = resolveScoreField_(config);
  if (!scoreField) {
    throw new Error('scoreSourceHeader に対応する項目が見つかりません。');
  }

  config.formSources.forEach(function(source) {
    if (!formStatusCache.spreadsheets.has(source.id)) formStatusCache.spreadsheets.set(source.id, SpreadsheetApp.openById(source.id));
    const formSs = formStatusCache.spreadsheets.get(source.id);
    formSs.getSheets().forEach(function(sheet) {
      const sheetName = sheet.getName();
      if (isManagedInternalSheet_(sheetName, source.id) || !sheetMatchesPrefixes_(sheetName, config.formSheetNamePrefix)) return;

      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow <= 1 || lastCol === 0) return;

      const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
      const displayValues = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
      const headers = values[0].map(String);
      const headerMap = createHeaderMap_(headers);
      formStatusCache.targets.set(JSON.stringify([source.id, sheetName, config.formStatusHeader]), {
        sheet: sheet, statusColumn: headers.indexOf(config.formStatusHeader) + 1
      });

      const requiredHeaders = [config.emailHeader, config.studentNameHeader, config.formStatusHeader];
      config.fields.forEach(function(field) {
        requiredHeaders.push(field.sourceHeader);
      });

      const missingHeaders = requiredHeaders.filter(function(header) {
        return !(header in headerMap);
      });
      if (missingHeaders.length > 0) {
        ensureErrorSheet_().appendRow([
          new Date(),
          sheetName,
          '-',
          '',
          '',
          'CONFIG',
          '不足ヘッダ: ' + missingHeaders.join(', ')
        ]);
        return;
      }

      const rowsToAppend = [];
      const statusUpdates = [];

      for (let r = 1; r < values.length; r++) {
        const row = values[r];
        const dispRow = displayValues[r];
        const email = String(row[headerMap[config.emailHeader]] || '').trim();
        const studentName = String(row[headerMap[config.studentNameHeader]] || '').trim();
        const status = String(row[headerMap[config.formStatusHeader]] || '').trim();
        const scoreRaw = row[headerMap[scoreField.sourceHeader]];

        if (!email) continue;
        if (status) continue;
        if (String(scoreRaw == null ? '' : scoreRaw).trim() === '') continue;
        const responseKey = makeResponseKey_(source.id, sheetName, r + 1);
        if (existingResponses.has(responseKey)) continue;

        const newRow = new Array(evalHeaders.length).fill('');
        newRow[evalHeaderMap['元SS_ID']] = source.id;
        newRow[evalHeaderMap['元シート名']] = sheetName;
        newRow[evalHeaderMap['元行番号']] = r + 1;
        newRow[evalHeaderMap['メールアドレス']] = email;
        newRow[evalHeaderMap['名前']] = studentName;
        newRow[evalHeaderMap['処理状態']] = '';

        config.fields.forEach(function(field) {
          const headerIndex = headerMap[field.sourceHeader];
          const processed = processConfiguredFieldValue_(field, row[headerIndex], dispRow[headerIndex], config);
          newRow[evalHeaderMap[getFieldEvalHeader_(field)]] = processed;
        });

        rowsToAppend.push(newRow);
        existingResponses.add(responseKey);
        statusUpdates.push({ rowIndex: r + 1, status: '反映〇' });
      }

      appendRows_(evalSheet, rowsToAppend);
      if (rowsToAppend.length > 0) SpreadsheetApp.flush();
      statusUpdates.forEach(function(update) {
        updateFormStatus_(config, source.id, sheetName, update.rowIndex, update.status, formStatusCache);
      });
    });
  });
}

function evalToSendSheet() {
  return withAppLock_(evalToSendSheetUnlocked_);
}

function evalToSendSheetUnlocked_() {
  if(typeof assertNoManagedFormUpdate_==='function')assertNoManagedFormUpdate_();
  const config = getConfig_();
  validateAppConfig_(config);
  const formStatusCache = createFormStatusCache_();

  const evalSheet = ensureEvalSheet_(config);
  const sendSheet = ensureSendSheet_(config);
  const studentSheet = getStudentSheet_();
  const mappingEntries = readMappingEntries_(true);

  const studentValues = studentSheet.getDataRange().getDisplayValues();
  if (studentValues.length <= 1) {
    throw new Error('生徒一覧にデータがありません。');
  }

  const studentHeaders = studentValues[0].map(String);
  const studentHeaderMap = createHeaderMap_(studentHeaders);
  ['メールアドレス', '名前', 'クラス名', 'コースID', 'studentId'].forEach(function(header) {
    if (!(header in studentHeaderMap)) {
      throw new Error('生徒一覧のヘッダに「' + header + '」が必要です。');
    }
  });

  const studentMap = new Map();
  const studentsByEmail = new Map();
  for (let i = 1; i < studentValues.length; i++) {
    const row = studentValues[i];
    const email = String(row[studentHeaderMap['メールアドレス']] || '').trim();
    const courseId = String(row[studentHeaderMap['コースID']] || '').trim();
    const lookupKey = makeStudentLookupKey_(courseId, email);
    if (!email || !courseId) continue;
    const record = {
      name: String(row[studentHeaderMap['名前']] || '').trim(),
      className: String(row[studentHeaderMap['クラス名']] || '').trim(),
      courseId: courseId,
      studentId: String(row[studentHeaderMap['studentId']] || '').trim()
    };
    studentMap.set(lookupKey, record);

    const emailKey = String(email || '').trim().toLowerCase();
    const bucket = studentsByEmail.get(emailKey) || [];
    bucket.push(record);
    studentsByEmail.set(emailKey, bucket);
  }

  const evalValues = evalSheet.getDataRange().getDisplayValues();
  if (evalValues.length <= 1) return;

  const evalHeaders = evalValues[0].map(String);
  const evalHeaderMap = createHeaderMap_(evalHeaders);
  const sendHeaders = sendSheet.getRange(1, 1, 1, sendSheet.getLastColumn()).getDisplayValues()[0].map(String);
  const sendHeaderMap = createHeaderMap_(sendHeaders);
  const existingResponses = collectResponseKeys_(sendSheet.getDataRange().getDisplayValues(), sendHeaderMap);

  let nextNo = 1;
  if (sendSheet.getLastRow() > 1 && sendHeaderMap.No != null) {
    const lastNo = Number(sendSheet.getRange(sendSheet.getLastRow(), sendHeaderMap.No + 1).getValue() || 0);
    if (!isNaN(lastNo) && lastNo > 0) nextNo = lastNo + 1;
  }

  const sendRows = [];
  const stateUpdates = [];
  const errorSheet = ensureErrorSheet_();

  for (let r = 1; r < evalValues.length; r++) {
    const row = evalValues[r];
    if (row.every(function(value) { return String(value == null ? '' : value).trim() === ''; })) continue;
    const state = String(row[evalHeaderMap['処理状態']] || '').trim();
    if (state) continue;

    const srcSsId = String(row[evalHeaderMap['元SS_ID']] || '').trim();
    const srcSheetName = String(row[evalHeaderMap['元シート名']] || '').trim();
    const srcRow = Number(row[evalHeaderMap['元行番号']] || 0);
    const email = String(row[evalHeaderMap['メールアドレス']] || '').trim();
    const evalName = String(row[evalHeaderMap['名前']] || '').trim();
    const responseKey = makeResponseKey_(srcSsId, srcSheetName, srcRow);
    if (!responseKey) {
      throw new Error('評価データの元回答情報が不正です。行 ' + (r + 1) + ' の元SS_ID・元シート名・元行番号を確認してください。');
    }
    if (existingResponses.has(responseKey)) {
      // 既存の送信行は保持し、元フォームの最終状態（済など）を巻き戻さない。
      stateUpdates.push({ rowNum: r + 1, evalState: '準備〇' });
      continue;
    }

    if (!email) {
      errorSheet.appendRow([new Date(), EVAL_SHEET_NAME, r + 1, '', '', 'CONFIG', 'メールアドレスが空のため本文を準備できません。評価データと元回答を確認してください。']);
      stateUpdates.push({ rowNum: r + 1, evalState: '準備×', srcState: '準備×', srcSsId: srcSsId, srcSheetName: srcSheetName, srcRow: srcRow });
      continue;
    }

    const mapping = mappingEntries.get(makeMappingLookupKey_(srcSsId, srcSheetName));
    const mappedCourseId = String(mapping && mapping.courseId || '').trim();
    let student = null;

    if (mappedCourseId) {
      student = studentMap.get(makeStudentLookupKey_(mappedCourseId, email)) || null;
      if (!student) {
        errorSheet.appendRow([
          new Date(),
          SEND_SHEET_NAME,
          '-',
          email,
          '',
          'CONFIG',
          '対応表の courseId と生徒一覧が一致しません: ' + srcSheetName + ' / ' + mappedCourseId
        ]);
        stateUpdates.push({ rowNum: r + 1, evalState: '準備×', srcState: '準備×', srcSsId: srcSsId, srcSheetName: srcSheetName, srcRow: srcRow });
        continue;
      }
    }

    const candidates = studentsByEmail.get(String(email).trim().toLowerCase()) || [];
    if (!student && candidates.length === 1) {
      student = candidates[0];
    }

    if (!student && candidates.length > 1) {
      errorSheet.appendRow([
        new Date(),
        SEND_SHEET_NAME,
        '-',
        email,
        '',
        'CONFIG',
        '同じメールアドレスが複数クラスに存在するため送信先クラスを特定できません。対応表シートで元シート名にクラスを割り当ててください'
      ]);
      stateUpdates.push({ rowNum: r + 1, evalState: '準備×', srcState: '準備×', srcSsId: srcSsId, srcSheetName: srcSheetName, srcRow: srcRow });
      continue;
    }

    if (!student) {
      errorSheet.appendRow([
        new Date(),
        SEND_SHEET_NAME,
        '-',
        email,
        '',
        'FATAL',
        '生徒一覧にメールアドレスが存在しません'
      ]);
      stateUpdates.push({ rowNum: r + 1, evalState: '準備×', srcState: '準備×', srcSsId: srcSsId, srcSheetName: srcSheetName, srcRow: srcRow });
      continue;
    }

    if (!student.courseId) {
      errorSheet.appendRow([
        new Date(),
        SEND_SHEET_NAME,
        '-',
        email,
        '',
        'FATAL',
        '生徒一覧にコースIDがありません'
      ]);
      stateUpdates.push({ rowNum: r + 1, evalState: '準備×', srcState: '準備×', srcSsId: srcSsId, srcSheetName: srcSheetName, srcRow: srcRow });
      continue;
    }

    if (!student.studentId) {
      errorSheet.appendRow([
        new Date(),
        SEND_SHEET_NAME,
        '-',
        email,
        '',
        'FATAL',
        '生徒一覧に studentId がありません。生徒一覧取得をやり直してください'
      ]);
      stateUpdates.push({ rowNum: r + 1, evalState: '準備×', srcState: '準備×', srcSsId: srcSsId, srcSheetName: srcSheetName, srcRow: srcRow });
      continue;
    }

    const rowNo = nextNo++;
    const context = {
      student_name: student.name || evalName,
      class_name: student.className,
      course_id: student.courseId,
      row_no: String(rowNo)
    };

    config.fields.forEach(function(field) {
      context[field.key] = String(row[evalHeaderMap[getFieldEvalHeader_(field)]] || '').trim();
    });

    const replyBody = renderTemplateWithSections_(config.messageTemplate, context, config);
    if (!replyBody) {
      errorSheet.appendRow([
        new Date(),
        SEND_SHEET_NAME,
        '-',
        email,
        '',
        'CONFIG',
        '返信テンプレートの結果が空になりました'
      ]);
      stateUpdates.push({ rowNum: r + 1, evalState: '準備×', srcState: '準備×', srcSsId: srcSsId, srcSheetName: srcSheetName, srcRow: srcRow });
      continue;
    }

    const sendRow = new Array(sendHeaders.length).fill('');
    sendRow[sendHeaderMap['元SS_ID']] = srcSsId;
    sendRow[sendHeaderMap['元シート名']] = srcSheetName;
    sendRow[sendHeaderMap['元行番号']] = srcRow;
    sendRow[sendHeaderMap['No']] = rowNo;
    sendRow[sendHeaderMap['メールアドレス']] = email;
    sendRow[sendHeaderMap['名前']] = student.name || evalName;
    sendRow[sendHeaderMap['クラス名']] = student.className;
    sendRow[sendHeaderMap['コースID']] = student.courseId;
    sendRow[sendHeaderMap['studentId']] = student.studentId;
    sendRow[sendHeaderMap['送信状態']] = '';
    sendRow[sendHeaderMap[config.replyBodyHeader]] = replyBody;

    config.fields.forEach(function(field) {
      const sendHeader = getFieldSendHeader_(field);
      if (sendHeader && sendHeaderMap[sendHeader] != null) {
        sendRow[sendHeaderMap[sendHeader]] = context[field.key];
      }
    });

    sendRows.push(sendRow);
    existingResponses.add(responseKey);
    stateUpdates.push({ rowNum: r + 1, evalState: '準備〇', srcState: '準備〇', srcSsId: srcSsId, srcSheetName: srcSheetName, srcRow: srcRow });
  }

  appendRows_(sendSheet, sendRows);
  if (sendRows.length > 0) SpreadsheetApp.flush();

  stateUpdates.forEach(function(update) {
    evalSheet.getRange(update.rowNum, evalHeaderMap['処理状態'] + 1).setValue(update.evalState);
    if (update.srcSsId && update.srcSheetName && update.srcRow) {
      updateFormStatus_(config, update.srcSsId, update.srcSheetName, update.srcRow, update.srcState, formStatusCache);
    }
  });
}

function postAnnouncementIndividual_(courseId, studentId, name, text) {
  const data = {
    courseId: String(courseId),
    text: name + 'さんへ(個別メッセージ)\n\n' + text,
    assigneeMode: 'INDIVIDUAL_STUDENTS',
    individualStudentsOptions: { studentIds: [String(studentId)] },
    state: 'PUBLISHED'
  };

  const maxRetry = 3;
  for (let t = 0; t < maxRetry; t++) {
    try {
      return Classroom.Courses.Announcements.create(data, String(courseId));
    } catch (e) {
      // Retry only an explicit rate-limit rejection. Internal errors/timeouts may
      // have happened after Classroom accepted the announcement.
      if (t < maxRetry - 1 && isClassroomRateLimitError_(e)) {
        Utilities.sleep(500 * (t + 1));
        continue;
      }
      throw e;
    }
  }
}

function isClassroomRateLimitError_(error) {
  return /\bHTTP\s+429\b|\bRESOURCE_EXHAUSTED\b|\brateLimitExceeded\b|\buserRateLimitExceeded\b/.test(String(error));
}

function isClassroomRejectedError_(error) {
  return /\bHTTP\s+(?:400|401|403|404)\b|\b(?:INVALID_ARGUMENT|UNAUTHENTICATED|PERMISSION_DENIED|NOT_FOUND)\b/.test(String(error));
}

function sendMessages() {
  const message = withAppLock_(sendMessagesUnlocked_);
  safeAlert_(message);
}

function sendMessagesUnlocked_() {
  if(typeof assertNoManagedFormUpdate_==='function')assertNoManagedFormUpdate_();
  const config = getConfig_();
  validateAppConfig_(config);
  const formStatusCache = createFormStatusCache_();
  const sendSheet = ensureSendSheet_(config);
  const lastRow = sendSheet.getLastRow();
  const lastCol = sendSheet.getLastColumn();
  if (lastRow <= 1) {
    return '送信シートにデータがありません。';
  }

  const values = sendSheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const headers = values[0].map(String);
  const headerMap = createHeaderMap_(headers);
  const requiredHeaders = ['元SS_ID', '元シート名', '元行番号', 'メールアドレス', '名前', 'コースID', 'studentId', '送信状態', config.replyBodyHeader];

  requiredHeaders.forEach(function(header) {
    if (!(header in headerMap)) {
      throw new Error('送信シートのヘッダに「' + header + '」が必要です。');
    }
  });

  const errorSheet = ensureErrorSheet_();
  let sentCount = 0;
  let errorCount = 0;
  let reviewCount = 0;
  const postRegistry = typeof getManagedRecords_ === 'function' ? {records:getManagedRecords_(),sheet:null} : null;
  const responseCounts = new Map();
  values.slice(1).forEach(function(row) {
    const key = makeResponseKey_(row[headerMap['元SS_ID']], row[headerMap['元シート名']], row[headerMap['元行番号']]);
    if (key) responseCounts.set(key, (responseCounts.get(key) || 0) + 1);
  });

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const status = String(row[headerMap['送信状態']] || '').trim();
    if (row.every(function(value) { return String(value == null ? '' : value).trim() === ''; })) continue;
    if (status === SEND_REVIEW_STATUS) {
      reviewCount++;
      continue;
    }
    if (status && status !== '未') continue;

    const srcSsId = String(row[headerMap['元SS_ID']] || '').trim();
    const srcSheetName = String(row[headerMap['元シート名']] || '').trim();
    const srcRow = Number(row[headerMap['元行番号']] || 0);
    const email = String(row[headerMap['メールアドレス']] || '').trim();
    const name = String(row[headerMap['名前']] || '').trim();
    const courseId = String(row[headerMap['コースID']] || '').trim();
    const studentId = String(row[headerMap['studentId']] || '').trim();
    const body = String(row[headerMap[config.replyBodyHeader]] || '').trim();
    const responseKey = makeResponseKey_(srcSsId, srcSheetName, srcRow);
    const statusCell = sendSheet.getRange(r + 1, headerMap['送信状態'] + 1);

    const missingFields = [];
    if (!responseKey) missingFields.push('有効な元SS_ID・元シート名・元行番号（2以上の整数）');
    if (!email) missingFields.push('メールアドレス');
    if (!name) missingFields.push('名前');
    if (!courseId) missingFields.push('コースID');
    if (!studentId) missingFields.push('studentId');
    if (!body) missingFields.push(config.replyBodyHeader);
    if (missingFields.length > 0) {
      sendSheet.getRange(r + 1, headerMap['送信状態'] + 1).setValue('エラー');
      errorSheet.appendRow([
        new Date(),
        SEND_SHEET_NAME,
        r + 1,
        email,
        'エラー',
        'CONFIG',
        '送信に必要な値が不足しています: ' + missingFields.join(', ')
      ]);
      if (responseKey) {
        updateFormStatus_(config, srcSsId, srcSheetName, srcRow, '送信エラー', formStatusCache);
      }
      errorCount++;
      continue;
    }

    if (responseCounts.get(responseKey) > 1) {
      statusCell.setValue(SEND_REVIEW_STATUS);
      errorSheet.appendRow([new Date(), SEND_SHEET_NAME, r + 1, email,
        SEND_REVIEW_STATUS, 'DUPLICATE_SOURCE',
        '同じ元回答の送信行が複数あります。Classroomの投稿と各行を確認してください。']);
      reviewCount++;
      continue;
    }

    // Persist the checkpoint before crossing the external API boundary. If the
    // execution stops here or the outcome is unknown, a later run must not post.
    statusCell.setValue(SEND_REVIEW_STATUS);
    SpreadsheetApp.flush();
    let announcement;
    try {
      announcement = postAnnouncementIndividual_(courseId, studentId, name, body);
    } catch (e) {
      const msg = String(e);
      let errType = 'SEND_RESULT_UNKNOWN';
      let newStatus = SEND_REVIEW_STATUS;
      if (isClassroomRateLimitError_(e)) {
        errType = 'RETRYABLE';
        newStatus = '未';
      } else if (isClassroomRejectedError_(e)) {
        errType = 'FATAL';
        newStatus = 'エラー';
      }
      if (newStatus !== SEND_REVIEW_STATUS) statusCell.setValue(newStatus);
      errorSheet.appendRow([
        new Date(),
        SEND_SHEET_NAME,
        r + 1,
        email,
        newStatus,
        errType,
        msg
      ]);
      if (newStatus === SEND_REVIEW_STATUS) {
        reviewCount++;
      } else {
        errorCount++;
        if (newStatus === 'エラー') updateFormStatus_(config, srcSsId, srcSheetName, srcRow, '送信エラー', formStatusCache);
      }
      continue;
    }

    // A successful post and a failed Sheets write are different outcomes.
    // Never roll a successfully posted row back to an automatically sent status.
    try {
      if (postRegistry) recordManagedAnnouncement_(announcement, {courseId:courseId,name:name,sourceKey:responseKey},postRegistry);
      statusCell.setValue('済');
      SpreadsheetApp.flush();
    } catch (e) {
      errorSheet.appendRow([new Date(), SEND_SHEET_NAME, r + 1, email,
        SEND_REVIEW_STATUS, 'SEND_RECORD_ERROR',
        'Classroom投稿は成功。投稿ID: ' + String(announcement && announcement.id || '不明') +
        '。送信状態の保存を確認してください。' + String(e)]);
      reviewCount++;
      continue;
    }
    sentCount++;
    updateFormStatus_(config, srcSsId, srcSheetName, srcRow, '済', formStatusCache);
    Utilities.sleep(120);
  }

  return '送信処理が終了しました。\n' +
    '送信成功：' + sentCount + '件\n' +
    'エラー　：' + errorCount + '件（エラーシート参照）\n' +
    '要確認　：' + reviewCount + '件（送信確認待ち・エラーシート参照）';
}

function remindUngradedAndErrors() {
  return withAppLock_(remindUngradedAndErrorsUnlocked_);
}

function remindUngradedAndErrorsUnlocked_() {
  if(typeof assertNoManagedFormUpdate_==='function')assertNoManagedFormUpdate_();
  const config = getConfig_();
  validateAppConfig_(config);

  if (!Array.isArray(config.reminderTo) || config.reminderTo.length === 0) {
    throw new Error('リマインドメール送信先が設定されていません。管理画面の「フォームと通知先」で通知先を入力してください。');
  }

  const submissionSourceHeader = config.emailHeader;
  const scoreField = resolveScoreField_(config);

  if (!submissionSourceHeader || !scoreField) {
    throw new Error('リマインド判定に必要な項目設定が不足しています。');
  }

  const lines = [];
  let totalUngraded = 0;
  let totalPending = 0;
  let totalError = 0;

  config.formSources.forEach(function(source) {
    const formSs = SpreadsheetApp.openById(source.id);
    const fileName = formSs.getName();

    formSs.getSheets().forEach(function(sheet) {
      const sheetName = sheet.getName();
      if (isManagedInternalSheet_(sheetName, source.id) || !sheetMatchesPrefixes_(sheetName, config.formSheetNamePrefix)) return;

      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow <= 1 || lastCol === 0) return;

      const values = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
      const headers = values[0].map(String);
      const headerMap = createHeaderMap_(headers);

      if (!(submissionSourceHeader in headerMap) || !(scoreField.sourceHeader in headerMap)) {
        return;
      }

      const idxSubmission = headerMap[submissionSourceHeader];
      const idxScore = headerMap[scoreField.sourceHeader];
      const idxStatus = headerMap[config.formStatusHeader];

      let ungraded = 0;
      let pending = 0;
      let errorCnt = 0;

      for (let r = 1; r < values.length; r++) {
        const row = values[r];
        const submission = String(row[idxSubmission] || '').trim();
        const score = String(row[idxScore] || '').trim();
        const status = idxStatus != null ? String(row[idxStatus] || '').trim() : '';

        if (submission && !score) ungraded++;
        if (status === '準備〇') pending++;
        if (status === '準備×' || status === '送信エラー') errorCnt++;
      }

      if (ungraded > 0 || pending > 0 || errorCnt > 0) {
        lines.push(
          fileName + ' / ' + sheetName +
          '：未採点 ' + ungraded + '件, ' +
          '未送信 ' + pending + '件, ' +
          'エラー ' + errorCnt + '件'
        );
        totalUngraded += ungraded;
        totalPending += pending;
        totalError += errorCnt;
      }
    });
  });

  if (totalUngraded === 0 && totalPending === 0 && totalError === 0) {
    return;
  }

  let body = '';
  body += '【フォーム別の状況】\n';
  lines.forEach(function(line) {
    body += '・' + line + '\n';
  });
  body += '\n';
  body += '【全体集計】\n';
  body += '・未採点　：' + totalUngraded + '件\n';
  body += '・未送信　：' + totalPending + '件（送信シートは作成済み）\n';
  body += '・エラー　：' + totalError + '件（準備×／送信エラー）\n\n';
  body += '（このメールはGASの時間トリガーで自動送信されています）\n';

  GmailApp.sendEmail(
    (config.reminderTo || []).join(','),
    '授業まとめ／未採点・送信状況のリマインド（Form基準）',
    body
  );
}

function clearEvaluationData() {
  const ss = getAppSpreadsheet_();
  const sheet = ss.getSheetByName(EVAL_SHEET_NAME);
  if (!sheet) {
    safeAlert_('「評価データ」シートが存在しません。');
    return;
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1) {
    safeAlert_('「評価データ」にはクリアする行がありません。');
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    '確認',
    '「評価データ」シートの 2 行目以降をすべて削除します。\nこの操作は取り消しできません。\n実行しますか？',
    ui.ButtonSet.OK_CANCEL
  );
  if (result !== ui.Button.OK) return;

  withAppLock_(function() {
    const current = ss.getSheetByName(EVAL_SHEET_NAME);
    if (current && current.getLastRow() > 1) {
      current.getRange(2, 1, current.getLastRow() - 1, current.getLastColumn()).clearContent();
    }
  });
  safeAlert_('「評価データ」シートをクリアしました。');
}

function clearSendSheet() {
  const ss = getAppSpreadsheet_();
  const sheet = ss.getSheetByName(SEND_SHEET_NAME);
  if (!sheet) {
    safeAlert_('「送信シート」が存在しません。');
    return;
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1) {
    safeAlert_('「送信シート」にはクリアする行がありません。');
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    '確認',
    '「送信シート」の 2 行目以降をすべて削除します。\nこの操作は取り消せません。\nよろしいですか？',
    ui.ButtonSet.OK_CANCEL
  );
  if (result !== ui.Button.OK) return;

  withAppLock_(function() {
    const current = ss.getSheetByName(SEND_SHEET_NAME);
    if (current && current.getLastRow() > 1) {
      current.getRange(2, 1, current.getLastRow() - 1, current.getLastColumn()).clearContent();
    }
  });
  safeAlert_('「送信シート」をクリアしました。');
}

function clearErrorLog() {
  const ss = getAppSpreadsheet_();
  const sheet = ss.getSheetByName(ERROR_SHEET_NAME);
  if (!sheet) {
    safeAlert_('「エラー」シートは存在しません。');
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    '確認',
    '「エラー」シートを削除します。\nこの操作は取り消せません。\n実行しますか？',
    ui.ButtonSet.OK_CANCEL
  );
  if (result !== ui.Button.OK) return;

  withAppLock_(function() {
    const current = ss.getSheetByName(ERROR_SHEET_NAME);
    if (current) ss.deleteSheet(current);
  });
  safeAlert_('「エラー」シートを削除しました。');
}
