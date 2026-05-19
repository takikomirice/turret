/******************************************************
 * Google Classroom 評価自動送信 GAS
 *
 * - 設定シート（手動編集可）: reminderTo・formSources・prefix の権威的ソース
 * - Config シート（GAS 管理）: フィールド・列名・テンプレートを JSON で保持
 * - ScriptProperties: ランタイムキャッシュ（設定シート変更は即時反映）
 * - 学期コピー時は設定シート・Config シートがそのまま引き継がれる
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

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('自動送信システム')
    .addItem('設定を開く', 'openSettingsDialog')
    .addSeparator()
    .addItem('1 初期設定シート作成', 'initializeSheets')
    .addSeparator()
    .addItem('2 クラス一覧取得', 'classroomdata')
    .addItem('3 生徒一覧取得（同期対象クラス）', 'studentdataMulti')
    .addItem('4 対応表を作成', 'createMappingSheet')
    .addSeparator()
    .addItem('5 Form → 評価データ取り込み', 'importFromFormsToEval')
    .addItem('6 評価データ → 送信シート生成', 'evalToSendSheet')
    .addSeparator()
    .addItem('7 生徒へメッセージ送信（Classroom）', 'sendMessages')
    .addSeparator()
    .addItem('8 未採点・エラーのリマインダー送信', 'remindUngradedAndErrors')
    .addSeparator()
    .addItem('9 評価データをクリア', 'clearEvaluationData')
    .addItem('10 送信シートをクリア', 'clearSendSheet')
    .addItem('11 エラーログのクリア', 'clearErrorLog')
    .addToUi();
}

function initializeSheets() {
  ensureSheet_('クラス一覧', ['クラス名', 'コースID', '同期対象(1)']);
  ensureStudentSheetForSync_();
  ensureSettingsInputSheet_();
  safeAlert_('シートを作成しました（クラス一覧・生徒一覧・設定シート）');
}

function openSettingsDialog() {
  const html = HtmlService.createHtmlOutputFromFile('SettingsDialog')
    .setTitle('自動送信設定')
    .setWidth(900)
    .setHeight(760);
  SpreadsheetApp.getUi().showModalDialog(html, '自動送信設定');
}

function getSettingsDialogData() {
  const config = getConfig_();
  const parseError = PropertiesService.getScriptProperties()
    .getProperty('APP_CONFIG_PARSE_ERROR') === '1';

  return {
    config: config,
    configParseError: parseError,
    availableFieldTypes: [
      { value: 'text', label: 'text' },
      { value: 'date', label: 'date' },
      { value: 'score_grade', label: 'score_grade' }
    ],
    templateGuide: [
      '簡易記法は ＜score＞ です。例: 評価：＜score＞',
      '固定トークンも同じ書き方です。例: ＜student_name＞, ＜class_name＞, ＜course_id＞, ＜row_no＞',
      '値だけの行は、値が空なら自動で非表示になります。',
      '括弧や見出しを付けた行も、対応する値が空ならその行は自動で消えます。'
    ].join('\n')
  };
}

function saveSettingsFromDialog(payload, confirmedReset) {
  // 設定ダイアログは reminderTo / formSources / formSheetNamePrefix を持たないので、
  // 既存の設定からそれらを引き継いでマージする。
  const existing = getConfig_();
  const merged = Object.assign({}, existing, payload);
  const config = normalizeAppConfig_(merged);
  validateAppConfig_(config);

  const resetInfo = getResetImpactSummary_();
  const needsConfirmation = resetInfo.evalRows > 0 || resetInfo.sendRows > 0;
  if (!confirmedReset && needsConfirmation) {
    return {
      ok: false,
      requiresConfirmation: true,
      summary: buildResetSummaryText_(resetInfo)
    };
  }

  saveConfig_(config);
  writeDisplaySheet_(config);
  resetConfiguredSheets_(config);

  return {
    ok: true,
    message: '設定を保存し、シートを再生成しました。'
  };
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
        return sheetMatchesPrefixes_(sheet.getName(), normalizedPrefixes);
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

/**
 * 設定シートを作成する。既に存在する場合は何もしない（内容を保護する）。
 * REMINDER_TO・FORM_SHEET_PREFIX・FORM_SS を直接編集して使う。
 */
function ensureSettingsInputSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(INPUT_SETTINGS_SHEET_NAME)) return;

  const sheet = ss.insertSheet(INPUT_SETTINGS_SHEET_NAME);
  const rows = [
    ['KEY', 'VALUE', 'NOTE'],
    ['REMINDER_TO', '', 'リマインドメール送信先（行を追加して複数入力可）'],
    ['FORM_SHEET_PREFIX', 'フォームの回答', '対象シート名の接頭辞（行を追加して複数入力可）'],
    ['FORM_SS', '', 'フォーム回答スプレッドシートのURL（行を追加して複数入力可）']
  ];
  sheet.getRange(1, 1, rows.length, 3).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 3);
}

function ensureSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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
 * prefixes が空または全て空文字の場合はすべてのシートを許可（フィルタなし）。
 * @param {string} sheetName
 * @param {string[]} prefixes
 * @returns {boolean}
 */
function sheetMatchesPrefixes_(sheetName, prefixes) {
  const active = (prefixes || []).filter(function(p) { return !!p; });
  if (active.length === 0) return true;
  return active.some(function(p) { return sheetName.indexOf(p) === 0; });
}

function buildDefaultConfig_() {
  return {
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

  return normalized;
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
  const normalizedFields = fields.length > 0
    ? fields.map(function(field, index) { return normalizeFieldConfig_(field, index); }).filter(function(field) {
        return field.sourceHeader;
      })
    : defaults.fields;
  let scoreSourceHeader = String(config.scoreSourceHeader != null ? config.scoreSourceHeader : '').trim();

  if (!scoreSourceHeader && config.scoreFieldKey != null) {
    const legacyField = getFieldByKey_({ fields: normalizedFields }, String(config.scoreFieldKey).trim());
    if (legacyField) scoreSourceHeader = legacyField.sourceHeader;
  }
  if (!scoreSourceHeader) {
    scoreSourceHeader = defaults.scoreSourceHeader;
  }

  return {
    reminderTo: normalizeStringArray_(config.reminderTo != null ? config.reminderTo : defaults.reminderTo),
    formSheetNamePrefix: normalizeStringArray_(config.formSheetNamePrefix != null ? config.formSheetNamePrefix : defaults.formSheetNamePrefix, defaults.formSheetNamePrefix),
    emailHeader: String(config.emailHeader != null ? config.emailHeader : defaults.emailHeader).trim() || defaults.emailHeader,
    studentNameHeader: String(config.studentNameHeader != null ? config.studentNameHeader : defaults.studentNameHeader).trim() || defaults.studentNameHeader,
    formStatusHeader: String(config.formStatusHeader != null ? config.formStatusHeader : defaults.formStatusHeader).trim() || defaults.formStatusHeader,
    scoreSourceHeader: scoreSourceHeader,
    replyBodyHeader: String(config.replyBodyHeader != null ? config.replyBodyHeader : defaults.replyBodyHeader).trim() || defaults.replyBodyHeader,
    messageTemplate: normalizeMultiline_(config.messageTemplate != null ? config.messageTemplate : defaults.messageTemplate),
    formSources: formSources.map(normalizeFormSource_).filter(function(item) { return !!item; }),
    fields: normalizedFields
  };
}

function validateAppConfig_(config) {
  // reminderTo は任意。リマインダー送信時のみチェックする。
  if (config.formSources.length === 0) {
    throw new Error('フォーム回答スプレッドシートURLを1件以上入力してください。');
  }
  if (config.fields.length === 0) {
    throw new Error('取り込み項目を1件以上設定してください。');
  }
  if (!config.messageTemplate.trim()) {
    throw new Error('返信テンプレートを入力してください。');
  }

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

  const match = text.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (match) return match[1];

  if (/^[a-zA-Z0-9-_]{20,}$/.test(text)) {
    return text;
  }

  throw new Error('スプレッドシートURLまたはIDの形式が不正です: ' + trunc_(text, 80));
}

/**
 * 設定(編集不可)シートに人間可読な内容を書き出す（表示専用）。
 * GAS はこのシートを設定ストレージとして読まない。
 * シートが存在しない場合は新規作成する。
 */
function writeDisplaySheet_(config) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_SHEET_NAME);
  }
  sheet.clear();

  const rows = [
    ['【注意】このシートは自動生成です。直接編集しても反映されません。', '', ''],
    ['', '', ''],
    ['KEY', 'VALUE', 'NOTE'],
    ...(config.reminderTo && config.reminderTo.length > 0
      ? config.reminderTo.map(function(addr, idx) {
          return ['REMINDER_TO', addr, idx === 0 ? 'リマインドメール送信先（行を追加して複数入力可）' : ''];
        })
      : [['REMINDER_TO', '', 'リマインドメール送信先']]),
    ...(config.formSheetNamePrefix && config.formSheetNamePrefix.length > 0
      ? config.formSheetNamePrefix.map(function(prefix, idx) {
          return ['FORM_SHEET_PREFIX', prefix, idx === 0 ? '対象シート名の接頭辞（行を追加して複数入力可）' : ''];
        })
      : [['FORM_SHEET_PREFIX', '', '対象シート名の接頭辞']]),
    ['EMAIL_HEADER', config.emailHeader, 'フォーム側のメール列名'],
    ['STUDENT_NAME_HEADER', config.studentNameHeader, 'フォーム側の名前列名'],
    ['FORM_STATUS_HEADER', config.formStatusHeader, 'フォーム側の送信状態列名'],
    ['SCORE_SOURCE_HEADER', config.scoreSourceHeader, '採点済み判定に使うフォーム側ヘッダ'],
    ['REPLY_BODY_HEADER', config.replyBodyHeader, '送信シートで返信本文を保持する列名'],
    ['MESSAGE_TEMPLATE', config.messageTemplate, '返信テンプレート'],
    ['', '', ''],
    ['TYPE', 'URL', 'NOTE']
  ];

  config.formSources.forEach(function(source, index) {
    rows.push([
      'FORM_SS',
      normalizeFormSource_(source).url,
      'フォーム回答スプレッドシート ' + (index + 1)
    ]);
  });

  rows.push(['', '', '']);
  rows.push(['FIELD_KEY', 'SUMMARY', 'NOTE']);
  config.fields.forEach(function(field) {
    rows.push([
      field.key,
      'source=' + field.sourceHeader +
      ' / eval=' + (field.evalHeader || '(sourceと同じ)') +
      ' / send=' + (field.sendHeader || '(sourceと同じ)') +
      ' / type=' + field.type,
      ''
    ]);
  });

  const normalizedRows = rows.map(function(row) {
    const copy = row.slice(0, 3);
    while (copy.length < 3) copy.push('');
    return copy;
  });

  sheet.getRange(1, 1, normalizedRows.length, 3).setValues(normalizedRows);
  sheet.autoResizeColumns(1, 3);
  sheet.setFrozenRows(1);
}

/**
 * Config シートから CONFIG_JSON を読み込む。
 * @returns {Object|null} パース済みオブジェクト、または null
 */
function readConfigSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet || sheet.getLastRow() === 0) return null;
  const json = String(sheet.getRange(1, 1).getValue() || '').trim();
  if (!json) return null;
  try {
    return JSON.parse(json);
  } catch (e) {
    Logger.log('[readConfigSheet_] CONFIG_JSON のパース失敗: ' + e);
    return null;
  }
}

/**
 * Config シートに CONFIG_JSON を書き込む。
 * シートが存在しない場合は作成する。
 * @param {Object} config
 */
function writeConfigSheet_(config) {
  if (config == null) throw new Error('[writeConfigSheet_] config が null/undefined です');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG_SHEET_NAME);
  }
  sheet.getRange(1, 1).setValue(JSON.stringify(config));
}

/**
 * 設定シートから REMINDER_TO・FORM_SHEET_PREFIX・FORM_SS を読み込む。
 * 値が空のキーは返さない（PropertiesService の既存データを上書きしないため）。
 * @returns {Object} 見つかった値だけを含む部分的な config オブジェクト
 */
function readSettingsInputSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(INPUT_SETTINGS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return {};

  const values = sheet.getDataRange().getDisplayValues();
  const result = {};
  const reminderTos = [];
  const prefixes = [];
  const formSources = [];

  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0] || '').trim();
    const value = String(values[i][1] || '').trim();
    if (!key || !value) continue;

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

  if (reminderTos.length > 0) result.reminderTo = reminderTos;
  if (prefixes.length > 0) result.formSheetNamePrefix = prefixes;
  if (formSources.length > 0) result.formSources = formSources;

  return result;
}

/**
 * 設定シートの REMINDER_TO・FORM_SHEET_PREFIX・FORM_SS 行を更新する。
 * シートの他の行は一切変更しない。
 * 既存行が見つからない場合はシート末尾に追加する。
 * FORM_SS は件数が増減した場合に行を追加／削除する。
 * @param {Object} config
 */
/**
 * 設定シートの特定 KEY の行を newValues に合わせて更新・追加・削除する。
 * 常に最低 1 行を維持する（空値でもよい）。
 * @param {Sheet} sheet
 * @param {number[]} existingRows - 該当 KEY の既存行番号（1始まり）
 * @param {string[]} newValues - 書き込む値の配列
 * @param {string} key - KEY 列の値
 * @param {string} defaultNote - 1行目の NOTE 列テキスト
 */
function syncMultiRows_(sheet, existingRows, newValues, key, defaultNote) {
  const targets = newValues.length > 0 ? newValues : [''];

  if (existingRows.length === 0) {
    targets.forEach(function(val, idx) {
      sheet.appendRow([key, val, idx === 0 ? defaultNote : '']);
    });
    return;
  }

  for (let j = 0; j < targets.length && j < existingRows.length; j++) {
    sheet.getRange(existingRows[j], 2).setValue(targets[j]);
  }
  for (let k = existingRows.length - 1; k >= targets.length; k--) {
    sheet.deleteRow(existingRows[k]);
  }
  for (let m = existingRows.length; m < targets.length; m++) {
    sheet.appendRow([key, targets[m], '']);
  }
}

function writeSettingsInputSheet_(config) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(INPUT_SETTINGS_SHEET_NAME);
  if (!sheet) {
    ensureSettingsInputSheet_();
    sheet = ss.getSheetByName(INPUT_SETTINGS_SHEET_NAME);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow === 0) return;

  const values = sheet.getRange(1, 1, lastRow, 2).getValues();

  const reminderToRows = [];
  const prefixRows = [];
  const formSsRows = [];

  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0] || '').trim();
    if (key === 'REMINDER_TO') reminderToRows.push(i + 1);
    else if (key === 'FORM_SHEET_PREFIX') prefixRows.push(i + 1);
    else if (key === 'FORM_SS') formSsRows.push(i + 1);
  }

  try {
    syncMultiRows_(sheet, reminderToRows, config.reminderTo || [],
      'REMINDER_TO', 'リマインドメール送信先（必須・行を追加して複数入力可）');
    syncMultiRows_(sheet, prefixRows, config.formSheetNamePrefix || [],
      'FORM_SHEET_PREFIX', '対象シート名の接頭辞（行を追加して複数入力可）');

    // courseId は対応表シートで管理するため、設定シートには URL のみ書き込む
    const newSources = (config.formSources || [])
      .map(function(s) { return normalizeFormSource_(s); })
      .filter(function(s) { return !!s; });

    for (let j = 0; j < newSources.length && j < formSsRows.length; j++) {
      sheet.getRange(formSsRows[j], 2).setValue(newSources[j].url);
    }
    for (let k = formSsRows.length - 1; k >= newSources.length; k--) {
      sheet.deleteRow(formSsRows[k]);
    }
    for (let m = formSsRows.length; m < newSources.length; m++) {
      sheet.appendRow(['FORM_SS', newSources[m].url, '']);
    }
  } catch (e) {
    Logger.log('[writeSettingsInputSheet_] シート書き込みに失敗しました: ' + e);
    throw e;
  }
}

/**
 * 設定を読み込む。
 * 設定シート（REMINDER_TO/FORM_SHEET_PREFIX/FORM_SS）は常に読んで優先適用する。
 * 複雑な設定（フィールド定義等）は PropertiesService → Config シート → デフォルト値の順にフォールバック。
 */
function loadConfig_() {
  // 設定シートは常に読んで overrides として適用（直接編集が即時反映される）
  const sheetOverrides = readSettingsInputSheet_();

  const props = PropertiesService.getScriptProperties();
  const json = props.getProperty('APP_CONFIG');
  if (json) {
    try {
      const cached = JSON.parse(json);
      return normalizeAppConfig_(Object.assign({}, cached, sheetOverrides));
    } catch (e) {
      Logger.log('[loadConfig_] APP_CONFIG のパース失敗。Config シートにフォールバック: ' + e);
      props.setProperty('APP_CONFIG_PARSE_ERROR', '1');
    }
  }

  // Config シートから読み込み（新規スプシコピー時のメインパス）
  const configSheetData = readConfigSheet_();
  if (configSheetData) {
    const config = normalizeAppConfig_(Object.assign({}, configSheetData, sheetOverrides));
    try {
      props.setProperty('APP_CONFIG', JSON.stringify(config));
      props.deleteProperty('APP_CONFIG_PARSE_ERROR');
    } catch (e) {
      Logger.log('[loadConfig_] PropertiesService キャッシュ書き込み失敗: ' + e);
    }
    return config;
  }

  // デフォルト値にシート overrides を適用
  return normalizeAppConfig_(Object.assign({}, buildDefaultConfig_(), sheetOverrides));
}

/**
 * 設定を保存する。
 * PropertiesService（キャッシュ）・Config シート・設定シートの 3 箇所に書き込む。
 * @param {Object} config
 */
function saveConfig_(config) {
  const json = JSON.stringify(config);
  if (json.length > 8000) {
    throw new Error(
      '設定データが大きすぎます（' + json.length + ' バイト）。' +
      'フォーム回答スプレッドシートの件数や取り込み項目の数を減らしてください。'
    );
  }
  const props = PropertiesService.getScriptProperties();
  props.setProperty('APP_CONFIG', json);
  props.deleteProperty('APP_CONFIG_PARSE_ERROR');

  writeConfigSheet_(config);
  writeSettingsInputSheet_(config);
}

function getConfig_() {
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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

  throw new Error('「' + name + '」シートの構成が設定と一致しません。設定画面から保存して再生成してください。');
}

function ensureEvalSheet_(config) {
  return ensureConfiguredSheet_(EVAL_SHEET_NAME, getConfiguredEvalHeaders_(config));
}

function ensureSendSheet_(config) {
  return ensureConfiguredSheet_(SEND_SHEET_NAME, getConfiguredSendHeaders_(config));
}

function ensureStudentSheetForSync_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(STUDENT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(STUDENT_SHEET_NAME);
  }

  if (!hasMatchingHeaders_(sheet, STUDENT_SHEET_HEADERS)) {
    sheet.clear();
    sheet.getRange(1, 1, 1, STUDENT_SHEET_HEADERS.length).setValues([STUDENT_SHEET_HEADERS]);
  }

  sheet.setFrozenRows(1);
  return sheet;
}

function getResetImpactSummary_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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

function processConfiguredFieldValue_(field, rawValue, displayValue) {
  if (field.type === 'date') {
    return toYmdString(displayValue || rawValue);
  }
  if (field.type === 'score_grade') {
    return scoreToGrade(rawValue) || String(rawValue == null ? '' : rawValue).trim();
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
  return /[＜<]\s*([A-Za-z_][A-Za-z0-9_]*)\s*[＞>]/g;
}

function friendlyPlaceholderOnlyPattern_() {
  return /^[＜<]\s*[A-Za-z_][A-Za-z0-9_]*\s*[＞>]$/;
}

function replaceFriendlyPlaceholdersOnly_(text) {
  return normalizeMultiline_(text || '').replace(friendlyPlaceholderPattern_(), '{{$1}}');
}

function convertFriendlyTemplateToSections_(template, config) {
  const labels = buildTemplateLabelMap_(config);
  const normalized = normalizeMultiline_(template || '').trim();
  if (!normalized) return '';

  return normalized.split(/\n{2,}/).map(function(block) {
    const lines = block.split('\n');
    const lineInfos = lines.map(function(line) {
      const keys = [];
      line.replace(friendlyPlaceholderPattern_(), function(_, key) {
        keys.push(key);
        return _;
      });

      return {
        raw: line,
        keys: keys,
        replaced: line.replace(friendlyPlaceholderPattern_(), '{{$1}}'),
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

function updateFormStatus_(config, ssId, sheetName, rowIndex, status) {
  if (!ssId || !sheetName || !rowIndex) return;

  try {
    const ss = SpreadsheetApp.openById(ssId);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(String);
    const idxFlag = headers.indexOf(config.formStatusHeader);
    if (idxFlag < 0) return;

    sheet.getRange(rowIndex, idxFlag + 1).setValue(status);
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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
    throw new Error('クラス一覧が空です。先に「2 クラス一覧取得」を実行してください。');
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
    throw new Error('設定シートに FORM_SS が1件もありません。');
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
      if (!sheetMatchesPrefixes_(sheetName, prefixes)) return;

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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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
      throw new Error('対応表シートのヘッダが不正です。「4 対応表を作成」を実行してください。');
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

  const ss = SpreadsheetApp.getActiveSpreadsheet();
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

  safeAlert_('対応表シートを作成/更新しました（' + targetSheets.length + '件）');
}

function appendRows_(sheet, rows) {
  if (!rows || rows.length === 0) return;
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
}

function classroomdata() {
  const sheet = ensureSheet_('クラス一覧', ['クラス名', 'コースID', '同期対象(1)']);
  const response = Classroom.Courses.list({
    teacherId: Session.getActiveUser(),
    courseStates: 'ACTIVE',
    pageSize: 100
  });
  const courses = (response && response.courses) ? response.courses : [];
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
  if (courses.length === 0) return;
  appendRows_(sheet, courses.map(function(course) {
    return [course.name, course.id, ''];
  }));
}

function studentdataMulti() {
  const classSheet = ensureSheet_('クラス一覧', ['クラス名', 'コースID', '同期対象(1)']);
  const last = classSheet.getLastRow();
  if (last < 2) {
    safeAlert_('クラス一覧が空です');
    return;
  }

  const studentSheet = ensureStudentSheetForSync_();
  const lastRow = studentSheet.getLastRow();
  if (lastRow > 1) {
    studentSheet.getRange(2, 1, lastRow - 1, studentSheet.getLastColumn()).clearContent();
  }

  const rows = [];
  let no = 1;
  const data = classSheet.getRange(2, 1, last - 1, 3).getValues();

  data.forEach(function(item) {
    const courseName = item[0];
    const courseId = String(item[1] || '').trim();
    const syncFlag = String(item[2] || '').trim();
    if (!courseId || syncFlag !== '1') return;

    getStudentListMax(courseId).forEach(function(student) {
      rows.push([no++, student.email, student.name, courseName || '', courseId, student.studentId]);
    });
  });

  if (rows.length === 0) {
    safeAlert_('同期対象クラスが無いか、生徒が取得できませんでした');
    return;
  }

  appendRows_(studentSheet, rows);
  safeAlert_('複数クラスの生徒一覧を作成/更新しました（' + rows.length + '名）');
}

function getStudentListMax(classId) {
  let pageToken = '';
  const all = [];

  do {
    const response = Classroom.Courses.Students.list(classId, { pageToken: pageToken });
    const list = (response && response.students) ? response.students : [];
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
  const config = getConfig_();
  validateAppConfig_(config);
  const evalSheet = ensureEvalSheet_(config);
  const evalHeaders = evalSheet.getRange(1, 1, 1, evalSheet.getLastColumn()).getDisplayValues()[0].map(String);
  const evalHeaderMap = createHeaderMap_(evalHeaders);

  const scoreField = resolveScoreField_(config);
  if (!scoreField) {
    throw new Error('scoreSourceHeader に対応する項目が見つかりません。');
  }

  config.formSources.forEach(function(source) {
    const formSs = SpreadsheetApp.openById(source.id);
    formSs.getSheets().forEach(function(sheet) {
      const sheetName = sheet.getName();
      if (!sheetMatchesPrefixes_(sheetName, config.formSheetNamePrefix)) return;

      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow <= 1 || lastCol === 0) return;

      const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
      const displayValues = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
      const headers = values[0].map(String);
      const headerMap = createHeaderMap_(headers);

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

        const newRow = new Array(evalHeaders.length).fill('');
        newRow[evalHeaderMap['元SS_ID']] = source.id;
        newRow[evalHeaderMap['元シート名']] = sheetName;
        newRow[evalHeaderMap['元行番号']] = r + 1;
        newRow[evalHeaderMap['メールアドレス']] = email;
        newRow[evalHeaderMap['名前']] = studentName;
        newRow[evalHeaderMap['処理状態']] = '';

        config.fields.forEach(function(field) {
          const headerIndex = headerMap[field.sourceHeader];
          const processed = processConfiguredFieldValue_(field, row[headerIndex], dispRow[headerIndex]);
          newRow[evalHeaderMap[getFieldEvalHeader_(field)]] = processed;
        });

        rowsToAppend.push(newRow);
        statusUpdates.push({ rowIndex: r + 1, status: '反映〇' });
      }

      appendRows_(evalSheet, rowsToAppend);
      statusUpdates.forEach(function(update) {
        updateFormStatus_(config, source.id, sheetName, update.rowIndex, update.status);
      });
    });
  });
}

function evalToSendSheet() {
  const config = getConfig_();
  validateAppConfig_(config);

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
    const state = String(row[evalHeaderMap['処理状態']] || '').trim();
    if (state) continue;

    const srcSsId = String(row[evalHeaderMap['元SS_ID']] || '').trim();
    const srcSheetName = String(row[evalHeaderMap['元シート名']] || '').trim();
    const srcRow = Number(row[evalHeaderMap['元行番号']] || 0);
    const email = String(row[evalHeaderMap['メールアドレス']] || '').trim();
    const evalName = String(row[evalHeaderMap['名前']] || '').trim();

    if (!email) {
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
    stateUpdates.push({ rowNum: r + 1, evalState: '準備〇', srcState: '準備〇', srcSsId: srcSsId, srcSheetName: srcSheetName, srcRow: srcRow });
  }

  appendRows_(sendSheet, sendRows);

  stateUpdates.forEach(function(update) {
    evalSheet.getRange(update.rowNum, evalHeaderMap['処理状態'] + 1).setValue(update.evalState);
    if (update.srcSsId && update.srcSheetName && update.srcRow) {
      updateFormStatus_(config, update.srcSsId, update.srcSheetName, update.srcRow, update.srcState);
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
      Classroom.Courses.Announcements.create(data, String(courseId));
      Utilities.sleep(120);
      return;
    } catch (e) {
      const msg = String(e);
      if (t < maxRetry - 1 && (msg.indexOf('429') >= 0 || msg.indexOf('Rate') >= 0 || msg.indexOf('Internal') >= 0)) {
        Utilities.sleep(500 * (t + 1));
        continue;
      }
      throw e;
    }
  }
}

function sendMessages() {
  const config = getConfig_();
  validateAppConfig_(config);
  const sendSheet = ensureSendSheet_(config);
  const lastRow = sendSheet.getLastRow();
  const lastCol = sendSheet.getLastColumn();
  if (lastRow <= 1) {
    safeAlert_('送信シートにデータがありません。');
    return;
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

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const status = String(row[headerMap['送信状態']] || '').trim();
    if (status && status !== '未') continue;

    const srcSsId = String(row[headerMap['元SS_ID']] || '').trim();
    const srcSheetName = String(row[headerMap['元シート名']] || '').trim();
    const srcRow = Number(row[headerMap['元行番号']] || 0);
    const email = String(row[headerMap['メールアドレス']] || '').trim();
    const name = String(row[headerMap['名前']] || '').trim();
    const courseId = String(row[headerMap['コースID']] || '').trim();
    const studentId = String(row[headerMap['studentId']] || '').trim();
    const body = String(row[headerMap[config.replyBodyHeader]] || '').trim();

    const missingFields = [];
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
      if (srcSsId && srcSheetName && srcRow) {
        updateFormStatus_(config, srcSsId, srcSheetName, srcRow, '送信エラー');
      }
      errorCount++;
      continue;
    }

    try {
      postAnnouncementIndividual_(courseId, studentId, name, body);
      sendSheet.getRange(r + 1, headerMap['送信状態'] + 1).setValue('済');
      updateFormStatus_(config, srcSsId, srcSheetName, srcRow, '済');
      sentCount++;
    } catch (e) {
      const msg = String(e);
      let errType = 'FATAL';
      let newStatus = 'エラー';

      if (msg.indexOf('429') >= 0 || msg.indexOf('Rate') >= 0 || msg.indexOf('Internal') >= 0) {
        errType = 'RETRYABLE';
        newStatus = '未';
      } else {
        updateFormStatus_(config, srcSsId, srcSheetName, srcRow, '送信エラー');
      }

      sendSheet.getRange(r + 1, headerMap['送信状態'] + 1).setValue(newStatus);
      errorSheet.appendRow([
        new Date(),
        SEND_SHEET_NAME,
        r + 1,
        email,
        newStatus,
        errType,
        msg
      ]);
      errorCount++;
    }
  }

  safeAlert_(
    '送信処理が終了しました。\n' +
    '送信成功：' + sentCount + '件\n' +
    'エラー　：' + errorCount + '件（エラーシート参照）'
  );
}

function remindUngradedAndErrors() {
  const config = getConfig_();
  validateAppConfig_(config);

  if (!Array.isArray(config.reminderTo) || config.reminderTo.length === 0) {
    throw new Error('リマインドメール送信先が設定されていません。設定シートの REMINDER_TO を入力してください。');
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
      if (!sheetMatchesPrefixes_(sheetName, config.formSheetNamePrefix)) return;

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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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

  sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  safeAlert_('「評価データ」シートをクリアしました。');
}

function clearSendSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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

  sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  safeAlert_('「送信シート」をクリアしました。');
}

function clearErrorLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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

  ss.deleteSheet(sheet);
  safeAlert_('「エラー」シートを削除しました。');
}
