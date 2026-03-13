# Settings → PropertiesService Migration Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 設定ストレージを設定シート（KEY-VALUE）から ScriptProperties に移行し、サイドバーで `reminderTo`・`formSources` を含む全設定を完結させる。

**Architecture:** ScriptProperties の `APP_CONFIG` キーに JSON 1件として保存する。初回アクセス時のみ既存設定シートからマイグレーションを行い、以降はシートを読まない。表示用に `設定(編集不可)` シートへ書き出すが、GAS はこのシートを読まない。

**Tech Stack:** Google Apps Script (V8), HTML Service (Sidebar.html)

**Spec:** `docs/superpowers/specs/2026-03-24-settings-to-propertiesservice-design.md`

---

## ファイル変更マップ

| ファイル | 変更種別 | 内容 |
|----------|----------|------|
| `Code.gs` | 修正 | 新規5関数追加、既存4関数修正、3関数・3定数削除 |
| `Sidebar.html` | 修正 | 新規UI2セクション追加、JS5関数修正・追加・削除 |

---

## Task 1: PropertiesService 読み書き層の追加 + getConfig_ 更新

**Files:**
- Modify: `Code.gs:635-641` (getConfig_), `Code.gs:9` (SETTINGS_SHEET_NAME定数)
- Add functions: `loadConfig_`, `saveConfig_`, `migrateFromLegacySheet_` を `getConfig_` の直前に挿入

このタスク完了後、既存機能はすべて動作し続ける（migrateFromLegacySheet_ が既存設定シートから読むため）。

- [ ] **Step 1: SETTINGS_SHEET_NAME 定数を変更する**

`Code.gs` の9行目を変更：

```javascript
// Before:
const SETTINGS_SHEET_NAME = '設定シート';

// After:
const SETTINGS_SHEET_NAME = '設定(編集不可)';
```

- [ ] **Step 2: loadConfig_, saveConfig_, migrateFromLegacySheet_ を Code.gs の getConfig_ 直前（635行付近）に追加する**

```javascript
/**
 * PropertiesService から設定を読み込む。
 * APP_CONFIG が存在しない場合のみ既存設定シートからマイグレーションを試みる。
 */
function loadConfig_() {
  const props = PropertiesService.getScriptProperties();
  const json = props.getProperty('APP_CONFIG');
  if (json) {
    try {
      return normalizeAppConfig_(JSON.parse(json));
    } catch (e) {
      Logger.log('[loadConfig_] APP_CONFIG のパース失敗。デフォルト設定を使用。: ' + e);
      props.setProperty('APP_CONFIG_PARSE_ERROR', '1');
      return buildDefaultConfig_();
    }
  }
  return migrateFromLegacySheet_();
}

/**
 * PropertiesService に設定を書き込む。
 * 9KB 上限チェック付き。保存成功時にパースエラーフラグをクリアする。
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
}

/**
 * 初回のみ: 既存設定シートから PropertiesService へ移行する。
 * 移行後は設定シートを「設定(編集不可)」にリネームする。
 *
 * NOTE: readConfigJsonFromSheet_, readConfigOverridesFromSheet_, parseLegacyConfig_ は
 * このマイグレーション関数のためだけに残している。全ユーザーの移行が完了したと
 * 判断したタイミングで migrateFromLegacySheet_ ごとこれらを削除してよい。
 */
function migrateFromLegacySheet_() {
  const LEGACY_SHEET_NAME = '設定シート';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const legacySheet = ss.getSheetByName(LEGACY_SHEET_NAME);

  if (!legacySheet) {
    Logger.log('[migrateFromLegacySheet_] 旧設定シートが存在しないためデフォルト設定を使用。');
    return buildDefaultConfig_();
  }

  try {
    const json = readConfigJsonFromSheet_(legacySheet);
    const baseConfig = json || parseLegacyConfig_(legacySheet);
    const overrides = readConfigOverridesFromSheet_(legacySheet);
    const config = normalizeAppConfig_(Object.assign({}, baseConfig || {}, overrides));

    saveConfig_(config);

    // PropertiesService への書き込み成功後にリネーム
    legacySheet.setName(SETTINGS_SHEET_NAME);
    Logger.log('[migrateFromLegacySheet_] 設定シートを PropertiesService に移行し「' + SETTINGS_SHEET_NAME + '」にリネームしました。');

    return config;
  } catch (e) {
    Logger.log('[migrateFromLegacySheet_] マイグレーション失敗。デフォルト設定を使用。: ' + e);
    return buildDefaultConfig_();
  }
}
```

- [ ] **Step 3: getConfig_ を loadConfig_ に委譲するよう変更する**

```javascript
// Before (Code.gs:635-641):
function getConfig_() {
  const sheet = ensureSettingsSheet_();
  const json = readConfigJsonFromSheet_(sheet);
  const baseConfig = json || parseLegacyConfig_(sheet);
  const overrides = readConfigOverridesFromSheet_(sheet);
  return normalizeAppConfig_(Object.assign({}, baseConfig || {}, overrides));
}

// After:
function getConfig_() {
  return loadConfig_();
}
```

- [ ] **Step 4: Apps Script エディタで `getConfig_` を手動実行して動作確認する**

実行ログで以下のいずれかが確認できればOK:
- 既存設定シートがある場合: `[migrateFromLegacySheet_] 設定シートを PropertiesService に移行し「設定(編集不可)」にリネームしました。`
- すでに APP_CONFIG がある場合: ログなし（正常）

- [ ] **Step 5: コミットする**

```bash
git add Code.gs
git commit -m "feat: PropertiesServiceへの設定ストレージ移行（loadConfig_/saveConfig_/migrateFromLegacySheet_追加）"
```

---

## Task 2: writeDisplaySheet_ の追加と ensureSettingsSheet_ の削除

**Files:**
- Modify: `Code.gs:528-530` (ensureSettingsSheet_ を削除し writeDisplaySheet_ に置き換え)
- Add function: `writeDisplaySheet_` を `ensureSettingsSheet_` の位置に挿入

- [ ] **Step 1: writeDisplaySheet_ を追加し、ensureSettingsSheet_ を削除する**

`Code.gs` の528行付近を以下に置き換える（`ensureSettingsSheet_` を `writeDisplaySheet_` に置換）：

```javascript
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
    ['REMINDER_TO', config.reminderTo, 'リマインドメール送信先'],
    ['FORM_SHEET_PREFIX', config.formSheetNamePrefix, '対象シート名の接頭辞'],
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
```

- [ ] **Step 2: コミットする**

```bash
git add Code.gs
git commit -m "feat: writeDisplaySheet_を追加しensureSettingsSheet_を削除"
```

---

## Task 3: saveSettingsFromSidebar の更新

> **前提条件:** Task 2 (`writeDisplaySheet_` の追加) が完了していること。このタスクで呼ぶ `writeDisplaySheet_` は Task 2 で追加される。

**Files:**
- Modify: `Code.gs:96-122` (saveSettingsFromSidebar, mergeSidebarPayloadWithSheetConfig_)

- [ ] **Step 1: saveSettingsFromSidebar を更新し、mergeSidebarPayloadWithSheetConfig_ を削除する**

`Code.gs` の96〜122行を以下に置き換える：

```javascript
function saveSettingsFromSidebar(payload, confirmedReset) {
  // サイドバーが全フィールドを送るので merge は不要。直接 normalize する。
  const config = normalizeAppConfig_(payload);
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
```

`mergeSidebarPayloadWithSheetConfig_` 関数（120〜122行）を削除する。

- [ ] **Step 2: コミットする**

```bash
git add Code.gs
git commit -m "refactor: saveSettingsFromSidebarをPropertiesService保存に切り替え、mergeSidebarPayloadWithSheetConfig_を削除"
```

---

## Task 4: fetchHeadersFromSources の追加 + collectAvailableSourceHeaders_ の更新 + getSettingsSidebarData の更新

> **重要:** Step 1 と Step 2 は必ず同じ編集セッションで一緒に適用すること。Step 1 だけ適用した状態では `getSettingsSidebarData` が旧シグネチャ `(config)` で `collectAvailableSourceHeaders_` を呼び続けるため、動作が壊れる。

**Files:**
- Modify: `Code.gs:75-94` (getSettingsSidebarData)
- Modify: `Code.gs:124-181` (collectAvailableSourceHeaders_)
- Add: `fetchHeadersFromSources` 関数（collectAvailableSourceHeaders_ の直後に追加）

- [ ] **Step 1: collectAvailableSourceHeaders_ の引数を変更する**

`Code.gs` の124〜181行の関数を以下に置き換える：

```javascript
/**
 * 指定された URL リストとシート名接頭辞から利用可能なヘッダー候補を収集する。
 * @param {string[]} urlList - 未正規化の URL 文字列の配列（サイドバーの入力値）
 * @param {string} prefix - 対象シート名の接頭辞
 * @returns {{ headers: string[], warnings: string[] }}
 */
function collectAvailableSourceHeaders_(urlList, prefix) {
  const headerSet = {};
  const headers = [];
  const warnings = [];
  let matchedSheets = 0;
  const normalizedPrefix = String(prefix || '').trim();

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
        const sheetName = sheet.getName();
        return !normalizedPrefix || sheetName.indexOf(normalizedPrefix) === 0;
      });

      if (sheets.length === 0) {
        warnings.push(
          ss.getName() + ' に対象シートがありません。' +
          (normalizedPrefix ? ' 接頭辞「' + normalizedPrefix + '」を確認してください。' : '')
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
 * サイドバーの「ヘッダを取得」ボタンから呼ばれる。
 * 入力中の URL リストとシート名接頭辞を受け取り、ヘッダー候補を返す（保存は行わない）。
 * @param {string[]} urlList - サイドバーで入力中の URL 文字列の配列
 * @param {string} prefix - サイドバーで入力中のシート名接頭辞
 * @returns {{ headers: string[], warnings: string[] }}
 */
function fetchHeadersFromSources(urlList, prefix) {
  return collectAvailableSourceHeaders_(urlList, prefix);
}
```

- [ ] **Step 2: getSettingsSidebarData を更新する**

`Code.gs` の75〜94行を以下に置き換える：

```javascript
function getSettingsSidebarData() {
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
```

- [ ] **Step 3: コミットする**

```bash
git add Code.gs
git commit -m "feat: fetchHeadersFromSources追加、collectAvailableSourceHeaders_のシグネチャを変更、getSettingsSidebarDataを簡略化"
```

---

## Task 5: 不要コードの削除

**Files:**
- Modify: `Code.gs:16-18` (定数3件削除)
- Modify: `Code.gs:643-694` (writeConfigToSettingsSheet_ 削除)

- [ ] **Step 1: 不要定数 3 件を削除する**

`Code.gs` 16〜18行を削除する：

```javascript
// 削除対象（この3行を削除）:
const SETTINGS_SHEET_HEADERS = ['KEY', 'VALUE', 'NOTE'];
const SETTINGS_FIELDS_HEADERS = ['FIELD_KEY', 'SUMMARY', 'NOTE'];
const SETTINGS_JSON_KEY = 'CONFIG_JSON';
```

- [ ] **Step 2: writeConfigToSettingsSheet_ を削除する**

`Code.gs` の `writeConfigToSettingsSheet_` 関数（643〜694行付近）をまるごと削除する。

- [ ] **Step 3: コミットする**

```bash
git add Code.gs
git commit -m "chore: writeConfigToSettingsSheet_と不要定数3件を削除"
```

---

## Task 6: Sidebar.html の全面更新

**Files:**
- Modify: `Sidebar.html` 全体

このタスクは Sidebar.html を一気に書き直す。変更点は以下の通り：

1. `<body>` に `reminderTo` カード追加（最上部）
2. `formSources` 動的リストカード追加（2番目）
3. 既存の「対象シート名の接頭辞」カードの `<p class="note">` 削除（設定シート案内文）
4. JS: `sidebarData` から `availableSourceHeaders`・`sourceHeaderWarnings` の参照を削除
5. JS: `renderSourceHeaderOptions()` を削除、`updateHeaderSelects(result)` を追加
6. JS: `loadConfig()` に `configParseError` 警告処理を追加
7. JS: `fillForm(config)` に `reminderTo`・`formSources` のセット処理を追加
8. JS: `collectPayload()` に `reminderTo`・`formSources` を追加
9. JS: `appendFormSourceRow()` 関数を追加（formSources の動的リスト行生成）
10. JS: `readFormSourceUrls()` 関数を追加（formSources の URL リスト読み取り）
11. JS: `fetchHeaders()` 関数を追加（「ヘッダを取得」ボタンのハンドラ）

- [ ] **Step 1: Sidebar.html を以下の内容に書き換える**

```html
<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
    <style>
      body {
        font-family: "Segoe UI", sans-serif;
        margin: 0;
        padding: 16px;
        color: #1f2937;
        background: #f7f8fa;
      }
      h1 {
        margin: 0;
        font-size: 18px;
      }
      h2 {
        margin: 20px 0 8px;
        font-size: 14px;
      }
      p.note {
        margin: 6px 0 0;
        font-size: 12px;
        color: #6b7280;
        white-space: pre-line;
      }
      .card {
        background: #fff;
        border: 1px solid #d9dee7;
        border-radius: 10px;
        padding: 12px;
        margin-bottom: 12px;
      }
      label {
        display: block;
        font-size: 12px;
        font-weight: 600;
        margin-bottom: 4px;
      }
      input, textarea, select, button {
        box-sizing: border-box;
        width: 100%;
        border-radius: 8px;
        border: 1px solid #c7ced9;
        padding: 8px 10px;
        font: inherit;
      }
      textarea {
        min-height: 88px;
        resize: vertical;
      }
      button {
        cursor: pointer;
        background: #0f766e;
        border-color: #0f766e;
        color: #fff;
        font-weight: 600;
      }
      button.secondary {
        background: #fff;
        color: #0f172a;
        border-color: #c7ced9;
      }
      .page-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 8px;
        margin-bottom: 12px;
      }
      .page-header button {
        width: auto;
        flex-shrink: 0;
      }
      .row {
        display: grid;
        grid-template-columns: 1fr;
        gap: 8px;
        margin-bottom: 8px;
      }
      .stack {
        display: grid;
        gap: 8px;
      }
      .field-row {
        border: 1px solid #e5e7eb;
        border-radius: 8px;
        padding: 10px;
        margin-bottom: 10px;
        background: #fbfcfe;
      }
      .field-actions {
        display: flex;
        gap: 8px;
        margin-top: 8px;
      }
      .field-actions button {
        width: auto;
        padding: 8px 12px;
      }
      .status {
        margin-top: 12px;
        font-size: 12px;
        white-space: pre-line;
      }
      .status.error {
        color: #b91c1c;
      }
      .status.success {
        color: #047857;
      }
      .status.warning {
        color: #b45309;
      }
      .inline-actions {
        display: flex;
        gap: 8px;
      }
      .inline-actions button {
        width: auto;
      }
      .form-source-row {
        display: flex;
        gap: 6px;
        margin-bottom: 6px;
        align-items: center;
      }
      .form-source-row input {
        flex: 1;
      }
      .form-source-row button {
        width: auto;
        flex-shrink: 0;
        padding: 8px 10px;
      }
    </style>
  </head>
  <body>
    <div class="page-header">
      <h1>送信設定</h1>
      <button type="button" class="secondary" id="reloadTopButton">再読込</button>
    </div>

    <div class="card">
      <label for="reminderTo">通知先メールアドレス</label>
      <input id="reminderTo" type="email" placeholder="teacher@example.com">
      <p class="note">未採点・エラーのリマインダーメールの送信先です。</p>
    </div>

    <div class="card">
      <h2>フォーム回答スプレッドシート</h2>
      <div id="formSources"></div>
      <div class="inline-actions" style="margin-top:8px;">
        <button type="button" class="secondary" id="addFormSourceButton">+ URLを追加</button>
        <button type="button" class="secondary" id="fetchHeadersButton">ヘッダを取得</button>
      </div>
      <p class="note">URLを入力後「ヘッダを取得」を押すと、下の列名選択肢が更新されます。</p>
    </div>

    <div class="card">
      <label for="formSheetNamePrefix">対象シート名の接頭辞</label>
      <input id="formSheetNamePrefix" type="text">
    </div>

    <div class="card">
      <h2>基本列</h2>
      <p class="note">自動送信の際、Form側のスプレッドシートで必要になる情報列を入力します。</p>
      <div class="stack">
        <div>
          <label for="emailHeader">メール列名</label>
          <select id="emailHeader"></select>
        </div>
        <div>
          <label for="studentNameHeader">名前列名</label>
          <select id="studentNameHeader"></select>
        </div>
        <div>
          <label for="formStatusHeader">送信状態列名</label>
          <select id="formStatusHeader"></select>
        </div>
        <div>
          <label for="replyBodyHeader">送信シートの返信本文列名</label>
          <input id="replyBodyHeader" type="text">
        </div>
        <div>
          <label for="scoreSourceHeader">採点済み判定に使うフォーム側ヘッダ</label>
          <select id="scoreSourceHeader"></select>
        </div>
      </div>
      <p class="note">未採点判定はメール列名を使います（メールアドレス列を推奨）。</p>
      <p class="note">採点時に必ず入力され、生徒に送信する列を選んでください。</p>
      <p class="note" id="sourceHeaderGuide"></p>
    </div>

    <div class="card">
      <div class="inline-actions">
        <h2 style="flex:1;margin:0;">取り込み項目</h2>
        <button type="button" class="secondary" id="addFieldButton">項目を追加</button>
      </div>
      <div id="fields"></div>
      <p class="note">テンプレート変数名は返信テンプレートで使う名前です。半角英字で始めて、英数字とアンダースコアのみ使えます。例: `score`, `submission_text`</p>
      <p class="note">評価データ列名と送信シート列名を空欄にすると、フォーム側ヘッダ名をそのまま使います。</p>
    </div>

    <div class="card">
      <h2>返信テンプレート</h2>
      <textarea id="messageTemplate"></textarea>
      <p class="note" id="templateGuide"></p>
    </div>

    <div class="inline-actions">
      <button type="button" id="saveButton">保存してシート再生成</button>
      <button type="button" class="secondary" id="reloadButton">再読込</button>
    </div>

    <div id="status" class="status"></div>

    <template id="fieldTemplate">
      <div class="field-row">
        <div class="row">
          <div>
            <label>テンプレート変数名</label>
            <input data-name="key" type="text">
          </div>
          <div>
            <label>type</label>
            <select data-name="type"></select>
          </div>
        </div>
        <div class="row">
          <div>
            <label>フォーム側ヘッダ</label>
            <select data-name="sourceHeader"></select>
          </div>
          <div>
            <label>評価データ列名（任意）</label>
            <input data-name="evalHeader" type="text">
          </div>
        </div>
        <div>
          <label>送信シート列名（任意）</label>
          <input data-name="sendHeader" type="text">
        </div>
        <div class="field-actions">
          <button type="button" class="secondary" data-action="remove">削除</button>
        </div>
      </div>
    </template>

    <script>
      let sidebarData = null;
      // 現在のヘッダー候補（fetchHeaders 実行後に更新）
      let currentHeaders = [];

      const defaultFieldTypes = [
        { value: 'text', label: 'text' },
        { value: 'date', label: 'date' },
        { value: 'score_grade', label: 'score_grade' }
      ];

      document.getElementById('addFieldButton').addEventListener('click', function() {
        appendFieldRow();
      });

      document.getElementById('reloadButton').addEventListener('click', function() {
        loadConfig();
      });
      document.getElementById('reloadTopButton').addEventListener('click', function() {
        loadConfig();
      });

      document.getElementById('saveButton').addEventListener('click', function() {
        save(false);
      });

      document.getElementById('addFormSourceButton').addEventListener('click', function() {
        appendFormSourceRow('');
      });

      document.getElementById('fetchHeadersButton').addEventListener('click', function() {
        fetchHeaders();
      });

      function setStatus(message, kind) {
        const el = document.getElementById('status');
        el.textContent = message || '';
        el.className = 'status' + (kind ? ' ' + kind : '');
      }

      function getAvailableFieldTypes() {
        return sidebarData && sidebarData.availableFieldTypes
          ? sidebarData.availableFieldTypes
          : defaultFieldTypes;
      }

      function populateTypeOptions(select) {
        select.innerHTML = '';
        getAvailableFieldTypes().forEach(function(item) {
          const option = document.createElement('option');
          option.value = item.value;
          option.textContent = item.label;
          select.appendChild(option);
        });
      }

      function populateHeaderSelect(select, value) {
        const headers = currentHeaders;
        const targetValue = value || '';

        select.innerHTML = '';
        const emptyOption = document.createElement('option');
        emptyOption.value = '';
        emptyOption.textContent = '';
        select.appendChild(emptyOption);

        headers.forEach(function(header) {
          const option = document.createElement('option');
          option.value = header;
          option.textContent = header;
          select.appendChild(option);
        });

        if (targetValue && !headers.includes(targetValue)) {
          const currentOption = document.createElement('option');
          currentOption.value = targetValue;
          currentOption.textContent = targetValue;
          select.appendChild(currentOption);
        }

        select.value = targetValue;
      }

      /**
       * fetchHeadersFromSources の結果を受け取り、全 select を更新する。
       */
      function updateHeaderSelects(result) {
        currentHeaders = result.headers || [];

        // 基本列の select を更新（現在の選択値を保持）
        ['emailHeader', 'studentNameHeader', 'formStatusHeader', 'scoreSourceHeader'].forEach(function(id) {
          const sel = document.getElementById(id);
          if (sel) populateHeaderSelect(sel, sel.value);
        });

        // 取り込み項目の sourceHeader select を更新
        document.querySelectorAll('[data-name="sourceHeader"]').forEach(function(sel) {
          populateHeaderSelect(sel, sel.value);
        });

        const guide = document.getElementById('sourceHeaderGuide');
        const lines = [];
        if (currentHeaders.length > 0) {
          lines.push('フォーム回答スプレッドシートのヘッダー候補から選べます。');
        } else {
          lines.push('ヘッダー候補を取得できていません。URLと接頭辞を確認して「ヘッダを取得」を押してください。');
        }
        (result.warnings || []).forEach(function(msg) {
          lines.push('注意: ' + msg);
        });
        guide.textContent = lines.join('\n');
      }

      /**
       * formSources リストに URL 入力行を1件追加する。
       */
      function appendFormSourceRow(url) {
        const container = document.getElementById('formSources');
        const row = document.createElement('div');
        row.className = 'form-source-row';

        const input = document.createElement('input');
        input.type = 'url';
        input.placeholder = 'https://docs.google.com/spreadsheets/d/...';
        input.value = url || '';

        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'secondary';
        btn.textContent = '削除';
        btn.addEventListener('click', function() {
          row.remove();
        });

        row.appendChild(input);
        row.appendChild(btn);
        container.appendChild(row);
      }

      /**
       * formSources コンテナから URL 文字列の配列を読み取る。
       */
      function readFormSourceUrls() {
        return Array.from(document.querySelectorAll('#formSources .form-source-row input'))
          .map(function(input) { return input.value.trim(); })
          .filter(function(url) { return url !== ''; });
      }

      function appendFieldRow(field) {
        const template = document.getElementById('fieldTemplate');
        const node = template.content.firstElementChild.cloneNode(true);
        const typeSelect = node.querySelector('select[data-name="type"]');
        const sourceHeaderSelect = node.querySelector('select[data-name="sourceHeader"]');
        populateTypeOptions(typeSelect);
        populateHeaderSelect(sourceHeaderSelect, field && field.sourceHeader ? field.sourceHeader : '');

        ['key', 'evalHeader', 'sendHeader'].forEach(function(name) {
          node.querySelector('[data-name="' + name + '"]').value = field && field[name] ? field[name] : '';
        });
        typeSelect.value = field && field.type ? field.type : 'text';

        node.querySelector('[data-action="remove"]').addEventListener('click', function() {
          node.remove();
        });

        document.getElementById('fields').appendChild(node);
      }

      function readFieldRows() {
        return Array.from(document.querySelectorAll('.field-row')).map(function(row) {
          return {
            key: row.querySelector('[data-name="key"]').value.trim(),
            sourceHeader: row.querySelector('[data-name="sourceHeader"]').value.trim(),
            evalHeader: row.querySelector('[data-name="evalHeader"]').value.trim(),
            sendHeader: row.querySelector('[data-name="sendHeader"]').value.trim(),
            type: row.querySelector('[data-name="type"]').value
          };
        }).filter(function(field) {
          return field.key || field.sourceHeader || field.evalHeader || field.sendHeader;
        });
      }

      function collectPayload() {
        return {
          reminderTo: document.getElementById('reminderTo').value.trim(),
          formSources: readFormSourceUrls(),
          formSheetNamePrefix: document.getElementById('formSheetNamePrefix').value.trim(),
          emailHeader: document.getElementById('emailHeader').value.trim(),
          studentNameHeader: document.getElementById('studentNameHeader').value.trim(),
          formStatusHeader: document.getElementById('formStatusHeader').value.trim(),
          replyBodyHeader: document.getElementById('replyBodyHeader').value.trim(),
          scoreSourceHeader: document.getElementById('scoreSourceHeader').value.trim(),
          fields: readFieldRows(),
          messageTemplate: document.getElementById('messageTemplate').value
        };
      }

      function fillForm(config) {
        document.getElementById('reminderTo').value = config.reminderTo || '';

        // formSources: 既存リストをクリアして再描画
        document.getElementById('formSources').innerHTML = '';
        const sources = config.formSources || [];
        if (sources.length === 0) {
          appendFormSourceRow('');
        } else {
          sources.forEach(function(source) {
            const url = (source && source.url) ? source.url : String(source || '');
            appendFormSourceRow(url);
          });
        }

        document.getElementById('formSheetNamePrefix').value = config.formSheetNamePrefix || '';
        populateHeaderSelect(document.getElementById('emailHeader'), config.emailHeader || '');
        populateHeaderSelect(document.getElementById('studentNameHeader'), config.studentNameHeader || '');
        populateHeaderSelect(document.getElementById('formStatusHeader'), config.formStatusHeader || '');
        document.getElementById('replyBodyHeader').value = config.replyBodyHeader || '';
        populateHeaderSelect(document.getElementById('scoreSourceHeader'), config.scoreSourceHeader || '');
        document.getElementById('messageTemplate').value = config.messageTemplate || '';
        document.getElementById('templateGuide').textContent = sidebarData && sidebarData.templateGuide || '';

        const fields = document.getElementById('fields');
        fields.innerHTML = '';
        (config.fields || []).forEach(function(field) {
          appendFieldRow(field);
        });
      }

      function loadConfig() {
        setStatus('設定を読み込んでいます…');
        google.script.run
          .withSuccessHandler(function(data) {
            sidebarData = data;
            fillForm(data.config);
            if (data.configParseError) {
              setStatus('設定データが壊れていました。デフォルト値を表示しています。保存すると上書きされます。', 'warning');
            } else {
              setStatus('');
            }
          })
          .withFailureHandler(function(error) {
            setStatus(error.message || String(error), 'error');
          })
          .getSettingsSidebarData();
      }

      function fetchHeaders() {
        const urlList = readFormSourceUrls();
        const prefix = document.getElementById('formSheetNamePrefix').value.trim();
        setStatus('ヘッダーを取得しています…');

        google.script.run
          .withSuccessHandler(function(result) {
            updateHeaderSelects(result);
            const hasWarnings = result.warnings && result.warnings.length > 0;
            setStatus(
              'ヘッダーを取得しました（' + result.headers.length + '件）' +
              (hasWarnings ? '\n注意: ' + result.warnings.join('\n注意: ') : ''),
              hasWarnings ? 'warning' : 'success'
            );
          })
          .withFailureHandler(function(error) {
            setStatus(error.message || String(error), 'error');
          })
          .fetchHeadersFromSources(urlList, prefix);
      }

      function save(confirmedReset) {
        const payload = collectPayload();
        setStatus('保存しています…');

        google.script.run
          .withSuccessHandler(function(result) {
            if (result.requiresConfirmation && !confirmedReset) {
              const ok = window.confirm(result.summary);
              if (ok) {
                save(true);
              } else {
                setStatus('保存をキャンセルしました。');
              }
              return;
            }
            setStatus(result.message || '保存しました。', 'success');
            loadConfig();
          })
          .withFailureHandler(function(error) {
            setStatus(error.message || String(error), 'error');
          })
          .saveSettingsFromSidebar(payload, confirmedReset);
      }

      loadConfig();
    </script>
  </body>
</html>
```

- [ ] **Step 2: Apps Script にデプロイして動作確認する**

以下の順に確認する：

1. サイドバーを開く → エラーなく表示されること
2. 通知先メールアドレス欄・フォーム回答SS欄が表示されること
3. フォーム回答SSにURLを1件入力し「ヘッダを取得」を押す → select の選択肢が更新されること
4. 設定を入力して「保存してシート再生成」を押す → 成功メッセージが出ること
5. 再読込し、入力した値が復元されること（reminderTo・formSources を含む）
6. スプレッドシートで `設定(編集不可)` シートが更新されていること

- [ ] **Step 3: コミットする**

```bash
git add Sidebar.html
git commit -m "feat: サイドバーにreminderTo・formSources追加、ヘッダを取得ボタン実装"
```

---

## Task 7: ファイルヘッダコメントの更新

**Files:**
- Modify: `Code.gs:1-7` (ファイル先頭コメント)

- [ ] **Step 1: ファイル先頭コメントを実態に合わせて更新する**

```javascript
/******************************************************
 * Google Classroom 評価自動送信 GAS
 *
 * - 設定は ScriptProperties に JSON で保存（唯一のストレージ）
 * - サイドバーからすべての設定を完結させる
 * - 設定(編集不可)シートは表示専用（GAS は読まない）
 * - 初回のみ旧設定シートから PropertiesService へ自動移行
 ******************************************************/
```

- [ ] **Step 2: コミットする**

```bash
git add Code.gs
git commit -m "docs: ファイルヘッダコメントをPropertiesService移行後の実態に更新"
```

---

## 完了確認チェックリスト

- [ ] `getConfig_()` が PropertiesService から読んでいる（Settings シートを参照しない）
- [ ] `saveSettingsFromSidebar()` が `saveConfig_()` + `writeDisplaySheet_()` を呼んでいる
- [ ] `設定(編集不可)` シートが「このシートは自動生成です」注記付きで生成されている
- [ ] サイドバーで reminderTo・formSources が保存・復元できる
- [ ] 「ヘッダを取得」が保存前の入力値から動作する
- [ ] `SETTINGS_SHEET_HEADERS`・`SETTINGS_FIELDS_HEADERS`・`SETTINGS_JSON_KEY` 定数が存在しない
- [ ] `writeConfigToSettingsSheet_`・`mergeSidebarPayloadWithSheetConfig_`・`ensureSettingsSheet_` が存在しない
