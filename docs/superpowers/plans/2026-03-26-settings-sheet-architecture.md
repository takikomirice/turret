# Settings Sheet Architecture Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 設定シートを REMINDER_TO・FORM_SS・PREFIX の権威的ソースにし、サイドバーを開かずに対応表を作成できるようにする。学期コピー時に設定がそのまま引き継がれるアーキテクチャへ移行する。

**Architecture:** 3層ストレージ。設定シート（手動編集可）は reminderTo・formSources・formSheetNamePrefix の権威的ソース。Config シート（GAS 管理）はフィールド/列名/テンプレートを CONFIG_JSON として保持。PropertiesService はランタイムキャッシュ。`loadConfig_()` は常に設定シートを読んで overrides として適用し、複雑な設定は PropertiesService → Config シート → デフォルト値の順にフォールバック。

**Tech Stack:** Google Apps Script (V8)

---

## ファイル変更マップ

| ファイル | 変更種別 | 内容 |
|----------|----------|------|
| `Code.gs` | 修正 | 定数1件追加、関数6件追加、関数2件修正、関数4件削除、ヘッダコメント更新、onOpen修正 |

---

## ⚠️ 重要: タスク実行順序について

`migrateFromLegacySheet_()` は内部で `LEGACY_SHEET_NAME = '設定シート'` という変数を持っており、**新しい設定シートと名前が衝突する**。Task 3 で `loadConfig_()` を更新して `migrateFromLegacySheet_()` の呼び出しを除去し、Task 5 で関数本体を削除する。Task 1〜6 の順番通りに実行すること。

---

## Task 1: Config シートの読み書きヘルパーを追加

**Files:**
- Modify: `Code.gs` — 定数1件追加（11行目付近）、関数2件追加（loadConfig_ の直前、約744行目付近）

- [ ] **Step 1: CONFIG_SHEET_NAME 定数を追加する**

`Code.gs` の11〜12行目を以下に変更（INPUT_SETTINGS_SHEET_NAME の直後に追加）：

```javascript
const INPUT_SETTINGS_SHEET_NAME = '設定シート';
const CONFIG_SHEET_NAME = 'Config';
```

- [ ] **Step 2: readConfigSheet_ と writeConfigSheet_ を loadConfig_ の直前（744行目付近）に追加する**

```javascript
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG_SHEET_NAME);
  }
  sheet.getRange(1, 1).setValue(JSON.stringify(config));
}
```

- [ ] **Step 3: Apps Script エディタで動作確認する**

以下のテスト関数を一時的に追加して実行する：

```javascript
function testConfigSheet() {
  writeConfigSheet_({ test: true, fields: [] });
  const result = readConfigSheet_();
  Logger.log('read back: ' + JSON.stringify(result));
  // 期待: {"test":true,"fields":[]}
}
```

確認項目：
- Config シートが作成されていること
- A1 セルに JSON 文字列が入っていること
- `readConfigSheet_()` が同じオブジェクトを返すこと

テスト関数は確認後に削除する。

- [ ] **Step 4: コミットする**

```bash
git add Code.gs
git commit -m "feat: ConfigシートのCONFIG_JSON読み書きヘルパーを追加（readConfigSheet_/writeConfigSheet_）"
```

---

## Task 2: 設定シートの読み書きヘルパーを追加

**Files:**
- Modify: `Code.gs` — 関数2件追加（Task 1 で追加した writeConfigSheet_ の直後）

- [ ] **Step 1: readSettingsInputSheet_ を追加する**

Task 1 で追加した `writeConfigSheet_()` の直後に挿入：

```javascript
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
  const formSources = [];

  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0] || '').trim();
    const value = String(values[i][1] || '').trim();
    if (!key || !value) continue;

    switch (key) {
      case 'REMINDER_TO':
        result.reminderTo = value;
        break;
      case 'FORM_SHEET_PREFIX':
        result.formSheetNamePrefix = value;
        break;
      case 'FORM_SS': {
        const normalized = normalizeFormSource_(value);
        if (normalized) formSources.push(normalized);
        break;
      }
    }
  }

  if (formSources.length > 0) {
    result.formSources = formSources;
  }

  return result;
}
```

- [ ] **Step 2: writeSettingsInputSheet_ を追加する**

`readSettingsInputSheet_()` の直後に挿入：

```javascript
/**
 * 設定シートの REMINDER_TO・FORM_SHEET_PREFIX・FORM_SS 行を更新する。
 * シートの他の行は一切変更しない。
 * 既存行が見つからない場合はシート末尾に追加する。
 * FORM_SS は件数が増減した場合に行を追加／削除する。
 * @param {Object} config
 */
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

  let reminderToRow = -1;
  let prefixRow = -1;
  const formSsRows = [];

  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0] || '').trim();
    if (key === 'REMINDER_TO' && reminderToRow < 0) reminderToRow = i + 1;
    else if (key === 'FORM_SHEET_PREFIX' && prefixRow < 0) prefixRow = i + 1;
    else if (key === 'FORM_SS') formSsRows.push(i + 1);
  }

  if (reminderToRow > 0) {
    sheet.getRange(reminderToRow, 2).setValue(config.reminderTo || '');
  } else {
    sheet.appendRow(['REMINDER_TO', config.reminderTo || '', 'リマインドメール送信先（必須）']);
  }

  if (prefixRow > 0) {
    sheet.getRange(prefixRow, 2).setValue(config.formSheetNamePrefix || '');
  } else {
    sheet.appendRow(['FORM_SHEET_PREFIX', config.formSheetNamePrefix || '', '対象シート名の接頭辞']);
  }

  const newSources = (config.formSources || [])
    .map(function(s) { return normalizeFormSource_(s); })
    .filter(function(s) { return !!s; });

  // 既存 FORM_SS 行を更新
  for (let j = 0; j < newSources.length && j < formSsRows.length; j++) {
    sheet.getRange(formSsRows[j], 2).setValue(newSources[j].url);
  }

  // 余分な FORM_SS 行を削除（下から順に）
  for (let k = formSsRows.length - 1; k >= newSources.length; k--) {
    sheet.deleteRow(formSsRows[k]);
  }

  // 不足分を末尾に追加
  for (let m = formSsRows.length; m < newSources.length; m++) {
    sheet.appendRow(['FORM_SS', newSources[m].url, '']);
  }
}
```

- [ ] **Step 3: Apps Script エディタで動作確認する**

設定シートの VALUE 列に以下を手動入力しておく：
```
KEY              VALUE
REMINDER_TO      test@example.com
FORM_SHEET_PREFIX  フォームの回答
FORM_SS          https://docs.google.com/spreadsheets/d/DUMMY_ID/edit
```

確認用関数を実行：

```javascript
function testReadSettingsInputSheet() {
  const result = readSettingsInputSheet_();
  Logger.log(JSON.stringify(result));
  // 期待: {"reminderTo":"test@example.com","formSheetNamePrefix":"フォームの回答","formSources":[{...}]}
}
```

`writeSettingsInputSheet_` の確認：

```javascript
function testWriteSettingsInputSheet() {
  const config = getConfig_();
  writeSettingsInputSheet_(config);
  Logger.log('done');
  // 設定シートの REMINDER_TO・FORM_SHEET_PREFIX・FORM_SS が config の値で更新されていることを確認
}
```

テスト関数は確認後に削除する。

- [ ] **Step 4: コミットする**

```bash
git add Code.gs
git commit -m "feat: 設定シートの読み書きヘルパーを追加（readSettingsInputSheet_/writeSettingsInputSheet_）"
```

---

## Task 3: loadConfig_() を更新

**Files:**
- Modify: `Code.gs:744-757` (loadConfig_ を置き換え)

新しいフォールバックチェーン：
1. 設定シートを**常に**読んで overrides を取得（REMINDER_TO・FORM_SS・PREFIX はここが優先）
2. PropertiesService にキャッシュがあれば overrides をマージして返す
3. Config シートに CONFIG_JSON があれば PropertiesService にキャッシュし、overrides をマージして返す
4. どちらもなければデフォルト値に overrides をマージして返す

- [ ] **Step 1: loadConfig_() を以下に置き換える**

既存の `loadConfig_()` 関数（744〜757行目）をまるごと以下で置き換える：

```javascript
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
```

- [ ] **Step 2: Apps Script エディタで動作確認する**

```javascript
function testLoadConfigFromSheets() {
  // PropertiesService を一時クリアしてシートから読む経路を確認
  PropertiesService.getScriptProperties().deleteProperty('APP_CONFIG');
  const config = loadConfig_();
  Logger.log('reminderTo: ' + config.reminderTo);
  Logger.log('formSources count: ' + config.formSources.length);
  Logger.log('fields count: ' + config.fields.length);
  // reminderTo と formSources は設定シートの値が入っていること
  // fields は Config シートまたはデフォルト値が入っていること
}
```

テスト関数は確認後に削除する。

- [ ] **Step 3: コミットする**

```bash
git add Code.gs
git commit -m "feat: loadConfig_を3層フォールバック対応に更新（設定シート常時読み込み・Configシートフォールバック）"
```

---

## Task 4: saveConfig_() を更新

**Files:**
- Modify: `Code.gs:763-774` (saveConfig_ を置き換え)

- [ ] **Step 1: saveConfig_() を以下に置き換える**

既存の `saveConfig_()` 関数（763〜774行目）をまるごと以下で置き換える：

```javascript
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
```

- [ ] **Step 2: Apps Script エディタで動作確認する**

```javascript
function testSaveConfig() {
  const config = loadConfig_();
  saveConfig_(config);
  Logger.log('save complete');
  // 確認項目:
  // 1. Config シートの A1 に JSON が入っている
  // 2. 設定シートの REMINDER_TO・FORM_SHEET_PREFIX・FORM_SS が config の値になっている
  // 3. PropertiesService に APP_CONFIG が保存されている
}
```

テスト関数は確認後に削除する。

- [ ] **Step 3: コミットする**

```bash
git add Code.gs
git commit -m "feat: saveConfig_をConfigシート・設定シートへの書き込みに対応"
```

---

## Task 5: 旧マイグレーションコード4件を削除・ファイルヘッダを更新

**Files:**
- Modify: `Code.gs` — 関数4件削除（637〜840行目付近）、ヘッダコメント更新（1〜8行目）

以下の関数は旧設定シート（KEY-VALUE シートベースストレージ）からのマイグレーション用であり、今回の変更で不要になる。また `migrateFromLegacySheet_()` は内部で `'設定シート'` を参照しており、新しい設定シートと名前が衝突するため必ず削除する。

削除対象：
- `function readConfigJsonFromSheet_(sheet)` （637行目付近）
- `function parseLegacyConfig_(sheet)` （650行目付近）
- `function readConfigOverridesFromSheet_(sheet)` （684行目付近）
- `function migrateFromLegacySheet_()` （784行目付近）

- [ ] **Step 1: 4関数を Code.gs からまるごと削除する**

各関数を `function xxx(` の行から末尾の `}` まで（JSDoc コメントも含めて）削除する。

削除後に以下の関数が Code.gs に存在しないことを grep で確認：

```bash
grep -n "function readConfigJsonFromSheet_\|function parseLegacyConfig_\|function readConfigOverridesFromSheet_\|function migrateFromLegacySheet_" Code.gs
# → 出力なしであればOK
```

- [ ] **Step 2: ファイルヘッダコメントを更新する**

`Code.gs` の1〜8行目を以下に置き換える：

```javascript
/******************************************************
 * Google Classroom 評価自動送信 GAS
 *
 * - 設定シート（手動編集可）: reminderTo・formSources・prefix の権威的ソース
 * - Config シート（GAS 管理）: フィールド・列名・テンプレートを JSON で保持
 * - ScriptProperties: ランタイムキャッシュ（設定シート変更は即時反映）
 * - 学期コピー時は設定シート・Config シートがそのまま引き継がれる
 ******************************************************/
```

- [ ] **Step 3: コミットする**

```bash
git add Code.gs
git commit -m "refactor: 旧マイグレーションコード4件削除・ファイルヘッダをアーキテクチャ変更後の実態に更新"
```

---

## Task 6: onOpen メニューを更新

**Files:**
- Modify: `Code.gs:42-64` (onOpen を置き換え)

「設定を開く」は番号なしで最上部に常設。対応表を4番に移動。

- [ ] **Step 1: onOpen() を以下に置き換える**

```javascript
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('自動送信システム')
    .addItem('設定を開く', 'openSettingsSidebar')
    .addSeparator()
    .addItem('1 クラス/生徒一覧・設定シート作成', 'initializeSheets')
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
```

- [ ] **Step 2: コミットする**

```bash
git add Code.gs
git commit -m "feat: onOpenメニュー更新（設定を開くを番号外に・対応表を4番に）"
```

---

## 完了確認チェックリスト

- [ ] 設定シートを直接編集した FORM_SS が、サイドバーを開かずに `getConfig_()` の結果に反映される
- [ ] PropertiesService を削除した状態で `loadConfig_()` を実行し、Config シートから設定が読めている
- [ ] `saveConfig_()` 実行後、Config シート・設定シート・PropertiesService の3箇所が更新されている
- [ ] 旧マイグレーション関数4件（readConfigJsonFromSheet_/parseLegacyConfig_/readConfigOverridesFromSheet_/migrateFromLegacySheet_）が Code.gs に存在しない
- [ ] 「設定を開く」がメニューの番号なしアイテムとして最上部にある
- [ ] 「4 対応表を作成」が設定シートに FORM_SS を入れた状態で動作する（サイドバー保存不要）
- [ ] 学期コピーのシナリオ（PropertiesService 空・Config シートあり・設定シートあり）で設定が正しく読み込まれる
