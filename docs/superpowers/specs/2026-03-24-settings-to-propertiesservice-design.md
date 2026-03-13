# 設計書：設定ストレージをPropertiesServiceへ移行 + サイドバー完結化

**日付:** 2026-03-24
**対象ファイル:** `Code.gs`, `Sidebar.html`

---

## 背景・目的

現在の設定ストレージは「設定シート」（KEY-VALUE行 + CONFIG_JSON行）で、サイドバーから保存するたびにシートを全消し・再生成する。`formSources`（FORM_SS）と `reminderTo` はサイドバー未対応で、設定シートを直接編集する必要があるが、次回サイドバー保存時に上書きされるデュアルソース問題がある。

**目標:**
- `ScriptProperties` を唯一の設定ストレージにする
- サイドバーですべての設定（`reminderTo`・`formSources` を含む）を完結させる
- 設定シートは読み取り専用の表示シート（`設定(編集不可)`）として残す
- ヘッダー候補の取得をサイドバー内で任意タイミングで実行できるようにする

---

## アーキテクチャ概要

```
[サイドバー UI]
      ↕ google.script.run
[Code.gs: saveSettingsFromSidebar / getSettingsSidebarData / fetchHeadersFromSources]
      ↕
[ScriptProperties] ← 唯一の設定ストレージ（APP_CONFIG キーにJSONを保存）
      ↓ （保存のたびに書き出し）
[設定(編集不可) シート] ← 人間が確認用に読むだけ、GASは読まない
```

---

## セクション1：データストアとマイグレーション

### 1-1. PropertiesServiceへの読み書き

```javascript
// 読み込み
function loadConfig_() {
  const props = PropertiesService.getScriptProperties();
  const json = props.getProperty('APP_CONFIG');
  if (json) {
    try {
      return normalizeAppConfig_(JSON.parse(json));
    } catch (e) {
      // パース失敗時はデフォルト設定を返すが、警告フラグを立てる
      // getSettingsSidebarData() でこのフラグを参照してサイドバーに警告表示
      Logger.log('[loadConfig_] APP_CONFIG のパース失敗。デフォルト設定を使用。: ' + e);
      PropertiesService.getScriptProperties().setProperty('APP_CONFIG_PARSE_ERROR', '1');
      return buildDefaultConfig_();
    }
  }
  // PropertiesServiceが空 = 初回のみマイグレーション試行
  return migrateFromLegacySheet_();
}

// 書き込み（9KB上限チェック付き）
// 保存成功時に APP_CONFIG_PARSE_ERROR フラグを削除する
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
  props.deleteProperty('APP_CONFIG_PARSE_ERROR'); // パースエラーフラグをクリア
}
```

### 1-2. マイグレーション（初回のみ）

`migrateFromLegacySheet_()` の処理手順：

1. `SETTINGS_SHEET_NAME`（旧名: `'設定シート'`）のシートを探す
2. シートが存在する場合、現行の `getConfig_()` と同一のロジックで設定を読む
   - `readConfigJsonFromSheet_` で `CONFIG_JSON` 行を優先読み取り
   - `readConfigOverridesFromSheet_` でKEY-VALUE行をオーバーライドとして適用
   - `normalizeAppConfig_` でマージ
3. 読み込みに成功したら `saveConfig_()` でPropertiesServiceに書き込む
4. 書き込み成功後に設定シートを `設定(編集不可)` にリネームする（シート名の変更はここで1回だけ行う）
5. シートが存在しない・読み込めない場合は `buildDefaultConfig_()` を返す。失敗はLoggerに記録する

**重要:** `SETTINGS_SHEET_NAME` 定数は `'設定(編集不可)'` に変更する。
マイグレーション中は旧シート名 `'設定シート'` をハードコードして参照し、
リネーム成功後は新定数 `'設定(編集不可)'` で統一する。
こうすることで、マイグレーション前に `ensureSettingsSheet_()` が呼ばれて
空の `設定(編集不可)` シートが誤生成されるのを防ぐ。

### 1-3. 表示用シート「設定(編集不可)」

- 定数変更：`SETTINGS_SHEET_NAME = '設定(編集不可)'`
- `writeDisplaySheet_(config)` を新規作成
  - シートをクリアして人間可読な内容を書き出す
  - 先頭行に `【注意】このシートは自動生成です。直接編集しても反映されません。` と注記を追加（絵文字不使用）
  - 以降の行は現在の `writeConfigToSettingsSheet_` と同等のKEY-VALUE形式
- `saveSettingsFromSidebar()` の中で `saveConfig_()` の直後に `writeDisplaySheet_()` を呼ぶ

---

## セクション2：サイドバーUIの変更

### 2-1. 新規追加する入力項目

| 項目 | HTML要素 | 備考 |
|------|----------|------|
| 通知先メールアドレス（`reminderTo`） | `<input type="email" id="reminderTo">` | 上部に配置 |
| フォーム回答SS（`formSources`） | 動的リスト | 下記参照 |

### 2-2. `formSources` の動的リストUI

```
[フォーム回答スプレッドシート]
  ┌────────────────────────────────────┬──────┐
  │ https://docs.google.com/...        │ 削除 │
  ├────────────────────────────────────┼──────┤
  │ https://docs.google.com/...        │ 削除 │
  └────────────────────────────────────┴──────┘
  [+ URLを追加]  [ヘッダを取得]
```

- `+ URLを追加` ボタン：新しいURL入力行を追加
- `ヘッダを取得` ボタン：現在入力中のURLリスト（未保存のものを含む）と
  現在入力中の接頭辞をGASに送り、ヘッダー候補を取得（保存は行わない）
- 取得結果はメール列名・名前列名などの `<select>` の選択肢に反映される

### 2-3. レイアウト順序（上から）

1. 通知先メールアドレス（新規）
2. フォーム回答スプレッドシート（新規、ヘッダを取得ボタン含む）
3. 対象シート名の接頭辞（既存）
4. 基本列（既存、変更なし）
5. 取り込み項目（既存、変更なし）
6. 返信テンプレート（既存、変更なし）
7. 保存してシート再生成 / 再読込（既存）

### 2-4. 削除する要素

- 「通知先メールアドレスとフォーム回答スプレッドシートは設定シートで管理します。」の注記

### 2-5. ヘッダー取得フロー変更

**現在の問題:**
- `getSettingsSidebarData()` が呼ばれると `collectAvailableSourceHeaders_(config)` が走る
- 初回は `formSources` が空 → ヘッダー候補が空 → `<select>` が空のまま

**変更後のフロー:**
1. サイドバーを開く → `getSettingsSidebarData()` は設定値のみ返す（ヘッダー取得は行わない）
2. URLと接頭辞を入力後に `ヘッダを取得` をクリック → `fetchHeadersFromSources(urlList, prefix)` を呼ぶ
3. GASが指定URLのシートからヘッダーを読んで返す
4. サイドバー側の `updateHeaderSelects(result)` が各 `<select>` の選択肢を更新する

**`getSettingsSidebarData` の戻り値から削除するフィールド:**
- `availableSourceHeaders`
- `sourceHeaderWarnings`

**新規GAS関数:**
```javascript
// urlList: 未正規化のURL文字列の配列（サイドバーの入力値をそのまま渡す）
// prefix: 対象シート名の接頭辞（サイドバーの現在入力値）
// 戻り値: { headers: string[], warnings: string[] }
function fetchHeadersFromSources(urlList, prefix) {
  // urlListの各エントリをnormalizeFormSource_でIDを抽出
  // 抽出失敗はwarningsに追記してスキップ（例外を投げない）
  // 各SSのシートを接頭辞でフィルタしてヘッダー行を収集
  // collectAvailableSourceHeaders_ の内部ロジックをベースに実装
}
```

**`collectAvailableSourceHeaders_` の変更:**
- 引数を `(config)` から `(urlList: string[], prefix: string)` に変更
- `getSettingsSidebarData` からは呼ばない
- `fetchHeadersFromSources` から呼ぶ

### 2-6. 変更するJavaScript関数（Sidebar.html）

| 関数名 | 変更内容 |
|--------|----------|
| `loadConfig()` | `getSettingsSidebarData` のレスポンスから `availableSourceHeaders`・`sourceHeaderWarnings` の参照を削除。設定値のみ `fillForm` に渡す。パースエラー警告フラグがある場合は警告メッセージを表示する |
| `renderSourceHeaderOptions()` | 削除。`fetchHeadersFromSources` のコールバック内の `updateHeaderSelects(result)` に置き換える |
| `updateHeaderSelects(result)` | 新規追加。`result.headers` で各 `<select>` を再描画し、`result.warnings` を表示する |
| `collectPayload()` | `reminderTo`・`formSources`（URLリスト）を追加して返す |
| `fillForm(config)` | `reminderTo`・`formSources` のフィールドに値をセットする処理を追加 |

---

## セクション3：Code.gsの変更範囲

### 3-1. 削除する関数・定数

| 対象 | 理由 |
|------|------|
| `readConfigJsonFromSheet_` | PropertiesServiceに移行（マイグレーション内で一時使用後に削除） |
| `readConfigOverridesFromSheet_` | 同上 |
| `writeConfigToSettingsSheet_` | `writeDisplaySheet_` に置き換え |
| `mergeSidebarPayloadWithSheetConfig_` | サイドバーが全フィールドを送るようになるため不要。`saveSettingsFromSidebar` は `normalizeAppConfig_(payload)` を直接呼ぶ |
| `SETTINGS_SHEET_HEADERS` | `writeConfigToSettingsSheet_` 削除に伴い不要 |
| `SETTINGS_FIELDS_HEADERS` | 同上 |
| `SETTINGS_JSON_KEY` | 同上 |

※ `readConfigJsonFromSheet_`・`readConfigOverridesFromSheet_`・`parseLegacyConfig_` の3関数は、
このPRでは削除せず `migrateFromLegacySheet_` から引き続き使用する。
マイグレーションが不要になったタイミング（全ユーザーが移行済みと判断した時点）に別PRで削除する。
その旨をコードコメントに明記する。削除テーブルの「マイグレーション内で一時使用後に削除」はこの意味である。

### 3-2. 新規追加する関数

| 関数名 | 役割 |
|--------|------|
| `loadConfig_()` | PropertiesService読み込み＋マイグレーション＋パースエラー処理 |
| `saveConfig_(config)` | 9KB上限チェック付きPropertiesService書き込み |
| `writeDisplaySheet_(config)` | 表示用シートへの書き出し |
| `migrateFromLegacySheet_()` | 初回移行（設定シートから読み込み → PropertiesService書き込み → シートリネーム） |
| `fetchHeadersFromSources(urlList, prefix)` | サイドバーから呼ぶヘッダー取得関数 |

### 3-3. 変更する関数・定数

| 対象 | 変更内容 |
|------|----------|
| `SETTINGS_SHEET_NAME` | `'設定シート'` → `'設定(編集不可)'` |
| `getConfig_()` | `loadConfig_()` の呼び出しに変更（または `loadConfig_` に統合して `getConfig_` を削除） |
| `getSettingsSidebarData()` | `availableSourceHeaders`・`sourceHeaderWarnings` を返さないように変更。パースエラーフラグを `configParseError: boolean` として返す |
| `saveSettingsFromSidebar(payload, confirmedReset)` | `mergeSidebarPayloadWithSheetConfig_` の代わりに `normalizeAppConfig_(payload)` を直接呼ぶ。`writeConfigToSettingsSheet_` → `saveConfig_` + `writeDisplaySheet_` に変更。`payload` に `reminderTo`・`formSources` が含まれることを前提とする |
| `collectAvailableSourceHeaders_` | 引数を `(config)` から `(urlList: string[], prefix: string)` に変更。`fetchHeadersFromSources` から呼ぶ |

### 3-4. 変更しない関数（業務ロジック）

- `normalizeAppConfig_`、`validateAppConfig_`、`normalizeFormSource_`
- `importFromFormsToEval`、`evalToSendSheet`、`sendMessages`
- `ensureEvalSheet_`、`ensureSendSheet_`、`ensureStudentSheetForSync_`
- `buildDefaultConfig_`、`normalizeFieldConfig_` など

### 3-5. `ensureSettingsSheet_()` の扱い

`ensureSettingsSheet_()` は現在 `ensureSheet_(SETTINGS_SHEET_NAME, SETTINGS_SHEET_HEADERS)` を呼んでいるが、
`SETTINGS_SHEET_HEADERS` は削除対象のため、このままでは実行時エラーになる。

対応：`ensureSettingsSheet_()` を削除し、呼び出し箇所を `writeDisplaySheet_()` に置き換える。
`writeDisplaySheet_()` はシートが存在しなければ新規作成する責務も持つ（`ensureSheet_` を内部で使う）。

---

## エラーハンドリング方針

- `fetchHeadersFromSources` でURLが不正・アクセス不可の場合は各URLごとに `try/catch` して `warnings` 配列に追記する。例外は外に伝播させない
- マイグレーション失敗時はデフォルト設定を返し、失敗理由をLoggerに記録する
- `loadConfig_` がPropertiesServiceの値のパースに失敗した場合はデフォルト設定を返し、`APP_CONFIG_PARSE_ERROR` フラグをPropertiesServiceに立てる。`getSettingsSidebarData` はこのフラグを確認して `configParseError: true` を返し、サイドバーに「設定データが壊れていました。デフォルト値を表示しています。保存すると上書きされます。」と警告を表示する
- `saveConfig_` が9KB上限を超える場合はユーザー向けエラーメッセージを throw する。サイドバーの `withFailureHandler` が受け取ってステータスエリアに表示する

---

## 影響を受けないもの

- `onOpen()` のメニュー構成
- `評価データ`・`送信シート`・`生徒一覧`・`対応表`・`エラー` シートの構造と操作
- Classroom API呼び出し部分
- テンプレートレンダリングロジック
