# フォーム一括作成・回答集約・資料配布 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking. Applicable user instructions take precedence over skill approval ceremonies; do not stop for a formal plan approval when implementation is already authorized. Delegate only when current instructions permit it.

**Goal:** 同期対象クラス向けにひな形フォームを複製し、turret自身に回答・採点欄を準備して、Classroomへ「資料」として配布し、既存の個別返却につなげる。

**Architecture:** クラスごとの作成・配布記録をIDで永続化し、「準備」と「受付・資料公開」を別の明示操作にする。既存の設定保存・対応表・返却処理を再利用し、回答タブを同じスプレッドシートに集約する。期間切替・評価集計・既存回答IDの全面移行は行わない。

**Tech Stack:** Google Apps Script V8、FormApp、Drive／Forms／Classroom API、Google Sheets、既存HTMLダイアログ、Node.js `node:test`。新規npm依存は原則不要。

**Spec:** [引き継ぎ・要件整理](../../handoff/2026-09-29-form-provisioning.md)。確定事項と推奨案を区別して読む。

## Global Constraints

- Apps Scriptへの配置は `Code.gs`、`Administration.gs`、`Automation.gs`、`SettingsDialog.html`、`appsscript.json` の5ファイルを維持する。超える必要が出た場合は増やす前に相談する。
- 評価集計は今回対象外。高度な学期管理も先行させない。
- Classroomへの新規配布は「授業」の「資料」。既存の返却先はタイムラインへの個別投稿のまま。
- 新規作成の回答先は現在のturret自身を推奨既定値とする。既存外部スプレッドシートの参照機能は維持する。
- 回答の識別 `元SS_ID + 元シート名 + 元行番号` は維持する。新規作成記録はフォームID・コースID・投稿IDを用いる。
- 設定や配布準備を保存しただけで、資料公開、回答受付開始、個別返却の日次トリガーを開始しない。
- 学校外の回答を許可しない。回答者権限と本人メール収集を別々に検証する。検証不能な場合は公開を止め、手動確認へ案内する。
- 既存回答、採点、名簿、送信行、既存資料を無断で削除・移動・再投稿しない。部分失敗後に作成済みファイルを自動削除しない。
- 空の初期設定・旧JSON v1/v2/v3互換性を維持。今回の作成履歴・リソースID・トリガー状態を既存設定JSONへ混入させない。
- 時刻は `Asia/Tokyo`。配布機能を加えるために既存の日次処理へ行単位の外部アクセスを増やさない。
- 今回は実装計画だけを作成している。チェックボックスは未実装・未検証を示す。

## Review Focus

1. 複製・接続・資料投稿がGoogle側で成功した直後に応答や記録が失われる：自動再実行で二重作成・投稿しない（Task 2・3・5）。
2. 確認後にクラス選択、設定、フォームの回答者権限が変更される：古いプレビューを使って配布しない（Task 1・5）。
3. 自己参照・タブ名の衝突・コピーされた古い作成履歴：管理シートや別期間のフォームを処理しない（Task 2・3）。
4. 回答送信の集中、共通ロック競合、外部名簿の読取失敗：回答を失わず名前の未補完を後から回復できる（Task 4）。
5. 学校アカウント制限の未対応・公開変更による受付状態の変化：学校外へ広げず、実際の最終状態を表示する（Task 0・5）。

## 実装境界とファイル配置

| ファイル | 追加・変更する責務 |
| --- | --- |
| `Administration.gs` | 配布プレビュー、ジョブ記録、フォーム複製・接続・採点欄準備、資料公開、受付操作、管理DTO。機能ごとに見出しと非公開ヘルパーをまとめる |
| `Code.gs` | 管理タブ除外、既存設定・対応表への安全な登録、名簿照合の再利用。取り込み・返却アルゴリズムは必要箇所以外変更しない |
| `Automation.gs` | 回答先SSのフォーム送信トリガー管理、名前補完・再補完。日次トリガーと管理対象を分ける |
| `SettingsDialog.html` | クラス手順から作成ウィザード、対象確認、作成進捗、公開確認、受付状態、回答タブへのリンク |
| `appsscript.json` | Task 0で確認したDrive v3・Forms v1サービスを必要に応じて追加。既存Classroom v1を維持 |
| `tests/helpers/form-environment.mjs` | 新規ローカルテスト専用。Form/Drive/Classroom/イベントの代替。実サービスへ通信しない |
| `tests/form-*.test.mjs` | 下記各工程の振る舞い・故障試験。GAS配置には含めない |

新規の配布設定・履歴は専用の`フォーム管理`シートへ保存する案を採る。既存の`設定シート`や`Config`を再定義しない。新規作成時に同名の未知のシートがあれば上書きせず止める。

## 共通の型・API案

以下は新規実装の契約案で、既存の関数ではない。Task 0でサービス仕様に起因する変更が必要なら、この節と呼び出し側・試験を一緒に更新する。

```text
FormSetupInput = {
  templateFormId, folderId,
  titlePattern, materialTitlePattern, materialDescription, topicId?,
  responseSheetPrefix,
  targets: [{courseId, label}],
  responderPolicy: {schoolDomains: string[]},
  nameSource: {kind: 'external'|'turret', spreadsheetId?, sheetName?, emailHeader, nameHeader},
  columns: {nameHeader, gradeHeader, commentHeader, statusHeader},
  gradeChoices: string[]
}
FormSetupPreview = {
  fingerprint, configRevision, ownerSpreadsheetId,
  targets: [{courseId, className, label, formTitle, responseSheetName, materialTitle}],
  warnings: string[], blockers: string[], input: FormSetupInput
}
FormJob = {
  jobId, ownerSpreadsheetId, revision, input: FormSetupInput,
  targets: FormTarget[], createdAt, updatedAt
}
FormTarget = {
  courseId, label, formId?, responseSheetId?, responseSheetName?, materialId?,
  stage: 'pending'|'creating'|'created'|'linking'|'linked'|'prepared'|'registered'|'publishing'|'published'|'needs_review'|'error',
  operation?, lastError?, updatedAt
}
```

- クラスの正当性・選択状態はサーバーで再取得する。ブラウザーからの任意courseIdを信用しない。
- `fingerprint`は対象クラスID、入力、設定revision、回答先SSを含む正規化値から生成。ひな形の重要な設定・質問構成も再確認対象にする。
- タイトル置換は `＜class_name＞` と `＜class_label＞` など、明示したトークンのみ。任意コードは実行しない。
- 表示ラベルは自動候補を編集できる。長いClassroom名から正規表現だけで勝手に `1-1` を推測して確定しない。
- 作成履歴は専用シートの版付き列形式で記録。ジョブ単位の設定JSONとクラス別成果物IDを保持し、行位置ではなく `jobId + courseId` で照合する。任意のタイトル・説明をシートへ書く際は数式として解釈させない。
- 名簿参照設定と追加列設定はFormJobに保持し、登録済み回答タブIDから名前補完の設定を引く。既存APP_CONFIGの8,000バイト制限へ全ジョブを詰め込まない。
- 配布履歴の自動引継ぎはしない。複製先ではownerSpreadsheetId不一致を検知し、過去フォームへの変更を拒否する。設定JSONの再利用は従来機能を使う。

## Task 0: 必要な実物とGoogle側の挙動を確認する

**Files:** `docs/handoff/2026-09-29-form-provisioning.md`、`docs/development.md`。実験メモ・退避はGit管理外の`.local/`。

**Interfaces:** 不足URL・権限・対応方式を確認して、FormSetupInputと後続のサービス呼び出し条件を確定する。

- [ ] ひな形、保存先、学校ドメイン、名簿列、専用の検証クラスを確認する。資料投稿の形式は確定済みなので再質問しない。
- [ ] 公式資料でフォーム複製後の公開状態・回答者権限、`VERIFIED`メール収集、複数フォームの同一SS接続、回答タブとフォームIDの対応取得を確認する。
- [ ] 許可された検証用資源だけで、回答0件でも回答タブと追加列を準備できることを確認する。タブが見つからない場合は未完了として保持し、ダミー生徒回答を本番フォームへ入れない。
- [ ] 検証フォームへの手動回答と回答イベントを照合し、複数フォームを接続した同一SSでイベントの対象タブ・回答行を特定できることを確認する。
- [ ] API・学校側ポリシーで自動設定できない条件を記録する。許可範囲を緩める代替は採らず、確認済みでない対象の公開を禁止する。

検証対象が未提供でもTask 1の入力モデル・試験は進められる。実機を試せていない部分を完了扱いにしない。

## Task 1: 作成プレビューと入力検証

**Files:** Modify `Administration.gs`、`SettingsDialog.html` / Create `tests/form-preview.test.mjs`、`tests/helpers/form-environment.mjs`。

**Interfaces:** `previewFormSetup(input: FormSetupInput, expectedConfigRevision: string): FormSetupPreview`。読み取り専用。`getSelectedClassRecords_()`と既存revisionを利用する。

- [ ] テストを追加する：同期対象1のクラスだけを採用、対象0件拒否、重複courseId拒否、タブ名衝突・不正文字・空タイトル、ひな形／フォルダ参照不可をblockerにする。`creates === 0`、`posts === 0`、`triggerChanges === 0`を確認する。
- [ ] `node --test tests/form-preview.test.mjs`で新規試験の失敗を確認する。
- [ ] 上記APIを実装する。既存の設定・クラス選択を変えず、予定フォーム名・回答先・投稿内容・共有範囲を返す。任意gradeChoicesは空でもよく、1〜5を全利用者へ強制しない。
- [ ] 選択クラスや設定が確認後に変わるケースを追加し、fingerprint再検証で差異を検出する。
- [ ] 対象試験の成功を確認してコミットする：`feat: preview class form provisioning`。

## Task 2: 作成履歴と部分失敗からの再開

**Files:** Modify `Administration.gs` / Create `tests/form-jobs.test.mjs`。

**Interfaces:** `createFormSetupJob(input, expectedFingerprint, expectedConfigRevision): FormJob`、`getFormSetupJob(jobId): FormJob`、内部 `readFormJobs_()` / `writeFormJob_(job, expectedRevision)`。

- [ ] テストを追加する：同じ要求の二重実行でジョブが増えない、古いrevision・異なるownerSpreadsheetId・未知schemaを拒否、既存の同名ユーザーシートを保持、数式に見えるタイトルを文字列保存する。
- [ ] `node --test tests/form-jobs.test.mjs`で失敗を確認する。
- [ ] 専用シートへjobId・所有SS・要求スナップショット・クラス別ID・進捗を保存する。既存共通ロックを利用し、外部操作の前後に意図と結果を記録する。
- [ ] 外部操作の前に記録できなければ実行しない。操作後のID記録失敗、`creating/linking/publishing`で中断したジョブは、再開時に`needs_review`へ送る。
- [ ] `resume`が結果不明の対象を再実行しない試験を追加する。確認済み既存リソースを採用する内部契約 `reconcileFormTarget(jobId, courseId, resourceIds, expectedJobRevision)` は所属・フォームリンク・回答先を再検証して記録する。
- [ ] 対象試験の成功を確認してコミットする：`feat: persist resumable form provisioning jobs`。

## Task 3: フォーム複製・自己スプシ接続・採点欄・turret登録

**Files:** Modify `Administration.gs`、`Code.gs`、`appsscript.json` / Create `tests/form-provisioning.test.mjs` / Extend `tests/settings-save.test.mjs`、`tests/pipeline-500.test.mjs`。

**Interfaces:** `prepareFormTarget(jobId, courseId, expectedJobRevision): FormJob`。1回の呼び出しで1クラスを処理し、UIから順に実行する。内部 `registerPreparedFormTarget_(job, target, expectedConfigRevision)` は既存の設定・対応表へ登録する。

- [ ] 試験を追加する：フォーム複製成功、途中失敗、既知formIdの再利用、遅延した回答タブ生成、タイトルとタブ名の衝突、追加列の二重作成防止、回答・採点の保持。
- [ ] `node --test tests/form-provisioning.test.mjs`で失敗を確認する。
- [ ] ひな形を指定フォルダへ複製し、タイトルを設定。Task 0で確認した方法で回答受付を閉じた準備状態にする。原本のひな形には書かない。
- [ ] 現在のturretへ接続し、フォームと回答タブの関連を実際に照合してsheetIdを記録する。「最後のタブ」や名前だけで推測しない。未出現なら成果物を残して再確認可能にする。
- [ ] 回答開始前にタブ名・名前・評価・コメント・送信状態の追加列を準備する。質問由来の列と衝突した場合は上書きせず入力修正を求める。入力規則は指定されたgradeChoicesだけを使う。
- [ ] `FORM_SS`へ自己SSを重複なく追加し、prefix・既存設定と整合させる。対応表は作成したタブとcourseIdを登録し、既存のクラス名・メモ・他の行を保持する。既存の全消去型ヘルパーを無条件に流用しない。
- [ ] 登録前の設定revisionを確認する。設定と対応表はスナップショット・復旧を備え、失敗時には作成済みフォームを保持し`prepared`以降から再登録できるようにする。新しい列構成が既存キューと競合する場合は止める。
- [ ] 自己参照の取り込み・診断・対応表生成・名前補完が、管理用タブ（固定名とフォーム管理）を除外する試験を追加する。既存のクリア操作が回答タブを変更しないことを検証する。
- [ ] 外部SS参照と既存500件経路の回帰試験を実行し、成功後コミットする：`feat: provision forms into the turret workbook`。

既存の列・本文設定を作成ウィザードが勝手に置き換えない。未設定の役割は入力候補を提示し、通常の設定検証・保存で確定する。フォーム配布に必要な準備と、返却の自動実行に必要な準備を別に判定する。

## Task 4: 名前の自動補完と再補完

**Files:** Modify `Automation.gs`、必要箇所の`Code.gs` / Create `tests/form-name-enrichment.test.mjs` / Extend `tests/automation.test.mjs`。

**Interfaces:** `configureFormResponseTrigger(expectedRegistryRevision): {enabled, triggerId}`、イベントハンドラ `managedFormResponseReceived_(e)`、`refreshFormResponseNames(jobId, courseId, expectedJobRevision): {updated, unmatched, skipped}`。名前補完はFormJob.nameSourceを参照する。

- [ ] テストを追加する：初回回答から名前表示、メールの大小文字・空白正規化、名簿不一致、同一メールに異なる氏名、名簿取得失敗、イベント再配信、未登録タブ、不正e.range、ロック競合。
- [ ] `node --test tests/form-name-enrichment.test.mjs`で失敗を確認する。
- [ ] 回答先SSに対するインストール型フォーム送信トリガーを、管理者・IDを記録して重複なく作る。既存の日次トリガー開始／停止から誤って削除・起動されないよう管理対象を分ける。
- [ ] `e.source`、登録済みsheetId、該当行・メール・見出しを検証し、その回答の名前列だけ補完する。評価・コメント・送信状態を上書きしない。名簿は一括読み取りしてMapで照合する。
- [ ] 名簿にない場合は明示的な未照合状態にし、任意の同名生徒やClassroom表示名で勝手に補うことを避ける。失敗・未補完をログ／画面で示す。
- [ ] 「名前を更新」で登録タブをまとめて補完できるようにする。既存の手入力氏名の上書きは初期動作に含めず、空欄・未照合状態を補完する。回復経路でも採点内容を維持する。
- [ ] 複製先SSで古いトリガーIDやフォーム所有情報を流用しない試験、日次自動実行との独立性を追加する。
- [ ] 対象試験と既存automation試験の成功後コミットする：`feat: populate respondent names from the roster`。

## Task 5: 資料投稿・回答受付の切替

**Files:** Modify `Administration.gs`、`appsscript.json` / Create `tests/form-publishing.test.mjs`。

**Interfaces:** `previewFormPublication(jobId): {fingerprint, targets, blockers}`、`publishFormTarget(jobId, courseId, confirmation, expectedJobRevision): FormJob`、`setManagedFormAcceptingResponses(jobId, courseId, enabled, confirmation, expectedJobRevision): FormJob`。

- [ ] テストを追加する：正しいクラスへ`CourseWorkMaterials`を1件作成、説明文と回答用URL一致、再実行で既存materialIdを再利用、権限・内容変更による古い確認拒否、投稿結果不明を確認待ちにする。
- [ ] `node --test tests/form-publishing.test.mjs`で失敗を確認する。
- [ ] フォームの回答者権限と本人メール収集を取得し、学校外・匿名アクセスが残っていないことを確認する。編集者権限を回答者制限と混同しない。API未対応／学校ポリシーで確認不能なら公開しない。
- [ ] 確認された対象だけフォーム公開・受付開始・資料投稿を行う。フォーム公開が受付状態を変える場合を考慮し、最終状態を読み直す。日次返却トリガーは変更しない。
- [ ] 資料投稿が失敗した場合もフォーム公開・受付の実際の状態を記録する。全工程を原子的に戻せるとは扱わない。結果不明でフォームや投稿を削除しない。
- [ ] 受付停止は回答・採点・返却キュー・資料を保持する。資料が公開済みでも受付を再開／停止でき、操作対象が登録フォームだけであることを検証する。
- [ ] 対象試験の成功後コミットする：`feat: publish classroom materials and manage response intake`。

## Task 6: 既存のコンパクトな管理画面へ接続

**Files:** Modify `SettingsDialog.html`、`Administration.gs` / Extend `tests/ui-dialog-contract.test.mjs` / Create `tests/form-ui.test.mjs`。

**Interfaces:** 上記APIと `getFormProvisioningState(): {jobs, registryRevision, warnings}` を使用。追加DTOを初回の重いDrive全件検索にしない。

- [ ] UI試験を追加する：同期対象の表示、未保存設定の保持、古いpreviewの失効、連打防止、クラスごとの部分成功、確認待ちで再投稿ボタン無効、未許可公開の禁止。
- [ ] `node --test tests/form-ui.test.mjs tests/ui-dialog-contract.test.mjs`で新規試験の失敗を確認する。
- [ ] クラス手順に「ひな形からフォームを作成」を追加。入力 → 対象・命名確認 → 準備結果 → 配布確認の短い画面にする。6ステップ全体を不用意に増やさない。
- [ ] 現在のクラス一覧の選択を使い、フォームURL／回答タブ／資料リンク／受付状態／未補完件数をコンパクトな一覧で表示。各タブへの「回答・採点を開く」を用意する。
- [ ] 作成は対象ごとに順次RPCを実行し、完了分を表示する。複数の更新RPCを並列にして共通ロックと競合させない。画面を閉じても履歴から再開できる。
- [ ] 資料公開と受付変更は対象・内容を確認する。準備時の保存ボタンや上部「変更を保存」に外部公開を混在させない。
- [ ] 準備した採点用列を既存の「列と項目」の候補へ引き渡す。未設定の返却本文などを残したままでも配布準備できる一方、返却の自動実行は既存の必須条件を維持する。
- [ ] 対象試験と画面密度・スクロール・下書き保持を確認し、コミットする：`feat: guide form setup and distribution in the console`。

## Task 7: 全体確認・文書・確認用への反映

**Files:** Modify `README.md`、`docs/development.md`、この計画・引き継ぎ文書。必要に応じて関連試験を補う。

- [ ] `pnpm test`を実行し、既存249件と追加試験がすべて成功することを確認する。件数は追加後の実数を記録する。
- [ ] 実コードの構文、`git diff --check`、設定JSON互換性、キュー保持、5ファイルだけの同期対象、秘密情報混入なしを確認する。
- [ ] 自己レビューで新規資源IDと旧回答行キーが混同されていないこと、別SSコピーの旧フォーム操作防止、受付停止と返却停止の分離を確認する。
- [ ] 許可済み確認用Apps Scriptを別ディレクトリへ退避し、クラウド側独自変更を比較する。同期先を再確認し、`pnpm gas:push`を実行する。
- [ ] 新しいGoogle権限が必要なら、要求内容を示して本人に承認してもらう。既存承認を理由に自動で認証操作しない。
- [ ] 許可された専用環境で、回答0件の準備 → 手動テスト回答 → 名前補完 → 手動採点 → 既存取り込み・本文準備を確認する。資料公開・個別返却を実際に試す場合は各対象への実行許可を確かめる。
- [ ] 受付停止後に新規回答ができず、既存の採点・返却は継続できることを確認する。再開・画面再読込でフォームや資料が増殖しないことを確認する。
- [ ] 同一SSと既存外部SSの双方を検証し、測定していない性能や送信を成功扱いにしない。ブラウザー制約などの未確認範囲を記録する。
- [ ] README・検証記録・残課題を更新し、変更を論理単位でコミット。現ブランチへ通常プッシュし、SHA・送信先・成否を報告する。保護ブランチへ直接プッシュしない。

## 完了条件

- ユーザーがひな形と対象クラスを指定すれば、回答開始前にクラス別フォームと採点用タブを準備できる。
- 人が回答開始後にXLOOKUPを貼る作業が不要になり、名前補完失敗にも回復手段がある。
- 「授業」の資料として正しいクラスへ配布でき、フォームごとの回答受付を切り替えられる。
- 作成結果からturretの参照元・クラス対応を登録でき、従来の採点・個別返却へつながる。
- 部分失敗・再実行・別期間コピーで、既存データの消失や自動二重投稿を起こさない。
- 実機未検証の条件があれば明示する。評価集計・学期コンソール・設定シート統合・回答ID全面移行を完成条件へ追加しない。

## 計画の自己レビュー記録

- [x] ユーザーの「資料」指定、集約案、1選択クラス、名前補完、評価集計後回しを反映。
- [x] 資料配布記録のID管理と既存回答の行番号管理を区別。
- [x] 各APIの入力・戻り値・配置先と、失敗ケースを担当Taskへ割当。
- [x] 学校側ポリシー、名簿参照先、検証用資源を未確認として明記。
- [x] 5ファイル制約、既存UI密度、外部SS互換性、実機操作の許可範囲を維持。
- [x] 計画のみ作成し、フォーム作成・権限変更・投稿・GAS同期は実行していない。
