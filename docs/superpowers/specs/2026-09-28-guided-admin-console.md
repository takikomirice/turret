# Guided administration console

## Intent

日常処理は既存トリガーで自動運転し、年に数回触る利用者が初見でも設定・保守できる画面にする。操作順は初期設定、保守・クリア、手動実行。名簿と対応表はシート編集を活かし、参照元・通知先・列・項目・文面はダイアログで編集する。

## Agreed behavior

- 新規設定は用途固有の列・項目・文面を持たない。旧版の保存済み設定は維持する。空欄や空配列を旧初期値で復活させない。
- 設定手順はシート準備、クラス選択、生徒取得、フォームと通知先、対応表、列と項目、本文、自動実行。進捗は現在の設定・シートから算出し、未確認を完了と表示しない。
- 入力は手順ごとに下書き保存できる。実行と自動実行開始は完全な設定を検証する。既存データを消す保存は行わない。列変更が既存キューと互換でない場合は変更を止め、保守手順を案内する。
- シート入力とダイアログは同じ設定を使う。画面表示後の変更をrevisionで検知して上書きを止める。
- JSON設定の書出し/読込、適用前の確認を提供する。文面・列・項目・時間帯を再利用し、参照元と通知先は書出し時に選択可能。名簿、評価、送信履歴、認証情報、トリガーIDは含めない。読み込みでトリガーを作らない/有効にしない。
- 埋め込み候補には日本語の意味、参照列、例を表示。プレビューは架空の値と明示し実送信しない。
- 名簿は全取得成功後に置換。クラス更新で同期対象を保持し全ページを取得。予約済み変数との衝突、不正入力を保存前に拒否する。
- 自動実行の時刻はダイアログで変更可。初期値は毎日Asia/Tokyoの取り込み5時台、準備と送信8時台、リマインダー16時台。時刻はGoogleの時間帯指定で、分単位の定刻保証ではない。
- 準備と送信は一つの管理用トリガーで順番に実行する。旧公開関数は維持する。開始・変更・停止は明示操作。開く/保存/インポートだけではトリガーを変更しない。
- 自分の管理対象トリガーだけを変更する。他の用途のトリガーは保持。旧送信処理のトリガーは確認して切り替える。他アカウントのトリガーは列挙できないため、単一管理者の運用を案内する。
- 既存の送信checkpoint、重複防止、共通ロックを維持。UIの警告は影響と次の操作を具体的に示す。

## Boundaries

本番反映・実送信・実トリガーの作成は検証で行わない。確認用Apps Scriptへの同期と送信を伴わない画面確認は実施する。GAS管理画面内で認証が必要な場合は本人が承認する。元回答行の移動をまたぐ永続ID導入は別改修。

## Components and contracts

- Code.gs: existing pipeline, draft settings save/revision, transactional rollback, roster safety, validation and migration.
- Administration.gs: console state/progress, whitelisted actions, safe sheet navigation, preview, JSON profiles, menu entry screens.
- Automation.gs: time settings, current-owner trigger lifecycle, ordered delivery handler, result status.
- SettingsDialog.html: dependency-free Japanese guided console with setup, maintenance, manual actions and profiles.

Public UI calls:

- getAdminConsoleData() -> {config,revision,progress:[{id,title,state,detail}],counts,sheets,initialPanel}; states complete/pending/attention. Progress ids prepare/classes/students/sources/mapping/fields/template/automation.
- saveSetupSection(section,payload,expectedRevision) -> {ok,message,config,revision}; sections sources/fields/template; fields section includes column roles and fields.
- runAdminAction(action,confirmed,expectedRevision) -> {message,data}; actions initialize/classes/students/mapping/import/prepare/send/remind/clearEval/clearSend/clearErrors. Send/remind/clear need explicit confirmation. Never accepts arbitrary handler names.
- openAdminSheet(key) -> {message}; keys classes/students/mapping/settings/evaluation/send/errors.
- fetchHeadersFromSources() -> {headers,warnings}; existing behavior.
- previewReplyTemplate(payload) -> {text,tokens:[{key,label,source,example}],warnings}; actual server renderer with fictional values.
- exportSettingsProfile(includeConnections) -> {filename,json}; includes desired automation times, never enabled state.
- inspectSettingsProfile(json) -> {summary,config,schedule,warnings}; no writes.
- applySettingsProfile(json,expectedRevision) -> {message,data}; explicit application, preserves queues or rejects incompatibility, does not change automation.
- getAutomationState() -> {schedule:{importHour,deliveryHour,reminderHour,reminderEnabled},timezone,enabled,legacyCount,triggers,warnings,revision,lastRuns}.
- configureAutomation(schedule,expectedRevision,replaceLegacy,expectedConfigRevision) -> automation state. Saves schedule and enables/updates after explicit confirmation; validates configuration and the viewed config revision first.
- stopAutomation(expectedRevision,includeLegacy) -> automation state. Only relevant current-user triggers.
- getAutomationSchedule_() -> schedule; safe read without trigger API for profile export. saveAutomationScheduleDraft_(schedule) validates and saves desired times without touching active triggers (import warns active schedule unchanged).

Private cross-component APIs: getConfigRevision_(config), applyConfigDraft_(config,expectedRevision) called under common lock. Automation uses existing unlocked pipeline workers, validates getConfig_(). All UI mutations serialize with withAppLock_.
