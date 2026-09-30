# Guided administration console implementation plan

**Goal:** 初期設定・保守・自動実行を案内付き画面へ統合し、空の新規設定と再利用可能なJSONを提供する。

**Architecture:** 既存パイプラインの公開入口と安全対策を保ち、管理・自動実行は専用GASファイルへ分ける。UIは同じ設定を読み書きし、revisionで古い画面からの上書きを防ぐ。

**Tech Stack:** Apps Script V8, dependency-free HTML/CSS/JS, Node test runner, clasp.

**Spec:** ../specs/2026-09-28-guided-admin-console.md

## Constraints

- 生徒・メールへの試験送信、実トリガー作成、本番反映をしない。
- 現在の公開関数、シート列、旧設定の読み込みを維持する。
- 各担当は所有ファイルだけを編集する。rootが統合し検証・コミット・プッシュする。

## Tasks

- [x] Backend: Code.gsと設定/名簿テスト。失敗テスト→保存によるデータ削除廃止、互換性検証、保存失敗の復旧、revision、空初期値と旧設定移行、予約変数と空配列、全ページ名簿取得→関連テスト成功。
- [x] Automation: Automation.gsと単体テスト。失敗テスト→時刻検証、現在の管理者のトリガー管理、既存トリガーの明示切替、部分失敗/孤児トリガー防止、順次準備送信、開始停止→成功。
- [x] Console: SettingsDialog.htmlとUIテスト。手順ナビ、設定入力、本文プレビュー、JSON、時刻変更、確認操作、通信失敗、未保存警告、アクセシビリティと狭い画面→成功。
- [x] Integration: Administration.gs、管理/JSONテスト。進捗算出、操作allowlist、確認境界、型付きプロファイル、読込だけでは書換しない、既存値保持→成功。
- [x] Verification: 全テスト・構文・差分・独立レビュー、確認用コード退避/同期、実機UI確認。README/開発手順更新。
- [x] Delivery: 今回の差分のみfeat/guided-admin-consoleにコミットしturretへプッシュ。

## Review focus

- 不完全な設定の下書きは保存できても、自動送信開始は拒否する。
- JSONや古いダイアログから旧URL/通知先/本文が意図せず復活しない。
- インポート/保存はキューを消さず、トリガーを勝手に作らない。
- トリガー切替失敗で重複実行や別用途のトリガー削除を起こさない。
- 既存v3設定と新規の空設定を区別し、空配列を初期値で補わない。

## Verification evidence

- 162 tests passed; independent review findings resolved.
- Five files synced to the test Apps Script after remote backup and comparison.
- Real UI: menus, progress, empty values, draft retention, fictional preview, editable hours, confirmation, initialization and empty settings save verified.
- Browser JSON file selection timed out in the extension; browser import end-to-end and downloaded file contents remain unverified. Server and UI profile logic passed local tests.
- No production sync, Classroom/Gmail sends, live roster sync or live trigger creation/firing. See docs/development.md for details.
