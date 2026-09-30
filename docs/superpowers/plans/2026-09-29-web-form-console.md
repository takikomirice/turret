# Web form console implementation plan

> **For agentic workers:** Use superpowers:executing-plans. User authorized implementation; no formal approval pause.

**Goal:** Web画面からフォーム準備・配布・受付管理・記録済み投稿削除を行う。
**Architecture:** 1プロジェクト1運用。共通getAppSpreadsheet_で接続先を固定、既存画面を共用。フォーム管理シートの版付き記録で外部操作を追跡。
**Tech Stack:** GAS, Forms/Drive/Classroom advanced services, HTML, node:test.
**Spec:** ../specs/2026-09-29-web-form-console.md

## Global Constraints
- 配置はCode.gs、Administration.gs、Automation.gs、SettingsDialog.html、appsscript.jsonの5ファイル。
- 生徒一覧はClassroom。科目別運用。回答IDは元SS＋タブ名＋行を維持。
- 本番や旧投稿を自動で変更しない。公開・削除は具体的な確認が必要。

## Review Focus
- Webで別運用へアクセスしない、コピーで旧接続を継承しない（Task 1）。
- 外部成功後の書込失敗を自動再作成・再投稿しない（Task 2）。
- 回答者権限とメール収集を公開直前にも照合（Task 2）。
- 名簿未照合／重複で行を消さず、採点を変更しない（Task 3）。
- 削除の結果不明／再実行で回答や送信状態を変更しない（Task 4）。

## Tasks
- [x] 1. Code.gsのgetAppSpreadsheet_、Administration.gsのWeb起動・新規作成・複製。tests/web-console.test.mjsを失敗→成功。既存テストを確認。
- [x] 2. Administration.gsにgetFormConsoleData、previewFormSetup、prepareFormTarget、previewManagedFormAction、runManagedFormAction。フォーム管理シートで進捗保存。tests/form-console.test.mjsを失敗→成功。
- [x] 3. Automation.gsのmanagedFormResponseReceived_、登録トリガー、名前再照合。新規列は既存フィールド設定へ候補提供し、既存キューを変えない。
- [x] 4. 資料と新規返却投稿のIDを記録し、プレビュー確認後削除。Code.gs送信成功直後のID保存と送信確認待ちを維持。削除／再実行試験。
- [x] 5. SettingsDialog.htmlでWeb起動、フォーム一覧・入力・進捗・確認操作。既存設定UIの回帰と構文確認。
- [ ] 6. 全テスト、独立レビュー、README・導入手順・改修レポート更新。確認用同期前バックアップ・差分照合。実機検証可能範囲を確認。ブランチへコミット・プッシュ。

## Execution ledger
- Ruling: 学校環境の権限制限はユーザー確認済み。追加機能の実動作だけ実機で確認する。
- Ruling: 任意の外部投稿は削除しない。投稿IDのない既存返却は一覧の対象外。
- Ruling: Webの複数運用同時制御を今回に混在させず、運用単位の分離を維持する。

- 実行記録：Web起動と既存設定読込まで実機確認。フォーム・投稿の実機試験は検証資源と実行許可の回答待ち。
- レビュー修正：名前補完の取り込み時追補、結果不明時の受付停止、コピー初期化の永続チェックポイントを追加し回帰試験。
