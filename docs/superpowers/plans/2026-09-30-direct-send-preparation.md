# 回答から送信シートを直接作成する実装計画

> **For agentic workers:** Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** 評価データを通常処理から外し、送信データの準備と送信の2段階にする。
**Architecture:** 採点済み回答から必要項目を変換し、既存の宛先解決・本文生成を通して送信シートへ保存する。旧評価データは移行元として保持し、未処理行を優先して引き継ぐ。送信のチェックポイントと状態管理は維持する。
**Tech Stack:** Apps Script V8、Google Sheets / Classroom、node:test。
**Spec:** この会話で承認された「準備→送信」、作成済み本文を自動更新しない設計。

## 制約と判断

- GAS配置は既存5ファイル。外部回答スプシの既存設定も利用可能とする。
- 現作業ブランチで実装する。作業開始時の未コミット差分はなし。
- 旧評価データは自動削除しない。新規作成・新規追記はしない。既存の未処理分を引き継ぐ。
- `importHour` と既存ハンドラー名・設定JSON形式は互換用に保持し、意味を送信データ準備へ変更する。送信時間には準備を再実行しない。
- 旧関数名も準備処理への互換入口として保持する。
- 作成済み本文は再作成せず、済・確認待ち・エラーを含む全送信行を重複判定に使う。

## Review Focus

- 旧評価データと元回答の重複時に旧取り込み値を優先する。
- 保存後の応答消失・flush失敗・状態更新失敗で二重作成しない。
- 旧準備済み/準備失敗行を意図せず新規回答として再処理しない。
- 旧評価データの完了済み履歴が新しい設定変更を不要に妨げない。
- 準備後の採点/本文変更は送信待ち本文へ自動反映せず、重複回答の保留は送信時にも確認する。

## Task 1: 準備処理と移行

Files: `Code.gs`, `tests/direct-send-preparation.test.mjs`, 関連pipelineテスト。
Interfaces: `prepareSendData()` / `prepareSendDataUnlocked_()`、既存送信シート形式。

- [x] 新規準備・移行・失敗再試行のテストを追加し、未実装で失敗することを確認。
- [x] 旧評価データの読み取りと回答からの候補作成、共通の宛先解決・本文生成を実装。
- [x] `node --test tests/direct-send-preparation.test.mjs tests/pipeline-reliability.test.mjs tests/pipeline-500.test.mjs` が成功することを確認。

## Task 2: 自動実行・画面・診断

Files: `Administration.gs`, `Setting.html`, 管理・自動実行・UIテスト。
Interfaces: Task 1の準備処理、既存のschedule/config形式。

- [x] 準備時刻には準備のみ、送信時刻には送信のみのテストへ更新。
- [x] 通常画面を2段階へ変更し、旧データがある場合だけ移行案内・管理操作を表示。
- [x] 診断、設定保存の列検証、初期化で評価データを不要にする。
- [x] 関連テストと`node --test`（`pnpm test`の実体）が成功することを確認。pnpmは起動時の取得がネットワーク制限で失敗したため直接実行。

## Task 3: 実機検証・レビュー・配送

Files: `README.md`, `docs/development.md`、検証記録。

- [x] READMEに流れ・時刻・移行・作成済み本文の扱いを反映。
- [x] 差分を独立レビューし、指摘を修正して必要なテストを実行。
- [x] 確認用GASに5ファイルを同期し、本人限定画面とテストクラスで準備→個別送信を確認。実行できない範囲は理由を記録。
- [x] 最終5ファイルの読み戻し・一致確認、本人限定のバージョン37の画面確認、差分確認を完了。
- 配送先は`feat/scoring-tool-integration`。コミット・プッシュの結果は完了報告に記載する。
