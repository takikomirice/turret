# 設定JSON画面の整理 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 設定JSONの出力と入力を2列1行にまとめ、出力は種類選択、入力は種類の自動判定にする。

**Architecture:** 既存5配置ファイルのまま、Administration.gsでJSONの識別と既存検証への接続を行う。Setting.htmlは出力プルダウンと共通入力欄を横並びにし、狭い画面では縦に並べる。確認済み種類と競合検知を使う適用処理は維持する。

**Tech Stack:** GAS V8、HTML/CSS、Node.js node:test。

**Spec:** 本会話の「出力する種類はプルダウン」「入力は自動認識」「出力と入力を2列1行」の指定。

## Global Constraints

- 配置対象は既存5ファイルを維持し、追加API・Cloud設定を必要にしない。
- 既存の個別JSON・一括JSONと適用前確認、競合検知、失敗時復元を維持する。
- 作業ブランチ feat/configuration-json-bundle で実装・検証・レビュー後にコミットとプッシュする。

## Review Focus

- 不明・矛盾する識別情報を採点設定として誤認しない。
- 旧管理設定とラッパーなし採点JSONを引き続き認識する。
- 検出後も管理設定の128KiB上限を適用する。
- 別ファイル選択や不正入力で古い適用候補を残さない。
- 出力種類変更で読み込み候補を変えず、狭い画面でも操作できる。

### Task 1: 種類の自動判定

**Files:** Administration.gs、tests/configuration-files.test.mjs。

**Interfaces:** inspectConfigurationFile('auto', json) は kind を含む確認結果を返す。既存の明示種類・applyConfigurationFile の契約は維持する。

- [x] 4種類・旧形式・不明形式・矛盾する種類・サイズ制限のテストを追加し、失敗を確認。
- [x] 保存形式を変更せず識別し、既存の解析と検査へ接続。
- [x] 関連テストの成功を確認。

### Task 2: 2列の入出力UIと実機確認

**Files:** Setting.html、tests/ui-dialog-contract.test.mjs、README.md、docs/development.md。

**Interfaces:** 出力選択を state.exportKind に保持し、共通ファイル欄で inspectConfigurationFile('auto', json) を呼び、返された kind を適用候補に保存する。

- [x] 1個の種類選択・出力ボタン・入力欄、種類保持・自動検出・無効入力のUIテストを追加し、失敗を確認。
- [x] 2列の出力・入力と狭幅の1列表示を実装し、利用手順を更新。
- [x] 全642テスト、構文、差分を確認し、確認用GASの5ファイルを読み戻して一致を確認。Chrome拡張で2列画面・種類切替・一括生成通知を確認。
- [x] 自己レビューと独立レビューを完了。今回変更を同ブランチへコミット・プッシュする。

## 実機確認の制約

ファイル選択がタイムアウトし、拡張のファイルURLアクセス設定を案内。実ファイル保存完了と選択から適用までの実機結合は未確認。確認用Webデプロイはバージョン43。本番反映や設定適用は行っていない。
