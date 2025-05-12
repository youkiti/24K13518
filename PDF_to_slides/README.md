# PDF to Slides 変換システム（Claude API版）

## 概要

このGoogle Apps Scriptは、Google Drive内のPDFファイルを自動的に検出し、以下の処理を実行します：

1. Google Drive内の指定フォルダからPDFファイルを検索
2. 処理済みかどうかをスプレッドシートで確認
3. Claude APIを使用してPDFを直接解析し、必要な情報を抽出
   - タイトルと日本語タイトル
   - 雑誌名（Journal）
   - 引用情報（Citation）
   - Limitation
   - Clinical/Research/Policy Implication
   - 論文の要約
4. 指定のGoogleスライドテンプレートを複製し、抽出した情報で置換
5. 生成したスライドの管理と処理状態の追跡

## 前提条件

- Google Workspace（Googleドライブ、スプレッドシート、スライド）へのアクセス権
- Claude API（APIキー）
  - [Anthropic Developer Console](https://console.anthropic.com/)でアカウント作成とAPIキー取得が必要
- Google Apps Scriptエディタへのアクセス
- スライドテンプレート（ID: 1Z4eV6RhyBrMHtRUxPpNRNPa2IG61cK3lAq84zm1muiA）へのアクセス権

## システム構成

このシステムは2つの主要なファイルで構成されています：

1. **ClaudeAPIHandler.gs** - Claude APIを使用してPDFを解析する機能
2. **SlideCreator.gs** - 解析結果からGoogleスライドを生成する機能

## セットアップ手順

### 1. Google Apps Script プロジェクトへのファイルのインポート

1. [Google Apps Script](https://script.google.com/) にアクセス
2. 「新しいプロジェクト」をクリック
3. このリポジトリの `ClaudeAPIHandler.gs` と `SlideCreator.gs` の内容をコピー＆ペースト

### 2. 必要なAPIサービスの有効化

スクリプトエディタのメニューから「サービス」を追加：

1. Google Drive API
2. Sheets API
3. Slides API

### 3. スクリプトプロパティの設定

1. スクリプトエディタで「プロジェクトの設定」→「スクリプトプロパティ」を開く
2. 以下のプロパティを設定：

| プロパティ名 | 説明 |
|------------|------|
| PDF_FOLDER_ID | PDFファイルが保存されているGoogleドライブフォルダのID |
| TRACKING_SPREADSHEET_ID | 処理状態を管理するスプレッドシートのID |
| CLAUDE_API_KEY | Claude APIキー |
| SLIDE_OUTPUT_FOLDER_ID | 生成したスライドを保存するフォルダID |

※IDはGoogleドライブ要素のURLから取得できます：
- フォルダID: `https://drive.google.com/drive/folders/FOLDER_ID` 
- スプレッドシートID: `https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit`

### 4. トリガーの設定

自動実行を設定するには：

1. `SlideCreator.gs` の `setupTrigger` 関数を実行して毎日の自動処理をセットアップ

## 使い方

### 1. PDF準備

1. 対象のPDFファイルを、設定した `PDF_FOLDER_ID` のGoogleドライブフォルダにアップロード

### 2. 処理の実行

トリガーを設定した場合：
- 毎日午前9時に自動的に処理が実行されます

手動で実行する場合：
- スプレッドシートに処理対象PDFを登録して `runManually` 関数を実行
- または特定フォルダ内の全PDFを処理する `processAllPDFsInFolder` 関数を実行

### 3. 結果確認

1. 処理状態は設定したスプレッドシートの「PDF処理状態」シートで確認できます
2. 生成されたスライドは `SLIDE_OUTPUT_FOLDER_ID` で指定したフォルダに保存されます

## トラブルシューティング

### 一般的な問題

- **権限エラー**: スクリプト初回実行時に必要な権限をすべて許可してください
- **実行時間制限**: 一度に処理するファイル数が多い場合、Google Apps Scriptの実行時間制限（6分）に達する可能性があります。処理するファイル数は `processAllPDFsInFolder` 関数内で制限されています。

### Claude API関連の問題

- **APIキーエラー**: Claude APIキーの有効性を確認してください
- **サイズ制限**: 非常に大きなPDFファイル（32MB以上または100ページ以上）は処理できません
- **レート制限**: Claude APIの利用制限に注意してください

## ログの確認

エラーやデバッグ情報は、スクリプトエディタの「実行」→「ログを表示」で確認できます。

## 注意事項

- Claude APIは有料サービスです。使用量と料金にご注意ください。
- テンプレートスライド(ID: 1Z4eV6RhyBrMHtRUxPpNRNPa2IG61cK3lAq84zm1muiA)に以下のプレースホルダーが必要です：
  - `{Title}` - 論文原題
  - `{Japanese_Title}` - 日本語タイトル
  - `{Journal}` - 雑誌名
  - `{Citation}` - 引用情報
  - `{Limitation}` - 研究の限界
  - `{ClinicalImplication}` - 臨床的意義
  - `{ResearchImplication}` - 研究的意義
  - `{PolicyImplication}` - 政策的意義
  - `{Summary}` - 要約
