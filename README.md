# Google Calendar → Notion 自動同期スクリプト

Google Apps Script と Notion API を使って、
Googleカレンダーの予定をNotionデータベースに自動反映するスクリプトです。

## 機能
- GoogleカレンダーのイベントをNotionに同期
- 削除された予定は自動でアーカイブ
- 毎日トリガーで自動実行（無料）

## 使い方
1. `Code.gs` の中身を Apps Script エディタに貼り付け
2. `CONF` のトークンとデータベースIDを設定
3. 手動で `runRollingSync()` を実行して動作確認
4. トリガーで自動化

## 必要環境
- Googleアカウント
- Notion Integration Token
- Notion Database ID

## ライセンス
MIT
