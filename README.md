# Google広告スクリプト セットアップガイド

## 概要
このリポジトリには、Google広告のレポートを自動取得し、Googleスプレッドシートに出力するスクリプトが含まれています。

## テンプレートスプレッドシート
以下のテンプレートスプレッドシートを使用して、レポートの出力先として設定します：

[広告レポート_テンプレート](https://docs.google.com/spreadsheets/d/1MBdF2FHNYp77P-MCGQxF7l5ZHbgtbzyYaVVVKJTQDJ8/edit?usp=sharing)

## スクリプト一覧
各スクリプトのGitHubリンクは以下の通りです：

- [daily.js](https://github.com/ShogoKumazaki/Ads_script/blob/main/daily.js) - 日次レポート
- [prefectures.js](https://github.com/ShogoKumazaki/Ads_script/blob/main/prefectures.js) - 都道府県別レポート
- [timeSlot.js](https://github.com/ShogoKumazaki/Ads_script/blob/main/timeSlot.js) - 時間帯別レポート
- [query.js](https://github.com/ShogoKumazaki/Ads_script/blob/main/query.js) - 検索クエリレポート
- [keyword.js](https://github.com/ShogoKumazaki/Ads_script/blob/main/keyword.js) - キーワードレポート
- [conversions_action_daily.js](https://github.com/ShogoKumazaki/Ads_script/blob/main/conversions_action_daily.js) - コンバージョンアクション別レポート

## セットアップ手順

### 1. スプレッドシートの準備
1. テンプレートスプレッドシートをコピーして新しいスプレッドシートを作成
2. スプレッドシートのIDをコピー（URLの `/d/` と `/edit` の間の文字列）

### 2. スクリプトの設定
各スクリプトで以下の手順を実行します：

1. Google広告のスクリプトエディタを開く
2. 新しいスクリプトを作成
3. 対応するスクリプトファイルの内容をコピー＆ペースト
4. `SPREADSHEET_ID`の値を、コピーしたスプレッドシートのIDに変更

```javascript
// 例：
const SPREADSHEET_ID = '1MBdF2FHNYp77P-MCGQxF7l5ZHbgtbzyYaVVVKJTQDJ8';
```

### 3. スクリプトの実行設定
1. スクリプトエディタで「トリガー」を設定
2. 実行する関数として`main`を選択
3. 実行頻度を設定（推奨：毎日1回）

## 各スクリプトの説明

### daily.js
- 日次の広告パフォーマンスデータを取得
- キャンペーン、広告グループ、コンバージョンアクションのデータを出力

### prefectures.js
- 都道府県別の広告パフォーマンスデータを取得
- 過去3ヶ月分のデータを出力

### timeSlot.js
- 時間帯別の広告パフォーマンスデータを取得
- 過去3ヶ月分のデータを出力

### query.js
- 検索クエリのパフォーマンスデータを取得
- 昨日、過去7日、過去30日、過去90日のデータを出力

### keyword.js
- キーワードのパフォーマンスデータを取得
- 昨日、過去7日、過去30日、過去90日のデータを出力

### conversions_action_daily.js
- コンバージョンアクション別のパフォーマンスデータを取得
- キャンペーンと広告グループの両方のレベルでデータを出力
- 過去3ヶ月分のデータを出力

## 注意事項
- スクリプトの実行には、Google広告アカウントへの適切なアクセス権限が必要です
- スプレッドシートへの書き込み権限が必要です
- 大量のデータを取得する場合は、実行時間制限に注意してください

## トラブルシューティング
- スクリプトが実行されない場合は、トリガーの設定を確認してください
- データが出力されない場合は、スプレッドシートIDが正しいか確認してください
- エラーが発生した場合は、Google広告アカウントの権限を確認してください
