# SharePoint ライブラリ評価カスタムWebパーツ

SharePoint ドキュメントライブラリの評価（レーティング）を閲覧・更新できる、SPFx 想定のカスタムWebパーツ実装です。

## 実装内容

- ライブラリ内アイテムの評価情報（平均値 / 件数）を一覧表示
- `Rating` UI からユーザー評価を更新
- Web パーツ設定で以下を変更可能
  - 対象ライブラリ名
  - 取得件数
  - 自分が評価したアイテムのみ表示

## 主なファイル

- `src/webparts/libraryRatings/LibraryRatingsWebPart.ts`
  - Webパーツ本体 / プロパティペイン定義
- `src/webparts/libraryRatings/components/LibraryRatings.tsx`
  - 一覧表示と評価操作 UI
- `src/webparts/libraryRatings/services/SharePointRatingsService.ts`
  - SharePoint REST API 呼び出し（一覧取得・評価更新）

## REST API

- 一覧取得:
  `/_api/web/lists/getByTitle('<library>')/items?$select=Id,Title,FileLeafRef,FileRef,AverageRating,RatingCount,Modified`
- 評価更新:
  `/_api/web/lists/getByTitle('<library>')/items(<id>)/SetRating(rating=<0-5>)`

## 利用時の注意

- 対象ライブラリで評価機能が有効になっている必要があります。
- SPFx プロジェクトへの組み込み時は、既存の `serve/build/package-solution` などの設定に合わせてファイルを統合してください。
- 必要に応じてアクセス権（読み取り/編集）を確認してください。
