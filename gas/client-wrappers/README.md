# クライアントラッパースクリプト

各顧客スプレッドシートのGASに貼り付けるラッパースクリプトを管理するフォルダです。

## ファイル構成

```
client-wrappers/
├── _TEMPLATE.gs    # 新規顧客用テンプレート
├── MK004.gs        # MK004用
└── README.md       # このファイル
```

## 新規顧客を追加する手順

1. `_TEMPLATE.gs` をコピーして `MKxxx.gs` を作成
2. `CLIENT_CONFIG` の値を編集
   - `PASSBOOK_FOLDER_ID`: Google Driveの通帳フォルダID
3. 顧客スプシのGASエディタを開く
4. ライブラリ「ReceiptEngine」を追加（識別子: ReceiptEngine）
5. 作成した `MKxxx.gs` の内容を貼り付け
6. 保存して動作確認

## ライブラリの追加方法

1. 顧客スプシ → 拡張機能 → Apps Script
2. 左側「ライブラリ」の「＋」をクリック
3. スクリプトID（ReceiptEngineのプロジェクトID）を入力
4. 「検索」をクリック
5. バージョンを選択、識別子は `ReceiptEngine`
6. 「追加」をクリック

## 注意事項

- ラッパースクリプトはGASエディタに直接貼り付けて使用
- claspでの自動デプロイは対象外（各顧客スプシは独立）
- 変更があった場合は手動で各顧客スプシに反映が必要
