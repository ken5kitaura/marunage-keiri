# 顧客管理GAS

LINE顧客管理スプレッドシート用のGASです。

## 機能

### 新規顧客セットアップ
- `📁 新規フォルダ50件作成（一括）` - フォルダ作成 → スプシ登録 → 権限付与を一括実行

### 個別機能
- `📁 フォルダのみ作成` - フォルダのみ作成
- `📝 スプシに登録` - 既存フォルダをシートに登録
- `🔑 権限付与（全フォルダ）` - 全フォルダにサービスアカウント権限を付与
- `🔄 既存顧客にサブフォルダ追加` - 既存顧客に通帳・クレカ明細フォルダを追加

## 作成されるフォルダ構成

```
MK001/
├── 領収書/
├── 通帳/
├── クレカ明細/
└── MK001_レシート読込（スプシ）
    └── ClientConfigシート
        ├── PASSBOOK_FOLDER_ID
        ├── RECEIPT_FOLDER_ID
        └── CC_STATEMENT_FOLDER_ID
```

## 顧客管理シートの列構成

| 列 | 内容 |
|----|------|
| A | line_user_id |
| B | customer_name |
| C | folder_id（領収書フォルダID） |
| D | registered_at |
| E | notified |
| F | sent_at |
| G | customer_code |
| H | status |
| I | email |
| R | spreadsheet_url |
| S | passbook_folder_id |
| T | cc_statement_folder_id |

## 設置方法

1. 顧客管理スプレッドシートのGASエディタを開く
2. `CustomerManagement.gs` の内容を貼り付け
3. スクリプトプロパティに `LINE_ACCESS_TOKEN` を設定

## 注意

このファイルはローカル管理用です。
clasp pushでReceiptEngineにはアップロードされません。
