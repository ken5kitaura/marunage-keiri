# LINE領収書システム 運用マニュアル

## 概要

顧客がLINEで領収書の写真を送ると、自動的にGoogle Driveの顧客別フォルダに保存されるシステム。

---

## システム構成

```
顧客のLINE
    ↓ 写真送信
LINE Messaging API
    ↓ Webhook
Google Cloud Run (line-receipt-webhook)
    ↓ 顧客フォルダ確認
Google Sheets (LINE顧客管理)
    ↓ 画像転送
Google Apps Script (LINE領収書アップロード)
    ↓ 保存
Google Drive (顧客別フォルダ)
```

---

## 各サービスへのアクセス

### LINE Official Account Manager
- URL: https://manager.line.biz/
- アカウント: まるなげ経理 (@095ylsrv)
- 用途: 顧客とのチャット、Webhook設定

### LINE Developers Console
- URL: https://developers.line.biz/
- 用途: チャネルアクセストークンの管理

### Google Cloud Console
- URL: https://console.cloud.google.com/
- プロジェクト: marunage-keiri
- サービス: line-receipt-webhook
- リージョン: asia-northeast2 (大阪)

### LINE顧客管理スプレッドシート
- URL: https://docs.google.com/spreadsheets/d/1sUqtfRUSjg4XAAymuxCJm8masTkmxBVle70rCGxglXU/edit
- 用途: LINEユーザーIDと顧客フォルダの紐付け

### Google Apps Script (LINE領収書アップロード)
- URL: https://script.google.com/macros/s/AKfycbywXPJFZyUJgp-v4-n10DVpjwASomRMoeWzQRiuiDeVXIVXhKtEVMRytrkPPQImfxoVfA/exec
- 用途: Google Driveへの画像アップロード

### Google Apps Script (顧客通知)
- 場所: LINE顧客管理スプレッドシート → 拡張機能 → Apps Script
- 用途: 設定完了時の自動LINE通知

---

## 環境変数 (Cloud Run)

| 変数名 | 用途 |
|--------|------|
| LINE_CHANNEL_SECRET | LINE署名検証用 |
| LINE_CHANNEL_ACCESS_TOKEN | LINE API認証用 |
| CUSTOMER_SHEET_ID | 顧客管理スプレッドシートID |
| GAS_UPLOAD_URL | GASのウェブアプリURL |

---

## スプレッドシートの構成

| 列 | ヘッダー | 内容 |
|----|----------|------|
| A | line_user_id | 顧客のLINEユーザーID（自動登録） |
| B | customer_name | 顧客名（手動入力） |
| C | folder_id | Google DriveのフォルダID（手動入力） |
| D | registered_at | 登録日時（自動登録） |
| E | notified | 通知チェックボックス（手動チェック） |
| F | sent_at | 通知送信日時（自動記録） |

---

## 新規顧客の追加手順

### ステップ1: 顧客がLINEで友だち追加

顧客に以下のいずれかで友だち追加してもらう：
- LINE ID: @095ylsrv
- QRコード（LINE Official Account Managerで取得可能）

### ステップ2: 顧客が画像を送信

顧客が試しに画像を送ると、以下のメッセージが返る：
```
⏳ 初回のお客様ですね。

担当者が設定を行いますので、しばらくお待ちください。
設定完了後、改めて領収書をお送りください。
```

この時点で、スプレッドシートに顧客のLINEユーザーIDが自動登録される。

### ステップ3: Google Driveで顧客用フォルダを作成

1. Google Driveで新しいフォルダを作成
   - 例: `04_レシート読み込み/顧客名_領収書`

2. フォルダをサービスアカウントと共有
   - メールアドレス: `845322634063-compute@developer.gserviceaccount.com`
   - 権限: **編集者**

3. フォルダIDをコピー
   - URLの `https://drive.google.com/drive/folders/XXXXX` の `XXXXX` 部分

### ステップ4: スプレッドシートで顧客情報を設定

LINE顧客管理スプレッドシートで：

| 列 | 内容 | 例 |
|----|------|-----|
| A列 | line_user_id | （自動入力済み） |
| B列 | customer_name | 山田商店 |
| C列 | folder_id | 1KDEiRDVzp3r7LN8bSo5Y8oJZEcIbjZ0T |
| D列 | registered_at | （自動入力済み） |

### ステップ5: 顧客に自動通知を送信

1. **E列のチェックボックス**にチェックを入れる
2. 自動でLINEに以下のメッセージが送信される：
   ```
   ✅ 設定が完了しました！
   
   領収書の写真をこのトークに送ってください。
   自動で保存されます。
   
   📸 複数枚まとめて送ってもOKです！
   ```
3. F列に送信日時が自動記録される

**注意**: チェックは1回だけ入れてください。F列に日時が入っていれば再送信されません。

---

## 日常運用

### 顧客が領収書を送信した場合

1. 自動でGoogle Driveの該当フォルダに保存される
2. LINEに「✅ 領収書を保存しました」と返信される
3. 特に作業は不要

### 新しい顧客が来た場合

上記「新規顧客の追加手順」を実行

---

## トラブルシューティング

### 「保存に失敗しました」エラー

**原因1: フォルダの共有設定**
- 顧客のフォルダがサービスアカウントと共有されているか確認
- メールアドレス: `845322634063-compute@developer.gserviceaccount.com`
- 権限: 編集者

**原因2: フォルダIDの誤り**
- スプレッドシートのC列に正しいフォルダIDが入っているか確認

### 「初回のお客様ですね」が何度も出る

- スプレッドシートのC列（folder_id）が空欄のまま
- 顧客のフォルダIDを入力してください

### LINE通知が送信されない

- E列のチェックボックスにチェックを入れたか確認
- C列（folder_id）が空欄だと送信されない
- Apps Scriptのスクリプトプロパティに `LINE_ACCESS_TOKEN` が設定されているか確認

### ログの確認方法

**Cloud Run のログ:**
1. Google Cloud Console → Cloud Run → line-receipt-webhook
2. 「ログ」タブをクリック
3. エラーメッセージを確認

**Apps Script のログ:**
1. LINE顧客管理スプレッドシート → 拡張機能 → Apps Script
2. 左メニュー「実行」をクリック
3. 実行履歴とログを確認

---

## サービスアカウント情報

- メールアドレス: `845322634063-compute@developer.gserviceaccount.com`
- 用途: Google Drive / Google Sheets へのアクセス
- 注意: 新しい顧客フォルダを作成するたびに、このアカウントを「編集者」として共有する必要がある

---

## 技術情報

### Cloud Run サービス
- 名前: line-receipt-webhook
- リージョン: asia-northeast2 (大阪)
- URL: https://line-receipt-webhook-845322634063.asia-northeast2.run.app
- ランタイム: Python 3.14

### LINE Webhook URL
- https://line-receipt-webhook-845322634063.asia-northeast2.run.app

### 有効化が必要なGoogle Cloud API
- Google Drive API
- Google Sheets API

---

## 改善アイデア（将来）

- [x] フォルダID設定時に自動で顧客にLINE通知を送る ← 実装済み
- [ ] 月次で保存枚数をレポートする機能
- [ ] 既存のReceipt Engineと連携して自動でOCR処理

---

## 作成日

2026年2月12日

## 更新履歴

- 2026-02-12: 初版作成
- 2026-02-12: 自動LINE通知機能を追加（E列チェックボックス、F列送信日時）
