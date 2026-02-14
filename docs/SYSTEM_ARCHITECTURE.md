# まるなげ経理 - システム全体構成図

最終更新: 2026-02-14

---

## 1. 全体アーキテクチャ

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                              お客様                                          │
└─────────────────────────────────────────────────────────────────────────────┘
        │                           │                           │
        ▼                           ▼                           ▼
┌───────────────┐          ┌───────────────┐          ┌───────────────┐
│   LINE App    │          │  Webサイト     │          │   Stripe      │
│ (友だち追加)   │          │ (LP閲覧)       │          │ (決済)        │
└───────────────┘          └───────────────┘          └───────────────┘
        │                           │                           │
        ▼                           ▼                           ▼
┌───────────────┐          ┌───────────────┐          ┌───────────────┐
│ LINE Platform │          │    Vercel     │          │Stripe Webhook │
│   Webhook     │          │  (hosting)    │          │               │
└───────────────┘          └───────────────┘          └───────────────┘
        │                                                       │
        ▼                                                       ▼
┌───────────────────────────────────────────────────────────────────────────┐
│                      Google Cloud Functions                                │
│  ┌─────────────────────────┐    ┌─────────────────────────┐               │
│  │  line-receipt-webhook   │    │    stripe-webhook       │               │
│  │  (領収書受付・OCR)       │    │  (決済完了処理)          │               │
│  └─────────────────────────┘    └─────────────────────────┘               │
└───────────────────────────────────────────────────────────────────────────┘
        │                                   │
        ▼                                   ▼
┌───────────────────────────────────────────────────────────────────────────┐
│                         Google Workspace                                   │
│  ┌─────────────────────┐    ┌─────────────────────┐                       │
│  │  スプレッドシート     │    │   Google Drive      │                       │
│  │  (顧客管理)          │    │  (領収書画像保存)    │                       │
│  └─────────────────────┘    └─────────────────────┘                       │
│                │                         │                                 │
│                └────────────┬────────────┘                                 │
│                             ▼                                              │
│  ┌─────────────────────────────────────────────┐                          │
│  │           GAS (Google Apps Script)          │                          │
│  │  ┌───────────────┐    ┌───────────────┐     │                          │
│  │  │ Receipt-Engine │    │ Client GAS   │     │                          │
│  │  │  (ライブラリ)   │◄───│ (顧客別)      │     │                          │
│  │  └───────────────┘    └───────────────┘     │                          │
│  └─────────────────────────────────────────────┘                          │
└───────────────────────────────────────────────────────────────────────────┘
        │
        ▼
┌───────────────┐
│   弥生会計    │
│ (CSV出力)     │
└───────────────┘
```

---

## 2. ローカルフォルダ構成

```
~/Desktop/marunage/
├── web/                          # LP (Vercelで公開)
│   ├── index.html                # メインLP
│   ├── bpo.html                  # 経理BPOページ
│   ├── tokushoho.html            # 特商法表記
│   ├── robots.txt
│   ├── sitemap.xml
│   ├── llms.txt
│   └── marunage_icon_*.png       # ロゴ画像
│
├── gas/                          # GAS (clasp pushでデプロイ)
│   ├── .clasp.json               # clasp設定 (※gitignore)
│   ├── appsscript.json           # GAS設定
│   ├── _Main.gs                  # メイン処理
│   ├── _Types.gs                 # 型定義
│   ├── Config.gs                 # 設定
│   ├── Logic_Accounting.gs       # 会計ロジック
│   ├── Mapping.js                # 勘定科目マッピング
│   ├── Output_Yayoi.gs           # 弥生出力
│   ├── Service_OCR.gs            # OCRサービス
│   ├── Service_Reconcile.gs      # 照合サービス
│   ├── Service_Verification.gs   # 検証サービス
│   ├── Sidebar.html              # サイドバーUI
│   ├── Tools.gs                  # ツール
│   ├── Utils.gs                  # ユーティリティ
│   ├── docs/                     # GAS関連ドキュメント
│   └── specs/                    # 仕様書
│
├── docs/                         # 運用ドキュメント
│   └── LINE領収書システム_運用マニュアル.md
│
├── .gitignore
└── README.md
```

---

## 3. 外部サービス一覧

### 3.1 LINE公式アカウント

| 項目 | 値 |
|------|-----|
| アカウント名 | まるなげ経理 |
| 友だち追加URL | https://lin.ee/KbUqcWG |
| Webhook URL | https://asia-northeast1-gourmetlabo.cloudfunctions.net/line-receipt-webhook |
| 管理者LINE ID | U6980f7c583babed09518d986f704e959 |

**リッチメニュー構成:**
- 左: 「お試し」
- 中: 「契約済み・利用開始」
- 右: 「ヘルプ」

**環境変数 (Cloud Functions):**
- `LINE_CHANNEL_SECRET`
- `LINE_CHANNEL_ACCESS_TOKEN`

---

### 3.2 Stripe

| 項目 | 値 |
|------|-----|
| ダッシュボード | https://dashboard.stripe.com/ |
| Webhook URL | https://asia-northeast1-gourmetlabo.cloudfunctions.net/stripe-webhook |
| Webhookイベント | checkout.session.completed |

**Payment Links:**

| プラン | 金額 | URL |
|--------|------|-----|
| ミニマル (30行) | ¥5,000/月 | https://buy.stripe.com/6oUeVdeuA2W0abA7aO3Ru00 |
| ライト (100行) | ¥10,000/月 | https://buy.stripe.com/dRm3cvaek7cgerQ9iW3Ru01 |
| スタンダード (200行) | ¥14,000/月 | https://buy.stripe.com/28EaEXgCI2W06ZoeDg3Ru02 |

**超過料金:** 20円/行

**環境変数 (Cloud Functions):**
- `STRIPE_WEBHOOK_SECRET`
- `STRIPE_SECRET_KEY`

---

### 3.3 Google Cloud Platform

**プロジェクト:** gourmetlabo

**Cloud Functions:**

| 関数名 | トリガー | 役割 |
|--------|----------|------|
| line-receipt-webhook | HTTP | LINE Webhook受信、OCR、画像保存 |
| stripe-webhook | HTTP | Stripe決済完了処理、コード割り当て |

**サービスアカウント:**
- `845322634063-compute@developer.gserviceaccount.com`
- Google Drive/Sheets APIへのアクセス権限

**使用API:**
- Cloud Functions
- Google Sheets API
- Google Drive API
- Gemini API (OCR用)

**環境変数 (共通):**
- `CUSTOMER_SHEET_ID`: 顧客管理スプレッドシートID
- `GAS_UPLOAD_URL`: GAS画像アップロードURL
- `GEMINI_API_KEY`: Gemini APIキー

---

### 3.4 Google Workspace

#### スプレッドシート: 顧客管理

| 項目 | 値 |
|------|-----|
| ID | 1sUqtfRUSjg4XAAymuxCJm8masTkmxBVle70rCGxglXU |
| シート名 | 顧客管理 |

**列構成:**

| 列 | 内容 |
|----|------|
| A | line_user_id |
| B | customer_name |
| C | folder_id (領収書フォルダ) |
| D | created_at |
| E | (未使用) |
| F | (未使用) |
| G | customer_code (MK001等) |
| H | status (未使用/契約済/お試し) |
| I | email |
| J | (未使用) |
| K | trial_count (お試し送信回数) |
| L | (未使用) |
| M | (未使用) |
| N | stripe_customer_id |
| O | plan (ミニマル/ライト/スタンダード) |
| P | monthly_price |
| Q | billing_start_date |

#### Google Drive: 顧客フォルダ

| 項目 | 値 |
|------|-----|
| 親フォルダID | 1_D9JjIRLhZ6PWgVgOj4MjrWSNCEuID3N |
| 構成 | MK005_顧客名/領収書/ |
| 事前作成済み | MK005〜MK030 (26フォルダ) |

---

### 3.5 GAS (Google Apps Script)

#### Receipt-Engine (ライブラリ)

| 項目 | 値 |
|------|-----|
| スクリプトID | (※.clasp.jsonに記載) |
| 役割 | OCR、検証、仕訳、弥生出力の共通ロジック |
| デプロイ | `cd ~/Desktop/marunage/gas && clasp push` |

#### Client GAS (顧客別)

| 項目 | 説明 |
|------|------|
| 役割 | 顧客のスプレッドシートに紐づく処理 |
| 設置タイミング | 契約後、領収書が溜まってから |
| ライブラリ読込 | Receipt-Engineを参照 |

---

### 3.6 Vercel (Webホスティング)

| 項目 | 値 |
|------|-----|
| プロジェクト名 | marunage-keiri |
| 本番URL | https://marunagekeiri.com |
| Root Directory | web |
| GitHubリポジトリ | (自動連携) |

**デプロイ:** `git push origin main` で自動デプロイ

---

### 3.7 Cloudflare (DNS)

| 項目 | 値 |
|------|-----|
| ドメイン | marunagekeiri.com |
| 役割 | DNSのみ (ホスティングはVercel) |

---

## 4. データフロー

### 4.1 お試しフロー

```
1. 顧客がLINE友だち追加
   └─→ LINE Platform → line-receipt-webhook
       └─→ スプシに新規登録 (status: お試し)

2. 顧客が領収書写真を送信
   └─→ LINE Platform → line-receipt-webhook
       ├─→ Gemini API (OCR)
       └─→ LINE返信 (読み取り結果 + プラン案内)
```

### 4.2 契約フロー

```
1. 顧客がPayment Linkで決済
   └─→ Stripe → stripe-webhook
       ├─→ スプシで未使用コード検索
       ├─→ 顧客情報を書き込み (status: 契約済)
       └─→ 管理者にLINE通知

2. 顧客がLINEでコード入力 (例: MK005)
   └─→ LINE Platform → line-receipt-webhook
       ├─→ スプシでLINE ID紐付け
       ├─→ Driveフォルダ名変更 (MK005 → MK005_顧客名)
       └─→ 管理者にLINE通知
```

### 4.3 領収書処理フロー (契約後)

```
1. 顧客がLINEで領収書送信
   └─→ line-receipt-webhook
       ├─→ GAS経由でDriveに画像保存
       └─→ LINE返信 (受領確認)

2. 月末にあなたが処理
   └─→ Client GAS (スプシから起動)
       ├─→ Receipt-Engine (OCR + 検証)
       └─→ 弥生CSV出力
```

---

## 5. 環境変数まとめ

### Cloud Functions: line-receipt-webhook

```
LINE_CHANNEL_SECRET=xxx
LINE_CHANNEL_ACCESS_TOKEN=xxx
CUSTOMER_SHEET_ID=1sUqtfRUSjg4XAAymuxCJm8masTkmxBVle70rCGxglXU
GAS_UPLOAD_URL=https://script.google.com/macros/s/xxx/exec
GEMINI_API_KEY=xxx
```

### Cloud Functions: stripe-webhook

```
STRIPE_WEBHOOK_SECRET=whsec_xxx
STRIPE_SECRET_KEY=sk_live_xxx
CUSTOMER_SHEET_ID=1sUqtfRUSjg4XAAymuxCJm8masTkmxBVle70rCGxglXU
LINE_CHANNEL_ACCESS_TOKEN=xxx
```

---

## 6. デプロイ手順

### LP更新

```bash
cd ~/Desktop/marunage
git add -A
git commit -m "更新内容"
git push origin main
# → Vercelが自動デプロイ
```

### GAS更新

```bash
cd ~/Desktop/marunage/gas
clasp push
```

### Cloud Functions更新

```bash
cd [Cloud Functions用フォルダ]
gcloud functions deploy line-receipt-webhook \
  --runtime python312 \
  --trigger-http \
  --allow-unauthenticated \
  --region asia-northeast1

gcloud functions deploy stripe-webhook \
  --runtime python312 \
  --trigger-http \
  --allow-unauthenticated \
  --region asia-northeast1
```

---

## 7. 料金プラン

| プラン | 行数/月 | 月額(税別) |
|--------|---------|-----------|
| ミニマル | 30行 | ¥5,000 |
| ライト | 100行 | ¥10,000 |
| スタンダード | 200行 | ¥14,000 |
| 超過 | - | ¥20/行 |

※全プラン: プロの最終チェック付き

---

## 8. 連絡先・URL一覧

| 項目 | URL/連絡先 |
|------|-----------|
| 本番サイト | https://marunagekeiri.com |
| BPOページ | https://marunagekeiri.com/bpo.html |
| LINE追加 | https://lin.ee/KbUqcWG |
| note | https://note.com/marunagekeiri |
| メール | info@marunagekeiri.com |

---

## 9. 未実装・TODO

- [ ] 顧客へのコード通知メール (SendGrid/AWS SES)
- [ ] 超過行数の自動請求フロー
- [ ] 顧客向けダッシュボード (月間利用状況)
