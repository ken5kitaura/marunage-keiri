# まるなげ経理

中小企業・個人事業主向けの記帳代行サービス

## フォルダ構成

```
marunage/
├── web/              # LP（GitHub → Vercelで自動デプロイ）
├── gas/              # GAS（clasp pushでデプロイ）
├── functions/        # Cloud Functions（gcloudでデプロイ）
│   ├── line-receipt-webhook/   # LINE Webhook
│   └── stripe-webhook/         # Stripe Webhook
├── docs/             # 運用ドキュメント
└── README.md
```

## デプロイ

### LP（web/）

```bash
cd ~/Desktop/marunage
git add -A && git commit -m "更新内容" && git push origin main
```

GitHub → Vercel 自動デプロイ（ルートディレクトリ: `web`）

### GAS（gas/）

```bash
cd ~/Desktop/marunage/gas
clasp push
```

### Cloud Functions（functions/）

```bash
# LINE Webhook
cd ~/Desktop/marunage/functions/line-receipt-webhook
gcloud functions deploy line-receipt-webhook \
  --runtime python312 \
  --trigger-http \
  --allow-unauthenticated \
  --region asia-northeast1 \
  --entry-point line_webhook

# Stripe Webhook
cd ~/Desktop/marunage/functions/stripe-webhook
gcloud functions deploy stripe-webhook \
  --runtime python312 \
  --trigger-http \
  --allow-unauthenticated \
  --region asia-northeast1 \
  --entry-point stripe_webhook
```

## 環境変数

### line-receipt-webhook

- `LINE_CHANNEL_SECRET`
- `LINE_CHANNEL_ACCESS_TOKEN`
- `CUSTOMER_SHEET_ID`
- `GAS_UPLOAD_URL`
- `GEMINI_API_KEY`

### stripe-webhook

- `STRIPE_WEBHOOK_SECRET`
- `STRIPE_API_KEY`
- `CUSTOMER_SHEET_ID`
- `LINE_CHANNEL_ACCESS_TOKEN`
- `LINE_NOTIFY_USER_ID`

## 関連サービス

- **LP**: https://marunagekeiri.com
- **LINE**: https://lin.ee/KbUqcWG
- **Cloud Functions**: asia-northeast1-gourmetlabo.cloudfunctions.net
- **スプレッドシート**: 顧客管理シート (ID: 1sUqtfRUSjg4XAAymuxCJm8masTkmxBVle70rCGxglXU)
