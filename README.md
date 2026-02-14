# まるなげ経理

中小企業・個人事業主向けの記帳代行サービス

## フォルダ構成

```
marunage/
├── web/          # LP（Cloudflare Pagesにデプロイ）
├── gas/          # GAS（clasp pushでデプロイ）
├── docs/         # 運用ドキュメント
└── README.md
```

## デプロイ

### LP（web/）

```bash
git add -A && git commit -m "更新内容" && git push origin main
```

Cloudflare Pagesが自動デプロイ（ルートディレクトリ: `web`）

### GAS（gas/）

```bash
cd gas
clasp push
```

## 関連サービス

- **LP**: https://marunagekeiri.com
- **LINE**: https://lin.ee/KbUqcWG
- **Cloud Functions**: line-receipt-webhook, stripe-webhook
- **スプレッドシート**: 顧客管理シート
