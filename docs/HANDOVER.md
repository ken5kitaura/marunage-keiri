# まるなげ経理 - システム全体引き継ぎドキュメント

最終更新: 2026-03-02（ハイライト修正・外貨レシート対応・トリガー再作成・税理士操作ガイド作成）

---

## 📋 目次

1. [システム概要](#1-システム概要)
2. [2つの事業モデル（B2C / B2B）](#2-2つの事業モデルb2c--b2b)
3. [ローカルプロジェクト構成](#3-ローカルプロジェクト構成)
4. [GASライブラリ構成](#4-gasライブラリ構成)
5. [Cloud Functions](#5-cloud-functions)
6. [LINE Webhook処理フロー](#6-line-webhook処理フロー)
7. [顧客管理システム（B2C / B2B別）](#7-顧客管理システムb2c--b2b別)
8. [請求管理システム](#8-請求管理システム)
9. [レシート処理・検証フロー](#9-レシート処理検証フロー)
10. [通帳処理フロー](#10-通帳処理フロー)
11. [外部サービス・認証情報](#11-外部サービス認証情報)
12. [完了済み機能](#12-完了済み機能)
13. [未完了・保留タスク](#13-未完了保留タスク)
14. [既知の問題点](#14-既知の問題点)
15. [設計上の重要な決定事項](#15-設計上の重要な決定事項)
16. [デプロイ手順](#16-デプロイ手順)
17. [新規パートナー追加チェックリスト](#17-新規パートナー追加チェックリスト)
18. [外貨レシート対応](#18-外貨レシート対応)

---

## 1. システム概要

### 1.1 サービス概要

「まるなげ経理」は、**LINEで領収書・通帳を送るだけで記帳代行を行うサービス**。

- 顧客がLINEで画像を送信
- Cloud Functions (line-receipt-webhook) が受信・分類・保存
- GASライブラリ (ReceiptEngine) がOCR・検証・仕訳・弥生CSV出力

### 1.2 設計ドキュメントの所在

| ドキュメント | パス | 内容 |
|-------------|------|------|
| システム全体構成図 | `docs/SYSTEM_ARCHITECTURE.md` | 全体アーキテクチャ、外部サービス一覧 |
| 絆パートナーズフロー | `docs/kizuna-partners-flow.md` | 税理士法人向けフロー詳細 |
| 請求管理設計 | `docs/billing-management.md` | 課金カウント・請求管理の設計 |
| 税理士向け操作ガイド | `docs/TAX_ACCOUNTANT_GUIDE.md` | 税理士が自分で操作するための手順書 |
| GAS仕様書群 | `gas/specs/` | 各機能の詳細仕様 |
| クライアントラッパー説明 | `gas/client-wrappers/README.md` | 顧客別GASの設置方法 |

### 1.3 ドキュメントの問題点

**現状の課題:**
- B2CとB2Bの分け方が明確に文書化されていない
- どのコードがどのシステム用かが分散している
- 全体像を把握するには複数ファイルを見る必要がある

---

## 2. 2つの事業モデル（B2C / B2B）

### 2.1 B2C（まるなげ経理 直販）

**対象:** 個人事業主・小規模法人が直接契約

**料金:**
| プラン | 行数/月 | 月額 |
|--------|---------|------|
| ミニマル | 30行 | ¥5,000 |
| ライト | 100行 | ¥10,000 |
| スタンダード | 200行 | ¥14,000 |
| 超過 | - | ¥20/行 |

**顧客コード体系:** `MK001` 〜 `MK999`

**決済:** Stripe Payment Links → stripe-webhook → 顧客管理シート更新

**LINE公式アカウント:** まるなげ経理（1つ）

### 2.2 B2B（税理士法人向けOEM）

**対象:** 税理士法人が自社顧問先に提供

**料金:** 出力行数 × 単価（例: ¥20/行）で税理士法人に請求

**顧客コード体系:** 税理士法人ごとにプレフィックス
- 絆パートナーズ: `KZ001` 〜 `KZ999`
- （将来追加可能）

**LINE公式アカウント:** 絆パートナーズ経理（@821hkrnz）— 作成済み

### 2.3 B2C / B2B の技術的な違い

| 項目 | B2C | B2B |
|------|-----|-----|
| LINE Webhook | `line-receipt-webhook` (共通) | 同上（プレフィックスで判別） |
| 顧客管理シート | `1sUqtfRUSjg4XAAymuxCJm8masTkmxBVle70rCGxglXU` | 税理士法人ごとに別シート |
| 顧客管理GAS | `gas/customer-management/` | `gas/customer-management-kizuna/` |
| 顧客登録方法 | Stripe決済 → 自動 | メニューから招待メール → LINE連携 |
| 請求 | 顧客に直接 | 税理士法人に集約請求 |

---

## 3. ローカルプロジェクト構成

```
~/Desktop/marunage/
├── docs/                              # 設計・運用ドキュメント
│   ├── SYSTEM_ARCHITECTURE.md         # 全体構成図
│   ├── kizuna-partners-flow.md        # 絆パートナーズ向けフロー
│   ├── billing-management.md          # 請求管理設計
│   └── kizuna-partners-flow-for-accountants.txt
│
├── functions/                         # Cloud Functions
│   ├── line-receipt-webhook/          # LINE Webhook受信
│   │   ├── main.py
│   │   └── requirements.txt
│   └── stripe-webhook/                # Stripe決済処理
│       ├── main.py
│       └── requirements.txt
│
├── gas/                               # GAS (ReceiptEngineライブラリ)
│   ├── .clasp.json                    # clasp設定（ライブラリID）
│   ├── appsscript.json
│   │
│   ├── _Main.gs                       # メニュー・エントリポイント
│   ├── _Types.gs                      # 型定義
│   ├── Config.gs                      # 設定管理
│   ├── Logic_Accounting.gs            # 会計ロジック（税計算・科目判定）
│   ├── Mapping.js                     # 勘定科目マッピング辞書
│   ├── Output_Yayoi.gs                # 弥生CSV出力 + 出力フラグ管理
│   ├── PassbookEngine.gs              # 通帳OCR・処理
│   ├── Service_OCR.gs                 # Gemini OCR
│   ├── Service_Reconcile.gs           # クレカ明細突合
│   ├── Service_Verification.gs        # GPT-5検証
│   ├── Sidebar.html                   # サイドバーUI
│   ├── Tools.gs                       # ツール類
│   ├── Utils.gs                       # ユーティリティ
│   │
│   ├── billing-management/            # 請求管理GAS（けんご専用）
│   │   ├── .clasp.json
│   │   ├── BillingManagement.gs       # 月次集計
│   │   └── ExportFlagUtils.gs         # 出力フラグユーティリティ
│   │
│   ├── client-wrappers/               # 顧客用ラッパー
│   │   ├── _TEMPLATE.gs               # テンプレート（✅ 2026-02-18 更新済み）
│   │   ├── MK004.gs                   # MK004用（参考）
│   │   └── README.md
│   │
│   ├── central-admin/                 # 中央管理GAS（✅ デプロイ済み）
│   │   ├── .clasp.json
│   │   ├── appsscript.json
│   │   └── CentralAdmin.gs
│   │
│   ├── customer-management/           # B2C顧客管理GAS（✅ 2026-02-18 更新済み）
│   │   ├── .clasp.json
│   │   └── CustomerManagement.gs
│   │
│   ├── customer-management-kizuna/    # B2B顧客管理GAS（✅ 2026-02-18 更新済み）
│   │   ├── .clasp.json
│   │   └── CustomerManagement.gs
│   │
│   ├── docs/                          # GAS関連ドキュメント
│   │   └── PassbookWrapper_Sample.gs
│   │
│   └── specs/                         # 機能仕様書
│       ├── 00_project_overview.md
│       ├── 01_auto_verification.md
│       └── ...
│
├── web/                               # LP (Vercelで公開)
│   ├── index.html                     # メインLP
│   ├── bpo.html                       # 経理BPOページ
│   └── ...
│
├── .gitignore
└── README.md
```

---

## 4. GASライブラリ構成

### 4.1 ReceiptEngine（メインライブラリ）

**スクリプトID:** `1PyeRXm9QGaHcI3e4a2--mv6yV_w7FwbkqNUO5_vtqcNlUDE__Vr2bBOJ`

**デプロイ:** `cd ~/Desktop/marunage/gas && clasp push`

**ファイル構成と責務:**

| ファイル | 責務 |
|----------|------|
| `_Main.gs` | メニュー作成、エントリポイント |
| `Config.gs` | 設定管理（スクリプトプロパティ、シート設定） |
| `Service_OCR.gs` | Gemini APIでレシートOCR |
| `Service_Verification.gs` | GPT-5で読み取り結果を検証 |
| `Logic_Accounting.gs` | 税計算、勘定科目判定 |
| `Output_Yayoi.gs` | 弥生CSV出力 + **出力フラグ管理（課金用）** |
| `PassbookEngine.gs` | 通帳OCR・スプシ出力 |
| `Service_Reconcile.gs` | クレカ明細との突合 |
| `Mapping.js` | 店舗名→勘定科目のマッピング辞書 |

### 4.2 顧客用ラッパー（✅ MK001〜MK004 配布完了）

各顧客のスプシには薄いラッパーGASを設置し、ReceiptEngineをライブラリとして参照。

**テンプレート:** `gas/client-wrappers/_TEMPLATE.gs`

**配布済み顧客:** MK001, MK002, MK003, MK004（全てReceiptEngineライブラリ追加済み）

**ラッパーの役割:**
- サイドバーUI表示
- 手動操作メニュー（レシート処理、検証、弥生CSV出力）
- ClientConfigシートからフォルダID取得

**設置手順:**
1. 顧客スプシのGASエディタを開く
2. `_TEMPLATE.gs` の内容を貼り付け
3. ライブラリ追加: `ReceiptEngine`（スクリプトID: `1PyeRXm9QGaHcI3e4a2--mv6yV_w7FwbkqNUO5_vtqcNlUDE__Vr2bBOJ`）
4. `ClientConfig` シートを作成し、フォルダIDを設定

### 4.3 顧客管理GAS（B2C / B2B別）（✅ 2026-02-18 ステータス統一済み）

| 用途 | パス | 顧客コード | 対象シート |
|------|------|-----------|-----------| 
| B2C | `gas/customer-management/` | MK | 顧客管理（まるなげ経理直販） |
| B2B | `gas/customer-management-kizuna/` | KZ | 顧客管理（絆パートナーズ） |

**ステータス定義（MK・KZ共通 ✅ 2026-02-18 統一）:**

| ステータス | 意味 | レシート処理対象？ |
|-----------|------|-------------------|
| 未使用 | フォルダ・スプシ作成済み、顧客未割当 | ✗ |
| 案内済 | 顧客に案内済み（MK: Stripe決済完了 / KZ: 招待メール送信済み） | ✗ |
| 契約済 | LINEコード連携完了、サービス利用中 | ✓ |
| 解約 | サービス終了 | ✗ |

**MK（B2C）の特徴:**
- Stripe決済完了 → 自動で「案内済」に変更
- LINEでコード入力 → 自動で「契約済」に変更
- 招待メール送信機能なし（Stripe経由のため不要）

**KZ（B2B）の特徴:**
- 顧客名・メールアドレス入力後、メニューから「📧 選択行に招待メール送信」を実行
- 送信成功 → 自動で「案内済」に変更
- LINEでコード入力 → 自動で「契約済」に変更
- 複数行選択で一括送信可能

**設計方針:** 各パートナーごとに列構成・招待方法が異なるため、個別のGASとして維持。ステータス定義は統一。一括処理は中央管理GASに集約。

### 4.4 中央管理GAS（✅ 2026-02-18 デプロイ・設定完了）

**パス:** `gas/central-admin/`

**3層構造:**
- **第1層: 中央管理GAS**（1つだけ） — 全パートナー横断で一括処理
- **第2層: パートナー別GAS**（MK用、KZ用、将来のXX用） — フォルダ作成・招待等
- **第3層: 顧客スプシ**（100枚〜） — ラッパーGAS + ReceiptEngine

**機能:**
- 全アクティブ顧客のレシート一括処理（1時間毎トリガー）
- 全アクティブ顧客のAI検証一括実行（1日1回トリガー）
- ラッパーGAS一括配布（Apps Script API）— テンプレート登録済み
- パートナー設定管理（「パートナー設定」シートで列番号・ステータス等を定義）
- 実行ログ記録

**完了済みセットアップ:**
- パートナー設定シート初期化済み（MK・KZ両方登録）
- ReceiptEngine ScriptID設定済み
- ラッパーテンプレート登録済み
- トリガー設定済み（1時間毎のレシート処理、1日1回の自動検証）
- アクティブステータス: 「契約済」（G列）

**パートナー設定シート列構成:**
| 列 | 名前 | 説明 |
|----|------|------|
| A | パートナー名 | 表示名 |
| B | シートID | 顧客管理スプレッドシートID |
| C | シート名 | シート名（通常「顧客管理」） |
| D | 顧客コード列 | 0始まりの列番号 |
| E | ステータス列 | 0始まりの列番号 |
| F | スプシURL列 | 0始まりの列番号 |
| G | アクティブステータス | カンマ区切り（例: 契約済） |
| H | 有効 | TRUE/FALSE |

### 4.5 請求管理GAS

**パス:** `gas/billing-management/`

**対象:** けんご専用の請求管理スプシ

**スプシURL:** `https://docs.google.com/spreadsheets/d/1STg-OxRkPjCfYlUYcQpPNKjY68-QXJLU2XV-jPm0x3w/`

**機能:**
- 税理士一覧の管理
- 月次出力行数の集計
- 請求金額の計算

---

## 5. Cloud Functions

### 5.1 line-receipt-webhook（✅ 2026-02-18 KZ対応追加）

**URL:** `https://line-receipt-webhook-845322634063.asia-northeast2.run.app`

**リージョン:** asia-northeast2（第2世代 Cloud Run）

**パス:** `functions/line-receipt-webhook/main.py`

**責務:**
1. LINEからのWebhook受信（署名検証）
2. 画像/PDF/テキストの種別判定
3. Geminiで分類（レシート/通帳/クレカ売上票）
4. Google Driveに保存
5. 顧客管理シートの更新
6. LINE返信

**処理する顧客コードプレフィックス:**
- `MK`（B2C）— CUSTOMER_SHEET_ID を参照
- `KZ`（B2B）— KZ_CUSTOMER_SHEET_ID を参照（✅ 2026-02-18 追加）

**KZ対応の仕組み:**
- `get_sheet_id_for_code()` — コードプレフィックスに応じてシートIDを返す
- `get_code_col_for_prefix()` — MKはG列(6)、KZはE列(4)
- `get_status_col_for_prefix()` — MKはH列、KZはF列
- `customer_code_exists()` / `link_user_with_customer_code()` — MK/KZ両方に対応

**LINE連携時の動作（MK・KZ共通）:**
- 顧客がLINEでコード入力 → LINE IDを紐付け → ステータスを「契約済」に変更

### 5.2 stripe-webhook（✅ 2026-02-18 ステータス変更）

**URL:** `https://stripe-webhook-845322634063.asia-northeast2.run.app`

**リージョン:** asia-northeast2（第2世代 Cloud Run）

**パス:** `functions/stripe-webhook/main.py`

**責務:**
- `checkout.session.completed` イベント受信
- 未使用の顧客コード（MK）を割り当て
- ステータスを「**案内済**」に変更（✅ 変更: 旧「契約済」→ 新「案内済」）
- 管理者へLINE通知

**ステータス変更の理由:**
Stripe決済完了時点ではまだLINE連携が済んでいないため「案内済」とし、
LINE連携完了時に初めて「契約済」にする（line-receipt-webhookが担当）。

---

## 6. LINE Webhook処理フロー

### 6.1 お試しフロー

```
顧客: LINE友だち追加
  ↓
line-receipt-webhook: ウェルカムメッセージ送信（登録はしない）
  ↓
顧客: 領収書画像送信
  ↓
line-receipt-webhook:
  1. 顧客未登録なら「お試し」として登録
  2. Geminiで分類・OCR
  3. 結果をLINE返信（料金プラン案内付き）
```

### 6.2 契約フロー（B2C）（✅ 2026-02-18 更新）

```
顧客: Stripe Payment Linkで決済
  ↓
stripe-webhook:
  1. 未使用のMKコードを検索
  2. 顧客情報を書き込み（status: 案内済 ← 変更点）
  3. 管理者にLINE通知
  ↓
顧客: LINEでコード入力（例: MK005）
  ↓
line-receipt-webhook:
  1. コードを正規化・検証
  2. LINE IDと紐付け
  3. ステータスを「契約済」に変更 ← ここで初めて契約済
  4. 同じLINE IDの「お試し」行を削除
  5. フォルダ名変更（MK005 → MK005_顧客名）
  6. 管理者にLINE通知
```

### 6.3 契約フロー（B2B - 税理士法人向け）（✅ 2026-02-18 更新）

```
税理士: 顧客管理シートで顧客登録
  1. 「未使用」行を探す
  2. 顧客名・メールを入力
  3. 行を選択してメニューから「📧 選択行に招待メール送信」
  ↓
GAS: 確認ダイアログ → 招待メール送信 → ステータスを「案内済」に変更
  ↓
顧客: 招待メール受信
  ↓
顧客: LINEで友だち追加 + コード入力（例: KZ001）
  ↓
line-receipt-webhook:
  1. コードを正規化・検証（MK/KZ両対応）
  2. KZ_CUSTOMER_SHEET_IDから該当シートを特定
  3. LINE IDと紐付け
  4. ステータスを「契約済」に変更
```

---

## 7. 顧客管理システム（B2C / B2B別）

### 7.1 B2C顧客管理シート

**シートID:** `1sUqtfRUSjg4XAAymuxCJm8masTkmxBVle70rCGxglXU`

**列構成（21列）:**

| 列 | 名前 | 説明 |
|----|------|------|
| A | line_user_id | LINE ユーザーID |
| B | customer_name | 顧客名 |
| C | folder_id | 領収書フォルダID |
| D | created_at | 登録日時 |
| E | (未使用) | - |
| F | (未使用) | - |
| G | customer_code | 顧客コード（MK001等） |
| H | status | ステータス（未使用/案内済/契約済/解約） |
| I | email | メールアドレス |
| J | (未使用) | - |
| K | trial_count | お試し送信回数 |
| L | (未使用) | - |
| M | (未使用) | - |
| N | stripe_customer_id | Stripe顧客ID |
| O | plan | プラン名 |
| P | monthly_price | 月額料金 |
| Q | billing_start_date | 課金開始日 |
| R | spreadsheet_url | 顧客用スプシURL |
| S | passbook_folder_id | 通帳フォルダID |
| T | cc_statement_folder_id | クレカ明細フォルダID |
| U | invitation_sent_at | 招待メール送信日時 |

### 7.2 B2B顧客管理シート（絆パートナーズ）

**シートID:** `1w8KfoYs6RFjNM6LZvH9qoD5e5Ow_vB6DrztsUF7nKzg`

**列構成（13列）:**

| 列 | 名前 | 説明 |
|----|------|------|
| A | line_user_id | LINE ユーザーID |
| B | customer_name | 顧客名 |
| C | folder_id | 領収書フォルダID |
| D | registered_at | 登録日時 |
| E | customer_code | 顧客コード（KZ001等） |
| F | status | ステータス（未使用/案内済/契約済） |
| G | email | メールアドレス |
| H | phone | 電話番号 |
| I | memo | メモ |
| J | passbook_folder_id | 通帳フォルダID |
| K | cc_statement_folder_id | クレカ明細フォルダID |
| L | spreadsheet_url | 顧客用スプシURL |
| M | invitation_sent_at | 招待メール送信日時 |

### 7.3 フォルダ構成

```
親フォルダ/
├── KZ001/
│   ├── 領収書/          ← folder_id（LINE画像の保存先）
│   ├── 通帳/            ← passbook_folder_id
│   └── クレカ明細/       ← cc_statement_folder_id
└── KZ002/
    └── ...
```

### 7.4 ステータス遷移図（✅ 2026-02-18 統一）

**MK（B2C）:**
```
未使用 → [Stripe決済] → 案内済 → [LINEコード入力] → 契約済 → [解約] → 解約
```

**KZ（B2B）:**
```
未使用 → [メニューからメール送信] → 案内済 → [LINEコード入力] → 契約済
```

**統一ルール:**
- 「契約済」になるタイミング = LINEでコード入力した時点（MK・KZ共通）
- レシート処理対象 = 「契約済」のみ
- 中央管理GASのアクティブステータス = 「契約済」

---

## 8. 請求管理システム

### 8.1 設計概要

**目的:** 税理士法人ごとの出力行数を集計し、請求金額を計算

**請求管理スプシ:** `https://docs.google.com/spreadsheets/d/1STg-OxRkPjCfYlUYcQpPNKjY68-QXJLU2XV-jPm0x3w/`

### 8.2 出力フラグの仕組み

各顧客スプシの「本番シート」に3列を追加:

| 列 | 名前 | 説明 |
|----|------|------|
| U | 出力済 | TRUE/FALSE |
| V | 出力日 | 出力実行日時 |
| W | 出力行数 | 仕訳行数（税率ごと） |

**出力行数の計算ロジック:**

```javascript
// レシート
出力行数 = 0;
if (対象額10% > 0) 出力行数++;
if (対象額8% > 0) 出力行数++;
if (不課税 > 0) 出力行数++;

// 通帳
出力行数 = 1;  // 常に1行
```

### 8.3 集計ロジック

1. 税理士一覧シートから有効な税理士を取得
2. 各税理士の顧客管理シートを開く
3. 各顧客のスプシを巡回
4. `出力済=TRUE` かつ `日付が該当月` の `出力行数` を合計
5. 月次集計シートに記録

### 8.4 実装状況

| ファイル | 状態 | 備考 |
|----------|------|------|
| `BillingManagement.gs` | ✅ 実装済（ローカル） | push未実行 |
| `ExportFlagUtils.gs` | ✅ 実装済（ローカル） | push未実行 |
| `Output_Yayoi.gs`（フラグ追加部分） | ✅ 実装済（ローカル） | push未実行 |

---

## 9. レシート処理・検証フロー

### 9.1 処理フロー

```
1. processReceipts() 実行（中央管理GASが1時間毎にトリガー）
   ↓
2. Driveフォルダから未処理ファイル取得
   ↓
3. Gemini APIでOCR（Service_OCR.gs）
   ↓
4. 税計算・勘定科目判定（Logic_Accounting.gs）
   ↓
5. スプシに書き込み
   ↓
6. ファイル名にステータスプレフィックス付与
   [CHK] / [ERR] / [CMP] / [OK]
```

### 9.2 検証フロー

```
1. runAutoVerification() 実行（中央管理GASが1日1回トリガー）
   ↓
2. CHECK/ERROR/COMPOUNDの行を抽出
   ↓
3. GPT-5で検証（Service_Verification.gs）
   ↓
4. 自動承認条件を判定:
   - 確度スコア >= 0.90
   - 合計金額が完全一致
   - severity: high の問題がゼロ
   ↓
5a. 条件クリア → 自動承認（OK）
5b. 条件未達 → 要確認マーク
```

### 9.3 ステータス体系

**A列（メインステータス）:**

| ステータス | 意味 | 検証対象？ |
|-----------|------|-----------|
| 🟢OK | 問題なし | ✗ |
| 🔴CHECK | 整合性に疑問 | ✓ |
| 🔴ERROR | 基本情報欠落 | ✓ |
| 🟡COMPOUND | 複合仕訳 | △ |
| 🖊️HAND | 手書き | ✗ |

**17列目（検証ステータス）:**

| ステータス | 意味 |
|-----------|------|
| 🤖 自動承認 | GPT-5検証で条件クリア |
| ✅ 手動承認 | 人間が確認してOK |
| 🟡要確認 | 条件未達、人間確認必要 |

---

## 10. 通帳処理フロー

### 10.1 処理フロー

```
1. processPassbooks() 実行（手動 or トリガー）
   ↓
2. 通帳フォルダから未処理ファイル取得
   ↓
3. Gemini APIでOCR（PassbookEngine.gs）
   - ページ順序指定対応
   - 複数ページ対応
   ↓
4. 「通帳」シートに書き込み
   ↓
5. 処理済みファイル名を変更（[DONE]プレフィックス）
```

### 10.2 通帳シート列構成

| 列 | 名前 |
|----|------|
| A | 日付 |
| B | 摘要 |
| C | お支払金額 |
| D | お預り金額 |
| E | 差引残高 |
| F | ソースファイル |
| G | ページ番号 |

---

## 11. 外部サービス・認証情報

### 11.1 LINE公式アカウント

| 項目 | 値 |
|------|-----|
| アカウント名 | まるなげ経理 |
| 友だち追加URL | https://lin.ee/KbUqcWG |
| Webhook URL | `https://line-receipt-webhook-845322634063.asia-northeast2.run.app` |
| 管理者LINE ID | `U6980f7c583babed09518d986f704e959` |

### 11.2 Stripe

| 項目 | 値 |
|------|-----|
| Webhook URL | `https://stripe-webhook-845322634063.asia-northeast2.run.app` |
| Webhookイベント | `checkout.session.completed` |

**Payment Links:**

| プラン | URL |
|--------|-----|
| ミニマル | `https://buy.stripe.com/6oUeVdeuA2W0abA7aO3Ru00` |
| ライト | `https://buy.stripe.com/dRm3cvaek7cgerQ9iW3Ru01` |
| スタンダード | `https://buy.stripe.com/28EaEXgCI2W06ZoeDg3Ru02` |

### 11.3 Google Cloud Platform

| 項目 | 値 |
|------|-----|
| プロジェクト | marunage-keiri |
| リージョン | asia-northeast2 |
| サービスアカウント | `845322634063-compute@developer.gserviceaccount.com` |

### 11.4 API Keys（環境変数）

**line-receipt-webhook:**
- `LINE_CHANNEL_SECRET`
- `LINE_CHANNEL_ACCESS_TOKEN`
- `CUSTOMER_SHEET_ID`（MK用: `1sUqtfRUSjg4XAAymuxCJm8masTkmxBVle70rCGxglXU`）
- `KZ_CUSTOMER_SHEET_ID`（KZ用: `1w8KfoYs6RFjNM6LZvH9qoD5e5Ow_vB6DrztsUF7nKzg`）← ✅ 2026-02-18 追加
- `GAS_UPLOAD_URL`
- `GEMINI_API_KEY`
- `DEFAULT_FOLDER_ID`

**stripe-webhook:**
- `STRIPE_WEBHOOK_SECRET`
- `STRIPE_API_KEY`
- `CUSTOMER_SHEET_ID`
- `LINE_CHANNEL_ACCESS_TOKEN`
- `LINE_NOTIFY_USER_ID`（`U6980f7c583babed09518d986f704e959`）

**GAS (ScriptProperties):**
- `GEMINI_API_KEY`
- `OPENAI_API_KEY`

---

## 12. 完了済み機能

### 12.1 コア機能

- [x] LINE Webhook受信・分類・保存
- [x] Gemini OCR（レシート）
- [x] GPT-5検証（自動検証 + 自動承認）
- [x] 弥生CSV出力
- [x] 税計算・勘定科目自動判定
- [x] クレカ明細突合
- [x] Gemini OCR（通帳）
- [x] 通帳スプシ出力

### 12.2 顧客管理

- [x] B2C顧客管理（Stripe連携）
- [x] B2B顧客管理（絆パートナーズ）
- [x] フォルダ一括作成
- [x] 顧客用スプシ自動作成（ClientConfig付き）
- [x] 招待メール送信機能（KZ: メニューから選択行送信）
- [x] ステータス統一（未使用/案内済/契約済/解約）— MK・KZ共通
- [x] 中央管理GAS（全パートナー横断一括処理）
- [x] ラッパーGAS配布（MK001〜MK004）

### 12.3 LP・Web

- [x] メインLP
- [x] 経理BPOページ
- [x] Vercel自動デプロイ

### 12.4 インフラ・デプロイ（✅ 2026-02-18）

- [x] 中央管理GASデプロイ・パートナー設定初期化・トリガー設定
- [x] MK顧客管理GASデプロイ（ステータス統一・メール送信削除）
- [x] KZ顧客管理GASデプロイ（ステータス統一・メニュー送信追加）
- [x] Stripe Webhookデプロイ（案内済に変更）
- [x] LINE Webhookデプロイ（KZ対応追加・KZ_CUSTOMER_SHEET_ID追加）
- [x] 両顧客管理シートで初期セットアップ実行済み

---

## 13. 未完了・保留タスク

### 13.1 高優先度

| タスク | 詳細 | ブロッカー |
|--------|------|-----------|
| 請求管理GAS push | `billing-management/`のデプロイ | 動作確認必要 |
| 仕訳数の自動通知 | 税理士向けに月次仕訳数を自動通知する仕組み | 設計未着手 |
| 外貨レシート：クレカ明細照合機能 | 外貨CHECK行とクレカ明細の自動突合・日本円確定 | 設計済み・実装は次回 |
| 既存データの背景色クリア | `clearRowBackgroundColorsById` を中央管理GASから手動実行 | 実行待ち |

### 13.2 中優先度

| タスク | 詳細 |
|--------|------|
| 顧客へのコード通知メール | Stripe決済後にメール送信 |
| 超過行数の自動請求フロー | 月末に超過分を自動請求 |
| 顧客向けダッシュボード | 月間利用状況を閲覧 |
| 請求書PDF自動生成 | 税理士法人向け |

### 13.3 低優先度

| タスク | 詳細 |
|--------|------|
| 精度レポート・ダッシュボード | 検証精度の可視化 |
| アラート・監視 | 確度低下検知、メール通知 |

---

## 14. 既知の問題点

### 14.1 トリガー問題（✅ 2026-02-18 修正済み）

**現象:**
- `processReceipts`（1時間毎）と `runAutoVerification`（1日1回）のトリガーが、気づくと無効化または消失していた

**原因（特定済み）:**
- タイムアウト時の継続トリガー作成処理が、「同じ関数名のトリガーを全削除」してから新規作成していた
- この「全削除」が、継続トリガーだけでなく定期トリガー（毎時・毎日）も巻き込んで削除していた

**修正内容:**
- 継続実行用に別名のラッパー関数を導入（`processReceipts_continue`, `runAutoVerification_continue`）
- 継続トリガーはこの `_continue` 関数を呼ぶため、削除時に定期トリガーを巻き込まない
- 対象ファイル: `_Main.gs`, `Service_Verification.gs`

**現在の運用:**
- 定期トリガーは中央管理GASに移行済み（顧客スプシ側のトリガーは不要）
- 中央管理GASが全アクティブ顧客を巡回して処理

### 14.2 clasp push後のトリガー停止問題（✅ 2026-03-02 原因特定・対処済み）

**現象:**
- `clasp push` 後に中央管理GASの時間トリガー（`batchProcessAllReceipts`・`batchRunAutoVerification`）が停止
- トリガー設定は正常に見えるが発火しない状態

**原因:**
- clasp push後にGoogleが再認証を要求。認証が完了するまでトリガーが無効化される

**対処法:**
- clasp push後は必ずGASエディタのトリガー画面を確認
- トリガーが停止していたら「削除→再作成」が確実（再認証ダイアログが出るので承認）
- トリガー設定: `batchProcessAllReceipts`=時間主導型・1時間おき / `batchRunAutoVerification`=日付ベース・毎日

### 14.3 GASライブラリのバージョン管理

**現状:**
- `clasp push`でライブラリを更新すると、参照している全顧客スプシに即時反映
- 破壊的変更を入れると全顧客に影響

**対応案（未実装）:**
- バージョン番号を明示的に管理
- 顧客別にライブラリバージョンを固定

---

## 15. 設計上の重要な決定事項

### 15.1 ライブラリ構成

**決定:** ロジックはすべて `ReceiptEngine` ライブラリに集約し、顧客スプシには薄いラッパーのみ設置

**理由:**
- 修正時に1箇所を直せば全顧客に反映
- 顧客ごとの差異を最小化

**注意点:**
- ライブラリの破壊的変更は全顧客に影響
- ラッパー配布は中央管理GASの一括配布機能で対応

### 15.2 顧客コード体系

**決定:** プレフィックス2文字 + 3桁数字（例: MK001, KZ001）

**理由:**
- LINE Webhookでプレフィックスから顧客管理シートを特定
- 将来の税理士法人追加に対応可能

### 15.3 ステータス統一（✅ 2026-02-18 決定）

**決定:** MK・KZ共通で4つのステータス: 未使用 / 案内済 / 契約済 / 解約

**「契約済」の統一ルール:** LINEでコードを入力した時点で「契約済」（MK・KZ共通）

**理由:**
- レシート処理対象の判定がシンプルになる
- 中央管理GASが「契約済」だけを見ればよい
- MK/KZで「案内済になるタイミング」は違っても、「契約済になるタイミング」は同じ

### 15.4 KZ招待メール送信方式（✅ 2026-02-18 決定）

**決定:** セル変更トリガーではなく、メニューから手動実行

**理由:**
- ステータスを純粋な「状態」の表示に保つ
- 誤入力での誤送信リスクを排除
- 確認ダイアログで送信前チェック可能
- 複数行選択で一括送信にも対応

### 15.5 課金カウント方式

**決定:** 出力時に行単位でフラグをセット、月末に集計

**理由:**
- 複数回出力しても重複カウントしない
- データ修正時にフラグをリセット → 再出力で再カウント

### 15.6 検証AIの選択

**決定:**
- OCR: Gemini 2.0 Flash（安価・高速）
- 検証: GPT-5（高精度）

**理由:**
- OCRは速度・コスト優先
- 検証は精度優先（顧客の信頼に直結）

---

## 16. デプロイ手順

### 16.1 GASライブラリ更新

```bash
cd ~/Desktop/marunage/gas
clasp push
```

**⚠️ clasp push後の必須確認:**
clasp push後、中央管理GASのトリガーが停止することがある（Googleが再認証を要求するため）。
push後は必ず中央管理GASのトリガー画面を確認し、停止していたら削除→再作成すること。

### 16.2 請求管理GAS更新

```bash
cd ~/Desktop/marunage/gas/billing-management
clasp push
```

### 16.3 顧客管理GAS更新（B2C）

```bash
cd ~/Desktop/marunage/gas/customer-management
clasp push --force
```

### 16.4 顧客管理GAS更新（絆パートナーズ）

```bash
cd ~/Desktop/marunage/gas/customer-management-kizuna
clasp push --force
```

### 16.5 中央管理GAS更新

```bash
cd ~/Desktop/marunage/gas/central-admin
clasp push --force
```

### 16.6 LP更新

```bash
cd ~/Desktop/marunage
git add -A
git commit -m "更新内容"
git push origin main
# → Vercelが自動デプロイ
```

### 16.7 Cloud Functions更新

**⚠️ 重要: 環境変数が消えないよう、Google Cloud Consoleから直接ソースを貼り替えてデプロイするのが安全。**

CLIでデプロイする場合（環境変数が消えるリスクあり）:
```bash
# LINE Webhook
cd ~/Desktop/marunage/functions/line-receipt-webhook
gcloud functions deploy line-receipt-webhook \
  --runtime python312 \
  --trigger-http \
  --allow-unauthenticated \
  --region asia-northeast2

# Stripe Webhook
cd ~/Desktop/marunage/functions/stripe-webhook
gcloud functions deploy stripe-webhook \
  --runtime python312 \
  --trigger-http \
  --allow-unauthenticated \
  --region asia-northeast2
```

---

## 17. 新規パートナー追加チェックリスト

新規税理士パートナーを追加する際の手順。上から順に実施する。

### 17.1 LINE公式アカウント

- [ ] LINE公式アカウント作成
- [ ] Messaging API 有効化
- [ ] Webhook URL 設定（`https://line-receipt-webhook-845322634063.asia-northeast2.run.app`）
- [ ] 応答メッセージ OFF（Webhook側で応答するため）
- [ ] チャネルアクセストークン発行

### 17.2 Cloud Functions 環境変数

- [ ] `XX_LINE_CHANNEL_SECRET` 追加（XXはプレフィックス）
- [ ] `XX_LINE_CHANNEL_ACCESS_TOKEN` 追加
- [ ] `XX_BOT_USER_ID` 追加（初回デプロイ後にログから確認）
- [ ] `XX_CUSTOMER_SHEET_ID` 追加

### 17.3 顧客管理シート

- [ ] 顧客管理スプレッドシート作成（列構成はKZシートを参考）
- [ ] サービスアカウント（`845322634063-compute@developer.gserviceaccount.com`）を**編集者**として共有追加
  - **🚨 これを忘れるとWebhookが403エラーになる。必ず最初にやること。**
- [ ] 入力規則の「無効なデータの場合」を「警告を表示」に設定（Webhook書き込み時の警告防止）
- [ ] 初期データ（フォルダ・スプシ）のセットアップ

### 17.4 顧客管理GAS

- [ ] `gas/customer-management-XX/` ディレクトリ作成（KZをベースにコピー）
- [ ] `CONFIG.LINE_URL` に友だち追加URL設定（デフォルトはプレースホルダーなので必ず変更）
- [ ] 招待メールテンプレートの QRコードURL を更新（`https://qr-official.line.me/gs/M_XXXXX_GW.png` 形式）
- [ ] clasp push でデプロイ

### 17.5 中央管理GAS・Webhook

- [ ] 中央管理GASの「パートナー設定」シートに新パートナー行を追加
- [ ] `main.py` の `CHANNELS` マップに新チャネルを追加
- [ ] `main.py` の `BOT_USER_ID` マッピングに追加
- [ ] GASラッパー一括配布（中央管理GASから実行）

---

## 📝 次のセッションへの引き継ぎ事項

### 完了済み（2026-02-24）
- ✅ 絆パートナーズLINE公式アカウント作成・設定完了（@821hkrnz）
- ✅ line-receipt-webhookマルチチャネル対応（MK/KZ両対応）デプロイ完了
- ✅ Bot User ID取得・環境変数設定完了
- ✅ KZ招待メールテンプレート改修（HTML形式・LINEボタン・QRコード付き）
- ✅ KZフロー全体テスト完了（招待メール→LINE紐付け→領収書送信→OCR処理）
- ✅ GASライブラリ統合メニュー（「📥 未処理を一括読み込み」でレシート・通帳・クレカを一括処理）
- ✅ 通帳シート自動作成（初回実行時に自動生成）
- ✅ ラッパーGAS全体再配布（processAll対応版）
- ✅ KZ全件GAS配布完了（KZ001〜KZ050）
- ✅ LP修正（郵送→デジタル送信メイン化、料金注釈追加、構造化データFAQ更新）

### 完了済み（2026-03-02）
- ✅ HANDOVER.md更新（KZ完了項目反映、チェックリスト強化）
- ✅ 税理士向け操作ガイド作成（`docs/TAX_ACCOUNTANT_GUIDE.md`）
- ✅ ハイライト問題修正 — `_Main.gs`の行全体setBackground削除、`Service_Verification.gs`のエラーハイライトをQ列のみに限定
- ✅ 外貨レシート対応 — OCR通貨判定追加、JPY以外は消費税スキップ・CHECKステータス・通貨列(21列目)追加
- ✅ 中央管理GASトリガー再作成（clasp push後の再認証切れが原因で停止していた）
- ✅ 背景色クリア用関数追加（`clearRowBackgroundColors` / `clearRowBackgroundColorsById`）

### 実行待ち
- **既存データ背景色クリア** — 中央管理GASから `clearRowBackgroundColorsById('1PuggH0KHMLslZHGbMxBnFtjmHy_jrXNG6L3GU3yr9bU')` を実行（一時関数作成が必要）

### 保留タスク
- **請求管理GAS push** — `billing-management/` のデプロイ（動作確認待ち）
- **仕訳数の自動通知** — 税理士向けに月次仕訳数を自動通知する仕組み
- **外貨レシート：クレカ明細照合** — 設計方針確定済み、実装は次回（詳細は下記「外貨レシート対応」参照）

---

## 18. 外貨レシート対応（✅ OCR・税計算実装済 / クレカ照合は次回）

### 18.1 実装済み（2026-03-02）

**OCR通貨判定:**
- `Service_OCR.gs`: プロンプトに`currency`抽出項目を追加。$→USD、€→EUR等を自動判定
- `parseGeminiResponse_`: `currency`フィールドを返却（デフォルトJPY）

**消費税スキップ:**
- `Logic_Accounting.gs`: currency ≠ JPY の場合、税計算を全スキップ（全税額0）

**ステータス処理:**
- `_Main.gs`: 外貨レシートはCHECKステータスに強制変更
- debugInfoに「外貨レシート（USD）- クレカ明細との照合が必要」等を記載

**通貨列:**
- 21列目に「通貨」列を追加。JPY以外の場合のみ通貨コードを記入

**自動検証除外:**
- `Service_Verification.gs`: 列21が空でない＆JPY以外の行は自動検証対象から除外

**弥生CSV出力:**
- 変更不要。外貨行はCHECKのまま残るので自動的にスキップされる

### 18.2 未実装（次回）: クレカ明細との自動照合

**設計方針（確定済み）:**

1. クレカ明細読み込み時に「外貨（CHECK）」行を自動検索
2. 「日付が近い（±5日）」＋「店舗名の類似度」でマッチング候補を提示
3. マッチしたらクレカ明細の日本円金額をレシート側に上書き
4. 税区分を「不課税」に設定
5. ステータスをOKに変更
6. マッチしなかった外貨行は「未突合」のまま残し手動対応

**店舗名突合のポイント:**
- SaaS系（ClaudeならAnthropic、AWSならAmazon）は比較的容易
- 海外出張の飲食・交通は店舗名が一致しにくいので、AIに突合判定させるのが確実

---

*このドキュメントは `/Users/kengokitaura/Desktop/marunage/docs/HANDOVER.md` に保存されています。*
