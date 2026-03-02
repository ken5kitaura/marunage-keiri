# 絆パートナーズ税理士法人 - 顧客管理GAS

## 概要
絆パートナーズ税理士法人専用の顧客管理システム

## 顧客コード
- プレフィックス: `KZ`
- 形式: `KZ001`, `KZ002`, ... `KZ999`

## 機能

### メニュー
- `📁 新規フォルダ50件作成（一括）` - フォルダ/スプシ作成 → 登録 → 権限付与
- `📁 フォルダのみ作成` - 任意の件数を作成
- `📝 スプシに登録` - 既存フォルダをシートに登録
- `🔑 権限付与（全フォルダ）` - サービスアカウントに権限付与
- `📧 招待メール送信（未送信分）` - 顧客名・メールが入力された行にメール送信
- `📊 次の顧客コードを確認`

### 自動メール送信
顧客名（B列）とメール（I列）が入力されると、自動で招待メールを送信
（`onEdit`トリガーを設定する必要あり）

## セットアップ手順

1. Google Driveの「絆パートナーズ税理士法人」フォルダ内に新しいスプシを作成
2. スクリプトエディタを開く
3. `CustomerManagement.gs`の内容を貼り付け
4. メニュー → `🔧 顧客管理` → `📁 新規フォルダ50件作成（一括）`を実行

## シート構成

| 列 | 内容 |
|----|------|
| A | line_user_id |
| B | customer_name |
| C | folder_id（領収書） |
| D | registered_at |
| E | notified |
| F | sent_at |
| G | customer_code |
| H | status |
| I | email |
| J | phone |
| K | trial_count |
| L | total_count |
| M | memo |
| N | passbook_folder_id |
| O | cc_statement_folder_id |
| P | spreadsheet_url |
| Q | invitation_sent |
| R | invitation_sent_at |
