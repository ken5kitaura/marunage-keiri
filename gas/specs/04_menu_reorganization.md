# メニュー構成の整理

## 背景

現在のメニュー構成が実装済み機能と合っていない。
- `runAutoVerification`（未検証行を一括検証）がメニューにない
- `approveSelectedRows`（選択行を承認）がメニューにない
- 日常的に使う機能と初期設定の機能が混在している

## 目的

メニューを整理し、使いやすくする。

---

## 新しいメニュー構成

### レシート処理

| 項目 | 関数名 | 説明 |
|------|--------|------|
| レシート読み込み開始 | processReceipts | Driveフォルダからレシートを読み込み |
| --- | | セパレータ |
| 未検証行を一括検証 | runAutoVerification | GPT-5で一括検証＆自動承認 |
| 選択行を検証 | verifySelectedRows | GPT-5で選択行のみ検証 |
| 選択行を承認 | approveSelectedRows | 人間確認後に手動承認 |
| --- | | セパレータ |
| 選択行のファイルを削除 | deleteSelectedFiles | 選択行のファイルを削除済みフォルダへ移動 |
| --- | | セパレータ |
| サイドバーを表示 | showSidebar | 手書き領収書入力などのサイドバー |

### データ修正

| 項目 | 関数名 | 説明 |
|------|--------|------|
| 店舗名を一括正規化 | normalizeExistingStoreNames | 既存の店舗名を正規化（名寄せ） |
| 勘定科目を一括再判定 | fixExistingAccountTitles | Mapping辞書で勘定科目を再判定 |
| --- | | セパレータ |
| CHK/ERRのみリセット | resetCheckErrorMarks | CHECK/ERRORのみ再読み込み可能に |
| 全マークをリセット | resetProcessedMarks | 全ファイルを再読み込み可能に |

### 設定

| 項目 | 関数名 | 説明 |
|------|--------|------|
| Configシートを作成 | setupConfigSheet | 基本設定シートを作成 |
| Config_Foldersシートを作成 | createConfigFoldersSheet | フォルダ設定シートを作成 |
| Config_Mappingシートを作成 | createConfigMappingSheet | 勘定科目マッピングシートを作成 |
| --- | | セパレータ |
| Gemini APIキーを設定 | promptGeminiApiKey | GeminiのAPIキーを設定 |

### クレカ突合

| 項目 | 関数名 | 説明 |
|------|--------|------|
| 明細とレシートを突合 | reconcileWithStatements | クレカ明細とレシートを突合 |
| 突合情報をリセット | resetReconcileInfo | 突合情報をクリア |
| --- | | セパレータ |
| 明細スプレッドシートを設定 | promptStatementSpreadsheetId | 明細スプレッドシートIDを設定 |

### エクスポート

| 項目 | 関数名 | 説明 |
|------|--------|------|
| 弥生CSV出力 | exportToYayoiCSV | 弥生会計用CSVを出力 |

---

## 実装内容

`_Main.gs` の `onOpen()` 関数を以下のように変更する：

```javascript
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('レシート処理')
    .addItem('レシート読み込み開始', 'processReceipts')
    .addSeparator()
    .addItem('未検証行を一括検証', 'runAutoVerification')
    .addItem('選択行を検証', 'verifySelectedRows')
    .addItem('選択行を承認', 'approveSelectedRows')
    .addSeparator()
    .addItem('選択行のファイルを削除', 'deleteSelectedFiles')
    .addSeparator()
    .addItem('サイドバーを表示', 'showSidebar')
    .addToUi();

  ui.createMenu('データ修正')
    .addItem('店舗名を一括正規化', 'normalizeExistingStoreNames')
    .addItem('勘定科目を一括再判定', 'fixExistingAccountTitles')
    .addSeparator()
    .addItem('CHK/ERRのみリセット', 'resetCheckErrorMarks')
    .addItem('全マークをリセット', 'resetProcessedMarks')
    .addToUi();

  ui.createMenu('設定')
    .addItem('Configシートを作成', 'setupConfigSheet')
    .addItem('Config_Foldersシートを作成', 'createConfigFoldersSheet')
    .addItem('Config_Mappingシートを作成', 'createConfigMappingSheet')
    .addSeparator()
    .addItem('Gemini APIキーを設定', 'promptGeminiApiKey')
    .addToUi();

  ui.createMenu('クレカ突合')
    .addItem('明細とレシートを突合', 'reconcileWithStatements')
    .addItem('突合情報をリセット', 'resetReconcileInfo')
    .addSeparator()
    .addItem('明細スプレッドシートを設定', 'promptStatementSpreadsheetId')
    .addToUi();

  ui.createMenu('エクスポート')
    .addItem('弥生CSV出力', 'exportToYayoiCSV')
    .addToUi();
}
```

---

## 注意事項

- 関数自体は既に実装済み。メニューの登録のみ変更する。
- 既存の関数名は変更しない。

---

## 実装完了記録

| 日付 | 内容 | 担当 |
|------|------|------|
| | onOpen()のメニュー構成を変更 | Claude Code |
