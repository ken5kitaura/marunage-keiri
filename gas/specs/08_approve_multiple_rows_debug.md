# 選択行を一括承認の修正（複数範囲対応）

## 問題

「選択行を承認」で複数行を選択しても、`numRows=1` となり1行しか承認されない。

## 原因

`getActiveRange()` は単一の連続範囲しか返さない。
複数の選択範囲（Ctrl+クリック）や、離れた行を選択した場合に対応できない。

## 解決策

`getActiveRangeList()` を使用して、複数の選択範囲に対応する。

---

## 修正内容

`Service_Verification.gs` の `approveSelectedRows()` 関数を以下のように修正する。

### 修正後のコード

```javascript
/**
 * 選択された行を手動で承認する。
 * 人間が17列目（検証結果）を確認した後に使うボタン。
 *
 * 対象: CHECK / ERROR / COMPOUND 行のみ（OKやHANDは何もしない）
 * 複数の選択範囲（Ctrl+クリック）にも対応。
 */
function approveSelectedRows() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (sheet.getName() !== CONFIG.SHEET_NAME.MAIN) {
    ui.alert('「' + CONFIG.SHEET_NAME.MAIN + '」シートで実行してください。');
    return;
  }

  // 複数の選択範囲に対応するため getActiveRangeList() を使用
  const rangeList = sheet.getActiveRangeList();
  if (!rangeList) {
    ui.alert('承認したい行を選択してください。');
    return;
  }

  // 選択されている全ての行番号を収集（重複排除）
  const selectedRows = new Set();
  const ranges = rangeList.getRanges();
  
  for (const range of ranges) {
    const startRow = range.getRow();
    const numRows = range.getNumRows();
    for (let row = startRow; row < startRow + numRows; row++) {
      if (row > 1) {  // ヘッダー行は除外
        selectedRows.add(row);
      }
    }
  }

  if (selectedRows.size === 0) {
    ui.alert('承認したい行を選択してください（ヘッダー行は除外されます）。');
    return;
  }

  // デバッグログ
  const rowArray = Array.from(selectedRows).sort((a, b) => a - b);
  console.log('approveSelectedRows: 対象行 = ' + rowArray.join(', ') + ' (' + rowArray.length + '件)');

  let approvedCount = 0;
  let skippedCount = 0;

  for (const row of rowArray) {
    const rawStatus = String(sheet.getRange(row, 1).getValue() || '');
    const status = rawStatus.replace(/^[🟢🔴🟡🟠🖊️]+/, '');

    console.log('行' + row + ': rawStatus="' + rawStatus + '", status="' + status + '"');

    // CHECK / ERROR / COMPOUND のみ承認可能
    if (status !== 'CHECK' && status !== 'ERROR' && status !== 'COMPOUND') {
      console.log('行' + row + ': スキップ（ステータスが対象外）');
      skippedCount++;
      continue;
    }

    applyApproval_(sheet, row, '✅ 手動承認');
    approvedCount++;
    console.log('行' + row + ': 承認完了');
  }

  const msg = '承認完了: ' + approvedCount + '件' +
              (skippedCount > 0 ? '（スキップ: ' + skippedCount + '件）' : '');
  console.log('approveSelectedRows 結果: ' + msg);
  ui.alert(msg);
}
```

---

## 変更点

1. `getActiveRange()` → `getActiveRangeList()` に変更
2. `getRanges()` で全ての選択範囲を取得
3. `Set` で重複する行番号を排除
4. 全ての選択行をループで処理

---

## 期待される効果

- 連続した複数行の選択に対応
- Ctrl+クリックで離れた行を選択した場合にも対応
- 行番号をドラッグして選択した場合にも対応

---

## 実装完了記録

| 日付 | 内容 | 担当 |
|------|------|------|
| | approveSelectedRows を getActiveRangeList() 対応に修正 | Claude Code |
