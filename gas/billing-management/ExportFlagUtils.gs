/**
 * ==========================================
 * 出力フラグ管理ユーティリティ
 * ==========================================
 * 
 * 顧客スプシの「出力済」「出力日」「出力行数」を管理
 */

/**
 * レシートシートにエクスポート列を追加
 */
function addExportColumnsToReceiptSheet(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // 既に存在するかチェック
  if (headers.includes('出力済')) {
    return { exported: headers.indexOf('出力済') + 1 };
  }
  
  // 新しい列を追加（U, V, W列）
  const lastCol = sheet.getLastColumn();
  sheet.getRange(1, lastCol + 1, 1, 3).setValues([['出力済', '出力日', '出力行数']]);
  
  return {
    exported: lastCol + 1,
    exportDate: lastCol + 2,
    exportRows: lastCol + 3
  };
}

/**
 * 通帳シートにエクスポート列を追加
 */
function addExportColumnsToPassbookSheet(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // 既に存在するかチェック
  if (headers.includes('出力済')) {
    return { exported: headers.indexOf('出力済') + 1 };
  }
  
  // 新しい列を追加
  const lastCol = sheet.getLastColumn();
  sheet.getRange(1, lastCol + 1, 1, 3).setValues([['出力済', '出力日', '出力行数']]);
  
  return {
    exported: lastCol + 1,
    exportDate: lastCol + 2,
    exportRows: lastCol + 3
  };
}

/**
 * レシート行の出力フラグをセット
 * @param {Sheet} sheet - 本番シート
 * @param {number} row - 行番号（1-indexed）
 * @param {object} colIndex - 列インデックス
 */
function setReceiptExportFlag(sheet, row, colIndex) {
  const data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // 出力行数を計算
  const amount10 = data[8] || 0;  // I列: 対象額(10%)
  const amount8 = data[10] || 0;  // K列: 対象額(8%)
  const taxFree = data[12] || 0;  // M列: 不課税
  
  let exportRows = 0;
  if (amount10 > 0) exportRows++;
  if (amount8 > 0) exportRows++;
  if (taxFree > 0) exportRows++;
  
  // 最低1行
  if (exportRows === 0) exportRows = 1;
  
  // フラグをセット
  sheet.getRange(row, colIndex.exported).setValue(true);
  sheet.getRange(row, colIndex.exportDate).setValue(new Date());
  sheet.getRange(row, colIndex.exportRows).setValue(exportRows);
  
  return exportRows;
}

/**
 * 通帳行の出力フラグをセット
 * @param {Sheet} sheet - 通帳シート
 * @param {number} row - 行番号（1-indexed）
 * @param {object} colIndex - 列インデックス
 */
function setPassbookExportFlag(sheet, row, colIndex) {
  sheet.getRange(row, colIndex.exported).setValue(true);
  sheet.getRange(row, colIndex.exportDate).setValue(new Date());
  sheet.getRange(row, colIndex.exportRows).setValue(1);
  
  return 1;
}

/**
 * 出力フラグをリセット（修正時に呼び出す）
 * @param {Sheet} sheet - シート
 * @param {number} row - 行番号（1-indexed）
 * @param {object} colIndex - 列インデックス
 */
function resetExportFlag(sheet, row, colIndex) {
  sheet.getRange(row, colIndex.exported).setValue(false);
  sheet.getRange(row, colIndex.exportDate).setValue('');
  sheet.getRange(row, colIndex.exportRows).setValue('');
}

/**
 * レシートシートの全行に出力フラグをセット（一括出力用）
 * @param {Sheet} sheet - 本番シート
 * @param {number[]} rows - 行番号の配列（1-indexed）
 * @returns {number} 出力行数の合計
 */
function setReceiptExportFlagBatch(sheet, rows) {
  const colIndex = addExportColumnsToReceiptSheet(sheet);
  let totalRows = 0;
  
  for (const row of rows) {
    totalRows += setReceiptExportFlag(sheet, row, colIndex);
  }
  
  return totalRows;
}

/**
 * 通帳シートの全行に出力フラグをセット（一括出力用）
 * @param {Sheet} sheet - 通帳シート
 * @param {number[]} rows - 行番号の配列（1-indexed）
 * @returns {number} 出力行数の合計
 */
function setPassbookExportFlagBatch(sheet, rows) {
  const colIndex = addExportColumnsToPassbookSheet(sheet);
  let totalRows = 0;
  
  for (const row of rows) {
    totalRows += setPassbookExportFlag(sheet, row, colIndex);
  }
  
  return totalRows;
}

/**
 * データ修正を検知してフラグをリセット
 * ※ 出力処理時に呼び出して、前回値と比較する
 * @param {Sheet} sheet - シート
 * @param {number} row - 行番号
 * @param {object} currentData - 現在のデータ
 * @param {object} previousData - 前回出力時のデータ（ハッシュ等）
 */
function checkAndResetIfModified(sheet, row, currentData, previousData, colIndex) {
  // 簡易的に金額系の列を比較
  const isModified = JSON.stringify(currentData) !== JSON.stringify(previousData);
  
  if (isModified) {
    resetExportFlag(sheet, row, colIndex);
    return true;
  }
  
  return false;
}
