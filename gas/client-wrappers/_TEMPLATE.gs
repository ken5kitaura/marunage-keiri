/**
 * クライアント用ラッパースクリプト テンプレート
 * 
 * ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 * このファイルは中央管理GASの「ラッパーGAS一括配布」で
 * 各顧客スプシに自動配布されます。
 * 
 * 手動で設定する場合:
 * 1. この内容を顧客スプシのGASエディタに貼り付け
 * 2. ライブラリ「ReceiptEngine」を追加（開発モード）
 * ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 */

// ============================================================
// メニュー
// ============================================================

function onOpen() {
  ReceiptEngine.onOpen();
}

// ============================================================
// レシート処理
// ============================================================

function processReceipts() {
  ReceiptEngine.processReceipts();
}

function processReceipts_continue() {
  ReceiptEngine.processReceipts();
}

function resetCheckErrorMarks() {
  ReceiptEngine.resetCheckErrorMarks();
}

function resetProcessedMarks() {
  ReceiptEngine.resetProcessedMarks();
}

function showSidebar() {
  ReceiptEngine.showSidebar();
}

// ============================================================
// 検証・承認
// ============================================================

function runAutoVerification() {
  ReceiptEngine.runAutoVerification();
}

function runAutoVerification_continue() {
  ReceiptEngine.runAutoVerification();
}

function verifySelectedRows() {
  ReceiptEngine.verifySelectedRows();
}

function approveSelectedRows() {
  ReceiptEngine.approveSelectedRows();
}

function applyVerificationFix(rowIndex, field, value) {
  return ReceiptEngine.applyVerificationFix(rowIndex, field, value);
}

// ============================================================
// サイドバー用（google.script.runから呼ばれる）
// ============================================================

function getSelectedRowData() {
  return ReceiptEngine.getSelectedRowData();
}

function updateHandReceipt(row, totalAmount, taxable10, tax10, taxable8, tax8, nonTaxable) {
  return ReceiptEngine.updateHandReceipt(row, totalAmount, taxable10, tax10, taxable8, tax8, nonTaxable);
}

function approveHandReceipt(row) {
  return ReceiptEngine.approveHandReceipt(row);
}

// ============================================================
// データ修正
// ============================================================

function fixExistingAccountTitles() {
  ReceiptEngine.fixExistingAccountTitles();
}

function normalizeExistingStoreNames() {
  ReceiptEngine.normalizeExistingStoreNames();
}

function deleteSelectedFiles() {
  ReceiptEngine.deleteSelectedFiles();
}

// ============================================================
// 設定
// ============================================================

function setupConfigSheet() {
  ReceiptEngine.setupConfigSheet();
}

function createConfigFoldersSheet() {
  ReceiptEngine.createConfigFoldersSheet();
}

function createConfigMappingSheet() {
  ReceiptEngine.createConfigMappingSheet();
}

function promptGeminiApiKey() {
  ReceiptEngine.promptGeminiApiKey();
}

// ============================================================
// クレカ突合
// ============================================================

function reconcileWithStatements() {
  ReceiptEngine.reconcileWithStatements();
}

function resetReconcileInfo() {
  ReceiptEngine.resetReconcileInfo();
}

function promptStatementSpreadsheetId() {
  ReceiptEngine.promptStatementSpreadsheetId();
}

// ============================================================
// エクスポート
// ============================================================

function exportToYayoiCSV() {
  ReceiptEngine.exportToYayoiCSV();
}

// ============================================================
// 一括読み込み
// ============================================================

function processAll() {
  ReceiptEngine.processAll();
}

// ============================================================
// 通帳処理
// ============================================================

function processPassbooks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const folderId = ReceiptEngine.getPassbookFolderId();

  if (!folderId) {
    SpreadsheetApp.getUi().alert('エラー: 通帳フォルダIDが設定されていません。\nClientConfigシートを確認してください。');
    return;
  }

  try {
    const count = ReceiptEngine.processPassbookFolder(folderId, ss);
    SpreadsheetApp.getUi().alert('完了: ' + count + '件の通帳を処理しました。');
  } catch (e) {
    SpreadsheetApp.getUi().alert('エラー: ' + e.message);
  }
}

function createPassbookSheetOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ReceiptEngine.createPassbookSheet(ss);
  SpreadsheetApp.getUi().alert('通帳シートを作成しました。');
}

// ============================================================
// ClientConfig
// ============================================================

function createClientConfigSheet() {
  ReceiptEngine.createClientConfigSheet();
}
