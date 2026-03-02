/**
 * MK004用クライアントラッパースクリプト
 * 
 * ライブラリ「ReceiptEngine」の全メニュー関数をラップ
 * 
 * ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 * 設定
 * ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 * 顧客コード: MK004
 * 設定はClientConfigシートで管理
 * ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 */

// ============================================================
// メニュー
// ============================================================

function onOpen() {
  ReceiptEngine.onOpen();
  
  // 通帳処理メニューを追加
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('通帳処理')
    .addItem('通帳読み込み開始', 'processPassbooks')
    .addSeparator()
    .addItem('通帳シートを作成', 'createPassbookSheetOnly')
    .addToUi();
}

// ============================================================
// レシート処理
// ============================================================

function processReceipts() {
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
// サイドバー用（google.script.runから呼ばれる）
// ============================================================

function getSelectedRowData() {
  return ReceiptEngine.getSelectedRowData();
}

// ============================================================
// データ修正
// ============================================================

function fixExistingAccountTitles() {
  ReceiptEngine.fixExistingAccountTitles();
}

function deleteSelectedFiles() {
  ReceiptEngine.deleteSelectedFiles();
}

function normalizeExistingStoreNames() {
  ReceiptEngine.normalizeExistingStoreNames();
}

// ============================================================
// 検証
// ============================================================

function verifySelectedRows() {
  ReceiptEngine.verifySelectedRows();
}

function applyVerificationFix(rowIndex, field, value) {
  return ReceiptEngine.applyVerificationFix(rowIndex, field, value);
}

function runAutoVerification() {
  ReceiptEngine.runAutoVerification();
}

function approveSelectedRows() {
  ReceiptEngine.approveSelectedRows();
}

// ============================================================
// 手書き領収証関連
// ============================================================

function simpleTest() {
  return ReceiptEngine.simpleTest();
}

function testHandwrittenUpdate() {
  return ReceiptEngine.testHandwrittenUpdate();
}

function updateHandReceipt(row, totalAmount, taxable10, tax10, taxable8, tax8, nonTaxable) {
  return ReceiptEngine.updateHandReceipt(row, totalAmount, taxable10, tax10, taxable8, tax8, nonTaxable);
}

function approveHandReceipt(row) {
  return ReceiptEngine.approveHandReceipt(row);
}

function testFromMenu() {
  return ReceiptEngine.testFromMenu();
}

function checkAuthStatus() {
  return ReceiptEngine.checkAuthStatus();
}

function showExecutionLog() {
  return ReceiptEngine.showExecutionLog();
}

// ============================================================
// 通帳処理
// ============================================================

/**
 * 通帳読み込みを実行
 * ClientConfigシートからPASSBOOK_FOLDER_IDを取得
 */
function processPassbooks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const folderId = ReceiptEngine.getPassbookFolderId_();
  
  if (!folderId) {
    console.error('エラー: PASSBOOK_FOLDER_IDが設定されていません。ClientConfigシートを確認してください。');
    return;
  }
  
  try {
    const count = ReceiptEngine.processPassbookFolder(folderId, ss);
    console.log('処理完了: ' + count + '件の通帳を処理しました。');
  } catch (e) {
    console.error('エラー: ' + e.message);
  }
}

/**
 * 通帳シートのみを作成
 */
function createPassbookSheetOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ReceiptEngine.createPassbookSheet(ss);
  console.log('通帳シートを作成しました。');
}

// ============================================================
// ClientConfig関連
// ============================================================

/**
 * ClientConfigシートを作成
 */
function createClientConfigSheet() {
  ReceiptEngine.createClientConfigSheet();
}
