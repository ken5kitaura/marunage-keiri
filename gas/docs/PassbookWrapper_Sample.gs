/**
 * 顧客スプシ用ラッパー関数（通帳読み込み）
 * 
 * このファイルは各顧客スプシ（MKxxx）のGASにコピーして使用する
 * 
 * ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 * 設定方法
 * ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 * 
 * 1. PassbookEngineをライブラリとして追加
 *    - スクリプトID: [PassbookEngineのスクリプトID]
 *    - 識別子: PassbookEngine
 * 
 * 2. 以下の関数を顧客スプシのGASに追加
 * 
 * 3. ScriptPropertiesにGEMINI_API_KEYを設定
 * 
 * 4. PASSBOOK_FOLDER_IDを設定
 */

// ============================================================
// 設定
// ============================================================

// 通帳画像フォルダID（各顧客で変更）
const PASSBOOK_FOLDER_ID = '';  // ←ここに通帳フォルダIDを入力

// ============================================================
// メニュー
// ============================================================

/**
 * スプレッドシート起動時のメニュー追加（既存のonOpenに追加）
 */
function addPassbookMenu() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('通帳処理')
    .addItem('通帳読み込み開始', 'processPassbooks')
    .addSeparator()
    .addItem('通帳シートを作成', 'createPassbookSheetOnly')
    .addToUi();
}

// ============================================================
// メイン処理
// ============================================================

/**
 * 通帳読み込みを実行
 */
function processPassbooks() {
  if (!PASSBOOK_FOLDER_ID) {
    SpreadsheetApp.getUi().alert('PASSBOOK_FOLDER_ID が設定されていません。');
    return;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    const count = PassbookEngine.processPassbookFolder(PASSBOOK_FOLDER_ID, ss);
    SpreadsheetApp.getUi().alert('処理完了: ' + count + '件の通帳を処理しました。');
  } catch (e) {
    SpreadsheetApp.getUi().alert('エラー: ' + e.message);
  }
}

/**
 * 通帳シートのみを作成
 */
function createPassbookSheetOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PassbookEngine.createPassbookSheet(ss);
  SpreadsheetApp.getUi().alert('通帳シートを作成しました。');
}
