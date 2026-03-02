/**
 * ==========================================
 * まるなげ経理 - 請求管理GAS
 * ==========================================
 * 
 * 機能:
 * - 全税理士法人の顧客スプシを巡回
 * - 月次の出力行数を集計
 * - 請求管理シートに記録
 */

// ========== 定数 ==========
const CONFIG = {
  // シート名
  SHEET_CLIENTS: '税理士一覧',
  SHEET_MONTHLY: '月次集計',
  SHEET_BILLING_HISTORY: '請求履歴',
  
  // 各顧客スプシの列構成
  CUSTOMER_SHEET: {
    NAME: '本番シート',  // レシート用シート名
    PASSBOOK_NAME: '通帳',
    COL: {
      // 本番シート（レシート）
      DATE: 5,           // E: 日付
      AMOUNT_10: 9,      // I: 対象額(10%)
      AMOUNT_8: 11,      // K: 対象額(8%)
      TAX_FREE: 13,      // M: 不課税
      EXPORTED: 21,      // U: 出力済
      EXPORT_DATE: 22,   // V: 出力日
      EXPORT_ROWS: 23,   // W: 出力行数
      
      // 通帳
      PB_DATE: 1,        // A: 日付（通帳）
      PB_EXPORTED: null, // 後で設定
      PB_EXPORT_DATE: null,
      PB_EXPORT_ROWS: null
    }
  },
  
  // 税理士一覧シートの列構成
  CLIENT_SHEET_COL: {
    NAME: 1,             // A: 税理士法人名
    SHEET_ID: 2,         // B: 顧客管理シートID
    CODE_PREFIX: 3,      // C: 顧客コードプレフィックス
    ACTIVE: 4,           // D: 有効フラグ
    MONTHLY_FEE: 5,      // E: 月額基本料金
    UNIT_PRICE: 6,       // F: 単価（円/行）
    MEMO: 7              // G: メモ
  }
};

// ========== メニュー ==========

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔧 請求管理')
    .addItem('📊 月次集計を実行', 'runMonthlyAggregation')
    .addItem('📊 指定月の集計', 'runMonthlyAggregationPrompt')
    .addSeparator()
    .addItem('⚙️ 初期セットアップ', 'initialSetup')
    .addToUi();
}

// ========== 初期セットアップ ==========

function initialSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // 税理士一覧シート
  let clientSheet = ss.getSheetByName(CONFIG.SHEET_CLIENTS);
  if (!clientSheet) {
    clientSheet = ss.insertSheet(CONFIG.SHEET_CLIENTS);
    const headers = ['税理士法人名', '顧客管理シートID', 'コードプレフィックス', '有効', '月額基本料金', '単価(円/行)', 'メモ'];
    clientSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    clientSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    clientSheet.setFrozenRows(1);
    
    // サンプルデータ
    clientSheet.getRange(2, 1, 1, 7).setValues([[
      '絆パートナーズ税理士法人',
      '1w8KfoYs6RFjNM6LZvH9qoD5e5Ow_vB6DrztsUF7nKzg',
      'KZ',
      true,
      0,
      20,
      ''
    ]]);
  }
  
  // 月次集計シート
  let monthlySheet = ss.getSheetByName(CONFIG.SHEET_MONTHLY);
  if (!monthlySheet) {
    monthlySheet = ss.insertSheet(CONFIG.SHEET_MONTHLY);
    const headers = ['税理士法人名', 'コード', '年月', 'レシート行数', '通帳行数', '合計行数', '単価', '金額', '集計日時'];
    monthlySheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    monthlySheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    monthlySheet.setFrozenRows(1);
  }
  
  // 請求履歴シート（将来用）
  let historySheet = ss.getSheetByName(CONFIG.SHEET_BILLING_HISTORY);
  if (!historySheet) {
    historySheet = ss.insertSheet(CONFIG.SHEET_BILLING_HISTORY);
    const headers = ['請求ID', '税理士法人名', '年月', '基本料金', '従量料金', '合計', '請求日', 'ステータス'];
    historySheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    historySheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    historySheet.setFrozenRows(1);
  }
  
  ui.alert('✅ 初期セットアップが完了しました');
}

// ========== 月次集計 ==========

/**
 * 当月の集計を実行
 */
function runMonthlyAggregation() {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1;
  
  runAggregationForMonth(year, month);
}

/**
 * 指定月の集計を実行（プロンプト）
 */
function runMonthlyAggregationPrompt() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    '指定月の集計',
    '年月を入力してください（例: 2025-01）:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const input = response.getResponseText();
  const match = input.match(/^(\d{4})-(\d{1,2})$/);
  
  if (!match) {
    ui.alert('形式が正しくありません。例: 2025-01');
    return;
  }
  
  const year = parseInt(match[1]);
  const month = parseInt(match[2]);
  
  runAggregationForMonth(year, month);
}

/**
 * 指定年月の集計を実行
 */
function runAggregationForMonth(year, month) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const clientSheet = ss.getSheetByName(CONFIG.SHEET_CLIENTS);
  const monthlySheet = ss.getSheetByName(CONFIG.SHEET_MONTHLY);
  
  if (!clientSheet || !monthlySheet) {
    ui.alert('❌ シートが見つかりません。初期セットアップを実行してください。');
    return;
  }
  
  const clientData = clientSheet.getDataRange().getValues();
  const yearMonth = `${year}-${String(month).padStart(2, '0')}`;
  const results = [];
  
  // 各税理士を処理
  for (let i = 1; i < clientData.length; i++) {
    const row = clientData[i];
    const clientName = row[CONFIG.CLIENT_SHEET_COL.NAME - 1];
    const sheetId = row[CONFIG.CLIENT_SHEET_COL.SHEET_ID - 1];
    const codePrefix = row[CONFIG.CLIENT_SHEET_COL.CODE_PREFIX - 1];
    const active = row[CONFIG.CLIENT_SHEET_COL.ACTIVE - 1];
    const unitPrice = row[CONFIG.CLIENT_SHEET_COL.UNIT_PRICE - 1] || 20;
    
    if (!active || !sheetId) continue;
    
    try {
      const counts = aggregateClientData(sheetId, codePrefix, year, month);
      
      results.push([
        clientName,
        codePrefix,
        yearMonth,
        counts.receipt,
        counts.passbook,
        counts.total,
        unitPrice,
        counts.total * unitPrice,
        new Date()
      ]);
      
      console.log(`${clientName}: レシート${counts.receipt}行, 通帳${counts.passbook}行`);
      
    } catch (e) {
      console.log(`${clientName}: エラー - ${e.message}`);
      results.push([
        clientName,
        codePrefix,
        yearMonth,
        'エラー',
        'エラー',
        'エラー',
        unitPrice,
        0,
        new Date()
      ]);
    }
  }
  
  // 既存の同月データを削除
  const existingData = monthlySheet.getDataRange().getValues();
  for (let i = existingData.length - 1; i >= 1; i--) {
    if (existingData[i][2] === yearMonth) {
      monthlySheet.deleteRow(i + 1);
    }
  }
  
  // 結果を書き込み
  if (results.length > 0) {
    monthlySheet.getRange(monthlySheet.getLastRow() + 1, 1, results.length, results[0].length)
      .setValues(results);
  }
  
  // 合計を計算
  let totalRows = 0;
  let totalAmount = 0;
  for (const r of results) {
    if (typeof r[5] === 'number') {
      totalRows += r[5];
      totalAmount += r[7];
    }
  }
  
  ui.alert(
    '✅ 集計完了',
    `${yearMonth} の集計結果:\n\n` +
    `税理士法人数: ${results.length}\n` +
    `合計行数: ${totalRows}\n` +
    `合計金額: ¥${totalAmount.toLocaleString()}`,
    ui.ButtonSet.OK
  );
}

/**
 * 1つの税理士の顧客管理シートを集計
 */
function aggregateClientData(managementSheetId, codePrefix, year, month) {
  const managementSS = SpreadsheetApp.openById(managementSheetId);
  const managementSheet = managementSS.getSheetByName('顧客管理');
  
  if (!managementSheet) {
    throw new Error('顧客管理シートが見つかりません');
  }
  
  const managementData = managementSheet.getDataRange().getValues();
  
  let totalReceipt = 0;
  let totalPassbook = 0;
  
  // 各顧客のスプシを処理
  for (let i = 1; i < managementData.length; i++) {
    const customerCode = managementData[i][4]; // E列: customer_code
    const spreadsheetUrl = managementData[i][11]; // L列: spreadsheet_url
    
    if (!spreadsheetUrl || !customerCode) continue;
    
    try {
      const ssId = extractSpreadsheetId(spreadsheetUrl);
      const customerSS = SpreadsheetApp.openById(ssId);
      
      // レシート集計
      const receiptSheet = customerSS.getSheetByName('本番シート');
      if (receiptSheet) {
        totalReceipt += countExportedRows(receiptSheet, year, month, 'receipt');
      }
      
      // 通帳集計
      const passbookSheet = customerSS.getSheetByName('通帳');
      if (passbookSheet) {
        totalPassbook += countExportedRows(passbookSheet, year, month, 'passbook');
      }
      
    } catch (e) {
      console.log(`${customerCode}: ${e.message}`);
    }
  }
  
  return {
    receipt: totalReceipt,
    passbook: totalPassbook,
    total: totalReceipt + totalPassbook
  };
}

/**
 * シートから出力済み行数をカウント
 */
function countExportedRows(sheet, year, month, type) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return 0;
  
  // 列位置を特定
  const headers = data[0];
  const dateColIndex = findColumnIndex(headers, ['日付', 'date']);
  const exportedColIndex = findColumnIndex(headers, ['出力済', 'exported']);
  const exportRowsColIndex = findColumnIndex(headers, ['出力行数', 'export_rows']);
  
  if (dateColIndex === -1 || exportedColIndex === -1) {
    return 0;
  }
  
  let count = 0;
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const date = row[dateColIndex];
    const exported = row[exportedColIndex];
    const exportRows = exportRowsColIndex !== -1 ? row[exportRowsColIndex] : 1;
    
    if (!exported || exported !== true) continue;
    
    // 日付チェック
    if (date instanceof Date) {
      if (date.getFullYear() === year && date.getMonth() + 1 === month) {
        count += (typeof exportRows === 'number' && exportRows > 0) ? exportRows : 1;
      }
    }
  }
  
  return count;
}

/**
 * ヘッダーから列インデックスを検索
 */
function findColumnIndex(headers, possibleNames) {
  for (let i = 0; i < headers.length; i++) {
    const header = String(headers[i]).toLowerCase().trim();
    for (const name of possibleNames) {
      if (header === name.toLowerCase()) {
        return i;
      }
    }
  }
  return -1;
}

/**
 * スプレッドシートURLからIDを抽出
 */
function extractSpreadsheetId(url) {
  const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (match) return match[1];
  throw new Error('Invalid spreadsheet URL');
}
