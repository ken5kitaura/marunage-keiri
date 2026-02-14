/**
 * Output_Yayoi.gs
 * 弥生会計インポート形式CSV出力
 *
 * 責務:
 * - 本番シートのデータを弥生インポート形式（27列）に変換
 * - 10%/8%/不課税を税率ごとに複数行へ分割
 * - 「弥生エクスポート」シートに書き出し
 * - UTF-8 BOM付きCSVファイルを生成・ダウンロード
 */

/**
 * 弥生CSV列定義（27列）
 */
const YAYOI_COLUMNS = [
  '識別フラグ',      // 0: 2000=仕訳
  '伝票No',          // 1: 空欄
  '決算',            // 2: 空欄
  '取引日付',        // 3: YYYYMMDD
  '借方勘定科目',    // 4
  '借方補助科目',    // 5
  '借方部門',        // 6
  '借方税区分',      // 7
  '借方金額',        // 8
  '借方税金額',      // 9: 空欄（弥生が自動計算）
  '貸方勘定科目',    // 10
  '貸方補助科目',    // 11
  '貸方部門',        // 12
  '貸方税区分',      // 13: 対象外
  '貸方金額',        // 14
  '貸方税金額',      // 15
  '摘要',            // 16
  '番号',            // 17
  '期日',            // 18
  'タイプ',          // 19: 0
  '生成元',          // 20
  '仕訳メモ',        // 21
  '付箋１',          // 22
  '付箋２',          // 23
  '調整',            // 24: no
  '借方取引先名',    // 25
  '貸方取引先名'     // 26
];

/**
 * 税区分コード
 */
const TAX_CODES = {
  TAXABLE_10: '課対仕入10%',
  TAXABLE_8:  '課対仕入8%（軽）',
  EXEMPT:     '対象外'
};

/**
 * 弥生エクスポートシート名
 */
const YAYOI_SHEET_NAME = '弥生エクスポート';

// ============================================================
// メイン関数
// ============================================================

/**
 * 弥生CSVを生成：シート書き出し + CSVダウンロード
 */
function generateYayoiCSV() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(CONFIG.SHEET_NAME.MAIN);

  if (!mainSheet) {
    ui.alert('「' + CONFIG.SHEET_NAME.MAIN + '」シートが見つかりません。');
    return;
  }

  const lastRow = mainSheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('データがありません。');
    return;
  }

  // ヘッダー取得
  const lastCol = mainSheet.getLastColumn();
  const header = mainSheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // 列インデックス取得（現行シートヘッダーに対応）
  const idx = {
    status:     findHeaderIndex(header, ['Status']),
    date:       findHeaderIndex(header, ['日付']),
    store:      findHeaderIndex(header, ['利用店舗名', '店舗名']),
    total:      findHeaderIndex(header, ['総合計']),
    subtotal10: findHeaderIndex(header, ['対象額(10%)']),
    tax10:      findHeaderIndex(header, ['消費税(10%)']),
    subtotal8:  findHeaderIndex(header, ['対象額(8%)']),
    tax8:       findHeaderIndex(header, ['消費税(8%)']),
    nonTaxable: findHeaderIndex(header, ['不課税']),
    account:    findHeaderIndex(header, ['勘定科目']),
    credit:     findHeaderIndex(header, ['貸方科目'])
  };

  // データ取得
  const data = mainSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  // 弥生形式の行を生成
  const yayoiRows = [];

  for (const row of data) {
    // ステータスチェック（OK/COMPOUNDのみ出力）
    const status = String(row[idx.status] || '');
    if (!status.includes('OK') && !status.includes('COMPOUND')) {
      continue;
    }

    const date = parseDate(row[idx.date]);
    if (!date) continue;

    const dateStr = formatDateYayoiExport_(date);
    const storeName = String(row[idx.store] || '');
    const accountTitle = String(row[idx.account] || '消耗品費');
    const creditAccount = String(row[idx.credit] || '現金');

    const subtotal10 = parseAmount(row[idx.subtotal10]) || 0;
    const tax10 = parseAmount(row[idx.tax10]) || 0;
    const subtotal8 = parseAmount(row[idx.subtotal8]) || 0;
    const tax8 = parseAmount(row[idx.tax8]) || 0;
    const nonTaxable = parseAmount(row[idx.nonTaxable]) || 0;

    // 10%課税行
    if (subtotal10 > 0) {
      const taxIncluded10 = subtotal10 + tax10;
      yayoiRows.push(createYayoiRow_({
        date: dateStr,
        debitAccount: accountTitle,
        debitTaxCode: TAX_CODES.TAXABLE_10,
        debitAmount: taxIncluded10,
        creditAccount: creditAccount,
        creditAmount: taxIncluded10,
        summary: storeName
      }));
    }

    // 8%軽減税率行
    if (subtotal8 > 0) {
      const taxIncluded8 = subtotal8 + tax8;
      yayoiRows.push(createYayoiRow_({
        date: dateStr,
        debitAccount: accountTitle,
        debitTaxCode: TAX_CODES.TAXABLE_8,
        debitAmount: taxIncluded8,
        creditAccount: creditAccount,
        creditAmount: taxIncluded8,
        summary: storeName
      }));
    }

    // 不課税行（入湯税等）
    if (nonTaxable > 0) {
      yayoiRows.push(createYayoiRow_({
        date: dateStr,
        debitAccount: '租税公課',
        debitTaxCode: TAX_CODES.EXEMPT,
        debitAmount: nonTaxable,
        creditAccount: creditAccount,
        creditAmount: nonTaxable,
        summary: storeName + '（入湯税等）'
      }));
    }

    // 10%も8%も不課税もない場合（総合計のみ）→ 10%として出力
    if (subtotal10 === 0 && subtotal8 === 0 && nonTaxable === 0) {
      const totalAmount = parseAmount(row[idx.total]) || 0;
      if (totalAmount > 0) {
        yayoiRows.push(createYayoiRow_({
          date: dateStr,
          debitAccount: accountTitle,
          debitTaxCode: TAX_CODES.TAXABLE_10,
          debitAmount: totalAmount,
          creditAccount: creditAccount,
          creditAmount: totalAmount,
          summary: storeName
        }));
      }
    }
  }

  if (yayoiRows.length === 0) {
    ui.alert('出力対象のデータがありません。\nステータスがOKまたはCOMPOUNDの行が必要です。');
    return;
  }

  // ── Step 1: 弥生エクスポートシートに書き出し ──
  writeToYayoiSheet_(ss, yayoiRows);

  // ── Step 2: UTF-8 BOM付きCSVファイルを生成 ──
  const csvRows = [YAYOI_COLUMNS].concat(yayoiRows);
  const csvContent = '\uFEFF' + csvRows.map(toCSVRow).join('\r\n');
  const fileName = 'yayoi_export_' + Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss') + '.csv';
  const blob = Utilities.newBlob(csvContent, 'text/csv; charset=UTF-8', fileName);
  const file = DriveApp.createFile(blob);
  const downloadUrl = file.getDownloadUrl();

  // 結果表示
  const html = HtmlService.createHtmlOutput(
    '<p>弥生CSV出力が完了しました。</p>' +
    '<p>仕訳行数: <b>' + yayoiRows.length + '件</b></p>' +
    '<p>「' + YAYOI_SHEET_NAME + '」シートにも書き出しました。</p>' +
    '<hr>' +
    '<p><a href="' + downloadUrl + '" target="_blank" style="font-size:16px;">CSVをダウンロード</a></p>' +
    '<p style="color:#666;font-size:12px;">ファイル名: ' + fileName + '</p>'
  ).setWidth(420).setHeight(220);

  ui.showModalDialog(html, '弥生CSV出力完了');
}

// ============================================================
// ヘルパー関数
// ============================================================

/**
 * 弥生エクスポートシートにデータを書き出す
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {Array<Array>} yayoiRows
 */
function writeToYayoiSheet_(ss, yayoiRows) {
  let sheet = ss.getSheetByName(YAYOI_SHEET_NAME);

  if (sheet) {
    // 既存シートをクリア
    sheet.clearContents();
  } else {
    // 新規作成
    sheet = ss.insertSheet(YAYOI_SHEET_NAME);
  }

  // ヘッダー書き込み
  sheet.getRange(1, 1, 1, YAYOI_COLUMNS.length).setValues([YAYOI_COLUMNS]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, YAYOI_COLUMNS.length).setFontWeight('bold');

  // データ書き込み
  if (yayoiRows.length > 0) {
    sheet.getRange(2, 1, yayoiRows.length, YAYOI_COLUMNS.length).setValues(yayoiRows);
  }

  // 列幅調整（主要列のみ）
  sheet.setColumnWidth(4, 100);  // 取引日付
  sheet.setColumnWidth(5, 130);  // 借方勘定科目
  sheet.setColumnWidth(8, 110);  // 借方税区分
  sheet.setColumnWidth(9, 90);   // 借方金額
  sheet.setColumnWidth(11, 130); // 貸方勘定科目
  sheet.setColumnWidth(14, 110); // 貸方税区分
  sheet.setColumnWidth(17, 200); // 摘要
}

/**
 * 弥生CSV1行を作成（27列）
 * @param {Object} params
 * @return {Array}
 */
function createYayoiRow_(params) {
  return [
    '2000',                        // [0]  識別フラグ
    '',                            // [1]  伝票No
    '',                            // [2]  決算
    params.date,                   // [3]  取引日付 (YYYYMMDD)
    params.debitAccount,           // [4]  借方勘定科目
    '',                            // [5]  借方補助科目
    '',                            // [6]  借方部門
    params.debitTaxCode,           // [7]  借方税区分
    params.debitAmount,            // [8]  借方金額
    '',                            // [9]  借方税金額（弥生が自動計算）
    params.creditAccount,          // [10] 貸方勘定科目
    '',                            // [11] 貸方補助科目
    '',                            // [12] 貸方部門
    TAX_CODES.EXEMPT,              // [13] 貸方税区分 = 対象外
    params.creditAmount,           // [14] 貸方金額
    '',                            // [15] 貸方税金額
    params.summary,                // [16] 摘要
    '',                            // [17] 番号
    '',                            // [18] 期日
    '0',                           // [19] タイプ
    '',                            // [20] 生成元
    '',                            // [21] 仕訳メモ
    '',                            // [22] 付箋１
    '',                            // [23] 付箋２
    'no',                          // [24] 調整
    '',                            // [25] 借方取引先名
    ''                             // [26] 貸方取引先名
  ];
}

/**
 * 日付をYYYYMMDD形式にフォーマット（弥生エクスポート用）
 * @param {Date} date
 * @return {string}
 */
function formatDateYayoiExport_(date) {
  if (!date || !(date instanceof Date)) return '';
  const y = date.getFullYear();
  const m = ('0' + (date.getMonth() + 1)).slice(-2);
  const d = ('0' + date.getDate()).slice(-2);
  return String(y) + m + d;
}
