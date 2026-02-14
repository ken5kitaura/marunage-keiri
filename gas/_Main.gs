/**
 * _Main.gs
 * メインエントリポイント
 *
 * 処理フロー:
 * 1. Driveフォルダからファイル取得
 * 2. Service_OCR: Gemini APIで画像→OCRResult
 * 3. Logic_Accounting: OCRResult→AccountingData（税計算・科目判定）
 * 4. スプレッドシートに出力
 * 5. (オプション) Service_Reconcile: クレカ明細と突合
 * 6. (オプション) Output_Yayoi: 弥生CSV出力
 */

/**
 * スプレッドシート起動時のメニュー作成
 */
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

/**
 * Gemini APIキー設定ダイアログ
 */
function promptGeminiApiKey() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'Gemini APIキー設定',
    'Gemini APIキーを入力してください:',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() === ui.Button.OK) {
    const key = result.getResponseText().trim();
    if (key) {
      setConfig_('GEMINI_API_KEY', key);
      ui.alert('APIキーを保存しました。');
    }
  }
}

/**
 * 明細スプレッドシートID設定ダイアログ
 */
function promptStatementSpreadsheetId() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'クレカ明細スプレッドシート設定',
    'スプレッドシートのURLまたはIDを入力してください:',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() === ui.Button.OK) {
    let input = result.getResponseText().trim();
    if (input) {
      // URLからIDを抽出
      const match = input.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
      if (match && match[1]) {
        input = match[1];
      }
      setConfig_('CC_STATEMENT_SPREADSHEET_ID', input);
      ui.alert('設定を保存しました。\nID: ' + input);
    }
  }
}

// ============================================================
// メイン処理
// ============================================================

/**
 * レシート処理メイン（多重実行ガード付き）
 */
function processReceipts() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    console.log('Skip: already running');
    return;
  }

  try {
    const startTime = Date.now();
    const folderConfigs = loadFolderConfigs_();
    let processedCount = 0;
    let hasNext = false;

    console.log('バッチ処理を開始...');

    for (const folderConfig of folderConfigs) {
      // タイムアウトチェック
      if (Date.now() - startTime > CONFIG.PROCESSING.MAX_EXECUTION_TIME_MS) {
        console.log('タイムアウト: トリガーを設定して中断');
        scheduleContinuation_('processReceipts');
        return;
      }

      let folder;
      try {
        folder = DriveApp.getFolderById(folderConfig.folderId);
      } catch (e) {
        console.error('フォルダ取得失敗 (' + folderConfig.label + '): ' + e.message);
        continue;
      }

      console.log('フォルダ処理開始: ' + folderConfig.label);
      const files = folder.getFiles();

      while (files.hasNext()) {
        // タイムアウトチェック
        if (Date.now() - startTime > CONFIG.PROCESSING.MAX_EXECUTION_TIME_MS) {
          console.log('タイムアウト: トリガーを設定して中断');
          scheduleContinuation_('processReceipts');
          return;
        }

        const file = files.next();
        const fileName = file.getName();

        // 処理済みスキップ
        if (isProcessedFile_(fileName)) {
          continue;
        }

        // 対応MIMEタイプチェック
        const mime = file.getMimeType();
        if (!isSupportedMimeType_(mime)) {
          console.log('SKIP (非対応MIME): ' + fileName);
          continue;
        }

        // 重複チェック
        if (isDuplicateInSheet_(fileName)) {
          console.log('SKIP (重複): ' + fileName);
          continue;
        }

        hasNext = true;

        try {
          console.log('処理中: ' + fileName);
          processOneReceipt_(file, folderConfig);
          processedCount++;
        } catch (e) {
          console.error('処理エラー (' + fileName + '): ' + e.message);
          markFileAsError_(file, e.message);
        }
      }
    }

    if (processedCount === 0 && !hasNext) {
      try {
        SpreadsheetApp.getUi().alert('処理完了: 未処理ファイルはありません。');
      } catch (e) { /* UI非対応環境 */ }
    } else {
      console.log('処理完了: ' + processedCount + '件');
    }

  } finally {
    lock.releaseLock();
  }
}

/**
 * 1ファイルの処理
 * @param {GoogleAppsScript.Drive.File} file
 * @param {FolderConfig} folderConfig
 */
function processOneReceipt_(file, folderConfig) {
  const fileName = file.getName();
  const fileUrl = file.getUrl();

  // Step 1: OCR抽出（Gemini API）
  /** @type {OCRResult} */
  const ocrResult = extractOCR_(file);

  // Step 2: 会計データ生成（税計算・科目判定）
  /** @type {AccountingData} */
  const accountingData = calculateAccounting_(ocrResult);

  // Step 3: 科目判定（Gemini推定値+会計データも渡す）
  const accountTitle = inferAccountTitle_(ocrResult.storeName, ocrResult.items, ocrResult._suggestedAccountTitle, ocrResult);

  // Step 4: ステータス判定（整合性チェック付き）
  const statusResult = determineStatus_(ocrResult, accountingData);
  let status = statusResult.status;
  const debugInfo = statusResult.errors.length > 0 ? statusResult.errors.join(' | ') : '';

  // 手書き判定（Geminiの判定を優先）
  const isHandwritten = ocrResult.isHandwritten === true;
  if (isHandwritten) {
    status = 'HAND';
    console.log('手書き領収証と判定: ' + fileName);
  }

  // Step 5: 摘要欄の生成（EC店舗の場合は品名を付加）
  const summaryStoreName = buildSummaryStoreName_(ocrResult.storeName, ocrResult.items);

  // Step 6: ReceiptData構築
  /** @type {ReceiptData} */
  const receiptData = {
    id: generateId_(fileName),
    status: status,
    fileUrl: fileUrl,
    fileName: fileName,
    processedAt: new Date(),
    ocr: ocrResult,
    accounting: accountingData,
    date: ocrResult.date,
    storeName: summaryStoreName,
    accountTitle: accountTitle || '',  // 判定不能なら空欄（後で人間が埋める）
    creditAccount: folderConfig.creditAccount,
    folderLabel: folderConfig.label,
    reconcile: null,
    debugInfo: debugInfo
  };

  // Step 7: スプレッドシートに出力
  appendReceiptToSheet_(receiptData);

  // Step 8: ファイル名に処理済みマークを付与
  markFileAsProcessed_(file, status);

  console.log('完了: ' + fileName + ' (' + status + ')' + (debugInfo ? ' [' + debugInfo + ']' : ''));
}

// ============================================================
// ヘルパー関数
// ============================================================

/**
 * 処理済みファイルか判定
 * @param {string} fileName
 * @return {boolean}
 */
function isProcessedFile_(fileName) {
  // 新形式: [OK], [CHK], [CMP], [ERR], [HAND]
  // 旧形式: 絵文字プレフィックス（後方互換）
  // processed_ プレフィックスもチェック
  return /^\[(OK|CHK|CMP|ERR|HAND|\?)\]/.test(fileName) ||
         /^[🟢🔴🟡🟠]/.test(fileName) ||
         fileName.startsWith('processed_');
}

/**
 * 対応MIMEタイプか判定
 * @param {string} mime
 * @return {boolean}
 */
function isSupportedMimeType_(mime) {
  return mime === 'application/pdf' || mime.startsWith('image/');
}

/**
 * シート内で重複チェック
 * @param {string} fileName
 * @return {boolean}
 */
function isDuplicateInSheet_(fileName) {
  const sheet = getOrCreateMainSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return false;

  // ヘッダーから「ファイル名」列を動的に検索
  const lastCol = sheet.getLastColumn();
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const fileNameColIndex = findHeaderIndex(header, ['ファイル名']) + 1; // 1-indexed
  if (fileNameColIndex < 1) return false;

  const data = sheet.getRange(2, fileNameColIndex, lastRow - 1, 1).getValues();
  const names = new Set(data.map(r => String(r[0])));
  return names.has(fileName);
}

/**
 * ステータス判定（整合性チェック付き）
 * @param {OCRResult} ocr
 * @param {AccountingData} accounting
 * @return {{status: string, errors: Array<string>}}
 */
function determineStatus_(ocr, accounting) {
  const errors = [];

  // 基本チェック
  if (!ocr.date) {
    errors.push('DATE_MISSING');
  } else if (!/^\d{4}-\d{2}-\d{2}$/.test(ocr.date)) {
    // YYYY-MM-DD形式でない日付は危険（令和表記がそのまま通過した等）
    errors.push('DATE_FORMAT_INVALID: ' + ocr.date);
  }
  if (ocr.storeName === 'PARSE_ERROR' || ocr.storeName === 'API_ERROR') {
    errors.push('OCR_FAILED: ' + ocr.storeName);
  } else if (!ocr.storeName || ocr.storeName === 'UNKNOWN') {
    errors.push('STORE_MISSING');
  }
  if (!ocr.totalAmount || ocr.totalAmount <= 0) {
    errors.push('TOTAL_MISSING');
  }

  // 基本情報が欠けていればERROR
  if (errors.length > 0) {
    return { status: 'ERROR', errors: errors };
  }

  // 整合性チェック
  const validation = validateAccountingConsistency_(ocr, accounting);
  if (!validation.isValid) {
    errors.push(...validation.errors);
    return { status: 'CHECK', errors: errors };
  }

  // 値引き等による補正が行われた場合はCHECK
  if (accounting.adjustmentNote) {
    return { status: 'CHECK', errors: [accounting.adjustmentNote] };
  }

  // 複合仕訳（入湯税等がある場合）
  if (accounting.isCompound) {
    return { status: 'COMPOUND', errors: [] };
  }

  return { status: 'OK', errors: [] };
}

/**
 * 会計データの整合性チェック
 * @param {OCRResult} ocr
 * @param {AccountingData} accounting
 * @return {{isValid: boolean, errors: Array<string>}}
 */
function validateAccountingConsistency_(ocr, accounting) {
  const errors = [];
  const totalAmount = ocr.totalAmount || 0;
  const subtotal = accounting.subtotal10 + accounting.subtotal8;
  const tax = accounting.tax10 + accounting.tax8;

  // 1. 税抜合計が0なのに総額がある場合
  if (totalAmount > 0 && subtotal === 0 && tax === 0 && accounting.rawNonTaxable === 0) {
    errors.push('SUBTOTAL_ZERO: 税抜合計が読み取れていません');
  }

  // 2. 計算された総額と実際の総額の差異チェック
  // COMPOUND（不課税あり）: 全額がレシートに明記されているため完全一致を要求
  // 通常: 内税逆算の丸め誤差があるため±3円まで許容
  const calculatedTotal = subtotal + tax + accounting.rawNonTaxable;
  if (calculatedTotal > 0) {
    const delta = Math.abs(calculatedTotal - totalAmount);
    const tolerance = accounting.isCompound ? 0 : 3;
    if (delta > tolerance) {
      errors.push('TOTAL_MISMATCH: 差異=' + delta + '円');
    }
  }

  // 3. 税率整合性チェック（税抜がある場合のみ）
  // ★小額の場合は丸め誤差が大きいため、絶対値で許容
  if (accounting.subtotal10 > 0 && accounting.tax10 > 0) {
    const ratio10 = accounting.tax10 / accounting.subtotal10;
    const expectedTax10 = Math.round(accounting.subtotal10 * 0.1);
    const taxDiff10 = Math.abs(accounting.tax10 - expectedTax10);

    // 小額（100円以下）は±2円まで許容、それ以外は比率チェック
    if (accounting.subtotal10 <= 100) {
      if (taxDiff10 > 2) {
        errors.push('TAX10_RATE_MISMATCH: 期待' + expectedTax10 + '円, 実際' + accounting.tax10 + '円');
      }
    } else {
      if (ratio10 < 0.08 || ratio10 > 0.12) {
        errors.push('TAX10_RATE_MISMATCH: ' + (ratio10 * 100).toFixed(1) + '%');
      }
    }
  }
  if (accounting.subtotal8 > 0 && accounting.tax8 > 0) {
    const ratio8 = accounting.tax8 / accounting.subtotal8;
    const expectedTax8 = Math.round(accounting.subtotal8 * 0.08);
    const taxDiff8 = Math.abs(accounting.tax8 - expectedTax8);

    // 小額（100円以下）は±2円まで許容、それ以外は比率チェック
    if (accounting.subtotal8 <= 100) {
      if (taxDiff8 > 2) {
        errors.push('TAX8_RATE_MISMATCH: 期待' + expectedTax8 + '円, 実際' + accounting.tax8 + '円');
      }
    } else {
      if (ratio8 < 0.06 || ratio8 > 0.10) {
        errors.push('TAX8_RATE_MISMATCH: ' + (ratio8 * 100).toFixed(1) + '%');
      }
    }
  }

  // 4. 負の値チェック
  if (accounting.subtotal10 < 0 || accounting.tax10 < 0 ||
      accounting.subtotal8 < 0 || accounting.tax8 < 0 ||
      accounting.rawNonTaxable < 0) {
    errors.push('NEGATIVE_VALUE');
  }

  return {
    isValid: errors.length === 0,
    errors: errors
  };
}

/**
 * 一意ID生成
 * @param {string} fileName
 * @return {string}
 */
function generateId_(fileName) {
  const timestamp = new Date().getTime();
  const hash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    fileName + timestamp
  ).map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('').slice(0, 8);
  return hash;
}

/**
 * ファイルに処理済みマークを付与
 * @param {GoogleAppsScript.Drive.File} file
 * @param {string} status
 */
function markFileAsProcessed_(file, status) {
  // ASCII文字のプレフィックスを使用（絵文字は文字化けの原因になる）
  const prefix = {
    'OK': '[OK]',
    'CHECK': '[CHK]',
    'COMPOUND': '[CMP]',
    'ERROR': '[ERR]',
    'HAND': '[HAND]'
  }[status] || '[?]';

  const currentName = file.getName();
  if (!isProcessedFile_(currentName)) {
    file.setName(prefix + currentName);
  }
}

/**
 * ファイルにエラーマークを付与
 * @param {GoogleAppsScript.Drive.File} file
 * @param {string} errorMsg
 */
function markFileAsError_(file, errorMsg) {
  const currentName = file.getName();
  if (!isProcessedFile_(currentName)) {
    file.setName('[ERR]' + currentName);
  }
}

/**
 * 摘要用の店舗名を生成
 * EC店舗（Amazon, 楽天等）の場合は品名を付加する
 * 例: "Amazon" + [{name: "USBケーブル 3本セット"}] → "Amazon USBケーブル 3本セット"
 *
 * @param {string} storeName - 店舗名
 * @param {Array<OCRLineItem>} items - 明細行
 * @return {string} 摘要用の店舗名
 */
function buildSummaryStoreName_(storeName, items) {
  const name = String(storeName || '');

  // EC店舗の判定（Amazon, 楽天 等）
  // ただし "Amazon Web Services" 等のIT系は除外
  const isEC = /^(amazon|楽天)$/i.test(name.trim()) ||
               (/amazon/i.test(name) && !/web\s*services|aws/i.test(name));

  if (!isEC) {
    return name;
  }

  // 品名を取得（最初の品目名を使用、長すぎる場合は30文字で切る）
  if (!items || items.length === 0) {
    return name;
  }

  const firstItemName = String(items[0].name || '').trim();
  if (!firstItemName) {
    return name;
  }

  // 品名が長すぎる場合は切り詰め
  const maxLen = 30;
  const itemSummary = firstItemName.length > maxLen
    ? firstItemName.slice(0, maxLen) + '…'
    : firstItemName;

  return name + ' ' + itemSummary;
}

/**
 * 継続処理のトリガーをスケジュール
 * @param {string} handlerName
 */
function scheduleContinuation_(handlerName) {
  // 既存の同一トリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === handlerName) {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 新しいトリガーを作成
  ScriptApp.newTrigger(handlerName)
    .timeBased()
    .after(CONFIG.PROCESSING.RETRY_DELAY_MINUTES * 60 * 1000)
    .create();
}

/**
 * メインシートを取得または作成
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateMainSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME.MAIN);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME.MAIN);
    // ヘッダーを設定（完全1行管理：10%/8%/不課税を分離）
    const headers = [
      'Status',           // 1: ステータスアイコン
      '画像確認',          // 2: ファイルURLへのリンク
      '処理日時',          // 3: processedAt
      '日付',             // 4: date
      '利用店舗名',        // 5: storeName
      '登録番号',          // 6: ocr.invoiceNumber
      '総合計',           // 7: accounting.totalAmount (税込総額)
      '対象額(10%)',      // 8: accounting.subtotal10 (税抜)
      '消費税(10%)',      // 9: accounting.tax10
      '対象額(8%)',       // 10: accounting.subtotal8 (税抜)
      '消費税(8%)',       // 11: accounting.tax8
      '不課税',           // 12: accounting.rawNonTaxable (入湯税など)
      '勘定科目',          // 13: accountTitle
      '貸方科目',          // 14: creditAccount
      'ファイル名',        // 15: fileName
      'Debug'            // 16: debugInfo
    ];
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }

  return sheet;
}

/**
 * ReceiptDataをシートに出力
 * @param {ReceiptData} data
 */
function appendReceiptToSheet_(data) {
  const sheet = getOrCreateMainSheet_();
  const acc = data.accounting;

  // ステータスアイコン
  const statusIcon = {
    'OK': '🟢OK',
    'CHECK': '🔴CHECK',
    'COMPOUND': '🟡COMPOUND',
    'ERROR': '🔴ERROR',
    'HAND': '🖊️HAND'
  }[data.status] || '🟡UNKNOWN';

  // 完全1行管理：10%/8%/不課税を分離して出力
  const row = [
    statusIcon,                                          // 1: Status
    '=HYPERLINK("' + data.fileUrl + '", "画像確認")',    // 2: 画像確認
    data.processedAt,                                    // 3: 処理日時
    data.date,                                           // 4: 日付
    data.storeName,                                      // 5: 利用店舗名
    data.ocr.invoiceNumber || '',                        // 6: 登録番号
    acc.totalAmount,                                     // 7: 総合計（税込総額）
    acc.subtotal10 || '',                                // 8: 対象額(10%)（税抜）
    acc.tax10 || '',                                     // 9: 消費税(10%)
    acc.subtotal8 || '',                                 // 10: 対象額(8%)（税抜）
    acc.tax8 || '',                                      // 11: 消費税(8%)
    acc.rawNonTaxable || '',                             // 12: 不課税（入湯税など）
    data.accountTitle,                                   // 13: 勘定科目
    data.creditAccount,                                  // 14: 貸方科目
    data.fileName,                                       // 15: ファイル名
    data.debugInfo                                       // 16: Debug
  ];

  sheet.appendRow(row);

  // 条件付き書式（エラー行を赤背景）
  const lastRow = sheet.getLastRow();
  if (data.status === 'ERROR' || data.status === 'CHECK') {
    sheet.getRange(lastRow, 1, 1, row.length).setBackground('#FFE6E6');
  } else if (data.status === 'COMPOUND') {
    sheet.getRange(lastRow, 1, 1, row.length).setBackground('#FFFACD');
  } else if (data.status === 'HAND') {
    sheet.getRange(lastRow, 1, 1, row.length).setBackground('#FFF3E0'); // オレンジ系の薄い背景
  }
}

// ============================================================
// 外部モジュール呼び出し（スタブ）
// 実装は各Service_*.gs, Logic_*.gsで行う
// ============================================================

/**
 * OCR抽出（Service_OCR.gsで実装）
 * @param {GoogleAppsScript.Drive.File} file
 * @return {OCRResult}
 */
function extractOCR_(file) {
  // Service_OCR.gs の extractOCRFromFile() を呼び出す
  if (typeof extractOCRFromFile === 'function') {
    return extractOCRFromFile(file);
  }

  // スタブ（未実装時）
  console.warn('extractOCRFromFile is not implemented');
  return {
    date: '',
    storeName: 'UNKNOWN',
    totalAmount: null,
    invoiceNumber: null,
    items: [],
    rawText: ''
  };
}

/**
 * 会計計算（Logic_Accounting.gsで実装）
 * @param {OCRResult} ocrResult
 * @return {AccountingData}
 */
function calculateAccounting_(ocrResult) {
  // Logic_Accounting.gs の calculateAccountingData() を呼び出す
  if (typeof calculateAccountingData === 'function') {
    return calculateAccountingData(ocrResult);
  }

  // スタブ（未実装時）
  console.warn('calculateAccountingData is not implemented');
  return {
    subtotal10: 0,
    tax10: 0,
    subtotal8: 0,
    tax8: 0,
    rawNonTaxable: 0,
    totalAmount: ocrResult.totalAmount || 0,
    isCompound: false,
    hasInvoice: !!ocrResult.invoiceNumber,
    taxType: 'standard'
  };
}

/**
 * 勘定科目推定（Logic_Accounting.gsで実装）
 * @param {string} storeName
 * @param {Array<OCRLineItem>} items
 * @param {string} [geminiSuggestion] - Geminiが推定した勘定科目
 * @param {OCRResult} [ocrResult] - OCRデータ（軽油税判定用）
 * @return {string|null}
 */
function inferAccountTitle_(storeName, items, geminiSuggestion, ocrResult) {
  // Logic_Accounting.gs の inferAccountTitleFromStore() を呼び出す
  if (typeof inferAccountTitleFromStore === 'function') {
    return inferAccountTitleFromStore(storeName, items, geminiSuggestion, ocrResult);
  }

  // スタブ（未実装時）- Config.gs のマッピングを使用
  const rules = loadMappingRules_();
  const storeStr = String(storeName || '').toLowerCase();

  for (const rule of rules) {
    if (storeStr.includes(rule.keyword.toLowerCase())) {
      return rule.accountTitle;
    }
  }

  return null;
}

// ============================================================
// ユーザー操作用関数
// ============================================================

/**
 * 処理済みマークをリセット
 */
function resetProcessedMarks() {
  const folderConfigs = loadFolderConfigs_();
  let count = 0;

  for (const config of folderConfigs) {
    try {
      const folder = DriveApp.getFolderById(config.folderId);
      const files = folder.getFiles();

      while (files.hasNext()) {
        const file = files.next();
        const name = file.getName();

        if (isProcessedFile_(name)) {
          // 新形式と旧形式の両方のプレフィックスを除去
          const newName = name
            .replace(/^\[(OK|CHK|CMP|ERR|\?)\]/, '')
            .replace(/^[🟢🔴🟡🟠]/, '');
          file.setName(newName);
          count++;
        }
      }
    } catch (e) {
      console.error('フォルダ処理エラー: ' + e.message);
    }
  }

  SpreadsheetApp.getUi().alert('リセット完了: ' + count + '件のファイル名を更新しました。');
}

/**
 * CHK/ERRマークのみリセット（OK/CMPは維持）
 */
function resetCheckErrorMarks() {
  const folderConfigs = loadFolderConfigs_();
  let count = 0;

  for (const config of folderConfigs) {
    try {
      const folder = DriveApp.getFolderById(config.folderId);
      const files = folder.getFiles();

      while (files.hasNext()) {
        const file = files.next();
        const name = file.getName();

        // [CHK] または [ERR] プレフィックスのみ除去
        if (/^\[(CHK|ERR)\]/.test(name)) {
          const newName = name.replace(/^\[(CHK|ERR)\]/, '');
          file.setName(newName);
          count++;
        }
        // 旧形式の🔴プレフィックスも対象
        if (/^🔴/.test(name)) {
          const newName = name.replace(/^🔴/, '');
          file.setName(newName);
          count++;
        }
      }
    } catch (e) {
      console.error('フォルダ処理エラー: ' + e.message);
    }
  }

  SpreadsheetApp.getUi().alert(
    'CHK/ERRリセット完了: ' + count + '件\n' +
    'OK/CMPのファイルはそのまま維持されています。\n' +
    '再読み込みを実行してください。'
  );
}

/**
 * クレカ突合（Service_Reconcile.gsで実装）
 */
function reconcileWithStatements() {
  if (typeof runReconciliation === 'function') {
    runReconciliation();
  } else {
    SpreadsheetApp.getUi().alert('Service_Reconcile.gs が読み込まれていません。');
  }
}

/**
 * 突合情報リセット（Service_Reconcile.gsで実装）
 */
function resetReconcileInfo() {
  if (typeof resetReconciliation === 'function') {
    resetReconciliation();
  } else {
    SpreadsheetApp.getUi().alert('Service_Reconcile.gs が読み込まれていません。');
  }
}

/**
 * 弥生CSV出力（Output_Yayoi.gsで実装）
 */
function exportToYayoiCSV() {
  if (typeof generateYayoiCSV === 'function') {
    generateYayoiCSV();
  } else {
    SpreadsheetApp.getUi().alert('Output_Yayoi.gs が読み込まれていません。');
  }
}

/**
 * サイドバー表示
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('レシート確認');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * 選択中の行のデータを取得（サイドバーから呼び出し）
 * @return {Object|null}
 */
function getSelectedRowData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getName() !== CONFIG.SHEET_NAME.MAIN) return null;

  const row = sheet.getActiveCell().getRow();
  if (row <= 1) return null; // ヘッダー行

  const lastCol = sheet.getLastColumn();
  const readCols = Math.max(lastCol, 20);
  const values = sheet.getRange(row, 1, 1, readCols).getValues()[0];
  const formulas = sheet.getRange(row, 2, 1, 1).getFormulas()[0];

  // 列B のHYPERLINK数式からファイルURLを抽出
  // 形式: =HYPERLINK("https://drive.google.com/file/d/xxx/view...", "画像確認")
  let fileUrl = '';
  let fileId = '';
  const formula = formulas[0] || '';
  const urlMatch = formula.match(/HYPERLINK\("([^"]+)"/);
  if (urlMatch) {
    fileUrl = urlMatch[1];
    const idMatch = fileUrl.match(/\/d\/([^\/]+)/);
    if (idMatch) {
      fileId = idMatch[1];
    }
  }

  // MIMEタイプをファイル名から推定（Drive API呼び出しを回避）
  const mimeType = guessMimeType_(String(values[14] || ''));

  // ステータスから絵文字を除去（例: '🖊️HAND' → 'HAND'）
  const rawStatus = String(values[0] || '');
  const status = rawStatus.replace(/^[🟢🔴🟡🟠🖊️]+/, ''); // 絵文字を削除

  // 画像表示用のURLを生成（通常モードと同じ軽量な方式）
  // Base64は重いので使わず、Drive の thumbnail/preview URL を使用
  let imageUrl = '';
  if (fileId) {
    if (mimeType && mimeType.startsWith('image/')) {
      // 画像の場合：thumbnail URL（軽量・高速）
      imageUrl = `https://drive.google.com/thumbnail?id=${fileId}&sz=w800`;
    } else {
      // PDF等の場合：preview URL（iframe用）
      imageUrl = `https://drive.google.com/file/d/${fileId}/preview`;
    }
    Logger.log(`[getSelectedRowData] imageUrl生成: ${mimeType} -> ${imageUrl}`);
  }

  return {
    row: row,
    status: status,        // 1: Status（絵文字なし）
    fileUrl: fileUrl,                         // 2: 画像確認URL
    fileId: fileId,
    imageUrl: imageUrl,                       // 画像表示用URL
    mimeType: mimeType,
    date: formatDateValue_(values[3]),         // 4: 日付
    storeName: String(values[4] || ''),      // 5: 利用店舗名
    invoiceNumber: String(values[5] || ''),  // 6: 登録番号
    totalAmount: values[6] || 0,             // 7: 総合計
    subtotal10: values[7] || 0,              // 8: 対象額(10%)
    tax10: values[8] || 0,                   // 9: 消費税(10%)
    subtotal8: values[9] || 0,               // 10: 対象額(8%)
    tax8: values[10] || 0,                   // 11: 消費税(8%)
    nonTaxable: values[11] || 0,             // 12: 不課税
    accountTitle: String(values[12] || ''),  // 13: 勘定科目
    creditAccount: String(values[13] || ''), // 14: 貸方科目
    fileName: String(values[14] || ''),      // 15: ファイル名
    debug: String(values[15] || ''),         // 16: Debug
    verificationStatus: String(values[16] || ''),  // 17: 検証ステータス
    verificationScore: values[17] || '',           // 18: 検証スコア
    verificationResult: String(values[18] || ''),  // 19: 検証結果
    verificationJSON: String(values[19] || '')     // 20: 修正案JSON
  };
}

/**
 * 日付値をYYYY-MM-DD文字列に変換
 * @param {*} val
 * @return {string}
 */
function formatDateValue_(val) {
  if (!val) return '';
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return y + '-' + m + '-' + d;
  }
  return String(val);
}

/**
 * ファイル名の拡張子からMIMEタイプを推定
 * @param {string} fileName
 * @return {string}
 */
function guessMimeType_(fileName) {
  const ext = String(fileName || '').split('.').pop().toLowerCase();
  const map = {
    'pdf': 'application/pdf',
    'jpg': 'image/jpeg',
    'jpeg': 'image/jpeg',
    'png': 'image/png',
    'gif': 'image/gif',
    'heic': 'image/heic',
    'webp': 'image/webp',
    'bmp': 'image/bmp',
    'tiff': 'image/tiff',
    'tif': 'image/tiff'
  };
  return map[ext] || 'application/pdf';
}

/**
 * 手書き領収証の金額を更新
 * @param {number} row - 行番号
 * @param {number} totalAmount - 総額
 * @param {number} taxable10 - 10%税抜
 * @param {number} tax10 - 10%税額
 * @param {number} taxable8 - 8%税抜
 * @param {number} tax8 - 8%税額
 * @param {number} nonTaxable - 不課税
 * @return {Object} 処理結果
 */
function updateHandReceipt(row, totalAmount, taxable10, tax10, taxable8, tax8, nonTaxable) {
  try {
    console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    console.log('[updateHandReceipt] 関数が呼ばれました！');
    console.log('[updateHandReceipt] row:', row);
    console.log('[updateHandReceipt] totalAmount:', totalAmount);
    console.log('[updateHandReceipt] taxable10:', taxable10);
    console.log('[updateHandReceipt] tax10:', tax10);
    console.log('[updateHandReceipt] taxable8:', taxable8);
    console.log('[updateHandReceipt] tax8:', tax8);
    console.log('[updateHandReceipt] nonTaxable:', nonTaxable);
    console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

    Logger.log('[updateHandReceipt] 手書き領収証を更新: Row ' + row);

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME.MAIN);

    if (!sheet) {
      return {
        success: false,
        message: 'シート「' + CONFIG.SHEET_NAME.MAIN + '」が見つかりません'
      };
    }

    // 現在のステータスを確認（A列 = 1列目）
    const rawStatus = String(sheet.getRange(row, 1).getValue() || '');
    const currentStatus = rawStatus.replace(/^[🟢🔴🟡🟠🖊️]+/, ''); // 絵文字除去

    if (currentStatus !== 'HAND') {
      return {
        success: false,
        message: 'この行は手書き領収証ではありません（現在のステータス: ' + currentStatus + '）'
      };
    }

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // 金額を更新
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    sheet.getRange(row, 7).setValue(Number(totalAmount) || 0);      // G列: 総額
    sheet.getRange(row, 8).setValue(Number(taxable10) || 0);        // H列: 10%税抜
    sheet.getRange(row, 9).setValue(Number(tax10) || 0);            // I列: 10%税額
    sheet.getRange(row, 10).setValue(Number(taxable8) || 0);        // J列: 8%税抜
    sheet.getRange(row, 11).setValue(Number(tax8) || 0);            // K列: 8%税額
    sheet.getRange(row, 12).setValue(Number(nonTaxable) || 0);      // L列: 不課税

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // ステータスを確定に変更
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    sheet.getRange(row, 1).setValue('🟢OK');  // A列: ステータス

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // 検証結果を更新（手動確認済み）
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    sheet.getRange(row, 17).setValue('🖊️ 手動確認済み');  // Q列: 検証ステータス
    sheet.getRange(row, 18).setValue(100);                 // R列: 検証スコア（100%）
    sheet.getRange(row, 19).setValue('');                  // S列: 検証結果（クリア）
    sheet.getRange(row, 20).setValue('');                  // T列: 修正案JSON（クリア）

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // ファイル名のプレフィックスを変更
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    // B列のHYPERLINK式からファイルIDを抽出
    const formulas = sheet.getRange(row, 2, 1, 1).getFormulas()[0];
    const formula = formulas[0] || '';
    let fileId = '';
    const urlMatch = formula.match(/HYPERLINK\("([^"]+)"/);
    if (urlMatch) {
      const fileUrl = urlMatch[1];
      const idMatch = fileUrl.match(/\/d\/([^\/]+)/);
      if (idMatch) {
        fileId = idMatch[1];
      }
    }

    const fileName = sheet.getRange(row, 15).getValue();  // O列: ファイル名（プレフィックスなし）

    Logger.log('[updateHandReceipt] ファイルID: "' + fileId + '"');
    Logger.log('[updateHandReceipt] 現在のファイル名（スプシ）: "' + fileName + '"');

    if (fileId) {
      try {
        // ファイルIDから直接ファイルを取得
        const file = DriveApp.getFileById(fileId);
        const actualFileName = file.getName();

        Logger.log('[updateHandReceipt] 実際のファイル名（Drive）: "' + actualFileName + '"');

        // [HAND] または 🖊️ を [OK] または 🟢 に変更
        let newFileName = actualFileName;
        newFileName = newFileName.replace(/^\[HAND\]/, '[OK]');
        newFileName = newFileName.replace(/^🖊️/, '🟢');

        Logger.log('[updateHandReceipt] 新しいファイル名: "' + newFileName + '"');

        if (actualFileName !== newFileName) {
          // ファイル名を変更
          file.setName(newFileName);
          Logger.log('[updateHandReceipt] ✅ ファイル名を変更: "' + actualFileName + '" → "' + newFileName + '"');

          // スプレッドシートのファイル名も更新（プレフィックスなしの基本ファイル名）
          const baseFileName = newFileName.replace(/^\[(HAND|OK|CHK|CMP|ERR)\]/, '').replace(/^[🖊️🟢🔴🟡🟠]/, '');
          sheet.getRange(row, 15).setValue(baseFileName);
          Logger.log('[updateHandReceipt] ✅ スプレッドシートのファイル名も更新: "' + baseFileName + '"');
        } else {
          Logger.log('[updateHandReceipt] ℹ️ ファイル名の変更は不要（既に正しいプレフィックス）');
        }

      } catch (error) {
        Logger.log('[updateHandReceipt] ❌ ファイル処理エラー: ' + error.toString());
        Logger.log('[updateHandReceipt] スタック: ' + error.stack);
        // エラーでも処理を継続（ファイル名変更は必須ではない）
      }
    } else {
      Logger.log('[updateHandReceipt] ⚠️ ファイルIDが取得できませんでした');
      Logger.log('[updateHandReceipt] HYPERLINK式: "' + formula + '"');
    }

    Logger.log('手書き領収証の更新完了: Row ' + row);

    return {
      success: true,
      message: '手書き領収証を確定しました'
    };

  } catch (error) {
    Logger.log('エラー: ' + error.toString());

    return {
      success: false,
      message: 'エラーが発生しました: ' + error.toString()
    };
  }
}

/**
 * Configシートを作成・初期化
 */
function setupConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'Config';
  let sheet = ss.getSheetByName(sheetName);

  if (sheet) {
    SpreadsheetApp.getUi().alert('Configシートは既に存在します。');
    return;
  }

  sheet = ss.insertSheet(sheetName);

  // ヘッダー
  sheet.appendRow(['項目キー', '設定値', '説明']);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold');

  // 初期データ（APIキーはScriptPropertiesで管理するため含めない）
  const data = [
    ['FOLDER_ID_RECEIPTS', '', 'レシート画像が格納されているフォルダID（必須）'],
    ['FOLDER_ID_PROCESSED', '', '処理が終わったファイルを移動するフォルダID（空欄なら移動しない）'],
    ['CC_STATEMENT_SPREADSHEET_ID', '', 'クレカ明細スプレッドシートID']
  ];

  data.forEach(function(row) {
    sheet.appendRow(row);
  });

  // 列幅調整
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 350);
  sheet.setColumnWidth(3, 400);

  SpreadsheetApp.getUi().alert('Configシートを作成しました。各項目の設定値を入力してください。');
}
