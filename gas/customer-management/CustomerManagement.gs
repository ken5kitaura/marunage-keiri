/**
 * ==========================================
 * まるなげ経理 - B2C顧客管理GAS
 * ==========================================
 *
 * 顧客コードプレフィックス: MK
 * フォルダ: まるなげ経理 顧客データ
 *
 * 【列構成】
 * A: line_user_id
 * B: customer_name
 * C: folder_id（領収書フォルダID）
 * D: created_at
 * E: notified
 * F: sent_at
 * G: customer_code
 * H: status
 * I: email
 * J: (未使用)
 * K: trial_count
 * L: (未使用)
 * M: (未使用)
 * N: stripe_customer_id
 * O: plan
 * P: monthly_price
 * Q: billing_start_date
 * R: spreadsheet_url
 * S: passbook_folder_id
 * T: cc_statement_folder_id
 * U: invitation_sent_at
 */

// ========== 定数 ==========
const CONFIG = {
  PARENT_FOLDER_ID: '1_D9JjIRLhZ6PWgVgOj4MjrWSNCEuID3N',
  SERVICE_ACCOUNT_EMAIL: '845322634063-compute@developer.gserviceaccount.com',
  SHEET_NAME: '顧客管理',
  FLOW_SHEET_NAME: 'フロー手順',
  CODE_PREFIX: 'MK',
  SERVICE_NAME: 'まるなげ経理',
  LINE_URL: 'https://lin.ee/KbUqcWG',
  STATUS: {
    UNUSED: '未使用',
    NOTIFIED: '案内済',
    CONTRACTED: '契約済',
    CANCELLED: '解約'
  },
  COL: {
    LINE_USER_ID: 1,        // A
    CUSTOMER_NAME: 2,       // B
    FOLDER_ID: 3,           // C
    CREATED_AT: 4,          // D
    NOTIFIED: 5,            // E
    SENT_AT: 6,             // F
    CUSTOMER_CODE: 7,       // G
    STATUS: 8,              // H
    EMAIL: 9,               // I
    UNUSED_J: 10,           // J
    TRIAL_COUNT: 11,        // K
    UNUSED_L: 12,           // L
    UNUSED_M: 13,           // M
    STRIPE_CUSTOMER_ID: 14, // N
    PLAN: 15,               // O
    MONTHLY_PRICE: 16,      // P
    BILLING_START_DATE: 17, // Q
    SPREADSHEET_URL: 18,    // R
    PASSBOOK_FOLDER_ID: 19, // S
    CC_STATEMENT_FOLDER_ID: 20, // T
    INVITATION_SENT_AT: 21  // U
  }
};

// ========== メニュー ==========

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔧 顧客管理')
    .addItem('📁 新規フォルダ50件作成（一括）', 'createNewCustomerFoldersWithAll')
    .addSeparator()
    .addItem('📁 フォルダのみ作成', 'createNewCustomerFolders')
    .addItem('📝 スプシに登録', 'registerNewFoldersToSheet')
    .addItem('🔑 権限付与（全フォルダ）', 'grantAccessToAllFoldersAndParents')
    .addSeparator()
    .addItem('🔄 既存顧客にサブフォルダ追加', 'addSubfoldersToExistingCustomers')
    .addItem('📊 次の顧客コードを確認', 'showNextCustomerCode')
    .addSeparator()
    .addItem('⚙️ 初期セットアップ', 'initialSetup')
    .addToUi();
}

// ========== 初期セットアップ ==========

/**
 * 初期セットアップ（ヘッダー・ドロップダウン・フロー手順シート）
 */
function initialSetup() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '初期セットアップ',
    '以下を実行します：\n\n' +
    '1. 顧客管理シートのヘッダー設定\n' +
    '2. フロー手順シートの作成\n' +
    '3. ステータス列のドロップダウン設定\n\n' +
    '続行しますか？',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('キャンセルしました');
    return;
  }

  setupSheet();
  createFlowSheet();

  ui.alert('✅ 初期セットアップが完了しました');
}

/**
 * 顧客管理シートのヘッダーを設定
 */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = ss.getActiveSheet();
    sheet.setName(CONFIG.SHEET_NAME);
  }

  const headers = [
    'line_user_id', 'customer_name', 'folder_id', 'created_at',
    'notified', 'sent_at', 'customer_code', 'status', 'email',
    '', 'trial_count', '', '', 'stripe_customer_id', 'plan',
    'monthly_price', 'billing_start_date', 'spreadsheet_url',
    'passbook_folder_id', 'cc_statement_folder_id', 'invitation_sent_at'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);

  // 列幅調整
  sheet.setColumnWidth(CONFIG.COL.LINE_USER_ID, 120);
  sheet.setColumnWidth(CONFIG.COL.CUSTOMER_NAME, 150);
  sheet.setColumnWidth(CONFIG.COL.FOLDER_ID, 280);
  sheet.setColumnWidth(CONFIG.COL.CUSTOMER_CODE, 100);
  sheet.setColumnWidth(CONFIG.COL.STATUS, 100);
  sheet.setColumnWidth(CONFIG.COL.EMAIL, 200);
  sheet.setColumnWidth(CONFIG.COL.SPREADSHEET_URL, 300);

  // ステータスドロップダウン設定
  setupStatusDropdown();
}

/**
 * ステータス列にドロップダウンを設定
 */
function setupStatusDropdown() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return;

  const statusList = [
    CONFIG.STATUS.UNUSED,
    CONFIG.STATUS.NOTIFIED,
    CONFIG.STATUS.CONTRACTED,
    CONFIG.STATUS.CANCELLED
  ];

  const range = sheet.getRange(2, CONFIG.COL.STATUS, lastRow - 1, 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(statusList, true)
    .setAllowInvalid(false)
    .build();

  range.setDataValidation(rule);
}

/**
 * フロー手順シートを作成
 */
function createFlowSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let flowSheet = ss.getSheetByName(CONFIG.FLOW_SHEET_NAME);
  if (flowSheet) {
    ss.deleteSheet(flowSheet);
  }

  flowSheet = ss.insertSheet(CONFIG.FLOW_SHEET_NAME);

  const flowContent = [
    ['まるなげ経理 - 顧客管理フロー手順書'],
    [''],
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    [''],
    ['■ ステータスの意味'],
    [''],
    ['ステータス', '意味'],
    ['未使用', '準備済み、顧客未割当'],
    ['案内済', 'Stripe決済完了、LINE連携待ち'],
    ['契約済', 'LINE連携完了、レシート処理対象'],
    ['解約', 'サービス終了'],
    [''],
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    [''],
    ['■ 新規顧客の登録フロー'],
    [''],
    ['【Stripe決済からの自動登録】'],
    ['  1. 顧客がStripe Payment Linkで決済'],
    ['  2. stripe-webhookが自動で「未使用」行にコードを割り当て'],
    ['  3. ステータスが「案内済」に自動変更'],
    ['  4. 顧客がLINEでコードを入力して連携完了'],
    ['  5. ステータスが「契約済」に自動変更'],
    [''],
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    [''],
    ['■ 顧客がLINE連携する流れ（顧客の作業）'],
    [''],
    ['1. 顧客が招待メールを受信（またはLP経由）'],
    ['2. LINEで公式アカウントを友達追加'],
    ['3. トーク画面で顧客コード（例: MK005）を入力'],
    ['4. 「設定が完了しました」と表示される'],
    [''],
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    [''],
    ['■ 領収書・通帳の送信（顧客の作業）'],
    [''],
    ['1. 顧客がLINEで写真を送信'],
    ['2. 自動で分類されて保存'],
    ['   ・レシート → 領収書フォルダ'],
    ['   ・通帳 → 通帳フォルダ'],
    ['3. 顧客に確認メッセージが返信される'],
    [''],
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    [''],
    ['■ 顧客枠が足りなくなったら'],
    [''],
    ['メニュー → 🔧 顧客管理 → 📁 新規フォルダ50件作成（一括）'],
    [''],
    ['次の番号から50件が自動作成されます。'],
    [''],
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    [''],
    ['■ トラブルシューティング'],
    [''],
    ['問題', '対処'],
    ['顧客コードが認識されない', 'シートにコードがあるか確認'],
    ['メールが届かない', 'メールアドレスを確認、迷惑メール確認'],
    ['ステータスがドロップダウンにならない', 'メニュー → ⚙️ 初期セットアップ を実行'],
    [''],
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━'],
  ];

  flowSheet.getRange(1, 1, flowContent.length, 2).setValues(
    flowContent.map(row => [row[0] || '', row[1] || ''])
  );

  flowSheet.getRange(1, 1).setFontSize(14).setFontWeight('bold');
  flowSheet.setColumnWidth(1, 400);
  flowSheet.setColumnWidth(2, 350);
  flowSheet.getRange(7, 1, 1, 2).setFontWeight('bold');
  flowSheet.getRange(62, 1, 1, 2).setFontWeight('bold');
}

// ========== 一括フォルダ作成 ==========

/**
 * 新規フォルダ50件作成 → スプシ登録 → 権限付与 を一括実行
 */
function createNewCustomerFoldersWithAll() {
  const ui = SpreadsheetApp.getUi();

  const nextCode = getNextCustomerCodeNumber();
  const startNum = nextCode;
  const endNum = nextCode + 49;

  const response = ui.alert(
    '新規フォルダ作成',
    `${CONFIG.CODE_PREFIX}${String(startNum).padStart(3, '0')} 〜 ${CONFIG.CODE_PREFIX}${String(endNum).padStart(3, '0')} の50件を作成します。\n\n` +
    `各フォルダには以下が含まれます：\n` +
    `・領収書サブフォルダ\n` +
    `・通帳サブフォルダ\n` +
    `・クレカ明細サブフォルダ\n` +
    `・レシート読み込み用スプシ（ClientConfig付き）\n\n` +
    `続行しますか？`,
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('キャンセルしました');
    return;
  }

  ui.alert('処理を開始します。完了までお待ちください...');

  try {
    const createdFolders = createFoldersInRange(startNum, endNum);
    const registeredCount = registerFoldersToSheetFromList(createdFolders);
    grantAccessToNewFolders(createdFolders);
    setupStatusDropdown();

    ui.alert(
      '✅ 完了',
      `📁 フォルダ作成: ${createdFolders.length}件\n` +
      `📝 シート登録: ${registeredCount}件\n` +
      `🔑 権限付与: 完了\n\n` +
      `${CONFIG.CODE_PREFIX}${String(startNum).padStart(3, '0')} 〜 ${CONFIG.CODE_PREFIX}${String(endNum).padStart(3, '0')}`,
      ui.ButtonSet.OK
    );

  } catch (e) {
    console.log('エラー: ' + e.message);
    ui.alert('❌ エラー', 'エラーが発生しました: ' + e.message, ui.ButtonSet.OK);
  }
}

// ========== フォルダ作成 ==========

/**
 * 指定範囲のフォルダを作成
 */
function createFoldersInRange(startNum, endNum) {
  const parentFolder = DriveApp.getFolderById(CONFIG.PARENT_FOLDER_ID);
  const createdFolders = [];

  for (let i = startNum; i <= endNum; i++) {
    const code = CONFIG.CODE_PREFIX + String(i).padStart(3, '0');

    const mainFolder = parentFolder.createFolder(code);
    const receiptFolder = mainFolder.createFolder('領収書');
    const passbookFolder = mainFolder.createFolder('通帳');
    const ccStatementFolder = mainFolder.createFolder('クレカ明細');

    const spreadsheet = createCustomerSpreadsheet(mainFolder, code, {
      receiptFolderId: receiptFolder.getId(),
      passbookFolderId: passbookFolder.getId(),
      ccStatementFolderId: ccStatementFolder.getId()
    });

    createdFolders.push({
      code: code,
      mainFolderId: mainFolder.getId(),
      receiptFolderId: receiptFolder.getId(),
      passbookFolderId: passbookFolder.getId(),
      ccStatementFolderId: ccStatementFolder.getId(),
      spreadsheetId: spreadsheet ? spreadsheet.getId() : '',
      spreadsheetUrl: spreadsheet ? spreadsheet.getUrl() : ''
    });

    console.log(`Created: ${code} - サブフォルダ & スプシ作成完了`);
  }

  return createdFolders;
}

/**
 * 顧客用スプレッドシートを作成（ClientConfig付き）
 */
function createCustomerSpreadsheet(parentFolder, customerCode, folderIds) {
  try {
    const spreadsheet = SpreadsheetApp.create(`${customerCode}_レシート読込`);
    const file = DriveApp.getFileById(spreadsheet.getId());

    parentFolder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);

    createClientConfigSheet(spreadsheet, folderIds);

    return spreadsheet;
  } catch (e) {
    console.log(`スプシ作成エラー (${customerCode}): ${e.message}`);
    return null;
  }
}

/**
 * ClientConfigシートを作成
 */
function createClientConfigSheet(spreadsheet, folderIds) {
  let sheet = spreadsheet.getSheets()[0];
  sheet.setName('ClientConfig');

  const configData = [
    ['PASSBOOK_FOLDER_ID', folderIds.passbookFolderId],
    ['RECEIPT_FOLDER_ID', folderIds.receiptFolderId],
    ['CC_STATEMENT_FOLDER_ID', folderIds.ccStatementFolderId]
  ];

  sheet.getRange(1, 1, configData.length, 2).setValues(configData);
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 350);
  sheet.getRange(1, 1, configData.length, 1).setFontWeight('bold');
}

/**
 * 新規フォルダ作成（メニューから個別実行用）
 */
function createNewCustomerFolders() {
  const ui = SpreadsheetApp.getUi();
  const nextCode = getNextCustomerCodeNumber();

  const response = ui.prompt(
    'フォルダ作成',
    `次の顧客コード: ${CONFIG.CODE_PREFIX}${String(nextCode).padStart(3, '0')}\n\n作成件数を入力してください（1〜100）:`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const count = parseInt(response.getResponseText());
  if (isNaN(count) || count < 1 || count > 100) {
    ui.alert('1〜100の数字を入力してください');
    return;
  }

  const createdFolders = createFoldersInRange(nextCode, nextCode + count - 1);
  registerFoldersToSheetFromList(createdFolders);
  grantAccessToNewFolders(createdFolders);
  setupStatusDropdown();

  ui.alert(`✅ ${createdFolders.length}件のフォルダを作成しました`);
}

// ========== 既存顧客にサブフォルダ追加 ==========

/**
 * 既存の顧客フォルダに通帳・クレカ明細サブフォルダを追加
 * スプシがなければ作成、あればClientConfigを追加/更新
 */
function addSubfoldersToExistingCustomers() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  const response = ui.alert(
    'サブフォルダ追加',
    `既存の全顧客フォルダに以下を追加/更新します：\n\n` +
    `・「通帳」サブフォルダ（なければ作成）\n` +
    `・「クレカ明細」サブフォルダ（なければ作成）\n` +
    `・顧客用スプシ（なければ作成）\n` +
    `・ClientConfigシート（なければ作成、あれば更新）\n\n` +
    `続行しますか？`,
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  let processedCount = 0;
  let skippedCount = 0;
  let errorCount = 0;
  let spreadsheetCreatedCount = 0;

  for (let i = 1; i < data.length; i++) {
    const receiptFolderId = data[i][CONFIG.COL.FOLDER_ID - 1];
    const customerCode = data[i][CONFIG.COL.CUSTOMER_CODE - 1];

    if (!receiptFolderId || !customerCode) {
      skippedCount++;
      continue;
    }

    try {
      const receiptFolder = DriveApp.getFolderById(receiptFolderId);
      const parents = receiptFolder.getParents();

      if (!parents.hasNext()) {
        skippedCount++;
        continue;
      }

      const mainFolder = parents.next();

      // 通帳フォルダ作成（既存チェック）
      let passbookFolderId = '';
      const existingPassbook = mainFolder.getFoldersByName('通帳');
      if (existingPassbook.hasNext()) {
        passbookFolderId = existingPassbook.next().getId();
      } else {
        const passbookFolder = mainFolder.createFolder('通帳');
        passbookFolderId = passbookFolder.getId();
        grantEditorAccess(passbookFolderId);
      }

      // クレカ明細フォルダ作成（既存チェック）
      let ccStatementFolderId = '';
      const existingCC = mainFolder.getFoldersByName('クレカ明細');
      if (existingCC.hasNext()) {
        ccStatementFolderId = existingCC.next().getId();
      } else {
        const ccFolder = mainFolder.createFolder('クレカ明細');
        ccStatementFolderId = ccFolder.getId();
        grantEditorAccess(ccStatementFolderId);
      }

      // スプシを探す or 作成
      let spreadsheetUrl = data[i][CONFIG.COL.SPREADSHEET_URL - 1];
      let spreadsheet = null;

      if (spreadsheetUrl) {
        try {
          const spreadsheetId = extractSpreadsheetId(spreadsheetUrl);
          if (spreadsheetId) {
            spreadsheet = SpreadsheetApp.openById(spreadsheetId);
          }
        } catch (ssError) {
          spreadsheet = null;
        }
      }

      if (!spreadsheet) {
        const existingFiles = mainFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
        if (existingFiles.hasNext()) {
          const existingFile = existingFiles.next();
          spreadsheet = SpreadsheetApp.openById(existingFile.getId());
          spreadsheetUrl = existingFile.getUrl();
        } else {
          spreadsheet = createCustomerSpreadsheet(mainFolder, customerCode, {
            receiptFolderId: receiptFolderId,
            passbookFolderId: passbookFolderId,
            ccStatementFolderId: ccStatementFolderId
          });
          if (spreadsheet) {
            spreadsheetUrl = spreadsheet.getUrl();
            spreadsheetCreatedCount++;
          }
        }
      }

      // ClientConfigシートを更新
      if (spreadsheet) {
        updateClientConfigSheet(spreadsheet, {
          receiptFolderId: receiptFolderId,
          passbookFolderId: passbookFolderId,
          ccStatementFolderId: ccStatementFolderId
        });
      }

      // シートに書き込み
      sheet.getRange(i + 1, CONFIG.COL.PASSBOOK_FOLDER_ID).setValue(passbookFolderId);
      sheet.getRange(i + 1, CONFIG.COL.CC_STATEMENT_FOLDER_ID).setValue(ccStatementFolderId);
      if (spreadsheetUrl) {
        sheet.getRange(i + 1, CONFIG.COL.SPREADSHEET_URL).setValue(spreadsheetUrl);
      }

      processedCount++;

    } catch (e) {
      console.log(`${customerCode}: エラー - ${e.message}`);
      errorCount++;
    }
  }

  ui.alert(
    '✅ 完了',
    `処理完了:\n\n` +
    `✅ 処理済み: ${processedCount}件\n` +
    `📄 スプシ新規作成: ${spreadsheetCreatedCount}件\n` +
    `⏭️ スキップ: ${skippedCount}件\n` +
    `❌ エラー: ${errorCount}件`,
    ui.ButtonSet.OK
  );
}

/**
 * スプレッドシートURLからIDを抽出
 */
function extractSpreadsheetId(url) {
  if (!url) return null;
  const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : null;
}

/**
 * ClientConfigシートを更新（既存があれば更新、なければ作成）
 */
function updateClientConfigSheet(spreadsheet, folderIds) {
  let sheet = spreadsheet.getSheetByName('ClientConfig');

  if (!sheet) {
    sheet = spreadsheet.insertSheet('ClientConfig');
  }

  sheet.clear();

  const configData = [
    ['PASSBOOK_FOLDER_ID', folderIds.passbookFolderId],
    ['RECEIPT_FOLDER_ID', folderIds.receiptFolderId],
    ['CC_STATEMENT_FOLDER_ID', folderIds.ccStatementFolderId]
  ];

  sheet.getRange(1, 1, configData.length, 2).setValues(configData);
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 350);
  sheet.getRange(1, 1, configData.length, 1).setFontWeight('bold');
}

// ========== スプシ登録 ==========

/**
 * 作成済みフォルダリストから顧客管理シートに登録
 */
function registerFoldersToSheetFromList(folderList) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  let count = 0;

  for (const folder of folderList) {
    // 列構成に合わせた行データ（A〜U列 = 21列）
    const rowData = new Array(21).fill('');
    rowData[CONFIG.COL.FOLDER_ID - 1] = folder.receiptFolderId;
    rowData[CONFIG.COL.CUSTOMER_CODE - 1] = folder.code;
    rowData[CONFIG.COL.STATUS - 1] = CONFIG.STATUS.UNUSED;
    rowData[CONFIG.COL.SPREADSHEET_URL - 1] = folder.spreadsheetUrl;
    rowData[CONFIG.COL.PASSBOOK_FOLDER_ID - 1] = folder.passbookFolderId;
    rowData[CONFIG.COL.CC_STATEMENT_FOLDER_ID - 1] = folder.ccStatementFolderId;

    sheet.appendRow(rowData);
    count++;
  }

  return count;
}

/**
 * 既存フォルダを顧客管理シートに登録（メニューから個別実行用）
 */
function registerNewFoldersToSheet() {
  const ui = SpreadsheetApp.getUi();
  const parentFolder = DriveApp.getFolderById(CONFIG.PARENT_FOLDER_ID);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);

  const existingData = sheet.getDataRange().getValues();
  const existingCodes = new Set();
  for (let i = 1; i < existingData.length; i++) {
    if (existingData[i][CONFIG.COL.CUSTOMER_CODE - 1]) {
      existingCodes.add(existingData[i][CONFIG.COL.CUSTOMER_CODE - 1]);
    }
  }

  const folders = parentFolder.getFolders();
  let addedCount = 0;

  while (folders.hasNext()) {
    const mainFolder = folders.next();
    const code = mainFolder.getName();

    if (code.startsWith(CONFIG.CODE_PREFIX) && !existingCodes.has(code)) {
      const receiptFolders = mainFolder.getFoldersByName('領収書');
      if (receiptFolders.hasNext()) {
        const receiptFolder = receiptFolders.next();
        let passbookFolderId = '', ccStatementFolderId = '', spreadsheetUrl = '';

        const passbookFolders = mainFolder.getFoldersByName('通帳');
        if (passbookFolders.hasNext()) passbookFolderId = passbookFolders.next().getId();

        const ccFolders = mainFolder.getFoldersByName('クレカ明細');
        if (ccFolders.hasNext()) ccStatementFolderId = ccFolders.next().getId();

        const files = mainFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
        if (files.hasNext()) spreadsheetUrl = files.next().getUrl();

        const rowData = new Array(21).fill('');
        rowData[CONFIG.COL.FOLDER_ID - 1] = receiptFolder.getId();
        rowData[CONFIG.COL.CUSTOMER_CODE - 1] = code;
        rowData[CONFIG.COL.STATUS - 1] = CONFIG.STATUS.UNUSED;
        rowData[CONFIG.COL.SPREADSHEET_URL - 1] = spreadsheetUrl;
        rowData[CONFIG.COL.PASSBOOK_FOLDER_ID - 1] = passbookFolderId;
        rowData[CONFIG.COL.CC_STATEMENT_FOLDER_ID - 1] = ccStatementFolderId;

        sheet.appendRow(rowData);
        addedCount++;
        console.log('Registered: ' + code);
      }
    }
  }

  if (addedCount > 0) setupStatusDropdown();
  ui.alert(`✅ ${addedCount}件を登録しました`);
}

// ========== 権限付与 ==========

/**
 * 新規作成フォルダに権限付与
 */
function grantAccessToNewFolders(folderList) {
  for (const folder of folderList) {
    try {
      grantEditorAccess(folder.receiptFolderId);
      grantEditorAccess(folder.passbookFolderId);
      grantEditorAccess(folder.ccStatementFolderId);
      grantEditorAccess(folder.mainFolderId);
    } catch (e) {
      console.log(`権限付与エラー (${folder.code}): ${e.message}`);
    }
  }
}

/**
 * 全フォルダに権限付与（メニューから個別実行用）
 */
function grantAccessToAllFoldersAndParents() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  let successCount = 0, skipCount = 0, errorCount = 0;

  for (let i = 1; i < data.length; i++) {
    const receiptFolderId = data[i][CONFIG.COL.FOLDER_ID - 1];
    const passbookFolderId = data[i][CONFIG.COL.PASSBOOK_FOLDER_ID - 1];
    const ccStatementFolderId = data[i][CONFIG.COL.CC_STATEMENT_FOLDER_ID - 1];

    if (!receiptFolderId) { skipCount++; continue; }

    try {
      grantEditorAccess(receiptFolderId);
      if (passbookFolderId) grantEditorAccess(passbookFolderId);
      if (ccStatementFolderId) grantEditorAccess(ccStatementFolderId);

      const receiptFolder = DriveApp.getFolderById(receiptFolderId);
      const parents = receiptFolder.getParents();
      if (parents.hasNext()) grantEditorAccess(parents.next().getId());

      successCount++;
    } catch (e) {
      errorCount++;
    }
  }

  SpreadsheetApp.getUi().alert(`完了!\n✅ 権限付与: ${successCount}件\n⏭️ スキップ: ${skipCount}件\n❌ エラー: ${errorCount}件`);
}

/**
 * 単一フォルダに権限付与
 */
function grantEditorAccess(folderId) {
  if (!folderId) return;

  const folder = DriveApp.getFolderById(folderId);
  const editors = folder.getEditors();
  for (const editor of editors) {
    if (editor.getEmail() === CONFIG.SERVICE_ACCOUNT_EMAIL) return;
  }

  folder.addEditor(CONFIG.SERVICE_ACCOUNT_EMAIL);
}

// ========== LINE通知（既存機能） ==========

function sendNotificationOnCheck(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = e.range;

  if (range.getColumn() !== CONFIG.COL.NOTIFIED) return;
  if (range.getValue() !== true) return;

  var row = range.getRow();
  if (row <= 1) return;

  var alreadySent = sheet.getRange(row, CONFIG.COL.SENT_AT).getValue();
  if (alreadySent) {
    Logger.log('既に送信済み: row ' + row);
    return;
  }

  var lineUserId = sheet.getRange(row, CONFIG.COL.LINE_USER_ID).getValue();
  var customerName = sheet.getRange(row, CONFIG.COL.CUSTOMER_NAME).getValue();
  var folderId = sheet.getRange(row, CONFIG.COL.FOLDER_ID).getValue();

  if (!folderId) {
    SpreadsheetApp.getUi().alert('folder_id が空です。先にフォルダIDを入力してください。');
    range.setValue(false);
    return;
  }

  var now = new Date();
  sheet.getRange(row, CONFIG.COL.SENT_AT).setValue(now);

  var success = sendLineNotification(lineUserId, customerName);

  if (!success) {
    SpreadsheetApp.getUi().alert('LINE送信に失敗しました。');
    range.setValue(false);
    sheet.getRange(row, CONFIG.COL.SENT_AT).setValue('');
  }
}

function sendLineNotification(userId, customerName) {
  var token = PropertiesService.getScriptProperties().getProperty('LINE_ACCESS_TOKEN');

  if (!token) {
    Logger.log('LINE_ACCESS_TOKEN が設定されていません');
    return false;
  }

  var message = '✅ 設定が完了しました！\n\n' +
                '領収書の写真をこのトークに送ってください。\n' +
                '自動で保存されます。\n\n' +
                '📸 複数枚まとめて送ってもOKです！';

  var payload = {
    'to': userId,
    'messages': [{ 'type': 'text', 'text': message }]
  };

  var options = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': { 'Authorization': 'Bearer ' + token },
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  try {
    var response = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', options);
    var code = response.getResponseCode();
    Logger.log('LINE API Response: ' + code + ' ' + response.getContentText());
    return code === 200;
  } catch (error) {
    Logger.log('LINE送信エラー: ' + error);
    return false;
  }
}

// ========== ユーティリティ ==========

/**
 * 次の顧客コード番号を取得
 */
function getNextCustomerCodeNumber() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return 1;

  const data = sheet.getDataRange().getValues();
  let maxNum = 0;

  for (let i = 1; i < data.length; i++) {
    const code = data[i][CONFIG.COL.CUSTOMER_CODE - 1];
    if (code && code.startsWith(CONFIG.CODE_PREFIX)) {
      const num = parseInt(code.replace(CONFIG.CODE_PREFIX, ''));
      if (!isNaN(num) && num > maxNum) maxNum = num;
    }
  }

  return maxNum + 1;
}

/**
 * 次の顧客コードを表示
 */
function showNextCustomerCode() {
  const nextNum = getNextCustomerCodeNumber();
  const nextCode = CONFIG.CODE_PREFIX + String(nextNum).padStart(3, '0');
  SpreadsheetApp.getUi().alert(`次の顧客コード: ${nextCode}`);
}

function testPermission() {
  var response = UrlFetchApp.fetch('https://www.google.com');
  Logger.log('OK: ' + response.getResponseCode());
}
