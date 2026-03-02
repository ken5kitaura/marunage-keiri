/**
 * ==========================================
 * 絆パートナーズ税理士法人 - 顧客管理GAS
 * ==========================================
 * 
 * 顧客コードプレフィックス: KZ
 * フォルダ: 絆パートナーズ税理士法人
 * 
 * 【列構成】
 * A: line_user_id
 * B: customer_name
 * C: folder_id（領収書フォルダID）
 * D: registered_at
 * E: customer_code
 * F: status
 * G: email
 * H: phone
 * I: memo
 * J: passbook_folder_id
 * K: cc_statement_folder_id
 * L: spreadsheet_url
 * M: invitation_sent_at
 */

// ========== 定数 ==========
const CONFIG = {
  PARENT_FOLDER_ID: '1Np68bNyn8QDxDSdSJay4CF-sWx7ckJtK',
  SERVICE_ACCOUNT_EMAIL: '845322634063-compute@developer.gserviceaccount.com',
  SHEET_NAME: '顧客管理',
  FLOW_SHEET_NAME: 'フロー手順',
  CODE_PREFIX: 'KZ',
  SERVICE_NAME: '絆パートナーズ税理士法人 記帳代行サービス',
  LINE_URL: 'https://line.me/R/ti/p/@821hkrnz',
  STATUS: {
    UNUSED: '未使用',
    NOTIFIED: '案内済',
    CONTRACTED: '契約済'
  },
  COL: {
    LINE_USER_ID: 1,
    CUSTOMER_NAME: 2,
    FOLDER_ID: 3,
    REGISTERED_AT: 4,
    CUSTOMER_CODE: 5,
    STATUS: 6,
    EMAIL: 7,
    PHONE: 8,
    MEMO: 9,
    PASSBOOK_FOLDER_ID: 10,
    CC_STATEMENT_FOLDER_ID: 11,
    SPREADSHEET_URL: 12,
    INVITATION_SENT_AT: 13
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
    .addItem('📧 選択行に招待メール送信', 'sendInvitationToSelectedRows')
    .addItem('📊 次の顧客コードを確認', 'showNextCustomerCode')
    .addSeparator()
    .addItem('⚙️ 初期セットアップ', 'initialSetup')
    .addToUi();
}

// ========== 初期セットアップ ==========

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

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.getActiveSheet();
    sheet.setName(CONFIG.SHEET_NAME);
  }
  
  const headers = [
    'line_user_id', 'customer_name', 'folder_id', 'registered_at',
    'customer_code', 'status', 'email', 'phone', 'memo',
    'passbook_folder_id', 'cc_statement_folder_id', 'spreadsheet_url', 'invitation_sent_at'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  
  sheet.setColumnWidth(CONFIG.COL.LINE_USER_ID, 120);
  sheet.setColumnWidth(CONFIG.COL.CUSTOMER_NAME, 150);
  sheet.setColumnWidth(CONFIG.COL.FOLDER_ID, 280);
  sheet.setColumnWidth(CONFIG.COL.CUSTOMER_CODE, 100);
  sheet.setColumnWidth(CONFIG.COL.STATUS, 100);
  sheet.setColumnWidth(CONFIG.COL.EMAIL, 200);
  sheet.setColumnWidth(CONFIG.COL.SPREADSHEET_URL, 300);
}

function createFlowSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let flowSheet = ss.getSheetByName(CONFIG.FLOW_SHEET_NAME);
  if (flowSheet) {
    ss.deleteSheet(flowSheet);
  }
  
  flowSheet = ss.insertSheet(CONFIG.FLOW_SHEET_NAME);
  
  const flowContent = [
    ['記帳代行サービス - 顧客管理フロー手順書'],
    [''],
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    [''],
    ['■ ステータスの意味'],
    [''],
    ['ステータス', '意味'],
    ['未使用', '準備済み、顧客未割当'],
    ['案内済', '招待メール送信完了、LINE連携待ち'],
    ['契約済', 'LINE連携完了、レシート処理対象'],
    [''],
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    [''],
    ['■ 新規顧客を登録する手順'],
    [''],
    ['【STEP 1】顧客管理シートを開く'],
    [''],
    ['【STEP 2】「未使用」ステータスの行を探す'],
    [''],
    ['【STEP 3】以下を入力'],
    ['  ・B列: 顧客名'],
    ['  ・G列: メールアドレス'],
    [''],
    ['【STEP 4】入力した行を選択し、メニューから「📧 選択行に招待メール送信」を実行'],
    ['  ・成功 → ステータスが「案内済」に自動変更'],
    ['  ・失敗 → エラーダイアログが表示'],
    [''],
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    [''],
    ['■ 顧客がLINE連携する流れ（顧客の作業）'],
    [''],
    ['1. 顧客が招待メールを受信'],
    ['2. LINEで公式アカウントを友達追加'],
    ['3. トーク画面で顧客コード（例: KZ001）を入力'],
    ['4. 「設定が完了しました」と表示される'],
    ['5. ステータスが「案内済」→「契約済」に自動変更'],
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
  flowSheet.getRange(64, 1, 1, 2).setFontWeight('bold');
}

function setupStatusDropdown() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) return;
  
  const statusList = [
    CONFIG.STATUS.UNUSED,
    CONFIG.STATUS.NOTIFIED,
    CONFIG.STATUS.CONTRACTED
  ];
  
  const range = sheet.getRange(2, CONFIG.COL.STATUS, lastRow - 1, 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(statusList, true)
    .setAllowInvalid(false)
    .build();
  
  range.setDataValidation(rule);
}

// ========== 一括処理 ==========

function createNewCustomerFoldersWithAll() {
  const ui = SpreadsheetApp.getUi();
  const nextCode = getNextCustomerCodeNumber();
  const startNum = nextCode;
  const endNum = nextCode + 49;
  
  const response = ui.alert(
    '新規フォルダ作成',
    `${CONFIG.CODE_PREFIX}${String(startNum).padStart(3, '0')} 〜 ${CONFIG.CODE_PREFIX}${String(endNum).padStart(3, '0')} の50件を作成します。\n\n続行しますか？`,
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    ui.alert('キャンセルしました');
    return;
  }
  
  ui.alert('処理を開始します。完了までお待ちください...');
  
  try {
    const createdFolders = createFoldersInRange(startNum, endNum);
    registerFoldersToSheetFromList(createdFolders);
    grantAccessToNewFolders(createdFolders);
    setupStatusDropdown();
    
    ui.alert(
      '✅ 完了',
      `📁 フォルダ作成: ${createdFolders.length}件\n📝 シート登録: 完了\n🔑 権限付与: 完了\n\n${CONFIG.CODE_PREFIX}${String(startNum).padStart(3, '0')} 〜 ${CONFIG.CODE_PREFIX}${String(endNum).padStart(3, '0')}`,
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('❌ エラー', 'エラーが発生しました: ' + e.message, ui.ButtonSet.OK);
  }
}

// ========== フォルダ作成 ==========

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
      spreadsheetUrl: spreadsheet ? spreadsheet.getUrl() : ''
    });
  }
  
  return createdFolders;
}

function createCustomerSpreadsheet(parentFolder, customerCode, folderIds) {
  try {
    const spreadsheet = SpreadsheetApp.create(`${customerCode}_レシート読込`);
    const file = DriveApp.getFileById(spreadsheet.getId());
    parentFolder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);
    createClientConfigSheet(spreadsheet, folderIds);
    return spreadsheet;
  } catch (e) {
    return null;
  }
}

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

// ========== スプシ登録 ==========

function registerFoldersToSheetFromList(folderList) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  
  for (const folder of folderList) {
    sheet.appendRow([
      '', '', folder.receiptFolderId, '', folder.code, CONFIG.STATUS.UNUSED,
      '', '', '', folder.passbookFolderId, folder.ccStatementFolderId,
      folder.spreadsheetUrl, ''
    ]);
  }
  
  return folderList.length;
}

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
        
        sheet.appendRow([
          '', '', receiptFolder.getId(), '', code, CONFIG.STATUS.UNUSED,
          '', '', '', passbookFolderId, ccStatementFolderId, spreadsheetUrl, ''
        ]);
        addedCount++;
      }
    }
  }
  
  if (addedCount > 0) setupStatusDropdown();
  ui.alert(`✅ ${addedCount}件を登録しました`);
}

// ========== 権限付与 ==========

function grantAccessToNewFolders(folderList) {
  for (const folder of folderList) {
    try {
      grantEditorAccess(folder.receiptFolderId);
      grantEditorAccess(folder.passbookFolderId);
      grantEditorAccess(folder.ccStatementFolderId);
      grantEditorAccess(folder.mainFolderId);
    } catch (e) {}
  }
}

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
    } catch (e) { errorCount++; }
  }
  
  SpreadsheetApp.getUi().alert(`完了!\n✅ 権限付与: ${successCount}件\n⏭️ スキップ: ${skipCount}件\n❌ エラー: ${errorCount}件`);
}

function grantEditorAccess(folderId) {
  if (!folderId) return;
  const folder = DriveApp.getFolderById(folderId);
  const editors = folder.getEditors();
  for (const editor of editors) {
    if (editor.getEmail() === CONFIG.SERVICE_ACCOUNT_EMAIL) return;
  }
  folder.addEditor(CONFIG.SERVICE_ACCOUNT_EMAIL);
}

// ========== 招待メール送信 ==========

/**
 * 選択した行の顧客に招待メールを送信する。
 * メニューから実行。送信後ステータスを「案内済」に変更。
 */
function sendInvitationToSelectedRows() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const selection = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveRange();

  if (sheet.getName() !== SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName()) {
    ui.alert('⚠️ 顧客管理シートで実行してください。');
    return;
  }

  const startRow = selection.getRow();
  const numRows = selection.getNumRows();

  if (startRow <= 1) {
    ui.alert('⚠️ ヘッダー行は選択できません。データ行を選択してください。');
    return;
  }

  // 送信対象を収集して確認
  const targets = [];
  for (let row = startRow; row < startRow + numRows; row++) {
    const customerName = sheet.getRange(row, CONFIG.COL.CUSTOMER_NAME).getValue();
    const email = sheet.getRange(row, CONFIG.COL.EMAIL).getValue();
    const customerCode = sheet.getRange(row, CONFIG.COL.CUSTOMER_CODE).getValue();
    const status = sheet.getRange(row, CONFIG.COL.STATUS).getValue();

    if (!customerName || !email) continue;
    if (status === CONFIG.STATUS.NOTIFIED || status === CONFIG.STATUS.CONTRACTED) continue;

    targets.push({ row, customerName, email, customerCode });
  }

  if (targets.length === 0) {
    ui.alert('⚠️ 送信対象がありません。\n\n顧客名とメールアドレスが入力済みで、ステータスが「未使用」の行を選択してください。');
    return;
  }

  // 確認ダイアログ
  const names = targets.map(t => `  ${t.customerCode}: ${t.customerName} (${t.email})`).join('\n');
  const confirm = ui.alert(
    '📧 招待メール送信確認',
    `以下の${targets.length}件に招待メールを送信します。\n\n${names}\n\n送信しますか？`,
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  // 送信実行
  let sentCount = 0, errorCount = 0;
  const errors = [];

  for (const target of targets) {
    try {
      sendInvitationEmail_(target.email, target.customerName, target.customerCode);
      sheet.getRange(target.row, CONFIG.COL.STATUS).setValue(CONFIG.STATUS.NOTIFIED);
      sheet.getRange(target.row, CONFIG.COL.INVITATION_SENT_AT).setValue(new Date());
      sentCount++;
    } catch (e) {
      errors.push(`${target.customerCode}: ${e.message}`);
      errorCount++;
    }
  }

  let message = `✅ 招待メール送信完了: ${sentCount}件`;
  if (errorCount > 0) message += `\n❌ エラー: ${errorCount}件\n\n${errors.join('\n')}`;
  ui.alert(message);
}

/**
 * 招待メールを送信（内部関数）
 */
function sendInvitationEmail_(email, customerName, customerCode) {
  const subject = `【${CONFIG.SERVICE_NAME}】ご利用開始のご案内`;

  const qrCodeUrl = 'https://qr-official.line.me/gs/M_821hkrnz_GW.png';

  const plainBody = customerName + ' 様\n\n' +
    'いつもお世話になっております。\n絆パートナーズ税理士法人です。\n\n' +
    '記帳代行サービスのご利用準備が整いましたので、\n下記の手順でサービスをご利用ください。\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━\n■ ご利用開始の手順\n━━━━━━━━━━━━━━━━━━━━━━\n\n' +
    '【STEP 1】LINEで友だち追加\n下記のリンクからLINE公式アカウントを友だち追加してください。\n' + CONFIG.LINE_URL + '\n\n' +
    '【STEP 2】顧客コードを入力\nLINEのトーク画面で、以下の顧客コードを入力してください。\n\n' +
    '━━━━━━━━━━━━━━━━━\nあなたの顧客コード: ' + customerCode + '\n━━━━━━━━━━━━━━━━━\n\n' +
    '【STEP 3】領収書を送信\n設定完了後、領収書や通帳の写真をLINEで送るだけ！\nあとは担当者がすべて対応いたします。\n\n' +
    'ご不明な点がございましたら、お気軽にお問い合わせください。\n\n今後ともよろしくお願いいたします。\n\n絆パートナーズ税理士法人';

  const htmlBody = '<div style="font-family: Helvetica Neue, Arial, Hiragino Kaku Gothic Pro, Meiryo, sans-serif; max-width: 600px; margin: 0 auto; color: #333;">' +
    '<p>' + customerName + ' 様</p>' +
    '<p>いつもお世話になっております。<br>絆パートナーズ税理士法人です。</p>' +
    '<p>記帳代行サービスのご利用準備が整いましたので、<br>下記の手順でサービスをご利用ください。</p>' +
    '<div style="background: #f8f8f5; border-radius: 12px; padding: 24px; margin: 24px 0;">' +
    '<h3 style="color: #2B5F3F; margin-top: 0;">■ ご利用開始の手順</h3>' +
    '<div style="margin-bottom: 20px;">' +
    '<p style="font-weight: bold; color: #2B5F3F;">STEP 1｜LINEで友だち追加</p>' +
    '<p>下のボタンまたはQRコードから、LINE公式アカウントを友だち追加してください。</p>' +
    '<div style="text-align: center; margin: 16px 0;">' +
    '<a href="' + CONFIG.LINE_URL + '" style="display: inline-block; background: #06C755; color: #fff; font-weight: bold; padding: 12px 32px; border-radius: 50px; text-decoration: none; font-size: 16px;">LINEで友だち追加</a>' +
    '</div>' +
    '<div style="text-align: center; margin: 16px 0;">' +
    '<img src="' + qrCodeUrl + '" alt="QRコード" width="160" height="160" style="border: 1px solid #ddd; border-radius: 8px;">' +
    '<p style="font-size: 12px; color: #888; margin-top: 4px;">スマホのカメラで読み取ってください</p>' +
    '</div></div>' +
    '<div style="margin-bottom: 20px;">' +
    '<p style="font-weight: bold; color: #2B5F3F;">STEP 2｜顧客コードを入力</p>' +
    '<p>LINEのトーク画面で、以下の顧客コードを入力してください。</p>' +
    '<div style="background: #fff; border: 2px solid #2B5F3F; border-radius: 8px; padding: 16px; text-align: center; margin: 12px 0;">' +
    '<p style="font-size: 12px; color: #888; margin: 0 0 4px 0;">あなたの顧客コード</p>' +
    '<p style="font-size: 28px; font-weight: bold; color: #2B5F3F; margin: 0; letter-spacing: 0.1em;">' + customerCode + '</p>' +
    '</div></div>' +
    '<div>' +
    '<p style="font-weight: bold; color: #2B5F3F;">STEP 3｜領収書を送信</p>' +
    '<p>設定完了後、領収書や通帳の写真をLINEで送るだけ！<br>あとは担当者がすべて対応いたします。</p>' +
    '</div></div>' +
    '<p>ご不明な点がございましたら、お気軽にお問い合わせください。</p>' +
    '<p>今後ともよろしくお願いいたします。</p>' +
    '<hr style="border: none; border-top: 1px solid #ddd; margin: 24px 0;">' +
    '<p style="font-size: 13px; color: #888;">絆パートナーズ税理士法人</p></div>';

  MailApp.sendEmail({
    to: email,
    subject: subject,
    body: plainBody,
    htmlBody: htmlBody
  });
}

// ========== ユーティリティ ==========

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

function showNextCustomerCode() {
  const nextNum = getNextCustomerCodeNumber();
  const nextCode = CONFIG.CODE_PREFIX + String(nextNum).padStart(3, '0');
  SpreadsheetApp.getUi().alert(`次の顧客コード: ${nextCode}`);
}
