/**
 * ==========================================
 * まるなげ経理 - 中央管理GAS
 * ==========================================
 *
 * 全パートナーの顧客を横断して一括処理する司令塔。
 * このGASは独立したスプレッドシートにバインドする。
 *
 * 【機能】
 * - 全顧客のレシート処理（1時間毎トリガー）
 * - 全顧客のAI検証（1日1回トリガー）
 * - ラッパーGASの一括配布
 * - パートナー管理（顧客管理シートの登録）
 *
 * 【パートナー設定】
 * 「パートナー設定」シートに各パートナーの情報を記載。
 * 列: パートナー名, シートID, シート名, 顧客コード列, ステータス列,
 *     スプシURL列, アクティブステータス（カンマ区切り）
 */

// ============================================================
// メニュー
// ============================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('🚀 一括処理')
    .addItem('📦 全顧客レシート処理', 'batchProcessAllReceipts')
    .addItem('🤖 全顧客AI検証', 'batchRunAutoVerification')
    .addSeparator()
    .addItem('📜 ラッパーGAS一括配布', 'deployWrapperToAllClients')
    .addToUi();

  ui.createMenu('⚙️ 設定')
    .addItem('📋 パートナー設定シートを初期化', 'initPartnerConfigSheet')
    .addItem('🔑 ReceiptEngine ScriptIDを設定', 'promptReceiptEngineScriptId')
    .addItem('📄 ラッパーテンプレートを登録（手動貼り付け）', 'promptWrapperTemplate')
    .addItem('📄 最新テンプレートを一括登録', 'registerLatestWrapperTemplate')
    .addSeparator()
    .addItem('🔍 パートナー設定を確認', 'showPartnerSummary')
    .addItem('📊 全顧客一覧を表示', 'showAllClientsSummary')
    .addToUi();
}

// ============================================================
// パートナー設定の管理
// ============================================================

const PARTNER_SHEET_NAME = 'パートナー設定';

/**
 * パートナー設定シートを初期化する。
 * 既に存在する場合はスキップ。
 */
function initPartnerConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PARTNER_SHEET_NAME);

  if (sheet) {
    SpreadsheetApp.getUi().alert('パートナー設定シートは既に存在します。');
    return;
  }

  sheet = ss.insertSheet(PARTNER_SHEET_NAME);

  // ヘッダー
  const headers = [
    'パートナー名',           // A
    'シートID',               // B
    'シート名',               // C
    '顧客コード列（0始まり）', // D
    'ステータス列（0始まり）', // E
    'スプシURL列（0始まり）',  // F
    'アクティブステータス',    // G（カンマ区切り）
    '有効',                   // H（TRUE/FALSE）
    'scriptId列（0始まり）'   // I（顧客管理シート内でGASのscriptIdを記録する列）
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);

  // サンプルデータ: メイン（MK）
  sheet.appendRow([
    'まるなげ経理（MK）',
    '',  // ← ここにメイン顧客管理シートのIDを入力
    '顧客管理',
    6,   // G列（0始まり）
    7,   // H列
    17,  // R列
    '利用中,トライアル',
    true
  ]);

  // サンプルデータ: 絆パートナーズ（KZ）
  sheet.appendRow([
    '絆パートナーズ（KZ）',
    '',  // ← ここに絆顧客管理シートのIDを入力
    '顧客管理',
    4,   // E列（0始まり）
    5,   // F列
    11,  // L列
    '契約済',
    true
  ]);

  // 列幅調整
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 350);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(7, 200);

  SpreadsheetApp.getUi().alert(
    '✅ パートナー設定シートを作成しました。\n\n' +
    'B列に各パートナーの顧客管理スプレッドシートIDを入力してください。'
  );
}

/**
 * パートナー設定を読み込む。
 * @return {Array<Object>}
 */
function loadPartnerConfigs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PARTNER_SHEET_NAME);

  if (!sheet) {
    throw new Error('パートナー設定シートがありません。「設定」メニューから初期化してください。');
  }

  const data = sheet.getDataRange().getValues();
  const partners = [];

  for (let i = 1; i < data.length; i++) {
    const enabled = data[i][7];
    if (enabled !== true && enabled !== 'TRUE') continue;

    const sheetId = String(data[i][1] || '').trim();
    if (!sheetId) continue;

    partners.push({
      name: String(data[i][0] || '').trim(),
      spreadsheetId: sheetId,
      sheetName: String(data[i][2] || '顧客管理').trim(),
      codeCol: parseInt(data[i][3]) || 0,
      statusCol: parseInt(data[i][4]) || 0,
      urlCol: parseInt(data[i][5]) || 0,
      activeStatuses: String(data[i][6] || '').split(',').map(s => s.trim()).filter(s => s),
      scriptIdCol: (data[i][8] !== undefined && data[i][8] !== '') ? parseInt(data[i][8]) : -1
    });
  }

  return partners;
}

// ============================================================
// 全パートナー横断の顧客一覧取得
// ============================================================

/**
 * 全パートナーからアクティブな顧客のスプシID一覧を取得する。
 * scriptIdCol が設定されている場合は既存のscriptIdも読み取る。
 * @return {Array<{partner: string, code: string, spreadsheetId: string, scriptId: string, _sheet: Sheet, _row: number, _scriptIdCol: number}>}
 */
function getAllActiveClients_() {
  const partners = loadPartnerConfigs_();
  const allClients = [];

  for (const partner of partners) {
    try {
      const ss = SpreadsheetApp.openById(partner.spreadsheetId);
      const sheet = ss.getSheetByName(partner.sheetName);

      if (!sheet) {
        console.warn('[' + partner.name + '] シート「' + partner.sheetName + '」が見つかりません');
        continue;
      }

      const data = sheet.getDataRange().getValues();

      for (let i = 1; i < data.length; i++) {
        const code = String(data[i][partner.codeCol] || '').trim();
        const status = String(data[i][partner.statusCol] || '').trim();
        const spreadsheetUrl = String(data[i][partner.urlCol] || '').trim();

        if (!code || !spreadsheetUrl) continue;
        if (!partner.activeStatuses.includes(status)) continue;

        // URLからスプシIDを抽出
        const match = spreadsheetUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
        if (!match) continue;

        // scriptIdを読み取る（列が設定されている場合）
        var scriptId = '';
        if (partner.scriptIdCol >= 0) {
          scriptId = String(data[i][partner.scriptIdCol] || '').trim();
        }

        allClients.push({
          partner: partner.name,
          code: code,
          spreadsheetId: match[1],
          scriptId: scriptId,
          _sheet: sheet,               // 書き戻し用
          _row: i + 1,                 // 1始まり行番号（書き戻し用）
          _scriptIdCol: partner.scriptIdCol  // 書き戻し用
        });
      }

      console.log('[' + partner.name + '] ' + allClients.filter(c => c.partner === partner.name).length + '件のアクティブ顧客');

    } catch (e) {
      console.error('[' + partner.name + '] 読み込みエラー: ' + e.message);
    }
  }

  return allClients;
}

// ============================================================
// 一括レシート処理
// ============================================================

/**
 * 全アクティブ顧客のレシートを一括処理する。
 * トリガー（1時間毎）から呼び出される。
 */
function batchProcessAllReceipts() {
  const startTime = Date.now();
  const MAX_TIME_MS = 5 * 60 * 1000;

  const clients = getAllActiveClients_();
  console.log('=== 一括レシート処理開始: ' + clients.length + '件 ===');

  let processedCount = 0;
  let errorCount = 0;

  for (const client of clients) {
    if (Date.now() - startTime > MAX_TIME_MS) {
      console.log('タイムアウト: ' + processedCount + '/' + clients.length + '件処理済み。残りは次回。');
      break;
    }

    try {
      console.log('[' + client.code + '] レシート処理開始...');
      ReceiptEngine.processReceiptsById(client.spreadsheetId);
      processedCount++;
      console.log('[' + client.code + '] 完了');
    } catch (e) {
      errorCount++;
      console.error('[' + client.code + '] エラー: ' + e.message);
    }
  }

  console.log('=== 一括処理完了: 成功=' + processedCount + ', エラー=' + errorCount + ' ===');

  // 実行ログをシートに記録
  logExecution_('レシート処理', clients.length, processedCount, errorCount);
}

// ============================================================
// 一括AI検証
// ============================================================

/**
 * 全アクティブ顧客のAI検証を一括実行する。
 * トリガー（1日1回）またはメニューから呼び出される。
 * タイムアウト時は継続トリガーで残りを自動再開する。
 */
function batchRunAutoVerification() {
  batchRunAutoVerification_(false);
}

/**
 * batchRunAutoVerification の継続実行用。
 * 継続トリガーから呼ばれ、前回の続きから処理する。
 */
function batchRunAutoVerification_continue() {
  console.log('batchRunAutoVerification_continue: 継続実行を開始');
  batchRunAutoVerification_(true);
}

/**
 * 一括AI検証の実体。
 * @param {boolean} isContinuation - 継続実行かどうか
 */
function batchRunAutoVerification_(isContinuation) {
  const startTime = Date.now();
  const MAX_TIME_MS = 4 * 60 * 1000; // 4分（ライブラリに残り時間を渡して制御するため余裕あり）

  const clients = getAllActiveClients_();

  // 前回の続きからスタートする場合、スキップ位置を取得
  let startIndex = 0;
  if (isContinuation) {
    const saved = PropertiesService.getScriptProperties().getProperty('BATCH_VERIFY_RESUME_INDEX');
    startIndex = saved ? parseInt(saved) : 0;
    if (startIndex >= clients.length) {
      console.log('全顧客処理済み。リセットします。');
      PropertiesService.getScriptProperties().deleteProperty('BATCH_VERIFY_RESUME_INDEX');
      deleteContinuationTrigger_central_('batchRunAutoVerification_continue');
      return;
    }
    console.log('=== 一括AI検証【継続】: ' + startIndex + '番目から再開（全' + clients.length + '件）===');
  } else {
    // 新規実行時はリセット
    PropertiesService.getScriptProperties().deleteProperty('BATCH_VERIFY_RESUME_INDEX');
    deleteContinuationTrigger_central_('batchRunAutoVerification_continue');
    console.log('=== 一括AI検証開始: ' + clients.length + '件 ===');
  }

  let processedCount = 0;
  let errorCount = 0;
  let timedOut = false;

  for (let i = startIndex; i < clients.length; i++) {
    const client = clients[i];

    if (Date.now() - startTime > MAX_TIME_MS) {
      console.log('タイムアウト: ' + processedCount + '/' + (clients.length - startIndex) + '件処理済み。');
      // 次回の再開位置を保存
      PropertiesService.getScriptProperties().setProperty('BATCH_VERIFY_RESUME_INDEX', String(i));
      // 1分後に継続トリガーを設定
      deleteContinuationTrigger_central_('batchRunAutoVerification_continue');
      ScriptApp.newTrigger('batchRunAutoVerification_continue')
        .timeBased()
        .after(1 * 60 * 1000)
        .create();
      console.log('継続トリガーを設定しました（1分後に ' + i + '番目=' + clients[i].code + 'から再開）');
      timedOut = true;
      break;
    }

    try {
      console.log('[' + client.code + '] AI検証開始...');
      const remainingMs = MAX_TIME_MS - (Date.now() - startTime) - 30000; // 30秒の余裕
      const safeTimeMs = Math.max(remainingMs, 60000); // 最低1分は確保
      const result = ReceiptEngine.runAutoVerificationById(client.spreadsheetId, safeTimeMs);
      processedCount++;
      console.log('[' + client.code + '] 完了');

      // ライブラリが途中でタイムアウトした場合、同じ顧客から再開
      if (result && result.timedOut) {
        console.log('[' + client.code + '] ライブラリ内タイムアウト（' + result.processed + '/' + result.total + '件処理済み）→ 同じ顧客から再開');
        PropertiesService.getScriptProperties().setProperty('BATCH_VERIFY_RESUME_INDEX', String(i));
        deleteContinuationTrigger_central_('batchRunAutoVerification_continue');
        ScriptApp.newTrigger('batchRunAutoVerification_continue')
          .timeBased()
          .after(1 * 60 * 1000)
          .create();
        console.log('継続トリガーを設定しました（1分後に ' + i + '番目=' + client.code + 'から再開）');
        timedOut = true;
        break;
      }

      // 中央管理側のタイムアウトチェック（ライブラリは完了したが次の顧客に進む時間がない場合）
      if (Date.now() - startTime > MAX_TIME_MS) {
        console.log('タイムアウト（顧客処理直後）: 次の顧客に進む時間なし');
        PropertiesService.getScriptProperties().setProperty('BATCH_VERIFY_RESUME_INDEX', String(i + 1));
        deleteContinuationTrigger_central_('batchRunAutoVerification_continue');
        ScriptApp.newTrigger('batchRunAutoVerification_continue')
          .timeBased()
          .after(1 * 60 * 1000)
          .create();
        console.log('継続トリガーを設定しました（1分後に ' + (i + 1) + '番目から再開）');
        timedOut = true;
        break;
      }
    } catch (e) {
      errorCount++;
      console.error('[' + client.code + '] エラー: ' + e.message);
    }
  }

  // 全件完了の場合はクリーンアップ
  if (!timedOut) {
    PropertiesService.getScriptProperties().deleteProperty('BATCH_VERIFY_RESUME_INDEX');
    deleteContinuationTrigger_central_('batchRunAutoVerification_continue');
    console.log('全顧客の検証が完了しました。');
  }

  console.log('=== 一括AI検証完了: 成功=' + processedCount + ', エラー=' + errorCount + ' ===');
  logExecution_('AI検証' + (isContinuation ? '（継続）' : ''), clients.length, processedCount, errorCount);
}

/**
 * 中央管理GAS用の継続トリガー削除。
 * @param {string} functionName - 削除対象の関数名
 */
function deleteContinuationTrigger_central_(functionName) {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
      console.log('既存の継続トリガーを削除: ' + functionName);
    }
  }
}

// ============================================================
// ラッパーGAS一括配布
// ============================================================

/**
 * 全パートナーからスプシURLが記録されている全顧客を取得する（ステータス不問）。
 * ラッパー配布専用。「未使用」「案内済」「解約」等もすべて含む。
 * @return {Array<{partner: string, code: string, spreadsheetId: string, scriptId: string, _sheet: Sheet, _row: number, _scriptIdCol: number}>}
 */
function getAllClientsForDeploy_() {
  const partners = loadPartnerConfigs_();
  const allClients = [];

  for (const partner of partners) {
    try {
      const ss = SpreadsheetApp.openById(partner.spreadsheetId);
      const sheet = ss.getSheetByName(partner.sheetName);

      if (!sheet) {
        console.warn('[' + partner.name + '] シート「' + partner.sheetName + '」が見つかりません');
        continue;
      }

      const data = sheet.getDataRange().getValues();

      for (let i = 1; i < data.length; i++) {
        const code = String(data[i][partner.codeCol] || '').trim();
        const spreadsheetUrl = String(data[i][partner.urlCol] || '').trim();

        if (!code || !spreadsheetUrl) continue;

        // URLからスプシIDを抽出
        const match = spreadsheetUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
        if (!match) continue;

        // scriptIdを読み取る（列が設定されている場合）
        var scriptId = '';
        if (partner.scriptIdCol >= 0) {
          scriptId = String(data[i][partner.scriptIdCol] || '').trim();
        }

        allClients.push({
          partner: partner.name,
          code: code,
          spreadsheetId: match[1],
          scriptId: scriptId,
          _sheet: sheet,
          _row: i + 1,
          _scriptIdCol: partner.scriptIdCol
        });
      }

      console.log('[' + partner.name + '] 配布対象: ' + allClients.filter(c => c.partner === partner.name).length + '件');

    } catch (e) {
      console.error('[' + partner.name + '] 読み込みエラー: ' + e.message);
    }
  }

  return allClients;
}

/**
 * ラッパー配布のドライラン。実際の作成・更新は行わず、対象リストと予定アクションをログ出力する。
 * GASエディタから手動実行してログを確認すること。
 */
function dryRunDeployWrapper() {
  const clients = getAllClientsForDeploy_();

  var updateCount = 0;
  var createCount = 0;

  console.log('========================================');
  console.log('ドライラン: ラッパーGAS一括配布');
  console.log('対象顧客数: ' + clients.length + '件');
  console.log('========================================');

  for (var i = 0; i < clients.length; i++) {
    var c = clients[i];
    var action = c.scriptId ? '更新' : '新規作成';
    if (c.scriptId) {
      updateCount++;
    } else {
      createCount++;
    }
    console.log('[' + (i + 1) + '] ' + c.code + ' | ' + action +
      ' | スプシ=' + c.spreadsheetId.substring(0, 12) + '...' +
      ' | scriptId=' + (c.scriptId || '(なし)'));
  }

  console.log('========================================');
  console.log('更新予定: ' + updateCount + '件（scriptId登録済み）');
  console.log('新規作成予定: ' + createCount + '件（scriptId未登録）');
  console.log('合計: ' + clients.length + '件');
  console.log('========================================');
}

/**
 * 全顧客のスプシにGASラッパーコードを一括配布する（メニューから呼ぶ）。
 * ステータスに関係なく、スプシURLが記録されている全顧客が対象。
 */
function deployWrapperToAllClients() {
  const ui = SpreadsheetApp.getUi();
  const clients = getAllClientsForDeploy_();

  if (clients.length === 0) {
    ui.alert('配布対象の顧客が見つかりません。');
    return;
  }

  const confirm = ui.alert(
    'ラッパーGAS一括配布',
    clients.length + '件の顧客スプシ（全ステータス）のGASコードを最新版に更新します。\n続行しますか？',
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  // 新規実行: リセットして開始
  PropertiesService.getScriptProperties().deleteProperty('DEPLOY_WRAPPER_RESUME_INDEX');
  deleteContinuationTrigger_central_('deployWrapperToAllClients_continue');
  deployWrapperToAllClients_(false);
}

/**
 * ラッパー配布の継続実行用。継続トリガーから呼ばれる。
 */
function deployWrapperToAllClients_continue() {
  console.log('deployWrapperToAllClients_continue: 継続実行を開始');
  deployWrapperToAllClients_(true);
}

/**
 * ラッパー配布の実体。タイムアウト時は継続トリガーで残りを自動再開する。
 * @param {boolean} isContinuation - 継続実行かどうか
 */
function deployWrapperToAllClients_(isContinuation) {
  const startTime = Date.now();
  const MAX_TIME_MS = 4 * 60 * 1000; // 4分で切り上げ

  const clients = getAllClientsForDeploy_();
  const wrapperCode = getLatestWrapperCode_();
  const token = ScriptApp.getOAuthToken();

  // 前回の続きから
  let startIndex = 0;
  if (isContinuation) {
    const saved = PropertiesService.getScriptProperties().getProperty('DEPLOY_WRAPPER_RESUME_INDEX');
    startIndex = saved ? parseInt(saved) : 0;
    if (startIndex >= clients.length) {
      console.log('全顧客配布済み。リセットします。');
      PropertiesService.getScriptProperties().deleteProperty('DEPLOY_WRAPPER_RESUME_INDEX');
      deleteContinuationTrigger_central_('deployWrapperToAllClients_continue');
      return;
    }
    console.log('=== ラッパー配布【継続】: ' + startIndex + '番目から再開（全' + clients.length + '件）===');
  } else {
    console.log('=== ラッパー配布開始: ' + clients.length + '件 ===');
  }

  let successCount = 0;
  let errorCount = 0;
  const errors = [];
  let timedOut = false;

  for (let i = startIndex; i < clients.length; i++) {
    // タイムアウトチェック（ループ先頭）
    if (Date.now() - startTime > MAX_TIME_MS) {
      console.log('タイムアウト: ' + successCount + '/' + (clients.length - startIndex) + '件処理済み。');
      PropertiesService.getScriptProperties().setProperty('DEPLOY_WRAPPER_RESUME_INDEX', String(i));
      deleteContinuationTrigger_central_('deployWrapperToAllClients_continue');
      ScriptApp.newTrigger('deployWrapperToAllClients_continue')
        .timeBased()
        .after(1 * 60 * 1000)
        .create();
      console.log('継続トリガーを設定しました（1分後に ' + i + '番目=' + clients[i].code + 'から再開）');
      timedOut = true;
      break;
    }

    const client = clients[i];
    try {
      if (client.scriptId) {
        // ── scriptId登録済み: そのまま更新 ──
        const updateResult = updateScriptContent_(client.scriptId, wrapperCode, token);
        if (updateResult.rateLimited) {
          // レート制限: 5分後に継続トリガーで再開
          console.warn('[' + client.code + '] レート制限（更新）。5分後に再開します。');
          PropertiesService.getScriptProperties().setProperty('DEPLOY_WRAPPER_RESUME_INDEX', String(i));
          deleteContinuationTrigger_central_('deployWrapperToAllClients_continue');
          ScriptApp.newTrigger('deployWrapperToAllClients_continue')
            .timeBased()
            .after(5 * 60 * 1000)
            .create();
          timedOut = true;
          break;
        }
        successCount++;
        console.log('[' + client.code + '] 更新完了: ' + client.scriptId);

      } else {
        // ── scriptId未登録: 新規作成 ──
        console.log('[' + client.code + '] scriptId未登録。新規作成...');
        const createResult = createBoundScript_(client.spreadsheetId, token);

        if (createResult.rateLimited) {
          // レート制限: 5分後に継続トリガーで再開
          console.warn('[' + client.code + '] レート制限（作成）。5分後に再開します。');
          PropertiesService.getScriptProperties().setProperty('DEPLOY_WRAPPER_RESUME_INDEX', String(i));
          deleteContinuationTrigger_central_('deployWrapperToAllClients_continue');
          ScriptApp.newTrigger('deployWrapperToAllClients_continue')
            .timeBased()
            .after(5 * 60 * 1000)
            .create();
          timedOut = true;
          break;
        }

        if (!createResult.scriptId) {
          errors.push(client.code + ': GASプロジェクト作成失敗');
          errorCount++;
          continue;
        }

        const updateResult = updateScriptContent_(createResult.scriptId, wrapperCode, token);
        if (updateResult.rateLimited) {
          // 作成は成功したがコード更新でレート制限 → scriptIdは書き戻してから中断
          if (client._scriptIdCol >= 0 && client._sheet && client._row) {
            client._sheet.getRange(client._row, client._scriptIdCol + 1).setValue(createResult.scriptId);
          }
          console.warn('[' + client.code + '] レート制限（更新）。scriptId記録済み。5分後に再開します。');
          PropertiesService.getScriptProperties().setProperty('DEPLOY_WRAPPER_RESUME_INDEX', String(i));
          deleteContinuationTrigger_central_('deployWrapperToAllClients_continue');
          ScriptApp.newTrigger('deployWrapperToAllClients_continue')
            .timeBased()
            .after(5 * 60 * 1000)
            .create();
          timedOut = true;
          break;
        }

        // scriptIdを顧客管理シートに書き戻す
        if (client._scriptIdCol >= 0 && client._sheet && client._row) {
          client._sheet.getRange(client._row, client._scriptIdCol + 1).setValue(createResult.scriptId);
          console.log('[' + client.code + '] scriptIdを顧客管理シートに記録: ' + createResult.scriptId);
        } else {
          console.warn('[' + client.code + '] scriptId列が未設定のため書き戻しできません。手動で記録してください: ' + createResult.scriptId);
        }

        successCount++;
        console.log('[' + client.code + '] 新規作成＋更新完了: ' + createResult.scriptId);

        // 新規作成後は3秒待つ（レート制限回避）
        Utilities.sleep(3000);
      }
    } catch (e) {
      errorCount++;
      errors.push(client.code + ': ' + e.message);
      console.error('[' + client.code + '] エラー: ' + e.message);
    }
  }

  // 全件完了の場合はクリーンアップ
  if (!timedOut) {
    PropertiesService.getScriptProperties().deleteProperty('DEPLOY_WRAPPER_RESUME_INDEX');
    deleteContinuationTrigger_central_('deployWrapperToAllClients_continue');
    console.log('全顧客の配布が完了しました。');
  }

  console.log('=== ラッパー配布完了: 成功=' + successCount + ', エラー=' + errorCount + ' ===');
  if (errors.length > 0) {
    console.log('エラー詳細:\n' + errors.slice(0, 20).join('\n'));
  }

  logExecution_('ラッパー配布' + (isContinuation ? '（継続）' : ''), clients.length, successCount, errorCount);

  // UI表示（初回メニュー実行時のみ、継続トリガー時はUIなし）
  if (!isContinuation && !timedOut) {
    try {
      let msg = '✅ 完了\n成功: ' + successCount + '件\nエラー: ' + errorCount + '件';
      if (errors.length > 0) {
        msg += '\n\nエラー詳細:\n' + errors.slice(0, 10).join('\n');
      }
      SpreadsheetApp.getUi().alert(msg);
    } catch (e) { /* トリガー実行時はUI非対応 */ }
  } else if (!isContinuation && timedOut) {
    try {
      SpreadsheetApp.getUi().alert(
        '⏱ タイムアウト\n' +
        '成功: ' + successCount + '件（' + startIndex + '〜' + (startIndex + successCount - 1) + '番目）\n' +
        '残り: ' + (clients.length - startIndex - successCount - errorCount) + '件\n\n' +
        '1分後に自動で続きを実行します。'
      );
    } catch (e) { /* トリガー実行時はUI非対応 */ }
  }
}

// ============================================================
// Apps Script API ヘルパー
// ============================================================

/**
 * 最新のラッパーGASコードを取得する。
 * @return {string}
 */
function getLatestWrapperCode_() {
  const code = PropertiesService.getScriptProperties().getProperty('WRAPPER_TEMPLATE');
  if (!code) {
    throw new Error('ラッパーテンプレートが未設定です。「設定」メニューから登録してください。');
  }
  return code;
}

/**
 * API呼び出しの429リトライ付きラッパー。
 * 429 (Rate Limit) の場合は30秒待ってリトライ（最大3回）。
 * @param {string} url
 * @param {Object} options - UrlFetchApp.fetchのオプション
 * @param {number} [maxRetries=3]
 * @return {{response: HTTPResponse, rateLimited: boolean}} rateLimited=trueは3回リトライしても429
 */
function fetchWithRetry_(url, options, maxRetries) {
  maxRetries = maxRetries || 3;
  for (var attempt = 1; attempt <= maxRetries; attempt++) {
    var response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 429) {
      return { response: response, rateLimited: false };
    }
    console.warn('429 Rate Limit (attempt ' + attempt + '/' + maxRetries + ')。30秒待機...');
    if (attempt < maxRetries) {
      Utilities.sleep(30000);
    }
  }
  // 3回リトライしても429
  console.error('429 Rate Limit: ' + maxRetries + '回リトライしても解消されず');
  return { response: response, rateLimited: true };
}

/**
 * スプシにバインドされたGASプロジェクトを新規作成する。
 * 429エラー時はリトライ付き。
 * @param {string} spreadsheetId
 * @param {string} token
 * @return {{scriptId: string|null, rateLimited: boolean}}
 */
function createBoundScript_(spreadsheetId, token) {
  const url = 'https://script.googleapis.com/v1/projects';
  const payload = {
    title: 'レシート処理',
    parentId: spreadsheetId
  };

  const result = fetchWithRetry_(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + token },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (result.rateLimited) {
    return { scriptId: null, rateLimited: true };
  }

  if (result.response.getResponseCode() !== 200) {
    console.error('GASプロジェクト作成失敗: ' + result.response.getContentText());
    return { scriptId: null, rateLimited: false };
  }

  const data = JSON.parse(result.response.getContentText());
  return { scriptId: data.scriptId || null, rateLimited: false };
}

/**
 * GASプロジェクトのソースコードを更新する。
 * 429エラー時はリトライ付き。
 * @param {string} scriptId
 * @param {string} code
 * @param {string} token
 * @return {{rateLimited: boolean}}
 */
function updateScriptContent_(scriptId, code, token) {
  const url = 'https://script.googleapis.com/v1/projects/' + scriptId + '/content';
  const engineId = getReceiptEngineScriptId_();

  const payload = {
    files: [
      {
        name: 'Code',
        type: 'SERVER_JS',
        source: code
      },
      {
        name: 'appsscript',
        type: 'JSON',
        source: JSON.stringify({
          timeZone: 'Asia/Tokyo',
          dependencies: {
            libraries: [{
              userSymbol: 'ReceiptEngine',
              libraryId: engineId,
              version: '0',
              developmentMode: true
            }]
          },
          exceptionLogging: 'STACKDRIVER',
          runtimeVersion: 'V8'
        })
      }
    ]
  };

  const result = fetchWithRetry_(url, {
    method: 'put',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + token },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (result.rateLimited) {
    return { rateLimited: true };
  }

  if (result.response.getResponseCode() !== 200) {
    throw new Error('GASコード更新失敗: ' + result.response.getContentText());
  }
  return { rateLimited: false };
}

/**
 * ReceiptEngineのスクリプトIDを取得する。
 * @return {string}
 */
function getReceiptEngineScriptId_() {
  const id = PropertiesService.getScriptProperties().getProperty('RECEIPT_ENGINE_SCRIPT_ID');
  if (!id) {
    throw new Error('RECEIPT_ENGINE_SCRIPT_ID が未設定です。「設定」メニューから設定してください。');
  }
  return id;
}

// ============================================================
// 設定用UI
// ============================================================

/**
 * ReceiptEngine ScriptIDを設定する。
 */
function promptReceiptEngineScriptId() {
  const ui = SpreadsheetApp.getUi();
  const current = PropertiesService.getScriptProperties().getProperty('RECEIPT_ENGINE_SCRIPT_ID') || '（未設定）';

  const result = ui.prompt(
    'ReceiptEngine ScriptID設定',
    '現在の値: ' + current + '\n\nReceiptEngineのスクリプトIDを入力してください:',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() === ui.Button.OK) {
    const id = result.getResponseText().trim();
    if (id) {
      PropertiesService.getScriptProperties().setProperty('RECEIPT_ENGINE_SCRIPT_ID', id);
      ui.alert('✅ 保存しました: ' + id);
    }
  }
}

/**
 * ラッパーテンプレートを登録する。
 * ※長いコードなのでプロンプトでは入力しきれない場合がある。
 * その場合はGASエディタから setWrapperTemplate() を直接実行。
 */
function promptWrapperTemplate() {
  const ui = SpreadsheetApp.getUi();
  const current = PropertiesService.getScriptProperties().getProperty('WRAPPER_TEMPLATE');
  const currentInfo = current ? '（登録済み: ' + current.length + '文字）' : '（未登録）';

  const result = ui.prompt(
    'ラッパーテンプレート登録',
    currentInfo + '\n\nラッパーGASのソースコード全文を貼り付けてください:\n' +
    '（長すぎる場合はGASエディタから setWrapperTemplate(code) を直接実行してください）',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() === ui.Button.OK) {
    const code = result.getResponseText().trim();
    if (code) {
      PropertiesService.getScriptProperties().setProperty('WRAPPER_TEMPLATE', code);
      ui.alert('✅ ラッパーテンプレートを保存しました（' + code.length + '文字）');
    }
  }
}

/**
 * setWrapperTemplate を直接呼ぶ用のラッパー。
 * GASエディタから実行する場合はこの関数に引数を渡す。
 * @param {string} code
 */
function setWrapperTemplate(code) {
  PropertiesService.getScriptProperties().setProperty('WRAPPER_TEMPLATE', code);
  console.log('ラッパーテンプレートを保存しました（' + code.length + '文字）');
}

/**
 * 最新のラッパーテンプレートをPropertiesServiceに登録する。
 * GASエディタからこの関数を直接実行すること。
 * テンプレートの内容を更新したら、この関数内のコードも更新してから実行する。
 */
function registerLatestWrapperTemplate() {
  var code = [
'/**',
' * クライアント用ラッパースクリプト テンプレート',
' * ',
' * ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
' * このファイルは中央管理GASの「ラッパーGAS一括配布」で',
' * 各顧客スプシに自動配布されます。',
' * ',
' * 手動で設定する場合:',
' * 1. この内容を顧客スプシのGASエディタに貼り付け',
' * 2. ライブラリ「ReceiptEngine」を追加（開発モード）',
' * ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
' */',
'',
'// ============================================================',
'// メニュー',
'// ============================================================',
'',
'function onOpen() {',
'  ReceiptEngine.onOpen();',
'}',
'',
'// ============================================================',
'// レシート処理',
'// ============================================================',
'',
'function processReceipts() {',
'  ReceiptEngine.processReceipts();',
'}',
'',
'function processReceipts_continue() {',
'  ReceiptEngine.processReceipts();',
'}',
'',
'function resetCheckErrorMarks() {',
'  ReceiptEngine.resetCheckErrorMarks();',
'}',
'',
'function resetProcessedMarks() {',
'  ReceiptEngine.resetProcessedMarks();',
'}',
'',
'function showSidebar() {',
'  ReceiptEngine.showSidebar();',
'}',
'',
'// ============================================================',
'// 検証・承認',
'// ============================================================',
'',
'function runAutoVerification() {',
'  ReceiptEngine.runAutoVerification();',
'}',
'',
'function runAutoVerification_continue() {',
'  ReceiptEngine.runAutoVerification();',
'}',
'',
'function verifySelectedRows() {',
'  ReceiptEngine.verifySelectedRows();',
'}',
'',
'function approveSelectedRows() {',
'  ReceiptEngine.approveSelectedRows();',
'}',
'',
'function applyVerificationFix(rowIndex, field, value) {',
'  return ReceiptEngine.applyVerificationFix(rowIndex, field, value);',
'}',
'',
'// ============================================================',
'// サイドバー用（google.script.runから呼ばれる）',
'// ============================================================',
'',
'function getSelectedRowData() {',
'  return ReceiptEngine.getSelectedRowData();',
'}',
'',
'function updateHandReceipt(row, totalAmount, taxable10, tax10, taxable8, tax8, nonTaxable) {',
'  return ReceiptEngine.updateHandReceipt(row, totalAmount, taxable10, tax10, taxable8, tax8, nonTaxable);',
'}',
'',
'function approveHandReceipt(row) {',
'  return ReceiptEngine.approveHandReceipt(row);',
'}',
'',
'// ============================================================',
'// データ修正',
'// ============================================================',
'',
'function fixExistingAccountTitles() {',
'  ReceiptEngine.fixExistingAccountTitles();',
'}',
'',
'function normalizeExistingStoreNames() {',
'  ReceiptEngine.normalizeExistingStoreNames();',
'}',
'',
'function deleteSelectedFiles() {',
'  ReceiptEngine.deleteSelectedFiles();',
'}',
'',
'// ============================================================',
'// 設定',
'// ============================================================',
'',
'function setupConfigSheet() {',
'  ReceiptEngine.setupConfigSheet();',
'}',
'',
'function createConfigFoldersSheet() {',
'  ReceiptEngine.createConfigFoldersSheet();',
'}',
'',
'function createConfigMappingSheet() {',
'  ReceiptEngine.createConfigMappingSheet();',
'}',
'',
'function promptGeminiApiKey() {',
'  ReceiptEngine.promptGeminiApiKey();',
'}',
'',
'// ============================================================',
'// クレカ突合',
'// ============================================================',
'',
'function reconcileWithStatements() {',
'  ReceiptEngine.reconcileWithStatements();',
'}',
'',
'function resetReconcileInfo() {',
'  ReceiptEngine.resetReconcileInfo();',
'}',
'',
'function promptStatementSpreadsheetId() {',
'  ReceiptEngine.promptStatementSpreadsheetId();',
'}',
'',
'// ============================================================',
'// エクスポート',
'// ============================================================',
'',
'function exportToYayoiCSV() {',
'  ReceiptEngine.exportToYayoiCSV();',
'}',
'',
'// ============================================================',
'// 一括読み込み',
'// ============================================================',
'',
'function processAll() {',
'  ReceiptEngine.processAll();',
'}',
'',
'// ============================================================',
'// 通帳処理',
'// ============================================================',
'',
'function processPassbooks() {',
'  const ss = SpreadsheetApp.getActiveSpreadsheet();',
'  const folderId = ReceiptEngine.getPassbookFolderId();',
'',
'  if (!folderId) {',
'    SpreadsheetApp.getUi().alert(\'エラー: 通帳フォルダIDが設定されていません。\\nClientConfigシートを確認してください。\');',
'    return;',
'  }',
'',
'  try {',
'    const count = ReceiptEngine.processPassbookFolder(folderId, ss);',
'    SpreadsheetApp.getUi().alert(\'完了: \' + count + \'件の通帳を処理しました。\');',
'  } catch (e) {',
'    SpreadsheetApp.getUi().alert(\'エラー: \' + e.message);',
'  }',
'}',
'',
'function createPassbookSheetOnly() {',
'  const ss = SpreadsheetApp.getActiveSpreadsheet();',
'  ReceiptEngine.createPassbookSheet(ss);',
'  SpreadsheetApp.getUi().alert(\'通帳シートを作成しました。\');',
'}',
'',
'// ============================================================',
'// ClientConfig',
'// ============================================================',
'',
'function createClientConfigSheet() {',
'  ReceiptEngine.createClientConfigSheet();',
'}'
  ].join('\n');

  PropertiesService.getScriptProperties().setProperty('WRAPPER_TEMPLATE', code);
  console.log('✅ ラッパーテンプレートを登録しました（' + code.length + '文字）');

  // UI表示（手動実行時のみ）
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('ラッパーテンプレートを登録しました（' + code.length + '文字）', '✅ 完了', 5);
  } catch (e) { /* トリガー実行時はUI非対応 */ }
}

// ============================================================
// サマリー表示
// ============================================================

/**
 * パートナー設定のサマリーを表示する。
 */
function showPartnerSummary() {
  const partners = loadPartnerConfigs_();

  if (partners.length === 0) {
    SpreadsheetApp.getUi().alert('有効なパートナー設定がありません。');
    return;
  }

  let msg = '登録パートナー: ' + partners.length + '件\n\n';
  for (const p of partners) {
    msg += '📋 ' + p.name + '\n';
    msg += '   シートID: ' + p.spreadsheetId.slice(0, 15) + '...\n';
    msg += '   アクティブ条件: ' + p.activeStatuses.join(', ') + '\n\n';
  }

  SpreadsheetApp.getUi().alert(msg);
}

/**
 * 全パートナー横断の顧客一覧サマリーを表示する。
 */
function showAllClientsSummary() {
  const clients = getAllActiveClients_();

  if (clients.length === 0) {
    SpreadsheetApp.getUi().alert('アクティブな顧客がいません。');
    return;
  }

  // パートナー別に集計
  const summary = {};
  for (const c of clients) {
    if (!summary[c.partner]) summary[c.partner] = 0;
    summary[c.partner]++;
  }

  let msg = '全アクティブ顧客: ' + clients.length + '件\n\n';
  for (const [partner, count] of Object.entries(summary)) {
    msg += partner + ': ' + count + '件\n';
  }

  SpreadsheetApp.getUi().alert(msg);
}

// ============================================================
// 実行ログ
// ============================================================

const LOG_SHEET_NAME = '実行ログ';

/**
 * 実行結果をログシートに記録する。
 * @param {string} operation
 * @param {number} total
 * @param {number} success
 * @param {number} errors
 */
function logExecution_(operation, total, success, errors) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LOG_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET_NAME);
    sheet.appendRow(['実行日時', '処理', '対象件数', '成功', 'エラー']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  sheet.appendRow([
    new Date(),
    operation,
    total,
    success,
    errors
  ]);
}

// ============================================================
// 初期セットアップ: KZ顧客管理シートにscriptId列を追加（一回限り）
// ============================================================

/**
 * KZ側の顧客管理シートにscriptId列を追加し、パートナー設定のI列を更新する。
 * GASエディタから1回だけ手動実行する。完了後は削除してよい。
 */
function setupKzScriptIdColumn() {
  // パートナー設定からKZ行を探す
  var centralSS = SpreadsheetApp.getActiveSpreadsheet();
  var partnerSheet = centralSS.getSheetByName('パートナー設定');
  if (!partnerSheet) {
    console.error('パートナー設定シートが見つかりません');
    return;
  }

  var partnerData = partnerSheet.getDataRange().getValues();
  var kzRow = -1;
  var kzSheetId = '';
  var kzSheetName = '';

  for (var r = 1; r < partnerData.length; r++) {
    var name = String(partnerData[r][0] || '').trim();
    if (name.indexOf('KZ') !== -1 || name.indexOf('絆') !== -1) {
      kzRow = r + 1; // 1始まり
      kzSheetId = String(partnerData[r][1] || '').trim();
      kzSheetName = String(partnerData[r][2] || '顧客管理').trim();
      break;
    }
  }

  if (!kzSheetId) {
    console.error('KZパートナーのシートIDが見つかりません');
    return;
  }

  console.log('KZパートナー: シートID=' + kzSheetId + ', シート名=' + kzSheetName);

  // KZ顧客管理シートを開く
  var ss = SpreadsheetApp.openById(kzSheetId);
  var sheet = ss.getSheetByName(kzSheetName);
  if (!sheet) {
    console.error('「' + kzSheetName + '」シートが見つかりません');
    return;
  }

  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  console.log('現在のヘッダー列数: ' + lastCol);
  for (var h = 0; h < headers.length; h++) {
    console.log('  [' + h + '] ' + (headers[h] || '(空)'));
  }

  // scriptIdヘッダーが既にあるか確認
  var scriptIdCol = -1;
  for (var c = 0; c < headers.length; c++) {
    if (String(headers[c]).trim() === 'scriptId') {
      scriptIdCol = c;
      break;
    }
  }

  if (scriptIdCol === -1) {
    scriptIdCol = lastCol; // 次の空き列
    sheet.getRange(1, scriptIdCol + 1).setValue('scriptId');
    sheet.getRange(1, scriptIdCol + 1).setFontWeight('bold');
    console.log('scriptId列を追加: 列 ' + scriptIdCol + ' (0始まり)');
  } else {
    console.log('scriptId列は既に存在: 列 ' + scriptIdCol + ' (0始まり)');
  }

  // パートナー設定のI列を更新
  partnerSheet.getRange(kzRow, 9).setValue(scriptIdCol); // I列 = 9列目
  console.log('パートナー設定更新: KZ行(' + kzRow + ') のscriptId列 → ' + scriptIdCol);

  console.log('========================================');
  console.log('✅ KZセットアップ完了');
  console.log('scriptId列: ' + scriptIdCol + ' (0始まり)');
  console.log('========================================');
}


