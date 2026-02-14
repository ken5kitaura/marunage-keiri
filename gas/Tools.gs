/**
 * Tools.gs
 * ワンタイム実行・メンテナンス用ユーティリティ
 */

/**
 * 【ワンタイム実行用】
 * シート上の既存データの「勘定科目」を、現在のMapping辞書に基づいて再判定し、上書き修正する関数。
 * ※辞書(Mapping.js)にヒットする店名だけを更新し、ヒットしないものはそのまま残します。
 */
function fixExistingAccountTitles() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME.MAIN);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('シート「' + CONFIG.SHEET_NAME.MAIN + '」が見つかりません。');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert('データがありません。');
    return;
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // ヘッダーから列インデックスを特定
  const colStore = headers.indexOf('利用店舗名');
  const colAccount = headers.indexOf('勘定科目');
  const colNonTaxable = headers.indexOf('不課税');

  if (colStore === -1 || colAccount === -1) {
    SpreadsheetApp.getUi().alert('必須のカラム（利用店舗名 / 勘定科目）が見つかりません。');
    return;
  }

  console.log('修正開始: 全' + (data.length - 1) + '行をチェックします...');

  let updatedCount = 0;
  const updates = []; // バッチ更新用

  for (let i = 1; i < data.length; i++) {
    const storeName = String(data[i][colStore] || '');
    const currentAccount = String(data[i][colAccount] || '');

    if (!storeName) continue;

    // L列（不課税）の値を軽油税情報として渡す
    // ※不課税列に値がある場合、軽油税の可能性がある（ガソリンスタンドなら確定）
    const nonTaxableValue = colNonTaxable !== -1 ? (Number(data[i][colNonTaxable]) || 0) : 0;
    const ocrForDiesel = nonTaxableValue > 0 ? { _subtotalInfo: { dieselTax: nonTaxableValue } } : null;

    // Mapping辞書から正解を引く（軽油税情報も渡す）
    const correctAccount = inferAccountTitleFromStore(storeName, [], null, ocrForDiesel);

    // 辞書にヒットし、かつ現在の値と違う場合のみ更新
    if (correctAccount && correctAccount !== currentAccount) {
      updates.push({
        row: i + 1,
        col: colAccount + 1,
        oldValue: currentAccount,
        newValue: correctAccount,
        storeName: storeName
      });
    }
  }

  // バッチでセルを更新
  for (const u of updates) {
    sheet.getRange(u.row, u.col).setValue(u.newValue);
    console.log('[行' + u.row + '] 更新: ' + u.storeName + ' | ' + u.oldValue + ' -> ' + u.newValue);
    updatedCount++;
  }

  const message = '勘定科目の修正が完了しました。\n' +
    '対象: ' + (data.length - 1) + '行\n' +
    '更新: ' + updatedCount + '件';
  console.log(message);
  SpreadsheetApp.getUi().alert(message);
}

/**
 * 【ワンタイム実行用】
 * 既存の「利用店舗名」を正規化（名寄せ）する。
 * - 法人格（株式会社・有限会社等）を除去
 * - 支店名・店舗番号（〇〇店・〇〇支店・No.XX等）を除去
 * - 住所・電話番号・郵便番号を除去
 * - 前後の空白・記号を整理
 * 変更前後をログに出力し、確認ダイアログ後に一括更新する。
 */
function normalizeExistingStoreNames() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME.MAIN);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('シート「' + CONFIG.SHEET_NAME.MAIN + '」が見つかりません。');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert('データがありません。');
    return;
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colStore = headers.indexOf('利用店舗名');

  if (colStore === -1) {
    SpreadsheetApp.getUi().alert('「利用店舗名」カラムが見つかりません。');
    return;
  }

  const updates = [];

  for (var i = 1; i < data.length; i++) {
    var original = String(data[i][colStore] || '').trim();
    if (!original) continue;

    var normalized = normalizeStoreName_(original);

    if (normalized && normalized !== original) {
      updates.push({
        row: i + 1,
        col: colStore + 1,
        oldValue: original,
        newValue: normalized
      });
    }
  }

  if (updates.length === 0) {
    SpreadsheetApp.getUi().alert('変更対象はありませんでした。');
    return;
  }

  // プレビューをログに出力
  console.log('=== 店舗名 正規化プレビュー (' + updates.length + '件) ===');
  for (var j = 0; j < updates.length; j++) {
    console.log('[行' + updates[j].row + '] ' + updates[j].oldValue + '  →  ' + updates[j].newValue);
  }

  // 確認ダイアログ（先頭10件を表示）
  var preview = updates.slice(0, 10).map(function(u) {
    return u.oldValue + '  →  ' + u.newValue;
  }).join('\n');
  if (updates.length > 10) {
    preview += '\n... 他 ' + (updates.length - 10) + '件（全件はログで確認）';
  }

  var confirm = SpreadsheetApp.getUi().alert(
    '店舗名の正規化',
    updates.length + '件の店舗名を変更します。\n\n' + preview + '\n\n実行しますか？',
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );

  if (confirm !== SpreadsheetApp.getUi().Button.YES) {
    SpreadsheetApp.getUi().alert('キャンセルしました。');
    return;
  }

  // 一括更新
  for (var k = 0; k < updates.length; k++) {
    sheet.getRange(updates[k].row, updates[k].col).setValue(updates[k].newValue);
  }

  // 正規化後に勘定科目も再判定
  var accountUpdated = 0;
  var colAccount = headers.indexOf('勘定科目');
  var colNonTaxable2 = headers.indexOf('不課税');
  if (colAccount !== -1) {
    for (var m = 0; m < updates.length; m++) {
      // L列（不課税）の値を軽油税情報として渡す
      var dataRowIdx = updates[m].row - 1; // data配列のインデックス（0-based）
      var nonTaxVal = (colNonTaxable2 !== -1 && dataRowIdx < data.length) ? (Number(data[dataRowIdx][colNonTaxable2]) || 0) : 0;
      var ocrForDiesel2 = nonTaxVal > 0 ? { _subtotalInfo: { dieselTax: nonTaxVal } } : null;

      var correctAccount = inferAccountTitleFromStore(updates[m].newValue, [], null, ocrForDiesel2);
      if (correctAccount) {
        var currentAccount = String(sheet.getRange(updates[m].row, colAccount + 1).getValue() || '');
        if (correctAccount !== currentAccount) {
          sheet.getRange(updates[m].row, colAccount + 1).setValue(correctAccount);
          accountUpdated++;
        }
      }
    }
  }

  var msg = '正規化完了\n店舗名更新: ' + updates.length + '件';
  if (accountUpdated > 0) {
    msg += '\n勘定科目も再判定: ' + accountUpdated + '件';
  }
  console.log(msg);
  SpreadsheetApp.getUi().alert(msg);
}

/**
 * 店舗名を正規化するルール
 * @param {string} name
 * @return {string}
 */
function normalizeStoreName_(name) {
  var s = name;

  // 1. 法人格を除去
  s = s.replace(/株式会社|㈱|（株）|\(株\)/g, '');
  s = s.replace(/有限会社|㈲|（有）|\(有\)/g, '');
  s = s.replace(/合同会社|合名会社|合資会社/g, '');
  s = s.replace(/一般社団法人|一般財団法人|公益社団法人|公益財団法人/g, '');
  s = s.replace(/社会福祉法人|医療法人|学校法人|宗教法人|特定非営利活動法人|NPO法人/g, '');

  // 2. 支店名・店舗番号を除去
  //    ★最後のスペース区切り部分のみ対象（スペースを越えてマッチしない）
  //    ★ブランド名自体を削らないよう保守的に判定
  // 「〇〇支店」「〇〇営業所」「〇〇出張所」（これらは確実に支店名）
  s = s.replace(/[　\s]+[^\s　]*(支店|営業所|出張所)$/g, '');
  // 「〇〇通り店」「〇〇駅前店」「〇〇XX号店」のような具体的な支店名
  s = s.replace(/[　\s]+[^\s　]*?(通り|駅前|駅ビル)[^\s　]*?店$/g, '');
  s = s.replace(/[　\s]+\S*\d+号?店$/g, '');
  // 「豊中店」「梅田店」のような地名+店（スペース区切りの場合のみ）
  s = s.replace(/[　\s]+[^\s　]{1,6}店$/g, '');

  // 3. 電話番号を除去（TEL/FAX/Tel含む）
  s = s.replace(/(?:TEL|FAX|Tel|Fax|tel|fax)[：:\s]*[\d\-()（）]+/g, '');
  s = s.replace(/\d{2,4}[-ー]\d{2,4}[-ー]\d{3,4}/g, '');

  // 4. 郵便番号を除去
  s = s.replace(/〒?\s*\d{3}[-ー]\d{4}/g, '');

  // 5. 住所パターンを除去（都道府県から始まる文字列）
  s = s.replace(/[　\s]+(?:北海道|東京都|大阪府|京都府|.{2,3}県).+$/g, '');

  // 6. 「No.XX」「#XX」等の番号を除去
  s = s.replace(/[　\s]*(?:No\.?|#|Ｎｏ．?)\s*\d+/gi, '');

  // 7. 全角スペース→半角、連続スペース→1つ、前後トリム
  s = s.replace(/　/g, ' ');
  s = s.replace(/\s+/g, ' ');
  s = s.trim();

  // 8. 先頭・末尾の記号を除去
  s = s.replace(/^[\s・\-_／/]+|[\s・\-_／/]+$/g, '');

  // 空文字になったら元の名前を返す（過剰除去防止）
  if (!s || s.length < 2) {
    return name.trim();
  }

  return s;
}

/**
 * 選択行のファイルを「削除済み」フォルダへ移動し、スプシの行も削除する
 * 共有ドライブでは投稿者権限ではゴミ箱操作ができないため、
 * ソースフォルダ内の「削除済み」サブフォルダに移動する方式を採用
 *
 * パフォーマンス最適化:
 * - データを一括読み取り（セル個別読み取りを回避）
 * - fetchAll() で親フォルダ取得・ファイル移動を並列実行
 * - 複数ソースフォルダに対応（ファイルごとに正しい親を使用）
 * - 入れ子防止: 親が「削除済み」の場合はスキップ
 */
function deleteSelectedFiles() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  // 本番シートかチェック
  if (sheet.getName() !== CONFIG.SHEET_NAME.MAIN) {
    ui.alert('「' + CONFIG.SHEET_NAME.MAIN + '」シートで実行してください。');
    return;
  }

  const selection = sheet.getActiveRange();
  if (!selection) {
    ui.alert('削除したい行を選択してください。');
    return;
  }

  const startRow = selection.getRow();
  const numRows = selection.getNumRows();

  // ヘッダー行は除外
  if (startRow <= 1) {
    ui.alert('ヘッダー行は削除できません。データ行を選択してください。');
    return;
  }

  // ヘッダーから列インデックスを取得
  const lastCol = sheet.getLastColumn();
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const idxStore = findHeaderIndex(header, ['利用店舗名', '店舗名']);
  const idxFileName = findHeaderIndex(header, ['ファイル名']);

  // 対象範囲を一括読み取り（値 + B列の数式）
  const values = sheet.getRange(startRow, 1, numRows, lastCol).getValues();
  const formulas = sheet.getRange(startRow, 2, numRows, 1).getFormulas();

  // 対象ファイル情報を収集
  const targets = [];
  for (let i = 0; i < numRows; i++) {
    const formula = formulas[i][0] || '';
    const urlMatch = formula.match(/HYPERLINK\("([^"]+)"/);
    let fileId = '';
    if (urlMatch) {
      const idMatch = urlMatch[1].match(/\/d\/([^\/]+)/);
      if (idMatch) fileId = idMatch[1];
    }

    const storeName = idxStore >= 0 ? String(values[i][idxStore] || '') : '';
    const fileName = idxFileName >= 0 ? String(values[i][idxFileName] || '') : '';

    targets.push({
      row: startRow + i,
      fileId: fileId,
      storeName: storeName,
      fileName: fileName
    });
  }

  const validTargets = targets.filter(function(t) { return t.fileId; });

  if (validTargets.length === 0) {
    ui.alert('削除対象のファイルが見つかりません。');
    return;
  }

  // 確認ダイアログ（件数が多い場合は先頭20件のみ表示）
  const displayTargets = targets.slice(0, 20);
  let fileList = displayTargets.map(function(t) {
    return '・' + (t.storeName || t.fileName || '(不明)');
  }).join('\n');
  if (targets.length > 20) {
    fileList += '\n... 他 ' + (targets.length - 20) + '件';
  }

  const confirmResult = ui.alert(
    'ファイル移動確認',
    targets.length + '件のファイルを「削除済み」フォルダへ移動し、スプレッドシートの行を削除します。\n\n' + fileList + '\n\nよろしいですか？',
    ui.ButtonSet.YES_NO
  );

  if (confirmResult !== ui.Button.YES) {
    return;
  }

  const token = ScriptApp.getOAuthToken();
  const BATCH_SIZE = 100;

  // ── Step 1: 全ファイルの親フォルダを並列取得 ──
  const parentMap = {}; // fileId -> parentId
  for (let batchStart = 0; batchStart < validTargets.length; batchStart += BATCH_SIZE) {
    const batch = validTargets.slice(batchStart, batchStart + BATCH_SIZE);
    const requests = batch.map(function(t) {
      return {
        url: 'https://www.googleapis.com/drive/v3/files/' + t.fileId +
             '?fields=parents&supportsAllDrives=true&includeItemsFromAllDrives=true',
        method: 'get',
        headers: { 'Authorization': 'Bearer ' + token },
        muteHttpExceptions: true
      };
    });

    const responses = UrlFetchApp.fetchAll(requests);
    for (let i = 0; i < responses.length; i++) {
      if (responses[i].getResponseCode() >= 200 && responses[i].getResponseCode() < 300) {
        try {
          const info = JSON.parse(responses[i].getContentText());
          if (info.parents && info.parents[0]) {
            parentMap[batch[i].fileId] = info.parents[0];
          }
        } catch (e) { /* ignore */ }
      }
    }
    console.log('親フォルダ取得: ' + (batchStart + batch.length) + '/' + validTargets.length);
  }

  // ── Step 2: 親フォルダごとに「削除済み」フォルダを取得/作成 ──
  const uniqueParentIds = [];
  const seen = {};
  for (var fid in parentMap) {
    var pid = parentMap[fid];
    if (!seen[pid]) {
      seen[pid] = true;
      uniqueParentIds.push(pid);
    }
  }

  const deletedFolderMap = {}; // parentId -> deletedFolderId
  for (const pid of uniqueParentIds) {
    try {
      deletedFolderMap[pid] = getOrCreateDeletedFolder_(token, pid);
    } catch (e) {
      console.error('「削除済み」フォルダ作成失敗 (parent=' + pid + '): ' + e.message);
    }
  }

  // ── Step 3: fetchAll() で並列移動 ──
  let movedCount = 0;
  let errorCount = 0;
  const errors = [];
  const movedFileIds = new Set();

  // fileIdなしのエラー
  const noIdErrors = targets.filter(function(t) { return !t.fileId; });
  errorCount += noIdErrors.length;
  for (const t of noIdErrors) {
    errors.push((t.storeName || t.fileName) + ': ファイルIDが取得できません');
  }

  // 移動可能なファイルを抽出（親が判明 & 削除済みフォルダが存在）
  const movableTargets = validTargets.filter(function(t) {
    var pid = parentMap[t.fileId];
    if (!pid) {
      errorCount++;
      errors.push((t.storeName || t.fileName) + ': 親フォルダ不明');
      return false;
    }
    if (!deletedFolderMap[pid]) {
      errorCount++;
      errors.push((t.storeName || t.fileName) + ': 削除済みフォルダ作成失敗');
      return false;
    }
    return true;
  });

  for (let batchStart = 0; batchStart < movableTargets.length; batchStart += BATCH_SIZE) {
    const batch = movableTargets.slice(batchStart, batchStart + BATCH_SIZE);
    const requests = batch.map(function(t) {
      var pid = parentMap[t.fileId];
      var destId = deletedFolderMap[pid];
      return {
        url: 'https://www.googleapis.com/drive/v3/files/' + t.fileId +
             '?addParents=' + destId +
             '&removeParents=' + pid +
             '&supportsAllDrives=true',
        method: 'patch',
        headers: { 'Authorization': 'Bearer ' + token },
        contentType: 'application/json',
        payload: '{}',
        muteHttpExceptions: true
      };
    });

    const responses = UrlFetchApp.fetchAll(requests);

    for (let i = 0; i < responses.length; i++) {
      const code = responses[i].getResponseCode();
      if (code >= 200 && code < 300) {
        movedCount++;
        movedFileIds.add(batch[i].fileId);
      } else {
        errorCount++;
        let errMsg = 'HTTP ' + code;
        try {
          const err = JSON.parse(responses[i].getContentText());
          errMsg = err.error.message || errMsg;
        } catch (e) { /* ignore */ }
        errors.push((batch[i].storeName || batch[i].fileName) + ': ' + errMsg);
      }
    }

    console.log('バッチ移動完了: ' + (batchStart + batch.length) + '/' + movableTargets.length);
  }

  // ── Step 4: 成功した行を削除（下から順にインデックスずれ防止）──
  if (movedCount > 0) {
    const rowsToDelete = targets
      .filter(function(t) { return movedFileIds.has(t.fileId); })
      .map(function(t) { return t.row; })
      .sort(function(a, b) { return b - a; });
    for (const row of rowsToDelete) {
      sheet.deleteRow(row);
    }
  }

  // 結果表示
  if (errorCount === 0) {
    ui.alert('完了\n「削除済み」フォルダへ移動: ' + movedCount + '件');
  } else {
    const errorPreview = errors.slice(0, 10).join('\n');
    const errorSuffix = errors.length > 10 ? '\n... 他 ' + (errors.length - 10) + '件（ログで確認）' : '';
    ui.alert('結果\n移動成功: ' + movedCount + '件\n失敗: ' + errorCount + '件\n\n' + errorPreview + errorSuffix);
    console.log('移動エラー詳細:\n' + errors.join('\n'));
  }
}

// ============================================================
// Drive REST API ヘルパー（共有ドライブ対応）
// ============================================================

/**
 * 親フォルダ内の「削除済み」サブフォルダを取得、なければ作成
 * 入れ子防止: 親自体が「削除済み」の場合はそのまま返す
 * @param {string} token
 * @param {string} parentId
 * @return {string} folderId
 */
function getOrCreateDeletedFolder_(token, parentId) {
  // 入れ子防止: 親フォルダ名が「削除済み」なら、その中にさらに作らない
  var parentInfo = driveApiRequest_(token, 'GET',
    'https://www.googleapis.com/drive/v3/files/' + parentId +
    '?fields=name&supportsAllDrives=true&includeItemsFromAllDrives=true');
  if (parentInfo.name === '削除済み') {
    console.log('入れ子防止: 親が既に「削除済み」フォルダのためスキップ (id=' + parentId + ')');
    return parentId;
  }

  // 既存の「削除済み」フォルダを検索
  var q = "name='削除済み' and mimeType='application/vnd.google-apps.folder' and '" + parentId + "' in parents and trashed=false";
  var searchUrl = 'https://www.googleapis.com/drive/v3/files' +
    '?q=' + encodeURIComponent(q) +
    '&supportsAllDrives=true&includeItemsFromAllDrives=true' +
    '&corpora=allDrives&fields=files(id)';

  var result = driveApiRequest_(token, 'GET', searchUrl);

  if (result.files && result.files.length > 0) {
    return result.files[0].id;
  }

  // なければ作成
  var createResult = driveApiRequest_(token, 'POST',
    'https://www.googleapis.com/drive/v3/files?supportsAllDrives=true',
    {
      name: '削除済み',
      mimeType: 'application/vnd.google-apps.folder',
      parents: [parentId]
    });

  return createResult.id;
}

/**
 * Drive REST API 汎用リクエスト
 * @param {string} token
 * @param {string} method - GET, POST, PATCH
 * @param {string} url
 * @param {Object} [body] - リクエストボディ（POST/PATCH時）
 * @return {Object} レスポンスJSON
 */
function driveApiRequest_(token, method, url, body) {
  var options = {
    method: method.toLowerCase(),
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true
  };

  if (body !== undefined) {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(body);
  }

  var response = UrlFetchApp.fetch(url, options);
  var code = response.getResponseCode();
  var text = response.getContentText();

  if (code < 200 || code >= 300) {
    var errMsg = 'HTTP ' + code;
    try {
      var err = JSON.parse(text);
      errMsg = err.error.message || errMsg;
    } catch (e) { /* ignore parse error */ }
    throw new Error(errMsg);
  }

  return text ? JSON.parse(text) : {};
}
