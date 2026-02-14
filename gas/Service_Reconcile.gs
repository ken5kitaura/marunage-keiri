/**
 * Service_Reconcile.gs
 * クレジットカード明細との突合処理
 *
 * 旧cc_reconcile.jsから移植・整理
 * - 日付Window（店舗グループ別に可変）
 * - 店名類似度スコアリング
 * - 双方向の書き戻し
 */

/**
 * 明細行データ
 * @typedef {Object} StatementRow
 * @property {string} tabName - タブ名
 * @property {number} rowNumber - 行番号
 * @property {Date} date - 利用日
 * @property {string} merchant - 利用先
 * @property {number} amount - 金額
 */

// ============================================================
// メイン突合処理
// ============================================================

/**
 * 突合処理を実行
 */
function runReconciliation() {
  const receiptSS = SpreadsheetApp.getActiveSpreadsheet();
  const statementSSId = getCCStatementSpreadsheetId_();

  if (!statementSSId) {
    SpreadsheetApp.getUi().alert('クレカ明細スプレッドシートが設定されていません。');
    return;
  }

  let statementSS;
  try {
    statementSS = SpreadsheetApp.openById(statementSSId);
  } catch (e) {
    SpreadsheetApp.getUi().alert('明細スプレッドシートを開けません: ' + e.message);
    return;
  }

  // 1. 明細タブを列挙
  const statementTabs = listStatementTabs_(statementSS);
  if (statementTabs.length === 0) {
    SpreadsheetApp.getUi().alert('YYYYMM形式のタブが見つかりませんでした。');
    return;
  }

  console.log('明細タブ数: ' + statementTabs.length);

  // 2. 明細行を読み込み
  const allStatements = [];
  for (const tab of statementTabs) {
    const rows = readStatementRows_(tab);
    Array.prototype.push.apply(allStatements, rows);
  }

  console.log('明細行数: ' + allStatements.length);

  if (allStatements.length === 0) {
    SpreadsheetApp.getUi().alert('明細データが見つかりませんでした。');
    return;
  }

  // 3. レシートシートを取得
  const receiptSheet = receiptSS.getSheetByName(CONFIG.SHEET_NAME.MAIN);
  if (!receiptSheet) {
    SpreadsheetApp.getUi().alert('レシートシートが見つかりません。');
    return;
  }

  // 4. 突合実行
  const matchCount = performReconciliation_(receiptSheet, allStatements, statementSS);

  SpreadsheetApp.getUi().alert('突合完了: ' + matchCount + '件が一致しました。');
}

/**
 * YYYYMM形式のタブを列挙
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @return {Array<GoogleAppsScript.Spreadsheet.Sheet>}
 */
function listStatementTabs_(ss) {
  const sheets = ss.getSheets();
  const pattern = /^\d{6}$/;
  return sheets.filter(s => pattern.test(s.getName()));
}

/**
 * 明細シートから行データを読み取り
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @return {Array<StatementRow>}
 */
function readStatementRows_(sheet) {
  const tabName = sheet.getName();
  const lastRow = sheet.getLastRow();

  if (lastRow < 3) return [];

  // A:D列を取得（A:利用日, B:利用店名, C:利用金額, D:備考）
  const data = sheet.getRange(3, 1, lastRow - 2, 4).getValues();
  const rows = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const date = parseDate(row[0]);
    if (!date) continue;

    const merchant = String(row[1] || '').trim();
    const amount = parseAmount(row[2]);

    if (!merchant || !amount || amount <= 0) continue;

    rows.push({
      tabName: tabName,
      rowNumber: i + 3,
      date: date,
      merchant: merchant,
      amount: amount
    });
  }

  return rows;
}

/**
 * 突合処理の実行
 * @param {GoogleAppsScript.Spreadsheet.Sheet} receiptSheet
 * @param {Array<StatementRow>} statements
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} statementSS
 * @return {number} マッチ数
 */
function performReconciliation_(receiptSheet, statements, statementSS) {
  const lastRow = receiptSheet.getLastRow();
  const lastCol = receiptSheet.getLastColumn();

  if (lastRow < 2) return 0;

  // ヘッダー取得
  const header = receiptSheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // 列インデックス取得
  const idxDate = findHeaderIndex(header, ['日付', '利用日', '購入日']);
  const idxStore = findHeaderIndex(header, ['利用店舗名', '店舗名', '店名']);
  const idxAmount = findHeaderIndex(header, ['総合計', '合計金額', '金額']);
  const idxFileName = findHeaderIndex(header, ['ファイル名']);
  const idxReconcileStatus = findHeaderIndex(header, ['突合ステータス']);
  const idxCreditAccount = findHeaderIndex(header, ['貸方科目']);

  if (idxDate === -1 || idxStore === -1 || idxAmount === -1) {
    console.error('必要な列が見つかりません');
    return 0;
  }

  // 必要列を追加
  ensureReconcileColumns_(receiptSheet);

  // ヘッダー再取得
  const newLastCol = receiptSheet.getLastColumn();
  const newHeader = receiptSheet.getRange(1, 1, 1, newLastCol).getValues()[0];
  const idxStatementId = findHeaderIndex(newHeader, ['明細ID']);
  const idxScore = findHeaderIndex(newHeader, ['突合スコア']);

  // データ取得
  const receiptData = receiptSheet.getRange(2, 1, lastRow - 1, newLastCol).getValues();

  let matchCount = 0;

  for (let i = 0; i < receiptData.length; i++) {
    const row = receiptData[i];
    const rowNum = i + 2;

    const receiptDate = parseDate(row[idxDate]);
    const receiptStore = String(row[idxStore] || '');
    const receiptAmount = parseAmount(row[idxAmount]);
    const receiptFileName = row[idxFileName] || '';

    if (!receiptDate || !receiptAmount || receiptAmount <= 0) continue;

    // 既に突合済みならスキップ
    const currentStatus = idxReconcileStatus !== -1 ? String(row[idxReconcileStatus] || '') : '';
    if (currentStatus === 'MATCHED') continue;

    // 店舗グループ判定
    const receiptGroupKey = detectMerchantGroup_(receiptStore);

    // 候補を探す
    const candidates = findCandidates_(
      receiptDate, receiptStore, receiptAmount, receiptGroupKey, statements
    );

    if (candidates.length === 0) continue;

    // 最高スコアの候補
    const best = candidates[0];

    // 同点が複数ある場合
    const sameScoreCount = candidates.filter(c => c.score === best.score).length;

    if (sameScoreCount > 1) {
      // MULTI
      if (idxReconcileStatus !== -1) {
        receiptSheet.getRange(rowNum, idxReconcileStatus + 1).setValue('MULTI');
      }
      continue;
    }

    // スコアが閾値未満
    if (best.score < CONFIG.RECONCILE.MIN_MATCH_SCORE) continue;

    // マッチ
    matchCount++;

    // レシート側を更新
    if (idxReconcileStatus !== -1) {
      receiptSheet.getRange(rowNum, idxReconcileStatus + 1).setValue('MATCHED');
    }
    if (idxStatementId !== -1) {
      const stmtId = best.statement.tabName + '_' + best.statement.rowNumber;
      receiptSheet.getRange(rowNum, idxStatementId + 1).setValue(stmtId);
    }
    if (idxScore !== -1) {
      receiptSheet.getRange(rowNum, idxScore + 1).setValue(best.score);
    }
    if (idxCreditAccount !== -1) {
      receiptSheet.getRange(rowNum, idxCreditAccount + 1).setValue('クレジットカード');
    }

    // 明細側を更新
    updateStatementRow_(statementSS, best.statement, receiptFileName, rowNum, best.score);
  }

  return matchCount;
}

/**
 * 候補を検索
 * @param {Date} receiptDate
 * @param {string} receiptStore
 * @param {number} receiptAmount
 * @param {string|null} receiptGroupKey
 * @param {Array<StatementRow>} statements
 * @return {Array<{statement: StatementRow, score: number}>}
 */
function findCandidates_(receiptDate, receiptStore, receiptAmount, receiptGroupKey, statements) {
  const candidates = [];

  for (const stmt of statements) {
    // 金額一致チェック（必須）
    if (Math.abs(stmt.amount - receiptAmount) > 0.5) continue;

    // 明細グループ判定
    const stmtGroupKey = detectMerchantGroup_(stmt.merchant);

    // 日付許容範囲
    let dateWindowDays = CONFIG.RECONCILE.DEFAULT_DATE_WINDOW_DAYS;
    if (receiptGroupKey && receiptGroupKey === stmtGroupKey) {
      const group = CONFIG.RECONCILE.MERCHANT_GROUPS.find(g => g.key === receiptGroupKey);
      if (group) dateWindowDays = group.dateWindowDays;
    }

    // 日付チェック
    const daysDiff = dateDiffDays(receiptDate, stmt.date);
    if (daysDiff > dateWindowDays) continue;

    // 店名類似度
    let merchantScore = similarityScore(receiptStore, stmt.merchant);
    if (receiptGroupKey && receiptGroupKey === stmtGroupKey) {
      merchantScore = 100;  // 同一グループなら満点
    }

    // 総合スコア（金額=50, 日付=25, 店名=25）
    const amountScore = 50;
    const dateScore = daysDiff === 0 ? 25 : (daysDiff === 1 ? 20 : 15);
    const totalScore = amountScore + dateScore + Math.min(25, Math.floor(merchantScore * 0.25));

    candidates.push({
      statement: stmt,
      score: totalScore
    });
  }

  // スコア降順ソート
  candidates.sort((a, b) => b.score - a.score);

  return candidates;
}

/**
 * 店舗グループを判定
 * @param {string} merchant
 * @return {string|null}
 */
function detectMerchantGroup_(merchant) {
  if (!merchant) return null;

  const normalized = normalizeString(merchant);

  for (const group of CONFIG.RECONCILE.MERCHANT_GROUPS) {
    for (const keyword of group.keywords) {
      const normalizedKeyword = normalizeString(keyword);
      if (normalized.indexOf(normalizedKeyword) !== -1) {
        return group.key;
      }
    }
  }

  return null;
}

/**
 * 突合用列を追加
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function ensureReconcileColumns_(sheet) {
  const lastCol = sheet.getLastColumn();
  const header = sheet.getRange(1, 1, 1, Math.max(lastCol, 1)).getValues()[0];
  const existing = header.map(h => String(h || '').trim());

  const required = ['突合ステータス', '明細ID', '突合スコア'];
  const toAdd = required.filter(col => existing.indexOf(col) === -1);

  if (toAdd.length > 0) {
    const startCol = existing.length + 1;
    sheet.getRange(1, startCol, 1, toAdd.length).setValues([toAdd]);
  }
}

/**
 * 明細側の行を更新
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {StatementRow} stmt
 * @param {string} receiptFileName
 * @param {number} receiptRowNum
 * @param {number} score
 */
function updateStatementRow_(ss, stmt, receiptFileName, receiptRowNum, score) {
  const sheet = ss.getSheetByName(stmt.tabName);
  if (!sheet) return;

  const lastCol = sheet.getLastColumn();
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // 必要列を追加
  const required = ['レシート突合ステータス', 'レシート_ファイル名', 'レシート_行番号', '突合スコア'];
  const existing = header.map(h => String(h || '').trim());
  const toAdd = required.filter(col => existing.indexOf(col) === -1);

  if (toAdd.length > 0) {
    sheet.getRange(1, existing.length + 1, 1, toAdd.length).setValues([toAdd]);
  }

  // 列インデックス再取得
  const newLastCol = sheet.getLastColumn();
  const newHeader = sheet.getRange(1, 1, 1, newLastCol).getValues()[0];

  const idxStatus = findHeaderIndex(newHeader, ['レシート突合ステータス']);
  const idxFileName = findHeaderIndex(newHeader, ['レシート_ファイル名']);
  const idxRowNum = findHeaderIndex(newHeader, ['レシート_行番号']);
  const idxScore = findHeaderIndex(newHeader, ['突合スコア']);

  if (idxStatus !== -1) {
    sheet.getRange(stmt.rowNumber, idxStatus + 1).setValue('MATCHED');
  }
  if (idxFileName !== -1) {
    sheet.getRange(stmt.rowNumber, idxFileName + 1).setValue(receiptFileName);
  }
  if (idxRowNum !== -1) {
    sheet.getRange(stmt.rowNumber, idxRowNum + 1).setValue(receiptRowNum);
  }
  if (idxScore !== -1) {
    sheet.getRange(stmt.rowNumber, idxScore + 1).setValue(score);
  }
}

// ============================================================
// リセット処理
// ============================================================

/**
 * 突合情報をリセット
 */
function resetReconciliation() {
  const receiptSS = SpreadsheetApp.getActiveSpreadsheet();
  const receiptSheet = receiptSS.getSheetByName(CONFIG.SHEET_NAME.MAIN);

  if (!receiptSheet) {
    SpreadsheetApp.getUi().alert('レシートシートが見つかりません。');
    return;
  }

  const lastRow = receiptSheet.getLastRow();
  const lastCol = receiptSheet.getLastColumn();

  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('データがありません。');
    return;
  }

  const header = receiptSheet.getRange(1, 1, 1, lastCol).getValues()[0];

  const idxReconcileStatus = findHeaderIndex(header, ['突合ステータス']);
  const idxStatementId = findHeaderIndex(header, ['明細ID']);
  const idxScore = findHeaderIndex(header, ['突合スコア']);
  const idxCreditAccount = findHeaderIndex(header, ['貸方科目']);

  // データ取得
  const data = receiptSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  for (let i = 0; i < data.length; i++) {
    const rowNum = i + 2;
    const row = data[i];

    const currentStatus = idxReconcileStatus !== -1 ? String(row[idxReconcileStatus] || '') : '';

    if (currentStatus === 'MATCHED' || currentStatus === 'MULTI') {
      if (idxReconcileStatus !== -1) {
        receiptSheet.getRange(rowNum, idxReconcileStatus + 1).clearContent();
      }
      if (idxStatementId !== -1) {
        receiptSheet.getRange(rowNum, idxStatementId + 1).clearContent();
      }
      if (idxScore !== -1) {
        receiptSheet.getRange(rowNum, idxScore + 1).clearContent();
      }
      if (idxCreditAccount !== -1 && row[idxCreditAccount] === 'クレジットカード') {
        receiptSheet.getRange(rowNum, idxCreditAccount + 1).setValue('現金');
      }
    }
  }

  // 明細側もリセット
  const statementSSId = getCCStatementSpreadsheetId_();
  if (statementSSId) {
    try {
      const statementSS = SpreadsheetApp.openById(statementSSId);
      const tabs = listStatementTabs_(statementSS);

      for (const tab of tabs) {
        resetStatementTab_(tab);
      }
    } catch (e) {
      console.error('明細リセットエラー: ' + e.message);
    }
  }

  SpreadsheetApp.getUi().alert('突合情報をリセットしました。');
}

/**
 * 明細タブの突合情報をリセット
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function resetStatementTab_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 3) return;

  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  const cols = ['レシート突合ステータス', 'レシート_ファイル名', 'レシート_行番号', '突合スコア'];

  for (const colName of cols) {
    const idx = findHeaderIndex(header, [colName]);
    if (idx !== -1 && idx >= 4) {  // E列以降のみ
      sheet.getRange(3, idx + 1, lastRow - 2, 1).clearContent();
    }
  }
}
