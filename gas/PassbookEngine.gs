/**
 * Passbook Engine（通帳読み取りライブラリ）
 * 
 * 通帳画像をGemini APIで読み取り、各行の取引データを抽出して
 * スプレッドシートの「通帳」タブに記録する
 * 
 * ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 * 使い方（各顧客スプシ側）
 * ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 * 
 * 1. このスクリプトをライブラリとして追加
 * 2. 顧客スプシのGASに以下のラッパー関数を作成:
 * 
 *    function processPassbooks() {
 *      PassbookEngine.processPassbookFolder(
 *        'フォルダID',           // 通帳画像フォルダ
 *        SpreadsheetApp.getActiveSpreadsheet()
 *      );
 *    }
 * 
 * ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 */

// ============================================================
// 設定
// ============================================================

const PASSBOOK_CONFIG = {
  SHEET_NAME: '通帳',
  GEMINI_MODEL: 'gemini-2.0-flash',
  GEMINI_MAX_TOKENS: 8192,
  GEMINI_TEMPERATURE: 0.1
};

// ============================================================
// メイン処理
// ============================================================

/**
 * 通帳フォルダを処理するメイン関数
 * @param {string} folderId - 通帳画像が格納されたフォルダID
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - 出力先スプレッドシート
 * @param {string} [apiKey] - Gemini APIキー（省略時はScriptPropertiesから取得）
 */
function processPassbookFolder(folderId, spreadsheet, apiKey) {
  const geminiKey = apiKey || PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!geminiKey) {
    throw new Error('GEMINI_API_KEY が設定されていません');
  }
  
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  
  // Step 1: 未処理ファイルを収集
  const unprocessedFiles = [];
  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    
    // 処理済みスキップ
    if (isPassbookProcessed_(fileName)) {
      continue;
    }
    
    // 対応MIMEタイプチェック
    const mime = file.getMimeType();
    if (!isSupportedPassbookMime_(mime)) {
      console.log('SKIP (非対応MIME): ' + fileName);
      continue;
    }
    
    unprocessedFiles.push(file);
  }
  
  if (unprocessedFiles.length === 0) {
    console.log('未処理の通帳ファイルはありません');
    return 0;
  }
  
  console.log('未処理ファイル数: ' + unprocessedFiles.length);
  
  // Step 2: 各ファイルをOCRして、ページ情報と取引データを取得
  const passbookPages = [];
  
  for (const file of unprocessedFiles) {
    const fileName = file.getName();
    const fileUrl = file.getUrl();
    
    try {
      console.log('読み取り中: ' + fileName);
      const passbookData = extractPassbookData_(file, geminiKey);
      
      if (!passbookData || !passbookData.transactions || passbookData.transactions.length === 0) {
        console.warn('取引データが抽出できませんでした: ' + fileName);
        markPassbookAsError_(file);
        continue;
      }
      
      // ページ情報を保存
      passbookPages.push({
        file: file,
        fileName: fileName,
        fileUrl: fileUrl,
        pageNumber: passbookData.pageNumber || null,
        firstBalance: passbookData.transactions[0].balance,
        lastBalance: passbookData.transactions[passbookData.transactions.length - 1].balance,
        transactions: passbookData.transactions
      });
      
      console.log('読み取り完了: ' + fileName + ' (ページ: ' + (passbookData.pageNumber || '不明') + ', ' + passbookData.transactions.length + '件)');
      
    } catch (e) {
      console.error('処理エラー (' + fileName + '): ' + e.message);
      markPassbookAsError_(file);
    }
  }
  
  if (passbookPages.length === 0) {
    console.log('有効な通帳データがありません');
    return 0;
  }
  
  // Step 3: ページ順にソート
  const sortedPages = sortPassbookPages_(passbookPages);
  
  // Step 4: スプレッドシートに出力
  const sheet = getOrCreatePassbookSheet_(spreadsheet);
  
  for (const page of sortedPages) {
    for (const tx of page.transactions) {
      appendPassbookTransaction_(sheet, {
        date: tx.date,
        description: tx.description,
        deposit: tx.deposit,
        withdrawal: tx.withdrawal,
        balance: tx.balance,
        accountTitle: '',
        subAccount: '',
        fileUrl: page.fileUrl,
        fileName: page.fileName,
        status: '未確認'
      });
    }
    
    // ファイルに処理済みマークを付与
    markPassbookAsProcessed_(page.file);
  }
  
  console.log('通帳処理完了: ' + sortedPages.length + '件');
  return sortedPages.length;
}

/**
 * 通帳ページをソート
 * 1. ページ番号がある場合はページ番号順
 * 2. ページ番号がない場合は残高の連続性でソート
 * @param {Array} pages
 * @return {Array}
 */
function sortPassbookPages_(pages) {
  if (pages.length <= 1) {
    return pages;
  }
  
  // ページ番号が全て取得できているかチェック
  const allHavePageNumber = pages.every(p => p.pageNumber !== null && p.pageNumber !== undefined);
  
  if (allHavePageNumber) {
    // ページ番号でソート
    console.log('ページ番号でソート');
    return pages.sort((a, b) => a.pageNumber - b.pageNumber);
  }
  
  // 残高の連続性でソート
  console.log('残高の連続性でソート');
  return sortByBalanceContinuity_(pages);
}

/**
 * 残高の連続性でソート
 * 前のページの最後の残高 = 次のページの最初の残高 となるように並べる
 * @param {Array} pages
 * @return {Array}
 */
function sortByBalanceContinuity_(pages) {
  if (pages.length <= 1) {
    return pages;
  }
  
  const sorted = [];
  const remaining = [...pages];
  
  // 最初のページを見つける（最初の残高が最も古い＝金額が小さいか、他のページの最後の残高と一致しない）
  // または、最初の取引日が最も古いページを選ぶ
  let firstPage = null;
  let firstPageIndex = -1;
  
  // 各ページの最初の残高が、他のページの最後の残高と一致するかチェック
  for (let i = 0; i < remaining.length; i++) {
    const page = remaining[i];
    const firstBal = page.firstBalance;
    
    // このページの最初の残高が、他のページの最後の残高と一致するか
    const isPrecededByOther = remaining.some((other, j) => {
      if (i === j) return false;
      return other.lastBalance === firstBal;
    });
    
    // 一致しない = これが最初のページの可能性が高い
    if (!isPrecededByOther) {
      if (firstPage === null) {
        firstPage = page;
        firstPageIndex = i;
      } else {
        // 複数候補がある場合は、最初の取引日が古い方を選ぶ
        const pageFirstDate = page.transactions[0]?.date || '';
        const currentFirstDate = firstPage.transactions[0]?.date || '';
        if (pageFirstDate < currentFirstDate) {
          firstPage = page;
          firstPageIndex = i;
        }
      }
    }
  }
  
  // 最初のページが見つからない場合は、最初の取引日が最も古いページを選ぶ
  if (firstPage === null) {
    remaining.sort((a, b) => {
      const dateA = a.transactions[0]?.date || '';
      const dateB = b.transactions[0]?.date || '';
      return dateA.localeCompare(dateB);
    });
    firstPage = remaining[0];
    firstPageIndex = 0;
  }
  
  // 最初のページを追加
  sorted.push(firstPage);
  remaining.splice(firstPageIndex, 1);
  
  // 残りのページを連結
  while (remaining.length > 0) {
    const lastPage = sorted[sorted.length - 1];
    const lastBalance = lastPage.lastBalance;
    
    // 次のページを探す（最初の残高が lastBalance と一致するもの）
    let nextPageIndex = remaining.findIndex(p => p.firstBalance === lastBalance);
    
    if (nextPageIndex === -1) {
      // 一致するものがない場合、最も近い残高を持つページを選ぶ
      let minDiff = Infinity;
      for (let i = 0; i < remaining.length; i++) {
        const diff = Math.abs(remaining[i].firstBalance - lastBalance);
        if (diff < minDiff) {
          minDiff = diff;
          nextPageIndex = i;
        }
      }
      console.warn('残高が一致しないページがあります。最も近い残高で連結: ' + 
        lastBalance + ' → ' + remaining[nextPageIndex].firstBalance);
    }
    
    sorted.push(remaining[nextPageIndex]);
    remaining.splice(nextPageIndex, 1);
  }
  
  return sorted;
}

// ============================================================
// Gemini API
// ============================================================

/**
 * 通帳画像からデータを抽出
 * @param {GoogleAppsScript.Drive.File} file
 * @param {string} apiKey
 * @return {Object} 通帳データ
 */
function extractPassbookData_(file, apiKey) {
  const blob = file.getBlob();
  const base64 = Utilities.base64Encode(blob.getBytes());
  const mimeType = blob.getContentType();
  
  return callGeminiForPassbook_(base64, mimeType, apiKey);
}

/**
 * Gemini APIを呼び出して通帳を読み取り
 * @param {string} base64Content
 * @param {string} mimeType
 * @param {string} apiKey
 * @return {Object}
 */
function callGeminiForPassbook_(base64Content, mimeType, apiKey) {
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' +
              PASSBOOK_CONFIG.GEMINI_MODEL + ':generateContent?key=' + apiKey;
  
  const prompt = buildPassbookOCRPrompt_();
  
  const payload = {
    contents: [{
      parts: [
        { text: prompt },
        {
          inline_data: {
            mime_type: mimeType,
            data: base64Content
          }
        }
      ]
    }],
    generationConfig: {
      temperature: PASSBOOK_CONFIG.GEMINI_TEMPERATURE,
      maxOutputTokens: PASSBOOK_CONFIG.GEMINI_MAX_TOKENS,
      responseMimeType: 'application/json'
    }
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  const statusCode = response.getResponseCode();
  
  if (statusCode !== 200) {
    throw new Error('Gemini API Error: ' + statusCode);
  }
  
  const result = JSON.parse(response.getContentText());
  return parsePassbookResponse_(result);
}

/**
 * 通帳OCR用のプロンプト
 * @return {string}
 */
function buildPassbookOCRPrompt_() {
  return `あなたは日本の銀行通帳のOCR専門家です。
通帳の画像から取引データを正確に抽出してください。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【重要】ページ番号の抽出
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

通帳の右上または左上にページ番号が記載されていることがあります。
例: 「4」「5」「P.4」「4/10」など

ページ番号が見つかった場合は必ず抽出してください。
見つからない場合はnullを返してください。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【最重要】銀行ごとのフォーマット差異への対応
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

日本の銀行通帳は銀行によってレイアウトが大きく異なります。
以下のパターンを理解し、正しく金額と摘要を分離してください。

■ パターンA: 標準形式（多くの銀行）
列構成: [年月日] [摘要] [お支払金額] [お預り金額] [残高]
例:
  01-05 | シヤカイホケン | 60,980 |        | 1,120,402
  01-05 | 振込 エクシード |        | 737,115 | 1,512,178

■ パターンB: 金額欄に摘要が結合（一部の地方銀行・信用金庫）
列構成: [年月日] [記号] [お支払金額(+摘要)] [お預り金額(+摘要)] [残高]
例:
  08-01-05 | 200 | *60,980シヤカイホケンリヨウ |           | *1,120,402
  08-01-05 | 振込カ)エクシート゛ |              | *737,115 | *1,512,178

このパターンでは：
- 出金時: 「記号」欄に"200"等のコード、「お支払金額」欄に"金額+摘要"が結合
- 入金時: 「記号」欄に"摘要"、「お預り金額」欄に金額のみ
→ 金額と摘要を正しく分離して抽出すること！

■ パターンC: 摘要が複数列に分かれる
列構成: [年月日] [取引種別] [摘要1] [摘要2] [出金] [入金] [残高]
→ 摘要1と摘要2を結合して1つの摘要として出力

■ パターンD: ネット銀行・Web通帳形式
列構成や表示が銀行独自。項目名を参考に判断。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【金額と摘要の分離ルール】
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

金額欄に "*60,980シヤカイホケンリヨウ" のように
金額と摘要が結合されている場合：

1. 先頭の * や記号を除去
2. 数字部分（カンマ含む）を金額として抽出
3. 数字以降の文字列を摘要として抽出

例: "*60,980シヤカイホケンリヨウ"
→ withdrawal: 60980
→ description: "シヤカイホケンリヨウ"

例: "*237,489セイサクコウコ(コクミン"
→ withdrawal: 237489
→ description: "セイサクコウコ(コクミン"

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【日付の変換ルール】
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

- 和暦は西暦に変換: 令和N年 = (2018+N)年
- "08-01-05" のような形式は "2008-01-05" ではなく、
  文脈から判断（通帳の発行時期、残高の推移等）
  → 2008年か2026年かを適切に判断
- 年が省略されている場合（"01-05"等）は、
  通帳上部や前後の取引から年を推定
- 出力形式: YYYY-MM-DD

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【記号・コードの解釈】
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

通帳に記載される記号の一般的な意味：
- "200" = 振込（出金）
- "振込" "振込カ)" = 振込入金（カ=カブシキガイシャ等の略）
- "ATM" = ATM取引
- "利息" = 利息入金
- "手数料" = 各種手数料

記号自体は description に含めず、
実際の取引内容（振込先名等）を description に抽出すること。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【抽出項目】
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1. pageNumber: ページ番号（通帳右上等に記載、なければnull）

各取引行について以下を抽出:
2. date: 取引日（YYYY-MM-DD形式）
3. description: 摘要（取引先名、取引内容など。記号コードではなく実際の内容）
4. withdrawal: お支払金額（出金額、なければnull）
5. deposit: お預り金額（入金額、なければnull）
6. balance: 差引残高

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【出力形式】
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

{
  "bankName": "銀行名（ヘッダーや通帳デザインから推定、不明なら空文字）",
  "accountType": "普通預金",
  "pageNumber": 4,
  "dateRange": "2026/01/05〜2026/01/27",
  "transactions": [
    {
      "date": "2026-01-05",
      "description": "シヤカイホケンリヨウ",
      "withdrawal": 60980,
      "deposit": null,
      "balance": 1120402
    },
    {
      "date": "2026-01-05",
      "description": "セイサクコウコ(コクミン",
      "withdrawal": 237489,
      "deposit": null,
      "balance": 882913
    },
    {
      "date": "2026-01-05",
      "description": "振込 エクシード",
      "withdrawal": null,
      "deposit": 737115,
      "balance": 1512178
    }
  ]
}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【重要な注意事項】
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1. ページ番号は必ず抽出を試みる（右上、左上などを確認）
2. 金額は必ず数値で出力（カンマなし）
3. 残高は必ず数値で出力
4. 空行や読み取れない行はスキップ
5. 手書きの斜線や取り消し線がある行は除外
6. 取引は上から下の順番で出力
7. 日付が同じ行が複数あっても、すべて個別の取引として抽出
8. 金額の前の * は除去（一部の銀行で使用される記号）
9. "200" などの記号コードは description に含めない
10. 入金の場合、「振込カ)エクシード」のような摘要は
    description: "振込 エクシード" のように出力（カ)は除去可）`;
}

/**
 * Geminiレスポンスをパース
 * @param {Object} apiResult
 * @return {Object}
 */
function parsePassbookResponse_(apiResult) {
  try {
    const text = apiResult.candidates?.[0]?.content?.parts?.[0]?.text || '';
    
    // JSONを抽出
    let jsonStr = text.trim();
    if (jsonStr.startsWith('```json')) {
      jsonStr = jsonStr.replace(/^```json\s*/, '').replace(/\s*```$/, '');
    } else if (jsonStr.startsWith('```')) {
      jsonStr = jsonStr.replace(/^```\s*/, '').replace(/\s*```$/, '');
    }
    
    const parsed = JSON.parse(jsonStr);
    
    // ページ番号を数値に変換
    let pageNumber = null;
    if (parsed.pageNumber !== null && parsed.pageNumber !== undefined) {
      const pn = parseInt(String(parsed.pageNumber).replace(/[^0-9]/g, ''), 10);
      if (!isNaN(pn)) {
        pageNumber = pn;
      }
    }
    
    // 日付の正規化
    let transactions = [];
    if (parsed.transactions) {
      transactions = parsed.transactions.map(function(tx) {
        return {
          date: normalizePassbookDate_(tx.date),
          description: String(tx.description || '').trim(),
          withdrawal: parsePassbookAmount_(tx.withdrawal),
          deposit: parsePassbookAmount_(tx.deposit),
          balance: parsePassbookAmount_(tx.balance)
        };
      });
    }
    
    return {
      bankName: parsed.bankName || '',
      accountType: parsed.accountType || '',
      pageNumber: pageNumber,
      dateRange: parsed.dateRange || '',
      transactions: transactions
    };
    
  } catch (e) {
    console.error('Passbook Response Parse Error: ' + e.message);
    return { transactions: [], pageNumber: null };
  }
}

// ============================================================
// シート操作
// ============================================================

/**
 * 通帳シートを取得または作成
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreatePassbookSheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(PASSBOOK_CONFIG.SHEET_NAME);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(PASSBOOK_CONFIG.SHEET_NAME);
    
    // ヘッダーを設定
    const headers = [
      '取引日',       // A
      '摘要',         // B
      '入金',         // C
      '出金',         // D
      '残高',         // E
      '勘定科目',     // F
      '補助科目',     // G
      '画像リンク',   // H
      'ステータス'    // I
    ];
    
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    
    // 列幅調整
    sheet.setColumnWidth(1, 100);  // 取引日
    sheet.setColumnWidth(2, 200);  // 摘要
    sheet.setColumnWidth(3, 100);  // 入金
    sheet.setColumnWidth(4, 100);  // 出金
    sheet.setColumnWidth(5, 120);  // 残高
    sheet.setColumnWidth(6, 120);  // 勘定科目
    sheet.setColumnWidth(7, 120);  // 補助科目
    sheet.setColumnWidth(8, 100);  // 画像リンク
    sheet.setColumnWidth(9, 80);   // ステータス
  }
  
  return sheet;
}

/**
 * 通帳取引をシートに追加
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object} data
 */
function appendPassbookTransaction_(sheet, data) {
  const row = [
    data.date,
    data.description,
    data.deposit || '',
    data.withdrawal || '',
    data.balance || '',
    data.accountTitle || '',
    data.subAccount || '',
    data.fileUrl ? '=HYPERLINK("' + data.fileUrl + '", "画像")' : '',
    data.status || '未確認'
  ];
  
  sheet.appendRow(row);
}

// ============================================================
// ヘルパー関数
// ============================================================

/**
 * 処理済み通帳ファイルか判定
 * @param {string} fileName
 * @return {boolean}
 */
function isPassbookProcessed_(fileName) {
  return /^\[(OK|ERR)\]/.test(fileName) || /^[🟢🔴]/.test(fileName);
}

/**
 * 対応MIMEタイプか判定
 * @param {string} mime
 * @return {boolean}
 */
function isSupportedPassbookMime_(mime) {
  return mime === 'application/pdf' || mime.startsWith('image/');
}

/**
 * ファイルに処理済みマークを付与
 * @param {GoogleAppsScript.Drive.File} file
 */
function markPassbookAsProcessed_(file) {
  const currentName = file.getName();
  if (!isPassbookProcessed_(currentName)) {
    file.setName('[OK]' + currentName);
  }
}

/**
 * ファイルにエラーマークを付与
 * @param {GoogleAppsScript.Drive.File} file
 */
function markPassbookAsError_(file) {
  const currentName = file.getName();
  if (!isPassbookProcessed_(currentName)) {
    file.setName('[ERR]' + currentName);
  }
}

/**
 * 日付を正規化
 * @param {string} dateStr
 * @return {string}
 */
function normalizePassbookDate_(dateStr) {
  if (!dateStr) return '';
  let str = String(dateStr).trim();
  
  // 全角数字を半角に
  str = str.replace(/[０-９]/g, function(s) {
    return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
  });
  
  // YYYY-MM-DD形式ならそのまま返す
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
    return str;
  }
  
  // MM-DD形式（年なし）の場合は現在の年を付加
  const mdMatch = str.match(/^(\d{1,2})[-\/](\d{1,2})$/);
  if (mdMatch) {
    const year = new Date().getFullYear();
    const month = ('0' + mdMatch[1]).slice(-2);
    const day = ('0' + mdMatch[2]).slice(-2);
    return year + '-' + month + '-' + day;
  }
  
  // その他のパターン
  const ymdMatch = str.match(/(\d{2,4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (ymdMatch) {
    let year = parseInt(ymdMatch[1]);
    if (year < 100) year += 2000;
    const month = ('0' + ymdMatch[2]).slice(-2);
    const day = ('0' + ymdMatch[3]).slice(-2);
    return year + '-' + month + '-' + day;
  }
  
  return str;
}

/**
 * 金額をパース
 * @param {*} value
 * @return {number|null}
 */
function parsePassbookAmount_(value) {
  if (value === null || value === undefined || value === '') {
    return null;
  }
  
  // 文字列の場合、カンマを除去
  if (typeof value === 'string') {
    value = value.replace(/,/g, '').trim();
  }
  
  const num = parseInt(value, 10);
  return isNaN(num) ? null : num;
}

// ============================================================
// 公開関数（ライブラリとして使用時）
// ============================================================

/**
 * 通帳シートのみを作成（処理は行わない）
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 */
function createPassbookSheet(spreadsheet) {
  getOrCreatePassbookSheet_(spreadsheet);
}

/**
 * バージョン情報
 * @return {string}
 */
function getVersion() {
  return '1.1.0';
}
