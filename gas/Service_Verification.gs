/**
 * Service_Verification.gs
 * AI検証レイヤー（プロトタイプ版）
 *
 * 責務:
 * - 既存の読み取り結果をGPT-5で再検証
 * - レシート画像と読み取りデータを比較し、誤りを検出
 * - 検証結果を17-20列目に書き込み
 *
 * 使用AI:
 * - 検証: GPT-5 (OpenAI, 2025年8月7日リリース)
 */

// ============================================================
// メイン関数
// ============================================================

/**
 * 選択行のレシートをAIで検証する
 */
function verifySelectedRows() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (sheet.getName() !== CONFIG.SHEET_NAME.MAIN) {
    ui.alert('「' + CONFIG.SHEET_NAME.MAIN + '」シートで実行してください。');
    return;
  }

  const selection = sheet.getActiveRange();
  if (!selection) {
    ui.alert('検証したい行を選択してください。');
    return;
  }

  const startRow = selection.getRow();
  const numRows = selection.getNumRows();

  if (startRow <= 1) {
    ui.alert('ヘッダー行は検証できません。データ行を選択してください。');
    return;
  }

  // 検証用列の確保
  ensureVerificationColumns_(sheet);

  ss.toast('検証を開始しました...', '検証中', -1);

  let processedCount = 0;
  let errorCount = 0;

  for (let row = startRow; row < startRow + numRows; row++) {
    try {
      ss.toast('行 ' + row + ' を検証中... (' + (row - startRow + 1) + '/' + numRows + ')', '検証中', -1);
      verifyOneRow_(sheet, row, null);
      processedCount++;
    } catch (e) {
      console.error('検証エラー (行' + row + '): ' + e.message);
      writeVerificationError_(sheet, row, e.message);
      errorCount++;
    }
  }

  const msg = processedCount + '行の検証が完了しました。' +
              (errorCount > 0 ? '（エラー: ' + errorCount + '件）' : '');
  ss.toast(msg, '検証完了', 5);
}

// ============================================================
// 検証処理
// ============================================================

/**
 * 計算の整合性をチェック（Gemini呼び出し前）
 * @param {Object} rowData - 行データ
 * @return {Array} issues - 検出された問題の配列
 */
function checkCalculations_(rowData) {
  const issues = [];

  // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  // チェック1: 消費税(10%)の計算
  // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  if (rowData.taxable10 > 0) {
    const expectedTax10 = Math.round(rowData.taxable10 * 0.1);
    const taxDiff10 = Math.abs(expectedTax10 - rowData.tax10);

    if (taxDiff10 > 1) {
      issues.push({
        category: 'tax',
        severity: 'high',
        field: 'tax10',
        currentValue: rowData.tax10,
        correctValue: expectedTax10,
        reason: `消費税(10%)が${taxDiff10}円ズレています。${rowData.taxable10}円 × 0.1 = ${expectedTax10}円のはずです`,
        confidence: 1.0,
        evidence: '計算結果'
      });
    }
  }

  // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  // チェック2: 消費税(8%)の計算
  // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  if (rowData.taxable8 > 0) {
    const expectedTax8 = Math.round(rowData.taxable8 * 0.08);
    const taxDiff8 = Math.abs(expectedTax8 - rowData.tax8);

    if (taxDiff8 > 1) {
      issues.push({
        category: 'tax',
        severity: 'high',
        field: 'tax8',
        currentValue: rowData.tax8,
        correctValue: expectedTax8,
        reason: `消費税(8%)が${taxDiff8}円ズレています。${rowData.taxable8}円 × 0.08 = ${expectedTax8}円のはずです`,
        confidence: 1.0,
        evidence: '計算結果'
      });
    }
  }

  // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  // チェック3: 不自然な桁数（誤読の可能性）
  // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  if (rowData.totalAmount >= 100000) {
    issues.push({
      category: 'amount',
      severity: 'medium',
      field: 'totalAmount',
      currentValue: rowData.totalAmount,
      correctValue: null,
      reason: `金額が${rowData.totalAmount.toLocaleString()}円と高額です。手書きの場合、桁数を誤読している可能性があります（例: ¥2,200を92,200と誤読）`,
      confidence: 0.7,
      evidence: '金額の範囲チェック'
    });
  }

  // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  // チェック5: ゼロ円チェック
  // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  if (rowData.totalAmount === 0) {
    issues.push({
      category: 'amount',
      severity: 'high',
      field: 'totalAmount',
      currentValue: 0,
      correctValue: null,
      reason: '総額が0円です。読み取りに失敗している可能性があります',
      confidence: 1.0,
      evidence: '金額チェック'
    });
  }

  return issues;
}

/**
 * 行データを検証用に取得
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} row
 * @return {Object} rowData
 */
function getRowDataForVerification_(sheet, row) {
  const values = sheet.getRange(row, 1, 1, 16).getValues()[0];

  return {
    rowIndex: row,
    date: formatCellValueForVerification_(values[3]),
    storeName: String(values[4] || ''),
    registrationNumber: String(values[5] || ''),
    totalAmount: Number(values[6]) || 0,
    taxable10: Number(values[7]) || 0,
    tax10: Number(values[8]) || 0,
    taxable8: Number(values[9]) || 0,
    tax8: Number(values[10]) || 0,
    nonTaxable: Number(values[11]) || 0,
    account: String(values[12] || '')
  };
}

/**
 * 1行の検証処理
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} row
 * @param {string} apiKey - 未使用（互換性のため残す）
 */
function verifyOneRow_(sheet, row, apiKey) {
  try {
    // ステータスを確認（手書きレシートはスキップ）
    const status = sheet.getRange(row, 2).getValue();

    if (status === 'HAND') {
      Logger.log('Row ' + row + ': 手書き領収証のため検証スキップ');

      // 検証結果欄に説明を書く
      sheet.getRange(row, 19).setValue('🖊️ 手書き領収証（目視確認してください）');
      sheet.getRange(row, 20).setValue('');

      return; // 検証処理を行わずに終了
    }

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // ステップ1: 画像ファイルを取得
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    const imageFile = getImageFileForRow_(sheet, row);
    if (!imageFile) {
      writeVerificationError_(sheet, row, '画像ファイルが見つかりません');
      return;
    }

    // Base64に変換
    const base64Image = convertToBase64_(imageFile);

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // ステップ2: 行データを構造化
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    const values = sheet.getRange(row, 1, 1, 16).getValues()[0];
    const rowData = {
      date: formatCellValueForVerification_(values[3]),
      storeName: String(values[4] || ''),
      registrationNumber: String(values[5] || ''),
      totalAmount: formatCellValueForVerification_(values[6]),
      taxable10: formatCellValueForVerification_(values[7]),
      tax10: formatCellValueForVerification_(values[8]),
      taxable8: formatCellValueForVerification_(values[9]),
      tax8: formatCellValueForVerification_(values[10]),
      nonTaxable: formatCellValueForVerification_(values[11]),
      account: String(values[12] || '')
    };

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // ステップ3: プロンプトを構築
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    const prompt = buildVerificationPrompt_(rowData, []);

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // ステップ4: GPT-5 APIで検証
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    Logger.log('Row ' + row + ': GPT-5検証開始');
    const response = callGPT5ForVerification_(imageFile, base64Image, prompt);
    const responseText = extractGPT5Text_(response);
    const result = parseVerificationResponse_(responseText);

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // ステップ5: 事後修正（内税/外税判定ミス）
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    const fixedResult = fixTaxCalculationError_(result, rowData.totalAmount);

    // 結果を書き込み
    writeVerificationResult_(sheet, row, fixedResult);
    Logger.log('Row ' + row + ': 検証完了');

  } catch (error) {
    Logger.log('Row ' + row + ': エラー - ' + error.toString());
    writeVerificationError_(sheet, row, error.toString());
  }
}

/**
 * セル値を検証プロンプト用の文字列に変換
 * @param {*} val
 * @return {string}
 */
function formatCellValueForVerification_(val) {
  if (val instanceof Date) {
    var y = val.getFullYear();
    var m = ('0' + (val.getMonth() + 1)).slice(-2);
    var d = ('0' + val.getDate()).slice(-2);
    return y + '-' + m + '-' + d;
  }
  if (val === '' || val === null || val === undefined) return '';
  return String(val);
}

// ============================================================
// Gemini API呼び出し
// ============================================================

/**
 * 検証用プロンプトを構築（ブラインド検証方式 - 強化版）
 * AIが先に独自読み取りを行い、その後で既存結果と比較する
 * @param {Object} rowData
 * @param {Array} calcIssues - 計算エラーの配列（オプション）
 * @return {string}
 */
function buildVerificationPrompt_(rowData, calcIssues) {
  calcIssues = calcIssues || [];

  var prompt = '🚨🚨🚨 最重要指示 🚨🚨🚨\n' +
'\n' +
'このタスクは2つのステップに分かれていますが、各ステップは完全に独立しています。\n' +
'\n' +
'【禁止事項】\n' +
'- ステップ1の実行中に、ステップ2の「既存の読み取り結果」を参照すること\n' +
'- 既存結果の数値を yourReading にコピーすること\n' +
'- 「既存と一致している」という理由で、画像を確認せずに値を記入すること\n' +
'\n' +
'【必須事項】\n' +
'- yourReading には、あなたが画像から直接読み取った値"のみ"を記入\n' +
'- 既存結果とあなたの読み取りが一致していても、必ず画像を見て確認\n' +
'- 不明な場合は null にする（既存結果からコピーしない）\n' +
'\n' +
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
'ステップ1: 画像のみを見て、あなた自身が読み取る\n' +
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
'\n' +
'⚠️ この段階では、下に書いてある「既存の読み取り結果」を絶対に参照しないでください。\n' +
'\n' +
'まるでこのレシートを初めて見るかのように、画像だけを観察してください。\n' +
'\n' +
'【読み取る項目】\n' +
'\n' +
'1. **発行者（店名・会社名・個人名）**\n' +
'   \n' +
'   確認方法：\n' +
'   - 画像のどこに店名が書いてありますか？\n' +
'   - 印鑑・ハンコは誰の名前ですか？\n' +
'   - 「上記正に領収いたしました」の主語は誰ですか？\n' +
'   - 画像下部に住所・電話番号と一緒に記載されている名前は？\n' +
'   \n' +
'   ⚠️ 注意：「宛名（〇〇様）」ではなく「発行者」を探す\n' +
'   \n' +
'   ⚠️ 重要：店舗名の扱い\n' +
'   \n' +
'   店舗名は「ブランド名のみ」で十分です。\n' +
'   支店名、店舗番号、法人格（株式会社など）は不要です。\n' +
'   \n' +
'   【正しい例】\n' +
'   ✅ "LAWSON"（支店名不要）\n' +
'   ✅ "Amazon"（.co.jp不要）\n' +
'   ✅ "Starbucks"（Coffee、渋谷店など不要）\n' +
'   ✅ "セブンイレブン"（◯◯店不要）\n' +
'   \n' +
'   【間違った例】\n' +
'   ❌ "LAWSON 門真月出町店"（支店名は不要）\n' +
'   ❌ "Amazon.co.jp"（法人格不要）\n' +
'   ❌ "株式会社○○"（法人格不要）\n' +
'   \n' +
'   例外：レシート上にブランド名がなく、個人名や\n' +
'   固有の店舗名しかない場合は、その名前を使用。\n' +
'   \n' +
'   comparison での店舗名の比較：\n' +
'   - ブランド名が一致していれば match: true\n' +
'   - 支店名の有無は無視してください\n' +
'   - 例: "LAWSON" vs "LAWSON 門真店" → match: true とすべき\n' +
'   - 支店名の違いで issue を作らないでください\n' +
'   \n' +
'   yourReading.storeName に記入する値：\n' +
'   → あなたが画像で見た店名をそのまま書く\n' +
'   → 既存結果とは無関係に、画像だけを見て判断\n' +
'\n' +
'2. **日付**\n' +
'   \n' +
'   確認方法：\n' +
'   - 画像のどこに日付が書いてありますか？\n' +
'   - 和暦（R7年など）ですか？西暦ですか？\n' +
'   - R7年 = 令和7年 = 2025年\n' +
'   \n' +
'   yourReading.date に記入する値：\n' +
'   → 画像に書いてある日付を西暦YYYY-MM-DD形式で\n' +
'\n' +
'3. **総合計**\n' +
'   \n' +
'   ⚠️ 重要：¥記号の識別方法\n' +
'   \n' +
'   ¥記号には必ず横2本線（=）が入っています。\n' +
'   たとえ「Y」の部分が数字の9や7に似ていても、\n' +
'   横2本線があれば、それは通貨記号であり数字ではありません。\n' +
'   \n' +
'   【正しい読み方】\n' +
'   ✅ ¥2,200 → 2,200円（¥記号の横線を確認）\n' +
'   ❌ ¥92,200 → 間違い（¥を9と誤認）\n' +
'   ❌ ¥72,200 → 間違い（¥を7と誤認）\n' +
'   \n' +
'   【識別手順】\n' +
'   1. ★や「合計」の後にある記号を確認\n' +
'   2. 横2本線（=）があれば、それは¥記号\n' +
'   3. ¥記号の"直後"から数字を読み始める\n' +
'   4. 桁数が異常に多い場合（6桁以上）は¥記号の誤認を疑う\n' +
'   \n' +
'   確認方法：\n' +
'   - 画像のどこに金額が書いてありますか？\n' +
'   - ★や「合計」などのマークがついていますか？\n' +
'   - ¥記号（横2本線）の直後の数字はいくつですか？\n' +
'   - 手書きの場合、￥記号と数字を区別できていますか？\n' +
'   \n' +
'   yourReading.totalAmount に記入する値：\n' +
'   → ¥記号の直後から読み取った金額（数値のみ）\n' +
'   → 桁数が多すぎる場合は再確認\n' +
'\n' +
'4. **税区分別の内訳**\n' +
'   \n' +
'   確認方法：\n' +
'   - 「外税10%」「税込」「税抜」などの表記を探す\n' +
'   - 10%対象額と消費税額を確認\n' +
'   - 8%対象額と消費税額を確認\n' +
'   \n' +
'   ⚠️ 重要：外税表記の解釈\n' +
'   - 「(外8% 対象 ¥398)」→ これは税抜398円です\n' +
'   - 「(外税8% ¥31)」→ これは消費税31円です\n' +
'   - 「(外10% 対象 ¥5)」→ これは税抜5円です\n' +
'   - 「外10% 対象」と「外税10%」は別物\n' +
'   \n' +
'   yourReading に記入する値：\n' +
'   - taxable10: 画像で「10%対象」と書いてある金額（税抜）\n' +
'   - tax10: 画像で「外税10%」または「消費税10%」と書いてある金額\n' +
'   - taxable8: 画像で「8%対象」と書いてある金額（税抜）\n' +
'   - tax8: 画像で「外税8%」または「消費税8%」と書いてある金額\n' +
'   - nonTaxable: 入湯税、宿泊税など\n' +
'\n' +
'   ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
'   ⚠️ 重要：税表記には2種類あります\n' +
'   ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
'\n' +
'   レシートの税表記には「外税表記」と「内税表記」があります。\n' +
'   必ず判定してから yourReading を記入してください。\n' +
'\n' +
'   【パターンA: 外税表記】\n' +
'   例：\n' +
'     小計: ¥1,733\n' +
'     (外8% 対象 ¥1,730)\n' +
'     外8% ¥138\n' +
'     (外10%対象 ¥3)\n' +
'     合計 ¥1,871\n' +
'\n' +
'   意味：\n' +
'   - 「対象額」は税抜金額\n' +
'   - 「外税」は別途加算される消費税\n' +
'   - 合計 = 対象額 + 消費税\n' +
'\n' +
'   yourReading記入例：\n' +
'     taxable8: 1730（税抜）\n' +
'     tax8: 138（消費税）\n' +
'     taxable10: 3（税抜）\n' +
'\n' +
'   検算：1730 + 138 + 3 = 1871 ✓\n' +
'\n' +
'\n' +
'   【パターンB: 内税表記】\n' +
'   例：\n' +
'     合計 ¥510\n' +
'     (10%対象 ¥3)\n' +
'     (内消費税額 ¥0)\n' +
'     (8%対象 ¥507)\n' +
'     (内消費税額 ¥37)\n' +
'\n' +
'   意味：\n' +
'   - 「対象額」は税込金額\n' +
'   - 「内消費税額」は対象額に含まれる税額\n' +
'   - 合計 = 対象額の合計（消費税は別途加算しない）\n' +
'\n' +
'   yourReading記入例：\n' +
'     taxable8: 470（税抜 = 507 - 37）\n' +
'     tax8: 37（消費税）\n' +
'     taxable10: 3（税抜 = 3 - 0）\n' +
'     tax10: 0（消費税）\n' +
'\n' +
'   検算：470 + 37 + 3 + 0 = 510 ✓\n' +
'\n' +
'\n' +
'   【判定方法】\n' +
'\n' +
'   Step 1: レシートに「外税」「外○%」という表記があるか？\n' +
'     → ある場合：外税表記\n' +
'\n' +
'   Step 2: レシートに「内消費税」「内税」という表記があるか？\n' +
'     → ある場合：内税表記\n' +
'\n' +
'   Step 3: 対象額の合計を計算\n' +
'     例：(10%対象 ¥3) + (8%対象 ¥507) = 510円\n' +
'\n' +
'     合計金額と一致する？\n' +
'     → 一致：内税表記（対象額は税込）\n' +
'     → 不一致：外税表記（対象額は税抜）\n' +
'\n' +
'   Step 4: yourReadingに記入する値\n' +
'     - 内税表記の場合：\n' +
'       taxableN = 対象額 - 内消費税額\n' +
'       taxN = 内消費税額\n' +
'\n' +
'     - 外税表記の場合：\n' +
'       taxableN = 対象額\n' +
'       taxN = 外税額\n' +
'\n' +
'5. **登録番号**\n' +
'   \n' +
'   確認方法：\n' +
'   - 「T」で始まる13桁の番号はありますか？\n' +
'   \n' +
'   yourReading.registrationNumber に記入する値：\n' +
'   → T+13桁、または null\n' +
'\n' +
'【あなたの読み取り結果を記録】\n' +
'\n' +
'yourReading: {\n' +
'  storeName: "画像で見た店名",\n' +
'  storeNameEvidence: "画像のどこに書いてあったか（例：中央下部の印鑑）",\n' +
'  date: "YYYY-MM-DD",\n' +
'  totalAmount: 数値,\n' +
'  taxable10: 数値,\n' +
'  tax10: 数値,\n' +
'  taxable8: 数値,\n' +
'  tax8: 数値,\n' +
'  nonTaxable: 数値,\n' +
'  registrationNumber: "T+13桁 または null"\n' +
'}\n' +
'\n' +
'⚠️ 再確認：上記の値は全て"画像から"読み取ったものですか？\n' +
'既存結果からコピーしていませんか？\n' +
'\n' +
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
'ステップ2: 既存の読み取り結果と比較する\n' +
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
'\n' +
'ここで初めて、既存の読み取り結果を見てください。\n' +
'\n' +
'【既存の読み取り結果】\n' +
'- 日付: ' + (rowData.date || 'なし') + '\n' +
'- 店名: ' + (rowData.storeName || 'なし') + '\n' +
'- 登録番号: ' + (rowData.registrationNumber || 'なし') + '\n' +
'- 総合計: ' + (rowData.totalAmount || 0) + '円\n' +
'- 対象額(10%): ' + (rowData.taxable10 || 0) + '円\n' +
'- 消費税(10%): ' + (rowData.tax10 || 0) + '円\n' +
'- 対象額(8%): ' + (rowData.taxable8 || 0) + '円\n' +
'- 消費税(8%): ' + (rowData.tax8 || 0) + '円\n' +
'- 不課税: ' + (rowData.nonTaxable || 0) + '円\n' +
'- 勘定科目: ' + (rowData.account || 'なし') + '\n' +
'\n' +
'【重要】既存値が空・0・なしの場合の扱い\n' +
'\n' +
'既存値が「空」「0」「なし」「null」「UNKNOWN」「PARSE_ERROR」で、\n' +
'あなたの読み取り値が有効な値（数値 > 0、または文字列）の場合：\n' +
'\n' +
'- これは「データ欠落」であり、必ず issues に含めること\n' +
'- severity: high として報告すること\n' +
'- comparison の match は false とすること\n' +
'\n' +
'例：\n' +
'- 既存の taxable8 = 0、あなたの読み取り = 696 → issue（severity: high）\n' +
'- 既存の tax8 = 0、あなたの読み取り = 55 → issue（severity: high）\n' +
'- 既存の registrationNumber = なし、あなたの読み取り = T123... → issue（severity: high）\n' +
'- 既存の storeName = PARSE_ERROR、あなたの読み取り = オーエスドラッグ → issue（severity: high）\n' +
'\n' +
'⚠️ 全てのフィールドについて、既存値と自分の読み取りを比較し、\n' +
'差異があれば漏れなく全て issues に含めてください。\n' +
'1回の検証で全ての問題を検出することが重要です。\n' +
'\n' +
'【比較してください】\n' +
'\n' +
'あなたが「ステップ1で画像から読み取った値」と、上記の「既存結果」を比較してください。\n' +
'\n' +
'各項目について：\n' +
'- match: true/false（一致しているか）\n' +
'- original: 既存の値\n' +
'- yours: あなたがステップ1で読み取った値\n' +
'- correct: どちらが正しいか\n' +
'- reason: なぜそう判断したか\n' +
'\n' +
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
'ステップ3: 差異の判定と修正提案\n' +
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
'\n' +
'差異がある項目について：\n' +
'1. どちらが正しいか判定\n' +
'2. 理由を明確に説明\n' +
'3. 画像のどこに証拠があるか示す\n' +
'\n' +
'【よくある誤読パターン】\n' +
'- 「宛名（お客様名）」を「店名（発行者）」と誤認\n' +
'- 手書きの「￥」を数字の「7」「2」と誤読\n' +
'- 手書きの「✓」を数字の「1」と誤読\n' +
'- 和暦の年号計算ミス（R7年を2027年と誤認）\n' +
'- 桁数の間違い（3円を30円、50円を500円、2,200円を92,200円）\n' +
'- 「外税」表記の誤解釈（税抜と消費税の取り違え）\n' +
'\n' +
'【端数値引き・値引きの処理ルール】\n' +
'\n' +
'レシートに「端数値引」「値引」「割引」「クーポン」などがある場合の注意点：\n' +
'\n' +
'1. 合計金額（totalAmount）を絶対正とする\n' +
'2. 税抜額 + 消費税 + 不課税 = 合計 が成立していれば正常\n' +
'3. 税抜額がレシート記載の「課税対象額」より数円〜数十円少ない場合がある\n' +
'   - これは値引き分を税抜額から差し引いているため\n' +
'   - 例：課税対象 ¥7,140 + 税 ¥714 - 値引 ¥4 = 合計 ¥7,850\n' +
'   - この場合、税抜額は 7,140 ではなく 7,136 が正しい\n' +
'\n' +
'4. 以下の場合はissueとして報告しない：\n' +
'   - taxable10/taxable8 がレシート記載値より少ないが、\n' +
'     合計金額が完全一致している場合\n' +
'   - 差額が「端数値引」「値引」等の金額と一致または近い場合\n' +
'\n' +
'5. 逆に、以下の場合はissueとして報告する：\n' +
'   - 合計金額が一致しない場合\n' +
'   - 差額が値引き額と大きく乖離している場合\n' +
'\n' +
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
'出力形式\n' +
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
'\n' +
'必ず以下のJSON形式で回答してください。説明文は不要です。\n' +
'\n' +
'{\n' +
'  "yourReading": {\n' +
'    "storeName": "あなたが画像から読み取った発行者名",\n' +
'    "storeNameEvidence": "画像のどこに書いてあったか",\n' +
'    "date": "YYYY-MM-DD",\n' +
'    "totalAmount": 数値,\n' +
'    "registrationNumber": "T+13桁 または null",\n' +
'    "taxable10": 数値,\n' +
'    "tax10": 数値,\n' +
'    "taxable8": 数値,\n' +
'    "tax8": 数値,\n' +
'    "nonTaxable": 数値\n' +
'  },\n' +
'  "comparison": {\n' +
'    "storeName": {\n' +
'      "match": true,\n' +
'      "original": "既存の値",\n' +
'      "yours": "あなたの値",\n' +
'      "correct": "正しい値",\n' +
'      "reason": "判定理由"\n' +
'    },\n' +
'    "date": {\n' +
'      "match": true,\n' +
'      "original": "既存の値",\n' +
'      "yours": "あなたの値",\n' +
'      "correct": "正しい値",\n' +
'      "reason": "判定理由"\n' +
'    },\n' +
'    "totalAmount": {\n' +
'      "match": true,\n' +
'      "original": 既存の値,\n' +
'      "yours": あなたの値,\n' +
'      "correct": 正しい値,\n' +
'      "reason": "判定理由"\n' +
'    },\n' +
'    "taxable10": {\n' +
'      "match": true,\n' +
'      "original": 既存の値,\n' +
'      "yours": あなたの値,\n' +
'      "correct": 正しい値,\n' +
'      "reason": "判定理由"\n' +
'    },\n' +
'    "tax10": {\n' +
'      "match": true,\n' +
'      "original": 既存の値,\n' +
'      "yours": あなたの値,\n' +
'      "correct": 正しい値,\n' +
'      "reason": "判定理由"\n' +
'    }\n' +
'  },\n' +
'  "overallStatus": "OK",\n' +
'  "overallConfidence": 0.95,\n' +
'  "hasHandwriting": false,\n' +
'  "isComplexReceipt": false,\n' +
'  "issues": [\n' +
'    {\n' +
'      "category": "storeName",\n' +
'      "severity": "high",\n' +
'      "field": "storeName",\n' +
'      "currentValue": "既存の誤った値",\n' +
'      "correctValue": "正しい値",\n' +
'      "reason": "詳細な理由",\n' +
'      "confidence": 0.85,\n' +
'      "evidence": "画像のどこに証拠があるか"\n' +
'    }\n' +
'  ],\n' +
'  "suggestions": []\n' +
'}\n' +
'\n' +
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
'最終チェックリスト（yourReading記入後に必ず確認）\n' +
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
'\n' +
'□ 合計金額の検算\n' +
'  taxable10 + tax10 + taxable8 + tax8 + nonTaxable = totalAmount\n' +
'\n' +
'  ⚠️ 差異が5円以上ある場合、内税/外税の判定が間違っている可能性\n' +
'\n' +
'□ 消費税の再計算\n' +
'  taxable10 × 0.1 ≒ tax10（±2円）\n' +
'  taxable8 × 0.08 ≒ tax8（±2円）\n' +
'\n' +
'  ⚠️ 大きくずれる場合、税抜/税込の判定が間違っている可能性\n' +
'\n' +
'【最終チェックリスト】\n' +
'以下を確認してからJSONを出力してください：\n' +
'\n' +
'□ yourReading の storeName は、画像から読み取りましたか？\n' +
'□ yourReading の totalAmount は、画像から読み取りましたか？\n' +
'□ yourReading の taxable10 は、画像から読み取りましたか？\n' +
'□ yourReading の tax10 は、画像から読み取りましたか？\n' +
'□ 既存結果からコピーした値はありませんか？\n' +
'□ 画像を実際に確認しましたか？\n' +
'□ 「外税」表記を正しく解釈しましたか？\n' +
'\n' +
'【重要な注意事項】\n' +
'- 問題がない場合でも、yourReading と comparison は必ず出力してください\n' +
'- yourReading の値が既存と一致していても、それは"画像から読み取った結果が一致した"であり、"コピーした"ではありません\n' +
'- 確信が持てない場合は confidence を低めに設定してください\n' +
'- 画像が不鮮明で判読できない場合は overallStatus を "ERROR" としてください\n';

  // 計算エラーがある場合の追加指示
  if (calcIssues.length > 0) {
    var calcErrorMsg = '\n\n' +
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
'【重要】計算の不整合が検出されています\n' +
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
'\n' +
'以下の計算エラーが検出されました：\n' +
calcIssues.map(function(issue) { return '- ' + issue.reason; }).join('\n') + '\n' +
'\n' +
'これは元の読み取り（既存結果）が間違っている可能性が高いです。\n' +
'画像を注意深く確認して、正しい内訳を提案してください。\n' +
'\n' +
'【特に注意すべき点】\n' +
'1. 「外税」「税込」「税抜」の表記を正しく解釈する\n' +
'   - 「(外8% 対象 ¥398)」→ これは税抜398円\n' +
'   - 「(外税8% ¥31)」→ これは消費税31円\n' +
'\n' +
'2. 税率と金額の対応を確認\n' +
'   - 10%対象額が50円なのに消費税が5円 → おかしい（50円 × 0.1 = 5円）\n' +
'   - 正しくは「税抜5円、消費税0円（端数切り捨て）」の可能性\n' +
'\n' +
'3. 合計金額から逆算\n' +
'   - 合計 = (税抜10% + 税10%) + (税抜8% + 税8%) + 不課税\n' +
'   - この式が成立する内訳を提案\n' +
'\n' +
'yourReading には、あなたが画像から読み取った正しい値を記録してください。';

    prompt = prompt + calcErrorMsg;
  }

  return prompt;
}

/*
 * 旧プロンプト（2026-02-01まで使用 — 読み取り結果を先に提示する方式）
 * 問題点: AIが既存結果に引きずられ、宛名を店名と誤認するケース等を見逃す
 *
 * function buildVerificationPrompt_OLD(rowData) {
 *   return 'あなたは経理処理の専門家です。\n' +
 *   'レシート画像と、AIが読み取った結果を比較して、間違いを指摘してください。\n' +
 *   '【読み取り結果】\n- 日付: ' + rowData.date + ' ...(以下略)';
 * }
 */

/**
 * Gemini APIを呼び出して検証結果を取得
 * @param {string} apiKey
 * @param {string} base64 - 画像のbase64データ
 * @param {string} mimeType
 * @param {Object} rowData - 読み取り結果
 * @param {Array} calcIssues - 計算エラーの配列（オプション）
 * @return {Object} 検証結果JSON
 */
function callGeminiForVerification_(apiKey, base64, mimeType, rowData, calcIssues) {
  var model = CONFIG.GEMINI.MODEL;
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' +
            model + ':generateContent?key=' + apiKey;

  var prompt = buildVerificationPrompt_(rowData, calcIssues);

  var payload = {
    contents: [{
      parts: [
        { text: prompt },
        {
          inline_data: {
            mime_type: mimeType,
            data: base64
          }
        }
      ]
    }],
    generationConfig: {
      temperature: 0.1,
      maxOutputTokens: 8192,
      responseMimeType: 'application/json'
    }
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var MAX_RETRIES = 3;
  var lastError = null;

  for (var attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    try {
      var response = UrlFetchApp.fetch(url, options);
      var statusCode = response.getResponseCode();

      if (statusCode !== 200) {
        lastError = 'HTTP ' + statusCode;
        console.warn('Verification API error (attempt ' + attempt + '/' + MAX_RETRIES + '): ' +
                     statusCode + ' - ' + response.getContentText().slice(0, 300));
        if (attempt < MAX_RETRIES) {
          Utilities.sleep(2000 * attempt);
          continue;
        }
        throw new Error('Gemini API エラー: ' + statusCode);
      }

      var apiResult = JSON.parse(response.getContentText());

      // レスポンス構造を検証
      if (!apiResult.candidates || !apiResult.candidates[0]) {
        lastError = 'レスポンスにcandidatesがありません';
        if (attempt < MAX_RETRIES) {
          Utilities.sleep(2000 * attempt);
          continue;
        }
        throw new Error(lastError);
      }

      var candidate = apiResult.candidates[0];

      // finishReasonチェック
      if (candidate.finishReason && candidate.finishReason !== 'STOP') {
        lastError = 'finishReason: ' + candidate.finishReason;
        if (attempt < MAX_RETRIES) {
          Utilities.sleep(2000 * attempt);
          continue;
        }
        throw new Error(lastError);
      }

      var text = (candidate.content && candidate.content.parts &&
                  candidate.content.parts[0] && candidate.content.parts[0].text) || '';

      if (!text.trim()) {
        lastError = 'レスポンスのテキストが空です';
        if (attempt < MAX_RETRIES) {
          Utilities.sleep(2000 * attempt);
          continue;
        }
        throw new Error(lastError);
      }

      return parseVerificationResponse_(text);

    } catch (e) {
      lastError = e.message;
      console.warn('Verification exception (attempt ' + attempt + '/' + MAX_RETRIES + '): ' + e.message);
      if (attempt < MAX_RETRIES) {
        Utilities.sleep(2000 * attempt);
        continue;
      }
      throw e;
    }
  }

  throw new Error('検証API: 全リトライ失敗 - ' + lastError);
}

/**
 * Geminiのレスポンスをパース
 * @param {string} text - APIレスポンスのテキスト
 * @return {Object} 検証結果
 */
function parseVerificationResponse_(responseText) {
  Logger.log('=== parseVerificationResponse_ デバッグ ===');
  Logger.log('入力の型: ' + typeof responseText);

  // 配列の場合は最初の要素を取得
  if (Array.isArray(responseText) && responseText.length > 0) {
    Logger.log('配列を検出: 最初の要素を使用');
    responseText = responseText[0];
  }

  // output_text 形式の場合、text フィールドを抽出
  if (typeof responseText === 'object' && responseText !== null) {
    if (responseText.type === 'output_text' && responseText.text) {
      Logger.log('output_text タイプを検出: text フィールドを抽出');
      responseText = responseText.text;
    }
  }

  // 安全なログ出力
  if (typeof responseText === 'string') {
    Logger.log('入力の長さ: ' + responseText.length);
    Logger.log('入力の最初の500文字: ' + responseText.substring(0, Math.min(500, responseText.length)));
  } else if (typeof responseText === 'object' && responseText !== null) {
    // すでにパース済みのJSONオブジェクトの場合
    if (responseText.yourReading || responseText.comparison || responseText.issues) {
      Logger.log('すでにパース済みのJSONオブジェクトです');
      return responseText;
    }
    var jsonStr = JSON.stringify(responseText);
    Logger.log('入力の長さ: ' + jsonStr.length);
    Logger.log('入力の最初の500文字: ' + jsonStr.substring(0, Math.min(500, jsonStr.length)));
  }

  // 以下、既存のパース処理を続行
  var text = responseText;

  if (typeof text !== 'string') {
    text = JSON.stringify(text);
  }

  if (!text) {
    throw new Error('responseText が空またはnullです');
  }

  text = text.trim();

  Logger.log('trim後の長さ: ' + text.length);
  Logger.log('trim後の最初の200文字: ' + text.substring(0, Math.min(200, text.length)));

  // JSONブロックを抽出
  var jsonMatch = text.match(/```json\s*([\s\S]*?)\s*```/) || text.match(/```\s*([\s\S]*?)\s*```/);

  if (jsonMatch) {
    text = jsonMatch[1].trim();
    Logger.log('JSONブロック抽出成功、長さ: ' + text.length);
  }

  // JSONをパース
  var parsed;
  try {
    parsed = JSON.parse(text);
    Logger.log('✅ JSONパース成功');
    Logger.log('result.yourReading が存在: ' + !!parsed.yourReading);
    Logger.log('result.comparison が存在: ' + !!parsed.comparison);
  } catch (error) {
    Logger.log('❌ JSONパースエラー: ' + error.toString());
    if (text && text.length > 0) {
      Logger.log('パース対象テキスト（最初の1000文字）: ' + text.substring(0, Math.min(1000, text.length)));
    } else {
      Logger.log('パース対象テキストが空です');
    }
    throw new Error('検証結果のJSONパースに失敗しました: ' + error.toString());
  }

  // 必須フィールドのバリデーション
  if (!parsed.overallStatus) {
    parsed.overallStatus = 'WARNING';
  }
  if (typeof parsed.overallConfidence !== 'number') {
    parsed.overallConfidence = 0.5;
  }
  if (!Array.isArray(parsed.issues)) {
    parsed.issues = [];
  }
  if (typeof parsed.hasHandwriting !== 'boolean') {
    parsed.hasHandwriting = false;
  }
  if (typeof parsed.isComplexReceipt !== 'boolean') {
    parsed.isComplexReceipt = false;
  }
  if (!Array.isArray(parsed.suggestions)) {
    parsed.suggestions = [];
  }

  // ブラインド検証方式の新フィールド
  if (!parsed.yourReading || typeof parsed.yourReading !== 'object') {
    Logger.log('⚠️ yourReading が空またはオブジェクトでないため、空オブジェクトを設定');
    parsed.yourReading = {};
  }
  if (!parsed.comparison || typeof parsed.comparison !== 'object') {
    Logger.log('⚠️ comparison が空またはオブジェクトでないため、空オブジェクトを設定');
    parsed.comparison = {};
  }

  return parsed;
}

/**
 * 内税/外税の判定ミスを自動修正
 * @param {Object} result - GPT-5の検証結果
 * @param {number} totalAmount - 合計金額
 * @return {Object} 修正後の結果
 */
function fixTaxCalculationError_(result, totalAmount) {
  if (!result.yourReading) {
    return result;
  }

  const yr = result.yourReading;

  // 計算された合計
  const calculated =
    (yr.taxable10 || 0) +
    (yr.tax10 || 0) +
    (yr.taxable8 || 0) +
    (yr.tax8 || 0) +
    (yr.nonTaxable || 0);

  const diff = Math.abs(calculated - totalAmount);

  // 差異が5円未満なら問題なし
  if (diff < 5) {
    return result;
  }

  // 消費税の合計
  const totalTax = (yr.tax10 || 0) + (yr.tax8 || 0);

  // 差異が消費税合計と近い場合、内税を外税と誤解している可能性
  if (Math.abs(diff - totalTax) < 5) {
    Logger.log('内税/外税の判定ミスを検出。自動修正を試みます。');
    Logger.log('修正前: calculated=' + calculated + ', total=' + totalAmount + ', diff=' + diff);

    // taxable から tax を引く
    yr.taxable10 = Math.max(0, (yr.taxable10 || 0) - (yr.tax10 || 0));
    yr.taxable8 = Math.max(0, (yr.taxable8 || 0) - (yr.tax8 || 0));

    // 再計算
    const recalculated =
      (yr.taxable10 || 0) +
      (yr.tax10 || 0) +
      (yr.taxable8 || 0) +
      (yr.tax8 || 0) +
      (yr.nonTaxable || 0);

    const newDiff = Math.abs(recalculated - totalAmount);

    Logger.log('修正後: recalculated=' + recalculated + ', newDiff=' + newDiff);

    // 修正後の差異が改善された場合
    if (newDiff < diff) {
      Logger.log('修正成功。内税表記として処理しました。');

      // issuesに記録
      if (!result.issues) {
        result.issues = [];
      }
      result.issues.push({
        category: 'taxCalculation',
        severity: 'medium',
        reason: '内税/外税の判定ミスを自動修正（差異 ' + diff + '円 → ' + newDiff + '円）',
        autoFixed: true
      });
    } else {
      // 修正が失敗した場合は元に戻す
      Logger.log('修正失敗。元の値を維持します。');
      yr.taxable10 = (yr.taxable10 || 0) + (yr.tax10 || 0);
      yr.taxable8 = (yr.taxable8 || 0) + (yr.tax8 || 0);
    }
  }

  return result;
}

// ============================================================
// GPT-5 API呼び出し
// ============================================================

/**
 * OpenAI APIキーを取得
 */
function getOpenAIApiKey_() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) {
    throw new Error('OpenAI APIキーが設定されていません。ScriptPropertiesに OPENAI_API_KEY を設定してください。');
  }
  return apiKey;
}

/**
 * GPT-5 APIを呼び出して検証を実行
 * @param {Blob} imageFile - 画像ファイル（PDF/画像）
 * @param {string} base64Image - Base64エンコードされたファイル
 * @param {string} prompt - プロンプト
 * @return {Object} GPT-5 APIのレスポンス
 */
function callGPT5ForVerification_(imageFile, base64Image, prompt) {
  const apiKey = getOpenAIApiKey_();
  const url = 'https://api.openai.com/v1/responses';

  Logger.log('=== API呼び出し詳細 ===');
  Logger.log('プロンプト長: ' + prompt.length + ' 文字');
  Logger.log('画像サイズ: ' + base64Image.length + ' 文字');

  const mimeType = imageFile.getContentType();
  const fileName = imageFile.getName() || 'receipt.pdf';

  Logger.log('ファイルタイプ: ' + mimeType + ', ファイル名: ' + fileName);

  // Responses API の正しい形式
  const input = [
    {
      type: 'message',  // これが正しい
      role: 'user',
      content: [
        {
          type: 'input_text',
          text: prompt
        },
        {
          type: 'input_file',
          filename: fileName,
          file_data: 'data:' + mimeType + ';base64,' + base64Image
        }
      ]
    }
  ];

  const payload = {
    model: 'gpt-5',
    input: input,
    max_output_tokens: 16000
  };

  Logger.log('ペイロード作成完了');

  const options = {
    method: 'post',
    headers: {
      'Authorization': 'Bearer ' + apiKey,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  let lastError;
  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      Logger.log('GPT-5 Responses API呼び出し（試行 ' + attempt + '/3）');
      Logger.log('API URL: ' + url);

      const startTime = Date.now();
      const response = UrlFetchApp.fetch(url, options);
      const elapsed = Date.now() - startTime;

      Logger.log('レスポンス受信: ' + elapsed + 'ms経過');

      const statusCode = response.getResponseCode();
      Logger.log('ステータスコード: ' + statusCode);

      if (statusCode === 200) {
        const responseText = response.getContentText();
        Logger.log('レスポンステキスト長: ' + responseText.length + ' 文字');

        const jsonResponse = JSON.parse(responseText);
        Logger.log('GPT-5 Responses API呼び出し成功');
        return jsonResponse;
      } else {
        const errorText = response.getContentText();
        Logger.log('GPT-5 Responses APIエラー: ' + errorText);
        lastError = new Error('API returned status ' + statusCode + ': ' + errorText);
      }

    } catch (error) {
      Logger.log('❌ 例外発生: ' + error.toString());
      Logger.log('例外の型: ' + error.name);
      lastError = error;

      if (attempt < 3) {
        const waitTime = Math.pow(2, attempt) * 1000;
        Logger.log('リトライ前に ' + waitTime + 'ms 待機');
        Utilities.sleep(waitTime);
      }
    }
  }

  throw lastError;
}

/**
 * GPT-5のレスポンスからテキストを抽出
 * Responses API形式とChat Completions API形式の両方に対応
 * @param {Object} response - GPT-5 APIのレスポンス
 * @return {string} 抽出されたテキスト
 */
function extractGPT5Text_(response) {
  Logger.log('=== extractGPT5Text_ デバッグ ===');
  Logger.log('レスポンスの型: ' + typeof response);
  Logger.log('レスポンスは配列: ' + Array.isArray(response));

  if (!response) {
    throw new Error('GPT-5 APIのレスポンスがnullです');
  }

  // 配列の場合は最初の要素を取得
  if (Array.isArray(response) && response.length > 0) {
    Logger.log('配列形式のレスポンスを検出: 最初の要素を使用');
    response = response[0];
  }

  // output_text 形式 (type フィールド付き)
  if (response && response.type === 'output_text' && response.text) {
    Logger.log('output_text 形式を検出');
    return response.text;
  }

  // output_text フィールド
  if (response && response.output_text) {
    Logger.log('output_text フィールドを検出');
    return response.output_text;
  }

  // output 配列がある場合
  if (response && response.output && Array.isArray(response.output)) {
    Logger.log('output 配列を検出');
    var textOutput = response.output.find(function(item) {
      return item.type === 'text' || item.type === 'message' || item.type === 'output_text';
    });
    if (textOutput && textOutput.text) {
      Logger.log('output配列からtextを取得');
      return textOutput.text;
    }
    if (textOutput && textOutput.content) {
      Logger.log('output配列からcontentを取得');
      return textOutput.content;
    }
  }

  // Chat Completions形式
  if (response && response.choices && Array.isArray(response.choices)) {
    if (response.choices.length === 0) {
      throw new Error('GPT-5 APIのレスポンスにテキストが含まれていません');
    }
    Logger.log('Chat Completions 形式を検出');
    return response.choices[0].message.content;
  }

  Logger.log('❌ 認識できないレスポンス形式');
  Logger.log('レスポンスキー: ' + Object.keys(response).join(', '));
  throw new Error('APIのレスポンス形式が不正です');
}

/**
 * GPT-5 API接続テスト
 */
function diagnoseGPT5Api() {
  Logger.log('=== GPT-5 API診断開始 ===');

  // ステップ1: APIキーの取得
  Logger.log('\n[ステップ1] APIキーの取得');
  try {
    const apiKey = getOpenAIApiKey_();
    Logger.log('✅ APIキー取得成功');
    Logger.log('   先頭10文字: ' + apiKey.substring(0, 10));
  } catch (error) {
    Logger.log('❌ APIキー取得失敗: ' + error.toString());
    return;
  }

  // ステップ2: 最小限のリクエストテスト
  Logger.log('\n[ステップ2] 最小限のAPIリクエスト');
  const apiKey = getOpenAIApiKey_();
  const url = 'https://api.openai.com/v1/chat/completions';

  const payload = {
    model: 'gpt-5',
    messages: [{
      role: 'user',
      content: 'Hi'
    }],
    max_output_tokens: 10
  };

  const options = {
    method: 'post',
    headers: {
      'Authorization': 'Bearer ' + apiKey,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();
    const responseText = response.getContentText();

    Logger.log('   ステータスコード: ' + statusCode);

    if (statusCode === 200) {
      Logger.log('✅ API接続成功！');
      const json = JSON.parse(responseText);
      Logger.log('   レスポンス: ' + json.choices[0].message.content);
    } else {
      Logger.log('❌ APIエラー（' + statusCode + '）');
      Logger.log('   レスポンス: ' + responseText);
    }
  } catch (error) {
    Logger.log('❌ リクエスト失敗: ' + error.toString());
  }

  Logger.log('\n=== 診断完了 ===');
}

/**
 * 画像ファイルをBase64にエンコード
 */
function convertToBase64_(file) {
  const bytes = file.getBytes();
  return Utilities.base64Encode(bytes);
}

/**
 * 行に対応する画像ファイルを取得
 */
function getImageFileForRow_(sheet, row) {
  // B列のHYPERLINK数式からファイルIDを抽出
  const formula = sheet.getRange(row, 2).getFormula() || '';
  let fileId = '';
  const urlMatch = formula.match(/HYPERLINK\("([^"]+)"/);
  if (urlMatch) {
    const idMatch = urlMatch[1].match(/\/d\/([^\/]+)/);
    if (idMatch) fileId = idMatch[1];
  }

  if (!fileId) {
    throw new Error('ファイルIDが取得できません（B列にHYPERLINKがありません）');
  }

  // Driveからファイルを取得
  try {
    const file = DriveApp.getFileById(fileId);
    return file.getBlob();
  } catch (e) {
    throw new Error('ファイル取得失敗: ' + e.message);
  }
}

// ============================================================
// 結果書き込み
// ============================================================

/**
 * 検証用の列ヘッダーを確保（17-20列目）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function ensureVerificationColumns_(sheet) {
  var lastCol = sheet.getLastColumn();

  // 既に17列目以降にヘッダーがあるかチェック
  if (lastCol >= 17) {
    var existing = sheet.getRange(1, 17, 1, 1).getValue();
    if (existing === '検証ステータス') {
      return; // 既に設定済み
    }
  }

  // ヘッダーを書き込み
  sheet.getRange(1, 17, 1, 4).setValues([['検証ステータス', '検証スコア', '検証結果', '修正案JSON']]);
  sheet.getRange(1, 17, 1, 4).setFontWeight('bold');

  // 列幅を調整
  sheet.setColumnWidth(17, 110); // 検証ステータス
  sheet.setColumnWidth(18, 80);  // 検証スコア
  sheet.setColumnWidth(19, 300); // 検証結果
  sheet.setColumnWidth(20, 200); // 修正案JSON
}

/**
 * フィールド名を日本語ラベルに変換
 * @param {string} field - フィールド名
 * @return {string} 日本語ラベル
 */
function getFieldLabel_(field) {
  var labels = {
    'storeName': '店名',
    'date': '日付',
    'totalAmount': '総額',
    'registrationNumber': '登録番号',
    'taxable10': '税抜10%',
    'tax10': '消費税10%',
    'taxable8': '税抜8%',
    'tax8': '消費税8%',
    'nonTaxable': '不課税',
    'account': '勘定科目'
  };
  return labels[field] || field;
}

/**
 * 簡潔な検証結果サマリーを生成
 * @param {Object} result - 検証結果
 * @return {string} 人間が読める説明
 */
function buildCompactSummary_(result) {
  if (!result.issues || result.issues.length === 0) {
    return '✅ 問題なし';
  }

  // 重大度が high の問題を優先的に表示
  var highPriorityIssues = result.issues
    .filter(function(issue) { return issue.severity === 'high'; })
    .slice(0, 2); // 最大2件

  if (highPriorityIssues.length === 0) {
    // high がない場合は、全体から2件
    var anyIssues = result.issues.slice(0, 2);
    return anyIssues.map(function(issue) {
      var label = getFieldLabel_(issue.field);
      return '⚠️ ' + label + ': ' + issue.currentValue + ' → ' + issue.correctValue;
    }).join('\n');
  }

  return highPriorityIssues.map(function(issue) {
    var label = getFieldLabel_(issue.field);
    var current = issue.currentValue || '（空）';
    var correct = issue.correctValue || '（空）';
    return '⚠️ ' + label + ': ' + current + ' → ' + correct;
  }).join('\n');
}

/**
 * ステータスに応じた絵文字を取得
 * @param {string} status - ステータス
 * @return {string} 絵文字付きステータス
 */
function getStatusEmoji_(status) {
  switch(status) {
    case 'OK':
      return '🟢自動確定';
    case 'WARNING':
      return '🟡要確認';
    case 'ERROR':
      return '🔴要入力';
    default:
      return '🔴エラー';
  }
}

/**
 * 検証結果を17-20列目に書き込み
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} row
 * @param {Object} result - 検証結果JSON
 */
function writeVerificationResult_(sheet, row, result) {
  // ステータス判定（issues があれば強制的に WARNING または ERROR）
  if (result.issues && result.issues.length > 0) {
    var highSeverityCount = result.issues.filter(function(issue) {
      return issue.severity === 'high';
    }).length;

    if (highSeverityCount >= 2) {
      result.overallStatus = 'ERROR';
    } else if (highSeverityCount >= 1) {
      result.overallStatus = 'WARNING';
    } else {
      result.overallStatus = 'WARNING';
    }
  }

  // 17列目：検証ステータス
  var statusEmoji = getStatusEmoji_(result.overallStatus);
  var statusCell = sheet.getRange(row, 17);
  statusCell.setValue(statusEmoji);

  // ステータスに応じて背景色を設定
  if (result.overallStatus === 'OK') {
    statusCell.setBackground('#d4edda'); // 緑
  } else if (result.overallStatus === 'WARNING') {
    statusCell.setBackground('#fff3cd'); // 黄
  } else {
    statusCell.setBackground('#f8d7da'); // 赤
  }

  // 18列目：検証スコア
  sheet.getRange(row, 18).setValue(result.overallConfidence);

  // 19列目：人間が読める説明（簡潔版）
  var summary = buildCompactSummary_(result);
  var summaryCell = sheet.getRange(row, 19);
  summaryCell.setValue(summary);
  summaryCell.setWrap(true); // 折り返しOK

  // 19列目の列幅を設定
  sheet.setColumnWidth(19, 250);

  // 20列目：修正案JSON（圧縮版、折り返しなし）
  var jsonString = JSON.stringify(result); // 改行なしの圧縮版
  var jsonCell = sheet.getRange(row, 20);
  jsonCell.setValue(jsonString);
  jsonCell.setWrap(false); // 折り返しNG

  // 20列目の列幅を固定
  sheet.setColumnWidth(20, 100);

  // 行の高さを60pxに固定（2-3行分）
  sheet.setRowHeight(row, 60);

  console.log('Row ' + row + ': 検証結果を書き込み完了');
}

/**
 * 検証エラーを書き込み
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} row
 * @param {string} errorMsg
 */
function writeVerificationError_(sheet, row, errorMsg) {
  var errorResult = {
    overallStatus: 'ERROR',
    overallConfidence: 0,
    hasHandwriting: false,
    isComplexReceipt: false,
    issues: [],
    suggestions: [],
    error: errorMsg
  };

  sheet.getRange(row, 17, 1, 4).setValues([
    ['🔴エラー', 0, 'エラー: ' + errorMsg, JSON.stringify(errorResult, null, 2)]
  ]);
  sheet.getRange(row, 17, 1, 4).setBackground('#FFEBEE');
}

// ============================================================
// ワンクリック修正適用（サイドバーから呼び出し）
// ============================================================

/**
 * フィールド名 → 列番号のマッピング
 */
var FIELD_COLUMN_MAP_ = {
  'date':               4,
  'storeName':          5,
  'registrationNumber': 6,
  'totalAmount':        7,
  'taxable10':          8,
  'subtotal10':         8,
  'tax10':              9,
  'taxable8':           10,
  'subtotal8':          10,
  'tax8':               11,
  'nonTaxable':         12,
  'accountTitle':       13,
  'account':            13
};

// ============================================================
// 自動検証＋自動承認
// ============================================================

/**
 * 未検証行を一括でGPT-5検証し、条件を満たせば自動でOKに承認する。
 *
 * 対象行: A列が CHECK / ERROR / COMPOUND（不課税0円の場合のみ）で、17列目が空欄
 * 自動承認条件（3つすべて満たす場合のみ）:
 *   1. 確度スコア >= 0.90
 *   2. 合計金額が完全一致（差0円）
 *   3. severity: high の問題がゼロ
 */
function runAutoVerification() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    console.log('runAutoVerification: Skip - already running');
    return;
  }

  try {
    const startTime = Date.now();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME.MAIN);

    if (!sheet) {
      console.log('runAutoVerification: シート「' + CONFIG.SHEET_NAME.MAIN + '」が見つかりません');
      return;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      console.log('runAutoVerification: データがありません');
      return;
    }

    // 検証用列の確保
    ensureVerificationColumns_(sheet);

    // 全データを一括取得（パフォーマンス最適化）
    const lastCol = Math.max(sheet.getLastColumn(), 20);
    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    // 対象行を抽出
    const targetRows = [];
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const rawStatus = String(row[0] || '');
      // 絵文字を除去してステータスを取得
      const status = rawStatus.replace(/^[🟢🔴🟡🟠🖊️]+/, '');
      const verificationStatus = String(row[16] || ''); // 17列目（0-indexed: 16）
      const nonTaxable = Number(row[11]) || 0; // 12列目（0-indexed: 11）

      // 対象判定
      if (verificationStatus !== '') continue; // 既に検証済み
      if (status === 'CHECK' || status === 'ERROR') {
        targetRows.push(i + 2); // シートの行番号（1-indexed、ヘッダー+1）
      } else if (status === 'COMPOUND' && nonTaxable === 0) {
        // COMPOUNDは不課税0円のみ対象
        targetRows.push(i + 2);
      }
    }

    console.log('runAutoVerification: 対象行 = ' + targetRows.length + '件');

    let processedCount = 0;
    let approvedCount = 0;
    let pendingCount = 0;
    let errorCount = 0;

    for (const row of targetRows) {
      // タイムアウトチェック
      if (Date.now() - startTime > CONFIG.PROCESSING.MAX_EXECUTION_TIME_MS) {
        console.log('runAutoVerification: タイムアウト（' + processedCount + '/' + targetRows.length + '件処理済み）');

        // 未処理行が残っている場合、継続トリガーを設定
        if (processedCount < targetRows.length) {
          deleteContinuationTrigger_('runAutoVerification');
          ScriptApp.newTrigger('runAutoVerification')
            .timeBased()
            .after(1 * 60 * 1000)  // 1分後
            .create();
          console.log('継続トリガーを設定しました（1分後に再実行）');
          try {
            ss.toast('タイムアウト。1分後に自動で続きを実行します。', '継続予定', 5);
          } catch (e) { /* トリガー実行時はUI非対応 */ }
        }
        break;
      }

      try {
        // GPT-5で検証（既存の verifyOneRow_ を利用）
        verifyOneRow_(sheet, row, null);
        processedCount++;

        // 検証結果を読み取り
        const verResult = readVerificationResult_(sheet, row);

        // 自動承認条件の判定
        if (shouldAutoApprove_(verResult)) {
          applyApproval_(sheet, row, '🤖 自動承認');
          approvedCount++;
          console.log('行' + row + ': 自動承認 (score=' + verResult.score + ')');
        } else {
          // 要確認マーク（17列目のみ更新、A列は変更しない）
          sheet.getRange(row, 17).setValue('🟡要確認');
          pendingCount++;
          console.log('行' + row + ': 要確認 (score=' + verResult.score + ')');
        }

      } catch (e) {
        console.error('行' + row + ': 検証エラー - ' + e.message);
        writeVerificationError_(sheet, row, e.message);
        errorCount++;
      }
    }

    // 全件処理完了の場合、継続トリガーを削除
    if (processedCount >= targetRows.length || targetRows.length === 0) {
      deleteContinuationTrigger_('runAutoVerification');
      console.log('全件処理完了。継続トリガーを削除しました。');
    }

    const summary = 'runAutoVerification 完了: ' +
      '対象=' + targetRows.length + '件, ' +
      '処理=' + processedCount + '件, ' +
      '自動承認=' + approvedCount + '件, ' +
      '要確認=' + pendingCount + '件, ' +
      'エラー=' + errorCount + '件';
    console.log(summary);

    // UI表示（トリガー実行時はUIがないためtry-catch）
    try {
      SpreadsheetApp.getUi().alert(summary);
    } catch (e) { /* トリガー実行時はUI非対応 */ }

  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// 手動承認
// ============================================================

/**
 * 選択された行を手動で承認する。
 * 人間が17列目（検証結果）を確認した後に使うボタン。
 *
 * 対象: CHECK / ERROR / COMPOUND 行のみ（OKやHANDは何もしない）
 * 複数の選択範囲（Ctrl+クリック）にも対応。
 */
function approveSelectedRows() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (sheet.getName() !== CONFIG.SHEET_NAME.MAIN) {
    ui.alert('「' + CONFIG.SHEET_NAME.MAIN + '」シートで実行してください。');
    return;
  }

  // 複数の選択範囲に対応するため getActiveRangeList() を使用
  const rangeList = sheet.getActiveRangeList();
  if (!rangeList) {
    ui.alert('承認したい行を選択してください。');
    return;
  }

  // 選択されている全ての行番号を収集（重複排除）
  const selectedRows = new Set();
  const ranges = rangeList.getRanges();

  for (const range of ranges) {
    const startRow = range.getRow();
    const numRows = range.getNumRows();
    for (let row = startRow; row < startRow + numRows; row++) {
      if (row > 1) {  // ヘッダー行は除外
        selectedRows.add(row);
      }
    }
  }

  if (selectedRows.size === 0) {
    ui.alert('承認したい行を選択してください（ヘッダー行は除外されます）。');
    return;
  }

  // デバッグログ
  const rowArray = Array.from(selectedRows).sort((a, b) => a - b);
  console.log('approveSelectedRows: 対象行 = ' + rowArray.join(', ') + ' (' + rowArray.length + '件)');

  let approvedCount = 0;
  let skippedCount = 0;

  for (const row of rowArray) {
    const rawStatus = String(sheet.getRange(row, 1).getValue() || '');
    const status = rawStatus.replace(/^[🟢🔴🟡🟠🖊️]+/, '');

    console.log('行' + row + ': rawStatus="' + rawStatus + '", status="' + status + '"');

    // CHECK / ERROR / COMPOUND のみ承認可能
    if (status !== 'CHECK' && status !== 'ERROR' && status !== 'COMPOUND') {
      console.log('行' + row + ': スキップ（ステータスが対象外）');
      skippedCount++;
      continue;
    }

    applyApproval_(sheet, row, '✅ 手動承認');
    approvedCount++;
    console.log('行' + row + ': 承認完了');
  }

  const msg = '承認完了: ' + approvedCount + '件' +
              (skippedCount > 0 ? '（スキップ: ' + skippedCount + '件）' : '');
  console.log('approveSelectedRows 結果: ' + msg);
  ui.alert(msg);
}

// ============================================================
// ヘルパー関数（自動検証・承認共通）
// ============================================================

/**
 * 承認処理の共通ロジック（自動承認・手動承認で共用）
 * - A列 → OK に変更
 * - 17列目 → ラベルを設定
 * - ファイル名 → [CHK]/[ERR]/[CMP] を [OK] に変更
 * - 行の背景色をリセット
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} row - 行番号
 * @param {string} label - 17列目に書き込むラベル（例: '🤖 自動承認', '✅ 手動承認'）
 */
function applyApproval_(sheet, row, label) {
  // A列をOKに変更
  sheet.getRange(row, 1).setValue('🟢OK');

  // 17列目にラベルを書き込み
  sheet.getRange(row, 17).setValue(label);
  sheet.getRange(row, 17).setBackground('#d4edda'); // 緑

  // 行の背景色をリセット（OK行のデフォルト = 白）
  const lastCol = Math.max(sheet.getLastColumn(), 20);
  sheet.getRange(row, 1, 1, lastCol).setBackground(null);

  // Google Drive上のファイル名プレフィックスを変更
  renameFilePrefix_(sheet, row);
}

/**
 * Drive上のファイル名のステータスプレフィックスを [OK] に変更する
 * [CHK] / [ERR] / [CMP] → [OK]
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} row - 行番号
 */
function renameFilePrefix_(sheet, row) {
  try {
    // B列のHYPERLINK数式からファイルIDを抽出
    const formula = sheet.getRange(row, 2).getFormula() || '';
    const urlMatch = formula.match(/HYPERLINK\("([^"]+)"/);
    if (!urlMatch) return;

    const idMatch = urlMatch[1].match(/\/d\/([^\/]+)/);
    if (!idMatch) return;

    const fileId = idMatch[1];
    const file = DriveApp.getFileById(fileId);
    const currentName = file.getName();

    // プレフィックスを [OK] に置換
    let newName = currentName;
    newName = newName.replace(/^\[(CHK|ERR|CMP)\]/, '[OK]');
    // 旧形式の絵文字プレフィックスにも対応
    newName = newName.replace(/^[🔴🟡]/, '🟢');

    if (newName !== currentName) {
      file.setName(newName);
      console.log('ファイル名変更: "' + currentName + '" → "' + newName + '"');
    }
  } catch (e) {
    // ファイル名変更はベストエフォート（失敗しても処理を継続）
    console.warn('ファイル名変更失敗 (行' + row + '): ' + e.message);
  }
}

/**
 * 検証結果を17-20列目から読み取る
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} row - 行番号
 * @return {{score: number, issues: Array, totalMatch: boolean}}
 */
function readVerificationResult_(sheet, row) {
  const values = sheet.getRange(row, 17, 1, 4).getValues()[0];
  const score = Number(values[1]) || 0;  // 18列目: 確度スコア
  const jsonStr = String(values[3] || ''); // 20列目: 修正案JSON

  let issues = [];
  let totalMatch = false;

  if (jsonStr) {
    try {
      const parsed = JSON.parse(jsonStr);
      issues = parsed.issues || [];

      // 合計金額の一致チェック: yourReading と既存の totalAmount を比較
      if (parsed.yourReading && parsed.comparison && parsed.comparison.totalAmount) {
        totalMatch = parsed.comparison.totalAmount.match === true;
      } else if (parsed.yourReading) {
        // comparison がない場合は、yourReading の検算で判定
        const yr = parsed.yourReading;
        const yrTotal = (Number(yr.taxable10) || 0) +
                        (Number(yr.tax10) || 0) +
                        (Number(yr.taxable8) || 0) +
                        (Number(yr.tax8) || 0) +
                        (Number(yr.nonTaxable) || 0);
        const yrAmount = Number(yr.totalAmount) || 0;
        totalMatch = (yrAmount > 0 && Math.abs(yrTotal - yrAmount) === 0);
      }
    } catch (e) {
      console.warn('検証JSON読み取りエラー (行' + row + '): ' + e.message);
    }
  }

  return {
    score: score,
    issues: issues,
    totalMatch: totalMatch
  };
}

/**
 * 自動承認の3条件を判定する
 *
 * @param {{score: number, issues: Array, totalMatch: boolean}} verResult
 * @return {boolean} 自動承認すべきか
 */
function shouldAutoApprove_(verResult) {
  // 条件1: 確度スコア >= 0.90
  if (verResult.score < 0.90) {
    return false;
  }

  // 条件2: 合計金額が完全一致
  if (!verResult.totalMatch) {
    return false;
  }

  // 条件3: severity: high の問題がゼロ
  const highCount = verResult.issues.filter(function(issue) {
    return issue.severity === 'high';
  }).length;
  if (highCount > 0) {
    return false;
  }

  return true;
}

// ============================================================
// 継続トリガー管理
// ============================================================

/**
 * 指定した関数名の時間ベーストリガーを削除
 * runAutoVerification の継続実行で使用する。
 * @param {string} functionName
 */
function deleteContinuationTrigger_(functionName) {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === functionName &&
        trigger.getEventType() === ScriptApp.EventType.CLOCK) {
      ScriptApp.deleteTrigger(trigger);
      console.log('既存の継続トリガーを削除: ' + functionName);
    }
  }
}

// ============================================================
// ワンクリック修正適用（サイドバーから呼び出し）
// ============================================================

/**
 * 検証結果の修正を本番シートに適用する
 * サイドバーからワンクリックで呼び出される
 * @param {number} row - 行番号
 * @param {string} field - フィールド名
 * @param {*} value - 修正後の値
 * @return {Object} 結果 {success: boolean, message: string}
 */
function applyVerificationFix(row, field, value) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME.MAIN);
    if (!sheet) {
      return { success: false, message: 'シートが見つかりません' };
    }

    var col = FIELD_COLUMN_MAP_[field];
    if (!col) {
      return { success: false, message: '不明なフィールド: ' + field };
    }

    // 値を適切な型に変換
    var cellValue = value;
    if (col >= 7 && col <= 12) {
      // 金額フィールドは数値に変換
      cellValue = Number(value) || 0;
    }

    // セルに書き込み
    sheet.getRange(row, col).setValue(cellValue);

    // 検証ステータスを「修正済み」に更新
    var currentStatus = String(sheet.getRange(row, 17).getValue() || '');
    if (currentStatus && !currentStatus.includes('修正済')) {
      sheet.getRange(row, 17).setValue('🔧修正済み');
      sheet.getRange(row, 17, 1, 4).setBackground('#E3F2FD');
    }

    console.log('修正適用: 行' + row + ' ' + field + ' = ' + cellValue);
    return { success: true, message: field + ' を修正しました' };

  } catch (e) {
    console.error('修正適用エラー: ' + e.message);
    return { success: false, message: 'エラー: ' + e.message };
  }
}
