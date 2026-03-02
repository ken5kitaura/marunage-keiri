/**
 * Service_OCR.gs
 * Gemini 2.5 Flash を使用したOCR処理
 *
 * 重要ルール:
 * - AIには「見る」ことだけをさせる
 * - 計算（税額計算・合計チェック）はGAS側で行う
 * - JSONのみを返すプロンプトを使用
 */

/**
 * ファイルからOCRデータを抽出
 * @param {GoogleAppsScript.Drive.File} file
 * @return {OCRResult}
 */
function extractOCRFromFile(file) {
  const mime = file.getMimeType();

  if (mime === 'application/pdf') {
    return extractOCRFromPDF_(file);
  } else if (mime.startsWith('image/')) {
    return extractOCRFromImage_(file);
  } else {
    throw new Error('非対応のファイル形式: ' + mime);
  }
}

/**
 * 画像ファイルからOCR
 * @param {GoogleAppsScript.Drive.File} file
 * @return {OCRResult}
 */
function extractOCRFromImage_(file) {
  const blob = file.getBlob();
  const base64 = Utilities.base64Encode(blob.getBytes());
  const mimeType = blob.getContentType();

  return callGeminiVision_(base64, mimeType);
}

/**
 * PDFファイルからOCR
 * @param {GoogleAppsScript.Drive.File} file
 * @return {OCRResult}
 */
function extractOCRFromPDF_(file) {
  const blob = file.getBlob();
  const base64 = Utilities.base64Encode(blob.getBytes());

  return callGeminiVision_(base64, 'application/pdf');
}

/**
 * Gemini Vision APIを呼び出し
 * @param {string} base64Content
 * @param {string} mimeType
 * @return {OCRResult}
 */
function callGeminiVision_(base64Content, mimeType) {
  const apiKey = getGeminiApiKey_();
  const model = CONFIG.GEMINI.MODEL;

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' +
              model + ':generateContent?key=' + apiKey;

  const prompt = buildOCRPrompt_();

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
      temperature: CONFIG.GEMINI.TEMPERATURE,
      maxOutputTokens: CONFIG.GEMINI.MAX_TOKENS,
      responseMimeType: 'application/json'
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  // リトライ付きAPI呼び出し（最大3回）
  const MAX_RETRIES = 3;
  let lastError = null;

  for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    try {
      const response = UrlFetchApp.fetch(url, options);
      const statusCode = response.getResponseCode();

      if (statusCode !== 200) {
        lastError = 'Gemini API HTTP ' + statusCode;
        console.warn('Gemini API Error (attempt ' + attempt + '/' + MAX_RETRIES + '): ' +
                     statusCode + ' - ' + response.getContentText().slice(0, 300));
        if (attempt < MAX_RETRIES) {
          Utilities.sleep(2000 * attempt);
          continue;
        }
        throw new Error('Gemini API エラー: ' + statusCode + '（' + MAX_RETRIES + '回リトライ後）');
      }

      const result = JSON.parse(response.getContentText());

      // レスポンス構造を検証
      const validation = validateGeminiResponse_(result);
      if (!validation.ok) {
        lastError = validation.reason;
        console.warn('Gemini response invalid (attempt ' + attempt + '/' + MAX_RETRIES + '): ' + validation.reason);
        if (attempt < MAX_RETRIES) {
          Utilities.sleep(2000 * attempt);
          continue;
        }
        // 最終試行でも無効 → エラー情報付きで返す
        return {
          date: '',
          storeName: 'API_ERROR',
          totalAmount: null,
          invoiceNumber: null,
          items: [],
          rawText: 'APIエラー: ' + validation.reason
        };
      }

      return parseGeminiResponse_(result);

    } catch (e) {
      lastError = e.message;
      console.warn('Gemini call exception (attempt ' + attempt + '/' + MAX_RETRIES + '): ' + e.message);
      if (attempt < MAX_RETRIES) {
        Utilities.sleep(2000 * attempt);
        continue;
      }
      throw e;
    }
  }

  throw new Error('Gemini API: 全リトライ失敗 - ' + lastError);
}

/**
 * Gemini APIレスポンスの構造を検証
 * @param {Object} result - APIレスポンス
 * @return {{ok: boolean, reason: string}}
 */
function validateGeminiResponse_(result) {
  // promptFeedbackによるブロック
  if (result.promptFeedback && result.promptFeedback.blockReason) {
    return { ok: false, reason: 'BLOCKED: ' + result.promptFeedback.blockReason };
  }

  // candidatesが空
  if (!result.candidates || result.candidates.length === 0) {
    return { ok: false, reason: 'NO_CANDIDATES: レスポンスが空です' };
  }

  const candidate = result.candidates[0];

  // finishReasonチェック
  if (candidate.finishReason && candidate.finishReason !== 'STOP') {
    return { ok: false, reason: 'FINISH_REASON: ' + candidate.finishReason };
  }

  // テキストコンテンツの存在チェック
  const text = candidate.content?.parts?.[0]?.text;
  if (!text || text.trim().length === 0) {
    return { ok: false, reason: 'EMPTY_RESPONSE: テキストが空です' };
  }

  return { ok: true, reason: '' };
}

/**
 * OCR用プロンプトを構築
 * @return {string}
 */
function buildOCRPrompt_() {
  // VALID_ACCOUNT_TITLES (Mapping.js) をプロンプトに埋め込む
  const accountTitlesList = VALID_ACCOUNT_TITLES.join('、');

  return `あなたはレシート・領収書のOCR専門家です。
画像から以下の情報を抽出し、JSONで返してください。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【最重要】この領収証は手書きですか？印刷ですか？
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

判定基準：
- 金額、日付、発行者名などが"手書き"で記入されている → 手書き
- レジで印刷されたレシート、プリンタで印刷された領収証 → 印刷
- 印鑑や署名"だけ"が手書き → 印刷（本体が印刷なら印刷扱い）

レスポンスのJSONに必ず以下を含めてください：
"isHandwritten": true（手書き）または false（印刷）

【重要ルール】
- 見たままの値を正確に抽出してください
- 計算や推測はしないでください
- 値が読み取れない場合はnullを返してください
- 税額や小計が記載されていれば必ず抽出してください

【日付の変換ルール - 重要】
- 和暦は西暦に変換してください: 令和N年 = (2018+N)年（例: 令和7年 = 2025年、R7 = 2025年）
- 2桁の年号は2000年代として解釈してください（例: 25年 = 2025年）
- 必ずYYYY-MM-DD形式で出力してください

【店舗名の正規化ルール - 重要】
storeNameは以下のルールで正規化し、集計しやすい「ブランド名」のみを出力すること。
1. 法人格を除去: 株式会社、有限会社、合同会社、(株)、(有)、Inc.、Co.、Ltd.、Corp.、LLC 等
2. 支店名・店舗番号を除去: 「○○店」「○○支店」「○○営業所」「No.123」等
3. 住所・電話番号を除去: 「大阪市北区…」「TEL 06-…」等
4. ブランド名が日本語ならそのまま日本語で、英語ならそのまま英語で出力（言語はレシート記載に従う）
5. チェーン店は一般的に知られているブランド名に統一

変換例:
- "株式会社セブン-イレブン・ジャパン 豊中上新田店" → "セブン-イレブン"
- "Starbucks Coffee 梅田茶屋町店" → "Starbucks"
- "ENEOS Dr.Driveセルフ 門真店" → "ENEOS"
- "タイムズ24 なんばパークス" → "タイムズ"
- "合同会社西友 荻窪店" → "西友"
- "割烹 花月" → "割烹 花月"（個人店はそのまま）

【抽出項目】
1. date: 日付（YYYY-MM-DD形式。和暦・2桁年は上記ルールで西暦4桁に変換）
2. storeName: 店舗名（上記の正規化ルールに従い、ブランド名のみ出力）
3. totalAmount: 支払総額（レジで支払った最終金額、「合計」「お支払い」等の金額）
4. invoiceNumber: インボイス登録番号（T+13桁、なければnull）
5. items: 明細行の配列
   - name: 品目名
   - quantity: 数量（不明なら1）
   - unitPrice: 単価（不明ならnull）
   - amount: 金額
   - taxMark: 軽減税率マーク（"※"や"*"など、なければnull）
6. subtotalInfo: 小計・税額情報（★重要：レシートに記載があれば必ず抽出）
   - subtotal10: 10%対象の税抜小計（「10%対象」「税抜」等）
   - tax10: 10%消費税額（「消費税(10%)」「内消費税」等）
   - subtotal8: 8%対象の税抜小計（「8%対象」「軽減税率対象」等）
   - tax8: 8%消費税額（「消費税(8%)」等）
   - dieselTax: 軽油税（「軽油税」「軽油引取税」、ガソリンスタンドのレシートで確認）
   - bathTax: 入湯税（温泉・旅館のレシートで確認）
   - accommodationTax: 宿泊税（ホテルのレシートで確認）
   - otherNonTaxable: その他の非課税・不課税金額
7. suggestedAccountTitle: 推定される勘定科目（下記の許可リストから1つだけ選択）
8. rawText: レシート全体のテキスト（デバッグ用、300文字まで）
9. currency: 通貨コード（"JPY", "USD", "EUR", "SGD", "THB"等。日本語のレシートで円表記ならJPY。$表記ならUSD。€表記ならEUR。通貨記号や表記から判定。不明ならJPY）

【勘定科目 - 許可リスト（厳守）】
以下のリストに含まれる勘定科目のみ使用してください。このリスト以外の単語（配送費、システム料、食費など）は絶対に使用禁止です。
許可リスト: ${accountTitlesList}

【勘定科目の推定ルール - 重要】

■ 基本方針
1. まず店名を見て、それで勘定科目が明確にわかればそれで判断
2. 店名だけでわからない場合（Amazon、楽天、個人名など）は、品名・明細を見て判断
3. 品名を見てもわからない場合は、suggestedAccountTitle を null にする

■ 店名だけでわかる例
- スターバックス → 会議費
- ENEOS → 車両費
- ダイソー → 消耗品費
- 高島屋 → 交際費

■ 店名だけでわからない例（品名を見て判断）
- Amazon + USBケーブル → 消耗品費
- Amazon + ビジネス書籍 → 新聞図書費
- Amazon + コーヒー豆 → 会議費（福利厚生費でも可）
- 楽天 + 事務用品 → 消耗品費

■ 個人名のみの領収書
- 肩書きあり（税理士、弁護士、司法書士など）→ 支払報酬料
- 肩書きなし → 飲食店と推定して会議費（または交際費）
  理由: 個人名のみの領収書は9割が小料理屋・スナック・居酒屋等

■ 判断の指針（上記で判定できない場合に参考）
- 飲食店・カフェ・コンビニでの飲食 → 会議費
- ゴルフ場・百貨店・贈答品・スナック・ラウンジ → 交際費
- 宅配便・配送 → 荷造運賃
- ガソリンスタンド・駐車場・ETC・高速道路 → 車両費
- 電車・バス・タクシー・航空券・宿泊 → 旅費交通費
- IT・クラウド・SaaS・決済手数料 → 支払手数料
- 人材派遣・業務委託 → 外注費
- 事務用品・日用品・物品購入 → 消耗品費
- 書籍・書店・新聞 → 新聞図書費
- 各種団体年会費 → 諸会費
- 税理士・弁護士等の報酬 → 支払報酬料

■ 出力ルール
- 確信がある場合: 許可リストから1つ選んで出力
- 確信がない場合: null を出力（後で人間が判断）
- 判断がつかない場合は null にする（消耗品費にしない）

【課税対象外の例】
- 軽油税：ガソリンスタンドで軽油を購入した場合に別記載される税金
- 入湯税：温泉旅館等で150円/人程度
- 宿泊税：ホテル等で100〜200円/人程度

【収入印紙について - 重要】
- 領収書に貼付されている収入印紙は完全に無視してください
- 収入印紙の金額（200円等）は抽出しないでください
- 収入印紙は店舗側の納税義務であり、支払金額とは無関係です

【出力形式】
以下のJSON形式で出力してください:
{
  "isHandwritten": false,
  "date": "2025-01-15",
  "storeName": "店舗名",
  "totalAmount": 1234,
  "invoiceNumber": "T1234567890123",
  "items": [
    {"name": "商品A", "quantity": 1, "unitPrice": 100, "amount": 100, "taxMark": null},
    {"name": "商品B", "quantity": 2, "unitPrice": 200, "amount": 400, "taxMark": "※"}
  ],
  "subtotalInfo": {
    "subtotal10": 1000,
    "tax10": 100,
    "subtotal8": 200,
    "tax8": 16,
    "dieselTax": null,
    "bathTax": null,
    "accommodationTax": null,
    "otherNonTaxable": null
  },
  "suggestedAccountTitle": "消耗品費",
  "rawText": "領収書のテキスト...",
  "currency": "JPY"
}`;
}

/**
 * Geminiのレスポンスをパース
 * @param {Object} apiResult
 * @return {OCRResult}
 */
function parseGeminiResponse_(apiResult) {
  try {
    const text = apiResult.candidates?.[0]?.content?.parts?.[0]?.text || '';

    // JSONを抽出（コードブロックを除去）
    let jsonStr = text.trim();
    if (jsonStr.startsWith('```json')) {
      jsonStr = jsonStr.replace(/^```json\s*/, '').replace(/\s*```$/, '');
    } else if (jsonStr.startsWith('```')) {
      jsonStr = jsonStr.replace(/^```\s*/, '').replace(/\s*```$/, '');
    }

    const parsed = JSON.parse(jsonStr);

    // 手書き判定のログ
    console.log('Gemini手書き判定: isHandwritten = ' + parsed.isHandwritten);

    // OCRResult形式に変換
    return {
      date: normalizeDate_(parsed.date || ''),
      storeName: (parsed.storeName || 'UNKNOWN').replace(/[\r\n]+/g, ' ').trim(),
      totalAmount: parseAmount(parsed.totalAmount),
      invoiceNumber: parsed.invoiceNumber || null,
      items: (parsed.items || []).map(item => ({
        name: item.name || '',
        quantity: item.quantity || 1,
        unitPrice: parseAmount(item.unitPrice),
        amount: parseAmount(item.amount) || 0,
        taxMark: item.taxMark || null
      })),
      rawText: parsed.rawText || '',
      // 手書き判定
      isHandwritten: parsed.isHandwritten === true,
      // 拡張情報（Logic_Accountingで使用）
      _subtotalInfo: parsed.subtotalInfo || null,
      // Geminiが推定した勘定科目
      _suggestedAccountTitle: parsed.suggestedAccountTitle || null,
      // 通貨コード
      currency: parsed.currency || 'JPY'
    };

  } catch (e) {
    const rawText = apiResult.candidates?.[0]?.content?.parts?.[0]?.text || '';
    const finishReason = apiResult.candidates?.[0]?.finishReason || 'N/A';
    console.error('Gemini Response Parse Error: ' + e.message);
    console.error('finishReason: ' + finishReason);
    console.error('Raw text (first 500 chars): ' + rawText.slice(0, 500));

    return {
      date: '',
      storeName: 'PARSE_ERROR',
      totalAmount: null,
      invoiceNumber: null,
      items: [],
      rawText: 'パースエラー: ' + e.message + ' | finishReason: ' + finishReason + ' | raw: ' + rawText.slice(0, 200)
    };
  }
}

/**
 * 日付文字列を正規化（和暦・2桁年・不正年の補正）
 * Geminiが変換に失敗した場合のフォールバック
 * @param {string} dateStr
 * @return {string} YYYY-MM-DD形式
 */
function normalizeDate_(dateStr) {
  if (!dateStr) return '';
  let str = String(dateStr).trim();

  // 全角数字を半角に変換
  str = str.replace(/[０-９]/g, function(s) {
    return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
  });

  // パターン1: 令和N年M月D日 / R.N.M.D / R N年M月D日
  const rewaMatch = str.match(/(?:令和|Ｒ|R)\.?\s*(\d{1,2})[年\.\/\-](\d{1,2})[月\.\/\-](\d{1,2})/);
  if (rewaMatch) {
    const year = 2018 + parseInt(rewaMatch[1]);
    const month = ('0' + rewaMatch[2]).slice(-2);
    const day = ('0' + rewaMatch[3]).slice(-2);
    return year + '-' + month + '-' + day;
  }

  // パターン2: YYYY-MM-DD（既に正しい形式）→ 年の妥当性チェック
  const isoMatch = str.match(/^(\d{2,4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
  if (isoMatch) {
    let year = parseInt(isoMatch[1]);
    const month = ('0' + isoMatch[2]).slice(-2);
    const day = ('0' + isoMatch[3]).slice(-2);

    // 2桁年の補正（00-99 → 2000-2099）
    if (year < 100) {
      year += 2000;
    }

    // 妥当性チェック（2000〜2099の範囲外は不正）
    if (year < 2000 || year > 2099) {
      console.warn('normalizeDate_: 不正な年を検出: ' + year + ' → 空文字を返します');
      return '';
    }

    return year + '-' + month + '-' + day;
  }

  // パターン3: N年M月D日（和暦の可能性がある2桁以下の年）
  const jpShortMatch = str.match(/^(\d{1,2})年(\d{1,2})月(\d{1,2})日/);
  if (jpShortMatch) {
    const rawYear = parseInt(jpShortMatch[1]);
    // 令和として解釈（1〜20の範囲なら和暦の可能性が高い）
    const year = rawYear <= 20 ? 2018 + rawYear : 2000 + rawYear;
    const month = ('0' + jpShortMatch[2]).slice(-2);
    const day = ('0' + jpShortMatch[3]).slice(-2);
    return year + '-' + month + '-' + day;
  }

  // パターン4: YYYY年M月D日（4桁年の日本語表記）
  const jpFullMatch = str.match(/(\d{4})年(\d{1,2})月(\d{1,2})日/);
  if (jpFullMatch) {
    const year = parseInt(jpFullMatch[1]);
    const month = ('0' + jpFullMatch[2]).slice(-2);
    const day = ('0' + jpFullMatch[3]).slice(-2);
    return year + '-' + month + '-' + day;
  }

  // パターン5: 最終フォールバック — 「令和」を含み数字が3つ以上ある場合
  // 上記パターン1で拾えなかった変則的な令和表記（特殊文字・不可視文字混入など）
  if (/令和/.test(str)) {
    const nums = str.match(/\d+/g);
    if (nums && nums.length >= 3) {
      const year = 2018 + parseInt(nums[0]);
      const month = parseInt(nums[1]);
      const day = parseInt(nums[2]);
      if (year >= 2019 && year <= 2099 && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
        return year + '-' + ('0' + month).slice(-2) + '-' + ('0' + day).slice(-2);
      }
    }
  }

  // パターン6: 最終フォールバック — 「R」+数字パターン（R3.4.16 等の変則表記）
  if (/^[RＲ]\d/.test(str)) {
    const nums = str.match(/\d+/g);
    if (nums && nums.length >= 3) {
      const year = 2018 + parseInt(nums[0]);
      const month = parseInt(nums[1]);
      const day = parseInt(nums[2]);
      if (year >= 2019 && year <= 2099 && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
        return year + '-' + ('0' + month).slice(-2) + '-' + ('0' + day).slice(-2);
      }
    }
  }

  // いずれにもマッチしない場合はそのまま返す
  // ★注意: determineStatus_でDATE_FORMAT_INVALIDとして検出される
  console.warn('normalizeDate_: 変換失敗: "' + str + '"');
  return str;
}

/**
 * Google Cloud Vision OCR（フォールバック用）
 * @param {GoogleAppsScript.Drive.File} file
 * @param {string} lang - 言語コード（'ja'等）
 * @return {string|null} - 抽出されたテキスト
 */
function ocrWithCloudVision_(file, lang) {
  // Cloud Vision APIを使用する場合の実装
  // 現時点ではGemini Flashで十分なため省略
  return null;
}
