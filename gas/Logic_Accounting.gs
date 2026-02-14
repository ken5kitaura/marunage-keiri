/**
 * Logic_Accounting.gs
 * 会計計算・勘定科目判定ロジック
 *
 * 責務:
 * - OCRResultからAccountingDataを生成
 * - 税額計算（内税）
 * - インボイス番号有無チェック
 * - 入湯税/宿泊税等の不課税検出と isCompound フラグ設定
 * - 店名から勘定科目を判定
 */

/**
 * OCRResultからAccountingDataを計算
 * @param {OCRResult} ocr
 * @return {AccountingData}
 */
function calculateAccountingData(ocr) {
  const totalAmount = parseAmount(ocr.totalAmount) || 0;

  // 1. 不課税額（入湯税・軽油税等）を抽出
  let rawNonTaxable = 0;
  if (ocr._subtotalInfo) {
    const info = ocr._subtotalInfo;
    const dieselTax = parseAmount(info.dieselTax) || 0;
    const bathTax = parseAmount(info.bathTax) || 0;
    const accommodationTax = parseAmount(info.accommodationTax) || 0;
    const otherNonTaxable = parseAmount(info.otherNonTaxable) || 0;
    rawNonTaxable = dieselTax + bathTax + accommodationTax + otherNonTaxable;
  }

  // 明細行から追加の不課税検出
  const nonTaxableFromItems = detectNonTaxableFromItems_(ocr.items);
  if (nonTaxableFromItems > 0 && rawNonTaxable === 0) {
    rawNonTaxable = nonTaxableFromItems;
  }

  // 2. 課税対象額を算出（総額から不課税を除く）
  const taxableAmount = totalAmount - rawNonTaxable;

  // 3. OCRから税額情報が取得できている場合はそれを優先
  let subtotal10 = 0;
  let tax10 = 0;
  let subtotal8 = 0;
  let tax8 = 0;

  if (ocr._subtotalInfo) {
    const info = ocr._subtotalInfo;
    const ocrSubtotal10 = parseAmount(info.subtotal10) || 0;
    const ocrTax10 = parseAmount(info.tax10) || 0;
    const ocrSubtotal8 = parseAmount(info.subtotal8) || 0;
    const ocrTax8 = parseAmount(info.tax8) || 0;

    // OCRの値が存在する場合は使用
    if (ocrSubtotal10 > 0 || ocrTax10 > 0 || ocrSubtotal8 > 0 || ocrTax8 > 0) {
      subtotal10 = ocrSubtotal10;
      tax10 = ocrTax10;
      subtotal8 = ocrSubtotal8;
      tax8 = ocrTax8;

      // ★パターンA修正: 税額のみで税抜がない場合、総額から逆算
      // 旧: Math.round(tax10 / 0.1) は丸め誤差で6〜10円ズレる
      // 新: totalAmount - 他の確定要素 で正確に算出
      if (tax10 > 0 && subtotal10 === 0) {
        const derived = taxableAmount - tax10 - subtotal8 - tax8;
        if (derived > 0) {
          subtotal10 = derived;
        } else {
          subtotal10 = Math.round(tax10 / 0.1);  // フォールバック
        }
      }
      if (tax8 > 0 && subtotal8 === 0) {
        const derived = taxableAmount - tax8 - subtotal10 - tax10;
        if (derived > 0) {
          subtotal8 = derived;
        } else {
          subtotal8 = Math.round(tax8 / 0.08);  // フォールバック
        }
      }

      // ★自動補正: OCRが「税込金額」を「税抜金額」として読み取った場合
      const calculatedTotal = subtotal10 + tax10 + subtotal8 + tax8 + rawNonTaxable;
      const totalTax = tax10 + tax8;
      if (totalAmount > 0 && calculatedTotal > 0) {
        const delta = calculatedTotal - totalAmount;
        if (totalTax > 0 && delta > 0 && delta >= totalTax * 0.9 && delta <= totalTax * 1.1) {
          if (subtotal10 > 0 && tax10 > 0) {
            subtotal10 = subtotal10 - tax10;
          }
          if (subtotal8 > 0 && tax8 > 0) {
            subtotal8 = subtotal8 - tax8;
          }
        }
      }

      // ★税率サニティチェック: OCRのsubtotalが明らかにおかしい場合、総額から再計算
      // 例: 軽油レシートで課税対象外¥1,065を10%対象額と誤認するケース
      if (subtotal10 > 0 && tax10 > 0) {
        const ratio10 = tax10 / subtotal10;
        if (ratio10 < 0.05 || ratio10 > 0.15) {
          // 税率が10%から大幅に外れている → subtotal10をtaxableAmountから再計算
          const derivedSubtotal10 = taxableAmount - tax10 - subtotal8 - tax8;
          if (derivedSubtotal10 > 0) {
            const derivedRatio10 = tax10 / derivedSubtotal10;
            if (derivedRatio10 >= 0.08 && derivedRatio10 <= 0.12) {
              subtotal10 = derivedSubtotal10;
            }
          }
        }
      }
      if (subtotal8 > 0 && tax8 > 0) {
        const ratio8 = tax8 / subtotal8;
        if (ratio8 < 0.03 || ratio8 > 0.13) {
          const derivedSubtotal8 = taxableAmount - tax8 - subtotal10 - tax10;
          if (derivedSubtotal8 > 0) {
            const derivedRatio8 = tax8 / derivedSubtotal8;
            if (derivedRatio8 >= 0.06 && derivedRatio8 <= 0.10) {
              subtotal8 = derivedSubtotal8;
            }
          }
        }
      }
    }
  }

  // 4. OCRに税情報がない場合は、課税対象額を10%内税として計算
  if (subtotal10 === 0 && tax10 === 0 && subtotal8 === 0 && tax8 === 0 && taxableAmount > 0) {
    tax10 = Math.floor(taxableAmount * 10 / 110);
    subtotal10 = taxableAmount - tax10;
  }

  // ★パターンB修正: 税抜金額があるのに税額が0の場合、総額から税額を逆算
  // 手書き領収証で「税抜金額37,800円 / 消費税等(10%)」のように税額が未記載のケース
  if (subtotal10 > 0 && tax10 === 0 && totalAmount > 0) {
    const derivedTax = totalAmount - subtotal10 - subtotal8 - tax8 - rawNonTaxable;
    if (derivedTax > 0) {
      const ratio = derivedTax / subtotal10;
      // 10%の妥当な範囲（8%〜12%）に収まるなら採用
      if (ratio >= 0.08 && ratio <= 0.12) {
        tax10 = derivedTax;
      }
    }
  }
  if (subtotal8 > 0 && tax8 === 0 && totalAmount > 0) {
    const derivedTax = totalAmount - subtotal8 - subtotal10 - tax10 - rawNonTaxable;
    if (derivedTax > 0) {
      const ratio = derivedTax / subtotal8;
      // 8%の妥当な範囲（6%〜10%）に収まるなら採用
      if (ratio >= 0.06 && ratio <= 0.10) {
        tax8 = derivedTax;
      }
    }
  }

  // ★パターンE修正: 軽油税等の不課税額が明らかにおかしい場合、総額から逆算
  // 例: OCRが単価@32.10を合計額と混同するケース
  if (rawNonTaxable > 0 && totalAmount > 0) {
    const calcTotal = subtotal10 + tax10 + subtotal8 + tax8 + rawNonTaxable;
    const gap = Math.abs(calcTotal - totalAmount);
    if (gap > 5) {
      // 不課税額を総額 - 課税分 で再計算
      const derivedNonTaxable = totalAmount - subtotal10 - tax10 - subtotal8 - tax8;
      if (derivedNonTaxable > 0 && derivedNonTaxable !== rawNonTaxable) {
        rawNonTaxable = derivedNonTaxable;
      }
    }
  }

  // 5. 値引き等による端数差異の補正
  // 総額を絶対正として、小さなミスマッチは税抜額を調整して整合させる
  // 税額は領収書記載値を保持（仕入税額控除に必要）
  let adjustmentNote = '';
  const finalCalcTotal = subtotal10 + tax10 + subtotal8 + tax8 + rawNonTaxable;
  if (totalAmount > 0 && finalCalcTotal > 0) {
    const finalDelta = finalCalcTotal - totalAmount;
    // 差異あり、かつ小額（総額の1%以内、最大100円）の場合のみ補正
    if (finalDelta !== 0 && Math.abs(finalDelta) <= Math.min(Math.max(Math.round(totalAmount * 0.01), 10), 100)) {
      if (subtotal10 > 0 && subtotal8 === 0) {
        subtotal10 = totalAmount - tax10 - rawNonTaxable;
        adjustmentNote = 'SUBTOTAL_ADJUSTED: 税抜額(10%)を' + (-finalDelta) + '円補正（値引等）';
      } else if (subtotal8 > 0 && subtotal10 === 0) {
        subtotal8 = totalAmount - tax8 - rawNonTaxable;
        adjustmentNote = 'SUBTOTAL_ADJUSTED: 税抜額(8%)を' + (-finalDelta) + '円補正（値引等）';
      } else if (subtotal10 > 0 && subtotal8 > 0) {
        // 混合税率: 10%側を調整（より大きい方を調整するのが一般的）
        subtotal10 = totalAmount - tax10 - subtotal8 - tax8 - rawNonTaxable;
        adjustmentNote = 'SUBTOTAL_ADJUSTED: 税抜額(10%)を' + (-finalDelta) + '円補正（値引等・混合税率）';
      }
    }
  }

  // 6. 税タイプ判定
  const taxType = determineTaxType_(subtotal10, subtotal8, rawNonTaxable);

  // 7. isCompound判定: 不課税がある場合はtrue
  const isCompound = rawNonTaxable > 0;

  return {
    subtotal10: subtotal10,
    tax10: tax10,
    subtotal8: subtotal8,
    tax8: tax8,
    rawNonTaxable: rawNonTaxable,
    totalAmount: totalAmount,
    isCompound: isCompound,
    hasInvoice: isValidInvoiceNumber(ocr.invoiceNumber),
    taxType: taxType,
    adjustmentNote: adjustmentNote
  };
}

/**
 * 明細行から非課税金額を検出
 * @param {Array<OCRLineItem>} items
 * @return {number}
 */
function detectNonTaxableFromItems_(items) {
  if (!items) return 0;

  let total = 0;
  for (const item of items) {
    if (isNonTaxableByItemName_(item.name)) {
      total += (item.amount || 0);
    }
  }
  return total;
}

/**
 * 品目名から非課税判定
 * @param {string} itemName
 * @return {boolean}
 */
function isNonTaxableByItemName_(itemName) {
  if (!itemName) return false;

  const keywords = CONFIG.NON_TAXABLE_KEYWORDS;
  const name = String(itemName).toLowerCase();

  for (const keyword of keywords) {
    if (name.indexOf(keyword.toLowerCase()) !== -1) {
      return true;
    }
  }
  return false;
}

/**
 * 税タイプを判定
 * @param {number} subtotal10
 * @param {number} subtotal8
 * @param {number} nonTaxable
 * @return {string}
 */
function determineTaxType_(subtotal10, subtotal8, nonTaxable) {
  const has10 = subtotal10 > 0;
  const has8 = subtotal8 > 0;
  const hasNonTax = nonTaxable > 0;

  if ((has10 && has8) || (has10 && hasNonTax) || (has8 && hasNonTax)) {
    return 'mixed';
  } else if (has8 && !has10) {
    return 'reduced';
  } else if (hasNonTax && !has10 && !has8) {
    return 'exempt';
  } else {
    return 'standard';
  }
}

// ============================================================
// 勘定科目判定
// ============================================================

/**
 * 店舗名と明細から勘定科目を推定
 * 優先順位: 軽油税 → Config_Mapping → STORE_ACCOUNT_MAP → 特殊ルール → Gemini推定 → VEHICLE_PATTERNS → null
 *
 * ★ 軽油税がある場合は100%ガソリンスタンドでの給油なので、無条件で車両費。
 * ★ Config_Mapping（税理士・ユーザー定義）が次に優先。
 *   ここに登録された店名→科目ルールは、汎用辞書やAI推定より常に優先される。
 *
 * @param {string} storeName
 * @param {Array<OCRLineItem>} items
 * @param {string} [geminiSuggestion] - Geminiが推定した勘定科目
 * @param {AccountingData} [accountingData] - 会計データ（軽油税判定用）
 * @return {string|null}
 */
function inferAccountTitleFromStore(storeName, items, geminiSuggestion, accountingData) {
  if (!storeName && (!items || items.length === 0) && !geminiSuggestion) {
    return null;
  }

  // ── 最優先: 軽油税（dieselTax）+ ガソリンスタンド判定 ──
  // 軽油税がありかつ店名がガソリンスタンド系 → 車両費確定
  // ※不課税列には入湯税・宿泊税も含まれるため、店名チェックで誤判定を防ぐ
  if (accountingData) {
    const dieselTax = (accountingData._subtotalInfo && accountingData._subtotalInfo.dieselTax)
      ? parseAmount(accountingData._subtotalInfo.dieselTax)
      : (accountingData.dieselTax ? parseAmount(accountingData.dieselTax) : 0);
    if (dieselTax > 0 && isDieselTaxStore_(storeName)) {
      console.log('軽油税検出 (' + dieselTax + '円) + GS店名 → 車両費確定: ' + storeName);
      return '車両費';
    }
  }

  const storeStr = String(storeName || '').toLowerCase();
  const itemsStr = (items || []).map(i => i.name || '').join(' ').toLowerCase();
  const originalStoreName = String(storeName || '');

  // ── 優先度1: Config_Mappingシート（税理士・ユーザー定義）──
  // 各社固有の仕訳ルールを最優先で適用
  const rules = loadMappingRules_();
  for (const rule of rules) {
    const keyword = rule.keyword.toLowerCase();
    if (storeStr.indexOf(keyword) !== -1 || itemsStr.indexOf(keyword) !== -1) {
      return validateAccountTitle_(rule.accountTitle);
    }
  }

  // ── 優先度2: STORE_ACCOUNT_MAP（Mapping.js 汎用辞書）──
  const mapResult = matchStoreAccountMap_(storeStr, itemsStr);
  if (mapResult) {
    return mapResult;
  }

  // ── 優先度3: 特殊ルール ──
  // ガソリンスタンドでの飲料は会議費（車両費より優先すべき例外）
  if (isGasStation_(storeStr)) {
    if (hasBeverageItem_(itemsStr)) {
      return '会議費';
    }
    return '車両費';
  }

  // 非課税・租税公課判定
  if (isNonTaxableStore_(storeStr, itemsStr)) {
    if (/印紙|市役所|区役所|法務局|証明書/.test(storeStr + itemsStr)) {
      return '租税公課';
    }
    if (/切手/.test(storeStr + itemsStr)) {
      return '通信費';
    }
    if (/振込|atm|手数料/.test(storeStr + itemsStr)) {
      return '支払手数料';
    }
  }

  // ── 優先度4: Gemini AIの推定値（許容リストチェック付き）──
  // レシート画像を実際に見て判断しているため信頼度が高い
  // VALID_ACCOUNT_TITLESに含まれるもののみ採用
  // さらに、店名カテゴリの許容リストで「明らかにおかしい」を弾く
  if (geminiSuggestion) {
    const validated = validateAccountTitle_(geminiSuggestion);
    if (validated) {
      // 許容リストチェック
      const allowCheck = checkAllowedAccountTitle_(originalStoreName, validated);
      if (allowCheck.allowed) {
        return validated;
      } else {
        // 許容リスト外 → 警告ログを出して、許容リストのデフォルト（先頭）を採用
        console.warn('許容リスト外の勘定科目: "' + validated + '" (店名: ' + originalStoreName + ', カテゴリ: ' + allowCheck.category + ') → ' + allowCheck.fallback + 'に変更');
        return allowCheck.fallback;
      }
    }
  }

  // ── 優先度5: 正規表現パターン（保守的・最終手段）──
  // VEHICLE_PATTERNS: 末尾SS等、キーワードでは拾えないパターンのみ
  if (matchVehiclePatterns_(originalStoreName)) {
    return '車両費';
  }

  // ── 該当なし ──
  return null;
}

/**
 * STORE_ACCOUNT_MAP から部分一致で勘定科目を検索
 * @param {string} storeStr - 小文字化済みの店舗名
 * @param {string} itemsStr - 小文字化済みの品目文字列
 * @return {string|null}
 */
function matchStoreAccountMap_(storeStr, itemsStr) {
  const combined = storeStr + ' ' + itemsStr;
  // 長いキーワードを先にマッチさせる（"Amazon Web Services" → "Amazon" の順）
  const keys = Object.keys(STORE_ACCOUNT_MAP).sort(function(a, b) {
    return b.length - a.length;
  });

  for (const key of keys) {
    if (combined.indexOf(key.toLowerCase()) !== -1) {
      return STORE_ACCOUNT_MAP[key];
    }
  }
  return null;
}

/**
 * VEHICLE_PATTERNS（Mapping.js）で車両費パターンを検出
 * キーワード部分一致では拾えない店名を正規表現で補完する
 * 例: "百舌鳥SS"(末尾SS), "P鶴見緑地"(先頭P), "信貴生駒スカイライン"(末尾ライン)
 * @param {string} storeName - 元の店舗名（大文字小文字そのまま）
 * @return {boolean}
 */
function matchVehiclePatterns_(storeName) {
  if (!storeName || typeof VEHICLE_PATTERNS === 'undefined') return false;

  for (const pattern of VEHICLE_PATTERNS) {
    if (pattern.test(storeName)) {
      return true;
    }
  }
  return false;
}

/**
 * 勘定科目がVALID_ACCOUNT_TITLESに含まれるか検証
 * 含まれなければnullを返す（不正な科目名の使用を防止）
 * @param {string} title - 検証する勘定科目
 * @return {string|null}
 */
function validateAccountTitle_(title) {
  if (!title) return null;
  const trimmed = String(title).trim();
  if (VALID_ACCOUNT_TITLES.indexOf(trimmed) !== -1) {
    return trimmed;
  }
  console.warn('無効な勘定科目を検出: "' + trimmed + '" → 無視します');
  return null;
}

/**
 * 軽油税が発生しうる店舗か判定（軽油税+店名の二重チェック用）
 * isGasStation_より広めのキーワードで判定する
 * @param {string} storeName - 元の店舗名（大文字小文字そのまま）
 * @return {boolean}
 */
function isDieselTaxStore_(storeName) {
  const s = String(storeName || '').toLowerCase();
  return /eneos|出光|コスモ|shell|石油|ガソリン|gs|costco|コストコ|給油|軽油|スタンド/.test(s);
}

/**
 * ガソリンスタンドか判定
 * @param {string} storeStr
 * @return {boolean}
 */
function isGasStation_(storeStr) {
  return /eneos|shell|出光|コスモ|石油|ガソリン|gs|スタンド/.test(storeStr);
}

/**
 * 飲料品目があるか判定
 * @param {string} itemsStr
 * @return {boolean}
 */
function hasBeverageItem_(itemsStr) {
  return /水|茶|コーヒー|ドリンク|coffee|tea|ジュース|缶/.test(itemsStr);
}

/**
 * 非課税取引店舗か判定
 * @param {string} storeStr
 * @param {string} itemsStr
 * @return {boolean}
 */
function isNonTaxableStore_(storeStr, itemsStr) {
  const combined = storeStr + ' ' + itemsStr;
  return /非課税|印紙|切手|手数料|atm|振込|市役所|行政|郵便局/.test(combined);
}

// ============================================================
// 許容リストチェック
// ============================================================

/**
 * Gemini推定の勘定科目が許容リスト内かチェック
 * ACCOUNT_TITLE_ALLOWED_MAP（Mapping.js）を使用
 *
 * @param {string} storeName - 店舗名
 * @param {string} accountTitle - チェックする勘定科目
 * @return {{allowed: boolean, category: string|null, fallback: string|null}}
 */
function checkAllowedAccountTitle_(storeName, accountTitle) {
  // ACCOUNT_TITLE_ALLOWED_MAP が未定義の場合はチェックスキップ（常に許可）
  if (typeof ACCOUNT_TITLE_ALLOWED_MAP === 'undefined' || typeof getStoreCategory_ !== 'function') {
    return { allowed: true, category: null, fallback: null };
  }

  const category = getStoreCategory_(storeName);

  // カテゴリが判定できない場合はチェックスキップ（許可）
  if (!category) {
    return { allowed: true, category: null, fallback: null };
  }

  const allowedList = ACCOUNT_TITLE_ALLOWED_MAP[category];

  // カテゴリに対応する許容リストがない場合はスキップ
  if (!allowedList || allowedList.length === 0) {
    return { allowed: true, category: category, fallback: null };
  }

  // 許容リスト内にあるか確認
  if (allowedList.indexOf(accountTitle) !== -1) {
    return { allowed: true, category: category, fallback: null };
  }

  // 許容リスト外 → フォールバック（リストの先頭）
  return {
    allowed: false,
    category: category,
    fallback: allowedList[0]
  };
}

// ============================================================
// 整合性チェック（validateAccountingDataは_Main.gsで使用）
// ============================================================

/**
 * AccountingDataの整合性をチェック
 * @param {AccountingData} acc
 * @param {number} expectedTotal - 期待される総額
 * @return {{isValid: boolean, errors: Array<string>}}
 */
function validateAccountingData(acc, expectedTotal) {
  const errors = [];

  // 計算された総額
  const calculatedTotal = acc.subtotal10 + acc.tax10 +
                          acc.subtotal8 + acc.tax8 +
                          acc.rawNonTaxable;

  // 総額との差異チェック
  // COMPOUND（不課税あり）: 全額明記のため完全一致を要求
  // 通常: 内税逆算の丸め誤差があるため±3円まで許容
  const delta = Math.abs(calculatedTotal - (expectedTotal || acc.totalAmount));
  const tolerance = acc.isCompound ? 0 : 3;
  if (delta > tolerance) {
    errors.push('TOTAL_MISMATCH: 差異=' + delta + '円');
  }

  // 負の値チェック
  if (acc.subtotal10 < 0 || acc.tax10 < 0 ||
      acc.subtotal8 < 0 || acc.tax8 < 0 ||
      acc.rawNonTaxable < 0) {
    errors.push('NEGATIVE_VALUE');
  }

  // 税額なしで税抜がある場合
  if (acc.subtotal10 > 0 && acc.tax10 === 0) {
    errors.push('TAX10_MISSING');
  }
  if (acc.subtotal8 > 0 && acc.tax8 === 0) {
    errors.push('TAX8_MISSING');
  }

  // 税率整合性チェック
  if (acc.subtotal10 > 0) {
    const ratio10 = acc.tax10 / acc.subtotal10;
    if (ratio10 < 0.09 || ratio10 > 0.11) {
      errors.push('TAX10_RATE_MISMATCH: ' + (ratio10 * 100).toFixed(1) + '%');
    }
  }
  if (acc.subtotal8 > 0) {
    const ratio8 = acc.tax8 / acc.subtotal8;
    if (ratio8 < 0.07 || ratio8 > 0.09) {
      errors.push('TAX8_RATE_MISMATCH: ' + (ratio8 * 100).toFixed(1) + '%');
    }
  }

  return {
    isValid: errors.length === 0,
    errors: errors
  };
}
