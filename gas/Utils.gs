/**
 * Utils.gs
 * 共通ユーティリティ関数
 */

/**
 * 金額文字列を数値に変換
 * カンマ・円記号・全角数字に対応
 * @param {*} val
 * @return {number|null}
 */
function parseAmount(val) {
  if (val == null || val === '') return null;
  if (typeof val === 'number') return val;

  const str = String(val)
    .replace(/[¥￥,，円\s　]/g, '')  // 記号除去
    .replace(/[０-９]/g, function(s) {  // 全角→半角
      return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
    });

  const num = parseFloat(str);
  return isNaN(num) ? null : num;
}

/**
 * 日付文字列をDateオブジェクトに変換
 * @param {string|Date} val
 * @return {Date|null}
 */
function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date) return val;

  const str = String(val).trim();

  // YYYY-MM-DD または YYYY/MM/DD
  const match1 = str.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (match1) {
    return new Date(parseInt(match1[1]), parseInt(match1[2]) - 1, parseInt(match1[3]));
  }

  // 令和X年M月D日
  const match2 = str.match(/令和(\d+)年(\d+)月(\d+)日/);
  if (match2) {
    const year = 2018 + parseInt(match2[1]);
    return new Date(year, parseInt(match2[2]) - 1, parseInt(match2[3]));
  }

  // 標準パース試行
  const parsed = new Date(str);
  return isNaN(parsed.getTime()) ? null : parsed;
}

/**
 * 日付をYYYY-MM-DD形式にフォーマット
 * @param {Date} date
 * @return {string}
 */
function formatDateISO(date) {
  if (!date || !(date instanceof Date)) return '';
  const y = date.getFullYear();
  const m = ('0' + (date.getMonth() + 1)).slice(-2);
  const d = ('0' + date.getDate()).slice(-2);
  return y + '-' + m + '-' + d;
}

/**
 * 日付をYYYY/MM/DD形式にフォーマット（弥生用）
 * @param {Date} date
 * @return {string}
 */
function formatDateYayoi(date) {
  if (!date || !(date instanceof Date)) return '';
  const y = date.getFullYear();
  const m = ('0' + (date.getMonth() + 1)).slice(-2);
  const d = ('0' + date.getDate()).slice(-2);
  return y + '/' + m + '/' + d;
}

/**
 * 2つの日付の差分日数を計算
 * @param {Date} d1
 * @param {Date} d2
 * @return {number}
 */
function dateDiffDays(d1, d2) {
  if (!d1 || !d2) return Infinity;
  const ms = Math.abs(d1.getTime() - d2.getTime());
  return Math.floor(ms / (1000 * 60 * 60 * 24));
}

/**
 * 文字列を正規化（全角→半角、大文字化、記号除去）
 * @param {string} str
 * @return {string}
 */
function normalizeString(str) {
  if (!str) return '';

  return String(str)
    .replace(/　/g, ' ')  // 全角空白
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) {
      return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
    })
    .toUpperCase()
    .replace(/[\s\-_\.・]/g, '')
    .replace(/株式会社|カブシキガイシャ|\(株\)|（株）|ＫＫ/g, '');
}

/**
 * 簡易的な文字列類似度スコア
 * @param {string} s1
 * @param {string} s2
 * @return {number} 0-100
 */
function similarityScore(s1, s2) {
  const str1 = normalizeString(s1);
  const str2 = normalizeString(s2);

  if (str1 === str2) return 100;
  if (!str1 || !str2) return 0;

  // 部分一致
  if (str1.indexOf(str2) !== -1 || str2.indexOf(str1) !== -1) {
    return 80;
  }

  // 文字一致率
  const maxLen = Math.max(str1.length, str2.length);
  let matches = 0;
  for (let i = 0; i < Math.min(str1.length, str2.length); i++) {
    if (str1[i] === str2[i]) matches++;
  }

  return Math.floor((matches / maxLen) * 100);
}

/**
 * インボイス登録番号のバリデーション
 * @param {string} invoiceNumber
 * @return {boolean}
 */
function isValidInvoiceNumber(invoiceNumber) {
  if (!invoiceNumber) return false;
  // T + 13桁の数字
  return /^T\d{13}$/.test(String(invoiceNumber).trim());
}

/**
 * ヘッダー行から列インデックスを取得
 * @param {Array} headerRow
 * @param {Array<string>} candidates
 * @return {number} 0-indexed, not found = -1
 */
function findHeaderIndex(headerRow, candidates) {
  if (!headerRow || !Array.isArray(headerRow)) return -1;

  for (let i = 0; i < headerRow.length; i++) {
    const cell = String(headerRow[i] || '').trim();
    for (const candidate of candidates) {
      if (cell === candidate) return i;
    }
  }
  return -1;
}

/**
 * 端数処理（四捨五入）
 * @param {number} value
 * @param {number} digits - 小数点以下桁数
 * @return {number}
 */
function roundTo(value, digits) {
  const factor = Math.pow(10, digits || 0);
  return Math.round(value * factor) / factor;
}

/**
 * 税込金額から税抜・税額を計算（内税計算）
 * @param {number} totalWithTax - 税込金額
 * @param {number} taxRate - 税率（10 or 8）
 * @return {{subtotal: number, tax: number}}
 */
function calculateFromTaxIncluded(totalWithTax, taxRate) {
  const rate = taxRate / 100;
  const subtotal = Math.round(totalWithTax / (1 + rate));
  const tax = totalWithTax - subtotal;
  return { subtotal: subtotal, tax: tax };
}

/**
 * 税抜金額から税込・税額を計算（外税計算）
 * @param {number} subtotal - 税抜金額
 * @param {number} taxRate - 税率（10 or 8）
 * @return {{total: number, tax: number}}
 */
function calculateFromTaxExcluded(subtotal, taxRate) {
  const rate = taxRate / 100;
  const tax = Math.round(subtotal * rate);
  const total = subtotal + tax;
  return { total: total, tax: tax };
}

/**
 * 非課税金額かどうかを判定
 * @param {string} itemName
 * @param {string} storeName
 * @return {boolean}
 */
function isNonTaxableItem(itemName, storeName) {
  const keywords = CONFIG.NON_TAXABLE_KEYWORDS;
  const combined = (itemName || '') + ' ' + (storeName || '');

  for (const keyword of keywords) {
    if (combined.indexOf(keyword) !== -1) {
      return true;
    }
  }
  return false;
}

/**
 * Shift-JISエンコード（弥生CSV用）
 * @param {string} text
 * @return {Blob}
 */
function encodeShiftJIS(text) {
  // GASのUtilities.newBlobでShift_JISを指定
  const blob = Utilities.newBlob('', 'text/csv', 'export.csv');

  // 実際のエンコードはSpreadsheet経由で行う必要あり
  // この関数は簡易実装
  return Utilities.newBlob(text, 'text/csv; charset=Shift_JIS', 'export.csv');
}

/**
 * CSVエスケープ
 * @param {string} value
 * @return {string}
 */
function escapeCSV(value) {
  if (value == null) return '';
  const str = String(value);
  if (str.includes(',') || str.includes('"') || str.includes('\n')) {
    return '"' + str.replace(/"/g, '""') + '"';
  }
  return str;
}

/**
 * 配列をCSV行に変換
 * @param {Array} row
 * @return {string}
 */
function toCSVRow(row) {
  return row.map(escapeCSV).join(',');
}
