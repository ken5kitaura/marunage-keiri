/**
 * _Types.gs
 * 型定義（JSDocによる型アノテーション）
 *
 * GASは静的型付けがないため、JSDocで型を定義し
 * エディタの補完・ドキュメント生成に活用する
 */

/**
 * @typedef {Object} OCRResult
 * @property {string} date - 日付（YYYY-MM-DD形式）
 * @property {string} storeName - 店舗名
 * @property {number|null} totalAmount - 総合計金額
 * @property {string|null} invoiceNumber - インボイス登録番号（T+13桁）
 * @property {Array<OCRLineItem>} items - 明細行
 * @property {string} rawText - 生テキスト（デバッグ用）
 */

/**
 * @typedef {Object} OCRLineItem
 * @property {string} name - 品目名
 * @property {number} quantity - 数量
 * @property {number} unitPrice - 単価
 * @property {number} amount - 金額
 * @property {string|null} taxMark - 軽減税率マーク（※, *など）
 */

/**
 * @typedef {Object} AccountingData
 * @property {number} subtotal10 - 10%対象税抜金額
 * @property {number} tax10 - 10%消費税額
 * @property {number} subtotal8 - 8%対象税抜金額
 * @property {number} tax8 - 8%消費税額
 * @property {number} rawNonTaxable - 不課税金額（入湯税・宿泊税等）
 * @property {number} totalAmount - 総合計
 * @property {boolean} isCompound - 複合仕訳フラグ（不課税がある場合true）
 * @property {boolean} hasInvoice - インボイス番号の有無
 * @property {string} taxType - 課税タイプ（'standard'|'reduced'|'mixed'|'exempt'）
 */

/**
 * @typedef {Object} ReceiptData
 * @property {string} id - 一意識別子（ファイル名ベース）
 * @property {string} status - 処理ステータス（'OK'|'CHECK'|'ERROR'）
 * @property {string} fileUrl - Google DriveファイルURL
 * @property {string} fileName - ファイル名
 * @property {Date} processedAt - 処理日時
 *
 * @property {OCRResult} ocr - OCR抽出結果
 * @property {AccountingData} accounting - 会計計算結果
 *
 * @property {string} date - 取引日（YYYY-MM-DD）
 * @property {string} storeName - 店舗名
 * @property {string} accountTitle - 勘定科目（借方）
 * @property {string} creditAccount - 貸方科目（現金/クレジットカード等）
 * @property {string} folderLabel - フォルダ区分
 *
 * @property {ReconcileResult|null} reconcile - クレカ突合結果
 * @property {string} debugInfo - デバッグ情報
 */

/**
 * @typedef {Object} ReconcileResult
 * @property {string} status - 突合ステータス（'MATCHED'|'MULTI'|'UNMATCHED'）
 * @property {string|null} statementId - 明細ID（タブ名_行番号）
 * @property {string|null} statementTab - 明細タブ名
 * @property {number|null} statementRow - 明細行番号
 * @property {Date|null} statementDate - 明細利用日
 * @property {string|null} statementMerchant - 明細利用先
 * @property {number|null} statementAmount - 明細金額
 * @property {number|null} score - 突合スコア
 * @property {Array<ReconcileCandidate>} candidates - 候補リスト
 */

/**
 * @typedef {Object} ReconcileCandidate
 * @property {string} tabName - 明細タブ名
 * @property {number} rowNumber - 明細行番号
 * @property {Date} date - 利用日
 * @property {string} merchant - 利用先
 * @property {number} amount - 金額
 * @property {number} score - スコア
 */

/**
 * @typedef {Object} YayoiRow
 * @property {string} date - 日付（YYYY/MM/DD）
 * @property {string} debitAccount - 借方勘定科目
 * @property {string} debitSubAccount - 借方補助科目
 * @property {number} debitAmount - 借方金額
 * @property {string} debitTaxType - 借方税区分
 * @property {string} creditAccount - 貸方勘定科目
 * @property {string} creditSubAccount - 貸方補助科目
 * @property {number} creditAmount - 貸方金額
 * @property {string} creditTaxType - 貸方税区分
 * @property {string} summary - 摘要
 */

/**
 * @typedef {Object} FolderConfig
 * @property {string} folderId - Google DriveフォルダID
 * @property {string} label - 表示ラベル（現金、クレカ等）
 * @property {string} creditAccount - デフォルト貸方科目
 */

/**
 * @typedef {Object} MappingRule
 * @property {string} keyword - マッチングキーワード
 * @property {string} accountTitle - 勘定科目
 * @property {string|null} subAccount - 補助科目
 */
