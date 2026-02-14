/**
 * Config.gs
 * 設定・定数管理
 *
 * フォルダIDやAPIキーはScriptPropertiesで管理し、
 * コードへの直書きを避ける
 */

const CONFIG = {
  // シート名
  SHEET_NAME: {
    MAIN: '本番シート',
    MAPPING: 'Config_Mapping',
    FOLDERS: 'Config_Folders'
  },

  // Gemini API
  GEMINI: {
    MODEL: 'gemini-2.0-flash',
    MAX_TOKENS: 4096,
    TEMPERATURE: 0.1  // 低温でより確定的な出力
  },

  // 処理設定
  PROCESSING: {
    MAX_EXECUTION_TIME_MS: 4.5 * 60 * 1000,  // 4分30秒
    RETRY_DELAY_MINUTES: 1
  },

  // 突合設定
  RECONCILE: {
    DEFAULT_DATE_WINDOW_DAYS: 2,
    MIN_MATCH_SCORE: 60,
    MERCHANT_GROUPS: [
      { key: 'starbucks', keywords: ['スターバックス', 'ｽﾀｰﾊﾞｯｸｽ', 'STARBUCKS', 'SBX'], dateWindowDays: 2 },
      { key: 'shinkansen', keywords: ['SMARTEX', 'EX予約', 'ＥＸ予約', 'JR東海', 'JR西日本', '新幹線', '東海道新幹線'], dateWindowDays: 5 },
      { key: 'toll', keywords: ['NEXCO', '阪神高速', '首都高', '高速道路', 'ETC'], dateWindowDays: 7 }
    ]
  },

  // 税率
  TAX_RATE: {
    STANDARD: 10,
    REDUCED: 8
  },

  // 非課税キーワード（入湯税・宿泊税等）
  NON_TAXABLE_KEYWORDS: ['入湯税', '宿泊税', '湯税', '滞在税', '観光税'],

  // 弥生CSV設定
  YAYOI: {
    ENCODING: 'Shift_JIS',
    DATE_FORMAT: 'yyyy/MM/dd'
  }
};

/**
 * 設定値を取得（Configシート → ScriptProperties の順で探索）
 * @param {string} key - プロパティキー
 * @param {string} defaultValue - デフォルト値
 * @return {string}
 */
function getConfig_(key, defaultValue) {
  // 1. Configシートから取得を試みる
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Config');
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
        for (const row of data) {
          if (String(row[0]).trim() === key) {
            const value = String(row[1] || '').trim();
            if (value) return value;
          }
        }
      }
    }
  } catch (e) {
    // シート取得に失敗した場合は無視してScriptPropertiesを使用
  }

  // 2. ScriptPropertiesから取得
  const props = PropertiesService.getScriptProperties();
  return props.getProperty(key) || defaultValue;
}

/**
 * ScriptPropertiesに設定を保存
 * @param {string} key - プロパティキー
 * @param {string} value - 値
 */
function setConfig_(key, value) {
  PropertiesService.getScriptProperties().setProperty(key, value);
}

/**
 * Gemini APIキーを取得（ScriptPropertiesから直接読み込み）
 * ライブラリ化時にConfigシート経由でのキー流出を防ぐため、
 * getConfig_()を経由せずScriptPropertiesのみを参照する
 * @return {string}
 */
function getGeminiApiKey_() {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '';
  if (!key) {
    throw new Error('GEMINI_API_KEY が設定されていません。\n' +
      '「設定」→「Gemini APIキーを設定」から入力するか、\n' +
      'スクリプトプロパティに直接設定してください。');
  }
  return key;
}

/**
 * クレカ明細スプレッドシートIDを取得
 * @return {string}
 */
function getCCStatementSpreadsheetId_() {
  return getConfig_('CC_STATEMENT_SPREADSHEET_ID', '');
}

/**
 * 処理済みファイル移動先フォルダIDを取得
 * @return {string} フォルダID（空文字なら移動しない）
 */
function getProcessedFolderId_() {
  return getConfig_('FOLDER_ID_PROCESSED', '');
}

/**
 * Config_FoldersシートまたはConfigシートからフォルダ設定を読み込む
 * @return {Array<FolderConfig>}
 */
function loadFolderConfigs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME.FOLDERS);

  if (!sheet) {
    // Config_Foldersシートがない場合、Configシートから取得を試みる
    const folderId = getConfig_('FOLDER_ID_RECEIPTS', '');
    if (folderId) {
      return [{
        folderId: folderId,
        label: '現金',
        creditAccount: '現金'
      }];
    }

    // 後方互換: DEFAULT_FOLDER_ID も確認
    const defaultFolderId = getConfig_('DEFAULT_FOLDER_ID', '');
    if (defaultFolderId) {
      return [{
        folderId: defaultFolderId,
        label: '現金',
        creditAccount: '現金'
      }];
    }

    throw new Error('フォルダIDが設定されていません。ConfigシートにFOLDER_ID_RECEIPTSを設定するか、Config_Foldersシートを作成してください。');
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const configs = [];

  data.forEach(function(row) {
    const folderId = String(row[0] || '').trim();
    const label = String(row[1] || '').trim();
    const creditAccount = String(row[2] || '現金').trim();

    if (folderId) {
      configs.push({
        folderId: folderId,
        label: label || 'デフォルト',
        creditAccount: creditAccount
      });
    }
  });

  return configs;
}

/**
 * Config_Mappingシートからマッピングルールを読み込む
 * @return {Array<MappingRule>}
 */
function loadMappingRules_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME.MAPPING);

  // Config_Mappingシートが無い場合は空配列を返す
  // （汎用辞書 STORE_ACCOUNT_MAP が代わりに処理する）
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const rules = [];

  data.forEach(function(row) {
    const keyword = String(row[0] || '').trim();
    const accountTitle = String(row[1] || '').trim();
    const subAccount = String(row[2] || '').trim();

    if (keyword && accountTitle) {
      rules.push({
        keyword: keyword,
        accountTitle: accountTitle,
        subAccount: subAccount || null
      });
    }
  });

  return rules;
}

/**
 * デフォルトのマッピングルールを返す
 * @return {Array<MappingRule>}
 */
function getDefaultMappingRules_() {
  return [
    // コンビニ
    { keyword: 'セブン', accountTitle: '消耗品費', subAccount: null },
    { keyword: 'ローソン', accountTitle: '消耗品費', subAccount: null },
    { keyword: 'ファミリーマート', accountTitle: '消耗品費', subAccount: null },
    { keyword: 'ファミマ', accountTitle: '消耗品費', subAccount: null },

    // 交通
    { keyword: 'JR', accountTitle: '旅費交通費', subAccount: null },
    { keyword: '鉄道', accountTitle: '旅費交通費', subAccount: null },
    { keyword: '電鉄', accountTitle: '旅費交通費', subAccount: null },
    { keyword: '新幹線', accountTitle: '旅費交通費', subAccount: null },
    { keyword: 'バス', accountTitle: '旅費交通費', subAccount: null },
    { keyword: 'タクシー', accountTitle: '旅費交通費', subAccount: null },
    { keyword: '駐車場', accountTitle: '旅費交通費', subAccount: null },

    // EC
    { keyword: 'Amazon', accountTitle: '消耗品費', subAccount: null },
    { keyword: '楽天', accountTitle: '消耗品費', subAccount: null },
    { keyword: 'ヨドバシ', accountTitle: '消耗品費', subAccount: null },

    // ガソリン
    { keyword: 'ENEOS', accountTitle: '車両費', subAccount: null },
    { keyword: '出光', accountTitle: '車両費', subAccount: null },
    { keyword: 'コスモ石油', accountTitle: '車両費', subAccount: null },
    { keyword: 'ガソリン', accountTitle: '車両費', subAccount: null },
    { keyword: '軽油', accountTitle: '車両費', subAccount: null },

    // 通信
    { keyword: 'ドコモ', accountTitle: '通信費', subAccount: null },
    { keyword: 'au', accountTitle: '通信費', subAccount: null },
    { keyword: 'ソフトバンク', accountTitle: '通信費', subAccount: null },
    { keyword: 'NTT', accountTitle: '通信費', subAccount: null },

    // 非課税・租税
    { keyword: '印紙', accountTitle: '租税公課', subAccount: null },
    { keyword: '市役所', accountTitle: '租税公課', subAccount: null },
    { keyword: '切手', accountTitle: '通信費', subAccount: null },
    { keyword: '振込', accountTitle: '支払手数料', subAccount: null },
    { keyword: 'ATM', accountTitle: '支払手数料', subAccount: null }
  ];
}

/**
 * Config_Foldersシートを作成
 */
function createConfigFoldersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME.FOLDERS);

  if (sheet) {
    SpreadsheetApp.getUi().alert('Config_Foldersシートは既に存在します。');
    return;
  }

  sheet = ss.insertSheet(CONFIG.SHEET_NAME.FOLDERS);

  // ヘッダー
  sheet.appendRow(['フォルダID', 'ラベル', '貸方科目']);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold');

  // サンプル行
  sheet.appendRow(['（ここにフォルダIDを入力）', '現金', '現金']);
  sheet.appendRow(['（クレカ用フォルダID）', 'クレカ', 'クレジットカード']);

  // 列幅調整
  sheet.setColumnWidth(1, 350);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 150);

  SpreadsheetApp.getUi().alert('Config_Foldersシートを作成しました。フォルダIDを入力してください。');
}

/**
 * Config_Mappingシートを作成
 */
function createConfigMappingSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME.MAPPING);

  if (sheet) {
    SpreadsheetApp.getUi().alert('Config_Mappingシートは既に存在します。');
    return;
  }

  sheet = ss.insertSheet(CONFIG.SHEET_NAME.MAPPING);

  // ヘッダー
  sheet.appendRow(['キーワード', '勘定科目', '補助科目']);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold');

  // デフォルトルールを挿入
  const defaultRules = getDefaultMappingRules_();
  defaultRules.forEach(function(rule) {
    sheet.appendRow([rule.keyword, rule.accountTitle, rule.subAccount || '']);
  });

  // 列幅調整
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 150);

  SpreadsheetApp.getUi().alert('Config_Mappingシートを作成しました。');
}
