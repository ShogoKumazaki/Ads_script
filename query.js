// 定数定義
const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';
const TIMEZONE_OFFSET = 14; // タイムゾーンのオフセット（時間）

// 取得期間の定義
const PERIODS = [
  { name: 'クエリ_昨日', days: 1 },
  { name: 'クエリ_過去7日', days: 7 },
  { name: 'クエリ_過去30日', days: 30 },
  { name: 'クエリ_過去90日', days: 90 }
];

function main() {
  // 現在の日付を取得
  const today = new Date();
  today.setHours(today.getHours() + TIMEZONE_OFFSET);
  const endDate = Utilities.formatDate(
    new Date(today.getFullYear(), today.getMonth(), today.getDate() + 1),
    "GMT+9",
    "yyyy-MM-dd"
  );

  // 各期間のレポートを取得
  PERIODS.forEach(period => {
    const startDate = Utilities.formatDate(
      new Date(today.getFullYear(), today.getMonth(), today.getDate() - period.days),
      "GMT+9",
      "yyyy-MM-dd"
    );
    console.log(`${period.name}の期間設定: ${startDate} から ${endDate}`);
    exportReport(period.name, startDate, endDate);
  });
}

/**
 * レポートをエクスポートする
 * @param {string} sheetName - シート名
 * @param {string} startDate - 開始日
 * @param {string} endDate - 終了日
 */
function exportReport(sheetName, startDate, endDate) {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(sheetName);

  // シートが存在しない場合は作成
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    console.log(`新しいシート '${sheetName}' を作成しました。`);
  }

  // レポートクエリの構築
  const query = buildReportQuery(startDate, endDate);
  console.log(`実行するクエリ: ${query}`);

  // レポートの取得とエクスポート
  const report = AdsApp.report(query);
  const range = sheet.getDataRange();
  range.clearContent();
  report.exportToSheet(sheet);

  // データの加工
  processReportData(sheet);
}

/**
 * レポートクエリを構築する
 * @param {string} startDate - 開始日
 * @param {string} endDate - 終了日
 * @returns {string} 構築されたクエリ
 */
function buildReportQuery(startDate, endDate) {
  const selectClause = `SELECT 
    search_term_view.search_term, 
    search_term_view.status, 
    segments.keyword.info.text, 
    segments.keyword.info.match_type, 
    campaign.name, 
    ad_group.name,  
    metrics.cost_micros,  
    metrics.impressions,  
    metrics.clicks, 
    metrics.ctr, 
    metrics.average_cpc,
    metrics.conversions,
    metrics.conversions_from_interactions_rate, 
    metrics.cost_per_conversion`;

  const fromClause = " FROM search_term_view";
  const whereClause = ` WHERE segments.date BETWEEN '${startDate}' AND '${endDate}'`;
  const orderByClause = " ORDER BY metrics.conversions DESC";

  return selectClause + fromClause + whereClause + orderByClause;
}

/**
 * レポートデータを加工する
 * @param {Sheet} sheet - スプレッドシートのシートオブジェクト
 */
function processReportData(sheet) {
  // ヘッダーの設定
  const headers = [
    "検索クエリ", 
    "ステータス", 
    "キーワード", 
    "マッチタイプ", 
    "キャンペーン名", 
    "広告グループ名", 
    "コスト",
    "インプ",  
    "クリック数", 
    "CTR", 
    "CPC", 
    "CV",
    "CVR",
    "CPA"
  ];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return; // データが存在しない場合は処理を終了

  // ステータスの変換
  convertStatus(sheet, lastRow);
  
  // マッチタイプの変換
  convertMatchType(sheet, lastRow);
  
  // 費用関連の変換
  convertCostMetrics(sheet, lastRow);
  
  // 数値フォーマットの設定
  formatNumberColumns(sheet, lastRow);
}

/**
 * ステータスを日本語に変換する
 * @param {Sheet} sheet - スプレッドシートのシートオブジェクト
 * @param {number} lastRow - 最終行
 */
function convertStatus(sheet, lastRow) {
  const statusRange = sheet.getRange(2, 2, lastRow - 1, 1);
  const statusValues = statusRange.getValues();
  const translatedStatuses = statusValues.map(row => {
    const status = row[0];
    switch (status) {
      case 'NONE':
        return ['なし'];
      case 'ADDED':
        return ['追加済み'];
      default:
        return [status];
    }
  });
  statusRange.setValues(translatedStatuses);
}

/**
 * マッチタイプを日本語に変換する
 * @param {Sheet} sheet - スプレッドシートのシートオブジェクト
 * @param {number} lastRow - 最終行
 */
function convertMatchType(sheet, lastRow) {
  const matchTypeRange = sheet.getRange(2, 4, lastRow - 1, 1);
  const matchTypeValues = matchTypeRange.getValues();
  const translatedMatchTypes = matchTypeValues.map(row => {
    const matchType = row[0];
    switch (matchType) {
      case 'EXACT':
        return ['完全一致'];
      case 'PHRASE':
        return ['フレーズ一致'];
      case 'BROAD':
        return ['インテントマッチ'];
      default:
        return [matchType];
    }
  });
  matchTypeRange.setValues(translatedMatchTypes);
}

/**
 * 費用関連のメトリクスを日本円に変換する
 * @param {Sheet} sheet - スプレッドシートのシートオブジェクト
 * @param {number} lastRow - 最終行
 */
function convertCostMetrics(sheet, lastRow) {
  // コストの変換
  convertCostToYen(sheet, lastRow, 7);
  
  // CPCの変換
  convertCostToYen(sheet, lastRow, 11);
  
  // CPAの変換
  convertCostToYen(sheet, lastRow, 14);
}

/**
 * 費用を日本円に変換する
 * @param {Sheet} sheet - スプレッドシートのシートオブジェクト
 * @param {number} lastRow - 最終行
 * @param {number} column - 変換対象の列番号
 */
function convertCostToYen(sheet, lastRow, column) {
  const costRange = sheet.getRange(2, column, lastRow - 1, 1);
  const costValues = costRange.getValues();
  const convertedCostValues = costValues.map(row => [Math.round(row[0] / 1000000)]);
  costRange.setValues(convertedCostValues);
}

/**
 * 数値列のフォーマット設定
 * @param {Sheet} sheet - スプレッドシートのシートオブジェクト
 * @param {number} lastRow - 最終行
 */
function formatNumberColumns(sheet, lastRow) {
  // インプレッション、クリック数をカンマ区切り数値フォーマット
  sheet.getRange(2, 8, lastRow - 1, 1).setNumberFormat('#,##0'); // インプレッション
  sheet.getRange(2, 9, lastRow - 1, 1).setNumberFormat('#,##0'); // クリック数
  sheet.getRange(2, 12, lastRow - 1, 1).setNumberFormat('#,##0.00'); // CV数

  // コスト、CPC、CPAを通貨フォーマット
  sheet.getRange(2, 7, lastRow - 1, 1).setNumberFormat('¥#,##0'); // コスト
  sheet.getRange(2, 11, lastRow - 1, 1).setNumberFormat('¥#,##0'); // CPC
  sheet.getRange(2, 14, lastRow - 1, 1).setNumberFormat('¥#,##0'); // CPA

  // CTR、CVRをパーセントフォーマット
  sheet.getRange(2, 10, lastRow - 1, 1).setNumberFormat('0.00%'); // CTR
  sheet.getRange(2, 13, lastRow - 1, 1).setNumberFormat('0.00%'); // CVR
}
  
  