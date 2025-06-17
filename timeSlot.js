// 定数定義
const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';
const TIMEZONE_OFFSET = 14; // タイムゾーンのオフセット（時間）
const MONTHS_AGO = 3; // 過去何ヶ月分のデータを取得するか

function main() {
  // 日付範囲の設定
  const dateRange = getDateRange();
  
  // レポートを取得し、指定されたシートにエクスポート
  exportReport('raw_campaign_time_slot', 'campaign', dateRange.startDate, dateRange.endDate);
}

/**
 * 日付範囲を取得する
 * @returns {Object} 開始日と終了日を含むオブジェクト
 */
function getDateRange() {
  const today = new Date();
  today.setHours(today.getHours() + TIMEZONE_OFFSET);
  
  const endDate = Utilities.formatDate(
    new Date(today.getFullYear(), today.getMonth(), today.getDate() + 1),
    "GMT+9",
    "yyyy-MM-dd"
  );
  
  // 開始日を指定月数前に設定
  const monthsAgo = new Date(today);
  monthsAgo.setMonth(today.getMonth() - MONTHS_AGO);
  const startDate = Utilities.formatDate(
    monthsAgo,
    "GMT+9",
    "yyyy-MM-dd"
  );

  console.log("期間設定: " + startDate + " から " + endDate + " (過去" + MONTHS_AGO + "ヶ月)");
  return { startDate, endDate };
}

/**
 * レポートをエクスポートする
 * @param {string} sheetName - シート名
 * @param {string} level - レポートレベル（campaign等）
 * @param {string} startDate - 開始日
 * @param {string} endDate - 終了日
 */
function exportReport(sheetName, level, startDate, endDate) {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(sheetName);

  // シートが存在しない場合は作成
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    console.log(`新しいシート '${sheetName}' を作成しました。`);
  }

  // レポートクエリの構築
  const query = buildReportQuery(level, startDate, endDate);
  
  // レポートの取得とエクスポート
  const report = AdsApp.report(query);
  report.exportToSheet(sheet);

  // データの加工
  processReportData(sheet);
}

/**
 * レポートクエリを構築する
 * @param {string} level - レポートレベル
 * @param {string} startDate - 開始日
 * @param {string} endDate - 終了日
 * @returns {string} 構築されたクエリ
 */
function buildReportQuery(level, startDate, endDate) {
  const selectClause = `SELECT 
    segments.date, 
    segments.hour,
    ${level}.name, 
    metrics.cost_micros, 
    metrics.clicks, 
    metrics.conversions, 
    metrics.impressions, 
    metrics.search_impression_share, 
    metrics.search_top_impression_share, 
    metrics.search_absolute_top_impression_share`;

  const fromClause = ` FROM ${level}`;
  const whereClause = ` WHERE segments.date BETWEEN '${startDate}' AND '${endDate}'`;
  const orderByClause = " ORDER BY segments.date";

  return selectClause + fromClause + whereClause + orderByClause;
}

/**
 * レポートデータを加工する
 * @param {Sheet} sheet - スプレッドシートのシートオブジェクト
 */
function processReportData(sheet) {
  // ヘッダーの設定
  const headers = [
    "日付",
    "時間",
    "キャンペーン名",
    "費用 (円)",
    "クリック数",
    "コンバージョン数",
    "インプレッション数",
    "インプレッションシェア",
    "上部インプレッションシェア",
    "最上部インプレッションシェア"
  ];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return; // データが存在しない場合は処理を終了

  // 費用の変換
  convertCostToYen(sheet, lastRow);
  
  // 数値フォーマットの設定
  formatNumberColumns(sheet, lastRow);
}

/**
 * 費用を日本円に変換する
 * @param {Sheet} sheet - スプレッドシートのシートオブジェクト
 * @param {number} lastRow - 最終行
 */
function convertCostToYen(sheet, lastRow) {
  const costRange = sheet.getRange(2, 4, lastRow - 1);
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
  // クリック数、コンバージョン数、インプレッション数をカンマ区切り数値フォーマット
  sheet.getRange(2, 5, lastRow - 1, 1).setNumberFormat('#,##0'); // クリック数
  sheet.getRange(2, 6, lastRow - 1, 1).setNumberFormat('#,##0.00'); // コンバージョン数
  sheet.getRange(2, 7, lastRow - 1, 1).setNumberFormat('#,##0'); // インプレッション数

  // 費用を通貨フォーマット
  sheet.getRange(2, 4, lastRow - 1, 1).setNumberFormat('¥#,##0');

  // インプレッションシェア関連をパーセントフォーマット
  sheet.getRange(2, 8, lastRow - 1, 3).setNumberFormat('0.00%');
}
  