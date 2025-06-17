// 定数定義
const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';
const TIMEZONE_OFFSET = 14; // タイムゾーンのオフセット（時間）
const MONTHS_AGO = 3; // 過去何ヶ月分のデータを取得するか

function main() {
  // 日付範囲の設定
  const dateRange = getDateRange();
  
  // レポートを取得し、指定されたシートにエクスポート
  exportReport('raw_prefectures', 'campaign', dateRange.startDate, dateRange.endDate);
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
 * @param {string} level - レポートレベル（campaign/adgroup）
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
    ${level}.name, 
    segments.geo_target_region,
    metrics.cost_micros,
    metrics.impressions,
    metrics.clicks,
    metrics.conversions,
    metrics.conversions_value`;
  const fromClause = " FROM geographic_view";
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
    "名前",
    "都道府県",
    "費用 (円)",
    "インプレッション",
    "クリック数",
    "CV数",
    "CV価値 (円)"
  ];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return; // データが存在しない場合は処理を終了

  // 費用の変換
  convertCostToYen(sheet, lastRow);
  
  // 都道府県名の変換
  convertPrefectureNames(sheet, lastRow);

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
  // インプレッション、クリック数をカンマ区切り数値フォーマット
  sheet.getRange(2, 5, lastRow - 1, 1).setNumberFormat('#,##0'); // インプレッション
  sheet.getRange(2, 6, lastRow - 1, 1).setNumberFormat('#,##0'); // クリック数
  sheet.getRange(2, 7, lastRow - 1, 1).setNumberFormat('#,##0.00'); // CV数

  // 費用、CV価値を通貨フォーマット
  sheet.getRange(2, 4, lastRow - 1, 1).setNumberFormat('¥#,##0'); // 費用
  sheet.getRange(2, 8, lastRow - 1, 1).setNumberFormat('¥#,##0'); // CV価値
}

/**
 * 都道府県名を日本語に変換する
 * @param {Sheet} sheet - スプレッドシートのシートオブジェクト
 * @param {number} lastRow - 最終行
 */
function convertPrefectureNames(sheet, lastRow) {
  const regionRange = sheet.getRange(2, 3, lastRow - 1);
  const regionValues = regionRange.getValues();
  const convertedRegionValues = regionValues.map(row => [translatePrefectureToJapanese(row[0])]);
  regionRange.setValues(convertedRegionValues);
}

/**
 * 都道府県名を英語から日本語に変換する
 * @param {string} prefectureName - 英語の都道府県名
 * @returns {string} 日本語の都道府県名
 */
function translatePrefectureToJapanese(prefectureName) {
  const prefectureMapping = {
    "Hokkaido": "北海道",
    "Aomori": "青森県",
    "Iwate": "岩手県",
    "Miyagi": "宮城県",
    "Akita": "秋田県",
    "Yamagata": "山形県",
    "Fukushima": "福島県",
    "Ibaraki": "茨城県",
    "Tochigi": "栃木県",
    "Gunma": "群馬県",
    "Saitama": "埼玉県",
    "Chiba": "千葉県",
    "Tokyo": "東京都",
    "Kanagawa": "神奈川県",
    "Niigata": "新潟県",
    "Toyama": "富山県",
    "Ishikawa": "石川県",
    "Fukui": "福井県",
    "Yamanashi": "山梨県",
    "Nagano": "長野県",
    "Gifu": "岐阜県",
    "Shizuoka": "静岡県",
    "Aichi": "愛知県",
    "Mie": "三重県",
    "Shiga": "滋賀県",
    "Kyoto": "京都府",
    "Osaka": "大阪府",
    "Hyogo": "兵庫県",
    "Nara": "奈良県",
    "Wakayama": "和歌山県",
    "Tottori": "鳥取県",
    "Shimane": "島根県",
    "Okayama": "岡山県",
    "Hiroshima": "広島県",
    "Yamaguchi": "山口県",
    "Tokushima": "徳島県",
    "Kagawa": "香川県",
    "Ehime": "愛媛県",
    "Kochi": "高知県",
    "Fukuoka": "福岡県",
    "Saga": "佐賀県",
    "Nagasaki": "長崎県",
    "Kumamoto": "熊本県",
    "Oita": "大分県",
    "Miyazaki": "宮崎県",
    "Kagoshima": "鹿児島県",
    "Okinawa": "沖縄県"
  };
  
  return prefectureMapping[prefectureName] || prefectureName;
}
  