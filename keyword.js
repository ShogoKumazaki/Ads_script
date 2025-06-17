// 定数定義
const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';
const TIMEZONE_OFFSET = 14; // タイムゾーンのオフセット（時間）

// 取得期間の定義
const PERIODS = [
  { name: 'キーワード_昨日', days: 1 },
  { name: 'キーワード_過去7日', days: 7 },
  { name: 'キーワード_過去30日', days: 30 },
  { name: 'キーワード_過去90日', days: 90 }
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
    ad_group_criterion.keyword.text, 
    ad_group_criterion.keyword.match_type, 
    campaign.name, 
    ad_group.name,    
    metrics.cost_micros,
    metrics.impressions, 
    metrics.clicks, 
    metrics.ctr, 
    metrics.average_cpc,
    metrics.conversions,
    metrics.conversions_from_interactions_rate, 
    metrics.cost_per_conversion, 
    metrics.search_impression_share, 
    metrics.search_top_impression_share, 
    metrics.search_absolute_top_impression_share,
    metrics.search_rank_lost_impression_share,
    metrics.search_rank_lost_top_impression_share,
    metrics.search_rank_lost_absolute_top_impression_share,
    metrics.search_exact_match_impression_share,
    metrics.historical_quality_score,
    metrics.historical_search_predicted_ctr,
    metrics.historical_landing_page_quality_score,
    metrics.historical_creative_quality_score`;

  const fromClause = " FROM keyword_view";
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
    "CPA",
    "IS",
    "上位IS",
    "最上位IS",
    "IS損失率(ランク)",
    "上位IS損失率(ランク)",
    "最上位IS損失率(ランク)",
    "完全一致のIS",
    "品質スコア",
    "推定CTR",
    "LPの利便性",
    "広告の関連性"
  ];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return; // データが存在しない場合は処理を終了

  // マッチタイプの変換
  convertMatchType(sheet, lastRow);
  
  // 費用関連の変換
  convertCostMetrics(sheet, lastRow);
  
  // 品質スコア関連の変換
  convertQualityScores(sheet, lastRow);
  
  // 数値フォーマットの設定
  formatNumberColumns(sheet, lastRow);
}

/**
 * マッチタイプを日本語に変換する
 * @param {Sheet} sheet - スプレッドシートのシートオブジェクト
 * @param {number} lastRow - 最終行
 */
function convertMatchType(sheet, lastRow) {
  const matchTypeRange = sheet.getRange(2, 2, lastRow - 1, 1);
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
  convertCostToYen(sheet, lastRow, 5);
  
  // CPCの変換
  convertCostToYen(sheet, lastRow, 9);
  
  // CPAの変換
  convertCostToYen(sheet, lastRow, 12);
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
 * 品質スコア関連の値を日本語に変換する
 * @param {Sheet} sheet - スプレッドシートのシートオブジェクト
 * @param {number} lastRow - 最終行
 */
function convertQualityScores(sheet, lastRow) {
  // 推定CTRの変換
  convertQualityScore(sheet, lastRow, 21, 'CTR');
  
  // LPの利便性の変換
  convertQualityScore(sheet, lastRow, 22, 'LP');
  
  // 広告の関連性の変換
  convertQualityScore(sheet, lastRow, 23, 'AD');
}

/**
 * 品質スコアを日本語に変換する
 * @param {Sheet} sheet - スプレッドシートのシートオブジェクト
 * @param {number} lastRow - 最終行
 * @param {number} column - 変換対象の列番号
 * @param {string} type - 変換タイプ（CTR/LP/AD）
 */
function convertQualityScore(sheet, lastRow, column, type) {
  const range = sheet.getRange(2, column, lastRow - 1, 1);
  const values = range.getValues();
  const translatedValues = values.map(row => {
    const value = row[0];
    switch (value) {
      case 'AVERAGE':
        return ['平均'];
      case 'BELOW_AVERAGE':
        return ['平均以下'];
      case 'ABOVE_AVERAGE':
        return ['平均以上'];
      default:
        return [value];
    }
  });
  range.setValues(translatedValues);
}

/**
 * 数値列のフォーマット設定
 * @param {Sheet} sheet - スプレッドシートのシートオブジェクト
 * @param {number} lastRow - 最終行
 */
function formatNumberColumns(sheet, lastRow) {
  // インプレッション、クリック数をカンマ区切り数値フォーマット
  sheet.getRange(2, 6, lastRow - 1, 1).setNumberFormat('#,##0'); // インプレッション
  sheet.getRange(2, 7, lastRow - 1, 1).setNumberFormat('#,##0'); // クリック数
  sheet.getRange(2, 10, lastRow - 1, 1).setNumberFormat('#,##0.00'); // CV数

  // コスト、CPC、CPAを通貨フォーマット
  sheet.getRange(2, 5, lastRow - 1, 1).setNumberFormat('¥#,##0'); // コスト
  sheet.getRange(2, 9, lastRow - 1, 1).setNumberFormat('¥#,##0'); // CPC
  sheet.getRange(2, 12, lastRow - 1, 1).setNumberFormat('¥#,##0'); // CPA

  // CTR、CVR、IS関連をパーセントフォーマット
  sheet.getRange(2, 8, lastRow - 1, 1).setNumberFormat('0.00%'); // CTR
  sheet.getRange(2, 11, lastRow - 1, 1).setNumberFormat('0.00%'); // CVR
  sheet.getRange(2, 13, lastRow - 1, 7).setNumberFormat('0.00%'); // IS関連

  // 品質スコアを数値フォーマット
  sheet.getRange(2, 20, lastRow - 1, 1).setNumberFormat('0.0'); // 品質スコア
}
  
  