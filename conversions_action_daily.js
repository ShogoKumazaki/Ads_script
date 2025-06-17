function main() {
  // スプレッドシートのID（実際の使用時に設定）
  const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';

  // 取得期間設定（月数）
  const MONTHS_AGO = 6;

  // 現在の日付を YYYY-MM-DD 形式で取得し、前日を終了日とする
  var today = new Date();
  today.setHours(today.getHours() + 14); // タイムゾーンのオフセットを考慮
  var endDate = Utilities.formatDate(new Date(today.getFullYear(), today.getMonth(), today.getDate() + 1), "GMT+9", "yyyy-MM-dd");

  // 開始日を指定月数前に設定
  var monthsAgo = new Date(today);
  monthsAgo.setMonth(today.getMonth() - MONTHS_AGO);
  var startDate = Utilities.formatDate(monthsAgo, "GMT+9", "yyyy-MM-dd");

  console.log("期間設定: " + startDate + " から " + endDate + " (過去" + MONTHS_AGO + "ヶ月)");

  // デイリーのコンバージョンアクション別レポートを取得
  exportConversionReport('raw_campaign_conversion_action', startDate, endDate, SPREADSHEET_ID);
  
  // 広告グループ単位のレポートを取得
  exportAdGroupConversionReport('raw_adgroup_conversion_action', startDate, endDate, SPREADSHEET_ID);
}

function exportConversionReport(sheetName, startDate, endDate, spreadsheetId) {
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheet = spreadsheet.getSheetByName(sheetName);

  // シートが存在しない場合は作成
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    console.log("新しいシート '" + sheetName + "' を作成しました。");
  }

  // 既存のデータをクリア
  var range = sheet.getDataRange();
  range.clearContent();

  // キャンペーン単位でのコンバージョンアクション別データを取得するクエリ
  // segments.conversion_action_name と互換性のないメトリクスを除外
  var selectClause = "SELECT " +
    "segments.date, " +
    "campaign.name, " +
    "campaign.id, " +
    "segments.conversion_action_name, " +
    "metrics.conversions, " +
    "metrics.conversions_value, " +
    "metrics.all_conversions, " +
    "metrics.all_conversions_value ";

  var fromClause = "FROM campaign ";

  var whereClause = "WHERE segments.date BETWEEN '" + startDate + "' AND '" + endDate + "' " +
    "AND metrics.conversions > 0 ";

  var orderByClause = "ORDER BY segments.date DESC, campaign.name ASC, segments.conversion_action_name ASC";

  var query = selectClause + fromClause + whereClause + orderByClause;

  console.log("実行するクエリ: " + query);

  try {
    var report = AdsApp.report(query);
    report.exportToSheet(sheet);

    // 日本語ヘッダーを設定
    var headers = [
      "日付",
      "キャンペーン名",
      "キャンペーンID",
      "コンバージョンアクション名",
      "CV数",
      "CV価値 (円)",
      "全CV数",
      "全CV価値 (円)"
    ];

    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);

    // ヘッダー行のスタイリング
    headerRange.setBackground('#34A853');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');

    var lastRow = sheet.getLastRow();

    if (lastRow > 1) {
      // CV価値を日本円に変換
      convertCostToYen(sheet, lastRow, 6);  // CV価値列
      convertCostToYen(sheet, lastRow, 8);  // 全CV価値列

      // 数値フォーマットの設定
      formatNumberColumns(sheet, lastRow);
    }

    // 列幅の自動調整
    sheet.autoResizeColumns(1, headers.length);

    console.log(sheetName + " のレポート作成が完了しました。取得行数: " + (lastRow - 1) + " 行");

  } catch (error) {
    console.log("エラーが発生しました: " + error.toString());
    throw error;
  }
}

// コンバージョンカテゴリを日本語に変換（CV詳細シート用）
function translateConversionCategory(sheet, lastRow) {
  var categoryRange = sheet.getRange(2, 2, lastRow - 1, 1); // CV詳細シートでは2列目
  var categoryValues = categoryRange.getValues();
  var translatedCategories = categoryValues.map(function(row) {
    var category = row[0];
    switch (category) {
      case 'PURCHASE':
        return ['購入'];
      case 'SIGNUP':
        return ['登録'];
      case 'LEAD':
        return ['リード'];
      case 'PAGE_VIEW':
        return ['ページビュー'];
      case 'DOWNLOAD':
        return ['ダウンロード'];
      case 'ADD_TO_CART':
        return ['カート追加'];
      case 'SUBMIT_LEAD_FORM':
        return ['リードフォーム送信'];
      case 'CONTACT':
        return ['お問い合わせ'];
      case 'BOOK_APPOINTMENT':
        return ['予約'];
      case 'GET_DIRECTIONS':
        return ['経路検索'];
      case 'OUTBOUND_CLICK':
        return ['外部クリック'];
      case 'OTHER':
        return ['その他'];
      default:
        return [category];
    }
  });
  categoryRange.setValues(translatedCategories);
}

// コンバージョンタイプを日本語に変換（CV詳細シート用）
function translateConversionType(sheet, lastRow) {
  var typeRange = sheet.getRange(2, 3, lastRow - 1, 1); // CV詳細シートでは3列目
  var typeValues = typeRange.getValues();
  var translatedTypes = typeValues.map(function(row) {
    var type = row[0];
    switch (type) {
      case 'WEBPAGE':
        return ['ウェブページ'];
      case 'APP_INSTALL':
        return ['アプリインストール'];
      case 'PHONE_CALL_CLICKS':
        return ['電話クリック'];
      case 'IMPORT':
        return ['インポート'];
      case 'GOOGLE_ANALYTICS_4':
        return ['GA4'];
      case 'GOOGLE_ANALYTICS':
        return ['GA'];
      case 'FIREBASE':
        return ['Firebase'];
      case 'CLICK_TO_CALL':
        return ['クリックトゥコール'];
      case 'SALESFORCE':
        return ['Salesforce'];
      case 'AD_CALL':
        return ['広告通話'];
      case 'STORE_SALES_DIRECT_UPLOAD':
        return ['店舗売上直接アップロード'];
      default:
        return [type];
    }
  });
  typeRange.setValues(translatedTypes);
}

// 費用をマイクロ単位から円に変換
function convertCostToYen(sheet, lastRow, columnIndex) {
  var costRange = sheet.getRange(2, columnIndex, lastRow - 1, 1);
  var costValues = costRange.getValues();
  var convertedCostValues = costValues.map(function(row) {
    return [Math.round(row[0] / 1000000)]; // マイクロ単位を円に変換
  });
  costRange.setValues(convertedCostValues);
}

// 数値列のフォーマット設定
function formatNumberColumns(sheet, lastRow) {
  // CV数、全CV数をカンマ区切り数値フォーマット
  sheet.getRange(2, 5, lastRow - 1, 1).setNumberFormat('#,##0.00'); // CV数
  sheet.getRange(2, 7, lastRow - 1, 1).setNumberFormat('#,##0.00'); // 全CV数

  // CV価値、全CV価値を通貨フォーマット
  sheet.getRange(2, 6, lastRow - 1, 1).setNumberFormat('¥#,##0'); // CV価値
  sheet.getRange(2, 8, lastRow - 1, 1).setNumberFormat('¥#,##0'); // 全CV価値
}

// 特定のコンバージョンアクションのみ抽出する関数（オプション）
function exportSpecificConversionAction(conversionActionName, spreadsheetId) {
  // 取得期間設定（月数） - 必要に応じて変更可能
  const MONTHS_AGO = 6;

  var today = new Date();
  today.setHours(today.getHours() + 14);
  var endDate = Utilities.formatDate(new Date(today.getFullYear(), today.getMonth(), today.getDate() + 1), "GMT+9", "yyyy-MM-dd");

  // 開始日を指定月数前に設定
  var monthsAgo = new Date(today);
  monthsAgo.setMonth(today.getMonth() - MONTHS_AGO);
  var startDate = Utilities.formatDate(monthsAgo, "GMT+9", "yyyy-MM-dd");

  var sheetName = 'CV_' + conversionActionName.replace(/[^\w\s]/gi, '').replace(/\s+/g, '_');

  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }

  sheet.clear();

  var selectClause = "SELECT " +
    "segments.date, " +
    "campaign.name, " +
    "segments.conversion_action_name, " +
    "metrics.conversions, " +
    "metrics.conversions_value ";

  var fromClause = "FROM campaign ";

  var whereClause = "WHERE segments.date BETWEEN '" + startDate + "' AND '" + endDate + "' " +
    "AND segments.conversion_action_name = '" + conversionActionName + "' " +
    "AND metrics.conversions > 0 ";

  var orderByClause = "ORDER BY segments.date DESC, metrics.conversions DESC";

  var query = selectClause + fromClause + whereClause + orderByClause;

  console.log("特定CVアクション用クエリ: " + query);

  try {
    var report = AdsApp.report(query);
    report.exportToSheet(sheet);

    var headers = [
      "日付",
      "キャンペーン名",
      "コンバージョンアクション名",
      "CV数",
      "CV価値 (円)"
    ];

    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#9C27B0')
      .setFontColor('white')
      .setFontWeight('bold');

    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      convertCostToYen(sheet, lastRow, 5); // CV価値列

      sheet.getRange(2, 4, lastRow - 1, 1).setNumberFormat('#,##0.00'); // CV数
      sheet.getRange(2, 5, lastRow - 1, 1).setNumberFormat('¥#,##0'); // CV価値
    }

    sheet.autoResizeColumns(1, headers.length);

    console.log("特定コンバージョンアクション '" + conversionActionName + "' のレポートを作成しました。");

  } catch (error) {
    console.log("エラー: " + error.toString());
  }
}

// 使用例:
// exportSpecificConversionAction('購入', SPREADSHEET_ID); // 「購入」CVアクションのみ抽出

// 広告グループ単位のコンバージョンレポートを取得する関数
function exportAdGroupConversionReport(sheetName, startDate, endDate, spreadsheetId) {
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheet = spreadsheet.getSheetByName(sheetName);

  // シートが存在しない場合は作成
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    console.log("新しいシート '" + sheetName + "' を作成しました。");
  }

  // 既存のデータをクリア
  var range = sheet.getDataRange();
  range.clearContent();

  // 広告グループ単位でのコンバージョンアクション別データを取得するクエリ
  var selectClause = "SELECT " +
    "segments.date, " +
    "campaign.name, " +
    "campaign.id, " +
    "ad_group.name, " +
    "ad_group.id, " +
    "segments.conversion_action_name, " +
    "metrics.conversions, " +
    "metrics.conversions_value, " +
    "metrics.all_conversions, " +
    "metrics.all_conversions_value ";

  var fromClause = "FROM ad_group ";

  var whereClause = "WHERE segments.date BETWEEN '" + startDate + "' AND '" + endDate + "' " +
    "AND metrics.conversions > 0 ";

  var orderByClause = "ORDER BY segments.date DESC, campaign.name ASC, ad_group.name ASC, segments.conversion_action_name ASC";

  var query = selectClause + fromClause + whereClause + orderByClause;

  console.log("実行するクエリ: " + query);

  try {
    var report = AdsApp.report(query);
    report.exportToSheet(sheet);

    // 日本語ヘッダーを設定
    var headers = [
      "日付",
      "キャンペーン名",
      "キャンペーンID",
      "広告グループ名",
      "広告グループID",
      "コンバージョンアクション名",
      "CV数",
      "CV価値 (円)",
      "全CV数",
      "全CV価値 (円)"
    ];

    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);

    // ヘッダー行のスタイリング
    headerRange.setBackground('#4285F4');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');

    var lastRow = sheet.getLastRow();

    if (lastRow > 1) {
      // CV価値を日本円に変換
      convertCostToYen(sheet, lastRow, 8);  // CV価値列
      convertCostToYen(sheet, lastRow, 10); // 全CV価値列

      // 数値フォーマットの設定
      formatAdGroupNumberColumns(sheet, lastRow);
    }

    // 列幅の自動調整
    sheet.autoResizeColumns(1, headers.length);

    console.log(sheetName + " のレポート作成が完了しました。取得行数: " + (lastRow - 1) + " 行");

  } catch (error) {
    console.log("エラーが発生しました: " + error.toString());
    throw error;
  }
}

// 広告グループレポート用の数値列フォーマット設定
function formatAdGroupNumberColumns(sheet, lastRow) {
  // CV数、全CV数をカンマ区切り数値フォーマット
  sheet.getRange(2, 7, lastRow - 1, 1).setNumberFormat('#,##0.00'); // CV数
  sheet.getRange(2, 9, lastRow - 1, 1).setNumberFormat('#,##0.00'); // 全CV数

  // CV価値、全CV価値を通貨フォーマット
  sheet.getRange(2, 8, lastRow - 1, 1).setNumberFormat('¥#,##0'); // CV価値
  sheet.getRange(2, 10, lastRow - 1, 1).setNumberFormat('¥#,##0'); // 全CV価値
}