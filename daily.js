function main() {
    // スプレッドシートのID（実際の使用時に設定）
    const SPREADSHEET_ID = '';

    // 取得期間設定（月数）
    const MONTHS_AGO = 6;

    // 日付設定
    const today = new Date();
    today.setHours(today.getHours() + 14); // タイムゾーンのオフセット考慮

    const endDate = formatDate(new Date(today.getFullYear(), today.getMonth(), today.getDate() + 1));

    // 開始日を指定月数前に設定
    const monthsAgo = new Date(today);
    monthsAgo.setMonth(today.getMonth() - MONTHS_AGO);
    const startDate = formatDate(monthsAgo);

    const allDate = '2024-01-01';

    // レポート出力
    // exportReport('raw_customer', 'customer', endDate, allDate, SPREADSHEET_ID);
    exportReport('raw_campaign', 'campaign', endDate, startDate, SPREADSHEET_ID);
    // exportReport('raw_ad_group', 'ad_group', endDate, startDate, SPREADSHEET_ID);
  }

  function formatDate(date) {
    return Utilities.formatDate(date, "GMT+9", "yyyy-MM-dd");
  }

  function exportReport(sheetName, level, endDate, startDate, spreadsheetId) {
    // スプレッドシートを開く
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);
    sheet.clear();

    // レポートクエリの作成
    const query = createReportQuery(level, startDate, endDate);

    // レポート取得とエクスポート
    const report = AdsApp.report(query);
    report.exportToSheet(sheet);

    // ヘッダーを日本語に置き換える
    setJapaneseHeaders(sheet, level);

    // 費用を日本円に変換
    convertCostToJPY(sheet);
  }

  function createReportQuery(level, startDate, endDate) {
    let selectClause;

    if (level === 'customer') {
      selectClause = "SELECT segments.date, customer.id, metrics.cost_micros, metrics.clicks, metrics.conversions, metrics.impressions, metrics.search_impression_share, metrics.ctr, metrics.average_cpc ";
    } else if (level === 'campaign') {
      selectClause = "SELECT segments.date, " + level + ".name, metrics.cost_micros, metrics.clicks, metrics.conversions, metrics.impressions, metrics.search_impression_share, metrics.search_top_impression_share, metrics.search_absolute_top_impression_share, metrics.search_rank_lost_impression_share, metrics.search_rank_lost_top_impression_share, metrics.search_rank_lost_absolute_top_impression_share, metrics.search_click_share, metrics.ctr, metrics.average_cpc";
    } else {
      selectClause = "SELECT segments.date, " + level + ".name, metrics.cost_micros, metrics.clicks, metrics.conversions, metrics.impressions, metrics.search_impression_share, metrics.search_top_impression_share, metrics.search_absolute_top_impression_share, metrics.search_rank_lost_impression_share, metrics.search_rank_lost_top_impression_share, metrics.search_rank_lost_absolute_top_impression_share, metrics.ctr, metrics.average_cpc";
    }

    const fromClause = " FROM " + level;
    const whereClause = " WHERE segments.date BETWEEN '" + startDate + "' AND '" + endDate + "' ";
    const orderByClause = " ORDER BY segments.date";

    return selectClause + fromClause + whereClause + orderByClause;
  }

  function setJapaneseHeaders(sheet, level) {
    const headers = level === 'customer'
      ? ["日付", "顧客ID", "費用 (円)", "クリック数", "コンバージョン数", "インプレッション数", "IS", "CTR", "CPC (円)"]
      : ["日付", "名前", "費用 (円)", "クリック数", "コンバージョン数", "インプレッション数", "IS", "上部IS", "最上部IS", "IS損失率(ランク)", "上位IS損失率(ランク)", "最上位IS損失率(ランク)", "クリックシェア", "CTR", "CPC (円)"];

    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
  }

  function convertCostToJPY(sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      // 費用を日本円に変換（3列目）
      const costRange = sheet.getRange(2, 3, lastRow - 1);
      const costValues = costRange.getValues();
      const convertedCostValues = costValues.map(row => [Math.round(row[0] / 1000000)]);
      costRange.setValues(convertedCostValues);

      // CPCを日本円に変換（最後の列）
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const cpcColumnIndex = headers.indexOf("CPC (円)") + 1;

      if (cpcColumnIndex > 0) {
        const cpcRange = sheet.getRange(2, cpcColumnIndex, lastRow - 1);
        const cpcValues = cpcRange.getValues();
        const convertedCPCValues = cpcValues.map(row => [Math.round(row[0] / 1000000)]);
        cpcRange.setValues(convertedCPCValues);
      }
    }
  }