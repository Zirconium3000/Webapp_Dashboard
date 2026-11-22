// ============================================================
// KONFIGURATION
// ============================================================

// Google-Sheet-ID eintragen.
// Nur die ID zwischen /d/ und /edit verwenden.
const SPREADSHEET_ID = '1zzrDZmACN5ov2uqyBvV1TULH8tmyZ6eUxHt_FBcguH0';


// ============================================================
// WEB-APP START
// ============================================================

function doGet(e) {
  try {
    const data = getDashboardData();
    const jsonOutput = JSON.stringify(data);
    const callback = e && e.parameter ? e.parameter.callback : '';

    // Optionaler JSONP-Zweig für Browser, die keinen CORS-Fetch erlauben.
    if (callback) {
      if (!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
        return ContentService.createTextOutput('Invalid callback')
          .setMimeType(ContentService.MimeType.TEXT);
      }
      return ContentService
        .createTextOutput(callback + '(' + jsonOutput + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }

    return ContentService
      .createTextOutput(jsonOutput)
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


// ============================================================
// GOOGLE SHEET HOLEN
// ============================================================

function getSheet(spreadsheet, sheetName) {
  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error('Tabellenblatt nicht gefunden: ' + sheetName);
  }

  return sheet;
}


// ============================================================
// DASHBOARD-DATEN LADEN
// ============================================================

function getDashboardData() {

  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getSheet(spreadsheet, 'MSD_small_2026');
  const trendSheet = getSheet(spreadsheet, 'MSD_trend_2026');
  const trendData = trendSheet.getRange('A2:F21').getValues();

  // ----------------------------------------------------------
  // INTRADAY DELTA
  //
  // E17 = Hauptwert
  // F17 = Veränderung zum Vortag
  // ----------------------------------------------------------

  const intradayDelta =
    Number(
      sheet
        .getRange('E17')
        .getValue()
    ) || 0;

  const previousDayDelta =
    Number(
      sheet
        .getRange('F17')
        .getValue()
    ) || 0;
  const weekPerformance = Number(sheet.getRange('D38').getValue()) || 0;
  const weekPercentage = Number(sheet.getRange('F38').getValue()) || 0;

  // ----------------------------------------------------------
  // PERFORMANCE-DIAGRAMM
  //
  // A2:A15 = Namen
  // E2:E15 = Werte in €
  // F2:F15 = Werte in %
  // ----------------------------------------------------------

  const names =
    sheet
      .getRange('A2:A15')
      .getDisplayValues();

  const values =
    sheet
      .getRange('E2:E15')
      .getValues();

  const percentages =
    sheet
      .getRange('F2:F15')
      .getValues();

  const prices = sheet.getRange('B2:C15').getValues();
  const weeklyNames = sheet.getRange('A23:A36').getDisplayValues();
  const weeklyPrices = sheet.getRange('B23:C36').getValues();
  const weeklyPerformance = sheet.getRange('D23:D36').getValues();
  const weeklyPercentages = sheet.getRange('F23:F36').getValues();
  const tradeTime = sheet.getRange('D17').getDisplayValue();
  const ytdValues = sheet.getRange('H2:H15').getValues();
  const ytdPercentages = sheet.getRange('I2:I15').getValues();
  const ytdTotal = Number(sheet.getRange('H17').getValue()) || 0;
  const ytdTotalPercentage = Number(sheet.getRange('I17').getValue()) || 0;
  const totalValues = sheet.getRange('K2:K15').getValues();
  const totalPercentages = sheet.getRange('L2:L15').getValues();
  const totalPortfolio = Number(sheet.getRange('K17').getValue()) || 0;
  const totalPortfolioPercentage = Number(sheet.getRange('L17').getValue()) || 0;

  // ----------------------------------------------------------
  // DATEN FÜR DIAGRAMM AUFBEREITEN
  // ----------------------------------------------------------

  const ytdData = [];
  const totalData = [];
  const chartData = [];
  const weeklyData = [];

  for (let i = 0; i < 14; i++) {

    const name =
      names[i][0];

    let value =
      Number(
        values[i][0]
      );

    // Falls keine gültige Zahl vorhanden ist
    if (isNaN(value)) {

      value = 0;

    }

    chartData.push({

      name: name,

      value: value,

      percentage: Number(percentages[i][0]) || 0,

      previousPrice: Number(prices[i][0]) || 0,

      currentPrice: Number(prices[i][1]) || 0,

    });
  }

  // ----------------------------------------------------------
  for (let i = 0; i < 14; i++) {
    weeklyData.push({
      name: weeklyNames[i][0],
      previousPrice: Number(weeklyPrices[i][0]) || 0,
      currentPrice: Number(weeklyPrices[i][1]) || 0,
      performance: Number(weeklyPerformance[i][0]) || 0,
      percentage: Number(weeklyPercentages[i][0]) || 0
    });
  }

  
  for (let i = 0; i < 14; i++) {
    ytdData.push({
      name: names[i][0],
      value: Number(ytdValues[i][0]) || 0,
      percentage: Number(ytdPercentages[i][0]) || 0
    });
    totalData.push({
      name: names[i][0],
      value: Number(totalValues[i][0]) || 0,
      percentage: Number(totalPercentages[i][0]) || 0
    });
  }
  // DATEN AN WEB-APP ZURÜCKGEBEN
  // ----------------------------------------------------------

  return {

    intradayDelta:
      intradayDelta,

    previousDayDelta:
      previousDayDelta,
    weekPerformance: weekPerformance,
    weekPercentage: weekPercentage,
    weeklyData: weeklyData,

    tradeTime: tradeTime,

    ytdTotal: ytdTotal,
    ytdTotalPercentage: ytdTotalPercentage,
    ytdData: ytdData,
    totalPortfolio: totalPortfolio,
    totalPortfolioPercentage: totalPortfolioPercentage,
    totalData: totalData,

    trendData: trendData,

    chartData:
      chartData

  };

}
















