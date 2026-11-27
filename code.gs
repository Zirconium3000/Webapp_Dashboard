// ============================================================
// KONFIGURATION
// ============================================================

// Google-Sheet-ID eintragen.
// Nur die ID zwischen /d/ und /edit verwenden.
const SPREADSHEET_ID = '1zzrDZmACN5ov2uqyBvV1TULH8tmyZ6eUxHt_FBcguH0';

// Geheimer Zugriffsschlüssel (PIN oder Passwort)
// Ändere diesen Wert nach Wunsch (z. B. '1234' oder dein eigenes Passwort).
// Wenn leer (''), ist der Zugriff ungeschützt.
const AUTH_TOKEN = '1234';

// Lang & Schwarz Instrument-IDs für den Intraday-Graphen.
// Schlüssel = exakter Name aus dem Google Sheet (Spalte A in MSD_small_2026),
// Wert = L&S instrumentId (z. B. '45389' für 'AT&S', oder '45389@1').
const LS_INSTRUMENT_IDS = {
  'AT&S': '45389'
};


// ============================================================
// WEB-APP START
// ============================================================

function doGet(e) {
  try {
    const providedToken = e && e.parameter ? e.parameter.token : '';

    // Authentifizierung prüfen (Möglichkeit 1)
    if (AUTH_TOKEN && providedToken !== AUTH_TOKEN) {
      return ContentService
        .createTextOutput(JSON.stringify({
          error: 'Unauthorized',
          message: 'Ungültiger oder fehlender Zugriffsschlüssel (PIN / Token).'
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    const action = e && e.parameter ? (e.parameter.action || '') : '';
    let data;
    if (action === 'intraday') {
      const stockName = e.parameter.name || '';
      const explicitId = e.parameter.id || '';
      data = getLsIntradayData(stockName, explicitId);
    } else {
      data = getDashboardData();
    }
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

    lsIds: LS_INSTRUMENT_IDS,

    chartData:
      chartData

  };

}

// ============================================================
// LANG & SCHWARZ INTRADAY-DATEN (JSON-RPC)
// ============================================================

function parseLsId_(rawId) {
  const str = String(rawId || '').trim();
  if (!str) return null;
  const parts = str.split('@');
  return {
    instrumentId: parts[0].trim(),
    marketId: (parts[1] || '1').trim()
  };
}

function getLsIntradayData(name, explicitId) {
  const rawId = explicitId || (name && LS_INSTRUMENT_IDS[name]) || '';
  const parsed = parseLsId_(rawId);

  if (!parsed || !parsed.instrumentId) {
    return {
      ok: false,
      configured: false,
      name: name || '',
      error: 'Keine L&S-ID für "' + (name || '') + '" in LS_INSTRUMENT_IDS hinterlegt.'
    };
  }

  const cache = CacheService.getScriptCache();
  const cacheKey = 'ls_intraday_' + parsed.instrumentId + '_' + parsed.marketId;
  const cached = cache.get(cacheKey);
  if (cached) {
    try {
      const parsedCache = JSON.parse(cached);
      parsedCache.name = name || parsedCache.name || '';
      parsedCache.cached = true;
      return parsedCache;
    } catch (e) {
      // Cache ignorieren falls korrupt
    }
  }

  const url = 'https://www.ls-tc.de/_rpc/json/instrument/chart/dataForInstrument'
    + '?instrumentId=' + encodeURIComponent(parsed.instrumentId)
    + '&marketId=' + encodeURIComponent(parsed.marketId)
    + '&quotetype=mid&series=intraday&localeId=1';

  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true,
    headers: {
      'User-Agent': 'Mozilla/5.0'
    }
  });

  if (response.getResponseCode() !== 200) {
    return {
      ok: false,
      configured: true,
      name: name || '',
      instrumentId: parsed.instrumentId,
      error: 'L&S HTTP-Fehler ' + response.getResponseCode()
    };
  }

  const json = JSON.parse(response.getContentText());
  const points = (json.series && json.series.intraday && json.series.intraday.data) || [];
  let previousDay = null;
  if (json.info && Array.isArray(json.info.plotlines)) {
    for (let i = 0; i < json.info.plotlines.length; i++) {
      if (json.info.plotlines[i].id === 'previousDay') {
        previousDay = Number(json.info.plotlines[i].value);
        break;
      }
    }
  }

  const lastPoint = points.length ? points[points.length - 1] : null;
  const result = {
    ok: true,
    configured: true,
    name: name || '',
    instrumentId: parsed.instrumentId,
    marketId: parsed.marketId,
    isin: (json.info && json.info.isin) || '',
    previousDay: isNaN(previousDay) ? null : previousDay,
    currentPrice: lastPoint ? Number(lastPoint[1]) : null,
    lastTimestamp: lastPoint ? Number(lastPoint[0]) : null,
    points: points
  };

  try {
    cache.put(cacheKey, JSON.stringify(result), 20);
  } catch (e) {
    // Cache-Limit ignorieren falls Serie sehr groß
  }

  return result;
}
















