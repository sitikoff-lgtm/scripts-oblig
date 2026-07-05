// ============================================================
// MoexLoader.gs — С ПАРСИНГОМ PREVPRICE (для выходных)
// ============================================================

function fetchAllMoexDataBatch() {
  var moexCache = {};
  
  var allIsins = [
    "RU000A108P46", "RU000A107RZ0", "RU000A10BPZ1", "RU000A1061K1",
    "RU000A1074G2", "RU000A103BR0", "RU000A101F94", "RU000A100EF5",
    "RU000A106UW3", "RU000A1077X0", "RU000A107XA1", "RU000A1098F3",
    "RU000A10EYJ1", "RU000A10CC99", "RU000A10DZW3", "RU000A10DSG1",
    "RU000A107UU5"
  ];
  
  writeToLogSheet("ИНФО", "Загрузка " + allIsins.length + " бумаг с MOEX (с PREVPRICE)...", 0);
  
  for (var i = 0; i < allIsins.length; i++) {
    var isin = allIsins[i];
    
    var url = "https://moex.com/iss/securities.json?q=" + isin + "&iss.meta=off&iss.only=securities,marketdata";
    
    try {
      if (i > 0) Utilities.sleep(300);
      
      var response = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        headers: { 'User-Agent': 'Mozilla/5.0' }
      });
      
      if (response.getResponseCode() !== 200) {
        writeToLogSheet("ПРЕДУПРЕЖДЕНИЕ", isin + " → код " + response.getResponseCode(), 0);
        continue;
      }
      
      var json = JSON.parse(response.getContentText());
      
      if (!json || !json.securities || !json.securities.data || json.securities.data.length === 0) {
        writeToLogSheet("ПРЕДУПРЕЖДЕНИЕ", isin + " → нет данных", 0);
        continue;
      }
      
      // Парсим ISIN
      var sCols = json.securities.columns;
      var sRow = json.securities.data[0];
      var isinIdx = sCols.indexOf("isin") !== -1 ? sCols.indexOf("isin") : sCols.indexOf("secid");
      var actualIsin = String(sRow[isinIdx]).trim().toUpperCase();
      
      // Базовая структура
      moexCache[actualIsin] = {
        price: 0,
        nkd: 0,
        faceValue: 1000,
        maturity: "",
        couponValue: 0,
        period: 2
      };
      
      // ============================================================
      // ПАРСИМ securities (статика)
      // ============================================================
      var matIdx = sCols.indexOf("matdate");
      if (matIdx !== -1 && sRow[matIdx] !== undefined && sRow[matIdx] !== null) {
        moexCache[actualIsin].maturity = String(sRow[matIdx]).trim();
      }
      
      var faceIdx = sCols.indexOf("facevalue");
      if (faceIdx !== -1 && sRow[faceIdx] !== undefined && sRow[faceIdx] !== null) {
        var face = parseFloat(sRow[faceIdx]);
        if (!isNaN(face) && face > 0) {
          moexCache[actualIsin].faceValue = face;
        }
      }
      
      var periodIdx = sCols.indexOf("couponfrequency");
      if (periodIdx === -1) periodIdx = sCols.indexOf("couponperiod");
      if (periodIdx !== -1 && sRow[periodIdx] !== undefined && sRow[periodIdx] !== null) {
        var period = parseInt(sRow[periodIdx], 10);
        if (!isNaN(period) && period > 0) {
          moexCache[actualIsin].period = period;
        }
      }
      
      var cpnIdx = sCols.indexOf("couponvalue");
      if (cpnIdx !== -1 && sRow[cpnIdx] !== undefined && sRow[cpnIdx] !== null) {
        var cpn = parseFloat(sRow[cpnIdx]);
        if (!isNaN(cpn) && cpn > 0) {
          moexCache[actualIsin].couponValue = cpn;
        }
      }
      
      // ============================================================
      // ПАРСИМ marketdata (цены и НКД)
      // ============================================================
      if (json.marketdata && json.marketdata.data && json.marketdata.data.length > 0) {
        var mCols = json.marketdata.columns;
        var mRow = json.marketdata.data[0];
        
        // Логируем все колонки для диагностики
        var logMsg = actualIsin + " → ";
        
        // === ЦЕНА: пробуем LAST, LCURRENTPRICE, MARKETPRICE, PREVPRICE ===
        var priceNames = ["LAST", "LCURRENTPRICE", "MARKETPRICE", "PREVPRICE"];
        var priceFound = false;
        
        for (var p = 0; p < priceNames.length; p++) {
          var idx = mCols.indexOf(priceNames[p]);
          if (idx !== -1 && mRow[idx] !== undefined && mRow[idx] !== null && mRow[idx] !== "") {
            var val = parseFloat(mRow[idx]);
            if (!isNaN(val) && val > 0) {
              moexCache[actualIsin].price = val;
              priceFound = true;
              logMsg += priceNames[p] + "=" + val + " ";
              break;
            }
          }
        }
        
        // === НКД: ACCRUEDINT ===
        var nkdIdx = mCols.indexOf("ACCRUEDINT");
        if (nkdIdx === -1) nkdIdx = mCols.indexOf("ACCRUEDINT_VALUE");
        
        if (nkdIdx !== -1 && mRow[nkdIdx] !== undefined && mRow[nkdIdx] !== null && mRow[nkdIdx] !== "") {
          var val = parseFloat(mRow[nkdIdx]);
          if (!isNaN(val) && val >= 0) {
            moexCache[actualIsin].nkd = val;
            logMsg += "НКД=" + val;
          }
        }
        
        writeToLogSheet("ИНФО", logMsg, 0);
        
        // Если цена не найдена — пишем предупреждение
        if (!priceFound) {
          writeToLogSheet("ПРЕДУПРЕЖДЕНИЕ", actualIsin + " → цена не найдена (выходной день?)", 0);
        }
        
      } else {
        writeToLogSheet("ПРЕДУПРЕЖДЕНИЕ", actualIsin + " → marketdata отсутствует", 0);
      }
      
    } catch(e) {
      writeToLogSheet("ОШИБКА", isin + " → " + e.message, 0);
    }
  }
  
  var totalFound = Object.keys(moexCache).length;
  writeToLogSheet("ИНФО", "Загрузка завершена: найдено " + totalFound + " из " + allIsins.length, totalFound);
  
  return moexCache;
}

function updateAllBondValuesFromCache() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Облигации");
  if (!sheet) {
    writeToLogSheet("ОШИБКА", "Лист 'Облигации' не найден", 0);
    return;
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    writeToLogSheet("ИНФО", "Лист 'Облигации' пуст", 0);
    return;
  }
  
  writeToLogSheet("ИНФО", "Обновление цен и НКД из кэша MOEX...", 0);
  
  var moexCache = fetchAllMoexDataBatch();
  
  if (!moexCache || Object.keys(moexCache).length === 0) {
    writeToLogSheet("ОШИБКА", "Не удалось загрузить данные с MOEX", 0);
    return;
  }
  
  var isins = sheet.getRange("A2:A" + lastRow).getValues();
  var priceData = [];
  var nkdData = [];
  
  for (var i = 0; i < isins.length; i++) {
    var isin = String(isins[i][0]).replace(/[^A-Z0-9]/gi, '').trim().toUpperCase();
    
    if (isin && moexCache[isin]) {
      priceData.push([moexCache[isin].price || 0]);
      nkdData.push([moexCache[isin].nkd || 0]);
    } else {
      priceData.push([0]);
      nkdData.push([0]);
    }
  }
  
  sheet.getRange("K2:K" + lastRow).setValues(priceData).setNumberFormat('0.00"%"');
  sheet.getRange("L2:L" + lastRow).setValues(nkdData).setNumberFormat('#,##0.00');
  SpreadsheetApp.flush();
  
  writeToLogSheet("УСПЕХ", "Цены и НКД обновлены для " + priceData.length + " бумаг", priceData.length);
}

function testFetchMoex() {
  var result = fetchAllMoexDataBatch();
  
  if (result === undefined) {
    Logger.log("❌ fetchAllMoexDataBatch() вернула undefined!");
  } else {
    var count = Object.keys(result).length;
    Logger.log("✅ Загружено: " + count + " бумаг");
    
    for (var isin in result) {
      Logger.log(isin + " → цена: " + result[isin].price + ", НКД: " + result[isin].nkd);
    }
  }
}