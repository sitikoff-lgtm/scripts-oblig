// ============================================================
// ПАРСЕР ЛЕНТЫ БКС
// ============================================================

function cleanBksVerticalText() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("КалендарьИнвестора");
  if (!sheet) return;
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return;
  
  var rawValues = sheet.getRange(1, 1, lastRow, 1).getDisplayValues();
  var lines = [];
  
  for (var k = 0; k < rawValues.length; k++) {
    var lineText = String(rawValues[k]).trim();
    if (lineText !== "") lines.push(lineText);
  }
  
  var cleanRows = [];
  var currentObj = null;
  
  for (var p = 0; p < lines.length; p++) {
    var line = lines[p];
    
    if (line === "Дата отсечки" || line === "Сумма" || line === "Купон" || line === "₽" || line === "%") {
      continue;
    }
    
    var isinMatch = line.match(/^(RU|SU)[A-Z0-9]{10}$/i) || (line === "MGKL1P3");
    
    if (isinMatch) {
      if (currentObj && currentObj.isin && currentObj.dirty > 0) {
        cleanRows.push(buildFinalBksArray(currentObj));
      }
      
      var finalIsin = line.toUpperCase().trim();
      if (finalIsin === "MGKL1P3") finalIsin = "RU000A10FDE3";
      
      currentObj = {
        name: (p >= 1) ? lines[p-1] : "Облигация",
        isin: finalIsin,
        date: "",
        qty: 0,
        rate: "0,00%",
        dirty: 0
      };
      
      if (currentObj.name === currentObj.isin && p >= 2) currentObj.name = lines[p-2];
      continue;
    }
    
    parseBksDetailsLine(line, currentObj);
  }
  
  flushBksCleanRowsToSheet(sheet, cleanRows, currentObj);
}

function parseBksDetailsLine(line, currentObj) {
  if (!currentObj) return;
  
  var dateMatch = line.match(/^\d{2}\.\d{2}\.\d{4}$/);
  if (dateMatch) {
    currentObj.date = line;
    return;
  }
  
  if (line.indexOf('%') !== -1) {
    currentObj.rate = line;
    return;
  }
  
  if (line.toLowerCase().indexOf('шт') !== -1 || line.toLowerCase().indexOf('pcs') !== -1) {
    currentObj.qty = parseFloat(line.replace(/[^0-9]/g, '')) || 0;
    return;
  }
  
  if (line.indexOf('₽') !== -1 || line.indexOf(',') !== -1 || !isNaN(parseFloat(line.replace(/\s/g, '').replace(',', '.')))) {
    var val = parseFloat(line.replace(/[^\d.,-]/g, '').replace(/\s/g, '').replace(',', '.')) || 0;
    if (val > 0 && currentObj.dirty === 0) {
      currentObj.dirty = val;
    }
  }
}

function buildFinalBksArray(obj) {
  var ndfl = Math.round((obj.dirty * 0.13) * 100) / 100;
  var clean = Math.round((obj.dirty - ndfl) * 100) / 100;
  
  return [
    obj.name,
    obj.isin,
    obj.date,
    obj.qty,
    obj.rate,
    obj.dirty,
    ndfl,
    clean
  ];
}

function flushBksCleanRowsToSheet(sheet, cleanRows, currentObj) {
  if (currentObj && currentObj.isin && currentObj.dirty > 0) {
    cleanRows.push(buildFinalBksArray(currentObj));
  }
  
  if (cleanRows.length === 0) return;
  
  sheet.clear();
  
  var headers = [["Эмитент / Выпуск", "ISIN", "Дата отсечки", "Кол-во (шт)", "Ставка (%)", "Сумма грязными (₽)", "НДФЛ 13% (₽)", "Сумма чистыми (₽)"]];
  sheet.getRange(1, 1, 1, 8).setValues(headers).setFontWeight("bold").setBackground("#e6effa");
  
  sheet.getRange(2, 1, cleanRows.length, 8).setValues(cleanRows);
  sheet.getRange("D2:D" + (cleanRows.length + 1)).setNumberFormat('#,##0');
  sheet.getRange("F2:H" + (cleanRows.length + 1)).setNumberFormat('#,##0.00');
  sheet.autoResizeColumns(1, 8);
  
  SpreadsheetApp.flush();
}

function autoAddNewBondsFromCalendar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var bondSheet = ss.getSheetByName("Облигации");
  var calSheet = ss.getSheetByName("КалендарьИнвестора");
  
  if (!bondSheet || !calSheet || calSheet.getLastRow() < 2) return;
  
  var existingIsins = {};
  var bondLastRow = bondSheet.getLastRow();
  if (bondLastRow >= 2) {
    var bondIsins = bondSheet.getRange("A2:A" + bondLastRow).getValues();
    for (var i = 0; i < bondIsins.length; i++) {
      var bIsin = String(bondIsins[i][0]).trim().toUpperCase();
      if (bIsin) existingIsins[bIsin] = true;
    }
  }
  
  var calLastRow = calSheet.getLastRow();
  var calIsins = calSheet.getRange("B2:B" + calLastRow).getValues();
  var calNames = calSheet.getRange("A2:A" + calLastRow).getValues();
  
  var moexCache = fetchAllMoexDataBatch();
  var newBondsData = [];
  
  for (var j = 0; j < calIsins.length; j++) {
    var targetIsin = String(calIsins[j][0]).trim().toUpperCase();
    var emitterName = String(calNames[j][0]).trim();
    
    if (targetIsin && !existingIsins[targetIsin] && moexCache[targetIsin]) {
      var moexInfo = moexCache[targetIsin];
      var newRow = new Array(20).fill("");
      
      newRow[0] = targetIsin;
      newRow[1] = emitterName || "Новый выпуск";
      newRow[6] = moexInfo.maturity || "";
      newRow[9] = moexInfo.period || 2;
      newRow[10] = moexInfo.price || 100;
      newRow[11] = moexInfo.nkd || 0;
      newRow[13] = moexInfo.faceValue || 1000;
      newRow[15] = moexInfo.couponValue || 0;
      
      newBondsData.push(newRow);
      existingIsins[targetIsin] = true;
    }
  }
  
  if (newBondsData.length > 0) {
    var startRow = bondSheet.getLastRow() + 1;
    bondSheet.getRange(startRow, 1, newBondsData.length, 20).setValues(newBondsData);
    updateAllBondValuesFromCache();
    writeToLogSheet("УСПЕХ", "Добавлено новых бумаг: " + newBondsData.length, newBondsData.length);
  }
}
