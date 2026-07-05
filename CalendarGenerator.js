// ============================================================
// ГЕНЕРАЦИЯ СВОДНОГО КАЛЕНДАРЯ ВЫПЛАТ
// ============================================================

function prepareDestSheetStructure() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var destSheetName = "Свод_Календарь";
  var destSheet = ss.getSheetByName(destSheetName);
  
  if (destSheet) {
    destSheet.clear();
  } else {
    destSheet = ss.insertSheet(destSheetName);
  }
  
  var headers = [["Месяц", "Эмитент / Выпуск", "ISIN", "Дата отсечки", "Кол-во (шт)", "Ставка (%)", "Сумма грязными (₽)", "НДФЛ 13% (₽)", "Сумма чистыми (₽)", "Рыночный сигнал (YTM)"]];
  destSheet.getRange("A1:J1").setValues(headers).setFontWeight("bold").setBackground("#f3f3f3").setHorizontalAlignment("center");
  
  generateCalendarDataRows(destSheet);
}

function generateCalendarDataRows(destSheet) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName("КалендарьИнвестора");
  if (!sourceSheet) return;
  
  var sourceLastRow = sourceSheet.getLastRow();
  if (sourceLastRow < 2) return;
  
  var calData = sourceSheet.getRange(2, 1, sourceLastRow - 1, 8).getDisplayValues();
  var amortizationMap = getAmortizationData();
  
  var finalData = [];
  var chartValues = [];
  var totalRowsIndices = [];
  var monthsNames = ["ЯНВАРЬ", "ФЕВРАЛЬ", "МАРТ", "АПРЕЛЬ", "МАЙ", "ИЮНЬ", "ИЮЛЬ", "АВГУСТ", "СЕНТЯБРЬ", "ОКТЯБРЬ", "НОЯБРЬ", "ДЕКАБРЬ"];
  
  processMonthsLoop(destSheet, calData, amortizationMap, finalData, chartValues, totalRowsIndices, monthsNames);
}

function processMonthsLoop(destSheet, calData, amortizationMap, finalData, chartValues, totalRowsIndices, monthsNames) {
  for (var m = 0; m < monthsNames.length; m++) {
    var targetMonthName = monthsNames[m];
    var runningDirty = 0, runningNdfl = 0, runningClean = 0;
    var hasBondsInMonth = false;
    
    for (var i = 0; i < calData.length; i++) {
      var dateStr = calData[i][2];
      if (!dateStr) continue;
      
      var parts = dateStr.split('.');
      if (parts.length !== 3) continue;
      
      var bondMonthIdx = parseInt(parts[1], 10) - 1;
      if (bondMonthIdx === m) {
        hasBondsInMonth = true;
        
        var name = calData[i][0];
        var isin = calData[i][1];
        var qty = parseFloat(calData[i][3]) || 0;
        var rateStr = calData[i][4];
        var dirtyValue = parseFloat(calData[i][5].replace(/\s/g, '').replace(',', '.')) || 0;
        var ndfl = parseFloat(calData[i][6].replace(/\s/g, '').replace(',', '.')) || 0;
        var clean = parseFloat(calData[i][7].replace(/\s/g, '').replace(',', '.')) || 0;
        
        var monthLabelToPrint = (!hasBondsInMonth) ? targetMonthName : "";
        var signalText = "⚪ Нет данных YTM";
        
        if (isin && amortizationMap[isin]) {
          var currentYTM = amortizationMap[isin].ytm;
          var maturityStr = amortizationMap[isin].maturity || "";
          signalText = generateMarketSignal(currentYTM, maturityStr);
        }
        
        finalData.push([monthLabelToPrint, name, isin, dateStr, qty, rateStr, dirtyValue, ndfl, clean, signalText]);
        hasBondsInMonth = true;
        
        runningDirty += dirtyValue;
        runningNdfl += ndfl;
        runningClean += clean;
      }
    }
    
    if (hasBondsInMonth) {
      var totalRowNum = finalData.length + 2;
      finalData.push([
        "ИТОГ " + targetMonthName, "ИТОГ_МЕСЯЦ", "", "-", "-", "-",
        Math.round(runningDirty * 100) / 100, Math.round(runningNdfl * 100) / 100,
        Math.round(runningClean * 100) / 100, "-"
      ]);
      totalRowsIndices.push(totalRowNum);
      chartValues.push([targetMonthName, Math.round(runningClean * 100) / 100]);
    }
  }
  
  flushCalendarToSheet(destSheet, finalData, totalRowsIndices, chartValues);
}

function flushCalendarToSheet(destSheet, finalData, totalRowsIndices, chartValues) {
  if (finalData.length === 0) return;
  
  var outputRange = destSheet.getRange(2, 1, finalData.length, 10);
  outputRange.setValues(finalData);
  
  destSheet.getRange("E2:E" + (finalData.length + 1)).setNumberFormat('#,##0');
  destSheet.getRange("G2:I" + (finalData.length + 1)).setNumberFormat('#,##0.00');
  destSheet.getRange(2, 1, finalData.length, 10).setHorizontalAlignment("center");
  
  for (var m = 0; m < totalRowsIndices.length; m++) {
    var rIndex = totalRowsIndices[m];
    destSheet.getRange(rIndex, 1, 1, 10).setFontWeight("bold").setBackground("#eaedf1");
  }
  
  autoUpdateCouponDates();
  renderAnalyticsChart2026(chartValues);
  SpreadsheetApp.flush();
}

function getAmortizationData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Облигации");
  var map = {};
  if (!sheet) return map;
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return map;
  
  var isins = sheet.getRange("A2:A" + lastRow).getValues();
  var currentNominals = sheet.getRange("E2:E" + lastRow).getValues();
  var maturityDates = sheet.getRange("G2:G" + lastRow).getValues();
  var initialNominals = sheet.getRange("H2:H" + lastRow).getValues();
  var ytmValues = sheet.getRange("P2:P" + lastRow).getValues();
  
  for (var i = 0; i < isins.length; i++) {
    var isin = String(isins[i]).replace(/[^A-Z0-9]/gi, '').trim();
    if (isin) {
      var rawYtm = ytmValues[i];
      var cleanYtm = 0;
      if (typeof rawYtm === "string") {
        cleanYtm = parseFloat(rawYtm.replace(",", ".").replace("%", "")) || 0;
        if (rawYtm.indexOf("%") === -1 && cleanYtm > 0) cleanYtm = cleanYtm * 100;
      } else {
        cleanYtm = Number(rawYtm) * 100 || 0;
      }
      
      var rawMatDate = maturityDates[i];
      var finalMatStr = "";
      
      if (rawMatDate instanceof Date) {
        var dd = String(rawMatDate.getDate()).padStart(2, '0');
        var mm = String(rawMatDate.getMonth() + 1).padStart(2, '0');
        finalMatStr = dd + "." + mm + "." + rawMatDate.getFullYear();
      } else if (!isNaN(Number(rawMatDate)) && Number(rawMatDate) > 0) {
        var jsDate = new Date((Number(rawMatDate) - 25569) * 86400 * 1000);
        var ddNum = String(jsDate.getDate()).padStart(2, '0');
        var mmNum = String(jsDate.getMonth() + 1).padStart(2, '0');
        finalMatStr = ddNum + "." + mmNum + "." + jsDate.getFullYear();
      } else {
        var strCheck = String(rawMatDate).trim();
        if (strCheck.indexOf("GMT") !== -1) {
          var matchYear = strCheck.match(/\d{4}/);
          if (matchYear) {
            var cleanCut = strCheck.substring(0, strCheck.indexOf(matchYear) + 4);
            var ts = Date.parse(cleanCut);
            if (!isNaN(ts)) {
              var dObj = new Date(ts);
              var dStr = String(dObj.getDate()).padStart(2, '0');
              var mStr = String(dObj.getMonth() + 1).padStart(2, '0');
              finalMatStr = dStr + "." + mStr + "." + dObj.getFullYear();
            }
          }
        } else {
          finalMatStr = strCheck;
        }
      }
      
      map[isin] = {
        current: Number(currentNominals[i]) || 0,
        initial: Number(initialNominals[i]) || 0,
        ytm: cleanYtm,
        maturity: finalMatStr
      };
    }
  }
  return map;
}

function renderAnalyticsChart2026(chartValues) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var chartSheetName = "Аналитика_Дохода";
  var chartSheet = ss.getSheetByName(chartSheetName);
  
  if (chartSheet) {
    var existingCharts = chartSheet.getCharts();
    for (var i = 0; i < existingCharts.length; i++) {
      chartSheet.removeChart(existingCharts[i]);
    }
    chartSheet.clear();
  } else {
    chartSheet = ss.insertSheet(chartSheetName);
  }
  
  chartSheet.getRange("A1:B1").setValues([["Месяц", "Доход (₽)"]]).setFontWeight("bold").setBackground("#f3f3f3");
  chartSheet.getRange(2, 1, chartValues.length, 2).setValues(chartValues);
  chartSheet.getRange("B2:B" + (chartValues.length + 1)).setNumberFormat('#,##0.00');
  
  var chart = chartSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(chartSheet.getRange("A1:B" + (chartValues.length + 1)))
    .setPosition(1, 4, 10, 0)
    .setOption('title', '💡 Чистый прогноз выплат 2026')
    .setOption('is3D', true)
    .setOption('colors', ['#198754'])
    .setOption('legend', {position: 'none'})
    .build();
    
  chartSheet.insertChart(chart);
}
