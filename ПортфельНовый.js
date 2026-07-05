// ============================================================
// БЛОК 1. ГЛАВНЫЙ ИНТЕРФЕЙС И АВТО-МЕНЮ С ЭМОДЗИ
// ============================================================
function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('💼 ')
      .addItem('🔄 1. Обновить календарь выплат', 'startUniqueBondCalendarGeneration2026')
      .addItem('📅 2. Синхронизировать с Google Календарем', 'syncPaymentsWithGoogleCalendar')
      .addToUi();
  } catch(e) {
    Logger.log("Не удалось добавить меню: " + e.message);
  }
}

// ============================================================
// БЛОК 2. АБСОЛЮТНАЯ ЗАЩИТА И ЗАГРУЗКА ЦЕН С МОЕХ
// ============================================================
function fetchMoexPriceBksCryptoSafe(targetIsinText) {
  if (!targetIsinText) return 0;
  
  var cleanIsin = String(targetIsinText).replace(/[^A-Z0-9]/gi, '').trim();
  if (cleanIsin.length > 12) cleanIsin = cleanIsin.slice(0, 12);
  if (cleanIsin.length < 5) return 0;
  
  var isOfz = (cleanIsin.indexOf("SU") === 0);
  var boardMode = isOfz ? "TQOB" : "TQCB";
  var directMoexUrl = "https://moex.com" + boardMode + "/securities/" + cleanIsin + ".json?iss.meta=off";
  
  try {
    var response = UrlFetchApp.fetch(directMoexUrl, {muteHttpExceptions: true});
    if (response.getResponseCode() !== 200) return 0;
    
    var json = JSON.parse(response.getContentText());
    
    if (json && json.marketdata && json.marketdata.data && json.marketdata.data.length > 0) {
      var columns = json.marketdata.columns;
      var rowData = json.marketdata.data[0];
      
      var priceIdx = columns.indexOf("LAST");
      if (priceIdx === -1 || !rowData[priceIdx]) priceIdx = columns.indexOf("LCURRENTPRICE");
      if (priceIdx === -1 || !rowData[priceIdx]) priceIdx = columns.indexOf("MARKETPRICE");
      
      if (priceIdx !== -1) {
        var price = parseFloat(rowData[priceIdx]);
        if (!isNaN(price) && price > 0) return price; 
      }
    }
    
    if (json && json.securities && json.securities.data && json.securities.data.length > 0) {
      var colsSec = json.securities.columns;
      var rowDataSec = json.securities.data[0];
      
      var idxSec = colsSec.indexOf("PREVPRICE");
      if (idxSec === -1) idxSec = colsSec.indexOf("LEGALCLOSEPRICE");
      
      if (idxSec !== -1) {
        var priceSec = parseFloat(rowDataSec[idxSec]);
        if (!isNaN(priceSec) && priceSec > 0) return priceSec;
      }
    }
    return 0;
  } catch(e) { return 0; }
}

// ============================================================
// БЛОК 3. АВТОМАТИЧЕСКОЕ ОБНОВЛЕНИЕ ДАТ ПРОШЕДШИХ КУПОНОВ
// ============================================================
function autoUpdateCouponDates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Облигации");
  if (!sheet) return;
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  var today = new Date();
  today.setHours(0, 0, 0, 0); 
  
  var isinList = sheet.getRange("A2:A" + lastRow).getDisplayValues();
  var couponRange = sheet.getRange("M2:M" + lastRow); 
  var couponList = couponRange.getDisplayValues();
  var maturityList = sheet.getRange("G2:G" + lastRow).getDisplayValues(); 
  var freqList = sheet.getRange("J2:J" + lastRow).getDisplayValues(); 
  
  var outputDates = [];
  var updated = false;
  
  for (var i = 0; i < isinList.length; i++) {
    var isin = String(isinList[i]).replace(/[^A-Z0-9]/gi, '').trim();
    var rawCouponDate = String(couponList[i]).trim();
    var rawMaturity = String(maturityList[i]).trim();
    var frequency = Number(String(freqList[i]).replace(/[^0-9]/g, '')) || 0;
    
    if (!rawCouponDate || rawCouponDate === "" || rawCouponDate.indexOf("NaN") !== -1 || frequency <= 0) {
      outputDates.push([rawCouponDate]);
      continue;
    }
    
    var couponDate = parseDateLocal(rawCouponDate);
    var maturityDate = parseDateLocal(rawMaturity);
    
    if (!couponDate) {
      outputDates.push([rawCouponDate]);
      continue;
    }
    
    if (couponDate <= today) {
      var monthsStep = 12 / frequency;
      var diffMonths = (today.getFullYear() - couponDate.getFullYear()) * 12 + (today.getMonth() - couponDate.getMonth());
      var stepsNeeded = Math.ceil(diffMonths / monthsStep);
      if (stepsNeeded <= 0) stepsNeeded = 1;
      
      couponDate.setMonth(couponDate.getMonth() + (stepsNeeded * monthsStep));
      if (maturityDate && couponDate > maturityDate) { couponDate = maturityDate; }
      
      var day = String(couponDate.getDate()).padStart(2, '0');
      var month = String(couponDate.getMonth() + 1).padStart(2, '0');
      var year = couponDate.getFullYear();
      outputDates.push([day + '.' + month + '.' + year]);
      updated = true;
    } else {
      outputDates.push([rawCouponDate]);
    }
  }
  
  if (updated) {
    couponRange.setNumberFormat("@"); 
    couponRange.setValues(outputDates);
    SpreadsheetApp.flush();
  }
}


// ============================================================
// БЛОК 4. ГЕНЕРАЦИЯ СВОДНОГО ЛИСТА ВЫПЛАТ (ИСПРАВЛЕННЫЙ ТЕКСТОВЫЙ РЕЖИМ)
// ============================================================
function startUniqueBondCalendarGeneration2026() {
  // 1. Сначала превращаем грязный вертикальный текст БКС в горизонтальную таблицу
  cleanBksVerticalText(); 
  
  // 2. Автоматически заносим новые бумаги на лист "Облигации"
  autoAddNewBondsFromCalendar(); 
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName("КалендарьИнвестора"); 
  if (!sourceSheet) return;
  
  var amortizationMap = getAmortizationData();
  
  var destSheetName = "Свод_Календарь";
  var destSheet = ss.getSheetByName(destSheetName);
  if (destSheet) { destSheet.clear(); } else { destSheet = ss.insertSheet(destSheetName); }
  
  var headers = [["Месяц", "Эмитент / Выпуск", "ISIN", "Дата отсечки", "Кол-во (шт)", "Ставка (%)", "Сумма грязными (₽)", "НДФЛ 13% (₽)", "Сумма чистыми (₽)", "Рыночный сигнал (YTM)"]];
  destSheet.getRange("A1:J1").setValues(headers).setFontWeight("bold").setBackground("#f3f3f3").setHorizontalAlignment("center");

  var sourceLastRow = sourceSheet.getLastRow();
  if (sourceLastRow < 2) return;
  
  // ИСПРАВЛЕНО: Читаем все данные как ТЕКСТ через getDisplayValues(), чтобы даты не превращались в объекты
  var calData = sourceSheet.getRange(2, 1, sourceLastRow - 1, 8).getDisplayValues();
  
  var finalData = [];
  var chartValues = [];
  var totalRowsIndices = [];
  
  var monthsNames = ["ЯНВАРЬ", "ФЕВРАЛЬ", "МАРТ", "АПРЕЛЬ", "МАЙ", "ИЮНЬ", "ИЮЛЬ", "АВГУСТ", "СЕНТЯБРЬ", "ОКТЯБРЬ", "НОЯБРЬ", "ДЕКАБРЬ"];
  
  // Группируем данные по месяцам года
  for (var m = 0; m < monthsNames.length; m++) {
    var targetMonthName = monthsNames[m];
    var runningDirty = 0, runningNdfl = 0, runningClean = 0;
    var hasBondsInMonth = false;
    
    for (var i = 0; i < calData.length; i++) {
      var name = calData[i][0];
      var isin = calData[i][1];
      var dateStr = calData[i][2];
      var qty = parseFloat(calData[i][3]) || 0;
      var rateStr = calData[i][4];
      var dirtyValue = parseFloat(calData[i][5].replace(/\s/g, '').replace(',', '.')) || 0;
      var ndfl = parseFloat(calData[i][6].replace(/\s/g, '').replace(',', '.')) || 0;
      var clean = parseFloat(calData[i][7].replace(/\s/g, '').replace(',', '.')) || 0;
      
      if (!dateStr || dateStr === "") continue;
      
      // Теперь метод split('.') сработает со 100% гарантией
      var parts = dateStr.split('.');
      if (parts.length !== 3) continue;
      var bondMonthIdx = parseInt(parts[1], 10) - 1;
      
      if (bondMonthIdx === m) {
        runningDirty += dirtyValue;
        runningNdfl += ndfl;
        runningClean += clean;
        
        var monthLabelToPrint = (!hasBondsInMonth) ? targetMonthName : "";
        var signalText = "⚪ Нет данных YTM на листе Облигации";
        
        if (isin && amortizationMap[isin]) {
          var currentYTM = amortizationMap[isin].ytm;
          var maturityStr = amortizationMap[isin].maturity || "";
          var dateLabel = maturityStr ? " до " + maturityStr : "";
          
          if (currentYTM > 0) {
            if (currentYTM > 23.0) {
              signalText = "🔥 КУПИТЬ ЕЩЕ (YTM " + currentYTM.toFixed(2) + "%" + dateLabel + ")";
            } else {
              signalText = "🟢 Держать (YTM " + currentYTM.toFixed(2) + "%" + dateLabel + ")";
            }
          }
        }
        
        finalData.push([monthLabelToPrint, name, isin, dateStr, qty, rateStr, dirtyValue, ndfl, clean, signalText]);
        hasBondsInMonth = true;
      }
    }
    
    // Выводим итог месяца
    if (hasBondsInMonth) {
      var totalRowNum = finalData.length + 2; 
      finalData.push([
        "ИТОГ " + targetMonthName, "ИТОГ_МЕСЯЦ", "", "-", "-", "-", 
        Math.round(runningDirty * 100) / 100, Math.round(runningNdfl * 100) / 100, Math.round(runningClean * 100) / 100, "-"
      ]);
      totalRowsIndices.push(totalRowNum);
      chartValues.push([targetMonthName, Math.round(runningClean * 100) / 100]);
    }
  }

  if (finalData.length === 0) return;

  var totalPortfolioRow = finalData.length + 2;
  
  // Выгружаем матрицу на лист
  var outputRange = destSheet.getRange(2, 1, finalData.length, 10);
  outputRange.setValues(finalData);
  
  destSheet.getRange("E2:E" + (finalData.length + 1)).setNumberFormat('#,##0'); 
  destSheet.getRange("G2:I" + (finalData.length + 1)).setNumberFormat('#,##0.00'); 
  
  destSheet.getRange("A1:J1").setWrap(true).setVerticalAlignment("middle");
  destSheet.getRange(2, 1, finalData.length, 10).setWrap(false).setVerticalAlignment("middle");
  destSheet.getRange(1, 1, finalData.length + 1, 10).setBorder(true, true, true, true, true, true, "#d9d9d9", SpreadsheetApp.BorderStyle.SOLID).setHorizontalAlignment("center");
  destSheet.getRange(1, 2, finalData.length + 1, 1).setHorizontalAlignment("left"); 

  destSheet.setRowHeight(1, 36); 
  for (var r = 2; r <= totalPortfolioRow; r++) { destSheet.setRowHeight(r, 20); }

  for (var m = 0; m < totalRowsIndices.length; m++) {
    var rIndex = totalRowsIndices[m];
    destSheet.getRange(rIndex, 1, 1, 10).setFontWeight("bold").setBackground("#eaedf1");
    destSheet.setRowHeight(rIndex, 24); 
  }

  autoUpdateCouponDates(); 
  createIncomeChart(chartValues);
  SpreadsheetApp.flush();
}



// ============================================================
// БЛОК 5. ЧТЕНИЕ ДАННЫХ ОБ АМОРТИЗАЦИИ, YTM И ДАТ ПОГАШЕНИЯ (ИТОГ)
// ============================================================
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

      // ПРЯМОЙ И СТРОГИЙ КОНВЕРТЕР ДАТ ДЛЯ РЫНОЧНОГО СИГНАЛА
      var rawMatDate = maturityDates[i];
      var finalMatStr = "";
      
      if (rawMatDate instanceof Date) {
        // Убираем длинную строку системного времени GMT
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
        // Защита: если внутри текста все-таки проскочила системная строка GMT
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



// ============================================================
// БЛОК 6. ГРАФИК ДОХОДОВ (ДИАПАЗОН СТРОГО С ДАННЫХ СТРОКИ 2)
// ============================================================
function createIncomeChart(chartValues) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var chartSheetName = "Аналитика_Дохода";
  var chartSheet = ss.getSheetByName(chartSheetName);
  
  if (chartSheet) { 
    var existingCharts = chartSheet.getCharts();
    for (var i = 0; i < existingCharts.length; i++) { chartSheet.removeChart(existingCharts[i]); }
    chartSheet.clear(); 
  } else { 
    chartSheet = ss.insertSheet(chartSheetName); 
  }
  
  chartSheet.getRange("A1").setValue("Месяц").setFontWeight("bold").setBackground("#f3f3f3").setHorizontalAlignment("center");
  chartSheet.getRange("B1").setValue("Доход (₽)").setFontWeight("bold").setBackground("#f3f3f3").setHorizontalAlignment("center");
  
  chartSheet.getRange("A2:B" + (chartValues.length + 1)).setValues(chartValues);
  chartSheet.getRange("B2:B" + (chartValues.length + 1)).setNumberFormat('#,##0.00');
  chartSheet.getRange("A1:B" + (chartValues.length + 1)).setBorder(true, true, true, true, true, true, "#d9d9d9", SpreadsheetApp.BorderStyle.SOLID);
  
  chartSheet.setColumnWidth(1, 110);
  chartSheet.setColumnWidth(2, 130);
  
  // ИСПРАВЛЕНО: Убрали А1 из диапазона самого графика, слово "Месяц" стерто с осей!
  var chartDataRange = chartSheet.getRange("A2:B" + (chartValues.length + 1));
  var chart = chartSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(chartDataRange)
    .setPosition(1, 4, 10, 0)
    .setOption('title', '💡 Чистый доход по месяцам')
    .setOption('titleTextStyle', {color: '#1a1a1a', fontSize: 13, bold: true})
    .setOption('legend', {position: 'none'})
    .setOption('colors', ['#198754']) 
    .setOption('vAxis', {gridlines: {count: 5}})
    .setOption('is3D', true)
    .setOption('width', 680)
    .setOption('height', 340)
    .build();
    
  chartSheet.insertChart(chart);
}

// ============================================================
// БЛОК 7. ВСЕЯДНЫЙ ПАРСЕР ДАТ С ПОДДЕРЖКОЙ СЛОЖНЫХ СТРОК
// ============================================================
function parseDateLocal(dateInput) {
  if (!dateInput) return null;
  if (dateInput instanceof Date) return dateInput;
  var str = String(dateInput).trim();
  if (str === "") return null;
  
  if (str.indexOf("GMT") !== -1) {
    var matchYear = str.match(/\d{4}/);
    if (matchYear) {
      var cleanCut = str.substring(0, str.indexOf(matchYear) + 4);
      var ts = Date.parse(cleanCut);
      if (!isNaN(ts)) return new Date(ts);
    }
  }
  
  var textParts = str.split('.');
  if (textParts.length === 3) {
    var dayNumber = parseInt(textParts[0], 10);
    var monthNumber = parseInt(textParts[1], 10) - 1; 
    var yearNumber = parseInt(textParts[2], 10);
    if (yearNumber < 100) yearNumber += 2000;
    if (!isNaN(dayNumber) && !isNaN(monthNumber) && !isNaN(yearNumber)) {
      return new Date(yearNumber, monthNumber, dayNumber);
    }
  }
  return null;
}

// ============================================================
// БЛОК 8. СИНХРОНИЗАЦИЯ С КАЛЕНДАРЕМ (СТРОГИЙ НАУЧНЫЙ ПАРСИНГ ДАТ)
// ============================================================
function syncPaymentsWithGoogleCalendar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Свод_Календарь");
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Ошибка: Сначала сгенерируйте календарь выплат на листе 'Свод_Календарь'!");
    return;
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  var data = sheet.getRange(2, 1, lastRow - 1, 10).getDisplayValues();
  var calName = "💼 Инвестиции БКС (Купоны)";
  var cal;
  
  try {
    var calendars = CalendarApp.getCalendarsByName(calName);
    if (calendars.length > 0) {
      cal = calendars[0];
    } else {
      cal = CalendarApp.createCalendar(calName);
    }
  } catch(e) {
    SpreadsheetApp.getUi().alert("Google заблокировал доступ к Календарю: " + e.toString());
    return;
  }
  
  var past = new Date(2020, 0, 1);
  var farFuture = new Date(3000, 0, 1);
  try {
    var oldEvents = cal.getEvents(past, farFuture);
    for (var i = 0; i < oldEvents.length; i++) { oldEvents[i].deleteEvent(); }
  } catch(e) {}
  
  var count = 0;
  
  for (var j = 0; j < data.length; j++) {
    var name = String(data[j][1]).trim();   
    var isin = String(data[j][2]).trim();   
    var rawDateStr = String(data[j][3]).trim(); 
    var cleanSum = data[j][8];              
    var signal = String(data[j][9]).trim(); 
    
    if (name === "ИТОГ_МЕСЯЦ" || name === "" || name.indexOf("ИТОГ") !== -1 || name.indexOf("ОБЩИЙ") !== -1) {
      continue;
    }
    
    var parts = rawDateStr.split('.');
    if (parts.length !== 3) continue;
    
    var day = parseInt(parts[0], 10);
    var month = parseInt(parts[1], 10) - 1; 
    var year = parseInt(parts[2], 10);
    
    if (isNaN(day) || isNaN(month) || isNaN(year)) continue;
    
    var eventDate = new Date(year, month, day, 10, 0, 0);
    if (isNaN(eventDate.getTime())) continue;
    
    var title = "💰 Купон: " + name + " (+" + cleanSum + " ₽)";
    if (signal.indexOf("🚨") !== -1) {
      title = "🚨 ПРОДАТЬ " + name + "!";
    } else if (signal.indexOf("🔥") !== -1) {
      title = "🔥 КУПИТЬ " + name + "!";
    }
    
    var actionPlan = "";
    if (signal.indexOf("🚨") !== -1) {
      actionPlan = "💡 АНАЛИТИКА БКС:\n⚠️ Бумага перегрета, текущая доходность слишком низкая. Рекомендуется ПРОДАТЬ этот выпуск в приложении БКС, зафиксировать прибыль от роста стоимости и переложить деньги в выпуски с более высокой YTM.";
    } else if (signal.indexOf("🔥") !== -1) {
      actionPlan = "💡 АНАЛИТИКА БКС:\n📈 Аномально высокая доходность! Биржевая цена занижена. Отличный момент, чтобы использовать поступающие купоны и докупить этот выпуск в портфель для разгона капитала.";
    } else {
      actionPlan = "💡 АНАЛИТИКА БКС:\n🟢 Стабильный баланс. Бумага приносит плановый доход, купоны отправляем на реинвестирование.";
    }
    
    var description = "📋 СВОДКА ПО ВЫПЛАТЕ:\n" +
                      "• Облигация: " + name + "\n" +
                      "• Код ISIN: " + isin + "\n" +
                      "• Сумма чистыми на счет: " + cleanSum + " ₽\n" +
                      "• Текущий статус YTM: " + signal + "\n\n" + 
                      "----------------------------------------\n" +
                      actionPlan;
    
    try {
      cal.createEvent(title, eventDate, eventDate, {description: description});
      count++;
    } catch(e) {
      Logger.log("Ошибка записи события: " + e.toString());
    }
  }
  
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert("🚀 Макс-версия: Синхронизация завершена! Успешно перенесено " + count + " инвестиционных событий на 2026 год.");
}


// ============================================================
// БЛОК 9. ИНТЕЛЛЕКТУАЛЬНЫЙ ПОТОКОВЫЙ ПАРСЕР ЛЕНТЫ БКС (ИСПРАВЛЕННЫЙ)
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
  
  // ИСПРАВЛЕНО: Цикл строго работает по индексу 'p' без посторонних букв 'r'
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
    
    if (!currentObj) continue;
    
    if (line.match(/\d{2}\.\d{2}\.\d{4}/) || line.indexOf("GMT") !== -1 || line.indexOf("стандартное") !== -1) {
      var dateMatch = line.match(/\d{2}\.\d{2}\.\d{4}/);
      if (dateMatch) {
        currentObj.date = dateMatch[0];
      } else {
        var matchYear = line.match(/\d{4}/);
        if (matchYear) {
          var cleanCut = line.substring(0, line.indexOf(matchYear[0]) + 4);
          var ts = Date.parse(cleanCut);
          if (!isNaN(ts)) {
            var dObj = new Date(ts);
            var dd = String(dObj.getDate()).padStart(2, '0');
            var mm = String(dObj.getMonth() + 1).padStart(2, '0');
            currentObj.date = dd + "." + mm + "." + dObj.getFullYear();
          }
        }
      }
      continue;
    }
    
    if (line.indexOf("на") !== -1 && line.indexOf("шт") !== -1) {
      currentObj.qty = parseFloat(line.replace(/[^0-9]/g, '')) || 0;
      continue;
    }
    
    if (line.indexOf("%") !== -1) {
      currentObj.rate = line;
      continue;
    }
    
    if (line.indexOf(",") !== -1 || line.indexOf(" ") !== -1 || !isNaN(parseFloat(line))) {
      var cleanNum = parseFloat(line.replace(/\s/g, '').replace(',', '.'));
      if (!isNaN(cleanNum) && cleanNum > 0 && cleanNum !== currentObj.qty) {
        if (cleanNum < 100 && currentObj.rate === "0,00%" && line.indexOf(",") !== -1) {
          currentObj.rate = line + "%";
        } else {
          if (currentObj.dirty === 0) currentObj.dirty = cleanNum;
        }
      }
    }
  }
  
  if (currentObj && currentObj.isin && currentObj.dirty > 0) {
    cleanRows.push(buildFinalBksArray(currentObj));
  }
  
  sheet.clear();
  var headers = [["Эмитент / Выпуск", "ISIN", "Дата отсечки", "Кол-во (шт)", "Ставка (%)", "Сумма грязными (₽)", "НДФЛ 13% (₽)", "Сумма чистыми (₽)"]];
  sheet.getRange("A1:H1").setValues(headers).setFontWeight("bold").setBackground("#f3f3f3").setHorizontalAlignment("center");
  
  if (cleanRows.length > 0) {
    sheet.getRange(2, 1, cleanRows.length, 8).setValues(cleanRows);
    sheet.getRange(2, 6, cleanRows.length, 3).setNumberFormat('#,##0.00');
    sheet.getRange(2, 4, cleanRows.length, 1).setNumberFormat('#,##0');
    for (var col = 1; col <= 8; col++) { sheet.autoResizeColumn(col); }
  }
}

function buildFinalBksArray(obj) {
  var ndfl = Math.round(obj.dirty * 0.13);
  var clean = Math.round((obj.dirty - ndfl) * 100) / 100;
  var finalDate = obj.date ? obj.date : "23.06.2026";
  return [obj.name, obj.isin, finalDate, obj.qty, obj.rate, obj.dirty, ndfl, clean];
}

// ============================================================
// БЛОК 10. АВТО-ДОБАВЛЕНИЕ СЕТЕВЫХ ДАННЫХ ИЗ ШЛЮЗА MOEX
// ============================================================
function autoAddNewBondsFromCalendar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var bondsSheet = ss.getSheetByName("Облигации");
  var calSheet = ss.getSheetByName("КалендарьИнвестора");
  if (!bondsSheet || !calSheet) return;
  
  var bondsLastRow = bondsSheet.getLastRow();
  var calLastRow = calSheet.getLastRow();
  if (calLastRow < 2) return;
  
  var existingIsins = {};
  if (bondsLastRow >= 2) {
    var bIsins = bondsSheet.getRange("A2:A" + bondsLastRow).getValues();
    for (var k = 0; k < bIsins.length; k++) {
      existingIsins[String(bIsins[k]).trim().toUpperCase()] = true;
    }
  }
  
  var calData = calSheet.getRange(2, 1, calLastRow - 1, 6).getValues();
  
  for (var i = 0; i < calData.length; i++) {
    var rawName = String(calData[i][0]).trim();
    var rawIsin = String(calData[i][1]).trim().toUpperCase();
    var rawDate = calData[i][2];
    var rawQty = calData[i][3];
    
    if (!rawIsin || rawIsin === "" || rawIsin.indexOf("ИТОГ") !== -1) continue;
    
    if (!existingIsins[rawIsin]) {
      var nextRow = bondsSheet.getLastRow() + 1;
      
      // ИСПРАВЛЕНО: Прямой монолитный сетевой запрос шлюза Мосбиржи без внешних зависимостей
      var moexCouponPercent = 0;
      var isOfz = (rawIsin.indexOf("SU") === 0);
      var boardMode = isOfz ? "TQOB" : "TQCB";
      var url = "https://moex.com" + boardMode + "/securities/" + rawIsin + ".json?iss.meta=off&iss.only=securities&securities.columns=COUPONPERCENT";
      
      try {
        var response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
        var json = JSON.parse(response.getContentText());
        if (json && json.securities && json.securities.data && json.securities.data.length > 0) {
          moexCouponPercent = parseFloat(json.securities.data[0][0]) || 0;
        }
      } catch(e) { Logger.log("Ошибка шлюза Мосбиржи: " + e.message); }
      
      bondsSheet.getRange(nextRow, 1).setValue(rawIsin);          
      bondsSheet.getRange(nextRow, 2).setValue(rawName);          
      bondsSheet.getRange(nextRow, 3).setValue("Нет");            
      bondsSheet.getRange(nextRow, 4).setValue(rawQty);           
      bondsSheet.getRange(nextRow, 6).setValue(rawDate);          
      bondsSheet.getRange(nextRow, 8).setValue(moexCouponPercent); 
      bondsSheet.getRange(nextRow, 10).setValue(12);              
      bondsSheet.getRange(nextRow, 14).setValue(1000);            
      
      bondsSheet.getRange(nextRow, 16).setFormula("=H" + nextRow + "/J" + nextRow); 
      
      existingIsins[rawIsin] = true;
    }
  }
  SpreadsheetApp.flush();
}