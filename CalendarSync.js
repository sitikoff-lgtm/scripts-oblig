// ============================================================
// СИНХРОНИЗАЦИЯ С GOOGLE КАЛЕНДАРЕМ
// ============================================================

function syncPaymentsWithGoogleCalendar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Свод_Календарь");
  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert("❌ Сначала сгенерируйте календарь выплат!");
    return;
  }
  
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getDisplayValues();
  var calName = "💼 Инвестиции БКС (Купоны)";
  var calendars = CalendarApp.getCalendarsByName(calName);
  var cal = (calendars.length > 0) ? calendars[0] : CalendarApp.createCalendar(calName);
  
  // Очистка старых событий
  var oldEvents = cal.getEvents(new Date(2020, 0, 1), new Date(3000, 0, 1));
  for (var i = 0; i < oldEvents.length; i++) {
    oldEvents[i].deleteEvent();
  }
  
  var count = 0;
  
  for (var j = 0; j < data.length; j++) {
    var name = String(data[j][1]).trim();
    var isin = String(data[j][2]).trim();
    var rawDateStr = String(data[j][3]).trim();
    var cleanSum = data[j][8];
    var signal = String(data[j][9]).trim();
    
    if (name === "ИТОГ_МЕСЯЦ" || name === "" || name.indexOf("ИТОГ") !== -1) {
      continue;
    }
    
    var parts = rawDateStr.split('.');
    if (parts.length !== 3) continue;
    
    var day = parseInt(parts[0], 10);
    var month = parseInt(parts[1], 10) - 1;
    var year = parseInt(parts[2], 10);
    
    var eventDate = new Date(year, month, day, 10, 0, 0);
    if (isNaN(eventDate.getTime())) continue;
    
    var title = "💰 Купон: " + name + " (+" + cleanSum + " ₽)";
    if (signal.indexOf("🚨") !== -1) title = "🚨 ПРОДАТЬ " + name + "!";
    if (signal.indexOf("🔥") !== -1) title = "🔥 КУПИТЬ " + name + "!";
    
    var description = generateEventDescription(name, isin, cleanSum, signal);
    
    try {
      cal.createEvent(title, eventDate, eventDate, {description: description});
      count++;
    } catch(e) {
      writeToLogSheet("ОШИБКА", "Ошибка записи события: " + e.message, 0);
    }
  }
  
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert("✅ Синхронизация завершена! Добавлено " + count + " событий.");
  writeToLogSheet("УСПЕХ", "Синхронизация с календарем: " + count + " событий", count);
}

function generateEventDescription(name, isin, cleanSum, signal) {
  var actionPlan = "";
  
  if (signal.indexOf("🚨") !== -1) {
    actionPlan = "⚠️ Бумага перегрета, доходность упала. Рекомендуется ПРОДАТЬ выпуск и зафиксировать прибыль.";
  } else if (signal.indexOf("🔥") !== -1) {
    actionPlan = "📈 Высокая доходность! Цена занижена. Рекомендуется ДОКУПИТЬ этот выпуск.";
  } else {
    actionPlan = "🟢 Стабильный плановый доход. Купоны отправляем на реинвестирование.";
  }
  
  return "📋 СВОДКА ПО ВЫПЛАТЕ:\n" +
         "• Облигация: " + name + "\n" +
         "• Код ISIN: " + isin + "\n" +
         "• Сумма чистыми: " + cleanSum + " ₽\n" +
         "• Статус YTM: " + signal + "\n\n" +
         "----------------------------------------\n" +
         actionPlan;
}
