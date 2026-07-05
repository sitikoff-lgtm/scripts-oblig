// ============================================================
// СИСТЕМА ЛОГИРОВАНИЯ
// ============================================================

function writeToLogSheet(status, message, count) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName("Логи");
    
    if (!logSheet) {
      logSheet = ss.insertSheet("Логи");
      logSheet.getRange("A1:D1").setValues([["Время (МСК)", "Статус", "Сообщение / Ошибка", "Обновлено бумаг"]])
               .setFontWeight("bold").setBackground("#f3f3f3");
      logSheet.setFrozenRows(1);
      logSheet.setColumnWidth(3, 400);
    }
    
    var timestamp = Utilities.formatDate(new Date(), "GMT+3", "dd.MM.yyyy HH:mm:ss");
    var newRow = [timestamp, status, message, count || 0];
    
    logSheet.insertRowBefore(2);
    logSheet.getRange(2, 1, 1, 4).setValues([newRow]);
    
    var statusCell = logSheet.getRange(2, 2);
    if (status === "УСПЕХ") statusCell.setBackground("#d4edda").setFontColor("#155724");
    if (status === "ОШИБКА") statusCell.setBackground("#f8d7da").setFontColor("#721c24");
    if (status === "ИНФО") statusCell.setBackground("#cce5ff").setFontColor("#004085");
    if (status === "ПРЕДУПРЕЖДЕНИЕ") statusCell.setBackground("#fff3cd").setFontColor("#856404");
    
    // Оставляем только 100 последних записей
    var maxRows = 100;
    var totalRows = logSheet.getLastRow();
    if (totalRows > maxRows) {
      logSheet.deleteRows(maxRows + 1, totalRows - maxRows);
    }
    
    SpreadsheetApp.flush();
  } catch(e) {
    Logger.log("Ошибка логирования: " + e.message);
  }
}

function clearLogs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("Логи");
  if (!logSheet) return;
  
  var lastRow = logSheet.getLastRow();
  if (lastRow > 1) {
    logSheet.deleteRows(2, lastRow - 1);
    SpreadsheetApp.getUi().alert("✅ Логи очищены.");
  }
}

function checkPortfolioHealth() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("Логи");
  if (!logSheet || logSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert("ℹ️ Логов пока нет.");
    return;
  }
  
  var logs = logSheet.getRange(2, 2, logSheet.getLastRow() - 1, 2).getValues();
  var stats = { УСПЕХ: 0, ОШИБКА: 0, ПРЕДУПРЕЖДЕНИЕ: 0, ИНФО: 0 };
  
  for (var i = 0; i < logs.length; i++) {
    var status = logs[i][0];
    if (stats.hasOwnProperty(status)) stats[status]++;
  }
  
  var message = "📊 **Статус портфеля**\n\n";
  message += "✅ Успешно: " + stats.УСПЕХ + "\n";
  message += "⚠️ Предупреждений: " + stats.ПРЕДУПРЕЖДЕНИЕ + "\n";
  message += "❌ Ошибок: " + stats.ОШИБКА + "\n";
  message += "ℹ️ Информации: " + stats.ИНФО + "\n";
  
  var health = stats.ОШИБКА > 0 ? "🔴 Требуется внимание!" : "🟢 Отлично";
  message += "\n💚 Состояние: " + health;
  
  SpreadsheetApp.getUi().alert(message);
}

function showLastLogs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("Логи");
  if (!logSheet || logSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert("ℹ️ Логов пока нет.");
    return;
  }
  
  var lastRow = logSheet.getLastRow();
  var startRow = Math.max(2, lastRow - 9);
  var logs = logSheet.getRange(startRow, 1, lastRow - startRow + 1, 4).getValues();
  
  var message = "📋 **Последние события:**\n\n";
  for (var i = 0; i < logs.length; i++) {
    var emoji = logs[i][1] === "УСПЕХ" ? "✅" : 
                logs[i][1] === "ОШИБКА" ? "❌" : 
                logs[i][1] === "ПРЕДУПРЕЖДЕНИЕ" ? "⚠️" : "ℹ️";
    message += emoji + " " + logs[i][0] + " | " + logs[i][1] + "\n   " + logs[i][2] + "\n\n";
  }
  
  SpreadsheetApp.getUi().alert(message);
}
