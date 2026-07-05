// ============================================================
// ГЛАВНЫЙ ФАЙЛ: МЕНЮ, ОРКЕСТРАТОР И ТРИГГЕРЫ
// ============================================================

function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();
    var menu = ui.createMenu('💼 ');
    menu.addItem('🔄 1. Обновить календарь и цены', 'startUniqueBondCalendarGeneration2026');
    menu.addItem('📅 2. Синхронизировать с Google Календарем', 'syncPaymentsWithGoogleCalendar');
    menu.addSeparator();
    menu.addItem('⏰ 3. Включить автообновление каждый день', 'setupDailyAutomatedTriggers');
    menu.addSeparator();
    
    var logsSubMenu = ui.createMenu('📋 Логи');
    logsSubMenu.addItem('📊 Проверить статус портфеля', 'checkPortfolioHealth');
    logsSubMenu.addItem('🔍 Показать последние логи', 'showLastLogs');
    logsSubMenu.addSeparator();
    logsSubMenu.addItem('🗑️ Очистить все логи', 'clearLogs');
    menu.addSubMenu(logsSubMenu);
    menu.addToUi();
  } catch(e) {
    Logger.log("Не удалось добавить меню: " + e.message);
  }
}

function startUniqueBondCalendarGeneration2026() {
  writeToLogSheet("ИНФО", "🚀 Запуск обновления портфеля...", 0);
  try {
    cleanBksVerticalText();
    writeToLogSheet("ИНФО", "✅ Парсинг БКС выполнен", 0);
    autoAddNewBondsFromCalendar();
    writeToLogSheet("ИНФО", "✅ Проверка новых бумаг завершена", 0);
    updateAllBondValuesFromCache();
    writeToLogSheet("ИНФО", "✅ Цены и НКД обновлены", 0);
    prepareDestSheetStructure();
    writeToLogSheet("УСПЕХ", "✅ Календарь выплат сгенерирован", 0);
  } catch(e) {
    writeToLogSheet("ОШИБКА", "❌ Критический сбой: " + e.message, 0);
    throw e;
  }
}

function setupDailyAutomatedTriggers() {
  var triggerName = 'startUniqueBondCalendarGeneration2026';
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === triggerName) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger(triggerName).timeBased().everyDays(1).atHour(20).create();
  writeToLogSheet("ИНФО", "Автообновление включено (20:00 МСК)", 0);
  SpreadsheetApp.getUi().alert("⏰ Автоматизация включена!");
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
  var message = "📋 **ПОСЛЕДНИЕ " + logs.length + " СОБЫТИЙ:**\n\n";
  for (var i = 0; i < logs.length; i++) {
    var emoji = logs[i][1] === "УСПЕХ" ? "✅" : logs[i][1] === "ОШИБКА" ? "❌" : logs[i][1] === "ПРЕДУПРЕЖДЕНИЕ" ? "⚠️" : "ℹ️";
    message += emoji + " " + logs[i][0] + " | " + logs[i][1] + "\n   " + logs[i][2] + "\n\n";
  }
  SpreadsheetApp.getUi().alert(message);
}

function checkPortfolioHealth() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("Логи");
  if (!logSheet || logSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert("ℹ️ Логов пока нет.");
    return;
  }
  var allLogs = logSheet.getRange(2, 2, logSheet.getLastRow() - 1, 2).getValues();
  var stats = { УСПЕХ: 0, ОШИБКА: 0, ПРЕДУПРЕЖДЕНИЕ: 0, ИНФО: 0 };
  var lastErrors = [];
  for (var i = 0; i < allLogs.length; i++) {
    var status = allLogs[i][0];
    if (stats.hasOwnProperty(status)) stats[status]++;
    if (status === "ОШИБКА" && lastErrors.length < 3) lastErrors.push(allLogs[i][1]);
  }
  var message = "📊 **СТАТУС ПОРТФЕЛЯ**\n\n";
  message += "✅ Успешно: " + stats.УСПЕХ + "\n";
  message += "⚠️ Предупреждений: " + stats.ПРЕДУПРЕЖДЕНИЕ + "\n";
  message += "❌ Ошибок: " + stats.ОШИБКА + "\n";
  message += "ℹ️ Информации: " + stats.ИНФО + "\n\n";
  var health = stats.ОШИБКА > 0 ? "🔴 Требуется внимание!" : "🟢 Отлично";
  message += "💚 Состояние: " + health;
  SpreadsheetApp.getUi().alert(message);
}

function clearLogs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("Логи");
  if (!logSheet || logSheet.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert("ℹ️ Логи уже пусты.");
    return;
  }
  var response = SpreadsheetApp.getUi().alert("⚠️ Подтверждение очистки", "Удалить все логи?", SpreadsheetApp.getUi().ButtonSet.YES_NO);
  if (response === SpreadsheetApp.getUi().Button.YES) {
    logSheet.deleteRows(2, logSheet.getLastRow() - 1);
    writeToLogSheet("ИНФО", "Логи очищены", 0);
    SpreadsheetApp.getUi().alert("✅ Логи очищены.");
  }
}