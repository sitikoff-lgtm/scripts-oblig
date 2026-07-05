// ============================================================
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
// ============================================================

function parseDateLocal(dateVal) {
  if (!dateVal) return null;
  
  if (dateVal instanceof Date) {
    var cleanDate = new Date(dateVal.getTime());
    cleanDate.setHours(0, 0, 0, 0);
    return cleanDate;
  }
  
  var dateStr = String(dateVal).trim();
  if (dateStr === "" || dateStr.indexOf("NaN") !== -1) return null;
  
  var day, month, year;
  
  if (dateStr.indexOf('.') !== -1 || dateStr.indexOf('/') !== -1) {
    var separator = dateStr.indexOf('.') !== -1 ? '.' : '/';
    var parts = dateStr.split(separator);
    if (parts.length === 3) {
      day = parseInt(parts[0], 10);
      month = parseInt(parts[1], 10) - 1;
      year = parseInt(parts[2], 10);
    }
  } else if (dateStr.indexOf('-') !== -1) {
    var isoParts = dateStr.split('-');
    if (isoParts.length === 3) {
      year = parseInt(isoParts[0], 10);
      month = parseInt(isoParts[1], 10) - 1;
      day = parseInt(isoParts[2], 10);
    }
  }
  
  if (isNaN(day) || isNaN(month) || isNaN(year)) {
    var fallbackDate = new Date(dateStr);
    if (!isNaN(fallbackDate.getTime())) {
      fallbackDate.setHours(0, 0, 0, 0);
      return fallbackDate;
    }
    return null;
  }
  
  var finalResultDate = new Date(year, month, day, 0, 0, 0, 0);
  return isNaN(finalResultDate.getTime()) ? null : finalResultDate;
}

function isValidIsinFormat(isin) {
  if (!isin || typeof isin !== 'string') return false;
  
  var cleanIsin = isin.trim().toUpperCase();
  if (cleanIsin.length !== 12) return false;
  
  var firstTwo = cleanIsin.substring(0, 2);
  if (!firstTwo.match(/^[A-Z]{2}$/)) return false;
  
  var rest = cleanIsin.substring(2);
  if (!rest.match(/^[A-Z0-9]{10}$/)) return false;
  
  return true;
}