// ============================================================
// РАСЧЕТЫ YTM И РЫНОЧНЫХ СИГНАЛОВ
// ============================================================

function calculateExactYTM(currentPricePct, faceValue, nkd, couponValue, frequency, maturityStr) {
  if (!currentPricePct || currentPricePct <= 0 || !maturityStr) return 0;
  
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  
  var maturityDate = parseDateLocal(maturityStr);
  if (!maturityDate || maturityDate <= today) return 0;
  
  var dirtyPrice = (currentPricePct / 100) * faceValue + nkd;
  var daysToMaturity = Math.ceil((maturityDate.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
  if (daysToMaturity <= 0) return 0;
  
  var monthsStep = 12 / (frequency || 2);
  var cashFlows = [];
  cashFlows.push({ date: new Date(today.getTime()), amount: -dirtyPrice });
  
  var tempDate = new Date(maturityDate.getTime());
  while (tempDate > today) {
    cashFlows.push({ date: new Date(tempDate.getTime()), amount: couponValue });
    tempDate.setMonth(tempDate.getMonth() - monthsStep);
  }
  
  if (cashFlows.length > 1) {
    cashFlows[1].amount += faceValue;
  }
  
  cashFlows.sort(function(a, b) { return a.date - b.date; });
  
  var rate = 0.15;
  for (var iteration = 0; iteration < 50; iteration++) {
    var fValue = 0;
    var fDerivative = 0;
    
    for (var i = 0; i < cashFlows.length; i++) {
      var d = Math.ceil((cashFlows[i].date - cashFlows[0].date) / (1000 * 60 * 60 * 24)) / 365;
      var expTerm = Math.pow(1 + rate, d);
      fValue += cashFlows[i].amount / expTerm;
      fDerivative -= d * cashFlows[i].amount / (expTerm * (1 + rate));
    }
    
    if (Math.abs(fDerivative) < 1e-10) break;
    
    var newRate = rate - fValue / fDerivative;
    if (Math.abs(newRate - rate) < 1e-6) {
      rate = newRate;
      break;
    }
    rate = newRate;
  }
  
  return (rate > 0 && rate < 2) ? rate : 0;
}

function generateMarketSignal(ytmPercent, maturityStr) {
  var dateLabel = maturityStr ? " до " + maturityStr : "";
  if (!ytmPercent || ytmPercent <= 0) {
    return "⚪ Нет данных YTM";
  }
  
  if (ytmPercent > 23.0) {
    return "🔥 КУПИТЬ ЕЩЕ (YTM " + ytmPercent.toFixed(2) + "%" + dateLabel + ")";
  } else if (ytmPercent >= 16.0) {
    return "🟢 Держать (YTM " + ytmPercent.toFixed(2) + "%" + dateLabel + ")";
  } else {
    return "🚨 ПРОДАТЬ/ПЕРЕЛОЖИТЬ (YTM " + ytmPercent.toFixed(2) + "%" + dateLabel + ")";
  }
}

function autoUpdateCouponDates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Облигации");
  if (!sheet) return;
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  
  var data = sheet.getRange("A2:M" + lastRow).getDisplayValues();
  var couponRange = sheet.getRange("M2:M" + lastRow);
  var outputDates = [];
  var updated = false;
  
  for (var i = 0; i < data.length; i++) {
    var rawCouponDate = data[i][12];
    var rawMaturity = data[i][6];
    var frequency = parseInt(data[i][9]) || 0;
    
    if (!rawCouponDate || rawCouponDate.indexOf("NaN") !== -1 || frequency <= 0) {
      outputDates.push([rawCouponDate]);
      continue;
    }
    
    var couponDate = parseDateLocal(rawCouponDate);
    var maturityDate = parseDateLocal(rawMaturity);
    
    if (couponDate && couponDate <= today) {
      var monthsStep = 12 / frequency;
      couponDate.setMonth(couponDate.getMonth() + Math.ceil((today - couponDate) / (1000 * 60 * 60 * 24 * 30.4 * (12/frequency))) * monthsStep);
      
      if (maturityDate && couponDate > maturityDate) {
        couponDate = maturityDate;
      }
      
      outputDates.push([Utilities.formatDate(couponDate, Session.getScriptZone(), "dd.MM.yyyy")]);
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
