/**
 * ====================================================================
 * WARGAKOE - CORE SERVER ENTRYPOINT & UTILITY HELPER
 * By Zettbos System (ZettBOT 3.1)
 * ====================================================================
 */

function doGet(e) {
  try {
    var template = HtmlService.createTemplateFromFile('index');
    return template.evaluate()
      .setTitle('Wargakoe - Sistem Pengelolaan Warga & Iuran RT 010')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return HtmlService.createHtmlOutput('<h3>Terjadi kesalahan memuat aplikasi: ' + err.toString() + '</h3>');
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSheet_(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    setupDatabase();
    sheet = ss.getSheetByName(sheetName);
  }
  return sheet;
}

function generateSequentialId_(prefix, sheetName) {
  var sheet = getSheet_(sheetName);
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return prefix + '-0001';
  }
  var idCol = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
  var maxNum = 0;
  for (var i = 0; i < idCol.length; i++) {
    var rawId = idCol[i][0];
    if (rawId && rawId.indexOf(prefix + '-') === 0) {
      var numPart = parseInt(rawId.replace(prefix + '-', ''), 10);
      if (!isNaN(numPart) && numPart > maxNum) {
        maxNum = numPart;
      }
    }
  }
  var nextNum = maxNum + 1;
  var padded = ('0000' + nextNum).slice(-4);
  return prefix + '-' + padded;
}

function getNextKuitansiNo() {
  try {
    var todayStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd');
    var sheet = getSheet_('IURAN');
    var lastRow = sheet.getLastRow();
    
    var countToday = 0;
    if (lastRow > 1) {
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
      var kwiIdx = headers.indexOf('No_Kuitansi');
      if (kwiIdx !== -1) {
        var data = sheet.getRange(2, kwiIdx + 1, lastRow - 1, 1).getDisplayValues();
        for (var i = 0; i < data.length; i++) {
          var kwi = data[i][0];
          if (kwi && kwi.indexOf(todayStr) === 0) {
            var seq = parseInt(kwi.substring(8), 10);
            if (!isNaN(seq) && seq > countToday) {
              countToday = seq;
            }
          }
        }
      }
    }
    
    var nextSeq = countToday + 1;
    var seqStr = ('0' + nextSeq).slice(-2);
    return { success: true, noKuitansi: todayStr + seqStr };
  } catch (error) {
    return { success: false, message: 'Gagal generate No Kuitansi: ' + error.toString() };
  }
}
