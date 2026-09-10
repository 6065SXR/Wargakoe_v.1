/**
 * ====================================================================
 * WARGAKOE - SUPER ADMIN PANEL MODULE
 * Modul cadangan (Backup JSON), pemulihan (Restore), dan reset database
 * terproteksi untuk akun Super Admin RT 010 / RW 05 Halim.
 * Mendukung reset per-sheet, multi-sheet (checkbox), maupun reset all.
 * ====================================================================
 */

var ADMIN_DATABASE_SHEETS = [
  'USERS',
  'ANGGOTA_KELUARGA',
  'IURAN',
  'KAS_RT',
  'MADING',
  'ASET',
  'PEMINJAMAN_ASET',
  'JADWAL_RONDA',
  'ABSENSI_RONDA',
  'PENGADUAN'
];

/**
 * Memvalidasi apakah user pemanggil adalah Super Admin aktif
 */
function verifySuperAdminAccess_(ss, adminNoKk, adminNoHp) {
  var userSheet = ss.getSheetByName('USERS');
  if (!userSheet || userSheet.getLastRow() <= 1) return false;

  var data = userSheet.getDataRange().getDisplayValues();
  var headers = data[0];
  var kkIdx = headers.indexOf('No_KK');
  var hpIdx = headers.indexOf('No_HP');
  var roleIdx = headers.indexOf('Role');

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var matchKk = (kkIdx !== -1 && String(row[kkIdx]).trim() === String(adminNoKk).trim());
    var matchHp = (hpIdx !== -1 && String(row[hpIdx]).trim() === String(adminNoHp).trim());
    var isSuper = (roleIdx !== -1 && String(row[roleIdx]).trim().toLowerCase() === 'super admin');

    if ((matchKk || matchHp) && isSuper) {
      return true;
    }
  }
  return false;
}

/**
 * 1. BACKUP DATABASE (Ekspor seluruh 10 sheet ke struktur JSON)
 */
function backupDatabase(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    payload = payload || {};
    var adminNoKk = payload.adminNoKk || '';
    var adminNoHp = payload.adminNoHp || '';

    if (!verifySuperAdminAccess_(ss, adminNoKk, adminNoHp)) {
      return { success: false, message: 'Akses ditolak! Fitur ini hanya untuk Super Admin terverifikasi.' };
    }

    var nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var fileTimestamp = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd_HHmmss');

    var backupPackage = {
      app: 'Wargakoe - RT 010 / RW 05 Halim Perdana Kusuma',
      version: '3.1.0',
      exportedAt: nowStr,
      exportedBy: payload.adminNama || 'Super Admin',
      sheetsCount: ADMIN_DATABASE_SHEETS.length,
      sheets: {}
    };

    for (var i = 0; i < ADMIN_DATABASE_SHEETS.length; i++) {
      var sheetName = ADMIN_DATABASE_SHEETS[i];
      var sheet = ss.getSheetByName(sheetName);

      if (sheet && sheet.getLastRow() > 0) {
        var rawData = sheet.getDataRange().getDisplayValues();
        backupPackage.sheets[sheetName] = {
          headers: rawData[0] || [],
          rowCount: rawData.length - 1,
          rows: rawData.slice(1)
        };
      } else {
        backupPackage.sheets[sheetName] = {
          headers: [],
          rowCount: 0,
          rows: []
        };
      }
    }

    return {
      success: true,
      backupJson: JSON.stringify(backupPackage, null, 2),
      filename: 'backup_wargakoe_rt010_' + fileTimestamp + '.json',
      message: 'Backup database berhasil dibuat! Siap diunduh.'
    };
  } catch (error) {
    return { success: false, message: 'Gagal membuat backup: ' + error.toString() };
  }
}

/**
 * 2. RESTORE DATABASE (Memulihkan seluruh sheet dari file cadangan JSON)
 * Bebas duplikasi karena seluruh sheet target dikosongkan (clearContents) sebelum ditimpa data baru
 */
function restoreDatabase(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    payload = payload || {};
    var adminNoKk = payload.adminNoKk || '';
    var adminNoHp = payload.adminNoHp || '';
    var backupJsonStr = payload.backupJsonStr || '';

    if (!verifySuperAdminAccess_(ss, adminNoKk, adminNoHp)) {
      return { success: false, message: 'Akses ditolak! Hanya Super Admin yang berhak merestore database.' };
    }

    if (!backupJsonStr) {
      return { success: false, message: 'File cadangan JSON tidak valid atau kosong!' };
    }

    var parsedPackage;
    try {
      parsedPackage = JSON.parse(backupJsonStr);
    } catch (e) {
      return { success: false, message: 'Format file JSON rusak dan tidak dapat dibaca.' };
    }

    if (!parsedPackage.sheets) {
      return { success: false, message: 'Format backup tidak cocok dengan struktur database Wargakoe.' };
    }

    var restoredSheets = [];

    for (var i = 0; i < ADMIN_DATABASE_SHEETS.length; i++) {
      var sheetName = ADMIN_DATABASE_SHEETS[i];
      var sheetData = parsedPackage.sheets[sheetName];

      if (sheetData && sheetData.headers && sheetData.headers.length > 0) {
        var targetSheet = ss.getSheetByName(sheetName);
        if (!targetSheet) {
          targetSheet = ss.insertSheet(sheetName);
        }

        // MENGHINDARI DUPLIKASI DATA: Kosongkan seluruh isi lama terlebih dahulu
        targetSheet.clearContents();

        var tableData = [sheetData.headers].concat(sheetData.rows || []);
        var range = targetSheet.getRange(1, 1, tableData.length, sheetData.headers.length);
        range.setValues(tableData);

        // Memformat ulang header tabel agar rapi
        var headerRange = targetSheet.getRange(1, 1, 1, sheetData.headers.length);
        headerRange.setBackground('#221d52');
        headerRange.setFontColor('#ffffff');
        headerRange.setFontWeight('bold');
        headerRange.setHorizontalAlignment('center');
        targetSheet.setFrozenRows(1);

        restoredSheets.push(sheetName);
      }
    }

    SpreadsheetApp.flush();

    return {
      success: true,
      message: 'Database berhasil dipulihkan (Restore)! Total ' + restoredSheets.length + ' sheet telah disinkronkan kembali tanpa duplikasi.'
    };
  } catch (error) {
    return { success: false, message: 'Gagal merestore database: ' + error.toString() };
  }
}

/**
 * 3. RESET DATABASE (Mereset per-sheet, multi-sheet checkbox, atau reset all dengan proteksi akun Super Admin)
 */
function resetDatabase(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    payload = payload || {};
    var adminNoKk = String(payload.adminNoKk || '').trim();
    var adminNoHp = String(payload.adminNoHp || '').trim();
    var confirmKeyword = String(payload.confirmKeyword || '').trim();
    var mode = String(payload.mode || 'SINGLE').toUpperCase(); // 'SINGLE' (atau 'CHECKBOX') atau 'ALL'
    var targetSheets = payload.targetSheets || [];

    // Kompatibilitas mundur jika dikirim string tunggal targetSheet
    if (payload.targetSheet && targetSheets.length === 0) {
      targetSheets = [payload.targetSheet];
    }

    if (!verifySuperAdminAccess_(ss, adminNoKk, adminNoHp)) {
      return { success: false, message: 'Akses ditolak! Hanya Super Admin yang berhak melakukan reset data.' };
    }

    if (confirmKeyword !== 'RESET') {
      return { success: false, message: 'Konfirmasi keamanan tidak sesuai! Ketik teks "RESET" dengan huruf kapital.' };
    }

    var resetResults = [];

    if (mode === 'ALL') {
      for (var i = 0; i < ADMIN_DATABASE_SHEETS.length; i++) {
        var sName = ADMIN_DATABASE_SHEETS[i];
        resetSingleSheetWithProtection_(ss, sName, adminNoKk);
        resetResults.push(sName);
      }
    } else {
      if (!Array.isArray(targetSheets) || targetSheets.length === 0) {
        return { success: false, message: 'Pilih minimal satu sheet yang akan di-reset!' };
      }

      for (var j = 0; j < targetSheets.length; j++) {
        var sheetTarget = String(targetSheets[j]).trim();
        if (ADMIN_DATABASE_SHEETS.indexOf(sheetTarget) !== -1) {
          resetSingleSheetWithProtection_(ss, sheetTarget, adminNoKk);
          resetResults.push(sheetTarget);
        }
      }
    }

    SpreadsheetApp.flush();

    var msg = (mode === 'ALL')
      ? 'Seluruh database (10 sheet) berhasil di-reset ke pengaturan awal! Akun Super Admin tetap diproteksi dan aman.'
      : 'Sebanyak ' + resetResults.length + ' sheet (' + resetResults.join(', ') + ') berhasil di-reset dengan aman.';

    return {
      success: true,
      message: msg,
      details: resetResults
    };
  } catch (error) {
    return { success: false, message: 'Gagal melakukan reset database: ' + error.toString() };
  }
}

/**
 * Helper internal untuk mengosongkan sheet tertentu dengan perlindungan akun Super Admin
 */
function resetSingleSheetWithProtection_(ss, sheetName, adminNoKk) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow <= 1) return; // Sudah kosong hanya ada baris header

  if (sheetName === 'USERS') {
    // PROTEKSI MAKSIMAL: Pertahankan baris akun Super Admin dan pemanggil
    var rawData = sheet.getDataRange().getDisplayValues();
    var headers = rawData[0];
    var roleIdx = headers.indexOf('Role');
    var kkIdx = headers.indexOf('No_KK');

    var preservedRows = [headers];

    for (var r = 1; r < rawData.length; r++) {
      var row = rawData[r];
      var isSuper = (roleIdx !== -1 && String(row[roleIdx]).trim().toLowerCase() === 'super admin');
      var isCaller = (kkIdx !== -1 && String(row[kkIdx]).trim() === adminNoKk);

      if (isSuper || isCaller) {
        preservedRows.push(row);
      }
    }

    sheet.clearContents();
    sheet.getRange(1, 1, preservedRows.length, headers.length).setValues(preservedRows);

    // Format ulang header
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#221d52');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
  } else {
    // Sheet selain USERS: hapus seluruh baris transaksi dan pertahankan baris 1 (headers)
    sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  }
}
