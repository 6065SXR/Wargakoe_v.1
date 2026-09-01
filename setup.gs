/**
 * ====================================================================
 * WARGAKOE - DATABASE SETUP & SAFE MIGRATION WITH DATA REPAIR
 * By Zettbos System (ZettBOT 3.1)
 * ====================================================================
 */

var DB_SCHEMA = {
  USERS: {
    name: 'USERS',
    headers: [
      'No_KK',
      'No_KTP',
      'Nama',
      'No_HP',
      'Password',
      'Role',
      'Status_User',
      'Tanggal_Lahir',
      'Umur',
      'Alamat',
      'Jenis_Kelamin',
      'Status_Keluarga',
      'Status_Rumah',
      'Pendidikan',
      'Pekerjaan',
      'Created_At'
    ]
  },
  ANGGOTA_KELUARGA: {
    name: 'ANGGOTA_KELUARGA',
    headers: [
      'ID_Anggota',
      'No_KK',
      'Nama_Anggota',
      'Hubungan_Keluarga',
      'Tanggal_Lahir',
      'Umur',
      'Jenis_Kelamin',
      'Created_At'
    ]
  },
  IURAN: {
    name: 'IURAN',
    headers: [
      'ID_Iuran',
      'No_Kuitansi',
      'No_KK',
      'Nama_Warga',
      'Bulan_Tahun',
      'Jenis_Iuran',
      'Nominal',
      'Status_Bayar',
      'Tanggal_Bayar',
      'Approved_By',
      'Created_At'
    ]
  },
  MADING: {
    name: 'MADING',
    headers: [
      'ID_Mading',
      'Judul',
      'Tanggal_Kegiatan',
      'Note',
      'Status_Publish',
      'Pembuat',
      'Created_At'
    ]
  },
  KAS_RT: {
    name: 'KAS_RT',
    headers: [
      'ID_Kas',
      'Tanggal',
      'Jenis_Kas',
      'Kategori',
      'Keterangan',
      'Nominal',
      'Bulan_Tahun',
      'Created_At'
    ]
  }
};

function setupDatabase() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var results = [];

    for (var key in DB_SCHEMA) {
      if (DB_SCHEMA.hasOwnProperty(key)) {
        var table = DB_SCHEMA[key];
        var sheet = ss.getSheetByName(table.name);

        if (!sheet) {
          sheet = ss.insertSheet(table.name);
          setupHeaders_(sheet, table.headers);
          results.push('Sheet ' + table.name + ' berhasil dibuat.');
        } else {
          migrateHeaders_(sheet, table.headers);
          results.push('Sheet ' + table.name + ' diverifikasi.');
        }
      }
    }

    // Perbaikan otomatis & pembersihan legacy data KWI-IUR-
    repairIuranSheet_(ss);
    seedInitialData_(ss);
    SpreadsheetApp.flush();
    return { success: true, message: 'Database Wargakoe berhasil diperbaiki & diinisialisasi.', details: results };
  } catch (error) {
    return { success: false, message: 'Gagal setup database: ' + error.toString() };
  }
}

function setupHeaders_(sheet, headers) {
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#221d52');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

function migrateHeaders_(sheet, expectedHeaders) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow === 0 || lastCol === 0) {
    setupHeaders_(sheet, expectedHeaders);
    return;
  }

  var currentHeaders = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  var missingHeaders = [];

  for (var i = 0; i < expectedHeaders.length; i++) {
    if (currentHeaders.indexOf(expectedHeaders[i]) === -1) {
      missingHeaders.push(expectedHeaders[i]);
    }
  }

  if (missingHeaders.length > 0) {
    var startCol = lastCol + 1;
    var newRange = sheet.getRange(1, startCol, 1, missingHeaders.length);
    newRange.setValues([missingHeaders]);
    newRange.setBackground('#221d52');
    newRange.setFontColor('#ffffff');
    newRange.setFontWeight('bold');
    newRange.setHorizontalAlignment('center');
    sheet.autoResizeColumns(startCol, missingHeaders.length);
  }
}

function repairIuranSheet_(ss) {
  try {
    var sheet = ss.getSheetByName('IURAN');
    if (!sheet) return;

    var canonicalHeaders = DB_SCHEMA.IURAN.headers;
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    if (lastRow <= 1) {
      setupHeaders_(sheet, canonicalHeaders);
      return;
    }

    var fullData = sheet.getRange(1, 1, lastRow, Math.max(lastCol, canonicalHeaders.length)).getDisplayValues();
    var currentHeaders = fullData[0];

    var repairedRows = [canonicalHeaders];

    for (var i = 1; i < fullData.length; i++) {
      var row = fullData[i];
      var idIuran = row[0] || '';
      var colB = row[1] || '';

      // Hapus & lewati total data legacy yang memiliki nomor KWI-IUR-
      if (idIuran.indexOf('KWI-IUR-') === 0 || colB.indexOf('KWI-IUR-') === 0) {
        continue;
      }

      if (!idIuran || idIuran.trim() === '') continue;

      var noKwi = '', noKk = '', nama = '', bulan = '', jenis = '', nominal = '', status = '', tglBayar = '', approvedBy = '', createdAt = '';

      var colK = row[10] || '';

      if (colB.length >= 15 && !isNaN(colB) && !colB.startsWith('2025') && !colB.startsWith('2026') && !colB.startsWith('2027')) {
        noKk = colB;
        nama = row[2] || '';
        bulan = row[3] || '';
        jenis = row[4] || '';
        nominal = row[5] || '';
        status = row[6] || '';
        tglBayar = row[7] || '';
        approvedBy = row[8] || '';
        createdAt = row[9] || '';
        noKwi = colK || ('KWI-' + idIuran);
      } else {
        noKwi = colB || row[currentHeaders.indexOf('No_Kuitansi')] || '';
        noKk = row[2] || row[currentHeaders.indexOf('No_KK')] || '';
        nama = row[3] || row[currentHeaders.indexOf('Nama_Warga')] || '';
        bulan = row[4] || row[currentHeaders.indexOf('Bulan_Tahun')] || '';
        jenis = row[5] || row[currentHeaders.indexOf('Jenis_Iuran')] || '';
        nominal = row[6] || row[currentHeaders.indexOf('Nominal')] || '';
        status = row[7] || row[currentHeaders.indexOf('Status_Bayar')] || '';
        tglBayar = row[8] || row[currentHeaders.indexOf('Tanggal_Bayar')] || '';
        approvedBy = row[9] || row[currentHeaders.indexOf('Approved_By')] || '';
        createdAt = row[10] || row[currentHeaders.indexOf('Created_At')] || '';
      }

      // Normalisasi status
      status = status.trim();
      if (status === 'Menunggu Approval' || status === 'Menunggu' || status === 'Pending') {
        status = 'Menunggu';
      } else if (status === 'Lunas' || status === 'Sudah bayar' || status === 'Sudah Bayar') {
        status = 'Sudah bayar';
      } else if (status === 'Ditolak') {
        status = 'Ditolak';
      } else if (!status || isNaN(status) === false) {
        status = 'Menunggu';
      }

      repairedRows.push([
        idIuran, noKwi, noKk, nama, bulan, jenis, nominal, status, tglBayar, approvedBy, createdAt
      ]);
    }

    sheet.clearContents();
    sheet.getRange(1, 1, repairedRows.length, canonicalHeaders.length).setValues(repairedRows);
    setupHeaders_(sheet, canonicalHeaders);
  } catch (err) {
    Logger.log('Gagal repair IURAN: ' + err.toString());
  }
}

function seedInitialData_(ss) {
  var nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
  var todayKwi = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd');
  
  var userSheet = ss.getSheetByName(DB_SCHEMA.USERS.name);
  if (userSheet && userSheet.getLastRow() <= 1) {
    var dummyUsers = [
      [
        '3171010000000001', '3171010101850001', 'Budi Santoso', '081234567890',
        'admin123', 'Super Admin', 'Pengurus RT', '15/01/1985', '39',
        'Kamp. Baru I Jl. Marga Mulya RT 010 RW 05 Halim Perdana Kusuma', 'Laki-laki', 'Suami', 'Pribadi',
        'Sarjana', 'Pegawai Negeri', nowStr
      ],
      [
        '3171010000000002', '3171010102900002', 'Siti Rahmawati', '081234567891',
        'bendahara123', 'Bendahara', 'Pengurus RT', '20/05/1990', '34',
        'Kamp. Baru I Jl. Marga Mulya RT 010 RW 05 Halim Perdana Kusuma', 'Perempuan', 'Istri', 'Pribadi',
        'Sarjana', 'Wiraswasta', nowStr
      ],
      [
        '3171010000000003', '3171010103950003', 'Agus Setiawan', '081234567892',
        'warga123', 'Warga', 'Warga', '10/08/1995', '29',
        'Kamp. Baru I Jl. Marga Mulya RT 010 RW 05 Halim Perdana Kusuma', 'Laki-laki', 'Suami', 'Kontrak',
        'Diploma', 'Pegawai Swasta', nowStr
      ],
      [
        '3171010000000004', '3171010104600004', 'Haji Abdullah', '081234567893',
        'warga123', 'Warga', 'Warga', '05/03/1960', '64',
        'Kamp. Baru I Jl. Marga Mulya RT 010 RW 05 Halim Perdana Kusuma', 'Laki-laki', 'Ayah', 'Pribadi',
        'SMA', 'Wiraswasta', nowStr
      ]
    ];
    userSheet.getRange(2, 1, dummyUsers.length, dummyUsers[0].length).setValues(dummyUsers);
  }

  var iurSheet = ss.getSheetByName(DB_SCHEMA.IURAN.name);
  if (iurSheet && iurSheet.getLastRow() <= 1) {
    var currentBulan = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'MMMM yyyy');
    var dummyIuran = [
      ['IUR-0001', todayKwi + '01', '3171010000000001', 'Budi Santoso', currentBulan, 'Iuran Kas', '50000', 'Sudah bayar', nowStr, 'Siti Rahmawati', nowStr],
      ['IUR-0002', todayKwi + '02', '3171010000000001', 'Budi Santoso', currentBulan, 'Iuran Duka', '20000', 'Sudah bayar', nowStr, 'Siti Rahmawati', nowStr],
      ['IUR-0003', todayKwi + '03', '3171010000000003', 'Agus Setiawan', currentBulan, 'Iuran Sampah', '25000', 'Menunggu', nowStr, '-', nowStr],
      ['IUR-0004', todayKwi + '04', '3171010000000003', 'Agus Setiawan', currentBulan, 'Iuran Sosial', '15000', 'Belum Bayar', '-', '-', nowStr],
      ['IUR-0005', todayKwi + '05', '3171010000000004', 'Haji Abdullah', currentBulan, 'Iuran Kas', '50000', 'Sudah bayar', nowStr, 'Siti Rahmawati', nowStr]
    ];
    iurSheet.getRange(2, 1, dummyIuran.length, dummyIuran[0].length).setValues(dummyIuran);
  }

  var kasSheet = ss.getSheetByName(DB_SCHEMA.KAS_RT.name);
  if (kasSheet && kasSheet.getLastRow() <= 1) {
    var curBulanYear = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'MMMM yyyy');
    var dummyKas = [
      ['KAS-0001', '01/09/2026', 'Pemasukan', 'Iuran Kas', 'Penerimaan Iuran Kas Bulanan RT 010', '100000', curBulanYear, nowStr],
      ['KAS-0002', '03/09/2026', 'Pengeluaran', 'Iuran Kas', 'Pembelian Alat Kebersihan Pos Ronda RT 010', '35000', curBulanYear, nowStr]
    ];
    kasSheet.getRange(2, 1, dummyKas.length, dummyKas[0].length).setValues(dummyKas);
  }
}
