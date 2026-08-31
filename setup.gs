/**
 * ====================================================================
 * WARGAKOE - DATABASE SETUP & SAFE MIGRATION
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

    seedInitialData_(ss);
    SpreadsheetApp.flush();
    return { success: true, message: 'Database Wargakoe berhasil diinisialisasi.', details: results };
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

function seedInitialData_(ss) {
  var nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
  
  var userSheet = ss.getSheetByName(DB_SCHEMA.USERS.name);
  if (userSheet && userSheet.getLastRow() <= 1) {
    var dummyUsers = [
      [
        '3171010000000001', '3171010101850001', 'Budi Santoso', '081234567890',
        'admin123', 'Super Admin', 'Pengurus RT', '15/01/1985', '39',
        'Kamp. Baru I Jl. Marga Mulya No. 01 RT 05/RW 02', 'Laki-laki', 'Suami', 'Pribadi',
        'Sarjana', 'Pegawai Negeri', nowStr
      ],
      [
        '3171010000000002', '3171010102900002', 'Siti Rahmawati', '081234567891',
        'bendahara123', 'Bendahara', 'Pengurus RT', '20/05/1990', '34',
        'Kamp. Baru I Jl. Marga Mulya No. 04 RT 05/RW 02', 'Perempuan', 'Istri', 'Pribadi',
        'Sarjana', 'Wiraswasta', nowStr
      ],
      [
        '3171010000000003', '3171010103950003', 'Agus Setiawan', '081234567892',
        'warga123', 'Warga', 'Warga', '10/08/1995', '29',
        'Jl. Anggrek No. 12 RT 05/RW 02', 'Laki-laki', 'Suami', 'Kontrak',
        'Diploma', 'Pegawai Swasta', nowStr
      ],
      [
        '3171010000000004', '3171010104600004', 'Haji Abdullah', '081234567893',
        'warga123', 'Warga', 'Warga', '05/03/1960', '64',
        'Kamp. Baru I Jl. Marga Mulya No. 08 RT 05/RW 02', 'Laki-laki', 'Ayah', 'Pribadi',
        'SMA', 'Wiraswasta', nowStr
      ]
    ];
    userSheet.getRange(2, 1, dummyUsers.length, dummyUsers[0].length).setValues(dummyUsers);
  }

  var agtSheet = ss.getSheetByName(DB_SCHEMA.ANGGOTA_KELUARGA.name);
  if (agtSheet && agtSheet.getLastRow() <= 1) {
    var dummyAnggota = [
      ['AGT-0001', '3171010000000001', 'Dewi Lestari', 'Istri', '12/04/1988', '36', 'Perempuan', nowStr],
      ['AGT-0002', '3171010000000001', 'Rian Santoso', 'Anak ke-1', '08/09/2012', '12', 'Laki-laki', nowStr],
      ['AGT-0003', '3171010000000001', 'Balita Alika', 'Anak ke-2', '15/02/2023', '1', 'Perempuan', nowStr],
      ['AGT-0004', '3171010000000003', 'Nina Marlina', 'Istri', '22/11/1997', '27', 'Perempuan', nowStr],
      ['AGT-0005', '3171010000000003', 'Rizky Pratama', 'Anak ke-1', '14/06/2021', '3', 'Laki-laki', nowStr]
    ];
    agtSheet.getRange(2, 1, dummyAnggota.length, dummyAnggota[0].length).setValues(dummyAnggota);
  }

  var iurSheet = ss.getSheetByName(DB_SCHEMA.IURAN.name);
  if (iurSheet && iurSheet.getLastRow() <= 1) {
    var currentBulan = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'MMMM yyyy');
    var dummyIuran = [
      ['IUR-0001', '3171010000000001', 'Budi Santoso', currentBulan, 'Iuran RT', '25000', 'Lunas', nowStr, 'Siti Rahmawati', nowStr],
      ['IUR-0002', '3171010000000001', 'Budi Santoso', currentBulan, 'Iuran Duka', '15000', 'Lunas', nowStr, 'Siti Rahmawati', nowStr],
      ['IUR-0003', '3171010000000003', 'Agus Setiawan', currentBulan, 'Iuran Sampah', '25000', 'Menunggu Approval', nowStr, '-', nowStr],
      ['IUR-0004', '3171010000000003', 'Agus Setiawan', currentBulan, 'Iuran Sosial', '15000', 'Belum Bayar', '-', '-', nowStr],
      ['IUR-0005', '3171010000000004', 'Haji Abdullah', currentBulan, 'Iuran Lain-lain', '50000', 'Lunas', nowStr, 'Siti Rahmawati', nowStr]
    ];
    iurSheet.getRange(2, 1, dummyIuran.length, dummyIuran[0].length).setValues(dummyIuran);
  }

  var madingSheet = ss.getSheetByName(DB_SCHEMA.MADING.name);
  if (madingSheet && madingSheet.getLastRow() <= 1) {
    var dummyMading = [
      ['MAD-0001', 'Kerja Bakti Bulanan RT 05', '10 September 2026', 'Diharapkan seluruh warga hadir membawa alat kebersihan masing-masing di pos ronda utama pukul 07:00 WIB.', 'Published', 'Budi Santoso (Ketua RT)', nowStr],
      ['MAD-0002', 'Rapat Pleno Warga & Pembahasan Kas', '25 September 2026', 'Pertemuan rutin bulanan bertempat di balai warga lantai 1 pukul 19:30 WIB.', 'Published', 'Siti Rahmawati (Bendahara)', nowStr]
    ];
    madingSheet.getRange(2, 1, dummyMading.length, dummyMading[0].length).setValues(dummyMading);
  }

  var kasSheet = ss.getSheetByName(DB_SCHEMA.KAS_RT.name);
  if (kasSheet && kasSheet.getLastRow() <= 1) {
    var curBulanYear = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'MMMM yyyy');
    var dummyKas = [
      ['KAS-0001', '01/08/2026', 'Pemasukan', 'Iuran RT', 'Penerimaan Iuran Bulanan Warga RT 05', '1250000', curBulanYear, nowStr],
      ['KAS-0002', '03/08/2026', 'Pemasukan', 'Donasi', 'Sumbangan Donatur Kegiatan RT', '500000', curBulanYear, nowStr],
      ['KAS-0003', '05/08/2026', 'Pengeluaran', 'Kebersihan', 'Pembelian Alat Kebersihan & Sapu Pos Ronda', '150000', curBulanYear, nowStr],
      ['KAS-0004', '12/08/2026', 'Pengeluaran', 'Fasilitas', 'Perbaikan Lampu Jalan Gang Utama RT 05', '200000', curBulanYear, nowStr]
    ];
    kasSheet.getRange(2, 1, dummyKas.length, dummyKas[0].length).setValues(dummyKas);
  }
}
