/**
 * ====================================================================
 * WARGAKOE - BACKEND LOGIC & CRUD OPERATIONS WITH DYNAMIC HEADERS
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

function loginUser(noHp, password) {
  try {
    if (!noHp || !password) {
      return { success: false, message: 'Nomor Handphone dan Password wajib diisi!' };
    }

    var sheet = getSheet_('USERS');
    var data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) {
      return { success: false, message: 'Data pengguna tidak ditemukan. Silakan registrasi terlebih dahulu.' };
    }

    var headers = data[0];
    var hpIdx = headers.indexOf('No_HP');
    var passIdx = headers.indexOf('Password');

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[hpIdx] === noHp.trim() && row[passIdx] === password.trim()) {
        var userObj = {
          noKk: row[headers.indexOf('No_KK')],
          noKtp: row[headers.indexOf('No_KTP')],
          nama: row[headers.indexOf('Nama')],
          noHp: row[hpIdx],
          role: row[headers.indexOf('Role')],
          statusUser: row[headers.indexOf('Status_User')],
          tanggalLahir: row[headers.indexOf('Tanggal_Lahir')],
          umur: row[headers.indexOf('Umur')],
          alamat: row[headers.indexOf('Alamat')],
          jenisKelamin: row[headers.indexOf('Jenis_Kelamin')],
          statusKeluarga: row[headers.indexOf('Status_Keluarga')],
          statusRumah: row[headers.indexOf('Status_Rumah')],
          pendidikan: row[headers.indexOf('Pendidikan')],
          pekerjaan: row[headers.indexOf('Pekerjaan')]
        };
        return { success: true, message: 'Login berhasil! Selamat datang, ' + userObj.nama, user: userObj };
      }
    }

    return { success: false, message: 'Nomor HP atau Password salah!' };
  } catch (error) {
    return { success: false, message: 'Gagal login: ' + error.toString() };
  }
}

function registerUser(formData) {
  try {
    var sheet = getSheet_('USERS');
    var data = sheet.getDataRange().getDisplayValues();
    var headers = data[0];

    var kkIdx = headers.indexOf('No_KK');
    var ktpIdx = headers.indexOf('No_KTP');
    var hpIdx = headers.indexOf('No_HP');

    for (var i = 1; i < data.length; i++) {
      if (data[i][kkIdx] === String(formData.noKk).trim()) {
        return { success: false, message: 'Nomor KK sudah terdaftar di sistem!' };
      }
      if (data[i][ktpIdx] === String(formData.noKtp).trim()) {
        return { success: false, message: 'Nomor KTP sudah terdaftar di sistem!' };
      }
      if (data[i][hpIdx] === String(formData.noHp).trim()) {
        return { success: false, message: 'Nomor HP sudah terdaftar. Silakan login!' };
      }
    }

    var nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var defaultRole = formData.statusUser === 'Pengurus RT' ? (formData.role || 'Pengurus RT') : 'Warga';

    var newRow = [
      String(formData.noKk || '').trim(),
      String(formData.noKtp || '').trim(),
      String(formData.nama || '').trim(),
      String(formData.noHp || '').trim(),
      String(formData.password || '').trim(),
      defaultRole,
      String(formData.statusUser || 'Warga').trim(),
      String(formData.tanggalLahir || '').trim(),
      String(formData.umur || '0').trim(),
      String(formData.alamat || '').trim(),
      String(formData.jenisKelamin || 'Laki-laki').trim(),
      String(formData.statusKeluarga || 'Suami').trim(),
      String(formData.statusRumah || 'Pribadi').trim(),
      String(formData.pendidikan || 'SMA').trim(),
      String(formData.pekerjaan || 'Wiraswasta').trim(),
      nowStr
    ];

    sheet.appendRow(newRow);
    SpreadsheetApp.flush();

    return { success: true, message: 'Registrasi berhasil! Silakan login.' };
  } catch (error) {
    return { success: false, message: 'Gagal registrasi: ' + error.toString() };
  }
}

function getDashboardSummary(userRole, userNoKk) {
  try {
    var userSheet = getSheet_('USERS');
    var agtSheet = getSheet_('ANGGOTA_KELUARGA');
    var iurSheet = getSheet_('IURAN');
    var madingSheet = getSheet_('MADING');

    var userData = userSheet.getDataRange().getDisplayValues();
    var agtData = agtSheet.getDataRange().getDisplayValues();
    var iurData = iurSheet.getDataRange().getDisplayValues();
    var madingData = madingSheet.getDataRange().getDisplayValues();

    var totalKk = Math.max(0, userData.length - 1);
    var totalWarga = totalKk + Math.max(0, agtData.length - 1);

    var stats = {
      balita: 0,
      remaja: 0,
      dewasa: 0,
      lansia: 0,
      totalWarga: totalWarga,
      totalKk: totalKk,
      kasTerkumpul: 0,
      kasRtBulanIni: 0,
      kasKematianBulanIni: 0,
      iuranSampahBulanIni: 0,
      iuranSosialBulanIni: 0,
      iuranLainBulanIni: 0,
      totalTerkumpul: 0,
      totalMenunggu: 0,
      totalBelumBayar: 0,
      unpaidDuesList: [],
      madingTerbaru: []
    };

    var uHeaders = userData[0];
    var uUmurIdx = uHeaders.indexOf('Umur');
    for (var u = 1; u < userData.length; u++) {
      var umur = parseInt(userData[u][uUmurIdx], 10) || 0;
      if (umur < 5) stats.balita++;
      else if (umur <= 17) stats.remaja++;
      else if (umur <= 59) stats.dewasa++;
      else stats.lansia++;
    }

    if (agtData.length > 1) {
      var aHeaders = agtData[0];
      var aUmurIdx = aHeaders.indexOf('Umur');
      for (var a = 1; a < agtData.length; a++) {
        var aUmur = parseInt(agtData[a][aUmurIdx], 10) || 0;
        if (aUmur < 5) stats.balita++;
        else if (aUmur <= 17) stats.remaja++;
        else if (aUmur <= 59) stats.dewasa++;
        else stats.lansia++;
      }
    }

    if (iurData.length > 1) {
      var iHeaders = iurData[0];
      var iKkIdx = iHeaders.indexOf('No_KK');
      var iJenisIdx = iHeaders.indexOf('Jenis_Iuran');
      var iNominalIdx = iHeaders.indexOf('Nominal');
      var iStatusIdx = iHeaders.indexOf('Status_Bayar');

      for (var k = 1; k < iurData.length; k++) {
        var row = iurData[k];
        var idIur = row[iHeaders.indexOf('ID_Iuran')] || '';
        var kwiNo = row[iHeaders.indexOf('No_Kuitansi')] || '';

        // Lewati data legacy KWI-IUR-
        if (idIur.indexOf('KWI-IUR-') === 0 || kwiNo.indexOf('KWI-IUR-') === 0) continue;

        var nominal = parseFloat(row[iNominalIdx]) || 0;
        var status = (row[iStatusIdx] || '').trim();
        var jenis = row[iJenisIdx];
        var kk = row[iKkIdx];

        if (status === 'Sudah bayar' || status === 'Lunas') {
          stats.totalTerkumpul += nominal;
          if (jenis === 'Iuran Kas' || jenis === 'Iuran RT') {
            stats.kasTerkumpul += nominal;
            stats.kasRtBulanIni += nominal;
          } else if (jenis === 'Iuran Duka' || jenis === 'Iuran Kematian') {
            stats.kasKematianBulanIni += nominal;
          } else if (jenis === 'Iuran Sampah') {
            stats.iuranSampahBulanIni += nominal;
          } else if (jenis === 'Iuran Sosial') {
            stats.iuranSosialBulanIni += nominal;
          } else {
            stats.iuranLainBulanIni += nominal;
          }
        } else if (status === 'Menunggu' || status === 'Menunggu Approval') {
          stats.totalMenunggu += nominal;
        } else {
          stats.totalBelumBayar += nominal;
        }

        if (userRole !== 'Super Admin' && kk === userNoKk) {
          if (status !== 'Sudah bayar' && status !== 'Lunas') {
            stats.unpaidDuesList.push({
              idIuran: idIur,
              noKuitansi: kwiNo,
              namaWarga: row[iHeaders.indexOf('Nama_Warga')],
              jenisIuran: jenis,
              bulanTahun: row[iHeaders.indexOf('Bulan_Tahun')],
              nominal: nominal,
              statusBayar: status
            });
          }
        }
      }
    }

    if (madingData.length > 1) {
      var mHeaders = madingData[0];
      var mList = [];
      for (var m = 1; m < madingData.length; m++) {
        if (madingData[m][mHeaders.indexOf('Status_Publish')] === 'Published') {
          mList.push({
            idMading: madingData[m][mHeaders.indexOf('ID_Mading')],
            judul: madingData[m][mHeaders.indexOf('Judul')],
            tanggalKegiatan: madingData[m][mHeaders.indexOf('Tanggal_Kegiatan')],
            note: madingData[m][mHeaders.indexOf('Note')],
            pembuat: madingData[m][mHeaders.indexOf('Pembuat')],
            createdAt: madingData[m][mHeaders.indexOf('Created_At')]
          });
        }
      }
      stats.madingTerbaru = mList.slice(-3).reverse();
    }

    return { success: true, data: stats };
  } catch (error) {
    return { success: false, message: 'Gagal memuat statistik: ' + error.toString() };
  }
}

function getKasRtDetails(bulanTahun) {
  try {
    var sheet = getSheet_('KAS_RT');
    var rawData = sheet.getDataRange().getDisplayValues();
    if (rawData.length <= 1) {
      return { success: true, data: [], totalPemasukan: 0, totalPengeluaran: 0, sisaSaldo: 0 };
    }

    var headers = rawData[0];
    var idIdx = headers.indexOf('ID_Kas');
    var tglIdx = headers.indexOf('Tanggal');
    var jenisIdx = headers.indexOf('Jenis_Kas');
    var katIdx = headers.indexOf('Kategori');
    var ketIdx = headers.indexOf('Keterangan');
    var nomIdx = headers.indexOf('Nominal');
    var bulanIdx = headers.indexOf('Bulan_Tahun');

    var list = [];
    var totalPemasukan = 0;
    var totalPengeluaran = 0;

    for (var i = 1; i < rawData.length; i++) {
      var row = rawData[i];
      var rowBulan = row[bulanIdx];
      var jenis = row[jenisIdx];
      var nominal = parseFloat(row[nomIdx]) || 0;

      if (!bulanTahun || bulanTahun === 'SEMUA' || rowBulan === bulanTahun) {
        if (jenis === 'Pemasukan') totalPemasukan += nominal;
        else if (jenis === 'Pengeluaran') totalPengeluaran += nominal;

        list.push({
          idKas: row[idIdx],
          tanggal: row[tglIdx],
          jenisKas: jenis,
          kategori: row[katIdx],
          keterangan: row[ketIdx],
          nominal: nominal,
          bulanTahun: rowBulan
        });
      }
    }

    var sisaSaldo = totalPemasukan - totalPengeluaran;
    return {
      success: true,
      data: list.reverse(),
      totalPemasukan: totalPemasukan,
      totalPengeluaran: totalPengeluaran,
      sisaSaldo: sisaSaldo
    };
  } catch (error) {
    return { success: false, message: 'Gagal memuat detail Kas RT: ' + error.toString() };
  }
}

function savePengeluaranRutin(payload) {
  try {
    var sheet = getSheet_('KAS_RT');
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var newId = generateSequentialId_('KAS', 'KAS_RT');

    sheet.appendRow([
      newId,
      String(payload.tanggal || '').trim(),
      'Pengeluaran',
      String(payload.kategori || 'Iuran Kas').trim(),
      String(payload.keterangan || '').trim(),
      String(payload.nominal || '0').trim(),
      String(payload.bulanTahun || '').trim(),
      nowStr
    ]);
    SpreadsheetApp.flush();
    return { success: true, message: 'Pengeluaran rutin berhasil dicatat!' };
  } catch (error) {
    return { success: false, message: 'Gagal mencatat pengeluaran: ' + error.toString() };
  }
}

function getIuranList(params) {
  try {
    params = params || {};
    var limit = parseInt(params.limit, 10) || 10;
    var offset = parseInt(params.offset, 10) || 0;
    var search = (params.search || '').toLowerCase().trim();
    var filterStatus = params.statusBayar || 'SEMUA';
    var filterJenis = params.jenisIuran || 'SEMUA';
    var userRole = params.role || 'Warga';
    var filterKk = params.noKk || '';

    var sheet = getSheet_('IURAN');
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];

    if (headers.indexOf('No_Kuitansi') !== 1) {
      repairIuranSheet_(SpreadsheetApp.getActiveSpreadsheet());
      sheet = getSheet_('IURAN');
      headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    }

    var rawData = sheet.getDataRange().getDisplayValues();
    if (rawData.length <= 1) return { success: true, data: [], total: 0, page: 1, totalPages: 0 };

    var filtered = [];
    for (var i = 1; i < rawData.length; i++) {
      var row = rawData[i];
      var id = row[headers.indexOf('ID_Iuran')] || '';
      var kwi = row[headers.indexOf('No_Kuitansi')] || '';

      // Hapus & lewati seluruh data legacy KWI-IUR-
      if (id.indexOf('KWI-IUR-') === 0 || kwi.indexOf('KWI-IUR-') === 0) {
        continue;
      }

      var kk = row[headers.indexOf('No_KK')] || '';
      var nama = (row[headers.indexOf('Nama_Warga')] || '').toLowerCase();
      var jenis = row[headers.indexOf('Jenis_Iuran')] || '';
      var status = (row[headers.indexOf('Status_Bayar')] || '').trim();

      if (status === 'Menunggu Approval' || status === 'Pending') status = 'Menunggu';
      if (status === 'Lunas') status = 'Sudah bayar';

      if (userRole === 'Warga' && filterKk && kk !== filterKk) {
        continue;
      }

      var matchSearch = !search || (nama.indexOf(search) !== -1 || kk.indexOf(search) !== -1 || id.toLowerCase().indexOf(search) !== -1 || kwi.toLowerCase().indexOf(search) !== -1);
      var matchStatus = filterStatus === 'SEMUA' || status === filterStatus || (filterStatus === 'Menunggu' && (status === 'Menunggu' || status === 'Menunggu Approval'));
      var matchJenis = filterJenis === 'SEMUA' || jenis === filterJenis;

      if (matchSearch && matchStatus && matchJenis) {
        filtered.push({
          idIuran: id,
          noKuitansi: kwi,
          noKk: kk,
          namaWarga: row[headers.indexOf('Nama_Warga')],
          bulanTahun: row[headers.indexOf('Bulan_Tahun')],
          jenisIuran: jenis,
          nominal: row[headers.indexOf('Nominal')],
          statusBayar: status,
          tanggalBayar: row[headers.indexOf('Tanggal_Bayar')],
          approvedBy: row[headers.indexOf('Approved_By')]
        });
      }
    }

    // PENGURUTAN PRIORITAS: Posisikan status "Menunggu" selalu di urutan TERATAS (Top of Page 1)
    filtered.sort(function(a, b) {
      var aIsPending = (a.statusBayar === 'Menunggu' || a.statusBayar === 'Menunggu Approval');
      var bIsPending = (b.statusBayar === 'Menunggu' || b.statusBayar === 'Menunggu Approval');
      if (aIsPending && !bIsPending) return -1;
      if (!aIsPending && bIsPending) return 1;
      return 0;
    });

    var total = filtered.length;
    var paginated = filtered.slice(offset, offset + limit);
    var totalPages = Math.ceil(total / limit) || 1;
    var currentPage = Math.floor(offset / limit) + 1;

    return { success: true, data: paginated, total: total, page: currentPage, totalPages: totalPages };
  } catch (error) {
    return { success: false, message: 'Gagal memuat data iuran: ' + error.toString() };
  }
}

function submitBayarIuran(payload) {
  try {
    var sheet = getSheet_('IURAN');
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var newId = generateSequentialId_('IUR', 'IURAN');

    var newRow = [
      newId,
      String(payload.noKuitansi || '').trim(),
      String(payload.noKk || '').trim(),
      String(payload.namaWarga || '').trim(),
      String(payload.bulanTahun || '').trim(),
      String(payload.jenisIuran || 'Iuran Kas').trim(),
      String(payload.nominal || '0').trim(),
      'Menunggu',
      String(payload.tanggalBayar || nowStr).trim(),
      '-',
      nowStr
    ];

    sheet.appendRow(newRow);
    SpreadsheetApp.flush();

    return { success: true, message: 'Pembayaran iuran berhasil dikirim! Status: Menunggu verifikasi Pengurus RT.' };
  } catch (error) {
    return { success: false, message: 'Gagal kirim pembayaran: ' + error.toString() };
  }
}

function approveIuran(idIuran, approverName, action) {
  try {
    var sheet = getSheet_('IURAN');
    var data = sheet.getDataRange().getDisplayValues();
    var headers = data[0];
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var newStatus = action === 'APPROVE' ? 'Sudah bayar' : 'Ditolak';

    var idIdx = headers.indexOf('ID_Iuran');
    var statusIdx = headers.indexOf('Status_Bayar');
    var approvedIdx = headers.indexOf('Approved_By');

    for (var i = 1; i < data.length; i++) {
      if (data[i][idIdx] === idIuran) {
        var r = i + 1;
        sheet.getRange(r, statusIdx + 1).setValue(newStatus);
        sheet.getRange(r, approvedIdx + 1).setValue(approverName || 'Pengurus RT');
        
        if (action === 'APPROVE') {
          var kasSheet = getSheet_('KAS_RT');
          var kasId = generateSequentialId_('KAS', 'KAS_RT');
          kasSheet.appendRow([
            kasId,
            nowStr.substring(0, 10),
            'Pemasukan',
            data[i][headers.indexOf('Jenis_Iuran')],
            'Iuran dari ' + data[i][headers.indexOf('Nama_Warga')] + ' (' + data[i][headers.indexOf('No_Kuitansi')] + ')',
            data[i][headers.indexOf('Nominal')],
            data[i][headers.indexOf('Bulan_Tahun')],
            nowStr
          ]);
        }
        
        SpreadsheetApp.flush();
        return { 
          success: true, 
          message: action === 'APPROVE' ? 'Pembayaran iuran DITERIMA (Sudah bayar).' : 'Pembayaran iuran DITOLAK.' 
        };
      }
    }
    return { success: false, message: 'ID Iuran ' + idIuran + ' tidak ditemukan.' };
  } catch (error) {
    return { success: false, message: 'Gagal memproses verifikasi: ' + error.toString() };
  }
}

function getLaporanKeuanganData(jenisFilter) {
  try {
    var iurSheet = getSheet_('IURAN');
    var iurData = iurSheet.getDataRange().getDisplayValues();
    if (iurData.length <= 1) return { success: true, data: [] };

    var headers = iurData[0];
    var list = [];
    var totalNominal = 0;

    for (var i = 1; i < iurData.length; i++) {
      var row = iurData[i];
      var idIur = row[headers.indexOf('ID_Iuran')] || '';
      var kwiNo = row[headers.indexOf('No_Kuitansi')] || '';

      if (idIur.indexOf('KWI-IUR-') === 0 || kwiNo.indexOf('KWI-IUR-') === 0) continue;

      var jenis = row[headers.indexOf('Jenis_Iuran')];
      var status = (row[headers.indexOf('Status_Bayar')] || '').trim();

      var matchFilter = (!jenisFilter || jenisFilter === 'SEMUA' || jenis === jenisFilter);
      if (matchFilter && (status === 'Sudah bayar' || status === 'Lunas')) {
        var nom = parseFloat(row[headers.indexOf('Nominal')]) || 0;
        totalNominal += nom;
        list.push({
          noKuitansi: kwiNo,
          namaWarga: row[headers.indexOf('Nama_Warga')],
          jenisIuran: jenis,
          bulanTahun: row[headers.indexOf('Bulan_Tahun')],
          nominal: nom,
          tanggalBayar: row[headers.indexOf('Tanggal_Bayar')],
          approvedBy: row[headers.indexOf('Approved_By')]
        });
      }
    }

    return {
      success: true,
      data: list,
      totalNominal: totalNominal,
      kopInfo: {
        rtRw: 'PENGURUS RT 010 RW 05',
        alamat: 'Kamp. Baru I Jl. Marga Mulya RT 010 RW 05 Halim Perdana Kusuma'
      }
    };
  } catch (error) {
    return { success: false, message: 'Gagal memuat laporan keuangan: ' + error.toString() };
  }
}

function getWargaList(params) {
  try {
    params = params || {};
    var limit = parseInt(params.limit, 10) || 10;
    var offset = parseInt(params.offset, 10) || 0;
    var search = (params.search || '').toLowerCase().trim();
    var filterRumah = params.statusRumah || 'SEMUA';
    var filterRole = params.role || 'SEMUA';

    var sheet = getSheet_('USERS');
    var rawData = sheet.getDataRange().getDisplayValues();
    if (rawData.length <= 1) return { success: true, data: [], total: 0, page: 1, totalPages: 0 };

    var headers = rawData[0];
    var filtered = [];
    for (var i = 1; i < rawData.length; i++) {
      var row = rawData[i];
      var nama = row[headers.indexOf('Nama')].toLowerCase();
      var noKk = row[headers.indexOf('No_KK')].toLowerCase();
      var noKtp = row[headers.indexOf('No_KTP')].toLowerCase();
      var noHp = row[headers.indexOf('No_HP')].toLowerCase();
      var rumah = row[headers.indexOf('Status_Rumah')];
      var role = row[headers.indexOf('Role')];

      var matchSearch = !search || (nama.indexOf(search) !== -1 || noKk.indexOf(search) !== -1 || noKtp.indexOf(search) !== -1 || noHp.indexOf(search) !== -1);
      var matchRumah = filterRumah === 'SEMUA' || rumah === filterRumah;
      var matchRole = filterRole === 'SEMUA' || role === filterRole;

      if (matchSearch && matchRumah && matchRole) {
        filtered.push({
          noKk: row[headers.indexOf('No_KK')],
          noKtp: row[headers.indexOf('No_KTP')],
          nama: row[headers.indexOf('Nama')],
          noHp: row[headers.indexOf('No_HP')],
          role: row[headers.indexOf('Role')],
          statusUser: row[headers.indexOf('Status_User')],
          alamat: row[headers.indexOf('Alamat')],
          umur: row[headers.indexOf('Umur')],
          statusRumah: row[headers.indexOf('Status_Rumah')],
          pekerjaan: row[headers.indexOf('Pekerjaan')]
        });
      }
    }

    var total = filtered.length;
    var paginated = filtered.slice(offset, offset + limit);
    var totalPages = Math.ceil(total / limit) || 1;
    var currentPage = Math.floor(offset / limit) + 1;

    return { success: true, data: paginated, total: total, page: currentPage, totalPages: totalPages };
  } catch (error) {
    return { success: false, message: 'Gagal mengambil data warga: ' + error.toString() };
  }
}

function getFamilyDetails(noKk) {
  try {
    var userSheet = getSheet_('USERS');
    var agtSheet = getSheet_('ANGGOTA_KELUARGA');

    var userData = userSheet.getDataRange().getDisplayValues();
    var agtData = agtSheet.getDataRange().getDisplayValues();

    var kepalaKeluarga = null;
    var uHeaders = userData[0];
    for (var i = 1; i < userData.length; i++) {
      if (userData[i][uHeaders.indexOf('No_KK')] === String(noKk).trim()) {
        kepalaKeluarga = {
          noKk: userData[i][uHeaders.indexOf('No_KK')],
          noKtp: userData[i][uHeaders.indexOf('No_KTP')],
          nama: userData[i][uHeaders.indexOf('Nama')],
          noHp: userData[i][uHeaders.indexOf('No_HP')],
          role: userData[i][uHeaders.indexOf('Role')],
          statusUser: userData[i][uHeaders.indexOf('Status_User')],
          tanggalLahir: userData[i][uHeaders.indexOf('Tanggal_Lahir')],
          umur: userData[i][uHeaders.indexOf('Umur')],
          alamat: userData[i][uHeaders.indexOf('Alamat')],
          jenisKelamin: userData[i][uHeaders.indexOf('Jenis_Kelamin')],
          statusKeluarga: userData[i][uHeaders.indexOf('Status_Keluarga')],
          statusRumah: userData[i][uHeaders.indexOf('Status_Rumah')],
          pendidikan: userData[i][uHeaders.indexOf('Pendidikan')],
          pekerjaan: userData[i][uHeaders.indexOf('Pekerjaan')]
        };
        break;
      }
    }

    var anggotaList = [];
    if (agtData.length > 1) {
      var aHeaders = agtData[0];
      for (var j = 1; j < agtData.length; j++) {
        if (agtData[j][aHeaders.indexOf('No_KK')] === String(noKk).trim()) {
          anggotaList.push({
            idAnggota: agtData[j][aHeaders.indexOf('ID_Anggota')],
            noKk: agtData[j][aHeaders.indexOf('No_KK')],
            namaAnggota: agtData[j][aHeaders.indexOf('Nama_Anggota')],
            hubunganKeluarga: agtData[j][aHeaders.indexOf('Hubungan_Keluarga')],
            tanggalLahir: agtData[j][aHeaders.indexOf('Tanggal_Lahir')],
            umur: agtData[j][aHeaders.indexOf('Umur')],
            jenisKelamin: agtData[j][aHeaders.indexOf('Jenis_Kelamin')]
          });
        }
      }
    }

    return { success: true, data: { kepalaKeluarga: kepalaKeluarga, anggota: anggotaList } };
  } catch (error) {
    return { success: false, message: 'Gagal mengambil detail keluarga: ' + error.toString() };
  }
}

function saveAnggotaKeluarga(payload) {
  try {
    var sheet = getSheet_('ANGGOTA_KELUARGA');
    var data = sheet.getDataRange().getDisplayValues();
    var headers = data[0];
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var isEdit = payload.idAnggota && payload.idAnggota.trim() !== '';

    if (isEdit) {
      for (var i = 1; i < data.length; i++) {
        if (data[i][headers.indexOf('ID_Anggota')] === payload.idAnggota) {
          var r = i + 1;
          sheet.getRange(r, headers.indexOf('Nama_Anggota') + 1).setValue(payload.namaAnggota);
          sheet.getRange(r, headers.indexOf('Hubungan_Keluarga') + 1).setValue(payload.hubunganKeluarga);
          sheet.getRange(r, headers.indexOf('Tanggal_Lahir') + 1).setValue(payload.tanggalLahir);
          sheet.getRange(r, headers.indexOf('Umur') + 1).setValue(payload.umur);
          sheet.getRange(r, headers.indexOf('Jenis_Kelamin') + 1).setValue(payload.jenisKelamin);
          SpreadsheetApp.flush();
          return { success: true, message: 'Data anggota keluarga berhasil diperbarui!' };
        }
      }
    }

    var newId = generateSequentialId_('AGT', 'ANGGOTA_KELUARGA');
    sheet.appendRow([newId, String(payload.noKk).trim(), String(payload.namaAnggota).trim(), String(payload.hubunganKeluarga).trim(), String(payload.tanggalLahir).trim(), String(payload.umur || '0').trim(), String(payload.jenisKelamin || 'Laki-laki').trim(), nowStr]);
    SpreadsheetApp.flush();
    return { success: true, message: 'Anggota keluarga baru berhasil ditambahkan!', idAnggota: newId };
  } catch (error) {
    return { success: false, message: 'Gagal menyimpan anggota keluarga: ' + error.toString() };
  }
}

function deleteAnggotaKeluarga(idAnggota) {
  try {
    var sheet = getSheet_('ANGGOTA_KELUARGA');
    var data = sheet.getDataRange().getDisplayValues();
    var idIdx = data[0].indexOf('ID_Anggota');
    for (var i = 1; i < data.length; i++) {
      if (data[i][idIdx] === idAnggota) {
        sheet.deleteRow(i + 1);
        SpreadsheetApp.flush();
        return { success: true, message: 'Anggota keluarga berhasil dihapus!' };
      }
    }
    return { success: false, message: 'ID Anggota tidak ditemukan.' };
  } catch (error) {
    return { success: false, message: 'Gagal menghapus: ' + error.toString() };
  }
}

function getMadingList(params) {
  try {
    params = params || {};
    var search = (params.search || '').toLowerCase().trim();
    var userRole = params.role || 'Warga';

    var sheet = getSheet_('MADING');
    var rawData = sheet.getDataRange().getDisplayValues();
    if (rawData.length <= 1) return { success: true, data: [] };

    var headers = rawData[0];
    var list = [];
    for (var i = 1; i < rawData.length; i++) {
      var row = rawData[i];
      var status = row[headers.indexOf('Status_Publish')];
      var judul = row[headers.indexOf('Judul')].toLowerCase();
      var note = row[headers.indexOf('Note')].toLowerCase();

      if (userRole === 'Warga' && status !== 'Published') {
        continue;
      }

      var matchSearch = !search || (judul.indexOf(search) !== -1 || note.indexOf(search) !== -1);
      if (matchSearch) {
        list.push({
          idMading: row[headers.indexOf('ID_Mading')],
          judul: row[headers.indexOf('Judul')],
          tanggalKegiatan: row[headers.indexOf('Tanggal_Kegiatan')],
          note: row[headers.indexOf('Note')],
          statusPublish: row[headers.indexOf('Status_Publish')],
          pembuat: row[headers.indexOf('Pembuat')],
          createdAt: row[headers.indexOf('Created_At')]
        });
      }
    }
    return { success: true, data: list.reverse() };
  } catch (error) {
    return { success: false, message: 'Gagal memuat mading: ' + error.toString() };
  }
}

function saveMading(payload) {
  try {
    var sheet = getSheet_('MADING');
    var data = sheet.getDataRange().getDisplayValues();
    var headers = data[0];
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var isEdit = payload.idMading && payload.idMading.trim() !== '';

    if (isEdit) {
      for (var i = 1; i < data.length; i++) {
        if (data[i][headers.indexOf('ID_Mading')] === payload.idMading) {
          var r = i + 1;
          sheet.getRange(r, headers.indexOf('Judul') + 1).setValue(payload.judul);
          sheet.getRange(r, headers.indexOf('Tanggal_Kegiatan') + 1).setValue(payload.tanggalKegiatan);
          sheet.getRange(r, headers.indexOf('Note') + 1).setValue(payload.note);
          sheet.getRange(r, headers.indexOf('Status_Publish') + 1).setValue(payload.statusPublish);
          SpreadsheetApp.flush();
          return { success: true, message: 'Informasi Mading berhasil diperbarui!' };
        }
      }
    }

    var newId = generateSequentialId_('MAD', 'MADING');
    sheet.appendRow([
      newId,
      String(payload.judul || '').trim(),
      String(payload.tanggalKegiatan || '').trim(),
      String(payload.note || '').trim(),
      String(payload.statusPublish || 'Published').trim(),
      String(payload.pembuat || 'Pengurus RT').trim(),
      nowStr
    ]);
    SpreadsheetApp.flush();
    return { success: true, message: 'Informasi Mading berhasil dipublikasikan!', idMading: newId };
  } catch (error) {
    return { success: false, message: 'Gagal menyimpan mading: ' + error.toString() };
  }
}

function deleteMading(idMading) {
  try {
    var sheet = getSheet_('MADING');
    var data = sheet.getDataRange().getDisplayValues();
    var idIdx = data[0].indexOf('ID_Mading');
    for (var i = 1; i < data.length; i++) {
      if (data[i][idIdx] === idMading) {
        sheet.deleteRow(i + 1);
        SpreadsheetApp.flush();
        return { success: true, message: 'Informasi Mading berhasil dihapus!' };
      }
    }
    return { success: false, message: 'ID Mading tidak ditemukan.' };
  } catch (error) {
    return { success: false, message: 'Gagal menghapus mading: ' + error.toString() };
  }
}
