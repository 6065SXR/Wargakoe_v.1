/**
 * ====================================================================
 * WARGAKOE - BACKEND LOGIC & CRUD OPERATIONS (REVISED)
 * By Zettbos System (ZettBOT 3.1)
 * ====================================================================
 */

function doGet(e) {
  try {
    var template = HtmlService.createTemplateFromFile('index');
    return template.evaluate()
      .setTitle('Wargakoe - Sistem Pengelolaan Warga & Iuran RT')
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
    var kkIdx = headers.indexOf('No_KK');
    var namaIdx = headers.indexOf('Nama');
    var roleIdx = headers.indexOf('Role');
    var statusUserIdx = headers.indexOf('Status_User');

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[hpIdx] === noHp.trim() && row[passIdx] === password.trim()) {
        var userObj = {
          noKk: row[kkIdx],
          noKtp: row[headers.indexOf('No_KTP')],
          nama: row[namaIdx],
          noHp: row[hpIdx],
          role: row[roleIdx],
          statusUser: row[statusUserIdx],
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
      kasRtBulanIni: 0,        // Iuran RT
      kasKematianBulanIni: 0,  // Iuran Duka
      iuranSampahBulanIni: 0,
      iuranSosialBulanIni: 0,
      iuranLainBulanIni: 0,
      totalTerkumpul: 0,
      totalMenunggu: 0,
      totalBelumBayar: 0,
      tagihanSaya: [],
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
        var nominal = parseFloat(row[iNominalIdx]) || 0;
        var status = row[iStatusIdx];
        var jenis = row[iJenisIdx];
        var kk = row[iKkIdx];

        if (status === 'Lunas') {
          stats.totalTerkumpul += nominal;
          if (jenis === 'Iuran RT') stats.kasRtBulanIni += nominal;
          else if (jenis === 'Iuran Duka') stats.kasKematianBulanIni += nominal;
          else if (jenis === 'Iuran Sampah') stats.iuranSampahBulanIni += nominal;
          else if (jenis === 'Iuran Sosial') stats.iuranSosialBulanIni += nominal;
          else stats.iuranLainBulanIni += nominal;
        } else if (status === 'Menunggu Approval') {
          stats.totalMenunggu += nominal;
        } else {
          stats.totalBelumBayar += nominal;
        }

        if (userRole === 'Warga' && kk === userNoKk && status !== 'Lunas') {
          stats.tagihanSaya.push({
            idIuran: row[iHeaders.indexOf('ID_Iuran')],
            jenisIuran: jenis,
            bulanTahun: row[iHeaders.indexOf('Bulan_Tahun')],
            nominal: nominal,
            statusBayar: status
          });
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

function getAllWargaForExport() {
  try {
    var sheet = getSheet_('USERS');
    var data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) return { success: true, data: [] };

    var headers = data[0];
    var list = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      list.push({
        noKk: row[headers.indexOf('No_KK')],
        noKtp: row[headers.indexOf('No_KTP')],
        nama: row[headers.indexOf('Nama')],
        noHp: row[headers.indexOf('No_HP')],
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
      });
    }
    return { success: true, data: list };
  } catch (error) {
    return { success: false, message: 'Gagal mengambil data export: ' + error.toString() };
  }
}

function importWargaBatch(wargaList) {
  try {
    if (!wargaList || !Array.isArray(wargaList) || wargaList.length === 0) {
      return { success: false, message: 'Data yang diimpor kosong atau format tidak valid.' };
    }

    var sheet = getSheet_('USERS');
    var data = sheet.getDataRange().getDisplayValues();
    var headers = data[0];

    var kkIdx = headers.indexOf('No_KK');
    var ktpIdx = headers.indexOf('No_KTP');

    var existingKk = {};
    var existingKtp = {};

    for (var i = 1; i < data.length; i++) {
      if (data[i][kkIdx]) existingKk[data[i][kkIdx].trim()] = true;
      if (data[i][ktpIdx]) existingKtp[data[i][ktpIdx].trim()] = true;
    }

    var newRows = [];
    var duplicates = [];
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');

    for (var j = 0; j < wargaList.length; j++) {
      var item = wargaList[j];
      var noKk = String(item.noKk || item.No_KK || '').trim();
      var noKtp = String(item.noKtp || item.No_KTP || '').trim();
      var nama = String(item.nama || item.Nama || 'Warga').trim();

      if (!noKk || !noKtp) continue;

      var isDup = false;
      if (existingKk[noKk]) {
        duplicates.push('No. KK ' + noKk + ' (' + nama + ')');
        isDup = true;
      }
      if (existingKtp[noKtp]) {
        duplicates.push('NIK/KTP ' + noKtp + ' (' + nama + ')');
        isDup = true;
      }

      if (!isDup) {
        existingKk[noKk] = true;
        existingKtp[noKtp] = true;

        newRows.push([
          noKk,
          noKtp,
          nama,
          String(item.noHp || item.No_HP || '').trim(),
          String(item.password || item.Password || 'warga123').trim(),
          String(item.role || item.Role || 'Warga').trim(),
          String(item.statusUser || item.Status_User || 'Warga').trim(),
          String(item.tanggalLahir || item.Tanggal_Lahir || '').trim(),
          String(item.umur || item.Umur || '0').trim(),
          String(item.alamat || item.Alamat || '').trim(),
          String(item.jenisKelamin || item.Jenis_Kelamin || 'Laki-laki').trim(),
          String(item.statusKeluarga || item.Status_Keluarga || 'Suami').trim(),
          String(item.statusRumah || item.Status_Rumah || 'Pribadi').trim(),
          String(item.pendidikan || item.Pendidikan || 'SMA').trim(),
          String(item.pekerjaan || item.Pekerjaan || 'Wiraswasta').trim(),
          nowStr
        ]);
      }
    }

    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
      SpreadsheetApp.flush();
    }

    return {
      success: true,
      importedCount: newRows.length,
      duplicates: duplicates,
      message: 'Proses import selesai. ' + newRows.length + ' data baru berhasil diimpor.'
    };
  } catch (error) {
    return { success: false, message: 'Gagal import data warga: ' + error.toString() };
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

function getIuranList(params) {
  try {
    params = params || {};
    var limit = parseInt(params.limit, 10) || 10;
    var offset = parseInt(params.offset, 10) || 0;
    var search = (params.search || '').toLowerCase().trim();
    var filterStatus = params.statusBayar || 'SEMUA';
    var filterJenis = params.jenisIuran || 'SEMUA';
    var filterKk = params.noKk || '';

    var sheet = getSheet_('IURAN');
    var rawData = sheet.getDataRange().getDisplayValues();
    if (rawData.length <= 1) return { success: true, data: [], total: 0, page: 1, totalPages: 0 };

    var headers = rawData[0];
    var filtered = [];
    for (var i = 1; i < rawData.length; i++) {
      var row = rawData[i];
      var id = row[headers.indexOf('ID_Iuran')];
      var kk = row[headers.indexOf('No_KK')];
      var nama = row[headers.indexOf('Nama_Warga')].toLowerCase();
      var jenis = row[headers.indexOf('Jenis_Iuran')];
      var status = row[headers.indexOf('Status_Bayar')];

      var matchKk = !filterKk || kk === filterKk;
      var matchSearch = !search || (nama.indexOf(search) !== -1 || kk.indexOf(search) !== -1 || id.toLowerCase().indexOf(search) !== -1);
      var matchStatus = filterStatus === 'SEMUA' || status === filterStatus;
      var matchJenis = filterJenis === 'SEMUA' || jenis === filterJenis;

      if (matchKk && matchSearch && matchStatus && matchJenis) {
        filtered.push({
          idIuran: row[headers.indexOf('ID_Iuran')],
          noKk: row[headers.indexOf('No_KK')],
          namaWarga: row[headers.indexOf('Nama_Warga')],
          bulanTahun: row[headers.indexOf('Bulan_Tahun')],
          jenisIuran: row[headers.indexOf('Jenis_Iuran')],
          nominal: row[headers.indexOf('Nominal')],
          statusBayar: row[headers.indexOf('Status_Bayar')],
          tanggalBayar: row[headers.indexOf('Tanggal_Bayar')],
          approvedBy: row[headers.indexOf('Approved_By')]
        });
      }
    }

    var total = filtered.length;
    var paginated = filtered.slice(offset, offset + limit);
    var totalPages = Math.ceil(total / limit) || 1;
    var currentPage = Math.floor(offset / limit) + 1;

    return { success: true, data: paginated, total: total, page: currentPage, totalPages: totalPages };
  } catch (error) {
    return { success: false, message: 'Gagal memuat data iuran: ' + error.toString() };
  }
}

function konfirmasiBayarIuran(idIuran, noKk) {
  try {
    var sheet = getSheet_('IURAN');
    var data = sheet.getDataRange().getDisplayValues();
    var headers = data[0];
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');

    for (var i = 1; i < data.length; i++) {
      if (data[i][headers.indexOf('ID_Iuran')] === idIuran && (noKk === '' || data[i][headers.indexOf('No_KK')] === noKk)) {
        var r = i + 1;
        sheet.getRange(r, headers.indexOf('Status_Bayar') + 1).setValue('Menunggu Approval');
        sheet.getRange(r, headers.indexOf('Tanggal_Bayar') + 1).setValue(nowStr);
        SpreadsheetApp.flush();
        return { success: true, message: 'Konfirmasi pembayaran berhasil dikirim!' };
      }
    }
    return { success: false, message: 'Tagihan iuran tidak ditemukan.' };
  } catch (error) {
    return { success: false, message: 'Gagal konfirmasi bayar: ' + error.toString() };
  }
}

function approveIuran(idIuran, approverName, action) {
  try {
    var sheet = getSheet_('IURAN');
    var data = sheet.getDataRange().getDisplayValues();
    var headers = data[0];
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var newStatus = action === 'APPROVE' ? 'Lunas' : 'Belum Bayar';

    for (var i = 1; i < data.length; i++) {
      if (data[i][headers.indexOf('ID_Iuran')] === idIuran) {
        var r = i + 1;
        sheet.getRange(r, headers.indexOf('Status_Bayar') + 1).setValue(newStatus);
        sheet.getRange(r, headers.indexOf('Approved_By') + 1).setValue(approverName || 'Pengurus RT');
        if (action === 'APPROVE' && !data[i][headers.indexOf('Tanggal_Bayar')]) {
          sheet.getRange(r, headers.indexOf('Tanggal_Bayar') + 1).setValue(nowStr);
        }
        SpreadsheetApp.flush();
        return { success: true, message: action === 'APPROVE' ? 'Iuran disetujui (LUNAS).' : 'Iuran ditolak.' };
      }
    }
    return { success: false, message: 'ID Iuran tidak ditemukan.' };
  } catch (error) {
    return { success: false, message: 'Gagal memproses approval: ' + error.toString() };
  }
}

function generateTagihanBulananMassal(bulanTahun, jenisIuran, nominal) {
  try {
    var userSheet = getSheet_('USERS');
    var iurSheet = getSheet_('IURAN');
    var userData = userSheet.getDataRange().getDisplayValues();
    if (userData.length <= 1) return { success: false, message: 'Tidak ada data warga.' };

    var iurData = iurSheet.getDataRange().getDisplayValues();
    var iHeaders = iurData[0];
    var uHeaders = userData[0];
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var newRows = [];
    var count = 0;

    for (var u = 1; u < userData.length; u++) {
      var noKk = userData[u][uHeaders.indexOf('No_KK')];
      var nama = userData[u][uHeaders.indexOf('Nama')];

      var exists = false;
      for (var k = 1; k < iurData.length; k++) {
        if (iurData[k][iHeaders.indexOf('No_KK')] === noKk && iurData[k][iHeaders.indexOf('Bulan_Tahun')] === bulanTahun && iurData[k][iHeaders.indexOf('Jenis_Iuran')] === jenisIuran) {
          exists = true;
          break;
        }
      }

      if (!exists) {
        var idIuran = generateSequentialId_('IUR', 'IURAN');
        newRows.push([idIuran, noKk, nama, bulanTahun, jenisIuran, String(nominal), 'Belum Bayar', '-', '-', nowStr]);
        count++;
      }
    }

    if (newRows.length > 0) {
      iurSheet.getRange(iurSheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
      SpreadsheetApp.flush();
      return { success: true, message: 'Berhasil menerbitkan ' + count + ' tagihan iuran baru.' };
    } else {
      return { success: false, message: 'Seluruh warga sudah memiliki tagihan tersebut.' };
    }
  } catch (error) {
    return { success: false, message: 'Gagal generate tagihan: ' + error.toString() };
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
