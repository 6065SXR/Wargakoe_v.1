/**
 * ====================================================================
 * WARGAKOE - CITIZEN & FAMILY DATA BACKEND MODULE
 * By Zettbos System (ZettBOT 3.1)
 * ====================================================================
 */

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
    var rawData = sheet.getDataRange().getDisplayValues();
    if (rawData.length <= 1) return { success: true, data: [] };
    var headers = rawData[0];
    var list = [];
    for (var i = 1; i < rawData.length; i++) {
      var row = rawData[i];
      var obj = {};
      for (var h = 0; h < headers.length; h++) {
        if (headers[h] !== 'Password') {
          obj[headers[h]] = row[h];
        }
      }
      list.push(obj);
    }
    return { success: true, data: list };
  } catch (err) {
    return { success: false, message: 'Gagal export data: ' + err.toString() };
  }
}

function importWargaBatch(wargaArray) {
  try {
    var sheet = getSheet_('USERS');
    var rawData = sheet.getDataRange().getDisplayValues();
    var headers = rawData[0];
    var existingKk = [];
    var existingKtp = [];
    for (var i = 1; i < rawData.length; i++) {
      existingKk.push(rawData[i][headers.indexOf('No_KK')]);
      existingKtp.push(rawData[i][headers.indexOf('No_KTP')]);
    }

    var importedCount = 0;
    var duplicates = [];
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');

    for (var j = 0; j < wargaArray.length; j++) {
      var w = wargaArray[j];
      var kk = String(w.No_KK || w.noKk || '').trim();
      var ktp = String(w.No_KTP || w.noKtp || '').trim();
      var nama = String(w.Nama || w.nama || '').trim();

      if (!kk || !ktp) continue;

      if (existingKk.indexOf(kk) !== -1 || existingKtp.indexOf(ktp) !== -1) {
        duplicates.push(nama + ' (KK: ' + kk + ')');
        continue;
      }

      sheet.appendRow([
        kk,
        ktp,
        nama,
        String(w.No_HP || w.noHp || '').trim(),
        String(w.Password || w.password || 'warga123').trim(),
        String(w.Role || w.role || 'Warga').trim(),
        String(w.Status_User || w.statusUser || 'Warga').trim(),
        String(w.Tanggal_Lahir || w.tanggalLahir || '').trim(),
        String(w.Umur || w.umur || '0').trim(),
        String(w.Alamat || w.alamat || 'Kamp. Baru I Jl. Marga Mulya RT 010 RW 05').trim(),
        String(w.Jenis_Kelamin || w.jenisKelamin || 'Laki-laki').trim(),
        String(w.Status_Keluarga || w.statusKeluarga || 'Suami').trim(),
        String(w.Status_Rumah || w.statusRumah || 'Pribadi').trim(),
        String(w.Pendidikan || w.pendidikan || 'SMA').trim(),
        String(w.Pekerjaan || w.pekerjaan || 'Wiraswasta').trim(),
        nowStr
      ]);
      existingKk.push(kk);
      existingKtp.push(ktp);
      importedCount++;
    }

    SpreadsheetApp.flush();
    return {
      success: true,
      importedCount: importedCount,
      duplicates: duplicates,
      message: 'Berhasil mengimpor ' + importedCount + ' data warga baru.'
    };
  } catch (err) {
    return { success: false, message: 'Gagal impor data: ' + err.toString() };
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
