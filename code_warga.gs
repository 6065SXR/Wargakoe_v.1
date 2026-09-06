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

function ensureAnggotaKtpHeader_(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];
  var headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  if (headers.indexOf('No_KTP') === -1) {
    var newCol = lastCol + 1;
    var cell = sheet.getRange(1, newCol);
    cell.setValue('No_KTP');
    cell.setBackground('#221d52');
    cell.setFontColor('#ffffff');
    cell.setFontWeight('bold');
    cell.setHorizontalAlignment('center');
    headers.push('No_KTP');
  }
  return headers;
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
      var aIdIdx = aHeaders.indexOf('ID_Anggota');
      var aKkIdx = aHeaders.indexOf('No_KK');
      var aKtpIdx = aHeaders.indexOf('No_KTP');
      var aNamaIdx = aHeaders.indexOf('Nama_Anggota');
      var aHubIdx = aHeaders.indexOf('Hubungan_Keluarga');
      var aTglIdx = aHeaders.indexOf('Tanggal_Lahir');
      var aUmurIdx = aHeaders.indexOf('Umur');
      var aGenIdx = aHeaders.indexOf('Jenis_Kelamin');

      for (var j = 1; j < agtData.length; j++) {
        if (agtData[j][aKkIdx] === String(noKk).trim()) {
          anggotaList.push({
            idAnggota: agtData[j][aIdIdx],
            noKk: agtData[j][aKkIdx],
            noKtp: aKtpIdx !== -1 ? agtData[j][aKtpIdx] : '',
            namaAnggota: agtData[j][aNamaIdx],
            hubunganKeluarga: agtData[j][aHubIdx],
            tanggalLahir: agtData[j][aTglIdx],
            umur: agtData[j][aUmurIdx],
            jenisKelamin: agtData[j][aGenIdx]
          });
        }
      }
    }

    return { success: true, data: { kepalaKeluarga: kepalaKeluarga, anggota: anggotaList } };
  } catch (error) {
    return { success: false, message: 'Gagal mengambil detail keluarga: ' + error.toString() };
  }
}

function updateProfilWarga(payload) {
  try {
    var sheet = getSheet_('USERS');
    var rawData = sheet.getDataRange().getDisplayValues();
    var headers = rawData[0];

    var origNoKk = String(payload.origNoKk || '').trim();
    var inputKtp = String(payload.noKtp || '').trim();
    var isSuperAdmin = Boolean(payload.isSuperAdmin);
    var targetNoKk = isSuperAdmin ? String(payload.noKk || origNoKk).trim() : origNoKk;

    if (!origNoKk) {
      return { success: false, message: 'Nomor KK asal tidak valid!' };
    }

    if (!inputKtp || inputKtp.length !== 16 || isNaN(inputKtp)) {
      return { success: false, message: 'Nomor NIK/KTP wajib 16 digit angka!' };
    }

    // 1. Temukan baris user di tabel USERS
    var userRowIndex = -1;
    var uKkIdx = headers.indexOf('No_KK');
    var uKtpIdx = headers.indexOf('No_KTP');
    var uNamaIdx = headers.indexOf('Nama');
    var uHpIdx = headers.indexOf('No_HP');
    var uPassIdx = headers.indexOf('Password');
    var uRoleIdx = headers.indexOf('Role');
    var uStatusUserIdx = headers.indexOf('Status_User');
    var uTglIdx = headers.indexOf('Tanggal_Lahir');
    var uUmurIdx = headers.indexOf('Umur');
    var uAlamatIdx = headers.indexOf('Alamat');
    var uGenIdx = headers.indexOf('Jenis_Kelamin');
    var uHubIdx = headers.indexOf('Status_Keluarga');
    var uRumahIdx = headers.indexOf('Status_Rumah');
    var uPendIdx = headers.indexOf('Pendidikan');
    var uPekIdx = headers.indexOf('Pekerjaan');

    for (var i = 1; i < rawData.length; i++) {
      if (rawData[i][uKkIdx] === origNoKk) {
        userRowIndex = i + 1; // 1-based index di Google Sheets
        break;
      }
    }

    if (userRowIndex === -1) {
      return { success: false, message: 'Data pengguna tidak ditemukan di database!' };
    }

    // 2. Validasi Anti-Duplikasi NIK ke akun lain di tabel USERS
    for (var u = 1; u < rawData.length; u++) {
      if (u + 1 !== userRowIndex && rawData[u][uKtpIdx] === inputKtp) {
        return { success: false, message: 'Nomor NIK/KTP (' + inputKtp + ') sudah digunakan oleh warga lain!' };
      }
    }

    // 3. Validasi Anti-Duplikasi NIK ke tabel ANGGOTA_KELUARGA
    var agtSheet = getSheet_('ANGGOTA_KELUARGA');
    var agtData = agtSheet.getDataRange().getDisplayValues();
    if (agtData.length > 1) {
      var aKtpIdx = agtData[0].indexOf('No_KTP');
      if (aKtpIdx !== -1) {
        for (var a = 1; a < agtData.length; a++) {
          if (agtData[a][aKtpIdx] === inputKtp) {
            return { success: false, message: 'Nomor NIK/KTP (' + inputKtp + ') sudah terdaftar pada anggota keluarga lain!' };
          }
        }
      }
    }

    // 4. Update data baris pengguna di tabel USERS
    if (isSuperAdmin && targetNoKk !== origNoKk) {
      sheet.getRange(userRowIndex, uKkIdx + 1).setValue(targetNoKk);
    }
    sheet.getRange(userRowIndex, uKtpIdx + 1).setValue(inputKtp);
    sheet.getRange(userRowIndex, uNamaIdx + 1).setValue(payload.nama);
    sheet.getRange(userRowIndex, uHpIdx + 1).setValue(payload.noHp);
    sheet.getRange(userRowIndex, uTglIdx + 1).setValue(payload.tanggalLahir);
    sheet.getRange(userRowIndex, uUmurIdx + 1).setValue(payload.umur);
    sheet.getRange(userRowIndex, uAlamatIdx + 1).setValue(payload.alamat);
    sheet.getRange(userRowIndex, uGenIdx + 1).setValue(payload.jenisKelamin);
    sheet.getRange(userRowIndex, uHubIdx + 1).setValue(payload.statusKeluarga);
    sheet.getRange(userRowIndex, uRumahIdx + 1).setValue(payload.statusRumah);
    sheet.getRange(userRowIndex, uPendIdx + 1).setValue(payload.pendidikan);
    sheet.getRange(userRowIndex, uPekIdx + 1).setValue(payload.pekerjaan);

    // Update password hanya jika kolom password baru diisi
    if (payload.password && String(payload.password).trim() !== '') {
      sheet.getRange(userRowIndex, uPassIdx + 1).setValue(String(payload.password).trim());
    }

    SpreadsheetApp.flush();

    // Susun objek user terbaru untuk sinkronisasi sesi frontend
    var updatedRow = sheet.getRange(userRowIndex, 1, 1, headers.length).getDisplayValues()[0];
    var updatedUser = {
      noKk: updatedRow[uKkIdx],
      noKtp: updatedRow[uKtpIdx],
      nama: updatedRow[uNamaIdx],
      noHp: updatedRow[uHpIdx],
      role: updatedRow[uRoleIdx],
      statusUser: updatedRow[uStatusUserIdx],
      tanggalLahir: updatedRow[uTglIdx],
      umur: updatedRow[uUmurIdx],
      alamat: updatedRow[uAlamatIdx],
      jenisKelamin: updatedRow[uGenIdx],
      statusKeluarga: updatedRow[uHubIdx],
      statusRumah: updatedRow[uRumahIdx],
      pendidikan: updatedRow[uPendIdx],
      pekerjaan: updatedRow[uPekIdx]
    };

    return { 
      success: true, 
      message: 'Data profil kepala keluarga berhasil diperbarui!',
      user: updatedUser 
    };
  } catch (error) {
    return { success: false, message: 'Gagal memperbarui profil: ' + error.toString() };
  }
}

function saveAnggotaKeluarga(payload) {
  try {
    var sheet = getSheet_('ANGGOTA_KELUARGA');
    var rawData = sheet.getDataRange().getDisplayValues();
    var headers = ensureAnggotaKtpHeader_(sheet);
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var isEdit = payload.idAnggota && payload.idAnggota.trim() !== '';
    var inputKtp = String(payload.noKtp || '').trim();

    if (!inputKtp || inputKtp.length !== 16 || isNaN(inputKtp)) {
      return { success: false, message: 'Nomor NIK/KTP wajib 16 digit angka!' };
    }

    // 1. VALIDASI ANTI-DUPLIKASI: Periksa tabel USERS (Kepala Keluarga / Pengurus RT)
    var userSheet = getSheet_('USERS');
    var userData = userSheet.getDataRange().getDisplayValues();
    if (userData.length > 1) {
      var uHeaders = userData[0];
      var uKtpIdx = uHeaders.indexOf('No_KTP');
      for (var u = 1; u < userData.length; u++) {
        if (userData[u][uKtpIdx] === inputKtp) {
          return { success: false, message: 'Nomor NIK/KTP (' + inputKtp + ') sudah terdaftar pada akun warga/kepala keluarga lain!' };
        }
      }
    }

    // 2. VALIDASI ANTI-DUPLIKASI: Periksa tabel ANGGOTA_KELUARGA
    var aKtpIdx = headers.indexOf('No_KTP');
    var aIdIdx = headers.indexOf('ID_Anggota');
    if (rawData.length > 1) {
      for (var a = 1; a < rawData.length; a++) {
        if (rawData[a][aKtpIdx] === inputKtp) {
          if (isEdit && rawData[a][aIdIdx] === payload.idAnggota) {
            continue; // data yang sama saat mode edit
          }
          return { success: false, message: 'Nomor NIK/KTP (' + inputKtp + ') sudah terdaftar pada anggota keluarga lain!' };
        }
      }
    }

    if (isEdit) {
      for (var i = 1; i < rawData.length; i++) {
        if (rawData[i][headers.indexOf('ID_Anggota')] === payload.idAnggota) {
          var r = i + 1;
          sheet.getRange(r, headers.indexOf('No_KTP') + 1).setValue(inputKtp);
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
    var newRow = new Array(headers.length);
    for (var h = 0; h < headers.length; h++) {
      var col = headers[h];
      if (col === 'ID_Anggota') newRow[h] = newId;
      else if (col === 'No_KK') newRow[h] = String(payload.noKk).trim();
      else if (col === 'No_KTP') newRow[h] = inputKtp;
      else if (col === 'Nama_Anggota') newRow[h] = String(payload.namaAnggota).trim();
      else if (col === 'Hubungan_Keluarga') newRow[h] = String(payload.hubunganKeluarga).trim();
      else if (col === 'Tanggal_Lahir') newRow[h] = String(payload.tanggalLahir).trim();
      else if (col === 'Umur') newRow[h] = String(payload.umur || '0').trim();
      else if (col === 'Jenis_Kelamin') newRow[h] = String(payload.jenisKelamin || 'Laki-laki').trim();
      else if (col === 'Created_At') newRow[h] = nowStr;
      else newRow[h] = '';
    }

    sheet.appendRow(newRow);
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
