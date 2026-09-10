/**
 * ====================================================================
 * WARGAKOE - DATA WARGA BACKEND MODULE (KEPENDUDUKAN & IMPOR/EKSPOR)
 * Mengelola direktori kependudukan, pengurutan alfabetis A-Z, filter abjad,
 * penggabungan anggota keluarga per No. KK, unduh CSV, dan impor batch.
 * ====================================================================
 */

function getWargaList(params) {
  try {
    params = params || {};
    var search = String(params.search || '').toLowerCase().trim();
    var statusRumah = String(params.statusRumah || 'SEMUA').trim();
    var alphabet = String(params.alphabet || 'SEMUA').toUpperCase().trim();
    var limit = parseInt(params.limit, 10) || 10;
    var offset = parseInt(params.offset, 10) || 0;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var userSheet = ss.getSheetByName('USERS');
    if (!userSheet || userSheet.getLastRow() <= 1) {
      return { success: true, data: [], total: 0, page: 1, totalPages: 1 };
    }

    var uData = userSheet.getDataRange().getDisplayValues();
    var uHeaders = uData[0];

    function findColIdx_(headers, candidates) {
      for (var i = 0; i < headers.length; i++) {
        var h = String(headers[i] || '').toLowerCase().replace(/[^a-z0-9]/g, '');
        for (var c = 0; c < candidates.length; c++) {
          if (h === candidates[c].toLowerCase().replace(/[^a-z0-9]/g, '')) return i;
        }
      }
      return -1;
    }

    var kkCol = findColIdx_(uHeaders, ['No_KK', 'nokk', 'no kk']);
    var ktpCol = findColIdx_(uHeaders, ['No_KTP', 'noktp', 'nik']);
    var namaCol = findColIdx_(uHeaders, ['Nama', 'nama_lengkap']);
    var hpCol = findColIdx_(uHeaders, ['No_HP', 'nohp', 'telepon', 'kontak']);
    var umurCol = findColIdx_(uHeaders, ['Umur', 'usia']);
    var alamatCol = findColIdx_(uHeaders, ['Alamat', 'alamat_domisili']);
    var jkCol = findColIdx_(uHeaders, ['Jenis_Kelamin', 'gender', 'jk']);
    var stRumahCol = findColIdx_(uHeaders, ['Status_Rumah', 'statusrumah', 'rumah']);
    var roleCol = findColIdx_(uHeaders, ['Role', 'peran']);
    var stUserCol = findColIdx_(uHeaders, ['Status_User', 'statususer']);
    var stKeluargaCol = findColIdx_(uHeaders, ['Status_Keluarga', 'statuskeluarga', 'hubungan_keluarga', 'hubungan']);

    // 1. KUMPULKAN ANGGOTA KELUARGA DALAM 1 NO. KK (GABUNGAN SHEET USERS & ANGGOTA_KELUARGA)
    var familyMembersMap = {};
    var seenMemberPerKk = {};

    // Ambil anggota dari sheet USERS
    for (var u = 1; u < uData.length; u++) {
      var rKk = String(kkCol !== -1 ? uData[u][kkCol] : '').trim();
      var rKtp = String(ktpCol !== -1 ? uData[u][ktpCol] : '').trim();
      var rNama = String(namaCol !== -1 ? uData[u][namaCol] : '').trim();
      var rHubungan = String(stKeluargaCol !== -1 ? uData[u][stKeluargaCol] : 'Kepala Keluarga').trim();
      var rJk = String(jkCol !== -1 ? uData[u][jkCol] : '-').trim();
      var rUmur = String(umurCol !== -1 ? uData[u][umurCol] : '-').trim();

      if (!rKk || !rNama) continue;

      if (!familyMembersMap[rKk]) {
        familyMembersMap[rKk] = [];
        seenMemberPerKk[rKk] = {};
      }

      var memberKey = rKtp || (rNama.toLowerCase() + '_' + rUmur);
      if (!seenMemberPerKk[rKk][memberKey]) {
        seenMemberPerKk[rKk][memberKey] = true;
        familyMembersMap[rKk].push({
          nama: rNama,
          hubungan: rHubungan || 'Kepala Keluarga',
          jenisKelamin: rJk,
          umur: rUmur
        });
      }
    }

    // Ambil anggota tambahan dari sheet ANGGOTA_KELUARGA
    var anggotaSheet = ss.getSheetByName('ANGGOTA_KELUARGA');
    if (anggotaSheet && anggotaSheet.getLastRow() > 1) {
      var aData = anggotaSheet.getDataRange().getDisplayValues();
      var aHeaders = aData[0];
      var aKkCol = findColIdx_(aHeaders, ['No_KK', 'nokk', 'no kk']);
      var aKtpCol = findColIdx_(aHeaders, ['No_KTP', 'noktp', 'nik']);
      var aNamaCol = findColIdx_(aHeaders, ['Nama_Anggota', 'nama', 'namalengkap']);
      var aHubCol = findColIdx_(aHeaders, ['Hubungan_Keluarga', 'hubungan', 'statuskeluarga', 'status_keluarga']);
      var aUmurCol = findColIdx_(aHeaders, ['Umur', 'usia']);
      var aJkCol = findColIdx_(aHeaders, ['Jenis_Kelamin', 'gender', 'jk']);

      if (aKkCol !== -1 && aNamaCol !== -1) {
        for (var a = 1; a < aData.length; a++) {
          var aKk = String(aData[a][aKkCol] || '').trim();
          var aKtp = String(aKtpCol !== -1 ? aData[a][aKtpCol] : '').trim();
          var aNama = String(aData[a][aNamaCol] || '').trim();
          var aHub = String(aHubCol !== -1 ? aData[a][aHubCol] : 'Anggota Keluarga').trim();
          var aUmur = String(aUmurCol !== -1 ? aData[a][aUmurCol] : '-').trim();
          var aJk = String(aJkCol !== -1 ? aData[a][aJkCol] : '-').trim();

          if (!aKk || !aNama) continue;

          if (!familyMembersMap[aKk]) {
            familyMembersMap[aKk] = [];
            seenMemberPerKk[aKk] = {};
          }

          var aMemberKey = aKtp || (aNama.toLowerCase() + '_' + aUmur);
          if (!seenMemberPerKk[aKk][aMemberKey]) {
            seenMemberPerKk[aKk][aMemberKey] = true;
            familyMembersMap[aKk].push({
              nama: aNama,
              hubungan: aHub || 'Anggota Keluarga',
              jenisKelamin: aJk,
              umur: aUmur
            });
          }
        }
      }
    }

    // 2. FILTER & KUMPULKAN DAFTAR WARGA
    var filtered = [];
    for (var i = 1; i < uData.length; i++) {
      var row = uData[i];
      var noKk = String(kkCol !== -1 ? row[kkCol] : '').trim();
      var noKtp = String(ktpCol !== -1 ? row[ktpCol] : '').trim();
      var nama = String(namaCol !== -1 ? row[namaCol] : '').trim();
      var noHp = String(hpCol !== -1 ? row[hpCol] : '').trim();
      var alamat = String(alamatCol !== -1 ? row[alamatCol] : '-').trim();
      var jk = String(jkCol !== -1 ? row[jkCol] : '-').trim();
      var umur = String(umurCol !== -1 ? row[umurCol] : '-').trim();
      var rumah = String(stRumahCol !== -1 ? row[stRumahCol] : 'Pribadi').trim();
      var role = String(roleCol !== -1 ? row[roleCol] : 'Warga').trim();
      var statusUser = String(stUserCol !== -1 ? row[stUserCol] : 'Warga').trim();

      if (!nama && !noKk) continue;

      // Filter status rumah (Pribadi / Kontrak)
      if (statusRumah !== 'SEMUA' && rumah.toLowerCase() !== statusRumah.toLowerCase()) {
        continue;
      }

      // Filter abjad nama (A - Z)
      if (alphabet && alphabet !== 'SEMUA') {
        var firstChar = nama.trim().charAt(0).toUpperCase();
        if (firstChar !== alphabet) {
          continue;
        }
      }

      // Filter pencarian
      if (search) {
        var match = (
          nama.toLowerCase().indexOf(search) !== -1 ||
          noKk.indexOf(search) !== -1 ||
          noKtp.indexOf(search) !== -1 ||
          noHp.indexOf(search) !== -1 ||
          alamat.toLowerCase().indexOf(search) !== -1
        );
        if (!match) continue;
      }

      var membersList = familyMembersMap[noKk] || [];
      var totKeluarga = membersList.length > 0 ? membersList.length : 1;

      filtered.push({
        noKk: noKk,
        noKtp: noKtp,
        nama: nama,
        alamat: alamat,
        jenisKelamin: jk,
        umur: umur,
        noHp: noHp,
        statusRumah: rumah,
        role: role,
        statusUser: statusUser,
        jumlahKeluarga: totKeluarga,
        anggotaKeluarga: membersList
      });
    }

    // 3. SORTING ALFABETIS (A SAMPAI Z) BERDASARKAN NAMA
    filtered.sort(function(a, b) {
      var nameA = String(a.nama || '').toLowerCase().trim();
      var nameB = String(b.nama || '').toLowerCase().trim();
      return nameA.localeCompare(nameB, 'id');
    });

    var total = filtered.length;
    var paginated = filtered.slice(offset, offset + limit);
    var totalPages = Math.ceil(total / limit) || 1;
    var currentPage = Math.floor(offset / limit) + 1;

    return {
      success: true,
      data: paginated,
      total: total,
      page: currentPage,
      totalPages: totalPages
    };
  } catch (err) {
    return { success: false, message: 'Gagal memuat data warga: ' + err.toString() };
  }
}

function getAllWargaForExport() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('USERS');
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, data: [] };

    var rawData = sheet.getDataRange().getDisplayValues();
    var headers = rawData[0];
    var list = [];

    for (var i = 1; i < rawData.length; i++) {
      var row = rawData[i];
      var rowObj = {};
      for (var h = 0; h < headers.length; h++) {
        rowObj[headers[h]] = row[h] || '';
      }
      list.push(rowObj);
    }

    // Urutkan data ekspor sesuai alfabet A-Z berdasarkan Nama
    list.sort(function(a, b) {
      var nA = String(a.Nama || a.nama || '').toLowerCase().trim();
      var nB = String(b.Nama || b.nama || '').toLowerCase().trim();
      return nA.localeCompare(nB, 'id');
    });

    return { success: true, data: list };
  } catch (err) {
    return { success: false, message: 'Gagal mengambil data ekspor: ' + err.toString() };
  }
}

function importWargaBatch(parsedData) {
  try {
    if (!parsedData || !Array.isArray(parsedData) || parsedData.length === 0) {
      return { success: false, message: 'Data CSV kosong atau format tidak valid.' };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('USERS');
    if (!sheet) throw new Error('Sheet USERS tidak ditemukan.');

    var rawData = sheet.getDataRange().getDisplayValues();
    var headers = rawData[0];

    function findColIdx_(headers, candidates) {
      for (var i = 0; i < headers.length; i++) {
        var h = String(headers[i] || '').toLowerCase().replace(/[^a-z0-9]/g, '');
        for (var c = 0; c < candidates.length; c++) {
          if (h === candidates[c].toLowerCase().replace(/[^a-z0-9]/g, '')) return i;
        }
      }
      return -1;
    }

    var ktpIdx = findColIdx_(headers, ['No_KTP', 'nik', 'noktp']);
    var kkIdx = findColIdx_(headers, ['No_KK', 'nokk']);
    var namaIdx = findColIdx_(headers, ['Nama', 'namalengkap']);

    var existingKtp = {};
    for (var r = 1; r < rawData.length; r++) {
      var ktpVal = String(ktpIdx !== -1 ? rawData[r][ktpIdx] : '').trim();
      if (ktpVal) existingKtp[ktpVal] = true;
    }

    var importedCount = 0;
    var duplicates = [];
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var newRows = [];

    for (var i = 0; i < parsedData.length; i++) {
      var item = parsedData[i];
      var inputKtp = String(item.No_KTP || item.noKtp || item.NIK || item.nik || '').trim();
      var inputNama = String(item.Nama || item.nama || '').trim();
      var inputKk = String(item.No_KK || item.noKk || '').trim();

      if (!inputNama && !inputKk) continue;

      if (inputKtp && existingKtp[inputKtp]) {
        duplicates.push(inputNama + ' (NIK: ' + inputKtp + ')');
        continue;
      }

      var rowValues = new Array(headers.length);
      for (var c = 0; c < headers.length; c++) {
        var hName = headers[c];
        var val = item[hName] !== undefined ? item[hName] : (item[hName.toLowerCase()] || '');

        if (!val) {
          if (hName === 'No_KK') val = inputKk;
          else if (hName === 'No_KTP') val = inputKtp;
          else if (hName === 'Nama') val = inputNama;
          else if (hName === 'No_HP') val = item.No_HP || item.noHp || item.no_hp || '-';
          else if (hName === 'Password') val = item.Password || item.password || 'warga123';
          else if (hName === 'Role') val = item.Role || item.role || 'Warga';
          else if (hName === 'Status_User') val = item.Status_User || item.statusUser || 'Warga';
          else if (hName === 'Tanggal_Lahir') val = item.Tanggal_Lahir || item.tanggalLahir || '-';
          else if (hName === 'Umur') val = item.Umur || item.umur || '-';
          else if (hName === 'Alamat') val = item.Alamat || item.alamat || 'RT 010 / RW 05 Halim';
          else if (hName === 'Jenis_Kelamin') val = item.Jenis_Kelamin || item.jenisKelamin || 'Laki-laki';
          else if (hName === 'Status_Keluarga') val = item.Status_Keluarga || item.statusKeluarga || 'Kepala Keluarga';
          else if (hName === 'Status_Rumah') val = item.Status_Rumah || item.statusRumah || 'Pribadi';
          else if (hName === 'Pendidikan') val = item.Pendidikan || item.pendidikan || 'SMA';
          else if (hName === 'Pekerjaan') val = item.Pekerjaan || item.pekerjaan || 'Wiraswasta';
          else if (hName === 'Created_At') val = nowStr;
        }

        rowValues[c] = String(val);
      }

      newRows.push(rowValues);
      if (inputKtp) existingKtp[inputKtp] = true;
      importedCount++;
    }

    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, headers.length).setValues(newRows);
      SpreadsheetApp.flush();
    }

    return {
      success: true,
      importedCount: importedCount,
      duplicates: duplicates,
      message: 'Berhasil mengimpor ' + importedCount + ' data warga baru.' + (duplicates.length > 0 ? ' (' + duplicates.length + ' data dilewati karena NIK sudah terdaftar)' : '')
    };
  } catch (err) {
    return { success: false, message: 'Gagal mengimpor data CSV: ' + err.toString() };
  }
}
