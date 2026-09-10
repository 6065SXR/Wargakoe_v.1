/**
 * ====================================================================
 * WARGAKOE - CORE SERVER ENTRYPOINT, UTILITY & MADING/RONDA BACKEND
 * By Zettbos System (ZettBOT 3.1)
 * ====================================================================
 */

function doGet(e) {
  try {
    var template = HtmlService.createTemplateFromFile('index');
    return template.evaluate()
      .setTitle('Wargakoe - RT 010 RW 05 Halim Perdana Kusuma')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return HtmlService.createHtmlOutput('<h3>Terjadi kesalahan memuat aplikasi: ' + err.toString() + '</h3>');
  }
}

function include(filename) {
  try {
    return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
  } catch (e) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  }
}

function getSheet_(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    if (typeof setupDatabase === 'function') setupDatabase();
    sheet = ss.getSheetByName(sheetName);
  }
  return sheet;
}

function generateSequentialId_(prefix, sheetName) {
  var sheet = getSheet_(sheetName);
  if (!sheet) return prefix + '-0001';
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
    if (!sheet) return { success: true, noKuitansi: todayStr + '01' };
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

// ====================================================================
// MODUL BACKEND: MADING & BERITA WARGA (SHEET: MADING)
// ====================================================================

function getMadingList(params) {
  try {
    params = params || {};
    var search = String(params.search || '').toLowerCase().trim();

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('MADING');
    if (!sheet || sheet.getLastRow() <= 1) {
      return { success: true, data: [] };
    }

    var rawData = sheet.getDataRange().getDisplayValues();
    var headers = rawData[0];

    function findCol_(candidates) {
      for (var i = 0; i < headers.length; i++) {
        var h = String(headers[i] || '').toLowerCase().replace(/[^a-z0-9]/g, '');
        for (var c = 0; c < candidates.length; c++) {
          if (h === candidates[c].toLowerCase().replace(/[^a-z0-9]/g, '')) return i;
        }
      }
      return -1;
    }

    var idCol = findCol_(['id_mading', 'idmading', 'id']);
    var judulCol = findCol_(['judul', 'judul_mading', 'title']);
    var tglCol = findCol_(['tanggal_kegiatan', 'tanggal', 'tgl', 'date']);
    var noteCol = findCol_(['note', 'isi', 'deskripsi', 'keterangan', 'pesan']);
    var pubCol = findCol_(['status_publish', 'statuspublish', 'status']);
    var authorCol = findCol_(['pembuat', 'penulis', 'author', 'createdby', 'oleh']);
    var createdCol = findCol_(['created_at', 'createdat', 'timestamp']);

    var list = [];
    for (var i = rawData.length - 1; i >= 1; i--) {
      var row = rawData[i];
      var judul = String(judulCol !== -1 ? row[judulCol] : row[1] || '').trim();
      if (!judul) continue;

      var note = String(noteCol !== -1 ? row[noteCol] : row[3] || '').trim();
      if (search && judul.toLowerCase().indexOf(search) === -1 && note.toLowerCase().indexOf(search) === -1) {
        continue;
      }

      var id = String(idCol !== -1 ? row[idCol] : row[0] || ('MAD-' + i)).trim();
      var tglKegiatan = String(tglCol !== -1 ? row[tglCol] : row[2] || '-').trim();
      var statusPublish = String(pubCol !== -1 ? row[pubCol] : row[4] || 'Published').trim();
      var pembuat = String(authorCol !== -1 ? row[authorCol] : row[5] || 'Pengurus RT 010').trim();
      var createdAt = String(createdCol !== -1 ? row[createdCol] : row[6] || '-').trim();

      list.push({
        idMading: id,
        judul: judul,
        tanggalKegiatan: tglKegiatan,
        note: note,
        statusPublish: statusPublish || 'Published',
        pembuat: pembuat,
        createdAt: createdAt
      });
    }

    return { success: true, data: list };
  } catch (err) {
    return { success: false, message: 'Gagal memuat mading: ' + err.toString(), data: [] };
  }
}

function getMadingData() { return getMadingList(); }
function getRecentMading() { return getMadingList(); }

function saveMading(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('MADING');
    if (!sheet) {
      if (typeof setupDatabase === 'function') setupDatabase();
      sheet = ss.getSheetByName('MADING');
    }

    var nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var judul = String(payload.judul || '').trim();
    var note = String(payload.note || payload.isi || '').trim();
    var tglKegiatan = String(payload.tanggalKegiatan || payload.tanggal || Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy')).trim();
    var pembuat = String(payload.pembuat || 'Pengurus RT').trim();
    var status = String(payload.statusPublish || 'Published').trim();
    var targetId = String(payload.idMading || '').trim();

    if (!judul || !note) {
      return { success: false, message: 'Judul dan isi pengumuman wajib diisi!' };
    }

    if (targetId && sheet.getLastRow() > 1) {
      var data = sheet.getDataRange().getDisplayValues();
      for (var r = 1; r < data.length; r++) {
        if (String(data[r][0]).trim() === targetId) {
          sheet.getRange(r + 1, 2).setValue(judul);
          sheet.getRange(r + 1, 3).setValue(tglKegiatan);
          sheet.getRange(r + 1, 4).setValue(note);
          sheet.getRange(r + 1, 5).setValue(status);
          sheet.getRange(r + 1, 6).setValue(pembuat);
          SpreadsheetApp.flush();
          return { success: true, message: 'Pengumuman mading berhasil diperbarui!' };
        }
      }
    }

    var newId = generateSequentialId_('MAD', 'MADING');
    sheet.appendRow([newId, judul, tglKegiatan, note, status, pembuat, nowStr]);
    SpreadsheetApp.flush();

    return { success: true, message: 'Pengumuman mading berhasil diterbitkan!' };
  } catch (err) {
    return { success: false, message: 'Gagal menyimpan mading: ' + err.toString() };
  }
}

function deleteMading(idMading) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('MADING');
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: 'Data mading tidak ditemukan.' };

    var data = sheet.getDataRange().getDisplayValues();
    var targetId = String(idMading || '').trim();
    var targetRow = -1;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === targetId) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow === -1) return { success: false, message: 'Pengumuman tidak ditemukan.' };

    sheet.deleteRow(targetRow);
    SpreadsheetApp.flush();
    return { success: true, message: 'Pengumuman mading berhasil dihapus.' };
  } catch (err) {
    return { success: false, message: 'Gagal menghapus mading: ' + err.toString() };
  }
}

// ====================================================================
// MODUL BACKEND: JADWAL & ABSENSI RONDA (SHEET: JADWAL_RONDA & ABSENSI_RONDA)
// ====================================================================

function parseWargaRondaListSafe_(rawVal) {
  if (!rawVal) return [];
  if (Array.isArray(rawVal)) {
    return rawVal.map(function(item) {
      if (typeof item === 'string') return { noKk: '', nama: item, noHp: '-', peran: 'Anggota' };
      return {
        noKk: String(item.noKk || item.nokk || ''),
        nama: String(item.nama || item.name || 'Warga'),
        noHp: String(item.noHp || item.nohp || '-'),
        peran: String(item.peran || item.role || 'Anggota')
      };
    });
  }

  var strVal = String(rawVal).trim();
  if (strVal.indexOf('[') === 0 && strVal.lastIndexOf(']') === strVal.length - 1) {
    try {
      var parsed = JSON.parse(strVal);
      if (Array.isArray(parsed)) {
        return parsed.map(function(item) {
          if (typeof item === 'string') return { noKk: '', nama: item, noHp: '-', peran: 'Anggota' };
          return {
            noKk: String(item.noKk || item.nokk || ''),
            nama: String(item.nama || item.name || 'Warga'),
            noHp: String(item.noHp || item.nohp || '-'),
            peran: String(item.peran || item.role || 'Anggota')
          };
        });
      }
    } catch (e) {}
  }

  var parts = strVal.split(',');
  var list = [];
  for (var p = 0; p < parts.length; p++) {
    var itemStr = parts[p].trim();
    if (itemStr) {
      list.push({ noKk: '', nama: itemStr, noHp: '-', peran: 'Anggota' });
    }
  }
  return list;
}

function getRondaSchedules() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('JADWAL_RONDA');
    if (!sheet || sheet.getLastRow() <= 1) {
      return { success: true, data: [], activeRonda: null };
    }

    var rawData = sheet.getDataRange().getDisplayValues();
    var headers = rawData[0];

    function findCol_(candidates) {
      for (var i = 0; i < headers.length; i++) {
        var h = String(headers[i] || '').toLowerCase().replace(/[^a-z0-9]/g, '');
        for (var c = 0; c < candidates.length; c++) {
          if (h === candidates[c].toLowerCase().replace(/[^a-z0-9]/g, '')) return i;
        }
      }
      return -1;
    }

    var idCol = findCol_(['id_jadwal', 'idjadwal', 'id']);
    var reguCol = findCol_(['nama_regu', 'namaregu', 'regu', 'nama']);
    var tglCol = findCol_(['tanggal_ronda', 'tanggal', 'tgl', 'date']);
    var ketuaCol = findCol_(['ketua_regu', 'ketuaregu', 'ketua', 'koordinator']);
    var posCol = findCol_(['lokasi_pos', 'lokasipos', 'pos', 'lokasi']);
    var shiftCol = findCol_(['jam_shift', 'jamshift', 'jam', 'waktu']);
    var wargaCol = findCol_(['daftar_warga_json', 'daftarwarga', 'warga', 'anggota', 'petugas']);
    var noteCol = findCol_(['catatan', 'note', 'keterangan']);

    var list = [];
    for (var i = 1; i < rawData.length; i++) {
      var row = rawData[i];
      var namaRegu = String(reguCol !== -1 ? row[reguCol] : row[1] || '').trim();
      if (!namaRegu) continue;

      var idJadwal = String(idCol !== -1 ? row[idCol] : row[0] || ('RND-' + i)).trim();
      var tglRonda = String(tglCol !== -1 ? row[tglCol] : row[2] || '-').trim();
      var ketuaRegu = String(ketuaCol !== -1 ? row[ketuaCol] : row[3] || 'Koordinator').trim();
      var lokasiPos = String(posCol !== -1 ? row[posCol] : row[4] || 'Pos Ronda Utama RT 010').trim();
      var jamShift = String(shiftCol !== -1 ? row[shiftCol] : row[5] || '22.00 - 03.00 WIB').trim();
      var rawWarga = wargaCol !== -1 ? row[wargaCol] : row[6] || '';
      var catatan = String(noteCol !== -1 ? row[noteCol] : row[7] || '').trim();

      var normalizedWarga = parseWargaRondaListSafe_(rawWarga);

      list.push({
        idJadwal: idJadwal,
        namaRegu: namaRegu,
        tanggalRonda: tglRonda,
        ketuaRegu: ketuaRegu,
        lokasiPos: lokasiPos,
        jamShift: jamShift,
        catatan: catatan,
        daftarWarga: normalizedWarga,
        wargaList: normalizedWarga
      });
    }

    var activeRonda = list.length > 0 ? list[list.length - 1] : null;

    return {
      success: true,
      data: list,
      activeRonda: activeRonda
    };
  } catch (err) {
    return { success: false, message: 'Gagal memuat jadwal ronda: ' + err.toString(), data: [], activeRonda: null };
  }
}

// Alias agar kompatibel dengan pemanggilan js_ronda.html
function getJadwalRondaList() { return getRondaSchedules(); }
function getRondaData() { return getRondaSchedules(); }
function getActiveRonda() { return getRondaSchedules(); }

function saveJadwalRonda(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('JADWAL_RONDA');
    if (!sheet) {
      if (typeof setupDatabase === 'function') setupDatabase();
      sheet = ss.getSheetByName('JADWAL_RONDA');
    }

    var nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var newId = generateSequentialId_('RND', 'JADWAL_RONDA');

    var namaRegu = String(payload.namaRegu || 'Regu Ronda').trim();
    var tglRonda = String(payload.tanggalRonda || '-').trim();
    var ketuaRegu = String(payload.ketuaRegu || '-').trim();
    var lokasiPos = String(payload.lokasiPos || 'Pos Ronda RT 010').trim();
    var jamShift = String(payload.jamShift || '22.00 - 03.00 WIB').trim();
    var catatan = String(payload.catatan || '').trim();
    var wargaJson = JSON.stringify(payload.daftarWarga || []);

    sheet.appendRow([newId, namaRegu, tglRonda, ketuaRegu, lokasiPos, jamShift, wargaJson, catatan, nowStr]);
    SpreadsheetApp.flush();

    return { success: true, message: 'Jadwal ronda berhasil disimpan!' };
  } catch (err) {
    return { success: false, message: 'Gagal menyimpan jadwal: ' + err.toString() };
  }
}

function deleteJadwalRonda(idJadwal) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('JADWAL_RONDA');
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: 'Data tidak ditemukan.' };

    var data = sheet.getDataRange().getDisplayValues();
    var targetId = String(idJadwal || '').trim();
    var targetRow = -1;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === targetId) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow === -1) return { success: false, message: 'Jadwal tidak ditemukan.' };

    sheet.deleteRow(targetRow);
    SpreadsheetApp.flush();
    return { success: true, message: 'Jadwal ronda berhasil dihapus.' };
  } catch (err) {
    return { success: false, message: 'Gagal menghapus jadwal: ' + err.toString() };
  }
}

function submitAbsenRonda(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('ABSENSI_RONDA');
    if (!sheet) {
      if (typeof setupDatabase === 'function') setupDatabase();
      sheet = ss.getSheetByName('ABSENSI_RONDA');
    }

    var now = new Date();
    var nowStr = Utilities.formatDate(now, 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var timeStr = Utilities.formatDate(now, 'Asia/Jakarta', 'HH:mm:ss') + ' WIB';
    var newId = generateSequentialId_('ABS', 'ABSENSI_RONDA');

    sheet.appendRow([
      newId,
      String(payload.idJadwal || '').trim(),
      String(payload.tanggalRonda || Utilities.formatDate(now, 'Asia/Jakarta', 'dd/MM/yyyy')).trim(),
      String(payload.noKk || '').trim(),
      String(payload.namaWarga || 'Warga RT 010').trim(),
      timeStr,
      'Hadir (Selfie Pos)',
      String(payload.fotoSelfie || ''),
      nowStr
    ]);
    SpreadsheetApp.flush();

    return { success: true, message: 'Presensi selfie ronda berhasil dicatat!' };
  } catch (err) {
    return { success: false, message: 'Gagal mencatat presensi: ' + err.toString() };
  }
}

function getRekapAbsenRonda(idJadwal) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('ABSENSI_RONDA');
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, data: [] };

    var data = sheet.getDataRange().getDisplayValues();
    var targetId = String(idJadwal || '').trim();
    var list = [];

    for (var i = data.length - 1; i >= 1; i--) {
      var row = data[i];
      if (!targetId || String(row[1]).trim() === targetId) {
        list.push({
          idAbsen: row[0],
          idJadwal: row[1],
          tanggalRonda: row[2],
          namaWarga: row[4],
          waktuAbsen: row[5],
          statusHadir: row[6] || 'Hadir',
          fotoSelfie: row[7] || ''
        });
      }
    }

    return { success: true, data: list };
  } catch (e) {
    return { success: false, data: [] };
  }
}

function getRondaAbsensiToday() {
  return getRekapAbsenRonda('');
}

// ====================================================================
// MODUL BACKEND: DATA KELUARGA (SHEET: USERS & ANGGOTA_KELUARGA)
// ====================================================================

function getFamilyDetails(noKk) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var targetKk = String(noKk || '').trim();
    var userSheet = ss.getSheetByName('USERS');
    var kepalaKeluarga = null;

    if (userSheet && userSheet.getLastRow() > 1) {
      var uData = userSheet.getDataRange().getDisplayValues();
      var uHeaders = uData[0];
      var kkIdx = -1;
      for (var h = 0; h < uHeaders.length; h++) {
        var hClean = String(uHeaders[h]).toLowerCase().replace(/[^a-z0-9]/g, '');
        if (hClean === 'nokk' || hClean === 'no_kk') kkIdx = h;
      }

      for (var r = 1; r < uData.length; r++) {
        var rKk = String(kkIdx !== -1 ? uData[r][kkIdx] : uData[r][0]).trim();
        if (rKk === targetKk) {
          kepalaKeluarga = {
            noKk: rKk,
            noKtp: uData[r][1] || '',
            nama: uData[r][2] || '',
            noHp: uData[r][3] || '',
            role: uData[r][5] || 'Warga',
            statusUser: uData[r][6] || 'Warga',
            tanggalLahir: uData[r][7] || '',
            umur: uData[r][8] || '',
            alamat: uData[r][9] || '',
            jenisKelamin: uData[r][10] || 'Laki-laki',
            statusKeluarga: uData[r][11] || 'Kepala Keluarga',
            statusRumah: uData[r][12] || 'Pribadi',
            pendidikan: uData[r][13] || 'SMA',
            pekerjaan: uData[r][14] || 'Wiraswasta'
          };
          break;
        }
      }
    }

    var anggotaList = [];
    var aSheet = ss.getSheetByName('ANGGOTA_KELUARGA');
    if (aSheet && aSheet.getLastRow() > 1) {
      var aData = aSheet.getDataRange().getDisplayValues();
      for (var a = 1; a < aData.length; a++) {
        var aKk = String(aData[a][1] || '').trim();
        if (aKk === targetKk) {
          anggotaList.push({
            idAnggota: aData[a][0] || '',
            noKk: aKk,
            noKtp: aData[a][2] || '',
            namaAnggota: aData[a][3] || '',
            hubunganKeluarga: aData[a][4] || 'Anggota',
            tanggalLahir: aData[a][5] || '',
            umur: aData[a][6] || '',
            jenisKelamin: aData[a][7] || 'Perempuan'
          });
        }
      }
    }

    return {
      success: true,
      data: {
        kepalaKeluarga: kepalaKeluarga,
        anggota: anggotaList
      }
    };
  } catch (err) {
    return { success: false, message: 'Gagal memuat keluarga: ' + err.toString() };
  }
}

function updateProfilWarga(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('USERS');
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: 'Sheet USERS kosong.' };

    var data = sheet.getDataRange().getDisplayValues();
    var origKk = String(payload.origNoKk || payload.noKk || '').trim();
    var targetRow = -1;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === origKk) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow === -1) return { success: false, message: 'Data warga tidak ditemukan.' };

    if (payload.noKtp) sheet.getRange(targetRow, 2).setValue(payload.noKtp);
    if (payload.nama) sheet.getRange(targetRow, 3).setValue(payload.nama);
    if (payload.noHp) sheet.getRange(targetRow, 4).setValue(payload.noHp);
    if (payload.password) sheet.getRange(targetRow, 5).setValue(payload.password);
    if (payload.tanggalLahir) sheet.getRange(targetRow, 8).setValue(payload.tanggalLahir);
    if (payload.umur) sheet.getRange(targetRow, 9).setValue(payload.umur);
    if (payload.alamat) sheet.getRange(targetRow, 10).setValue(payload.alamat);
    if (payload.jenisKelamin) sheet.getRange(targetRow, 11).setValue(payload.jenisKelamin);
    if (payload.statusKeluarga) sheet.getRange(targetRow, 12).setValue(payload.statusKeluarga);
    if (payload.statusRumah) sheet.getRange(targetRow, 13).setValue(payload.statusRumah);
    if (payload.pendidikan) sheet.getRange(targetRow, 14).setValue(payload.pendidikan);
    if (payload.pekerjaan) sheet.getRange(targetRow, 15).setValue(payload.pekerjaan);

    SpreadsheetApp.flush();

    return {
      success: true,
      message: 'Profil warga berhasil diperbarui!',
      user: {
        noKk: String(data[targetRow - 1][0]),
        noKtp: payload.noKtp || data[targetRow - 1][1],
        nama: payload.nama || data[targetRow - 1][2],
        noHp: payload.noHp || data[targetRow - 1][3],
        role: data[targetRow - 1][5],
        statusUser: data[targetRow - 1][6],
        alamat: payload.alamat || data[targetRow - 1][9]
      }
    };
  } catch (err) {
    return { success: false, message: 'Gagal update profil: ' + err.toString() };
  }
}

function saveAnggotaKeluarga(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('ANGGOTA_KELUARGA');
    if (!sheet) {
      if (typeof setupDatabase === 'function') setupDatabase();
      sheet = ss.getSheetByName('ANGGOTA_KELUARGA');
    }

    var nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var targetId = String(payload.idAnggota || '').trim();

    if (targetId && sheet.getLastRow() > 1) {
      var data = sheet.getDataRange().getDisplayValues();
      for (var r = 1; r < data.length; r++) {
        if (String(data[r][0]).trim() === targetId) {
          sheet.getRange(r + 1, 3).setValue(payload.noKtp || '');
          sheet.getRange(r + 1, 4).setValue(payload.namaAnggota || '');
          sheet.getRange(r + 1, 5).setValue(payload.hubunganKeluarga || '');
          sheet.getRange(r + 1, 6).setValue(payload.tanggalLahir || '');
          sheet.getRange(r + 1, 7).setValue(payload.umur || '');
          sheet.getRange(r + 1, 8).setValue(payload.jenisKelamin || '');
          SpreadsheetApp.flush();
          return { success: true, message: 'Data anggota keluarga berhasil diperbarui!' };
        }
      }
    }

    var newId = generateSequentialId_('AGT', 'ANGGOTA_KELUARGA');
    sheet.appendRow([
      newId,
      String(payload.noKk || '').trim(),
      String(payload.noKtp || '').trim(),
      String(payload.namaAnggota || '').trim(),
      String(payload.hubunganKeluarga || 'Anggota').trim(),
      String(payload.tanggalLahir || '-').trim(),
      String(payload.umur || '0').trim(),
      String(payload.jenisKelamin || 'Perempuan').trim(),
      nowStr
    ]);
    SpreadsheetApp.flush();

    return { success: true, message: 'Anggota keluarga berhasil ditambahkan!' };
  } catch (err) {
    return { success: false, message: 'Gagal menyimpan anggota: ' + err.toString() };
  }
}

function deleteAnggotaKeluarga(idAnggota) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('ANGGOTA_KELUARGA');
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: 'Data tidak ditemukan.' };

    var data = sheet.getDataRange().getDisplayValues();
    var targetId = String(idAnggota || '').trim();
    var targetRow = -1;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === targetId) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow === -1) return { success: false, message: 'Anggota tidak ditemukan.' };

    sheet.deleteRow(targetRow);
    SpreadsheetApp.flush();
    return { success: true, message: 'Anggota keluarga berhasil dihapus.' };
  } catch (err) {
    return { success: false, message: 'Gagal menghapus anggota: ' + err.toString() };
  }
}

// ====================================================================
// MODUL BACKEND: DATA RINGKASAN DASHBOARD UTAMA
// ====================================================================

function getDashboardData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Hitung Total Warga & KK
    var totalWarga = 0;
    var kkSet = {};

    var userSheet = ss.getSheetByName('USERS');
    if (userSheet && userSheet.getLastRow() > 1) {
      var uData = userSheet.getDataRange().getDisplayValues();
      var uHeaders = uData[0];
      var kkIdx = -1, namaIdx = -1;
      for (var h = 0; h < uHeaders.length; h++) {
        var hClean = String(uHeaders[h]).toLowerCase().replace(/[^a-z0-9]/g, '');
        if (hClean === 'nokk' || hClean === 'no_kk') kkIdx = h;
        if (hClean === 'nama') namaIdx = h;
      }
      for (var u = 1; u < uData.length; u++) {
        var rowKk = String(kkIdx !== -1 ? uData[u][kkIdx] : '').trim();
        var rowNama = String(namaIdx !== -1 ? uData[u][namaIdx] : '').trim();
        if (rowNama || rowKk) {
          totalWarga++;
          if (rowKk) kkSet[rowKk] = true;
        }
      }
    }

    var anggotaSheet = ss.getSheetByName('ANGGOTA_KELUARGA');
    if (anggotaSheet && anggotaSheet.getLastRow() > 1) {
      var aData = anggotaSheet.getDataRange().getDisplayValues();
      for (var a = 1; a < aData.length; a++) {
        var aNama = String(aData[a][3] || aData[a][2] || '').trim();
        if (aNama) totalWarga++;
      }
    }

    var totalKk = Object.keys(kkSet).length;

    // 2. Hitung Kas Terkumpul & Menunggu Approval dari sheet IURAN
    var kasTerkumpul = 0;
    var menungguApprovalCount = 0;
    var iurSheet = ss.getSheetByName('IURAN');
    if (iurSheet && iurSheet.getLastRow() > 1) {
      var iData = iurSheet.getDataRange().getDisplayValues();
      var iHeaders = iData[0];
      var statusCol = -1, nomCol = -1;
      for (var ih = 0; ih < iHeaders.length; ih++) {
        var ihClean = String(iHeaders[ih]).toLowerCase().replace(/[^a-z0-9]/g, '');
        if (ihClean.indexOf('status') !== -1) statusCol = ih;
        if (ihClean === 'nominal' || ihClean === 'jumlah') nomCol = ih;
      }

      for (var r = 1; r < iData.length; r++) {
        var st = String(statusCol !== -1 ? iData[r][statusCol] : '').toLowerCase().trim();
        var rawNom = nomCol !== -1 ? iData[r][nomCol] : 0;
        var nomVal = parseInt(String(rawNom).replace(/[^0-9]/g, ''), 10) || 0;

        if (st === 'sudah bayar' || st === 'lunas' || st === 'approved') {
          kasTerkumpul += nomVal;
        } else if (st === 'menunggu' || st === 'pending') {
          menungguApprovalCount++;
        }
      }
    }

    // 3. Ambil Pengumuman Mading Terkini (Maksimal 3 pengumuman)
    var madingRes = getMadingList();
    var recentMading = (madingRes && madingRes.success && madingRes.data) ? madingRes.data.slice(0, 3) : [];

    // 4. Ambil Jadwal Ronda Aktif
    var rondaRes = getRondaSchedules();
    var activeRonda = (rondaRes && rondaRes.success) ? (rondaRes.activeRonda || (rondaRes.data && rondaRes.data.length > 0 ? rondaRes.data[0] : null)) : null;

    return {
      success: true,
      totalWarga: totalWarga,
      totalKk: totalKk,
      kasTerkumpul: kasTerkumpul,
      menungguApproval: menungguApprovalCount,
      recentMading: recentMading,
      activeRonda: activeRonda
    };
  } catch (err) {
    return {
      success: false,
      message: 'Gagal memuat dashboard: ' + err.toString(),
      totalWarga: 0,
      totalKk: 0,
      kasTerkumpul: 0,
      menungguApproval: 0,
      recentMading: [],
      activeRonda: null
    };
  }
}

function getDashboardSummary() { return getDashboardData(); }
