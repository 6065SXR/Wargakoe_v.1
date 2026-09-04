/**
 * ====================================================================
 * WARGAKOE - RONDA NIGHT SCHEDULE & SELFIE ATTENDANCE BACKEND
 * By Zettbos System (ZettBOT 3.1)
 * ====================================================================
 */

function getJadwalRondaList() {
  try {
    var sheet = getSheet_('JADWAL_RONDA');
    var rawData = sheet.getDataRange().getDisplayValues();
    if (rawData.length <= 1) return { success: true, data: [] };

    var headers = rawData[0];
    var list = [];

    for (var i = 1; i < rawData.length; i++) {
      var row = rawData[i];
      var daftarWarga = [];
      try {
        daftarWarga = JSON.parse(row[headers.indexOf('Daftar_Warga_JSON')] || '[]');
      } catch (e) {
        daftarWarga = [];
      }

      list.push({
        idJadwal: row[headers.indexOf('ID_Jadwal')],
        namaRegu: row[headers.indexOf('Nama_Regu')],
        tanggalRonda: row[headers.indexOf('Tanggal_Ronda')],
        ketuaRegu: row[headers.indexOf('Ketua_Regu')],
        lokasiPos: row[headers.indexOf('Lokasi_Pos')],
        jamShift: row[headers.indexOf('Jam_Shift')],
        daftarWarga: daftarWarga,
        catatan: row[headers.indexOf('Catatan')],
        createdAt: row[headers.indexOf('Created_At')]
      });
    }

    list.sort(function(a, b) {
      return new Date(a.tanggalRonda) - new Date(b.tanggalRonda);
    });

    return { success: true, data: list };
  } catch (error) {
    return { success: false, message: 'Gagal memuat jadwal ronda: ' + error.toString() };
  }
}

function saveJadwalRonda(payload) {
  try {
    var sheet = getSheet_('JADWAL_RONDA');
    var data = sheet.getDataRange().getDisplayValues();
    var headers = data[0];
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var isEdit = payload.idJadwal && payload.idJadwal.trim() !== '';

    var jsonWargaStr = JSON.stringify(payload.daftarWarga || []);

    if (isEdit) {
      for (var i = 1; i < data.length; i++) {
        if (data[i][headers.indexOf('ID_Jadwal')] === payload.idJadwal) {
          var r = i + 1;
          sheet.getRange(r, headers.indexOf('Nama_Regu') + 1).setValue(payload.namaRegu);
          sheet.getRange(r, headers.indexOf('Tanggal_Ronda') + 1).setValue(payload.tanggalRonda);
          sheet.getRange(r, headers.indexOf('Ketua_Regu') + 1).setValue(payload.ketuaRegu);
          sheet.getRange(r, headers.indexOf('Lokasi_Pos') + 1).setValue(payload.lokasiPos);
          sheet.getRange(r, headers.indexOf('Jam_Shift') + 1).setValue(payload.jamShift);
          sheet.getRange(r, headers.indexOf('Daftar_Warga_JSON') + 1).setValue(jsonWargaStr);
          sheet.getRange(r, headers.indexOf('Catatan') + 1).setValue(payload.catatan);
          SpreadsheetApp.flush();
          return { success: true, message: 'Jadwal ronda regu berhasil diperbarui!' };
        }
      }
    }

    var newId = generateSequentialId_('RND', 'JADWAL_RONDA');
    sheet.appendRow([
      newId,
      String(payload.namaRegu || '').trim(),
      String(payload.tanggalRonda || '').trim(),
      String(payload.ketuaRegu || '').trim(),
      String(payload.lokasiPos || 'Pos Ronda RT 010').trim(),
      String(payload.jamShift || '22.00 - 03.00 WIB').trim(),
      jsonWargaStr,
      String(payload.catatan || '').trim(),
      nowStr
    ]);

    SpreadsheetApp.flush();
    return { success: true, message: 'Jadwal ronda malam minggu berhasil ditambahkan!', idJadwal: newId };
  } catch (error) {
    return { success: false, message: 'Gagal menyimpan jadwal ronda: ' + error.toString() };
  }
}

function deleteJadwalRonda(idJadwal) {
  try {
    var sheet = getSheet_('JADWAL_RONDA');
    var data = sheet.getDataRange().getDisplayValues();
    var idIdx = data[0].indexOf('ID_Jadwal');

    for (var i = 1; i < data.length; i++) {
      if (data[i][idIdx] === idJadwal) {
        sheet.deleteRow(i + 1);
        SpreadsheetApp.flush();
        return { success: true, message: 'Jadwal ronda regu berhasil dihapus!' };
      }
    }
    return { success: false, message: 'ID Jadwal tidak ditemukan.' };
  } catch (error) {
    return { success: false, message: 'Gagal menghapus jadwal ronda: ' + error.toString() };
  }
}

function submitAbsenRonda(payload) {
  try {
    var sheet = getSheet_('ABSENSI_RONDA');
    var data = sheet.getDataRange().getDisplayValues();
    var headers = data[0];
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var todayIso = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');

    // Cek duplikasi absen hari ini
    var idJadwalIdx = headers.indexOf('ID_Jadwal');
    var kkIdx = headers.indexOf('No_KK');
    var tglIdx = headers.indexOf('Tanggal_Ronda');

    for (var i = 1; i < data.length; i++) {
      if (data[i][idJadwalIdx] === payload.idJadwal && data[i][kkIdx] === String(payload.noKk).trim() && data[i][tglIdx] === payload.tanggalRonda) {
        return { success: false, message: 'Anda sudah melakukan absensi selfie untuk jadwal ronda hari ini!' };
      }
    }

    var newId = generateSequentialId_('ABS', 'ABSENSI_RONDA');
    sheet.appendRow([
      newId,
      String(payload.idJadwal || '').trim(),
      String(payload.tanggalRonda || todayIso).trim(),
      String(payload.noKk || '').trim(),
      String(payload.namaWarga || '').trim(),
      nowStr,
      'Hadir (Selfie Pos)',
      String(payload.fotoSelfie || '').trim(),
      nowStr
    ]);

    SpreadsheetApp.flush();
    return { success: true, message: 'Absensi selfie ronda berhasil disimpan! Terima kasih atas partisipasinya.' };
  } catch (error) {
    return { success: false, message: 'Gagal merekam absensi selfie: ' + error.toString() };
  }
}

function getRekapAbsenRonda(idJadwal) {
  try {
    var sheet = getSheet_('ABSENSI_RONDA');
    var data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) return { success: true, data: [] };

    var headers = data[0];
    var list = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!idJadwal || row[headers.indexOf('ID_Jadwal')] === idJadwal) {
        list.push({
          idAbsen: row[headers.indexOf('ID_Absen')],
          idJadwal: row[headers.indexOf('ID_Jadwal')],
          tanggalRonda: row[headers.indexOf('Tanggal_Ronda')],
          noKk: row[headers.indexOf('No_KK')],
          namaWarga: row[headers.indexOf('Nama_Warga')],
          waktuAbsen: row[headers.indexOf('Waktu_Absen')],
          statusHadir: row[headers.indexOf('Status_Hadir')],
          fotoSelfie: row[headers.indexOf('Foto_Selfie')]
        });
      }
    }

    return { success: true, data: list.reverse() };
  } catch (error) {
    return { success: false, message: 'Gagal mengambil rekap absensi: ' + error.toString() };
  }
}

function getRondaWargaNotification(noKk) {
  try {
    if (!noKk) return { hasRonda: false };

    var rondaSheet = getSheet_('JADWAL_RONDA');
    var rawData = rondaSheet.getDataRange().getDisplayValues();
    if (rawData.length <= 1) return { hasRonda: false };

    var headers = rawData[0];
    var todayObj = new Date();
    todayObj.setHours(0,0,0,0);

    var upcoming = null;

    for (var i = 1; i < rawData.length; i++) {
      var row = rawData[i];
      var tglStr = row[headers.indexOf('Tanggal_Ronda')];
      var rDate = new Date(tglStr);
      rDate.setHours(0,0,0,0);

      var daftarWarga = [];
      try {
        daftarWarga = JSON.parse(row[headers.indexOf('Daftar_Warga_JSON')] || '[]');
      } catch (e) {
        daftarWarga = [];
      }

      var isMember = false;
      for (var w = 0; w < daftarWarga.length; w++) {
        if (daftarWarga[w].noKk === String(noKk).trim()) {
          isMember = true;
          break;
        }
      }

      if (isMember && rDate >= todayObj) {
        if (!upcoming || rDate < new Date(upcoming.tanggalRonda)) {
          // Cek status absen warga ini
          var isAbsen = checkWargaAlreadyAbsen_(row[headers.indexOf('ID_Jadwal')], noKk, tglStr);

          upcoming = {
            idJadwal: row[headers.indexOf('ID_Jadwal')],
            namaRegu: row[headers.indexOf('Nama_Regu')],
            tanggalRonda: tglStr,
            ketuaRegu: row[headers.indexOf('Ketua_Regu')],
            lokasiPos: row[headers.indexOf('Lokasi_Pos')],
            jamShift: row[headers.indexOf('Jam_Shift')],
            catatan: row[headers.indexOf('Catatan')],
            daftarWarga: daftarWarga,
            isToday: (rDate.getTime() === todayObj.getTime()),
            alreadyAbsen: isAbsen
          };
        }
      }
    }

    return upcoming ? { hasRonda: true, ronda: upcoming } : { hasRonda: false };
  } catch (error) {
    return { hasRonda: false, error: error.toString() };
  }
}

function checkWargaAlreadyAbsen_(idJadwal, noKk, tanggalRonda) {
  var sheet = getSheet_('ABSENSI_RONDA');
  var data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return false;

  var headers = data[0];
  var idJadwalIdx = headers.indexOf('ID_Jadwal');
  var kkIdx = headers.indexOf('No_KK');
  var tglIdx = headers.indexOf('Tanggal_Ronda');

  for (var i = 1; i < data.length; i++) {
    if (data[i][idJadwalIdx] === idJadwal && data[i][kkIdx] === String(noKk).trim() && data[i][tglIdx] === tanggalRonda) {
      return true;
    }
  }
  return false;
}
