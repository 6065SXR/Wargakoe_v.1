/**
 * ====================================================================
 * WARGAKOE - PENGADUAN WARGA BACKEND MODULE
 * Mengelola penyampaian aduan warga, unggah foto bukti terkompresi,
 * notifikasi pengurus (red dot), tindak lanjut aksi RT, serta penghapusan aduan selesai.
 * ====================================================================
 */

function getPengaduanSheetSafe_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('PENGADUAN');
  if (!sheet) {
    sheet = ss.insertSheet('PENGADUAN');
    var headers = [
      'ID_Pengaduan',
      'No_KK',
      'Nama_Pelapor',
      'No_HP',
      'Kategori',
      'Judul_Aduan',
      'Lokasi',
      'Deskripsi',
      'Foto_Bukti',
      'Status',
      'Catatan_Pengurus',
      'Petugas_Penangan',
      'Created_At',
      'Updated_At'
    ];
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setBackground('#221d52');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getPengaduanList(params) {
  try {
    params = params || {};
    var search = String(params.search || '').toLowerCase().trim();
    var filterKategori = String(params.kategori || 'SEMUA').trim();
    var filterStatus = String(params.status || 'SEMUA').trim();
    var onlyMine = Boolean(params.onlyMine);
    var userNoKk = String(params.userNoKk || '').trim();

    var sheet = getPengaduanSheetSafe_();
    var rawData = sheet.getDataRange().getDisplayValues();
    if (rawData.length <= 1) {
      return { success: true, data: [], pendingCount: 0 };
    }

    var headers = rawData[0];
    var idIdx = headers.indexOf('ID_Pengaduan');
    var kkIdx = headers.indexOf('No_KK');
    var namaIdx = headers.indexOf('Nama_Pelapor');
    var hpIdx = headers.indexOf('No_HP');
    var katIdx = headers.indexOf('Kategori');
    var judulIdx = headers.indexOf('Judul_Aduan');
    var lokasiIdx = headers.indexOf('Lokasi');
    var deskIdx = headers.indexOf('Deskripsi');
    var fotoIdx = headers.indexOf('Foto_Bukti');
    var statusIdx = headers.indexOf('Status');
    var catatanIdx = headers.indexOf('Catatan_Pengurus');
    var petugasIdx = headers.indexOf('Petugas_Penangan');
    var createdIdx = headers.indexOf('Created_At');
    var updatedIdx = headers.indexOf('Updated_At');

    var list = [];
    var pendingCount = 0;

    for (var i = rawData.length - 1; i >= 1; i--) {
      var row = rawData[i];
      var rowStatus = String(row[statusIdx] || 'Menunggu').trim();
      var rowKk = String(row[kkIdx] || '').trim();
      var rowNama = String(row[namaIdx] || '').trim();
      var rowJudul = String(row[judulIdx] || '').trim();
      var rowLokasi = String(row[lokasiIdx] || '').trim();
      var rowKat = String(row[katIdx] || 'Lain-lain').trim();

      if (rowStatus.toLowerCase() === 'menunggu') {
        pendingCount++;
      }

      if (onlyMine && userNoKk && rowKk !== userNoKk) {
        continue;
      }

      if (filterKategori !== 'SEMUA' && rowKat.toLowerCase() !== filterKategori.toLowerCase()) {
        continue;
      }

      if (filterStatus !== 'SEMUA' && rowStatus.toLowerCase() !== filterStatus.toLowerCase()) {
        continue;
      }

      var matchSearch = !search || (
        rowJudul.toLowerCase().indexOf(search) !== -1 ||
        rowNama.toLowerCase().indexOf(search) !== -1 ||
        rowLokasi.toLowerCase().indexOf(search) !== -1 ||
        rowKat.toLowerCase().indexOf(search) !== -1
      );

      if (matchSearch) {
        list.push({
          idPengaduan: row[idIdx],
          noKk: rowKk,
          namaPelapor: rowNama,
          noHp: row[hpIdx] || '-',
          kategori: rowKat,
          judulAduan: rowJudul,
          lokasi: rowLokasi,
          deskripsi: row[deskIdx] || '-',
          fotoBukti: row[fotoIdx] || '',
          status: rowStatus,
          catatanPengurus: row[catatanIdx] || '-',
          petugasPenangan: row[petugasIdx] || '-',
          createdAt: row[createdIdx] || '-',
          updatedAt: row[updatedIdx] || '-'
        });
      }
    }

    list.sort(function(a, b) {
      var aIsWait = (a.status.toLowerCase() === 'menunggu');
      var bIsWait = (b.status.toLowerCase() === 'menunggu');
      if (aIsWait && !bIsWait) return -1;
      if (!aIsWait && bIsWait) return 1;
      return 0;
    });

    return {
      success: true,
      data: list,
      pendingCount: pendingCount
    };
  } catch (error) {
    return { success: false, message: 'Gagal memuat daftar pengaduan: ' + error.toString() };
  }
}

function savePengaduan(payload) {
  try {
    var sheet = getPengaduanSheetSafe_();
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var newId = generatePengaduanSequentialId_();

    var judul = String(payload.judulAduan || '').trim();
    var kategori = String(payload.kategori || 'Fasilitas Umum').trim();
    var lokasi = String(payload.lokasi || 'Lingkungan RT 010').trim();
    var deskripsi = String(payload.deskripsi || '').trim();
    var fotoBukti = String(payload.fotoBukti || '').trim();

    if (!judul || !deskripsi) {
      return { success: false, message: 'Judul aduan dan rincian masalah wajib diisi!' };
    }

    sheet.appendRow([
      newId,
      String(payload.noKk || '').trim(),
      String(payload.namaPelapor || 'Warga RT 010').trim(),
      String(payload.noHp || '-').trim(),
      kategori,
      judul,
      lokasi,
      deskripsi,
      fotoBukti,
      'Menunggu',
      '-',
      '-',
      nowStr,
      nowStr
    ]);

    SpreadsheetApp.flush();
    return {
      success: true,
      message: 'Pengaduan berhasil disampaikan! Pengurus RT telah menerima laporan Anda.',
      idPengaduan: newId
    };
  } catch (error) {
    return { success: false, message: 'Gagal mengirim pengaduan: ' + error.toString() };
  }
}

function updateStatusPengaduan(payload) {
  try {
    var sheet = getPengaduanSheetSafe_();
    var data = sheet.getDataRange().getDisplayValues();
    var headers = data[0];

    var idIdx = headers.indexOf('ID_Pengaduan');
    var statusIdx = headers.indexOf('Status');
    var catatanIdx = headers.indexOf('Catatan_Pengurus');
    var petugasIdx = headers.indexOf('Petugas_Penangan');
    var updatedIdx = headers.indexOf('Updated_At');

    var targetId = String(payload.idPengaduan || '').trim();
    var targetRow = -1;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idIdx]).trim() === targetId) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow === -1) {
      return { success: false, message: 'Data pengaduan tidak ditemukan.' };
    }

    // Proteksi status selesai: jika sudah selesai, tidak boleh diubah mundur kembali
    var curStatus = String(data[targetRow - 1][statusIdx]).trim().toLowerCase();
    var newStatus = String(payload.status || 'Diterima').trim();
    if (curStatus === 'selesai' && newStatus.toLowerCase() !== 'selesai') {
      return { success: false, message: 'Laporan pengaduan ini sudah dinyatakan selesai dan tidak dapat diubah kembali.' };
    }

    var nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var catatan = String(payload.catatanPengurus || '-').trim();
    var petugas = String(payload.petugasPenangan || 'Pengurus RT 010').trim();

    sheet.getRange(targetRow, statusIdx + 1).setValue(newStatus);
    sheet.getRange(targetRow, catatanIdx + 1).setValue(catatan);
    sheet.getRange(targetRow, petugasIdx + 1).setValue(petugas);
    sheet.getRange(targetRow, updatedIdx + 1).setValue(nowStr);

    SpreadsheetApp.flush();
    return {
      success: true,
      message: 'Status pengaduan berhasil diperbarui menjadi: ' + newStatus
    };
  } catch (error) {
    return { success: false, message: 'Gagal memperbarui status pengaduan: ' + error.toString() };
  }
}

function deletePengaduan(idPengaduan) {
  try {
    var sheet = getPengaduanSheetSafe_();
    var data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) {
      return { success: false, message: 'Data pengaduan tidak ditemukan.' };
    }

    var idIdx = data[0].indexOf('ID_Pengaduan');
    var targetId = String(idPengaduan || '').trim();
    var targetRow = -1;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idIdx]).trim() === targetId) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow === -1) {
      return { success: false, message: 'Data pengaduan tidak ditemukan atau sudah dihapus.' };
    }

    sheet.deleteRow(targetRow);
    SpreadsheetApp.flush();

    return {
      success: true,
      message: 'Laporan pengaduan ' + targetId + ' berhasil dihapus.'
    };
  } catch (error) {
    return { success: false, message: 'Gagal menghapus pengaduan: ' + error.toString() };
  }
}

function getPengaduanPendingCount() {
  try {
    var sheet = getPengaduanSheetSafe_();
    var data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) return { success: true, pendingCount: 0 };

    var statusIdx = data[0].indexOf('Status');
    var count = 0;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][statusIdx]).toLowerCase().trim() === 'menunggu') {
        count++;
      }
    }
    return { success: true, pendingCount: count };
  } catch (e) {
    return { success: false, pendingCount: 0 };
  }
}

function generatePengaduanSequentialId_() {
  try {
    var sheet = getPengaduanSheetSafe_();
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return 'ADU-0001';

    var data = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
    var maxNum = 0;
    for (var i = 0; i < data.length; i++) {
      var id = String(data[i][0]);
      if (id.indexOf('ADU-') === 0) {
        var num = parseInt(id.replace('ADU-', ''), 10);
        if (!isNaN(num) && num > maxNum) maxNum = num;
      }
    }
    var nextNum = maxNum + 1;
    return 'ADU-' + ('0000' + nextNum).slice(-4);
  } catch (e) {
    return 'ADU-' + Math.floor(1000 + Math.random() * 9000);
  }
}
