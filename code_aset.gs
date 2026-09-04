/**
 * ====================================================================
 * WARGAKOE - ASSET INVENTORY & BORROWING BACKEND MODULE
 * By Zettbos System (ZettBOT 3.1)
 * ====================================================================
 */

function getAsetList() {
  try {
    var sheet = getSheet_('ASET');
    var data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) return { success: true, data: [] };

    var headers = data[0];
    var list = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      list.push({
        idAset: row[headers.indexOf('ID_Aset')],
        namaAset: row[headers.indexOf('Nama_Aset')],
        kategori: row[headers.indexOf('Kategori')],
        jumlahTotal: parseInt(row[headers.indexOf('Jumlah_Total')], 10) || 0,
        jumlahTersedia: parseInt(row[headers.indexOf('Jumlah_Tersedia')], 10) || 0,
        kondisi: row[headers.indexOf('Kondisi')],
        lokasiPenyimpanan: row[headers.indexOf('Lokasi_Penyimpanan')],
        keterangan: row[headers.indexOf('Keterangan')],
        createdAt: row[headers.indexOf('Created_At')]
      });
    }
    return { success: true, data: list };
  } catch (error) {
    return { success: false, message: 'Gagal mengambil data aset: ' + error.toString() };
  }
}

function saveAset(payload) {
  try {
    var sheet = getSheet_('ASET');
    var data = sheet.getDataRange().getDisplayValues();
    var headers = data[0];
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var isEdit = payload.idAset && payload.idAset.trim() !== '';

    var totalBaru = parseInt(payload.jumlahTotal, 10) || 0;

    if (isEdit) {
      for (var i = 1; i < data.length; i++) {
        if (data[i][headers.indexOf('ID_Aset')] === payload.idAset) {
          var r = i + 1;
          var totalLama = parseInt(data[i][headers.indexOf('Jumlah_Total')], 10) || 0;
          var tersediaLama = parseInt(data[i][headers.indexOf('Jumlah_Tersedia')], 10) || 0;
          var sedangDipinjam = Math.max(0, totalLama - tersediaLama);
          var tersediaBaru = Math.max(0, totalBaru - sedangDipinjam);

          sheet.getRange(r, headers.indexOf('Nama_Aset') + 1).setValue(payload.namaAset);
          sheet.getRange(r, headers.indexOf('Kategori') + 1).setValue(payload.kategori);
          sheet.getRange(r, headers.indexOf('Jumlah_Total') + 1).setValue(totalBaru);
          sheet.getRange(r, headers.indexOf('Jumlah_Tersedia') + 1).setValue(tersediaBaru);
          sheet.getRange(r, headers.indexOf('Kondisi') + 1).setValue(payload.kondisi);
          sheet.getRange(r, headers.indexOf('Lokasi_Penyimpanan') + 1).setValue(payload.lokasiPenyimpanan);
          sheet.getRange(r, headers.indexOf('Keterangan') + 1).setValue(payload.keterangan);
          SpreadsheetApp.flush();
          return { success: true, message: 'Data aset berhasil diperbarui!' };
        }
      }
    }

    var newId = generateSequentialId_('AST', 'ASET');
    sheet.appendRow([
      newId,
      String(payload.namaAset || '').trim(),
      String(payload.kategori || 'Peralatan Umum').trim(),
      totalBaru,
      totalBaru,
      String(payload.kondisi || 'Baik').trim(),
      String(payload.lokasiPenyimpanan || 'Pos Ronda RT 010').trim(),
      String(payload.keterangan || '').trim(),
      nowStr
    ]);
    SpreadsheetApp.flush();
    return { success: true, message: 'Aset RT baru berhasil ditambahkan!', idAset: newId };
  } catch (error) {
    return { success: false, message: 'Gagal menyimpan data aset: ' + error.toString() };
  }
}

function deleteAset(idAset) {
  try {
    var sheet = getSheet_('ASET');
    var data = sheet.getDataRange().getDisplayValues();
    var idIdx = data[0].indexOf('ID_Aset');
    for (var i = 1; i < data.length; i++) {
      if (data[i][idIdx] === idAset) {
        sheet.deleteRow(i + 1);
        SpreadsheetApp.flush();
        return { success: true, message: 'Aset RT berhasil dihapus!' };
      }
    }
    return { success: false, message: 'ID Aset tidak ditemukan.' };
  } catch (error) {
    return { success: false, message: 'Gagal menghapus aset: ' + error.toString() };
  }
}

function submitPinjamAset(payload) {
  try {
    var asetSheet = getSheet_('ASET');
    var asetData = asetSheet.getDataRange().getDisplayValues();
    var aHeaders = asetData[0];
    var idAsetIdx = aHeaders.indexOf('ID_Aset');
    var tersediaIdx = aHeaders.indexOf('Jumlah_Tersedia');
    var namaAsetIdx = aHeaders.indexOf('Nama_Aset');

    var targetAset = null;
    var jumlahDiminta = parseInt(payload.jumlahPinjam, 10) || 1;

    for (var i = 1; i < asetData.length; i++) {
      if (asetData[i][idAsetIdx] === payload.idAset) {
        targetAset = {
          nama: asetData[i][namaAsetIdx],
          tersedia: parseInt(asetData[i][tersediaIdx], 10) || 0
        };
        break;
      }
    }

    if (!targetAset) {
      return { success: false, message: 'Aset yang dipilih tidak ditemukan di inventaris.' };
    }

    if (targetAset.tersedia < jumlahDiminta) {
      return { 
        success: false, 
        message: 'Stok barang tidak mencukupi! Saat ini hanya tersedia ' + targetAset.tersedia + ' unit.' 
      };
    }

    var pinjamSheet = getSheet_('PEMINJAMAN_ASET');
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var newPinjamId = generateSequentialId_('PJM', 'PEMINJAMAN_ASET');

    pinjamSheet.appendRow([
      newPinjamId,
      payload.idAset,
      targetAset.nama,
      String(payload.noKk || '').trim(),
      String(payload.namaPeminjam || '').trim(),
      String(payload.noHp || '').trim(),
      jumlahDiminta,
      String(payload.tanggalPinjam || '').trim(),
      String(payload.tanggalKembali || '').trim(),
      String(payload.keperluan || '').trim(),
      'Menunggu',
      '-',
      '-',
      nowStr
    ]);

    SpreadsheetApp.flush();
    return { 
      success: true, 
      message: 'Pengajuan peminjaman aset berhasil dikirim! Menunggu persetujuan Pengurus RT.' 
    };
  } catch (error) {
    return { success: false, message: 'Gagal mengajukan peminjaman: ' + error.toString() };
  }
}

function getPeminjamanList(params) {
  try {
    params = params || {};
    var userRole = params.role || 'Warga';
    var userKk = params.noKk || '';
    var filterStatus = params.status || 'SEMUA';

    var sheet = getSheet_('PEMINJAMAN_ASET');
    var rawData = sheet.getDataRange().getDisplayValues();
    if (rawData.length <= 1) return { success: true, data: [] };

    var headers = rawData[0];
    var list = [];

    for (var i = 1; i < rawData.length; i++) {
      var row = rawData[i];
      var kk = row[headers.indexOf('No_KK')];
      var status = (row[headers.indexOf('Status')] || '').trim();

      if (userRole === 'Warga' && userKk && kk !== userKk) {
        continue;
      }

      if (filterStatus !== 'SEMUA' && status !== filterStatus) {
        continue;
      }

      list.push({
        idPinjam: row[headers.indexOf('ID_Pinjam')],
        idAset: row[headers.indexOf('ID_Aset')],
        namaAset: row[headers.indexOf('Nama_Aset')],
        noKk: kk,
        namaPeminjam: row[headers.indexOf('Nama_Peminjam')],
        noHp: row[headers.indexOf('No_HP')],
        jumlahPinjam: parseInt(row[headers.indexOf('Jumlah_Pinjam')], 10) || 0,
        tanggalPinjam: row[headers.indexOf('Tanggal_Pinjam')],
        tanggalKembali: row[headers.indexOf('Tanggal_Kembali')],
        keperluan: row[headers.indexOf('Keperluan')],
        status: status,
        catatanPengurus: row[headers.indexOf('Catatan_Pengurus')],
        disetujuiOleh: row[headers.indexOf('Disetujui_Oleh')],
        createdAt: row[headers.indexOf('Created_At')]
      });
    }

    list.sort(function(a, b) {
      var aWait = (a.status === 'Menunggu');
      var bWait = (b.status === 'Menunggu');
      if (aWait && !bWait) return -1;
      if (!aWait && bWait) return 1;
      return 0;
    });

    return { success: true, data: list };
  } catch (error) {
    return { success: false, message: 'Gagal memuat daftar peminjaman: ' + error.toString() };
  }
}

function approvePeminjamanAset(idPinjam, action, catatan, approverName) {
  try {
    var pSheet = getSheet_('PEMINJAMAN_ASET');
    var pData = pSheet.getDataRange().getDisplayValues();
    var pHeaders = pData[0];

    var idPinjamIdx = pHeaders.indexOf('ID_Pinjam');
    var statusIdx = pHeaders.indexOf('Status');
    var catatanIdx = pHeaders.indexOf('Catatan_Pengurus');
    var approverIdx = pHeaders.indexOf('Disetujui_Oleh');
    var idAsetIdx = pHeaders.indexOf('ID_Aset');
    var qtyIdx = pHeaders.indexOf('Jumlah_Pinjam');

    var targetRowIndex = -1;
    var targetIdAset = '';
    var targetQty = 0;
    var currentStatus = '';

    for (var i = 1; i < pData.length; i++) {
      if (pData[i][idPinjamIdx] === idPinjam) {
        targetRowIndex = i + 1;
        targetIdAset = pData[i][idAsetIdx];
        targetQty = parseInt(pData[i][qtyIdx], 10) || 0;
        currentStatus = (pData[i][statusIdx] || '').trim();
        break;
      }
    }

    if (targetRowIndex === -1) {
      return { success: false, message: 'Data peminjaman tidak ditemukan.' };
    }

    var aSheet = getSheet_('ASET');
    var aData = aSheet.getDataRange().getDisplayValues();
    var aHeaders = aData[0];
    var aIdIdx = aHeaders.indexOf('ID_Aset');
    var aTersediaIdx = aHeaders.indexOf('Jumlah_Tersedia');

    var asetRowIndex = -1;
    var stokTersedia = 0;

    for (var j = 1; j < aData.length; j++) {
      if (aData[j][aIdIdx] === targetIdAset) {
        asetRowIndex = j + 1;
        stokTersedia = parseInt(aData[j][aTersediaIdx], 10) || 0;
        break;
      }
    }

    var newStatus = '';
    var msgSuccess = '';

    if (action === 'APPROVE') {
      if (stokTersedia < targetQty) {
        return { success: false, message: 'Gagal menyetujui: Stok barang saat ini tidak cukup (' + stokTersedia + ' unit).' };
      }
      newStatus = 'Disetujui';
      msgSuccess = 'Peminjaman disetujui! Stok barang otomatis dikurangi.';
      if (asetRowIndex !== -1) {
        aSheet.getRange(asetRowIndex, aTersediaIdx + 1).setValue(Math.max(0, stokTersedia - targetQty));
      }
    } else if (action === 'REJECT') {
      newStatus = 'Ditolak';
      msgSuccess = 'Peminjaman telah ditolak.';
    } else if (action === 'SELESAI') {
      newStatus = 'Selesai';
      msgSuccess = 'Aset berhasil dikonfirmasi kembali! Stok tersedia bertambah.';
      if (currentStatus === 'Disetujui' && asetRowIndex !== -1) {
        aSheet.getRange(asetRowIndex, aTersediaIdx + 1).setValue(stokTersedia + targetQty);
      }
    }

    pSheet.getRange(targetRowIndex, statusIdx + 1).setValue(newStatus);
    pSheet.getRange(targetRowIndex, catatanIdx + 1).setValue(catatan || '-');
    pSheet.getRange(targetRowIndex, approverIdx + 1).setValue(approverName || 'Pengurus RT');

    SpreadsheetApp.flush();
    return { success: true, message: msgSuccess };
  } catch (error) {
    return { success: false, message: 'Gagal memproses aksi peminjaman: ' + error.toString() };
  }
}
