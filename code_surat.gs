/**
 * ====================================================================
 * WARGAKOE - SURAT PENGANTAR RT BACKEND MODULE (E-SURAT MANDIRI)
 * Mengelola permohonan surat pengantar warga, penomoran otomatis resmi,
 * persetujuan pengurus RT, dan validasi dokumen digital RT 010.
 * ====================================================================
 */

function getSuratSheetSafe_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('SURAT_PENGANTAR');
  if (!sheet) {
    sheet = ss.insertSheet('SURAT_PENGANTAR');
    var headers = [
      'ID_Surat',
      'No_Surat',
      'No_KK',
      'Nama_Pemohon',
      'NIK',
      'No_HP',
      'Jenis_Surat',
      'Keperluan',
      'Keterangan_Tambahan',
      'Status',
      'Catatan_Pengurus',
      'Disetujui_Oleh',
      'Tanggal_Pengajuan',
      'Tanggal_Disetujui',
      'Created_At'
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

function getSuratList(params) {
  try {
    params = params || {};
    var search = String(params.search || '').toLowerCase().trim();
    var filterStatus = String(params.status || 'SEMUA').trim();
    var filterJenis = String(params.jenis || 'SEMUA').trim();
    var onlyMine = Boolean(params.onlyMine);
    var userNoKk = String(params.userNoKk || '').trim();
    var userRole = String(params.userRole || 'Warga').trim();
    var userStatusUser = String(params.userStatusUser || '').trim();

    // Verifikasi apakah pemanggil adalah Pengurus RT atau Super Admin
    var isPengurus = (
      userRole === 'Super Admin' ||
      userRole === 'Ketua RT' ||
      userRole === 'Pengurus RT' ||
      userRole === 'Bendahara' ||
      userRole === 'Sekretaris' ||
      userStatusUser === 'Pengurus RT'
    );

    var sheet = getSuratSheetSafe_();
    var rawData = sheet.getDataRange().getDisplayValues();
    if (rawData.length <= 1) {
      return { success: true, data: [], pendingCount: 0 };
    }

    var headers = rawData[0];
    var idIdx = headers.indexOf('ID_Surat');
    var noSuratIdx = headers.indexOf('No_Surat');
    var kkIdx = headers.indexOf('No_KK');
    var namaIdx = headers.indexOf('Nama_Pemohon');
    var nikIdx = headers.indexOf('NIK');
    var hpIdx = headers.indexOf('No_HP');
    var jenisIdx = headers.indexOf('Jenis_Surat');
    var kepIdx = headers.indexOf('Keperluan');
    var ketIdx = headers.indexOf('Keterangan_Tambahan');
    var statusIdx = headers.indexOf('Status');
    var catIdx = headers.indexOf('Catatan_Pengurus');
    var ttdIdx = headers.indexOf('Disetujui_Oleh');
    var tglPengajuanIdx = headers.indexOf('Tanggal_Pengajuan');
    var tglDisetujuiIdx = headers.indexOf('Tanggal_Disetujui');
    var crtIdx = headers.indexOf('Created_At');

    var list = [];
    var pendingCount = 0;

    for (var i = rawData.length - 1; i >= 1; i--) {
      var row = rawData[i];
      var rowStatus = String(row[statusIdx] || 'Menunggu').trim();
      var rowKk = String(row[kkIdx] || '').trim();
      var rowNama = String(row[namaIdx] || '').trim();
      var rowNoSurat = String(row[noSuratIdx] || '-').trim();
      var rowJenis = String(row[jenisIdx] || 'Surat Pengantar Umum').trim();
      var rowKep = String(row[kepIdx] || '').trim();

      // Hitung pending approval khusus untuk pengurus
      if (rowStatus.toLowerCase() === 'menunggu') {
        if (isPengurus || (userNoKk && rowKk === userNoKk)) {
          pendingCount++;
        }
      }

      // PRIVASI KETAT: Jika role Warga, HANYA boleh melihat surat miliknya sendiri (berdasarkan No. KK)
      if (!isPengurus) {
        if (!userNoKk || rowKk !== userNoKk) {
          continue;
        }
      } else if (onlyMine && userNoKk && rowKk !== userNoKk) {
        continue;
      }

      if (filterStatus !== 'SEMUA' && rowStatus.toLowerCase() !== filterStatus.toLowerCase()) {
        continue;
      }

      if (filterJenis !== 'SEMUA' && rowJenis.toLowerCase() !== filterJenis.toLowerCase()) {
        continue;
      }

      var matchSearch = !search || (
        rowNama.toLowerCase().indexOf(search) !== -1 ||
        rowNoSurat.toLowerCase().indexOf(search) !== -1 ||
        rowJenis.toLowerCase().indexOf(search) !== -1 ||
        rowKep.toLowerCase().indexOf(search) !== -1 ||
        rowKk.indexOf(search) !== -1
      );

      if (matchSearch) {
        list.push({
          idSurat: row[idIdx],
          noSurat: rowNoSurat,
          noKk: rowKk,
          namaPemohon: rowNama,
          nik: row[nikIdx] || '-',
          noHp: row[hpIdx] || '-',
          jenisSurat: rowJenis,
          keperluan: rowKep,
          keteranganTambahan: row[ketIdx] || '-',
          status: rowStatus,
          catatanPengurus: row[catIdx] || '-',
          disetujuiOleh: row[ttdIdx] || '-',
          tanggalPengajuan: row[tglPengajuanIdx] || '-',
          tanggalDisetujui: row[tglDisetujuiIdx] || '-',
          createdAt: row[crtIdx] || '-'
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
    return { success: false, message: 'Gagal memuat daftar surat pengantar: ' + error.toString() };
  }
}

function savePengajuanSurat(payload) {
  try {
    var sheet = getSuratSheetSafe_();
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var todayStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy');
    var newId = generateSequentialSuratId_();

    var jenisSurat = String(payload.jenisSurat || 'Surat Keterangan Domisili').trim();
    var keperluan = String(payload.keperluan || '').trim();
    var nama = String(payload.namaPemohon || 'Warga RT 010').trim();

    if (!keperluan) {
      return { success: false, message: 'Keperluan surat wajib diisi dengan jelas!' };
    }

    sheet.appendRow([
      newId,
      '-', // Nomor surat resmi diterbitkan saat disetujui Pengurus RT
      String(payload.noKk || '').trim(),
      nama,
      String(payload.nik || '-').trim(),
      String(payload.noHp || '-').trim(),
      jenisSurat,
      keperluan,
      String(payload.keteranganTambahan || '-').trim(),
      'Menunggu',
      '-',
      '-',
      todayStr,
      '-',
      nowStr
    ]);

    SpreadsheetApp.flush();
    return {
      success: true,
      message: 'Permohonan surat pengantar berhasil dikirim! Pengurus RT akan meninjau pengajuan Anda.',
      idSurat: newId
    };
  } catch (error) {
    return { success: false, message: 'Gagal mengirim pengajuan surat: ' + error.toString() };
  }
}

function updateStatusSurat(payload) {
  try {
    var sheet = getSuratSheetSafe_();
    var data = sheet.getDataRange().getDisplayValues();
    var headers = data[0];

    var idIdx = headers.indexOf('ID_Surat');
    var noSuratIdx = headers.indexOf('No_Surat');
    var statusIdx = headers.indexOf('Status');
    var catIdx = headers.indexOf('Catatan_Pengurus');
    var ttdIdx = headers.indexOf('Disetujui_Oleh');
    var tglDisetujuiIdx = headers.indexOf('Tanggal_Disetujui');

    var targetId = String(payload.idSurat || '').trim();
    var targetRow = -1;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idIdx]).trim() === targetId) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow === -1) {
      return { success: false, message: 'Data permohonan surat tidak ditemukan.' };
    }

    var newStatus = String(payload.status || 'Disetujui').trim();
    var catatan = String(payload.catatanPengurus || '-').trim();
    var petugas = String(payload.petugas || 'Ketua RT 010').trim();
    var todayStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy');

    var existingNoSurat = String(data[targetRow - 1][noSuratIdx]).trim();
    var finalNoSurat = existingNoSurat;

    if (newStatus === 'Disetujui' && (existingNoSurat === '-' || existingNoSurat === '')) {
      finalNoSurat = generateNomorSuratResmi_(sheet);
    }

    sheet.getRange(targetRow, noSuratIdx + 1).setValue(finalNoSurat);
    sheet.getRange(targetRow, statusIdx + 1).setValue(newStatus);
    sheet.getRange(targetRow, catIdx + 1).setValue(catatan);
    sheet.getRange(targetRow, ttdIdx + 1).setValue(newStatus === 'Disetujui' ? petugas : '-');
    sheet.getRange(targetRow, tglDisetujuiIdx + 1).setValue(newStatus === 'Disetujui' ? todayStr : '-');

    SpreadsheetApp.flush();
    return {
      success: true,
      message: 'Status surat berhasil diperbarui menjadi: ' + newStatus + (newStatus === 'Disetujui' ? ' (No: ' + finalNoSurat + ')' : ''),
      noSurat: finalNoSurat
    };
  } catch (error) {
    return { success: false, message: 'Gagal memperbarui status surat: ' + error.toString() };
  }
}

function deleteSurat(idSurat) {
  try {
    var sheet = getSuratSheetSafe_();
    var data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) return { success: false, message: 'Data tidak ditemukan.' };

    var idIdx = data[0].indexOf('ID_Surat');
    var targetId = String(idSurat || '').trim();
    var targetRow = -1;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idIdx]).trim() === targetId) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow === -1) {
      return { success: false, message: 'Data surat tidak ditemukan atau sudah dihapus.' };
    }

    sheet.deleteRow(targetRow);
    SpreadsheetApp.flush();
    return { success: true, message: 'Data surat pengantar ' + targetId + ' berhasil dihapus.' };
  } catch (err) {
    return { success: false, message: 'Gagal menghapus surat: ' + err.toString() };
  }
}

function getSuratPendingCount() {
  try {
    var sheet = getSuratSheetSafe_();
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

function generateSequentialSuratId_() {
  try {
    var sheet = getSuratSheetSafe_();
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return 'SP-0001';

    var data = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
    var maxNum = 0;
    for (var i = 0; i < data.length; i++) {
      var id = String(data[i][0]);
      if (id.indexOf('SP-') === 0) {
        var num = parseInt(id.replace('SP-', ''), 10);
        if (!isNaN(num) && num > maxNum) maxNum = num;
      }
    }
    var nextNum = maxNum + 1;
    return 'SP-' + ('0000' + nextNum).slice(-4);
  } catch (e) {
    return 'SP-' + Math.floor(1000 + Math.random() * 9000);
  }
}

function generateNomorSuratResmi_(sheet) {
  try {
    var d = new Date();
    var romawiBulan = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII'];
    var bulanRomawi = romawiBulan[d.getMonth()];
    var tahun = d.getFullYear();

    var countApproved = 0;
    if (sheet.getLastRow() > 1) {
      var data = sheet.getDataRange().getDisplayValues();
      var noSuratIdx = data[0].indexOf('No_Surat');
      if (noSuratIdx !== -1) {
        for (var i = 1; i < data.length; i++) {
          var val = String(data[i][noSuratIdx]);
          if (val && val !== '-' && val.indexOf('/SP-RT010/') !== -1) {
            countApproved++;
          }
        }
      }
    }

    var nextSeq = countApproved + 1;
    var seqPadded = ('000' + nextSeq).slice(-3);
    return seqPadded + '/SP-RT010/RW05/' + bulanRomawi + '/' + tahun;
  } catch (e) {
    return '001/SP-RT010/RW05/IX/2026';
  }
}
