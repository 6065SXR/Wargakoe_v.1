/**
 * ====================================================================
 * WARGAKOE - IURAN & KAS RT BACKEND MODULE (CASHFLOW & 3-IN-1 ENGINE)
 * Mengelola pembukuan Kas RT, verifikasi pembayaran iuran warga,
 * Dompet 4 Pos Terpisah, Monitoring Kepatuhan KK, & Buku Belanja RT.
 * ====================================================================
 */

function getKasRtDetails(targetBulan, targetTahun) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var curBulan = String(targetBulan || 'September').trim();
    var curTahun = String(targetTahun || '2026').trim();

    // 1. DOMPET 4 POS TERPISAH & BUKU BELANJA (DARI SHEET KAS_RT)
    var dompet = {
      kasRt: { masuk: 0, keluar: 0, saldo: 0 },
      kasDuka: { masuk: 0, keluar: 0, saldo: 0 },
      kasSampah: { masuk: 0, keluar: 0, saldo: 0 },
      kasSosial: { masuk: 0, keluar: 0, saldo: 0 }
    };

    var pengeluaranList = [];
    var totalPemasukanAll = 0;
    var totalPengeluaranAll = 0;

    var kasSheet = getKasSheetSafe_(ss);
    if (kasSheet && kasSheet.getLastRow() > 1) {
      var rawData = kasSheet.getDataRange().getDisplayValues();
      var headers = rawData[0];

      var idIdx = findColIdxFlexible_(headers, ['id_kas', 'idkas', 'id'], 0);
      var tglIdx = findColIdxFlexible_(headers, ['tanggal', 'tgl', 'date'], 1);
      var jenisIdx = findColIdxFlexible_(headers, ['jenis_kas', 'jeniskas', 'jenis', 'tipe'], 2);
      var katIdx = findColIdxFlexible_(headers, ['kategori', 'category', 'pos'], 3);
      var ketIdx = findColIdxFlexible_(headers, ['keterangan', 'ket', 'deskripsi', 'rincian'], 4);
      var nomIdx = findColIdxFlexible_(headers, ['nominal', 'jumlah', 'total', 'rp'], 5);
      var bulanTahunIdx = findColIdxFlexible_(headers, ['bulan_tahun', 'bulantahun', 'periode', 'bulan'], 6);
      var createdIdx = findColIdxFlexible_(headers, ['created_at', 'createdat', 'timestamp'], 7);

      for (var i = rawData.length - 1; i >= 1; i--) {
        var row = rawData[i];
        if (!row || row.length === 0) continue;

        var colId = String(idIdx !== -1 && row[idIdx] ? row[idIdx] : ('KAS-' + i)).trim();
        var colTgl = String(tglIdx !== -1 && row[tglIdx] ? row[tglIdx] : '-').trim();
        var colJenis = String(jenisIdx !== -1 && row[jenisIdx] ? row[jenisIdx] : '').trim().toLowerCase();
        var colKat = String(katIdx !== -1 && row[katIdx] ? row[katIdx] : '').trim();
        var colKet = String(ketIdx !== -1 && row[ketIdx] ? row[ketIdx] : '-').trim();
        var colNom = parseKasNominalSafe_(nomIdx !== -1 ? row[nomIdx] : 0);
        var colBulan = String(bulanTahunIdx !== -1 && row[bulanTahunIdx] ? row[bulanTahunIdx] : '').trim();
        var colCreated = String(createdIdx !== -1 && row[createdIdx] ? row[createdIdx] : '').trim();

        var isMasuk = (colJenis.indexOf('masuk') !== -1 || colJenis.indexOf('terima') !== -1);
        var isKeluar = (colJenis.indexOf('keluar') !== -1 || colJenis.indexOf('belanja') !== -1 || colJenis.indexOf('biaya') !== -1);

        var pocketKey = 'kasRt';
        var katLower = colKat.toLowerCase();

        if (katLower.indexOf('duka') !== -1 || katLower.indexOf('kematian') !== -1) {
          pocketKey = 'kasDuka';
        } else if (katLower.indexOf('sampah') !== -1) {
          pocketKey = 'kasSampah';
        } else if (katLower.indexOf('sosial') !== -1) {
          pocketKey = 'kasSosial';
        } else {
          pocketKey = 'kasRt';
        }

        if (isMasuk) {
          dompet[pocketKey].masuk += colNom;
          totalPemasukanAll += colNom;
        } else if (isKeluar) {
          dompet[pocketKey].keluar += colNom;
          totalPengeluaranAll += colNom;

          pengeluaranList.push({
            idKas: colId,
            tanggal: colTgl !== '-' ? colTgl : (colCreated ? colCreated.split(' ')[0] : '-'),
            kategori: colKat || 'Operasional RT',
            keterangan: colKet,
            nominal: colNom,
            bulanTahun: colBulan
          });
        }
      }

      dompet.kasRt.saldo = dompet.kasRt.masuk - dompet.kasRt.keluar;
      dompet.kasDuka.saldo = dompet.kasDuka.masuk - dompet.kasDuka.keluar;
      dompet.kasSampah.saldo = dompet.kasSampah.masuk - dompet.kasSampah.keluar;
      dompet.kasSosial.saldo = dompet.kasSosial.masuk - dompet.kasSosial.keluar;
    }

    // 2. MONITORING KEPATUHAN KEPALA KELUARGA (USERS + IURAN)
    var totalKk = 0;
    var lunasKk = 0;
    var kkMap = {};

    var userSheet = ss.getSheetByName('USERS') || ss.getSheetByName('Users') || ss.getSheetByName('users');
    if (userSheet && userSheet.getLastRow() > 1) {
      var uData = userSheet.getDataRange().getDisplayValues();
      var uHeaders = uData[0];
      var uKkIdx = findColIdxFlexible_(uHeaders, ['no_kk', 'nokk', 'no kk', 'kk'], 0);
      var uNamaIdx = findColIdxFlexible_(uHeaders, ['nama', 'nama_lengkap', 'namawarga'], 2);
      var uHpIdx = findColIdxFlexible_(uHeaders, ['no_hp', 'nohp', 'hp', 'telepon'], 3);
      var uAlamatIdx = findColIdxFlexible_(uHeaders, ['alamat', 'rumah', 'lokasi'], 9);

      for (var u = 1; u < uData.length; u++) {
        var noKk = String((uKkIdx !== -1 ? uData[u][uKkIdx] : '') || ('KK-' + u)).trim();
        if (noKk && !kkMap[noKk]) {
          totalKk++;
          kkMap[noKk] = {
            noKk: noKk,
            nama: String((uNamaIdx !== -1 ? uData[u][uNamaIdx] : '') || 'Warga').trim(),
            noHp: String((uHpIdx !== -1 ? uData[u][uHpIdx] : '') || '-').trim(),
            alamat: String((uAlamatIdx !== -1 ? uData[u][uAlamatIdx] : '') || 'RT 010 RW 05').trim(),
            isLunas: false
          };
        }
      }
    }

    var iurSheet = ss.getSheetByName('IURAN') || ss.getSheetByName('Iuran') || ss.getSheetByName('iuran');
    if (iurSheet && iurSheet.getLastRow() > 1) {
      var iData = iurSheet.getDataRange().getDisplayValues();
      var iHeaders = iData[0];
      var iKkIdx = findColIdxFlexible_(iHeaders, ['no_kk', 'nokk', 'no kk', 'kk'], 2);
      var iBulanIdx = findColIdxFlexible_(iHeaders, ['bulan_tahun', 'bulantahun', 'periode', 'bulan'], 4);
      var iStatusIdx = findColIdxFlexible_(iHeaders, ['status_bayar', 'statusbayar', 'status'], 7);

      var curBulanClean = curBulan.toLowerCase();
      var curTahunClean = curTahun.toLowerCase();

      for (var j = 1; j < iData.length; j++) {
        var rKk = String(iKkIdx !== -1 ? iData[j][iKkIdx] : '').trim();
        var rBulan = String(iBulanIdx !== -1 ? iData[j][iBulanIdx] : '').toLowerCase().trim();
        var rStatus = String(iStatusIdx !== -1 ? iData[j][iStatusIdx] : '').toLowerCase().trim();

        var isMatchMonth = (rBulan.indexOf(curBulanClean) !== -1 || rBulan.indexOf('09') !== -1 || rBulan.indexOf('sep') !== -1);
        var isPaid = (rStatus === 'sudah bayar' || rStatus === 'lunas' || rStatus === 'approved' || rStatus === 'sukses' || rStatus === 'sudah');

        if (isMatchMonth && isPaid && kkMap[rKk] && !kkMap[rKk].isLunas) {
          kkMap[rKk].isLunas = true;
          lunasKk++;
        }
      }
    }

    var belumLunasList = [];
    for (var k in kkMap) {
      if (kkMap.hasOwnProperty(k) && !kkMap[k].isLunas) {
        belumLunasList.push(kkMap[k]);
      }
    }

    var persentaseKepatuhan = totalKk > 0 ? Math.round((lunasKk / totalKk) * 100) : 0;

    var kepatuhan = {
      periode: curBulan + ' ' + curTahun,
      totalKk: totalKk,
      lunasKk: lunasKk,
      belumLunasCount: belumLunasList.length,
      persentase: persentaseKepatuhan,
      belumLunasList: belumLunasList
    };

    var payload = {
      dompet: dompet,
      kepatuhan: kepatuhan,
      pengeluaranList: pengeluaranList,
      totalPemasukan: dompet.kasRt.masuk,
      totalPengeluaran: dompet.kasRt.keluar,
      sisaSaldo: dompet.kasRt.saldo,
      mutasi: pengeluaranList
    };

    return {
      success: true,
      data: payload
    };
  } catch (err) {
    return {
      success: false,
      message: 'Gagal memuat rincian kas: ' + err.toString()
    };
  }
}

function savePengeluaranRutin(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var kasSheet = getKasSheetSafe_(ss);
    if (!kasSheet) throw new Error('Sheet KAS_RT tidak ditemukan!');

    var rawHeaders = kasSheet.getRange(1, 1, 1, kasSheet.getLastColumn()).getDisplayValues()[0];
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var newId = generateIuranSequentialId_('KAS', 'KAS_RT');

    var idCol = findColIdxFlexible_(rawHeaders, ['id_kas', 'idkas', 'id'], 0);
    var tglCol = findColIdxFlexible_(rawHeaders, ['tanggal', 'tgl', 'date'], 1);
    var jenisCol = findColIdxFlexible_(rawHeaders, ['jenis_kas', 'jeniskas', 'jenis', 'tipe'], 2);
    var katCol = findColIdxFlexible_(rawHeaders, ['kategori', 'category', 'pos'], 3);
    var ketCol = findColIdxFlexible_(rawHeaders, ['keterangan', 'ket', 'deskripsi', 'rincian'], 4);
    var nomCol = findColIdxFlexible_(rawHeaders, ['nominal', 'jumlah', 'total', 'rp'], 5);
    var blnCol = findColIdxFlexible_(rawHeaders, ['bulan_tahun', 'bulantahun', 'periode', 'bulan'], 6);
    var crtCol = findColIdxFlexible_(rawHeaders, ['created_at', 'createdat', 'timestamp'], 7);

    var maxLen = Math.max(rawHeaders.length, 8);
    var newRow = new Array(maxLen);
    for (var c = 0; c < maxLen; c++) newRow[c] = '';

    if (idCol !== -1) newRow[idCol] = newId;
    if (tglCol !== -1) newRow[tglCol] = String(payload.tanggal || Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy')).trim();
    if (jenisCol !== -1) newRow[jenisCol] = 'Pengeluaran';
    if (katCol !== -1) newRow[katCol] = String(payload.kategori || 'Iuran Kas').trim();
    if (ketCol !== -1) newRow[ketCol] = String(payload.keterangan || '').trim();
    if (nomCol !== -1) newRow[nomCol] = String(payload.nominal || 0);
    if (blnCol !== -1) newRow[blnCol] = String(payload.bulanTahun || '').trim();
    if (crtCol !== -1) newRow[crtCol] = nowStr;

    kasSheet.appendRow(newRow);
    SpreadsheetApp.flush();

    return {
      success: true,
      message: 'Pengeluaran rutin RT sebesar Rp ' + parseInt(payload.nominal, 10).toLocaleString('id-ID') + ' berhasil dicatat!'
    };
  } catch (err) {
    return { success: false, message: 'Gagal mencatat pengeluaran: ' + err.toString() };
  }
}

function getIuranList(params) {
  try {
    params = params || {};
    var limit = parseInt(params.limit, 10) || 10;
    var offset = parseInt(params.offset, 10) || 0;
    var search = (params.search || '').toLowerCase().trim();
    var filterStatus = (params.status || 'SEMUA').toLowerCase().trim();
    var filterJenis = (params.jenis || 'SEMUA').toLowerCase().trim();

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var iurSheet = ss.getSheetByName('IURAN') || ss.getSheetByName('Iuran') || ss.getSheetByName('iuran');

    if (!iurSheet || iurSheet.getLastRow() <= 1) {
      return { success: true, data: [], total: 0, page: 1, totalPages: 0 };
    }

    var rawData = iurSheet.getDataRange().getDisplayValues();
    var headers = rawData[0];

    var idIdx = findColIdxFlexible_(headers, ['id_iuran', 'idiuran', 'id'], 0);
    var kwiIdx = findColIdxFlexible_(headers, ['no_kuitansi', 'kuitansi', 'kwitansi'], 1);
    var kkIdx = findColIdxFlexible_(headers, ['no_kk', 'nokk', 'no kk', 'kk'], 2);
    var namaIdx = findColIdxFlexible_(headers, ['nama_warga', 'namawarga', 'nama'], 3);
    var bulanIdx = findColIdxFlexible_(headers, ['bulan_tahun', 'bulantahun', 'periode', 'bulan'], 4);
    var jenisIdx = findColIdxFlexible_(headers, ['jenis_iuran', 'jenisiuran', 'jenis'], 5);
    var nomIdx = findColIdxFlexible_(headers, ['nominal', 'jumlah', 'total', 'rp'], 6);
    var statusIdx = findColIdxFlexible_(headers, ['status_bayar', 'statusbayar', 'status'], 7);
    var tglIdx = findColIdxFlexible_(headers, ['tanggal_bayar', 'tanggal', 'tgl'], 8);
    var appIdx = findColIdxFlexible_(headers, ['approved_by', 'approvedby', 'petugas', 'verifikator'], 9);

    var filtered = [];

    for (var i = rawData.length - 1; i >= 1; i--) {
      var row = rawData[i];
      var idIuran = idIdx !== -1 ? row[idIdx] : '';
      var noKwi = kwiIdx !== -1 ? row[kwiIdx] : '';
      var noKk = kkIdx !== -1 ? row[kkIdx] : '';
      var namaWarga = namaIdx !== -1 ? row[namaIdx] : '';
      var bulanTahun = bulanIdx !== -1 ? row[bulanIdx] : '';
      var jenisIuran = jenisIdx !== -1 ? row[jenisIdx] : '';
      var nominal = parseKasNominalSafe_(nomIdx !== -1 ? row[nomIdx] : 0);
      var statusBayar = statusIdx !== -1 ? row[statusIdx] : 'Menunggu';
      var tglBayar = tglIdx !== -1 ? row[tglIdx] : '-';
      var approvedBy = appIdx !== -1 ? row[appIdx] : '-';

      var matchSearch = !search || 
        (namaWarga.toLowerCase().indexOf(search) !== -1 || 
         noKwi.toLowerCase().indexOf(search) !== -1 || 
         noKk.toLowerCase().indexOf(search) !== -1);

      var matchStatus = filterStatus === 'semua' || statusBayar.toLowerCase() === filterStatus;
      var matchJenis = filterJenis === 'semua' || jenisIuran.toLowerCase() === filterJenis;

      if (matchSearch && matchStatus && matchJenis) {
        filtered.push({
          idIuran: idIuran,
          noKuitansi: noKwi,
          noKk: noKk,
          namaWarga: namaWarga,
          bulanTahun: bulanTahun,
          jenisIuran: jenisIuran,
          nominal: nominal,
          statusBayar: statusBayar,
          tanggalBayar: tglBayar,
          approvedBy: approvedBy
        });
      }
    }

    // PRIORITAS UTAMA: Status 'Menunggu' selalu diletakkan paling atas halaman 1
    filtered.sort(function(a, b) {
      var aIsMenunggu = String(a.statusBayar || '').toLowerCase().trim() === 'menunggu';
      var bIsMenunggu = String(b.statusBayar || '').toLowerCase().trim() === 'menunggu';
      if (aIsMenunggu && !bIsMenunggu) return -1;
      if (!aIsMenunggu && bIsMenunggu) return 1;
      return 0;
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
  } catch (error) {
    return { success: false, message: 'Gagal memuat daftar iuran: ' + error.toString() };
  }
}

function saveBayarIuran(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var iurSheet = ss.getSheetByName('IURAN') || ss.getSheetByName('Iuran') || ss.getSheetByName('iuran');
    if (!iurSheet) throw new Error('Sheet IURAN tidak ditemukan!');

    var nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var rawHeaders = iurSheet.getRange(1, 1, 1, iurSheet.getLastColumn()).getDisplayValues()[0];

    var newId = generateIuranSequentialId_('IUR', 'IURAN');
    var noKwi = String(payload.noKuitansi || '').trim();
    if (!noKwi) {
      noKwi = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd') + Math.floor(1000 + Math.random() * 9000);
    }

    var isSuper = payload.isPengurus || false;
    var statusAwal = isSuper ? 'Sudah bayar' : 'Menunggu';
    var tglBayar = payload.tanggalBayar || Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy');
    var approvedBy = isSuper ? (payload.petugas || 'Pengurus RT') : '-';

    var idCol = findColIdxFlexible_(rawHeaders, ['id_iuran', 'idiuran', 'id'], 0);
    var kwiCol = findColIdxFlexible_(rawHeaders, ['no_kuitansi', 'kuitansi', 'kwitansi'], 1);
    var kkCol = findColIdxFlexible_(rawHeaders, ['no_kk', 'nokk', 'no kk', 'kk'], 2);
    var namaCol = findColIdxFlexible_(rawHeaders, ['nama_warga', 'namawarga', 'nama'], 3);
    var blnCol = findColIdxFlexible_(rawHeaders, ['bulan_tahun', 'bulantahun', 'periode', 'bulan'], 4);
    var jnsCol = findColIdxFlexible_(rawHeaders, ['jenis_iuran', 'jenisiuran', 'jenis'], 5);
    var nomCol = findColIdxFlexible_(rawHeaders, ['nominal', 'jumlah', 'total', 'rp'], 6);
    var stCol = findColIdxFlexible_(rawHeaders, ['status_bayar', 'statusbayar', 'status'], 7);
    var tglCol = findColIdxFlexible_(rawHeaders, ['tanggal_bayar', 'tanggal', 'tgl'], 8);
    var appCol = findColIdxFlexible_(rawHeaders, ['approved_by', 'approvedby', 'petugas'], 9);
    var crtCol = findColIdxFlexible_(rawHeaders, ['created_at', 'createdat', 'timestamp'], 10);

    var maxLen = Math.max(rawHeaders.length, 11);
    var newRow = new Array(maxLen);
    for (var c = 0; c < maxLen; c++) newRow[c] = '';

    if (idCol !== -1) newRow[idCol] = newId;
    if (kwiCol !== -1) newRow[kwiCol] = noKwi;
    if (kkCol !== -1) newRow[kkCol] = String(payload.noKk || '').trim();
    if (namaCol !== -1) newRow[namaCol] = String(payload.namaWarga || '').trim();
    if (blnCol !== -1) newRow[blnCol] = String(payload.bulanTahun || '').trim();
    if (jnsCol !== -1) newRow[jnsCol] = String(payload.jenisIuran || '').trim();
    if (nomCol !== -1) newRow[nomCol] = String(payload.nominal || 0);
    if (stCol !== -1) newRow[stCol] = statusAwal;
    if (tglCol !== -1) newRow[tglCol] = tglBayar;
    if (appCol !== -1) newRow[appCol] = approvedBy;
    if (crtCol !== -1) newRow[crtCol] = nowStr;

    iurSheet.appendRow(newRow);

    if (statusAwal === 'Sudah bayar') {
      logIuranToKasRt_(ss, {
        tanggal: tglBayar,
        kategori: payload.jenisIuran,
        keterangan: 'Iuran dari ' + payload.namaWarga + ' (' + noKwi + ')',
        nominal: payload.nominal,
        bulanTahun: payload.bulanTahun
      });
    }

    SpreadsheetApp.flush();
    return {
      success: true,
      message: 'Pembayaran iuran berhasil dikirim!' + (isSuper ? ' (Status: Lunas)' : ' (Menunggu persetujuan bendahara)')
    };
  } catch (err) {
    return { success: false, message: 'Gagal menyimpan pembayaran: ' + err.toString() };
  }
}

function updateStatusIuran(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var iurSheet = ss.getSheetByName('IURAN') || ss.getSheetByName('Iuran') || ss.getSheetByName('iuran');
    if (!iurSheet) throw new Error('Sheet IURAN tidak ditemukan!');

    var data = iurSheet.getDataRange().getDisplayValues();
    var headers = data[0];

    var idIdx = findColIdxFlexible_(headers, ['id_iuran', 'idiuran', 'id'], 0);
    var statusIdx = findColIdxFlexible_(headers, ['status_bayar', 'statusbayar', 'status'], 7);
    var appIdx = findColIdxFlexible_(headers, ['approved_by', 'approvedby', 'petugas'], 9);
    var tglIdx = findColIdxFlexible_(headers, ['tanggal_bayar', 'tanggal', 'tgl'], 8);
    var namaIdx = findColIdxFlexible_(headers, ['nama_warga', 'namawarga', 'nama'], 3);
    var kwiIdx = findColIdxFlexible_(headers, ['no_kuitansi', 'kuitansi', 'kwitansi'], 1);
    var jenisIdx = findColIdxFlexible_(headers, ['jenis_iuran', 'jenisiuran', 'jenis'], 5);
    var nomIdx = findColIdxFlexible_(headers, ['nominal', 'jumlah', 'total', 'rp'], 6);
    var bulanIdx = findColIdxFlexible_(headers, ['bulan_tahun', 'bulantahun', 'periode', 'bulan'], 4);

    var targetId = String(payload.idIuran || '').trim();
    var newStatus = String(payload.status || 'Sudah bayar').trim();
    var verifikator = String(payload.verifikator || 'Pengurus RT').trim();
    var todayStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy');

    var foundRow = -1;
    var rowData = null;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idIdx]).trim() === targetId) {
        foundRow = i + 1;
        rowData = data[i];
        break;
      }
    }

    if (foundRow === -1) {
      return { success: false, message: 'Data iuran tidak ditemukan.' };
    }

    iurSheet.getRange(foundRow, statusIdx + 1).setValue(newStatus);
    iurSheet.getRange(foundRow, appIdx + 1).setValue(newStatus === 'Sudah bayar' ? verifikator : '-');
    if (newStatus === 'Sudah bayar') {
      iurSheet.getRange(foundRow, tglIdx + 1).setValue(todayStr);
    }

    if (newStatus === 'Sudah bayar' && rowData) {
      logIuranToKasRt_(ss, {
        tanggal: todayStr,
        kategori: rowData[jenisIdx],
        keterangan: 'Iuran dari ' + rowData[namaIdx] + ' (' + rowData[kwiIdx] + ')',
        nominal: rowData[nomIdx],
        bulanTahun: rowData[bulanIdx]
      });
    }

    SpreadsheetApp.flush();
    return {
      success: true,
      message: 'Status iuran berhasil diperbarui menjadi: ' + newStatus
    };
  } catch (err) {
    return { success: false, message: 'Gagal update status iuran: ' + err.toString() };
  }
}

function getAllIuranForExport(filterJenis) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var iurSheet = ss.getSheetByName('IURAN') || ss.getSheetByName('Iuran') || ss.getSheetByName('iuran');
    if (!iurSheet || iurSheet.getLastRow() <= 1) return { success: true, data: [] };

    var data = iurSheet.getDataRange().getDisplayValues();
    var headers = data[0];

    var kwiIdx = findColIdxFlexible_(headers, ['no_kuitansi', 'kuitansi', 'kwitansi'], 1);
    var namaIdx = findColIdxFlexible_(headers, ['nama_warga', 'namawarga', 'nama'], 3);
    var bulanIdx = findColIdxFlexible_(headers, ['bulan_tahun', 'bulantahun', 'periode', 'bulan'], 4);
    var jenisIdx = findColIdxFlexible_(headers, ['jenis_iuran', 'jenisiuran', 'jenis'], 5);
    var nomIdx = findColIdxFlexible_(headers, ['nominal', 'jumlah', 'total', 'rp'], 6);
    var statusIdx = findColIdxFlexible_(headers, ['status_bayar', 'statusbayar', 'status'], 7);
    var tglIdx = findColIdxFlexible_(headers, ['tanggal_bayar', 'tanggal', 'tgl'], 8);

    var filter = String(filterJenis || 'SEMUA').toLowerCase().trim();
    var list = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var status = String(row[statusIdx]).toLowerCase().trim();
      var jenis = String(row[jenisIdx]).trim();

      if (status === 'sudah bayar' || status === 'lunas' || status === 'approved') {
        if (filter === 'semua' || jenis.toLowerCase() === filter) {
          list.push({
            noKuitansi: row[kwiIdx] || '-',
            namaWarga: row[namaIdx] || '-',
            jenisIuran: jenis,
            bulanTahun: row[bulanIdx] || '-',
            nominal: parseKasNominalSafe_(row[nomIdx]),
            tanggalBayar: row[tglIdx] || '-'
          });
        }
      }
    }

    return { success: true, data: list };
  } catch (err) {
    return { success: false, message: 'Gagal mengambil data cetak: ' + err.toString() };
  }
}

function logIuranToKasRt_(ss, item) {
  try {
    var kasSheet = getKasSheetSafe_(ss);
    if (!kasSheet) return;

    var newId = generateIuranSequentialId_('KAS', 'KAS_RT');
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var headers = kasSheet.getRange(1, 1, 1, kasSheet.getLastColumn()).getDisplayValues()[0];

    var idCol = findColIdxFlexible_(headers, ['id_kas', 'idkas', 'id'], 0);
    var tglCol = findColIdxFlexible_(headers, ['tanggal', 'tgl', 'date'], 1);
    var jenisCol = findColIdxFlexible_(headers, ['jenis_kas', 'jeniskas', 'jenis', 'tipe'], 2);
    var katCol = findColIdxFlexible_(headers, ['kategori', 'category', 'pos'], 3);
    var ketCol = findColIdxFlexible_(headers, ['keterangan', 'ket', 'deskripsi'], 4);
    var nomCol = findColIdxFlexible_(headers, ['nominal', 'jumlah', 'total', 'rp'], 5);
    var blnCol = findColIdxFlexible_(headers, ['bulan_tahun', 'bulantahun', 'periode', 'bulan'], 6);
    var crtCol = findColIdxFlexible_(headers, ['created_at', 'createdat', 'timestamp'], 7);

    var maxLen = Math.max(headers.length, 8);
    var row = new Array(maxLen);
    for (var c = 0; c < maxLen; c++) row[c] = '';

    if (idCol !== -1) row[idCol] = newId;
    if (tglCol !== -1) row[tglCol] = item.tanggal;
    if (jenisCol !== -1) row[jenisCol] = 'Pemasukan';
    if (katCol !== -1) row[katCol] = item.kategori || 'Iuran Kas';
    if (ketCol !== -1) row[ketCol] = item.keterangan || 'Iuran Masuk';
    if (nomCol !== -1) row[nomCol] = String(item.nominal || 0);
    if (blnCol !== -1) row[blnCol] = item.bulanTahun;
    if (crtCol !== -1) row[crtCol] = nowStr;

    kasSheet.appendRow(row);
  } catch (e) {
    Logger.log('Gagal log kas RT: ' + e.toString());
  }
}

function getKasSheetSafe_(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var candidates = ['KAS_RT', 'Kas_RT', 'KAS RT', 'kas_rt', 'KASRT', 'KAS'];
  for (var i = 0; i < candidates.length; i++) {
    var sh = ss.getSheetByName(candidates[i]);
    if (sh) return sh;
  }

  var allSheets = ss.getSheets();
  for (var s = 0; s < allSheets.length; s++) {
    var sName = allSheets[s].getName().trim().toUpperCase().replace(/[\s_]/g, '');
    if (sName === 'KASRT' || sName === 'KAS') {
      return allSheets[s];
    }
  }
  return null;
}

function findColIdxFlexible_(headers, candidates, fallbackIdx) {
  if (!headers || !Array.isArray(headers)) return fallbackIdx !== undefined ? fallbackIdx : -1;
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i] || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
    for (var c = 0; c < candidates.length; c++) {
      var cand = candidates[c].toLowerCase().replace(/[^a-z0-9]/g, '');
      if (h === cand || h.indexOf(cand) !== -1 || cand.indexOf(h) !== -1) {
        return i;
      }
    }
  }
  return fallbackIdx !== undefined ? fallbackIdx : -1;
}

function parseKasNominalSafe_(val) {
  if (val === undefined || val === null || val === '') return 0;
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  var clean = String(val).replace(/[^0-9]/g, '');
  return parseInt(clean, 10) || 0;
}

function generateIuranSequentialId_(prefix, sheetName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = (sheetName === 'KAS_RT') ? getKasSheetSafe_(ss) : ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) return prefix + '-0001';

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
    var maxNum = 0;
    for (var i = 0; i < data.length; i++) {
      var id = String(data[i][0]);
      if (id.indexOf(prefix + '-') === 0) {
        var num = parseInt(id.replace(prefix + '-', ''), 10);
        if (!isNaN(num) && num > maxNum) maxNum = num;
      }
    }
    var nextNum = maxNum + 1;
    var padded = ('0000' + nextNum).slice(-4);
    return prefix + '-' + padded;
  } catch (e) {
    return prefix + '-' + Math.floor(1000 + Math.random() * 9000);
  }
}
