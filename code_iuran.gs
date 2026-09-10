/**
 * ====================================================================
 * WARGAKOE - IURAN & KAS RT BACKEND MODULE (CASHFLOW & DOMPET 4 POS)
 * Menghitung Saldo Awal Mengendap (Bulan Lalu), Arus Kas Masuk & Keluar
 * Bulan Berjalan, serta Total Sisa Saldo Akhir per masing-masing pos.
 * ====================================================================
 */

function getKasRtDetails(targetBulan, targetTahun) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Daftar bulan dalam bahasa Indonesia untuk pemetaan periode kronologis
    var ID_MONTHS = ['januari', 'februari', 'maret', 'april', 'mei', 'juni', 'juli', 'agustus', 'september', 'oktober', 'november', 'desember'];

    // Menentukan index periode aktif yang sedang dilihat
    var tBulanLower = String(targetBulan || 'September').toLowerCase().trim();
    var tMonthIdx = ID_MONTHS.indexOf(tBulanLower);
    if (tMonthIdx === -1) tMonthIdx = 8; // Default September (0-indexed)
    var tYear = parseInt(targetTahun, 10) || 2026;
    var targetPeriodVal = (tYear * 12) + tMonthIdx;

    // Helper untuk mengekstrak nilai periode (Year * 12 + MonthIdx) dari string
    function extractPeriodVal_(blnStr, tglStr, crtStr) {
      blnStr = String(blnStr || '').toLowerCase().trim();
      for (var m = 0; m < ID_MONTHS.length; m++) {
        if (blnStr.indexOf(ID_MONTHS[m]) !== -1) {
          var yMatch = blnStr.match(/20\d\d/);
          var yr = yMatch ? parseInt(yMatch[0], 10) : tYear;
          return (yr * 12) + m;
        }
      }
      var dateCandidate = tglStr || crtStr || '';
      if (dateCandidate) {
        if (dateCandidate.indexOf('/') !== -1) {
          var p = dateCandidate.split('/');
          if (p.length >= 3) {
            var mNum = parseInt(p[1], 10) - 1;
            var yNum = parseInt(p[2].substring(0, 4), 10);
            if (!isNaN(mNum) && !isNaN(yNum)) return (yNum * 12) + mNum;
          }
        } else if (dateCandidate.indexOf('-') !== -1) {
          var p2 = dateCandidate.split('-');
          if (p2.length >= 3) {
            var yNum2 = parseInt(p2[0], 10);
            var mNum2 = parseInt(p2[1], 10) - 1;
            if (!isNaN(mNum2) && !isNaN(yNum2)) return (yNum2 * 12) + mNum2;
          }
        }
      }
      return targetPeriodVal; // Fallback jika tidak terdeteksi
    }

    // Inisialisasi Dompet 4 Pos Terpisah secara independen
    var dompet = {
      kasRt: { saldoAwal: 0, masuk: 0, keluar: 0, arusBulanIni: 0, saldo: 0 },
      kasDuka: { saldoAwal: 0, masuk: 0, keluar: 0, arusBulanIni: 0, saldo: 0 },
      kasSampah: { saldoAwal: 0, masuk: 0, keluar: 0, arusBulanIni: 0, saldo: 0 },
      kasSosial: { saldoAwal: 0, masuk: 0, keluar: 0, arusBulanIni: 0, saldo: 0 }
    };

    var pengeluaranList = [];

    // 1. HITUNG UANG MASUK DARI SHEET IURAN (TERPISAH SEBELUM BULAN INI VS BULAN INI)
    var iurSheet = getSheetByNameFlexible_(ss, 'IURAN');
    if (iurSheet && iurSheet.getLastRow() > 1) {
      var iData = iurSheet.getDataRange().getDisplayValues();
      var iHeaders = iData[0];

      var iBulanIdx = findColIdxFlexible_(iHeaders, ['bulan_tahun', 'bulantahun', 'periode', 'bulan'], 4);
      var iJenisIdx = findColIdxFlexible_(iHeaders, ['jenis_iuran', 'jenisiuran', 'jenis'], 5);
      var iNomIdx = findColIdxFlexible_(iHeaders, ['nominal', 'jumlah', 'total', 'rp'], 6);
      var iStatusIdx = findColIdxFlexible_(iHeaders, ['status_bayar', 'statusbayar', 'status'], 7);
      var iTglIdx = findColIdxFlexible_(iHeaders, ['tanggal_bayar', 'tanggal', 'tgl'], 8);
      var iCrtIdx = findColIdxFlexible_(iHeaders, ['created_at', 'createdat', 'timestamp'], 10);

      for (var j = 1; j < iData.length; j++) {
        var rowIur = iData[j];
        if (!rowIur || rowIur.length === 0) continue;

        var status = String(iStatusIdx !== -1 ? rowIur[iStatusIdx] : '').toLowerCase().trim();
        var isPaid = (status === 'sudah bayar' || status === 'lunas' || status === 'approved' || status === 'sukses');

        if (isPaid) {
          var nominal = parseKasNominalSafe_(iNomIdx !== -1 ? rowIur[iNomIdx] : 0);
          var jenis = String(iJenisIdx !== -1 ? rowIur[iJenisIdx] : '').toLowerCase().trim();
          var blnStr = iBulanIdx !== -1 ? rowIur[iBulanIdx] : '';
          var tglStr = iTglIdx !== -1 ? rowIur[iTglIdx] : '';
          var crtStr = iCrtIdx !== -1 ? rowIur[iCrtIdx] : '';

          var rowPeriodVal = extractPeriodVal_(blnStr, tglStr, crtStr);
          var pocketKey = null;

          if (jenis.indexOf('duka') !== -1 || jenis.indexOf('kematian') !== -1) {
            pocketKey = 'kasDuka';
          } else if (jenis.indexOf('sampah') !== -1 || jenis.indexOf('kebersihan') !== -1) {
            pocketKey = 'kasSampah';
          } else if (jenis.indexOf('sosial') !== -1) {
            pocketKey = 'kasSosial';
          } else if (jenis.indexOf('kas') !== -1 || jenis === 'iuran kas' || jenis === 'kas rt') {
            pocketKey = 'kasRt';
          }

          if (pocketKey && dompet[pocketKey]) {
            if (rowPeriodVal < targetPeriodVal) {
              // Uang masuk dari bulan-bulan sebelum bulan aktif -> mengendap di Saldo Awal
              dompet[pocketKey].saldoAwal += nominal;
            } else {
              // Uang masuk pada bulan aktif berjalan
              dompet[pocketKey].masuk += nominal;
            }
          }
        }
      }
    }

    // 2. HITUNG UANG KELUAR DARI SHEET KAS_RT (TERPISAH SEBELUM BULAN INI VS BULAN INI)
    var kasSheet = getSheetByNameFlexible_(ss, 'KAS_RT');
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

      for (var i = 1; i < rawData.length; i++) {
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

        var isKeluar = (
          colJenis === 'pengeluaran' ||
          colJenis.indexOf('pengeluaran') !== -1 ||
          colJenis.indexOf('keluar') !== -1 ||
          colJenis.indexOf('belanja') !== -1 ||
          colJenis.indexOf('biaya') !== -1
        );

        if (!isKeluar && colJenis.indexOf('masuk') === -1 && colJenis.indexOf('terima') === -1 && colNom > 0) {
          isKeluar = true;
        }

        if (isKeluar) {
          var katLower = colKat.toLowerCase();
          var pKey = null;

          if (katLower.indexOf('duka') !== -1 || katLower.indexOf('kematian') !== -1) {
            pKey = 'kasDuka';
          } else if (katLower.indexOf('sampah') !== -1 || katLower.indexOf('kebersihan') !== -1) {
            pKey = 'kasSampah';
          } else if (katLower.indexOf('sosial') !== -1) {
            pKey = 'kasSosial';
          } else if (
            katLower === 'iuran kas' ||
            katLower === 'kas rt' ||
            katLower === 'kas' ||
            (katLower.indexOf('kas') !== -1 &&
             katLower.indexOf('duka') === -1 &&
             katLower.indexOf('kematian') === -1 &&
             katLower.indexOf('sampah') === -1 &&
             katLower.indexOf('sosial') === -1 &&
             katLower.indexOf('kebersihan') === -1 &&
             katLower.indexOf('fasilitas') === -1)
          ) {
            pKey = 'kasRt';
          }

          var rowPerValKas = extractPeriodVal_(colBulan, colTgl, colCreated);

          if (pKey && dompet[pKey]) {
            if (rowPerValKas < targetPeriodVal) {
              // Pengeluaran dari bulan sebelum bulan aktif -> memotong Saldo Awal
              dompet[pKey].saldoAwal -= colNom;
            } else {
              // Pengeluaran pada bulan aktif berjalan
              dompet[pKey].keluar += colNom;
            }
          }

          // Catatan pengeluaran masuk ke Buku Transparansi
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
    }

    // 3. HITUNG ARUS KAS BULAN INI DAN TOTAL SISA SALDO AKHIR
    var pocketNames = ['kasRt', 'kasDuka', 'kasSampah', 'kasSosial'];
    for (var k = 0; k < pocketNames.length; k++) {
      var key = pocketNames[k];
      dompet[key].arusBulanIni = dompet[key].masuk - dompet[key].keluar;
      dompet[key].saldo = dompet[key].saldoAwal + dompet[key].arusBulanIni;
    }

    pengeluaranList.reverse();

    return {
      success: true,
      data: {
        dompet: dompet,
        pengeluaranList: pengeluaranList,
        targetBulan: targetBulan,
        targetTahun: targetTahun
      }
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
    var kasSheet = getSheetByNameFlexible_(ss, 'KAS_RT');
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

    var tglFormatted = String(payload.tanggal || Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy')).trim();
    if (tglFormatted.indexOf('-') !== -1) {
      var parts = tglFormatted.split('-');
      if (parts.length === 3) tglFormatted = parts[2] + '/' + parts[1] + '/' + parts[0];
    }

    if (idCol !== -1) newRow[idCol] = newId;
    if (tglCol !== -1) newRow[tglCol] = tglFormatted;
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
    var iurSheet = getSheetByNameFlexible_(ss, 'IURAN');

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

      if (filterJenis === 'iuran sampah' && (jenisIuran.toLowerCase().indexOf('sampah') !== -1 || jenisIuran.toLowerCase().indexOf('kebersihan') !== -1)) {
        matchJenis = true;
      }

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
    var iurSheet = getSheetByNameFlexible_(ss, 'IURAN');
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
    var iurSheet = getSheetByNameFlexible_(ss, 'IURAN');
    if (!iurSheet) throw new Error('Sheet IURAN tidak ditemukan!');

    var data = iurSheet.getDataRange().getDisplayValues();
    var headers = data[0];

    var idIdx = findColIdxFlexible_(headers, ['id_iuran', 'idiuran', 'id'], 0);
    var statusIdx = findColIdxFlexible_(headers, ['status_bayar', 'statusbayar', 'status'], 7);
    var appIdx = findColIdxFlexible_(headers, ['approved_by', 'approvedby', 'petugas'], 9);
    var tglIdx = findColIdxFlexible_(headers, ['tanggal_bayar', 'tanggal', 'tgl'], 8);

    var targetId = String(payload.idIuran || '').trim();
    var newStatus = String(payload.status || 'Sudah bayar').trim();
    var verifikator = String(payload.verifikator || 'Pengurus RT').trim();
    var todayStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy');

    var foundRow = -1;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idIdx]).trim() === targetId) {
        foundRow = i + 1;
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
    var iurSheet = getSheetByNameFlexible_(ss, 'IURAN');
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
        var match = (filter === 'semua' || jenis.toLowerCase() === filter);
        if (filter === 'iuran sampah' && (jenis.toLowerCase().indexOf('sampah') !== -1 || jenis.toLowerCase().indexOf('kebersihan') !== -1)) {
          match = true;
        }

        if (match) {
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

function getSheetByNameFlexible_(ss, targetName) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var direct = ss.getSheetByName(targetName);
  if (direct) return direct;

  var targetClean = targetName.toLowerCase().replace(/[^a-z0-9]/g, '');
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sNameClean = sheets[i].getName().toLowerCase().replace(/[^a-z0-9]/g, '');
    if (sNameClean === targetClean) {
      return sheets[i];
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
      if (h === cand) return i;
    }
  }

  for (var j = 0; j < headers.length; j++) {
    var h2 = String(headers[j] || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
    for (var k = 0; k < candidates.length; k++) {
      var cand2 = candidates[k].toLowerCase().replace(/[^a-z0-9]/g, '');
      if (h2.indexOf(cand2) !== -1 || cand2.indexOf(h2) !== -1) return j;
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
    var sheet = getSheetByNameFlexible_(ss, sheetName);
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
