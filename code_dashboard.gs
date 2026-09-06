/**
 * ====================================================================
 * WARGAKOE - DASHBOARD STATISTICS & SUMMARY BACKEND MODULE
 * Menghitung demografi usia (USERS + ANGGOTA_KELUARGA),
 * realisasi iuran bulan berjalan, kas terkumpul (murni cashflow pos Iuran Kas:
 * penerimaan Iuran Kas dikurangi pengeluaran kategori Iuran Kas), dan notifikasi.
 * ====================================================================
 */

function getDashboardStats(noKk, targetMonthYear) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var totalWarga = 0;
    var totalKk = 0;
    var balita = 0;   // < 5 tahun
    var remaja = 0;   // 5 - 17 tahun
    var dewasa = 0;   // 18 - 59 tahun
    var lansia = 0;   // >= 60 tahun

    // 1. HITUNG DEMOGRAFI DARI TABEL USERS (KEPALA KELUARGA)
    var userSheet = ss.getSheetByName('USERS');
    if (userSheet && userSheet.getLastRow() > 1) {
      var uData = userSheet.getDataRange().getDisplayValues();
      var uHeaders = uData[0];
      var uUmurIdx = uHeaders.indexOf('Umur');
      var uTglIdx = uHeaders.indexOf('Tanggal_Lahir');

      totalKk = uData.length - 1;
      totalWarga += totalKk;

      for (var u = 1; u < uData.length; u++) {
        var uRow = uData[u];
        var uUmurVal = uUmurIdx !== -1 ? uRow[uUmurIdx] : '';
        var uTglVal = uTglIdx !== -1 ? uRow[uTglIdx] : '';
        var uAge = parseDashboardAge_(uUmurVal, uTglVal);

        if (uAge >= 0) {
          if (uAge < 5) balita++;
          else if (uAge <= 17) remaja++;
          else if (uAge < 60) dewasa++;
          else lansia++;
        }
      }
    }

    // 2. HITUNG DEMOGRAFI DARI TABEL ANGGOTA_KELUARGA
    var agtSheet = ss.getSheetByName('ANGGOTA_KELUARGA');
    if (agtSheet && agtSheet.getLastRow() > 1) {
      var aData = agtSheet.getDataRange().getDisplayValues();
      var aHeaders = aData[0];
      var aUmurIdx = aHeaders.indexOf('Umur');
      var aTglIdx = aHeaders.indexOf('Tanggal_Lahir');

      totalWarga += (aData.length - 1);

      for (var a = 1; a < aData.length; a++) {
        var aRow = aData[a];
        var aUmurVal = aUmurIdx !== -1 ? aRow[aUmurIdx] : '';
        var aTglVal = aTglIdx !== -1 ? aRow[aTglIdx] : '';
        var aAge = parseDashboardAge_(aUmurVal, aTglVal);

        if (aAge >= 0) {
          if (aAge < 5) balita++;
          else if (aAge <= 17) remaja++;
          else if (aAge < 60) dewasa++;
          else lansia++;
        }
      }
    }

    // 3. HITUNG REALISASI IURAN & PENERIMAAN IURAN KAS DARI TABEL IURAN
    var now = new Date();
    var monthsIndo = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
    var currentMonthName = monthsIndo[now.getMonth()];
    var currentYear = String(now.getFullYear());

    var activeMonthStr = (targetMonthYear && String(targetMonthYear).trim() !== '') 
      ? String(targetMonthYear).trim() 
      : (currentMonthName + ' ' + currentYear);

    var totalPenerimaanIuranKas = 0;
    var menungguApproval = 0;
    var pendingIuranCount = 0;

    var realisasi = {
      kasRt: 0,
      kasKematian: 0,
      iuranSampah: 0,
      iuranSosial: 0
    };

    var iurSheet = ss.getSheetByName('IURAN');
    if (iurSheet && iurSheet.getLastRow() > 1) {
      var iData = iurSheet.getDataRange().getDisplayValues();
      var iHeaders = iData[0];

      var iNoKkIdx = iHeaders.indexOf('No_KK');
      var iBulanIdx = iHeaders.indexOf('Bulan_Tahun');
      var iJenisIdx = iHeaders.indexOf('Jenis_Iuran');
      var iNominalIdx = iHeaders.indexOf('Nominal');
      var iStatusIdx = iHeaders.indexOf('Status_Bayar');

      var targetNoKk = String(noKk || '').trim();

      for (var i = 1; i < iData.length; i++) {
        var row = iData[i];
        var rNoKk = iNoKkIdx !== -1 ? String(row[iNoKkIdx]).trim() : '';
        var rBulan = iBulanIdx !== -1 ? String(row[iBulanIdx]).trim() : '';
        var rJenis = iJenisIdx !== -1 ? String(row[iJenisIdx]).trim() : '';
        var rNominal = parseDashboardNominal_(iNominalIdx !== -1 ? row[iNominalIdx] : 0);
        var rStatus = iStatusIdx !== -1 ? String(row[iStatusIdx]).trim() : '';

        var isPaid = isDashboardPaidStatus_(rStatus);
        var isPending = isDashboardPendingStatus_(rStatus);

        var jLower = rJenis.toLowerCase();
        var isKasRt = (jLower.indexOf('kas') !== -1 && jLower.indexOf('duka') === -1 && jLower.indexOf('kematian') === -1) || jLower === 'iuran kas';

        // Akumulasi penerimaan kas terkumpul HANYA dari pos Iuran Kas yang sudah lunas
        if (isPaid && isKasRt) {
          totalPenerimaanIuranKas += rNominal;
        }

        // Akumulasi iuran yang sedang menunggu verifikasi pengurus
        if (isPending) {
          menungguApproval += rNominal;
        }

        // Pengecekan realisasi iuran bulan berjalan untuk widget dashboard 4 pos iuran
        if (isDashboardMatchMonth_(rBulan, activeMonthStr, currentMonthName, currentYear) && isPaid) {
          if (isKasRt) {
            realisasi.kasRt += rNominal;
          } else if (jLower.indexOf('duka') !== -1 || jLower.indexOf('kematian') !== -1) {
            realisasi.kasKematian += rNominal;
          } else if (jLower.indexOf('sampah') !== -1) {
            realisasi.iuranSampah += rNominal;
          } else if (jLower.indexOf('sosial') !== -1) {
            realisasi.iuranSosial += rNominal;
          }
        }

        // Pengecekan tagihan belum lunas untuk akun warga yang sedang login
        if (targetNoKk && rNoKk === targetNoKk && !isPaid) {
          pendingIuranCount++;
        }
      }
    }

    // 4. HITUNG PENGELUARAN MURNI DARI KATEGORI "IURAN KAS"
    var totalPengeluaranKas = 0;
    var kasSheet = ss.getSheetByName('KAS_RT');
    if (kasSheet && kasSheet.getLastRow() > 1) {
      var kData = kasSheet.getDataRange().getDisplayValues();
      var kHeaders = kData[0];
      var kJenisIdx = kHeaders.indexOf('Jenis_Kas');
      var kKategoriIdx = kHeaders.indexOf('Kategori');
      var kNominalIdx = kHeaders.indexOf('Nominal');

      for (var k = 1; k < kData.length; k++) {
        var kRow = kData[k];
        var kJenis = kJenisIdx !== -1 ? String(kRow[kJenisIdx]).trim().toLowerCase() : '';
        var kKategori = kKategoriIdx !== -1 ? String(kRow[kKategoriIdx]).trim().toLowerCase() : '';
        var kNominal = parseDashboardNominal_(kNominalIdx !== -1 ? kRow[kNominalIdx] : 0);

        // HANYA ambil pengeluaran yang MURNI berkategori 'Iuran Kas' / 'Kas RT'
        // Kategori lain (Kebersihan, Fasilitas, Iuran Sampah, Duka, Sosial) TIDAK dicampurkan ke kas ini
        if (kJenis === 'pengeluaran' || kJenis === 'keluar') {
          var isPureKasExpense = (kKategori === 'iuran kas' || kKategori === 'kas rt' || kKategori === 'kas' || 
            (kKategori.indexOf('kas') !== -1 && 
             kKategori.indexOf('duka') === -1 && 
             kKategori.indexOf('kematian') === -1 && 
             kKategori.indexOf('sampah') === -1 && 
             kKategori.indexOf('sosial') === -1 && 
             kKategori.indexOf('kebersihan') === -1 && 
             kKategori.indexOf('fasilitas') === -1));

          if (isPureKasExpense) {
            totalPengeluaranKas += kNominal;
          }
        }
      }
    }

    // Kas terkumpul bersih = Total Penerimaan Iuran Kas - Pengeluaran Murni Iuran Kas
    var kasTerkumpul = totalPenerimaanIuranKas - totalPengeluaranKas;

    // 5. HITUNG PERMOHONAN PINJAM ASET YANG MENUNGGU (UNTUK PENGURUS)
    var pendingAsetCount = 0;
    var pinjamSheet = ss.getSheetByName('PEMINJAMAN_ASET');
    if (pinjamSheet && pinjamSheet.getLastRow() > 1) {
      var pData = pinjamSheet.getDataRange().getDisplayValues();
      var pHeaders = pData[0];
      var pStatusIdx = pHeaders.indexOf('Status');
      if (pStatusIdx !== -1) {
        for (var p = 1; p < pData.length; p++) {
          var pStatus = String(pData[p][pStatusIdx]).trim().toLowerCase();
          if (pStatus === 'menunggu' || pStatus === 'pending') {
            pendingAsetCount++;
          }
        }
      }
    }

    // 6. AMBIL JADWAL RONDA WARGA TERKAIT
    var myRonda = null;
    var rondaSheet = ss.getSheetByName('JADWAL_RONDA');
    if (rondaSheet && rondaSheet.getLastRow() > 1) {
      var rData = rondaSheet.getDataRange().getDisplayValues();
      var rHeaders = rData[0];
      var rNamaIdx = rHeaders.indexOf('Nama_Regu');
      var rTglIdx = rHeaders.indexOf('Tanggal_Ronda');
      var rShiftIdx = rHeaders.indexOf('Jam_Shift');
      var rPosIdx = rHeaders.indexOf('Lokasi_Pos');
      var rJsonIdx = rHeaders.indexOf('Daftar_Warga_JSON');

      var targetKk = String(noKk || '').trim();

      for (var r = rData.length - 1; r >= 1; r--) {
        var rRow = rData[r];
        var rawJson = rJsonIdx !== -1 ? rRow[rJsonIdx] : '[]';
        var isMember = false;

        if (targetKk) {
          try {
            var listWarga = JSON.parse(rawJson);
            if (Array.isArray(listWarga)) {
              for (var w = 0; w < listWarga.length; w++) {
                if (String(listWarga[w].noKk).trim() === targetKk) {
                  isMember = true;
                  break;
                }
              }
            }
          } catch (e) {
            if (rawJson.indexOf(targetKk) !== -1) isMember = true;
          }
        } else {
          isMember = true;
        }

        if (isMember) {
          myRonda = {
            namaRegu: rNamaIdx !== -1 ? rRow[rNamaIdx] : 'Regu Ronda',
            tanggalRonda: rTglIdx !== -1 ? rRow[rTglIdx] : '',
            shift: rShiftIdx !== -1 ? rRow[rShiftIdx] : '22.00 - 03.00 WIB',
            lokasi: rPosIdx !== -1 ? rRow[rPosIdx] : 'Pos Ronda Utama RT 010'
          };
          break;
        }
      }
    }

    return {
      success: true,
      data: {
        totalWarga: totalWarga,
        totalKk: totalKk,
        kasTerkumpul: kasTerkumpul,
        menungguApproval: menungguApproval,
        demografi: {
          balita: balita,
          remaja: remaja,
          dewasa: dewasa,
          lansia: lansia
        },
        realisasi: {
          kasRt: realisasi.kasRt,
          kasKematian: realisasi.kasKematian,
          iuranSampah: realisasi.iuranSampah,
          iuranSosial: realisasi.iuranSosial
        },
        pendingIuranCount: pendingIuranCount,
        pendingAsetCount: pendingAsetCount,
        myRonda: myRonda
      }
    };
  } catch (error) {
    return {
      success: false,
      message: 'Gagal mengambil data statistik dashboard: ' + error.toString()
    };
  }
}

// Alias untuk kompatibilitas frontend
function getDashboardSummary(noKk, targetMonthYear) {
  return getDashboardStats(noKk, targetMonthYear);
}

/**
 * Membaca umur dari nilai kolom 'Umur' (angka/teks) atau menghitung dari 'Tanggal_Lahir'
 */
function parseDashboardAge_(ageVal, birthDateStr) {
  if (ageVal !== undefined && ageVal !== null && String(ageVal).trim() !== '') {
    var cleanNum = String(ageVal).replace(/[^0-9]/g, '');
    if (cleanNum !== '') {
      var parsed = parseInt(cleanNum, 10);
      if (!isNaN(parsed) && parsed >= 0) {
        return parsed;
      }
    }
  }

  if (birthDateStr) {
    var str = String(birthDateStr).trim();
    var bDate = null;

    if (str.indexOf('/') !== -1) {
      var parts = str.split('/');
      if (parts.length === 3) {
        bDate = new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
      }
    } else if (str.indexOf('-') !== -1) {
      var pIso = str.split('-');
      if (pIso.length === 3) {
        bDate = new Date(parseInt(pIso[0], 10), parseInt(pIso[1], 10) - 1, parseInt(pIso[2], 10));
      }
    }

    if (bDate && !isNaN(bDate.getTime())) {
      var today = new Date();
      var calculatedAge = today.getFullYear() - bDate.getFullYear();
      var m = today.getMonth() - bDate.getMonth();
      if (m < 0 || (m === 0 && today.getDate() < bDate.getDate())) {
        calculatedAge--;
      }
      return Math.max(0, calculatedAge);
    }
  }

  return -1;
}

function parseDashboardNominal_(val) {
  if (!val) return 0;
  if (typeof val === 'number') return val;
  var clean = String(val).replace(/[^0-9]/g, '');
  return parseInt(clean, 10) || 0;
}

function isDashboardPaidStatus_(statusStr) {
  if (!statusStr) return false;
  var s = String(statusStr).toLowerCase().trim();
  return s === 'sudah bayar' || s === 'lunas' || s === 'approved' || s === 'sukses' || s === 'sudah';
}

function isDashboardPendingStatus_(statusStr) {
  if (!statusStr) return false;
  var s = String(statusStr).toLowerCase().trim();
  return s === 'menunggu' || s === 'menunggu approval' || s === 'pending';
}

function isDashboardMatchMonth_(rowMonthStr, targetMonthYear, curMonthName, curYear) {
  if (!rowMonthStr) return false;
  var row = String(rowMonthStr).toLowerCase().trim();
  var target = String(targetMonthYear).toLowerCase().trim();
  var curM = String(curMonthName).toLowerCase().trim();
  var curY = String(curYear).toLowerCase().trim();

  if (row === target) return true;

  if (target.indexOf(' ') !== -1) {
    var parts = target.split(' ');
    if (parts.length >= 2) {
      if (row.indexOf(parts[0]) !== -1 && row.indexOf(parts[1]) !== -1) return true;
    }
  }

  if (row.indexOf(curM) !== -1 && (row.indexOf(curY) !== -1 || row === curM)) return true;

  return false;
}
