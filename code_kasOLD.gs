/**
 * ====================================================================
 * WARGAKOE - KAS RT LEDGER & DASHBOARD SUMMARY BACKEND MODULE
 * By Zettbos System (ZettBOT 3.1)
 * ====================================================================
 */

function getKasRtDetails(bulanTahun) {
  try {
    var sheet = getSheet_('KAS_RT');
    var rawData = sheet.getDataRange().getDisplayValues();
    if (rawData.length <= 1) {
      return { success: true, data: [], totalPemasukan: 0, totalPengeluaran: 0, sisaSaldo: 0 };
    }

    var headers = rawData[0];
    var idIdx = headers.indexOf('ID_Kas');
    var tglIdx = headers.indexOf('Tanggal');
    var jenisIdx = headers.indexOf('Jenis_Kas');
    var katIdx = headers.indexOf('Kategori');
    var ketIdx = headers.indexOf('Keterangan');
    var nomIdx = headers.indexOf('Nominal');
    var bulanIdx = headers.indexOf('Bulan_Tahun');

    var list = [];
    var totalPemasukan = 0;
    var totalPengeluaran = 0;

    for (var i = 1; i < rawData.length; i++) {
      var row = rawData[i];
      var rowBulan = row[bulanIdx];
      var jenis = row[jenisIdx];
      var nominal = parseFloat(row[nomIdx]) || 0;

      if (!bulanTahun || bulanTahun === 'SEMUA' || rowBulan === bulanTahun) {
        if (jenis === 'Pemasukan') totalPemasukan += nominal;
        else if (jenis === 'Pengeluaran') totalPengeluaran += nominal;

        list.push({
          idKas: row[idIdx],
          tanggal: row[tglIdx],
          jenisKas: jenis,
          kategori: row[katIdx],
          keterangan: row[ketIdx],
          nominal: nominal,
          bulanTahun: rowBulan
        });
      }
    }

    var sisaSaldo = totalPemasukan - totalPengeluaran;
    return {
      success: true,
      data: list.reverse(),
      totalPemasukan: totalPemasukan,
      totalPengeluaran: totalPengeluaran,
      sisaSaldo: sisaSaldo
    };
  } catch (error) {
    return { success: false, message: 'Gagal memuat detail Kas RT: ' + error.toString() };
  }
}

function savePengeluaranRutin(payload) {
  try {
    var sheet = getSheet_('KAS_RT');
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var newId = generateSequentialId_('KAS', 'KAS_RT');

    sheet.appendRow([
      newId,
      String(payload.tanggal || '').trim(),
      'Pengeluaran',
      String(payload.kategori || 'Iuran Kas').trim(),
      String(payload.keterangan || '').trim(),
      String(payload.nominal || '0').trim(),
      String(payload.bulanTahun || '').trim(),
      nowStr
    ]);
    SpreadsheetApp.flush();
    return { success: true, message: 'Pengeluaran rutin berhasil dicatat!' };
  } catch (error) {
    return { success: false, message: 'Gagal mencatat pengeluaran: ' + error.toString() };
  }
}

function getDashboardSummary(userRole, userNoKk) {
  try {
    var userSheet = getSheet_('USERS');
    var agtSheet = getSheet_('ANGGOTA_KELUARGA');
    var iurSheet = getSheet_('IURAN');
    var madingSheet = getSheet_('MADING');
    var pinjamSheet = getSheet_('PEMINJAMAN_ASET');

    var userData = userSheet.getDataRange().getDisplayValues();
    var agtData = agtSheet.getDataRange().getDisplayValues();
    var iurData = iurSheet.getDataRange().getDisplayValues();
    var madingData = madingSheet.getDataRange().getDisplayValues();
    var pinjamData = pinjamSheet.getDataRange().getDisplayValues();

    var totalKk = Math.max(0, userData.length - 1);
    var totalWarga = totalKk + Math.max(0, agtData.length - 1);

    var stats = {
      balita: 0,
      remaja: 0,
      dewasa: 0,
      lansia: 0,
      totalWarga: totalWarga,
      totalKk: totalKk,
      kasTerkumpul: 0,
      kasRtBulanIni: 0,
      kasKematianBulanIni: 0,
      iuranSampahBulanIni: 0,
      iuranSosialBulanIni: 0,
      iuranLainBulanIni: 0,
      totalTerkumpul: 0,
      totalMenunggu: 0,
      totalBelumBayar: 0,
      unpaidDuesList: [],
      madingTerbaru: [],
      pendingAsetCount: 0,
      upcomingRonda: null
    };

    var uHeaders = userData[0];
    var uUmurIdx = uHeaders.indexOf('Umur');
    for (var u = 1; u < userData.length; u++) {
      var umur = parseInt(userData[u][uUmurIdx], 10) || 0;
      if (umur < 5) stats.balita++;
      else if (umur <= 17) stats.remaja++;
      else if (umur <= 59) stats.dewasa++;
      else stats.lansia++;
    }

    if (agtData.length > 1) {
      var aHeaders = agtData[0];
      var aUmurIdx = aHeaders.indexOf('Umur');
      for (var a = 1; a < agtData.length; a++) {
        var aUmur = parseInt(agtData[a][aUmurIdx], 10) || 0;
        if (aUmur < 5) stats.balita++;
        else if (aUmur <= 17) stats.remaja++;
        else if (aUmur <= 59) stats.dewasa++;
        else stats.lansia++;
      }
    }

    if (iurData.length > 1) {
      var iHeaders = iurData[0];
      var iKkIdx = iHeaders.indexOf('No_KK');
      var iJenisIdx = iHeaders.indexOf('Jenis_Iuran');
      var iNominalIdx = iHeaders.indexOf('Nominal');
      var iStatusIdx = iHeaders.indexOf('Status_Bayar');

      for (var k = 1; k < iurData.length; k++) {
        var row = iurData[k];
        var idIur = row[iHeaders.indexOf('ID_Iuran')] || '';
        var kwiNo = row[iHeaders.indexOf('No_Kuitansi')] || '';

        if (idIur.indexOf('KWI-IUR-') === 0 || kwiNo.indexOf('KWI-IUR-') === 0) continue;

        var nominal = parseFloat(row[iNominalIdx]) || 0;
        var status = (row[iStatusIdx] || '').trim();
        var jenis = row[iJenisIdx];
        var kk = row[iKkIdx];

        if (status === 'Sudah bayar' || status === 'Lunas') {
          stats.totalTerkumpul += nominal;
          if (jenis === 'Iuran Kas' || jenis === 'Iuran RT') {
            stats.kasTerkumpul += nominal;
            stats.kasRtBulanIni += nominal;
          } else if (jenis === 'Iuran Duka' || jenis === 'Iuran Kematian') {
            stats.kasKematianBulanIni += nominal;
          } else if (jenis === 'Iuran Sampah') {
            stats.iuranSampahBulanIni += nominal;
          } else if (jenis === 'Iuran Sosial') {
            stats.iuranSosialBulanIni += nominal;
          } else {
            stats.iuranLainBulanIni += nominal;
          }
        } else if (status === 'Menunggu' || status === 'Menunggu Approval') {
          stats.totalMenunggu += nominal;
        } else {
          stats.totalBelumBayar += nominal;
        }

        if (userRole !== 'Super Admin' && kk === userNoKk) {
          if (status !== 'Sudah bayar' && status !== 'Lunas') {
            stats.unpaidDuesList.push({
              idIuran: idIur,
              noKuitansi: kwiNo,
              namaWarga: row[iHeaders.indexOf('Nama_Warga')],
              jenisIuran: jenis,
              bulanTahun: row[iHeaders.indexOf('Bulan_Tahun')],
              nominal: nominal,
              statusBayar: status
            });
          }
        }
      }
    }

    if (madingData.length > 1) {
      var mHeaders = madingData[0];
      var mList = [];
      for (var m = 1; m < madingData.length; m++) {
        if (madingData[m][mHeaders.indexOf('Status_Publish')] === 'Published') {
          mList.push({
            idMading: madingData[m][mHeaders.indexOf('ID_Mading')],
            judul: madingData[m][mHeaders.indexOf('Judul')],
            tanggalKegiatan: madingData[m][mHeaders.indexOf('Tanggal_Kegiatan')],
            note: madingData[m][mHeaders.indexOf('Note')],
            pembuat: madingData[m][mHeaders.indexOf('Pembuat')],
            createdAt: madingData[m][mHeaders.indexOf('Created_At')]
          });
        }
      }
      stats.madingTerbaru = mList.slice(-3).reverse();
    }

    if (pinjamData.length > 1) {
      var pHeaders = pinjamData[0];
      var pStatusIdx = pHeaders.indexOf('Status');
      for (var p = 1; p < pinjamData.length; p++) {
        var pStatus = (pinjamData[p][pStatusIdx] || '').trim();
        if (pStatus === 'Menunggu') {
          stats.pendingAsetCount++;
        }
      }
    }

    // Ambil pengingat ronda mendatang untuk warga
    var rondaNotif = getRondaWargaNotification(userNoKk);
    if (rondaNotif.hasRonda) {
      stats.upcomingRonda = rondaNotif.ronda;
    }

    return { success: true, data: stats };
  } catch (error) {
    return { success: false, message: 'Gagal memuat statistik: ' + error.toString() };
  }
}
