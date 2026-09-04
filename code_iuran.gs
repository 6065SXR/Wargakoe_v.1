/**
 * ====================================================================
 * WARGAKOE - DUES PAYMENT & FINANCIAL REPORT BACKEND MODULE
 * By Zettbos System (ZettBOT 3.1)
 * ====================================================================
 */

function getIuranList(params) {
  try {
    params = params || {};
    var limit = parseInt(params.limit, 10) || 10;
    var offset = parseInt(params.offset, 10) || 0;
    var search = (params.search || '').toLowerCase().trim();
    var filterStatus = params.statusBayar || 'SEMUA';
    var filterJenis = params.jenisIuran || 'SEMUA';
    var userRole = params.role || 'Warga';
    var filterKk = params.noKk || '';

    var sheet = getSheet_('IURAN');
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];

    if (headers.indexOf('No_Kuitansi') !== 1) {
      repairIuranSheet_(SpreadsheetApp.getActiveSpreadsheet());
      sheet = getSheet_('IURAN');
      headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    }

    var rawData = sheet.getDataRange().getDisplayValues();
    if (rawData.length <= 1) return { success: true, data: [], total: 0, page: 1, totalPages: 0 };

    var filtered = [];
    for (var i = 1; i < rawData.length; i++) {
      var row = rawData[i];
      var id = row[headers.indexOf('ID_Iuran')] || '';
      var kwi = row[headers.indexOf('No_Kuitansi')] || '';

      if (id.indexOf('KWI-IUR-') === 0 || kwi.indexOf('KWI-IUR-') === 0) continue;

      var kk = row[headers.indexOf('No_KK')] || '';
      var nama = (row[headers.indexOf('Nama_Warga')] || '').toLowerCase();
      var jenis = row[headers.indexOf('Jenis_Iuran')] || '';
      var status = (row[headers.indexOf('Status_Bayar')] || '').trim();

      if (status === 'Menunggu Approval' || status === 'Pending') status = 'Menunggu';
      if (status === 'Lunas') status = 'Sudah bayar';

      if (userRole === 'Warga' && filterKk && kk !== filterKk) {
        continue;
      }

      var matchSearch = !search || (nama.indexOf(search) !== -1 || kk.indexOf(search) !== -1 || id.toLowerCase().indexOf(search) !== -1 || kwi.toLowerCase().indexOf(search) !== -1);
      var matchStatus = filterStatus === 'SEMUA' || status === filterStatus || (filterStatus === 'Menunggu' && (status === 'Menunggu' || status === 'Menunggu Approval'));
      var matchJenis = filterJenis === 'SEMUA' || jenis === filterJenis;

      if (matchSearch && matchStatus && matchJenis) {
        filtered.push({
          idIuran: id,
          noKuitansi: kwi,
          noKk: kk,
          namaWarga: row[headers.indexOf('Nama_Warga')],
          bulanTahun: row[headers.indexOf('Bulan_Tahun')],
          jenisIuran: jenis,
          nominal: row[headers.indexOf('Nominal')],
          statusBayar: status,
          tanggalBayar: row[headers.indexOf('Tanggal_Bayar')],
          approvedBy: row[headers.indexOf('Approved_By')]
        });
      }
    }

    filtered.sort(function(a, b) {
      var aIsPending = (a.statusBayar === 'Menunggu' || a.statusBayar === 'Menunggu Approval');
      var bIsPending = (b.statusBayar === 'Menunggu' || b.statusBayar === 'Menunggu Approval');
      if (aIsPending && !bIsPending) return -1;
      if (!aIsPending && bIsPending) return 1;
      return 0;
    });

    var total = filtered.length;
    var paginated = filtered.slice(offset, offset + limit);
    var totalPages = Math.ceil(total / limit) || 1;
    var currentPage = Math.floor(offset / limit) + 1;

    return { success: true, data: paginated, total: total, page: currentPage, totalPages: totalPages };
  } catch (error) {
    return { success: false, message: 'Gagal memuat data iuran: ' + error.toString() };
  }
}

function submitBayarIuran(payload) {
  try {
    var sheet = getSheet_('IURAN');
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var newId = generateSequentialId_('IUR', 'IURAN');

    var newRow = [
      newId,
      String(payload.noKuitansi || '').trim(),
      String(payload.noKk || '').trim(),
      String(payload.namaWarga || '').trim(),
      String(payload.bulanTahun || '').trim(),
      String(payload.jenisIuran || 'Iuran Kas').trim(),
      String(payload.nominal || '0').trim(),
      'Menunggu',
      String(payload.tanggalBayar || nowStr).trim(),
      '-',
      nowStr
    ];

    sheet.appendRow(newRow);
    SpreadsheetApp.flush();

    return { success: true, message: 'Pembayaran iuran berhasil dikirim! Status: Menunggu verifikasi Pengurus RT.' };
  } catch (error) {
    return { success: false, message: 'Gagal kirim pembayaran: ' + error.toString() };
  }
}

function approveIuran(idIuran, approverName, action) {
  try {
    var sheet = getSheet_('IURAN');
    var data = sheet.getDataRange().getDisplayValues();
    var headers = data[0];
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var newStatus = action === 'APPROVE' ? 'Sudah bayar' : 'Ditolak';

    var idIdx = headers.indexOf('ID_Iuran');
    var statusIdx = headers.indexOf('Status_Bayar');
    var approvedIdx = headers.indexOf('Approved_By');

    for (var i = 1; i < data.length; i++) {
      if (data[i][idIdx] === idIuran) {
        var r = i + 1;
        sheet.getRange(r, statusIdx + 1).setValue(newStatus);
        sheet.getRange(r, approvedIdx + 1).setValue(approverName || 'Pengurus RT');
        
        if (action === 'APPROVE') {
          var kasSheet = getSheet_('KAS_RT');
          var kasId = generateSequentialId_('KAS', 'KAS_RT');
          kasSheet.appendRow([
            kasId,
            nowStr.substring(0, 10),
            'Pemasukan',
            data[i][headers.indexOf('Jenis_Iuran')],
            'Iuran dari ' + data[i][headers.indexOf('Nama_Warga')] + ' (' + data[i][headers.indexOf('No_Kuitansi')] + ')',
            data[i][headers.indexOf('Nominal')],
            data[i][headers.indexOf('Bulan_Tahun')],
            nowStr
          ]);
        }
        
        SpreadsheetApp.flush();
        return { 
          success: true, 
          message: action === 'APPROVE' ? 'Pembayaran iuran DITERIMA (Sudah bayar).' : 'Pembayaran iuran DITOLAK.' 
        };
      }
    }
    return { success: false, message: 'ID Iuran ' + idIuran + ' tidak ditemukan.' };
  } catch (error) {
    return { success: false, message: 'Gagal memproses verifikasi: ' + error.toString() };
  }
}

function getLaporanKeuanganData(jenisFilter) {
  try {
    var iurSheet = getSheet_('IURAN');
    var iurData = iurSheet.getDataRange().getDisplayValues();
    if (iurData.length <= 1) return { success: true, data: [] };

    var headers = iurData[0];
    var list = [];
    var totalNominal = 0;

    for (var i = 1; i < iurData.length; i++) {
      var row = iurData[i];
      var idIur = row[headers.indexOf('ID_Iuran')] || '';
      var kwiNo = row[headers.indexOf('No_Kuitansi')] || '';

      if (idIur.indexOf('KWI-IUR-') === 0 || kwiNo.indexOf('KWI-IUR-') === 0) continue;

      var jenis = row[headers.indexOf('Jenis_Iuran')];
      var status = (row[headers.indexOf('Status_Bayar')] || '').trim();

      var matchFilter = (!jenisFilter || jenisFilter === 'SEMUA' || jenis === jenisFilter);
      if (matchFilter && (status === 'Sudah bayar' || status === 'Lunas')) {
        var nom = parseFloat(row[headers.indexOf('Nominal')]) || 0;
        totalNominal += nom;
        list.push({
          noKuitansi: kwiNo,
          namaWarga: row[headers.indexOf('Nama_Warga')],
          jenisIuran: jenis,
          bulanTahun: row[headers.indexOf('Bulan_Tahun')],
          nominal: nom,
          tanggalBayar: row[headers.indexOf('Tanggal_Bayar')],
          approvedBy: row[headers.indexOf('Approved_By')]
        });
      }
    }

    return {
      success: true,
      data: list,
      totalNominal: totalNominal,
      kopInfo: {
        rtRw: 'PENGURUS RT 010 RW 05',
        alamat: 'Kamp. Baru I Jl. Marga Mulya RT 010 RW 05 Halim Perdana Kusuma'
      }
    };
  } catch (error) {
    return { success: false, message: 'Gagal memuat laporan keuangan: ' + error.toString() };
  }
}
