/**
 * ====================================================================
 * WARGAKOE - BULLETIN BOARD (MADING) BACKEND MODULE
 * By Zettbos System (ZettBOT 3.1)
 * ====================================================================
 */

function getMadingList(params) {
  try {
    params = params || {};
    var search = (params.search || '').toLowerCase().trim();
    var userRole = params.role || 'Warga';

    var sheet = getSheet_('MADING');
    var rawData = sheet.getDataRange().getDisplayValues();
    if (rawData.length <= 1) return { success: true, data: [] };

    var headers = rawData[0];
    var list = [];
    for (var i = 1; i < rawData.length; i++) {
      var row = rawData[i];
      var status = row[headers.indexOf('Status_Publish')];
      var judul = row[headers.indexOf('Judul')].toLowerCase();
      var note = row[headers.indexOf('Note')].toLowerCase();

      if (userRole === 'Warga' && status !== 'Published') {
        continue;
      }

      var matchSearch = !search || (judul.indexOf(search) !== -1 || note.indexOf(search) !== -1);
      if (matchSearch) {
        list.push({
          idMading: row[headers.indexOf('ID_Mading')],
          judul: row[headers.indexOf('Judul')],
          tanggalKegiatan: row[headers.indexOf('Tanggal_Kegiatan')],
          note: row[headers.indexOf('Note')],
          statusPublish: row[headers.indexOf('Status_Publish')],
          pembuat: row[headers.indexOf('Pembuat')],
          createdAt: row[headers.indexOf('Created_At')]
        });
      }
    }
    return { success: true, data: list.reverse() };
  } catch (error) {
    return { success: false, message: 'Gagal memuat mading: ' + error.toString() };
  }
}

function saveMading(payload) {
  try {
    var sheet = getSheet_('MADING');
    var data = sheet.getDataRange().getDisplayValues();
    var headers = data[0];
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var isEdit = payload.idMading && payload.idMading.trim() !== '';

    if (isEdit) {
      for (var i = 1; i < data.length; i++) {
        if (data[i][headers.indexOf('ID_Mading')] === payload.idMading) {
          var r = i + 1;
          sheet.getRange(r, headers.indexOf('Judul') + 1).setValue(payload.judul);
          sheet.getRange(r, headers.indexOf('Tanggal_Kegiatan') + 1).setValue(payload.tanggalKegiatan);
          sheet.getRange(r, headers.indexOf('Note') + 1).setValue(payload.note);
          sheet.getRange(r, headers.indexOf('Status_Publish') + 1).setValue(payload.statusPublish);
          SpreadsheetApp.flush();
          return { success: true, message: 'Informasi Mading berhasil diperbarui!' };
        }
      }
    }

    var newId = generateSequentialId_('MAD', 'MADING');
    sheet.appendRow([
      newId,
      String(payload.judul || '').trim(),
      String(payload.tanggalKegiatan || '').trim(),
      String(payload.note || '').trim(),
      String(payload.statusPublish || 'Published').trim(),
      String(payload.pembuat || 'Pengurus RT').trim(),
      nowStr
    ]);
    SpreadsheetApp.flush();
    return { success: true, message: 'Informasi Mading berhasil dipublikasikan!', idMading: newId };
  } catch (error) {
    return { success: false, message: 'Gagal menyimpan mading: ' + error.toString() };
  }
}

function deleteMading(idMading) {
  try {
    var sheet = getSheet_('MADING');
    var data = sheet.getDataRange().getDisplayValues();
    var idIdx = data[0].indexOf('ID_Mading');
    for (var i = 1; i < data.length; i++) {
      if (data[i][idIdx] === idMading) {
        sheet.deleteRow(i + 1);
        SpreadsheetApp.flush();
        return { success: true, message: 'Informasi Mading berhasil dihapus!' };
      }
    }
    return { success: false, message: 'ID Mading tidak ditemukan.' };
  } catch (error) {
    return { success: false, message: 'Gagal menghapus mading: ' + error.toString() };
  }
}
