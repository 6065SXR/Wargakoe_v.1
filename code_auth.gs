/**
 * ====================================================================
 * WARGAKOE - USER AUTHENTICATION & REGISTRATION MODULE
 * By Zettbos System (ZettBOT 3.1)
 * ====================================================================
 */

function loginUser(noHp, password) {
  try {
    if (!noHp || !password) {
      return { success: false, message: 'Nomor Handphone dan Password wajib diisi!' };
    }

    var sheet = getSheet_('USERS');
    var data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) {
      return { success: false, message: 'Data pengguna tidak ditemukan. Silakan registrasi terlebih dahulu.' };
    }

    var headers = data[0];
    var hpIdx = headers.indexOf('No_HP');
    var passIdx = headers.indexOf('Password');

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[hpIdx] === noHp.trim() && row[passIdx] === password.trim()) {
        var userObj = {
          noKk: row[headers.indexOf('No_KK')],
          noKtp: row[headers.indexOf('No_KTP')],
          nama: row[headers.indexOf('Nama')],
          noHp: row[hpIdx],
          role: row[headers.indexOf('Role')],
          statusUser: row[headers.indexOf('Status_User')],
          tanggalLahir: row[headers.indexOf('Tanggal_Lahir')],
          umur: row[headers.indexOf('Umur')],
          alamat: row[headers.indexOf('Alamat')],
          jenisKelamin: row[headers.indexOf('Jenis_Kelamin')],
          statusKeluarga: row[headers.indexOf('Status_Keluarga')],
          statusRumah: row[headers.indexOf('Status_Rumah')],
          pendidikan: row[headers.indexOf('Pendidikan')],
          pekerjaan: row[headers.indexOf('Pekerjaan')]
        };
        return { success: true, message: 'Login berhasil! Selamat datang, ' + userObj.nama, user: userObj };
      }
    }

    return { success: false, message: 'Nomor HP atau Password salah!' };
  } catch (error) {
    return { success: false, message: 'Gagal login: ' + error.toString() };
  }
}

function registerUser(formData) {
  try {
    var sheet = getSheet_('USERS');
    var data = sheet.getDataRange().getDisplayValues();
    var headers = data[0];

    var kkIdx = headers.indexOf('No_KK');
    var ktpIdx = headers.indexOf('No_KTP');
    var hpIdx = headers.indexOf('No_HP');

    for (var i = 1; i < data.length; i++) {
      if (data[i][kkIdx] === String(formData.noKk).trim()) {
        return { success: false, message: 'Nomor KK sudah terdaftar di sistem!' };
      }
      if (data[i][ktpIdx] === String(formData.noKtp).trim()) {
        return { success: false, message: 'Nomor KTP sudah terdaftar di sistem!' };
      }
      if (data[i][hpIdx] === String(formData.noHp).trim()) {
        return { success: false, message: 'Nomor HP sudah terdaftar. Silakan login!' };
      }
    }

    var nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var defaultRole = formData.statusUser === 'Pengurus RT' ? (formData.role || 'Pengurus RT') : 'Warga';

    var newRow = [
      String(formData.noKk || '').trim(),
      String(formData.noKtp || '').trim(),
      String(formData.nama || '').trim(),
      String(formData.noHp || '').trim(),
      String(formData.password || '').trim(),
      defaultRole,
      String(formData.statusUser || 'Warga').trim(),
      String(formData.tanggalLahir || '').trim(),
      String(formData.umur || '0').trim(),
      String(formData.alamat || '').trim(),
      String(formData.jenisKelamin || 'Laki-laki').trim(),
      String(formData.statusKeluarga || 'Suami').trim(),
      String(formData.statusRumah || 'Pribadi').trim(),
      String(formData.pendidikan || 'SMA').trim(),
      String(formData.pekerjaan || 'Wiraswasta').trim(),
      nowStr
    ];

    sheet.appendRow(newRow);
    SpreadsheetApp.flush();

    return { success: true, message: 'Registrasi berhasil! Silakan login.' };
  } catch (error) {
    return { success: false, message: 'Gagal registrasi: ' + error.toString() };
  }
}
