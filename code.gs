const SPREADSHEET_ID = '1vgTDc8Ut8mlwAn_U-OOrDpRVAS1HcBEnfmmAs3u8Vlw';

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Sistem Presensi & Indisipliner - Cabdin Wil IV')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// --- AMBIL SEMUA DATA DARI GOOGLE SPREADSHEET (SEKOLAH, ABSENSI, PEGAWAI) ---
function ambilSemuaDataFromSheet() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // 1. Data Sekolah (Tab: Data_Sekolah)
    let sheetSekolah = ss.getSheetByName('Data_Sekolah') || ss.getSheetByName('Data Sekolah');
    let dataSekolah = [];
    if (sheetSekolah) {
      let raw = sheetSekolah.getDataRange().getValues();
      if (raw.length > 1) {
        raw.shift(); // Hapus Header
        dataSekolah = raw.filter(row => row[0] || row[1]).map(row => ({
          id: row[0] || Date.now() + Math.random(),
          nama: String(row[1] || row[0] || '').trim(),
          npsn: String(row[2] || '').trim(),
          kepsek: String(row[3] || '').trim(),
          nip: String(row[4] || '').trim()
        }));
      }
    }

    // 2. Data Rekap Absensi (Tab: Rekap_Absen)
    let sheetRekap = ss.getSheetByName('Rekap_Absen') || ss.getSheetByName('Rekap Absensi');
    let dataAbsensi = [];
    if (sheetRekap) {
      let raw = sheetRekap.getDataRange().getValues();
      if (raw.length > 1) {
        raw.shift(); // Hapus Header
        dataAbsensi = raw.filter(row => row[0] || row[1]).map(row => ({
          nip: String(row[0] || '').trim(),
          nama: String(row[1] || '').trim(),
          sekolah: String(row[2] || '').trim(),
          hariKerja: parseInt(row[3]) || 0,
          hadir: parseInt(row[4]) || 0,
          terlambatMin: parseInt(row[5]) || 0,
          cepatPulangMin: parseInt(row[6]) || 0,
          izinSakit: parseInt(row[7]) || 0,
          tanpaKet: parseInt(row[8]) || 0
        }));
      }
    }

    // 3. Master Data Pegawai ASN (Tab: Data_Pegawai)
    let sheetPegawai = ss.getSheetByName('Data_Pegawai') || ss.getSheetByName('Data Pegawai');
    let dataPegawai = [];
    if (sheetPegawai) {
      let raw = sheetPegawai.getDataRange().getValues();
      if (raw.length > 1) {
        raw.shift(); // Hapus Header
        dataPegawai = raw.filter(row => row[0] || row[1]).map(row => ({
          nip: String(row[0] || '').trim(),
          nama: String(row[1] || '').trim(),
          sekolah: String(row[2] || '').trim(),
          photo_uploaded: String(row[3] || 'False').trim()
        }));
      }
    }

  // 4. Data Log Harian (Tab: Log_Harian)
    let sheetLog = ss.getSheetByName('Log_Harian');
    let dataLogHarian = [];
    if (sheetLog) {
      let raw = sheetLog.getDataRange().getValues();
      if (raw.length > 1) {
        raw.shift(); // Hapus Header
        dataLogHarian = raw.filter(row => row[0]).map(row => ({
          nip: String(row[0] || '').trim(),
          nama: String(row[1] || '').trim(),
          sekolah: String(row[2] || '').trim(),
          tanggalStr: String(row[3] || '').trim(),
          yyyy: parseInt(row[4]) || 0,
          mm: parseInt(row[5]) || 0,
          dd: parseInt(row[6]) || 0,
          hadir: parseInt(row[7]) || 0,
          terlambatMin: parseInt(row[8]) || 0,
          cepatPulangMin: parseInt(row[9]) || 0,
          izinSakit: parseInt(row[10]) || 0
        }));
      }
    }

    return { sekolah: dataSekolah, absensi: dataAbsensi, pegawai: dataPegawai, logHarian: dataLogHarian };
  } catch(err) {
    return { error: err.message, sekolah: [], absensi: [], pegawai: [] };
  }
}

// --- SIMPAN MASTER DATA PEGAWAI KE SPREADSHEET ---
function simpanDataPegawaiKeSheet(listPegawai) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Data_Pegawai') || ss.getSheetByName('Data Pegawai') || ss.insertSheet('Data_Pegawai');
    sheet.clearContents();
    let rows = [['NIP', 'Nama_Pegawai', 'Sekolah', 'Foto_Acuan']];
    if (listPegawai && listPegawai.length > 0) {
      listPegawai.forEach(p => rows.push([p.nip, p.nama, p.sekolah, p.photo_uploaded || 'False']));
    }
    sheet.getRange(1, 1, rows.length, 4).setValues(rows);
    return "OK";
  } catch(e) {
    return "Error: " + e.message;
  }
}

// --- SIMPAN DATA SEKOLAH KE SPREADSHEET ---
function simpanDataSekolahKeSheet(listSekolah) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Data_Sekolah') || ss.getSheetByName('Data Sekolah') || ss.insertSheet('Data_Sekolah');
    sheet.clearContents();
    let rows = [['ID', 'Nama_Sekolah', 'NPSN', 'Nama_Kepsek', 'NIP_Kepsek']];
    if (listSekolah && listSekolah.length > 0) {
      listSekolah.forEach(s => rows.push([s.id, s.nama, s.npsn, s.kepsek, s.nip]));
    }
    sheet.getRange(1, 1, rows.length, 5).setValues(rows);
    return "OK";
  } catch(e) {
    return "Error: " + e.message;
  }
}

// --- SIMPAN REKAP ABSEN KE SPREADSHEET ---
function simpanRekapAbsenKeSheet(listAbsensi) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Rekap_Absen') || ss.getSheetByName('Rekap Absensi') || ss.insertSheet('Rekap_Absen');
    sheet.clearContents();
    let rows = [['NIP', 'Nama_Pegawai', 'Sekolah', 'Hari_Kerja', 'Hadir', 'Keterlambatan', 'Cepat_Pulang', 'Izin_Sakit_DL', 'Tanpa_Ket']];
    if (listAbsensi && listAbsensi.length > 0) {
      listAbsensi.forEach(a => rows.push([a.nip, a.nama, a.sekolah, a.hariKerja || 0, a.hadir, a.terlambatMin, a.cepatPulangMin, a.izinSakit, a.tanpaKet]));
    }
    sheet.getRange(1, 1, rows.length, 9).setValues(rows);
    return "OK";
  } catch(e) {
    return "Error: " + e.message;
  }
}

// --- SIMPAN LOG HARIAN KE SPREADSHEET ---
function simpanLogHarianKeSheet(listLog) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Log_Harian') || ss.insertSheet('Log_Harian');
    sheet.clearContents();
    let rows = [['NIP', 'Nama_Pegawai', 'Sekolah', 'Tanggal', 'Tahun', 'Bulan', 'Hari', 'Hadir', 'Keterlambatan', 'Cepat_Pulang', 'Izin_Sakit_DL']];
    if (listLog && listLog.length > 0) {
      listLog.forEach(l => rows.push([l.nip, l.nama, l.sekolah, l.tanggalStr, l.yyyy, l.mm, l.dd, l.hadir, l.terlambatMin, l.cepatPulangMin, l.izinSakit]));
    }
    sheet.getRange(1, 1, rows.length, 11).setValues(rows);
    return "OK";
  } catch(e) {
    return "Error: " + e.message;
  }
}

// --- FUNGSI KALENDER LIBUR ---
function getLiburOtomatisDariGoogle(tahun) {
  try {
    var cal = CalendarApp.getCalendarById('id.indonesian#holiday@group.v.calendar.google.com');
    if (!cal) return []; 
    var events = cal.getEvents(new Date(tahun, 0, 1), new Date(tahun, 11, 31));
    var liburArray = events.map(e => {
      let d = e.getStartTime();
      return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
    });
    return [...new Set(liburArray)];
  } catch (e) { 
    return []; 
  }
}
