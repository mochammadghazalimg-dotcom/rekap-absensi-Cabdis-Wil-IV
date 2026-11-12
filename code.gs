const SPREADSHEET_ID = '1vgTDc8Ut8mlwAn_U-OOrDpRVAS1HcBEnfmmAs3u8Vlw';
// ID SPREADSHEET KE 2 (DATABASE ARSIP ABSENSI)
const BACKUP_SPREADSHEET_ID = '1xRDKpYo-JeKKIXzC3s_rCRanv_9dPAqlu1npOVw3kRM'; 

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Sistem Presensi & Indisipliner - Cabdin Wil IV')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// --- AMBIL SEMUA DATA DARI GOOGLE SPREADSHEET ---
function ambilSemuaDataFromSheet() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    let sheetSekolah = ss.getSheetByName('Data_Sekolah') || ss.getSheetByName('Data Sekolah');
    let dataSekolah = [];
    if (sheetSekolah) {
      let raw = sheetSekolah.getDataRange().getValues();
      if (raw.length > 1) {
        raw.shift();
        dataSekolah = raw.filter(row => row[0] || row[1]).map(row => ({
          id: row[0] || Date.now() + Math.random(), 
          nama: String(row[1] || row[0] || '').trim(),
          npsn: String(row[2] || '').trim(), 
          kepsek: String(row[3] || '').trim(), 
          nip: String(row[4] || '').replace(/['\s]/g, '').trim()
        }));
      }
    }

    // Sinkronkan Rekap_Absen dari Log_Harian terlebih dahulu agar selalu akurat
    sinkronkanLogHarianKeRekapAbsen();

    let sheetRekap = ss.getSheetByName('Rekap_Absen') || ss.getSheetByName('Rekap Absensi');
    let dataAbsensi = [];
    if (sheetRekap) {
      let raw = sheetRekap.getDataRange().getValues();
      if (raw.length > 1) {
        raw.shift();
        dataAbsensi = raw.filter(row => row[0] || row[1]).map(row => ({
          nip: String(row[0] || '').replace(/['\s]/g, '').trim(), 
          nama: String(row[1] || '').trim(), 
          sekolah: String(row[2] || '').trim(),
          hariKerja: parseInt(row[3]) || 0, 
          hadir: parseInt(row[4]) || 0, 
          terlambatMin: parseInt(row[5]) || 0,
          cepatPulangMin: parseInt(row[6]) || 0, 
          izin: parseInt(row[7]) || 0,
          cuti: parseInt(row[8]) || 0,
          sakit: parseInt(row[9]) || 0,
          dinasLuar: parseInt(row[10]) || 0,
          tanpaKet: parseInt(row[11]) || 0
        }));
      }
    }

    let sheetPegawai = ss.getSheetByName('Data_Pegawai') || ss.getSheetByName('Data Pegawai');
    let dataPegawai = [];
    if (sheetPegawai) {
      let raw = sheetPegawai.getDataRange().getValues();
      if (raw.length > 1) {
        raw.shift();
        dataPegawai = raw.filter(row => row[0] || row[1]).map(row => ({
          nip: String(row[0] || '').replace(/['\s]/g, '').trim(), 
          nama: String(row[1] || '').trim(),
          sekolah: String(row[2] || '').trim(), 
          photo_uploaded: String(row[3] || 'False').trim()
        }));
      }
    }

    let sheetLog = ss.getSheetByName('Log_Harian');
    let dataLogHarian = [];
    if (sheetLog) {
      let raw = sheetLog.getDataRange().getValues();
      if (raw.length > 1) {
        raw.shift();
        dataLogHarian = raw.filter(row => row[0]).map(row => {
          let yr = parseInt(row[3]) || 0;
          if (yr > 0 && yr < 100) yr += 2000;
          let mm = parseInt(row[4]) || 0;
          let dd = parseInt(row[5]) || 0;
          
          let mmStr = String(mm).padStart(2, '0');
          let ddStr = String(dd).padStart(2, '0');
          let tglStr = `${yr}-${mmStr}-${ddStr}`;

          return {
            nip: String(row[0] || '').replace(/['\s]/g, '').trim(), 
            nama: String(row[1] || '').trim(), 
            sekolah: String(row[2] || '').trim(),
            tanggalStr: tglStr, 
            yyyy: yr, 
            mm: mm, 
            dd: dd,
            hadir: parseInt(row[6]) || 0, 
            terlambatMin: parseInt(row[7]) || 0, 
            cepatPulangMin: parseInt(row[8]) || 0, 
            izin: parseInt(row[9]) || 0,
            cuti: parseInt(row[10]) || 0,
            sakit: parseInt(row[11]) || 0,
            dinasLuar: parseInt(row[12]) || 0
          };
        });
      }
    }

    // MEMUAT ARSIP DASHBOARD (Data Snapshot)
    let sheetArsip = ss.getSheetByName('Arsip_Dashboard');
    let dataArsip = [];
    if (sheetArsip) {
      let raw = sheetArsip.getDataRange().getValues();
      if (raw.length > 1) {
        raw.shift();
        dataArsip = raw.filter(row => row[0]).map(row => ({
          id_periode: String(row[0] || ''), 
          bulan: parseInt(row[1]) || 0, 
          tahun: parseInt(row[2]) || 0, 
          snapshot: String(row[3] || '')
        }));
      }
    }

    return { sekolah: dataSekolah, absensi: dataAbsensi, pegawai: dataPegawai, logHarian: dataLogHarian, arsip: dataArsip };
  } catch(err) {
    return { error: err.message, sekolah: [], absensi: [], pegawai: [], logHarian: [], arsip: [] };
  }
}

// --- FUNGSI PROSES BACKUP DAN RESET BULANAN KETAT DENGAN URUTAN ALUR PERSIS ---
function prosesBackupDanResetBulanan(snapshotJSON, targetBulan, targetTahun) {
  try {
    const ssSource = SpreadsheetApp.openById(SPREADSHEET_ID);
    const ssBackup = SpreadsheetApp.openById(BACKUP_SPREADSHEET_ID);
    
    let sheetLog = ssSource.getSheetByName('Log_Harian');
    let sheetRekap = ssSource.getSheetByName('Rekap_Absen') || ssSource.getSheetByName('Rekap Absensi');
    
    if (!sheetLog || !sheetRekap) return { status: 'error', message: 'Sheet sumber tidak ditemukan.' };
    
    let dataLog = sheetLog.getDataRange().getValues();
    let dataRekap = sheetRekap.getDataRange().getValues();
    
    let now = new Date();
    if (!targetBulan) { 
      targetBulan = now.getMonth(); 
      targetTahun = now.getFullYear(); 
      if (targetBulan === 0) { targetBulan = 12; targetTahun -= 1; } 
    }
    
    let namaBulanBck = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Ags", "Sep", "Okt", "Nov", "Des"];
    let labelWaktu = `${namaBulanBck[targetBulan - 1]}_${targetTahun}`; 
    let namaSheetLogBaru = `Log_Harian_${labelWaktu}`;
    let namaSheetRekapBaru = `Rekap_Absen_${labelWaktu}`;
    
    // Cek apakah data bulan ini sudah pernah dibackup ke Spreadsheet 2
    if (ssBackup.getSheetByName(namaSheetLogBaru) || ssBackup.getSheetByName(namaSheetRekapBaru)) {
      return { status: 'already_backed_up', message: 'Data bulan ini sudah pernah dibackup sebelumnya.' };
    }

    if (dataLog.length <= 1 && dataRekap.length <= 1) {
      return { status: 'no_data', message: 'Data absensi masih kosong.' };
    }

    // ALUR A: BACKUP DATA MENTAH KE SPREADSHEET 2 (Database Arsip Absensi)
    let backupLogSheet = ssBackup.insertSheet(namaSheetLogBaru);
    let backupRekapSheet = ssBackup.insertSheet(namaSheetRekapBaru);
    
    if (dataLog.length > 0) backupLogSheet.getRange(1, 1, dataLog.length, dataLog[0].length).setValues(dataLog);
    if (dataRekap.length > 0) backupRekapSheet.getRange(1, 1, dataRekap.length, dataRekap[0].length).setValues(dataRekap);
    
    SpreadsheetApp.flush(); 
    
    // ALUR B: PEMERIKSAAN KETAT KELENGKAPAN DATA (Double Check)
    let verifLog = backupLogSheet.getDataRange().getValues();
    let verifRekap = backupRekapSheet.getDataRange().getValues();
    
    let isLogRowsValid = (verifLog.length === dataLog.length) && (verifLog[0].length === dataLog[0].length);
    let isRekapRowsValid = (verifRekap.length === dataRekap.length) && (verifRekap[0].length === dataRekap[0].length);
    
    let lastRowLogIdx = dataLog.length - 1;
    let isLogContentValid = (String(verifLog[lastRowLogIdx][0]) === String(dataLog[lastRowLogIdx][0]));
    let lastRowRekapIdx = dataRekap.length - 1;
    let isRekapContentValid = (String(verifRekap[lastRowRekapIdx][0]) === String(dataRekap[lastRowRekapIdx][0]));
    
    let isBackupValid = isLogRowsValid && isRekapRowsValid && isLogContentValid && isRekapContentValid;

    if (!isBackupValid) {
      // Jika data ada yang tertinggal/tidak cocok, hapus sheet backup (rollback)
      ssBackup.deleteSheet(backupLogSheet);
      ssBackup.deleteSheet(backupRekapSheet);
      return { status: 'fail', message: 'Verifikasi backup gagal. Data tidak cocok.' };
    }

    // ALUR C: JIKA BACKUP BENAR LENGKAP -> LAKUKAN SNAPSHOT DASHBOARD KE SHEET Arsip_Dashboard
    let snapshotBerhasil = false;
    if (snapshotJSON) {
      let sheetArsip = ssSource.getSheetByName('Arsip_Dashboard');
      if (!sheetArsip) {
        sheetArsip = ssSource.insertSheet('Arsip_Dashboard');
        sheetArsip.appendRow(['ID_Periode', 'Bulan', 'Tahun', 'Data_Snapshot']);
      }
      let id_periode = `${targetBulan}_${targetTahun}`;
      let dataA = sheetArsip.getDataRange().getValues();
      let rowIndex = -1;
      for (let i = 1; i < dataA.length; i++) {
        if (String(dataA[i][0]) === String(id_periode)) { rowIndex = i + 1; break; }
      }
      if (rowIndex > -1) {
        sheetArsip.getRange(rowIndex, 2, 1, 3).setValues([[targetBulan, targetTahun, snapshotJSON]]);
      } else {
        sheetArsip.appendRow([id_periode, targetBulan, targetTahun, snapshotJSON]);
      }
      SpreadsheetApp.flush();

      // ALUR D: PERIKSA SHEET Arsip_Dashboard JIKA BENAR SUDAH TERSIMPAN DAN TERKUNCI
      let verifArsip = sheetArsip.getDataRange().getValues();
      snapshotBerhasil = verifArsip.some(row => String(row[0]) === String(id_periode) && row[3] !== '');
    }

    // ALUR E: JIKA SNAPSHOT SUDAH BENAR TERSIMPAN & TERKUNCI -> BARU HAPUS DATA MENTAH DIBULAN LALU
    if (snapshotBerhasil) {
      if (dataLog.length > 1) sheetLog.getRange(2, 1, dataLog.length - 1, dataLog[0].length).clearContent();
      if (dataRekap.length > 1) sheetRekap.getRange(2, 1, dataRekap.length - 1, dataRekap[0].length).clearContent();
      SpreadsheetApp.flush();
      return { status: 'success', message: 'Backup & Snapshot berhasil dikunci. Data berhasil di-reset.' };
    } else {
      return { status: 'fail', message: 'Penyimpanan Snapshot ke Arsip_Dashboard gagal. Reset dibatalkan.' };
    }
  } catch(e) {
    return { status: 'fail', error: e.message };
  }
}
// ==========================================================
// FUNGSI SYSTEM LOGIN ADMIN (DATABASE SPREADSHEET)
// ==========================================================
function prosesLoginAdmin(username, password) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheetAdmin = ss.getSheetByName('Data_Admin') || ss.getSheetByName('Data Admin');
    
    // 1. Jika sheet belum ada, buat otomatis & masukkan akun default
    if (!sheetAdmin) {
      sheetAdmin = ss.insertSheet('Data_Admin');
      sheetAdmin.appendRow(['Username', 'Password', 'Nama', 'Role']);
      sheetAdmin.appendRow(['admin', 'admin123', 'Admin Cabdin', 'Super Admin']);
      SpreadsheetApp.flush();
    }
    
    let raw = sheetAdmin.getDataRange().getValues();
    if (raw.length <= 1) {
      sheetAdmin.appendRow(['admin', 'admin123', 'Admin Cabdin', 'Super Admin']);
      raw = sheetAdmin.getDataRange().getValues();
    }
    
    // 2. Cocokkan Username & Password
    let userClean = String(username || '').trim().toLowerCase();
    let passClean = String(password || '').trim();

    for (let i = 1; i < raw.length; i++) {
      let u = String(raw[i][0] || '').trim().toLowerCase();
      let p = String(raw[i][1] || '').trim();
      let nama = String(raw[i][2] || 'Admin Cabdin').trim();
      let role = String(raw[i][3] || 'Administrator').trim();
      
      if (u === userClean && p === passClean) {
        return {
          status: 'success',
          user: { username: raw[i][0], nama: nama, role: role }
        };
      }
    }
    
    return { status: 'fail', message: 'Username atau Kata Sandi salah!' };
  } catch(e) {
    if (String(username).trim() === 'admin' && String(password).trim() === 'admin123') {
      return { status: 'success', user: { username: 'admin', nama: 'Admin Cabdin', role: 'Super Admin' } };
    }
    return { status: 'fail', message: 'Terjadi kesalahan sistem: ' + e.message };
  }
}

function simpanDataPegawaiKeSheet(listPegawai) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Data_Pegawai') || ss.getSheetByName('Data Pegawai') || ss.insertSheet('Data_Pegawai');
    sheet.clearContents();
    let rows = [['NIP', 'Nama_Pegawai', 'Sekolah', 'Foto_Acuan']];
    if (listPegawai && listPegawai.length > 0) {
      listPegawai.forEach(p => {
        let cleanNip = String(p.nip || '').replace(/['\s]/g, '');
        rows.push(["'" + cleanNip, p.nama, p.sekolah, p.photo_uploaded || 'False']);
      });
    }
    sheet.getRange(1, 1, rows.length, 4).setValues(rows);
    return "OK";
  } catch(e) { return "Error: " + e.message; }
}

function simpanDataSekolahKeSheet(listSekolah) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Data_Sekolah') || ss.getSheetByName('Data Sekolah') || ss.insertSheet('Data_Sekolah');
    sheet.clearContents();
    let rows = [['ID', 'Nama_Sekolah', 'NPSN', 'Nama_Kepsek', 'NIP_Kepsek']];
    if (listSekolah && listSekolah.length > 0) listSekolah.forEach(s => rows.push([s.id, s.nama, s.npsn, s.kepsek, s.nip]));
    sheet.getRange(1, 1, rows.length, 5).setValues(rows);
    return "OK";
  } catch(e) { return "Error: " + e.message; }
}

function simpanRekapAbsenKeSheet(listAbsensi) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Rekap_Absen') || ss.getSheetByName('Rekap Absensi') || ss.insertSheet('Rekap_Absen');
    sheet.clearContents();
    let rows = [['NIP', 'Nama_Pegawai', 'Sekolah', 'Hari_Kerja', 'Hadir', 'Keterlambatan', 'Cepat_Pulang', 'Izin', 'Cuti', 'Sakit', 'Dinas_Luar', 'Tanpa_Ket']];
    if (listAbsensi && listAbsensi.length > 0) {
      listAbsensi.forEach(a => {
        let cleanNip = String(a.nip || '').replace(/['\s]/g, '');
        rows.push([
          "'" + cleanNip, 
          a.nama, 
          a.sekolah, 
          a.hariKerja || 0, 
          a.hadir || 0, 
          a.terlambatMin || 0, 
          a.cepatPulangMin || 0, 
          a.izin || 0,
          a.cuti || 0,
          a.sakit || 0,
          a.dinasLuar || 0,
          a.tanpaKet || 0
        ]);
      });
    }
    sheet.getRange(1, 1, rows.length, 12).setValues(rows);
    return "OK";
  } catch(e) { return "Error: " + e.message; }
}

function simpanLogHarianKeSheet(listLog) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Log_Harian') || ss.insertSheet('Log_Harian');
    sheet.clearContents();
    
    let rows = [['NIP', 'Nama_Pegawai', 'Sekolah', 'Tahun', 'Bulan', 'Hari', 'Hadir', 'Keterlambatan', 'Cepat_Pulang', 'Izin', 'Cuti', 'Sakit', 'Dinas_Luar']];
    
    if (listLog && listLog.length > 0) {
      listLog.forEach(l => {
        let cleanNip = String(l.nip || '').replace(/['\s]/g, '');
        let yr = parseInt(l.yyyy) || 0;
        if (yr > 0 && yr < 100) yr += 2000;
        
        rows.push([
          "'" + cleanNip, 
          l.nama, 
          l.sekolah, 
          yr, 
          l.mm, 
          l.dd, 
          l.hadir || 0, 
          l.terlambatMin || 0, 
          l.cepatPulangMin || 0, 
          l.izin || 0,
          l.cuti || 0,
          l.sakit || 0,
          l.dinasLuar || 0
        ]);
      });
    }
    
    sheet.getRange(1, 1, rows.length, 13).setValues(rows);
    return "OK";
  } catch(e) { return "Error: " + e.message; }
}

function sinkronkanLogHarianKeRekapAbsen(targetBulan, targetTahun) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return null;

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetPegawai = ss.getSheetByName('Data_Pegawai') || ss.getSheetByName('Data Pegawai');
    const sheetLog = ss.getSheetByName('Log_Harian');
    let sheetRekap = ss.getSheetByName('Rekap_Absen') || ss.getSheetByName('Rekap Absensi');
    if (!sheetRekap) sheetRekap = ss.insertSheet('Rekap_Absen');
    if (!sheetPegawai || !sheetLog) return null;

    const now = new Date();
    const bulan = targetBulan || (now.getMonth() + 1);
    const tahun = targetTahun || now.getFullYear();
    const rawLog = sheetLog.getDataRange().getValues();
    if (rawLog.length <= 1) return null;

    const logMap = {};
    let adaLogBulanIni = false;

    for (let i = 1; i < rawLog.length; i++) {
      const row = rawLog[i];
      const nip = String(row[0] || '').replace(/['\s]/g, '').trim();
      if (!nip) continue;
      let yr = parseInt(row[3]) || 0;
      if (yr > 0 && yr < 100) yr += 2000;
      const mm = parseInt(row[4]) || 0;

      if (yr === tahun && mm === bulan) {
        adaLogBulanIni = true;
        if (!logMap[nip]) logMap[nip] = { hadir: 0, terlambat: 0, cepatPulang: 0, izin: 0, cuti: 0, sakit: 0, dinasLuar: 0 };
        logMap[nip].hadir += (parseInt(row[6]) || 0);
        logMap[nip].terlambat += (parseInt(row[7]) || 0);
        logMap[nip].cepatPulang += (parseInt(row[8]) || 0);
        logMap[nip].izin += (parseInt(row[9]) || 0);
        logMap[nip].cuti += (parseInt(row[10]) || 0);
        logMap[nip].sakit += (parseInt(row[11]) || 0);
        logMap[nip].dinasLuar += (parseInt(row[12]) || 0);
      }
    }
    if (!adaLogBulanIni) return null;

    const liburList = getLiburOtomatisDariGoogle(tahun);
    const isBulanBerjalan = (bulan === (now.getMonth() + 1) && tahun === now.getFullYear());
    const daysInMonth = new Date(tahun, bulan, 0).getDate();
    const maxDay = isBulanBerjalan ? Math.min(now.getDate(), daysInMonth) : daysInMonth;

    let hariKerja = 0;
    for (let d = 1; d <= maxDay; d++) {
      const dt = new Date(tahun, bulan - 1, d);
      const dateStr = tahun + '-' + String(bulan).padStart(2, '0') + '-' + String(d).padStart(2, '0');
      if (dt.getDay() !== 0 && dt.getDay() !== 6 && liburList.indexOf(dateStr) === -1) hariKerja++;
    }

    const rawPegawai = sheetPegawai.getDataRange().getValues();
    const rowsRekap = [['NIP', 'Nama_Pegawai', 'Sekolah', 'Hari_Kerja', 'Hadir', 'Keterlambatan', 'Cepat_Pulang', 'Izin', 'Cuti', 'Sakit', 'Dinas_Luar', 'Tanpa_Ket']];

    for (let i = 1; i < rawPegawai.length; i++) {
      const pRow = rawPegawai[i];
      const nip = String(pRow[0] || '').replace(/['\s]/g, '').trim();
      const nama = String(pRow[1] || '').trim();
      const sekolah = String(pRow[2] || '').trim();
      if (!nip || !nama) continue;

      const stat = logMap[nip] || { hadir: 0, terlambat: 0, cepatPulang: 0, izin: 0, cuti: 0, sakit: 0, dinasLuar: 0 };
      const tanpaKet = Math.max(0, hariKerja - (stat.hadir + stat.izin + stat.cuti + stat.sakit + stat.dinasLuar));
      rowsRekap.push(["'" + nip, nama, sekolah, hariKerja, stat.hadir, stat.terlambat, stat.cepatPulang, stat.izin, stat.cuti, stat.sakit, stat.dinasLuar, tanpaKet]);
    }

    sheetRekap.clearContents();
    sheetRekap.getRange(1, 1, rowsRekap.length, 12).setValues(rowsRekap);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}

function autoSyncRekapAbsenTrigger(e) {
  sinkronkanLogHarianKeRekapAbsen();
}

function pasangTriggerOtomatis() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'autoSyncRekapAbsenTrigger') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('autoSyncRekapAbsenTrigger').forSpreadsheet(ss).onChange().create();
}

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
  } catch (e) { return []; }
}
