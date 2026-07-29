// ============================================================
// CODE.GS GABUNGAN — TBR REGION 10
// Berisi 2 bagian:
//   BAGIAN A: Backend Web App (doGet) — dipanggil dari GitHub Pages
//             (index.html & laporan.html) untuk submit & ambil laporan.
//   BAGIAN B: Rekap Cabang — menu sidebar di Google Sheets untuk
//             memantau cabang yang sudah/belum lapor.
// Cara pasang: ganti SEMUA isi Code.gs di Apps Script dengan file ini,
// lalu Deploy ulang Web App (pilih "New version" di deployment yang sudah ada,
// supaya URL /exec tidak berubah).
// ============================================================


// ============================================================
// BAGIAN A — BACKEND WEB APP (doGet)
// ============================================================

var CONFIG = {
  NAMA_SHEET: "Laporan Harian",         // Nama tab sheet di Google Spreadsheet
  NAMA_SHEET_LOG: "Log Laporan",        // Nama tab log (dibuat otomatis)
  NAMA_SHEET_LIVIN: "Monitoring Livin' Food",  // Nama tab monitoring Livin' Food
  NAMA_SHEET_PIPELINE: "Pipeline EDC to LVM"   // Nama tab pipeline konversi EDC to LVM
};

// Tahapan status progress pipeline akuisisi merchant yang valid.
var STATUS_EDC_LVM_VALID = ["Target", "Penawaran", "Done Konversi to LVM", "Merchant Menolak", "Merchant Tutup", "Perlu Kunjungan MTI"];

/**
 * Semua request dari GitHub Pages masuk ke sini via GET.
 * Parameter "action" menentukan apa yang dilakukan:
 *   - action=submit  → simpan data laporan baru
 *   - action=getData → ambil data laporan berdasarkan ID
 */
function doGet(e) {
  var action = (e.parameter && e.parameter.action) ? e.parameter.action : '';

  // ── SUBMIT: simpan laporan baru ──────────────────────────────
  if (action === 'submit') {
    try {
      var result = prosesSubmit(e.parameter);
      return jsonResponse({ success: true, idLaporan: result.idLaporan });
    } catch (err) {
      return jsonResponse({ success: false, error: err.message });
    }
  }

  // ── GET DATA: ambil data laporan by ID ───────────────────────
  if (action === 'getData') {
    var id = e.parameter.id || '';
    var laporan = cariLaporan(id);
    if (laporan) {
      return jsonResponse(laporan);
    } else {
      return jsonResponse({ error: 'Laporan tidak ditemukan: ' + id });
    }
  }

  // ── GET MERCHANT EDC: ambil daftar merchant EDC (pipeline to LVM)
  //    milik satu cabang, lengkap dengan status progress terkininya ──
  if (action === 'getMerchantEDC') {
    var kodeCabangQ = e.parameter.kodeCabang || '';
    try {
      var merchants = getMerchantEDCByCabang(kodeCabangQ);
      return jsonResponse({ merchants: merchants });
    } catch (err) {
      return jsonResponse({ merchants: [], error: err.message });
    }
  }

  // ── GET PIPELINE ALL: ambil semua data pipeline EDC to LVM,
  //    opsional difilter per kodeCabang (dipakai pipeline.html) ──
  if (action === 'getPipelineAll') {
    var kodeFilter = e.parameter.kodeCabang || '';
    try {
      var pipelineList = getPipelineAll(kodeFilter);
      return jsonResponse({ data: pipelineList });
    } catch (err) {
      return jsonResponse({ data: [], error: err.message });
    }
  }

  // ── UPDATE STATUS: update status progress satu merchant pipeline,
  //    dipanggil langsung dari pipeline.html saat admin klik merchant
  //    lalu pilih status baru di modal ──
  if (action === 'updateStatus') {
    try {
      var ssUpd = SpreadsheetApp.getActiveSpreadsheet();
      var kodeCabangUpd   = e.parameter.kodeCabang   || '';
      var namaMerchantUpd = e.parameter.namaMerchant || '';
      var statusBaruUpd   = e.parameter.status       || '';
      var midUpd          = e.parameter.mid          || '';
      var alasanUpd       = e.parameter.alasan       || '';

      if (statusBaruUpd === 'Done Konversi to LVM' && !midUpd.trim()) {
        return jsonResponse({ success: false, error: 'MNDI/MID/Nomor Rekening wajib diisi untuk status Done Konversi to LVM.' });
      }
      if (statusBaruUpd === 'Merchant Menolak' && !alasanUpd.trim()) {
        return jsonResponse({ success: false, error: 'Alasan wajib diisi untuk status Merchant Menolak.' });
      }

      var berhasilUpd = updatePipelineStatus(ssUpd, kodeCabangUpd, namaMerchantUpd, statusBaruUpd, midUpd, alasanUpd);
      if (berhasilUpd) {
        return jsonResponse({ success: true });
      } else {
        return jsonResponse({ success: false, error: 'Merchant tidak ditemukan atau status tidak valid.' });
      }
    } catch (err) {
      return jsonResponse({ success: false, error: err.message });
    }
  }

  // ── Default: API aktif ───────────────────────────────────────
  return jsonResponse({ status: 'API aktif', versi: '2.4' });
}

// Helper: buat JSON response
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// PROSES SUBMIT: simpan data ke Google Sheet
// ============================================================
function prosesSubmit(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(ss);

  // Ambil & validasi data dari form
  var tanggal          = p.tanggal          || '';
  var area             = p.area             || '';
  var kodeCabang       = p.kodeCabang       || '';
  var namaCabang       = p.namaCabang       || '';
  var jumlahLVM        = Number(p.jumlahLVM)            || 0;
  var jumlahEDC        = Number(p.jumlahEDC)            || 0;
  var jumlahEDCPOT     = Number(p.jumlahEDCPOT)         || 0;
  var jumlahPemasLVM   = Number(p.jumlahPemasanganLVM)  || 0;
  var jumlahRetensiEDC = Number(p.jumlahRetensiEDC)     || 0;
  var totalLeadsCakra     = Number(p.totalLeadsCakra)     || 0;
  var totalKunjunganCakra = Number(p.totalKunjunganCakra) || 0;
  var gapCakra             = Number(p.gapCakra)            || 0;
  var jumlahLivinFood      = Number(p.jumlahTransaksiLivinFood) || 0;
  var namaMerchantLivinFood = p.namaMerchantLivinFood || '-';
  var kendala          = p.kendala          || '-';
  var keterangan       = p.keterangan       || '-';
  var totalAkuisisi    = jumlahLVM + jumlahEDC + jumlahEDCPOT;
  var tanggalFormatted = formatTanggal(tanggal);
  var timestamp        = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

  // Ambil OTOMATIS semua merchant Pipeline EDC to LVM milik cabang ini yang
  // statusnya diupdate hari yang sama dengan tanggal laporan (lewat pipeline.html) —
  // jadi tidak dibatasi cuma 1 merchant, semua kunjungan hari itu ikut kebawa.
  var updatesEDCLVM = getPipelineUpdatesHariIni(kodeCabang, tanggal);

  // Generate ID unik
  var idLaporan = generateID(kodeCabang);

  // Susun teks WA
  var teksWA = susunTeksWA(
    tanggalFormatted, area, namaCabang, kodeCabang,
    jumlahLVM, jumlahEDC, jumlahEDCPOT, totalAkuisisi,
    jumlahPemasLVM, jumlahRetensiEDC,
    totalLeadsCakra, totalKunjunganCakra, gapCakra,
    jumlahLivinFood, namaMerchantLivinFood,
    updatesEDCLVM,
    kendala, keterangan
  );

  // URL laporan (di GitHub Pages) — diisi kosong, ditentukan di sisi frontend
  var linkLaporan = '';

  // Ringkasan update EDC to LVM hari ini, buat ditulis ke sheet "Laporan Harian"
  // (kolom lama dipertahankan supaya nggak perlu migrasi skema baru lagi).
  var ringkasanMerchantEDCLVM = updatesEDCLVM.length
    ? updatesEDCLVM.map(function (u) { return u.namaMerchant + ' (' + u.status + ')'; }).join('; ')
    : '-';
  var ringkasanStatusEDCLVM = updatesEDCLVM.length
    ? updatesEDCLVM.length + ' merchant diupdate'
    : '-';

  // Susun baris berdasarkan NAMA HEADER asli di sheet (bukan posisi tetap),
  // supaya aman walau urutan kolom "Jumlah Akuisisi EDC POT" berbeda-beda.
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var nilaiKolom = {
    'Timestamp'               : timestamp,
    'Tanggal Laporan'         : tanggal,
    'Area'                    : area,
    'Nama Cabang'             : namaCabang,
    'Kode Cabang'             : kodeCabang,
    'Jumlah Akuisisi LVM'     : jumlahLVM,
    'Jumlah Akuisisi EDC'     : jumlahEDC,
    'Jumlah Akuisisi EDC POT' : jumlahEDCPOT,
    'Jumlah Pemasangan LVM'   : jumlahPemasLVM,
    'Jumlah Retensi EDC'      : jumlahRetensiEDC,
    'Total Akuisisi'          : totalAkuisisi,
    'Total Leads Cakra'       : totalLeadsCakra,
    'Total Kunjungan Cakra'   : totalKunjunganCakra,
    'Gap (sesuai Cakra)'      : gapCakra,
    "Jumlah Transaksi Livin' Food" : jumlahLivinFood,
    "Nama Merchant Livin' Food"    : namaMerchantLivinFood,
    'Merchant EDC to LVM'          : ringkasanMerchantEDCLVM,
    'Status Progress EDC to LVM'   : ringkasanStatusEDCLVM,
    'Kendala'                 : kendala,
    'Keterangan'              : keterangan,
    'ID Laporan'              : idLaporan,
    'Teks WA'                 : teksWA,
    'Link Laporan'            : linkLaporan
  };

  var row = headers.map(function (h) {
    return nilaiKolom.hasOwnProperty(h) ? nilaiKolom[h] : '';
  });

  sheet.appendRow(row);

  // Update monitoring Livin' Food (jika eligible)
  if (namaMerchantLivinFood && namaMerchantLivinFood !== '-') {
    try {
      updateMonitoringLivinFood(ss, tanggal, kodeCabang, namaMerchantLivinFood, jumlahLivinFood);
    } catch (e) {
      Logger.log('Monitoring Livin Food error: ' + e.message);
    }
  }

  // Catatan: update status Pipeline EDC to LVM sekarang dilakukan LANGSUNG
  // dari pipeline.html (action=updateStatus) saat merchant diklik & disimpan —
  // bukan lagi lewat form laporan harian ini. Baris ini hanya MEMBACA hasil
  // update yang sudah tersimpan di sheet (lihat updatesEDCLVM di atas).

  // Catat log
  catatLog(ss, idLaporan, tanggalFormatted, namaCabang, kodeCabang);

  Logger.log('Submit berhasil: ' + idLaporan);
  return { idLaporan: idLaporan, success: true };
}

// ============================================================
// CARI LAPORAN BERDASARKAN ID
// ============================================================
function cariLaporan(idLaporan) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.NAMA_SHEET);
  if (!sheet) return null;

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;

  var headers = data[0];
  var idxID = headers.indexOf('ID Laporan');

  for (var i = 1; i < data.length; i++) {
    if (data[i][idxID] === idLaporan) {
      var row = data[i];
      return {
        idLaporan           : row[headers.indexOf('ID Laporan')],
        timestamp           : formatTanggal(row[headers.indexOf('Timestamp')]),
        tanggal             : formatTanggal(row[headers.indexOf('Tanggal Laporan')]),
        area                : row[headers.indexOf('Area')]               || '-',
        namaCabang          : row[headers.indexOf('Nama Cabang')]        || '-',
        kodeCabang          : row[headers.indexOf('Kode Cabang')]        || '-',
        jumlahLVM           : Number(row[headers.indexOf('Jumlah Akuisisi LVM')])       || 0,
        jumlahEDC           : Number(row[headers.indexOf('Jumlah Akuisisi EDC')])       || 0,
        jumlahEDCPOT        : Number(row[headers.indexOf('Jumlah Akuisisi EDC POT')])   || 0,
        jumlahPemasanganLVM : Number(row[headers.indexOf('Jumlah Pemasangan LVM')])     || 0,
        jumlahRetensiEDC    : Number(row[headers.indexOf('Jumlah Retensi EDC')])        || 0,
        total               : Number(row[headers.indexOf('Total Akuisisi')])             || 0,
        totalLeadsCakra     : Number(row[headers.indexOf('Total Leads Cakra')])          || 0,
        totalKunjunganCakra : Number(row[headers.indexOf('Total Kunjungan Cakra')])      || 0,
        gapCakra            : Number(row[headers.indexOf('Gap (sesuai Cakra)')])         || 0,
        jumlahTransaksiLivinFood : Number(row[headers.indexOf("Jumlah Transaksi Livin' Food")]) || 0,
        namaMerchantLivinFood    : row[headers.indexOf("Nama Merchant Livin' Food")] || '-',
        // Diambil ULANG dari sheet Pipeline (bukan dibaca dari kolom ringkasan di
        // sheet ini), supaya kalau ada update susulan buat tanggal yang sama,
        // laporan.html tetap nampilin data terkini.
        updatesEDCLVM : getPipelineUpdatesHariIni(
          row[headers.indexOf('Kode Cabang')],
          row[headers.indexOf('Tanggal Laporan')]
        ),
        kendala             : row[headers.indexOf('Kendala')]            || '-',
        keterangan          : row[headers.indexOf('Keterangan')]         || '-',
        teksWA              : row[headers.indexOf('Teks WA')]            || ''
      };
    }
  }
  return null;
}

// ============================================================
// SUSUN TEKS WHATSAPP
// ============================================================
function susunTeksWA(tanggal, area, cabang, kode, lvm, edc, edcPot, total, plasLVM, retEDC, leadsCakra, kunjunganCakra, gapCakra, livinFood, namaMerchant, updatesEDCLVM, kendala, keterangan) {
  var t = '';
  t += 'Mohon izin melaporkan hasil akuisisi harian:\n\n';
  t += 'Tanggal : ' + tanggal + '\n';
  t += 'Area    : ' + area + '\n';
  t += 'Cabang  : ' + cabang + '\n';
  t += 'Kode    : ' + kode + '\n\n';
  t += '📊 Hasil Akuisisi:\n';
  t += '• LVM         : ' + lvm + '\n';
  t += '• EDC         : ' + edc + '\n';
  t += '• EDC POT     : ' + edcPot + '\n';
  t += '• Total       : ' + total + '\n\n';
  t += '🔧 Pemasangan & Retensi:\n';
  t += '• Pemasangan LVM : ' + plasLVM + '\n';
  t += '• Retensi EDC    : ' + retEDC + '\n\n';
  t += '📊 Hasil Kunjungan Cakra:\n';
  t += '• Leads     : ' + leadsCakra + '\n';
  t += '• Realisasi : ' + kunjunganCakra + '\n';
  t += '• Gap       : ' + gapCakra + '\n\n';
  t += "🛒 Livin' Food:\n";
  if (namaMerchant && namaMerchant !== '-') {
    t += '• Merchant  : ' + namaMerchant + '\n';
  }
  t += '• Transaksi : ' + livinFood + '\n\n';
  if (updatesEDCLVM && updatesEDCLVM.length) {
    t += '🔄 Progress Konversi EDC to LVM (' + updatesEDCLVM.length + ' merchant):\n';
    updatesEDCLVM.forEach(function (u) {
      var baris = '• ' + u.namaMerchant + ' — ' + u.status;
      if (u.status === 'Done Konversi to LVM' && u.mid && u.mid !== '-') {
        baris += ' (MID: ' + u.mid + ')';
      }
      if (u.status === 'Merchant Menolak' && u.alasan && u.alasan !== '-') {
        baris += ' (Alasan: ' + u.alasan + ')';
      }
      t += baris + '\n';
    });
    t += '\n';
  }
  t += 'Kendala    : ' + kendala + '\n';
  t += 'Keterangan : ' + keterangan + '\n\n';
  t += 'Terima kasih.';
  return t;
}

// ============================================================
// HELPER FUNCTIONS (BAGIAN A)
// ============================================================
function generateID(kodeCabang) {
  var now = new Date();
  var tgl = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd');
  var rnd = Math.floor(1000 + Math.random() * 9000);
  var kode = (kodeCabang || 'XXX').toString().replace(/\s/g, '').toUpperCase();
  return 'RPT-' + kode + '-' + tgl + '-' + rnd;
}

function formatTanggal(tgl) {
  if (!tgl) return '-';
  try {
    var d = new Date(tgl);
    if (isNaN(d.getTime())) return tgl.toString();
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  } catch (e) { return tgl.toString(); }
}

function catatLog(ss, idLaporan, tanggal, cabang, kode) {
  try {
    var log = ss.getSheetByName(CONFIG.NAMA_SHEET_LOG);
    if (!log) {
      log = ss.insertSheet(CONFIG.NAMA_SHEET_LOG);
      log.getRange(1, 1, 1, 5).setValues([['Timestamp Log', 'ID Laporan', 'Tanggal', 'Cabang', 'Kode']]);
      log.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#34a853').setFontColor('#ffffff');
    }
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
    log.appendRow([now, idLaporan, tanggal, cabang, kode]);
  } catch (e) { Logger.log('Log error: ' + e.message); }
}

// ============================================================
// SETUP: Buat sheet dengan header yang benar (kalau belum ada)
// ============================================================
function getOrCreateSheet(ss) {
  var sheet = ss.getSheetByName(CONFIG.NAMA_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.NAMA_SHEET);
    var headers = [
      'Timestamp', 'Tanggal Laporan', 'Area', 'Nama Cabang', 'Kode Cabang',
      'Jumlah Akuisisi LVM', 'Jumlah Akuisisi EDC', 'Jumlah Akuisisi EDC POT',
      'Jumlah Pemasangan LVM', 'Jumlah Retensi EDC', 'Total Akuisisi',
      'Total Leads Cakra', 'Total Kunjungan Cakra', 'Gap (sesuai Cakra)',
      "Jumlah Transaksi Livin' Food",
      "Nama Merchant Livin' Food",
      'Kendala', 'Keterangan', 'ID Laporan', 'Teks WA', 'Link Laporan'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Jalankan fungsi ini sekali dari Apps Script Editor
 * untuk membuat sheet dengan header yang benar (kalau sheet belum ada).
 */
function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  getOrCreateSheet(ss);
  SpreadsheetApp.getUi().alert('✅ Sheet "' + CONFIG.NAMA_SHEET + '" berhasil disiapkan!');
}

/**
 * Jalankan fungsi ini SEKALI dari Apps Script Editor (pilih dari dropdown
 * fungsi, lalu klik Run) untuk otomatis menambahkan kolom
 * "Jumlah Akuisisi EDC POT" ke sheet "Laporan Harian" kalau belum ada.
 * Aman dijalankan berkali-kali — kalau kolomnya sudah ada, tidak akan
 * dibuat dobel, cuma kasih pesan "sudah ada".
 */
function tambahKolomEDCPOT() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.NAMA_SHEET); // "Laporan Harian"
  if (!sheet) {
    beriTahu('Sheet "' + CONFIG.NAMA_SHEET + '" tidak ditemukan!');
    return;
  }

  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var namaKolomBaru = 'Jumlah Akuisisi EDC POT';

  // Kalau kolom sudah ada, tidak usah buat lagi
  if (headers.indexOf(namaKolomBaru) !== -1) {
    beriTahu('✅ Kolom "' + namaKolomBaru + '" sudah ada, tidak perlu ditambah lagi.');
    return;
  }

  // Cari posisi kolom "Jumlah Akuisisi EDC" supaya kolom baru disisipkan tepat setelahnya
  var idxEDC = headers.indexOf('Jumlah Akuisisi EDC'); // 0-based
  var posisiSisip;

  if (idxEDC !== -1) {
    sheet.insertColumnAfter(idxEDC + 1); // idxEDC+1 = posisi 1-based kolom EDC
    posisiSisip = idxEDC + 2;            // posisi 1-based kolom baru
  } else {
    // Kalau kolom "Jumlah Akuisisi EDC" tidak ditemukan, tambahkan di paling akhir
    sheet.insertColumnAfter(lastCol);
    posisiSisip = lastCol + 1;
  }

  sheet.getRange(1, posisiSisip).setValue(namaKolomBaru);
  sheet.getRange(1, posisiSisip)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff');

  beriTahu('✅ Kolom "' + namaKolomBaru + '" berhasil ditambahkan di posisi kolom ' + posisiSisip + '.');
}

/**
 * Jalankan fungsi ini SEKALI dari Apps Script Editor untuk menambahkan
 * 3 kolom baru ke sheet "Laporan Harian": "Total Leads Cakra",
 * "Total Kunjungan Cakra", dan "Gap (sesuai Cakra)" — kalau belum ada.
 * Aman dijalankan berkali-kali, kolom yang sudah ada tidak akan dibuat dobel.
 */
function tambahKolomCakra() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.NAMA_SHEET); // "Laporan Harian"
  if (!sheet) {
    beriTahu('Sheet "' + CONFIG.NAMA_SHEET + '" tidak ditemukan!');
    return;
  }

  var kolomBaru = ['Total Leads Cakra', 'Total Kunjungan Cakra', 'Gap (sesuai Cakra)'];
  var ditambahkan = [];

  kolomBaru.forEach(function (nama) {
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    if (headers.indexOf(nama) !== -1) return; // sudah ada, lewati

    sheet.insertColumnAfter(lastCol);
    sheet.getRange(1, lastCol + 1).setValue(nama);
    sheet.getRange(1, lastCol + 1)
      .setFontWeight('bold')
      .setBackground('#1a73e8')
      .setFontColor('#ffffff');
    ditambahkan.push(nama);
  });

  if (ditambahkan.length === 0) {
    beriTahu('✅ Semua kolom Cakra sudah ada, tidak perlu ditambah lagi.');
  } else {
    beriTahu('✅ Kolom berhasil ditambahkan: ' + ditambahkan.join(', '));
  }
}

// Helper: tampilkan alert kalau memang ada UI aktif (misal dipanggil dari
// menu/sidebar). Kalau dijalankan langsung dari tombol Run di Apps Script
// Editor, tidak ada UI aktif sehingga getUi() akan error — di kondisi itu,
// cukup catat pesannya di Execution Log saja (tidak masalah, pekerjaan
// utamanya tetap sudah selesai sebelum baris ini dijalankan).
function beriTahu(pesan) {
  try {
    SpreadsheetApp.getUi().alert(pesan);
  } catch (e) {
    Logger.log(pesan);
  }
}

/**
 * Jalankan fungsi ini SEKALI dari Apps Script Editor untuk menambahkan
 * kolom "Jumlah Transaksi Livin' Food" ke sheet "Laporan Harian" kalau belum ada.
 * Aman dijalankan berkali-kali.
 */
function tambahKolomLivinFood() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.NAMA_SHEET);
  if (!sheet) {
    beriTahu('Sheet "' + CONFIG.NAMA_SHEET + '" tidak ditemukan!');
    return;
  }
  var namaKolom = "Jumlah Transaksi Livin' Food";
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers.indexOf(namaKolom) !== -1) {
    beriTahu('✅ Kolom "' + namaKolom + '" sudah ada, tidak perlu ditambah lagi.');
    return;
  }
  sheet.insertColumnAfter(lastCol);
  sheet.getRange(1, lastCol + 1).setValue(namaKolom);
  sheet.getRange(1, lastCol + 1)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff');
  beriTahu('✅ Kolom "' + namaKolom + '" berhasil ditambahkan.');
}

/**
 * Jalankan fungsi ini SEKALI dari Apps Script Editor untuk menambahkan
 * kolom "Nama Merchant Livin' Food" ke sheet "Laporan Harian" kalau belum ada.
 * Aman dijalankan berkali-kali.
 */
function tambahKolomNamaMerchantLivinFood() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.NAMA_SHEET);
  if (!sheet) {
    beriTahu('Sheet "' + CONFIG.NAMA_SHEET + '" tidak ditemukan!');
    return;
  }
  var namaKolom = "Nama Merchant Livin' Food";
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers.indexOf(namaKolom) !== -1) {
    beriTahu('✅ Kolom "' + namaKolom + '" sudah ada, tidak perlu ditambah lagi.');
    return;
  }
  // Sisipkan setelah kolom "Jumlah Transaksi Livin' Food" jika ada
  var idxLivin = headers.indexOf("Jumlah Transaksi Livin' Food");
  var posisi;
  if (idxLivin !== -1) {
    sheet.insertColumnAfter(idxLivin + 1);
    posisi = idxLivin + 2;
  } else {
    sheet.insertColumnAfter(lastCol);
    posisi = lastCol + 1;
  }
  sheet.getRange(1, posisi).setValue(namaKolom);
  sheet.getRange(1, posisi)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff');
  beriTahu('✅ Kolom "' + namaKolom + '" berhasil ditambahkan.');
}

// ============================================================
// PIPELINE EDC TO LVM — daftar merchant & progress konversi
// ============================================================

/**
 * Ambil / buat sheet "Pipeline EDC to LVM". Kalau baru dibuat, cuma
 * diisi header saja — data merchant (Kode Cabang, Nama Cabang, Area,
 * Nama Merchant EDC) diisi manual oleh admin di Google Sheets.
 */
function getOrCreateSheetPipeline(ss) {
  var sheet = ss.getSheetByName(CONFIG.NAMA_SHEET_PIPELINE);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.NAMA_SHEET_PIPELINE);
    var headers = [
      'Kode Cabang', 'Nama Cabang', 'Area', 'Nama Merchant EDC', 'Alamat', 'Kota / Kab',
      'Status Progress', 'MID', 'Alasan', 'Tanggal Update Terakhir', 'Catatan'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Jalankan SEKALI dari Apps Script Editor untuk menyiapkan sheet
 * "Pipeline EDC to LVM" kalau belum ada. Setelah itu isi manual
 * baris-barisnya: Kode Cabang, Nama Cabang, Area, Nama Merchant EDC
 * (kolom Status Progress & Tanggal Update biar terisi otomatis lewat form).
 */
function setupSheetPipeline() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  getOrCreateSheetPipeline(ss);
  beriTahu('✅ Sheet "' + CONFIG.NAMA_SHEET_PIPELINE + '" berhasil disiapkan! Silakan isi data merchant-nya secara manual.');
}

/**
 * Ambil daftar merchant EDC (pipeline konversi ke LVM) milik satu cabang,
 * lengkap dengan status progress terkininya. Dipanggil dari form (index.html)
 * lewat action=getMerchantEDC saat cabang dipilih.
 */
function getMerchantEDCByCabang(kodeCabang) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.NAMA_SHEET_PIPELINE);
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  var headers      = data[0];
  var idxKode      = headers.indexOf('Kode Cabang');
  var idxMerchant  = headers.indexOf('Nama Merchant EDC');
  var idxStatus    = headers.indexOf('Status Progress');
  var idxTgl       = headers.indexOf('Tanggal Update Terakhir');

  var kodeCari = String(kodeCabang || '').trim();
  var hasil = [];
  for (var i = 1; i < data.length; i++) {
    var kodeBaris = String(data[i][idxKode] || '').trim();
    var namaMerchant = String(data[i][idxMerchant] || '').trim();
    if (kodeBaris === kodeCari && namaMerchant) {
      hasil.push({
        namaMerchant  : namaMerchant,
        status        : data[i][idxStatus] || 'Target',
        tanggalUpdate : data[i][idxTgl] ? formatTanggal(data[i][idxTgl]) : '-'
      });
    }
  }
  return hasil;
}

/**
 * Ambil SEMUA data pipeline EDC to LVM (lintas cabang), opsional
 * difilter per kodeCabang. Dipanggil dari pipeline.html lewat
 * action=getPipelineAll — kalau parameter kodeCabang dikirim,
 * cuma baris cabang itu yang dikembalikan (dipakai buat link
 * per-cabang, misal pipeline.html?kode=15200).
 */
function getPipelineAll(kodeCabangFilter) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.NAMA_SHEET_PIPELINE);
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  var headers      = data[0];
  var idxKode      = headers.indexOf('Kode Cabang');
  var idxNama      = headers.indexOf('Nama Cabang');
  var idxArea      = headers.indexOf('Area');
  var idxMerchant  = headers.indexOf('Nama Merchant EDC');
  var idxAlamat    = headers.indexOf('Alamat');
  var idxKota      = headers.indexOf('Kota / Kab');
  var idxStatus    = headers.indexOf('Status Progress');
  var idxMid       = headers.indexOf('MID');
  var idxAlasan    = headers.indexOf('Alasan');
  var idxTgl       = headers.indexOf('Tanggal Update Terakhir');
  var idxCatatan   = headers.indexOf('Catatan');

  var kodeCari = String(kodeCabangFilter || '').trim();
  var hasil = [];
  for (var i = 1; i < data.length; i++) {
    var kodeBaris     = String(data[i][idxKode] || '').trim();
    var namaMerchant  = String(data[i][idxMerchant] || '').trim();
    if (!kodeBaris || !namaMerchant) continue;
    if (kodeCari && kodeBaris !== kodeCari) continue;

    hasil.push({
      kodeCabang    : kodeBaris,
      namaCabang    : data[i][idxNama] || '-',
      area          : data[i][idxArea] || '-',
      namaMerchant  : namaMerchant,
      alamat        : idxAlamat !== -1 ? (data[i][idxAlamat] || '-') : '-',
      kota          : idxKota   !== -1 ? (data[i][idxKota]   || '-') : '-',
      status        : data[i][idxStatus] || 'Target',
      mid           : idxMid !== -1 ? (data[i][idxMid] || '-') : '-',
      alasan        : idxAlasan !== -1 ? (data[i][idxAlasan] || '-') : '-',
      tanggalUpdate : data[i][idxTgl] ? formatTanggal(data[i][idxTgl]) : '-',
      catatan       : idxCatatan !== -1 ? (data[i][idxCatatan] || '-') : '-'
    });
  }
  return hasil;
}

/**
 * Ambil semua merchant di "Pipeline EDC to LVM" milik satu cabang yang
 * "Tanggal Update Terakhir"-nya SAMA dengan tanggal laporan yang sedang
 * disubmit. Ini yang bikin teks WA laporan harian otomatis nampilin
 * SEMUA merchant yang statusnya diupdate hari itu (lewat pipeline.html),
 * bukan cuma 1 merchant — jadi kalau cabang kunjungan 5 merchant dalam
 * sehari dan update ke-5nya di pipeline.html, ke-5nya otomatis muncul
 * di teks WA laporan harian tanpa perlu pilih manual di form.
 */
function getPipelineUpdatesHariIni(kodeCabang, tanggalLaporan) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.NAMA_SHEET_PIPELINE);
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  var headers     = data[0];
  var idxKode     = headers.indexOf('Kode Cabang');
  var idxMerchant = headers.indexOf('Nama Merchant EDC');
  var idxStatus   = headers.indexOf('Status Progress');
  var idxMid      = headers.indexOf('MID');
  var idxAlasan   = headers.indexOf('Alasan');
  var idxTgl      = headers.indexOf('Tanggal Update Terakhir');
  if (idxKode === -1 || idxMerchant === -1 || idxTgl === -1) return [];

  var kodeCari = String(kodeCabang || '').trim();
  var tglCari  = formatTanggal(tanggalLaporan); // dd/MM/yyyy, sama format dgn Tanggal Update Terakhir

  var hasil = [];
  for (var i = 1; i < data.length; i++) {
    var kodeBaris = String(data[i][idxKode] || '').trim();
    if (kodeBaris !== kodeCari) continue;

    var tglBaris = data[i][idxTgl] ? formatTanggal(data[i][idxTgl]) : '';
    if (!tglBaris || tglBaris !== tglCari) continue;

    hasil.push({
      namaMerchant : data[i][idxMerchant] || '-',
      status       : data[i][idxStatus] || 'Target',
      mid          : idxMid !== -1 ? (data[i][idxMid] || '-') : '-',
      alasan       : idxAlasan !== -1 ? (data[i][idxAlasan] || '-') : '-'
    });
  }
  return hasil;
}

/**
 * Update status progress satu merchant di sheet "Pipeline EDC to LVM".
 * Dipanggil dari prosesSubmit() saat user memilih merchant + status baru
 * di form, dan dari action=updateStatus (pipeline.html). Mencari baris
 * berdasarkan kombinasi Kode Cabang + Nama Merchant, lalu menimpa kolom
 * Status Progress, MID (kalau diisi), Alasan (kalau diisi), & Tanggal
 * Update Terakhir baris itu (bukan menambah baris baru — 1 merchant =
 * 1 baris terus terupdate).
 * Parameter "mid" & "alasan" opsional — hanya ditulis kalau ada isinya.
 * Return true kalau baris ditemukan & berhasil diupdate, false kalau tidak.
 */
function updatePipelineStatus(ss, kodeCabang, namaMerchant, statusBaru, mid, alasan) {
  var sheet = ss.getSheetByName(CONFIG.NAMA_SHEET_PIPELINE);
  if (!sheet) return false;

  if (STATUS_EDC_LVM_VALID.indexOf(statusBaru) === -1) {
    Logger.log('Status progress EDC to LVM tidak valid: ' + statusBaru);
    return false;
  }

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return false;

  var headers     = data[0];
  var idxKode     = headers.indexOf('Kode Cabang');
  var idxMerchant = headers.indexOf('Nama Merchant EDC');
  var idxStatus   = headers.indexOf('Status Progress');
  var idxMid      = headers.indexOf('MID');
  var idxAlasan   = headers.indexOf('Alasan');
  var idxTgl      = headers.indexOf('Tanggal Update Terakhir');
  if (idxKode === -1 || idxMerchant === -1 || idxStatus === -1) return false;

  var kodeCari     = String(kodeCabang || '').trim();
  var merchantCari = String(namaMerchant || '').trim();

  for (var i = 1; i < data.length; i++) {
    var kodeBaris     = String(data[i][idxKode] || '').trim();
    var merchantBaris = String(data[i][idxMerchant] || '').trim();
    if (kodeBaris === kodeCari && merchantBaris === merchantCari) {
      sheet.getRange(i + 1, idxStatus + 1).setValue(statusBaru);
      if (idxMid !== -1 && mid && String(mid).trim()) {
        sheet.getRange(i + 1, idxMid + 1).setValue(String(mid).trim());
      }
      if (idxAlasan !== -1 && alasan && String(alasan).trim()) {
        sheet.getRange(i + 1, idxAlasan + 1).setValue(String(alasan).trim());
      }
      if (idxTgl !== -1) {
        sheet.getRange(i + 1, idxTgl + 1).setValue(
          Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy')
        );
      }
      return true;
    }
  }
  return false;
}

/**
 * Jalankan SEKALI dari Apps Script Editor setelah update status pipeline
 * (skema lama → Target/Done Konversi to LVM/Merchant Menolak/Merchant
 * Tutup/Perlu Kunjungan MTI), supaya baris-baris lama di sheet "Pipeline
 * EDC to LVM" yang masih pakai istilah status sebelumnya ("Proses",
 * "Sudah LVM", "DONE KONVERSI TO LVM" huruf kapital semua, "Kunjungan
 * Awal", dst.) ikut termigrasi ke istilah baru.
 * Aman dijalankan berkali-kali — baris yang sudah pakai status baru dilewati.
 */
function migrasiStatusPipeline() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.NAMA_SHEET_PIPELINE);
  if (!sheet) {
    beriTahu('Sheet "' + CONFIG.NAMA_SHEET_PIPELINE + '" tidak ditemukan!');
    return;
  }

  var PETA_MIGRASI = {
    'Proses'               : 'Target',
    'Kunjungan Awal'       : 'Target',
    'Sudah LVM'            : 'Done Konversi to LVM',
    'Done Akuisisi'        : 'Done Konversi to LVM',
    'DONE KONVERSI TO LVM' : 'Done Konversi to LVM'
  };

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    beriTahu('Sheet "' + CONFIG.NAMA_SHEET_PIPELINE + '" belum ada data.');
    return;
  }

  var headers   = data[0];
  var idxStatus = headers.indexOf('Status Progress');
  if (idxStatus === -1) {
    beriTahu('Kolom "Status Progress" tidak ditemukan di sheet.');
    return;
  }

  var jumlahDiubah = 0;
  for (var i = 1; i < data.length; i++) {
    var statusLama = String(data[i][idxStatus] || '').trim();
    if (PETA_MIGRASI.hasOwnProperty(statusLama)) {
      sheet.getRange(i + 1, idxStatus + 1).setValue(PETA_MIGRASI[statusLama]);
      jumlahDiubah++;
    }
  }

  beriTahu(jumlahDiubah > 0
    ? '✅ ' + jumlahDiubah + ' baris berhasil dimigrasikan ke status baru.'
    : '✅ Tidak ada baris dengan status lama — semua sudah pakai istilah baru.');
}

/**
 * Jalankan SEKALI dari Apps Script Editor untuk menambahkan kolom
 * "Merchant EDC to LVM" dan "Status Progress EDC to LVM" ke sheet
 * "Laporan Harian" kalau belum ada. Aman dijalankan berkali-kali.
 */
function tambahKolomEDCLVM() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.NAMA_SHEET);
  if (!sheet) {
    beriTahu('Sheet "' + CONFIG.NAMA_SHEET + '" tidak ditemukan!');
    return;
  }
  var kolomBaru = ['Merchant EDC to LVM', 'Status Progress EDC to LVM'];
  var ditambahkan = [];
  kolomBaru.forEach(function (nama) {
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    if (headers.indexOf(nama) !== -1) return;
    sheet.insertColumnAfter(lastCol);
    sheet.getRange(1, lastCol + 1).setValue(nama);
    sheet.getRange(1, lastCol + 1)
      .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
    ditambahkan.push(nama);
  });
  if (ditambahkan.length === 0) {
    beriTahu('✅ Kolom EDC to LVM sudah ada, tidak perlu ditambah lagi.');
  } else {
    beriTahu('✅ Kolom berhasil ditambahkan: ' + ditambahkan.join(', '));
  }
}

// ============================================================
// MONITORING LIVIN' FOOD — Sheet baru otomatis tiap minggu
// ============================================================

// Data merchant tetap (urutan = urutan baris di sheet)
var LIVIN_MERCHANTS = [
  { region: 'REGION X/SULAWESI & MALUKU', nama: 'Coto Dg Tayang', outlet: 'Coto Daeng Tayang' },
  { region: 'REGION X/SULAWESI & MALUKU', nama: 'Dari kopi',      outlet: 'Dari kopi'          },
  { region: 'REGION X/SULAWESI & MALUKU', nama: 'Eksposed',       outlet: 'EKSPOSED SIGNATURE' },
  { region: 'REGION X/SULAWESI & MALUKU', nama: 'Konijiwa',       outlet: 'Konijiwa'           },
  { region: 'REGION X/SULAWESI & MALUKU', nama: 'Mas Daeng',      outlet: 'Mas Daeng Kuliner'  }
];

// Peta hari (JS getDay()) → offset kolom dari G (col 7)
// Setiap hari = 2 kolom: Kode Cabang & #order
// 0=Minggu→6, 1=Senin→0, 2=Selasa→1, ..., 6=Sabtu→5
var HARI_OFFSET = { 1:0, 2:1, 3:2, 4:3, 5:4, 6:5, 0:6 };

/**
 * Hitung nomor minggu ISO dari sebuah tanggal.
 * Minggu ISO dimulai Senin, minggu pertama = minggu yang mengandung Kamis pertama tahun tsb.
 */
function getISOWeek(date) {
  var d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  var day = d.getUTCDay() || 7; // jadikan Minggu = 7
  d.setUTCDate(d.getUTCDate() + 4 - day);
  var yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return { week: Math.ceil(((d - yearStart) / 86400000 + 1) / 7), year: d.getUTCFullYear() };
}

/**
 * Dapatkan tanggal Senin dan Minggu dari sebuah tanggal (untuk label header sheet).
 */
function getRangeMinggu(date) {
  var d = new Date(date);
  var day = d.getDay() || 7; // Minggu = 7
  var senin = new Date(d); senin.setDate(d.getDate() - day + 1);
  var minggu = new Date(senin); minggu.setDate(senin.getDate() + 6);
  var bulan = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
  return (senin.getDate() + ' ' + bulan[senin.getMonth()])
       + ' – '
       + (minggu.getDate() + ' ' + bulan[minggu.getMonth()] + ' ' + minggu.getFullYear());
}

/**
 * Update monitoring Livin' Food di sheet minggu yang sesuai.
 * Sheet dibuat otomatis jika belum ada untuk minggu tersebut.
 */
function updateMonitoringLivinFood(ss, tanggal, kodeCabang, namaMerchant, jumlahOrder) {
  var tglObj = new Date(tanggal);
  var wk     = getISOWeek(tglObj);
  var sheetName = "Livin' Food W" + wk.week + ' ' + wk.year;

  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = buatSheetLivinMinggu(ss, sheetName, tglObj);

  // Tentukan kolom hari
  var dayJS    = tglObj.getDay();
  var offset   = HARI_OFFSET[dayJS];
  var colKode  = 7 + offset * 2;
  var colOrder = 7 + offset * 2 + 1;

  // Cari baris merchant (mulai baris 5 karena baris 1 = title, 2-3 = header, 4 = kosong/range)
  var DATA_START_ROW = 5;
  var merchantNorm = (namaMerchant || '').toLowerCase().trim();
  var targetRow = -1;

  for (var i = 0; i < LIVIN_MERCHANTS.length; i++) {
    if (LIVIN_MERCHANTS[i].nama.toLowerCase() === merchantNorm) {
      targetRow = DATA_START_ROW + i; break;
    }
    var sheetNama = sheet.getRange(DATA_START_ROW + i, 2).getValue();
    if (String(sheetNama).toLowerCase().trim() === merchantNorm) {
      targetRow = DATA_START_ROW + i; break;
    }
  }
  if (targetRow === -1) {
    Logger.log('Merchant tidak ditemukan: ' + namaMerchant); return;
  }

  // Akumulasi #order; gabung kode cabang jika berbeda
  var existingKode  = sheet.getRange(targetRow, colKode).getValue();
  var existingOrder = Number(sheet.getRange(targetRow, colOrder).getValue()) || 0;
  var newOrder = existingOrder + jumlahOrder;
  var newKode  = (!existingKode || existingKode === '')
    ? kodeCabang
    : (String(existingKode).indexOf(kodeCabang) === -1 ? existingKode + ', ' + kodeCabang : existingKode);

  sheet.getRange(targetRow, colKode).setValue(newKode);
  sheet.getRange(targetRow, colOrder).setValue(newOrder);
  Logger.log('Monitoring updated [' + sheetName + ']: ' + namaMerchant + ' day=' + dayJS + ' order=' + newOrder);
}

/**
 * Buat sheet monitoring baru untuk satu minggu.
 * Nama sheet: "Livin' Food W27 2026"
 */
function buatSheetLivinMinggu(ss, sheetName, tglObj) {
  var sheet    = ss.insertSheet(sheetName);
  var hdrFill  = '#002060';
  var hdrFont  = '#FFFFFF';
  var days     = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu'];
  var rangeMgg = getRangeMinggu(tglObj);

  // Baris 1: periode minggu (A1:F1) + judul Cabang Pemesan (G1:T1)
  sheet.getRange('A1:F1').merge();
  sheet.getRange('A1').setValue("Periode: " + rangeMgg);
  styleHeader(sheet.getRange('A1'), hdrFill, hdrFont, true, 'left');

  sheet.getRange('G1:T1').merge();
  sheet.getRange('G1').setValue('Cabang Pemesan (min order 10 per hari)');
  styleHeader(sheet.getRange('G1'), hdrFill, hdrFont, true, 'center');

  // Baris 2–3: header kolom tetap (merge rows 2:3 per kolom)
  var fixedHeaders = ['Region Merchant', 'Nama Merchant', 'Nama Outlet', 'Area',
                      'Nama PIC Area/ Jabatan', 'No HP  PIC Area'];
  for (var f = 0; f < fixedHeaders.length; f++) {
    var r = sheet.getRange(2, f + 1, 2, 1);
    r.merge(); r.setValue(fixedHeaders[f]);
    styleHeader(r, hdrFill, hdrFont, true, 'center');
  }

  // Baris 2–3: header hari
  for (var d = 0; d < days.length; d++) {
    var col = 7 + d * 2;
    var rng2 = sheet.getRange(2, col, 1, 2);
    rng2.merge(); rng2.setValue(days[d]);
    styleHeader(rng2, hdrFill, hdrFont, true, 'center');

    styleHeader(sheet.getRange(3, col),     hdrFill, hdrFont, true, 'center').setValue('Kode Cabang');
    styleHeader(sheet.getRange(3, col + 1), hdrFill, hdrFont, true, 'center').setValue('#order');
  }

  // Baris 4: kosong (spacer) — baris data merchant mulai baris 5
  sheet.setRowHeight(3, 32);

  // Lebar kolom
  var widths = [200, 130, 160, 85, 95, 95, 65, 55, 65, 55, 65, 55, 65, 55, 65, 55, 65, 55, 65, 55];
  for (var w = 0; w < widths.length; w++) {
    sheet.setColumnWidth(w + 1, widths[w]);
  }

  // Data rows merchant (mulai baris 5)
  for (var m = 0; m < LIVIN_MERCHANTS.length; m++) {
    var row = 5 + m;
    sheet.getRange(row, 1).setValue(LIVIN_MERCHANTS[m].region);
    sheet.getRange(row, 2).setValue(LIVIN_MERCHANTS[m].nama);
    sheet.getRange(row, 3).setValue(LIVIN_MERCHANTS[m].outlet);
  }

  // Border seluruh tabel (baris 1–(5+jumlah merchant-1), kolom A–T)
  sheet.getRange(1, 1, 4 + LIVIN_MERCHANTS.length, 20)
    .setBorder(true, true, true, true, true, true);

  sheet.setFrozenRows(3);
  Logger.log('Sheet baru dibuat: ' + sheetName + ' (' + rangeMgg + ')');
  return sheet;
}

function styleHeader(range, bgHex, fontHex, bold, halign) {
  range.setBackground(bgHex)
       .setFontColor(fontHex)
       .setFontWeight(bold ? 'bold' : 'normal')
       .setHorizontalAlignment(halign || 'center')
       .setVerticalAlignment('middle')
       .setWrap(true);
  return range;
}

/**
 * Test manual: buat sheet monitoring untuk minggu ini.
 * Jalankan dari Apps Script Editor untuk cek hasilnya.
 */
function testBuatSheetLivinMingguIni() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tgl = new Date();
  var wk  = getISOWeek(tgl);
  var sheetName = "Livin' Food W" + wk.week + ' ' + wk.year;
  var existing = ss.getSheetByName(sheetName);
  if (existing) ss.deleteSheet(existing);
  buatSheetLivinMinggu(ss, sheetName, tgl);
  beriTahu('✅ Sheet "' + sheetName + '" berhasil dibuat!');
}

/**
 * Test: cek apakah API berjalan dengan benar.
 * Jalankan dari Apps Script Editor, lihat hasilnya di Logs (View > Logs).
 */
function testAPI() {
  Logger.log('URL Web App: ' + ScriptApp.getService().getUrl());
  Logger.log('Sheet: ' + CONFIG.NAMA_SHEET);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.NAMA_SHEET);
  Logger.log('Sheet ada: ' + (sheet ? 'YA' : 'BELUM — jalankan setupSheet()'));
}


// ============================================================
// BAGIAN B — REKAP CABANG (menu & sidebar di Google Sheets)
// ============================================================

const SHEET_LAPORAN = "Laporan Harian";
const SHEET_MASTER  = "Master Cabang"; // opsional, lihat keterangan di bawah

// Nama kolom di sheet "Laporan Harian" — HARUS sama persis dengan tulisan
// header di baris 1 sheet tersebut. Kalau nanti ada kolom baru, cukup
// tambahkan barisnya di sini, tidak perlu ubah bagian lain.
const KOLOM = {
  TANGGAL     : "Tanggal Laporan",
  AREA        : "Area",
  CABANG      : "Nama Cabang",
  KODE        : "Kode Cabang",
  LVM         : "Jumlah Akuisisi LVM",
  EDC         : "Jumlah Akuisisi EDC",
  EDC_POT     : "Jumlah Akuisisi EDC POT",
  PASANG_LVM  : "Jumlah Pemasangan LVM",
  RETENSI_EDC : "Jumlah Retensi EDC",
  TOTAL       : "Total Akuisisi"
};

// Baca header baris 1 dari sebuah sheet, lalu cocokkan dengan KOLOM di atas.
// Hasilnya: object berisi index kolom (0-based). -1 berarti kolom tidak ditemukan.
function getColMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  for (const key in KOLOM) {
    map[key] = headers.indexOf(KOLOM[key]);
  }
  return map;
}

// Ambil angka dari sebuah baris dengan aman. Kalau kolomnya tidak ada
// (index -1, misal header belum dibuat di sheet), otomatis dianggap 0.
function ambilAngka(row, idx) {
  if (idx < 0) return 0;
  return Number(row[idx]) || 0;
}

// =====================================================================
// 1. MENU - otomatis muncul saat file dibuka
// =====================================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📊 Rekap Cabang")
    .addItem("Buka Panel Rekap", "bukaPanel")
    .addItem("Buat Sheet Rekap Hari Ini", "buatSheetRekapHariIni")
    .addToUi();
}

// =====================================================================
// 2. BUKA SIDEBAR
// =====================================================================
function bukaPanel() {
  const html = HtmlService.createHtmlOutput(getHtmlSidebar())
    .setTitle("Rekap Laporan Cabang")
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

// =====================================================================
// 3. AMBIL DAFTAR TANGGAL (dipanggil dari HTML)
// =====================================================================
function getDaftarTanggal() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_LAPORAN);
  if (!sheet) return [];

  const col         = getColMap(sheet);
  const data        = sheet.getDataRange().getValues();
  const tanggalSet  = new Set();

  for (let i = 1; i < data.length; i++) {
    const tgl = data[i][col.TANGGAL];
    if (tgl instanceof Date && !isNaN(tgl)) {
      tanggalSet.add(formatTanggalISO(tgl));
    }
  }

  // Urutkan dari terbaru
  return Array.from(tanggalSet).sort((a, b) => b.localeCompare(a));
}

// =====================================================================
// 4. AMBIL SEMUA CABANG DARI MASTER CABANG
// Prioritas utama: sheet "Master Cabang"
// Fallback: ambil dari data unik di "Laporan Harian" (jika master kosong)
// =====================================================================
function getAllCabang() {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const master      = {};
  const sheetMaster = ss.getSheetByName(SHEET_MASTER);

  // ✅ Prioritas utama: baca dari sheet Master Cabang
  if (sheetMaster && sheetMaster.getLastRow() > 1) {
    const dataMaster = sheetMaster.getDataRange().getValues();
    for (let i = 1; i < dataMaster.length; i++) {
      const kode = String(dataMaster[i][0]).trim();
      const nama = String(dataMaster[i][1]).trim();
      const area = String(dataMaster[i][2] || "-").trim();
      if (kode && nama && kode !== "" && nama !== "") {
        master[kode] = { nama: nama, area: area };
      }
    }
    return master;
  }

  // ⚠️ Fallback: ambil cabang unik dari Laporan Harian jika master belum diisi
  const sheetLaporan = ss.getSheetByName(SHEET_LAPORAN);
  if (!sheetLaporan) return {};

  const col  = getColMap(sheetLaporan);
  const data = sheetLaporan.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const kode = String(data[i][col.KODE]).trim();
    const nama = String(data[i][col.CABANG]).trim();
    const area = String(data[i][col.AREA] || "-").trim();
    if (kode && nama) {
      master[kode] = { nama: nama, area: area };
    }
  }

  return master;
}

// =====================================================================
// 5. REKAP BERDASARKAN TANGGAL (dipanggil dari HTML)
// =====================================================================
function getRekap(tanggalStr) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_LAPORAN);
  if (!sheet) return { sudahLapor: [], belumLapor: [], error: "Sheet tidak ditemukan!" };

  const col         = getColMap(sheet);
  const data        = sheet.getDataRange().getValues();
  const allCabang   = getAllCabang();
  const sudahLapor  = {};

  // Cari cabang yang sudah lapor di tanggal tsb
  for (let i = 1; i < data.length; i++) {
    const tgl  = data[i][col.TANGGAL];
    const kode = data[i][col.KODE];
    const nama = data[i][col.CABANG];
    const area = data[i][col.AREA];

    if (tgl instanceof Date && formatTanggalISO(tgl) === tanggalStr && kode) {
      if (!sudahLapor[kode]) {
        sudahLapor[kode] = {
          nama          : nama,
          area          : area || "-",
          jumlahLaporan : 0,
          lvm           : 0,
          edc           : 0,
          edcPot        : 0,
          pasangLvm     : 0,
          retensiEdc    : 0,
          total         : 0
        };
      }
      sudahLapor[kode].jumlahLaporan++;
      sudahLapor[kode].lvm        += ambilAngka(data[i], col.LVM);
      sudahLapor[kode].edc        += ambilAngka(data[i], col.EDC);
      sudahLapor[kode].edcPot     += ambilAngka(data[i], col.EDC_POT);
      sudahLapor[kode].pasangLvm  += ambilAngka(data[i], col.PASANG_LVM);
      sudahLapor[kode].retensiEdc += ambilAngka(data[i], col.RETENSI_EDC);
      sudahLapor[kode].total      += ambilAngka(data[i], col.TOTAL);
    }
  }

  // Cabang yang belum lapor
  const belumLapor = [];
  for (const kode in allCabang) {
    if (!sudahLapor[kode]) {
      belumLapor.push({
        kode : kode,
        nama : allCabang[kode].nama,
        area : allCabang[kode].area
      });
    }
  }

  // Format hasil sudah lapor
  const sudahLaporArr = Object.entries(sudahLapor).map(([kode, info]) => ({
    kode          : kode,
    nama          : info.nama,
    area          : info.area,
    jumlahLaporan : info.jumlahLaporan,
    lvm           : info.lvm,
    edc           : info.edc,
    edcPot        : info.edcPot,
    pasangLvm     : info.pasangLvm,
    retensiEdc    : info.retensiEdc,
    total         : info.total
  }));

  // Urutkan berdasarkan area lalu nama
  const sortByAreaNama = (a, b) => a.area.localeCompare(b.area) || a.nama.localeCompare(b.nama);
  sudahLaporArr.sort(sortByAreaNama);
  belumLapor.sort(sortByAreaNama);

  return {
    tanggal      : tanggalStr,
    totalCabang  : Object.keys(allCabang).length,
    sudahLapor   : sudahLaporArr,
    belumLapor   : belumLapor
  };
}

// =====================================================================
// 6. BUAT SHEET REKAP (tombol ekspor)
// =====================================================================
function buatSheetRekapHariIni() {
  const tanggal = formatTanggalISO(new Date());
  buatSheetRekap(tanggal);
}

function buatSheetRekap(tanggalStr) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const rekap  = getRekap(tanggalStr);
  const namaSheet = "Rekap " + tanggalStr;

  // Hapus sheet lama jika sudah ada
  const existing = ss.getSheetByName(namaSheet);
  if (existing) ss.deleteSheet(existing);

  const sheet = ss.insertSheet(namaSheet);

  // Header utama
  sheet.getRange("A1").setValue("REKAP LAPORAN CABANG");
  sheet.getRange("A2").setValue("Tanggal: " + tanggalStr);
  sheet.getRange("A3").setValue(
    `Total: ${rekap.sudahLapor.length} sudah lapor, ${rekap.belumLapor.length} belum lapor dari ${rekap.totalCabang} cabang`
  );

  // ---- Tabel SUDAH LAPOR ----
  const headerSudah = ["No", "Kode", "Area", "Nama Cabang", "Akuisisi LVM", "Akuisisi EDC", "Akuisisi EDC POT", "Pemasangan LVM", "Retensi EDC", "Total Akuisisi", "Jml Laporan"];
  const JML_KOLOM = headerSudah.length;
  sheet.getRange(5, 1, 1, JML_KOLOM).setValues([headerSudah]);
  sheet.getRange(5, 1, 1, JML_KOLOM)
    .setBackground("#34a853").setFontColor("white").setFontWeight("bold");

  rekap.sudahLapor.forEach((c, i) => {
    sheet.getRange(6 + i, 1, 1, JML_KOLOM).setValues([[
      i + 1, c.kode, c.area, c.nama,
      c.lvm, c.edc, c.edcPot, c.pasangLvm, c.retensiEdc, c.total,
      c.jumlahLaporan
    ]]);
    if (i % 2 === 0) sheet.getRange(6 + i, 1, 1, JML_KOLOM).setBackground("#d9ead3");
  });

  // Baris total
  const totalRow = 6 + rekap.sudahLapor.length;
  const totals = rekap.sudahLapor.reduce((acc, c) => {
    acc.lvm        += c.lvm;
    acc.edc        += c.edc;
    acc.edcPot     += c.edcPot;
    acc.pasangLvm  += c.pasangLvm;
    acc.retensiEdc += c.retensiEdc;
    acc.total      += c.total;
    return acc;
  }, { lvm: 0, edc: 0, edcPot: 0, pasangLvm: 0, retensiEdc: 0, total: 0 });

  sheet.getRange(totalRow, 1, 1, JML_KOLOM).setValues([[
    "", "", "", "TOTAL",
    totals.lvm, totals.edc, totals.edcPot, totals.pasangLvm, totals.retensiEdc, totals.total, ""
  ]]);
  sheet.getRange(totalRow, 1, 1, JML_KOLOM).setBackground("#b6d7a8").setFontWeight("bold");

  // ---- Tabel BELUM LAPOR ----
  const startRow = 6 + rekap.sudahLapor.length + 3;
  sheet.getRange(startRow, 1).setValue("BELUM LAPOR");
  const headerBelum = ["No", "Kode", "Area", "Nama Cabang"];
  sheet.getRange(startRow + 1, 1, 1, headerBelum.length).setValues([headerBelum]);
  sheet.getRange(startRow + 1, 1, 1, headerBelum.length)
    .setBackground("#ea4335").setFontColor("white").setFontWeight("bold");

  rekap.belumLapor.forEach((c, i) => {
    sheet.getRange(startRow + 2 + i, 1, 1, 4).setValues([[i + 1, c.kode, c.area, c.nama]]);
    if (i % 2 === 0) sheet.getRange(startRow + 2 + i, 1, 1, 4).setBackground("#fce8e6");
  });

  sheet.autoResizeColumns(1, JML_KOLOM);
  SpreadsheetApp.getUi().alert(`Sheet rekap "${namaSheet}" berhasil dibuat!`);
  return namaSheet;
}

// =====================================================================
// 7. HELPER (BAGIAN B)
// Catatan: nama fungsi ini "formatTanggalISO" (bukan "formatTanggal")
// karena nama "formatTanggal" sudah dipakai di Bagian A untuk keperluan
// berbeda (format dd/MM/yyyy untuk teks WA & tampilan laporan).
// =====================================================================
function formatTanggalISO(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

// =====================================================================
// 8. HTML SIDEBAR
// =====================================================================
function getHtmlSidebar() {
  return `
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: Arial, sans-serif; font-size: 13px; background: #f8f9fa; }

  .header {
    background: linear-gradient(135deg, #1a73e8, #0d47a1);
    color: white; padding: 14px 16px; text-align: center;
  }
  .header h2 { font-size: 15px; margin-bottom: 3px; }
  .header p  { font-size: 11px; opacity: 0.85; }

  .content { padding: 14px; }

  .form-group { margin-bottom: 12px; }
  label { display: block; font-weight: bold; margin-bottom: 5px; color: #333; }
  select {
    width: 100%; padding: 8px 10px; border: 1px solid #ccc;
    border-radius: 6px; font-size: 13px; background: white;
  }

  .btn {
    width: 100%; padding: 9px; border: none; border-radius: 6px;
    font-size: 13px; font-weight: bold; cursor: pointer; margin-bottom: 6px;
  }
  .btn-primary { background: #1a73e8; color: white; }
  .btn-primary:hover { background: #1558b0; }
  .btn-export  { background: #34a853; color: white; }
  .btn-export:hover { background: #2d8f47; }

  .summary {
    display: flex; gap: 8px; margin-bottom: 14px;
  }
  .stat-box {
    flex: 1; text-align: center; padding: 10px 6px;
    border-radius: 8px; font-weight: bold;
  }
  .stat-box .num  { font-size: 22px; }
  .stat-box .lbl  { font-size: 10px; margin-top: 2px; }
  .stat-sudah { background: #e6f4ea; color: #1e7e34; border: 1px solid #b7dfbc; }
  .stat-belum { background: #fce8e6; color: #c0392b; border: 1px solid #f5b7b1; }
  .stat-total { background: #e8f0fe; color: #1a73e8; border: 1px solid #aecbfa; }

  .tab-header { display: flex; margin-bottom: 8px; border-bottom: 2px solid #ddd; }
  .tab-btn {
    flex: 1; padding: 8px; background: none; border: none;
    cursor: pointer; font-size: 12px; font-weight: bold; color: #666;
  }
  .tab-btn.active { color: #1a73e8; border-bottom: 3px solid #1a73e8; margin-bottom: -2px; }

  .tab-content { display: none; }
  .tab-content.active { display: block; }

  .search-box {
    width: 100%; padding: 6px 10px; margin-bottom: 8px;
    border: 1px solid #ccc; border-radius: 6px; font-size: 12px;
  }

  .list-item {
    padding: 7px 10px; border-radius: 6px; margin-bottom: 4px;
    border-left: 4px solid;
  }
  .item-sudah { background: #f0faf2; border-color: #34a853; }
  .item-belum { background: #fff5f5; border-color: #ea4335; }
  .item-nama  { font-weight: bold; font-size: 12px; color: #333; }
  .item-info  { font-size: 11px; color: #666; margin-top: 2px; }
  .item-stats { display: flex; flex-wrap: wrap; gap: 4px; margin-top: 5px; }
  .stat-pill  {
    font-size: 10px; padding: 2px 6px; border-radius: 10px;
    background: #e8f5e9; color: #2e7d32; border: 1px solid #c8e6c9;
  }
  .stat-pill.total { background: #e3f2fd; color: #1565c0; border-color: #bbdefb; font-weight: bold; }
  .badge {
    float: right; font-size: 10px; padding: 2px 6px;
    border-radius: 10px; margin-top: 1px;
  }
  .badge-sudah { background: #34a853; color: white; }
  .badge-belum { background: #ea4335; color: white; }

  .loading { text-align: center; padding: 30px; color: #999; }
  .empty   { text-align: center; padding: 20px; color: #999; font-style: italic; }
  .error   { background: #fce8e6; color: #c0392b; padding: 10px; border-radius: 6px; font-size: 12px; }
</style>
</head>
<body>

<div class="header">
  <h2>📊 Rekap Laporan Cabang</h2>
  <p>Pantau cabang yang sudah & belum lapor</p>
</div>

<div class="content">

  <div class="form-group">
    <label>📅 Pilih Tanggal</label>
    <select id="selTanggal">
      <option value="">-- Memuat tanggal... --</option>
    </select>
  </div>

  <button class="btn btn-primary" onclick="cariRekap()">🔍 Tampilkan Rekap</button>
  <button class="btn btn-export" onclick="eksporSheet()" id="btnEkspor" style="display:none">
    📄 Buat Sheet Rekap
  </button>

  <div id="hasil"></div>

</div>

<script>
  let hasilGlobal = null;

  // Load daftar tanggal saat halaman siap
  window.onload = function() {
    google.script.run
      .withSuccessHandler(function(tglList) {
        const sel = document.getElementById('selTanggal');
        sel.innerHTML = '<option value="">-- Pilih Tanggal --</option>';
        tglList.forEach(t => {
          const opt = document.createElement('option');
          opt.value = t;
          opt.textContent = formatTampil(t);
          sel.appendChild(opt);
        });
        // Auto pilih tanggal terbaru
        if (tglList.length > 0) {
          sel.value = tglList[0];
          cariRekap();
        }
      })
      .withFailureHandler(showError)
      .getDaftarTanggal();
  };

  function cariRekap() {
    const tgl = document.getElementById('selTanggal').value;
    if (!tgl) { alert('Pilih tanggal dulu ya!'); return; }

    document.getElementById('hasil').innerHTML =
      '<div class="loading">⏳ Sedang memuat data...</div>';
    document.getElementById('btnEkspor').style.display = 'none';

    google.script.run
      .withSuccessHandler(tampilkanHasil)
      .withFailureHandler(showError)
      .getRekap(tgl);
  }

  function tampilkanHasil(data) {
    hasilGlobal = data;
    const sudah = data.sudahLapor;
    const belum = data.belumLapor;

    let html = \`
      <div class="summary">
        <div class="stat-box stat-total">
          <div class="num">\${data.totalCabang}</div>
          <div class="lbl">Total Cabang</div>
        </div>
        <div class="stat-box stat-sudah">
          <div class="num">\${sudah.length}</div>
          <div class="lbl">Sudah Lapor</div>
        </div>
        <div class="stat-box stat-belum">
          <div class="num">\${belum.length}</div>
          <div class="lbl">Belum Lapor</div>
        </div>
      </div>

      <div class="tab-header">
        <button class="tab-btn active" onclick="switchTab('sudah', this)">
          ✅ Sudah Lapor (\${sudah.length})
        </button>
        <button class="tab-btn" onclick="switchTab('belum', this)">
          ❌ Belum Lapor (\${belum.length})
        </button>
      </div>

      <input class="search-box" type="text" placeholder="🔎 Cari nama cabang..." onkeyup="filterList(this.value)" />

      <div id="tab-sudah" class="tab-content active">
    \`;

    if (sudah.length === 0) {
      html += '<div class="empty">Tidak ada cabang yang lapor di tanggal ini.</div>';
    } else {
      sudah.forEach((c, i) => {
        const dup = c.jumlahLaporan > 1 ? \` (\${c.jumlahLaporan}x)\` : '';
        html += \`
          <div class="list-item item-sudah" data-nama="\${c.nama.toLowerCase()}" data-area="\${c.area.toLowerCase()}">
            <span class="badge badge-sudah">✓\${dup}</span>
            <div class="item-nama">\${c.nama}</div>
            <div class="item-info">\${c.area} &bull; Kode: \${c.kode}</div>
            <div class="item-stats">
              <span class="stat-pill">LVM: \${c.lvm}</span>
              <span class="stat-pill">EDC: \${c.edc}</span>
              <span class="stat-pill">EDC POT: \${c.edcPot}</span>
              <span class="stat-pill">Pasang: \${c.pasangLvm}</span>
              <span class="stat-pill">Retensi: \${c.retensiEdc}</span>
              <span class="stat-pill total">Total: \${c.total}</span>
            </div>
          </div>\`;
      });
    }

    html += '</div><div id="tab-belum" class="tab-content">';

    if (belum.length === 0) {
      html += '<div class="empty">🎉 Semua cabang sudah lapor!</div>';
    } else {
      belum.forEach(c => {
        html += \`
          <div class="list-item item-belum" data-nama="\${c.nama.toLowerCase()}" data-area="\${c.area.toLowerCase()}">
            <span class="badge badge-belum">✗</span>
            <div class="item-nama">\${c.nama}</div>
            <div class="item-info">\${c.area} &bull; Kode: \${c.kode}</div>
          </div>\`;
      });
    }

    html += '</div>';

    document.getElementById('hasil').innerHTML = html;
    document.getElementById('btnEkspor').style.display = 'block';
  }

  function switchTab(tab, btn) {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById('tab-' + tab).classList.add('active');
    document.querySelector('.search-box').value = '';
    filterList('');
  }

  function filterList(keyword) {
    const kw = keyword.toLowerCase();
    document.querySelectorAll('.tab-content.active .list-item').forEach(el => {
      const cocok = el.dataset.nama.includes(kw) || el.dataset.area.includes(kw);
      el.style.display = cocok ? 'block' : 'none';
    });
  }

  function eksporSheet() {
    const tgl = document.getElementById('selTanggal').value;
    document.getElementById('btnEkspor').textContent = '⏳ Membuat sheet...';
    google.script.run
      .withSuccessHandler(function(nama) {
        document.getElementById('btnEkspor').textContent = '📄 Buat Sheet Rekap';
        alert('Sheet "' + nama + '" berhasil dibuat!');
      })
      .withFailureHandler(showError)
      .buatSheetRekap(tgl);
  }

  function showError(err) {
    document.getElementById('hasil').innerHTML =
      '<div class="error">⚠️ Error: ' + err.message + '</div>';
  }

  function formatTampil(tglStr) {
    const [y, m, d] = tglStr.split('-');
    const bulan = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
    return d + ' ' + bulan[parseInt(m)-1] + ' ' + y;
  }
</script>
</body>
</html>
`;
}
