// ============================================================
//  SISTEM HRIS PT JAPA INDOTAMA
//  File: Code.gs
//  Versi: 3.1 — Offboarding: Data DIPINDAHKAN ke db_alumni
//  Perubahan dari v3.0:
//    - prosesOffboarding() → DELETE baris dari db_karyawan,
//      PINDAHKAN ke db_alumni (bukan hanya tandai non-aktif)
//    - getSemuaKaryawan() hanya mengembalikan karyawan AKTIF
//    - getDataAlumni() membaca dari db_alumni
//    - batalkanOffboarding() memindahkan kembali dari db_alumni
//      ke db_karyawan
// ============================================================

var SHEET_DB      = "db_karyawan";
var SHEET_HISTORY = "db_history";
var SHEET_KONTRAK = "db_kontrak";
var SHEET_ALUMNI  = "db_alumni";

// ── FASE 2: REKRUTMEN & ATS ──
var SHEET_MPP       = "db_mpp";
var SHEET_KANDIDAT  = "db_kandidat";

// Header db_mpp — satu baris per posisi yang dibuka
var HEADERS_MPP = [
  "MPP ID",         // 0  — auto: MPP-YYYYMM-001
  "TGL BUKA",       // 1
  "DEPARTEMEN",     // 2
  "UNIT",           // 3
  "POSISI / JABATAN",// 4
  "LEVEL",          // 5
  "HUB. KERJA",     // 6  — Permanent/Kontrak/dll
  "KUOTA",          // 7  — jumlah orang dibutuhkan
  "ALASAN REKRUTMEN",// 8 — Penggantian/Tambahan/dll
  "TARGET SELESAI", // 9
  "STATUS MPP",     // 10 — Open/On Progress/Closed/Cancelled
  "HIRED",          // 11 — jumlah sudah dihire
  "CATATAN",        // 12
  "DIBUAT OLEH",    // 13
  "TIMESTAMP"       // 14
];

// Header db_kandidat — satu baris per pelamar per MPP
var HEADERS_KANDIDAT = [
  "KANDIDAT ID",    // 0  — auto: KND-YYYYMM-001
  "MPP ID",         // 1  — referensi ke db_mpp
  "POSISI",         // 2  — dari MPP
  "DEPARTEMEN",     // 3  — dari MPP
  "NAMA KANDIDAT",  // 4
  "NO HP",          // 5
  "EMAIL",          // 6
  "SUMBER",         // 7  — Jobstreet/LinkedIn/Referral/Walk-in/dll
  "LINK CV",        // 8  — Google Drive link
  "TGL DAFTAR",     // 9
  "TAHAPAN",        // 10 — tahapan aktif saat ini
  "STATUS",         // 11 — Proses/Hired/Reject/Withdraw
  // ── Hasil per tahapan ──
  "TGL SCREENING",  // 12
  "HASIL SCREENING",// 13 — Lanjut/Tidak Lanjut
  "CATATAN SCREENING",// 14
  "TGL TES TERTULIS",// 15
  "NILAI TES TERTULIS",// 16 — angka 0–100
  "HASIL TES TERTULIS",// 17 — Lanjut/Tidak Lanjut
  "CATATAN TES TERTULIS",// 18
  "TGL PSIKOTES",   // 19
  "HASIL PSIKOTES", // 20 — Lanjut/Tidak Lanjut
  "CATATAN PSIKOTES",// 21
  "TGL ITV HRD",    // 22
  "HASIL ITV HRD",  // 23
  "CATATAN ITV HRD",// 24
  "TGL ITV USER",   // 25
  "HASIL ITV USER", // 26
  "CATATAN ITV USER",// 27
  "TGL OFFERING",   // 28
  "GAJI DITAWARKAN",// 29
  "STATUS OFFERING",// 30 — Diterima/Ditolak/Negosiasi
  "CATATAN OFFERING",// 31
  "TGL HIRED",      // 32
  "TERAKHIR DIUPDATE",// 33
  "TIMESTAMP INPUT", // 34
  "TGL LAHIR"        // 35 — untuk scoring IST
];

// Header db_karyawan — sama persis dengan v2.1 (index 0–20)
// Tidak ada kolom tambahan. Data alumni sudah pindah ke db_alumni.
var HEADERS_DB = [
  "NO",             // 0
  "NIP",            // 1
  "NAMA",           // 2
  "DEPARTEMEN",     // 3
  "UNIT",           // 4
  "JABATAN",        // 5
  "LEVEL",          // 6
  "NIK",            // 7
  "NO TLP",         // 8
  "ALAMAT",         // 9
  "HUBUNGAN KERJA", // 10
  "AGAMA",          // 11
  "TANGGAL LAHIR",  // 12
  "USIA",           // 13
  "PENDIDIKAN",     // 14
  "JENIS KELAMIN",  // 15
  "TANGGAL MASUK",  // 16
  "MASA KERJA",     // 17
  "TANGGAL INPUT",  // 18
  "TERAKHIR DIUBAH",// 19
  "AKHIR KONTRAK"   // 20
];

var HEADERS_HIST = [
  "NIP", "NAMA", "TIPE PERUBAHAN", "TANGGAL EFEKTIF",
  "KETERANGAN", "DICATAT OLEH", "TIMESTAMP INPUT"
];

var HEADERS_KONTRAK = [
  "NO", "NIP", "NAMA", "DEPARTEMEN", "UNIT", "JABATAN", "LEVEL",
  "HUB. KERJA", "TGL MASUK", "AKHIR KONTRAK", "SISA HARI",
  "STATUS PKWT", "TERAKHIR DIUPDATE"
];

// db_alumni menyimpan SEMUA data karyawan yang sudah keluar
// Kolom 0–20 = salinan penuh dari db_karyawan
// Kolom 21–25 = info offboarding
var HEADERS_ALUMNI = [
  "NO",             // 0
  "NIP",            // 1
  "NAMA",           // 2
  "DEPARTEMEN",     // 3
  "UNIT",           // 4
  "JABATAN",        // 5
  "LEVEL",          // 6
  "NIK",            // 7
  "NO TLP",         // 8
  "ALAMAT",         // 9
  "HUBUNGAN KERJA", // 10
  "AGAMA",          // 11
  "TANGGAL LAHIR",  // 12
  "USIA SAAT KELUAR",// 13
  "PENDIDIKAN",     // 14
  "JENIS KELAMIN",  // 15
  "TANGGAL MASUK",  // 16
  "MASA KERJA TOTAL",// 17
  "TANGGAL INPUT",  // 18
  "DICATAT OLEH",   // 19
  "AKHIR KONTRAK",  // 20
  "TGL KELUAR",     // 21
  "ALASAN KELUAR",  // 22
  "KETERANGAN EXTRA",// 23
  "CHECKLIST",      // 24
  "TIMESTAMP OFFBOARDING" // 25
];

var STATUS_KONTRAK = ['Kontrak', 'Probation', 'Magang', 'Outsourcing'];

// ============================================================
// doGet()
// ============================================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('HRIS Dashboard - PT Japa Indotama')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================================
// simpanDataKaryawan() — Simpan karyawan BARU ke db_karyawan
// ============================================================
function simpanDataKaryawan(dataForm) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getOrCreateSheet(ss, SHEET_DB, HEADERS_DB, '#1a3a5c');

    var lastRow  = sheet.getLastRow();
    var newNo    = lastRow > 1 ? parseInt(sheet.getRange(lastRow, 1).getValue()) + 1 : 1;
    var now      = new Date();
    var tglInput = formatTimestamp(now);
    var usia     = hitungUsia(dataForm.tgllahir);
    var masaKerja= hitungMasaKerja(dataForm.tglmasuk);

    var tglLahirObj  = parseFlexDate(dataForm.tgllahir) || '';
    var tglMasukObj  = parseFlexDate(dataForm.tglmasuk) || '';
    var hubKerja     = dataForm.hubkerja || '';
    var akhirKontrak = '';
    if (STATUS_KONTRAK.indexOf(hubKerja) >= 0 && dataForm.akhirKontrak) {
      akhirKontrak = parseFlexDate(dataForm.akhirKontrak) || '';
    }

    var rowData = [
      newNo,
      dataForm.nip,
      (dataForm.nama || '').toUpperCase(),
      dataForm.departemen, dataForm.unit, dataForm.jabatan, dataForm.level || '',
      dataForm.nik, dataForm.notlp, dataForm.alamat, dataForm.hubkerja,
      dataForm.agama, tglLahirObj, usia, dataForm.pendidikan, dataForm.jeniskelamin,
      tglMasukObj, masaKerja, tglInput, tglInput, akhirKontrak
    ];

    sheet.appendRow(rowData);
    applyRowFormat(sheet, lastRow + 1, rowData.length);

    // Catat ke history sebagai entri "Masuk"
    simpanHistory({
      nip: dataForm.nip,
      nama: (dataForm.nama || '').toUpperCase(),
      tipe: 'Masuk',
      tglEfektif: dataForm.tglmasuk,
      keterangan: dataForm.hubkerja + ' · ' + dataForm.jabatan + ' · ' + dataForm.departemen
    });

    syncKontrakSheet();

    return {
      status: 'success',
      message: 'Data karyawan ' + dataForm.nama + ' berhasil disimpan! (NIP: ' + dataForm.nip + ')'
    };
  } catch (err) {
    return { status: 'error', message: err.message };
  }
}

// ============================================================
// getSemuaKaryawan() — HANYA karyawan di db_karyawan (semua aktif)
// Data yang sudah dioffboarding sudah pindah ke db_alumni,
// jadi sheet ini murisi karyawan aktif saja.
// ============================================================
function getSemuaKaryawan() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_DB);
    if (!sheet || sheet.getLastRow() <= 1) return [];

    var numCols = HEADERS_DB.length;
    var data    = sheet.getRange(2, 1, sheet.getLastRow() - 1, numCols).getValues();

    data = data.map(function(row) {
      while (row.length < numCols) row.push('');

      // Konversi Date objects → string DD/MM/YYYY
      row = row.map(function(cell) {
        if (cell instanceof Date && !isNaN(cell.getTime())) return formatTanggal(cell);
        return cell;
      });

      // Hitung ulang usia & masa kerja
      if (row[12]) { try { row[13] = hitungUsia(row[12]); } catch(e) {} }
      if (row[16]) { try { row[17] = hitungMasaKerja(row[16]); } catch(e) {} }

      return row;
    });

    // Filter baris kosong (NIP kosong)
    return data.filter(function(row) {
      return row[1] !== '' && row[1] !== null && row[1] !== undefined;
    });

  } catch (err) {
    console.error('getSemuaKaryawan ERROR: ' + err.message);
    return [];
  }
}

// ============================================================
// editDataKaryawan() — Update data karyawan aktif
// ============================================================
function editDataKaryawan(dataForm) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_DB);
    if (!sheet) return { status: 'error', message: 'Sheet db_karyawan tidak ditemukan.' };

    var rowNum = parseInt(dataForm.rowNum);
    if (isNaN(rowNum) || rowNum < 2) return { status: 'error', message: 'Nomor baris tidak valid.' };

    var now        = new Date();
    var tglUbah    = formatTimestamp(now);
    var usia       = hitungUsia(dataForm.tgllahir);
    var masaKerja  = hitungMasaKerja(dataForm.tglmasuk);
    var numCols    = HEADERS_DB.length;

    var existingRow = sheet.getRange(rowNum, 1, 1, numCols).getValues()[0];
    while (existingRow.length < numCols) existingRow.push('');

    var tglLahirObj  = parseFlexDate(dataForm.tgllahir) || '';
    var tglMasukObj  = parseFlexDate(dataForm.tglmasuk) || '';
    var hubKerja     = dataForm.hubkerja || '';
    var akhirKontrak = '';
    if (STATUS_KONTRAK.indexOf(hubKerja) >= 0 && dataForm.akhirKontrak) {
      akhirKontrak = parseFlexDate(dataForm.akhirKontrak) || '';
    }

    var updatedRow = [
      existingRow[0], existingRow[1],               // NO & NIP tetap
      (dataForm.nama || '').toUpperCase(),
      dataForm.departemen, dataForm.unit, dataForm.jabatan, dataForm.level || '',
      dataForm.nik, dataForm.notlp, dataForm.alamat, dataForm.hubkerja,
      dataForm.agama, tglLahirObj, usia, dataForm.pendidikan, dataForm.jeniskelamin,
      tglMasukObj, masaKerja,
      existingRow[18] || tglUbah,  // TANGGAL INPUT tetap
      tglUbah,                      // TERAKHIR DIUBAH
      akhirKontrak
    ];

    sheet.getRange(rowNum, 1, 1, updatedRow.length).setValues([updatedRow]);
    syncKontrakSheet();

    return { status: 'success', message: 'Data ' + dataForm.nama + ' berhasil diperbarui.' };
  } catch (err) {
    return { status: 'error', message: err.message };
  }
}

// ============================================================
// prosesOffboarding() — PINDAHKAN baris dari db_karyawan ke db_alumni
//
// Alur:
//   1. Baca baris karyawan dari db_karyawan
//   2. Susun baris alumni (data lengkap + info offboarding)
//   3. Append ke db_alumni
//   4. DELETE baris dari db_karyawan
//   5. Renumber kolom NO di db_karyawan
//   6. Catat ke db_history
//   7. Sync db_kontrak
// ============================================================
function prosesOffboarding(dataForm) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_DB);
    if (!sheet) return { status: 'error', message: 'Sheet db_karyawan tidak ditemukan.' };

    var rowNum = parseInt(dataForm.rowNum);
    if (isNaN(rowNum) || rowNum < 2) return { status: 'error', message: 'Nomor baris tidak valid.' };

    var numCols     = HEADERS_DB.length;
    var existingRow = sheet.getRange(rowNum, 1, 1, numCols).getValues()[0];
    while (existingRow.length < numCols) existingRow.push('');

    // Konversi Date objects di baris yang dibaca
    var cleanRow = existingRow.map(function(cell) {
      if (cell instanceof Date && !isNaN(cell.getTime())) return formatTanggal(cell);
      return cell;
    });

    var now       = new Date();
    var tsNow     = formatTimestamp(now);
    var tglKeluar = parseFlexDate(dataForm.tglKeluar) || now;

    // Hitung masa kerja total (dari tgl masuk hingga tgl keluar)
    var masaKerjaTotal = '';
    if (cleanRow[16]) {
      try {
        var tglMasukDate = parseFlexDate(cleanRow[16].toString());
        if (tglMasukDate) {
          var thn = tglKeluar.getFullYear() - tglMasukDate.getFullYear();
          var bln = tglKeluar.getMonth() - tglMasukDate.getMonth();
          if (bln < 0) { thn--; bln += 12; }
          if (tglKeluar.getDate() < tglMasukDate.getDate()) { bln--; if (bln < 0) { thn--; bln += 12; } }
          masaKerjaTotal = thn + ' Tahun ' + bln + ' Bulan';
        }
      } catch(e) {}
    }

    // Ringkasan checklist
    var cl     = dataForm.checklist || {};
    var clKeys = ['serahTerimaPekerjaan','pengembalianAset','clearanceIT',
                  'clearanceKeuangan','bpjsSelesai','suratKeluar',
                  'wawancaraKeluar','hakAkhirDibayar'];
    var clLabels = {
      serahTerimaPekerjaan:'Serah Terima', pengembalianAset:'Pengembalian Aset',
      clearanceIT:'Clearance IT',           clearanceKeuangan:'Clearance Keuangan',
      bpjsSelesai:'BPJS Selesai',          suratKeluar:'Surat Keluar',
      wawancaraKeluar:'Exit Interview',     hakAkhirDibayar:'Hak Akhir Dibayar'
    };
    var doneCt = 0;
    var clParts = clKeys.map(function(k) {
      var done = cl[k] === true || cl[k] === 'true' || cl[k] === 1;
      if (done) doneCt++;
      return (done ? '✓' : '✗') + ' ' + clLabels[k];
    });
    var checklistStr = doneCt + '/' + clKeys.length + ' — ' + clParts.join(' | ');

    // ---- 1. Siapkan sheet db_alumni ----
    var alumniSheet = getOrCreateSheet(ss, SHEET_ALUMNI, HEADERS_ALUMNI, '#374151');

    var alumniLastRow = alumniSheet.getLastRow();
    var alumniNo = alumniLastRow > 1
      ? parseInt(alumniSheet.getRange(alumniLastRow, 1).getValue()) + 1
      : 1;

    // Kolom 0–20 = salinan penuh dari db_karyawan
    // Kolom 21–25 = info offboarding
    var alumniRow = [
      alumniNo,          // 0  NO (di alumni)
      cleanRow[1],       // 1  NIP
      cleanRow[2],       // 2  NAMA
      cleanRow[3],       // 3  DEPARTEMEN
      cleanRow[4],       // 4  UNIT
      cleanRow[5],       // 5  JABATAN
      cleanRow[6],       // 6  LEVEL
      cleanRow[7],       // 7  NIK
      cleanRow[8],       // 8  NO TLP
      cleanRow[9],       // 9  ALAMAT
      cleanRow[10],      // 10 HUBUNGAN KERJA
      cleanRow[11],      // 11 AGAMA
      cleanRow[12],      // 12 TANGGAL LAHIR
      cleanRow[13],      // 13 USIA SAAT KELUAR
      cleanRow[14],      // 14 PENDIDIKAN
      cleanRow[15],      // 15 JENIS KELAMIN
      cleanRow[16],      // 16 TANGGAL MASUK
      masaKerjaTotal,    // 17 MASA KERJA TOTAL
      cleanRow[18],      // 18 TANGGAL INPUT (asli)
      dataForm.dicatatOleh || 'Admin HR', // 19 DICATAT OLEH
      cleanRow[20],      // 20 AKHIR KONTRAK (jika ada)
      tglKeluar,         // 21 TGL KELUAR (Date obj → locale-safe)
      dataForm.alasanKeluar, // 22 ALASAN KELUAR
      dataForm.keteranganExtra || '', // 23 KETERANGAN EXTRA
      checklistStr,      // 24 CHECKLIST
      tsNow              // 25 TIMESTAMP OFFBOARDING
    ];

    alumniSheet.appendRow(alumniRow);
    applyRowFormat(alumniSheet, alumniSheet.getLastRow(), HEADERS_ALUMNI.length);

    // Warnai baris alumni sesuai alasan
    var alasanBg = {
      'Resign':         '#fffbeb',
      'PHK':            '#fef2f2',
      'Pensiun':        '#eff6ff',
      'Kontrak Habis':  '#f5f3ff',
      'Meninggal Dunia':'#f3f4f6',
      'Lainnya':        '#f9fafb'
    };
    var bg = alasanBg[dataForm.alasanKeluar] || '#f9fafb';
    alumniSheet.getRange(alumniSheet.getLastRow(), 1, 1, HEADERS_ALUMNI.length).setBackground(bg);

    // ---- 2. Hapus baris dari db_karyawan ----
    sheet.deleteRow(rowNum);

    // ---- 3. Renumber kolom NO di db_karyawan ----
    renumberSheet(sheet);

    // ---- 4. Catat ke db_history ----
    simpanHistory({
      nip:        cleanRow[1],
      nama:       cleanRow[2],
      tipe:       'Keluar',
      tglEfektif: tglKeluar,
      keterangan: dataForm.alasanKeluar
                  + (dataForm.keteranganExtra ? ': ' + dataForm.keteranganExtra : '')
                  + ' | Checklist: ' + doneCt + '/' + clKeys.length
    });

    // ---- 5. Sync db_kontrak (baris sudah terhapus dari db_karyawan) ----
    syncKontrakSheet();

    return {
      status: 'success',
      message: cleanRow[2] + ' berhasil dipindahkan ke Alumni. '
               + 'Checklist: ' + doneCt + '/' + clKeys.length + ' selesai.'
    };
  } catch (err) {
    return { status: 'error', message: err.message };
  }
}

// ============================================================
// batalkanOffboarding() — Pindahkan kembali dari db_alumni ke db_karyawan
// Berguna jika terjadi kesalahan input offboarding
// ============================================================
function batalkanOffboarding(dataForm) {
  try {
    var ss          = SpreadsheetApp.getActiveSpreadsheet();
    var alumniSheet = ss.getSheetByName(SHEET_ALUMNI);
    if (!alumniSheet) return { status: 'error', message: 'Sheet db_alumni tidak ditemukan.' };

    var alumniRow = parseInt(dataForm.alumniRow);
    if (isNaN(alumniRow) || alumniRow < 2) return { status: 'error', message: 'Nomor baris alumni tidak valid.' };

    var numColsAlumni = HEADERS_ALUMNI.length;
    var aRow = alumniSheet.getRange(alumniRow, 1, 1, numColsAlumni).getValues()[0];
    while (aRow.length < numColsAlumni) aRow.push('');

    // Konversi Date objects
    var cleanARow = aRow.map(function(cell) {
      if (cell instanceof Date && !isNaN(cell.getTime())) return formatTanggal(cell);
      return cell;
    });

    // Ambil sheet db_karyawan, buat jika belum ada
    var dbSheet  = getOrCreateSheet(ss, SHEET_DB, HEADERS_DB, '#1a3a5c');
    var lastRow  = dbSheet.getLastRow();
    var newNo    = lastRow > 1 ? parseInt(dbSheet.getRange(lastRow, 1).getValue()) + 1 : 1;
    var now      = new Date();
    var tglUbah  = formatTimestamp(now);

    // Rekonstruksi baris db_karyawan dari kolom 0–20 di alumni
    // Hitung ulang usia & masa kerja agar segar
    var usia     = hitungUsia(cleanARow[12]);
    var masaKerja= hitungMasaKerja(cleanARow[16]);

    var newRow = [
      newNo,
      cleanARow[1],   // NIP
      cleanARow[2],   // NAMA
      cleanARow[3],   // DEPARTEMEN
      cleanARow[4],   // UNIT
      cleanARow[5],   // JABATAN
      cleanARow[6],   // LEVEL
      cleanARow[7],   // NIK
      cleanARow[8],   // NO TLP
      cleanARow[9],   // ALAMAT
      cleanARow[10],  // HUBUNGAN KERJA
      cleanARow[11],  // AGAMA
      parseFlexDate(cleanARow[12]) || cleanARow[12], // TANGGAL LAHIR (Date obj)
      usia,
      cleanARow[14],  // PENDIDIKAN
      cleanARow[15],  // JENIS KELAMIN
      parseFlexDate(cleanARow[16]) || cleanARow[16], // TANGGAL MASUK (Date obj)
      masaKerja,
      cleanARow[18],  // TANGGAL INPUT ASLI
      tglUbah,        // TERAKHIR DIUBAH
      parseFlexDate(cleanARow[20]) || ''  // AKHIR KONTRAK (Date obj jika ada)
    ];

    dbSheet.appendRow(newRow);
    applyRowFormat(dbSheet, dbSheet.getLastRow(), newRow.length);

    // Hapus dari db_alumni
    alumniSheet.deleteRow(alumniRow);
    renumberSheet(alumniSheet);

    // Catat ke history
    simpanHistory({
      nip:        cleanARow[1],
      nama:       cleanARow[2],
      tipe:       'Reaktivasi',
      tglEfektif: formatTanggal(now),
      keterangan: 'Offboarding dibatalkan — ' + (dataForm.alasan || 'Tidak ada keterangan')
    });

    syncKontrakSheet();

    return { status: 'success', message: cleanARow[2] + ' berhasil dipindahkan kembali ke Data Karyawan.' };
  } catch (err) {
    return { status: 'error', message: err.message };
  }
}

// ============================================================
// getDataAlumni() — Ambil semua data dari db_alumni
// ============================================================
function getDataAlumni() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_ALUMNI);
    if (!sheet || sheet.getLastRow() <= 1) return [];

    var numCols = HEADERS_ALUMNI.length;
    var data    = sheet.getRange(2, 1, sheet.getLastRow() - 1, numCols).getValues();

    data = data.map(function(row) {
      while (row.length < numCols) row.push('');
      return row.map(function(cell) {
        if (cell instanceof Date && !isNaN(cell.getTime())) return formatTanggal(cell);
        return cell;
      });
    });

    // Filter baris kosong
    return data.filter(function(row) { return row[1] !== '' && row[1] !== null; });

  } catch (err) {
    return [];
  }
}

// ============================================================
// simpanHistory() — Catat ke db_history
// Tipe: Masuk · Keluar · Pengangkatan Tetap · Mutasi · Promosi · Demosi · Reaktivasi
// ============================================================
function simpanHistory(dataForm) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getOrCreateSheet(ss, SHEET_HISTORY, HEADERS_HIST, '#132a47');

    var tglDate = parseFlexDate(dataForm.tglEfektif) || dataForm.tglEfektif || '';
    var rowData = [
      dataForm.nip, dataForm.nama, dataForm.tipe,
      tglDate,
      dataForm.keterangan || '',
      'Admin HR',
      formatTimestamp(new Date())
    ];

    sheet.appendRow(rowData);
    applyRowFormat(sheet, sheet.getLastRow(), rowData.length);
    return { status: 'success', message: 'Perubahan berhasil dicatat.' };
  } catch (err) {
    return { status: 'error', message: err.message };
  }
}

// ============================================================
// catatPerubahan() — BARU v3.2
// Catat perubahan kepegawaian DAN otomatis update db_karyawan
//
// dataForm fields:
//   tipe        : 'Pengangkatan Tetap' | 'Mutasi' | 'Promosi' | 'Demosi'
//   nip         : NIP karyawan
//   rowNum      : nomor baris di db_karyawan (2-based)
//   tglEfektif  : YYYY-MM-DD
//   noSK        : nomor SK/surat keputusan (opsional)
//
//   -- Pengangkatan Tetap --
//   newDept, newUnit, newJabatan, newLevel
//   (Hubungan Kerja otomatis → Permanent, Akhir Kontrak dihapus)
//
//   -- Mutasi --
//   newDept, newUnit, newJabatan (jabatan bisa sama)
//   alasanMutasi
//
//   -- Promosi --
//   newJabatan, newLevel
//   alasanPromosi
//
//   -- Demosi --
//   newJabatan, newLevel
//   alasanDemosi
// ============================================================
function catatPerubahan(dataForm) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_DB);
    if (!sheet) return { status: 'error', message: 'Sheet db_karyawan tidak ditemukan.' };

    var rowNum = parseInt(dataForm.rowNum);
    if (isNaN(rowNum) || rowNum < 2) return { status: 'error', message: 'Nomor baris tidak valid.' };

    var numCols = HEADERS_DB.length;
    var existing = sheet.getRange(rowNum, 1, 1, numCols).getValues()[0];
    while (existing.length < numCols) existing.push('');

    // Konversi Date objects di baris lama
    var clean = existing.map(function(cell) {
      if (cell instanceof Date && !isNaN(cell.getTime())) return formatTanggal(cell);
      return cell;
    });

    var now     = new Date();
    var tglUbah = formatTimestamp(now);
    var tipe    = dataForm.tipe || '';

    // ── Nilai lama (untuk teks history) ──
    var oldDept    = clean[3] || '-';
    var oldUnit    = clean[4] || '-';
    var oldJabatan = clean[5] || '-';
    var oldLevel   = clean[6] || '-';
    var oldHub     = clean[10] || '-';

    // ── Nilai baru (default = nilai lama) ──
    var newDept    = dataForm.newDept    || clean[3] || '';
    var newUnit    = dataForm.newUnit    || clean[4] || '';
    var newJabatan = dataForm.newJabatan || clean[5] || '';
    var newLevel   = dataForm.newLevel   || clean[6] || '';
    var newHub     = clean[10] || '';           // Hubungan Kerja tetap kecuali Angkat Tetap
    var newAkhir   = existing[20];              // Akhir Kontrak tetap kecuali Angkat Tetap

    // ── Logika per tipe ──
    var keteranganHistory = '';
    var noSK = dataForm.noSK ? ' | No. SK: ' + dataForm.noSK : '';

    if (tipe === 'Pengangkatan Tetap') {
      newHub     = 'Permanent';
      newAkhir   = '';             // Hapus akhir kontrak
      newDept    = dataForm.newDept    || clean[3] || '';
      newUnit    = dataForm.newUnit    || clean[4] || '';
      newJabatan = dataForm.newJabatan || clean[5] || '';
      newLevel   = dataForm.newLevel   || clean[6] || '';
      keteranganHistory = oldHub + ' → Permanent'
        + ' | Jabatan: ' + oldJabatan + ' → ' + newJabatan
        + ' | Level: ' + oldLevel + ' → ' + newLevel
        + ' | Dept: ' + newDept + noSK;
    }
    else if (tipe === 'Mutasi') {
      keteranganHistory = 'Dept: ' + oldDept + ' → ' + newDept
        + ' | Unit: ' + oldUnit + ' → ' + newUnit
        + ' | Jabatan: ' + oldJabatan + ' → ' + newJabatan
        + (dataForm.alasanMutasi ? ' | Alasan: ' + dataForm.alasanMutasi : '')
        + noSK;
    }
    else if (tipe === 'Promosi') {
      keteranganHistory = 'Jabatan: ' + oldJabatan + ' → ' + newJabatan
        + ' | Level: ' + oldLevel + ' → ' + newLevel
        + (dataForm.alasanPromosi ? ' | ' + dataForm.alasanPromosi : '')
        + noSK;
    }
    else if (tipe === 'Demosi') {
      keteranganHistory = 'Jabatan: ' + oldJabatan + ' → ' + newJabatan
        + ' | Level: ' + oldLevel + ' → ' + newLevel
        + (dataForm.alasanDemosi ? ' | Alasan: ' + dataForm.alasanDemosi : '')
        + noSK;
    }

    // ── Update kolom yang berubah di db_karyawan ──
    // Kolom 4 = DEPARTEMEN, 5 = UNIT, 6 = JABATAN, 7 = LEVEL,
    // 11 = HUBUNGAN KERJA, 20 = TERAKHIR DIUBAH, 21 = AKHIR KONTRAK

    // Departemen (col 4 → index 3 → sheet col 4)
    sheet.getRange(rowNum, 4).setValue(newDept);
    // Unit (col 5)
    sheet.getRange(rowNum, 5).setValue(newUnit);
    // Jabatan (col 6)
    sheet.getRange(rowNum, 6).setValue(newJabatan);
    // Level (col 7)
    sheet.getRange(rowNum, 7).setValue(newLevel);
    // Hubungan Kerja (col 11)
    sheet.getRange(rowNum, 11).setValue(newHub);
    // Akhir Kontrak (col 21)
    if (tipe === 'Pengangkatan Tetap') {
      sheet.getRange(rowNum, 21).setValue('');   // Hapus akhir kontrak
    }
    // Terakhir Diubah (col 20)
    sheet.getRange(rowNum, 20).setValue(tglUbah);

    // ── Catat ke db_history ──
    simpanHistory({
      nip:        clean[1],
      nama:       clean[2],
      tipe:       tipe,
      tglEfektif: dataForm.tglEfektif,
      keterangan: keteranganHistory
    });

    // ── Sync db_kontrak (Angkat Tetap memindahkan dari non-permanent ke permanent) ──
    syncKontrakSheet();

    return {
      status: 'success',
      message: tipe + ' untuk ' + clean[2] + ' berhasil dicatat & data karyawan telah diperbarui.',
      updatedData: {
        departemen:   newDept,
        unit:         newUnit,
        jabatan:      newJabatan,
        level:        newLevel,
        hubunganKerja:newHub
      }
    };
  } catch (err) {
    return { status: 'error', message: err.message };
  }
}

// ============================================================
// getHistory() — Ambil riwayat berdasarkan NIP
// ============================================================
function getHistory(nip) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_HISTORY);
    if (!sheet || sheet.getLastRow() <= 1) return [];

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS_HIST.length).getValues();

    var result = data.filter(function(row) {
      return row[0].toString() === nip.toString();
    });

    result.sort(function(a, b) {
      var da = parseFlexDate(a[3]), db = parseFlexDate(b[3]);
      if (!da && !db) return 0;
      if (!da) return -1;
      if (!db) return 1;
      return da - db;
    });

    return result.map(function(row) {
      return row.map(function(cell) {
        if (cell instanceof Date && !isNaN(cell.getTime())) return formatTanggal(cell);
        return cell;
      });
    });
  } catch (err) {
    return [];
  }
}

// ============================================================
// getDashboardStats() — Ambil statistik ringkas untuk dashboard
// ============================================================
function getDashboardStats() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var dbSheet = ss.getSheetByName(SHEET_DB);
    var total   = dbSheet && dbSheet.getLastRow() > 1 ? dbSheet.getLastRow() - 1 : 0;
    var perm = 0, kontrak = 0, dept = {};

    if (dbSheet && total > 0) {
      var cols = dbSheet.getRange(2, 1, total, HEADERS_DB.length).getValues();
      cols.forEach(function(row) {
        if (!row[1]) return;
        if (row[10] === 'Permanent') perm++;
        if (STATUS_KONTRAK.indexOf(row[10]) >= 0) kontrak++;
        if (row[3]) dept[row[3]] = 1;
      });
    }

    var alumniSheet = ss.getSheetByName(SHEET_ALUMNI);
    var alumni = alumniSheet && alumniSheet.getLastRow() > 1 ? alumniSheet.getLastRow() - 1 : 0;

    return {
      total:   total,
      perm:    perm,
      kontrak: kontrak,
      dept:    Object.keys(dept).length,
      alumni:  alumni
    };
  } catch (err) {
    return { total: 0, perm: 0, kontrak: 0, dept: 0, alumni: 0 };
  }
}

// ============================================================
// syncKontrakSheet() — Perbarui sheet db_kontrak
// Hanya membaca dari db_karyawan (sudah 100% aktif)
// ============================================================
function syncKontrakSheet() {
  try {
    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var dbSheet = ss.getSheetByName(SHEET_DB);
    if (!dbSheet || dbSheet.getLastRow() <= 1) return;

    var kontrakSheet = getOrCreateSheet(ss, SHEET_KONTRAK, HEADERS_KONTRAK, '#7c3aed');
    var lastExisting = kontrakSheet.getLastRow();
    if (lastExisting > 1) {
      kontrakSheet.getRange(2, 1, lastExisting - 1, HEADERS_KONTRAK.length).clear();
    }

    var numCols = HEADERS_DB.length;
    var allData = dbSheet.getRange(2, 1, dbSheet.getLastRow() - 1, numCols).getValues();

    var kontrakData = allData.filter(function(row) {
      while (row.length < numCols) row.push('');
      return STATUS_KONTRAK.indexOf((row[10] || '').toString()) >= 0;
    });

    if (!kontrakData.length) return;

    var now   = new Date();
    var today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    var upd   = formatTimestamp(now);

    kontrakData.sort(function(a, b) {
      var da = parseFlexDate(a[20]) || null;
      var db = parseFlexDate(b[20]) || null;
      if (!da && !db) return 0;
      if (!da) return 1;
      if (!db) return -1;
      return da - db;
    });

    var rows = kontrakData.map(function(row, idx) {
      var akhirStr  = cellToDateStr(row[20]);
      var tglMasuk  = cellToDateStr(row[16]);
      var akhirDate = akhirStr ? parseFlexDate(akhirStr) : null;
      var sisaHari  = '-', statusPkwt = 'BELUM DIISI';

      if (akhirDate) {
        var days = Math.ceil((akhirDate.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
        sisaHari = days;
        if      (days < 0)   statusPkwt = 'SUDAH BERAKHIR';
        else if (days <= 30) statusPkwt = 'KRITIS';
        else if (days <= 60) statusPkwt = 'SEGERA';
        else if (days <= 90) statusPkwt = 'PERHATIAN';
        else                 statusPkwt = 'AMAN';
      }

      return [
        idx + 1, row[1], row[2], row[3], row[4], row[5], row[6],
        row[10], tglMasuk || '', akhirStr || 'Belum diisi',
        sisaHari, statusPkwt, upd
      ];
    });

    if (rows.length > 0) {
      kontrakSheet.getRange(2, 1, rows.length, HEADERS_KONTRAK.length).setValues(rows);
      rows.forEach(function(row, i) {
        var st = row[11];
        var bg = '#ffffff', tc = '#1f2937';
        if      (st === 'KRITIS')         { bg = '#fef2f2'; tc = '#dc2626'; }
        else if (st === 'SEGERA')         { bg = '#fffbeb'; tc = '#d97706'; }
        else if (st === 'PERHATIAN')      { bg = '#fefce8'; tc = '#ca8a04'; }
        else if (st === 'AMAN')           { bg = '#ecfdf5'; tc = '#059669'; }
        else if (st === 'SUDAH BERAKHIR') { bg = '#f3f4f6'; tc = '#6b7280'; }
        else if (st === 'BELUM DIISI')    { bg = '#f8fafc'; tc = '#94a3b8'; }
        kontrakSheet.getRange(i + 2, 1, 1, HEADERS_KONTRAK.length).setBackground(bg);
        kontrakSheet.getRange(i + 2, 12).setFontColor(tc).setFontWeight('bold');
        if (row[10] !== '-') kontrakSheet.getRange(i + 2, 11).setFontColor(tc).setFontWeight('bold');
        applyRowFormat(kontrakSheet, i + 2, HEADERS_KONTRAK.length);
      });
      kontrakSheet.autoResizeColumns(1, HEADERS_KONTRAK.length);
    }
  } catch (e) {
    console.error('syncKontrakSheet ERROR: ' + e.message);
  }
}

function syncKontrakManual() {
  syncKontrakSheet();
  SpreadsheetApp.getUi().alert('✅ Sheet db_kontrak berhasil diperbarui!');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ HRIS Tools')
    .addItem('🔄 Sinkronisasi db_kontrak', 'syncKontrakManual')
    .addToUi();
}

// ============================================================
// renumberSheet() — Nomor ulang kolom NO setelah baris dihapus
// ============================================================
function renumberSheet(sheet) {
  try {
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;
    for (var i = 2; i <= lastRow; i++) {
      sheet.getRange(i, 1).setValue(i - 1);
    }
  } catch(e) {
    console.error('renumberSheet ERROR: ' + e.message);
  }
}

// ============================================================
//  HELPERS
// ============================================================
function cellToDateStr(val) {
  if (!val) return '';
  if (val instanceof Date) return isNaN(val.getTime()) ? '' : formatTanggal(val);
  var str = val.toString().trim();
  if (!str) return '';
  var d = parseFlexDate(str);
  return d ? formatTanggal(d) : str;
}

function getOrCreateSheet(ss, name, headers, bg) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    var hr = sheet.getRange(1, 1, 1, headers.length);
    hr.setBackground(bg || '#1a3a5c').setFontColor('#ffffff')
      .setFontWeight('bold').setHorizontalAlignment('center').setFontSize(9);
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(2);
    sheet.autoResizeColumns(1, headers.length);
  }
  return sheet;
}

function applyRowFormat(sheet, rowNum, colCount) {
  try {
    sheet.getRange(rowNum, 1, 1, colCount)
         .setBorder(true,true,true,true,true,true,'#e2e8f0',SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(rowNum, 2).setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange(rowNum, 1).setHorizontalAlignment('center');
  } catch(e) {}
}

function hitungUsia(str) {
  if (!str) return '';
  var d = parseFlexDate(str.toString());
  if (!d) return '';
  var now = new Date();
  var usia = now.getFullYear() - d.getFullYear();
  var bln  = now.getMonth() - d.getMonth();
  if (bln < 0 || (bln === 0 && now.getDate() < d.getDate())) usia--;
  return usia + ' Tahun';
}

function hitungMasaKerja(str) {
  if (!str) return '';
  var d = parseFlexDate(str.toString());
  if (!d) return '';
  var now = new Date();
  var thn = now.getFullYear() - d.getFullYear();
  var bln = now.getMonth() - d.getMonth();
  if (bln < 0) { thn--; bln += 12; }
  if (now.getDate() < d.getDate()) { bln--; if (bln < 0) { thn--; bln += 12; } }
  return thn + ' Tahun ' + bln + ' Bulan';
}

function parseFlexDate(str) {
  if (!str) return null;
  if (str instanceof Date) return isNaN(str.getTime()) ? null : str;
  str = str.toString().trim();
  if (str.length < 6) return null;
  var m1 = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m1) return new Date(parseInt(m1[3]), parseInt(m1[2])-1, parseInt(m1[1]));
  var m2 = str.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m2) return new Date(parseInt(m2[1]), parseInt(m2[2])-1, parseInt(m2[3]));
  var d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}

function formatTanggal(str) {
  if (!str) return '';
  var d = parseFlexDate(str);
  if (!d) return str.toString();
  return d.getDate().toString().padStart(2,'0') + '/'
       + (d.getMonth()+1).toString().padStart(2,'0') + '/'
       + d.getFullYear();
}

function formatTimestamp(date) {
  return Utilities.formatDate(date, 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
}

// ================================================================
//  FASE 2: REKRUTMEN & ATS
//  Fungsi-fungsi di bawah ini menangani MPP dan pipeline kandidat
// ================================================================

// ── HELPER: Generate ID unik ──
function generateMppId(sheet) {
  var now    = new Date();
  var prefix = 'MPP-' + now.getFullYear() + (now.getMonth()+1).toString().padStart(2,'0') + '-';
  var lastRow = sheet.getLastRow();
  var seq = 1;
  if (lastRow > 1) {
    var ids = sheet.getRange(2, 1, lastRow-1, 1).getValues();
    var maxSeq = 0;
    ids.forEach(function(r) {
      var id = r[0].toString();
      if (id.indexOf('MPP-') === 0) {
        var n = parseInt(id.split('-').pop()) || 0;
        if (n > maxSeq) maxSeq = n;
      }
    });
    seq = maxSeq + 1;
  }
  return prefix + seq.toString().padStart(3,'0');
}

function generateKandidatId(sheet) {
  var now    = new Date();
  var prefix = 'KND-' + now.getFullYear() + (now.getMonth()+1).toString().padStart(2,'0') + '-';
  var lastRow = sheet.getLastRow();
  var seq = 1;
  if (lastRow > 1) {
    var ids = sheet.getRange(2, 1, lastRow-1, 1).getValues();
    var maxSeq = 0;
    ids.forEach(function(r) {
      var id = r[0].toString();
      if (id.indexOf('KND-') === 0) {
        var n = parseInt(id.split('-').pop()) || 0;
        if (n > maxSeq) maxSeq = n;
      }
    });
    seq = maxSeq + 1;
  }
  return prefix + seq.toString().padStart(3,'0');
}

// ────────────────────────────────────────────────────────────────
// buatMPP() — Buat permintaan rekrutmen baru
// ────────────────────────────────────────────────────────────────
function buatMPP(dataForm) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getOrCreateSheet(ss, SHEET_MPP, HEADERS_MPP, '#1a3a5c');

    var mppId = generateMppId(sheet);
    var now   = formatTimestamp(new Date());
    var tglBuka = parseFlexDate(dataForm.tglBuka) || new Date();

    var row = [
      mppId,
      tglBuka,
      dataForm.departemen   || '',
      dataForm.unit         || '',
      dataForm.posisi       || '',
      dataForm.level        || '',
      dataForm.hubKerja     || '',
      parseInt(dataForm.kuota) || 1,
      dataForm.alasan       || '',
      parseFlexDate(dataForm.targetSelesai) || '',
      'Open',
      0,                      // HIRED: 0 awalnya
      dataForm.catatan      || '',
      'Admin HR',
      now
    ];

    sheet.appendRow(row);
    applyRowFormat(sheet, sheet.getLastRow(), HEADERS_MPP.length);
    // Highlight baris baru dengan warna hijau muda
    sheet.getRange(sheet.getLastRow(), 11)
         .setBackground('#ecfdf5').setFontColor('#059669').setFontWeight('bold');

    return { status: 'success', message: 'MPP ' + mppId + ' berhasil dibuat untuk posisi ' + dataForm.posisi + '.', mppId: mppId };
  } catch(err) {
    return { status: 'error', message: err.message };
  }
}

// ────────────────────────────────────────────────────────────────
// getSemuaMPP() — Ambil semua MPP untuk ditampilkan di tabel
// ────────────────────────────────────────────────────────────────
function getSemuaMPP() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_MPP);
    if (!sheet || sheet.getLastRow() <= 1) return [];

    var data = sheet.getRange(2, 1, sheet.getLastRow()-1, HEADERS_MPP.length).getValues();

    return data.map(function(row) {
      return row.map(function(cell) {
        if (cell instanceof Date && !isNaN(cell.getTime())) return formatTanggal(cell);
        return cell;
      });
    }).filter(function(row) { return row[0] !== ''; });
  } catch(err) { return []; }
}

// ────────────────────────────────────────────────────────────────
// updateStatusMPP() — Ubah status MPP (Open/On Progress/Closed/Cancelled)
// ────────────────────────────────────────────────────────────────
function updateStatusMPP(dataForm) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_MPP);
    if (!sheet) return { status: 'error', message: 'Sheet db_mpp tidak ditemukan.' };

    var rowNum = parseInt(dataForm.rowNum);
    if (isNaN(rowNum) || rowNum < 2) return { status: 'error', message: 'Nomor baris tidak valid.' };

    sheet.getRange(rowNum, 11).setValue(dataForm.statusBaru);

    // Warnai kolom status sesuai nilai
    var colors = { 'Open':'#ecfdf5', 'On Progress':'#eff6ff', 'Closed':'#f3f4f6', 'Cancelled':'#fef2f2' };
    var txtColors = { 'Open':'#059669', 'On Progress':'#2563eb', 'Closed':'#6b7280', 'Cancelled':'#dc2626' };
    var bg  = colors[dataForm.statusBaru]    || '#f9fafb';
    var tc  = txtColors[dataForm.statusBaru] || '#374151';
    sheet.getRange(rowNum, 11).setBackground(bg).setFontColor(tc).setFontWeight('bold');

    return { status: 'success', message: 'Status MPP berhasil diubah ke ' + dataForm.statusBaru };
  } catch(err) {
    return { status: 'error', message: err.message };
  }
}

// ────────────────────────────────────────────────────────────────
// tambahKandidat() — Tambah pelamar baru ke pipeline MPP
// ────────────────────────────────────────────────────────────────
function tambahKandidat(dataForm) {
  try {
    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var sheet   = getOrCreateSheet(ss, SHEET_KANDIDAT, HEADERS_KANDIDAT, '#132a47');
    var mppSheet= ss.getSheetByName(SHEET_MPP);

    var kndId   = generateKandidatId(sheet);
    var now     = formatTimestamp(new Date());
    var tglDftr = parseFlexDate(dataForm.tglDaftar) || new Date();
    var tglLahirK = parseFlexDate(dataForm.tglLahir)  || '';

    // Ambil info posisi dari MPP
    var posisi = dataForm.posisi || '';
    var dept   = dataForm.departemen || '';

    // Baris kandidat — kolom 12–34 (tahapan) dikosongkan dulu
    var row = [
      kndId,
      dataForm.mppId       || '',
      posisi,
      dept,
      (dataForm.nama       || '').toUpperCase(),
      dataForm.noHp        || '',
      dataForm.email       || '',
      dataForm.sumber      || '',
      dataForm.linkCv      || '',
      tglDftr,
      'Screening',           // TAHAPAN awal
      'Proses',              // STATUS awal
      // Kolom 12–34: kosong semua
      '','','',  // Screening
      '','','','', // Tes Tertulis
      '','','',  // Psikotes
      '','','',  // ITV HRD
      '','','',  // ITV User
      '','','','', // Offering
      '',        // TGL HIRED
      now,       // TERAKHIR DIUPDATE
      now,        // TIMESTAMP INPUT
      tglLahirK  // TGL LAHIR
    ];

    sheet.appendRow(row);
    applyRowFormat(sheet, sheet.getLastRow(), HEADERS_KANDIDAT.length);

    return {
      status: 'success',
      message: 'Kandidat ' + dataForm.nama + ' berhasil ditambahkan ke pipeline ' + dataForm.mppId + '.',
      kndId: kndId
    };
  } catch(err) {
    return { status: 'error', message: err.message };
  }
}

// ────────────────────────────────────────────────────────────────
// getKandidatByMPP() — Ambil semua kandidat untuk satu MPP
// ────────────────────────────────────────────────────────────────
function getKandidatByMPP(mppId) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_KANDIDAT);
    if (!sheet || sheet.getLastRow() <= 1) return [];

    var numCols = HEADERS_KANDIDAT.length;
    var data    = sheet.getRange(2, 1, sheet.getLastRow()-1, numCols).getValues();

    return data
      .filter(function(row) { return row[1].toString() === mppId.toString(); })
      .map(function(row) {
        return row.map(function(cell) {
          if (cell instanceof Date && !isNaN(cell.getTime())) return formatTanggal(cell);
          return cell;
        });
      });
  } catch(err) { return []; }
}

// ────────────────────────────────────────────────────────────────
// updateTahapan() — Geser kandidat ke tahapan berikutnya +
//                   simpan hasil tahapan yang baru selesai
//
// TAHAPAN (urutan): Screening → Tes Tertulis → Psikotes →
//                  Interview HRD → Interview User → Offering → Hired
//
// dataForm:
//   rowNum         : baris di db_kandidat
//   tahapanSelesai : tahapan yang baru selesai diisi hasilnya
//   hasil          : 'Lanjut' | 'Tidak Lanjut' | 'Diterima' | 'Ditolak' | 'Negosiasi'
//   tglPelaksanaan : YYYY-MM-DD
//   nilaiTes       : angka (khusus tes tertulis)
//   catatan        : teks bebas
//   gaji           : angka (khusus offering)
// ────────────────────────────────────────────────────────────────
function updateTahapan(dataForm) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_KANDIDAT);
    if (!sheet) return { status: 'error', message: 'Sheet db_kandidat tidak ditemukan.' };

    var rowNum = parseInt(dataForm.rowNum);
    if (isNaN(rowNum) || rowNum < 2) return { status: 'error', message: 'Nomor baris tidak valid.' };

    var now    = formatTimestamp(new Date());
    var tglP   = parseFlexDate(dataForm.tglPelaksanaan) || new Date();
    var hasil  = dataForm.hasil  || '';
    var catatan= dataForm.catatan|| '';
    var lanjut = (hasil === 'Lanjut' || hasil === 'Diterima' || hasil === 'Negosiasi');

    // Mapping tahapan → kolom mulai (0-based) di HEADERS_KANDIDAT
    // Screening=12, TesTortulis=15, Psikotes=19, ItvHrd=22, ItvUser=25, Offering=28
    var TAHAPAN_MAP = {
      'Screening':    { startCol: 13, jumlahCol: 3,  tglCol: 13, hasilCol: 14, catatanCol: 15 },
      'Tes Tertulis': { startCol: 16, jumlahCol: 4,  tglCol: 16, nilaiCol: 17, hasilCol: 18, catatanCol: 19 },
      'Psikotes':     { startCol: 20, jumlahCol: 3,  tglCol: 20, hasilCol: 21, catatanCol: 22 },
      'Interview HRD':{ startCol: 23, jumlahCol: 3,  tglCol: 23, hasilCol: 24, catatanCol: 25 },
      'Interview User':{ startCol: 26, jumlahCol: 3, tglCol: 26, hasilCol: 27, catatanCol: 28 },
      'Offering':     { startCol: 29, jumlahCol: 4,  tglCol: 29, gajiCol: 30, hasilCol: 31, catatanCol: 32 }
    };

    // Urutan tahapan untuk menentukan tahapan BERIKUTNYA
    var URUTAN = ['Screening','Tes Tertulis','Psikotes','Interview HRD','Interview User','Offering'];

    var t = dataForm.tahapanSelesai;
    var tm = TAHAPAN_MAP[t];
    if (!tm) return { status: 'error', message: 'Tahapan tidak dikenal: ' + t };

    // Update kolom hasil tahapan yang selesai
    sheet.getRange(rowNum, tm.tglCol).setValue(tglP);
    sheet.getRange(rowNum, tm.hasilCol).setValue(hasil);
    sheet.getRange(rowNum, tm.catatanCol).setValue(catatan);

    if (t === 'Tes Tertulis' && tm.nilaiCol && dataForm.nilaiTes !== undefined) {
      sheet.getRange(rowNum, tm.nilaiCol).setValue(parseFloat(dataForm.nilaiTes) || 0);
    }
    if (t === 'Offering' && tm.gajiCol && dataForm.gaji) {
      sheet.getRange(rowNum, tm.gajiCol).setValue(dataForm.gaji);
    }

    // Tentukan tahapan & status berikutnya
    var tahapanBaru = t;
    var statusBaru  = 'Proses';

    if (!lanjut) {
      // Tidak lanjut → Reject
      statusBaru  = 'Reject';
      tahapanBaru = t;  // tetap di tahapan ini
    } else {
      var idx = URUTAN.indexOf(t);
      if (idx >= 0 && idx < URUTAN.length - 1) {
        tahapanBaru = URUTAN[idx + 1];
        statusBaru  = 'Proses';
      } else if (t === 'Offering' && (hasil === 'Diterima')) {
        // Hired!
        tahapanBaru = 'Hired';
        statusBaru  = 'Hired';
        sheet.getRange(rowNum, 33).setValue(tglP);  // TGL HIRED

        // Update jumlah HIRED di db_mpp
        var mppId = sheet.getRange(rowNum, 2).getValue();
        updateHiredCountMPP(mppId);
      } else if (t === 'Offering' && hasil === 'Ditolak') {
        statusBaru = 'Withdraw';
      }
    }

    // Update kolom TAHAPAN (11) dan STATUS (12) dan TERAKHIR DIUPDATE (34)
    sheet.getRange(rowNum, 11).setValue(tahapanBaru);
    sheet.getRange(rowNum, 12).setValue(statusBaru);
    sheet.getRange(rowNum, 34).setValue(now);

    // Warnai status
    var stBg = { 'Proses':'#eff6ff','Hired':'#ecfdf5','Reject':'#fef2f2','Withdraw':'#f3f4f6' };
    var stTc = { 'Proses':'#2563eb','Hired':'#059669','Reject':'#dc2626','Withdraw':'#6b7280' };
    sheet.getRange(rowNum, 12)
         .setBackground(stBg[statusBaru] || '#f9fafb')
         .setFontColor(stTc[statusBaru] || '#374151')
         .setFontWeight('bold');

    return {
      status: 'success',
      message: 'Hasil ' + t + ' berhasil disimpan. Tahapan berikutnya: ' + tahapanBaru + '.',
      tahapanBaru: tahapanBaru,
      statusBaru:  statusBaru
    };
  } catch(err) {
    return { status: 'error', message: err.message };
  }
}

// ────────────────────────────────────────────────────────────────
// updateHiredCountMPP() — Update kolom HIRED di db_mpp
// ────────────────────────────────────────────────────────────────
function updateHiredCountMPP(mppId) {
  try {
    var ss         = SpreadsheetApp.getActiveSpreadsheet();
    var mppSheet   = ss.getSheetByName(SHEET_MPP);
    var kndSheet   = ss.getSheetByName(SHEET_KANDIDAT);
    if (!mppSheet || !kndSheet) return;

    // Hitung jumlah kandidat Hired untuk MPP ini
    var kndData = kndSheet.getRange(2, 1, kndSheet.getLastRow()-1, 12).getValues();
    var hiredCt = kndData.filter(function(r) {
      return r[1].toString() === mppId.toString() && r[11].toString() === 'Hired';
    }).length;

    // Cari baris MPP
    var mppData = mppSheet.getRange(2, 1, mppSheet.getLastRow()-1, 1).getValues();
    for (var i = 0; i < mppData.length; i++) {
      if (mppData[i][0].toString() === mppId.toString()) {
        var mppRow = i + 2;
        mppSheet.getRange(mppRow, 12).setValue(hiredCt);  // Kolom HIRED

        // Auto-close MPP jika hired >= kuota
        var kuota = parseInt(mppSheet.getRange(mppRow, 8).getValue()) || 1;
        if (hiredCt >= kuota) {
          mppSheet.getRange(mppRow, 11).setValue('Closed')
                  .setBackground('#f3f4f6').setFontColor('#6b7280').setFontWeight('bold');
        } else {
          mppSheet.getRange(mppRow, 11).setValue('On Progress')
                  .setBackground('#eff6ff').setFontColor('#2563eb').setFontWeight('bold');
        }
        break;
      }
    }
  } catch(e) {
    console.error('updateHiredCountMPP ERROR: ' + e.message);
  }
}

// ────────────────────────────────────────────────────────────────
// hiredKeKaryawan() — Konversi kandidat Hired → pre-fill form karyawan
// Mengembalikan data yang siap diisi ke form input karyawan
// ────────────────────────────────────────────────────────────────
function hiredKeKaryawan(dataForm) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_KANDIDAT);
    if (!sheet) return { status: 'error', message: 'Sheet db_kandidat tidak ditemukan.' };

    var rowNum = parseInt(dataForm.rowNum);
    var numCols= HEADERS_KANDIDAT.length;
    var row    = sheet.getRange(rowNum, 1, 1, numCols).getValues()[0];

    // Kembalikan data pre-fill untuk form Input Karyawan
    return {
      status: 'success',
      prefill: {
        nama:       row[4] || '',   // NAMA KANDIDAT
        noTlp:      row[5] || '',   // NO HP
        email:      row[6] || '',   // EMAIL
        departemen: row[3] || '',   // DEPARTEMEN
        jabatan:    row[2] || '',   // POSISI
        hubkerja:   '',             // diisi manual (kontrak/permanent)
        kndId:      row[0] || '',
        mppId:      row[1] || ''
      }
    };
  } catch(err) {
    return { status: 'error', message: err.message };
  }
}

// ────────────────────────────────────────────────────────────────
// getRekrutmenStats() — Statistik cepat untuk dashboard Fase 2
// ────────────────────────────────────────────────────────────────
function getRekrutmenStats() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();

    var mppSheet = ss.getSheetByName(SHEET_MPP);
    var kndSheet = ss.getSheetByName(SHEET_KANDIDAT);

    var mppOpen = 0, mppOnProgress = 0, mppClosed = 0;
    if (mppSheet && mppSheet.getLastRow() > 1) {
      var mppData = mppSheet.getRange(2, 1, mppSheet.getLastRow()-1, 11).getValues();
      mppData.forEach(function(r) {
        var s = (r[10] || '').toString();
        if (s === 'Open')        mppOpen++;
        else if (s === 'On Progress') mppOnProgress++;
        else if (s === 'Closed') mppClosed++;
      });
    }

    var totalKnd = 0, hired = 0, reject = 0, proses = 0;
    if (kndSheet && kndSheet.getLastRow() > 1) {
      var kndData = kndSheet.getRange(2, 1, kndSheet.getLastRow()-1, 12).getValues();
      kndData.forEach(function(r) {
        if (!r[0]) return;
        totalKnd++;
        var s = (r[11] || '').toString();
        if (s === 'Hired')  hired++;
        else if (s === 'Reject' || s === 'Withdraw') reject++;
        else proses++;
      });
    }

    return {
      mppOpen: mppOpen, mppOnProgress: mppOnProgress, mppClosed: mppClosed,
      totalKnd: totalKnd, hired: hired, reject: reject, proses: proses
    };
  } catch(err) {
    return { mppOpen:0, mppOnProgress:0, mppClosed:0, totalKnd:0, hired:0, reject:0, proses:0 };
  }
}

// ================================================================
// getSemuaHasilIST() — Ambil semua hasil tes dari db_hasil_psikotes
// Dipanggil dari halaman Psikotes HRIS → tab Hasil Tes
// ================================================================
function getSemuaHasilIST() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('db_hasil_psikotes');
    if (!sheet || sheet.getLastRow() <= 1) return [];
    var COLS = 28; // sesuai HEADERS_HASIL di Code_Tes.gs
    var data = sheet.getRange(2, 1, sheet.getLastRow()-1, COLS).getValues();
    return data
      .map(function(row) {
        return row.map(function(c) {
          if (c instanceof Date && !isNaN(c.getTime())) {
            return c.getDate().toString().padStart(2,'0')+'/'+(c.getMonth()+1).toString().padStart(2,'0')+'/'+c.getFullYear();
          }
          return c;
        });
      })
      .filter(function(row){ return row[0] !== ''; });
  } catch(e) { return []; }
}

// ================================================================
//  PSIKOTES IST — Fungsi yang dipanggil dari HRIS Web App
//  (Code.gs HRIS harus punya fungsi ini agar bisa dipanggil
//   oleh google.script.run di Index.html)
//
//  Semua fungsi ini membaca/menulis ke sheet di Spreadsheet HRIS
//  yang sama (db_token_psikotes & db_hasil_psikotes)
// ================================================================

// Header sheet db_token_psikotes (mirror dari Code_Tes.gs)
var SHEET_TOKEN_PSI = 'db_token_psikotes';
var SHEET_HASIL_PSI = 'db_hasil_psikotes';

var HEADERS_TOKEN_PSI = [
  'TOKEN','KANDIDAT_ID','MPP_ID','NAMA_PESERTA','EMAIL_PESERTA',
  'POSISI','TGL_DIBUAT','TGL_EXPIRED','STATUS',
  'TGL_MULAI','TGL_SELESAI','IP_ADDRESS','TAB_SWITCH_COUNT','DIBUAT_OLEH','USIA_PESERTA'
];

// ────────────────────────────────────────────────────────────────
// getDaftarToken() — Ambil semua token dari db_token_psikotes
// ────────────────────────────────────────────────────────────────
function getDaftarToken() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_TOKEN_PSI);

    // Buat sheet jika belum ada
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_TOKEN_PSI);
      var hr = sheet.appendRow(HEADERS_TOKEN_PSI)
                    .getRange ? null : null;
      sheet.appendRow(HEADERS_TOKEN_PSI);
      var hRange = sheet.getRange(1, 1, 1, HEADERS_TOKEN_PSI.length);
      hRange.setBackground('#132a47').setFontColor('#fff')
            .setFontWeight('bold').setHorizontalAlignment('center');
      sheet.setFrozenRows(1);
      // Sheet baru → belum ada data
      return [];
    }

    if (sheet.getLastRow() <= 1) return [];

    var data = sheet.getRange(2, 1, sheet.getLastRow()-1, HEADERS_TOKEN_PSI.length).getValues();

    return data
      .map(function(row) {
        return row.map(function(cell) {
          if (cell instanceof Date && !isNaN(cell.getTime())) {
            return Utilities.formatDate(cell, 'Asia/Jakarta', 'dd/MM/yyyy HH:mm');
          }
          return cell;
        });
      })
      .filter(function(row) { return row[0] !== '' && row[0] !== null; })
      // Urutkan: terbaru di atas
      .reverse();

  } catch (err) {
    console.error('getDaftarToken ERROR: ' + err.message);
    return [];
  }
}

// ────────────────────────────────────────────────────────────────
// buatTokenPeserta() — Generate token baru & simpan ke sheet
// dipanggil dari halaman Psikotes IST → tab Generate Token
// ────────────────────────────────────────────────────────────────
function buatTokenPeserta(dataForm) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_TOKEN_PSI);

    // Buat sheet jika belum ada
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_TOKEN_PSI);
      sheet.appendRow(HEADERS_TOKEN_PSI);
      var hRange = sheet.getRange(1, 1, 1, HEADERS_TOKEN_PSI.length);
      hRange.setBackground('#132a47').setFontColor('#fff')
            .setFontWeight('bold').setHorizontalAlignment('center');
      sheet.setFrozenRows(1);
      sheet.autoResizeColumns(1, HEADERS_TOKEN_PSI.length);
    }

    // Generate token unik format XXXX-XXXX-XXXX-XXXX
    var token   = generateTokenIST();
    var now     = new Date();
    var expired = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000); // +7 hari

    // Hitung usia dari tglLahir kandidat
    var usiaPeserta = 25; // default jika tidak ada data
    if (dataForm.tglLahir) {
      var tglLahirObj = parseFlexDate(dataForm.tglLahir);
      if (tglLahirObj) {
        var th  = now.getFullYear() - tglLahirObj.getFullYear();
        var bln = now.getMonth()    - tglLahirObj.getMonth();
        if (bln < 0 || (bln === 0 && now.getDate() < tglLahirObj.getDate())) th--;
        usiaPeserta = th;
      }
    } else if (dataForm.usia) {
      usiaPeserta = parseInt(dataForm.usia) || 25;
    }

    var row = [
      token,
      dataForm.kandidatId  || '',
      dataForm.mppId       || '',
      (dataForm.nama       || '').toUpperCase(),
      dataForm.email       || '',
      dataForm.posisi      || '',
      Utilities.formatDate(now,     'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss'),
      Utilities.formatDate(expired, 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss'),
      'Belum Mulai',   // STATUS
      '',              // TGL_MULAI
      '',              // TGL_SELESAI
      '',              // IP_ADDRESS
      0,               // TAB_SWITCH_COUNT
      'Admin HR',       // DIBUAT_OLEH
      usiaPeserta      // USIA_PESERTA
    ];

    sheet.appendRow(row);

    // Format baris baru
    var newRow = sheet.getLastRow();
    sheet.getRange(newRow, 1, 1, HEADERS_TOKEN_PSI.length)
         .setBorder(true,true,true,true,true,true,'#e2e8f0',SpreadsheetApp.BorderStyle.SOLID);
    // Warnai kolom TOKEN
    sheet.getRange(newRow, 1)
         .setFontWeight('bold')
         .setFontColor('#1a3a5c')
         .setBackground('#eff6ff');
    // Warnai kolom STATUS
    sheet.getRange(newRow, 9)
         .setBackground('#ecfdf5')
         .setFontColor('#059669')
         .setFontWeight('bold');

    // Format tanggal expired untuk tampilan
    var expiredDisplay = expired.getDate().toString().padStart(2,'0') + '/'
      + (expired.getMonth()+1).toString().padStart(2,'0') + '/'
      + expired.getFullYear();

    return {
      status:  'success',
      token:   token,
      expired: expiredDisplay,
      nama:    (dataForm.nama || '').toUpperCase()
    };

  } catch (err) {
    return { status: 'error', pesan: err.message };
  }
}

// ────────────────────────────────────────────────────────────────
// generateTokenIST() — Generate token unik XXXX-XXXX-XXXX-XXXX
// ────────────────────────────────────────────────────────────────
function generateTokenIST() {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  // Pastikan token unik (cek duplikat dengan sheet)
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var sheet  = ss.getSheetByName(SHEET_TOKEN_PSI);
  var existing = [];
  if (sheet && sheet.getLastRow() > 1) {
    existing = sheet.getRange(2, 1, sheet.getLastRow()-1, 1)
                    .getValues()
                    .map(function(r){ return r[0].toString(); });
  }

  var token, attempt = 0;
  do {
    token = '';
    for (var i = 0; i < 16; i++) {
      if (i > 0 && i % 4 === 0) token += '-';
      token += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    attempt++;
  } while (existing.indexOf(token) >= 0 && attempt < 100);

  return token;
}
