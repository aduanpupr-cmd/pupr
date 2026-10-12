/**
 * =========================================================================
 * E-TAMU PUPR KOTA GORONTALO
 * Buku Tamu Elektronik — Dinas Pekerjaan Umum dan Penataan Ruang Kota Gorontalo
 * File: Code.gs (Google Apps Script Backend & Web App Engine)
 * =========================================================================
 * 
 * Pengembang: Tim IT Dinas PUPR Kota Gorontalo
 * Zona Waktu: Asia/Makassar (WITA, UTC+8)
 * Database: Google Sheets ("DATABASE E-TAMU PUPR KOTA GORONTALO")
 */

// Konfigurasi Database
var SPREADSHEET_NAME = "DATABASE E-TAMU PUPR KOTA GORONTALO";
var SHEET_NAME = "Data Tamu";
var TIMEZONE = "Asia/Makassar";

// 17 Header Kolom Resmi E-Tamu PUPR
var HEADERS = [
  "ID Kunjungan",          // Kolom 1 (No. Tiket Tamu / Aduan)
  "Timestamp",             // Kolom 2 (WITA)
  "Nama Lengkap",          // Kolom 3 (Sesuai KTP)
  "Jenis Kelamin",         // Kolom 4 (Laki-laki / Perempuan)
  "Usia",                  // Kolom 5 (Tahun)
  "Alamat Lengkap",        // Kolom 6 (Sesuai KTP)
  "Nomor Handphone",       // Kolom 7 (Teks, awalan 0)
  "Kategori Instansi",     // Kolom 8 (13 Kategori Pilihan)
  "Nama Instansi",         // Kolom 9 (Nama Lembaga / Perusahaan)
  "Tujuan Kunjungan",      // Kolom 10
  "Ada Pengaduan",         // Kolom 11 (Ada / Tidak Ada)
  "Uraian Pengaduan",      // Kolom 12
  "Status Kunjungan",      // Kolom 13 (Menunggu / Sedang Diproses / Selesai / Perlu Tindak Lanjut)
  "Progres Aduan",         // Kolom 14 (Diterima / Verifikasi & Disposisi / Penanganan Lapangan / Selesai Ditindaklanjuti)
  "Bidang Terkait",        // Kolom 15 (Bidang Penangan di PUPR)
  "Tindak Lanjut",         // Kolom 16 (Uraian Tindak Lanjut / Solusi)
  "Estimasi Selesai"       // Kolom 17 (Target / Tanggal Selesai)
];

/**
 * Routing Web App (doGet)
 */
function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) ? e.parameter.page.toLowerCase() : "home";
  var template;

  if (page === "admin") {
    template = HtmlService.createTemplateFromFile("Admin");
  } else if (page === "login") {
    template = HtmlService.createTemplateFromFile("Login");
  } else {
    template = HtmlService.createTemplateFromFile("Index");
  }

  return template.evaluate()
    .setTitle("E-Tamu PUPR Kota Gorontalo — Buku Tamu Elektronik")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Webhook API Endpoint (doPost) untuk integrasi Frontend / Aplikasi Lain
 */
function doPost(e) {
  try {
    var rawData = e.postData.contents;
    var data = JSON.parse(rawData);
    var action = data.action;

    if (action === "addGuest") {
      var result = submitGuestRegistration(data);
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } else if (action === "login") {
      var loginResult = loginAdmin(data.username, data.password);
      return ContentService.createTextOutput(JSON.stringify(loginResult))
        .setMimeType(ContentService.MimeType.JSON);
    } else if (action === "checkProgress") {
      var progressResult = checkComplaintProgress(data.ticketOrPhone);
      return ContentService.createTextOutput(JSON.stringify(progressResult))
        .setMimeType(ContentService.MimeType.JSON);
    } else if (action === "updateProgress") {
      var updateProgressResult = updateComplaintProgressAdmin(data.token, data.id, data);
      return ContentService.createTextOutput(JSON.stringify(updateProgressResult))
        .setMimeType(ContentService.MimeType.JSON);
    } else if (action === "updateStatus") {
      var updateResult = updateVisitStatus(data.token, data.id, data.statusKunjungan, data.catatanPetugas);
      return ContentService.createTextOutput(JSON.stringify(updateResult))
        .setMimeType(ContentService.MimeType.JSON);
    } else if (action === "ping") {
      return ContentService.createTextOutput(JSON.stringify({ success: true, message: "PUPR Webhook Aktif", time: formatWitaNow() }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ success: false, message: "Aksi tidak dikenal" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Helper include file HTML modular
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Sanitasi Spreadsheet Formula Injection (=, +, -, @, \t, \r)
 */
function sanitizeForSpreadsheet(val) {
  if (val === null || val === undefined) return "";
  var str = String(val).trim();
  if (/^[=+\-@\t\r]/.test(str)) {
    return "'" + str;
  }
  return str;
}

/**
 * Format Waktu Asia/Makassar (WITA)
 */
function formatWitaNow() {
  var d = new Date();
  return Utilities.formatDate(d, TIMEZONE, "dd/MM/yyyy, HH:mm:ss") + " WITA";
}

/**
 * Pembuatan Nomor Tiket / ID Kunjungan Unik (PUPR-YYYYMMDD-XXXX)
 */
function generateVisitId() {
  var d = new Date();
  var dateStr = Utilities.formatDate(d, TIMEZONE, "yyyyMMdd");
  var randomNum = Math.floor(1000 + Math.random() * 9000);
  return "PUPR-" + dateStr + "-" + randomNum;
}

/**
 * Inisialisasi Database Google Sheets & Format Kolom
 */
function setupInitialDatabase() {
  var ss;
  var files = DriveApp.getFilesByName(SPREADSHEET_NAME);
  
  if (files.hasNext()) {
    ss = SpreadsheetApp.open(files.next());
  } else {
    ss = SpreadsheetApp.create(SPREADSHEET_NAME);
  }

  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    var defaultSheet = ss.getSheetByName("Sheet1");
    if (defaultSheet && ss.getSheets().length > 1) {
      try { ss.deleteSheet(defaultSheet); } catch (e) {}
    }
  }

  // Tulis Header jika baris 1 kosong
  var firstCell = sheet.getRange(1, 1).getValue();
  if (!firstCell) {
    var headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setValues([HEADERS]);
    headerRange.setBackground("#1e40af");
    headerRange.setFontColor("#ffffff");
    headerRange.setFontWeight("bold");
    headerRange.setFontFamily("Roboto");
    headerRange.setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
    
    // Format nomor HP (kolom 7) sebagai teks
    sheet.getRange(2, 7, 2000, 1).setNumberFormat("@");
    
    for (var i = 1; i <= HEADERS.length; i++) {
      sheet.autoResizeColumn(i);
    }
  }

  // Set default Admin credentials jika belum ada
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty("ADMIN_USERNAME")) {
    props.setProperty("ADMIN_USERNAME", "admin");
    props.setProperty("ADMIN_PASSWORD_HASH", hashPassword("puprgorontalo2026!"));
  }

  return {
    success: true,
    spreadsheetId: ss.getId(),
    spreadsheetUrl: ss.getUrl()
  };
}

/**
 * Dapatkan Spreadsheet aktif
 */
function getSpreadsheet() {
  var files = DriveApp.getFilesByName(SPREADSHEET_NAME);
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  }
  try {
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    setupInitialDatabase();
    var newFiles = DriveApp.getFilesByName(SPREADSHEET_NAME);
    return SpreadsheetApp.open(newFiles.next());
  }
}

/**
 * Pendaftaran Tamu — Menyimpan ke Google Sheets dengan LockService
 */
function submitGuestRegistration(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    // 1. Validasi
    var nama = String(data.namaLengkap || "").trim();
    if (!nama || nama.length < 3) {
      return { success: false, message: "Nama lengkap wajib diisi minimal 3 karakter." };
    }

    var jenisKelamin = data.jenisKelamin;
    if (jenisKelamin !== "Laki-laki" && jenisKelamin !== "Perempuan") {
      return { success: false, message: "Silakan pilih jenis kelamin." };
    }

    var usia = parseInt(data.usia, 10);
    if (isNaN(usia) || usia < 12 || usia > 120) {
      return { success: false, message: "Usia harus berupa angka antara 12 dan 120 tahun." };
    }

    var alamat = String(data.alamat || "").trim();
    if (!alamat || alamat.length < 5) {
      return { success: false, message: "Alamat lengkap wajib diisi." };
    }

    var noHp = String(data.nomorHp || "").replace(/\s+/g, "").trim();
    if (!/^(08|\+628|628)[0-9]{8,12}$/.test(noHp)) {
      return { success: false, message: "Nomor handphone tidak valid (gunakan format Indonesia)." };
    }
    if (noHp.indexOf("+62") === 0) noHp = "0" + noHp.substring(3);
    else if (noHp.indexOf("62") === 0) noHp = "0" + noHp.substring(2);

    var kategoriInstansi = String(data.kategoriInstansi || "Masyarakat Umum / Perorangan").trim();
    var namaInstansi = String(data.namaInstansi || "").trim();

    var tujuan = String(data.tujuanKunjungan || "").trim();
    if (!tujuan) {
      return { success: false, message: "Tujuan kunjungan wajib dipilih." };
    }
    if (tujuan === "Lainnya" && data.tujuanDetail) {
      tujuan = "Lainnya: " + String(data.tujuanDetail).trim();
    }

    var adaPengaduan = (data.hasPengaduan === true || data.adaPengaduan === "Ada") ? "Ada" : "Tidak Ada";
    var uraianPengaduan = "";
    if (adaPengaduan === "Ada") {
      uraianPengaduan = String(data.uraianPengaduan || "").trim();
      if (!uraianPengaduan || uraianPengaduan.length < 5) {
        return { success: false, message: "Uraian keluhan/pengaduan wajib diisi ketika opsi dicentang." };
      }
    }

    // 2. ID Kunjungan & Waktu Server
    var visitId = data.id || generateVisitId();
    var timestampStr = formatWitaNow();

    // 3. Baris Data 17 Kolom
    var rowData = [
      visitId,
      timestampStr,
      sanitizeForSpreadsheet(nama),
      jenisKelamin,
      usia,
      sanitizeForSpreadsheet(alamat),
      "'" + noHp,
      sanitizeForSpreadsheet(kategoriInstansi),
      sanitizeForSpreadsheet(namaInstansi),
      sanitizeForSpreadsheet(tujuan),
      adaPengaduan,
      sanitizeForSpreadsheet(uraianPengaduan),
      "Menunggu",
      adaPengaduan === "Ada" ? "Diterima" : "-",
      adaPengaduan === "Ada" ? "Sekretariat / Tim Verifikasi Dinas PUPR" : "-",
      adaPengaduan === "Ada" ? "Aduan telah resmi teregistrasi di sistem E-Tamu PUPR" : "-",
      "-"
    ];

    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      setupInitialDatabase();
      sheet = ss.getSheetByName(SHEET_NAME);
    }

    sheet.appendRow(rowData);

    return {
      success: true,
      message: "Pendaftaran berhasil disimpan ke database.",
      data: {
        id: visitId,
        namaLengkap: nama,
        tujuanKunjungan: tujuan,
        kategoriInstansi: kategoriInstansi,
        adaPengaduan: adaPengaduan,
        timestamp: timestampStr
      }
    };
  } catch (err) {
    return {
      success: false,
      message: "Pendaftaran belum berhasil disimpan. " + err.toString()
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Hashing Password (SHA-256)
 */
function hashPassword(pass) {
  var rawBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pass, Utilities.Charset.UTF_8);
  var hashStr = "";
  for (var i = 0; i < rawBytes.length; i++) {
    var byteVal = rawBytes[i];
    if (byteVal < 0) byteVal += 256;
    var byteHex = byteVal.toString(16);
    if (byteHex.length === 1) byteHex = "0" + byteHex;
    hashStr += byteHex;
  }
  return hashStr;
}

/**
 * Autentikasi Admin Server-Side
 */
function loginAdmin(username, password) {
  var props = PropertiesService.getScriptProperties();
  var storedUser = props.getProperty("ADMIN_USERNAME") || "admin";
  var storedHash = props.getProperty("ADMIN_PASSWORD_HASH") || hashPassword("puprgorontalo2026!");

  if (username === storedUser && hashPassword(password) === storedHash) {
    var token = Utilities.getUuid();
    var cache = CacheService.getScriptCache();
    cache.put("SESSION_" + token, username, 7200);

    return {
      success: true,
      message: "Login berhasil.",
      token: token,
      username: username
    };
  } else {
    return {
      success: false,
      message: "Username atau password salah."
    };
  }
}

/**
 * Verifikasi Sesi Admin Server-Side
 */
function verifySession(token) {
  if (!token) return false;
  var cache = CacheService.getScriptCache();
  var user = cache.get("SESSION_" + token);
  return !!user;
}

/**
 * Logout Admin
 */
function logoutAdmin(token) {
  if (token) {
    var cache = CacheService.getScriptCache();
    cache.remove("SESSION_" + token);
  }
  return { success: true };
}

/**
 * Mengambil Seluruh Data Kunjungan (Hanya untuk Admin Terotentikasi)
 */
function getAdminVisits(token) {
  if (!verifySession(token)) {
    throw new Error("Akses ditolak. Sesi tidak sah atau telah kedaluwarsa.");
  }

  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return [];
  }

  var dataRange = sheet.getRange(2, 1, lastRow - 1, HEADERS.length);
  var values = dataRange.getValues();
  var results = [];

  for (var i = values.length - 1; i >= 0; i--) {
    var row = values[i];
    if (!row[0]) continue;
    results.push({
      rowIndex: i + 2,
      id: String(row[0]),
      timestamp: String(row[1]),
      namaLengkap: String(row[2]),
      jenisKelamin: String(row[3]),
      usia: row[4],
      alamat: String(row[5]),
      nomorHp: String(row[6]).replace(/^'/, ""),
      kategoriInstansi: String(row[7] || ""),
      namaInstansi: String(row[8] || ""),
      tujuanKunjungan: String(row[9]),
      adaPengaduan: String(row[10]),
      uraianPengaduan: String(row[11]),
      statusKunjungan: String(row[12]) || "Menunggu",
      progresAduan: String(row[13] || "-"),
      bidangTerkait: String(row[14] || "-"),
      tindakLanjut: String(row[15] || "-"),
      estimasiSelesai: String(row[16] || "-")
    });
  }

  return results;
}

/**
 * Perbarui Status Kunjungan
 */
function updateVisitStatus(token, visitId, newStatus, catatan) {
  if (!verifySession(token)) {
    return { success: false, message: "Akses ditolak. Silakan login kembali." };
  }

  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();

  for (var r = 1; r < data.length; r++) {
    if (String(data[r][0]) === String(visitId)) {
      sheet.getRange(r + 1, 13).setValue(newStatus);
      return { success: true, message: "Status kunjungan berhasil diperbarui." };
    }
  }

  return { success: false, message: "Data kunjungan tidak ditemukan." };
}

/**
 * Hapus Data Kunjungan
 */
function deleteVisit(token, visitId) {
  if (!verifySession(token)) {
    return { success: false, message: "Akses ditolak." };
  }

  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();

  for (var r = 1; r < data.length; r++) {
    if (String(data[r][0]) === String(visitId)) {
      sheet.deleteRow(r + 1);
      return { success: true, message: "Data berhasil dihapus dari Google Sheets." };
    }
  }

  return { success: false, message: "Data tidak ditemukan." };
}

/**
 * Cek Progres Penyelesaian Aduan (Akses Publik via Nomor Tiket / HP)
 */
function checkComplaintProgress(query) {
  if (!query) return { success: false, message: "Nomor tiket atau nomor handphone wajib diisi." };
  
  var cleanQuery = String(query).trim().toLowerCase();
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) {
    return { success: false, message: "Belum ada data aduan terdaftar." };
  }
  
  var values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  var results = [];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var id = String(row[0] || "").toLowerCase();
    var phone = String(row[6] || "").replace(/[^0-9]/g, "");
    var cleanPhoneQuery = cleanQuery.replace(/[^0-9]/g, "");

    var isMatch = (id === cleanQuery || id.indexOf(cleanQuery) !== -1);
    if (!isMatch && cleanPhoneQuery && cleanPhoneQuery.length >= 6) {
      isMatch = (phone.indexOf(cleanPhoneQuery) !== -1);
    }

    if (isMatch) {
      results.push({
        id: String(row[0]),
        timestamp: String(row[1]),
        namaLengkap: String(row[2]),
        nomorHp: String(row[6]).replace(/^'/, ""),
        kategoriInstansi: String(row[7] || ""),
        namaInstansi: String(row[8] || ""),
        tujuanKunjungan: String(row[9]),
        adaPengaduan: String(row[10]),
        uraianPengaduan: String(row[11]),
        statusKunjungan: String(row[12]) || "Menunggu",
        progresAduan: String(row[13] || "Diterima"),
        bidangTerkait: String(row[14] || "Sekretariat / Tim Verifikasi Dinas PUPR"),
        tindakLanjut: String(row[15] || "Aduan telah resmi teregistrasi di sistem"),
        estimasiSelesai: String(row[16] || "-")
      });
    }
  }

  if (results.length === 0) {
    return { success: false, message: "Data aduan tidak ditemukan. Periksa kembali nomor tiket Anda." };
  }

  return { success: true, data: results };
}

/**
 * Update Progres Aduan oleh Admin ke Google Sheets
 */
function updateComplaintProgressAdmin(token, visitId, params) {
  if (!verifySession(token)) {
    return { success: false, message: "Akses ditolak. Sesi tidak sah." };
  }

  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();

  for (var r = 1; r < data.length; r++) {
    if (String(data[r][0]) === String(visitId)) {
      if (params.statusKunjungan) {
        sheet.getRange(r + 1, 13).setValue(params.statusKunjungan);
      }
      if (params.progresAduan) {
        sheet.getRange(r + 1, 14).setValue(params.progresAduan);
      }
      if (params.bidangTerkait) {
        sheet.getRange(r + 1, 15).setValue(params.bidangTerkait);
      }
      if (params.tindakLanjut) {
        sheet.getRange(r + 1, 16).setValue(params.tindakLanjut);
      }
      if (params.estimasiSelesai) {
        sheet.getRange(r + 1, 17).setValue(params.estimasiSelesai);
      }
      return {
        success: true,
        message: "Progres aduan berhasil diperbarui ke Google Sheets."
      };
    }
  }

  return { success: false, message: "Data aduan tidak ditemukan." };
}
