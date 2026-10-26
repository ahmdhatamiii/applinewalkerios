/**
 * PLN UPT Semarang - Apps Script Initialization & Safe Migrate Protocol
 * File: setup.gs
 * System: Management & Monitoring Line Walker UPT Semarang
 */

var SPREADSHEET_NAME_LW = "Data LW";
var SPREADSHEET_NAME_ACUAN = "AcuanKendala";
var SPREADSHEET_NAME_BA = "BeritaAcara";
var SPREADSHEET_NAME_WO = "WorkOrder";
var SPREADSHEET_NAME_CEKLIS = "Ceklis";
var SPREADSHEET_NAME_ABSENSI = "Absensi";
var SPREADSHEET_NAME_LEMBURAN = "Lemburan";
var SPREADSHEET_NAME_REQUEST_JADWAL = "RequestJadwal";
var SPREADSHEET_NAME_SURAT_PERINGATAN = "SuratPeringatan";
var SPREADSHEET_NAME_DOCUMENTS = "DataDokumen";
var SPREADSHEET_NAME_SPOTS = "TitikTandaTangan";
var SPREADSHEET_NAME_APPROVER = "Master_Approver";
var SPREADSHEET_NAME_TOWER_SPAN = "TowerSpan";
var DRIVE_FOLDER_NAME = "UPT_Semarang_WorkOrders_PDF";
var DRIVE_FOLDER_CEKLIS_BA = "UPT_Semarang_Ceklis_BeritaAcara";
var DRIVE_FOLDER_SP_PDF = "UPT_Semarang_SuratPeringatan_PDF";
var DRIVE_FOLDER_APPROVAL_STORAGE = "UPT_Semarang_Approval_Storage";
var DRIVE_FOLDER_FOTO_LW = "UPT_Semarang_Foto_LineWalker";

/**
 * Safe Migrate Initialization
 * Creates sheets and adds header & initial seed data if sheets are empty.
 * Will NEVER overwrite or clear existing records.
 */
function setupDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Setup Sheet Data LW with Full Master Dataset (9 Columns termasuk FOTO_URL)
  var sheetLW = getOrCreateSheet(ss, SPREADSHEET_NAME_LW);
  var headerLW = ["LINE WALKER", "KLW", "NIK BARU", "POSISI", "ULTG", "GI", "TL GI", "Pengahantar", "FOTO_URL"];
  ensureHeaderAndData(sheetLW, headerLW, getMasterLineWalkerDataset());

  // 2. Setup Sheet AcuanKendala
  var sheetAcuan = getOrCreateSheet(ss, SPREADSHEET_NAME_ACUAN);
  var headerAcuan = [
    "JENIS_BA", 
    "NAMA", 
    "JABATAN", 
    "NIP", 
    "PENGHANTAR", 
    "GARDU_INDUK", 
    "ULTG", 
    "TANGGAL", 
    "BULAN", 
    "KETERANGAN_KENDALA", 
    "KLW", 
    "TL_GI", 
    "TANGGAL_SURAT", 
    "KETERANGAN"
  ];
  ensureHeaderAndData(sheetAcuan, headerAcuan, getDummyAcuanKendalaData());

  // 3. Setup Sheet BeritaAcara
  var sheetBA = getOrCreateSheet(ss, SPREADSHEET_NAME_BA);
  var headerBA = [
    "ID_BA", 
    "JENIS_BA", 
    "NAMA", 
    "JABATAN", 
    "NIP", 
    "PENGHANTAR", 
    "GARDU_INDUK", 
    "ULTG", 
    "TANGGAL", 
    "BULAN", 
    "KETERANGAN_KENDALA", 
    "KLW", 
    "TL_GI", 
    "TANGGAL_SURAT", 
    "STATUS", 
    "TANGGAL_INPUT",
    "FILE_SIGNED_PDF_URL",
    "FILE_SIGNED_PDF_ID"
  ];
  ensureHeaderAndData(sheetBA, headerBA, getDummyBeritaAcaraData());

  // 4. Setup Sheet WorkOrder
  var sheetWO = getOrCreateSheet(ss, SPREADSHEET_NAME_WO);
  var headerWO = ["ID_WO", "Nama_GI", "Bulan_Order", "Deskripsi", "File_PDF_Url", "File_PDF_Id", "Tanggal_Input", "JSON_TANGGAL_PERSONEL", "STATUS_TTD", "PENANDATANGAN", "FILE_TTD_URL"];
  ensureHeaderAndData(sheetWO, headerWO, getDummyWorkOrderData());

  // 5. Setup Sheet Ceklis
  var sheetCeklis = getOrCreateSheet(ss, SPREADSHEET_NAME_CEKLIS);
  var headerCeklis = ["ULTG", "Nama_Line_Walker", "NIP", "Bulan", "Tahun", "Jenis_Kendala", "Status_Ceklis", "Tanggal_Update", "File_BA_Url", "File_BA_Id", "Tanggal_Kendala"];
  ensureHeaderAndData(sheetCeklis, headerCeklis, getDummyCeklisData());

  // 6. Setup Sheet Absensi
  var sheetAbsensi = getOrCreateSheet(ss, SPREADSHEET_NAME_ABSENSI);
  var headerAbsensi = ["ID_ABSEN", "NAMA_LW", "NIP", "TANGGAL", "BULAN", "TAHUN", "JAM_MASUK", "JAM_KELUAR", "STATUS_KETERLAMBATAN", "STATUS_PULANG"];
  ensureHeaderAndData(sheetAbsensi, headerAbsensi, getDummyAbsensiData());

  // 7. Setup Sheet Lemburan
  var sheetLemburan = getOrCreateSheet(ss, SPREADSHEET_NAME_LEMBURAN);
  var headerLemburan = ["ID_LEMBUR", "BULAN_ORDER", "TAHUN", "NAMA_FILE", "FILE_EXCEL_URL", "FILE_EXCEL_ID", "TANGGAL_UPLOAD"];
  ensureHeaderAndData(sheetLemburan, headerLemburan, getDummyLemburanData());

  // 8. Setup Sheet RequestJadwal
  var sheetRequestJadwal = getOrCreateSheet(ss, SPREADSHEET_NAME_REQUEST_JADWAL);
  var headerRequestJadwal = ["ID_REQUEST", "NAMA_LW", "NIP", "BULAN_TARGET", "DETAIL_PERUBAHAN", "STATUS", "TANGGAL_SUBMIT"];
  ensureHeaderAndData(sheetRequestJadwal, headerRequestJadwal, getDummyRequestJadwalData());

  // 9. Setup Sheet SuratPeringatan
  var sheetSP = getOrCreateSheet(ss, SPREADSHEET_NAME_SURAT_PERINGATAN);
  var headerSP = ["ID_SP", "NAMA_LW", "NIP", "JENIS_SURAT", "NOMOR_SURAT", "TANGGAL_SURAT", "MASA_BERLAKU_BULAN", "TANGGAL_BERAKHIR", "KETERANGAN", "FILE_PDF_URL", "FILE_PDF_ID", "STATUS_AKTIF", "TANGGAL_UPLOAD"];
  ensureHeaderAndData(sheetSP, headerSP, getDummySuratPeringatanData());

  // 10. Setup Sheet DataDokumen (Approval Studio)
  var sheetDoc = getOrCreateSheet(ss, SPREADSHEET_NAME_DOCUMENTS);
  var headerDoc = ["ID", "TITLE", "CREATED_AT", "ORIGINAL_FILE_NAME", "DRIVE_FILE_ID_ORIGINAL", "DRIVE_FILE_ID_SIGNED", "TOTAL_PAGES", "TOTAL_SPOTS", "SIGNED_COUNT", "STATUS"];
  ensureHeaderAndData(sheetDoc, headerDoc, []);

  // 11. Setup Sheet TitikTandaTangan (Approval Studio)
  var sheetSpot = getOrCreateSheet(ss, SPREADSHEET_NAME_SPOTS);
  var headerSpot = ["SPOT_ID", "DOC_ID", "SIGNER_NAME", "SIGNER_ROLE", "PAGE_NUMBER", "X_PERCENT", "Y_PERCENT", "WIDTH_PERCENT", "HEIGHT_PERCENT", "IS_SIGNED", "SIGNED_AT", "VERIFICATION_HASH"];
  ensureHeaderAndData(sheetSpot, headerSpot, []);

  // 12. Setup Sheet Master_Approver (PLN ES UP2 & UPT Semarang)
  var sheetApprover = getOrCreateSheet(ss, SPREADSHEET_NAME_APPROVER);
  var headerApprover = ["ID", "UNIT", "KATEGORI", "ULTG", "GI", "NAMA", "JABATAN", "NIP_NIK", "KETERANGAN"];
  ensureHeaderAndData(sheetApprover, headerApprover, getMasterApproverDataset());

  // 13. Setup Sheet TowerSpan untuk data lokasi tower
  var sheetTower = getOrCreateSheet(ss, SPREADSHEET_NAME_TOWER_SPAN);
  var headerTower = ["TOWER_ID", "NAMA_TOWER", "LAT", "LNG", "PENGHANTAR", "ULTG", "GI", "LINE_WALKER", "KLW", "DESKRIPSI"];
  ensureHeaderAndData(sheetTower, headerTower, []);

  // 14. Setup Google Drive Folders
  getOrCreateDriveFolder(DRIVE_FOLDER_NAME);
  getOrCreateDriveFolder(DRIVE_FOLDER_CEKLIS_BA);
  getOrCreateDriveFolder(DRIVE_FOLDER_SP_PDF);
  getOrCreateDriveFolder(DRIVE_FOLDER_APPROVAL_STORAGE);

  SpreadsheetApp.flush();
  Logger.log("Safe Migrate Finished Successfully!");
  return { status: "success", message: "Inisialisasi Database Line Walker UPT Semarang Berhasil!" };
}

/**
 * Inisialisasi Khusus Modul Approval Dokumen PDF
 */
function setupApprovalDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDoc = getOrCreateSheet(ss, SPREADSHEET_NAME_DOCUMENTS);
  var headerDoc = ["ID", "TITLE", "CREATED_AT", "ORIGINAL_FILE_NAME", "DRIVE_FILE_ID_ORIGINAL", "DRIVE_FILE_ID_SIGNED", "TOTAL_PAGES", "TOTAL_SPOTS", "SIGNED_COUNT", "STATUS"];
  ensureHeaderAndData(sheetDoc, headerDoc, []);

  var sheetSpot = getOrCreateSheet(ss, SPREADSHEET_NAME_SPOTS);
  var headerSpot = ["SPOT_ID", "DOC_ID", "SIGNER_NAME", "SIGNER_ROLE", "PAGE_NUMBER", "X_PERCENT", "Y_PERCENT", "WIDTH_PERCENT", "HEIGHT_PERCENT", "IS_SIGNED", "SIGNED_AT", "VERIFICATION_HASH"];
  ensureHeaderAndData(sheetSpot, headerSpot, []);

  var sheetApprover = getOrCreateSheet(ss, SPREADSHEET_NAME_APPROVER);
  var headerApprover = ["ID", "UNIT", "KATEGORI", "ULTG", "GI", "NAMA", "JABATAN", "NIP_NIK", "KETERANGAN"];
  ensureHeaderAndData(sheetApprover, headerApprover, getMasterApproverDataset());

  getOrCreateDriveFolder(DRIVE_FOLDER_APPROVAL_STORAGE);
  SpreadsheetApp.flush();
  return { status: "success", message: "Database Approval Dokumen Siap Digunakan!" };
}

function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    var all = ss.getSheets();
    var cleanTarget = name.replace(/[\s_\-]+/g, "").toLowerCase();
    for (var i = 0; i < all.length; i++) {
      if (all[i].getName().replace(/[\s_\-]+/g, "").toLowerCase() === cleanTarget) {
        return all[i];
      }
    }
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function ensureHeaderAndData(sheet, headers, seedData) {
  var lastRow = sheet.getLastRow();
  
  if (lastRow === 0) {
    sheet.appendRow(headers);
    if (seedData && seedData.length > 0) {
      sheet.getRange(2, 1, seedData.length, headers.length).setValues(seedData);
    }
  } else {
    var currentHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getDisplayValues()[0];
    if (currentHeaders.length < headers.length || currentHeaders[0] === "") {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }
}

function getOrCreateDriveFolder(folderName) {
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder(folderName);
  }
}

function getMasterLineWalkerDataset() {
  return [
    ["EKA FEBRIANTO", "EKA FEBRIANTO", "9125904TJBT", "Koordinator Line Walker", "REMBANG", "GI REMBANG", "ARTIKA KURNIANINGSIH", "-"],
    ["ADITIYA ANDI NUGROHO", "EKA FEBRIANTO", "9825999TJBT", "Line Walker", "REMBANG", "GI SLUKE", "ARTIKA KURNIANINGSIH", "TRS 150kV INC REMBANG-PATI"],
    ["ABDUL ROCHMAN", "EKA FEBRIANTO", "7525641TJBT", "Line Walker", "REMBANG", "GI REMBANG", "ARTIKA KURNIANINGSIH", "TRS 150kV REMBANG-BLORA"],
    ["IMAM SANTOSO", "EKA FEBRIANTO", "7825709TJBT", "Line Walker", "REMBANG", "GI PATI", "IHYA AYYUBI MIRZA", "TRS 150kV INC REMBANG-PATI"],
    ["HERLAMBANG PUTRA MAULANA", "EKA FEBRIANTO", "0325009TJBT", "Line Walker", "REMBANG", "GI PATI", "IHYA AYYUBI MIRZA", "TRS 150kV INC JEKULO-PATI"],
    ["MUSLIMIN", "EKA FEBRIANTO", "8825864TJBT", "Line Walker", "REMBANG", "GI SLUKE", "IHYA AYYUBI MIRZA", "TRS 150kV PLTU SLUKE-INC SLUKE (PATI)"],
    ["NANANG ADI PRABOWO", "EKA FEBRIANTO", "9425952TJBT", "Line Walker", "REMBANG", "GI PATI", "IHYA AYYUBI MIRZA", "TRS 150kV INC REMBANG-PATI"],
    ["PRIHADI", "EKA FEBRIANTO", "7825712TJBT", "Line Walker", "REMBANG", "GI REMBANG", "ARTIKA KURNIANINGSIH", "TRS 150kV REMBANG-BLORA"],
    ["SUGIONO", "EKA FEBRIANTO", "9225920TJBT", "Line Walker", "REMBANG", "GI REMBANG", "ARTIKA KURNIANINGSIH", "TRS 150kV PLTU SLUKE-REMBANG"],
    ["BAMBANG RIYANTO", "BAMBANG RIYANTO", "8425819TJBT", "Koordinator Line Walker", "KUDUS", "GI PURWODADI", "SANDI PRATAMA RAKHMADI", "-"],
    ["AHMAD AMIN DHUKA", "BAMBANG RIYANTO", "9125902TJBT", "Line Walker", "KUDUS", "GI PURWODADI", "SANDI PRATAMA RAKHMADI", "TRS 500kV TANJUNG JATI-UNGARAN"],
    ["ATOK SUDARMANTO", "BAMBANG RIYANTO", "8325806TJBT", "Line Walker", "KUDUS", "GI PURWODADI", "SANDI PRATAMA RAKHMADI", "TRS 500kV UNGARAN-KRIAN"],
    ["BAMBANG HERNOWO", "BAMBANG RIYANTO", "8425818TJBT", "Line Walker", "KUDUS", "GI PURWODADI", "SANDI PRATAMA RAKHMADI", "TRS 150kV KUDUS-PURWODADI"],
    ["EKO SUYANTO", "BAMBANG RIYANTO", "8125770TJBT", "Line Walker", "KUDUS", "GI PURWODADI", "SANDI PRATAMA RAKHMADI", "TRS 500kV TANJUNG JATI-UNGARAN"],
    ["MUKAHAR", "BAMBANG RIYANTO", "7725673TJBT", "Line Walker", "KUDUS", "GI PURWODADI", "SANDI PRATAMA RAKHMADI", "TRS 150kV INC MRANGGEN-PURWODADI"],
    ["AHMAD RIZKI FITRIANTO", "BAMBANG RIYANTO", "9525960TJBT", "Line Walker", "KUDUS", "GI PURWODADI", "SANDI PRATAMA RAKHMADI", "TRS 500kV TANJUNG JATI-PEMALANG"],
    ["SARWADI", "BAMBANG RIYANTO", "8625844TJBT", "Line Walker", "KUDUS", "GI PURWODADI", "SANDI PRATAMA RAKHMADI", "TRS 150kV KUDUS-PURWODADI"],
    ["KRISNA ARI WIJAYA", "BAMBANG RIYANTO", "0325006TJBT", "Line Walker", "KUDUS", "GI PURWODADI", "SANDI PRATAMA RAKHMADI", "TRS 500kV TANJUNG JATI-PEMALANG"],
    ["MUHAMMAD DENDI RENALDI", "MUHAMMAD DENDI RENALDI", "9925015TJBT", "Koordinator Line Walker", "KUDUS", "GI TANJUNG JATI", "DANIAL ZULKARNIANSYAH", "-"],
    ["AGUS HARIYANTO", "MUHAMMAD DENDI RENALDI", "8625841TJBT", "Line Walker", "KUDUS", "GI TANJUNG JATI", "DANIAL ZULKARNIANSYAH", "TRS 150kV INC JEPARA-TANJUNG JATI"],
    ["AHMAD ASRO", "MUHAMMAD DENDI RENALDI", "9325927TJBT", "Line Walker", "KUDUS", "GI TANJUNG JATI", "DANIAL ZULKARNIANSYAH", "TRS 500kV TANJUNG JATI-NEW TANJUNG JATI"],
    ["AHMAD RIDWAN", "MUHAMMAD DENDI RENALDI", "7325618TJBT", "Line Walker", "KUDUS", "GI TANJUNG JATI", "DANIAL ZULKARNIANSYAH", "TRS 500kV TANJUNG JATI-UNGARAN"],
    ["ERWIN ARDI DIRTANA", "MUHAMMAD DENDI RENALDI", "8125771TJBT", "Line Walker", "KUDUS", "GI TANJUNG JATI", "DANIAL ZULKARNIANSYAH", "TRS 150kV INC JEPARA-TANJUNG JATI"],
    ["M. AIWAN HANDOKO", "MUHAMMAD DENDI RENALDI", "9625973TJBT", "Line Walker", "KUDUS", "GI TANJUNG JATI", "DANIAL ZULKARNIANSYAH", "TRS 150kV TANJUNGJATI-SAYUNG"],
    ["DEDI SISWOYO", "MUHAMMAD DENDI RENALDI", "9125903TJBT", "Line Walker", "KUDUS", "GI TANJUNG JATI", "DANIAL ZULKARNIANSYAH", "TRS 500kV TANJUNG JATI-PEMALANG"],
    ["MUHAMMAD UFRON", "MUHAMMAD DENDI RENALDI", "9125907TJBT", "Line Walker", "KUDUS", "GI TANJUNG JATI", "DANIAL ZULKARNIANSYAH", "TRS 150kV TANJUNGJATI-SAYUNG"],
    ["ROHMAD SYAIFUL MUSTAKIM", "MUHAMMAD DENDI RENALDI", "9325930TJBT", "Line Walker", "KUDUS", "GI TANJUNG JATI", "DANIAL ZULKARNIANSYAH", "TRS 150kV TANJUNGJATI-SAYUNG"],
    ["SHOLIHUN", "MUHAMMAD DENDI RENALDI", "8425824TJBT", "Line Walker", "KUDUS", "GI TANJUNG JATI", "DANIAL ZULKARNIANSYAH", "TRS 500kV TANJUNG JATI-UNGARAN"]
  ];
}

function getDummyAcuanKendalaData() {
  return [
    [
      "BERITA ACARA KENDALA ABSENSI & COC", "ABDUL ROCHMAN", "Line Walker", "7525641TJBT", 
      "TRS 150kV REMBANG-BLORA", "GI REMBANG", "REMBANG", "19", "OCTOBER", 
      "Terlambat presisi lokasi titik patroli", "EKA FEBRIANTO", "ARTIKA KURNIANINGSIH", "21 October 2026", "-"
    ],
    [
      "BERITA ACARA KENDALA SOSIALISASI TEGAKAN", "ADITIYA ANDI NUGROHO", "Line Walker", "9825999TJBT", 
      "TRS 150kV INC REMBANG-PATI", "GI SLUKE", "REMBANG", "17", "OCTOBER", 
      "Pemilik lahan belum sepakat kompensasi", "EKA FEBRIANTO", "ARTIKA KURNIANINGSIH", "21 October 2026", "-"
    ]
  ];
}

function getDummyBeritaAcaraData() {
  var nowStr = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");
  return [
    [
      "BA-20261002-0001", "BERITA ACARA KENDALA ABSENSI & COC", "ABDUL ROCHMAN", "Line Walker", "7525641TJBT", 
      "TRS 150kV REMBANG-BLORA", "GI REMBANG", "REMBANG", "19", "OCTOBER", 
      "Terlambat presisi lokasi titik patroli", "EKA FEBRIANTO", "ARTIKA KURNIANINGSIH", "21 October 2026", "Selesai", nowStr, "", ""
    ]
  ];
}

function getDummyWorkOrderData() {
  var nowStr = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");
  return [
    ["WO-20261002-0001", "GI REMBANG", "October 2026", "Penggantian Isolator Flashover T.12 Jalur Rembang-Blora", "#", "dummy-id-1", nowStr, "[]", "Belum TTD", "-", ""]
  ];
}

function getDummyCeklisData() {
  var nowStr = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");
  return [
    ["REMBANG", "ABDUL ROCHMAN", "7525641TJBT", "OCTOBER", "2026", "BERITA ACARA KENDALA SOSIALISASI TEGAKAN", "TRUE", nowStr, "#", "", "19"]
  ];
}

function getDummyAbsensiData() {
  return [
    ["ABS-001", "ABDUL ROCHMAN", "7525641TJBT", "02", "OCTOBER", "2026", "08:15", "-", "Terlambat 15 Menit", "Tidak Absen Checkout"]
  ];
}

function getDummyLemburanData() {
  var nowStr = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");
  return [
    ["LMB-202610-001", "OKTOBER", "2026", "Rekap_Lembur_LW_Oktober_2026.xlsx", "#", "dummy-excel-id-1", nowStr]
  ];
}

function getDummyRequestJadwalData() {
  var nowStr = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");
  return [
    ["REQ-20261002-0001", "ABDUL ROCHMAN", "7525641TJBT", "NOVEMBER 2026", "Tukar jadwal patroli Tgl 10 shift pagi.", "Pending", nowStr]
  ];
}

function getDummySuratPeringatanData() {
  var nowStr = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");
  return [
    ["SP-202610-0001", "ABDUL ROCHMAN", "7525641TJBT", "ST", "012/ST/UPT-SMG/X/2026", "01/10/2026", 3, "01/01/2027", "Tidak absen checkout patroli 3 kali berturut-turut", "#", "dummy-pdf-1", "Aktif", nowStr],
    ["SP-202609-0002", "FIRMAN ADI PRASETYO", "000000", "SP1", "005/SP1/UPT-SMG/IX/2026", "15/09/2026", 6, "15/03/2027", "Keterlambatan presensi dan patroli tanpa keterangan", "#", "dummy-pdf-2", "Aktif", nowStr]
  ];
}

function getMasterApproverDataset() {
  return [
    // 1. UNIT: PLN ES UP2
    ["APP-UP2-01", "PLN ES UP2", "ASMAN", "-", "-", "MAHMUD FAUZI ISWORO", "ASMAN TRANSMISI & GI UP2 PLN ES", "8206012Z", "Asman Transmisi & GI UP2"],
    ["APP-UP2-02", "PLN ES UP2", "TL", "-", "-", "AGUS SETIAWAN", "TL OPERASI & PEMELIHARAAN TRANSMISI UP2", "8809021Z", "Team Leader Op & Har Transmisi UP2"],
    ["APP-UP2-03", "PLN ES UP2", "TL", "-", "-", "TRI WAHYUDI", "TL SCADA & TELEKOMUNIKASI UP2", "8911032Z", "Team Leader Scada & Tel UP2"],
    ["APP-UP2-04", "PLN ES UP2", "TL", "-", "-", "HENDRA WIJAYA", "TL K3L & KEAMANAN UP2", "9012044Z", "Team Leader K3L & Kam UP2"],
    ["APP-UP2-05", "PLN ES UP2", "MANAJER", "-", "-", "PRIHADI UTOMO", "MANAJER PLN ES UP2", "7905001Z", "Manajer UP2"],

    // 2. UNIT: UPT SEMARANG (STRUKTURAL UPT)
    ["APP-UPT-01", "UPT SEMARANG", "MANAJER", "SEMUA", "-", "MANAJER UPT SEMARANG", "MANAJER UPT SEMARANG", "7804001Z", "Manajer UPT"],
    ["APP-UPT-02", "UPT SEMARANG", "ASMAN", "SEMUA", "-", "ASMAN JARINGAN UPT SEMARANG", "ASMAN JARINGAN UPT SEMARANG", "8305012Z", "Asman Jaringan UPT"],
    ["APP-UPT-03", "UPT SEMARANG", "ASMAN", "SEMUA", "-", "ASMAN TRANSAKSI ENERGI & PEMELIHARAAN", "ASMAN TE & HAR UPT SEMARANG", "8407023Z", "Asman TE & Har UPT"],
    ["APP-UPT-04", "UPT SEMARANG", "ASMAN", "SEMUA", "-", "ASMAN PERENCANAAN & EVALUASI", "ASMAN RENVAL UPT SEMARANG", "8509015Z", "Asman Renval UPT"],
    ["APP-UPT-05", "UPT SEMARANG", "ASMAN", "SEMUA", "-", "ASMAN KEUANGAN & UMUM", "ASMAN KEU & UMUM UPT SEMARANG", "8610022Z", "Asman Keu & Umum UPT"],

    // 3. UNIT: UPT SEMARANG - WILAYAH ULTG REMBANG
    ["APP-RMB-01", "UPT SEMARANG", "TL GI", "REMBANG", "GI REMBANG", "ARTIKA KURNIANINGSIH", "TL GI REMBANG", "-", "Team Leader GI Rembang"],
    ["APP-RMB-02", "UPT SEMARANG", "TL GI", "REMBANG", "GI SLUKE", "ARTIKA KURNIANINGSIH", "TL GI SLUKE", "-", "Team Leader GI Sluke"],
    ["APP-RMB-03", "UPT SEMARANG", "TL GI", "REMBANG", "GI PATI", "IHYA AYYUBI MIRZA", "TL GI PATI", "-", "Team Leader GI Pati"],
    ["APP-RMB-04", "UPT SEMARANG", "TL GI", "REMBANG", "GI BLORA", "IHYA AYYUBI MIRZA", "TL GI BLORA", "-", "Team Leader GI Blora"],
    ["APP-RMB-05", "UPT SEMARANG", "KLW", "REMBANG", "GI REMBANG", "EKA FEBRIANTO", "Koordinator Line Walker REMBANG", "9125904TJBT", "Koordinator LW Rembang"],
    ["APP-RMB-06", "UPT SEMARANG", "LW", "REMBANG", "GI REMBANG", "ABDUL ROCHMAN", "Line Walker GI REMBANG", "7525641TJBT", "Line Walker"],
    ["APP-RMB-07", "UPT SEMARANG", "LW", "REMBANG", "GI SLUKE", "ADITIYA ANDI NUGROHO", "Line Walker GI SLUKE", "9825999TJBT", "Line Walker"],
    ["APP-RMB-08", "UPT SEMARANG", "LW", "REMBANG", "GI PATI", "IMAM SANTOSO", "Line Walker GI PATI", "7825709TJBT", "Line Walker"],
    ["APP-RMB-09", "UPT SEMARANG", "LW", "REMBANG", "GI PATI", "HERLAMBANG PUTRA MAULANA", "Line Walker GI PATI", "0325009TJBT", "Line Walker"],
    ["APP-RMB-10", "UPT SEMARANG", "LW", "REMBANG", "GI SLUKE", "MUSLIMIN", "Line Walker GI SLUKE", "8825864TJBT", "Line Walker"],
    ["APP-RMB-11", "UPT SEMARANG", "LW", "REMBANG", "GI PATI", "NANANG ADI PRABOWO", "Line Walker GI PATI", "9425952TJBT", "Line Walker"],
    ["APP-RMB-12", "UPT SEMARANG", "LW", "REMBANG", "GI REMBANG", "PRIHADI", "Line Walker GI REMBANG", "7825712TJBT", "Line Walker"],
    ["APP-RMB-13", "UPT SEMARANG", "LW", "REMBANG", "GI REMBANG", "SUGIONO", "Line Walker GI REMBANG", "9225920TJBT", "Line Walker"],

    // 4. UNIT: UPT SEMARANG - WILAYAH ULTG KUDUS
    ["APP-KDS-01", "UPT SEMARANG", "TL GI", "KUDUS", "GI PURWODADI", "SANDI PRATAMA RAKHMADI", "TL GI PURWODADI", "-", "Team Leader GI Purwodadi"],
    ["APP-KDS-02", "UPT SEMARANG", "TL GI", "KUDUS", "GI TANJUNG JATI", "DANIAL ZULKARNIANSYAH", "TL GI TANJUNG JATI", "-", "Team Leader GI Tanjung Jati"],
    ["APP-KDS-03", "UPT SEMARANG", "TL GI", "KUDUS", "GI KUDUS", "SANDI PRATAMA RAKHMADI", "TL GI KUDUS", "-", "Team Leader GI Kudus"],
    ["APP-KDS-04", "UPT SEMARANG", "KLW", "KUDUS", "GI PURWODADI", "BAMBANG RIYANTO", "Koordinator Line Walker KUDUS", "8425819TJBT", "Koordinator LW Purwodadi"],
    ["APP-KDS-05", "UPT SEMARANG", "KLW", "KUDUS", "GI TANJUNG JATI", "MUHAMMAD DENDI RENALDI", "Koordinator Line Walker KUDUS", "9925015TJBT", "Koordinator LW TJB"],
    ["APP-KDS-06", "UPT SEMARANG", "LW", "KUDUS", "GI PURWODADI", "AHMAD AMIN DHUKA", "Line Walker GI PURWODADI", "9125902TJBT", "Line Walker"],
    ["APP-KDS-07", "UPT SEMARANG", "LW", "KUDUS", "GI PURWODADI", "ATOK SUDARMANTO", "Line Walker GI PURWODADI", "8325806TJBT", "Line Walker"],
    ["APP-KDS-08", "UPT SEMARANG", "LW", "KUDUS", "GI PURWODADI", "BAMBANG HERNOWO", "Line Walker GI PURWODADI", "8425818TJBT", "Line Walker"],
    ["APP-KDS-09", "UPT SEMARANG", "LW", "KUDUS", "GI PURWODADI", "EKO SUYANTO", "Line Walker GI PURWODADI", "8125770TJBT", "Line Walker"],
    ["APP-KDS-10", "UPT SEMARANG", "LW", "KUDUS", "GI PURWODADI", "MUKAHAR", "Line Walker GI PURWODADI", "7725673TJBT", "Line Walker"],
    ["APP-KDS-11", "UPT SEMARANG", "LW", "KUDUS", "GI TANJUNG JATI", "AGUS HARIYANTO", "Line Walker GI TANJUNG JATI", "8625841TJBT", "Line Walker"],
    ["APP-KDS-12", "UPT SEMARANG", "LW", "KUDUS", "GI TANJUNG JATI", "AHMAD ASRO", "Line Walker GI TANJUNG JATI", "9325927TJBT", "Line Walker"],
    ["APP-KDS-13", "UPT SEMARANG", "LW", "KUDUS", "GI TANJUNG JATI", "AHMAD RIDWAN", "Line Walker GI TANJUNG JATI", "7325618TJBT", "Line Walker"],
    ["APP-KDS-14", "UPT SEMARANG", "LW", "KUDUS", "GI TANJUNG JATI", "ERWIN ARDI DIRTANA", "Line Walker GI TANJUNG JATI", "8125771TJBT", "Line Walker"],
    ["APP-KDS-15", "UPT SEMARANG", "LW", "KUDUS", "GI TANJUNG JATI", "DEDI SISWOYO", "Line Walker GI TANJUNG JATI", "9125903TJBT", "Line Walker"],
    ["APP-KDS-16", "UPT SEMARANG", "LW", "KUDUS", "GI TANJUNG JATI", "MUHAMMAD UFRON", "Line Walker GI TANJUNG JATI", "9125907TJBT", "Line Walker"],
    ["APP-KDS-17", "UPT SEMARANG", "LW", "KUDUS", "GI TANJUNG JATI", "SHOLIHUN", "Line Walker GI TANJUNG JATI", "8425824TJBT", "Line Walker"]
  ];
}