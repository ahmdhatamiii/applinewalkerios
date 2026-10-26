/**
 * PLN UPT Semarang - Backend API Handler
 * File: code.gs
 * Description: Core logic, CRUD operations, Live Data LW Synchronization,
 * Multi-Date Work Order SPKL, Form_Responses integration, and Presensi Analytics with Attendance Scoring.
 */

function doGet(e) {
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("Line Walker UPT Semarang - Management System")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0, viewport-fit=cover")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * API Router untuk menerima request dari frontend Vercel (Opsi Hybrid)
 */
function doPost(e) {
  try {
    var raw = (e && e.postData && e.postData.contents) ? e.postData.contents : "{}";
    var payload = JSON.parse(raw);
    var action = payload.action;
    var params = payload.params || [];

    if (!action) {
      throw new Error("Parameter 'action' wajib disertakan.");
    }

    // Akses global scope GAS
    var globalScope = this;
    if (typeof globalScope[action] === "function") {
      var result = globalScope[action].apply(globalScope, params);
      return ContentService.createTextOutput(JSON.stringify(result !== undefined ? result : { status: "success" }))
        .setMimeType(ContentService.MimeType.JSON);
    } else {
      throw new Error("Fungsi backend '" + action + "' tidak ditemukan.");
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function generateSequentialId(prefix, sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var todayStr = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyyMMdd");
  var fullPrefix = prefix + "-" + todayStr + "-";
  
  if (!sheet || sheet.getLastRow() <= 1) {
    return fullPrefix + "0001";
  }
  
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
  var maxNum = 0;
  
  for (var i = 0; i < values.length; i++) {
    var id = values[i][0];
    if (id && id.indexOf(fullPrefix) === 0) {
      var suffix = id.slice(fullPrefix.length);
      if (/^\d+$/.test(suffix)) maxNum = Math.max(maxNum, Number(suffix));
    }
  }
  
  var nextNum = maxNum + 1;
  var numPadded = String(nextNum).padStart(4, "0");
  return fullPrefix + numPadded;
}

function normalizeMonthString(str) {
  if (!str) return "";
  var s = str.toString().toUpperCase().trim();
  if (s.indexOf("OKT") !== -1 || s.indexOf("OCT") !== -1) return "OCTOBER";
  if (s.indexOf("SEP") !== -1) return "SEPTEMBER";
  if (s.indexOf("AGU") !== -1 || s.indexOf("AUG") !== -1) return "AGUSTUS";
  if (s.indexOf("JUL") !== -1) return "JULI";
  if (s.indexOf("JUN") !== -1) return "JUNI";
  if (s.indexOf("MEI") !== -1 || s.indexOf("MAY") !== -1) return "MEI";
  if (s.indexOf("APR") !== -1) return "APRIL";
  if (s.indexOf("MAR") !== -1) return "MARET";
  if (s.indexOf("FEB") !== -1) return "FEBRUARI";
  if (s.indexOf("JAN") !== -1) return "JANUARI";
  if (s.indexOf("NOV") !== -1) return "NOVEMBER";
  if (s.indexOf("DES") !== -1 || s.indexOf("DEC") !== -1) return "DESEMBER";
  return s;
}

function normalizeTanggal(str) {
  if (str === null || str === undefined) return "";
  var s = str.toString().trim();
  if (s === "") return "";
  var m = s.match(/^(\d{1,2})/);
  if (m) {
    return parseInt(m[1], 10).toString();
  }
  var num = parseInt(s, 10);
  return isNaN(num) ? s.toLowerCase() : num.toString();
}

function getGoogleFormResponses(searchQuery) {
  try {
    var formSsId = "1rdWG9RD_Av8lA-XCvvRdh8AESgNyb6NneD831ojzZ8M";
    var formSs = SpreadsheetApp.openById(formSsId);
    var sheet = formSs.getSheetByName("Form_Responses") || formSs.getSheets()[0];
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return { status: "success", data: [] };
    }

    var lastCol = sheet.getLastColumn();
    var rawData = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getDisplayValues();
    var headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];

    var search = (searchQuery || "").toLowerCase();
    var filtered = [];

    // Definisi Konstanta Pemetaan Kolom Baku (Sesuai Urutan Lembar Kerja):
    var COL_RESPON = {
      ID_BA: 0,
      JENIS_BA: 1,
      NAMA: 2,
      JABATAN: 3,
      NIP: 4,
      PENGHANTAR: 5,
      GARDU_INDUK: 6,
      ULTG: 7,
      TANGGAL: 8,
      BULAN: 9,
      KETERANGAN_KENDALA: 10,
      KLW: 11,
      TL_GI: 12,
      TANGGAL_SURAT: 13,
      STATUS: 14,
      TANGGAL_INPUT: 15
    };

    for (var i = rawData.length - 1; i >= 0; i--) {
      var row = rawData[i];
      var rowStr = row.join(" ").toLowerCase();

      if (search === "" || rowStr.indexOf(search) !== -1) {
        var mergedUrl = "#";
        for (var c = row.length - 1; c >= 0; c--) {
          if (row[c] && row[c].toString().indexOf("http") === 0) {
            mergedUrl = row[c];
            break;
          }
        }

        var driveId = extractDriveFileId(row[16]) || extractDriveFileId(row[17]) || extractDriveFileId(row[15]) || extractDriveFileId(mergedUrl);
        var previewUrl = "";
        var downloadUrl = "";
        var hasValidFile = false;

        if (driveId) {
          hasValidFile = true;
          previewUrl = "https://drive.google.com/file/d/" + driveId + "/preview";
          downloadUrl = "https://drive.google.com/uc?export=download&id=" + driveId;
        } else if (mergedUrl && mergedUrl !== "#" && mergedUrl.indexOf("http") === 0) {
          hasValidFile = true;
          previewUrl = mergedUrl;
          downloadUrl = mergedUrl;
        }

        var idBaVal = row[COL_RESPON.ID_BA] || ("BA-" + (i + 1));
        var statusVal = row[COL_RESPON.STATUS] || row[18] || row[17] || "Merged";
        var tglInputVal = row[COL_RESPON.TANGGAL_INPUT] || "-";

        filtered.push({
          rowId: i + 2,
          idBa: idBaVal,
          jenisKendala: row[COL_RESPON.JENIS_BA] || "-",
          jenisBa: row[COL_RESPON.JENIS_BA] || "-",
          nama: row[COL_RESPON.NAMA] || "-",
          jabatan: row[COL_RESPON.JABATAN] || "Line Walker",
          nip: row[COL_RESPON.NIP] || "-",
          penghantar: row[COL_RESPON.PENGHANTAR] || "-",
          garduInduk: row[COL_RESPON.GARDU_INDUK] || "-",
          ultg: row[COL_RESPON.ULTG] || "-",
          tanggal: row[COL_RESPON.TANGGAL] || "-",
          bulan: row[COL_RESPON.BULAN] || "-",
          keteranganKendala: row[COL_RESPON.KETERANGAN_KENDALA] || "-",
          klw: row[COL_RESPON.KLW] || "-",
          tlGi: row[COL_RESPON.TL_GI] || "-",
          tanggalSurat: row[COL_RESPON.TANGGAL_SURAT] || "-",
          status: statusVal,
          tanggalInput: tglInputVal,
          timestamp: tglInputVal,
          mergedDocId: driveId || "-",
          mergedDocUrl: mergedUrl,
          previewUrl: previewUrl,
          downloadUrl: downloadUrl,
          hasValidFile: hasValidFile,
          documentMergeStatus: statusVal,
          headers: headers,
          fullRow: row
        });
      }
    }

    return { status: "success", data: filtered };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function getDashboardAnalytics(periodMonths, selectedMonth, selectedYear) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetAcuan = ss.getSheetByName("AcuanKendala");
    var sheetBA = ss.getSheetByName("BeritaAcara");
    var sheetCeklis = ss.getSheetByName("Ceklis");
    var sheetLW = ss.getSheetByName("Data LW");
    var sheetSP = getSuratPeringatanSheet(ss);

    var lwMap = {};
    if (sheetLW && sheetLW.getLastRow() > 1) {
      var lwData = sheetLW.getRange(2, 1, sheetLW.getLastRow() - 1, Math.max(sheetLW.getLastColumn(), 8)).getDisplayValues();
      for (var i = 0; i < lwData.length; i++) {
        var namaLW = (lwData[i][0] || "").trim();
        if (namaLW !== "") {
          lwMap[namaLW] = {
            idLw: lwData[i][2] || ("LW-" + (i + 1)),
            nama: namaLW,
            koordinator: lwData[i][1] || "-",
            posisi: lwData[i][3] || "Line Walker",
            ultg: lwData[i][4] || "-",
            gi: lwData[i][5] || "-",
            tlGi: lwData[i][6] || "-",
            penghantar: lwData[i][7] || "-"
          };
        }
      }
    }

    var totalAcuanKendala = 0;
    var totalBeritaAcara = 0;
    var belumTerbackupBA = 0;
    var totalSelesaiBA = 0;
    
    var monthNames = ["JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", "JULI", "AGUSTUS", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DESEMBER"];
    var filterMonthName = "ALL";
    if (selectedMonth && selectedMonth !== "ALL") {
      var mIndex = parseInt(selectedMonth, 10) - 1;
      if (mIndex >= 0 && mIndex < 12) {
        filterMonthName = monthNames[mIndex];
      } else {
        filterMonthName = selectedMonth.toString().toUpperCase();
      }
    }

    // Peta Ceklis untuk verifikasi status Ceklis & Upload PDF (spesifik per tanggal kendala)
    var ceklisStatusMap = {};
    if (sheetCeklis && sheetCeklis.getLastRow() > 1) {
      var lastColC = Math.max(sheetCeklis.getLastColumn(), 11);
      var rawCeklis = sheetCeklis.getRange(2, 1, sheetCeklis.getLastRow() - 1, lastColC).getDisplayValues();
      for (var c = 0; c < rawCeklis.length; c++) {
        var cNama = (rawCeklis[c][1] || "").trim().toLowerCase();
        var cNip = (rawCeklis[c][2] || "").trim();
        var cBulan = normalizeMonthString(rawCeklis[c][3]);
        var cTahun = (rawCeklis[c][4] || "").toString();
        var cCat = (rawCeklis[c][5] || "").trim().toUpperCase();
        var isChecked = (rawCeklis[c][6] || "").toUpperCase() === "TRUE";
        var fileUrl = rawCeklis[c][8] || "";
        var fileId = rawCeklis[c][9] || "";
        var cTanggal = rawCeklis[c].length >= 11 ? normalizeTanggal(rawCeklis[c][10]) : "";

        var cleanCat = cCat.replace(/^BERITA\s+ACARA\s+/i, "").replace(/^KENDALA\s+/i, "").trim();
        var statusObj = { isCeklis: isChecked, fileBaUrl: fileUrl, fileBaId: fileId, cTanggal: cTanggal };
        var validNipC = (cNip && cNip !== "000000" && cNip !== "-");

        if (cTanggal) {
          if (validNipC) {
            ceklisStatusMap[(cNip + "|" + cTanggal + "|" + cBulan + "|" + cTahun + "|" + cCat).toLowerCase()] = statusObj;
            ceklisStatusMap[(cNip + "|" + cTanggal + "|" + cBulan + "|" + cTahun + "|" + cleanCat).toLowerCase()] = statusObj;
            ceklisStatusMap[(cNip + "|" + cTanggal + "|" + cBulan + "|" + cleanCat).toLowerCase()] = statusObj;
            ceklisStatusMap[(cNip + "|" + cTanggal + "|" + cBulan).toLowerCase()] = statusObj;
          }
          if (cNama) {
            ceklisStatusMap[(cNama + "|" + cTanggal + "|" + cBulan + "|" + cTahun + "|" + cCat).toLowerCase()] = statusObj;
            ceklisStatusMap[(cNama + "|" + cTanggal + "|" + cBulan + "|" + cTahun + "|" + cleanCat).toLowerCase()] = statusObj;
            ceklisStatusMap[(cNama + "|" + cTanggal + "|" + cBulan + "|" + cleanCat).toLowerCase()] = statusObj;
            ceklisStatusMap[(cNama + "|" + cTanggal + "|" + cBulan).toLowerCase()] = statusObj;
          }
        }
      }
    }

    var acuanList = [];
    if (sheetAcuan && sheetAcuan.getLastRow() > 1) {
      var acuanRaw = sheetAcuan.getRange(2, 1, sheetAcuan.getLastRow() - 1, 14).getDisplayValues();
      for (var a = 0; a < acuanRaw.length; a++) {
        var jenisBa = (acuanRaw[a][0] || "").trim().toUpperCase();
        var namaAcuan = (acuanRaw[a][1] || "").trim();
        var nipAcuan = (acuanRaw[a][3] || "").trim();
        var tglAcuan = acuanRaw[a][7];
        var bulanAcuan = (acuanRaw[a][8] || "").toUpperCase();
        var tglSuratAcuan = acuanRaw[a][12] || "";

        var yearMatch = tglSuratAcuan.match(/\b(20\d\d)\b/);
        var itemYear = yearMatch ? yearMatch[1] : "2026";

        var matchMonth = (filterMonthName === "ALL" || bulanAcuan.indexOf(filterMonthName) !== -1 || normalizeMonthString(bulanAcuan) === normalizeMonthString(filterMonthName));
        var matchYear = (selectedYear === "ALL" || !selectedYear || itemYear === selectedYear.toString());

        if ((jenisBa || namaAcuan) && matchMonth && matchYear) {
          totalAcuanKendala++;
          acuanList.push({
            jenisBA: jenisBa,
            nama: namaAcuan,
            nip: nipAcuan,
            tanggal: tglAcuan,
            bulan: bulanAcuan,
            tahun: itemYear,
            isBackedUp: false
          });
        }
      }
    }

    // Peta Berita Acara untuk pencocokan apakah item di Ceklis sudah masuk Berita Acara (spesifik per tanggal)
    var baMap = {};
    var totalBeritaAcara = 0;
    if (sheetBA && sheetBA.getLastRow() > 1) {
      var baData = sheetBA.getRange(2, 1, sheetBA.getLastRow() - 1, sheetBA.getLastColumn()).getDisplayValues();

      for (var j = 0; j < baData.length; j++) {
        var row = baData[j];
        var idBA = row[0];
        var jenisBA = (row[1] || "").trim().toUpperCase();
        var namaLW = (row[2] || "").trim();
        var nipLW = (row[4] || "").trim();
        var baTanggal = normalizeTanggal(row[8]);
        var bulanBA = (row[9] || "").toUpperCase();
        var bulanNormBA = normalizeMonthString(bulanBA);
        var tglSuratBA = row[13] || "";
        var status = (row[14] || "").trim().toLowerCase();
        var fileSignedPdfUrl = row[16] || "";
        var fileSignedPdfId = row[17] || "";

        var yearMatchBA = tglSuratBA.match(/\b(20\d\d)\b/);
        var itemYearBA = yearMatchBA ? yearMatchBA[1] : "2026";

        var matchMonthBA = (filterMonthName === "ALL" || bulanBA.indexOf(filterMonthName) !== -1 || normalizeMonthString(bulanBA) === normalizeMonthString(filterMonthName));
        var matchYearBA = (selectedYear === "ALL" || !selectedYear || itemYearBA === selectedYear.toString());

        if (matchMonthBA && matchYearBA && status !== "dibatalkan") {
          totalBeritaAcara++;
          var hasPdf = (fileSignedPdfUrl && fileSignedPdfUrl !== "#" && fileSignedPdfUrl.indexOf("http") === 0) || (fileSignedPdfId && fileSignedPdfId !== "");
          var baObj = {
            idBA: idBA,
            jenisBA: jenisBA,
            nama: namaLW,
            nip: nipLW,
            tanggal: baTanggal,
            status: status,
            hasPdf: hasPdf,
            fileUrl: fileSignedPdfUrl,
            fileId: fileSignedPdfId
          };

          var cleanJenisBA = jenisBA.replace(/^BERITA\s+ACARA\s+/i, "").replace(/^KENDALA\s+/i, "").trim();
          var validNipBA = (nipLW && nipLW !== "000000" && nipLW !== "-");

          if (baTanggal) {
            if (validNipBA) {
              baMap[(nipLW + "|" + baTanggal + "|" + bulanNormBA + "|" + itemYearBA + "|" + jenisBA).toLowerCase()] = baObj;
              baMap[(nipLW + "|" + baTanggal + "|" + bulanNormBA + "|" + itemYearBA + "|" + cleanJenisBA).toLowerCase()] = baObj;
              baMap[(nipLW + "|" + baTanggal + "|" + bulanNormBA + "|" + cleanJenisBA).toLowerCase()] = baObj;
              baMap[(nipLW + "|" + baTanggal + "|" + bulanNormBA).toLowerCase()] = baObj;
            }
            if (namaLW) {
              baMap[(namaLW.toLowerCase() + "|" + baTanggal + "|" + bulanNormBA + "|" + itemYearBA + "|" + jenisBA).toLowerCase()] = baObj;
              baMap[(namaLW.toLowerCase() + "|" + baTanggal + "|" + bulanNormBA + "|" + itemYearBA + "|" + cleanJenisBA).toLowerCase()] = baObj;
              baMap[(namaLW.toLowerCase() + "|" + baTanggal + "|" + bulanNormBA + "|" + cleanJenisBA).toLowerCase()] = baObj;
              baMap[(namaLW.toLowerCase() + "|" + baTanggal + "|" + bulanNormBA).toLowerCase()] = baObj;
            }
          }
        }
      }
    }

    // HITUNG DATA KENDALA DI CEKLIS:
    // Setiap data kendala yang SUDAH DICEKLIS DAN SUDAH MASUK BERITA ACARA TERHITUNG 1 (berdasarkan tanggal spesifik)
    var lwKendalaCount = {};
    for (var u = 0; u < acuanList.length; u++) {
      var ac = acuanList[u];
      var acTanggalNorm = normalizeTanggal(ac.tanggal);
      var acBulanNorm = normalizeMonthString(ac.bulan);
      var acJenisNorm = (ac.jenisBA || "").trim().toUpperCase();
      var acCleanJenis = acJenisNorm.replace(/^BERITA\s+ACARA\s+/i, "").replace(/^KENDALA\s+/i, "").trim();
      var acNamaLow = (ac.nama || "").trim().toLowerCase();
      var acNip = (ac.nip || "").trim();
      var acTahun = ac.tahun || "2026";
      var validNipAc = (acNip && acNip !== "000000" && acNip !== "-");

      // 1. Ceklis status (spesifik per tanggal)
      var cKeyNip = (acNip + "|" + acTanggalNorm + "|" + acBulanNorm + "|" + acTahun + "|" + acJenisNorm).toLowerCase();
      var cKeyCleanNip = (acNip + "|" + acTanggalNorm + "|" + acBulanNorm + "|" + acTahun + "|" + acCleanJenis).toLowerCase();
      var cKeyNama = (acNamaLow + "|" + acTanggalNorm + "|" + acBulanNorm + "|" + acTahun + "|" + acJenisNorm).toLowerCase();
      var cKeyCleanNama = (acNamaLow + "|" + acTanggalNorm + "|" + acBulanNorm + "|" + acTahun + "|" + acCleanJenis).toLowerCase();
      var cKeyNamaJenis = (acNamaLow + "|" + acTanggalNorm + "|" + acCleanJenis).toLowerCase();
      var cKeyCleanNipJenis = (acNip + "|" + acTanggalNorm + "|" + acCleanJenis).toLowerCase();
      var cKeyNipBulan = (acNip + "|" + acTanggalNorm + "|" + acBulanNorm).toLowerCase();
      var cKeyNamaBulan = (acNamaLow + "|" + acTanggalNorm + "|" + acBulanNorm).toLowerCase();

      var ceklisInfo = (validNipAc ? (ceklisStatusMap[cKeyNip] || ceklisStatusMap[cKeyCleanNip] || ceklisStatusMap[cKeyCleanNipJenis] || ceklisStatusMap[cKeyNipBulan]) : null) ||
                       ceklisStatusMap[cKeyNama] || ceklisStatusMap[cKeyCleanNama] || ceklisStatusMap[cKeyNamaJenis] || ceklisStatusMap[cKeyNamaBulan] || null;
      var isCeklis = ceklisInfo ? ceklisInfo.isCeklis : false;

      // 2. Berita Acara status (spesifik per tanggal)
      var baInfo = (validNipAc ? (baMap[cKeyNip] || baMap[cKeyCleanNip] || baMap[cKeyCleanNipJenis] || baMap[cKeyNipBulan]) : null) ||
                   baMap[cKeyNama] || baMap[cKeyCleanNama] || baMap[cKeyNamaJenis] || baMap[cKeyNamaBulan] || null;

      var hasFileInCeklis = ceklisInfo && ((ceklisInfo.fileBaUrl && ceklisInfo.fileBaUrl !== "#" && ceklisInfo.fileBaUrl.indexOf("http") === 0) || (ceklisInfo.fileBaId && ceklisInfo.fileBaId !== ""));
      var masukBA = Boolean(baInfo || hasFileInCeklis);

      if (baInfo && (baInfo.status === "selesai" || baInfo.hasPdf)) {
        isCeklis = true;
      }

      // KETIKA SUDAH DICEKLIS DAN MASUK BERITA ACARA -> TERHITUNG 1
      var isTuntas = Boolean(isCeklis && masukBA);

      if (isTuntas) {
        totalSelesaiBA++; // Terhitung 1
        ac.isBackedUp = true;
      } else {
        belumTerbackupBA++;
        ac.isBackedUp = false;
      }

      var namaLW = ac.nama || "-";
      if (!lwKendalaCount[namaLW]) {
        lwKendalaCount[namaLW] = { total: 0, pending: 0, selesai: 0 };
      }
      lwKendalaCount[namaLW].total++;
      if (isTuntas) {
        lwKendalaCount[namaLW].selesai++;
      } else {
        lwKendalaCount[namaLW].pending++;
      }
    }

    // Jika acuan kendala kosong tapi ada berita acara terupload, fallback ke totalBeritaAcara
    var targetTotalKendala = totalAcuanKendala > 0 ? totalAcuanKendala : totalBeritaAcara;
    if (totalAcuanKendala === 0 && totalBeritaAcara > 0) {
      totalSelesaiBA = totalBeritaAcara;
      belumTerbackupBA = 0;
    }

    var lwList = [];
    for (var n in lwKendalaCount) {
      var info = lwMap[n] || { idLw: "LW-000", nama: n, koordinator: "-" };
      lwList.push({
        idLw: info.idLw,
        nama: n,
        koordinator: info.koordinator,
        total: lwKendalaCount[n].total,
        selesai: lwKendalaCount[n].selesai,
        pending: lwKendalaCount[n].pending
      });
    }

    lwList.sort(function(a, b) { return b.pending - a.pending; });

    // Target SLA Rate: Persentase Berita Acara yang tuntas dari total kendala di Ceklis
    var completionRate = targetTotalKendala > 0 ? Math.round((totalSelesaiBA / targetTotalKendala) * 100) : 0;
    var completionStats = {
      rate: completionRate,
      selesai: totalSelesaiBA,
      total: targetTotalKendala,
      isSemuaSelesai: (targetTotalKendala > 0 && totalSelesaiBA >= targetTotalKendala)
    };

    var worstLW = lwList.length > 0 ? lwList[0] : { nama: "-", total: 0, pending: 0, selesai: 0 };

    // Agregasi Data Dashboard Surat Peringatan (ST & SP)
    var spSummary = {
      totalAktif: 0,
      totalKadaluarsa: 0,
      stCount: 0,
      sp1Count: 0,
      sp2Count: 0,
      sp3Count: 0,
      spTerakhirCount: 0
    };

    if (sheetSP && sheetSP.getLastRow() > 1) {
      var spData = sheetSP.getRange(2, 1, sheetSP.getLastRow() - 1, Math.min(sheetSP.getLastColumn(), 13)).getDisplayValues();
      var today = new Date();
      today.setHours(0, 0, 0, 0);

      for (var s = 0; s < spData.length; s++) {
        var spRow = spData[s];
        var namaLW = (spRow[1] || "").trim();
        var idSp = (spRow[0] || "").trim();
        // Lewati baris kosong / ghost rows
        if (!namaLW && !idSp) continue;

        var jSurat = (spRow[3] || "ST").trim().toUpperCase();
        var tglBerakhirStr = spRow[7];
        var endDate = parseDateFlexible(tglBerakhirStr);
        endDate.setHours(23, 59, 59, 999);

        var diffTime = endDate.getTime() - today.getTime();
        var isAktif = diffTime >= 0;

        var rawStatus = (spRow[11] || "").toLowerCase();
        if (rawStatus.indexOf("kadaluarsa") !== -1 || rawStatus.indexOf("eskalasi") !== -1 || rawStatus.indexOf("selesai") !== -1 || rawStatus.indexOf("bebas") !== -1) {
          isAktif = false;
        }

        if (isAktif) {
          spSummary.totalAktif++;
          if (jSurat === "ST") spSummary.stCount++;
          else if (jSurat === "SP1") spSummary.sp1Count++;
          else if (jSurat === "SP2") spSummary.sp2Count++;
          else if (jSurat === "SP3") spSummary.sp3Count++;
          else if (jSurat.indexOf("TERAKHIR") !== -1 || jSurat === "SP1&TERAKHIR") spSummary.spTerakhirCount++;
        } else {
          spSummary.totalKadaluarsa++;
        }
      }
    }

    return {
      status: "success",
      totalKendala: targetTotalKendala,
      totalBA: totalSelesaiBA,
      totalSelesaiBA: totalSelesaiBA,
      totalDokumenBA: totalBeritaAcara,
      belumTerbackupBA: belumTerbackupBA,
      worstLW: worstLW,
      rankingLW: lwList.slice(0, 5),
      completionStats: completionStats,
      spSummary: spSummary,
      acuanList: acuanList
    };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function createAcuanKendalaInput(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetAcuan = ss.getSheetByName("AcuanKendala");
    if (!sheetAcuan) {
      sheetAcuan = ss.insertSheet("AcuanKendala");
    }

    var sheetBA = ss.getSheetByName("BeritaAcara");
    if (!sheetBA) {
      sheetBA = ss.insertSheet("BeritaAcara");
    }

    var sheetLW = ss.getSheetByName("Data LW");
    var masterMap = {};

    if (sheetLW && sheetLW.getLastRow() > 1) {
      var lastCol = Math.max(sheetLW.getLastColumn(), 8);
      var lwData = sheetLW.getRange(2, 1, sheetLW.getLastRow() - 1, lastCol).getDisplayValues();
      for (var i = 0; i < lwData.length; i++) {
        var namaLw = (lwData[i][0] || "").trim();
        if (namaLw !== "") {
          masterMap[namaLw.toLowerCase()] = {
            klw: lwData[i][1] || "-",
            nip: lwData[i][2] || "-",
            posisi: lwData[i][3] || payload.jabatan || "Line Walker",
            ultg: lwData[i][4] || "-",
            gi: lwData[i][5] || "-",
            tlGi: lwData[i][6] || "-",
            penghantar: lwData[i][7] || "-"
          };
        }
      }
    }

    var targetName = (payload.nama || "").trim().toLowerCase();
    var matchInfo = masterMap[targetName] || {
      nip: "000000",
      posisi: payload.jabatan || "Line Walker",
      klw: "-",
      ultg: "-",
      gi: "-",
      tlGi: "-",
      penghantar: "-"
    };

    var now = new Date();
    var todaySuratStr = Utilities.formatDate(now, "Asia/Jakarta", "dd MMMM yyyy");
    var nowStr = Utilities.formatDate(now, "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");
    var monthNames = ["JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", "JULI", "AGUSTUS", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DESEMBER"];
    var currentMonthName = monthNames[now.getMonth()];

    var rowData = [
      payload.jenisBA || "BERITA ACARA KENDALA ABSENSI & COC",
      payload.nama,
      matchInfo.posisi,
      matchInfo.nip,
      matchInfo.penghantar,
      matchInfo.gi,
      matchInfo.ultg,
      payload.tanggalKendala || now.getDate().toString(),
      currentMonthName,
      payload.keteranganKendala || "-",
      matchInfo.klw,
      matchInfo.tlGi,
      todaySuratStr,
      "-"
    ];

    sheetAcuan.appendRow(rowData);

    var newBaId = generateSequentialId("BA", "BeritaAcara");

    sheetBA.appendRow([
      newBaId,
      payload.jenisBA || "BERITA ACARA KENDALA ABSENSI & COC",
      payload.nama,
      matchInfo.posisi,
      matchInfo.nip,
      matchInfo.penghantar,
      matchInfo.gi,
      matchInfo.ultg,
      payload.tanggalKendala || now.getDate().toString(),
      currentMonthName,
      payload.keteranganKendala || "-",
      matchInfo.klw,
      matchInfo.tlGi,
      todaySuratStr,
      "Selesai",
      nowStr,
      "#",
      ""
    ]);

    SpreadsheetApp.flush();

    return { 
      status: "success", 
      message: "Berhasil menginput Kendala BA " + newBaId + "!",
      idBA: newBaId
    };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function getAcuanKendala(searchQuery, kendalaFilter, roleFilter, monthFilter, yearFilter, page, limit, statusBaFilter) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("AcuanKendala");
    var sheetBA = ss.getSheetByName("BeritaAcara");
    if (!sheet || sheet.getLastRow() <= 1) {
      return { data: [], totalPages: 1, currentPage: 1, totalItems: 0 };
    }

    var usedAcuanMap = {};
    var baPdfMap = {};
    if (sheetBA && sheetBA.getLastRow() > 1) {
      var baData = sheetBA.getRange(2, 1, sheetBA.getLastRow() - 1, 18).getDisplayValues();
      for (var b = 0; b < baData.length; b++) {
        var baRow = baData[b];
        var baKey = (baRow[1] + "|" + baRow[2] + "|" + baRow[5] + "|" + baRow[6] + "|" + baRow[8]).toLowerCase();
        usedAcuanMap[baKey] = true;
        baPdfMap[baKey] = baRow[16] || "#";
      }
    }

    var rawData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getDisplayValues();
    var filtered = [];
    var search = (searchQuery || "").toLowerCase();
    var kendala = (kendalaFilter || "ALL").toLowerCase();
    var role = (roleFilter || "ALL").toLowerCase();
    var monthNorm = normalizeMonthString(monthFilter);
    var year = (yearFilter || "ALL").toString();
    var statusBa = (statusBaFilter || "ALL").toUpperCase();

    for (var i = 0; i < rawData.length; i++) {
      var jenisBa = rawData[i][0];
      var nama = rawData[i][1];
      var jabatan = rawData[i][2];
      var nip = rawData[i][3];
      var penghantar = rawData[i][4];
      var gi = rawData[i][5];
      var ultg = rawData[i][6];
      var tanggal = rawData[i][7];
      var bulan = (rawData[i][8] || "").toUpperCase();
      var bulanNorm = normalizeMonthString(bulan);
      var ketKendala = rawData[i][9];
      var klw = rawData[i][10];
      var tlGi = rawData[i][11];
      var tglSurat = rawData[i][12];
      var keterangan = rawData[i][13];

      var yearMatch = tglSurat.match(/\b(20\d\d)\b/);
      var itemYear = yearMatch ? yearMatch[1] : "2026";

      var matchKendala = (kendala === "all" || jenisBa.toLowerCase().indexOf(kendala) !== -1 || kendala.indexOf(jenisBa.toLowerCase()) !== -1);
      var matchRole = (role === "all" || jabatan.toLowerCase().indexOf(role) !== -1 || (role.indexOf("koordinator") !== -1 && klw.toLowerCase().indexOf(search) !== -1));
      var matchMonth = (monthFilter === "ALL" || monthNorm === "" || bulanNorm === monthNorm || bulan.indexOf(monthFilter.toString().toUpperCase()) !== -1);
      var matchYear = (year === "ALL" || itemYear === year || year === "");

      var searchBlob = (jenisBa + " " + nama + " " + nip + " " + penghantar + " " + gi + " " + ultg + " " + ketKendala + " " + klw + " " + tlGi + " " + bulan + " " + itemYear).toLowerCase();
      var matchSearch = (searchBlob.indexOf(search) !== -1);

      if (matchKendala && matchRole && matchMonth && matchYear && matchSearch) {
        var acKey = (jenisBa + "|" + nama + "|" + penghantar + "|" + gi + "|" + tanggal).toLowerCase();
        var hasBA = !!usedAcuanMap[acKey];
        var pdfUrl = baPdfMap[acKey] || "#";

        var matchStatusBa = (statusBa === "ALL") || 
                            (statusBa === "BELUM_BA" && !hasBA) || 
                            (statusBa === "SUDAH_BA" && hasBA);

        if (matchStatusBa) {
          filtered.push({
            rowId: i + 2,
            jenisBA: jenisBa,
            nama: nama,
            jabatan: jabatan,
            nip: nip,
            penghantar: penghantar,
            garduInduk: gi,
            ultg: ultg,
            tanggal: tanggal,
            bulan: bulan,
            tahun: itemYear,
            keteranganKendala: ketKendala,
            klw: klw,
            tlGi: tlGi,
            tanggalSurat: tglSurat,
            keterangan: keterangan,
            hasBA: hasBA,
            fileSignedPdfUrl: pdfUrl
          });
        }
      }
    }

    var p = page || 1;
    var l = limit || 10;
    var totalItems = filtered.length;
    var totalPages = Math.ceil(totalItems / l) || 1;
    var startIndex = (p - 1) * l;
    var paginatedData = filtered.slice(startIndex, startIndex + l);

    return {
      status: "success",
      data: paginatedData,
      allFilteredData: filtered,
      totalPages: totalPages,
      currentPage: p,
      totalItems: totalItems
    };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function getBeritaAcaraList(searchQuery, statusFilter, monthFilter, yearFilter, page, limit, jenisBAFilter) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetBA = ss.getSheetByName("BeritaAcara");

    if (!sheetBA || sheetBA.getLastRow() <= 1) {
      return { data: [], totalPages: 1, currentPage: 1, totalItems: 0 };
    }

    var lastCol = Math.max(sheetBA.getLastColumn(), 18);
    var rawData = sheetBA.getRange(2, 1, sheetBA.getLastRow() - 1, lastCol).getDisplayValues();
    var filtered = [];
    var search = (searchQuery || "").toLowerCase();
    var st = (statusFilter || "ALL").toLowerCase();
    var monthNorm = normalizeMonthString(monthFilter);
    var year = (yearFilter || "ALL").toString();
    var jenisBA = (jenisBAFilter || "ALL").toLowerCase();

    for (var i = rawData.length - 1; i >= 0; i--) {
      var idBa = (rawData[i][0] || "").toLowerCase();
      var jenisBa = rawData[i][1] || "";
      var nama = rawData[i][2] || "";
      var jabatan = rawData[i][3] || "";
      var nip = rawData[i][4] || "";
      var penghantar = rawData[i][5] || "";
      var gi = rawData[i][6] || "";
      var ultg = rawData[i][7] || "";
      var tanggal = rawData[i][8] || "";
      var bulan = (rawData[i][9] || "").toUpperCase();
      var bulanNorm = normalizeMonthString(bulan);
      var ketKendala = rawData[i][10] || "";
      var klw = rawData[i][11] || "";
      var tlGi = rawData[i][12] || "";
      var tglSurat = rawData[i][13] || "";
      var status = rawData[i][14] || "";
      var tglInput = rawData[i][15] || "";
      var fileSignedPdfUrl = rawData[i][16] || "";
      var fileSignedPdfId = rawData[i][17] || "";

      var yearMatch = (tglSurat + " " + tglInput).match(/\b(20\d\d)\b/);
      var itemYear = yearMatch ? yearMatch[1] : "2026";

      var matchStatus = (st === "all" || status.toLowerCase() === st);
      var matchMonth = (monthFilter === "ALL" || monthNorm === "" || bulanNorm === monthNorm || bulan.indexOf(monthFilter.toString().toUpperCase()) !== -1);
      var matchYear = (year === "ALL" || itemYear === year || year === "");
      var matchJenis = (jenisBA === "all" || jenisBa.toLowerCase().indexOf(jenisBA) !== -1 || jenisBA.indexOf(jenisBa.toLowerCase()) !== -1);

      var searchBlob = (idBa + " " + jenisBa + " " + nama + " " + nip + " " + penghantar + " " + gi + " " + ultg + " " + ketKendala + " " + klw + " " + tlGi + " " + bulan + " " + itemYear).toLowerCase();
      var matchSearch = (searchBlob.indexOf(search) !== -1);

      if (matchStatus && matchMonth && matchYear && matchJenis && matchSearch) {
        var driveId = extractDriveFileId(fileSignedPdfId) || extractDriveFileId(fileSignedPdfUrl);
        var previewUrl = "";
        var downloadUrl = "";
        var hasValidFile = false;

        if (driveId) {
          hasValidFile = true;
          previewUrl = "https://drive.google.com/file/d/" + driveId + "/preview";
          downloadUrl = "https://drive.google.com/uc?export=download&id=" + driveId;
        } else if (fileSignedPdfUrl && fileSignedPdfUrl !== "#" && fileSignedPdfUrl.indexOf("http") === 0) {
          hasValidFile = true;
          previewUrl = fileSignedPdfUrl;
          downloadUrl = fileSignedPdfUrl;
        }

        filtered.push({
          rowId: i + 2,
          idBA: rawData[i][0],
          jenisBA: jenisBa,
          nama: nama,
          jabatan: jabatan,
          nip: nip,
          penghantar: penghantar,
          garduInduk: gi,
          ultg: ultg,
          tanggal: tanggal,
          bulan: bulan,
          tahun: itemYear,
          keteranganKendala: ketKendala,
          klw: klw,
          tlGi: tlGi,
          tanggalSurat: tglSurat,
          status: status,
          tanggalInput: tglInput,
          fileSignedPdfUrl: fileSignedPdfUrl,
          fileSignedPdfId: fileSignedPdfId,
          previewUrl: previewUrl,
          downloadUrl: downloadUrl || fileSignedPdfUrl,
          hasValidFile: hasValidFile
        });
      }
    }

    var p = page || 1;
    var l = limit || 10;
    var totalItems = filtered.length;
    var totalPages = Math.ceil(totalItems / l) || 1;
    var startIndex = (p - 1) * l;
    var paginatedData = filtered.slice(startIndex, startIndex + l);

    return {
      status: "success",
      data: paginatedData,
      allFilteredData: filtered,
      totalPages: totalPages,
      currentPage: p,
      totalItems: totalItems
    };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function getCeklisData(ultgFilter, monthFilter, yearFilter, searchQuery, categoryFilter, weekFilter, klwFilter, giFilter) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetAcuan = ss.getSheetByName("AcuanKendala");
    var sheetCeklis = ss.getSheetByName("Ceklis");
    var sheetBA = ss.getSheetByName("BeritaAcara");

    if (!sheetAcuan || sheetAcuan.getLastRow() <= 1) {
      return { status: "success", data: [], checkedCount: 0, totalCount: 0, klwList: [], giList: [] };
    }

    // 1. Peta Berita Acara (Sheet BeritaAcara) - mendeteksi tanggal mana yang sudah mengirimkan/menerbitkan BA
    var baStatusMap = {};
    if (sheetBA && sheetBA.getLastRow() > 1) {
      var lastColBA = Math.max(sheetBA.getLastColumn(), 18);
      var rawBA = sheetBA.getRange(2, 1, sheetBA.getLastRow() - 1, lastColBA).getDisplayValues();
      for (var b = 0; b < rawBA.length; b++) {
        var bId = rawBA[b][0] || "";
        var bJenis = (rawBA[b][1] || "").trim().toUpperCase();
        var bNama = (rawBA[b][2] || "").trim().toLowerCase();
        var bNip = (rawBA[b][4] || "").trim();
        var bTanggal = normalizeTanggal(rawBA[b][8]);
        var bBulan = normalizeMonthString(rawBA[b][9]);
        var bTglSurat = rawBA[b][13] || "";
        var yMatchBA = bTglSurat.match(/\b(20\d\d)\b/);
        var bTahun = yMatchBA ? yMatchBA[1] : "2026";
        var bStatus = (rawBA[b][14] || "").trim().toLowerCase();
        var bUpdate = rawBA[b][15] || "";
        var bFileUrl = rawBA[b][16] || "";
        var bFileId = rawBA[b][17] || "";

        if (bStatus !== "dibatalkan" && bTanggal !== "") {
          var cleanCatBA = bJenis.replace(/^BERITA\s+ACARA\s+/i, "").replace(/^KENDALA\s+/i, "").trim();
          var baObj = {
            isCeklis: true,
            tanggalUpdate: bUpdate,
            fileBaUrl: bFileUrl,
            fileBaId: bFileId,
            idBA: bId
          };

          var validNipBA = (bNip && bNip !== "000000" && bNip !== "-");
          if (validNipBA) {
            baStatusMap[(bNip + "|" + bTanggal + "|" + bBulan + "|" + bTahun + "|" + bJenis).toLowerCase()] = baObj;
            baStatusMap[(bNip + "|" + bTanggal + "|" + bBulan + "|" + bTahun + "|" + cleanCatBA).toLowerCase()] = baObj;
            baStatusMap[(bNip + "|" + bTanggal + "|" + bBulan + "|" + bJenis).toLowerCase()] = baObj;
            baStatusMap[(bNip + "|" + bTanggal + "|" + bBulan + "|" + cleanCatBA).toLowerCase()] = baObj;
            baStatusMap[(bNip + "|" + bTanggal + "|" + bBulan).toLowerCase()] = baObj;
          }
          if (bNama) {
            baStatusMap[(bNama + "|" + bTanggal + "|" + bBulan + "|" + bTahun + "|" + bJenis).toLowerCase()] = baObj;
            baStatusMap[(bNama + "|" + bTanggal + "|" + bBulan + "|" + bTahun + "|" + cleanCatBA).toLowerCase()] = baObj;
            baStatusMap[(bNama + "|" + bTanggal + "|" + bBulan + "|" + bJenis).toLowerCase()] = baObj;
            baStatusMap[(bNama + "|" + bTanggal + "|" + bBulan + "|" + cleanCatBA).toLowerCase()] = baObj;
            baStatusMap[(bNama + "|" + bTanggal + "|" + bBulan).toLowerCase()] = baObj;
          }
        }
      }
    }

    // 2. Peta Ceklis Manual/Upload (Sheet Ceklis)
    var ceklisMap = {};
    if (sheetCeklis && sheetCeklis.getLastRow() > 1) {
      var lastColC = Math.max(sheetCeklis.getLastColumn(), 11);
      var rawCeklis = sheetCeklis.getRange(2, 1, sheetCeklis.getLastRow() - 1, lastColC).getDisplayValues();
      for (var c = 0; c < rawCeklis.length; c++) {
        var cNama = (rawCeklis[c][1] || "").trim().toLowerCase();
        var cNip = (rawCeklis[c][2] || "").trim();
        var cBulan = normalizeMonthString(rawCeklis[c][3]);
        var cTahun = (rawCeklis[c][4] || "").toString();
        var cCat = (rawCeklis[c][5] || "").trim().toUpperCase();
        var isChecked = (rawCeklis[c][6] || "").toUpperCase() === "TRUE";
        var cUpdate = rawCeklis[c][7] || "";
        var cFileUrl = rawCeklis[c][8] || "";
        var cFileId = rawCeklis[c][9] || "";
        var cTanggal = rawCeklis[c].length >= 11 ? normalizeTanggal(rawCeklis[c][10]) : "";

        // Jika cTanggal kosong pada baris lama, cari tanggal dari sheetBA berdasarkan cFileId
        if (!cTanggal && cFileId && baStatusMap) {
          for (var k in baStatusMap) {
            if (baStatusMap[k] && baStatusMap[k].fileBaId === cFileId) {
              var parts = k.split("|");
              if (parts.length >= 2 && parts[1]) {
                cTanggal = parts[1];
                break;
              }
            }
          }
        }

        var cleanCatC = cCat.replace(/^BERITA\s+ACARA\s+/i, "").replace(/^KENDALA\s+/i, "").trim();
        var statusObj = {
          isCeklis: isChecked,
          tanggalUpdate: cUpdate,
          fileBaUrl: cFileUrl,
          fileBaId: cFileId
        };

        var validNipC = (cNip && cNip !== "000000" && cNip !== "-");

        if (cTanggal) {
          if (validNipC) {
            ceklisMap[(cNip + "|" + cTanggal + "|" + cBulan + "|" + cTahun + "|" + cCat).toLowerCase()] = statusObj;
            ceklisMap[(cNip + "|" + cTanggal + "|" + cBulan + "|" + cTahun + "|" + cleanCatC).toLowerCase()] = statusObj;
            ceklisMap[(cNip + "|" + cTanggal + "|" + cBulan + "|" + cCat).toLowerCase()] = statusObj;
            ceklisMap[(cNip + "|" + cTanggal + "|" + cBulan + "|" + cleanCatC).toLowerCase()] = statusObj;
            ceklisMap[(cNip + "|" + cTanggal + "|" + cBulan).toLowerCase()] = statusObj;
          }
          if (cNama) {
            ceklisMap[(cNama + "|" + cTanggal + "|" + cBulan + "|" + cTahun + "|" + cCat).toLowerCase()] = statusObj;
            ceklisMap[(cNama + "|" + cTanggal + "|" + cBulan + "|" + cTahun + "|" + cleanCatC).toLowerCase()] = statusObj;
            ceklisMap[(cNama + "|" + cTanggal + "|" + cBulan + "|" + cCat).toLowerCase()] = statusObj;
            ceklisMap[(cNama + "|" + cTanggal + "|" + cBulan + "|" + cleanCatC).toLowerCase()] = statusObj;
            ceklisMap[(cNama + "|" + cTanggal + "|" + cBulan).toLowerCase()] = statusObj;
          }
        }
      }
    }

    var lastColA = Math.max(sheetAcuan.getLastColumn(), 14);
    var rawAcuan = sheetAcuan.getRange(2, 1, sheetAcuan.getLastRow() - 1, lastColA).getDisplayValues();

    var filtered = [];
    var ultg = (ultgFilter || "ALL").toLowerCase();
    var monthNorm = normalizeMonthString(monthFilter);
    var year = (yearFilter || "ALL").toString();
    var q = (searchQuery || "").toLowerCase();
    var cat = (categoryFilter || "ALL").trim().toUpperCase();
    var week = (weekFilter || "ALL").toUpperCase();
    var klwNorm = (klwFilter || "ALL").toString().trim().toLowerCase();
    var giNorm = (giFilter || "ALL").toString().trim().toLowerCase();

    var checkedCount = 0;
    var totalCount = 0;
    var seenKLW = {};
    var seenGI = {};
    var klwList = [];
    var giList = [];

    // Kumpulkan seluruh KLW & GI yang tersedia dari AcuanKendala
    for (var a = 0; a < rawAcuan.length; a++) {
      var aKlw = (rawAcuan[a][10] || "").trim();
      var aGi = (rawAcuan[a][5] || "").trim();
      if (aKlw && aKlw !== "-" && !seenKLW[aKlw.toUpperCase()]) {
        seenKLW[aKlw.toUpperCase()] = true;
        klwList.push(aKlw);
      }
      if (aGi && aGi !== "-" && !seenGI[aGi.toUpperCase()]) {
        seenGI[aGi.toUpperCase()] = true;
        giList.push(aGi);
      }
    }
    klwList.sort(function(x, y) { return x.localeCompare(y); });
    giList.sort(function(x, y) { return x.localeCompare(y); });

    for (var i = 0; i < rawAcuan.length; i++) {
      var jenisBa = rawAcuan[i][0] || "";
      var nama = rawAcuan[i][1] || "";
      var jabatan = rawAcuan[i][2] || "";
      var nip = rawAcuan[i][3] || "";
      var penghantar = rawAcuan[i][4] || "";
      var gi = (rawAcuan[i][5] || "").trim();
      var itemUltg = rawAcuan[i][6] || "";
      var tanggal = rawAcuan[i][7] || "";
      var itemBulanRaw = rawAcuan[i][8] || "";
      var itemBulanNorm = normalizeMonthString(itemBulanRaw);
      var ketKendala = rawAcuan[i][9] || "";
      var klw = (rawAcuan[i][10] || "").trim();
      var tlGi = rawAcuan[i][11] || "";
      var tglSurat = rawAcuan[i][12] || "";

      var yearMatch = tglSurat.match(/\b(20\d\d)\b/);
      var itemYear = yearMatch ? yearMatch[1] : "2026";

      var matchUltg = (ultg === "all" || itemUltg.toLowerCase().indexOf(ultg) !== -1);
      var matchMonth = (monthFilter === "ALL" || monthNorm === "" || itemBulanNorm === monthNorm || itemBulanRaw.toUpperCase().indexOf(monthFilter.toString().toUpperCase()) !== -1);
      var matchYear = (year === "ALL" || itemYear === year || year === "");
      
      var normalizedJenis = jenisBa.toUpperCase();
      var matchCategory = (cat === "ALL" || normalizedJenis.indexOf(cat) !== -1 || cat.indexOf(normalizedJenis) !== -1);
      var matchSearch = (q === "" || (nama + " " + nip + " " + itemUltg + " " + gi + " " + penghantar + " " + klw).toLowerCase().indexOf(q) !== -1);
      var matchKlw = (klwNorm === "all" || klw.toLowerCase().indexOf(klwNorm) !== -1);
      var matchGi = (giNorm === "all" || gi.toLowerCase().indexOf(giNorm) !== -1);

      var matchWeek = true;
      if (week !== "ALL" && tanggal) {
        var tglNum = parseInt(normalizeTanggal(tanggal), 10);
        if (!isNaN(tglNum)) {
          if (week === "M1") matchWeek = (tglNum >= 1 && tglNum <= 7);
          else if (week === "M2") matchWeek = (tglNum >= 8 && tglNum <= 14);
          else if (week === "M3") matchWeek = (tglNum >= 15 && tglNum <= 21);
          else if (week === "M4") matchWeek = (tglNum >= 22 && tglNum <= 28);
          else if (week === "M5") matchWeek = (tglNum >= 29 && tglNum <= 31);
        }
      }

      if (matchUltg && matchMonth && matchYear && matchCategory && matchSearch && matchWeek && matchKlw && matchGi) {
        totalCount++;

        var cleanJenis = normalizedJenis.replace(/^BERITA\s+ACARA\s+/i, "").replace(/^KENDALA\s+/i, "").trim();
        var normTgl = normalizeTanggal(tanggal);
        var namaLow = nama.trim().toLowerCase();
        var validNip = (nip && nip !== "000000" && nip !== "-");

        var cKeyNip = (nip + "|" + normTgl + "|" + itemBulanNorm + "|" + itemYear + "|" + normalizedJenis).toLowerCase();
        var cKeyCleanNip = (nip + "|" + normTgl + "|" + itemBulanNorm + "|" + itemYear + "|" + cleanJenis).toLowerCase();
        var cKeyNipShort = (nip + "|" + normTgl + "|" + itemBulanNorm + "|" + cleanJenis).toLowerCase();
        var cKeyNipDate = (nip + "|" + normTgl + "|" + itemBulanNorm).toLowerCase();

        var cKeyNama = (namaLow + "|" + normTgl + "|" + itemBulanNorm + "|" + itemYear + "|" + normalizedJenis).toLowerCase();
        var cKeyCleanNama = (namaLow + "|" + normTgl + "|" + itemBulanNorm + "|" + itemYear + "|" + cleanJenis).toLowerCase();
        var cKeyNamaShort = (namaLow + "|" + normTgl + "|" + itemBulanNorm + "|" + cleanJenis).toLowerCase();
        var cKeyNamaDate = (namaLow + "|" + normTgl + "|" + itemBulanNorm).toLowerCase();

        // Cari kecocokan status Ceklis & status Berita Acara yang spesifik untuk tanggal ini
        var cStatus = (validNip ? (ceklisMap[cKeyNip] || ceklisMap[cKeyCleanNip] || ceklisMap[cKeyNipShort] || ceklisMap[cKeyNipDate]) : null) ||
                      ceklisMap[cKeyNama] || ceklisMap[cKeyCleanNama] || ceklisMap[cKeyNamaShort] || ceklisMap[cKeyNamaDate] || null;

        var bStatus = (validNip ? (baStatusMap[cKeyNip] || baStatusMap[cKeyCleanNip] || baStatusMap[cKeyNipShort] || baStatusMap[cKeyNipDate]) : null) ||
                      baStatusMap[cKeyNama] || baStatusMap[cKeyCleanNama] || baStatusMap[cKeyNamaShort] || baStatusMap[cKeyNamaDate] || null;

        var rowIsCeklis = false;
        var rowTanggalUpdate = "";
        var rowFileBaUrl = "";
        var rowFileBaId = "";

        if (cStatus && cStatus.isCeklis) {
          rowIsCeklis = true;
          rowTanggalUpdate = cStatus.tanggalUpdate || "";
          rowFileBaUrl = cStatus.fileBaUrl || "";
          rowFileBaId = cStatus.fileBaId || "";
        }

        if (bStatus && bStatus.isCeklis) {
          rowIsCeklis = true;
          if (!rowFileBaUrl && bStatus.fileBaUrl) rowFileBaUrl = bStatus.fileBaUrl;
          if (!rowFileBaId && bStatus.fileBaId) rowFileBaId = bStatus.fileBaId;
          if (!rowTanggalUpdate && bStatus.tanggalUpdate) rowTanggalUpdate = bStatus.tanggalUpdate;
        }

        if (rowIsCeklis) checkedCount++;

        filtered.push({
          rowId: i + 2,
          ultg: itemUltg,
          nama: nama,
          nip: nip,
          tanggal: tanggal,
          bulan: itemBulanRaw.toUpperCase(),
          tahun: itemYear,
          jenisKendala: jenisBa,
          penghantar: penghantar,
          garduInduk: gi,
          klw: klw,
          tlGi: tlGi,
          keteranganKendala: ketKendala,
          tanggalSurat: tglSurat,
          isCeklis: rowIsCeklis,
          tanggalUpdate: rowTanggalUpdate,
          fileBaUrl: rowFileBaUrl,
          fileBaId: rowFileBaId
        });
      }
    }

    return {
      status: "success",
      data: filtered,
      checkedCount: checkedCount,
      totalCount: totalCount,
      klwList: klwList,
      giList: giList
    };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function uploadBatchCeklisBA(batchData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetCeklis = ss.getSheetByName("Ceklis");
    var sheetBA = ss.getSheetByName("BeritaAcara");
    var folder = getOrCreateDriveFolder("UPT_Semarang_Ceklis_BeritaAcara");

    if (!batchData || !batchData.items || batchData.items.length === 0) {
      return { status: "error", message: "Tidak ada item Ceklis yang dipilih!" };
    }

    // Pastikan header Ceklis memiliki kolom Tanggal_Kendala
    if (sheetCeklis) {
      var lastColHeader = Math.max(sheetCeklis.getLastColumn(), 1);
      var currentHeaders = sheetCeklis.getRange(1, 1, 1, lastColHeader).getDisplayValues()[0];
      if (currentHeaders.length < 11 || !currentHeaders[10]) {
        sheetCeklis.getRange(1, 11).setValue("Tanggal_Kendala");
      }
    }

    var nowStr = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");
    var todaySuratStr = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd MMMM yyyy");
    var processedCount = 0;

    for (var b = 0; b < batchData.items.length; b++) {
      var item = batchData.items[b];
      var base64Data = item.fileBase64;
      var fileName = item.fileName || ("BA_Signed_" + (item.nip || item.nama) + ".pdf");

      var fileUrl = "#";
      var fileId = "";

      if (base64Data) {
        var bytes = Utilities.base64Decode(base64Data.split(",")[1] || base64Data);
        var blob = Utilities.newBlob(bytes, "application/pdf", fileName);
        var driveFile = folder.createFile(blob);
        driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fileUrl = driveFile.getUrl();
        fileId = driveFile.getId();
      }

      var targetCat = (item.jenisKendala || "BERITA ACARA KENDALA SOSIALISASI TEGAKAN").trim().toUpperCase();
      var targetBulan = (item.bulan || "OCTOBER").toUpperCase();
      var targetBulanNorm = normalizeMonthString(targetBulan);
      var targetTahun = (item.tahun || "2026").toString();
      var targetTanggal = normalizeTanggal(item.tanggal || "1");
      var targetNama = (item.nama || "").trim().toLowerCase();
      var targetNip = (item.nip || "").trim();

      var lastColCeklis = Math.max(sheetCeklis.getLastColumn(), 11);
      var ceklisData = sheetCeklis.getRange(2, 1, Math.max(sheetCeklis.getLastRow() - 1, 1), lastColCeklis).getDisplayValues();
      var foundCeklisIdx = -1;

      for (var c = 0; c < ceklisData.length; c++) {
        var rNama = (ceklisData[c][1] || "").trim().toLowerCase();
        var rNip = (ceklisData[c][2] || "").trim();
        var rBulan = normalizeMonthString(ceklisData[c][3] || "");
        var rTahun = (ceklisData[c][4] || "").toString();
        var rCat = (ceklisData[c][5] || "").trim().toUpperCase();
        var rTanggal = ceklisData[c].length >= 11 ? normalizeTanggal(ceklisData[c][10]) : "";

        var validNip = (targetNip && targetNip !== "000000" && targetNip !== "-");
        var matchPerson = (validNip && rNip === targetNip) || (rNama === targetNama);
        var matchBulan = (rBulan === targetBulanNorm);
        var matchTahun = (!rTahun || rTahun === targetTahun);
        var matchCat = (!rCat || rCat === targetCat || targetCat.indexOf(rCat) !== -1 || rCat.indexOf(targetCat) !== -1);
        var matchTanggal = (!rTanggal || rTanggal === targetTanggal);

        if (matchPerson && matchBulan && matchTahun && matchCat && matchTanggal) {
          foundCeklisIdx = c + 2;
          break;
        }
      }

      if (foundCeklisIdx !== -1) {
        sheetCeklis.getRange(foundCeklisIdx, 7).setValue("TRUE");
        sheetCeklis.getRange(foundCeklisIdx, 8).setValue(nowStr);
        sheetCeklis.getRange(foundCeklisIdx, 9).setValue(fileUrl);
        sheetCeklis.getRange(foundCeklisIdx, 10).setValue(fileId);
        sheetCeklis.getRange(foundCeklisIdx, 11).setValue(targetTanggal);
      } else {
        sheetCeklis.appendRow([item.ultg, item.nama, item.nip, targetBulan, targetTahun, targetCat, "TRUE", nowStr, fileUrl, fileId, targetTanggal]);
      }

      if (sheetBA) {
        var newBaId = generateSequentialId("BA", "BeritaAcara");
        sheetBA.appendRow([
          newBaId,
          item.jenisKendala || ("BERITA ACARA " + targetCat),
          item.nama || "-",
          item.jabatan || "Line Walker",
          item.nip || "-",
          item.penghantar || "-",
          item.garduInduk || "-",
          item.ultg || "-",
          item.tanggal || "1",
          targetBulan,
          item.keteranganKendala || ("Kendala Pemenuhan Target " + targetCat),
          item.klw || "-",
          item.tlGi || "-",
          item.tanggalSurat || todaySuratStr,
          "Selesai",
          nowStr,
          fileUrl,
          fileId
        ]);
      }

      processedCount++;
    }

    SpreadsheetApp.flush();
    return {
      status: "success",
      message: "Berhasil mengunggah Berita Acara TTD dan menerbitkan " + processedCount + " Ceklis BA!"
    };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function uncheckCeklisWithFileDeletion(nip, bulan, tahun, category, fileBaId, reason, tanggal, nama) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetCeklis = ss.getSheetByName("Ceklis");
    var sheetBA = ss.getSheetByName("BeritaAcara");

    if (fileBaId && fileBaId !== "" && fileBaId !== "#") {
      try {
        var file = DriveApp.getFileById(fileBaId);
        file.setTrashed(true);
      } catch (e) {
        Logger.log("File Drive tidak ditemukan atau gagal dihapus: " + e.toString());
      }
    }

    var targetCat = (category || "BERITA ACARA KENDALA SOSIALISASI TEGAKAN").trim().toUpperCase();
    var targetBulan = (bulan || "OCTOBER").toUpperCase();
    var targetBulanNorm = normalizeMonthString(targetBulan);
    var targetTahun = (tahun || "2026").toString();
    var targetTanggal = normalizeTanggal(tanggal);
    var targetNama = (nama || "").trim().toLowerCase();
    var targetNip = (nip || "").trim();

    if (sheetCeklis && sheetCeklis.getLastRow() > 1) {
      var lastCol = Math.max(sheetCeklis.getLastColumn(), 11);
      var data = sheetCeklis.getRange(2, 1, sheetCeklis.getLastRow() - 1, lastCol).getDisplayValues();
      var nowStr = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");

      for (var i = 0; i < data.length; i++) {
        var rNama = (data[i][1] || "").trim().toLowerCase();
        var rNip = (data[i][2] || "").trim();
        var rBulan = normalizeMonthString(data[i][3] || "");
        var rTahun = data[i][4].toString();
        var rCat = (data[i][5] || "").trim().toUpperCase();
        var rTanggal = data[i].length >= 11 ? normalizeTanggal(data[i][10]) : "";
        var rFileId = data[i][9] || "";

        var matchFile = (fileBaId && fileBaId !== "#" && rFileId === fileBaId);
        var validNip = (targetNip && targetNip !== "000000" && targetNip !== "-");
        var matchPerson = (validNip && rNip === targetNip) || (rNama === targetNama);
        var matchBulan = (rBulan === targetBulanNorm);
        var matchTahun = (!rTahun || rTahun === targetTahun);
        var matchCat = (!rCat || rCat === targetCat || targetCat.indexOf(rCat) !== -1 || rCat.indexOf(targetCat) !== -1);
        var matchTanggal = (!targetTanggal || !rTanggal || rTanggal === targetTanggal);

        if (matchFile || (matchPerson && matchBulan && matchTahun && matchCat && matchTanggal)) {
          var rowIndex = i + 2;
          sheetCeklis.getRange(rowIndex, 7).setValue("FALSE");
          sheetCeklis.getRange(rowIndex, 8).setValue(nowStr);
          sheetCeklis.getRange(rowIndex, 9).setValue("");
          sheetCeklis.getRange(rowIndex, 10).setValue("");
          break;
        }
      }
    }

    if (sheetBA && sheetBA.getLastRow() > 1) {
      var baData = sheetBA.getRange(2, 1, sheetBA.getLastRow() - 1, 18).getDisplayValues();
      for (var b = baData.length - 1; b >= 0; b--) {
        var baFileId = baData[b][17];
        var baNama = (baData[b][2] || "").trim().toLowerCase();
        var baNip = (baData[b][4] || "").trim();
        var baTanggal = normalizeTanggal(baData[b][8]);
        var baBulan = normalizeMonthString(baData[b][9]);

        var matchFile = (fileBaId && fileBaId !== "#" && baFileId === fileBaId);
        var validNipBA = (targetNip && targetNip !== "000000" && targetNip !== "-");
        var matchPerson = (validNipBA && baNip === targetNip) || (baNama === targetNama);
        var matchDate = (baBulan === targetBulanNorm && (!targetTanggal || baTanggal === targetTanggal));

        if (matchFile || (matchPerson && matchDate)) {
          sheetBA.deleteRow(b + 2);
          break;
        }
      }
    }

    SpreadsheetApp.flush();
    return {
      status: "success",
      message: "Ceklis dibatalkan & Berita Acara di Drive berhasil dihapus. Alasan: " + reason
    };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function getWorkOrderList(searchQuery, page, limit) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("WorkOrder");
    if (!sheet || sheet.getLastRow() <= 1) {
      return { data: [], totalPages: 1, currentPage: 1, totalItems: 0 };
    }

    var lastCol = Math.max(sheet.getLastColumn(), 11);
    var rawData = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getDisplayValues();
    var filtered = [];
    var search = (searchQuery || "").toLowerCase();

    for (var i = rawData.length - 1; i >= 0; i--) {
      var idWo = rawData[i][0].toLowerCase();
      var gi = rawData[i][1].toLowerCase();
      var deskripsi = rawData[i][3].toLowerCase();
      var jsonDatesStr = rawData[i][7] || "[]";
      var statusTTD = rawData[i][8] || "Belum TTD";
      var penandatangan = rawData[i][9] || "-";

      var tanggalLemburList = [];
      try { tanggalLemburList = JSON.parse(jsonDatesStr); } catch(e) {}

      var searchBlob = (idWo + " " + gi + " " + deskripsi + " " + penandatangan).toLowerCase();
      if (searchBlob.indexOf(search) !== -1) {
        filtered.push({
          idWO: rawData[i][0],
          namaGI: rawData[i][1],
          bulanOrder: rawData[i][2],
          deskripsi: rawData[i][3],
          filePdfUrl: rawData[i][4],
          filePdfId: rawData[i][5],
          tanggalInput: rawData[i][6],
          tanggalLemburList: tanggalLemburList,
          statusTTD: statusTTD,
          penandatangan: penandatangan
        });
      }
    }

    var p = page || 1;
    var l = limit || 10;
    var totalItems = filtered.length;
    var totalPages = Math.ceil(totalItems / l) || 1;
    var startIndex = (p - 1) * l;
    var paginatedData = filtered.slice(startIndex, startIndex + l);

    return {
      status: "success",
      data: paginatedData,
      totalPages: totalPages,
      currentPage: p,
      totalItems: totalItems
    };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function createWorkOrder(formObj) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("WorkOrder");
    if (!sheet) {
      sheet = ss.insertSheet("WorkOrder");
      sheet.appendRow(["ID_WO", "Nama_GI", "Bulan_Order", "Deskripsi", "File_PDF_Url", "File_PDF_Id", "Tanggal_Input", "JSON_TANGGAL_PERSONEL", "STATUS_TTD", "PENANDATANGAN", "FILE_TTD_URL"]);
    }

    var newId = generateSequentialId("WO", "WorkOrder");
    var fileUrl = "#";
    var fileId = "";

    if (formObj.fileBase64 && formObj.fileName) {
      var folder = getOrCreateDriveFolder("UPT_Semarang_WorkOrders_PDF");
      var bytes = Utilities.base64Decode(formObj.fileBase64.split(",")[1] || formObj.fileBase64);
      var blob = Utilities.newBlob(bytes, "application/pdf", newId + "_" + formObj.fileName);
      var driveFile = folder.createFile(blob);
      driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = driveFile.getUrl();
      fileId = driveFile.getId();
    }

    var nowStr = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");
    var jsonDatesStr = JSON.stringify(formObj.tanggalLemburList || []);

    var rowData = [
      newId,
      formObj.namaGI || "-",
      formObj.bulanOrder || "Oktober 2026",
      formObj.deskripsi || "-",
      fileUrl,
      fileId,
      nowStr,
      jsonDatesStr,
      "Belum TTD",
      "-",
      ""
    ];

    sheet.appendRow(rowData);
    SpreadsheetApp.flush();

    return { status: "success", message: "SPKL " + newId + " berhasil dibuat!", idWO: newId };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function uploadWorkOrderPdf(idWO, fileBase64, fileName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("WorkOrder");
    if (!sheet || sheet.getLastRow() <= 1) {
      return { status: "error", message: "Sheet WorkOrder tidak ditemukan!" };
    }

    var folder = getOrCreateDriveFolder("UPT_Semarang_WorkOrders_PDF");
    var bytes = Utilities.base64Decode(fileBase64.split(",")[1] || fileBase64);
    var blob = Utilities.newBlob(bytes, "application/pdf", idWO + "_" + fileName);
    var driveFile = folder.createFile(blob);
    driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileUrl = driveFile.getUrl();
    var fileId = driveFile.getId();

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getDisplayValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === idWO) {
        var rowIndex = i + 2;
        sheet.getRange(rowIndex, 5).setValue(fileUrl);
        sheet.getRange(rowIndex, 6).setValue(fileId);
        break;
      }
    }

    SpreadsheetApp.flush();
    return { status: "success", message: "Berkas PDF SPKL " + idWO + " berhasil diunggah!", fileUrl: fileUrl };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function signWorkOrder(idWO, sigBase64, namaPenandatangan) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("WorkOrder");
    if (!sheet || sheet.getLastRow() <= 1) {
      return { status: "error", message: "Sheet WorkOrder tidak ditemukan!" };
    }

    if (typeof idWO !== "string" || !idWO.trim() ||
        typeof sigBase64 !== "string" || !sigBase64.trim() ||
        typeof namaPenandatangan !== "string" || !namaPenandatangan.trim()) {
      throw new Error("ID WO, tanda tangan, dan nama penandatangan wajib diisi.");
    }
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getDisplayValues();
    var rowIndex = -1;
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === idWO) { rowIndex = i + 2; break; }
    }
    if (rowIndex === -1) throw new Error("WorkOrder " + idWO + " tidak ditemukan.");

    var folder = getOrCreateDriveFolder("UPT_Semarang_TTD_WO");
    var bytes = Utilities.base64Decode(sigBase64.split(",")[1] || sigBase64);
    var blob = Utilities.newBlob(bytes, "image/png", "TTD_" + idWO + "_" + namaPenandatangan + ".png");
    var driveFile = folder.createFile(blob);
    driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var ttdUrl = driveFile.getUrl();

    sheet.getRange(rowIndex, 9, 1, 3).setValues([["Sudah TTD", namaPenandatangan, ttdUrl]]);

    SpreadsheetApp.flush();
    return { status: "success", message: "Tanda Tangan Digital untuk " + idWO + " berhasil disimpan!" };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function getAbsensiCalendarData(monthFilter, yearFilter) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Absensi");
    if (!sheet || sheet.getLastRow() <= 1) {
      return { status: "success", data: {} };
    }

    var rawData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getDisplayValues();
    var m = (monthFilter || "OCTOBER").toUpperCase();
    var mNorm = normalizeMonthString(m);
    var y = (yearFilter || "2026").toString();

    var calendarMap = {};

    for (var i = 0; i < rawData.length; i++) {
      var namaLw = rawData[i][1];
      var nip = rawData[i][2];
      var tgl = ("0" + rawData[i][3]).slice(-2);
      var bulan = (rawData[i][4] || "").toUpperCase();
      var bulanNorm = normalizeMonthString(bulan);
      var tahun = (rawData[i][5] || "").toString();
      var jamMasuk = rawData[i][6] || "-";
      var jamKeluar = rawData[i][7] || "-";
      var stTerlambat = rawData[i][8] || "";
      var stPulang = rawData[i][9] || "";

      if ((bulan === m || bulanNorm === mNorm) && tahun === y) {
        // Logic 1: Tidak Absen Checkout (Nilai = 0)
        var isTidakCheckout = (jamKeluar === "-" || !jamKeluar || jamKeluar.trim() === "" || jamKeluar === jamMasuk || stPulang.toLowerCase().indexOf("belum") !== -1 || stPulang.toLowerCase().indexOf("tidak") !== -1);

        // Logic 2: Terlambat Presensi (Nilai < 1, contoh: 0.75)
        var isTerlambat = (stTerlambat.toLowerCase().indexOf("terlambat") !== -1);
        var nilaiMasuk = 1;

        if (isTidakCheckout) {
          nilaiMasuk = 0;
        } else if (isTerlambat) {
          nilaiMasuk = 0.75;
        }

        if (!isTerlambat && jamMasuk !== "-" && jamMasuk !== "") {
          var timeParts = jamMasuk.split(":");
          if (timeParts.length >= 2) {
            var hour = parseInt(timeParts[0], 10);
            var min = parseInt(timeParts[1], 10);
            if (hour > 8 || (hour === 8 && min > 15) || hour > 9 || (hour === 9 && min > 45)) {
              isTerlambat = true;
              nilaiMasuk = 0.75;
              if (!stTerlambat) stTerlambat = "Terlambat (Checkin " + jamMasuk + ")";
            }
          }
        }

        // HANYA proses jika tergolong Tidak Checkout ATAU Terlambat
        if (isTidakCheckout || isTerlambat) {
          if (!calendarMap[tgl]) {
            calendarMap[tgl] = {
              date: tgl,
              terlambat: [],
              tidakCheckout: []
            };
          }

          if (isTidakCheckout) {
            var existsTC = calendarMap[tgl].tidakCheckout.some(function(tc) {
              return tc.nama === namaLw || (tc.nip && tc.nip === nip && nip !== "000000");
            });
            if (!existsTC) {
              calendarMap[tgl].tidakCheckout.push({
                nama: namaLw,
                nip: nip,
                jamMasuk: jamMasuk,
                nilai: 0
              });
            }
          } else if (isTerlambat) {
            var existsTB = calendarMap[tgl].terlambat.some(function(tb) {
              return tb.nama === namaLw || (tb.nip && tb.nip === nip && nip !== "000000");
            });
            if (!existsTB) {
              calendarMap[tgl].terlambat.push({
                nama: namaLw,
                nip: nip,
                jamMasuk: jamMasuk,
                keterangan: stTerlambat || "Terlambat Presensi",
                nilai: 0.75
              });
            }
          }
        }
      }
    }

    var sheetAcuan = ss.getSheetByName("AcuanKendala");
    var syncedCeklisKeys = {};
    if (sheetAcuan && sheetAcuan.getLastRow() > 1) {
      var lastColAcuan = Math.max(sheetAcuan.getLastColumn(), 10);
      var rawAcuan = sheetAcuan.getRange(2, 1, sheetAcuan.getLastRow() - 1, lastColAcuan).getDisplayValues();
      for (var a = 0; a < rawAcuan.length; a++) {
        var aJenis = (rawAcuan[a][0] || "").trim().toLowerCase();
        var aKet = (rawAcuan[a][9] || "").trim().toLowerCase();
        if (aJenis.indexOf("absensi") !== -1 || aJenis.indexOf("terlambat") !== -1 || aJenis.indexOf("checkout") !== -1 || aKet.indexOf("tidak checkout") !== -1 || aKet.indexOf("terlambat") !== -1) {
          var aNama = (rawAcuan[a][1] || "").trim().toLowerCase();
          var aTgl = parseInt(rawAcuan[a][7] || "0", 10);
          var aBln = normalizeMonthString(rawAcuan[a][8] || "");
          if (aNama && aTgl && aBln) {
            syncedCeklisKeys[aNama + "|" + aTgl + "|" + aBln] = true;
          }
        }
      }
    }

    return { status: "success", data: calendarMap, syncedCeklisKeys: syncedCeklisKeys };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function addPersonelAbsensi(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Absensi");
    if (!sheet) {
      sheet = ss.insertSheet("Absensi");
      sheet.appendRow(["ID_ABSEN", "NAMA_LW", "NIP", "TANGGAL", "BULAN", "TAHUN", "JAM_MASUK", "JAM_KELUAR", "STATUS_KETERLAMBATAN", "STATUS_PULANG"]);
    }

    var newId = generateSequentialId("ABS", "Absensi");
    var rowData = [
      newId,
      payload.nama,
      payload.nip,
      payload.tanggal,
      (payload.bulan || "OCTOBER").toUpperCase(),
      payload.tahun || "2026",
      payload.jamMasuk || "08:00",
      payload.jamKeluar || "-",
      payload.status.indexOf("Terlambat") !== -1 ? "Terlambat" : "Tepat Waktu",
      payload.status
    ];

    sheet.appendRow(rowData);
    SpreadsheetApp.flush();

    return { status: "success", message: "Personel " + payload.nama + " berhasil ditambahkan ke Absensi tanggal " + payload.tanggal + "!" };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function importBatchAbsensiData(records) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Absensi");
    if (!sheet) {
      sheet = ss.insertSheet("Absensi");
      sheet.appendRow(["ID_ABSEN", "NAMA_LW", "NIP", "TANGGAL", "BULAN", "TAHUN", "JAM_MASUK", "JAM_KELUAR", "STATUS_KETERLAMBATAN", "STATUS_PULANG"]);
    }

    if (!records || !Array.isArray(records) || records.length === 0) {
      return { status: "error", message: "Data impor absensi kosong atau format tidak valid!" };
    }

    var rowsToAdd = [];
    var todayStr = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyyMMdd");
    var lastRow = sheet.getLastRow();
    var startNum = lastRow > 1 ? lastRow : 1;

    for (var i = 0; i < records.length; i++) {
      var item = records[i];
      var newId = "ABS-" + todayStr + "-" + ("000" + (startNum + i)).slice(-4);
      
      var jamMasuk = item.jamMasuk || item.checkin || "-";
      var jamKeluar = item.jamKeluar || item.checkout || "-";
      
      var stTerlambat = item.statusKeterlambatan || "Tepat Waktu";
      var stPulang = item.statusPulang || (jamKeluar === "-" || !jamKeluar || jamKeluar === jamMasuk ? "Tidak Absen Checkout" : "Sudah Checkout");

      if (jamMasuk !== "-" && jamMasuk !== "") {
        var parts = jamMasuk.split(":");
        if (parts.length >= 2) {
          var h = parseInt(parts[0], 10);
          var m = parseInt(parts[1], 10);
          var jadwalMulai = item.jadwalMulai || "09:45";
          var jParts = jadwalMulai.split(":");
          var jh = jParts[0] ? parseInt(jParts[0], 10) : 9;
          var jm = jParts[1] ? parseInt(jParts[1], 10) : 45;

          if (h > jh || (h === jh && m > jm)) {
            stTerlambat = "Terlambat";
          }
        }
      }

      rowsToAdd.push([
        newId,
        item.nama || "-",
        item.nip || "-",
        ("0" + (item.tanggal || "01")).slice(-2),
        (item.bulan || "OCTOBER").toUpperCase(),
        (item.tahun || "2026").toString(),
        jamMasuk,
        jamKeluar,
        stTerlambat,
        stPulang
      ]);
    }

    if (rowsToAdd.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, 10).setValues(rowsToAdd);
      SpreadsheetApp.flush();
    }

    return { 
      status: "success", 
      message: "Berhasil mengimpor " + rowsToAdd.length + " data presensi harian dari Excel ke Sheet Absensi!" 
    };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function extractDriveFileId(str) {
  if (!str || str === "#") return "";
  var s = str.toString().trim();
  // Jika sudah berupa ID file 20+ karakter (huruf/angka/tanda minus/underscore)
  if (/^[a-zA-Z0-9_-]{20,}$/.test(s)) return s;
  // Format /d/{ID}
  var match = s.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (match && match[1]) return match[1];
  // Format ?id={ID} atau ?key={ID}
  match = s.match(/[?&](?:id|key)=([a-zA-Z0-9_-]+)/);
  if (match && match[1]) return match[1];
  return "";
}

function getLemburanList(searchQuery, monthFilter, yearFilter) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Lemburan");
    if (!sheet || sheet.getLastRow() <= 1) {
      return { status: "success", data: [] };
    }

    var rawData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getDisplayValues();
    var filtered = [];
    var q = (searchQuery || "").toLowerCase();
    var m = (monthFilter || "ALL").toUpperCase();
    var mNorm = normalizeMonthString(m);
    var y = (yearFilter || "ALL").toString();

    for (var i = rawData.length - 1; i >= 0; i--) {
      var idLembur = rawData[i][0];
      var bulanOrder = (rawData[i][1] || "").toUpperCase();
      var bulanOrderNorm = normalizeMonthString(bulanOrder);
      var tahun = (rawData[i][2] || "").toString();
      var namaFile = rawData[i][3];
      var fileExcelUrl = (rawData[i][4] || "").trim();
      var fileExcelId = (rawData[i][5] || "").trim();
      var tglUpload = rawData[i][6];

      var matchMonth = (m === "ALL" || bulanOrder.indexOf(m) !== -1 || bulanOrderNorm === mNorm);
      var matchYear = (y === "ALL" || tahun === y);
      var matchSearch = (q === "" || (idLembur + " " + namaFile + " " + bulanOrder).toLowerCase().indexOf(q) !== -1);

      if (matchMonth && matchYear && matchSearch) {
        var driveId = extractDriveFileId(fileExcelId) || extractDriveFileId(fileExcelUrl);
        var previewUrl = "";
        var directViewUrl = "";
        var downloadUrl = "";
        var hasValidFile = false;

        if (driveId) {
          hasValidFile = true;
          previewUrl = "https://drive.google.com/file/d/" + driveId + "/preview";
          directViewUrl = "https://drive.google.com/file/d/" + driveId + "/view?usp=sharing";
          downloadUrl = "https://drive.google.com/uc?export=download&id=" + driveId;

          if (fileExcelUrl.indexOf("spreadsheets") !== -1) {
            previewUrl = "https://docs.google.com/spreadsheets/d/" + driveId + "/preview";
            directViewUrl = "https://docs.google.com/spreadsheets/d/" + driveId + "/edit?usp=sharing";
          }
        } else if (fileExcelUrl && fileExcelUrl !== "#" && fileExcelUrl.indexOf("http") === 0) {
          hasValidFile = true;
          previewUrl = fileExcelUrl;
          directViewUrl = fileExcelUrl;
          downloadUrl = fileExcelUrl;
        }

        filtered.push({
          idLembur: idLembur,
          bulanOrder: bulanOrder,
          tahun: tahun,
          namaFile: namaFile,
          fileExcelUrl: directViewUrl || fileExcelUrl,
          previewUrl: previewUrl,
          directViewUrl: directViewUrl || previewUrl,
          downloadUrl: downloadUrl || fileExcelUrl,
          hasValidFile: hasValidFile,
          tanggalUpload: tglUpload
        });
      }
    }

    return { status: "success", data: filtered };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function saveLemburan(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Lemburan");
    if (!sheet) {
      sheet = ss.insertSheet("Lemburan");
      sheet.appendRow(["ID Lembur", "Bulan", "Tahun", "Nama File", "File URL", "File ID", "Tanggal Upload"]);
    }

    var now = new Date();
    var nowStr = Utilities.formatDate(now, "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");
    var idLembur = generateSequentialId("LMB", "Lemburan");

    var fileUrl = payload.fileUrl || "";
    var fileId = extractDriveFileId(fileUrl) || "";

    if (payload.fileBase64 && payload.fileName) {
      try {
        var folder = getOrCreateSharedFolder("Lemburan_Uploads");
        var parts = payload.fileBase64.split(",");
        var pureBase64 = parts.length > 1 ? parts[1] : parts[0];
        var decoded = Utilities.base64Decode(pureBase64);
        var blob = Utilities.newBlob(decoded, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", payload.fileName);
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fileId = file.getId();
        fileUrl = file.getUrl();
      } catch (fErr) {
        // fallback jika upload drive gagal
      }
    }

    sheet.appendRow([
      idLembur,
      (payload.bulanOrder || "JANUARI").toUpperCase(),
      payload.tahun || now.getFullYear().toString(),
      payload.namaFile || ("LEMBUR " + payload.bulanOrder + " " + (payload.tahun || "2026")),
      fileUrl,
      fileId,
      nowStr
    ]);

    return { status: "success", message: "Berkas Lemburan " + idLembur + " berhasil disimpan!", idLembur: idLembur };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function getRequestJadwalList(searchQuery, monthFilter, yearFilter) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("RequestJadwal");
    if (!sheet || sheet.getLastRow() <= 1) {
      return { status: "success", data: [] };
    }

    var rawData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getDisplayValues();
    var filtered = [];
    var q = (searchQuery || "").toLowerCase();
    var m = (monthFilter || "ALL").toUpperCase();
    var mNorm = normalizeMonthString(m);

    for (var i = rawData.length - 1; i >= 0; i--) {
      var idReq = rawData[i][0];
      var namaLw = rawData[i][1];
      var nip = rawData[i][2];
      var bulanTarget = (rawData[i][3] || "").toUpperCase();
      var bulanTargetNorm = normalizeMonthString(bulanTarget);
      var detailPerubahan = rawData[i][4];
      var status = rawData[i][5];
      var tglSubmit = rawData[i][6];

      var matchMonth = (m === "ALL" || bulanTarget.indexOf(m) !== -1 || bulanTargetNorm === mNorm);
      var matchSearch = (q === "" || (idReq + " " + namaLw + " " + nip + " " + detailPerubahan).toLowerCase().indexOf(q) !== -1);

      if (matchMonth && matchSearch) {
        filtered.push({
          idRequest: idReq,
          nama: namaLw,
          nip: nip,
          bulanTarget: bulanTarget,
          detailPerubahan: detailPerubahan,
          status: status,
          tanggalSubmit: tglSubmit
        });
      }
    }

    return { status: "success", data: filtered };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function submitRequestJadwal(formObj) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("RequestJadwal");
    if (!sheet) {
      sheet = ss.insertSheet("RequestJadwal");
      sheet.appendRow(["ID_REQUEST", "NAMA_LW", "NIP", "BULAN_TARGET", "DETAIL_PERUBAHAN", "STATUS", "TANGGAL_SUBMIT"]);
    }

    var newId = generateSequentialId("REQ", "RequestJadwal");
    var nowStr = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");

    var rowData = [
      newId,
      formObj.namaLW || "-",
      formObj.nip || "-",
      (formObj.bulanTarget || "NOVEMBER 2026").toUpperCase(),
      formObj.detailPerubahan || "-",
      "Pending",
      nowStr
    ];

    sheet.appendRow(rowData);
    SpreadsheetApp.flush();

    return { status: "success", message: "Permohonan perubahan jadwal " + newId + " berhasil diajukan!" };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function getLineWalkerOptions() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Data LW");
    if (!sheet || sheet.getLastRow() <= 1) {
      return { status: "success", data: [] };
    }

    // Pastikan minimal 9 kolom agar membaca kolom FOTO_URL
    var lastCol = Math.max(sheet.getLastColumn(), 9);
    var rawData = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getDisplayValues();
    var result = [];
    for (var i = 0; i < rawData.length; i++) {
      var nama = (rawData[i][0] || "").trim();
      if (nama !== "") {
        result.push({
          rowIndex: i + 2,
          idLW: rawData[i][2] || ("LW-" + (i + 1)),
          nama: nama,
          koordinator: rawData[i][1] || "-",
          nip: rawData[i][2] || "-",
          posisi: rawData[i][3] || "Line Walker",
          ultg: rawData[i][4] || "-",
          gi: rawData[i][5] || "-",
          tlGi: rawData[i][6] || "-",
          penghantar: rawData[i][7] || "-",
          fotoUrl: rawData[i][8] || ""
        });
      }
    }
    return { status: "success", data: result };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function updateLineWalker(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Data LW");
    if (!sheet || sheet.getLastRow() <= 1) {
      return { status: "error", message: "Sheet Data LW tidak ditemukan!" };
    }

    var lastCol = Math.max(sheet.getLastColumn(), 9);
    var rawData = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getDisplayValues();
    var foundRow = -1;

    // 1. Cari berdasarkan rowIndex jika valid
    if (payload.rowIndex && payload.rowIndex >= 2 && payload.rowIndex <= sheet.getLastRow()) {
      var checkRowNama = (sheet.getRange(payload.rowIndex, 1).getValue() || "").toString().trim().toLowerCase();
      var checkOldNama = (payload.oldNama || payload.nama || "").trim().toLowerCase();
      if (checkRowNama === checkOldNama || checkRowNama !== "") {
        foundRow = payload.rowIndex;
      }
    }

    // 2. Jika belum ketemu, cari pencocokan berdasarkan NIP lama atau Nama lama
    if (foundRow === -1) {
      var targetNip = (payload.oldNip || payload.nip || "").trim().toLowerCase();
      var targetNama = (payload.oldNama || payload.nama || "").trim().toLowerCase();
      for (var i = 0; i < rawData.length; i++) {
        var rNama = (rawData[i][0] || "").trim().toLowerCase();
        var rNip = (rawData[i][2] || "").trim().toLowerCase();
        if ((targetNip !== "" && rNip === targetNip) || (targetNama !== "" && rNama === targetNama)) {
          foundRow = i + 2;
          break;
        }
      }
    }

    if (foundRow === -1) {
      return { status: "error", message: "Data Line Walker tidak ditemukan di database sheet Data LW!" };
    }

    // Kolom: [LINE WALKER, KLW, NIK BARU, POSISI, ULTG, GI, TL GI, Pengahantar, FOTO_URL]
    sheet.getRange(foundRow, 1).setValue(payload.nama || "");
    sheet.getRange(foundRow, 2).setValue(payload.koordinator || "-");
    sheet.getRange(foundRow, 3).setValue(payload.nip || "-");
    sheet.getRange(foundRow, 4).setValue(payload.posisi || "Line Walker");
    sheet.getRange(foundRow, 5).setValue(payload.ultg || "-");
    sheet.getRange(foundRow, 6).setValue(payload.gi || "-");
    sheet.getRange(foundRow, 7).setValue(payload.tlGi || "-");
    sheet.getRange(foundRow, 8).setValue(payload.penghantar || "-");
    if (payload.fotoUrl !== undefined) {
      sheet.getRange(foundRow, 9).setValue(payload.fotoUrl || "");
    }

    // SINKRONISASI BERANTAI TL GI UNTUK SELURUH PERSONEL DI GI YANG SAMA
    var targetGi = (payload.gi || "").toString().trim().toLowerCase();
    var newTlGi = (payload.tlGi || "").toString().trim();
    var syncedGiCount = 0;

    if (targetGi !== "" && newTlGi !== "") {
      for (var k = 0; k < rawData.length; k++) {
        var rowNum = k + 2;
        if (rowNum !== foundRow) {
          var rGi = (rawData[k][5] || "").toString().trim().toLowerCase();
          if (rGi === targetGi) {
            sheet.getRange(rowNum, 7).setValue(newTlGi);
            syncedGiCount++;
          }
        }
      }
    }

    SpreadsheetApp.flush();

    var successMsg = "Data Line Walker " + (payload.nama || "") + " berhasil diperbarui!";
    if (syncedGiCount > 0) {
      successMsg += " (" + syncedGiCount + " personel lain di " + payload.gi + " otomatis disinkronkan ke TL GI: " + newTlGi + ")";
    }

    return {
      status: "success",
      message: successMsg,
      syncedGiCount: syncedGiCount,
      targetGi: payload.gi,
      newTlGi: newTlGi,
      updatedData: {
        rowIndex: foundRow,
        idLW: payload.nip,
        nama: payload.nama,
        koordinator: payload.koordinator,
        nip: payload.nip,
        posisi: payload.posisi,
        ultg: payload.ultg,
        gi: payload.gi,
        tlGi: payload.tlGi,
        penghantar: payload.penghantar,
        fotoUrl: payload.fotoUrl || ""
      }
    };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

/**
 * Upload Foto Line Walker ke Google Drive dan simpan URL ke Sheet Data LW
 */
function uploadFotoLineWalker(rowIndex, nip, nama, fileBase64, mimeType) {
  try {
    if (!fileBase64) {
      return { status: "error", message: "File foto tidak valid atau kosong!" };
    }

    var folder = getOrCreateDriveFolder("UPT_Semarang_Foto_LineWalker");
    var cleanBase64 = fileBase64;
    var detectedMime = mimeType || "image/jpeg";

    if (fileBase64.indexOf(",") !== -1) {
      var parts = fileBase64.split(",");
      cleanBase64 = parts[1];
      var match = parts[0].match(/:(.*?);/);
      if (match && match[1]) detectedMime = match[1];
    }

    var bytes = Utilities.base64Decode(cleanBase64);
    var extension = detectedMime.indexOf("png") !== -1 ? ".png" : ".jpg";
    var safeName = (nama || "LW").replace(/[^a-zA-Z0-9]/g, "_");
    var safeNip = (nip || "NONIP").replace(/[^a-zA-Z0-9]/g, "_");
    var fileName = "FOTO_" + safeNip + "_" + safeName + "_" + Date.now() + extension;

    var blob = Utilities.newBlob(bytes, detectedMime, fileName);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Google Drive direct image URL
    var fileId = file.getId();
    var fotoUrl = "https://lh3.googleusercontent.com/d/" + fileId;

    // Simpan ke Google Sheet Data LW jika rowIndex tersedia
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Data LW");
    var actualRow = rowIndex;

    if (sheet) {
      // Pastikan header kolom 9 ada
      var col9Header = sheet.getRange(1, 9).getValue();
      if (!col9Header || col9Header.toString().trim() === "") {
        sheet.getRange(1, 9).setValue("FOTO_URL");
      }

      if (!actualRow || actualRow < 2) {
        // Cari baris berdasarkan NIP / Nama
        var lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          var nips = sheet.getRange(2, 3, lastRow - 1, 1).getDisplayValues();
          for (var i = 0; i < nips.length; i++) {
            if (nips[i][0] && nips[i][0].toString().trim() === (nip || "").toString().trim()) {
              actualRow = i + 2;
              break;
            }
          }
        }
      }

      if (actualRow && actualRow >= 2) {
        sheet.getRange(actualRow, 9).setValue(fotoUrl);
        SpreadsheetApp.flush();
      }
    }

    return {
      status: "success",
      message: "Foto profil berhasil diunggah!",
      fileId: fileId,
      fotoUrl: fotoUrl,
      rowIndex: actualRow
    };
  } catch (err) {
    return { status: "error", message: err.toString() };
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

function getSuratPeringatanSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return null;

  var all = ss.getSheets();
  var best = null;
  var bestScore = -1;

  for (var i = 0; i < all.length; i++) {
    var sh = all[i];
    var nama = sh.getName();
    var n = nama.trim().toLowerCase().replace(/[\s_\-\&\.\/]+/g, "");

    // Skor nama tab
    var nameScore = 0;
    if (n === "suratperingatan" || n === "datasp" || n === "suratteguran") nameScore = 3;
    else if (n === "sp" || n === "teguran" || n === "peringatan" || n === "disiplin") nameScore = 2;
    else if (n.indexOf("suratperingatan") !== -1 || n.indexOf("peringatan") !== -1) nameScore = 2;
    else if (n.indexOf("teguran") !== -1 && n.indexOf("surat") !== -1) nameScore = 1;

    // Skor header kolom khas SP
    var headerScore = isSpHeaderSheet(sh) ? 4 : 0;

    if (nameScore === 0 && headerScore === 0) continue;

    // Skor isi data (tab berisi data selalu menang atas tab kosong)
    var hasData = false;
    try { hasData = sh.getLastRow() > 1; } catch (e) { hasData = false; }
    var dataScore = hasData ? 6 : 0;

    var total = nameScore + headerScore + dataScore;
    if (total > bestScore) {
      bestScore = total;
      best = sh;
    }
  }

  return best;
}

/**
 * Cek apakah sheet memiliki sidik header kolom Surat Peringatan.
 * Minimal 3 kecocokan dari kolom khas SP.
 */
function isSpHeaderSheet(sheet) {
  try {
    var cols = Math.min(sheet.getLastColumn(), 30);
    if (cols < 3) return false;
    var head = sheet.getRange(1, 1, 1, cols).getDisplayValues()[0];
    var keys = ["IDSP", "NAMALW", "NIP", "JENISSURAT", "NOMORSURAT", "TANGGALSURAT", "TANGGALBERAKHIR", "MASABERLAKU"];
    var hit = 0;
    for (var c = 0; c < head.length; c++) {
      var h = String(head[c] || "").trim().toUpperCase().replace(/[\s_\-\/\.]+/g, "");
      if (keys.indexOf(h) !== -1) hit++;
    }
    return hit >= 3;
  } catch (e) {
    return false;
  }
}

/**
 * DIAGNOSTIK: Jalankan manual di Apps Script Editor untuk melihat isi spreadsheet.
 * Hasilnya di menu Execution log / Logger.
 */
function diagnoseSuratPeringatanSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out = {
    spreadsheet: ss ? ss.getName() : "TIDAK ADA SPREADSHEET AKTIF",
    spreadsheetId: ss ? ss.getId() : "-",
    sheets: [],
    terpilih: "-",
    headerTerpilih: [],
    jumlahBarisData: 0,
    contohBaris: []
  };

  if (!ss) {
    Logger.log(JSON.stringify(out, null, 2));
    return out;
  }

  ss.getSheets().forEach(function (sh) {
    var firstRow = [];
    try {
      if (sh.getLastColumn() > 0 && sh.getLastRow() > 0) {
        firstRow = sh.getRange(1, 1, 1, Math.min(sh.getLastColumn(), 12)).getDisplayValues()[0];
      }
    } catch (e) { }
    out.sheets.push({
      nama: sh.getName(),
      baris: sh.getLastRow(),
      kolom: sh.getLastColumn(),
      header: firstRow
    });
  });

  var target = getSuratPeringatanSheet(ss);
  if (target) {
    out.terpilih = target.getName();
    out.jumlahBarisData = Math.max(target.getLastRow() - 1, 0);
    out.headerTerpilih = target.getRange(1, 1, 1, Math.min(target.getLastColumn(), 15)).getDisplayValues()[0];
    if (target.getLastRow() > 1) {
      out.contohBaris = target.getRange(2, 1, Math.min(3, target.getLastRow() - 1), Math.min(target.getLastColumn(), 15)).getDisplayValues();
    }
  }

  Logger.log(JSON.stringify(out, null, 2));
  return out;
}

function parseDateFlexible(dateStr) {
  if (!dateStr) return new Date();
  if (dateStr instanceof Date) return dateStr;
  var str = dateStr.toString().trim();
  var cleanStr = str.split(" ")[0];
  // Format pemisah tanda hubung (-)
  if (cleanStr.indexOf("-") !== -1) {
    var parts = cleanStr.split("-");
    if (parts.length === 3) {
      if (parts[0].length === 4) return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
      if (parts[2].length === 4) return new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
    }
  }
  // Format pemisah garis miring (/)
  if (cleanStr.indexOf("/") !== -1) {
    var parts2 = cleanStr.split("/");
    if (parts2.length === 3) {
      if (parts2[2].length === 4) return new Date(parseInt(parts2[2], 10), parseInt(parts2[1], 10) - 1, parseInt(parts2[0], 10));
      if (parts2[0].length === 4) return new Date(parseInt(parts2[0], 10), parseInt(parts2[1], 10) - 1, parseInt(parts2[2], 10));
    }
  }
  var d = new Date(str);
  return isNaN(d.getTime()) ? new Date() : d;
}

function getSuratPeringatanList(searchQuery, jenisFilter, statusFilter, page, limit) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getSuratPeringatanSheet(ss);
    if (!sheet) {
      sheet = ss.insertSheet("SuratPeringatan");
      sheet.appendRow(["ID_SP", "NAMA_LW", "NIP", "JENIS_SURAT", "NOMOR_SURAT", "TANGGAL_SURAT", "MASA_BERLAKU_BULAN", "TANGGAL_BERAKHIR", "KETERANGAN", "FILE_PDF_URL", "FILE_PDF_ID", "STATUS_AKTIF", "TANGGAL_UPLOAD"]);
      if (typeof getDummySuratPeringatanData === "function") {
        var dummies = getDummySuratPeringatanData();
        for (var d = 0; d < dummies.length; d++) {
          sheet.appendRow(dummies[d]);
        }
      }
      SpreadsheetApp.flush();
    }

    if (sheet.getLastRow() <= 1) {
      return {
        status: "success",
        data: [],
        total: 0,
        totalPages: 1,
        debug: {
          sheetUsed: sheet.getName(),
          rowCount: 0,
          headers: sheet.getLastColumn() > 0 ? sheet.getRange(1, 1, 1, Math.min(sheet.getLastColumn(), 15)).getDisplayValues()[0] : [],
          allSheets: ss.getSheets().map(function (s) { return s.getName() + " (" + s.getLastRow() + "x" + s.getLastColumn() + ")"; })
        }
      };
    }

    var lastCol = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    if (lastCol < 1 || lastRow <= 1) {
      return { status: "success", data: [], total: 0, totalPages: 1 };
    }

    // Mapping header kolom secara dinamis & toleran terhadap format spreadsheet pengguna
    var headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
    var colMap = {
      idSp: -1,
      namaLw: -1,
      nip: -1,
      jenisSurat: -1,
      nomorSurat: -1,
      tglSurat: -1,
      masaBulan: -1,
      tglBerakhir: -1,
      keterangan: -1,
      filePdfUrl: -1,
      filePdfId: -1,
      statusAktif: -1,
      klw: -1,
      tglUpload: -1
    };

    for (var c = 0; c < headers.length; c++) {
      var h = String(headers[c] || "").trim().toUpperCase().replace(/[\s_\-\/\.]+/g, "");
      if (colMap.idSp === -1 && (h === "IDSP" || h === "ID" || h === "NOID" || h === "KODESURAT" || h === "IDPERINGATAN")) colMap.idSp = c;
      if (colMap.namaLw === -1 && (h === "NAMALW" || h === "NAMA" || h === "LINEWALKER" || h === "NAMALINEWALKER" || h === "PERSONEL" || h === "PETUGAS")) colMap.namaLw = c;
      if (colMap.nip === -1 && (h === "NIP" || h === "NIK" || h === "NIPNIK" || h === "NIKBARU")) colMap.nip = c;
      if (colMap.jenisSurat === -1 && (h === "JENISSURAT" || h === "JENIS" || h === "STATUSSP" || h === "SURAT" || h === "JENISPERINGATAN" || h === "SANKSI")) colMap.jenisSurat = c;
      if (colMap.nomorSurat === -1 && (h === "NOMORSURAT" || h === "NOSURAT" || h === "NOMOR" || h === "NO" || h === "NOBERKAS")) colMap.nomorSurat = c;
      if (colMap.tglSurat === -1 && (h === "TANGGALSURAT" || h === "TANGGAL" || h === "TGLSURAT" || h === "TGL" || h === "TANGGALTERBIT" || h === "TANGGALSP")) colMap.tglSurat = c;
      if (colMap.masaBulan === -1 && (h === "MASABERLAKUBULAN" || h === "MASABERLAKU" || h === "MASA" || h === "DURASIBULAN" || h === "BULAN")) colMap.masaBulan = c;
      if (colMap.tglBerakhir === -1 && (h === "TANGGALBERAKHIR" || h === "TGLBERAKHIR" || h === "BERAKHIR" || h === "BATASAKHIR" || h === "EXPIRED" || h === "TGLEXPIRED")) colMap.tglBerakhir = c;
      if (colMap.keterangan === -1 && (h === "KETERANGAN" || h === "ALASAN" || h === "PELANGGARAN" || h === "KENDALA" || h === "DESKRIPSI" || h === "CATATAN" || h === "PERIHAL")) colMap.keterangan = c;
      if (colMap.filePdfUrl === -1 && (h === "FILEPDFURL" || h === "FILEURL" || h === "PDFURL" || h === "LINKPDF" || h === "LINKDRIVE" || h === "URLPDF" || h === "FILE" || h === "BERKASPDF" || h === "LINKBERKAS")) colMap.filePdfUrl = c;
      if (colMap.filePdfId === -1 && (h === "FILEPDFID" || h === "FILEID" || h === "PDFID" || h === "DRIVEID" || h === "IDDRIVE" || h === "IDFILE")) colMap.filePdfId = c;
      if (colMap.statusAktif === -1 && (h === "STATUSAKTIF" || h === "STATUS" || h === "STATUSSANKSI" || h === "STATUSSURAT")) colMap.statusAktif = c;
      if (colMap.klw === -1 && (h === "KLW" || h === "KOORDINATOR" || h === "KOORDINATORLW")) colMap.klw = c;
      if (colMap.tglUpload === -1 && (h === "TANGGALUPLOAD" || h === "TGLUPLOAD" || h === "CREATEDAT" || h === "WAKTUINPUT" || h === "TIMESTAMP")) colMap.tglUpload = c;
    }

    // Default indeks jika kolom belum bernama standar (Indeks 0 s/d 12 baku)
    if (colMap.idSp === -1) colMap.idSp = 0;
    if (colMap.namaLw === -1) colMap.namaLw = 1;
    if (colMap.nip === -1) colMap.nip = 2;
    if (colMap.jenisSurat === -1) colMap.jenisSurat = 3;
    if (colMap.nomorSurat === -1) colMap.nomorSurat = 4;
    if (colMap.tglSurat === -1) colMap.tglSurat = 5;
    if (colMap.masaBulan === -1) colMap.masaBulan = 6;
    if (colMap.tglBerakhir === -1) colMap.tglBerakhir = 7;
    if (colMap.keterangan === -1) colMap.keterangan = 8;
    if (colMap.filePdfUrl === -1) colMap.filePdfUrl = 9;
    if (colMap.filePdfId === -1) colMap.filePdfId = 10;
    if (colMap.statusAktif === -1) colMap.statusAktif = 11;
    if (colMap.tglUpload === -1) colMap.tglUpload = 12;

    // Cache Data LW untuk fallback mapping KLW
    var sheetLW = ss.getSheetByName("Data LW") || ss.getSheetByName("DataLW");
    var lwKlwMap = {};
    if (sheetLW && sheetLW.getLastRow() > 1) {
      var dLw = sheetLW.getRange(2, 1, sheetLW.getLastRow() - 1, Math.max(sheetLW.getLastColumn(), 2)).getDisplayValues();
      for (var l = 0; l < dLw.length; l++) {
        var nLw = String(dLw[l][0] || "").trim().toUpperCase();
        var kLw = String(dLw[l][1] || "").trim();
        if (nLw) lwKlwMap[nLw] = kLw;
      }
    }

    var rawData = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
    var q = (searchQuery || "").toLowerCase().trim();
    var jf = (jenisFilter || "ALL").toUpperCase().trim();
    var sf = (statusFilter || "ALL").toLowerCase().trim();
    var pg = parseInt(page, 10) || 1;
    var lmt = parseInt(limit, 10) || 10;

    var today = new Date();
    today.setHours(0, 0, 0, 0);

    var filtered = [];

    for (var i = rawData.length - 1; i >= 0; i--) {
      var row = rawData[i];
      var rawIdSp = (colMap.idSp >= 0 ? String(row[colMap.idSp] || "") : "").trim();
      var namaLw = (colMap.namaLw >= 0 ? String(row[colMap.namaLw] || "") : "").trim();

      // Lewati baris kosong / ghost rows di Google Sheets
      if (!namaLw && !rawIdSp) continue;

      var idSp = rawIdSp || ("SP-" + (i + 1));
      var nip = (colMap.nip >= 0 ? row[colMap.nip] : "") || "-";
      var rawJenis = (colMap.jenisSurat >= 0 ? row[colMap.jenisSurat] : "") || "ST";
      var jenisSurat = rawJenis.trim();
      var nomorSurat = (colMap.nomorSurat >= 0 ? row[colMap.nomorSurat] : "") || "-";
      var tglSurat = (colMap.tglSurat >= 0 ? row[colMap.tglSurat] : "") || "";
      var masaBulan = parseInt(colMap.masaBulan >= 0 ? row[colMap.masaBulan] : "", 10) || (jenisSurat.indexOf("ST") !== -1 ? 3 : 6);
      var tglBerakhirStr = (colMap.tglBerakhir >= 0 ? row[colMap.tglBerakhir] : "") || "";
      var keterangan = (colMap.keterangan >= 0 ? row[colMap.keterangan] : "") || "-";
      var filePdfUrl = ((colMap.filePdfUrl >= 0 ? row[colMap.filePdfUrl] : "") || "").trim();
      var filePdfId = ((colMap.filePdfId >= 0 ? row[colMap.filePdfId] : "") || "").trim();
      var rawStatus = (colMap.statusAktif >= 0 ? row[colMap.statusAktif] : "") || "";
      var klw = (colMap.klw >= 0 ? row[colMap.klw] : "") || lwKlwMap[namaLw.toUpperCase()] || "LAINNYA / UMUM";
      var tglUpload = (colMap.tglUpload >= 0 ? row[colMap.tglUpload] : "") || "";

      // Kalkulasi tanggal berakhir otomatis jika kosong
      var endDate = null;
      if (tglBerakhirStr) {
        endDate = parseDateFlexible(tglBerakhirStr);
      } else if (tglSurat) {
        var sDate = parseDateFlexible(tglSurat);
        endDate = new Date(sDate.getTime());
        endDate.setMonth(endDate.getMonth() + masaBulan);
        tglBerakhirStr = Utilities.formatDate(endDate, "Asia/Jakarta", "dd/MM/yyyy");
      } else {
        endDate = new Date(today.getTime());
        endDate.setMonth(endDate.getMonth() + masaBulan);
        tglBerakhirStr = Utilities.formatDate(endDate, "Asia/Jakarta", "dd/MM/yyyy");
      }

      var diffDays = 0;
      var isAktif = true;
      try {
        endDate.setHours(23, 59, 59, 999);
        var diffTime = endDate.getTime() - today.getTime();
        diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
        isAktif = diffDays >= 0;
      } catch (errDate) {
        isAktif = true;
        diffDays = 30;
      }

      if (rawStatus.toLowerCase().indexOf("kadaluarsa") !== -1 || rawStatus.toLowerCase().indexOf("eskalasi") !== -1 || rawStatus.toLowerCase().indexOf("selesai") !== -1 || rawStatus.toLowerCase().indexOf("bebas") !== -1) {
        isAktif = false;
      }
      var statusReal = isAktif ? "Aktif" : "Kadaluarsa";

      // Filter check
      var matchSearch = (q === "" || (idSp + " " + namaLw + " " + nip + " " + nomorSurat + " " + keterangan + " " + klw).toLowerCase().indexOf(q) !== -1);
      var matchJenis = (jf === "ALL" || jenisSurat.toUpperCase() === jf || jenisSurat.toUpperCase().indexOf(jf) !== -1);
      var matchStatus = (sf === "all" || statusReal.toLowerCase() === sf);

      if (matchSearch && matchJenis && matchStatus) {
        var driveId = extractDriveFileId(filePdfId) || extractDriveFileId(filePdfUrl);
        var previewUrl = "";
        var downloadUrl = "";
        var hasValidFile = false;

        if (driveId) {
          hasValidFile = true;
          previewUrl = "https://drive.google.com/file/d/" + driveId + "/preview";
          downloadUrl = "https://drive.google.com/uc?export=download&id=" + driveId;
        } else if (filePdfUrl && filePdfUrl !== "#" && filePdfUrl.indexOf("http") === 0) {
          hasValidFile = true;
          previewUrl = filePdfUrl;
          downloadUrl = filePdfUrl;
        }

        filtered.push({
          idSp: idSp,
          namaLw: namaLw,
          nip: nip,
          klw: klw,
          jenisSurat: jenisSurat,
          nomorSurat: nomorSurat,
          tglSurat: tglSurat,
          masaBulan: masaBulan,
          tglBerakhir: tglBerakhirStr,
          keterangan: keterangan,
          filePdfUrl: filePdfUrl,
          filePdfId: filePdfId,
          previewUrl: previewUrl,
          downloadUrl: downloadUrl,
          hasValidFile: hasValidFile,
          statusAktif: statusReal,
          sisaHari: diffDays > 0 ? diffDays : 0,
          tglUpload: tglUpload
        });
      }
    }

    var total = filtered.length;
    var totalPages = Math.ceil(total / lmt) || 1;
    var startIndex = (pg - 1) * lmt;
    var pagedData = filtered.slice(startIndex, startIndex + lmt);

    return {
      status: "success",
      data: pagedData,
      total: total,
      page: pg,
      totalPages: totalPages,
      debug: {
        sheetUsed: sheet.getName(),
        rowCount: lastRow - 1,
        headerRow: headers.slice(0, 15),
        passedFilter: total,
        colMap: colMap
      }
    };
  } catch (err) {
    Logger.log("Error getSuratPeringatanList: " + err.toString());
    return { status: "error", message: err.toString() };
  }
}

function saveSuratPeringatan(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getSuratPeringatanSheet(ss);
    if (!sheet) {
      sheet = ss.insertSheet("SuratPeringatan");
      sheet.appendRow(["ID_SP", "NAMA_LW", "NIP", "JENIS_SURAT", "NOMOR_SURAT", "TANGGAL_SURAT", "MASA_BERLAKU_BULAN", "TANGGAL_BERAKHIR", "KETERANGAN", "FILE_PDF_URL", "FILE_PDF_ID", "STATUS_AKTIF", "TANGGAL_UPLOAD"]);
    }

    var jenisSurat = (payload.jenisSurat || "ST").trim();
    var masaBulan = (jenisSurat === "ST") ? 3 : 6;

    // Hitung tanggal berakhir
    var startDate = parseDateFlexible(payload.tanggalSurat);
    var endDate = new Date(startDate.getTime());
    endDate.setMonth(endDate.getMonth() + masaBulan);

    var tglSuratFormatted = Utilities.formatDate(startDate, "Asia/Jakarta", "dd/MM/yyyy");
    var tglBerakhirFormatted = Utilities.formatDate(endDate, "Asia/Jakarta", "dd/MM/yyyy");
    var nowStr = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");

    var newId = generateSequentialId("SP", sheet.getName());

    var filePdfUrl = (payload.fileUrl || "").trim() || "#";
    var filePdfId = "";

    // Upload file PDF ke Google Drive jika ada base64
    if (payload.fileBase64 && payload.fileBase64.trim() !== "") {
      try {
        var base64Data = payload.fileBase64;
        if (base64Data.indexOf("base64,") !== -1) {
          base64Data = base64Data.split("base64,")[1];
        }
        var decoded = Utilities.base64Decode(base64Data);
        var fileName = payload.fileName || (newId + "_" + jenisSurat + "_" + payload.namaLW.replace(/[^a-zA-Z0-9]/g, "_") + ".pdf");
        var blob = Utilities.newBlob(decoded, "application/pdf", fileName);

        var folder = getOrCreateDriveFolder("UPT_Semarang_SuratPeringatan_PDF");
        var driveFile = folder.createFile(blob);
        driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

        filePdfId = driveFile.getId();
        filePdfUrl = driveFile.getUrl();
      } catch (uploadErr) {
        Logger.log("Gagal upload PDF ke Drive: " + uploadErr.toString());
      }
    } else if (filePdfUrl !== "#") {
      filePdfId = extractDriveFileId(filePdfUrl);
    }

    var rowData = [
      newId,
      payload.namaLW,
      payload.nip || "000000",
      jenisSurat,
      payload.nomorSurat || "-",
      tglSuratFormatted,
      masaBulan,
      tglBerakhirFormatted,
      payload.keterangan || "-",
      filePdfUrl,
      filePdfId,
      "Aktif",
      nowStr
    ];

    sheet.appendRow(rowData);
    SpreadsheetApp.flush();

    return { 
      status: "success", 
      message: "Surat " + jenisSurat + " (" + payload.namaLW + ") berhasil disimpan dengan masa berlaku " + masaBulan + " bulan!" 
    };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function updateSuratPeringatanPdf(idSp, fileBase64, fileName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getSuratPeringatanSheet(ss);
    if (!sheet || sheet.getLastRow() <= 1) {
      return { status: "error", message: "Sheet SuratPeringatan tidak ditemukan atau kosong!" };
    }

    var folder = getOrCreateDriveFolder("UPT_Semarang_SuratPeringatan_PDF");
    var base64Data = fileBase64;
    if (base64Data.indexOf("base64,") !== -1) {
      base64Data = base64Data.split("base64,")[1];
    }
    var bytes = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(bytes, "application/pdf", idSp + "_" + fileName);
    var driveFile = folder.createFile(blob);
    driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileUrl = driveFile.getUrl();
    var fileId = driveFile.getId();

    var numCols = Math.max(sheet.getLastColumn(), 13);
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, numCols).getDisplayValues();
    var found = false;
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === idSp) {
        var rowIdx = i + 2;
        sheet.getRange(rowIdx, 10).setValue(fileUrl);
        sheet.getRange(rowIdx, 11).setValue(fileId);
        sheet.getRange(rowIdx, 13).setValue(Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss"));
        found = true;
        break;
      }
    }

    if (!found) {
      return { status: "error", message: "Data Surat Peringatan " + idSp + " tidak ditemukan!" };
    }

    SpreadsheetApp.flush();
    return { status: "success", message: "Berkas PDF untuk " + idSp + " berhasil diperbarui!", fileUrl: fileUrl };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function eskalasiSuratPeringatan(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getSuratPeringatanSheet(ss);
    if (!sheet) {
      sheet = ss.insertSheet("SuratPeringatan");
      sheet.appendRow(["ID_SP", "NAMA_LW", "NIP", "JENIS_SURAT", "NOMOR_SURAT", "TANGGAL_SURAT", "MASA_BERLAKU_BULAN", "TANGGAL_BERAKHIR", "KETERANGAN", "FILE_PDF_URL", "FILE_PDF_ID", "STATUS_AKTIF", "TANGGAL_UPLOAD"]);
    }

    var jenisSurat = (payload.jenisSuratBaru || "SP1").trim();
    var masaBulan = (jenisSurat === "ST") ? 3 : 6;

    var startDate = parseDateFlexible(payload.tanggalSuratBaru || new Date());
    var endDate = new Date(startDate.getTime());
    endDate.setMonth(endDate.getMonth() + masaBulan);

    var tglSuratFormatted = Utilities.formatDate(startDate, "Asia/Jakarta", "dd/MM/yyyy");
    var tglBerakhirFormatted = Utilities.formatDate(endDate, "Asia/Jakarta", "dd/MM/yyyy");
    var nowStr = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");

    var newId = generateSequentialId("SP", sheet.getName());
    var filePdfUrl = "#";
    var filePdfId = "";

    if (payload.fileBase64 && payload.fileBase64.trim() !== "") {
      try {
        var base64Data = payload.fileBase64;
        if (base64Data.indexOf("base64,") !== -1) {
          base64Data = base64Data.split("base64,")[1];
        }
        var decoded = Utilities.base64Decode(base64Data);
        var fileName = payload.fileName || (newId + "_" + jenisSurat + "_" + (payload.namaLW || "LW").replace(/[^a-zA-Z0-9]/g, "_") + ".pdf");
        var blob = Utilities.newBlob(decoded, "application/pdf", fileName);

        var folder = getOrCreateDriveFolder("UPT_Semarang_SuratPeringatan_PDF");
        var driveFile = folder.createFile(blob);
        driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

        filePdfId = driveFile.getId();
        filePdfUrl = driveFile.getUrl();
      } catch (uploadErr) {
        Logger.log("Gagal upload PDF eskalasi ke Drive: " + uploadErr.toString());
      }
    }

    // Catat eskalasi pada surat lama jika ada idSpLama
    if (payload.idSpLama && sheet.getLastRow() > 1) {
      var numColsEskalasi = Math.max(sheet.getLastColumn(), 13);
      var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, numColsEskalasi).getDisplayValues();
      for (var i = 0; i < data.length; i++) {
        if (data[i][0] === payload.idSpLama) {
          var rowIdx = i + 2;
          var ketLama = data[i][8] || "";
          sheet.getRange(rowIdx, 9).setValue(ketLama + " [Dieskalasi ke " + jenisSurat + " No: " + (payload.nomorSuratBaru || "-") + " (" + newId + ")]");
          sheet.getRange(rowIdx, 12).setValue("Kadaluarsa (Dieskalasi)");
          break;
        }
      }
    }

    var rowData = [
      newId,
      payload.namaLW,
      payload.nip || "000000",
      jenisSurat,
      payload.nomorSuratBaru || "-",
      tglSuratFormatted,
      masaBulan,
      tglBerakhirFormatted,
      payload.keteranganBaru || ("Eskalasi pelanggaran sebelum masa sanksi usai dari surat sebelumnya (" + (payload.idSpLama || "-") + ")"),
      filePdfUrl,
      filePdfId,
      "Aktif",
      nowStr
    ];

    sheet.appendRow(rowData);
    SpreadsheetApp.flush();

    return {
      status: "success",
      message: "Berhasil memperbarui & mengeskalasi sanksi menjadi " + jenisSurat + " (" + payload.namaLW + ") dengan masa berlaku baru " + masaBulan + " bulan!",
      idSp: newId
    };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function masukkanAbsensiKeCeklis(records) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetAcuan = ss.getSheetByName("AcuanKendala");
    if (!sheetAcuan) {
      sheetAcuan = ss.insertSheet("AcuanKendala");
      sheetAcuan.appendRow(["JENIS_BA", "NAMA", "JABATAN", "NIP", "PENGHANTAR", "GARDU_INDUK", "ULTG", "TANGGAL", "BULAN", "KETERANGAN_KENDALA", "KLW", "TL_GI", "TANGGAL_SURAT", "KETERANGAN"]);
    }

    var sheetLW = ss.getSheetByName("Data LW");
    var masterMap = {};
    if (sheetLW && sheetLW.getLastRow() > 1) {
      var lastCol = Math.max(sheetLW.getLastColumn(), 8);
      var lwData = sheetLW.getRange(2, 1, sheetLW.getLastRow() - 1, lastCol).getDisplayValues();
      for (var i = 0; i < lwData.length; i++) {
        var n = (lwData[i][0] || "").trim().toLowerCase();
        if (n !== "") {
          masterMap[n] = {
            klw: lwData[i][1] || "-",
            nip: lwData[i][2] || "-",
            posisi: lwData[i][3] || "Line Walker",
            ultg: lwData[i][4] || "-",
            gi: lwData[i][5] || "-",
            tlGi: lwData[i][6] || "-",
            penghantar: lwData[i][7] || "-"
          };
        }
      }
    }

    if (!records || !Array.isArray(records) || records.length === 0) {
      return { status: "error", message: "Tidak ada personel yang dipilih untuk dimasukkan ke Ceklis!" };
    }

    // Peta existing acuan untuk mencegah duplikasi absensi
    var existingKeys = {};
    if (sheetAcuan.getLastRow() > 1) {
      var existRaw = sheetAcuan.getRange(2, 1, sheetAcuan.getLastRow() - 1, sheetAcuan.getLastColumn()).getDisplayValues();
      for (var e = 0; e < existRaw.length; e++) {
        var exNama = (existRaw[e][1] || "").trim().toLowerCase();
        var exTgl = parseInt(existRaw[e][7] || "0", 10);
        var exBln = normalizeMonthString(existRaw[e][8] || "");
        if (exNama && exTgl && exBln) {
          existingKeys[exNama + "|" + exTgl + "|" + exBln] = true;
        }
      }
    }

    var now = new Date();
    var todaySuratStr = Utilities.formatDate(now, "Asia/Jakarta", "dd MMMM yyyy");
    var insertedCount = 0;

    for (var r = 0; r < records.length; r++) {
      var item = records[r];
      var namaKey = (item.nama || "").trim().toLowerCase();
      var tglNum = parseInt(item.tanggal || now.getDate(), 10);
      var tglStr = ("0" + tglNum).slice(-2);
      var blnRaw = (item.bulan || "OCTOBER").toUpperCase();
      var blnNorm = normalizeMonthString(blnRaw);
      var jenisBA = "BERITA ACARA KENDALA ABSENSI & COC";

      var checkKey = namaKey + "|" + tglNum + "|" + blnNorm;
      if (!existingKeys[checkKey]) {
        var match = masterMap[namaKey] || {
          klw: "-",
          nip: item.nip || "000000",
          posisi: "Line Walker",
          ultg: item.ultg || "-",
          gi: "-",
          tlGi: "-",
          penghantar: "-"
        };

        var ket = item.keterangan || (item.status + " pada Tgl " + tglStr + " " + blnRaw);

        sheetAcuan.appendRow([
          jenisBA,
          item.nama,
          match.posisi,
          match.nip !== "-" ? match.nip : (item.nip || "000000"),
          match.penghantar,
          match.gi,
          match.ultg,
          tglStr,
          blnRaw,
          ket,
          match.klw,
          match.tlGi,
          todaySuratStr,
          "Pelanggaran Absensi Mingguan (" + (item.periodeMinggu || "M") + ")"
        ]);

        existingKeys[checkKey] = true;
        insertedCount++;
      }
    }

    SpreadsheetApp.flush();
    return {
      status: "success",
      count: insertedCount,
      message: insertedCount > 0
        ? "Berhasil menyinkronkan " + insertedCount + " pelanggaran absensi ke Menu Ceklis & Acuan Kendala!"
        : "Data absensi yang dipilih sudah ada di Menu Ceklis."
    };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function deleteCeklisItem(rowId, nip, tanggal, bulan, tahun, jenisKendala, nama) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetAcuan = ss.getSheetByName("AcuanKendala");
    if (!sheetAcuan || sheetAcuan.getLastRow() <= 1) {
      return { status: "error", message: "Sheet AcuanKendala kosong atau tidak ditemukan!" };
    }

    var lastColA = Math.max(sheetAcuan.getLastColumn(), 14);
    var rawAcuan = sheetAcuan.getRange(2, 1, sheetAcuan.getLastRow() - 1, lastColA).getDisplayValues();
    var targetRowIndex = -1;

    var targetNama = (nama || "").trim().toLowerCase();
    var targetNip = (nip || "").trim();
    var targetTgl = parseInt(tanggal || "0", 10);
    var targetBln = normalizeMonthString(bulan || "");
    var targetJenis = (jenisKendala || "").trim().toLowerCase();

    // 1. Cek kecocokan rowId bila dalam rentang yang valid
    if (rowId && rowId >= 2 && rowId <= rawAcuan.length + 1) {
      var candidate = rawAcuan[rowId - 2];
      var cNama = (candidate[1] || "").trim().toLowerCase();
      var cTgl = parseInt(candidate[7] || "0", 10);
      var cBln = normalizeMonthString(candidate[8] || "");
      if ((!targetNama || cNama === targetNama) && (!targetTgl || cTgl === targetTgl) && (!targetBln || cBln === targetBln)) {
        targetRowIndex = rowId;
      }
    }

    // 2. Jika rowId belum cocok, scan baris yang sesuai dari AcuanKendala
    if (targetRowIndex === -1) {
      for (var i = rawAcuan.length - 1; i >= 0; i--) {
        var rowNama = (rawAcuan[i][1] || "").trim().toLowerCase();
        var rowNip = (rawAcuan[i][3] || "").trim();
        var rowTgl = parseInt(rawAcuan[i][7] || "0", 10);
        var rowBln = normalizeMonthString(rawAcuan[i][8] || "");
        var rowJenis = (rawAcuan[i][0] || "").trim().toLowerCase();

        var matchNama = (!targetNama || rowNama === targetNama);
        var matchNip = (!targetNip || rowNip === targetNip || rowNip === "000000" || targetNip === "000000");
        var matchTgl = (!targetTgl || rowTgl === targetTgl);
        var matchBln = (!targetBln || rowBln === targetBln);
        var matchJenis = (!targetJenis || rowJenis.indexOf(targetJenis) !== -1 || targetJenis.indexOf(rowJenis) !== -1);

        if (matchNama && matchTgl && matchBln && (matchJenis || matchNip)) {
          targetRowIndex = i + 2;
          break;
        }
      }
    }

    if (targetRowIndex === -1) {
      return { status: "error", message: "Data tidak ditemukan di AcuanKendala!" };
    }

    // Hapus baris dari AcuanKendala
    sheetAcuan.deleteRow(targetRowIndex);

    // Hapus juga status dari sheet Ceklis bila ada (spesifik per tanggal kendala)
    var sheetCeklis = ss.getSheetByName("Ceklis");
    if (sheetCeklis && sheetCeklis.getLastRow() > 1) {
      var lastColC = Math.max(sheetCeklis.getLastColumn(), 11);
      var rawCeklis = sheetCeklis.getRange(2, 1, sheetCeklis.getLastRow() - 1, lastColC).getDisplayValues();
      for (var c = rawCeklis.length - 1; c >= 0; c--) {
        var cNama = (rawCeklis[c][1] || "").trim().toLowerCase();
        var cNip = (rawCeklis[c][2] || "").trim();
        var cBulan = normalizeMonthString(rawCeklis[c][3] || "");
        var cCat = (rawCeklis[c][5] || "").trim().toLowerCase();
        var cTanggal = rawCeklis[c].length >= 11 ? normalizeTanggal(rawCeklis[c][10]) : "";

        var validNip = (targetNip && targetNip !== "000000" && targetNip !== "-");
        var matchPerson = (validNip && cNip === targetNip) || (targetNama && cNama === targetNama);
        var matchBulan = (cBulan === targetBln);
        var matchCat = (!targetJenis || cCat.indexOf(targetJenis) !== -1 || targetJenis.indexOf(cCat) !== -1);
        var matchTanggal = (!targetTgl || !cTanggal || parseInt(cTanggal, 10) === targetTgl);

        if (matchPerson && matchBulan && matchCat && matchTanggal) {
          sheetCeklis.deleteRow(c + 2);
          break;
        }
      }
    }

    SpreadsheetApp.flush();
    return {
      status: "success",
      message: "Data berhasil dihapus dari Menu Ceklis! Absensi ini kini dapat disinkronkan kembali dari Rekap Mingguan."
    };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

// ============================================================================
// MODUL BERITA ACARA & AUTOCRAT FORM TRIGGER CONTROLLER (CLEAN ARCHITECTURE)
// ============================================================================

var GOOGLE_FORM_ID = "1V-syWXWHLox6bRMXrb5MnjO3kvd82_foBWJt0QoyHIY";
var GOOGLE_FORM_SPREADSHEET_ID = "1rdWG9RD_Av8lA-XCvvRdh8AESgNyb6NneD831ojzZ8M";
var SHEET_DATA_LW = "DataLW";

/**
 * Generate ID Registrasi Berita Acara Baru Sekuensial (BA-YYYYMMDD-XXXX)
 */
function getNewBeritaAcaraId() {
  try {
    return generateSequentialId("BA", "Form_Responses");
  } catch (e) {
    var todayStr = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyyMMdd");
    return "BA-" + todayStr + "-0001";
  }
}

/**
 * Buka instance Google Form resmi dengan fallback
 */
function getGoogleFormInstance() {
  if (GOOGLE_FORM_ID) {
    try {
      var form = FormApp.openById(GOOGLE_FORM_ID);
      if (form) return form;
    } catch (e1) {
      Logger.log("FormApp.openById error: " + e1.toString());
    }
  }

  if (GOOGLE_FORM_SPREADSHEET_ID) {
    try {
      var formSs = SpreadsheetApp.openById(GOOGLE_FORM_SPREADSHEET_ID);
      if (formSs && formSs.getFormUrl()) {
        var form = FormApp.openByUrl(formSs.getFormUrl());
        if (form) return form;
      }
    } catch (e2) {
      Logger.log("FormApp.openByUrl error: " + e2.toString());
    }
  }

  return null;
}

/**
 * Mengambil Master Data Petugas Line Walker dari Sheet DataLW
 * Untuk dropdown dan autofill personel di Form Berita Acara
 */
function getLineWalkerMaster() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_DATA_LW) || ss.getSheetByName("Data LW") || ss.getSheetByName("DataLineWalker");
    
    if (!sheet) {
      return { success: true, data: [] };
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, data: [] };
    }
    
    var values = sheet.getRange(2, 1, lastRow - 1, 8).getDisplayValues();
    var masterList = values.map(function(row) {
      return {
        nama: (row[0] || "").trim(),
        klw: (row[1] || "").trim(),
        nip: (row[2] || "").trim(),
        posisi: (row[3] || "").trim(),
        ultg: (row[4] || "").trim(),
        gi: (row[5] || "").trim(),
        tlGi: (row[6] || "").trim(),
        penghantar: (row[7] || "").trim()
      };
    }).filter(function(item) {
      return item.nama !== "";
    });
    
    return {
      success: true,
      data: masterList
    };
  } catch (err) {
    Logger.log("Error getLineWalkerMaster: " + err.toString());
    return {
      success: false,
      message: "Gagal memuat master data Line Walker: " + err.toString(),
      data: []
    };
  }
}

/**
 * Simpan Data Berita Acara ke Sheet Form_Responses & Pemicu Trigger AutoCrat Kosong
 */
function submitBeritaAcaraForm(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    
    if (!payload || typeof payload !== "object") {
      throw new Error("Data formulir tidak valid.");
    }
    
    if (!payload.nama || !payload.nama.trim()) throw new Error("Nama Petugas wajib dipilih.");
    if (!payload.jenisBa || !payload.jenisBa.trim()) throw new Error("Jenis Berita Acara wajib dipilih.");
    if (!payload.keteranganKendala || !payload.keteranganKendala.trim()) throw new Error("Uraian keterangan kendala wajib diisi.");

    var now = new Date();
    var idBa = payload.idBa && payload.idBa.indexOf("BA-") === 0
      ? payload.idBa
      : generateSequentialId("BA", "Form_Responses");
    var nowTimestamp = Utilities.formatDate(now, "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");
    
    var tanggalSurat = payload.tanggalSurat && payload.tanggalSurat.trim() !== ""
      ? payload.tanggalSurat.trim()
      : Utilities.formatDate(now, "Asia/Jakarta", "dd MMMM yyyy");
      
    var statusDoc = payload.status || "Tercatat";
    
    // 1. Simpan baris data 16 kolom ke Sheet Respon Form ("Form_Responses")
    var formSs = null;
    try {
      if (GOOGLE_FORM_SPREADSHEET_ID) {
        formSs = SpreadsheetApp.openById(GOOGLE_FORM_SPREADSHEET_ID);
      }
    } catch (ssErr) {
      Logger.log("openById(formSsId) error: " + ssErr.toString());
    }

    if (!formSs) {
      try {
        formSs = SpreadsheetApp.getActiveSpreadsheet();
      } catch (e2) {}
    }

    if (formSs) {
      var sheetRespon = formSs.getSheetByName("Form_Responses") || formSs.getSheetByName("Form Responses 1") || formSs.getSheetByName("Respon Formulir 1") || formSs.getSheets()[0];
      if (sheetRespon) {
        var rowData = [
          idBa,
          payload.jenisBa.trim(),
          payload.nama.trim(),
          (payload.jabatan || "Line Walker").trim(),
          (payload.nip || "").trim(),
          (payload.penghantar || "-").trim(),
          (payload.garduInduk || "-").trim(),
          (payload.ultg || "-").trim(),
          (payload.tanggal || Utilities.formatDate(now, "Asia/Jakarta", "dd")).toString().trim(),
          (payload.bulan || Utilities.formatDate(now, "Asia/Jakarta", "MMMM")).toUpperCase().trim(),
          payload.keteranganKendala.trim(),
          (payload.klw || "").trim(),
          (payload.tlGi || "").trim(),
          tanggalSurat,
          statusDoc,
          nowTimestamp
        ];
        
        sheetRespon.appendRow(rowData);
        SpreadsheetApp.flush();
      }
    }
    
    // 2. Memicu trigger submit Google Form kosong (AutoCrat onFormSubmit)
    var formSubmitResult = forwardResponseToGoogleForm();
    
    var msgFeedback = formSubmitResult.success
      ? "Berhasil menyimpan Berita Acara (" + idBa + ") & memicu pencetakan dokumen AutoCrat!"
      : "Data tersimpan ke Respon Form. (Trigger Form: " + (formSubmitResult.error || "Tercatat") + ")";

    return {
      status: "success",
      success: true,
      idBa: idBa,
      message: msgFeedback,
      formSynced: formSubmitResult.success,
      data: {
        idBa: idBa,
        jenisBa: payload.jenisBa,
        nama: payload.nama,
        tanggalInput: nowTimestamp,
        penghantar: payload.penghantar || "-",
        garduInduk: payload.garduInduk || "-",
        status: statusDoc
      }
    };
  } catch (err) {
    Logger.log("Error submitBeritaAcaraForm: " + err.toString());
    return {
      status: "error",
      success: false,
      message: "Gagal menyimpan Berita Acara: " + err.toString()
    };
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}

function submitBeritaAcaraToGoogleForm(payload) {
  return submitBeritaAcaraForm(payload);
}

/**
 * Submit respons kosong ke Google Form resmi untuk memicu event onFormSubmit AutoCrat
 */
function forwardResponseToGoogleForm() {
  try {
    var form = getGoogleFormInstance();
    if (!form) {
      Logger.log("Info: Google Form tidak terdeteksi via FormApp.");
      return { success: false, reason: "NO_FORM_INSTANCE" };
    }
    
    // Submit respons kosong murni sebagai trigger AutoCrat
    var formResponse = form.createResponse();
    formResponse.submit();
    
    Logger.log("Sukses: Submit Google Form trigger berhasil dieksekusi!");
    return { success: true };
  } catch (err) {
    Logger.log("Warning forwardResponseToGoogleForm: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

// ==============================================================================
// MODUL DIGITAL APPROVAL STUDIO & MULTI-SIGNATURE PDF
// 1. Multi-Titik Tanda Tangan Dinamis (1, 2, 3+ Approver, Drag & Resize Koordinat X,Y)
// 2. Penempatan Tepat Sasaran (Bottom-Left Standard PDF Coordinates via pdf-lib)
// 3. Tanda Tangan Digital (Canvas Touch/Mouse, Upload PNG/JPG, QR Verification)
// 4. Penyimpanan Otomatis Google Drive (Folder Terorganisir, Status Tracking Sheets)
// ==============================================================================

var SHEET_DOCUMENTS = "DataDokumen";
var SHEET_SPOTS = "TitikTandaTangan";
var SHEET_APPROVERS = "Master_Approver";
var STORAGE_FOLDER_APPROVAL_NAME = "UPT_Semarang_Approval_Storage";

const COL_DOC = {
  ID: 0,                     // Kolom A
  TITLE: 1,                  // Kolom B
  CREATED_AT: 2,             // Kolom C
  ORIGINAL_FILE_NAME: 3,     // Kolom D
  DRIVE_FILE_ID_ORIGINAL: 4, // Kolom E
  DRIVE_FILE_ID_SIGNED: 5,   // Kolom F
  TOTAL_PAGES: 6,            // Kolom G
  TOTAL_SPOTS: 7,            // Kolom H
  SIGNED_COUNT: 8,           // Kolom I
  STATUS: 9                  // Kolom J
};

const COL_SPOT = {
  SPOT_ID: 0,                // Kolom A
  DOC_ID: 1,                 // Kolom B
  SIGNER_NAME: 2,            // Kolom C
  SIGNER_ROLE: 3,            // Kolom D
  PAGE_NUMBER: 4,            // Kolom E
  X_PERCENT: 5,              // Kolom F
  Y_PERCENT: 6,              // Kolom G
  WIDTH_PERCENT: 7,          // Kolom H
  HEIGHT_PERCENT: 8,         // Kolom I
  IS_SIGNED: 9,              // Kolom J
  SIGNED_AT: 10,             // Kolom K
  VERIFICATION_HASH: 11      // Kolom L
};

/**
 * Helper: Parse angka float aman, tahan koma desimal Indonesia & nilai 0
 */
function parseSafeFloat(val, fallback) {
  if (val === null || val === undefined || val === '') return (fallback !== undefined ? fallback : 0);
  if (typeof val === 'number') return isNaN(val) ? (fallback !== undefined ? fallback : 0) : val;
  var s = String(val).replace(',', '.').trim();
  var n = parseFloat(s);
  return isNaN(n) ? (fallback !== undefined ? fallback : 0) : n;
}

/**
 * Mendapatkan atau membuat folder penyimpanan Google Drive khusus Approval Dokumen
 */
function getApprovalStorageFolder() {
  var props = PropertiesService.getScriptProperties();
  var savedFolderId = props.getProperty("STORAGE_FOLDER_APPROVAL_ID");
  
  if (savedFolderId) {
    try {
      return DriveApp.getFolderById(savedFolderId);
    } catch (e) {
      Logger.log("Folder approval lama tidak valid, mencari ulang...");
    }
  }

  var existingFolders = DriveApp.getFoldersByName(STORAGE_FOLDER_APPROVAL_NAME);
  if (existingFolders.hasNext()) {
    var folder = existingFolders.next();
    props.setProperty("STORAGE_FOLDER_APPROVAL_ID", folder.getId());
    return folder;
  }

  var newFolder = DriveApp.createFolder(STORAGE_FOLDER_APPROVAL_NAME);
  props.setProperty("STORAGE_FOLDER_APPROVAL_ID", newFolder.getId());
  return newFolder;
}

/**
 * Generator ID Dokumen Sekuensial Anti-Tabrakan (DOC-YYYYMM-XXXX)
 */
function apiGetNextApprovalDocId() {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_DOCUMENTS);
    if (!sheet) {
      if (typeof setupApprovalDatabase === "function") setupApprovalDatabase();
      sheet = ss.getSheetByName(SHEET_DOCUMENTS);
    }

    var now = new Date();
    var yearMonth = Utilities.formatDate(now, "Asia/Jakarta", "yyyyMM");
    var prefix = "DOC-" + yearMonth + "-";

    var lastRow = sheet ? sheet.getLastRow() : 1;
    var maxNum = 0;

    if (lastRow > 1) {
      var idValues = sheet.getRange(2, COL_DOC.ID + 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < idValues.length; i++) {
        var val = String(idValues[i][0] || "");
        if (val.indexOf(prefix) === 0) {
          var numPart = parseInt(val.replace(prefix, ""), 10);
          if (!isNaN(numPart) && numPart > maxNum) {
            maxNum = numPart;
          }
        }
      }
    }

    var nextFormatted = ("000" + (maxNum + 1)).slice(-4);
    return prefix + nextFormatted;
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**
 * API: Mengambil Data Ringkasan Dashboard Approval & Seluruh Dokumen
 */
function apiGetApprovalDashboardData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_DOCUMENTS);
    if (!sheet) {
      if (typeof setupApprovalDatabase === "function") setupApprovalDatabase();
      sheet = ss.getSheetByName(SHEET_DOCUMENTS);
    }

    if (!sheet || sheet.getLastRow() <= 1) {
      return { success: true, docs: [], stats: { total: 0, pending: 0, completed: 0 } };
    }

    var lastRow = sheet.getLastRow();
    var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    var docs = [];
    var pendingCount = 0;
    var completedCount = 0;

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var docId = String(row[COL_DOC.ID] || "");
      if (!docId) continue;

      var status = String(row[COL_DOC.STATUS] || "WAITING_APPROVAL");
      if (status === "COMPLETED") {
        completedCount++;
      } else {
        pendingCount++;
      }

      docs.push({
        id: docId,
        title: String(row[COL_DOC.TITLE] || ""),
        createdAt: row[COL_DOC.CREATED_AT] instanceof Date 
          ? Utilities.formatDate(row[COL_DOC.CREATED_AT], "Asia/Jakarta", "dd/MM/yyyy HH:mm")
          : String(row[COL_DOC.CREATED_AT] || ""),
        originalFileName: String(row[COL_DOC.ORIGINAL_FILE_NAME] || ""),
        driveFileIdOriginal: String(row[COL_DOC.DRIVE_FILE_ID_ORIGINAL] || ""),
        driveFileIdSigned: String(row[COL_DOC.DRIVE_FILE_ID_SIGNED] || ""),
        totalPages: Number(row[COL_DOC.TOTAL_PAGES] || 1),
        totalSpots: Number(row[COL_DOC.TOTAL_SPOTS] || 0),
        signedCount: Number(row[COL_DOC.SIGNED_COUNT] || 0),
        status: status
      });
    }

    // Urutkan dari dokumen yang paling baru dibuat
    docs.reverse();

    return {
      success: true,
      docs: docs,
      stats: {
        total: docs.length,
        pending: pendingCount,
        completed: completedCount
      }
    };
  } catch (err) {
    Logger.log("Error apiGetApprovalDashboardData: " + err.toString());
    return { success: false, message: err.toString(), docs: [], stats: { total: 0, pending: 0, completed: 0 } };
  }
}

/**
 * API: Menyimpan Dokumen Baru ke Google Drive dan Google Sheets
 */
function apiSaveNewApprovalDocument(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var docSheet = ss.getSheetByName(SHEET_DOCUMENTS);
    var spotSheet = ss.getSheetByName(SHEET_SPOTS);

    if (!docSheet || !spotSheet) {
      if (typeof setupApprovalDatabase === "function") setupApprovalDatabase();
      docSheet = ss.getSheetByName(SHEET_DOCUMENTS);
      spotSheet = ss.getSheetByName(SHEET_SPOTS);
    }

    var folder = getApprovalStorageFolder();
    var cleanBase64 = payload.pdfBase64.replace(/^data:application\/pdf;base64,/, "");
    var pdfBlob = Utilities.newBlob(Utilities.base64Decode(cleanBase64), "application/pdf", payload.originalFileName || "Dokumen.pdf");

    // 1. Simpan PDF Asli di Drive
    var originalFile = folder.createFile(pdfBlob);
    originalFile.setName(payload.docId + "_Original_" + payload.originalFileName);

    // 2. Simpan PDF Versi Awal untuk Signed di Drive
    var signedFile = folder.createFile(pdfBlob);
    signedFile.setName(payload.docId + "_Signed_" + payload.originalFileName);

    // Buka akses viewing siapa saja dengan link
    try {
      signedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      originalFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {
      Logger.log("Sharing notice: " + e.message);
    }

    var nowIso = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");

    // 3. Simpan Baris ke Sheet DataDokumen
    var docRow = [];
    docRow[COL_DOC.ID] = payload.docId;
    docRow[COL_DOC.TITLE] = payload.title || "Dokumen Tanpa Judul";
    docRow[COL_DOC.CREATED_AT] = nowIso;
    docRow[COL_DOC.ORIGINAL_FILE_NAME] = payload.originalFileName || "Dokumen.pdf";
    docRow[COL_DOC.DRIVE_FILE_ID_ORIGINAL] = originalFile.getId();
    docRow[COL_DOC.DRIVE_FILE_ID_SIGNED] = signedFile.getId();
    docRow[COL_DOC.TOTAL_PAGES] = Number(payload.totalPages || 1);
    docRow[COL_DOC.TOTAL_SPOTS] = payload.spots ? payload.spots.length : 0;
    docRow[COL_DOC.SIGNED_COUNT] = 0;
    docRow[COL_DOC.STATUS] = "WAITING_APPROVAL";

    docSheet.appendRow(docRow);

    // 4. Simpan Baris-Baris ke Sheet TitikTandaTangan
    if (payload.spots && payload.spots.length > 0) {
      var spotRows = [];
      payload.spots.forEach(function(spot, idx) {
        var spotId = spot.spotId || ("SPOT-" + payload.docId + "-" + (idx + 1));
        var row = [];
        row[COL_SPOT.SPOT_ID] = spotId;
        row[COL_SPOT.DOC_ID] = payload.docId;
        row[COL_SPOT.SIGNER_NAME] = spot.signerName || "Approver " + (idx + 1);
        row[COL_SPOT.SIGNER_ROLE] = spot.signerRole || "Penandatangan";
        row[COL_SPOT.PAGE_NUMBER] = Math.round(parseSafeFloat(spot.pageNumber, 1));
        row[COL_SPOT.X_PERCENT] = parseSafeFloat(spot.xPercent, 35);
        row[COL_SPOT.Y_PERCENT] = parseSafeFloat(spot.yPercent, 75);
        row[COL_SPOT.WIDTH_PERCENT] = parseSafeFloat(spot.widthPercent, 24);
        row[COL_SPOT.HEIGHT_PERCENT] = parseSafeFloat(spot.heightPercent, 7.5);
        row[COL_SPOT.IS_SIGNED] = false;
        row[COL_SPOT.SIGNED_AT] = "";
        row[COL_SPOT.VERIFICATION_HASH] = "";
        spotRows.push(row);
      });
      spotSheet.getRange(spotSheet.getLastRow() + 1, 1, spotRows.length, spotRows[0].length).setValues(spotRows);
    }

    SpreadsheetApp.flush();

    return {
      success: true,
      docId: payload.docId,
      message: "Dokumen berhasil disimpan ke Google Drive & alur approval telah dibuat!"
    };
  } catch (err) {
    Logger.log("Error apiSaveNewApprovalDocument: " + err.toString());
    return { success: false, message: "Gagal menyimpan dokumen: " + err.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**
 * API: Mengambil Detail Dokumen & File PDF dari Google Drive untuk Ruang Approval
 */
function apiGetApprovalDocumentDetail(docId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var docSheet = ss.getSheetByName(SHEET_DOCUMENTS);
    var spotSheet = ss.getSheetByName(SHEET_SPOTS);

    if (!docSheet) throw new Error("Sheet DataDokumen tidak ditemukan.");

    var docData = docSheet.getDataRange().getValues();
    var targetDoc = null;

    for (var i = 1; i < docData.length; i++) {
      if (String(docData[i][COL_DOC.ID]) === String(docId)) {
        var row = docData[i];
        targetDoc = {
          id: String(row[COL_DOC.ID]),
          title: String(row[COL_DOC.TITLE]),
          createdAt: row[COL_DOC.CREATED_AT] instanceof Date 
            ? Utilities.formatDate(row[COL_DOC.CREATED_AT], "Asia/Jakarta", "dd/MM/yyyy HH:mm")
            : String(row[COL_DOC.CREATED_AT]),
          originalFileName: String(row[COL_DOC.ORIGINAL_FILE_NAME]),
          driveFileIdSigned: String(row[COL_DOC.DRIVE_FILE_ID_SIGNED]),
          driveFileIdOriginal: String(row[COL_DOC.DRIVE_FILE_ID_ORIGINAL]),
          totalPages: Number(row[COL_DOC.TOTAL_PAGES]),
          totalSpots: Number(row[COL_DOC.TOTAL_SPOTS]),
          signedCount: Number(row[COL_DOC.SIGNED_COUNT]),
          status: String(row[COL_DOC.STATUS])
        };
        break;
      }
    }

    if (!targetDoc) throw new Error("Dokumen dengan ID " + docId + " tidak ditemukan.");

    var spots = [];
    if (spotSheet && spotSheet.getLastRow() > 1) {
      var spotData = spotSheet.getDataRange().getValues();
      for (var j = 1; j < spotData.length; j++) {
        var sRow = spotData[j];
        if (String(sRow[COL_SPOT.DOC_ID]) === String(docId)) {
          spots.push({
            spotId: String(sRow[COL_SPOT.SPOT_ID]),
            documentId: String(sRow[COL_SPOT.DOC_ID]),
            signerName: String(sRow[COL_SPOT.SIGNER_NAME]),
            signerRole: String(sRow[COL_SPOT.SIGNER_ROLE]),
            pageNumber: Math.round(parseSafeFloat(sRow[COL_SPOT.PAGE_NUMBER], 1)),
            xPercent: parseSafeFloat(sRow[COL_SPOT.X_PERCENT], 35),
            yPercent: parseSafeFloat(sRow[COL_SPOT.Y_PERCENT], 75),
            widthPercent: parseSafeFloat(sRow[COL_SPOT.WIDTH_PERCENT], 24),
            heightPercent: parseSafeFloat(sRow[COL_SPOT.HEIGHT_PERCENT], 7.5),
            isSigned: Boolean(sRow[COL_SPOT.IS_SIGNED]),
            signedAt: sRow[COL_SPOT.SIGNED_AT] instanceof Date 
              ? Utilities.formatDate(sRow[COL_SPOT.SIGNED_AT], "Asia/Jakarta", "dd/MM/yyyy HH:mm")
              : String(sRow[COL_SPOT.SIGNED_AT] || ""),
            verificationHash: String(sRow[COL_SPOT.VERIFICATION_HASH] || "")
          });
        }
      }
    }

    // Ambil Byte File PDF dari Google Drive
    var driveFileId = targetDoc.driveFileIdSigned || targetDoc.driveFileIdOriginal;
    var driveFile = DriveApp.getFileById(driveFileId);
    var pdfBytes = driveFile.getBlob().getBytes();
    var pdfBase64 = Utilities.base64Encode(pdfBytes);

    return {
      success: true,
      doc: targetDoc,
      spots: spots,
      pdfBase64: pdfBase64
    };
  } catch (err) {
    Logger.log("Error apiGetApprovalDocumentDetail: " + err.toString());
    return { success: false, message: err.toString() };
  }
}

/**
 * API: Konfirmasi Tanda Tangan & Pembaruan File PDF di Google Drive
 */
function apiConfirmApprovalSignature(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    if (!payload || typeof payload !== "object" ||
        typeof payload.docId !== "string" || !payload.docId.trim() ||
        typeof payload.spotId !== "string" || !payload.spotId.trim() ||
        typeof payload.signedPdfBase64 !== "string" || !payload.signedPdfBase64.trim()) {
      throw new Error("Dokumen, titik tanda tangan, dan PDF wajib diisi.");
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var docSheet = ss.getSheetByName(SHEET_DOCUMENTS);
    var spotSheet = ss.getSheetByName(SHEET_SPOTS);
    if (!docSheet || !spotSheet) throw new Error("Database approval belum tersedia.");

    var docId = payload.docId;
    var spotId = payload.spotId;

    // 1. Perbarui File PDF di Google Drive dengan Byte Tertanda Tangan
    var docData = docSheet.getDataRange().getValues();
    var docRowIndex = -1;
    var oldDriveFileIdSigned = "";
    var docTitle = "";

    for (var i = 1; i < docData.length; i++) {
      if (String(docData[i][COL_DOC.ID]) === String(docId)) {
        docRowIndex = i + 1;
        oldDriveFileIdSigned = String(docData[i][COL_DOC.DRIVE_FILE_ID_SIGNED]);
        docTitle = String(docData[i][COL_DOC.TITLE]);
        break;
      }
    }

    if (docRowIndex === -1) throw new Error("Dokumen " + docId + " tidak ditemukan.");

    var spotData = spotSheet.getDataRange().getValues();
    var spotRowIndex = -1;
    for (var j = 1; j < spotData.length; j++) {
      if (String(spotData[j][COL_SPOT.SPOT_ID]) === String(spotId) &&
          String(spotData[j][COL_SPOT.DOC_ID]) === String(docId)) {
        spotRowIndex = j + 1;
        var isSigned = spotData[j][COL_SPOT.IS_SIGNED];
        if (isSigned === true || String(isSigned).toLowerCase() === "true") {
          throw new Error("Titik tanda tangan sudah ditandatangani.");
        }
        break;
      }
    }
    if (spotRowIndex === -1) throw new Error("Titik tanda tangan tidak ditemukan pada dokumen ini.");

    var cleanBase64 = payload.signedPdfBase64.replace(/^data:application\/pdf;base64,/, "");
    var newBytes = Utilities.base64Decode(cleanBase64);
    var folder = getApprovalStorageFolder();
    var updatedBlob = Utilities.newBlob(newBytes, "application/pdf", docId + "_Signed.pdf");
    var newFile = folder.createFile(updatedBlob);

    try {
      newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {}

    // Update ID file signed terbaru di spreadsheet
    docSheet.getRange(docRowIndex, COL_DOC.DRIVE_FILE_ID_SIGNED + 1).setValue(newFile.getId());

    // 2. Perbarui Baris Titik TTD di Sheet TitikTandaTangan
    spotSheet.getRange(spotRowIndex, COL_SPOT.IS_SIGNED + 1).setValue(true);
    spotSheet.getRange(spotRowIndex, COL_SPOT.SIGNED_AT + 1).setValue(Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss"));
    spotSheet.getRange(spotRowIndex, COL_SPOT.VERIFICATION_HASH + 1).setValue(payload.verificationHash || "");

    // 3. Hitung Jumlah TTD Selesai
    var updatedSpotData = spotSheet.getDataRange().getValues();
    var signedCount = 0;
    var totalSpots = 0;

    for (var k = 1; k < updatedSpotData.length; k++) {
      if (String(updatedSpotData[k][COL_SPOT.DOC_ID]) === String(docId)) {
        totalSpots++;
        if (updatedSpotData[k][COL_SPOT.IS_SIGNED] === true || String(updatedSpotData[k][COL_SPOT.IS_SIGNED]).toLowerCase() === "true") {
          signedCount++;
        }
      }
    }

    var newDocStatus = (totalSpots > 0 && signedCount >= totalSpots) ? "COMPLETED" : "WAITING_APPROVAL";

    docSheet.getRange(docRowIndex, COL_DOC.SIGNED_COUNT + 1).setValue(signedCount);
    docSheet.getRange(docRowIndex, COL_DOC.STATUS + 1).setValue(newDocStatus);

    SpreadsheetApp.flush();

    return {
      success: true,
      status: newDocStatus,
      signedCount: signedCount,
      totalSpots: totalSpots,
      newFileId: newFile.getId(),
      message: "Tanda tangan berhasil dibubuhkan ke PDF dan disimpan ke Google Drive!"
    };
  } catch (err) {
    Logger.log("Error apiConfirmApprovalSignature: " + err.toString());
    return { success: false, message: err.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**
 * API: Hapus Dokumen Approval Beserta Seluruh Titik & Berkas di Google Drive
 */
function apiDeleteApprovalDocument(docId) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var docSheet = ss.getSheetByName(SHEET_DOCUMENTS);
    var spotSheet = ss.getSheetByName(SHEET_SPOTS);

    if (!docSheet) return { success: false, message: "Sheet DataDokumen tidak ditemukan." };

    var docData = docSheet.getDataRange().getValues();
    for (var i = 1; i < docData.length; i++) {
      if (String(docData[i][COL_DOC.ID]) === String(docId)) {
        var origId = String(docData[i][COL_DOC.DRIVE_FILE_ID_ORIGINAL] || "");
        var signedId = String(docData[i][COL_DOC.DRIVE_FILE_ID_SIGNED] || "");
        
        // Hapus file di Drive
        if (origId) { try { DriveApp.getFileById(origId).setTrashed(true); } catch (e) {} }
        if (signedId) { try { DriveApp.getFileById(signedId).setTrashed(true); } catch (e) {} }

        docSheet.deleteRow(i + 1);
        break;
      }
    }

    // Hapus titik tanda tangan terkait
    if (spotSheet && spotSheet.getLastRow() > 1) {
      var spotRows = spotSheet.getDataRange().getValues();
      for (var j = spotRows.length - 1; j >= 1; j--) {
        if (String(spotRows[j][COL_SPOT.DOC_ID]) === String(docId)) {
          spotSheet.deleteRow(j + 1);
        }
      }
    }

    SpreadsheetApp.flush();
    return { success: true, message: "Dokumen approval berhasil dihapus." };
  } catch (err) {
    Logger.log("Error apiDeleteApprovalDocument: " + err.toString());
    return { success: false, message: err.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**
 * API: Mengambil Dokumen Berita Acara, Work Order, dan Surat Peringatan yang memiliki PDF
 * Memudahkan pengguna mengimpor dokumen operasional langsung ke Approval Studio
 */
function apiGetSystemDocumentsForApproval() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var result = {
      beritaAcara: [],
      workOrders: [],
      suratPeringatan: []
    };

    // 1. Berita Acara (Cek dari Sheet BeritaAcara maupun Form_Responses)
    var sheetBA = ss.getSheetByName("BeritaAcara");
    var hasBAData = false;

    if (sheetBA && sheetBA.getLastRow() > 1) {
      var dataBA = sheetBA.getRange(2, 1, sheetBA.getLastRow() - 1, Math.min(sheetBA.getLastColumn(), 18)).getDisplayValues();
      dataBA.forEach(function(row) {
        var pdfUrl = row[16] || "";
        var pdfId = row[17] || "";
        if (!pdfId && pdfUrl) pdfId = extractDriveFileId(pdfUrl);

        if (row[0] && row[0] !== "-") {
          hasBAData = true;
          result.beritaAcara.push({
            id: row[0],
            title: (row[1] || "BA") + " - " + (row[2] || ""),
            nama: row[2],
            tanggal: row[8] + " " + row[9],
            pdfUrl: pdfUrl,
            pdfId: pdfId
          });
        }
      });
    }

    // Jika Sheet BeritaAcara belum terisi, coba ambil dari sheet Form_Responses (Google Form Spreadsheet)
    if (!hasBAData && typeof GOOGLE_FORM_SPREADSHEET_ID !== "undefined" && GOOGLE_FORM_SPREADSHEET_ID) {
      try {
        var formSs = SpreadsheetApp.openById(GOOGLE_FORM_SPREADSHEET_ID);
        var sheetResp = formSs.getSheetByName("Form_Responses") || formSs.getSheets()[0];
        if (sheetResp && sheetResp.getLastRow() > 1) {
          var respValues = sheetResp.getRange(2, 1, sheetResp.getLastRow() - 1, sheetResp.getLastColumn()).getDisplayValues();
          respValues.forEach(function(r) {
            var idBa = r[0] || "";
            var namaLw = r[2] || "";
            var jenisBa = r[1] || "BA";
            var tglBa = (r[8] || "") + " " + (r[9] || "");
            var foundUrl = "";
            var foundId = "";

            for (var c = r.length - 1; c >= 0; c--) {
              var cellVal = (r[c] || "").toString().trim();
              if (cellVal.indexOf("http") === 0) {
                foundUrl = cellVal;
                foundId = extractDriveFileId(cellVal);
                break;
              } else if (cellVal.length >= 20 && extractDriveFileId(cellVal)) {
                foundId = extractDriveFileId(cellVal);
                break;
              }
            }

            if (idBa && idBa !== "-") {
              result.beritaAcara.push({
                id: idBa,
                title: jenisBa + " - " + namaLw,
                nama: namaLw,
                tanggal: tglBa,
                pdfUrl: foundUrl,
                pdfId: foundId
              });
            }
          });
        }
      } catch (errForm) {
        Logger.log("Gagal membaca Google Form Responses: " + errForm.toString());
      }
    }

    // 2. Work Order SPKL
    var sheetWO = ss.getSheetByName("WorkOrder");
    if (sheetWO && sheetWO.getLastRow() > 1) {
      var dataWO = sheetWO.getRange(2, 1, sheetWO.getLastRow() - 1, Math.min(sheetWO.getLastColumn(), 11)).getDisplayValues();
      dataWO.forEach(function(row) {
        var pUrl = row[4] || "";
        var pId = row[5] || "";
        if (!pId && pUrl) pId = extractDriveFileId(pUrl);

        if (row[0] && row[0] !== "-") {
          result.workOrders.push({
            id: row[0],
            title: "SPKL " + (row[1] || "") + " (" + (row[2] || "") + ")",
            gi: row[1],
            bulan: row[2],
            pdfUrl: pUrl,
            pdfId: pId
          });
        }
      });
    }

    // 3. Surat Peringatan
    var sheetSP = getSuratPeringatanSheet(ss);
    if (sheetSP && sheetSP.getLastRow() > 1) {
      var dataSP = sheetSP.getRange(2, 1, sheetSP.getLastRow() - 1, Math.min(sheetSP.getLastColumn(), 13)).getDisplayValues();
      dataSP.forEach(function(row) {
        var namaLw = (row[1] || "").trim();
        var idSp = (row[0] || "").trim();
        if (!namaLw && !idSp) return;

        var spUrl = row[9] || "";
        var spId = row[10] || "";
        if (!spId && spUrl) spId = extractDriveFileId(spUrl);

        if (idSp && idSp !== "-") {
          result.suratPeringatan.push({
            id: idSp,
            title: (row[3] || "SP") + " - " + (namaLw || "-") + " (" + (row[4] || "") + ")",
            nama: namaLw,
            jenis: row[3],
            nomor: row[4],
            pdfUrl: spUrl,
            pdfId: spId
          });
        }
      });
    }

    return { success: true, data: result };
  } catch (err) {
    Logger.log("Error apiGetSystemDocumentsForApproval: " + err.toString());
    return { success: false, message: err.toString(), data: { beritaAcara: [], workOrders: [], suratPeringatan: [] } };
  }
}

/**
 * API: Mengambil byte PDF dokumen sistem untuk diimpor ke Digital Approval Studio
 * Mencegah error "Dokumen tidak ditemukan" dan mengekstrak berkas langsung dari Google Drive
 */
function apiGetSystemDocumentPdf(payload) {
  try {
    if (!payload || typeof payload !== "object") {
      throw new Error("Parameter tidak valid.");
    }

    var cat = payload.cat || "";
    var docId = (payload.docId || "").toString().trim();
    var pdfId = (payload.pdfId || "").toString().trim();
    var pdfUrl = (payload.pdfUrl || "").toString().trim();
    var targetFileId = "";

    // 1. Coba dari pdfId langsung
    if (pdfId && pdfId !== "#" && pdfId.length >= 15) {
      targetFileId = extractDriveFileId(pdfId) || pdfId;
    }

    // 2. Coba ekstrak dari pdfUrl
    if (!targetFileId && pdfUrl && pdfUrl !== "#") {
      targetFileId = extractDriveFileId(pdfUrl);
    }

    // 3. Jika belum ditemukan, lakukan pencarian di Sheet terkait berdasarkan docId
    if (!targetFileId && docId) {
      var ss = SpreadsheetApp.getActiveSpreadsheet();

      // Cek sheet BeritaAcara
      if (cat === "ba" || docId.indexOf("BA-") === 0) {
        var sheetBA = ss.getSheetByName("BeritaAcara");
        if (sheetBA && sheetBA.getLastRow() > 1) {
          var rowsBA = sheetBA.getDataRange().getDisplayValues();
          for (var b = 1; b < rowsBA.length; b++) {
            if (rowsBA[b][0] === docId) {
              targetFileId = extractDriveFileId(rowsBA[b][17]) || 
                             extractDriveFileId(rowsBA[b][16]) || 
                             extractDriveFileId(rowsBA[b][15]);
              if (!targetFileId) {
                for (var colB = rowsBA[b].length - 1; colB >= 0; colB--) {
                  var idExt = extractDriveFileId(rowsBA[b][colB]);
                  if (idExt) { targetFileId = idExt; break; }
                }
              }
              break;
            }
          }
        }

        // Cek Form_Responses
        if (!targetFileId && typeof GOOGLE_FORM_SPREADSHEET_ID !== "undefined" && GOOGLE_FORM_SPREADSHEET_ID) {
          try {
            var formSs = SpreadsheetApp.openById(GOOGLE_FORM_SPREADSHEET_ID);
            var sheetResp = formSs.getSheetByName("Form_Responses") || formSs.getSheets()[0];
            if (sheetResp && sheetResp.getLastRow() > 1) {
              var rowsR = sheetResp.getDataRange().getDisplayValues();
              for (var r = 1; r < rowsR.length; r++) {
                if (rowsR[r][0] === docId) {
                  for (var cR = rowsR[r].length - 1; cR >= 0; cR--) {
                    var exR = extractDriveFileId(rowsR[r][cR]);
                    if (exR) { targetFileId = exR; break; }
                  }
                  break;
                }
              }
            }
          } catch (eForm) {
            Logger.log("Error baca Form_Responses: " + eForm.toString());
          }
        }
      } else if (cat === "spkl" || docId.indexOf("WO-") === 0) {
        var sheetWO = ss.getSheetByName("WorkOrder");
        if (sheetWO && sheetWO.getLastRow() > 1) {
          var rowsWO = sheetWO.getDataRange().getDisplayValues();
          for (var w = 1; w < rowsWO.length; w++) {
            if (rowsWO[w][0] === docId) {
              targetFileId = extractDriveFileId(rowsWO[w][5]) || extractDriveFileId(rowsWO[w][4]);
              break;
            }
          }
        }
      } else if (cat === "sp" || docId.indexOf("SP-") === 0) {
        var sheetSP = getSuratPeringatanSheet(ss);
        if (sheetSP && sheetSP.getLastRow() > 1) {
          var rowsSP = sheetSP.getDataRange().getDisplayValues();
          for (var s = 1; s < rowsSP.length; s++) {
            if (rowsSP[s][0] === docId) {
              targetFileId = extractDriveFileId(rowsSP[s][10]) || extractDriveFileId(rowsSP[s][9]);
              break;
            }
          }
        }
      }
    }

    if (!targetFileId) {
      throw new Error("Berkas PDF untuk dokumen " + docId + " belum tersedia di Google Drive. Silakan buat/generate berkas terlebih dahulu atau unggah PDF secara manual.");
    }

    // 4. Ambil berkas dari Google Drive
    var file = DriveApp.getFileById(targetFileId);
    var mimeType = file.getMimeType();
    var blob;

    if (mimeType === MimeType.PDF) {
      blob = file.getBlob();
    } else {
      blob = file.getAs(MimeType.PDF);
    }

    var bytes = blob.getBytes();
    var pdfBase64 = Utilities.base64Encode(bytes);

    return {
      success: true,
      pdfBase64: pdfBase64,
      fileName: file.getName() || (docId + ".pdf"),
      fileId: targetFileId,
      docId: docId
    };
  } catch (err) {
    Logger.log("Error apiGetSystemDocumentPdf: " + err.toString());
    return { success: false, message: err.toString() };
  }
}

/**
 * API: Mengambil data Master Approver (PLN ES UP2 & UPT Semarang)
 * Mencakup MAHMUD FAUZI ISWORO (Asman Transmisi & GI UP2 PLN ES), TL UP2, Asman UPT, TL GI, KLW, dan Line Walker.
 * Otomatis sinkronisasi dengan sheet Data LW dan auto-seed sheet Master_Approver jika belum ada.
 */
function apiGetMasterApprovers() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_APPROVERS);
    if (!sheet || sheet.getLastRow() <= 1) {
      initMasterApproverSheet(ss);
      sheet = ss.getSheetByName(SHEET_APPROVERS);
    }

    var approvers = [];
    if (sheet && sheet.getLastRow() > 1) {
      var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.min(sheet.getLastColumn(), 9)).getDisplayValues();
      rows.forEach(function(r) {
        var nama = (r[5] || "").trim();
        if (nama !== "") {
          approvers.push({
            id: r[0] || ("APP-" + (approvers.length + 1)),
            unit: (r[1] || "UPT SEMARANG").trim(),
            kategori: (r[2] || "-").trim(),
            ultg: (r[3] || "-").trim(),
            gi: (r[4] || "-").trim(),
            nama: nama,
            jabatan: (r[6] || "-").trim(),
            nip: (r[7] || "-").trim(),
            keterangan: (r[8] || "").trim()
          });
        }
      });
    }

    // Sinkronisasi data dinamis dari Sheet "Data LW" jika ada
    try {
      var sheetLW = ss.getSheetByName("Data LW");
      if (sheetLW && sheetLW.getLastRow() > 1) {
        var dataLW = sheetLW.getRange(2, 1, sheetLW.getLastRow() - 1, Math.min(sheetLW.getLastColumn(), 8)).getDisplayValues();
        dataLW.forEach(function(lw) {
          var namaLW = (lw[0] || "").trim();
          var klw = (lw[1] || "").trim();
          var nik = (lw[2] || "").trim();
          var posisi = (lw[3] || "Line Walker").trim();
          var ultg = (lw[4] || "-").trim();
          var gi = (lw[5] || "-").trim();
          var tlGi = (lw[6] || "").trim();

          // Cek TL GI belum terdaftar
          if (tlGi && tlGi !== "-" && !approvers.some(function(a) { return a.nama.toLowerCase() === tlGi.toLowerCase(); })) {
            approvers.push({
              id: "APP-AUTO-TL-" + approvers.length,
              unit: "UPT SEMARANG",
              kategori: "TL GI",
              ultg: ultg,
              gi: gi,
              nama: tlGi,
              jabatan: "TL " + gi,
              nip: "-",
              keterangan: "Team Leader Gardu Induk"
            });
          }

          // Cek KLW belum terdaftar
          if (klw && klw !== "-" && !approvers.some(function(a) { return a.nama.toLowerCase() === klw.toLowerCase(); })) {
            approvers.push({
              id: "APP-AUTO-KLW-" + approvers.length,
              unit: "UPT SEMARANG",
              kategori: "KLW",
              ultg: ultg,
              gi: gi,
              nama: klw,
              jabatan: "Koordinator Line Walker " + ultg,
              nip: "-",
              keterangan: "Koordinator Line Walker"
            });
          }

          // Cek LW belum terdaftar
          if (namaLW && namaLW !== "-" && !approvers.some(function(a) { return a.nama.toLowerCase() === namaLW.toLowerCase(); })) {
            approvers.push({
              id: "APP-AUTO-LW-" + approvers.length,
              unit: "UPT SEMARANG",
              kategori: "LW",
              ultg: ultg,
              gi: gi,
              nama: namaLW,
              jabatan: posisi + " " + gi,
              nip: nik,
              keterangan: "Line Walker"
            });
          }
        });
      }
    } catch (eLW) {
      Logger.log("Info: Sinkronisasi Data LW ke Master Approvers dilewati: " + eLW.toString());
    }

    return { success: true, approvers: approvers };
  } catch (err) {
    Logger.log("Error apiGetMasterApprovers: " + err.toString());
    return { success: false, message: err.toString(), approvers: [] };
  }
}

/**
 * Inisialisasi Sheet Master_Approver jika belum ada
 */
function initMasterApproverSheet(ss) {
  try {
    var sheet = ss.getSheetByName(SHEET_APPROVERS);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_APPROVERS);
    }

    var header = ["ID", "UNIT", "KATEGORI", "ULTG", "GI", "NAMA", "JABATAN", "NIP_NIK", "KETERANGAN"];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(header);
      sheet.getRange(1, 1, 1, header.length).setFontWeight("bold").setBackground("#e2e8f0");
      sheet.setFrozenRows(1);
    }

    if (sheet.getLastRow() === 1) {
      var seedData = (typeof getMasterApproverDataset === "function") ? getMasterApproverDataset() : [
        ["APP-UP2-01", "PLN ES UP2", "ASMAN", "-", "-", "MAHMUD FAUZI ISWORO", "ASMAN TRANSMISI & GI UP2 PLN ES", "8206012Z", "Asman Transmisi & GI UP2"],
        ["APP-UP2-02", "PLN ES UP2", "TL", "-", "-", "AGUS SETIAWAN", "TL OPERASI & PEMELIHARAAN TRANSMISI UP2", "8809021Z", "Team Leader Op & Har UP2"],
        ["APP-UP2-03", "PLN ES UP2", "TL", "-", "-", "TRI WAHYUDI", "TL SCADA & TELEKOMUNIKASI UP2", "8911032Z", "Team Leader Scada & Tel UP2"],
        ["APP-UP2-04", "PLN ES UP2", "TL", "-", "-", "HENDRA WIJAYA", "TL K3L & KEAMANAN UP2", "9012044Z", "Team Leader K3L & Kam UP2"],
        ["APP-UP2-05", "PLN ES UP2", "MANAJER", "-", "-", "PRIHADI UTOMO", "MANAJER PLN ES UP2", "7905001Z", "Manajer UP2"],
        ["APP-UPT-01", "UPT SEMARANG", "MANAJER", "SEMUA", "-", "MANAJER UPT SEMARANG", "MANAJER UPT SEMARANG", "7804001Z", "Manajer UPT"],
        ["APP-UPT-02", "UPT SEMARANG", "ASMAN", "SEMUA", "-", "ASMAN JARINGAN UPT SEMARANG", "ASMAN JARINGAN UPT SEMARANG", "8305012Z", "Asman Jaringan"],
        ["APP-UPT-03", "UPT SEMARANG", "ASMAN", "SEMUA", "-", "ASMAN TRANSAKSI ENERGI & PEMELIHARAAN", "ASMAN TE & HAR UPT SEMARANG", "8407023Z", "Asman TE & Har"]
      ];

      if (seedData && seedData.length > 0) {
        sheet.getRange(2, 1, seedData.length, seedData[0].length).setValues(seedData);
      }
    }

    SpreadsheetApp.flush();
  } catch (e) {
    Logger.log("Error initMasterApproverSheet: " + e.toString());
  }
}

/**
 * Import Tower dari KML Content langsung (dipanggil dari web UI)
 * @param {string} kmlContent - Konten KML file
 * @param {boolean} replaceMode - true: clear semua data sebelum import
 * @return {object} status import
 */
function importTowerKMLContent(kmlContent, replaceMode) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetTower = ss.getSheetByName(SPREADSHEET_NAME_TOWER_SPAN);
    var sheetLW = ss.getSheetByName(SPREADSHEET_NAME_LW);
    
    if (!sheetTower) {
      throw new Error("Sheet TowerSpan belum ada. Jalankan setupDatabase() terlebih dahulu.");
    }
    
    // Clear data jika replace mode
    if (replaceMode && sheetTower.getLastRow() > 1) {
      sheetTower.deleteRows(2, sheetTower.getLastRow() - 1);
      SpreadsheetApp.flush();
    }
    
    // Parse KML untuk extract Placemark
    var placemarks = extractPlacemarksFromKML(kmlContent);
    
    if (placemarks.length === 0) {
      return { status: "warning", message: "Tidak ada data tower ditemukan di KML." };
    }
    
    // Buat map Line Walker untuk auto-linking
    var lwMap = {};
    if (sheetLW && sheetLW.getLastRow() > 1) {
      var lwData = sheetLW.getRange(2, 1, sheetLW.getLastRow() - 1, 8).getDisplayValues();
      for (var i = 0; i < lwData.length; i++) {
        var penghantar = (lwData[i][7] || "").trim().toUpperCase();
        if (penghantar && penghantar !== "-") {
          lwMap[penghantar] = {
            nama: lwData[i][0],
            klw: lwData[i][1],
            ultg: lwData[i][4],
            gi: lwData[i][5]
          };
        }
      }
    }
    
    var imported = [];
    var skipped = 0;
    
    for (var p = 0; p < placemarks.length; p++) {
      var pm = placemarks[p];
      
      // Skip jika tidak ada koordinat
      if (!pm.lat || !pm.lng) {
        skipped++;
        continue;
      }
      
      // Coba deteksi penghantar dari nama atau deskripsi
      var detectedPenghantar = detectPenghantarFromText(pm.name + " " + pm.description);
      var lwInfo = detectedPenghantar ? lwMap[detectedPenghantar] : null;
      
      var towerId = pm.name.replace(/[^A-Z0-9\.\-]/gi, "_").substring(0, 50) || ("TOWER_" + (p + 1));
      
      imported.push([
        towerId,
        pm.name,
        pm.lat,
        pm.lng,
        detectedPenghantar || "-",
        lwInfo ? lwInfo.ultg : "-",
        lwInfo ? lwInfo.gi : "-",
        lwInfo ? lwInfo.nama : "-",
        lwInfo ? lwInfo.klw : "-",
        pm.description || "-"
      ]);
    }
    
    if (imported.length > 0) {
      sheetTower.getRange(sheetTower.getLastRow() + 1, 1, imported.length, 10).setValues(imported);
      SpreadsheetApp.flush();
    }
    
    return {
      status: "success",
      message: "Berhasil import " + imported.length + " tower dari KML. Dilewati: " + skipped,
      imported: imported.length,
      skipped: skipped
    };
  } catch (err) {
    Logger.log("Error importTowerKMLContent: " + err.toString());
    return { status: "error", message: err.toString() };
  }
}

function extractPlacemarksFromKML(kmlContent) {
  var placemarks = [];
  
  // Regex sederhana untuk extract <Placemark>
  var placemarkRegex = /<Placemark[^>]*>([\s\S]*?)<\/Placemark>/gi;
  var matches = kmlContent.match(placemarkRegex);
  
  if (!matches) return placemarks;
  
  for (var i = 0; i < matches.length; i++) {
    var pm = matches[i];
    
    var name = extractTagContent(pm, "name");
    var description = extractTagContent(pm, "description");
    var coordinates = extractTagContent(pm, "coordinates");
    
    if (coordinates) {
      // Format KML coordinates: lng,lat,altitude (dipisah koma)
      var coords = coordinates.trim().split(",");
      if (coords.length >= 2) {
        placemarks.push({
          name: name || "Tower " + (i + 1),
          description: description || "",
          lng: parseFloat(coords[0]),
          lat: parseFloat(coords[1])
        });
      }
    }
  }
  
  return placemarks;
}

function extractTagContent(xml, tagName) {
  var regex = new RegExp("<" + tagName + "[^>]*>([\\s\\S]*?)<\\/" + tagName + ">", "i");
  var match = xml.match(regex);
  return match ? match[1].replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, "$1").trim() : "";
}

function detectPenghantarFromText(text) {
  if (!text) return null;
  
  var upperText = text.toUpperCase();
  
  // Pola umum nama penghantar di PLN
  var patterns = [
    /TRS\s+\d{2,3}kV\s+[A-Z\s\-]+/i,
    /SALURAN\s+[A-Z\s\-]+/i,
    /JALUR\s+[A-Z\s\-]+/i
  ];
  
  for (var i = 0; i < patterns.length; i++) {
    var match = text.match(patterns[i]);
    if (match) return match[0].trim().toUpperCase();
  }
  
  return null;
}

/**
 * API: Get Line Walker List untuk dropdown filter (termasuk list unik KLW dan GI)
 * Sesuai Standar ZettBOT: Index-Based Mapping (COL_LW)
 */
function getLineWalkerList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = (typeof SPREADSHEET_NAME_LW !== "undefined") ? SPREADSHEET_NAME_LW : "Data LW";
    var sheet = ss.getSheetByName(sheetName);
    
    var COL_LW = {
      NAMA: 0,
      KLW: 1,
      NIK: 2,
      POSISI: 3,
      ULTG: 4,
      GI: 5,
      TL_GI: 6,
      PENGHANTAR: 7
    };
    
    var result = [];
    var seenNama = {};
    var seenKLW = {};
    var seenGI = {};
    var klwList = [];
    var giList = [];
    
    if (sheet && sheet.getLastRow() > 1) {
      var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getDisplayValues();
      for (var i = 0; i < data.length; i++) {
        var nama = (data[i][COL_LW.NAMA] || "").trim();
        var klw = (data[i][COL_LW.KLW] || "").trim();
        var gi = (data[i][COL_LW.GI] || "").trim();
        
        if (klw && klw !== "-" && !seenKLW[klw.toUpperCase()]) {
          seenKLW[klw.toUpperCase()] = true;
          klwList.push(klw);
        }
        if (gi && gi !== "-" && !seenGI[gi.toUpperCase()]) {
          seenGI[gi.toUpperCase()] = true;
          giList.push(gi);
        }
        
        if (!nama || nama === "-" || seenNama[nama.toUpperCase()]) continue;
        seenNama[nama.toUpperCase()] = true;
        
        result.push({
          nama: nama,
          klw: klw || "-",
          nik: data[i][COL_LW.NIK] || "-",
          posisi: data[i][COL_LW.POSISI] || "Line Walker",
          ultg: data[i][COL_LW.ULTG] || "-",
          gi: gi || "-",
          tlGi: data[i][COL_LW.TL_GI] || "-",
          penghantar: data[i][COL_LW.PENGHANTAR] || "-"
        });
      }
    }
    
    // Scan TowerSpan sheet to include any additional KLW / GI / Line Walker
    var sheetTowerName = (typeof SPREADSHEET_NAME_TOWER_SPAN !== "undefined") ? SPREADSHEET_NAME_TOWER_SPAN : "TowerSpan";
    var sheetTower = ss.getSheetByName(sheetTowerName);
    if (sheetTower && sheetTower.getLastRow() > 1) {
      var tData = sheetTower.getRange(2, 1, sheetTower.getLastRow() - 1, 10).getDisplayValues();
      for (var j = 0; j < tData.length; j++) {
        var tKlw = (tData[j][8] || "").trim();
        var tGi = (tData[j][6] || "").trim();
        var tLw = (tData[j][7] || "").trim();
        if (tKlw && tKlw !== "-" && !seenKLW[tKlw.toUpperCase()]) {
          seenKLW[tKlw.toUpperCase()] = true;
          klwList.push(tKlw);
        }
        if (tGi && tGi !== "-" && !seenGI[tGi.toUpperCase()]) {
          seenGI[tGi.toUpperCase()] = true;
          giList.push(tGi);
        }
        if (tLw && tLw !== "-" && !seenNama[tLw.toUpperCase()]) {
          seenNama[tLw.toUpperCase()] = true;
          result.push({
            nama: tLw,
            klw: tKlw || "-",
            nik: "-",
            posisi: "Line Walker",
            ultg: (tData[j][5] || "").trim() || "-",
            gi: tGi || "-",
            tlGi: "-",
            penghantar: (tData[j][4] || "").trim() || "-"
          });
        }
      }
    }
    
    result.sort(function(a, b) { return a.nama.localeCompare(b.nama); });
    klwList.sort(function(a, b) { return a.localeCompare(b); });
    giList.sort(function(a, b) { return a.localeCompare(b); });
    
    return { status: "success", data: result, klwList: klwList, giList: giList };
  } catch (err) {
    Logger.log("Error getLineWalkerList: " + err.toString());
    return { status: "error", message: err.toString(), data: [], klwList: [], giList: [] };
  }
}

/**
 * API: Get Tower & Span Data berdasarkan Line Walker, KLW, GI, dan filter pencarian
 * Sesuai Standar ZettBOT: Index-Based Mapping (COL_TOWER) + Self-Healing Filter Dropdowns
 */
function getTowerDataByLineWalker(lwNama, searchQuery, klwFilter, giFilter, penghantarFilter) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = (typeof SPREADSHEET_NAME_TOWER_SPAN !== "undefined") ? SPREADSHEET_NAME_TOWER_SPAN : "TowerSpan";
    var sheetTower = ss.getSheetByName(sheetName);
    
    var COL_TOWER = {
      ID: 0,
      NAMA: 1,
      LAT: 2,
      LNG: 3,
      PENGHANTAR: 4,
      ULTG: 5,
      GI: 6,
      LINE_WALKER: 7,
      KLW: 8,
      DESKRIPSI: 9
    };

    var seenKLW = {};
    var seenGI = {};
    var seenLW = {};
    var seenPenghantar = {};
    var klwList = [];
    var giList = [];
    var lwList = [];
    var penghantarList = [];

    // 1. Ekstrak master Line Walker dari Data LW jika ada
    var lwSheetName = (typeof SPREADSHEET_NAME_LW !== "undefined") ? SPREADSHEET_NAME_LW : "Data LW";
    var sheetLW = ss.getSheetByName(lwSheetName);
    if (sheetLW && sheetLW.getLastRow() > 1) {
      var dLw = sheetLW.getRange(2, 1, sheetLW.getLastRow() - 1, 8).getDisplayValues();
      for (var k = 0; k < dLw.length; k++) {
        var mName = (dLw[k][0] || "").trim();
        var mKlw = (dLw[k][1] || "").trim();
        var mGi = (dLw[k][5] || "").trim();
        var mUltg = (dLw[k][4] || "").trim();
        var mPenghantar = (dLw[k][7] || "").trim();

        if (mKlw && mKlw !== "-" && !seenKLW[mKlw.toUpperCase()]) {
          seenKLW[mKlw.toUpperCase()] = true;
          klwList.push(mKlw);
        }
        if (mGi && mGi !== "-" && !seenGI[mGi.toUpperCase()]) {
          seenGI[mGi.toUpperCase()] = true;
          giList.push(mGi);
        }
        if (mPenghantar && mPenghantar !== "-" && !seenPenghantar[mPenghantar.toUpperCase()]) {
          seenPenghantar[mPenghantar.toUpperCase()] = true;
          penghantarList.push(mPenghantar);
        }
        if (mName && mName !== "-" && !seenLW[mName.toUpperCase()]) {
          seenLW[mName.toUpperCase()] = true;
          lwList.push({
            nama: mName,
            klw: mKlw || "-",
            gi: mGi || "-",
            ultg: mUltg || "-",
            penghantar: mPenghantar || "-"
          });
        }
      }
    }
    
    if (!sheetTower || sheetTower.getLastRow() <= 1) {
      klwList.sort(function(a, b) { return a.localeCompare(b); });
      giList.sort(function(a, b) { return a.localeCompare(b); });
      penghantarList.sort(function(a, b) { return a.localeCompare(b); });
      lwList.sort(function(a, b) { return a.nama.localeCompare(b.nama); });
      return { status: "success", data: [], total: 0, klwList: klwList, giList: giList, lwList: lwList, penghantarList: penghantarList };
    }
    
    var data = sheetTower.getRange(2, 1, sheetTower.getLastRow() - 1, 10).getDisplayValues();
    var filtered = [];
    var search = (searchQuery || "").trim().toLowerCase();
    var targetLW = (lwNama || "").trim().toLowerCase();
    var targetKLW = (klwFilter || "").trim().toLowerCase();
    var targetGI = (giFilter || "").trim().toLowerCase();
    var targetPenghantar = (penghantarFilter || "").trim().toLowerCase();
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var rawLW = (row[COL_TOWER.LINE_WALKER] || "").trim();
      var rawKLW = (row[COL_TOWER.KLW] || "").trim();
      var rawGI = (row[COL_TOWER.GI] || "").trim();
      var rawPenghantar = (row[COL_TOWER.PENGHANTAR] || "").trim();
      var rawUltg = (row[COL_TOWER.ULTG] || "").trim();

      if (rawKLW && rawKLW !== "-" && !seenKLW[rawKLW.toUpperCase()]) {
        seenKLW[rawKLW.toUpperCase()] = true;
        klwList.push(rawKLW);
      }
      if (rawGI && rawGI !== "-" && !seenGI[rawGI.toUpperCase()]) {
        seenGI[rawGI.toUpperCase()] = true;
        giList.push(rawGI);
      }
      if (rawPenghantar && rawPenghantar !== "-" && !seenPenghantar[rawPenghantar.toUpperCase()]) {
        seenPenghantar[rawPenghantar.toUpperCase()] = true;
        penghantarList.push(rawPenghantar);
      }
      if (rawLW && rawLW !== "-" && !seenLW[rawLW.toUpperCase()]) {
        seenLW[rawLW.toUpperCase()] = true;
        lwList.push({
          nama: rawLW,
          klw: rawKLW || "-",
          gi: rawGI || "-",
          ultg: rawUltg || "-",
          penghantar: rawPenghantar || "-"
        });
      }

      var towerLW = rawLW.toLowerCase();
      var towerKLW = rawKLW.toLowerCase();
      var towerGI = rawGI.toLowerCase();
      var towerPenghantar = rawPenghantar.toLowerCase();
      
      // Filter KLW
      if (targetKLW && targetKLW !== "all" && towerKLW !== targetKLW && towerKLW.indexOf(targetKLW) === -1) {
        continue;
      }
      
      // Filter GI
      if (targetGI && targetGI !== "all" && towerGI !== targetGI && towerGI.indexOf(targetGI) === -1) {
        continue;
      }

      // Filter Penghantar
      if (targetPenghantar && targetPenghantar !== "all" && towerPenghantar !== targetPenghantar && towerPenghantar.indexOf(targetPenghantar) === -1) {
        continue;
      }
      
      // Filter Line Walker
      if (targetLW && targetLW !== "all" && towerLW !== targetLW && towerLW.indexOf(targetLW) === -1) {
        continue;
      }
      
      // Filter Search: cek di semua kolom
      if (search !== "") {
        var rowText = (row[COL_TOWER.ID] + " " + row[COL_TOWER.NAMA] + " " + row[COL_TOWER.PENGHANTAR] + " " + row[COL_TOWER.ULTG] + " " + row[COL_TOWER.GI] + " " + row[COL_TOWER.LINE_WALKER] + " " + row[COL_TOWER.KLW] + " " + row[COL_TOWER.DESKRIPSI]).toLowerCase();
        if (rowText.indexOf(search) === -1) continue;
      }
      
      var latVal = parseFloat(row[COL_TOWER.LAT]) || 0;
      var lngVal = parseFloat(row[COL_TOWER.LNG]) || 0;
      
      filtered.push({
        towerId: row[COL_TOWER.ID] || ("T-" + (i + 1)),
        namaTower: row[COL_TOWER.NAMA] || ("Tower " + (i + 1)),
        lat: latVal,
        lng: lngVal,
        penghantar: rawPenghantar || "-",
        ultg: rawUltg || "-",
        gi: rawGI || "-",
        lineWalker: rawLW || "-",
        klw: rawKLW || "-",
        deskripsi: row[COL_TOWER.DESKRIPSI] || "-"
      });
    }
    
    klwList.sort(function(a, b) { return a.localeCompare(b); });
    giList.sort(function(a, b) { return a.localeCompare(b); });
    penghantarList.sort(function(a, b) { return a.localeCompare(b); });
    lwList.sort(function(a, b) { return a.nama.localeCompare(b.nama); });

    return {
      status: "success",
      data: filtered,
      total: filtered.length,
      klwList: klwList,
      giList: giList,
      lwList: lwList,
      penghantarList: penghantarList
    };
  } catch (err) {
    Logger.log("Error getTowerDataByLineWalker: " + err.toString());
    return { status: "error", message: err.toString(), data: [], total: 0, klwList: [], giList: [], lwList: [] };
  }
}


