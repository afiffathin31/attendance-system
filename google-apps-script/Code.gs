/**
 * ============================================================================
 * SISTEM PRESENSI RUTIN KAJIAN & EVALUASI PEMATERI (100% CLOUD & SERVERLESS)
 * Google Sheets + Google Forms + Ekstensi Meet/Zoom (Google Apps Script)
 * Versi: 3.3 (14 Departemen Baru, Auto Form Sync, Unified Dashboard)
 * ============================================================================
 */

// DAFTAR RESMI 14 DEPARTEMEN / DIVISI
const DEPARTEMEN_LIST = [
  "EBS",
  "Solution Architect",
  "Could Automation",
  "Research and innovation",
  "Infrastructur Analytic",
  "FAT",
  "Sales",
  "Cyber Cecurity",
  "PMO",
  "BOD",
  "Modern Application",
  "Data Analytics",
  "HCMC",
  "PKL/Magang"
];

// ----------------------------------------------------------------------------
// 1. MENU TOOLBAR OTOMATIS
// ----------------------------------------------------------------------------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🕌 Presensi Kajian")
    .addItem("⚙️ Setup / Perbarui Sistem Lengkap", "setupSistemAbsensiOtomatis")
    .addItem("👥 Buka Master Data Karyawan", "bukaTabDataKaryawan")
    .addItem("🧹 Kosongkan & Reset Tab Data Karyawan", "resetDanKosongkanTabDataKaryawan")
    .addItem("🔄 Sinkronkan Divisi di Log Absensi", "sinkronkanDepartemenLogDariDataKaryawan")
    .addItem("📊 Buat / Gambar Ulang Semua Grafik", "gambarUlangSemuaGrafik")
    .addItem("🔍 Buka Filter Kalender & Departemen", "bukaTabFilter")
    .addItem("🧹 Bersihkan Kolom Ganda Form Responses", "perbaikiKolomGandaFormResponses")
    .addItem("📋 Buat / Update Form Dropdown (14 Departemen)", "buatGoogleFormDropdownBaru")
    .addItem("🧹 Bersihkan Data Sampah / Uji Coba", "bersihkanDataLogKotor")
    .addItem("📅 Buka Tab Jadwal Kajian", "bukaJadwalKajian")
    .addItem("📋 Buka Google Form Jamaah", "bukaGoogleForm")
    .addToUi();
}

/**
 * Membuka / mengaktifkan tab '👥 Data Karyawan'
 */
function bukaTabDataKaryawan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("👥 Data Karyawan") || ss.getSheetByName("Data Karyawan");
  if (!sheet) {
    setupTabDataKaryawan();
    sheet = ss.getSheetByName("👥 Data Karyawan") || ss.getSheetByName("Data Karyawan");
  }
  if (sheet) ss.setActiveSheet(sheet);
}

/**
 * Membuka / mengaktifkan tab '🔍 Cari & Filter Data'
 */
function bukaTabFilter() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("🔍 Cari & Filter Data");
  if (!sheet) {
    setupTabFilterDanPencarian();
    sheet = ss.getSheetByName("🔍 Cari & Filter Data");
  }
  if (sheet) ss.setActiveSheet(sheet);
}

/**
 * Merapikan kolom ganda 'Departemen / Divisi' pada tab 'Form Responses 1'
 * Menggabungkan isi kolom dan menghapus kolom duplikat yang kosong.
 */
function perbaikiKolomGandaFormResponses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  let fixedCount = 0;

  for (let sheet of allSheets) {
    const sName = sheet.getName().toLowerCase();
    if (sName.startsWith("form response") || sName.startsWith("form_response") || sName.startsWith("tanggapan formulir")) {
      const lastCol = sheet.getLastColumn();
      const lastRow = sheet.getLastRow();
      if (lastCol < 3 || lastRow < 1) continue;

      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      const deptCols = [];

      for (let c = 0; c < headers.length; c++) {
        const h = String(headers[c] || '').toLowerCase().trim();
        if (h.includes("departemen") || h.includes("divisi")) {
          deptCols.push(c + 1); // 1-based index
        }
      }

      // Jika ada lebih dari 1 kolom departemen (kolom ganda)
      if (deptCols.length > 1) {
        const primaryCol = deptCols[0]; // Kolom C

        if (lastRow > 1) {
          for (let i = 1; i < deptCols.length; i++) {
            const extraCol = deptCols[i];
            const primaryRange = sheet.getRange(2, primaryCol, lastRow - 1, 1);
            const extraRange = sheet.getRange(2, extraCol, lastRow - 1, 1);
            const pVals = primaryRange.getValues();
            const eVals = extraRange.getValues();

            for (let r = 0; r < pVals.length; r++) {
              const valP = String(pVals[r][0] || '').trim();
              const valE = String(eVals[r][0] || '').trim();
              if (!valP && valE) {
                pVals[r][0] = valE;
              }
            }
            primaryRange.setValues(pVals);
          }
        }

        // Hapus kolom duplikat dari belakang agar indeks tidak bergeser
        for (let i = deptCols.length - 1; i >= 1; i--) {
          sheet.deleteColumn(deptCols[i]);
          fixedCount++;
        }

        sheet.getRange(1, primaryCol).setValue("Departemen / Divisi");
      }
    }
  }

  // Juga bersihkan Google Form jika ada item ganda
  try {
    let formUrl = PropertiesService.getScriptProperties().getProperty("GOOGLE_FORM_URL");
    let formEditUrl = PropertiesService.getScriptProperties().getProperty("GOOGLE_FORM_EDIT_URL");
    let form = null;
    if (formEditUrl) { try { form = FormApp.openByUrl(formEditUrl); } catch(e) {} }
    if (!form && formUrl) { try { form = FormApp.openByUrl(formUrl); } catch(e) {} }
    if (!form) { try { const bUrl = ss.getFormUrl(); if (bUrl) form = FormApp.openByUrl(bUrl); } catch(e) {} }

    if (form) {
      const items = form.getItems();
      const deptFormItems = [];
      for (let it of items) {
        const t = it.getTitle().toLowerCase();
        if (t.includes("departemen") || t.includes("divisi")) {
          deptFormItems.push(it);
        }
      }
      while (deptFormItems.length > 1) {
        const extra = deptFormItems.pop();
        form.deleteItem(extra.getIndex());
      }
    }
  } catch(e) {}

  ss.toast("Kolom ganda pada Form Responses berhasil dirapikan menjadi 1 kolom bersih!", "✅ Berhasil Dirapikan", 6);
  return fixedCount;
}

/**
 * Membuat tab khusus '🔍 Cari & Filter Data' dengan:
 * 1. Filter Tanggal langsung dari KALENDER VISUAL (klik 2x sel)
 * 2. Filter Departemen langsung dari DROPDOWN 14 Departemen Resmi
 * 3. Hasil penyaringan presensi real-time otomatis
 */
function setupTabFilterDanPencarian() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("🔍 Cari & Filter Data");
  
  if (!sheet) {
    sheet = ss.insertSheet("🔍 Cari & Filter Data", 2);
  } else {
    try { sheet.getDataRange().breakApart(); } catch(e) {}
    sheet.clear();
  }
  sheet.setTabColor("#0284c7");

  // 1. Header Title
  sheet.getRange("A1:J1").merge()
    .setValue("🔍 PENCARIAN & FILTER PRESENSI KAJIAN (REAL-TIME)")
    .setBackground("#0f172a")
    .setFontColor("#38bdf8")
    .setFontSize(13)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(1, 40);

  // 2. Kotak Filter Interaktif (Baris 3)
  sheet.getRange("A3").setValue("📅 Tanggal (Kalender):").setFontWeight("bold").setHorizontalAlignment("right");
  
  const ruleDate = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(true)
    .build();
  sheet.getRange("B3")
    .setDataValidation(ruleDate)
    .setValue(new Date())
    .setNumberFormat("yyyy-mm-dd")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground("#fef08a")
    .setBorder(true, true, true, true, true, true, "#ca8a04", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange("C3").setValue("🏢 Pilih Departemen:").setFontWeight("bold").setHorizontalAlignment("right");
  
  const deptFilterOptions = ["Semua Departemen", ...DEPARTEMEN_LIST];
  const ruleDept = SpreadsheetApp.newDataValidation()
    .requireValueInList(deptFilterOptions, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange("D3")
    .setDataValidation(ruleDept)
    .setValue("Semua Departemen")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground("#dbeafe")
    .setBorder(true, true, true, true, true, true, "#2563eb", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange("E3").setValue("Hasil Filter:").setFontWeight("bold").setHorizontalAlignment("right");
  sheet.getRange("F3").setFormula(`="✅ Ditemukan: " & IFERROR(COUNTA(A7:A) - IF(ISNUMBER(MATCH("ℹ️*"; A7:A; 0)); 1; 0); 0) & " Jamaah"`)
    .setFontWeight("bold")
    .setFontColor("#065f46")
    .setBackground("#d1fae5")
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true, "#059669", SpreadsheetApp.BorderStyle.SOLID);

  // Petunjuk Pemakaian (Baris 4)
  sheet.getRange("B4").setValue("👆 Klik 2x sel B3 (Popup Kalender)")
    .setFontSize(9).setFontStyle("italic").setFontColor("#64748b").setHorizontalAlignment("center");
  sheet.getRange("D4").setValue("👆 Klik tanda panah D3 (Pilih Divisi)")
    .setFontSize(9).setFontStyle("italic").setFontColor("#64748b").setHorizontalAlignment("center");
  sheet.getRange("F4").setValue("Tabel Otomatis Terupdate")
    .setFontSize(9).setFontStyle("italic").setFontColor("#64748b").setHorizontalAlignment("center");

  // 3. Header Tabel Hasil (Baris 6)
  const headers = [
    "Timestamp", "Tanggal", "Jam Masuk", "Nama Lengkap", 
    "Departemen", "Metode", "Tema / Judul Kajian", "Nama Pemateri", 
    "Status Kehadiran", "Keterlambatan (Menit)"
  ];
  sheet.getRange("A6:J6").setValues([headers])
    .setBackground("#1e293b")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(6, 32);

  // 4. Formula Filter Otomatis (Baris 7)
  sheet.getRange("A7").setFormula(
    `=IFERROR(FILTER('Log Absensi'!A2:J; (IF(ISBLANK($B$3); 1; 'Log Absensi'!B2:B = TEXT($B$3; "yyyy-mm-dd"))) * (IF(OR(ISBLANK($D$3); $D$3="Semua Departemen"); 1; 'Log Absensi'!E2:E = $D$3))); "ℹ️ Tidak ada data presensi yang sesuai dengan tanggal & departemen yang Anda pilih.")`
  );

  // Format Kolom
  sheet.getRange("A7:A").setNumberFormat("@");
  sheet.getRange("B7:B").setNumberFormat("yyyy-mm-dd").setHorizontalAlignment("center");
  sheet.getRange("C7:C").setNumberFormat("hh:mm:ss").setHorizontalAlignment("center");
  sheet.getRange("I7:J").setHorizontalAlignment("center");

  // Conditional Formatting
  const ruleOnTime = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("TEPAT WAKTU")
    .setBackground("#d1fae5")
    .setFontColor("#065f46")
    .setBold(true)
    .setRanges([sheet.getRange("I7:I1000")])
    .build();

  const ruleLate = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("TERLAMBAT")
    .setBackground("#fee2e2")
    .setFontColor("#991b1b")
    .setBold(true)
    .setRanges([sheet.getRange("I7:I1000")])
    .build();

  sheet.setConditionalFormatRules([ruleOnTime, ruleLate]);

  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 90);
  sheet.setColumnWidth(4, 180);
  sheet.setColumnWidth(5, 140);
  sheet.setColumnWidth(6, 130);
  sheet.setColumnWidth(7, 200);
  sheet.setColumnWidth(8, 180);
  sheet.setColumnWidth(9, 130);
  sheet.setColumnWidth(10, 130);

  sheet.setFrozenRows(6);
}

/**
 * Membuat atau memperbarui tab khusus '👥 Data Karyawan' dengan:
 * 1. Kolom Input Nama Lengkap & Dropdown 14 Departemen Resmi (Kolom B & C - DIBIARKAN KOSONG untuk diisi pengguna)
 * 2. Monitoring Performa Individu Real-time (Total Hadir, Tepat Waktu, Terlambat, % On-Time, Status)
 * 3. Tabel Ringkasan Monitoring Performa 14 Divisi Real-time (Kolom J - N)
 */
function setupTabDataKaryawan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("👥 Data Karyawan");
  if (!sheet) sheet = ss.getSheetByName("Data Karyawan");

  if (!sheet) {
    sheet = ss.insertSheet("👥 Data Karyawan", 2);
  }
  sheet.setTabColor("#10b981");

  // Pastikan jumlah baris minimal 200 baris
  if (sheet.getMaxRows() < 200) {
    sheet.insertRowsAfter(sheet.getMaxRows(), 200 - sheet.getMaxRows());
  }

  // Header Banner Atas (Baris 1)
  sheet.getRange("A1:H1").merge()
    .setValue("👥 MASTER DATA KARYAWAN & MONITORING PERFORMA INDIVIDU")
    .setBackground("#0f172a")
    .setFontColor("#38bdf8")
    .setFontSize(12)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  sheet.getRange("J1:N1").merge()
    .setValue("🏢 MONITORING PERFORMA DIVISI (REAL-TIME)")
    .setBackground("#065f46")
    .setFontColor("#ffffff")
    .setFontSize(12)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(1, 38);

  // Header Tabel Individu (Baris 2)
  const headersIndividu = [
    "No", "Nama Lengkap Karyawan", "Departemen / Divisi", 
    "Total Hadir", "Tepat Waktu", "Terlambat", 
    "% On-Time", "Status Performa"
  ];
  sheet.getRange("A2:H2").setValues([headersIndividu])
    .setBackground("#1e293b")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  // Header Tabel Ringkasan Divisi (Baris 2)
  const headersDivisi = [
    "Departemen (14 Divisi)", "Jml Karyawan", "Total Hadir", "Tepat Waktu", "% On-Time Divisi"
  ];
  sheet.getRange("J2:N2").setValues([headersDivisi])
    .setBackground("#047857")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(2, 32);

  // 1. Validasi Dropdown Departemen (Kolom C) - 14 Departemen Resmi
  const ruleDept = SpreadsheetApp.newDataValidation()
    .requireValueInList(DEPARTEMEN_LIST, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange("C3:C200").setDataValidation(ruleDept);

  // Format Text untuk Kolom Nama Lengkap (B3:B200)
  sheet.getRange("B3:B200").setNumberFormat("@");

  // 2. Set Formula Otomatis untuk Tabel Individu (Baris 3 s/d 200)
  // Formula hanya aktif jika Kolom B diisi nama karyawan (ISBLANK = "")
  const indFormulas = [];
  for (let r = 3; r <= 200; r++) {
    indFormulas.push([
      `=IF(ISBLANK(B${r}); ""; ROW()-2)`,
      `=IF(ISBLANK(B${r}); ""; COUNTIFS('Log Absensi'!D:D; B${r}))`,
      `=IF(ISBLANK(B${r}); ""; COUNTIFS('Log Absensi'!D:D; B${r}; 'Log Absensi'!I:I; "TEPAT WAKTU"))`,
      `=IF(ISBLANK(B${r}); ""; COUNTIFS('Log Absensi'!D:D; B${r}; 'Log Absensi'!I:I; "TERLAMBAT"))`,
      `=IF(ISBLANK(B${r}); ""; IFERROR(IF(D${r}>0; E${r}/D${r}; 0); 0))`,
      `=IF(ISBLANK(B${r}); ""; IF(D${r}=0; "Belum Hadir"; IF(G${r}*100>=85; "⭐ Sangat Baik"; IF(G${r}*100>=70; "👍 Baik"; "⚠️ Perlu Ditingkatkan"))))`
    ]);
  }
  sheet.getRange("A3:A200").setFormulas(indFormulas.map(f => [f[0]]));
  sheet.getRange("D3:H200").setFormulas(indFormulas.map(f => f.slice(1)));

  // Format Angka & Align Individu
  sheet.getRange("A3:A200").setHorizontalAlignment("center");
  sheet.getRange("D3:F200").setHorizontalAlignment("center");
  sheet.getRange("G3:G200").setNumberFormat("0.0%").setHorizontalAlignment("center").setFontWeight("bold");
  sheet.getRange("H3:H200").setHorizontalAlignment("center");

  // 3. Set Formula Otomatis untuk Tabel Ringkasan Divisi (Baris 3 s/d 16: 14 Departemen)
  const divRows = [];
  for (let i = 0; i < DEPARTEMEN_LIST.length; i++) {
    const deptName = DEPARTEMEN_LIST[i];
    const r = 3 + i;
    divRows.push([
      deptName,
      `=COUNTIF($C$3:$C$200; "${deptName}")`,
      `=COUNTIFS('Log Absensi'!E:E; "${deptName}")`,
      `=COUNTIFS('Log Absensi'!E:E; "${deptName}"; 'Log Absensi'!I:I; "TEPAT WAKTU")`,
      `=IFERROR(IF(L${r}>0; M${r}/L${r}; 0); 0)`
    ]);
  }
  sheet.getRange("J3:N16").setValues(divRows.map(row => [row[0], "", "", "", ""]));
  sheet.getRange("K3:N16").setFormulas(divRows.map(row => row.slice(1)));
  sheet.getRange("J3:J16").setFontWeight("bold");
  sheet.getRange("K3:M16").setHorizontalAlignment("center");
  sheet.getRange("N3:N16").setNumberFormat("0.0%").setHorizontalAlignment("center").setFontWeight("bold");

  // Total Keseluruhan Divisi (Baris 17)
  sheet.getRange("J17").setValue("TOTAL KESELURUHAN").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("K17").setFormula("=SUM(K3:K16)").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("L17").setFormula("=SUM(L3:L16)").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("M17").setFormula("=SUM(M3:M16)").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("N17").setFormula("=IFERROR(IF(L17>0; M17/L17; 0); 0)").setNumberFormat("0.0%").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("J17:N17").setBackground("#f1f5f9").setBorder(true, true, true, true, true, true, "#94a3b8", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange("J2:N16").setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);

  // CATATAN: Baris data karyawan (Kolom B & C baris 3 ke bawah) SENGAJA DIBIARKAN KOSONG
  // agar pengguna dapat mengisi nama dan memilih divisi melalui dropdown secara mandiri.

  // 4. Atur Lebar Kolom
  sheet.setColumnWidth(1, 45);   // No
  sheet.setColumnWidth(2, 230);  // Nama Karyawan
  sheet.setColumnWidth(3, 190);  // Departemen Dropdown
  sheet.setColumnWidth(4, 95);   // Total Hadir
  sheet.setColumnWidth(5, 95);   // Tepat Waktu
  sheet.setColumnWidth(6, 95);   // Terlambat
  sheet.setColumnWidth(7, 100);  // % On-Time
  sheet.setColumnWidth(8, 140);  // Status Performa
  sheet.setColumnWidth(9, 25);   // Pemisah
  sheet.setColumnWidth(10, 200); // Departemen Divisi
  sheet.setColumnWidth(11, 100); // Jml Karyawan
  sheet.setColumnWidth(12, 95);  // Total Hadir
  sheet.setColumnWidth(13, 95);  // Tepat Waktu
  sheet.setColumnWidth(14, 120); // % On-Time Divisi

  sheet.setFrozenRows(2);
  sheet.getRange("A2:H200").setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
}

/**
 * Mengosongkan seluruh kolom Nama (B) dan Divisi (C) pada tab 'Data Karyawan',
 * serta memperbarui formula Status Performa agar tidak terjadi #ERROR!.
 */
function resetDanKosongkanTabDataKaryawan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Data Karyawan") || ss.getSheetByName("👥 Data Karyawan");
  if (!sheet) {
    setupTabDataKaryawan();
    sheet = ss.getSheetByName("Data Karyawan") || ss.getSheetByName("👥 Data Karyawan");
  }
  if (sheet) {
    const maxR = Math.max(sheet.getLastRow(), 200);
    // Hapus seluruh isi Kolom B & C dari baris 3 ke bawah
    sheet.getRange(3, 2, maxR - 2, 2).clearContent();
    setupTabDataKaryawan();
    ss.setActiveSheet(sheet);
    SpreadsheetApp.getActiveSpreadsheet().toast("Tab Data Karyawan berhasil dikosongkan dan rumus diperbaiki!", "✅ Berhasil", 6);
  }
}

/**
 * Alias untuk menu kosongkan data
 */
function kosongkanDataKaryawan() {
  resetDanKosongkanTabDataKaryawan();
}

/**
 * Mencari departemen terdaftar berdasarkan nama karyawan dari tab '👥 Data Karyawan'
 * Menghubungkan peserta Zoom / Google Meet / Form langsung dengan divisi resminya
 * 
 * @param {string} name - Nama peserta yang masuk dari Zoom / Google Meet / Form
 * @return {string|null} - Nama divisi resmi yang terdaftar, atau null
 */
function lookupRegisteredDepartment(name) {
  if (!name) return null;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetKaryawan = ss.getSheetByName("👥 Data Karyawan") || ss.getSheetByName("Data Karyawan");
    if (!sheetKaryawan) return null;

    const lastRow = sheetKaryawan.getLastRow();
    if (lastRow < 3) return null;

    const data = sheetKaryawan.getRange(3, 2, lastRow - 2, 2).getValues(); // Col B (Nama), Col C (Dept)
    let cleanTarget = String(name).trim().toLowerCase();

    // 1. Bersihkan badge role Zoom / Meet jika ada
    cleanTarget = cleanTarget.replace(/\((anda|you|host|penyelenggara rapat|penyelenggara|co-host|host,\s*me|me|guest|tamu)\)/gi, "").trim();
    // Hilangkan spasi berlebih
    cleanTarget = cleanTarget.replace(/\s+/g, ' ');

    if (!cleanTarget) return null;

    // Tahap 1: Pencarian Persis (Exact match - case insensitive)
    for (let row of data) {
      const empName = String(row[0] || '').trim().toLowerCase().replace(/\s+/g, ' ');
      const empDept = String(row[1] || '').trim();
      if (empName && empName === cleanTarget && empDept && empDept !== "Umum") {
        return empDept;
      }
    }

    // Tahap 2: Sub-string / Partial match (jika panjang nama >= 4 karakter)
    // Membantu jika di Zoom peserta menggunakan imbuhan gelar atau nama panggilan
    if (cleanTarget.length >= 4) {
      for (let row of data) {
        const empName = String(row[0] || '').trim().toLowerCase().replace(/\s+/g, ' ');
        const empDept = String(row[1] || '').trim();
        if (empName && empDept && empDept !== "Umum" && empName.length >= 4) {
          if (cleanTarget.includes(empName) || empName.includes(cleanTarget)) {
            return empDept;
          }
        }
      }
    }
  } catch(e) {
    Logger.log("Error in lookupRegisteredDepartment: " + e.message);
  }
  return null;
}

/**
 * Menyinkronkan ulang kolom Departemen di tab 'Log Absensi' 
 * berdasarkan Master Data Karyawan yang telah diisi pengguna.
 * Sangat bermanfaat setelah pengguna memasukkan daftar nama & divisi baru.
 */
function sinkronkanDepartemenLogDariDataKaryawan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetLog = ss.getSheetByName("Log Absensi");
  const sheetKaryawan = ss.getSheetByName("👥 Data Karyawan") || ss.getSheetByName("Data Karyawan");

  if (!sheetLog || !sheetKaryawan) {
    SpreadsheetApp.getActiveSpreadsheet().toast("Sheet 'Log Absensi' atau 'Data Karyawan' belum siap.", "Perhatian", 5);
    return 0;
  }

  const logLastRow = sheetLog.getLastRow();
  if (logLastRow < 2) {
    SpreadsheetApp.getActiveSpreadsheet().toast("Belum ada data di Log Absensi.", "Info", 4);
    return 0;
  }

  // Bangun peta master: Nama -> Departemen
  const masterMap = new Map();
  const karLastRow = sheetKaryawan.getLastRow();
  if (karLastRow >= 3) {
    const karData = sheetKaryawan.getRange(3, 2, karLastRow - 2, 2).getValues();
    karData.forEach(r => {
      const n = String(r[0] || '').trim().toLowerCase().replace(/\s+/g, ' ');
      const d = String(r[1] || '').trim();
      if (n && d && d !== "Umum") {
        masterMap.set(n, d);
      }
    });
  }

  if (masterMap.size === 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast("Belum ada data karyawan di tab '👥 Data Karyawan'. Silakan isi terlebih dahulu.", "Info", 6);
    return 0;
  }

  // Baca Kolom D (Nama Jamaah) dan Kolom E (Departemen) dari Log Absensi
  const logRange = sheetLog.getRange(2, 4, logLastRow - 1, 2);
  const logValues = logRange.getValues();
  let updatedCount = 0;

  for (let i = 0; i < logValues.length; i++) {
    const rawName = String(logValues[i][0] || '').trim();
    const currentDept = String(logValues[i][1] || '').trim();
    const cleanName = rawName.toLowerCase().replace(/\s+/g, ' ');

    if (!cleanName) continue;

    // Jika departemen kosong, "Umum", atau belum terdaftar resmi
    if (!currentDept || currentDept === "Umum" || !DEPARTEMEN_LIST.includes(currentDept)) {
      let matchedDept = masterMap.get(cleanName);
      if (!matchedDept && cleanName.length >= 4) {
        // Coba partial match
        for (let [kName, kDept] of masterMap.entries()) {
          if (cleanName.includes(kName) || kName.includes(cleanName)) {
            matchedDept = kDept;
            break;
          }
        }
      }
      if (matchedDept && matchedDept !== currentDept) {
        logValues[i][1] = matchedDept;
        updatedCount++;
      }
    }
  }

  if (updatedCount > 0) {
    logRange.setValues(logValues);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Berhasil mengidentifikasi & menyinkronkan ${updatedCount} baris departemen di Log Absensi!`,
      "✅ Sinkronisasi Berhasil",
      6
    );
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Semua departemen di Log Absensi sudah sinkron dengan Data Karyawan.",
      "ℹ️ Sudah Sinkron",
      5
    );
  }

  return updatedCount;
}

/**
 * Trigger otomatis ketika pengguna mengedit sel secara langsung di spreadsheet
 * Jika mengedit/mengetik Nama Jamaah di 'Log Absensi', otomatis mengisi departemennya
 */
function onEdit(e) {
  if (!e || !e.range) return;
  try {
    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();
    const row = e.range.getRow();
    const col = e.range.getColumn();

    // Jika mengedit kolom 'Nama Jamaah' (Kolom 4 / D) di 'Log Absensi' pada baris >= 2
    if (sheetName === "Log Absensi" && col === 4 && row >= 2) {
      const nama = String(e.value || sheet.getRange(row, 4).getValue() || '').trim();
      const currentDept = String(sheet.getRange(row, 5).getValue() || '').trim();
      if (nama && (!currentDept || currentDept === "Umum")) {
        const foundDept = lookupRegisteredDepartment(nama);
        if (foundDept) {
          sheet.getRange(row, 5).setValue(foundDept);
        }
      }
    }
  } catch(err) {
    // Non-blocking trigger
  }
}

/**
 * Menu trigger untuk memasang filter departemen & tanggal pada tab raw data
 */
function menuPasangFilterData() {
  const count = pasangFilterDepartemenDanTanggal();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Filter aktif di ${count} tab (Log Absensi & Form Responses). Klik tanda panah filter di header Departemen & Tanggal.`,
    "🔍 Filter Berhasil Dipasang",
    6
  );
}

/**
 * Memasang filter otomatis (Google Sheets Native Filter) pada tab 'Log Absensi' 
 * dan tab 'Form Responses 1' (Tanggapan Formulir) sehingga pengguna dapat memfilter 
 * Departemen dan Tanggal dengan 1 klik pada header kolom.
 */
function pasangFilterDepartemenDanTanggal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  let countUpdated = 0;

  for (let sheet of allSheets) {
    const sName = sheet.getName().trim();
    const isLog = sName === "Log Absensi";
    const isForm = sName.toLowerCase().startsWith("form response") || 
                  sName.toLowerCase().startsWith("tanggapan formulir");

    if (isLog || isForm) {
      try {
        const lastRow = Math.max(sheet.getLastRow(), 2);
        const lastCol = sheet.getLastColumn();
        if (lastCol < 1) continue;

        // Lepas filter lama jika ada agar range diperbarui mencakup seluruh baris data terbaru
        const existingFilter = sheet.getFilter();
        if (existingFilter) {
          existingFilter.remove();
        }

        // Pasang filter baru mencakup baris 1 sampai baris data terakhir dan kolom terakhir
        sheet.getRange(1, 1, lastRow, lastCol).createFilter();
        sheet.setFrozenRows(1);
        countUpdated++;
      } catch (err) {
        Logger.log(`Warning pasang filter di sheet ${sName}: ${err.message}`);
      }
    }
  }

  return countUpdated;
}

function bukaJadwalKajian() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Jadwal Kajian");
  if (sheet) ss.setActiveSheet(sheet);
}

function bukaGoogleForm() {
  const formUrl = PropertiesService.getScriptProperties().getProperty("GOOGLE_FORM_URL");
  if (formUrl) {
    const html = `<script>window.open("${formUrl}", "_blank");google.script.host.close();</script>`;
    SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(200).setHeight(50), "Membuka Form...");
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast("Jalankan setup atau buat form terlebih dahulu.", "Perhatian", 5);
  }
}

/**
 * Membuat atau memperbarui Google Form dengan tipe pertanyaan Dropdown (List) berisi 14 departemen resmi
 */
function buatGoogleFormDropdownBaru() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    let formUrl = PropertiesService.getScriptProperties().getProperty("GOOGLE_FORM_URL");
    let formEditUrl = PropertiesService.getScriptProperties().getProperty("GOOGLE_FORM_EDIT_URL");
    let form = null;

    if (formEditUrl) {
      try { form = FormApp.openByUrl(formEditUrl); } catch(e) {}
    }
    if (!form && formUrl) {
      try { form = FormApp.openByUrl(formUrl); } catch(e) {}
    }
    if (!form) {
      try {
        const boundUrl = ss.getFormUrl();
        if (boundUrl) form = FormApp.openByUrl(boundUrl);
      } catch(e) {}
    }

    if (!form) {
      form = FormApp.create("Presensi Kajian Rutin Jamaah");
      form.setDescription("Formulir absensi mandiri kajian rutin jamaah. Silakan masukkan nama lengkap dan pilih departemen/divisi Anda.");
      form.setAllowResponseEdits(false);
      form.setCollectEmail(false);
      form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    }

    // Periksa apakah item sudah ada; pastikan pertanyaan Departemen bertipe DROPDOWN (LIST)
    const existingItems = form.getItems();
    let hasNameItem = false;
    let deptItem = null;

    for (let item of existingItems) {
      const title = item.getTitle().toLowerCase();
      if (title.includes("nama")) {
        hasNameItem = true;
      }
      if (title.includes("departemen") || title.includes("divisi")) {
        deptItem = item;
      }
    }

    // Jika belum ada pertanyaan nama
    if (!hasNameItem) {
      form.addTextItem().setTitle("Nama Lengkap").setHelpText("Masukkan nama lengkap Anda").setRequired(true);
    }

    // Pasang / Ganti item departemen menjadi DROPDOWN (LIST)
    if (deptItem) {
      if (deptItem.getType() === FormApp.ItemType.LIST) {
        deptItem.asListItem().setChoiceValues(DEPARTEMEN_LIST);
      } else {
        // Jika sebelumnya Multiple Choice, ganti menjadi Dropdown List
        const idx = deptItem.getIndex();
        form.deleteItem(idx);
        const newDeptList = form.addListItem();
        newDeptList.setTitle("Departemen / Divisi").setHelpText("Pilih departemen dari dropdown").setRequired(true);
        newDeptList.setChoiceValues(DEPARTEMEN_LIST);
      }
    } else {
      const newDeptList = form.addListItem();
      newDeptList.setTitle("Departemen / Divisi").setHelpText("Pilih departemen dari dropdown").setRequired(true);
      newDeptList.setChoiceValues(DEPARTEMEN_LIST);
    }

    formUrl = form.getPublishedUrl();
    formEditUrl = form.getEditUrl();

    PropertiesService.getScriptProperties().setProperty("GOOGLE_FORM_URL", formUrl);
    PropertiesService.getScriptProperties().setProperty("GOOGLE_FORM_EDIT_URL", formEditUrl);

    // Update kotak Google Form di Dashboard
    // Update kotak Google Form di Dashboard (D3:E6)
    const sheetDash = ss.getSheetByName("Dashboard & Analitik");
    if (sheetDash) {
      sheetDash.getRange("D4").setValue("Form Jamaah:");
      sheetDash.getRange("E4").setFormula(`=HYPERLINK("${formUrl}"; "👉 Klik Isi Absensi")`).setFontWeight("bold");
      sheetDash.getRange("D5").setValue("Link URL:");
      sheetDash.getRange("E5").setValue(formUrl);
      sheetDash.getRange("D6").setValue("Edit Form:");
      sheetDash.getRange("E6").setFormula(`=HYPERLINK("${formEditUrl || formUrl}"; "⚙️ Pengaturan Soal")`);
    }

    // Pastikan trigger onGoogleFormSubmit aktif
    const triggers = ScriptApp.getProjectTriggers();
    const hasTrigger = triggers.some(t => t.getHandlerFunction() === "onGoogleFormSubmit");
    if (!hasTrigger) {
      ScriptApp.newTrigger("onGoogleFormSubmit").forSpreadsheet(ss).onFormSubmit().create();
    }

    // Tampilkan pop-up dialog siap pakai
    const htmlOutput = HtmlService.createHtmlOutput(`
      <div style="font-family: Arial, sans-serif; padding: 15px; color: #1e293b; line-height: 1.5;">
        <h3 style="color: #059669; margin-top: 0;">✅ Google Form Dropdown Siap Digunakan!</h3>
        <p>Pertanyaan departemen telah disetel menjadi <b>Dropdown List</b> berisi 14 departemen resmi.</p>
        <div style="margin: 20px 0;">
          <a href="${formUrl}" target="_blank" style="background: #10b981; color: white; padding: 10px 18px; text-decoration: none; border-radius: 6px; font-weight: bold; display: inline-block;">
            📋 Buka Form Jamaah (Dropdown)
          </a>
          <a href="${formEditUrl}" target="_blank" style="background: #e2e8f0; color: #334155; padding: 10px 18px; text-decoration: none; border-radius: 6px; font-weight: bold; display: inline-block; margin-left: 10px;">
            ⚙️ Edit Form
          </a>
        </div>
        <p style="font-size: 12px; color: #64748b;">Respons jamaah otomatis masuk ke tab <b>Log Absensi</b> dan teranalisis di <b>Dashboard</b>.</p>
      </div>
    `).setWidth(480).setHeight(250);

    SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Google Form Dropdown Siap");

  } catch (err) {
    Logger.log("Error create dropdown form: " + err.message);
    SpreadsheetApp.getActiveSpreadsheet().toast("Error: " + err.message, "Gagal Membuat Form", 8);
  }
}

function sinkronisasiDepartemenGoogleForm() {
  return buatGoogleFormDropdownBaru();
}

/**
 * Membersihkan data uji coba / sampah seperti "Participants (1)", "Afif Fathin Invite...", dll.
 * Merapikan kolom tema dan pemateri yang terisi judul window rapat.
 */
function bersihkanDataLogKotor() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetLog = ss.getSheetByName("Log Absensi");
  const sheetDash = ss.getSheetByName("Dashboard & Analitik");
  if (!sheetLog) return;

  const lastRow = sheetLog.getLastRow();
  if (lastRow <= 1) {
    ss.toast("Log Absensi masih kosong.", "Info", 3);
    return;
  }

  const maxCols = Math.max(sheetLog.getLastColumn(), 10);
  const data = sheetLog.getRange(2, 1, lastRow - 1, maxCols).getValues();
  const cleanData = [];
  const seenKeys = new Set();

  for (let rawRow of data) {
    let row = rawRow.slice(0, 10);
    while (row.length < 10) row.push("-");

    const rawName = String(row[3] || '').trim();
    const lower = rawName.toLowerCase();

    // Saring baris yang merupakan sampah / artifact Zoom
    if (!rawName || rawName.length < 3) continue;
    if (lower.startsWith('participants') || lower.startsWith('peserta')) continue;
    if (lower.includes('invite') || lower.includes('mute all') || lower.includes('mute')) continue;
    if (lower === 'j' || lower === 'afif fathin j') continue;

    // Bersihkan nama jika ada embel-embel role Zoom
    let cleanName = rawName.replace(/\((Host|Co-host|Host,\s*me|You|Me|Guest|Penyelenggara rapat|Anda)[^)]*\)/gi, '').trim();
    cleanName = cleanName.replace(/\b(invite|mute all|more|mute)\b/gi, '').trim();
    row[3] = cleanName;

    // Format Tanggal (Kolom 2)
    let rDate = row[1];
    if (rDate instanceof Date) {
      rDate = Utilities.formatDate(rDate, "Asia/Jakarta", "yyyy-MM-dd");
    } else {
      rDate = String(rDate).trim();
    }
    row[1] = rDate;

    // Format Jam Masuk (Kolom 3)
    let timeStr = row[2];
    if (timeStr instanceof Date) {
      timeStr = Utilities.formatDate(timeStr, "Asia/Jakarta", "HH:mm:ss");
    } else {
      timeStr = String(timeStr).trim();
    }
    row[2] = timeStr;

    // Format Timestamp (Kolom 1) - selalu seragam "yyyy-MM-dd HH:mm:ss"
    let ts = row[0];
    if (ts instanceof Date) {
      row[0] = Utilities.formatDate(ts, "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");
    } else {
      let tsStr = String(ts).trim();
      if (!tsStr.includes(":") && timeStr) {
        row[0] = `${rDate} ${timeStr}`;
      } else if (tsStr.includes("/")) {
        row[0] = `${rDate} ${timeStr}`;
      } else {
        row[0] = tsStr || `${rDate} ${timeStr}`;
      }
    }

    // Deduplikasi per tanggal & nama
    const dedupeKey = `${rDate}_${cleanName.toLowerCase()}`;
    if (seenKeys.has(dedupeKey)) continue;
    seenKeys.add(dedupeKey);

    // Ambil jadwal resmi hari tersebut
    const jadwal = getJadwalByDate(rDate);
    const officialTema = jadwal ? jadwal.tema : "Makan Bergizi Gratis";
    const officialPemateri = jadwal ? jadwal.pemateri : "Ustadz Ahmad Fauzi, Lc.";

    let jamTarget = "08:00";
    let toleransi = 15;
    if (jadwal && jadwal.jamMulai) {
      jamTarget = jadwal.jamMulai;
    }
    if (sheetDash) {
      const valJam = sheetDash.getRange("H6").getDisplayValue() || sheetDash.getRange("E4").getDisplayValue();
      if (valJam && valJam.includes(":")) jamTarget = valJam.trim();

      const valTol = sheetDash.getRange("K6").getValue() || sheetDash.getRange("E5").getValue();
      if (valTol && !isNaN(valTol) && Number(valTol) > 0) toleransi = parseInt(valTol, 10);
    }

    // Hitung ulang status kehadiran dan keterlambatan akurat
    const evalStatus = evaluateOnTime(timeStr, jamTarget, toleransi);

    // Deteksi jika baris mengalami pergeseran kolom (akibat deploy Web App versi lama yang hanya 9 kolom)
    // Ciri: row[7] berisi "TERLAMBAT" atau "TEPAT WAKTU" bukannya nama pemateri
    const col7Str = String(row[7] || '').trim().toUpperCase();
    if (col7Str === "TERLAMBAT" || col7Str === "TEPAT WAKTU" || !row[7] || row[7] === "-") {
      row[7] = officialPemateri;
    }

    // Bersihkan Tema jika terisi nama window zoom atau default 'Kajian Rutin'
    const col6Str = String(row[6] || '').trim();
    if (col6Str.includes("Zoom Meeting") || col6Str === "Online Meeting" || col6Str === "Kajian Rutin" || !col6Str) {
      row[6] = officialTema;
    }

    // Tetapkan Status Kehadiran (Kolom 9) & Keterlambatan (Kolom 10) yang akurat
    row[8] = evalStatus.status;
    row[9] = evalStatus.keterlambatanText;

    cleanData.push(row);
  }

  // Format kolom Log Absensi
  sheetLog.getRange("A2:A").setNumberFormat("@");
  sheetLog.getRange("B2:B").setNumberFormat("yyyy-mm-dd");
  sheetLog.getRange("C2:C").setNumberFormat("hh:mm:ss");
  sheetLog.getRange("I2:J").setHorizontalAlignment("center");

  // Hapus semua baris data lama
  sheetLog.getRange(2, 1, lastRow - 1, 10).clearContent();

  // Tulis ulang hanya data yang bersih, rapi, dan terformat sempurna
  if (cleanData.length > 0) {
    sheetLog.getRange(2, 1, cleanData.length, 10).setValues(cleanData);
  }

  // Pasang filter departemen & tanggal agar data yang baru dibersihkan langsung siap difilter
  pasangFilterDepartemenDanTanggal();

  ss.toast(`Pembersihan sukses! Seluruh data (${cleanData.length} baris) kini rapi 10 kolom, waktu seragam, dan status akurat.`, "✅ Berhasil Dibersihkan", 6);
}

// ----------------------------------------------------------------------------
// 2. SETUP UTAMA SISTEM KAJIAN (SUPER AMAN & ANTI-CRASH)
// ----------------------------------------------------------------------------
function setupSistemAbsensiOtomatis() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.setSpreadsheetTimeZone("Asia/Jakarta");

    const now = new Date();
    const todayStr = Utilities.formatDate(now, "Asia/Jakarta", "yyyy-MM-dd");

    Logger.log("=== MEMULAI SETUP SISTEM PRESENSI KAJIAN & PEMATERI ===");

    // ==========================================================================
    // TAB 1: MASTER JADWAL KAJIAN
    // ==========================================================================
    let sheetJadwal = ss.getSheetByName("Jadwal Kajian");
    if (!sheetJadwal) {
      sheetJadwal = ss.insertSheet("Jadwal Kajian", 0);
    }

    if (sheetJadwal.getLastRow() === 0) {
      const jadwalHeaders = ["Tanggal (YYYY-MM-DD)", "Jam Mulai", "Tema / Judul Kajian", "Nama Pemateri / Ustadz", "Keterangan"];
      sheetJadwal.appendRow(jadwalHeaders);

      const d1 = new Date(now.getTime() + 7 * 86400000);
      const d2 = new Date(now.getTime() + 14 * 86400000);
      const d3 = new Date(now.getTime() + 21 * 86400000);

      const s1 = Utilities.formatDate(d1, "Asia/Jakarta", "yyyy-MM-dd");
      const s2 = Utilities.formatDate(d2, "Asia/Jakarta", "yyyy-MM-dd");
      const s3 = Utilities.formatDate(d3, "Asia/Jakarta", "yyyy-MM-dd");

      sheetJadwal.appendRow([todayStr, "08:00", "Makan Bergizi Gratis", "Ustadz Ahmad Fauzi, Lc.", "Rutin Kajian Pekanan"]);
      sheetJadwal.appendRow([s1, "08:00", "Fiqih Muamalah Kontemporer", "Dr. Ustadz Khalid Al-Mansur", "Rutin Kajian Pekanan"]);
      sheetJadwal.appendRow([s2, "08:00", "Tazkiyatun Nafs: Kebersihan Hati", "Ustadz Hanif Al-Banjari", "Rutin Kajian Pekanan"]);
      sheetJadwal.appendRow([s3, "08:00", "Adab & Akhlak Penuntut Ilmu", "Ustadz Ahmad Fauzi, Lc.", "Rutin Kajian Pekanan"]);
    }

    sheetJadwal.getRange("A1:E1")
      .setBackground("#065f46")
      .setFontColor("#ffffff")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
    sheetJadwal.setRowHeight(1, 35);
    sheetJadwal.setFrozenRows(1);
    sheetJadwal.setColumnWidth(1, 140);
    sheetJadwal.setColumnWidth(2, 90);
    sheetJadwal.setColumnWidth(3, 240);
    sheetJadwal.setColumnWidth(4, 200);
    sheetJadwal.setColumnWidth(5, 180);

    // ==========================================================================
    // TAB 2: LOG ABSENSI LENGKAP (10 KOLOM)
    // ==========================================================================
    let sheetLog = ss.getSheetByName("Log Absensi");
    if (!sheetLog) {
      sheetLog = ss.insertSheet("Log Absensi", 1);
    }

    if (sheetLog.getLastRow() > 1 && sheetLog.getLastColumn() === 9) {
      sheetLog.insertColumnBefore(8);
      const numRows = sheetLog.getLastRow() - 1;
      const defaultPemateri = [];
      for (let k = 0; k < numRows; k++) {
        defaultPemateri.push(["Ustadz Ahmad Fauzi, Lc."]);
      }
      sheetLog.getRange(2, 8, numRows, 1).setValues(defaultPemateri);
    }

    const logHeaders = [
      "Timestamp", "Tanggal", "Jam Masuk", "Nama Lengkap", 
      "Departemen", "Metode", "Tema / Judul Kajian", "Nama Pemateri", 
      "Status Kehadiran", "Keterlambatan (Menit)"
    ];
    sheetLog.getRange(1, 1, 1, 10).setValues([logHeaders]);

    sheetLog.getRange("A1:J1")
      .setBackground("#1e3a8a")
      .setFontColor("#ffffff")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
    sheetLog.setRowHeight(1, 35);
    sheetLog.setFrozenRows(1);

    sheetLog.setColumnWidth(1, 150);
    sheetLog.setColumnWidth(2, 100);
    sheetLog.setColumnWidth(3, 90);
    sheetLog.setColumnWidth(4, 180);
    sheetLog.setColumnWidth(5, 120);
    sheetLog.setColumnWidth(6, 130);
    sheetLog.setColumnWidth(7, 200);
    sheetLog.setColumnWidth(8, 180);
    sheetLog.setColumnWidth(9, 130);
    sheetLog.setColumnWidth(10, 130);

    const ruleOnTime = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("TEPAT WAKTU")
      .setBackground("#d1fae5")
      .setFontColor("#065f46")
      .setBold(true)
      .setRanges([sheetLog.getRange("I2:I1000")])
      .build();

    const ruleLate = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("TERLAMBAT")
      .setBackground("#fee2e2")
      .setFontColor("#991b1b")
      .setBold(true)
      .setRanges([sheetLog.getRange("I2:I1000")])
      .build();

    sheetLog.setConditionalFormatRules([ruleOnTime, ruleLate]);
    sheetLog.getRange("A2:A").setNumberFormat("@");
    sheetLog.getRange("B2:B").setNumberFormat("yyyy-mm-dd");
    sheetLog.getRange("C2:C").setNumberFormat("hh:mm:ss");
    sheetLog.getRange("I2:J").setHorizontalAlignment("center");
    sheetJadwal.getRange("B2:B").setNumberFormat("@");

    // ==========================================================================
    // TAB 3: DASHBOARD & ANALITIK
    // ==========================================================================
    let sheetDash = ss.getSheetByName("Dashboard & Analitik");
    if (!sheetDash) {
      sheetDash = ss.insertSheet("Dashboard & Analitik", 2);
    } else {
      // Unmerge semua cell lama agar tidak crash saat merge ulang
      try {
        sheetDash.getDataRange().breakApart();
      } catch (e) {}

      // Hapus chart lama
      const existingCharts = sheetDash.getCharts();
      for (let c of existingCharts) {
        try { sheetDash.removeChart(c); } catch (e) {}
      }
      sheetDash.clear();
    }

    // Header Title
    sheetDash.getRange("A1:L1").merge()
      .setValue("📊 DASHBOARD ANALITIK RUTIN KAJIAN & PERFORMA PEMATERI (REAL-TIME)")
      .setBackground("#0f172a")
      .setFontColor("#38bdf8")
      .setFontSize(14)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
    sheetDash.setRowHeight(1, 45);

    // 1. KOTAK FILTER PERIODE (A3:B6)
    sheetDash.getRange("A3:B3").merge().setValue("🎛️ FILTER PERIODE KAJIAN").setFontWeight("bold").setBackground("#e2e8f0").setFontColor("#1e293b");
    sheetDash.getRange("A4").setValue("Pilih Periode:");
    
    const rulePeriod = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Semua Periode", "Hari Ini", "Minggu Ini", "Bulan Ini", "Tahun Ini"], true)
      .setAllowInvalid(false)
      .build();
    sheetDash.getRange("B4").setDataValidation(rulePeriod).setValue("Semua Periode").setFontWeight("bold").setBackground("#fef08a");

    sheetDash.getRange("A5").setValue("Tanggal Mulai (Kalender):");
    sheetDash.getRange("B5").setFormula(`=IF(B4="Hari Ini"; TODAY(); IF(B4="Minggu Ini"; TODAY()-WEEKDAY(TODAY();2)+1; IF(B4="Bulan Ini"; DATE(YEAR(TODAY()); MONTH(TODAY()); 1); IF(B4="Tahun Ini"; DATE(YEAR(TODAY()); 1; 1); DATE(2020; 1; 1)))))`).setNumberFormat("yyyy-mm-dd");
    sheetDash.getRange("A6").setValue("Tanggal Akhir (Kalender):");
    sheetDash.getRange("B6").setFormula(`=IF(B4="Hari Ini"; TODAY(); IF(B4="Minggu Ini"; TODAY()-WEEKDAY(TODAY();2)+7; IF(B4="Bulan Ini"; EOMONTH(TODAY(); 0); IF(B4="Tahun Ini"; DATE(YEAR(TODAY()); 12; 31); DATE(2099; 12; 31)))))`).setNumberFormat("yyyy-mm-dd");
    sheetDash.getRange("B5:B6").setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build());
    sheetDash.getRange("A3:B6").setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);

    // 2. KOTAK GOOGLE FORM OFFLINE (D3:E6) - Menggantikan tabel Pengaturan Kajian yang dihapus
    let formUrl = PropertiesService.getScriptProperties().getProperty("GOOGLE_FORM_URL");
    let formEditUrl = PropertiesService.getScriptProperties().getProperty("GOOGLE_FORM_EDIT_URL");

    sheetDash.getRange("D3:E3").merge().setValue("📋 GOOGLE FORM (OFFLINE)").setFontWeight("bold").setBackground("#e2e8f0").setFontColor("#1e293b");
    if (formUrl) {
      sheetDash.getRange("D4").setValue("Form Jamaah:");
      sheetDash.getRange("E4").setFormula(`=HYPERLINK("${formUrl}"; "👉 Klik Isi Absensi")`).setFontWeight("bold");
      sheetDash.getRange("D5").setValue("Link URL:");
      sheetDash.getRange("E5").setValue(formUrl);
      sheetDash.getRange("D6").setValue("Edit Form:");
      sheetDash.getRange("E6").setFormula(`=HYPERLINK("${formEditUrl || formUrl}"; "⚙️ Pengaturan Soal")`);
    } else {
      sheetDash.getRange("D4").setValue("Status Form:");
      sheetDash.getRange("E4").setValue("Form Aktif");
      sheetDash.getRange("D5").setValue("Keterangan:");
      sheetDash.getRange("E5").setValue("Jalankan setup form");
      sheetDash.getRange("D6").setValue("Dropdown:");
      sheetDash.getRange("E6").setValue("14 Departemen");
    }
    sheetDash.getRange("D3:E6").setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);

    // 3. KOTAK JADWAL KAJIAN HARI INI & PENGATURAN WAKTU (G3:L7) - Atur Maksimal Keterlambatan di Sini
    sheetDash.getRange("G3:L3").merge()
      .setValue("📅 JADWAL KAJIAN HARI INI & PENGATURAN WAKTU")
      .setFontWeight("bold")
      .setBackground("#dcfce7")
      .setFontColor("#166534")
      .setHorizontalAlignment("center");

    sheetDash.getRange("G4").setValue("Tema Kajian:").setFontWeight("bold");
    sheetDash.getRange("H4:L4").merge()
      .setFormula(`=IFERROR(INDEX('Jadwal Kajian'!C2:C; MATCH(TODAY(); 'Jadwal Kajian'!A2:A; 0)); IFERROR(INDEX('Jadwal Kajian'!C2:C; MATCH(TEXT(TODAY(); "yyyy-mm-dd"); 'Jadwal Kajian'!A2:A; 0)); IFERROR(INDEX('Jadwal Kajian'!C2:C; 1); "Kajian Rutin")))`)
      .setFontWeight("bold");

    sheetDash.getRange("G5").setValue("Pemateri / Ustadz:").setFontWeight("bold");
    sheetDash.getRange("H5:L5").merge()
      .setFormula(`=IFERROR(INDEX('Jadwal Kajian'!D2:D; MATCH(TODAY(); 'Jadwal Kajian'!A2:A; 0)); IFERROR(INDEX('Jadwal Kajian'!D2:D; MATCH(TEXT(TODAY(); "yyyy-mm-dd"); 'Jadwal Kajian'!A2:A; 0)); IFERROR(INDEX('Jadwal Kajian'!D2:D; 1); "-")))`)
      .setFontWeight("bold")
      .setFontColor("#059669");

    sheetDash.getRange("G6").setValue("Jam Mulai:");
    sheetDash.getRange("H6:I6").merge()
      .setFormula(`=TEXT(IFERROR(INDEX('Jadwal Kajian'!B2:B; MATCH(TODAY(); 'Jadwal Kajian'!A2:A; 0)); IFERROR(INDEX('Jadwal Kajian'!B2:B; MATCH(TEXT(TODAY(); "yyyy-mm-dd"); 'Jadwal Kajian'!A2:A; 0)); IFERROR(INDEX('Jadwal Kajian'!B2:B; 1); "08:00"))); "HH:mm")`)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setNumberFormat("@");

    sheetDash.getRange("J6").setValue("Maks. Telat (Menit):").setFontWeight("bold").setHorizontalAlignment("right");
    sheetDash.getRange("K6:L6").merge()
      .setValue(15)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setBackground("#fef9c3")
      .setFontColor("#854d0e")
      .setNumberFormat("0");

    sheetDash.getRange("G7:H7").merge().setValue("Batas On-Time Resmi:").setFontWeight("bold");
    sheetDash.getRange("I7:L7").merge()
      .setFormula(`=TEXT(TIMEVALUE(H6) + TIME(0; K6; 0); "HH:mm") & " (Toleransi " & K6 & " Menit)"`)
      .setFontWeight("bold")
      .setFontColor("#15803d")
      .setBackground("#f0fdf4")
      .setHorizontalAlignment("center");

    sheetDash.getRange("G3:L7").setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);

    // 4. KPI SUMMARY STATISTIK PERIODE (A8:B12)
    sheetDash.getRange("A8:B8").merge().setValue("📈 RINGKASAN SESUAI PERIODE").setFontWeight("bold").setBackground("#e2e8f0").setFontColor("#1e293b");
    sheetDash.getRange("A9").setValue("Total Jamaah Hadir:");
    sheetDash.getRange("B9").setFormula(`=COUNTIFS('Log Absensi'!B:B; ">="&TEXT(B5; "yyyy-mm-dd"); 'Log Absensi'!B:B; "<="&TEXT(B6; "yyyy-mm-dd"); 'Log Absensi'!D:D; "<>")`).setFontWeight("bold");
    sheetDash.getRange("A10").setValue("Tepat Waktu:");
    sheetDash.getRange("B10").setFormula(`=COUNTIFS('Log Absensi'!B:B; ">="&TEXT(B5; "yyyy-mm-dd"); 'Log Absensi'!B:B; "<="&TEXT(B6; "yyyy-mm-dd"); 'Log Absensi'!I:I; "TEPAT WAKTU")`).setFontColor("#059669").setFontWeight("bold");
    sheetDash.getRange("A11").setValue("Terlambat:");
    sheetDash.getRange("B11").setFormula(`=COUNTIFS('Log Absensi'!B:B; ">="&TEXT(B5; "yyyy-mm-dd"); 'Log Absensi'!B:B; "<="&TEXT(B6; "yyyy-mm-dd"); 'Log Absensi'!I:I; "TERLAMBAT")`).setFontColor("#dc2626").setFontWeight("bold");
    sheetDash.getRange("A12").setValue("Persentase On-Time:");
    sheetDash.getRange("B12").setFormula(`=IFERROR(IF(AND(ISNUMBER(B9); B9>0); B10/B9; 0); 0)`).setNumberFormat("0.0%").setFontWeight("bold");
    sheetDash.getRange("A8:B12").setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);

    // 5. TABEL PERFORMA DEPARTEMEN / DIVISI (A14:E29)
    sheetDash.getRange("A14:E14").merge().setValue("🏢 KEHADIRAN PER DEPARTEMEN / DIVISI").setFontWeight("bold").setBackground("#e2e8f0").setFontColor("#1e293b");
    sheetDash.getRange("A15:E15").setValues([["Departemen", "Tepat Waktu", "Terlambat", "Total", "% On-Time"]]);
    sheetDash.getRange("A15:E15").setBackground("#334155").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");

    const depts = DEPARTEMEN_LIST;
    const endRow = 15 + depts.length; // 15 + 14 = 29
    const deptRows = [];
    for (let i = 0; i < depts.length; i++) {
      const row = 16 + i;
      const d = depts[i];
      deptRows.push([
        `=COUNTIFS('Log Absensi'!E:E; "${d}"; 'Log Absensi'!I:I; "TEPAT WAKTU"; 'Log Absensi'!B:B; ">="&TEXT($B$5; "yyyy-mm-dd"); 'Log Absensi'!B:B; "<="&TEXT($B$6; "yyyy-mm-dd"))`,
        `=COUNTIFS('Log Absensi'!E:E; "${d}"; 'Log Absensi'!I:I; "TERLAMBAT"; 'Log Absensi'!B:B; ">="&TEXT($B$5; "yyyy-mm-dd"); 'Log Absensi'!B:B; "<="&TEXT($B$6; "yyyy-mm-dd"))`,
        `=B${row}+C${row}`,
        `=IFERROR(IF(AND(ISNUMBER(D${row}); D${row}>0); B${row}/D${row}; 0); 0)`
      ]);
    }
    sheetDash.getRange(`A16:A${endRow}`).setValues(depts.map(d => [d])).setFontWeight("bold");
    sheetDash.getRange(`B16:E${endRow}`).setFormulas(deptRows);
    sheetDash.getRange(`B16:B${endRow}`).setFontColor("#059669").setFontWeight("bold").setHorizontalAlignment("center");
    sheetDash.getRange(`C16:C${endRow}`).setFontColor("#dc2626").setFontWeight("bold").setHorizontalAlignment("center");
    sheetDash.getRange(`D16:D${endRow}`).setFontWeight("bold").setHorizontalAlignment("center");
    sheetDash.getRange(`E16:E${endRow}`).setNumberFormat("0.0%").setHorizontalAlignment("center");
    sheetDash.getRange(`A15:E${endRow}`).setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);

    // 6. TABEL EVALUASI PERFORMA PEMATERI (G8:L29)
    sheetDash.getRange("G8:L8").merge().setValue("🏆 EVALUASI PERFORMA PEMATERI (USTADZ)").setFontWeight("bold").setBackground("#065f46").setFontColor("#ffffff");
    sheetDash.getRange("G9:L9").setValues([["Nama Pemateri / Ustadz", "Total Jamaah", "Total Sesi", "Rata-rata / Sesi", "% On-Time", "Antusiasme"]]);
    sheetDash.getRange("G9:L9").setBackground("#1e293b").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");

    sheetDash.getRange("G10").setFormula(`=IFERROR(UNIQUE(FILTER('Jadwal Kajian'!D2:D; 'Jadwal Kajian'!D2:D<>"")); "Belum ada pemateri")`);

    const pemateriFormulas = [];
    for (let r = 10; r <= endRow; r++) {
      pemateriFormulas.push([
        `=IF(OR(G${r}=""; G${r}="Belum ada pemateri"); 0; COUNTIFS('Log Absensi'!H:H; G${r}; 'Log Absensi'!B:B; ">="&TEXT($B$5; "yyyy-mm-dd"); 'Log Absensi'!B:B; "<="&TEXT($B$6; "yyyy-mm-dd")))`,
        `=IF(OR(G${r}=""; G${r}="Belum ada pemateri"); 0; COUNTIFS('Jadwal Kajian'!D:D; G${r}; 'Jadwal Kajian'!A:A; ">="&TEXT($B$5; "yyyy-mm-dd"); 'Jadwal Kajian'!A:A; "<="&TEXT($B$6; "yyyy-mm-dd")))`,
        `=IF(OR(G${r}=""; G${r}="Belum ada pemateri"); 0; IFERROR(IF(AND(ISNUMBER(I${r}); I${r}>0); ROUND(H${r}/I${r}; 1); IF(AND(ISNUMBER(H${r}); H${r}>0); H${r}; 0)); 0))`,
        `=IF(OR(G${r}=""; G${r}="Belum ada pemateri"); 0; IFERROR(IF(AND(ISNUMBER(H${r}); H${r}>0); COUNTIFS('Log Absensi'!H:H; G${r}; 'Log Absensi'!I:I; "TEPAT WAKTU"; 'Log Absensi'!B:B; ">="&TEXT($B$5; "yyyy-mm-dd"); 'Log Absensi'!B:B; "<="&TEXT($B$6; "yyyy-mm-dd")) / H${r}; 0); 0))`,
        `=IF(OR(G${r}=""; G${r}="Belum ada pemateri"); "-"; IF(AND(ISNUMBER(H${r}); H${r}>=20); "🔥 Sangat Tinggi"; IF(AND(ISNUMBER(H${r}); H${r}>=10); "⭐ Tinggi"; IF(AND(ISNUMBER(H${r}); H${r}>0); "👍 Baik"; "-"))))`
      ]);
    }
    sheetDash.getRange(`H10:L${endRow}`).setFormulas(pemateriFormulas);
    sheetDash.getRange(`H10:H${endRow}`).setFontWeight("bold").setFontColor("#0284c7").setHorizontalAlignment("center");
    sheetDash.getRange(`I10:J${endRow}`).setHorizontalAlignment("center");
    sheetDash.getRange(`K10:K${endRow}`).setNumberFormat("0.0%").setHorizontalAlignment("center");
    sheetDash.getRange(`L10:L${endRow}`).setHorizontalAlignment("center");
    sheetDash.getRange(`G8:L${endRow}`).setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);

    sheetDash.setColumnWidth(1, 160);
    sheetDash.setColumnWidth(2, 110);
    sheetDash.setColumnWidth(3, 95);
    sheetDash.setColumnWidth(4, 120);
    sheetDash.setColumnWidth(5, 120);
    sheetDash.setColumnWidth(6, 30);
    sheetDash.setColumnWidth(7, 180);
    sheetDash.setColumnWidth(8, 100);
    sheetDash.setColumnWidth(9, 90);
    sheetDash.setColumnWidth(10, 110);
    sheetDash.setColumnWidth(11, 90);
    sheetDash.setColumnWidth(12, 115);

    // ==========================================================================
    // 7. PASANG 3 DIAGRAM GRAFIK ELEGAN & BERDAMPINGAN DI BARIS 31
    // ==========================================================================
    gambarUlangSemuaGrafik();

    // Otomatis bersihkan kolom ganda pada Form Responses jika ada
    perbaikiKolomGandaFormResponses();

    // Buat / Update tab khusus Pencarian & Filter dengan Kalender dan Dropdown Departemen
    setupTabFilterDanPencarian();

    // Buat / Update tab khusus Master Data Karyawan dengan Dropdown 14 Departemen & Monitoring
    setupTabDataKaryawan();

    // Pasang filter otomatis pada tab Log Absensi & Form Responses 1
    pasangFilterDepartemenDanTanggal();

    // Otomatis sinkronisasi 14 departemen ke Google Form
    sinkronisasiDepartemenGoogleForm();

    Logger.log("=== SETUP SELESAI DENGAN SUKSES! ===");
    ss.toast("Setup Presensi Kajian Berhasil! Semua tabel, 14 departemen, dan grafik siap.", "✅ Setup Sukses", 6);

  } catch (err) {
    Logger.log("ERROR SETUP: " + err.message + "\n" + err.stack);
    SpreadsheetApp.getActiveSpreadsheet().toast("Info: " + err.message, "Perhatian", 8);
  }
}

/**
 * ============================================================================
 * FUNGSI STANDALONE: BUAT / GAMBAR ULANG SEMUA GRAFIK DASHBOARD
 * Menu Toolbar: 🕌 Presensi Kajian ➜ 📊 Buat / Gambar Ulang Semua Grafik
 * ============================================================================
 * Membersihkan dan menggambar ulang 3 diagram grafik BERDAMPINGAN di Baris 31:
 * 1. Donut Chart (Kolom A-C): Rasio Ketepatan Waktu Jamaah (Hijau Tepat vs Merah Telat)
 * 2. Stacked Column Chart (Kolom D-H): Kehadiran 14 Departemen (Tepat Waktu vs Terlambat)
 * 3. Column Chart (Kolom I-L): Total Kehadiran Jamaah per Pemateri (Ustadz)
 */
function gambarUlangSemuaGrafik() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetDash = ss.getSheetByName("Dashboard & Analitik");
  if (!sheetDash) {
    ss.toast("Tab 'Dashboard & Analitik' tidak ditemukan.", "Peringatan", 4);
    return;
  }

  // 1. Pastikan jumlah baris cukup banyak (minimal 60 baris) agar posisi baris 31 selalu valid
  if (sheetDash.getMaxRows() < 60) {
    sheetDash.insertRowsAfter(sheetDash.getMaxRows(), 60 - sheetDash.getMaxRows());
  }

  // 2. Hapus semua grafik lama agar tidak bertumpuk
  const existingCharts = sheetDash.getCharts();
  for (let c of existingCharts) {
    try { sheetDash.removeChart(c); } catch (e) {}
  }

  const endRow = 15 + DEPARTEMEN_LIST.length; // 29

  // 3. Pastikan formula pemateri menghasilkan angka 0
  const pemateriFormulas = [];
  for (let r = 10; r <= endRow; r++) {
    pemateriFormulas.push([
      `=IF(OR(G${r}=""; G${r}="Belum ada pemateri"); 0; COUNTIFS('Log Absensi'!H:H; G${r}; 'Log Absensi'!B:B; ">="&TEXT($B$5; "yyyy-mm-dd"); 'Log Absensi'!B:B; "<="&TEXT($B$6; "yyyy-mm-dd")))`,
      `=IF(OR(G${r}=""; G${r}="Belum ada pemateri"); 0; COUNTIFS('Jadwal Kajian'!D:D; G${r}; 'Jadwal Kajian'!A:A; ">="&TEXT($B$5; "yyyy-mm-dd"); 'Jadwal Kajian'!A:A; "<="&TEXT($B$6; "yyyy-mm-dd")))`,
      `=IF(OR(G${r}=""; G${r}="Belum ada pemateri"); 0; IFERROR(IF(AND(ISNUMBER(I${r}); I${r}>0); ROUND(H${r}/I${r}; 1); IF(AND(ISNUMBER(H${r}); H${r}>0); H${r}; 0)); 0))`,
      `=IF(OR(G${r}=""; G${r}="Belum ada pemateri"); 0; IFERROR(IF(AND(ISNUMBER(H${r}); H${r}>0); COUNTIFS('Log Absensi'!H:H; G${r}; 'Log Absensi'!I:I; "TEPAT WAKTU"; 'Log Absensi'!B:B; ">="&TEXT($B$5; "yyyy-mm-dd"); 'Log Absensi'!B:B; "<="&TEXT($B$6; "yyyy-mm-dd")) / H${r}; 0); 0))`,
      `=IF(OR(G${r}=""; G${r}="Belum ada pemateri"); "-"; IF(AND(ISNUMBER(H${r}); H${r}>=20); "🔥 Sangat Tinggi"; IF(AND(ISNUMBER(H${r}); H${r}>=10); "⭐ Tinggi"; IF(AND(ISNUMBER(H${r}); H${r}>0); "👍 Baik"; "-"))))`
    ]);
  }
  sheetDash.getRange(`H10:L${endRow}`).setFormulas(pemateriFormulas);

  SpreadsheetApp.flush();

  let berhasil = 0;
  const errors = [];

  // --------------------------------------------------------------------------
  // GRAFIK 1: DONUT - Rasio Ketepatan Waktu Jamaah (Kolom A, Baris 31)
  // --------------------------------------------------------------------------
  try {
    const donutChart = sheetDash.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(sheetDash.getRange("A10:B11"))
      .setPosition(31, 1, 0, 0)
      .setOption("title", "Rasio Ketepatan Waktu Jamaah")
      .setOption("pieHole", 0.45)
      .setOption("colors", ["#10b981", "#ef4444"])
      .setOption("width", 360)
      .setOption("height", 270)
      .build();
    sheetDash.insertChart(donutChart);
    berhasil++;
  } catch (e1) {
    errors.push("Donut: " + e1.message);
    Logger.log("Grafik 1 Error: " + e1.message);
  }

  // --------------------------------------------------------------------------
  // GRAFIK 2: STACKED COLUMN - Kehadiran 14 Departemen (Kolom D, Baris 31)
  // --------------------------------------------------------------------------
  try {
    const deptChart = sheetDash.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sheetDash.getRange(`A15:C${endRow}`))
      .setNumHeaders(1)
      .setPosition(31, 4, 0, 0)
      .setOption("title", "Kehadiran per Departemen (Tepat Waktu vs Terlambat)")
      .setOption("isStacked", true)
      .setOption("colors", ["#10b981", "#ef4444"])
      .setOption("width", 540)
      .setOption("height", 270)
      .build();
    sheetDash.insertChart(deptChart);
    berhasil++;
  } catch (e2) {
    errors.push("Departemen: " + e2.message);
    Logger.log("Grafik 2 Error: " + e2.message);
  }

  // --------------------------------------------------------------------------
  // GRAFIK 3: COLUMN - Total Kehadiran Jamaah per Pemateri (Kolom I, Baris 31)
  // --------------------------------------------------------------------------
  try {
    let pemateriEnd = 10;
    try {
      const gVals = sheetDash.getRange("G10:G29").getValues();
      for (let idx = 0; idx < gVals.length; idx++) {
        const val = String(gVals[idx][0] || '').trim();
        if (val && val !== "Belum ada pemateri" && val !== "#N/A" && val !== "-") {
          pemateriEnd = 10 + idx;
        }
      }
    } catch (eFind) {
      pemateriEnd = 10;
    }
    pemateriEnd = Math.max(pemateriEnd, 10);

    const pemateriChart = sheetDash.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sheetDash.getRange(`G9:H${pemateriEnd}`))
      .setNumHeaders(1)
      .setPosition(31, 9, 0, 0)
      .setOption("title", "Total Kehadiran Jamaah per Pemateri")
      .setOption("colors", ["#0284c7"])
      .setOption("width", 420)
      .setOption("height", 270)
      .build();
    sheetDash.insertChart(pemateriChart);
    berhasil++;
  } catch (e3) {
    errors.push("Pemateri: " + e3.message);
    Logger.log("Grafik 3 Error: " + e3.message);
  }

  if (errors.length > 0) {
    ss.toast(`Peringatan: ${errors.join(" | ")}`, "Detail Kendala", 8);
  } else {
    ss.toast(`✅ Berhasil memuat ${berhasil} dari 3 grafik berdampingan di Baris 31!`, "📊 Grafik Siap", 5);
  }
}

// ----------------------------------------------------------------------------
// 3. FUNGSI LOOKUP MASTER JADWAL KAJIAN
// ----------------------------------------------------------------------------
function parseJamMulai(val) {
  if (!val) return "08:00";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Jakarta", "HH:mm");
  }
  if (typeof val === "number") {
    let frac = val;
    if (frac > 1) {
      if (frac <= 24) return `${String(Math.floor(frac)).padStart(2, '0')}:00`;
      if (frac > 24 && frac <= 100) frac = frac / 100;
      else frac = frac - Math.floor(frac);
    }
    const totalMinutes = Math.round(frac * 1440);
    const h = Math.floor(totalMinutes / 60) % 24;
    const m = totalMinutes % 60;
    return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
  }
  const str = String(val).trim().replace(".", ":");
  const match = str.match(/(\d{1,2}):(\d{2})/);
  if (match) {
    return `${match[1].padStart(2, '0')}:${match[2]}`;
  }
  return "08:00";
}

function getJadwalByDate(dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetJadwal = ss.getSheetByName("Jadwal Kajian");
  if (!sheetJadwal) return null;

  const lastRow = sheetJadwal.getLastRow();
  if (lastRow <= 1) return null;

  const values = sheetJadwal.getRange(2, 1, lastRow - 1, 5).getValues();
  for (let row of values) {
    let rDate = row[0];
    if (rDate instanceof Date) {
      rDate = Utilities.formatDate(rDate, "Asia/Jakarta", "yyyy-MM-dd");
    } else {
      rDate = String(rDate).trim();
    }

    if (rDate === dateStr) {
      return {
        tanggal: rDate,
        jamMulai: parseJamMulai(row[1]),
        tema: String(row[2] || "Kajian Rutin").trim(),
        pemateri: String(row[3] || "Ustadz Pembina").trim(),
        keterangan: String(row[4] || "").trim()
      };
    }
  }

  // FALLBACK CERDAS: Jika tanggal hari ini belum tertulis di master jadwal,
  // gunakan jadwal terdekat/aktif di baris pertama agar tema & pemateri selalu terisi rapi.
  const firstRow = values[0];
  if (firstRow && firstRow[2]) {
    return {
      tanggal: dateStr,
      jamMulai: parseJamMulai(firstRow[1]),
      tema: String(firstRow[2] || "Makan Bergizi Gratis").trim(),
      pemateri: String(firstRow[3] || "Ustadz Ahmad Fauzi, Lc.").trim(),
      keterangan: String(firstRow[4] || "").trim()
    };
  }

  return null;
}

// ----------------------------------------------------------------------------
// 4. TRIGGER PRESENSI OFFLINE (GOOGLE FORM SUBMIT)
// ----------------------------------------------------------------------------
function onGoogleFormSubmit(e) {
  if (!e) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetLog = ss.getSheetByName("Log Absensi");
  const sheetDash = ss.getSheetByName("Dashboard & Analitik");
  if (!sheetLog) return;

  const now = new Date();
  const timeStr = Utilities.formatDate(now, "Asia/Jakarta", "HH:mm:ss");
  const dateStr = Utilities.formatDate(now, "Asia/Jakarta", "yyyy-MM-dd");
  const timestampStr = Utilities.formatDate(now, "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");

  let nama = "Tanpa Nama";
  let dept = "Umum";

  if (e.namedValues) {
    if (e.namedValues["Nama Lengkap"] && e.namedValues["Nama Lengkap"][0]) {
      nama = e.namedValues["Nama Lengkap"][0].trim();
    }
    const deptCandidates = [
      ...(e.namedValues["Departemen / Divisi"] || []),
      ...(e.namedValues["Departemen"] || [])
    ];
    for (let cand of deptCandidates) {
      if (cand && String(cand).trim() !== "") {
        dept = String(cand).trim();
        break;
      }
    }
  } else if (e.values) {
    nama = e.values[1] ? e.values[1].trim() : "Tanpa Nama";
    dept = e.values.slice(2).find(v => v && String(v).trim() !== '') || "Umum";
  }

  // Jika departemen belum terdeteksi atau masih "Umum", cari dari Master Data Karyawan
  if (dept === "Umum" || !dept) {
    const regDept = lookupRegisteredDepartment(nama);
    if (regDept) dept = regDept;
  }

  const jadwal = getJadwalByDate(dateStr);
  const temaKajian = jadwal ? jadwal.tema : "Makan Bergizi Gratis";
  const namaPemateri = jadwal ? jadwal.pemateri : "Ustadz Ahmad Fauzi, Lc.";

  let jamTarget = "08:00";
  let toleransi = 15;
  if (jadwal && jadwal.jamMulai) {
    jamTarget = jadwal.jamMulai;
  }
  if (sheetDash) {
    const valJam = sheetDash.getRange("H6").getDisplayValue() || sheetDash.getRange("E4").getDisplayValue();
    if (valJam && valJam.includes(":")) jamTarget = valJam.trim();

    const valTol = sheetDash.getRange("K6").getValue() || sheetDash.getRange("E5").getValue();
    if (valTol && !isNaN(valTol) && Number(valTol) > 0) toleransi = parseInt(valTol, 10);
  }

  const evalStatus = evaluateOnTime(timeStr, jamTarget, toleransi);

  sheetLog.appendRow([
    timestampStr,
    dateStr,
    timeStr,
    nama,
    dept,
    "Google Form (Offline)",
    temaKajian,
    namaPemateri,
    evalStatus.status,
    evalStatus.keterlambatanText
  ]);

  // Pastikan filter di Log Absensi dan Form Responses tetap melingkupi baris data baru
  pasangFilterDepartemenDanTanggal();
}

// ----------------------------------------------------------------------------
// 5. API WEB APP (GOOGLE MEET & ZOOM EXTENSION)
// ----------------------------------------------------------------------------
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: "No payload received"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    const contents = JSON.parse(e.postData.contents);
    const participants = contents.participants || [];
    const sessionName = contents.sessionName || "Online Meeting";
    const platform = contents.platform === "ZOOM" ? "Zoom Meeting" : "Google Meet";

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetLog = ss.getSheetByName("Log Absensi");
    const sheetDash = ss.getSheetByName("Dashboard & Analitik");
    if (!sheetLog) throw new Error("Sheet 'Log Absensi' tidak ditemukan. Jalankan setup terlebih dahulu.");

    const now = new Date();
    const timeStr = Utilities.formatDate(now, "Asia/Jakarta", "HH:mm:ss");
    const dateStr = Utilities.formatDate(now, "Asia/Jakarta", "yyyy-MM-dd");
    const timestampStr = Utilities.formatDate(now, "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");

    const jadwal = getJadwalByDate(dateStr);
    const temaKajian = jadwal ? jadwal.tema : "Makan Bergizi Gratis";
    const namaPemateri = jadwal ? jadwal.pemateri : "Ustadz Ahmad Fauzi, Lc.";

    let jamTarget = "08:00";
    let toleransi = 15;
    if (jadwal && jadwal.jamMulai) {
      jamTarget = jadwal.jamMulai;
    }
    if (sheetDash) {
      const valJam = sheetDash.getRange("H6").getDisplayValue() || sheetDash.getRange("E4").getDisplayValue();
      if (valJam && valJam.includes(":")) jamTarget = valJam.trim();

      const valTol = sheetDash.getRange("K6").getValue() || sheetDash.getRange("E5").getValue();
      if (valTol && !isNaN(valTol) && Number(valTol) > 0) toleransi = parseInt(valTol, 10);
    }

    const evalStatus = evaluateOnTime(timeStr, jamTarget, toleransi);

    const lastRow = sheetLog.getLastRow();
    const existingSet = new Set();
    if (lastRow > 1) {
      // Dapatkan indeks kolom secara dinamis dari header baris 1
      const headers = sheetLog.getRange(1, 1, 1, Math.max(sheetLog.getLastColumn(), 10)).getValues()[0];
      let colDate = -1;
      let colName = -1;
      let colDept = -1;

      for (let c = 0; c < headers.length; c++) {
        const h = String(headers[c] || '').toLowerCase().trim();
        if (h.includes("tanggal")) colDate = c;
        else if (h.includes("nama")) colName = c;
        else if (h.includes("departemen") || h.includes("divisi")) colDept = c;
      }

      // Fallback kolom standar jika header tidak ditemukan: B (col 1), D (col 3), E (col 4)
      if (colDate === -1) colDate = 1;
      if (colName === -1) colName = 3;
      if (colDept === -1) colDept = 4;

      const numCols = Math.max(colDate, colName, colDept) + 1;
      const data = sheetLog.getRange(2, 1, lastRow - 1, numCols).getValues();

      data.forEach(row => {
        let rDate = row[colDate];
        if (rDate instanceof Date) {
          rDate = Utilities.formatDate(rDate, "Asia/Jakarta", "yyyy-MM-dd");
        } else {
          rDate = String(rDate || '').trim();
        }

        const rName = String(row[colName] || '').toLowerCase().trim();
        const rDept = String(row[colDept] || '').toLowerCase().trim();

        if (rDate === dateStr && rName) {
          // Simpan identitas nama dan kombinasi nama|departemen
          existingSet.add(`${rName}|${rDept}`);
          existingSet.add(rName);
        }
      });
    }

    let recordedCount = 0;
    const newRows = [];

    for (let rawName of participants) {
      const parsed = parseParticipantName(rawName);
      if (!parsed || !parsed.name) continue;

      // Jika departemen belum terdeteksi atau masih "Umum", cari dari Master Data Karyawan
      if (!parsed.department || parsed.department === "Umum") {
        const regDept = lookupRegisteredDepartment(parsed.name);
        if (regDept) parsed.department = regDept;
      }

      const cleanName = parsed.name.toLowerCase().trim();
      const cleanDept = (parsed.department || '').toLowerCase().trim();
      const compositeKey = `${cleanName}|${cleanDept}`;

      // Deduplikasi berbasis identitas nama peserta, BUKAN departemen!
      if (existingSet.has(compositeKey) || existingSet.has(cleanName)) continue;

      existingSet.add(compositeKey);
      existingSet.add(cleanName);
      newRows.push([
        timestampStr,
        dateStr,
        timeStr,
        parsed.name,
        parsed.department,
        platform,
        temaKajian,
        namaPemateri,
        evalStatus.status,
        evalStatus.keterlambatanText
      ]);
      recordedCount++;
    }

    if (newRows.length > 0) {
      sheetLog.getRange(sheetLog.getLastRow() + 1, 1, newRows.length, 10).setValues(newRows);
      pasangFilterDepartemenDanTanggal();
    }

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      count: recordedCount,
      tema: temaKajian,
      pemateri: namaPemateri,
      message: `Berhasil mencatat ${recordedCount} jamaah untuk kajian '${temaKajian}' bersama ${namaPemateri}!`
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: "online",
    service: "Sistem Presensi Rutin Kajian & Evaluasi Pemateri",
    time: new Date()
  })).setMimeType(ContentService.MimeType.JSON);
}

function evaluateOnTime(timeStr, targetTimeStr, toleranceMinutes) {
  let [h, m] = (timeStr || "00:00").split(":").map(Number);
  if (isNaN(h)) h = 0;
  if (isNaN(m)) m = 0;

  const targetParsed = parseJamMulai(targetTimeStr);
  let [th, tm] = targetParsed.split(":").map(Number);
  if (isNaN(th)) th = 8;
  if (isNaN(tm)) tm = 0;

  const currentMin = h * 60 + m;
  const targetMin = th * 60 + tm;
  const diff = currentMin - targetMin;
  const tol = (toleranceMinutes !== undefined && toleranceMinutes !== null && !isNaN(toleranceMinutes)) ? Number(toleranceMinutes) : 15;

  if (diff <= tol) {
    return { 
      status: "TEPAT WAKTU", 
      lateMinutes: 0,
      keterlambatanText: "-"
    };
  } else {
    return { 
      status: "TERLAMBAT", 
      lateMinutes: diff,
      keterlambatanText: `${diff} Menit`
    };
  }
}

function parseParticipantName(rawName) {
  if (!rawName) return null;
  let clean = String(rawName).trim();

  // 1. Abaikan header/jumlah peserta (e.g. "Participants (1)", "Peserta (2)")
  if (/^(participants|peserta)\s*\(/i.test(clean)) return null;

  // 2. Abaikan teks yang mengandung tombol Zoom
  if (/invite|mute all|mute|unmute|whiteboards|screen share/i.test(clean)) return null;

  // 3. Batas panjang dan format (minimal 2 karakter, maksimal 70 karakter)
  if (clean.includes("\n") || clean.length < 2 || clean.length > 70) return null;

  // 4. Pastikan memiliki minimal 2 huruf alfabet
  const alphaOnly = clean.replace(/[^a-zA-Z]/g, '');
  if (alphaOnly.length < 2) return null;

  const lower = clean.toLowerCase();
  const junkWords = [
    "more_vert", "visual_effects", "frame_person", "lock_person", "call_end", 
    "mic", "videocam", "audio", "video", "settings", "security", "preview"
  ];
  for (let j of junkWords) {
    if (lower.includes(j)) return null;
  }

  // 5. Bersihkan badge role
  clean = clean.replace(/\((Anda|You|Host|Penyelenggara rapat|Penyelenggara|Co-host|Host,\s*me|Me|Guest|Tamu)\)/gi, "").trim();

  // 6. Hapus tombol yang menempel
  clean = clean.replace(/\b(invite|mute all|more|mute)\b/gi, "").trim();

  if (!clean || clean.length < 2) return null;

  const knownDeptMap = {
    // 14 Departemen Resmi & Aliases
    "ebs": "EBS",
    "solution architect": "Solution Architect",
    "solution architecture": "Solution Architect",
    "architect": "Solution Architect",
    "sa": "Solution Architect",
    "could automation": "Could Automation",
    "cloud automation": "Could Automation",
    "ca": "Could Automation",
    "research and innovation": "Research and innovation",
    "research & innovation": "Research and innovation",
    "rni": "Research and innovation",
    "r&i": "Research and innovation",
    "research": "Research and innovation",
    "infrastructur analytic": "Infrastructur Analytic",
    "infrastructure analytic": "Infrastructur Analytic",
    "infrastructure analytics": "Infrastructur Analytic",
    "infrastruktur analitik": "Infrastructur Analytic",
    "ia": "Infrastructur Analytic",
    "fat": "FAT",
    "finance, accounting & tax": "FAT",
    "finance": "FAT",
    "keuangan": "FAT",
    "accounting": "FAT",
    "tax": "FAT",
    "pajak": "FAT",
    "sales": "Sales",
    "marketing": "Sales",
    "pemasaran": "Sales",
    "bd": "Sales",
    "business development": "Sales",
    "cyber cecurity": "Cyber Cecurity",
    "cyber security": "Cyber Cecurity",
    "security": "Cyber Cecurity",
    "keamanan siber": "Cyber Cecurity",
    "cs": "Cyber Cecurity",
    "pmo": "PMO",
    "project management": "PMO",
    "bod": "BOD",
    "board of directors": "BOD",
    "direksi": "BOD",
    "modern application": "Modern Application",
    "modern applications": "Modern Application",
    "modern app": "Modern Application",
    "modern apps": "Modern Application",
    "ma": "Modern Application",
    "data analytics": "Data Analytics",
    "data analytic": "Data Analytics",
    "data analyst": "Data Analytics",
    "data science": "Data Analytics",
    "da": "Data Analytics",
    "hcmc": "HCMC",
    "hrd": "HCMC",
    "hr": "HCMC",
    "human capital": "HCMC",
    "personalia": "HCMC",
    "pkl/magang": "PKL/Magang",
    "pkl": "PKL/Magang",
    "magang": "PKL/Magang",
    "intern": "PKL/Magang",
    "internship": "PKL/Magang",
    "it": "IT",
    "ti": "IT",
    "legal": "Legal",
    "operasional": "Operasional",
    "ops": "Operasional",
    "admin": "Admin",
    "umum": "Umum"
  };

  function normalizeDept(str) {
    if (!str) return null;
    const raw = str.trim().toLowerCase();
    if (knownDeptMap[raw]) return knownDeptMap[raw];
    const cleaned = raw.replace(/[\s\-_]+/g, ' ');
    if (knownDeptMap[cleaned]) return knownDeptMap[cleaned];
    return null;
  }

  // Pola 1: [Departemen] Nama Lengkap
  const bracketPrefix = clean.match(/^[\[\{]([^\]\}]+)[\]\}]\s*[-–—:]?\s*(.+)$/);
  if (bracketPrefix) {
    const rawDept = bracketPrefix[1].trim();
    return {
      department: normalizeDept(rawDept) || rawDept,
      name: bracketPrefix[2].trim()
    };
  }

  // Pola 2: Nama Lengkap [Departemen]
  const bracketSuffix = clean.match(/^(.+?)\s*[-–—:]?\s*[\[\{]([^\]\}]+)[\]\}]$/);
  if (bracketSuffix) {
    const rawDept = bracketSuffix[2].trim();
    return {
      name: bracketSuffix[1].trim(),
      department: normalizeDept(rawDept) || rawDept
    };
  }

  // Pola 3: (Departemen) Nama Lengkap
  const parenPrefix = clean.match(/^\(([^)]+)\)\s*[-–—:]?\s*(.+)$/);
  if (parenPrefix) {
    const rawDept = parenPrefix[1].trim();
    return {
      department: normalizeDept(rawDept) || rawDept,
      name: parenPrefix[2].trim()
    };
  }

  // Pola 4: Nama Lengkap (Departemen)
  const parenMatch = clean.match(/^(.+?)\s*\(([^)]+)\)$/);
  if (parenMatch) {
    const p1 = parenMatch[1].trim();
    const p2 = parenMatch[2].trim();
    if (normalizeDept(p2)) {
      return { name: p1, department: normalizeDept(p2) };
    } else if (normalizeDept(p1)) {
      return { department: normalizeDept(p1), name: p2 };
    } else {
      return { name: p1, department: p2 };
    }
  }

  // Pola 5: Separator -, –, —, |, /, :, _
  if (/[-–—|:_/]/.test(clean)) {
    const parts = clean.split(/[-–—|:_]/).map(function(s) { return s.trim(); }).filter(Boolean);
    if (parts.length >= 2) {
      const p1 = parts[0];
      const p2 = parts.slice(1).join(' ').trim();

      const p1Dept = normalizeDept(p1);
      const p2Dept = normalizeDept(p2);

      if (p1Dept && !p2Dept) {
        return { department: p1Dept, name: p2 };
      }
      if (p2Dept && !p1Dept) {
        return { name: p1, department: p2Dept };
      }

      // Heuristik akronim (misal FAT, PMO, BOD, MA, CA, SA, RNI)
      const p1IsAcronym = /^[A-Za-z0-9&]{2,5}$/.test(p1) && (p1 === p1.toUpperCase() || p1.length <= 4);
      const p2IsAcronym = /^[A-Za-z0-9&]{2,5}$/.test(p2) && (p2 === p2.toUpperCase() || p2.length <= 4);

      if (p1IsAcronym && !p2IsAcronym) {
        return { department: p1, name: p2 };
      }
      if (p2IsAcronym && !p1IsAcronym) {
        return { name: p1, department: p2 };
      }

      // Standar penamaan meeting kantor: [Departemen] - [Nama Lengkap]
      return { department: p1, name: p2 };
    }
  }

  return { name: clean, department: "Umum" };
}
