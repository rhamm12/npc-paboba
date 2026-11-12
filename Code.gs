/**
 * NEXT PAGE CHALLENGE (NPC) by PABOBA
 * File: Code.gs (Google Apps Script Backend - HIGH PERFORMANCE OPTIMIZED)
 * Deskripsi: Mengelola penyimpanan database terpusat secara aman, cepat, dan real-time.
 */

function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * doGet:
 *  - ?action=load  -> mengembalikan seluruh database sebagai JSON (dipakai versi Vercel)
 *  - tanpa action  -> menampilkan halaman index (mode lama langsung di Apps Script)
 */
function doGet(e) {
  const action = e && e.parameter ? e.parameter.action : "";

  if (action === "load") {
    try {
      return jsonOutput({ ok: true, data: getDatabase() });
    } catch (err) {
      return jsonOutput({ ok: false, error: String(err) });
    }
  }

  try {
    return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Next Page Challenge (NPC) - PABOBA')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return jsonOutput({ ok: true, message: "NPC API aktif. Gunakan ?action=load untuk membaca data." });
  }
}

/**
 * doPost:
 *  - body JSON (dikirim sebagai text/plain): { action: "save", state: {...} }
 *  - dilindungi LockService agar dua penyimpanan bersamaan tidak saling menimpa
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);

    const raw = e && e.postData ? e.postData.contents : "";
    const body = JSON.parse(raw || "{}");

    if (body.action !== "save" || !body.state) {
      return jsonOutput({ ok: false, error: "Permintaan tidak valid" });
    }

    const st = body.state;
    if (!Array.isArray(st.participants) || !Array.isArray(st.laporan) ||
        !Array.isArray(st.partners) || !Array.isArray(st.piketPj)) {
      return jsonOutput({ ok: false, error: "Struktur data tidak lengkap" });
    }

    saveStateToSheet(st);
    return jsonOutput({ ok: true });
  } catch (err) {
    return jsonOutput({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function formatDateString(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(val);
}

function getSpreadsheetRobust() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  
  const id = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  if (id) {
    return SpreadsheetApp.openById(id);
  }
  
  throw new Error("Spreadsheet tidak terdeteksi! Pastikan script ini dibuat dari dalam Google Sheets.");
}

function getDatabase() {
  const ss = getSpreadsheetRobust();
  const db = {
    participants: [],
    laporan: [],
    partners: [],
    piketPj: [],
    securityPin: "123456",
    startDate: "2026-05-01",
    endDate: "2026-05-31"
  };

  try {
    // 1. Ambil atau Inisialisasi Sheet "Peserta"
    let sheetPeserta = ss.getSheetByName("Peserta");
    if (!sheetPeserta) {
      sheetPeserta = ss.insertSheet("Peserta");
      sheetPeserta.appendRow(["Nama Peserta"]);
      const defaultNames = [
        "Fatih", "Devi", "Miah", "Ida", "Vina", "Septi", "Rii", "Wari", "Nurjanah", "Yola", 
        "Umi", "Annisa", "Salma", "Rezki", "Azmi", "Bii", "Chairunnisa", "Mawar", "Suro", "Nuha", 
        "Immaduddin", "Hikma", "Fadhilah", "Dino", "Emmy", "Icca", "Umy", "Eka", "Izzah", "Mimi", 
        "Agung", "Dimas", "Jirfani", "Indra", "Dian", "Anita", "Sukarini", "Kartika", "Hasanah", 
        "Amelia", "Putri", "Idha Sasongko", "Muslimah Annisa", "Fajriy", "Dindsan", "Aqilah", "Nur", 
        "Melly", "Dika Septiamy"
      ];
      defaultNames.forEach(name => sheetPeserta.appendRow([name]));
    }
    const dataPeserta = sheetPeserta.getDataRange().getValues();
    for (let i = 1; i < dataPeserta.length; i++) {
      if (dataPeserta[i][0]) db.participants.push(dataPeserta[i][0]);
    }

    // 2. Ambil atau Inisialisasi Sheet "Laporan"
    let sheetLaporan = ss.getSheetByName("Laporan");
    if (!sheetLaporan) {
      sheetLaporan = ss.insertSheet("Laporan");
      sheetLaporan.appendRow(["id", "date", "name", "status", "insight", "time", "insightText", "bookTitle", "pageRange", "author"]);
    }
    const dataLaporan = sheetLaporan.getDataRange().getValues();
    for (let i = 1; i < dataLaporan.length; i++) {
      if (dataLaporan[i][0]) {
        db.laporan.push({
          id: dataLaporan[i][0],
          date: formatDateString(dataLaporan[i][1]),
          name: dataLaporan[i][2],
          status: dataLaporan[i][3],
          insight: dataLaporan[i][4],
          time: dataLaporan[i][5],
          insightText: dataLaporan[i][6] ? String(dataLaporan[i][6]) : "",
          bookTitle: dataLaporan[i][7] ? String(dataLaporan[i][7]) : "",
          pageRange: (dataLaporan[i][8] && !(dataLaporan[i][8] instanceof Date)) ? String(dataLaporan[i][8]) : "",
          author: dataLaporan[i][9] ? String(dataLaporan[i][9]) : ""
        });
      }
    }

    // 3. Ambil atau Inisialisasi Sheet "Partners"
    let sheetPartners = ss.getSheetByName("Partners");
    if (!sheetPartners) {
      sheetPartners = ss.insertSheet("Partners");
      sheetPartners.appendRow(["type", "name1", "name2", "hp", "active", "notes"]);
    }
    const dataPartners = sheetPartners.getDataRange().getValues();
    const headersPartners = dataPartners[0];
    const hasNotesCol = headersPartners.indexOf("notes") !== -1;

    for (let i = 1; i < dataPartners.length; i++) {
      if (dataPartners[i][0]) {
        db.partners.push({
          type: dataPartners[i][0],
          name1: dataPartners[i][1],
          name2: dataPartners[i][2],
          hp: Number(dataPartners[i][3]),
          active: dataPartners[i][4] === "true" || dataPartners[i][4] === true || dataPartners[i][4] === "TRUE",
          notes: hasNotesCol ? String(dataPartners[i][headersPartners.indexOf("notes")]) : ""
        });
      }
    }

    // 4. Ambil atau Inisialisasi Sheet "Piket"
    let sheetPiket = ss.getSheetByName("Piket");
    if (!sheetPiket) {
      sheetPiket = ss.insertSheet("Piket");
      sheetPiket.appendRow(["id", "date", "day", "name"]);
    }
    const dataPiket = sheetPiket.getDataRange().getValues();
    const headersPiket = dataPiket[0];
    const hasDateCol = headersPiket.indexOf("date") !== -1;

    for (let i = 1; i < dataPiket.length; i++) {
      if (dataPiket[i][0]) {
        let dateVal = "";
        let dayVal = "";
        let nameVal = "";
        if (hasDateCol) {
          dateVal = formatDateString(dataPiket[i][headersPiket.indexOf("date")]);
          dayVal = dataPiket[i][headersPiket.indexOf("day")];
          nameVal = dataPiket[i][headersPiket.indexOf("name")];
        } else {
          dayVal = dataPiket[i][1];
          nameVal = dataPiket[i][2];
          dateVal = ""; 
        }
        db.piketPj.push({
          id: dataPiket[i][0],
          date: dateVal,
          day: dayVal,
          name: nameVal
        });
      }
    }

    // 5. Ambil atau Inisialisasi Sheet "Pengaturan"
    let sheetPengaturan = ss.getSheetByName("Pengaturan");
    if (!sheetPengaturan) {
      sheetPengaturan = ss.insertSheet("Pengaturan");
      sheetPengaturan.appendRow(["Kunci", "Nilai"]);
      sheetPengaturan.appendRow(["securityPin", "123456"]);
      sheetPengaturan.appendRow(["startDate", "2026-05-01"]);
      sheetPengaturan.appendRow(["endDate", "2026-05-31"]);
    }
    const dataPengaturan = sheetPengaturan.getDataRange().getValues();
    for (let i = 1; i < dataPengaturan.length; i++) {
      if (dataPengaturan[i][0] === "securityPin") db.securityPin = String(dataPengaturan[i][1]);
      if (dataPengaturan[i][0] === "startDate") db.startDate = formatDateString(dataPengaturan[i][1]);
      if (dataPengaturan[i][0] === "endDate") db.endDate = formatDateString(dataPengaturan[i][1]);
    }

  } catch (error) {
    Logger.log("Kesalahan saat memuat spreadsheet: " + error.toString());
  }

  return db;
}

function saveStateToSheet(state) {
  const ss = getSpreadsheetRobust();
  
  // 1. Perbarui Sheet "Peserta"
  let sheetPeserta = ss.getSheetByName("Peserta") || ss.insertSheet("Peserta");
  sheetPeserta.clearContents();
  let rowsPeserta = [["Nama Peserta"]];
  state.participants.forEach(name => {
    if (name) rowsPeserta.push([name]);
  });
  sheetPeserta.getRange(1, 1, rowsPeserta.length, rowsPeserta[0].length).setValues(rowsPeserta);

  // 2. Perbarui Sheet "Laporan"
  let sheetLaporan = ss.getSheetByName("Laporan") || ss.insertSheet("Laporan");
  sheetLaporan.clearContents();
  let rowsLaporan = [["id", "date", "name", "status", "insight", "time", "insightText", "bookTitle", "pageRange", "author"]];
  state.laporan.forEach(log => {
    if (log.id) {
      rowsLaporan.push([log.id, log.date, log.name, log.status, log.insight, log.time, log.insightText || "", log.bookTitle || "", log.pageRange || "", log.author || ""]);
    }
  });
  // Kolom "pageRange" (kolom 9) dipaksa berformat teks agar "10-20" tidak berubah jadi tanggal
  sheetLaporan.getRange(1, 9, sheetLaporan.getMaxRows(), 1).setNumberFormat("@");
  sheetLaporan.getRange(1, 1, rowsLaporan.length, rowsLaporan[0].length).setValues(rowsLaporan);

  // 3. Perbarui Sheet "Partners"
  let sheetPartners = ss.getSheetByName("Partners") || ss.insertSheet("Partners");
  sheetPartners.clearContents();
  let rowsPartners = [["type", "name1", "name2", "hp", "active", "notes"]];
  state.partners.forEach(partner => {
    if (partner.type) {
      rowsPartners.push([partner.type, partner.name1, partner.name2, partner.hp, String(partner.active), partner.notes || ""]);
    }
  });
  sheetPartners.getRange(1, 1, rowsPartners.length, rowsPartners[0].length).setValues(rowsPartners);

  // 4. Perbarui Sheet "Piket"
  let sheetPiket = ss.getSheetByName("Piket") || ss.insertSheet("Piket");
  sheetPiket.clearContents();
  let rowsPiket = [["id", "date", "day", "name"]];
  state.piketPj.forEach(piket => {
    if (piket.id) {
      rowsPiket.push([piket.id, piket.date || "", piket.day || "", piket.name]);
    }
  });
  sheetPiket.getRange(1, 1, rowsPiket.length, rowsPiket[0].length).setValues(rowsPiket);

  // 5. Perbarui Sheet "Pengaturan"
  let sheetPengaturan = ss.getSheetByName("Pengaturan") || ss.insertSheet("Pengaturan");
  sheetPengaturan.clearContents();
  let rowsPengaturan = [
    ["Kunci", "Nilai"],
    ["securityPin", state.securityPin],
    ["startDate", state.startDate],
    ["endDate", state.endDate]
  ];
  sheetPengaturan.getRange(1, 1, rowsPengaturan.length, rowsPengaturan[0].length).setValues(rowsPengaturan);

  // 6. Perbarui Sheet "Gugur" (Otomatis menyaring klan gugur & resign secara real-time)
  let sheetGugur = ss.getSheetByName("Gugur") || ss.insertSheet("Gugur");
  sheetGugur.clearContents();
  let rowsGugur = [["Nama Peserta Gugur", "Status", "Keterangan"]];
  
  let disqNames = [];
  state.partners.forEach(team => {
    let isInactive = false;
    if (team.active === false || team.active === "false" || team.active === "FALSE") {
      isInactive = true;
    } else if (team.active === true || team.active === "true" || team.active === "TRUE") {
      isInactive = false;
    } else {
      isInactive = !team.active;
    }

    if (isInactive) {
      if (team.name1 && team.name1 !== "-") {
        let reason = "HP Habis / Tereliminasi Otomatis dari Batch";
        if (team.notes) {
          reason = team.notes;
        } else if (team.hp === 0) {
          reason = "Mengundurkan Diri (Resign)";
        }
        disqNames.push({ name: team.name1, reason: reason });
      }
      if (team.name2 && team.name2 !== "-") {
        let reason = "HP Habis / Tereliminasi Otomatis dari Batch";
        if (team.notes) {
          reason = team.notes;
        }
        disqNames.push({ name: team.name2, reason: reason });
      }
    }
  });
  
  disqNames.forEach(item => {
    rowsGugur.push([item.name, "Gugur", item.reason]);
  });
  
  if (rowsGugur.length > 1) {
    sheetGugur.getRange(1, 1, rowsGugur.length, rowsGugur[0].length).setValues(rowsGugur);
  } else {
    sheetGugur.appendRow(["Belum ada peserta gugur", "-", "-"]);
  }

  return true;
}