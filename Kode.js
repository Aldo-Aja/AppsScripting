/**
 * KONFIGURASI & KONSTANTA
 * Menggunakan engine 'IVE' yang dioptimalkan, tapi disesuaikan tanpa Rapor.
 */
const CFG = {
  SHEET_SNBP: "SNBP",
  SHEET_ENGINE: "Engine",
  SHEET_MAP: "Mapping",
  SHEET_SIN: "MappingSinonim",
  SHEET_MAPEL: "MappingMapel",
  COL_OUT_DATA: 27, // Kolom AA (Data mentah untuk VLOOKUP/Indeks)
  COL_OUT_TEXT: 32, // Kolom AF (Teks Hasil Rekomendasi)
  PTN_PSIKO_SAINTEK: new Set(['UNPAD','UNS','UNHAS','UNAND','UNSRI']),
  OWNER: 'revaldoparikesit@gmail.com'
};

/**
 * 1. TRIGGER OTOMATIS
 * Dimodifikasi: Menghapus fungsi reset nilai rapor (baris 18++).
 */
function onEdit(e) {
  if (!e || e.source.getSheetName() !== CFG.SHEET_SNBP) return;
  
  const r = e.range.rowStart;
  const c = e.range.columnStart;
  const sh = e.range.getSheet();

  // A. Trigger: Reset Jurusan (Edit di C8)
  if (r === 8 && c === 3) {
    // 1. Hapus Jurusan (C9)
    sh.getRange("C9").clearContent();
    
    // 2. Bersihkan Output Mapel (J8, J9 - Sesuai request Code A)
    sh.getRangeList(["J8", "J9"]).setValue("Belum Ditentukan");
    
    // 3. Clear Output Engine Cepat (AA & AF)
    const shEng = e.source.getSheetByName(CFG.SHEET_ENGINE);
    if (shEng) {
      const last = getLastRowInColumn(shEng, CFG.COL_OUT_DATA);
      if (last > 1) {
        shEng.getRange(2, CFG.COL_OUT_DATA, last, 10).clearContent();
        shEng.getRange(2, CFG.COL_OUT_TEXT, last, 2).clearContent();
      }
    }
    return;
  }

  // B. Trigger: Ganti Jurusan (C9) atau PTN (M11)
  if ((r === 9 && c === 3) || (r === 11 && c === 13)) {
    rekomendasiKampus();
    rekomendasiMapel();
  }
}

/**
 * 2. CORE LOGIC: REKOMENDASI KAMPUS (High Performance)
 * Menggunakan sistem Code B tapi mengambil rasio dari C14
 */
function rekomendasiKampus() {
  const ss = SpreadsheetApp.getActive();
  const shSNBP = ss.getSheetByName(CFG.SHEET_SNBP);
  const shEng  = ss.getSheetByName(CFG.SHEET_ENGINE);
  
  // Baca Input Sekaligus (Batch Read)
  const inputVals = shSNBP.getRange("C9:M14").getValues(); 
  const jurInputRaw = inputVals[0][0]; // C9
  const ptnAsal     = inputVals[2][10]; // M11
  const rasioRaw    = inputVals[5][0]; // C14 (Rasio Indeks Sekolah/Alumni)

  if (!jurInputRaw || !ptnAsal) return;

  const rasioAsal = parseNum(rasioRaw);
  let jurInput = norm(jurInputRaw);
  const inputJenjang = getJenjangGroup(jurInputRaw);

  // Normalisasi Nama Jurusan Khusus
  if (jurInput.includes("produksi media")) 
    jurInput = "produksi media komunikasi penyiaran televisi film multimedia";

  // --- PRE-LOAD DATA (CACHE) ---
  const uinData = getUINAndMapData(ss);
  const synonyms = getSynonyms(ss, jurInput);
  
  // Flagging Filter
  const isGigi   = jurInput.includes("gigi");
  const isHewan  = jurInput.includes("hewan");
  const isDokter = (jurInput.includes("dokter") || jurInput.includes("kedokteran")) && !isGigi && !isHewan;
  const isTeknik = jurInput.startsWith("teknik");
  const isInputUIN = uinData.codes.has(ptnAsal.toString().toLowerCase());

  // --- ENGINE PROCESSING ---
  const lastRowEng = shEng.getLastRow();
  if (lastRowEng < 2) return;
  
  // Ambil database engine (A2:I)
  const data = shEng.getRange(2, 1, lastRowEng - 1, 9).getValues(); 
  
  let results = [];
  
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    // Filter Cepat
    if (!r[3] || !r[4]) continue; 
    if (r[3].toString().toLowerCase() === ptnAsal.toString().toLowerCase()) continue; // Skip PTN Asal
    if (inputJenjang !== getJenjangGroup(r[5])) continue; // Skip beda jenjang (D3 vs S1)
    
    const rasio = parseNum(r[8]);
    if (rasio === 0) continue;

    const jurTarget = norm(r[4]);

    // Filter Strict Keywords
    if (isTeknik !== jurTarget.startsWith("teknik")) continue;
    if (isGigi && !jurTarget.includes("gigi")) continue;
    if (!isGigi && jurTarget.includes("gigi")) continue;
    if (isHewan && !jurTarget.includes("hewan")) continue;
    if (!isHewan && jurTarget.includes("hewan")) continue;
    if (isDokter && (!jurTarget.includes("dokter") && !jurTarget.includes("kedokteran"))) continue;

    // Filter Berat (Regex/Sinonim)
    if (!isJurusanMatch(jurInput, jurTarget, synonyms)) continue;

    // Logika Rasio (Setara/Longgar)
    const diff = rasio - rasioAsal;
    if ((diff >= -0.02 && diff <= 0.90) || (rasioAsal === 0 && diff > 0)) {
       const isUinTarget = uinData.codes.has(r[3].toString().toLowerCase());
       
       // PERUBAHAN DI SINI:
       // Menggunakan r[4] langsung (Nama Jurusan Asli) tanpa formatJurusanDisplay
       const namaJurusanFinal = r[4]; 
       
       results.push({
         row: [r[3], r[4], r[6], r[7], r[8]],
         text: `${namaJurusanFinal} - ${uinData.map[r[3]] || r[3]} (${(rasio * 100).toFixed(2).replace('.', ',')}%)`,
         val: rasio,
         isUin: isUinTarget
       });
    }
  }

  // Sorting: Prioritaskan UIN jika input UIN, lalu urutkan berdasarkan ketetatan/rasio
  results.sort((a, b) => {
    if (isInputUIN && a.isUin !== b.isUin) return a.isUin ? -1 : 1;
    return a.val - b.val;
  });

  // --- BATCH WRITING (Kecepatan Tinggi) ---
  const limit = Math.min(results.length, 10); // Ambil Top 10
  const outData = [];
  const outText = [];

  for (let i = 0; i < limit; i++) {
    outData.push(results[i].row);
    outText.push([results[i].text]);
  }

  // Tulis ke Engine (AA dan AF)
  if (limit > 0) {
    shEng.getRange(2, CFG.COL_OUT_DATA, limit, 5).setValues(outData);
    shEng.getRange(2, CFG.COL_OUT_TEXT, limit, 1).setValues(outText);
  }

  // Bersihkan Sisa Data Lama
  const lastOut = getLastRowInColumn(shEng, CFG.COL_OUT_DATA);
  const rowsToClean = lastOut - (limit + 1);
  
  if (rowsToClean > 0) {
    const safeClean = Math.min(rowsToClean, 50); 
    shEng.getRange(2 + limit, CFG.COL_OUT_DATA, safeClean, 5).clearContent();
    shEng.getRange(2 + limit, CFG.COL_OUT_TEXT, safeClean, 1).clearContent();
  }
}

/**
 * 3. LOGIC: REKOMENDASI MAPEL
 * Output dikembalikan ke J8 dan J9 (Sesuai Kode A)
 */
function rekomendasiMapel() {
  const ss = SpreadsheetApp.getActive();
  const shSNBP = ss.getSheetByName(CFG.SHEET_SNBP);
  const shMapel = ss.getSheetByName(CFG.SHEET_MAPEL);
  if (!shMapel) return;

  const jurRaw = shSNBP.getRange("C9").getValue();
  const ptnKode = shSNBP.getRange("M11").getValue();

  if (!jurRaw) {
    shSNBP.getRangeList(['J8','J9']).setValue("Belum Ditentukan");
    return;
  }

  let jurNorm = resolveKimiaContext(norm(jurRaw));
  
  // Konteks Khusus (Psikologi/Kedokteran)
  if (jurNorm.includes('psikologi')) {
    jurNorm = CFG.PTN_PSIKO_SAINTEK.has(ptnKode) ? 'psikologi saintek' : 'psikologi soshum';
  } else if (jurNorm.includes('dokter') || jurNorm.includes('kedokteran')) {
    if (jurNorm.includes('gigi')) jurNorm = 'kedokteran gigi';
    else if (jurNorm.includes('hewan')) jurNorm = 'kedokteran hewan';
    else jurNorm = 'kedokteran';
  }

  const data = shMapel.getDataRange().getValues();
  const scores = {};

  // Hitung Skor
  for (let i = 1; i < data.length; i++) { 
    const key = norm(data[i][0]);
    if (!key || !jurNorm.includes(key)) continue;
    
    const mapel = data[i][1];
    scores[mapel] = (scores[mapel] || 0) + Number(data[i][2]);
  }

  // Bonus Points Logic
  if (jurNorm.match(/indonesia.*(bahasa|sastra)/)) scores['Bahasa Indonesia'] = 1000;
  else if (jurNorm.includes("pendidikan")) scores['Bahasa Indonesia'] = (scores['Bahasa Indonesia'] || 0) + 50;

  // Default Fallback
  if (Object.keys(scores).length === 0) {
    if (isBioKimiaRumpun(jurNorm) && !jurNorm.includes('teknik kimia')) {
      scores['Biologi'] = 10; scores['Kimia'] = 9;
    } else if (jurNorm.includes('teknik')) {
      scores['Matematika Tingkat Lanjut'] = 10; scores['Fisika'] = 9;
    }
  }

  const sorted = Object.keys(scores).sort((a, b) => scores[b] - scores[a]);
  const final = [...new Set([...sorted, 'Bahasa Indonesia', 'Matematika Tingkat Lanjut'])];

  // OUTPUT KE J8 & J9 (Sesuai permintaan Code A)
  shSNBP.getRange("J8").setValue(final[0]);
  shSNBP.getRange("J9").setValue(final[1]);
}

/**
 * UTILITIES & HELPER (JANGAN DIHAPUS - INI ENGINE PENCARINYA)
 */
function getLastRowInColumn(sheet, colIndex) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const range = sheet.getRange(1, colIndex, lastRow, 1).getValues();
  for (let i = lastRow - 1; i >= 0; i--) {
    if (range[i][0] !== "" && range[i][0] != null) return i + 1;
  }
  return 0;
}

function getUINAndMapData(ss) {
  const sh = ss.getSheetByName(CFG.SHEET_MAP);
  const res = { codes: new Set(), map: {} };
  if (!sh) return res;
  const data = sh.getRange("A2:B" + sh.getLastRow()).getValues();
  for (let r of data) {
    if (r[1]) {
      const k = r[1].toString();
      res.map[k] = r[0];
      if (r[0].toString().toUpperCase().match(/UIN |IAIN |STAIN |ISLAM NEGERI/)) res.codes.add(k.toLowerCase());
    }
  }
  return res;
}

function getSynonyms(ss, jurInput) {
  const sh = ss.getSheetByName(CFG.SHEET_SIN);
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  for (let row of data) {
    if (row.join(" ").toLowerCase().includes(jurInput)) {
      return row.filter(String).map(s => norm(s));
    }
  }
  return [];
}

function norm(text) {
  if (!text) return "";
  return text.toString().toLowerCase()
    .replace(/\(.*?\)/g, "")
    .replace(/[^a-z0-9 ]/g, " ")
    .replace(/\b(prodi|program|studi|jurusan|ilmu|fakultas|sekolah|departemen|dan|of|and|the)\b/g, "")
    .replace(/\s+/g, " ").trim();
}

function getJenjangGroup(text) {
  const s = String(text).toUpperCase();
  return (s.match(/(D3|D4|DIPLOMA|TERAPAN)/)) ? "VOKASI" : "S1";
}

function formatJurusanDisplay(nama, jenjang) {
  const j = String(jenjang).trim().toUpperCase();
  let p = "";
  if (j === "SARJANA" || j === "S1") p = "(S1)";
  else if (j.includes("DIPLOMA TIGA") || j === "D3") p = "(D3)";
  else if (j.includes("SARJANA TERAPAN") || j === "D4") p = "(D4)";
  return p ? p + " " + nama : nama;
}

function parseNum(val) {
  if (!val) return 0;
  if (typeof val === 'number') return (val > 1) ? val / 100 : val;
  const str = val.toString().replace('%','').replace(',','.').trim();
  const num = parseFloat(str);
  return (val.toString().includes('%')) ? num / 100 : num;
}

function resolveKimiaContext(jur) {
  if (jur.includes('kimia') && jur.includes('pendidikan')) return 'pendidikan kimia';
  if (jur.includes('kimia') && jur.includes('teknik')) return 'teknik kimia';
  return jur.includes('kimia') ? 'kimia' : jur;
}

function isBioKimiaRumpun(jur) {
  return /kehutanan|kelautan|peternakan|perikanan|perairan|pertanian|agro|agribisnis|agrikultur|teknologi pangan|pangan|bio|biologi|bioproses|hayati|kesehatan|gizi|kultur/.test(jur);
}

function isJurusanMatch(input, target, synonyms) {
  // 1. CEK SINONIM DULU (Prioritas Utama)
  // Jika target mengandung salah satu kata dari mapping sinonim, langsung TRUE
  if (synonyms.length > 0) {
    if (synonyms.some(s => target.includes(s) || s.includes(target))) return true;
  }

  // 2. Filter Blacklist (Tetap ada untuk akurasi)
  if (input.includes('manajemen')) {
    if (target.match(/hutan|informatika|sumber|perairan|pendidikan|rekayasa|industri/)) return false;
  }
  
  // 3. Pengecekan Kata Per Kata
  const words = input.split(" ").filter(w => w.length > 2);
  if (words.length === 0) return false;
  
  let matches = 0;
  for (let w of words) if (target.includes(w)) matches++;
  
  // Jika 50% kata cocok, anggap masuk
  return (matches / words.length) >= 0.5;
}
// AUTH
function isAuthorizedUser_() {
  const lic = SpreadsheetApp.getActive().getSheetByName('LICENSE').getRange('A1').getValue();
  const email = Session.getActiveUser().getEmail();
  return (email === CFG.OWNER || email === lic);
}
function doPost(e) {
  if (!e || !e.postData) return ContentService.createTextOutput("Error");
  const d = JSON.parse(e.postData.contents);
  SpreadsheetApp.openById(d.spreadsheetId).getSheetByName('LICENSE').getRange('A1').setValue(d.email);
  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }));
}
