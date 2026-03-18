/**
 * =========================================================================
 * BACKEND GOOGLE APPS SCRIPT - E-PLANNING DINKES
 * =========================================================================
 * Silakan copy-paste seluruh kode ini ke dalam file Code.gs di project GAS Anda.
 */

// NAMA-NAMA SHEET SESUAI DATABASE ANDA
const SHEET_USER = "Master_User";
const SHEET_PENGATURAN = "Pengaturan_Sistem";
const SHEET_USULAN = "Data_Usulan";
const SHEET_RINCIAN = "Rincian_Usulan";
const SHEET_REKENING = "Master_Rekening";
const SHEET_SUBKEG = "Master_SubKegiatan";
const SHEET_DANA = "Master_SumberDana";
const SHEET_SATUAN = "Master_Satuan";

// Peta Kolom (Memastikan key konsisten antara Frontend dan Backend)
const COLUMN_MAP = {
  "id_usulan": "ID_Usulan",
  "bidang": "Bidang",
  "sub_kegiatan": "Sub_Kegiatan",
  "indikator": "Indikator",
  "target": "Target",
  "total_anggaran": "Total_Anggaran",
  "status": "Status",
  "pembuat": "Pembuat", // Track siapa pengusulnya (Level Program)
  "id_rincian": "ID_Rincian",
  "kode_rekening": "Kode_Rekening",
  "nama_rekening": "Nama_Rekening",
  "sumber_dana": "Sumber_Dana",
  "komponen": "Komponen",
  "spesifikasi": "Spesifikasi",
  "keterangan": "Keterangan",
  "koefisien": "Koefisien",
  "volume": "Volume",
  "harga_satuan": "Harga_Satuan",
  "sub_total": "Sub_Total",
  "status_item": "Status_Item",
  "catatan": "Catatan",
  "username": "Username",
  "nama": "Nama", // Kolom Nama Lengkap User
  "password": "Password",
  "level_akses": "Level_Akses", // Level user (Level Program / Level Bidang / Admin)
  "kode_subkegiatan": "Kode_SubKegiatan",
  "nama_subkegiatan": "Nama_SubKegiatan",
  "satuan": "Satuan",
  "kode": "Kode"
};

// ==========================================
// CORE SYSTEM UTILITIES
// ==========================================
function getDb() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("Script belum di-bind ke Spreadsheet.");
  }
  return ss;
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('E-Planning Dinkes')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDataFromSheet(sheetName) {
  try {
    const sheet = getDb().getSheetByName(sheetName);
    if (!sheet) return []; 
    
    // Menggunakan getDisplayValues() untuk mencegah Error Crash akibat format Tanggal (Date) di Sheets
    const data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) return []; 
    
    const headers = data[0].map(h => String(h).trim().replace(/\r/g, ''));
    const result = [];
    
    for (let i = 1; i < data.length; i++) {
      let obj = {};
      for (let j = 0; j < headers.length; j++) {
        if(headers[j]) obj[headers[j]] = data[i][j];
      }
      result.push(obj);
    }
    return result;
  } catch (e) {
    console.error("Error reading " + sheetName + ": " + e);
    return []; 
  }
}

function generateId(prefix) {
  return prefix + "-" + new Date().getTime() + "-" + Math.floor(Math.random() * 1000);
}

// ==========================================
// FUNGSI SETUP DATABASE OTOMATIS
// ==========================================
function setupDatabase() {
  const ss = getDb();
  // Struktur Header Terbaru
  const schemas = {
    [SHEET_USER]: ["Username", "Nama", "Password", "Level_Akses", "Bidang"],
    [SHEET_PENGATURAN]: ["Parameter (Key)", "Nilai (Value)"],
    [SHEET_USULAN]: ["ID_Usulan", "Bidang", "Pembuat", "Sub_Kegiatan", "Indikator", "Target", "Total_Anggaran", "Status"],
    [SHEET_RINCIAN]: ["ID_Usulan", "ID_Rincian", "Kode_Rekening", "Sumber_Dana", "Komponen", "Spesifikasi", "Keterangan", "Koefisien", "Volume", "Harga_Satuan", "Sub_Total", "Status_Item", "Catatan"],
    [SHEET_REKENING]: ["Kode_Rekening", "Nama_Rekening"],
    [SHEET_SUBKEG]: ["Kode_SubKegiatan", "Nama_SubKegiatan", "Indikator", "Satuan"],
    [SHEET_DANA]: ["Sumber_Dana"],
    [SHEET_SATUAN]: ["Satuan"]
  };

  for (const sheetName in schemas) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) { sheet = ss.insertSheet(sheetName); }
    const headers = schemas[sheetName];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#d9ead3");
  }
  
  const userSheet = ss.getSheetByName(SHEET_USER);
  if (userSheet && userSheet.getLastRow() <= 1) {
     userSheet.appendRow(["admin", "Admin Perencanaan", "admin123", "Admin Verifikator", "Admin Verifikator"]);
  }
  
  const defaultSheet = ss.getSheetByName("Sheet1");
  if (defaultSheet && ss.getSheets().length > 1) { ss.deleteSheet(defaultSheet); }
  
  return "Database berhasil di-generate!";
}

// ==========================================
// AUTH & MANAJEMEN DATA AWAL
// ==========================================
function verifyLogin(username, password) {
  try {
    const users = getDataFromSheet(SHEET_USER);
    if(users.length === 0) return { status: 'error', message: 'Database Master User kosong!' };

    const user = users.find(u => u[COLUMN_MAP.username] === username && u[COLUMN_MAP.password] === password);
    
    if (user) {
      return { 
        status: 'success', 
        bidang: user[COLUMN_MAP.bidang], 
        username: user[COLUMN_MAP.username],
        nama: user[COLUMN_MAP.nama], // Mengambil kolom Nama
        level: user[COLUMN_MAP.level_akses] || 'Level Program' // Default fallback jika kosong
      };
    }
    return { status: 'error', message: 'Username atau password salah!' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function getInitialData() {
  try {
    const ss = getDb();
    let pengaturanData = {};
    const sheetPengaturan = ss.getSheetByName(SHEET_PENGATURAN);
    if (sheetPengaturan) {
      const rawPengaturan = sheetPengaturan.getDataRange().getDisplayValues();
      for (let i = 0; i < rawPengaturan.length; i++) {
        let key = String(rawPengaturan[i][0] || "").trim();
        if (key && key !== "Parameter (Key)") { 
          pengaturanData[key] = rawPengaturan[i][1];
        }
      }
    }

    return { 
      status: 'success', 
      pengaturan: pengaturanData,
      subKegiatan: getDataFromSheet(SHEET_SUBKEG),
      rekening: getDataFromSheet(SHEET_REKENING),
      sumberDana: getDataFromSheet(SHEET_DANA).map(r => r[COLUMN_MAP.sumber_dana]).filter(v => v),
      satuan: getDataFromSheet(SHEET_SATUAN).map(r => r[COLUMN_MAP.satuan]).filter(v => v)
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

// ==========================================
// FUNGSI CRUD USULAN (HEADER)
// ==========================================
function getUsulanData(bidang, username, level) {
  try {
    let data = getDataFromSheet(SHEET_USULAN);
    
    // BACKEND FILTERING: Sangat penting untuk keamanan akses data antar Level
    if (level === 'Admin Verifikator' || bidang === 'Admin Verifikator') {
        // Admin melihat semua data, tidak di-filter sama sekali
    } else if (level === 'Level Bidang' || level === 'Kepala Bidang') {
        // Level Bidang melihat seluruh data yang ada dalam Unit Kerjanya
        data = data.filter(item => item[COLUMN_MAP.bidang] === bidang);
    } else {
        // Level Program HANYA melihat usulan miliknya sendiri
        data = data.filter(item => item[COLUMN_MAP.bidang] === bidang && item[COLUMN_MAP.pembuat] === username);
    }
    
    return { status: 'success', data: data };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function saveUsulan(payload) {
  try {
    const sheet = getDb().getSheetByName(SHEET_USULAN);
    if (!sheet) throw new Error("Sheet Data_Usulan tidak ditemukan!");

    const data = sheet.getDataRange().getDisplayValues();
    const headers = data[0].map(h => String(h).trim().replace(/\r/g, ''));
    
    let idUsulan = payload[COLUMN_MAP.id_usulan];
    let isNew = false;
    
    if (!idUsulan) {
      idUsulan = generateId("USL");
      isNew = true;
    }

    const rowData = headers.map(h => {
      if (h === COLUMN_MAP.id_usulan) return idUsulan;
      if (h === COLUMN_MAP.status && isNew) return "DRAFT"; 
      return payload[h] !== undefined ? payload[h] : "";
    });

    if (isNew) {
      sheet.appendRow(rowData);
    } else {
      let found = false;
      for (let i = 1; i < data.length; i++) {
        if (data[i][headers.indexOf(COLUMN_MAP.id_usulan)] === idUsulan) {
          sheet.getRange(i + 1, 1, 1, headers.length).setValues([rowData]);
          found = true; break;
        }
      }
      if (!found) sheet.appendRow(rowData); 
    }
    
    return { status: 'success', message: 'Usulan berhasil disimpan!', id: idUsulan };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function deleteUsulan(idUsulan) {
  try {
    const ss = getDb();
    const sheetUsulan = ss.getSheetByName(SHEET_USULAN);
    if(sheetUsulan) {
      const dataUsulan = sheetUsulan.getDataRange().getDisplayValues();
      const idIndexUsl = dataUsulan[0].map(h => String(h).trim()).indexOf(COLUMN_MAP.id_usulan);
      for (let i = dataUsulan.length - 1; i >= 1; i--) {
        if (dataUsulan[i][idIndexUsl] === idUsulan) {
          sheetUsulan.deleteRow(i + 1);
          break;
        }
      }
    }

    const sheetRincian = ss.getSheetByName(SHEET_RINCIAN);
    if(sheetRincian) {
      const dataRincian = sheetRincian.getDataRange().getDisplayValues();
      const idIndexRnc = dataRincian[0].map(h => String(h).trim()).indexOf(COLUMN_MAP.id_usulan);
      for (let i = dataRincian.length - 1; i >= 1; i--) {
        if (dataRincian[i][idIndexRnc] === idUsulan) {
          sheetRincian.deleteRow(i + 1);
        }
      }
    }
    return { status: 'success', message: 'Usulan dan rinciannya berhasil dihapus.' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function updateStatusUsulan(idUsulan, statusBaru) {
  try {
    const sheet = getDb().getSheetByName(SHEET_USULAN);
    if (!sheet) throw new Error("Sheet Data_Usulan tidak ditemukan!");

    const data = sheet.getDataRange().getDisplayValues();
    const headers = data[0].map(h => String(h).trim());
    const idIndex = headers.indexOf(COLUMN_MAP.id_usulan);
    const statusIndex = headers.indexOf(COLUMN_MAP.status);

    for (let i = 1; i < data.length; i++) {
      if (data[i][idIndex] === idUsulan) {
        sheet.getRange(i + 1, statusIndex + 1).setValue(statusBaru);
        return { status: 'success', message: 'Status berhasil diperbarui.' };
      }
    }
    return { status: 'error', message: 'Usulan tidak ditemukan.' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

// ==========================================
// FUNGSI CRUD RINCIAN
// ==========================================
function getRincianData(idUsulan = null) {
  try {
    let data = getDataFromSheet(SHEET_RINCIAN);
    if (idUsulan) {
       data = data.filter(item => item[COLUMN_MAP.id_usulan] === idUsulan);
    }
    return { status: 'success', data: data };
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}

function saveRincian(payload) {
  try {
    const sheet = getDb().getSheetByName(SHEET_RINCIAN);
    if (!sheet) throw new Error("Sheet Rincian_Usulan tidak ditemukan!");

    const data = sheet.getDataRange().getDisplayValues();
    const headers = data[0].map(h => String(h).trim().replace(/\r/g, ''));
    
    let idRincian = payload[COLUMN_MAP.id_rincian];
    let isNew = false;
    
    if (!idRincian) {
      idRincian = generateId("RNC");
      isNew = true;
    }

    const rowData = headers.map(h => {
      if (h === COLUMN_MAP.id_rincian) return idRincian;
      return payload[h] !== undefined ? payload[h] : "";
    });

    if (isNew) {
      sheet.appendRow(rowData);
    } else {
      let found = false;
      for (let i = 1; i < data.length; i++) {
        if (data[i][headers.indexOf(COLUMN_MAP.id_rincian)] === idRincian) {
          sheet.getRange(i + 1, 1, 1, headers.length).setValues([rowData]);
          found = true; break;
        }
      }
      if (!found) sheet.appendRow(rowData);
    }

    recalculateTotalAnggaran(payload[COLUMN_MAP.id_usulan]);
    return { status: 'success', message: 'Rincian berhasil disimpan!' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function deleteRincian(idRincian, idUsulan) {
  try {
    const sheet = getDb().getSheetByName(SHEET_RINCIAN);
    if (!sheet) throw new Error("Sheet Rincian_Usulan tidak ditemukan!");

    const data = sheet.getDataRange().getDisplayValues();
    const idIndex = data[0].map(h => String(h).trim()).indexOf(COLUMN_MAP.id_rincian);
    
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][idIndex] === idRincian) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    
    if (idUsulan) recalculateTotalAnggaran(idUsulan);
    return { status: 'success', message: 'Rincian berhasil dihapus.' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function recalculateTotalAnggaran(idUsulan) {
  const ss = getDb();
  const sheetRincian = ss.getSheetByName(SHEET_RINCIAN);
  if(!sheetRincian) return;

  const dataRincian = sheetRincian.getDataRange().getDisplayValues();
  const hRincian = dataRincian[0].map(h => String(h).trim());
  
  const idUslIdxRnc = hRincian.indexOf(COLUMN_MAP.id_usulan);
  const subTotalIdx = hRincian.indexOf(COLUMN_MAP.sub_total);
  
  let totalAnggaran = 0;
  for (let i = 1; i < dataRincian.length; i++) {
    if (dataRincian[i][idUslIdxRnc] === idUsulan) {
      // Membersihkan string menjadi angka sebelum kalkulasi
      let val = dataRincian[i][subTotalIdx];
      let num = 0;
      if (val) {
        let str = String(val);
        num = Number(str.replace(/[^0-9]/g, ''));
        if (str.includes('-')) num = -num;
      }
      totalAnggaran += num;
    }
  }
  
  const sheetUsulan = ss.getSheetByName(SHEET_USULAN);
  if(!sheetUsulan) return;

  const dataUsulan = sheetUsulan.getDataRange().getDisplayValues();
  const hUsulan = dataUsulan[0].map(h => String(h).trim());
  const idUslIdx = hUsulan.indexOf(COLUMN_MAP.id_usulan);
  const totalIdx = hUsulan.indexOf(COLUMN_MAP.total_anggaran);
  
  for (let i = 1; i < dataUsulan.length; i++) {
    if (dataUsulan[i][idUslIdx] === idUsulan) {
      sheetUsulan.getRange(i + 1, totalIdx + 1).setValue(totalAnggaran);
      break;
    }
  }
}

// ==========================================
// ENDPOINT KHUSUS ADMIN
// ==========================================
function getUsers() {
  return getDataFromSheet(SHEET_USER);
}

// Terpadat parameter nama dan level untuk management Multi-Level
function saveUser(oldUsername, newUsername, nama, password, bidang, level) {
  try {
    const sheet = getDb().getSheetByName(SHEET_USER);
    const data = sheet.getDataRange().getDisplayValues();
    const headers = data[0].map(h => String(h).trim().replace(/\r/g, ''));
    
    // Create Baru
    if (!oldUsername) {
      const isExist = data.some(row => row[headers.indexOf(COLUMN_MAP.username)] === newUsername);
      if (isExist) return { status: 'error', message: 'Username sudah digunakan!' };
      
      const newRow = headers.map(h => {
        if (h === COLUMN_MAP.username) return newUsername;
        if (h === COLUMN_MAP.nama) return nama;
        if (h === COLUMN_MAP.password) return password;
        if (h === COLUMN_MAP.bidang) return bidang;
        if (h === COLUMN_MAP.level_akses) return level;
        return "";
      });
      sheet.appendRow(newRow);
      return { status: 'success', message: 'User berhasil ditambahkan.' };
    }
    
    // Update Mode
    for (let i = 1; i < data.length; i++) {
      if (data[i][headers.indexOf(COLUMN_MAP.username)] === oldUsername) {
        if(headers.indexOf(COLUMN_MAP.username) > -1) sheet.getRange(i + 1, headers.indexOf(COLUMN_MAP.username) + 1).setValue(newUsername);
        if(headers.indexOf(COLUMN_MAP.nama) > -1) sheet.getRange(i + 1, headers.indexOf(COLUMN_MAP.nama) + 1).setValue(nama);
        if(headers.indexOf(COLUMN_MAP.password) > -1) sheet.getRange(i + 1, headers.indexOf(COLUMN_MAP.password) + 1).setValue(password);
        if(headers.indexOf(COLUMN_MAP.bidang) > -1) sheet.getRange(i + 1, headers.indexOf(COLUMN_MAP.bidang) + 1).setValue(bidang);
        if(headers.indexOf(COLUMN_MAP.level_akses) > -1) sheet.getRange(i + 1, headers.indexOf(COLUMN_MAP.level_akses) + 1).setValue(level);
        return { status: 'success', message: 'User berhasil diupdate.' };
      }
    }
    return { status: 'error', message: 'User tidak ditemukan.' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function deleteUser(username) {
  try {
    const sheet = getDb().getSheetByName(SHEET_USER);
    const data = sheet.getDataRange().getDisplayValues();
    const uIndex = data[0].map(h => String(h).trim().replace(/\r/g, '')).indexOf(COLUMN_MAP.username);
    
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][uIndex] === username) {
        sheet.deleteRow(i + 1);
        return { status: 'success', message: 'User berhasil dihapus.' };
      }
    }
    return { status: 'error', message: 'User tidak ditemukan.' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function savePengaturanSistem(tahapan, deadline, tahun) {
  try {
    const ss = getDb();
    let sheet = ss.getSheetByName(SHEET_PENGATURAN);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_PENGATURAN);
      sheet.appendRow(["Parameter (Key)", "Nilai (Value)"]);
    }
    
    const data = sheet.getDataRange().getDisplayValues();
    function updateParam(key, value) {
      let found = false;
      for (let i = 0; i < data.length; i++) {
        if (data[i][0] == key) {
          sheet.getRange(i + 1, 2).setValue(value);
          found = true; break;
        }
      }
      if (!found) sheet.appendRow([key, value]);
    }
    
    updateParam("Tahapan_Aktif", tahapan);
    updateParam("Batas_Waktu", deadline);
    updateParam("Tahun_Anggaran", tahun);
    
    return { status: 'success', message: 'Pengaturan berhasil disimpan!' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}
