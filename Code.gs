var SPREADSHEET_ID = "1LWCOYr3Lwi8fcyoq7yO7GCqV8fdIq76GUGzQ79mKUMk";

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Sistem BBM Solar')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSS() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getOrCreateSheet(ss, sheetName, headers) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (headers && headers.length > 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e9ecef");
    }
  }
  return sheet;
}

// === SISTEM FITUR PIN USER & ADMIN TERPISAH ===
function getUserPin() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var savedPin = scriptProperties.getProperty('USER_PIN');
  if (!savedPin) {
    savedPin = "1111"; // PIN Bawaan Awal User
    scriptProperties.setProperty('USER_PIN', savedPin);
  }
  return savedPin;
}

function getAdminPin() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var savedPin = scriptProperties.getProperty('ADMIN_PIN');
  if (!savedPin) {
    savedPin = "1234"; // PIN Bawaan Awal Admin
    scriptProperties.setProperty('ADMIN_PIN', savedPin);
  }
  return savedPin;
}

function verifyUserPin(inputPin) {
  var currentPin = getUserPin();
  if (String(inputPin) === String(currentPin)) {
    return { success: true, message: "PIN User Benar!" };
  } else {
    return { success: false, message: "PIN User Salah! Akses ditolak." };
  }
}

function verifyAdminPin(inputPin) {
  var currentPin = getAdminPin();
  if (String(inputPin) === String(currentPin)) {
    return { success: true, message: "PIN Admin Benar!" };
  } else {
    return { success: false, message: "PIN Admin Salah! Akses ditolak." };
  }
}

function changeUserPin(oldPin, newPin) {
  var currentPin = getUserPin();
  if (String(oldPin) !== String(currentPin)) {
    return { success: false, message: "PIN User Lama yang Anda masukkan salah!" };
  }
  if (!newPin || String(newPin).trim().length < 4) {
    return { success: false, message: "PIN Baru minimal harus 4 digit/karakter!" };
  }
  
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('USER_PIN', String(newPin).trim());
  return { success: true, message: "PIN User berhasil diperbarui!" };
}

function changeAdminPin(oldPin, newPin) {
  var currentPin = getAdminPin();
  if (String(oldPin) !== String(currentPin)) {
    return { success: false, message: "PIN Admin Lama yang Anda masukkan salah!" };
  }
  if (!newPin || String(newPin).trim().length < 4) {
    return { success: false, message: "PIN Baru minimal harus 4 digit/karakter!" };
  }
  
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('ADMIN_PIN', String(newPin).trim());
  return { success: true, message: "PIN Admin berhasil diperbarui!" };
}

// === GOOGLE DRIVE FOLDER MANAGEMENT (OTOMATIS TAHUN & BULAN) ===
function getOrCreateFolder(tglInput) {
  var d = tglInput ? new Date(tglInput) : new Date();
  var bln = ("0" + (d.getMonth() + 1)).slice(-2);
  var thn = d.getFullYear();
  var folderName = "Lampiran_BBM_" + thn + "-" + bln;
  
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder(folderName);
  }
}

// === AMBIL REKAPAN DATA ===
function ambilRekapanData() {
  try {
    var ss = getSS();
    var sheetP = getOrCreateSheet(ss, 'Penggunaan', ["Tanggal Input", "Waktu ON", "Waktu OFF", "Durasi (Menit)", "Durasi (Jam)", "Mode", "Solar Terpakai (L)"]);
    var sheetS = getOrCreateSheet(ss, 'Stok', ["Tanggal", "Keterangan", "Jumlah Masuk (L)", "Lampiran File"]);
    
    var totalMasuk = 0;
    var totalKeluar = 0;
    
    // Read Data Stok
    var listStok = [];
    var lastRowS = sheetS.getLastRow();
    if (lastRowS > 1) {
      var dataS = sheetS.getRange(2, 1, lastRowS - 1, 4).getValues();
      for (var j = dataS.length - 1; j >= 0; j--) {
        if (!dataS[j][0] && !dataS[j][2]) continue;
        
        var val = dataS[j][2];
        if (typeof val === 'string') val = val.replace(',', '.');
        var num = parseFloat(val) || 0;
        totalMasuk += num;
        listStok.push({
          rowIndex: j + 2,
          tanggal: formatDate(dataS[j][0]),
          keterangan: String(dataS[j][1] || "-"),
          jumlah: num,
          lampiran: String(dataS[j][3] || "-")
        });
      }
    }
    
    // Read Data Penggunaan
    var listPenggunaan = [];
    var lastRowP = sheetP.getLastRow();
    if (lastRowP > 1) {
      var dataP = sheetP.getRange(2, 1, lastRowP - 1, 7).getValues();
      for (var i = dataP.length - 1; i >= 0; i--) {
        if (!dataP[i][0] && !dataP[i][6]) continue;
        
        var valK = dataP[i][6];
        if (typeof valK === 'string') valK = valK.replace(',', '.');
        var numK = parseFloat(valK) || 0;
        totalKeluar += numK;
        listPenggunaan.push({
          tanggalInput: formatDate(dataP[i][0]),
          waktuOn: formatDate(dataP[i][1]),
          waktuOff: formatDate(dataP[i][2]),
          menit: dataP[i][3] || 0,
          jam: dataP[i][4] || 0,
          mode: String(dataP[i][5] || "-"),
          solar: numK
        });
      }
    }
    
    var sisa = (totalMasuk - totalKeluar).toFixed(2);
    
    return {
      success: true,
      sisaStok: sisa,
      penggunaan: listPenggunaan,
      stok: listStok
    };
  } catch(e) {
    return { 
      success: false, 
      sisaStok: "0.00", 
      penggunaan: [], 
      stok: [], 
      error: "Detail Error: " + e.message
    };
  }
}

// === INPUT PENGGUNAAN (USER) ===
function simpanPenggunaan(waktuOn, waktuOff, mode) {
  try {
    var ss = getSS();
    var sheet = getOrCreateSheet(ss, 'Penggunaan', ["Tanggal Input", "Waktu ON", "Waktu OFF", "Durasi (Menit)", "Durasi (Jam)", "Mode", "Solar Terpakai (L)"]);
    
    var dOn = new Date(waktuOn);
    var dOff = new Date(waktuOff);
    
    if (isNaN(dOn.getTime()) || isNaN(dOff.getTime())) {
      return { success: false, message: "Format tanggal & waktu tidak valid." };
    }
    if (dOff <= dOn) {
      return { success: false, message: "Waktu OFF harus lebih besar / setelah Waktu ON!" };
    }
    
    var diffMs = dOff - dOn;
    var durasiMenit = Math.round(diffMs / (1000 * 60));
    var durasiJam = parseFloat((durasiMenit / 60).toFixed(4));
    
    var faktorLoad = (mode === "Pemanasan") ? 0.30 : 1.00;
    var totalSolar = parseFloat((0.21 * 125 * faktorLoad * durasiJam).toFixed(2));
    
    var dataRekap = ambilRekapanData();
    var sisaStokSaatIni = parseFloat(dataRekap.sisaStok);
    
    if (totalSolar > sisaStokSaatIni) {
      return { success: false, message: "Stok BBM tidak cukup! Sisa stok: " + sisaStokSaatIni + " Liter" };
    }
    
    var tglInput = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    var strOn = Utilities.formatDate(dOn, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    var strOff = Utilities.formatDate(dOff, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    sheet.appendRow([tglInput, strOn, strOff, durasiMenit, parseFloat(durasiJam.toFixed(2)), mode, totalSolar]);
    
    return { success: true, durasiMenit: durasiMenit, totalSolar: totalSolar };
  } catch(err) {
    return { success: false, message: "Error: " + err.toString() };
  }
}

// === TAMBAH STOK (ADMIN) ===
function tambahStok(tglInput, jumlah, ket, filesData, adminPin) {
  try {
    var verify = verifyAdminPin(adminPin);
    if (!verify.success) {
      return verify;
    }
    if (!jumlah || parseFloat(jumlah) <= 0) {
      return { success: false, message: "Jumlah stok masuk harus lebih besar dari 0!" };
    }
    var folder = getOrCreateFolder(tglInput);
    var urlLampiran = [];
    if (filesData && filesData.length > 0) {
      for (var i = 0; i < filesData.length; i++) {
        var fileObj = filesData[i];
        var bytes = Utilities.base64Decode(fileObj.bytes);
        var blob = Utilities.newBlob(bytes, fileObj.mimeType, fileObj.fileName);
        
        var createdFile = folder.createFile(blob);
        createdFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        
        urlLampiran.push(createdFile.getUrl());
      }
    }
    var stringLampiran = urlLampiran.length > 0 ? urlLampiran.join(" | ") : "-";
    var ss = getSS();
    var sheet = getOrCreateSheet(ss, 'Stok', ["Tanggal", "Keterangan", "Jumlah Masuk (L)", "Lampiran File"]);
    
    var tglMasuk = tglInput ? new Date(tglInput) : new Date();
    var strTgl = Utilities.formatDate(tglMasuk, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    sheet.appendRow([strTgl, ket || "-", parseFloat(jumlah), stringLampiran]);
    return { success: true, message: "Stok BBM berhasil ditambahkan!" };
  } catch(e) {
    return { success: false, message: "Gagal menyimpan stok: " + e.toString() };
  }
}

// === HAPUS RIWAYAT STOK (ADMIN) ===
function hapusStok(rowIndex, adminPin) {
  try {
    var verify = verifyAdminPin(adminPin);
    if (!verify.success) {
      return verify;
    }
    var ss = getSS();
    var sheet = ss.getSheetByName('Stok');
    if (!sheet) {
      return { success: false, message: "Sheet Stok tidak ditemukan." };
    }
    sheet.deleteRow(rowIndex);
    return { success: true, message: "Riwayat stok berhasil dihapus!" };
  } catch(e) {
    return { success: false, message: "Gagal menghapus data: " + e.toString() };
  }
}

// === FORMATTER TANGGAL & JAM ===
function formatDate(date) {
  if (!date) return "-";
  if (date instanceof Date) {
    if (date.getFullYear() === 1899) {
      return Utilities.formatDate(date, Session.getScriptTimeZone(), "HH:mm");
    }
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
  }
  return String(date);
}

function mintaIzinOtorisasi() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var folder = getOrCreateFolder();
  Logger.log("Izin Berhasil Diaktifkan!");
}