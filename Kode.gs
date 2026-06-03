// WAJIB: ISI DENGAN LINK SPREADSHEET ADMIN ANDA DI BAWAH INI
// ========================================================================
const ADMIN_DB_URL = "https://docs.google.com/spreadsheets/d/1jBwX1TpBqhGbpqVZ6em0u6WAkl0iP3p6Mlc1QeADKm8/edit?usp=drivesdk";

function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('MBG (Manage Budget Gue)')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=0.86, user-scalable=no');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ================= SISTEM AKUN & AUTENTIKASI =================
function getAdminSS() {
  if(!ADMIN_DB_URL || ADMIN_DB_URL === "ISI_DENGAN_LINK_SPREADSHEET_ANDA_DISINI") {
    throw new Error("Admin belum mengatur link Master Database di Code.gs");
  }
  return SpreadsheetApp.openByUrl(ADMIN_DB_URL);
}

function registerUser(name, email, phone, ref, password) {
  if(!email || !password || !name) throw new Error("Nama, Email, dan Password wajib diisi!");
  
  const ss = getAdminSS();
  let sheet = ss.getSheetByName("Users");
  if(!sheet) {
    sheet = ss.insertSheet("Users");
    sheet.appendRow(["Tanggal Daftar", "Nama", "Email", "No HP", "Kode Reff", "Password", "SS_ID", "Status"]);
    sheet.getRange("A1:H1").setFontWeight("bold");
  }
  
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
    if(data[i][2].toString().toLowerCase() === email.toLowerCase()) throw new Error("Email sudah terdaftar!");
  }
  
  const newSS = SpreadsheetApp.create("MBG_DB_" + name);
  const newId = newSS.getId();
  
  sheet.appendRow([new Date(), name, email.toLowerCase(), phone, ref, password, newId, "Pending"]);
  initSheets(newId); 
  
  return { success: true, message: "Pendaftaran berhasil! Menunggu verifikasi admin." };
}

function loginUser(email, password) {
  if (email === 'admin' && password === 'admin123') {
    return { success: true, role: 'admin', name: 'Administrator', email: 'admin' };
  }

  const ss = getAdminSS();
  const sheet = ss.getSheetByName("Users");
  if(!sheet) throw new Error("Sistem belum memiliki pengguna.");

  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
    if(data[i][2].toString().toLowerCase() === email.toLowerCase() && data[i][5].toString() === password) {
      if(data[i][7] !== "Verified") throw new Error("Akun Anda sedang menunggu verifikasi admin.");
      return { success: true, ssId: data[i][6], name: data[i][1], email: email.toLowerCase(), role: 'user' };
    }
  }
  throw new Error("Email atau Password salah!");
}

function updateUserProfile(oldEmail, newName, newEmail, newPass) {
  const sheet = getAdminSS().getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
    if(data[i][2].toString().toLowerCase() === oldEmail.toLowerCase()) {
      sheet.getRange(i+1, 2).setValue(newName);
      sheet.getRange(i+1, 3).setValue(newEmail.toLowerCase());
      if(newPass) sheet.getRange(i+1, 6).setValue(newPass);
      return { name: newName, email: newEmail.toLowerCase() };
    }
  }
  throw new Error("Akun tidak ditemukan.");
}

function getBackupUrl(ssId) {
  if(!ssId) throw new Error("Sesi tidak valid.");
  return "https://docs.google.com/spreadsheets/d/" + ssId + "/export?format=xlsx";
}

// ================= FITUR PANEL ADMIN =================
function getAdminUsers() {
  const sheet = getAdminSS().getSheetByName("Users");
  if(!sheet) return [];
  const data = sheet.getDataRange().getValues();
  data.shift(); 
  let users = [];
  data.forEach((r, idx) => {
    if(r[1]) {
      let tglStr = "";
      try { tglStr = Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"); } 
      catch(e) { tglStr = r[0] ? r[0].toString() : "-"; }
      
      users.push({ row: idx+2, tgl: tglStr, nama: r[1].toString(), email: r[2].toString(), hp: r[3].toString(), ref: r[4] ? r[4].toString() : "-", pass: r[5].toString(), status: r[7] ? r[7].toString() : "Pending" });
    }
  });
  return users.reverse(); 
}

function adminVerifyUser(rowIdx) { getAdminSS().getSheetByName("Users").getRange(rowIdx, 8).setValue("Verified"); return "Akun berhasil diverifikasi!"; }
function adminDeleteUser(rowIdx) { getAdminSS().getSheetByName("Users").deleteRow(rowIdx); return "Akun berhasil dihapus!"; }
function adminEditUser(rowIdx, nama, email, hp, ref, pass) {
  const sheet = getAdminSS().getSheetByName("Users");
  sheet.getRange(rowIdx, 2, 1, 5).setValues([[nama, email, hp, ref, pass]]);
  return "Data user berhasil diperbarui!";
}

// ================= INISIALISASI SHEET KEUANGAN USER =================
function initSheets(ssId) {
  const ss = SpreadsheetApp.openById(ssId);
  const requiredSheets = [
    { name: "Pemasukan", headers: ["Tanggal", "Sumber", "Nominal", "Dompet"] },
    { name: "Pengeluaran", headers: ["Tanggal", "Barang", "Kategori", "Nominal", "Keterangan", "Dompet"] },
    { name: "Dompet", headers: ["Nama Dompet", "Saldo Awal", "Kategori", "No Rekening"] },
    { name: "Assets", headers: ["Nama Aset", "Nilai/Saldo"] },
    { name: "Tagihan", headers: ["Tanggal", "Nama Tagihan", "Nominal", "Keterangan", "Terbayar", "Status"] },
    { name: "Hutang", headers: ["Tanggal", "Hutang Apa", "Tempat/Siapa", "Nominal", "Keterangan", "Terbayar", "Status"] },
    { name: "Transfer", headers: ["Tanggal", "Dari", "Ke", "Nominal"] }
  ];
  
  requiredSheets.forEach(s => {
    let sheet = ss.getSheetByName(s.name);
    if (!sheet) {
      sheet = ss.insertSheet(s.name);
      sheet.appendRow(s.headers);
      sheet.getRange(1, 1, 1, s.headers.length).setFontWeight("bold");
    }
  });
}

// ================= ENGINE DATA KEUANGAN USER =================
function getDashboardData(ssId) {
  if(!ssId) throw new Error("Sesi tidak valid.");
  try {
    const ss = SpreadsheetApp.openById(ssId);
    initSheets(ssId); 
    
    const res = { totalSaldo: 0, totalHutang: 0, totalAssets: 0, rekap: { h: {i:0, o:0}, w: {i:0, o:0}, m: {i:0, o:0} }, topExpenses: [], history: [], wallets: [], assets: [], tagihan: [], hutang: [] };
    const now = new Date(); now.setHours(0,0,0,0); const t0 = now.getTime(), t7 = t0 - (7*86400000), t30 = t0 - (30*86400000);

    const getSafeData = (name) => { const sh = ss.getSheetByName(name); return (!sh || sh.getLastRow()<2) ? [] : sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).getValues(); };

    const walletData = getSafeData("Dompet"), assetData = getSafeData("Assets"), mskData = getSafeData("Pemasukan"), kelData = getSafeData("Pengeluaran"), tagihanData = getSafeData("Tagihan"), hutangData = getSafeData("Hutang");

    let walletMap = {};
    walletData.forEach(r => { if(r[0]) walletMap[r[0].toString()] = { nama: r[0].toString(), saldo: Number(r[1])||0, kategori: r[2]||'Cash', rek: r[3]||'' }; });

    assetData.forEach(r => { if(r[0]) { const val = Number(r[1])||0; res.totalAssets += val; res.assets.push({ nama: r[0].toString(), saldo: val }); } });

    mskData.forEach(r => { 
      if(r[0] && r[1]) {
        const nom = Number(r[2])||0; const dompet = r[3]?r[3].toString():"";
        if(walletMap[dompet]) walletMap[dompet].saldo += nom;
        let d = new Date(r[0]); d.setHours(0,0,0,0); let tm = d.getTime();
        if(tm >= t0) res.rekap.h.i += nom; if(tm >= t7) res.rekap.w.i += nom; if(tm >= t30) res.rekap.m.i += nom;
        res.history.push({ tgl: Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd"), nama: r[1].toString(), nominal: nom, tipe: 'income', detail: dompet });
      }
    });

    let tempKel = [];
    kelData.forEach(r => {
      if(r[0] && r[1]) {
        const nom = Number(r[3])||0; const dompetDigunakan = r[5]?r[5].toString():""; 
        tempKel.push({nama: r[1].toString(), nominal: nom});
        if(walletMap[dompetDigunakan]) walletMap[dompetDigunakan].saldo -= nom;
        let d = new Date(r[0]); d.setHours(0,0,0,0); let tm = d.getTime();
        if(tm >= t0) res.rekap.h.o += nom; if(tm >= t7) res.rekap.w.o += nom; if(tm >= t30) res.rekap.m.o += nom;
        res.history.push({ tgl: Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd"), nama: r[1].toString(), nominal: nom, tipe: 'expense', detail: dompetDigunakan||r[2].toString() });
      }
    });

    tagihanData.forEach((r, idx) => { if(r[0]&&r[1]) res.tagihan.push({ row: idx+2, tgl: Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), "yyyy-MM-dd"), nama: r[1].toString(), nom: Number(r[2])||0, ket: r[3], terbayar: Number(r[4])||0, status: r[5]||"Belum Lunas" }); });
    hutangData.forEach((r, idx) => { if(r[0]&&r[1]) { let sisa = (Number(r[3])||0) - (Number(r[5])||0); res.totalHutang += sisa>0?sisa:0; res.hutang.push({ row: idx+2, tgl: Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), "yyyy-MM-dd"), nama: r[1].toString(), tempat: r[2], nom: Number(r[3])||0, ket: r[4], terbayar: Number(r[5])||0, status: r[6]||"Belum Lunas" }); }});

    res.wallets = Object.values(walletMap);
    res.wallets.forEach(w => { res.totalSaldo += w.saldo; });
    res.topExpenses = tempKel.sort((a, b) => b.nominal - a.nominal).slice(0, 3);
    res.history.sort((a, b) => new Date(b.tgl) - new Date(a.tgl));

    return res;
  } catch(e) { throw new Error(e.message); }
}

// ================= MANAGE (CRUD & PAYMENTS USER) =================
function manageMasterItem(ssId, sheetName, action, originalName, name, balance, kategori, rek) {
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName(sheetName); const data = sheet.getDataRange().getValues();
  let rowData = [name, balance]; if(sheetName === 'Dompet') rowData = [name, balance, kategori||'Cash', rek||''];

  if(action === 'add') { sheet.appendRow(rowData); return "Berhasil ditambahkan!"; }
  for(let i=1; i<data.length; i++) {
    if(data[i][0] === originalName) {
      if(action === 'edit') { sheet.getRange(i+1, 1, 1, rowData.length).setValues([rowData]); return "Berhasil diubah!"; } 
      else if(action === 'delete') { sheet.deleteRow(i+1); return "Berhasil dihapus!"; }
    }
  }
  throw new Error("Data tidak ditemukan.");
}

function processTransfer(ssId, fromDompet, toDompet, nominal) {
  const ss = SpreadsheetApp.openById(ssId); const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  ss.getSheetByName("Transfer").appendRow([dateStr, fromDompet, toDompet, nominal]);
  ss.getSheetByName("Pengeluaran").appendRow([dateStr, "Transfer ke "+toDompet, "Transfer Internal", nominal, "Pemindahan Dana", fromDompet]);
  ss.getSheetByName("Pemasukan").appendRow([dateStr, "Transfer dari "+fromDompet, nominal, toDompet]);
  return "Transfer Rp " + Number(nominal).toLocaleString('id-ID') + " Berhasil!";
}

function processPayment(ssId, type, rowIndex, amount, dompet) {
  const ss = SpreadsheetApp.openById(ssId); const sheet = ss.getSheetByName(type);
  const data = sheet.getRange(rowIndex, 1, 1, 7).getValues()[0];
  const isTagihan = (type === 'Tagihan'); const nominal = Number(data[2])||0; let terbayar = Number(data[isTagihan?4:5])||0;
  
  terbayar += Number(amount); let status = (terbayar >= nominal) ? "Lunas" : "Cicil / Belum Lunas";
  sheet.getRange(rowIndex, isTagihan?5:6).setValue(terbayar); sheet.getRange(rowIndex, isTagihan?6:7).setValue(status);

  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  ss.getSheetByName("Pengeluaran").appendRow([dateStr, "Bayar " + type + ": " + data[1], type, amount, "Pembayaran via Manage", dompet]);
  return "Pembayaran Berhasil!";
}

function saveData(ssId, type, data) {
  try { SpreadsheetApp.openById(ssId).getSheetByName(type).appendRow(data); return "Data Berhasil Disimpan!"; } catch(e) { throw new Error(e.message); }
}
