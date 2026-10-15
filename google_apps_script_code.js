// KODE GOOGLE APPS SCRIPT (Copy-Paste kode ini ke Apps Script Editor Google Sheet Anda)

// 1. FUNGSI UNTUK MENGAMBIL DATA DARI SPREADSHEET KE APLIKASI WEB (doGet)
function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Read Clients
    var sheetClient = ss.getSheetByName("Clients");
    var clients = [];
    if (sheetClient && sheetClient.getLastRow() > 1) {
      var clientValues = sheetClient.getRange(2, 1, sheetClient.getLastRow() - 1, 8).getValues();
      clients = clientValues.map(function(r) {
        return {
          id: String(r[0]),
          nama: String(r[1]),
          lokasi: String(r[2]),
          hp: String(r[3]),
          akun: Number(r[4]),
          biayaLain: Number(r[5]),
          ketBiayaLain: String(r[6]),
          total: Number(r[7])
        };
      });
    }

    // Read Payments
    var sheetPayment = ss.getSheetByName("Payments");
    var payments = [];
    if (sheetPayment && sheetPayment.getLastRow() > 1) {
      var payValues = sheetPayment.getRange(2, 1, sheetPayment.getLastRow() - 1, 7).getValues();
      payments = payValues.map(function(r) {
        return {
          id: String(r[0]),
          clientId: String(r[1]),
          startMonth: String(r[2]),
          endMonth: String(r[3]),
          amount: Number(r[4]),
          notes: String(r[5]),
          date: String(r[6])
        };
      });
    }

    var result = JSON.stringify({ status: "success", clients: clients, payments: payments });
    return ContentService.createTextOutput(result).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 2. FUNGSI UNTUK MENYIMPAN DATA DARI APLIKASI WEB KE SPREADSHEET (doPost)
function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var data = JSON.parse(e.postData.contents);

    // --- SIMPAN / OVERWRITE TABEL CLIENTS ---
    var sheetClient = ss.getSheetByName("Clients");
    if (!sheetClient) {
      sheetClient = ss.insertSheet("Clients");
    }
    sheetClient.clearContents();
    // Buat Header
    sheetClient.appendRow(["ID", "Nama Klien", "Lokasi", "No HP/WA", "Jumlah Akun", "Biaya Lain", "Ket Biaya Lain", "Total Tagihan/Bln"]);

    if (data.clients && data.clients.length > 0) {
      data.clients.forEach(function(c) {
        sheetClient.appendRow([
          c.id, 
          c.nama, 
          c.lokasi, 
          "'" + c.hp, // Tanda petik agar nomor HP 08... tidak berubah jadi angka tunggal
          c.akun, 
          c.biayaLain || 0, 
          c.ketBiayaLain || '', 
          c.total
        ]);
      });
    }

    // --- SIMPAN / OVERWRITE TABEL PAYMENTS ---
    var sheetPayment = ss.getSheetByName("Payments");
    if (!sheetPayment) {
      sheetPayment = ss.insertSheet("Payments");
    }
    sheetPayment.clearContents();
    // Buat Header
    sheetPayment.appendRow(["ID Transaksi", "ID Klien", "Bulan Mulai", "Bulan Sampai", "Jumlah Dibayar", "Catatan/Metode", "Tanggal Transaksi"]);

    if (data.payments && data.payments.length > 0) {
      data.payments.forEach(function(p) {
        sheetPayment.appendRow([
          p.id, 
          p.clientId, 
          p.startMonth, 
          p.endMonth, 
          p.amount, 
          p.notes || '', 
          p.date
        ]);
      });
    }

    return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Data berhasil disimpan ke Google Sheet!" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}