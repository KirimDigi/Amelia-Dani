// Code.gs - Copy paste ke Apps Script (Ekstensi > Apps Script)
// Spreadsheet ID FORM CATIN
var SPREADSHEET_ID = '1pQJrFVEhobOQoZZz0onkFzYGHvkmuiRMKEM_e308O-g';
var SHEET_NAME = 'RSVP';

function doPost(e) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(['Timestamp', 'Nama', 'Kehadiran', 'Jumlah Tamu', 'Ucapan', 'Waktu ISO']);
      sheet.getRange(1,1,1,6).setFontWeight('bold').setBackground('#E2078A').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
      sheet.autoResizeColumns(1,6);
    }
    // e.postData.contents berisi JSON string dari fetch
    var data = {};
    if (e.postData && e.postData.contents) {
      try { data = JSON.parse(e.postData.contents); } catch(err) { data = e.parameter; }
    } else {
      data = e.parameter;
    }
    var nama = (data.nama || data.Nama || '').toString().trim();
    var hadir = (data.hadir || data.Kehadiran || '').toString().trim();
    var tamu = (data.tamu || data['Jumlah Tamu'] || '0').toString().trim();
    if (hadir === 'Hadir' && (!tamu || tamu === '0')) tamu = '1';
    if (hadir === 'Tidak Hadir') tamu = '0';
    var ucapan = (data.ucapan || data.Ucapan || '').toString().trim();
    var timeISO = (data.time || new Date().toISOString()).toString();
    
    // Format tanggal jam WIB
    var now = new Date();
    // Simpan timestamp asli
    sheet.appendRow([now, nama, hadir, tamu, ucapan, timeISO]);
    
    return ContentService.createTextOutput(JSON.stringify({result:'success', message:'Tersimpan'}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({result:'error', error: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    // Jika ?action=list -> kembalikan JSON daftar ucapan
    if (e && e.parameter && e.parameter.action === 'list') {
      var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      var sheet = ss.getSheetByName(SHEET_NAME);
      if (!sheet) {
        return ContentService.createTextOutput(JSON.stringify([]))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var values = sheet.getDataRange().getValues();
      // values[0] = header, mulai dari baris 2
      var out = [];
      for (var i = 1; i < values.length; i++) {
        var row = values[i];
        // row: 0 Timestamp, 1 Nama, 2 Hadir, 3 Tamu, 4 Ucapan, 5 ISO
        if (!row[1]) continue; // skip jika nama kosong
        var ts = row[0];
        var tanggalJam = '';
        try {
          // Format ke WIB dd/MM/yyyy HH:mm
          tanggalJam = Utilities.formatDate(new Date(ts), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm');
        } catch(e2) { tanggalJam = ts.toString(); }
        out.push({
          tanggal: tanggalJam,
          nama: row[1],
          hadir: row[2],
          tamu: row[3],
          ucapan: row[4],
          iso: row[5] || ''
        });
      }
      // urut terbaru dulu
      out.reverse();
      return ContentService.createTextOutput(JSON.stringify(out))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput('RSVP Active - Gunakan POST untuk kirim, GET?action=list untuk ambil data')
      .setMimeType(ContentService.MimeType.TEXT);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({result:'error', error: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Untuk test manual di editor
function testAdd() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) sheet = SpreadsheetApp.openById(SPREADSHEET_ID).insertSheet(SHEET_NAME);
  sheet.appendRow([new Date(), 'Test Nama', 'Hadir', '2', 'Selamat ya!', new Date().toISOString()]);
}
