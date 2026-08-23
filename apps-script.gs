function doGet(e) {
  var action = e.parameter && e.parameter.action;

  if (action === 'lookup') {
    try {
      return ContentService.createTextOutput(JSON.stringify(buildTujuanLookup()))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(
        JSON.stringify({ error: err.message })
      ).setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (action === 'summary') {
    try {
      var dari = e.parameter && e.parameter.dari;
      var sampai = e.parameter && e.parameter.sampai;
      return ContentService.createTextOutput(JSON.stringify(buildSummary(dari, sampai)))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(
        JSON.stringify({ error: err.message })
      ).setMimeType(ContentService.MimeType.JSON);
    }
  }

  return ContentService.createTextOutput('API aktif. Gunakan POST untuk kirim data.');
}

function readTailRows(sheet, startRow, numCols, tailCount) {
  var lastRow = sheet.getLastRow();
  if (lastRow < startRow) return [];
  var from = Math.max(startRow, lastRow - tailCount + 1);
  return sheet.getRange(from, 1, lastRow - from + 1, numCols).getValues();
}

// Kolom TANGGAL isinya campur, tergantung baris itu dibuat oleh siapa:
//  - baris dari aplikasi  -> teks "23 Agu 2026"
//  - baris yang diketik manual di Sheet -> nilai Date asli, atau teks
//    "23/08/2026" / "2026-08-23"
// Versi lama hanya mengenali bentuk pertama, sehingga hampir semua baris
// lama dianggap tak bertanggal dan totalnya selalu Rp0.
function parseIndoDate(val) {
  if (val === null || val === undefined || val === '') return null;

  if (Object.prototype.toString.call(val) === '[object Date]') {
    if (isNaN(val.getTime())) return null;
    return new Date(val.getFullYear(), val.getMonth(), val.getDate());
  }

  var str = String(val).trim();
  if (!str) return null;

  // "23 Agu 2026" / "23 Agustus 2026"
  var parts = str.split(/\s+/);
  if (parts.length >= 3) {
    var day = parseInt(parts[0], 10);
    var monthNames = ['jan', 'feb', 'mar', 'apr', 'mei', 'jun', 'jul', 'agu', 'sep', 'okt', 'nov', 'des'];
    var mIdx = monthNames.indexOf(parts[1].toLowerCase().slice(0, 3));
    var year = parseInt(parts[2], 10);
    if (!isNaN(day) && mIdx !== -1 && !isNaN(year)) return new Date(year, mIdx, day);
  }

  // "23/08/2026", "23-08-2026", "23.08.2026"
  var m = str.match(/^(\d{1,2})[.\-\/](\d{1,2})[.\-\/](\d{2,4})$/);
  if (m) {
    var th = Number(m[3]);
    if (th < 100) th += 2000;
    return new Date(th, Number(m[2]) - 1, Number(m[1]));
  }

  // "2026-08-23"
  m = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));

  return null;
}

// Tanggal batas dikirim frontend sebagai "YYYY-MM-DD". Dibuat di zona waktu
// lokal (bukan new Date(str) yang diperlakukan UTC), supaya batas hari tidak
// bergeser dan baris di tanggal ujung tidak ikut hilang.
function parseIsoDate(str) {
  if (!str) return null;
  var m = String(str).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

function buildTotals(pb, uj, dari, sampai) {
  // Rentang mencakup kedua ujungnya: dari jam 00:00 tanggal awal sampai
  // 23:59 tanggal akhir.
  var awal = parseIsoDate(dari);
  var akhir = parseIsoDate(sampai);
  if (akhir) akhir = new Date(akhir.getFullYear(), akhir.getMonth(), akhir.getDate(), 23, 59, 59);

  function didalam(d) {
    if (!d) return false;
    if (awal && d < awal) return false;
    if (akhir && d > akhir) return false;
    return true;
  }

  var totals = { perbaikan: 0, uangJalan: 0, invoice: 0 };

  var pbLast = pb.getLastRow();
  if (pbLast >= 3) {
    // B TANGGAL, C TINDAKAN, D HARGA
    var pbData = pb.getRange(3, 2, pbLast - 2, 3).getValues();
    pbData.forEach(function(r) {
      if (!didalam(parseIndoDate(r[0]))) return;
      totals.perbaikan += Number(r[2]) || 0;
    });
  }

  var ujLast = uj.getLastRow();
  if (ujLast >= 3) {
    // B TANGGAL, C NO DO/SPE, D TUJUAN, E UANG JALAN, F, G, H INVOICE
    var ujData = uj.getRange(3, 2, ujLast - 2, 7).getValues();
    ujData.forEach(function(r) {
      if (!didalam(parseIndoDate(r[0]))) return;
      totals.uangJalan += Number(r[3]) || 0;
      totals.invoice += Number(r[6]) || 0;
    });
  }

  return totals;
}

function buildSummary(dari, sampai) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pb = ss.getSheetByName('PERBAIKAN');
  var uj = ss.getSheetByName('UANG JALAN');
  var tc = ss.getSheetByName('TAGIHAN DAN CICILAN');

  // Daftar "terakhir" ditampilkan di kotak yang bisa di-scroll, jadi boleh
  // lebih banyak dari sebelumnya tanpa membuat halaman memanjang.
  var pbRows = readTailRows(pb, 3, 7, 25);
  var ujRows = readTailRows(uj, 3, 9, 25);
  var tcRows = readTailRows(tc, 3, 5, 25);

  return {
    periode: { dari: dari || '', sampai: sampai || '' },
    recent: {
      perbaikan: pbRows.map(function(r) {
        return { tanggal: r[1], tindakan: r[2], harga: Number(r[3]) || 0, bayar: r[4], keterangan: r[6] };
      }).reverse(),
      uangJalan: ujRows.map(function(r) {
        return { tanggal: r[1], noDo: r[2], tujuan: r[3], uangJalan: Number(r[4]) || 0, invoice: Number(r[7]) || 0, customer: r[8] };
      }).reverse(),
      tagihan: tcRows.map(function(r) {
        return { bulan: r[1], tagihan: Number(r[2]) || 0, cicilan: Number(r[3]) || 0, operasional: Number(r[4]) || 0 };
      }).reverse()
    },
    totals: buildTotals(pb, uj, dari, sampai)
  };
}

function buildTujuanLookup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('UANG JALAN');
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return { tujuan: {}, customers: [] };

  // D=TUJUAN, E=UANG JALAN, F,G=(diabaikan), H=INVOICE, I=CUSTOMER
  var data = sheet.getRange(3, 4, lastRow - 2, 6).getValues();
  var groups = {};
  var customerCounts = {};

  data.forEach(function(row) {
    var tujuan = String(row[0] || '').trim();
    var uangJalan = Number(row[1]) || 0;
    var invoice = Number(row[4]) || 0;

    // Daftar nama customer yang pernah dipakai, untuk mengoreksi typo OCR
    // di sisi frontend. Dikumpulkan terpisah dari rute karena satu customer
    // bisa muncul di banyak rute.
    var customer = String(row[5] || '').trim();
    if (customer) {
      customerCounts[customer] = (customerCounts[customer] || 0) + 1;
    }

    if (!tujuan) return;

    var key = tujuan.toLowerCase().replace(/\s+/g, ' ');
    if (!groups[key]) {
      groups[key] = { tujuan: tujuan, uangJalanCounts: {}, invoiceCounts: {} };
    }
    if (uangJalan > 0) {
      groups[key].uangJalanCounts[uangJalan] = (groups[key].uangJalanCounts[uangJalan] || 0) + 1;
    }
    if (invoice > 0) {
      groups[key].invoiceCounts[invoice] = (groups[key].invoiceCounts[invoice] || 0) + 1;
    }
  });

  // Ambil nilai yang paling sering muncul (modus) per rute, supaya baris
  // kosong atau nilai anomali sesekali tidak ikut jadi acuan.
  function mode(counts) {
    var best = 0, bestCount = 0;
    Object.keys(counts).forEach(function(k) {
      if (counts[k] > bestCount) { bestCount = counts[k]; best = Number(k); }
    });
    return best;
  }

  var map = {};
  Object.keys(groups).forEach(function(key) {
    var g = groups[key];
    map[key] = {
      tujuan: g.tujuan,
      uangJalan: mode(g.uangJalanCounts),
      invoice: mode(g.invoiceCounts)
    };
  });

  // Urutkan customer dari yang paling sering muncul, supaya kalau ada dua nama
  // mirip, yang lebih sering dipakai jadi acuan utama.
  var customers = Object.keys(customerCounts).sort(function(a, b) {
    return customerCounts[b] - customerCounts[a];
  });

  return { tujuan: map, customers: customers };
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var result;

    switch (body.sheet) {
      case 'PERBAIKAN':
        result = addPerbaikan(ss, body);
        break;
      case 'UANG_JALAN':
        result = addUangJalan(ss, body);
        break;
      case 'TAGIHAN_CICILAN':
        result = addTagihanCicilan(ss, body);
        break;
      default:
        throw new Error('sheet tidak dikenali: ' + body.sheet);
    }

    return ContentService.createTextOutput(
      JSON.stringify({ status: 'ok', row: result })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: 'error', message: err.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function nextNo(sheet, col, startRow) {
  var lastRow = sheet.getLastRow();
  if (lastRow < startRow) return 1;
  var values = sheet.getRange(startRow, col, lastRow - startRow + 1, 1).getValues();
  var max = 0;
  values.forEach(function(r) {
    var n = Number(r[0]);
    if (!isNaN(n) && n > max) max = n;
  });
  return max + 1;
}

function addPerbaikan(ss, body) {
  var sheet = ss.getSheetByName('PERBAIKAN');
  var row = sheet.getLastRow() + 1;
  var no = nextNo(sheet, 1, 3);

  sheet.getRange(row, 1).setValue(no);            // A NO
  sheet.getRange(row, 2).setValue(body.tanggal);  // B TANGGAL
  sheet.getRange(row, 3).setValue(body.tindakan); // C TINDAKAN
  sheet.getRange(row, 4).setValue(body.harga);    // D HARGA
  sheet.getRange(row, 5).setValue(body.bayarKeBru); // E BAYAR KE BRU
  sheet.getRange(row, 6).setFormula('=D' + row);  // F mirror formula
  sheet.getRange(row, 7).setValue(body.keterangan || ''); // G KETERANGAN

  return row;
}

function addUangJalan(ss, body) {
  var sheet = ss.getSheetByName('UANG JALAN');
  var row = sheet.getLastRow() + 1;
  var no = nextNo(sheet, 1, 3);

  sheet.getRange(row, 1).setValue(no);              // A NO
  sheet.getRange(row, 2).setValue(body.tanggal);    // B TANGGAL
  sheet.getRange(row, 3).setValue(body.noDoSpe);    // C NO DO/SPE
  sheet.getRange(row, 4).setValue(body.tujuan);     // D TUJUAN
  sheet.getRange(row, 5).setValue(body.uangJalan);  // E UANG JALAN
  // F, G sengaja dilewati (hidden/tidak dipakai)
  sheet.getRange(row, 8).setValue(body.invoice);    // H INVOICE
  sheet.getRange(row, 9).setValue(body.customer);   // I CUSTOMER

  return row;
}

function addTagihanCicilan(ss, body) {
  var sheet = ss.getSheetByName('TAGIHAN DAN CICILAN');
  var row = sheet.getLastRow() + 1;
  var no = nextNo(sheet, 1, 3);

  sheet.getRange(row, 1).setValue(no);               // A NO
  sheet.getRange(row, 2).setValue(body.bulan);       // B BULAN
  sheet.getRange(row, 3).setValue(body.tagihan);     // C TAGIHAN
  sheet.getRange(row, 4).setValue(body.cicilan);     // D CICILAN
  sheet.getRange(row, 5).setValue(body.operasional); // E OPERASIONAL

  return row;
}
