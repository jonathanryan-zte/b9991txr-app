  // Gudang asal untuk armada ini selalu sama, jadi tidak perlu (dan tidak boleh)
  // ikut ditebak dari hasil OCR -- yang dibaca cukup kota tujuannya saja.
  const GUDANG_ASAL = "TFJ Pandeglang";

  function levenshtein(a, b) {
    if (a === b) return 0;
    if (!a.length) return b.length;
    if (!b.length) return a.length;
    let prev = [];
    for (let j = 0; j <= b.length; j++) prev[j] = j;
    for (let i = 1; i <= a.length; i++) {
      const cur = [i];
      for (let j = 1; j <= b.length; j++) {
        const cost = a.charAt(i - 1) === b.charAt(j - 1) ? 0 : 1;
        cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + cost);
      }
      prev = cur;
    }
    return prev[b.length];
  }

  // Daftar kota tujuan yang pernah dipakai, diambil dari histori Sheet.
  // Dipakai untuk membetulkan typo hasil OCR (mis. "SERAWG" -> "SERANG").
  function knownKotaList() {
    const seen = {};
    Object.keys(tujuanLookup || {}).forEach((key) => {
      const entry = tujuanLookup[key];
      const label = String((entry && entry.tujuan) || key);
      const parts = label.split("-");
      if (parts.length < 2) return;
      const kota = parts[parts.length - 1].trim().toUpperCase();
      if (kota) seen[kota] = true;
    });
    return Object.keys(seen);
  }

  function bestKotaMatch(raw) {
    const cand = String(raw || "").toUpperCase().replace(/[^A-Z0-9 ]/g, "").trim();
    if (!cand) return null;
    const list = knownKotaList();
    let best = null;
    let bestDist = Infinity;
    list.forEach((kota) => {
      const d = levenshtein(cand, kota);
      if (d < bestDist) { bestDist = d; best = kota; }
    });
    if (!best) return null;
    // toleransi sekitar sepertiga panjang kata, minimal 1 huruf
    const tol = Math.max(1, Math.floor(Math.max(cand.length, best.length) / 3));
    if (bestDist > tol) return null;
    return { kota: best, dist: bestDist };
  }

  function looksLikeGudangAsal(word) {
    const u = String(word || "").toUpperCase().replace(/[^A-Z]/g, "");
    if (!u) return true;
    return levenshtein(u, "TFJ") <= 1 || levenshtein(u, "PANDEGLANG") <= 4;
  }

  function cleanCustomer(str) {
    const words = String(str || "").trim().split(/\s+/).filter(Boolean);
    while (words.length > 1) {
      const last = words[words.length - 1].replace(/[,.]+$/, "");
      if (/^[^A-Za-z0-9]+$/.test(last) || (last.length <= 2 && !/^(PT|CV|UD)$/i.test(last))) {
        words.pop();
      } else {
        break;
      }
    }
    return words.join(" ").replace(/[,.]+$/, "").trim();
  }

  function toIsoDate(d, m, y) {
    let year = String(y);
    if (year.length === 2) year = "20" + year;
    return `${year}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
  }

  function parseSpr(text) {
    const lines = text.split("\n").map((l) => l.trim()).filter(Boolean);
    const result = {};

    // Tanggal diambil dari field "Tgl Muat" kalau ada; kalau tidak terbaca,
    // baru jatuh ke tanggal pertama yang muncul di dokumen.
    const muatMatch = text.match(
      /(?:tgl|tanggal)?\.?\s*muat[^0-9]{0,15}(\d{1,2})[.\-\/](\d{1,2})[.\-\/](\d{2,4})/i
    );
    if (muatMatch) {
      result.tanggal = toIsoDate(muatMatch[1], muatMatch[2], muatMatch[3]);
    } else {
      const dateMatch = text.match(/(\d{1,2})[.\-\/](\d{1,2})[.\-\/](\d{4})/);
      if (dateMatch) result.tanggal = toIsoDate(dateMatch[1], dateMatch[2], dateMatch[3]);
    }

    // No SPR/SPE (pojok kanan atas) jadi sumber utama No DO/SPE
    const sprMatch = text.match(/No\.?\s*SP[RE][:\.\s]*([0-9]{4,})/i);
    if (sprMatch) result.noDoSpe = sprMatch[1];
    if (!result.noDoSpe) {
      for (const line of lines) {
        const m = line.match(/No\.?\s*D\s*[\/il]?\s*[O0][:\.\s]*([0-9]{5,})/i);
        if (m) { result.noDoSpe = m[1]; break; }
      }
    }

    // Baris data utama: Tgl D/O, No D/O, (kode lain), Gudang Asal, Tujuan, Langganan
    // semuanya ada dalam SATU baris yang sama, dipisah spasi. Contoh nyata:
    // "9-08-2026 76781579 1708 TFJ Pandeglang SERANG SEKAR ARUM MAKMUR, P"
    const dataLineMatch = text.match(/\d{1,2}[.\-\/]\d{1,2}[.\-\/]\d{2,4}\s+(\d{5,})\s+(.+)/);
    let dataLineIdx = -1;
    if (dataLineMatch) {
      if (!result.noDoSpe) result.noDoSpe = dataLineMatch[1];
      dataLineIdx = lines.findIndex((l) => l.indexOf(dataLineMatch[1]) !== -1);

      var rest = dataLineMatch[2];
      // buang kode angka tambahan di depan (mis. jam bongkar "1708")
      rest = rest.replace(/^(\s*\d{1,6}\s+)+/, "");
      var words = rest.trim().split(/\s+/).filter(Boolean);

      // buang token gudang asal ("TFJ Pandeglang", termasuk versi typo OCR
      // seperti "TFQ Pandegla") -- nilainya sudah pasti, tidak perlu dibaca.
      var dibuang = 0;
      while (words.length > 1 && dibuang < 3 && looksLikeGudangAsal(words[0])) {
        words.shift();
        dibuang++;
      }

      // Kota tujuan: coba cocokkan 1 atau 2 kata pertama ke daftar kota historis,
      // supaya typo hasil OCR otomatis dibetulkan.
      var kota = null;
      var kotaLen = 0;
      var kandidat = [];
      if (words.length >= 1) kandidat.push({ raw: words[0], len: 1 });
      if (words.length >= 2) kandidat.push({ raw: words[0] + " " + words[1], len: 2 });

      var terbaik = null;
      kandidat.forEach(function (k) {
        var m = bestKotaMatch(k.raw);
        if (m && (!terbaik || m.dist < terbaik.dist)) {
          terbaik = { kota: m.kota, dist: m.dist, len: k.len };
        }
      });

      if (terbaik) {
        kota = terbaik.kota;
        kotaLen = terbaik.len;
      } else if (words.length) {
        // tidak ada rute historis yang cocok: pakai apa adanya
        kota = words[0].toUpperCase().replace(/[^A-Z0-9]/g, "");
        kotaLen = kota ? 1 : 0;
      }

      if (kota) result.tujuan = `${GUDANG_ASAL} - ${kota}`;

      var customer = cleanCustomer(words.slice(kotaLen).join(" "));
      if (customer.length >= 3) result.customer = customer;
    }

    // fallback untuk format dokumen lain: label dan isi di baris/segmen yang sama
    if (!result.tujuan) {
      for (const line of lines) {
        const m = line.match(/Tujuan[:\.\s]+(.{3,})/i);
        if (m) {
          const cocok = bestKotaMatch(m[1].split(/\s+/)[0]);
          result.tujuan = `${GUDANG_ASAL} - ${cocok ? cocok.kota : m[1].trim().toUpperCase()}`;
          break;
        }
      }
    }
    if (!result.customer) {
      for (const line of lines) {
        const m = line.match(/Langganan[:\.\s]+(.{3,})/i);
        if (m) { result.customer = cleanCustomer(m[1]); break; }
      }
    }
    if (!result.customer && dataLineIdx !== -1) {
      // kadang nama customer terpotong ke baris berikutnya oleh OCR
      for (let i = dataLineIdx + 1; i < Math.min(lines.length, dataLineIdx + 3); i++) {
        const kandidatBaris = cleanCustomer(lines[i].replace(/[^A-Za-z0-9,.\s]/g, " "));
        const hurufBesar = kandidatBaris.replace(/[^A-Z]/g, "").length;
        if (kandidatBaris.length >= 6 && hurufBesar >= kandidatBaris.length * 0.6) {
          result.customer = kandidatBaris;
          break;
        }
      }
    }

    return result;
  }

const tujuanLookup = {
  "tfj pandeglang - serang": { tujuan: "TFJ Pandeglang - SERANG", uangJalan: 450000 },
  "tfj pandeglang - tangerang": { tujuan: "TFJ Pandeglang - TANGERANG", uangJalan: 550000 }
};

const samples = {
  "asli-1": "PayLBAA YewawWl] MEOE6HOWEWETOAY\nNo.SPE : 76781579\nTgl Muat : 09-08-2026\nTgl D/O No D/O Gudang Asal Tujuan Langganan\n9-08-2026 76781579 1708 TFQ Pandegla SERANG SEKAR ARUM MAKMUR, P",
  "kota-typo": "No.SPE : 76781580\nTgl Muat : 10-08-2026\n10-08-2026 76781580 1708 TFJ Pandegla TANGERAWG PT MAJU JAYA ABADI",
  "customer-baris-lain": "No.SPE : 76781581\nTgl Muat: 11-08-2026\n11-08-2026 76781581 TFJ Pandegla SERANG\nSEKAR ARUM MAKMUR",
  "tanpa-tgl-muat": "No.SPE : 76781582\n12-08-2026 76781582 1708 TFJ Pandegla SERANG TOKO SUMBER REJEKI"
};

Object.keys(samples).forEach((k) => {
  console.log(k, JSON.stringify(parseSpr(samples[k])));
});
