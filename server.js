import express from "express";
import fs from "fs";
import path from "path";
import xlsx from "xlsx";

const app = express();
const PORT = 3000;
app.use(express.static("."));

// --- Fungsi Aman Parsing Excel ---
function parseAntigramExcel(filePath) {
  const wb = xlsx.readFile(filePath);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const raw = xlsx.utils.sheet_to_json(ws, { header: 1, defval: "" });

  if (!raw || raw.length < 2) {
    return { ok: false, error: "File antigram kosong atau format salah." };
  }

  // Jika baris pertama bukan header antigen (misal meta seperti Merk/Lot/Exp)
  // maka ambil baris berikutnya sebagai header antigen.
  let headerRow = raw[0];
  const headerLower = headerRow.map(v => String(v).toLowerCase());
  const hasSelHeader = headerLower.includes("sel") || headerLower.includes("cell");

  if (!hasSelHeader && raw.length > 1) {
    headerRow = raw[1];
    raw.shift(); // buang baris meta
  }

  // Header & deteksi kolom
  const header = headerRow.map(h => String(h || "").trim());
  const testHeaders = ["20", "37", "iat", "gel"];
  const testIndex = header.findIndex(h => testHeaders.some(t => h.toLowerCase().includes(t)));
  const antigenStart = header.findIndex(h => h.toLowerCase() === "ref") + 1;
  const antigenEnd = testIndex > -1 ? testIndex - 1 : header.length - 1;
  const antigenKeys = header.slice(antigenStart, antigenEnd + 1).filter(k => k);

  // Meta info (cari baris yang berisi kata "Merk" atau "Lot")
  let meta = { merk: "-", lot: "-", exp: "-" };
  const metaRow = raw.find(r => r.join(" ").toLowerCase().includes("merk"));
  if (metaRow) {
    meta = {
      merk: metaRow[1] || "-",
      lot: metaRow[3] || "-",
      exp: metaRow[5] || "-"
    };
  }

  // Baris data mulai setelah header
  const dataRows = raw.slice(raw.indexOf(headerRow) + 1).filter(r => r.some(v => v !== ""));
  const cells = [];

  for (const row of dataRows) {
    const sel = String(row[0] || "").trim();
    if (!sel) continue;

    // Jika baris auto kontrol
    if (sel.toLowerCase().includes("auto")) {
      cells.push({
        sel: "Auto Kontrol",
        ref: "",
        antigen: {},
        isAuto: true
      });
      continue;
    }

    const ref = String(row[1] || "").trim();
    const antigenObj = {};
    antigenKeys.forEach((k, i) => {
      const colIndex = antigenStart + i;
      antigenObj[k] = row[colIndex] ? String(row[colIndex]).trim() : "";
    });

    cells.push({ sel, ref, antigen: antigenObj });
  }

  return {
    ok: true,
    header,
    data: { meta, cells }
  };
}

// --- Endpoint untuk kirim JSON ---
app.get("/panel.json", (req, res) => {
  const filePath = path.resolve("./antigram.xlsx");
  if (!fs.existsSync(filePath)) {
    return res.json({ ok: false, error: "File antigram.xlsx tidak ditemukan" });
  }

  try {
    const parsed = parseAntigramExcel(filePath);
    res.json(parsed);
  } catch (err) {
    res.json({ ok: false, error: "Gagal membaca Excel: " + err.message });
  }
});

app.listen(PORT, () => {
  console.log(`✅ Server aktif di http://localhost:${PORT}`);
});
