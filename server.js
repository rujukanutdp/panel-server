import express from "express";
import fs from "fs";
import path from "path";
import xlsx from "xlsx";

const app = express();
const PORT = 3000;

// Folder statis (tempat file HTML kamu)
app.use(express.static("."));

// Fungsi: parsing Excel antigram
function parseAntigramExcel(filePath) {
  const wb = xlsx.readFile(filePath);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const raw = xlsx.utils.sheet_to_json(ws, { header: 1 });

  if (!raw || raw.length < 2) {
    return { ok: false, error: "Data Excel kosong atau tidak lengkap" };
  }

  // Baris pertama = meta (merk, lot, exp)
  const metaRow = raw[0].map(x => (x ? String(x).trim() : ""));
  const meta = {
    merk: metaRow[1] || "-",
    lot: metaRow[3] || "-",
    exp: metaRow[5] || "-"
  };

  // Baris kedua = header antigen
  const headerRow = raw[1].map(x => (x ? String(x).trim() : ""));
  let selIndex = headerRow.findIndex(x =>
    ["sel", "cell", "no", "ref"].includes(x.toLowerCase())
  );
  if (selIndex === -1) selIndex = 0;

  // Kolom antigen mulai setelah “Sel / Ref”, berhenti sebelum kolom hasil tes
  const testHeaders = ["20", "20°", "37", "37°", "iat", "gel"];
  const firstTestIndex = headerRow.findIndex(h =>
    testHeaders.some(t => h.toLowerCase().includes(t))
  );
  const antigenStart = selIndex + 1;
  const antigenEnd =
    firstTestIndex > -1 ? firstTestIndex - 1 : headerRow.length - 1;

  const antigenKeys = headerRow.slice(antigenStart, antigenEnd + 1);

  // Baris data
  const dataRows = raw.slice(2).filter(
    r => r.some(v => v !== null && v !== undefined && v !== "")
  );

  const cells = [];
  for (const row of dataRows) {
    const label = String(row[selIndex] || "").trim();

    // Baris auto kontrol
    if (label.toLowerCase().includes("auto")) {
      cells.push({
        sel: "Auto Kontrol",
        ref: "",
        antigen: {},
        isAuto: true
      });
      continue;
    }

    const antigenObj = {};
    antigenKeys.forEach((k, idx) => {
      const val = row[antigenStart + idx];
      antigenObj[k] = val !== undefined ? String(val).trim() : "";
    });

    cells.push({
      sel: label || String(cells.length + 1),
      ref: row[selIndex + 1] || "",
      antigen: antigenObj
    });
  }

  return {
    ok: true,
    header: headerRow,
    data: { meta, cells }
  };
}

// API: kirim panel JSON
app.get("/panel.json", (req, res) => {
  const filePath = path.resolve("./antigram.xlsx");
  if (!fs.existsSync(filePath)) {
    return res.json({ ok: false, error: "File antigram.xlsx tidak ditemukan" });
  }

  const parsed = parseAntigramExcel(filePath);
  res.json(parsed);
});

app.listen(PORT, () => {
  console.log(`✅ Server berjalan di http://localhost:${PORT}`);
});
