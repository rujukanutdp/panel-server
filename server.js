// server.js
const express = require("express");
const cors = require("cors");
const path = require("path");
const xlsx = require("xlsx");
const app = express();

app.use(cors());
app.use(express.static(__dirname)); // supaya panel_client.html bisa diakses langsung

// === Fungsi baca Excel ===
function readAntigram() {
  try {
    const filePath = path.join(__dirname, "antigram.xlsx");
    const workbook = xlsx.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = xlsx.utils.sheet_to_json(sheet, { defval: "" });

    // Baris pertama = metadata (merk, lot, exp)
    const meta = json[0] || {};
    // Sisanya = data sel
    const rows = json.slice(1);

    return { meta, rows };
  } catch (err) {
    console.error("Gagal baca antigram.xlsx:", err.message);
    return { meta: {}, rows: [] };
  }
}

// === Endpoint kirim data ke client ===
app.get("/data", (req, res) => {
  const data = readAntigram();
  res.json(data);
});

// === Jalankan server ===
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server jalan di port ${PORT}`));
