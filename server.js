const express = require("express");
const cors = require("cors");
const path = require("path");
const xlsx = require("xlsx");
const app = express();

app.use(cors());
app.use(express.static(__dirname));

function readAntigram() {
  try {
    const filePath = path.join(__dirname, "antigram.xlsx");
    const workbook = xlsx.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = xlsx.utils.sheet_to_json(sheet, { defval: "" });

    if (json.length === 0) return { meta: {}, rows: [] };

    // baris pertama = metadata
    const meta = {
      merk: json[0].Merk || json[0].merk || "-",
      lot: json[0].Lot || json[0].lot || "-",
      exp: json[0].Exp || json[0].exp || "-"
    };

    // sisanya = data tabel sel
    const rows = json.slice(1);
    return { meta, rows };
  } catch (err) {
    console.error("Gagal membaca antigram.xlsx:", err);
    return { meta: {}, rows: [] };
  }
}

app.get("/data", (req, res) => {
  res.json(readAntigram());
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`✅ Server jalan di port ${PORT}`));
