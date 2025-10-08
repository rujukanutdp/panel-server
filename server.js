import express from "express";
import cors from "cors";
import xlsx from "xlsx";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const app = express();

app.use(cors());
app.use(express.static(__dirname));

app.get("/panel.json", (req, res) => {
  try {
    const filePath = path.join(__dirname, "antigram.xlsx");
    const workbook = xlsx.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = xlsx.utils.sheet_to_json(sheet, { header: 1 });

    if (rows.length < 3) {
      return res.json({ ok: false, message: "File antigram tidak lengkap" });
    }

    const metaRow = rows[0];
    const meta = { merk: metaRow[0] || "-", lot: metaRow[1] || "-", exp: metaRow[2] || "-" };

    const header = rows[1].map(h => String(h || "").trim());
    const antigenStartIndex = 2;
    const antigenKeys = header.slice(antigenStartIndex);

    const cells = [];
    for (let i = 2; i < rows.length; i++) {
      const row = rows[i];
      if (!row || !row.length) continue;
      const sel = row[0] || "";
      const ref = row[1] || "";
      const antigen = {};
      antigenKeys.forEach((key, j) => {
        antigen[key] = row[antigenStartIndex + j] ?? "";
      });
      cells.push({ sel, ref, antigen });
    }

    res.json({ ok: true, header, data: { meta, cells } });
  } catch (err) {
    console.error(err);
    res.json({ ok: false, error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`✅ Server panel aktif di port ${PORT}`));
