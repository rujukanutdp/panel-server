// server.js
const express = require('express');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const multer = require('multer');
const cors = require('cors');

const app = express();
app.use(express.static(__dirname));
app.use(cors());
app.use(express.json());

// ----- Configuration -----
const DATA_DIR = path.join(__dirname, 'data');
const EXCEL_FILENAME = 'antigram.xlsx';
const EXCEL_PATH = path.join(DATA_DIR, EXCEL_FILENAME);

// Upload token (optional) for security. Set env UPLOAD_TOKEN to enable simple protection.
const UPLOAD_TOKEN = process.env.UPLOAD_TOKEN || null;

// Ensure data dir exists
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

// Serve static client
app.use('/', express.static(path.join(__dirname, 'static')));

// Helper: read Excel and convert to structured JSON
function readExcelToPanel() {
  if (!fs.existsSync(EXCEL_PATH)) {
    return { ok: false, error: `File not found: ${EXCEL_PATH}` };
  }

  const workbook = XLSX.readFile(EXCEL_PATH, { cellNF: false, cellDates: true });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });

  // Find header row (try to detect row that contains 'Sel' and 'Ref')
  let headerRowIndex = null;
  for (let i = 0; i < Math.min(10, aoa.length); i++) {
    const row = aoa[i].map(c => (c || '').toString().trim().toLowerCase());
    if (row.includes('sel') && (row.includes('ref') || row.includes('ref.'))) {
      headerRowIndex = i;
      break;
    }
  }
  if (headerRowIndex === null) headerRowIndex = 2; // fallback

  const header = (aoa[headerRowIndex] || []).map(h => (h || '').toString().trim());
  // data rows: rows after header until a fully empty row
  const dataRows = [];
  for (let r = headerRowIndex + 1; r < aoa.length; r++) {
    const row = aoa[r];
    if (!row || row.every(c => c === undefined || c === null || String(c).trim() === '')) break;
    dataRows.push(row);
  }

  // Map header names (lowercase) -> index
  const idx = {};
  header.forEach((h, i) => {
    if (!h) return;
    idx[h.toString().toLowerCase()] = i;
  });

  const cell = (row, nameCandidates = []) => {
    for (const name of nameCandidates) {
      const k = name.toString().toLowerCase();
      if (k in idx) {
        const v = row[idx[k]];
        if (v === undefined || v === null) return '';
        return String(v).trim();
      }
    }
    return '';
  };

  // collect meta info (scan rows above header)
  const meta = { merk: '', lot: '', exp: '' };
  for (let r = 0; r < headerRowIndex; r++) {
    const row = aoa[r] || [];
   const lower = row.map(c => (c || '').toString().toLowerCase());

  // Cari kata kunci di kolom mana pun
  lower.forEach((cell, i) => {
    if (cell.includes('merk') && row[i + 1]) meta.merk = row[i + 1].toString().trim();
    if ((cell.includes('lot') || cell.includes('no. lot') || cell.includes('no.lot')) && row[i + 1]) meta.lot = row[i + 1].toString().trim();
    if (cell.includes('exp') && row[i + 1]) meta.exp = row[i + 1].toString().trim();
  });
}

  const result = { meta, cells: [], auto: { '20c': '', '37c': '', 'iat': '', 'gel': '' } };

  // parse rows
  dataRows.forEach(row => {
    // detect auto row: any cell containing 'auto' or 'auto kontrol'
    const anyAuto = row.some(c => (c || '').toString().toLowerCase().includes('auto'));
    if (anyAuto) {
      // try to read last 4 columns or named columns
      result.auto['20c'] = cell(row, ['20oc','20oc','20oc','20oc','20oc','20oc','20oc','20oc']) || cell(row, ['20oc','20oc']) || cell(row, ['20oC','20°C']) || (row[header.length - 4] || '').toString().trim();
      result.auto['37c'] = cell(row, ['37oc','37oC','37°C']) || (row[header.length - 3] || '').toString().trim();
      result.auto['iat'] = cell(row, ['ict','iat','ICT','IAT']) || (row[header.length - 2] || '').toString().trim(); // ICT->IAT
      result.auto['gel'] = cell(row, ['gel']) || (row[header.length - 1] || '').toString().trim();
      return;
    }

    // normal sel row
    const sel = cell(row, ['sel']) || (row[0] || '').toString().trim();
    const ref = cell(row, ['ref', 'ref.']) || (row[1] || '').toString().trim();

    // antigen mapping attempt: try header names; fallback to positional mapping based on typical layout
    const ant = {};
    const mapNames = [
      ['d'], ['c'], ['e'], // careful: 'C' and 'c' collision handled by header detection
      ['k','kell'], ['k','kell'],
      ['fya'], ['fyb'],
      ['jka'], ['jkb'],
      ['m'], ['n'], ['s'], ['s'],
      ['p1'], ['lea'], ['leb'], ['lua'], ['lub']
    ];
    // Better to fetch by explicit common names:
    ant.D = cell(row, ['d','D']);
    ant.C = cell(row, ['c','C']);
    ant.E = cell(row, ['e','E']);
    // fallback: since header may contain both uppercase and lowercase same label, attempt by header indexes directly
    ant.c = cell(row, ['c','C']);
    ant.e = cell(row, ['e','E']);
    ant.K = cell(row, ['k','K','kell']);
    ant.k = cell(row, ['k','K','kell']);
    ant.Fya = cell(row, ['fya']);
    ant.Fyb = cell(row, ['fyb']);
    ant.Jka = cell(row, ['jka']);
    ant.Jkb = cell(row, ['jkb']);
    ant.M = cell(row, ['m']);
    ant.N = cell(row, ['n']);
    ant.S = cell(row, ['s']);
    ant.s = cell(row, ['s']);
    ant.P1 = cell(row, ['p1']);
    ant.Lea = cell(row, ['lea']);
    ant.Leb = cell(row, ['leb']);
    ant.Lua = cell(row, ['lua']);
    ant.Lub = cell(row, ['lub']);

    // fallback positional extraction if many empty
    // assume typical Excel layout (Ref at col 1, then D.. etc)
    // try minimal safe positional fallbacks (only when key is empty)
    const fallback = (index) => (row[index] === undefined ? '' : String(row[index]).trim());
    if (!ant.D) ant.D = fallback(2);
    if (!ant.C) ant.C = fallback(3);
    if (!ant.E) ant.E = fallback(4);
    if (!ant.c) ant.c = fallback(5);
    if (!ant.e) ant.e = fallback(6);
    if (!ant.K) ant.K = fallback(7);
    if (!ant.k) ant.k = fallback(8);
    if (!ant.Fya) ant.Fya = fallback(9);
    if (!ant.Fyb) ant.Fyb = fallback(10);
    if (!ant.Jka) ant.Jka = fallback(11);
    if (!ant.Jkb) ant.Jkb = fallback(12);
    if (!ant.M) ant.M = fallback(13);
    if (!ant.N) ant.N = fallback(14);
    if (!ant.S) ant.S = fallback(15);
    if (!ant.s) ant.s = fallback(16);
    if (!ant.P1) ant.P1 = fallback(17);
    if (!ant.Lea) ant.Lea = fallback(18);
    if (!ant.Leb) ant.Leb = fallback(19);
    if (!ant.Lua) ant.Lua = fallback(20);
    if (!ant.Lub) ant.Lub = fallback(21);

    result.cells.push({
      sel,
      ref,
      antigen: ant
    });
  });

  return { ok: true, source: sheetName, header, data: result };
}

// GET /panel.json
app.get('/panel.json', (req, res) => {
  try {
    // Pastikan file selalu dibaca ulang
    delete require.cache[require.resolve('xlsx')];
    
    const out = readExcelToPanel();
    if (!out.ok) return res.status(404).json(out);
    return res.json(out);
  } catch (err) {
    console.error(err);
    return res.status(500).json({ ok: false, error: err.message });
  }
});

// Upload endpoint (replace antigram.xlsx)
// Protected with simple token if UPLOAD_TOKEN env var is set
const upload = multer({ dest: path.join(__dirname, 'uploads') });
app.post('/upload', upload.single('file'), (req, res) => {
  try {
    if (UPLOAD_TOKEN) {
      const token = req.headers['x-upload-token'] || req.body.token;
      if (!token || token !== UPLOAD_TOKEN) {
        if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
        return res.status(401).json({ ok: false, error: 'Unauthorized' });
      }
    }
    if (!req.file) return res.status(400).json({ ok: false, error: 'No file posted (field name: file)' });
    // move file to data/antigram.xlsx (overwrite)
    const tmp = req.file.path;
    const dest = EXCEL_PATH;
    fs.copyFileSync(tmp, dest);
    fs.unlinkSync(tmp);
    return res.json({ ok: true, message: `Replaced ${EXCEL_FILENAME}` });
  } catch (err) {
    return res.status(500).json({ ok: false, error: err.message });
  }
});

const port = process.env.PORT || 3000;
app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "panel_client.html"));
});

app.listen(port, () => {
  console.log(`Server berjalan di http://localhost:${port}`);
});



