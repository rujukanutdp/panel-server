const express = require('express');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const multer = require('multer');
const cors = require('cors');

const app = express();
app.use(cors());
app.use(express.json());

const DATA_DIR = path.join(__dirname, 'data');
const EXCEL_FILENAME = 'antigram.xlsx';
const EXCEL_PATH = path.join(DATA_DIR, EXCEL_FILENAME);

if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'panel_client.html'));
});

function norm(v){ return (v === undefined || v === null) ? '' : String(v).trim(); }
function normLower(v){ return norm(v).toLowerCase(); }

function readExcelToPanel(){
  if (!fs.existsSync(EXCEL_PATH)) {
    return { ok:false, error: `File not found: ${EXCEL_PATH}` };
  }

  const workbook = XLSX.readFile(EXCEL_PATH, { cellNF:false, cellDates:true });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const aoa = XLSX.utils.sheet_to_json(sheet, { header:1, defval: '' });

  let headerRowIndex = null;
  for (let i=0; i < Math.min(10, aoa.length); i++){
    const row = (aoa[i] || []).map(c => normLower(c));
    if (row.some(c => c === 'sel' || c === 'cell' || c === 'no' || c === 'no.' )){
      headerRowIndex = i;
      break;
    }
  }
  if (headerRowIndex === null) headerRowIndex = 1;

  const headerRaw = aoa[headerRowIndex] || [];
  const header = headerRaw.map(h => norm(h));

  const meta = { merk:'', lot:'', exp:'' };
  for (let r = 0; r < headerRowIndex; r++){
    const row = aoa[r] || [];
    for (let c = 0; c < row.length; c++){
      const v = normLower(row[c]);
      if (v.includes('merk')) meta.merk = norm(row[c+1] || '');
      if (v.includes('lot') || v.includes('no.lot') || v.includes('no. lot')) meta.lot = norm(row[c+1] || '');
      if (v.includes('exp') || v.includes('expiry') || v.includes('kedaluwarsa')) meta.exp = norm(row[c+1] || '');
    }
  }

  const testsLike = ['20','20°','37','37°','iat','gel','hasil'];
  let antigenHeaders = [];
  if (header.length >= 3){
    for (let i = 2; i < header.length; i++){
      const h = norm(header[i]);
      const low = normLower(h);
      if (testsLike.some(t => low.includes(t))) break;
      if (h !== '') antigenHeaders.push(h);
    }
  }
  if (antigenHeaders.length === 0){
    for (let i = 2; i < header.length; i++){
      const h = norm(header[i]);
      if (h !== '') antigenHeaders.push(h);
    }
  }

  const dataRows = [];
  for (let r = headerRowIndex + 1; r < aoa.length; r++){
    const row = aoa[r];
    if (!row || row.every(c => c === '' || c === null || c === undefined)) break;
    dataRows.push(row);
  }

  const cells = [];
  for (const row of dataRows){
    const sel = norm(row[0]);
    const anyAuto = (row || []).some(c => normLower(c).includes('auto'));
    if (anyAuto){
      cells.push({ sel: 'Auto Kontrol', ref:'', antigen:{}, isAuto:true, raw: row });
      continue;
    }
    const ref = norm(row[1]);
    const antigen = {};
    antigenHeaders.forEach((hdr, idx) => {
      const colIdx = 2 + idx;
      antigen[hdr] = (colIdx < row.length) ? norm(row[colIdx]) : '';
    });
    cells.push({ sel: sel || '', ref, antigen });
  }

  return { ok:true, header, data:{ meta, cells, antigenHeaders }, source: sheetName };
}

app.get('/panel.json', (req, res) => {
  try {
    const out = readExcelToPanel();
    if (!out.ok) return res.status(404).json(out);
    return res.json(out);
  } catch (err) {
    console.error(err);
    return res.status(500).json({ ok:false, error: err.message });
  }
});

const upload = multer({ dest: path.join(__dirname, 'uploads') });
app.post('/upload', upload.single('file'), (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ ok:false, error:'No file posted (field name: file)' });
    const tmp = req.file.path;
    if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
    fs.copyFileSync(tmp, EXCEL_PATH);
    fs.unlinkSync(tmp);
    return res.json({ ok:true, message: 'Replaced antigram.xlsx' });
  } catch (err) {
    return res.status(500).json({ ok:false, error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server listening on http://localhost:${PORT}`));
