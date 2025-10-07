// server.js
// Dynamic Excel -> panel.json server
// - Baca antigram.xlsx dari folder ./data/antigram.xlsx
// - Header detection: kolom 0 = Sel, kolom 1 = Ref, kolom 2...N = antigen (dinamis)
// - Mendeteksi baris "Auto Kontrol" (kata 'auto' / 'auto kontrol' pada salah satu sel) -> isi auto.{20c,37c,iat,gel}
// - Endpoint:
//    GET  /panel.json   => JSON struktur lengkap
//    POST /upload       => upload file (multipart form field name "file") -> akan menimpa data/antigram.xlsx
// - Serve static client (panel_client.html) on GET /

const express = require('express');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const multer = require('multer');
const cors = require('cors');

const app = express();
app.use(cors());
app.use(express.json());

// CONFIG
const DATA_DIR = path.join(__dirname, 'data');
const EXCEL_FILENAME = 'antigram.xlsx';
const EXCEL_PATH = path.join(DATA_DIR, EXCEL_FILENAME);

// optional: allow simple upload token protection (set env UPLOAD_TOKEN)
const UPLOAD_TOKEN = process.env.UPLOAD_TOKEN || null;

// ensure data dir
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

// serve client file (panel_client.html should be in same dir as server.js)
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'panel_client.html'));
});

// Serve static files (if you want a static folder)
// app.use('/static', express.static(path.join(__dirname, 'static')));

// Helper: safe lower-case normalized string
function norm(v){
  if(v === undefined || v === null) return '';
  return String(v).trim();
}
function normLower(v){ return norm(v).toLowerCase(); }

// Helper: read Excel and return dynamic JSON
function readExcelToPanel(){
  if (!fs.existsSync(EXCEL_PATH)) {
    return { ok:false, error: `File not found: ${EXCEL_PATH}` };
  }

  // Read workbook
  const workbook = XLSX.readFile(EXCEL_PATH, { cellNF: false, cellDates: true });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  // Convert sheet to array-of-arrays
  const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });

  // find header row index by searching for a row that contains 'sel' and 'ref' (case-insensitive)
  let headerRowIndex = null;
  for(let r = 0; r < Math.min(15, aoa.length); r++){
    const row = (aoa[r] || []).map(c => normLower(c));
    if(row.some(c => c === 'sel' || c === 'cell' || c === 'sel.')) {
      // require also presence of 'ref' or assume second column is ref
      if(row.some(c => c === 'ref' || c === 'ref.' || c === 'no.ref' || c === 'no. ref' || c === 'reference') || row.length > 1){
        headerRowIndex = r;
        break;
      }
    }
  }
  if(headerRowIndex === null){
    // fallback: pick row 0 or 1 if it looks like header
    headerRowIndex = 0;
  }

  const headerRaw = aoa[headerRowIndex] || [];
  const header = headerRaw.map(h => norm(h));

  // collect meta info from rows above header (merk, lot, exp)
  const meta = { merk: '', lot: '', exp: '' };
  for(let r = 0; r < headerRowIndex; r++){
    const row = aoa[r] || [];
    for(let c = 0; c < (row.length || 0); c++){
      const v = normLower(row[c]);
      if(v.includes('merk')) meta.merk = norm(row[c+1] || row[c+1]);
      if(v.includes('lot') || v.includes('no.lot') || v.includes('no. lot')) meta.lot = norm(row[c+1] || row[c+1]);
      if(v.includes('exp') || v.includes('expiry') || v.includes('kedaluwarsa')) meta.exp = norm(row[c+1] || row[c+1]);
    }
  }

  // data rows start after header
  const dataRows = [];
  for(let r = headerRowIndex + 1; r < aoa.length; r++){
    const row = aoa[r];
    // stop at fully empty row
    if(!row || row.every(c => c === undefined || c === null || norm(c) === '')) break;
    dataRows.push(row);
  }

  // Determine antigen headers:
  // We assume header[0] = Sel, header[1] = Ref, antigen headers start at index 2
  // We'll stop the antigen list if we encounter header labels that clearly indicate tests (rare in the Excel).
  const testsLike = ['20','20°','20°c','37','37°','37°c','iat','gel','hasil','hasil pemeriksaan'];
  let antigenHeaders = [];
  if(header.length >= 3){
    for(let i = 2; i < header.length; i++){
      const h = norm(header[i]);
      const low = normLower(h);
      // if header seems like test columns, break
      if(testsLike.some(t => low.includes(t))) break;
      // ignore empty header
      if(h !== '') antigenHeaders.push(h);
    }
  }

  // If antigenHeaders is empty, fallback to scanning dataRows for union of keys (rare)
  if(antigenHeaders.length === 0){
    const set = new Set();
    dataRows.forEach(row => {
      // treat columns index >=2 as antigen
      for(let c = 2; c < row.length; c++){
        const v = header[c] ? norm(header[c]) : '';
        if(v) set.add(v);
      }
    });
    antigenHeaders = Array.from(set);
  }

  // Build cells: for each data row, map sel, ref, and antigen object
  const cells = [];
  dataRows.forEach(row => {
    // get sel and ref
    const sel = (row[0] !== undefined && row[0] !== null) ? norm(row[0]) : '';
    const ref = (row[1] !== undefined && row[1] !== null) ? norm(row[1]) : '';

    // build antigen object mapping header -> value (try case-insensitive matching)
    const antigen = {};
    // If headerRaw had names, we will map based on header indices
    antigenHeaders.forEach(hdr => {
      // attempt to find hdr in headerRaw (case-insensitive) -> get its index
      let idx = headerRaw.findIndex(x => normLower(x) === normLower(hdr));
      if(idx === -1){
        // fallback: try to find by trimmed match
        idx = headerRaw.findIndex(x => norm(x) === norm(hdr));
      }
      // if still -1, fallback to sequential mapping starting at col 2
      if(idx === -1){
        // find position of hdr in antigenHeaders to map to column index 2 + pos
        const pos = antigenHeaders.indexOf(hdr);
        idx = 2 + pos;
      }
      const rawVal = (row[idx] === undefined || row[idx] === null) ? '' : String(row[idx]).trim();
      antigen[hdr] = rawVal;
    });

    cells.push({ sel, ref, antigen });
  });

  // Detect Auto Kontrol row if any row contains 'auto' string (case-insensitive) in any cell.
  // Attempt to read last 4 columns of that auto-row as 20c,37c,iat,gel OR read by header names if available.
  const auto = { '20c':'', '37c':'', 'iat':'', 'gel':'' };
  let autoFound = false;
  for(let r = headerRowIndex + 1; r < aoa.length; r++){
    const row = aoa[r] || [];
    const anyAuto = row.some(c => typeof c === 'string' && normLower(c).includes('auto'));
    if(anyAuto){
      autoFound = true;
      // try to get by header indexes if they exist
      // find header indexes for test-like headers
      const idx20 = headerRaw.findIndex(h => h && normLower(h).includes('20'));
      const idx37 = headerRaw.findIndex(h => h && normLower(h).includes('37'));
      const idxIAT = headerRaw.findIndex(h => h && normLower(h).includes('iat'));
      const idxGel = headerRaw.findIndex(h => h && normLower(h).includes('gel'));
      if(idx20 !== -1) auto['20c'] = norm(row[idx20]);
      if(idx37 !== -1) auto['37c'] = norm(row[idx37]);
      if(idxIAT !== -1) auto['iat'] = norm(row[idxIAT]);
      if(idxGel !== -1) auto['gel'] = norm(row[idxGel]);

      // fallback: try read last 4 columns of that row
      if(!auto['20c'] && row.length >= 4) auto['20c'] = norm(row[row.length - 4]);
      if(!auto['37c'] && row.length >= 3) auto['37c'] = norm(row[row.length - 3]);
      if(!auto['iat'] && row.length >= 2) auto['iat'] = norm(row[row.length - 2]);
      if(!auto['gel'] && row.length >= 1) auto['gel'] = norm(row[row.length - 1]);
      break;
    }
  }

  return {
    ok: true,
    header,
    data: {
      meta,
      cells,
      auto: autoFound ? auto : { '20c':'','37c':'','iat':'','gel':'' }
    },
    sourceSheet: sheetName
  };
}

// GET /panel.json
app.get('/panel.json', (req, res) => {
  try {
    const out = readExcelToPanel();
    if(!out.ok) return res.status(404).json(out);
    return res.json(out);
  } catch(err) {
    console.error('readExcelToPanel error:', err);
    return res.status(500).json({ ok:false, error: err.message });
  }
});

// Upload endpoint: field name "file"
const upload = multer({ dest: path.join(__dirname, 'uploads') });
app.post('/upload', upload.single('file'), (req, res) => {
  try {
    // optional token protection
    if (UPLOAD_TOKEN) {
      const token = req.headers['x-upload-token'] || req.body.token;
      if (!token || token !== UPLOAD_TOKEN) {
        if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
        return res.status(401).json({ ok:false, error:'Unauthorized' });
      }
    }
    if (!req.file) return res.status(400).json({ ok:false, error: 'No file posted (field name: file)' });

    // move uploaded to data/antigram.xlsx (overwrite)
    const tmp = req.file.path;
    const dest = EXCEL_PATH;
    // ensure data dir
    if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
    fs.copyFileSync(tmp, dest);
    fs.unlinkSync(tmp);

    return res.json({ ok:true, message: `Replaced ${EXCEL_FILENAME}` });
  } catch(err) {
    console.error('upload error', err);
    return res.status(500).json({ ok:false, error: err.message });
  }
});

// Health
app.get('/_health', (req, res) => res.json({ ok:true, now: new Date().toISOString() }));

const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log(`Panel server listening on port ${port} — GET /panel.json, POST /upload (field 'file')`);
});
