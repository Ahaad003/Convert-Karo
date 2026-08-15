// ConvertKaro Backend Server
// Handles: PDF -> Word, Word -> PDF, and PDF editing (via pdf-lib)

const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const os = require('os');
const cors = require('cors');
const { execFile } = require('child_process');
const { PDFDocument, rgb, StandardFonts, degrees } = require('pdf-lib');
const fontkit = require('@pdf-lib/fontkit');
const pdfjsLib = require('pdfjs-dist/legacy/build/pdf.js');
const XLSX = require('xlsx');
const JSZip = require('jszip');

const app = express();
const PORT = process.env.PORT || 8080;

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, '..', 'public')));

// Temporary upload storage (auto-cleaned)
const upload = multer({ dest: os.tmpdir() });

// In-memory store for documents being edited (docId -> file path)
// NOTE: for production with multiple users, replace with a real DB + file storage (S3 etc.)
const editSessions = {};

// ---------- LibreOffice conversion helper ----------
// Saves the input with a real file extension (e.g. .pdf, .docx) so LibreOffice
// can correctly auto-detect the format, then calls soffice directly.
// This avoids a bug in the libreoffice-convert npm package where it saves the
// temp file with NO extension, which makes LibreOffice fail to recognize PDFs
// as an input format (it only worked for Word/Excel because those formats
// happen to be detectable from content alone).
function convertWithSoffice(inputBuffer, inputExt, outputExt, callback) {
  const id = Date.now() + '-' + Math.random().toString(36).slice(2, 8);
  const tmpDir = path.join(os.tmpdir(), 'lo-' + id);
  const userInstallDir = path.join(os.tmpdir(), 'lo-profile-' + id);
  try {
    fs.mkdirSync(tmpDir, { recursive: true });
  } catch (e) {
    return callback(e);
  }
  const inputPath = path.join(tmpDir, 'input.' + inputExt);
  fs.writeFileSync(inputPath, inputBuffer);

  // When the source is a PDF, LibreOffice defaults to opening it as a Draw
  // (image/vector) document, which has no export path to docx/xlsx and fails
  // silently. Forcing the Writer PDF-import filter makes it reconstruct the
  // PDF as text/paragraphs instead, which CAN be exported to docx/xlsx.
  const args = [
    '-env:UserInstallation=file://' + userInstallDir,
    '--headless',
    '--norestore',
  ];
  if (inputExt === 'pdf') {
    args.push('--infilter=writer_pdf_import');
  }
  args.push('--convert-to', outputExt, '--outdir', tmpDir, inputPath);

  execFile('soffice', args, { timeout: 90000 }, (err, stdout, stderr) => {
    if (err) {
      cleanup();
      return callback(new Error((stderr || err.message || 'soffice failed').toString().slice(0, 400)));
    }
    const outputPath = path.join(tmpDir, 'input.' + outputExt);
    if (!fs.existsSync(outputPath)) {
      cleanup();
      return callback(new Error('Output file was not created. ' + (stdout || '').toString().slice(0, 300)));
    }
    let outBuf;
    try {
      outBuf = fs.readFileSync(outputPath);
    } catch (e) {
      cleanup();
      return callback(e);
    }
    cleanup();
    callback(null, outBuf);
  });

  function cleanup() {
    fs.rm(tmpDir, { recursive: true, force: true }, () => {});
    fs.rm(userInstallDir, { recursive: true, force: true }, () => {});
  }
}

// ---------- PDF -> Excel: real table extraction (no LibreOffice) ----------
// LibreOffice has no PDF-import filter for Calc, so a spreadsheet can't be
// reconstructed via soffice. Instead we read the PDF's text with its x/y
// position on the page (via pdfjs), group nearby text into rows and columns
// by position, and write that grid straight into an .xlsx file.
// This works well for PDFs that already look like a clean table; it won't
// be perfect for scanned PDFs or very irregular layouts.
async function extractPdfTableData(buffer) {
  const loadingTask = pdfjsLib.getDocument({ data: buffer, disableFontFace: true });
  const pdf = await loadingTask.promise;
  const pages = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const textContent = await page.getTextContent();
    const items = textContent.items
      .map(it => ({ text: it.str, x: it.transform[4], y: it.transform[5] }))
      .filter(it => it.text && it.text.trim().length > 0);
    pages.push(items);
  }
  return pages;
}

function itemsToGrid(items) {
  if (!items.length) return [[]];
  const yTolerance = 3;
  const rowBuckets = [];
  items.forEach(it => {
    let bucket = rowBuckets.find(b => Math.abs(b.y - it.y) <= yTolerance);
    if (!bucket) { bucket = { y: it.y, items: [] }; rowBuckets.push(bucket); }
    bucket.items.push(it);
  });
  rowBuckets.sort((a, b) => b.y - a.y); // PDF y grows upward, so top row first

  const xTolerance = 15;
  const xClusters = [];
  items.forEach(it => {
    let c = xClusters.find(c => Math.abs(c.center - it.x) <= xTolerance);
    if (c) { c.xs.push(it.x); c.center = c.xs.reduce((a, b) => a + b, 0) / c.xs.length; }
    else xClusters.push({ center: it.x, xs: [it.x] });
  });
  xClusters.sort((a, b) => a.center - b.center);

  return rowBuckets.map(bucket => {
    const cells = new Array(xClusters.length).fill('');
    bucket.items.sort((a, b) => a.x - b.x).forEach(it => {
      let bestIdx = 0, bestDist = Infinity;
      xClusters.forEach((c, idx) => {
        const d = Math.abs(c.center - it.x);
        if (d < bestDist) { bestDist = d; bestIdx = idx; }
      });
      cells[bestIdx] = cells[bestIdx] ? cells[bestIdx] + ' ' + it.text : it.text;
    });
    return cells;
  });
}

function gridsToXlsxBuffer(pagesGrids) {
  const wb = XLSX.utils.book_new();
  pagesGrids.forEach((grid, idx) => {
    const ws = XLSX.utils.aoa_to_sheet(grid);
    XLSX.utils.book_append_sheet(wb, ws, 'Page' + (idx + 1));
  });
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

// ---------- Word page size (A4 / A3 / Legal) post-processing ----------
// LibreOffice's PDF->Word reconstruction already tries to match the source
// PDF's page size/orientation. If the person picks a specific paper size
// instead, we rewrite the docx's page-size XML directly (docx is just a zip
// of XML files) so the exported file uses exactly that paper size.
const PAGE_SIZES_TWIPS = { // [width, height] in twips, portrait orientation
  a4: [11907, 16840],
  a3: [16840, 23811],
  legal: [12240, 20160]
};
async function setDocxPageSize(docxBuffer, sizeKey) {
  const dims = PAGE_SIZES_TWIPS[sizeKey];
  if (!dims) return docxBuffer; // 'original' or unknown -> leave untouched
  try {
    const zip = await JSZip.loadAsync(docxBuffer);
    const docFile = zip.file('word/document.xml');
    if (!docFile) return docxBuffer;
    let xml = await docFile.async('string');
    const w = dims[0], h = dims[1];
    const pgSzTag = '<w:pgSz w:w="' + w + '" w:h="' + h + '"/>';
    if (/<w:pgSz [^\/]*\/>/.test(xml)) {
      xml = xml.replace(/<w:pgSz [^\/]*\/>/g, pgSzTag);
    } else if (/<w:sectPr[^>]*>/.test(xml)) {
      xml = xml.replace(/(<w:sectPr[^>]*>)/, '$1' + pgSzTag);
    }
    zip.file('word/document.xml', xml);
    return await zip.generateAsync({ type: 'nodebuffer' });
  } catch (e) {
    console.error('setDocxPageSize error:', e);
    return docxBuffer; // fall back to the unmodified file rather than fail
  }
}

// ---------- Health check ----------
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', service: 'ConvertKaro backend' });
});

// ---------- Debug: check if soffice/gs are actually installed ----------
// Temporary diagnostic route — visit /api/debug/tools in the browser to see
// exactly what's available on the server. Safe to remove once things work.
app.get('/api/debug/tools', (req, res) => {
  const { exec } = require('child_process');
  exec('echo PATH=$PATH; echo ---; which soffice; echo ---; which libreoffice; echo ---; which gs; echo ---; ls /nix/store 2>/dev/null | grep -i libreoffice | head -5; echo ---; ls /nix/store 2>/dev/null | grep -i ghostscript | head -5',
    { timeout: 10000 },
    (err, stdout, stderr) => {
      res.type('text/plain').send('STDOUT:\n' + stdout + '\n\nSTDERR:\n' + stderr + '\n\nERR:\n' + (err ? err.message : 'none'));
    }
  );
});

// ---------- PDF -> Word ----------
app.post('/api/convert/pdf-to-word', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

  const inputPath = req.file.path;
  const inputBuf = fs.readFileSync(inputPath);
  const pageSize = (req.body && req.body.pageSize) || 'original';

  convertWithSoffice(inputBuf, 'pdf', 'docx', async (err, outBuf) => {
    fs.unlink(inputPath, () => {}); // cleanup uploaded temp file
    if (err) {
      console.error('PDF->Word conversion error:', err);
      return res.status(500).json({ error: 'Conversion failed', detail: err.message || String(err) });
    }
    let finalBuf = outBuf;
    if (pageSize !== 'original') {
      finalBuf = await setDocxPageSize(outBuf, pageSize);
    }
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename=converted.docx');
    res.send(finalBuf);
  });
});

// ---------- Word -> PDF ----------
app.post('/api/convert/word-to-pdf', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

  const inputPath = req.file.path;
  const inputBuf = fs.readFileSync(inputPath);

  convertWithSoffice(inputBuf, 'docx', 'pdf', (err, outBuf) => {
    fs.unlink(inputPath, () => {});
    if (err) {
      console.error('Word->PDF conversion error:', err);
      return res.status(500).json({ error: 'Conversion failed', detail: err.message || String(err) });
    }
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename=converted.pdf');
    res.send(outBuf);
  });
});

// ---------- Excel -> PDF ----------
app.post('/api/convert/excel-to-pdf', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

  const inputPath = req.file.path;
  const inputBuf = fs.readFileSync(inputPath);

  convertWithSoffice(inputBuf, 'xlsx', 'pdf', (err, outBuf) => {
    fs.unlink(inputPath, () => {});
    if (err) {
      console.error('Excel->PDF conversion error:', err);
      return res.status(500).json({ error: 'Conversion failed', detail: err.message || String(err) });
    }
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename=converted.pdf');
    res.send(outBuf);
  });
});

// ---------- PDF -> Excel (real text-position based table extraction) ----------
app.post('/api/convert/pdf-to-excel', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

  const inputPath = req.file.path;
  try {
    const inputBuf = fs.readFileSync(inputPath);
    const pagesItems = await extractPdfTableData(inputBuf);
    const grids = pagesItems.map(itemsToGrid);
    const outBuf = gridsToXlsxBuffer(grids);
    fs.unlink(inputPath, () => {});
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=converted.xlsx');
    res.send(outBuf);
  } catch (err) {
    fs.unlink(inputPath, () => {});
    console.error('PDF->Excel extraction error:', err);
    res.status(500).json({ error: 'Extraction failed', detail: err.message || String(err) });
  }
});

// ---------- Compress PDF ----------
// Uses Ghostscript to re-encode/downsample images inside the PDF.
app.post('/api/compress', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

  const inputPath = req.file.path;
  const outputPath = path.join(os.tmpdir(), 'compressed-' + Date.now() + '.pdf');

  // /screen = smallest size (lower quality), /ebook = good balance (default here)
  const gsArgs = [
    '-sDEVICE=pdfwrite',
    '-dCompatibilityLevel=1.4',
    '-dPDFSETTINGS=/ebook',
    '-dNOPAUSE', '-dQUIET', '-dBATCH',
    '-sOutputFile=' + outputPath,
    inputPath
  ];

  execFile('gs', gsArgs, (err) => {
    fs.unlink(inputPath, () => {});
    if (err) {
      console.error('Compress error:', err);
      return res.status(500).json({ error: 'Compression failed', detail: err.message });
    }
    res.download(outputPath, 'compressed.pdf', () => {
      fs.unlink(outputPath, () => {});
    });
  });
});

// ---------- PDF Editor: upload & start session ----------
app.post('/api/edit/upload', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

  const docId = Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
  const storedPath = path.join(os.tmpdir(), 'edit-' + docId + '.pdf');
  fs.renameSync(req.file.path, storedPath);
  editSessions[docId] = storedPath;

  res.json({ docId });
});

// ---------- PDF Editor: get page count / basic info ----------
app.get('/api/edit/:docId/info', async (req, res) => {
  const filePath = editSessions[req.params.docId];
  if (!filePath || !fs.existsSync(filePath)) {
    return res.status(404).json({ error: 'Document not found or session expired' });
  }
  const bytes = fs.readFileSync(filePath);
  const pdfDoc = await PDFDocument.load(bytes);
  res.json({ pageCount: pdfDoc.getPageCount() });
});

// ---------- PDF Editor: apply a text correction (white-box + new text overlay) ----------
// body: { pageIndex, x, y, width, height, newText, fontSize, color:{r,g,b} }
// Supports Hindi (Devanagari) text if font files are present in server/fonts/
// (see README for the two font files you need to add — StandardFonts can only draw English/Latin text).
app.post('/api/edit/:docId/apply', async (req, res) => {
  const filePath = editSessions[req.params.docId];
  if (!filePath || !fs.existsSync(filePath)) {
    return res.status(404).json({ error: 'Document not found or session expired' });
  }
  const { pageIndex, x, y, width, height, newText, fontSize, color, bold, italic, underline } = req.body;

  try {
    const bytes = fs.readFileSync(filePath);
    const pdfDoc = await PDFDocument.load(bytes);
    const page = pdfDoc.getPage(pageIndex);

    // Cover old text with a white rectangle first
    page.drawRectangle({ x, y, width, height, color: rgb(1, 1, 1) });

    const text = newText || '';
    if (text.trim().length > 0) {
      const font = await pickFont(pdfDoc, text, !!bold);
      const c = color || { r: 0, g: 0, b: 0 };
      const size = fontSize || 14;

      const drawOpts = {
        x, y: y + 2,
        size,
        font,
        color: rgb(c.r, c.g, c.b),
      };
      // Synthetic italic: pdf-lib supports skewing text (xSkew) — this slants
      // the glyphs like real italic, since Noto Sans Devanagari has no italic style.
      if (italic) drawOpts.xSkew = degrees(12);

      page.drawText(text, drawOpts);

      if (underline) {
        const textWidth = font.widthOfTextAtSize(text, size);
        page.drawLine({
          start: { x, y: y - 1 },
          end: { x: x + textWidth, y: y - 1 },
          thickness: Math.max(1, size * 0.06),
          color: rgb(c.r, c.g, c.b),
        });
      }
    }

    const updatedBytes = await pdfDoc.save();
    fs.writeFileSync(filePath, updatedBytes);
    res.json({ success: true });
  } catch (err) {
    console.error('Edit apply error:', err);
    res.status(500).json({ error: 'Edit failed', detail: err.message });
  }
});

// Picks a Latin or Devanagari (Hindi) font depending on the text's script,
// and embeds a Unicode-capable TTF via fontkit so non-Latin characters render.
// Falls back to the built-in Helvetica font if the custom font files aren't present
// (Hindi text will then fail to draw — see README to add the font files).
// Bold uses a separate -Bold.ttf file if present; Devanagari has no italic style
// in Noto Sans, so italic only visually applies to English/Latin text.
const DEVANAGARI_RANGE = /[\u0900-\u097F]/;
async function pickFont(pdfDoc, text, bold) {
  const isHindi = DEVANAGARI_RANGE.test(text);
  const base = isHindi ? 'NotoSansDevanagari' : 'NotoSans';
  const weight = bold ? '-Bold' : '-Regular';
  const fontPath = path.join(__dirname, 'fonts', base + weight + '.ttf');
  const fallbackPath = path.join(__dirname, 'fonts', base + '-Regular.ttf');

  const finalPath = fs.existsSync(fontPath) ? fontPath
    : (fs.existsSync(fallbackPath) ? fallbackPath : null);

  if (finalPath) {
    pdfDoc.registerFontkit(fontkit);
    const fontBytes = fs.readFileSync(finalPath);
    return await pdfDoc.embedFont(fontBytes);
  }
  // Fallback — only works for plain English/Latin text
  return await pdfDoc.embedFont(bold ? StandardFonts.HelveticaBold : StandardFonts.Helvetica);
}


// ---------- PDF Editor: download final edited file ----------
app.get('/api/edit/:docId/download', (req, res) => {
  const filePath = editSessions[req.params.docId];
  if (!filePath || !fs.existsSync(filePath)) {
    return res.status(404).json({ error: 'Document not found or session expired' });
  }
  res.download(filePath, 'edited.pdf');
});

// Cleanup old edit sessions every hour (files older than 2 hours)
setInterval(() => {
  const now = Date.now();
  for (const [docId, filePath] of Object.entries(editSessions)) {
    try {
      const stat = fs.statSync(filePath);
      if (now - stat.mtimeMs > 2 * 60 * 60 * 1000) {
        fs.unlinkSync(filePath);
        delete editSessions[docId];
      }
    } catch (e) {
      delete editSessions[docId];
    }
  }
}, 60 * 60 * 1000);

app.listen(PORT, () => {
  console.log('ConvertKaro backend running on port ' + PORT);
});
