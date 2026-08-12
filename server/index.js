// ConvertKaro Backend Server
// Handles: PDF -> Word, Word -> PDF, and PDF editing (via pdf-lib)

const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const os = require('os');
const cors = require('cors');
const libre = require('libreoffice-convert');
const { PDFDocument, rgb, StandardFonts } = require('pdf-lib');

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

// ---------- Health check ----------
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', service: 'ConvertKaro backend' });
});

// ---------- PDF -> Word ----------
app.post('/api/convert/pdf-to-word', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

  const inputPath = req.file.path;
  const inputBuf = fs.readFileSync(inputPath);

  libre.convert(inputBuf, '.docx', undefined, (err, done) => {
    fs.unlink(inputPath, () => {}); // cleanup uploaded temp file
    if (err) {
      console.error('PDF->Word conversion error:', err);
      return res.status(500).json({ error: 'Conversion failed' });
    }
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename=converted.docx');
    res.send(done);
  });
});

// ---------- Word -> PDF ----------
app.post('/api/convert/word-to-pdf', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

  const inputPath = req.file.path;
  const inputBuf = fs.readFileSync(inputPath);

  libre.convert(inputBuf, '.pdf', undefined, (err, done) => {
    fs.unlink(inputPath, () => {});
    if (err) {
      console.error('Word->PDF conversion error:', err);
      return res.status(500).json({ error: 'Conversion failed' });
    }
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename=converted.pdf');
    res.send(done);
  });
});

// ---------- PDF Editor: upload & start session ----------
app.post('/api/edit/upload', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

  const docId = Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
  const storedPath = path.join(os.tmpdir(), `edit-${docId}.pdf`);
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
// body: { pageIndex, x, y, width, height, newText, fontSize }
app.post('/api/edit/:docId/apply', async (req, res) => {
  const filePath = editSessions[req.params.docId];
  if (!filePath || !fs.existsSync(filePath)) {
    return res.status(404).json({ error: 'Document not found or session expired' });
  }
  const { pageIndex, x, y, width, height, newText, fontSize } = req.body;

  try {
    const bytes = fs.readFileSync(filePath);
    const pdfDoc = await PDFDocument.load(bytes);
    const page = pdfDoc.getPage(pageIndex);
    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);

    // Cover old text with a white rectangle
    page.drawRectangle({
      x, y, width, height,
      color: rgb(1, 1, 1),
    });

    // Draw new text in its place
    page.drawText(newText || '', {
      x, y: y + 2,
      size: fontSize || 12,
      font,
      color: rgb(0, 0, 0),
    });

    const updatedBytes = await pdfDoc.save();
    fs.writeFileSync(filePath, updatedBytes);
    res.json({ success: true });
  } catch (err) {
    console.error('Edit apply error:', err);
    res.status(500).json({ error: 'Edit failed' });
  }
});

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
  console.log(`ConvertKaro backend running on port ${PORT}`);
});
