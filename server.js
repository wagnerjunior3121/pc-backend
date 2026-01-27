require("dotenv").config();
const express = require("express");
const mongoose = require("mongoose");
const cors = require("cors");
const { Server } = require("socket.io");
const http = require("http");
const xlsx = require("xlsx");
let uploadAvailable = false;
let upload = null;
try {
  const multer = require('multer');
  upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024 } }); // 10MB limit
  uploadAvailable = true;
  console.log('✅ multer carregado — endpoint de upload habilitado');
} catch (e) {
  console.warn('⚠️ multer não encontrado — endpoint de upload estará desabilitado. Para habilitar execute: npm install multer');
}

const app = express();
const PORT = process.env.PORT || 5000;
const server = http.createServer(app);

const io = new Server(server, {
  cors: {
    origin: "*",
    methods: ["GET", "POST", "DELETE", "PUT"]
  }
});

app.use(cors({ origin: "*" }));
app.use(express.json({ limit: '100mb' }));

mongoose.connect(process.env.MONGO_URI, {
  useNewUrlParser: true,
  useUnifiedTopology: true,
})
.then(() => console.log("✅ MongoDB conectado"))
.catch(err => console.log("❌ Erro MongoDB:", err));

const assetSchema = new mongoose.Schema({
  name: String,
  parentId: { type: String, default: null },
  isCritical: { type: Boolean, default: false },
  isPinned: { type: Boolean, default: false },
  itemErp: { type: String, default: "" },
  equipmentFunction: { type: String, default: "" },
  quantidade: { type: Number, default: 0 },
});

const Asset = mongoose.model("Asset", assetSchema);

app.get("/", (req, res) => {
  res.send("🚀 Backend ativo e funcionando!");
});

app.get(["/assets", "/api/assets"], async (req, res) => {
  try {
    const assets = await Asset.find();
    res.json(assets);
  } catch (err) {
    res.status(500).json({ error: "Erro ao buscar ativos" });
  }
});

app.post(["/assets", "/api/assets"], async (req, res) => {
  try {
    const { name, parentId, isCritical, itemErp, equipmentFunction, quantidade } = req.body;
    const newAsset = new Asset({ name, parentId, isCritical, itemErp, equipmentFunction, quantidade });
    await newAsset.save();

    io.emit("asset-updated");
    res.status(201).json(newAsset);
  } catch (err) {
    res.status(500).json({ error: "Erro ao criar ativo" });
  }
});

app.put(["/assets/:id", "/api/assets/:id"], async (req, res) => {
  try {
    const { id } = req.params;
    const { name, parentId, isCritical, isPinned, itemErp, equipmentFunction, quantidade } = req.body;

    const updated = await Asset.findByIdAndUpdate(
      id,
      { name, parentId: parentId ?? null, isCritical, isPinned, itemErp, equipmentFunction, quantidade },
      { new: true }
    );

    if (!updated) return res.status(404).json({ error: "Ativo não encontrado" });

    io.emit("asset-updated");
    res.json(updated);
  } catch (err) {
    res.status(500).json({ error: "Erro ao atualizar ativo" });
  }
});

app.delete(["/assets/:id", "/api/assets/:id"], async (req, res) => {
  try {
    const { id } = req.params;
    const deleted = await Asset.findByIdAndDelete(id);
    if (!deleted) return res.status(404).json({ error: "Ativo não encontrado" });

    io.emit("asset-updated");
    res.status(200).json({ message: "Ativo excluído com sucesso" });
  } catch (err) {
    res.status(500).json({ error: "Erro ao excluir ativo" });
  }
});

let assetsHistory = [];

app.post(["/assets/saveState", "/api/assets/saveState"], (req, res) => {
  console.log("📥 Recebido /assets/saveState ou /api/assets/saveState");
  const { assets } = req.body;

  if (!assets) return res.status(400).json({ error: "Nenhum dado recebido" });

  assetsHistory.push(JSON.stringify(assets));
  console.log("📦 Histórico salvo. Total:", assetsHistory.length);
  res.sendStatus(200);
});

app.post(["/assets/restoreState", "/api/assets/restoreState"], (req, res) => {
  const { index } = req.body;
  if (assetsHistory[index]) {
    const restoredAssets = JSON.parse(assetsHistory[index]);
    Asset.deleteMany({})
      .then(() => Asset.insertMany(restoredAssets))
      .then(() => res.json({ success: true }))
      .catch(err => res.status(500).json({ error: err.message }));
  } else {
    res.status(404).json({ error: "Estado não encontrado" });
  }
});

io.on("connection", (socket) => {
  console.log("🟢 Cliente conectado:", socket.id);
  socket.on("disconnect", () => {
    console.log("🔴 Cliente desconectado:", socket.id);
  });
});

app.post('/assets/import-quantidades', async (req, res) => {
  try {
    const defaultPath = '\\licnt\\tecnica\\10_M_Op_G\\Planejamento\\PCM\\Pecas_Criticas_Oficial\\mapa_de_estoque.xlsx';
    const filePath = req.body && req.body.path ? req.body.path : defaultPath;

    const workbook = xlsx.readFile(filePath, { cellDates: true });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    const range = xlsx.utils.decode_range(sheet['!ref']);
    const codeToQty = {};
    const parseFailures = [];

    for (let R = range.s.r; R <= range.e.r; ++R) {
      const cellA = sheet[xlsx.utils.encode_cell({ r: R, c: 0 })]; // coluna A
      const cellK = sheet[xlsx.utils.encode_cell({ r: R, c: 10 })]; // coluna K (0-based)
      if (!cellA) continue;
      const raw = (cellA.v || '').toString().trim();
      if (!raw) continue;

      const digits = raw.replace(/\D/g, '');
      const noLeadingZeros = digits.replace(/^0+/, '') || digits;
      const keys = new Set([raw, raw.toLowerCase(), digits, noLeadingZeros]);

      let qty = 0;
      if (cellK && (cellK.v !== undefined && cellK.v !== null)) {
        let rawQty = cellK.v;
        if (typeof rawQty === 'string') {
          rawQty = rawQty.trim();
          rawQty = rawQty.replace(/\u00A0/g, '');
          if (rawQty.includes(',') && rawQty.includes('.')) {
            rawQty = rawQty.replace(/\./g, '').replace(/,/g, '.');
          } else {
            rawQty = rawQty.replace(/,/g, '.');
          }
          rawQty = rawQty.replace(/[^\d\.\-]/g, '');
        }
        const n = Number(rawQty);
        if (isNaN(n)) {
          parseFailures.push({ row: R + 1, code: raw, rawValue: cellK.v });
          qty = 0;
        } else {
          qty = n;
        }
      }

      for (const k of keys) {
        if (k) codeToQty[k] = qty;
      }
    }

    const assets = await Asset.find();
    const unmatched = [];
    const bulkOps = [];

    for (const a of assets) {
      const candidates = [];
      if (a.itemErp) {
        const s = a.itemErp.toString().trim();
        candidates.push(s, s.toLowerCase(), s.replace(/\D/g, ''), s.replace(/\D/g, '').replace(/^0+/, '') || s.replace(/\D/g, ''));
      }
      if (a.name) {
        const s = a.name.toString().trim();
        candidates.push(s, s.toLowerCase(), s.replace(/\s+/g, ' ').toLowerCase());
      }
      const seen = new Set();
      let foundKey = null;
      for (const c of candidates) {
        if (!c) continue;
        if (seen.has(c)) continue;
        seen.add(c);
        if (codeToQty.hasOwnProperty(c)) {
          foundKey = c;
          break;
        }
      }
      if (foundKey !== null) {
        bulkOps.push({
          updateOne: {
            filter: { _id: a._id },
            update: { $set: { quantidade: codeToQty[foundKey] } }
          }
        });
      } else {
        if (unmatched.length < 20) unmatched.push({ id: a._id, name: a.name, itemErp: a.itemErp });
      }
    }

    console.log(`📥 Importado ${Object.keys(codeToQty).length} chaves da planilha. Assets: ${assets.length}, atualizados: ${bulkOps.length}, não encontrados (ex.: ${unmatched.length}):`, unmatched.slice(0, 10));
    if (parseFailures.length > 0) {
      console.log(`⚠️ ${parseFailures.length} linhas tiveram erro ao parsear quantidade (ex.):`, parseFailures.slice(0, 10));
    }

    if (bulkOps.length > 0) {
      await Asset.bulkWrite(bulkOps);
    }

    io.emit('asset-updated');
    res.json({ success: true, updated: bulkOps.length, parsedKeys: Object.keys(codeToQty).length, unmatchedSample: unmatched.slice(0, 10), parseFailures: parseFailures.slice(0, 10) });
  } catch (err) {
    console.error('Erro ao importar quantidades:', err.message || err);
    res.status(500).json({ error: 'Erro ao importar quantidades', details: err.message });
  }
});

if (uploadAvailable) {
  app.post('/assets/import-quantidades/upload', upload.single('file'), async (req, res) => {
    try {
      if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

      const original = req.file.originalname || '';
      if (!original.toLowerCase().endsWith('.xlsx') && !original.toLowerCase().endsWith('.xls')) {
        return res.status(400).json({ error: 'Invalid file type. Please upload .xlsx or .xls' });
      }

      const buffer = req.file.buffer;
      const workbook = xlsx.read(buffer, { type: 'buffer', cellDates: true });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];

      const range = xlsx.utils.decode_range(sheet['!ref']);
      const codeToQty = {};
      const parseFailures = [];

      for (let R = range.s.r; R <= range.e.r; ++R) {
        const cellA = sheet[xlsx.utils.encode_cell({ r: R, c: 0 })]; // coluna A
        const cellK = sheet[xlsx.utils.encode_cell({ r: R, c: 10 })]; // coluna K (0-based)
        if (!cellA) continue;
        const raw = (cellA.v || '').toString().trim();
        if (!raw) continue;

        const digits = raw.replace(/\D/g, '');
        const noLeadingZeros = digits.replace(/^0+/, '') || digits;
        const keys = new Set([raw, raw.toLowerCase(), digits, noLeadingZeros]);

        let qty = 0;
        if (cellK && (cellK.v !== undefined && cellK.v !== null)) {
          let rawQty = cellK.v;
          if (typeof rawQty === 'string') {
            rawQty = rawQty.trim();
            rawQty = rawQty.replace(/\u00A0/g, '');
            if (rawQty.includes(',') && rawQty.includes('.')) {
              rawQty = rawQty.replace(/\./g, '').replace(/,/g, '.');
            } else {
              rawQty = rawQty.replace(/,/g, '.');
            }
            rawQty = rawQty.replace(/[^\d\.\-]/g, '');
          }
          const n = Number(rawQty);
          if (isNaN(n)) {
            parseFailures.push({ row: R + 1, code: raw, rawValue: cellK.v });
            qty = 0;
          } else {
            qty = n;
          }
        }

        for (const k of keys) {
          if (k) codeToQty[k] = qty;
        }
      }

      const assets = await Asset.find();
      const unmatched = [];
      const bulkOps = [];

      for (const a of assets) {
        const candidates = [];
        if (a.itemErp) {
          const s = a.itemErp.toString().trim();
          candidates.push(s, s.toLowerCase(), s.replace(/\D/g, ''), s.replace(/\D/g, '').replace(/^0+/, '') || s.replace(/\D/g, ''));
        }
        if (a.name) {
          const s = a.name.toString().trim();
          candidates.push(s, s.toLowerCase(), s.replace(/\s+/g, ' ').toLowerCase());
        }
        const seen = new Set();
        let foundKey = null;
        for (const c of candidates) {
          if (!c) continue;
          if (seen.has(c)) continue;
          seen.add(c);
          if (codeToQty.hasOwnProperty(c)) {
            foundKey = c;
            break;
          }
        }
        if (foundKey !== null) {
          bulkOps.push({
            updateOne: {
              filter: { _id: a._id },
              update: { $set: { quantidade: codeToQty[foundKey] } }
            }
          });
        } else {
          if (unmatched.length < 20) unmatched.push({ id: a._id, name: a.name, itemErp: a.itemErp });
        }
      }

      console.log(`📥 (UPLOAD) Importado ${Object.keys(codeToQty).length} chaves da planilha. Assets: ${assets.length}, atualizados: ${bulkOps.length}, não encontrados (ex.: ${unmatched.length}):`, unmatched.slice(0, 10));
      if (parseFailures.length > 0) {
        console.log(`⚠️ (UPLOAD) ${parseFailures.length} linhas tiveram erro ao parsear quantidade (ex.):`, parseFailures.slice(0, 10));
      }

      if (bulkOps.length > 0) {
        await Asset.bulkWrite(bulkOps);
      }   

      io.emit('asset-updated');
      res.json({ success: true, updated: bulkOps.length, parsedKeys: Object.keys(codeToQty).length, unmatchedSample: unmatched.slice(0, 10), parseFailures: parseFailures.slice(0, 10) });
    } catch (err) {
      console.error('Erro ao importar quantidades (upload):', err.message || err);
      res.status(500).json({ error: 'Erro ao importar quantidades (upload)', details: err.message });
    }
  });
  
  const uploadedSheetSchema = new mongoose.Schema({
    name: String,
    type: { type: String, default: 'generic' },
    parsedRows: { type: Array, default: [] },
    uploadedAt: { type: Date, default: Date.now },
    expireAt: { type: Date, default: () => new Date(Date.now() + 24 * 60 * 60 * 1000), index: { expires: '24h' } },
  });
  const UploadedSheet = mongoose.model('UploadedSheet', uploadedSheetSchema);

  app.post('/sheets/upload', upload.single('file'), async (req, res) => {
    try {
      if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
      const t = (req.query.type && String(req.query.type)) || 'generic';
      const original = req.file.originalname || 'upload.xlsx';
      const buffer = req.file.buffer;
      const workbook = xlsx.read(buffer, { type: 'buffer', cellDates: true });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

      const doc = new UploadedSheet({ name: original, type: t, parsedRows: jsonData });
      await doc.save();

      res.json({ success: true, id: doc._id, parsedRows: jsonData, uploadedAt: doc.uploadedAt });
    } catch (err) {
      console.error('Erro ao fazer upload de planilha:', err.message || err);
      res.status(500).json({ error: 'Erro ao processar upload', details: err.message });
    }
  });

  app.get('/sheets/latest', async (req, res) => {
    try {
      const t = (req.query.type && String(req.query.type)) || 'generic';
      const doc = await UploadedSheet.findOne({ type: t }).sort({ uploadedAt: -1 }).lean();
      if (!doc) return res.status(404).json({ error: 'No sheet found for type' });
      res.json({ success: true, id: doc._id, parsedRows: doc.parsedRows, uploadedAt: doc.uploadedAt });
    } catch (err) {
      console.error('Erro ao buscar sheet mais recente:', err.message || err);
      res.status(500).json({ error: 'Erro ao buscar sheet', details: err.message });
    }
  });
} else {
  app.post('/assets/import-quantidades/upload', async (req, res) => {
    res.status(501).json({ error: 'Upload endpoint not available: multer is not installed on server. Install multer and redeploy.' });
  });
}

server.listen(PORT, () => {
  console.log(`🚀 Backend rodando em http://localhost:${PORT}`);
});
