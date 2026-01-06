const express = require("express");
const mongoose = require("mongoose");
const cors = require("cors");
const { Server } = require("socket.io");
const http = require("http");
const xlsx = require("xlsx");

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

// Schema do ativo
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

// Rota de teste
app.get("/", (req, res) => {
  res.send("🚀 Backend ativo e funcionando!");
});

// Buscar todos os ativos
app.get(["/assets", "/api/assets"], async (req, res) => {
  try {
    const assets = await Asset.find();
    res.json(assets);
  } catch (err) {
    res.status(500).json({ error: "Erro ao buscar ativos" });
  }
});

// Criar novo ativo
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

// Atualizar ativo
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

// Excluir ativo
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

// Histórico de estados
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

// WebSocket
io.on("connection", (socket) => {
  console.log("🟢 Cliente conectado:", socket.id);
  socket.on("disconnect", () => {
    console.log("🔴 Cliente desconectado:", socket.id);
  });
});

// Endpoint para importar quantidades da planilha
// Recebe JSON opcional: { path: "\\licnt\\tecnica\\...\\mapa_de_estoque.xlsx" }
app.post('/assets/import-quantidades', async (req, res) => {
  try {
    const defaultPath = '\\licnt\\tecnica\\10_M_Op_G\\Planejamento\\PCM\\Pecas_Criticas_Oficial\\mapa_de_estoque.xlsx';
    const filePath = req.body && req.body.path ? req.body.path : defaultPath;

    // Abre planilha
    const workbook = xlsx.readFile(filePath, { cellDates: true });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Converte em JSON por linhas (A -> coluna 1, K -> coluna 11)
    const range = xlsx.utils.decode_range(sheet['!ref']);
    const codeToQty = {};

    for (let R = range.s.r; R <= range.e.r; ++R) {
      const cellA = sheet[xlsx.utils.encode_cell({ r: R, c: 0 })]; // coluna A
      const cellK = sheet[xlsx.utils.encode_cell({ r: R, c: 10 })]; // coluna K (0-based)
      if (!cellA) continue;
      const raw = (cellA.v || '').toString().trim();
      if (!raw) continue;

      // Normaliza possíveis variações do código (texto, números, sem zeros à esquerda)
      const digits = raw.replace(/\D/g, '');
      const noLeadingZeros = digits.replace(/^0+/, '') || digits;
      const keys = new Set([raw, raw.toLowerCase(), digits, noLeadingZeros]);

      let qty = 0;
      if (cellK && (cellK.v !== undefined && cellK.v !== null)) {
        // tenta converter para número
        const n = Number(cellK.v);
        qty = isNaN(n) ? 0 : n;
      }

      // Armazena a mesma quantidade em várias chaves normalizadas para facilitar o lookup
      for (const k of keys) {
        if (k) codeToQty[k] = qty;
      }
    }

    // Buscar assets e atualizar quantidade
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
      // dedupe candidates
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

    if (bulkOps.length > 0) {
      await Asset.bulkWrite(bulkOps);
    }

    io.emit('asset-updated');
    res.json({ success: true, updated: bulkOps.length, parsedKeys: Object.keys(codeToQty).length, unmatchedSample: unmatched.slice(0, 10) });
  } catch (err) {
    console.error('Erro ao importar quantidades:', err.message || err);
    res.status(500).json({ error: 'Erro ao importar quantidades', details: err.message });
  }
});

// Inicialização do servidor
server.listen(PORT, () => {
  console.log(`🚀 Backend rodando em http://localhost:${PORT}`);
});
