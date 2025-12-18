const express = require("express");
const mongoose = require("mongoose");
const cors = require("cors");
const { Server } = require("socket.io");
const http = require("http");
const xlsx = require("xlsx");
const fs = require("fs");

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
    console.log("📥 Iniciando importação de quantidades...");
    const defaultPath = '\\\\licnt\\tecnica\\10_M_Op_G\\Planejamento\\PCM\\Pecas_Críticas_Oficial\\mapa_de_estoque.xlsx';
    const filePath = req.body?.path || defaultPath;
    console.log("📂 Caminho do arquivo recebido:", filePath);

    // 🔎 Verifica se o arquivo existe
    if (!fs.existsSync(filePath)) {
      console.error("❌ Arquivo NÃO encontrado no caminho informado");
      return res.status(404).json({ error: "Arquivo não encontrado", path: filePath });
    }

    console.log("✅ Arquivo encontrado");

    // Abre planilha
    const workbook = xlsx.readFile(filePath, { cellDates: true });
    console.log("📘 Planilha carregada com sucesso");
    console.log("📄 Sheets disponíveis:", workbook.SheetNames);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    if (!sheet || !sheet['!ref']) {
     return res.status(400).json({ error: "Planilha vazia ou inválida" });
    }

    console.log("📄 Usando sheet:", sheetName);
    console.log("📐 Intervalo da planilha (!ref):", sheet['!ref']);

    // Converte em JSON por linhas (A -> coluna 1, K -> coluna 11)
    const range = xlsx.utils.decode_range(sheet['!ref']);
    console.log("📊 Range decodificado:", range);
    const codeToQty = {};
    let linhasLidas = 0;

    for (let R = range.s.r; R <= range.e.r; ++R) {
      const cellA = sheet[xlsx.utils.encode_cell({ r: R, c: 0 })];
      const cellK = sheet[xlsx.utils.encode_cell({ r: R, c: 10 })];
    
      if (!cellA) continue;
    
      const code = (cellA.v || '').toString().trim();
      if (!code) continue;
    
      let qty = 0;
      if (cellK && cellK.v !== undefined && cellK.v !== null) {
        const n = Number(cellK.v);
        qty = isNaN(n) ? 0 : n;
      }
    
      codeToQty[code] = qty;
      linhasLidas++;
    
      // 🔍 loga apenas as 5 primeiras linhas válidas
      if (linhasLidas <= 5) {
        console.log(`📌 Linha ${R + 1} | Código: "${code}" | Qtd: ${qty}`);
      }
    }

    console.log(`📦 Total de linhas válidas lidas da planilha: ${linhasLidas}`);
    console.log(`🔑 Total de códigos únicos encontrados: ${Object.keys(codeToQty).length}`);

    // Buscar assets e atualizar quantidade
    const assets = await Asset.find();
    console.log(`🗃️ Total de assets no banco: ${assets.length}`);
    let encontrados = 0;
    const bulkOps = assets.map(a => {
      const lookupKey = (a.itemErp || a.name || '').toString().trim();
      if (lookupKey && codeToQty.hasOwnProperty(lookupKey)) {
        encontrados++;
        return {
          updateOne: {
            filter: { _id: a._id },
            update: { $set: { quantidade: codeToQty[lookupKey] } }
          }
        };
      }
      return null;
    }).filter(Boolean);

    console.log(`🔄 Assets com correspondência encontrada: ${encontrados}`);
    console.log(`📝 Operações de update preparadas: ${bulkOps.length}`);

    if (bulkOps.length > 0) {
      await Asset.bulkWrite(bulkOps);
      console.log("✅ Atualização em massa realizada com sucesso");
    } else {
      console.warn("⚠️ Nenhum asset correspondeu aos códigos da planilha");
    }

    io.emit('asset-updated');
    res.json({ success: true, updated: bulkOps.length, linhasLidas, totalAssets: assets.length });
  } catch (err) {
    console.error('Erro ao importar quantidades:', err.message || err);
    res.status(500).json({ error: 'Erro ao importar quantidades', details: err.message });
  }
});

// Inicialização do servidor
server.listen(PORT, () => {
  console.log(`🚀 Backend rodando em http://localhost:${PORT}`);
});
