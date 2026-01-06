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
app.use(express.json({ limit: "100mb" }));

mongoose.connect(process.env.MONGO_URI, {
  useNewUrlParser: true,
  useUnifiedTopology: true,
})
.then(() => console.log("✅ MongoDB conectado"))
.catch(err => console.log("❌ Erro MongoDB:", err));

/* =========================
   Schema
========================= */
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

/* =========================
   Utilitário
========================= */
function normalize(v) {
  return (v ?? "")
    .toString()
    .trim()
    .toUpperCase()
    .replace(/\s+/g, "");
}

/* =========================
   Rotas básicas
========================= */
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

/* =========================
   Histórico de estados
========================= */
let assetsHistory = [];

app.post(["/assets/saveState", "/api/assets/saveState"], (req, res) => {
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

/* =========================
   WebSocket
========================= */
io.on("connection", (socket) => {
  console.log("🟢 Cliente conectado:", socket.id);
  socket.on("disconnect", () => {
    console.log("🔴 Cliente desconectado:", socket.id);
  });
});

/* =========================
   IMPORTAÇÃO DE QUANTIDADES
========================= */
app.post("/assets/import-quantidades", async (req, res) => {
  try {
    const defaultPath =
      "\\\\licnt\\tecnica\\10_M_Op_G\\Planejamento\\PCM\\Pecas_Criticas_Oficial\\mapa_de_estoque.xlsx";

    const filePath = req.body?.path || defaultPath;

    console.log("📄 Lendo planilha:", filePath);

    const workbook = xlsx.readFile(filePath, { cellDates: true });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    const range = xlsx.utils.decode_range(sheet["!ref"]);
    const codeToQty = {};

    for (let R = range.s.r; R <= range.e.r; ++R) {
      const cellA = sheet[xlsx.utils.encode_cell({ r: R, c: 0 })];   // coluna A
      const cellK = sheet[xlsx.utils.encode_cell({ r: R, c: 10 })];  // coluna K

      if (!cellA) continue;

      const code = normalize(cellA.v);
      if (!code) continue;

      let qty = 0;
      if (cellK?.v !== undefined && cellK?.v !== null) {
        const n = Number(cellK.v);
        qty = isNaN(n) ? 0 : n;
      }

      codeToQty[code] = qty;
    }

    console.log("📊 Códigos lidos da planilha:", Object.keys(codeToQty).length);

    const assets = await Asset.find();
    console.log("📦 Assets no banco:", assets.length);

    const bulkOps = assets
      .map(a => {
        const lookupKey = normalize(a.itemErp || a.name);
        if (lookupKey && codeToQty.hasOwnProperty(lookupKey)) {
          return {
            updateOne: {
              filter: { _id: a._id },
              update: { $set: { quantidade: codeToQty[lookupKey] } }
            }
          };
        }
        return null;
      })
      .filter(Boolean);

    console.log("✅ Registros a atualizar:", bulkOps.length);

    if (bulkOps.length > 0) {
      await Asset.bulkWrite(bulkOps);
    }

    io.emit("asset-updated");
    res.json({ success: true, updated: bulkOps.length });

  } catch (err) {
    console.error("❌ Erro ao importar quantidades:", err);
    res.status(500).json({
      error: "Erro ao importar quantidades",
      details: err.message
    });
  }
});

/* =========================
   Inicialização
========================= */
server.listen(PORT, () => {
  console.log(`🚀 Backend rodando em http://localhost:${PORT}`);
});
