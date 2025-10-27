const express = require("express");
const mongoose = require("mongoose");
const cors = require("cors");
const { Server } = require("socket.io");
const http = require("http");

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
app.use(express.json());

mongoose.connect(
  "mongodb+srv://wagnerjuniorsql:Daledoly12@cluster0.oc1at.mongodb.net/pcm?retryWrites=true&w=majority",
  { useNewUrlParser: true, useUnifiedTopology: true }
)
.then(() => console.log("✅ MongoDB conectado"))
.catch(err => console.log("❌ Erro MongoDB:", err));

const assetSchema = new mongoose.Schema({
  name: String,
  parentId: { type: String, default: null },
  isCritical: { type: Boolean, default: false },
  isPinned: { type: Boolean, default: false },
  itemErp: { type: String, default: "" },
  equipmentFunction: { type: String, default: "" },
});

const Asset = mongoose.model("Asset", assetSchema);

app.get("/assets", async (req, res) => {
  try {
    const assets = await Asset.find();
    res.json(assets);
  } catch (err) {
    res.status(500).json({ error: "Erro ao buscar ativos" });
  }
});

app.put("/assets/:id", async (req, res) => {
  try {
    const { id } = req.params;
    const { name, parentId, isCritical, isPinned, itemErp, equipmentFunction } = req.body;

    const updated = await Asset.findByIdAndUpdate(
      id,
      { 
        name,
        parentId: parentId !== undefined ? parentId : null,
        isCritical,
        isPinned, 
        itemErp,
        equipmentFunction
      },
      { new: true }
    );

    if (!updated) return res.status(404).json({ error: "Ativo não encontrado" });

    io.emit("asset-updated"); 
    res.json(updated);
  } catch (err) {
    res.status(500).json({ error: "Erro ao atualizar ativo" });
  }
});

let assetsHistory = [];

app.post(["/assets/saveState", "/api/assets/saveState"], (req, res) => {
  console.log("📥 Recebido /assets/saveState ou /api/assets/saveState");
  const { assets } = req.body;

  if (!assets) {
    return res.status(400).json({ error: "Nenhum dado recebido" });
  }

  assetsHistory.push(JSON.stringify(assets));
  console.log("📦 Histórico salvo. Total:", assetsHistory.length);
  res.sendStatus(200);
});

app.post("/api/assets/saveState", (req, res) => {
  console.log("Recebido:", req.body); 
  const { assets } = req.body;
  assetsHistory.push(JSON.stringify(assets));
  res.sendStatus(200);
});

app.post("/api/assets/restoreState", (req, res) => {
  const { index } = req.body;
  if (assetsHistory[index]) {
    const restoredAssets = JSON.parse(assetsHistory[index]);
    Asset.deleteMany({})
      .then(() => Asset.insertMany(restoredAssets))
      .then(() => res.json({ success: true }))
      .catch((err) => res.status(500).json({ error: err.message }));
  } else {
    res.status(404).json({ error: "Estado não encontrado" });
  }
});

app.post("/assets", async (req, res) => {
  try {
    const { name, parentId, isCritical, itemErp, equipmentFunction } = req.body;
    const newAsset = new Asset({ name, parentId, isCritical, itemErp, equipmentFunction });
    await newAsset.save();

    io.emit("asset-updated");
    res.status(201).json(newAsset);
  } catch (err) {
    res.status(500).json({ error: "Erro ao criar ativo" });
  }
});

app.delete("/assets/:id", async (req, res) => {
  const { id } = req.params;
  try {
    const deleted = await Asset.findByIdAndDelete(id);
    if (!deleted) return res.status(404).json({ error: "Ativo não encontrado" });

    io.emit("asset-updated"); 
    res.status(200).json({ message: "Ativo excluído com sucesso" });
  } catch (err) {
    res.status(500).json({ error: "Erro ao excluir ativo" });
  }
});

io.on("connection", (socket) => {
  console.log("🟢 Cliente conectado:", socket.id);
  socket.on("disconnect", () => {
    console.log("🔴 Cliente desconectado:", socket.id);
  });
});

server.listen(PORT, () => {
  console.log(`🚀 Backend rodando em http://localhost:${PORT}`);
});