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
    origin: "*", // 🔓 Libera o frontend
    methods: ["GET", "POST", "DELETE", "PUT"]
  }
});

app.use(cors({ origin: "*" }));
app.use(express.json());

// ✅ Conexão com o MongoDB Atlas
mongoose.connect(process.env.MONGO_URI, {
  useNewUrlParser: true,
  useUnifiedTopology: true,
})
.then(() => console.log("✅ MongoDB conectado"))
.catch(err => console.error("❌ Erro MongoDB:", err));

const assetSchema = new mongoose.Schema({
  name: String,
  parentId: { type: String, default: null },
  isCritical: { type: Boolean, default: false },
  isPinned: { type: Boolean, default: false },
  itemErp: { type: String, default: "" },
  equipmentFunction: { type: String, default: "" },
});

const Asset = mongoose.model("Asset", assetSchema);

// ✅ Rota inicial para teste
app.get("/", (req, res) => {
  res.send("✅ API do PCM Backend está online!");
});

// === ROTAS CRUD ===
app.get("/assets", async (req, res) => {
  try {
    const assets = await Asset.find();
    res.json(assets);
  } catch (err) {
    res.status(500).json({ error: "Erro ao buscar ativos" });
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

app.put("/assets/:id", async (req, res) => {
  try {
    const { id } = req.params;
    const { name, parentId, isCritical, isPinned, itemErp, equipmentFunction } = req.body;

    const updated = await Asset.findByIdAndUpdate(
      id,
      { name, parentId, isCritical, isPinned, itemErp, equipmentFunction },
      { new: true }
    );

    if (!updated) return res.status(404).json({ error: "Ativo não encontrado" });

    io.emit("asset-updated");
    res.json(updated);
  } catch (err) {
    res.status(500).json({ error: "Erro ao atualizar ativo" });
  }
});

app.delete("/assets/:id", async (req, res) => {
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

// === WebSocket ===
io.on("connection", (socket) => {
  console.log("🟢 Cliente conectado:", socket.id);
  socket.on("disconnect", () => {
    console.log("🔴 Cliente desconectado:", socket.id);
  });
});

// ✅ Escuta a porta correta (obrigatório para Render / Fly.io)
server.listen(PORT, "0.0.0.0", () => {
  console.log(`🚀 Backend rodando na porta ${PORT}`);
});
