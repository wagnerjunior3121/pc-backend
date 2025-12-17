const mongoose = require("mongoose");

const AssetSchema = new mongoose.Schema({
  name: { type: String, required: true },
  parentId: { type: mongoose.Schema.Types.ObjectId, ref: "Asset", default: null },
  isCritical: { type: Boolean, default: false },
  // código/identificador usado para procurar na planilha (ERP/item code)
  itemErp: { type: String, default: "" },
  // quantidade sincronizada a partir da planilha
  quantidade: { type: Number, default: 0 },
  equipmentFunction: { type: String, default: "" },
  isPinned: { type: Boolean, default: false }
});

module.exports = mongoose.model("Asset", AssetSchema);