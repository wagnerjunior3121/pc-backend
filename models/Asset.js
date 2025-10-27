const mongoose = require("mongoose");

const AssetSchema = new mongoose.Schema({
  name: { type: String, required: true },
  parentId: { type: mongoose.Schema.Types.ObjectId, ref: "Asset", default: null },
  isCritical: { type: Boolean, default: false }
});

module.exports = mongoose.model("Asset", AssetSchema);