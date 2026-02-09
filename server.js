require("dotenv").config();
const express = require("express");
const mongoose = require("mongoose");
const cors = require("cors");
const { Server } = require("socket.io");
const http = require("http");
const xlsx = require("xlsx");
const cron = require("node-cron");
const nodemailer = require("nodemailer");
const PDFDocument = require("pdfkit");
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

// Model para planilhas carregadas (pendentes, completed, solicitações etc.)
const uploadedSheetSchema = new mongoose.Schema({
  name: String,
  type: { type: String, default: 'generic' },
  parsedRows: { type: Array, default: [] },
  uploadedAt: { type: Date, default: Date.now },
  expireAt: { type: Date, default: () => new Date(Date.now() + 24 * 60 * 60 * 1000), index: { expires: '24h' } },
});
const UploadedSheet = mongoose.model('UploadedSheet', uploadedSheetSchema);

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
    const parseFailures = [];

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
        let rawQty = cellK.v;
        // Normalize strings like "7,00" or "1.234,56" to a form Number() understands (e.g. "1234.56")
        if (typeof rawQty === 'string') {
          rawQty = rawQty.trim();
          // Remove non-breaking spaces
          rawQty = rawQty.replace(/\u00A0/g, '');
          // If both dot and comma present, assume dot is thousands separator and comma is decimal
          if (rawQty.includes(',') && rawQty.includes('.')) {
            rawQty = rawQty.replace(/\./g, '').replace(/,/g, '.');
          } else {
            // Replace comma with dot (handles "7,00")
            rawQty = rawQty.replace(/,/g, '.');
          }
          // Keep only digits, dot and minus
          rawQty = rawQty.replace(/[^\d\.\-]/g, '');
        }
        const n = Number(rawQty);
        if (isNaN(n)) {
          // save small sample for diagnostics
          parseFailures.push({ row: R + 1, code: raw, rawValue: cellK.v });
          qty = 0;
        } else {
          qty = n;
        }
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
  // Novo endpoint: upload XLSX para importar quantidades (multipart/form-data, campo 'file')
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
  // Upload generic sheet and persist parsed rows for 24 hours. Query param `type` optional.
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

  // Get latest uploaded sheet for a given type (if any)
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

// Inicialização do servidor
server.listen(PORT, () => {
  console.log(`🚀 Backend rodando em http://localhost:${PORT}`);
});

// Agendamento semanal para envio automático de resumo do relatório Processa Plano
// Horário padrão: toda segunda-feira às 08:00 (horário do servidor)
const REPORT_RECIPIENT = 'pcm.lic@jbs.com.br';

const mailTransporter = nodemailer.createTransport({
  host: process.env.SMTP_HOST,
  port: Number(process.env.SMTP_PORT) || 587,
  secure: process.env.SMTP_SECURE === 'true',
  auth: process.env.SMTP_USER && process.env.SMTP_PASS
    ? {
        user: process.env.SMTP_USER,
        pass: process.env.SMTP_PASS,
      }
    : undefined,
});

async function buildWeeklyReportEmailHtml() {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1;
  const monthStr = String(month).padStart(2, '0');

  // Pegamos apenas se existir pelo menos uma planilha de pendentes
  const pendentesDoc = await UploadedSheet.findOne({ type: 'pendentes' }).sort({ uploadedAt: -1 }).lean();
  if (!pendentesDoc) {
    return null;
  }

  const baseInfo = `<p><strong>Mês de referência:</strong> ${monthStr}/${year}</p>`;

  // Por enquanto, enviamos um resumo textual simples e um lembrete para abrir o sistema e gerar o PDF detalhado.
  // Se quisermos no futuro, podemos replicar toda a lógica de indicadores aqui para bater 100% com o dashboard.
  const html = `
    <div style="font-family: system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; font-size: 14px; color: #111827;">
      <h2 style="font-size:18px; margin-bottom:4px;">Relatório semanal - Processa Plano</h2>
      ${baseInfo}
      <p>Este é um envio automático gerado pelo backend do sistema de PCM.</p>
      <p>
        Para visualizar o relatório consolidado completo (com todas as guias: Preventiva Nível 1, Preventiva Nível 2,
        Lubrificação, Preditivas e Solicitações), acesse o módulo <strong>Processa Plano</strong> no sistema e use o botão
        <strong>"Gerar relatório em PDF"</strong> na barra superior.
      </p>
      <p>
        Caso deseje que este e-mail traga também os mesmos indicadores numéricos do dashboard (pendentes, realizadas,
        meta, aderência detalhada por guia), podemos evoluir este job para recalcular os indicadores diretamente no backend.
      </p>
      <p style="margin-top:16px; font-size:12px; color:#6b7280;">
        Mensagem enviada automaticamente pelo backend em ${now.toLocaleString('pt-BR')}.
      </p>
    </div>
  `;

  return html;
}

// Gera um PDF simples com um resumo das últimas planilhas carregadas
async function buildWeeklyReportPdfBuffer() {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1;
  const monthStr = String(month).padStart(2, '0');

  const pendentesDoc = await UploadedSheet.findOne({ type: 'pendentes' }).sort({ uploadedAt: -1 }).lean();
  if (!pendentesDoc) {
    return null;
  }

  const completedDoc = await UploadedSheet.findOne({ type: 'completed' }).sort({ uploadedAt: -1 }).lean();
  const solicitacoesDoc = await UploadedSheet.findOne({ type: 'solicitacoes' }).sort({ uploadedAt: -1 }).lean();

  const safeLen = (rows) => Array.isArray(rows) ? Math.max(0, rows.length - 1) : 0; // assume primeira linha como cabeçalho

  const totalPendentes = safeLen(pendentesDoc.parsedRows);
  const totalCompleted = completedDoc ? safeLen(completedDoc.parsedRows) : 0;
  const totalSolicitacoes = solicitacoesDoc ? safeLen(solicitacoesDoc.parsedRows) : 0;

  return new Promise((resolve, reject) => {
    try {
      const doc = new PDFDocument({ size: 'A4', margin: 50 });
      const chunks = [];

      doc.on('data', (chunk) => chunks.push(chunk));
      doc.on('end', () => resolve(Buffer.concat(chunks)));
      doc.on('error', (err) => reject(err));

      doc.fontSize(18).text('Relatório semanal - Processa Plano', { align: 'center' });
      doc.moveDown(0.5);
      doc.fontSize(12).text(`Mês de referência: ${monthStr}/${year}`, { align: 'center' });
      doc.moveDown(1.5);

      doc.fontSize(11).text('Resumo das planilhas mais recentes carregadas no sistema:', { align: 'left' });
      doc.moveDown(0.8);

      doc.fontSize(11).text(`• Pendentes: ${totalPendentes} linhas (último upload em ${pendentesDoc.uploadedAt ? new Date(pendentesDoc.uploadedAt).toLocaleString('pt-BR') : 'N/D'})`);

      if (completedDoc) {
        doc.text(`• Concluídas / realizadas: ${totalCompleted} linhas (último upload em ${completedDoc.uploadedAt ? new Date(completedDoc.uploadedAt).toLocaleString('pt-BR') : 'N/D'})`);
      }

      if (solicitacoesDoc) {
        doc.text(`• Solicitações: ${totalSolicitacoes} linhas (último upload em ${solicitacoesDoc.uploadedAt ? new Date(solicitacoesDoc.uploadedAt).toLocaleString('pt-BR') : 'N/D'})`);
      }

      doc.moveDown(1.2);
      doc.text('Para o detalhamento completo por tipo (Preventiva N1, N2, Lubrificação, Preditivas e Solicitações), utilize o módulo Processa Plano e o botão "Gerar relatório em PDF" na interface.', {
        align: 'left',
      });

      doc.moveDown(1.2);
      doc.fontSize(9).fillColor('#555555').text(`Relatório gerado automaticamente em ${now.toLocaleString('pt-BR')}.`, { align: 'right' });

      doc.end();
    } catch (err) {
      reject(err);
    }
  });
}

// Função que executa o fluxo completo do relatório semanal (usada pelo cron e pelo endpoint de teste)
async function runWeeklyReportJob() {
  try {
    if (!process.env.SMTP_HOST || !process.env.SMTP_USER || !process.env.SMTP_PASS) {
      console.warn('[relatorio-semanal] SMTP não configurado (SMTP_HOST/SMTP_USER/SMTP_PASS ausentes). E-mail não será enviado.');
      return;
    }

    const html = await buildWeeklyReportEmailHtml();
    if (!html) {
      console.warn('[relatorio-semanal] Nenhuma planilha pendente encontrada para montar o e-mail.');
      return;
    }

    const pdfBuffer = await buildWeeklyReportPdfBuffer();
    if (!pdfBuffer) {
      console.warn('[relatorio-semanal] Não foi possível gerar o PDF do relatório semanal. E-mail será enviado sem anexo.');
    }

    const now = new Date();
    const subject = `Relatório semanal - Processa Plano (${now.toLocaleDateString('pt-BR')})`;
    const filename = `Relatorio-Semanal-Processa-Plano-${now.toISOString().slice(0, 10)}.pdf`;

    await mailTransporter.sendMail({
      from: process.env.SMTP_FROM || process.env.SMTP_USER,
      to: REPORT_RECIPIENT,
      subject,
      html,
      attachments: pdfBuffer
        ? [
            {
              filename,
              content: pdfBuffer,
            },
          ]
        : undefined,
    });

    console.log(`[relatorio-semanal] E-mail enviado para ${REPORT_RECIPIENT} em`, now.toISOString());
  } catch (err) {
    console.error('[relatorio-semanal] Falha ao enviar e-mail semanal:', err && err.message ? err.message : err);
  }
}

// Endpoint de teste para disparar o relatório semanal manualmente
// Aceita tanto POST (para Postman/curl) quanto GET (para testar via navegador)
app.all('/debug/send-weekly-report', async (req, res) => {
  try {
    await runWeeklyReportJob();
    res.json({ success: true });
  } catch (err) {
    console.error('[relatorio-semanal] Erro ao executar job via endpoint de debug:', err && err.message ? err.message : err);
    res.status(500).json({ error: 'Erro ao executar job de relatório semanal' });
  }
});

// Cron: minuto hora dia-mês mês dia-semana => 0 8 * * 1  (segunda às 08:00)
cron.schedule('0 8 * * 1', async () => {
  await runWeeklyReportJob();
});
