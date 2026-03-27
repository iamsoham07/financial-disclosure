require('dotenv').config();
const express    = require('express');
const cors       = require('cors');
const multer     = require('multer');
const path       = require('path');
const { pool, setupDatabase } = require('./db');
const { extractData, fillAssistedTemplate, fillNegotiationTemplate } = require('./xlsxFiller');

const app    = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

app.use(cors());
app.use(express.json({ limit: '10mb' }));


// ─── LAZY DB INIT (safe for serverless cold starts) ─────────────────────────
let _dbReady = null;
const ensureDb = () => { if (!_dbReady) _dbReady = setupDatabase(); return _dbReady; };
app.use((req, res, next) => ensureDb().then(next).catch(err => res.status(500).json({ error: 'DB init failed: ' + err.message })));

// ─── HEALTH CHECK ────────────────────────────────────────────────────────────
app.get('/api/health', (req, res) => res.json({ status: 'ok' }));

// ─── TEMPLATE ENDPOINTS ──────────────────────────────────────────────────────

/** GET /api/templates — list all stored templates */
app.get('/api/templates', async (req, res) => {
  try {
    const result = await pool.query(
      'SELECT id, service, name, file_name, uploaded_at, updated_at FROM xlsx_templates ORDER BY service'
    );
    res.json(result.rows);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

/** POST /api/templates — upload/replace a template */
app.post('/api/templates', upload.single('file'), async (req, res) => {
  try {
    const { service, name } = req.body;
    if (!service || !['Assisted', 'Negotiation'].includes(service)) {
      return res.status(400).json({ error: 'service must be "Assisted" or "Negotiation"' });
    }
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

    await pool.query(`
      INSERT INTO xlsx_templates (service, name, file_data, file_name, updated_at)
      VALUES ($1, $2, $3, $4, NOW())
      ON CONFLICT (service) DO UPDATE SET
        name = EXCLUDED.name,
        file_data = EXCLUDED.file_data,
        file_name = EXCLUDED.file_name,
        updated_at = NOW()
    `, [service, name || req.file.originalname, req.file.buffer, req.file.originalname]);

    res.json({ success: true, message: `Template for "${service}" saved` });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

/** DELETE /api/templates/:service — remove a template */
app.delete('/api/templates/:service', async (req, res) => {
  try {
    const { service } = req.params;
    await pool.query('DELETE FROM xlsx_templates WHERE service = $1', [service]);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

/** GET /api/templates/:service/download — download raw template */
app.get('/api/templates/:service/download', async (req, res) => {
  try {
    const { service } = req.params;
    const result = await pool.query(
      'SELECT file_data, file_name FROM xlsx_templates WHERE service = $1',
      [service]
    );
    if (!result.rows.length) return res.status(404).json({ error: 'Template not found' });
    const { file_data, file_name } = result.rows[0];
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${file_name}"`);
    res.send(file_data);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ─── MAIN PROCESS ENDPOINT ───────────────────────────────────────────────────
/**
 * POST /api/process
 * Body: { "Service": "Negotiation" | "Assisted", "hs_object_id": 358206255303, "consent_order_json": {...} }
 *
 * 1. Loads the correct template from Postgres
 * 2. Fills the template with data from consent_order_json
 * 3. Returns the filled xlsx as binary along with hs_object_id and Service headers
 */
app.post('/api/process', async (req, res) => {
  const { Service, hs_object_id, consent_order_json } = req.body;

  if (!Service || !hs_object_id || !consent_order_json) {
    return res.status(400).json({ error: 'Service, hs_object_id and consent_order_json are required' });
  }
  if (!['Assisted', 'Negotiation'].includes(Service)) {
    return res.status(400).json({ error: 'Service must be "Assisted" or "Negotiation"' });
  }

  // Log the attempt
  let logId;
  try {
    const logResult = await pool.query(
      'INSERT INTO processing_log (hs_object_id, service, status) VALUES ($1, $2, $3) RETURNING id',
      [hs_object_id, Service, 'pending']
    );
    logId = logResult.rows[0].id;
  } catch (e) { /* non-fatal */ }

  const updateLog = async (status, error = null, fileBuffer = null, fileName = null) => {
    if (!logId) return;
    try {
      await pool.query(
        'UPDATE processing_log SET status=$1, error_message=$2, generated_file=$3, generated_file_name=$4 WHERE id=$5',
        [status, error, fileBuffer, fileName, logId]
      );
    } catch (e) { /* non-fatal */ }
  };

  try {
    // ── 1. Load template from Postgres ────────────────────────────────────
    const tmplResult = await pool.query(
      'SELECT file_data FROM xlsx_templates WHERE service = $1',
      [Service]
    );
    if (!tmplResult.rows.length) {
      throw new Error(`No template found for service "${Service}". Please upload one first.`);
    }
    const templateBuffer = tmplResult.rows[0].file_data;

    // ── 2. Extract data and fill template ────────────────────────────────
    const consentOrder = typeof consent_order_json === 'string'
      ? JSON.parse(consent_order_json)
      : consent_order_json;

    const data = extractData(consentOrder);
    const filledBuffer = Service === 'Assisted'
      ? await fillAssistedTemplate(templateBuffer, data)
      : await fillNegotiationTemplate(templateBuffer, data);

    const resName = data.resName.replace(/[^A-Za-z0-9_-]/g, '_') || 'Unknown';
    const outName = `Financial_Disclosure_${Service}_${resName}_${new Date().toISOString().slice(0, 10)}.xlsx`;

    await updateLog('success', null, filledBuffer, outName);

    // ── 3. Return the filled xlsx with metadata headers ───────────────────
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${outName}"`);
    res.setHeader('X-HS-Object-ID', String(hs_object_id));
    res.setHeader('X-Service', Service);
    res.send(filledBuffer);

  } catch (err) {
    await updateLog('error', err.message);
    console.error('[/api/process]', err.message);
    res.status(500).json({ error: err.message });
  }
});

/** GET /api/logs — recent processing log */
app.get('/api/logs', async (req, res) => {
  try {
    const result = await pool.query(
      'SELECT id, hs_object_id, service, status, error_message, generated_file_name, created_at FROM processing_log ORDER BY created_at DESC LIMIT 50'
    );
    res.json(result.rows);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

/** GET /api/logs/:id/download — download the generated xlsx for a log entry */
app.get('/api/logs/:id/download', async (req, res) => {
  try {
    const { id } = req.params;
    const result = await pool.query(
      'SELECT generated_file, generated_file_name FROM processing_log WHERE id = $1',
      [id]
    );
    if (!result.rows.length) return res.status(404).json({ error: 'Log entry not found' });
    const { generated_file, generated_file_name } = result.rows[0];
    if (!generated_file) return res.status(404).json({ error: 'No file stored for this log entry' });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${generated_file_name}"`);
    res.send(generated_file);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ─── EXPORT FOR VERCEL / IMPORT ──────────────────────────────────────────────
module.exports = app;

// ─── START (only when run directly, not when imported by Vercel) ─────────────
if (require.main === module) {
  const PORT = process.env.PORT || 3001;
  setupDatabase()
    .then(() => app.listen(PORT, () => console.log(`✅ Server running on port ${PORT}`)))
    .catch(err => { console.error('DB setup failed:', err); process.exit(1); });
}
