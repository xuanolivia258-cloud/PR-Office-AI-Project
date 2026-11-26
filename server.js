const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');
const { Pool } = require('pg');

const root = __dirname;
const port = Number(process.env.PORT || 8090);
const initialDataFile = path.join(root, 'initial-data.json');
const serveStatic = process.env.SERVE_STATIC !== 'false';
if (!process.env.DATABASE_URL) throw new Error('DATABASE_URL is required.');
const pool = new Pool({ connectionString: process.env.DATABASE_URL, max: 15 });
const publicFiles = new Map([
  ['/', ['index.html', 'text/html; charset=utf-8']], ['/index.html', ['index.html', 'text/html; charset=utf-8']],
  ['/script.js', ['script.js', 'text/javascript; charset=utf-8']], ['/styles.css', ['styles.css', 'text/css; charset=utf-8']],
  ['/tracker.css', ['tracker.css', 'text/css; charset=utf-8']]
]);

function normalizeState(value) {
  return {
    records: Array.isArray(value?.records) ? value.records : [],
    remarks: value?.remarks && typeof value.remarks === 'object' ? value.remarks : {},
    dismissedHighlights: Array.isArray(value?.dismissedHighlights) ? value.dismissedHighlights : []
  };
}

async function initializeDatabase() {
  await pool.query(`CREATE TABLE IF NOT EXISTS app_state (
    id SMALLINT PRIMARY KEY CHECK (id = 1), records JSONB NOT NULL DEFAULT '[]'::jsonb,
    remarks JSONB NOT NULL DEFAULT '{}'::jsonb, dismissed_highlights JSONB NOT NULL DEFAULT '[]'::jsonb,
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW())`);
  const initial = normalizeState(JSON.parse(fs.readFileSync(initialDataFile, 'utf8')));
  await pool.query(`INSERT INTO app_state (id, records, remarks, dismissed_highlights)
    VALUES (1, $1::jsonb, $2::jsonb, $3::jsonb) ON CONFLICT (id) DO NOTHING`,
    [JSON.stringify(initial.records), JSON.stringify(initial.remarks), JSON.stringify(initial.dismissedHighlights)]);
}

async function getState(client = pool) {
  const { rows } = await client.query('SELECT records, remarks, dismissed_highlights FROM app_state WHERE id = 1');
  const row = rows[0];
  return normalizeState({ records: row.records, remarks: row.remarks, dismissedHighlights: row.dismissed_highlights });
}

function sendJson(response, status, value) {
  response.writeHead(status, { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' });
  response.end(JSON.stringify(value));
}

function readBody(request) {
  return new Promise((resolve, reject) => {
    let body = '';
    request.on('data', chunk => { body += chunk; if (body.length > 25_000_000) reject(new Error('Request body is too large.')); });
    request.on('end', () => { try { resolve(body ? JSON.parse(body) : {}); } catch { reject(new Error('Invalid JSON body.')); } });
    request.on('error', reject);
  });
}

async function withTransaction(work) {
  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    await client.query('SELECT id FROM app_state WHERE id = 1 FOR UPDATE');
    const result = await work(client);
    await client.query('COMMIT');
    return result;
  } catch (error) {
    await client.query('ROLLBACK'); throw error;
  } finally { client.release(); }
}

async function handleApi(request, response, pathname) {
  if (request.method === 'GET' && pathname === '/api/health') {
    await pool.query('SELECT 1'); return sendJson(response, 200, { status: 'ok', database: 'connected' });
  }
  if (request.method === 'GET' && pathname === '/api/state') return sendJson(response, 200, await getState());

  if (request.method === 'POST' && pathname === '/api/import') {
    const body = await readBody(request);
    if (!Array.isArray(body.records)) return sendJson(response, 400, { error: 'records must be an array.' });
    const state = await withTransaction(async client => {
      await client.query(`UPDATE app_state SET records = $1::jsonb, dismissed_highlights = '[]'::jsonb,
        updated_at = NOW() WHERE id = 1`, [JSON.stringify(body.records)]);
      return getState(client);
    });
    return sendJson(response, 200, state);
  }

  if (request.method === 'PUT' && pathname === '/api/remark') {
    const body = await readBody(request);
    if (typeof body.key !== 'string' || typeof body.value !== 'string' || body.key.length > 300 || body.value.length > 5000)
      return sendJson(response, 400, { error: 'Invalid remark.' });
    await withTransaction(client => client.query(`UPDATE app_state SET
      remarks = jsonb_set(remarks, ARRAY[$1::text], to_jsonb($2::text), true), updated_at = NOW() WHERE id = 1`,
      [body.key, body.value]));
    return sendJson(response, 200, { saved: true });
  }

  if (request.method === 'PUT' && pathname === '/api/highlight') {
    const body = await readBody(request);
    if (typeof body.key !== 'string' || body.key.length > 500 || typeof body.dismissed !== 'boolean')
      return sendJson(response, 400, { error: 'Invalid highlight state.' });
    await withTransaction(async client => {
      const state = await getState(client); const dismissed = new Set(state.dismissedHighlights);
      if (body.dismissed) dismissed.add(body.key); else dismissed.delete(body.key);
      await client.query('UPDATE app_state SET dismissed_highlights = $1::jsonb, updated_at = NOW() WHERE id = 1',
        [JSON.stringify([...dismissed])]);
    });
    return sendJson(response, 200, { saved: true });
  }
  return sendJson(response, 404, { error: 'Not found.' });
}

const server = http.createServer(async (request, response) => {
  try {
    const pathname = new URL(request.url, `http://${request.headers.host || 'localhost'}`).pathname;
    if (pathname.startsWith('/api/')) return await handleApi(request, response, pathname);
    if (!serveStatic) return sendJson(response, 404, { error: 'Not found.' });
    const publicFile = publicFiles.get(pathname); if (!publicFile) return sendJson(response, 404, { error: 'Not found.' });
    const [file, contentType] = publicFile; response.writeHead(200, { 'Content-Type': contentType });
    fs.createReadStream(path.join(root, file)).pipe(response);
  } catch (error) { console.error(error); sendJson(response, 500, { error: 'Internal server error.' }); }
});

initializeDatabase().then(() => server.listen(port, '0.0.0.0', () =>
  console.log(`PR Office backend listening on http://0.0.0.0:${port}`)))
  .catch(error => { console.error('Database initialization failed:', error); process.exit(1); });

async function shutdown() { server.close(async () => { await pool.end(); process.exit(0); }); }
process.on('SIGTERM', shutdown); process.on('SIGINT', shutdown);
