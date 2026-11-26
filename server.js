const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');

const root = __dirname;
const port = Number(process.env.PORT || 8090);
const dataFile = process.env.DATA_FILE || path.join(root, 'data', 'store.json');
const initialDataFile = path.join(root, 'initial-data.json');
const publicFiles = new Map([
  ['/', ['index.html', 'text/html; charset=utf-8']],
  ['/index.html', ['index.html', 'text/html; charset=utf-8']],
  ['/script.js', ['script.js', 'text/javascript; charset=utf-8']],
  ['/styles.css', ['styles.css', 'text/css; charset=utf-8']],
  ['/tracker.css', ['tracker.css', 'text/css; charset=utf-8']]
]);

function readJson(file) {
  return JSON.parse(fs.readFileSync(file, 'utf8'));
}

function normalizeState(value) {
  return {
    records: Array.isArray(value?.records) ? value.records : [],
    remarks: value?.remarks && typeof value.remarks === 'object' ? value.remarks : {},
    dismissedHighlights: Array.isArray(value?.dismissedHighlights) ? value.dismissedHighlights : []
  };
}

function writeState(state) {
  fs.mkdirSync(path.dirname(dataFile), { recursive: true });
  const temporary = `${dataFile}.tmp`;
  fs.writeFileSync(temporary, JSON.stringify(normalizeState(state), null, 2));
  fs.renameSync(temporary, dataFile);
}

function loadState() {
  if (fs.existsSync(dataFile)) return normalizeState(readJson(dataFile));
  const initial = normalizeState(readJson(initialDataFile));
  writeState(initial);
  return initial;
}

let state = loadState();

function sendJson(response, status, value) {
  response.writeHead(status, { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' });
  response.end(JSON.stringify(value));
}

function readBody(request) {
  return new Promise((resolve, reject) => {
    let body = '';
    request.on('data', chunk => {
      body += chunk;
      if (body.length > 5_000_000) reject(new Error('Request body is too large.'));
    });
    request.on('end', () => {
      try {
        resolve(body ? JSON.parse(body) : {});
      } catch {
        reject(new Error('Invalid JSON body.'));
      }
    });
    request.on('error', reject);
  });
}

async function handleApi(request, response, pathname) {
  if (request.method === 'GET' && pathname === '/api/health') return sendJson(response, 200, { status: 'ok' });
  if (request.method === 'GET' && pathname === '/api/state') return sendJson(response, 200, state);

  if (request.method === 'POST' && pathname === '/api/import') {
    const body = await readBody(request);
    if (!Array.isArray(body.records)) return sendJson(response, 400, { error: 'records must be an array.' });
    state.records = body.records;
    state.dismissedHighlights = [];
    writeState(state);
    return sendJson(response, 200, state);
  }

  if (request.method === 'PUT' && pathname === '/api/remark') {
    const body = await readBody(request);
    if (typeof body.key !== 'string' || typeof body.value !== 'string' || body.key.length > 300 || body.value.length > 5000) {
      return sendJson(response, 400, { error: 'Invalid remark.' });
    }
    state.remarks[body.key] = body.value;
    writeState(state);
    return sendJson(response, 200, { saved: true });
  }

  if (request.method === 'PUT' && pathname === '/api/highlight') {
    const body = await readBody(request);
    if (typeof body.key !== 'string' || body.key.length > 500 || typeof body.dismissed !== 'boolean') {
      return sendJson(response, 400, { error: 'Invalid highlight state.' });
    }
    const dismissed = new Set(state.dismissedHighlights);
    if (body.dismissed) dismissed.add(body.key);
    else dismissed.delete(body.key);
    state.dismissedHighlights = [...dismissed];
    writeState(state);
    return sendJson(response, 200, { saved: true });
  }

  return sendJson(response, 404, { error: 'Not found.' });
}

const server = http.createServer(async (request, response) => {
  try {
    const pathname = new URL(request.url, `http://${request.headers.host || 'localhost'}`).pathname;
    if (pathname.startsWith('/api/')) return await handleApi(request, response, pathname);
    const publicFile = publicFiles.get(pathname);
    if (!publicFile) {
      response.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      return response.end('Not found');
    }
    const [file, contentType] = publicFile;
    response.writeHead(200, { 'Content-Type': contentType });
    fs.createReadStream(path.join(root, file)).pipe(response);
  } catch (error) {
    sendJson(response, 500, { error: error.message || 'Internal server error.' });
  }
});

server.listen(port, '0.0.0.0', () => {
  console.log(`PR Office server listening on http://0.0.0.0:${port}`);
  console.log(`Persistent data: ${dataFile}`);
});
