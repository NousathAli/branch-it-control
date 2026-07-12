import http from 'http';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const PORT = process.env.PORT || 3000;
const root = path.join(__dirname, 'dist');
const dataDir = path.join(__dirname, 'data');
const dbPath = path.join(dataDir, 'db.json');
const backupPath = path.join(dataDir, 'db.json.bak');

const types = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.ico': 'image/x-icon',
};

function sendJson(res, status, data) {
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Cache-Control': 'no-store, no-cache, must-revalidate, proxy-revalidate',
    'Pragma': 'no-cache',
    'Expires': '0',
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET,POST,OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
  });
  res.end(JSON.stringify(data));
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    let body = '';
    req.on('data', (chunk) => {
      body += chunk;
      if (body.length > 20 * 1024 * 1024) {
        req.destroy();
        reject(new Error('Request body too large'));
      }
    });
    req.on('end', () => resolve(body));
    req.on('error', reject);
  });
}

function readDatabase() {
  try {
    if (!fs.existsSync(dbPath)) return null;
    return JSON.parse(fs.readFileSync(dbPath, 'utf8'));
  } catch (error) {
    try {
      if (fs.existsSync(backupPath)) return JSON.parse(fs.readFileSync(backupPath, 'utf8'));
    } catch {}
    throw error;
  }
}

function writeDatabase(data) {
  fs.mkdirSync(dataDir, { recursive: true });
  if (fs.existsSync(dbPath)) {
    try { fs.copyFileSync(dbPath, backupPath); } catch {}
  }
  const tempPath = `${dbPath}.tmp`;
  fs.writeFileSync(tempPath, JSON.stringify(data, null, 2), 'utf8');
  fs.renameSync(tempPath, dbPath);
}

function noCacheHeaders(ext) {
  return {
    'Content-Type': types[ext] || 'application/octet-stream',
    'Cache-Control': 'no-store, no-cache, must-revalidate, proxy-revalidate',
    'Pragma': 'no-cache',
    'Expires': '0',
  };
}

const server = http.createServer(async (req, res) => {
  const urlPath = decodeURIComponent(req.url.split('?')[0]);

  if (req.method === 'OPTIONS') {
    sendJson(res, 200, { ok: true });
    return;
  }

  try {
    if (urlPath === '/api/status') {
      sendJson(res, 200, {
        ok: true,
        app: 'Branch IT Control Dashboard',
        version: 'v8.47-modal-contrast-upgrade',
        databaseExists: fs.existsSync(dbPath),
        databasePath: 'data/db.json',
      });
      return;
    }

    if (urlPath === '/api/db' && req.method === 'GET') {
      const data = readDatabase();
      sendJson(res, 200, data || { ok: false, message: 'No database file found yet.' });
      return;
    }

    if (urlPath === '/api/db' && req.method === 'POST') {
      const body = await readBody(req);
      const data = JSON.parse(body || '{}');
      if (!Array.isArray(data.branches)) {
        sendJson(res, 400, { ok: false, message: 'Invalid database payload: branches array missing.' });
        return;
      }
      writeDatabase(data);
      sendJson(res, 200, { ok: true, savedAt: new Date().toISOString(), branches: data.branches.length });
      return;
    }
  } catch (error) {
    sendJson(res, 500, { ok: false, message: error.message || 'Server error' });
    return;
  }

  let filePath = path.join(root, urlPath === '/' ? 'index.html' : urlPath);
  if (!filePath.startsWith(root)) {
    res.writeHead(403);
    res.end('Forbidden');
    return;
  }

  fs.stat(filePath, (err, stat) => {
    if (err || !stat.isFile()) filePath = path.join(root, 'index.html');
    const ext = path.extname(filePath);
    res.writeHead(200, noCacheHeaders(ext));
    fs.createReadStream(filePath).pipe(res);
  });
});

server.listen(PORT, '0.0.0.0', () => {
  console.log('Branch IT Control Dashboard v8.47 running');
  console.log(`Local:    http://localhost:${PORT}`);
  console.log(`Network:  http://<your-tailscale-ip>:${PORT}`);
  console.log(`Status:   http://localhost:${PORT}/api/status`);
  console.log('Shared data file: data/db.json');
  console.log('Press Ctrl+C to stop.');
});
