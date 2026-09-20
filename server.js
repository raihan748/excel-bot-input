const http = require('http');
const fs = require('fs');
const path = require('path');
const url = require('url');

const PORT = process.env.PORT || 3000;
const PUBLIC_DIR = path.join(__dirname, 'public');

// Handlers serverless functions
const processHandler = require('./api/process');
const exportHandler = require('./api/export-excel');
const mergeHandler = require('./api/merge-excel');

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.json': 'application/json',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon'
};

const server = http.createServer(async (req, res) => {
  const parsedUrl = url.parse(req.url, true);
  const pathname = parsedUrl.pathname;

  // Enhance res with helper methods matching Vercel Serverless Function signature
  res.status = function(code) {
    res.statusCode = code;
    return res;
  };
  res.json = function(data) {
    res.setHeader('Content-Type', 'application/json');
    res.end(JSON.stringify(data));
    return res;
  };
  res.send = function(data) {
    res.end(data);
    return res;
  };

  // Helper untuk membaca request body JSON
  async function readJsonBody() {
    return new Promise((resolve, reject) => {
      let body = '';
      req.on('data', chunk => {
        body += chunk;
        // Limit 60MB untuk upload gambar beresolusi tinggi
        if (body.length > 60 * 1024 * 1024) {
          req.destroy();
          reject(new Error('Ukuran request terlalu besar (maks 60MB).'));
        }
      });
      req.on('end', () => {
        if (!body) return resolve({});
        try {
          resolve(JSON.parse(body));
        } catch (e) {
          reject(new Error('Invalid JSON body'));
        }
      });
      req.on('error', reject);
    });
  }

  // 1. Routing API Endpoints
  if (pathname.startsWith('/api/')) {
    try {
      req.body = await readJsonBody();
    } catch (e) {
      return res.status(400).json({ error: e.message });
    }

    if (pathname === '/api/process') {
      return processHandler(req, res);
    } else if (pathname === '/api/export-excel') {
      return exportHandler(req, res);
    } else if (pathname === '/api/merge-excel') {
      return mergeHandler(req, res);
    } else {
      return res.status(404).json({ error: 'Endpoint API tidak ditemukan.' });
    }
  }

  // 2. Routing File Statis (public)
  let filePath = path.join(PUBLIC_DIR, pathname === '/' ? 'index.html' : pathname);

  fs.stat(filePath, (err, stats) => {
    if (err || !stats.isFile()) {
      filePath = path.join(PUBLIC_DIR, 'index.html');
    }

    const ext = path.extname(filePath).toLowerCase();
    const contentType = MIME_TYPES[ext] || 'application/octet-stream';

    fs.readFile(filePath, (readErr, content) => {
      if (readErr) {
        res.status(500).send('Internal Server Error');
        return;
      }
      res.writeHead(200, { 'Content-Type': contentType });
      res.end(content);
    });
  });
});

server.listen(PORT, () => {
  console.log('=====================================================');
  console.log(`🚀 ExcelBot AI Server berjalan di: http://localhost:${PORT}`);
  console.log('Siap digunakan lokal atau dideploy langsung ke Vercel!');
  console.log('=====================================================');
});
