const express = require('express');
const http = require('http');
const https = require('https');
const { WebSocketServer } = require('ws');
const path = require('path');
const url = require('url');
const fs = require('fs');

const app = express();
const server = http.createServer(app);
const wss = new WebSocketServer({ server });

const PORT = 3001;

// --- Static files ---
app.use('/public', express.static(path.join(__dirname, 'public')));

// 提供 pptx-preview 的 UMD 文件给浏览器
const pptxPreviewPath = path.join(__dirname, 'node_modules', 'pptx-preview', 'dist', 'pptx-preview.umd.js');
if (fs.existsSync(pptxPreviewPath)) {
  app.get('/public/pptx-preview.umd.js', (req, res) => {
    res.sendFile(pptxPreviewPath);
  });
} else {
  // fallback: copy es module
  const esPath = path.join(__dirname, 'node_modules', 'pptx-preview', 'dist', 'pptx-preview.es.js');
  if (fs.existsSync(esPath)) {
    app.get('/public/pptx-preview.umd.js', (req, res) => {
      res.sendFile(esPath);
    });
  }
}

// --- JSON body parser for upload ---
app.use(express.json({ limit: '100mb' }));
app.use(express.urlencoded({ limit: '100mb', extended: true }));

// 上传目录
const uploadDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadDir)) {
  fs.mkdirSync(uploadDir, { recursive: true });
}

// --- Pages ---
app.get('/view', (req, res) => {
  res.sendFile(path.join(__dirname, 'view.html'));
});

app.get('/control', (req, res) => {
  res.sendFile(path.join(__dirname, 'control.html'));
});

app.get('/', (req, res) => {
  res.redirect('/view');
});

// ===== PPT 文件下载接口 =====
// 从 URL 下载 ppt 文件并存到本地，返回文件访问路径
app.post('/api/download-ppt', async (req, res) => {
  const { url: pptUrl } = req.body;
  if (!pptUrl) {
    return res.status(400).json({ error: 'missing url' });
  }

  let targetUrl = pptUrl.trim();
  if (!/^https?:\/\//i.test(targetUrl)) {
    targetUrl = 'http://' + targetUrl;
  }

  try {
    const parsed = new URL(targetUrl);
    const fileName = parsed.pathname.split('/').pop() || 'presentation.pptx';
    // 如果文件名没有扩展名，加上 .pptx
    const safeName = fileName.includes('.') ? fileName : fileName + '.pptx';
    const destPath = path.join(uploadDir, safeName);

    // 检查是否已下载过
    if (fs.existsSync(destPath)) {
      const stat = fs.statSync(destPath);
      // 如果文件在1小时内下载过，直接返回
      if (Date.now() - stat.mtimeMs < 3600000) {
        return res.json({ file: `/uploads/${safeName}`, name: safeName });
      }
    }

    await downloadFile(targetUrl, destPath);
    console.log(`[Download] ${targetUrl} -> ${destPath}`);
    res.json({ file: `/uploads/${safeName}`, name: safeName });
  } catch (err) {
    console.error(`[Download Error] ${pptUrl}:`, err.message);
    res.status(500).json({ error: err.message });
  }
});

// PPT 文件上传接口
const multer = require('multer');
const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, uploadDir),
  filename: (req, file, cb) => {
    const safeName = Date.now() + '-' + file.originalname.replace(/[^a-zA-Z0-9._-]/g, '_');
    cb(null, safeName);
  }
});
const upload = multer({
  storage,
  limits: { fileSize: 100 * 1024 * 1024 }, // 100MB
  fileFilter: (req, file, cb) => {
    if (file.mimetype.includes('presentation') || file.originalname.match(/\.pptx?$/i)) {
      cb(null, true);
    } else {
      cb(new Error('Only .pptx files allowed'));
    }
  }
});

app.post('/api/upload-ppt', upload.single('file'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'no file uploaded' });
  }
  console.log(`[Upload] ${req.file.originalname} -> ${req.file.filename}`);
  res.json({
    file: `/uploads/${req.file.filename}`,
    name: req.file.originalname
  });
});

// 提供上传文件访问
app.use('/uploads', express.static(uploadDir));

// ===== 错误处理 =====
app.use((err, req, res, next) => {
  if (err instanceof multer.MulterError) {
    return res.status(400).json({ error: err.message });
  }
  if (err) {
    return res.status(400).json({ error: err.message });
  }
  next();
});

// --- HTTP Proxy for iframe (已有的) ---
function targetToProxyUrl(targetUrl) {
  return '/proxy/' + targetUrl;
}

function rewriteHtml(body, targetBaseUrl) {
  body = body.replace(
    /(src|href|action|data)\s*=\s*"(https?:\/\/[^"]+)"/gi,
    (match, attr, val) => `${attr}="${targetToProxyUrl(val)}"`
  );
  body = body.replace(
    /(src|href|action|data)\s*=\s*'(https?:\/\/[^']+)'/gi,
    (match, attr, val) => `${attr}='${targetToProxyUrl(val)}'`
  );
  body = body.replace(
    /srcset\s*=\s*"([^"]+)"/gi,
    (match, val) => {
      const newVal = val.split(',').map(part => {
        part = part.trim();
        const [s, desc] = part.split(/\s+/, 2);
        if (/^https?:\/\//i.test(s)) {
          return targetToProxyUrl(s) + (desc ? ' ' + desc : '');
        }
        return part;
      }).join(', ');
      return `srcset="${newVal}"`;
    }
  );
  body = body.replace(
    /url\(\s*["']?(https?:\/\/[^"')]+)["']?\s*\)/gi,
    (match, val) => `url("${targetToProxyUrl(val)}")`
  );
  return body;
}

app.use('/proxy/', (req, res) => {
  let responded = false;
  const safeRespond = (statusCode, headers, body) => {
    if (responded) return;
    responded = true;
    if (body !== undefined) {
      res.writeHead(statusCode, headers);
      res.end(body);
    } else {
      res.writeHead(statusCode, headers);
      res.end();
    }
  };

  let target = req.originalUrl.slice(7);

  if (!target) {
    safeRespond(400, { 'Content-Type': 'text/plain' }, 'Missing target URL');
    return;
  }

  if (!/^https?:\/\//i.test(target)) {
    target = 'http://' + target;
  }

  const parsed = url.parse(target);
  const isHttps = parsed.protocol === 'https:';
  const port = parsed.port || (isHttps ? 443 : 80);

  const opts = {
    hostname: parsed.hostname,
    port: port,
    path: parsed.path + (parsed.search || ''),
    method: req.method,
    headers: { ...req.headers, host: parsed.host, 'accept-encoding': 'identity' },
    rejectUnauthorized: false,
    timeout: 15000,
  };
  delete opts.headers['proxy-connection'];

  delete opts.headers['sec-fetch-site'];
  delete opts.headers['sec-fetch-mode'];
  delete opts.headers['sec-fetch-dest'];
  delete opts.headers['sec-fetch-user'];
  delete opts.headers['x-forwarded-for'];
  delete opts.headers['x-forwarded-proto'];
  delete opts.headers['x-forwarded-host'];

  const proxyReq = (isHttps ? https : http).request(opts, (proxyRes) => {
    const contentType = proxyRes.headers['content-type'] || '';
    const isHtml = contentType.includes('text/html');

    if (proxyRes.statusCode >= 300 && proxyRes.statusCode < 400 && proxyRes.headers.location) {
      const loc = proxyRes.headers.location;
      if (/^https?:\/\//i.test(loc)) {
        proxyRes.headers.location = targetToProxyUrl(loc);
      } else if (loc.startsWith('//')) {
        proxyRes.headers.location = targetToProxyUrl(parsed.protocol + loc);
      } else if (loc.startsWith('/')) {
        proxyRes.headers.location = targetToProxyUrl(parsed.protocol + '//' + parsed.host + loc);
      } else {
        const basePath = parsed.pathname.substring(0, parsed.pathname.lastIndexOf('/') + 1);
        proxyRes.headers.location = targetToProxyUrl(parsed.protocol + '//' + parsed.host + basePath + loc);
      }
    }

    delete proxyRes.headers['content-security-policy'];
    delete proxyRes.headers['x-frame-options'];
    delete proxyRes.headers['content-security-policy-report-only'];
    delete proxyRes.headers['x-content-security-policy'];
    delete proxyRes.headers['x-webkit-csp'];

    if (isHtml && proxyRes.statusCode === 200) {
      let chunks = [];
      proxyRes.on('data', (chunk) => chunks.push(chunk));
      proxyRes.on('end', () => {
        let body = Buffer.concat(chunks).toString('utf-8');
        body = rewriteHtml(body, target);
        proxyRes.headers['content-length'] = Buffer.byteLength(body);
        if (!proxyRes.headers['content-type']) proxyRes.headers['content-type'] = 'text/html; charset=utf-8';
        safeRespond(proxyRes.statusCode, proxyRes.headers, body);
      });
    } else {
      if (!responded) {
        responded = true;
        res.writeHead(proxyRes.statusCode, proxyRes.headers);
        proxyRes.pipe(res);
      }
    }
  });

  proxyReq.on('error', (err) => {
    safeRespond(502, { 'Content-Type': 'text/plain' }, 'Proxy error: ' + err.message);
  });

  proxyReq.on('timeout', () => {
    proxyReq.destroy();
    safeRespond(504, { 'Content-Type': 'text/plain' }, 'Proxy timeout');
  });

  if (req.method === 'GET' || req.method === 'HEAD') {
    proxyReq.end();
  } else {
    req.pipe(proxyReq);
  }
});

// --- WebSocket ---
const sessions = new Map();

function getOrCreateSession(sessionId) {
  if (!sessions.has(sessionId)) {
    sessions.set(sessionId, { view: null, control: null });
  }
  return sessions.get(sessionId);
}

function broadcastToView(sessionId, data) {
  const session = sessions.get(sessionId);
  if (session && session.view && session.view.readyState === 1) {
    session.view.send(JSON.stringify(data));
  }
}

function broadcastToControl(sessionId, data) {
  const session = sessions.get(sessionId);
  if (session && session.control && session.control.readyState === 1) {
    session.control.send(JSON.stringify(data));
  }
}

wss.on('connection', (ws, req) => {
  const u = new URL(req.url, `http://${req.headers.host}`);
  const role = u.searchParams.get('role');
  const sessionId = u.searchParams.get('session') || 'default';

  if (role !== 'view' && role !== 'control') {
    ws.close(4001, 'Invalid role');
    return;
  }

  const session = getOrCreateSession(sessionId);

  if (role === 'view') {
    session.view = ws;
  } else {
    session.control = ws;
  }

  console.log(`[WS] ${role} joined session=${sessionId}`);

  ws.on('message', (raw) => {
    let data;
    try {
      data = JSON.parse(raw.toString());
    } catch {
      return;
    }

    if (role === 'control' && data.type === 'key') {
      console.log(`[WS] key event: ${data.event} key=${data.key} session=${sessionId}`);
      broadcastToView(sessionId, { type: 'key', key: data.key, event: data.event });
    }

    if (role === 'view' && data.type === 'nav') {
      console.log(`[WS] nav event: ${data.url} session=${sessionId}`);
      broadcastToControl(sessionId, { type: 'nav', url: data.url });
    }
  });

  ws.on('close', () => {
    if (role === 'view') {
      session.view = null;
    } else {
      session.control = null;
    }
    console.log(`[WS] ${role} left session=${sessionId}`);

    if (!session.view && !session.control) {
      sessions.delete(sessionId);
    }
  });
});

// ===== Helper: download file =====
function downloadFile(url, destPath) {
  return new Promise((resolve, reject) => {
    const parsed = new URL(url);
    const isHttps = parsed.protocol === 'https:';
    const port = parsed.port || (isHttps ? 443 : 80);

    const opts = {
      hostname: parsed.hostname,
      port,
      path: parsed.pathname + parsed.search,
      method: 'GET',
      headers: {
        'User-Agent': 'Mozilla/5.0 (compatible; PPTRemoteViewer/1.0)',
        'Accept': '*/*',
      },
      rejectUnauthorized: false,
      timeout: 30000,
    };

    const requester = isHttps ? https : http;
    const req = requester.request(opts, (res) => {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        // Follow redirect
        const redirectUrl = res.headers.location.startsWith('http')
          ? res.headers.location
          : `${parsed.protocol}//${parsed.host}${res.headers.location}`;
        return downloadFile(redirectUrl, destPath).then(resolve).catch(reject);
      }

      if (res.statusCode !== 200) {
        return reject(new Error(`Download failed with status ${res.statusCode}`));
      }

      const file = fs.createWriteStream(destPath);
      res.pipe(file);
      file.on('finish', () => {
        file.close();
        resolve();
      });
      file.on('error', reject);
    });

    req.on('error', reject);
    req.on('timeout', () => {
      req.destroy();
      reject(new Error('Download timeout'));
    });
    req.end();
  });
}

server.listen(PORT, '0.0.0.0', () => {
  console.log(`Server running on http://0.0.0.0:${PORT}`);
});
