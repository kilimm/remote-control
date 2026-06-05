const express = require('express');
const http = require('http');
const https = require('https');
const { WebSocketServer } = require('ws');
const path = require('path');
const url = require('url');

const app = express();
const server = http.createServer(app);
const wss = new WebSocketServer({ server });

const PORT = 3001;

// --- Static files ---
app.use('/public', express.static(path.join(__dirname, 'public')));

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

// --- HTTP Proxy for iframe (makes target pages same-origin) ---
function targetToProxyUrl(targetUrl) {
  return '/proxy/' + targetUrl;
}

function rewriteHtml(body, targetBaseUrl) {
  // Rewrite absolute URLs in the HTML to go through our proxy
  body = body.replace(
    /(src|href|action|data)\s*=\s*"(https?:\/\/[^"]+)"/gi,
    (match, attr, val) => `${attr}="${targetToProxyUrl(val)}"`
  );
  body = body.replace(
    /(src|href|action|data)\s*=\s*'(https?:\/\/[^']+)'/gi,
    (match, attr, val) => `${attr}='${targetToProxyUrl(val)}'`
  );
  // Rewrite srcset
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
  // Rewrite CSS url() references
  body = body.replace(
    /url\(\s*["']?(https?:\/\/[^"')]+)["']?\s*\)/gi,
    (match, val) => `url("${targetToProxyUrl(val)}")`
  );
  return body;
}

// Proxy endpoint — handles all /proxy/* requests
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

  let target = req.originalUrl.slice(7); // Remove '/proxy/'

  if (!target) {
    safeRespond(400, { 'Content-Type': 'text/plain' }, 'Missing target URL');
    return;
  }

  if (!/^https?:\/\//i.test(target)) {
    target = 'http://' + target;
  }

  console.log(`[Proxy] ${req.method} ${target}`);

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

  // DO NOT forward fetch metadata / service-worker headers that affect routing
  delete opts.headers['sec-fetch-site'];
  delete opts.headers['sec-fetch-mode'];
  delete opts.headers['sec-fetch-dest'];
  delete opts.headers['sec-fetch-user'];
  // Also strip proxy-specific headers
  delete opts.headers['x-forwarded-for'];
  delete opts.headers['x-forwarded-proto'];
  delete opts.headers['x-forwarded-host'];

  const proxyReq = (isHttps ? https : http).request(opts, (proxyRes) => {
    const contentType = proxyRes.headers['content-type'] || '';
    const isHtml = contentType.includes('text/html');

    // Rewrite redirects
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

    // Remove headers that block iframe embedding
    delete proxyRes.headers['content-security-policy'];
    delete proxyRes.headers['x-frame-options'];
    delete proxyRes.headers['content-security-policy-report-only'];
    delete proxyRes.headers['x-content-security-policy'];
    delete proxyRes.headers['x-webkit-csp'];

    if (isHtml && proxyRes.statusCode === 200) {
      // Collect HTML body and rewrite URLs
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
      // Stream non-HTML content — pipe directly
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

  // End the request - for GET/HEAD this sends immediately; for POST/PUT we pipe body first
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

server.listen(PORT, '0.0.0.0', () => {
  console.log(`Server running on http://0.0.0.0:${PORT}`);
});
