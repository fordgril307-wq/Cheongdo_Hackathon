// 로컬 미리보기 서버 (Azure 없이 화면·API 확인용)
// 실행: node local/dev-server.js  →  http://localhost:4280
const http = require('http');
const fs = require('fs');
const path = require('path');
const { parseRequest } = require('../api/src/lib/parser');
const { matchRides } = require('../api/src/lib/matcher');
const { villages, destinations } = require('../api/src/lib/places');
const { demoRequests, demoProfiles } = require('../api/src/lib/demo');

const ROOT = path.join(__dirname, '..', 'app');
const TYPES = { '.html': 'text/html; charset=utf-8', '.css': 'text/css', '.js': 'text/javascript', '.json': 'application/json' };

function send(res, status, body) {
  res.writeHead(status, { 'Content-Type': 'application/json; charset=utf-8' });
  res.end(JSON.stringify(body));
}

async function readJson(req) {
  let data = '';
  for await (const chunk of req) data += chunk;
  try { return JSON.parse(data || '{}'); } catch { return {}; }
}

http.createServer(async (req, res) => {
  const url = new URL(req.url, 'http://localhost');
  if (url.pathname.startsWith('/api/')) {
    const route = url.pathname.slice(5);
    if (route === 'health') return send(res, 200, { ok: true, ai: Boolean(process.env.AZURE_OPENAI_API_KEY) });
    if (route === 'places') return send(res, 200, { villages, destinations });
    if (route === 'demo') return send(res, 200, { requests: demoRequests(), profiles: demoProfiles() });
    if (route === 'parse' && req.method === 'POST') {
      const body = await readJson(req);
      if (!body.text) return send(res, 400, { error: 'text가 필요합니다' });
      return send(res, 200, await parseRequest(body.text, body.profile || {}));
    }
    if (route === 'match' && req.method === 'POST') {
      const body = await readJson(req);
      return send(res, 200, matchRides(body.requests || [], body.fleet || { van: 0 }));
    }
    return send(res, 404, { error: 'not found' });
  }
  const file = path.join(ROOT, url.pathname === '/' ? 'index.html' : url.pathname);
  if (!file.startsWith(ROOT) || !fs.existsSync(file)) {
    res.writeHead(404); return res.end('not found');
  }
  res.writeHead(200, { 'Content-Type': TYPES[path.extname(file)] || 'application/octet-stream' });
  fs.createReadStream(file).pipe(res);
}).listen(process.env.PORT || 4280, () => console.log('청도효도택시 로컬: http://localhost:' + (process.env.PORT || 4280)));
