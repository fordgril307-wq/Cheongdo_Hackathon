// Azure Functions (Node.js v4 프로그래밍 모델) — Static Web Apps 관리형 API
const { app } = require('@azure/functions');
const { parseRequest } = require('../lib/parser');
const { matchRides } = require('../lib/matcher');
const { villages, destinations } = require('../lib/places');
const { demoRequests, demoProfiles } = require('../lib/demo');

const json = (body, status = 200) => ({ status, jsonBody: body });

app.http('health', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'health',
  handler: async () =>
    json({ ok: true, service: 'hyodo-taxi', ai: Boolean(process.env.AZURE_OPENAI_ENDPOINT && process.env.AZURE_OPENAI_API_KEY) })
});

app.http('places', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'places',
  handler: async () => json({ villages, destinations })
});

app.http('demo', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'demo',
  handler: async () => json({ requests: demoRequests(), profiles: demoProfiles() })
});

app.http('parse', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'parse',
  handler: async (request) => {
    const body = await request.json().catch(() => ({}));
    if (!body.text) return json({ error: 'text가 필요합니다' }, 400);
    const result = await parseRequest(body.text, body.profile || {});
    return json(result);
  }
});

app.http('match', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'match',
  handler: async (request) => {
    const body = await request.json().catch(() => ({}));
    const requests = Array.isArray(body.requests) ? body.requests : [];
    return json(matchRides(requests, body.fleet || { van: 0 }));
  }
});
