// Kite order relay (v5) - sends orders to Zerodha from a STATIC IP, through an HTTP proxy (TrueIP or any provider with the same
// host:port:user:pass model). Zerodha accepts API orders only from the IP registered in the account's developer profile.
//
// v5: ONE relay can serve many people. A request may carry its own `proxy_url` (that person's own static-IP proxy); without one the
// relay uses STATIC_IP_PROXY_URL (the relay owner's own IP). Adds /modify-order, /cancel-order and /check-proxy.
//
// Security: every request must include RELAY_SECRET. The relay never stores or logs api keys, access tokens or proxy URLs.

const express = require('express');
const https = require('https');
const crypto = require('crypto');
const { HttpsProxyAgent } = require('https-proxy-agent');

const app = express();
app.use(express.json({ limit: '50kb' }));

const PORT = process.env.PORT || 8080;
const RELAY_SECRET = process.env.RELAY_SECRET;
const PROXY_URL = process.env.STATIC_IP_PROXY_URL; // optional in v5: http://USER:PASS@gw.trueip.in:PORT (the owner's own IP)

if (!RELAY_SECRET) {
  console.error('RELAY_SECRET is not set. Refusing to start - this relay would accept unauthenticated order requests otherwise.');
  process.exit(1);
}
if (!PROXY_URL) {
  console.warn('STATIC_IP_PROXY_URL is not set. Requests that do not bring their own proxy_url will be refused (they would leave from this server\'s own IP, which Kite rejects).');
}

const defaultAgent = PROXY_URL ? new HttpsProxyAgent(PROXY_URL) : null;
const agentCache = new Map(); // proxy_url -> agent (small LRU)

function secretOk(given) {
  const a = Buffer.from(String(given || '')), b = Buffer.from(String(RELAY_SECRET));
  return a.length === b.length && crypto.timingSafeEqual(a, b);
}

// A caller-supplied proxy must be a plain http(s) proxy on a PUBLIC host - never the relay's own network.
function isPrivateHost(h) {
  h = String(h || '').toLowerCase().replace(/^\[|\]$/g, '');
  if (!h || h === 'localhost' || h.endsWith('.local') || h.endsWith('.internal') || h.endsWith('.localhost')) return true;
  if (/^(::1|::|fc|fd|fe[89ab])/i.test(h) && h.includes(':')) return true;
  const m = /^(\d+)\.(\d+)\.(\d+)\.(\d+)$/.exec(h);
  if (m) {
    const [a, b] = [Number(m[1]), Number(m[2])];
    if (a === 0 || a === 10 || a === 127 || (a === 169 && b === 254) || (a === 172 && b >= 16 && b <= 31) || (a === 192 && b === 168) || a >= 224) return true;
  }
  return false;
}
function agentFor(proxyUrl) {
  if (!proxyUrl) {
    if (!defaultAgent) throw new Error('This relay has no default proxy and the request did not include proxy_url.');
    return defaultAgent;
  }
  const s = String(proxyUrl);
  if (s.length > 500) throw new Error('proxy_url is too long.');
  let u;
  try { u = new URL(s); } catch (e) { throw new Error('proxy_url is not a valid URL.'); }
  if (u.protocol !== 'http:' && u.protocol !== 'https:') throw new Error('proxy_url must start with http:// or https://');
  if (isPrivateHost(u.hostname)) throw new Error('proxy_url must point to a public proxy host.');
  const port = Number(u.port || (u.protocol === 'https:' ? 443 : 80));
  if (!(port >= 1 && port <= 65535)) throw new Error('proxy_url has an invalid port.');
  let ag = agentCache.get(s);
  if (!ag) {
    ag = new HttpsProxyAgent(s);
    agentCache.set(s, ag);
    if (agentCache.size > 50) agentCache.delete(agentCache.keys().next().value);
  }
  return ag;
}

function log(body, outcome) {
  console.log(new Date().toISOString(), outcome, (body && body.tradingsymbol) || (body && body.order_id) || '', (body && body.transaction_type) || '', (body && body.quantity) || '', (body && body.order_type) || '');
}

// One HTTPS request to api.kite.trade through the chosen proxy.
function kiteRequest(method, path, params, headers, agent) {
  return new Promise((resolve, reject) => {
    const payload = params ? params.toString() : '';
    const h = Object.assign({}, headers);
    if (payload) { h['Content-Type'] = 'application/x-www-form-urlencoded'; h['Content-Length'] = Buffer.byteLength(payload); }
    const req = https.request({ hostname: 'api.kite.trade', path, method, agent, headers: h }, (res) => {
      let body = '';
      res.on('data', (c) => { body += c; });
      res.on('end', () => {
        try { resolve({ statusCode: res.statusCode, data: JSON.parse(body) }); }
        catch (e) { reject(new Error('Kite returned a non-JSON response: ' + body.slice(0, 200))); }
      });
    });
    req.on('error', reject);
    if (payload) req.write(payload);
    req.end();
  });
}

function authHeaders(body) {
  return { 'X-Kite-Version': '3', 'Authorization': `token ${body.api_key}:${body.access_token}` };
}
function pick(body, keys) {
  const f = new URLSearchParams();
  keys.forEach((k) => { if (body[k] !== undefined && body[k] !== null && body[k] !== '') f.append(k, body[k]); });
  return f;
}
function guard(req, res, needOrderId) {
  const body = req.body || {};
  if (!secretOk(body.relay_secret)) { log(body, 'REJECTED (bad secret)'); res.status(403).json({ success: false, error: 'Invalid relay secret.' }); return null; }
  if (!body.api_key || !body.access_token) { res.status(400).json({ success: false, error: 'Missing api_key or access_token.' }); return null; }
  if (needOrderId && !/^[0-9A-Za-z_-]{6,40}$/.test(String(body.order_id || ''))) { res.status(400).json({ success: false, error: 'Missing or invalid order_id.' }); return null; }
  let agent;
  try { agent = agentFor(body.proxy_url); } catch (e) { res.status(400).json({ success: false, error: e.message }); return null; }
  return { body, agent };
}
function sendKite(res, body, kiteData, okLabel) {
  if (kiteData.status !== 'success') {
    log(body, 'REJECTED by Kite: ' + kiteData.error_type);
    return res.json({ success: false, error: kiteData.message || 'Kite rejected the request.', error_type: kiteData.error_type });
  }
  log(body, okLabel + ' ' + (kiteData.data && kiteData.data.order_id));
  return res.json({ success: true, order_id: kiteData.data && kiteData.data.order_id });
}
const VARIETIES = ['regular', 'amo', 'co', 'iceberg', 'auction'];
const variety = (b) => (VARIETIES.includes(String(b.variety || 'regular').toLowerCase()) ? String(b.variety || 'regular').toLowerCase() : 'regular');

// ---- place ----
app.post('/place-order', async (req, res) => {
  const g = guard(req, res, false); if (!g) return;
  const body = g.body;
  if (!body.tradingsymbol || !body.transaction_type || !body.quantity || !body.order_type) {
    return res.status(400).json({ success: false, error: 'Missing tradingsymbol, transaction_type, quantity, or order_type.' });
  }
  // Everything Kite's order endpoint accepts, forwarded as-is (market_protection is required by Kite on API MARKET orders).
  const params = pick(body, ['tradingsymbol', 'exchange', 'transaction_type', 'order_type', 'quantity', 'product', 'validity', 'price', 'trigger_price', 'disclosed_quantity', 'squareoff', 'stoploss', 'trailing_stoploss', 'tag', 'market_protection']);
  try {
    const { data } = await kiteRequest('POST', '/orders/regular', params, authHeaders(body), g.agent);
    return sendKite(res, body, data, 'PLACED');
  } catch (err) {
    console.error(new Date().toISOString(), 'RELAY ERROR', err.message);
    return res.status(502).json({ success: false, error: 'Relay could not reach Kite through the proxy: ' + err.message });
  }
});

// ---- modify a pending order ----
app.post('/modify-order', async (req, res) => {
  const g = guard(req, res, true); if (!g) return;
  const body = g.body;
  const params = pick(body, ['order_type', 'quantity', 'price', 'trigger_price', 'validity', 'disclosed_quantity']);
  try {
    const { data } = await kiteRequest('PUT', `/orders/${variety(body)}/${body.order_id}`, params, authHeaders(body), g.agent);
    return sendKite(res, body, data, 'MODIFIED');
  } catch (err) {
    console.error(new Date().toISOString(), 'RELAY ERROR', err.message);
    return res.status(502).json({ success: false, error: 'Relay could not reach Kite through the proxy: ' + err.message });
  }
});

// ---- cancel a pending order ----
app.post('/cancel-order', async (req, res) => {
  const g = guard(req, res, true); if (!g) return;
  const body = g.body;
  try {
    const { data } = await kiteRequest('DELETE', `/orders/${variety(body)}/${body.order_id}`, null, authHeaders(body), g.agent);
    return sendKite(res, body, data, 'CANCELLED');
  } catch (err) {
    console.error(new Date().toISOString(), 'RELAY ERROR', err.message);
    return res.status(502).json({ success: false, error: 'Relay could not reach Kite through the proxy: ' + err.message });
  }
});

function egressIp(agent) {
  return new Promise((resolve, reject) => {
    const r = https.request({ hostname: 'ifconfig.me', path: '/ip', method: 'GET', agent }, (resp) => {
      let b = ''; resp.on('data', (c) => (b += c)); resp.on('end', () => resolve(b.trim()));
    });
    r.on('error', reject); r.end();
  });
}

// ---- which IP does a proxy leave from? (so a user can check it matches the IP whitelisted on Zerodha) ----
app.post('/check-proxy', async (req, res) => {
  const body = req.body || {};
  if (!secretOk(body.relay_secret)) return res.status(403).json({ ok: false, error: 'Invalid relay secret.' });
  try {
    const ip = await egressIp(agentFor(body.proxy_url));
    return res.json({ ok: true, egressIp: ip, usingOwnProxy: !!body.proxy_url });
  } catch (e) {
    return res.json({ ok: false, error: 'Proxy check failed: ' + e.message });
  }
});

// Confirms the relay is alive AND that its DEFAULT proxy works (the IP shown must be the one you whitelisted with Zerodha).
// Also a keep-alive target for hosts that put idle services to sleep (Render's free tier).
app.get('/health', async (req, res) => {
  if (!defaultAgent) return res.json({ ok: true, time: new Date().toISOString(), defaultProxy: false });
  try {
    res.json({ ok: true, time: new Date().toISOString(), egressIpThroughProxy: await egressIp(defaultAgent) });
  } catch (e) {
    res.status(500).json({ ok: false, error: 'Relay is up but the proxy leg failed: ' + e.message });
  }
});

app.listen(PORT, () => console.log(`Kite order relay (v5) listening on port ${PORT}${PROXY_URL ? ', default proxy configured' : ', no default proxy (callers bring their own)'}`));
