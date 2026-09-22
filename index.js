// Kite order relay — routes the actual order-placement call through TrueIP's proxy (or any
// HTTP-proxy-based static-IP provider using the same host:port:user:pass model), since that's
// what satisfies Kite's static-IP requirement for placing orders. This server itself does NOT
// need its own static IP — TrueIP's infrastructure makes the final hop to Kite from its own
// dedicated address, so this can run anywhere reachable over HTTPS from Apps Script.
//
// Security model: a single shared secret (RELAY_SECRET) that only this server and your Apps
// Script's Script Properties know. Every request must include it or gets rejected outright. This
// server never stores your Kite API key/access token — it just forwards them through to Kite on
// each request, exactly as sent, and never logs them.

const express = require('express');
const https = require('https');
const { HttpsProxyAgent } = require('https-proxy-agent');

const app = express();
app.use(express.json());

const PORT = process.env.PORT || 8080;
const RELAY_SECRET = process.env.RELAY_SECRET;
const PROXY_URL = process.env.STATIC_IP_PROXY_URL; // e.g. http://USER:PASS@gw.trueip.in:PORT

if (!RELAY_SECRET) {
  console.error('RELAY_SECRET is not set. Refusing to start — this relay would accept unauthenticated order requests otherwise.');
  process.exit(1);
}
if (!PROXY_URL) {
  console.error('STATIC_IP_PROXY_URL is not set. Refusing to start — without it, the order request would leave from this server\'s own IP, which Kite will reject.');
  process.exit(1);
}

const proxyAgent = new HttpsProxyAgent(PROXY_URL);

// Never logs the secret, api_key, or access_token — only what's useful for debugging a rejected
// or failed order.
function logRequest(body, outcome) {
  console.log(new Date().toISOString(), outcome, body.tradingsymbol, body.transaction_type, body.quantity, body.order_type);
}

// Routes through proxyAgent explicitly, so the request to Kite always egresses from the static IP
// regardless of what network interfaces this server itself has (the IPv4-leak issue some Kite
// Connect users have hit when their own server had both IPv4 and IPv6 available — routing through
// an explicit proxy sidesteps that, since the final hop to Kite happens inside the proxy's own
// infrastructure, never on this machine's own interfaces).
function postToKiteViaProxy(params, headers) {
  return new Promise((resolve, reject) => {
    const req = https.request({
      hostname: 'api.kite.trade',
      path: '/orders/regular',
      method: 'POST',
      agent: proxyAgent,
      headers: Object.assign({ 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(params) }, headers),
    }, (res) => {
      let body = '';
      res.on('data', (chunk) => { body += chunk; });
      res.on('end', () => {
        try { resolve({ statusCode: res.statusCode, data: JSON.parse(body) }); }
        catch (e) { reject(new Error('Kite returned a non-JSON response: ' + body.slice(0, 200))); }
      });
    });
    req.on('error', reject);
    req.write(params);
    req.end();
  });
}

app.post('/place-order', async (req, res) => {
  const body = req.body || {};

  if (body.relay_secret !== RELAY_SECRET) {
    logRequest(body, 'REJECTED (bad secret)');
    return res.status(403).json({ success: false, error: 'Invalid relay secret.' });
  }
  if (!body.api_key || !body.access_token) {
    return res.status(400).json({ success: false, error: 'Missing api_key or access_token.' });
  }
  if (!body.tradingsymbol || !body.transaction_type || !body.quantity || !body.order_type) {
    return res.status(400).json({ success: false, error: 'Missing tradingsymbol, transaction_type, quantity, or order_type.' });
  }

  // Everything Kite's order endpoint accepts, forwarded through as-is — this relay doesn't
  // second-guess the order shape beyond the presence checks above; that validation already
  // happened in kitePlaceOrderFromWeb before this request was ever sent.
  const kiteParams = new URLSearchParams();
  ['tradingsymbol', 'exchange', 'transaction_type', 'order_type', 'quantity', 'product', 'validity', 'price', 'trigger_price', 'disclosed_quantity', 'squareoff', 'stoploss', 'trailing_stoploss', 'tag']
    .forEach((k) => { if (body[k] !== undefined && body[k] !== null && body[k] !== '') kiteParams.append(k, body[k]); });

  try {
    const { data: kiteData } = await postToKiteViaProxy(kiteParams.toString(), {
      'X-Kite-Version': '3',
      'Authorization': `token ${body.api_key}:${body.access_token}`,
    });

    if (kiteData.status !== 'success') {
      logRequest(body, 'REJECTED by Kite: ' + kiteData.error_type);
      // Passed straight through — kitePlaceOrderFromWeb on the Apps Script side surfaces
      // error_type + message directly rather than a generic failure.
      return res.json({ success: false, error: kiteData.message || 'Kite rejected the order.', error_type: kiteData.error_type });
    }

    logRequest(body, 'PLACED ' + kiteData.data.order_id);
    return res.json({ success: true, order_id: kiteData.data.order_id });
  } catch (err) {
    console.error(new Date().toISOString(), 'RELAY ERROR', err.message);
    return res.status(502).json({ success: false, error: 'Relay could not reach Kite through the proxy: ' + err.message });
  }
});

// Confirms the relay is alive AND that the proxy leg itself works, by asking TrueIP's own gateway
// (or your provider's equivalent) what IP it sees — check this returns your registered static IP
// before ever placing a real order through /place-order. Also doubles as a keepalive target if
// your host reclaims idle instances (Oracle Cloud does this).
app.get('/health', async (req, res) => {
  try {
    const result = await new Promise((resolve, reject) => {
      const r = https.request({ hostname: 'ifconfig.me', path: '/ip', method: 'GET', agent: proxyAgent }, (resp) => {
        let b = ''; resp.on('data', (c) => b += c); resp.on('end', () => resolve(b.trim()));
      });
      r.on('error', reject); r.end();
    });
    res.json({ ok: true, time: new Date().toISOString(), egressIpThroughProxy: result });
  } catch (e) {
    res.status(500).json({ ok: false, error: 'Relay is up but the proxy leg failed: ' + e.message });
  }
});

app.listen(PORT, () => console.log(`Kite order relay listening on port ${PORT}, routing orders through the configured static-IP proxy`));
