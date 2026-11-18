// ---------------------------------------------------------------------------
// Brillienta CRM: forward website leads to the CRM
// Add this to your Vercel function  api/lead.js  (the endpoint your
// book-demo and contact forms already post to).
//
// 1. Vercel › Project › Settings › Environment Variables:
//      BRILLIENTA_CRM_FLOW_URL = <HTTP URL of the flow
//                                "Brillienta CRM - Website leads and demo bookings">
// 2. Paste forwardToCrm() into api/lead.js.
// 3. In your handler, AFTER your validation / spam checks pass and BEFORE you
//    send the success response, add:     await forwardToCrm(body);
//    (body = the parsed JSON your forms send: kind, name, email, company, role,
//     size, country, notes, products, duration, start, timeZone, attribution…)
//
// The flow URL stays on the server (never in your website's HTML).
// If the CRM is slow or unreachable, the visitor still gets your normal
// response and your existing email notification is unaffected.
// ---------------------------------------------------------------------------

async function forwardToCrm(payload) {
  const url = process.env.BRILLIENTA_CRM_FLOW_URL;
  if (!url) return { skipped: true };
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), 5000);
  try {
    const res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
      signal: ctrl.signal
    });
    if (!res.ok) console.error('Brillienta CRM: flow returned', res.status);
    return { ok: res.ok, status: res.status };
  } catch (err) {
    console.error('Brillienta CRM: forwarding failed', err && err.message);
    return { ok: false };
  } finally {
    clearTimeout(timer);
  }
}

module.exports.forwardToCrm = forwardToCrm; // optional (CommonJS); remove if api/lead.js uses ESM
