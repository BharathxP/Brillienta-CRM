// Issues a license key for one customer tenant.
// Usage:
//   node tools/license/issue-license.mjs --customer "Contoso Ltd" --tenant <AAD tenant GUID> --plan enterprise --seats 50 --expires 2027-12-31
// The customer's tenant ID is shown in the CRM under Settings > License.
import { createPrivateKey, sign, randomUUID } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const args = Object.fromEntries(process.argv.slice(2).reduce((acc, a, i, arr) => {
  if (a.startsWith('--')) { acc.push([a.slice(2), arr[i + 1]]); }
  return acc;
}, []));
for (const k of ['customer', 'tenant', 'plan', 'expires']) {
  if (!args[k]) { console.error(`Missing --${k}`); process.exit(1); }
}
if (!['standard', 'enterprise'].includes(args.plan)) { console.error('--plan must be standard or enterprise'); process.exit(1); }
const here = dirname(fileURLToPath(import.meta.url));
const key = createPrivateKey(readFileSync(join(here, 'keys/private-key.pem')));
const payload = {
  id: randomUUID(), customer: args.customer, tenantId: args.tenant.toLowerCase(), plan: args.plan,
  seats: Number(args.seats || 0), issued: new Date().toISOString().slice(0, 10), expires: args.expires
};
const b64u = buf => Buffer.from(buf).toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
const body = b64u(JSON.stringify(payload));
const sig = sign('sha256', Buffer.from(body), { key, dsaEncoding: 'ieee-p1363' });
console.log(JSON.stringify(payload, null, 2));
console.log('\nLICENSE KEY:\n' + body + '.' + b64u(sig));
