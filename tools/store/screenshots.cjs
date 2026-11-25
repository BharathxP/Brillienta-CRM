// Store screenshots (1366 x 768, as AppSource / the SharePoint Store require), taken from the local preview with the
// sample CRM (preview/demoData.js). Run after `npm run preview`:  node tools/store/screenshots.cjs
// Writes sharepoint/images/screenshot-1..5.png for brillienta-crm.sppkg.
const { chromium } = require('playwright');
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..');
const URL = 'file://' + path.join(ROOT, 'preview-dist', 'index.html');
const OUT = path.join(ROOT, 'sharepoint', 'images');

(async () => {
  const browser = await chromium.launch();
  const p = await browser.newPage({ viewport: { width: 1366, height: 768 }, deviceScaleFactor: 1 });
  const errors = [];
  p.on('pageerror', (e) => errors.push(e.message));
  const go = async (as, page) => {
    await p.goto(`${URL}?as=${as}&page=${encodeURIComponent(page)}`);
    await p.waitForSelector('.crm-nav-inner', { timeout: 20000 });
    await p.waitForTimeout(2500);
  };
  const shot = (n) => p.screenshot({ path: path.join(OUT, `screenshot-${n}.png`) });

  await p.goto(`${URL}?as=asha&reset=1`);
  await p.click('button:has-text("Set up CRM")');
  await p.waitForSelector('text=AI BRIEFING', { timeout: 30000 });
  await p.evaluate(() => window.__crm.seedDemo());
  // Company name on quotes and the dashboard
  await p.evaluate(() => {
    const st = window.__crm.state();
    const row = st.lists.BrillCRM_Settings.items[0];
    const v = JSON.parse(row.Value);
    v.company = { ...v.company, name: 'Brillienta Demo' };
    window.__crm.patchItem('BrillCRM_Settings', row.Id, { Value: JSON.stringify(v) });
  });
  // Maintenance would add renewal deals; keep the sample exactly as seeded.
  await p.evaluate(() => localStorage.setItem('brillienta-crm:maintenance:https://contoso.sharepoint.com/sites/crm', String(Date.now())));

  // 1. Dashboard with AI briefing, KPIs, next best actions and forecast
  await go('asha', 'dashboard');
  await shot(1);
  // 2. Pipeline board
  await go('chen', 'entity:opportunities');
  await shot(2);
  // 3. A deal with its smart summary, health and activity
  await go('chen', 'entity:opportunities');
  await p.click('.crm-main >> text="Contoso HR portal rollout"');
  await p.waitForTimeout(1800);
  await shot(3);
  // 4. A quote with line items
  await go('chen', 'entity:quotes');
  await p.click('.crm-main >> text="Contoso HR portal - 3 years"');
  await p.waitForTimeout(1200);
  await p.locator('.ms-Panel.is-open button:has-text("Line items")').first().click();
  await p.waitForTimeout(800);
  await shot(4);
  // 5. Forecast against quota (sales manager)
  await go('ben', 'forecast');
  await shot(5);

  await browser.close();
  if (errors.length) { console.error('Browser errors:', errors.join(' | ')); process.exit(1); }
  console.log(`Screenshots written to ${OUT}`);
})();
