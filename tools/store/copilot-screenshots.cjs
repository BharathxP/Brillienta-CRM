// Store screenshots for brillienta-crm-copilot.sppkg (1366 x 768): the Copilot cards with the sample CRM, from the
// local preview (preview-dist/copilot.html). Run after `npm run preview`:  node tools/store/copilot-screenshots.cjs
// Writes copilot-solution/images/screenshot-1..4.png.
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..');
const BASE = 'file://' + path.join(ROOT, 'preview-dist');
const OUT = path.join(ROOT, 'copilot-solution', 'images');
fs.mkdirSync(OUT, { recursive: true });

(async () => {
  const browser = await chromium.launch();
  const p = await browser.newPage({ viewport: { width: 1366, height: 768 }, deviceScaleFactor: 1 });
  const errors = [];
  p.on('pageerror', (e) => errors.push(e.message));
  // The sample CRM, set up once through the web part preview (same browser storage as copilot.html).
  await p.goto(`${BASE}/index.html?as=asha&reset=1`);
  await p.click('button:has-text("Set up CRM")');
  await p.waitForSelector('text=AI BRIEFING', { timeout: 30000 });
  await p.evaluate(() => window.__crm.seedDemo());
  const shots = [['pipeline', 'chen'], ['account', 'chen'], ['lead', 'chen'], ['cases', 'dana']];
  let n = 0;
  for (const [tool, as] of shots) {
    await p.goto(`${BASE}/copilot.html?tool=${tool}&as=${as}`);
    await p.waitForTimeout(3500);
    await p.screenshot({ path: path.join(OUT, `screenshot-${++n}.png`) });
  }
  await browser.close();
  if (errors.length) { console.error('Browser errors:', errors.join(' | ')); process.exit(1); }
  console.log(`Copilot screenshots written to ${OUT}`);
})();
