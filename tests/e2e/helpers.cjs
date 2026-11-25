// Shared helpers for the end-to-end journeys (npm run e2e). They drive the real CRM in the local preview
// (preview-dist/index.html) against the in-memory SharePoint, whose state is read back to check what was saved.
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..');
const URL = 'file://' + path.join(ROOT, 'preview-dist', 'index.html');
const OUT = path.join(__dirname, 'output');
fs.mkdirSync(OUT, { recursive: true });

const results = [];
/** A label as written, optionally followed by " (USD)", "(%)" or similar. */
const startsRe = (label) => new RegExp('^' + label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(\\s*\\(.*?\\))?(\\s|$)');
const labelRe = (label) => new RegExp('^' + label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(\\s*\\(.*\\))?\\s*\\*?$');
function check(name, ok, detail) {
  results.push({ name, ok: !!ok, detail });
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${detail !== undefined && detail !== '' ? '  — ' + String(detail).slice(0, 300) : ''}`);
}

async function start() {
  const browser = await chromium.launch();
  const ctx = await browser.newContext({ viewport: { width: 1366, height: 900 }, acceptDownloads: true });
  const p = await ctx.newPage();
  const errors = [];
  p.on('pageerror', (e) => errors.push(e.message));
  p.on('console', (m) => {
    const t = m.text();
    // Errors, and warnings from background jobs (sequences, maintenance, SLA updates) that only log when they fail.
    if ((m.type() === 'error' && !/Failed to load resource/.test(t)) || (m.type() === 'warning' && /fail|error|Territory routing|Sequences|Signed quotes/i.test(t))) { errors.push(`${m.type()}: ${t.slice(0, 300)}`); }
  });
  const answers = [];
  p.on('dialog', (d) => d.accept(answers.length ? String(answers.shift()) : d.defaultValue()));
  const panel = () => p.locator('.ms-Panel.is-open').last();
  const h = {
    p, errors,
    /** Answers for the next window.prompt() calls, in order (confirm() is always accepted). */
    answer(...values) { answers.push(...values); },
    /** Stands in for a Power Automate flow updating an item. */
    async flow(list, id, fields) { await p.evaluate(([l, i, f]) => window.__crm.patchItem(l, i, f), [list, id, fields]); },
    /** Opens the CRM as someone (asha admin, ben manager, chen/maya sales, dana support, eli reader) on a page. */
    async go(as, page, extra = '') {
      await p.goto(`${URL}?as=${as}&page=${encodeURIComponent(page || 'dashboard')}${extra}`);
      await p.waitForSelector('.crm-nav-inner, :text("Welcome to Brillienta CRM")', { timeout: 20000 });
      await p.waitForTimeout(900);
    },
    async nav(label) {
      await p.click(`.crm-nav-inner button:has-text("${label}")`);
      await p.waitForTimeout(900);
    },
    panel,
    async fill(label, v, scope) {
      const root = scope || panel();
      await root.getByLabel(labelRe(label)).first().fill(String(v));
    },
    async choose(label, option, scope) {
      const root = scope || panel();
      await root.getByRole('combobox', { name: startsRe(label) }).first().click();
      await p.getByRole('option', { name: option, exact: true }).first().click();
      await p.waitForTimeout(150);
    },
    /** A record command, from the command bar or its "More actions" overflow menu. */
    async command(name) {
      const btn = panel().locator(`.ms-CommandBar button:has-text("${name}")`).first();
      if (await btn.isVisible()) { await btn.click(); } else {
        await panel().locator('button[aria-label="More actions"]').first().click(); await p.waitForTimeout(300);
        await p.getByRole('menuitem', { name }).first().click();
      }
      await p.waitForTimeout(1500);
    },
    /** Is a record command shown but disabled? */
    async commandDisabled(name) {
      const btn = panel().locator(`.ms-CommandBar button:has-text("${name}")`).first();
      if (await btn.isVisible()) { return btn.isDisabled(); }
      await panel().locator('button[aria-label="More actions"]').first().click(); await p.waitForTimeout(300);
      const item = p.getByRole('menuitem', { name }).first();
      const off = (await item.getAttribute('aria-disabled')) === 'true' || await item.isDisabled();
      await p.keyboard.press('Escape'); await p.waitForTimeout(200);
      return off;
    },
    async click(selector) { await p.click(selector); await p.waitForTimeout(700); },
    async save(button = 'Create & close') {
      await panel().locator(`button:has-text("${button}")`).first().click();
      await p.waitForTimeout(1500);
    },
    async openRecord(text) { await p.click(`.crm-main >> text="${text}"`); await p.waitForTimeout(1100); },
    async toast() { return (await p.$$eval('.crm-toasts .ms-MessageBar, .ms-MessageBar', (m) => m.map((x) => x.innerText.trim()).join(' | '))) || ''; },
    async state() { return p.evaluate(() => window.__crm.state()); },
    async items(list) { return (await h.state()).lists[list].items; },
    async visible(selector) { return p.isVisible(selector); },
    async text(selector = 'body') { return (await p.innerText(selector)).replace(/\s+/g, ' '); },
    async shot(name, full = false) { await p.screenshot({ path: path.join(OUT, name + '.png'), fullPage: full }); }
  };
  return { browser, h };
}

function summary() {
  const failed = results.filter((r) => !r.ok);
  console.log(`\n${results.length - failed.length} passed, ${failed.length} failed`);
  failed.forEach((f) => console.log(`  FAILED: ${f.name}${f.detail ? ' — ' + f.detail : ''}`));
  return failed.length;
}

module.exports = { URL, OUT, ROOT, check, start, summary, results };
