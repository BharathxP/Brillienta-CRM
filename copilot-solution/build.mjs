// Builds the second SharePoint package, brillienta-crm-copilot.sppkg: the Brillienta CRM Copilot agent and its six
// Copilot Components, as a solution of its own.
//
// Why a second package: in SPFx 1.24 "Sync to Teams" publishes one Teams app per .sppkg, and when a package has both
// Copilot Components and a Teams web part it publishes only the agent (sp-dev-docs#11041). Keeping the agent in its own
// package lets brillienta-crm.sppkg publish the CRM Teams app as before, and this package publish the agent.
//
// The Copilot Components share the CRM's code (src/webparts/brillientaCrm), so this script builds a staging copy of the
// project in temp/copilot-solution with this folder's config (bundle, package-solution, copilot-agent) and agent files
// (copilot/), then copies the package to sharepoint/solution/brillienta-crm-copilot.sppkg.
import { cpSync, existsSync, mkdirSync, rmSync, symlinkSync, copyFileSync, readdirSync } from 'node:fs';
import { execSync } from 'node:child_process';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = join(HERE, '..');
const STAGE = join(ROOT, 'temp', 'copilot-solution');

rmSync(STAGE, { recursive: true, force: true });
mkdirSync(join(STAGE, 'config'), { recursive: true });

// Project files shared with the CRM.
for (const f of ['package.json', 'tsconfig.json', 'eslint.config.js']) { copyFileSync(join(ROOT, f), join(STAGE, f)); }
cpSync(join(ROOT, 'src'), join(STAGE, 'src'), { recursive: true });
// Store images: the app icon (shared with the CRM) and this package's own screenshots (images/screenshot-*.png).
mkdirSync(join(STAGE, 'sharepoint', 'images'), { recursive: true });
copyFileSync(join(ROOT, 'sharepoint', 'images', 'AppIcon.png'), join(STAGE, 'sharepoint', 'images', 'AppIcon.png'));
cpSync(join(HERE, 'images'), join(STAGE, 'sharepoint', 'images'), { recursive: true });
symlinkSync(join(ROOT, 'node_modules'), join(STAGE, 'node_modules'), 'dir');

// Build config: the CRM's, with this package's bundle, solution and agent instead.
const own = new Set(readdirSync(join(HERE, 'config')));
for (const f of readdirSync(join(ROOT, 'config'))) {
  if (!own.has(f) && f !== 'jest.config.json') { copyFileSync(join(ROOT, 'config', f), join(STAGE, 'config', f)); }
}
for (const f of own) { copyFileSync(join(HERE, 'config', f), join(STAGE, 'config', f)); }
cpSync(join(HERE, 'copilot'), join(STAGE, 'copilot'), { recursive: true });
mkdirSync(join(STAGE, 'teams'), { recursive: true });

// Tests run in the main CRM build; here only build and package.
const run = (cmd) => execSync(cmd, { cwd: STAGE, stdio: 'inherit' });
run('npx heft build --clean --production');
run('npx heft package-solution --production');

const built = join(STAGE, 'sharepoint', 'solution', 'brillienta-crm-copilot.sppkg');
if (!existsSync(built)) { throw new Error(`Package not found: ${built}`); }
mkdirSync(join(ROOT, 'sharepoint', 'solution'), { recursive: true });
copyFileSync(built, join(ROOT, 'sharepoint', 'solution', 'brillienta-crm-copilot.sppkg'));
console.log('Brillienta CRM Copilot → sharepoint/solution/brillienta-crm-copilot.sppkg');
