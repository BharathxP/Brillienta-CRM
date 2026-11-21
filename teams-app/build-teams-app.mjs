// Builds the Microsoft Teams app package for Brillienta CRM.
//  - teams/TeamsSPFxApp.zip  → included in the .sppkg; used by "Sync to Teams" in the App Catalog
//  - dist/BrillientaCRM-Teams.zip → for manual upload in the Teams admin center
import { copyFileSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { execSync } from 'node:child_process';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = join(HERE, '..');
const wpText = readFileSync(join(ROOT, 'src/webparts/brillientaCrm/BrillientaCrmWebPart.manifest.json'), 'utf8');
const wp = { id: (wpText.match(/"id"\s*:\s*"([0-9a-f-]{36})"/i) || [])[1] };
const pkg = JSON.parse(readFileSync(join(ROOT, 'package.json'), 'utf8'));
const sol = JSON.parse(readFileSync(join(ROOT, 'config/package-solution.json'), 'utf8')).solution;
const ID = wp.id;
const version = sol.version.split('.').slice(0, 3).join('.');

const manifest = {
  $schema: 'https://developer.microsoft.com/en-us/json-schemas/teams/v1.17/MicrosoftTeams.schema.json',
  manifestVersion: '1.17',
  version,
  id: ID,
  packageName: 'com.brillienta.crm',
  developer: {
    name: 'Brillienta',
    websiteUrl: 'https://brillienta.com',
    privacyUrl: 'https://brillienta.com/privacy',
    termsOfUseUrl: 'https://brillienta.com/terms'
  },
  name: { short: 'Brillienta CRM', full: 'Brillienta CRM: sales, marketing and service for Microsoft 365' },
  description: {
    short: 'CRM inside Teams & SharePoint: leads, deals, quotes, demos, cases, forecasts.',
    full: 'Brillienta CRM is a complete CRM that runs inside Microsoft Teams and SharePoint and keeps all data in your own Microsoft 365 tenant. ' +
      'Capture leads from Microsoft AppSource, web forms and imports with automatic scoring and assignment; manage accounts and contacts; ' +
      'move deals through a drag-and-drop pipeline; create branded quotes with discount approvals; schedule product demos as Teams meetings; ' +
      'track support cases against SLAs; and see dashboards, forecasts with quotas and reports. Integrates with Outlook, Teams and Power Automate.'
  },
  icons: { outline: 'outline.png', color: 'color.png' },
  accentColor: '#008D62',
  staticTabs: [
    {
      entityId: 'brillienta-crm',
      name: 'CRM',
      contentUrl: `https://{teamSiteDomain}/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/_layouts/15/teamshostedapp.aspx%3Fteams%26personal%26componentId=${ID}%26forceLocale={locale}`,
      websiteUrl: 'https://brillienta.com',
      scopes: ['personal']
    },
    { entityId: 'about', scopes: ['personal'] }
  ],
  configurableTabs: [
    {
      configurationUrl: `https://{teamSiteDomain}{teamSitePath}/_layouts/15/TeamsLogon.aspx?SPFX=true&dest={teamSitePath}/_layouts/15/teamshostedapp.aspx%3FopenPropertyPane=true%26teams%26componentId=${ID}%26forceLocale={locale}`,
      canUpdateConfiguration: true,
      scopes: ['team'],
      context: ['channelTab', 'privateChatTab', 'meetingChatTab']
    }
  ],
  permissions: ['identity'],
  validDomains: ['*.login.microsoftonline.com', '*.sharepoint.com', '*.sharepoint-df.com', 'spoppe-a.akamaihd.net', 'spoprod-a.akamaihd.net', 'resourceseng.blob.core.windows.net', 'msft.spoppe.com'],
  webApplicationInfo: { resource: 'https://{teamSiteDomain}', id: '00000003-0000-0ff1-ce00-000000000000' }
};

// ---- checks against the Teams manifest rules we rely on
const errs = [];
if (manifest.name.short.length > 30) { errs.push('name.short > 30'); }
if (manifest.name.full.length > 100) { errs.push('name.full > 100'); }
if (manifest.description.short.length > 80) { errs.push(`description.short > 80 (${manifest.description.short.length})`); }
if (manifest.description.full.length > 4000) { errs.push('description.full > 4000'); }
if (!/^[0-9a-f-]{36}$/i.test(manifest.id)) { errs.push('id must be a GUID'); }
if (!/^\d+\.\d+\.\d+$/.test(manifest.version)) { errs.push('version must be semver'); }
if (!/^#[0-9A-F]{6}$/i.test(manifest.accentColor)) { errs.push('accentColor'); }
if (errs.length) { console.error(errs.join('\n')); process.exit(1); }

const OUT = join(HERE, 'dist');
rmSync(OUT, { recursive: true, force: true });
mkdirSync(join(OUT, 'pkg'), { recursive: true });
writeFileSync(join(OUT, 'pkg', 'manifest.json'), JSON.stringify(manifest, null, 2));
copyFileSync(join(ROOT, 'teams', `${ID}_color.png`), join(OUT, 'pkg', 'color.png'));
copyFileSync(join(ROOT, 'teams', `${ID}_outline.png`), join(OUT, 'pkg', 'outline.png'));
execSync(`cd "${join(OUT, 'pkg')}" && zip -qX ../BrillientaCRM-Teams.zip manifest.json color.png outline.png`);
copyFileSync(join(OUT, 'BrillientaCRM-Teams.zip'), join(ROOT, 'teams', 'TeamsSPFxApp.zip'));
console.log(`Teams app ${manifest.name.short} v${version} (${pkg.name}) → teams/TeamsSPFxApp.zip, teams-app/dist/BrillientaCRM-Teams.zip`);
