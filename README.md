# Brillienta CRM: enterprise CRM for SharePoint & Microsoft Teams

A complete CRM delivered as one SharePoint Framework (SPFx 1.24) web part. All data is stored in SharePoint lists **in the customer's own Microsoft 365 tenant**. There is no external database or server.

## Modules

| Area | Features |
|---|---|
| **Home** | Role-aware dashboard (quarter KPIs with change vs. last quarter, 12-month revenue chart, pipeline by stage, conversion funnel, lead sources, hottest leads, biggest deals, demos, tasks) · **My work** queue (overdue/today/upcoming tasks, follow-ups, new leads, demos, deals closing, cases at risk) · **Calendar** (demos, tasks, expected closes, SLA deadlines) |
| **Leads** | AppSource/Partner Center, web form, import, manual entry · automatic **lead scoring** (0–100, configurable rules) · **round-robin assignment** · duplicate detection · **conversion wizard** (account + contact + opportunity, linking existing records) · disqualify |
| **Accounts & contacts** | **Customer 360** (contacts, opportunities, quotes, cases, subsidiaries, activity, audit) · parent accounts · GDPR consent fields |
| **Opportunities** | Drag-and-drop **Kanban pipeline** with configurable stages · automatic probability & forecast category · won/lost with required lost reason · competitors, next step |
| **Quotes & products** | Price list · quote line items (discount and tax per line) · automatic numbering `Q-2026-0001` · branded **print/PDF** with logo and terms · **manager approval for large discounts** · email the quote · accept a quote to close the deal |
| **Demo requests** | Created from leads or manually · **one-click Teams meeting** with Outlook invitations · agenda · no-show tracking · convert to opportunity |
| **Campaigns** | Budget/cost · lead & deal attribution · **ROI** |
| **Cases (service)** | Numbering `CS-2026-0001` · priorities with **SLA deadlines** · breach alerts · resolve/escalate |
| **Activities** | Calls, emails, meetings, notes, tasks on every record · timeline |
| **Insights** | **Forecast** per quarter (closed / commit / best case / pipeline) with **quotas & attainment**, slipping deals · **Reports**: leaderboard, win/loss reasons, funnel, source performance, campaign ROI, quote and service metrics · CSV export everywhere |
| **Productivity** | Global search · quick-create · notifications · saved views · column chooser · bulk edit/assign/delete · CSV import with column mapping · mobile layout · dark mode / section themes · Teams tab & personal app |

## Integrations

| Integration | How |
|---|---|
| **Outlook email** | Send from the user's mailbox via Microsoft Graph (`Mail.Send`), with templates and placeholders; automatically logged to the record |
| **Teams meetings** | Creates an Outlook event with a Teams link (`Calendars.ReadWrite`) and stores the join URL on the demo |
| **Teams channel alerts** | Teams *Workflows* webhook: new leads, demo requests, deals won, cases |
| **AppSource / Partner Center** | HTTPS endpoint → Power Automate → leads list (see below) |
| **Webhooks** | 17 events (`lead.created`, `opportunity.won`, `case.created`, …) to Power Automate, Logic Apps, Zapier, Make, n8n or your own API |
| **Data** | CSV import/export (Excel, HubSpot, Salesforce, Dynamics exports) · full JSON backup |
| **Power Automate flow pack** | 8 ready-made flows as an importable solution (`flows/dist/BrillientaCRMFlows_1_3_1_0.zip`): AppSource lead intake, website leads, demo bookings & support requests (brillienta.com `api/lead.js` → `CRM_WEBHOOK_URL`), new lead alert, daily sales digest, overdue reminders, case SLA escalation, deal-won handoff, quote discount approval. See `flows/README.md` |

## Enterprise features

* **Role-based access** through SharePoint groups: Administrators, Sales managers, Sales, Support, Readers. Setting them up creates the groups and grants list permissions. You can optionally restrict reps to their own records.
* **Audit log** of who changed which field, when, and from what to what. It's shown on every record and on an admin page.
* **Custom fields** for any record type. Admins add them in Settings and the SharePoint columns are created automatically.
* Configurable pipeline stages, pick lists, scoring, SLAs, quotas, branding, currency and fiscal year.
* **Licensing**: a 30-day trial, then signed licence keys tied to the customer's tenant (Standard / Enterprise plans).
* Lists over 5,000 items are paged automatically. Lookup columns are indexed.

## Architecture

```
src/webparts/brillientaCrm/
  BrillientaCrmWebPart.ts       web part, property pane
  models/model.ts               entity metadata: every module is defined here (fields, relations, board field)
  models/settings.ts            tenant settings (stored as JSON in the "CRM Settings" list)
  models/license.ts             ECDSA P-256 licence verification
  services/SpRest.ts            SharePoint REST client
  services/DataService.ts       provisioning, CRUD mapping, audit, roles, permissions
  services/Engine.ts            business rules: numbering, scoring, SLA, forecast, conversion, events, permissions
  services/Integrations.ts      Microsoft Graph (Outlook/Teams) and webhooks
  components/core/              generic grid, record panel, fields, charts, dialogs, quote editor
  components/pages/             dashboard, my work, calendar, forecast, reports, settings, audit, setup
tools/license/                  key generation and licence issuing (vendor-only, not shipped)
```

Every list, form, grid, board, filter, import and export is generated from `models/model.ts`. To add a module, add an entity definition there.

**Lists created by Setup:** `AppSource Leads`, `CRM Accounts`, `CRM Contacts`, `CRM Opportunities`, `CRM Quotes`, `CRM Products`, `CRM Demo Requests`, `CRM Cases`, `CRM Campaigns`, `CRM Activities`, `CRM Settings`, `CRM Audit Log`. You can rename any of them per site in the web part properties. Version 1 data is upgraded in place, because Setup only adds the missing columns.

## Build

Requires Node 22 LTS.

```bash
npm install
npm run build        # → sharepoint/solution/brillienta-crm.sppkg
npm start            # local dev server for the hosted workbench
```

## Install at a customer

1. Upload `brillienta-crm.sppkg` to the **App Catalog** and click **Deploy**.
2. **API access:** in the SharePoint admin center go to *Advanced › API access* and approve the *Microsoft Graph* requests `Mail.Send`, `Calendars.ReadWrite` and `User.ReadBasic.All`. These are needed for Outlook email and Teams meetings.
3. Create a site (for example `/sites/crm`), add a page and add the **Brillienta CRM** web part. A full-width section works best. You can also use *Sync to Teams* to add it as a Teams app.
4. A site owner opens the page and clicks **Set up CRM**.
5. Go to *Admin › Settings*. Set up branding, the pipeline, quotas and integrations. Under *Security*, click **Create groups & apply list permissions** and add users to the groups.
6. **Full screen:** the CRM opens full screen by default, covering the SharePoint header and navigation. Users can switch it with the ⛶ button in the CRM header, and their choice is remembered. Admins can turn it off in the web part properties. While the page is being edited, the CRM never covers the SharePoint editor. For the cleanest result, create the page with **Site contents › New › Page › App**, then choose Brillienta CRM.

## AppSource leads → CRM

In Partner Center, go to your offer › *Customer leads* › **HTTPS endpoint**. Point it at a Power Automate flow: trigger *When an HTTP request is received* → action *SharePoint: Create item* in `AppSource Leads`.

Map the fields as follows:

| Partner Center field | CRM column |
|---|---|
| `UserDetails.FirstName`, `LastName`, `Email`, `Phone`, `Company`, `Country` | same names |
| `UserDetails.Title` | `JobTitle` |
| `OfferTitle` | `Offer` |
| `LeadSource`, `ActionCode`, `Description` | same names |

Also set `Status = New`. The flow must end with a Response of HTTP 200.

## Microsoft Teams

The CRM runs as a **Teams personal app** (left rail) and as a **channel tab**, and also in Outlook and the Microsoft 365 app.

1. **One CRM for everyone:** open the CRM on your CRM site, go to *Settings › Microsoft Teams* and click **Use this site as the default CRM**. Teams tabs run on the team's own site and the personal app runs on the root site, so this setting tells them where the data lives. It's stored as a tenant property (`BrillientaCRM.SiteUrl`).

   This needs app catalog admin rights. The PowerShell alternative is:

   ```powershell
   Set-PnPStorageEntity -Key "BrillientaCRM.SiteUrl" -Value "https://contoso.sharepoint.com/sites/crm" -Scope Tenant
   ```

2. **Publish to Teams:** in the App Catalog, select *Brillienta CRM* › **Sync to Teams**. The solution includes a branded Teams app package (`teams/TeamsSPFxApp.zip`). Alternatively, upload `teams-app/dist/BrillientaCRM-Teams.zip` in the Teams admin center › Manage apps.
3. **Channel tab:** the tab settings show the **CRM site URL** field. Leave it empty to use the default CRM, or point the team at another site, such as a separate CRM for one team.

In Teams, the CRM follows the Teams theme (light, dark or high contrast). It uses an app-style layout with fixed navigation. If pop-ups are blocked, quotes download as an HTML file instead of opening a print window. The CRM site URL can also be set per web part in SharePoint (*CRM data* in the property pane).

To rebuild the Teams package, run `node teams-app/build-teams-app.mjs`. `npm run build` runs it automatically.

## Selling it: licences

**Do this once:**

```bash
node tools/license/generate-keys.mjs --force
```

This creates your private signing key in `tools/license/keys/` (keep it secret and back it up offline). It also writes the matching public key into the solution. Then rebuild with `npm run build`.

**For each customer:** the customer sends you the tenant ID shown in *Settings › Licence*. Issue a key with:

```bash
node tools/license/issue-license.mjs --customer "Contoso Ltd" --tenant <tenant-id> --plan enterprise --seats 50 --expires 2027-12-31
```

The customer pastes the key into *Settings › Licence* and saves.

| Plan | Includes |
|---|---|
| Trial (30 days) | Everything |
| Standard | Leads, accounts, contacts, opportunities, quotes, products, demos, activities, dashboards, reports, calendar, Outlook/Teams, import/export |
| Enterprise | Standard plus cases & SLA, campaigns & ROI, forecast & quotas, webhooks & Teams alerts, custom fields, audit log, security groups |

When a trial or licence lapses, the CRM becomes **read-only**. No data is lost.

The licence check runs in the browser. A determined tenant admin could bypass it, so treat it as commercial protection, not DRM. Your contract is the real enforcement.

## Before you submit to AppSource

* In `config/package-solution.json`, fill in `developer.name`, `websiteUrl`, `privacyUrl`, `termsOfUseUrl` and `mpnId`, and add screenshots.
* In Partner Center, answer *"My product requires purchase of a service or offers in-app purchases"*. Under *Notes for certification*, explain how testers get a licence key, or point out that the 30-day trial works without one.
* Increase `version` in `package-solution.json` for every submission.
