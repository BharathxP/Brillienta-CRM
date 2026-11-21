# Brillienta CRM: enterprise CRM for SharePoint & Microsoft Teams

A complete CRM delivered as one SharePoint Framework (SPFx 1.24) web part. All data is stored in SharePoint lists **in the customer's own Microsoft 365 tenant**. There is no external database or server.

## What's new in 3.5

* **Brillienta CRM in Microsoft 365 Copilot** (SPFx 1.24 Copilot Components). The `.sppkg` now also contains a Copilot agent, **Brillienta CRM**, whose six tools render live CRM data as interactive cards in the Copilot chat:

  | Tool | Example prompt | Shows |
  |---|---|---|
  | `GetMyPipeline` | "Which of my deals are at risk?" | Open deals with stage, amount, close date and AI deal health; total, weighted and at-risk KPIs. Filters: mine / team / all, period, health, minimum amount |
  | `GetAccountSummary` | "Summarise Contoso" | Account 360: open deals, open cases with SLA, subscriptions and ARR, unpaid invoices, activity. Picks from a list when several accounts match |
  | `GetRenewals` | "Which subscriptions renew in the next 60 days?" | Renewals in the window with ARR, seats and auto-renew; past-due pending renewals first |
  | `GetOverdueInvoices` | "Invoices more than 30 days overdue" | Unpaid invoices past due, with open balance and days overdue |
  | `GetSlaRiskCases` | "Which cases are breaching SLA?" | Breached cases, late first replies and cases due soon (Enterprise plan) |
  | `CreateLead` | "Add a lead: Priya Shah, CTO at Contoso, met at the expo" | A prefilled lead form with duplicate check. Nothing is saved until the user presses **Save**; then the normal scoring, round-robin assignment, enrichment, alerts and audit run |

  Every card has **Open** links to the records, and tells Copilot what it shows, so users can ask follow-up questions ("draft an email to the first customer"). Data is read with the user's own permissions through the same engine as the web part: record-level security, "reps see only their own records", licence plan and read-only trial expiry all apply. Nothing is copied out of SharePoint. See [Microsoft 365 Copilot](#microsoft-365-copilot) for setup.

## What's new in 3.4

* **Brillienta Helpdesk integration** (Settings › Service › Brillienta Helpdesk): connect the Helpdesk site and test the connection.
  * Accounts, contacts and cases get a **Support tickets** tab: an account shows the tickets of all its contacts (matched by email), a contact their own; open first, resolved and closed on request; each ticket opens in Helpdesk.
  * **New ticket** from an account, contact or case. The ticket is handed to Helpdesk the same way as Microsoft 365 Copilot tickets: the next Helpdesk page an agent opens numbers, routes and notifies it and starts the SLA. CRM tags (CRM, crm-acc-, crm-con-, crm-case-) link it back to the record, also for customers outside the organisation.
  * Dashboard: **Helpdesk tickets** widget (customers' open tickets, with the account) and KPI tile.
  * Permissions: Helpdesk agents see every ticket; other users only tickets they raised.
* **Settings, easier to use:** plain-language labels and help everywhere; pick lists, holidays, countries and email domains as tags; lead scoring as a points table per rating, source and job title; flow links show *Connected / Not connected* with step-by-step “How do I get this link?”; email and WhatsApp templates collapse to one line and have clickable placeholders; time zone picker and time fields for business hours; friendly empty states; technical details (PowerShell, webhook format, Power BI relationships, admin consent) moved into collapsible notes.

## What's new in 3.3

* **Dashboard redesign:** a compact, modern home page. A coloured header holds the greeting, the *Ask your CRM* box and the AI briefing; below it are KPI tiles with 12-month trend lines and a grid of widgets with fixed heights that scroll inside, so the default layout fits on about one screen.
* **Customisable:** *Customise* lets each user add, remove and reorder widgets and KPI tiles (drag and drop, or the ⋯ menu), and choose each widget's width (⅓, ½, ⅔ or full). The layout is saved per user; *Reset* goes back to the default, and admins can *Save as company default* for everyone.
* **Widgets:** Next best actions, AI forecast, Revenue won with AI projection, Pipeline by stage, Deals at risk, Hottest leads, My open tasks, Biggest open deals, Conversion funnel, Leads by source, Upcoming demos and Open support cases. **KPIs:** Won this quarter, Open pipeline, New leads, Win rate, Average deal, Sales cycle, Demos to schedule, MRR, Unpaid invoices and Open cases.
* **AI** (computed in the browser from the records the user can see; nothing is sent anywhere):
  * *AI briefing*: quarter outlook against quota, deals at risk and the main reason, lead flow against the same point last quarter, leads waiting for first contact, best converting source, SLA breaches, overdue invoices and renewals.
  * *Deal health*: each open deal gets a win likelihood (0–100) from its stage, activity, close date, quotes and e-signature, contact, next steps, age and the company's own win history; hover the badge for the reasons.
  * *AI forecast*: won, commit, forecast (deals × win likelihood) and best case for the quarter against quota.
  * *Next best actions*: a ranked to-do list across leads, deals, quotes, cases, invoices, renewals and tasks, with a one-click email draft.
  * *Ask your CRM*: plain-language questions such as "deals closing this month over 5k", "hot leads from LinkedIn" or "will we hit quota this quarter?".
  * *Ask Copilot*: copies a summary of the dashboard as a prompt and opens Microsoft 365 Copilot.

## What's new in 3.0

* **Multi-currency and price books:** each deal, quote, invoice and subscription has its own currency, while dashboards, forecasts, reports and grid totals convert to the company currency (Settings › Company › Currencies). Price books (for example "Partners EUR") set a quote's currency and prices, with specific prices per product or a discount on the converted list price.
* **Territories and teams:** teams have a manager, members and countries. New leads are routed round-robin within their territory team, including leads arriving from flows. The Forecast can be filtered by team and has a per-team rollup, and the Dashboard has a "Team" scope.
* **Record-level security in SharePoint:** optionally, every lead, account, contact, deal, quote, case, subscription and invoice gets its own permissions. Access goes to the owner, the owner's team manager, team members (read or edit), and the CRM groups, so reps can't reach other reps' records anywhere in Microsoft 365. Administrators apply it in one click; the *Record security* flow keeps it up to date.
* **Offline:** lists are kept on the device. With no connection, the CRM still opens from the saved session, shows the saved data, and queues edits and new records (including follow-up activities). These are sent in order when the connection returns, and new records get their real IDs. Deleting is disabled while offline.
* **Languages:** English, Deutsch, Français, Español and हिन्दी, chosen per user (the default follows the SharePoint language). Navigation, record types, fields, sections, choices, statuses, commands and buttons are translated. Dashboards, settings and messages are still in English.
* **Lead enrichment:** new leads with a business email get their website and company name at once. The *Lead / Account enrichment* flows then read the company website (name, description, LinkedIn page) and, with an Apollo.io key, add industry, size, location and revenue.
* **Customer portal:** customers sign in with an email link and see their company's requests. They can follow and reply to the conversation (which reopens resolved cases), open new requests, and search published knowledge articles. Contacts get an "Invite to portal" command.
* **Outlook add-in** (`outlook-addin/`): shows the sender's contact or lead with open deals and cases, logs the email to their timeline (no duplicates with Outlook sync), and creates a lead for unknown senders. It's hosted at brillienta.com/crm-addin.
* **AppSource transactable SaaS** (BrillientaWebSite): a landing page and fulfillment webhook that activate subscriptions, issue and email tenant-bound licence keys, handle plan and seat changes and renewals, and record each sale in Brillienta's CRM.

Upgrading from 2.x: deploy the `.sppkg`, click **Update lists** (adds the new columns), then re-import the flow pack (2.0, 23 flows). Then paste the *Customer portal* flow URL into the flow's own "Customer portal URL" variable and into Settings › Company.

## What's new in 2.2

* **Sales sequences:** timed emails, calls, WhatsApp messages and tasks for leads and contacts. Email steps are sent automatically or with one click, and a reply stops the sequence.
* **Duplicate merge** for leads, contacts and accounts: choose which values to keep, and every deal, quote, case and activity moves over.
* **Knowledge base**, with matching articles suggested on each case, "reply with this", and usage counts.
* **Customer satisfaction:** a one-click 1–5 survey when a case is resolved, low-score alerts, and CSAT in Reports.
* **Click-to-call** (Teams or phone) with a call timer, outcome, notes and a follow-up task.
* **GDPR:** personal-data export (access request), right to erasure, and a consent history on the timeline.
* **Email marketing:** campaign emails to consented leads and contacts with a personal unsubscribe page, or export the audience to Mailchimp / Brevo.
* **Inbound REST API** (create / upsert / get) for other systems, plus a **Power BI** connection guide.

Upgrading from 2.1: deploy the `.sppkg`, click **Update lists** (adds the `CRM Knowledge Base` list and new columns), then re-import the flow pack (1.5, 20 flows).

## What's new in 2.1

* **Outlook sync:** email and meetings with customers are logged automatically.
* **Subscriptions, renewals and invoicing:** MRR/ARR, renewal opportunities, invoices, Razorpay/Stripe pay links, and Zoho Books / QuickBooks / Xero / Tally export.
* **Service:** email-to-case, plus first-response and resolution SLAs in business hours, with pause while waiting on the customer.
* **E-signature on quotes**, completing quote-to-cash.
* **AI:** a smart summary, risks, next best actions and follow-up drafts on every record, Ask Copilot, and a Microsoft 365 Copilot agent.
* **WhatsApp Business:** click to chat, or the Cloud API with templates, replies and delivery status.

Upgrading from 2.0: deploy the new `.sppkg`, open the CRM as an administrator and click **Update lists**. This adds 2 lists and the new columns, and existing data is kept. Then download the flow pack again from Settings › Integrations and import it over the old one.

## Modules

| Area | Features |
|---|---|
| **Home** | Role-aware dashboard (quarter KPIs with change vs. last quarter, 12-month revenue chart, pipeline by stage, conversion funnel, lead sources, hottest leads, biggest deals, demos, tasks) · **My work** queue (overdue/today/upcoming tasks, follow-ups, new leads, demos, deals closing, cases at risk) · **Calendar** (demos, tasks, expected closes, SLA deadlines) |
| **Leads** | AppSource/Partner Center, web form, import, manual entry · automatic **lead scoring** (0–100, configurable rules) · **round-robin assignment** · duplicate detection · **conversion wizard** (account + contact + opportunity, linking existing records) · disqualify |
| **Accounts & contacts** | **Customer 360** (contacts, opportunities, quotes, cases, subsidiaries, activity, audit) · parent accounts · GDPR consent fields |
| **Opportunities** | Drag-and-drop **Kanban pipeline** with configurable stages · automatic probability & forecast category · won/lost with required lost reason · competitors, next step |
| **Quotes & products** | Price list · quote line items (discount and tax per line) · automatic numbering `Q-2026-0001` · branded **print/PDF** with logo and terms · **manager approval for large discounts** · email the quote · **e-signature**: a single-use signing link; the customer signs online, and name, job title, IP address, time and a snapshot of the signed quote are recorded. A signed quote closes the deal and creates its subscriptions automatically |
| **Subscriptions & invoicing** | Subscriptions with seats, billing cycle, **MRR/ARR**, renewal date and auto-renew, created from a quote's recurring lines · **automatic renewal opportunities** N days before annual renewals (a won renewal extends the term; a lost one cancels it) · monthly roll-over and expiry · **invoices** `INV-2026-0001` from quotes or subscriptions, with payment terms, partial payments, overdue tracking, branded print/PDF and bank details · **Razorpay / Stripe payment links** with automatic "paid" updates · **accounting export** for Zoho Books, QuickBooks Online, Xero, Tally (XML vouchers) and generic CSV |
| **Demo requests** | Created from leads or manually · **one-click Teams meeting** with Outlook invitations · agenda · no-show tracking · convert to opportunity |
| **Campaigns** | Budget/cost · lead & deal attribution · **ROI** · **email campaigns** to consented leads and contacts (filters by tag, source and country; personal unsubscribe link; company footer; Outlook-safe sending pace; up to 500 per run) · audience export for Mailchimp / Brevo |
| **Sequences** | Multi-step cadences (email, call, WhatsApp, task on day N). Bulk-enrol from the grid. Due steps run when the owner opens the CRM. Email steps go out automatically or become one-click tasks. A reply by email or WhatsApp stops the sequence |
| **Knowledge base** | Articles with category, product, keywords and a public link. On each case the best matches are suggested (published only); "Reply with this" puts the answer into the email and links the article. Usage counts show which answers help most |
| **Cases (service)** | **Email-to-case** from a shared support mailbox: replies with `[CS-2026-0042]` in the subject thread onto the case and reopen it, and new senders get an automatic acknowledgement · numbering `CS-YEAR-ID` · **first-response and resolution SLAs** per priority, counted in **business hours** (time zone, working days, holidays) · the SLA clock pauses while *Waiting on customer* · breach alerts · reply, resolve, escalate |
| **Activities** | Calls (**click-to-call with Teams or phone**, timer, outcome, follow-up), emails, meetings, WhatsApp, notes, tasks on every record · timeline · **automatic Outlook logging** of email and meetings with known contacts and leads |
| **AI assistance** | A **smart summary** on every record (situation, risks, next best actions) · **follow-up email drafts** · an **Ask Copilot** hand-off to Microsoft 365 Copilot · a downloadable **Microsoft 365 Copilot agent** grounded in the CRM site |
| **Insights** | **Forecast** per quarter (closed / commit / best case / pipeline) with **quotas & attainment**, slipping deals · **Reports**: leaderboard, win/loss reasons, funnel, source performance, campaign ROI, quote and service metrics · CSV export everywhere |
| **Productivity** | Global search · quick-create · notifications · saved views · column chooser · bulk edit/assign/delete · CSV import with column mapping · mobile layout · dark mode / section themes · Teams tab & personal app |

## Integrations

| Integration | How |
|---|---|
| **Outlook email** | Send from the user's mailbox via Microsoft Graph (`Mail.Send`), with templates and placeholders; automatically logged to the record |
| **Outlook sync** | Per-user opt-in (My work › Turn on). Every 15 minutes while the CRM is open, it reads the user's inbox, sent items and calendar (`Mail.Read`, `Calendars.ReadWrite`) and logs conversations and meetings with known contacts and leads on their timelines, with a link back to Outlook. Internal mail is never logged, and items are de-duplicated |
| **WhatsApp Business** | **Click to chat** (free; opens WhatsApp with the message and a template filled in) or the **WhatsApp Business Cloud API** through Power Automate, which sends approved templates and free text and receives replies, with delivery and read status. Unknown numbers become WhatsApp leads. Everything is logged on the timeline |
| **Microsoft 365 Copilot** | **Built-in agent** (3.5): six SPFx Copilot Components ship in the `.sppkg` and render live pipeline, account, renewal, invoice, case and new-lead cards in Copilot. See [Microsoft 365 Copilot](#microsoft-365-copilot). **Knowledge agent**: Settings › Integrations › **Download Copilot agent** builds a declarative agent (Teams manifest 1.19) grounded in the CRM site, for text answers across all lists. Upload it in the Teams admin center. Both need Microsoft 365 Copilot licences |
| **Inbound REST API** | Power Automate flow *Inbound API*: `POST {entity, action: create/upsert/get, match: {field, value}, fields}` with an `x-api-key` header, for ERPs, billing systems, Zapier or Make |
| **Power BI** | Settings › Integrations › Power BI: connect Power BI Desktop to the CRM lists (SharePoint Online list 2.0), with the relationships documented. Data stays protected by SharePoint permissions |
| **Payments** | Razorpay and Stripe payment links and payment webhooks through Power Automate. API keys never reach the browser |
| **Teams meetings** | Creates an Outlook event with a Teams link (`Calendars.ReadWrite`) and stores the join URL on the demo |
| **Teams channel alerts** | Teams *Workflows* webhook: new leads, demo requests, deals won, cases |
| **AppSource / Partner Center** | HTTPS endpoint → Power Automate → leads list (see below) |
| **Webhooks** | 27 events (`lead.created`, `opportunity.won`, `case.created`, …) to Power Automate, Logic Apps, Zapier, Make, n8n or your own API |
| **Data** | CSV import/export (Excel, HubSpot, Salesforce, Dynamics exports) · full JSON backup |
| **Power Automate flow pack** | 23 ready-made flows as an importable solution (`flows/dist/BrillientaCRMFlows_2_0_1_0.zip`, also downloadable from Settings › Integrations): AppSource lead intake, website leads/demos/support, new lead alert, daily digest, overdue reminders, case SLA escalation, deal-won handoff, quote discount approval, **quote e-signature**, **email to case**, **Razorpay/Stripe payment link and payment received**, **WhatsApp send/inbound**, **case satisfaction survey**, **unsubscribe**, **inbound API**, **record security**, **lead/account enrichment**, **customer portal**. See `flows/README.md` |
| **Outlook add-in** | `outlook-addin/`: sender lookup, track email, create lead, open deals and cases, in Outlook desktop, web and mobile. See `outlook-addin/README.md` |

## Enterprise features

* **Role-based access** through SharePoint groups: Administrators, Sales managers, Sales, Support, Readers. Setting them up creates the groups and grants list permissions. You can optionally restrict reps to their own records.
* **Audit log** of who changed which field, when, and from what to what. It's shown on every record and on an admin page.
* **Duplicate merge**, with references and timelines moved to the surviving record.
* **Privacy / GDPR:** export everything held about a person as JSON, erase personal data (record, timeline, copies on deals, demos, cases and quotes, plus the audit history), unsubscribe flags, and consent changes logged automatically.
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
  services/OutlookSync.ts       automatic Outlook email & meeting logging
  services/BusinessHours.ts     business-hours SLA arithmetic (IANA time zones, holidays)
  services/Accounting.ts        invoice exports: Zoho Books, QuickBooks Online, Xero, Tally, CSV
  services/Insights.ts          smart summary, risks, next best actions, follow-up drafts, Copilot prompt
  services/CopilotAgent.ts      Microsoft 365 Copilot declarative agent package
  services/WhatsApp.ts          WhatsApp numbers, click-to-chat, Cloud API queue
  services/FlowPackage.ts       Power Automate solution zip, built in the browser
  services/OfflineStore.ts      IndexedDB copy of lists, offline edit queue
  services/Enrichment.ts        lead/account enrichment (domain → website/company)
  services/Marketing.ts         email campaigns, unsubscribe links, audience export
  i18n/                         UI languages (index.ts) and translation table (table.ts)
src/copilotComponents/          Microsoft 365 Copilot tools (SPFx 1.24 Copilot Components)
  shared/crmSession.ts          finds the CRM site (tenant property) and opens a CRM Engine session as the user
  shared/crmQueries.ts          what each tool shows (pure functions over the Engine; unit-tested)
  shared/CrmCopilotComponentBase.tsx  React root, Copilot theme, links, model context
  crmPipeline/ crmAccountSummary/ crmRenewals/ crmOverdueInvoices/ crmSlaRiskCases/ crmCreateLead/
                                one folder per tool: manifest (tool name + description), parameter schema, view
copilot/                        the Copilot agent: Teams manifest, declarative agent, instructions, plugin seed
config/copilot-agent.json       which components the agent exposes (the build generates teams/brillienta-crm.zip)
outlook-addin/                  Outlook add-in (manifest.xml + web/ task pane)
  components/core/              generic grid, record panel, fields, charts, dialogs, quote editor
  components/pages/             dashboard, my work, calendar, forecast, reports, settings, audit, setup
tools/license/                  key generation and licence issuing (vendor-only, not shipped)
```

Every list, form, grid, board, filter, import and export is generated from `models/model.ts`. To add a module, add an entity definition there.

**Lists created by Setup:** `AppSource Leads`, `CRM Accounts`, `CRM Contacts`, `CRM Opportunities`, `CRM Quotes`, `CRM Products`, `CRM Demo Requests`, `CRM Cases`, `CRM Campaigns`, `CRM Subscriptions`, `CRM Invoices`, `CRM Knowledge Base`, `CRM Activities`, `CRM Settings`, `CRM Audit Log`. You can rename any of them per site in the web part properties. Version 1 data is upgraded in place, because Setup only adds the missing columns.

## Build

Requires Node 22 LTS.

```bash
npm install
npm run build        # → sharepoint/solution/brillienta-crm.sppkg
npm start            # local dev server for the hosted workbench
```

## Install at a customer

1. Upload `brillienta-crm.sppkg` to the **App Catalog** and click **Deploy**.
2. **API access:** in the SharePoint admin center go to *Advanced › API access* and approve the *Microsoft Graph* requests `Mail.Send`, `Mail.Read`, `Calendars.ReadWrite` and `User.ReadBasic.All`. These are needed for Outlook email, Outlook sync and Teams meetings.
3. Create a site (for example `/sites/crm`), add a page and add the **Brillienta CRM** web part. A full-width section works best. You can also use *Sync to Teams* to add it as a Teams app.
4. A site owner opens the page and clicks **Set up CRM**.
5. Go to *Admin › Settings*. Set up branding, the pipeline, quotas and integrations. Under *Security*, click **Create groups & apply list permissions** and add users to the groups.
6. **Full screen:** the CRM opens full screen by default, covering the SharePoint header and navigation. Users can switch it with the ⛶ button in the CRM header, and their choice is remembered. Admins can turn it off in the web part properties. While the page is being edited, the CRM never covers the SharePoint editor. For the cleanest result, create the page with **Site contents › New › Page › App**, then choose Brillienta CRM.

## Microsoft 365 Copilot

Version 3.5 ships a Copilot agent inside the `.sppkg` (SPFx 1.24 Copilot Components). There is no separate app to upload and no API key: the tools call SharePoint as the signed-in user.

**Requirements**

* SharePoint Framework 1.24 Copilot Components must be available in the tenant (rolling out; Targeted release tenants get it first).
* Users need a Microsoft 365 Copilot licence.
* Nothing else to configure. Copilot does not run on the CRM site, so it finds the CRM the same way as the Teams personal app: with SharePoint Search, for the signed-in person (no admin rights or PowerShell). A SharePoint admin can optionally pin one CRM site when the tenant has several (see [Microsoft Teams](#microsoft-teams)).

**Install**

1. Deploy `brillienta-crm.sppkg` to the App Catalog as usual. The package contains the six Copilot Components and the agent package (`ClientSideAssets/brillienta-crm.zip`).
2. Make the agent available in the Microsoft 365 admin center (*Agents* / *Integrated apps*) if your tenant requires admin approval for agents.
3. In Copilot, open the **Brillienta CRM** agent and try a conversation starter, such as "Show my open deals closing this quarter."

**One Teams app for the CRM and the agent.** In SPFx 1.24, when a solution has both Copilot Components and a Teams-enabled web part, *Sync to Teams* publishes only the agent's Teams app ([sp-dev-docs#11041](https://github.com/SharePoint/sp-dev-docs/issues/11041)). So the agent's app (`copilot/manifest.json`) also carries the CRM's personal tab and channel tab: deploying the `.sppkg` and choosing **Sync to Teams** gives users the CRM app (left rail and channel tabs) and the Copilot agent in one app. Nothing else to upload. `toolSchemas.test.ts` checks that the tabs and the agent stay in that manifest.

**How it behaves**

* Each tool reads only the lists it needs (for example the pipeline reads opportunities, accounts, activities and quotes) once per Copilot session.
* Users see exactly what they can see in the CRM. If they have no access to the CRM site, or search does not find it yet, the card says so.
* `CreateLead` never saves by itself. Copilot prefills the form; the user reviews it and presses **Save**. The save goes through the same rules as the web part (permissions, read-only licence, scoring, assignment, enrichment, webhooks, audit).
* "Open" links go to the record's SharePoint item page.

**Change or add a tool**

Tools live in `src/copilotComponents/<tool>/`. The tool name and the description Copilot uses to choose it are in the `*.manifest.json`, the parameters are in `*Properties.ts` (plain JSON Schema, portable plugin v2.4 subset, no imports), and the agent's routing rules are in `copilot/instruction.txt`. New components also need an entry in `config/config.json` (bundle `brillienta-crm-copilot`) and `config/copilot-agent.json`; `toolSchemas.test.ts` checks this. Microsoft recommends creating new components with `yo @microsoft/sharepoint --component-type copilotComponent`.

**Not yet verified in a live tenant:** the tools are unit-tested and render-tested with sample data, and the package builds, but routing by Copilot, SharePoint calls from the Copilot host, and the Copilot iframe's content security policy have to be checked in a tenant where SPFx Copilot Components are enabled.

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

1. **One CRM for everyone, found automatically:** Teams tabs run on the team's own site and the personal app runs on the root site, so the CRM finds its data by itself, in this order:
   * the site the tab runs on, if it has the CRM lists;
   * otherwise SharePoint Search finds the `CRM Opportunities` list the person can read. Nothing to set up and no admin rights or PowerShell needed. A brand-new CRM site appears in search within minutes to a few hours.

   *Settings › Microsoft Teams › Main CRM site* shows whether search already finds the CRM. **Optional:** when a tenant has more than one CRM, a SharePoint admin can pin one for everyone with **Pin this site** (needs app catalog admin rights) or `Set-PnPStorageEntity -Key "BrillientaCRM.SiteUrl" -Value "https://contoso.sharepoint.com/sites/crm" -Scope Tenant`. A pinned site always wins.
2. **Publish to Teams:** in the App Catalog, select *Brillienta CRM* › **Sync to Teams**. From 3.5 this one Teams app contains the CRM tabs and the Copilot agent (see [Microsoft 365 Copilot](#microsoft-365-copilot)). The `.sppkg` contains only this one Teams app (no `teams/TeamsSPFxApp.zip`), because two Teams apps for the same web part make Sync to Teams fail with "solution already exists".
3. **Channel tab:** the tab settings show the **CRM site URL** field. Leave it empty to find the CRM automatically, or point the team at another site, such as a separate CRM for one team.

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
