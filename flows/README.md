# Brillienta CRM: Power Automate flow pack

`dist/BrillientaCRMFlows_1_3_1_0.zip` is an **unmanaged Power Platform solution** containing 8 ready-made cloud flows, 3 connection references and 7 environment variables. To rebuild it, run `node flows/build-solution.mjs`. The build checks every flow before zipping.

## The flows

| Flow | Trigger | What it does | Licence |
|---|---|---|---|
| **AppSource lead intake** | HTTP request (Partner Center → *Customer leads* → HTTPS endpoint) | Creates the lead in `AppSource Leads` (source AppSource or Azure Marketplace, status New, raw payload kept) and always answers HTTP 200 | **Premium** |
| **Website leads & demo bookings** | brillienta.com `api/lead.js` sends every validated lead (`CRM_WEBHOOK_URL`) | **Demo booking:** Hot lead + demo request (products, requested slot in UTC and the visitor's time zone, length) + email to sales. **Support request** (Issue, Feature request, How-to, Licence & billing, Security): **case** with the `BR-…` reference, priority from the visitor's impact, SLA deadline, linked to the contact if the email matches, + email to support. **Other enquiry:** lead. Keeps UTM, landing page and referrer. Checks the `x-lead-key` header | **Premium** |
| **New lead alert** | New item in `AppSource Leads` | Emails the lead owner. Unassigned leads go to the sales manager | Standard |
| **Daily sales digest** | Weekdays 07:00 UTC | New leads, deals won (last 24 h), deals closing in 7 days and overdue tasks, sent to the sales manager | Standard |
| **Overdue reminders** | Weekdays 06:00 UTC | **One email per person** listing their overdue tasks and lead follow-ups | Standard |
| **Case SLA escalation** | Every hour | Cases that breached their SLA in the last hour go to the support manager, with the case owner in Cc (high importance) | Standard |
| **Deal won handoff** | Opportunity changes to *Closed Won* | Emails finance the deal details (owner in Cc) and creates `Customers/<customer name>` in a document library | Standard |
| **Quote discount approval** | Quote's *Approval* changes to *Pending* (the rep clicks **Request approval** in the CRM) | Sends an **Approvals** request to the sales manager in Teams and Outlook. Writes *Approved* / *Rejected* and the comments back to the quote, and emails the rep | Standard |

The CRM enforces the approval step. A quote whose total discount is above **Settings › Company › Quotes › "Discount needing manager approval"** (default 15%) can't be emailed or marked accepted until the flow approves it. If the total changes after approval, the quote needs approving again.

## Install at a customer

1. **Import:** go to make.powerautomate.com (choose the right environment) › **Solutions › Import solution** and select the zip.
2. **Connections:** the wizard asks for three: **SharePoint**, **Office 365 Outlook** and **Approvals**. Sign in with a service account if possible. That account sends the emails, so it needs edit access to the CRM lists.
3. **Environment variables:** fill these in during import, or later under Solutions › Brillienta CRM Flows › Environment variables:

   | Variable | Example |
   |---|---|
   | CRM site URL | `https://contoso.sharepoint.com/sites/crm` |
   | CRM page URL | `https://contoso.sharepoint.com/sites/crm/SitePages/CRM.aspx` |
   | Sales manager email | `sales.manager@contoso.com` |
   | Finance email | `finance@contoso.com` |
   | Support manager email | `support@contoso.com` |
   | Website shared key | same as `CRM_WEBHOOK_KEY` in Vercel, or `off` |
   | Customer folder library | `Documents` |

4. **Turn on** the flows you want. They're imported switched **off**, so nothing runs before it's configured.
5. **Website (brillienta.com):** open *Website leads and demo bookings*, copy its **HTTP URL**, and add it in Vercel › Settings › Environment Variables as `CRM_WEBHOOK_URL`. Optionally add `CRM_WEBHOOK_KEY` and set the flow's **Website shared key** variable to the same value. Then redeploy. This is already supported in the website code (`api/lead.js`). It runs alongside your existing `LEAD_WEBHOOK_URL` flow and email, and never replaces them.
6. **AppSource only:** open *AppSource lead intake*, copy the **HTTP URL** from the trigger, paste it in Partner Center (*Customer leads › HTTPS endpoint*) and click **Validate**.

## Notes and requirements

* **Schedules are in UTC.** To use local time, open the flow, edit the *Recurrence* trigger and set your time zone.
* **Deal won handoff** uses *Get changes for an item* to fire only at the moment the stage *changes* to Closed Won. That needs **version history** on `CRM Opportunities`, which SharePoint lists have on by default.
* **Renamed lists:** the flows use the default list names (`AppSource Leads`, `CRM Opportunities`, …). If a site renamed a list in the web part properties, change the list name in the affected flow actions.
* **Licensing:** only the two HTTP-triggered flows (AppSource lead intake and Website leads & demo bookings) need **Power Automate Premium**. Everything else runs on standard Microsoft 365 licences.
* **Other events:** for anything the pack doesn't cover, use **Settings › Integrations › Outgoing webhooks** in the CRM. It offers 17 events, including `opportunity.won`, `case.created` and `quote.approvalRequested`, sent to any flow's HTTP trigger.

## If the import fails

The solution was generated by a script and checked structurally, but it hasn't been imported into a live Power Platform environment yet. If an environment rejects it, every flow's full definition is in `dist/solution/Workflows/*.json`. Each action there maps one-to-one to a step in the flow designer, so the flows can be rebuilt by hand. Please send me the import error and I'll fix the generator.
