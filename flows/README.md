# Brillienta CRM: Power Automate flow pack

`dist/BrillientaCRMFlows_1_4_0_0.zip` is an **unmanaged Power Platform solution** containing 16 ready-made cloud flows, 4 connection references and 15 environment variables. Admins can also download it from the CRM (**Settings › Integrations › Download Power Automate solution**), already filled in for their site. To rebuild it, run `node flows/build-solution.mjs`. The build checks every flow before zipping.

## The flows

| Flow | Trigger | What it does | Licence |
|---|---|---|---|
| **AppSource lead intake** | HTTP request (Partner Center → *Customer leads* → HTTPS endpoint) | Creates the lead in `AppSource Leads` (source AppSource or Azure Marketplace, status New, raw payload kept) and always answers HTTP 200 | **Premium** |
| **Website leads & demo bookings** | brillienta.com `api/lead.js` sends every validated lead (`CRM_WEBHOOK_URL`) | **Demo booking:** Hot lead + demo request (products, requested slot in UTC and the visitor's time zone, length) + email to sales. **Support request** (Issue, Feature request, How-to, Licence & billing, Security): **case** with the `BR-…` reference, priority from the visitor's impact (the CRM sets the business-hours SLA deadlines), linked to the contact if the email matches, + email to support. **Other enquiry:** lead. Keeps UTM, landing page and referrer. Checks the `x-lead-key` header | **Premium** |
| **New lead alert** | New item in `AppSource Leads` | Emails the lead owner. Unassigned leads go to the sales manager | Standard |
| **Daily sales digest** | Weekdays 07:00 UTC | New leads, deals won (last 24 h), deals closing in 7 days and overdue tasks, sent to the sales manager | Standard |
| **Overdue reminders** | Weekdays 06:00 UTC | **One email per person** listing their overdue tasks and lead follow-ups | Standard |
| **Case SLA escalation** | Every hour | Cases that breached their SLA in the last hour go to the support manager, with the case owner in Cc (high importance) | Standard |
| **Deal won handoff** | Opportunity changes to *Closed Won* | Emails finance the deal details (owner in Cc) and creates `Customers/<customer name>` in a document library | Standard |
| **Quote e-signature** | HTTP request (the customer opens the signing link) | Serves the **signing page**: the quote, its lines, total and terms, plus name, job title, comments, an "I agree" box, and Accept or Decline. On *Accept* it records the signer, title, IP address, time, a reference and a snapshot of the signed quote, sets the quote to *Accepted*, makes the link single-use, and emails the quote owner and the signer. The CRM then closes the deal and creates subscriptions | **Premium** |
| **Email to case** | New email in the support shared mailbox | A subject with `[CS-2026-0042]` (or `[BR-…]`) logs the email on that case and reopens it if it was resolved or waiting. Any other email creates a case (contact and account matched by sender, numbered `PREFIX-YEAR-ID`) and sends an automatic acknowledgement with the case number. Auto-replies and no-reply senders are ignored | Standard |
| **Razorpay payment link** | Invoice: *Create Razorpay link* clicked | Creates a Razorpay Payment Link for the balance (reference = invoice number), Razorpay emails it to the billing contact, and the short URL is written back to the invoice | **Premium** |
| **Razorpay payment received** | Razorpay webhook `payment_link.paid` | Re-reads the payment link from the Razorpay API (so forged calls are ignored), then marks the invoice *Paid* and logs the payment | **Premium** |
| **Stripe payment link** | Invoice: *Create Stripe link* clicked | Creates a Stripe Price and Payment Link for the balance and writes the URL back to the invoice | **Premium** |
| **Stripe payment received** | Stripe webhook `checkout.session.completed` | Re-reads the Checkout Session from the Stripe API, then marks the invoice *Paid* | **Premium** |
| **WhatsApp send** | Activity with WhatsApp status *Queued* (CRM › WhatsApp › Send) | Sends an approved template (with its values) or free text through the WhatsApp Business Cloud API, and writes *Sent* / *Failed* back | **Premium** |
| **WhatsApp inbound** | Meta webhook | Answers Meta's verification, logs incoming messages on the contact or lead whose number matches (unknown numbers become WhatsApp leads), and updates the delivery and read status of sent messages | **Premium** |
| **Quote discount approval** | Quote's *Approval* changes to *Pending* (the rep clicks **Request approval** in the CRM) | Sends an **Approvals** request to the sales manager in Teams and Outlook. Writes *Approved* / *Rejected* and the comments back to the quote, and emails the rep | Standard |

The CRM enforces the approval step. A quote whose total discount is above **Settings › Company › Quotes › "Discount needing manager approval"** (default 15%) can't be emailed or marked accepted until the flow approves it. If the total changes after approval, the quote needs approving again.

## Install at a customer

1. **Import:** go to make.powerautomate.com (choose the right environment) › **Solutions › Import solution** and select the zip.
2. **Connections:** the wizard asks for four: **SharePoint**, **Office 365 Outlook**, **Approvals** and **Content Conversion** (HTML to text for email-to-case). Sign in with a service account if possible. That account sends the emails, so it needs edit access to the CRM lists.
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
   | Support mailbox | `support@contoso.com` (shared mailbox for email-to-case; the connection account needs Full Access and Send As) |
   | Case number prefix | `CS` (same as in the CRM settings) |
   | Razorpay key id / key secret | from Razorpay › API keys, or leave `not-set` |
   | Stripe secret key | restricted key `rk_…`, or leave `not-set` |
   | WhatsApp phone number ID / access token / webhook verify token | from Meta › WhatsApp › API setup, or leave `not-set` |

   The payment and WhatsApp keys are **not** asked for in the CRM's download dialog, so they never pass through the browser. Enter them here, and use Azure Key Vault secret environment variables if your organisation requires that.

4. **Turn on** the flows you want. They're imported switched **off**, so nothing runs before it's configured.
5. **Website (brillienta.com):** open *Website leads and demo bookings*, copy its **HTTP URL**, and add it in Vercel › Settings › Environment Variables as `CRM_WEBHOOK_URL`. Optionally add `CRM_WEBHOOK_KEY` and set the flow's **Website shared key** variable to the same value. Then redeploy. This is already supported in the website code (`api/lead.js`). It runs alongside your existing `LEAD_WEBHOOK_URL` flow and email, and never replaces them.
6. **E-signature:** turn on *Quote e-signature*, copy its **HTTP URL** and paste it in the CRM under **Settings › Company › E-signature link**. Reps then see **Send for e-signature** on approved quotes.
7. **Payments:** turn on the Razorpay or Stripe flows, choose the provider in **Settings › Company › Online payment links**, and add the *payment received* flow's HTTP URL as a webhook: Razorpay › Webhooks (`payment_link.paid`) or Stripe › Developers › Webhooks (`checkout.session.completed`).
8. **WhatsApp Cloud API:** create the approved templates in WhatsApp Manager (the defaults are `crm_intro`, `crm_demo_reminder`, `crm_quote_followup` and `crm_payment_reminder`, or rename them in the CRM). Turn on both WhatsApp flows, set Meta › WhatsApp › Configuration › Webhook to the *WhatsApp inbound* HTTP URL with your verify token, subscribe to **messages**, and set **Settings › Integrations › WhatsApp › How messages are sent** to Cloud API.
9. **AppSource only:** open *AppSource lead intake*, copy the **HTTP URL** from the trigger, paste it in Partner Center (*Customer leads › HTTPS endpoint*) and click **Validate**.

## Notes and requirements

* **Schedules are in UTC.** To use local time, open the flow, edit the *Recurrence* trigger and set your time zone.
* **Deal won handoff** uses *Get changes for an item* to fire only at the moment the stage *changes* to Closed Won. That needs **version history** on `CRM Opportunities`, which SharePoint lists have on by default.
* **Renamed lists:** the flows use the default list names (`AppSource Leads`, `CRM Opportunities`, …). If a site renamed a list in the web part properties, change the list name in the affected flow actions.
* **Licensing:** flows that use HTTP (AppSource, website, e-signature, payments, WhatsApp) need **Power Automate Premium** for the connection owner. Everything else runs on standard Microsoft 365 licences.
* **Webhook security:** Power Automate can't check HMAC signatures, so the payment flows **re-read the payment from the provider's API** before marking an invoice paid. WhatsApp inbound only accepts events for your phone number ID. The e-signature link is a random 48-character, single-use token that expires with the quote.
* **SLAs:** cases created by flows get their SLA deadlines from the CRM, in business hours, the next time someone in support or a manager opens it.
* **Amounts** for Razorpay and Stripe are sent in the smallest currency unit (× 100). Zero-decimal currencies such as JPY need the multiplier changed in the flow.
* **Other events:** for anything the pack doesn't cover, use **Settings › Integrations › Outgoing webhooks** in the CRM. It offers 27 events, including `opportunity.won`, `quote.signed`, `invoice.paid`, `subscription.renewed` and `case.created`, sent to any flow's HTTP trigger.

## If the import fails

The solution was generated by a script and checked structurally, but it hasn't been imported into a live Power Platform environment yet. If an environment rejects it, every flow's full definition is in `dist/solution/Workflows/*.json`. Each action there maps one-to-one to a step in the flow designer, so the flows can be rebuilt by hand. Please send me the import error and I'll fix the generator.
