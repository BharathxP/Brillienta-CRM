# Brillienta CRM for Outlook

An Outlook add-in (desktop, web, new Outlook and mobile) that works with a customer's own Brillienta CRM site:

* shows the sender as a CRM **contact or lead**, with job title, phone, **open deals** and **open cases**
* **Track email**: saves the email on the contact's or lead's timeline (same ID as Outlook sync, so it's never logged twice)
* **Create lead and track email** for unknown senders. Company and website come from a business email domain, and the lead is queued for enrichment.
* **Open CRM** goes to the CRM page

All data stays in the customer's SharePoint. The add-in calls SharePoint REST with the signed-in user's own permissions (Microsoft Entra nested app authentication, with a popup as fallback).

## Files

| File | Purpose |
|---|---|
| `manifest.xml` | Add-in manifest (Mailbox 1.5, read mode, desktop + mobile command surfaces) |
| `web/taskpane.html`, `taskpane.js`, `taskpane.css` | Task pane (vanilla JS, no build step) |
| `web/config.js` | Client ID of Brillienta's multi-tenant app registration |
| `web/msal-browser.min.js` | MSAL.js 3.28.1 (MIT, `msal-browser.LICENSE`), vendored so no third-party CDN is needed |
| `web/icon-*.png` | Icons |

## Publisher set-up (Brillienta, once)

1. **App registration** in Brillienta's Microsoft Entra tenant. Make it multi-tenant, with these SPA redirect URIs: `brk-multihub://brillienta.com`, `https://brillienta.com/crm-addin/taskpane.html` and `https://brillienta.com/marketplace`. Add the delegated permissions Microsoft Graph `User.Read` and SharePoint `AllSites.Write`. The same app is used for the AppSource SaaS offer.
2. Put the client ID in `web/config.js` (and in the website's `crm-addin/config.js` and `js/marketplace-config.js`).
3. The files are served from `https://brillienta.com/crm-addin/` (BrillientaWebSite repository, folder `crm-addin`). After a change here, copy `web/*` and `manifest.xml` there.
4. Optional: submit the add-in to AppSource (Partner Center › Office add-in) so customers find it in Outlook's *Get add-ins*.

## Customer set-up

1. Microsoft 365 admin center › **Settings › Integrated apps › Upload custom apps** › *Office Add-in* › URL `https://brillienta.com/crm-addin/manifest.xml` › assign to users.
2. A Global Admin grants consent once: `https://login.microsoftonline.com/organizations/adminconsent?client_id=<client ID>`
3. Each user opens an email › **CRM** button › enters the CRM site URL once (it's remembered in their mailbox settings).

## Limits

* It uses the default list names (`BrillCRM_Contacts`, `BrillCRM_Leads`, …). Sites that renamed lists in the web part properties aren't supported by the add-in yet.
* It doesn't work in compose mode (read mode only).
