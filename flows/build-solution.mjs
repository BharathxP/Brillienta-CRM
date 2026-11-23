// Builds "Brillienta CRM Flows" — an importable (unmanaged) Power Platform solution with 7 cloud flows.
// Usage: node flows/build-solution.mjs            → flows/dist/BrillientaCRMFlows_1_0_0_0.zip
import { createHash } from 'node:crypto';
import { mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { execSync } from 'node:child_process';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));
const VERSION = '2.0.1.0';
const SOLUTION = 'BrillientaCRMFlows';
const PREFIX = 'bri';

/** Stable GUIDs so re-importing a newer version upgrades the same flows. */
const guid = (seed) => {
  const h = createHash('sha1').update(`brillienta-crm-flows:${seed}`).digest('hex');
  return `${h.slice(0, 8)}-${h.slice(8, 12)}-5${h.slice(13, 16)}-a${h.slice(17, 20)}-${h.slice(20, 32)}`;
};

// ------------------------------------------------------------------ environment variables
const ENV = {
  site: { schema: `${PREFIX}_CrmSiteUrl`, name: 'CRM site URL', def: 'https://contoso.sharepoint.com/sites/crm', desc: 'SharePoint site that hosts the Brillienta CRM lists.' },
  page: { schema: `${PREFIX}_CrmPageUrl`, name: 'CRM page URL', def: 'https://contoso.sharepoint.com/sites/crm/SitePages/CRM.aspx', desc: 'Link to the page with the CRM web part (used in emails).' },
  sales: { schema: `${PREFIX}_SalesManagerEmail`, name: 'Sales manager email', def: 'sales.manager@contoso.com', desc: 'Receives the daily digest, unassigned-lead alerts and quote approvals.' },
  finance: { schema: `${PREFIX}_FinanceEmail`, name: 'Finance email', def: 'finance@contoso.com', desc: 'Notified when a deal is won.' },
  support: { schema: `${PREFIX}_SupportManagerEmail`, name: 'Support manager email', def: 'support.manager@contoso.com', desc: 'Notified when a case breaches its SLA.' },
  key: { schema: `${PREFIX}_WebsiteKey`, name: 'Website shared key', def: 'off', desc: 'Same value as CRM_WEBHOOK_KEY (or LEAD_WEBHOOK_KEY) in Vercel; checked against the x-lead-key header. "off" accepts every request.' },
  rzpId: { schema: `${PREFIX}_RazorpayKeyId`, name: 'Razorpay key id', def: 'not-set', desc: 'Razorpay Dashboard › Account & Settings › API keys (rzp_live_… or rzp_test_…). Leave empty if you do not use Razorpay.' },
  rzpSecret: { schema: `${PREFIX}_RazorpayKeySecret`, name: 'Razorpay key secret', def: 'not-set', desc: 'Razorpay key secret. Treat as a password: restrict who can edit these flows, or switch this variable to an Azure Key Vault secret.' },
  stripeKey: { schema: `${PREFIX}_StripeSecretKey`, name: 'Stripe secret key', def: 'not-set', desc: 'Stripe restricted key (rk_…) with write access to Prices and Payment Links and read access to Checkout Sessions. Leave empty if you do not use Stripe.' },
  waPhoneId: { schema: `${PREFIX}_WhatsAppPhoneNumberId`, name: 'WhatsApp phone number ID', def: 'not-set', desc: 'Meta › WhatsApp › API setup › Phone number ID of your WhatsApp Business number.' },
  waToken: { schema: `${PREFIX}_WhatsAppAccessToken`, name: 'WhatsApp access token', def: 'not-set', desc: 'Permanent System User access token with whatsapp_business_messaging. Treat as a password (consider an Azure Key Vault secret).' },
  waVerify: { schema: `${PREFIX}_WhatsAppVerifyToken`, name: 'WhatsApp webhook verify token', def: 'not-set', desc: 'Any long random text; enter the same value as the Verify token in Meta › WhatsApp › Configuration › Webhook.' },
  portalUrl: { schema: `${PREFIX}_PortalUrl`, name: 'Customer portal URL', def: 'https://example.invalid/portal?not-set=1', desc: 'After import: copy the HTTP URL of the flow "Customer portal" here (it is used in the portal links and sign-in emails).' },
  apolloKey: { schema: `${PREFIX}_ApolloApiKey`, name: 'Apollo.io API key', def: 'not-set', desc: 'Optional. Apollo.io API key for company firmographics (industry, size, location, revenue). Without it the enrichment flows only read the company website.' },
  apiKey: { schema: `${PREFIX}_InboundApiKey`, name: 'Inbound API key', def: 'not-set', desc: 'Long random secret that callers of the Inbound API flow send in the x-api-key header. The API is closed while this is not-set.' },
  mailbox: { schema: `${PREFIX}_SupportMailbox`, name: 'Support mailbox', def: 'support@contoso.com', desc: 'Shared mailbox watched by the Email to case flow (the flow connection needs Full Access or Send As on it).' },
  casePrefix: { schema: `${PREFIX}_CasePrefix`, name: 'Case number prefix', def: 'CS', desc: 'Same as Settings › Company › Case number prefix in the CRM. New email cases are numbered PREFIX-YEAR-ID.' },
  library: { schema: `${PREFIX}_CustomerFolderLibrary`, name: 'Customer folder library', def: 'Documents', desc: 'Document library (on the CRM site) where a folder is created for every won customer.' }
};
const pname = (k) => `${ENV[k].name} (${ENV[k].schema})`;
const P = (k) => `@parameters('${pname(k)}')`;
const Pi = (k) => `@{parameters('${pname(k)}')}`;

// ------------------------------------------------------------------ connectors
const CONN = {
  sp: { api: 'shared_sharepointonline', ref: `${PREFIX}_BrillientaCRM_SharePoint`, display: 'Brillienta CRM – SharePoint' },
  outlook: { api: 'shared_office365', ref: `${PREFIX}_BrillientaCRM_Outlook`, display: 'Brillienta CRM – Office 365 Outlook' },
  conv: { api: 'shared_conversionservice', ref: `${PREFIX}_BrillientaCRM_ContentConversion`, display: 'Brillienta CRM – Content Conversion' },
  approvals: { api: 'shared_approvals', ref: `${PREFIX}_BrillientaCRM_Approvals`, display: 'Brillienta CRM – Approvals' }
};
const host = (c, op) => ({ connectionName: CONN[c].api, operationId: op, apiId: `/providers/Microsoft.PowerApps/apis/${CONN[c].api}` });
const api = (c, op, parameters, runAfter = {}, extra = {}) => ({
  runAfter, type: 'OpenApiConnection', inputs: { host: host(c, op), parameters, authentication: "@parameters('$authentication')" }, ...extra
});
const LISTS = { demos: 'BrillCRM_DemoRequests', leads: 'BrillCRM_Leads', accounts: 'BrillCRM_Accounts', opps: 'BrillCRM_Opportunities', quotes: 'BrillCRM_Quotes', cases: 'BrillCRM_Cases', acts: 'BrillCRM_Activities', invoices: 'BrillCRM_Invoices', subs: 'BrillCRM_Subscriptions', contacts: 'BrillCRM_Contacts', settingsList: 'BrillCRM_Settings', articles: 'BrillCRM_KnowledgeBase' };

const getItems = (list, filter, runAfter = {}, top = 500) =>
  api('sp', 'GetItems', { dataset: P('site'), table: LISTS[list], '$filter': filter, '$top': top }, runAfter);
const email = (to, subject, body, runAfter = {}, extra = {}) =>
  api('outlook', 'SendEmailV2', { 'emailMessage/To': to, 'emailMessage/Subject': subject, 'emailMessage/Body': body, 'emailMessage/Importance': 'Normal', ...extra }, runAfter);
const spPoll = (op, list, conditions) => ({
  recurrence: { interval: 5, frequency: 'Minute' },
  splitOn: "@triggerOutputs()?['body/value']",
  ...(conditions ? { conditions: conditions.map(expression => ({ expression })) } : {}),
  type: 'OpenApiConnection',
  inputs: { host: host('sp', op), parameters: { dataset: P('site'), table: LISTS[list] }, authentication: "@parameters('$authentication')" }
});
const schedule = (frequency, hours, weekDays) => ({
  recurrence: { frequency, interval: 1, timeZone: 'UTC', ...(hours ? { schedule: { hours, minutes: [0], ...(weekDays ? { weekDays } : {}) } } : {}) },
  metadata: { operationMetadataId: guid(`rec-${frequency}-${hours}`) },
  type: 'Recurrence'
});
const WEEKDAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];

// Shared HTML wrapper for emails
const STYLE = '<style>body{font-family:Segoe UI,Arial,sans-serif;color:#121317}table{border-collapse:collapse;margin:6px 0 16px}th{background:#e6f4ee;text-align:left;padding:6px 10px;border-bottom:1px solid #b9dccd}td{padding:6px 10px;border-bottom:1px solid #eee}h2{color:#008D62;margin:18px 0 6px;font-size:16px}a.btn{background:#008D62;color:#fff;padding:8px 14px;border-radius:6px;text-decoration:none}</style>';
const html = (title, inner) => `<html><head>${STYLE}</head><body><p style="font-size:18px;font-weight:600;color:#008D62">${title}</p>${inner}<p style="margin-top:20px"><a class="btn" href="${Pi('page')}">Open Brillienta CRM</a></p><p style="color:#888;font-size:12px">Sent by the Brillienta CRM flow pack.</p></body></html>`;
const tr = (label, value) => `<tr><td style="color:#666">${label}</td><td><b>${value}</b></td></tr>`;

// ------------------------------------------------------------------ flows
const FLOWS = [];
const flow = (name, description, connections, envKeys, triggers, actions) => FLOWS.push({ name, description, connections, envKeys, triggers, actions, id: guid(name) });

// 1 — AppSource / Partner Center lead intake (HTTP endpoint for Partner Center "Customer leads")
flow('Brillienta CRM - AppSource lead intake',
  'Receives Microsoft AppSource / Azure Marketplace leads from Partner Center (HTTPS endpoint) and creates them in the CRM. Requires Power Automate Premium (HTTP trigger).',
  ['sp'], ['site'],
  {
    manual: {
      type: 'Request', kind: 'Http',
      inputs: {
        triggerAuthenticationType: 'All',
        method: 'POST',
        schema: {
          type: 'object',
          properties: {
            UserDetails: { type: 'object', properties: { FirstName: { type: 'string' }, LastName: { type: 'string' }, Email: { type: 'string' }, Phone: { type: 'string' }, Country: { type: 'string' }, Company: { type: 'string' }, Title: { type: 'string' } } },
            LeadSource: { type: 'string' }, ActionCode: { type: 'string' }, OfferTitle: { type: 'string' }, Description: { type: 'string' }
          }
        }
      }
    }
  },
  {
    Create_lead: api('sp', 'PostItem', {
      dataset: P('site'), table: LISTS.leads,
      'item/Title': "@{trim(concat(coalesce(triggerBody()?['UserDetails']?['FirstName'], ''), ' ', coalesce(triggerBody()?['UserDetails']?['LastName'], '')))}",
      'item/FirstName': "@triggerBody()?['UserDetails']?['FirstName']",
      'item/LastName': "@triggerBody()?['UserDetails']?['LastName']",
      'item/Email': "@triggerBody()?['UserDetails']?['Email']",
      'item/Phone': "@triggerBody()?['UserDetails']?['Phone']",
      'item/Company': "@triggerBody()?['UserDetails']?['Company']",
      'item/JobTitle': "@triggerBody()?['UserDetails']?['Title']",
      'item/Country': "@triggerBody()?['UserDetails']?['Country']",
      'item/LeadSource/Value': "@{if(contains(toLower(coalesce(triggerBody()?['LeadSource'], '')), 'azure'), 'Azure Marketplace', 'AppSource')}",
      'item/ActionCode': "@triggerBody()?['ActionCode']",
      'item/Offer': "@triggerBody()?['OfferTitle']",
      'item/Description': "@triggerBody()?['Description']",
      'item/RawPayload': '@{string(triggerBody())}',
      'item/Status/Value': 'New',
      'item/Rating/Value': 'Warm'
    }),
    Response: { runAfter: { Create_lead: ['Succeeded', 'Failed', 'Skipped', 'TimedOut'] }, type: 'Response', kind: 'Http', inputs: { statusCode: 200 } }
  });

// 1b — Website leads, demo bookings & support requests.
// brillienta.com api/lead.js validates every form submission and POSTs a clean record here (CRM_WEBHOOK_URL).
const L = (k) => `triggerBody()?['${k}']`;
const SUPPORT_TYPES = "createArray('Issue', 'Feature request', 'How-to question', 'Licence & billing', 'Security or compliance')";
flow('Brillienta CRM - Website leads and demo bookings',
  'Receives validated submissions from the brillienta.com /api/lead function (book-demo and contact forms). Demo bookings → Hot lead + demo request; support requests → case with SLA; other enquiries → lead. Checks the x-lead-key header. Requires Power Automate Premium (HTTP trigger).',
  ['sp', 'outlook'], ['site', 'page', 'sales', 'support', 'key'],
  {
    manual: {
      type: 'Request', kind: 'Http',
      inputs: {
        triggerAuthenticationType: 'All', method: 'POST',
        schema: {
          type: 'object',
          properties: {
            kind: { type: 'string' }, name: { type: 'string' }, email: { type: 'string' }, company: { type: 'string' }, role: { type: 'string' },
            size: { type: 'string' }, country: { type: 'string' }, message: { type: 'string' },
            interests: { type: 'array', items: { type: 'string' } }, products: { type: 'array', items: { type: 'string' } },
            data_location: { type: 'string' }, timeline: { type: 'string' }, request_type: { type: 'string' }, product: { type: 'string' },
            subject: { type: 'string' }, impact: { type: 'string' }, version: { type: 'string' }, reference: { type: 'string' },
            duration_minutes: { type: ['integer', 'null'] }, requested_start: { type: 'string' }, visitor_time_zone: { type: 'string' },
            consent: { type: 'boolean' }, page: { type: 'string' }, referrer: { type: 'string' }, landing_page: { type: 'string' },
            attribution: { type: 'object' }, received_at: { type: 'string' }
          }
        }
      }
    }
  },
  {
    Key_ok: {
      runAfter: {}, type: 'If',
      expression: { or: [
        { equals: [`@toLower(${P('key').slice(1)})`, 'off'] },
        { equals: ["@coalesce(triggerOutputs()?['headers']?['x-lead-key'], '')", P('key')] }
      ] },
      actions: {
        Full_name: { runAfter: {}, type: 'Compose', inputs: `@trim(coalesce(${L('name')}, ''))` },
        First_name: { runAfter: { Full_name: ['Succeeded'] }, type: 'Compose', inputs: "@first(split(outputs('Full_name'), ' '))" },
        Last_name: { runAfter: { First_name: ['Succeeded'] }, type: 'Compose', inputs: "@if(contains(outputs('Full_name'), ' '), trim(substring(outputs('Full_name'), indexOf(outputs('Full_name'), ' '))), '')" },
        Products: { runAfter: {}, type: 'Compose', inputs: `@join(union(coalesce(${L('products')}, json('[]')), coalesce(${L('interests')}, json('[]')), if(empty(coalesce(${L('product')}, '')), json('[]'), createArray(${L('product')}))), ', ')` },
        Kind: { runAfter: {}, type: 'Compose', inputs: `@if(equals(${L('kind')}, 'demo'), 'demo', if(contains(${SUPPORT_TYPES}, coalesce(${L('request_type')}, '')), 'support', 'lead'))` },
        Details: {
          runAfter: { Products: ['Succeeded'], Kind: ['Succeeded'] }, type: 'Compose',
          inputs: `@{coalesce(${L('message')}, '')}

— brillienta.com @{coalesce(${L('kind')}, '')} · @{coalesce(${L('reference')}, '')} —
Request: @{coalesce(${L('request_type')}, '')}
Subject: @{coalesce(${L('subject')}, '')}
Products / interests: @{outputs('Products')}
Impact: @{coalesce(${L('impact')}, '')}   Version: @{coalesce(${L('version')}, '')}
Requested time (UTC): @{coalesce(${L('requested_start')}, '')}   Visitor time zone: @{coalesce(${L('visitor_time_zone')}, '')}   Length: @{coalesce(${L('duration_minutes')}, '')} min
Timeline: @{coalesce(${L('timeline')}, '')}   Data location: @{coalesce(${L('data_location')}, '')}   Company size: @{coalesce(${L('size')}, '')}
UTM: @{coalesce(${L('attribution')}?['utm_source'], '')} / @{coalesce(${L('attribution')}?['utm_medium'], '')} / @{coalesce(${L('attribution')}?['utm_campaign'], '')}
Page: @{coalesce(${L('page')}, '')}   Landing page: @{coalesce(${L('landing_page')}, '')}   Referrer: @{coalesce(${L('referrer')}, '')}`
        },
        Route: {
          runAfter: { Details: ['Succeeded'], Last_name: ['Succeeded'] }, type: 'If',
          expression: { and: [{ equals: ["@outputs('Kind')", 'support'] }] },
          // ---------------- support request → case
          actions: {
            Find_contact: api('sp', 'GetItems', { dataset: P('site'), table: 'BrillCRM_Contacts', '$filter': `Email eq '@{replace(toLower(${L('email')}), '''', '''''')}'`, '$top': 1 }),
            Priority: { runAfter: { Find_contact: ['Succeeded', 'Failed'] }, type: 'Compose', inputs: `@if(equals(${L('impact')}, 'Everyone is blocked'), 'Critical', if(equals(${L('impact')}, 'Some people cannot work'), 'High', if(equals(${L('impact')}, 'Minor or cosmetic'), 'Low', 'Normal')))` },
            Create_case: api('sp', 'PostItem', {
              dataset: P('site'), table: LISTS.cases,
              'item/Title': `@{if(empty(coalesce(${L('subject')}, '')), concat(${L('request_type')}, ' from ', outputs('Full_name')), ${L('subject')})}`,
              'item/CaseNumber': `@{${L('reference')}}`,
              'item/Priority/Value': "@{outputs('Priority')}",
              'item/CaseStatus/Value': 'New',
              'item/Category/Value': `@{if(equals(${L('request_type')}, 'Issue'), 'Problem', if(equals(${L('request_type')}, 'Feature request'), 'Feature request', if(equals(${L('request_type')}, 'Licence & billing'), 'Billing', 'Question')))}`,
              'item/Origin/Value': 'Web',
              // SLA deadlines are calculated by the CRM (business hours, first response) when support opens it.
              'item/RequesterEmail': `@{toLower(${L('email')})}`,
              'item/ContactRef': "@first(coalesce(body('Find_contact')?['value'], json('[]')))?['ID']",
              'item/AccountRef': "@first(coalesce(body('Find_contact')?['value'], json('[]')))?['AccountRef']",
              'item/Description': `From: @{outputs('Full_name')} <@{${L('email')}}> · @{coalesce(${L('company')}, '')}\n\n@{outputs('Details')}`
            }, { Priority: ['Succeeded'] }),
            Alert_support: email(P('support'), `[@{${L('reference')}}] @{${L('request_type')}} (@{outputs('Priority')}): @{coalesce(${L('subject')}, outputs('Full_name'))}`,
              html('New support request from brillienta.com', `<table>${tr('Reference', `@{${L('reference')}}`)}${tr('Type', `@{${L('request_type')}}`)}${tr('Priority', "@{outputs('Priority')}")}${tr('Product', `@{${L('product')}} @{${L('version')}}`)}${tr('From', `@{outputs('Full_name')} &lt;@{${L('email')}}&gt; · @{${L('company')}}`)}${tr('Impact', `@{${L('impact')}}`)}${tr('Subject', `@{${L('subject')}}`)}${tr('Message', `@{${L('message')}}`)}</table><p>The case is in the CRM under Cases.</p>`),
              { Create_case: ['Succeeded'] }, { 'emailMessage/Importance': "@{if(contains(createArray('Critical', 'High'), outputs('Priority')), 'High', 'Normal')}" })
          },
          else: {
            actions: {
              // ---------------- demo booking / enquiry → lead
              Create_lead: api('sp', 'PostItem', {
                dataset: P('site'), table: LISTS.leads,
                'item/Title': "@{outputs('Full_name')}",
                'item/FirstName': "@{outputs('First_name')}",
                'item/LastName': "@{outputs('Last_name')}",
                'item/Email': `@{${L('email')}}`,
                'item/Company': `@{${L('company')}}`,
                'item/JobTitle': `@{${L('role')}}`,
                'item/Country': `@{${L('country')}}`,
                'item/LeadSource/Value': "@{if(equals(outputs('Kind'), 'demo'), 'Website Demo Form', 'Website')}",
                'item/Status/Value': 'New',
                'item/Rating/Value': "@{if(equals(outputs('Kind'), 'demo'), 'Hot', 'Warm')}",
                'item/ActionCode': `@{${L('reference')}}`,
                'item/Offer': "@{outputs('Products')}",
                'item/Description': "@{outputs('Details')}",
                'item/RawPayload': '@{string(triggerBody())}'
              }),
              Demo_booking: {
                runAfter: { Create_lead: ['Succeeded'] }, type: 'If',
                expression: { and: [{ equals: ["@outputs('Kind')", 'demo'] }] },
                actions: {
                  Create_demo_request: api('sp', 'PostItem', {
                    dataset: P('site'), table: LISTS.demos,
                    'item/Title': `Demo - @{${L('company')}} (@{outputs('Products')})`,
                    'item/ContactName': "@{outputs('Full_name')}",
                    'item/Email': `@{${L('email')}}`,
                    'item/Company': `@{${L('company')}}`,
                    'item/LeadRef': "@body('Create_lead')?['ID']",
                    'item/PreferredDate': `@if(empty(coalesce(${L('requested_start')}, '')), null, ${L('requested_start')})`,
                    'item/DurationMinutes': `@coalesce(${L('duration_minutes')}, 30)`,
                    'item/DemoStatus/Value': 'Requested',
                    'item/Notes': "@{outputs('Details')}"
                  }),
                  Alert_sales: email(P('sales'), `Demo booked: @{outputs('Full_name')} (@{${L('company')}}) – @{outputs('Products')}`,
                    html('A visitor booked a demo on brillienta.com', `<table>${tr('Reference', `@{${L('reference')}}`)}${tr('Name', "@{outputs('Full_name')}")}${tr('Company', `@{${L('company')}}`)}${tr('Role', `@{${L('role')}}`)}${tr('Email', `@{${L('email')}}`)}${tr('Products', "@{outputs('Products')}")}${tr('Requested time (UTC)', `@{${L('requested_start')}}`)}${tr('Visitor time zone', `@{${L('visitor_time_zone')}}`)}${tr('Length', `@{${L('duration_minutes')}} minutes`)}${tr('Company size', `@{${L('size')}}`)}${tr('Notes', `@{${L('message')}}`)}</table><p>Confirm it in the CRM: Demo requests › Schedule Teams meeting (sends the Teams invite).</p>`),
                    { Create_demo_request: ['Succeeded'] }, { 'emailMessage/Importance': 'High' })
                },
                else: { actions: {} }
              }
            }
          }
        },
        Respond_ok: {
          runAfter: { Route: ['Succeeded'] }, type: 'Response', kind: 'Http',
          inputs: { statusCode: 200, headers: { 'Content-Type': 'application/json' }, body: { ok: true, reference: `@{${L('reference')}}` } }
        },
        Respond_failed: {
          runAfter: { Route: ['Failed', 'TimedOut'] }, type: 'Response', kind: 'Http',
          inputs: { statusCode: 502, headers: { 'Content-Type': 'application/json' }, body: { ok: false } }
        }
      },
      else: {
        actions: {
          Respond_unauthorized: { runAfter: {}, type: 'Response', kind: 'Http', inputs: { statusCode: 401, headers: { 'Content-Type': 'application/json' }, body: { ok: false, error: 'key' } } }
        }
      }
    }
  });

// 2 — New lead alert
flow('Brillienta CRM - New lead alert',
  'Emails the lead owner when a new lead arrives; unassigned leads go to the sales manager.',
  ['sp', 'outlook'], ['site', 'page', 'sales'],
  { When_a_lead_is_created: spPoll('GetOnNewItems', 'leads') },
  {
    Lead_summary: {
      runAfter: {}, type: 'Compose',
      inputs: `<table>${tr('Name', "@{triggerOutputs()?['body/Title']}")}${tr('Company', "@{triggerOutputs()?['body/Company']}")}${tr('Email', "@{triggerOutputs()?['body/Email']}")}${tr('Phone', "@{triggerOutputs()?['body/Phone']}")}${tr('Source', "@{triggerOutputs()?['body/LeadSource/Value']}")}${tr('Product', "@{triggerOutputs()?['body/Offer']}")}${tr('Message', "@{triggerOutputs()?['body/Description']}")}</table>`
    },
    Is_assigned: {
      runAfter: { Lead_summary: ['Succeeded'] }, type: 'If',
      expression: { and: [{ not: { equals: ["@coalesce(triggerOutputs()?['body/LeadOwner/Email'], '')", ''] } }] },
      actions: {
        Email_owner: email("@triggerOutputs()?['body/LeadOwner/Email']", "New lead for you: @{triggerOutputs()?['body/Title']} (@{triggerOutputs()?['body/Company']})",
          html('You have a new lead', "@{outputs('Lead_summary')}<p>Reach out within 24 hours — fast follow-up doubles conversion.</p>"))
      },
      else: {
        actions: {
          Email_sales_manager: email(P('sales'), "Unassigned lead: @{triggerOutputs()?['body/Title']} (@{triggerOutputs()?['body/Company']})",
            html('A new lead needs an owner', "@{outputs('Lead_summary')}<p>Assign it in the CRM (Leads › Round-robin assign or Assign to me).</p>"))
        }
      }
    }
  });

// 3 — Daily sales digest
const tableFrom = (name, source, map, runAfter) => ({
  [`Select_${name}`]: { runAfter, type: 'Select', inputs: { from: source, select: map } },
  [`Table_${name}`]: { runAfter: { [`Select_${name}`]: ['Succeeded'] }, type: 'Table', inputs: { from: `@body('Select_${name}')`, format: 'HTML' } }
});
flow('Brillienta CRM - Daily sales digest',
  'Every weekday morning: new leads, deals won, deals closing this week and overdue tasks, emailed to the sales manager.',
  ['sp', 'outlook'], ['site', 'page', 'sales'],
  { Every_weekday_morning: schedule('Week', [7], WEEKDAYS) },
  {
    Get_new_leads: getItems('leads', "Created ge '@{addDays(utcNow(), -1)}'"),
    Get_won_deals: getItems('opps', "Stage eq 'Closed Won' and ActualCloseDate ge '@{addDays(utcNow(), -1)}'"),
    Get_closing_deals: getItems('opps', "Stage ne 'Closed Won' and Stage ne 'Closed Lost' and ExpectedCloseDate le '@{addDays(utcNow(), 7)}'"),
    Get_overdue_tasks: getItems('acts', "ActivityType eq 'Task' and Completed eq 0 and DueDate lt '@{utcNow()}'"),
    ...tableFrom('leads', "@body('Get_new_leads')?['value']", { Lead: "@item()?['Title']", Company: "@item()?['Company']", Source: "@item()?['LeadSource']?['Value']", Score: "@item()?['LeadScore']", Owner: "@coalesce(item()?['LeadOwner']?['DisplayName'], 'Unassigned')" }, { Get_new_leads: ['Succeeded'] }),
    ...tableFrom('won', "@body('Get_won_deals')?['value']", { Deal: "@item()?['Title']", Amount: "@item()?['Amount']", Owner: "@coalesce(item()?['DealOwner']?['DisplayName'], '')" }, { Get_won_deals: ['Succeeded'] }),
    ...tableFrom('closing', "@body('Get_closing_deals')?['value']", { Deal: "@item()?['Title']", Stage: "@item()?['Stage']?['Value']", Amount: "@item()?['Amount']", Close: "@formatDateTime(item()?['ExpectedCloseDate'], 'dd MMM')", Owner: "@coalesce(item()?['DealOwner']?['DisplayName'], '')" }, { Get_closing_deals: ['Succeeded'] }),
    ...tableFrom('tasks', "@body('Get_overdue_tasks')?['value']", { Task: "@item()?['Title']", Due: "@formatDateTime(item()?['DueDate'], 'dd MMM HH:mm')", Owner: "@coalesce(item()?['TaskOwner']?['DisplayName'], '')" }, { Get_overdue_tasks: ['Succeeded'] }),
    Send_digest: email(P('sales'), "Sales digest @{formatDateTime(utcNow(), 'dddd dd MMM')}: @{length(body('Get_new_leads')?['value'])} new leads, @{length(body('Get_won_deals')?['value'])} deals won",
      html('Your daily sales digest',
        "<h2>New leads (@{length(body('Get_new_leads')?['value'])})</h2>@{body('Table_leads')}" +
        "<h2>Deals won (@{length(body('Get_won_deals')?['value'])})</h2>@{body('Table_won')}" +
        "<h2>Closing in the next 7 days (@{length(body('Get_closing_deals')?['value'])})</h2>@{body('Table_closing')}" +
        "<h2>Overdue tasks (@{length(body('Get_overdue_tasks')?['value'])})</h2>@{body('Table_tasks')}"),
      { Table_leads: ['Succeeded'], Table_won: ['Succeeded'], Table_closing: ['Succeeded'], Table_tasks: ['Succeeded'] })
  });

// 4 — Personal reminders (overdue tasks + lead follow-ups), one email per owner
flow('Brillienta CRM - Overdue reminders',
  'Every weekday morning each user gets one email with their overdue tasks and overdue lead follow-ups.',
  ['sp', 'outlook'], ['site', 'page'],
  { Every_weekday_morning: schedule('Week', [6], WEEKDAYS) },
  {
    Get_overdue_tasks: getItems('acts', "ActivityType eq 'Task' and Completed eq 0 and DueDate lt '@{utcNow()}'"),
    Get_overdue_followups: getItems('leads', "NextFollowUp lt '@{utcNow()}' and Status ne 'Converted' and Status ne 'Unqualified'"),
    Select_task_owners: { runAfter: { Get_overdue_tasks: ['Succeeded'] }, type: 'Select', inputs: { from: "@body('Get_overdue_tasks')?['value']", select: "@coalesce(item()?['TaskOwner']?['Email'], '')" } },
    Select_lead_owners: { runAfter: { Get_overdue_followups: ['Succeeded'] }, type: 'Select', inputs: { from: "@body('Get_overdue_followups')?['value']", select: "@coalesce(item()?['LeadOwner']?['Email'], '')" } },
    Owners: { runAfter: { Select_task_owners: ['Succeeded'], Select_lead_owners: ['Succeeded'] }, type: 'Compose', inputs: "@union(body('Select_task_owners'), body('Select_lead_owners'))" },
    For_each_owner: {
      runAfter: { Owners: ['Succeeded'] }, type: 'Foreach', foreach: "@outputs('Owners')",
      actions: {
        Has_email: {
          runAfter: {}, type: 'If', expression: { and: [{ not: { equals: ["@items('For_each_owner')", ''] } }] },
          actions: {
            My_tasks: { runAfter: {}, type: 'Query', inputs: { from: "@body('Get_overdue_tasks')?['value']", where: "@equals(coalesce(item()?['TaskOwner']?['Email'], ''), items('For_each_owner'))" } },
            My_followups: { runAfter: {}, type: 'Query', inputs: { from: "@body('Get_overdue_followups')?['value']", where: "@equals(coalesce(item()?['LeadOwner']?['Email'], ''), items('For_each_owner'))" } },
            Select_my_tasks: { runAfter: { My_tasks: ['Succeeded'] }, type: 'Select', inputs: { from: "@body('My_tasks')", select: { Task: "@item()?['Title']", Due: "@formatDateTime(item()?['DueDate'], 'dd MMM HH:mm')", 'Related to': "@concat(item()?['RelatedType'], ' #', item()?['RelatedId'])" } } },
            Select_my_followups: { runAfter: { My_followups: ['Succeeded'] }, type: 'Select', inputs: { from: "@body('My_followups')", select: { Lead: "@item()?['Title']", Company: "@item()?['Company']", 'Was due': "@formatDateTime(item()?['NextFollowUp'], 'dd MMM HH:mm')" } } },
            Table_tasks: { runAfter: { Select_my_tasks: ['Succeeded'] }, type: 'Table', inputs: { from: "@body('Select_my_tasks')", format: 'HTML' } },
            Table_followups: { runAfter: { Select_my_followups: ['Succeeded'] }, type: 'Table', inputs: { from: "@body('Select_my_followups')", format: 'HTML' } },
            Send_reminder: email("@items('For_each_owner')", "You have @{add(length(body('My_tasks')), length(body('My_followups')))} overdue CRM item(s)",
              html('Overdue items in Brillienta CRM', "<h2>Tasks (@{length(body('My_tasks'))})</h2>@{body('Table_tasks')}<h2>Lead follow-ups (@{length(body('My_followups'))})</h2>@{body('Table_followups')}"),
              { Table_tasks: ['Succeeded'], Table_followups: ['Succeeded'] })
          },
          else: { actions: {} }
        }
      }
    }
  });

// 5 — Case SLA breach escalation
flow('Brillienta CRM - Case SLA escalation',
  'Every hour: cases that breached their SLA in the last hour are escalated to the support manager (case owner in Cc).',
  ['sp', 'outlook'], ['site', 'page', 'support'],
  { Every_hour: schedule('Hour') },
  {
    Get_breached_cases: getItems('cases', "SlaDue lt '@{utcNow()}' and SlaDue ge '@{addHours(utcNow(), -1)}' and CaseStatus ne 'Resolved' and CaseStatus ne 'Closed'"),
    For_each_case: {
      runAfter: { Get_breached_cases: ['Succeeded'] }, type: 'Foreach', foreach: "@body('Get_breached_cases')?['value']",
      actions: {
        Escalate: email(P('support'), "SLA breached: @{items('For_each_case')?['CaseNumber']} @{items('For_each_case')?['Title']}",
          html('A case has breached its SLA', `<table>${tr('Case', "@{items('For_each_case')?['CaseNumber']} @{items('For_each_case')?['Title']}")}${tr('Priority', "@{items('For_each_case')?['Priority']?['Value']}")}${tr('Status', "@{items('For_each_case')?['CaseStatus']?['Value']}")}${tr('SLA due', "@{items('For_each_case')?['SlaDue']}")}${tr('Owner', "@{coalesce(items('For_each_case')?['CaseOwner']?['DisplayName'], 'Unassigned')}")}</table>`),
          {}, { 'emailMessage/Cc': "@coalesce(items('For_each_case')?['CaseOwner']?['Email'], '')", 'emailMessage/Importance': 'High' })
      }
    }
  });

// 6 — Deal won → finance + customer folder
flow('Brillienta CRM - Deal won handoff',
  'When an opportunity changes to Closed Won: notifies finance with the deal details and creates a customer folder in the document library.',
  ['sp', 'outlook'], ['site', 'page', 'finance', 'library'],
  { When_a_deal_is_updated: spPoll('GetOnUpdatedItems', 'opps', ["@equals(triggerOutputs()?['body/Stage/Value'], 'Closed Won')"]) },
  {
    Customer_name: { runAfter: {}, type: 'InitializeVariable', inputs: { variables: [{ name: 'customerName', type: 'string', value: "@{coalesce(triggerOutputs()?['body/Company'], triggerOutputs()?['body/Title'])}" }] } },
    Get_changes: api('sp', 'GetItemChanges', {
      dataset: P('site'), table: LISTS.opps, id: "@triggerOutputs()?['body/ID']",
      since: "@triggerOutputs()?['body/{TriggerWindowStartToken}']", includeDrafts: false
    }, { Customer_name: ['Succeeded'] }),
    Stage_just_changed: {
      runAfter: { Get_changes: ['Succeeded'] }, type: 'If',
      expression: { and: [{ equals: ["@body('Get_changes')?['ColumnHasChanged']?['Stage']", true] }] },
      actions: {
        Has_account: {
          runAfter: {}, type: 'If', expression: { and: [{ greater: ["@coalesce(triggerOutputs()?['body/AccountRef'], 0)", 0] }] },
          actions: {
            Get_account: api('sp', 'GetItem', { dataset: P('site'), table: LISTS.accounts, id: "@triggerOutputs()?['body/AccountRef']" }),
            Use_account_name: { runAfter: { Get_account: ['Succeeded'] }, type: 'SetVariable', inputs: { name: 'customerName', value: "@body('Get_account')?['Title']" } }
          },
          else: { actions: {} }
        },
        Notify_finance: email(P('finance'), "Deal won: @{triggerOutputs()?['body/Title']} (@{triggerOutputs()?['body/Amount']})",
          html('A deal has been won — please prepare invoicing', `<table>${tr('Customer', "@{variables('customerName')}")}${tr('Deal', "@{triggerOutputs()?['body/Title']}")}${tr('Amount', "@{triggerOutputs()?['body/Amount']}")}${tr('Licenses / seats', "@{triggerOutputs()?['body/Licenses']}")}${tr('Contact', "@{triggerOutputs()?['body/ContactName']} @{triggerOutputs()?['body/Email']}")}${tr('Owner', "@{triggerOutputs()?['body/DealOwner/DisplayName']}")}${tr('Closed on', "@{triggerOutputs()?['body/ActualCloseDate']}")}</table>`),
          { Has_account: ['Succeeded'] }, { 'emailMessage/Cc': "@coalesce(triggerOutputs()?['body/DealOwner/Email'], '')" }),
        Create_customer_folder: api('sp', 'CreateNewFolder', {
          dataset: P('site'), table: P('library'),
          'parameters/path': "Customers/@{trim(replace(replace(replace(replace(replace(variables('customerName'), '/', '-'), ':', '-'), '#', ''), '%', ''), '\"', ''))}"
        }, { Notify_finance: ['Succeeded', 'Failed'] })
      },
      else: { actions: {} }
    }
  });

// 7 — Quote discount approval
flow('Brillienta CRM - Quote discount approval',
  'When a rep clicks "Request approval" on a quote, the sales manager gets an Approvals request (Teams/Outlook). The decision is written back to the quote.',
  ['sp', 'outlook', 'approvals'], ['site', 'page', 'sales'],
  { When_approval_is_requested: spPoll('GetOnUpdatedItems', 'quotes', ["@equals(triggerOutputs()?['body/ApprovalStatus/Value'], 'Pending')"]) },
  {
    Discount_percent: {
      runAfter: {}, type: 'Compose',
      inputs: "@if(greater(float(coalesce(triggerOutputs()?['body/Subtotal'], 0)), 0), div(mul(float(coalesce(triggerOutputs()?['body/DiscountTotal'], 0)), 100), float(triggerOutputs()?['body/Subtotal'])), 0)"
    },
    Ask_approval: {
      runAfter: { Discount_percent: ['Succeeded'] }, type: 'OpenApiConnectionWebhook',
      inputs: {
        host: host('approvals', 'StartAndWaitForAnApproval'),
        parameters: {
          approvalType: 'Basic',
          'WebhookApprovalCreationInput/title': "Approve discount on quote @{triggerOutputs()?['body/QuoteNumber']} (@{formatNumber(outputs('Discount_percent'), '0.#')}% off)",
          'WebhookApprovalCreationInput/assignedTo': P('sales'),
          'WebhookApprovalCreationInput/details': "**Quote:** @{triggerOutputs()?['body/QuoteNumber']} – @{triggerOutputs()?['body/Title']}\n\n**Subtotal:** @{triggerOutputs()?['body/Subtotal']}\n\n**Discount:** @{triggerOutputs()?['body/DiscountTotal']} (@{formatNumber(outputs('Discount_percent'), '0.#')}%)\n\n**Total:** @{triggerOutputs()?['body/GrandTotal']}\n\n**Requested by:** @{triggerOutputs()?['body/QuoteOwner/DisplayName']}",
          'WebhookApprovalCreationInput/itemLink': "@triggerOutputs()?['body/{Link}']",
          'WebhookApprovalCreationInput/itemLinkDescription': "Quote @{triggerOutputs()?['body/QuoteNumber']}",
          'WebhookApprovalCreationInput/enableNotifications': true,
          'WebhookApprovalCreationInput/enableReassignment': true
        },
        authentication: "@parameters('$authentication')"
      }
    },
    Approved: {
      runAfter: { Ask_approval: ['Succeeded'] }, type: 'If',
      expression: { and: [{ equals: ["@body('Ask_approval')?['outcome']", 'Approve'] }] },
      actions: {
        Mark_approved: api('sp', 'PatchItem', {
          dataset: P('site'), table: LISTS.quotes, id: "@triggerOutputs()?['body/ID']", 'item/Title': "@triggerOutputs()?['body/Title']",
          'item/ApprovalStatus/Value': 'Approved',
          'item/ApprovalNote': "Approved by @{body('Ask_approval')?['responses']?[0]?['responder']?['displayName']}: @{coalesce(body('Ask_approval')?['responses']?[0]?['comments'], '')}"
        })
      },
      else: {
        actions: {
          Mark_rejected: api('sp', 'PatchItem', {
            dataset: P('site'), table: LISTS.quotes, id: "@triggerOutputs()?['body/ID']", 'item/Title': "@triggerOutputs()?['body/Title']",
            'item/ApprovalStatus/Value': 'Rejected',
            'item/ApprovalNote': "Rejected by @{body('Ask_approval')?['responses']?[0]?['responder']?['displayName']}: @{coalesce(body('Ask_approval')?['responses']?[0]?['comments'], '')}"
          })
        }
      }
    },
    Tell_the_rep: email("@coalesce(triggerOutputs()?['body/QuoteOwner/Email'], parameters('" + pname('sales') + "'))",
      "Quote @{triggerOutputs()?['body/QuoteNumber']} was @{if(equals(body('Ask_approval')?['outcome'], 'Approve'), 'approved', 'rejected')}",
      html("Quote @{triggerOutputs()?['body/QuoteNumber']}: @{if(equals(body('Ask_approval')?['outcome'], 'Approve'), 'approved ✔', 'rejected ✖')}",
        "<p>Comments: @{coalesce(body('Ask_approval')?['responses']?[0]?['comments'], '—')}</p>"),
      { Approved: ['Succeeded'] })
  });


// ------------------------------------------------------------------ payments (Razorpay / Stripe)
const T = (k) => `triggerOutputs()?['body/${k}']`;
const balance = `sub(float(coalesce(${T('GrandTotal')}, 0)), float(coalesce(${T('AmountPaid')}, 0)))`;
const minor = `int(formatNumber(mul(${balance}, 100), '0', 'en-US'))`;
const http = (method, uri, headers, body, runAfter = {}, extra = {}) => ({
  runAfter, type: 'Http', inputs: { method, uri, headers, ...(body !== undefined ? { body } : {}), ...extra }
});
const patchInvoice = (id, fields, runAfter = {}) => api('sp', 'PatchItem', { dataset: P('site'), table: LISTS.invoices, id, ...fields }, runAfter);
const logInvoice = (id, title, details, runAfter) => api('sp', 'PostItem', {
  dataset: P('site'), table: LISTS.acts, 'item/Title': title, 'item/Details': details, 'item/ActivityType/Value': 'Note',
  'item/RelatedType': 'Invoice', 'item/RelatedId': id, 'item/Completed': true, 'item/ActivitySource/Value': 'Flow'
}, runAfter);
const linkRequested = (provider) => spPoll('GetOnUpdatedItems', 'invoices', [
  `@equals(${T('PaymentLinkStatus/Value')}, 'Requested')`, `@equals(${T('PaymentProvider/Value')}, '${provider}')`
]);
const contactLookup = {
  Contact_email: { runAfter: {}, type: 'InitializeVariable', inputs: { variables: [{ name: 'contactEmail', type: 'string', value: '' }] } },
  Has_contact: {
    runAfter: { Contact_email: ['Succeeded'] }, type: 'If', expression: { and: [{ greater: [`@coalesce(${T('ContactRef')}, 0)`, 0] }] },
    actions: {
      Get_contact: api('sp', 'GetItem', { dataset: P('site'), table: 'BrillCRM_Contacts', id: `@${T('ContactRef')}` }),
      Use_contact_email: { runAfter: { Get_contact: ['Succeeded'] }, type: 'SetVariable', inputs: { name: 'contactEmail', value: "@{coalesce(body('Get_contact')?['Email'], '')}" } }
    },
    else: { actions: {} }
  }
};
const respond = (status, runAfter) => ({ runAfter, type: 'Response', kind: 'Http', inputs: { statusCode: status } });
const webhookTrigger = { manual: { type: 'Request', kind: 'Http', inputs: { triggerAuthenticationType: 'All', method: 'POST', schema: { type: 'object' } } } };
const markPaid = (lookupFilter, amountExpr, provider, reference, runAfter) => ({
  Find_invoice: getItems('invoices', lookupFilter, runAfter, 1),
  For_the_invoice: {
    runAfter: { Find_invoice: ['Succeeded'] }, type: 'Foreach', foreach: "@body('Find_invoice')?['value']",
    actions: {
      Mark_paid: patchInvoice("@items('For_the_invoice')?['ID']", {
        'item/Title': "@items('For_the_invoice')?['Title']", 'item/InvoiceStatus/Value': 'Paid', 'item/PaidDate': '@utcNow()',
        'item/AmountPaid': `@add(float(coalesce(items('For_the_invoice')?['AmountPaid'], 0)), ${amountExpr})`, 'item/PaymentProvider/Value': provider
      }),
      Log_payment: logInvoice("@items('For_the_invoice')?['ID']", `Payment received via ${provider}`, `@{${amountExpr}} · ${reference}`, { Mark_paid: ['Succeeded'] })
    }
  }
});

// 9 — Razorpay payment link
flow('Brillienta CRM - Razorpay payment link',
  'When "Create Razorpay link" is clicked on an invoice: creates a Razorpay Payment Link for the balance (reference = invoice number), emails it to the billing contact and writes the short URL back to the invoice. Requires Power Automate Premium (HTTP).',
  ['sp'], ['site', 'rzpId', 'rzpSecret'],
  { When_a_link_is_requested: linkRequested('Razorpay') },
  {
    ...contactLookup,
    Create_link: http('POST', 'https://api.razorpay.com/v1/payment_links', { 'Content-Type': 'application/json' }, {
      amount: `@${minor}`, currency: "@{coalesce(" + T('Currency') + ", 'INR')}", accept_partial: false,
      reference_id: `@{${T('InvoiceNumber')}}`, description: `Invoice @{${T('InvoiceNumber')}} - @{${T('Title')}}`,
      customer: { email: "@{variables('contactEmail')}" }, notify: { email: "@not(empty(variables('contactEmail')))", sms: false }, reminder_enable: true,
      notes: { crm_invoice_id: `@{${T('ID')}}` }
    }, { Has_contact: ['Succeeded'] }, { authentication: { type: 'Basic', username: P('rzpId'), password: P('rzpSecret') } }),
    Save_link: patchInvoice(`@${T('ID')}`, {
      'item/Title': `@${T('Title')}`, 'item/PaymentLink': "@body('Create_link')?['short_url']",
      'item/ExternalPaymentId': "@body('Create_link')?['id']", 'item/PaymentLinkStatus/Value': 'Created'
    }, { Create_link: ['Succeeded'] }),
    Save_failure: patchInvoice(`@${T('ID')}`, { 'item/Title': `@${T('Title')}`, 'item/PaymentLinkStatus/Value': 'Failed' }, { Create_link: ['Failed', 'TimedOut'] }),
    Log_failure: logInvoice(`@${T('ID')}`, 'Razorpay payment link failed', "@{body('Create_link')}", { Save_failure: ['Succeeded'] })
  });

// 10 — Razorpay payment received
flow('Brillienta CRM - Razorpay payment received',
  'Razorpay webhook (event payment_link.paid). The payment link is re-read from the Razorpay API before the invoice is marked Paid, so forged webhook calls are ignored. Requires Power Automate Premium (HTTP trigger).',
  ['sp'], ['site', 'rzpId', 'rzpSecret'],
  webhookTrigger,
  {
    Link_id: { runAfter: {}, type: 'Compose', inputs: "@{coalesce(triggerBody()?['payload']?['payment_link']?['entity']?['id'], '')}" },
    Is_link_event: {
      runAfter: { Link_id: ['Succeeded'] }, type: 'If',
      expression: { and: [{ equals: ["@triggerBody()?['event']", 'payment_link.paid'] }, { startsWith: ["@outputs('Link_id')", 'plink_'] }] },
      actions: {
        Verify_link: http('GET', "https://api.razorpay.com/v1/payment_links/@{outputs('Link_id')}", {}, undefined, {}, { authentication: { type: 'Basic', username: P('rzpId'), password: P('rzpSecret') } }),
        Is_paid: {
          runAfter: { Verify_link: ['Succeeded'] }, type: 'If',
          expression: { and: [{ equals: ["@body('Verify_link')?['status']", 'paid'] }] },
          actions: markPaid("ExternalPaymentId eq '@{outputs('Link_id')}'", "div(float(body('Verify_link')?['amount_paid']), 100)", 'Razorpay', "@{outputs('Link_id')}", {}),
          else: { actions: {} }
        }
      },
      else: { actions: {} }
    },
    Response: respond(200, { Is_link_event: ['Succeeded', 'Failed', 'Skipped', 'TimedOut'] })
  });

// 11 — Stripe payment link
const stripeAuth = { Authorization: `Bearer ${Pi('stripeKey')}`, 'Content-Type': 'application/x-www-form-urlencoded' };
flow('Brillienta CRM - Stripe payment link',
  'When "Create Stripe link" is clicked on an invoice: creates a Stripe Price and Payment Link for the balance and writes the URL back to the invoice. Requires Power Automate Premium (HTTP).',
  ['sp'], ['site', 'stripeKey'],
  { When_a_link_is_requested: linkRequested('Stripe') },
  {
    Create_price: http('POST', 'https://api.stripe.com/v1/prices', stripeAuth,
      `currency=@{toLower(coalesce(${T('Currency')}, 'usd'))}&unit_amount=@{${minor}}&product_data[name]=@{encodeUriComponent(concat('Invoice ', ${T('InvoiceNumber')}))}`),
    Create_link: http('POST', 'https://api.stripe.com/v1/payment_links', stripeAuth,
      `line_items[0][price]=@{body('Create_price')?['id']}&line_items[0][quantity]=1&metadata[invoice]=@{encodeUriComponent(${T('InvoiceNumber')})}&metadata[crm_invoice_id]=@{${T('ID')}}&payment_intent_data[metadata][invoice]=@{encodeUriComponent(${T('InvoiceNumber')})}`,
      { Create_price: ['Succeeded'] }),
    Save_link: patchInvoice(`@${T('ID')}`, {
      'item/Title': `@${T('Title')}`, 'item/PaymentLink': "@body('Create_link')?['url']",
      'item/ExternalPaymentId': "@body('Create_link')?['id']", 'item/PaymentLinkStatus/Value': 'Created'
    }, { Create_link: ['Succeeded'] }),
    Save_failure: patchInvoice(`@${T('ID')}`, { 'item/Title': `@${T('Title')}`, 'item/PaymentLinkStatus/Value': 'Failed' }, { Create_link: ['Failed', 'TimedOut', 'Skipped'] }),
    Log_failure: logInvoice(`@${T('ID')}`, 'Stripe payment link failed', "@{coalesce(body('Create_link'), body('Create_price'))}", { Save_failure: ['Succeeded'] })
  });

// 12 — Stripe payment received
flow('Brillienta CRM - Stripe payment received',
  'Stripe webhook (event checkout.session.completed). The Checkout Session is re-read from the Stripe API before the invoice is marked Paid, so forged webhook calls are ignored. Requires Power Automate Premium (HTTP trigger).',
  ['sp'], ['site', 'stripeKey'],
  webhookTrigger,
  {
    Session_id: { runAfter: {}, type: 'Compose', inputs: "@{coalesce(triggerBody()?['data']?['object']?['id'], '')}" },
    Is_checkout_event: {
      runAfter: { Session_id: ['Succeeded'] }, type: 'If',
      expression: { and: [{ equals: ["@triggerBody()?['type']", 'checkout.session.completed'] }, { startsWith: ["@outputs('Session_id')", 'cs_'] }] },
      actions: {
        Verify_session: http('GET', "https://api.stripe.com/v1/checkout/sessions/@{outputs('Session_id')}", { Authorization: `Bearer ${Pi('stripeKey')}` }),
        Is_paid: {
          runAfter: { Verify_session: ['Succeeded'] }, type: 'If',
          expression: { and: [{ equals: ["@body('Verify_session')?['payment_status']", 'paid'] }] },
          actions: markPaid("ExternalPaymentId eq '@{body('Verify_session')?['payment_link']}'", "div(float(body('Verify_session')?['amount_total']), 100)", 'Stripe', "@{outputs('Session_id')}", {}),
          else: { actions: {} }
        }
      },
      else: { actions: {} }
    },
    Response: respond(200, { Is_checkout_event: ['Succeeded', 'Failed', 'Skipped', 'TimedOut'] })
  });


// 13 — Email to case
const M = (k) => `triggerOutputs()?['body/${k}']`;
const ESC = (e) => `replace(${e}, , '')`;
flow('Brillienta CRM - Email to case',
  'Watches the support shared mailbox. Replies that carry a case number in the subject (e.g. [CS-2026-0042]) are logged on that case and re-open it; other emails create a new case (contact and account matched by sender), numbered PREFIX-YEAR-ID, with an automatic acknowledgement. The CRM then sets business-hours SLA deadlines.',
  ['sp', 'outlook', 'conv'], ['site', 'page', 'mailbox', 'casePrefix', 'support'],
  {
    When_an_email_arrives: {
      recurrence: { interval: 3, frequency: 'Minute' },
      splitOn: "@triggerOutputs()?['body/value']",
      type: 'OpenApiConnection',
      inputs: {
        host: host('outlook', 'SharedMailboxOnNewEmailV2'),
        parameters: { mailboxAddress: P('mailbox'), folderId: 'Inbox', importance: 'Any', includeAttachments: false },
        authentication: "@parameters('$authentication')"
      }
    }
  },
  {
    Subject: { runAfter: {}, type: 'Compose', inputs: `@{coalesce(${M('subject')}, '(no subject)')}` },
    Sender: { runAfter: {}, type: 'Compose', inputs: `@{toLower(coalesce(${M('from')}, ''))}` },
    Tag: { runAfter: { Subject: ['Succeeded'] }, type: 'Compose', inputs: "@{if(and(contains(outputs('Subject'), '['), contains(outputs('Subject'), ']')), trim(first(split(split(outputs('Subject'), '[')[1], ']'))), '')}" },
    Case_id: { runAfter: {}, type: 'InitializeVariable', inputs: { variables: [{ name: 'caseId', type: 'integer', value: 0 }] } },
    Plain_text: api('conv', 'HtmlToText', { Content: `@{coalesce(${M('body')}, '')}` }, {}),
    Not_automated: {
      runAfter: { Sender: ['Succeeded'], Tag: ['Succeeded'], Case_id: ['Succeeded'], Plain_text: ['Succeeded', 'Failed'] }, type: 'If',
      expression: { and: [
        { not: { equals: ["@outputs('Sender')", `@toLower(${P('mailbox').slice(1)})`] } },
        { not: { contains: ["@outputs('Sender')", 'noreply'] } }, { not: { contains: ["@outputs('Sender')", 'no-reply'] } },
        { not: { contains: ["@outputs('Sender')", 'mailer-daemon'] } },
        { not: { startsWith: ["@toLower(outputs('Subject'))", 'automatic reply'] } }, { not: { startsWith: ["@toLower(outputs('Subject'))", 'out of office'] } }
      ] },
      actions: {
        Email_text: { runAfter: {}, type: 'Compose', inputs: `@{take(coalesce(body('Plain_text'), ${M('bodyPreview')}, ''), 60000)}` },
        Has_tag: {
          runAfter: { Email_text: ['Succeeded'] }, type: 'If', expression: { and: [{ not: { equals: ["@outputs('Tag')", ''] } }] },
          actions: {
            Find_case: getItems('cases', `CaseNumber eq '@{${ESC("outputs('Tag')")}}'`, {}, 1),
            Remember_case: { runAfter: { Find_case: ['Succeeded'] }, type: 'SetVariable', inputs: { name: 'caseId', value: "@coalesce(first(body('Find_case')?['value'])?['ID'], 0)" } }
          },
          else: { actions: {} }
        },
        Existing_case: {
          runAfter: { Has_tag: ['Succeeded'] }, type: 'If', expression: { and: [{ greater: ["@variables('caseId')", 0] }] },
          actions: {
            Log_reply: api('sp', 'PostItem', {
              dataset: P('site'), table: LISTS.acts, 'item/Title': "@{take(concat('Received: ', outputs('Subject')), 255)}",
              'item/Details': "From: @{outputs('Sender')}\n\n@{outputs('Email_text')}", 'item/ActivityType/Value': 'Email',
              'item/RelatedType': 'Case', 'item/RelatedId': "@variables('caseId')", 'item/Completed': true, 'item/DueDate': '@utcNow()',
              'item/ActivitySource/Value': 'Outlook', 'item/ExternalId': `@{take(concat('mail:', coalesce(${M('internetMessageId')}, ${M('id')}), '|cases:', variables('caseId')), 255)}`
            }),
            Needs_reopen: {
              runAfter: { Log_reply: ['Succeeded'] }, type: 'If',
              expression: { and: [{ contains: ["@createArray('Resolved', 'Closed', 'Waiting on customer')", "@first(body('Find_case')?['value'])?['CaseStatus']?['Value']"] }] },
              actions: {
                Reopen: api('sp', 'PatchItem', {
                  dataset: P('site'), table: LISTS.cases, id: "@variables('caseId')", 'item/Title': "@first(body('Find_case')?['value'])?['Title']", 'item/CaseStatus/Value': 'In progress'
                }),
                Tell_owner: email(`@{coalesce(first(body('Find_case')?['value'])?['CaseOwner']?['Email'], ${P('support').slice(1)})}`,
                  "Customer replied: @{outputs('Tag')} @{first(body('Find_case')?['value'])?['Title']}",
                  html("A customer replied to @{outputs('Tag')} — the case is open again", `<table>${tr('From', "@{outputs('Sender')}")}${tr('Subject', "@{outputs('Subject')}")}</table><p>@{take(outputs('Email_text'), 2000)}</p>`),
                  { Reopen: ['Succeeded'] })
              },
              else: { actions: {} }
            }
          },
          else: {
            actions: {
              Find_contact: getItems('contacts', "Email eq '@{" + ESC("outputs('Sender')") + "}'", {}, 1),
              Create_case: api('sp', 'PostItem', {
                dataset: P('site'), table: LISTS.cases, 'item/Title': "@{take(outputs('Subject'), 255)}",
                'item/Priority/Value': `@{if(equals(toLower(coalesce(${M('importance')}, '')), 'high'), 'High', 'Normal')}`,
                'item/CaseStatus/Value': 'New', 'item/Origin/Value': 'Email', 'item/Category/Value': 'Question',
                'item/RequesterEmail': "@{outputs('Sender')}",
                'item/ContactRef': "@first(coalesce(body('Find_contact')?['value'], json('[]')))?['ID']",
                'item/AccountRef': "@first(coalesce(body('Find_contact')?['value'], json('[]')))?['AccountRef']",
                'item/Description': "From: @{outputs('Sender')}\n\n@{outputs('Email_text')}"
              }, { Find_contact: ['Succeeded', 'Failed'] }),
              Case_number: { runAfter: { Create_case: ['Succeeded'] }, type: 'Compose', inputs: `@{concat(${P('casePrefix').slice(1)}, '-', formatDateTime(utcNow(), 'yyyy'), '-', formatNumber(body('Create_case')?['ID'], '0000'))}` },
              Set_number: api('sp', 'PatchItem', {
                dataset: P('site'), table: LISTS.cases, id: "@body('Create_case')?['ID']", 'item/Title': "@body('Create_case')?['Title']", 'item/CaseNumber': "@{outputs('Case_number')}"
              }, { Case_number: ['Succeeded'] }),
              Log_email: api('sp', 'PostItem', {
                dataset: P('site'), table: LISTS.acts, 'item/Title': "@{take(concat('Received: ', outputs('Subject')), 255)}",
                'item/Details': "From: @{outputs('Sender')}\n\n@{outputs('Email_text')}", 'item/ActivityType/Value': 'Email',
                'item/RelatedType': 'Case', 'item/RelatedId': "@body('Create_case')?['ID']", 'item/Completed': true, 'item/DueDate': '@utcNow()',
                'item/ActivitySource/Value': 'Outlook'
              }, { Set_number: ['Succeeded'] }),
              Acknowledge: api('outlook', 'SharedMailboxSendEmailV2', {
                'emailMessage/MailboxAddress': P('mailbox'), 'emailMessage/To': "@{outputs('Sender')}",
                'emailMessage/Subject': "[@{outputs('Case_number')}] @{outputs('Subject')}",
                'emailMessage/Body': `<html><head>${STYLE}</head><body><p style="font-size:18px;font-weight:600;color:#008D62">We have received your request — reference @{outputs('Case_number')}</p><p>Thank you for contacting us. Our support team will get back to you as soon as possible.</p><p>Please keep <b>[@{outputs('Case_number')}]</b> in the subject when you reply, so your message is added to the same case.</p></body></html>`,
                'emailMessage/Importance': 'Normal'
              }, { Set_number: ['Succeeded'] })
            }
          }
        }
      },
      else: { actions: {} }
    }
  });


// 14 — Quote e-signature (signing page served by the flow itself)
const Q = (k) => `body('Get_quote')?['${k}']`;
const H = (e) => `replace(replace(replace(replace(string(coalesce(${e}, '')), '&', '&amp;'), '<', '&lt;'), '>', '&gt;'), '"', '&quot;')`;
const money2 = (e) => `formatNumber(float(coalesce(${e}, 0)), 'N2', 'en-US')`;
const PAGE_CSS = 'body{font-family:Segoe UI,Arial,sans-serif;color:#121317;margin:0;background:#f5f7f6}main{max-width:760px;margin:24px auto;background:#fff;padding:28px;border-radius:10px;box-shadow:0 1px 4px rgba(0,0,0,.08)}h1{font-size:22px;margin:0 0 4px;color:#008D62}.m{color:#5b616e}table{width:100%;border-collapse:collapse;margin:16px 0}th{text-align:left;background:#e6f4ee;padding:8px;font-weight:600}td{padding:8px;border-bottom:1px solid #eee}.n{text-align:right;white-space:nowrap}.tot{font-size:18px;font-weight:700;text-align:right}.terms{white-space:pre-wrap;font-size:13px;color:#333;background:#fafafa;padding:12px;border-radius:6px}label{display:block;margin:12px 0 4px;font-weight:600}input[type=text],textarea{width:100%;box-sizing:border-box;padding:9px;border:1px solid #bbb;border-radius:6px;font:inherit}.chk{display:flex;gap:8px;align-items:flex-start;margin:14px 0;font-weight:400}.btns{display:flex;gap:10px;flex-wrap:wrap;margin-top:16px}button{font:inherit;padding:10px 18px;border-radius:6px;border:1px solid #008D62;cursor:pointer}.ok{background:#008D62;color:#fff}.no{background:#fff;color:#a4262c;border-color:#a4262c}.err{background:#fde7e9;color:#a4262c;padding:10px;border-radius:6px}.done{background:#e6f4ee;padding:14px;border-radius:6px}@media(max-width:600px){main{margin:0;border-radius:0;padding:16px}}';
const page = (title, inner) => `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta name="robots" content="noindex"><title>${title}</title><style>${PAGE_CSS}</style></head><body><main>${inner}</main></body></html>`;
const htmlResponse = (body, runAfter = {}) => ({
  runAfter, type: 'Response', kind: 'Http',
  inputs: { statusCode: 200, headers: { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-store', 'X-Frame-Options': 'DENY', 'Referrer-Policy': 'no-referrer', 'X-Robots-Tag': 'noindex' }, body }
});
const quoteSummary = `<h1>Quote @{${H(Q('QuoteNumber'))}}</h1><p class="m">@{${H(Q('Title'))}}@{if(empty(${Q('ValidUntil')}), '', concat(' · valid until ', formatDateTime(${Q('ValidUntil')}, 'd MMM yyyy')))}</p>
<table><thead><tr><th>Item</th><th class="n">Qty</th><th class="n">Unit price</th><th class="n">Discount</th></tr></thead><tbody>@{join(body('Rows'), '')}</tbody></table>
<p class="tot">Total @{${money2(Q('GrandTotal'))}} @{${H(Q('Currency'))}} <span class="m" style="font-size:13px;font-weight:400">(incl. tax @{${money2(Q('TaxTotal'))}})</span></p>
@{if(empty(coalesce(${Q('Terms')}, '')), '', concat('<p><b>Terms &amp; conditions</b></p><div class="terms">', ${H(Q('Terms'))}, '</div>'))}`;
const signForm = (errorExpr) => page('Review and sign quote', `${quoteSummary}
${errorExpr}
<form method="post" action="" autocomplete="on">
<label for="n">Your full name</label><input type="text" id="n" name="name" maxlength="120" required autocomplete="name">
<label for="t">Job title</label><input type="text" id="t" name="title" maxlength="120" autocomplete="organization-title">
<label for="c">Comments (optional)</label><textarea id="c" name="note" rows="2" maxlength="1000"></textarea>
<label class="chk"><input type="checkbox" name="agree" value="yes" required> I have authority to accept this quote and agree to its terms. Typing my name above is my electronic signature.</label>
<div class="btns"><button class="ok" type="submit" name="decision" value="accept">Accept and sign</button><button class="no" type="submit" name="decision" value="decline" formnovalidate>Decline</button></div>
</form>`);
const formField = (key) => ({ type: 'Query', inputs: { from: "@coalesce(triggerBody()?['$formdata'], json('[]'))", where: `@equals(item()?['key'], '${key}')` } });
const fv = (key) => `trim(string(coalesce(first(body('Form_${key}'))?['value'], '')))`;
const ipExpr = "@{take(coalesce(triggerOutputs()?['headers']?['X-Forwarded-For'], triggerOutputs()?['headers']?['x-forwarded-for'], triggerOutputs()?['headers']?['X-Real-IP'], ''), 255)}";
const uaExpr = "coalesce(triggerOutputs()?['headers']?['User-Agent'], triggerOutputs()?['headers']?['user-agent'], '')";
flow('Brillienta CRM - Quote e-signature',
  'Serves the customer signing page for quotes sent with "Send for e-signature" (single-use link). Records the signer name, job title, IP address, time, comments and a snapshot of the signed quote, then marks the quote Accepted (or Rejected when declined) and emails the quote owner and the signer. The CRM completes the deal (opportunity won, subscriptions) when it is next opened. Requires Power Automate Premium (HTTP trigger).',
  ['sp', 'outlook'], ['site', 'page', 'sales'],
  { manual: { type: 'Request', kind: 'Http', inputs: { triggerAuthenticationType: 'All' } } },
  {
    Quote_id: { runAfter: {}, type: 'Compose', inputs: "@int(coalesce(triggerOutputs()?['queries']?['q'], '0'))" },
    Token: { runAfter: {}, type: 'Compose', inputs: "@{coalesce(triggerOutputs()?['queries']?['t'], '')}" },
    Get_quote: api('sp', 'GetItem', { dataset: P('site'), table: LISTS.quotes, id: "@outputs('Quote_id')" }, { Quote_id: ['Succeeded'] }),
    Link_not_found: htmlResponse(page('Link not valid', '<h1>This link is not valid</h1><p class="m">Please ask your contact for a new link.</p>'), { Get_quote: ['Failed', 'Skipped', 'TimedOut'] }),
    Is_valid: {
      runAfter: { Get_quote: ['Succeeded'], Token: ['Succeeded'] }, type: 'Compose',
      inputs: `@and(not(empty(outputs('Token'))), equals(outputs('Token'), coalesce(${Q('SignatureToken')}, '')), contains(createArray('Sent', 'Viewed'), coalesce(${Q('SignatureStatus')}?['Value'], '')), or(empty(${Q('ValidUntil')}), greaterOrEquals(ticks(addDays(${Q('ValidUntil')}, 1)), ticks(utcNow()))))`
    },
    Rows: {
      runAfter: { Is_valid: ['Succeeded'] }, type: 'Select',
      inputs: {
        from: `@json(coalesce(${Q('LineItems')}, '[]'))`,
        select: `@concat('<tr><td>', ${H("item()?['name']")}, '</td><td class=\"n\">', string(coalesce(item()?['quantity'], 0)), '</td><td class=\"n\">', formatNumber(float(coalesce(item()?['unitPrice'], 0)), 'N2', 'en-US'), '</td><td class=\"n\">', if(greater(float(coalesce(item()?['discountPct'], 0)), 0), concat(string(item()?['discountPct']), '%'), ''), '</td></tr>')`
      }
    },
    Valid_link: {
      runAfter: { Rows: ['Succeeded'] }, type: 'If',
      expression: { and: [{ equals: ["@outputs('Is_valid')", true] }] },
      actions: {
        Is_form_post: {
          runAfter: {}, type: 'If',
          expression: { and: [{ contains: ["@toLower(string(coalesce(triggerOutputs()?['headers']?['Content-Type'], triggerOutputs()?['headers']?['content-type'], '')))", 'application/x-www-form-urlencoded'] }] },
          actions: {
            Form_name: { runAfter: {}, ...formField('name') },
            Form_title: { runAfter: {}, ...formField('title') },
            Form_note: { runAfter: {}, ...formField('note') },
            Form_agree: { runAfter: {}, ...formField('agree') },
            Form_decision: { runAfter: {}, ...formField('decision') },
            Decision: {
              runAfter: { Form_name: ['Succeeded'], Form_title: ['Succeeded'], Form_note: ['Succeeded'], Form_agree: ['Succeeded'], Form_decision: ['Succeeded'] }, type: 'Compose',
              inputs: `@if(equals(${fv('decision')}, 'decline'), 'decline', if(and(equals(${fv('decision')}, 'accept'), equals(${fv('agree')}, 'yes'), greaterOrEquals(length(${fv('name')}), 2)), 'accept', 'incomplete'))`
            },
            Accepted: {
              runAfter: { Decision: ['Succeeded'] }, type: 'If', expression: { and: [{ equals: ["@outputs('Decision')", 'accept'] }] },
              actions: {
                Signature_ref: { runAfter: {}, type: 'Compose', inputs: "@{toUpper(take(replace(guid(), '-', ''), 16))}" },
                Save_signature: api('sp', 'PatchItem', {
                  dataset: P('site'), table: LISTS.quotes, id: "@outputs('Quote_id')", 'item/Title': `@${Q('Title')}`,
                  'item/QuoteStatus/Value': 'Accepted', 'item/SignatureStatus/Value': 'Signed',
                  'item/SignatureName': `@{take(${fv('name')}, 255)}`, 'item/SignatureTitle': `@{take(${fv('title')}, 255)}`,
                  'item/SignatureNote': `@{${fv('note')}}`, 'item/SignedDate': '@utcNow()', 'item/SignatureIp': ipExpr,
                  'item/SignatureId': "@{outputs('Signature_ref')}", 'item/SignatureToken': '',
                  'item/SignedSnapshot': `@{concat('Quote ', ${Q('QuoteNumber')}, ' | total ', string(${Q('GrandTotal')}), ' ', coalesce(${Q('Currency')}, ''), ' | valid until ', coalesce(${Q('ValidUntil')}, ''), ' | signer ', ${fv('name')}, ' <', coalesce(${Q('SignatureEmail')}, ''), '> | ip ', coalesce(triggerOutputs()?['headers']?['X-Forwarded-For'], ''), ' | user agent ', ${uaExpr}, ' | signed ', utcNow(), ' | lines ', coalesce(${Q('LineItems')}, '[]'), ' | terms ', coalesce(${Q('Terms')}, ''))}`
                }, { Signature_ref: ['Succeeded'] }),
                Log_signature: api('sp', 'PostItem', {
                  dataset: P('site'), table: LISTS.acts, 'item/Title': `@{take(concat('Quote signed by ', ${fv('name')}), 255)}`,
                  'item/Details': `@{concat(${fv('name')}, if(empty(${fv('title')}), '', concat(', ', ${fv('title')})), ' accepted quote ', ${Q('QuoteNumber')}, ' electronically. Reference ', outputs('Signature_ref'), '. ', ${fv('note')})}`,
                  'item/ActivityType/Value': 'Note', 'item/RelatedType': 'Quote', 'item/RelatedId': "@outputs('Quote_id')", 'item/Completed': true, 'item/ActivitySource/Value': 'Flow'
                }, { Save_signature: ['Succeeded'] }),
                Tell_owner_signed: email(`@{coalesce(${Q('QuoteOwner')}?['Email'], ${P('sales').slice(1)})}`,
                  `✍ Quote @{${Q('QuoteNumber')}} signed by @{${fv('name')}}`,
                  html(`Quote @{${Q('QuoteNumber')}} was accepted and signed`, `<table>${tr('Signed by', `@{${H(fv('name'))}} @{${H(fv('title'))}}`)}${tr('Email', `@{${H(Q('SignatureEmail'))}}`)}${tr('Total', `@{${money2(Q('GrandTotal'))}} @{${H(Q('Currency'))}}`)}${tr('Comments', `@{${H(fv('note'))}}`)}${tr('Reference', "@{outputs('Signature_ref')}")}</table><p>Open the CRM to close the deal and create subscriptions automatically.</p>`),
                  { Save_signature: ['Succeeded'] }),
                Confirm_signer: email(`@{${Q('SignatureEmail')}}`, `Your acceptance of quote @{${Q('QuoteNumber')}}`,
                  `<html><head>${STYLE}</head><body><p style="font-size:18px;font-weight:600;color:#008D62">Thank you — quote @{${H(Q('QuoteNumber'))}} is accepted</p><table>${tr('Quote', `@{${H(Q('Title'))}}`)}${tr('Total', `@{${money2(Q('GrandTotal'))}} @{${H(Q('Currency'))}}`)}${tr('Signed by', `@{${H(fv('name'))}}`)}${tr('Signed on (UTC)', "@{formatDateTime(utcNow(), 'd MMM yyyy HH:mm')}")}${tr('Reference', "@{outputs('Signature_ref')}")}</table><p>Keep this email for your records.</p></body></html>`,
                  { Save_signature: ['Succeeded'] }),
                Thank_you: htmlResponse(page('Quote accepted', `<h1>Thank you</h1><div class="done">Quote <b>@{${H(Q('QuoteNumber'))}}</b> is accepted and signed by <b>@{${H(fv('name'))}}</b>.<br>Reference @{outputs('Signature_ref')} · a confirmation email is on its way.</div>`), { Save_signature: ['Succeeded'] })
              },
              else: {
                actions: {
                  Declined: {
                    runAfter: {}, type: 'If', expression: { and: [{ equals: ["@outputs('Decision')", 'decline'] }] },
                    actions: {
                      Save_decline: api('sp', 'PatchItem', {
                        dataset: P('site'), table: LISTS.quotes, id: "@outputs('Quote_id')", 'item/Title': `@${Q('Title')}`,
                        'item/QuoteStatus/Value': 'Rejected', 'item/SignatureStatus/Value': 'Declined', 'item/SignatureToken': '',
                        'item/SignatureName': `@{take(${fv('name')}, 255)}`, 'item/SignatureNote': `@{${fv('note')}}`, 'item/SignedDate': '@utcNow()', 'item/SignatureIp': ipExpr
                      }),
                      Tell_owner_declined: email(`@{coalesce(${Q('QuoteOwner')}?['Email'], ${P('sales').slice(1)})}`, `Quote @{${Q('QuoteNumber')}} was declined`,
                        html(`Quote @{${Q('QuoteNumber')}} was declined`, `<table>${tr('By', `@{${H(fv('name'))}} @{${H(Q('SignatureEmail'))}}`)}${tr('Comments', `@{${H(fv('note'))}}`)}</table>`), { Save_decline: ['Succeeded'] }),
                      Declined_page: htmlResponse(page('Quote declined', '<h1>Thank you for letting us know</h1><div class="done">We have told your contact that you declined this quote.</div>'), { Save_decline: ['Succeeded'] })
                    },
                    else: { actions: { Incomplete_page: htmlResponse(signForm('<p class="err">Please enter your full name and tick the box to accept.</p>')) } }
                  }
                }
              }
            }
          },
          else: {
            actions: {
              First_view: {
                runAfter: {}, type: 'If', expression: { and: [{ equals: [`@coalesce(${Q('SignatureStatus')}?['Value'], '')`, 'Sent'] }] },
                actions: {
                  Mark_viewed: api('sp', 'PatchItem', { dataset: P('site'), table: LISTS.quotes, id: "@outputs('Quote_id')", 'item/Title': `@${Q('Title')}`, 'item/SignatureStatus/Value': 'Viewed' })
                },
                else: { actions: {} }
              },
              Signing_page: htmlResponse(signForm(''), { First_view: ['Succeeded', 'Failed'] })
            }
          }
        }
      },
      else: {
        actions: {
          Link_used: htmlResponse(page('Link not valid', `<h1>@{if(equals(coalesce(${Q('SignatureStatus')}?['Value'], ''), 'Signed'), 'This quote is already signed', 'This link is no longer valid')}</h1><p class="m">@{if(equals(coalesce(${Q('SignatureStatus')}?['Value'], ''), 'Signed'), concat('Signed by ', ${H(Q('SignatureName'))}, ' on ', formatDateTime(${Q('SignedDate')}, 'd MMM yyyy'), '.'), 'The quote may have expired, been answered or been sent again with a new link. Please ask your contact for a new link.')}</p>`))
        }
      }
    }
  });


// ------------------------------------------------------------------ WhatsApp Business Cloud API
const A = (k) => `triggerOutputs()?['body/${k}']`;
const GRAPH = `https://graph.facebook.com/v21.0/${Pi('waPhoneId')}/messages`;
const waHeaders = { Authorization: `Bearer ${Pi('waToken')}`, 'Content-Type': 'application/json' };
const actPatch = (fields, runAfter = {}) => api('sp', 'PatchItem', { dataset: P('site'), table: LISTS.acts, id: `@${A('ID')}`, 'item/Title': `@${A('Title')}`, ...fields }, runAfter);

// 15 — WhatsApp send
flow('Brillienta CRM - WhatsApp send',
  'Sends WhatsApp messages queued in the CRM (WhatsApp › Send) through the WhatsApp Business Cloud API: approved templates with their values, or free text inside the 24-hour customer service window. Writes Sent/Failed back to the activity. Requires Power Automate Premium (HTTP).',
  ['sp'], ['site', 'waPhoneId', 'waToken'],
  { When_a_message_is_queued: spPoll('GetOnUpdatedItems', 'acts', [`@equals(${A('WaStatus/Value')}, 'Queued')`]) },
  {
    Template_values: {
      runAfter: {}, type: 'Select',
      inputs: { from: `@json(coalesce(${A('WaParams')}, '[]'))`, select: { type: 'text', text: "@{take(string(item()), 1000)}" } }
    },
    Uses_template: {
      runAfter: { Template_values: ['Succeeded'] }, type: 'If', expression: { and: [{ not: { equals: [`@coalesce(${A('WaTemplate')}, '')`, ''] } }] },
      actions: {
        Send_template: http('POST', GRAPH, waHeaders, {
          messaging_product: 'whatsapp', to: `@{${A('WaTo')}}`, type: 'template',
          template: {
            name: `@{first(split(${A('WaTemplate')}, '|'))}`, language: { code: `@{last(split(${A('WaTemplate')}, '|'))}` },
            components: [{ type: 'body', parameters: "@body('Template_values')" }]
          }
        })
      },
      else: {
        actions: {
          Send_text: http('POST', GRAPH, waHeaders, {
            messaging_product: 'whatsapp', to: `@{${A('WaTo')}}`, type: 'text', text: { preview_url: true, body: `@{take(coalesce(${A('Details')}, ''), 4096)}` }
          })
        }
      }
    },
    Message_id: { runAfter: { Uses_template: ['Succeeded'] }, type: 'Compose', inputs: "@{coalesce(body('Send_template')?['messages']?[0]?['id'], body('Send_text')?['messages']?[0]?['id'], '')}" },
    Mark_sent: actPatch({ 'item/WaStatus/Value': 'Sent', 'item/Completed': true, 'item/ExternalId': "@{concat('wa:', outputs('Message_id'))}" }, { Message_id: ['Succeeded'] }),
    Mark_failed: actPatch({
      'item/WaStatus/Value': 'Failed',
      'item/Details': `@{concat(coalesce(${A('Details')}, ''), '\n\n— Not sent: ', coalesce(body('Send_template')?['error']?['message'], body('Send_text')?['error']?['message'], 'see the flow run history'))}`
    }, { Uses_template: ['Failed', 'TimedOut'] })
  });

// 16 — WhatsApp inbound
const V = "triggerBody()?['entry']?[0]?['changes']?[0]?['value']";
flow('Brillienta CRM - WhatsApp inbound',
  'Webhook for the WhatsApp Business Cloud API (Meta › WhatsApp › Configuration). Answers the verification request, logs incoming messages on the matching contact or lead (unknown numbers become new WhatsApp leads) and updates delivery status (sent, delivered, read, failed) of messages sent from the CRM. Requires Power Automate Premium (HTTP trigger).',
  ['sp'], ['site', 'waPhoneId', 'waVerify'],
  { manual: { type: 'Request', kind: 'Http', inputs: { triggerAuthenticationType: 'All' } } },
  {
    Is_verification: {
      runAfter: {}, type: 'If', expression: { and: [{ equals: ["@coalesce(triggerOutputs()?['queries']?['hub.mode'], '')", 'subscribe'] }] },
      actions: {
        Token_ok: {
          runAfter: {}, type: 'If', expression: { and: [{ equals: ["@coalesce(triggerOutputs()?['queries']?['hub.verify_token'], '')", P('waVerify')] }, { not: { equals: [P('waVerify'), 'not-set'] } }] },
          actions: { Challenge: { runAfter: {}, type: 'Response', kind: 'Http', inputs: { statusCode: 200, headers: { 'Content-Type': 'text/plain' }, body: "@{triggerOutputs()?['queries']?['hub.challenge']}" } } },
          else: { actions: { Forbidden: respond(403, {}) } }
        }
      },
      else: {
        actions: {
          Our_number: {
            runAfter: {}, type: 'If', expression: { and: [{ equals: [`@coalesce(${V}?['metadata']?['phone_number_id'], '')`, P('waPhoneId')] }] },
            actions: {
              For_each_message: {
                runAfter: {}, type: 'Foreach', foreach: `@coalesce(${V}?['messages'], json('[]'))`,
                actions: {
                  Sender: { runAfter: {}, type: 'Compose', inputs: "@{replace(coalesce(items('For_each_message')?['from'], ''), '''', '')}" },
                  Text: { runAfter: {}, type: 'Compose', inputs: "@{coalesce(items('For_each_message')?['text']?['body'], items('For_each_message')?['button']?['text'], items('For_each_message')?['interactive']?['button_reply']?['title'], items('For_each_message')?['interactive']?['list_reply']?['title'], items('For_each_message')?['image']?['caption'], items('For_each_message')?['document']?['filename'], concat('[', items('For_each_message')?['type'], ' message]'))}" },
                  Profile_name: { runAfter: {}, type: 'Compose', inputs: `@{coalesce(first(coalesce(${V}?['contacts'], json('[]')))?['profile']?['name'], '')}` },
                  Find_contact: getItems('contacts', "Whatsapp eq '@{outputs('Sender')}'", { Sender: ['Succeeded'] }, 1),
                  Find_lead: getItems('leads', "Whatsapp eq '@{outputs('Sender')}'", { Sender: ['Succeeded'] }, 1),
                  Target: {
                    runAfter: { Find_contact: ['Succeeded'], Find_lead: ['Succeeded'], Text: ['Succeeded'], Profile_name: ['Succeeded'] }, type: 'Compose',
                    inputs: "@if(greater(length(body('Find_contact')?['value']), 0), concat('Contact|', string(first(body('Find_contact')?['value'])?['ID'])), if(greater(length(body('Find_lead')?['value']), 0), concat('Lead|', string(first(body('Find_lead')?['value'])?['ID'])), 'new'))"
                  },
                  Known_sender: {
                    runAfter: { Target: ['Succeeded'] }, type: 'If', expression: { and: [{ not: { equals: ["@outputs('Target')", 'new'] } }] },
                    actions: {
                      Log_message: api('sp', 'PostItem', {
                        dataset: P('site'), table: LISTS.acts, 'item/Title': "@{take(concat('WhatsApp from +', outputs('Sender'), ': ', outputs('Text')), 255)}",
                        'item/Details': "@{outputs('Text')}", 'item/ActivityType/Value': 'WhatsApp', 'item/RelatedType': "@{first(split(outputs('Target'), '|'))}",
                        'item/RelatedId': "@int(last(split(outputs('Target'), '|')))", 'item/Completed': true, 'item/DueDate': '@utcNow()',
                        'item/ActivitySource/Value': 'WhatsApp', 'item/WaStatus/Value': 'Received', 'item/WaTo': "@{outputs('Sender')}",
                        'item/ExternalId': "@{concat('wa:', items('For_each_message')?['id'])}"
                      })
                    },
                    else: {
                      actions: {
                        Create_lead: api('sp', 'PostItem', {
                          dataset: P('site'), table: LISTS.leads,
                          'item/Title': "@{if(empty(outputs('Profile_name')), concat('WhatsApp +', outputs('Sender')), outputs('Profile_name'))}",
                          'item/FirstName': "@{first(split(if(empty(outputs('Profile_name')), 'WhatsApp', outputs('Profile_name')), ' '))}",
                          'item/LastName': "@{if(contains(outputs('Profile_name'), ' '), trim(substring(outputs('Profile_name'), indexOf(outputs('Profile_name'), ' '))), '')}",
                          'item/Phone': "@{concat('+', outputs('Sender'))}", 'item/Whatsapp': "@{outputs('Sender')}",
                          'item/LeadSource/Value': 'WhatsApp', 'item/Status/Value': 'New', 'item/Rating/Value': 'Warm',
                          'item/Description': "@{concat('First WhatsApp message: ', outputs('Text'))}"
                        }),
                        Log_first_message: api('sp', 'PostItem', {
                          dataset: P('site'), table: LISTS.acts, 'item/Title': "@{take(concat('WhatsApp from +', outputs('Sender'), ': ', outputs('Text')), 255)}",
                          'item/Details': "@{outputs('Text')}", 'item/ActivityType/Value': 'WhatsApp', 'item/RelatedType': 'Lead',
                          'item/RelatedId': "@body('Create_lead')?['ID']", 'item/Completed': true, 'item/DueDate': '@utcNow()',
                          'item/ActivitySource/Value': 'WhatsApp', 'item/WaStatus/Value': 'Received', 'item/WaTo': "@{outputs('Sender')}",
                          'item/ExternalId': "@{concat('wa:', items('For_each_message')?['id'])}"
                        }, { Create_lead: ['Succeeded'] })
                      }
                    }
                  }
                }
              },
              For_each_status: {
                runAfter: { For_each_message: ['Succeeded', 'Failed'] }, type: 'Foreach', foreach: `@coalesce(${V}?['statuses'], json('[]'))`,
                actions: {
                  Find_sent: getItems('acts', "ExternalId eq 'wa:@{replace(items('For_each_status')?['id'], '''', '')}'", {}, 1),
                  For_the_activity: {
                    runAfter: { Find_sent: ['Succeeded'] }, type: 'Foreach', foreach: "@body('Find_sent')?['value']",
                    actions: {
                      Update_status: api('sp', 'PatchItem', {
                        dataset: P('site'), table: LISTS.acts, id: "@items('For_the_activity')?['ID']", 'item/Title': "@items('For_the_activity')?['Title']",
                        'item/WaStatus/Value': "@{if(equals(items('For_each_status')?['status'], 'read'), 'Read', if(equals(items('For_each_status')?['status'], 'delivered'), 'Delivered', if(equals(items('For_each_status')?['status'], 'failed'), 'Failed', 'Sent')))}"
                      })
                    }
                  }
                }
              }
            },
            else: { actions: {} }
          },
          Accepted: respond(200, { Our_number: ['Succeeded', 'Failed', 'Skipped', 'TimedOut'] })
        }
      }
    }
  });


// 17 — Case satisfaction survey (page served by the flow)
const C = (k) => `body('Get_case')?['${k}']`;
flow('Brillienta CRM - Case satisfaction survey',
  'Serves the one-click 1–5 satisfaction survey that the CRM emails when a case is resolved. Saves the score on the case immediately, then lets the customer add a comment. Low scores (1–2) alert the support manager. Requires Power Automate Premium (HTTP trigger).',
  ['sp', 'outlook'], ['site', 'page', 'support'],
  { manual: { type: 'Request', kind: 'Http', inputs: { triggerAuthenticationType: 'All' } } },
  {
    Case_id: { runAfter: {}, type: 'Compose', inputs: "@int(coalesce(triggerOutputs()?['queries']?['q'], '0'))" },
    Rating: { runAfter: {}, type: 'Compose', inputs: "@if(contains(createArray('1', '2', '3', '4', '5'), coalesce(triggerOutputs()?['queries']?['r'], '')), int(triggerOutputs()?['queries']?['r']), 0)" },
    Get_case: api('sp', 'GetItem', { dataset: P('site'), table: LISTS.cases, id: "@outputs('Case_id')" }, { Case_id: ['Succeeded'] }),
    Case_not_found: htmlResponse(page('Link not valid', '<h1>This link is not valid</h1>'), { Get_case: ['Failed', 'Skipped', 'TimedOut'] }),
    Valid_link: {
      runAfter: { Get_case: ['Succeeded'], Rating: ['Succeeded'] }, type: 'If',
      expression: { and: [{ not: { equals: [`@coalesce(${C('CsatToken')}, '')`, ''] } }, { equals: ["@coalesce(triggerOutputs()?['queries']?['t'], '')", `@coalesce(${C('CsatToken')}, '')`] }] },
      actions: {
        Is_comment: {
          runAfter: {}, type: 'If',
          expression: { and: [{ contains: ["@toLower(string(coalesce(triggerOutputs()?['headers']?['Content-Type'], triggerOutputs()?['headers']?['content-type'], '')))", 'application/x-www-form-urlencoded'] }] },
          actions: {
            Form_comment: { runAfter: {}, ...formField('comment') },
            Save_comment: api('sp', 'PatchItem', {
              dataset: P('site'), table: LISTS.cases, id: "@outputs('Case_id')", 'item/Title': `@${C('Title')}`,
              'item/CsatComment': "@{take(trim(string(coalesce(first(body('Form_comment'))?['value'], ''))), 2000)}"
            }, { Form_comment: ['Succeeded'] }),
            Thanks_comment: htmlResponse(page('Thank you', '<h1>Thank you for your feedback</h1><div class="done">We read every comment and use it to improve our support.</div>'), { Save_comment: ['Succeeded', 'Failed'] })
          },
          else: {
            actions: {
              Has_rating: {
                runAfter: {}, type: 'If', expression: { and: [{ greater: ["@outputs('Rating')", 0] }] },
                actions: {
                  Save_rating: api('sp', 'PatchItem', { dataset: P('site'), table: LISTS.cases, id: "@outputs('Case_id')", 'item/Title': `@${C('Title')}`, 'item/CsatScore': "@outputs('Rating')" }),
                  Low_score: {
                    runAfter: { Save_rating: ['Succeeded'] }, type: 'If', expression: { and: [{ lessOrEquals: ["@outputs('Rating')", 2] }] },
                    actions: {
                      Alert_support: email(P('support'), `Low satisfaction (@{outputs('Rating')}/5) on @{${C('CaseNumber')}}`,
                        html(`A customer rated case @{${C('CaseNumber')}} @{outputs('Rating')}/5`, `<table>${tr('Case', `@{${H(C('Title'))}}`)}${tr('Owner', `@{coalesce(${C('CaseOwner')}?['DisplayName'], 'Unassigned')}`)}</table><p>Consider following up with the customer.</p>`),
                        {}, { 'emailMessage/Cc': `@{coalesce(${C('CaseOwner')}?['Email'], '')}`, 'emailMessage/Importance': 'High' })
                    },
                    else: { actions: {} }
                  }
                },
                else: { actions: {} }
              },
              Comment_page: htmlResponse(page('Thank you', `<h1>Thank you!</h1><p class="m">@{if(greater(outputs('Rating'), 0), concat('You rated case ', ${H(C('CaseNumber'))}, ' ', string(outputs('Rating')), ' out of 5.'), '')}</p>
<form method="post" action=""><label for="c">Anything we could do better? (optional)</label><textarea id="c" name="comment" rows="4" maxlength="2000"></textarea><div class="btns"><button class="ok" type="submit">Send feedback</button></div></form>`), { Has_rating: ['Succeeded', 'Failed'] })
            }
          }
        }
      },
      else: { actions: { Link_invalid: htmlResponse(page('Link not valid', '<h1>This survey link is no longer valid</h1>')) } }
    }
  });


// ------------------------------------------------------------------ SharePoint REST helper (generic list access)
const spHttp = (method, uri, body, runAfter = {}, extraHeaders = {}) => api('sp', 'HttpRequest', {
  dataset: P('site'), 'parameters/method': method, 'parameters/uri': uri,
  'parameters/headers': { Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata', ...extraHeaders },
  ...(body !== undefined ? { 'parameters/body': body } : {})
}, runAfter);
const jsonResponse = (status, body, runAfter = {}) => ({ runAfter, type: 'Response', kind: 'Http', inputs: { statusCode: status, headers: { 'Content-Type': 'application/json' }, body } });

// 18 — Unsubscribe
const UL = "@{if(startsWith(coalesce(triggerOutputs()?['queries']?['e'], ''), 'L'), 'BrillCRM_Leads', 'BrillCRM_Contacts')}";
flow('Brillienta CRM - Unsubscribe',
  'Personal unsubscribe page for CRM marketing emails. Shows a confirmation button (so mail scanners cannot unsubscribe people by opening the link), then marks the lead or contact Unsubscribed, removes marketing consent and logs it on the timeline. Requires Power Automate Premium (HTTP trigger).',
  ['sp'], ['site'],
  { manual: { type: 'Request', kind: 'Http', inputs: { triggerAuthenticationType: 'All' } } },
  {
    Person_id: { runAfter: {}, type: 'Compose', inputs: "@int(substring(concat(coalesce(triggerOutputs()?['queries']?['e'], 'X0'), ''), 1))" },
    List_name: { runAfter: {}, type: 'Compose', inputs: UL },
    Get_person: spHttp('GET', "_api/web/lists/getbytitle('@{outputs('List_name')}')/items(@{outputs('Person_id')})?$select=Id,Title,UnsubToken", undefined, { Person_id: ['Succeeded'], List_name: ['Succeeded'] }),
    Not_found: htmlResponse(page('Link not valid', '<h1>This link is not valid</h1>'), { Get_person: ['Failed', 'Skipped', 'TimedOut'] }),
    Token_ok: {
      runAfter: { Get_person: ['Succeeded'] }, type: 'If',
      expression: { and: [{ not: { equals: ["@coalesce(body('Get_person')?['UnsubToken'], '')", ''] } }, { equals: ["@coalesce(triggerOutputs()?['queries']?['t'], '')", "@coalesce(body('Get_person')?['UnsubToken'], '')"] }] },
      actions: {
        Is_confirm: {
          runAfter: {}, type: 'If', expression: { and: [{ contains: ["@toLower(string(coalesce(triggerOutputs()?['headers']?['Content-Type'], triggerOutputs()?['headers']?['content-type'], '')))", 'application/x-www-form-urlencoded'] }] },
          actions: {
            Unsubscribe: spHttp('POST', "_api/web/lists/getbytitle('@{outputs('List_name')}')/items(@{outputs('Person_id')})", { Unsubscribed: true, MarketingConsent: false }, {}, { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' }),
            Log_unsubscribe: api('sp', 'PostItem', {
              dataset: P('site'), table: LISTS.acts, 'item/Title': 'Consent: Unsubscribed from marketing (email link)', 'item/ActivityType/Value': 'Note',
              'item/RelatedType': "@{if(startsWith(triggerOutputs()?['queries']?['e'], 'L'), 'Lead', 'Contact')}", 'item/RelatedId': "@outputs('Person_id')", 'item/Completed': true,
              'item/Details': "Unsubscribed via the link in a marketing email on @{utcNow()}.", 'item/ActivitySource/Value': 'Flow'
            }, { Unsubscribe: ['Succeeded'] }),
            Done_page: htmlResponse(page('Unsubscribed', '<h1>You are unsubscribed</h1><div class="done">You will no longer receive marketing emails from us. Service messages about your account or support requests may still be sent.</div>'), { Unsubscribe: ['Succeeded'] })
          },
          else: { actions: { Confirm_page: htmlResponse(page('Unsubscribe', '<h1>Unsubscribe from marketing emails?</h1><form method="post" action=""><input type="hidden" name="confirm" value="yes"><div class="btns"><button class="ok" type="submit">Unsubscribe</button></div></form>')) } }
        }
      },
      else: { actions: { Bad_token: htmlResponse(page('Link not valid', '<h1>This link is not valid</h1><p class="m">Please use the link in the most recent email.</p>')) } }
    }
  });

// 19 — Inbound API
const B = (k) => `triggerBody()?['${k}']`;
const LIST_MAP = "{\"leads\":\"BrillCRM_Leads\",\"contacts\":\"BrillCRM_Contacts\",\"accounts\":\"BrillCRM_Accounts\",\"opportunities\":\"BrillCRM_Opportunities\",\"cases\":\"BrillCRM_Cases\",\"activities\":\"BrillCRM_Activities\",\"demos\":\"BrillCRM_DemoRequests\",\"products\":\"BrillCRM_Products\",\"subscriptions\":\"BrillCRM_Subscriptions\",\"invoices\":\"BrillCRM_Invoices\"}";
const MATCH_FIELDS = "createArray('Email', 'Title', 'Id', 'ExternalId', 'Whatsapp', 'CaseNumber', 'InvoiceNumber', 'QuoteNumber', 'Sku', 'Phone')";
flow('Brillienta CRM - Inbound API',
  'REST endpoint for other systems (ERP, billing, web apps, Zapier/Make). POST JSON {"entity": "leads", "action": "create" | "upsert" | "get", "match": {"field": "Email", "value": "x@y.com"}, "fields": {SharePoint column: value}} with the header x-api-key. Returns {"id": …, "action": …}. Requires Power Automate Premium (HTTP trigger).',
  ['sp'], ['site', 'apiKey'],
  {
    manual: {
      type: 'Request', kind: 'Http',
      inputs: {
        triggerAuthenticationType: 'All', method: 'POST',
        schema: { type: 'object', properties: { entity: { type: 'string' }, action: { type: 'string' }, match: { type: 'object', properties: { field: { type: 'string' }, value: {} } }, fields: { type: 'object' } } }
      }
    }
  },
  {
    Authorised: {
      runAfter: {}, type: 'If',
      expression: { and: [{ not: { equals: [P('apiKey'), 'not-set'] } }, { equals: ["@coalesce(triggerOutputs()?['headers']?['x-api-key'], triggerOutputs()?['headers']?['X-Api-Key'], '')", P('apiKey')] }] },
      actions: {
        List_title: { runAfter: {}, type: 'Compose', inputs: `@coalesce(json('${LIST_MAP}')?[toLower(coalesce(${B('entity')}, ''))], '')` },
        Action: { runAfter: {}, type: 'Compose', inputs: `@{toLower(coalesce(${B('action')}, 'create'))}` },
        Match_field: { runAfter: {}, type: 'Compose', inputs: `@{coalesce(${B('match')}?['field'], '')}` },
        Valid_request: {
          runAfter: { List_title: ['Succeeded'], Action: ['Succeeded'], Match_field: ['Succeeded'] }, type: 'If',
          expression: { and: [
            { not: { equals: ["@outputs('List_title')", ''] } },
            { contains: ["@createArray('create', 'upsert', 'get')", "@outputs('Action')"] },
            { or: [{ equals: ["@outputs('Action')", 'create'] }, { contains: [`@${MATCH_FIELDS}`, "@outputs('Match_field')"] }] }
          ] },
          actions: {
            Needs_lookup: {
              runAfter: {}, type: 'If', expression: { and: [{ not: { equals: ["@outputs('Action')", 'create'] } }] },
              actions: {
                Find: spHttp('GET', `_api/web/lists/getbytitle('@{outputs('List_title')}')/items?$top=1&$filter=@{outputs('Match_field')} eq @{if(equals(outputs('Match_field'), 'Id'), string(${B('match')}?['value']), concat('''', replace(string(${B('match')}?['value']), '''', ''''''), ''''))}`)
              },
              else: { actions: {} }
            },
            Found: { runAfter: { Needs_lookup: ['Succeeded'] }, type: 'Compose', inputs: "@first(coalesce(body('Find')?['value'], json('[]')))" },
            Route: {
              runAfter: { Found: ['Succeeded'] }, type: 'If', expression: { and: [{ equals: ["@outputs('Action')", 'get'] }] },
              actions: {
                Get_result: {
                  runAfter: {}, type: 'If', expression: { and: [{ not: { equals: ["@outputs('Found')", null] } }] },
                  actions: { Return_item: jsonResponse(200, "@outputs('Found')") },
                  else: { actions: { Return_404: jsonResponse(404, { error: 'not found' }) } }
                }
              },
              else: {
                actions: {
                  Update_or_create: {
                    runAfter: {}, type: 'If', expression: { and: [{ equals: ["@outputs('Action')", 'upsert'] }, { not: { equals: ["@outputs('Found')", null] } }] },
                    actions: {
                      Update_item: spHttp('POST', "_api/web/lists/getbytitle('@{outputs('List_title')}')/items(@{outputs('Found')?['Id']})", `@coalesce(${B('fields')}, json('{}'))`, {}, { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' }),
                      Updated: jsonResponse(200, { id: "@outputs('Found')?['Id']", action: 'updated' }, { Update_item: ['Succeeded'] }),
                      Update_failed: jsonResponse(400, { error: "@{coalesce(body('Update_item')?['odata.error']?['message']?['value'], 'update failed')}" }, { Update_item: ['Failed', 'TimedOut'] })
                    },
                    else: {
                      actions: {
                        Create_item: spHttp('POST', "_api/web/lists/getbytitle('@{outputs('List_title')}')/items", `@coalesce(${B('fields')}, json('{}'))`),
                        Created: jsonResponse(201, { id: "@body('Create_item')?['Id']", action: 'created' }, { Create_item: ['Succeeded'] }),
                        Create_failed: jsonResponse(400, { error: "@{coalesce(body('Create_item')?['odata.error']?['message']?['value'], 'create failed')}" }, { Create_item: ['Failed', 'TimedOut'] })
                      }
                    }
                  }
                }
              }
            }
          },
          else: { actions: { Bad_request: jsonResponse(400, { error: 'entity must be one of leads, contacts, accounts, opportunities, cases, activities, demos, products, subscriptions, invoices; action create, upsert or get; match.field one of Email, Title, Id, ExternalId, Whatsapp, CaseNumber, InvoiceNumber, QuoteNumber, Sku, Phone' }) } }
        }
      },
      else: { actions: { Unauthorised: jsonResponse(401, { error: 'missing or wrong x-api-key' }) } }
    }
  });


// 20 — Record security (item-level permissions by owner and team)
const SEC_LISTS = [
  { list: 'BrillCRM_Leads', owner: 'LeadOwner', cases: false }, { list: 'BrillCRM_Accounts', owner: 'AccountOwner', cases: false },
  { list: 'BrillCRM_Contacts', owner: 'ContactOwner', cases: false }, { list: 'BrillCRM_Opportunities', owner: 'DealOwner', cases: false },
  { list: 'BrillCRM_Quotes', owner: 'QuoteOwner', cases: false }, { list: 'BrillCRM_Cases', owner: 'CaseOwner', cases: true },
  { list: 'BrillCRM_Subscriptions', owner: 'SubscriptionOwner', cases: false }, { list: 'BrillCRM_Invoices', owner: 'InvoiceOwner', cases: false }
];
const S = (k) => `outputs('Settings')?['${k}']`;
const grp = (name, key) => ({ [`Group_${name}`]: spHttp('GET', `_api/web/sitegroups/getbyname('@{replace(${S('groups')}?['${key}'], '''', '''''')}')?$select=Id`) });
const role = (name, type) => ({ [`Role_${name}`]: spHttp('GET', `_api/web/roledefinitions/getbytype(${type})?$select=Id`) });
const ITEM = "_api/web/lists/getbytitle('@{items('For_each_list')?['list']}')/items(@{items('For_each_item')?['Id']})";
const OWNER = "items('For_each_item')?[concat(items('For_each_list')?['owner'], 'Id')]";
const addRole = (principal, roleExpr, runAfter) => spHttp('POST', `${ITEM}/roleassignments/addroleassignment(principalid=@{${principal}},roledefid=@{${roleExpr}})`, undefined, runAfter);
flow('Brillienta CRM - Record security',
  'Every 15 minutes: for records created or reassigned since the last runs, gives each lead, account, contact, deal, quote, case, subscription and invoice its own SharePoint permissions (owner, owner\'s team manager and team per CRM settings, CRM groups). Runs only when Settings › Security › Record-level security is on. The connection must be a site owner (Full Control).',
  ['sp'], ['site'],
  { Every_15_minutes: { recurrence: { frequency: 'Minute', interval: 15 }, type: 'Recurrence' } },
  {
    Get_settings: getItems('settingsList', "Title eq 'settings'", {}, 1),
    Settings: { runAfter: { Get_settings: ['Succeeded'] }, type: 'Compose', inputs: "@json(coalesce(first(body('Get_settings')?['value'])?['Value'], '{}'))" },
    Enabled: {
      runAfter: { Settings: ['Succeeded'] }, type: 'If', expression: { and: [{ equals: [`@coalesce(${S('security')}?['itemLevel'], false)`, true] }] },
      actions: {
        ...grp('admins', 'admins'), ...grp('managers', 'managers'), ...grp('support', 'support'), ...grp('readers', 'readers'),
        ...role('read', 2), ...role('contribute', 3), ...role('full', 5),
        For_each_list: {
          runAfter: { Group_admins: ['Succeeded'], Group_managers: ['Succeeded'], Group_support: ['Succeeded', 'Failed'], Group_readers: ['Succeeded', 'Failed'], Role_read: ['Succeeded'], Role_contribute: ['Succeeded'], Role_full: ['Succeeded'] },
          type: 'Foreach', foreach: `@json('${JSON.stringify(SEC_LISTS)}')`, runtimeConfiguration: { concurrency: { repetitions: 1 } },
          actions: {
            Changed_items: spHttp('GET', "_api/web/lists/getbytitle('@{items('For_each_list')?['list']}')/items?$select=Id,SecuredOwner,@{items('For_each_list')?['owner']}Id&$filter=Modified ge datetime'@{addMinutes(utcNow(), -40)}'&$top=1000"),
            For_each_item: {
              runAfter: { Changed_items: ['Succeeded'] }, type: 'Foreach', foreach: "@coalesce(body('Changed_items')?['value'], json('[]'))",
              runtimeConfiguration: { concurrency: { repetitions: 4 } },
              actions: {
                Needs_securing: {
                  runAfter: {}, type: 'If',
                  expression: { and: [{ not: { equals: [`@coalesce(${OWNER}, 0)`, 0] } }, { not: { equals: [`@string(coalesce(items('For_each_item')?['SecuredOwner'], ''))`, `@string(${OWNER})`] } }] },
                  actions: {
                    Break_inheritance: spHttp('POST', `${ITEM}/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)`),
                    Grant_admins: addRole("body('Group_admins')?['Id']", "body('Role_full')?['Id']", { Break_inheritance: ['Succeeded'] }),
                    Grant_managers: addRole("body('Group_managers')?['Id']", "body('Role_contribute')?['Id']", { Grant_admins: ['Succeeded'] }),
                    Grant_support: addRole("body('Group_support')?['Id']", "if(items('For_each_list')?['cases'], body('Role_contribute')?['Id'], body('Role_read')?['Id'])", { Grant_managers: ['Succeeded'] }),
                    Grant_readers: addRole("body('Group_readers')?['Id']", "body('Role_read')?['Id']", { Grant_support: ['Succeeded', 'Failed'] }),
                    Grant_owner: addRole(OWNER, "body('Role_contribute')?['Id']", { Grant_readers: ['Succeeded', 'Failed'] }),
                    Owner_teams: {
                      runAfter: { Grant_owner: ['Succeeded'] }, type: 'Query',
                      inputs: { from: `@coalesce(${S('teams')}, json('[]'))`, where: `@contains(string(item()), concat('"id":', string(${OWNER}), ','))` }
                    },
                    For_each_team: {
                      runAfter: { Owner_teams: ['Succeeded'] }, type: 'Foreach', foreach: "@body('Owner_teams')", runtimeConfiguration: { concurrency: { repetitions: 1 } },
                      actions: {
                        Has_manager: {
                          runAfter: {}, type: 'If', expression: { and: [{ greater: ["@coalesce(items('For_each_team')?['manager']?['id'], 0)", 0] }] },
                          actions: { Grant_team_manager: addRole("items('For_each_team')?['manager']?['id']", "body('Role_contribute')?['Id']", {}) },
                          else: { actions: {} }
                        },
                        Share_with_team: {
                          runAfter: { Has_manager: ['Succeeded', 'Failed'] }, type: 'If', expression: { and: [{ not: { equals: [`@coalesce(${S('security')}?['teamAccess'], 'read')`, 'none'] } }] },
                          actions: {
                            For_each_member: {
                              runAfter: {}, type: 'Foreach', foreach: "@coalesce(items('For_each_team')?['members'], json('[]'))", runtimeConfiguration: { concurrency: { repetitions: 1 } },
                              actions: {
                                Grant_member: addRole("items('For_each_member')?['id']", `if(equals(${S('security')}?['teamAccess'], 'edit'), body('Role_contribute')?['Id'], body('Role_read')?['Id'])`, {})
                              }
                            }
                          },
                          else: { actions: {} }
                        }
                      }
                    },
                    Mark_secured: spHttp('POST', `${ITEM}/ValidateUpdateListItem`, { formValues: [{ FieldName: 'SecuredOwner', FieldValue: `@{string(${OWNER})}` }], bNewDocumentUpdate: true }, { For_each_team: ['Succeeded', 'Failed'] })
                  },
                  else: { actions: {} }
                }
              }
            }
          }
        }
      },
      else: { actions: {} }
    }
  });


// 21/22 — Lead & account enrichment (company website + optional Apollo.io)
const enrichFlow = (listKey, label) => {
  const E = (k) => `triggerOutputs()?['body/${k}']`;
  const DOMAIN = `@{toLower(first(split(replace(replace(replace(coalesce(${E('Website')}, concat('https://', last(split(coalesce(${E('Email')}, ''), '@')))), 'https://', ''), 'http://', ''), 'www.', ''), '/')))}`;
  const HTML = "string(body('Read_website'))";
  const between = (start, end) => `if(contains(${HTML}, '${start}'), trim(first(split(last(split(${HTML}, '${start}')), '${end}'))), '')`;
  flow(`Brillienta CRM - ${label} enrichment`,
    `When "Enrich" is clicked on a ${label.toLowerCase()} (or a new lead arrives with a business email): reads the company website (name, description, LinkedIn page) and, when an Apollo.io API key is set, adds industry, size, location and revenue. Writes the results to the record. Requires Power Automate Premium (HTTP).`,
    ['sp'], ['site', 'apolloKey'],
    { When_enrichment_is_requested: spPoll('GetOnUpdatedItems', listKey, [`@equals(${E('EnrichStatus/Value')}, 'Requested')`]) },
    {
      Domain: { runAfter: {}, type: 'Compose', inputs: DOMAIN },
      Read_website: http('GET', "https://@{outputs('Domain')}", { 'User-Agent': 'Mozilla/5.0 (compatible; BrillientaCRM/3.0)', Accept: 'text/html' }, undefined, { Domain: ['Succeeded'] }),
      Site_title: { runAfter: { Read_website: ['Succeeded', 'Failed', 'TimedOut'] }, type: 'Compose', inputs: `@{if(contains(${HTML}, '</title>'), trim(last(split(first(split(${HTML}, '</title>')), '>'))), '')}` },
      Site_description: { runAfter: { Read_website: ['Succeeded', 'Failed', 'TimedOut'] }, type: 'Compose', inputs: `@{take(if(contains(${HTML}, 'name="description"'), first(split(last(split(first(split(last(split(${HTML}, 'name="description"')), '>')), 'content="')), '"')), ${between('property="og:description" content="', '"')}), 500)}` },
      Site_linkedin: { runAfter: { Read_website: ['Succeeded', 'Failed', 'TimedOut'] }, type: 'Compose', inputs: `@{if(contains(${HTML}, 'linkedin.com/company/'), concat('https://www.linkedin.com/company/', first(split(first(split(first(split(last(split(${HTML}, 'linkedin.com/company/')), '"')), '/')), '?'))), '')}` },
      Use_apollo: {
        runAfter: { Site_title: ['Succeeded'], Site_description: ['Succeeded'], Site_linkedin: ['Succeeded'] }, type: 'If',
        expression: { and: [{ not: { equals: [P('apolloKey'), 'not-set'] } }] },
        actions: { Apollo: http('GET', "https://api.apollo.io/api/v1/organizations/enrich?domain=@{encodeUriComponent(outputs('Domain'))}", { 'X-Api-Key': Pi('apolloKey'), Accept: 'application/json' }) },
        else: { actions: {} }
      },
      Org: { runAfter: { Use_apollo: ['Succeeded', 'Failed', 'TimedOut'] }, type: 'Compose', inputs: "@coalesce(body('Apollo')?['organization'], json('{}'))" },
      Found: { runAfter: { Org: ['Succeeded'] }, type: 'Compose', inputs: `@or(not(empty(outputs('Site_title'))), not(empty(outputs('Org')?['name'])))` },
      Save: api('sp', 'PatchItem', {
        dataset: P('site'), table: LISTS[listKey], id: `@${E('ID')}`, 'item/Title': `@${E('Title')}`,
        'item/EnrichStatus/Value': "@{if(outputs('Found'), 'Done', 'Not found')}", 'item/EnrichedAt': '@utcNow()',
        'item/LinkedIn': `@{coalesce(${E('LinkedIn')}, outputs('Org')?['linkedin_url'], if(empty(outputs('Site_linkedin')), null, outputs('Site_linkedin')))}`,
        ...(listKey === 'leads' ? { 'item/CompanySize': `@{coalesce(${E('CompanySize')}, string(outputs('Org')?['estimated_num_employees']), '')}` } : { 'item/Employees': `@coalesce(${E('Employees')}, outputs('Org')?['estimated_num_employees'])` }),
        'item/CompanyInfo': "@{take(concat(coalesce(outputs('Org')?['name'], outputs('Site_title')), if(empty(outputs('Site_description')), '', concat('\\n', outputs('Site_description'))), if(empty(coalesce(outputs('Org')?['industry'], '')), '', concat('\\nIndustry: ', outputs('Org')?['industry'])), if(equals(outputs('Org')?['estimated_num_employees'], null), '', concat('\\nEmployees: ', string(outputs('Org')?['estimated_num_employees']))), if(empty(coalesce(outputs('Org')?['city'], outputs('Org')?['country'], '')), '', concat('\\nLocation: ', coalesce(outputs('Org')?['city'], ''), ' ', coalesce(outputs('Org')?['country'], ''))), if(empty(coalesce(outputs('Org')?['annual_revenue_printed'], '')), '', concat('\\nRevenue: ', outputs('Org')?['annual_revenue_printed'])), concat('\\nSource: ', outputs('Domain'))), 4000)}",
        'item/Country': `@{coalesce(${E('Country')}, outputs('Org')?['country'], '')}`
      }, { Found: ['Succeeded'] })
    });
};
enrichFlow('leads', 'Lead');
enrichFlow('accounts', 'Account');


// 23 — Customer portal (served by the flow): sign-in by email link, my cases, case conversation, new case, knowledge base
const Q_ = (k) => `triggerOutputs()?['queries']?['${k}']`;
const PORTAL = Pi('portalUrl');
const link = (pg, extra = '') => `${PORTAL}&p=${pg}&s=@{outputs('Token')}${extra}`;
const ME = "outputs('Session')";
const pform = (key) => ({ type: 'Query', inputs: { from: "@coalesce(triggerBody()?['$formdata'], json('[]'))", where: `@equals(item()?['key'], '${key}')` } });
const pv = (key) => `trim(string(coalesce(first(body('F_${key}'))?['value'], '')))`;
const nav = `<p class="m" style="margin-bottom:18px"><a href="${link('cases')}">My requests</a> · <a href="${link('new')}">New request</a> · <a href="${link('kb')}">Help articles</a> · <a href="${PORTAL}&p=login">Sign out</a></p>`;
const redirect = (url, runAfter = {}) => ({ runAfter, type: 'Response', kind: 'Http', inputs: { statusCode: 303, headers: { Location: url, 'Cache-Control': 'no-store' } } });
const ownsCase = `or(equals(body('Get_case')?['ContactRef'], ${ME}?['ID']), and(not(equals(${ME}?['AccountRef'], null)), equals(body('Get_case')?['AccountRef'], ${ME}?['AccountRef'])))`;
flow('Brillienta CRM - Customer portal',
  'Self-service portal for customers, served by this flow: sign in with a one-time email link, see the requests (cases) of their company, follow the conversation, reply, open new requests and search published help articles. Requires Power Automate Premium (HTTP trigger). Set the "Customer portal URL" variable to this flow\'s HTTP URL.',
  ['sp', 'outlook'], ['site', 'page', 'support', 'portalUrl', 'casePrefix'],
  { manual: { type: 'Request', kind: 'Http', inputs: { triggerAuthenticationType: 'All' } } },
  {
    Page: { runAfter: {}, type: 'Compose', inputs: `@{toLower(coalesce(${Q_('p')}, 'cases'))}` },
    Token: { runAfter: {}, type: 'Compose', inputs: `@{replace(coalesce(${Q_('s')}, ''), '''', '')}` },
    Is_post: { runAfter: {}, type: 'Compose', inputs: "@contains(toLower(string(coalesce(triggerOutputs()?['headers']?['Content-Type'], triggerOutputs()?['headers']?['content-type'], ''))), 'application/x-www-form-urlencoded')" },
    F_email: { runAfter: {}, ...pform('email') }, F_subject: { runAfter: {}, ...pform('subject') }, F_description: { runAfter: {}, ...pform('description') },
    F_impact: { runAfter: {}, ...pform('impact') }, F_message: { runAfter: {}, ...pform('message') },
    Find_session: getItems('contacts', "@{if(empty(outputs('Token')), 'Id eq 0', concat('PortalToken eq ''', outputs('Token'), ''''))}", { Token: ['Succeeded'] }, 1),
    Session: { runAfter: { Find_session: ['Succeeded'] }, type: 'Compose', inputs: "@first(body('Find_session')?['value'])" },
    Signed_in: {
      runAfter: { Session: ['Succeeded'], Page: ['Succeeded'], Is_post: ['Succeeded'], F_email: ['Succeeded'], F_subject: ['Succeeded'], F_description: ['Succeeded'], F_impact: ['Succeeded'], F_message: ['Succeeded'] },
      type: 'If',
      expression: { and: [{ not: { equals: [`@${ME}`, null] } }, { not: { equals: ["@outputs('Page')", 'login'] } }, { greater: [`@ticks(coalesce(${ME}?['PortalTokenExpires'], '2000-01-01T00:00:00Z'))`, '@ticks(utcNow())'] }] },
      actions: {
        Route: {
          runAfter: {}, type: 'Switch', expression: "@outputs('Page')",
          cases: {
            Case_page: {
              case: 'case',
              actions: {
                Get_case: api('sp', 'GetItem', { dataset: P('site'), table: LISTS.cases, id: `@int(coalesce(${Q_('id')}, '0'))` }),
                Not_yours: htmlResponse(page('Not found', `${nav}<h1>Request not found</h1>`), { Get_case: ['Failed', 'TimedOut'] }),
                Is_owner: {
                  runAfter: { Get_case: ['Succeeded'] }, type: 'If', expression: { and: [{ equals: [`@${ownsCase}`, true] }] },
                  actions: {
                    Posted_reply: {
                      runAfter: {}, type: 'If', expression: { and: [{ equals: ["@outputs('Is_post')", true] }, { not: { equals: [`@${pv('message')}`, ''] } }] },
                      actions: {
                        Add_reply: api('sp', 'PostItem', {
                          dataset: P('site'), table: LISTS.acts, 'item/Title': `@{take(concat('Received: ', body('Get_case')?['Title']), 255)}`,
                          'item/Details': `From: @{${ME}?['Title']} <@{${ME}?['Email']}> (customer portal)\n\n@{take(${pv('message')}, 8000)}`,
                          'item/ActivityType/Value': 'Email', 'item/RelatedType': 'Case', 'item/RelatedId': "@body('Get_case')?['ID']", 'item/Completed': true,
                          'item/DueDate': '@utcNow()', 'item/ActivitySource/Value': 'Website'
                        }),
                        Reopen_case: {
                          runAfter: { Add_reply: ['Succeeded'] }, type: 'If', expression: { and: [{ contains: ["@createArray('Resolved', 'Closed', 'Waiting on customer')", "@body('Get_case')?['CaseStatus']?['Value']"] }] },
                          actions: { Set_in_progress: api('sp', 'PatchItem', { dataset: P('site'), table: LISTS.cases, id: "@body('Get_case')?['ID']", 'item/Title': "@body('Get_case')?['Title']", 'item/CaseStatus/Value': 'In progress' }) },
                          else: { actions: {} }
                        },
                        Tell_agent: email(`@{coalesce(body('Get_case')?['CaseOwner']?['Email'], ${P('support').slice(1)})}`, "Customer replied in the portal: @{body('Get_case')?['CaseNumber']}",
                          html("New portal reply on @{body('Get_case')?['CaseNumber']}", `<p>@{${H(pv('message'))}}</p>`), { Add_reply: ['Succeeded'] }),
                        After_reply: redirect(`${link('case', "&id=@{body('Get_case')?['ID']}")}`, { Tell_agent: ['Succeeded', 'Failed'] })
                      },
                      else: {
                        actions: {
                          Conversation: getItems('acts', "RelatedType eq 'Case' and RelatedId eq @{body('Get_case')?['ID']} and ActivityType eq 'Email'", {}, 200),
                          Messages: {
                            runAfter: { Conversation: ['Succeeded'] }, type: 'Select',
                            inputs: { from: "@body('Conversation')?['value']", select: `@concat('<div class="done" style="margin:8px 0;background:', if(startsWith(coalesce(item()?['Title'], ''), 'Received'), '#f3f6fa', '#e6f4ee'), '"><b>', if(startsWith(coalesce(item()?['Title'], ''), 'Received'), 'You', 'Support'), '</b> <span class="m">', formatDateTime(coalesce(item()?['DueDate'], item()?['Created']), 'd MMM yyyy HH:mm'), '</span><div style="white-space:pre-wrap;margin-top:6px">', ${H("item()?['Details']")}, '</div></div>')` }
                          },
                          Case_view: htmlResponse(page('Request', `${nav}<h1>@{${H("body('Get_case')?['CaseNumber']")}} · @{${H("body('Get_case')?['Title']")}}</h1>
<p class="m">Status: <b>@{body('Get_case')?['CaseStatus']?['Value']}</b> · Priority: @{body('Get_case')?['Priority']?['Value']} · Opened @{formatDateTime(body('Get_case')?['Created'], 'd MMM yyyy')}</p>
@{if(empty(coalesce(body('Get_case')?['Description'], '')), '', concat('<div class="terms">', ${H("body('Get_case')?['Description']")}, '</div>'))}
@{if(empty(coalesce(body('Get_case')?['Resolution'], '')), '', concat('<p><b>Resolution</b></p><div class="done" style="white-space:pre-wrap">', ${H("body('Get_case')?['Resolution']")}, '</div>'))}
<h2 style="font-size:16px;margin-top:22px">Conversation</h2>@{join(body('Messages'), '')}
<form method="post" action=""><label for="m">Reply</label><textarea id="m" name="message" rows="4" required maxlength="8000"></textarea><div class="btns"><button class="ok" type="submit">Send reply</button></div></form>`), { Messages: ['Succeeded'] })
                        }
                      }
                    }
                  },
                  else: { actions: { Not_owner: htmlResponse(page('Not found', `${nav}<h1>Request not found</h1>`)) } }
                }
              }
            },
            New_page: {
              case: 'new',
              actions: {
                Posted_case: {
                  runAfter: {}, type: 'If', expression: { and: [{ equals: ["@outputs('Is_post')", true] }, { not: { equals: [`@${pv('subject')}`, ''] } }] },
                  actions: {
                    Create_portal_case: api('sp', 'PostItem', {
                      dataset: P('site'), table: LISTS.cases, 'item/Title': `@{take(${pv('subject')}, 255)}`, 'item/Description': `@{take(${pv('description')}, 60000)}`,
                      'item/Priority/Value': `@{if(contains(createArray('Low', 'Normal', 'High', 'Critical'), ${pv('impact')}), ${pv('impact')}, 'Normal')}`,
                      'item/CaseStatus/Value': 'New', 'item/Origin/Value': 'Portal', 'item/Category/Value': 'Question',
                      'item/ContactRef': `@${ME}?['ID']`, 'item/AccountRef': `@${ME}?['AccountRef']`, 'item/RequesterEmail': `@{${ME}?['Email']}`
                    }),
                    Number_case: api('sp', 'PatchItem', {
                      dataset: P('site'), table: LISTS.cases, id: "@body('Create_portal_case')?['ID']", 'item/Title': "@body('Create_portal_case')?['Title']",
                      'item/CaseNumber': `@{concat(${P('casePrefix').slice(1)}, '-', formatDateTime(utcNow(), 'yyyy'), '-', formatNumber(body('Create_portal_case')?['ID'], '0000'))}`
                    }, { Create_portal_case: ['Succeeded'] }),
                    Alert_new_case: email(P('support'), "New portal request: @{body('Create_portal_case')?['Title']}",
                      html('New request from the customer portal', `<table>${tr('From', `@{${H(`${ME}?['Title']`)}} &lt;@{${ME}?['Email']}&gt;`)}${tr('Subject', `@{${H("body('Create_portal_case')?['Title']")}}`)}</table>`), { Number_case: ['Succeeded', 'Failed'] }),
                    After_create: redirect(`${link('case', "&id=@{body('Create_portal_case')?['ID']}")}`, { Alert_new_case: ['Succeeded', 'Failed'] })
                  },
                  else: {
                    actions: {
                      New_form: htmlResponse(page('New request', `${nav}<h1>New request</h1><form method="post" action="">
<label for="s">Subject</label><input type="text" id="s" name="subject" required maxlength="200">
<label for="d">Describe the problem or question</label><textarea id="d" name="description" rows="6" maxlength="20000"></textarea>
<label for="i">How urgent is it?</label><select id="i" name="impact" style="padding:9px;border-radius:6px;border:1px solid #bbb"><option value="Low">Low – minor or cosmetic</option><option value="Normal" selected>Normal</option><option value="High">High – some people cannot work</option><option value="Critical">Critical – everyone is blocked</option></select>
<div class="btns"><button class="ok" type="submit">Submit request</button></div></form>`))
                    }
                  }
                }
              }
            },
            Kb_page: {
              case: 'kb',
              actions: {
                Search_text: { runAfter: {}, type: 'Compose', inputs: `@{replace(trim(coalesce(${Q_('q')}, '')), '''', '''''')}` },
                Articles: getItems('articles', "@{if(empty(outputs('Search_text')), 'ArticleStatus eq ''Published''', concat('ArticleStatus eq ''Published'' and (substringof(''', outputs('Search_text'), ''', Title) or substringof(''', outputs('Search_text'), ''', Keywords))'))}", { Search_text: ['Succeeded'] }, 50),
                Article_list: {
                  runAfter: { Articles: ['Succeeded'] }, type: 'Select',
                  inputs: { from: "@body('Articles')?['value']", select: `@concat('<details style="margin:8px 0"><summary style="cursor:pointer;font-weight:600">', ${H("item()?['Title']")}, '</summary><div class="terms" style="margin-top:6px">', ${H("item()?['Body']")}, '</div></details>')` }
                },
                Kb_view: htmlResponse(page('Help articles', `${nav}<h1>Help articles</h1><form method="get" action=""><input type="hidden" name="p" value="kb"><input type="hidden" name="s" value="@{outputs('Token')}"><input type="text" name="q" placeholder="Search" value="@{${H(Q_('q'))}}"></form>
@{if(empty(body('Article_list')), '<p class="m">No articles found.</p>', join(body('Article_list'), ''))}`), { Article_list: ['Succeeded'] })
              }
            }
          },
          default: {
            actions: {
              Account_filter: { runAfter: {}, type: 'Compose', inputs: `@{if(equals(${ME}?['AccountRef'], null), concat('ContactRef eq ', string(${ME}?['ID'])), concat('ContactRef eq ', string(${ME}?['ID']), ' or AccountRef eq ', string(${ME}?['AccountRef'])))}` },
              My_cases: getItems('cases', "@{outputs('Account_filter')}", { Account_filter: ['Succeeded'] }, 200),
              Case_rows: {
                runAfter: { My_cases: ['Succeeded'] }, type: 'Select',
                inputs: { from: "@body('My_cases')?['value']", select: `@concat('<tr><td><a href="${PORTAL}&p=case&s=', outputs('Token'), '&id=', string(item()?['ID']), '">', ${H("coalesce(item()?['CaseNumber'], string(item()?['ID']))")}, '</a></td><td>', ${H("item()?['Title']")}, '</td><td>', coalesce(item()?['CaseStatus']?['Value'], ''), '</td><td class="n">', formatDateTime(item()?['Modified'], 'd MMM yyyy'), '</td></tr>')` }
              },
              Touch_login: api('sp', 'PatchItem', { dataset: P('site'), table: LISTS.contacts, id: `@${ME}?['ID']`, 'item/Title': `@${ME}?['Title']`, 'item/PortalLastLogin': '@utcNow()' }),
              Cases_view: htmlResponse(page('My requests', `${nav}<h1>Hello @{${H(`${ME}?['FirstName']`)}}</h1>
<p class="m">Requests from you@{if(equals(${ME}?['AccountRef'], null), '', ' and your colleagues')}.</p>
<table><thead><tr><th>Ref</th><th>Subject</th><th>Status</th><th class="n">Updated</th></tr></thead><tbody>@{join(body('Case_rows'), '')}</tbody></table>
@{if(empty(body('Case_rows')), '<p class="m">No requests yet.</p>', '')}<div class="btns"><a class="ok" style="padding:10px 18px;border-radius:6px;text-decoration:none" href="${link('new')}">New request</a></div>`), { Case_rows: ['Succeeded'] })
            }
          }
        }
      },
      else: {
        actions: {
          Login_post: {
            runAfter: {}, type: 'If', expression: { and: [{ equals: ["@outputs('Is_post')", true] }, { contains: [`@${pv('email')}`, '@'] }] },
            actions: {
              Find_contact_by_email: getItems('contacts', `Email eq '@{replace(toLower(${pv('email')}), '''', '''''')}'`, {}, 1),
              Known_contact: {
                runAfter: { Find_contact_by_email: ['Succeeded'] }, type: 'If', expression: { and: [{ greater: ["@length(body('Find_contact_by_email')?['value'])", 0] }] },
                actions: {
                  New_token: { runAfter: {}, type: 'Compose', inputs: "@{concat(replace(guid(), '-', ''), replace(guid(), '-', ''))}" },
                  Save_token: api('sp', 'PatchItem', {
                    dataset: P('site'), table: LISTS.contacts, id: "@first(body('Find_contact_by_email')?['value'])?['ID']", 'item/Title': "@first(body('Find_contact_by_email')?['value'])?['Title']",
                    'item/PortalToken': "@{outputs('New_token')}", 'item/PortalTokenExpires': '@{addDays(utcNow(), 7)}'
                  }, { New_token: ['Succeeded'] }),
                  Send_link: email("@{first(body('Find_contact_by_email')?['value'])?['Email']}", 'Your sign-in link',
                    `<html><head>${STYLE}</head><body><p style="font-size:18px;font-weight:600;color:#008D62">Sign in to the customer portal</p><p><a class="btn" href="${PORTAL}&p=cases&s=@{outputs('New_token')}">Open my requests</a></p><p style="color:#888;font-size:12px">The link works for 7 days. If you did not ask for it, you can ignore this email.</p></body></html>`,
                    { Save_token: ['Succeeded'] })
                },
                else: { actions: {} }
              },
              Check_email: htmlResponse(page('Check your email', '<h1>Check your email</h1><div class="done">If this address is registered with us, we have sent you a sign-in link. It works for 7 days.</div>'), { Known_contact: ['Succeeded', 'Failed'] })
            },
            else: {
              actions: {
                Login_form: htmlResponse(page('Customer portal', `<h1>Customer portal</h1><p class="m">Follow your support requests, reply to our team and open new requests.</p>
<form method="post" action="${PORTAL}&p=login"><label for="e">Your work email</label><input type="text" id="e" name="email" required autocomplete="email"><div class="btns"><button class="ok" type="submit">Email me a sign-in link</button></div></form>`))
              }
            }
          }
        }
      }
    }
  });

// ------------------------------------------------------------------ validation
function validate(f) {
  const errors = [];
  const names = new Set();
  const walk = (actions, scope) => {
    Object.keys(actions || {}).forEach(k => {
      if (names.has(k)) { errors.push(`duplicate action name ${k}`); }
      names.add(k);
      const a = actions[k];
      Object.keys(a.runAfter || {}).forEach(r => { if (!(r in actions)) { errors.push(`${k}: runAfter "${r}" not in same scope (${scope})`); } });
      if (a.actions) { walk(a.actions, k); }
      if (a.else && a.else.actions) { walk(a.else.actions, `${k}/else`); }
      if (a.cases) { Object.keys(a.cases).forEach(c => walk(a.cases[c].actions, `${k}/${c}`)); }
      if (a.default && a.default.actions) { walk(a.default.actions, `${k}/default`); }
    });
  };
  walk(f.actions, 'root');
  const text = JSON.stringify(f);
  for (const m of text.matchAll(/(?:body|outputs|items|actions)\('([^']+)'\)/g)) {
    if (!names.has(m[1]) && !(m[1] in f.triggers)) { errors.push(`reference to unknown action "${m[1]}"`); }
  }
  for (const m of text.matchAll(/parameters\('([^']+)'\)/g)) {
    if (m[1] !== '$authentication' && m[1] !== '$connections' && !f.envKeys.some(k => pname(k) === m[1])) { errors.push(`unknown parameter "${m[1]}"`); }
  }
  for (const m of text.matchAll(/variables\('([^']+)'\)/g)) {
    if (!text.includes(`"name":"${m[1]}"`)) { errors.push(`unknown variable ${m[1]}`); }
  }
  // Functions Power Automate rejects without arguments (use json('[]') for an empty array).
  for (const fn of ['createArray', 'concat', 'coalesce', 'union', 'intersection', 'first', 'last', 'join', 'split']) {
    if (text.includes(`${fn}()`)) { errors.push(`${fn}() needs arguments`); }
  }
  // balanced parentheses inside every @{...} / @expr
  const strings = [];
  const collect = (o) => { if (typeof o === 'string') { strings.push(o); } else if (o && typeof o === 'object') { Object.values(o).forEach(collect); } };
  collect(f.actions); collect(f.triggers);
  strings.forEach(s => {
    const exprs = s.startsWith('@') && !s.startsWith('@{') ? [s.slice(1)] : [...s.matchAll(/@\{([^}]*(?:\}[^@]*?)?)\}/g)].map(x => x[1]);
    exprs.forEach(e => {
      let depth = 0, q = false;
      for (const ch of e) { if (ch === "'") { q = !q; } else if (!q) { if (ch === '(' || ch === '[') { depth++; } if (ch === ')' || ch === ']') { depth--; } } }
      if (depth !== 0 || q) { errors.push(`unbalanced expression: ${e.slice(0, 80)}`); }
    });
  });
  return errors;
}

// ------------------------------------------------------------------ write solution
const OUT = join(HERE, 'dist', 'solution');
rmSync(join(HERE, 'dist'), { recursive: true, force: true });
mkdirSync(join(OUT, 'Workflows'), { recursive: true });

let failed = false;
const xmlEsc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const workflowXml = [];
FLOWS.forEach(f => {
  const errs = validate(f);
  if (errs.length) { failed = true; console.error(`✖ ${f.name}\n  - ${errs.join('\n  - ')}`); } else { console.log(`✔ ${f.name}`); }
  const parameters = {
    $connections: { defaultValue: {}, type: 'Object' },
    $authentication: { defaultValue: {}, type: 'SecureObject' }
  };
  f.envKeys.forEach(k => { parameters[pname(k)] = { defaultValue: ENV[k].def, type: 'String', metadata: { schemaName: ENV[k].schema, description: ENV[k].desc } }; });
  const connectionReferences = {};
  f.connections.forEach(c => {
    connectionReferences[CONN[c].api] = { runtimeSource: 'embedded', connection: { connectionReferenceLogicalName: CONN[c].ref }, api: { name: CONN[c].api } };
  });
  const json = {
    properties: {
      connectionReferences,
      definition: {
        $schema: 'https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#',
        contentVersion: '1.0.0.0',
        parameters, triggers: f.triggers, actions: f.actions
      },
      templateName: ''
    },
    schemaVersion: '1.0.0.0'
  };
  const file = `${f.name.replace(/[^A-Za-z0-9]/g, '')}-${f.id.toUpperCase()}.json`;
  writeFileSync(join(OUT, 'Workflows', file), JSON.stringify(json, null, 2));
  workflowXml.push(`    <Workflow WorkflowId="{${f.id}}" Name="${xmlEsc(f.name)}">
      <JsonFileName>/Workflows/${file}</JsonFileName>
      <Type>1</Type>
      <Subprocess>0</Subprocess>
      <Category>5</Category>
      <Mode>0</Mode>
      <Scope>4</Scope>
      <OnDemand>0</OnDemand>
      <TriggerOnCreate>0</TriggerOnCreate>
      <TriggerOnDelete>0</TriggerOnDelete>
      <AsyncAutodelete>0</AsyncAutodelete>
      <SyncWorkflowLogOnFailure>0</SyncWorkflowLogOnFailure>
      <StateCode>0</StateCode>
      <StatusCode>1</StatusCode>
      <RunAs>1</RunAs>
      <IsTransacted>1</IsTransacted>
      <IntroducedVersion>${VERSION}</IntroducedVersion>
      <IsCustomizable>1</IsCustomizable>
      <BusinessProcessType>0</BusinessProcessType>
      <IsCustomProcessingStepAllowedForOtherPublishers>1</IsCustomProcessingStepAllowedForOtherPublishers>
      <PrimaryEntity>none</PrimaryEntity>
      <LocalizedNames>
        <LocalizedName languagecode="1033" description="${xmlEsc(f.name)}" />
      </LocalizedNames>
      <Descriptions>
        <Description languagecode="1033" description="${xmlEsc(f.description)}" />
      </Descriptions>
    </Workflow>`);
});
if (failed) { process.exit(1); }

const usedConns = [...new Set(FLOWS.flatMap(f => f.connections))];
const connXml = usedConns.map(c => `    <connectionreference connectionreferencelogicalname="${CONN[c].ref}">
      <connectionreferencedisplayname>${xmlEsc(CONN[c].display)}</connectionreferencedisplayname>
      <connectorid>/providers/Microsoft.PowerApps/apis/${CONN[c].api}</connectorid>
      <iscustomizable>1</iscustomizable>
      <promptingbehavior>0</promptingbehavior>
      <statecode>0</statecode>
      <statuscode>1</statuscode>
    </connectionreference>`).join('\n');

writeFileSync(join(OUT, 'customizations.xml'), `<?xml version="1.0" encoding="utf-8"?>
<ImportExportXml xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" OrganizationVersion="9.2.0.0" OrganizationSchemaType="Standard" CRMServerServiceabilityVersion="9.2.0.0">
  <Entities />
  <Roles />
  <Workflows>
${workflowXml.join('\n')}
  </Workflows>
  <FieldSecurityProfiles />
  <Templates />
  <EntityMaps />
  <EntityRelationships />
  <OrganizationSettings />
  <optionsets />
  <CustomControls />
  <EntityDataProviders />
  <connectionreferences>
${connXml}
  </connectionreferences>
  <Languages>
    <Language>1033</Language>
  </Languages>
</ImportExportXml>
`);

Object.keys(ENV).forEach(k => {
  const e = ENV[k];
  mkdirSync(join(OUT, 'environmentvariabledefinitions', e.schema), { recursive: true });
  writeFileSync(join(OUT, 'environmentvariabledefinitions', e.schema, 'environmentvariabledefinition.xml'), `<environmentvariabledefinition schemaname="${e.schema}">
  <defaultvalue>${xmlEsc(e.def)}</defaultvalue>
  <description default="${xmlEsc(e.desc)}">
    <label description="${xmlEsc(e.desc)}" languagecode="1033" />
  </description>
  <displayname default="${xmlEsc(e.name)}">
    <label description="${xmlEsc(e.name)}" languagecode="1033" />
  </displayname>
  <introducedversion>${VERSION}</introducedversion>
  <iscustomizable>1</iscustomizable>
  <isrequired>0</isrequired>
  <secretstore>0</secretstore>
  <type>100000000</type>
</environmentvariabledefinition>
`);
});

const nil = (t) => `<${t} xsi:nil="true"></${t}>`;
const address = (n) => `        <Address>
          <AddressNumber>${n}</AddressNumber>
          <AddressTypeCode>1</AddressTypeCode>
          ${['City', 'County', 'Country', 'Fax'].map(nil).join('')}
          <FreightTermsCode xsi:nil="true"></FreightTermsCode>
          ${['ImportSequenceNumber', 'Latitude', 'Line1', 'Line2', 'Line3', 'Longitude', 'Name', 'PostalCode', 'PostOfficeBox', 'PrimaryContactName'].map(nil).join('')}
          <ShippingMethodCode>1</ShippingMethodCode>
          ${['StateOrProvince', 'Telephone1', 'Telephone2', 'Telephone3', 'TimeZoneRuleVersionNumber', 'UPSZone', 'UTCOffset', 'UTCConversionTimeZoneCode'].map(nil).join('')}
        </Address>`;

writeFileSync(join(OUT, 'solution.xml'), `<?xml version="1.0" encoding="utf-8"?>
<ImportExportXml version="9.2.0.0" SolutionPackageVersion="9.2" languagecode="1033" generatedBy="CrmLive" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" OrganizationVersion="9.2.0.0" OrganizationSchemaType="Standard" CRMServerServiceabilityVersion="9.2.0.0">
  <SolutionManifest>
    <UniqueName>${SOLUTION}</UniqueName>
    <LocalizedNames>
      <LocalizedName description="Brillienta CRM Flows" languagecode="1033" />
    </LocalizedNames>
    <Descriptions>
      <Description description="${xmlEsc('Ready-made Power Automate flows for Brillienta CRM: AppSource lead intake, website leads, demo bookings & support requests, new lead alerts, daily sales digest, overdue reminders, case SLA escalation, deal-won handoff and quote discount approval.')}" languagecode="1033" />
    </Descriptions>
    <Version>${VERSION}</Version>
    <Managed>0</Managed>
    <Publisher>
      <UniqueName>brillienta</UniqueName>
      <LocalizedNames>
        <LocalizedName description="Brillienta" languagecode="1033" />
      </LocalizedNames>
      <Descriptions>
        <Description description="Brillienta – workplace products for Microsoft 365" languagecode="1033" />
      </Descriptions>
      <EMailAddress xsi:nil="true"></EMailAddress>
      <SupportingWebsiteUrl>https://brillienta.com</SupportingWebsiteUrl>
      <CustomizationPrefix>${PREFIX}</CustomizationPrefix>
      <CustomizationOptionValuePrefix>61937</CustomizationOptionValuePrefix>
      <Addresses>
${address(1)}
${address(2)}
      </Addresses>
    </Publisher>
    <RootComponents>
${FLOWS.map(f => `      <RootComponent type="29" id="{${f.id}}" behavior="0" />`).join('\n')}
${Object.keys(ENV).map(k => `      <RootComponent type="380" schemaName="${ENV[k].schema}" behavior="0" />`).join('\n')}
    </RootComponents>
    <MissingDependencies />
  </SolutionManifest>
</ImportExportXml>
`);

writeFileSync(join(OUT, '[Content_Types].xml'), '<?xml version="1.0" encoding="utf-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/octet-stream" /><Default Extension="json" ContentType="application/octet-stream" /></Types>');

// Fail the build if any generated file is not well-formed (Power Platform rejects the import otherwise).
{
  const { readdirSync: rd, statSync: st, readFileSync: rf } = await import('node:fs');
  const bad = [];
  const walk = (dir) => rd(dir).forEach(n => {
    const full = join(dir, n);
    if (st(full).isDirectory()) { walk(full); return; }
    const text = rf(full, 'utf8');
    if (n.endsWith('.json')) { try { JSON.parse(text); } catch (e) { bad.push(`${full}: ${e.message}`); } }
    if (n.endsWith('.xml')) {
      const noCdata = text.replace(/<!\[CDATA\[[\s\S]*?\]\]>/g, '');
      const amp = /&(?!(amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)/.exec(noCdata);
      if (amp) { bad.push(`${full}: unescaped "&" near "${noCdata.slice(Math.max(0, amp.index - 30), amp.index + 10)}"`); }
      const tags = noCdata.replace(/<\?[\s\S]*?\?>/g, '').match(/<\/?[A-Za-z][^>]*?\/?>/g) || [];
      const stack = [];
      tags.forEach(t => {
        if (t.endsWith('/>')) { return; }
        const name = t.replace(/^<\/?/, '').split(/[\s>]/)[0];
        if (t.startsWith('</')) { if (stack.pop() !== name) { bad.push(`${full}: mismatched </${name}>`); } } else { stack.push(name); }
      });
      if (stack.length) { bad.push(`${full}: unclosed <${stack[stack.length - 1]}>`); }
    }
  });
  walk(OUT);
  if (bad.length) { console.error('✖ Malformed solution files:\n  - ' + bad.join('\n  - ')); process.exit(1); }
  console.log('✔ All solution files are well-formed');
}
const zip = join(HERE, 'dist', `${SOLUTION}_${VERSION.replace(/\./g, '_')}.zip`);
execSync(`cd "${OUT}" && zip -qr -X "${zip}" "[Content_Types].xml" solution.xml customizations.xml Workflows environmentvariabledefinitions`);
console.log(`\nSolution: ${zip}`);

// ------------------------------------------------------------------ embed in the CRM web part
// The CRM builds the zip in the browser (Settings › Integrations › Download), pre-filled with the customer's values.
import('node:fs').then(({ readdirSync, statSync, readFileSync: rf }) => {
  const files = {};
  const walk = (dir, prefix = '') => readdirSync(dir).sort().forEach(n => {
    const full = join(dir, n);
    if (statSync(full).isDirectory()) { walk(full, `${prefix}${n}/`); } else { files[`${prefix}${n}`] = rf(full, 'utf8'); }
  });
  walk(OUT);
  const env = Object.keys(ENV).map(k => ({ schema: ENV[k].schema, name: ENV[k].name, description: ENV[k].desc, defaultValue: ENV[k].def }));
  const flows = FLOWS.map(f => ({ name: f.name, description: f.description }));
  const target = join(HERE, '..', 'src', 'webparts', 'brillientaCrm', 'models', 'flowPack.generated.ts');
  writeFileSync(target, `// GENERATED by flows/build-solution.mjs — do not edit. Power Platform solution files for the CRM's download button.\n` +
    `export const FLOW_PACK_VERSION = ${JSON.stringify(VERSION)};\n` +
    `export const FLOW_PACK_SOLUTION = ${JSON.stringify(SOLUTION)};\n` +
    `export const FLOW_PACK_ENV: { schema: string; name: string; description: string; defaultValue: string }[] = ${JSON.stringify(env, null, 2)};\n` +
    `export const FLOW_PACK_FLOWS: { name: string; description: string }[] = ${JSON.stringify(flows, null, 2)};\n` +
    `export const FLOW_PACK_FILES: Record<string, string> = ${JSON.stringify(files)};\n`);
  console.log(`Embedded in the web part: ${target} (${Object.keys(files).length} files)`);
});
