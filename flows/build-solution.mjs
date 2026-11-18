// Builds "Brillienta CRM Flows" — an importable (unmanaged) Power Platform solution with 7 cloud flows.
// Usage: node flows/build-solution.mjs            → flows/dist/BrillientaCRMFlows_1_0_0_0.zip
import { createHash } from 'node:crypto';
import { mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { execSync } from 'node:child_process';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));
const VERSION = '1.2.0.0';
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
  library: { schema: `${PREFIX}_CustomerFolderLibrary`, name: 'Customer folder library', def: 'Documents', desc: 'Document library (on the CRM site) where a folder is created for every won customer.' }
};
const pname = (k) => `${ENV[k].name} (${ENV[k].schema})`;
const P = (k) => `@parameters('${pname(k)}')`;
const Pi = (k) => `@{parameters('${pname(k)}')}`;

// ------------------------------------------------------------------ connectors
const CONN = {
  sp: { api: 'shared_sharepointonline', ref: `${PREFIX}_BrillientaCRM_SharePoint`, display: 'Brillienta CRM – SharePoint' },
  outlook: { api: 'shared_office365', ref: `${PREFIX}_BrillientaCRM_Outlook`, display: 'Brillienta CRM – Office 365 Outlook' },
  approvals: { api: 'shared_approvals', ref: `${PREFIX}_BrillientaCRM_Approvals`, display: 'Brillienta CRM – Approvals' }
};
const host = (c, op) => ({ connectionName: CONN[c].api, operationId: op, apiId: `/providers/Microsoft.PowerApps/apis/${CONN[c].api}` });
const api = (c, op, parameters, runAfter = {}, extra = {}) => ({
  runAfter, type: 'OpenApiConnection', inputs: { host: host(c, op), parameters, authentication: "@parameters('$authentication')" }, ...extra
});
const LISTS = { demos: 'CRM Demo Requests', leads: 'AppSource Leads', accounts: 'CRM Accounts', opps: 'CRM Opportunities', quotes: 'CRM Quotes', cases: 'CRM Cases', acts: 'CRM Activities' };

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

// 1b — Website leads & demo bookings (brillienta.com /api/lead forwards each submission here)
const tb = (k) => `triggerBody()?['${k}']`;
flow('Brillienta CRM - Website leads and demo bookings',
  'Receives submissions from the website (book-demo and contact forms) forwarded by the site\'s /api/lead function. Creates a lead (with UTM attribution); for demo bookings also creates a demo request with the requested time slot and length, and alerts the sales manager. Requires Power Automate Premium (HTTP trigger).',
  ['sp', 'outlook'], ['site', 'page', 'sales'],
  {
    manual: {
      type: 'Request', kind: 'Http',
      inputs: {
        triggerAuthenticationType: 'All', method: 'POST',
        schema: {
          type: 'object',
          properties: {
            kind: { type: 'string' }, name: { type: 'string' }, email: { type: 'string' }, company: { type: 'string' }, role: { type: 'string' },
            size: { type: 'string' }, country: { type: 'string' }, notes: { type: 'string' }, message: { type: 'string' }, subject: { type: 'string' },
            request_type: { type: 'string' }, products: { type: 'array', items: { type: 'string' } }, duration: { type: 'integer' }, start: { type: 'string' },
            timeZone: { type: 'string' }, website: { type: 'string' }, attribution: { type: 'object' }, landing_page: { type: 'string' }, referrer: { type: 'string' }, page: { type: 'string' }
          }
        }
      }
    }
  },
  {
    Full_name: { runAfter: {}, type: 'Compose', inputs: `@trim(coalesce(${tb('name')}, ''))` },
    First_name: { runAfter: { Full_name: ['Succeeded'] }, type: 'Compose', inputs: "@first(split(outputs('Full_name'), ' '))" },
    Last_name: { runAfter: { First_name: ['Succeeded'] }, type: 'Compose', inputs: "@if(contains(outputs('Full_name'), ' '), trim(substring(outputs('Full_name'), indexOf(outputs('Full_name'), ' '))), '')" },
    Products: { runAfter: {}, type: 'Compose', inputs: `@join(coalesce(${tb('products')}, createArray()), ', ')` },
    Is_demo: { runAfter: {}, type: 'Compose', inputs: `@equals(toLower(coalesce(${tb('kind')}, '')), 'demo')` },
    Details: {
      runAfter: { Products: ['Succeeded'], Is_demo: ['Succeeded'] }, type: 'Compose',
      inputs: `@{coalesce(${tb('notes')}, ${tb('message')}, '')}

— Website ${'@'}{coalesce(${tb('kind')}, 'lead')} —
Subject: @{coalesce(${tb('subject')}, ${tb('request_type')}, '')}
Products: @{outputs('Products')}
Requested time: @{coalesce(${tb('start')}, '')} (@{coalesce(${tb('timeZone')}, '')}), @{coalesce(${tb('duration')}, '')} min
Company size: @{coalesce(${tb('size')}, '')}
UTM: @{coalesce(${tb('attribution')}?['utm_source'], '')} / @{coalesce(${tb('attribution')}?['utm_medium'], '')} / @{coalesce(${tb('attribution')}?['utm_campaign'], '')}
Landing page: @{coalesce(${tb('landing_page')}, '')}
Referrer: @{coalesce(${tb('referrer')}, '')}`
    },
    Is_human: {
      runAfter: { Details: ['Succeeded'], Last_name: ['Succeeded'] }, type: 'If',
      // the website's hidden "website" field is a spam trap
      expression: { and: [{ equals: [`@coalesce(${tb('website')}, '')`, ''] }, { not: { equals: [`@coalesce(${tb('email')}, '')`, ''] } }] },
      actions: {
        Create_lead: api('sp', 'PostItem', {
          dataset: P('site'), table: LISTS.leads,
          'item/Title': "@{outputs('Full_name')}",
          'item/FirstName': "@{outputs('First_name')}",
          'item/LastName': "@{outputs('Last_name')}",
          'item/Email': `@{${tb('email')}}`,
          'item/Company': `@{${tb('company')}}`,
          'item/JobTitle': `@{${tb('role')}}`,
          'item/Country': `@{${tb('country')}}`,
          'item/LeadSource/Value': "@{if(outputs('Is_demo'), 'Website Demo Form', 'Website')}",
          'item/Status/Value': 'New',
          'item/Rating/Value': "@{if(outputs('Is_demo'), 'Hot', 'Warm')}",
          'item/ActionCode': `@{coalesce(${tb('kind')}, 'lead')}`,
          'item/Offer': "@{outputs('Products')}",
          'item/Description': "@{outputs('Details')}",
          'item/RawPayload': '@{string(triggerBody())}'
        }),
        Demo_booking: {
          runAfter: { Create_lead: ['Succeeded'] }, type: 'If',
          expression: { and: [{ equals: ["@outputs('Is_demo')", true] }] },
          actions: {
            Create_demo_request: api('sp', 'PostItem', {
              dataset: P('site'), table: LISTS.demos,
              'item/Title': `Demo - @{${tb('company')}} (@{outputs('Products')})`,
              'item/ContactName': "@{outputs('Full_name')}",
              'item/Email': `@{${tb('email')}}`,
              'item/Company': `@{${tb('company')}}`,
              'item/LeadRef': "@body('Create_lead')?['ID']",
              'item/PreferredDate': `@if(empty(coalesce(${tb('start')}, '')), null, ${tb('start')})`,
              'item/DurationMinutes': `@coalesce(${tb('duration')}, 30)`,
              'item/DemoStatus/Value': 'Requested',
              'item/Notes': "@{outputs('Details')}"
            }),
            Alert_sales: email(P('sales'), `Demo booked: @{outputs('Full_name')} (@{${tb('company')}}) – @{outputs('Products')}`,
              html('A visitor booked a demo on brillienta.com', `<table>${tr('Name', "@{outputs('Full_name')}")}${tr('Company', `@{${tb('company')}}`)}${tr('Role', `@{${tb('role')}}`)}${tr('Email', `@{${tb('email')}}`)}${tr('Products', "@{outputs('Products')}")}${tr('Requested time (UTC)', `@{${tb('start')}}`)}${tr('Visitor time zone', `@{${tb('timeZone')}}`)}${tr('Length', `@{${tb('duration')}} minutes`)}${tr('Company size', `@{${tb('size')}}`)}${tr('Notes', `@{${tb('notes')}}`)}</table><p>Confirm it in the CRM: Demo requests › Schedule Teams meeting (sends the Teams invite).</p>`),
              { Create_demo_request: ['Succeeded'] }, { 'emailMessage/Importance': 'High' })
          },
          else: { actions: {} }
        }
      },
      else: { actions: {} }
    },
    Response: {
      runAfter: { Is_human: ['Succeeded', 'Failed', 'Skipped', 'TimedOut'] }, type: 'Response', kind: 'Http',
      inputs: { statusCode: 200, headers: { 'Content-Type': 'application/json' }, body: { ok: true } }
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
      <Description description="Ready-made Power Automate flows for Brillienta CRM: AppSource lead intake, website leads & demo bookings, new lead alerts, daily sales digest, overdue reminders, case SLA escalation, deal-won handoff and quote discount approval." languagecode="1033" />
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

const zip = join(HERE, 'dist', `${SOLUTION}_${VERSION.replace(/\./g, '_')}.zip`);
execSync(`cd "${OUT}" && zip -qr -X "${zip}" "[Content_Types].xml" solution.xml customizations.xml Workflows environmentvariabledefinitions`);
console.log(`\nSolution: ${zip}`);
