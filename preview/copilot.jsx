// Local preview of the Brillienta CRM Copilot cards (copilot.html): each tool's card, as Microsoft 365 Copilot shows it
// inline, with the sample CRM. The CRM site is found and read exactly as in Copilot (crmSession.ts).
//   ?tool=pipeline|account|renewals|invoices|cases|lead   ?as=chen   ?dark=1   ?props=<JSON tool parameters>
import * as React from 'react';
import { createRoot } from 'react-dom/client';
import { USERS, WEB, createFakeHttp } from './fakeSharePoint';
import { getCrmSession } from '../lib/copilotComponents/shared/crmSession';
import { Frame, ErrorBox, Loading } from '../lib/copilotComponents/shared/ui';
import { PipelineView } from '../lib/copilotComponents/crmPipeline/CrmPipelineCopilotComponent';
import { AccountSummaryView } from '../lib/copilotComponents/crmAccountSummary/CrmAccountSummaryCopilotComponent';
import { RenewalsView } from '../lib/copilotComponents/crmRenewals/CrmRenewalsCopilotComponent';
import { InvoicesView } from '../lib/copilotComponents/crmOverdueInvoices/CrmOverdueInvoicesCopilotComponent';
import { CasesView } from '../lib/copilotComponents/crmSlaRiskCases/CrmSlaRiskCasesCopilotComponent';
import { CreateLeadView } from '../lib/copilotComponents/crmCreateLead/CrmCreateLeadCopilotComponent';

const params = new URLSearchParams(window.location.search);
const dark = !!params.get('dark');
const TOOLS = {
  pipeline: ['Pipeline', PipelineView, ['entity', 'opportunities', 'accounts', 'activities', 'quotes'], 'Which of my deals close this quarter?'],
  account: ['Account summary', AccountSummaryView, ['accounts', 'contacts', 'opportunities', 'quotes', 'cases', 'subscriptions', 'invoices', 'activities'], 'Summarise Contoso.', { accountName: 'Contoso' }],
  renewals: ['Upcoming renewals', RenewalsView, ['subscriptions', 'accounts'], 'Which subscriptions renew in the next 60 days?'],
  invoices: ['Overdue invoices', InvoicesView, ['invoices', 'accounts'], 'Who owes us money?'],
  cases: ['Cases at SLA risk', CasesView, ['cases', 'accounts'], 'Which support cases are breaching their SLA?'],
  lead: ['New lead', CreateLeadView, ['leads', 'activities', 'demos'], 'Add a lead: Kavya Iyer, Head of HR at Litware Health, kavya@litwarehealth.com, met her at the HR Tech expo.',
    { firstName: 'Kavya', lastName: 'Iyer', email: 'kavya@litwarehealth.com', company: 'Litware Health', jobTitle: 'Head of HR', leadSource: 'Event', description: 'Met at the HR Tech expo.' }]
};
const [title, View, lists, prompt, defaults] = TOOLS[params.get('tool') || 'pipeline'];
const properties = params.get('props') ? JSON.parse(params.get('props')) : defaults || {};
const user = USERS[params.get('as') || 'chen'] || USERS.chen;
window.history.replaceState(null, '', `${window.location.pathname}?${params.toString()}`);
// The fake SharePoint reads ?as= from the address.
if (!params.get('as')) { params.set('as', 'chen'); window.history.replaceState(null, '', `${window.location.pathname}?${params.toString()}`); }

const ctx = { spHttpClient: createFakeHttp(), pageContext: { web: { absoluteUrl: 'https://contoso.sharepoint.com' }, aadInfo: { tenantId: '6f1c2a9e' } } };
window.__copilot = { context: '' };

const Card = () => {
  const [state, setState] = React.useState({});
  React.useEffect(() => {
    getCrmSession(ctx).then(async (s) => { await s.engine.ensure(lists.filter((l) => l !== 'entity')); setState({ s }); }).catch((e) => setState({ e }));
  }, []);
  const host = state.s && {
    session: state.s, dark, fullscreen: false,
    openLink: (u) => console.log('open', u), expand: () => console.log('expand'),
    setModelContext: (t) => { window.__copilot.context = t; }, sendPrompt: (t) => console.log('prompt', t)
  };
  const site = state.s ? (state.s.engine.settings.company.name || state.s.siteUrl.replace(/^https:\/\//, '')) : '';
  return (
    <Frame dark={dark} title={title} subtitle={site ? `Brillienta CRM · ${site}` : 'Brillienta CRM'}>
      {state.e ? <ErrorBox error={state.e} /> : host ? <View host={host} properties={properties} /> : <Loading />}
    </Frame>
  );
};

document.body.className = dark ? 'dark' : '';
createRoot(document.getElementById('root')).render(
  <div className="chat">
    <div className="msg user"><span className="avatar">{user.Title.split(' ').map((x) => x[0]).join('')}</span><div className="bubble">{prompt}</div></div>
    <div className="msg agent"><span className="avatar agent-avatar" aria-hidden="true" /><div className="card"><Card /></div></div>
  </div>
);
void WEB;
