/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
import {
  Pivot, PivotItem, Stack, TextField, PrimaryButton, DefaultButton, Toggle, Dropdown, MessageBar, MessageBarType, Text, IconButton,
  Link, Checkbox, Spinner, SpinnerSize, Label, Dialog, DialogType, DialogFooter
} from '@fluentui/react';
import { downloadFlowPack, IFlowPackInfo, loadFlowPack } from '../../services/FlowPackage';
import { ENTITIES, EntityKey, FieldType } from '../../models/model';
import { ICrmSettings, ICustomField, IStageDef, IWebhook, WEBHOOK_EVENTS } from '../../models/settings';
import { useCrm } from '../CrmContext';
import { UserPicker } from '../core/controls';
import { errorText, uid } from '../utils';
import { Card } from './Dashboard';

const lines = (a: string[]): string => a.join('\n');
const parseLines = (s: string): string[] => s.split('\n').map(x => x.trim()).filter((x, i, arr) => x && arr.indexOf(x) === i);
const kv = (o: Record<string, number>): string => Object.keys(o).map(k => `${k} = ${o[k]}`).join('\n');
const parseKv = (s: string): Record<string, number> => {
  const out: Record<string, number> = {};
  s.split('\n').forEach(l => { const m = l.split('='); if (m.length === 2 && m[0].trim()) { out[m[0].trim()] = Number(m[1]) || 0; } });
  return out;
};

const TENANT_SITE_KEY = 'BrillientaCRM.SiteUrl';

const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

/** Download the Power Automate solution, pre-filled with this customer's site, page and recipients. */
const FlowPackDialog: React.FC<{ onClose: () => void }> = ({ onClose }) => {
  const { engine, notify } = useCrm();
  const [pack, setPack] = React.useState<IFlowPackInfo>();
  const [values, setValues] = React.useState<Record<string, string>>({});
  const [error, setError] = React.useState<string>();
  const [done, setDone] = React.useState<string>();
  React.useEffect(() => {
    loadFlowPack().then(p => {
      const site = engine.ds.rest.webUrl.replace(/\/$/, '');
      const here = window.location.href.split(/[?#]/)[0];
      const page = /\.aspx$/i.test(here) && here.toLowerCase().indexOf(site.toLowerCase()) === 0 ? here : `${site}/SitePages/Home.aspx`;
      const me = engine.me?.email || '';
      const pre: Record<string, string> = {};
      p.env.forEach(e => {
        const n = e.schema.toLowerCase();
        pre[e.schema] = n.indexOf('crmsiteurl') >= 0 ? site
          : n.indexOf('crmpageurl') >= 0 ? page
            : /email$/.test(n) ? (me || e.defaultValue)
              : e.defaultValue;
      });
      setValues(pre);
      setPack(p);
    }).catch(e => setError(errorText(e)));
  }, []);
  const invalid = (schema: string, v: string): string | undefined => {
    const n = schema.toLowerCase();
    if (!v.trim()) { return 'Required'; }
    if (/url$/.test(n) && !/^https:\/\/\S+$/i.test(v)) { return 'Must start with https://'; }
    if (/email$/.test(n) && !EMAIL_RE.test(v)) { return 'Enter a valid email address'; }
    return undefined;
  };
  // Payment API keys are never typed into the browser: they are entered in Power Automate after import.
  const isSecret = (schema: string): boolean => /razorpay|stripe|whatsapp/i.test(schema);
  const shown = (pack?.env || []).filter(e => !isSecret(e.schema));
  const anyInvalid = !pack || shown.some(e => !!invalid(e.schema, values[e.schema] || ''));
  const download = (): void => {
    downloadFlowPack(values).then(name => { setDone(name); notify(`Downloaded ${name}.`, MessageBarType.success); })
      .catch(e => setError(errorText(e)));
  };
  return (
    <Dialog hidden={false} onDismiss={onClose} minWidth={620} maxWidth={760}
      dialogContentProps={{ type: DialogType.largeHeader, title: 'Download Power Automate solution', subText: pack ? `Brillienta CRM Flows v${pack.version} · ${pack.flows.length} ready-made flows` : undefined }}>
      {!pack && !error && <Spinner label="Preparing the solution..." />}
      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
      {pack && !done && (
        <Stack tokens={{ childrenGap: 8 }}>
          <Text variant="small">These values are written into the solution so the import is pre-filled. You can still change them during or after import.</Text>
          {shown.map(e => (
            <TextField key={e.schema} label={e.name} description={e.description} value={values[e.schema] || ''}
              errorMessage={invalid(e.schema, values[e.schema] || '')} onChange={(_, v) => setValues({ ...values, [e.schema]: v || '' })} />
          ))}
        </Stack>
      )}
      {pack && done && (
        <Stack tokens={{ childrenGap: 8 }}>
          <MessageBar messageBarType={MessageBarType.success}>Downloaded <b>{done}</b>.</MessageBar>
          <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>Next steps</Text>
          <ol className="crm-ul">
            <li>Open <Link href="https://make.powerautomate.com" target="_blank">make.powerautomate.com</Link> and choose the right environment.</li>
            <li><b>Solutions › Import solution</b> › select the zip › Next.</li>
            <li>Create or pick the connections: <b>SharePoint</b>, <b>Office 365 Outlook</b>, <b>Approvals</b> (a service account is best).</li>
            <li>Check the environment variables (already filled in) › <b>Import</b>.</li>
            <li>Open the solution and <b>turn on</b> the flows you want.</li>
            <li>Payments: in the solution › <b>Environment variables</b>, set the Razorpay key id/secret or the Stripe secret key, then add the HTTP URL of the “payment received” flow as a webhook in Razorpay (<code>payment_link.paid</code>) or Stripe (<code>checkout.session.completed</code>).</li>
            <li>Website / AppSource: copy the HTTP URL of the flow into Vercel (<code>CRM_WEBHOOK_URL</code>) or Partner Center.</li>
          </ol>
          <Text variant="small" className="crm-muted">Flows included: {pack.flows.map(f => f.name.replace('Brillienta CRM - ', '')).join(' · ')}</Text>
        </Stack>
      )}
      <DialogFooter>
        {!done && <PrimaryButton iconProps={{ iconName: 'Download' }} text="Download .zip" disabled={anyInvalid} onClick={download} />}
        <DefaultButton text={done ? 'Close' : 'Cancel'} onClick={onClose} />
      </DialogFooter>
    </Dialog>
  );
};

const CopilotAgentCard: React.FC = () => {
  const { engine, notify } = useCrm();
  const company = engine.settings.company.name || '';
  const [name, setName] = React.useState(`${company ? `${company} ` : ''}CRM Copilot`.slice(0, 30));
  const [busy, setBusy] = React.useState(false);
  const site = engine.ds.rest.webUrl.replace(/\/$/, '');
  const here = window.location.href.split(/[?#]/)[0];
  const page = /\.aspx$/i.test(here) && here.toLowerCase().indexOf(site.toLowerCase()) === 0 ? here : `${site}/SitePages/Home.aspx`;
  const download = (): void => {
    setBusy(true);
    import(/* webpackChunkName: 'brillienta-copilot-agent' */ '../../services/CopilotAgent')
      .then(m => m.downloadCopilotAgent({ siteUrl: site, pageUrl: page, companyName: company, agentName: name.trim() || 'CRM Copilot' }))
      .then(f => { setBusy(false); notify(`Downloaded ${f}.`, MessageBarType.success); })
      .catch(e => { setBusy(false); notify(errorText(e), MessageBarType.error); });
  };
  return (
    <Card title="Microsoft 365 Copilot agent" wide>
      <Text variant="small" block>
        A declarative agent for Microsoft 365 Copilot that answers from this CRM site (deals, accounts, cases, renewals, invoices) with each person&apos;s own SharePoint permissions,
        and drafts customer emails. Needs Microsoft 365 Copilot licences. In every record, <b>Smart summary › Ask Copilot</b> hands the record over to Copilot.
      </Text>
      <Stack horizontal wrap verticalAlign="end" tokens={{ childrenGap: 12 }} style={{ margin: '10px 0 6px' }}>
        <TextField label="Agent name" value={name} maxLength={30} onChange={(_, v) => setName(v || '')} styles={{ root: { width: 260 } }} />
        <PrimaryButton iconProps={{ iconName: 'Download' }} text={busy ? 'Preparing…' : 'Download Copilot agent'} disabled={busy || !name.trim()} onClick={download} />
      </Stack>
      <ol className="crm-ul">
        <li>Teams admin center › <b>Manage apps › Upload new app</b> (or Microsoft 365 admin center › Integrated apps › Upload custom apps) › select the zip.</li>
        <li>Make it available to the sales and support teams. It appears in Copilot under <b>Agents</b>.</li>
        <li>Download again after renaming the agent; the same app is updated.</li>
      </ol>
    </Card>
  );
};

const TeamsSettings: React.FC<{ host: string; tenantSiteUrl: string }> = ({ host, tenantSiteUrl }) => {
  const { engine, notify } = useCrm();
  const here = engine.ds.rest.webUrl;
  const [current, setCurrent] = React.useState<string>(tenantSiteUrl);
  const [busy, setBusy] = React.useState(false);
  React.useEffect(() => { engine.ds.getTenantCrmSite(TENANT_SITE_KEY).then(v => setCurrent(v)).catch(() => undefined); }, []);
  const isDefault = !!current && current.replace(/\/$/, '').toLowerCase() === here.replace(/\/$/, '').toLowerCase();
  const setDefault = (): void => {
    setBusy(true);
    engine.ds.setTenantCrmSite(TENANT_SITE_KEY, here).then(() => { setBusy(false); setCurrent(here); notify('This site is now the default CRM for Teams and other sites.', MessageBarType.success); })
      .catch(e => { setBusy(false); notify(`${errorText(e)} — requires admin rights on the tenant app catalog. Alternative: Set-PnPStorageEntity -Key "${TENANT_SITE_KEY}" -Value "${here}" -Scope Tenant`, MessageBarType.error); });
  };
  return (
    <div className="crm-dash-grid" style={{ paddingTop: 12 }}>
      <Card title="Default CRM site">
        <Text block>Teams tabs run on the team’s own site and the personal app runs on the root site. The <b>default CRM site</b> tells them where the CRM data lives, so everyone works in one CRM.</Text>
        <MessageBar messageBarType={isDefault ? MessageBarType.success : MessageBarType.info} isMultiline styles={{ root: { margin: '10px 0' } }}>
          {current ? <>Current default: <b>{current}</b>{isDefault ? ' (this site)' : ''}</> : 'No default CRM site is set yet.'}
        </MessageBar>
        <PrimaryButton text={isDefault ? 'This site is the default' : 'Use this site as the default CRM'} disabled={busy || isDefault || host !== 'sharepoint'} onClick={setDefault} />
        {host !== 'sharepoint' && <Text variant="small" block className="crm-muted" style={{ marginTop: 6 }}>Open the CRM in SharePoint to change this.</Text>}
        <Text variant="small" block className="crm-muted" style={{ marginTop: 10 }}>Requires a SharePoint admin (app catalog). PowerShell alternative:<br />
          <code>Set-PnPStorageEntity -Key &quot;{TENANT_SITE_KEY}&quot; -Value &quot;{here}&quot; -Scope Tenant</code></Text>
      </Card>
      <Card title="Add Brillienta CRM to Teams">
        <ol className="crm-ul">
          <li>SharePoint admin center › <b>More features › Apps › App catalog</b> › select <i>Brillienta CRM</i> › <b>Sync to Teams</b>. (Or upload <code>BrillientaCRM-Teams.zip</code> in the Teams admin center › Manage apps.)</li>
          <li>Users open <b>Apps</b> in Teams, search “Brillienta CRM” and click <b>Add</b> — it appears as a personal app in the left rail.</li>
          <li>To add it to a team channel: <b>+</b> (Add a tab) › Brillienta CRM › leave the CRM site empty to use the default, or enter a site.</li>
          <li>Approve the Microsoft Graph API access once (SharePoint admin center › Advanced › API access) so email and Teams meetings work in Teams too.</li>
        </ol>
      </Card>
      <Card title="What works in Teams">
        <ul className="crm-ul">
          <li>Everything in the CRM: dashboards, records, pipeline, quotes, forecasting, reports, settings.</li>
          <li>Follows the Teams theme (light, dark, high contrast) automatically.</li>
          <li>Quotes: if Teams blocks the print window, the quote is downloaded as an HTML file — open it and choose Print › Save as PDF.</li>
          <li>Demo meetings are created in Outlook/Teams; links open in Teams.</li>
        </ul>
      </Card>
    </div>
  );
};

export const SettingsPage: React.FC<{ tenantId: string; host: string; tenantSiteUrl: string }> = ({ tenantId, host, tenantSiteUrl }) => {
  const { engine, settings, updateSettings, notify } = useCrm();
  const [s, setS] = React.useState<ICrmSettings>(() => JSON.parse(JSON.stringify(settings)));
  const [dirty, setDirty] = React.useState(false);
  const [busy, setBusy] = React.useState(false);
  const [progress, setProgress] = React.useState<string>();
  const upd = (patch: Partial<ICrmSettings>): void => { setS({ ...s, ...patch }); setDirty(true); };
  const isAdmin = engine.can('admin');

  const save = (after?: () => Promise<void>): void => {
    setBusy(true);
    updateSettings(s).then(() => (after ? after() : Promise.resolve())).then(() => { setBusy(false); setDirty(false); notify('Settings saved.', MessageBarType.success); })
      .catch(e => { setBusy(false); notify(errorText(e), MessageBarType.error); });
  };

  if (!isAdmin) {
    return <MessageBar messageBarType={MessageBarType.warning}>Only CRM administrators can change settings.</MessageBar>;
  }

  const setStage = (i: number, p: Partial<IStageDef>): void => upd({ stages: s.stages.map((x, j) => (j === i ? { ...x, ...p } : x)) });
  const moveStage = (i: number, d: number): void => {
    const a = [...s.stages];
    const j = i + d;
    if (j < 0 || j >= a.length) { return; }
    const t = a[i]; a[i] = a[j]; a[j] = t;
    upd({ stages: a });
  };
  const setHook = (id: string, p: Partial<IWebhook>): void => upd({ webhooks: s.webhooks.map(h => (h.id === id ? { ...h, ...p } : h)) });

  const [cf, setCf] = React.useState<ICustomField>({ entity: 'leads', key: '', label: '', type: 'text', choices: [], grid: false });
  const addCustom = (): void => {
    const key = (cf.key || cf.label).replace(/[^A-Za-z0-9]/g, '').slice(0, 30);
    if (!key || !cf.label) { return; }
    if (s.customFields.some(x => x.entity === cf.entity && x.key === key)) { notify('A custom field with this name already exists.', MessageBarType.warning); return; }
    upd({ customFields: [...s.customFields, { ...cf, key }] });
    setCf({ ...cf, key: '', label: '', choices: [] });
  };

  const backup = (): void => {
    setBusy(true);
    const keys = engine.defs.map(d => d.key);
    engine.ensure(keys).then(() => {
      const payload: any = { exportedAt: new Date().toISOString(), settings: s, data: {} };
      keys.forEach(k => { payload.data[k] = engine.data[k]; });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' }));
      a.download = `brillienta-crm-backup-${new Date().toISOString().slice(0, 10)}.json`;
      a.click();
      setBusy(false);
    }).catch(e => { setBusy(false); notify(errorText(e), MessageBarType.error); });
  };

  const lic = engine.license;
  const [flowDialog, setFlowDialog] = React.useState(false);

  return (
    <Stack tokens={{ childrenGap: 12 }}>
      <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center" className="crm-sticky-bar">
        <PrimaryButton text="Save settings" iconProps={{ iconName: 'Save' }} disabled={!dirty || busy} onClick={() => save()} />
        <DefaultButton text="Discard" disabled={!dirty || busy} onClick={() => { setS(JSON.parse(JSON.stringify(settings))); setDirty(false); }} />
        {busy && <Spinner size={SpinnerSize.small} label={progress} labelPosition="right" />}
        {dirty && <Text className="crm-muted">Unsaved changes</Text>}
      </Stack>
      {flowDialog && <FlowPackDialog onClose={() => setFlowDialog(false)} />}
      <Pivot overflowBehavior="menu">
        <PivotItem headerText="Company" itemIcon="CityNext">
          <div className="crm-dash-grid" style={{ paddingTop: 12 }}>
            <Card title="Company & branding (used on quotes)">
              <Stack tokens={{ childrenGap: 8 }}>
                <TextField label="Company name" value={s.company.name} onChange={(_, v) => upd({ company: { ...s.company, name: v || '' } })} />
                <TextField label="Address" multiline rows={3} value={s.company.address} onChange={(_, v) => upd({ company: { ...s.company, address: v || '' } })} />
                <TextField label="Email" value={s.company.email} onChange={(_, v) => upd({ company: { ...s.company, email: v || '' } })} />
                <TextField label="Phone" value={s.company.phone} onChange={(_, v) => upd({ company: { ...s.company, phone: v || '' } })} />
                <TextField label="Website" value={s.company.website} onChange={(_, v) => upd({ company: { ...s.company, website: v || '' } })} />
                <TextField label="Logo URL (https)" value={s.company.logoUrl} onChange={(_, v) => upd({ company: { ...s.company, logoUrl: v || '' } })} description="Upload the logo to a SharePoint library and paste its link." />
                <TextField label="Tax / VAT / GST ID" value={s.company.taxId} onChange={(_, v) => upd({ company: { ...s.company, taxId: v || '' } })} />
              </Stack>
            </Card>
            <Card title="Regional">
              <Stack tokens={{ childrenGap: 8 }}>
                <TextField label="Currency (ISO code)" value={s.currency} onChange={(_, v) => upd({ currency: (v || '').toUpperCase().slice(0, 3) })} description="e.g. USD, EUR, GBP, INR, AUD" />
                <Dropdown label="Fiscal year starts in" selectedKey={s.fiscalYearStartMonth}
                  options={Array.from({ length: 12 }, (_, i) => ({ key: i + 1, text: new Date(2000, i, 1).toLocaleDateString(undefined, { month: 'long' }) }))}
                  onChange={(_, o) => upd({ fiscalYearStartMonth: Number(o?.key) })} />
              </Stack>
            </Card>
            <Card title="Quotes & cases">
              <Stack tokens={{ childrenGap: 8 }}>
                <TextField label="Quote number prefix" value={s.quote.prefix} onChange={(_, v) => upd({ quote: { ...s.quote, prefix: v || 'Q' } })} />
                <TextField label="Quote validity (days)" type="number" value={String(s.quote.validityDays)} onChange={(_, v) => upd({ quote: { ...s.quote, validityDays: Number(v) || 30 } })} />
                <TextField label="Default terms & conditions" multiline rows={4} value={s.quote.terms} onChange={(_, v) => upd({ quote: { ...s.quote, terms: v || '' } })} />
                <TextField label="Quote footer" value={s.quote.footer} onChange={(_, v) => upd({ quote: { ...s.quote, footer: v || '' } })} />
                <TextField label="Discount needing manager approval (%)" type="number" value={String(s.quote.approvalDiscountPct)}
                  description="Quotes with a higher total discount must be approved (Power Automate flow “Quote discount approval”). 0 = off."
                  onChange={(_, v) => upd({ quote: { ...s.quote, approvalDiscountPct: Math.max(0, Number(v) || 0) } })} />
                <TextField label="E-signature link (HTTP URL of the flow “Quote e-signature”)" value={s.quote.signingUrl} placeholder="https://…"
                  description="Customers open this link to review and sign quotes. Leave empty to turn e-signature off."
                  errorMessage={s.quote.signingUrl && !/^https:\/\//i.test(s.quote.signingUrl) ? 'Must start with https://' : undefined}
                  onChange={(_, v) => upd({ quote: { ...s.quote, signingUrl: (v || '').trim() } })} />
                <Toggle label="Signed quote closes the opportunity as won" inlineLabel checked={s.quote.signatureClosesOpportunity}
                  onChange={(_, v) => upd({ quote: { ...s.quote, signatureClosesOpportunity: !!v } })} />
                <TextField label="Case number prefix" value={s.casePrefix} onChange={(_, v) => upd({ casePrefix: v || 'CS' })} />
              </Stack>
            </Card>
            <Card title="Invoices & subscriptions">
              <Stack tokens={{ childrenGap: 8 }}>
                <TextField label="Invoice number prefix" value={s.invoice.prefix} onChange={(_, v) => upd({ invoice: { ...s.invoice, prefix: v || 'INV' } })} />
                <TextField label="Payment terms (days)" type="number" value={String(s.invoice.paymentTermsDays)} onChange={(_, v) => upd({ invoice: { ...s.invoice, paymentTermsDays: Math.max(0, Number(v) || 0) } })} />
                <TextField label="Bank transfer details (printed on invoices)" multiline rows={3} value={s.invoice.bankDetails} onChange={(_, v) => upd({ invoice: { ...s.invoice, bankDetails: v || '' } })} />
                <TextField label="Invoice footer" value={s.invoice.footer} onChange={(_, v) => upd({ invoice: { ...s.invoice, footer: v || '' } })} />
                <TextField label="Create renewal opportunity (days before renewal)" type="number" value={String(s.renewals.leadDays)}
                  onChange={(_, v) => upd({ renewals: { ...s.renewals, leadDays: Math.max(1, Number(v) || 60) } })} />
                <Toggle label="Create renewal opportunities automatically" inlineLabel checked={s.renewals.autoCreateOpportunities}
                  onChange={(_, v) => upd({ renewals: { ...s.renewals, autoCreateOpportunities: !!v } })} />
                <Dropdown label="Online payment links" selectedKey={s.payments.provider}
                  options={[{ key: 'none', text: 'Off' }, { key: 'Razorpay', text: 'Razorpay (flow “Razorpay payment link”)' }, { key: 'Stripe', text: 'Stripe (flow “Stripe payment link”)' }]}
                  onChange={(_, o) => upd({ payments: { provider: (o?.key as 'none' | 'Razorpay' | 'Stripe') || 'none' } })} />
                <Text variant="small" className="crm-muted">API keys stay in Power Automate (environment variables), never in the browser.</Text>
              </Stack>
            </Card>
            <Card title="Accounting export">
              <Stack tokens={{ childrenGap: 8 }}>
                <Text variant="small" className="crm-muted">Invoices › ⋯ › Export for accounting: Zoho Books, QuickBooks Online, Xero, Tally or generic CSV.</Text>
                <TextField label="Tally sales ledger" value={s.accounting.tallySalesLedger} onChange={(_, v) => upd({ accounting: { ...s.accounting, tallySalesLedger: v || '' } })} />
                <TextField label="Tally tax ledger" value={s.accounting.tallyTaxLedger} onChange={(_, v) => upd({ accounting: { ...s.accounting, tallyTaxLedger: v || '' } })} />
                <TextField label="Xero sales account code" value={s.accounting.xeroSalesAccount} onChange={(_, v) => upd({ accounting: { ...s.accounting, xeroSalesAccount: v || '' } })} />
                <TextField label="Xero tax type" value={s.accounting.xeroTaxType} onChange={(_, v) => upd({ accounting: { ...s.accounting, xeroTaxType: v || '' } })} />
              </Stack>
            </Card>
          </div>
        </PivotItem>

        <PivotItem headerText="Sales process" itemIcon="Flow">
          <div className="crm-dash-grid" style={{ paddingTop: 12 }}>
            <Card title="Pipeline stages" wide>
              <Text variant="small" className="crm-muted" block>Order matters. “Closed Won” and “Closed Lost” must exist. Probability and forecast category are applied automatically when a deal changes stage.</Text>
              <div className="crm-stage-table" role="table" aria-label="Pipeline stages">
                <span className="crm-stage-head" role="columnheader">Stage</span>
                <span className="crm-stage-head" role="columnheader">Probability %</span>
                <span className="crm-stage-head crm-stage-head-forecast" role="columnheader">Forecast</span>
                <span />
                {s.stages.map((st, i) => {
                  const fixed = st.name === 'Closed Won' || st.name === 'Closed Lost';
                  return (
                    <React.Fragment key={i}>
                      <TextField ariaLabel="Stage name" value={st.name} disabled={fixed} onChange={(_, v) => setStage(i, { name: v || '' })} />
                      <TextField ariaLabel="Probability %" type="number" value={String(st.probability)} disabled={fixed} onChange={(_, v) => setStage(i, { probability: Number(v) || 0 })} />
                      <Dropdown className="crm-stage-forecast" ariaLabel="Forecast category" selectedKey={st.forecast} disabled={fixed} options={['Pipeline', 'Best case', 'Commit', 'Closed', 'Omitted'].map(x => ({ key: x, text: x }))}
                        onChange={(_, o) => setStage(i, { forecast: String(o?.key) })} />
                      <Stack horizontal>
                        <IconButton iconProps={{ iconName: 'Up' }} ariaLabel="Move up" disabled={i === 0} onClick={() => moveStage(i, -1)} />
                        <IconButton iconProps={{ iconName: 'Down' }} ariaLabel="Move down" disabled={i === s.stages.length - 1} onClick={() => moveStage(i, 1)} />
                        <IconButton iconProps={{ iconName: 'Delete' }} ariaLabel="Remove stage" disabled={fixed} style={fixed ? { visibility: 'hidden' } : undefined} onClick={() => upd({ stages: s.stages.filter((_, j) => j !== i) })} />
                      </Stack>
                    </React.Fragment>
                  );
                })}
              </div>
              <DefaultButton iconProps={{ iconName: 'Add' }} text="Add stage" style={{ marginTop: 8 }}
                onClick={() => upd({ stages: [...s.stages.slice(0, -2), { name: 'New stage', probability: 50, forecast: 'Pipeline' }, ...s.stages.slice(-2)] })} />
            </Card>
            <Card title="Pick lists (one per line)">
              <Stack tokens={{ childrenGap: 8 }}>
                <TextField label="Lead sources" multiline autoAdjustHeight rows={5} value={lines(s.leadSources)} onChange={(_, v) => upd({ leadSources: parseLines(v || '') })} />
                <TextField label="Industries" multiline autoAdjustHeight rows={5} value={lines(s.industries)} onChange={(_, v) => upd({ industries: parseLines(v || '') })} />
                <TextField label="Lost reasons" multiline autoAdjustHeight rows={4} value={lines(s.lostReasons)} onChange={(_, v) => upd({ lostReasons: parseLines(v || '') })} />
                <TextField label="Case categories" multiline autoAdjustHeight rows={4} value={lines(s.caseCategories)} onChange={(_, v) => upd({ caseCategories: parseLines(v || '') })} />
              </Stack>
            </Card>
            <Card title="Quarterly sales quotas">
              {s.quotas.map((qt, i) => (
                <Stack key={qt.user.id} horizontal verticalAlign="end" tokens={{ childrenGap: 8 }}>
                  <Text styles={{ root: { width: 160, paddingBottom: 6 } }}>{qt.user.title}</Text>
                  <TextField type="number" ariaLabel="Quota" value={String(qt.quarterly)} onChange={(_, v) => upd({ quotas: s.quotas.map((x, j) => (j === i ? { ...x, quarterly: Number(v) || 0 } : x)) })} styles={{ root: { width: 140 } }} />
                  <IconButton iconProps={{ iconName: 'Delete' }} ariaLabel="Remove" onClick={() => upd({ quotas: s.quotas.filter((_, j) => j !== i) })} />
                </Stack>
              ))}
              <UserPicker label="Add rep" showMe={false} onChange={u => { if (u && !s.quotas.some(x => x.user.id === u.id)) { upd({ quotas: [...s.quotas, { user: u, quarterly: 0 }] }); } }} />
            </Card>
          </div>
        </PivotItem>

        <PivotItem headerText="Automation" itemIcon="Robot">
          <div className="crm-dash-grid" style={{ paddingTop: 12 }}>
            <Card title="Lead assignment">
              <Toggle label="Automatically assign new leads (round-robin)" checked={s.assignment.autoAssignNewLeads} onChange={(_, c) => upd({ assignment: { ...s.assignment, autoAssignNewLeads: !!c } })} />
              <Label>Round-robin reps</Label>
              {s.assignment.roundRobin.map((u, i) => (
                <Stack key={u.id} horizontal verticalAlign="center"><Text style={{ flex: 1 }}>{i + 1}. {u.title}</Text>
                  <IconButton iconProps={{ iconName: 'Delete' }} ariaLabel="Remove" onClick={() => upd({ assignment: { ...s.assignment, roundRobin: s.assignment.roundRobin.filter(x => x.id !== u.id) } })} /></Stack>
              ))}
              <UserPicker showMe={false} onChange={u => { if (u && !s.assignment.roundRobin.some(x => x.id === u.id)) { upd({ assignment: { ...s.assignment, roundRobin: [...s.assignment.roundRobin, u] } }); } }} />
              <Text variant="small" className="crm-muted" block>Leads created in the CRM are assigned on save. Marketplace leads arriving via Power Automate can be distributed with “Round-robin assign” in the Leads grid.</Text>
            </Card>
            <Card title="Lead scoring (0–100)">
              <Stack tokens={{ childrenGap: 8 }}>
                <TextField label="Rating points" multiline rows={3} value={kv(s.scoring.rating)} onChange={(_, v) => upd({ scoring: { ...s.scoring, rating: parseKv(v || '') } })} />
                <TextField label="Source points" multiline rows={5} value={kv(s.scoring.sources)} onChange={(_, v) => upd({ scoring: { ...s.scoring, sources: parseKv(v || '') } })} />
                <TextField label="Job title keywords (best match counts)" multiline rows={5}
                  value={s.scoring.titleKeywords.map(k => `${k.keyword} = ${k.points}`).join('\n')}
                  onChange={(_, v) => { const m = parseKv(v || ''); upd({ scoring: { ...s.scoring, titleKeywords: Object.keys(m).map(k => ({ keyword: k, points: m[k] })) } }); }} />
                <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
                  <TextField label="Has phone" type="number" value={String(s.scoring.hasPhone)} onChange={(_, v) => upd({ scoring: { ...s.scoring, hasPhone: Number(v) || 0 } })} styles={{ root: { width: 100 } }} />
                  <TextField label="Has company" type="number" value={String(s.scoring.hasCompany)} onChange={(_, v) => upd({ scoring: { ...s.scoring, hasCompany: Number(v) || 0 } })} styles={{ root: { width: 100 } }} />
                  <TextField label="Per activity" type="number" value={String(s.scoring.perActivity)} onChange={(_, v) => upd({ scoring: { ...s.scoring, perActivity: Number(v) || 0 } })} styles={{ root: { width: 100 } }} />
                  <TextField label="Activity max" type="number" value={String(s.scoring.maxActivityPoints)} onChange={(_, v) => upd({ scoring: { ...s.scoring, maxActivityPoints: Number(v) || 0 } })} styles={{ root: { width: 100 } }} />
                </Stack>
              </Stack>
            </Card>
            <Card title="Service level agreements">
              <div className="crm-sla-table">
                <Text variant="small" className="crm-muted">Priority</Text><Text variant="small" className="crm-muted">First response (h)</Text><Text variant="small" className="crm-muted">Resolve (h)</Text>
                {['Critical', 'High', 'Normal', 'Low'].map(p => (
                  <React.Fragment key={p}>
                    <Text>{p}</Text>
                    <TextField ariaLabel={`${p} first response hours`} type="number" value={String(s.firstResponseHours[p] ?? '')} onChange={(_, v) => upd({ firstResponseHours: { ...s.firstResponseHours, [p]: Math.max(0, Number(v) || 0) } })} />
                    <TextField ariaLabel={`${p} resolution hours`} type="number" value={String(s.slaHours[p] ?? '')} onChange={(_, v) => upd({ slaHours: { ...s.slaHours, [p]: Math.max(0, Number(v) || 0) } })} />
                  </React.Fragment>
                ))}
              </div>
              <Toggle label="Pause the SLA clock while “Waiting on customer”" checked={s.slaPauseWhenWaiting} onChange={(_, c) => upd({ slaPauseWhenWaiting: !!c })} />
              <Toggle label="Count business hours only" checked={s.businessHours.enabled} onChange={(_, c) => upd({ businessHours: { ...s.businessHours, enabled: !!c } })} />
              {s.businessHours.enabled && (
                <Stack tokens={{ childrenGap: 8 }}>
                  <TextField label="Time zone (IANA)" value={s.businessHours.timeZone} description="e.g. Asia/Kolkata, Europe/London, America/New_York"
                    onChange={(_, v) => upd({ businessHours: { ...s.businessHours, timeZone: (v || '').trim() } })} />
                  <Stack horizontal tokens={{ childrenGap: 8 }}>
                    <TextField label="Opens" value={s.businessHours.start} placeholder="09:00" styles={{ root: { width: 100 } }} onChange={(_, v) => upd({ businessHours: { ...s.businessHours, start: v || '' } })} />
                    <TextField label="Closes" value={s.businessHours.end} placeholder="18:00" styles={{ root: { width: 100 } }} onChange={(_, v) => upd({ businessHours: { ...s.businessHours, end: v || '' } })} />
                  </Stack>
                  <Stack horizontal wrap tokens={{ childrenGap: 10 }}>
                    {['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'].map((d, i) => (
                      <Checkbox key={d} label={d} checked={s.businessHours.days.indexOf(i) >= 0}
                        onChange={(_, c) => upd({ businessHours: { ...s.businessHours, days: c ? [...s.businessHours.days, i].sort() : s.businessHours.days.filter(x => x !== i) } })} />
                    ))}
                  </Stack>
                  <TextField label="Holidays (YYYY-MM-DD, one per line)" multiline rows={3} value={s.businessHours.holidays.join('\n')}
                    onChange={(_, v) => upd({ businessHours: { ...s.businessHours, holidays: (v || '').split(/\s+/).filter(x => /^\d{4}-\d{2}-\d{2}$/.test(x)) } })} />
                </Stack>
              )}
            </Card>
            <Card title="Data visibility & compliance">
              <Toggle label="Sales reps only see their own records" checked={s.restrictRepsToOwnRecords} onChange={(_, c) => upd({ restrictRepsToOwnRecords: !!c })} />
              <Text variant="small" className="crm-muted" block>Applies to the CRM screens. For hard security, use item-level permissions in SharePoint.</Text>
              <Toggle label="Audit log (who changed what)" checked={s.auditEnabled} disabled={!engine.hasFeature('audit')} onChange={(_, c) => upd({ auditEnabled: !!c })} />
            </Card>
          </div>
        </PivotItem>

        <PivotItem headerText="Integrations" itemIcon="PlugConnected">
          <div className="crm-dash-grid" style={{ paddingTop: 12 }}>
            <Card title="Microsoft 365 (Outlook & Teams)">
              <Toggle label="Send emails from Outlook (Microsoft Graph Mail.Send)" checked={s.integrations.graphEmail} onChange={(_, c) => upd({ integrations: { ...s.integrations, graphEmail: !!c } })} />
              <Toggle label="Create Teams meetings for demos (Calendars.ReadWrite)" checked={s.integrations.graphMeetings} onChange={(_, c) => upd({ integrations: { ...s.integrations, graphMeetings: !!c } })} />
              <Toggle label="Allow automatic Outlook email & meeting logging (Mail.Read)" checked={s.integrations.outlookSync} onChange={(_, c) => upd({ integrations: { ...s.integrations, outlookSync: !!c } })} />
              {s.integrations.outlookSync && (
                <Stack tokens={{ childrenGap: 6 }} style={{ marginBottom: 8 }}>
                  <Text variant="small" className="crm-muted">Each user switches it on in <b>My work</b>. Only email and meetings with contacts and leads already in the CRM are logged.</Text>
                  <TextField label="First sync reads the last (days)" type="number" value={String(s.integrations.syncLookbackDays)} onChange={(_, v) => upd({ integrations: { ...s.integrations, syncLookbackDays: Math.max(1, Math.min(365, Number(v) || 30)) } })} styles={{ root: { maxWidth: 220 } }} />
                  <TextField label="Never log these domains (one per line)" multiline autoAdjustHeight rows={2} value={lines(s.integrations.internalDomains)}
                    description="Your own email domain is always excluded." onChange={(_, v) => upd({ integrations: { ...s.integrations, internalDomains: parseLines(v || '') } })} />
                </Stack>
              )}
              <MessageBar isMultiline>A SharePoint/Global admin must approve the API access requests once: <b>SharePoint admin center › Advanced › API access</b> → approve “Microsoft Graph: Mail.Send, Mail.Read, Calendars.ReadWrite, User.ReadBasic.All”.</MessageBar>
            </Card>
            <Card title="Teams channel notifications">
              <TextField label="Teams Workflows webhook URL" value={s.integrations.teamsWebhookUrl} disabled={!engine.hasFeature('webhooks')}
                onChange={(_, v) => upd({ integrations: { ...s.integrations, teamsWebhookUrl: v || '' } })}
                description='In Teams: channel › ⋯ › Workflows › "Post to a channel when a webhook request is received". Posts deals won, new leads, demo requests, cases.' />
              <DefaultButton text="Send test" disabled={!s.integrations.teamsWebhookUrl} style={{ marginTop: 8 }}
                onClick={() => engine.webhooks.test(s.integrations.teamsWebhookUrl).then(() => notify('Test sent.', MessageBarType.success)).catch(e => notify(errorText(e), MessageBarType.error))} />
            </Card>
            <CopilotAgentCard />
            <Card title="WhatsApp Business" wide>
              <Stack tokens={{ childrenGap: 8 }}>
                <Stack horizontal wrap tokens={{ childrenGap: 16 }} verticalAlign="end">
                  <Toggle label="WhatsApp on records" checked={s.whatsapp.enabled} onChange={(_, c) => upd({ whatsapp: { ...s.whatsapp, enabled: !!c } })} />
                  <Dropdown label="How messages are sent" selectedKey={s.whatsapp.mode} styles={{ root: { width: 300 } }}
                    options={[{ key: 'click', text: 'Click to chat (WhatsApp / WhatsApp Business app)' }, { key: 'cloud', text: 'WhatsApp Business Cloud API (Power Automate)' }]}
                    onChange={(_, o) => upd({ whatsapp: { ...s.whatsapp, mode: (o?.key as 'click' | 'cloud') || 'click' } })} />
                  <TextField label="Default country code" prefix="+" value={s.whatsapp.defaultCountryCode} styles={{ root: { width: 120 } }}
                    onChange={(_, v) => upd({ whatsapp: { ...s.whatsapp, defaultCountryCode: (v || '').replace(/\D/g, '').slice(0, 4) } })} />
                </Stack>
                <Text variant="small" className="crm-muted">
                  Cloud API: import the flows <b>WhatsApp send</b> and <b>WhatsApp inbound</b>, set the phone number ID and access token in Power Automate, and add the inbound flow URL as the webhook in Meta › WhatsApp › Configuration.
                  Every message, sent or received, is logged on the contact or lead timeline; unknown senders become new leads.
                </Text>
                {s.whatsapp.templates.map((t, i) => {
                  const setT = (patch: Partial<typeof t>): void => upd({ whatsapp: { ...s.whatsapp, templates: s.whatsapp.templates.map((x, j) => (j === i ? { ...x, ...patch } : x)) } });
                  return (
                    <div key={t.id} className="crm-wa-tpl">
                      <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="end">
                        <TextField label="Template" value={t.name} onChange={(_, v) => setT({ name: v || '' })} styles={{ root: { width: 180 } }} />
                        <TextField label="Approved template name (Cloud API)" value={t.cloudName || ''} onChange={(_, v) => setT({ cloudName: (v || '').trim() })} styles={{ root: { width: 220 } }} />
                        <TextField label="Language" value={t.language || ''} placeholder="en" onChange={(_, v) => setT({ language: (v || '').trim() })} styles={{ root: { width: 80 } }} />
                        <TextField label="{{1}}, {{2}}… values" value={(t.params || []).join(', ')} placeholder="firstName, quoteNumber"
                          onChange={(_, v) => setT({ params: (v || '').split(',').map(x => x.trim()).filter(Boolean) })} styles={{ root: { width: 240 } }} />
                        <IconButton iconProps={{ iconName: 'Delete' }} ariaLabel="Remove template" onClick={() => upd({ whatsapp: { ...s.whatsapp, templates: s.whatsapp.templates.filter((_, j) => j !== i) } })} />
                      </Stack>
                      <TextField label="Message" multiline rows={2} value={t.body} onChange={(_, v) => setT({ body: v || '' })} />
                    </div>
                  );
                })}
                <DefaultButton iconProps={{ iconName: 'Add' }} text="Add WhatsApp template" styles={{ root: { alignSelf: 'flex-start' } }}
                  onClick={() => upd({ whatsapp: { ...s.whatsapp, templates: [...s.whatsapp.templates, { id: `t${Date.now()}`, name: 'New template', body: 'Hi {{firstName}}, ', language: 'en', params: ['firstName'] }] } })} />
                <Text variant="small" className="crm-muted">Placeholders: {'{{firstName}} {{fullName}} {{company}} {{senderName}} {{companyName}} {{quoteNumber}} {{invoiceNumber}} {{total}} {{paymentLink}} {{scheduledDate}} {{meetingLink}} {{caseNumber}}'}</Text>
              </Stack>
            </Card>
            <Card title="Power Automate flow pack" wide>
              <Stack horizontal wrap verticalAlign="center" tokens={{ childrenGap: 12 }} style={{ marginBottom: 6 }}>
                <PrimaryButton iconProps={{ iconName: 'Download' }} text="Download Power Automate solution" onClick={() => setFlowDialog(true)} />
                <Text variant="small" className="crm-muted">Pre-filled for this site · import in <b>Power Automate › Solutions › Import solution</b></Text>
              </Stack>
              <ul className="crm-ul">
                <li><b>AppSource lead intake</b> – Partner Center HTTPS endpoint → leads list <i>(Power Automate Premium)</i></li>
                <li><b>Website leads &amp; demo bookings</b> – brillienta.com demo bookings → lead + demo request; support requests → case with SLA; enquiries → lead <i>(Premium)</i></li>
                <li><b>New lead alert</b> – emails the owner; unassigned leads go to the sales manager</li>
                <li><b>Daily sales digest</b> – new leads, deals won, deals closing, overdue tasks (weekdays)</li>
                <li><b>Overdue reminders</b> – one email per person with their overdue tasks and follow-ups</li>
                <li><b>Case SLA escalation</b> – hourly; escalates newly breached cases</li>
                <li><b>Deal won handoff</b> – notifies finance and creates a customer folder</li>
                <li><b>Quote discount approval</b> – Approvals in Teams/Outlook for quotes over {s.quote.approvalDiscountPct}% discount</li>
                <li><b>Quote e-signature</b> – signing page for customers; records name, title, IP and time <i>(Premium)</i></li>
                <li><b>WhatsApp send / inbound</b> – WhatsApp Business Cloud API: queued messages and templates out, replies and new leads in <i>(Premium)</i></li>
                <li><b>Email to case</b> – support mailbox → cases; replies with [case number] go to the same case</li>
                <li><b>Razorpay / Stripe payment link</b> – creates a pay link when an invoice asks for one <i>(Premium)</i></li>
                <li><b>Razorpay / Stripe payment received</b> – payment webhook marks the invoice Paid <i>(Premium)</i></li>
              </ul>
            </Card>
            <Card title="Microsoft AppSource / Partner Center leads">
              <Text variant="small" block>
                Partner Center › your offer › Customer leads › <b>HTTPS endpoint</b> → Power Automate flow “When an HTTP request is received” → SharePoint “Create item” in <b>{engine.def('leads').listTitle}</b>.
                Map UserDetails.FirstName/LastName/Email/Phone/Company/Country, Title → JobTitle, OfferTitle → Offer, LeadSource, ActionCode, Description; set Status = New and LeadSource = AppSource.
                New marketplace leads are then scored and appear in the CRM automatically.
              </Text>
            </Card>
            <Card title="Outgoing webhooks (Power Automate, Zapier, Make, n8n, your API)" wide>
              {!engine.hasFeature('webhooks') && <MessageBar messageBarType={MessageBarType.warning}>Webhooks require the Enterprise plan.</MessageBar>}
              <Text variant="small" className="crm-muted" block>Each event is POSTed as JSON (Content-Type text/plain; in Power Automate use <code>json(triggerBody())</code>). Payload: {'{ event, occurredAt, tenant, summary, data }'}.</Text>
              {s.webhooks.map(h => (
                <div key={h.id} className="crm-card" style={{ marginTop: 8 }}>
                  <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="end">
                    <TextField label="Name" value={h.name} onChange={(_, v) => setHook(h.id, { name: v || '' })} styles={{ root: { width: 180 } }} />
                    <TextField label="URL (https)" value={h.url} onChange={(_, v) => setHook(h.id, { url: v || '' })} styles={{ root: { flex: 1, minWidth: 260 } }} />
                    <Toggle label="Active" checked={h.active} onChange={(_, c) => setHook(h.id, { active: !!c })} />
                    <DefaultButton text="Test" onClick={() => engine.webhooks.test(h.url).then(() => notify('Test sent.', MessageBarType.success)).catch(e => notify(errorText(e), MessageBarType.error))} />
                    <IconButton iconProps={{ iconName: 'Delete' }} ariaLabel="Delete webhook" onClick={() => upd({ webhooks: s.webhooks.filter(x => x.id !== h.id) })} />
                  </Stack>
                  <Dropdown label="Events" multiSelect selectedKeys={h.events} options={[{ key: '*', text: 'All events' }, ...WEBHOOK_EVENTS.map(e => ({ key: e, text: e }))]}
                    onChange={(_, o) => o && setHook(h.id, { events: o.selected ? [...h.events, String(o.key)] : h.events.filter(x => x !== o.key) })} />
                </div>
              ))}
              <DefaultButton iconProps={{ iconName: 'Add' }} text="Add webhook" style={{ marginTop: 8 }} disabled={!engine.hasFeature('webhooks')}
                onClick={() => upd({ webhooks: [...s.webhooks, { id: uid(), name: 'New webhook', url: '', events: ['opportunity.won'], active: true }] })} />
            </Card>
            <Card title="Email templates" wide>
              <Text variant="small" className="crm-muted" block>{'Placeholders: {{firstName}} {{lastName}} {{fullName}} {{company}} {{product}} {{ownerName}} {{senderName}} {{scheduledDate}} {{meetingLink}} {{quoteNumber}} {{total}} {{companyName}}'}</Text>
              {s.emailTemplates.map((t, i) => (
                <div key={t.id} className="crm-card" style={{ marginTop: 8 }}>
                  <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="end">
                    <TextField label="Name" value={t.name} onChange={(_, v) => upd({ emailTemplates: s.emailTemplates.map((x, j) => (j === i ? { ...x, name: v || '' } : x)) })} styles={{ root: { width: 200 } }} />
                    <TextField label="Subject" value={t.subject} onChange={(_, v) => upd({ emailTemplates: s.emailTemplates.map((x, j) => (j === i ? { ...x, subject: v || '' } : x)) })} styles={{ root: { flex: 1 } }} />
                    <IconButton iconProps={{ iconName: 'Delete' }} ariaLabel="Delete template" onClick={() => upd({ emailTemplates: s.emailTemplates.filter((_, j) => j !== i) })} />
                  </Stack>
                  <TextField label="Body" multiline rows={5} value={t.body} onChange={(_, v) => upd({ emailTemplates: s.emailTemplates.map((x, j) => (j === i ? { ...x, body: v || '' } : x)) })} />
                </div>
              ))}
              <DefaultButton iconProps={{ iconName: 'Add' }} text="Add template" style={{ marginTop: 8 }} onClick={() => upd({ emailTemplates: [...s.emailTemplates, { id: uid(), name: 'New template', subject: '', body: '' }] })} />
            </Card>
          </div>
        </PivotItem>

        <PivotItem headerText="Custom fields" itemIcon="FieldNotChanged">
          <div className="crm-dash-grid" style={{ paddingTop: 12 }}>
            {!engine.hasFeature('customFields') && <MessageBar messageBarType={MessageBarType.warning}>Custom fields require the Enterprise plan.</MessageBar>}
            <Card title="Add a field">
              <Stack tokens={{ childrenGap: 8 }}>
                <Dropdown label="Record type" selectedKey={cf.entity} options={ENTITIES.map(e => ({ key: e.key, text: e.singular }))} onChange={(_, o) => setCf({ ...cf, entity: o?.key as EntityKey })} />
                <TextField label="Label" value={cf.label} onChange={(_, v) => setCf({ ...cf, label: v || '' })} />
                <Dropdown label="Type" selectedKey={cf.type} onChange={(_, o) => setCf({ ...cf, type: o?.key as FieldType })} options={[
                  { key: 'text', text: 'Text' }, { key: 'note', text: 'Multi-line text' }, { key: 'number', text: 'Number' }, { key: 'currency', text: 'Currency' },
                  { key: 'date', text: 'Date' }, { key: 'datetime', text: 'Date & time' }, { key: 'boolean', text: 'Yes / No' }, { key: 'choice', text: 'Choice' },
                  { key: 'email', text: 'Email' }, { key: 'phone', text: 'Phone' }, { key: 'url', text: 'Link' }, { key: 'user', text: 'Person' }
                ]} />
                {cf.type === 'choice' && <TextField label="Choices (one per line)" multiline rows={4} value={lines(cf.choices || [])} onChange={(_, v) => setCf({ ...cf, choices: parseLines(v || '') })} />}
                <Checkbox label="Show as a grid column" checked={!!cf.grid} onChange={(_, c) => setCf({ ...cf, grid: !!c })} />
                <DefaultButton text="Add field" disabled={!cf.label || !engine.hasFeature('customFields')} onClick={addCustom} />
              </Stack>
            </Card>
            <Card title="Custom fields">
              {s.customFields.length === 0 && <Text className="crm-muted">None yet.</Text>}
              {s.customFields.map(c => (
                <Stack key={`${c.entity}-${c.key}`} horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                  <Text style={{ flex: 1 }}><b>{c.label}</b> <span className="crm-muted">· {engine.def(c.entity).singular} · {c.type} · column cf{c.key}</span></Text>
                  <IconButton iconProps={{ iconName: 'Delete' }} ariaLabel="Remove field" title="Hide field (the SharePoint column and its data are kept)"
                    onClick={() => upd({ customFields: s.customFields.filter(x => !(x.entity === c.entity && x.key === c.key)) })} />
                </Stack>
              ))}
              <PrimaryButton text="Save & create columns" style={{ marginTop: 12 }} disabled={busy}
                onClick={() => save(() => engine.ds.provision(setProgress))} />
            </Card>
          </div>
        </PivotItem>

        <PivotItem headerText="Security" itemIcon="Lock">
          <div className="crm-dash-grid" style={{ paddingTop: 12 }}>
            <Card title="Roles">
              <Text block>Your role: <b>{engine.role}</b></Text>
              <ul className="crm-ul">
                <li><b>Administrators</b> – everything, including settings, security and licence.</li>
                <li><b>Sales managers</b> – all records, delete, import, bulk assignment, forecasts.</li>
                <li><b>Sales</b> – work with leads, accounts, contacts, deals, quotes, demos, tasks.</li>
                <li><b>Support</b> – cases, accounts, contacts, activities.</li>
                <li><b>Readers</b> – read-only.</li>
              </ul>
              <Text variant="small" className="crm-muted" block>Site owners are always administrators. Users in no CRM group get Sales manager rights if they can edit the lists, otherwise read-only.</Text>
            </Card>
            <Card title="SharePoint groups">
              <Stack tokens={{ childrenGap: 6 }}>
                {(['admins', 'managers', 'sales', 'support', 'readers'] as const).map(k => (
                  <TextField key={k} label={k[0].toUpperCase() + k.slice(1)} value={s.groups[k]} onChange={(_, v) => upd({ groups: { ...s.groups, [k]: v || '' } })} />
                ))}
                <PrimaryButton text="Create groups & apply list permissions" disabled={busy || !engine.hasFeature('security')} onClick={() => save(async () => { await engine.ds.applySecurity(setProgress); })} />
                {!engine.hasFeature('security') && <Text variant="small">Requires the Enterprise plan.</Text>}
                <Text variant="small" className="crm-muted">Then add people to the groups in <Link href={`${engine.ds.rest.webUrl}/_layouts/15/groups.aspx`} target="_blank">Site permissions › Groups</Link>. Existing site permissions are kept.</Text>
              </Stack>
            </Card>
          </div>
        </PivotItem>

        <PivotItem headerText="Microsoft Teams" itemIcon="TeamsLogo">
          <TeamsSettings host={host} tenantSiteUrl={tenantSiteUrl} />
        </PivotItem>

        <PivotItem headerText="Licence" itemIcon="Certificate">
          <div className="crm-dash-grid" style={{ paddingTop: 12 }}>
            <Card title="Licence">
              <MessageBar messageBarType={lic.status === 'active' ? MessageBarType.success : lic.status === 'trial' ? MessageBarType.info : MessageBarType.severeWarning} isMultiline>{lic.message}</MessageBar>
              <TextField label="Your Microsoft 365 tenant ID" readOnly value={tenantId} description="Send this to your vendor to receive a licence key." />
              <TextField label="Licence key" multiline rows={4} value={s.licenseKey} onChange={(_, v) => upd({ licenseKey: (v || '').trim() })} />
              <Text variant="small" className="crm-muted" block>Save settings to activate the key.</Text>
            </Card>
            <Card title="Plans">
              <ul className="crm-ul">
                <li><b>Standard</b> – leads, accounts, contacts, opportunities & pipeline, quotes & products, demo requests, activities & tasks, dashboards, reports, calendar, Outlook email & Teams meetings, import/export, AppSource leads.</li>
                <li><b>Enterprise</b> – everything in Standard plus cases & SLAs, campaigns & ROI, forecasting & quotas, webhooks & Teams notifications, custom fields, audit log, role-based security groups.</li>
                <li><b>Trial</b> – all Enterprise features for 30 days.</li>
              </ul>
            </Card>
            <Card title="Data">
              <DefaultButton iconProps={{ iconName: 'Download' }} text="Download full backup (JSON)" disabled={busy} onClick={backup} />
              <Text variant="small" className="crm-muted" block style={{ marginTop: 8 }}>All CRM data also lives in SharePoint lists on this site and is covered by Microsoft 365 retention, versioning and backup.</Text>
            </Card>
          </div>
        </PivotItem>
      </Pivot>
    </Stack>
  );
};
