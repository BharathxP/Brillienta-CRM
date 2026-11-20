/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
import {
  Stack, TextField, PrimaryButton, DefaultButton, Toggle, Dropdown, MessageBar, MessageBarType, Text, IconButton,
  Link, Checkbox, Spinner, SpinnerSize, Label, Dialog, DialogType, DialogFooter, Icon, SearchBox
} from '@fluentui/react';
import { downloadFlowPack, IFlowPackInfo, loadFlowPack } from '../../services/FlowPackage';
import { ENTITIES, EntityKey, FieldType } from '../../models/model';
import { ICrmSettings, ICustomField, IStageDef, IWebhook, WEBHOOK_EVENTS } from '../../models/settings';
import { useCrm, useEntities } from '../CrmContext';
import { UserPicker } from '../core/controls';
import { errorText, lsGet, lsSet, uid } from '../utils';

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
  const isSecret = (schema: string): boolean => /razorpay|stripe|whatsapp|apikey|apollo|portalurl/i.test(schema);
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
            <li>Create or pick the connections: <b>SharePoint</b>, <b>Office 365 Outlook</b>, <b>Approvals</b> and <b>Content Conversion</b> (a service account is best).</li>
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

// ================================================================ layout primitives

/** A titled block of settings on a settings page. */
const Group: React.FC<{ title: string; description?: React.ReactNode; action?: React.ReactNode; children?: React.ReactNode }> = ({ title, description, action, children }) => (
  <section className="crm-set-group">
    <div className="crm-set-group-head">
      <div style={{ flex: 1, minWidth: 0 }}>
        <Text as="h3" styles={{ root: { fontSize: 15, fontWeight: 600, margin: 0 } }}>{title}</Text>
        {description && <Text variant="small" className="crm-muted" block styles={{ root: { marginTop: 2 } }}>{description}</Text>}
      </div>
      {action}
    </div>
    <div className="crm-set-group-body">{children}</div>
  </section>
);

/** One setting: label and help on the left, the control on the right (stacked on narrow screens). */
const Row: React.FC<{ label: string; help?: React.ReactNode; wide?: boolean; children?: React.ReactNode }> = ({ label, help, wide, children }) => (
  <div className={`crm-set-row${wide ? ' crm-set-row-wide' : ''}`}>
    <div className="crm-set-label">
      <Text styles={{ root: { fontWeight: 600 } }}>{label}</Text>
      {help && <Text variant="small" className="crm-muted" block>{help}</Text>}
    </div>
    <div className="crm-set-control">{children}</div>
  </div>
);

/** A repeatable item (stage, template, team…) inside a group. */
const Item: React.FC<{ children?: React.ReactNode }> = ({ children }) => <div className="crm-set-item">{children}</div>;

/** A field with a unit before or after it ("+", "days", "%"). Keeps the field's own accessible name, which Fluent's prefix/suffix would replace. */
const Unit: React.FC<{ before?: string; after?: string; children?: React.ReactNode }> = ({ before, after, children }) => (
  <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 6 }}>
    {before && <Text className="crm-muted" aria-hidden styles={{ root: { paddingBottom: 6 } }}>{before}</Text>}{children}{after && <Text className="crm-muted" aria-hidden styles={{ root: { paddingBottom: 6 } }}>{after}</Text>}
  </Stack>
);

const urlError = (v: string): string | undefined => (v && !/^https:\/\//i.test(v) ? 'Must start with https://' : undefined);

// ================================================================ self-contained sections

const CurrenciesGroup: React.FC<{ s: ICrmSettings; upd: (patch: Partial<ICrmSettings>) => void }> = ({ s, upd }) => (
  <Group title="Other currencies" description={`Deals, quotes, invoices and subscriptions can use these currencies. Dashboards, forecasts and reports convert them to ${s.currency}.`}
    action={<DefaultButton iconProps={{ iconName: 'Add' }} text="Add currency" onClick={() => upd({ currencies: [...s.currencies, { code: '', rate: 1 }] })} />}>
    {s.currencies.length === 0 && <Text className="crm-muted" block styles={{ root: { padding: '10px 0' } }}>Only {s.currency} is used.</Text>}
    {s.currencies.map((c, i) => (
      <Row key={i} label={c.code || 'New currency'} help={c.code && c.rate ? `1 ${c.code} = ${c.rate} ${s.currency}` : 'Code and exchange rate'}>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
          <TextField ariaLabel="Currency code" value={c.code} placeholder="EUR" styles={{ root: { width: 90 } }}
            onChange={(_, v) => upd({ currencies: s.currencies.map((x, j) => (j === i ? { ...x, code: (v || '').toUpperCase().slice(0, 3) } : x)) })} />
          <Unit before="=" after={s.currency}><TextField ariaLabel={`Rate in ${s.currency}`} type="number" value={String(c.rate)} styles={{ root: { width: 170 } }}
            onChange={(_, v) => upd({ currencies: s.currencies.map((x, j) => (j === i ? { ...x, rate: Number(v) || 0 } : x)) })} /></Unit>
          <IconButton iconProps={{ iconName: 'Delete' }} ariaLabel="Remove currency" onClick={() => upd({ currencies: s.currencies.filter((_, j) => j !== i) })} />
        </Stack>
      </Row>
    ))}
  </Group>
);

const PriceBooksGroup: React.FC<{ s: ICrmSettings; upd: (patch: Partial<ICrmSettings>) => void }> = ({ s, upd }) => {
  const { rows } = useEntities('products');
  const products = rows('products').filter(p => p.active !== false);
  const [open, setOpen] = React.useState<string>();
  const codes = [s.currency, ...s.currencies.map(c => c.code)].filter(Boolean);
  return (
    <Group title="Price books" description="A price book sets a quote's currency and prices (for example “Partners EUR” or “Education INR”). Products without a specific price use the list price, converted and minus the book discount."
      action={<DefaultButton iconProps={{ iconName: 'Add' }} text="New price book"
        onClick={() => upd({ priceBooks: [...s.priceBooks, { id: `pb${Date.now()}`, name: 'New price book', currency: s.currency, discountPct: 0, prices: {}, active: true }] })} />}>
      {s.priceBooks.length === 0 && <Text className="crm-muted" block styles={{ root: { padding: '10px 0' } }}>No price books yet: quotes use the product list prices.</Text>}
      {s.priceBooks.map((b, i) => {
        const setB = (patch: Partial<typeof b>): void => upd({ priceBooks: s.priceBooks.map((x, j) => (j === i ? { ...x, ...patch } : x)) });
        const count = Object.keys(b.prices).filter(k => b.prices[k] !== undefined).length;
        return (
          <Item key={b.id}>
            <Stack horizontal wrap verticalAlign="end" tokens={{ childrenGap: 10 }}>
              <TextField label="Name" value={b.name} onChange={(_, v) => setB({ name: v || '' })} styles={{ root: { width: 220 } }} />
              <Dropdown label="Currency" selectedKey={b.currency} options={codes.map(c => ({ key: c, text: c }))} onChange={(_, o) => setB({ currency: String(o?.key) })} styles={{ root: { width: 100 } }} />
              <Unit after="%"><TextField label="Discount on list" type="number" value={String(b.discountPct)} onChange={(_, v) => setB({ discountPct: Math.max(0, Math.min(100, Number(v) || 0)) })} styles={{ root: { width: 130 } }} /></Unit>
              <Toggle label="Active" checked={b.active} onChange={(_, c) => setB({ active: !!c })} />
              <DefaultButton iconProps={{ iconName: open === b.id ? 'ChevronUp' : 'ChevronDown' }} text={`Specific prices (${count})`} onClick={() => setOpen(open === b.id ? undefined : b.id)} />
              <IconButton iconProps={{ iconName: 'Delete' }} ariaLabel="Remove price book" onClick={() => upd({ priceBooks: s.priceBooks.filter((_, j) => j !== i) })} />
            </Stack>
            {open === b.id && (
              <div className="crm-sla-table" style={{ marginTop: 12 }}>
                <Text variant="small" className="crm-muted">Product</Text><Text variant="small" className="crm-muted">List ({s.currency})</Text><Text variant="small" className="crm-muted">Price ({b.currency})</Text>
                {products.map(p => (
                  <React.Fragment key={p.id}>
                    <Text>{p.title}</Text><Text>{p.unitPrice}</Text>
                    <TextField ariaLabel={`${p.title} price`} type="number" placeholder="automatic" value={b.prices[String(p.id)] === undefined ? '' : String(b.prices[String(p.id)])}
                      onChange={(_, v) => { const prices = { ...b.prices }; if (v === '' || v === undefined) { delete prices[String(p.id)]; } else { prices[String(p.id)] = Number(v); } setB({ prices }); }} />
                  </React.Fragment>
                ))}
              </div>
            )}
          </Item>
        );
      })}
    </Group>
  );
};

const CopilotGroup: React.FC = () => {
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
    <>
      <Group title="Microsoft 365 Copilot agent" description="An agent for Microsoft 365 Copilot that answers from this CRM site (deals, accounts, cases, renewals, invoices) with each person's own SharePoint permissions. Needs Microsoft 365 Copilot licences.">
        <Row label="Agent name" help="Shown in Copilot under Agents.">
          <TextField ariaLabel="Agent name" value={name} maxLength={30} onChange={(_, v) => setName(v || '')} />
        </Row>
        <Row label="Agent package" help="Download again after renaming; the same app is updated.">
          <PrimaryButton iconProps={{ iconName: 'Download' }} text={busy ? 'Preparing…' : 'Download Copilot agent'} disabled={busy || !name.trim()} onClick={download} />
        </Row>
      </Group>
      <Group title="Deploy it">
        <ol className="crm-ul">
          <li>Teams admin center › <b>Manage apps › Upload new app</b> (or Microsoft 365 admin center › Integrated apps › Upload custom apps) › select the zip.</li>
          <li>Make it available to the sales and support teams.</li>
          <li>In any record, <b>Smart summary › Ask Copilot</b> hands that record to Copilot.</li>
        </ol>
      </Group>
    </>
  );
};

const TeamsGroups: React.FC<{ host: string; tenantSiteUrl: string }> = ({ host, tenantSiteUrl }) => {
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
    <>
      <Group title="Default CRM site" description="Teams tabs run on the team's own site and the personal app on the root site. The default CRM site tells them where the CRM data lives, so everyone works in one CRM.">
        <Row label="Current default" help={isDefault ? 'This site.' : current ? 'Another site is the default.' : 'Not set yet.'}>
          <Text styles={{ root: { wordBreak: 'break-all' } }}>{current || '—'}</Text>
        </Row>
        <Row label="Make this site the default" help={host !== 'sharepoint' ? 'Open the CRM in SharePoint to change this.' : 'Needs a SharePoint admin (tenant app catalog).'}>
          <PrimaryButton text={isDefault ? 'This site is the default' : 'Use this site'} disabled={busy || isDefault || host !== 'sharepoint'} onClick={setDefault} />
        </Row>
        <Text variant="small" className="crm-muted" block styles={{ root: { padding: '8px 0' } }}>PowerShell alternative: <code>Set-PnPStorageEntity -Key &quot;{TENANT_SITE_KEY}&quot; -Value &quot;{here}&quot; -Scope Tenant</code></Text>
      </Group>
      <Group title="Add Brillienta CRM to Teams">
        <ol className="crm-ul">
          <li>SharePoint admin center › <b>More features › Apps › App catalog</b> › <i>Brillienta CRM</i> › <b>Sync to Teams</b> (or upload <code>BrillientaCRM-Teams.zip</code> in the Teams admin center).</li>
          <li>Users open <b>Apps</b> in Teams, search “Brillienta CRM” and click <b>Add</b>; it appears in the left rail.</li>
          <li>Channel tab: <b>+</b> › Brillienta CRM › leave the site empty to use the default.</li>
          <li>The CRM follows the Teams theme. If Teams blocks the print window, quotes download as HTML (open › Print › Save as PDF).</li>
        </ol>
      </Group>
    </>
  );
};

// ================================================================ settings center

interface ISection { key: string; group: string; title: string; icon: string; description: string; keywords: string; render: () => React.ReactNode }

const FLOWS: [string, string, boolean][] = [
  ['AppSource lead intake', 'Partner Center HTTPS endpoint → leads', true],
  ['Website leads & demo bookings', 'brillienta.com bookings → lead + demo; support → case; enquiries → lead', true],
  ['New lead alert', 'emails the owner; unassigned leads go to the sales manager', false],
  ['Daily sales digest', 'new leads, deals won and closing, overdue tasks (weekdays)', false],
  ['Overdue reminders', 'one email per person with their overdue tasks', false],
  ['Case SLA escalation', 'hourly; escalates newly breached cases', false],
  ['Deal won handoff', 'notifies finance and creates a customer folder', false],
  ['Quote discount approval', 'Approvals in Teams/Outlook for large discounts', false],
  ['Quote e-signature', 'customer signing page for quotes', true],
  ['Email to case', 'support mailbox → cases; replies go to the same case', false],
  ['Case satisfaction survey', 'one-click 1–5 rating page', true],
  ['Customer portal', 'customers follow and raise their requests', true],
  ['Unsubscribe', 'personal unsubscribe page for campaign emails', true],
  ['Razorpay / Stripe payment link & payment received', 'pay links on invoices, marked paid automatically', true],
  ['WhatsApp send / inbound', 'WhatsApp Business Cloud API', true],
  ['Lead / account enrichment', 'company website and Apollo.io details', true],
  ['Record security', 'item permissions by owner and team (every 15 min)', false],
  ['Inbound API', 'other systems create, update or read records', true]
];

export const SettingsPage: React.FC<{ tenantId: string; host: string; tenantSiteUrl: string }> = ({ tenantId, host, tenantSiteUrl }) => {
  const { engine, settings, updateSettings, notify, isNarrow } = useCrm();
  const [s, setS] = React.useState<ICrmSettings>(() => JSON.parse(JSON.stringify(settings)));
  const [dirty, setDirty] = React.useState(false);
  const [busy, setBusy] = React.useState(false);
  const [progress, setProgress] = React.useState<string>();
  const [section, setSection] = React.useState<string>(() => lsGet('brillienta-crm:settings-section', 'overview'));
  const [query, setQuery] = React.useState('');
  const [cf, setCf] = React.useState<ICustomField>({ entity: 'leads', key: '', label: '', type: 'text', choices: [], grid: false });
  const [flowDialog, setFlowDialog] = React.useState(false);
  const mainRef = React.useRef<HTMLDivElement>(null);
  const upd = (patch: Partial<ICrmSettings>): void => { setS({ ...s, ...patch }); setDirty(true); };

  const save = (after?: () => Promise<void>): void => {
    setBusy(true);
    updateSettings(s).then(() => (after ? after() : Promise.resolve())).then(() => { setBusy(false); setDirty(false); notify('Settings saved.', MessageBarType.success); })
      .catch(e => { setBusy(false); notify(errorText(e), MessageBarType.error); });
  };
  const discard = (): void => { setS(JSON.parse(JSON.stringify(settings))); setDirty(false); };
  const go = (key: string): void => { setSection(key); lsSet('brillienta-crm:settings-section', key); setQuery(''); mainRef.current?.scrollIntoView({ block: 'start', behavior: 'smooth' }); };

  if (!engine.can('admin')) {
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
  const month = (i: number): string => new Date(2000, i, 1).toLocaleDateString(undefined, { month: 'long' });
  const toggleRow = (label: string, help: React.ReactNode, checked: boolean, onChange: (v: boolean) => void, disabled?: boolean): React.ReactNode => (
    <Row label={label} help={help}><Toggle ariaLabel={label} checked={checked} disabled={disabled} onText="On" offText="Off" onChange={(_, c) => onChange(!!c)} /></Row>
  );

  // ---------------------------------------------------------------- checklist (overview)
  const checks: { label: string; done: boolean; key: string; optional?: boolean }[] = [
    { label: 'Company name and address (printed on quotes and invoices)', done: !!(s.company.name && s.company.address), key: 'company' },
    { label: 'Company logo', done: /^https:\/\//i.test(s.company.logoUrl), key: 'company', optional: true },
    { label: 'Licence key', done: lic.status === 'active', key: 'licence' },
    { label: 'Outlook email and Teams meetings switched on', done: s.integrations.graphEmail && s.integrations.graphMeetings, key: 'email' },
    { label: 'Pipeline stages reviewed', done: s.stages.length > 2, key: 'pipeline' },
    { label: 'Products in the price list', done: (engine.data.products || []).length > 0, key: 'pricebooks', optional: true },
    { label: 'Teams channel notifications', done: !!s.integrations.teamsWebhookUrl, key: 'teams', optional: true },
    { label: 'Quote e-signature link', done: !!s.quote.signingUrl, key: 'quotes', optional: true },
    { label: 'Customer portal and satisfaction survey links', done: !!(s.service.portalUrl && s.service.surveyUrl), key: 'portal', optional: true },
    { label: 'Unsubscribe link for email campaigns', done: !!s.marketing.unsubscribeUrl, key: 'marketing', optional: true },
    { label: 'Security groups and record-level security', done: s.security.itemLevel, key: 'security', optional: true }
  ];
  const required = checks.filter(c => !c.optional);

  // ---------------------------------------------------------------- sections
  const sections: ISection[] = [
    {
      key: 'overview', group: '', title: 'Overview', icon: 'Home', keywords: 'setup checklist start',
      description: 'What is set up, and what is left to do.',
      render: () => (
        <>
          <Group title="Setup checklist" description={`${required.filter(c => c.done).length} of ${required.length} essentials done · ${checks.filter(c => c.optional && c.done).length} of ${checks.filter(c => c.optional).length} optional`}>
            {checks.map(c => (
              <div key={c.label} className="crm-set-check">
                <Icon iconName={c.done ? 'CompletedSolid' : 'CircleRing'} className={c.done ? 'crm-set-done' : 'crm-muted'} />
                <Text style={{ flex: 1 }}>{c.label}{c.optional && <span className="crm-muted"> · optional</span>}</Text>
                <Link onClick={() => go(c.key)}>{c.done ? 'Review' : 'Set up'}</Link>
              </div>
            ))}
          </Group>
          <Group title="All settings">
            <div className="crm-set-tiles">
              {sections.filter(x => x.key !== 'overview').map(x => (
                <button key={x.key} type="button" className="crm-set-tile" onClick={() => go(x.key)}>
                  <Icon iconName={x.icon} /><span><b>{x.title}</b><span className="crm-muted">{x.description}</span></span>
                </button>
              ))}
            </div>
          </Group>
        </>
      )
    },
    // ============================================================ General
    {
      key: 'company', group: 'General', title: 'Company profile', icon: 'CityNext', keywords: 'name address email phone website logo tax vat gst branding',
      description: 'Your company details, printed on quotes, invoices and emails.',
      render: () => (
        <Group title="Company">
          <Row label="Company name"><TextField ariaLabel="Company name" value={s.company.name} onChange={(_, v) => upd({ company: { ...s.company, name: v || '' } })} /></Row>
          <Row label="Address" help="As it should appear on documents."><TextField ariaLabel="Address" multiline rows={3} value={s.company.address} onChange={(_, v) => upd({ company: { ...s.company, address: v || '' } })} /></Row>
          <Row label="Email"><TextField ariaLabel="Email" value={s.company.email} onChange={(_, v) => upd({ company: { ...s.company, email: v || '' } })} /></Row>
          <Row label="Phone"><TextField ariaLabel="Phone" value={s.company.phone} onChange={(_, v) => upd({ company: { ...s.company, phone: v || '' } })} /></Row>
          <Row label="Website"><TextField ariaLabel="Website" value={s.company.website} onChange={(_, v) => upd({ company: { ...s.company, website: v || '' } })} /></Row>
          <Row label="Tax / VAT / GST ID"><TextField ariaLabel="Tax ID" value={s.company.taxId} onChange={(_, v) => upd({ company: { ...s.company, taxId: v || '' } })} /></Row>
          <Row label="Logo" help="Upload the logo to a SharePoint library and paste its https link.">
            <Stack tokens={{ childrenGap: 6 }}>
              <TextField ariaLabel="Logo URL" placeholder="https://…" value={s.company.logoUrl} errorMessage={urlError(s.company.logoUrl)} onChange={(_, v) => upd({ company: { ...s.company, logoUrl: v || '' } })} />
              {/^https:\/\//i.test(s.company.logoUrl) && <img src={s.company.logoUrl} alt="Logo preview" style={{ maxHeight: 48, maxWidth: 200, objectFit: 'contain' }} />}
            </Stack>
          </Row>
        </Group>
      )
    },
    {
      key: 'regional', group: 'General', title: 'Currency & regional', icon: 'Globe', keywords: 'currency exchange rate fiscal year multi-currency',
      description: 'Company currency, fiscal year and other currencies.',
      render: () => (
        <>
          <Group title="Regional">
            <Row label="Company currency" help="ISO code, e.g. INR, USD, EUR. Reports and dashboards use it.">
              <TextField ariaLabel="Company currency" value={s.currency} styles={{ root: { width: 120 } }} onChange={(_, v) => upd({ currency: (v || '').toUpperCase().slice(0, 3) })} />
            </Row>
            <Row label="Fiscal year starts in" help="Used for quarters in the forecast and dashboard.">
              <Dropdown ariaLabel="Fiscal year start" selectedKey={s.fiscalYearStartMonth} options={Array.from({ length: 12 }, (_, i) => ({ key: i + 1, text: month(i) }))}
                onChange={(_, o) => upd({ fiscalYearStartMonth: Number(o?.key) })} />
            </Row>
          </Group>
          <CurrenciesGroup s={s} upd={upd} />
        </>
      )
    },
    {
      key: 'licence', group: 'General', title: 'Licence', icon: 'Certificate', keywords: 'licence license key plan trial enterprise standard tenant',
      description: 'Your plan and licence key.',
      render: () => (
        <>
          <Group title="Licence">
            <MessageBar messageBarType={lic.status === 'active' ? MessageBarType.success : lic.status === 'trial' ? MessageBarType.info : MessageBarType.severeWarning} isMultiline styles={{ root: { margin: '10px 0' } }}>{lic.message}</MessageBar>
            <Row label="Microsoft 365 tenant ID" help="Send this to Brillienta to receive a licence key."><TextField ariaLabel="Tenant ID" readOnly value={tenantId} /></Row>
            <Row label="Licence key" help="Paste the key and save the settings to activate it."><TextField ariaLabel="Licence key" multiline rows={4} value={s.licenseKey} onChange={(_, v) => upd({ licenseKey: (v || '').trim() })} /></Row>
          </Group>
          <Group title="Plans">
            <ul className="crm-ul">
              <li><b>Standard</b> – leads, accounts, contacts, pipeline, quotes and products, demos, activities, dashboards, reports, Outlook email and Teams meetings, import/export.</li>
              <li><b>Enterprise</b> – everything in Standard plus cases and SLAs, knowledge base, campaigns, forecasting and quotas, webhooks, custom fields, audit log and security groups.</li>
              <li><b>Trial</b> – all Enterprise features for 30 days.</li>
            </ul>
          </Group>
        </>
      )
    },
    // ============================================================ Sales
    {
      key: 'pipeline', group: 'Sales', title: 'Pipeline stages', icon: 'Flow', keywords: 'stages probability forecast category opportunity deal',
      description: 'The stages a deal moves through, with probability and forecast category.',
      render: () => (
        <Group title="Stages" description="Order matters. “Closed Won” and “Closed Lost” must stay. Probability and forecast category are set automatically when a deal changes stage."
          action={<DefaultButton iconProps={{ iconName: 'Add' }} text="Add stage" onClick={() => upd({ stages: [...s.stages.slice(0, -2), { name: 'New stage', probability: 50, forecast: 'Pipeline' }, ...s.stages.slice(-2)] })} />}>
          <div className="crm-stage-table" role="table" aria-label="Pipeline stages" style={{ padding: '10px 0' }}>
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
        </Group>
      )
    },
    {
      key: 'picklists', group: 'Sales', title: 'Pick lists', icon: 'BulletedList', keywords: 'lead sources industries lost reasons case categories choices dropdown',
      description: 'Choices offered in drop-downs. One value per line.',
      render: () => (
        <Group title="Pick lists" description="One value per line. Removing a value does not change existing records.">
          <Row label="Lead sources"><TextField ariaLabel="Lead sources" multiline autoAdjustHeight rows={5} value={lines(s.leadSources)} onChange={(_, v) => upd({ leadSources: parseLines(v || '') })} /></Row>
          <Row label="Industries"><TextField ariaLabel="Industries" multiline autoAdjustHeight rows={5} value={lines(s.industries)} onChange={(_, v) => upd({ industries: parseLines(v || '') })} /></Row>
          <Row label="Lost reasons" help="Asked when a deal is marked lost."><TextField ariaLabel="Lost reasons" multiline autoAdjustHeight rows={4} value={lines(s.lostReasons)} onChange={(_, v) => upd({ lostReasons: parseLines(v || '') })} /></Row>
          <Row label="Case categories"><TextField ariaLabel="Case categories" multiline autoAdjustHeight rows={4} value={lines(s.caseCategories)} onChange={(_, v) => upd({ caseCategories: parseLines(v || '') })} /></Row>
        </Group>
      )
    },
    {
      key: 'leads', group: 'Sales', title: 'Leads: routing, scoring & enrichment', icon: 'People', keywords: 'round robin assignment scoring points enrichment apollo',
      description: 'How new leads are assigned, scored and completed.',
      render: () => (
        <>
          <Group title="Assignment" description="Leads from a team's countries go to that team first (Teams & territories). The rest use this round-robin.">
            {toggleRow('Assign new leads automatically', 'Round-robin across the reps below.', s.assignment.autoAssignNewLeads, v => upd({ assignment: { ...s.assignment, autoAssignNewLeads: v } }))}
            <Row label="Round-robin reps" help="In this order.">
              <Stack tokens={{ childrenGap: 4 }}>
                {s.assignment.roundRobin.map((u, i) => (
                  <Stack key={u.id} horizontal verticalAlign="center"><Text style={{ flex: 1 }}>{i + 1}. {u.title}</Text>
                    <IconButton iconProps={{ iconName: 'Cancel' }} ariaLabel={`Remove ${u.title}`} onClick={() => upd({ assignment: { ...s.assignment, roundRobin: s.assignment.roundRobin.filter(x => x.id !== u.id) } })} /></Stack>
                ))}
                <UserPicker showMe={false} onChange={u => { if (u && !s.assignment.roundRobin.some(x => x.id === u.id)) { upd({ assignment: { ...s.assignment, roundRobin: [...s.assignment.roundRobin, u] } }); } }} />
              </Stack>
            </Row>
          </Group>
          <Group title="Lead scoring (0–100)" description="Points per line, written as “value = points”.">
            <Row label="Rating"><TextField ariaLabel="Rating points" multiline rows={3} value={kv(s.scoring.rating)} onChange={(_, v) => upd({ scoring: { ...s.scoring, rating: parseKv(v || '') } })} /></Row>
            <Row label="Lead source"><TextField ariaLabel="Source points" multiline rows={5} value={kv(s.scoring.sources)} onChange={(_, v) => upd({ scoring: { ...s.scoring, sources: parseKv(v || '') } })} /></Row>
            <Row label="Job title keywords" help="The best match counts.">
              <TextField ariaLabel="Job title keywords" multiline rows={5} value={s.scoring.titleKeywords.map(k => `${k.keyword} = ${k.points}`).join('\n')}
                onChange={(_, v) => { const m = parseKv(v || ''); upd({ scoring: { ...s.scoring, titleKeywords: Object.keys(m).map(k => ({ keyword: k, points: m[k] })) } }); }} />
            </Row>
            <Row label="Other signals">
              <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
                <TextField label="Has phone" type="number" value={String(s.scoring.hasPhone)} onChange={(_, v) => upd({ scoring: { ...s.scoring, hasPhone: Number(v) || 0 } })} styles={{ root: { width: 90 } }} />
                <TextField label="Has company" type="number" value={String(s.scoring.hasCompany)} onChange={(_, v) => upd({ scoring: { ...s.scoring, hasCompany: Number(v) || 0 } })} styles={{ root: { width: 90 } }} />
                <TextField label="Per activity" type="number" value={String(s.scoring.perActivity)} onChange={(_, v) => upd({ scoring: { ...s.scoring, perActivity: Number(v) || 0 } })} styles={{ root: { width: 90 } }} />
                <TextField label="Activity max" type="number" value={String(s.scoring.maxActivityPoints)} onChange={(_, v) => upd({ scoring: { ...s.scoring, maxActivityPoints: Number(v) || 0 } })} styles={{ root: { width: 90 } }} />
              </Stack>
            </Row>
          </Group>
          <Group title="Enrichment">
            {toggleRow('Enrich new leads automatically', 'Leads with a business email get their website and company at once; the “Lead enrichment” flow adds company details.', s.enrichment.autoNewLeads, v => upd({ enrichment: { autoNewLeads: v } }))}
          </Group>
        </>
      )
    },
    {
      key: 'teams-territories', group: 'Sales', title: 'Teams, territories & quotas', icon: 'Group', keywords: 'team manager members territory countries quota forecast',
      description: 'Sales teams, the countries they cover, and quarterly quotas.',
      render: () => (
        <>
          <Group title="Teams & territories" description="New leads from a team's countries are shared round-robin within that team. Dashboards and the forecast can be filtered by team."
            action={<DefaultButton iconProps={{ iconName: 'Add' }} text="New team" onClick={() => upd({ teams: [...s.teams, { id: `team${Date.now()}`, name: 'New team', members: [], territories: [], lastIndex: -1 }] })} />}>
            {s.teams.length === 0 && <Text className="crm-muted" block styles={{ root: { padding: '10px 0' } }}>No teams yet.</Text>}
            {s.teams.map((t, i) => {
              const setT = (patch: Partial<typeof t>): void => upd({ teams: s.teams.map((x, j) => (j === i ? { ...x, ...patch } : x)) });
              return (
                <Item key={t.id}>
                  <Stack horizontal wrap verticalAlign="end" tokens={{ childrenGap: 10 }}>
                    <TextField label="Team" value={t.name} onChange={(_, v) => setT({ name: v || '' })} styles={{ root: { width: 200 } }} />
                    <TextField label="Territories (countries, comma separated)" value={t.territories.join(', ')} placeholder="India, UAE, Saudi Arabia" styles={{ root: { flex: 1, minWidth: 240 } }}
                      onChange={(_, v) => setT({ territories: (v || '').split(',').map(x => x.trim()).filter(Boolean) })} />
                    <IconButton iconProps={{ iconName: 'Delete' }} ariaLabel="Remove team" onClick={() => upd({ teams: s.teams.filter((_, j) => j !== i) })} />
                  </Stack>
                  <Stack horizontal wrap tokens={{ childrenGap: 24 }} style={{ marginTop: 8 }}>
                    <Stack tokens={{ childrenGap: 4 }} styles={{ root: { minWidth: 220 } }}>
                      <Label>Manager</Label>
                      {t.manager ? <Stack horizontal verticalAlign="center"><Text style={{ flex: 1 }}>{t.manager.title}</Text><IconButton iconProps={{ iconName: 'Cancel' }} ariaLabel="Remove manager" onClick={() => setT({ manager: undefined })} /></Stack>
                        : <UserPicker showMe={false} onChange={u => { if (u) { setT({ manager: u }); } }} />}
                    </Stack>
                    <Stack tokens={{ childrenGap: 4 }} styles={{ root: { minWidth: 260 } }}>
                      <Label>Members</Label>
                      {t.members.map(m => (
                        <Stack key={m.id} horizontal verticalAlign="center"><Text style={{ flex: 1 }}>{m.title}</Text>
                          <IconButton iconProps={{ iconName: 'Cancel' }} ariaLabel={`Remove ${m.title}`} onClick={() => setT({ members: t.members.filter(x => x.id !== m.id) })} /></Stack>
                      ))}
                      <UserPicker showMe={false} onChange={u => { if (u && !t.members.some(x => x.id === u.id)) { setT({ members: [...t.members, u] }); } }} />
                    </Stack>
                  </Stack>
                </Item>
              );
            })}
          </Group>
          <Group title="Quarterly quotas" description={`Target per rep per quarter, in ${s.currency}. Shown in the forecast.`}>
            {s.quotas.map((qt, i) => (
              <Row key={qt.user.id} label={qt.user.title}>
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                  <Unit after={s.currency}><TextField type="number" ariaLabel={`Quota for ${qt.user.title}`} value={String(qt.quarterly)} onChange={(_, v) => upd({ quotas: s.quotas.map((x, j) => (j === i ? { ...x, quarterly: Number(v) || 0 } : x)) })} styles={{ root: { width: 200 } }} /></Unit>
                  <IconButton iconProps={{ iconName: 'Delete' }} ariaLabel="Remove quota" onClick={() => upd({ quotas: s.quotas.filter((_, j) => j !== i) })} />
                </Stack>
              </Row>
            ))}
            <Row label="Add a rep"><UserPicker showMe={false} onChange={u => { if (u && !s.quotas.some(x => x.user.id === u.id)) { upd({ quotas: [...s.quotas, { user: u, quarterly: 0 }] }); } }} /></Row>
          </Group>
        </>
      )
    },
    {
      key: 'quotes', group: 'Sales', title: 'Quotes & e-signature', icon: 'Document', keywords: 'quote numbering validity terms footer discount approval signature sign',
      description: 'Quote numbering, terms, discount approval and online signing.',
      render: () => (
        <>
          <Group title="Quotes">
            <Row label="Number prefix" help={`Quotes are numbered ${s.quote.prefix || 'Q'}-${new Date().getFullYear()}-0001.`}><TextField ariaLabel="Quote prefix" value={s.quote.prefix} styles={{ root: { width: 120 } }} onChange={(_, v) => upd({ quote: { ...s.quote, prefix: v || 'Q' } })} /></Row>
            <Row label="Valid for"><Unit after="days"><TextField ariaLabel="Validity days" type="number" value={String(s.quote.validityDays)} styles={{ root: { width: 160 } }} onChange={(_, v) => upd({ quote: { ...s.quote, validityDays: Number(v) || 30 } })} /></Unit></Row>
            <Row label="Default terms & conditions"><TextField ariaLabel="Terms" multiline rows={4} value={s.quote.terms} onChange={(_, v) => upd({ quote: { ...s.quote, terms: v || '' } })} /></Row>
            <Row label="Footer"><TextField ariaLabel="Quote footer" value={s.quote.footer} onChange={(_, v) => upd({ quote: { ...s.quote, footer: v || '' } })} /></Row>
          </Group>
          <Group title="Discount approval">
            <Row label="Approval needed above" help="Quotes with a higher total discount need manager approval (flow “Quote discount approval”). 0 = off.">
              <Unit after="%"><TextField ariaLabel="Approval threshold" type="number" value={String(s.quote.approvalDiscountPct)} styles={{ root: { width: 140 } }} onChange={(_, v) => upd({ quote: { ...s.quote, approvalDiscountPct: Math.max(0, Number(v) || 0) } })} /></Unit>
            </Row>
          </Group>
          <Group title="E-signature" description="Customers review and sign quotes online through the flow “Quote e-signature”.">
            <Row label="Signing link" help="HTTP URL of the flow. Leave empty to turn e-signature off.">
              <TextField ariaLabel="Signing link" placeholder="https://…" value={s.quote.signingUrl} errorMessage={urlError(s.quote.signingUrl)} onChange={(_, v) => upd({ quote: { ...s.quote, signingUrl: (v || '').trim() } })} />
            </Row>
            {toggleRow('Signed quote wins the deal', 'Marks the linked opportunity as Closed Won.', s.quote.signatureClosesOpportunity, v => upd({ quote: { ...s.quote, signatureClosesOpportunity: v } }))}
          </Group>
        </>
      )
    },
    {
      key: 'pricebooks', group: 'Sales', title: 'Price books', icon: 'Money', keywords: 'price book discount partner currency products prices',
      description: 'Price lists per currency or customer segment.',
      render: () => <PriceBooksGroup s={s} upd={upd} />
    },
    {
      key: 'sequences', group: 'Sales', title: 'Sequences', icon: 'Send', keywords: 'cadence sequence steps follow up automation',
      description: 'Timed follow-up steps for leads and contacts.',
      render: () => (
        <Group title="Sequences" description="Enrol leads or contacts from their grid or record. Due steps run when the owner opens the CRM. A reply by email or WhatsApp stops the sequence."
          action={<DefaultButton iconProps={{ iconName: 'Add' }} text="New sequence" onClick={() => upd({ sequences: [...s.sequences, { id: `seq${Date.now()}`, name: 'New sequence', autoSend: false, steps: [{ day: 0, type: 'Email', templateId: s.emailTemplates[0]?.id, title: 'First email' }] }] })} />}>
          {s.sequences.map((q, qi) => {
            const setQ = (patch: Partial<typeof q>): void => upd({ sequences: s.sequences.map((x, j) => (j === qi ? { ...x, ...patch } : x)) });
            return (
              <Item key={q.id}>
                <Stack horizontal wrap verticalAlign="end" tokens={{ childrenGap: 12 }}>
                  <TextField label="Sequence" value={q.name} onChange={(_, v) => setQ({ name: v || '' })} styles={{ root: { width: 260 } }} />
                  <Toggle label="Send email steps automatically" inlineLabel checked={q.autoSend} onChange={(_, c) => setQ({ autoSend: !!c })} />
                  <IconButton iconProps={{ iconName: 'Delete' }} ariaLabel="Remove sequence" onClick={() => upd({ sequences: s.sequences.filter((_, j) => j !== qi) })} />
                </Stack>
                <div className="crm-seq-steps">
                  {q.steps.map((st, si) => {
                    const setStep = (patch: Partial<typeof st>): void => setQ({ steps: q.steps.map((x, j) => (j === si ? { ...x, ...patch } : x)) });
                    const tplOptions = st.type === 'Email' ? s.emailTemplates.map(t => ({ key: t.id, text: t.name })) : st.type === 'WhatsApp' ? s.whatsapp.templates.map(t => ({ key: t.id, text: t.name })) : [];
                    return (
                      <Stack key={si} horizontal wrap verticalAlign="end" tokens={{ childrenGap: 8 }}>
                        <TextField label={si ? undefined : 'Day'} ariaLabel="Day" type="number" value={String(st.day)} onChange={(_, v) => setStep({ day: Math.max(0, Number(v) || 0) })} styles={{ root: { width: 70 } }} />
                        <Dropdown label={si ? undefined : 'Step'} ariaLabel="Step type" selectedKey={st.type} options={['Email', 'Call', 'WhatsApp', 'Task'].map(x => ({ key: x, text: x }))}
                          onChange={(_, o) => setStep({ type: o?.key as typeof st.type, templateId: undefined })} styles={{ root: { width: 120 } }} />
                        <TextField label={si ? undefined : 'Title'} ariaLabel="Step title" value={st.title} onChange={(_, v) => setStep({ title: v || '' })} styles={{ root: { width: 240 } }} />
                        {tplOptions.length > 0 && <Dropdown label={si ? undefined : 'Template'} ariaLabel="Template" selectedKey={st.templateId || null} options={tplOptions} onChange={(_, o) => setStep({ templateId: String(o?.key) })} styles={{ root: { width: 180 } }} />}
                        <IconButton iconProps={{ iconName: 'Cancel' }} ariaLabel="Remove step" onClick={() => setQ({ steps: q.steps.filter((_, j) => j !== si) })} />
                      </Stack>
                    );
                  })}
                </div>
                <DefaultButton iconProps={{ iconName: 'Add' }} text="Add step" styles={{ root: { marginTop: 8 } }}
                  onClick={() => setQ({ steps: [...q.steps, { day: (q.steps[q.steps.length - 1]?.day || 0) + 3, type: 'Task', title: 'Follow up' }] })} />
              </Item>
            );
          })}
        </Group>
      )
    },
    // ============================================================ Billing
    {
      key: 'invoices', group: 'Billing', title: 'Invoices & renewals', icon: 'Money', keywords: 'invoice numbering payment terms bank details renewal subscription',
      description: 'Invoice numbering, payment terms and automatic renewals.',
      render: () => (
        <>
          <Group title="Invoices">
            <Row label="Number prefix" help={`Invoices are numbered ${s.invoice.prefix || 'INV'}-${new Date().getFullYear()}-0001.`}><TextField ariaLabel="Invoice prefix" value={s.invoice.prefix} styles={{ root: { width: 120 } }} onChange={(_, v) => upd({ invoice: { ...s.invoice, prefix: v || 'INV' } })} /></Row>
            <Row label="Payment terms"><Unit after="days"><TextField ariaLabel="Payment terms" type="number" value={String(s.invoice.paymentTermsDays)} styles={{ root: { width: 160 } }} onChange={(_, v) => upd({ invoice: { ...s.invoice, paymentTermsDays: Math.max(0, Number(v) || 0) } })} /></Unit></Row>
            <Row label="Bank transfer details" help="Printed on invoices."><TextField ariaLabel="Bank details" multiline rows={3} value={s.invoice.bankDetails} onChange={(_, v) => upd({ invoice: { ...s.invoice, bankDetails: v || '' } })} /></Row>
            <Row label="Footer"><TextField ariaLabel="Invoice footer" value={s.invoice.footer} onChange={(_, v) => upd({ invoice: { ...s.invoice, footer: v || '' } })} /></Row>
          </Group>
          <Group title="Subscriptions & renewals">
            {toggleRow('Create subscriptions when a quote is accepted', 'Monthly and annual quote lines become subscriptions.', s.invoice.createSubscriptionsOnAccept, v => upd({ invoice: { ...s.invoice, createSubscriptionsOnAccept: v } }))}
            {toggleRow('Create renewal opportunities', 'For annual subscriptions, ahead of the renewal date.', s.renewals.autoCreateOpportunities, v => upd({ renewals: { ...s.renewals, autoCreateOpportunities: v } }))}
            <Row label="Days before renewal"><Unit after="days"><TextField ariaLabel="Days before renewal" type="number" value={String(s.renewals.leadDays)} styles={{ root: { width: 160 } }} onChange={(_, v) => upd({ renewals: { ...s.renewals, leadDays: Math.max(1, Number(v) || 60) } })} /></Unit></Row>
          </Group>
        </>
      )
    },
    {
      key: 'payments', group: 'Billing', title: 'Online payments', icon: 'PaymentCard', keywords: 'razorpay stripe payment link pay online',
      description: 'Razorpay or Stripe payment links on invoices.',
      render: () => (
        <Group title="Payment links" description="Invoices get a “Create payment link” button; payments mark the invoice as paid. API keys stay in Power Automate, never in the browser.">
          <Row label="Provider" help="Import and turn on the matching “payment link” and “payment received” flows.">
            <Dropdown ariaLabel="Payment provider" selectedKey={s.payments.provider} options={[{ key: 'none', text: 'Off' }, { key: 'Razorpay', text: 'Razorpay' }, { key: 'Stripe', text: 'Stripe' }]}
              onChange={(_, o) => upd({ payments: { provider: (o?.key as 'none' | 'Razorpay' | 'Stripe') || 'none' } })} />
          </Row>
        </Group>
      )
    },
    {
      key: 'accounting', group: 'Billing', title: 'Accounting export', icon: 'Financial', keywords: 'tally xero zoho quickbooks ledger export',
      description: 'Settings for Tally and Xero exports (Invoices › ⋯ › Export for accounting).',
      render: () => (
        <Group title="Accounting export" description="Zoho Books, QuickBooks Online and CSV need no settings.">
          <Row label="Tally sales ledger"><TextField ariaLabel="Tally sales ledger" value={s.accounting.tallySalesLedger} onChange={(_, v) => upd({ accounting: { ...s.accounting, tallySalesLedger: v || '' } })} /></Row>
          <Row label="Tally tax ledger"><TextField ariaLabel="Tally tax ledger" value={s.accounting.tallyTaxLedger} onChange={(_, v) => upd({ accounting: { ...s.accounting, tallyTaxLedger: v || '' } })} /></Row>
          <Row label="Xero sales account code"><TextField ariaLabel="Xero account" value={s.accounting.xeroSalesAccount} onChange={(_, v) => upd({ accounting: { ...s.accounting, xeroSalesAccount: v || '' } })} /></Row>
          <Row label="Xero tax type"><TextField ariaLabel="Xero tax type" value={s.accounting.xeroTaxType} onChange={(_, v) => upd({ accounting: { ...s.accounting, xeroTaxType: v || '' } })} /></Row>
        </Group>
      )
    },
    // ============================================================ Service
    {
      key: 'cases', group: 'Service', title: 'Cases & SLAs', icon: 'Headset', keywords: 'case sla response resolution business hours holidays pause',
      description: 'Case numbering, response and resolution targets, business hours.',
      render: () => (
        <>
          <Group title="Cases">
            <Row label="Number prefix" help={`New cases are numbered ${s.casePrefix || 'CS'}-${new Date().getFullYear()}-<ID>.`}><TextField ariaLabel="Case prefix" value={s.casePrefix} styles={{ root: { width: 120 } }} onChange={(_, v) => upd({ casePrefix: v || 'CS' })} /></Row>
          </Group>
          <Group title="Service level agreements" description="Hours to the first response and to resolution, per priority.">
            <div className="crm-sla-table" style={{ padding: '10px 0' }}>
              <Text variant="small" className="crm-muted">Priority</Text><Text variant="small" className="crm-muted">First response (h)</Text><Text variant="small" className="crm-muted">Resolve (h)</Text>
              {['Critical', 'High', 'Normal', 'Low'].map(p => (
                <React.Fragment key={p}>
                  <Text>{p}</Text>
                  <TextField ariaLabel={`${p} first response hours`} type="number" value={String(s.firstResponseHours[p] ?? '')} onChange={(_, v) => upd({ firstResponseHours: { ...s.firstResponseHours, [p]: Math.max(0, Number(v) || 0) } })} />
                  <TextField ariaLabel={`${p} resolution hours`} type="number" value={String(s.slaHours[p] ?? '')} onChange={(_, v) => upd({ slaHours: { ...s.slaHours, [p]: Math.max(0, Number(v) || 0) } })} />
                </React.Fragment>
              ))}
            </div>
            {toggleRow('Pause while waiting on the customer', 'The SLA clock stops while a case is “Waiting on customer”.', s.slaPauseWhenWaiting, v => upd({ slaPauseWhenWaiting: v }))}
          </Group>
          <Group title="Business hours">
            {toggleRow('Count business hours only', 'Otherwise SLAs run around the clock.', s.businessHours.enabled, v => upd({ businessHours: { ...s.businessHours, enabled: v } }))}
            {s.businessHours.enabled && (
              <>
                <Row label="Time zone" help="IANA name, e.g. Asia/Kolkata, Europe/London."><TextField ariaLabel="Time zone" value={s.businessHours.timeZone} onChange={(_, v) => upd({ businessHours: { ...s.businessHours, timeZone: (v || '').trim() } })} /></Row>
                <Row label="Opening hours">
                  <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                    <TextField ariaLabel="Opens" value={s.businessHours.start} placeholder="09:00" styles={{ root: { width: 90 } }} onChange={(_, v) => upd({ businessHours: { ...s.businessHours, start: v || '' } })} />
                    <Text>to</Text>
                    <TextField ariaLabel="Closes" value={s.businessHours.end} placeholder="18:00" styles={{ root: { width: 90 } }} onChange={(_, v) => upd({ businessHours: { ...s.businessHours, end: v || '' } })} />
                  </Stack>
                </Row>
                <Row label="Working days">
                  <Stack horizontal wrap tokens={{ childrenGap: 10 }}>
                    {['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'].map((d, i) => (
                      <Checkbox key={d} label={d} checked={s.businessHours.days.indexOf(i) >= 0}
                        onChange={(_, c) => upd({ businessHours: { ...s.businessHours, days: c ? [...s.businessHours.days, i].sort() : s.businessHours.days.filter(x => x !== i) } })} />
                    ))}
                  </Stack>
                </Row>
                <Row label="Holidays" help="YYYY-MM-DD, one per line."><TextField ariaLabel="Holidays" multiline rows={3} value={s.businessHours.holidays.join('\n')}
                  onChange={(_, v) => upd({ businessHours: { ...s.businessHours, holidays: (v || '').split(/\s+/).filter(x => /^\d{4}-\d{2}-\d{2}$/.test(x)) } })} /></Row>
              </>
            )}
          </Group>
        </>
      )
    },
    {
      key: 'portal', group: 'Service', title: 'Customer portal & surveys', icon: 'Emoji2', keywords: 'portal csat satisfaction survey customer self service',
      description: 'Self-service portal and satisfaction surveys (Power Automate flows).',
      render: () => (
        <>
          <Group title="Customer portal" description="Customers sign in with an email link, follow their requests, reply and open new ones.">
            <Row label="Portal link" help="HTTP URL of the flow “Customer portal”. Contacts then get “Invite to portal”.">
              <TextField ariaLabel="Portal link" placeholder="https://…" value={s.service.portalUrl} errorMessage={urlError(s.service.portalUrl)} onChange={(_, v) => upd({ service: { ...s.service, portalUrl: (v || '').trim() } })} />
            </Row>
          </Group>
          <Group title="Satisfaction surveys">
            <Row label="Survey link" help="HTTP URL of the flow “Case satisfaction survey”.">
              <TextField ariaLabel="Survey link" placeholder="https://…" value={s.service.surveyUrl} errorMessage={urlError(s.service.surveyUrl)} onChange={(_, v) => upd({ service: { ...s.service, surveyUrl: (v || '').trim() } })} />
            </Row>
            {toggleRow('Send a survey when a case is resolved', 'A one-click 1–5 rating email to the customer.', s.service.surveyOnResolve, v => upd({ service: { ...s.service, surveyOnResolve: v } }))}
          </Group>
        </>
      )
    },
    // ============================================================ Communication
    {
      key: 'email', group: 'Communication', title: 'Outlook & email', icon: 'Mail', keywords: 'outlook graph mail send teams meetings sync logging calendar',
      description: 'Send email, create Teams meetings and log Outlook conversations.',
      render: () => (
        <>
          <Group title="Microsoft 365">
            {toggleRow('Send email from Outlook', 'From the user’s own mailbox (Microsoft Graph Mail.Send).', s.integrations.graphEmail, v => upd({ integrations: { ...s.integrations, graphEmail: v } }))}
            {toggleRow('Create Teams meetings for demos', 'With Outlook invitations (Calendars.ReadWrite).', s.integrations.graphMeetings, v => upd({ integrations: { ...s.integrations, graphMeetings: v } }))}
            {toggleRow('Allow Outlook sync', 'Each user turns it on in My work. Only email and meetings with known contacts and leads are logged (Mail.Read).', s.integrations.outlookSync, v => upd({ integrations: { ...s.integrations, outlookSync: v } }))}
            {s.integrations.outlookSync && (
              <>
                <Row label="First sync reads the last"><Unit after="days"><TextField ariaLabel="Lookback days" type="number" value={String(s.integrations.syncLookbackDays)} styles={{ root: { width: 160 } }} onChange={(_, v) => upd({ integrations: { ...s.integrations, syncLookbackDays: Math.max(1, Math.min(365, Number(v) || 30)) } })} /></Unit></Row>
                <Row label="Never log these domains" help="One per line. Your own domain is always excluded."><TextField ariaLabel="Excluded domains" multiline autoAdjustHeight rows={2} value={lines(s.integrations.internalDomains)} onChange={(_, v) => upd({ integrations: { ...s.integrations, internalDomains: parseLines(v || '') } })} /></Row>
              </>
            )}
            <MessageBar isMultiline styles={{ root: { margin: '10px 0' } }}>One-time admin step: <b>SharePoint admin center › Advanced › API access</b> › approve Microsoft Graph Mail.Send, Mail.Read, Calendars.ReadWrite and User.ReadBasic.All.</MessageBar>
          </Group>
          <Group title="Email templates" description={'Placeholders: {{firstName}} {{lastName}} {{fullName}} {{company}} {{product}} {{ownerName}} {{senderName}} {{scheduledDate}} {{meetingLink}} {{quoteNumber}} {{total}} {{companyName}}'}
            action={<DefaultButton iconProps={{ iconName: 'Add' }} text="Add template" onClick={() => upd({ emailTemplates: [...s.emailTemplates, { id: uid(), name: 'New template', subject: '', body: '' }] })} />}>
            {s.emailTemplates.map((t, i) => (
              <Item key={t.id}>
                <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="end">
                  <TextField label="Name" value={t.name} onChange={(_, v) => upd({ emailTemplates: s.emailTemplates.map((x, j) => (j === i ? { ...x, name: v || '' } : x)) })} styles={{ root: { width: 200 } }} />
                  <TextField label="Subject" value={t.subject} onChange={(_, v) => upd({ emailTemplates: s.emailTemplates.map((x, j) => (j === i ? { ...x, subject: v || '' } : x)) })} styles={{ root: { flex: 1, minWidth: 220 } }} />
                  <IconButton iconProps={{ iconName: 'Delete' }} ariaLabel="Delete template" onClick={() => upd({ emailTemplates: s.emailTemplates.filter((_, j) => j !== i) })} />
                </Stack>
                <TextField label="Body" multiline rows={5} value={t.body} onChange={(_, v) => upd({ emailTemplates: s.emailTemplates.map((x, j) => (j === i ? { ...x, body: v || '' } : x)) })} />
              </Item>
            ))}
          </Group>
        </>
      )
    },
    {
      key: 'whatsapp', group: 'Communication', title: 'WhatsApp', icon: 'Chat', keywords: 'whatsapp business cloud api templates country code',
      description: 'Click-to-chat or the WhatsApp Business Cloud API, and message templates.',
      render: () => (
        <>
          <Group title="WhatsApp">
            {toggleRow('WhatsApp on records', 'Shows the WhatsApp command on leads, contacts, deals, cases and invoices.', s.whatsapp.enabled, v => upd({ whatsapp: { ...s.whatsapp, enabled: v } }))}
            <Row label="How messages are sent" help="Cloud API needs the flows “WhatsApp send” and “WhatsApp inbound”; replies and new leads then come in automatically.">
              <Dropdown ariaLabel="WhatsApp mode" selectedKey={s.whatsapp.mode} options={[{ key: 'click', text: 'Click to chat (WhatsApp app)' }, { key: 'cloud', text: 'WhatsApp Business Cloud API' }]}
                onChange={(_, o) => upd({ whatsapp: { ...s.whatsapp, mode: (o?.key as 'click' | 'cloud') || 'click' } })} />
            </Row>
            <Row label="Default country code" help="Added to local numbers."><Unit before="+"><TextField ariaLabel="Default country code" value={s.whatsapp.defaultCountryCode} styles={{ root: { width: 120 } }} onChange={(_, v) => upd({ whatsapp: { ...s.whatsapp, defaultCountryCode: (v || '').replace(/\D/g, '').slice(0, 4) } })} /></Unit></Row>
          </Group>
          <Group title="Templates" description={'Placeholders: {{firstName}} {{company}} {{senderName}} {{companyName}} {{quoteNumber}} {{invoiceNumber}} {{total}} {{paymentLink}} {{scheduledDate}} {{meetingLink}} {{caseNumber}}'}
            action={<DefaultButton iconProps={{ iconName: 'Add' }} text="Add template" onClick={() => upd({ whatsapp: { ...s.whatsapp, templates: [...s.whatsapp.templates, { id: `t${Date.now()}`, name: 'New template', body: 'Hi {{firstName}}, ', language: 'en', params: ['firstName'] }] } })} />}>
            {s.whatsapp.templates.map((t, i) => {
              const setT = (patch: Partial<typeof t>): void => upd({ whatsapp: { ...s.whatsapp, templates: s.whatsapp.templates.map((x, j) => (j === i ? { ...x, ...patch } : x)) } });
              return (
                <Item key={t.id}>
                  <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="end">
                    <TextField label="Template" value={t.name} onChange={(_, v) => setT({ name: v || '' })} styles={{ root: { width: 180 } }} />
                    <TextField label="Approved name (Cloud API)" value={t.cloudName || ''} onChange={(_, v) => setT({ cloudName: (v || '').trim() })} styles={{ root: { width: 200 } }} />
                    <TextField label="Language" value={t.language || ''} placeholder="en" onChange={(_, v) => setT({ language: (v || '').trim() })} styles={{ root: { width: 80 } }} />
                    <TextField label="{{1}}, {{2}}… values" value={(t.params || []).join(', ')} placeholder="firstName, quoteNumber" onChange={(_, v) => setT({ params: (v || '').split(',').map(x => x.trim()).filter(Boolean) })} styles={{ root: { width: 220 } }} />
                    <IconButton iconProps={{ iconName: 'Delete' }} ariaLabel="Remove template" onClick={() => upd({ whatsapp: { ...s.whatsapp, templates: s.whatsapp.templates.filter((_, j) => j !== i) } })} />
                  </Stack>
                  <TextField label="Message" multiline rows={2} value={t.body} onChange={(_, v) => setT({ body: v || '' })} />
                </Item>
              );
            })}
          </Group>
        </>
      )
    },
    {
      key: 'marketing', group: 'Communication', title: 'Email marketing', icon: 'Megaphone', keywords: 'campaign unsubscribe consent opt-in newsletter',
      description: 'Campaign emails to consented leads and contacts.',
      render: () => (
        <Group title="Campaign email" description="Campaigns › Send email sends one by one from your mailbox, with a personal unsubscribe link.">
          <Row label="Unsubscribe link" help="HTTP URL of the flow “Unsubscribe”.">
            <TextField ariaLabel="Unsubscribe link" placeholder="https://…" value={s.marketing.unsubscribeUrl} errorMessage={urlError(s.marketing.unsubscribeUrl)} onChange={(_, v) => upd({ marketing: { ...s.marketing, unsubscribeUrl: (v || '').trim() } })} />
          </Row>
          <Row label="Footer" help="Company name and postal address (required by anti-spam laws)."><TextField ariaLabel="Marketing footer" multiline rows={2} value={s.marketing.footer} onChange={(_, v) => upd({ marketing: { ...s.marketing, footer: v || '' } })} /></Row>
          {toggleRow('Leads need explicit opt-in', 'Otherwise only contacts need marketing consent.', s.marketing.requireLeadOptIn, v => upd({ marketing: { ...s.marketing, requireLeadOptIn: v } }))}
          <Row label="Pause between emails" help="Outlook allows about 30 emails a minute per mailbox."><Unit after="seconds"><TextField ariaLabel="Seconds between emails" type="number" value={String(s.marketing.secondsBetweenEmails)} styles={{ root: { width: 170 } }} onChange={(_, v) => upd({ marketing: { ...s.marketing, secondsBetweenEmails: Math.max(1, Number(v) || 2) } })} /></Unit></Row>
        </Group>
      )
    },
    {
      key: 'teams', group: 'Communication', title: 'Microsoft Teams', icon: 'TeamsLogo', keywords: 'teams channel notifications webhook default site personal app',
      description: 'Channel notifications, the default CRM site and the Teams app.',
      render: () => (
        <>
          <Group title="Channel notifications" description="Posts deals won, new leads, demo requests and cases to a Teams channel.">
            <Row label="Workflows webhook URL" help='Teams: channel › ⋯ › Workflows › "Post to a channel when a webhook request is received".'>
              <Stack horizontal tokens={{ childrenGap: 8 }}>
                <TextField ariaLabel="Teams webhook URL" styles={{ root: { flex: 1 } }} value={s.integrations.teamsWebhookUrl} disabled={!engine.hasFeature('webhooks')} onChange={(_, v) => upd({ integrations: { ...s.integrations, teamsWebhookUrl: v || '' } })} />
                <DefaultButton text="Test" disabled={!s.integrations.teamsWebhookUrl} onClick={() => engine.webhooks.test(s.integrations.teamsWebhookUrl).then(() => notify('Test sent.', MessageBarType.success)).catch(e => notify(errorText(e), MessageBarType.error))} />
              </Stack>
            </Row>
          </Group>
          <TeamsGroups host={host} tenantSiteUrl={tenantSiteUrl} />
        </>
      )
    },
    // ============================================================ Integrations
    {
      key: 'flows', group: 'Integrations', title: 'Power Automate', icon: 'Flow', keywords: 'power automate flows solution download import',
      description: 'The ready-made flow pack, pre-filled for this site.',
      render: () => (
        <>
          <Group title="Flow pack" description="One importable solution with every flow. Import it in Power Automate › Solutions › Import solution, then turn on the flows you need.">
            <Row label="Download" help="Pre-filled with this site, page and your email.">
              <PrimaryButton iconProps={{ iconName: 'Download' }} text="Download Power Automate solution" onClick={() => setFlowDialog(true)} />
            </Row>
          </Group>
          <Group title={`Included flows (${FLOWS.length})`} description="Premium = needs Power Automate Premium for the connection owner.">
            {FLOWS.map(([name, what, premium]) => (
              <div key={name} className="crm-set-check">
                <Icon iconName="Flow" className="crm-muted" />
                <Text style={{ flex: 1 }}><b>{name}</b> <span className="crm-muted">– {what}</span></Text>
                {premium && <span className="crm-pill" style={{ background: 'var(--crm-accent-soft)', color: 'var(--crm-accent)' }}>Premium</span>}
              </div>
            ))}
          </Group>
        </>
      )
    },
    {
      key: 'copilot', group: 'Integrations', title: 'Copilot', icon: 'Robot', keywords: 'copilot agent ai declarative',
      description: 'Microsoft 365 Copilot agent grounded in this CRM.',
      render: () => <CopilotGroup />
    },
    {
      key: 'webhooks', group: 'Integrations', title: 'Webhooks', icon: 'PlugConnected', keywords: 'webhook zapier make n8n api events',
      description: 'Send CRM events to Power Automate, Zapier, Make, n8n or your API.',
      render: () => (
        <Group title="Outgoing webhooks" description={'Each event is POSTed as JSON (text/plain; in Power Automate use json(triggerBody())). Payload: { event, occurredAt, tenant, summary, data }.'}
          action={<DefaultButton iconProps={{ iconName: 'Add' }} text="Add webhook" disabled={!engine.hasFeature('webhooks')} onClick={() => upd({ webhooks: [...s.webhooks, { id: uid(), name: 'New webhook', url: '', events: ['opportunity.won'], active: true }] })} />}>
          {!engine.hasFeature('webhooks') && <MessageBar messageBarType={MessageBarType.warning} styles={{ root: { margin: '10px 0' } }}>Webhooks require the Enterprise plan.</MessageBar>}
          {s.webhooks.length === 0 && <Text className="crm-muted" block styles={{ root: { padding: '10px 0' } }}>No webhooks yet.</Text>}
          {s.webhooks.map(h => (
            <Item key={h.id}>
              <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="end">
                <TextField label="Name" value={h.name} onChange={(_, v) => setHook(h.id, { name: v || '' })} styles={{ root: { width: 180 } }} />
                <TextField label="URL (https)" value={h.url} onChange={(_, v) => setHook(h.id, { url: v || '' })} styles={{ root: { flex: 1, minWidth: 260 } }} />
                <Toggle label="Active" checked={h.active} onChange={(_, c) => setHook(h.id, { active: !!c })} />
                <DefaultButton text="Test" onClick={() => engine.webhooks.test(h.url).then(() => notify('Test sent.', MessageBarType.success)).catch(e => notify(errorText(e), MessageBarType.error))} />
                <IconButton iconProps={{ iconName: 'Delete' }} ariaLabel="Delete webhook" onClick={() => upd({ webhooks: s.webhooks.filter(x => x.id !== h.id) })} />
              </Stack>
              <Dropdown label="Events" multiSelect selectedKeys={h.events} options={[{ key: '*', text: 'All events' }, ...WEBHOOK_EVENTS.map(e => ({ key: e, text: e }))]}
                onChange={(_, o) => o && setHook(h.id, { events: o.selected ? [...h.events, String(o.key)] : h.events.filter(x => x !== o.key) })} />
            </Item>
          ))}
        </Group>
      )
    },
    {
      key: 'connect', group: 'Integrations', title: 'AppSource & Power BI', icon: 'BarChartVertical', keywords: 'appsource partner center leads power bi reporting',
      description: 'Marketplace leads and Power BI reporting.',
      render: () => (
        <>
          <Group title="Microsoft AppSource / Partner Center leads">
            <Text block styles={{ root: { padding: '10px 0' } }}>
              Partner Center › your offer › Customer leads › <b>HTTPS endpoint</b> › paste the HTTP URL of the flow <b>AppSource lead intake</b> › Validate.
              New marketplace leads then arrive in <b>{engine.def('leads').listTitle}</b>, are scored and routed automatically.
            </Text>
          </Group>
          <Group title="Power BI" description="Data stays protected by SharePoint permissions.">
            <Row label="CRM site" help="Power BI Desktop › Get data › SharePoint Online list › implementation 2.0 › select the CRM lists.">
              <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
                <Text styles={{ root: { wordBreak: 'break-all', flex: 1 } }}><code>{engine.ds.rest.webUrl}</code></Text>
                <DefaultButton iconProps={{ iconName: 'Copy' }} text="Copy" onClick={() => { (navigator.clipboard ? navigator.clipboard.writeText(engine.ds.rest.webUrl) : Promise.reject(new Error('no clipboard'))).then(() => notify('Site URL copied.')).catch(() => window.prompt('Site URL', engine.ds.rest.webUrl)); }} />
              </Stack>
            </Row>
            <Row label="Relationships" help="Set these in the Power BI model.">
              <Text variant="small">AccountRef, ContactRef, OpportunityRef, QuoteRef, SubscriptionRef → ID of the matching list. Activities: RelatedType + RelatedId.</Text>
            </Row>
          </Group>
        </>
      )
    },
    // ============================================================ Data & security
    {
      key: 'security', group: 'Data & security', title: 'Roles & permissions', icon: 'Lock', keywords: 'roles groups permissions record level security item sharepoint',
      description: 'CRM roles, SharePoint groups and record-level security.',
      render: () => (
        <>
          <Group title="Roles" description={`Your role: ${engine.role}. Site owners are always administrators.`}>
            <ul className="crm-ul">
              <li><b>Administrators</b> – everything, including settings, security and licence.</li>
              <li><b>Sales managers</b> – all records, delete, import, bulk assignment, forecasts.</li>
              <li><b>Sales</b> – leads, accounts, contacts, deals, quotes, demos, subscriptions, invoices, tasks.</li>
              <li><b>Support</b> – cases, knowledge base, accounts, contacts, activities.</li>
              <li><b>Readers</b> – read-only.</li>
            </ul>
          </Group>
          <Group title="SharePoint groups" description="Group names used for each role. Then add people in Site permissions › Groups."
            action={<PrimaryButton text="Create groups & apply list permissions" disabled={busy || !engine.hasFeature('security')} onClick={() => save(async () => { await engine.ds.applySecurity(setProgress); })} />}>
            {(['admins', 'managers', 'sales', 'support', 'readers'] as const).map(k => (
              <Row key={k} label={{ admins: 'Administrators', managers: 'Sales managers', sales: 'Sales', support: 'Support', readers: 'Readers' }[k]}>
                <TextField ariaLabel={`${k} group`} value={s.groups[k]} onChange={(_, v) => upd({ groups: { ...s.groups, [k]: v || '' } })} />
              </Row>
            ))}
            <Text variant="small" className="crm-muted" block styles={{ root: { padding: '8px 0' } }}>
              {!engine.hasFeature('security') ? 'Requires the Enterprise plan. ' : ''}<Link href={`${engine.ds.rest.webUrl}/_layouts/15/groups.aspx`} target="_blank">Open Site permissions › Groups</Link>. Existing site permissions are kept.
            </Text>
          </Group>
          <Group title="Record-level security" description="Each lead, account, contact, deal, quote, case, subscription and invoice gets its own SharePoint permissions: owner, owner's team manager, team members (as set below) and the CRM groups. Reps no longer reach other reps' records anywhere in Microsoft 365.">
            {toggleRow('Enforce record-level security', 'For records others create, turn on the flow “Record security” (runs every 15 minutes).', s.security.itemLevel, v => upd({ security: { ...s.security, itemLevel: v } }), !engine.hasFeature('security'))}
            <Row label="Team members">
              <Dropdown ariaLabel="Team access" selectedKey={s.security.teamAccess} options={[{ key: 'none', text: 'No access to each other' }, { key: 'read', text: "Can read each other's records" }, { key: 'edit', text: "Can edit each other's records" }]}
                onChange={(_, o) => upd({ security: { ...s.security, teamAccess: (o?.key as 'none' | 'read' | 'edit') || 'read' } })} />
            </Row>
            <Row label={settings.security.itemLevel ? 'Apply to existing records' : 'Remove record permissions'} help={dirty ? 'Save the settings first.' : 'Up to 50,000 uniquely-permissioned items per list are supported by SharePoint.'}>
              <DefaultButton text={settings.security.itemLevel ? 'Apply now' : 'Remove now'} disabled={busy || dirty || !engine.hasFeature('security')}
                onClick={() => save(async () => {
                  const r = await engine.applyItemSecurity(settings.security.itemLevel, setProgress);
                  setProgress(`${settings.security.itemLevel ? 'Secured' : 'Reset'} ${r.secured} record(s); ${r.skipped} already up to date.`);
                })} />
            </Row>
            {toggleRow('Reps only see their own records in the CRM', 'A screen filter only. Use record-level security for real protection.', s.restrictRepsToOwnRecords, v => upd({ restrictRepsToOwnRecords: v }))}
          </Group>
        </>
      )
    },
    {
      key: 'fields', group: 'Data & security', title: 'Custom fields', icon: 'FieldNotChanged', keywords: 'custom field column add',
      description: 'Add your own fields to any record type.',
      render: () => (
        <>
          {!engine.hasFeature('customFields') && <MessageBar messageBarType={MessageBarType.warning} styles={{ root: { marginBottom: 12 } }}>Custom fields require the Enterprise plan.</MessageBar>}
          <Group title="Add a field">
            <Row label="Record type"><Dropdown ariaLabel="Record type" selectedKey={cf.entity} options={ENTITIES.map(e => ({ key: e.key, text: engine.def(e.key).singular }))} onChange={(_, o) => setCf({ ...cf, entity: o?.key as EntityKey })} /></Row>
            <Row label="Label"><TextField ariaLabel="Field label" value={cf.label} onChange={(_, v) => setCf({ ...cf, label: v || '' })} /></Row>
            <Row label="Type">
              <Dropdown ariaLabel="Field type" selectedKey={cf.type} onChange={(_, o) => setCf({ ...cf, type: o?.key as FieldType })} options={[
                { key: 'text', text: 'Text' }, { key: 'note', text: 'Multi-line text' }, { key: 'number', text: 'Number' }, { key: 'currency', text: 'Currency' },
                { key: 'date', text: 'Date' }, { key: 'datetime', text: 'Date & time' }, { key: 'boolean', text: 'Yes / No' }, { key: 'choice', text: 'Choice' },
                { key: 'email', text: 'Email' }, { key: 'phone', text: 'Phone' }, { key: 'url', text: 'Link' }, { key: 'user', text: 'Person' }
              ]} />
            </Row>
            {cf.type === 'choice' && <Row label="Choices" help="One per line."><TextField ariaLabel="Choices" multiline rows={4} value={lines(cf.choices || [])} onChange={(_, v) => setCf({ ...cf, choices: parseLines(v || '') })} /></Row>}
            <Row label="Show in grids"><Checkbox ariaLabel="Show as a grid column" checked={!!cf.grid} onChange={(_, c) => setCf({ ...cf, grid: !!c })} /></Row>
            <Row label=""><DefaultButton iconProps={{ iconName: 'Add' }} text="Add field" disabled={!cf.label || !engine.hasFeature('customFields')} onClick={addCustom} /></Row>
          </Group>
          <Group title={`Your fields (${s.customFields.length})`} description="Removing a field hides it; the SharePoint column and its data are kept."
            action={<PrimaryButton text="Save & create columns" disabled={busy} onClick={() => save(() => engine.ds.provision(setProgress))} />}>
            {s.customFields.length === 0 && <Text className="crm-muted" block styles={{ root: { padding: '10px 0' } }}>None yet.</Text>}
            {s.customFields.map(c => (
              <div key={`${c.entity}-${c.key}`} className="crm-set-check">
                <Icon iconName="FieldNotChanged" className="crm-muted" />
                <Text style={{ flex: 1 }}><b>{c.label}</b> <span className="crm-muted">· {engine.def(c.entity).singular} · {c.type} · column cf{c.key}</span></Text>
                <IconButton iconProps={{ iconName: 'Delete' }} ariaLabel={`Remove ${c.label}`} onClick={() => upd({ customFields: s.customFields.filter(x => !(x.entity === c.entity && x.key === c.key)) })} />
              </div>
            ))}
          </Group>
        </>
      )
    },
    {
      key: 'data', group: 'Data & security', title: 'Data, audit & backup', icon: 'Database', keywords: 'audit log backup export data retention',
      description: 'Audit log and a full backup of your CRM data.',
      render: () => (
        <>
          <Group title="Audit">
            {toggleRow('Audit log', 'Records who changed which field, when, from what to what.', s.auditEnabled, v => upd({ auditEnabled: v }), !engine.hasFeature('audit'))}
          </Group>
          <Group title="Backup" description="All CRM data also lives in SharePoint lists on this site, covered by Microsoft 365 retention, versioning and backup.">
            <Row label="Full backup" help="Every record and all settings as one JSON file.">
              <DefaultButton iconProps={{ iconName: 'Download' }} text="Download backup" disabled={busy} onClick={backup} />
            </Row>
          </Group>
        </>
      )
    }
  ];

  const q = query.trim().toLowerCase();
  const visible = q ? sections.filter(x => `${x.title} ${x.description} ${x.keywords} ${x.group}`.toLowerCase().indexOf(q) >= 0) : sections;
  const current = sections.find(x => x.key === section) || sections[0];
  const groups = visible.map(x => x.group).filter((g, i, a) => a.indexOf(g) === i);

  return (
    <div className={`crm-settings${isNarrow ? ' crm-settings-narrow' : ''}`}>
      {flowDialog && <FlowPackDialog onClose={() => setFlowDialog(false)} />}
      {isNarrow ? (
        <Dropdown ariaLabel="Settings page" selectedKey={current.key} onChange={(_, o) => go(String(o?.key))}
          options={sections.map(x => ({ key: x.key, text: x.group ? `${x.group} › ${x.title}` : x.title }))} />
      ) : (
        <nav className="crm-set-nav" aria-label="Settings">
          <SearchBox placeholder="Search settings" value={query} onChange={(_, v) => setQuery(v || '')} styles={{ root: { marginBottom: 8 } }} />
          {groups.map(g => (
            <div key={g || 'top'}>
              {g && <div className="crm-set-nav-group">{g}</div>}
              {visible.filter(x => x.group === g).map(x => (
                <button key={x.key} type="button" className={`crm-set-link${x.key === current.key ? ' on' : ''}`} aria-current={x.key === current.key ? 'page' : undefined} onClick={() => go(x.key)}>
                  <Icon iconName={x.icon} /><span>{x.title}</span>
                </button>
              ))}
            </div>
          ))}
          {!visible.length && <Text className="crm-muted" block styles={{ root: { padding: 8 } }}>No settings match “{query}”.</Text>}
        </nav>
      )}
      <div className="crm-set-main" ref={mainRef}>
        {(dirty || busy) && (
          <div className="crm-set-savebar" role="status">
            {busy ? <Spinner size={SpinnerSize.small} label={progress || 'Saving…'} labelPosition="right" /> : <><Icon iconName="Edit" /><Text style={{ flex: 1 }}>Unsaved changes</Text></>}
            {!busy && <DefaultButton text="Discard" onClick={discard} />}
            {!busy && <PrimaryButton iconProps={{ iconName: 'Save' }} text="Save changes" onClick={() => save()} />}
          </div>
        )}
        <div className="crm-set-head">
          <span className="crm-set-icon"><Icon iconName={current.icon} /></span>
          <div>
            {current.group && <Text variant="small" className="crm-muted" block>{current.group}</Text>}
            <Text as="h2" styles={{ root: { fontSize: 22, fontWeight: 600, margin: 0 } }}>{current.title}</Text>
            <Text className="crm-muted" block>{current.description}</Text>
          </div>
        </div>
        {current.render()}
      </div>
    </div>
  );
};
