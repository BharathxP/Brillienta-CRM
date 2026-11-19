// Microsoft 365 Copilot declarative agent for Brillienta CRM, built in the browser for the customer's own site.
// The agent answers from the CRM site's SharePoint content with each user's own permissions (no extra API, no data copy).
import { ENTITIES } from '../models/model';
import { MARK } from '../components/logo';
import { downloadBlob, zipFiles } from './FlowPackage';

export interface IAgentOptions { siteUrl: string; pageUrl: string; companyName: string; agentName: string; version?: string }

/** Stable GUID from a string (FNV-1a), so re-downloading for the same site updates the same app. */
export function stableGuid(seed: string): string {
  let hex = '';
  for (let round = 0; hex.length < 32; round++) {
    let h = 0x811c9dc5 ^ round;
    for (let i = 0; i < seed.length; i++) { h ^= seed.charCodeAt(i); h = Math.imul(h, 0x01000193) >>> 0; }
    hex += ('0000000' + h.toString(16)).slice(-8);
  }
  const x = hex.slice(0, 32).split('');
  x[12] = '4';
  x[16] = '89ab'[parseInt(x[16], 16) % 4];
  const s = x.join('');
  return `${s.slice(0, 8)}-${s.slice(8, 12)}-${s.slice(12, 16)}-${s.slice(16, 20)}-${s.slice(20)}`;
}

export function agentInstructions(o: IAgentOptions): string {
  const lists = ENTITIES.map(d => `- ${d.listTitle}: ${d.plural.toLowerCase()} (${d.description})`).join('\n');
  return `You are ${o.agentName}, the CRM assistant for ${o.companyName || 'our company'}. Answer questions about customers, leads, deals, quotes, invoices, subscriptions and support cases using only the Brillienta CRM data on the SharePoint site ${o.siteUrl}.

The CRM is stored in these SharePoint lists:
${lists}
- CRM Activities: emails, meetings, calls, WhatsApp messages, notes and tasks; RelatedType and RelatedId link an activity to a record.

How records link: AccountRef, ContactRef, OpportunityRef, QuoteRef and SubscriptionRef hold the ID of the related item. Opportunity Stage is one of the pipeline stages; "Closed Won" and "Closed Lost" are closed.

Guidelines:
- Be concise. Lead with the answer, then 2-5 bullets. Use tables for lists of records (name, owner, amount, date, status).
- Always say which list and record you used, e.g. "Opportunity: Contoso - HR portal". Never invent records, amounts or dates; if you cannot find something, say so and suggest where to look in the CRM.
- For a summary of a customer: open deals and amounts, last activities, open cases (SLA), active subscriptions and renewal dates, unpaid invoices.
- For drafting emails: friendly and professional, short paragraphs, include a clear next step. Do not include internal notes, discounts that need approval, or other customers' data.
- Money is in the CRM currency. Dates: show as day month year.
- You cannot change CRM data. To update a record, give the user this link: ${o.pageUrl}`;
}

export function agentFiles(o: IAgentOptions): Record<string, string> {
  const id = stableGuid(`brillienta-crm-agent|${o.siteUrl.toLowerCase()}`);
  const short = o.agentName.slice(0, 30);
  const manifest = {
    $schema: 'https://developer.microsoft.com/en-us/json-schemas/teams/v1.19/MicrosoftTeams.schema.json',
    manifestVersion: '1.19',
    version: o.version || '1.0.0',
    id,
    developer: { name: 'Brillienta', websiteUrl: 'https://brillienta.com', privacyUrl: 'https://brillienta.com/privacy', termsOfUseUrl: 'https://brillienta.com/terms' },
    icons: { color: 'color.png', outline: 'outline.png' },
    name: { short, full: `${o.agentName} (Brillienta CRM)`.slice(0, 100) },
    description: {
      short: 'Ask about customers, deals, cases and renewals in Brillienta CRM.',
      full: `Copilot agent for Brillienta CRM on ${o.siteUrl}. Summarises accounts and deals, finds renewals, overdue invoices and SLA breaches, and drafts customer emails from CRM data. Uses each user's own SharePoint permissions.`.slice(0, 4000)
    },
    accentColor: '#008D62',
    copilotAgents: { declarativeAgents: [{ id: 'brillientaCrmAgent', file: 'declarativeAgent.json' }] },
    permissions: ['identity'],
    validDomains: []
  };
  const agent = {
    $schema: 'https://developer.microsoft.com/json-schemas/copilot/declarative-agent/v1.3/schema.json',
    version: 'v1.3',
    name: short,
    description: `CRM assistant for ${o.companyName || 'your team'}: customers, deals, quotes, invoices, renewals and support cases from Brillienta CRM.`.slice(0, 1000),
    instructions: agentInstructions(o).slice(0, 8000),
    capabilities: [{ name: 'OneDriveAndSharePoint', items_by_url: [{ url: o.siteUrl }] }],
    conversation_starters: [
      { title: 'My pipeline', text: 'Show my open opportunities closing this quarter with amount, stage and next step.' },
      { title: 'Account summary', text: 'Summarise the account Contoso: open deals, recent activity, open cases and renewals.' },
      { title: 'Renewals', text: 'Which subscriptions renew in the next 60 days, and what is their ARR?' },
      { title: 'Overdue invoices', text: 'List overdue invoices with customer, amount and days overdue.' },
      { title: 'SLA risk', text: 'Which open support cases have breached or are close to breaching their SLA?' },
      { title: 'Draft follow-up', text: 'Draft a follow-up email for my most recent demo, based on the CRM notes.' }
    ]
  };
  return { 'manifest.json': JSON.stringify(manifest, null, 2), 'declarativeAgent.json': JSON.stringify(agent, null, 2) };
}

function canvasPng(size: number, draw: (ctx: CanvasRenderingContext2D) => Promise<void> | void): Promise<Uint8Array> {
  const c = document.createElement('canvas');
  c.width = size; c.height = size;
  const ctx = c.getContext('2d')!;
  return Promise.resolve(draw(ctx)).then(() => new Promise<Uint8Array>((resolve, reject) => c.toBlob(b => {
    if (!b) { reject(new Error('Could not create the icon.')); return; }
    const r = new FileReader();
    r.onload = () => resolve(new Uint8Array(r.result as ArrayBuffer));
    r.onerror = () => reject(r.error);
    r.readAsArrayBuffer(b);
  }, 'image/png')));
}

/** 192×192 colour icon (logo mark on white) and 32×32 white outline icon on transparent. */
export async function agentIcons(): Promise<{ color: Uint8Array; outline: Uint8Array }> {
  const color = await canvasPng(192, ctx => new Promise<void>((resolve, reject) => {
    ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, 192, 192);
    const img = new Image();
    img.onload = () => { const w = 140; const h = w * 442 / 497; ctx.drawImage(img, (192 - w) / 2, (192 - h) / 2, w, h); resolve(); };
    img.onerror = () => reject(new Error('Could not draw the logo.'));
    img.src = MARK;
  }));
  const outline = await canvasPng(32, ctx => {
    ctx.fillStyle = '#ffffff';
    const bar = (x: number, y: number, w: number, h: number): void => {
      const r = h / 2;
      ctx.beginPath(); ctx.moveTo(x + r, y); ctx.lineTo(x + w, y); ctx.lineTo(x + w, y + h); ctx.lineTo(x + r, y + h);
      ctx.arc(x + r, y + r, r, Math.PI / 2, Math.PI * 1.5); ctx.closePath(); ctx.fill();
    };
    bar(13, 4, 17, 7); bar(8, 13, 17, 7); bar(3, 22, 17, 7);
  });
  return { color, outline };
}

export async function downloadCopilotAgent(o: IAgentOptions): Promise<string> {
  const icons = await agentIcons();
  const files: Record<string, string | Uint8Array> = { ...agentFiles(o), 'color.png': icons.color, 'outline.png': icons.outline };
  const name = `BrillientaCRM-CopilotAgent-${o.siteUrl.replace(/^https:\/\//, '').replace(/[^\w]+/g, '-').replace(/-+$/, '').slice(0, 60)}.zip`;
  downloadBlob(zipFiles(files), name);
  return name;
}
