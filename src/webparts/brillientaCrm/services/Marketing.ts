// Email campaigns to consented leads and contacts, sent one by one from the user's Outlook mailbox
// (Exchange Online allows about 30 messages a minute and 10,000 recipients a day per mailbox).
// Every email carries a personal unsubscribe link served by the Power Automate flow "Unsubscribe".
import { EntityKey, IRecord } from '../models/model';
import type { Engine } from './Engine';
import { exportCsv } from '../components/utils';

export interface IAudienceFilter {
  leads: boolean;
  contacts: boolean;
  tag: string;
  leadSource: string;
  country: string;
  campaignLeadsOnly: boolean;
}

export interface IRecipient { key: EntityKey; record: IRecord; email: string }

export const MAX_PER_RUN = 500;

export function audience(engine: Engine, f: IAudienceFilter, campaignId?: number): { ok: IRecipient[]; excluded: { noConsent: number; unsubscribed: number; noEmail: number } } {
  const ok: IRecipient[] = [];
  const excluded = { noConsent: 0, unsubscribed: 0, noEmail: 0 };
  const seen = new Set<string>();
  const match = (r: IRecord): boolean =>
    (!f.tag || (Array.isArray(r.tags) ? r.tags : String(r.tags || '').split(',')).some((t: string) => t.trim().toLowerCase() === f.tag.trim().toLowerCase())) &&
    (!f.country || String(r.country || '').toLowerCase() === f.country.toLowerCase());
  const take = (key: EntityKey, rows: IRecord[], needsOptIn: boolean): void => {
    rows.filter(match).forEach(r => {
      const email = String(r.email || '').trim().toLowerCase();
      if (!email || !/@/.test(email)) { excluded.noEmail++; return; }
      if (r.unsubscribed) { excluded.unsubscribed++; return; }
      if (needsOptIn && !r.marketingConsent) { excluded.noConsent++; return; }
      if (seen.has(email)) { return; }
      seen.add(email);
      ok.push({ key, record: r, email });
    });
  };
  if (f.contacts) { take('contacts', engine.data.contacts || [], true); }
  if (f.leads) {
    const leads = (engine.data.leads || []).filter(l => l.status !== 'Converted' && l.status !== 'Unqualified' &&
      (!f.leadSource || l.leadSource === f.leadSource) && (!f.campaignLeadsOnly || (campaignId && l.campaignRef === campaignId)));
    take('leads', leads, engine.settings.marketing.requireLeadOptIn);
  }
  return { ok, excluded };
}

function token(): string {
  const b = new Uint8Array(12);
  window.crypto.getRandomValues(b);
  return Array.prototype.map.call(b, (x: number) => ('0' + x.toString(16)).slice(-2)).join('');
}

export function unsubscribeLink(engine: Engine, key: EntityKey, r: IRecord): string {
  const base = (engine.settings.marketing.unsubscribeUrl || '').trim();
  if (!base || !r.unsubToken) { return ''; }
  return `${base}${base.indexOf('?') >= 0 ? '&' : '?'}e=${key === 'leads' ? 'L' : 'C'}${r.id}&t=${r.unsubToken}`;
}

const sleep = (ms: number): Promise<void> => new Promise(resolve => setTimeout(resolve, ms));

/** Sends the campaign; returns the number sent. Stops early when `cancelled()` returns true. */
export async function sendCampaign(engine: Engine, campaign: IRecord, recipients: IRecipient[], subject: string, body: string,
  onProgress: (sent: number, failed: number) => void, cancelled: () => boolean): Promise<{ sent: number; failed: string[] }> {
  if (!/^https:\/\//i.test(engine.settings.marketing.unsubscribeUrl || '')) { throw new Error('Set the unsubscribe link in Settings › Integrations › Email marketing first (flow “Unsubscribe”).'); }
  if (!engine.graph.available || !engine.settings.integrations.graphEmail) { throw new Error('Outlook email is not available. Export the audience to Mailchimp or Brevo instead.'); }
  const list = recipients.slice(0, MAX_PER_RUN);
  const delay = Math.max(1, engine.settings.marketing.secondsBetweenEmails || 2) * 1000;
  const footer = engine.settings.marketing.footer || engine.settings.company.address || engine.settings.company.name || '';
  let sent = 0;
  const failed: string[] = [];
  for (const rc of list) {
    if (cancelled()) { break; }
    try {
      let r = rc.record;
      if (!r.unsubToken) { const t = token(); await engine.ds.save(rc.key, { id: r.id, unsubToken: t }); r = { ...r, unsubToken: t }; }
      const link = unsubscribeLink(engine, rc.key, r);
      const text = `${engine.fillTemplate(body, r)}\n\n—\n${footer}\nYou receive this email because you are in contact with ${engine.settings.company.name || 'us'}. Unsubscribe: ${link}`;
      await engine.graph.sendMail([rc.email], engine.fillTemplate(subject, r), text, []);
      await engine.ds.save('activities', {
        title: `Campaign email: ${engine.fillTemplate(subject, r)}`.slice(0, 255), details: `Campaign: ${campaign.title}`, activityType: 'Email',
        relatedType: engine.def(rc.key).activityName, relatedId: r.id, completed: true, dueDate: new Date().toISOString(), owner: engine.me, activitySource: 'CRM'
      });
      if (rc.key === 'leads' && !r.campaignRef) { await engine.ds.save('leads', { id: r.id, campaignRef: campaign.id }); }
      sent++;
    } catch {
      failed.push(rc.email);
    }
    onProgress(sent, failed.length);
    if (!cancelled()) { await sleep(delay); }
  }
  await engine.patch('campaigns', campaign.id!, { emailsSent: (Number(campaign.emailsSent) || 0) + sent, lastEmailAt: new Date().toISOString(), campaignStatus: campaign.campaignStatus === 'Planned' ? 'Active' : campaign.campaignStatus });
  await engine.logActivity('campaigns', campaign.id!, `Email sent to ${sent} recipient(s)`, `Subject: ${subject}${failed.length ? `\nFailed: ${failed.join(', ')}` : ''}`, 'Email');
  await Promise.all([engine.refresh('activities'), engine.refresh('leads'), engine.refresh('contacts')]);
  return { sent, failed };
}

/** Audience file in Mailchimp / Brevo import format. */
export function exportAudience(recipients: IRecipient[], campaign: IRecord): void {
  exportCsv(`audience-${String(campaign.title || 'campaign').replace(/[^\w-]+/g, '-')}.csv`, recipients.map(r => ({
    'Email Address': r.email, 'First Name': r.record.firstName || '', 'Last Name': r.record.lastName || '',
    Company: r.record.company || '', Country: r.record.country || '', Tags: [r.key === 'leads' ? 'Lead' : 'Contact', ...(Array.isArray(r.record.tags) ? r.record.tags : [])].join(', ')
  })));
}
