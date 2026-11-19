// Lead and account enrichment. Instant part (no network): website and company name from the email domain.
// The "Lead enrichment" Power Automate flow then reads the company website (title, description, LinkedIn)
// and, when an Apollo.io key is set, firmographics (industry, size, location).
import { EntityKey, IRecord } from '../models/model';
import type { Engine } from './Engine';

const FREE_MAIL = ['gmail.com', 'googlemail.com', 'yahoo.com', 'yahoo.co.in', 'hotmail.com', 'outlook.com', 'live.com', 'msn.com', 'icloud.com', 'me.com',
  'aol.com', 'proton.me', 'protonmail.com', 'zoho.com', 'yandex.com', 'gmx.com', 'gmx.de', 'web.de', 'mail.com', 'rediffmail.com', 'qq.com', '163.com'];

export function companyDomain(r: IRecord): string {
  const fromSite = String(r.website || '').trim().replace(/^https?:\/\//i, '').replace(/^www\./i, '').split(/[/?#]/)[0].toLowerCase();
  if (fromSite) { return fromSite; }
  const d = String(r.email || '').split('@')[1]?.toLowerCase() || '';
  return d && FREE_MAIL.indexOf(d) < 0 ? d : '';
}

/** Fills what can be derived without any lookup; returns the fields that changed. */
export function quickEnrich(r: IRecord): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  const domain = companyDomain(r);
  if (!domain) { return out; }
  if (!r.website) { out.website = `https://${domain}`; }
  if ('company' in r || r.firstName !== undefined) {
    if (!r.company) { const base = domain.split('.')[0]; out.company = base.charAt(0).toUpperCase() + base.slice(1); }
  }
  return out;
}

export async function requestEnrichment(engine: Engine, key: EntityKey, r: IRecord): Promise<IRecord> {
  const quick = quickEnrich(key === 'leads' ? r : { ...r, company: r.title });
  if (key === 'accounts') { delete quick.company; }
  if (!companyDomain({ ...r, ...quick })) { throw new Error('No company website or business email to enrich from.'); }
  return engine.patch(key, r.id!, { ...quick, enrichStatus: 'Requested' });
}
