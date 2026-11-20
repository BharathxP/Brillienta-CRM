// Brillienta Helpdesk connector: reads tickets from the Helpdesk site and raises new ones.
//
// Helpdesk keeps tickets in the SharePoint list "<prefix>Tickets" on its own site. The CRM reads them with the
// signed-in user's permissions (Helpdesk agents see every ticket, everyone else only their own requests).
// New tickets are saved with CopilotPending = true, the hand-off Helpdesk already uses for Microsoft 365 Copilot:
// the next Helpdesk page an agent opens numbers the ticket, routes it, applies the SLA and notifies people.
// Tickets raised from the CRM carry tags (CRM, crm-acc-<id>, crm-con-<id>, crm-case-<id>) so they can be found
// again even when the customer is not a Microsoft 365 user.
import { Engine } from './Engine';
import { IRecord } from '../models/model';

export interface IHelpdeskTicket {
  id: number;
  number: string;
  title: string;
  status: string;
  /** New | Open | Pending | Resolved | Closed */
  kind: string;
  priority: string;
  requester: string;
  requesterEmail: string;
  assignee: string;
  created: string;
  modified: string;
  resolutionDue?: string;
  breached: boolean;
  tags: string[];
  /** Saved by the CRM and waiting for Helpdesk to number and route it. */
  pending: boolean;
  link: string;
}

export interface IHelpdeskTarget {
  /** Email addresses whose tickets belong to this record. */
  emails: string[];
  /** Tags whose tickets belong to this record (crm-acc-12, …). */
  tags: string[];
}

const SELECT = 'Id,Title,TicketNumber,TicketKey,StatusName,StatusKind,PriorityName,RequesterName,RequesterEmail,AssigneeName,Created,Modified,ResolutionDue,SlaResolutionBreached,TagsText,CopilotPending';
const CACHE_MS = 60000;
const RECENT = 1000;

const esc = (s: string): string => s.replace(/'/g, "''");
const parseTags = (v: string): string[] => String(v || '').split(/[;,]/).map(t => t.trim()).filter(Boolean);
const textToHtml = (t: string): string => String(t || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').split(/\r?\n/).map(l => `<p>${l || '&nbsp;'}</p>`).join('');

export const isOpenTicket = (t: IHelpdeskTicket): boolean => t.kind !== 'Resolved' && t.kind !== 'Closed';

export class Helpdesk {
  private cache?: { at: number; key: string; rows: IHelpdeskTicket[] };

  constructor(private readonly engine: Engine) {}

  private get cfg(): { enabled: boolean; siteUrl: string; listPrefix: string; pageUrl: string } { return this.engine.settings.helpdesk; }

  public get configured(): boolean { return !!(this.cfg.enabled && /^https:\/\//i.test(this.cfg.siteUrl)); }

  private site(siteUrl?: string): string { return (siteUrl || this.cfg.siteUrl).replace(/\/+$/, '').replace(/\/SitePages\/.*$/i, ''); }

  private listUrl(siteUrl?: string, prefix?: string): string {
    const title = `${(prefix ?? this.cfg.listPrefix) || 'HD_'}Tickets`;
    return `${this.site(siteUrl)}/_api/web/lists/getbytitle('${encodeURIComponent(esc(title))}')`;
  }

  public pageUrl(): string { return this.cfg.pageUrl || `${this.site()}/SitePages/Helpdesk.aspx`; }

  private map(r: Record<string, unknown>): IHelpdeskTicket {
    const id = Number(r.Id);
    return {
      id, number: String(r.TicketNumber || (r.CopilotPending ? 'New' : `#${id}`)), title: String(r.Title || ''),
      status: String(r.StatusName || r.StatusKind || 'New'), kind: String(r.StatusKind || 'New'), priority: String(r.PriorityName || ''),
      requester: String(r.RequesterName || ''), requesterEmail: String(r.RequesterEmail || '').toLowerCase(), assignee: String(r.AssigneeName || ''),
      created: String(r.Created || ''), modified: String(r.Modified || ''), resolutionDue: r.ResolutionDue ? String(r.ResolutionDue) : undefined,
      breached: !!r.SlaResolutionBreached, tags: parseTags(String(r.TagsText || '')), pending: !!r.CopilotPending,
      link: `${this.pageUrl()}#/ticket/${r.TicketKey || id}`
    };
  }

  /** Checks the connection; resolves with the number of tickets visible to the current user. */
  public async test(siteUrl: string, prefix: string): Promise<{ visible: number }> {
    const rows = await this.engine.ds.rest.get<{ value: unknown[] }>(`${this.listUrl(siteUrl, prefix)}/items?$select=Id&$top=500`);
    return { visible: (rows.value || []).length };
  }

  /** The most recent tickets (cached for a minute). */
  public async recent(force?: boolean): Promise<IHelpdeskTicket[]> {
    const key = `${this.cfg.siteUrl}|${this.cfg.listPrefix}`;
    if (!force && this.cache && this.cache.key === key && Date.now() - this.cache.at < CACHE_MS) { return this.cache.rows; }
    const res = await this.engine.ds.rest.get<{ value: Record<string, unknown>[] }>(`${this.listUrl()}/items?$select=${SELECT}&$orderby=Id desc&$top=${RECENT}`);
    const rows = (res.value || []).map(r => this.map(r));
    this.cache = { at: Date.now(), key, rows };
    return rows;
  }

  /** Who and which tags a record's tickets belong to. */
  public targetFor(entity: string, r: IRecord): IHelpdeskTarget {
    const contacts = this.engine.data.contacts || [];
    const mail = (x?: string): string => String(x || '').trim().toLowerCase();
    if (entity === 'accounts') {
      return { emails: contacts.filter(c => c.accountRef === r.id).map(c => mail(c.email)).filter(Boolean), tags: [`crm-acc-${r.id}`] };
    }
    if (entity === 'contacts') { return { emails: [mail(r.email)].filter(Boolean), tags: [`crm-con-${r.id}`] }; }
    if (entity === 'cases') { return { emails: [], tags: [`crm-case-${r.id}`] }; }
    return { emails: [], tags: [] };
  }

  public matches(t: IHelpdeskTicket, target: IHelpdeskTarget): boolean {
    return (!!t.requesterEmail && target.emails.indexOf(t.requesterEmail) >= 0) || t.tags.some(x => target.tags.indexOf(x) >= 0);
  }

  public async ticketsFor(entity: string, r: IRecord, force?: boolean): Promise<IHelpdeskTicket[]> {
    const target = this.targetFor(entity, r);
    return (await this.recent(force)).filter(t => this.matches(t, target));
  }

  /** Tickets of any CRM customer (contacts' emails or CRM tags). */
  public async customerTickets(): Promise<IHelpdeskTicket[]> {
    const emails: Record<string, true> = {};
    (this.engine.data.contacts || []).forEach(c => { const e = String(c.email || '').trim().toLowerCase(); if (e) { emails[e] = true; } });
    return (await this.recent()).filter(t => emails[t.requesterEmail] || t.tags.indexOf('CRM') >= 0);
  }

  /** The CRM account a ticket belongs to, if any. */
  public accountOf(t: IHelpdeskTicket): IRecord | undefined {
    const tag = t.tags.find(x => /^crm-acc-\d+$/.test(x));
    if (tag) { return this.engine.find('accounts', Number(tag.slice(8))); }
    const c = (this.engine.data.contacts || []).find(x => String(x.email || '').toLowerCase() === t.requesterEmail);
    return c?.accountRef ? this.engine.find('accounts', c.accountRef) : undefined;
  }

  /** Raises a ticket. Helpdesk numbers, routes and notifies it the next time an agent has Helpdesk open. */
  public async create(input: { title: string; details: string; requesterEmail?: string; accountId?: number; contactId?: number; caseId?: number }): Promise<IHelpdeskTicket> {
    const tags = ['CRM'];
    if (input.accountId) { tags.push(`crm-acc-${input.accountId}`); }
    if (input.contactId) { tags.push(`crm-con-${input.contactId}`); }
    if (input.caseId) { tags.push(`crm-case-${input.caseId}`); }
    const body = {
      Title: input.title.slice(0, 255),
      Details: textToHtml(input.details),
      RequesterEmail: String(input.requesterEmail || '').trim().toLowerCase(),
      TicketSource: 'Copilot',
      TagsText: `;${tags.join(';')};`,
      CopilotPending: true
    };
    const res = await this.engine.ds.rest.post<{ Id?: number; ID?: number }>(`${this.listUrl()}/items`, body);
    this.cache = undefined;
    const id = Number(res.Id || res.ID);
    return this.map({ ...body, Id: id, Created: new Date().toISOString(), Modified: new Date().toISOString(), StatusKind: 'New' });
  }
}
