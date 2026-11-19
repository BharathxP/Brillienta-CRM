import { EntityKey, IRecord } from '../models/model';
import { Engine } from './Engine';
import { IGraphAddress, IGraphEvent, IGraphMessage } from './Integrations';

export interface ISyncResult {
  emails: number;
  meetings: number;
  scanned: number;
  at: string;
}

const LS_PREFIX = 'brillienta-crm:outlook-sync:';
const addr = (a?: IGraphAddress): string => String(a?.emailAddress?.address || '').trim().toLowerCase();
const domainOf = (email: string): string => email.split('@')[1] || '';

function lsGet(key: string): string | undefined {
  try { return window.localStorage.getItem(LS_PREFIX + key) || undefined; } catch { return undefined; }
}
function lsSet(key: string, v: string): void {
  try { window.localStorage.setItem(LS_PREFIX + key, v); } catch { /* storage unavailable */ }
}

/** Per-user state: opt-in and last successful sync time. */
export const syncState = {
  enabled: (userId?: number): boolean => lsGet(`on:${userId}`) === '1',
  setEnabled: (userId: number | undefined, on: boolean): void => lsSet(`on:${userId}`, on ? '1' : '0'),
  last: (userId?: number): string | undefined => lsGet(`last:${userId}`),
  setLast: (userId: number | undefined, iso: string): void => lsSet(`last:${userId}`, iso)
};

interface IMatch { key: EntityKey; record: IRecord }

/**
 * Logs the current user's Outlook email and meetings with CRM contacts and leads as activities.
 * Runs in the browser with the user's own Microsoft Graph permissions (Mail.Read, Calendars.ReadWrite).
 * Only conversations with known contacts/leads are logged; internal email is never logged.
 */
export async function syncOutlook(engine: Engine, opts: { full?: boolean } = {}): Promise<ISyncResult> {
  const s = engine.settings.integrations;
  const me = engine.me;
  if (!s.outlookSync) { throw new Error('Outlook sync is switched off in Settings › Integrations.'); }
  if (!engine.graph.available) { throw new Error('Microsoft Graph is not available here.'); }
  if (!engine.license.canWrite) { throw new Error(engine.license.message); }
  await engine.ensure(['contacts', 'leads', 'activities']);

  // email → record (contacts win over leads; converted/unqualified leads are skipped)
  const index: Record<string, IMatch> = {};
  (engine.data.leads || []).forEach(l => {
    const e = String(l.email || '').toLowerCase();
    if (e && l.status !== 'Converted' && l.status !== 'Unqualified') { index[e] = { key: 'leads', record: l }; }
  });
  (engine.data.contacts || []).forEach(c => { const e = String(c.email || '').toLowerCase(); if (e) { index[e] = { key: 'contacts', record: c }; } });

  const internal = new Set<string>([domainOf(String(me?.email || '').toLowerCase()), ...s.internalDomains.map(d => d.trim().toLowerCase().replace(/^@/, ''))].filter(Boolean));
  const isInternal = (email: string): boolean => internal.has(domainOf(email));
  const known = new Set((engine.data.activities || []).map(a => String(a.externalId || '')).filter(Boolean));

  const now = new Date();
  const last = opts.full ? undefined : syncState.last(me?.id);
  const since = last ? new Date(new Date(last).getTime() - 10 * 60000) : new Date(now.getTime() - s.syncLookbackDays * 86400000);
  const result: ISyncResult = { emails: 0, meetings: 0, scanned: 0, at: now.toISOString() };

  const log = async (m: IMatch, ext: string, rec: IRecord): Promise<boolean> => {
    if (known.has(ext)) { return false; }
    known.add(ext);
    await engine.ds.save('activities', {
      ...rec, relatedType: engine.def(m.key).activityName, relatedId: m.record.id, completed: rec.completed ?? true,
      owner: me, externalId: ext.slice(0, 255), activitySource: 'Outlook'
    });
    return true;
  };

  // ---------------- email (received and sent)
  const [inbox, sent] = await Promise.all([
    engine.graph.listMessages('inbox', since.toISOString()),
    engine.graph.listMessages('sentitems', since.toISOString())
  ]);
  const messages: { m: IGraphMessage; outbound: boolean }[] = [
    ...inbox.map(m => ({ m, outbound: false })), ...sent.map(m => ({ m, outbound: true }))
  ];
  for (const { m, outbound } of messages) {
    result.scanned++;
    const people = [m.from, ...(m.toRecipients || []), ...(m.ccRecipients || [])].map(addr).filter(e => e && !isInternal(e));
    const matches = people.map(e => index[e]).filter((x, i, arr) => !!x && arr.indexOf(x) === i) as IMatch[];
    if (!matches.length) { continue; }
    const when = (outbound ? m.sentDateTime : m.receivedDateTime) || now.toISOString();
    const from = m.from?.emailAddress?.name || addr(m.from);
    const to = (m.toRecipients || []).map(r => r.emailAddress?.name || addr(r)).join(', ');
    for (const match of matches) {
      const saved = await log(match, `mail:${m.internetMessageId || m.id}|${match.key}:${match.record.id}`, {
        title: `${outbound ? 'Sent' : 'Received'}: ${m.subject || '(no subject)'}`.slice(0, 255),
        activityType: 'Email', dueDate: when,
        details: `From: ${from}\nTo: ${to}\n\n${(m.bodyPreview || '').trim()}`,
        webLink: m.webLink || ''
      });
      if (saved) { result.emails++; }
    }
  }

  // ---------------- meetings (past lookback window and the next 30 days)
  const evStart = new Date(since.getTime());
  const evEnd = new Date(now.getTime() + 30 * 86400000);
  const events: IGraphEvent[] = await engine.graph.listEvents(evStart.toISOString(), evEnd.toISOString());
  for (const ev of events) {
    result.scanned++;
    if (ev.isCancelled) { continue; }
    const people = [ev.organizer, ...(ev.attendees || [])].map(addr).filter(e => e && !isInternal(e));
    const matches = people.map(e => index[e]).filter((x, i, arr) => !!x && arr.indexOf(x) === i) as IMatch[];
    if (!matches.length) { continue; }
    const start = ev.start ? new Date(`${ev.start.dateTime}${/Z|[+-]\d\d:?\d\d$/.test(ev.start.dateTime) ? '' : 'Z'}`) : now;
    const end = ev.end ? new Date(`${ev.end.dateTime}${/Z|[+-]\d\d:?\d\d$/.test(ev.end.dateTime) ? '' : 'Z'}`) : start;
    const attendees = (ev.attendees || []).map(a => a.emailAddress?.name || addr(a)).join(', ');
    for (const match of matches) {
      const saved = await log(match, `event:${ev.iCalUId || ev.id}|${match.key}:${match.record.id}`, {
        title: `Meeting: ${ev.subject || '(no subject)'}`.slice(0, 255),
        activityType: 'Meeting', dueDate: start.toISOString(), completed: end.getTime() < now.getTime(),
        details: `${start.toLocaleString()} – ${end.toLocaleTimeString()}\nAttendees: ${attendees}${ev.onlineMeeting?.joinUrl ? `\nTeams: ${ev.onlineMeeting.joinUrl}` : ''}\n\n${(ev.bodyPreview || '').trim()}`,
        webLink: ev.webLink || ''
      });
      if (saved) { result.meetings++; }
    }
  }

  syncState.setLast(me?.id, result.at);
  if (result.emails || result.meetings) { await engine.refresh('activities'); }
  return result;
}
