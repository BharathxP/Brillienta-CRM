// Dashboard AI: deal health and win likelihood, an AI forecast, a daily briefing, ranked next best actions and
// plain-language questions about the CRM data. Everything is computed in the browser from the records the user can
// already see; nothing is sent anywhere. The Copilot prompt is only copied when the user asks for it.
import { EntityKey, IRecord, LOST, WON } from '../models/model';
import { Engine } from './Engine';
import { relatedActivities } from './Insights';
import { fmtDate, fmtMoney, inRange, isOverdue, slaBreached } from '../components/utils';

const DAY = 86400000;
const daysSince = (iso?: string): number | undefined => (iso ? Math.floor((Date.now() - new Date(iso).getTime()) / DAY) : undefined);
const daysUntil = (iso?: string): number | undefined => (iso ? Math.ceil((new Date(iso).getTime() - Date.now()) / DAY) : undefined);
const clamp = (n: number, lo: number, hi: number): number => Math.max(lo, Math.min(hi, n));
const plural = (n: number, w: string): string => `${n} ${w}${n === 1 ? '' : 's'}`;

/** The records the dashboard currently shows (already filtered to the chosen scope). */
export interface IScope {
  leads: IRecord[];
  opps: IRecord[];
  demos: IRecord[];
  cases: IRecord[];
  quotes: IRecord[];
  invoices: IRecord[];
  subs: IRecord[];
  tasks: IRecord[];
  /** User ids whose quotas count (undefined = everyone). */
  quotaUsers?: number[];
  quarter: { start: Date; end: Date; label: string };
  prevQuarter: { start: Date; end: Date; label: string };
}

/* ------------------------------------------------------------------ deal health */

export type HealthLabel = 'Healthy' | 'Watch' | 'At risk';

export interface IDealHealth {
  /** Likelihood to win, 0–100. */
  score: number;
  label: HealthLabel;
  /** Reasons that lower the score, most important first. */
  risks: string[];
  /** Signals that raise it. */
  strengths: string[];
  /** Short, countable reason for each risk (same order). */
  kinds: string[];
}

/** Historical win rate (0–1) and average days from creation to a win, from the company's closed deals. */
export function history(engine: Engine): { winRate?: number; cycleDays?: number } {
  const all = engine.data.opportunities || [];
  const closed = all.filter(o => o.stage === WON || o.stage === LOST);
  const won = closed.filter(o => o.stage === WON);
  const cycles = won.map(o => (o.actualCloseDate && o.created ? (new Date(o.actualCloseDate).getTime() - new Date(o.created).getTime()) / DAY : NaN)).filter(n => isFinite(n) && n >= 0);
  return {
    winRate: closed.length >= 5 ? won.length / closed.length : undefined,
    cycleDays: cycles.length >= 3 ? cycles.reduce((t, n) => t + n, 0) / cycles.length : undefined
  };
}

export function dealHealth(engine: Engine, o: IRecord, hist: { winRate?: number; cycleDays?: number } = history(engine)): IDealHealth {
  if (o.stage === WON) { return { score: 100, label: 'Healthy', risks: [], strengths: ['Won'], kinds: [] }; }
  if (o.stage === LOST) { return { score: 0, label: 'At risk', risks: ['Lost'], strengths: [], kinds: [] }; }
  const risks: { w: number; t: string; k: string }[] = [];
  const strengths: string[] = [];
  let s = Number(o.probability) || 10;

  const acts = relatedActivities(engine, 'opportunities', o);
  const done = acts.filter(a => a.completed);
  const last = daysSince(done[0]?.dueDate || done[0]?.created) ?? daysSince(o.modified);
  if (last !== undefined) {
    if (last > 30) { s -= 25; risks.push({ w: 25, t: `No activity for ${last} days`, k: 'no recent activity' }); }
    else if (last > 14) { s -= 12; risks.push({ w: 12, t: `Quiet for ${last} days`, k: 'no recent activity' }); }
    else if (last <= 3) { s += 5; strengths.push('Recent activity'); }
  }
  const close = daysUntil(o.expectedCloseDate);
  if (close === undefined) { s -= 6; risks.push({ w: 6, t: 'No close date', k: 'no close date' }); }
  else if (close < 0) { s -= 18; risks.push({ w: 18, t: `Close date passed ${-close} days ago`, k: 'close date passed' }); }
  if (!o.contactRef) { s -= 8; risks.push({ w: 8, t: 'No contact linked', k: 'no contact linked' }); }
  const open = acts.filter(a => a.activityType === 'Task' && !a.completed);
  if (open.some(a => isOverdue(a.dueDate))) { s -= 8; risks.push({ w: 8, t: 'Overdue task', k: 'overdue tasks' }); }
  else if (!open.length) { s -= 5; risks.push({ w: 5, t: 'No next step planned', k: 'no next step planned' }); }
  else { strengths.push('Next step planned'); }

  const quotes = (engine.data.quotes || []).filter(q => q.opportunityRef === o.id);
  if (quotes.some(q => q.quoteStatus === 'Accepted' || q.signatureStatus === 'Signed')) { s += 15; strengths.push('Quote accepted'); }
  else if (quotes.some(q => q.signatureStatus === 'Viewed')) { s += 8; strengths.push('Customer viewed the quote'); }
  else if (quotes.some(q => q.quoteStatus === 'Sent' || q.signatureStatus === 'Sent')) { s += 3; strengths.push('Quote sent'); }
  else if ((Number(o.probability) || 0) >= 50) { s -= 6; risks.push({ w: 6, t: 'No quote at 50%+', k: 'no quote sent' }); }
  if (quotes.some(q => q.approvalStatus === 'Pending')) { s -= 4; risks.push({ w: 4, t: 'Discount approval pending', k: 'discount approval pending' }); }
  if (done.filter(a => a.activityType === 'Meeting').length >= 2) { s += 4; strengths.push('Several meetings held'); }

  const age = daysSince(o.created);
  if (hist.cycleDays && age !== undefined && age > hist.cycleDays * 1.75) { s -= 10; risks.push({ w: 10, t: `Open ${age} days (wins usually take ${Math.round(hist.cycleDays)})`, k: 'taking longer than usual' }); }
  // Lean on the company's own track record when there is one.
  if (hist.winRate !== undefined) { s = s * 0.8 + hist.winRate * 100 * 0.2; }

  const score = Math.round(clamp(s, 1, 99));
  return {
    score, label: score >= 60 ? 'Healthy' : score >= 35 ? 'Watch' : 'At risk',
    risks: risks.sort((a, b) => b.w - a.w).map(r => r.t), kinds: risks.map(r => r.k), strengths
  };
}

/* ------------------------------------------------------------------ forecast */

export interface IForecast {
  won: number;
  /** Won + open deals closing this quarter × win likelihood. */
  predicted: number;
  /** Won + deals with likelihood ≥ 70. */
  commit: number;
  /** Won + every open deal closing this quarter. */
  best: number;
  quota: number;
  /** Share of the quarter that has passed, 0–1. */
  elapsed: number;
  daysLeft: number;
  deals: number;
}

export function aiForecast(engine: Engine, sc: IScope, health: (o: IRecord) => IDealHealth): IForecast {
  const q = sc.quarter;
  const won = sc.opps.filter(o => o.stage === WON && inRange(o.actualCloseDate, q.start, q.end)).reduce((t, o) => t + engine.baseAmount(o), 0);
  // Open deals due this quarter, plus overdue ones (they still have to land somewhere).
  const due = sc.opps.filter(o => engine.openStages().indexOf(o.stage) >= 0 && o.expectedCloseDate && new Date(o.expectedCloseDate).getTime() < q.end.getTime());
  let predicted = won;
  let commit = won;
  let best = won;
  due.forEach(o => {
    const v = engine.baseAmount(o);
    const h = health(o).score;
    predicted += (v * h) / 100;
    best += v;
    if (h >= 70) { commit += v; }
  });
  const quotas = engine.settings.quotas || [];
  const quota = quotas.filter(x => !sc.quotaUsers || sc.quotaUsers.indexOf(x.user.id) >= 0).reduce((t, x) => t + (Number(x.quarterly) || 0), 0);
  const span = q.end.getTime() - q.start.getTime();
  return {
    won, predicted: Math.round(predicted), commit, best, quota,
    elapsed: clamp((Date.now() - q.start.getTime()) / span, 0, 1),
    daysLeft: Math.max(0, Math.ceil((q.end.getTime() - Date.now()) / DAY)), deals: due.length
  };
}

/** Expected revenue for the `n` months after the current one, from open deals × win likelihood. */
export function projectedMonths(engine: Engine, opps: IRecord[], health: (o: IRecord) => IDealHealth, n: number = 3): { label: string; short: string; value: number }[] {
  const now = new Date();
  const out: { label: string; short: string; value: number }[] = [];
  for (let i = 1; i <= n; i++) {
    const s = new Date(now.getFullYear(), now.getMonth() + i, 1);
    const e = new Date(now.getFullYear(), now.getMonth() + i + 1, 1);
    const value = opps.filter(o => engine.openStages().indexOf(o.stage) >= 0 && inRange(o.expectedCloseDate, s, e))
      .reduce((t, o) => t + (engine.baseAmount(o) * health(o).score) / 100, 0);
    out.push({ label: `${s.toLocaleDateString(undefined, { month: 'long', year: 'numeric' })} (AI projection)`, short: s.toLocaleDateString(undefined, { month: 'short' }), value: Math.round(value) });
  }
  return out;
}

/* ------------------------------------------------------------------ briefing */

export type Tone = 'good' | 'warn' | 'bad' | 'info';
export interface IBriefItem { tone: Tone; icon: string; text: string; page?: string }

export function briefing(engine: Engine, sc: IScope, fc: IForecast, health: (o: IRecord) => IDealHealth): IBriefItem[] {
  const money = (n: number): string => fmtMoney(n, engine.settings.currency);
  const out: IBriefItem[] = [];

  // 1. Quarter outlook
  if (fc.quota > 0) {
    const pct = Math.round((fc.predicted / fc.quota) * 100);
    out.push({
      tone: pct >= 100 ? 'good' : pct >= 80 ? 'warn' : 'bad', icon: 'Trophy2', page: 'forecast',
      text: `${money(fc.won)} won in ${sc.quarter.label} (${Math.round((fc.won / fc.quota) * 100)}% of quota), ${fc.daysLeft} days left. AI expects about ${money(fc.predicted)}, ${pct}% of quota${pct < 100 ? `: ${money(fc.quota - fc.predicted)} short` : ''}.`
    });
  } else if (fc.deals || fc.won) {
    out.push({ tone: 'info', icon: 'Trophy2', page: 'forecast', text: `${money(fc.won)} won in ${sc.quarter.label}, ${fc.daysLeft} days left. AI expects about ${money(fc.predicted)} by quarter end from ${plural(fc.deals, 'deal')} still due.` });
  }

  // 2. Deals at risk
  const open = sc.opps.filter(o => engine.openStages().indexOf(o.stage) >= 0);
  const atRisk = open.filter(o => health(o).label === 'At risk');
  if (atRisk.length) {
    const reasons: Record<string, number> = {};
    atRisk.forEach(o => { const k = health(o).kinds[0]; if (k) { reasons[k] = (reasons[k] || 0) + 1; } });
    const top = Object.keys(reasons).sort((a, b) => reasons[b] - reasons[a])[0];
    const why = top ? ` Main reason: ${top} (${reasons[top]} of ${atRisk.length}).` : '';
    out.push({ tone: 'bad', icon: 'Warning', page: 'entity:opportunities', text: `${plural(atRisk.length, 'deal')} worth ${money(atRisk.reduce((t, o) => t + engine.baseAmount(o), 0))} ${atRisk.length === 1 ? 'is' : 'are'} at risk.${why}` });
  }

  // 3. Lead flow compared with the same point last quarter
  const qLeads = sc.leads.filter(l => inRange(l.created, sc.quarter.start, sc.quarter.end)).length;
  const pq = sc.prevQuarter;
  const pqSameTime = sc.leads.filter(l => inRange(l.created, pq.start, new Date(pq.start.getTime() + (pq.end.getTime() - pq.start.getTime()) * fc.elapsed))).length;
  if (pqSameTime >= 3 && fc.elapsed > 0.1) {
    const d = Math.round(((qLeads - pqSameTime) / pqSameTime) * 100);
    if (Math.abs(d) >= 15) {
      out.push({ tone: d > 0 ? 'good' : 'warn', icon: d > 0 ? 'Market' : 'MarketDown', page: 'entity:leads', text: `New leads are ${Math.abs(d)}% ${d > 0 ? 'ahead of' : 'behind'} last quarter at the same point (${qLeads} vs ${pqSameTime}).` });
    }
  }

  // 4. Speed to lead
  const waiting = sc.leads.filter(l => l.status === 'New' && (daysSince(l.created) || 0) >= 1);
  if (waiting.length) { out.push({ tone: 'warn', icon: 'Clock', page: 'entity:leads', text: `${plural(waiting.length, 'new lead')} ${waiting.length === 1 ? 'has' : 'have'} waited more than a day for a first contact.` }); }

  // 5. Best converting source
  const bySrc: Record<string, { n: number; c: number }> = {};
  sc.leads.forEach(l => { const k = l.leadSource || 'Other'; bySrc[k] = bySrc[k] || { n: 0, c: 0 }; bySrc[k].n++; if (l.status === 'Converted') { bySrc[k].c++; } });
  const srcs = Object.keys(bySrc).filter(k => bySrc[k].n >= 3 && bySrc[k].c > 0).sort((a, b) => bySrc[b].c / bySrc[b].n - bySrc[a].c / bySrc[a].n);
  if (srcs.length >= 2) { const b = bySrc[srcs[0]]; out.push({ tone: 'info', icon: 'Lightbulb', page: 'reports', text: `${srcs[0]} leads convert best: ${Math.round((b.c / b.n) * 100)}% become customers.` }); }

  // 6. Service and money
  const breached = sc.cases.filter(c => slaBreached(c));
  if (breached.length) { out.push({ tone: 'bad', icon: 'Headset', page: 'entity:cases', text: `${plural(breached.length, 'support case')} past the SLA.` }); }
  const overdue = sc.invoices.filter(i => i.invoiceStatus === 'Overdue');
  if (overdue.length) { out.push({ tone: 'warn', icon: 'Money', page: 'entity:invoices', text: `${plural(overdue.length, 'invoice')} overdue, ${money(overdue.reduce((t, i) => t + engine.toBase((i.grandTotal || 0) - (i.amountPaid || 0), i.currency), 0))} outstanding.` }); }
  const renew = sc.subs.filter(x => (x.subStatus === 'Active' || x.subStatus === 'Pending renewal') && (daysUntil(x.renewalDate) ?? 999) <= 30);
  if (renew.length) { out.push({ tone: 'info', icon: 'RecurringEvent', page: 'entity:subscriptions', text: `${plural(renew.length, 'renewal')} in the next 30 days, MRR ${money(renew.reduce((t, x) => t + engine.toBase(x.mrr, x.currency), 0))}.` }); }

  if (!out.length) { out.push({ tone: 'good', icon: 'CheckMark', text: 'Nothing urgent. Pipeline, leads and service all look on track.' }); }
  return out.slice(0, 6);
}

/* ------------------------------------------------------------------ next best actions */

export interface IAction {
  key: EntityKey;
  r: IRecord;
  title: string;
  why: string;
  /** Higher is more urgent. */
  rank: number;
  /** A follow-up email makes sense for this action. */
  canDraft: boolean;
}

export function nextActions(engine: Engine, sc: IScope, health: (o: IRecord) => IDealHealth, max: number = 8): IAction[] {
  const money = (n: number): string => fmtMoney(n, engine.settings.currency);
  const out: IAction[] = [];
  const big = Math.max(1, ...sc.opps.map(o => engine.baseAmount(o)));

  sc.leads.filter(l => l.status === 'New').forEach(l => {
    const age = daysSince(l.created) || 0;
    if (!l.owner) { out.push({ key: 'leads', r: l, title: `Assign ${l.title}`, why: 'New lead with no owner', rank: 70 + (l.leadScore || 0) / 5, canDraft: false }); }
    else if (age >= 1 || (l.leadScore || 0) >= 60) { out.push({ key: 'leads', r: l, title: `Contact ${l.title}${l.company ? ` (${l.company})` : ''}`, why: `${age ? `Waiting ${plural(age, 'day')}` : 'Hot lead'} · score ${l.leadScore || 0}`, rank: 55 + (l.leadScore || 0) / 3 + Math.min(age, 10), canDraft: true }); }
  });
  sc.opps.filter(o => engine.openStages().indexOf(o.stage) >= 0).forEach(o => {
    const h = health(o);
    if (h.label === 'Healthy' || !h.risks.length) { return; }
    const v = engine.baseAmount(o);
    out.push({ key: 'opportunities', r: o, title: `${h.label === 'At risk' ? 'Rescue' : 'Move forward'}: ${o.title}`, why: `${money(v)} · ${h.risks[0]} · ${h.score}% likely`, rank: (h.label === 'At risk' ? 50 : 35) + (v / big) * 30, canDraft: true });
  });
  sc.quotes.filter(q => q.signatureStatus === 'Viewed').forEach(q => out.push({ key: 'quotes', r: q, title: `Call about quote ${q.quoteNumber}`, why: 'Customer viewed it but has not signed', rank: 72, canDraft: true }));
  sc.quotes.filter(q => q.approvalStatus === 'Pending' && engine.role !== 'sales').forEach(q => out.push({ key: 'quotes', r: q, title: `Approve discount on ${q.quoteNumber}`, why: 'Waiting for approval', rank: 66, canDraft: false }));
  sc.cases.filter(c => slaBreached(c)).forEach(c => out.push({ key: 'cases', r: c, title: `Update ${c.caseNumber}: ${c.title}`, why: `SLA breached · ${c.priority || 'Normal'} priority`, rank: c.priority === 'Urgent' || c.priority === 'High' ? 90 : 75, canDraft: true }));
  sc.invoices.filter(i => i.invoiceStatus === 'Overdue').forEach(i => out.push({ key: 'invoices', r: i, title: `Chase invoice ${i.invoiceNumber}`, why: `Overdue since ${fmtDate(i.dueDate)}`, rank: 50 + Math.min(20, daysSince(i.dueDate) || 0), canDraft: true }));
  sc.subs.filter(x => x.subStatus === 'Pending renewal' || ((x.subStatus === 'Active') && (daysUntil(x.renewalDate) ?? 999) <= 21)).forEach(x => out.push({ key: 'subscriptions', r: x, title: `Confirm renewal: ${x.title}`, why: `Renews ${fmtDate(x.renewalDate)}`, rank: 48, canDraft: true }));
  sc.tasks.filter(t => isOverdue(t.dueDate)).forEach(t => out.push({ key: 'activities', r: t, title: t.title || 'Task', why: `Task overdue since ${fmtDate(t.dueDate)}`, rank: 60, canDraft: false }));

  return out.sort((a, b) => b.rank - a.rank).slice(0, max);
}

/* ------------------------------------------------------------------ ask your CRM */

export interface IAnswer {
  text: string;
  items: { key: EntityKey; r: IRecord; sub: string }[];
  /** The question was not understood. */
  unknown?: boolean;
}

export const ASK_EXAMPLES = [
  'Which deals are at risk?',
  'Deals closing this month',
  'Will we hit quota this quarter?',
  'Hot leads not contacted',
  'Overdue invoices',
  'Win rate by source'
];

const ENT: { key: EntityKey; re: RegExp }[] = [
  { key: 'leads', re: /\bleads?\b|prospects?/ },
  { key: 'opportunities', re: /\bdeals?\b|opportunit|pipeline/ },
  { key: 'cases', re: /\bcases?\b|tickets?|support/ },
  { key: 'invoices', re: /invoices?|unpaid|payments?/ },
  { key: 'quotes', re: /quotes?|proposals?/ },
  { key: 'demos', re: /\bdemos?\b/ },
  { key: 'subscriptions', re: /subscriptions?|renewals?|\bmrr\b/ },
  { key: 'activities', re: /\btasks?\b|to-?dos?/ }
];

export function askCrm(engine: Engine, sc: IScope, question: string, health: (o: IRecord) => IDealHealth, fc: IForecast): IAnswer {
  const q = question.toLowerCase().trim();
  const money = (n: number): string => fmtMoney(n, engine.settings.currency);
  const meId = engine.me?.id;
  if (!q) { return { text: '', items: [] }; }

  if (/quota|forecast|on track|will we (hit|make)|predict|end of (the )?quarter/.test(q)) {
    const text = fc.quota > 0
      ? `AI forecast for ${sc.quarter.label}: ${money(fc.predicted)} (${Math.round((fc.predicted / fc.quota) * 100)}% of the ${money(fc.quota)} quota). Won so far ${money(fc.won)}, commit ${money(fc.commit)}, best case ${money(fc.best)}, ${fc.daysLeft} days left.`
      : `AI forecast for ${sc.quarter.label}: ${money(fc.predicted)}. Won so far ${money(fc.won)}, commit ${money(fc.commit)}, best case ${money(fc.best)}, ${fc.daysLeft} days left. Set quotas in Settings › Teams, territories & quotas to compare against a target.`;
    const items = sc.opps.filter(o => engine.openStages().indexOf(o.stage) >= 0 && o.expectedCloseDate && new Date(o.expectedCloseDate).getTime() < sc.quarter.end.getTime())
      .sort((a, b) => engine.baseAmount(b) * health(b).score - engine.baseAmount(a) * health(a).score).slice(0, 8)
      .map(o => ({ key: 'opportunities' as EntityKey, r: o, sub: `${money(engine.baseAmount(o))} · ${health(o).score}% likely · close ${fmtDate(o.expectedCloseDate)}` }));
    return { text, items };
  }

  if (/win ?rate|convert|conversion/.test(q)) {
    const by = /source/.test(q) || /lead/.test(q);
    if (by) {
      const m: Record<string, { n: number; c: number }> = {};
      sc.leads.forEach(l => { const k = l.leadSource || 'Other'; m[k] = m[k] || { n: 0, c: 0 }; m[k].n++; if (l.status === 'Converted') { m[k].c++; } });
      const rows = Object.keys(m).sort((a, b) => m[b].c / m[b].n - m[a].c / m[a].n);
      return { text: rows.length ? `Lead conversion by source: ${rows.map(k => `${k} ${Math.round((m[k].c / m[k].n) * 100)}% (${m[k].c}/${m[k].n})`).join(', ')}.` : 'No leads yet.', items: [] };
    }
    const closed = sc.opps.filter(o => o.stage === WON || o.stage === LOST);
    const won = closed.filter(o => o.stage === WON).length;
    return { text: closed.length ? `Win rate: ${Math.round((won / closed.length) * 100)}% (${won} won of ${closed.length} closed deals).` : 'No closed deals yet.', items: [] };
  }

  const ent = ENT.find(e => e.re.test(q))?.key || (/at risk|stuck|closing|biggest/.test(q) ? 'opportunities' : /hot|uncontacted|not contacted/.test(q) ? 'leads' : undefined);
  if (!ent) {
    return { unknown: true, items: [], text: 'I can answer questions about leads, deals, quotes, cases, invoices, renewals and tasks, for example “deals closing this month” or “hot leads from LinkedIn”. For anything else, use Ask Copilot.' };
  }
  const pool: Record<string, IRecord[]> = { leads: sc.leads, opportunities: sc.opps, cases: sc.cases, invoices: sc.invoices, quotes: sc.quotes, demos: sc.demos, subscriptions: sc.subs, activities: sc.tasks };
  let rows = (pool[ent] || []).slice();
  const notes: string[] = [];
  const isOpenDeal = (o: IRecord): boolean => engine.openStages().indexOf(o.stage) >= 0;
  const dateField: Record<string, string> = { opportunities: 'expectedCloseDate', invoices: 'dueDate', subscriptions: 'renewalDate', demos: 'scheduledDate', activities: 'dueDate', cases: 'slaDue', quotes: 'validUntil', leads: 'created' };

  // Status filters
  if (/\bwon\b/.test(q) && ent === 'opportunities') { rows = rows.filter(o => o.stage === WON); notes.push('won'); }
  else if (/\blost\b/.test(q) && ent === 'opportunities') { rows = rows.filter(o => o.stage === LOST); notes.push('lost'); }
  else if (ent === 'opportunities') { rows = rows.filter(isOpenDeal); }
  if (/at risk|risky/.test(q) && ent === 'opportunities') { rows = rows.filter(o => health(o).label === 'At risk'); notes.push('at risk'); }
  else if (/attention|stuck|cold|slipping|watch/.test(q) && ent === 'opportunities') { rows = rows.filter(o => health(o).label !== 'Healthy'); notes.push('need attention'); }
  if (/overdue|late|past due|breach/.test(q)) {
    if (ent === 'cases') { rows = rows.filter(c => slaBreached(c)); } else if (ent === 'invoices') { rows = rows.filter(i => i.invoiceStatus === 'Overdue'); } else { rows = rows.filter(r => isOverdue(r[dateField[ent]])); }
    notes.push('overdue');
  }
  if (/unpaid|outstanding/.test(q) && ent === 'invoices') { rows = rows.filter(i => i.invoiceStatus === 'Sent' || i.invoiceStatus === 'Overdue'); notes.push('unpaid'); }
  if (ent === 'activities') { rows = rows.filter(t => !t.completed); }
  if (ent === 'cases' && !/closed|resolved/.test(q)) { rows = rows.filter(c => c.caseStatus !== 'Resolved' && c.caseStatus !== 'Closed'); }
  if (/hot|top|best/.test(q) && ent === 'leads') { rows = rows.filter(l => (l.leadScore || 0) >= 50 && l.status !== 'Converted' && l.status !== 'Unqualified'); notes.push('hot'); }
  if (/not contacted|uncontacted|\bnew\b|untouched/.test(q) && ent === 'leads') { rows = rows.filter(l => l.status === 'New'); notes.push('not contacted'); }
  if (/unassigned|no owner/.test(q)) { rows = rows.filter(r => !r.owner); notes.push('unassigned'); }
  if (/\b(my|mine|me)\b/.test(q) && meId) { rows = rows.filter(r => r.owner?.id === meId); notes.push('yours'); }
  if (/urgent|high priority/.test(q) && ent === 'cases') { rows = rows.filter(c => c.priority === 'Urgent' || c.priority === 'High'); notes.push('high priority'); }

  // Period
  const now = new Date();
  const period = /this week|next 7 days/.test(q) ? { s: new Date(now.getFullYear(), now.getMonth(), now.getDate() - ((now.getDay() + 6) % 7)), days: 7, name: 'this week' }
    : /next week/.test(q) ? { s: new Date(now.getFullYear(), now.getMonth(), now.getDate() - ((now.getDay() + 6) % 7) + 7), days: 7, name: 'next week' }
      : /this month/.test(q) ? { s: new Date(now.getFullYear(), now.getMonth(), 1), e: new Date(now.getFullYear(), now.getMonth() + 1, 1), name: 'this month' }
        : /next month/.test(q) ? { s: new Date(now.getFullYear(), now.getMonth() + 1, 1), e: new Date(now.getFullYear(), now.getMonth() + 2, 1), name: 'next month' }
          : /this quarter/.test(q) ? { s: sc.quarter.start, e: sc.quarter.end, name: sc.quarter.label }
            : /last quarter/.test(q) ? { s: sc.prevQuarter.start, e: sc.prevQuarter.end, name: sc.prevQuarter.label } : undefined;
  if (period) {
    const end = period.e || new Date(period.s.getTime() + (period.days || 7) * DAY);
    const f = /closing|close|due|renew|scheduled|upcoming/.test(q) ? dateField[ent] : /\bwon\b/.test(q) ? 'actualCloseDate' : ent === 'opportunities' && !/created|new/.test(q) ? 'expectedCloseDate' : ent === 'leads' || /created|new|came in/.test(q) ? 'created' : dateField[ent];
    rows = rows.filter(r => inRange(r[f], period.s, end));
    notes.push(period.name);
  }

  // "from LinkedIn", "in Germany", "in healthcare"
  const from = /\b(?:from|via|source)\s+([a-z0-9 .&-]{2,30}?)(?:\s+(?:this|next|last|over|above|with|that|which)\b|[?.!]|$)/.exec(q);
  if (from && ent === 'leads') {
    const v = from[1].trim();
    rows = rows.filter(l => String(l.leadSource || '').toLowerCase().indexOf(v) >= 0 || String(l.company || '').toLowerCase().indexOf(v) >= 0);
    notes.push(`from ${v}`);
  }
  const inX = /\bin\s+([a-z][a-z .&-]{2,30}?)(?:\s+(?:this|next|last|over|above|with)\b|[?.!]|$)/.exec(q);
  if (inX && !/in (the )?(next|this|last)/.test(inX[0])) {
    const v = inX[1].trim();
    const match = (r: IRecord): boolean => {
      const acc = r.accountRef ? engine.find('accounts', r.accountRef) : undefined;
      return [r.country, r.industry, r.stage, acc?.country, acc?.industry, acc?.title].some(x => String(x || '').toLowerCase().indexOf(v) >= 0);
    };
    const narrowed = rows.filter(match);
    if (narrowed.length || rows.length === 0) { rows = narrowed; notes.push(`in ${v}`); }
  }

  // Amount threshold: "over 10000", "above 5k"
  const amt = /(?:over|above|more than|greater than|>)\s*([\d.,]+)\s*(k|m|lakh|lakhs|cr|crore)?/.exec(q);
  if (amt) {
    const mult: Record<string, number> = { k: 1e3, m: 1e6, lakh: 1e5, lakhs: 1e5, cr: 1e7, crore: 1e7 };
    const n = Number(amt[1].replace(/,/g, '')) * (mult[amt[2] || ''] || 1);
    const val = (r: IRecord): number => (ent === 'opportunities' ? engine.baseAmount(r) : Number(r.grandTotal || r.mrr || r.amount) || 0);
    rows = rows.filter(r => val(r) > n);
    notes.push(`over ${money(n)}`);
  }

  // Ranking
  const value = (r: IRecord): number => (ent === 'opportunities' ? engine.baseAmount(r) : ent === 'invoices' ? engine.toBase((r.grandTotal || 0) - (r.amountPaid || 0), r.currency) : ent === 'subscriptions' ? engine.toBase(r.mrr, r.currency) : Number(r.grandTotal) || 0);
  if (ent === 'leads') { rows.sort((a, b) => (b.leadScore || 0) - (a.leadScore || 0)); }
  else if (ent === 'opportunities' && /at risk|risky|attention|stuck|cold|slipping|watch/.test(q)) { rows.sort((a, b) => health(a).score - health(b).score || value(b) - value(a)); }
  else if (/biggest|largest|top|highest/.test(q) || ['opportunities', 'invoices', 'subscriptions', 'quotes'].indexOf(ent) >= 0) { rows.sort((a, b) => value(b) - value(a)); }
  else { const f = dateField[ent]; rows.sort((a, b) => String(a[f] || '9').localeCompare(String(b[f] || '9'))); }

  const d = engine.def(ent);
  const n = rows.length;
  const total = rows.reduce((t, r) => t + value(r), 0);
  const label = `${n === 1 ? d.singular.toLowerCase() : d.plural.toLowerCase()}${notes.length ? ` (${notes.join(', ')})` : ''}`;
  const showMoney = ['opportunities', 'invoices', 'subscriptions', 'quotes'].indexOf(ent) >= 0 && total > 0;
  const text = n ? `${n} ${label}${showMoney ? `, ${ent === 'subscriptions' ? 'MRR ' : 'total '}${money(total)}` : ''}.` : `No ${label} found.`;

  const sub = (r: IRecord): string => {
    switch (ent) {
      case 'opportunities': { const h = health(r); return `${money(engine.baseAmount(r))} · ${r.stage} · ${h.score}% likely${h.risks[0] ? ` · ${h.risks[0]}` : ''}`; }
      case 'leads': return `${r.company || ''}${r.company ? ' · ' : ''}${r.leadSource || ''} · score ${r.leadScore || 0} · ${r.status}`;
      case 'invoices': return `${money(value(r))} due ${fmtDate(r.dueDate)} · ${r.invoiceStatus}`;
      case 'cases': return `${r.caseNumber} · ${r.priority || 'Normal'} · ${r.caseStatus}${r.slaDue ? ` · SLA ${fmtDate(r.slaDue, true)}` : ''}`;
      case 'quotes': return `${r.quoteNumber} · ${money(Number(r.grandTotal) || 0)} · ${r.quoteStatus || 'Draft'}`;
      case 'subscriptions': return `MRR ${money(value(r))} · renews ${fmtDate(r.renewalDate)}`;
      case 'demos': return `${fmtDate(r.scheduledDate, true)} · ${r.status}`;
      default: return r.dueDate ? `Due ${fmtDate(r.dueDate, true)}` : '';
    }
  };
  return { text, items: rows.slice(0, 8).map(r => ({ key: ent, r, sub: sub(r) })) };
}

/* ------------------------------------------------------------------ Copilot hand-off */

export function dashboardPrompt(engine: Engine, sc: IScope, fc: IForecast, brief: IBriefItem[], actions: IAction[], health: (o: IRecord) => IDealHealth): string {
  const money = (n: number): string => fmtMoney(n, engine.settings.currency);
  const open = sc.opps.filter(o => engine.openStages().indexOf(o.stage) >= 0).sort((a, b) => engine.baseAmount(b) - engine.baseAmount(a)).slice(0, 12);
  return [
    `You are my sales and service analyst. Using the CRM snapshot below for ${engine.settings.company.name || 'our company'}, give me: 1) a 5-bullet executive summary, 2) the three biggest risks to this quarter and what to do about each, 3) a prioritised plan for this week. Be specific and refer to deals by name.`,
    '', `Quarter ${sc.quarter.label}: won ${money(fc.won)}, AI forecast ${money(fc.predicted)}, commit ${money(fc.commit)}, best case ${money(fc.best)}${fc.quota ? `, quota ${money(fc.quota)}` : ''}, ${fc.daysLeft} days left.`,
    '', 'Signals:', ...brief.map(b => `- ${b.text}`),
    '', 'Largest open deals:', ...open.map(o => { const h = health(o); return `- ${o.title}: ${money(engine.baseAmount(o))}, ${o.stage}, close ${fmtDate(o.expectedCloseDate) || 'not set'}, win likelihood ${h.score}%${h.risks.length ? ` (${h.risks.slice(0, 2).join('; ')})` : ''}`; }),
    ...(actions.length ? ['', 'Suggested actions:', ...actions.map(a => `- ${a.title}: ${a.why}`)] : [])
  ].join('\n');
}
