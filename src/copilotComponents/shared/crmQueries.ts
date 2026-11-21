// What each Copilot tool shows, computed from the records the signed-in user can already see in the CRM.
// Pure functions over the Engine (no SharePoint calls), so they are unit-tested and behave the same as the
// dashboard: deal health, currency conversion, SLA rules and "reps see only their own records" are reused.
import type { Engine } from '../../webparts/brillientaCrm/services/Engine';
import { dealHealth, history, IDealHealth } from '../../webparts/brillientaCrm/services/DashboardAI';
import { relatedActivities } from '../../webparts/brillientaCrm/services/Insights';
import { CASE_PRIORITIES, EntityKey, IRecord, LOST, WON } from '../../webparts/brillientaCrm/models/model';
import { fmtDate, fmtMoney, inRange, quarterRange, slaBreached, firstResponseLate } from '../../webparts/brillientaCrm/components/utils';
import { HEALTH_FILTERS, LEAD_RATINGS, PIPELINE_PERIODS, PIPELINE_SCOPES } from './toolEnums';

export { HEALTH_FILTERS, LEAD_RATINGS, PIPELINE_PERIODS, PIPELINE_SCOPES };

const DAY = 86400000;
const HOUR = 3600000;

/* ------------------------------------------------------------------ input normalisation */

type Bag = Record<string, unknown>;
const bag = (v: unknown): Bag => (v && typeof v === 'object' ? (v as Bag) : {});
const str = (v: unknown, max: number = 200): string => (typeof v === 'string' ? v.trim().slice(0, max) : '');
const oneOf = <T extends string>(v: unknown, allowed: readonly T[], fallback: T): T => {
  const s = str(v).toLowerCase();
  return allowed.find(a => a.toLowerCase() === s) || fallback;
};
const int = (v: unknown, fallback: number, lo: number, hi: number): number => {
  const n = typeof v === 'number' ? v : typeof v === 'string' && v.trim() ? Number(v) : NaN;
  return isFinite(n) ? Math.max(lo, Math.min(hi, Math.round(n))) : fallback;
};
const bool = (v: unknown, fallback: boolean): boolean => (typeof v === 'boolean' ? v : v === 'true' ? true : v === 'false' ? false : fallback);
const lc = (s: unknown): string => String(s || '').trim().toLowerCase();

const daysUntil = (iso: string | undefined, now: number): number | undefined =>
  (iso && !isNaN(new Date(iso).getTime()) ? Math.ceil((new Date(iso).getTime() - now) / DAY) : undefined);
const startOfDay = (now: number): number => { const d = new Date(now); d.setHours(0, 0, 0, 0); return d.getTime(); };

/** Records of `key` the current user may see. Returns [] for features the licence plan does not include. */
const rows = (engine: Engine, key: EntityKey): IRecord[] => (engine.hasFeature(key) ? engine.visible(key) : []);
const ownedBy = (r: IRecord, ids: number[]): boolean => !!r.owner && ids.indexOf(r.owner.id) >= 0;
const money = (engine: Engine, n: number): string => fmtMoney(n, engine.settings.currency);
const accountName = (engine: Engine, id?: number): string => (id ? engine.display('accounts', id) : '');

/** Subscription statuses that still renew. */
const ACTIVE_SUBS = ['Active', 'Pending renewal', 'Trial'];
/** Invoice statuses that can be overdue (Draft is not sent yet; Paid and Void are settled). */
const UNPAID = ['Sent', 'Overdue'];
const balance = (i: IRecord): number => (Number(i.grandTotal) || 0) - (Number(i.amountPaid) || 0);
const isOpenCase = (c: IRecord): boolean => c.caseStatus !== 'Resolved' && c.caseStatus !== 'Closed';

export interface IRow {
  key: EntityKey;
  id: number;
  title: string;
}

/* ------------------------------------------------------------------ GetMyPipeline */


export interface IPipelineParams {
  scope: typeof PIPELINE_SCOPES[number];
  period: typeof PIPELINE_PERIODS[number];
  health: typeof HEALTH_FILTERS[number];
  minAmount: number;
}

export function normalizePipeline(input: unknown): IPipelineParams {
  const b = bag(input);
  return {
    scope: oneOf(b.scope, PIPELINE_SCOPES, 'mine'),
    period: oneOf(b.period, PIPELINE_PERIODS, 'thisQuarter'),
    health: oneOf(b.health, HEALTH_FILTERS, 'any'),
    minAmount: int(b.minAmount, 0, 0, 1e12)
  };
}

export interface IDealRow extends IRow {
  account: string;
  stage: string;
  amount: number;
  amountText: string;
  close?: string;
  closeText: string;
  owner: string;
  health: IDealHealth;
  nextStep: string;
}

export interface IPipelineView {
  params: IPipelineParams;
  /** Human description of the filter, e.g. "My deals closing in Q3 2026". */
  label: string;
  /** Set when the requested scope had to be narrowed (e.g. no team). */
  note?: string;
  deals: IDealRow[];
  total: number;
  weighted: number;
  atRisk: number;
}

export function pipelineView(engine: Engine, p: IPipelineParams, now: number = Date.now()): IPipelineView {
  const open = engine.openStages();
  let deals = rows(engine, 'opportunities').filter(o => open.indexOf(o.stage) >= 0 && o.stage !== WON && o.stage !== LOST);
  let note: string | undefined;
  let who = 'All';
  const me = engine.me?.id;
  if (p.scope === 'mine' || (p.scope === 'team' && !engine.teamMemberIds().length)) {
    if (p.scope === 'team') { note = 'You are not in a sales team, so only your own deals are shown.'; }
    deals = me ? deals.filter(o => ownedBy(o, [me])) : [];
    who = 'My';
  } else if (p.scope === 'team') {
    const ids = engine.teamMemberIds();
    deals = deals.filter(o => ownedBy(o, ids));
    who = 'My team’s';
  }

  const fy = engine.settings.fiscalYearStartMonth || 1;
  const ref = new Date(now);
  let when = '';
  if (p.period === 'thisQuarter' || p.period === 'nextQuarter') {
    const q = quarterRange(p.period === 'thisQuarter' ? 0 : 1, fy, ref);
    deals = deals.filter(o => inRange(o.expectedCloseDate, q.start, q.end));
    when = ` closing in ${q.label}`;
  } else if (p.period === 'thisMonth') {
    deals = deals.filter(o => { const d = o.expectedCloseDate ? new Date(o.expectedCloseDate) : undefined; return !!d && d.getFullYear() === ref.getFullYear() && d.getMonth() === ref.getMonth(); });
    when = ' closing this month';
  } else if (p.period === 'overdue') {
    deals = deals.filter(o => !!o.expectedCloseDate && new Date(o.expectedCloseDate).getTime() < startOfDay(now));
    when = ' past their close date';
  }
  if (p.minAmount > 0) { deals = deals.filter(o => engine.baseAmount(o) >= p.minAmount); }

  const hist = history(engine);
  let out: IDealRow[] = deals.map(o => {
    const amount = engine.baseAmount(o);
    return {
      key: 'opportunities' as EntityKey, id: o.id!, title: o.title || `#${o.id}`, account: accountName(engine, o.accountRef) || o.company || '',
      stage: o.stage || '', amount, amountText: money(engine, amount), close: o.expectedCloseDate, closeText: fmtDate(o.expectedCloseDate) || 'No close date',
      owner: o.owner?.title || '', health: dealHealth(engine, o, hist), nextStep: o.nextStep || ''
    };
  });
  const wanted = p.health === 'atRisk' ? 'At risk' : p.health === 'watch' ? 'Watch' : p.health === 'healthy' ? 'Healthy' : '';
  if (wanted) { out = out.filter(d => d.health.label === wanted); }
  out.sort((a, b) => String(a.close || '9999').localeCompare(String(b.close || '9999')) || b.amount - a.amount);

  const healthWord = wanted ? ` (${wanted.toLowerCase()})` : '';
  return {
    params: p, note, deals: out,
    label: `${who} open deals${when}${healthWord}${p.minAmount ? ` over ${money(engine, p.minAmount)}` : ''}`,
    total: out.reduce((t, d) => t + d.amount, 0),
    weighted: out.reduce((t, d) => t + d.amount * d.health.score / 100, 0),
    atRisk: out.filter(d => d.health.label === 'At risk').length
  };
}

export function pipelineSummary(engine: Engine, v: IPipelineView): string {
  if (!v.deals.length) { return `${v.label}: none found.`; }
  const top = v.deals.slice(0, 10).map(d => `- ${d.title}${d.account ? ` (${d.account})` : ''}: ${d.amountText}, ${d.stage}, closes ${d.closeText}, ${d.health.score}% likely${d.health.risks[0] ? `, risk: ${d.health.risks[0]}` : ''}`);
  return [`${v.label}: ${v.deals.length} deal(s), ${money(engine, v.total)} total, ${money(engine, v.weighted)} weighted, ${v.atRisk} at risk.`, ...top].join('\n');
}

/* ------------------------------------------------------------------ GetAccountSummary */

export interface IAccountParams { accountName: string }

export function normalizeAccount(input: unknown): IAccountParams {
  return { accountName: str(bag(input).accountName, 120) };
}

/** Exact name first, then names that contain the text, then all words in any order. */
export function findAccounts(engine: Engine, name: string, max: number = 6): IRecord[] {
  const q = lc(name);
  if (!q) { return []; }
  const all = rows(engine, 'accounts');
  const exact = all.filter(a => lc(a.title) === q);
  if (exact.length) { return exact.slice(0, max); }
  const words = q.split(/\s+/).filter(Boolean);
  const contains = all.filter(a => lc(a.title).indexOf(q) >= 0);
  const allWords = all.filter(a => contains.indexOf(a) < 0 && words.every(w => lc(a.title).indexOf(w) >= 0));
  return [...contains, ...allWords].slice(0, max);
}

export interface IAccountView {
  account: IRecord;
  type: string;
  owner: string;
  contacts: number;
  deals: IDealRow[];
  openPipeline: number;
  cases: ICaseRow[];
  subscriptions: IRenewalRow[];
  arr: number;
  invoices: IInvoiceRow[];
  outstanding: number;
  activities: { id: number; title: string; type: string; when: string }[];
  /** False when the licence plan has no service module. */
  hasCases: boolean;
}

export function accountView(engine: Engine, account: IRecord, now: number = Date.now()): IAccountView {
  const id = account.id!;
  const hist = history(engine);
  const open = engine.openStages();
  const deals: IDealRow[] = rows(engine, 'opportunities').filter(o => o.accountRef === id && open.indexOf(o.stage) >= 0).map(o => {
    const amount = engine.baseAmount(o);
    return {
      key: 'opportunities' as EntityKey, id: o.id!, title: o.title || `#${o.id}`, account: account.title || '', stage: o.stage || '', amount,
      amountText: money(engine, amount), close: o.expectedCloseDate, closeText: fmtDate(o.expectedCloseDate) || 'No close date',
      owner: o.owner?.title || '', health: dealHealth(engine, o, hist), nextStep: o.nextStep || ''
    };
  }).sort((a, b) => b.amount - a.amount);
  const cases = caseRows(engine, rows(engine, 'cases').filter(c => c.accountRef === id && isOpenCase(c)), 24, now);
  const subscriptions = renewalRows(engine, rows(engine, 'subscriptions').filter(s => s.accountRef === id && ACTIVE_SUBS.indexOf(s.subStatus) >= 0), now);
  const invoices = invoiceRows(engine, rows(engine, 'invoices').filter(i => i.accountRef === id && balance(i) > 0.005 && UNPAID.indexOf(i.invoiceStatus) >= 0), now);
  const activities = relatedActivities(engine, 'accounts', account).slice(0, 5).map(a => ({
    id: a.id!, title: a.title || a.activityType || 'Activity', type: a.activityType || '', when: fmtDate(a.dueDate || a.created)
  }));
  return {
    account, type: account.accountType || '', owner: account.owner?.title || '',
    contacts: rows(engine, 'contacts').filter(c => c.accountRef === id).length,
    deals, openPipeline: deals.reduce((t, d) => t + d.amount, 0),
    cases, subscriptions, arr: subscriptions.reduce((t, s) => t + s.arr, 0),
    invoices, outstanding: invoices.reduce((t, i) => t + i.balance, 0),
    activities, hasCases: engine.hasFeature('cases')
  };
}

export function accountSummary(engine: Engine, v: IAccountView): string {
  const lines = [
    `Account ${v.account.title}${v.type ? ` (${v.type})` : ''}${v.owner ? `, owner ${v.owner}` : ''}, ${v.contacts} contact(s).`,
    `Open deals: ${v.deals.length}, ${money(engine, v.openPipeline)}.${v.deals.slice(0, 5).map(d => ` ${d.title}: ${d.amountText}, ${d.stage}, ${d.health.score}% likely.`).join('')}`,
    v.hasCases ? `Open cases: ${v.cases.length}${v.cases.filter(c => c.state === 'breached').length ? `, ${v.cases.filter(c => c.state === 'breached').length} past SLA` : ''}.` : '',
    `Active subscriptions: ${v.subscriptions.length}, ${money(engine, v.arr)} ARR${v.subscriptions[0] ? `, next renewal ${v.subscriptions[0].renewalText}` : ''}.`,
    `Unpaid invoices: ${v.invoices.length}, ${money(engine, v.outstanding)} outstanding.`,
    v.activities.length ? `Recent activity: ${v.activities.map(a => `${a.when} ${a.type} "${a.title}"`).join('; ')}.` : 'No recent activity.'
  ];
  return lines.filter(Boolean).join('\n');
}

/* ------------------------------------------------------------------ GetRenewals */

export interface IRenewalParams { days: number; onlyMine: boolean; includeAutoRenew: boolean }

export function normalizeRenewals(input: unknown): IRenewalParams {
  const b = bag(input);
  return { days: int(b.days, 60, 1, 365), onlyMine: bool(b.onlyMine, false), includeAutoRenew: bool(b.includeAutoRenew, true) };
}

export interface IRenewalRow extends IRow {
  account: string;
  status: string;
  renewal?: string;
  renewalText: string;
  daysLeft?: number;
  arr: number;
  arrText: string;
  autoRenew: boolean;
  seats: number;
}

function renewalRows(engine: Engine, subs: IRecord[], now: number): IRenewalRow[] {
  return subs.map(s => {
    const arr = engine.toBase(Number(s.arr) || (Number(s.mrr) || 0) * 12, s.currency);
    return {
      key: 'subscriptions' as EntityKey, id: s.id!, title: s.title || `#${s.id}`, account: accountName(engine, s.accountRef), status: s.subStatus || '',
      renewal: s.renewalDate, renewalText: fmtDate(s.renewalDate) || 'No renewal date', daysLeft: daysUntil(s.renewalDate, now),
      arr, arrText: money(engine, arr), autoRenew: s.autoRenew !== false, seats: Number(s.quantity) || 0
    };
  }).sort((a, b) => String(a.renewal || '9999').localeCompare(String(b.renewal || '9999')));
}

export interface IRenewalView { params: IRenewalParams; label: string; renewals: IRenewalRow[]; arr: number; atRisk: number }

export function renewalView(engine: Engine, p: IRenewalParams, now: number = Date.now()): IRenewalView {
  const me = engine.me?.id;
  const subs = rows(engine, 'subscriptions').filter(s => {
    if (ACTIVE_SUBS.indexOf(s.subStatus) < 0) { return false; }
    const d = daysUntil(s.renewalDate, now);
    // Renewals already past their date but still pending are included: they need attention most.
    if (d === undefined || d > p.days || (d < 0 && s.subStatus !== 'Pending renewal')) { return false; }
    if (!p.includeAutoRenew && s.autoRenew !== false) { return false; }
    return !p.onlyMine || (!!me && ownedBy(s, [me]));
  });
  const renewals = renewalRows(engine, subs, now);
  return {
    params: p, renewals,
    label: `${p.onlyMine ? 'My subscriptions' : 'Subscriptions'} renewing in the next ${p.days} days${p.includeAutoRenew ? '' : ' (manual renewal only)'}`,
    arr: renewals.reduce((t, r) => t + r.arr, 0),
    atRisk: renewals.filter(r => !r.autoRenew || (r.daysLeft !== undefined && r.daysLeft < 0)).length
  };
}

export function renewalSummary(engine: Engine, v: IRenewalView): string {
  if (!v.renewals.length) { return `${v.label}: none.`; }
  return [`${v.label}: ${v.renewals.length}, ${money(engine, v.arr)} ARR.`,
    ...v.renewals.slice(0, 15).map(r => `- ${r.title}${r.account ? ` (${r.account})` : ''}: renews ${r.renewalText}, ${r.arrText} ARR, ${r.autoRenew ? 'auto-renew' : 'manual renewal'}`)].join('\n');
}

/* ------------------------------------------------------------------ GetOverdueInvoices */

export interface IInvoiceParams { minDaysOverdue: number; accountName: string; onlyMine: boolean }

export function normalizeInvoices(input: unknown): IInvoiceParams {
  const b = bag(input);
  return { minDaysOverdue: int(b.minDaysOverdue, 1, 0, 3650), accountName: str(b.accountName, 120), onlyMine: bool(b.onlyMine, false) };
}

export interface IInvoiceRow extends IRow {
  number: string;
  account: string;
  due?: string;
  dueText: string;
  daysOverdue: number;
  balance: number;
  balanceText: string;
  paymentLink: string;
}

function invoiceRows(engine: Engine, invoices: IRecord[], now: number): IInvoiceRow[] {
  const today = startOfDay(now);
  return invoices.map(i => {
    const b = engine.toBase(balance(i), i.currency);
    const due = i.dueDate ? new Date(i.dueDate).getTime() : NaN;
    return {
      key: 'invoices' as EntityKey, id: i.id!, title: i.title || i.invoiceNumber || `#${i.id}`, number: i.invoiceNumber || '', account: accountName(engine, i.accountRef),
      due: i.dueDate, dueText: fmtDate(i.dueDate) || 'No due date', daysOverdue: isFinite(due) ? Math.max(0, Math.floor((today - startOfDay(due)) / DAY)) : 0,
      balance: b, balanceText: money(engine, b), paymentLink: i.paymentLink || ''
    };
  }).sort((a, b) => b.daysOverdue - a.daysOverdue || b.balance - a.balance);
}

export interface IInvoiceView { params: IInvoiceParams; label: string; invoices: IInvoiceRow[]; outstanding: number }

export function invoiceView(engine: Engine, p: IInvoiceParams, now: number = Date.now()): IInvoiceView {
  const me = engine.me?.id;
  const q = lc(p.accountName);
  const candidates = rows(engine, 'invoices').filter(i => UNPAID.indexOf(i.invoiceStatus) >= 0 && balance(i) > 0.005 && !!i.dueDate
    && (!q || lc(accountName(engine, i.accountRef)).indexOf(q) >= 0) && (!p.onlyMine || (!!me && ownedBy(i, [me]))));
  const invoices = invoiceRows(engine, candidates, now).filter(i => i.daysOverdue >= Math.max(1, p.minDaysOverdue));
  const who = p.accountName ? ` for "${p.accountName}"` : '';
  return {
    params: p, invoices,
    label: `${p.onlyMine ? 'My overdue' : 'Overdue'} invoices${who}${p.minDaysOverdue > 1 ? ` (${p.minDaysOverdue}+ days)` : ''}`,
    outstanding: invoices.reduce((t, i) => t + i.balance, 0)
  };
}

export function invoiceSummary(engine: Engine, v: IInvoiceView): string {
  if (!v.invoices.length) { return `${v.label}: none.`; }
  return [`${v.label}: ${v.invoices.length}, ${money(engine, v.outstanding)} outstanding.`,
    ...v.invoices.slice(0, 15).map(i => `- ${i.number || i.title}${i.account ? ` (${i.account})` : ''}: ${i.balanceText}, ${i.daysOverdue} day(s) overdue, due ${i.dueText}`)].join('\n');
}

/* ------------------------------------------------------------------ GetSlaRiskCases */

export interface ICaseParams { windowHours: number; priority: string; onlyMine: boolean }

export function normalizeCases(input: unknown): ICaseParams {
  const b = bag(input);
  const pr = CASE_PRIORITIES.find(x => lc(x) === lc(b.priority)) || '';
  return { windowHours: int(b.windowHours, 24, 1, 720), priority: pr, onlyMine: bool(b.onlyMine, false) };
}

export type CaseState = 'breached' | 'firstResponseLate' | 'dueSoon' | 'paused' | 'ok';

export interface ICaseRow extends IRow {
  number: string;
  account: string;
  priority: string;
  status: string;
  owner: string;
  slaDue?: string;
  slaText: string;
  state: CaseState;
  /** Hours until (positive) or since (negative) the SLA deadline. */
  hoursLeft?: number;
}

function caseRows(engine: Engine, cases: IRecord[], windowHours: number, now: number): ICaseRow[] {
  return cases.map(c => {
    const due = c.slaDue ? new Date(c.slaDue).getTime() : NaN;
    const hoursLeft = isFinite(due) ? Math.round((due - now) / HOUR) : undefined;
    const state: CaseState = c.slaPausedAt ? 'paused' : slaBreached(c) ? 'breached' : firstResponseLate(c) ? 'firstResponseLate'
      : hoursLeft !== undefined && hoursLeft <= windowHours ? 'dueSoon' : 'ok';
    return {
      key: 'cases' as EntityKey, id: c.id!, title: c.title || `#${c.id}`, number: c.caseNumber || '', account: accountName(engine, c.accountRef),
      priority: c.priority || 'Normal', status: c.caseStatus || '', owner: c.owner?.title || '', slaDue: c.slaDue, slaText: fmtDate(c.slaDue, true) || 'No SLA',
      state, hoursLeft
    };
  });
}

const STATE_ORDER: CaseState[] = ['breached', 'firstResponseLate', 'dueSoon', 'paused', 'ok'];

export interface ICaseView { params: ICaseParams; label: string; cases: ICaseRow[]; breached: number; dueSoon: number; available: boolean }

export function caseView(engine: Engine, p: ICaseParams, now: number = Date.now()): ICaseView {
  const available = engine.hasFeature('cases');
  const me = engine.me?.id;
  const open = rows(engine, 'cases').filter(c => isOpenCase(c) && (!p.priority || c.priority === p.priority) && (!p.onlyMine || (!!me && ownedBy(c, [me]))));
  const cases = caseRows(engine, open, p.windowHours, now)
    .filter(c => c.state === 'breached' || c.state === 'firstResponseLate' || c.state === 'dueSoon')
    .sort((a, b) => STATE_ORDER.indexOf(a.state) - STATE_ORDER.indexOf(b.state) || (a.hoursLeft ?? 1e9) - (b.hoursLeft ?? 1e9));
  return {
    params: p, cases, available,
    label: `${p.onlyMine ? 'My open' : 'Open'} ${p.priority ? `${p.priority.toLowerCase()}-priority ` : ''}cases past SLA or due within ${p.windowHours} h`,
    breached: cases.filter(c => c.state === 'breached').length,
    dueSoon: cases.filter(c => c.state !== 'breached').length
  };
}

export function caseSummary(v: ICaseView): string {
  if (!v.available) { return 'Cases are part of the Brillienta CRM Enterprise plan, which this tenant does not have.'; }
  if (!v.cases.length) { return `${v.label}: none.`; }
  const word: Record<CaseState, string> = { breached: 'SLA breached', firstResponseLate: 'first response late', dueSoon: 'due soon', paused: 'paused', ok: 'on track' };
  return [`${v.label}: ${v.breached} breached, ${v.dueSoon} at risk.`,
    ...v.cases.slice(0, 15).map(c => `- ${c.number} ${c.title}${c.account ? ` (${c.account})` : ''}: ${c.priority}, ${word[c.state]}, SLA ${c.slaText}${c.owner ? `, owner ${c.owner}` : ''}`)].join('\n');
}

/* ------------------------------------------------------------------ CreateLead */

export interface ILeadDraft {
  firstName: string;
  lastName: string;
  email: string;
  phone: string;
  company: string;
  jobTitle: string;
  country: string;
  leadSource: string;
  rating: 'Hot' | 'Warm' | 'Cold';
  description: string;
}

/** Prompt values only prefill the form; nothing is saved until the user presses Save. */
export function normalizeLeadDraft(input: unknown, leadSources: string[]): ILeadDraft {
  const b = bag(input);
  const source = str(b.leadSource, 80);
  return {
    firstName: str(b.firstName, 80), lastName: str(b.lastName, 80), email: str(b.email, 200).toLowerCase(), phone: str(b.phone, 40),
    company: str(b.company, 120), jobTitle: str(b.jobTitle, 120), country: str(b.country, 80),
    leadSource: leadSources.find(s => lc(s) === lc(source)) || '',
    rating: oneOf(b.rating, LEAD_RATINGS, 'Warm'),
    description: str(b.description, 2000)
  };
}

const EMAIL = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

/** Field errors, keyed by field. Empty when the draft can be saved. */
export function validateLeadDraft(d: ILeadDraft): Partial<Record<keyof ILeadDraft, string>> {
  const e: Partial<Record<keyof ILeadDraft, string>> = {};
  if (!d.firstName.trim()) { e.firstName = 'First name is required.'; }
  if (d.email && !EMAIL.test(d.email)) { e.email = 'Enter a valid email address.'; }
  if (!d.email && !d.phone && !d.company) { e.email = 'Add an email, phone or company so the lead can be followed up.'; }
  return e;
}

export function leadRecord(d: ILeadDraft): IRecord {
  const r: IRecord = { firstName: d.firstName, lastName: d.lastName, rating: d.rating, status: 'New' };
  (['email', 'phone', 'company', 'jobTitle', 'country', 'leadSource', 'description'] as (keyof ILeadDraft)[]).forEach(k => { if (d[k]) { r[k] = d[k]; } });
  return r;
}
