/* eslint-disable @typescript-eslint/no-explicit-any */
import { DataService } from '../../webparts/brillientaCrm/services/DataService';
import { Engine } from '../../webparts/brillientaCrm/services/Engine';
import { GraphService } from '../../webparts/brillientaCrm/services/Integrations';
import type { ILicenseInfo } from '../../webparts/brillientaCrm/models/license';
import type { CrmData } from '../../webparts/brillientaCrm/models/model';
import {
  accountView, caseView, findAccounts, invoiceView, leadRecord, normalizeAccount, normalizeCases, normalizeInvoices,
  normalizeLeadDraft, normalizePipeline, normalizeRenewals, pipelineView, renewalView, validateLeadDraft
} from './crmQueries';

// Fixed "now": Friday 15 May 2026, 12:00 local time (calendar fiscal year, so Q2 = April to June).
const NOW = new Date(2026, 4, 15, 12, 0, 0).getTime();
const day = (offset: number): string => new Date(NOW + offset * 86400000).toISOString();
// The CRM's own helpers (SLA, overdue) read the clock, so pin it for every test.
beforeAll(() => { jest.useFakeTimers({ now: NOW }); });
afterAll(() => { jest.useRealTimers(); });
const ME = { id: 1, title: 'Asha Rao', email: 'asha@contoso.com' };
const OTHER = { id: 2, title: 'Ben Ortiz', email: 'ben@contoso.com' };

const enterprise: ILicenseInfo = { status: 'active', plan: 'enterprise', features: 'enterprise', canWrite: true, message: '' };
const standard: ILicenseInfo = { status: 'active', plan: 'standard', features: 'standard', canWrite: true, message: '' };

function makeEngine(data: CrmData, license: ILicenseInfo = enterprise, role: 'sales' | 'manager' = 'sales'): Engine {
  const engine = new Engine(new DataService({} as any, 'https://contoso.sharepoint.com/sites/crm'), new GraphService(undefined), 'tenant', () => undefined);
  engine.setSession(ME, role, license);
  engine.data = data;
  return engine;
}

const accounts = [
  { id: 10, title: 'Contoso Ltd', accountType: 'Customer', owner: ME },
  { id: 11, title: 'Contoso Pharma', accountType: 'Prospect', owner: OTHER },
  { id: 12, title: 'Fabrikam', accountType: 'Customer', owner: ME },
  { id: 13, title: 'Northwind Traders', accountType: 'Customer', owner: OTHER }
];

const opportunities = [
  // mine, this quarter
  { id: 100, title: 'Contoso HR portal', accountRef: 10, contactRef: 5, stage: 'Proposal', amount: 20000, probability: 50, expectedCloseDate: day(20), owner: ME, created: day(-10), modified: day(-1) },
  // mine, this quarter, no contact, close date close: lower health
  { id: 101, title: 'Fabrikam renewal', accountRef: 12, stage: 'Negotiation', amount: 5000, probability: 75, expectedCloseDate: day(5), owner: ME, created: day(-40), modified: day(-35) },
  // mine, next quarter
  { id: 102, title: 'Contoso expansion', accountRef: 10, stage: 'Prospecting', amount: 50000, probability: 10, expectedCloseDate: day(60), owner: ME, created: day(-5), modified: day(-2) },
  // mine, close date passed last quarter
  { id: 103, title: 'Old deal', accountRef: 12, stage: 'Qualification', amount: 3000, probability: 25, expectedCloseDate: day(-50), owner: ME, created: day(-90), modified: day(-60) },
  // someone else's
  { id: 104, title: 'Northwind CRM', accountRef: 13, stage: 'Proposal', amount: 70000, probability: 50, expectedCloseDate: day(10), owner: OTHER, created: day(-10), modified: day(-1) },
  // closed: never shown
  { id: 105, title: 'Won deal', accountRef: 10, stage: 'Closed Won', amount: 9000, expectedCloseDate: day(1), owner: ME },
  { id: 106, title: 'Lost deal', accountRef: 10, stage: 'Closed Lost', amount: 9000, expectedCloseDate: day(1), owner: ME }
];

describe('GetMyPipeline', () => {
  it('normalises prompt values and falls back to safe defaults', () => {
    expect(normalizePipeline(undefined)).toEqual({ scope: 'mine', period: 'thisQuarter', health: 'any', minAmount: 0 });
    expect(normalizePipeline({ scope: 'TEAM', period: 'nextquarter', health: 'atrisk', minAmount: '10000' }))
      .toEqual({ scope: 'team', period: 'nextQuarter', health: 'atRisk', minAmount: 10000 });
    expect(normalizePipeline({ scope: 'everyone', period: 'yesterday', minAmount: -5 })).toEqual({ scope: 'mine', period: 'thisQuarter', health: 'any', minAmount: 0 });
  });

  it('shows my open deals closing this quarter, soonest first, without closed deals', () => {
    const engine = makeEngine({ opportunities, accounts, activities: [], quotes: [] });
    const v = pipelineView(engine, normalizePipeline({}), NOW);
    expect(v.deals.map(d => d.id)).toEqual([101, 100]);
    expect(v.total).toBe(25000);
    expect(v.deals[1].account).toBe('Contoso Ltd');
    expect(v.label).toBe('My open deals closing in Q2 FY26');
    expect(v.weighted).toBeGreaterThan(0);
    expect(v.weighted).toBeLessThan(v.total);
  });

  it('filters by period, amount and scope', () => {
    const engine = makeEngine({ opportunities, accounts, activities: [], quotes: [] });
    expect(pipelineView(engine, normalizePipeline({ period: 'nextQuarter' }), NOW).deals.map(d => d.id)).toEqual([102]);
    expect(pipelineView(engine, normalizePipeline({ period: 'overdue' }), NOW).deals.map(d => d.id)).toEqual([103]);
    expect(pipelineView(engine, normalizePipeline({ period: 'all', minAmount: 10000 }), NOW).deals.map(d => d.id)).toEqual([100, 102]);
    expect(pipelineView(engine, normalizePipeline({ scope: 'all' }), NOW).deals.map(d => d.id)).toEqual([101, 104, 100]);
  });

  it('treats the quarter end as exclusive', () => {
    const july1 = new Date(2026, 6, 1, 0, 0, 0).toISOString();
    const engine = makeEngine({ opportunities: [{ id: 1, title: 'Next quarter', stage: 'Proposal', amount: 1, expectedCloseDate: july1, owner: ME }], activities: [], quotes: [] });
    expect(pipelineView(engine, normalizePipeline({}), NOW).deals).toEqual([]);
    expect(pipelineView(engine, normalizePipeline({ period: 'nextQuarter' }), NOW).deals.length).toBe(1);
  });

  it('falls back to my deals with a note when I am not in a team', () => {
    const engine = makeEngine({ opportunities, accounts, activities: [], quotes: [] });
    const v = pipelineView(engine, normalizePipeline({ scope: 'team' }), NOW);
    expect(v.note).toMatch(/not in a sales team/);
    expect(v.deals.every(d => d.owner === ME.title)).toBe(true);
  });

  it('uses the team when there is one', () => {
    const engine = makeEngine({ opportunities, accounts, activities: [], quotes: [] });
    engine.ds.settings = { ...engine.ds.settings, teams: [{ id: 't1', name: 'West', manager: ME, members: [OTHER], countries: [] } as any] };
    const v = pipelineView(engine, normalizePipeline({ scope: 'team', period: 'all' }), NOW);
    expect(v.note).toBeUndefined();
    expect(v.deals.map(d => d.id).sort()).toEqual([100, 101, 102, 103, 104]);
  });

  it('filters by deal health', () => {
    const engine = makeEngine({ opportunities, accounts, activities: [], quotes: [] });
    const all = pipelineView(engine, normalizePipeline({ period: 'all' }), NOW);
    const atRisk = pipelineView(engine, normalizePipeline({ period: 'all', health: 'atRisk' }), NOW);
    expect(atRisk.deals.length).toBe(all.deals.filter(d => d.health.label === 'At risk').length);
    expect(atRisk.deals.every(d => d.health.label === 'At risk')).toBe(true);
    expect(atRisk.atRisk).toBe(atRisk.deals.length);
  });

  it('respects "reps see only their own records"', () => {
    const engine = makeEngine({ opportunities, accounts, activities: [], quotes: [] });
    engine.ds.settings = { ...engine.ds.settings, restrictRepsToOwnRecords: true };
    expect(pipelineView(engine, normalizePipeline({ scope: 'all', period: 'all' }), NOW).deals.some(d => d.id === 104)).toBe(false);
  });
});

describe('GetAccountSummary', () => {
  it('finds accounts by exact name, then partial name, then all words', () => {
    const engine = makeEngine({ accounts });
    expect(findAccounts(engine, 'fabrikam').map(a => a.id)).toEqual([12]);
    expect(findAccounts(engine, 'Contoso').map(a => a.id)).toEqual([10, 11]);
    expect(findAccounts(engine, 'traders northwind').map(a => a.id)).toEqual([13]);
    expect(findAccounts(engine, '  ')).toEqual([]);
    expect(normalizeAccount({ accountName: '  Contoso  ' })).toEqual({ accountName: 'Contoso' });
  });

  it('pulls deals, cases, subscriptions, invoices and contacts for one account', () => {
    const engine = makeEngine({
      accounts, opportunities, quotes: [],
      contacts: [{ id: 5, title: 'Priya Shah', accountRef: 10 }, { id: 6, title: 'Other', accountRef: 12 }],
      cases: [
        { id: 300, caseNumber: 'CS-2026-0300', title: 'Login fails', accountRef: 10, caseStatus: 'In progress', priority: 'High', slaDue: day(-0.5) },
        { id: 301, caseNumber: 'CS-2026-0301', title: 'Closed one', accountRef: 10, caseStatus: 'Closed', slaDue: day(-5) }
      ],
      subscriptions: [{ id: 400, title: 'HR Suite', accountRef: 10, subStatus: 'Active', arr: 12000, renewalDate: day(30) }],
      invoices: [
        { id: 500, invoiceNumber: 'INV-2026-0500', accountRef: 10, invoiceStatus: 'Overdue', grandTotal: 1000, amountPaid: 400, dueDate: day(-12) },
        { id: 501, invoiceNumber: 'INV-2026-0501', accountRef: 10, invoiceStatus: 'Paid', grandTotal: 1000, amountPaid: 1000, dueDate: day(-40) }
      ],
      activities: [{ id: 600, title: 'Kick-off call', activityType: 'Call', relatedType: 'Opportunity', relatedId: 100, completed: true, dueDate: day(-1) }]
    });
    const v = accountView(engine, accounts[0], NOW);
    expect(v.contacts).toBe(1);
    expect(v.deals.map(d => d.id)).toEqual([102, 100]);
    expect(v.openPipeline).toBe(70000);
    expect(v.cases.map(c => [c.id, c.state])).toEqual([[300, 'breached']]);
    expect(v.arr).toBe(12000);
    expect(v.invoices.map(i => i.id)).toEqual([500]);
    expect(v.outstanding).toBe(600);
    expect(v.activities.map(a => a.title)).toEqual(['Kick-off call']);
  });

  it('hides cases on the Standard plan', () => {
    const engine = makeEngine({ accounts, cases: [{ id: 300, accountRef: 10, caseStatus: 'New', slaDue: day(-1) }] }, standard);
    const v = accountView(engine, accounts[0], NOW);
    expect(v.hasCases).toBe(false);
    expect(v.cases).toEqual([]);
  });
});

describe('GetRenewals', () => {
  const subscriptions = [
    { id: 1, title: 'In 10 days', accountRef: 10, subStatus: 'Active', arr: 1200, renewalDate: day(10), owner: ME },
    { id: 2, title: 'In 50 days, manual', accountRef: 12, subStatus: 'Active', mrr: 100, autoRenew: false, renewalDate: day(50), owner: OTHER },
    { id: 3, title: 'In 90 days', accountRef: 12, subStatus: 'Active', arr: 5000, renewalDate: day(90), owner: ME },
    { id: 4, title: 'Pending, date passed', accountRef: 13, subStatus: 'Pending renewal', arr: 800, renewalDate: day(-3), owner: ME },
    { id: 5, title: 'Active, date passed', accountRef: 13, subStatus: 'Active', arr: 800, renewalDate: day(-3), owner: ME },
    { id: 6, title: 'Cancelled', accountRef: 13, subStatus: 'Cancelled', arr: 800, renewalDate: day(5), owner: ME }
  ];

  it('lists renewals in the window, oldest first, with past-due pending renewals', () => {
    const engine = makeEngine({ subscriptions, accounts });
    const v = renewalView(engine, normalizeRenewals({}), NOW);
    expect(v.renewals.map(r => r.id)).toEqual([4, 1, 2]);
    expect(v.renewals[2].arr).toBe(1200); // from MRR × 12
    expect(v.arr).toBe(800 + 1200 + 1200);
    expect(v.atRisk).toBe(2); // past due + manual renewal
  });

  it('honours the window, "mine" and manual-only', () => {
    const engine = makeEngine({ subscriptions, accounts });
    expect(renewalView(engine, normalizeRenewals({ days: 120 }), NOW).renewals.map(r => r.id)).toEqual([4, 1, 2, 3]);
    expect(renewalView(engine, normalizeRenewals({ onlyMine: true }), NOW).renewals.map(r => r.id)).toEqual([4, 1]);
    expect(renewalView(engine, normalizeRenewals({ includeAutoRenew: false }), NOW).renewals.map(r => r.id)).toEqual([2]);
    expect(normalizeRenewals({ days: 9999 }).days).toBe(365);
  });
});

describe('GetOverdueInvoices', () => {
  const invoices = [
    { id: 1, invoiceNumber: 'INV-1', accountRef: 10, invoiceStatus: 'Overdue', grandTotal: 1000, amountPaid: 0, dueDate: day(-45), owner: ME },
    { id: 2, invoiceNumber: 'INV-2', accountRef: 12, invoiceStatus: 'Sent', grandTotal: 500, amountPaid: 100, dueDate: day(-3), owner: OTHER },
    { id: 3, invoiceNumber: 'INV-3', accountRef: 12, invoiceStatus: 'Sent', grandTotal: 500, dueDate: day(5), owner: ME },
    { id: 4, invoiceNumber: 'INV-4', accountRef: 12, invoiceStatus: 'Paid', grandTotal: 500, amountPaid: 500, dueDate: day(-30), owner: ME },
    { id: 5, invoiceNumber: 'INV-5', accountRef: 12, invoiceStatus: 'Draft', grandTotal: 500, dueDate: day(-30), owner: ME },
    { id: 6, invoiceNumber: 'INV-6', accountRef: 13, invoiceStatus: 'Overdue', grandTotal: 300, amountPaid: 300, dueDate: day(-30), owner: ME }
  ];

  it('lists unpaid invoices past due, most overdue first, with the open balance', () => {
    const engine = makeEngine({ invoices, accounts });
    const v = invoiceView(engine, normalizeInvoices({}), NOW);
    expect(v.invoices.map(i => i.id)).toEqual([1, 2]);
    expect(v.invoices[0].daysOverdue).toBe(45);
    expect(v.invoices[1].balance).toBe(400);
    expect(v.outstanding).toBe(1400);
  });

  it('filters by days overdue, account and owner', () => {
    const engine = makeEngine({ invoices, accounts });
    expect(invoiceView(engine, normalizeInvoices({ minDaysOverdue: 30 }), NOW).invoices.map(i => i.id)).toEqual([1]);
    expect(invoiceView(engine, normalizeInvoices({ accountName: 'fabri' }), NOW).invoices.map(i => i.id)).toEqual([2]);
    expect(invoiceView(engine, normalizeInvoices({ onlyMine: true }), NOW).invoices.map(i => i.id)).toEqual([1]);
  });
});

describe('GetSlaRiskCases', () => {
  const H = 1 / 24;
  const cases = [
    { id: 1, caseNumber: 'CS-1', title: 'Breached', accountRef: 10, caseStatus: 'In progress', priority: 'High', slaDue: day(-2 * H), owner: ME },
    { id: 2, caseNumber: 'CS-2', title: 'Due in 5h', accountRef: 10, caseStatus: 'New', priority: 'Normal', slaDue: day(5 * H), firstRespondedAt: day(-1), owner: OTHER },
    { id: 3, caseNumber: 'CS-3', title: 'Due in 3 days', accountRef: 10, caseStatus: 'New', priority: 'Low', slaDue: day(3), firstRespondedAt: day(-1), owner: ME },
    { id: 4, caseNumber: 'CS-4', title: 'Paused', accountRef: 10, caseStatus: 'Waiting on customer', priority: 'High', slaDue: day(-1), slaPausedAt: day(-2), owner: ME },
    { id: 5, caseNumber: 'CS-5', title: 'Resolved', accountRef: 10, caseStatus: 'Resolved', priority: 'High', slaDue: day(-1), owner: ME },
    { id: 6, caseNumber: 'CS-6', title: 'Reply late', accountRef: 10, caseStatus: 'New', priority: 'Critical', slaDue: day(2), firstResponseDue: day(-H), owner: ME }
  ];

  it('lists breached cases first, then late replies, then cases due soon; paused and resolved are left out', () => {
    const engine = makeEngine({ cases, accounts });
    const v = caseView(engine, normalizeCases({}), NOW);
    expect(v.cases.map(c => [c.id, c.state])).toEqual([[1, 'breached'], [6, 'firstResponseLate'], [2, 'dueSoon']]);
    expect(v.breached).toBe(1);
    expect(v.dueSoon).toBe(2);
    expect(v.cases[0].hoursLeft).toBe(-2);
  });

  it('honours the window, priority and owner', () => {
    const engine = makeEngine({ cases, accounts });
    expect(caseView(engine, normalizeCases({ windowHours: 96 }), NOW).cases.map(c => c.id)).toEqual([1, 6, 2, 3]);
    expect(caseView(engine, normalizeCases({ priority: 'high' }), NOW).cases.map(c => c.id)).toEqual([1]);
    expect(caseView(engine, normalizeCases({ onlyMine: true }), NOW).cases.map(c => c.id)).toEqual([1, 6]);
    expect(normalizeCases({ priority: 'Urgent' }).priority).toBe('');
  });

  it('reports that cases need the Enterprise plan', () => {
    const v = caseView(makeEngine({ cases, accounts }, standard), normalizeCases({}), NOW);
    expect(v.available).toBe(false);
    expect(v.cases).toEqual([]);
  });
});

describe('CreateLead', () => {
  const sources = ['Event', 'Referral', 'Website'];

  it('prefills only what the prompt gives and matches the lead source to the CRM list', () => {
    const d = normalizeLeadDraft({ firstName: ' Priya ', lastName: 'Shah', email: 'Priya@Contoso.com', company: 'Contoso', leadSource: 'event', rating: 'hot', phone: 42 }, sources);
    expect(d).toEqual({
      firstName: 'Priya', lastName: 'Shah', email: 'priya@contoso.com', phone: '', company: 'Contoso', jobTitle: '', country: '',
      leadSource: 'Event', rating: 'Hot', description: ''
    });
    expect(normalizeLeadDraft({ leadSource: 'Trade show', rating: 'boiling' }, sources)).toMatchObject({ leadSource: '', rating: 'Warm' });
  });

  it('requires a first name and a way to follow up', () => {
    const empty = normalizeLeadDraft({}, sources);
    expect(Object.keys(validateLeadDraft(empty)).sort()).toEqual(['email', 'firstName']);
    expect(validateLeadDraft({ ...empty, firstName: 'Priya', email: 'not-an-email' }).email).toMatch(/valid email/);
    expect(validateLeadDraft({ ...empty, firstName: 'Priya', company: 'Contoso' })).toEqual({});
  });

  it('builds a new lead record without empty fields', () => {
    const r = leadRecord(normalizeLeadDraft({ firstName: 'Priya', company: 'Contoso' }, sources));
    expect(r).toEqual({ firstName: 'Priya', lastName: '', rating: 'Warm', status: 'New', company: 'Contoso' });
  });

  it('is scored and checked for duplicates by the CRM engine', () => {
    const engine = makeEngine({ leads: [{ id: 9, title: 'Priya Shah', email: 'priya@contoso.com', company: 'Contoso' }] });
    const draft = normalizeLeadDraft({ firstName: 'Priya', lastName: 'Shah', email: 'PRIYA@contoso.com' }, sources);
    expect(engine.findDuplicates('leads', leadRecord(draft)).map(l => l.id)).toEqual([9]);
    expect(engine.scoreLead({ ...leadRecord(draft), rating: 'Hot' })).toBeGreaterThan(engine.scoreLead({ ...leadRecord(draft), rating: 'Cold' }));
  });
});
