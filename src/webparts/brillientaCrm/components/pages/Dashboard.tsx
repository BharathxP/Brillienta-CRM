import * as React from 'react';
import {
  ContextualMenuItemType, DefaultButton, Dropdown, Icon, IconButton, IContextualMenuItem, Link, PrimaryButton, Stack, Text, TooltipHost
} from '@fluentui/react';
import { EntityKey, IRecord, LOST, WON } from '../../models/model';
import { IDashboardLayout, WidgetSize } from '../../models/settings';
import { useCrm, useEntities } from '../CrmContext';
import { BarList, ColumnChart, Donut, Funnel } from '../core/charts';
import { Pill, ScoreBadge } from '../core/controls';
import { EmailDialog } from '../core/dialogs';
import { COPILOT_CHAT_URL } from '../core/SmartSummary';
import { draftFollowUp } from '../../services/Insights';
import { IHelpdeskTicket, isOpenTicket } from '../../services/Helpdesk';
import { TicketRow } from '../core/HelpdeskTickets';
import {
  aiForecast, askCrm, ASK_EXAMPLES, briefing, dashboardPrompt, dealHealth, history, IAnswer, IDealHealth, IScope, nextActions, projectedMonths
} from '../../services/DashboardAI';
import { fmtDate, fmtMoney, groupCount, inRange, isOverdue, lsGet, lsSet, monthLabel, quarterRange, slaBreached, sum } from '../utils';

/* ================================================================ building blocks */

/** Card used by the dashboard widgets and other pages (reports, forecast). */
export const Card: React.FC<{ title: string; action?: React.ReactNode; children?: React.ReactNode; wide?: boolean; ai?: boolean; className?: string; icon?: string }> =
  ({ title, action, children, wide, ai, className, icon }) => (
    <div className={`crm-card${wide ? ' crm-wide' : ''}${className ? ` ${className}` : ''}`}>
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center" className="crm-card-title" tokens={{ childrenGap: 8 }}>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} styles={{ root: { minWidth: 0 } }}>
          {icon && <span className={`crm-w-icon${ai ? ' crm-w-icon-ai' : ''}`}><Icon iconName={icon} aria-hidden /></span>}
          <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>{title}</Text>
          {ai && <span className="crm-ai-tag">AI</span>}
        </Stack>
        {action}
      </Stack>
      {children}
    </div>
  );

export const HealthBadge: React.FC<{ h: IDealHealth }> = ({ h }) => (
  <TooltipHost content={[...h.risks.map(x => `▼ ${x}`), ...h.strengths.map(x => `▲ ${x}`)].join('\n') || 'No signals yet'}>
    <span className={`crm-health crm-health-${h.label === 'Healthy' ? 'good' : h.label === 'Watch' ? 'warn' : 'bad'}`} aria-label={`Win likelihood ${h.score}%, ${h.label}`}>{h.score}%</span>
  </TooltipHost>
);

const Row: React.FC<{ onOpen: () => void; children?: React.ReactNode; aside?: React.ReactNode }> = ({ onOpen, children, aside }) => (
  <div className="crm-list-row crm-row-flex">
    <div className="crm-row-main" role="button" tabIndex={0} onClick={onOpen} onKeyDown={e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); onOpen(); } }}>{children}</div>
    {aside && <div className="crm-row-aside">{aside}</div>}
  </div>
);

const Empty: React.FC<{ text: string }> = ({ text }) => <div className="crm-w-empty"><Icon iconName="CheckMark" aria-hidden /> {text}</div>;

/** Small trend line for a KPI tile. */
const Spark: React.FC<{ data: number[]; color: string }> = ({ data, color }) => {
  if (data.length < 2 || !data.some(v => v > 0)) { return null; }
  const max = Math.max(...data, 1);
  const pts = data.map((v, i) => `${(i / (data.length - 1)) * 100},${28 - (v / max) * 24 - 2}`);
  return (
    <svg className="crm-spark" viewBox="0 0 100 28" preserveAspectRatio="none" aria-hidden>
      <polygon points={`0,28 ${pts.join(' ')} 100,28`} style={{ fill: color, opacity: 0.12 }} />
      <polyline points={pts.join(' ')} style={{ stroke: color }} vectorEffect="non-scaling-stroke" />
    </svg>
  );
};

/* ================================================================ layout */

interface IKpiDef { id: string; label: string; icon: string; color: string; value: string | number; hint?: string; delta?: number; spark?: number[]; page?: string; available: boolean }
interface IWidgetDef { id: string; title: string; icon: string; ai?: boolean; size: WidgetSize; available: boolean; action?: React.ReactNode; render: () => React.ReactNode }

const LAYOUT_KEY = 'brillienta-crm:dash-layout';
const SCOPE_KEY = 'brillienta-crm:dash-scope';
const DEFAULT_LAYOUT: IDashboardLayout = {
  kpis: ['won', 'pipeline', 'leads', 'winrate', 'mrr', 'cases', 'helpdesk'],
  widgets: [
    { id: 'actions', size: 'l' }, { id: 'forecast', size: 's' },
    { id: 'revenue', size: 'l' }, { id: 'pipeline', size: 's' },
    { id: 'attention', size: 's' }, { id: 'hotLeads', size: 's' }, { id: 'tasks', size: 's' },
    { id: 'helpdesk', size: 'm' }
  ]
};
const SIZES: { key: WidgetSize; text: string }[] = [{ key: 's', text: 'Small (⅓)' }, { key: 'm', text: 'Medium (½)' }, { key: 'l', text: 'Large (⅔)' }, { key: 'xl', text: 'Full width' }];

function move<T>(list: T[], from: number, to: number): T[] {
  const out = list.slice();
  const [x] = out.splice(from, 1);
  out.splice(Math.max(0, Math.min(out.length, to)), 0, x);
  return out;
}

/* ================================================================ dashboard */

export const Dashboard: React.FC = () => {
  const { engine, currency, navigate, open, settings, data, notify, width, updateSettings } = useCrm();
  const { rows } = useEntities('leads', 'opportunities', 'demos', 'activities', 'cases', 'accounts', 'subscriptions', 'invoices', 'quotes');
  const [scope, setScopeState] = React.useState<string>(() => lsGet(SCOPE_KEY, engine.role === 'sales' ? 'me' : 'all'));
  const setScope = (v: string): void => { setScopeState(v); lsSet(SCOPE_KEY, v); };
  const [layout, setLayoutState] = React.useState<IDashboardLayout>(() => lsGet<IDashboardLayout | null>(LAYOUT_KEY, null) || settings.dashboard || DEFAULT_LAYOUT);
  const setLayout = (l: IDashboardLayout): void => { setLayoutState(l); lsSet(LAYOUT_KEY, l); };
  const [editing, setEditing] = React.useState(false);
  const [drag, setDrag] = React.useState<{ kind: 'kpi' | 'widget'; index: number }>();
  const [question, setQuestion] = React.useState('');
  const [answer, setAnswer] = React.useState<IAnswer>();
  const [draft, setDraft] = React.useState<{ key: EntityKey; r: IRecord }>();
  const [briefOpen, setBriefOpen] = React.useState(false);
  const hdOn = engine.helpdesk.configured;
  const [tickets, setTickets] = React.useState<IHelpdeskTicket[]>();
  const [hdError, setHdError] = React.useState<string>();
  React.useEffect(() => {
    if (!hdOn) { return; }
    engine.helpdesk.customerTickets().then(setTickets).catch(e => { setTickets([]); setHdError(String((e as Error).message || e)); });
  }, [hdOn, data.contacts]);

  const teamIds = scope.indexOf('team:') === 0 ? engine.teamMemberIds(scope.slice(5)) : [];
  const mine = (r: IRecord): boolean => scope === 'all' || (scope === 'me' ? r.owner?.id === engine.me?.id : !!r.owner && teamIds.indexOf(r.owner.id) >= 0);
  const money = (n: number): string => fmtMoney(n, currency);
  const now = new Date();

  const m = React.useMemo(() => {
    const q = quarterRange(0, settings.fiscalYearStartMonth);
    const pq = quarterRange(-1, settings.fiscalYearStartMonth);
    const sc: IScope = {
      leads: rows('leads').filter(mine), opps: rows('opportunities').filter(mine), demos: rows('demos').filter(mine), cases: rows('cases').filter(mine),
      quotes: rows('quotes').filter(mine), invoices: rows('invoices').filter(mine), subs: rows('subscriptions').filter(mine),
      tasks: rows('activities').filter(a => a.activityType === 'Task' && !a.completed && (scope === 'all' ? true : mine(a))),
      quotaUsers: scope === 'all' ? undefined : scope === 'me' ? (engine.me ? [engine.me.id] : []) : teamIds,
      quarter: q, prevQuarter: pq
    };
    const hist = history(engine);
    const cache: Record<number, IDealHealth> = {};
    const health = (o: IRecord): IDealHealth => { const id = o.id as number; if (!cache[id]) { cache[id] = dealHealth(engine, o, hist); } return cache[id]; };
    const fc = aiForecast(engine, sc, health);
    return { q, pq, sc, health, hist, fc, brief: briefing(engine, sc, fc, health), actions: nextActions(engine, sc, health, 20) };
  }, [data, scope, settings]);
  const { q, pq, sc, health, hist, fc, brief, actions } = m;
  const { leads, opps, demos, cases } = sc;

  /* ---------------- numbers */
  const isOpen = (o: IRecord): boolean => engine.openStages().indexOf(o.stage) >= 0;
  const open_ = opps.filter(isOpen);
  const wonIn = (s: Date, e: Date): IRecord[] => opps.filter(o => o.stage === WON && inRange(o.actualCloseDate, s, e));
  const wonQv = sum(wonIn(q.start, q.end), o => engine.baseAmount(o));
  const wonPQv = sum(wonIn(pq.start, pq.end), o => engine.baseAmount(o));
  const closed = opps.filter(o => o.stage === WON || o.stage === LOST);
  const winRate = closed.length ? Math.round((closed.filter(o => o.stage === WON).length / closed.length) * 100) : 0;
  const samePointPQ = new Date(pq.start.getTime() + (pq.end.getTime() - pq.start.getTime()) * fc.elapsed);
  const leadsQ = leads.filter(l => inRange(l.created, q.start, q.end)).length;
  const leadsPQ = leads.filter(l => inRange(l.created, pq.start, samePointPQ)).length;
  const openCases = cases.filter(c => c.caseStatus !== 'Resolved' && c.caseStatus !== 'Closed');
  const breached = openCases.filter(c => slaBreached(c));
  const liveSubs = sc.subs.filter(x => x.subStatus === 'Active' || x.subStatus === 'Pending renewal');
  const mrr = sum(liveSubs, x => engine.toBase(x.mrr, x.currency));
  const unpaid = sc.invoices.filter(x => x.invoiceStatus === 'Sent' || x.invoiceStatus === 'Overdue');
  const wonAll = opps.filter(o => o.stage === WON);

  // 12 monthly buckets for charts and sparklines
  const buckets = Array.from({ length: 12 }, (_, i) => { const s = new Date(now.getFullYear(), now.getMonth() - 11 + i, 1); return { s, e: new Date(s.getFullYear(), s.getMonth() + 1, 1) }; });
  const series = (f: (s: Date, e: Date) => number): number[] => buckets.map(b => f(b.s, b.e));
  const wonByMonth = series((s, e) => sum(wonIn(s, e), o => engine.baseAmount(o)));
  const months = buckets.map((b, i) => ({ label: monthLabel(b.s), short: b.s.toLocaleDateString(undefined, { month: 'short' }), value: wonByMonth[i] }));
  const proj = projectedMonths(engine, opps, health, 3);

  /* ---------------- KPI tiles */
  const kpis: IKpiDef[] = [
    { id: 'won', label: `Won ${q.label}`, icon: 'Trophy2', color: 'var(--series-1)', value: money(wonQv), delta: wonPQv ? ((wonQv - wonPQv) / wonPQv) * 100 : undefined, hint: `vs ${money(wonPQv)} last quarter`, spark: wonByMonth, page: 'forecast', available: true },
    { id: 'pipeline', label: 'Open pipeline', icon: 'Money', color: 'var(--series-3)', value: money(sum(open_, o => engine.baseAmount(o))), hint: `${open_.length} deals · AI weighted ${money(sum(open_, o => (engine.baseAmount(o) * health(o).score) / 100))}`, spark: series((s, e) => sum(opps.filter(o => inRange(o.created, s, e)), o => engine.baseAmount(o))), page: 'entity:opportunities', available: true },
    { id: 'leads', label: 'New leads', icon: 'People', color: 'var(--series-2)', value: leadsQ, delta: leadsPQ ? ((leadsQ - leadsPQ) / leadsPQ) * 100 : undefined, hint: `${q.label} · vs ${leadsPQ} at this point last quarter`, spark: series((s, e) => leads.filter(l => inRange(l.created, s, e)).length), page: 'entity:leads', available: true },
    { id: 'winrate', label: 'Win rate', icon: 'Ribbon', color: 'var(--series-4)', value: `${winRate}%`, hint: `${closed.length} closed deals`, spark: series((s, e) => { const c = opps.filter(o => (o.stage === WON || o.stage === LOST) && inRange(o.actualCloseDate, s, e)); return c.length ? (c.filter(o => o.stage === WON).length / c.length) * 100 : 0; }), page: 'reports', available: true },
    { id: 'avgdeal', label: 'Average deal', icon: 'BarChartVertical', color: 'var(--series-5)', value: money(wonAll.length ? sum(wonAll, o => engine.baseAmount(o)) / wonAll.length : 0), hint: `${wonAll.length} won deals`, page: 'reports', available: true },
    { id: 'cycle', label: 'Sales cycle', icon: 'Timer', color: 'var(--series-6)', value: hist.cycleDays ? `${Math.round(hist.cycleDays)} days` : '—', hint: 'average from created to won', page: 'reports', available: true },
    { id: 'demos', label: 'Demos to schedule', icon: 'Presentation', color: 'var(--series-2)', value: demos.filter(d => d.status === 'Requested').length, hint: `${demos.filter(d => d.status === 'Scheduled').length} scheduled`, spark: series((s, e) => demos.filter(d => inRange(d.created, s, e)).length), page: 'entity:demos', available: true },
    { id: 'mrr', label: 'MRR', icon: 'RecurringEvent', color: 'var(--series-3)', value: money(mrr), hint: `ARR ${money(mrr * 12)}`, page: 'entity:subscriptions', available: engine.hasFeature('subscriptions') && liveSubs.length > 0 },
    { id: 'unpaid', label: 'Unpaid invoices', icon: 'Money', color: 'var(--series-5)', value: money(sum(unpaid, x => engine.toBase((x.grandTotal || 0) - (x.amountPaid || 0), x.currency))), hint: `${unpaid.filter(x => x.invoiceStatus === 'Overdue').length} overdue`, page: 'entity:invoices', available: engine.hasFeature('invoices') },
    { id: 'helpdesk', label: 'Helpdesk tickets', icon: 'Headset', color: 'var(--series-4)', value: tickets ? tickets.filter(isOpenTicket).length : '…', hint: tickets ? `${tickets.filter(t => isOpenTicket(t) && t.breached).length} past SLA · customers' open tickets` : 'loading from Brillienta Helpdesk', page: undefined, available: hdOn },
    { id: 'cases', label: 'Open cases', icon: 'Headset', color: 'var(--series-6)', value: openCases.length, hint: breached.length ? `${breached.length} past SLA` : 'all within SLA', spark: series((s, e) => cases.filter(c => inRange(c.created, s, e)).length), page: 'entity:cases', available: engine.hasFeature('cases') }
  ];
  const kpiById = (id: string): IKpiDef | undefined => kpis.find(k => k.id === id && k.available);

  /* ---------------- widgets */
  const myTasks = rows('activities').filter(a => a.activityType === 'Task' && !a.completed && a.owner?.id === engine.me?.id)
    .sort((a, b) => String(a.dueDate || '9').localeCompare(String(b.dueDate || '9'))).slice(0, 8);
  const upcoming = demos.filter(d => d.status === 'Scheduled' && d.scheduledDate && new Date(d.scheduledDate).getTime() > Date.now() - 3600000)
    .sort((a, b) => String(a.scheduledDate).localeCompare(String(b.scheduledDate))).slice(0, 8);
  const hotLeads = leads.filter(l => l.status !== 'Converted' && l.status !== 'Unqualified').sort((a, b) => (b.leadScore || 0) - (a.leadScore || 0)).slice(0, 8);
  const watchDeals = open_.filter(o => health(o).label !== 'Healthy').sort((a, b) => health(a).score - health(b).score || engine.baseAmount(b) - engine.baseAmount(a)).slice(0, 8);
  const topDeals = [...open_].sort((a, b) => engine.baseAmount(b) - engine.baseAmount(a)).slice(0, 8);
  const scale = Math.max(fc.quota, fc.best, fc.predicted, 1);
  const pct = (n: number): string => `${Math.min(100, (n / scale) * 100)}%`;
  const caseList = [...openCases].sort((a, b) => String(a.slaDue || '9').localeCompare(String(b.slaDue || '9'))).slice(0, 8);

  function ask(text: string): void {
    setQuestion(text);
    setAnswer(text.trim() ? askCrm(engine, sc, text, health, fc) : undefined);
  }

  const widgets: IWidgetDef[] = [
    {
      id: 'actions', title: 'Next best actions', icon: 'AutoEnhanceOn', ai: true, size: 'l', available: true,
      render: () => (actions.length === 0 ? <Empty text="You are all caught up." /> : actions.map((a, i) => (
        <Row key={`${a.key}-${a.r.id}-${i}`} onOpen={() => open(a.key, a.r)}
          aside={a.canDraft ? <IconButton iconProps={{ iconName: 'Mail' }} title="Draft email" ariaLabel={`Draft email: ${a.title}`} onClick={() => setDraft({ key: a.key, r: a.r })} /> : undefined}>
          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
            <span className={`crm-prio crm-prio-${a.rank >= 70 ? 'high' : a.rank >= 50 ? 'med' : 'low'}`} aria-label={a.rank >= 70 ? 'High priority' : a.rank >= 50 ? 'Medium priority' : 'Low priority'} />
            <Text styles={{ root: { fontWeight: 600 } }}>{a.title}</Text>
          </Stack>
          <Text variant="small" block className="crm-muted" styles={{ root: { paddingLeft: 16 } }}>{a.why}</Text>
        </Row>
      )))
    },
    {
      id: 'forecast', title: `AI forecast · ${q.label}`, icon: 'Chart', ai: true, size: 's', available: true,
      action: <Link onClick={() => navigate('forecast')}>Details</Link>,
      render: () => (
        <>
          <div className="crm-fc-big">{money(fc.predicted)}</div>
          <Text variant="small" className="crm-muted" block>{fc.quota > 0 ? `${Math.round((fc.predicted / fc.quota) * 100)}% of ${money(fc.quota)} quota` : 'expected by quarter end'} · {fc.daysLeft} days left</Text>
          <div className="crm-fc-track" role="img" aria-label={`Won ${money(fc.won)}, commit ${money(fc.commit)}, forecast ${money(fc.predicted)}, best case ${money(fc.best)}${fc.quota ? `, quota ${money(fc.quota)}` : ''}`}>
            <span className="crm-fc-best" style={{ width: pct(fc.best) }} />
            <span className="crm-fc-pred" style={{ width: pct(fc.predicted) }} />
            <span className="crm-fc-won" style={{ width: pct(fc.won) }} />
            {fc.quota > 0 && <span className="crm-fc-quota" style={{ left: pct(fc.quota) }} title="Quota" />}
          </div>
          <div className="crm-fc-legend">
            <span><i className="crm-fc-won" />Won<b>{money(fc.won)}</b></span>
            <span><i className="crm-fc-commit" />Commit<b>{money(fc.commit)}</b></span>
            <span><i className="crm-fc-pred" />Forecast<b>{money(fc.predicted)}</b></span>
            <span><i className="crm-fc-best" />Best case<b>{money(fc.best)}</b></span>
          </div>
        </>
      )
    },
    {
      id: 'revenue', title: 'Revenue won — last 12 months', icon: 'BarChartVertical', size: 'l', available: true,
      action: proj.some(p => p.value > 0) ? <span className="crm-legend-proj"><i />AI projection</span> : undefined,
      render: () => <ColumnChart data={[...months, ...proj.map(p => ({ ...p, projected: true }))]} format={money} ariaLabel="Revenue won by month and projected revenue for the next three months" />
    },
    {
      id: 'pipeline', title: 'Pipeline by stage', icon: 'FunnelChart', size: 's', available: true,
      action: <Link onClick={() => navigate('entity:opportunities')}>Board</Link>,
      render: () => <BarList data={engine.openStages().map(st => ({ label: st, value: sum(open_.filter(o => o.stage === st), o => engine.baseAmount(o)), hint: `${open_.filter(o => o.stage === st).length} deals` }))} format={money} />
    },
    {
      id: 'attention', title: 'Deals at risk', icon: 'Warning', ai: true, size: 's', available: true,
      action: <Link onClick={() => ask('Deals that need attention')}>All</Link>,
      render: () => (watchDeals.length === 0 ? <Empty text="Every open deal looks healthy." /> : watchDeals.map(o => (
        <Row key={o.id} onOpen={() => open('opportunities', o)} aside={<HealthBadge h={health(o)} />}>
          <Text styles={{ root: { fontWeight: 600 } }}>{o.title}</Text>
          <Text variant="small" block className="crm-muted">{money(engine.baseAmount(o))} · {health(o).risks[0] || o.stage}</Text>
        </Row>
      )))
    },
    {
      id: 'hotLeads', title: 'Hottest leads', icon: 'People', size: 's', available: true,
      action: <Link onClick={() => navigate('entity:leads')}>All</Link>,
      render: () => (hotLeads.length === 0 ? <Empty text="No open leads." /> : hotLeads.map(l => (
        <Row key={l.id} onOpen={() => open('leads', l)} aside={<ScoreBadge score={l.leadScore} />}>
          <Text styles={{ root: { fontWeight: 600 } }}>{l.title}</Text>
          <Text variant="small" block className="crm-muted">{[l.company, l.leadSource, fmtDate(l.created)].filter(Boolean).join(' · ')}</Text>
        </Row>
      )))
    },
    {
      id: 'tasks', title: 'My open tasks', icon: 'Clock', size: 's', available: true,
      action: <Link onClick={() => navigate('mywork')}>My work</Link>,
      render: () => (myTasks.length === 0 ? <Empty text="Nothing due." /> : myTasks.map(a => (
        <Row key={a.id} onOpen={() => open('activities', a)} aside={isOverdue(a.dueDate) ? <Pill text="Overdue" /> : undefined}>
          <Text>{a.title}</Text>
          <Text variant="small" block className="crm-muted">{a.dueDate ? fmtDate(a.dueDate, true) : 'No due date'}</Text>
        </Row>
      )))
    },
    {
      id: 'topDeals', title: 'Biggest open deals', icon: 'Trophy2', size: 's', available: true,
      render: () => (topDeals.length === 0 ? <Empty text="No open deals." /> : topDeals.map(o => (
        <Row key={o.id} onOpen={() => open('opportunities', o)} aside={<HealthBadge h={health(o)} />}>
          <Text styles={{ root: { fontWeight: 600 } }}>{o.title}</Text>
          <Text variant="small" block className="crm-muted">{fmtMoney(o.amount || 0, o.currency || currency)} · {o.stage} · close {fmtDate(o.expectedCloseDate)}</Text>
        </Row>
      )))
    },
    {
      id: 'funnel', title: 'Conversion funnel', icon: 'FunnelChart', size: 's', available: true,
      render: () => <Funnel data={[
        { label: 'Leads', value: leads.length },
        { label: 'Qualified', value: leads.filter(l => ['Qualified', 'Converted'].indexOf(l.status) >= 0).length },
        { label: 'Demos', value: demos.filter(d => d.status !== 'Cancelled').length },
        { label: 'Opportunities', value: opps.length },
        { label: 'Won', value: wonAll.length }
      ]} />
    },
    { id: 'sources', title: 'Leads by source', icon: 'DonutChart', size: 's', available: true, render: () => <Donut data={groupCount(leads, l => l.leadSource)} ariaLabel="Leads by source" centerLabel="leads" /> },
    {
      id: 'demos', title: 'Upcoming demos', icon: 'Presentation', size: 's', available: true,
      action: <Link onClick={() => navigate('calendar')}>Calendar</Link>,
      render: () => (upcoming.length === 0 ? <Empty text="No scheduled demos." /> : upcoming.map(d => (
        <Row key={d.id} onOpen={() => open('demos', d)}>
          <Text styles={{ root: { fontWeight: 600 } }}>{d.title}</Text>
          <Text variant="small" block className="crm-muted">{fmtDate(d.scheduledDate, true)} · {d.owner?.title || 'No presenter'}</Text>
        </Row>
      )))
    },
    {
      id: 'helpdesk', title: 'Helpdesk tickets', icon: 'Headset', size: 'm', available: hdOn,
      action: <Link href={engine.helpdesk.pageUrl()} target="_blank" data-interception="off">Open Helpdesk</Link>,
      render: () => (hdError ? <Text className="crm-muted">Could not read Brillienta Helpdesk: {hdError}</Text>
        : !tickets ? <Text className="crm-muted">Loading…</Text>
          : tickets.filter(isOpenTicket).length === 0 ? <Empty text="No open tickets from your customers." />
            : tickets.filter(isOpenTicket).slice(0, 10).map(t => <TicketRow key={t.id} t={t} sub={engine.helpdesk.accountOf(t)?.title || t.requester} />))
    },
    {
      id: 'cases', title: 'Open support cases', icon: 'Headset', size: 's', available: engine.hasFeature('cases'),
      action: <Link onClick={() => navigate('entity:cases')}>All</Link>,
      render: () => (caseList.length === 0 ? <Empty text="No open cases." /> : caseList.map(c => (
        <Row key={c.id} onOpen={() => open('cases', c)} aside={slaBreached(c) ? <Pill text="SLA" /> : undefined}>
          <Text styles={{ root: { fontWeight: 600 } }}>{c.caseNumber} {c.title}</Text>
          <Text variant="small" block className="crm-muted">{c.priority || 'Normal'} · {c.caseStatus}{c.slaDue ? ` · due ${fmtDate(c.slaDue, true)}` : ''}</Text>
        </Row>
      )))
    }
  ];
  const widgetById = (id: string): IWidgetDef | undefined => widgets.find(w => w.id === id && w.available);

  const shownKpis = layout.kpis.filter(id => !!kpiById(id));
  const shownWidgets = layout.widgets.filter(w => !!widgetById(w.id));
  const hiddenKpis = kpis.filter(k => k.available && shownKpis.indexOf(k.id) < 0);
  const hiddenWidgets = widgets.filter(w => w.available && !shownWidgets.some(x => x.id === w.id));

  /* ---------------- actions */
  const askCopilot = (): void => {
    const prompt = dashboardPrompt(engine, sc, fc, brief, actions.slice(0, 10), health);
    const done = (): void => { notify('Prompt copied. Paste it into Microsoft 365 Copilot (Ctrl+V).'); window.open(COPILOT_CHAT_URL, '_blank', 'noopener'); };
    if (navigator.clipboard) { navigator.clipboard.writeText(prompt).then(done).catch(() => window.prompt('Copy this prompt into Copilot', prompt)); } else { window.prompt('Copy this prompt into Copilot', prompt); }
  };
  const setWidgets = (w: IDashboardLayout['widgets']): void => setLayout({ ...layout, widgets: w });
  const setKpis = (k: string[]): void => setLayout({ ...layout, kpis: k });
  const reset = (): void => { setLayoutState(settings.dashboard || DEFAULT_LAYOUT); lsSet(LAYOUT_KEY, null); notify('Dashboard reset to the default layout.'); };
  const saveDefault = (): void => {
    updateSettings({ ...settings, dashboard: { kpis: shownKpis, widgets: shownWidgets } })
      .then(() => notify('Saved as the default dashboard for everyone. People who customised their own keep it until they reset.'))
      .catch(() => notify('Could not save the default layout.'));
  };
  const dropOn = (kind: 'kpi' | 'widget', index: number): void => {
    if (!drag || drag.kind !== kind || drag.index === index) { setDrag(undefined); return; }
    if (kind === 'kpi') { setKpis(move(shownKpis, drag.index, index)); } else { setWidgets(move(shownWidgets, drag.index, index)); }
    setDrag(undefined);
  };
  const dragProps = (kind: 'kpi' | 'widget', index: number): React.HTMLAttributes<HTMLDivElement> & { draggable?: boolean } => (editing ? {
    draggable: true,
    onDragStart: e => { setDrag({ kind, index }); e.dataTransfer.effectAllowed = 'move'; try { e.dataTransfer.setData('text/plain', String(index)); } catch { /* older browsers */ } },
    onDragOver: e => { if (drag?.kind === kind) { e.preventDefault(); } },
    onDrop: e => { e.preventDefault(); dropOn(kind, index); },
    onDragEnd: () => setDrag(undefined)
  } : {});

  const cols = width < 760 ? 1 : width < 1100 ? 2 : 3;
  const span = (s: WidgetSize): number => (cols === 1 ? 12 : cols === 2 ? (s === 's' || s === 'm' ? 6 : 12) : ({ s: 4, m: 6, l: 8, xl: 12 } as Record<WidgetSize, number>)[s]);
  const hour = now.getHours();
  const briefShown = briefOpen ? brief : brief.slice(0, 3);
  const addWidgetItems: IContextualMenuItem[] = hiddenWidgets.map(w => ({ key: w.id, text: w.title, iconProps: { iconName: w.icon }, onClick: () => setWidgets([...shownWidgets, { id: w.id, size: w.size }]) }));

  const widgetMenu = (index: number, size: WidgetSize): IContextualMenuItem[] => [
    ...SIZES.map(sz => ({ key: sz.key, text: sz.text, canCheck: true, checked: sz.key === size, onClick: () => setWidgets(shownWidgets.map((w, i) => (i === index ? { ...w, size: sz.key } : w))) })),
    { key: 'd1', itemType: ContextualMenuItemType.Divider },
    { key: 'up', text: 'Move earlier', iconProps: { iconName: 'ChevronUp' }, disabled: index === 0, onClick: () => setWidgets(move(shownWidgets, index, index - 1)) },
    { key: 'down', text: 'Move later', iconProps: { iconName: 'ChevronDown' }, disabled: index === shownWidgets.length - 1, onClick: () => setWidgets(move(shownWidgets, index, index + 1)) },
    { key: 'd2', itemType: ContextualMenuItemType.Divider },
    { key: 'rm', text: 'Remove', iconProps: { iconName: 'Cancel' }, onClick: () => setWidgets(shownWidgets.filter((_, i) => i !== index)) }
  ];

  return (
    <div className={`crm-dash${cols === 1 ? ' crm-dash-narrow' : ''}${editing ? ' crm-dash-editing' : ''}`}>
      {/* ---------- hero: greeting, Ask and AI briefing ---------- */}
      <div className="crm-hero">
        <div className="crm-hero-main">
          <div className="crm-hero-title">Good {hour < 12 ? 'morning' : hour < 18 ? 'afternoon' : 'evening'}, {(engine.me?.title || '').split(' ')[0]}</div>
          <div className="crm-hero-sub">{now.toLocaleDateString(undefined, { weekday: 'long', day: 'numeric', month: 'long' })} · {q.label}, {fc.daysLeft} days left</div>
          <form className="crm-hero-ask" role="search" onSubmit={e => { e.preventDefault(); ask(question); }}>
            <Icon iconName="AutoEnhanceOn" aria-hidden />
            <input value={question} onChange={e => { setQuestion(e.target.value); if (!e.target.value) { setAnswer(undefined); } }} aria-label="Ask about your CRM"
              placeholder="Ask anything about your CRM…" />
            {question && <button type="button" className="crm-hero-x" aria-label="Clear question" onClick={() => ask('')}><Icon iconName="Cancel" /></button>}
            <button type="submit" className="crm-hero-go" disabled={!question.trim()}>Ask</button>
          </form>
          <div className="crm-hero-chips">
            {ASK_EXAMPLES.slice(0, 4).map(x => <button key={x} type="button" onClick={() => ask(x)}>{x}</button>)}
          </div>
        </div>
        <div className="crm-hero-brief" aria-label="AI briefing" role="region">
          <div className="crm-hero-brief-h"><Icon iconName="AutoEnhanceOn" aria-hidden /> AI briefing</div>
          <ul>
            {briefShown.map((b, i) => (
              <li key={i} className={`crm-hb-${b.tone}`}>
                <span className="crm-hb-dot" aria-hidden />
                {b.page ? <button type="button" onClick={() => navigate(b.page!)}>{b.text}</button> : <span>{b.text}</span>}
              </li>
            ))}
          </ul>
          {brief.length > 3 && <button type="button" className="crm-hero-more" onClick={() => setBriefOpen(!briefOpen)}>{briefOpen ? 'Show less' : `${brief.length - 3} more`}</button>}
        </div>
      </div>

      {answer && (
        <div className="crm-card crm-ask-answer" aria-live="polite">
          <Stack horizontal horizontalAlign="space-between" verticalAlign="start" tokens={{ childrenGap: 8 }}>
            <Text styles={{ root: { fontWeight: 600 } }}><Icon iconName="AutoEnhanceOn" className="crm-ai-icon" aria-hidden /> {answer.text}</Text>
            <IconButton iconProps={{ iconName: 'Cancel' }} ariaLabel="Close answer" onClick={() => ask('')} />
          </Stack>
          {answer.items.map(it => (
            <Row key={`${it.key}-${it.r.id}`} onOpen={() => open(it.key, it.r)} aside={it.key === 'opportunities' && isOpen(it.r) ? <HealthBadge h={health(it.r)} /> : undefined}>
              <Text styles={{ root: { fontWeight: 600 } }}>{it.r.title || it.r.invoiceNumber || it.r.quoteNumber}</Text>
              <Text variant="small" block className="crm-muted">{it.sub}</Text>
            </Row>
          ))}
          {answer.unknown && <Link onClick={askCopilot}>Ask Microsoft 365 Copilot instead</Link>}
        </div>
      )}

      {/* ---------- toolbar ---------- */}
      <div className="crm-dash-bar">
        {editing ? (
          <>
            <Text styles={{ root: { fontWeight: 600 } }}><Icon iconName="Personalize" /> Drag tiles and cards to reorder. Use ⋯ to resize or remove.</Text>
            <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
              <DefaultButton iconProps={{ iconName: 'Add' }} text="Add widget" disabled={!hiddenWidgets.length} menuProps={{ items: addWidgetItems }} />
              <DefaultButton iconProps={{ iconName: 'Add' }} text="Add KPI" disabled={!hiddenKpis.length}
                menuProps={{ items: hiddenKpis.map(k => ({ key: k.id, text: k.label, iconProps: { iconName: k.icon }, onClick: () => setKpis([...shownKpis, k.id]) })) }} />
              <DefaultButton iconProps={{ iconName: 'Undo' }} text="Reset" onClick={reset} />
              {engine.role === 'admin' && <DefaultButton iconProps={{ iconName: 'Save' }} text="Save as company default" onClick={saveDefault} />}
              <PrimaryButton text="Done" onClick={() => setEditing(false)} />
            </Stack>
          </>
        ) : (
          <>
            <Dropdown selectedKey={scope} ariaLabel="Scope" styles={{ root: { width: 170 } }} onChange={(_, o) => setScope(String(o?.key))}
              options={[{ key: 'all', text: 'Whole company' }, ...(engine.role === 'sales' ? engine.teamsOf(engine.me?.id) : settings.teams).map(t => ({ key: `team:${t.id}`, text: `Team ${t.name}` })), { key: 'me', text: 'My records' }]} />
            <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
              <TooltipHost content="Copies a summary of this dashboard as a prompt and opens Microsoft 365 Copilot for a deeper analysis.">
                <DefaultButton iconProps={{ iconName: 'Robot' }} text="Ask Copilot" onClick={askCopilot} />
              </TooltipHost>
              <DefaultButton iconProps={{ iconName: 'Personalize' }} text="Customise" onClick={() => setEditing(true)} />
            </Stack>
          </>
        )}
      </div>

      {/* ---------- KPI tiles ---------- */}
      {shownKpis.length > 0 && (
        <div className="crm-tiles">
          {shownKpis.map((id, i) => {
            const k = kpiById(id)!;
            const go = (): void => { if (!editing && k.page) { navigate(k.page); } };
            return (
              <div key={id} className={`crm-tile${drag?.kind === 'kpi' && drag.index === i ? ' crm-dragging' : ''}`} {...dragProps('kpi', i)}
                role={editing ? undefined : 'button'} tabIndex={editing ? undefined : 0} onClick={go} onKeyDown={e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); go(); } }}
                style={{ ['--tile' as string]: k.color } as React.CSSProperties}>
                <div className="crm-tile-head">
                  <span className="crm-tile-icon"><Icon iconName={k.icon} aria-hidden /></span>
                  <span className="crm-tile-label">{k.label}</span>
                  {editing && <IconButton className="crm-tile-x" iconProps={{ iconName: 'Cancel' }} ariaLabel={`Remove ${k.label}`} onClick={() => setKpis(shownKpis.filter(x => x !== id))} />}
                </div>
                <div className="crm-tile-value">{k.value}</div>
                <div className="crm-tile-foot">
                  {k.delta !== undefined && isFinite(k.delta) && <span className={k.delta >= 0 ? 'crm-up' : 'crm-down'}>{k.delta >= 0 ? '▲' : '▼'} {Math.abs(Math.round(k.delta))}%</span>}
                  <span className="crm-tile-hint">{k.hint}</span>
                </div>
                {k.spark && <Spark data={k.spark} color={k.color} />}
              </div>
            );
          })}
        </div>
      )}

      {/* ---------- widgets ---------- */}
      <div className="crm-wgrid">
        {shownWidgets.map((w, i) => {
          const def = widgetById(w.id)!;
          return (
            <div key={w.id} className={`crm-w crm-w-${w.id}${drag?.kind === 'widget' && drag.index === i ? ' crm-dragging' : ''}`} style={{ gridColumn: `span ${span(w.size)}` }} {...dragProps('widget', i)}>
              <Card title={def.title} icon={def.icon} ai={def.ai}
                action={editing ? (
                  <Stack horizontal verticalAlign="center">
                    <Icon iconName="GripperDotsVertical" className="crm-grip" aria-hidden />
                    <IconButton iconProps={{ iconName: 'More' }} ariaLabel={`Options for ${def.title}`} menuProps={{ items: widgetMenu(i, w.size) }} onRenderMenuIcon={() => null} />
                  </Stack>
                ) : def.action}>
                <div className="crm-w-body">{def.render()}</div>
              </Card>
            </div>
          );
        })}
        {editing && hiddenWidgets.length > 0 && (
          <div className="crm-w crm-w-add" style={{ gridColumn: `span ${span('s')}` }}>
            <DefaultButton iconProps={{ iconName: 'Add' }} text="Add widget" menuProps={{ items: addWidgetItems }} />
          </div>
        )}
      </div>
      {shownWidgets.length === 0 && !editing && <Empty text="No widgets on your dashboard. Use Customise to add some." />}

      {draft && (() => {
        const d = draftFollowUp(engine, draft.key, draft.r);
        return <EmailDialog entity={draft.key} record={draft.r} to={d.to} subject={d.subject} body={d.body} onClose={() => setDraft(undefined)} />;
      })()}
    </div>
  );
};
