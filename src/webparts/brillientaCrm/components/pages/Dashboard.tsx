import * as React from 'react';
import { DefaultButton, Dropdown, Icon, IconButton, Link, PrimaryButton, Stack, Text, TextField, TooltipHost } from '@fluentui/react';
import { EntityKey, IRecord, LOST, WON } from '../../models/model';
import { useCrm, useEntities } from '../CrmContext';
import { BarList, ColumnChart, Donut, Funnel } from '../core/charts';
import { Pill, ScoreBadge } from '../core/controls';
import { EmailDialog } from '../core/dialogs';
import { COPILOT_CHAT_URL } from '../core/SmartSummary';
import { draftFollowUp } from '../../services/Insights';
import {
  aiForecast, askCrm, ASK_EXAMPLES, briefing, dashboardPrompt, dealHealth, history, IAction, IAnswer, IDealHealth, IScope, nextActions, projectedMonths
} from '../../services/DashboardAI';
import { fmtDate, fmtMoney, groupCount, inRange, isOverdue, lsGet, lsSet, monthLabel, quarterRange, slaBreached, sum } from '../utils';

const Kpi: React.FC<{ label: string; value: string | number; hint?: string; icon: string; onClick?: () => void; delta?: number }> = p => (
  <div className="crm-kpi" onClick={p.onClick} role={p.onClick ? 'button' : undefined} tabIndex={p.onClick ? 0 : undefined}
    onKeyDown={e => { if (p.onClick && (e.key === 'Enter' || e.key === ' ')) { p.onClick(); } }}>
    <Icon iconName={p.icon} className="crm-kpi-icon" />
    <div className="crm-kpi-value">{p.value}</div>
    <div className="crm-kpi-label">{p.label}</div>
    {(p.hint || p.delta !== undefined) && (
      <div className="crm-kpi-hint">
        {p.delta !== undefined && isFinite(p.delta) && <span style={{ color: p.delta >= 0 ? 'var(--good)' : 'var(--bad)' }}>{p.delta >= 0 ? '▲' : '▼'} {Math.abs(Math.round(p.delta))}% </span>}
        {p.hint}
      </div>
    )}
  </div>
);

export const Card: React.FC<{ title: string; action?: React.ReactNode; children?: React.ReactNode; wide?: boolean; ai?: boolean; className?: string }> = ({ title, action, children, wide, ai, className }) => (
  <div className={`crm-card${wide ? ' crm-wide' : ''}${ai ? ' crm-card-ai' : ''}${className ? ` ${className}` : ''}`}>
    <Stack horizontal horizontalAlign="space-between" verticalAlign="center" className="crm-card-title" tokens={{ childrenGap: 8 }}>
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
        {ai && <Icon iconName="AutoEnhanceOn" className="crm-ai-icon" aria-hidden />}
        <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>{title}</Text>
        {ai && <span className="crm-ai-tag">AI</span>}
      </Stack>
      {action}
    </Stack>
    {children}
  </div>
);

const Section: React.FC<{ title: string; children?: React.ReactNode }> = ({ title, children }) => (
  <section className="crm-dash-section" aria-label={title}>
    <h3 className="crm-dash-h">{title}</h3>
    {children}
  </section>
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

const SCOPE_KEY = 'brillienta-crm:dash-scope';

export const Dashboard: React.FC = () => {
  const { engine, currency, navigate, open, settings, data, notify, isNarrow } = useCrm();
  const { rows } = useEntities('leads', 'opportunities', 'demos', 'activities', 'cases', 'accounts', 'subscriptions', 'invoices', 'quotes');
  const [scope, setScopeState] = React.useState<string>(() => lsGet(SCOPE_KEY, engine.role === 'sales' ? 'me' : 'all'));
  const setScope = (v: string): void => { setScopeState(v); lsSet(SCOPE_KEY, v); };
  const [question, setQuestion] = React.useState('');
  const [answer, setAnswer] = React.useState<IAnswer>();
  const [draft, setDraft] = React.useState<{ key: EntityKey; r: IRecord }>();
  const [allActions, setAllActions] = React.useState(false);

  const teamIds = scope.indexOf('team:') === 0 ? engine.teamMemberIds(scope.slice(5)) : [];
  const mine = (r: IRecord): boolean => scope === 'all' || (scope === 'me' ? r.owner?.id === engine.me?.id : !!r.owner && teamIds.indexOf(r.owner.id) >= 0);
  const money = (n: number): string => fmtMoney(n, currency);
  const now = new Date();

  // Everything derived from the data is computed once per data change / scope change.
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
    const brief = briefing(engine, sc, fc, health);
    const actions = nextActions(engine, sc, health, 20);
    return { q, pq, sc, health, fc, brief, actions };
  }, [data, scope, settings]);
  const { q, pq, sc, health, fc, brief, actions } = m;
  const { leads, opps, demos, cases } = sc;

  const open_ = opps.filter(o => engine.openStages().indexOf(o.stage) >= 0);
  const wonQ = opps.filter(o => o.stage === WON && inRange(o.actualCloseDate, q.start, q.end));
  const wonPQ = opps.filter(o => o.stage === WON && inRange(o.actualCloseDate, pq.start, pq.end));
  const wonQv = sum(wonQ, o => engine.baseAmount(o));
  const wonPQv = sum(wonPQ, o => engine.baseAmount(o));
  const closed = opps.filter(o => o.stage === WON || o.stage === LOST);
  const winRate = closed.length ? Math.round((closed.filter(o => o.stage === WON).length / closed.length) * 100) : 0;
  const leadsQ = leads.filter(l => inRange(l.created, q.start, q.end)).length;
  // Compared with the same point of last quarter, so a quarter that has just started is not shown as a drop.
  const leadsPQ = leads.filter(l => inRange(l.created, pq.start, new Date(pq.start.getTime() + (pq.end.getTime() - pq.start.getTime()) * fc.elapsed))).length;
  const openCases = cases.filter(c => c.caseStatus !== 'Resolved' && c.caseStatus !== 'Closed');
  const breached = openCases.filter(c => slaBreached(c));
  const liveSubs = sc.subs.filter(x => x.subStatus === 'Active' || x.subStatus === 'Pending renewal');
  const mrr = sum(liveSubs, x => engine.toBase(x.mrr, x.currency));
  const renewSoon = liveSubs.filter(x => x.renewalDate && new Date(x.renewalDate).getTime() - Date.now() < settings.renewals.leadDays * 86400000);
  const unpaid = sc.invoices.filter(x => x.invoiceStatus === 'Sent' || x.invoiceStatus === 'Overdue');
  const overdueInv = unpaid.filter(x => x.invoiceStatus === 'Overdue');

  // Revenue won, last 12 months, followed by the AI projection for the next 3.
  const months: { label: string; value: number; short?: string; projected?: boolean }[] = [];
  for (let i = 11; i >= 0; i--) {
    const s = new Date(now.getFullYear(), now.getMonth() - i, 1);
    const e = new Date(now.getFullYear(), now.getMonth() - i + 1, 1);
    months.push({ label: monthLabel(s), short: s.toLocaleDateString(undefined, { month: 'short' }), value: sum(opps.filter(o => o.stage === WON && inRange(o.actualCloseDate, s, e)), o => engine.baseAmount(o)) });
  }
  const proj = projectedMonths(engine, opps, health, 3);
  const chart = [...months, ...proj.map(p => ({ ...p, projected: true }))];

  const byStage = engine.openStages().map(st => ({ label: st, value: sum(open_.filter(o => o.stage === st), o => engine.baseAmount(o)), hint: `${open_.filter(o => o.stage === st).length} deals` }));
  const bySource = groupCount(leads, l => l.leadSource);
  const funnel = [
    { label: 'Leads', value: leads.length },
    { label: 'Qualified', value: leads.filter(l => ['Qualified', 'Converted'].indexOf(l.status) >= 0).length },
    { label: 'Demos', value: demos.filter(d => d.status !== 'Cancelled').length },
    { label: 'Opportunities', value: opps.length },
    { label: 'Won', value: opps.filter(o => o.stage === WON).length }
  ];
  const myTasks = rows('activities').filter(a => a.activityType === 'Task' && !a.completed && a.owner?.id === engine.me?.id)
    .sort((a, b) => String(a.dueDate || '9').localeCompare(String(b.dueDate || '9'))).slice(0, 6);
  const upcoming = demos.filter(d => d.status === 'Scheduled' && d.scheduledDate && new Date(d.scheduledDate).getTime() > Date.now() - 3600000)
    .sort((a, b) => String(a.scheduledDate).localeCompare(String(b.scheduledDate))).slice(0, 5);
  const hotLeads = leads.filter(l => l.status !== 'Converted' && l.status !== 'Unqualified').sort((a, b) => (b.leadScore || 0) - (a.leadScore || 0)).slice(0, 6);
  const watchDeals = [...open_].sort((a, b) => health(a).score - health(b).score || engine.baseAmount(b) - engine.baseAmount(a)).filter(o => health(o).label !== 'Healthy').slice(0, 6);
  const topDeals = [...open_].sort((a, b) => engine.baseAmount(b) - engine.baseAmount(a)).slice(0, 6);
  const healthMix = ['Healthy', 'Watch', 'At risk'].map(l => ({ label: l, value: sum(open_.filter(o => health(o).label === l), o => engine.baseAmount(o)), hint: `${open_.filter(o => health(o).label === l).length} deals` }));

  const ask = (text: string): void => {
    setQuestion(text);
    setAnswer(text.trim() ? askCrm(engine, sc, text, health, fc) : undefined);
  };
  const askCopilot = (): void => {
    const prompt = dashboardPrompt(engine, sc, fc, brief, actions.slice(0, 10), health);
    const done = (): void => { notify('Prompt copied. Paste it into Microsoft 365 Copilot (Ctrl+V).'); window.open(COPILOT_CHAT_URL, '_blank', 'noopener'); };
    if (navigator.clipboard) { navigator.clipboard.writeText(prompt).then(done).catch(() => window.prompt('Copy this prompt into Copilot', prompt)); } else { window.prompt('Copy this prompt into Copilot', prompt); }
  };

  const hour = now.getHours();
  const shown: IAction[] = allActions ? actions : actions.slice(0, 6);
  const scale = Math.max(fc.quota, fc.best, fc.predicted, 1);
  const pct = (n: number): string => `${Math.min(100, (n / scale) * 100)}%`;

  return (
    <div className={`crm-dash${isNarrow ? ' crm-dash-narrow' : ''}`}>
      <div className="crm-dash-top">
        <div>
          <Text variant="xLarge" block styles={{ root: { fontWeight: 600 } }}>Good {hour < 12 ? 'morning' : hour < 18 ? 'afternoon' : 'evening'}, {(engine.me?.title || '').split(' ')[0]}</Text>
          <Text className="crm-muted">{now.toLocaleDateString(undefined, { weekday: 'long', day: 'numeric', month: 'long' })} · {q.label}, {fc.daysLeft} days left</Text>
        </div>
        <Stack horizontal wrap verticalAlign="center" tokens={{ childrenGap: 8 }}>
          <Dropdown selectedKey={scope} ariaLabel="Scope" styles={{ root: { width: 170 } }} onChange={(_, o) => setScope(String(o?.key))}
            options={[{ key: 'all', text: 'Whole company' }, ...(engine.role === 'sales' ? engine.teamsOf(engine.me?.id) : settings.teams).map(t => ({ key: `team:${t.id}`, text: `Team ${t.name}` })), { key: 'me', text: 'My records' }]} />
          <TooltipHost content="Copies a summary of this dashboard (forecast, risks and largest deals) as a prompt and opens Microsoft 365 Copilot for a deeper analysis.">
            <DefaultButton iconProps={{ iconName: 'Robot' }} text="Ask Copilot" onClick={askCopilot} />
          </TooltipHost>
        </Stack>
      </div>

      {/* Ask your CRM */}
      <div className="crm-ask" role="search">
        <form className="crm-ask-bar" onSubmit={e => { e.preventDefault(); ask(question); }}>
          <Icon iconName="AutoEnhanceOn" className="crm-ai-icon" aria-hidden />
          <TextField borderless value={question} onChange={(_, v) => { setQuestion(v || ''); if (!v) { setAnswer(undefined); } }} ariaLabel="Ask about your CRM"
            placeholder="Ask about your CRM, e.g. “deals closing this month over 5k”" styles={{ root: { flex: 1 }, field: { fontSize: 15 } }} />
          {question && <IconButton iconProps={{ iconName: 'Cancel' }} ariaLabel="Clear question" onClick={() => ask('')} />}
          <PrimaryButton type="submit" text="Ask" disabled={!question.trim()} />
        </form>
        {!answer && (
          <div className="crm-ask-chips">
            {ASK_EXAMPLES.map(x => <button key={x} type="button" className="crm-chip" onClick={() => ask(x)}>{x}</button>)}
          </div>
        )}
        {answer && (
          <div className="crm-ask-answer" aria-live="polite">
            <Text block styles={{ root: { fontWeight: 600, marginBottom: answer.items.length ? 6 : 0 } }}>{answer.text}</Text>
            {answer.items.map(it => (
              <Row key={`${it.key}${it.r.id}`} onOpen={() => open(it.key, it.r)}
                aside={it.key === 'opportunities' && engine.openStages().indexOf(it.r.stage) >= 0 ? <HealthBadge h={health(it.r)} /> : undefined}>
                <Text styles={{ root: { fontWeight: 600 } }}>{it.r.title || it.r.invoiceNumber || it.r.quoteNumber}</Text>
                <Text variant="small" block className="crm-muted">{it.sub}</Text>
              </Row>
            ))}
            {answer.unknown && <Link onClick={askCopilot}>Ask Microsoft 365 Copilot instead</Link>}
          </div>
        )}
      </div>

      {/* Briefing and forecast */}
      <div className="crm-dash-hero">
        <Card title="Today’s briefing" ai>
          <ul className="crm-brief">
            {brief.map((b, i) => (
              <li key={i} className={`crm-brief-${b.tone}`}>
                <Icon iconName={b.icon} aria-hidden />
                <span>{b.text}{b.page && <> <Link onClick={() => navigate(b.page!)}>View</Link></>}</span>
              </li>
            ))}
          </ul>
        </Card>
        <Card title={`AI forecast · ${q.label}`} ai action={<Link onClick={() => navigate('forecast')}>Forecast</Link>}>
          <div className="crm-fc-big">{money(fc.predicted)}</div>
          <Text variant="small" className="crm-muted" block>
            {fc.quota > 0 ? `${Math.round((fc.predicted / fc.quota) * 100)}% of ${money(fc.quota)} quota` : 'Predicted revenue by quarter end'} · from {fc.deals} open deal{fc.deals === 1 ? '' : 's'}
          </Text>
          <div className="crm-fc-track" role="img" aria-label={`Won ${money(fc.won)}, commit ${money(fc.commit)}, forecast ${money(fc.predicted)}, best case ${money(fc.best)}${fc.quota ? `, quota ${money(fc.quota)}` : ''}`}>
            <span className="crm-fc-best" style={{ width: pct(fc.best) }} />
            <span className="crm-fc-pred" style={{ width: pct(fc.predicted) }} />
            <span className="crm-fc-won" style={{ width: pct(fc.won) }} />
            {fc.quota > 0 && <span className="crm-fc-quota" style={{ left: pct(fc.quota) }} title="Quota" />}
            <span className="crm-fc-now" style={{ left: `${fc.elapsed * 100}%` }} title="Today" />
          </div>
          <div className="crm-fc-legend">
            <span><i className="crm-fc-won" />Won <b>{money(fc.won)}</b></span>
            <span><i className="crm-fc-commit" />Commit <b>{money(fc.commit)}</b></span>
            <span><i className="crm-fc-pred" />Forecast <b>{money(fc.predicted)}</b></span>
            <span><i className="crm-fc-best" />Best case <b>{money(fc.best)}</b></span>
          </div>
          <Text variant="small" className="crm-muted" block>Weighted by each deal’s AI win likelihood: activity, close date, quotes, contacts and your own win history.</Text>
        </Card>
      </div>

      <div className="crm-kpi-grid">
        <Kpi icon="Trophy2" label={`Won ${q.label}`} value={money(wonQv)} delta={wonPQv ? ((wonQv - wonPQv) / wonPQv) * 100 : undefined} hint={`vs ${money(wonPQv)} last quarter`} onClick={() => navigate('forecast')} />
        <Kpi icon="Money" label="Open pipeline" value={money(sum(open_, o => engine.baseAmount(o)))} hint={`${open_.length} deals · weighted ${money(sum(open_, o => (engine.baseAmount(o) * (o.probability || 0)) / 100))}`} onClick={() => navigate('entity:opportunities')} />
        <Kpi icon="People" label={`New leads ${q.label}`} value={leadsQ} delta={leadsPQ ? ((leadsQ - leadsPQ) / leadsPQ) * 100 : undefined} hint={`vs ${leadsPQ} at this point last quarter · ${leads.filter(l => l.status === 'New').length} still new`} onClick={() => navigate('entity:leads')} />
        <Kpi icon="Presentation" label="Demos to schedule" value={demos.filter(d => d.status === 'Requested').length} hint={`${upcoming.length} upcoming`} onClick={() => navigate('entity:demos')} />
        <Kpi icon="Org" label="Win rate" value={`${winRate}%`} hint={`${closed.length} closed deals`} onClick={() => navigate('reports')} />
        {liveSubs.length > 0 && <Kpi icon="RecurringEvent" label="MRR" value={money(mrr)} hint={`ARR ${money(mrr * 12)} · ${renewSoon.length} renewal${renewSoon.length === 1 ? '' : 's'} due`} onClick={() => navigate('entity:subscriptions')} />}
        {(unpaid.length > 0 || sc.invoices.length > 0) && <Kpi icon="Money" label="Unpaid invoices" value={money(sum(unpaid, x => engine.toBase((x.grandTotal || 0) - (x.amountPaid || 0), x.currency)))} hint={overdueInv.length ? `${overdueInv.length} overdue` : `${unpaid.length} open`} onClick={() => navigate('entity:invoices')} />}
        {engine.hasFeature('cases') && <Kpi icon="Headset" label="Open cases" value={openCases.length} hint={breached.length ? `${breached.length} SLA breached` : 'all within SLA'} onClick={() => navigate('entity:cases')} />}
      </div>

      <Section title="Focus">
        <div className="crm-dash-grid">
          <Card title="Next best actions" ai wide action={actions.length > 6 ? <Link onClick={() => setAllActions(!allActions)}>{allActions ? 'Show fewer' : `Show all ${actions.length}`}</Link> : undefined}>
            {shown.length === 0 && <Text className="crm-muted">You are all caught up.</Text>}
            {shown.map((a, i) => (
              <Row key={`${a.key}${a.r.id}${i}`} onOpen={() => open(a.key, a.r)}
                aside={a.canDraft ? <DefaultButton iconProps={{ iconName: 'Edit' }} text="Draft email" ariaLabel={`Draft email: ${a.title}`} onClick={() => setDraft({ key: a.key, r: a.r })} styles={{ root: { minWidth: 0, padding: '0 8px' } }} /> : undefined}>
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                  <span className={`crm-prio crm-prio-${a.rank >= 70 ? 'high' : a.rank >= 50 ? 'med' : 'low'}`} aria-label={a.rank >= 70 ? 'High priority' : a.rank >= 50 ? 'Medium priority' : 'Low priority'} />
                  <Text styles={{ root: { fontWeight: 600 } }}>{a.title}</Text>
                </Stack>
                <Text variant="small" block className="crm-muted" styles={{ root: { paddingLeft: 16 } }}>{a.why}</Text>
              </Row>
            ))}
          </Card>
          <Card title="Deals that need attention" ai action={<Link onClick={() => ask('Deals that need attention')}>List them</Link>}>
            <BarList data={healthMix} format={money} />
            <div style={{ marginTop: 8 }}>
              {watchDeals.length === 0 && <Text className="crm-muted">Every open deal looks healthy.</Text>}
              {watchDeals.map(o => (
                <Row key={o.id} onOpen={() => open('opportunities', o)} aside={<HealthBadge h={health(o)} />}>
                  <Text styles={{ root: { fontWeight: 600 } }}>{o.title}</Text>
                  <Text variant="small" block className="crm-muted">{money(engine.baseAmount(o))} · {health(o).risks[0] || o.stage}</Text>
                </Row>
              ))}
            </div>
          </Card>
        </div>
      </Section>

      <Section title="Performance">
        <div className="crm-dash-grid">
          <Card title="Revenue won — last 12 months" wide action={proj.some(p => p.value > 0) ? <span className="crm-legend-proj">Next 3 months: AI projection</span> : undefined}>
            <ColumnChart data={chart} format={money} ariaLabel="Revenue won by month and projected revenue for the next three months" />
          </Card>
          <Card title="Pipeline by stage" action={<Link onClick={() => navigate('entity:opportunities')}>Open board</Link>}>
            <BarList data={byStage} format={money} />
          </Card>
          <Card title="Conversion funnel">
            <Funnel data={funnel} />
          </Card>
          <Card title="Leads by source">
            <Donut data={bySource} ariaLabel="Leads by source" centerLabel="leads" />
          </Card>
          <Card title="Biggest open deals">
            {topDeals.length === 0 && <Text className="crm-muted">No open deals.</Text>}
            {topDeals.map(o => (
              <Row key={o.id} onOpen={() => open('opportunities', o)} aside={<HealthBadge h={health(o)} />}>
                <Stack horizontal horizontalAlign="space-between" tokens={{ childrenGap: 8 }}>
                  <Text styles={{ root: { fontWeight: 600 } }}>{o.title}</Text><Text styles={{ root: { whiteSpace: 'nowrap' } }}>{fmtMoney(o.amount || 0, o.currency || currency)}</Text>
                </Stack>
                <Text variant="small" block className="crm-muted"><Pill text={o.stage} /> close {fmtDate(o.expectedCloseDate)} · {o.owner?.title || 'Unassigned'}</Text>
              </Row>
            ))}
          </Card>
        </div>
      </Section>

      <Section title="Your day">
        <div className="crm-dash-grid">
          <Card title="Hottest leads" action={<Link onClick={() => navigate('entity:leads')}>All leads</Link>}>
            {hotLeads.length === 0 && <Text className="crm-muted">No open leads.</Text>}
            {hotLeads.map(l => (
              <Row key={l.id} onOpen={() => open('leads', l)} aside={<ScoreBadge score={l.leadScore} />}>
                <Text styles={{ root: { fontWeight: 600 } }}>{l.title}</Text>
                <Text variant="small" block className="crm-muted">{l.company} · {l.leadSource} · {fmtDate(l.created)}</Text>
              </Row>
            ))}
          </Card>
          <Card title="Upcoming demos" action={<Link onClick={() => navigate('calendar')}>Calendar</Link>}>
            {upcoming.length === 0 && <Text className="crm-muted">No scheduled demos.</Text>}
            {upcoming.map(d => (
              <Row key={d.id} onOpen={() => open('demos', d)}>
                <Text styles={{ root: { fontWeight: 600 } }}>{d.title}</Text>
                <Text variant="small" block className="crm-muted">{fmtDate(d.scheduledDate, true)} · {d.owner?.title || 'No presenter'}</Text>
              </Row>
            ))}
          </Card>
          <Card title="My open tasks" action={<Link onClick={() => navigate('mywork')}>My work</Link>}>
            {myTasks.length === 0 && <Text className="crm-muted">Nothing due.</Text>}
            {myTasks.map(a => (
              <Row key={a.id} onOpen={() => open('activities', a)}>
                <Text>{a.title}</Text>
                <Text variant="small" block className="crm-muted">{a.dueDate ? fmtDate(a.dueDate, true) : 'No due date'} {isOverdue(a.dueDate) && <Pill text="Overdue" />}</Text>
              </Row>
            ))}
          </Card>
        </div>
      </Section>

      {draft && (() => {
        const d = draftFollowUp(engine, draft.key, draft.r);
        return <EmailDialog entity={draft.key} record={draft.r} to={d.to} subject={d.subject} body={d.body} onClose={() => setDraft(undefined)} />;
      })()}
    </div>
  );
};
