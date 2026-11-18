import * as React from 'react';
import { Icon, Link, Stack, Text, Dropdown } from '@fluentui/react';
import { IRecord, LOST, WON } from '../../models/model';
import { useCrm, useEntities } from '../CrmContext';
import { BarList, ColumnChart, Donut, Funnel } from '../core/charts';
import { Pill, ScoreBadge } from '../core/controls';
import { fmtDate, fmtMoney, groupCount, inRange, isOverdue, monthLabel, quarterRange, sum } from '../utils';

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

export const Card: React.FC<{ title: string; action?: React.ReactNode; children?: React.ReactNode; wide?: boolean }> = ({ title, action, children, wide }) => (
  <div className={`crm-card${wide ? ' crm-wide' : ''}`}>
    <Stack horizontal horizontalAlign="space-between" verticalAlign="center" className="crm-card-title">
      <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>{title}</Text>
      {action}
    </Stack>
    {children}
  </div>
);

export const Dashboard: React.FC = () => {
  const { engine, currency, navigate, open, settings } = useCrm();
  const { rows } = useEntities('leads', 'opportunities', 'demos', 'activities', 'cases', 'accounts');
  const [scope, setScope] = React.useState<'all' | 'me'>(engine.role === 'sales' ? 'me' : 'all');
  const mine = (r: IRecord): boolean => scope === 'all' || r.owner?.id === engine.me?.id;
  const leads = rows('leads').filter(mine);
  const opps = rows('opportunities').filter(mine);
  const demos = rows('demos').filter(mine);
  const cases = rows('cases').filter(mine);
  const acts = rows('activities');
  const money = (n: number): string => fmtMoney(n, currency);

  const open_ = opps.filter(o => engine.openStages().indexOf(o.stage) >= 0);
  const q = quarterRange(0, settings.fiscalYearStartMonth);
  const pq = quarterRange(-1, settings.fiscalYearStartMonth);
  const wonQ = opps.filter(o => o.stage === WON && inRange(o.actualCloseDate, q.start, q.end));
  const wonPQ = opps.filter(o => o.stage === WON && inRange(o.actualCloseDate, pq.start, pq.end));
  const wonQv = sum(wonQ, o => o.amount);
  const wonPQv = sum(wonPQ, o => o.amount);
  const closed = opps.filter(o => o.stage === WON || o.stage === LOST);
  const winRate = closed.length ? Math.round((closed.filter(o => o.stage === WON).length / closed.length) * 100) : 0;
  const leadsQ = leads.filter(l => inRange(l.created, q.start, q.end)).length;
  const leadsPQ = leads.filter(l => inRange(l.created, pq.start, pq.end)).length;
  const openCases = cases.filter(c => c.caseStatus !== 'Resolved' && c.caseStatus !== 'Closed');
  const breached = openCases.filter(c => isOverdue(c.slaDue));

  // revenue won by month, last 12 months
  const months: { label: string; value: number; short?: string }[] = [];
  const now = new Date();
  for (let i = 11; i >= 0; i--) {
    const s = new Date(now.getFullYear(), now.getMonth() - i, 1);
    const e = new Date(now.getFullYear(), now.getMonth() - i + 1, 1);
    months.push({ label: monthLabel(s), short: s.toLocaleDateString(undefined, { month: 'short' }), value: sum(opps.filter(o => o.stage === WON && inRange(o.actualCloseDate, s, e)), o => o.amount) });
  }

  const byStage = engine.openStages().map(st => ({ label: st, value: sum(open_.filter(o => o.stage === st), o => o.amount), hint: `${open_.filter(o => o.stage === st).length} deals` }));
  const bySource = groupCount(leads, l => l.leadSource);
  const funnel = [
    { label: 'Leads', value: leads.length },
    { label: 'Qualified', value: leads.filter(l => ['Qualified', 'Converted'].indexOf(l.status) >= 0).length },
    { label: 'Demos', value: demos.filter(d => d.status !== 'Cancelled').length },
    { label: 'Opportunities', value: opps.length },
    { label: 'Won', value: opps.filter(o => o.stage === WON).length }
  ];
  const myTasks = acts.filter(a => a.activityType === 'Task' && !a.completed && a.owner?.id === engine.me?.id)
    .sort((a, b) => String(a.dueDate || '9').localeCompare(String(b.dueDate || '9'))).slice(0, 6);
  const upcoming = demos.filter(d => d.status === 'Scheduled' && d.scheduledDate && new Date(d.scheduledDate).getTime() > Date.now() - 3600000)
    .sort((a, b) => String(a.scheduledDate).localeCompare(String(b.scheduledDate))).slice(0, 5);
  const hotLeads = leads.filter(l => l.status !== 'Converted' && l.status !== 'Unqualified').sort((a, b) => (b.leadScore || 0) - (a.leadScore || 0)).slice(0, 6);
  const topDeals = [...open_].sort((a, b) => (b.amount || 0) - (a.amount || 0)).slice(0, 6);

  return (
    <Stack tokens={{ childrenGap: 16 }}>
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center" wrap>
        <Text variant="large" styles={{ root: { fontWeight: 600 } }}>Good {now.getHours() < 12 ? 'morning' : now.getHours() < 18 ? 'afternoon' : 'evening'}, {(engine.me?.title || '').split(' ')[0]}</Text>
        <Dropdown selectedKey={scope} ariaLabel="Scope" styles={{ root: { width: 160 } }} onChange={(_, o) => setScope(o?.key as 'all' | 'me')}
          options={[{ key: 'all', text: 'Whole team' }, { key: 'me', text: 'My records' }]} />
      </Stack>
      <div className="crm-kpi-grid">
        <Kpi icon="Trophy2" label={`Won ${q.label}`} value={money(wonQv)} delta={wonPQv ? ((wonQv - wonPQv) / wonPQv) * 100 : undefined} hint={`vs ${money(wonPQv)} last quarter`} onClick={() => navigate('forecast')} />
        <Kpi icon="Money" label="Open pipeline" value={money(sum(open_, o => o.amount))} hint={`${open_.length} deals · weighted ${money(sum(open_, o => (o.amount * (o.probability || 0)) / 100))}`} onClick={() => navigate('entity:opportunities')} />
        <Kpi icon="People" label={`New leads ${q.label}`} value={leadsQ} delta={leadsPQ ? ((leadsQ - leadsPQ) / leadsPQ) * 100 : undefined} hint={`${leads.filter(l => l.status === 'New').length} still new`} onClick={() => navigate('entity:leads')} />
        <Kpi icon="Presentation" label="Demos to schedule" value={demos.filter(d => d.status === 'Requested').length} hint={`${upcoming.length} upcoming`} onClick={() => navigate('entity:demos')} />
        <Kpi icon="Org" label="Win rate" value={`${winRate}%`} hint={`${closed.length} closed deals`} onClick={() => navigate('reports')} />
        {engine.hasFeature('cases') && <Kpi icon="Headset" label="Open cases" value={openCases.length} hint={breached.length ? `${breached.length} SLA breached` : 'all within SLA'} onClick={() => navigate('entity:cases')} />}
      </div>

      <div className="crm-dash-grid">
        <Card title="Revenue won — last 12 months" wide>
          <ColumnChart data={months} format={money} ariaLabel="Revenue won by month" />
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
        <Card title="Hottest leads" action={<Link onClick={() => navigate('entity:leads')}>All leads</Link>}>
          {hotLeads.length === 0 && <Text className="crm-muted">No open leads.</Text>}
          {hotLeads.map(l => (
            <div key={l.id} className="crm-list-row" role="button" tabIndex={0} onClick={() => open('leads', l)} onKeyDown={e => { if (e.key === 'Enter') { open('leads', l); } }}>
              <Stack horizontal horizontalAlign="space-between" tokens={{ childrenGap: 8 }}>
                <Text styles={{ root: { fontWeight: 600 } }}>{l.title}</Text><ScoreBadge score={l.leadScore} />
              </Stack>
              <Text variant="small" className="crm-muted">{l.company} · {l.leadSource} · {fmtDate(l.created)}</Text>
            </div>
          ))}
        </Card>
        <Card title="Biggest open deals">
          {topDeals.length === 0 && <Text className="crm-muted">No open deals.</Text>}
          {topDeals.map(o => (
            <div key={o.id} className="crm-list-row" role="button" tabIndex={0} onClick={() => open('opportunities', o)} onKeyDown={e => { if (e.key === 'Enter') { open('opportunities', o); } }}>
              <Stack horizontal horizontalAlign="space-between" tokens={{ childrenGap: 8 }}>
                <Text styles={{ root: { fontWeight: 600 } }}>{o.title}</Text><Text>{money(o.amount)}</Text>
              </Stack>
              <Text variant="small" className="crm-muted"><Pill text={o.stage} /> {o.probability}% · close {fmtDate(o.expectedCloseDate)} · {o.owner?.title || 'Unassigned'}</Text>
            </div>
          ))}
        </Card>
        <Card title="Upcoming demos" action={<Link onClick={() => navigate('calendar')}>Calendar</Link>}>
          {upcoming.length === 0 && <Text className="crm-muted">No scheduled demos.</Text>}
          {upcoming.map(d => (
            <div key={d.id} className="crm-list-row" role="button" tabIndex={0} onClick={() => open('demos', d)} onKeyDown={e => { if (e.key === 'Enter') { open('demos', d); } }}>
              <Text styles={{ root: { fontWeight: 600 } }}>{d.title}</Text>
              <Text variant="small" className="crm-muted">{fmtDate(d.scheduledDate, true)} · {d.owner?.title || 'No presenter'}</Text>
            </div>
          ))}
        </Card>
        <Card title="My open tasks" action={<Link onClick={() => navigate('mywork')}>My work</Link>}>
          {myTasks.length === 0 && <Text className="crm-muted">Nothing due.</Text>}
          {myTasks.map(a => (
            <div key={a.id} className="crm-list-row" role="button" tabIndex={0} onClick={() => open('activities', a)} onKeyDown={e => { if (e.key === 'Enter') { open('activities', a); } }}>
              <Text>{a.title}</Text>
              <Text variant="small" className="crm-muted">{a.dueDate ? fmtDate(a.dueDate, true) : 'No due date'} {isOverdue(a.dueDate) && <Pill text="Overdue" />}</Text>
            </div>
          ))}
        </Card>
      </div>
    </Stack>
  );
};
