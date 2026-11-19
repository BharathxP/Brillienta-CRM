import * as React from 'react';
import {
  Checkbox, DefaultButton, DetailsList, DetailsListLayoutMode, Dropdown, IColumn, IconButton, SelectionMode, Stack, Text
} from '@fluentui/react';
import { EntityKey, IRecord, LOST, WON } from '../../models/model';
import { useCrm, useEntities } from '../CrmContext';
import { BarList, ColumnChart, Donut, Funnel, Meter } from '../core/charts';
import { Card } from './Dashboard';
import { exportCsv, fmtMoney, groupCount, inRange, isOverdue, monthLabel, quarterRange, sum } from '../utils';

/* ================================================================ Calendar */

interface ICalItem { key: EntityKey; r: IRecord; date: Date; label: string; kind: string }

export const CalendarPage: React.FC = () => {
  const { engine, open } = useCrm();
  const { rows } = useEntities('demos', 'activities', 'opportunities', 'cases');
  const [month, setMonth] = React.useState(() => { const d = new Date(); return new Date(d.getFullYear(), d.getMonth(), 1); });
  const [show, setShow] = React.useState<Record<string, boolean>>({ Demo: true, Task: true, Close: true, SLA: engine.hasFeature('cases') });
  const items: ICalItem[] = [];
  rows('demos').forEach(d => { if (d.scheduledDate && d.status !== 'Cancelled') { items.push({ key: 'demos', r: d, date: new Date(d.scheduledDate), label: d.title || '', kind: 'Demo' }); } });
  rows('activities').forEach(a => { if (a.activityType === 'Task' && a.dueDate && !a.completed) { items.push({ key: 'activities', r: a, date: new Date(a.dueDate), label: a.title || '', kind: 'Task' }); } });
  rows('opportunities').forEach(o => { if (o.expectedCloseDate && o.stage !== WON && o.stage !== LOST) { items.push({ key: 'opportunities', r: o, date: new Date(o.expectedCloseDate), label: o.title || '', kind: 'Close' }); } });
  rows('cases').forEach(c => { if (c.slaDue && c.caseStatus !== 'Resolved' && c.caseStatus !== 'Closed') { items.push({ key: 'cases', r: c, date: new Date(c.slaDue), label: `${c.caseNumber} ${c.title}`, kind: 'SLA' }); } });

  const first = new Date(month);
  const startOffset = (first.getDay() + 6) % 7; // Monday first
  const gridStart = new Date(first.getFullYear(), first.getMonth(), 1 - startOffset);
  const days = Array.from({ length: 42 }, (_, i) => new Date(gridStart.getFullYear(), gridStart.getMonth(), gridStart.getDate() + i));
  const same = (a: Date, b: Date): boolean => a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
  const today = new Date();

  return (
    <Stack tokens={{ childrenGap: 10 }}>
      <Stack horizontal verticalAlign="center" wrap tokens={{ childrenGap: 8 }}>
        <IconButton iconProps={{ iconName: 'ChevronLeft' }} ariaLabel="Previous month" onClick={() => setMonth(new Date(month.getFullYear(), month.getMonth() - 1, 1))} />
        <Text variant="large" styles={{ root: { fontWeight: 600, minWidth: 170, textAlign: 'center' } }}>{month.toLocaleDateString(undefined, { month: 'long', year: 'numeric' })}</Text>
        <IconButton iconProps={{ iconName: 'ChevronRight' }} ariaLabel="Next month" onClick={() => setMonth(new Date(month.getFullYear(), month.getMonth() + 1, 1))} />
        <DefaultButton text="Today" onClick={() => setMonth(new Date(today.getFullYear(), today.getMonth(), 1))} />
        {Object.keys(show).map(k => (
          <Checkbox key={k} label={k === 'Close' ? 'Expected closes' : k === 'SLA' ? 'Case SLAs' : `${k}s`} checked={show[k]}
            onRenderLabel={p => <span className={`crm-cal-legend crm-cal-${k}`}>{p?.label}</span>} onChange={(_, c) => setShow({ ...show, [k]: !!c })} />
        ))}
      </Stack>
      <div className="crm-cal">
        {['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'].map(d => <div key={d} className="crm-cal-head">{d}</div>)}
        {days.map(day => {
          const its = items.filter(i => show[i.kind] && same(i.date, day)).sort((a, b) => a.date.getTime() - b.date.getTime());
          return (
            <div key={day.toISOString()} className={`crm-cal-day${day.getMonth() !== month.getMonth() ? ' crm-cal-out' : ''}${same(day, today) ? ' crm-cal-today' : ''}`}>
              <span className="crm-cal-num">{day.getDate()}</span>
              {its.slice(0, 4).map(i => (
                <button key={`${i.key}-${i.r.id}`} className={`crm-cal-item crm-cal-${i.kind}`} title={`${i.kind}: ${i.label}`} onClick={() => open(i.key, i.r)}>
                  {i.kind === 'Demo' || i.kind === 'Task' ? `${i.date.toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' })} ` : ''}{i.label}
                </button>
              ))}
              {its.length > 4 && <span className="crm-muted" style={{ fontSize: 11 }}>+{its.length - 4} more</span>}
            </div>
          );
        })}
      </div>
    </Stack>
  );
};

/* ================================================================ Forecast */

export const ForecastPage: React.FC = () => {
  const { engine, currency, settings, open } = useCrm();
  const { rows } = useEntities('opportunities');
  const [offset, setOffset] = React.useState(0);
  const q = quarterRange(offset, settings.fiscalYearStartMonth);
  const money = (n: number): string => fmtMoney(n, currency);
  const opps = rows('opportunities');
  const inQ = (o: IRecord): boolean => inRange(o.stage === WON ? o.actualCloseDate : o.expectedCloseDate, q.start, q.end);
  const qOpps = opps.filter(o => inQ(o) && o.stage !== LOST);

  const owners: { id: number; title: string }[] = [];
  qOpps.forEach(o => { if (o.owner && !owners.some(x => x.id === o.owner.id)) { owners.push(o.owner); } });
  settings.quotas.forEach(qt => { if (!owners.some(x => x.id === qt.user.id)) { owners.push(qt.user); } });

  const calc = (list: IRecord[]): { won: number; commit: number; best: number; pipeline: number; weighted: number } => ({
    won: sum(list.filter(o => o.stage === WON), o => o.amount),
    commit: sum(list.filter(o => o.stage !== WON && o.forecastCategory === 'Commit'), o => o.amount),
    best: sum(list.filter(o => o.stage !== WON && o.forecastCategory === 'Best case'), o => o.amount),
    pipeline: sum(list.filter(o => o.stage !== WON && (o.forecastCategory === 'Pipeline' || !o.forecastCategory)), o => o.amount),
    weighted: sum(list.filter(o => o.stage !== WON), o => (o.amount * (o.probability || 0)) / 100)
  });
  const team = calc(qOpps);
  const teamQuota = sum(settings.quotas, x => x.quarterly);

  const rowsData = owners.map(u => {
    const c = calc(qOpps.filter(o => o.owner?.id === u.id));
    const quota = settings.quotas.find(x => x.user.id === u.id)?.quarterly || 0;
    return { key: String(u.id), owner: u.title, quota, ...c, projected: c.won + c.commit };
  }).sort((a, b) => b.won - a.won);
  const unassigned = calc(qOpps.filter(o => !o.owner));
  if (unassigned.won + unassigned.commit + unassigned.best + unassigned.pipeline > 0) {
    rowsData.push({ key: 'none', owner: 'Unassigned', quota: 0, ...unassigned, projected: unassigned.won + unassigned.commit });
  }

  const cols: IColumn[] = [
    { key: 'owner', name: 'Owner', fieldName: 'owner', minWidth: 140 },
    { key: 'quota', name: 'Quota', fieldName: 'quota', minWidth: 90, onRender: r => (r.quota ? money(r.quota) : '—') },
    { key: 'won', name: 'Closed won', fieldName: 'won', minWidth: 100, onRender: r => <b>{money(r.won)}</b> },
    { key: 'commit', name: 'Commit', fieldName: 'commit', minWidth: 90, onRender: r => money(r.commit) },
    { key: 'best', name: 'Best case', fieldName: 'best', minWidth: 90, onRender: r => money(r.best) },
    { key: 'pipeline', name: 'Pipeline', fieldName: 'pipeline', minWidth: 90, onRender: r => money(r.pipeline) },
    { key: 'att', name: 'Attainment', minWidth: 140, onRender: r => <Meter value={r.won} target={r.quota} label={`${money(r.won)} of ${money(r.quota)}`} /> },
    { key: 'proj', name: 'Won + commit vs quota', minWidth: 160, onRender: r => <Meter value={r.projected} target={r.quota} /> }
  ];

  const months = [0, 1, 2].map(i => {
    const s = new Date(q.start.getFullYear(), q.start.getMonth() + i, 1);
    const e = new Date(q.start.getFullYear(), q.start.getMonth() + i + 1, 1);
    return { label: monthLabel(s), short: s.toLocaleDateString(undefined, { month: 'short' }), value: sum(qOpps.filter(o => inRange(o.stage === WON ? o.actualCloseDate : o.expectedCloseDate, s, e)), o => o.amount) };
  });
  const slipping = opps.filter(o => engine.openStages().indexOf(o.stage) >= 0 && isOverdue(o.expectedCloseDate));

  return (
    <Stack tokens={{ childrenGap: 16 }}>
      <Stack horizontal verticalAlign="center" wrap tokens={{ childrenGap: 8 }}>
        <IconButton iconProps={{ iconName: 'ChevronLeft' }} ariaLabel="Previous quarter" onClick={() => setOffset(offset - 1)} />
        <Text variant="large" styles={{ root: { fontWeight: 600 } }}>{q.label} · {q.start.toLocaleDateString()} – {new Date(q.end.getTime() - 1).toLocaleDateString()}</Text>
        <IconButton iconProps={{ iconName: 'ChevronRight' }} ariaLabel="Next quarter" onClick={() => setOffset(offset + 1)} />
        {offset !== 0 && <DefaultButton text="This quarter" onClick={() => setOffset(0)} />}
        <DefaultButton iconProps={{ iconName: 'ExcelDocument' }} text="Export" onClick={() => exportCsv(`forecast-${q.label}.csv`, rowsData.map(r => ({
          Owner: r.owner, Quota: r.quota, ClosedWon: r.won, Commit: r.commit, BestCase: r.best, Pipeline: r.pipeline, Weighted: Math.round(r.weighted)
        })))} />
      </Stack>
      <div className="crm-kpi-grid">
        <div className="crm-kpi"><div className="crm-kpi-value">{money(team.won)}</div><div className="crm-kpi-label">Closed won</div>
          {teamQuota > 0 && <Meter value={team.won} target={teamQuota} />}</div>
        <div className="crm-kpi"><div className="crm-kpi-value">{money(team.won + team.commit)}</div><div className="crm-kpi-label">Won + commit (forecast)</div></div>
        <div className="crm-kpi"><div className="crm-kpi-value">{money(team.won + team.commit + team.best)}</div><div className="crm-kpi-label">Upside (incl. best case)</div></div>
        <div className="crm-kpi"><div className="crm-kpi-value">{money(team.weighted)}</div><div className="crm-kpi-label">Weighted open pipeline</div></div>
        <div className="crm-kpi"><div className="crm-kpi-value">{teamQuota ? money(teamQuota) : '—'}</div><div className="crm-kpi-label">Team quota</div><div className="crm-kpi-hint">Set quotas in Settings › Sales</div></div>
      </div>
      <div className="crm-dash-grid">
        <Card title="Forecast by rep" wide>
          <DetailsList items={rowsData} columns={cols} selectionMode={SelectionMode.none} layoutMode={DetailsListLayoutMode.justified} />
          {!rowsData.length && <Text className="crm-muted">No deals closing in this quarter.</Text>}
        </Card>
        <Card title="Expected / won by month"><ColumnChart data={months} format={money} ariaLabel="Forecast by month" height={140} /></Card>
        <Card title="Forecast categories">
          <BarList data={[{ label: 'Closed won', value: team.won }, { label: 'Commit', value: team.commit }, { label: 'Best case', value: team.best }, { label: 'Pipeline', value: team.pipeline }]} format={money} />
        </Card>
        <Card title={`Slipping deals (${slipping.length})`}>
          {slipping.length === 0 && <Text className="crm-muted">No open deals past their close date.</Text>}
          {slipping.slice(0, 8).map(o => (
            <div key={o.id} className="crm-list-row" role="button" tabIndex={0} onClick={() => open('opportunities', o)} onKeyDown={e => { if (e.key === 'Enter') { open('opportunities', o); } }}>
              <Text styles={{ root: { fontWeight: 600 } }}>{o.title}</Text>
              <Text variant="small" className="crm-muted">{money(o.amount)} · was due {new Date(o.expectedCloseDate).toLocaleDateString()} · {o.owner?.title || 'Unassigned'}</Text>
            </div>
          ))}
        </Card>
      </div>
    </Stack>
  );
};

/* ================================================================ Reports */

const RANGES = [
  { key: 'q', text: 'This quarter' }, { key: 'pq', text: 'Last quarter' }, { key: 'ytd', text: 'Last 12 months' }, { key: 'all', text: 'All time' }
];

export const ReportsPage: React.FC = () => {
  const { engine, currency, settings } = useCrm();
  const { rows } = useEntities('leads', 'opportunities', 'demos', 'cases', 'campaigns', 'quotes', 'activities');
  const [range, setRange] = React.useState('ytd');
  const money = (n: number): string => fmtMoney(n, currency);
  const r = ((): { start: Date; end: Date } => {
    const now = new Date();
    if (range === 'q') { return quarterRange(0, settings.fiscalYearStartMonth); }
    if (range === 'pq') { return quarterRange(-1, settings.fiscalYearStartMonth); }
    if (range === 'ytd') { return { start: new Date(now.getFullYear() - 1, now.getMonth(), now.getDate()), end: new Date(now.getTime() + 86400000) }; }
    return { start: new Date(2000, 0, 1), end: new Date(2100, 0, 1) };
  })();
  const created = (x: IRecord): boolean => inRange(x.created, r.start, r.end);
  const closedIn = (o: IRecord): boolean => inRange(o.actualCloseDate, r.start, r.end);

  const leads = rows('leads').filter(created);
  const opps = rows('opportunities');
  const closed = opps.filter(o => (o.stage === WON || o.stage === LOST) && closedIn(o));
  const won = closed.filter(o => o.stage === WON);
  const lost = closed.filter(o => o.stage === LOST);
  const cycle = won.length ? Math.round(sum(won, o => (new Date(o.actualCloseDate).getTime() - new Date(String(o.created)).getTime()) / 86400000) / won.length) : 0;

  // owner leaderboard
  const ownerMap: Record<string, { owner: string; won: number; deals: number; lost: number; pipeline: number; activities: number }> = {};
  const ensure = (name: string): { owner: string; won: number; deals: number; lost: number; pipeline: number; activities: number } =>
    (ownerMap[name] = ownerMap[name] || { owner: name, won: 0, deals: 0, lost: 0, pipeline: 0, activities: 0 });
  won.forEach(o => { const e = ensure(o.owner?.title || 'Unassigned'); e.won += o.amount || 0; e.deals++; });
  lost.forEach(o => { ensure(o.owner?.title || 'Unassigned').lost++; });
  opps.filter(o => engine.openStages().indexOf(o.stage) >= 0).forEach(o => { ensure(o.owner?.title || 'Unassigned').pipeline += o.amount || 0; });
  rows('activities').filter(a => created(a) && a.activityType !== 'Note').forEach(a => { if (a.author) { ensure(a.author.title).activities++; } });
  const board = Object.keys(ownerMap).map(k => ownerMap[k]).sort((a, b) => b.won - a.won);

  // source performance
  const sources = groupCount(leads, l => l.leadSource).map(s => {
    const ls = leads.filter(l => (l.leadSource || '(none)') === s.label);
    const ids = ls.map(l => l.id);
    const rev = sum(won.filter(o => o.leadRef && ids.indexOf(o.leadRef) >= 0), o => o.amount);
    return { key: s.label, source: s.label, leads: s.value, converted: ls.filter(l => l.status === 'Converted').length, revenue: rev };
  });

  // campaign ROI
  const campaigns = rows('campaigns').map(c => {
    const cl = rows('leads').filter(l => l.campaignRef === c.id).length;
    const rev = sum(opps.filter(o => o.campaignRef === c.id && o.stage === WON), o => o.amount);
    const cost = Number(c.actualCost || c.budget || 0);
    return { key: String(c.id), name: c.title, leads: cl, revenue: rev, cost, roi: cost ? Math.round(((rev - cost) / cost) * 100) : undefined };
  });

  // cases
  const cs = rows('cases').filter(created);
  const resolved = cs.filter(c => c.resolvedDate);
  const slaMet = resolved.filter(c => !c.slaDue || new Date(c.resolvedDate).getTime() <= new Date(c.slaDue).getTime()).length;
  const avgRes = resolved.length ? Math.round(sum(resolved, c => (new Date(c.resolvedDate).getTime() - new Date(String(c.created)).getTime()) / 3600000) / resolved.length) : 0;

  const demos = rows('demos').filter(created);
  const quotes = rows('quotes').filter(created);

  return (
    <Stack tokens={{ childrenGap: 16 }}>
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} wrap>
        <Dropdown selectedKey={range} options={RANGES} onChange={(_, o) => setRange(String(o?.key))} styles={{ root: { width: 180 } }} ariaLabel="Date range" />
        <DefaultButton iconProps={{ iconName: 'ExcelDocument' }} text="Export leaderboard" onClick={() => exportCsv('sales-leaderboard.csv', board.map(b => ({ ...b })))} />
      </Stack>
      <div className="crm-kpi-grid">
        <div className="crm-kpi"><div className="crm-kpi-value">{money(sum(won, o => o.amount))}</div><div className="crm-kpi-label">Revenue won</div><div className="crm-kpi-hint">{won.length} deals</div></div>
        <div className="crm-kpi"><div className="crm-kpi-value">{closed.length ? Math.round((won.length / closed.length) * 100) : 0}%</div><div className="crm-kpi-label">Win rate</div><div className="crm-kpi-hint">{lost.length} lost</div></div>
        <div className="crm-kpi"><div className="crm-kpi-value">{money(won.length ? sum(won, o => o.amount) / won.length : 0)}</div><div className="crm-kpi-label">Average deal size</div></div>
        <div className="crm-kpi"><div className="crm-kpi-value">{cycle} days</div><div className="crm-kpi-label">Average sales cycle</div></div>
        <div className="crm-kpi"><div className="crm-kpi-value">{leads.length}</div><div className="crm-kpi-label">Leads created</div><div className="crm-kpi-hint">{leads.length ? Math.round((leads.filter(l => l.status === 'Converted').length / leads.length) * 100) : 0}% converted</div></div>
        <div className="crm-kpi"><div className="crm-kpi-value">{demos.filter(d => d.status === 'Completed').length}/{demos.length}</div><div className="crm-kpi-label">Demos completed</div><div className="crm-kpi-hint">{demos.filter(d => d.status === 'No-show').length} no-shows</div></div>
      </div>
      <div className="crm-dash-grid">
        <Card title="Sales leaderboard" wide>
          <DetailsList items={board} selectionMode={SelectionMode.none} layoutMode={DetailsListLayoutMode.justified} columns={[
            { key: 'o', name: 'Owner', fieldName: 'owner', minWidth: 90, maxWidth: 160 },
            { key: 'w', name: 'Won', fieldName: 'won', minWidth: 70, maxWidth: 110, onRender: x => <b>{money(x.won)}</b> },
            { key: 'd', name: 'Deals', fieldName: 'deals', minWidth: 40, maxWidth: 60 },
            { key: 'l', name: 'Lost', fieldName: 'lost', minWidth: 36, maxWidth: 50 },
            { key: 'wr', name: 'Win %', minWidth: 44, maxWidth: 60, onRender: x => (x.deals + x.lost ? `${Math.round((x.deals / (x.deals + x.lost)) * 100)}%` : '—') },
            { key: 'p', name: 'Pipeline', fieldName: 'pipeline', minWidth: 70, maxWidth: 110, onRender: x => money(x.pipeline) },
            { key: 'a', name: 'Activities', fieldName: 'activities', minWidth: 70, maxWidth: 90 }
          ]} />
        </Card>
        <Card title="Win / loss reasons"><Donut data={groupCount(lost, o => o.lostReason || 'Not given')} ariaLabel="Lost reasons" centerLabel="lost deals" /></Card>
        <Card title="Lead conversion funnel">
          <Funnel data={[
            { label: 'Leads', value: leads.length },
            { label: 'Contacted', value: leads.filter(l => l.status !== 'New').length },
            { label: 'Qualified', value: leads.filter(l => l.status === 'Qualified' || l.status === 'Converted').length },
            { label: 'Converted', value: leads.filter(l => l.status === 'Converted').length }
          ]} />
        </Card>
        <Card title="Lead sources" wide>
          <DetailsList items={sources} selectionMode={SelectionMode.none} layoutMode={DetailsListLayoutMode.justified} columns={[
            { key: 's', name: 'Source', fieldName: 'source', minWidth: 140 },
            { key: 'l', name: 'Leads', fieldName: 'leads', minWidth: 60 },
            { key: 'c', name: 'Converted', fieldName: 'converted', minWidth: 76 },
            { key: 'r', name: 'Conv. rate', minWidth: 70, onRender: x => `${x.leads ? Math.round((x.converted / x.leads) * 100) : 0}%` },
            { key: 'v', name: 'Revenue won', fieldName: 'revenue', minWidth: 100, onRender: x => money(x.revenue) }
          ]} />
        </Card>
        <Card title="Pipeline by owner"><BarList data={board.filter(b => b.pipeline > 0).map(b => ({ label: b.owner, value: b.pipeline }))} format={money} /></Card>
        <Card title="Quotes">
          <BarList data={['Draft', 'Sent', 'Accepted', 'Rejected', 'Expired'].map(s => ({ label: s, value: quotes.filter(x => x.quoteStatus === s).length }))} />
          <Text variant="small" className="crm-muted" block style={{ marginTop: 8 }}>Accepted value: {money(sum(quotes.filter(x => x.quoteStatus === 'Accepted'), x => x.grandTotal))}</Text>
        </Card>
        {engine.hasFeature('campaigns') && (
          <Card title="Campaign ROI" wide>
            <DetailsList items={campaigns} selectionMode={SelectionMode.none} layoutMode={DetailsListLayoutMode.justified} columns={[
              { key: 'n', name: 'Campaign', fieldName: 'name', minWidth: 160 },
              { key: 'l', name: 'Leads', fieldName: 'leads', minWidth: 60 },
              { key: 'c', name: 'Cost', fieldName: 'cost', minWidth: 90, onRender: x => money(x.cost) },
              { key: 'r', name: 'Revenue won', fieldName: 'revenue', minWidth: 100, onRender: x => money(x.revenue) },
              { key: 'roi', name: 'ROI', fieldName: 'roi', minWidth: 60, onRender: x => (x.roi === undefined ? '—' : `${x.roi}%`) }
            ]} />
            {!campaigns.length && <Text className="crm-muted">No campaigns yet.</Text>}
          </Card>
        )}
        {engine.hasFeature('cases') && (
          <Card title="Service performance">
            <BarList data={['Critical', 'High', 'Normal', 'Low'].map(p => ({ label: p, value: cs.filter(c => c.priority === p).length }))} />
            <Text variant="small" block style={{ marginTop: 8 }}>SLA met: <b>{resolved.length ? Math.round((slaMet / resolved.length) * 100) : 0}%</b> · Avg. resolution: <b>{avgRes} h</b> · Open: <b>{cs.filter(c => !c.resolvedDate).length}</b></Text>
          </Card>
        )}
      </div>
    </Stack>
  );
};
