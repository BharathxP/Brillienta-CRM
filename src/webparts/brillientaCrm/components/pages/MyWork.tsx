import * as React from 'react';
import { Stack, Text, Link, PrimaryButton } from '@fluentui/react';
import { EntityKey, IRecord, WON, LOST } from '../../models/model';
import { useCrm, useEntities } from '../CrmContext';
import { ActivityRow } from '../core/ActivityTimeline';
import { Pill, ScoreBadge } from '../core/controls';
import { Card } from './Dashboard';
import { OutlookSyncCard } from '../core/OutlookSyncCard';
import { fmtDate, fmtMoney, isOverdue, slaBreached } from '../utils';

/** Personal work queue: tasks, follow-ups, demos, deals closing, cases at risk. */
export const MyWork: React.FC = () => {
  const { engine, open, currency } = useCrm();
  const { rows } = useEntities('activities', 'leads', 'demos', 'opportunities', 'cases');
  const me = engine.me?.id;
  const endOfToday = new Date(); endOfToday.setHours(23, 59, 59, 999);
  const in7 = Date.now() + 7 * 86400000;

  const tasks = rows('activities').filter(a => a.activityType === 'Task' && !a.completed && a.owner?.id === me)
    .sort((a, b) => String(a.dueDate || '9').localeCompare(String(b.dueDate || '9')));
  const overdue = tasks.filter(t => isOverdue(t.dueDate));
  const today = tasks.filter(t => t.dueDate && !isOverdue(t.dueDate) && new Date(t.dueDate).getTime() <= endOfToday.getTime());
  const later = tasks.filter(t => overdue.indexOf(t) < 0 && today.indexOf(t) < 0);
  const followUps = rows('leads').filter(l => l.owner?.id === me && l.status !== 'Converted' && l.status !== 'Unqualified' && l.nextFollowUp && new Date(l.nextFollowUp).getTime() <= in7)
    .sort((a, b) => String(a.nextFollowUp).localeCompare(String(b.nextFollowUp)));
  const newLeads = rows('leads').filter(l => l.owner?.id === me && l.status === 'New');
  const demos = rows('demos').filter(d => d.owner?.id === me && (d.status === 'Requested' || (d.status === 'Scheduled' && d.scheduledDate && new Date(d.scheduledDate).getTime() <= in7)));
  const closing = rows('opportunities').filter(o => o.owner?.id === me && o.stage !== WON && o.stage !== LOST && o.expectedCloseDate && new Date(o.expectedCloseDate).getTime() <= Date.now() + 30 * 86400000)
    .sort((a, b) => String(a.expectedCloseDate).localeCompare(String(b.expectedCloseDate)));
  const cases = rows('cases').filter(c => c.owner?.id === me && c.caseStatus !== 'Resolved' && c.caseStatus !== 'Closed')
    .sort((a, b) => String(a.slaDue || '9').localeCompare(String(b.slaDue || '9')));

  const Row: React.FC<{ k: EntityKey; r: IRecord; sub: React.ReactNode }> = ({ k, r, sub }) => (
    <div className="crm-list-row" role="button" tabIndex={0} onClick={() => open(k, r)} onKeyDown={e => { if (e.key === 'Enter') { open(k, r); } }}>
      <Text styles={{ root: { fontWeight: 600 } }}>{r.title || r.caseNumber}</Text>
      <Text variant="small" className="crm-muted">{sub}</Text>
    </div>
  );
  const total = overdue.length + today.length + followUps.length + newLeads.length + demos.length + cases.length;

  return (
    <Stack tokens={{ childrenGap: 16 }}>
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center" wrap>
        <Text variant="large" styles={{ root: { fontWeight: 600 } }}>{total ? `You have ${total} item(s) that need attention` : 'You are all caught up'}</Text>
        {engine.can('create', 'activities') && <PrimaryButton iconProps={{ iconName: 'Add' }} text="New task" onClick={() => open('activities', { activityType: 'Task', completed: false, owner: engine.me })} />}
      </Stack>
      <OutlookSyncCard />
      <div className="crm-dash-grid">
        <Card title={`Overdue tasks (${overdue.length})`}>{overdue.length ? overdue.map(a => <ActivityRow key={a.id} a={a} showRelated />) : <Text className="crm-muted">None 🎉</Text>}</Card>
        <Card title={`Due today (${today.length})`}>{today.length ? today.map(a => <ActivityRow key={a.id} a={a} showRelated />) : <Text className="crm-muted">Nothing due today.</Text>}</Card>
        <Card title={`Upcoming tasks (${later.length})`}>{later.length ? later.slice(0, 10).map(a => <ActivityRow key={a.id} a={a} showRelated />) : <Text className="crm-muted">No upcoming tasks.</Text>}</Card>
        <Card title={`Lead follow-ups (${followUps.length})`}>
          {followUps.length ? followUps.map(l => <Row key={l.id} k="leads" r={l} sub={<>{fmtDate(l.nextFollowUp, true)} {isOverdue(l.nextFollowUp) && <Pill text="Overdue" />} · {l.company}</>} />) : <Text className="crm-muted">No follow-ups this week.</Text>}
        </Card>
        <Card title={`New leads assigned to me (${newLeads.length})`}>
          {newLeads.length ? newLeads.slice(0, 10).map(l => <Row key={l.id} k="leads" r={l} sub={<>{l.company} · {l.leadSource} · score <ScoreBadge score={l.leadScore} /></>} />) : <Text className="crm-muted">No new leads.</Text>}
        </Card>
        <Card title={`My demos (${demos.length})`}>
          {demos.length ? demos.map(d => <Row key={d.id} k="demos" r={d} sub={<><Pill text={d.status} /> {d.scheduledDate ? fmtDate(d.scheduledDate, true) : `preferred ${fmtDate(d.preferredDate, true) || '—'}`}</>} />) : <Text className="crm-muted">No demos this week.</Text>}
        </Card>
        <Card title={`Deals closing in 30 days (${closing.length})`}>
          {closing.length ? closing.map(o => <Row key={o.id} k="opportunities" r={o} sub={<>{fmtMoney(o.amount, o.currency || currency)} · {o.stage} · {fmtDate(o.expectedCloseDate)} {isOverdue(o.expectedCloseDate) && <Pill text="Overdue" />}</>} />) : <Text className="crm-muted">No deals closing soon.</Text>}
        </Card>
        {engine.hasFeature('cases') && (
          <Card title={`My open cases (${cases.length})`} action={<Link onClick={() => undefined}> </Link>}>
            {cases.length ? cases.map(c => <Row key={c.id} k="cases" r={c} sub={<><Pill text={c.priority} /> SLA {fmtDate(c.slaDue, true)} {slaBreached(c) && <Pill text="Breached" />}{c.slaPausedAt && <span className="crm-muted"> paused</span>}</>} />) : <Text className="crm-muted">No open cases.</Text>}
          </Card>
        )}
      </div>
    </Stack>
  );
};
