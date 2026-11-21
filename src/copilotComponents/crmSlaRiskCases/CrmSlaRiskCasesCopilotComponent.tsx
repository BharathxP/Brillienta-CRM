import * as React from 'react';
import type { EntityKey } from '../../webparts/brillientaCrm/models/model';
import CrmCopilotComponentBase from '../shared/CrmCopilotComponentBase';
import { caseSummary, caseView, CaseState, normalizeCases } from '../shared/crmQueries';
import { Badge, Empty, ICrmToolHost, Kpis, Note, openRecord, Row, RowList, Tone, useRowLimit } from '../shared/ui';
import type { ICrmSlaRiskCasesCopilotComponentProperties } from './CrmSlaRiskCasesCopilotComponentProperties';

const STATE: Record<CaseState, { tone: Tone; text: (h?: number) => string }> = {
  breached: { tone: 'bad', text: h => (h !== undefined ? `Breached ${Math.abs(h)} h ago` : 'SLA breached') },
  firstResponseLate: { tone: 'bad', text: () => 'First reply late' },
  dueSoon: { tone: 'warn', text: h => (h !== undefined ? `Due in ${Math.max(0, h)} h` : 'Due soon') },
  paused: { tone: 'info', text: () => 'Paused' },
  ok: { tone: 'good', text: () => 'On track' }
};

export const CasesView: React.FC<{ host: ICrmToolHost; properties: unknown }> = ({ host, properties }) => {
  const { engine } = host.session;
  const v = React.useMemo(() => caseView(engine, normalizeCases(properties)), [engine, properties]);
  React.useEffect(() => {
    host.setModelContext(caseSummary(v), {
      tool: 'GetSlaRiskCases', filter: v.label,
      cases: v.cases.slice(0, 25).map(c => ({ id: c.id, number: c.number, subject: c.title, account: c.account, priority: c.priority, state: c.state, slaDue: c.slaDue, owner: c.owner }))
    });
  }, [v]);
  const { visible, more } = useRowLimit(v.cases, host);
  if (!v.available) {
    return <Note>Support cases are part of the Brillienta CRM Enterprise plan. Ask your CRM administrator about upgrading.</Note>;
  }
  return (
    <>
      <p style={{ margin: '0 0 10px' }}>{v.label}</p>
      <Kpis items={[
        { label: 'SLA breached', value: String(v.breached), tone: v.breached ? 'bad' : 'good' },
        { label: 'At risk', value: String(v.dueSoon), tone: v.dueSoon ? 'warn' : undefined }
      ]} />
      {v.cases.length === 0 ? <Empty>No open cases are breaching or close to their SLA.</Empty> : (
        <RowList label="Cases at risk">
          {visible.map(c => (
            <Row key={c.id} title={`${c.number} ${c.title}`}
              meta={[c.account, c.priority, c.status, c.owner].filter(Boolean).join(' · ')}
              detail={`SLA ${c.slaText}`}
              badge={<Badge tone={STATE[c.state].tone}>{STATE[c.state].text(c.hoursLeft)}</Badge>}
              onOpen={openRecord(host, 'cases', c.id)} />
          ))}
        </RowList>
      )}
      {more}
    </>
  );
};

export default class CrmSlaRiskCasesCopilotComponent extends CrmCopilotComponentBase<ICrmSlaRiskCasesCopilotComponentProperties> {
  protected readonly toolTitle = 'Cases at SLA risk';
  protected readonly lists: EntityKey[] = ['cases', 'accounts'];
  protected renderView(host: ICrmToolHost, properties: unknown): React.ReactNode {
    return <CasesView host={host} properties={properties} />;
  }
}
