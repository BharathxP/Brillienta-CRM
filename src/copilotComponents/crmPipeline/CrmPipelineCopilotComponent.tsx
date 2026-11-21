import * as React from 'react';
import type { EntityKey } from '../../webparts/brillientaCrm/models/model';
import CrmCopilotComponentBase from '../shared/CrmCopilotComponentBase';
import { normalizePipeline, pipelineSummary, pipelineView } from '../shared/crmQueries';
import { Empty, HealthBadge, ICrmToolHost, Kpis, Note, openRecord, Row, RowList, useRowLimit } from '../shared/ui';
import { fmtMoney } from '../../webparts/brillientaCrm/components/utils';
import type { ICrmPipelineCopilotComponentProperties } from './CrmPipelineCopilotComponentProperties';

export const PipelineView: React.FC<{ host: ICrmToolHost; properties: unknown }> = ({ host, properties }) => {
  const { engine } = host.session;
  const v = React.useMemo(() => pipelineView(engine, normalizePipeline(properties)), [engine, properties]);
  React.useEffect(() => {
    host.setModelContext(pipelineSummary(engine, v), {
      tool: 'GetMyPipeline', filter: v.label,
      deals: v.deals.slice(0, 25).map(d => ({ id: d.id, name: d.title, account: d.account, stage: d.stage, amount: d.amount, close: d.close, winLikelihood: d.health.score, topRisk: d.health.risks[0] || '' }))
    });
  }, [v]);
  const { visible, more } = useRowLimit(v.deals, host);
  const money = (n: number): string => fmtMoney(n, engine.settings.currency);
  return (
    <>
      <p style={{ margin: '0 0 10px' }}>{v.label}</p>
      {v.note && <Note>{v.note}</Note>}
      <Kpis items={[
        { label: 'Deals', value: String(v.deals.length) },
        { label: 'Total', value: money(v.total) },
        { label: 'Weighted', value: money(v.weighted), tone: 'good' },
        { label: 'At risk', value: String(v.atRisk), tone: v.atRisk ? 'bad' : undefined }
      ]} />
      {v.deals.length === 0 ? <Empty>No open deals match. Try a different period or scope.</Empty> : (
        <RowList label="Deals">
          {visible.map(d => (
            <Row key={d.id} title={d.title}
              meta={[d.account, d.stage, `closes ${d.closeText}`, host.fullscreen || v.params.scope !== 'mine' ? d.owner : ''].filter(Boolean).join(' · ')}
              detail={d.health.risks[0] ? `⚠ ${d.health.risks[0]}` : d.nextStep ? `Next: ${d.nextStep}` : undefined}
              amount={d.amountText} badge={<HealthBadge health={d.health} />}
              onOpen={openRecord(host, 'opportunities' as EntityKey, d.id)} />
          ))}
        </RowList>
      )}
      {more}
    </>
  );
};

export default class CrmPipelineCopilotComponent extends CrmCopilotComponentBase<ICrmPipelineCopilotComponentProperties> {
  protected readonly toolTitle = 'Pipeline';
  protected readonly lists: EntityKey[] = ['opportunities', 'accounts', 'activities', 'quotes'];
  protected renderView(host: ICrmToolHost, properties: unknown): React.ReactNode {
    return <PipelineView host={host} properties={properties} />;
  }
}
