import * as React from 'react';
import type { EntityKey } from '../../webparts/brillientaCrm/models/model';
import CrmCopilotComponentBase from '../shared/CrmCopilotComponentBase';
import { normalizeRenewals, renewalSummary, renewalView } from '../shared/crmQueries';
import { Badge, Empty, ICrmToolHost, Kpis, openRecord, Row, RowList, useRowLimit } from '../shared/ui';
import { fmtMoney } from '../../webparts/brillientaCrm/components/utils';
import type { ICrmRenewalsCopilotComponentProperties } from './CrmRenewalsCopilotComponentProperties';

export const RenewalsView: React.FC<{ host: ICrmToolHost; properties: unknown }> = ({ host, properties }) => {
  const { engine } = host.session;
  const v = React.useMemo(() => renewalView(engine, normalizeRenewals(properties)), [engine, properties]);
  React.useEffect(() => {
    host.setModelContext(renewalSummary(engine, v), {
      tool: 'GetRenewals', filter: v.label,
      renewals: v.renewals.slice(0, 25).map(r => ({ id: r.id, name: r.title, account: r.account, renewal: r.renewal, arr: r.arr, autoRenew: r.autoRenew, seats: r.seats }))
    });
  }, [v]);
  const { visible, more } = useRowLimit(v.renewals, host);
  return (
    <>
      <p style={{ margin: '0 0 10px' }}>{v.label}</p>
      <Kpis items={[
        { label: 'Renewals', value: String(v.renewals.length) },
        { label: 'ARR up for renewal', value: fmtMoney(v.arr, engine.settings.currency), tone: 'good' },
        { label: 'Need attention', value: String(v.atRisk), tone: v.atRisk ? 'warn' : undefined }
      ]} />
      {v.renewals.length === 0 ? <Empty>No subscriptions renew in this window.</Empty> : (
        <RowList label="Renewals">
          {visible.map(r => {
            const late = r.daysLeft !== undefined && r.daysLeft < 0;
            return (
              <Row key={r.id} title={r.title}
                meta={[r.account, `renews ${r.renewalText}`, r.seats ? `${r.seats} seat${r.seats === 1 ? '' : 's'}` : ''].filter(Boolean).join(' · ')}
                amount={`${r.arrText} ARR`}
                badge={late ? <Badge tone="bad">Renewal date passed</Badge>
                  : !r.autoRenew ? <Badge tone="warn">Manual renewal</Badge>
                  : <Badge tone="info">{r.daysLeft === 0 ? 'Today' : `In ${r.daysLeft} d`}</Badge>}
                onOpen={openRecord(host, 'subscriptions', r.id)} />
            );
          })}
        </RowList>
      )}
      {more}
    </>
  );
};

export default class CrmRenewalsCopilotComponent extends CrmCopilotComponentBase<ICrmRenewalsCopilotComponentProperties> {
  protected readonly toolTitle = 'Upcoming renewals';
  protected readonly lists: EntityKey[] = ['subscriptions', 'accounts'];
  protected renderView(host: ICrmToolHost, properties: unknown): React.ReactNode {
    return <RenewalsView host={host} properties={properties} />;
  }
}
