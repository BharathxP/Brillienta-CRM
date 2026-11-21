import * as React from 'react';
import type { EntityKey } from '../../webparts/brillientaCrm/models/model';
import CrmCopilotComponentBase from '../shared/CrmCopilotComponentBase';
import { invoiceSummary, invoiceView, normalizeInvoices } from '../shared/crmQueries';
import { Badge, Empty, ICrmToolHost, Kpis, openRecord, Row, RowList, useRowLimit } from '../shared/ui';
import { fmtMoney } from '../../webparts/brillientaCrm/components/utils';
import type { ICrmOverdueInvoicesCopilotComponentProperties } from './CrmOverdueInvoicesCopilotComponentProperties';

export const InvoicesView: React.FC<{ host: ICrmToolHost; properties: unknown }> = ({ host, properties }) => {
  const { engine } = host.session;
  const v = React.useMemo(() => invoiceView(engine, normalizeInvoices(properties)), [engine, properties]);
  React.useEffect(() => {
    host.setModelContext(invoiceSummary(engine, v), {
      tool: 'GetOverdueInvoices', filter: v.label,
      invoices: v.invoices.slice(0, 25).map(i => ({ id: i.id, number: i.number, account: i.account, due: i.due, daysOverdue: i.daysOverdue, balance: i.balance }))
    });
  }, [v]);
  const { visible, more } = useRowLimit(v.invoices, host);
  const oldest = v.invoices[0]?.daysOverdue || 0;
  return (
    <>
      <p style={{ margin: '0 0 10px' }}>{v.label}</p>
      <Kpis items={[
        { label: 'Invoices', value: String(v.invoices.length) },
        { label: 'Outstanding', value: fmtMoney(v.outstanding, engine.settings.currency), tone: v.outstanding ? 'bad' : undefined },
        { label: 'Oldest', value: oldest ? `${oldest} days` : '—', tone: oldest > 30 ? 'bad' : oldest ? 'warn' : undefined }
      ]} />
      {v.invoices.length === 0 ? <Empty>No overdue invoices. Nice.</Empty> : (
        <RowList label="Overdue invoices">
          {visible.map(i => (
            <Row key={i.id} title={i.number ? `${i.number} · ${i.account || i.title}` : i.title}
              meta={`Due ${i.dueText}`}
              amount={i.balanceText}
              badge={<Badge tone={i.daysOverdue > 30 ? 'bad' : 'warn'}>{i.daysOverdue} d overdue</Badge>}
              onOpen={openRecord(host, 'invoices', i.id)} />
          ))}
        </RowList>
      )}
      {more}
    </>
  );
};

export default class CrmOverdueInvoicesCopilotComponent extends CrmCopilotComponentBase<ICrmOverdueInvoicesCopilotComponentProperties> {
  protected readonly toolTitle = 'Overdue invoices';
  protected readonly lists: EntityKey[] = ['invoices', 'accounts'];
  protected renderView(host: ICrmToolHost, properties: unknown): React.ReactNode {
    return <InvoicesView host={host} properties={properties} />;
  }
}
