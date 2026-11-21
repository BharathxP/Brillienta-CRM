import * as React from 'react';
import type { EntityKey, IRecord } from '../../webparts/brillientaCrm/models/model';
import CrmCopilotComponentBase from '../shared/CrmCopilotComponentBase';
import { accountSummary, accountView, findAccounts, normalizeAccount } from '../shared/crmQueries';
import {
  Actions, Badge, Button, Empty, HealthBadge, ICrmToolHost, Kpis, openRecord, Row, RowList, Section, styles
} from '../shared/ui';
import { fmtMoney } from '../../webparts/brillientaCrm/components/utils';
import type { ICrmAccountSummaryCopilotComponentProperties } from './CrmAccountSummaryCopilotComponentProperties';

const CASE_BADGE: Record<string, { tone: 'bad' | 'warn' | 'info'; text: string }> = {
  breached: { tone: 'bad', text: 'SLA breached' },
  firstResponseLate: { tone: 'bad', text: 'Reply late' },
  dueSoon: { tone: 'warn', text: 'Due soon' },
  paused: { tone: 'info', text: 'Waiting on customer' }
};

const AccountDetails: React.FC<{ host: ICrmToolHost; account: IRecord; onBack?: () => void }> = ({ host, account, onBack }) => {
  const { engine } = host.session;
  const v = React.useMemo(() => accountView(engine, account), [engine, account]);
  React.useEffect(() => {
    host.setModelContext(accountSummary(engine, v), {
      tool: 'GetAccountSummary', accountId: account.id, account: account.title,
      openDeals: v.deals.map(d => ({ id: d.id, name: d.title, stage: d.stage, amount: d.amount, close: d.close, winLikelihood: d.health.score })),
      openCases: v.cases.map(c => ({ id: c.id, number: c.number, subject: c.title, priority: c.priority, sla: c.state })),
      subscriptions: v.subscriptions.map(s => ({ id: s.id, name: s.title, renewal: s.renewal, arr: s.arr })),
      unpaidInvoices: v.invoices.map(i => ({ id: i.id, number: i.number, balance: i.balance, daysOverdue: i.daysOverdue }))
    });
  }, [v]);
  const money = (n: number): string => fmtMoney(n, engine.settings.currency);
  const limit = host.fullscreen ? 50 : 4;
  return (
    <>
      <RowList label="Account">
        <Row title={account.title || `#${account.id}`} meta={[v.type, v.owner && `Owner ${v.owner}`, `${v.contacts} contact${v.contacts === 1 ? '' : 's'}`].filter(Boolean).join(' · ')}
          onOpen={openRecord(host, 'accounts', account.id!)} />
      </RowList>
      <div style={{ height: 10 }} />
      <Kpis items={[
        { label: 'Open pipeline', value: money(v.openPipeline) },
        { label: 'ARR', value: money(v.arr), tone: 'good' },
        { label: 'Outstanding', value: money(v.outstanding), tone: v.outstanding > 0 ? 'warn' : undefined },
        ...(v.hasCases ? [{ label: 'Open cases', value: String(v.cases.length), tone: v.cases.some(c => c.state === 'breached') ? 'bad' as const : undefined }] : [])
      ]} />

      <Section title={`Open deals (${v.deals.length})`}>
        {v.deals.length ? <RowList label="Open deals">{v.deals.slice(0, limit).map(d => (
          <Row key={d.id} title={d.title} meta={`${d.stage} · closes ${d.closeText}`} amount={d.amountText} badge={<HealthBadge health={d.health} />} onOpen={openRecord(host, 'opportunities', d.id)} />
        ))}</RowList> : <Empty>No open deals.</Empty>}
      </Section>

      {v.hasCases && (
        <Section title={`Open cases (${v.cases.length})`}>
          {v.cases.length ? <RowList label="Open cases">{v.cases.slice(0, limit).map(c => (
            <Row key={c.id} title={`${c.number} ${c.title}`} meta={`${c.priority} · ${c.status} · SLA ${c.slaText}`}
              badge={CASE_BADGE[c.state] ? <Badge tone={CASE_BADGE[c.state].tone}>{CASE_BADGE[c.state].text}</Badge> : undefined}
              onOpen={openRecord(host, 'cases', c.id)} />
          ))}</RowList> : <Empty>No open cases.</Empty>}
        </Section>
      )}

      <Section title={`Subscriptions (${v.subscriptions.length})`}>
        {v.subscriptions.length ? <RowList label="Subscriptions">{v.subscriptions.slice(0, limit).map(s => (
          <Row key={s.id} title={s.title} meta={`${s.status} · renews ${s.renewalText}${s.autoRenew ? '' : ' · manual renewal'}`} amount={`${s.arrText} ARR`} onOpen={openRecord(host, 'subscriptions', s.id)} />
        ))}</RowList> : <Empty>No active subscriptions.</Empty>}
      </Section>

      {v.invoices.length > 0 && (
        <Section title={`Unpaid invoices (${v.invoices.length})`}>
          <RowList label="Unpaid invoices">{v.invoices.slice(0, limit).map(i => (
            <Row key={i.id} title={i.number || i.title} meta={`Due ${i.dueText}`} amount={i.balanceText}
              badge={i.daysOverdue > 0 ? <Badge tone="bad">{i.daysOverdue} d overdue</Badge> : undefined} onOpen={openRecord(host, 'invoices', i.id)} />
          ))}</RowList>
        </Section>
      )}

      <Section title="Activity">
        {v.activities.length ? <RowList label="Activity">{v.activities.map(a => (
          <Row key={a.id} title={a.title} meta={[a.type, a.when].filter(Boolean).join(' · ')} />
        ))}</RowList> : <Empty>No activity logged yet.</Empty>}
      </Section>

      <Actions>
        {host.expand && <Button onClick={host.expand}>Expand</Button>}
        <Button onClick={() => host.sendPrompt(`Draft a short follow-up email to ${account.title} based on this account summary.`)}>Draft follow-up</Button>
        {onBack && <Button onClick={onBack}>Choose another account</Button>}
      </Actions>
    </>
  );
};

export const AccountSummaryView: React.FC<{ host: ICrmToolHost; properties: unknown }> = ({ host, properties }) => {
  const { engine } = host.session;
  const initial = normalizeAccount(properties).accountName;
  const [query, setQuery] = React.useState(initial);
  const [submitted, setSubmitted] = React.useState(initial);
  const matches = React.useMemo(() => findAccounts(engine, submitted), [engine, submitted]);
  const [picked, setPicked] = React.useState<IRecord | undefined>(undefined);
  const account = picked || (matches.length === 1 ? matches[0] : undefined);

  React.useEffect(() => {
    if (!account) {
      host.setModelContext(submitted
        ? (matches.length ? `${matches.length} accounts match "${submitted}": ${matches.map(m => m.title).join(', ')}. The user is choosing one.` : `No account matches "${submitted}".`)
        : 'The user is searching for an account.');
    }
  }, [account, matches, submitted]);

  if (account) {
    return <AccountDetails host={host} account={account} onBack={matches.length > 1 || !initial ? () => setPicked(undefined) : undefined} />;
  }
  return (
    <>
      <form className={styles.search} onSubmit={e => { e.preventDefault(); setPicked(undefined); setSubmitted(query.trim()); }}>
        <label className={styles.visuallyHidden} htmlFor="crm-account-search">Account name</label>
        <input id="crm-account-search" value={query} placeholder="Account name" onChange={e => setQuery(e.target.value)} />
        <Button type="submit" primary onClick={() => undefined}>Find</Button>
      </form>
      <div style={{ height: 10 }} />
      {!submitted ? <Empty>Type a customer name to see its deals, cases, renewals and invoices.</Empty>
        : !matches.length ? <Empty>No account matches “{submitted}”.</Empty>
        : (
          <RowList label="Matching accounts">
            {matches.map(m => (
              <li key={m.id} className={styles.row}>
                <div className={styles.rowMain}>
                  <div className={styles.rowTitle}>{m.title}</div>
                  <div className={styles.rowMeta}>{[m.accountType, m.city, m.country].filter(Boolean).join(' · ')}</div>
                </div>
                <button type="button" className={styles.button} onClick={() => setPicked(m)}>Show</button>
              </li>
            ))}
          </RowList>
        )}
    </>
  );
};

export default class CrmAccountSummaryCopilotComponent extends CrmCopilotComponentBase<ICrmAccountSummaryCopilotComponentProperties> {
  protected readonly toolTitle = 'Account summary';
  protected readonly lists: EntityKey[] = ['accounts', 'contacts', 'opportunities', 'quotes', 'cases', 'subscriptions', 'invoices', 'activities'];
  protected renderView(host: ICrmToolHost, properties: unknown): React.ReactNode {
    return <AccountSummaryView host={host} properties={properties} />;
  }
}
