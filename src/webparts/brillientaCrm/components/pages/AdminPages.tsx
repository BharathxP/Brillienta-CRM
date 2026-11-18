import * as React from 'react';
import {
  DetailsList, DetailsListLayoutMode, Dropdown, Icon, Link, MessageBar, MessageBarType, PrimaryButton, SearchBox, SelectionMode,
  Spinner, SpinnerSize, Stack, Text, DefaultButton
} from '@fluentui/react';
import { IRecord } from '../../models/model';
import { DataService, IListStatus } from '../../services/DataService';
import { useCrm } from '../CrmContext';
import { errorText, exportCsv, fmtDate } from '../utils';

export const AuditLogPage: React.FC = () => {
  const { engine, open } = useCrm();
  const [rows, setRows] = React.useState<IRecord[]>();
  const [entity, setEntity] = React.useState('');
  const [q, setQ] = React.useState('');
  const [error, setError] = React.useState<string>();
  React.useEffect(() => {
    setRows(undefined);
    engine.ds.getAudit(entity || undefined, undefined, 1000).then(setRows).catch(e => setError(errorText(e)));
  }, [entity]);
  const filtered = (rows || []).filter(r => !q || `${r.recordTitle} ${r.author?.title} ${r.title} ${JSON.stringify(r.changes || {})}`.toLowerCase().indexOf(q.toLowerCase()) >= 0);
  return (
    <Stack tokens={{ childrenGap: 8 }}>
      <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
        <SearchBox placeholder="Search changes, users, records" value={q} onChange={(_, v) => setQ(v || '')} styles={{ root: { width: 280 } }} />
        <Dropdown selectedKey={entity} styles={{ root: { width: 180 } }} ariaLabel="Record type"
          options={[{ key: '', text: 'All record types' }, ...engine.defs.map(d => ({ key: d.key, text: d.plural }))]} onChange={(_, o) => setEntity(String(o?.key || ''))} />
        <DefaultButton iconProps={{ iconName: 'ExcelDocument' }} text="Export" onClick={() => exportCsv('audit-log.csv', filtered.map(r => ({
          When: r.created, User: r.author?.title || '', Action: r.title, Type: r.entityType, RecordId: r.recordId, Record: r.recordTitle, Changes: JSON.stringify(r.changes || {})
        })))} />
      </Stack>
      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
      {!rows && !error && <Spinner label="Loading audit log..." />}
      {rows && (
        <DetailsList items={filtered} selectionMode={SelectionMode.none} layoutMode={DetailsListLayoutMode.justified} columns={[
          { key: 'w', name: 'When', minWidth: 140, onRender: r => fmtDate(r.created, true) },
          { key: 'u', name: 'User', minWidth: 120, onRender: r => r.author?.title },
          { key: 'a', name: 'Action', fieldName: 'title', minWidth: 70 },
          { key: 't', name: 'Type', fieldName: 'entityType', minWidth: 90 },
          { key: 'r', name: 'Record', minWidth: 160, onRender: r => {
            const d = engine.defs.find(x => x.key === r.entityType);
            const rec = d ? engine.find(d.key, r.recordId) : undefined;
            return rec && d ? <Link onClick={() => open(d.key, rec)}>{r.recordTitle || `#${r.recordId}`}</Link> : `${r.recordTitle || ''} #${r.recordId}`;
          } },
          { key: 'c', name: 'Changes', minWidth: 260, isMultiline: true, onRender: r => r.changes ? Object.keys(r.changes).map(k => (
            <div key={k}><span className="crm-muted">{k}:</span> {String(r.changes[k][0] || '—')} → <b>{String(r.changes[k][1] || '—')}</b></div>
          )) : '' }
        ]} />
      )}
    </Stack>
  );
};

export const SetupPage: React.FC<{ ds: DataService; statuses?: IListStatus[]; canProvision: boolean; onDone: () => void; embedded?: boolean; logo?: string }> =
  ({ ds, statuses, canProvision, onDone, embedded, logo }) => {
    const [busy, setBusy] = React.useState(false);
    const [msg, setMsg] = React.useState<string>();
    const [error, setError] = React.useState<string>();
    const ready = !!statuses && statuses.every(s => s.exists && !s.missingFields.length);
    const run = (): void => {
      setBusy(true);
      setError(undefined);
      ds.provision(setMsg).then(() => { setBusy(false); onDone(); }).catch(e => {
        setBusy(false);
        const m = errorText(e);
        setError(/denied|403|unauthori/i.test(m) ? `${m} — creating lists requires a site owner (Manage Lists permission).` : m);
      });
    };
    return (
      <Stack tokens={{ childrenGap: 12 }} className={embedded ? undefined : 'crm-card'} style={embedded ? undefined : { maxWidth: 900, margin: '24px auto' }}>
        {!embedded && logo && <img src={logo} alt="Brillienta CRM" style={{ height: 40, alignSelf: 'flex-start' }} />}
        {!embedded && <Text variant="xxLarge" styles={{ root: { fontWeight: 600 } }}>Welcome to Brillienta CRM</Text>}
        <Text>Brillienta CRM stores all data in SharePoint lists on this site — your data never leaves your Microsoft 365 tenant. {ready ? 'All lists are ready.' : 'The following lists and columns will be created or upgraded:'}</Text>
        {statuses && (
          <div className="crm-setup-grid">
            {statuses.map(s => (
              <div key={s.key} className="crm-setup-row">
                <Icon iconName={s.exists && !s.missingFields.length ? 'CompletedSolid' : 'Warning'} style={{ color: s.exists && !s.missingFields.length ? 'var(--good)' : 'var(--warn)' }} />
                <div>
                  <Text block styles={{ root: { fontWeight: 600 } }}>{s.url ? <Link href={s.url} target="_blank">{s.title}</Link> : s.title}</Text>
                  <Text variant="small" className="crm-muted">{!s.exists ? 'Will be created' : s.missingFields.length ? `${s.missingFields.length} new column(s)` : 'Ready'}</Text>
                </div>
              </div>
            ))}
          </div>
        )}
        {error && <MessageBar messageBarType={MessageBarType.error} isMultiline>{error}</MessageBar>}
        {busy && <Spinner size={SpinnerSize.small} label={msg || 'Working...'} labelPosition="right" />}
        {canProvision
          ? <Stack horizontal tokens={{ childrenGap: 8 }}>
              <PrimaryButton text={ready ? 'Re-check & repair' : 'Set up CRM'} iconProps={{ iconName: 'Build' }} disabled={busy} onClick={run} />
              {ready && !embedded && <DefaultButton text="Open CRM" onClick={onDone} />}
            </Stack>
          : !ready && <MessageBar messageBarType={MessageBarType.warning}>A site owner needs to open this page once and click “Set up CRM”.</MessageBar>}
      </Stack>
    );
  };
