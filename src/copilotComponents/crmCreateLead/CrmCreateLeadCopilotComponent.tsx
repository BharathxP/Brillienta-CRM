import * as React from 'react';
import type { EntityKey, IRecord } from '../../webparts/brillientaCrm/models/model';
import CrmCopilotComponentBase from '../shared/CrmCopilotComponentBase';
import { ILeadDraft, LEAD_RATINGS, leadRecord, normalizeLeadDraft, validateLeadDraft } from '../shared/crmQueries';
import { Actions, Button, ICrmToolHost, Note, openRecord, Row, RowList, styles } from '../shared/ui';
import type { ICrmCreateLeadCopilotComponentProperties } from './CrmCreateLeadCopilotComponentProperties';

type FieldKey = keyof ILeadDraft;
type FieldEvent = React.ChangeEvent<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>;

const TEXT_FIELDS: { key: FieldKey; label: string; type?: string; full?: boolean; autoComplete?: string }[] = [
  { key: 'firstName', label: 'First name *', autoComplete: 'given-name' },
  { key: 'lastName', label: 'Last name', autoComplete: 'family-name' },
  { key: 'email', label: 'Email', type: 'email', autoComplete: 'email' },
  { key: 'phone', label: 'Phone', type: 'tel', autoComplete: 'tel' },
  { key: 'company', label: 'Company', autoComplete: 'organization' },
  { key: 'jobTitle', label: 'Job title', autoComplete: 'organization-title' },
  { key: 'country', label: 'Country', autoComplete: 'country-name' }
];

export const CreateLeadView: React.FC<{ host: ICrmToolHost; properties: unknown }> = ({ host, properties }) => {
  const { engine } = host.session;
  const sources = engine.settings.leadSources || [];
  const [draft, setDraft] = React.useState<ILeadDraft>(() => normalizeLeadDraft(properties, sources));
  const [touched, setTouched] = React.useState(false);
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState<string>();
  const [saved, setSaved] = React.useState<IRecord>();
  const errors = validateLeadDraft(draft);
  const canCreate = engine.can('create', 'leads');
  const duplicates = React.useMemo(() => engine.findDuplicates('leads', leadRecord(draft)), [engine, draft.email, draft.firstName, draft.lastName, draft.company]);
  const set = (k: FieldKey): ((e: FieldEvent) => void) => (e: FieldEvent): void => setDraft(d => ({ ...d, [k]: e.target.value }));

  React.useEffect(() => {
    if (saved) {
      host.setModelContext(`The user saved a new lead in Brillienta CRM: ${saved.title} (ID ${saved.id})${saved.company ? `, ${saved.company}` : ''}. Lead score ${saved.leadScore ?? 0}${saved.owner ? `, assigned to ${saved.owner.title}` : ''}.`,
        { tool: 'CreateLead', saved: true, leadId: saved.id, name: saved.title, owner: saved.owner?.title || '' });
    } else {
      host.setModelContext('A new-lead form is open in Brillienta CRM with the values from the prompt. It has NOT been saved; the user must review it and press Save.',
        { tool: 'CreateLead', saved: false, draft });
    }
  }, [saved]);

  const save = (): void => {
    setTouched(true);
    if (Object.keys(errors).length || busy) { return; }
    setBusy(true);
    setError(undefined);
    engine.save('leads', leadRecord(draft))
      .then(r => { setSaved(r); setBusy(false); })
      .catch(e => { setError((e as Error)?.message || String(e)); setBusy(false); });
  };

  if (saved) {
    return (
      <>
        <div className={styles.success} role="status">Lead saved to Brillienta CRM.</div>
        <div style={{ height: 10 }} />
        <RowList label="New lead">
          <Row title={saved.title || 'Lead'} meta={[saved.company, saved.leadSource, saved.owner ? `Owner ${saved.owner.title}` : 'Unassigned'].filter(Boolean).join(' · ')}
            amount={`Score ${saved.leadScore ?? 0}`} onOpen={openRecord(host, 'leads', saved.id!)} />
        </RowList>
      </>
    );
  }

  return (
    <form noValidate onSubmit={e => { e.preventDefault(); save(); }}>
      {!canCreate && <Note>{engine.license?.canWrite === false ? engine.license.message : 'You do not have permission to add leads in Brillienta CRM.'}</Note>}
      <p style={{ margin: '0 0 10px' }}>Check the details, then save. Nothing is added to the CRM until you press Save.</p>
      <div className={styles.form}>
        {TEXT_FIELDS.map(f => (
          <div key={f.key} className={styles.field}>
            <label htmlFor={`crm-lead-${f.key}`}>{f.label}</label>
            <input id={`crm-lead-${f.key}`} type={f.type || 'text'} autoComplete={f.autoComplete} value={draft[f.key]} onChange={set(f.key)}
              aria-invalid={touched && !!errors[f.key]} aria-describedby={touched && errors[f.key] ? `crm-lead-${f.key}-err` : undefined} />
            {touched && errors[f.key] && <span id={`crm-lead-${f.key}-err`} className={styles.error}>{errors[f.key]}</span>}
          </div>
        ))}
        <div className={styles.field}>
          <label htmlFor="crm-lead-source">Lead source</label>
          <select id="crm-lead-source" value={draft.leadSource} onChange={set('leadSource')}>
            <option value="">—</option>
            {sources.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
        </div>
        <div className={styles.field}>
          <label htmlFor="crm-lead-rating">Rating</label>
          <select id="crm-lead-rating" value={draft.rating} onChange={set('rating')}>
            {LEAD_RATINGS.map(r => <option key={r} value={r}>{r}</option>)}
          </select>
        </div>
        <div className={`${styles.field} ${styles.full}`}>
          <label htmlFor="crm-lead-description">Notes</label>
          <textarea id="crm-lead-description" value={draft.description} onChange={set('description')} />
        </div>
      </div>
      {duplicates.length > 0 && (
        <div style={{ marginTop: 12 }}>
          <Note>Possible duplicate{duplicates.length > 1 ? 's' : ''} already in the CRM:</Note>
          <RowList label="Possible duplicates">
            {duplicates.map(d => <Row key={d.id} title={d.title || `#${d.id}`} meta={[d.company, d.email, d.status].filter(Boolean).join(' · ')} onOpen={openRecord(host, 'leads', d.id!)} />)}
          </RowList>
        </div>
      )}
      {error && <div className={styles.note} role="alert" style={{ marginTop: 12 }}>{error}</div>}
      <Actions>
        <Button type="submit" primary disabled={busy || !canCreate} onClick={() => undefined}>{busy ? 'Saving…' : duplicates.length ? 'Save anyway' : 'Save lead'}</Button>
      </Actions>
    </form>
  );
};

export default class CrmCreateLeadCopilotComponent extends CrmCopilotComponentBase<ICrmCreateLeadCopilotComponentProperties> {
  protected readonly toolTitle = 'New lead';
  // Leads are loaded for duplicate detection; activities and demos feed the lead score.
  protected readonly lists: EntityKey[] = ['leads', 'activities', 'demos'];
  protected renderView(host: ICrmToolHost, properties: unknown): React.ReactNode {
    return <CreateLeadView host={host} properties={properties} />;
  }
}
