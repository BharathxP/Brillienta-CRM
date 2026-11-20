import * as React from 'react';
import { DefaultButton, Dialog, DialogFooter, DialogType, Dropdown, Icon, Link, MessageBar, MessageBarType, PrimaryButton, Spinner, Stack, Text, TextField } from '@fluentui/react';
import { EntityKey, IRecord } from '../../models/model';
import { IHelpdeskTicket, isOpenTicket } from '../../services/Helpdesk';
import { useCrm } from '../CrmContext';
import { errorText, fmtDate } from '../utils';

/** Status pill colour by Helpdesk status kind. */
export const TicketStatus: React.FC<{ t: IHelpdeskTicket }> = ({ t }) => (
  <span className={`crm-tk-status crm-tk-${t.pending ? 'pending' : t.kind.toLowerCase()}`}>{t.pending ? 'Sent to Helpdesk' : t.status}</span>
);

export const TicketRow: React.FC<{ t: IHelpdeskTicket; sub?: string }> = ({ t, sub }) => (
  <a className="crm-tk-row" href={t.link} target="_blank" rel="noopener noreferrer" data-interception="off">
    <div className="crm-tk-main">
      <Text styles={{ root: { fontWeight: 600 } }}>{t.title}</Text>
      <Text variant="small" block className="crm-muted">
        {[t.pending ? undefined : t.number, sub, t.priority, t.assignee ? `with ${t.assignee}` : 'unassigned', fmtDate(t.created)].filter(Boolean).join(' · ')}
      </Text>
    </div>
    {t.breached && isOpenTicket(t) && <span className="crm-tk-sla" title="Resolution SLA breached">SLA</span>}
    <TicketStatus t={t} />
    <Icon iconName="OpenInNewWindow" className="crm-muted" aria-hidden />
  </a>
);

/** Dialog to raise a Brillienta Helpdesk ticket for a CRM record. */
export const NewTicketDialog: React.FC<{ entity: EntityKey; record: IRecord; onClose: (created?: IHelpdeskTicket) => void }> = ({ entity, record, onClose }) => {
  const { engine, notify } = useCrm();
  const account = entity === 'accounts' ? record : record.accountRef ? engine.find('accounts', record.accountRef) : undefined;
  const people = (engine.data.contacts || []).filter(c => (entity === 'contacts' ? c.id === record.id : account && c.accountRef === account.id) && c.email);
  const caseContact = entity === 'cases' && record.contactRef ? engine.find('contacts', record.contactRef) : undefined;
  const [contactId, setContactId] = React.useState<number | undefined>(entity === 'contacts' ? record.id : caseContact?.id || people[0]?.id);
  const [title, setTitle] = React.useState<string>(entity === 'cases' ? String(record.title || '') : '');
  const [details, setDetails] = React.useState<string>(entity === 'cases' ? String(record.description || '') : '');
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState<string>();
  const contact = contactId ? engine.find('contacts', contactId) : undefined;
  const email = String(contact?.email || record.requesterEmail || '');

  const submit = (): void => {
    setBusy(true); setError(undefined);
    const header = [
      account ? `Customer: ${account.title}` : '',
      contact ? `Contact: ${contact.title}${contact.email ? ` <${contact.email}>` : ''}${contact.phone ? `, ${contact.phone}` : ''}` : '',
      entity === 'cases' ? `CRM case: ${record.caseNumber || record.id}` : '',
      `Raised from Brillienta CRM by ${engine.me?.title || ''}`
    ].filter(Boolean).join('\n');
    engine.helpdesk.create({
      title: title.trim(), details: `${details.trim()}\n\n${header}`, requesterEmail: email,
      accountId: account?.id, contactId: contact?.id, caseId: entity === 'cases' ? record.id : undefined
    }).then(t => {
      setBusy(false);
      notify('Ticket sent to Brillienta Helpdesk. It gets its number and is assigned as soon as a Helpdesk agent has Helpdesk open.', MessageBarType.success);
      onClose(t);
    }).catch(e => { setBusy(false); setError(errorText(e)); });
  };

  return (
    <Dialog hidden={false} onDismiss={() => onClose()} minWidth={520}
      dialogContentProps={{ type: DialogType.normal, title: 'New Helpdesk ticket', subText: 'The ticket is created in Brillienta Helpdesk, where the support team works on it.' }}>
      <Stack tokens={{ childrenGap: 12 }}>
        {people.length > 1 && (
          <Dropdown label="Requester" selectedKey={contactId} options={people.map(c => ({ key: c.id as number, text: `${c.title} (${c.email})` }))} onChange={(_, o) => setContactId(Number(o?.key))} />
        )}
        {people.length <= 1 && <Text><b>Requester:</b> {contact ? `${contact.title} (${email || 'no email'})` : 'You'}</Text>}
        <TextField label="Subject" required value={title} onChange={(_, v) => setTitle(v || '')} />
        <TextField label="Description" multiline rows={5} value={details} onChange={(_, v) => setDetails(v || '')} placeholder="What does the customer need?" />
        {error && <MessageBar messageBarType={MessageBarType.error} isMultiline>{error}</MessageBar>}
      </Stack>
      <DialogFooter>
        <PrimaryButton text={busy ? 'Sending…' : 'Create ticket'} disabled={busy || !title.trim()} onClick={submit} />
        <DefaultButton text="Cancel" onClick={() => onClose()} />
      </DialogFooter>
    </Dialog>
  );
};

/** "Support tickets" tab on accounts, contacts and cases. */
export const HelpdeskTickets: React.FC<{ entity: EntityKey; record: IRecord }> = ({ entity, record }) => {
  const { engine } = useCrm();
  const [rows, setRows] = React.useState<IHelpdeskTicket[]>();
  const [error, setError] = React.useState<string>();
  const [dialog, setDialog] = React.useState(false);
  const [showClosed, setShowClosed] = React.useState(false);
  const load = (force?: boolean): void => {
    setError(undefined);
    engine.helpdesk.ticketsFor(entity, record, force).then(setRows).catch(e => { setRows([]); setError(errorText(e)); });
  };
  React.useEffect(() => load(), [record.id]);
  const open = (rows || []).filter(isOpenTicket);
  const closed = (rows || []).filter(t => !isOpenTicket(t));
  return (
    <div style={{ paddingTop: 12 }}>
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center" wrap tokens={{ childrenGap: 8 }}>
        <Text className="crm-muted">{rows ? `${open.length} open · ${closed.length} resolved or closed` : 'Loading tickets from Brillienta Helpdesk…'}</Text>
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <DefaultButton iconProps={{ iconName: 'Refresh' }} text="Refresh" onClick={() => load(true)} />
          <PrimaryButton iconProps={{ iconName: 'Add' }} text="New ticket" onClick={() => setDialog(true)} />
        </Stack>
      </Stack>
      {!rows && <Spinner styles={{ root: { marginTop: 16 } }} />}
      {error && (
        <MessageBar messageBarType={MessageBarType.warning} isMultiline styles={{ root: { marginTop: 12 } }}>
          Could not read Brillienta Helpdesk: {error}. Check Settings › Brillienta Helpdesk, and that you have access to the Helpdesk site.
        </MessageBar>
      )}
      {rows && !error && rows.length === 0 && (
        <div className="crm-empty-state">
          <Icon iconName="Headset" />
          <Text block styles={{ root: { fontWeight: 600 } }}>No Helpdesk tickets yet</Text>
          <Text block className="crm-muted">Tickets from {entity === 'accounts' ? 'this customer’s contacts' : entity === 'contacts' ? 'this contact' : 'this case'} appear here. Only Helpdesk agents see other people’s tickets.</Text>
        </div>
      )}
      <div style={{ marginTop: 8 }}>
        {open.map(t => <TicketRow key={t.id} t={t} sub={entity === 'accounts' ? t.requester : undefined} />)}
        {closed.length > 0 && <Link onClick={() => setShowClosed(!showClosed)} styles={{ root: { display: 'block', margin: '10px 0' } }}>{showClosed ? 'Hide' : 'Show'} {closed.length} resolved or closed</Link>}
        {showClosed && closed.map(t => <TicketRow key={t.id} t={t} sub={entity === 'accounts' ? t.requester : undefined} />)}
      </div>
      {dialog && <NewTicketDialog entity={entity} record={record} onClose={t => { setDialog(false); if (t) { setRows([t, ...(rows || [])]); } }} />}
    </div>
  );
};
