/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
import {
  Checkbox, DefaultButton, Dialog, DialogFooter, DialogType, Dropdown, MessageBar, MessageBarType, PrimaryButton,
  ProgressIndicator, SpinButton, Stack, Text, TextField, ChoiceGroup, Link, Position
} from '@fluentui/react';
import { EntityKey, IRecord } from '../../models/model';
import { useCrm } from '../CrmContext';
import { DateTimeField, LookupPicker, UserPicker } from './controls';
import { choicesFor } from './fields';
import { errorText, parseCsv } from '../utils';

/* ---------------------------------------------------------------- convert lead */

export const ConvertLeadDialog: React.FC<{ lead: IRecord; onClose: (done?: boolean) => void }> = ({ lead, onClose }) => {
  const { engine, notify, data, open } = useCrm();
  React.useEffect(() => { engine.ensure(['accounts', 'contacts']).catch(() => undefined); }, []);
  const accMatch = (data.accounts || []).find(a => lead.company && String(a.title).toLowerCase() === String(lead.company).toLowerCase());
  const conMatch = (data.contacts || []).find(c => lead.email && String(c.email).toLowerCase() === String(lead.email).toLowerCase());
  const [accMode, setAccMode] = React.useState<string>(accMatch ? 'existing' : (lead.company ? 'new' : 'none'));
  const [accountId, setAccountId] = React.useState<number | undefined>(accMatch?.id);
  const [conMode, setConMode] = React.useState<string>(conMatch ? 'existing' : 'new');
  const [contactId, setContactId] = React.useState<number | undefined>(conMatch?.id);
  const [createOpp, setCreateOpp] = React.useState(true);
  const [oppName, setOppName] = React.useState(`${lead.company || lead.title} - ${lead.offer || 'New deal'}`);
  const [amount, setAmount] = React.useState<number>(0);
  const [close, setClose] = React.useState<string | undefined>(new Date(Date.now() + 30 * 86400000).toISOString());
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState<string>();

  React.useEffect(() => {
    if (accMatch && !accountId) { setAccMode('existing'); setAccountId(accMatch.id); }
    if (conMatch && !contactId) { setConMode('existing'); setContactId(conMatch.id); }
  }, [accMatch?.id, conMatch?.id]);

  const run = (): void => {
    setBusy(true);
    setError(undefined);
    engine.convertLead(lead, {
      accountId: accMode === 'existing' ? accountId : undefined, createAccount: accMode === 'new',
      contactId: conMode === 'existing' ? contactId : undefined, createContact: conMode === 'new',
      createOpportunity: createOpp, opportunityName: oppName, amount, expectedCloseDate: close
    }).then(res => {
      notify('Lead converted.', MessageBarType.success);
      onClose(true);
      if (res.opportunityId) { const o = engine.find('opportunities', res.opportunityId); if (o) { open('opportunities', o); } }
    }).catch(e => { setBusy(false); setError(errorText(e)); });
  };

  return (
    <Dialog hidden={false} onDismiss={() => onClose()} minWidth={520} dialogContentProps={{ type: DialogType.largeHeader, title: `Convert ${lead.title}` }}>
      <Stack tokens={{ childrenGap: 12 }}>
        {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
        <ChoiceGroup label="Account" selectedKey={accMode} onChange={(_, o) => setAccMode(String(o?.key))} options={[
          { key: 'new', text: `Create account "${lead.company || ''}"`, disabled: !lead.company },
          { key: 'existing', text: 'Link to existing account' },
          { key: 'none', text: 'No account' }
        ]} />
        {accMode === 'existing' && <LookupPicker entity="accounts" value={accountId} onChange={setAccountId} />}
        <ChoiceGroup label="Contact" selectedKey={conMode} onChange={(_, o) => setConMode(String(o?.key))} options={[
          { key: 'new', text: `Create contact "${lead.title}"` },
          { key: 'existing', text: 'Link to existing contact' },
          { key: 'none', text: 'No contact' }
        ]} />
        {conMode === 'existing' && <LookupPicker entity="contacts" value={contactId} onChange={setContactId} />}
        <Checkbox label="Create an opportunity" checked={createOpp} onChange={(_, c) => setCreateOpp(!!c)} />
        {createOpp && (
          <Stack tokens={{ childrenGap: 8 }}>
            <TextField label="Opportunity name" value={oppName} onChange={(_, v) => setOppName(v || '')} />
            <TextField label={`Estimated amount (${engine.settings.currency})`} type="number" value={String(amount)} onChange={(_, v) => setAmount(Number(v) || 0)} />
            <DateTimeField label="Expected close" showTime={false} value={close} onChange={setClose} />
          </Stack>
        )}
      </Stack>
      <DialogFooter>
        <PrimaryButton text={busy ? 'Converting...' : 'Convert'} disabled={busy || (accMode === 'existing' && !accountId) || (conMode === 'existing' && !contactId)} onClick={run} />
        <DefaultButton text="Cancel" onClick={() => onClose()} />
      </DialogFooter>
    </Dialog>
  );
};

/* ---------------------------------------------------------------- email (Outlook via Graph) */

export const EmailDialog: React.FC<{ entity: EntityKey; record: IRecord; to?: string; subject?: string; body?: string; onClose: () => void }> =
  ({ entity, record, to: initialTo, subject: s0, body: b0, onClose }) => {
    const { engine, notify } = useCrm();
    const templates = engine.settings.emailTemplates;
    const [to, setTo] = React.useState(initialTo || record.email || '');
    const [cc, setCc] = React.useState('');
    const [subject, setSubject] = React.useState(s0 || '');
    const [body, setBody] = React.useState(b0 || '');
    const [busy, setBusy] = React.useState(false);
    const [error, setError] = React.useState<string>();
    const graphOn = engine.settings.integrations.graphEmail && engine.graph.available;

    const apply = (id: string): void => {
      const t = templates.find(x => x.id === id);
      if (t) { setSubject(engine.fillTemplate(t.subject, record)); setBody(engine.fillTemplate(t.body, record)); }
    };
    const send = (): void => {
      setBusy(true);
      setError(undefined);
      engine.sendEmail(to, subject, body, record.id ? { key: entity, id: record.id } : undefined, cc.split(/[;,]/).map(x => x.trim()).filter(Boolean))
        .then(() => { notify('Email sent from your Outlook mailbox and logged.', MessageBarType.success); onClose(); })
        .catch(e => { setBusy(false); setError(errorText(e)); });
    };
    const mailtoHref = `mailto:${encodeURIComponent(to)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;

    return (
      <Dialog hidden={false} onDismiss={onClose} minWidth={600} dialogContentProps={{ type: DialogType.normal, title: 'Send email' }}>
        <Stack tokens={{ childrenGap: 8 }}>
          {!graphOn && <MessageBar>Outlook integration is off. “Open in mail app” will open your default mail client.</MessageBar>}
          {error && <MessageBar messageBarType={MessageBarType.error} isMultiline>{error}</MessageBar>}
          {templates.length > 0 && (
            <Dropdown label="Template" placeholder="Choose a template" options={templates.map(t => ({ key: t.id, text: t.name }))} onChange={(_, o) => apply(String(o?.key))} />
          )}
          <TextField label="To" value={to} onChange={(_, v) => setTo(v || '')} required />
          <TextField label="Cc" value={cc} onChange={(_, v) => setCc(v || '')} placeholder="Separate addresses with ;" />
          <TextField label="Subject" value={subject} onChange={(_, v) => setSubject(v || '')} required />
          <TextField label="Message" multiline rows={10} value={body} onChange={(_, v) => setBody(v || '')} />
        </Stack>
        <DialogFooter>
          {graphOn && <PrimaryButton text={busy ? 'Sending...' : 'Send'} iconProps={{ iconName: 'Send' }} disabled={busy || !to || !subject} onClick={send} />}
          <DefaultButton text="Open in mail app" href={mailtoHref} target="_blank" onClick={() => {
            if (record.id) { engine.logActivity(entity, record.id, `Email: ${subject}`, `To: ${to}\n\n${body}`, 'Email').catch(() => undefined); }
            onClose();
          }} />
          <DefaultButton text="Cancel" onClick={onClose} />
        </DialogFooter>
      </Dialog>
    );
  };

/* ---------------------------------------------------------------- Teams meeting */

export const MeetingDialog: React.FC<{ demo: IRecord; onClose: () => void }> = ({ demo, onClose }) => {
  const { engine, notify } = useCrm();
  const [start, setStart] = React.useState<string | undefined>(demo.scheduledDate || demo.preferredDate);
  const [minutes, setMinutes] = React.useState<number>(demo.durationMinutes || 30);
  const [attendees, setAttendees] = React.useState<string>([demo.email].filter(Boolean).join('; '));
  const [subject, setSubject] = React.useState(`${engine.settings.company.name || 'Product'} demo - ${demo.company || demo.contactName}`);
  const [agenda, setAgenda] = React.useState(`Agenda:\n- Introductions\n- Your requirements\n- Live product demo\n- Q&A and next steps\n\n${demo.notes || ''}`);
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState<string>();
  const graphOn = engine.settings.integrations.graphMeetings && engine.graph.available;

  const go = (): void => {
    if (!start) { return; }
    setBusy(true);
    engine.scheduleDemo(demo, start, minutes, attendees.split(/[;,]/).map(x => x.trim()).filter(Boolean), subject, agenda)
      .then(() => { notify(graphOn ? 'Teams meeting created and invitations sent.' : 'Demo scheduled.', MessageBarType.success); onClose(); })
      .catch(e => { setBusy(false); setError(errorText(e)); });
  };

  return (
    <Dialog hidden={false} onDismiss={onClose} minWidth={560} dialogContentProps={{ type: DialogType.normal, title: 'Schedule demo' }}>
      <Stack tokens={{ childrenGap: 8 }}>
        {graphOn
          ? <MessageBar messageBarType={MessageBarType.info}>A Teams meeting is created in your Outlook calendar and invitations are sent to the attendees.</MessageBar>
          : <MessageBar>Teams integration is off. The demo is scheduled in the CRM only; paste a meeting link on the record.</MessageBar>}
        {error && <MessageBar messageBarType={MessageBarType.error} isMultiline>{error}</MessageBar>}
        <TextField label="Subject" value={subject} onChange={(_, v) => setSubject(v || '')} />
        <DateTimeField label="Start" value={start} onChange={setStart} required />
        <SpinButton label="Duration (minutes)" labelPosition={Position.top} min={15} max={240} step={15} value={String(minutes)} onChange={(_, v) => setMinutes(Number(v) || 30)} />
        <TextField label="Attendees" value={attendees} onChange={(_, v) => setAttendees(v || '')} placeholder="email; email" />
        <TextField label="Agenda" multiline rows={6} value={agenda} onChange={(_, v) => setAgenda(v || '')} />
      </Stack>
      <DialogFooter>
        <PrimaryButton text={busy ? 'Scheduling...' : 'Schedule'} disabled={busy || !start} onClick={go} />
        <DefaultButton text="Cancel" onClick={onClose} />
      </DialogFooter>
    </Dialog>
  );
};

/* ---------------------------------------------------------------- bulk update */

export const BulkEditDialog: React.FC<{ entity: EntityKey; ids: number[]; onClose: () => void }> = ({ entity, ids, onClose }) => {
  const { engine, notify, settings } = useCrm();
  const def = engine.def(entity);
  const editable = def.fields.filter(f => !f.readOnly && !f.hidden && (f.type === 'choice' || f.type === 'user' || f.type === 'boolean' || f.type === 'date' || f.type === 'datetime'));
  const [fieldKey, setFieldKey] = React.useState<string>(editable[0]?.key);
  const [value, setValue] = React.useState<any>();
  const [done, setDone] = React.useState(0);
  const [busy, setBusy] = React.useState(false);
  const fd = editable.find(f => f.key === fieldKey);

  const run = (): void => {
    setBusy(true);
    engine.bulkPatch(entity, ids, { [fieldKey]: value }, setDone)
      .then(() => { notify(`Updated ${ids.length} record(s).`, MessageBarType.success); onClose(); })
      .catch(e => { setBusy(false); notify(errorText(e), MessageBarType.error); });
  };

  return (
    <Dialog hidden={false} onDismiss={onClose} minWidth={460} dialogContentProps={{ type: DialogType.normal, title: `Update ${ids.length} ${def.plural.toLowerCase()}` }}>
      <Stack tokens={{ childrenGap: 10 }}>
        <Dropdown label="Field" selectedKey={fieldKey} options={editable.map(f => ({ key: f.key, text: f.label }))} onChange={(_, o) => { setFieldKey(String(o?.key)); setValue(undefined); }} />
        {fd && fd.type === 'choice' && <Dropdown label="New value" options={choicesFor(fd, settings).map(c => ({ key: c, text: c }))} selectedKey={value} onChange={(_, o) => setValue(o?.key)} />}
        {fd && fd.type === 'boolean' && <Checkbox label="Value" checked={!!value} onChange={(_, c) => setValue(!!c)} />}
        {fd && (fd.type === 'date' || fd.type === 'datetime') && <DateTimeField label="New value" value={value} showTime={fd.type === 'datetime'} onChange={setValue} />}
        {fd && fd.type === 'user' && <UserPicker label="New value" value={value} onChange={setValue} />}
        {busy && <ProgressIndicator percentComplete={done / ids.length} label={`${done} / ${ids.length}`} />}
      </Stack>
      <DialogFooter>
        <PrimaryButton text="Update" disabled={busy || !fd} onClick={run} />
        <DefaultButton text="Cancel" onClick={onClose} />
      </DialogFooter>
    </Dialog>
  );
};

/* ---------------------------------------------------------------- CSV import */

const norm = (s: string): string => s.toLowerCase().replace(/[^a-z0-9]/g, '');

export const ImportDialog: React.FC<{ entity: EntityKey; onClose: () => void }> = ({ entity, onClose }) => {
  const { engine, notify } = useCrm();
  const def = engine.def(entity);
  const importable = def.fields.filter(f => !f.readOnly && f.type !== 'json' && f.type !== 'user' && f.type !== 'lookup');
  const [rows, setRows] = React.useState<string[][]>([]);
  const [map, setMap] = React.useState<Record<number, string>>({});
  const [skipDup, setSkipDup] = React.useState(true);
  const [progress, setProgress] = React.useState<{ done: number; ok: number; skipped: number; failed: number }>();
  const [error, setError] = React.useState<string>();
  const headers = rows[0] || [];

  const onFile = (file?: File): void => {
    if (!file) { return; }
    const reader = new FileReader();
    reader.onload = () => {
      const parsed = parseCsv(String(reader.result || ''));
      if (parsed.length < 2) { setError('The file has no data rows.'); return; }
      setRows(parsed);
      const m: Record<number, string> = {};
      parsed[0].forEach((h, i) => {
        const n = norm(h);
        const hit = importable.find(f => norm(f.label) === n || norm(f.key) === n || norm(f.column) === n) ||
          (n === 'name' || n === 'fullname' ? importable.find(f => f.key === 'title') : undefined) ||
          (n === 'company' || n === 'organization' ? importable.find(f => f.key === 'company') : undefined);
        if (hit) { m[i] = hit.key; }
      });
      setMap(m);
    };
    reader.readAsText(file);
  };

  const convert = (key: string, raw: string): any => {
    const fd = importable.find(f => f.key === key)!;
    const v = raw.trim();
    if (!v) { return undefined; }
    switch (fd.type) {
      case 'number': case 'currency': case 'percent': return Number(v.replace(/[^\d.-]/g, '')) || 0;
      case 'boolean': return /^(1|y|yes|true|x)$/i.test(v);
      case 'date': case 'datetime': { const d = new Date(v); return isNaN(d.getTime()) ? undefined : d.toISOString(); }
      case 'tags': return v.split(/[;,]/).map(x => x.trim()).filter(Boolean);
      default: return v;
    }
  };

  const run = async (): Promise<void> => {
    await engine.ensure([entity]);
    const data = rows.slice(1);
    const st = { done: 0, ok: 0, skipped: 0, failed: 0 };
    for (const row of data) {
      const rec: IRecord = {};
      Object.keys(map).forEach(i => { if (map[Number(i)]) { const v = convert(map[Number(i)], row[Number(i)] || ''); if (v !== undefined) { rec[map[Number(i)]] = v; } } });
      def.fields.forEach(f => { if (f.defaultValue !== undefined && rec[f.key] === undefined) { rec[f.key] = f.defaultValue; } });
      if (entity === 'leads' && !rec.leadSource) { rec.leadSource = 'Other'; }
      try {
        if (skipDup && engine.findDuplicates(entity, rec).length) { st.skipped++; }
        else { await engine.save(entity, rec); st.ok++; }
      } catch {
        st.failed++;
      }
      st.done++;
      setProgress({ ...st });
    }
    notify(`Import finished: ${st.ok} created, ${st.skipped} duplicates skipped, ${st.failed} failed.`, st.failed ? MessageBarType.warning : MessageBarType.success);
  };

  const template = (): void => {
    const csv = importable.map(f => `"${f.label}"`).join(',') + '\r\n';
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob(['﻿' + csv], { type: 'text/csv' }));
    a.download = `${def.plural.toLowerCase().replace(/\s/g, '-')}-template.csv`;
    a.click();
  };

  return (
    <Dialog hidden={false} onDismiss={onClose} minWidth={640} dialogContentProps={{ type: DialogType.normal, title: `Import ${def.plural.toLowerCase()} from CSV` }}>
      <Stack tokens={{ childrenGap: 10 }}>
        {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
        <Text>Upload a CSV file (export from Excel, HubSpot, Salesforce, Dynamics, etc.). Columns are matched automatically; adjust below. <Link onClick={template}>Download template</Link></Text>
        <input type="file" accept=".csv,text/csv" onChange={e => onFile(e.target.files ? e.target.files[0] : undefined)} aria-label="CSV file" />
        {headers.length > 0 && (
          <>
            <Text variant="small">{rows.length - 1} row(s) found. Map the columns:</Text>
            <div className="crm-import-map">
              {headers.map((h, i) => (
                <React.Fragment key={i}>
                  <Text title={rows[1]?.[i]}>{h || `Column ${i + 1}`}</Text>
                  <Dropdown selectedKey={map[i] || ''} ariaLabel={`Map ${h}`}
                    options={[{ key: '', text: '— skip —' }, ...importable.map(f => ({ key: f.key, text: f.label }))]}
                    onChange={(_, o) => setMap({ ...map, [i]: String(o?.key || '') })} />
                </React.Fragment>
              ))}
            </div>
            <Checkbox label="Skip rows that look like duplicates (same email / name)" checked={skipDup} onChange={(_, c) => setSkipDup(!!c)} />
          </>
        )}
        {progress && <ProgressIndicator percentComplete={progress.done / Math.max(1, rows.length - 1)}
          label={`${progress.done} / ${rows.length - 1} · ${progress.ok} created · ${progress.skipped} skipped · ${progress.failed} failed`} />}
      </Stack>
      <DialogFooter>
        <PrimaryButton text="Import" disabled={!headers.length || !!progress} onClick={() => { run().catch(e => setError(errorText(e))); }} />
        <DefaultButton text={progress && progress.done === rows.length - 1 ? 'Close' : 'Cancel'} onClick={onClose} />
      </DialogFooter>
    </Dialog>
  );
};
