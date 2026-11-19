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
import { logWhatsApp, waLink, waNumber, waRecipient } from '../../services/WhatsApp';
import { audience, exportAudience, IAudienceFilter, MAX_PER_RUN, sendCampaign } from '../../services/Marketing';

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

export const EmailDialog: React.FC<{ entity: EntityKey; record: IRecord; to?: string; subject?: string; body?: string; onClose: () => void; onSent?: () => void }> =
  ({ entity, record, to: initialTo, subject: s0, body: b0, onClose, onSent }) => {
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
        .then(() => { notify('Email sent from your Outlook mailbox and logged.', MessageBarType.success); onSent?.(); onClose(); })
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
            onSent?.();
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

/* ---------------------------------------------------------------- WhatsApp */

export const WhatsAppDialog: React.FC<{ entity: EntityKey; record: IRecord; onClose: () => void; onSent?: () => void; templateId?: string }> = ({ entity, record, onClose, onSent, templateId }) => {
  const { engine, notify } = useCrm();
  const wa = engine.settings.whatsapp;
  const rcpt = waRecipient(engine, entity, record);
  const [to, setTo] = React.useState(rcpt.number ? `+${rcpt.number}` : '');
  const initialTpl = wa.templates.find(t => t.id === templateId);
  const [tplId, setTplId] = React.useState<string>(initialTpl ? initialTpl.id : '');
  const [text, setText] = React.useState(initialTpl ? engine.fillTemplate(initialTpl.body, record) : '');
  const [useTemplate, setUseTemplate] = React.useState(true);
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState<string>();
  const tpl = wa.templates.find(t => t.id === tplId);
  const number = waNumber(to, wa.defaultCountryCode);
  const cloud = wa.mode === 'cloud';
  const pick = (id: string): void => {
    setTplId(id);
    const t = wa.templates.find(x => x.id === id);
    if (t) { setText(engine.fillTemplate(t.body, record)); setUseTemplate(!!t.cloudName); }
  };
  const log = (asCloud: boolean): Promise<void> => logWhatsApp(engine, entity, record, number, text, { cloud: asCloud, template: asCloud && useTemplate ? tpl : undefined });
  const openChat = (): void => {
    // Open synchronously (pop-up blockers), then log.
    window.open(waLink(number, text), '_blank', 'noopener');
    log(false).then(() => { notify('WhatsApp opened and logged on the timeline.'); onSent?.(); onClose(); }).catch(e => setError(errorText(e)));
  };
  const send = (): void => {
    setBusy(true);
    log(true).then(() => { setBusy(false); notify('WhatsApp message queued. Power Automate sends it within a minute.'); onSent?.(); onClose(); })
      .catch(e => { setBusy(false); setError(errorText(e)); });
  };
  return (
    <Dialog hidden={false} onDismiss={onClose} minWidth={480} maxWidth={620}
      dialogContentProps={{ type: DialogType.normal, title: 'WhatsApp', subText: rcpt.name ? `To ${rcpt.name}` : undefined }}>
      <Stack tokens={{ childrenGap: 8 }}>
        {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
        <TextField label="WhatsApp number" value={to} placeholder={`+${wa.defaultCountryCode || ''}…`} onChange={(_, v) => setTo(v || '')}
          errorMessage={to && !number ? 'Enter the number with country code, e.g. +91 98765 43210' : undefined} required />
        <Dropdown label="Template" placeholder="Choose a template (optional)" selectedKey={tplId || null} options={wa.templates.map(t => ({ key: t.id, text: t.name }))} onChange={(_, o) => pick(String(o?.key))} />
        <TextField label="Message" multiline rows={5} value={text} onChange={(_, v) => setText(v || '')} />
        {cloud && tpl?.cloudName && (
          <Checkbox label={`Send as approved template “${tpl.cloudName}” (needed when the customer has not written in the last 24 hours)`} checked={useTemplate} onChange={(_, c) => setUseTemplate(!!c)} />
        )}
        {cloud && <Text variant="small" className="crm-muted">Sent from your WhatsApp Business number by the Power Automate flow “WhatsApp send”. Replies appear on the timeline.</Text>}
      </Stack>
      <DialogFooter>
        {cloud && <PrimaryButton iconProps={{ iconName: 'Send' }} text={busy ? 'Queuing…' : 'Send'} disabled={busy || !number || !text.trim()} onClick={send} />}
        {cloud ? <DefaultButton iconProps={{ iconName: 'Chat' }} text="Open in WhatsApp" disabled={!number} onClick={openChat} />
          : <PrimaryButton iconProps={{ iconName: 'Chat' }} text="Open WhatsApp" disabled={!number} onClick={openChat} />}
        <DefaultButton text="Cancel" onClick={onClose} />
      </DialogFooter>
    </Dialog>
  );
};

/* ---------------------------------------------------------------- merge duplicates */

export const MergeDialog: React.FC<{ entity: EntityKey; master: IRecord; duplicate: IRecord; onClose: (merged?: IRecord) => void }> = ({ entity, master, duplicate, onClose }) => {
  const { engine, notify } = useCrm();
  const def = engine.def(entity);
  const show = (v: unknown, fd: { type: string; lookup?: string }): string => {
    if (v === undefined || v === null || v === '') { return '—'; }
    if (typeof v === 'object' && 'title' in (v as any)) { return (v as any).title; }
    if (fd.type === 'lookup') { return engine.display(fd.lookup as EntityKey, v as number) || String(v); }
    if (Array.isArray(v)) { return v.join(', '); }
    if (typeof v === 'boolean') { return v ? 'Yes' : 'No'; }
    if (fd.type.indexOf('date') === 0) { return new Date(String(v)).toLocaleDateString(); }
    return String(v);
  };
  const diffs = def.fields.filter(fd => !fd.readOnly && !fd.hidden && fd.type !== 'json' && show(master[fd.key], fd) !== show(duplicate[fd.key], fd));
  const [take, setTake] = React.useState<Record<string, boolean>>({});
  const [busy, setBusy] = React.useState<string>();
  const [error, setError] = React.useState<string>();
  const merge = (): void => {
    setBusy('Merging…');
    engine.mergeRecords(entity, master, duplicate, Object.keys(take).filter(k => take[k]), m => setBusy(`Moving linked ${m}`))
      .then(r => { notify(`Merged into ${r.title || 'the record'}.`, MessageBarType.success); onClose(r); })
      .catch(e => { setBusy(undefined); setError(errorText(e)); });
  };
  return (
    <Dialog hidden={false} onDismiss={() => onClose()} minWidth={640} maxWidth={820}
      dialogContentProps={{ type: DialogType.normal, title: `Merge duplicate ${def.singular.toLowerCase()}`, subText: `“${duplicate.title}” is merged into “${master.title}” and then deleted. Its activities, deals, quotes, cases and other links move to “${master.title}”.` }}>
      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
      {diffs.length ? (
        <div className="crm-merge-grid" role="table">
          <Text className="crm-muted" role="columnheader">Field</Text><Text className="crm-muted" role="columnheader">Keep (this record)</Text><Text className="crm-muted" role="columnheader">Duplicate</Text>
          {diffs.map(fd => (
            <React.Fragment key={fd.key}>
              <Text role="rowheader" styles={{ root: { fontWeight: 600 } }}>{fd.label}</Text>
              <ChoiceGroup selectedKey={take[fd.key] ? 'dup' : 'master'} options={[{ key: 'master', text: show(master[fd.key], fd) }]} onChange={() => setTake({ ...take, [fd.key]: false })} aria-label={`${fd.label}: keep`} />
              <ChoiceGroup selectedKey={take[fd.key] ? 'dup' : 'master'} options={[{ key: 'dup', text: show(duplicate[fd.key], fd) }]} onChange={() => setTake({ ...take, [fd.key]: true })} aria-label={`${fd.label}: duplicate`} />
            </React.Fragment>
          ))}
        </div>
      ) : <Text>The two records have the same values.</Text>}
      <Text variant="small" className="crm-muted" block style={{ marginTop: 8 }}>Empty fields on this record are filled from the duplicate automatically.</Text>
      {busy && <ProgressIndicator label={busy} />}
      <DialogFooter>
        <PrimaryButton iconProps={{ iconName: 'Merge' }} text="Merge" disabled={!!busy} onClick={merge} />
        <DefaultButton text="Cancel" disabled={!!busy} onClick={() => onClose()} />
      </DialogFooter>
    </Dialog>
  );
};

/* ---------------------------------------------------------------- call (Teams / phone) with logging */

export const CallDialog: React.FC<{ entity: EntityKey; record: IRecord; onClose: () => void }> = ({ entity, record, onClose }) => {
  const { engine, notify } = useCrm();
  const cc = engine.settings.whatsapp.defaultCountryCode;
  const contact = record.contactRef ? engine.find('contacts', record.contactRef) : undefined;
  const raw = record.mobile || record.phone || contact?.mobile || contact?.phone || '';
  const [number, setNumber] = React.useState(raw ? `+${waNumber(raw, cc) || String(raw).replace(/[^\d+]/g, '')}` : '');
  const [started, setStarted] = React.useState<number>();
  const [now, setNow] = React.useState(Date.now());
  const [outcome, setOutcome] = React.useState('Connected');
  const [notes, setNotes] = React.useState('');
  const [followUp, setFollowUp] = React.useState<string>();
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState<string>();
  React.useEffect(() => { if (!started) { return undefined; } const t = setInterval(() => setNow(Date.now()), 1000); return () => clearInterval(t); }, [started]);
  const e164 = waNumber(number, cc);
  const secs = started ? Math.max(0, Math.round((now - started) / 1000)) : 0;
  const mmss = `${Math.floor(secs / 60)}:${('0' + (secs % 60)).slice(-2)}`;
  const start = (): void => { if (!started) { setStarted(Date.now()); setNow(Date.now()); } };
  const save = (): void => {
    setBusy(true);
    const minutes = Math.max(started ? 1 : 0, Math.round(secs / 60));
    (async () => {
      await engine.logActivity(entity, record.id!, `Call: ${outcome}${record.title ? ` – ${record.title}` : ''}`, `${e164 ? `+${e164}` : number}${started ? ` · ${mmss} min` : ''}\n${notes}`.trim(), 'Call');
      if (followUp) {
        await engine.save('activities', { title: `Follow up: ${record.title || ''}`.slice(0, 255), activityType: 'Task', completed: false, dueDate: followUp, relatedType: engine.def(entity).activityName, relatedId: record.id, owner: engine.me, details: notes });
      }
      if (entity === 'leads' && record.status === 'New' && outcome !== 'Wrong number') { await engine.patch('leads', record.id!, { status: 'Contacted' }); }
      return minutes;
    })().then(() => { notify('Call logged.', MessageBarType.success); onClose(); }).catch(e => { setBusy(false); setError(errorText(e)); });
  };
  return (
    <Dialog hidden={false} onDismiss={onClose} minWidth={460} maxWidth={560} dialogContentProps={{ type: DialogType.normal, title: `Call ${record.title || ''}` }}>
      <Stack tokens={{ childrenGap: 8 }}>
        {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
        <TextField label="Number" value={number} onChange={(_, v) => setNumber(v || '')} />
        <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="center">
          <PrimaryButton iconProps={{ iconName: 'TeamsLogo' }} text="Call with Teams" disabled={!e164} href={`https://teams.microsoft.com/l/call/0/0?users=4:+${e164}`} target="_blank" onClick={start} />
          <DefaultButton iconProps={{ iconName: 'Phone' }} text="Call with phone" disabled={!e164} href={`tel:+${e164}`} onClick={start} />
          {started ? <Text styles={{ root: { fontVariantNumeric: 'tabular-nums' } }} aria-live="polite">⏱ {mmss}</Text> : <Link onClick={start}>Start timer</Link>}
        </Stack>
        <ChoiceGroup label="Outcome" selectedKey={outcome} onChange={(_, o) => setOutcome(String(o?.key))}
          options={['Connected', 'Left voicemail', 'No answer', 'Busy', 'Wrong number'].map(x => ({ key: x, text: x }))} styles={{ flexContainer: { display: 'flex', flexWrap: 'wrap', gap: '0 14px' } }} />
        <TextField label="Notes" multiline rows={3} value={notes} onChange={(_, v) => setNotes(v || '')} />
        <DateTimeField label="Follow-up task (optional)" value={followUp} onChange={v => setFollowUp(v)} />
        <Text variant="small" className="crm-muted">Teams calls to phone numbers need Teams Phone. Calls are logged on the timeline when you click Log call.</Text>
      </Stack>
      <DialogFooter>
        <PrimaryButton text={busy ? 'Saving…' : 'Log call'} disabled={busy} onClick={save} />
        <DefaultButton text="Cancel" onClick={onClose} />
      </DialogFooter>
    </Dialog>
  );
};

/* ---------------------------------------------------------------- email campaign */

export const CampaignEmailDialog: React.FC<{ campaign: IRecord; onClose: () => void }> = ({ campaign, onClose }) => {
  const { engine, notify, settings } = useCrm();
  React.useEffect(() => { engine.ensure(['leads', 'contacts']).catch(() => undefined); }, []);
  const [f, setF] = React.useState<IAudienceFilter>({ leads: true, contacts: true, tag: '', leadSource: '', country: '', campaignLeadsOnly: false });
  const [subject, setSubject] = React.useState(campaign.title || '');
  const [body, setBody] = React.useState('Hi {{firstName}},\n\n\n\nBest regards,\n{{senderName}}\n{{companyName}}');
  const [progress, setProgress] = React.useState<{ sent: number; failed: number }>();
  const [error, setError] = React.useState<string>();
  const stop = React.useRef(false);
  const aud = audience(engine, f, campaign.id);
  const n = Math.min(aud.ok.length, MAX_PER_RUN);
  const minutes = Math.ceil((n * Math.max(1, settings.marketing.secondsBetweenEmails)) / 60);
  const send = (): void => {
    if (!window.confirm(`Send “${subject}” to ${n} recipient(s) from your mailbox? This takes about ${minutes} minute(s); keep this window open.`)) { return; }
    stop.current = false;
    setProgress({ sent: 0, failed: 0 });
    sendCampaign(engine, campaign, aud.ok, subject, body, (sent, failed) => setProgress({ sent, failed }), () => stop.current)
      .then(r => { notify(`Campaign email sent to ${r.sent} recipient(s)${r.failed.length ? `, ${r.failed.length} failed` : ''}.`, r.failed.length ? MessageBarType.warning : MessageBarType.success); onClose(); })
      .catch(e => { setProgress(undefined); setError(errorText(e)); });
  };
  const busy = !!progress;
  return (
    <Dialog hidden={false} onDismiss={busy ? undefined : onClose} minWidth={620} maxWidth={760} modalProps={{ isBlocking: busy }}
      dialogContentProps={{ type: DialogType.normal, title: `Email campaign: ${campaign.title}` }}>
      <Stack tokens={{ childrenGap: 8 }}>
        {error && <MessageBar messageBarType={MessageBarType.error} isMultiline>{error}</MessageBar>}
        <Stack horizontal wrap tokens={{ childrenGap: 16 }}>
          <Checkbox label="Contacts" checked={f.contacts} disabled={busy} onChange={(_, c) => setF({ ...f, contacts: !!c })} />
          <Checkbox label="Open leads" checked={f.leads} disabled={busy} onChange={(_, c) => setF({ ...f, leads: !!c })} />
          <Checkbox label="Only leads of this campaign" checked={f.campaignLeadsOnly} disabled={busy || !f.leads} onChange={(_, c) => setF({ ...f, campaignLeadsOnly: !!c })} />
        </Stack>
        <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
          <TextField label="Tag" value={f.tag} disabled={busy} onChange={(_, v) => setF({ ...f, tag: v || '' })} styles={{ root: { width: 160 } }} />
          <Dropdown label="Lead source" selectedKey={f.leadSource} disabled={busy || !f.leads} options={[{ key: '', text: 'Any' }, ...settings.leadSources.map(x => ({ key: x, text: x }))]}
            onChange={(_, o) => setF({ ...f, leadSource: String(o?.key || '') })} styles={{ root: { width: 180 } }} />
          <TextField label="Country" value={f.country} disabled={busy} onChange={(_, v) => setF({ ...f, country: v || '' })} styles={{ root: { width: 140 } }} />
        </Stack>
        <MessageBar messageBarType={aud.ok.length ? MessageBarType.info : MessageBarType.warning} isMultiline>
          <b>{aud.ok.length}</b> recipient(s) with marketing consent. Excluded: {aud.excluded.noConsent} without opt-in, {aud.excluded.unsubscribed} unsubscribed, {aud.excluded.noEmail} without email.
          {aud.ok.length > MAX_PER_RUN && <> Only the first {MAX_PER_RUN} are sent per run (Outlook limits); export the audience for larger sends.</>}
        </MessageBar>
        <TextField label="Subject" value={subject} disabled={busy} onChange={(_, v) => setSubject(v || '')} required />
        <TextField label="Message" multiline rows={8} value={body} disabled={busy} onChange={(_, v) => setBody(v || '')} />
        <Text variant="small" className="crm-muted">Placeholders such as {'{{firstName}}'} are filled per person. Your company footer and a personal unsubscribe link are added automatically.</Text>
        {progress && <ProgressIndicator label={`Sending… ${progress.sent} of ${n}${progress.failed ? ` (${progress.failed} failed)` : ''}`} percentComplete={n ? (progress.sent + progress.failed) / n : 0} />}
      </Stack>
      <DialogFooter>
        {!busy && <PrimaryButton iconProps={{ iconName: 'Send' }} text={`Send to ${n}`} disabled={!n || !subject.trim() || !body.trim()} onClick={send} />}
        {!busy && <DefaultButton iconProps={{ iconName: 'Download' }} text="Export for Mailchimp / Brevo" disabled={!aud.ok.length} onClick={() => exportAudience(aud.ok, campaign)} />}
        {busy ? <DefaultButton text="Stop" onClick={() => { stop.current = true; }} /> : <DefaultButton text="Cancel" onClick={onClose} />}
      </DialogFooter>
    </Dialog>
  );
};
