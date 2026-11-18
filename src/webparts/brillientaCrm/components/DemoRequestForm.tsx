import * as React from 'react';
import { MessageBar, MessageBarType, PrimaryButton, Stack, Text, TextField, Icon, DefaultButton, Checkbox } from '@fluentui/react';
import { Engine } from '../services/Engine';
import { DateTimeField } from './core/controls';
import { errorText } from './utils';

interface IProps {
  engine: Engine;
  title: string;
  intro: string;
  productName: string;
}

interface IForm {
  firstName: string; lastName: string; email: string; phone: string; company: string; jobTitle: string; country: string;
  preferredDate?: string; message: string; consent: boolean;
}

const EMPTY: IForm = { firstName: '', lastName: '', email: '', phone: '', company: '', jobTitle: '', country: '', message: '', consent: false };
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

/** Public-facing "Request a demo" form: creates a scored, auto-assigned Lead + a Demo request. */
export const DemoRequestForm: React.FC<IProps> = ({ engine, title, intro, productName }) => {
  const [f, setF] = React.useState<IForm>(EMPTY);
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState<string>();
  const [done, setDone] = React.useState(false);
  const [touched, setTouched] = React.useState(false);
  const set = (p: Partial<IForm>): void => setF({ ...f, ...p });
  const invalid = !f.firstName || !f.lastName || !EMAIL_RE.test(f.email) || !f.company;

  const submit = async (): Promise<void> => {
    setTouched(true);
    if (invalid) { return; }
    if (!engine.license.canWrite) { setError('Demo requests are temporarily unavailable.'); return; }
    setBusy(true);
    setError(undefined);
    try {
      const ds = engine.ds;
      const name = `${f.firstName} ${f.lastName}`.trim();
      const lead = {
        title: name, firstName: f.firstName, lastName: f.lastName, email: f.email, phone: f.phone, company: f.company,
        jobTitle: f.jobTitle, country: f.country, leadSource: 'Website Demo Form', actionCode: 'DemoRequest', offer: productName,
        description: f.message, status: 'New', rating: 'Hot', owner: undefined as unknown
      };
      const scored = { ...lead, leadScore: engine.scoreLead(lead) };
      if (engine.settings.assignment.autoAssignNewLeads && engine.settings.assignment.roundRobin.length) {
        const rr = engine.settings.assignment;
        const idx = (rr.lastIndex + 1) % rr.roundRobin.length;
        scored.owner = rr.roundRobin[idx];
        ds.saveSettings({ ...engine.settings, assignment: { ...rr, lastIndex: idx } }).catch(() => undefined);
      }
      const leadId = await ds.save('leads', scored);
      const demoId = await ds.save('demos', {
        title: `Demo - ${f.company}`, contactName: name, email: f.email, phone: f.phone, company: f.company, leadRef: leadId,
        preferredDate: f.preferredDate, meetingLink: '', status: 'Requested', notes: f.message, durationMinutes: 30, owner: scored.owner
      });
      ds.writeAudit('Created', 'leads', leadId, name, { Source: ['', 'Demo request form'] }).catch(() => undefined);
      engine.webhooks.emit('demo.requested', { leadId, demoId, name, email: f.email, company: f.company, preferredDate: f.preferredDate, message: f.message },
        `New demo request: ${name} (${f.company})`);
      setDone(true);
      setF(EMPTY);
      setTouched(false);
    } catch (e) {
      setError(errorText(e));
    } finally {
      setBusy(false);
    }
  };

  if (done) {
    return (
      <Stack className="crm-card" tokens={{ childrenGap: 12 }} horizontalAlign="center" style={{ padding: 32 }}>
        <Icon iconName="CompletedSolid" style={{ fontSize: 40, color: 'var(--good)' }} />
        <Text variant="xLarge">Thank you!</Text>
        <Text>Your demo request was received. Our team will contact you shortly to confirm a time.</Text>
        <DefaultButton text="Submit another request" onClick={() => setDone(false)} />
      </Stack>
    );
  }
  const req = (v: string): string | undefined => (touched && !v ? 'Required' : undefined);
  return (
    <Stack className="crm-card" tokens={{ childrenGap: 12 }}>
      <Text variant="xLarge" styles={{ root: { fontWeight: 600 } }}>{title}</Text>
      {intro && <Text>{intro}</Text>}
      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
      <div className="crm-form-grid">
        <TextField label="First name" required value={f.firstName} errorMessage={req(f.firstName)} onChange={(_, v) => set({ firstName: v || '' })} />
        <TextField label="Last name" required value={f.lastName} errorMessage={req(f.lastName)} onChange={(_, v) => set({ lastName: v || '' })} />
        <TextField label="Work email" type="email" required value={f.email} errorMessage={touched && !EMAIL_RE.test(f.email) ? 'Enter a valid email' : undefined} onChange={(_, v) => set({ email: v || '' })} />
        <TextField label="Phone" value={f.phone} onChange={(_, v) => set({ phone: v || '' })} />
        <TextField label="Company" required value={f.company} errorMessage={req(f.company)} onChange={(_, v) => set({ company: v || '' })} />
        <TextField label="Job title" value={f.jobTitle} onChange={(_, v) => set({ jobTitle: v || '' })} />
        <TextField label="Country" value={f.country} onChange={(_, v) => set({ country: v || '' })} />
      </div>
      <DateTimeField label="Preferred date & time" value={f.preferredDate} onChange={v => set({ preferredDate: v })} />
      <TextField label="What would you like to see?" multiline rows={4} value={f.message} onChange={(_, v) => set({ message: v || '' })} />
      <Checkbox label="I agree to be contacted about this request." checked={f.consent} onChange={(_, c) => set({ consent: !!c })} />
      <Stack horizontal><PrimaryButton text={busy ? 'Sending...' : 'Request demo'} disabled={busy || !f.consent} onClick={() => { submit().catch(() => undefined); }} /></Stack>
    </Stack>
  );
};
