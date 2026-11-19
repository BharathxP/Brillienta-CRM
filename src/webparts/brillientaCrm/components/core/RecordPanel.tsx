/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
import {
  Panel, PanelType, PrimaryButton, DefaultButton, Stack, Pivot, PivotItem, MessageBar, MessageBarType, Dialog, DialogFooter,
  DialogType, Link, Text, CommandBar, ICommandBarItemProps, IconButton, Spinner, SpinnerSize, Dropdown, TextField
} from '@fluentui/react';
import { EntityKey, IFieldDef, IRecord, LOST, WON } from '../../models/model';
import { addPeriod, ILineItem, quoteTotals } from '../../services/Engine';
import { useCrm } from '../CrmContext';
import { FieldEditor, renderValue, choicesFor } from './fields';
import { Pill, ScoreBadge } from './controls';
import { ActivityTimeline } from './ActivityTimeline';
import { ConvertLeadDialog, EmailDialog, MeetingDialog, WhatsAppDialog } from './dialogs';
import { waRecipient } from '../../services/WhatsApp';
import { QuoteLines, printQuote } from './QuoteLines';
import { SmartSummary } from './SmartSummary';
import { draftFollowUp, INSIGHT_ENTITIES } from '../../services/Insights';
import { errorText, firstResponseLate, fmtDate, fmtMoney, slaBreached, sum } from '../utils';

const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

const RelatedList: React.FC<{ entity: EntityKey; rows: IRecord[]; loaded: boolean; onNew?: () => void }> = ({ entity, rows, loaded, onNew }) => {
  const { engine, open, currency } = useCrm();
  React.useEffect(() => { engine.ensure([entity]).catch(() => undefined); }, [entity]);
  const def = engine.def(entity);
  const cols = def.fields.filter(f => f.grid && f.key !== 'title').slice(0, 3);
  return (
    <div style={{ paddingTop: 12 }}>
      {onNew && <DefaultButton iconProps={{ iconName: 'Add' }} text={`New ${def.singular.toLowerCase()}`} onClick={onNew} style={{ marginBottom: 8 }} />}
      {!loaded && <Spinner size={SpinnerSize.small} />}
      {loaded && rows.length === 0 && <Text className="crm-muted" block>None yet.</Text>}
      {rows.map(r => (
        <div key={r.id} className="crm-related-row" role="button" tabIndex={0} onClick={() => open(entity, r)} onKeyDown={e => { if (e.key === 'Enter') { open(entity, r); } }}>
          <Text styles={{ root: { fontWeight: 600 } }}>{r.title || r.quoteNumber || `#${r.id}`}</Text>
          <Stack horizontal wrap tokens={{ childrenGap: 10 }}>
            {cols.map(c => <Text key={c.key} variant="small" className="crm-muted">{renderValue(c, r, engine, currency)}</Text>)}
          </Stack>
        </div>
      ))}
    </div>
  );
};

interface IProps {
  entity: EntityKey;
  record: IRecord;
  onClose: () => void;
  canGoBack: boolean;
}

export const RecordPanel: React.FC<IProps> = ({ entity, record, onClose, canGoBack }) => {
  const { engine, open, notify, currency, settings, data, isNarrow } = useCrm();
  const def = engine.def(entity);
  const isNew = !record.id;
  const live = (record.id && engine.find(entity, record.id)) || record;
  const [draft, setDraft] = React.useState<IRecord>(() => {
    if (!isNew) { return { ...live }; }
    const init: IRecord = {};
    def.fields.forEach(f => { if (f.defaultValue !== undefined) { init[f.key] = f.defaultValue; } });
    if (def.fields.some(f => f.key === 'owner')) { init.owner = engine.me; }
    if (entity === 'quotes' || entity === 'invoices') { init.lineItems = init.lineItems || []; }
    return { ...init, ...record };
  });
  const [dirty, setDirty] = React.useState(false);
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState<string>();
  const [errors, setErrors] = React.useState<Record<string, string>>({});
  const [confirmDelete, setConfirmDelete] = React.useState(false);
  const [dialog, setDialog] = React.useState<'convert' | 'email' | 'meeting' | 'lost' | 'emailQuote' | 'emailInvoice' | 'emailCase' | 'emailSign' | 'draft' | 'whatsapp'>();
  const [signLink, setSignLink] = React.useState('');
  const [history, setHistory] = React.useState<IRecord[]>();
  const [tab, setTab] = React.useState('details');

  // keep in sync with background refreshes when nothing was edited
  React.useEffect(() => {
    if (!dirty && !isNew && live !== record) { setDraft({ ...live }); }
  }, [live]);

  const canEdit = engine.can(isNew ? 'create' : 'edit', entity);
  const set = (key: string, v: any): void => { setDraft(d => ({ ...d, [key]: v })); setDirty(true); };

  const duplicates = (entity === 'leads' || entity === 'contacts' || entity === 'accounts') && (isNew || dirty)
    ? engine.findDuplicates(entity, { ...draft, title: entity === 'accounts' ? draft.title : `${draft.firstName || ''} ${draft.lastName || ''}`.trim() })
    : [];

  const validate = (): boolean => {
    const e: Record<string, string> = {};
    def.fields.forEach(f => {
      const v = draft[f.key];
      if (f.required && !f.readOnly && (v === undefined || v === null || v === '')) { e[f.key] = 'Required'; }
      if (f.type === 'email' && v && !EMAIL_RE.test(String(v))) { e[f.key] = 'Invalid email address'; }
    });
    setErrors(e);
    return !Object.keys(e).length;
  };

  const save = (close: boolean, override?: IRecord): void => {
    const rec = override || draft;
    if (!override && !validate()) { setTab('details'); return; }
    if (entity === 'opportunities' && rec.stage === LOST && !rec.lostReason) { setDialog('lost'); return; }
    setBusy(true);
    setError(undefined);
    engine.save(entity, rec).then(saved => {
      setBusy(false);
      setDirty(false);
      notify(`${def.singular} ${isNew ? 'created' : 'saved'}.`, MessageBarType.success);
      if (close) { onClose(); } else { setDraft({ ...saved }); if (isNew) { onClose(); open(entity, saved); } }
    }).catch(e => { setBusy(false); setError(errorText(e)); });
  };

  const quickPatch = (fields: Record<string, unknown>, msg: string): void => {
    if (dirty) { save(false, { ...draft, ...fields }); return; }
    setBusy(true);
    engine.patch(entity, draft.id!, fields).then(saved => { setBusy(false); setDraft({ ...saved }); notify(msg, MessageBarType.success); })
      .catch(e => { setBusy(false); setError(errorText(e)); });
  };

  /** Runs a composite action with the busy indicator; the promise resolves to a success message. */
  const run = (fn: () => Promise<string>): void => {
    setBusy(true);
    fn().then(msg => { setBusy(false); notify(msg, MessageBarType.success); }).catch(e => { setBusy(false); setError(errorText(e)); });
  };

  const close = (): void => {
    if (dirty && !window.confirm('Discard unsaved changes?')) { return; }
    onClose();
  };

  const remove = (): void => {
    setBusy(true);
    engine.remove(entity, draft.id!).then(() => { notify(`${def.singular} deleted (moved to the recycle bin).`, MessageBarType.success); onClose(); })
      .catch(e => { setBusy(false); setError(errorText(e)); });
  };

  const loadHistory = (): void => {
    if (!history && draft.id && engine.hasFeature('audit')) {
      engine.ds.getAudit(entity, draft.id, 200).then(setHistory).catch(() => setHistory([]));
    }
  };

  // ------------------------------------------------------------------ commands
  const cmds: ICommandBarItemProps[] = [];
  if (!isNew) {
    const writable = engine.can('edit', entity);
    if (draft.email && (entity !== 'quotes')) {
      cmds.push({ key: 'email', text: 'Email', iconProps: { iconName: 'Mail' }, onClick: () => setDialog('email') });
    }
    if (entity === 'leads' && writable) {
      cmds.push({ key: 'convert', text: 'Convert', iconProps: { iconName: 'Switch' }, disabled: draft.status === 'Converted', onClick: () => setDialog('convert') });
      cmds.push({ key: 'demo', text: 'Request demo', iconProps: { iconName: 'Presentation' }, onClick: () => open('demos', {
        contactName: draft.title, email: draft.email, phone: draft.phone, company: draft.company, leadRef: draft.id, notes: draft.description, preferredDate: undefined
      }) });
      cmds.push({ key: 'dq', text: 'Disqualify', iconProps: { iconName: 'Blocked' }, disabled: draft.status === 'Unqualified' || draft.status === 'Converted', onClick: () => quickPatch({ status: 'Unqualified' }, 'Lead disqualified.') });
    }
    if (entity === 'demos' && writable) {
      cmds.push({ key: 'meet', text: 'Schedule Teams meeting', iconProps: { iconName: 'TeamsLogo' }, onClick: () => setDialog('meeting') });
      if (/^https:\/\//i.test(draft.meetingLink || '')) {
        cmds.push({ key: 'join', text: 'Join', iconProps: { iconName: 'Video' }, href: draft.meetingLink, target: '_blank' });
      }
      cmds.push({ key: 'done', text: 'Completed', iconProps: { iconName: 'CheckMark' }, disabled: draft.status === 'Completed', onClick: () => quickPatch({ status: 'Completed', outcome: draft.outcome || 'Interested' }, 'Demo completed.') });
      cmds.push({ key: 'noshow', text: 'No-show', iconProps: { iconName: 'UserRemove' }, onClick: () => quickPatch({ status: 'No-show' }, 'Marked as no-show.') });
      cmds.push({ key: 'opp', text: 'Create opportunity', iconProps: { iconName: 'Money' }, onClick: () => open('opportunities', {
        title: `${draft.company || draft.contactName} - after demo`, company: draft.company, contactName: draft.contactName, email: draft.email,
        phone: draft.phone, demoRef: draft.id, leadRef: draft.leadRef, accountRef: draft.accountRef, contactRef: draft.contactRef,
        stage: engine.openStages()[2] || engine.openStages()[0], owner: draft.owner || engine.me
      }) });
    }
    if (entity === 'opportunities' && writable) {
      cmds.push({ key: 'won', text: 'Won', iconProps: { iconName: 'Trophy2' }, disabled: draft.stage === WON, onClick: () => save(false, { ...draft, stage: WON }) });
      cmds.push({ key: 'lost', text: 'Lost', iconProps: { iconName: 'Cancel' }, disabled: draft.stage === LOST, onClick: () => { setDraft({ ...draft, stage: LOST }); setDialog('lost'); } });
      if (engine.can('create', 'quotes')) {
        cmds.push({ key: 'quote', text: 'New quote', iconProps: { iconName: 'Document' }, onClick: () => open('quotes', {
          title: draft.title, opportunityRef: draft.id, accountRef: draft.accountRef, contactRef: draft.contactRef, lineItems: [], owner: draft.owner || engine.me
        }) });
      }
    }
    if (entity === 'quotes') {
      const appr = engine.quoteApproval(draft);
      cmds.push({ key: 'print', text: 'Print / PDF', iconProps: { iconName: 'Print' }, onClick: () => printQuote(draft, engine) });
      if (writable) {
        if (appr.required && draft.approvalStatus !== 'Approved') {
          cmds.push({ key: 'approval', text: draft.approvalStatus === 'Pending' ? 'Approval pending' : 'Request approval', iconProps: { iconName: 'Shield' },
            disabled: draft.approvalStatus === 'Pending' || dirty, onClick: () => {
              setBusy(true);
              engine.requestQuoteApproval(draft).then(s => { setBusy(false); setDraft({ ...s }); notify('Approval requested. The approver is notified by Power Automate.', MessageBarType.success); })
                .catch(e => { setBusy(false); setError(errorText(e)); });
            } });
        }
        cmds.push({ key: 'emailq', text: 'Email quote', iconProps: { iconName: 'Mail' }, disabled: !appr.approved, onClick: () => setDialog('emailQuote') });
        if (draft.quoteStatus !== 'Accepted' && draft.quoteStatus !== 'Rejected' && draft.signatureStatus !== 'Signed') {
          cmds.push({ key: 'sign', text: draft.signatureStatus === 'Sent' || draft.signatureStatus === 'Viewed' ? 'Resend for signature' : 'Send for e-signature', iconProps: { iconName: 'PenWorkspace' },
            disabled: !appr.approved || dirty, onClick: () => {
              const suggested = draft.signatureEmail || engine.find('contacts', draft.contactRef)?.email || '';
              const email = window.prompt('Signer email address', suggested);
              if (!email) { return; }
              if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email.trim())) { setError('Enter a valid email address for the signer.'); return; }
              setBusy(true);
              engine.sendForSignature(draft, email).then(r => { setBusy(false); setDraft({ ...r.quote }); setSignLink(r.link); setDialog('emailSign'); })
                .catch(e => { setBusy(false); setError(errorText(e)); });
            } });
        }
        if ((draft.signatureStatus === 'Sent' || draft.signatureStatus === 'Viewed') && engine.signingLink(draft)) {
          cmds.push({ key: 'copysign', text: 'Copy signing link', iconProps: { iconName: 'Link' }, onClick: () => {
            const link = engine.signingLink(draft);
            (navigator.clipboard ? navigator.clipboard.writeText(link) : Promise.reject(new Error('no clipboard')))
              .then(() => notify('Signing link copied.', MessageBarType.success)).catch(() => window.prompt('Copy the signing link', link));
          } });
        }
        cmds.push({ key: 'accept', text: 'Accepted', iconProps: { iconName: 'Accept' }, disabled: draft.quoteStatus === 'Accepted' || !appr.approved, onClick: () => {
          const closeOpp = !!draft.opportunityRef && window.confirm('Mark the linked opportunity as Closed Won with the quote total?');
          setBusy(true);
          const recurring = (draft.lineItems || []).filter((i: ILineItem) => i.billing === 'Monthly' || i.billing === 'Annual').length;
          const makeSubs = recurring > 0 && engine.settings.invoice.createSubscriptionsOnAccept && engine.can('create', 'subscriptions') &&
            window.confirm(`Create ${recurring} subscription${recurring === 1 ? '' : 's'} for the recurring line${recurring === 1 ? '' : 's'} (renewals are then tracked automatically)?`);
          engine.acceptQuote(draft, closeOpp)
            .then(() => (makeSubs ? engine.createSubscriptionsFromQuote(draft) : Promise.resolve([])))
            .then(subs => { setBusy(false); setDraft({ ...(engine.find('quotes', draft.id) || draft) }); notify(subs.length ? `Quote accepted · ${subs.length} subscription(s) created.` : 'Quote accepted.', MessageBarType.success); })
            .catch(e => { setBusy(false); setError(errorText(e)); });
        } });
        cmds.push({ key: 'reject', text: 'Rejected', iconProps: { iconName: 'Cancel' }, onClick: () => quickPatch({ quoteStatus: 'Rejected' }, 'Quote rejected.') });
        if (draft.quoteStatus === 'Accepted' && engine.can('create', 'invoices')) {
          cmds.push({ key: 'inv', text: 'Create invoice', iconProps: { iconName: 'Money' }, disabled: dirty, onClick: () => run(() => engine.createInvoiceFromQuote(draft).then(inv => { open('invoices', inv); return 'Invoice created.'; })) });
        }
        if (draft.quoteStatus === 'Accepted' && engine.can('create', 'subscriptions') && !(data.subscriptions || []).some(x => x.quoteRef === draft.id) && (draft.lineItems || []).some((i: ILineItem) => i.billing === 'Monthly' || i.billing === 'Annual')) {
          cmds.push({ key: 'subs', text: 'Create subscriptions', iconProps: { iconName: 'RecurringEvent' }, disabled: dirty, onClick: () => run(() => engine.createSubscriptionsFromQuote(draft).then(r => `${r.length} subscription(s) created.`)) });
        }
      }
    }
    if (entity === 'invoices') {
      cmds.push({ key: 'print', text: 'Print / PDF', iconProps: { iconName: 'Print' }, onClick: () => printQuote(draft, engine) });
      if (writable) {
        const open_ = draft.invoiceStatus !== 'Paid' && draft.invoiceStatus !== 'Void';
        if (/^https:\/\//i.test(draft.paymentLink || '')) { cmds.push({ key: 'paylink', text: 'Open payment link', iconProps: { iconName: 'Link' }, href: draft.paymentLink, target: '_blank' }); }
        if (open_ && engine.settings.payments.provider !== 'none' && !draft.paymentLink) {
          cmds.push({ key: 'reqlink', text: `Create ${engine.settings.payments.provider} link`, iconProps: { iconName: 'PaymentCard' }, disabled: dirty || draft.paymentLinkStatus === 'Requested',
            onClick: () => quickPatch({ paymentLinkStatus: 'Requested', paymentProvider: engine.settings.payments.provider }, 'Payment link requested. Power Automate adds it in a few seconds.') });
        }
        cmds.push({ key: 'emailinv', text: 'Email invoice', iconProps: { iconName: 'Mail' }, disabled: !open_, onClick: () => setDialog('emailInvoice') });
        cmds.push({ key: 'sent', text: 'Mark sent', iconProps: { iconName: 'Send' }, disabled: draft.invoiceStatus !== 'Draft', onClick: () => quickPatch({ invoiceStatus: 'Sent' }, 'Invoice marked as sent.') });
        cmds.push({ key: 'paid', text: 'Record payment', iconProps: { iconName: 'Money' }, disabled: !open_ || dirty, onClick: () => {
          const due = Math.max(0, (draft.grandTotal || 0) - (Number(draft.amountPaid) || 0));
          const v = window.prompt(`Amount received (${currency})`, String(due));
          const amount = Number(v);
          if (v === null || !(amount > 0)) { return; }
          run(() => engine.recordPayment(draft, amount, draft.paymentProvider || 'Bank transfer').then(x => { setDraft({ ...x }); return 'Payment recorded.'; }));
        } });
        cmds.push({ key: 'void', text: 'Void', iconProps: { iconName: 'Blocked' }, disabled: !open_, onClick: () => { if (window.confirm('Void this invoice?')) { quickPatch({ invoiceStatus: 'Void' }, 'Invoice voided.'); } } });
      }
    }
    if (entity === 'subscriptions' && writable) {
      if (engine.can('create', 'invoices')) { cmds.push({ key: 'inv', text: 'Invoice next period', iconProps: { iconName: 'Money' }, disabled: dirty, onClick: () => run(() => engine.createInvoiceFromSubscription(draft).then(inv => { open('invoices', inv); return 'Invoice created.'; })) }); }
      cmds.push({ key: 'renew', text: 'Renew now', iconProps: { iconName: 'Refresh' }, disabled: dirty || draft.subStatus === 'Cancelled', onClick: () => quickPatch({ subStatus: 'Active', renewalDate: addPeriod(draft.renewalDate, draft.billingCycle) }, 'Renewed for one more period.') });
      cmds.push({ key: 'cancel', text: 'Cancel', iconProps: { iconName: 'Cancel' }, disabled: draft.subStatus === 'Cancelled', onClick: () => { if (window.confirm('Cancel this subscription at its renewal date?')) { quickPatch({ subStatus: 'Cancelled', autoRenew: false }, 'Subscription cancelled.'); } } });
    }
    if (entity === 'cases' && writable) {
      if (draft.owner?.id !== engine.me?.id) { cmds.push({ key: 'take', text: 'Assign to me', iconProps: { iconName: 'Contact' }, onClick: () => quickPatch({ owner: engine.me, caseStatus: draft.caseStatus === 'New' ? 'In progress' : draft.caseStatus }, 'Case assigned to you.') }); }
      cmds.push({ key: 'resolve', text: 'Resolve', iconProps: { iconName: 'Completed' }, disabled: draft.caseStatus === 'Resolved' || draft.caseStatus === 'Closed', onClick: () => {
        if (!draft.resolution) { setTab('details'); setErrors({ resolution: 'Describe the resolution before resolving' }); return; }
        quickPatch({ caseStatus: 'Resolved' }, 'Case resolved.');
      } });
      cmds.push({ key: 'reply', text: 'Reply to customer', iconProps: { iconName: 'Reply' }, disabled: dirty, onClick: () => setDialog('emailCase') });
      cmds.push({ key: 'wait', text: 'Waiting on customer', iconProps: { iconName: 'Clock' }, disabled: draft.caseStatus === 'Waiting on customer' || draft.caseStatus === 'Resolved' || draft.caseStatus === 'Closed', onClick: () => quickPatch({ caseStatus: 'Waiting on customer' }, engine.settings.slaPauseWhenWaiting ? 'Waiting on customer: SLA clock paused.' : 'Waiting on customer.') });
      cmds.push({ key: 'esc', text: 'Escalate', iconProps: { iconName: 'Warning' }, disabled: draft.priority === 'Critical', onClick: () => quickPatch({ priority: 'Critical' }, 'Case escalated to Critical.') });
    }
    if ((entity === 'accounts' || entity === 'contacts') && engine.can('create', 'opportunities')) {
      const base = entity === 'accounts' ? { accountRef: draft.id } : { contactRef: draft.id, accountRef: draft.accountRef };
      cmds.push({ key: 'newopp', text: 'New opportunity', iconProps: { iconName: 'Money' }, onClick: () => open('opportunities', { ...base, title: `${draft.title} - new deal` }) });
      if (engine.can('create', 'cases')) { cmds.push({ key: 'newcase', text: 'New case', iconProps: { iconName: 'Headset' }, onClick: () => open('cases', { ...base }) }); }
      if (entity === 'accounts' && engine.can('create', 'contacts')) { cmds.push({ key: 'newcon', text: 'New contact', iconProps: { iconName: 'AddFriend' }, onClick: () => open('contacts', { accountRef: draft.id }) }); }
    }
    if (settings.whatsapp.enabled && engine.license?.canWrite && ['leads', 'contacts', 'opportunities', 'demos', 'quotes', 'cases', 'invoices', 'subscriptions'].indexOf(entity) >= 0 && waRecipient(engine, entity, draft).number) {
      cmds.push({ key: 'whatsapp', text: 'WhatsApp', iconProps: { iconName: 'Chat' }, onClick: () => setDialog('whatsapp') });
    }
  }

  // ------------------------------------------------------------------ header summary
  const header = (): React.ReactNode => {
    if (isNew) { return null; }
    const bits: React.ReactNode[] = [];
    if (def.boardField && draft[def.boardField]) { bits.push(<Pill key="b" text={draft[def.boardField]} />); }
    if (entity === 'leads') { bits.push(<Pill key="r" text={draft.rating} />, <span key="s">Score <ScoreBadge score={draft.leadScore} /></span>); }
    if (entity === 'opportunities') { bits.push(<b key="a">{fmtMoney(draft.amount || 0, currency)}</b>, <span key="p">{draft.probability ?? 0}%</span>, <span key="c">close {fmtDate(draft.expectedCloseDate)}</span>); }
    if (entity === 'invoices') {
      bits.push(<b key="n">{draft.invoiceNumber}</b>, <b key="t">{fmtMoney(quoteTotals(draft.lineItems || []).grandTotal, currency)}</b>, <span key="d">due {fmtDate(draft.dueDate)}</span>);
      if (draft.amountPaid) { bits.push(<span key="p">paid {fmtMoney(draft.amountPaid, currency)}</span>); }
    }
    if (entity === 'subscriptions') { bits.push(<b key="m">MRR {fmtMoney(draft.mrr || 0, currency)}</b>, <span key="a">ARR {fmtMoney(draft.arr || 0, currency)}</span>, <span key="r">renews {fmtDate(draft.renewalDate)}</span>); }
    if (entity === 'quotes' && draft.signatureStatus) {
      bits.push(<Pill key="sig" text={draft.signatureStatus === 'Signed' ? `Signed by ${draft.signatureName}` : `E-signature: ${draft.signatureStatus}`} />);
    }
    if (entity === 'quotes') { bits.push(<b key="q">{draft.quoteNumber}</b>, <b key="t">{fmtMoney(quoteTotals(draft.lineItems || []).grandTotal, currency)}</b>); }
    if (entity === 'cases') { bits.push(<b key="n">{draft.caseNumber}</b>, <Pill key="pr" text={draft.priority} />); if (slaBreached(draft)) { bits.push(<Pill key="sla" text="Breached" />); } if (draft.slaPausedAt) { bits.push(<span key="pz" className="crm-muted">SLA paused</span>); } if (firstResponseLate(draft)) { bits.push(<Pill key="fr" text="First response late" />); } }
    if (entity === 'accounts') {
      const opps = (data.opportunities || []).filter(o => o.accountRef === draft.id);
      bits.push(<span key="w">Won {fmtMoney(sum(opps.filter(o => o.stage === WON), o => o.amount), currency)}</span>,
        <span key="o">Open pipeline {fmtMoney(sum(opps.filter(o => engine.openStages().indexOf(o.stage) >= 0), o => o.amount), currency)}</span>);
    }
    if (entity === 'campaigns') {
      const leads = (data.leads || []).filter(l => l.campaignRef === draft.id);
      const won = sum((data.opportunities || []).filter(o => o.campaignRef === draft.id && o.stage === WON), o => o.amount);
      const cost = Number(draft.actualCost || draft.budget || 0);
      bits.push(<span key="l">{leads.length} leads</span>, <span key="r">Won {fmtMoney(won, currency)}</span>,
        <span key="roi">ROI {cost ? `${Math.round(((won - cost) / cost) * 100)}%` : '—'}</span>);
    }
    if (draft.owner) { bits.push(<span key="ow" className="crm-muted">{draft.owner.title}</span>); }
    return <Stack horizontal wrap verticalAlign="center" tokens={{ childrenGap: 10 }} className="crm-panel-summary">{bits}</Stack>;
  };

  // ------------------------------------------------------------------ details
  const sections: Record<string, IFieldDef[]> = {};
  def.fields.filter(f => !f.hidden && f.type !== 'json' && !(isNew && f.readOnly && f.key !== 'title')).forEach(f => {
    if ((entity === 'leads' || entity === 'contacts') && f.key === 'title') { return; }
    if (entity === 'opportunities' && f.key === 'lostReason' && draft.stage !== LOST && !draft.lostReason) { return; }
    if (entity === 'opportunities' && f.key === 'actualCloseDate' && !draft.actualCloseDate) { return; }
    const s = f.section || 'General';
    (sections[s] = sections[s] || []).push(f);
  });

  const related = (def.related || []).filter(r => engine.hasFeature(r.entity));
  React.useEffect(() => {
    if (!isNew) { engine.ensure([...related.map(r => r.entity), 'activities']).catch(() => undefined); }
  }, [entity, record.id]);

  return (
    <Panel isOpen type={isNarrow ? PanelType.smallFluid : PanelType.medium} onDismiss={close} isLightDismiss={false} closeButtonAriaLabel="Close"
      isFooterAtBottom
      onRenderHeader={() => (
        <div className="crm-panel-header">
          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
            {canGoBack && <IconButton iconProps={{ iconName: 'Back' }} ariaLabel="Back" title="Back" onClick={close} />}
            <Text variant="xLarge" styles={{ root: { fontWeight: 600, wordBreak: 'break-word' } }}>
              {isNew ? `New ${def.singular.toLowerCase()}` : (draft.title || draft.quoteNumber || `${def.singular} #${draft.id}`)}
            </Text>
            {busy && <Spinner size={SpinnerSize.small} />}
          </Stack>
          {header()}
        </div>
      )}
      onRenderFooterContent={() => (
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          {canEdit && <PrimaryButton text={isNew ? 'Create' : 'Save'} disabled={busy || (!dirty && !isNew)} onClick={() => save(false)} />}
          {canEdit && !isNarrow && <DefaultButton text={isNew ? 'Create & close' : 'Save & close'} disabled={busy || (!dirty && !isNew)} onClick={() => save(true)} />}
          <DefaultButton text={canEdit ? 'Cancel' : 'Close'} onClick={close} />
          {!isNew && engine.can('delete', entity) && (isNarrow
            ? <IconButton iconProps={{ iconName: 'Delete' }} title="Delete" ariaLabel="Delete" style={{ marginLeft: 'auto' }} onClick={() => setConfirmDelete(true)} />
            : <DefaultButton text="Delete" iconProps={{ iconName: 'Delete' }} style={{ marginLeft: 'auto' }} onClick={() => setConfirmDelete(true)} />)}
        </Stack>
      )}
    >
      {cmds.length > 0 && <CommandBar items={cmds} styles={{ root: { padding: 0, height: 40 } }} overflowButtonProps={{ ariaLabel: 'More actions' }} />}
      {!canEdit && <MessageBar>{engine.license.canWrite ? 'You have read-only access to this record.' : engine.license.message}</MessageBar>}
      {error && <MessageBar messageBarType={MessageBarType.error} isMultiline onDismiss={() => setError(undefined)}>{error}</MessageBar>}
      {entity === 'quotes' && !isNew && engine.quoteApproval(draft).required && draft.approvalStatus !== 'Approved' && (
        <MessageBar messageBarType={draft.approvalStatus === 'Rejected' ? MessageBarType.error : MessageBarType.warning} isMultiline>
          {draft.approvalStatus === 'Pending'
            ? 'Waiting for manager approval. You will be notified when it is decided.'
            : draft.approvalStatus === 'Rejected'
              ? `Approval rejected${draft.approvalNote ? `: ${draft.approvalNote}` : ''}. Adjust the discount or request approval again.`
              : `Discount ${engine.quoteApproval(draft).discountPct}% is above the ${settings.quote.approvalDiscountPct}% limit. Request manager approval before sending.`}
        </MessageBar>
      )}
      {duplicates.length > 0 && (
        <MessageBar messageBarType={MessageBarType.warning} isMultiline>
          Possible duplicate{duplicates.length > 1 ? 's' : ''}:{' '}
          {duplicates.map((d, i) => <React.Fragment key={d.id}>{i > 0 && ', '}<Link onClick={() => open(entity, d)}>{d.title || d.email}</Link></React.Fragment>)}
        </MessageBar>
      )}

      <Pivot selectedKey={tab} onLinkClick={i => { const k = String(i?.props.itemKey); setTab(k); if (k === 'history') { loadHistory(); } }} overflowBehavior="menu">
        <PivotItem itemKey="details" headerText="Details" itemIcon="Edit">
          {!isNew && INSIGHT_ENTITIES.indexOf(entity) >= 0 && <SmartSummary entity={entity} record={draft} onDraft={() => setDialog('draft')} />}
          <div className="crm-form">
            {Object.keys(sections).map(s => (
              <fieldset key={s} className="crm-section">
                <legend>{s}</legend>
                <div className="crm-form-grid">
                  {sections[s].map(f => (
                    <div key={f.key} className={f.type === 'note' || f.type === 'tags' || f.type === 'user' || f.type === 'datetime' ? 'crm-span' : undefined}>
                      <FieldEditor fd={f} value={draft[f.key]} disabled={!canEdit} error={errors[f.key]} onChange={v => {
                        set(f.key, v);
                        if (entity === 'opportunities' && f.key === 'stage') {
                          const st = engine.stage(v);
                          setDraft(d => ({ ...d, stage: v, probability: st.probability, forecastCategory: st.forecast }));
                        }
                      }} />
                    </div>
                  ))}
                </div>
              </fieldset>
            ))}
            {!isNew && (
              <Text variant="small" className="crm-muted" block style={{ marginTop: 12 }}>
                Created {fmtDate(draft.created, true)} by {draft.author?.title || '—'} · Modified {fmtDate(draft.modified, true)} by {draft.editor?.title || '—'}
              </Text>
            )}
          </div>
        </PivotItem>
        {(entity === 'quotes' || entity === 'invoices') && (
          <PivotItem itemKey="lines" headerText="Line items" itemIcon="BulletedList" itemCount={(draft.lineItems || []).length}>
            <div style={{ paddingTop: 12 }}>
              <QuoteLines items={(draft.lineItems || []) as ILineItem[]} disabled={!canEdit} onChange={items => set('lineItems', items)} />
            </div>
          </PivotItem>
        )}
        {!isNew && related.map(rel => {
          const rows = (data[rel.entity] || []).filter(r => r[rel.field] === draft.id);
          return (
            <PivotItem key={rel.label} itemKey={`rel-${rel.label}`} headerText={rel.label} itemCount={data[rel.entity] ? rows.length : undefined} onRenderItemLink={undefined}>
              <RelatedList entity={rel.entity} rows={rows} loaded={!!data[rel.entity]} onNew={engine.can('create', rel.entity) ? () => open(rel.entity, {
                [rel.field]: draft.id, ...(entity === 'contacts' && rel.entity !== 'accounts' ? { accountRef: draft.accountRef } : {})
              }) : undefined} />
            </PivotItem>
          );
        })}
        {!isNew && entity !== 'activities' && (
          <PivotItem itemKey="activity" headerText="Activity" itemIcon="History">
            <div style={{ paddingTop: 12 }}><ActivityTimeline entity={entity} id={draft.id} /></div>
          </PivotItem>
        )}
        {!isNew && engine.hasFeature('audit') && settings.auditEnabled && (
          <PivotItem itemKey="history" headerText="Audit" itemIcon="ComplianceAudit">
            <div style={{ paddingTop: 12 }}>
              {!history && <Spinner label="Loading history..." />}
              {history && history.length === 0 && <Text className="crm-muted">No recorded changes.</Text>}
              {history && history.map(h => (
                <div key={h.id} className="crm-timeline-item">
                  <div style={{ flex: 1 }}>
                    <Text styles={{ root: { fontWeight: 600 } }}>{h.title}</Text>
                    <Text variant="small" className="crm-muted"> · {h.author?.title} · {fmtDate(h.created, true)}</Text>
                    {h.changes && Object.keys(h.changes).map(k => (
                      <Text key={k} variant="small" block>{k}: <s className="crm-muted">{String(h.changes[k][0] || '—')}</s> → <b>{String(h.changes[k][1] || '—')}</b></Text>
                    ))}
                  </div>
                </div>
              ))}
            </div>
          </PivotItem>
        )}
      </Pivot>

      <Dialog hidden={!confirmDelete} onDismiss={() => setConfirmDelete(false)}
        dialogContentProps={{ type: DialogType.normal, title: `Delete ${def.singular.toLowerCase()}?`, subText: 'The record moves to the site recycle bin and can be restored by a site owner.' }}>
        <DialogFooter>
          <PrimaryButton text="Delete" onClick={() => { setConfirmDelete(false); remove(); }} />
          <DefaultButton text="Cancel" onClick={() => setConfirmDelete(false)} />
        </DialogFooter>
      </Dialog>

      {dialog === 'lost' && (
        <Dialog hidden={false} onDismiss={() => setDialog(undefined)} dialogContentProps={{ type: DialogType.normal, title: 'Why was this deal lost?' }}>
          <Dropdown label="Lost reason" required selectedKey={draft.lostReason} options={choicesFor(def.fields.find(f => f.key === 'lostReason')!, settings).map(c => ({ key: c, text: c }))}
            onChange={(_, o) => setDraft({ ...draft, lostReason: String(o?.key) })} />
          <TextField label="Competitor (optional)" value={draft.competitors || ''} onChange={(_, v) => setDraft({ ...draft, competitors: v })} />
          <DialogFooter>
            <PrimaryButton text="Mark as lost" disabled={!draft.lostReason} onClick={() => { setDialog(undefined); save(false, { ...draft, stage: LOST }); }} />
            <DefaultButton text="Cancel" onClick={() => setDialog(undefined)} />
          </DialogFooter>
        </Dialog>
      )}
      {dialog === 'convert' && <ConvertLeadDialog lead={draft} onClose={() => setDialog(undefined)} />}
      {dialog === 'email' && <EmailDialog entity={entity} record={draft} onClose={() => setDialog(undefined)} />}
      {dialog === 'meeting' && <MeetingDialog demo={draft} onClose={() => setDialog(undefined)} />}
      {dialog === 'emailQuote' && (
        <EmailDialog entity={entity} record={draft} to={engine.find('contacts', draft.contactRef)?.email || ''}
          subject={`Quote ${draft.quoteNumber} - ${draft.title}`}
          body={`Hello,\n\nPlease find our quote ${draft.quoteNumber} below.\n\nTotal: ${fmtMoney(quoteTotals(draft.lineItems || []).grandTotal, currency)}\nValid until: ${fmtDate(draft.validUntil)}\n\n${(draft.lineItems || []).map((i: ILineItem) => `• ${i.name} × ${i.quantity}`).join('\n')}\n\nUse "Print / PDF" in the CRM to attach a PDF copy if needed.\n\nBest regards,\n${engine.me?.title || ''}`}
          onClose={() => { setDialog(undefined); if (draft.quoteStatus === 'Draft') { quickPatch({ quoteStatus: 'Sent' }, 'Quote marked as sent.'); } }} />
      )}
      {dialog === 'emailCase' && (
        <EmailDialog entity={entity} record={draft} to={engine.find('contacts', draft.contactRef)?.email || draft.requesterEmail || ''}
          subject={`[${draft.caseNumber}] ${String(draft.title || '').replace(/^\s*\[[^\]]+\]\s*/, '')}`}
          body={`Hello,\n\n\n\nKind regards,\n${engine.me?.title || ''}\n${engine.settings.company.name}\n\nPlease keep [${draft.caseNumber}] in the subject when you reply.`}
          onClose={() => { setDialog(undefined); const c = engine.find('cases', draft.id); if (c) { setDraft({ ...c }); } }} />
      )}
      {dialog === 'whatsapp' && <WhatsAppDialog entity={entity} record={draft} onClose={() => setDialog(undefined)} />}
      {dialog === 'draft' && (() => {
        const d = draftFollowUp(engine, entity, draft);
        return <EmailDialog entity={entity} record={draft} to={d.to} subject={d.subject} body={d.body} onClose={() => setDialog(undefined)} />;
      })()}
      {dialog === 'emailSign' && (
        <EmailDialog entity={entity} record={draft} to={draft.signatureEmail || ''}
          subject={`Please review and sign quote ${draft.quoteNumber} - ${engine.settings.company.name}`}
          body={`Hello,\n\nPlease review our quote ${draft.quoteNumber} (${fmtMoney(quoteTotals(draft.lineItems || []).grandTotal, currency)}, valid until ${fmtDate(draft.validUntil)}) and accept it online:\n\n${signLink}\n\nThe link is personal to you and works once.\n\nBest regards,\n${engine.me?.title || ''}`}
          onClose={() => setDialog(undefined)} />
      )}
      {dialog === 'emailInvoice' && (
        <EmailDialog entity={entity} record={draft} to={engine.find('contacts', draft.contactRef)?.email || ''}
          subject={`Invoice ${draft.invoiceNumber} - ${engine.settings.company.name}`}
          body={`Hello,\n\nPlease find invoice ${draft.invoiceNumber} for ${fmtMoney(Math.max(0, (draft.grandTotal || 0) - (Number(draft.amountPaid) || 0)), currency)}, due ${fmtDate(draft.dueDate)}.\n\n${(draft.lineItems || []).map((i: ILineItem) => `• ${i.name} × ${i.quantity}`).join('\n')}\n${/^https:\/\//i.test(draft.paymentLink || '') ? `\nPay online: ${draft.paymentLink}\n` : ''}${engine.settings.invoice.bankDetails ? `\nBank transfer:\n${engine.settings.invoice.bankDetails}\n` : ''}\nBest regards,\n${engine.me?.title || ''}`}
          onClose={() => { setDialog(undefined); if (draft.invoiceStatus === 'Draft') { quickPatch({ invoiceStatus: 'Sent' }, 'Invoice marked as sent.'); } }} />
      )}
    </Panel>
  );
};

