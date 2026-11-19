// Record insights: a short summary, risks and next best actions computed from CRM data,
// a follow-up email draft, and a ready-made prompt for Microsoft 365 Copilot.
// Everything here runs locally in the browser; no record data leaves the tenant.
import { EntityKey, IRecord, LOST, WON } from '../models/model';
import { Engine } from './Engine';
import { fmtDate, fmtMoney, slaBreached, firstResponseLate } from '../components/utils';

export interface IBrief {
  headline: string;
  facts: string[];
  risks: string[];
  next: string[];
}

const DAY = 86400000;
const daysAgo = (iso?: string): number | undefined => (iso ? Math.floor((Date.now() - new Date(iso).getTime()) / DAY) : undefined);
const daysTo = (iso?: string): number | undefined => (iso ? Math.ceil((new Date(iso).getTime() - Date.now()) / DAY) : undefined);
const plural = (n: number, w: string): string => `${n} ${w}${n === 1 ? '' : 's'}`;

export const INSIGHT_ENTITIES: EntityKey[] = ['leads', 'contacts', 'accounts', 'opportunities', 'quotes', 'cases', 'subscriptions', 'invoices'];

/** Activities linked to the record (and, for accounts, to its contacts and deals). */
export function relatedActivities(engine: Engine, key: EntityKey, r: IRecord): IRecord[] {
  const acts = engine.data.activities || [];
  const pairs: { type: string; id: number }[] = [{ type: engine.def(key).activityName, id: r.id! }];
  if (key === 'accounts') {
    (engine.data.contacts || []).filter(c => c.accountRef === r.id).forEach(c => pairs.push({ type: engine.def('contacts').activityName, id: c.id! }));
    (engine.data.opportunities || []).filter(o => o.accountRef === r.id).forEach(o => pairs.push({ type: engine.def('opportunities').activityName, id: o.id! }));
  }
  return acts.filter(a => pairs.some(p => p.type === a.relatedType && p.id === a.relatedId))
    .sort((a, b) => String(b.dueDate || b.created || '').localeCompare(String(a.dueDate || a.created || '')));
}

export function recordBrief(engine: Engine, key: EntityKey, r: IRecord): IBrief {
  const cur = engine.settings.currency;
  const money = (n: number): string => fmtMoney(n || 0, cur);
  const acts = relatedActivities(engine, key, r);
  const done = acts.filter(a => a.completed);
  const openTasks = acts.filter(a => a.activityType === 'Task' && !a.completed);
  const overdueTasks = openTasks.filter(a => a.dueDate && new Date(a.dueDate).getTime() < Date.now());
  const last = done[0];
  const lastDays = daysAgo(last?.dueDate || last?.created);
  const facts: string[] = [];
  const risks: string[] = [];
  const next: string[] = [];
  let headline = r.title || '';

  if (last) { facts.push(`Last touch ${lastDays === 0 ? 'today' : `${lastDays} day(s) ago`}: ${last.title}`); } else { facts.push('No logged activity yet.'); }
  const emails = done.filter(a => a.activityType === 'Email').length;
  const meetings = done.filter(a => a.activityType === 'Meeting').length;
  const chats = done.filter(a => a.activityType === 'WhatsApp').length;
  if (emails + meetings + chats) { facts.push(`${plural(emails, 'email')}, ${plural(meetings, 'meeting')}${chats ? `, ${plural(chats, 'WhatsApp message')}` : ''} logged`); }
  if (openTasks.length) { facts.push(`${plural(openTasks.length, 'open task')}${overdueTasks.length ? ` (${overdueTasks.length} overdue)` : ''}`); }
  if (overdueTasks.length) { risks.push(`${plural(overdueTasks.length, 'task')} overdue: ${overdueTasks.slice(0, 2).map(t => t.title).join('; ')}`); }

  switch (key) {
    case 'leads': {
      headline = `${r.status || 'New'} ${r.rating ? `${String(r.rating).toLowerCase()} ` : ''}lead from ${r.leadSource || 'unknown source'}${r.company ? ` at ${r.company}` : ''}, score ${r.leadScore ?? 0}/100`;
      const age = daysAgo(r.created);
      if (age !== undefined) { facts.push(`Created ${age} day(s) ago`); }
      if (r.status === 'New' && (age || 0) >= 2) { risks.push('Still “New” after 2+ days: speed to lead matters most in the first hours.'); }
      if (!r.owner) { risks.push('No owner assigned.'); next.push('Assign an owner or use round-robin.'); }
      if (r.nextFollowUp && daysTo(r.nextFollowUp)! < 0) { risks.push(`Follow-up was due ${fmtDate(r.nextFollowUp)}.`); }
      if (r.status === 'New') { next.push('Call or email within 24 hours, then set status to Contacted.'); }
      if ((r.leadScore || 0) >= 60 && r.status !== 'Converted') { next.push('High score: offer a demo, then convert to an opportunity.'); }
      if (!r.nextFollowUp && r.status !== 'Converted' && r.status !== 'Unqualified') { next.push('Set a next follow-up date.'); }
      break;
    }
    case 'opportunities': {
      const st = r.stage || '';
      headline = `${fmtMoney(r.amount || 0, r.currency || cur)} ${st} deal${r.dealType && r.dealType !== 'New business' ? ` (${String(r.dealType).toLowerCase()})` : ''}, ${r.probability ?? 0}% · close ${fmtDate(r.expectedCloseDate) || 'not set'}`;
      const quotes = (engine.data.quotes || []).filter(q => q.opportunityRef === r.id);
      if (quotes.length) { facts.push(`Quotes: ${quotes.map(q => `${q.quoteNumber} ${q.quoteStatus || ''}${q.signatureStatus ? ` (${q.signatureStatus})` : ''}`).join(', ')}`); }
      if (st !== WON && st !== LOST) {
        const toClose = daysTo(r.expectedCloseDate);
        if (toClose !== undefined && toClose < 0) { risks.push(`Close date passed ${-toClose} day(s) ago: update the date or the stage.`); }
        if (lastDays !== undefined && lastDays > 14) { risks.push(`No activity for ${lastDays} days: the deal may be going cold.`); }
        if (!r.contactRef) { risks.push('No contact linked: find the decision maker.'); }
        if (quotes.some(q => q.approvalStatus === 'Pending')) { risks.push('A quote is waiting for discount approval.'); }
        if (!quotes.length && (r.probability || 0) >= 50) { next.push('Send a quote: the deal is at 50%+ without one.'); }
        if (quotes.some(q => q.signatureStatus === 'Sent')) { next.push('Quote sent for signature but not opened yet: send a reminder.'); }
        if (quotes.some(q => q.signatureStatus === 'Viewed')) { next.push('The customer viewed the quote: call to answer questions and ask for the signature.'); }
        if (!openTasks.length) { next.push('Plan the next step as a task with a due date.'); }
      }
      break;
    }
    case 'accounts': {
      const opps = (engine.data.opportunities || []).filter(o => o.accountRef === r.id);
      const open = opps.filter(o => engine.openStages().indexOf(o.stage) >= 0);
      const won = opps.filter(o => o.stage === WON);
      const subs = (engine.data.subscriptions || []).filter(s => s.accountRef === r.id && (s.subStatus === 'Active' || s.subStatus === 'Pending renewal'));
      const cases = (engine.data.cases || []).filter(c => c.accountRef === r.id && c.caseStatus !== 'Resolved' && c.caseStatus !== 'Closed');
      const unpaid = (engine.data.invoices || []).filter(i => i.accountRef === r.id && (i.invoiceStatus === 'Sent' || i.invoiceStatus === 'Overdue'));
      const mrr = subs.reduce((t, s) => t + (s.mrr || 0), 0);
      headline = `${r.accountType || 'Account'}${r.industry ? ` in ${r.industry}` : ''}: won ${money(won.reduce((t, o) => t + engine.baseAmount(o), 0))}, open pipeline ${money(open.reduce((t, o) => t + engine.baseAmount(o), 0))}${mrr ? `, MRR ${money(mrr)}` : ''}`;
      facts.push(`${plural((engine.data.contacts || []).filter(c => c.accountRef === r.id).length, 'contact')}, ${plural(open.length, 'open deal')}, ${plural(cases.length, 'open case')}`);
      subs.filter(s => (daysTo(s.renewalDate) ?? 999) <= engine.settings.renewals.leadDays).forEach(s => next.push(`Renewal of “${s.title}” due ${fmtDate(s.renewalDate)}.`));
      if (cases.some(c => slaBreached(c))) { risks.push('A support case has breached its SLA.'); }
      if (unpaid.some(i => i.invoiceStatus === 'Overdue')) { risks.push(`Overdue invoice(s): ${unpaid.filter(i => i.invoiceStatus === 'Overdue').map(i => i.invoiceNumber).join(', ')}`); }
      if (r.accountType === 'Customer' && !open.length && !subs.length) { next.push('No open deals: look for an upsell or a renewal.'); }
      if (lastDays === undefined || lastDays > 30) { risks.push('No contact with this account for 30+ days.'); }
      break;
    }
    case 'contacts': {
      headline = `${r.jobTitle || 'Contact'}${r.accountRef ? ` at ${engine.display('accounts', r.accountRef)}` : ''}`;
      const opps = (engine.data.opportunities || []).filter(o => o.contactRef === r.id && engine.openStages().indexOf(o.stage) >= 0);
      if (opps.length) { facts.push(`Involved in ${plural(opps.length, 'open deal')}: ${opps.slice(0, 3).map(o => o.title).join(', ')}`); }
      if (r.marketingConsent === false) { risks.push('Do not use for marketing email (no consent).'); }
      if (lastDays === undefined || lastDays > 30) { next.push('Reconnect: no conversation in the last 30 days.'); }
      break;
    }
    case 'quotes': {
      const appr = engine.quoteApproval(r);
      headline = `${r.quoteNumber} · ${money(r.grandTotal)} · ${r.quoteStatus || 'Draft'}${r.signatureStatus ? ` · e-signature ${r.signatureStatus.toLowerCase()}` : ''}`;
      facts.push(`Discount ${appr.discountPct}%${appr.required ? ` (approval ${r.approvalStatus || 'needed'})` : ''} · valid until ${fmtDate(r.validUntil) || '—'}`);
      if (r.validUntil && daysTo(r.validUntil)! < 0 && r.quoteStatus !== 'Accepted') { risks.push('The quote has expired: extend it or send a new one.'); }
      if (appr.required && !appr.approved) { next.push('Request discount approval before sending.'); }
      if (r.signatureStatus === 'Viewed') { next.push('Viewed but not signed: follow up personally.'); }
      if (r.quoteStatus === 'Accepted' && !(engine.data.invoices || []).some(i => i.quoteRef === r.id)) { next.push('Accepted: create the invoice.'); }
      break;
    }
    case 'cases': {
      headline = `${r.caseNumber || 'Case'} · ${r.priority || 'Normal'} · ${r.caseStatus || 'New'}${r.origin ? ` · via ${r.origin}` : ''}`;
      if (r.slaDue) { facts.push(`${r.slaPausedAt ? 'SLA paused' : `SLA due ${fmtDate(r.slaDue, true)}`}`); }
      if (slaBreached(r)) { risks.push('SLA breached: escalate or update the customer now.'); }
      if (firstResponseLate(r)) { risks.push('First response is late.'); }
      if (r.caseStatus === 'New') { next.push('Reply to the customer and set the case to In progress.'); }
      if (r.caseStatus === 'Waiting on customer' && (daysAgo(r.slaPausedAt) || 0) >= 3) { next.push('Waiting 3+ days: send a reminder or resolve the case.'); }
      const others = (engine.data.cases || []).filter(c => c.id !== r.id && c.accountRef && c.accountRef === r.accountRef && c.caseStatus !== 'Closed');
      if (others.length) { facts.push(`${plural(others.length, 'other case')} for this account`); }
      break;
    }
    case 'subscriptions': {
      const toRenew = daysTo(r.renewalDate);
      headline = `${r.subStatus || 'Active'} · ${r.quantity || 0} × ${money(r.unitPrice)} ${String(r.billingCycle || '').toLowerCase()} · MRR ${money(r.mrr)}`;
      if (toRenew !== undefined) { facts.push(toRenew >= 0 ? `Renews in ${toRenew} day(s) (${fmtDate(r.renewalDate)})` : `Renewal date passed ${-toRenew} day(s) ago`); }
      if (!r.autoRenew) { risks.push('Auto-renew is off: the subscription expires at the renewal date.'); }
      if (toRenew !== undefined && toRenew <= 30 && r.subStatus === 'Pending renewal') { next.push('Renewal within 30 days: confirm seats and send the renewal quote.'); }
      break;
    }
    case 'invoices': {
      const bal = Math.max(0, (r.grandTotal || 0) - (r.amountPaid || 0));
      headline = `${r.invoiceNumber} · ${money(r.grandTotal)} · ${r.invoiceStatus || 'Draft'}${bal && r.invoiceStatus !== 'Draft' ? ` · balance ${money(bal)}` : ''}`;
      if (r.invoiceStatus === 'Overdue') { risks.push(`Overdue since ${fmtDate(r.dueDate)}.`); next.push('Send a payment reminder with the payment link.'); }
      if (r.invoiceStatus === 'Draft') { next.push('Email the invoice to the billing contact.'); }
      if (!r.paymentLink && engine.settings.payments.provider !== 'none' && r.invoiceStatus !== 'Paid') { next.push(`Create a ${engine.settings.payments.provider} payment link.`); }
      break;
    }
  }
  return { headline, facts, risks, next: next.slice(0, 4) };
}

/** Follow-up email draft based on the record's situation. */
export function draftFollowUp(engine: Engine, key: EntityKey, r: IRecord): { to: string; subject: string; body: string } {
  const contact = r.contactRef ? engine.find('contacts', r.contactRef) : key === 'contacts' ? r : undefined;
  const to = String(r.email || contact?.email || r.requesterEmail || r.signatureEmail || '');
  const first = String(r.firstName || contact?.firstName || String(r.contactName || '').split(' ')[0] || '').trim();
  const me = engine.me?.title || '';
  const company = engine.settings.company.name || '';
  const hello = `Hello${first ? ` ${first}` : ''},`;
  const sign = `\n\nBest regards,\n${me}${company ? `\n${company}` : ''}`;
  const acts = relatedActivities(engine, key, r).filter(a => a.completed);
  const last = acts[0];
  const recap = last ? `Thank you for your time${last.activityType === 'Meeting' ? ' in our meeting' : ''}${last.title ? ` (${String(last.title).replace(/^(Received|Sent|Meeting):\s*/i, '')})` : ''}.` : 'I hope you are well.';
  switch (key) {
    case 'leads':
      return { to, subject: `${company ? `${company}: ` : ''}next steps`, body: `${hello}\n\nThank you for your interest${r.offer ? ` in ${r.offer}` : ''}. I would be happy to show you how it works for ${r.company || 'your team'} in a short demo.\n\nWould one of these times suit you this week? Just reply with what works best.${sign}` };
    case 'opportunities': {
      const q = (engine.data.quotes || []).filter(x => x.opportunityRef === r.id).sort((a, b) => String(b.modified).localeCompare(String(a.modified)))[0];
      return { to, subject: `${r.title}: next steps`, body: `${hello}\n\n${recap}\n\n${q ? `Our quote ${q.quoteNumber} is ready${q.validUntil ? ` and valid until ${fmtDate(q.validUntil)}` : ''}. Do you have any questions before you decide?` : 'As discussed, I will prepare a proposal. Could you confirm the number of users and your timeline?'}\n\nI suggest a short call to agree on the next steps.${sign}` };
    }
    case 'quotes':
      return { to, subject: `Quote ${r.quoteNumber}`, body: `${hello}\n\nI wanted to check whether you had a chance to review quote ${r.quoteNumber}${r.validUntil ? ` (valid until ${fmtDate(r.validUntil)})` : ''}. I am happy to walk you through it or adjust it.${sign}` };
    case 'cases':
      return { to, subject: `[${r.caseNumber}] ${String(r.title || '').replace(/^\s*\[[^\]]+\]\s*/, '')}`, body: `${hello}\n\nThank you for your patience. Here is an update on your request:\n\n\n\nPlease keep [${r.caseNumber}] in the subject when you reply.${sign}` };
    case 'invoices':
      return { to, subject: `Invoice ${r.invoiceNumber}`, body: `${hello}\n\nThis is a friendly reminder that invoice ${r.invoiceNumber} for ${fmtMoney(Math.max(0, (r.grandTotal || 0) - (r.amountPaid || 0)), engine.settings.currency)} was due on ${fmtDate(r.dueDate)}.${/^https:\/\//i.test(r.paymentLink || '') ? `\n\nYou can pay online here: ${r.paymentLink}` : ''}\n\nPlease let me know if you have any questions.${sign}` };
    case 'subscriptions':
      return { to, subject: `Your ${r.title} renewal`, body: `${hello}\n\nYour subscription renews on ${fmtDate(r.renewalDate)} (${r.quantity} seats). Would you like to keep the same number of seats, or adjust it for the coming term?${sign}` };
    default:
      return { to, subject: 'Checking in', body: `${hello}\n\n${recap}\n\nIs there anything I can help you with at the moment?${sign}` };
  }
}

/** A grounded prompt for Microsoft 365 Copilot (or any assistant the organisation approves). */
export function copilotPrompt(engine: Engine, key: EntityKey, r: IRecord): string {
  const d = engine.def(key);
  const b = recordBrief(engine, key, r);
  const fields = d.fields.filter(f => !f.hidden && f.type !== 'json' && r[f.key] !== undefined && r[f.key] !== null && r[f.key] !== '')
    .slice(0, 18).map(f => {
      const v = r[f.key];
      const text = v && typeof v === 'object' && 'title' in v ? v.title : f.type === 'lookup' ? engine.display(f.lookup as EntityKey, v) : Array.isArray(v) ? v.join(', ') : f.type.indexOf('date') === 0 ? fmtDate(v, f.type === 'datetime') : String(v);
      return `- ${f.label}: ${String(text).slice(0, 300)}`;
    });
  const acts = relatedActivities(engine, key, r).slice(0, 8).map(a => `- ${fmtDate(a.dueDate || a.created)} ${a.activityType}: ${a.title}${a.details ? ` — ${String(a.details).replace(/\s+/g, ' ').slice(0, 200)}` : ''}`);
  return [
    `You are helping me with ${/^[aeiou]/i.test(d.singular) ? 'an' : 'a'} ${d.singular.toLowerCase()} in our CRM. Summarise the situation in 3 bullets, list risks, and suggest the next best action. Then draft a short, friendly follow-up email I can send.`,
    '', `${d.singular}: ${r.title || r.quoteNumber || r.caseNumber || ''}`, b.headline, '', 'Fields:', ...fields,
    ...(acts.length ? ['', 'Recent activity:', ...acts] : []),
    ...(b.risks.length ? ['', 'Signals from the CRM:', ...b.risks.map(x => `- ${x}`)] : [])
  ].join('\n');
}
