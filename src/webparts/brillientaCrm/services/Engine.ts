/* eslint-disable @typescript-eslint/no-explicit-any */
import { DataService, Role } from './DataService';
import { GraphService, WebhookService } from './Integrations';
import { CrmData, EntityKey, IEntityDef, IRecord, IUserRef, LOST, WON } from '../models/model';
import { ICrmSettings } from '../models/settings';
import { ILicenseInfo } from '../models/license';
import { addBusinessHours, businessHoursBetween } from './BusinessHours';
import { waNumber } from './WhatsApp';
import { fmtMoney } from '../components/utils';

export type Action = 'read' | 'create' | 'edit' | 'delete' | 'import' | 'export' | 'bulk' | 'admin' | 'reports' | 'forecast';

export interface ILineItem {
  id: string;
  productId?: number;
  name: string;
  description?: string;
  quantity: number;
  unitPrice: number;
  discountPct: number;
  taxPct: number;
  billing?: string;
}

export interface IQuoteTotals {
  subtotal: number;
  discountTotal: number;
  taxTotal: number;
  grandTotal: number;
}

export interface IConvertOptions {
  accountId?: number;
  createAccount: boolean;
  contactId?: number;
  createContact: boolean;
  createOpportunity: boolean;
  opportunityName: string;
  amount: number;
  expectedCloseDate?: string;
}

const SALES_EDIT: EntityKey[] = ['leads', 'accounts', 'contacts', 'opportunities', 'quotes', 'demos', 'activities', 'cases', 'subscriptions', 'invoices'];
const SUPPORT_EDIT: EntityKey[] = ['cases', 'contacts', 'accounts', 'activities'];
const ENTERPRISE_FEATURES = ['cases', 'campaigns', 'forecast', 'webhooks', 'customFields', 'audit', 'security'];

export function quoteTotals(items: ILineItem[] = []): IQuoteTotals {
  let subtotal = 0;
  let discountTotal = 0;
  let taxTotal = 0;
  items.forEach(i => {
    const gross = (Number(i.quantity) || 0) * (Number(i.unitPrice) || 0);
    const disc = gross * ((Number(i.discountPct) || 0) / 100);
    const net = gross - disc;
    subtotal += gross;
    discountTotal += disc;
    taxTotal += net * ((Number(i.taxPct) || 0) / 100);
  });
  const r = (n: number): number => Math.round(n * 100) / 100;
  return { subtotal: r(subtotal), discountTotal: r(discountTotal), taxTotal: r(taxTotal), grandTotal: r(subtotal - discountTotal + taxTotal) };
}

const DAY = 86400000;
const dayKey = (iso?: string): string => (iso ? new Date(iso).toISOString().slice(0, 10) : '');

/** Adds whole billing periods to a date (Monthly → months, Annual/other → years). */
export function addPeriod(iso: string | undefined, cycle: string, n: number = 1): string {
  const d = iso ? new Date(iso) : new Date();
  if (cycle === 'Monthly') { d.setMonth(d.getMonth() + n); } else { d.setFullYear(d.getFullYear() + n); }
  return d.toISOString();
}

/** Monthly recurring revenue of a subscription. */
export function subscriptionMrr(s: IRecord): number {
  const period = (Number(s.quantity) || 0) * (Number(s.unitPrice) || 0);
  return Math.round((s.billingCycle === 'Monthly' ? period : period / 12) * 100) / 100;
}

export interface IMaintenanceResult { renewals: number; overdue: number; expired: number; rolled: number }

export class Engine {
  public data: CrmData = {};
  public me?: IUserRef;
  public role: Role = 'reader';
  public license!: ILicenseInfo;
  public readonly webhooks: WebhookService;
  private loading: Partial<Record<EntityKey, Promise<IRecord[]>>> = {};

  constructor(
    public readonly ds: DataService,
    public readonly graph: GraphService,
    tenantId: string,
    private readonly onData: (key: EntityKey, records: IRecord[]) => void
  ) {
    this.webhooks = new WebhookService(
      () => (this.hasFeature('webhooks') ? this.settings.webhooks : []),
      () => (this.hasFeature('webhooks') ? this.settings.integrations.teamsWebhookUrl : ''),
      tenantId
    );
  }

  public setSession(me: IUserRef | undefined, role: Role, license: ILicenseInfo): void {
    this.me = me;
    this.role = role;
    this.license = license;
  }

  public get settings(): ICrmSettings {
    return this.ds.settings;
  }

  public get defs(): IEntityDef[] {
    return this.ds.defs.filter(d => this.hasFeature(d.key));
  }

  public def(key: EntityKey): IEntityDef {
    return this.ds.def(key);
  }

  // ------------------------------------------------------------------ licensing & permissions

  public hasFeature(feature: string): boolean {
    if (ENTERPRISE_FEATURES.indexOf(feature) < 0) { return true; }
    return this.license?.features === 'enterprise';
  }

  public can(action: Action, key?: EntityKey): boolean {
    const r = this.role;
    if (key && !this.hasFeature(key)) { return false; }
    if (action === 'read' || action === 'export') { return true; }
    if (action === 'admin') { return r === 'admin'; }
    if (action === 'reports') { return true; }
    if (action === 'forecast') { return this.hasFeature('forecast') && (r === 'admin' || r === 'manager' || r === 'sales'); }
    if (!this.license?.canWrite) { return false; }
    if (r === 'admin' || r === 'manager') { return true; }
    if (r === 'reader') { return false; }
    if (action === 'delete' || action === 'import') { return false; }
    if (!key) { return action === 'bulk' && r === 'sales'; }
    if (r === 'sales') { return SALES_EDIT.indexOf(key) >= 0; }
    if (r === 'support') { return SUPPORT_EDIT.indexOf(key) >= 0; }
    return false;
  }

  /** Records the current user may see (row filtering for sales reps when enabled). */
  public visible(key: EntityKey): IRecord[] {
    const all = this.data[key] || [];
    if (!(this.settings.restrictRepsToOwnRecords && this.role === 'sales' && this.me)) { return all; }
    if (key === 'products' || key === 'campaigns') { return all; }
    const me = this.me.id;
    return all.filter(r => !r.owner || r.owner.id === me || r.author?.id === me);
  }

  // ------------------------------------------------------------------ loading

  public ensure(keys: EntityKey[]): Promise<void> {
    return Promise.all(keys.filter(k => this.hasFeature(k)).map(k => (this.data[k] ? Promise.resolve(this.data[k]!) : this.refresh(k)))).then(() => undefined);
  }

  public refresh(key: EntityKey): Promise<IRecord[]> {
    if (!this.loading[key]) {
      this.loading[key] = this.ds.getAll(key).then(rows => {
        this.data[key] = rows;
        this.onData(key, rows);
        delete this.loading[key];
        return rows;
      }, e => { delete this.loading[key]; throw e; });
    }
    return this.loading[key]!;
  }

  public find(key: EntityKey, id?: number): IRecord | undefined {
    return id ? (this.data[key] || []).find(r => r.id === id) : undefined;
  }

  public display(key: EntityKey, id?: number): string {
    const r = this.find(key, id);
    return r ? (r.title || r.quoteNumber || r.caseNumber || `#${r.id}`) : (id ? `#${id}` : '');
  }

  private upsertCache(key: EntityKey, rec: IRecord): void {
    const rows = [...(this.data[key] || [])];
    const i = rows.findIndex(r => r.id === rec.id);
    if (i >= 0) { rows[i] = rec; } else { rows.unshift(rec); }
    this.data[key] = rows;
    this.onData(key, rows);
  }

  private removeCache(key: EntityKey, id: number): void {
    const rows = (this.data[key] || []).filter(r => r.id !== id);
    this.data[key] = rows;
    this.onData(key, rows);
  }

  // ------------------------------------------------------------------ automation helpers

  public stage(name: string): { probability: number; forecast: string } {
    const s = this.settings.stages.find(x => x.name === name);
    return s ? s : { probability: name === WON ? 100 : 0, forecast: name === WON ? 'Closed' : 'Pipeline' };
  }

  public openStages(): string[] {
    return this.settings.stages.map(s => s.name).filter(n => n !== WON && n !== LOST);
  }

  public scoreLead(l: IRecord): number {
    const s = this.settings.scoring;
    let score = 0;
    score += s.rating[l.rating] || 0;
    score += s.sources[l.leadSource] || 0;
    if (l.phone) { score += s.hasPhone; }
    if (l.company) { score += s.hasCompany; }
    const jt = String(l.jobTitle || '').toLowerCase();
    let best = 0;
    s.titleKeywords.forEach(k => { if (k.keyword && jt.indexOf(k.keyword.toLowerCase()) >= 0) { best = Math.max(best, k.points); } });
    score += best;
    if (l.id) {
      const acts = (this.data.activities || []).filter(a => a.relatedType === 'Lead' && a.relatedId === l.id).length;
      score += Math.min(s.maxActivityPoints, acts * s.perActivity);
      if ((this.data.demos || []).some(d => d.leadRef === l.id)) { score += 15; }
    }
    return Math.max(0, Math.min(100, score));
  }

  public nextNumber(key: EntityKey, field: string, prefix: string): string {
    const year = new Date().getFullYear();
    const head = `${prefix}-${year}-`;
    let max = 0;
    (this.data[key] || []).forEach(r => {
      const v = String(r[field] || '');
      if (v.indexOf(head) === 0) { max = Math.max(max, Number(v.slice(head.length)) || 0); }
    });
    return `${head}${('000' + (max + 1)).slice(-Math.max(4, String(max + 1).length))}`;
  }

  /** Resolution SLA deadline, counted in business hours when they are switched on. */
  public slaDue(priority: string, fromIso?: string): string {
    const hours = this.settings.slaHours[priority] ?? 24;
    return addBusinessHours(fromIso ? new Date(fromIso) : new Date(), hours, this.settings.businessHours).toISOString();
  }

  public firstResponseDue(priority: string, fromIso?: string): string {
    const hours = this.settings.firstResponseHours[priority] ?? 8;
    return addBusinessHours(fromIso ? new Date(fromIso) : new Date(), hours, this.settings.businessHours).toISOString();
  }

  public caseNumberFor(id: number, createdIso?: string): string {
    const year = (createdIso ? new Date(createdIso) : new Date()).getFullYear();
    return `${this.settings.casePrefix || 'CS'}-${year}-${('000' + id).slice(-Math.max(4, String(id).length))}`;
  }

  public findDuplicates(key: EntityKey, r: IRecord): IRecord[] {
    const rows = this.data[key] || [];
    const email = String(r.email || '').trim().toLowerCase();
    const name = String(r.title || `${r.firstName || ''} ${r.lastName || ''}`).trim().toLowerCase();
    const company = String(r.company || '').trim().toLowerCase();
    return rows.filter(x => x.id !== r.id && (
      (email && String(x.email || '').toLowerCase() === email) ||
      (key === 'accounts' && name && String(x.title || '').trim().toLowerCase() === name) ||
      (key !== 'accounts' && name.length > 3 && company && String(x.title || '').trim().toLowerCase() === name && String(x.company || '').toLowerCase() === company)
    )).slice(0, 5);
  }

  private async nextRoundRobin(): Promise<IUserRef | undefined> {
    const a = this.settings.assignment;
    if (!a.roundRobin.length) { return undefined; }
    const idx = (a.lastIndex + 1) % a.roundRobin.length;
    const next = { ...this.settings, assignment: { ...a, lastIndex: idx } };
    await this.ds.saveSettings(next).catch(() => undefined);
    return a.roundRobin[idx];
  }

  // ------------------------------------------------------------------ save pipeline

  private async beforeSave(key: EntityKey, r: IRecord, prev?: IRecord): Promise<IRecord> {
    const x: IRecord = { ...r };
    const isNew = !x.id;
    switch (key) {
      case 'leads':
        x.title = `${x.firstName || ''} ${x.lastName || ''}`.trim() || x.email || x.company || 'Lead';
        x.leadScore = this.scoreLead({ ...prev, ...x });
        if ('phone' in x) { x.whatsapp = waNumber(x.phone, this.settings.whatsapp.defaultCountryCode); }
        if (isNew && !x.owner && this.settings.assignment.autoAssignNewLeads) {
          // eslint-disable-next-line require-atomic-updates -- x is a local copy
          x.owner = await this.nextRoundRobin();
        }
        break;
      case 'contacts':
        x.title = `${x.firstName || ''} ${x.lastName || ''}`.trim() || x.email || 'Contact';
        if ('phone' in x || 'mobile' in x) { x.whatsapp = waNumber(x.mobile || x.phone, this.settings.whatsapp.defaultCountryCode); }
        if (x.marketingConsent && !(prev && prev.marketingConsent) && !x.consentDate) { x.consentDate = new Date().toISOString(); }
        break;
      case 'opportunities': {
        const stageChanged = !prev || prev.stage !== x.stage;
        if (x.stage && stageChanged) {
          const st = this.stage(x.stage);
          if (!prev || x.probability === undefined || x.probability === prev.probability) { x.probability = st.probability; }
          x.forecastCategory = st.forecast;
          if (x.stage === WON || x.stage === LOST) {
            x.actualCloseDate = x.actualCloseDate || new Date().toISOString();
          } else if (prev && (prev.stage === WON || prev.stage === LOST)) {
            x.actualCloseDate = undefined;
          }
        }
        if (!x.title) { x.title = `${this.display('accounts', x.accountRef) || x.company || 'New'} deal`; }
        break;
      }
      case 'quotes': {
        if (isNew || !x.quoteNumber) {
          await this.ensure(['quotes']);
          // eslint-disable-next-line require-atomic-updates -- x is a local copy
          x.quoteNumber = this.nextNumber('quotes', 'quoteNumber', this.settings.quote.prefix || 'Q');
        }
        if (!x.validUntil) { x.validUntil = new Date(Date.now() + this.settings.quote.validityDays * 86400000).toISOString(); }
        if (!x.terms && isNew) { x.terms = this.settings.quote.terms; }
        const t = quoteTotals(x.lineItems || []);
        if (prev && prev.signatureStatus === 'Signed' && (Math.abs((prev.grandTotal || 0) - t.grandTotal) > 0.005 || JSON.stringify(prev.lineItems || []) !== JSON.stringify(x.lineItems || []))) {
          throw new Error('This quote was signed electronically, so its line items can no longer change. Create a new quote for the changes.');
        }
        if (!x.currency) { x.currency = this.settings.currency || 'USD'; }
        Object.keys(t).forEach(k => { x[k] = (t as any)[k]; });
        // Any price change after approval requires a new approval.
        if (prev && prev.approvalStatus === 'Approved' && Math.abs((prev.grandTotal || 0) - t.grandTotal) > 0.005) {
          x.approvalStatus = '';
          x.approvalNote = 'Approval reset: the quote total changed after approval.';
        }
        if (!x.title) { x.title = `Quote for ${this.display('accounts', x.accountRef) || 'customer'}`; }
        break;
      }
      case 'cases': {
        // Case numbers come from the SharePoint item ID (set right after the first save), so cases created by
        // Power Automate (email-to-case, website) and in the CRM never collide.
        if (!isNew && !x.caseNumber) { x.caseNumber = this.caseNumberFor(x.id!, prev?.created); }
        const created = prev?.created || new Date().toISOString();
        if (isNew || (prev && prev.priority !== x.priority) || !x.slaDue) {
          x.slaDue = this.slaDue(x.priority || 'Normal', created);
        }
        if (!x.firstResponseDue && !x.firstRespondedAt) { x.firstResponseDue = this.firstResponseDue(x.priority || 'Normal', created); }
        if (prev && !prev.firstRespondedAt && !x.firstRespondedAt && prev.caseStatus === 'New' && x.caseStatus && x.caseStatus !== 'New') {
          x.firstRespondedAt = new Date().toISOString();
        }
        const waiting = x.caseStatus === 'Waiting on customer';
        if (waiting && this.settings.slaPauseWhenWaiting && !(prev && prev.slaPausedAt)) { x.slaPausedAt = new Date().toISOString(); }
        if (!waiting && (prev?.slaPausedAt || x.slaPausedAt)) {
          // Resume: push the deadline out by the working time spent waiting on the customer.
          const paused = businessHoursBetween(new Date(prev?.slaPausedAt || x.slaPausedAt), new Date(), this.settings.businessHours);
          if (x.slaDue && paused > 0) { x.slaDue = addBusinessHours(new Date(x.slaDue), paused, this.settings.businessHours).toISOString(); }
          x.slaPausedAt = undefined;
        }
        if ((x.caseStatus === 'Resolved' || x.caseStatus === 'Closed') && !(prev && prev.resolvedDate)) { x.resolvedDate = new Date().toISOString(); }
        if (x.caseStatus && x.caseStatus !== 'Resolved' && x.caseStatus !== 'Closed' && prev?.resolvedDate) { x.resolvedDate = undefined; }
        break;
      }
      case 'subscriptions': {
        if (!x.billingCycle) { x.billingCycle = 'Annual'; }
        if (x.quantity === undefined || x.quantity === null) { x.quantity = 1; }
        if (!x.startDate && isNew) { x.startDate = new Date().toISOString(); }
        if (!x.renewalDate && x.startDate) { x.renewalDate = addPeriod(x.startDate, x.billingCycle); }
        x.mrr = subscriptionMrr(x);
        x.arr = Math.round(x.mrr * 12 * 100) / 100;
        if ((x.subStatus === 'Cancelled' || x.subStatus === 'Expired') && !x.endDate) { x.endDate = x.renewalDate || new Date().toISOString(); }
        if (!x.title) { x.title = `${this.display('products', x.productRef) || 'Subscription'} - ${this.display('accounts', x.accountRef) || 'customer'}`; }
        break;
      }
      case 'invoices': {
        const inv = this.settings.invoice;
        if (isNew || !x.invoiceNumber) {
          await this.ensure(['invoices']);
          // eslint-disable-next-line require-atomic-updates -- x is a local copy
          x.invoiceNumber = this.nextNumber('invoices', 'invoiceNumber', inv.prefix || 'INV');
        }
        if (!x.issueDate) { x.issueDate = new Date().toISOString(); }
        if (!x.currency) { x.currency = this.settings.currency || 'USD'; }
        if (!x.dueDate) { x.dueDate = new Date(new Date(x.issueDate).getTime() + (inv.paymentTermsDays || 0) * DAY).toISOString(); }
        if (!x.terms && isNew) { x.terms = inv.footer; }
        const t = quoteTotals(x.lineItems || []);
        Object.keys(t).forEach(k => { x[k] = (t as any)[k]; });
        if (x.invoiceStatus !== 'Void' && t.grandTotal > 0 && (Number(x.amountPaid) || 0) >= t.grandTotal - 0.005) { x.invoiceStatus = 'Paid'; }
        if (x.invoiceStatus === 'Paid') {
          if (!x.paidDate) { x.paidDate = new Date().toISOString(); }
          if (!x.amountPaid) { x.amountPaid = t.grandTotal; }
        }
        if (!x.title) { x.title = `Invoice for ${this.display('accounts', x.accountRef) || 'customer'}`; }
        break;
      }
      case 'demos':
        if (x.scheduledDate && (x.status === 'Requested' || !x.status)) { x.status = 'Scheduled'; }
        if (!x.title) { x.title = `Demo - ${x.company || x.contactName || ''}`.trim(); }
        break;
      case 'activities':
        if (x.activityType !== 'Task' && x.completed === undefined) { x.completed = true; }
        break;
    }
    return x;
  }

  private diff(key: EntityKey, prev: IRecord | undefined, next: IRecord): Record<string, [unknown, unknown]> {
    const out: Record<string, [unknown, unknown]> = {};
    const norm = (v: any): any => (v && typeof v === 'object' && 'title' in v && 'id' in v ? v.title : (Array.isArray(v) ? v.join(', ') : v));
    this.def(key).fields.forEach(fd => {
      if (!(fd.key in next) || fd.type === 'json') { return; }
      const a = norm(prev?.[fd.key]);
      const b = norm(next[fd.key]);
      if ((a ?? '') !== (b ?? '') && !(fd.type.indexOf('date') === 0 && a && b && new Date(a).getTime() === new Date(b).getTime())) {
        out[fd.label] = [a ?? '', b ?? ''];
      }
    });
    return out;
  }

  private afterSave(key: EntityKey, prev: IRecord | undefined, x: IRecord): void {
    const isNew = !prev;
    const d = this.def(key);
    const summary = (s: string): string => `${s}: ${x.title || x.quoteNumber || ''}`;
    const data = { id: x.id, entity: key, ...this.plain(x) };
    if (isNew) {
      const ev = key === 'demos' ? 'demo.requested' : `${d.singular.toLowerCase().replace(/\s/g, '')}.created`;
      this.webhooks.emit(ev, data, summary(`New ${d.singular.toLowerCase()}`));
    }
    if (key === 'leads' && prev && x.owner && prev.owner?.id !== x.owner.id) {
      this.webhooks.emit('lead.assigned', data, `${x.title} assigned to ${x.owner.title}`);
    }
    if (key === 'opportunities' && prev && prev.stage !== x.stage) {
      this.logActivity('opportunities', x.id!, `Stage: ${prev.stage} → ${x.stage}`, x.stage === LOST ? `Lost reason: ${x.lostReason || '-'}` : '').catch(() => undefined);
      this.webhooks.emit('opportunity.stageChanged', { ...data, previousStage: prev.stage }, `${x.title} moved to ${x.stage}`);
    }
    if (key === 'opportunities' && x.stage === WON && (!prev || prev.stage !== WON)) {
      this.webhooks.emit('opportunity.won', data, `🏆 Deal won: ${x.title} (${x.amount || 0} ${this.settings.currency})`);
    }
    if (key === 'opportunities' && x.stage === LOST && (!prev || prev.stage !== LOST)) {
      this.webhooks.emit('opportunity.lost', data, `Deal lost: ${x.title}`);
    }
    if (key === 'demos' && prev && prev.status !== x.status) {
      if (x.status === 'Scheduled') { this.webhooks.emit('demo.scheduled', data, summary('Demo scheduled')); }
      if (x.status === 'Completed') { this.webhooks.emit('demo.completed', data, summary('Demo completed')); }
    }
    if (key === 'quotes' && prev && prev.quoteStatus !== x.quoteStatus) {
      if (x.quoteStatus === 'Sent') { this.webhooks.emit('quote.sent', data, summary('Quote sent')); }
      if (x.quoteStatus === 'Accepted') { this.webhooks.emit('quote.accepted', data, summary('Quote accepted')); }
    }
    if (key === 'invoices' && prev && prev.invoiceStatus !== x.invoiceStatus) {
      if (x.invoiceStatus === 'Sent') { this.webhooks.emit('invoice.sent', data, summary('Invoice sent')); }
      if (x.invoiceStatus === 'Paid') { this.webhooks.emit('invoice.paid', data, `💰 Invoice paid: ${x.invoiceNumber} (${x.grandTotal || 0} ${this.settings.currency})`); }
      if (x.invoiceStatus === 'Overdue') { this.webhooks.emit('invoice.overdue', data, summary('Invoice overdue')); }
    }
    if (key === 'subscriptions' && prev && prev.subStatus !== x.subStatus && x.subStatus === 'Cancelled') {
      this.webhooks.emit('subscription.cancelled', data, summary('Subscription cancelled'));
    }
    if (key === 'opportunities' && x.dealType === 'Renewal' && x.subscriptionRef && prev && prev.stage !== x.stage && (x.stage === WON || x.stage === LOST)) {
      this.applyRenewalOutcome(x).catch(() => undefined);
    }
    if (key === 'cases' && prev && prev.caseStatus !== x.caseStatus && x.caseStatus === 'Resolved') {
      this.webhooks.emit('case.resolved', data, summary('Case resolved'));
    }
  }

  private plain(x: IRecord): Record<string, unknown> {
    const out: Record<string, unknown> = {};
    Object.keys(x).forEach(k => {
      const v = x[k];
      out[k] = v && typeof v === 'object' && 'id' in v && 'title' in v ? { id: v.id, name: v.title, email: v.email } : v;
    });
    return out;
  }

  /** Full save pipeline: business rules → SharePoint → audit → events → cache. */
  public async save(key: EntityKey, r: IRecord): Promise<IRecord> {
    if (!this.can(r.id ? 'edit' : 'create', key)) {
      throw new Error(this.license?.canWrite ? 'You do not have permission to change this record.' : this.license.message);
    }
    const prev = r.id ? this.find(key, r.id) : undefined;
    const x = await this.beforeSave(key, r, prev);
    const id = await this.ds.save(key, x);
    if (key === 'cases' && !x.caseNumber) { await this.ds.save('cases', { id, caseNumber: this.caseNumberFor(id) }); }
    const fresh = await this.ds.getOne(key, id);
    this.upsertCache(key, fresh);
    const changes = this.diff(key, prev, x);
    if (!prev || Object.keys(changes).length) {
      this.ds.writeAudit(prev ? 'Updated' : 'Created', key, id, fresh.title || fresh.quoteNumber || '', prev ? changes : undefined).catch(() => undefined);
    }
    this.afterSave(key, prev, fresh);
    return fresh;
  }

  public async patch(key: EntityKey, id: number, fields: Record<string, unknown>): Promise<IRecord> {
    const prev = this.find(key, id) || await this.ds.getOne(key, id);
    return this.save(key, { ...prev, ...fields, id });
  }

  public async bulkPatch(key: EntityKey, ids: number[], fields: Record<string, unknown>, onProgress?: (n: number) => void): Promise<void> {
    let n = 0;
    for (const id of ids) {
      await this.patch(key, id, fields);
      onProgress?.(++n);
    }
  }

  public async remove(key: EntityKey, id: number): Promise<void> {
    if (!this.can('delete', key)) { throw new Error('You do not have permission to delete records.'); }
    const prev = this.find(key, id);
    await this.ds.remove(key, id);
    this.removeCache(key, id);
    this.ds.writeAudit('Deleted', key, id, prev?.title || '', undefined).catch(() => undefined);
  }

  public async logActivity(key: EntityKey, id: number, title: string, details: string = '', type: string = 'Note'): Promise<void> {
    if (!this.license?.canWrite) { return; }
    const rec = await this.ds.save('activities', {
      title, details, activityType: type, relatedType: this.def(key).activityName, relatedId: id, completed: true
    });
    if (this.data.activities) {
      this.upsertCache('activities', await this.ds.getOne('activities', rec));
    }
  }

  // ------------------------------------------------------------------ composite operations

  public async assignRoundRobin(leadIds: number[]): Promise<number> {
    const users = this.settings.assignment.roundRobin;
    if (!users.length) { throw new Error('Add sales reps to the round-robin list in Settings > Automation first.'); }
    let idx = this.settings.assignment.lastIndex;
    for (const id of leadIds) {
      idx = (idx + 1) % users.length;
      await this.patch('leads', id, { owner: users[idx] });
    }
    await this.ds.saveSettings({ ...this.settings, assignment: { ...this.settings.assignment, lastIndex: idx } });
    return leadIds.length;
  }

  public async rescoreLeads(): Promise<number> {
    await this.ensure(['leads', 'activities', 'demos']);
    let n = 0;
    for (const l of this.data.leads || []) {
      const s = this.scoreLead(l);
      if (s !== l.leadScore) {
        await this.ds.save('leads', { id: l.id, leadScore: s });
        n++;
      }
    }
    if (n) { await this.refresh('leads'); }
    return n;
  }

  public async convertLead(lead: IRecord, o: IConvertOptions): Promise<{ accountId?: number; contactId?: number; opportunityId?: number }> {
    let accountId = o.accountId;
    let contactId = o.contactId;
    let opportunityId: number | undefined;
    if (!accountId && o.createAccount && lead.company) {
      accountId = (await this.save('accounts', {
        title: lead.company, accountType: 'Prospect', industry: lead.industry, website: lead.website, phone: lead.phone,
        country: lead.country, owner: lead.owner || this.me, description: `Created from lead ${lead.title}`
      })).id;
    }
    if (!contactId && o.createContact) {
      contactId = (await this.save('contacts', {
        firstName: lead.firstName, lastName: lead.lastName || '-', email: lead.email, phone: lead.phone, jobTitle: lead.jobTitle,
        country: lead.country, accountRef: accountId, owner: lead.owner || this.me
      })).id;
    }
    if (o.createOpportunity) {
      opportunityId = (await this.save('opportunities', {
        title: o.opportunityName, accountRef: accountId, contactRef: contactId, leadRef: lead.id, campaignRef: lead.campaignRef,
        stage: this.openStages()[1] || this.openStages()[0], amount: o.amount || 0, expectedCloseDate: o.expectedCloseDate,
        owner: lead.owner || this.me, company: lead.company, contactName: lead.title, email: lead.email, phone: lead.phone, notes: lead.description
      })).id;
    }
    await this.patch('leads', lead.id!, { status: 'Converted', accountRef: accountId, contactRef: contactId, opportunityRef: opportunityId });
    await this.logActivity('leads', lead.id!, 'Lead converted', [
      accountId ? `Account #${accountId}` : '', contactId ? `Contact #${contactId}` : '', opportunityId ? `Opportunity #${opportunityId}` : ''
    ].filter(Boolean).join(', '));
    this.webhooks.emit('lead.converted', { id: lead.id, accountId, contactId, opportunityId, email: lead.email, company: lead.company }, `Lead converted: ${lead.title}`);
    return { accountId, contactId, opportunityId };
  }

  public async acceptQuote(q: IRecord, closeOpportunity: boolean): Promise<void> {
    await this.patch('quotes', q.id!, { quoteStatus: 'Accepted' });
    if (closeOpportunity && q.opportunityRef) {
      await this.ensure(['opportunities']);
      await this.patch('opportunities', q.opportunityRef, { stage: WON, amount: q.grandTotal });
    }
    if (q.accountRef) {
      await this.ensure(['accounts']);
      const acc = this.find('accounts', q.accountRef);
      if (acc && acc.accountType === 'Prospect') { await this.patch('accounts', acc.id!, { accountType: 'Customer' }); }
    }
  }

  // ------------------------------------------------------------------ subscriptions, renewals and invoicing

  /** Recurring quote lines (Monthly/Annual) become subscriptions on the quote's account. */
  public async createSubscriptionsFromQuote(q: IRecord): Promise<IRecord[]> {
    const lines = ((q.lineItems || []) as ILineItem[]).filter(i => i.billing === 'Monthly' || i.billing === 'Annual');
    const out: IRecord[] = [];
    for (const i of lines) {
      const perSeat = (Number(i.unitPrice) || 0) * (1 - (Number(i.discountPct) || 0) / 100);
      out.push(await this.save('subscriptions', {
        title: `${i.name} - ${this.display('accounts', q.accountRef) || 'customer'}`,
        accountRef: q.accountRef, contactRef: q.contactRef, productRef: i.productId, opportunityRef: q.opportunityRef, quoteRef: q.id,
        subStatus: 'Active', billingCycle: i.billing, quantity: Number(i.quantity) || 1, unitPrice: Math.round(perSeat * 100) / 100,
        startDate: new Date().toISOString(), autoRenew: true, owner: q.owner || this.me
      }));
    }
    if (out.length) {
      await this.logActivity('quotes', q.id!, `${out.length} subscription${out.length === 1 ? '' : 's'} created`, out.map(s => s.title).join('\n'));
      out.forEach(s => this.webhooks.emit('subscription.created', { id: s.id, entity: 'subscriptions', ...this.plain(s) }, `New subscription: ${s.title}`));
    }
    return out;
  }

  public async createInvoiceFromQuote(q: IRecord): Promise<IRecord> {
    return this.save('invoices', {
      title: q.title, accountRef: q.accountRef, contactRef: q.contactRef, quoteRef: q.id, opportunityRef: q.opportunityRef,
      invoiceStatus: 'Draft', lineItems: (q.lineItems || []).map((i: ILineItem) => ({ ...i })), owner: q.owner || this.me
    });
  }

  /** Invoice for the next billing period of a subscription. */
  public async createInvoiceFromSubscription(s: IRecord): Promise<IRecord> {
    const from = s.renewalDate && new Date(s.renewalDate).getTime() > Date.now() ? addPeriod(s.renewalDate, s.billingCycle, -1) : (s.renewalDate || s.startDate);
    const to = addPeriod(from, s.billingCycle);
    const product = s.productRef ? this.find('products', s.productRef) : undefined;
    const period = `${new Date(from).toLocaleDateString()} – ${new Date(new Date(to).getTime() - DAY).toLocaleDateString()}`;
    return this.save('invoices', {
      title: `${s.title} (${period})`, accountRef: s.accountRef, contactRef: s.contactRef, subscriptionRef: s.id, invoiceStatus: 'Draft',
      lineItems: [{
        id: String(Date.now()), productId: s.productRef, name: product?.title || s.title || 'Subscription', description: `${s.billingCycle} subscription, ${period}`,
        quantity: Number(s.quantity) || 1, unitPrice: Number(s.unitPrice) || 0, discountPct: 0, taxPct: Number(product?.taxRate) || 0, billing: s.billingCycle
      }],
      owner: s.owner || this.me
    });
  }

  public async recordPayment(inv: IRecord, amount: number, provider?: string, reference?: string): Promise<IRecord> {
    const paid = Math.round(((Number(inv.amountPaid) || 0) + amount) * 100) / 100;
    const saved = await this.patch('invoices', inv.id!, {
      amountPaid: paid, paymentProvider: provider || inv.paymentProvider, externalPaymentId: reference || inv.externalPaymentId,
      ...(paid >= (inv.grandTotal || 0) - 0.005 ? { invoiceStatus: 'Paid', paidDate: new Date().toISOString() } : {})
    });
    await this.logActivity('invoices', inv.id!, `Payment recorded: ${amount} ${this.settings.currency}`, [provider, reference].filter(Boolean).join(' · '));
    return saved;
  }

  /** A won renewal extends the subscription one period; a lost renewal cancels it at the renewal date. */
  private async applyRenewalOutcome(opp: IRecord): Promise<void> {
    await this.ensure(['subscriptions']);
    const s = this.find('subscriptions', opp.subscriptionRef);
    if (!s) { return; }
    if (opp.stage === WON) {
      const saved = await this.patch('subscriptions', s.id!, { subStatus: 'Active', renewalDate: addPeriod(s.renewalDate, s.billingCycle), endDate: undefined });
      await this.logActivity('subscriptions', s.id!, 'Renewed', `Renewal ${opp.title} won. Next renewal ${new Date(saved.renewalDate).toLocaleDateString()}.`);
      this.webhooks.emit('subscription.renewed', { id: s.id, entity: 'subscriptions', ...this.plain(saved) }, `Subscription renewed: ${s.title}`);
    } else {
      await this.patch('subscriptions', s.id!, { subStatus: 'Cancelled', autoRenew: false, endDate: s.renewalDate });
    }
  }

  /**
   * Housekeeping that a server would normally do, run when a manager/admin opens the app (at most every 6 hours):
   * renewal opportunities ahead of annual renewals, overdue invoices, expired subscriptions, monthly roll-over.
   */
  public async runMaintenance(force: boolean = false): Promise<IMaintenanceResult> {
    const res: IMaintenanceResult = { renewals: 0, overdue: 0, expired: 0, rolled: 0 };
    if (!this.license?.canWrite || !(this.role === 'admin' || this.role === 'manager')) { return res; }
    const lsKey = `brillienta-crm:maintenance:${this.ds.rest.webUrl}`;
    try {
      const last = Number(window.localStorage.getItem(lsKey) || 0);
      if (!force && Date.now() - last < 6 * 3600000) { return res; }
      window.localStorage.setItem(lsKey, String(Date.now()));
    } catch { /* storage unavailable: run anyway */ }
    await this.ensure(['subscriptions', 'invoices', 'opportunities']);
    const now = Date.now();
    const r = this.settings.renewals;
    for (const s of this.data.subscriptions || []) {
      const due = s.renewalDate ? new Date(s.renewalDate).getTime() : 0;
      const live = s.subStatus === 'Active' || s.subStatus === 'Pending renewal' || s.subStatus === 'Trial';
      if (!live || !due) { continue; }
      if (due < now && !s.autoRenew) {
        await this.patch('subscriptions', s.id!, { subStatus: 'Expired', endDate: s.renewalDate });
        res.expired++;
      } else if (due < now && s.billingCycle === 'Monthly') {
        let next = s.renewalDate;
        while (new Date(next).getTime() < now) { next = addPeriod(next, 'Monthly'); }
        await this.patch('subscriptions', s.id!, { renewalDate: next });
        res.rolled++;
      } else if (r.autoCreateOpportunities && s.billingCycle !== 'Monthly' && due - now <= r.leadDays * DAY) {
        const exists = (this.data.opportunities || []).some(o => o.subscriptionRef === s.id && o.dealType === 'Renewal' && dayKey(o.expectedCloseDate) === dayKey(s.renewalDate));
        if (exists) { continue; }
        const stages = this.openStages();
        await this.save('opportunities', {
          title: `Renewal ${new Date(s.renewalDate).getFullYear()} - ${s.title}`, dealType: 'Renewal', subscriptionRef: s.id,
          accountRef: s.accountRef, contactRef: s.contactRef, stage: stages[Math.min(2, stages.length - 1)] || stages[0],
          amount: Math.round((s.arr || subscriptionMrr(s) * 12) * 100) / 100, expectedCloseDate: s.renewalDate,
          owner: s.owner || this.me, notes: `Automatic renewal opportunity for subscription #${s.id}.`
        });
        if (s.subStatus !== 'Pending renewal') { await this.patch('subscriptions', s.id!, { subStatus: 'Pending renewal' }); }
        res.renewals++;
      }
    }
    for (const inv of this.data.invoices || []) {
      if (inv.invoiceStatus === 'Sent' && inv.dueDate && new Date(inv.dueDate).getTime() < now - DAY) {
        await this.patch('invoices', inv.id!, { invoiceStatus: 'Overdue' });
        res.overdue++;
      }
    }
    return res;
  }

  /**
   * Completes cases created outside the CRM (email-to-case and website flows): case number, business-hours SLA
   * deadlines, and resuming SLAs that a flow re-opened while paused. Runs for support/managers/admins on open.
   */
  public async normalizeCases(): Promise<number> {
    if (!this.license?.canWrite || !this.hasFeature('cases') || ['admin', 'manager', 'support'].indexOf(this.role) < 0) { return 0; }
    await this.ensure(['cases']);
    let n = 0;
    for (const c of this.data.cases || []) {
      const open = c.caseStatus !== 'Resolved' && c.caseStatus !== 'Closed';
      const resume = !!c.slaPausedAt && c.caseStatus !== 'Waiting on customer';
      if (!c.caseNumber || (open && (!c.slaDue || resume || (!c.firstResponseDue && !c.firstRespondedAt)))) {
        await this.patch('cases', c.id!, {});
        n++;
      }
    }
    return n;
  }

  // ------------------------------------------------------------------ e-signature

  public signingLink(q: IRecord): string {
    const base = (this.settings.quote.signingUrl || '').trim();
    if (!base || !q.signatureToken) { return ''; }
    return `${base}${base.indexOf('?') >= 0 ? '&' : '?'}q=${q.id}&t=${encodeURIComponent(q.signatureToken)}`;
  }

  /** Issues a new single-use signing link for the quote (any earlier link stops working). */
  public async sendForSignature(q: IRecord, email: string): Promise<{ quote: IRecord; link: string }> {
    if (!/^https:\/\//i.test(this.settings.quote.signingUrl || '')) {
      throw new Error('Set the e-signature link in Settings › Company › Quotes first (the HTTP URL of the flow “Quote e-signature”).');
    }
    if (!this.quoteApproval(q).approved) { throw new Error('This quote needs discount approval before it can be sent for signature.'); }
    const bytes = new Uint8Array(24);
    window.crypto.getRandomValues(bytes);
    const token = Array.prototype.map.call(bytes, (b: number) => ('0' + b.toString(16)).slice(-2)).join('');
    const quote = await this.patch('quotes', q.id!, {
      signatureToken: token, signatureStatus: 'Sent', signatureEmail: email.trim().toLowerCase(), signatureNote: '',
      ...(q.quoteStatus === 'Draft' || !q.quoteStatus ? { quoteStatus: 'Sent' } : {})
    });
    await this.logActivity('quotes', q.id!, 'Sent for e-signature', `To: ${email}`);
    return { quote, link: this.signingLink(quote) };
  }

  /**
   * Completes quotes the customer signed on the signing page (the flow only records the signature):
   * closes the opportunity as won, makes the account a customer, creates subscriptions. Runs on open.
   */
  public async processSignedQuotes(): Promise<number> {
    if (!this.can('edit', 'quotes')) { return 0; }
    await this.ensure(['quotes']);
    const todo = (this.data.quotes || []).filter(q => (q.signatureStatus === 'Signed' || q.signatureStatus === 'Declined') && !q.signatureProcessed);
    for (const q of todo) {
      await this.patch('quotes', q.id!, { signatureProcessed: true });
      if (q.signatureStatus === 'Declined') {
        this.webhooks.emit('quote.declined', { id: q.id, quoteNumber: q.quoteNumber }, `Quote declined: ${q.quoteNumber}`);
        continue;
      }
      await this.acceptQuote(q, !!q.opportunityRef && this.settings.quote.signatureClosesOpportunity);
      if (this.settings.invoice.createSubscriptionsOnAccept && this.can('create', 'subscriptions') && !(this.data.subscriptions || []).some(x => x.quoteRef === q.id)) {
        await this.ensure(['subscriptions']);
        if (!(this.data.subscriptions || []).some(x => x.quoteRef === q.id)) { await this.createSubscriptionsFromQuote(q); }
      }
      this.webhooks.emit('quote.signed', { id: q.id, quoteNumber: q.quoteNumber, signedBy: q.signatureName, signedDate: q.signedDate, total: q.grandTotal },
        `✍ Quote signed: ${q.quoteNumber} by ${q.signatureName} (${q.grandTotal} ${this.settings.currency})`);
    }
    return todo.length;
  }

  /** Discount as % of subtotal and whether it needs manager approval (Settings › Company › Quotes). */
  public quoteApproval(q: IRecord): { discountPct: number; required: boolean; approved: boolean } {
    const t = quoteTotals(q.lineItems || []);
    const discountPct = t.subtotal > 0 ? Math.round((t.discountTotal / t.subtotal) * 1000) / 10 : 0;
    const threshold = this.settings.quote.approvalDiscountPct;
    const required = threshold > 0 && discountPct > threshold;
    return { discountPct, required, approved: !required || q.approvalStatus === 'Approved' };
  }

  public async requestQuoteApproval(q: IRecord): Promise<IRecord> {
    const saved = await this.patch('quotes', q.id!, { approvalStatus: 'Pending', approvalNote: '' });
    await this.logActivity('quotes', q.id!, 'Approval requested', `Discount ${this.quoteApproval(saved).discountPct}% (threshold ${this.settings.quote.approvalDiscountPct}%).`);
    this.webhooks.emit('quote.approvalRequested', { id: q.id, quoteNumber: q.quoteNumber, total: saved.grandTotal }, `Approval requested: ${q.quoteNumber}`);
    return saved;
  }

  public async sendEmail(to: string, subject: string, body: string, related?: { key: EntityKey; id: number }, cc: string[] = []): Promise<void> {
    await this.graph.sendMail([to], subject, body, cc);
    if (related) {
      await this.logActivity(related.key, related.id, `Email: ${subject}`, `To: ${to}\n\n${body}`, 'Email');
      const c = related.key === 'cases' ? this.find('cases', related.id) : undefined;
      if (c && !c.firstRespondedAt) {
        await this.patch('cases', c.id!, { firstRespondedAt: new Date().toISOString(), ...(c.caseStatus === 'New' ? { caseStatus: 'In progress' } : {}) });
      }
    }
  }

  public async scheduleDemo(demo: IRecord, startIso: string, minutes: number, attendees: string[], subject: string, agenda: string): Promise<IRecord> {
    let joinUrl = demo.meetingLink || '';
    if (this.settings.integrations.graphMeetings) {
      const m = await this.graph.createTeamsMeeting(subject, startIso, minutes, attendees, agenda);
      joinUrl = m.joinUrl || joinUrl;
    }
    const saved = await this.patch('demos', demo.id!, {
      scheduledDate: startIso, durationMinutes: minutes, meetingLink: joinUrl, status: 'Scheduled', owner: demo.owner || this.me
    });
    await this.logActivity('demos', demo.id!, `Teams meeting scheduled`, `${new Date(startIso).toLocaleString()} (${minutes} min)\nAttendees: ${attendees.join(', ')}\n${joinUrl}`, 'Meeting');
    return saved;
  }

  public fillTemplate(text: string, r: IRecord): string {
    const vars: Record<string, string> = {
      firstName: r.firstName || (r.contactRef ? this.find('contacts', r.contactRef)?.firstName : '') || String(r.contactName || (r.contactRef ? '' : r.title) || '').split(' ')[0] || '',
      lastName: r.lastName || '',
      fullName: r.title || r.contactName || '',
      company: r.company || this.display('accounts', r.accountRef) || '',
      product: r.offer || this.settings.company.name || '',
      ownerName: r.owner?.title || '',
      senderName: this.me?.title || '',
      scheduledDate: r.scheduledDate ? new Date(r.scheduledDate).toLocaleString() : '',
      meetingLink: r.meetingLink || '',
      quoteNumber: r.quoteNumber || '',
      invoiceNumber: r.invoiceNumber || '',
      paymentLink: r.paymentLink || '',
      caseNumber: r.caseNumber || '',
      total: r.grandTotal !== undefined ? fmtMoney(Math.max(0, r.grandTotal - (Number(r.amountPaid) || 0)), this.settings.currency || 'USD') : '',
      companyName: this.settings.company.name || ''
    };
    return text.replace(/\{\{\s*(\w+)\s*\}\}/g, (_, k) => (k in vars ? vars[k] : ''));
  }
}
