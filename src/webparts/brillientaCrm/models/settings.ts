import { IBusinessHours, localTimeZone } from '../services/BusinessHours';
import { EntityKey, FieldType, IUserRef } from './model';

export interface IStageDef {
  name: string;
  probability: number;
  forecast: string;
}

export interface IWebhook {
  id: string;
  name: string;
  url: string;
  events: string[];
  active: boolean;
}

export interface IWhatsAppTemplate {
  id: string;
  name: string;
  /** Message text with {{placeholders}} (same as email templates). */
  body: string;
  /** Name of the approved template in WhatsApp Manager (Cloud API); its {{1}}, {{2}}… take the values in `params`. */
  cloudName?: string;
  language?: string;
  /** Placeholders that fill the approved template's {{1}}, {{2}}…, e.g. ["firstName", "quoteNumber"]. */
  params?: string[];
}

export interface IEmailTemplate {
  id: string;
  name: string;
  subject: string;
  body: string;
}

export interface ICustomField {
  entity: EntityKey;
  key: string;
  label: string;
  type: FieldType;
  choices?: string[];
  grid?: boolean;
}

export interface IQuota {
  user: IUserRef;
  quarterly: number;
}

export interface ICrmSettings {
  version: number;
  installDate: string;
  licenseKey: string;
  company: {
    name: string; address: string; email: string; phone: string; website: string; logoUrl: string; taxId: string;
  };
  currency: string;
  fiscalYearStartMonth: number;
  stages: IStageDef[];
  leadSources: string[];
  industries: string[];
  caseCategories: string[];
  lostReasons: string[];
  scoring: {
    rating: Record<string, number>;
    sources: Record<string, number>;
    titleKeywords: { keyword: string; points: number }[];
    hasPhone: number;
    hasCompany: number;
    perActivity: number;
    maxActivityPoints: number;
    hotThreshold: number;
  };
  assignment: {
    autoAssignNewLeads: boolean;
    roundRobin: IUserRef[];
    lastIndex: number;
  };
  slaHours: Record<string, number>;
  firstResponseHours: Record<string, number>;
  businessHours: IBusinessHours;
  /** SLA clock stops while a case is "Waiting on customer". */
  slaPauseWhenWaiting: boolean;
  restrictRepsToOwnRecords: boolean;
  auditEnabled: boolean;
  integrations: {
    graphEmail: boolean;
    graphMeetings: boolean;
    teamsWebhookUrl: string;
    /** Users may switch on automatic logging of Outlook email & meetings with CRM contacts/leads. */
    outlookSync: boolean;
    /** Days of mail/calendar to read the first time a user switches sync on. */
    syncLookbackDays: number;
    /** Never log mail with these domains (your own domain is always excluded). */
    internalDomains: string[];
  };
  webhooks: IWebhook[];
  emailTemplates: IEmailTemplate[];
  whatsapp: {
    enabled: boolean;
    /** "click": opens WhatsApp with the message ready (free, any number). "cloud": sends through the WhatsApp Business Cloud API flow. */
    mode: 'click' | 'cloud';
    /** Country calling code added to local numbers, e.g. "91". */
    defaultCountryCode: string;
    templates: IWhatsAppTemplate[];
  };
  customFields: ICustomField[];
  quotas: IQuota[];
  quote: { prefix: string; validityDays: number; terms: string; footer: string; approvalDiscountPct: number; signingUrl: string; signatureClosesOpportunity: boolean };
  casePrefix: string;
  invoice: { prefix: string; paymentTermsDays: number; bankDetails: string; footer: string; createSubscriptionsOnAccept: boolean };
  renewals: { leadDays: number; autoCreateOpportunities: boolean };
  payments: { provider: 'none' | 'Razorpay' | 'Stripe' };
  accounting: { tallySalesLedger: string; tallyTaxLedger: string; xeroSalesAccount: string; xeroTaxType: string };
  groups: { admins: string; managers: string; sales: string; support: string; readers: string };
}

export const WEBHOOK_EVENTS = [
  'lead.created', 'lead.converted', 'lead.assigned',
  'opportunity.created', 'opportunity.stageChanged', 'opportunity.won', 'opportunity.lost',
  'demo.requested', 'demo.scheduled', 'demo.completed',
  'quote.sent', 'quote.accepted', 'quote.approvalRequested', 'quote.signed', 'quote.declined',
  'subscription.created', 'subscription.renewed', 'subscription.cancelled',
  'invoice.created', 'invoice.sent', 'invoice.paid', 'invoice.overdue',
  'case.created', 'case.resolved', 'case.slaBreached',
  'account.created', 'contact.created'
];

export function defaultSettings(): ICrmSettings {
  return {
    version: 2,
    installDate: new Date().toISOString(),
    licenseKey: '',
    company: { name: '', address: '', email: '', phone: '', website: '', logoUrl: '', taxId: '' },
    currency: 'USD',
    fiscalYearStartMonth: 1,
    stages: [
      { name: 'Prospecting', probability: 10, forecast: 'Pipeline' },
      { name: 'Qualification', probability: 25, forecast: 'Pipeline' },
      { name: 'Proposal', probability: 50, forecast: 'Best case' },
      { name: 'Negotiation', probability: 75, forecast: 'Commit' },
      { name: 'Closed Won', probability: 100, forecast: 'Closed' },
      { name: 'Closed Lost', probability: 0, forecast: 'Omitted' }
    ],
    leadSources: ['AppSource', 'Azure Marketplace', 'Website', 'Website Demo Form', 'WhatsApp', 'Referral', 'Event', 'Webinar', 'Email Campaign', 'Partner', 'Cold outreach', 'Other'],
    industries: ['Technology', 'Financial services', 'Healthcare', 'Manufacturing', 'Retail', 'Education', 'Government', 'Professional services', 'Non-profit', 'Other'],
    caseCategories: ['Question', 'Problem', 'Feature request', 'Billing', 'Onboarding', 'Other'],
    lostReasons: ['Price', 'Features', 'Competitor', 'Timing', 'No budget', 'No decision', 'Other'],
    scoring: {
      rating: { Hot: 30, Warm: 15, Cold: 0 },
      sources: { AppSource: 20, 'Website Demo Form': 25, Referral: 20, Partner: 15, Event: 10, Webinar: 10 },
      titleKeywords: [
        { keyword: 'CEO', points: 20 }, { keyword: 'CIO', points: 20 }, { keyword: 'CTO', points: 20 },
        { keyword: 'Director', points: 15 }, { keyword: 'Head', points: 12 }, { keyword: 'Manager', points: 8 }
      ],
      hasPhone: 5,
      hasCompany: 5,
      perActivity: 5,
      maxActivityPoints: 25,
      hotThreshold: 60
    },
    assignment: { autoAssignNewLeads: false, roundRobin: [], lastIndex: -1 },
    slaHours: { Critical: 4, High: 8, Normal: 24, Low: 72 },
    firstResponseHours: { Critical: 1, High: 2, Normal: 8, Low: 24 },
    businessHours: { enabled: false, timeZone: localTimeZone(), days: [1, 2, 3, 4, 5], start: '09:00', end: '18:00', holidays: [] },
    slaPauseWhenWaiting: true,
    restrictRepsToOwnRecords: false,
    auditEnabled: true,
    integrations: { graphEmail: true, graphMeetings: true, teamsWebhookUrl: '', outlookSync: true, syncLookbackDays: 30, internalDomains: [] },
    webhooks: [],
    whatsapp: {
      enabled: true, mode: 'click', defaultCountryCode: '91',
      templates: [
        { id: 'hello', name: 'Introduction', body: 'Hi {{firstName}}, this is {{senderName}} from {{companyName}}. Thanks for your interest in {{product}}. Is now a good time for a quick chat?', cloudName: 'crm_intro', language: 'en', params: ['firstName', 'senderName', 'companyName'] },
        { id: 'demo', name: 'Demo reminder', body: 'Hi {{firstName}}, a reminder of our demo on {{scheduledDate}}. Join here: {{meetingLink}}', cloudName: 'crm_demo_reminder', language: 'en', params: ['firstName', 'scheduledDate', 'meetingLink'] },
        { id: 'quote', name: 'Quote follow-up', body: 'Hi {{firstName}}, did you have a chance to look at quote {{quoteNumber}}? Happy to answer any questions.', cloudName: 'crm_quote_followup', language: 'en', params: ['firstName', 'quoteNumber'] },
        { id: 'pay', name: 'Payment reminder', body: 'Hi {{firstName}}, a friendly reminder that invoice {{invoiceNumber}} ({{total}}) is due. Pay online: {{paymentLink}}', cloudName: 'crm_payment_reminder', language: 'en', params: ['firstName', 'invoiceNumber', 'total', 'paymentLink'] }
      ]
    },
    emailTemplates: [
      {
        id: 'intro', name: 'Introduction',
        subject: 'Thanks for your interest in {{product}}',
        body: 'Hi {{firstName}},\n\nThank you for your interest. I would love to learn more about what {{company}} is looking for and show you how we can help.\n\nWould you have 30 minutes this week for a short call?\n\nBest regards,\n{{senderName}}'
      },
      {
        id: 'demo', name: 'Demo confirmation',
        subject: 'Your demo is confirmed',
        body: 'Hi {{firstName}},\n\nYour demo is confirmed for {{scheduledDate}}.\nJoin link: {{meetingLink}}\n\nLooking forward to speaking with you.\n\n{{senderName}}'
      },
      {
        id: 'followup', name: 'Follow-up',
        subject: 'Following up',
        body: 'Hi {{firstName}},\n\nI wanted to follow up on our recent conversation. Do you have any questions I can help with?\n\nBest regards,\n{{senderName}}'
      }
    ],
    customFields: [],
    quotas: [],
    quote: {
      prefix: 'Q', validityDays: 30,
      terms: 'Prices exclude applicable taxes unless stated. Payment due within 30 days of invoice. This quote is valid until the date shown.',
      footer: 'Thank you for your business.',
      approvalDiscountPct: 15, signingUrl: '', signatureClosesOpportunity: true
    },
    casePrefix: 'CS',
    invoice: { prefix: 'INV', paymentTermsDays: 30, bankDetails: '', footer: 'Thank you for your business.', createSubscriptionsOnAccept: true },
    renewals: { leadDays: 60, autoCreateOpportunities: true },
    payments: { provider: 'none' },
    accounting: { tallySalesLedger: 'Sales', tallyTaxLedger: 'Output GST', xeroSalesAccount: '200', xeroTaxType: 'OUTPUT' },
    groups: { admins: 'CRM Administrators', managers: 'CRM Sales Managers', sales: 'CRM Sales', support: 'CRM Support', readers: 'CRM Readers' }
  };
}

/** Deep-merges saved settings over defaults so new settings appear after upgrades. */
export function mergeSettings(saved: Partial<ICrmSettings> | undefined): ICrmSettings {
  const d = defaultSettings();
  if (!saved) {
    return d;
  }
  const out: ICrmSettings = { ...d, ...saved } as ICrmSettings;
  out.company = { ...d.company, ...(saved.company || {}) };
  out.scoring = { ...d.scoring, ...(saved.scoring || {}) };
  out.assignment = { ...d.assignment, ...(saved.assignment || {}) };
  out.integrations = { ...d.integrations, ...(saved.integrations || {}) };
  out.quote = { ...d.quote, ...(saved.quote || {}) };
  out.whatsapp = { ...d.whatsapp, ...(saved.whatsapp || {}) };
  out.invoice = { ...d.invoice, ...(saved.invoice || {}) };
  out.renewals = { ...d.renewals, ...(saved.renewals || {}) };
  out.payments = { ...d.payments, ...(saved.payments || {}) };
  out.accounting = { ...d.accounting, ...(saved.accounting || {}) };
  out.groups = { ...d.groups, ...(saved.groups || {}) };
  out.slaHours = { ...d.slaHours, ...(saved.slaHours || {}) };
  out.firstResponseHours = { ...d.firstResponseHours, ...(saved.firstResponseHours || {}) };
  out.businessHours = { ...d.businessHours, ...(saved.businessHours || {}) };
  return out;
}
