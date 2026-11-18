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
  restrictRepsToOwnRecords: boolean;
  auditEnabled: boolean;
  integrations: {
    graphEmail: boolean;
    graphMeetings: boolean;
    teamsWebhookUrl: string;
  };
  webhooks: IWebhook[];
  emailTemplates: IEmailTemplate[];
  customFields: ICustomField[];
  quotas: IQuota[];
  quote: { prefix: string; validityDays: number; terms: string; footer: string; approvalDiscountPct: number };
  casePrefix: string;
  groups: { admins: string; managers: string; sales: string; support: string; readers: string };
}

export const WEBHOOK_EVENTS = [
  'lead.created', 'lead.converted', 'lead.assigned',
  'opportunity.created', 'opportunity.stageChanged', 'opportunity.won', 'opportunity.lost',
  'demo.requested', 'demo.scheduled', 'demo.completed',
  'quote.sent', 'quote.accepted', 'quote.approvalRequested',
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
    leadSources: ['AppSource', 'Azure Marketplace', 'Website', 'Website Demo Form', 'Referral', 'Event', 'Webinar', 'Email Campaign', 'Partner', 'Cold outreach', 'Other'],
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
    restrictRepsToOwnRecords: false,
    auditEnabled: true,
    integrations: { graphEmail: true, graphMeetings: true, teamsWebhookUrl: '' },
    webhooks: [],
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
      approvalDiscountPct: 15
    },
    casePrefix: 'CS',
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
  out.groups = { ...d.groups, ...(saved.groups || {}) };
  out.slaHours = { ...d.slaHours, ...(saved.slaHours || {}) };
  return out;
}
