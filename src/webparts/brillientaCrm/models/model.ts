/* eslint-disable @typescript-eslint/no-explicit-any */

export interface IUserRef {
  id: number;
  title: string;
  email?: string;
}

export type FieldType =
  | 'text' | 'email' | 'phone' | 'url' | 'note' | 'choice' | 'number' | 'currency' | 'percent'
  | 'date' | 'datetime' | 'boolean' | 'user' | 'lookup' | 'tags' | 'json';

/** Choice lists that admins can edit in Settings. */
export type ChoiceSource = 'stages' | 'leadSources' | 'industries' | 'caseCategories' | 'lostReasons';

export interface IFieldDef {
  key: string;
  /** SharePoint internal column name. */
  column: string;
  label: string;
  type: FieldType;
  choices?: string[];
  choicesFrom?: ChoiceSource;
  lookup?: EntityKey;
  required?: boolean;
  section?: string;
  /** Show in the grid by default. */
  grid?: boolean;
  /** Offer as a grid filter. */
  filter?: boolean;
  /** Not editable in forms (computed by the engine). */
  readOnly?: boolean;
  /** Not shown in forms. */
  hidden?: boolean;
  defaultValue?: any;
  custom?: boolean;
  width?: number;
}

export type EntityKey =
  | 'leads' | 'accounts' | 'contacts' | 'opportunities' | 'quotes' | 'products'
  | 'demos' | 'cases' | 'campaigns' | 'activities' | 'subscriptions' | 'invoices';

export type SystemListKey = 'settings' | 'audit';
export type ListKey = EntityKey | SystemListKey;

export interface IRelatedDef {
  entity: EntityKey;
  field: string;
  label: string;
}

export interface IEntityDef {
  key: EntityKey;
  listTitle: string;
  singular: string;
  plural: string;
  icon: string;
  description: string;
  /** Name used in the Activities RelatedType column. */
  activityName: string;
  fields: IFieldDef[];
  searchFields: string[];
  related?: IRelatedDef[];
  /** Choice field used for the Kanban board. */
  boardField?: string;
  plan: 'standard' | 'enterprise';
  /** Area in the navigation. */
  area: 'Sales' | 'Marketing' | 'Service' | 'Catalog' | 'Work';
}

export interface IRecord {
  id?: number;
  title?: string;
  created?: string;
  modified?: string;
  author?: IUserRef;
  editor?: IUserRef;
  [key: string]: any;
}

export type CrmData = Partial<Record<EntityKey, IRecord[]>>;

/* ---------------------------------------------------------------- field helpers */

const f = (key: string, label: string, type: FieldType = 'text', extra: Partial<IFieldDef> = {}): IFieldDef => ({
  key, column: key.charAt(0).toUpperCase() + key.slice(1), label, type, ...extra
});
const title = (label: string, extra: Partial<IFieldDef> = {}): IFieldDef =>
  ({ key: 'title', column: 'Title', label, type: 'text', grid: true, required: true, section: 'General', ...extra });
const owner = (column: string): IFieldDef =>
  ({ key: 'owner', column, label: 'Owner', type: 'user', grid: true, filter: true, section: 'General' });
const ref = (key: string, label: string, lookup: EntityKey, extra: Partial<IFieldDef> = {}): IFieldDef =>
  f(key, label, 'lookup', { lookup, section: 'Relationships', ...extra });
const tags: IFieldDef = f('tags', 'Tags', 'tags', { section: 'General' });

export const LEAD_STATUSES = ['New', 'Contacted', 'Qualified', 'Unqualified', 'Converted'];
export const DEMO_STATUSES = ['Requested', 'Scheduled', 'Completed', 'No-show', 'Cancelled'];
export const QUOTE_STATUSES = ['Draft', 'Sent', 'Accepted', 'Rejected', 'Expired'];
export const CASE_STATUSES = ['New', 'In progress', 'Waiting on customer', 'Resolved', 'Closed'];
export const CASE_PRIORITIES = ['Low', 'Normal', 'High', 'Critical'];
export const FORECAST_CATEGORIES = ['Pipeline', 'Best case', 'Commit', 'Omitted', 'Closed'];
export const SUBSCRIPTION_STATUSES = ['Trial', 'Active', 'Pending renewal', 'Cancelled', 'Expired'];
export const INVOICE_STATUSES = ['Draft', 'Sent', 'Paid', 'Overdue', 'Void'];
export const DEAL_TYPES = ['New business', 'Renewal', 'Upsell'];
export const ACTIVITY_TYPES = ['Call', 'Email', 'Meeting', 'WhatsApp', 'Note', 'Task'];
export const WON = 'Closed Won';
export const LOST = 'Closed Lost';

/* ---------------------------------------------------------------- entities */

export const ENTITIES: IEntityDef[] = [
  {
    key: 'leads', listTitle: 'AppSource Leads', singular: 'Lead', plural: 'Leads', icon: 'People', area: 'Sales', plan: 'standard',
    description: 'Prospects from AppSource, web forms, events and manual entry', activityName: 'Lead', boardField: 'status',
    searchFields: ['title', 'email', 'company', 'phone', 'country', 'offer'],
    fields: [
      title('Name', { readOnly: true, required: false }),
      f('firstName', 'First name', 'text', { section: 'General', required: true }),
      f('lastName', 'Last name', 'text', { section: 'General' }),
      f('email', 'Email', 'email', { section: 'General' }),
      f('phone', 'Phone', 'phone', { section: 'General' }),
      f('whatsapp', 'WhatsApp ID', 'text', { hidden: true }),
      f('company', 'Company', 'text', { section: 'General', grid: true }),
      f('jobTitle', 'Job title', 'text', { section: 'General' }),
      f('country', 'Country', 'text', { section: 'General', filter: true }),
      f('industry', 'Industry', 'choice', { choicesFrom: 'industries', section: 'General', filter: true }),
      f('website', 'Website', 'url', { section: 'General' }),
      f('leadSource', 'Lead source', 'choice', { choicesFrom: 'leadSources', section: 'Qualification', grid: true, filter: true }),
      f('status', 'Status', 'choice', { choices: LEAD_STATUSES, defaultValue: 'New', section: 'Qualification', grid: true, filter: true }),
      f('rating', 'Rating', 'choice', { choices: ['Hot', 'Warm', 'Cold'], defaultValue: 'Warm', section: 'Qualification', grid: true, filter: true }),
      f('leadScore', 'Score', 'number', { readOnly: true, section: 'Qualification', grid: true }),
      { ...owner('LeadOwner'), section: 'Qualification' },
      f('nextFollowUp', 'Next follow-up', 'datetime', { section: 'Qualification', grid: true }),
      f('offer', 'Product / offer', 'text', { section: 'Qualification' }),
      f('actionCode', 'Marketplace action', 'text', { section: 'Qualification', readOnly: true }),
      ref('campaignRef', 'Campaign', 'campaigns'),
      ref('accountRef', 'Converted account', 'accounts', { readOnly: true }),
      ref('contactRef', 'Converted contact', 'contacts', { readOnly: true }),
      ref('opportunityRef', 'Converted opportunity', 'opportunities', { readOnly: true }),
      { ...tags, section: 'Qualification' },
      f('description', 'Description', 'note', { section: 'Details' }),
      f('rawPayload', 'Raw marketplace payload', 'note', { hidden: true })
    ]
  },
  {
    key: 'accounts', listTitle: 'CRM Accounts', singular: 'Account', plural: 'Accounts', icon: 'CityNext', area: 'Sales', plan: 'standard',
    description: 'Companies you sell to, partner with or compete against', activityName: 'Account', boardField: 'accountType',
    searchFields: ['title', 'website', 'email', 'phone', 'city', 'country'],
    related: [
      { entity: 'contacts', field: 'accountRef', label: 'Contacts' },
      { entity: 'opportunities', field: 'accountRef', label: 'Opportunities' },
      { entity: 'quotes', field: 'accountRef', label: 'Quotes' },
      { entity: 'subscriptions', field: 'accountRef', label: 'Subscriptions' },
      { entity: 'invoices', field: 'accountRef', label: 'Invoices' },
      { entity: 'cases', field: 'accountRef', label: 'Cases' },
      { entity: 'accounts', field: 'parentAccountRef', label: 'Subsidiaries' }
    ],
    fields: [
      title('Account name'),
      f('accountType', 'Type', 'choice', { choices: ['Prospect', 'Customer', 'Partner', 'Competitor', 'Other'], defaultValue: 'Prospect', section: 'General', grid: true, filter: true }),
      f('industry', 'Industry', 'choice', { choicesFrom: 'industries', section: 'General', grid: true, filter: true }),
      f('website', 'Website', 'url', { section: 'General' }),
      f('email', 'Email', 'email', { section: 'General' }),
      f('phone', 'Phone', 'phone', { section: 'General', grid: true }),
      f('employees', 'Employees', 'number', { section: 'General' }),
      f('annualRevenue', 'Annual revenue', 'currency', { section: 'General' }),
      owner('AccountOwner'),
      f('street', 'Street', 'text', { section: 'Address' }),
      f('city', 'City', 'text', { section: 'Address', grid: true }),
      f('state', 'State / region', 'text', { section: 'Address' }),
      f('postalCode', 'Postal code', 'text', { section: 'Address' }),
      f('country', 'Country', 'text', { section: 'Address', grid: true, filter: true }),
      ref('parentAccountRef', 'Parent account', 'accounts'),
      tags,
      f('description', 'Description', 'note', { section: 'Details' })
    ]
  },
  {
    key: 'contacts', listTitle: 'CRM Contacts', singular: 'Contact', plural: 'Contacts', icon: 'Contact', area: 'Sales', plan: 'standard',
    description: 'People at your accounts', activityName: 'Contact',
    searchFields: ['title', 'email', 'phone', 'mobile', 'jobTitle'],
    related: [
      { entity: 'opportunities', field: 'contactRef', label: 'Opportunities' },
      { entity: 'quotes', field: 'contactRef', label: 'Quotes' },
      { entity: 'cases', field: 'contactRef', label: 'Cases' }
    ],
    fields: [
      title('Full name', { readOnly: true, required: false }),
      f('firstName', 'First name', 'text', { section: 'General', required: true }),
      f('lastName', 'Last name', 'text', { section: 'General', required: true }),
      f('email', 'Email', 'email', { section: 'General', grid: true }),
      f('phone', 'Phone', 'phone', { section: 'General', grid: true }),
      f('mobile', 'Mobile', 'phone', { section: 'General' }),
      f('whatsapp', 'WhatsApp ID', 'text', { hidden: true }),
      f('jobTitle', 'Job title', 'text', { section: 'General', grid: true }),
      ref('accountRef', 'Account', 'accounts', { section: 'General', grid: true, filter: true }),
      owner('ContactOwner'),
      f('linkedIn', 'LinkedIn', 'url', { section: 'General' }),
      f('country', 'Country', 'text', { section: 'General', filter: true }),
      f('preferredContact', 'Preferred contact method', 'choice', { choices: ['Email', 'Phone', 'Teams', 'Any'], section: 'Privacy' }),
      f('marketingConsent', 'Marketing consent (opt-in)', 'boolean', { section: 'Privacy', filter: true }),
      f('consentDate', 'Consent date', 'date', { section: 'Privacy' }),
      { ...tags, section: 'Privacy' },
      f('description', 'Notes', 'note', { section: 'Details' })
    ]
  },
  {
    key: 'opportunities', listTitle: 'CRM Opportunities', singular: 'Opportunity', plural: 'Opportunities', icon: 'Money', area: 'Sales', plan: 'standard',
    description: 'Deals in your sales pipeline', activityName: 'Opportunity', boardField: 'stage',
    searchFields: ['title', 'company', 'contactName', 'email', 'nextStep'],
    related: [{ entity: 'quotes', field: 'opportunityRef', label: 'Quotes' }, { entity: 'subscriptions', field: 'opportunityRef', label: 'Subscriptions' }],
    fields: [
      title('Deal name'),
      ref('accountRef', 'Account', 'accounts', { section: 'General', grid: true }),
      ref('contactRef', 'Primary contact', 'contacts', { section: 'General' }),
      f('stage', 'Stage', 'choice', { choicesFrom: 'stages', defaultValue: 'Prospecting', section: 'General', grid: true, filter: true }),
      f('amount', 'Amount', 'currency', { section: 'General', grid: true }),
      f('probability', 'Probability', 'percent', { section: 'General', grid: true }),
      f('forecastCategory', 'Forecast category', 'choice', { choices: FORECAST_CATEGORIES, section: 'General', filter: true }),
      f('dealType', 'Deal type', 'choice', { choices: DEAL_TYPES, defaultValue: 'New business', section: 'General', filter: true }),
      f('expectedCloseDate', 'Expected close', 'date', { section: 'General', grid: true }),
      f('actualCloseDate', 'Actual close', 'date', { section: 'General' }),
      { ...owner('DealOwner') },
      f('nextStep', 'Next step', 'text', { section: 'General' }),
      f('licenses', 'Licenses / seats', 'number', { section: 'Details' }),
      f('competitors', 'Competitors', 'text', { section: 'Details' }),
      f('lostReason', 'Lost reason', 'choice', { choicesFrom: 'lostReasons', section: 'Details', filter: true }),
      f('company', 'Company (no account)', 'text', { section: 'Details' }),
      f('contactName', 'Contact name (no contact)', 'text', { section: 'Details' }),
      f('email', 'Contact email', 'email', { section: 'Details' }),
      f('phone', 'Contact phone', 'phone', { section: 'Details' }),
      ref('leadRef', 'Source lead', 'leads'),
      ref('demoRef', 'Demo', 'demos'),
      ref('campaignRef', 'Campaign', 'campaigns'),
      ref('subscriptionRef', 'Subscription (renewal)', 'subscriptions'),
      { ...tags, section: 'Details' },
      f('notes', 'Notes', 'note', { section: 'Details' })
    ]
  },
  {
    key: 'quotes', listTitle: 'CRM Quotes', singular: 'Quote', plural: 'Quotes', icon: 'Document', area: 'Sales', plan: 'standard',
    description: 'Priced proposals with line items, printable as PDF', activityName: 'Quote', boardField: 'quoteStatus',
    searchFields: ['title', 'quoteNumber'],
    related: [{ entity: 'invoices', field: 'quoteRef', label: 'Invoices' }, { entity: 'subscriptions', field: 'quoteRef', label: 'Subscriptions' }],
    fields: [
      f('quoteNumber', 'Quote #', 'text', { readOnly: true, section: 'General', grid: true }),
      title('Quote name'),
      ref('opportunityRef', 'Opportunity', 'opportunities', { section: 'General', grid: true }),
      ref('accountRef', 'Account', 'accounts', { section: 'General', grid: true }),
      ref('contactRef', 'Contact', 'contacts', { section: 'General' }),
      f('quoteStatus', 'Status', 'choice', { choices: QUOTE_STATUSES, defaultValue: 'Draft', section: 'General', grid: true, filter: true }),
      f('validUntil', 'Valid until', 'date', { section: 'General', grid: true }),
      f('approvalStatus', 'Approval', 'choice', { choices: ['Pending', 'Approved', 'Rejected'], section: 'General', grid: true, filter: true, readOnly: true }),
      f('approvalNote', 'Approver comments', 'note', { section: 'General', readOnly: true }),
      f('signatureStatus', 'E-signature', 'choice', { choices: ['Sent', 'Viewed', 'Signed', 'Declined'], section: 'Signature', readOnly: true, filter: true }),
      f('signatureEmail', 'Sent for signature to', 'email', { section: 'Signature', readOnly: true }),
      f('signatureName', 'Signed by', 'text', { section: 'Signature', readOnly: true }),
      f('signatureTitle', 'Signer job title', 'text', { section: 'Signature', readOnly: true }),
      f('signedDate', 'Signed on', 'datetime', { section: 'Signature', readOnly: true }),
      f('signatureIp', 'Signer IP address', 'text', { section: 'Signature', readOnly: true }),
      f('signatureId', 'Signature reference', 'text', { section: 'Signature', readOnly: true }),
      f('signatureNote', 'Signer comments', 'note', { section: 'Signature', readOnly: true }),
      f('signatureToken', 'Signing token', 'text', { hidden: true }),
      f('signedSnapshot', 'Signed content', 'note', { hidden: true }),
      f('signatureProcessed', 'Signature processed', 'boolean', { hidden: true }),
      f('currency', 'Currency', 'text', { hidden: true }),
      owner('QuoteOwner'),
      f('lineItems', 'Line items', 'json', { hidden: true }),
      f('subtotal', 'Subtotal', 'currency', { readOnly: true, hidden: true }),
      f('discountTotal', 'Discount', 'currency', { readOnly: true, hidden: true }),
      f('taxTotal', 'Tax', 'currency', { readOnly: true, hidden: true }),
      f('grandTotal', 'Total', 'currency', { readOnly: true, hidden: true, grid: true }),
      f('terms', 'Terms & conditions', 'note', { section: 'Terms' })
    ]
  },
  {
    key: 'products', listTitle: 'CRM Products', singular: 'Product', plural: 'Products', icon: 'Product', area: 'Catalog', plan: 'standard',
    description: 'Your price list, used on quotes', activityName: 'Product',
    searchFields: ['title', 'sku', 'category'],
    fields: [
      title('Product name'),
      f('sku', 'SKU', 'text', { section: 'General', grid: true }),
      f('category', 'Category', 'text', { section: 'General', grid: true, filter: true }),
      f('unitPrice', 'Unit price', 'currency', { section: 'General', grid: true, required: true }),
      f('billingCycle', 'Billing', 'choice', { choices: ['One-time', 'Monthly', 'Annual'], defaultValue: 'Annual', section: 'General', grid: true, filter: true }),
      f('taxRate', 'Tax rate', 'percent', { section: 'General' }),
      f('active', 'Active', 'boolean', { defaultValue: true, section: 'General', grid: true, filter: true }),
      f('description', 'Description', 'note', { section: 'Details' })
    ]
  },
  {
    key: 'demos', listTitle: 'CRM Demo Requests', singular: 'Demo request', plural: 'Demo requests', icon: 'Presentation', area: 'Marketing', plan: 'standard',
    description: 'Product demo requests and scheduled demos', activityName: 'Demo', boardField: 'status',
    searchFields: ['title', 'contactName', 'email', 'company', 'notes'],
    fields: [
      title('Demo title'),
      f('contactName', 'Contact name', 'text', { section: 'General', grid: true, required: true }),
      f('email', 'Email', 'email', { section: 'General' }),
      f('phone', 'Phone', 'phone', { section: 'General' }),
      f('company', 'Company', 'text', { section: 'General', grid: true }),
      { ...f('status', 'Status', 'choice', { choices: DEMO_STATUSES, defaultValue: 'Requested', section: 'General', grid: true, filter: true }), column: 'DemoStatus' },
      f('outcome', 'Outcome', 'choice', { choices: ['Interested', 'Follow-up needed', 'Not interested'], section: 'General', grid: true, filter: true }),
      f('preferredDate', 'Preferred date', 'datetime', { section: 'Scheduling', grid: true }),
      f('scheduledDate', 'Scheduled date', 'datetime', { section: 'Scheduling', grid: true }),
      f('durationMinutes', 'Duration (minutes)', 'number', { section: 'Scheduling', defaultValue: 30 }),
      { key: 'owner', column: 'Presenter', label: 'Presenter', type: 'user', section: 'Scheduling', grid: true, filter: true },
      f('meetingLink', 'Meeting link', 'url', { section: 'Scheduling' }),
      ref('leadRef', 'Lead', 'leads'),
      ref('accountRef', 'Account', 'accounts'),
      ref('contactRef', 'Contact', 'contacts'),
      f('notes', 'Notes / requirements', 'note', { section: 'Details' })
    ]
  },
  {
    key: 'cases', listTitle: 'CRM Cases', singular: 'Case', plural: 'Cases', icon: 'Headset', area: 'Service', plan: 'enterprise',
    description: 'Customer support tickets with SLA tracking', activityName: 'Case', boardField: 'caseStatus',
    searchFields: ['title', 'caseNumber', 'description', 'resolution'],
    fields: [
      f('caseNumber', 'Case #', 'text', { readOnly: true, section: 'General', grid: true }),
      title('Subject'),
      ref('accountRef', 'Account', 'accounts', { section: 'General', grid: true }),
      ref('contactRef', 'Contact', 'contacts', { section: 'General' }),
      f('priority', 'Priority', 'choice', { choices: CASE_PRIORITIES, defaultValue: 'Normal', section: 'General', grid: true, filter: true }),
      f('caseStatus', 'Status', 'choice', { choices: CASE_STATUSES, defaultValue: 'New', section: 'General', grid: true, filter: true }),
      f('category', 'Category', 'choice', { choicesFrom: 'caseCategories', section: 'General', filter: true }),
      f('origin', 'Origin', 'choice', { choices: ['Email', 'Phone', 'Web', 'Teams', 'Marketplace', 'Other'], section: 'General', filter: true }),
      owner('CaseOwner'),
      f('slaDue', 'SLA due', 'datetime', { readOnly: true, section: 'General', grid: true }),
      f('firstResponseDue', 'First response due', 'datetime', { readOnly: true, section: 'SLA' }),
      f('firstRespondedAt', 'First responded', 'datetime', { readOnly: true, section: 'SLA' }),
      f('slaPausedAt', 'SLA paused since', 'datetime', { readOnly: true, section: 'SLA' }),
      f('requesterEmail', 'Requester email', 'email', { section: 'General' }),
      f('resolvedDate', 'Resolved on', 'datetime', { readOnly: true, section: 'Resolution' }),
      f('description', 'Description', 'note', { section: 'Details' }),
      f('resolution', 'Resolution', 'note', { section: 'Resolution' })
    ]
  },
  {
    key: 'campaigns', listTitle: 'CRM Campaigns', singular: 'Campaign', plural: 'Campaigns', icon: 'Megaphone', area: 'Marketing', plan: 'enterprise',
    description: 'Marketing campaigns with lead attribution and ROI', activityName: 'Campaign', boardField: 'campaignStatus',
    searchFields: ['title', 'description'],
    related: [
      { entity: 'leads', field: 'campaignRef', label: 'Leads' },
      { entity: 'opportunities', field: 'campaignRef', label: 'Opportunities' }
    ],
    fields: [
      title('Campaign name'),
      f('campaignType', 'Type', 'choice', { choices: ['Email', 'Event', 'Webinar', 'Paid ads', 'Content', 'Partner', 'Marketplace', 'Other'], section: 'General', grid: true, filter: true }),
      f('campaignStatus', 'Status', 'choice', { choices: ['Planned', 'Active', 'Completed', 'Cancelled'], defaultValue: 'Planned', section: 'General', grid: true, filter: true }),
      f('startDate', 'Start', 'date', { section: 'General', grid: true }),
      f('endDate', 'End', 'date', { section: 'General', grid: true }),
      f('budget', 'Budget', 'currency', { section: 'Budget', grid: true }),
      f('actualCost', 'Actual cost', 'currency', { section: 'Budget' }),
      f('expectedRevenue', 'Expected revenue', 'currency', { section: 'Budget' }),
      owner('CampaignOwner'),
      f('description', 'Description', 'note', { section: 'Details' })
    ]
  },
  {
    key: 'subscriptions', listTitle: 'CRM Subscriptions', singular: 'Subscription', plural: 'Subscriptions', icon: 'RecurringEvent', area: 'Sales', plan: 'standard',
    description: 'Licences and recurring revenue with renewal dates', activityName: 'Subscription', boardField: 'subStatus',
    searchFields: ['title', 'licenseKey', 'notes'],
    related: [
      { entity: 'invoices', field: 'subscriptionRef', label: 'Invoices' },
      { entity: 'opportunities', field: 'subscriptionRef', label: 'Renewals' }
    ],
    fields: [
      title('Subscription'),
      ref('accountRef', 'Account', 'accounts', { section: 'General', grid: true, filter: true, required: true }),
      ref('productRef', 'Product', 'products', { section: 'General', grid: true }),
      f('subStatus', 'Status', 'choice', { choices: SUBSCRIPTION_STATUSES, defaultValue: 'Active', section: 'General', grid: true, filter: true }),
      f('billingCycle', 'Billing', 'choice', { choices: ['Monthly', 'Annual'], defaultValue: 'Annual', section: 'General', filter: true }),
      f('quantity', 'Seats / quantity', 'number', { section: 'General', defaultValue: 1 }),
      f('unitPrice', 'Price per seat per period', 'currency', { section: 'General' }),
      f('mrr', 'MRR', 'currency', { section: 'General', readOnly: true, grid: true }),
      f('arr', 'ARR', 'currency', { section: 'General', readOnly: true }),
      f('startDate', 'Start', 'date', { section: 'Term' }),
      f('renewalDate', 'Renewal date', 'date', { section: 'Term', grid: true }),
      f('endDate', 'Ended on', 'date', { section: 'Term' }),
      f('autoRenew', 'Auto-renew', 'boolean', { defaultValue: true, section: 'Term', filter: true }),
      owner('SubscriptionOwner'),
      f('licenseKey', 'Licence key', 'note', { section: 'Licence' }),
      ref('contactRef', 'Billing contact', 'contacts'),
      ref('opportunityRef', 'Opportunity', 'opportunities'),
      ref('quoteRef', 'Quote', 'quotes'),
      f('notes', 'Notes', 'note', { section: 'Details' })
    ]
  },
  {
    key: 'invoices', listTitle: 'CRM Invoices', singular: 'Invoice', plural: 'Invoices', icon: 'Money', area: 'Sales', plan: 'standard',
    description: 'Invoices with payment links and accounting export', activityName: 'Invoice', boardField: 'invoiceStatus',
    searchFields: ['title', 'invoiceNumber', 'externalRef'],
    fields: [
      f('invoiceNumber', 'Invoice #', 'text', { readOnly: true, section: 'General', grid: true }),
      title('Invoice title'),
      ref('accountRef', 'Account', 'accounts', { section: 'General', grid: true, filter: true, required: true }),
      ref('contactRef', 'Billing contact', 'contacts', { section: 'General' }),
      f('invoiceStatus', 'Status', 'choice', { choices: INVOICE_STATUSES, defaultValue: 'Draft', section: 'General', grid: true, filter: true }),
      f('issueDate', 'Issue date', 'date', { section: 'General', grid: true }),
      f('dueDate', 'Due date', 'date', { section: 'General', grid: true }),
      owner('InvoiceOwner'),
      f('lineItems', 'Line items', 'json', { hidden: true }),
      f('subtotal', 'Subtotal', 'currency', { readOnly: true, hidden: true }),
      f('discountTotal', 'Discount', 'currency', { readOnly: true, hidden: true }),
      f('taxTotal', 'Tax', 'currency', { readOnly: true, hidden: true }),
      f('grandTotal', 'Total', 'currency', { readOnly: true, hidden: true, grid: true }),
      f('currency', 'Currency', 'text', { hidden: true }),
      f('amountPaid', 'Amount paid', 'currency', { section: 'Payment' }),
      f('paidDate', 'Paid on', 'date', { section: 'Payment' }),
      f('paymentLink', 'Payment link', 'url', { section: 'Payment' }),
      f('paymentProvider', 'Payment via', 'choice', { choices: ['Razorpay', 'Stripe', 'Bank transfer', 'Other'], section: 'Payment', filter: true }),
      f('paymentLinkStatus', 'Payment link status', 'choice', { choices: ['Requested', 'Created', 'Failed'], section: 'Payment', readOnly: true }),
      f('externalPaymentId', 'Payment reference', 'text', { section: 'Payment', readOnly: true }),
      f('externalRef', 'Accounting reference', 'text', { section: 'Payment' }),
      ref('quoteRef', 'Quote', 'quotes'),
      ref('subscriptionRef', 'Subscription', 'subscriptions'),
      ref('opportunityRef', 'Opportunity', 'opportunities'),
      f('terms', 'Notes & payment terms', 'note', { section: 'Terms' })
    ]
  },
  {
    key: 'activities', listTitle: 'CRM Activities', singular: 'Activity', plural: 'Activities', icon: 'TaskManager', area: 'Work', plan: 'standard',
    description: 'Calls, emails, meetings, notes and tasks', activityName: 'Activity',
    searchFields: ['title', 'details'],
    fields: [
      title('Subject'),
      f('activityType', 'Type', 'choice', { choices: ACTIVITY_TYPES, defaultValue: 'Note', section: 'General', grid: true, filter: true }),
      f('relatedType', 'Related to', 'text', { section: 'General', grid: true, readOnly: true }),
      f('relatedId', 'Related ID', 'number', { section: 'General', readOnly: true }),
      f('dueDate', 'Due', 'datetime', { section: 'General', grid: true }),
      f('completed', 'Completed', 'boolean', { section: 'General', grid: true, filter: true }),
      { key: 'owner', column: 'TaskOwner', label: 'Assigned to', type: 'user', section: 'General', grid: true, filter: true },
      f('details', 'Details', 'note', { section: 'General' }),
      f('waStatus', 'WhatsApp status', 'choice', { choices: ['Queued', 'Sent', 'Delivered', 'Read', 'Received', 'Failed'], section: 'General', readOnly: true }),
      f('waTo', 'WhatsApp number', 'text', { hidden: true }),
      f('waTemplate', 'WhatsApp template', 'text', { hidden: true }),
      f('waParams', 'WhatsApp template values', 'json', { hidden: true }),
      f('externalId', 'External ID', 'text', { hidden: true }),
      f('activitySource', 'Source', 'choice', { choices: ['CRM', 'Outlook', 'Teams', 'WhatsApp', 'Website', 'Flow'], hidden: true }),
      f('webLink', 'Open in Outlook', 'url', { section: 'General', readOnly: true })
    ]
  }
];

export const SYSTEM_LISTS: Record<SystemListKey, { listTitle: string; description: string; fields: IFieldDef[] }> = {
  settings: {
    listTitle: 'CRM Settings', description: 'Brillienta CRM configuration (do not edit by hand)',
    fields: [f('value', 'Value', 'note')]
  },
  audit: {
    listTitle: 'CRM Audit Log', description: 'Brillienta CRM change history',
    fields: [
      f('entityType', 'Entity', 'text'),
      f('recordId', 'Record ID', 'number'),
      f('recordTitle', 'Record', 'text'),
      f('changes', 'Changes', 'note')
    ]
  }
};

export function entityDef(key: EntityKey, defs: IEntityDef[] = ENTITIES): IEntityDef {
  return defs.find(e => e.key === key)!;
}

export function entityByActivityName(name: string, defs: IEntityDef[] = ENTITIES): IEntityDef | undefined {
  return defs.find(e => e.activityName === name);
}
