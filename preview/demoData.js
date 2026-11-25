// A realistic sample CRM for the preview and the store screenshots (window.__crm.seedDemo(), after "Set up CRM").
// Dates are relative to today, so the dashboard, forecast and renewals always look current.
import { seedItems } from './fakeSharePoint';

const DAY = 86400000;
const at = (days) => new Date(Date.now() + days * DAY).toISOString();
const ASHA = 7, BEN = 8, CHEN = 9, MAYA = 12, DANA = 10;

export function seedDemo() {
  const accounts = [
    ['Contoso Ltd', 'Customer', 'Technology', 'Seattle', 'United States', CHEN, 'https://contoso.com'],
    ['Fabrikam Inc', 'Customer', 'Manufacturing', 'Pune', 'India', MAYA, 'https://fabrikam.com'],
    ['Northwind Traders', 'Customer', 'Retail', 'London', 'United Kingdom', CHEN, 'https://northwindtraders.com'],
    ['Adatum Corporation', 'Prospect', 'Financial services', 'Frankfurt', 'Germany', MAYA, 'https://adatum.com'],
    ['Woodgrove Bank', 'Prospect', 'Financial services', 'Toronto', 'Canada', CHEN, 'https://woodgrovebank.com'],
    ['Tailspin Toys', 'Customer', 'Retail', 'Austin', 'United States', MAYA, 'https://tailspintoys.com'],
    ['Litware Health', 'Prospect', 'Healthcare', 'Bengaluru', 'India', CHEN, 'https://litwarehealth.com'],
    ['Proseware Education', 'Customer', 'Education', 'Melbourne', 'Australia', MAYA, 'https://proseware.com']
  ];
  const acc = seedItems('BrillCRM_Accounts', accounts.map(([Title, AccountType, Industry, City, Country, owner, Website]) => ({ Title, AccountType, Industry, City, Country, Website, AccountOwnerId: owner, Created: at(-200) })));
  const people = [
    ['Priya', 'Shah', 'CTO', 0], ['Tom', 'Lee', 'Head of HR', 0], ['Anita', 'Rao', 'COO', 1], ['James', 'Hill', 'IT Director', 2],
    ['Sofia', 'Becker', 'CFO', 3], ['Liam', 'Chen', 'VP Operations', 4], ['Emma', 'Wilson', 'Procurement Lead', 5], ['Ravi', 'Kumar', 'CIO', 6], ['Grace', 'Ng', 'Principal', 7]
  ];
  const con = seedItems('BrillCRM_Contacts', people.map(([FirstName, LastName, JobTitle, a]) => ({
    Title: `${FirstName} ${LastName}`, FirstName, LastName, JobTitle, AccountRef: acc[a], ContactOwnerId: accounts[a][5],
    Email: `${FirstName.toLowerCase()}@${accounts[a][6].replace('https://', '')}`, Phone: '+1 425 555 01' + String(10 + a), MarketingConsent: true
  })));
  const stage = { Prospecting: [10, 'Pipeline'], Qualification: [25, 'Pipeline'], Proposal: [50, 'Best case'], Negotiation: [75, 'Commit'], 'Closed Won': [100, 'Closed'] };
  const deals = [
    ['Contoso HR portal rollout', 0, 0, 'Negotiation', 48000, 12, CHEN, 'Pricing call with the CFO'],
    ['Fabrikam service desk', 1, 2, 'Proposal', 22500, 20, MAYA, 'Send revised proposal'],
    ['Northwind store analytics', 2, 3, 'Qualification', 31000, 35, CHEN, 'Discovery workshop'],
    ['Adatum risk dashboard', 3, 4, 'Prospecting', 18000, 50, MAYA, ''],
    ['Woodgrove onboarding suite', 4, 5, 'Proposal', 64000, 28, CHEN, 'Security review'],
    ['Tailspin loyalty CRM', 5, 6, 'Negotiation', 27500, 9, MAYA, 'Legal review of the contract'],
    ['Litware patient portal', 6, 7, 'Qualification', 41000, 42, CHEN, 'Demo for the clinical team'],
    ['Proseware campus licences', 7, 8, 'Prospecting', 15500, 60, MAYA, ''],
    ['Contoso analytics add-on', 0, 1, 'Proposal', 15000, 18, CHEN, 'Quote viewed by customer']
  ];
  const opp = seedItems('BrillCRM_Opportunities', deals.map(([Title, a, c, s, Amount, days, owner, NextStep]) => ({
    Title, AccountRef: acc[a], ContactRef: con[c], Stage: s, Amount, AmountBase: Amount, Currency: 'USD', Probability: stage[s][0], ForecastCategory: stage[s][1],
    ExpectedCloseDate: at(days), DealOwnerId: owner, NextStep, DealType: 'New business', Created: at(-40 - days)
  })));
  // Won deals over the last year (revenue chart, win rate, sales cycle)
  const won = [];
  for (let m = 1; m <= 11; m++) {
    won.push({ Title: `${accounts[m % 8][0]} expansion ${m}`, AccountRef: acc[m % 8], Stage: 'Closed Won', Probability: 100, ForecastCategory: 'Closed', Amount: 8000 + ((m * 3700) % 21000), AmountBase: 8000 + ((m * 3700) % 21000), Currency: 'USD', ExpectedCloseDate: at(-30 * m), ActualCloseDate: at(-30 * m), DealOwnerId: m % 2 ? CHEN : MAYA, DealType: m % 3 ? 'New business' : 'Upsell', Created: at(-30 * m - 45) });
  }
  won.push({ Title: 'Adatum pilot', AccountRef: acc[3], Stage: 'Closed Lost', Probability: 0, ForecastCategory: 'Omitted', Amount: 9000, AmountBase: 9000, LostReason: 'Price', ExpectedCloseDate: at(-50), ActualCloseDate: at(-50), DealOwnerId: MAYA, Created: at(-110) });
  seedItems('BrillCRM_Opportunities', won);
  const leads = [
    ['Arjun', 'Mehta', 'Northwind Traders', 'Event', 'Hot', 78, CHEN, 'India'], ['Nina', 'Patel', 'Fabrikam', 'Webinar', 'Warm', 55, MAYA, 'United States'],
    ['Omar', 'Haddad', 'Adatum', 'Referral', 'Hot', 82, CHEN, 'Germany'], ['Lena', 'Gomez', 'Woodgrove Bank', 'Website Demo Form', 'Warm', 61, MAYA, 'Canada'],
    ['Kenji', 'Sato', 'Tailspin Toys', 'AppSource', 'Hot', 74, CHEN, 'Japan'], ['Chloe', 'Martin', 'Litware Health', 'LinkedIn', 'Cold', 22, MAYA, 'France']
  ];
  seedItems('BrillCRM_Leads', leads.map(([FirstName, LastName, Company, LeadSource, Rating, LeadScore, owner, Country], i) => ({
    Title: `${FirstName} ${LastName}`, FirstName, LastName, Company, LeadSource, Rating, LeadScore, LeadOwnerId: owner, Country, Status: i < 2 ? 'Contacted' : 'New',
    Email: `${FirstName.toLowerCase()}@${Company.toLowerCase().replace(/[^a-z]/g, '')}.com`, JobTitle: ['IT Manager', 'Head of Sales', 'CTO', 'VP Finance', 'COO', 'HR Lead'][i], Created: at(-i * 3 - 1)
  })));
  const year = new Date().getFullYear();
  seedItems('BrillCRM_Cases', [
    { CaseNumber: `CS-${year}-0042`, Title: 'SSO login fails for new users', AccountRef: acc[0], ContactRef: con[0], Priority: 'High', CaseStatus: 'In progress', CaseOwnerId: DANA, SlaDue: at(-0.2), FirstResponseDue: at(-0.6), FirstRespondedAt: at(-0.7), Created: at(-1) },
    { CaseNumber: `CS-${year}-0045`, Title: 'Invoice PDF missing logo', AccountRef: acc[1], Priority: 'Normal', CaseStatus: 'New', CaseOwnerId: DANA, SlaDue: at(0.4), FirstResponseDue: at(0.1), Created: at(-0.2) },
    { CaseNumber: `CS-${year}-0047`, Title: 'Export to Excel times out', AccountRef: acc[2], Priority: 'Critical', CaseStatus: 'In progress', CaseOwnerId: DANA, SlaDue: at(1.5), FirstResponseDue: at(-0.4), FirstRespondedAt: at(-0.5), Created: at(-0.6) }
  ], DANA);
  seedItems('BrillCRM_Subscriptions', [
    { Title: 'HR Suite - 250 seats', AccountRef: acc[0], SubStatus: 'Active', BillingCycle: 'Annual', Quantity: 250, UnitPrice: 144, Mrr: 3000, Arr: 36000, StartDate: at(-347), RenewalDate: at(18), AutoRenew: true, SubscriptionOwnerId: CHEN, Currency: 'USD' },
    { Title: 'Service Desk - 40 seats', AccountRef: acc[1], SubStatus: 'Active', BillingCycle: 'Annual', Quantity: 40, UnitPrice: 240, Mrr: 800, Arr: 9600, StartDate: at(-330), RenewalDate: at(35), AutoRenew: false, SubscriptionOwnerId: MAYA, Currency: 'USD' },
    { Title: 'Analytics - 60 seats', AccountRef: acc[2], SubStatus: 'Active', BillingCycle: 'Annual', Quantity: 60, UnitPrice: 300, Mrr: 1500, Arr: 18000, StartDate: at(-120), RenewalDate: at(245), AutoRenew: true, SubscriptionOwnerId: CHEN, Currency: 'USD' },
    { Title: 'Campus licences', AccountRef: acc[7], SubStatus: 'Active', BillingCycle: 'Monthly', Quantity: 500, UnitPrice: 4, Mrr: 2000, Arr: 24000, StartDate: at(-200), RenewalDate: at(5), AutoRenew: true, SubscriptionOwnerId: MAYA, Currency: 'USD' }
  ]);
  seedItems('BrillCRM_Invoices', [
    { InvoiceNumber: `INV-${year}-0118`, Title: 'HR Suite annual licence', AccountRef: acc[1], InvoiceStatus: 'Overdue', IssueDate: at(-71), DueDate: at(-41), Subtotal: 9600, GrandTotal: 9600, AmountPaid: 0, Currency: 'USD', InvoiceOwnerId: MAYA },
    { InvoiceNumber: `INV-${year}-0131`, Title: 'Onboarding services', AccountRef: acc[0], InvoiceStatus: 'Sent', IssueDate: at(-25), DueDate: at(5), Subtotal: 12000, GrandTotal: 12000, AmountPaid: 4000, Currency: 'USD', InvoiceOwnerId: CHEN },
    { InvoiceNumber: `INV-${year}-0140`, Title: 'Analytics licence', AccountRef: acc[2], InvoiceStatus: 'Paid', IssueDate: at(-60), DueDate: at(-30), PaidDate: at(-33), Subtotal: 18000, GrandTotal: 18000, AmountPaid: 18000, Currency: 'USD', InvoiceOwnerId: CHEN }
  ]);
  const lines = [{ id: 'l1', name: 'HR Suite', description: 'Annual licence per user', quantity: 250, unitPrice: 144, discountPct: 10, taxPct: 0, billing: 'Annual' },
    { id: 'l2', name: 'Onboarding service', description: 'Implementation and training', quantity: 1, unitPrice: 6500, discountPct: 0, taxPct: 0, billing: 'One-time' }];
  seedItems('BrillCRM_Quotes', [{ QuoteNumber: `Q-${year}-0007`, Title: 'Contoso HR portal - 3 years', OpportunityRef: opp[0], AccountRef: acc[0], ContactRef: con[0], QuoteStatus: 'Sent', ValidUntil: at(21), Currency: 'USD', LineItems: JSON.stringify(lines), Subtotal: 42500, DiscountTotal: 3600, TaxTotal: 0, GrandTotal: 38900, QuoteOwnerId: CHEN, SignatureStatus: 'Viewed' }]);
  seedItems('BrillCRM_Activities', [
    { Title: 'Discovery call with Priya', ActivityType: 'Call', RelatedType: 'Opportunity', RelatedId: opp[0], DueDate: at(-6), Completed: true, TaskOwnerId: CHEN },
    { Title: 'Demo for the HR team', ActivityType: 'Meeting', RelatedType: 'Opportunity', RelatedId: opp[0], DueDate: at(-2), Completed: true, TaskOwnerId: CHEN },
    { Title: 'Send revised proposal', ActivityType: 'Task', RelatedType: 'Opportunity', RelatedId: opp[0], DueDate: at(1), Completed: false, TaskOwnerId: CHEN },
    { Title: 'Follow up on security questionnaire', ActivityType: 'Task', RelatedType: 'Opportunity', RelatedId: opp[4], DueDate: at(-1), Completed: false, TaskOwnerId: CHEN },
    { Title: 'Renewal call with Anita', ActivityType: 'Task', RelatedType: 'Account', RelatedId: acc[1], DueDate: at(3), Completed: false, TaskOwnerId: MAYA }
  ]);
  seedItems('BrillCRM_Products', [
    { Title: 'HR Suite', Sku: 'HR-01', Category: 'Software', UnitPrice: 144, BillingCycle: 'Annual', Active: true },
    { Title: 'Analytics add-on', Sku: 'AN-01', Category: 'Software', UnitPrice: 300, BillingCycle: 'Annual', Active: true },
    { Title: 'Onboarding service', Sku: 'SRV-01', Category: 'Services', UnitPrice: 6500, BillingCycle: 'One-time', Active: true }
  ]);
}
