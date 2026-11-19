// Invoice exports for accounting systems. Files are created in the browser; nothing is sent anywhere.
// Each accounting product also offers a column-mapping step on import, so these follow their documented
// import templates as closely as possible and stay readable when a column needs to be remapped.
import { IRecord } from '../models/model';
import { Engine, ILineItem } from './Engine';
import { exportCsv } from '../components/utils';
import { downloadBlob } from './FlowPackage';

export type AccountingFormat = 'csv' | 'zoho' | 'quickbooks' | 'xero' | 'tally';

export const ACCOUNTING_FORMATS: { key: AccountingFormat; text: string }[] = [
  { key: 'csv', text: 'Invoice lines (generic CSV)' },
  { key: 'zoho', text: 'Zoho Books (CSV)' },
  { key: 'quickbooks', text: 'QuickBooks Online (CSV)' },
  { key: 'xero', text: 'Xero (CSV)' },
  { key: 'tally', text: 'Tally Prime / ERP 9 (XML)' }
];

const iso = (v?: string): string => (v ? new Date(v).toISOString().slice(0, 10) : '');
const r2 = (n: number): number => Math.round(n * 100) / 100;
const net = (i: ILineItem): number => r2((Number(i.quantity) || 0) * (Number(i.unitPrice) || 0) * (1 - (Number(i.discountPct) || 0) / 100));
const xml = (s: unknown): string => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

interface ILine { inv: IRecord; item: ILineItem; customer: string; email: string }

function lines(engine: Engine, invoices: IRecord[]): ILine[] {
  const out: ILine[] = [];
  invoices.filter(inv => inv.invoiceStatus !== 'Void').forEach(inv => {
    const customer = engine.display('accounts', inv.accountRef) || 'Customer';
    const email = engine.find('contacts', inv.contactRef)?.email || '';
    ((inv.lineItems || []) as ILineItem[]).forEach(item => out.push({ inv, item, customer, email }));
  });
  return out;
}

const STATUS_ZOHO: Record<string, string> = { Draft: 'Draft', Sent: 'Sent', Overdue: 'Overdue', Paid: 'Paid' };

/** Builds the rows (or XML) for one format. Exported for tests. */
export function accountingRows(engine: Engine, invoices: IRecord[], format: AccountingFormat): Record<string, unknown>[] | string {
  const s = engine.settings;
  const cur = s.currency;
  const ls = lines(engine, invoices);
  switch (format) {
    case 'zoho':
      return ls.map(({ inv, item, customer }) => ({
        'Invoice Date': iso(inv.issueDate), 'Invoice Number': inv.invoiceNumber, 'Invoice Status': STATUS_ZOHO[inv.invoiceStatus] || 'Draft',
        'Customer Name': customer, 'Due Date': iso(inv.dueDate), 'Currency Code': cur, 'Item Name': item.name, 'Item Desc': item.description || '',
        Quantity: item.quantity, 'Item Price': item.unitPrice, 'Discount(%)': item.discountPct || 0, 'Item Tax %': item.taxPct || 0,
        'Reference#': inv.title || '', Notes: inv.terms || ''
      }));
    case 'quickbooks':
      return ls.map(({ inv, item, customer }) => ({
        InvoiceNo: inv.invoiceNumber, Customer: customer, InvoiceDate: iso(inv.issueDate), DueDate: iso(inv.dueDate),
        Terms: `Net ${s.invoice.paymentTermsDays}`, Memo: inv.title || '', 'Item(Product/Service)': item.name, ItemDescription: item.description || '',
        ItemQuantity: item.quantity, ItemRate: r2(item.unitPrice * (1 - (item.discountPct || 0) / 100)), ItemAmount: net(item),
        Taxable: item.taxPct ? 'Y' : 'N', TaxRate: item.taxPct ? `${item.taxPct}%` : '', Currency: cur
      }));
    case 'xero':
      return ls.map(({ inv, item, customer, email }) => ({
        ContactName: customer, EmailAddress: email, InvoiceNumber: inv.invoiceNumber, Reference: inv.title || '',
        InvoiceDate: iso(inv.issueDate), DueDate: iso(inv.dueDate), Description: [item.name, item.description].filter(Boolean).join(' - '),
        Quantity: item.quantity, UnitAmount: item.unitPrice, Discount: item.discountPct || '', AccountCode: s.accounting.xeroSalesAccount,
        TaxType: s.accounting.xeroTaxType, Currency: cur
      }));
    case 'tally': {
      const a = s.accounting;
      const vouchers = invoices.filter(inv => inv.invoiceStatus !== 'Void').map(inv => {
        const items = (inv.lineItems || []) as ILineItem[];
        const sales = r2(items.reduce((t, i) => t + net(i), 0));
        const tax = r2(Number(inv.taxTotal) || 0);
        const total = r2(sales + tax);
        const party = engine.display('accounts', inv.accountRef) || 'Customer';
        const entry = (ledger: string, positive: boolean, amount: number): string =>
          `<ALLLEDGERENTRIES.LIST><LEDGERNAME>${xml(ledger)}</LEDGERNAME><ISDEEMEDPOSITIVE>${positive ? 'Yes' : 'No'}</ISDEEMEDPOSITIVE><AMOUNT>${amount.toFixed(2)}</AMOUNT></ALLLEDGERENTRIES.LIST>`;
        return `<TALLYMESSAGE xmlns:UDF="TallyUDF"><VOUCHER VCHTYPE="Sales" ACTION="Create" OBJVIEW="Accounting Voucher View">
<DATE>${iso(inv.issueDate).replace(/-/g, '')}</DATE><VOUCHERTYPENAME>Sales</VOUCHERTYPENAME><VOUCHERNUMBER>${xml(inv.invoiceNumber)}</VOUCHERNUMBER>
<REFERENCE>${xml(inv.invoiceNumber)}</REFERENCE><PARTYLEDGERNAME>${xml(party)}</PARTYLEDGERNAME><PERSISTEDVIEW>Accounting Voucher View</PERSISTEDVIEW>
<NARRATION>${xml(inv.title || '')}</NARRATION>
${entry(party, true, -total)}
${entry(a.tallySalesLedger || 'Sales', false, sales)}
${tax ? entry(a.tallyTaxLedger || 'Output GST', false, tax) : ''}
</VOUCHER></TALLYMESSAGE>`;
      });
      return `<?xml version="1.0" encoding="UTF-8"?>
<ENVELOPE><HEADER><TALLYREQUEST>Import Data</TALLYREQUEST></HEADER><BODY><IMPORTDATA>
<REQUESTDESC><REPORTNAME>Vouchers</REPORTNAME><STATICVARIABLES><SVCURRENTCOMPANY>${xml(s.company.name)}</SVCURRENTCOMPANY></STATICVARIABLES></REQUESTDESC>
<REQUESTDATA>
${vouchers.join('\n')}
</REQUESTDATA></IMPORTDATA></BODY></ENVELOPE>`;
    }
    default:
      return ls.map(({ inv, item, customer, email }) => ({
        'Invoice #': inv.invoiceNumber, Status: inv.invoiceStatus, Customer: customer, Email: email, 'Issue date': iso(inv.issueDate), 'Due date': iso(inv.dueDate),
        Item: item.name, Description: item.description || '', Quantity: item.quantity, 'Unit price': item.unitPrice, 'Discount %': item.discountPct || 0,
        'Tax %': item.taxPct || 0, 'Net amount': net(item), 'Tax amount': r2(net(item) * (item.taxPct || 0) / 100), 'Invoice total': inv.grandTotal,
        'Amount paid': inv.amountPaid || 0, 'Paid on': iso(inv.paidDate), Currency: cur
      }));
  }
}

export function exportInvoices(engine: Engine, invoices: IRecord[], format: AccountingFormat): number {
  const out = accountingRows(engine, invoices, format);
  const stamp = new Date().toISOString().slice(0, 10);
  if (typeof out === 'string') {
    downloadBlob(new Blob([out], { type: 'application/xml' }), `invoices-tally-${stamp}.xml`);
    return invoices.length;
  }
  exportCsv(`invoices-${format}-${stamp}.csv`, out);
  return out.length;
}
