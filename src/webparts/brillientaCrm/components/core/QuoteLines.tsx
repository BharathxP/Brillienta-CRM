import * as React from 'react';
import { ComboBox, DefaultButton, IconButton, Text, TextField, Stack, IComboBoxOption, IComboBox } from '@fluentui/react';
import { IRecord } from '../../models/model';
import { Engine, ILineItem, quoteTotals } from '../../services/Engine';
import { useCrm, useEntities } from '../CrmContext';
import { fmtDate, fmtMoney, uid } from '../utils';

/** Free-form product picker that also opens its list when the text box is clicked. */
const ProductCombo: React.FC<{ options: IComboBoxOption[]; text: string; productId?: number; disabled?: boolean; onPick: (key?: string | number, text?: string) => void }> =
  ({ options, text, productId, disabled, onPick }) => {
    const ref = React.useRef<IComboBox>(null);
    return (
      <ComboBox componentRef={ref} allowFreeform autoComplete="on" options={options} text={text} selectedKey={productId ?? null} disabled={disabled}
        onClick={() => ref.current?.focus(true)} onChange={(_, o, __, v) => onPick(o?.key, v)} ariaLabel="Product" placeholder="Pick or type" useComboBoxAsMenuWidth={false} />
    );
  };

export const QuoteLines: React.FC<{ items: ILineItem[]; onChange: (items: ILineItem[]) => void; disabled?: boolean }> = ({ items, onChange, disabled }) => {
  const { currency } = useCrm();
  const { rows } = useEntities('products');
  const products = rows('products').filter(p => p.active !== false);
  const opts: IComboBoxOption[] = products.map(p => ({ key: p.id!, text: `${p.title}${p.sku ? ` (${p.sku})` : ''}` }));
  const set = (i: number, patch: Partial<ILineItem>): void => onChange(items.map((x, j) => (j === i ? { ...x, ...patch } : x)));
  const add = (): void => onChange([...items, { id: uid(), name: '', quantity: 1, unitPrice: 0, discountPct: 0, taxPct: 0 }]);
  const pick = (i: number, key?: string | number, text?: string): void => {
    const p = products.find(x => x.id === Number(key));
    if (p) {
      set(i, { productId: p.id, name: p.title || '', unitPrice: Number(p.unitPrice) || 0, taxPct: Number(p.taxRate) || 0, billing: p.billingCycle, description: p.description });
    } else if (text) {
      set(i, { productId: undefined, name: text });
    }
  };
  const t = quoteTotals(items);
  const money = (n: number): string => fmtMoney(n, currency).replace(/\.00$/, '');

  return (
    <div>
      <div className="crm-lines">
        <div className="crm-lines-head"><span>Product / service</span><span>Qty</span><span>Unit price</span><span>Disc %</span><span>Tax %</span><span>Total</span><span /></div>
        {items.map((it, i) => {
          const gross = it.quantity * it.unitPrice;
          const net = gross * (1 - it.discountPct / 100);
          return (
            <div key={it.id} className="crm-lines-row">
              <ProductCombo options={opts} text={it.name} productId={it.productId} disabled={disabled} onPick={(k, t) => pick(i, k, t)} />
              <TextField type="number" min={0} value={String(it.quantity)} disabled={disabled} ariaLabel="Quantity" onChange={(_, v) => set(i, { quantity: Number(v) || 0 })} />
              <TextField type="number" min={0} value={String(it.unitPrice)} disabled={disabled} ariaLabel="Unit price" onChange={(_, v) => set(i, { unitPrice: Number(v) || 0 })} />
              <TextField type="number" min={0} max={100} value={String(it.discountPct)} disabled={disabled} ariaLabel="Discount percent" onChange={(_, v) => set(i, { discountPct: Math.min(100, Number(v) || 0) })} />
              <TextField type="number" min={0} max={100} value={String(it.taxPct)} disabled={disabled} ariaLabel="Tax percent" onChange={(_, v) => set(i, { taxPct: Number(v) || 0 })} />
              <Text className="crm-num">{money(net)}</Text>
              {!disabled ? <IconButton iconProps={{ iconName: 'Delete' }} ariaLabel="Remove line" onClick={() => onChange(items.filter((_, j) => j !== i))} /> : <span />}
            </div>
          );
        })}
      </div>
      {!disabled && <DefaultButton iconProps={{ iconName: 'Add' }} text="Add line" onClick={add} style={{ marginTop: 8 }} />}
      <Stack horizontalAlign="end" tokens={{ childrenGap: 2 }} style={{ marginTop: 12 }}>
        <Text>Subtotal: <b>{money(t.subtotal)}</b></Text>
        {t.discountTotal > 0 && <Text>Discount: <b>−{money(t.discountTotal)}</b></Text>}
        <Text>Tax: <b>{money(t.taxTotal)}</b></Text>
        <Text variant="large">Total: <b>{money(t.grandTotal)}</b></Text>
      </Stack>
    </div>
  );
};

const esc = (s: unknown): string => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** Builds a branded, print-ready quote document. */
export function quoteHtml(q: IRecord, engine: Engine): string {
  const s = engine.settings;
  const isInv = !!q.invoiceNumber;
  const number = isInv ? q.invoiceNumber : q.quoteNumber;
  const c = s.company;
  const cur = s.currency;
  const money = (n: number): string => esc(fmtMoney(n, cur));
  const acc = engine.find('accounts', q.accountRef);
  const con = engine.find('contacts', q.contactRef);
  const items: ILineItem[] = q.lineItems || [];
  const t = quoteTotals(items);
  const logo = /^https:\/\//i.test(c.logoUrl) ? `<img src="${esc(c.logoUrl)}" alt="" style="max-height:56px;max-width:220px">` : '';
  const rows = items.map(i => {
    const net = i.quantity * i.unitPrice * (1 - i.discountPct / 100);
    return `<tr><td><b>${esc(i.name)}</b>${i.description ? `<div class="d">${esc(i.description)}</div>` : ''}${i.billing ? `<div class="d">${esc(i.billing)}</div>` : ''}</td>
      <td class="n">${esc(i.quantity)}</td><td class="n">${money(i.unitPrice)}</td><td class="n">${i.discountPct ? `${esc(i.discountPct)}%` : ''}</td>
      <td class="n">${i.taxPct ? `${esc(i.taxPct)}%` : ''}</td><td class="n">${money(net)}</td></tr>`;
  }).join('');
  return `<!doctype html><html><head><meta charset="utf-8"><title>${esc(number)} - ${esc(q.title)}</title>
<style>
 body{font-family:'Segoe UI',Arial,sans-serif;color:#222;margin:40px;font-size:13px}
 .top{display:flex;justify-content:space-between;align-items:flex-start;border-bottom:3px solid #0f6cbd;padding-bottom:16px}
 h1{font-size:26px;margin:0;color:#0f6cbd} .muted{color:#666} .grid{display:flex;gap:40px;margin:24px 0}
 table{width:100%;border-collapse:collapse;margin-top:8px} th{background:#f3f6fa;text-align:left;font-weight:600;padding:8px;border-bottom:1px solid #ccd}
 td{padding:8px;border-bottom:1px solid #eee;vertical-align:top} .n{text-align:right;white-space:nowrap} .d{color:#666;font-size:12px}
 .tot{margin-left:auto;width:300px;margin-top:16px} .tot div{display:flex;justify-content:space-between;padding:4px 0}
 .grand{font-size:16px;font-weight:700;border-top:2px solid #222;margin-top:4px;padding-top:8px!important}
 .terms{margin-top:32px;font-size:12px;color:#444;white-space:pre-wrap} .foot{margin-top:40px;text-align:center;color:#666}
 .sign{display:flex;gap:40px;margin-top:48px} .sign div{flex:1;border-top:1px solid #999;padding-top:6px;color:#666}
 @media print{body{margin:16mm} .noprint{display:none}}
</style></head><body>
<div class="noprint" style="text-align:right;margin-bottom:12px"><button onclick="window.print()">Print / Save as PDF</button></div>
<div class="top"><div>${logo}<div style="margin-top:6px"><b>${esc(c.name)}</b><br><span class="muted">${esc(c.address).replace(/\n/g, '<br>')}</span><br>
<span class="muted">${esc([c.email, c.phone, c.website].filter(Boolean).join(' · '))}</span>${c.taxId ? `<br><span class="muted">Tax ID: ${esc(c.taxId)}</span>` : ''}</div></div>
<div style="text-align:right"><h1>${isInv ? (q.invoiceStatus === 'Paid' ? 'INVOICE · PAID' : 'INVOICE') : 'QUOTE'}</h1><div><b>${esc(number)}</b></div>
<div class="muted">Date: ${esc(fmtDate(q.issueDate || q.created || new Date().toISOString()))}</div>
<div class="muted">${isInv ? `Due: ${esc(fmtDate(q.dueDate))}` : `Valid until: ${esc(fmtDate(q.validUntil))}`}</div></div></div>
<div class="grid"><div><div class="muted">${isInv ? 'Bill to' : 'Prepared for'}</div><b>${esc(acc?.title || '')}</b><br>${esc(con?.title || '')}${con?.email ? `<br>${esc(con.email)}` : ''}
${acc ? `<br><span class="muted">${esc([acc.street, acc.city, acc.state, acc.postalCode, acc.country].filter(Boolean).join(', '))}</span>` : ''}</div>
<div><div class="muted">Prepared by</div><b>${esc(q.owner?.title || engine.me?.title || '')}</b><br>${esc(q.owner?.email || engine.me?.email || '')}</div>
<div><div class="muted">Reference</div><b>${esc(q.title)}</b></div></div>
<table><thead><tr><th>Description</th><th class="n">Qty</th><th class="n">Unit price</th><th class="n">Discount</th><th class="n">Tax</th><th class="n">Amount</th></tr></thead><tbody>${rows}</tbody></table>
<div class="tot"><div><span>Subtotal</span><span>${money(t.subtotal)}</span></div>${t.discountTotal ? `<div><span>Discount</span><span>−${money(t.discountTotal)}</span></div>` : ''}
<div><span>Tax</span><span>${money(t.taxTotal)}</span></div><div class="grand"><span>Total (${esc(cur)})</span><span>${money(t.grandTotal)}</span></div>
${isInv && q.amountPaid ? `<div><span>Paid</span><span>−${money(q.amountPaid)}</span></div><div class="grand"><span>Balance due</span><span>${money(Math.max(0, t.grandTotal - q.amountPaid))}</span></div>` : ''}</div>
${isInv && /^https:\/\//i.test(q.paymentLink || '') && q.invoiceStatus !== 'Paid' ? `<div class="terms"><b>Pay online:</b> <a href="${esc(q.paymentLink)}">${esc(q.paymentLink)}</a></div>` : ''}
${isInv && s.invoice.bankDetails ? `<div class="terms"><b>Bank transfer</b><br>${esc(s.invoice.bankDetails)}</div>` : ''}
${q.terms ? `<div class="terms"><b>${isInv ? 'Notes' : 'Terms &amp; conditions'}</b><br>${esc(q.terms)}</div>` : ''}
${isInv ? '' : (q.signatureName ? `<div class="sign"><div><b>${esc(q.signatureName)}</b>${q.signatureTitle ? `, ${esc(q.signatureTitle)}` : ''}<br>Accepted electronically ${esc(fmtDate(q.signedDate, true))}</div><div>Ref ${esc(q.signatureId || '')}</div></div>` : '<div class="sign"><div>Accepted by (name, signature)</div><div>Date</div></div>')}
<div class="foot">${esc(isInv ? s.invoice.footer : s.quote.footer)}</div></body></html>`;
}

export function printQuote(q: IRecord, engine: Engine): void {
  const doc = quoteHtml(q, engine);
  let w: Window | null = null;
  try {
    w = window.open('', '_blank');
  } catch {
    w = null;
  }
  if (w) {
    w.document.open();
    w.document.write(doc);
    w.document.close();
    return;
  }
  // Pop-ups blocked (e.g. Teams desktop): download the quote; open it and use Print / Save as PDF.
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([doc], { type: 'text/html' }));
  a.download = `${String(q.invoiceNumber || q.quoteNumber || 'quote').replace(/[^\w-]/g, '_')}.html`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}
