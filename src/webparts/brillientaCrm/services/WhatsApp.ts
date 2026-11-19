// WhatsApp: click-to-chat links (free, works with any WhatsApp account) and queued sends for the
// WhatsApp Business Cloud API flow (the access token lives only in Power Automate).
import { EntityKey, IRecord } from '../models/model';
import { IWhatsAppTemplate } from '../models/settings';
import type { Engine } from './Engine';

/** International number in WhatsApp format (digits only, with country code), or '' when it is not usable. */
export function waNumber(raw: unknown, defaultCountryCode: string = ''): string {
  let s = String(raw || '').trim();
  if (!s) { return ''; }
  const plus = s.charAt(0) === '+';
  s = s.replace(/\D/g, '');
  if (!plus) {
    if (s.indexOf('00') === 0) { s = s.slice(2); } else if (s.charAt(0) === '0') { s = `${defaultCountryCode}${s.replace(/^0+/, '')}`; } else if (s.length <= 10 && defaultCountryCode) { s = `${defaultCountryCode}${s}`; }
  }
  return s.length >= 8 && s.length <= 15 ? s : '';
}

export function waLink(number: string, text: string): string {
  return `https://wa.me/${number}${text ? `?text=${encodeURIComponent(text)}` : ''}`;
}

/** Where a WhatsApp message for this record goes: the record's own number, else its contact's. */
export function waRecipient(engine: Engine, key: EntityKey, r: IRecord): { number: string; name: string; target: { key: EntityKey; id: number } } {
  const cc = engine.settings.whatsapp.defaultCountryCode;
  const own = waNumber(r.mobile || r.phone, cc);
  if (own && (key === 'leads' || key === 'contacts' || key === 'demos' || key === 'opportunities')) {
    return { number: own, name: r.title || r.contactName || '', target: { key, id: r.id! } };
  }
  const c = r.contactRef ? engine.find('contacts', r.contactRef) : undefined;
  return { number: waNumber(c?.mobile || c?.phone, cc) || own, name: c?.title || r.contactName || '', target: { key, id: r.id! } };
}

export function templateValues(engine: Engine, t: IWhatsAppTemplate, r: IRecord): string[] {
  return (t.params || []).map(p => engine.fillTemplate(`{{${p}}}`, r) || '-');
}

/** Logs a WhatsApp message; in Cloud mode it is queued for the "WhatsApp send" flow. */
export async function logWhatsApp(engine: Engine, key: EntityKey, r: IRecord, to: string, text: string, opts: { cloud: boolean; template?: IWhatsAppTemplate }): Promise<void> {
  const cloud = opts.cloud;
  const tpl = cloud && opts.template?.cloudName ? opts.template : undefined;
  await engine.ds.save('activities', {
    title: `WhatsApp to +${to}: ${text.replace(/\s+/g, ' ').slice(0, 120)}`.slice(0, 255),
    details: text, activityType: 'WhatsApp', relatedType: engine.def(key).activityName, relatedId: r.id,
    completed: !cloud, dueDate: new Date().toISOString(), owner: engine.me, activitySource: 'WhatsApp',
    waTo: to, waStatus: cloud ? 'Queued' : undefined, waTemplate: tpl ? `${tpl.cloudName}|${tpl.language || 'en'}` : '',
    waParams: tpl ? templateValues(engine, tpl, r) : undefined
  });
  if (engine.data.activities) { await engine.refresh('activities'); }
}
