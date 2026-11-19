// UI languages. English text is the key; translations come from table.ts.
import { TABLE } from './table';

export type Lang = 'en' | 'de' | 'fr' | 'es' | 'hi';
export const LANGS: { key: Lang; text: string }[] = [
  { key: 'en', text: 'English' }, { key: 'de', text: 'Deutsch' }, { key: 'fr', text: 'Français' }, { key: 'es', text: 'Español' }, { key: 'hi', text: 'हिन्दी' }
];
const COLS: Lang[] = ['en', 'de', 'fr', 'es', 'hi'];
let lang: Lang = 'en';
let dict: Record<string, string> | undefined;

function build(): Record<string, string> {
  const d: Record<string, string> = {};
  const col = COLS.indexOf(lang);
  TABLE.split('\n').forEach(line => {
    const parts = line.split('|');
    if (parts.length === COLS.length && parts[col]) { d[parts[0]] = parts[col]; }
  });
  return d;
}

export function getLang(): Lang { return lang; }

export function setLang(l: string | undefined): Lang {
  const next = (COLS.indexOf(l as Lang) >= 0 ? l : 'en') as Lang;
  if (next !== lang) { lang = next; dict = undefined; }
  return lang;
}

/** Best match for a culture such as "de-DE" or "hi-IN"; English otherwise. */
export function detectLang(culture?: string): Lang {
  const c = String(culture || '').slice(0, 2).toLowerCase();
  return (COLS.indexOf(c as Lang) >= 0 ? c : 'en') as Lang;
}

/** Translates a UI string (returns it unchanged when there is no translation). */
export function tr(s: string): string {
  if (lang === 'en' || !s) { return s; }
  if (!dict) { dict = build(); }
  return dict[s] || s;
}

/** Lower-cases a noun inside an English sentence only (other languages keep their capitalisation). */
export function lc(s: string): string { return lang === 'en' ? s.toLowerCase() : s; }

/** "New lead" / "Neu: Lead" style labels. */
export function newOf(singular: string): string { return lang === 'en' ? `New ${singular.toLowerCase()}` : `${tr('New')}: ${singular}`; }
