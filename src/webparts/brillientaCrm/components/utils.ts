export function fmtDate(v?: string, withTime: boolean = false): string {
  if (!v) {
    return '';
  }
  const dt = new Date(v);
  if (isNaN(dt.getTime())) {
    return '';
  }
  return withTime
    ? dt.toLocaleString(undefined, { year: 'numeric', month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' })
    : dt.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' });
}

export function fmtMoney(v: number, currency: string): string {
  try {
    return new Intl.NumberFormat(undefined, { style: 'currency', currency: currency || 'USD', maximumFractionDigits: 0 }).format(v || 0);
  } catch {
    return `${currency} ${Math.round(v || 0).toLocaleString()}`;
  }
}

export function isOverdue(v?: string): boolean {
  return !!v && new Date(v).getTime() < Date.now();
}

export function isSameMonth(v: string | undefined, ref: Date = new Date()): boolean {
  if (!v) {
    return false;
  }
  const dt = new Date(v);
  return dt.getFullYear() === ref.getFullYear() && dt.getMonth() === ref.getMonth();
}

export function daysFromNow(v?: string): number | undefined {
  if (!v) {
    return undefined;
  }
  return Math.round((new Date(v).getTime() - Date.now()) / 86400000);
}

export function matchesSearch(text: string, values: (string | undefined)[]): boolean {
  if (!text) {
    return true;
  }
  const q = text.toLowerCase();
  return values.some(v => !!v && v.toLowerCase().indexOf(q) >= 0);
}

/** Sort helper usable with DetailsList column clicks. */
export function sortBy<T>(items: T[], key: keyof T, desc: boolean): T[] {
  return [...items].sort((a, b) => {
    const av = a[key] as unknown;
    const bv = b[key] as unknown;
    if (av === bv) {
      return 0;
    }
    if (av === undefined || av === null || av === '') {
      return 1;
    }
    if (bv === undefined || bv === null || bv === '') {
      return -1;
    }
    const r = typeof av === 'number' && typeof bv === 'number' ? av - bv : String(av).localeCompare(String(bv));
    return desc ? -r : r;
  });
}

export function exportCsv(fileName: string, rows: Record<string, unknown>[]): void {
  if (!rows.length) {
    return;
  }
  const headers = Object.keys(rows[0]);
  const cell = (v: unknown): string => {
    const str = v === undefined || v === null ? '' : String(v);
    // Neutralise spreadsheet formula injection and escape quotes.
    const safe = /^[=+\-@]/.test(str) ? `'${str}` : str;
    return `"${safe.replace(/"/g, '""')}"`;
  };
  const csv = [headers.map(cell).join(','), ...rows.map(r => headers.map(h => cell(r[h])).join(','))].join('\r\n');
  const blob = new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}

export function toInputDate(v?: string): Date | undefined {
  if (!v) {
    return undefined;
  }
  const dt = new Date(v);
  return isNaN(dt.getTime()) ? undefined : dt;
}

export function errorText(e: unknown): string {
  return e instanceof Error ? e.message : String(e);
}

export function mailto(email: string, subject: string = ''): string {
  return `mailto:${encodeURIComponent(email)}${subject ? `?subject=${encodeURIComponent(subject)}` : ''}`;
}

export function uid(): string {
  return Math.random().toString(36).slice(2, 10) + Date.now().toString(36);
}

export function pad(n: number, len: number = 2): string {
  let s = String(n);
  while (s.length < len) { s = '0' + s; }
  return s;
}

/** RFC 4180 CSV parser (quoted fields, embedded commas/newlines, BOM). */
export function parseCsv(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = '';
  let q = false;
  const src = text.replace(/^﻿/, '');
  const delim = (src.split('\n')[0].match(/;/g) || []).length > (src.split('\n')[0].match(/,/g) || []).length ? ';' : ',';
  for (let i = 0; i < src.length; i++) {
    const c = src[i];
    if (q) {
      if (c === '"') {
        if (src[i + 1] === '"') { cell += '"'; i++; } else { q = false; }
      } else { cell += c; }
    } else if (c === '"') { q = true; }
    else if (c === delim) { row.push(cell); cell = ''; }
    else if (c === '\n' || c === '\r') {
      if (c === '\r' && src[i + 1] === '\n') { i++; }
      row.push(cell); cell = '';
      if (row.some(x => x !== '')) { rows.push(row); }
      row = [];
    } else { cell += c; }
  }
  row.push(cell);
  if (row.some(x => x !== '')) { rows.push(row); }
  return rows;
}

export function quarterRange(offset: number = 0, fyStartMonth: number = 1, ref: Date = new Date()): { start: Date; end: Date; label: string } {
  const m0 = fyStartMonth - 1;
  const rel = (ref.getMonth() - m0 + 12) % 12;
  const qIdx = Math.floor(rel / 3) + offset;
  const startMonthAbs = ref.getFullYear() * 12 + ref.getMonth() - rel + (qIdx * 3);
  const start = new Date(Math.floor(startMonthAbs / 12), startMonthAbs % 12, 1);
  const end = new Date(start.getFullYear(), start.getMonth() + 3, 1);
  const fyYear = start.getMonth() >= m0 && m0 > 0 ? start.getFullYear() + 1 : start.getFullYear();
  const qn = ((Math.floor(((start.getMonth() - m0 + 12) % 12) / 3)) + 1);
  return { start, end, label: `Q${qn} FY${String(fyYear).slice(2)}` };
}

export function inRange(v: string | undefined, start: Date, end: Date): boolean {
  if (!v) { return false; }
  const t = new Date(v).getTime();
  return t >= start.getTime() && t < end.getTime();
}

export function monthLabel(d: Date): string {
  return d.toLocaleDateString(undefined, { month: 'short', year: '2-digit' });
}

export function lsGet<T>(key: string, fallback: T): T {
  try {
    const v = window.localStorage.getItem(key);
    return v ? (JSON.parse(v) as T) : fallback;
  } catch {
    return fallback;
  }
}

export function lsSet(key: string, value: unknown): void {
  try {
    window.localStorage.setItem(key, JSON.stringify(value));
  } catch {
    /* storage unavailable */
  }
}

export function sum<T>(items: T[], fn: (x: T) => number): number {
  return items.reduce((s, x) => s + (fn(x) || 0), 0);
}

export function groupCount<T>(items: T[], fn: (x: T) => string): { label: string; value: number }[] {
  const m: Record<string, number> = {};
  items.forEach(x => { const k = fn(x) || '(none)'; m[k] = (m[k] || 0) + 1; });
  return Object.keys(m).map(label => ({ label, value: m[label] })).sort((a, b) => b.value - a.value);
}

/** Compact date for tables and cards: "18 Nov 10:00" (year only when not this year). */
export function fmtShort(v?: string, withTime: boolean = true): string {
  if (!v) { return ''; }
  const d = new Date(v);
  if (isNaN(d.getTime())) { return ''; }
  const sameYear = d.getFullYear() === new Date().getFullYear();
  const date = d.toLocaleDateString(undefined, sameYear ? { day: 'numeric', month: 'short' } : { day: 'numeric', month: 'short', year: 'numeric' });
  return withTime ? `${date} ${d.toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' })}` : date;
}

/** Open, not paused (waiting on customer) and past its SLA deadline. */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export function slaBreached(c: Record<string, any>): boolean {
  return c.caseStatus !== 'Resolved' && c.caseStatus !== 'Closed' && !c.slaPausedAt && isOverdue(c.slaDue);
}

/** Open case whose first response is late. */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export function firstResponseLate(c: Record<string, any>): boolean {
  return c.caseStatus === 'New' && !c.firstRespondedAt && isOverdue(c.firstResponseDue);
}
