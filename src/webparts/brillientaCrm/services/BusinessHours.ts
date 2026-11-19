// Business-hours arithmetic for SLAs, in the support team's own time zone (IANA name, DST-aware).

export interface IBusinessHours {
  enabled: boolean;
  timeZone: string;
  /** Working days, 0 = Sunday … 6 = Saturday. */
  days: number[];
  /** "09:00" */
  start: string;
  /** "18:00" */
  end: string;
  /** Non-working dates, "YYYY-MM-DD" in the support time zone. */
  holidays: string[];
}

const MIN = 60000;

const toMin = (hhmm: string): number => {
  const m = /^(\d{1,2}):(\d{2})$/.exec(String(hhmm || '').trim());
  return m ? Math.min(24 * 60, Number(m[1]) * 60 + Number(m[2])) : 0;
};

const fmtCache: Record<string, Intl.DateTimeFormat> = {};
function parts(t: number, tz: string): { y: number; mo: number; d: number; h: number; mi: number; wd: number } {
  let f = fmtCache[tz];
  if (!f) {
    const opts = (zone: string): Intl.DateTimeFormatOptions => ({ timeZone: zone, hour12: false, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', weekday: 'short' });
    try { f = new Intl.DateTimeFormat('en-US', opts(tz)); } catch { f = new Intl.DateTimeFormat('en-US', opts('UTC')); }
    fmtCache[tz] = f;
  }
  const o: Record<string, string> = {};
  // eslint-disable-next-line @typescript-eslint/no-explicit-any -- formatToParts is ES2017; supported by every browser SPFx runs in
  ((f as any).formatToParts(new Date(t)) as { type: string; value: string }[]).forEach(p => { o[p.type] = p.value; });
  return {
    y: Number(o.year), mo: Number(o.month), d: Number(o.day), h: Number(o.hour) % 24, mi: Number(o.minute),
    wd: ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'].indexOf(o.weekday)
  };
}

/** UTC instant of a wall-clock time in `tz` (handles DST by re-checking the offset). */
function wallToUtc(y: number, mo: number, d: number, minutes: number, tz: string): number {
  const guess = Date.UTC(y, mo - 1, d, 0, minutes);
  let t = guess;
  for (let i = 0; i < 2; i++) {
    const p = parts(t, tz);
    const asUtc = Date.UTC(p.y, p.mo - 1, p.d, p.h, p.mi);
    t = t - (asUtc - guess);
  }
  return t;
}

const pad = (n: number): string => (n < 10 ? `0${n}` : String(n));

/** Working windows [startUtc, endUtc) for the local day that contains `t`, and the start of the next local day. */
function dayWindow(t: number, c: IBusinessHours): { open?: [number, number]; nextDay: number } {
  const p = parts(t, c.timeZone);
  const key = `${p.y}-${pad(p.mo)}-${pad(p.d)}`;
  const next = new Date(Date.UTC(p.y, p.mo - 1, p.d + 1));
  const nextDay = wallToUtc(next.getUTCFullYear(), next.getUTCMonth() + 1, next.getUTCDate(), 0, c.timeZone);
  const s = toMin(c.start);
  const e = toMin(c.end);
  if (c.days.indexOf(p.wd) < 0 || c.holidays.indexOf(key) >= 0 || e <= s) { return { nextDay }; }
  return { open: [wallToUtc(p.y, p.mo, p.d, s, c.timeZone), wallToUtc(p.y, p.mo, p.d, e, c.timeZone)], nextDay };
}

const usable = (c?: IBusinessHours): c is IBusinessHours =>
  !!c && c.enabled && (c.days || []).length > 0 && toMin(c.end) > toMin(c.start);

/** Adds `hours` of working time to `from`. Without business hours (or when off) this is plain clock time. */
export function addBusinessHours(from: Date, hours: number, c?: IBusinessHours): Date {
  if (!usable(c)) { return new Date(from.getTime() + hours * 3600000); }
  let left = Math.round(hours * 60) * MIN;
  let t = from.getTime();
  for (let guard = 0; guard < 3660; guard++) {
    const w = dayWindow(t, c);
    if (w.open) {
      const s = Math.max(t, w.open[0]);
      if (s < w.open[1]) {
        const avail = w.open[1] - s;
        if (left <= avail) { return new Date(s + left); }
        left -= avail;
      }
    }
    t = w.nextDay;
  }
  return new Date(t + left);
}

/** Working time (hours) between two instants. */
export function businessHoursBetween(a: Date, b: Date, c?: IBusinessHours): number {
  if (b.getTime() <= a.getTime()) { return 0; }
  if (!usable(c)) { return (b.getTime() - a.getTime()) / 3600000; }
  let total = 0;
  let t = a.getTime();
  const end = b.getTime();
  for (let guard = 0; guard < 3660 && t < end; guard++) {
    const w = dayWindow(t, c);
    if (w.open) {
      const s = Math.max(t, w.open[0]);
      const e = Math.min(end, w.open[1]);
      if (e > s) { total += e - s; }
    }
    t = w.nextDay;
  }
  return total / 3600000;
}

export function localTimeZone(): string {
  try { return Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC'; } catch { return 'UTC'; }
}
