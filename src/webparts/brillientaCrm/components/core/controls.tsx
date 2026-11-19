import { tr } from '../../i18n';
import * as React from 'react';
import {
  ActionButton, DatePicker, DayOfWeek, IPersonaProps, ITag, Label, NormalPeoplePicker, Stack, TagPicker, TextField
} from '@fluentui/react';
import { EntityKey, IUserRef } from '../../models/model';
import { useCrm } from '../CrmContext';
import { pad, toInputDate } from '../utils';

/* ---------------------------------------------------------------- people */

export const UserPicker: React.FC<{ label?: string; value?: IUserRef; onChange: (u?: IUserRef) => void; disabled?: boolean; showMe?: boolean }> =
  ({ label, value, onChange, disabled, showMe = true }) => {
    const { engine } = useCrm();
    const me = engine.me;
    const selected: IPersonaProps[] = value ? [{ key: String(value.id), text: value.title, secondaryText: value.email }] : [];
    const resolve = async (q: string): Promise<IPersonaProps[]> => {
      if (!q || q.length < 2) { return []; }
      try {
        return (await engine.ds.searchPeople(q)).map(p => ({ key: p.key, text: p.text, secondaryText: p.email }));
      } catch {
        return [];
      }
    };
    const change = async (items?: IPersonaProps[]): Promise<void> => {
      const p = items && items[0];
      if (!p) { onChange(undefined); return; }
      if (value && String(p.key) === String(value.id)) { return; }
      onChange(await engine.ds.ensureUser(String(p.key)));
    };
    return (
      <div>
        {label && (
          <Stack horizontal horizontalAlign="space-between" verticalAlign="end">
            <Label disabled={disabled}>{label}</Label>
            {showMe && !disabled && me && (!value || value.id !== me.id) && (
              <ActionButton iconProps={{ iconName: 'Contact' }} styles={{ root: { height: 24 } }} onClick={() => onChange(me)}>Assign to me</ActionButton>
            )}
          </Stack>
        )}
        <NormalPeoplePicker
          selectedItems={selected}
          itemLimit={1}
          disabled={disabled}
          onResolveSuggestions={resolve}
          onChange={items => { change(items).catch(() => undefined); }}
          pickerSuggestionsProps={{ noResultsFoundText: 'No people found', loadingText: 'Searching...' }}
          inputProps={{ placeholder: value ? '' : 'Search people', 'aria-label': label || 'Person' }}
          resolveDelay={300}
        />
      </div>
    );
  };

/* ---------------------------------------------------------------- lookup */

export const LookupPicker: React.FC<{ label?: string; entity: EntityKey; value?: number; onChange: (id?: number) => void; disabled?: boolean; filter?: (r: { [k: string]: unknown }) => boolean }> =
  ({ label, entity, value, onChange, disabled, filter }) => {
    const { engine, data } = useCrm();
    React.useEffect(() => { engine.ensure([entity]).catch(() => undefined); }, [entity]);
    const rows = (data[entity] || []).filter(r => !filter || filter(r));
    const selected: ITag[] = value ? [{ key: value, name: engine.display(entity, value) }] : [];
    const suggest = (q: string): ITag[] => {
      const s = q.toLowerCase();
      return rows.filter(r => !s || String(r.title || '').toLowerCase().indexOf(s) >= 0 || String(r.email || '').toLowerCase().indexOf(s) >= 0)
        .slice(0, 25).map(r => ({ key: r.id!, name: r.title || r.quoteNumber || `#${r.id}` }));
    };
    return (
      <div>
        {label && <Label disabled={disabled}>{label}</Label>}
        <TagPicker
          selectedItems={selected}
          itemLimit={1}
          disabled={disabled}
          onResolveSuggestions={suggest}
          onEmptyResolveSuggestions={() => suggest('')}
          onChange={items => onChange(items && items[0] ? Number(items[0].key) : undefined)}
          pickerSuggestionsProps={{ noResultsFoundText: `No ${engine.def(entity).plural.toLowerCase()} found` }}
          inputProps={{ placeholder: value ? '' : `Search ${engine.def(entity).plural.toLowerCase()}`, 'aria-label': label || entity }}
        />
      </div>
    );
  };

/* ---------------------------------------------------------------- date/time */

export const DateTimeField: React.FC<{ label?: string; value?: string; onChange: (iso?: string) => void; showTime?: boolean; required?: boolean; disabled?: boolean }> =
  ({ label, value, onChange, showTime = true, required, disabled }) => {
    const dt = toInputDate(value);
    const time = dt ? `${pad(dt.getHours())}:${pad(dt.getMinutes())}` : '';
    const setDate = (date?: Date | null): void => {
      if (!date) { onChange(undefined); return; }
      const next = new Date(date);
      if (showTime) { next.setHours(dt ? dt.getHours() : 10, dt ? dt.getMinutes() : 0, 0, 0); } else { next.setHours(12, 0, 0, 0); }
      onChange(next.toISOString());
    };
    const setTime = (t?: string): void => {
      if (!t) { return; }
      const [h, m] = t.split(':').map(Number);
      const next = dt ? new Date(dt) : new Date();
      next.setHours(h || 0, m || 0, 0, 0);
      onChange(next.toISOString());
    };
    return (
      <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="end">
        <Stack.Item grow>
          <DatePicker label={label} isRequired={required} value={dt} onSelectDate={setDate} disabled={disabled}
            firstDayOfWeek={DayOfWeek.Monday} formatDate={x => (x ? x.toLocaleDateString() : '')} placeholder="Select a date" allowTextInput />
        </Stack.Item>
        {showTime && <TextField type="time" value={time} disabled={disabled} onChange={(_, v) => setTime(v)} styles={{ root: { width: 116 } }} ariaLabel={`${label || 'Date'} time`} />}
        {dt && !disabled && <ActionButton iconProps={{ iconName: 'Clear' }} title="Clear" ariaLabel="Clear date" onClick={() => onChange(undefined)} />}
      </Stack>
    );
  };

/* ---------------------------------------------------------------- pill */

type Tone = 'info' | 'warn' | 'good' | 'bad' | 'neutral' | 'accent';
const TONES: Record<Tone, [string, string]> = {
  info: ['#deecf9', '#004578'], warn: ['#fff4ce', '#6b4f00'], good: ['#dff6dd', '#0b5a08'],
  bad: ['#fde7e9', '#a4262c'], neutral: ['#f3f2f1', '#3b3a39'], accent: ['#ece6f6', '#4b2a8a']
};
const AUTO: Record<string, Tone> = {
  New: 'info', Contacted: 'warn', Qualified: 'good', Unqualified: 'neutral', Converted: 'accent',
  Hot: 'bad', Warm: 'warn', Cold: 'info',
  Requested: 'info', Scheduled: 'warn', Completed: 'good', 'No-show': 'bad', Cancelled: 'neutral',
  'Closed Won': 'good', 'Closed Lost': 'bad', Won: 'good', Lost: 'bad',
  Draft: 'neutral', Sent: 'info', Accepted: 'good', Rejected: 'bad', Expired: 'neutral',
  'In progress': 'warn', 'Waiting on customer': 'accent', Resolved: 'good', Closed: 'neutral',
  Low: 'neutral', Normal: 'info', High: 'warn', Critical: 'bad',
  Planned: 'info', Active: 'good', Prospect: 'info', Customer: 'good', Partner: 'accent', Competitor: 'bad',
  Pending: 'warn', Approved: 'good', Overdue: 'bad', Breached: 'bad', 'At risk': 'warn', Done: 'good', Open: 'info',
  Commit: 'good', 'Best case': 'warn', Pipeline: 'info', Omitted: 'neutral', Interested: 'good', 'Not interested': 'neutral', 'Follow-up needed': 'warn'
};

export const Pill: React.FC<{ text?: string; tone?: Tone; icon?: string }> = ({ text, tone }) => {
  if (!text) { return null; }
  const [bg, fg] = TONES[tone || AUTO[text] || 'neutral'];
  return <span className="crm-pill" style={{ background: bg, color: fg }} title={tr(text)}>{tr(text)}</span>;
};

export const ScoreBadge: React.FC<{ score?: number }> = ({ score }) => {
  if (score === undefined || score === null) { return null; }
  const tone: Tone = score >= 60 ? 'bad' : score >= 30 ? 'warn' : 'info';
  return <Pill text={`${score}`} tone={tone} />;
};
