/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
import { Checkbox, Dropdown, Link, SpinButton, TagPicker, TextField, ITag, Label, Icon, Position } from '@fluentui/react';
import { IFieldDef, IRecord } from '../../models/model';
import { ICrmSettings } from '../../models/settings';
import { Engine } from '../../services/Engine';
import { useCrm } from '../CrmContext';
import { DateTimeField, LookupPicker, Pill, ScoreBadge, UserPicker } from './controls';
import { fmtDate, fmtMoney, fmtShort, isOverdue } from '../utils';

export function choicesFor(fd: IFieldDef, s: ICrmSettings, current?: string): string[] {
  let list: string[] = fd.choices || [];
  switch (fd.choicesFrom) {
    case 'stages': list = s.stages.map(x => x.name); break;
    case 'leadSources': list = s.leadSources; break;
    case 'industries': list = s.industries; break;
    case 'caseCategories': list = s.caseCategories; break;
    case 'lostReasons': list = s.lostReasons; break;
  }
  return current && list.indexOf(current) < 0 ? [current, ...list] : list;
}

const safeUrl = (v: string): string | undefined => (/^https?:\/\//i.test(v) ? v : (/^[\w.-]+\.[a-z]{2,}/i.test(v) ? `https://${v}` : undefined));

/** Read-only rendering of a field value (grids, cards, read-only forms). */
export function renderValue(fd: IFieldDef, r: IRecord, engine: Engine, currency: string, compact: boolean = false): React.ReactNode {
  const v = r[fd.key];
  if (v === undefined || v === null || v === '' || (Array.isArray(v) && !v.length)) {
    return fd.type === 'boolean' ? 'No' : '';
  }
  switch (fd.type) {
    case 'currency': return fmtMoney(Number(v), currency);
    case 'percent': return `${v}%`;
    case 'number': return fd.key === 'leadScore' ? <ScoreBadge score={Number(v)} /> : Number(v).toLocaleString();
    case 'date': return compact ? fmtShort(v, false) : fmtDate(v);
    case 'datetime': {
      const overdue = (fd.key === 'nextFollowUp' || fd.key === 'slaDue' || fd.key === 'dueDate') && isOverdue(v) &&
        !(r.completed || r.caseStatus === 'Resolved' || r.caseStatus === 'Closed' || r.status === 'Converted');
      return <span style={overdue ? { color: 'var(--bad)', fontWeight: 600 } : undefined} title={fmtDate(v, true)}>{compact ? fmtShort(v) : fmtDate(v, true)}</span>;
    }
    case 'boolean': return v ? 'Yes' : 'No';
    case 'user': return v.title;
    case 'lookup': return engine.display(fd.lookup!, Number(v));
    case 'choice': return <Pill text={String(v)} />;
    case 'tags': return (v as string[]).map(t => <span key={t} className="crm-tag">{t}</span>);
    case 'email': return <Link href={`mailto:${v}`} onClick={e => e.stopPropagation()}>{v}</Link>;
    case 'phone': return <Link href={`tel:${String(v).replace(/[^\d+]/g, '')}`} onClick={e => e.stopPropagation()}>{v}</Link>;
    case 'url': {
      const u = safeUrl(String(v));
      return u ? <Link href={u} target="_blank" rel="noopener noreferrer" onClick={e => e.stopPropagation()}>{String(v).replace(/^https?:\/\//, '').slice(0, 60)}</Link> : String(v);
    }
    case 'json': return '';
    default: return String(v);
  }
}

/** Plain text value for CSV export / search. */
export function textValue(fd: IFieldDef, r: IRecord, engine: Engine): string {
  const v = r[fd.key];
  if (v === undefined || v === null) { return ''; }
  switch (fd.type) {
    case 'user': return v.title || '';
    case 'lookup': return engine.display(fd.lookup!, Number(v));
    case 'boolean': return v ? 'Yes' : 'No';
    case 'tags': return (v as string[]).join(', ');
    case 'json': return JSON.stringify(v);
    default: return String(v);
  }
}

export const FieldEditor: React.FC<{ fd: IFieldDef; value: any; onChange: (v: any) => void; disabled?: boolean; error?: string }> =
  ({ fd, value, onChange, disabled, error }) => {
    const { settings, engine } = useCrm();
    const ro = disabled || fd.readOnly;
    if (fd.readOnly && !disabled) {
      // computed fields: show value as text
      return (
        <div>
          <Label>{fd.label} <Icon iconName="Lock" title="Calculated automatically" style={{ fontSize: 10, opacity: 0.6 }} /></Label>
          <div style={{ minHeight: 32, display: 'flex', alignItems: 'center', gap: 4, flexWrap: 'wrap' }}>
            {renderValue(fd, { [fd.key]: value }, engine, settings.currency) || <span style={{ opacity: 0.5 }}>—</span>}
          </div>
        </div>
      );
    }
    switch (fd.type) {
      case 'note':
        return <TextField label={fd.label} multiline autoAdjustHeight rows={3} value={value || ''} disabled={ro} required={fd.required} errorMessage={error} onChange={(_, v) => onChange(v)} />;
      case 'choice': {
        const opts = choicesFor(fd, settings, value);
        return <Dropdown label={fd.label} selectedKey={value || ''} disabled={ro} required={fd.required} errorMessage={error}
          options={[...(fd.required ? [] : [{ key: '', text: '—' }]), ...opts.map(o => ({ key: o, text: o }))]} onChange={(_, o) => onChange(o?.key ? String(o.key) : '')} />;
      }
      case 'number':
        return <TextField label={fd.label} type="number" value={value === undefined || value === null ? '' : String(value)} disabled={ro} required={fd.required} errorMessage={error}
          onChange={(_, v) => onChange(v === '' || v === undefined ? undefined : Number(v))} />;
      case 'currency':
        return <TextField label={`${fd.label} (${settings.currency})`} type="number" min={0} step="0.01" value={value === undefined || value === null ? '' : String(value)} disabled={ro}
          required={fd.required} errorMessage={error} onChange={(_, v) => onChange(v === '' || v === undefined ? undefined : Number(v))} />;
      case 'percent':
        return <SpinButton label={`${fd.label} (%)`} labelPosition={Position.top} min={0} max={100} step={5} value={String(value ?? 0)} disabled={ro}
          onChange={(_, v) => onChange(Math.max(0, Math.min(100, Number(v) || 0)))} styles={{ label: { fontWeight: 600 } }} />;
      case 'date':
        return <DateTimeField label={fd.label} value={value} showTime={false} required={fd.required} disabled={ro} onChange={onChange} />;
      case 'datetime':
        return <DateTimeField label={fd.label} value={value} required={fd.required} disabled={ro} onChange={onChange} />;
      case 'boolean':
        return <div style={{ paddingTop: 30 }}><Checkbox label={fd.label} checked={!!value} disabled={ro} onChange={(_, c) => onChange(!!c)} /></div>;
      case 'user':
        return <UserPicker label={fd.label} value={value} disabled={ro} onChange={onChange} />;
      case 'lookup':
        return <LookupPicker label={fd.label} entity={fd.lookup!} value={value} disabled={ro} onChange={onChange} />;
      case 'tags': {
        const tags: string[] = value || [];
        return (
          <div>
            <Label disabled={ro}>{fd.label}</Label>
            <TagPicker
              selectedItems={tags.map(t => ({ key: t, name: t }))}
              disabled={ro}
              onResolveSuggestions={(q: string): ITag[] => (q.trim() ? [{ key: q.trim(), name: q.trim() }] : [])}
              onChange={items => onChange((items || []).map(i => i.name))}
              inputProps={{ placeholder: 'Type a tag and press Enter', 'aria-label': fd.label }}
            />
          </div>
        );
      }
      case 'json':
        return null;
      default:
        return <TextField label={fd.label} type={fd.type === 'email' ? 'email' : fd.type === 'phone' ? 'tel' : 'text'} value={value || ''} disabled={ro}
          required={fd.required} errorMessage={error} onChange={(_, v) => onChange(v)} />;
    }
  };
