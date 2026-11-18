import * as React from 'react';
import { Checkbox, DefaultButton, Icon, MessageBarType, PrimaryButton, Stack, Text, TextField, Link } from '@fluentui/react';
import { ACTIVITY_TYPES, EntityKey, IRecord, entityByActivityName } from '../../models/model';
import { useCrm, useEntities } from '../CrmContext';
import { DateTimeField, Pill, UserPicker } from './controls';
import { errorText, fmtDate, isOverdue } from '../utils';

const ICONS: Record<string, string> = { Call: 'Phone', Email: 'Mail', Meeting: 'Group', Note: 'QuickNote', Task: 'TaskSolid' };

export const ActivityIcon: React.FC<{ type: string }> = ({ type }) => <span className="crm-act-icon"><Icon iconName={ICONS[type] || 'QuickNote'} /></span>;

export const ActivityRow: React.FC<{ a: IRecord; showRelated?: boolean }> = ({ a, showRelated }) => {
  const { engine, notify, open } = useCrm();
  const canEdit = engine.can('edit', 'activities');
  const toggle = (c: boolean): void => {
    engine.patch('activities', a.id!, { completed: c }).catch(e => notify(errorText(e), MessageBarType.error));
  };
  const relDef = showRelated && a.relatedType ? entityByActivityName(a.relatedType, engine.defs) : undefined;
  return (
    <div className="crm-timeline-item">
      <ActivityIcon type={a.activityType} />
      <div style={{ flex: 1, minWidth: 0 }}>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center" tokens={{ childrenGap: 8 }}>
          <Link onClick={() => open('activities', a)} styles={{ root: { fontWeight: 600, textDecoration: a.activityType === 'Task' && a.completed ? 'line-through' : undefined } }}>{a.title}</Link>
          {a.activityType === 'Task' && <Checkbox label="Done" checked={!!a.completed} disabled={!canEdit} onChange={(_, c) => toggle(!!c)} />}
        </Stack>
        {a.details && <Text variant="small" block style={{ whiteSpace: 'pre-wrap', wordBreak: 'break-word' }}>{a.details.length > 600 ? `${a.details.slice(0, 600)}…` : a.details}</Text>}
        <Text variant="small" block className="crm-muted">
          {relDef && a.relatedId ? <><Link onClick={() => { const r = engine.find(relDef.key, a.relatedId); if (r) { open(relDef.key, r); } }}>{relDef.singular}: {engine.display(relDef.key, a.relatedId)}</Link> · </> : null}
          {a.activityType} · {a.author?.title || ''} · {fmtDate(a.created, true)}
          {a.dueDate && <> · due {fmtDate(a.dueDate, true)}</>}
          {a.owner && <> · {a.owner.title}</>}
          {a.activityType === 'Task' && !a.completed && isOverdue(a.dueDate) && <> <Pill text="Overdue" /></>}
        </Text>
      </div>
    </div>
  );
};

export const ActivityTimeline: React.FC<{ entity: EntityKey; id?: number; record?: IRecord }> = ({ entity, id }) => {
  const { engine, notify } = useCrm();
  const { rows } = useEntities('activities');
  const [draft, setDraft] = React.useState<IRecord>();
  const [busy, setBusy] = React.useState(false);
  if (!id) { return <Text variant="small">Save the record first to log activities.</Text>; }
  const relatedType = engine.def(entity).activityName;
  const items = rows('activities').filter(a => a.relatedType === relatedType && a.relatedId === id)
    .sort((a, b) => String(b.created || '').localeCompare(String(a.created || '')));
  const canAdd = engine.can('create', 'activities');

  const save = (): void => {
    if (!draft) { return; }
    setBusy(true);
    engine.save('activities', {
      ...draft, title: draft.title || draft.activityType, relatedType, relatedId: id,
      completed: draft.activityType !== 'Task', owner: draft.activityType === 'Task' ? (draft.owner || engine.me) : undefined
    }).then(() => { setDraft(undefined); setBusy(false); })
      .catch(e => { setBusy(false); notify(errorText(e), MessageBarType.error); });
  };

  return (
    <Stack tokens={{ childrenGap: 8 }}>
      {canAdd && !draft && (
        <Stack horizontal wrap tokens={{ childrenGap: 6 }}>
          {ACTIVITY_TYPES.map(t => (
            <DefaultButton key={t} iconProps={{ iconName: ICONS[t] }} text={t === 'Task' ? 'Add task' : `Log ${t.toLowerCase()}`}
              onClick={() => setDraft({ activityType: t, title: '', details: '' })} />
          ))}
        </Stack>
      )}
      {draft && (
        <Stack tokens={{ childrenGap: 8 }} className="crm-card">
          <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>{draft.activityType === 'Task' ? 'New task' : `Log ${draft.activityType.toLowerCase()}`}</Text>
          <TextField label="Subject" value={draft.title} onChange={(_, v) => setDraft({ ...draft, title: v })} autoFocus />
          <TextField label={draft.activityType === 'Call' ? 'Call notes' : 'Details'} multiline rows={3} value={draft.details} onChange={(_, v) => setDraft({ ...draft, details: v })} />
          {draft.activityType === 'Task' && (
            <>
              <DateTimeField label="Due" value={draft.dueDate} onChange={v => setDraft({ ...draft, dueDate: v })} />
              <UserPicker label="Assigned to" value={draft.owner || engine.me} onChange={u => setDraft({ ...draft, owner: u })} />
            </>
          )}
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton text="Save" disabled={busy} onClick={save} />
            <DefaultButton text="Cancel" onClick={() => setDraft(undefined)} />
          </Stack>
        </Stack>
      )}
      {items.length === 0 && <Text variant="small" className="crm-muted">No activity yet.</Text>}
      {items.map(a => <ActivityRow key={a.id} a={a} />)}
    </Stack>
  );
};
