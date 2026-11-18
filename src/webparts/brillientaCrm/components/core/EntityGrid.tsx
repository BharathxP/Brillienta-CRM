import * as React from 'react';
import {
  CommandBar, ICommandBarItemProps, SearchBox, Dropdown, Stack, Text, IColumn, Selection, SelectionMode, ShimmeredDetailsList,
  DetailsListLayoutMode, MessageBarType, Dialog, DialogType, DialogFooter, PrimaryButton, DefaultButton, Checkbox, TextField,
  IContextualMenuItem, ConstrainMode
} from '@fluentui/react';
import { EntityKey, IFieldDef, IRecord, LOST, WON } from '../../models/model';
import { useCrm, useEntities } from '../CrmContext';
import { choicesFor, renderValue, textValue } from './fields';
import { BulkEditDialog, ImportDialog } from './dialogs';
import { errorText, exportCsv, fmtMoney, lsGet, lsSet, sortBy, sum } from '../utils';

interface IView { name: string; q: string; filters: Record<string, string>; sortKey: string; desc: boolean; mode: 'list' | 'board' }

const ALL = '__all';

export const EntityGrid: React.FC<{ entity: EntityKey; title?: string; initialFilters?: Record<string, string>; defaultMode?: 'list' | 'board' }> =
  ({ entity, initialFilters = {}, defaultMode }) => {
    const { engine, open, notify, currency, settings, isNarrow } = useCrm();
    const def = engine.def(entity);
    const { loading, rows } = useEntities(entity, ...(def.fields.filter(f => f.type === 'lookup').map(f => f.lookup!) as EntityKey[]));
    const viewsKey = `brillienta-crm:views:${entity}`;
    const colsKey = `brillienta-crm:cols:${entity}`;
    const [q, setQ] = React.useState('');
    const [filters, setFilters] = React.useState<Record<string, string>>(initialFilters);
    const [sortKey, setSortKey] = React.useState<string>(entity === 'leads' ? 'created' : 'modified');
    const [desc, setDesc] = React.useState(true);
    const [mode, setMode] = React.useState<'list' | 'board'>(defaultMode || (isNarrow ? 'list' : 'list'));
    const [views, setViews] = React.useState<IView[]>(() => lsGet<IView[]>(viewsKey, []));
    const [visibleCols, setVisibleCols] = React.useState<string[]>(() => lsGet<string[]>(colsKey, def.fields.filter(f => f.grid).map(f => f.key)));
    const [selected, setSelected] = React.useState<IRecord[]>([]);
    const [dialog, setDialog] = React.useState<'import' | 'bulk' | 'cols' | 'saveView' | 'delete'>();
    const [viewName, setViewName] = React.useState('');
    const [dragId, setDragId] = React.useState<number>();
    const [over, setOver] = React.useState<string>();
    const selection = React.useMemo((): Selection => {
      const s: Selection = new Selection({ onSelectionChanged: (): void => setSelected(s.getSelection() as IRecord[]) });
      return s;
    }, []);

    const filterFields = def.fields.filter(f => f.filter);
    const all = rows(entity);
    const filtered = sortBy(all.filter(r => {
      for (const k of Object.keys(filters)) {
        const v = filters[k];
        if (!v || v === ALL) { continue; }
        const fd = def.fields.find(f => f.key === k);
        if (!fd) { continue; }
        if (fd.type === 'user') {
          if (v === 'me' && r[k]?.id !== engine.me?.id) { return false; }
          if (v === 'none' && r[k]) { return false; }
        } else if (fd.type === 'boolean') {
          if ((v === 'yes') !== !!r[k]) { return false; }
        } else if (k === '__open') {
          continue;
        } else if (String(r[k] || '') !== v) {
          return false;
        }
      }
      if (filters.__open === 'open') {
        if (entity === 'opportunities' && (r.stage === WON || r.stage === LOST)) { return false; }
        if (entity === 'leads' && (r.status === 'Converted' || r.status === 'Unqualified')) { return false; }
        if (entity === 'cases' && (r.caseStatus === 'Resolved' || r.caseStatus === 'Closed')) { return false; }
        if (entity === 'activities' && r.completed) { return false; }
        if (entity === 'demos' && !(r.status === 'Requested' || r.status === 'Scheduled')) { return false; }
      }
      if (!q) { return true; }
      const s = q.toLowerCase();
      return def.searchFields.some(k => String(r[k] || '').toLowerCase().indexOf(s) >= 0) ||
        (r.tags || []).some((t: string) => t.toLowerCase().indexOf(s) >= 0);
    }), sortKey as keyof IRecord, desc);

    const moneyField = def.fields.find(f => f.type === 'currency' && f.grid);

    const sortableValue = (fd: IFieldDef): string => fd.key;
    const columns: IColumn[] = def.fields.filter(f => visibleCols.indexOf(f.key) >= 0 || f.key === 'title').map((f): IColumn => ({
      key: f.key, name: f.label, fieldName: sortableValue(f), minWidth: f.key === 'title' ? 160 : (f.type === 'choice' ? 90 : 100), maxWidth: f.key === 'title' ? 280 : 200,
      isResizable: true, isSorted: sortKey === f.key, isSortedDescending: desc,
      onColumnClick: () => { if (sortKey === f.key) { setDesc(!desc); } else { setSortKey(f.key); setDesc(false); } },
      onRender: (r: IRecord) => f.key === 'title'
        ? <span className="crm-link-cell">{r.title || r.quoteNumber || `#${r.id}`}</span>
        : renderValue(f, r, engine, currency)
    }));
    // title first
    columns.sort((a, b) => (a.key === 'title' ? -1 : b.key === 'title' ? 1 : 0));
    columns.push({ key: 'modified', name: 'Modified', fieldName: 'modified', minWidth: 90, isSorted: sortKey === 'modified', isSortedDescending: desc,
      onColumnClick: () => { if (sortKey === 'modified') { setDesc(!desc); } else { setSortKey('modified'); setDesc(true); } },
      onRender: (r: IRecord) => (r.modified ? new Date(r.modified).toLocaleDateString() : '') });

    const doExport = (): void => {
      const fields = def.fields.filter(f => f.type !== 'json' && f.key !== 'rawPayload');
      exportCsv(`${def.plural.toLowerCase().replace(/\s/g, '-')}.csv`, filtered.map(r => {
        const row: Record<string, unknown> = { ID: r.id };
        fields.forEach(f => { row[f.label] = textValue(f, r, engine); });
        row.Created = r.created; row['Created by'] = r.author?.title || ''; row.Modified = r.modified;
        return row;
      }));
    };

    const bulkDelete = (): void => {
      setDialog(undefined);
      (async () => {
        for (const r of selected) { await engine.remove(entity, r.id!); }
      })().then(() => notify(`${selected.length} record(s) deleted.`, MessageBarType.success)).catch(e => notify(errorText(e), MessageBarType.error));
    };

    const saveView = (): void => {
      const next = [...views.filter(v => v.name !== viewName), { name: viewName, q, filters, sortKey, desc, mode }];
      setViews(next);
      lsSet(viewsKey, next);
      setDialog(undefined);
    };
    const applyView = (v: IView): void => { setQ(v.q); setFilters(v.filters); setSortKey(v.sortKey); setDesc(v.desc); setMode(v.mode); };

    const viewItems: IContextualMenuItem[] = [
      { key: 'save', text: 'Save current view…', iconProps: { iconName: 'Save' }, onClick: () => { setViewName(''); setDialog('saveView'); } },
      { key: 'reset', text: 'Reset filters', iconProps: { iconName: 'ClearFilter' }, onClick: () => { setQ(''); setFilters({}); } },
      ...(views.length ? [{ key: 'div', itemType: 1 } as IContextualMenuItem] : []),
      ...views.map(v => ({
        key: `v-${v.name}`, text: v.name, onClick: () => applyView(v),
        subMenuProps: { items: [{ key: 'apply', text: 'Apply', onClick: () => applyView(v) }, { key: 'del', text: 'Delete view', onClick: () => { const n = views.filter(x => x.name !== v.name); setViews(n); lsSet(viewsKey, n); } }] }
      }))
    ];

    const sel = selected.length;
    const items: ICommandBarItemProps[] = [];
    if (engine.can('create', entity)) {
      items.push({ key: 'new', text: `New ${def.singular.toLowerCase()}`, iconProps: { iconName: 'Add' }, onClick: () => open(entity, {}) });
    }
    if (sel && engine.can('edit', entity)) {
      items.push({ key: 'bulk', text: `Edit ${sel}`, iconProps: { iconName: 'Edit' }, onClick: () => setDialog('bulk') });
      if (def.fields.some(f => f.key === 'owner')) {
        items.push({ key: 'mine', text: 'Assign to me', iconProps: { iconName: 'Contact' }, onClick: () => {
          engine.bulkPatch(entity, selected.map(r => r.id!), { owner: engine.me }).then(() => notify(`Assigned ${sel} to you.`, MessageBarType.success)).catch(e => notify(errorText(e), MessageBarType.error));
        } });
      }
      if (entity === 'leads' && engine.can('import', entity)) {
        items.push({ key: 'rr', text: 'Round-robin assign', iconProps: { iconName: 'Sync' }, onClick: () => {
          engine.assignRoundRobin(selected.map(r => r.id!)).then(n => notify(`Distributed ${n} lead(s).`, MessageBarType.success)).catch(e => notify(errorText(e), MessageBarType.error));
        } });
      }
    }
    if (sel && engine.can('delete', entity)) {
      items.push({ key: 'del', text: `Delete ${sel}`, iconProps: { iconName: 'Delete' }, onClick: () => setDialog('delete') });
    }

    const far: ICommandBarItemProps[] = [
      ...(def.boardField ? [{ key: 'mode', text: mode === 'board' ? 'List' : 'Board', iconProps: { iconName: mode === 'board' ? 'List' : 'BacklogBoard' }, onClick: () => setMode(mode === 'board' ? 'list' : 'board') }] : []),
      { key: 'views', text: 'Views', iconProps: { iconName: 'ViewList' }, subMenuProps: { items: viewItems } },
      {
        key: 'more', text: '', iconOnly: true, ariaLabel: 'More', iconProps: { iconName: 'More' }, subMenuProps: {
          items: [
            { key: 'cols', text: 'Choose columns', iconProps: { iconName: 'TripleColumn' }, onClick: () => setDialog('cols') },
            { key: 'export', text: 'Export to CSV / Excel', iconProps: { iconName: 'ExcelDocument' }, onClick: doExport },
            ...(engine.can('import', entity) ? [{ key: 'import', text: 'Import from CSV', iconProps: { iconName: 'Upload' }, onClick: () => setDialog('import') }] : []),
            ...(entity === 'leads' && engine.can('import', entity) ? [{ key: 'rescore', text: 'Recalculate lead scores', iconProps: { iconName: 'Calculator' }, onClick: () => {
              engine.rescoreLeads().then(n => notify(`${n} lead score(s) updated.`, MessageBarType.success)).catch(e => notify(errorText(e), MessageBarType.error));
            } }] : []),
            { key: 'refresh', text: 'Refresh', iconProps: { iconName: 'Refresh' }, onClick: () => { engine.refresh(entity).catch(e => notify(errorText(e), MessageBarType.error)); } }
          ]
        }
      }
    ];

    // ---------------------------------------------------------------- board
    const boardField = def.fields.find(f => f.key === def.boardField);
    const board = (): React.ReactNode => {
      if (!boardField) { return null; }
      const lanes = choicesFor(boardField, settings);
      const extra = filtered.filter(r => lanes.indexOf(r[boardField.key]) < 0 && r[boardField.key]).map(r => r[boardField.key]);
      const allLanes = [...lanes, ...extra.filter((x, i) => extra.indexOf(x) === i)];
      const drop = (lane: string): void => {
        setOver(undefined);
        const id = dragId;
        const rec = id ? filtered.find(r => r.id === id) : undefined;
        if (!rec || rec[boardField.key] === lane || !engine.can('edit', entity)) { return; }
        const patch: Record<string, unknown> = { [boardField.key]: lane };
        if (entity === 'opportunities' && lane === LOST && !rec.lostReason) {
          open(entity, rec);
          notify('Choose "Lost" in the record to add a lost reason.', MessageBarType.info);
          return;
        }
        engine.patch(entity, id!, patch).then(() => notify(`Moved to ${lane}.`, MessageBarType.success)).catch(e => notify(errorText(e), MessageBarType.error));
      };
      const cardFields = def.fields.filter(f => f.grid && f.key !== 'title' && f.key !== boardField.key && f.type !== 'tags').slice(0, 3);
      return (
        <div className="crm-board">
          {allLanes.map(lane => {
            const cards = filtered.filter(r => (r[boardField.key] || '') === lane);
            return (
              <div key={lane} className={`crm-lane ${over === lane ? 'crm-lane-over' : ''}`}
                onDragOver={e => { e.preventDefault(); setOver(lane); }} onDragLeave={() => setOver(undefined)} onDrop={e => { e.preventDefault(); drop(lane); }}>
                <div className="crm-lane-head">
                  <Text styles={{ root: { fontWeight: 600 } }}>{lane}</Text>
                  <Text variant="small" className="crm-muted">{cards.length}{moneyField ? ` · ${fmtMoney(sum(cards, c => Number(c[moneyField.key]) || 0), currency)}` : ''}</Text>
                </div>
                {cards.slice(0, 200).map(r => (
                  <div key={r.id} className="crm-card-item" draggable={engine.can('edit', entity)} tabIndex={0} role="button" aria-label={`${r.title}, ${lane}`}
                    onDragStart={e => { e.dataTransfer.setData('text/plain', String(r.id)); setDragId(r.id); }} onDragEnd={() => setDragId(undefined)}
                    onClick={() => open(entity, r)} onKeyDown={e => { if (e.key === 'Enter') { open(entity, r); } }}>
                    <Text block styles={{ root: { fontWeight: 600 } }}>{r.title || r.quoteNumber || r.caseNumber}</Text>
                    {cardFields.map(f => { const v = renderValue(f, r, engine, currency); return v ? <div key={f.key} className="crm-card-line">{v}</div> : null; })}
                  </div>
                ))}
                {!cards.length && <Text variant="small" className="crm-muted" style={{ padding: 8 }}>{engine.can('edit', entity) ? 'Drop here' : 'Empty'}</Text>}
              </div>
            );
          })}
        </div>
      );
    };

    return (
      <Stack tokens={{ childrenGap: 8 }}>
        <CommandBar items={items} farItems={far} styles={{ root: { padding: 0 } }} />
        <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="end">
          <SearchBox placeholder={`Search ${def.plural.toLowerCase()}...`} value={q} onChange={(_, v) => setQ(v || '')} styles={{ root: { width: isNarrow ? '100%' : 240 } }} />
          {['opportunities', 'leads', 'cases', 'activities', 'demos'].indexOf(entity) >= 0 && (
            <Dropdown ariaLabel="Open or all" selectedKey={filters.__open || ALL} styles={{ root: { width: 120 } }}
              options={[{ key: ALL, text: 'All' }, { key: 'open', text: 'Open only' }]} onChange={(_, o) => setFilters({ ...filters, __open: String(o?.key) })} />
          )}
          {filterFields.map(f => {
            let opts: { key: string; text: string }[];
            if (f.type === 'user') { opts = [{ key: 'me', text: `My ${def.plural.toLowerCase()}` }, { key: 'none', text: 'Unassigned' }]; }
            else if (f.type === 'boolean') { opts = [{ key: 'yes', text: 'Yes' }, { key: 'no', text: 'No' }]; }
            else {
              const distinct = all.map(r => String(r[f.key] || '')).filter((v, i, a) => v && a.indexOf(v) === i);
              const base = f.type === 'choice' ? choicesFor(f, settings) : [];
              opts = [...base, ...distinct.filter(d => base.indexOf(d) < 0)].map(v => ({ key: v, text: v }));
            }
            return (
              <Dropdown key={f.key} ariaLabel={f.label} placeholder={f.label} selectedKey={filters[f.key] || ALL} styles={{ root: { width: 150 } }}
                options={[{ key: ALL, text: `${f.label}: all` }, ...opts]} onChange={(_, o) => setFilters({ ...filters, [f.key]: String(o?.key) })} />
            );
          })}
        </Stack>
        <Text variant="small" className="crm-muted">
          {filtered.length} of {all.length} {def.plural.toLowerCase()}
          {moneyField && filtered.length > 0 && <> · {moneyField.label} total {fmtMoney(sum(filtered, r => Number(r[moneyField.key]) || 0), currency)}</>}
          {sel > 0 && <> · {sel} selected</>}
        </Text>

        {mode === 'board' && def.boardField ? board() : (
          <ShimmeredDetailsList
            enableShimmer={loading && !all.length}
            items={filtered}
            columns={columns}
            selection={selection}
            selectionMode={engine.can('edit', entity) ? SelectionMode.multiple : SelectionMode.none}
            layoutMode={DetailsListLayoutMode.justified}
            constrainMode={ConstrainMode.unconstrained}
            onItemInvoked={(r: IRecord) => open(entity, r)}
            onRenderRow={(props, render) => props && render ? (
              <div onClick={e => { if (!(e.target as HTMLElement).closest('[data-selection-toggle],a,button,.ms-Check')) { open(entity, props.item); } }} style={{ cursor: 'pointer' }}>
                {render(props)}
              </div>
            ) : null}
            setKey={entity}
            ariaLabelForGrid={def.plural}
          />
        )}
        {!loading && filtered.length === 0 && (
          <Stack horizontalAlign="center" tokens={{ childrenGap: 8 }} className="crm-empty">
            <Text>{all.length ? 'No records match the current filters.' : `No ${def.plural.toLowerCase()} yet.`}</Text>
            {!all.length && engine.can('create', entity) && <PrimaryButton text={`Create the first ${def.singular.toLowerCase()}`} onClick={() => open(entity, {})} />}
          </Stack>
        )}

        {dialog === 'import' && <ImportDialog entity={entity} onClose={() => setDialog(undefined)} />}
        {dialog === 'bulk' && <BulkEditDialog entity={entity} ids={selected.map(r => r.id!)} onClose={() => setDialog(undefined)} />}
        {dialog === 'delete' && (
          <Dialog hidden={false} onDismiss={() => setDialog(undefined)} dialogContentProps={{ type: DialogType.normal, title: `Delete ${sel} record(s)?`, subText: 'They will be moved to the site recycle bin.' }}>
            <DialogFooter><PrimaryButton text="Delete" onClick={bulkDelete} /><DefaultButton text="Cancel" onClick={() => setDialog(undefined)} /></DialogFooter>
          </Dialog>
        )}
        {dialog === 'saveView' && (
          <Dialog hidden={false} onDismiss={() => setDialog(undefined)} dialogContentProps={{ type: DialogType.normal, title: 'Save view' }}>
            <TextField label="View name" value={viewName} onChange={(_, v) => setViewName(v || '')} autoFocus />
            <DialogFooter><PrimaryButton text="Save" disabled={!viewName} onClick={saveView} /><DefaultButton text="Cancel" onClick={() => setDialog(undefined)} /></DialogFooter>
          </Dialog>
        )}
        {dialog === 'cols' && (
          <Dialog hidden={false} onDismiss={() => setDialog(undefined)} dialogContentProps={{ type: DialogType.normal, title: 'Choose columns' }}>
            <Stack tokens={{ childrenGap: 6 }} style={{ maxHeight: 360, overflowY: 'auto' }}>
              {def.fields.filter(f => f.key !== 'title' && f.type !== 'json' && !f.hidden || f.key === 'grandTotal').map(f => (
                <Checkbox key={f.key} label={f.label} checked={visibleCols.indexOf(f.key) >= 0} onChange={(_, c) => {
                  const next = c ? [...visibleCols, f.key] : visibleCols.filter(k => k !== f.key);
                  setVisibleCols(next);
                  lsSet(colsKey, next);
                }} />
              ))}
            </Stack>
            <DialogFooter><PrimaryButton text="Done" onClick={() => setDialog(undefined)} /></DialogFooter>
          </Dialog>
        )}
      </Stack>
    );
  };
