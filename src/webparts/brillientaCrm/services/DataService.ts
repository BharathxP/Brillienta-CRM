/* eslint-disable @typescript-eslint/no-explicit-any */
import { isNetworkError, offlineStore } from './OfflineStore';
import { tr } from '../i18n';
import { SPHttpClient } from '@microsoft/sp-http';
import { SpRest } from './SpRest';
import {
  ENTITIES, EntityKey, IEntityDef, IFieldDef, IRecord, IUserRef, ListKey, SYSTEM_LISTS
} from '../models/model';
import { ICrmSettings, mergeSettings } from '../models/settings';

export interface IListStatus {
  key: ListKey;
  title: string;
  exists: boolean;
  missingFields: string[];
  url?: string;
}

export type Role = 'admin' | 'manager' | 'sales' | 'support' | 'reader';

/** Builds the runtime entity definitions: built-ins + custom fields + per-site list names. */
export function buildEntityDefs(settings: ICrmSettings, listNames: Partial<Record<ListKey, string>> = {}): IEntityDef[] {
  return ENTITIES.map(e => {
    const custom: IFieldDef[] = settings.customFields.filter(c => c.entity === e.key).map(c => ({
      key: `cf_${c.key}`, column: `cf${c.key}`, label: c.label, type: c.type, choices: c.choices,
      section: 'Additional information', grid: !!c.grid, filter: c.type === 'choice', custom: true
    }));
    // Labels in the user's language (column names and keys never change).
    return {
      ...e, singular: tr(e.singular), plural: tr(e.plural), description: tr(e.description), listTitle: listNames[e.key] || e.listTitle,
      related: e.related?.map(r => ({ ...r, label: tr(r.label) })),
      fields: [...e.fields, ...custom].map(f => ({ ...f, label: tr(f.label), section: f.section ? tr(f.section) : f.section }))
    };
  });
}

function spFieldXml(fd: IFieldDef): string {
  const esc = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const base = `Name="${fd.column}" StaticName="${fd.column}" DisplayName="${esc(fd.label)}"`;
  switch (fd.type) {
    case 'note': case 'json': case 'url':
      return `<Field Type="Note" ${base} NumLines="6" RichText="FALSE" UnlimitedLengthInDocumentLibrary="TRUE" />`;
    case 'choice': {
      const ch = (fd.choices || []).map(c => `<CHOICE>${esc(c)}</CHOICE>`).join('');
      const def = fd.defaultValue ? `<Default>${esc(String(fd.defaultValue))}</Default>` : '';
      return `<Field Type="Choice" ${base} Format="Dropdown" FillInChoice="TRUE">${def}<CHOICES>${ch}</CHOICES></Field>`;
    }
    case 'number': case 'percent':
      return `<Field Type="Number" ${base} />`;
    case 'lookup':
      return `<Field Type="Number" ${base} Decimals="0" Indexed="TRUE" />`;
    case 'currency':
      return `<Field Type="Currency" ${base} Decimals="2" />`;
    case 'date':
      return `<Field Type="DateTime" ${base} Format="DateOnly" />`;
    case 'datetime':
      return `<Field Type="DateTime" ${base} Format="DateTime" />`;
    case 'boolean':
      return `<Field Type="Boolean" ${base}><Default>${fd.defaultValue ? 1 : 0}</Default></Field>`;
    case 'user':
      return `<Field Type="User" ${base} UserSelectionMode="PeopleOnly" UserSelectionScope="0" List="UserInfo" ShowField="ImnName" />`;
    default:
      return `<Field Type="Text" ${base} MaxLength="255" />`;
  }
}

const userSel = (c: string): string => `${c}/Id,${c}/Title,${c}/EMail`;
const toUser = (u: any): IUserRef | undefined => (u && u.Id ? { id: u.Id, title: u.Title, email: u.EMail } : undefined);

export class DataService {
  public readonly rest: SpRest;
  public defs: IEntityDef[] = ENTITIES;
  public settings: ICrmSettings = mergeSettings(undefined);
  private settingsItemId?: number;
  private me?: IUserRef;

  constructor(client: SPHttpClient, webUrl: string, public listNames: Partial<Record<ListKey, string>> = {}) {
    this.rest = new SpRest(client, webUrl);
  }

  public def(key: EntityKey): IEntityDef {
    return this.defs.find(d => d.key === key)!;
  }

  public listTitle(key: ListKey): string {
    if (key === 'settings' || key === 'audit') {
      return this.listNames[key] || SYSTEM_LISTS[key].listTitle;
    }
    return this.def(key).listTitle;
  }

  private items(key: ListKey): string {
    return `${this.rest.listUrl(this.listTitle(key))}/items`;
  }

  // ------------------------------------------------------------------ setup

  /** One request that returns all lists and their columns. */
  public async checkLists(): Promise<IListStatus[]> {
    const res: any = await this.rest.get(`/_api/web/lists?$select=Title,RootFolder/ServerRelativeUrl,Fields/InternalName,Fields/Hidden&$expand=RootFolder,Fields&$filter=Hidden eq false`);
    const lists: any[] = res.value || [];
    const status = (key: ListKey, fields: IFieldDef[]): IListStatus => {
      const title = this.listTitle(key);
      const l = lists.find(x => String(x.Title).toLowerCase() === title.toLowerCase());
      if (!l) {
        return { key, title, exists: false, missingFields: fields.filter(x => x.column !== 'Title').map(x => x.column) };
      }
      const names = new Set<string>((l.Fields || []).map((x: any) => x.InternalName));
      return {
        key, title, exists: true, url: l.RootFolder?.ServerRelativeUrl,
        missingFields: fields.filter(x => x.column !== 'Title' && !names.has(x.column)).map(x => x.column)
      };
    };
    return [
      ...this.defs.map(d => status(d.key, d.fields)),
      status('settings', SYSTEM_LISTS.settings.fields),
      status('audit', SYSTEM_LISTS.audit.fields)
    ];
  }

  public async provision(onProgress: (m: string) => void = () => undefined): Promise<void> {
    const statuses = await this.checkLists();
    const all: { key: ListKey; description: string; fields: IFieldDef[] }[] = [
      { key: 'settings', ...SYSTEM_LISTS.settings },
      { key: 'audit', ...SYSTEM_LISTS.audit },
      ...this.defs.map(d => ({ key: d.key as ListKey, description: d.description, fields: d.fields }))
    ];
    for (const l of all) {
      const st = statuses.find(s => s.key === l.key)!;
      const title = this.listTitle(l.key);
      if (!st.exists) {
        onProgress(`Creating list "${title}"...`);
        await this.rest.post('/_api/web/lists', { Title: title, Description: l.description, BaseTemplate: 100, ContentTypesEnabled: false });
        if (l.key === 'audit' || l.key === 'settings') {
          // keep system lists out of navigation/search clutter
          await this.rest.merge(this.rest.listUrl(title), { OnQuickLaunch: false, NoCrawl: true }).catch(() => undefined);
        }
      }
      for (const fd of l.fields) {
        if (st.missingFields.indexOf(fd.column) < 0) {
          continue;
        }
        onProgress(`Adding column "${fd.label}" to "${title}"...`);
        await this.rest.post(`${this.rest.listUrl(title)}/fields/createfieldasxml`, {
          parameters: { SchemaXml: spFieldXml(fd), Options: fd.hidden ? 8 : (fd.grid ? 24 : 8) }
        });
      }
      for (const fd of l.fields.filter(x => x.type === 'choice' && x.choices && st.missingFields.indexOf(x.column) < 0)) {
        await this.syncChoices(title, fd.column, fd.choices!).catch(() => undefined);
      }
    }
    onProgress('All lists are ready.');
  }

  /** Adds any missing choices to an existing choice column. */
  public async syncChoices(listTitle: string, column: string, wanted: string[]): Promise<void> {
    const url = `${this.rest.listUrl(listTitle)}/fields/getbyinternalnameortitle('${column}')`;
    const existing: any = await this.rest.get(`${url}?$select=Choices,TypeAsString`);
    if (existing.TypeAsString !== 'Choice') {
      return;
    }
    const current: string[] = existing.Choices || [];
    const merged = [...current, ...wanted.filter(c => current.indexOf(c) < 0)];
    if (merged.length !== current.length) {
      await this.rest.postVerbose(url, {
        __metadata: { type: 'SP.FieldChoice' },
        Choices: { __metadata: { type: 'Collection(Edm.String)' }, results: merged },
        FillInChoice: true
      }, { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' });
    }
  }

  // ------------------------------------------------------------------ settings

  /** Uses settings saved on this device (offline start). */
  public useSettings(s: ICrmSettings): void {
    this.settings = mergeSettings(s);
    this.defs = buildEntityDefs(this.settings, this.listNames);
  }

  public async loadSettings(): Promise<ICrmSettings> {
    const res: any = await this.rest.get(`${this.items('settings')}?$select=Id,Title,Value&$filter=Title eq 'settings'&$top=1`);
    const item = (res.value || [])[0];
    let saved: any;
    if (item) {
      this.settingsItemId = item.Id;
      try {
        saved = JSON.parse(item.Value || '{}');
      } catch {
        saved = undefined;
      }
    }
    this.settings = mergeSettings(saved);
    this.defs = buildEntityDefs(this.settings, this.listNames);
    if (!item) {
      // first run: persist install date for the trial clock
      await this.saveSettings(this.settings).catch(() => undefined);
    }
    return this.settings;
  }

  public async saveSettings(s: ICrmSettings): Promise<void> {
    const body = { Title: 'settings', Value: JSON.stringify(s) };
    if (this.settingsItemId) {
      await this.rest.merge(`${this.items('settings')}(${this.settingsItemId})`, body);
    } else {
      const created: any = await this.rest.post(this.items('settings'), body);
      this.settingsItemId = created.Id;
    }
    this.settings = s;
    this.defs = buildEntityDefs(s, this.listNames);
  }

  // ------------------------------------------------------------------ users & roles

  public async currentUser(): Promise<IUserRef> {
    if (!this.me) {
      const u: any = await this.rest.get('/_api/web/currentuser?$select=Id,Title,Email');
      this.me = { id: u.Id, title: u.Title, email: u.Email };
    }
    return this.me;
  }

  public async resolveRole(isSiteAdmin: boolean): Promise<Role> {
    const g = this.settings.groups;
    let groups: string[] = [];
    try {
      const res: any = await this.rest.get('/_api/web/currentuser/groups?$select=Title');
      groups = (res.value || []).map((x: any) => x.Title);
    } catch {
      /* ignore */
    }
    if (isSiteAdmin || groups.indexOf(g.admins) >= 0) { return 'admin'; }
    if (groups.indexOf(g.managers) >= 0) { return 'manager'; }
    if (groups.indexOf(g.sales) >= 0) { return 'sales'; }
    if (groups.indexOf(g.support) >= 0) { return 'support'; }
    if (groups.indexOf(g.readers) >= 0) { return 'reader'; }
    // No CRM groups configured for this user: fall back to SharePoint permissions.
    try {
      const perms: any = await this.rest.get(`${this.rest.listUrl(this.def('leads').listTitle)}/EffectiveBasePermissions`);
      const low = Number(perms.Low ?? perms.EffectiveBasePermissions?.Low ?? 0);
      // EditListItems = 0x4
      return (low & 0x4) === 0x4 ? 'manager' : 'reader';
    } catch {
      return 'reader';
    }
  }

  /** Is the current user an owner/admin of the CRM site (which may differ from the host site in Teams)? */
  public async isCrmSiteAdmin(): Promise<boolean> {
    try {
      const [u, perms]: any[] = await Promise.all([
        this.rest.get('/_api/web/currentuser?$select=IsSiteAdmin'),
        this.rest.get('/_api/web/EffectiveBasePermissions')
      ]);
      const low = Number(perms.Low ?? perms.EffectiveBasePermissions?.Low ?? 0);
      // ManageWeb = 0x40000000 (low word)
      return !!u.IsSiteAdmin || (low & 0x40000000) === 0x40000000;
    } catch {
      return false;
    }
  }

  /** Reads the tenant-wide default CRM site (tenant property / storage entity). */
  public async getTenantCrmSite(key: string): Promise<string> {
    try {
      const r: any = await this.rest.get(`/_api/web/GetStorageEntity('${key}')`);
      return String(r?.Value || '');
    } catch {
      return '';
    }
  }

  /** Writes the tenant-wide default CRM site. Requires admin rights on the tenant app catalog. */
  public async setTenantCrmSite(key: string, value: string): Promise<void> {
    const t: any = await this.rest.get('/_api/SP_TenantSettings_Current');
    const catalog = t?.CorporateCatalogUrl;
    if (!catalog) {
      throw new Error('No tenant app catalog found. Use PowerShell: Set-PnPStorageEntity -Key "' + key + '" -Value "' + value + '" -Scope Tenant');
    }
    await this.rest.post(`${catalog}/_api/web/SetStorageEntity`, {
      key, value, description: 'Brillienta CRM: site that holds the CRM lists (used by Teams and other sites)', comments: 'Set from Brillienta CRM settings'
    });
  }

  public async searchPeople(query: string): Promise<{ key: string; text: string; email: string }[]> {
    const res: any = await this.rest.post('/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser', {
      queryParams: {
        QueryString: query, MaximumEntitySuggestions: 10, AllowEmailAddresses: false, AllowMultipleEntities: false,
        PrincipalSource: 15, PrincipalType: 1, AllUrlZones: false
      }
    });
    const arr: any[] = JSON.parse(res.value || '[]');
    return arr.map(x => ({ key: x.Key, text: x.DisplayText, email: x.EntityData?.Email || '' }));
  }

  public async ensureUser(logonName: string): Promise<IUserRef> {
    const u: any = await this.rest.post('/_api/web/ensureuser', { logonName });
    return { id: u.Id, title: u.Title, email: u.Email };
  }

  /** Creates the CRM SharePoint groups and grants them access to the CRM lists. */
  public async applySecurity(onProgress: (m: string) => void = () => undefined): Promise<Record<string, number>> {
    const g = this.settings.groups;
    const ids: Record<string, number> = {};
    for (const name of [g.admins, g.managers, g.sales, g.support, g.readers]) {
      try {
        const existing: any = await this.rest.get(`/_api/web/sitegroups/getbyname('${encodeURIComponent(name.replace(/'/g, "''"))}')?$select=Id`);
        ids[name] = existing.Id;
      } catch {
        onProgress(`Creating group "${name}"...`);
        const created: any = await this.rest.post('/_api/web/sitegroups', { Title: name, Description: 'Brillienta CRM role group' });
        ids[name] = created.Id;
      }
    }
    const roleDef = async (type: number): Promise<number> => ((await this.rest.get(`/_api/web/roledefinitions/getbytype(${type})?$select=Id`)) as any).Id;
    const [readId, contributeId, fullId] = await Promise.all([roleDef(2), roleDef(3), roleDef(5)]);
    const grant = async (list: string, principal: number, role: number): Promise<void> => {
      await this.rest.post(`${this.rest.listUrl(list)}/roleassignments/addroleassignment(principalid=${principal},roledefid=${role})`);
    };
    const keys: ListKey[] = [...this.defs.map(d => d.key as ListKey), 'settings', 'audit'];
    for (const key of keys) {
      const list = this.listTitle(key);
      onProgress(`Securing "${list}"...`);
      await this.rest.post(`${this.rest.listUrl(list)}/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=true)`);
      await grant(list, ids[g.admins], fullId);
      const writeRole = key === 'settings' ? readId : contributeId;
      await grant(list, ids[g.managers], writeRole);
      await grant(list, ids[g.sales], key === 'products' ? readId : writeRole);
      await grant(list, ids[g.support], writeRole);
      await grant(list, ids[g.readers], key === 'audit' ? contributeId : readId);
    }
    onProgress('Security applied.');
    return ids;
  }

  // ------------------------------------------------------------------ item-level security

  private secCache?: { groups: Record<string, number>; read: number; contribute: number; full: number };

  private async securityIds(): Promise<{ groups: Record<string, number>; read: number; contribute: number; full: number }> {
    if (this.secCache) { return this.secCache; }
    const g = this.settings.groups;
    const groups: Record<string, number> = {};
    for (const [k, name] of [['admins', g.admins], ['managers', g.managers], ['support', g.support], ['readers', g.readers]] as [string, string][]) {
      try { groups[k] = ((await this.rest.get(`/_api/web/sitegroups/getbyname('${encodeURIComponent(name.replace(/'/g, "''"))}')?$select=Id`)) as any).Id; } catch { /* group not created yet */ }
    }
    const roleDef = async (type: number): Promise<number> => ((await this.rest.get(`/_api/web/roledefinitions/getbytype(${type})?$select=Id`)) as any).Id;
    const [read, contribute, full] = await Promise.all([roleDef(2), roleDef(3), roleDef(5)]);
    this.secCache = { groups, read, contribute, full };
    return this.secCache;
  }

  /**
   * Gives one record its own permissions: CRM admins (full), managers (edit), support (edit on cases, read
   * otherwise), readers (read), the owner (edit) and the given extra users. The Sales group gets no access,
   * so reps reach only records shared with them. Needs Manage Permissions (CRM administrators).
   */
  public async secureItem(key: EntityKey, id: number, ownerId: number, extra: { userId: number; edit: boolean }[]): Promise<void> {
    const ids = await this.securityIds();
    const item = `${this.items(key)}(${id})`;
    await this.rest.post(`${item}/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)`);
    const grant = async (principal: number | undefined, role: number): Promise<void> => {
      if (principal) { await this.rest.post(`${item}/roleassignments/addroleassignment(principalid=${principal},roledefid=${role})`); }
    };
    await grant(ids.groups.admins, ids.full);
    await grant(ids.groups.managers, ids.contribute);
    await grant(ids.groups.support, key === 'cases' ? ids.contribute : ids.read);
    await grant(ids.groups.readers, ids.read);
    await grant(ownerId, ids.contribute);
    for (const x of extra) { if (x.userId !== ownerId) { await grant(x.userId, x.edit ? ids.contribute : ids.read); } }
    // Remember who it is secured for without changing Modified / Modified by.
    await this.rest.post(`${item}/ValidateUpdateListItem`, { formValues: [{ FieldName: 'SecuredOwner', FieldValue: String(ownerId) }], bNewDocumentUpdate: true });
  }

  /** Back to the list's permissions (when item-level security is switched off). */
  public async unsecureItem(key: EntityKey, id: number): Promise<void> {
    const item = `${this.items(key)}(${id})`;
    await this.rest.post(`${item}/resetroleinheritance`);
    await this.rest.post(`${item}/ValidateUpdateListItem`, { formValues: [{ FieldName: 'SecuredOwner', FieldValue: '' }], bNewDocumentUpdate: true });
  }

  // ------------------------------------------------------------------ records

  private selectFor(d: IEntityDef): { select: string; expand: string } {
    const sel = ['Id', 'Title', 'Created', 'Modified', userSel('Author'), userSel('Editor')];
    const exp = ['Author', 'Editor'];
    d.fields.forEach(fd => {
      if (fd.column === 'Title') { return; }
      if (fd.type === 'user') {
        sel.push(userSel(fd.column));
        exp.push(fd.column);
      } else {
        sel.push(fd.column);
      }
    });
    return { select: sel.join(','), expand: exp.join(',') };
  }

  public fromSp(d: IEntityDef, i: any): IRecord {
    const r: IRecord = {
      id: i.Id, created: i.Created, modified: i.Modified, author: toUser(i.Author), editor: toUser(i.Editor)
    };
    d.fields.forEach(fd => {
      const v = i[fd.column];
      switch (fd.type) {
        case 'user': r[fd.key] = toUser(v); break;
        case 'number': case 'currency': case 'percent': case 'lookup':
          r[fd.key] = v === null || v === undefined || v === '' ? undefined : Number(v); break;
        case 'boolean': r[fd.key] = !!v; break;
        case 'json':
          try { r[fd.key] = v ? JSON.parse(v) : undefined; } catch { r[fd.key] = undefined; }
          break;
        case 'tags': r[fd.key] = v ? String(v).split(',').map(s => s.trim()).filter(Boolean) : []; break;
        case 'date': case 'datetime': r[fd.key] = v || undefined; break;
        default: r[fd.key] = v === null || v === undefined ? '' : String(v);
      }
    });
    return r;
  }

  public toSp(d: IEntityDef, r: IRecord): Record<string, any> {
    const body: Record<string, any> = {};
    d.fields.forEach(fd => {
      if (!(fd.key in r)) { return; }
      const v = r[fd.key];
      switch (fd.type) {
        case 'user': body[`${fd.column}Id`] = v?.id ?? null; break;
        case 'number': case 'currency': case 'percent': case 'lookup':
          body[fd.column] = v === undefined || v === null || v === '' || isNaN(Number(v)) ? null : Number(v); break;
        case 'boolean': body[fd.column] = !!v; break;
        case 'json': body[fd.column] = v === undefined ? null : JSON.stringify(v); break;
        case 'tags': body[fd.column] = Array.isArray(v) ? v.join(', ') : (v || ''); break;
        case 'date': case 'datetime': body[fd.column] = v ? new Date(v).toISOString() : null; break;
        case 'text': case 'email': case 'phone': body[fd.column] = v === undefined || v === null ? '' : String(v).slice(0, 255); break;
        default: body[fd.column] = v === undefined || v === null ? '' : String(v);
      }
    });
    return body;
  }

  /** Set while lists are served from the offline copy: when that copy was saved. */
  public offlineSince?: string;
  public onOfflineChange?: () => void;

  public get cacheScope(): string { return this.rest.webUrl.toLowerCase(); }

  public async getAll(key: EntityKey): Promise<IRecord[]> {
    const d = this.def(key);
    const { select, expand } = this.selectFor(d);
    const cacheKey = `${this.cacheScope}|list|${key}`;
    try {
      const items = await this.rest.getAll(`${this.items(key)}?$select=${select}&$expand=${expand}&$top=2000&$orderby=Id desc`);
      const rows = items.map(i => this.fromSp(d, i));
      offlineStore.put(cacheKey, { at: new Date().toISOString(), rows }).catch(() => undefined);
      if (this.offlineSince) { this.offlineSince = undefined; this.onOfflineChange?.(); }
      return rows;
    } catch (e) {
      if (!isNetworkError(e)) { throw e; }
      const cached = await offlineStore.get<{ at: string; rows: IRecord[] }>(cacheKey);
      if (!cached) { throw new Error('You are offline and this list has not been opened on this device before.'); }
      if (!this.offlineSince || cached.at < this.offlineSince) { this.offlineSince = cached.at; this.onOfflineChange?.(); }
      return cached.rows;
    }
  }

  public async getOne(key: EntityKey, id: number): Promise<IRecord> {
    const d = this.def(key);
    const { select, expand } = this.selectFor(d);
    return this.fromSp(d, await this.rest.get(`${this.items(key)}(${id})?$select=${select}&$expand=${expand}`));
  }

  /** Creates or updates; only keys present on the record are written. Returns the item id. */
  public async save(key: EntityKey, r: IRecord): Promise<number> {
    const body = this.toSp(this.def(key), r);
    if (r.id) {
      await this.rest.merge(`${this.items(key)}(${r.id})`, body);
      return r.id;
    }
    const created: any = await this.rest.post(this.items(key), body);
    return created.Id ?? created.ID;
  }

  /** Moves to the recycle bin. */
  public remove(key: EntityKey, id: number): Promise<void> {
    return this.rest.post(`${this.items(key)}(${id})/recycle()`);
  }

  // ------------------------------------------------------------------ audit

  public async writeAudit(action: string, entity: string, recordId: number, recordTitle: string, changes: unknown): Promise<void> {
    if (!this.settings.auditEnabled) { return; }
    await this.rest.post(this.items('audit'), {
      Title: action.slice(0, 255), EntityType: entity, RecordId: recordId, RecordTitle: (recordTitle || '').slice(0, 255),
      Changes: changes ? JSON.stringify(changes) : ''
    }).catch(() => undefined);
  }

  /** Permanently deletes the audit entries of a record (right to erasure). */
  public async purgeAudit(entity: string, recordId: number): Promise<number> {
    const rows = await this.getAudit(entity, recordId, 5000);
    for (const r of rows) { await this.rest.delete(`${this.items('audit')}(${r.id})`); }
    return rows.length;
  }

  public async getAudit(entity?: string, recordId?: number, top: number = 500): Promise<IRecord[]> {
    const filters: string[] = [];
    if (entity) { filters.push(`EntityType eq '${entity}'`); }
    if (recordId) { filters.push(`RecordId eq ${recordId}`); }
    const res: any = await this.rest.get(`${this.items('audit')}?$select=Id,Title,EntityType,RecordId,RecordTitle,Changes,Created,${userSel('Author')}&$expand=Author&$orderby=Id desc&$top=${top}${filters.length ? `&$filter=${filters.join(' and ')}` : ''}`);
    return (res.value || []).map((i: any) => ({
      id: i.Id, title: i.Title, entityType: i.EntityType, recordId: i.RecordId, recordTitle: i.RecordTitle,
      changes: (() => { try { return i.Changes ? JSON.parse(i.Changes) : undefined; } catch { return undefined; } })(),
      created: i.Created, author: toUser(i.Author)
    }));
  }
}
