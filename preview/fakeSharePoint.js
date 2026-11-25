// In-memory stand-in for the SharePoint REST API, for the local preview and the end-to-end tests (npm run preview,
// npm run e2e). It implements every call the CRM makes and is strict where SharePoint is strict, so mistakes surface
// here first: writing or selecting a column the list does not have is a 400, an unknown list is a 404, text columns
// take at most 255 characters, readers cannot write, and items with their own permissions are hidden from people
// they are not shared with.
//
// State (lists, items, groups) is kept in localStorage, so a test can sign in as someone else (?as=chen) and see the
// same CRM. ?reset=1 starts from an empty site.

const KEY = 'brillienta-crm-preview';
export const WEB = 'https://contoso.sharepoint.com/sites/crm';
const HOST = 'https://contoso.sharepoint.com';

/** People in the preview tenant. `role` decides which CRM group they belong to once the groups exist. */
export const USERS = {
  asha: { Id: 7, Title: 'Asha Rao', Email: 'asha.rao@contoso.com', role: 'admin' },
  ben: { Id: 8, Title: 'Ben Ortiz', Email: 'ben.ortiz@contoso.com', role: 'managers' },
  chen: { Id: 9, Title: 'Chen Wei', Email: 'chen.wei@contoso.com', role: 'sales' },
  dana: { Id: 10, Title: 'Dana Fox', Email: 'dana.fox@contoso.com', role: 'support' },
  eli: { Id: 11, Title: 'Eli Park', Email: 'eli.park@contoso.com', role: 'readers' },
  maya: { Id: 12, Title: 'Maya Singh', Email: 'maya.singh@contoso.com', role: 'sales' }
};
const byId = (id) => Object.values(USERS).find((u) => u.Id === Number(id));
const login = (u) => `i:0#.f|membership|${u.Email}`;

const BUILTIN = { Id: 'Counter', ID: 'Counter', Title: 'Text', Created: 'DateTime', Modified: 'DateTime', Author: 'User', Editor: 'User', Attachments: 'Attachments', ContentTypeId: 'ContentTypeId', GUID: 'Guid' };
const READ_ONLY = new Set(['Id', 'ID', 'Created', 'Modified', 'Author', 'Editor', 'AuthorId', 'EditorId', 'GUID', 'ContentTypeId']);

let state = load();
function blank() {
  return { lists: {}, nextItem: 1, groups: {}, nextGroup: 500, storage: {}, log: [] };
}
function load() {
  try {
    if (new URLSearchParams(window.location.search).get('reset')) { localStorage.removeItem(KEY); }
    const raw = localStorage.getItem(KEY);
    return raw ? JSON.parse(raw) : blank();
  } catch {
    return blank();
  }
}
function save() {
  try { localStorage.setItem(KEY, JSON.stringify(state)); } catch { /* storage full or blocked: keep in memory */ }
}
/** For tests: the raw state. */
export function previewState() { return state; }

const nowIso = () => new Date().toISOString();
function response(status, body) {
  const text = body === undefined ? '' : typeof body === 'string' ? body : JSON.stringify(body);
  return {
    ok: status >= 200 && status < 300, status, statusText: String(status),
    headers: { get: () => null },
    text: async () => text, json: async () => JSON.parse(text), blob: async () => new Blob([text])
  };
}
const err = (status, message) => response(status, { 'odata.error': { code: String(status), message: { lang: 'en-US', value: message } } });

// ------------------------------------------------------------------ OData helpers

function query(url) {
  const q = {};
  const i = url.indexOf('?');
  if (i < 0) { return q; }
  url.substring(i + 1).split('&').forEach((pair) => {
    const j = pair.indexOf('=');
    if (j > 0) { q[decodeURIComponent(pair.substring(0, j))] = decodeURIComponent(pair.substring(j + 1).replace(/\+/g, ' ')); }
  });
  return q;
}

/** eq/ne/lt/le/gt/ge on strings, numbers, booleans, null and datetime'…', joined by and/or with parentheses. */
function parseFilter(src, fieldOk) {
  if (!src) { return () => true; }
  const tokens = [];
  const re = /\s*(\(|\)|'(?:[^']|'')*'|datetime'[^']*'|[\w/.:-]+)/gy;
  let m;
  while ((m = re.exec(src))) { tokens.push(m[1]); }
  let i = 0;
  const literal = (t) => {
    if (/^datetime'/.test(t)) { return { date: Date.parse(t.slice(9, -1)) }; }
    if (/^'/.test(t)) { return t.slice(1, -1).replace(/''/g, "'"); }
    if (t === 'true' || t === 'false') { return t === 'true'; }
    if (t === 'null') { return null; }
    return /^-?\d+(\.\d+)?$/.test(t) ? Number(t) : t;
  };
  const compare = (field, op, value) => {
    fieldOk(field.split('/')[0]);
    return (row) => {
      const v = field.split('/').reduce((o, k) => (o == null ? o : o[k]), row);
      if (value && typeof value === 'object' && 'date' in value) {
        const x = Date.parse(v);
        if (isNaN(x)) { return op === 'ne'; }
        return { eq: x === value.date, ne: x !== value.date, lt: x < value.date, le: x <= value.date, gt: x > value.date, ge: x >= value.date }[op];
      }
      if (value === null) { return op === 'eq' ? v == null || v === '' : !(v == null || v === ''); }
      if (op === 'eq' || op === 'ne') {
        const same = String(v) === String(value) || (typeof value === 'boolean' && !!v === value);
        return op === 'eq' ? same : !same;
      }
      const a = typeof value === 'number' ? Number(v) : String(v);
      return { lt: a < value, le: a <= value, gt: a > value, ge: a >= value }[op];
    };
  };
  const primary = () => {
    if (tokens[i] === '(') { i++; const e = orExpr(); i++; return e; }
    const field = tokens[i++];
    const op = (tokens[i++] || '').toLowerCase();
    const value = literal(tokens[i++] || '');
    if (['eq', 'ne', 'lt', 'le', 'gt', 'ge'].indexOf(op) < 0) { throw new Error(`Unsupported $filter near '${field} ${op}'`); }
    return compare(field, op, value);
  };
  const andExpr = () => { const parts = [primary()]; while ((tokens[i] || '').toLowerCase() === 'and') { i++; parts.push(primary()); } return (r) => parts.every((t) => t(r)); };
  const orExpr = () => { const parts = [andExpr()]; while ((tokens[i] || '').toLowerCase() === 'or') { i++; parts.push(andExpr()); } return (r) => parts.some((t) => t(r)); };
  return orExpr();
}

// ------------------------------------------------------------------ permissions

const PERM = {
  full: { High: '2147483647', Low: '4294967295' },
  edit: { High: '432', Low: '1011030767' }, // Contribute: includes EditListItems (0x4), not ManageWeb
  read: { High: '176', Low: '138612833' } // Read: ViewListItems (0x1), no EditListItems
};
function currentUser() {
  const as = new URLSearchParams(window.location.search).get('as') || 'asha';
  return USERS[as] || USERS.asha;
}
/** Groups the person is in: the CRM group for their role, once an admin has created it. */
function groupsOf(user, settingsGroups) {
  const g = settingsGroups || {};
  const name = g[user.role];
  return Object.values(state.groups).filter((x) => x.Title === name || (x.members || []).indexOf(user.Id) >= 0);
}
function settingsGroups() {
  const l = state.lists.BrillCRM_Settings;
  const row = l && l.items.find((x) => x.Title === 'settings');
  try { return row ? (JSON.parse(row.Value || '{}').groups || defaultGroups()) : defaultGroups(); } catch { return defaultGroups(); }
}
function defaultGroups() {
  return { admins: 'BrillCRM_Administrators', managers: 'BrillCRM_SalesManagers', sales: 'BrillCRM_Sales', support: 'BrillCRM_Support', readers: 'BrillCRM_Readers' };
}
function level(user) {
  if (user.role === 'admin') { return 'full'; }
  if (user.role === 'readers') { return 'read'; }
  return 'edit';
}
/** Can the person see this item? Items with their own permissions are visible to their principals only. */
function canSee(user, row) {
  if (!row.__unique || user.role === 'admin') { return true; }
  const mine = [user.Id, ...groupsOf(user, settingsGroups()).map((g) => g.Id)];
  return (row.__principals || []).some((p) => mine.indexOf(p) >= 0);
}

// ------------------------------------------------------------------ lists and fields

function fieldXml(xml) {
  const attr = (n) => { const m = new RegExp(`\\b${n}="([^"]*)"`).exec(xml); return m ? m[1] : ''; };
  const choices = [];
  const re = /<CHOICE>([^<]*)<\/CHOICE>/g;
  let m;
  while ((m = re.exec(xml))) { choices.push(m[1].replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')); }
  return { name: attr('Name'), type: attr('Type'), choices };
}
function typeOf(list, name) {
  if (BUILTIN[name]) { return BUILTIN[name]; }
  const f = list.fields[name];
  return f ? f.type : undefined;
}
/** A writable property name → its column, or undefined (User columns are written as <Name>Id). */
function columnForWrite(list, key) {
  if (list.fields[key] && list.fields[key].type !== 'User') { return key; }
  if (key === 'Title') { return key; }
  const m = /^(.*)Id$/.exec(key);
  if (m && list.fields[m[1]] && list.fields[m[1]].type === 'User') { return key; }
  return undefined;
}
function validate(list, title, data) {
  for (const k of Object.keys(data)) {
    if (k === '__metadata') { continue; }
    if (READ_ONLY.has(k)) { return `The property '${k}' cannot be set.`; }
    const col = columnForWrite(list, k);
    if (!col) { return `The property '${k}' does not exist on type 'SP.Data.${title}ListItem'. Make sure to only use property names that are defined by the type.`; }
    const type = k === 'Title' ? 'Text' : (list.fields[k] || {}).type;
    const v = data[k];
    if (v === null || v === undefined) { continue; }
    if (type === 'Text' && String(v).length > 255) { return `The value of '${k}' is too long (${String(v).length} > 255 characters).`; }
    if ((type === 'Number' || type === 'Currency') && v !== '' && isNaN(Number(v))) { return `Cannot convert '${v}' to a number for '${k}'.`; }
    if (type === 'DateTime' && v !== '' && isNaN(Date.parse(v))) { return `Cannot convert '${v}' to a date for '${k}'.`; }
  }
  return undefined;
}
function project(list, row, select, expand) {
  const out = {};
  const ex = new Set((expand || '').split(',').map((s) => s.trim()).filter(Boolean));
  const fields = select ? select.split(',').map((s) => s.trim()).filter(Boolean) : ['Id', 'Title', 'Created', 'Modified', ...Object.keys(list.fields)];
  for (const f of fields) {
    const [name, sub] = f.split('/');
    const type = typeOf(list, name);
    if (type === 'User') {
      if (!ex.has(name)) { if (!sub) { out[name + 'Id'] = row[name + 'Id'] ?? null; } continue; }
      const u = byId(row[name + 'Id']);
      const obj = out[name] || (u ? {} : null);
      if (u && sub) { obj[sub] = sub === 'Id' ? u.Id : sub === 'Title' ? u.Title : sub === 'EMail' ? u.Email : null; }
      out[name] = obj;
      continue;
    }
    out[name] = row[name] === undefined ? null : row[name];
  }
  out.Id = row.Id;
  out.ID = row.Id;
  return out;
}

// ------------------------------------------------------------------ request handling

function handle(method, url, bodyText, headers) {
  const user = currentUser();
  const verb = ((headers && (headers['X-HTTP-Method'] || headers['x-http-method'])) || method).toUpperCase();
  const body = bodyText ? JSON.parse(bodyText) : undefined;
  const path = decodeURIComponent(url.replace(/^https?:\/\/[^/]+/, '').split('?')[0]);
  const q = query(url);
  const write = verb !== 'GET';
  if (write) { state.log.push(`${verb} ${path}`); }

  // ---- tenant and web
  if (/\/_api\/SP_TenantSettings_Current$/.test(path)) { return response(200, { CorporateCatalogUrl: `${HOST}/sites/appcatalog` }); }
  const se = /\/_api\/web\/GetStorageEntity\('([^']+)'\)$/.exec(path);
  if (se) { return response(200, state.storage[se[1]] ? { Value: state.storage[se[1]] } : { 'odata.null': true }); }
  if (/\/_api\/web\/SetStorageEntity$/.test(path)) {
    if (user.role !== 'admin') { return err(403, 'Access is denied. (Exception from HRESULT: 0x80070005 (E_ACCESSDENIED))'); }
    state.storage[body.key] = body.value; save(); return response(200, {});
  }
  if (/\/_api\/search\/query$/.test(path)) {
    const m = /Title:"([^"]+)"/.exec(q.querytext || '');
    const rows = Object.keys(state.lists).filter((t) => m && t === m[1] && !state.lists[t].noCrawl)
      .map((t) => ({ Cells: [{ Key: 'Path', Value: `${WEB}/Lists/${t}/AllItems.aspx` }, { Key: 'SPWebUrl', Value: WEB }, { Key: 'LastModifiedTime', Value: state.lists[t].modified }] }));
    return response(200, { PrimaryQueryResult: { RelevantResults: { RowCount: rows.length, Table: { Rows: rows } } } });
  }
  if (/\/_api\/web\/currentuser\/groups$/.test(path)) { return response(200, { value: groupsOf(user, settingsGroups()).map((g) => ({ Id: g.Id, Title: g.Title })) }); }
  if (/\/_api\/web\/currentuser$/.test(path)) { return response(200, { Id: user.Id, Title: user.Title, Email: user.Email, LoginName: login(user), IsSiteAdmin: user.role === 'admin' }); }
  if (/\/_api\/web\/EffectiveBasePermissions$/.test(path)) { return response(200, PERM[level(user)]); }
  if (/clientPeoplePickerSearchUser$/.test(path)) {
    const s = String(body.queryParams.QueryString || '').toLowerCase();
    const hits = Object.values(USERS).filter((u) => u.Title.toLowerCase().indexOf(s) >= 0 || u.Email.indexOf(s) >= 0)
      .map((u) => ({ Key: login(u), DisplayText: u.Title, EntityData: { Email: u.Email } }));
    return response(200, { value: JSON.stringify(hits) });
  }
  if (/\/_api\/web\/ensureuser$/.test(path)) {
    const u = Object.values(USERS).find((x) => login(x) === body.logonName || x.Email === body.logonName);
    return u ? response(200, { Id: u.Id, Title: u.Title, Email: u.Email, LoginName: login(u) }) : err(404, 'The specified user could not be found.');
  }
  const rd = /\/_api\/web\/roledefinitions\/getbytype\((\d+)\)$/.exec(path);
  if (rd) { return response(200, { Id: { 2: 1073741826, 3: 1073741827, 5: 1073741829 }[rd[1]] || 1073741826 }); }
  const gbn = /\/_api\/web\/sitegroups\/getbyname\('(.*)'\)$/.exec(path);
  if (gbn) {
    const g = Object.values(state.groups).find((x) => x.Title === gbn[1].replace(/''/g, "'"));
    return g ? response(200, { Id: g.Id, Title: g.Title }) : err(404, 'Group cannot be found.');
  }
  if (/\/_api\/web\/sitegroups$/.test(path) && method === 'POST') {
    if (user.role !== 'admin') { return err(403, 'Access denied.'); }
    const g = { Id: state.nextGroup++, Title: body.Title, members: [] };
    state.groups[g.Id] = g; save();
    return response(201, { Id: g.Id, Title: g.Title });
  }

  // ---- lists
  if (/\/_api\/web\/lists$/.test(path)) {
    if (method === 'GET') {
      const value = Object.keys(state.lists).map((t) => ({
        Title: t, RootFolder: { ServerRelativeUrl: `/sites/crm/Lists/${t}` },
        Fields: [...Object.keys(BUILTIN), ...Object.keys(state.lists[t].fields)].map((n) => ({ InternalName: n, Hidden: false }))
      }));
      return response(200, { value });
    }
    if (user.role !== 'admin') { return err(403, 'Access denied. You do not have permission to perform this action or access this resource.'); }
    if (state.lists[body.Title]) { return err(500, `A list, survey, discussion board, or document library with the specified title already exists in this Web site.  Please choose another title.`); }
    state.lists[body.Title] = { fields: {}, items: [], modified: nowIso(), description: body.Description || '' };
    save();
    return response(201, { Title: body.Title });
  }
  const lm = /\/_api\/web\/lists\/getbytitle\('(.*?)'\)(.*)$/.exec(path);
  if (!lm) { return err(404, `Not implemented in the preview: ${method} ${path}`); }
  const title = lm[1].replace(/''/g, "'");
  const rest = lm[2];
  const list = state.lists[title];
  if (!list) { return err(404, `List '${title}' does not exist at site with URL '${WEB}'.`); }
  const deny = () => err(403, 'Access denied. You do not have permission to perform this action or access this resource.');

  if (rest === '' && verb === 'MERGE') { if (body.NoCrawl !== undefined) { list.noCrawl = !!body.NoCrawl; } save(); return response(204); }
  if (rest === '/EffectiveBasePermissions') { return response(200, PERM[level(user)]); }
  if (rest === '/fields/createfieldasxml') {
    if (user.role !== 'admin') { return deny(); }
    const f = fieldXml(body.parameters.SchemaXml);
    if (!f.name) { return err(400, 'Invalid field schema.'); }
    if (list.fields[f.name]) { return err(500, `A duplicate field name "${f.name}" was found.`); }
    list.fields[f.name] = { type: f.type, choices: f.choices };
    save();
    return response(201, { InternalName: f.name, TypeAsString: f.type });
  }
  const fm = /^\/fields\/getbyinternalnameortitle\('(.*)'\)$/.exec(rest);
  if (fm) {
    const f = list.fields[fm[1]];
    if (!f) { return err(404, `Column '${fm[1]}' does not exist.`); }
    if (verb === 'MERGE') { f.choices = (body.Choices && body.Choices.results) || f.choices; save(); return response(204); }
    return response(200, { TypeAsString: f.type, Choices: f.choices });
  }
  if (/^\/(breakroleinheritance|roleassignments)/.test(rest)) { list.unique = true; save(); return response(200, {}); }

  // ---- items
  const im = /^\/items(?:\((\d+)\))?(\/.*)?$/.exec(rest);
  if (!im) { return err(404, `Not implemented in the preview: ${method} ${path}`); }
  const id = im[1] ? Number(im[1]) : undefined;
  const sub = im[2] || '';
  const row = id !== undefined ? list.items.find((x) => x.Id === id) : undefined;
  if (id !== undefined && (!row || !canSee(user, row))) { return err(404, 'Item does not exist. It may have been deleted by another user.'); }
  const fieldOk = (name) => {
    if (!typeOf(list, name)) { throw Object.assign(new Error(`The field or property '${name}' does not exist.`), { status: 400 }); }
  };
  try {
    if (sub === '/recycle()') {
      if (level(user) === 'read') { return deny(); }
      list.items = list.items.filter((x) => x.Id !== id); list.modified = nowIso(); save();
      return response(200, { value: 'recycled' });
    }
    if (/^\/breakroleinheritance/.test(sub)) {
      if (user.role !== 'admin') { return deny(); }
      row.__unique = true; row.__principals = /copyRoleAssignments=true/.test(sub) ? row.__principals || [] : []; save();
      return response(200, {});
    }
    if (sub === '/resetroleinheritance') { if (user.role !== 'admin') { return deny(); } row.__unique = false; row.__principals = []; save(); return response(200, {}); }
    const ra = /^\/roleassignments\/addroleassignment\(principalid=(\d+),roledefid=(\d+)\)$/.exec(sub);
    if (ra) { if (user.role !== 'admin') { return deny(); } row.__principals = [...(row.__principals || []), Number(ra[1])]; save(); return response(200, {}); }
    if (sub === '/ValidateUpdateListItem') {
      if (level(user) === 'read') { return deny(); }
      for (const fv of body.formValues) { fieldOk(fv.FieldName); row[fv.FieldName] = fv.FieldValue === '' ? null : fv.FieldValue; }
      save();
      return response(200, { value: body.formValues.map((f) => ({ FieldName: f.FieldName, HasException: false })) });
    }
    if (sub) { return err(404, `Not implemented in the preview: ${method} ${path}`); }

    if (verb === 'GET') {
      (q.$select || '').split(',').filter(Boolean).forEach((s) => fieldOk(s.split('/')[0]));
      (q.$expand || '').split(',').filter(Boolean).forEach((s) => { fieldOk(s); if (typeOf(list, s) !== 'User') { throw Object.assign(new Error(`The field '${s}' is not a lookup and cannot be expanded.`), { status: 400 }); } });
      if (row) { return response(200, project(list, row, q.$select, q.$expand)); }
      let rows = list.items.filter((r) => canSee(user, r)).filter(parseFilter(q.$filter, fieldOk));
      if (q.$orderby) {
        const [field, dir] = q.$orderby.split(',')[0].trim().split(/\s+/);
        fieldOk(field);
        rows = rows.slice().sort((a, b) => (a[field] < b[field] ? -1 : a[field] > b[field] ? 1 : 0) * (dir === 'desc' ? -1 : 1));
      }
      if (q.$top) { rows = rows.slice(0, Number(q.$top)); }
      return response(200, { value: rows.map((r) => project(list, r, q.$select, q.$expand)) });
    }
    if (level(user) === 'read') { return deny(); }
    if (verb === 'DELETE') { list.items = list.items.filter((x) => x.Id !== id); list.modified = nowIso(); save(); return response(204); }
    const problem = validate(list, title, body || {});
    if (problem) { return err(400, problem); }
    const data = { ...body };
    delete data.__metadata;
    if (verb === 'MERGE') {
      Object.assign(row, data, { Modified: nowIso(), EditorId: user.Id });
      list.modified = nowIso(); save();
      return response(204);
    }
    const created = { Id: state.nextItem++, Created: nowIso(), Modified: nowIso(), AuthorId: user.Id, EditorId: user.Id, ...data };
    list.items.push(created); list.modified = nowIso(); save();
    return response(201, project(list, created));
  } catch (e) {
    return err(e.status || 400, e.message);
  }
}

/** An SPHttpClient look-alike over the in-memory site (with a short delay, like a network call). */
export function createFakeHttp() {
  const run = (method) => (url, _config, options) => {
    const o = options || {};
    return new Promise((resolve) => setTimeout(() => resolve(handle(method, url, o.body, o.headers || {})), 5));
  };
  return { get: run('GET'), post: run('POST'), fetch: (url, _c, o) => run((o && o.method) || 'GET')(url, _c, o) };
}

/** Puts the preview users into the CRM groups (in a real tenant an admin adds them in SharePoint). */
export function seedGroupMembers() {
  const g = settingsGroups();
  Object.values(state.groups).forEach((grp) => {
    const role = Object.keys(g).find((k) => g[k] === grp.Title);
    grp.members = Object.values(USERS).filter((u) => u.role === role).map((u) => u.Id);
  });
  save();
}

/** Stands in for a Power Automate flow writing to an item (approvals, payment webhooks, e-signature, …). */
export function patchItem(title, id, fields) {
  const row = state.lists[title] && state.lists[title].items.find((x) => x.Id === id);
  if (!row) { throw new Error(`No item ${id} in ${title}`); }
  Object.assign(row, fields, { Modified: nowIso() });
  state.lists[title].modified = nowIso();
  save();
  return row;
}

/** Adds rows to a list as if created by `authorId` (the demo data uses it). Returns the new ids. */
export function seedItems(title, rows, authorId = 7) {
  const list = state.lists[title];
  if (!list) { throw new Error(`No list ${title}`); }
  const ids = rows.map((r) => {
    const row = { Id: state.nextItem++, Created: r.Created || nowIso(), Modified: nowIso(), AuthorId: authorId, EditorId: authorId, ...r };
    list.items.push(row);
    return row.Id;
  });
  list.modified = nowIso();
  save();
  return ids;
}
