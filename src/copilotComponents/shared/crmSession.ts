// Opens a Brillienta CRM session inside a Microsoft 365 Copilot host.
//
// The Copilot host does not run on the CRM site, so the CRM site is found the same way the Teams
// personal app finds it: from the tenant storage entity set in Settings › Microsoft Teams › Main CRM site
// ("Use this site"). All reads and writes then go through the same DataService and
// Engine as the web part, as the signed-in user, so SharePoint permissions, record-level security,
// "reps see only their own records", licensing, lead scoring, assignment and the audit log all
// behave exactly as they do in the CRM itself.
import { SPHttpClient } from '@microsoft/sp-http';
import { DataService } from '../../webparts/brillientaCrm/services/DataService';
import { Engine } from '../../webparts/brillientaCrm/services/Engine';
import { GraphService } from '../../webparts/brillientaCrm/services/Integrations';
import { evaluateLicense } from '../../webparts/brillientaCrm/models/license';
import { EntityKey } from '../../webparts/brillientaCrm/models/model';

/** Same key as the web part and Settings › Microsoft Teams. */
export const TENANT_SITE_KEY = 'BrillientaCRM.SiteUrl';

/** The parts of the SPFx context the session needs (kept small so tests can fake it). */
export interface ICrmHostContext {
  spHttpClient: SPHttpClient;
  pageContext: {
    web: { absoluteUrl: string };
    aadInfo?: { tenantId?: { toString(): string } };
  };
}

export interface ICrmSession {
  engine: Engine;
  /** The SharePoint site that holds the CRM lists. */
  siteUrl: string;
  /** Absolute list URLs by entity, used for "Open in CRM" links. */
  listUrls: Partial<Record<EntityKey, string>>;
}

export class CrmNotConfiguredError extends Error {
  public constructor() {
    super('No main CRM site is set for this tenant yet. A CRM administrator can open Brillienta CRM on the CRM site and choose Settings › Microsoft Teams › Main CRM site › "Use this site".');
    this.name = 'CrmNotConfiguredError';
  }
}

const trimSlash = (u: string): string => u.replace(/\/+$/, '');

/** Reads the tenant-wide CRM site URL. Storage entities are tenant scoped, so any site can read them. */
export async function resolveCrmSiteUrl(ctx: ICrmHostContext): Promise<string> {
  const host = trimSlash(ctx.pageContext.web.absoluteUrl);
  try {
    const res = await ctx.spHttpClient.get(`${host}/_api/web/GetStorageEntity('${TENANT_SITE_KEY}')`, SPHttpClient.configurations.v1,
      { headers: { Accept: 'application/json;odata=nometadata', 'odata-version': '' } });
    if (res.ok) {
      const j = await res.json();
      const v = trimSlash(String(j?.Value || '').trim());
      if (v) { return v; }
    }
  } catch {
    /* fall through */
  }
  throw new CrmNotConfiguredError();
}

async function openSession(ctx: ICrmHostContext): Promise<ICrmSession> {
  const siteUrl = await resolveCrmSiteUrl(ctx);
  const tenantId = String(ctx.pageContext.aadInfo?.tenantId || '');
  const ds = new DataService(ctx.spHttpClient, siteUrl);
  // Graph is not needed by any Copilot tool (no email, calendar or Teams calls), so none is passed.
  const engine = new Engine(ds, new GraphService(undefined), tenantId, () => undefined);
  await ds.loadSettings().catch(e => { if ((e as { status?: number }).status !== 404) { throw e; } });
  const [lists, me] = await Promise.all([ds.checkLists(), ds.currentUser()]);
  const license = await evaluateLicense(engine.settings.licenseKey, tenantId, engine.settings.installDate);
  const role = await ds.resolveRole(await ds.isCrmSiteAdmin());
  engine.setSession(me, role, license);
  const origin = siteUrl.replace(/^(https:\/\/[^/]+).*$/i, '$1');
  const listUrls: Partial<Record<EntityKey, string>> = {};
  lists.forEach(l => { if (l.url && l.key !== 'settings' && l.key !== 'audit') { listUrls[l.key as EntityKey] = /^https?:/i.test(l.url) ? l.url : `${origin}${l.url}`; } });
  return { engine, siteUrl, listUrls };
}

let shared: Promise<ICrmSession> | undefined;

/** One session per page load; every Copilot tool in the bundle reuses it. A failed attempt is retried next time. */
export function getCrmSession(ctx: ICrmHostContext): Promise<ICrmSession> {
  if (!shared) {
    shared = openSession(ctx);
    shared.catch(() => { shared = undefined; });
  }
  return shared;
}

/** Link to a record's SharePoint display form (always available, respects permissions). */
export function recordUrl(session: ICrmSession, key: EntityKey, id?: number): string {
  const list = session.listUrls[key];
  if (!list || !id) { return session.siteUrl; }
  return `${list}/DispForm.aspx?ID=${id}`;
}
