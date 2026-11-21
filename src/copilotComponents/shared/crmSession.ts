// Opens a Brillienta CRM session inside a Microsoft 365 Copilot host.
//
// The Copilot host does not run on the CRM site, so the CRM site is found the same way as in the Teams personal app:
// the main CRM site if a SharePoint admin pinned one, otherwise SharePoint Search finds the CRM lists the signed-in
// person can read (no admin rights or PowerShell needed; see CrmSiteLocator). All reads and writes then go through
// the same DataService and Engine as the web part, as the signed-in user, so SharePoint permissions, record-level
// security, "reps see only their own records", licensing, lead scoring, assignment and the audit log all behave
// exactly as they do in the CRM itself.
import { SPHttpClient } from '@microsoft/sp-http';
import { DataService } from '../../webparts/brillientaCrm/services/DataService';
import { Engine } from '../../webparts/brillientaCrm/services/Engine';
import { GraphService } from '../../webparts/brillientaCrm/services/Integrations';
import { evaluateLicense } from '../../webparts/brillientaCrm/models/license';
import { EntityKey } from '../../webparts/brillientaCrm/models/model';
import { locateCrmSite } from '../../webparts/brillientaCrm/services/CrmSiteLocator';

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
    super('No Brillienta CRM was found for your account. Ask a CRM administrator to give you access to the CRM site. If the CRM was set up in the last few minutes, try again shortly: Copilot finds it through SharePoint search.');
    this.name = 'CrmNotConfiguredError';
  }
}

const trimSlash = (u: string): string => u.replace(/\/+$/, '');

/** The CRM site for the signed-in person: pinned main CRM site, else the CRM lists SharePoint Search finds. */
export async function resolveCrmSiteUrl(ctx: ICrmHostContext): Promise<string> {
  const found = await locateCrmSite(ctx.spHttpClient, trimSlash(ctx.pageContext.web.absoluteUrl));
  if (!found) { throw new CrmNotConfiguredError(); }
  return found.url;
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
