// Finds the SharePoint site that holds the CRM lists when the CRM runs somewhere else: the Teams personal app (root
// site), a Teams channel tab on another team's site, another SharePoint page, or Microsoft 365 Copilot.
//
// No admin rights and no PowerShell are needed:
//   1. "pinned": the tenant property BrillientaCRM.SiteUrl, if a SharePoint admin set one (optional).
//   2. "here":   the current site, when it has the CRM lists.
//   3. "search": SharePoint Search for the CRM's lists, with the signed-in person's permissions. Every CRM user can
//                read the CRM lists, so search finds the CRM for them. New sites appear once search has indexed them
//                (usually within minutes).
import { SPHttpClient } from '@microsoft/sp-http';

/** Same key as Settings › Microsoft Teams and the PnP PowerShell command. */
export const TENANT_SITE_KEY = 'BrillientaCRM.SiteUrl';

/** The CRM list that search looks for. Crawled (unlike BrillCRM_Settings), present on every CRM site and rarely renamed. */
export const MARKER_LIST = 'BrillCRM_Opportunities';
/** Lists that, together, show a site holds a CRM (checked on the current site). */
const HERE_LISTS = ['BrillCRM_Settings', 'BrillCRM_Opportunities'];

export type CrmSiteSource = 'pinned' | 'here' | 'search';

export interface ICrmSiteLocation {
  url: string;
  source: CrmSiteSource;
  /** Other CRM sites search found (when there is more than one CRM in the tenant). */
  others?: string[];
}

export interface ISearchHit {
  webUrl: string;
  path: string;
  modified: string;
}

const HEADERS = { Accept: 'application/json;odata=nometadata', 'odata-version': '' };
const trim = (u: string): string => u.trim().replace(/\/+$/, '');
const same = (a: string, b: string): boolean => trim(a).toLowerCase() === trim(b).toLowerCase();

type SearchResponse = {
  PrimaryQueryResult?: { RelevantResults?: { Table?: { Rows?: Array<{ Cells: Array<{ Key: string; Value?: string }> }> } } };
};

/** The search query for CRM marker lists. */
export function crmSearchQuery(listTitle: string = MARKER_LIST): string {
  return `Title:"${listTitle}" contentclass:STS_List_GenericList`;
}

/**
 * CRM sites from a search response: one per site, the most recently changed first. Only results whose address is the
 * list itself (/Lists/<name>, usually reported as its default view /Lists/<name>/AllItems.aspx) are accepted, so a
 * document or item that merely mentions the name is ignored.
 */
export function parseCrmSearch(res: SearchResponse | undefined, listTitle: string = MARKER_LIST): ISearchHit[] {
  const rows = res?.PrimaryQueryResult?.RelevantResults?.Table?.Rows || [];
  const listName = listTitle.replace(/\s+/g, '').toLowerCase();
  const hits: ISearchHit[] = [];
  rows.forEach(r => {
    const cell = (k: string): string => String((r.Cells.find(c => c.Key === k) || { Value: '' }).Value || '');
    const path = trim(cell('Path'));
    const webUrl = trim(cell('SPWebUrl'));
    let decoded = path;
    try { decoded = decodeURIComponent(path); } catch { /* keep as is */ }
    const marker = `/lists/${listName}`;
    const at = decoded.replace(/\s+/g, '').toLowerCase().lastIndexOf(marker);
    const rest = at >= 0 ? decoded.replace(/\s+/g, '').toLowerCase().slice(at + marker.length) : undefined;
    const isList = rest !== undefined && (rest === '' || /^\/(forms\/)?[^/]*\.aspx$/.test(rest) || rest === '/');
    if (!webUrl || !/^https:\/\//i.test(webUrl) || !isList) { return; }
    if (hits.some(h => same(h.webUrl, webUrl))) { return; }
    hits.push({ webUrl, path, modified: cell('LastModifiedTime') });
  });
  return hits.sort((a, b) => b.modified.localeCompare(a.modified));
}

async function getJson<T>(client: SPHttpClient, url: string): Promise<T | undefined> {
  try {
    const res = await client.get(url, SPHttpClient.configurations.v1, { headers: HEADERS });
    return res.ok ? (await res.json()) as T : undefined;
  } catch {
    return undefined;
  }
}

/** The pinned main CRM site (tenant property), or '' when none is set. Readable by everyone. */
export async function readPinnedCrmSite(client: SPHttpClient, anyWebUrl: string): Promise<string> {
  const j = await getJson<{ Value?: string | null }>(client, `${trim(anyWebUrl)}/_api/web/GetStorageEntity('${TENANT_SITE_KEY}')`);
  return trim(String(j?.Value || ''));
}

/** True when the site has the CRM lists (and the person can read them). */
export async function hasCrmLists(client: SPHttpClient, webUrl: string): Promise<boolean> {
  const filter = HERE_LISTS.map(t => `Title eq '${t}'`).join(' or ');
  const j = await getJson<{ value?: { Title: string }[] }>(client, `${trim(webUrl)}/_api/web/lists?$select=Title&$filter=${encodeURIComponent(filter)}`);
  return (j?.value || []).length === HERE_LISTS.length;
}

/** CRM sites the signed-in person can see, found with SharePoint Search (most recently changed first). */
export async function searchCrmSites(client: SPHttpClient, anyWebUrl: string, listTitle: string = MARKER_LIST): Promise<ISearchHit[]> {
  const q = encodeURIComponent(crmSearchQuery(listTitle).replace(/'/g, "''"));
  const res = await getJson<SearchResponse>(client,
    `${trim(anyWebUrl)}/_api/search/query?querytext='${q}'&selectproperties='Path,SPWebUrl,LastModifiedTime'&rowlimit=20&trimduplicates=false`);
  return parseCrmSearch(res, listTitle);
}

/** Where the CRM lives, seen from `hostWebUrl`. Undefined when no CRM can be found for this person. */
export async function locateCrmSite(client: SPHttpClient, hostWebUrl: string): Promise<ICrmSiteLocation | undefined> {
  const [pinned, here] = await Promise.all([readPinnedCrmSite(client, hostWebUrl), hasCrmLists(client, hostWebUrl)]);
  if (pinned) { return { url: pinned, source: 'pinned' }; }
  if (here) { return { url: trim(hostWebUrl), source: 'here' }; }
  const hits = await searchCrmSites(client, hostWebUrl);
  if (!hits.length) { return undefined; }
  return { url: hits[0].webUrl, source: 'search', others: hits.slice(1).map(h => h.webUrl) };
}
