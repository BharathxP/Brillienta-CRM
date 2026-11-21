/* eslint-disable @typescript-eslint/no-explicit-any */
import { crmSearchQuery, locateCrmSite, parseCrmSearch } from './CrmSiteLocator';

const row = (path: string, web: string, modified: string): any => ({
  Cells: [{ Key: 'Path', Value: path }, { Key: 'SPWebUrl', Value: web }, { Key: 'LastModifiedTime', Value: modified }]
});
const searchResult = (...rows: any[]): any => ({ PrimaryQueryResult: { RelevantResults: { Table: { Rows: rows } } } });

describe('parseCrmSearch', () => {
  it('returns one entry per CRM site, most recently changed first', () => {
    const hits = parseCrmSearch(searchResult(
      row('https://t.sharepoint.com/sites/old-crm/Lists/CRM Opportunities', 'https://t.sharepoint.com/sites/old-crm', '2025-01-01T00:00:00Z'),
      row('https://t.sharepoint.com/sites/crm/Lists/CRMOpportunities/', 'https://t.sharepoint.com/sites/crm/', '2026-11-01T00:00:00Z'),
      row('https://t.sharepoint.com/sites/crm/Lists/CRM%20Opportunities', 'https://t.sharepoint.com/sites/crm', '2026-10-01T00:00:00Z')
    ));
    expect(hits.map(h => h.webUrl)).toEqual(['https://t.sharepoint.com/sites/crm', 'https://t.sharepoint.com/sites/old-crm']);
  });

  it('ignores results that are not the CRM list itself', () => {
    const hits = parseCrmSearch(searchResult(
      row('https://t.sharepoint.com/sites/x/Shared Documents/CRM Opportunities.docx', 'https://t.sharepoint.com/sites/x', '2026-01-01'),
      row('https://t.sharepoint.com/sites/y/Lists/Other', 'https://t.sharepoint.com/sites/y', '2026-01-01'),
      row('https://t.sharepoint.com/sites/z/Lists/CRM Opportunities', '', '2026-01-01')
    ));
    expect(hits).toEqual([]);
    expect(parseCrmSearch(undefined)).toEqual([]);
    expect(parseCrmSearch({})).toEqual([]);
  });

  it('searches for the list by title and type', () => {
    expect(crmSearchQuery()).toBe('Title:"CRM Opportunities" contentclass:STS_List_GenericList');
  });
});

describe('locateCrmSite', () => {
  const HOST = 'https://t.sharepoint.com';
  /** A fake SPHttpClient that answers by URL. */
  const client = (answers: { pinned?: string; hereLists?: string[]; search?: any }): any => ({
    get: async (url: string) => {
      let body: any;
      if (url.indexOf('GetStorageEntity') >= 0) { body = { Value: answers.pinned || null }; }
      else if (url.indexOf('/_api/web/lists') >= 0) { body = { value: (answers.hereLists || []).map(Title => ({ Title })) }; }
      else if (url.indexOf('/_api/search/query') >= 0) { body = answers.search || searchResult(); }
      return { ok: true, json: async () => body };
    }
  });

  it('uses a pinned site first', async () => {
    const found = await locateCrmSite(client({ pinned: 'https://t.sharepoint.com/sites/pinned/', hereLists: ['CRM Settings', 'CRM Opportunities'] }), HOST);
    expect(found).toEqual({ url: 'https://t.sharepoint.com/sites/pinned', source: 'pinned' });
  });

  it('uses the current site when it has the CRM lists', async () => {
    const found = await locateCrmSite(client({ hereLists: ['CRM Settings', 'CRM Opportunities'] }), `${HOST}/sites/team/`);
    expect(found).toEqual({ url: `${HOST}/sites/team`, source: 'here' });
  });

  it('otherwise finds the CRM with SharePoint Search, without admin rights', async () => {
    const found = await locateCrmSite(client({
      hereLists: ['CRM Settings'],
      search: searchResult(
        row(`${HOST}/sites/crm/Lists/CRM Opportunities`, `${HOST}/sites/crm`, '2026-11-01'),
        row(`${HOST}/sites/crm2/Lists/CRM Opportunities`, `${HOST}/sites/crm2`, '2026-01-01')
      )
    }), HOST);
    expect(found).toEqual({ url: `${HOST}/sites/crm`, source: 'search', others: [`${HOST}/sites/crm2`] });
  });

  it('returns undefined when nothing is found, and survives failed requests', async () => {
    expect(await locateCrmSite(client({}), HOST)).toBeUndefined();
    const failing: any = { get: async () => { throw new Error('offline'); } };
    expect(await locateCrmSite(failing, HOST)).toBeUndefined();
  });
});
