import { SPHttpClient } from '@microsoft/sp-http';
import { MSGraphClientFactory } from '@microsoft/sp-http-msgraph';
import { ListKey } from '../models/model';

export type CrmMode = 'crm' | 'demoForm';
export type CrmHost = 'sharepoint' | 'teamsTab' | 'teamsPersonal' | 'outlook' | 'office';

export interface IBrillientaCrmProps {
  spHttpClient: SPHttpClient;
  graphFactory?: MSGraphClientFactory;
  /** Site that holds the CRM lists. */
  webUrl: string;
  /** Site the web part is running on (team site in Teams, root site for the personal app). */
  hostWebUrl: string;
  host: CrmHost;
  /** Main CRM site pinned by a SharePoint admin (tenant property), if any. Optional: the CRM is also found automatically. */
  tenantSiteUrl: string;
  /** CRM site found automatically with SharePoint Search (Teams and other hosts, when nothing is pinned). */
  foundSiteUrl?: string;
  tenantId: string;
  /** SharePoint UI culture, e.g. "de-DE" (default language). */
  culture?: string;
  /** Current user can manage the site (site owner) → CRM administrator. */
  isSiteAdmin: boolean;
  listNames: Partial<Record<ListKey, string>>;
  mode: CrmMode;
  title: string;
  fullScreen: boolean;
  /** Page is in edit mode: never cover the SharePoint editor. */
  pageEditing: boolean;
  defaultPage: string;
  formTitle: string;
  formIntro: string;
  productName: string;
  isDarkTheme: boolean;
  theme?: object;
}
