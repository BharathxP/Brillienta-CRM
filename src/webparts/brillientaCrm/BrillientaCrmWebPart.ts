import * as React from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneDropdown, PropertyPaneLabel, PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme, ThemeProvider, ThemeChangedEventArgs } from '@microsoft/sp-component-base';
import { SPPermission } from '@microsoft/sp-page-context';
import { initializeIcons, createTheme, PartialTheme } from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';

import * as strings from 'BrillientaCrmWebPartStrings';
import BrillientaCrm from './components/BrillientaCrm';
import { CrmHost, CrmMode, IBrillientaCrmProps } from './components/IBrillientaCrmProps';
import { ENTITIES, ListKey, SYSTEM_LISTS } from './models/model';

export interface IBrillientaCrmWebPartProps {
  mode: CrmMode;
  /** Fill the whole browser window (hides SharePoint header & navigation) when the page is viewed. */
  fullScreen?: boolean;
  title: string;
  defaultTab: string;
  formTitle: string;
  formIntro: string;
  productName: string;
  /** SharePoint site that holds the CRM lists. Empty = tenant default (see Settings › Microsoft Teams) or this site. */
  crmSiteUrl?: string;
  /** Optional per-site list name overrides, e.g. { "leads": "Marketplace Leads" }. */
  listNames?: Partial<Record<ListKey, string>>;
  // v1 properties (kept for upgrades)
  leadsList?: string;
  demosList?: string;
  opportunitiesList?: string;
  activitiesList?: string;
}

let iconsInitialized = false;

export const TENANT_SITE_KEY = 'BrillientaCRM.SiteUrl';

/** Brillienta-branded themes for Microsoft Teams (default / dark / high contrast). */
function teamsTheme(kind: string): PartialTheme {
  if (kind === 'dark') {
    return createTheme({
      isInverted: true,
      palette: {
        themePrimary: '#19c48a', themeLighterAlt: '#010805', themeLighter: '#04201a', themeLight: '#073b2c', themeTertiary: '#0f7654', themeSecondary: '#16ad7a',
        themeDarkAlt: '#2ec995', themeDark: '#4dd2a6', themeDarker: '#7bdfbf',
        neutralLighterAlt: '#292929', neutralLighter: '#2e2e2e', neutralLight: '#3d3d3d', neutralQuaternaryAlt: '#454545', neutralQuaternary: '#4c4c4c',
        neutralTertiaryAlt: '#6b6b6b', neutralTertiary: '#c8c8c8', neutralSecondary: '#d0d0d0', neutralPrimaryAlt: '#dadada', neutralPrimary: '#ffffff',
        neutralDark: '#f4f4f4', black: '#f8f8f8', white: '#1f1f1f'
      }
    });
  }
  if (kind === 'contrast') {
    return createTheme({
      isInverted: true,
      palette: {
        themePrimary: '#ffff01', themeLighterAlt: '#000000', themeLighter: '#1a1a00', themeLight: '#333300', themeTertiary: '#999900', themeSecondary: '#e6e600',
        themeDarkAlt: '#ffff1a', themeDark: '#ffff4d', themeDarker: '#ffff80',
        neutralLighterAlt: '#0a0a0a', neutralLighter: '#141414', neutralLight: '#ffffff', neutralQuaternaryAlt: '#ffffff', neutralQuaternary: '#ffffff',
        neutralTertiaryAlt: '#ffffff', neutralTertiary: '#ffffff', neutralSecondary: '#ffffff', neutralPrimaryAlt: '#ffffff', neutralPrimary: '#ffffff',
        neutralDark: '#ffffff', black: '#ffffff', white: '#000000'
      }
    });
  }
  return createTheme({
    palette: {
      themePrimary: '#008d62', themeLighterAlt: '#f0faf6', themeLighter: '#c6ebdd', themeLight: '#9ad9c1', themeTertiary: '#4bb48c', themeSecondary: '#109a6e',
      themeDarkAlt: '#007f58', themeDark: '#006b4b', themeDarker: '#004f37',
      neutralLighterAlt: '#f5f5f5', neutralLighter: '#f0f0f0', neutralLight: '#e1e1e1', neutralQuaternaryAlt: '#d1d1d1', neutralQuaternary: '#c8c8c8',
      neutralTertiaryAlt: '#c0c0c0', neutralTertiary: '#a19f9d', neutralSecondary: '#605e5c', neutralPrimaryAlt: '#3b3a39', neutralPrimary: '#242424',
      neutralDark: '#201f1e', black: '#000000', white: '#ffffff'
    }
  });
}

export default class BrillientaCrmWebPart extends BaseClientSideWebPart<IBrillientaCrmWebPartProps> {
  private _root: Root | undefined;
  private _theme: IReadonlyTheme | undefined;
  private _themeProvider: ThemeProvider | undefined;
  private _host: CrmHost = 'sharepoint';
  private _teamsTheme: string | undefined;
  private _tenantSiteUrl: string = '';

  protected async onInit(): Promise<void> {
    await super.onInit();
    if (!iconsInitialized) {
      initializeIcons(undefined, { disableWarnings: true });
      iconsInitialized = true;
    }
    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
    this._theme = this._themeProvider.tryGetTheme();
    this._themeProvider.themeChangedEvent.add(this, this._onThemeChangedEvent);

    // Microsoft Teams / Outlook / Microsoft 365 app host
    const teams = this.context.sdks.microsoftTeams;
    if (teams) {
      try {
        const tctx = await teams.teamsJs.app.getContext();
        const hostName = String(tctx.app.host.name);
        this._host = hostName === 'Outlook' ? 'outlook' : hostName === 'Office' ? 'office' : (tctx.page?.frameContext === 'content' && !tctx.team ? 'teamsPersonal' : 'teamsTab');
        this._teamsTheme = tctx.app.theme || 'default';
        teams.teamsJs.app.registerOnThemeChangeHandler((theme: string) => { this._teamsTheme = theme; this.render(); });
      } catch {
        this._host = 'teamsTab';
        this._teamsTheme = 'default';
      }
    }
    // Tenant-wide default CRM site (lets the Teams personal app and other sites use one CRM)
    try {
      const res = await this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/GetStorageEntity('${TENANT_SITE_KEY}')`,
        SPHttpClient.configurations.v1, { headers: { Accept: 'application/json;odata=nometadata', 'odata-version': '' } });
      if (res.ok) {
        const j = await res.json();
        this._tenantSiteUrl = String(j?.Value || '').replace(/\/$/, '');
      }
    } catch {
      /* storage entity not available */
    }
  }

  private get crmWebUrl(): string {
    const own = (this.properties.crmSiteUrl || '').trim().replace(/\/$/, '');
    return own || this._tenantSiteUrl || this.context.pageContext.web.absoluteUrl;
  }

  private get listNames(): Partial<Record<ListKey, string>> {
    const p = this.properties;
    const names: Partial<Record<ListKey, string>> = { ...(p.listNames || {}) };
    if (p.leadsList && !names.leads) { names.leads = p.leadsList; }
    if (p.demosList && !names.demos) { names.demos = p.demosList; }
    if (p.opportunitiesList && !names.opportunities) { names.opportunities = p.opportunitiesList; }
    if (p.activitiesList && !names.activities) { names.activities = p.activitiesList; }
    Object.keys(names).forEach(k => { if (!names[k as ListKey]) { delete names[k as ListKey]; } });
    return names;
  }

  public render(): void {
    const p = this.properties;
    const ctx = this.context;
    const inTeams = this._host !== 'sharepoint';
    const crmWeb = this.crmWebUrl;
    const sameSite = crmWeb.toLowerCase() === ctx.pageContext.web.absoluteUrl.toLowerCase();
    const theme = inTeams ? teamsTheme(this._teamsTheme || 'default') : (this._theme as unknown as PartialTheme);
    const element: React.ReactElement<IBrillientaCrmProps> = React.createElement(BrillientaCrm, {
      spHttpClient: ctx.spHttpClient,
      graphFactory: ctx.msGraphClientFactory,
      webUrl: crmWeb,
      hostWebUrl: ctx.pageContext.web.absoluteUrl,
      host: this._host,
      tenantSiteUrl: this._tenantSiteUrl,
      tenantId: ctx.pageContext.aadInfo?.tenantId?.toString() || '',
      isSiteAdmin: sameSite && (!!ctx.pageContext.legacyPageContext?.isSiteAdmin || ctx.pageContext.web.permissions.hasPermission(SPPermission.manageWeb)),
      listNames: this.listNames,
      mode: p.mode || 'crm',
      title: p.title || 'Brillienta CRM',
      fullScreen: p.fullScreen !== false && !inTeams,
      pageEditing: this.displayMode === DisplayMode.Edit,
      defaultPage: p.defaultTab === 'leads' || p.defaultTab === 'demos' || p.defaultTab === 'opportunities' ? `entity:${p.defaultTab}` : (p.defaultTab === 'pipeline' ? 'entity:opportunities' : (p.defaultTab || 'dashboard')),
      formTitle: p.formTitle || 'Request a demo',
      formIntro: p.formIntro || '',
      productName: p.productName || '',
      isDarkTheme: inTeams ? this._teamsTheme === 'dark' || this._teamsTheme === 'contrast' : !!this._theme?.isInverted,
      theme: theme as unknown as object
    });
    if (!this._root) {
      this._root = createRoot(this.domElement);
    }
    this._root.render(element);
  }

  private _onThemeChangedEvent(args: ThemeChangedEventArgs): void {
    this._theme = args.theme;
    this.render();
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) { return; }
    this._theme = currentTheme;
    const { semanticColors } = currentTheme;
    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDisplayModeChanged(): void {
    this.render();
  }

  protected onDispose(): void {
    this._themeProvider?.themeChangedEvent.remove(this, this._onThemeChangedEvent);
    this._root?.unmount();
    this._root = undefined;
  }

  protected get dataVersion(): Version {
    return Version.parse('2.0');
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const isForm = this.properties.mode === 'demoForm';
    const listField = (key: ListKey, label: string, placeholder: string): ReturnType<typeof PropertyPaneTextField> =>
      PropertyPaneTextField(`listNames.${key}`, { label, placeholder });
    return {
      pages: [{
        header: { description: strings.PropertyPaneDescription },
        groups: [
          {
            groupName: strings.BasicGroupName,
            groupFields: [
              ...(this.properties.mode === 'demoForm' ? [PropertyPaneDropdown('mode', { label: strings.ModeFieldLabel, options: [{ key: 'crm', text: 'Full CRM app' }, { key: 'demoForm', text: 'Request-a-demo form (legacy)' }] })] : []),
              PropertyPaneToggle('fullScreen', { label: strings.FullScreenFieldLabel, onText: 'On', offText: 'Off', checked: this.properties.fullScreen !== false }),
              ...(isForm ? [
                PropertyPaneTextField('formTitle', { label: strings.FormTitleFieldLabel }),
                PropertyPaneTextField('formIntro', { label: strings.FormIntroFieldLabel, multiline: true }),
                PropertyPaneTextField('productName', { label: strings.ProductNameFieldLabel })
              ] : [
                PropertyPaneTextField('title', { label: strings.TitleFieldLabel }),
                PropertyPaneDropdown('defaultTab', {
                  label: strings.DefaultTabFieldLabel,
                  options: [
                    { key: 'dashboard', text: 'Dashboard' }, { key: 'mywork', text: 'My work' }, { key: 'leads', text: 'Leads' },
                    { key: 'opportunities', text: 'Opportunities (pipeline)' }, { key: 'demos', text: 'Demo requests' },
                    { key: 'calendar', text: 'Calendar' }, { key: 'reports', text: 'Reports' }
                  ]
                }),
                PropertyPaneLabel('settingsInfo', { text: strings.SettingsInfo })
              ])
            ]
          },
          {
            groupName: strings.ListsGroupName,
            isCollapsed: !this.properties.crmSiteUrl,
            groupFields: [
              PropertyPaneTextField('crmSiteUrl', {
                label: strings.CrmSiteFieldLabel,
                placeholder: this._tenantSiteUrl || this.context.pageContext.web.absoluteUrl,
                description: strings.CrmSiteFieldDescription
              }),
              PropertyPaneLabel('listsInfo', { text: strings.ListsInfo }),
              ...ENTITIES.map(e => listField(e.key, e.plural, e.listTitle)),
              listField('settings', 'Settings', SYSTEM_LISTS.settings.listTitle),
              listField('audit', 'Audit log', SYSTEM_LISTS.audit.listTitle)
            ]
          }
        ]
      }]
    };
  }
}
