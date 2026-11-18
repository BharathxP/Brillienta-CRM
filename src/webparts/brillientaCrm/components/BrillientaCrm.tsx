import * as React from 'react';
import {
  MessageBar, MessageBarType, Spinner, Stack, Text, IconButton, ThemeProvider, PartialTheme, SearchBox, Callout, DirectionalHint,
  Panel, PanelType, Icon, CommandButton, IContextualMenuItem, Link, FocusZone, TooltipHost, Customizer, LayerHost
} from '@fluentui/react';
import type { IBrillientaCrmProps } from './IBrillientaCrmProps';
import { CrmData, EntityKey, IRecord } from '../models/model';
import { ICrmSettings } from '../models/settings';
import { evaluateLicense } from '../models/license';
import { DataService, IListStatus } from '../services/DataService';
import { GraphService } from '../services/Integrations';
import { Engine } from '../services/Engine';
import { CrmContext, ICrmContext, IOpenRecord } from './CrmContext';
import { RecordPanel } from './core/RecordPanel';
import { EntityGrid } from './core/EntityGrid';
import { Dashboard } from './pages/Dashboard';
import { MyWork } from './pages/MyWork';
import { CalendarPage, ForecastPage, ReportsPage } from './pages/Insights';
import { SettingsPage } from './pages/SettingsPage';
import { AuditLogPage, SetupPage } from './pages/AdminPages';
import { DemoRequestForm } from './DemoRequestForm';
import { errorText, isOverdue, lsGet, lsSet } from './utils';
import { LOGO_DARK, LOGO_LIGHT, MARK, MARK_DARK } from './logo';
import styles from './BrillientaCrm.module.scss';

interface IToast { id: number; text: string; type: MessageBarType }

/** Maps the SharePoint/Teams theme (incl. section background variants) onto the app's CSS variables. */
function themeVars(theme?: PartialTheme): React.CSSProperties {
  const sc = (theme?.semanticColors || {}) as Record<string, string | undefined>;
  const pl = (theme?.palette || {}) as Record<string, string | undefined>;
  const vars: Record<string, string | undefined> = {
    '--crm-text': sc.bodyText, '--crm-subtle': sc.bodySubtext || pl.neutralSecondary, '--crm-surface': sc.bodyBackground,
    '--crm-surface-2': sc.bodyStandoutBackground || pl.neutralLighterAlt, '--crm-surface-3': sc.bodyBackgroundHovered || pl.neutralLighter,
    '--crm-border': sc.bodyDivider || pl.neutralLight, '--crm-accent': pl.themePrimary, '--crm-accent-soft': pl.themeLighter
  };
  const out: Record<string, string> = {};
  Object.keys(vars).forEach(k => { if (vars[k]) { out[k] = vars[k]!; } });
  return out as React.CSSProperties;
}

const Shell: React.FC<{ theme?: PartialTheme; className: string; hostId: string; children?: React.ReactNode }> = ({ theme, className, hostId, children }) => (
  <ThemeProvider theme={theme} className={className} style={themeVars(theme)}>
    <Customizer scopedSettings={{ Layer: { hostId } }}>
      <>{children}</>
    </Customizer>
    <LayerHost id={hostId} className="crm-layer-host" />
  </ThemeProvider>
);
interface INav { key: string; label: string; icon: string; group: string; show: boolean; count?: number }

/* ---------------------------------------------------------------- global search */

const SEARCH_ENTITIES: EntityKey[] = ['leads', 'accounts', 'contacts', 'opportunities', 'quotes', 'cases', 'demos', 'campaigns', 'products'];

const GlobalSearch: React.FC = () => {
  const { engine, data, open, isNarrow } = React.useContext(CrmContext)!;
  const [q, setQ] = React.useState('');
  const [show, setShow] = React.useState(false);
  const ref = React.useRef<HTMLDivElement>(null);
  const s = q.trim().toLowerCase();
  const results: { key: EntityKey; r: IRecord }[] = [];
  if (s.length >= 2) {
    SEARCH_ENTITIES.filter(k => engine.hasFeature(k)).forEach(k => {
      const d = engine.def(k);
      engine.visible(k).forEach(r => {
        if (results.filter(x => x.key === k).length >= 5) { return; }
        if (d.searchFields.some(f => String(r[f] || '').toLowerCase().indexOf(s) >= 0)) { results.push({ key: k, r }); }
      });
    });
  }
  return (
    <div ref={ref} className="crm-search">
      <SearchBox placeholder={isNarrow ? 'Search' : 'Search leads, accounts, contacts, deals, quotes, cases…'} value={q}
        onFocus={() => { setShow(true); engine.ensure(SEARCH_ENTITIES).catch(() => undefined); }}
        onChange={(_, v) => { setQ(v || ''); setShow(true); }} onEscape={() => setShow(false)} />
      {show && s.length >= 2 && (
        <Callout target={ref} isBeakVisible={false} directionalHint={DirectionalHint.bottomLeftEdge} onDismiss={() => setShow(false)} setInitialFocus={false}
          calloutWidth={ref.current ? Math.max(320, ref.current.offsetWidth) : 380} calloutMaxHeight={420}>
          <div className="crm-search-results" role="listbox">
            {results.length === 0 && <Text className="crm-muted" style={{ padding: 12, display: 'block' }}>{SEARCH_ENTITIES.some(k => engine.hasFeature(k) && !data[k]) ? 'Loading…' : 'No matches'}</Text>}
            {results.map(({ key, r }) => {
              const d = engine.def(key);
              return (
                <button key={`${key}-${r.id}`} className="crm-search-hit" role="option" aria-selected={false} onClick={() => { setShow(false); setQ(''); open(key, r); }}>
                  <Icon iconName={d.icon} />
                  <span><b>{r.title || r.quoteNumber}</b><small>{d.singular}{r.email ? ` · ${r.email}` : ''}{r.company ? ` · ${r.company}` : ''}{r.stage ? ` · ${r.stage}` : ''}</small></span>
                </button>
              );
            })}
          </div>
        </Callout>
      )}
    </div>
  );
};

/* ---------------------------------------------------------------- notifications */

const Notifications: React.FC<{ count: number; groups: { title: string; key: EntityKey; rows: IRecord[] }[]; onSeen: () => void }> = ({ count, groups, onSeen }) => {
  const { open } = React.useContext(CrmContext)!;
  const [show, setShow] = React.useState(false);
  const ref = React.useRef<HTMLDivElement>(null);
  return (
    <div ref={ref} style={{ position: 'relative' }}>
      <IconButton iconProps={{ iconName: count ? 'RingerSolid' : 'Ringer' }} title="Notifications" ariaLabel={`Notifications (${count})`} onClick={() => setShow(!show)} />
      {count > 0 && <span className="crm-badge">{count > 99 ? '99+' : count}</span>}
      {show && (
        <Callout target={ref} onDismiss={() => { setShow(false); onSeen(); }} directionalHint={DirectionalHint.bottomRightEdge} calloutWidth={340} calloutMaxHeight={440}>
          <div style={{ padding: 12 }}>
            <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>Notifications</Text>
            {count === 0 && <Text block className="crm-muted" style={{ marginTop: 8 }}>You are all caught up.</Text>}
            {groups.filter(g => g.rows.length).map(g => (
              <div key={g.title} style={{ marginTop: 10 }}>
                <Text variant="small" styles={{ root: { fontWeight: 600 } }} className="crm-muted">{g.title} ({g.rows.length})</Text>
                {g.rows.slice(0, 6).map(r => (
                  <button key={r.id} className="crm-search-hit" onClick={() => { setShow(false); open(g.key, r); }}>
                    <span><b>{r.title || r.caseNumber}</b><small>{r.company || r.relatedType || r.priority || ''}</small></span>
                  </button>
                ))}
              </div>
            ))}
          </div>
        </Callout>
      )}
    </div>
  );
};

const BrillientaCrm: React.FC<IBrillientaCrmProps> = props => {
  const [data, setData] = React.useState<CrmData>({});
  const engine = React.useMemo(() => {
    const ds = new DataService(props.spHttpClient, props.webUrl, props.listNames);
    return new Engine(ds, new GraphService(props.graphFactory), props.tenantId, (key, rows) => setData(d => ({ ...d, [key]: rows })));
  }, [props.webUrl, JSON.stringify(props.listNames)]);

  const [phase, setPhase] = React.useState<'loading' | 'setup' | 'ready' | 'error'>('loading');
  const [statuses, setStatuses] = React.useState<IListStatus[]>();
  const [settings, setSettings] = React.useState<ICrmSettings>(engine.settings);
  const [fatal, setFatal] = React.useState<string>();
  const [page, setPage] = React.useState<string>(props.defaultPage || 'dashboard');
  const [stack, setStack] = React.useState<IOpenRecord[]>([]);
  const [toasts, setToasts] = React.useState<IToast[]>([]);
  const [navOpen, setNavOpen] = React.useState(false);
  const [collapsed, setCollapsed] = React.useState<boolean>(() => lsGet('brillienta-crm:nav-collapsed', false));
  const [width, setWidth] = React.useState(1200);
  const rootRef = React.useRef<HTMLDivElement>(null);
  const isNarrow = width < 760;
  const [fsPref, setFsPref] = React.useState<boolean | undefined>(() => lsGet<boolean | undefined>('brillienta-crm:fullscreen', undefined));
  const toggleFullScreen = (): void => { const next = !(fsPref === undefined ? props.fullScreen : fsPref); setFsPref(next); lsSet('brillienta-crm:fullscreen', next); };
  const layerHostId = React.useMemo(() => `crm-layers-${Math.random().toString(36).slice(2, 9)}`, []);

  React.useEffect(() => {
    const el = rootRef.current;
    if (!el || typeof ResizeObserver === 'undefined') { return; }
    const ro = new ResizeObserver(entries => setWidth(entries[0].contentRect.width));
    ro.observe(el);
    return () => ro.disconnect();
  }, [phase]);

  const notify = React.useCallback((text: string, type: MessageBarType = MessageBarType.info): void => {
    const id = Date.now() + Math.random();
    setToasts(t => [...t.slice(-3), { id, text, type }]);
    setTimeout(() => setToasts(t => t.filter(x => x.id !== id)), type === MessageBarType.error ? 9000 : 4000);
  }, []);

  const init = React.useCallback(async (): Promise<void> => {
    setPhase('loading');
    try {
      await engine.ds.loadSettings().catch(e => { if ((e as { status?: number }).status !== 404) { throw e; } });
      const [st, me] = await Promise.all([engine.ds.checkLists(), engine.ds.currentUser()]);
      setStatuses(st);
      const license = await evaluateLicense(engine.settings.licenseKey, props.tenantId, engine.settings.installDate);
      const role = await engine.ds.resolveRole(props.isSiteAdmin || await engine.ds.isCrmSiteAdmin());
      engine.setSession(me, role, license);
      setSettings(engine.settings);
      if (!st.every(s => s.exists && !s.missingFields.length)) { setPhase('setup'); return; }
      if (props.mode !== 'demoForm') {
        // warm the most-used lists in the background
        engine.ensure(['leads', 'opportunities', 'activities', 'demos']).catch(e => notify(errorText(e), MessageBarType.error));
      }
      setPhase('ready');
    } catch (e) {
      setFatal(errorText(e));
      setPhase('error');
    }
  }, [engine]);

  React.useEffect(() => { init().catch(() => undefined); }, [init]);

  const updateSettings = async (s: ICrmSettings): Promise<void> => {
    await engine.ds.saveSettings(s);
    const license = await evaluateLicense(s.licenseKey, props.tenantId, s.installDate);
    engine.setSession(engine.me, engine.role, license);
    setSettings(s);
    setData(d => ({ ...d }));
  };

  const ctx: ICrmContext = {
    engine, data, settings, currency: settings.currency || 'USD', isNarrow, notify, updateSettings,
    open: (key: EntityKey, record: IRecord = {}) => setStack(s => [...s.slice(-4), { key, record }]),
    navigate: (p: string) => { setPage(p); setNavOpen(false); }
  };

  const theme = props.theme as PartialTheme | undefined;
  const inTeams = props.host !== 'sharepoint';
  const canFullScreen = !inTeams && !props.pageEditing && props.mode !== 'demoForm';
  const fullScreen = canFullScreen && (fsPref === undefined ? props.fullScreen : fsPref);
  const rootClass = `${styles.brillientaCrm} ${props.isDarkTheme ? 'crm-dark' : ''} ${inTeams ? 'crm-teams' : ''} ${fullScreen ? 'crm-immersive' : ''}`;

  // Full screen: stop the SharePoint page behind the CRM from scrolling.
  React.useEffect(() => {
    if (!fullScreen) { return undefined; }
    const html = document.documentElement;
    const prev = html.style.overflow;
    html.style.overflow = 'hidden';
    return () => { html.style.overflow = prev; };
  }, [fullScreen]);

  // ------------------------------------------------------------------ non-ready states
  if (phase === 'loading') {
    return <Shell theme={theme} className={rootClass} hostId={layerHostId}><Spinner label="Loading Brillienta CRM..." style={{ padding: 48 }} /></Shell>;
  }
  if (phase === 'error') {
    return (
      <Shell theme={theme} className={rootClass} hostId={layerHostId}>
        <MessageBar messageBarType={MessageBarType.error} isMultiline actions={<Link onClick={() => { init().catch(() => undefined); }}>Retry</Link>}>{fatal}</MessageBar>
      </Shell>
    );
  }
  if (phase === 'setup') {
    return (
      <Shell theme={theme} className={rootClass} hostId={layerHostId}>
        {inTeams && (
          <MessageBar messageBarType={MessageBarType.info} isMultiline styles={{ root: { maxWidth: 900, margin: '16px auto 0' } }}>
            <b>Running in Microsoft Teams.</b> {props.host === 'teamsTab'
              ? <>This tab is connected to the team’s site ({props.webUrl}). To use your company CRM instead, open the tab’s <b>Settings</b> and enter the CRM site URL — or ask an admin to set the default CRM site in <i>CRM › Settings › Microsoft Teams</i>. You can also set up a separate CRM for this team below.</>
              : <>The personal app needs the default CRM site. An admin sets it once in <i>CRM › Settings › Microsoft Teams</i> on the CRM site (currently checking {props.webUrl}).</>}
          </MessageBar>
        )}
        <SetupPage ds={engine.ds} statuses={statuses} canProvision={engine.role === 'admin' && props.host !== 'teamsPersonal'} onDone={() => { init().catch(() => undefined); }} logo={props.isDarkTheme ? LOGO_DARK : LOGO_LIGHT} />
      </Shell>
    );
  }
  if (props.mode === 'demoForm') {
    return (
      <Shell theme={theme} className={rootClass} hostId={layerHostId}>
        <CrmContext.Provider value={ctx}>
          <DemoRequestForm engine={engine} title={props.formTitle} intro={props.formIntro} productName={props.productName} />
        </CrmContext.Provider>
      </Shell>
    );
  }

  // ------------------------------------------------------------------ navigation
  const me = engine.me?.id;
  const myOverdue = (data.activities || []).filter(a => a.activityType === 'Task' && !a.completed && a.owner?.id === me && isOverdue(a.dueDate));
  const myBreached = (data.cases || []).filter(c => c.owner?.id === me && c.caseStatus !== 'Resolved' && c.caseStatus !== 'Closed' && isOverdue(c.slaDue));
  const lastSeen = lsGet<string>('brillienta-crm:last-seen', '');
  const newForMe = (data.leads || []).filter(l => l.owner?.id === me && l.status === 'New' && (!lastSeen || String(l.created) > lastSeen));
  const todayEnd = new Date(); todayEnd.setHours(23, 59, 59, 999);
  const demosToday = (data.demos || []).filter(d => d.owner?.id === me && d.status === 'Scheduled' && d.scheduledDate && new Date(d.scheduledDate).getTime() <= todayEnd.getTime() && new Date(d.scheduledDate).getTime() >= Date.now() - 3600000);
  const notifCount = myOverdue.length + myBreached.length + newForMe.length + demosToday.length;

  const nav: INav[] = [
    { key: 'dashboard', label: 'Dashboard', icon: 'ViewDashboard', group: 'Home', show: true },
    { key: 'mywork', label: 'My work', icon: 'TaskManager', group: 'Home', show: true, count: myOverdue.length || undefined },
    { key: 'calendar', label: 'Calendar', icon: 'Calendar', group: 'Home', show: true },
    ...engine.defs.filter(d => d.key !== 'activities').map(d => ({
      key: `entity:${d.key}`, label: d.plural, icon: d.icon, group: d.area, show: true,
      count: d.key === 'leads' ? (data.leads || []).filter(l => l.status === 'New').length || undefined
        : d.key === 'demos' ? (data.demos || []).filter(x => x.status === 'Requested').length || undefined
        : d.key === 'cases' ? (data.cases || []).filter(x => x.caseStatus === 'New').length || undefined : undefined
    })),
    { key: 'entity:activities', label: 'Activities', icon: 'BulletedList', group: 'Work', show: true },
    { key: 'forecast', label: 'Forecast', icon: 'Chart', group: 'Insights', show: engine.can('forecast') },
    { key: 'reports', label: 'Reports', icon: 'ReportDocument', group: 'Insights', show: true },
    { key: 'audit', label: 'Audit log', icon: 'ComplianceAudit', group: 'Admin', show: engine.can('admin') && engine.hasFeature('audit') },
    { key: 'settings', label: 'Settings', icon: 'Settings', group: 'Admin', show: engine.can('admin') },
    { key: 'setup', label: 'Setup & health', icon: 'Repair', group: 'Admin', show: engine.can('admin') }
  ].filter(n => n.show);
  const groups = ['Home', 'Sales', 'Marketing', 'Service', 'Catalog', 'Work', 'Insights', 'Admin'];
  const current = nav.find(n => n.key === page) || nav[0];

  const navList = (compact: boolean): React.ReactNode => (
    <FocusZone as="nav" className="crm-nav-inner" aria-label="CRM navigation">
      {groups.map(g => {
        const items = nav.filter(n => n.group === g);
        if (!items.length) { return null; }
        return (
          <div key={g} className="crm-nav-group">
            {!compact && <div className="crm-nav-group-title">{g}</div>}
            {items.map(n => (
              <TooltipHost key={n.key} content={compact ? n.label : undefined} directionalHint={DirectionalHint.rightCenter}>
                <button className={`crm-nav-item${n.key === current.key ? ' crm-nav-active' : ''}`} aria-current={n.key === current.key ? 'page' : undefined}
                  onClick={() => ctx.navigate(n.key)} aria-label={n.label}>
                  <Icon iconName={n.icon} />
                  {!compact && <span className="crm-nav-label">{n.label}</span>}
                  {!compact && n.count ? <span className="crm-nav-count">{n.count}</span> : null}
                </button>
              </TooltipHost>
            ))}
          </div>
        );
      })}
    </FocusZone>
  );

  const newMenu: IContextualMenuItem[] = engine.defs.filter(d => engine.can('create', d.key)).map(d => ({
    key: d.key, text: d.singular, iconProps: { iconName: d.icon }, onClick: () => ctx.open(d.key, d.key === 'activities' ? { activityType: 'Task', completed: false } : {})
  }));

  const pageBody = (): React.ReactNode => {
    if (current.key.indexOf('entity:') === 0) {
      const key = current.key.slice(7) as EntityKey;
      return <EntityGrid key={key} entity={key} defaultMode={key === 'opportunities' && !isNarrow ? 'board' : 'list'}
        initialFilters={key === 'opportunities' || key === 'cases' || key === 'leads' ? { __open: 'open' } : {}} />;
    }
    switch (current.key) {
      case 'mywork': return <MyWork />;
      case 'calendar': return <CalendarPage />;
      case 'forecast': return <ForecastPage />;
      case 'reports': return <ReportsPage />;
      case 'settings': return <SettingsPage tenantId={props.tenantId} host={props.host} tenantSiteUrl={props.tenantSiteUrl} />;
      case 'audit': return <AuditLogPage />;
      case 'setup': return <SetupPage embedded ds={engine.ds} statuses={statuses} canProvision onDone={() => { engine.ds.checkLists().then(setStatuses).catch(() => undefined); }} />;
      default: return <Dashboard />;
    }
  };

  const isDefaultTitle = !props.title || /^brillienta crm$/i.test(props.title.trim());
  const brand = isDefaultTitle
    ? <img src={props.isDarkTheme ? LOGO_DARK : LOGO_LIGHT} alt="Brillienta CRM" className="crm-brand-logo" />
    : <span className="crm-brand-custom"><img src={props.isDarkTheme ? MARK_DARK : MARK} alt="" className="crm-brand-mark" /><Text styles={{ root: { fontWeight: 700 } }}>{props.title}</Text></span>;
  const top = stack[stack.length - 1];
  const lic = engine.license;

  return (
    <Shell theme={theme} className={rootClass} hostId={layerHostId}>
      <CrmContext.Provider value={ctx}>
        <div ref={rootRef} className={`crm-shell${isNarrow ? ' crm-narrow' : ''}${collapsed ? ' crm-collapsed' : ''}`}>
          {!isNarrow && (
            <aside className="crm-nav">
              <div className="crm-brand">
                {collapsed && <img src={props.isDarkTheme ? MARK_DARK : MARK} alt="Brillienta CRM" className="crm-brand-mark crm-brand-collapsed" />}
                <IconButton iconProps={{ iconName: 'GlobalNavButton' }} ariaLabel={collapsed ? 'Expand navigation' : 'Collapse navigation'}
                  onClick={() => { setCollapsed(!collapsed); lsSet('brillienta-crm:nav-collapsed', !collapsed); }} />
                {collapsed ? null : brand}
              </div>
              {navList(collapsed)}
            </aside>
          )}
          <main className="crm-main">
            <div className="crm-topbar">
              {isNarrow && <IconButton iconProps={{ iconName: 'GlobalNavButton' }} ariaLabel="Open navigation" onClick={() => setNavOpen(true)} />}
              <Text variant="large" className="crm-page-title">{current.label}</Text>
              <GlobalSearch />
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 2 }}>
                {newMenu.length > 0 && <CommandButton iconProps={{ iconName: 'Add' }} text={isNarrow ? '' : 'New'} ariaLabel="New record" menuProps={{ items: newMenu }} />}
                <Notifications count={notifCount} groups={[
                  { title: 'Overdue tasks', key: 'activities', rows: myOverdue },
                  { title: 'SLA breached', key: 'cases', rows: myBreached },
                  { title: 'Demos today', key: 'demos', rows: demosToday },
                  { title: 'New leads for you', key: 'leads', rows: newForMe }
                ]} onSeen={() => lsSet('brillienta-crm:last-seen', new Date().toISOString())} />
                {canFullScreen && (
                  <IconButton iconProps={{ iconName: fullScreen ? 'BackToWindow' : 'FullScreen' }} title={fullScreen ? 'Exit full screen (show SharePoint)' : 'Full screen'}
                    ariaLabel={fullScreen ? 'Exit full screen' : 'Full screen'} onClick={toggleFullScreen} />
                )}
                <IconButton iconProps={{ iconName: 'Refresh' }} title="Refresh data" ariaLabel="Refresh data"
                  onClick={() => { Promise.all(Object.keys(data).map(k => engine.refresh(k as EntityKey))).then(() => notify('Data refreshed.')).catch(e => notify(errorText(e), MessageBarType.error)); }} />
              </Stack>
            </div>
            {(lic.status !== 'active' || (lic.daysLeft !== undefined && lic.daysLeft <= 14)) && (
              <MessageBar isMultiline={false} messageBarType={lic.canWrite ? (lic.status === 'trial' ? MessageBarType.info : MessageBarType.warning) : MessageBarType.severeWarning}
                actions={engine.can('admin') ? <Link onClick={() => ctx.navigate('settings')}>Manage licence</Link> : undefined}>
                {lic.status === 'active' ? `Licence expires in ${lic.daysLeft} day(s).` : lic.message}
              </MessageBar>
            )}
            <div className="crm-toasts" aria-live="polite">
              {toasts.map(t => <MessageBar key={t.id} messageBarType={t.type} onDismiss={() => setToasts(x => x.filter(y => y.id !== t.id))}>{t.text}</MessageBar>)}
            </div>
            <div className="crm-page">{pageBody()}</div>
          </main>
        </div>

        {isNarrow && (
          <Panel isOpen={navOpen} type={PanelType.smallFixedNear} onDismiss={() => setNavOpen(false)} isLightDismiss
            onRenderHeader={() => <div style={{ padding: '14px 20px' }}>{brand}</div>}>
            {navList(false)}
          </Panel>
        )}
        {top && <RecordPanel key={`${top.key}-${top.record.id ?? 'new'}-${stack.length}`} entity={top.key} record={top.record} canGoBack={stack.length > 1}
          onClose={() => setStack(s => s.slice(0, -1))} />}
      </CrmContext.Provider>
    </Shell>
  );
};

export default BrillientaCrm;

