import * as React from 'react';
import styles from './CrmCopilot.module.scss';
import { MARK, MARK_DARK } from '../../webparts/brillientaCrm/components/logo';
import type { IDealHealth } from '../../webparts/brillientaCrm/services/DashboardAI';
import type { EntityKey } from '../../webparts/brillientaCrm/models/model';
import { CrmNotConfiguredError, ICrmSession, recordUrl } from './crmSession';

export type Tone = 'good' | 'warn' | 'bad' | 'info' | 'neutral';

/** What a tool's view gets from the Copilot host. */
export interface ICrmToolHost {
  session: ICrmSession;
  dark: boolean;
  fullscreen: boolean;
  /** Opens a URL through the Copilot host (iframes cannot navigate on their own). */
  openLink: (url: string) => void;
  /** Present when the host can show this tool full screen. */
  expand?: () => void;
  /** Tells Copilot what is on screen so it can answer follow-up questions about it. */
  setModelContext: (summary: string, structured?: Record<string, unknown>) => void;
  /** Sends a follow-up prompt to Copilot as the user. */
  sendPrompt: (text: string) => void;
}

const cx = (...c: (string | false | undefined)[]): string => c.filter(Boolean).join(' ');

export const Frame: React.FC<{ dark: boolean; title: string; subtitle?: string; children?: React.ReactNode }> = ({ dark, title, subtitle, children }) => (
  <div className={cx(styles.root, dark && styles.dark)} role="region" aria-label={title}>
    <div className={styles.header}>
      <img className={styles.mark} src={dark ? MARK_DARK : MARK} alt="" />
      <div className={styles.headText}>
        <h2 className={styles.title}>{title}</h2>
        {subtitle && <p className={styles.subtitle} title={subtitle}>{subtitle}</p>}
      </div>
    </div>
    {children}
  </div>
);

export const Loading: React.FC<{ text?: string }> = ({ text }) => <div className={styles.spinner} role="status">{text || 'Loading from Brillienta CRM…'}</div>;

export const ErrorBox: React.FC<{ error: unknown; onRetry?: () => void }> = ({ error, onRetry }) => {
  const notConfigured = error instanceof CrmNotConfiguredError;
  const status = (error as { status?: number })?.status;
  const text = notConfigured ? (error as Error).message
    : status === 403 || status === 401 ? 'You do not have access to the Brillienta CRM site. Ask a CRM administrator to add you to a CRM group.'
    : status === 404 ? 'The CRM lists were not found on the CRM site. Open the CRM once as an administrator to set it up.'
    : `Brillienta CRM could not be loaded: ${(error as Error)?.message || String(error)}`;
  return (
    <div>
      <div className={styles.note} role="alert">{text}</div>
      {onRetry && !notConfigured && <div className={styles.actions}><button type="button" className={styles.button} onClick={onRetry}>Try again</button></div>}
    </div>
  );
};

export interface IKpi { label: string; value: string; tone?: Tone }

export const Kpis: React.FC<{ items: IKpi[] }> = ({ items }) => (
  <div className={styles.kpis}>
    {items.map(k => (
      <div key={k.label} className={cx(styles.kpi, k.tone === 'bad' && styles.kpiBad, k.tone === 'warn' && styles.kpiWarn, k.tone === 'good' && styles.kpiGood)}>
        <div className={styles.kpiLabel}>{k.label}</div>
        <div className={styles.kpiValue} title={k.value}>{k.value}</div>
      </div>
    ))}
  </div>
);

export const Badge: React.FC<{ tone?: Tone; children: React.ReactNode; title?: string }> = ({ tone = 'neutral', children, title }) => (
  <span className={cx(styles.badge, tone !== 'neutral' && styles[tone])} title={title}>{children}</span>
);

export const HealthBadge: React.FC<{ health: IDealHealth }> = ({ health }) => {
  const tone: Tone = health.label === 'Healthy' ? 'good' : health.label === 'Watch' ? 'warn' : 'bad';
  const why = [...health.risks.map(r => `− ${r}`), ...health.strengths.map(s => `+ ${s}`)].join('\n');
  return <Badge tone={tone} title={why || health.label}>{health.score}% · {health.label}</Badge>;
};

export const Note: React.FC<{ children: React.ReactNode }> = ({ children }) => <div className={styles.note}>{children}</div>;

export const Empty: React.FC<{ children: React.ReactNode }> = ({ children }) => <div className={styles.empty}>{children}</div>;

export const Section: React.FC<{ title: string; children: React.ReactNode }> = ({ title, children }) => (
  <section className={styles.section}>
    <h3 className={styles.sectionTitle}>{title}</h3>
    {children}
  </section>
);

export interface IRowProps {
  title: string;
  meta?: string;
  detail?: React.ReactNode;
  amount?: string;
  badge?: React.ReactNode;
  onOpen?: () => void;
  openLabel?: string;
}

export const RowList: React.FC<{ children: React.ReactNode; label: string }> = ({ children, label }) => <ul className={styles.list} aria-label={label}>{children}</ul>;

export const Row: React.FC<IRowProps> = ({ title, meta, detail, amount, badge, onOpen, openLabel }) => (
  <li className={styles.row}>
    <div className={styles.rowMain}>
      <div className={styles.rowTitle} title={title}>{title}</div>
      {meta && <div className={styles.rowMeta} title={meta}>{meta}</div>}
      {detail && <div className={styles.rowDetail}>{detail}</div>}
    </div>
    <div className={styles.rowRight}>
      {amount && <span className={styles.amount}>{amount}</span>}
      {badge}
      {onOpen && <button type="button" className={styles.linkButton} onClick={onOpen} aria-label={openLabel || `Open ${title} in Brillienta CRM`}>Open</button>}
    </div>
  </li>
);

export const Actions: React.FC<{ children: React.ReactNode }> = ({ children }) => <div className={styles.actions}>{children}</div>;

export const Button: React.FC<{ onClick: () => void; primary?: boolean; disabled?: boolean; children: React.ReactNode; type?: 'button' | 'submit' }> = ({ onClick, primary, disabled, children, type = 'button' }) => (
  <button type={type} className={primary ? styles.primary : styles.button} onClick={onClick} disabled={disabled}>{children}</button>
);

/** Number of rows shown inline before "Show all". */
export const INLINE_ROWS = 6;

/** Shows the first rows inline; "Show all" opens full screen when the host supports it, otherwise expands in place. */
export function useRowLimit<T>(items: T[], host: ICrmToolHost): { visible: T[]; more: React.ReactNode } {
  const [all, setAll] = React.useState(false);
  if (host.fullscreen || all || items.length <= INLINE_ROWS) { return { visible: items, more: undefined }; }
  const showAll = (): void => (host.expand ? host.expand() : setAll(true));
  return {
    visible: items.slice(0, INLINE_ROWS),
    more: <Actions><Button onClick={showAll}>Show all {items.length}</Button></Actions>
  };
}

export const openRecord = (host: ICrmToolHost, key: EntityKey, id: number): (() => void) => (): void => host.openLink(recordUrl(host.session, key, id));

export { styles };
