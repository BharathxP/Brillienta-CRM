import * as React from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { BaseCopilotComponent, createCopilotTextContent } from '@microsoft/sp-copilot-component';
import type { EntityKey } from '../../webparts/brillientaCrm/models/model';
import { getCrmSession, ICrmSession } from './crmSession';
import { ErrorBox, Frame, ICrmToolHost, Loading } from './ui';

interface IToolFrameProps {
  title: string;
  dark: boolean;
  lists: EntityKey[];
  load: () => Promise<ICrmSession>;
  makeHost: (session: ICrmSession) => ICrmToolHost;
  renderView: (host: ICrmToolHost) => React.ReactNode;
}

/** Opens the CRM session, loads the lists the tool needs, then renders the tool's view. */
const ToolFrame: React.FC<IToolFrameProps> = ({ title, dark, lists, load, makeHost, renderView }) => {
  const [state, setState] = React.useState<{ session?: ICrmSession; error?: unknown }>({});
  const [attempt, setAttempt] = React.useState(0);
  React.useEffect(() => {
    let live = true;
    setState({});
    load()
      .then(async s => { await s.engine.ensure(lists); return s; })
      .then(s => { if (live) { setState({ session: s }); } })
      .catch(e => { if (live) { setState({ error: e }); } });
    return () => { live = false; };
  }, [attempt]);
  const site = state.session?.engine.settings.company.name || state.session?.siteUrl.replace(/^https:\/\//i, '');
  return (
    <Frame dark={dark} title={title} subtitle={site ? `Brillienta CRM · ${site}` : 'Brillienta CRM'}>
      {state.error ? <ErrorBox error={state.error} onRetry={() => setAttempt(n => n + 1)} />
        : state.session ? renderView(makeHost(state.session))
        : <Loading />}
    </Frame>
  );
};

/**
 * Base for every Brillienta CRM Copilot Component: one React 18 root, the Copilot theme, the CRM session,
 * links through the Copilot host and model-context updates. Subclasses only describe their view.
 */
export default abstract class CrmCopilotComponentBase<TProperties> extends BaseCopilotComponent<TProperties> {
  private _root: Root | undefined;

  /** Heading shown in the card. */
  protected abstract readonly toolTitle: string;
  /** CRM lists the view reads (loaded once per session). */
  protected abstract readonly lists: EntityKey[];
  /** The tool's view, given the normalised prompt properties. */
  protected abstract renderView(host: ICrmToolHost, properties: unknown): React.ReactNode;

  protected render(): void {
    if (!this._root) { this._root = createRoot(this.context.domElement); }
    const dark = this.hostContext.theme === 'dark';
    const fullscreen = this.hostContext.displayMode === 'fullscreen';
    const canExpand = !fullscreen && (this.hostContext.availableDisplayModes || []).indexOf('fullscreen') >= 0;
    const properties = this.properties as unknown;
    // New prompt values start the view afresh (for example, a new draft lead).
    const key = JSON.stringify(properties || {});
    this._root.render(
      <ToolFrame key={key} title={this.toolTitle} dark={dark} lists={this.lists}
        load={() => getCrmSession(this.context)}
        makeHost={session => ({
          session, dark, fullscreen,
          openLink: this._openLink,
          expand: canExpand ? this._expand : undefined,
          setModelContext: this._setModelContext,
          sendPrompt: this._sendPrompt
        })}
        renderView={host => this.renderView(host, properties)} />
    );
  }

  protected async onTeardown(reason: string | undefined): Promise<void> {
    this._root?.unmount();
    this._root = undefined;
    await super.onTeardown(reason);
  }

  private _openLink = (url: string): void => {
    this.context.copilotBridge.openLinkAsync(url).catch(() => undefined);
  };

  private _expand = (): void => {
    this.requestDisplayModeAsync('fullscreen').catch(() => undefined);
  };

  private _lastContext = '';

  private _setModelContext = (summary: string, structured?: Record<string, unknown>): void => {
    if (summary === this._lastContext) { return; }
    this._lastContext = summary;
    this.context.copilotBridge.updateModelContextAsync({
      content: [createCopilotTextContent(summary)],
      structuredContent: structured
    }).catch(() => undefined);
  };

  private _sendPrompt = (text: string): void => {
    this.context.copilotBridge.sendFollowUpMessageAsync([createCopilotTextContent(text)]).catch(() => undefined);
  };
}
