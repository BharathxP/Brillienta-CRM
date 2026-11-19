import * as React from 'react';
import { DefaultButton, MessageBar, MessageBarType, Stack, Text, Toggle, Spinner, SpinnerSize } from '@fluentui/react';
import { useCrm } from '../CrmContext';
import { syncOutlook, syncState } from '../../services/OutlookSync';
import { errorText, fmtDate } from '../utils';

/** Per-user control for automatic Outlook email & meeting logging. */
export const OutlookSyncCard: React.FC = () => {
  const { engine, notify } = useCrm();
  const me = engine.me?.id;
  const [on, setOn] = React.useState(() => syncState.enabled(me));
  const [last, setLast] = React.useState(() => syncState.last(me));
  const [busy, setBusy] = React.useState(false);
  if (!engine.settings.integrations.outlookSync || !engine.graph.available) { return null; }

  const run = (full = false): void => {
    setBusy(true);
    syncOutlook(engine, { full }).then(r => {
      setBusy(false);
      setLast(r.at);
      notify(r.emails || r.meetings ? `Outlook: logged ${r.emails} email(s) and ${r.meetings} meeting(s) with your contacts and leads.` : 'Outlook is up to date.', MessageBarType.success);
    }).catch(e => { setBusy(false); notify(errorText(e), MessageBarType.error); });
  };
  const toggle = (v: boolean): void => {
    syncState.setEnabled(me, v);
    setOn(v);
    if (v) { run(); }
  };

  if (!on) {
    return (
      <MessageBar messageBarType={MessageBarType.info} isMultiline
        actions={<DefaultButton text="Turn on" iconProps={{ iconName: 'OutlookLogo' }} onClick={() => toggle(true)} />}>
        <b>Log your Outlook email and meetings automatically.</b> Conversations with your CRM contacts and leads appear on their timeline. Only external contacts already in the CRM are logged; internal email never is.
      </MessageBar>
    );
  }
  return (
    <Stack horizontal wrap verticalAlign="center" tokens={{ childrenGap: 12 }} className="crm-sync-bar">
      <Toggle inlineLabel label="Outlook logging" checked={on} onChange={(_, v) => toggle(!!v)} styles={{ root: { margin: 0 } }} />
      {busy ? <Spinner size={SpinnerSize.small} label="Syncing Outlook…" labelPosition="right" />
        : <Text variant="small" className="crm-muted">{last ? `Last synced ${fmtDate(last, true)}` : 'Not synced yet'} · runs every 15 minutes while the CRM is open</Text>}
      <DefaultButton text="Sync now" iconProps={{ iconName: 'Sync' }} disabled={busy} onClick={() => run()} />
    </Stack>
  );
};
