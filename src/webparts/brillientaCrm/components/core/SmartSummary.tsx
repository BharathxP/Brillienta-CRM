import * as React from 'react';
import { DefaultButton, Icon, Stack, Text, TooltipHost } from '@fluentui/react';
import { EntityKey, IRecord } from '../../models/model';
import { copilotPrompt, recordBrief } from '../../services/Insights';
import { useCrm } from '../CrmContext';
import { lsGet, lsSet } from '../utils';

export const COPILOT_CHAT_URL = 'https://m365.cloud.microsoft/chat';

/** Summary, risks and next best actions for a record, with a follow-up draft and a Copilot hand-off. */
export const SmartSummary: React.FC<{ entity: EntityKey; record: IRecord; onDraft: () => void }> = ({ entity, record, onDraft }) => {
  const { engine, data, notify } = useCrm();
  const [open, setOpen] = React.useState<boolean>(() => lsGet('brillienta-crm:summary-open', true));
  const brief = React.useMemo(() => recordBrief(engine, entity, record), [record, data]);
  const toggle = (): void => { setOpen(!open); lsSet('brillienta-crm:summary-open', !open); };
  const askCopilot = (): void => {
    const prompt = copilotPrompt(engine, entity, record);
    const done = (): void => { notify('Prompt copied. Paste it into Microsoft 365 Copilot (Ctrl+V).'); window.open(COPILOT_CHAT_URL, '_blank', 'noopener'); };
    if (navigator.clipboard) { navigator.clipboard.writeText(prompt).then(done).catch(() => window.prompt('Copy this prompt into Copilot', prompt)); } else { window.prompt('Copy this prompt into Copilot', prompt); }
  };
  return (
    <div className="crm-smart" role="region" aria-label="Smart summary">
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} className="crm-smart-head" onClick={toggle} style={{ cursor: 'pointer' }}>
        <Icon iconName="Lightbulb" />
        <Text styles={{ root: { fontWeight: 600, flex: 1 } }}>{brief.headline}</Text>
        <Icon iconName={open ? 'ChevronUp' : 'ChevronDown'} aria-label={open ? 'Collapse' : 'Expand'} />
      </Stack>
      {open && (
        <>
          <ul className="crm-smart-list">
            {brief.facts.map((x, i) => <li key={`f${i}`}>{x}</li>)}
            {brief.risks.map((x, i) => <li key={`r${i}`} className="crm-smart-risk"><Icon iconName="Warning" /> {x}</li>)}
            {brief.next.map((x, i) => <li key={`n${i}`} className="crm-smart-next"><Icon iconName="Forward" /> {x}</li>)}
          </ul>
          <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
            <DefaultButton iconProps={{ iconName: 'Edit' }} text="Draft follow-up" onClick={onDraft} />
            <TooltipHost content="Copies a prompt with this record's details and opens Microsoft 365 Copilot. Copilot uses your organisation's own licence and data protection.">
              <DefaultButton iconProps={{ iconName: 'Robot' }} text="Ask Copilot" onClick={askCopilot} />
            </TooltipHost>
          </Stack>
        </>
      )}
    </div>
  );
};
