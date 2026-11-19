import * as React from 'react';
import { Link, Stack, Text, Icon } from '@fluentui/react';
import { IRecord } from '../../models/model';
import { useCrm, useEntities } from '../CrmContext';

/** Knowledge articles that match the case, with one-click reply and linking. */
export const KbSuggestions: React.FC<{ caseRecord: IRecord; onReply: (article: IRecord) => void; onLinked: (c: IRecord) => void }> = ({ caseRecord, onReply, onLinked }) => {
  const { engine, open, notify } = useCrm();
  useEntities('articles');
  if (!engine.hasFeature('articles')) { return null; }
  const linked = caseRecord.articleRef ? engine.find('articles', caseRecord.articleRef) : undefined;
  const list = linked ? [linked] : engine.suggestArticles(caseRecord);
  if (!list.length) { return null; }
  const canEdit = engine.can('edit', 'cases');
  return (
    <div className="crm-smart" role="region" aria-label="Knowledge base">
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
        <Icon iconName="KnowledgeArticle" />
        <Text styles={{ root: { fontWeight: 600 } }}>{linked ? 'Linked knowledge article' : 'Suggested answers from the knowledge base'}</Text>
      </Stack>
      <ul className="crm-smart-list">
        {list.map(a => (
          <li key={a.id}>
            <Link onClick={() => open('articles', a)}>{a.title}</Link>
            {canEdit && <> · <Link onClick={() => onReply(a)}>Reply with this</Link></>}
            {canEdit && !linked && <> · <Link onClick={() => { engine.linkArticle(caseRecord, a).then(c => { onLinked(c); notify('Article linked to the case.'); }).catch(() => undefined); }}>Link</Link></>}
          </li>
        ))}
      </ul>
    </div>
  );
};
