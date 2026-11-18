import * as React from 'react';
import { MessageBarType } from '@fluentui/react';
import { Engine } from '../services/Engine';
import { CrmData, EntityKey, IRecord } from '../models/model';
import { ICrmSettings } from '../models/settings';

export interface IOpenRecord {
  key: EntityKey;
  record: IRecord;
}

export interface ICrmContext {
  engine: Engine;
  data: CrmData;
  settings: ICrmSettings;
  currency: string;
  /** Open a record (or a new one, without id) in the side panel. */
  open: (key: EntityKey, record?: IRecord) => void;
  navigate: (page: string) => void;
  notify: (message: string, type?: MessageBarType) => void;
  updateSettings: (s: ICrmSettings) => Promise<void>;
  isNarrow: boolean;
}

export const CrmContext = React.createContext<ICrmContext | undefined>(undefined);

export function useCrm(): ICrmContext {
  const ctx = React.useContext(CrmContext);
  if (!ctx) { throw new Error('CrmContext missing'); }
  return ctx;
}

/** Loads the given entities on demand and returns the visible records. */
export function useEntities(...keys: EntityKey[]): { loading: boolean; rows: (key: EntityKey) => IRecord[] } {
  const { engine, data, notify } = useCrm();
  const missing = keys.filter(k => engine.hasFeature(k) && !data[k]);
  const sig = missing.join(',');
  React.useEffect(() => {
    if (missing.length) {
      engine.ensure(missing).catch(e => notify(String((e as Error).message || e), MessageBarType.error));
    }
  }, [sig]);
  return { loading: missing.length > 0, rows: (k: EntityKey) => (data[k] ? engine.visible(k) : []) };
}
