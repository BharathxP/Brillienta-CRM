// Offline support: the last loaded copy of every list and the user's session are kept in IndexedDB,
// and changes made without a connection are queued and replayed when the connection returns.
/* eslint-disable @typescript-eslint/no-explicit-any */

const DB = 'brillienta-crm';
const VERSION = 1;
let dbp: Promise<IDBDatabase | undefined> | undefined;
const memory: Record<string, any> = {};
let memQueue: IQueued[] = [];
let memSeq = 1;

export interface IQueued { seq?: number; scope: string; key: string; record: any; tempId?: number; at: string }

function open(): Promise<IDBDatabase | undefined> {
  if (!dbp) {
    dbp = new Promise(resolve => {
      try {
        const req = window.indexedDB.open(DB, VERSION);
        req.onupgradeneeded = () => {
          const db = req.result;
          if (!db.objectStoreNames.contains('kv')) { db.createObjectStore('kv'); }
          if (!db.objectStoreNames.contains('queue')) { db.createObjectStore('queue', { keyPath: 'seq', autoIncrement: true }); }
        };
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => resolve(undefined);
      } catch { resolve(undefined); }
    });
  }
  return dbp;
}

function tx<T>(store: string, mode: IDBTransactionMode, fn: (s: IDBObjectStore) => IDBRequest): Promise<T | undefined> {
  return open().then(db => new Promise<T | undefined>(resolve => {
    if (!db) { resolve(undefined); return; }
    try {
      const r = fn(db.transaction(store, mode).objectStore(store));
      r.onsuccess = () => resolve(r.result as T);
      r.onerror = () => resolve(undefined);
    } catch { resolve(undefined); }
  }));
}

export const offlineStore = {
  async get<T>(key: string): Promise<T | undefined> {
    const v = await tx<T>('kv', 'readonly', s => s.get(key));
    return v !== undefined ? v : memory[key];
  },
  async put(key: string, value: any): Promise<void> {
    memory[key] = value;
    await tx('kv', 'readwrite', s => s.put(value, key));
  },
  async enqueue(item: IQueued): Promise<void> {
    const db = await open();
    if (!db) { memQueue.push({ ...item, seq: memSeq++ }); return; }
    await tx('queue', 'readwrite', s => s.add(item));
  },
  async queue(scope: string): Promise<IQueued[]> {
    const db = await open();
    const all = db ? ((await tx<IQueued[]>('queue', 'readonly', s => s.getAll())) || []) : memQueue;
    return all.filter(q => q.scope === scope).sort((a, b) => (a.seq || 0) - (b.seq || 0));
  },
  async remove(seq: number): Promise<void> {
    const db = await open();
    if (!db) { memQueue = memQueue.filter(q => q.seq !== seq); return; }
    await tx('queue', 'readwrite', s => s.delete(seq));
  },
  async update(item: IQueued): Promise<void> {
    const db = await open();
    if (!db) { memQueue = memQueue.map(q => (q.seq === item.seq ? item : q)); return; }
    await tx('queue', 'readwrite', s => s.put(item));
  }
};

/** True for errors caused by a missing connection (not by SharePoint refusing the request). */
export function isNetworkError(e: unknown): boolean {
  if (typeof navigator !== 'undefined' && navigator.onLine === false) { return true; }
  const m = String((e as any)?.message || e || '');
  return e instanceof TypeError || /failed to fetch|networkerror|network request failed|load failed|internet disconnected|ERR_INTERNET/i.test(m);
}
