// Копия несохранённого документа на случай сбоя (ADR-005, §7). В браузере — IndexedDB.
// Исходные файлы не затрагиваются: копия хранится отдельно и удаляется после сохранения.

import type { DocMeta } from './docx';

export interface RecoveryRecord {
  id: string;
  name: string;
  savedAt: number;
  /** Документ ProseMirror в виде JSON. */
  doc: Record<string, unknown>;
  meta: DocMeta;
  /** История версий (сжатая строка). */
  history?: string;
}

const DB = 'ens-text';
const STORE = 'recovery';

function open(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(STORE, { keyPath: 'id' });
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function tx<T>(mode: IDBTransactionMode, f: (s: IDBObjectStore) => IDBRequest<T>): Promise<T> {
  const db = await open();
  return new Promise((resolve, reject) => {
    const t = db.transaction(STORE, mode);
    const r = f(t.objectStore(STORE));
    t.oncomplete = () => { db.close(); resolve(r.result); };
    t.onerror = () => { db.close(); reject(t.error); };
    t.onabort = () => { db.close(); reject(t.error); };
  });
}

export const saveRecovery = (rec: RecoveryRecord) => tx('readwrite', (s) => s.put(rec)).catch(() => undefined);
export const clearRecovery = (id: string) => tx('readwrite', (s) => s.delete(id)).catch(() => undefined);
export const listRecovery = () => tx<RecoveryRecord[]>('readonly', (s) => s.getAll());
