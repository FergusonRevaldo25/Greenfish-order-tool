/** Remembers an uploaded template between visits (IndexedDB). If the browser blocks storage, the app still works for the session. */
const DB = "greenfish", STORE = "kv", KEY = "template";
export type SavedTemplate = { name: string; buf: ArrayBuffer; savedAt: number };

const open = () => new Promise<IDBDatabase>((res, rej) => {
  const rq = indexedDB.open(DB, 1);
  rq.onupgradeneeded = () => rq.result.createObjectStore(STORE);
  rq.onsuccess = () => res(rq.result);
  rq.onerror = () => rej(rq.error);
});
const run = async <T,>(mode: IDBTransactionMode, f: (s: IDBObjectStore) => IDBRequest<T>) => {
  const db = await open();
  return new Promise<T>((res, rej) => {
    const rq = f(db.transaction(STORE, mode).objectStore(STORE));
    rq.onsuccess = () => { db.close(); res(rq.result); };
    rq.onerror = () => { db.close(); rej(rq.error); };
  });
};

export const loadSavedTemplate = async (): Promise<SavedTemplate | null> => {
  try { return ((await run("readonly", s => s.get(KEY))) as SavedTemplate | undefined) ?? null; } catch { return null; }
};
export const saveTemplate = async (t: SavedTemplate) => { try { await run("readwrite", s => s.put(t, KEY)); return true; } catch { return false; } };
export const clearSavedTemplate = async () => { try { await run("readwrite", s => s.delete(KEY)); } catch {} };
