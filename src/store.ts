// Documents live on this device, in IndexedDB. Each keeps the original .docx and, once
// edited, the edited word/document.xml. With sync on (src/sync.ts) they are also copied to
// a private GitHub repository.
import { createStore, del, entries, get, set, update } from "idb-keyval";

export interface DocRecord {
  id: string;
  name: string;
  file: ArrayBuffer;
  xml?: string;
  created: number;
  updated: number;
  /** Counts local changes (missing = 1). The document needs uploading while rev ≠ syncedRev. */
  rev?: number;
  syncedRev?: number;
  /** Git blob shas of docs/<id>.docx and docs/<id>.json last seen in the sync repository. */
  remoteSha?: string;
  metaSha?: string;
}

const store = createStore("individualnyi-plan", "docs");

export const getDoc = (id: string) => get<DocRecord>(id, store);
export const putDoc = (rec: DocRecord) => set(rec.id, rec, store);
export const deleteDoc = (id: string) => del(id, store);
export async function listDocs() {
  return (await entries<string, DocRecord>(store)).map(([, d]) => d).sort((a, b) => b.updated - a.updated);
}

/** Read-modify-write in one transaction, so the editor and sync never overwrite each other. */
export function updateDoc(id: string, fn: (cur: DocRecord) => DocRecord) {
  return update<DocRecord | undefined>(id, (cur) => (cur ? fn(cur) : cur), store);
}

/** A change made on this phone: bumps the revision so sync uploads it. */
export function changeDoc(id: string, patch: Partial<DocRecord>) {
  return updateDoc(id, (cur) => ({ ...cur, ...patch, updated: Date.now(), rev: (cur.rev ?? 1) + 1 }));
}

export const isDirty = (d: DocRecord) => (d.rev ?? 1) !== (d.syncedRev ?? 0);

export const newId = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
