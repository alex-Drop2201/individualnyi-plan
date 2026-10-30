// Documents live only on this device, in IndexedDB. Each keeps the original .docx and,
// once edited, the edited word/document.xml.
import { createStore, del, entries, get, set } from "idb-keyval";

export interface DocRecord {
  id: string;
  name: string;
  file: ArrayBuffer;
  xml?: string;
  created: number;
  updated: number;
}

const store = createStore("individualnyi-plan", "docs");

export const getDoc = (id: string) => get<DocRecord>(id, store);
export const putDoc = (rec: DocRecord) => set(rec.id, rec, store);
export const deleteDoc = (id: string) => del(id, store);
export async function listDocs() {
  return (await entries<string, DocRecord>(store)).map(([, d]) => d).sort((a, b) => b.updated - a.updated);
}

export const newId = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
