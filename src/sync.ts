// Two-way sync with a private GitHub repository through the GitHub API.
// Layout of the repository: docs/<id>.docx (the document as it is now) and docs/<id>.json
// ({ name, created, updated }). The phone keeps the key (a fine-grained token with
// Contents read/write on that one repository) in localStorage; nothing goes to this app's
// public repository.
//
// One run: delete what was deleted here, download what changed there, upload what changed here.
// If a document changed in both places, the repository's version wins and the phone's
// version is kept as a new document "… (с телефона)", so nothing is lost.
import { WordDoc } from "./docx/document";
import { isDirty, listDocs, newId, putDoc, updateDoc, type DocRecord } from "./store";

export interface SyncConfig {
  repo: string;
  token: string;
}
export type SyncState = { status: "off" | "syncing" | "ok" | "error"; last?: number; error?: string };

const CONFIG_KEY = "ip.sync";
const DELETED_KEY = "ip.sync.deleted";
const API = "https://api.github.com";

export const DEFAULT_REPO = "alex-Drop2201/individualnyi-plan-data";

function readJson<T>(key: string, fallback: T): T {
  try {
    return JSON.parse(localStorage.getItem(key) ?? "") as T;
  } catch {
    return fallback;
  }
}
function writeJson(key: string, value: unknown) {
  try {
    if (value == null) localStorage.removeItem(key);
    else localStorage.setItem(key, JSON.stringify(value));
  } catch {
    /* private mode */
  }
}

export const getConfig = () => readJson<SyncConfig | null>(CONFIG_KEY, null);

let state: SyncState = { status: getConfig() ? "ok" : "off" };
const listeners = new Set<(s: SyncState, changed: Set<string>) => void>();
/** Called after each run with the state and the ids of documents that changed here from the repository. */
export function onSync(fn: (s: SyncState, changed: Set<string>) => void) {
  listeners.add(fn);
  return () => listeners.delete(fn);
}
export const syncState = () => state;
function emit(s: SyncState, changed = new Set<string>()) {
  state = s;
  for (const fn of listeners) fn(s, changed);
}

let notice: (text: string) => void = () => {};
export function onNotice(fn: (text: string) => void) {
  notice = fn;
}

// ---------- GitHub API ----------

class HttpError extends Error {
  constructor(readonly status: number, message: string) {
    super(message);
  }
}

async function api(c: SyncConfig, path: string, init: RequestInit & { accept?: string } = {}) {
  const res = await fetch(`${API}/repos/${c.repo}/${path}`, {
    ...init,
    cache: "no-store",
    headers: {
      Authorization: `Bearer ${c.token}`,
      Accept: init.accept ?? "application/vnd.github+json",
      "X-GitHub-Api-Version": "2022-11-28",
      ...(init.body ? { "Content-Type": "application/json" } : {}),
    },
  });
  if (!res.ok) throw new HttpError(res.status, `GitHub ответил ${res.status}`);
  return res;
}

const raw = (c: SyncConfig, path: string) => api(c, `contents/${path}`, { accept: "application/vnd.github.raw" });

function toBase64(bytes: Uint8Array) {
  let s = "";
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(s);
}

async function putFile(c: SyncConfig, path: string, bytes: Uint8Array, sha: string | undefined, message: string) {
  const res = await api(c, `contents/${path}`, { method: "PUT", body: JSON.stringify({ message, content: toBase64(bytes), ...(sha ? { sha } : {}) }) });
  return ((await res.json()) as { content: { sha: string } }).content.sha;
}

async function deleteFile(c: SyncConfig, path: string, sha: string, message: string) {
  try {
    await api(c, `contents/${path}`, { method: "DELETE", body: JSON.stringify({ message, sha }) });
  } catch (e) {
    if (!(e instanceof HttpError && e.status === 404)) throw e;
  }
}

/** Checks a key before saving it. Returns an error message, or null when it works. */
export async function checkConfig(c: SyncConfig): Promise<string | null> {
  try {
    const res = await api(c, "");
    const repo = (await res.json()) as { private: boolean; permissions?: { push?: boolean } };
    if (!repo.private) return "Это хранилище публичное. Для документов нужно приватное (Private).";
    if (repo.permissions && !repo.permissions.push) return "Ключ даёт только чтение. Нужно право Contents: Read and write.";
    return null;
  } catch (e) {
    if (e instanceof HttpError && e.status === 401) return "Ключ не подходит (возможно, скопирован не целиком или истёк).";
    if (e instanceof HttpError && e.status === 404) return "Хранилище не найдено, или ключ не даёт к нему доступа.";
    return "Нет связи с GitHub. Проверьте интернет.";
  }
}

export function setConfig(c: SyncConfig | null) {
  writeJson(CONFIG_KEY, c);
  emit({ status: c ? "ok" : "off" });
  if (c) void syncNow();
}

// ---------- local deletions ----------

interface Tombstone { id: string; doc?: string; meta?: string }

/** Remembers a deleted document so the next run deletes it in the repository too. */
export function forget(d: DocRecord) {
  if (!getConfig() || !d.remoteSha) return;
  writeJson(DELETED_KEY, [...readJson<Tombstone[]>(DELETED_KEY, []), { id: d.id, doc: d.remoteSha, meta: d.metaSha }]);
  syncSoon();
}

// ---------- running ----------

let running: Promise<void> | null = null;
let again = false;
let timer = 0;

/** Syncs a little later, so a burst of edits makes one upload. */
export function syncSoon(delay = 3000) {
  if (!getConfig()) return;
  clearTimeout(timer);
  timer = window.setTimeout(() => void syncNow(), delay);
}

export function syncNow(): Promise<void> {
  const c = getConfig();
  if (!c) return Promise.resolve();
  if (running) {
    again = true;
    return running;
  }
  running = (async () => {
    let rounds = 0;
    do {
      again = false;
      rounds++;
      emit({ ...state, status: "syncing" });
      try {
        const changed = await run(c);
        emit({ status: "ok", last: Date.now() }, changed);
      } catch (e) {
        const offline = !navigator.onLine || e instanceof TypeError;
        emit({ status: "error", last: state.last, error: offline ? "нет интернета" : (e as Error).message });
      }
    } while (again && rounds < 3);
    running = null;
  })();
  return running;
}

async function run(c: SyncConfig): Promise<Set<string>> {
  const changed = new Set<string>();

  // What the repository has
  const remote = new Map<string, { doc?: string; meta?: string }>();
  try {
    const list = (await (await api(c, "contents/docs")).json()) as { name: string; sha: string }[];
    for (const f of list) {
      const m = /^([\w-]+)\.(docx|json)$/.exec(f.name);
      if (!m) continue;
      const r = remote.get(m[1]) ?? {};
      r[m[2] === "docx" ? "doc" : "meta"] = f.sha;
      remote.set(m[1], r);
    }
  } catch (e) {
    if (!(e instanceof HttpError && e.status === 404)) throw e; // no docs/ folder yet
  }

  // 1. Deletions made on the phone
  for (const t of readJson<Tombstone[]>(DELETED_KEY, [])) {
    const r = remote.get(t.id);
    if (r?.doc) await deleteFile(c, `docs/${t.id}.docx`, r.doc, "Удалён на телефоне");
    if (r?.meta) await deleteFile(c, `docs/${t.id}.json`, r.meta, "Удалён на телефоне");
    remote.delete(t.id);
    writeJson(DELETED_KEY, readJson<Tombstone[]>(DELETED_KEY, []).filter((x) => x.id !== t.id));
  }

  // 2. Changes in the repository
  const local = new Map((await listDocs()).map((d) => [d.id, d]));
  for (const [id, r] of remote) {
    if (!r.doc) continue;
    const d = local.get(id);
    const docChanged = r.doc !== d?.remoteSha;
    const metaChanged = !!r.meta && r.meta !== d?.metaSha;
    if (d && !docChanged && !metaChanged) continue;
    const meta = r.meta ? ((await (await raw(c, `docs/${id}.json`)).json()) as Partial<DocRecord>) : {};
    const file = docChanged ? await (await raw(c, `docs/${id}.docx`)).arrayBuffer() : undefined;
    if (file) await WordDoc.open(file); // never take in a broken file
    const now = Date.now();

    if (!d) {
      await putDoc({ id, name: meta.name ?? id, file: file!, created: meta.created ?? now, updated: meta.updated ?? now, remoteSha: r.doc, metaSha: r.meta, rev: 1, syncedRev: 1 });
      changed.add(id);
      continue;
    }
    let copy: DocRecord | undefined;
    await updateDoc(id, (cur) => {
      if (isDirty(cur)) {
        if (!docChanged) return { ...cur, metaSha: r.meta }; // only the name changed there; ours is newer
        copy = { ...cur, id: newId(), name: `${cur.name} (с телефона)`, remoteSha: undefined, metaSha: undefined, rev: 1, syncedRev: 0 };
      }
      return {
        ...cur,
        ...(file ? { file, xml: undefined } : {}),
        name: meta.name ?? cur.name,
        updated: meta.updated ?? now,
        remoteSha: r.doc,
        metaSha: r.meta,
        syncedRev: cur.rev ?? 1,
      };
    });
    if (copy) {
      await putDoc(copy);
      notice(`«${copy.name.replace(/ \(с телефона\)$/, "")}» изменили и здесь, и в хранилище. Взята версия из хранилища, ваша сохранена как «${copy.name}».`);
    }
    changed.add(id);
  }

  // Synced before but gone from the repository: upload again rather than lose it here.
  for (const d of local.values())
    if (d.remoteSha && !remote.has(d.id))
      await updateDoc(d.id, (cur) => ({ ...cur, remoteSha: undefined, metaSha: undefined, syncedRev: 0 }));

  // 3. Changes made on the phone
  for (const d of await listDocs()) {
    if (!isDirty(d)) continue;
    const rev = d.rev ?? 1;
    const bytes = new Uint8Array(d.xml ? await (await (await WordDoc.open(d.file, d.xml)).toDocx()).arrayBuffer() : d.file);
    const json = new TextEncoder().encode(JSON.stringify({ name: d.name, created: d.created, updated: d.updated }, null, 2) + "\n");
    let docSha: string, metaSha: string;
    try {
      docSha = await putFile(c, `docs/${d.id}.docx`, bytes, d.remoteSha, `Телефон: ${d.name}`);
      metaSha = await putFile(c, `docs/${d.id}.json`, json, d.metaSha, `Телефон: ${d.name}`);
    } catch (e) {
      // Changed in the repository meanwhile: the next run downloads it and keeps ours as a copy.
      if (e instanceof HttpError && (e.status === 409 || e.status === 422)) {
        again = true;
        continue;
      }
      throw e;
    }
    await updateDoc(d.id, (cur) => ({ ...cur, remoteSha: docSha, metaSha, syncedRev: rev }));
  }
  return changed;
}
