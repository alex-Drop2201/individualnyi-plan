import "./style.css";
import { DOC_CSS } from "./docx/css";
import { DOCX_TYPE, WordDoc } from "./docx/document";
import { canShareFiles, shareFile } from "./files";
import { makePdf } from "./pdf";
import { deleteDoc, getDoc, listDocs, newId, putDoc, type DocRecord } from "./store";

const app = document.getElementById("app")!;
document.head.append(Object.assign(document.createElement("style"), { textContent: DOC_CSS }));

type Child = Node | string | null | undefined | false;
function h<K extends keyof HTMLElementTagNameMap>(tag: K, props: Partial<HTMLElementTagNameMap[K]> & { class?: string } = {}, ...children: Child[]) {
  const el = document.createElement(tag);
  const { class: cls, ...rest } = props;
  if (cls) el.className = cls;
  Object.assign(el, rest);
  for (const c of children) if (c) el.append(c);
  return el;
}

const dateFmt = new Intl.DateTimeFormat("ru", { day: "numeric", month: "long", hour: "2-digit", minute: "2-digit" });

// ---------- sheets and messages ----------

function sheet(title: string, ...content: Child[]) {
  const box = h("div", { class: "sheet" }, h("div", { class: "sheet-title" }, title), ...content);
  const overlay = h("div", { class: "overlay" }, box);
  const close = () => overlay.remove();
  overlay.onclick = (e) => e.target === overlay && close();
  box.append(h("button", { class: "btn wide", onclick: close }, "Отмена"));
  document.body.append(overlay);
  return close;
}

function actions(title: string, items: [string, () => void, string?][]) {
  const close = sheet(
    title,
    ...items.map(([label, fn, cls]) => h("button", { class: `btn wide ${cls ?? ""}`, onclick: () => { close(); fn(); } }, label)),
  );
}

function busy(text: string) {
  const el = h("div", { class: "overlay" }, h("div", { class: "busy" }, h("div", { class: "spinner" }), text));
  document.body.append(el);
  return () => el.remove();
}

async function withBusy<T>(text: string, job: () => Promise<T>) {
  const done = busy(text);
  await new Promise((r) => setTimeout(r, 30)); // let the message show before heavy work
  try {
    return await job();
  } catch (e) {
    alert(`Не получилось: ${(e as Error).message}`);
  } finally {
    done();
  }
}

/** Sharing needs a fresh tap after the file is made, so the file is offered on a button. */
function offer(blob: Blob, name: string) {
  if (!canShareFiles()) {
    void shareFile(blob, name);
    return;
  }
  const close = sheet(
    "Файл готов",
    h("div", { class: "sheet-note" }, name),
    h("button", { class: "btn wide primary", onclick: () => { close(); void shareFile(blob, name); } }, "Сохранить или отправить"),
  );
}

const fileName = (rec: DocRecord, ext: string) => `${rec.name.replace(/[\\/:*?"<>|]/g, "_")}.${ext}`;

async function exportDocx(rec: DocRecord, doc?: WordDoc) {
  const blob = await withBusy("Готовлю файл Word…", async () => (doc ?? (await WordDoc.open(rec.file, rec.xml))).toDocx());
  if (blob) offer(blob, fileName(rec, "docx"));
}

async function exportPdf(rec: DocRecord, doc?: WordDoc) {
  const blob = await withBusy("Готовлю PDF…", async () => makePdf(doc ?? (await WordDoc.open(rec.file, rec.xml))));
  if (blob) offer(blob, fileName(rec, "pdf"));
}

// ---------- list ----------

const picker = h("input", { type: "file", accept: `.docx,${DOCX_TYPE}`, hidden: true });
picker.onchange = () => {
  const f = picker.files?.[0];
  picker.value = "";
  if (f) void importFile(f);
};
document.body.append(picker);

async function importFile(f: File) {
  if (!/\.docx$/i.test(f.name) && f.type !== DOCX_TYPE) {
    alert(
      /\.(doc|pdf|pages|odt|rtf)$/i.test(f.name)
        ? "Нужен файл Word в формате .docx. Откройте этот файл в Word или Pages и сохраните (экспортируйте) его как .docx, затем загрузите сюда."
        : "Это не документ Word (.docx).",
    );
    return;
  }
  const file = await f.arrayBuffer();
  const rec = await withBusy("Открываю…", async () => {
    await WordDoc.open(file); // checks the file before keeping it
    const now = Date.now();
    const rec: DocRecord = { id: newId(), name: f.name.replace(/\.docx$/i, ""), file, created: now, updated: now };
    await putDoc(rec);
    return rec;
  });
  if (rec) location.hash = `doc=${rec.id}`;
}

async function rename(rec: DocRecord) {
  const name = prompt("Название документа", rec.name)?.trim();
  if (!name || name === rec.name) return false;
  rec.name = name;
  await putDoc(rec);
  return true;
}

async function showList() {
  const docs = await listDocs();
  const items = docs.map((d) =>
    h(
      "li",
      { class: "doc" },
      h(
        "button",
        { class: "doc-open", onclick: () => (location.hash = `doc=${d.id}`) },
        h("span", { class: "doc-name" }, d.name),
        h("span", { class: "doc-date" }, `Изменён ${dateFmt.format(d.updated)}`),
      ),
      h("button", { class: "icon-btn", ariaLabel: "Действия", onclick: () => docActions(d), textContent: "⋯" }),
    ),
  );
  app.replaceChildren(
    h(
      "header",
      { class: "bar" },
      h("h1", {}, "Документы"),
      h("button", { class: "btn primary", onclick: () => picker.click() }, "Открыть файл"),
    ),
    h(
      "main",
      { class: "list" },
      docs.length
        ? h("ul", { class: "docs" }, ...items)
        : h(
            "div",
            { class: "empty" },
            h("p", {}, "Здесь пока нет документов."),
            h("p", {}, "Нажмите «Открыть файл» и выберите документ Word (.docx). Он сохранится на этом телефоне, и его можно будет править и сохранять в Word или PDF."),
          ),
    ),
  );
}

function docActions(d: DocRecord) {
  actions(d.name, [
    ["Открыть", () => (location.hash = `doc=${d.id}`)],
    ["Сохранить как Word (.docx)", () => void exportDocx(d)],
    ["Сохранить как PDF", () => void exportPdf(d)],
    ["Сделать копию", async () => {
      const now = Date.now();
      await putDoc({ ...d, id: newId(), name: `${d.name} (копия)`, created: now, updated: now });
      void showList();
    }],
    ["Переименовать", async () => { if (await rename(d)) void showList(); }],
    ["Удалить", async () => {
      if (!confirm(`Удалить «${d.name}»? Это нельзя отменить.`)) return;
      await deleteDoc(d.id);
      void showList();
    }, "danger"],
  ]);
}

// ---------- editor ----------

const ZOOMS = [1, 1.5, 2, 2.5, 3];
let saveNow: (() => Promise<void>) | null = null;

async function showEditor(id: string) {
  const rec = await getDoc(id);
  if (!rec) {
    location.hash = "";
    return;
  }
  let doc: WordDoc;
  try {
    doc = await WordDoc.open(rec.file, rec.xml);
  } catch (e) {
    alert(`Не удалось открыть документ: ${(e as Error).message}`);
    location.hash = "";
    return;
  }

  const { width, top, right, bottom, left, height } = doc.section;
  const body = h("div", { class: "body" }, ...doc.render(true));
  const page = h("div", { class: "page editing" }, body);
  Object.assign(page.style, { width: `${width}pt`, minHeight: `${height}pt`, padding: `${top}pt ${right}pt ${bottom}pt ${left}pt` });
  const zoomBox = h("div", { class: "zoom" }, page);
  const scroller = h("main", { class: "scroller" }, zoomBox);

  const status = h("span", { class: "status" }, "Сохранено");
  const title = h("button", { class: "title", onclick: async () => { if (await rename(rec)) title.textContent = rec.name; } }, rec.name);

  let zi = 0;
  try { zi = Math.min(Math.max(0, +(localStorage.getItem("ip.zoom") ?? 0) || 0), ZOOMS.length - 1); } catch { /* private mode */ }
  const fit = () => (scroller.clientWidth - 16) / (width * (4 / 3));
  const applyZoom = () => {
    zoomBox.style.zoom = String(fit() * ZOOMS[zi]);
    minus.disabled = zi === 0;
    plus.disabled = zi === ZOOMS.length - 1;
    try { localStorage.setItem("ip.zoom", String(zi)); } catch { /* private mode */ }
  };
  const minus = h("button", { class: "icon-btn", ariaLabel: "Мельче", textContent: "−", onclick: () => { zi = Math.max(0, zi - 1); applyZoom(); } });
  const plus = h("button", { class: "icon-btn", ariaLabel: "Крупнее", textContent: "+", onclick: () => { zi = Math.min(ZOOMS.length - 1, zi + 1); applyZoom(); } });

  // Saving: every edit updates the XML at once; it is written to storage shortly after.
  let timer = 0;
  const save = async () => {
    clearTimeout(timer);
    timer = 0;
    rec.xml = doc.serialize();
    rec.updated = Date.now();
    await putDoc(rec);
    status.textContent = "Сохранено";
  };
  saveNow = () => (timer ? save() : Promise.resolve());
  const changed = () => {
    status.textContent = "Сохраняю…";
    clearTimeout(timer);
    timer = window.setTimeout(() => void save(), 700);
  };

  const paraOf = (t: EventTarget | null) => (t instanceof Element ? (t.closest(".p") as HTMLElement | null) : null);
  body.addEventListener("beforeinput", (e) => {
    if (e.inputType.startsWith("format")) e.preventDefault();
    else if (e.inputType === "insertParagraph") {
      // Enter starts a new line inside the same paragraph (a line break in Word).
      e.preventDefault();
      document.execCommand("insertLineBreak");
    }
  });
  body.addEventListener("paste", (e) => {
    e.preventDefault();
    const text = e.clipboardData?.getData("text/plain") ?? "";
    document.execCommand("insertText", false, text.replace(/\r\n?/g, "\n"));
  });
  body.addEventListener("drop", (e) => e.preventDefault());
  body.addEventListener("input", (e) => {
    const div = paraOf(e.target);
    if (div && doc.sync(div)) changed();
  });
  body.addEventListener("focusout", (e) => {
    const div = paraOf(e.target);
    if (!div) return;
    if (doc.sync(div)) changed();
    doc.refresh(div);
  });

  const flush = () => {
    const div = paraOf(document.activeElement);
    if (div && doc.sync(div)) changed();
    return saveNow?.();
  };
  const back = async () => {
    await flush();
    location.hash = "";
  };

  let hint: HTMLElement | null = null;
  if (!localStorage.getItem("ip.hint")) {
    hint = h(
      "div",
      { class: "hint" },
      h("span", {}, "Нажмите на любой текст, чтобы изменить его. В пустые клетки таблиц и строки «____» тоже можно писать. Всё сохраняется само."),
      h("button", { class: "icon-btn", ariaLabel: "Скрыть", textContent: "✕", onclick: () => { hint?.remove(); try { localStorage.setItem("ip.hint", "1"); } catch { /* */ } } }),
    );
  }

  app.replaceChildren(
    h("header", { class: "bar" }, h("button", { class: "btn back", onclick: back }, "‹ Документы"), h("div", { class: "titles" }, title, status)),
    ...(hint ? [hint] : []),
    scroller,
    h(
      "footer",
      { class: "tools" },
      minus,
      plus,
      h("span", { class: "spacer" }),
      h("button", { class: "btn", onclick: async () => { await flush(); void exportDocx(rec, doc); } }, "Word"),
      h("button", { class: "btn primary", onclick: async () => { await flush(); void exportPdf(rec, doc); } }, "PDF"),
    ),
  );
  applyZoom();
  window.onresize = applyZoom;
}

// ---------- start ----------

async function route() {
  await saveNow?.();
  saveNow = null;
  window.onresize = null;
  const id = /doc=([\w-]+)/.exec(location.hash)?.[1];
  if (id) await showEditor(id);
  else await showList();
}
window.addEventListener("hashchange", () => void route());
document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "hidden") void saveNow?.();
});
window.addEventListener("pagehide", () => void saveNow?.());
void navigator.storage?.persist?.();
void route();

if ("serviceWorker" in navigator && import.meta.env.PROD) void navigator.serviceWorker.register("./sw.js");
