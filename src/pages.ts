// Pages: where the document breaks, close to how Word does it. Breaks fall between blocks
// (paragraphs, tables) and between table rows (rows tied by a merged cell stay together),
// at Word's own page breaks, and a "keep with next" paragraph moves with what follows it.
// The editor shows pages by putting gaps (`.pg-gap`) into the one flowing column, so nothing
// being edited is ever moved; the PDF uses the same plan to cut real pages.
import type { Section } from "./docx/document";

export const PX = 4 / 3; // CSS px per pt
export const GAP = 24; // px between sheets in the editor

/** A page starts at `at` (a block, or a table row); `used` px were filled on the page before. */
export interface PageBreak {
  at: HTMLElement;
  used: number;
}

const isGap = (el: Element) => el.classList.contains("pg-gap");

function outer(el: HTMLElement) {
  const cs = getComputedStyle(el);
  return el.offsetHeight + parseFloat(cs.marginTop) + parseFloat(cs.marginBottom);
}

function rowsOf(table: HTMLTableElement) {
  return Array.from(table.tBodies[0]?.rows ?? []).filter((r) => !isGap(r));
}

/** Groups of rows that a page break must not cut: rows joined by a cell spanning them. */
export function rowUnits(table: HTMLTableElement): HTMLTableRowElement[][] {
  const rows = rowsOf(table);
  const units: HTMLTableRowElement[][] = [];
  let reach = -1;
  rows.forEach((tr, i) => {
    if (i > reach) units.push([]);
    units[units.length - 1].push(tr);
    for (const td of Array.from(tr.cells)) reach = Math.max(reach, i + td.rowSpan - 1);
  });
  return units;
}

const sum = (xs: number[]) => xs.reduce((a, b) => a + b, 0);

function tableParts(table: HTMLTableElement) {
  const units = rowUnits(table);
  const heights = units.map((u) => sum(u.map((r) => r.offsetHeight)));
  const gaps = sum(Array.from(table.tBodies[0]?.rows ?? []).filter(isGap).map((r) => r.offsetHeight));
  // Borders and the like that belong to the table rather than to a row: count them with the first rows.
  const extra = Math.max(0, outer(table) - gaps - sum(heights));
  if (heights.length) heights[0] += extra;
  return { units, heights };
}

/** Where pages begin, for blocks laid out in one column. `limit`: the height of a page's text area, px. */
export function planBreaks(blocks: HTMLElement[], limit: number): PageBreak[] {
  const out: PageBreak[] = [];
  let used = 0;
  let forced = false;
  const brk = (at: HTMLElement) => {
    out.push({ at, used });
    used = 0;
  };
  const firstPart = (b: HTMLElement) => (b instanceof HTMLTableElement ? (tableParts(b).heights[0] ?? 0) : outer(b));

  for (let i = 0; i < blocks.length; i++) {
    const b = blocks[i];
    if ((forced || b.dataset.break === "before") && used > 0) brk(b);
    forced = false;
    if (b instanceof HTMLTableElement) {
      const { units, heights } = tableParts(b);
      units.forEach((u, k) => {
        if (used + heights[k] > limit && used > 0) brk(k === 0 ? b : u[0]);
        used += heights[k];
      });
    } else {
      const own = outer(b);
      let need = own;
      // "Keep with next": the paragraph goes to the next page together with the start of what follows.
      for (let j = i; blocks[j]?.dataset.keep !== undefined && j + 1 < blocks.length; j++) {
        need += firstPart(blocks[j + 1]);
        if (blocks[j + 1].dataset.keep === undefined) break;
      }
      if (need > limit) need = own; // a chain taller than a page cannot stay together
      if (used + need > limit && used > 0) brk(b);
      used += own;
    }
    if (b.dataset.break === "after") forced = true;
  }
  return out;
}

/** Shows the document as sheets in the editor: gaps in the flow plus white sheets behind it. */
export class Pager {
  private breaks: PageBreak[] = [];
  private sheets = document.createElement("div");
  private queued = 0;

  constructor(private page: HTMLElement, private body: HTMLElement, private section: Section) {
    this.sheets.className = "sheets";
    page.prepend(this.sheets);
    new ResizeObserver(() => this.schedule()).observe(body);
    void document.fonts?.ready.then(() => this.update());
  }

  private get limit() {
    return (this.section.height - this.section.top - this.section.bottom) * PX;
  }

  private blocks() {
    return Array.from(this.body.children).filter((el): el is HTMLElement => !isGap(el));
  }

  schedule() {
    if (this.queued) return;
    this.queued = requestAnimationFrame(() => {
      this.queued = 0;
      this.update();
    });
  }

  /** Recomputes the pages; changes the gaps only when the pages changed. */
  update(force = false) {
    const breaks = planBreaks(this.blocks(), this.limit);
    const same =
      breaks.length === this.breaks.length && breaks.every((b, i) => b.at === this.breaks[i].at && Math.abs(b.used - this.breaks[i].used) < 1);
    if (same && !force) return;
    this.breaks = breaks;
    for (const g of Array.from(this.body.querySelectorAll(".pg-gap"))) g.remove();
    const { top, bottom } = this.section;
    for (const b of breaks) {
      const h = Math.max(0, this.limit - b.used) + (top + bottom) * PX + GAP;
      if (b.at instanceof HTMLTableRowElement) {
        const tr = document.createElement("tr");
        tr.className = "pg-gap";
        const td = document.createElement("td");
        td.colSpan = b.at.closest("table")?.querySelectorAll("col").length || 1;
        td.style.height = `${h}px`;
        tr.append(td);
        b.at.before(tr);
      } else {
        const gap = document.createElement("div");
        gap.className = "pg-gap";
        gap.contentEditable = "false";
        gap.style.height = `${h}px`;
        b.at.before(gap);
      }
    }
    this.paintSheets();
  }

  private paintSheets() {
    const n = this.breaks.length + 1;
    const h = this.section.height * PX;
    while (this.sheets.children.length > n) this.sheets.lastElementChild!.remove();
    while (this.sheets.children.length < n) {
      const sheet = document.createElement("div");
      const no = document.createElement("span");
      no.className = "sheet-no";
      sheet.append(no);
      this.sheets.append(sheet);
    }
    Array.from(this.sheets.children).forEach((s, i) => {
      const sheet = s as HTMLElement;
      sheet.style.top = `${i * (h + GAP)}px`;
      sheet.style.height = `${h}px`;
      sheet.firstElementChild!.textContent = `${i + 1}`;
    });
    this.page.style.minHeight = `${n * h + (n - 1) * GAP}px`;
  }

  /** The pages: where each starts, and the beginning of its text. */
  pages() {
    const starts: HTMLElement[] = [this.blocks()[0], ...this.breaks.map((b) => b.at)].filter(Boolean);
    return starts.map((start, i) => {
      // Text from this page's start up to the next page's start.
      const next = starts[i + 1];
      let text = "";
      const walker = document.createTreeWalker(this.body, NodeFilter.SHOW_TEXT);
      walker.currentNode = start;
      for (let n = walker.nextNode(); n && text.length < 80; n = walker.nextNode()) {
        if (next && (next.contains(n) || next.compareDocumentPosition(n) & Node.DOCUMENT_POSITION_FOLLOWING)) break;
        // Spaces and fill-in lines are not what a page is about.
        if (!n.parentElement?.closest("[data-skip]")) text = `${text} ${n.nodeValue}`.replace(/[\s_]+/g, " ").trimStart();
      }
      return { start, text: text.trim() };
    });
  }

  sheet(i: number) {
    return this.sheets.children[i] as HTMLElement | undefined;
  }
}
