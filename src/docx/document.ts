// A .docx opened for editing. The original file stays as it is; only the text inside
// word/document.xml changes, so fonts, tables and layout come back exactly as they were.
//
// Rendering: every w:p becomes its own editable <div data-p>, every editable w:r a
// <span data-r> inside it. Runs that hold anything besides text (pictures, fields, page
// breaks, hidden text) become locked spans (data-l) that the user cannot change.
// sync() reads a paragraph's DOM back into its runs; refresh() redraws it from the XML.
import JSZip from "jszip";
import { Numbering } from "./numbering";
import {
  kid, kids, merge, readBorders, readPPr, readRPr, Styles, Theme, val, W, wattr,
  type Border, type Borders, type ParaProps, type RunProps,
} from "./props";

const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const V = "urn:schemas-microsoft-com:vml";
const XML_NS = "http://www.w3.org/XML/1998/namespace";
const DOC_PATH = "word/document.xml";
export const DOCX_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

/** Run children that are plain text: such runs can be edited and rewritten. */
const TEXTY = new Set(["rPr", "t", "tab", "br", "cr", "noBreakHyphen", "softHyphen", "lastRenderedPageBreak"]);
/** Paragraph children that only wrap runs. */
const RUN_WRAPPERS = new Set(["hyperlink", "smartTag", "ins", "customXml", "fldSimple", "moveTo", "dir", "bdo"]);

const pt = (twips: number) => `${twips / 20}pt`;
const isPageBreak = (br: Element) => ["page", "column"].includes(wattr(br, "type") ?? "");

export interface Section {
  width: number;
  height: number;
  top: number;
  bottom: number;
  left: number;
  right: number;
}

export function fontStack(font?: string) {
  const sans = /calibri|arial|helvetica|verdana|tahoma|segoe|sans/i.test(font ?? "");
  const fallback = sans ? `Arial, "Liberation Sans", sans-serif` : `"Times New Roman", "Liberation Serif", Tinos, serif`;
  return font ? `"${font}", ${fallback}` : fallback;
}

/** Word's single line spacing is a little more than 1em and depends on the font. */
const lineFactor = (font?: string) => (/calibri|cambria/i.test(font ?? "") ? 1.22 : 1.15);

function cssBorder(b: Border | null | undefined) {
  if (!b) return "none";
  const style = b.style === "double" ? "double" : /dash/i.test(b.style) ? "dashed" : /dot/i.test(b.style) ? "dotted" : "solid";
  return `${b.style === "double" ? Math.max(b.width, 2.25) : b.width}pt ${style} ${b.color}`;
}

export class WordDoc {
  private ids = new WeakMap<Element, string>();
  private nodes = new Map<string, Element>();
  private seq = 0;
  private labels = new Map<Element, string>();
  readonly body: Element;
  readonly section: Section;

  private constructor(
    private zip: JSZip,
    readonly xml: XMLDocument,
    private styles: Styles,
    private numbering: Numbering,
    private media: Map<string, string>,
  ) {
    this.body = kid(xml.documentElement, "body")!;
    const ps = Array.from(this.body.getElementsByTagNameNS(W, "sectPr"));
    const sect = kid(this.body, "sectPr") ?? ps[ps.length - 1];
    const sz = kid(sect, "pgSz");
    const mar = kid(sect, "pgMar");
    const n = (el: Element | undefined, a: string, d: number) => Math.abs(+(wattr(el, a) ?? d)) / 20;
    this.section = {
      width: n(sz, "w", 11906), height: n(sz, "h", 16838),
      top: n(mar, "top", 1134), bottom: n(mar, "bottom", 1134), left: n(mar, "left", 1701), right: n(mar, "right", 850),
    };
  }

  /** Opens a .docx. `documentXml` is a saved edit of word/document.xml to use instead of the file's own. */
  static async open(bytes: ArrayBuffer, documentXml?: string) {
    const zip = await JSZip.loadAsync(bytes);
    const read = async (path: string) => {
      const f = zip.file(path);
      return f ? new DOMParser().parseFromString(await f.async("string"), "application/xml") : undefined;
    };
    const text = documentXml ?? (await zip.file(DOC_PATH)?.async("string"));
    if (!text) throw new Error("В файле нет текста документа (word/document.xml).");
    const xml = new DOMParser().parseFromString(text, "application/xml");
    if (xml.getElementsByTagName("parsererror").length || !kid(xml.documentElement, "body"))
      throw new Error("Не удалось прочитать документ.");
    const theme = new Theme(await read("word/theme/theme1.xml"));
    const styles = new Styles(await read("word/styles.xml"), theme);
    const numbering = new Numbering(await read("word/numbering.xml"), theme);
    const media = new Map<string, string>();
    const rels = await read("word/_rels/document.xml.rels");
    for (const rel of Array.from(rels?.getElementsByTagName("Relationship") ?? [])) {
      if (!/\/image$/.test(rel.getAttribute("Type") ?? "") || rel.getAttribute("TargetMode") === "External") continue;
      const target = (rel.getAttribute("Target") ?? "").replace(/^\//, "");
      const file = zip.file(target.startsWith("word/") ? target : "word/" + target);
      if (!file) continue;
      const ext = target.split(".").pop()!.toLowerCase();
      const mime = ext === "jpg" || ext === "jpeg" ? "image/jpeg" : ext === "svg" ? "image/svg+xml" : `image/${ext}`;
      media.set(rel.getAttribute("Id") ?? "", `data:${mime};base64,${await file.async("base64")}`);
    }
    return new WordDoc(zip, xml, styles, numbering, media);
  }

  serialize() {
    const s = new XMLSerializer().serializeToString(this.xml);
    return s.startsWith("<?xml") ? s : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n${s}`;
  }

  async toDocx(): Promise<Blob> {
    this.zip.file(DOC_PATH, this.serialize());
    return this.zip.generateAsync({ type: "blob", mimeType: DOCX_TYPE, compression: "DEFLATE" });
  }

  private id(el: Element) {
    let id = this.ids.get(el);
    if (!id) {
      id = String(++this.seq);
      this.ids.set(el, id);
      this.nodes.set(id, el);
    }
    return id;
  }
  private node(id: string | undefined) {
    return id ? this.nodes.get(id) : undefined;
  }

  // ---------- formatting ----------

  private paraStyle(p: Element) {
    return this.styles.chain(val(kid(kid(p, "pPr"), "pStyle")) ?? this.styles.defaultPara);
  }
  private paraProps(p: Element): ParaProps {
    const pPr = kid(p, "pPr");
    const st = this.paraStyle(p);
    const own = readPPr(pPr);
    const listed = merge<ParaProps>(st.p, own);
    const lv = this.numbering.level(listed.numId, listed.ilvl);
    return merge<ParaProps>(this.styles.defP, st.p, lv?.p, own);
  }
  private markProps(p: Element): RunProps {
    return merge<RunProps>(this.styles.defR, this.paraStyle(p).r, readRPr(kid(kid(p, "pPr"), "rPr"), this.styles.theme));
  }
  private runProps(r: Element, p: Element): RunProps {
    const rPr = kid(r, "rPr");
    return merge<RunProps>(this.styles.defR, this.paraStyle(p).r, this.styles.chain(val(kid(rPr, "rStyle"))).r, readRPr(rPr, this.styles.theme));
  }

  // ---------- reading runs ----------

  /** Runs of a paragraph in order, looking inside hyperlinks, fields and content controls; deleted text is skipped. */
  private runs(p: Element): Element[] {
    const out: Element[] = [];
    const walk = (el: Element) => {
      for (const c of kids(el)) {
        if (c.localName === "r") out.push(c);
        else if (RUN_WRAPPERS.has(c.localName)) walk(c);
        else if (c.localName === "sdt") { const content = kid(c, "sdtContent"); if (content) walk(content); }
      }
    };
    walk(p);
    return out;
  }

  private isEditable(r: Element, p: Element) {
    return kids(r).every((c) => TEXTY.has(c.localName) && !(c.localName === "br" && isPageBreak(c))) && !this.runProps(r, p).vanish;
  }

  private runText(r: Element) {
    let s = "";
    for (const c of kids(r))
      switch (c.localName) {
        case "t": s += c.textContent ?? ""; break;
        case "tab": s += "\t"; break;
        case "br": case "cr": s += "\n"; break;
        case "noBreakHyphen": s += "‑"; break;
        case "softHyphen": s += "­"; break;
      }
    return s;
  }

  private setRunText(r: Element, text: string) {
    if (this.runText(r) === text) return;
    for (const c of kids(r)) if (c.localName !== "rPr") c.remove();
    const add = (name: string) => r.appendChild(this.xml.createElementNS(W, "w:" + name));
    for (const part of text.split(/([\n\t‑­])/)) {
      if (part === "") continue;
      if (part === "\n") add("br");
      else if (part === "\t") add("tab");
      else if (part === "‑") add("noBreakHyphen");
      else if (part === "­") add("softHyphen");
      else {
        const t = add("t");
        t.setAttributeNS(XML_NS, "xml:space", "preserve");
        t.textContent = part;
      }
    }
  }

  // ---------- rendering ----------

  /** The whole document as a list of blocks (paragraphs and tables). */
  render(editable: boolean): HTMLElement[] {
    this.numbering.reset();
    this.labels.clear();
    return this.blocks(this.body, editable);
  }

  private blocks(container: Element, editable: boolean): HTMLElement[] {
    const out: HTMLElement[] = [];
    for (const c of kids(container)) {
      if (c.localName === "p") out.push(this.paragraph(c, editable));
      else if (c.localName === "tbl") out.push(this.table(c, editable));
      else if (c.localName === "sdt") out.push(...this.blocks(kid(c, "sdtContent") ?? c, editable));
      else if (c.localName === "customXml") out.push(...this.blocks(c, editable));
    }
    return out;
  }

  /** Redraws one paragraph from the XML (after editing it), keeping its list label. */
  refresh(div: HTMLElement) {
    const p = this.node(div.dataset.p);
    if (p) div.replaceWith(this.paragraph(p, true, false));
  }

  private paragraph(p: Element, editable: boolean, count = true): HTMLElement {
    const pp = this.paraProps(p);
    const mark = this.markProps(p);
    const div = document.createElement("div");
    div.className = "p";
    div.dataset.p = this.id(p);
    const s = div.style;
    const size = mark.size ?? 10;
    s.fontSize = `${size}pt`;
    s.fontFamily = fontStack(mark.font);
    const line = pp.line ?? 240;
    const auto = (line / 240) * lineFactor(mark.font);
    if (pp.lineRule === "exact") s.lineHeight = pt(line);
    else if (pp.lineRule === "atLeast") s.lineHeight = line / 20 > size * auto ? pt(line) : String(auto);
    else s.lineHeight = String(auto);
    s.minHeight = pp.lineRule === "exact" ? pt(line) : `${auto}em`;
    if (pp.before) s.marginTop = pt(pp.before);
    if (pp.after) s.marginBottom = pt(pp.after);
    if (pp.left) s.paddingLeft = pt(pp.left);
    if (pp.right) s.paddingRight = pt(pp.right);
    if (pp.hanging) s.textIndent = pt(-pp.hanging);
    else if (pp.firstLine) s.textIndent = pt(pp.firstLine);
    s.textAlign = { both: "justify", distribute: "justify", center: "center", right: "right", end: "right" }[pp.jc ?? ""] ?? "left";
    // Neighbours with the same borders share one box, as in Word: no line between them unless "between" is set.
    const b = { ...pp.borders };
    const sig = JSON.stringify(b);
    const same = (el: Element | null) => el?.namespaceURI === W && el.localName === "p" && JSON.stringify(this.paraProps(el).borders ?? {}) === sig;
    if (sig !== "{}") {
      if (same(p.previousElementSibling)) b.top = b.between;
      if (same(p.nextElementSibling)) b.bottom = undefined;
    }
    for (const side of ["top", "bottom", "left", "right"] as const) if (b[side]) s.setProperty(`border-${side}`, cssBorder(b[side]));
    if (pp.shd) s.background = pp.shd;
    if (pp.pageBreakBefore) div.dataset.break = "before";
    if (editable) {
      div.contentEditable = "true";
      div.spellcheck = false;
    }

    // List label
    if (pp.numId && pp.numId !== "0") {
      if (count) this.labels.set(p, this.numbering.next(pp.numId, pp.ilvl ?? 0) ?? "");
      const label = this.labels.get(p);
      if (label) {
        const span = document.createElement("span");
        span.dataset.skip = "";
        span.contentEditable = "false";
        span.textContent = label + "\t";
        this.styleRun(span, merge<RunProps>(mark, this.numbering.level(pp.numId, pp.ilvl)?.r));
        div.append(span);
      }
    }

    let hasText = false;
    for (const r of this.runs(p)) {
      const props = this.runProps(r, p);
      const span = document.createElement("span");
      this.styleRun(span, props);
      if (this.isEditable(r, p)) {
        span.dataset.r = this.id(r);
        span.textContent = this.runText(r);
        hasText ||= span.textContent !== "";
        div.append(span);
        continue;
      }
      if (props.vanish) continue;
      // Locked run: show what it contains, never edit it.
      span.dataset.l = this.id(r);
      span.contentEditable = "false";
      for (const c of kids(r)) {
        const name = c.localName;
        if (name === "t") span.append(c.textContent ?? "");
        else if (name === "tab") span.append("\t");
        else if (name === "br" && isPageBreak(c)) div.dataset.break = hasText ? "after" : "before";
        else if (name === "br" || name === "cr") span.append("\n");
        else if (name === "sym") {
          // Symbol-font characters (F0xx) have no Unicode meaning; show them as a bullet.
          const code = parseInt(wattr(c, "char") ?? "20", 16) || 32;
          span.append(String.fromCharCode(code >= 0xf000 ? 0x2022 : code));
        }
        else if (name === "drawing" || name === "pict" || name === "object") this.picture(c, span);
      }
      if (span.childNodes.length) {
        hasText = true;
        div.append(span);
      }
    }
    return div;
  }

  private styleRun(span: HTMLElement, rp: RunProps) {
    const s = span.style;
    s.fontFamily = fontStack(rp.font);
    s.fontSize = `${rp.size ?? 10}pt`;
    if (rp.b) s.fontWeight = "bold";
    if (rp.i) s.fontStyle = "italic";
    const deco = [rp.u && "underline", rp.strike && "line-through"].filter(Boolean).join(" ");
    if (deco) s.textDecoration = deco;
    if (rp.caps) s.textTransform = "uppercase";
    if (rp.smallCaps) s.fontVariant = "small-caps";
    if (rp.color) s.color = rp.color;
    if (rp.highlight || rp.shd) s.background = (rp.highlight ?? rp.shd)!;
    if (rp.vert === "superscript" || rp.vert === "subscript") {
      s.verticalAlign = rp.vert === "superscript" ? "super" : "sub";
      s.fontSize = `${(rp.size ?? 10) * 0.65}pt`;
    }
    if (rp.spacing) s.letterSpacing = pt(rp.spacing);
    if (rp.bdr) s.border = cssBorder(rp.bdr);
  }

  private picture(el: Element, into: HTMLElement) {
    const blip = el.getElementsByTagNameNS(A, "blip")[0];
    const vml = el.getElementsByTagNameNS(V, "imagedata")[0];
    const src = this.media.get(blip?.getAttributeNS(R, "embed") ?? vml?.getAttributeNS(R, "id") ?? "");
    if (!src) return;
    const img = document.createElement("img");
    img.src = src;
    const ext = el.getElementsByTagNameNS(WP, "extent")[0];
    if (ext) {
      img.style.width = `${+(ext.getAttribute("cx") ?? 0) / 12700}pt`;
      img.style.height = `${+(ext.getAttribute("cy") ?? 0) / 12700}pt`;
    } else {
      const style = el.getElementsByTagNameNS(V, "shape")[0]?.getAttribute("style") ?? "";
      img.style.width = /width:([\d.]+\w*)/.exec(style)?.[1] ?? "";
      img.style.height = /height:([\d.]+\w*)/.exec(style)?.[1] ?? "";
    }
    into.append(img);
  }

  private table(tbl: Element, editable: boolean): HTMLElement {
    const tblPr = kid(tbl, "tblPr");
    const prs = [...this.styles.tablePrs(val(kid(tblPr, "tblStyle")) ?? this.styles.defaultTable), ...(tblPr ? [tblPr] : [])];
    let borders: Borders = {};
    const mar = { top: 0, bottom: 0, left: 108, right: 108 };
    for (const pr of prs) {
      borders = { ...borders, ...readBorders(kid(pr, "tblBorders")) };
      readMargins(kid(pr, "tblCellMar"), mar);
    }
    const grid = kids(kid(tbl, "tblGrid"), "gridCol").map((g) => +(wattr(g, "w") ?? 0));
    const rows = kids(tbl).flatMap((c) => (c.localName === "tr" ? [c] : c.localName === "sdt" ? kids(kid(c, "sdtContent"), "tr") : []));

    interface Cell { tc: Element; col: number; span: number; vm?: "restart" | "cont"; rowspan: number }
    const layout = rows.map((tr) => {
      let col = +(val(kid(kid(tr, "trPr"), "gridBefore")) ?? 0);
      const tcs = kids(tr).flatMap((c) => (c.localName === "tc" ? [c] : c.localName === "sdt" ? kids(kid(c, "sdtContent"), "tc") : []));
      return tcs.map((tc): Cell => {
        const tcPr = kid(tc, "tcPr");
        const span = +(val(kid(tcPr, "gridSpan")) ?? 1);
        const vmEl = kid(tcPr, "vMerge");
        const cell: Cell = { tc, col, span, vm: vmEl ? (val(vmEl) === "restart" ? "restart" : "cont") : undefined, rowspan: 1 };
        col += span;
        return cell;
      });
    });
    layout.forEach((cells, r) => {
      for (const c of cells) {
        if (c.vm !== "restart") continue;
        for (let k = r + 1; k < layout.length; k++) {
          if (!layout[k].some((d) => d.col === c.col && d.vm === "cont")) break;
          c.rowspan++;
        }
      }
    });

    const table = document.createElement("table");
    table.className = "tbl";
    const total = grid.reduce((a, b) => a + b, 0);
    table.style.width = pt(total);
    const jc = val(kid(tblPr, "jc"));
    if (jc === "center") table.style.alignSelf = "center";
    else if (jc === "right" || jc === "end") table.style.alignSelf = "flex-end";
    const ind = wattr(kid(tblPr, "tblInd"), "w");
    if (ind && +ind) table.style.marginLeft = pt(+ind);
    const colgroup = document.createElement("colgroup");
    for (const w of grid) {
      const col = document.createElement("col");
      col.style.width = pt(w);
      colgroup.append(col);
    }
    table.append(colgroup);
    const tbody = document.createElement("tbody");
    table.append(tbody);

    rows.forEach((trEl, r) => {
      const tr = document.createElement("tr");
      const trPr = kid(trEl, "trPr");
      const h = kid(trPr, "trHeight");
      const height = h ? +(val(h) ?? 0) : 0;
      const exact = wattr(h, "hRule") === "exact";
      if (height) tr.style.height = pt(height);
      const before = +(val(kid(trPr, "gridBefore")) ?? 0);
      if (before) {
        const td = document.createElement("td");
        td.colSpan = before;
        td.className = "gap";
        tr.append(td);
      }
      for (const c of layout[r]) {
        if (c.vm === "cont") continue;
        const tcPr = kid(c.tc, "tcPr");
        const td = document.createElement("td");
        if (c.span > 1) td.colSpan = c.span;
        if (c.rowspan > 1) td.rowSpan = c.rowspan;
        const own = readBorders(kid(tcPr, "tcBorders")) ?? {};
        const lastRow = r + c.rowspan - 1 >= rows.length - 1;
        const side = (name: "top" | "bottom" | "left" | "right", outer: boolean, inner: "insideH" | "insideV") =>
          cssBorder(own[name] !== undefined ? own[name] : outer ? borders[name] : borders[inner]);
        td.style.borderTop = side("top", r === 0, "insideH");
        td.style.borderBottom = side("bottom", lastRow, "insideH");
        td.style.borderLeft = side("left", c.col === 0, "insideV");
        td.style.borderRight = side("right", c.col + c.span >= grid.length, "insideV");
        const m = readMargins(kid(tcPr, "tcMar"), { ...mar });
        td.style.padding = `${pt(m.top)} ${pt(m.right)} ${pt(m.bottom)} ${pt(m.left)}`;
        td.style.verticalAlign = { center: "middle", bottom: "bottom" }[val(kid(tcPr, "vAlign")) ?? ""] ?? "top";
        const shd = wattr(kid(tcPr, "shd"), "fill");
        if (shd && /^[0-9a-f]{6}$/i.test(shd)) td.style.background = "#" + shd;
        const inner = document.createElement("div");
        inner.className = "cell";
        inner.append(...this.blocks(c.tc, editable));
        if (exact && height) {
          inner.style.maxHeight = pt(height);
          inner.style.overflow = "hidden";
        }
        const dir = val(kid(tcPr, "textDirection"));
        if (dir === "btLr" || dir === "tbRl" || dir === "tbRlV") {
          // Rotated text: turn the cell's content a quarter turn inside the row's height.
          const width = grid.slice(c.col, c.col + c.span).reduce((a, b) => a + b, 0) - m.left - m.right;
          const rowsHeight = rows.slice(r, r + c.rowspan).reduce((a, tr) => a + +(val(kid(kid(tr, "trPr"), "trHeight")) ?? 0), 0);
          if (rowsHeight) {
            td.style.position = "relative";
            Object.assign(inner.style, {
              position: "absolute", left: "50%", top: "50%", width: pt(rowsHeight - m.top - m.bottom), maxHeight: pt(width),
              transform: `translate(-50%, -50%) rotate(${dir === "btLr" ? -90 : 90}deg)`,
            });
          } else {
            inner.style.writingMode = "vertical-rl";
            if (dir === "btLr") inner.style.transform = "rotate(180deg)";
          }
        }
        td.append(inner);
        tr.append(td);
      }
      tbody.append(tr);
    });
    return table;
  }

  // ---------- editing ----------

  /** Writes a paragraph's edited text back into its runs. Returns true when the XML changed. */
  sync(div: HTMLElement): boolean {
    const p = this.node(div.dataset.p);
    if (!p) return false;
    const was = new XMLSerializer().serializeToString(p);
    const editable = this.runs(p).filter((r) => this.isEditable(r, p));

    interface Seg { run?: Element; lock?: Element; text: string; from: Element; br?: boolean }
    const segs: Seg[] = [];
    const walk = (node: Node, owner: Element | undefined) => {
      for (const ch of Array.from(node.childNodes)) {
        if (ch.nodeType === Node.TEXT_NODE) {
          if (ch.nodeValue) segs.push({ run: owner, text: ch.nodeValue, from: ch.parentElement ?? div });
          continue;
        }
        if (!(ch instanceof HTMLElement)) continue;
        if (ch.dataset.skip !== undefined) continue;
        if (ch.dataset.l) {
          const lock = this.node(ch.dataset.l);
          if (lock) segs.push({ lock, text: "", from: ch });
          continue;
        }
        if (ch.tagName === "BR") {
          segs.push({ run: owner, text: "\n", from: node as Element, br: true });
          continue;
        }
        // Some browsers wrap a new line in its own <div>.
        if ((ch.tagName === "DIV" || ch.tagName === "P") && segs.some((s) => s.text)) segs.push({ run: owner, text: "\n", from: ch });
        const run = this.node(ch.dataset.r);
        walk(ch, run && editable.includes(run) ? run : owner);
      }
    };
    walk(div, undefined);
    // A <br> at the very end only holds the empty last line open.
    for (let i = segs.length - 1; i >= 0; i--) {
      if (segs[i].lock) continue;
      if (segs[i].br) segs.splice(i, 1);
      break;
    }

    // Text typed outside any run takes the formatting of the run it looks like.
    const inDom = new Set(segs.map((s) => s.run).filter(Boolean));
    let prev: Element | undefined;
    for (const s of segs) {
      if (s.lock) continue;
      if (!s.run && s.text) s.run = this.likeliestRun(s.from, div, p, editable, inDom, prev);
      prev = s.run ?? prev;
    }

    const used = new Set<Element>();
    let last: Element | undefined;
    const after = (a: Element, b: Element | undefined) => !b || !!(b.compareDocumentPosition(a) & Node.DOCUMENT_POSITION_FOLLOWING);
    for (const s of segs) {
      if (s.lock) {
        if (after(s.lock, last)) last = s.lock;
        continue;
      }
      if (!s.text) continue;
      let target = s.run;
      if (target && last === target && used.has(target)) {
        this.setRunText(target, this.runText(target) + s.text);
        continue;
      }
      if (!target || used.has(target) || !after(target, last)) {
        const fresh = this.xml.createElementNS(W, "w:r");
        const rPr = target ? kid(target, "rPr") : kid(kid(p, "pPr"), "rPr");
        if (rPr) {
          const copy = rPr.cloneNode(true) as Element;
          for (const c of kids(copy)) if (["ins", "del", "moveFrom", "moveTo", "rPrChange"].includes(c.localName)) c.remove();
          fresh.append(copy);
        }
        if (last) last.parentNode!.insertBefore(fresh, last.nextSibling);
        else {
          const first = kids(p).find((c) => c.localName !== "pPr");
          p.insertBefore(fresh, first ?? null);
        }
        target = fresh;
      }
      this.setRunText(target, s.text);
      used.add(target);
      last = target;
    }
    for (const r of editable) if (!used.has(r)) r.remove();
    return new XMLSerializer().serializeToString(p) !== was;
  }

  private likeliestRun(from: Element, div: HTMLElement, p: Element, runs: Element[], inDom: Set<Element | undefined>, prev?: Element) {
    if (!runs.length) return undefined;
    const cs = getComputedStyle(from);
    const bold = cs.fontWeight === "bold" || +cs.fontWeight >= 600;
    const italic = cs.fontStyle === "italic";
    let underline = false;
    for (let el: Element | null = from; el && el !== div.parentElement; el = el.parentElement) {
      const d = getComputedStyle(el);
      if ((d.textDecorationLine || d.textDecoration || "").includes("underline")) underline = true;
      if (el === div) break;
    }
    let best = runs[0];
    let bestScore = -1;
    for (const r of runs) {
      const rp = this.runProps(r, p);
      const score = 3 * +(!!rp.b === bold) + 3 * +(!!rp.i === italic) + 3 * +(!!rp.u === underline) + +!inDom.has(r) + 0.5 * +(r === prev);
      if (score > bestScore) [best, bestScore] = [r, score];
    }
    return best;
  }
}

function readMargins(el: Element | undefined, into: { top: number; bottom: number; left: number; right: number }) {
  for (const [side, alt] of [["top", "top"], ["bottom", "bottom"], ["left", "start"], ["right", "end"]] as const) {
    const m = kid(el, side) ?? kid(el, alt);
    const w = wattr(m, "w");
    if (m && w != null && (wattr(m, "type") ?? "dxa") === "dxa") into[side] = +w;
  }
  return into;
}
