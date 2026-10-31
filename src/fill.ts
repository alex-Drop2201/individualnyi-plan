// Fill-in lines: "______" or underlined spaces, the lines a form leaves for handwriting.
// Typing over one must not push the line along, as handwriting would not: typed text takes
// the place of as much of the line as it covers, and deleting text gives the line back. Text
// typed onto "____" is underlined, so the line stays under the letters.
// The renderer marks such runs with data-blank ("_" or " "); text typed onto "____" goes into a
// span with data-like (the run to copy formatting from), data-u (add an underline) and, on a
// line that is nothing but "____", data-font (the run whose typeface the text takes: the label
// before the line, since such lines often carry a font of their own).

const ctx = document.createElement("canvas").getContext("2d")!;

function width(text: string, el: Element) {
  const cs = getComputedStyle(el);
  ctx.font = `${cs.fontStyle} ${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`;
  return ctx.measureText(text).width;
}

/** Total width of a paragraph's text on one line (wrapping ignored). */
export function lineWidth(div: HTMLElement) {
  let w = 0;
  const walker = document.createTreeWalker(div, NodeFilter.SHOW_TEXT);
  for (let n = walker.nextNode(); n; n = walker.nextNode()) if (n.nodeValue && n.parentElement) w += width(n.nodeValue, n.parentElement);
  return w;
}

export const hasLine = (div: HTMLElement) => !!div.querySelector("[data-blank]");

function caret() {
  const sel = getSelection();
  return sel?.anchorNode ? { node: sel.anchorNode, offset: sel.anchorOffset } : null;
}

function setCaret(node: Node, offset: number) {
  const sel = getSelection();
  if (!sel) return;
  const r = document.createRange();
  r.setStart(node, offset);
  r.collapse(true);
  sel.removeAllRanges();
  sel.addRange(r);
}

/**
 * Typing with the caret on "____": puts the text in a new underlined piece at the caret
 * instead of between the underscores. Returns false (and does nothing) anywhere else.
 */
export function typeOnLine(text: string): boolean {
  const sel = getSelection();
  if (!sel?.isCollapsed || !(sel.anchorNode instanceof Text)) return false;
  const node = sel.anchorNode;
  const span = node.parentElement;
  if (!span || span.dataset.blank !== "_" || !span.dataset.r) return false;
  const k = sel.anchorOffset;
  if (node.data[k] !== "_" && node.data[k - 1] !== "_") return false;

  const typed = document.createElement("span");
  typed.style.cssText = span.style.cssText;
  const label = /[^_\s]/.test(span.textContent ?? "") ? null : labelBefore(span);
  if (label) {
    // The label's typeface, at the line's own size, so the lines keep their height.
    typed.style.fontFamily = label.style.fontFamily;
    typed.dataset.font = label.dataset.r;
  }
  typed.style.textDecoration = `${span.style.textDecoration.includes("line-through") ? "line-through " : ""}underline`;
  typed.dataset.like = span.dataset.r;
  typed.dataset.u = "";
  typed.textContent = text;
  if (k === 0) span.before(typed);
  else if (k === node.data.length) span.after(typed);
  else {
    const rest = span.cloneNode(false) as HTMLElement;
    rest.textContent = node.data.slice(k);
    node.data = node.data.slice(0, k);
    span.after(typed, rest);
  }
  setCaret(typed.firstChild!, text.length);
  return true;
}

/** The last piece of real text before a line: in its own paragraph, or else the paragraph above. */
function labelBefore(span: HTMLElement): HTMLElement | null {
  const textOf = (el: Element | null) =>
    el ? Array.from(el.querySelectorAll<HTMLElement>("[data-r]:not([data-blank])")).filter((s) => /\S/.test(s.textContent ?? "")) : [];
  const div = span.closest(".p");
  const own = textOf(div).filter((s) => !!(s.compareDocumentPosition(span) & Node.DOCUMENT_POSITION_FOLLOWING));
  if (own.length) return own[own.length - 1];
  const above = textOf(div?.previousElementSibling?.classList.contains("p") ? div.previousElementSibling : null);
  return above[above.length - 1] ?? null;
}

/** How a paragraph's text lies: its width on one line, and where its wrapped lines end. */
export interface Shape { width: number; lines: number; right: number }

export function shapeOf(div: HTMLElement): Shape {
  const r = document.createRange();
  r.selectNodeContents(div);
  const rects = Array.from(r.getClientRects()).filter((x) => x.width > 0);
  const tops: number[] = [];
  for (const x of rects) if (!tops.some((t) => Math.abs(t - x.top) < 3)) tops.push(x.top);
  const lastTop = Math.max(...tops);
  const right = Math.max(...rects.filter((x) => Math.abs(x.top - lastTop) < 3).map((x) => x.right));
  return { width: lineWidth(div), lines: tops.length, right: rects.length ? right : 0 };
}

/**
 * After an edit, gives the line nearest the caret back or takes it away, so the paragraph
 * keeps its shape from before the edit: same number of lines, ending at the same place.
 */
export function keepLineLength(div: HTMLElement, before: Shape) {
  const lines = Array.from(div.querySelectorAll<HTMLElement>("[data-blank]"));
  const at = caret();
  if (!lines.length || !at) return;
  // The first line at or after the caret, otherwise the last one before it.
  let line = lines.find((l) => l.contains(at.node) || !!(at.node.compareDocumentPosition(l) & Node.DOCUMENT_POSITION_FOLLOWING));
  const beforeCaret = !line;
  line ??= lines[lines.length - 1];
  const ch = line.dataset.blank!;
  const charWidth = width(ch, line);
  if (!charWidth) return;
  let text = Array.from(line.childNodes).find((c): c is Text => c instanceof Text);
  if (!text) text = line.appendChild(document.createTextNode(""));
  const node = text;
  const inside = node === at.node;
  let caretAt = inside ? at.offset : -1;

  /** Adds (n > 0) or removes (n < 0) line characters next to the caret. Returns how many changed. */
  const change = (n: number) => {
    const data = node.data;
    if (n > 0) {
      const where = inside ? caretAt : beforeCaret ? data.lastIndexOf(ch) + 1 || data.length : Math.max(0, data.indexOf(ch));
      node.data = data.slice(0, where) + ch.repeat(n) + data.slice(where);
      return n;
    }
    // Take line characters after the caret, or else before it.
    const after = beforeCaret ? -1 : data.indexOf(ch, inside ? caretAt : 0);
    let start: number, end: number;
    if (after >= 0) {
      start = end = after;
      while (end < data.length && data[end] === ch && end - start < -n) end++;
    } else {
      const last = data.lastIndexOf(ch, inside ? caretAt - 1 : data.length - 1);
      if (last < 0) return 0;
      start = end = last + 1;
      while (start > 0 && data[start - 1] === ch && end - start < -n) start--;
      if (inside && caretAt >= end) caretAt -= end - start;
    }
    node.data = data.slice(0, start) + data.slice(end);
    return end - start;
  };

  // First by width, then by how the lines actually wrap (a word that does not fit leaves a gap).
  const n = Math.round((before.width - lineWidth(div)) / charWidth);
  if (n) change(n);
  const scale = div.getBoundingClientRect().width / (div.offsetWidth || 1);
  const tolerance = charWidth * scale * 0.6;
  for (let i = 0; i < 80; i++) {
    const now = shapeOf(div);
    if (now.lines > before.lines || (now.lines === before.lines && now.right > before.right + tolerance)) {
      if (!change(-1)) break;
    } else if (now.lines < before.lines || now.right < before.right - tolerance) {
      change(1);
      const next = shapeOf(div);
      if (next.lines > before.lines || (next.lines === before.lines && next.right > before.right + tolerance)) {
        change(-1);
        break;
      }
    } else break;
  }
  if (inside) setCaret(node, Math.max(0, Math.min(caretAt, node.data.length)));
}
