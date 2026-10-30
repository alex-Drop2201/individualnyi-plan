// Reading Word formatting (styles.xml, theme, direct w:rPr / w:pPr) into plain objects.
// Units: sizes in points; spacing, indents and widths in twips (1/20 pt) as Word stores them.

export const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";

export const kids = (el: Element | null | undefined, name?: string) =>
  el ? Array.from(el.children).filter((c) => c.namespaceURI === W && (!name || c.localName === name)) : [];
export const kid = (el: Element | null | undefined, name: string) => kids(el, name)[0] as Element | undefined;
export const wattr = (el: Element | null | undefined, name: string) => el?.getAttributeNS(W, name) ?? null;
export const val = (el: Element | null | undefined) => wattr(el, "val");
const num = (s: string | null) => (s == null || s === "" || isNaN(+s) ? undefined : +s);

/** On/off properties (w:b, w:i, …): present without w:val means on. */
const toggle = (el: Element | undefined) => (el ? !["0", "false", "off", "none"].includes(val(el) ?? "") : undefined);

export interface RunProps {
  font?: string;
  size?: number;
  b?: boolean;
  i?: boolean;
  u?: boolean;
  strike?: boolean;
  caps?: boolean;
  smallCaps?: boolean;
  vanish?: boolean;
  color?: string;
  highlight?: string;
  shd?: string;
  vert?: string;
  spacing?: number;
  bdr?: Border | null;
}

export interface Border {
  style: string;
  width: number;
  color: string;
}
export type Borders = Partial<Record<"top" | "bottom" | "left" | "right" | "between" | "insideH" | "insideV", Border | null>>;

export interface ParaProps {
  jc?: string;
  before?: number;
  after?: number;
  line?: number;
  lineRule?: string;
  left?: number;
  right?: number;
  firstLine?: number;
  hanging?: number;
  borders?: Borders;
  shd?: string;
  numId?: string;
  ilvl?: number;
  pageBreakBefore?: boolean;
}

/** Later sources override earlier ones, field by field (undefined never overrides). */
export function merge<T extends object>(...parts: (Partial<T> | undefined)[]): T {
  const out: Record<string, unknown> = {};
  for (const p of parts)
    if (p)
      for (const [k, v] of Object.entries(p))
        if (v !== undefined) out[k] = k === "borders" ? { ...(out[k] as object), ...(v as object) } : v;
  return out as T;
}

const HIGHLIGHT: Record<string, string> = {
  yellow: "#ffff00", green: "#00ff00", cyan: "#00ffff", magenta: "#ff00ff", blue: "#0000ff", red: "#ff0000",
  darkBlue: "#000080", darkCyan: "#008080", darkGreen: "#008000", darkMagenta: "#800080", darkRed: "#800000",
  darkYellow: "#808000", darkGray: "#808080", lightGray: "#c0c0c0", black: "#000000",
};
const color = (s: string | null) => (s && /^[0-9a-f]{6}$/i.test(s) ? "#" + s : undefined);
const fill = (shd: Element | undefined) => (shd ? color(wattr(shd, "fill")) : undefined);

export function readBorder(el: Element | undefined): Border | null | undefined {
  if (!el) return undefined;
  const style = val(el) ?? "none";
  if (style === "none" || style === "nil") return null;
  return { style, width: Math.max(0.5, (num(wattr(el, "sz")) ?? 4) / 8), color: color(wattr(el, "color")) ?? "#000" };
}

export function readBorders(el: Element | undefined): Borders | undefined {
  if (!el) return undefined;
  const out: Borders = {};
  for (const side of ["top", "bottom", "left", "right", "between", "insideH", "insideV"] as const) {
    const alt = side === "left" ? "start" : side === "right" ? "end" : undefined;
    const b = readBorder(kid(el, side) ?? (alt ? kid(el, alt) : undefined));
    if (b !== undefined) out[side] = b;
  }
  return out;
}

export class Theme {
  minor?: string;
  major?: string;
  constructor(xml?: XMLDocument) {
    const font = (tag: string) => xml?.getElementsByTagNameNS(A, tag)[0]?.getElementsByTagNameNS(A, "latin")[0]?.getAttribute("typeface") || undefined;
    this.minor = font("minorFont");
    this.major = font("majorFont");
  }
}

export function readRPr(rPr: Element | undefined, theme: Theme): Partial<RunProps> | undefined {
  if (!rPr) return undefined;
  const f = kid(rPr, "rFonts");
  const themeFont = wattr(f, "asciiTheme") ?? wattr(f, "hAnsiTheme");
  const u = kid(rPr, "u");
  return {
    font: wattr(f, "ascii") ?? wattr(f, "hAnsi") ?? (themeFont ? (themeFont.startsWith("major") ? theme.major : theme.minor) : undefined),
    size: num(val(kid(rPr, "sz"))) !== undefined ? num(val(kid(rPr, "sz")))! / 2 : undefined,
    b: toggle(kid(rPr, "b")),
    i: toggle(kid(rPr, "i")),
    u: u ? val(u) !== "none" : undefined,
    strike: toggle(kid(rPr, "strike")) || toggle(kid(rPr, "dstrike")),
    caps: toggle(kid(rPr, "caps")),
    smallCaps: toggle(kid(rPr, "smallCaps")),
    vanish: toggle(kid(rPr, "vanish")),
    color: color(val(kid(rPr, "color"))),
    highlight: HIGHLIGHT[val(kid(rPr, "highlight")) ?? ""],
    shd: fill(kid(rPr, "shd")),
    vert: val(kid(rPr, "vertAlign")) ?? undefined,
    spacing: num(val(kid(rPr, "spacing"))),
    bdr: readBorder(kid(rPr, "bdr")),
  };
}

export function readPPr(pPr: Element | undefined): Partial<ParaProps> | undefined {
  if (!pPr) return undefined;
  const sp = kid(pPr, "spacing");
  const ind = kid(pPr, "ind");
  const numPr = kid(pPr, "numPr");
  return {
    jc: val(kid(pPr, "jc")) ?? undefined,
    before: num(wattr(sp, "before")),
    after: num(wattr(sp, "after")),
    line: num(wattr(sp, "line")),
    lineRule: wattr(sp, "lineRule") ?? (wattr(sp, "line") ? "auto" : undefined),
    left: num(wattr(ind, "left") ?? wattr(ind, "start")),
    right: num(wattr(ind, "right") ?? wattr(ind, "end")),
    firstLine: num(wattr(ind, "firstLine")),
    hanging: num(wattr(ind, "hanging")),
    borders: readBorders(kid(pPr, "pBdr")),
    shd: fill(kid(pPr, "shd")),
    numId: val(kid(numPr, "numId")) ?? undefined,
    ilvl: num(val(kid(numPr, "ilvl"))),
    pageBreakBefore: toggle(kid(pPr, "pageBreakBefore")),
  };
}

export class Styles {
  private byId = new Map<string, Element>();
  private cache = new Map<string, { p: ParaProps; r: RunProps }>();
  defR: RunProps;
  defP: ParaProps;
  defaultPara?: string;
  defaultTable?: string;

  constructor(xml: XMLDocument | undefined, readonly theme: Theme) {
    const root = xml?.documentElement;
    for (const s of kids(root, "style")) {
      const id = wattr(s, "styleId");
      if (!id) continue;
      this.byId.set(id, s);
      if (["1", "true"].includes(wattr(s, "default") ?? "")) {
        if (wattr(s, "type") === "paragraph") this.defaultPara = id;
        if (wattr(s, "type") === "table") this.defaultTable = id;
      }
    }
    const dd = kid(root, "docDefaults");
    this.defR = merge<RunProps>({ size: 10 }, readRPr(kid(kid(dd, "rPrDefault"), "rPr"), theme));
    this.defP = merge<ParaProps>(readPPr(kid(kid(dd, "pPrDefault"), "pPr")));
  }

  /** Paragraph or character style with everything it inherits through w:basedOn. */
  chain(id: string | null | undefined, seen = new Set<string>()): { p: ParaProps; r: RunProps } {
    if (!id || seen.has(id) || !this.byId.has(id)) return { p: {}, r: {} };
    const hit = this.cache.get(id);
    if (hit) return hit;
    seen.add(id);
    const s = this.byId.get(id)!;
    const base = this.chain(val(kid(s, "basedOn")), seen);
    const out = { p: merge<ParaProps>(base.p, readPPr(kid(s, "pPr"))), r: merge<RunProps>(base.r, readRPr(kid(s, "rPr"), this.theme)) };
    this.cache.set(id, out);
    return out;
  }

  /** w:tblPr elements of a table style and its bases, base first. */
  tablePrs(id: string | null | undefined, seen = new Set<string>()): Element[] {
    if (!id || seen.has(id) || !this.byId.has(id)) return [];
    seen.add(id);
    const s = this.byId.get(id)!;
    const own = kid(s, "tblPr");
    return [...this.tablePrs(val(kid(s, "basedOn")), seen), ...(own ? [own] : [])];
  }
}
