// List numbers ("1.", "а)", "•") from numbering.xml. Word stores only the rule, not the text,
// so the label is computed here in document order, for display and for the PDF.
import { kid, kids, readPPr, readRPr, val, wattr, type ParaProps, type RunProps, type Theme } from "./props";

interface Level {
  start: number;
  fmt: string;
  text: string;
  p?: Partial<ParaProps>;
  r?: Partial<RunProps>;
}

const ROMAN: [number, string][] = [[1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"], [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"]];
const roman = (n: number) => ROMAN.reduce((s, [v, r]) => { while (n >= v) { s += r; n -= v; } return s; }, "");
const letter = (n: number, abc: string) => abc[(n - 1) % abc.length].repeat(Math.floor((n - 1) / abc.length) + 1);
const LATIN = "abcdefghijklmnopqrstuvwxyz";
const CYRILLIC = "абвгдежзиклмнопрстуфхцчшщэюя";

function format(n: number, fmt: string) {
  switch (fmt) {
    case "lowerLetter": return letter(n, LATIN);
    case "upperLetter": return letter(n, LATIN).toUpperCase();
    case "lowerRoman": return roman(n);
    case "upperRoman": return roman(n).toUpperCase();
    case "russianLower": return letter(n, CYRILLIC);
    case "russianUpper": return letter(n, CYRILLIC).toUpperCase();
    case "decimalZero": return n < 10 ? "0" + n : String(n);
    case "none": return "";
    default: return String(n);
  }
}

export class Numbering {
  private levels = new Map<string, Level[]>();
  private counters = new Map<string, number[]>();

  constructor(xml: XMLDocument | undefined, theme: Theme) {
    const root = xml?.documentElement;
    const abstract = new Map<string, Element>();
    for (const a of kids(root, "abstractNum")) abstract.set(wattr(a, "abstractNumId") ?? "", a);
    for (const n of kids(root, "num")) {
      const a = abstract.get(val(kid(n, "abstractNumId")) ?? "");
      if (!a) continue;
      const levels: Level[] = [];
      for (const l of kids(a, "lvl")) {
        const i = +(wattr(l, "ilvl") ?? 0);
        levels[i] = {
          start: +(val(kid(l, "start")) ?? 1),
          fmt: val(kid(l, "numFmt")) ?? "decimal",
          text: val(kid(l, "lvlText")) ?? "",
          p: readPPr(kid(l, "pPr")),
          r: readRPr(kid(l, "rPr"), theme),
        };
      }
      for (const o of kids(n, "lvlOverride")) {
        const lv = levels[+(wattr(o, "ilvl") ?? 0)];
        const start = val(kid(o, "startOverride"));
        if (lv && start) lv.start = +start;
      }
      this.levels.set(wattr(n, "numId") ?? "", levels);
    }
  }

  /** Call before each full render: labels count from the start of the document. */
  reset() {
    this.counters.clear();
  }

  level(numId: string | undefined, ilvl = 0) {
    return numId ? this.levels.get(numId)?.[ilvl] : undefined;
  }

  /** The label for the next paragraph of this list, advancing its counter. */
  next(numId: string, ilvl = 0): string | undefined {
    const levels = this.levels.get(numId);
    const lv = levels?.[ilvl];
    if (!levels || !lv) return undefined;
    const c = this.counters.get(numId) ?? [];
    for (let i = 0; i < ilvl; i++) c[i] ??= levels[i]?.start ?? 1;
    c[ilvl] = c[ilvl] === undefined ? lv.start : c[ilvl] + 1;
    c.length = ilvl + 1;
    this.counters.set(numId, c);
    if (lv.fmt === "bullet") return lv.text === "" ? "" : "•";
    return lv.text.replace(/%(\d)/g, (_, d) => format(c[+d - 1] ?? levels[+d - 1]?.start ?? 1, levels[+d - 1]?.fmt ?? "decimal"));
  }
}
