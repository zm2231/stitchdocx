// Numbering labels: what Word renders in front of a numbered paragraph.
// Counters are per numbering instance (w:numId), levels restart deeper levels,
// w:lvlOverride/w:startOverride apply per instance, numPr may be inherited
// from a paragraph style chain, and %n in w:lvlText reads level n-1. A fresh
// counter set is created for every projection pass — labels never depend on
// earlier calls.
import { decodeAttrValue, kid, kids, parseXml, type XElement } from "@stitchdocx/package-io";

interface LevelDef {
  fmt: string;
  text: string;
  start: number;
}

interface NumDef {
  abstract: string;
  /** per-ilvl start override (w:lvlOverride/w:startOverride) */
  overrides: Map<string, number>;
}

interface StyleDef {
  basedOn?: string;
  numId?: string;
  ilvl?: string;
}

const attrOf = (el: XElement, name: string): string | null => {
  const a = el.attrs.find((x) => x.name === name);
  return a ? decodeAttrValue(a.raw) : null;
};

export class NumberingModel {
  constructor(
    private abstracts: Map<string, Map<number, LevelDef>>,
    private nums: Map<string, NumDef>,
    private styles: Map<string, StyleDef>,
  ) {}

  /** the numbering reference a paragraph renders, via pPr or its style chain */
  private numPrOf(p: XElement): { numId: string; ilvl: string } | null {
    const pPr = kid(p, "w:pPr");
    const numPr = pPr ? kid(pPr, "w:numPr") : undefined;
    if (numPr) {
      const numId = attrOf(kid(numPr, "w:numId") ?? (el0 as XElement), "w:val");
      if (numId === null) return null;
      if (numId === "0") return null;
      const ilvl = attrOf(kid(numPr, "w:ilvl") ?? (el0 as XElement), "w:val") ?? "0";
      return { numId, ilvl };
    }
    // inherited from a paragraph style (walking the basedOn chain)
    const pStyle = pPr ? kid(pPr, "w:pStyle") : undefined;
    const styleId = pStyle ? attrOf(pStyle, "w:val") : null;
    if (styleId === null) return null;
    const seen = new Set<string>();
    let cur: string | undefined = styleId;
    while (cur && !seen.has(cur)) {
      seen.add(cur);
      const def = this.styles.get(cur);
      if (!def) return null;
      if (def.numId !== undefined) {
        if (def.numId === "0") return null;
        return { numId: def.numId, ilvl: def.ilvl ?? "0" };
      }
      cur = def.basedOn;
    }
    return null;
  }

  /** true when the paragraph renders a numbering label (numbering from pPr or style) */
  numberedFor(p: XElement): boolean {
    return this.numPrOf(p) !== null;
  }

  /**
   * The label for `p`, advancing the counters in `counters` (one entry per
   * numId; fresh per projection pass).
   */
  labelFor(p: XElement, counters: Map<string, (number | undefined)[]>): string | undefined {
    const ref = this.numPrOf(p);
    if (!ref) return undefined;
    const num = this.nums.get(ref.numId);
    if (!num) return undefined;
    const levels = this.abstracts.get(num.abstract);
    const level = parseInt(ref.ilvl, 10) || 0;
    const def = levels?.get(level);
    if (!def) return undefined;
    const arr = counters.get(ref.numId) ?? [];
    const start = num.overrides.get(ref.ilvl) ?? def.start;
    arr[level] = (arr[level] ?? start - 1) + 1;
    // a step at this level restarts every deeper level
    for (let i = level + 1; i < arr.length; i++) arr[i] = undefined;
    counters.set(ref.numId, arr);
    if (def.fmt === "none") return undefined;
    if (def.fmt === "bullet") return def.text;
    const at = (i: number): number => {
      const v = arr[i];
      if (v !== undefined) return v;
      const lvl = levels?.get(i);
      const s = lvl ? num.overrides.get(String(i)) ?? lvl.start : 1;
      return s;
    };
    return def.text.replace(/%(\d)/g, (_, d: string) => formatNumber(at(parseInt(d, 10) - 1), def.fmt));
  }
}

const el0: XElement = { kind: "el", name: "x", attrs: [], children: [], parent: null, rawOpen: "x", rawClose: "x", selfClosing: true, attrsDirty: false };

function formatNumber(n: number, fmt: string): string {
  switch (fmt) {
    case "decimal":
    case "decimalEnclosedCircle":
      return String(n);
    case "decimalZero":
      return String(n).padStart(2, "0");
    case "lowerLetter":
      return letters(n).toLowerCase();
    case "upperLetter":
      return letters(n).toUpperCase();
    case "lowerRoman":
      return roman(n).toLowerCase();
    case "upperRoman":
      return roman(n);
    case "ordinal":
      return String(n) + ordinalSuffix(n);
    default:
      return String(n);
  }
}

/** 1..26 → a..z, 27.. → aa.. (Excel-style) */
function letters(n: number): string {
  let out = "";
  let v = n;
  while (v > 0) {
    const rem = (v - 1) % 26;
    out = String.fromCharCode(97 + rem) + out;
    v = Math.floor((v - 1) / 26);
  }
  return out || "a";
}

function roman(n: number): string {
  if (n <= 0) return String(n);
  const table: Array<[number, string]> = [
    [1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"], [90, "XC"],
    [50, "L"], [40, "XL"], [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"],
  ];
  let out = "";
  let v = n;
  for (const [val, sym] of table) {
    while (v >= val) {
      out += sym;
      v -= val;
    }
  }
  return out;
}

function ordinalSuffix(n: number): string {
  const mod100 = n % 100;
  if (mod100 >= 11 && mod100 <= 13) return "th";
  switch (n % 10) {
    case 1: return "st";
    case 2: return "nd";
    case 3: return "rd";
    default: return "th";
  }
}

export const emptyNumbering: NumberingModel = new NumberingModel(new Map(), new Map(), new Map());

/** parse word/numbering.xml plus word/styles.xml into a label model. Missing
 *  parts yield the empty model (no labels). */
export function parseNumberingModel(numberingXml: string | undefined, stylesXml: string | undefined): NumberingModel {
  const abstracts = new Map<string, Map<number, LevelDef>>();
  const nums = new Map<string, NumDef>();
  const styles = new Map<string, StyleDef>();
  if (numberingXml) {
    const tree = parseXml(numberingXml);
    const root = kid(tree as unknown as XElement, "w:numbering");
    if (root) {
      for (const a of kids(root, "w:abstractNum")) {
        const id = attrOf(a, "w:abstractNumId");
        if (id === null) continue;
        const levels = new Map<number, LevelDef>();
        for (const lvl of kids(a, "w:lvl")) {
          const ilvl = attrOf(lvl, "w:ilvl");
          if (ilvl === null) continue;
          const fmt = attrOf(kid(lvl, "w:numFmt") ?? el0, "w:val") ?? "decimal";
          const text = attrOf(kid(lvl, "w:lvlText") ?? el0, "w:val") ?? "";
          const start = parseInt(attrOf(kid(lvl, "w:start") ?? el0, "w:val") ?? "1", 10) || 1;
          levels.set(parseInt(ilvl, 10) || 0, { fmt, text, start });
        }
        abstracts.set(id, levels);
      }
      for (const n of kids(root, "w:num")) {
        const numId = attrOf(n, "w:numId");
        const abs = attrOf(kid(n, "w:abstractNumId") ?? el0, "w:val");
        if (!numId || !abs) continue;
        const overrides = new Map<string, number>();
        for (const o of kids(n, "w:lvlOverride")) {
          const ilvl = attrOf(o, "w:ilvl");
          const so = kid(o, "w:startOverride");
          const val = so ? parseInt(attrOf(so, "w:val") ?? "", 10) : NaN;
          if (ilvl !== null && Number.isInteger(val)) overrides.set(ilvl, val);
        }
        nums.set(numId, { abstract: abs, overrides });
      }
    }
  }
  if (stylesXml) {
    const tree = parseXml(stylesXml);
    const root = kid(tree as unknown as XElement, "w:styles");
    if (root) {
      for (const s of kids(root, "w:style")) {
        const styleId = attrOf(s, "w:styleId");
        if (!styleId) continue;
        const def: StyleDef = {};
        const basedOn = kid(s, "w:basedOn");
        if (basedOn) def.basedOn = attrOf(basedOn, "w:val") ?? undefined;
        const pPr = kid(s, "w:pPr");
        const numPr = pPr ? kid(pPr, "w:numPr") : undefined;
        if (numPr) {
          def.numId = attrOf(kid(numPr, "w:numId") ?? el0, "w:val") ?? undefined;
          def.ilvl = attrOf(kid(numPr, "w:ilvl") ?? el0, "w:val") ?? undefined;
        }
        styles.set(styleId, def);
      }
    }
  }
  return new NumberingModel(abstracts, nums, styles);
}
