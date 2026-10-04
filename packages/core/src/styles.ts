// Word's visibility rule, checked against Word: a run's own w:vanish is
// absolute; otherwise the paragraph and character style chains combine by
// XOR, each taking its nearest definition along basedOn; docDefaults apply
// only when neither chain defines it; an absent or unknown pStyle means the
// default paragraph style. specVanish and webHidden runs print in Word.
import { el, kid, kids, parseXml, type XElement } from "@stitchdocx/package-io";
import { attrVal } from "./projection.js";
import { cloneRunPr } from "./revisions.js";

interface StyleDef {
  basedOn?: string;
  vanish?: boolean;
}

function toggle(rPr: XElement | undefined, name: string): boolean | undefined {
  const prop = rPr ? kid(rPr, name) : undefined;
  if (!prop) return undefined;
  const v = attrVal(prop, "w:val");
  return v === null || !(v === "0" || v === "false" || v === "off");
}

export class StyleSheet {
  private constructor(
    private readonly paragraphStyles: Map<string, StyleDef>,
    private readonly characterStyles: Map<string, StyleDef>,
    private readonly defaultParagraphStyle: string | undefined,
    private readonly defaultVanish: boolean,
  ) {}

  static readonly empty = new StyleSheet(new Map(), new Map(), undefined, false);

  static parse(stylesXml: string | undefined): StyleSheet {
    if (!stylesXml) return StyleSheet.empty;
    const root = kid(parseXml(stylesXml) as unknown as XElement, "w:styles");
    if (!root) return StyleSheet.empty;
    const paragraph = new Map<string, StyleDef>();
    const character = new Map<string, StyleDef>();
    let defaultParagraph: string | undefined;
    for (const s of kids(root, "w:style")) {
      const id = attrVal(s, "w:styleId");
      if (!id) continue;
      const type = attrVal(s, "w:type") ?? "paragraph";
      const def: StyleDef = { vanish: toggle(kid(s, "w:rPr"), "w:vanish") };
      const basedOn = kid(s, "w:basedOn");
      if (basedOn) def.basedOn = attrVal(basedOn, "w:val") ?? undefined;
      if (type === "paragraph") {
        paragraph.set(id, def);
        const d = attrVal(s, "w:default");
        if (defaultParagraph === undefined && (d === "1" || d === "true" || d === "on")) defaultParagraph = id;
      } else if (type === "character") character.set(id, def);
    }
    const rPrDefault = kid(kid(kid(root, "w:docDefaults"), "w:rPrDefault"), "w:rPr");
    return new StyleSheet(paragraph, character, defaultParagraph, toggle(rPrDefault, "w:vanish") ?? false);
  }

  runHidden(run: XElement, paragraph: XElement | null): boolean {
    const rPr = kid(run, "w:rPr");
    const direct = toggle(rPr, "w:vanish");
    if (direct !== undefined) return direct;
    return this.inherited(rPr, paragraph);
  }

  private inherited(rPr: XElement | undefined, paragraph: XElement | null): boolean {
    const pStyle = paragraph ? kid(kid(paragraph, "w:pPr"), "w:pStyle") : undefined;
    const pId = pStyle ? attrVal(pStyle, "w:val") : null;
    const fromParagraph = chainValue(this.paragraphStyles, pId !== null && this.paragraphStyles.has(pId) ? pId : this.defaultParagraphStyle);
    const rStyle = rPr ? kid(rPr, "w:rStyle") : undefined;
    const fromCharacter = chainValue(this.characterStyles, rStyle ? attrVal(rStyle, "w:val") ?? undefined : undefined);
    if (fromParagraph === undefined && fromCharacter === undefined) return this.defaultVanish;
    return (fromParagraph ?? false) !== (fromCharacter ?? false);
  }

  /** the source run's formatting for a run we insert into `paragraph`,
   *  forced visible with w:vanish w:val="0" when its styles would hide it */
  insertedRunPr(source: XElement | null, paragraph: XElement): XElement | null {
    let rPr = cloneRunPr(source);
    if (!this.inherited(rPr ?? undefined, paragraph)) return rPr;
    rPr ??= el("w:rPr");
    const vanish = el("w:vanish", { "w:val": "0" });
    const at = rPr.children.findIndex((c) => c.kind === "el" && AFTER_VANISH.has(c.name));
    rPr.children.splice(at < 0 ? rPr.children.length : at, 0, vanish);
    vanish.parent = rPr;
    return rPr;
  }
}

function chainValue(styles: Map<string, StyleDef>, id: string | undefined): boolean | undefined {
  const seen = new Set<string>();
  while (id !== undefined && !seen.has(id)) {
    seen.add(id);
    const def = styles.get(id);
    if (!def) return undefined;
    if (def.vanish !== undefined) return def.vanish;
    id = def.basedOn;
  }
  return undefined;
}

/** CT_RPr children that follow w:vanish (schema order) */
const AFTER_VANISH = new Set([
  "w:webHidden", "w:color", "w:spacing", "w:w", "w:kern", "w:position", "w:sz", "w:szCs", "w:highlight", "w:u",
  "w:effect", "w:bdr", "w:shd", "w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang", "w:eastAsianLayout",
  "w:specVanish", "w:oMath", "w:rPrChange",
]);
