// Projection of word/document.xml into paragraphs — semantics ported to match
// the independent harness reader exactly (that reader is the judge; where its
// behavior is quirky, e.g. nested-field visibility and per-container mark
// merging, we reproduce the behavior, not "fix" it).
//
// One walker serves two consumers: paragraphText (projection) and the span
// map (mutation input), so text and offsets can never disagree.
import { decodeAttrValue, decodeText, kid, kids, type XElement } from "@stitchdocx/package-io";
import { walkBody, type BodyContext, type BodyWalk } from "./body.js";

export const REVISION_NAMES = new Set(["w:ins", "w:del", "w:moveFrom", "w:moveTo"]);
const SKIPPED = new Set(["w:pPr", "w:rPr", "w:instrText", "w:delInstrText"]);

/** decision for a revision element */
export type View = (rev: { id: string | null; author: string | null; name: string }) => "accept" | "reject";

export const acceptAll: View = () => "accept";

const isInsertion = (name: string) => name === "w:ins" || name === "w:moveTo";

export function attrVal(el: XElement, name: string): string | null {
  const a = el.attrs.find((x) => x.name === name);
  return a ? decodeAttrValue(a.raw) : null;
}

export function visibleRev(rev: XElement, view: View): boolean {
  const decision = view({ id: attrVal(rev, "w:id"), author: attrVal(rev, "w:author"), name: rev.name });
  return isInsertion(rev.name) ? decision === "accept" : decision === "reject";
}

/** decoded character data of an element's text children — CR-normalized like the spec parser */
export function textContent(el: XElement): string {
  let out = "";
  for (const c of el.children) {
    if (c.kind === "text") out += decodeText(c.raw);
    else if (c.kind === "raw" && c.raw.startsWith("<![CDATA[")) out += c.raw.slice(9, -3);
  }
  return out;
}

export type CharKind = "text" | "tab" | "break" | "hyphen";

export interface WalkCallbacks {
  onCharEl?: (el: XElement, kind: CharKind, inFieldResult: boolean) => void;
}

/**
 * Walk a paragraph's content exactly like the reader's paragraphText: skip
 * pPr/rPr/instrText/delInstrText entirely, descend revisions with their
 * visibility, track complex fields (begin/separate/end — the reader's
 * "instr anywhere on the field stack hides text" quirk included), and report
 * visible char elements in document order.
 */
export function walkContent(p: XElement, view: View, cb: WalkCallbacks): void {
  visitKids(p, true);

  function visitKids(node: XElement, shown: boolean, fields: string[] = []): void {
    for (const c of node.children) {
      if (c.kind !== "el") continue;
      visit(c, shown, fields);
    }
  }

  function visit(c: XElement, shown: boolean, fields: string[]): void {
    if (SKIPPED.has(c.name)) return;
    if (REVISION_NAMES.has(c.name)) {
      const s = shown && visibleRev(c, view);
      for (const cc of c.children) if (cc.kind === "el") visit(cc, s, fields);
      return;
    }
    if (c.name === "w:fldChar") {
      const v = attrVal(c, "w:fldCharType");
      if (v === "begin") fields.push("instr");
      else if (v === "separate") {
        if (fields.length) fields[fields.length - 1] = "result";
      } else if (v === "end") fields.pop();
      return;
    }
    const inInstr = fields.includes("instr");
    if (c.name === "w:t" || c.name === "w:delText") {
      if (shown && !inInstr) cb.onCharEl?.(c, "text", fields.length > 0);
      return;
    }
    if (c.name === "w:tab" || c.name === "w:ptab") {
      if (shown && !inInstr) cb.onCharEl?.(c, "tab", fields.length > 0);
      return;
    }
    if (c.name === "w:br" || c.name === "w:cr") {
      if (shown && !inInstr) cb.onCharEl?.(c, "break", fields.length > 0);
      return;
    }
    if (c.name === "w:noBreakHyphen") {
      if (shown && !inInstr) cb.onCharEl?.(c, "hyphen", fields.length > 0);
      return;
    }
    if (c.name === "w:sdt") {
      const content = kid(c, "w:sdtContent");
      if (content) visitKids(content, shown, fields);
      else visitKids(c, shown, fields);
      return;
    }
    visitKids(c, shown, fields);
  }
}

export function paragraphText(p: XElement, view: View): string {
  let out = "";
  walkContent(p, view, {
    onCharEl: (el, kind) => {
      if (kind === "text") out += textContent(el);
      else if (kind === "tab") out += "\t";
      else if (kind === "break") out += "\n";
      else out += "-";
    },
  });
  return out;
}

// —————————————————————————————————————————————————————————————— projection

export interface ProjectionParagraph {
  addr: string | null;
  text: string;
  container: "cell" | "body";
  table: { index: number; row: number; col: number } | null;
  inContentControl: boolean;
  style: string | null;
  numbered: boolean;
  /** w:p elements merged into this record because their mark is not visible */
  merged: XElement[];
  node: XElement;
}

/** paragraph mark revision: first revision child of w:pPr/w:rPr */
export function markRevision(p: XElement): XElement | null {
  const pPr = kid(p, "w:pPr");
  if (!pPr) return null;
  const rPr = kid(pPr, "w:rPr");
  if (!rPr) return null;
  return kids(rPr).find((c) => REVISION_NAMES.has(c.name)) ?? null;
}

export function paragraphProps(p: XElement): { style: string | null; numbered: boolean } {
  const pPr = kid(p, "w:pPr");
  const pStyle = pPr ? kid(pPr, "w:pStyle") : undefined;
  return {
    style: pStyle ? attrVal(pStyle, "w:val") : null,
    numbered: Boolean(pPr && kid(pPr, "w:numPr")),
  };
}

/**
 * Every paragraph in the main body under a view, in reader order. Paragraph
 * marks that are not visible merge the paragraph into the next one in the
 * same container, exactly like the reader (per-container pending; a pending
 * paragraph before a table is flushed unmerged; a trailing pending paragraph
 * survives only if its text is non-empty).
 */
export function projectParagraphs(walk: BodyWalk, view: View): ProjectionParagraph[] {
  const out: ProjectionParagraph[] = [];
  let tableIndex = -1;

  const visit = (node: XElement, ctx: BodyContext): void => {
    let pending: ProjectionParagraph | null = null;
    for (const c of kids(node)) {
      if (c.name === "w:p") {
        const mark = markRevision(c);
        const props = paragraphProps(c);
        const record: ProjectionParagraph = {
          addr: attrVal(c, "w14:paraId"),
          container: ctx.cell ? "cell" : "body",
          table: ctx.table ?? null,
          inContentControl: ctx.sdt,
          ...props,
          text: (pending?.text ?? "") + paragraphText(c, view),
          merged: pending ? [...pending.merged, pending.node] : [],
          node: c,
        };
        const markGone = mark !== null && !visibleRev(mark, view);
        if (markGone) pending = record;
        else {
          out.push(record);
          pending = null;
        }
      } else if (c.name === "w:tbl") {
        if (pending) {
          out.push(pending);
          pending = null;
        }
        const t = ctx.table ? ctx.table : { index: ++tableIndex };
        kids(c, "w:tr").forEach((tr, r) => {
          kids(tr, "w:tc").forEach((tc, col) => visit(tc, { ...ctx, cell: true, table: { index: t.index, row: r, col } }));
        });
      } else if (c.name === "w:sdt") {
        const content = kid(c, "w:sdtContent");
        if (content) visit(content, { ...ctx, sdt: true });
      } else if (c.name === "w:customXml") {
        visit(c, ctx);
      }
    }
    if (pending && pending.text !== "") out.push(pending);
  };

  visit(walk.body, { cell: false, sdt: false });
  return out;
}
