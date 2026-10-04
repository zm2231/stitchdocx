// Projection of word/document.xml into paragraphs — semantics follow Word's
// documented tracked-change behavior (paragraph marks that are not visible
// merge the paragraph into the next one in the same container; text inside a
// complex field's instruction part is hidden), matching what Word and
// LibreOffice show, not any particular test reader.
//
// One walker serves two consumers: paragraphText (projection) and the span
// map (mutation input), so text and offsets can never disagree.
import { decodeAttrValue, decodeText, kid, kids, type XElement } from "@stitchdocx/package-io";
import { walkBody, type BodyContext, type BodyWalk } from "./body.js";

export const REVISION_NAMES = new Set(["w:ins", "w:del", "w:moveFrom", "w:moveTo"]);
const SKIPPED = new Set(["w:pPr", "w:rPr", "w:instrText", "w:delInstrText"]);
/** text boxes are not part of the paragraph's text and are never edited; Word
 *  renders mc:Choice, never mc:Fallback */
const UNSHOWN = new Set(["w:txbxContent", "mc:Fallback"]);

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

/**
 * Text of a w:t or w:delText as Word reads it: without xml:space="preserve",
 * leading and trailing space/tab/newline are dropped; in either mode a newline
 * or tab character reads as a space.
 */
export function wordText(el: XElement): string {
  let v = textContent(el);
  if (attrVal(el, "xml:space") !== "preserve") v = v.replace(/^[ \t\r\n]+|[ \t\r\n]+$/g, "");
  return v.replace(/[\t\r\n]/g, " ");
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
    if (UNSHOWN.has(c.name)) return;
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
      if (kind === "text") out += wordText(el);
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
  /** name of the revision element on this paragraph's removed mark, if the
   *  mark is not visible under the view ("w:del" = a deleted mark under
   *  accept, "w:ins" = an inserted mark under reject) */
  removedMark: string | null;
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
 * Every paragraph in the main body under a view, in reader order. A paragraph
 * whose mark is removed (an inserted mark rejected, a deleted mark accepted)
 * merges into the next paragraph, as Word does:
 * - directly before a table, into the first paragraph of the table's first cell;
 * - at the end of a cell, a removed deletion that leaves the paragraph empty
 *   disappears, and any other removed mark stays;
 * - at the end of the body or of a content control, the paragraph stays,
 *   empty or not.
 */
export function projectParagraphs(walk: BodyWalk, view: View): ProjectionParagraph[] {
  const out: ProjectionParagraph[] = [];
  let tableIndex = -1;

  /** returns the container's trailing removed-mark paragraph, if any */
  const visit = (node: XElement, ctx: BodyContext, carry: ProjectionParagraph | null = null): ProjectionParagraph | null => {
    let pending = carry;
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
          removedMark: null,
          merged: pending ? [...pending.merged, pending.node] : [],
          node: c,
        };
        const markGone = mark !== null && !visibleRev(mark, view);
        if (markGone) {
          record.removedMark = mark!.name;
          pending = record;
        } else {
          out.push(record);
          pending = null;
        }
      } else if (c.name === "w:tbl") {
        const t = ctx.table ? ctx.table : { index: ++tableIndex };
        kids(c, "w:tr").forEach((tr, r) => {
          kids(tr, "w:tc").forEach((tc, col) => {
            // a removed mark before the table merges into the first cell's
            // first paragraph; only that cell receives the carry
            const left = visit(tc, { ...ctx, cell: true, table: { index: t.index, row: r, col } }, pending);
            pending = null;
            if (left && !(left.removedMark === "w:del" && left.text === "")) out.push(left);
          });
        });
      } else if (c.name === "w:sdt") {
        const content = kid(c, "w:sdtContent");
        if (content) {
          const left = visit(content, { ...ctx, sdt: true }, pending);
          pending = null;
          if (left) out.push(left);
        }
      } else if (c.name === "w:customXml") {
        pending = visit(c, ctx, pending);
      }
    }
    return pending;
  };

  const left = visit(walk.body, { cell: false, sdt: false });
  if (left) out.push(left);
  return out;
}
