// Projection of word/document.xml into paragraphs — semantics follow Word's
// documented tracked-change behavior (paragraph marks that are not visible
// merge the paragraph into the next one in the same container; text inside a
// complex field's instruction part is hidden), matching what Word and
// LibreOffice show, not any particular test reader.
//
// One walker serves two consumers: paragraphText (projection) and the span
// map (mutation input), so text and offsets can never disagree.
import { decodeAttrValue, decodeText, kid, kids, type XElement } from "@stitchdocx/package-io";
import { walkBody, rowsOf, cellsOf, rowRevision, type BodyContext, type BodyWalk } from "./body.js";
import { stepFields } from "./fields.js";

export const REVISION_NAMES = new Set(["w:ins", "w:del", "w:moveFrom", "w:moveTo"]);
const SKIPPED = new Set(["w:pPr", "w:rPr", "w:instrText", "w:delInstrText"]);
/** text boxes are not part of the paragraph's text and are never edited;
 *  AlternateContent branches are picked by mcBranch (the Word-supported
 *  Choice, else the Fallback) */
const UNSHOWN = new Set(["w:txbxContent"]);

/** run properties that hide a run from Word's display: hidden text is never
 *  part of the listed text, and never copied onto inserted runs */
const HIDDEN_PR = new Set(["w:vanish", "w:specVanish", "w:webHidden"]);

/** namespace prefixes an AlternateContent Choice may require for us to read
 *  that branch (the prefixes this engine understands) */
const MC_SUPPORTED = new Set(["w", "w14", "w15"]);

/** decision for a revision element */
export type View = (rev: { id: string | null; author: string | null; name: string }) => "accept" | "reject";

export const acceptAll: View = () => "accept";

const isInsertion = (name: string) => name === "w:ins" || name === "w:moveTo";

export function attrVal(el: XElement, name: string): string | null {
  const a = el.attrs.find((x) => x.name === name);
  return a ? decodeAttrValue(a.raw) : null;
}

/** true when the run carries a hide-from-display run property */
export function runHidden(r: XElement): boolean {
  const rPr = kid(r, "w:rPr");
  if (!rPr) return false;
  return kids(rPr).some((c) => HIDDEN_PR.has(c.name));
}

/** the AlternateContent branch Word would render: the first Choice whose
 *  Requires namespaces we understand, else the Fallback */
export function mcBranch(ac: XElement): XElement | null {
  for (const choice of kids(ac, "mc:Choice")) {
    const req = attrVal(choice, "Requires") ?? "";
    const prefixes = req.trim().split(/\s+/).filter(Boolean);
    if (prefixes.length > 0 && prefixes.every((p) => MC_SUPPORTED.has(p))) return choice;
  }
  return kid(ac, "mc:Fallback") ?? null;
}

/** the one character a w:sym contributes: the code point named by w:char */
export function symChar(el: XElement): string | null {
  let raw = attrVal(el, "w:char");
  if (raw === null) return null;
  raw = raw.replace(/^0[xX]/, "");
  const n = parseInt(raw, 16);
  return Number.isInteger(n) && n > 0 && n <= 0x10ffff ? String.fromCodePoint(n) : null;
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

export type CharKind = "text" | "tab" | "break" | "hyphen" | "sym";

export interface WalkCallbacks {
  onCharEl?: (el: XElement, kind: CharKind, inFieldResult: boolean, locked: boolean) => void;
}

/**
 * Walk a paragraph's content exactly like the reader's paragraphText: skip
 * pPr/rPr/instrText/delInstrText entirely, descend revisions with their
 * visibility, thread the complex-field cursor (begin/separate/end may sit in
 * different paragraphs — pass the shared cursor in), hide hidden runs, read
 * w:sym as one character, treat w:fldSimple content as a field result, pick
 * the Word-supported AlternateContent branch (whose content is listed but
 * locked), and report visible char elements in document order.
 */
export function walkContent(p: XElement, view: View, cb: WalkCallbacks, fields: string[] = []): void {
  visitKids(p, true, fields, false);

  function visitKids(node: XElement, shown: boolean, fields: string[], locked: boolean): void {
    for (const c of node.children) {
      if (c.kind !== "el") continue;
      visit(c, shown, fields, locked);
    }
  }

  function visit(c: XElement, shown: boolean, fields: string[], locked: boolean): void {
    if (SKIPPED.has(c.name)) return;
    if (UNSHOWN.has(c.name)) return;
    if (REVISION_NAMES.has(c.name)) {
      const s = shown && visibleRev(c, view);
      for (const cc of c.children) if (cc.kind === "el") visit(cc, s, fields, locked);
      return;
    }
    if (c.name === "w:fldChar") {
      stepFields(fields, c);
      return;
    }
    if (c.name === "w:fldSimple") {
      // the content of a simple field is a field result: shown, but never
      // edit-safe (Word discards the edit on the next field update)
      fields.push("result");
      visitKids(c, shown, fields, locked);
      fields.pop();
      return;
    }
    const inInstr = fields.includes("instr");
    if (c.name === "w:r") {
      if (runHidden(c)) return;
      visitKids(c, shown, fields, locked);
      return;
    }
    if (c.name === "w:sym") {
      if (shown && !inInstr) cb.onCharEl?.(c, "sym", fields.length > 0, locked);
      return;
    }
    if (c.name === "m:oMath" || c.name === "m:oMathPara") {
      // office math is listed but locked: the engine cannot re-write it
      visitKids(c, shown, fields, true);
      return;
    }
    if (c.name === "m:t") {
      if (shown && !inInstr) cb.onCharEl?.(c, "text", fields.length > 0, true);
      return;
    }
    if (c.name === "mc:AlternateContent") {
      const branch = mcBranch(c);
      if (branch) visitKids(branch, shown, fields, true);
      return;
    }
    if (c.name === "w:t" || c.name === "w:delText") {
      if (shown && !inInstr) cb.onCharEl?.(c, "text", fields.length > 0, locked);
      return;
    }
    if (c.name === "w:tab" || c.name === "w:ptab") {
      if (shown && !inInstr) cb.onCharEl?.(c, "tab", fields.length > 0, locked);
      return;
    }
    if (c.name === "w:br" || c.name === "w:cr") {
      if (shown && !inInstr) cb.onCharEl?.(c, "break", fields.length > 0, locked);
      return;
    }
    if (c.name === "w:noBreakHyphen") {
      if (shown && !inInstr) cb.onCharEl?.(c, "hyphen", fields.length > 0, locked);
      return;
    }
    if (c.name === "w:sdt") {
      const content = kid(c, "w:sdtContent");
      if (content) visitKids(content, shown, fields, locked);
      else visitKids(c, shown, fields, locked);
      return;
    }
    visitKids(c, shown, fields, locked);
  }
}

export function paragraphText(p: XElement, view: View, fields: string[] = []): string {
  let out = "";
  walkContent(p, view, {
    onCharEl: (el, kind) => {
      if (kind === "text") out += wordText(el);
      else if (kind === "sym") out += symChar(el) ?? "";
      else if (kind === "tab") out += "\t";
      else if (kind === "break") out += "\n";
      else out += "-";
    },
  }, fields);
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
  /** true when this record's text begins with content carried from a removed
   *  mark in front of it */
  carried: boolean;
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
  // the complex-field cursor is threaded through the whole story: a TOC's
  // begin/end fldChars sit in different paragraphs
  const fields: string[] = [];

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
          text: (pending?.text ?? "") + paragraphText(c, view, fields),
          removedMark: null,
          carried: pending !== null,
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
        let r = 0;
        for (const tr of rowsOf(c)) {
          // row-level tracked changes are applied to the view the way Word
          // applies them: a deleted row disappears under accept, an inserted
          // row under reject
          const rev = rowRevision(tr);
          if (rev && !visibleRev(rev, view)) {
            r++;
            continue;
          }
          let col = 0;
          for (const tc of cellsOf(tr)) {
            // a removed mark before the table merges into the first cell's
            // first paragraph; only that cell receives the carry
            const left = visit(tc, { ...ctx, cell: true, table: { index: t.index, row: r, col } }, pending);
            pending = null;
            if (left && !(left.removedMark === "w:del" && left.text === "")) out.push(left);
            col++;
          }
          r++;
        }
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
