// Projection of word/document.xml into paragraphs — semantics follow Word's
// documented tracked-change behavior (paragraph marks that are not visible
// merge the paragraph into the next one in the same container; text inside a
// complex field's instruction part is hidden), matching what Word and
// LibreOffice show, not any particular test reader.
//
// One walker serves two consumers: paragraphText (projection) and the span
// map (mutation input), so text and offsets can never disagree.
import { decodeAttrValue, decodeText, isPending, kid, kids, pendingSource, readView, type XElement } from "@stitchdocx/package-io";
import { walkBody, rowsOf, cellsOf, rowRevision, type BodyContext, type BodyWalk } from "./body.js";
import { stepFields } from "./fields.js";
import type { StyleSheet } from "./styles.js";

export const REVISION_NAMES = new Set(["w:ins", "w:del", "w:moveFrom", "w:moveTo"]);
const SKIPPED = new Set(["w:pPr", "w:rPr", "w:instrText", "w:delInstrText"]);
/** text boxes are not part of the paragraph's text and are never edited;
 *  AlternateContent branches are picked by mcBranch (the Word-supported
 *  Choice, else the Fallback) */
const UNSHOWN = new Set(["w:txbxContent"]);

/** namespace prefixes current Word renders in an AlternateContent Choice */
const MC_SUPPORTED = new Set([
  "w14", "w15", "w16", "w16se", "w16cid", "w16cex", "w16du", "w16sdtdh",
  "wps", "wpg", "wpc", "wpi", "wp14", "a14", "v", "o", "w10", "m",
]);

/** decision for a revision element */
export type View = (rev: { id: string | null; author: string | null; name: string }) => "accept" | "reject";

export const acceptAll: View = () => "accept";

const isInsertion = (name: string) => name === "w:ins" || name === "w:moveTo";

export function attrVal(el: XElement, name: string): string | null {
  const a = el.attrs.find((x) => x.name === name);
  return a ? decodeAttrValue(a.raw) : null;
}

/** the AlternateContent branch Word renders: the first Choice whose
 *  Requires prefixes Word all supports, else the Fallback */
export function mcBranch(ac: XElement): XElement | null {
  for (const choice of kids(ac, "mc:Choice")) {
    const req = attrVal(choice, "Requires") ?? "";
    const prefixes = req.trim().split(/\s+/).filter(Boolean);
    if (prefixes.length > 0 && prefixes.every((p) => MC_SUPPORTED.has(p))) return choice;
  }
  return kid(ac, "mc:Fallback") ?? null;
}

/** the one character a w:sym (w:char) or w16se:symEx (w16se:char) contributes */
export function symChar(el: XElement): string | null {
  let raw = attrVal(el, el.name === "w16se:symEx" ? "w16se:char" : "w:char");
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
export function walkContent(p: XElement, view: View, styles: StyleSheet, cb: WalkCallbacks, fields: string[] = []): void {
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
      if (styles.runHidden(c, p)) return;
      visitKids(c, shown, fields, locked);
      return;
    }
    if (c.name === "w:sym" || c.name === "w16se:symEx") {
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

export function paragraphText(p: XElement, view: View, styles: StyleSheet, fields: string[] = []): string {
  let out = "";
  walkContent(p, view, styles, {
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
  /** the paragraph's own w:numPr reference, when it has one */
  numPr: NumPrRef | null;
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

/** every revision on a paragraph's mark, in source order */
function markRevisions(p: XElement): XElement[] {
  const rPr = kid(kid(p, "w:pPr"), "w:rPr");
  return rPr ? kids(rPr).filter((c) => REVISION_NAMES.has(c.name)) : [];
}

/** a w:numPr as written: numId and ilvl values, each null when absent */
export interface NumPrRef {
  numId: string | null;
  ilvl: string | null;
}

export function paragraphProps(p: XElement): { style: string | null; numbered: boolean; numPr: NumPrRef | null } {
  const pPr = kid(p, "w:pPr");
  const pStyle = pPr ? kid(pPr, "w:pStyle") : undefined;
  const numPr = pPr ? kid(pPr, "w:numPr") : undefined;
  const val = (name: string) => {
    const e = numPr ? kid(numPr, name) : undefined;
    return e ? attrVal(e, "w:val") : null;
  };
  return {
    style: pStyle ? attrVal(pStyle, "w:val") : null,
    numbered: Boolean(numPr),
    numPr: numPr ? { numId: val("w:numId"), ilvl: val("w:ilvl") } : null,
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
 *
 * With `only`, the result is just the records with those addresses. Every
 * paragraph is still walked, since a removed mark or a field can reach across
 * paragraphs, but an unparsed paragraph outside `only` whose source holds no
 * revision or field markup cannot (its mark is visible and it leaves the field
 * cursor as it found it), so it is passed over without being parsed.
 */
export function projectParagraphs(walk: BodyWalk, view: View, styles: StyleSheet, only?: ReadonlySet<string>): ProjectionParagraph[] {
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
        if (only && !only.has(attrVal(c, "w14:paraId") ?? "") && isPending(c) && !CROSSES_PARAGRAPHS.test(pendingSource(c)!)) {
          // its record, carried text and all, is not one asked for
          pending = null;
          continue;
        }
        // an unparsed paragraph is read through a view dropped after this
        // record, so listing never holds more than one paragraph parsed
        const content = readView(c);
        const mark = markRevision(content);
        const props = paragraphProps(content);
        const record: ProjectionParagraph = {
          addr: attrVal(c, "w14:paraId"),
          container: ctx.cell ? "cell" : "body",
          table: ctx.table ?? null,
          inContentControl: ctx.sdt,
          ...props,
          text: (pending?.text ?? "") + paragraphText(content, view, styles, fields),
          removedMark: null,
          carried: pending !== null,
          merged: pending ? [...pending.merged, pending.node] : [],
          node: c,
        };
        // a mark inserted and later deleted carries both revisions, and is
        // removed under either view
        const hidden = mark ? markRevisions(content).find((r) => !visibleRev(r, view)) : undefined;
        if (hidden) {
          record.removedMark = hidden.name;
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
  return only ? out.filter((r) => only.has(r.addr ?? "")) : out;
}

/** markup through which a paragraph can affect another's record: a tracked
 *  change (which may sit on its mark) or a complex or simple field */
const CROSSES_PARAGRAPHS = /<w:(?:ins|del|moveFrom|moveTo)[\s/>]|<w:fld(?:Char|Simple)[\s/>]/;
