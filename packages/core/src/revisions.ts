// Revision machinery: id allocation, tracked-change element construction,
// del/ins wrapping, and accept/reject transforms.
import {
  cloneNode,
  decodeEntities,
  el,
  encodeText,
  getAttr,
  kid,
  kids,
  removeChild,
  setAttr,
  textNode,
  type XElement,
  type XNode,
} from "@stitchdocx/package-io";
import { REVISION_NAMES, markRevision } from "./projection.js";

export const REVISION_ALL = [...REVISION_NAMES, "w:rPrChange", "w:pPrChange", "w:sectPrChange", "w:tblPrChange"];

export function isRevisionEl(e: XElement): boolean {
  return REVISION_ALL.includes(e.name);
}

export function* revisionElements(root: XElement): Generator<XElement> {
  for (const e of walkAll(root)) if (isRevisionEl(e)) yield e;
}

/** revision elements in document order (walkAll's stack walk is reversed) */
function revisionElementsInOrder(root: XElement): XElement[] {
  const out: XElement[] = [];
  const visit = (node: XElement): void => {
    if (isRevisionEl(node)) out.push(node);
    for (const c of node.children) if (c.kind === "el") visit(c);
  };
  visit(root);
  return out;
}

export function* walkAll(root: XElement): Generator<XElement> {
  const stack: XElement[] = [root];
  while (stack.length) {
    const cur = stack.pop()!;
    yield cur;
    for (const c of cur.children) if (c.kind === "el") stack.push(c);
  }
}

/** monotonic w:id allocator, seeded above every existing revision id */
export class RevisionIdAllocator {
  private next: number;
  constructor(root: XElement) {
    let max = 0;
    for (const r of revisionElements(root)) {
      const id = getAttr(r, "w:id");
      const n = id === null ? NaN : Number(id);
      if (Number.isInteger(n) && n > max) max = n;
    }
    this.next = max + 1;
  }
  take(): string {
    return String(this.next++);
  }
}

const PARAID_MAX = 0x7fffffff;

/** deterministic w14:paraId allocator over the document's existing ids */
export class ParaIdAllocator {
  private next: number;
  private used: Set<string>;
  constructor(paragraphs: XElement[]) {
    this.used = new Set();
    let max = 0;
    for (const p of paragraphs) {
      const id = getAttr(p, "w14:paraId");
      if (id === null) continue;
      this.used.add(id.toUpperCase());
      const n = parseInt(id, 16);
      if (Number.isInteger(n) && n > max && n <= PARAID_MAX) max = n;
    }
    this.next = max >= PARAID_MAX ? 1 : max + 1;
  }
  take(): string {
    let id = "";
    do {
      id = this.next.toString(16).toUpperCase().padStart(8, "0");
      this.next = this.next >= PARAID_MAX ? 1 : this.next + 1;
    } while (this.used.has(id));
    this.used.add(id);
    return id;
  }
}

export function newRevisionEl(name: "w:ins" | "w:del" | "w:pPrChange", id: string, author: string, date: string): XElement {
  return el(name, { "w:id": id, "w:author": author, "w:date": date });
}

/** rename an element in place, keeping its attribute source form */
export function renameEl(e: XElement, newName: string): void {
  if (!e.attrsDirty) {
    const attrsPart = e.rawOpen.slice(e.name.length);
    e.rawOpen = newName + attrsPart;
  }
  e.name = newName;
  e.rawClose = newName; // the end tag must match the new name
}

/** replace an element with its own children (used by accept/reject) */
export function unwrap(el_: XElement): void {
  const parent = el_.parent;
  if (!parent) return;
  const i = parent.children.indexOf(el_);
  if (i < 0) return;
  parent.children.splice(i, 1, ...el_.children);
  for (const c of parent.children) if (c.kind === "el") c.parent = parent;
}

/** run content from contract text: \n becomes w:br, \t becomes w:tab */
export function textToRunChildren(text: string): XNode[] {
  const children: XNode[] = [];
  let buf = "";
  const flush = (): void => {
    if (buf.length) {
      const t = el("w:t", { "xml:space": "preserve" });
      t.children.push(textNode(buf));
      children.push(t);
      buf = "";
    }
  };
  for (const ch of text) {
    if (ch === "\n") {
      flush();
      children.push(el("w:br"));
    } else if (ch === "\t") {
      flush();
      children.push(el("w:tab"));
    } else buf += ch;
  }
  flush();
  return children;
}

/** clone of a run's w:rPr, or null (null-safe on runs without formatting) */
export function cloneRunPr(run: XElement | null): XElement | null {
  const rPr = run ? kid(run, "w:rPr") : null;
  return rPr ? (cloneNode(rPr) as XElement) : null;
}

export function makeRun(rPr: XElement | null, content: XNode[]): XElement {
  const run = el("w:r");
  if (rPr) run.children.push(rPr);
  for (const c of content) run.children.push(c);
  for (const c of run.children) if (c.kind === "el") c.parent = run;
  return run;
}

// —————————————————————————————————————————————————————— accept / reject

/**
 * Resolve matched revisions in place, across every revision kind the engine
 * recognizes. w:ins/w:moveTo (insertions): accept unwraps (a paragraph-mark
 * ins is simply removed), reject removes. w:del/w:moveFrom (deletions):
 * accept removes, reject unwraps and restores the original text. Property
 * changes: accept removes the marker (current props stay), reject restores
 * the saved props carried inside the marker.
 *
 * Rejecting a paragraph-mark ins merges the paragraph into the next w:p in
 * its container (that is what the mark means: OOXML tracks the mark, and
 * rejecting it un-paragraphs the text). When there is no next paragraph —
 * container end, or a table follows — Word cannot remove the mark either, so
 * the paragraph stays and only the tracking is removed. When the next
 * paragraph carries a same-author w:pPrChange that is rejected in the same
 * call (the markup applyInsertParagraph writes at a container end, and the
 * markup Word itself writes there), the pair dissolves instead: the inserted
 * paragraph is removed and the anchor survives untouched.
 */
export function resolveRevisions(root: XElement, match: (rev: XElement) => boolean, decision: "accept" | "reject"): number {
  // document order matters: a paragraph-mark ins must be resolved while the
  // paired w:pPrChange on the following paragraph is still in place
  const targets = revisionElementsInOrder(root).filter((r) => match(r));
  const targetSet = new Set(targets);
  let n = 0;
  for (const t of targets) {
    if (!attached(t, root)) {
      n++; // already gone: it lived inside a paragraph removed earlier in this pass
      continue;
    }
    if (t.name === "w:ins" || t.name === "w:moveTo") {
      if (decision === "accept") {
        const parent = t.parent;
        if (parent && parent.name === "w:rPr") removeChild(t); // paragraph mark
        else unwrap(t);
      } else {
        const rPr = t.parent;
        const pPr = rPr?.parent;
        const para = pPr?.parent;
        if (rPr && rPr.name === "w:rPr" && pPr && pPr.name === "w:pPr" && para && para.name === "w:p") {
          rejectParagraphMark(para, t, targetSet, root);
        } else removeChild(t);
      }
    } else if (t.name === "w:del" || t.name === "w:moveFrom") {
      if (decision === "accept") removeChild(t);
      else {
        for (const d of descendants(t)) if (d.name === "w:delText") renameEl(d, "w:t");
        unwrap(t);
      }
    } else if (t.name.endsWith("PrChange")) {
      const parent = t.parent;
      if (!parent) continue;
      if (decision === "accept") {
        removeChild(t);
      } else {
        // restore the saved props: the marker's element child holds the old
        // props, which REPLACE the current ones (no merge)
        const saved = t.children.find((c) => c.kind === "el");
        const i = parent.children.indexOf(t);
        if (i >= 0) parent.children.splice(i, 1);
        if (saved) {
          parent.children = [...saved.children];
          for (const c of parent.children) if (c.kind === "el") c.parent = parent;
        }
      }
    }
    n++;
  }
  return n;
}

/** block-level markers that carry no content and may sit between runs */
const CONTENT_NEUTRAL = new Set([
  "w:bookmarkStart",
  "w:bookmarkEnd",
  "w:proofErr",
  "w:commentRangeStart",
  "w:commentRangeEnd",
  "w:commentReference",
]);

/**
 * True when `para` is the inserted half of a container-end insert pair written
 * by `author` (the markup applyInsertParagraph writes, and Word writes for the
 * same edit): it carries a same-author w:pPrChange, its own mark is untracked
 * (so it is the last paragraph of its chain), and every content child is
 * tracked-inserted by that same author. Only then can resolving the pair
 * dissolve the paragraph without losing content the user kept.
 */
export function isPairedInsertParagraph(para: XElement, author: string): boolean {
  const pPr = kid(para, "w:pPr");
  const change = pPr ? kid(pPr, "w:pPrChange") : null;
  if (!change || attrOf(change, "w:author") !== author) return false;
  if (markRevision(para)) return false; // a tracked mark means a chain middle: merge, don't dissolve
  let sawOwnInsertion = false;
  for (const c of para.children) {
    if (c.kind !== "el" || c.name === "w:pPr" || CONTENT_NEUTRAL.has(c.name)) continue;
    if ((c.name === "w:ins" || c.name === "w:moveTo") && attrOf(c, "w:author") === author) {
      sawOwnInsertion = true;
      continue;
    }
    return false; // plain, deleted, or other-authored content must survive
  }
  return sawOwnInsertion;
}

/**
 * Reject the tracked-inserted paragraph mark `mark` on paragraph `para`.
 * Merges `para` into the next w:p sibling when there is one and the mark can
 * meaningfully be removed; dissolves a same-author paired insert (next
 * paragraph carries a rejected same-author w:pPrChange — the container-end
 * insert form) by removing the inserted paragraph and restoring `para`
 * exactly; otherwise only untracks the mark.
 */
function rejectParagraphMark(para: XElement, mark: XElement, rejected: Set<XElement>, root: XElement): void {
  const container = para.parent;
  if (!container) {
    removeChild(mark);
    return;
  }
  const next = container.children[container.children.indexOf(para) + 1];
  if (next && next.kind === "el" && next.name === "w:p") {
    const nextPPr = kid(next, "w:pPr");
    const change = nextPPr ? kid(nextPPr, "w:pPrChange") : null;
    if (change && rejected.has(change) && isPairedInsertParagraph(next, attrOf(mark, "w:author") ?? "")) {
      // paired insert: the next paragraph exists only because of this mark —
      // drop it and leave the anchor as it was before the op
      removeChild(next);
      removeChild(mark);
      return;
    }
    // a paragraph carrying a section break cannot dissolve into the next one
    const paraPPr = kid(para, "w:pPr");
    if (paraPPr && kid(paraPPr, "w:sectPr")) {
      removeChild(mark);
      return;
    }
    // merge: para's content moves in front of the next paragraph's content
    const at = next.children[0] && next.children[0].kind === "el" && next.children[0].name === "w:pPr" ? 1 : 0;
    const moving = para.children.filter((c) => !(c.kind === "el" && c.name === "w:pPr"));
    next.children.splice(at, 0, ...moving);
    for (const c of moving) if (c.kind === "el") c.parent = next;
    container.children.splice(container.children.indexOf(para), 1);
    para.parent = null;
    return;
  }
  // container end, or a table follows: Word cannot remove the mark either
  removeChild(mark);
}

function attached(e: XElement, root: XElement): boolean {
  let cur: XElement | null = e;
  while (cur) {
    if (cur === root) return true;
    cur = cur.parent;
  }
  return false;
}

function* descendants(root: XElement): Generator<XElement> {
  for (const c of root.children) {
    if (c.kind === "el") {
      yield c;
      yield* descendants(c);
    }
  }
}

/** every revision authored by `author` (exact match) */
export function revisionsByAuthor(root: XElement, author: string): XElement[] {
  return [...revisionElements(root)].filter((r) => getAttr(r, "w:author") === author);
}

export function revisionIdsOf(els: XElement[]): string[] {
  return els.map((e) => getAttr(e, "w:id") ?? "").filter((s) => s !== "");
}

export const attrOf = getAttr;
export { encodeText };
