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
import { REVISION_NAMES } from "./projection.js";

export const REVISION_ALL = [...REVISION_NAMES, "w:rPrChange", "w:pPrChange", "w:sectPrChange", "w:tblPrChange"];

export function isRevisionEl(e: XElement): boolean {
  return REVISION_ALL.includes(e.name);
}

export function* revisionElements(root: XElement): Generator<XElement> {
  for (const e of walkAll(root)) if (isRevisionEl(e)) yield e;
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

export function newRevisionEl(name: "w:ins" | "w:del", id: string, author: string, date: string): XElement {
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
 */
export function resolveRevisions(root: XElement, match: (rev: XElement) => boolean, decision: "accept" | "reject"): number {
  const targets = [...revisionElements(root)].filter((r) => match(r));
  let n = 0;
  for (const t of targets) {
    if (t.name === "w:ins" || t.name === "w:moveTo") {
      if (decision === "accept") {
        const parent = t.parent;
        if (parent && parent.name === "w:rPr") removeChild(t); // paragraph mark
        else unwrap(t);
      } else removeChild(t);
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
