// Document: the in-memory engine. The XML tree is the document; the model is
// an index over it. Save re-serializes only what changed — untouched
// paragraphs and untouched package entries are byte-identical by construction.
import {
  cloneNode,
  decodePart,
  el,
  encodePart,
  getAttr,
  kid,
  kids,
  mainPartName,
  readPackage,
  setAttr,
  writePackage,
  parseXml,
  serializeXml,
  type Entries,
  type XElement,
} from "@stitchdocx/package-io";
import { walkBody, type BodyWalk } from "./body.js";
import { acceptAll, paragraphText, projectParagraphs, REVISION_NAMES, type ProjectionParagraph, type View } from "./projection.js";
import { executeApply, type ApplyDeps } from "./ops.js";
import { cloneRunPr, isPairedInsertParagraph, ParaIdAllocator, resolveRevisions, RevisionIdAllocator, textToRunChildren, makeRun } from "./revisions.js";
import { normalizeDocx } from "./normalize.js";
import type { Addr, Decision, DecisionResult, Op, Paragraph, Receipt, RevisionDecision } from "./types.js";

/** the parsed document root is an XDoc; kid() works on it identically */
const treeAsEl = (d: unknown) => d as XElement;

export class Document {
  private entries: Entries;
  private mainPart: string;
  private tree: ReturnType<typeof parseXml>;
  private root: XElement;
  private walk: BodyWalk;
  private paraByAddr: Map<Addr, XElement>;
  private ids: RevisionIdAllocator;
  private paraIds: ParaIdAllocator;
  private originalBytes: Uint8Array;
  private dirty = false;
  private numbering?: NumberingModel;

  private constructor(bytes: Uint8Array) {
    this.originalBytes = bytes;
    this.entries = readPackage(bytes);
    this.mainPart = mainPartName(this.entries);
    const docXml = this.entries.get(this.mainPart);
    if (!docXml) throw new Error(`${this.mainPart} missing`);
    this.tree = parseXml(decodePart(docXml));
    this.root = kid(treeAsEl(this.tree), "w:document")!;
    if (!this.root) throw new Error("word/document.xml has no w:document root");
    this.walk = walkBody(this.root);
    this.paraByAddr = new Map();
    for (const p of this.walk.paragraphs) {
      const addr = getAttr(p, "w14:paraId");
      if (addr !== null) this.paraByAddr.set(addr, p);
    }
    this.ids = new RevisionIdAllocator(this.root);
    this.paraIds = new ParaIdAllocator(this.walk.paragraphs);
  }

  static open(bytes: Uint8Array): Document {
    return new Document(bytes);
  }

  /** projection under a view; addresses are w14:paraId (present after normalize) */
  projection(view: View = acceptAll): Paragraph[] {
    return projectParagraphs(this.walk, view).map((p) => this.toPublicParagraph(p));
  }

  /** raw projection records (engine-internal shape with merged nodes) */
  projectionRecords(view: View = acceptAll): ProjectionParagraph[] {
    return projectParagraphs(this.walk, view);
  }

  private toPublicParagraph(p: ProjectionParagraph): Paragraph {
    const kind: Paragraph["kind"] = p.container === "cell" ? "cell" : p.numbered ? "listItem" : /(^|[^a-z])heading|title/i.test(p.style ?? "") ? "heading" : "body";
    // a record whose text includes content carried across a removed mark (its
    // own or a preceding paragraph's) is a merged VIEW: its offsets span more
    // than the one paragraph an op may edit, so it is not addressable as listed
    const mergedView = p.removedMark !== null || p.carried;
    const out: Paragraph = {
      addr: p.addr ?? "",
      text: p.text,
      kind,
      editable: !mergedView,
      readOnlyReason: mergedView ? "merged_removed_mark" : undefined,
    };
    if (p.table) out.table = p.table;
    if (p.style) out.style = p.style;
    if (p.numbered) out.numbering = this.numberingLabel(p);
    return out;
  }

  private numberingLabel(p: ProjectionParagraph): string | undefined {
    if (!this.numbering) {
      const numberingXml = this.entries.get("word/numbering.xml");
      this.numbering = numberingXml ? parseNumbering(decodePart(numberingXml)) : emptyNumbering;
    }
    return this.numbering.labelFor(p.node);
  }

  apply(ops: Op[], opts: { date: string }): Receipt[] {
    this.recordTextCache = null;
    const deps: ApplyDeps = {
      paraByAddr: this.paraByAddr,
      ids: this.ids,
      paraIds: this.paraIds,
      date: opts.date,
      // ownership is scoped to THIS apply batch: revisions from earlier calls
      // are pre-existing tracked changes and must refuse, not reposition
      ownedRevisions: new Set<XElement>(),
      recordText: (addr, view) => this.recordText(addr, view),
    };
    const receipts = executeApply(deps, ops);
    if (receipts.some((r) => r.status === "applied")) this.markDirty();
    // register paragraphs created by inserts so later calls can address them
    for (const r of receipts) if (r.insertedAddr) this.paraByAddr.set(r.insertedAddr, this.findByAddrIn(r.insertedAddr)!);
    return receipts;
  }

  /** record-space text for receipts; the acceptAll walk is cached per apply */
  private recordTextCache: { view: View; byAddr: Map<Addr, string> } | null = null;

  private recordText(addr: Addr, view: View): string | null {
    if (view === acceptAll) {
      if (this.recordTextCache === null) {
        const byAddr = new Map<Addr, string>();
        for (const rec of projectParagraphs(this.walk, view)) if (rec.addr) byAddr.set(rec.addr, rec.text);
        this.recordTextCache = { view, byAddr };
      }
      return this.recordTextCache.byAddr.get(addr) ?? null;
    }
    for (const rec of projectParagraphs(this.walk, view)) if (rec.addr === addr) return rec.text;
    return null;
  }

  private findByAddrIn(addr: Addr): XElement | null {
    for (const p of this.walk.paragraphs) if (getAttr(p, "w14:paraId") === addr) return p;
    // inserted paragraphs are not in the cached walk list; rescan the body
    const fresh = walkBody(this.root);
    for (const p of fresh.paragraphs) {
      if (getAttr(p, "w14:paraId") === addr) {
        this.walk = fresh;
        return p;
      }
    }
    return null;
  }

  /**
   * Pure accept/reject of revisions, by changeIds or by author
   * (research/04 §3). Returns one result per decision with the change ids it
   * resolved. The model must be saved by the caller.
   */
  decide(decisions: RevisionDecision[]): DecisionResult[] {
    return decisions.map((d, index): DecisionResult => {
      const matched = [...allRevisions(this.root)].filter((rev) => {
        if ("changeIds" in d) {
          const id = getAttr(rev, "w:id");
          return id !== null && d.changeIds.includes(id);
        }
        return getAttr(rev, "w:author") === d.author;
      });
      const ids = matched.map((rev) => getAttr(rev, "w:id") ?? "").filter((s) => s !== "");
      const resolved = resolveRevisions(this.root, (rev) => matched.includes(rev), d.decision);
      if (resolved > 0) this.markDirty();
      return { index, decision: d.decision, changeIds: ids };
    });
  }

  /**
   * accept-with-caller-wording: the flagged span reads as the original with
   * the caller's text in place of our revision. Tree-level, exact: removes our
   * insertions, replaces the first deleted segment with the caller's wording,
   * and drops any further deleted segments.
   */
  acceptWithText(decision: Decision): number {
    if (decision.text === undefined) return 0;
    let touched = 0;
    const byParagraph = new Map<XElement, { ins: XElement[]; dels: XElement[]; markIns: XElement | null }>();
    for (const rev of allRevisions(this.root)) {
      const author = getAttr(rev, "w:author");
      if (author !== decision.author) continue;
      const paragraph = this.paragraphOf(rev);
      if (!paragraph) continue;
      const rec = byParagraph.get(paragraph) ?? { ins: [], dels: [], markIns: null };
      if (rev.name === "w:del") rec.dels.push(rev);
      else if (rev.name === "w:ins") {
        const pPr = kid(paragraph, "w:pPr");
        const rPr = pPr ? kid(pPr, "w:rPr") : undefined;
        if (rPr && rev.parent === rPr) rec.markIns = rev;
        else rec.ins.push(rev);
      }
      byParagraph.set(paragraph, rec);
    }
    // container-end insert pairs: an anchor paragraph with our tracked mark
    // whose next paragraph is the paired inserted half (isPairedInsertParagraph).
    // Only those paired paragraphs' w:pPrChange records belong to this edit.
    const paired = new Set<XElement>();
    for (const [para, rec] of byParagraph) {
      if (!rec.markIns) continue;
      const next = para.parent?.children[para.parent.children.indexOf(para) + 1];
      if (next && next.kind === "el" && next.name === "w:p" && isPairedInsertParagraph(next, decision.author)) {
        paired.add(next);
      }
    }
    for (const [para, rec] of byParagraph) {
      if (rec.markIns && rec.ins.length === 0) {
        // container-end insert form (Word's): the tracked mark sits on the
        // ANCHOR paragraph and the inserted paragraph follows it with a
        // w:pPrChange. Accepting the wording restores the anchor by just
        // removing the mark; the inserted paragraph is handled by its own
        // record below.
        const pPr = kid(para, "w:pPr");
        const rPrEl = pPr ? kid(pPr, "w:rPr") : undefined;
        if (rPrEl) {
          const mark = kids(rPrEl).find((c) => REVISION_NAMES.has(c.name));
          if (mark) removeChildEl(rPrEl, mark);
          if (rPrEl.children.length === 0 && pPr) removeChildEl(pPr, rPrEl);
        }
        touched++;
        continue;
      }
      if (rec.markIns) {
        // our inserted paragraph: strip the marks, put the caller's text in plain runs
        const firstIns = rec.ins[0];
        const rPr = firstIns ? cloneRunPr(firstIns) : null;
        let insertAt = para.children.length;
        for (const ins of rec.ins) {
          const i = para.children.indexOf(ins);
          if (i >= 0) {
            if (insertAt === para.children.length) insertAt = i;
            para.children.splice(i, 1);
          } else if (ins.parent) removeChildEl(ins.parent, ins);
        }
        const run = makeRun(rPr, textToRunChildren(decision.text!));
        para.children.splice(insertAt, 0, run);
        run.parent = para;
        const pPr = kid(para, "w:pPr");
        const rPrEl = pPr ? kid(pPr, "w:rPr") : undefined;
        if (rPrEl) {
          const mark = kids(rPrEl).find((c) => REVISION_NAMES.has(c.name));
          if (mark) removeChildEl(rPrEl, mark);
        }
        touched++;
        continue;
      }
      // insertion-only edits (empty-span replace): the caller's wording takes the
      // insertion's place, in the insertion's own parent (which may be a
      // hyperlink rather than the paragraph)
      if (rec.dels.length === 0 && rec.ins.length > 0) {
        const anchored = rec.ins[0]!;
        const parent = anchored.parent ?? para;
        const rPr = cloneRunPr(anchored);
        let insertAt = parent.children.indexOf(anchored);
        if (insertAt < 0) insertAt = parent.children.length;
        for (const ins of rec.ins) {
          const holder = ins.parent ?? para;
          const i = holder.children.indexOf(ins);
          if (i >= 0) holder.children.splice(i, 1);
        }
        const run = makeRun(rPr, textToRunChildren(decision.text!));
        parent.children.splice(Math.min(insertAt, parent.children.length), 0, run);
        run.parent = parent;
        // a paired container-end insert records the anchor's original
        // paragraph properties in a w:pPrChange; with the wording in place
        // that record is resolved too. An unrelated same-author property
        // change in any other paragraph is left alone.
        if (paired.has(para)) {
          const pPr = kid(para, "w:pPr");
          const change = pPr ? kids(pPr).find((c) => c.name === "w:pPrChange" && getAttr(c, "w:author") === decision.author) : undefined;
          if (change && pPr) removeChildEl(pPr, change);
        }
        touched++;
        continue;
      }
      // replace/delete: the caller's wording replaces the original span
      for (const ins of rec.ins) {
        const i = para.children.indexOf(ins);
        if (i >= 0) para.children.splice(i, 1);
        else if (ins.parent) removeChildEl(ins.parent, ins);
      }
      const parent = rec.dels[0]?.parent ?? para;
      let placed = false;
      for (const del of rec.dels) {
        if (!placed) {
          const i = parent.children.indexOf(del);
          if (i >= 0) {
            const firstRunInside = (() => {
              for (const c of del.children) if (c.kind === "el" && c.name === "w:r") return c;
              return null;
            })();
            const rPr = firstRunInside ? cloneRunPr(firstRunInside) : null;
            const run = makeRun(rPr, textToRunChildren(decision.text!));
            parent.children.splice(i, 1, run);
            run.parent = parent;
            placed = true;
            continue;
          }
        }
        const i = parent.children.indexOf(del);
        if (i >= 0) parent.children.splice(i, 1);
        else if (del.parent) removeChildEl(del.parent, del);
      }
      touched++;
    }
    if (touched > 0) this.markDirty();
    return touched;
  }

  private paragraphOf(rev: XElement): XElement | null {
    let cur: XElement | null = rev.parent;
    while (cur) {
      if (cur.name === "w:p") return cur;
      cur = cur.parent;
    }
    return null;
  }

  private markDirty(): void {
    this.dirty = true;
  }

  save(): Uint8Array {
    if (!this.dirty) return this.originalBytes;
    this.entries.set(this.mainPart, encodePart(serializeXml(this.tree)));
    return writePackage(this.entries);
  }
}

function removeChildEl(parent: XElement, child: XElement): void {
  const i = parent.children.indexOf(child);
  if (i >= 0) parent.children.splice(i, 1);
}

function* allRevisions(root: XElement): Generator<XElement> {
  const stack: XElement[] = [root];
  while (stack.length) {
    const cur = stack.pop()!;
    for (const c of cur.children) {
      if (c.kind !== "el") continue;
      if (REVISION_NAMES.has(c.name) || c.name.endsWith("PrChange")) yield c;
      stack.push(c);
    }
  }
}

// —————————————————————————————————————————————————————————————— numbering

interface NumLevel {
  fmt: string;
  text: string;
  start: number;
}
interface NumberingModel {
  labelFor(p: XElement): string | undefined;
}
const emptyNumbering: NumberingModel = { labelFor: () => undefined };

function parseNumbering(xml: string): NumberingModel {
  const tree = parseXml(xml);
  const root = kid(treeAsEl(tree), "w:numbering");
  if (!root) return emptyNumbering;
  const abstracts = new Map<string, Map<string, NumLevel>>();
  for (const a of kids(root, "w:abstractNum")) {
    const id = getAttr(a, "w:abstractNumId");
    if (id === null) continue;
    const levels = new Map<string, NumLevel>();
    for (const lvl of kids(a, "w:lvl")) {
      const ilvl = getAttr(lvl, "w:ilvl");
      if (ilvl === null) continue;
      const fmt = getAttr(kid(lvl, "w:numFmt") ?? (el("x") as XElement), "w:val") ?? "decimal";
      const text = getAttr(kid(lvl, "w:lvlText") ?? (el("x") as XElement), "w:val") ?? "";
      const start = parseInt(getAttr(kid(lvl, "w:start") ?? (el("x") as XElement), "w:val") ?? "1", 10) || 1;
      levels.set(ilvl, { fmt, text, start });
    }
    abstracts.set(id, levels);
  }
  const numToAbstract = new Map<string, string>();
  for (const n of kids(root, "w:num")) {
    const numId = getAttr(n, "w:numId");
    const abs = getAttr(kid(n, "w:abstractNumId") ?? (el("x") as XElement), "w:val");
    if (numId && abs) numToAbstract.set(numId, abs);
  }

  const counters = new Map<string, number[]>();
  return {
    labelFor(p: XElement): string | undefined {
      const pPr = kid(p, "w:pPr");
      const numPr = pPr ? kid(pPr, "w:numPr") : undefined;
      if (!numPr) return undefined;
      const numId = getAttr(kid(numPr, "w:numId") ?? (el("x") as XElement), "w:val") ?? "0";
      const ilvl = getAttr(kid(numPr, "w:ilvl") ?? (el("x") as XElement), "w:val") ?? "0";
      const abs = numToAbstract.get(numId);
      const levels = abs ? abstracts.get(abs) : undefined;
      const lvl = levels?.get(ilvl);
      if (!lvl) return undefined;
      const key = `${numId}`;
      const arr = counters.get(key) ?? [];
      const level = parseInt(ilvl, 10) || 0;
      arr[level] = (arr[level] ?? lvl.start - 1) + 1;
      for (let i = level + 1; i < arr.length; i++) arr[i] = lvl.start - 1;
      counters.set(key, arr);
      if (lvl.fmt === "bullet" || lvl.fmt === "none") return lvl.fmt === "none" ? "" : lvl.text;
      return lvl.text.replace(/%\d/g, (m) => {
        const idx = parseInt(m.slice(1), 10);
        const v = arr[idx] ?? lvl.start;
        return String(v);
      });
    },
  };
}

// —————————————————————————————————————————————————————————————————— public

export function openDocument(bytes: Uint8Array): Document {
  return Document.open(bytes);
}

export { normalizeDocx };
