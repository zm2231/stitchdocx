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
  openPackage,
  relatedPart,
  setAttr,
  writePackage,
  parseXml,
  serializeXml,
  type EntryData,
  type Entries,
  type XElement,
} from "@stitchdocx/package-io";
import { walkBody, type BodyWalk } from "./body.js";
import { acceptAll, attrVal, paragraphText, projectParagraphs, REVISION_NAMES, type ProjectionParagraph, type View } from "./projection.js";
import { buildSpanMap } from "./spanmap.js";
import { descendantsOf, executeApply, splitAt, type ApplyDeps } from "./ops.js";
import { cloneRunPr, isPairedInsertParagraph, ParaIdAllocator, renameEl, resolveRevisions, RevisionIdAllocator, textToRunChildren, makeRun, unwrap } from "./revisions.js";
import { normalizeDocx } from "./normalize.js";
import { NumberingModel, parseNumberingModel } from "./numbering.js";
import { otherPartRoots, paraIdsAcrossParts } from "./parts.js";
import type { Addr, Decision, DecisionResult, Op, Paragraph, Receipt, RevisionDecision } from "./types.js";

/** the parsed document root is an XDoc; kid() works on it identically */
const treeAsEl = (d: unknown) => d as XElement;

/** RFC 3339 timestamp — the w:date form Word writes */
const ISO_DATE = /^\d{4}-\d{2}-\d{2}[Tt]\d{2}:\d{2}:\d{2}(\.\d+)?([Zz]|[+-]\d{2}:\d{2})$/;

export function isIsoDate(value: unknown): value is string {
  return typeof value === "string" && ISO_DATE.test(value) && !Number.isNaN(Date.parse(value));
}

export class Document {
  private reader: ReturnType<typeof openPackage>;
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
    this.reader = openPackage(bytes);
    this.mainPart = mainPartName(this.reader);
    const docXml = this.reader.part(this.mainPart);
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
    // revision ids are unique across every part; seed above the largest id on
    // any revision-type element anywhere (the main part's tree is already parsed)
    this.ids = new RevisionIdAllocator([this.root, ...otherPartRoots(this.reader, this.mainPart)]);
    this.paraIds = new ParaIdAllocator(paraIdsAcrossParts(this.reader, this.mainPart, this.walk.paragraphs));
    // only the main part comes inflated; everything else stays raw until the
    // writer needs it or a command reads it
    this.entries = this.reader.readAll((name) => name === this.mainPart);
  }

  static open(bytes: Uint8Array): Document {
    return new Document(bytes);
  }

  /** projection under a view; addresses are w14:paraId (present after normalize).
   *  Numbering labels are computed per call — no counter state survives. */
  projection(view: View = acceptAll): Paragraph[] {
    const model = this.numberingModel();
    const counters = new Map<string, (number | undefined)[]>();
    return projectParagraphs(this.walk, view).map((p) => this.toPublicParagraph(p, model, counters));
  }

  /** raw projection records (engine-internal shape with merged nodes) */
  projectionRecords(view: View = acceptAll): ProjectionParagraph[] {
    return projectParagraphs(this.walk, view);
  }

  private toPublicParagraph(p: ProjectionParagraph, model: NumberingModel, counters: Map<string, (number | undefined)[]>): Paragraph {
    const kind: Paragraph["kind"] = p.container === "cell" ? "cell" : model.numberedFor(p.node) ? "listItem" : /(^|[^a-z])heading|title/i.test(p.style ?? "") ? "heading" : "body";
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
    if (model.numberedFor(p.node)) {
      const label = model.labelFor(p.node, counters);
      if (label !== undefined) out.numbering = label;
    }
    return out;
  }

  /** numbering + styles model, read through the package relationships */
  private numberingModel(): NumberingModel {
    if (!this.numbering) {
      const numberingXml = this.partText(relatedPart(this.reader, this.mainPart, "numbering") ?? "word/numbering.xml");
      const stylesXml = this.partText(relatedPart(this.reader, this.mainPart, "styles") ?? "word/styles.xml");
      this.numbering = parseNumberingModel(numberingXml, stylesXml);
    }
    return this.numbering;
  }

  private partText(name: string): string | undefined {
    const bytes = this.reader.part(name);
    return bytes ? decodePart(bytes) : undefined;
  }

  apply(ops: Op[], opts: { date: string }): Receipt[] {
    if (!isIsoDate(opts.date)) throw new Error(`date must be an ISO 8601 timestamp, got ${JSON.stringify(opts.date)}`);
    this.recordTextCache = null;
    // the addressable surface is the projection's listing: unlisted addresses
    // and merged view records refuse before any span check runs
    const records = projectParagraphs(this.walk, acceptAll);
    const listed = new Set<Addr>();
    const editable = new Set<Addr>();
    for (const rec of records) {
      if (!rec.addr) continue;
      listed.add(rec.addr);
      if (rec.removedMark === null && !rec.carried) editable.add(rec.addr);
    }
    const deps: ApplyDeps = {
      paraByAddr: this.paraByAddr,
      ids: this.ids,
      paraIds: this.paraIds,
      date: opts.date,
      // ownership is scoped to THIS apply batch: revisions from earlier calls
      // are pre-existing tracked changes and must refuse, not reposition
      ownedRevisions: new Set<XElement>(),
      recordText: (addr, view) => this.recordText(addr, view),
      addressable: (addr) => (!listed.has(addr) ? "unlisted" : editable.has(addr) ? "ok" : "merged"),
      fieldEntry: (para) => {
        const state = this.walk.fieldStates.get(para);
        return state ? [...state] : [];
      },
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
    // a decision names either the receipt's change ids or (legacy) an author
    const changeIdSet = decision.changeIds ? new Set(decision.changeIds) : null;
    const matched = [...allRevisions(this.root)].filter((rev) =>
      changeIdSet ? changeIdSet.has(getAttr(rev, "w:id") ?? "") : getAttr(rev, "w:author") === decision.author,
    );
    const author = decision.author ?? (matched[0] ? getAttr(matched[0], "w:author") : null) ?? "";
    for (const rev of matched) {
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
      if (next && next.kind === "el" && next.name === "w:p" && isPairedInsertParagraph(next, author)) {
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
          const change = pPr ? kids(pPr).find((c) => c.name === "w:pPrChange" && getAttr(c, "w:author") === author) : undefined;
          if (change && pPr) removeChildEl(pPr, change);
        }
        touched++;
        continue;
      }
      // replace/delete: the caller's wording replaces the whole span the op
      // changed — every deleted segment, the unchanged words between them
      // included — with the member's text as one run. The span is recovered
      // from the paragraph itself: rejecting this decision's revisions reads
      // the original, accepting them reads the replacement, and what remains
      // after trimming the common head and tail is exactly the region the op
      // changed.
      const idSet = new Set(matched.map((rev) => getAttr(rev, "w:id") ?? ""));
      const rejectThis: View = (rev) => (rev.id !== null && idSet.has(rev.id) ? "reject" : "accept");
      const originalText = paragraphText(para, rejectThis);
      const replacedText = paragraphText(para, acceptAll);
      let start = 0;
      while (start < originalText.length && start < replacedText.length && originalText[start] === replacedText[start]) start++;
      let end = originalText.length;
      let replEnd = replacedText.length;
      while (end > start && replEnd > start && originalText[end - 1] === replacedText[replEnd - 1]) {
        end--;
        replEnd--;
      }
      for (const ins of rec.ins) {
        const i = para.children.indexOf(ins);
        if (i >= 0) para.children.splice(i, 1);
        else if (ins.parent) removeChildEl(ins.parent, ins);
      }
      for (const del of rec.dels) {
        for (const d of descendantsOf(del)) if (d.name === "w:delText") renameEl(d, "w:t");
        unwrap(del);
      }
      if (end > start) {
        // the paragraph now reads exactly `originalText`; strike [start,end)
        // plainly (the wording is accepted, not tracked) and drop the
        // member's text in as one run
        let map = buildSpanMap(para, acceptAll);
        splitAt(map, start, () => this.ids.take());
        map = buildSpanMap(para, acceptAll);
        splitAt(map, end, () => this.ids.take());
        map = buildSpanMap(para, acceptAll);
        const covered = map.pieces.filter((pc) => pc.start >= start && pc.end <= end);
        const runs: XElement[] = [];
        for (const pc of covered) if (!runs.includes(pc.run)) runs.push(pc.run);
        if (runs.length > 0 && new Set(runs.map((r) => r.parent)).size === 1) {
          const parent = runs[0]!.parent!;
          const firstRunPr = cloneRunPr(runs[0]!);
          const at = parent.children.indexOf(runs[0]!);
          for (const r of runs) {
            const i = parent.children.indexOf(r);
            if (i >= 0) parent.children.splice(i, 1);
          }
          const run = makeRun(firstRunPr, textToRunChildren(decision.text!));
          parent.children.splice(Math.min(at, parent.children.length), 0, run);
          run.parent = parent;
        }
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

// —————————————————————————————————————————————————————————————————— public

export function openDocument(bytes: Uint8Array): Document {
  return Document.open(bytes);
}

export { normalizeDocx };
