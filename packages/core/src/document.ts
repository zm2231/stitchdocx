// Document: the in-memory engine. The XML tree is the document; the model is
// an index over it. Save re-serializes only what changed — untouched
// paragraphs and untouched package entries are byte-identical by construction.
import {
  cloneNode,
  decodePart,
  el,
  encodePart,
  getAttr,
  isPending,
  kid,
  kids,
  mainPartName,
  materialize,
  pendingElements,
  pendingSource,
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
import { parseMainPart, walkBody, type BodyWalk } from "./body.js";
import { acceptAll, attrVal, paragraphText, projectParagraphs, REVISION_NAMES, type ProjectionParagraph, type View } from "./projection.js";
import { buildSpanMap } from "./spanmap.js";
import { executeApply, splitAt, type ApplyDeps } from "./ops.js";
import { mayHoldRevisionId, ParaIdAllocator, resolveRevisions, RevisionIdAllocator, textToRunChildren, makeRun, REVISION_ID_NAMES } from "./revisions.js";
import { normalizeDocx } from "./normalize.js";
import { StyleSheet } from "./styles.js";
import { NumberingModel, parseNumberingModel } from "./numbering.js";
import { maxRevisionIdAcrossParts, paraIdsAcrossParts, wordprocessingPartNames } from "./parts.js";
import { mainPartSource, parsePart } from "./limits.js";
import type { Addr, Decision, DecisionResult, Op, Paragraph, Receipt } from "./types.js";

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
  readonly styles: StyleSheet;

  private constructor(bytes: Uint8Array) {
    this.originalBytes = bytes;
    this.reader = openPackage(bytes);
    this.mainPart = mainPartName(this.reader);
    // paragraphs stay source until an edit reaches them; read-only passes
    // parse each one transiently, and save copies the untouched ones verbatim
    const source = mainPartSource(this.reader, this.mainPart);
    ({ tree: this.tree, root: this.root } = parseMainPart(source));
    this.walk = walkBody(this.root);
    this.styles = StyleSheet.parse(this.partText(relatedPart(this.reader, this.mainPart, "styles") ?? "word/styles.xml"));
    this.paraByAddr = new Map();
    for (const p of this.walk.paragraphs) {
      const addr = getAttr(p, "w14:paraId");
      if (addr !== null) this.paraByAddr.set(addr, p);
    }
    // revision ids are unique across every part; seed above the largest id on
    // any revision-type element anywhere
    this.ids = new RevisionIdAllocator(maxRevisionIdAcrossParts(this.reader, this.mainPart, source));
    this.paraIds = new ParaIdAllocator(paraIdsAcrossParts(this.reader, this.mainPart, source));
    // every entry stays compressed; save replaces the main part
    this.entries = this.reader.readAll();
  }

  static open(bytes: Uint8Array): Document {
    return new Document(bytes);
  }

  /** projection under a view; addresses are w14:paraId (present after normalize).
   *  Numbering labels are computed per call — no counter state survives. */
  projection(view: View = acceptAll): Paragraph[] {
    this.assertUsable();
    const model = this.numberingModel();
    const counters = new Map<string, (number | undefined)[]>();
    return projectParagraphs(this.walk, view, this.styles).map((p) => this.toPublicParagraph(p, model, counters));
  }

  /** raw projection records (engine-internal shape with merged nodes) */
  projectionRecords(view: View = acceptAll): ProjectionParagraph[] {
    this.assertUsable();
    return projectParagraphs(this.walk, view, this.styles);
  }

  private toPublicParagraph(p: ProjectionParagraph, model: NumberingModel, counters: Map<string, (number | undefined)[]>): Paragraph {
    const kind: Paragraph["kind"] = p.container === "cell" ? "cell" : model.numberedFor(p) ? "listItem" : /(^|[^a-z])heading|title/i.test(p.style ?? "") ? "heading" : "body";
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
    if (model.numberedFor(p)) {
      const label = model.labelFor(p, counters);
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
    return this.reader.xml(name);
  }

  apply(ops: Op[], opts: { date: string }): Receipt[] {
    this.assertUsable();
    if (!isIsoDate(opts.date)) throw new Error(`date must be an ISO 8601 timestamp, got ${JSON.stringify(opts.date)}`);
    this.recordTextCache = null;
    // the addressable surface is the projection's listing: unlisted addresses
    // and merged view records refuse before any span check runs
    const records = projectParagraphs(this.walk, acceptAll, this.styles);
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
      styles: this.styles,
      fieldEntry: (para) => {
        const state = this.walk.fieldStates.get(para);
        return state ? [...state] : [];
      },
    };
    // refusals come back as receipts; anything thrown (a node budget spent
    // mid-batch, say) may leave the tree partly edited
    let receipts: Receipt[];
    try {
      receipts = executeApply(deps, ops);
    } catch (e) {
      this.broken = e instanceof Error ? e.message : String(e);
      throw e;
    }
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
        for (const rec of projectParagraphs(this.walk, view, this.styles)) if (rec.addr) byAddr.set(rec.addr, rec.text);
        this.recordTextCache = { view, byAddr };
      }
      return this.recordTextCache.byAddr.get(addr) ?? null;
    }
    return projectParagraphs(this.walk, view, this.styles, new Set([addr]))[0]?.text ?? null;
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
   * Resolve each decision's revisions, named by the receipt's changeIds.
   * Throws, naming the decision, when an id is not a revision in the
   * document, when two decisions name the same id, or when accept-with-wording
   * cannot place the member's text. Id checks run before anything changes;
   * a failure after that leaves this Document partly decided, so every later
   * call on it throws.
   */
  decide(decisions: Decision[]): DecisionResult[] {
    this.assertUsable();
    // only paragraphs whose source names one of the ids are parsed
    const holds = mayHoldRevisionId(decisions.flatMap((d) => d.changeIds));
    for (const p of [...pendingElements(this.root)]) if (holds(pendingSource(p)!)) materialize(p);
    const byId = new Map<string, XElement[]>();
    for (const rev of allRevisions(this.root)) {
      const id = getAttr(rev, "w:id");
      if (id !== null) byId.set(id, [...(byId.get(id) ?? []), rev]);
    }
    const claimed = new Map<string, number>();
    decisions.forEach((d, index) => {
      if (d.changeIds.length === 0) throw new Error(`decision ${index} (${d.ref}) names no changeIds`);
      if (d.text !== undefined && d.decision !== "accept") throw new Error(`decision ${index} (${d.ref}) gives text with a reject`);
      for (const id of d.changeIds) {
        if (!byId.has(id)) throw new Error(`decision ${index} (${d.ref}) names change ${id}, which is not a revision in the document`);
        const prior = claimed.get(id);
        if (prior !== undefined) throw new Error(`decisions ${prior} and ${index} both name change ${id}`);
        claimed.set(id, index);
      }
    });
    // from here decisions mutate the tree; a failure leaves it partly decided
    let results: DecisionResult[];
    try {
      results = decisions.map((d, index): DecisionResult => {
        const matched = new Set(d.changeIds.flatMap((id) => byId.get(id) ?? []));
        if (d.text === undefined) resolveRevisions(this.root, (rev) => matched.has(rev), d.decision);
        else this.acceptWithText(d, index, matched);
        return { index, ref: d.ref, decision: d.decision, changeIds: [...d.changeIds] };
      });
    } catch (e) {
      this.broken = e instanceof Error ? e.message : String(e);
      throw e;
    }
    if (results.length > 0) this.markDirty();
    return results;
  }

  /**
   * Accept every tracked change in every story part: the body, headers,
   * footers, footnotes, endnotes and comments. Afterwards the document holds
   * no tracked change of any kind, and the body reads as its accept projection
   * did before. Cell, numbering and table-grid changes, which decide() cannot
   * name, are accepted too: a deleted cell goes, the others keep their current
   * form.
   */
  acceptAllChanges(): void {
    this.assertUsable();
    for (const p of [...pendingElements(this.root)]) if (TRACKED.test(pendingSource(p)!)) materialize(p);
    let accepted = acceptEverything(this.root);
    this.walk = walkBody(this.root);
    this.paraByAddr = new Map();
    for (const p of this.walk.paragraphs) {
      const addr = getAttr(p, "w14:paraId");
      if (addr !== null) this.paraByAddr.set(addr, p);
    }
    for (const part of wordprocessingPartNames(this.reader, this.mainPart).slice(1)) {
      const source = this.reader.xml(part);
      if (source === undefined || !TRACKED.test(source)) continue;
      const tree = parsePart(source);
      const partRoot = kid(treeAsEl(tree));
      if (!partRoot) continue;
      if (acceptEverything(partRoot) === 0) continue;
      this.entries.set(part, encodePart(serializeXml(tree)));
      accepted++;
    }
    if (accepted > 0) this.markDirty();
  }

  private broken: string | null = null;

  private assertUsable(): void {
    if (this.broken !== null) throw new Error(`document is unusable after a failed edit: ${this.broken}`);
  }

  /**
   * Accept with the member's wording: the document reads as the original with
   * the flagged span (or the inserted paragraph) replaced by the text, untracked.
   */
  private acceptWithText(d: Decision, index: number, matched: ReadonlySet<XElement>): void {
    const isMark = (rev: XElement) => rev.parent?.name === "w:rPr" && rev.parent.parent?.name === "w:pPr";
    if ([...matched].some((rev) => rev.name === "w:ins" && isMark(rev))) {
      this.acceptInsertWithText(d.text!, matched);
      return;
    }
    if (d.expectText === undefined) throw new Error(`decision ${index} (${d.ref}) accepts with wording but gives no expectText`);
    const paragraphs = new Set([...matched].map((rev) => this.paragraphOf(rev)));
    if (paragraphs.size !== 1 || paragraphs.has(null)) throw new Error(`decision ${index} (${d.ref}) accepts with wording but its changes are not in one paragraph`);
    const para = [...paragraphs][0]!;

    // an empty-span replace is a pure insertion: the wording takes its place,
    // inside whatever holds it (a hyperlink, say)
    if (d.expectText === "") {
      const ins = [...matched].filter((rev) => rev.name === "w:ins");
      if (ins.length !== matched.size) throw new Error(`decision ${index} (${d.ref}): expectText is empty but the change deletes text`);
      this.replaceInsertions(ins, d.text!);
      return;
    }

    // where this decision's changes sit, in the paragraph's text with them
    // rejected: a deletion covers its characters, an insertion is a point
    const ids = new Set([...matched].map((m) => getAttr(m, "w:id")));
    const showOurs: View = (rev) => (ids.has(rev.id) && (rev.name === "w:del" || rev.name === "w:moveFrom") ? "reject" : "accept");
    const shown = buildSpanMap(para, showOurs, this.styles);
    let pos = 0;
    let lo = Infinity;
    let hi = -Infinity;
    for (const pc of shown.pieces) {
      const rev = shown.runInfo.get(pc.run)?.rev ?? null;
      const len = pc.end - pc.start;
      if (rev && matched.has(rev) && (rev.name === "w:ins" || rev.name === "w:moveTo")) {
        lo = Math.min(lo, pos);
        hi = Math.max(hi, pos);
        continue;
      }
      if (rev && matched.has(rev)) {
        lo = Math.min(lo, pos);
        hi = Math.max(hi, pos + len);
      }
      pos += len;
    }
    if (lo === Infinity) throw new Error(`decision ${index} (${d.ref}) accepts with wording but none of its changes is visible text`);

    resolveRevisions(this.root, (rev) => matched.has(rev), "reject");

    const text = paragraphText(para, acceptAll, this.styles);
    const want = d.expectText;
    const starts: number[] = [];
    for (let at = text.indexOf(want); at >= 0; at = text.indexOf(want, at + 1)) {
      if (at <= lo && at + want.length >= hi) starts.push(at);
    }
    if (starts.length !== 1) {
      throw new Error(`decision ${index} (${d.ref}): expectText ${JSON.stringify(want)} covers its changes ${starts.length === 0 ? "nowhere" : `${starts.length} times`} in the paragraph`);
    }
    this.replacePlain(para, starts[0]!, starts[0]! + want.length, d.text!, index, d.ref);
  }

  /** replace [start,end) of the paragraph's text with `text` as one plain run
   *  carrying the formatting of the span's first run */
  private replacePlain(para: XElement, start: number, end: number, text: string, index: number, ref: string): void {
    const take = () => this.ids.take();
    splitAt(buildSpanMap(para, acceptAll, this.styles), start, take);
    splitAt(buildSpanMap(para, acceptAll, this.styles), end, take);
    const map = buildSpanMap(para, acceptAll, this.styles);
    const inside = (pc: (typeof map.pieces)[number]) =>
      pc.end > pc.start ? pc.start >= start && pc.end <= end : pc.start > start && pc.start < end;
    const runs: XElement[] = [];
    for (const pc of map.pieces) if (inside(pc) && !runs.includes(pc.run)) runs.push(pc.run);
    for (const r of runs) {
      if (map.runInfo.get(r)?.rev) throw new Error(`decision ${index} (${ref}): the flagged span holds another revision`);
    }
    if (new Set(runs.map((r) => r.parent)).size > 1) throw new Error(`decision ${index} (${ref}): the flagged span crosses a structure boundary`);
    if (runs.length === 0) throw new Error(`decision ${index} (${ref}): the flagged span holds no runs`);
    const parent = runs[0]!.parent!;
    const at = parent.children.indexOf(runs[0]!);
    const rPr = this.styles.insertedRunPr(runs[0]!, para);
    for (const r of runs) parent.children.splice(parent.children.indexOf(r), 1);
    if (text === "") return;
    const run = makeRun(rPr, textToRunChildren(text));
    parent.children.splice(at, 0, run);
    run.parent = parent;
  }

  /** put `text` as one plain run where the first insertion sits, with its
   *  formatting, and drop the insertions */
  private replaceInsertions(ins: XElement[], text: string): void {
    const first = ins[0]!;
    const parent = first.parent!;
    const at = parent.children.indexOf(first);
    const runPr = this.styles.insertedRunPr(kid(first, "w:r") ?? null, this.paragraphOf(first)!);
    for (const i of ins) if (i.parent) removeChildEl(i.parent, i);
    const run = makeRun(runPr, textToRunChildren(text));
    parent.children.splice(Math.min(at, parent.children.length), 0, run);
    run.parent = parent;
  }

  /** accept-with-wording for an inserted paragraph: the inserted paragraph
   *  keeps its place and reads `text`, untracked */
  private acceptInsertWithText(text: string, matched: ReadonlySet<XElement>): void {
    const byParagraph = new Map<XElement, { ins: XElement[]; markIns: XElement | null; pPrChange: XElement | null }>();
    for (const rev of matched) {
      const para = this.paragraphOf(rev);
      if (!para) continue;
      const rec = byParagraph.get(para) ?? { ins: [], markIns: null, pPrChange: null };
      if (rev.name === "w:ins" && rev.parent?.name === "w:rPr" && rev.parent.parent?.name === "w:pPr") rec.markIns = rev;
      else if (rev.name === "w:ins") rec.ins.push(rev);
      else if (rev.name === "w:pPrChange") rec.pPrChange = rev;
      byParagraph.set(para, rec);
    }
    for (const [para, rec] of byParagraph) {
      // the anchor of Word's container-end form carries only our mark
      if (rec.markIns) {
        const rPr = rec.markIns.parent!;
        removeChildEl(rPr, rec.markIns);
        const pPr = rPr.parent;
        if (rPr.children.length === 0 && pPr) removeChildEl(pPr, rPr);
      }
      if (rec.ins.length > 0) this.replaceInsertions(rec.ins, text);
      if (rec.pPrChange?.parent) removeChildEl(rec.pPrChange.parent, rec.pPrChange);
    }
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
    this.assertUsable();
    if (!this.dirty) return this.originalBytes;
    this.entries.set(this.mainPart, encodePart(serializeXml(this.tree)));
    return writePackage(this.entries);
  }
}

function removeChildEl(parent: XElement, child: XElement): void {
  const i = parent.children.indexOf(child);
  if (i >= 0) parent.children.splice(i, 1);
}

/** source that may hold a tracked change of any kind */
const TRACKED = new RegExp(`<(?:${[...REVISION_ID_NAMES].join("|")})[\\s/>]`);

/** accept every tracked change under `root`; returns how many there were */
function acceptEverything(root: XElement): number {
  const resolved = resolveRevisions(root, () => true, "accept");
  const leftovers: XElement[] = [];
  const visit = (node: XElement): void => {
    for (const c of node.children) {
      if (c.kind !== "el") continue;
      if (REVISION_ID_NAMES.has(c.name)) leftovers.push(c);
      else if (isPending(c)) continue; // its source holds no tracked change
      else visit(c);
    }
  };
  visit(root);
  for (const rev of leftovers) {
    if (rev.name === "w:cellDel") {
      const tc = rev.parent?.parent;
      if (tc?.name === "w:tc") removeChildEl(tc.parent!, tc);
    }
    if (rev.parent) removeChildEl(rev.parent, rev);
  }
  return resolved + leftovers.length;
}

/** revision elements outside unparsed paragraphs */
function* allRevisions(root: XElement): Generator<XElement> {
  const stack: XElement[] = [root];
  while (stack.length) {
    const cur = stack.pop()!;
    if (isPending(cur)) continue;
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

/** Resolve the member's decisions on a redline. All or nothing: a failing
 *  decision throws and no document is produced. */
export function finalizeDocx(bytes: Uint8Array, decisions: Decision[]): { docx: Uint8Array; results: DecisionResult[] } {
  const doc = Document.open(bytes);
  const results = doc.decide(decisions);
  return { docx: doc.save(), results };
}

/** The document with every tracked change accepted, in every story part;
 *  the input itself when it holds none. */
export function acceptAllDocx(bytes: Uint8Array): Uint8Array {
  const doc = Document.open(bytes);
  doc.acceptAllChanges();
  return doc.save();
}

export { normalizeDocx };
