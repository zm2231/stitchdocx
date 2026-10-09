// apply: validate every op against the pristine document (refusals are
// all-or-nothing and never stop later ops), then mutate per paragraph
// right-to-left so all offsets — which the contract defines in the paragraph's
// original text coordinates — stay valid without adjustment.
import {
  cloneNode,
  el,
  kid,
  setAttr,
  textNode,
  type XElement,
  type XNode,
} from "@stitchdocx/package-io";
import { buildSpanMap, checkSpan, outsideLinks, pieceAt, type SpanMap } from "./spanmap.js";
import { attrVal, markRevision, paragraphText, textContent, wordText, type View } from "./projection.js";
import {
  makeRun,
  newRevisionEl,
  renameEl,
  reidRevisionRecords,
  stripRevisionRecords,
  RevisionIdAllocator,
  ParaIdAllocator,
  textToRunChildren,
} from "./revisions.js";
import { readableWordEdits, wordEdits, type WordEdit } from "./worddiff.js";
import type { StyleSheet } from "./styles.js";
import type { Addr, Op, Receipt } from "./types.js";

export interface ApplyDeps {
  paraByAddr: Map<Addr, XElement>;
  ids: RevisionIdAllocator;
  paraIds: ParaIdAllocator;
  date: string;
  /** revision elements created by THIS apply call — the only ones a later op
   *  may treat as its own (sibling placement); everything else is untouchable */
  ownedRevisions: Set<XElement>;
  /** the addressed paragraph's text in projection-record space under `view` —
   *  records merge removed-mark content, so a record can read differently
   *  than the physical paragraph node; receipts report in record space.
   *  Returns null when the address has no record (caller falls back). */
  recordText(addr: Addr, view: View): string | null;
  /** whether the projection lists `addr` as an editable record ("ok"), lists
   *  it only as a merged/uneditable view ("merged"), or never lists it
   *  ("unlisted") — ops are only valid against listed, editable records */
  addressable(addr: Addr): "ok" | "unlisted" | "merged";
  /** the paragraph's complex-field cursor state at entry (shared with the
   *  whole-story walk); callers receive a fresh copy */
  fieldEntry(para: XElement): string[];
  styles: StyleSheet;
}

const acceptAllView: View = () => "accept";

interface Plan {
  op: Op;
  index: number;
  status: "applied" | "refused";
  reason?: Receipt["reason"];
  result?: { changeIds: string[]; insertedAddr?: string };
  /** the word-granularity edits validation accepted, applied as validated */
  edits?: WordEdit[];
}

export function executeApply(deps: ApplyDeps, ops: Op[]): Receipt[] {
  // — phase 1: validate in request order against the pristine document —
  const plans: Plan[] = ops.map((op, index) => ({ op, index, status: "applied" }));
  const maps = new Map<Addr, SpanMap>();
  const priorApplied = new Map<Addr, Array<{ start: number; end: number }>>();
  const ownedRevisions = deps.ownedRevisions;
  for (const plan of plans) {
    const op = plan.op;
    const para = deps.paraByAddr.get(op.addr);
    if (!para) {
      refuse(plan, "addr_not_found");
      continue;
    }
    // ops are authored against the projection's listing: an address that is
    // not a listed, editable record cannot be operated on (the listing's
    // coordinates would not match the paragraph's own span map)
    const addrState = deps.addressable(op.addr);
    if (addrState === "unlisted") {
      refuse(plan, "addr_not_found");
      continue;
    }
    if (addrState === "merged") {
      refuse(plan, "paragraph_not_editable");
      continue;
    }
    if (op.op === "insertParagraphAfter") {
      if (!op.text || op.text.trim() === "") refuse(plan, "empty_text");
      else if (op.formatFrom !== undefined && (!deps.paraByAddr.has(op.formatFrom) || deps.addressable(op.formatFrom) === "unlisted")) refuse(plan, "addr_not_found");
      else if (op.formatFrom !== undefined && deps.addressable(op.formatFrom) === "merged") refuse(plan, "paragraph_not_editable");
      else if (op.scope === "table" && !hasAncestor(para, "w:tbl")) refuse(plan, "addr_not_found");
      else if (op.scope !== "table" && inPlainTextControl(para)) refuse(plan, "insert_in_plain_text_control");
      else if (op.scope !== "table" && nextBlockSibling(para)?.name !== "w:p" && markRevision(para)) {
        // at a container end the insert must track the ANCHOR's paragraph
        // mark (Word's form); an anchor whose mark is already tracked by any
        // author cannot carry ours, and Word could not reject the result
        refuse(plan, "paragraph_not_editable");
      }
      continue;
    }
    const map = getMap(maps, para, deps);
    if (
      !Number.isInteger(op.start) ||
      !Number.isInteger(op.end) ||
      op.start < 0 ||
      op.end < op.start ||
      op.end > map.text.length
    ) {
      refuse(plan, "span_out_of_range");
      continue;
    }
    // an offset that falls between the halves of a surrogate pair would
    // silently destroy the character (the parser strips lone halves)
    if (splitsSurrogate(map.text, op.start) || splitsSurrogate(map.text, op.end)) {
      refuse(plan, "span_out_of_range");
      continue;
    }
    if (map.text.slice(op.start, op.end) !== op.expectText) {
      refuse(plan, "span_text_mismatch");
      continue;
    }
    if (op.op === "replace" && (!op.text || op.text.trim() === "")) {
      refuse(plan, "empty_text");
      continue;
    }
    const { issue, edits } = spanIssue(map, op, priorApplied.get(op.addr) ?? [], ownedRevisions);
    plan.edits = edits;
    if (issue) {
      refuse(plan, issue.reason);
      continue;
    }
    const list = priorApplied.get(op.addr) ?? [];
    list.push({ start: op.start, end: op.end });
    priorApplied.set(op.addr, list);
  }

  // — phase 2: mutate. span ops per paragraph right-to-left; inserts in order —
  const changeIdsByIndex = new Map<number, string[]>();
  const insertedAddrByIndex = new Map<number, string>();
  const spanPlansByAddr = new Map<Addr, Plan[]>();
  for (const plan of plans) {
    if (plan.status !== "applied") continue;
    if (plan.op.op === "replace" || plan.op.op === "delete") {
      const list = spanPlansByAddr.get(plan.op.addr) ?? [];
      list.push(plan);
      spanPlansByAddr.set(plan.op.addr, list);
    }
  }
  for (const [addr, list] of spanPlansByAddr) {
    const para = deps.paraByAddr.get(addr)!;
    list.sort((a, b) => {
      const pa = a.op as SpanOp;
      const pb = b.op as SpanOp;
      return pb.start - pa.start;
    });
    for (const plan of list) {
      changeIdsByIndex.set(plan.index, applySpanOp(deps, para, plan.op as SpanOp, plan.edits));
    }
  }
  const insertChains = new Map<XElement, XElement>();
  const insertPlans = plans.filter((p) => p.status === "applied" && p.op.op === "insertParagraphAfter");
  // decided up front against the pristine tree: whether a paragraph follows
  // the chain decides the markup form (inside the loop the tree is already
  // mutating)
  const chainFollowsParagraph = new Map<XElement, boolean>();
  for (const plan of insertPlans) {
    const op = plan.op as Extract<Op, { op: "insertParagraphAfter" }>;
    const key = chainKeyOf(op, deps.paraByAddr.get(op.addr)!);
    if (!chainFollowsParagraph.has(key)) chainFollowsParagraph.set(key, nextBlockSibling(key)?.name === "w:p");
  }
  for (const plan of insertPlans) {
    const op = plan.op as Extract<Op, { op: "insertParagraphAfter" }>;
    const key = chainKeyOf(op, deps.paraByAddr.get(op.addr)!);
    const { insertedAddr, ids, element } = applyInsertParagraph(deps, op, insertChains, {
      followsParagraph: chainFollowsParagraph.get(key)!,
    });
    insertedAddrByIndex.set(plan.index, insertedAddr);
    changeIdsByIndex.set(plan.index, ids);
    deps.paraByAddr.set(insertedAddr, element); // receipts (and later ops) can address it
  }

  // — phase 3: receipts (before/after computed on the final tree) —
  return plans.map((plan): Receipt => {
    const base: Receipt = {
      index: plan.index,
      ref: plan.op.ref,
      author: plan.op.author,
      status: plan.status,
      addr: plan.op.addr,
      changeIds: [],
    };
    if (plan.status === "refused") return { ...base, reason: plan.reason };
    const changeIds = changeIdsByIndex.get(plan.index) ?? [];
    const para = deps.paraByAddr.get(plan.op.addr)!;
    if (plan.op.op === "insertParagraphAfter") {
      const insertedAddr = insertedAddrByIndex.get(plan.index)!;
      return {
        ...base,
        changeIds,
        insertedAddr,
        before: paragraphText(para, acceptAllView, deps.styles),
        after: paragraphText(deps.paraByAddr.get(insertedAddr) ?? para, acceptAllView, deps.styles),
      };
    }
    const rejectMine: View = (rev) => (rev.id !== null && changeIds.includes(rev.id) ? "reject" : "accept");
    return {
      ...base,
      changeIds,
      before: deps.recordText(plan.op.addr, rejectMine) ?? paragraphText(para, rejectMine, deps.styles),
      after: deps.recordText(plan.op.addr, acceptAllView) ?? paragraphText(para, acceptAllView, deps.styles),
    };
  });
}

// word granularity mutates only the changed tokens, so only those ranges are
// checked; the readable grouping is preferred, and the finest edits are the
// fallback when grouping would reach into a link or other fixed structure
function spanIssue(map: SpanMap, op: Extract<Op, { op: "replace" | "delete" }>, prior: Array<{ start: number; end: number }>, owned: ReadonlySet<XElement>): { issue: ReturnType<typeof checkSpan>; edits?: WordEdit[] } {
  if (op.op !== "replace" || op.granularity !== "word") return { issue: checkSpan(map, op.start, op.end, prior, owned) };
  const firstIssue = (edits: WordEdit[]) => {
    for (const e of edits) {
      const start = op.start + (e.delStart ?? e.insAt ?? 0);
      const end = e.delEnd !== undefined ? op.start + e.delEnd : start;
      const issue = checkSpan(map, start, end, prior, owned);
      if (issue) return issue;
    }
    return null;
  };
  const readable = readableWordEdits(op.expectText, op.text);
  if (!firstIssue(readable)) return { issue: null, edits: readable };
  const finest = wordEdits(op.expectText, op.text);
  return { issue: firstIssue(finest), edits: finest };
}

function refuse(plan: Plan, reason: NonNullable<Receipt["reason"]>): void {
  plan.status = "refused";
  plan.reason = reason;
}

function getMap(maps: Map<Addr, SpanMap>, para: XElement, deps: ApplyDeps): SpanMap {
  const addr = paraAddr(para);
  let map = addr ? maps.get(addr) : undefined;
  if (!map) {
    map = buildSpanMap(para, acceptAllView, deps.styles, deps.fieldEntry(para));
    if (addr) maps.set(addr, map);
  }
  return map;
}

/** true when `pos` falls between the high and low halves of a surrogate pair */
function splitsSurrogate(text: string, pos: number): boolean {
  if (pos <= 0 || pos >= text.length) return false;
  const hi = text.charCodeAt(pos - 1);
  const lo = text.charCodeAt(pos);
  return hi >= 0xd800 && hi <= 0xdbff && lo >= 0xdc00 && lo <= 0xdfff;
}

function paraAddr(p: XElement): Addr | null {
  const a = p.attrs.find((x) => x.name === "w14:paraId");
  return a ? a.raw : null;
}

function hasAncestor(p: XElement, name: string): boolean {
  let cur = p.parent;
  while (cur) {
    if (cur.name === name) return true;
    cur = cur.parent;
  }
  return false;
}

/** true when the paragraph lives inside a plain-text content control (an sdt
 *  whose sdtPr declares w:text): such a control may hold only one paragraph,
 *  and Word rejects the file when a second is added */
function inPlainTextControl(p: XElement): boolean {
  let cur = p.parent;
  while (cur) {
    if (cur.name === "w:sdt") {
      const sdtPr = kid(cur, "w:sdtPr");
      if (sdtPr && kid(sdtPr, "w:text")) return true;
    }
    cur = cur.parent;
  }
  return false;
}

// —————————————————————————————————————————————————————————————— span ops

type SpanOp = Extract<Op, { op: "replace" }> | Extract<Op, { op: "delete" }>;

function applySpanOp(deps: ApplyDeps, para: XElement, op: SpanOp, validated?: WordEdit[]): string[] {
  if (op.op === "delete") return applyDeletion(deps, para, op.start, op.end, op.author).ids;
  if (op.granularity === "word") {
    const ids: string[] = [];
    const edits = validated ?? wordEdits(op.expectText, op.text);
    const ordered = [...edits].sort((a, b) => {
      const pa = a.delStart ?? a.insAt ?? 0;
      const pb = b.delStart ?? b.insAt ?? 0;
      if (pb !== pa) return pb - pa;
      return (a.delStart !== undefined ? 0 : 1) - (b.delStart !== undefined ? 0 : 1);
    });
    for (const e of ordered) {
      if (e.delStart !== undefined && e.delEnd !== undefined && e.delEnd > e.delStart) {
        ids.push(...applyDeletion(deps, para, op.start + e.delStart, op.start + e.delEnd, op.author).ids);
      }
      if (e.insAt !== undefined && e.insText) {
        ids.push(applyPointInsertion(deps, para, op.start + e.insAt, e.insText, op.author));
      }
    }
    return ids;
  }
  // span granularity: one deletion (possibly segmented around bookmarks and
  // links) + one insertion right after it
  const del = applyDeletion(deps, para, op.start, op.end, op.author);
  if (del.insertAfter) {
    const insId = insertInsEl(deps, del.insertAfter, op.text, op.author, del.insertRunPr);
    return [...del.ids, insId];
  }
  // pure insertion (empty span)
  return [applyPointInsertion(deps, para, op.start, op.text, op.author)];
}

interface DeletionResult {
  ids: string[];
  /** formatting for replacement text: the first deleted run, or the first one
   *  outside a link when the deletion reached into one */
  insertRunPr: XElement | null;
  /** where replacement text goes: after the last deleted segment, or after
   *  its link when the deletion also covered text outside it */
  insertAfter: XElement | null;
}

function applyDeletion(deps: ApplyDeps, para: XElement, start: number, end: number, author: string): DeletionResult {
  if (end === start) return { ids: [], insertRunPr: null, insertAfter: null };
  // rebuild the map after each split: splitting changes the tree (text is
  // preserved, so offsets stay valid), and stale piece references would
  // split already-split runs and corrupt the paragraph
  let map = buildSpanMap(para, acceptAllView, deps.styles, deps.fieldEntry(para));
  splitAt(map, start, () => deps.ids.take());
  map = buildSpanMap(para, acceptAllView, deps.styles, deps.fieldEntry(para));
  splitAt(map, end, () => deps.ids.take());
  map = buildSpanMap(para, acceptAllView, deps.styles, deps.fieldEntry(para));
  const covered = map.pieces.filter((pc) => pc.start >= start && pc.end <= end);
  const runs: XElement[] = [];
  for (const pc of covered) if (!runs.includes(pc.run)) runs.push(pc.run);
  if (runs.length === 0) return { ids: [], insertRunPr: null, insertAfter: null };
  // defense in depth (checkSpan already gates this): covered runs may differ
  // only by the hyperlinks or smart tags that wrap them
  if (new Set(runs.map((r) => outsideLinks(r).parent)).size > 1) {
    throw new Error("mutation safety: covered runs span multiple containers");
  }
  const segmented = new Set(runs.map((r) => r.parent)).size > 1;
  const formatRun = segmented ? runs.find((r) => outsideLinks(r) === r) ?? runs[0]! : runs[0]!;
  const insertRunPr = deps.styles.insertedRunPr(formatRun, para);

  // contiguous sibling segments per parent: bookmark/comment markers between
  // runs, and the edges of a hyperlink or smart tag, split the deletion into
  // several w:del elements, each in the parent of its runs
  const coveredSet = new Set(runs);
  const segments: XElement[][] = [];
  for (const parent of new Set(runs.map((r) => r.parent!))) {
    let current: XElement[] = [];
    for (const c of parent.children) {
      if (c.kind === "el" && coveredSet.has(c)) current.push(c);
      else if (current.length) {
        segments.push(current);
        current = [];
      }
    }
    if (current.length) segments.push(current);
  }
  segments.sort((a, b) => runs.indexOf(a[0]!) - runs.indexOf(b[0]!));

  const delEls: XElement[] = [];
  for (let i = segments.length - 1; i >= 0; i--) {
    const seg = segments[i]!;
    const parent = seg[0]!.parent!;
    const rev = newRevisionEl("w:del", deps.ids.take(), author, deps.date);
    const segStart = parent.children.indexOf(seg[0]!);
    parent.children.splice(segStart, seg.length, rev);
    rev.children = seg;
    for (const r of seg) r.parent = rev;
    rev.parent = parent;
    delEls.unshift(rev);
    deps.ownedRevisions.add(rev);
  }
  for (const del of delEls) {
    for (const d of descendantsOf(del)) {
      if (d.name === "w:t") renameEl(d, "w:delText"); // w:t lives inside runs
    }
  }
  const lastDel = delEls[delEls.length - 1]!;
  return { ids: delEls.map((d) => d.attrs.find((a) => a.name === "w:id")!.raw), insertRunPr, insertAfter: segmented ? outsideLinks(lastDel) : lastDel };
}

export function* descendantsOf(root: XElement): Generator<XElement> {
  for (const c of root.children) {
    if (c.kind === "el") {
      yield c;
      yield* descendantsOf(c);
    }
  }
}

/** insertion immediately after an existing element (the deleted span's last segment) */
function insertInsEl(deps: ApplyDeps, afterEl: XElement, text: string, author: string, rPr: XElement | null): string {
  const ins = newRevisionEl("w:ins", deps.ids.take(), author, deps.date);
  const run = makeRun(rPr, textToRunChildren(text));
  ins.children.push(run);
  run.parent = ins;
  const parent = afterEl.parent!;
  parent.children.splice(parent.children.indexOf(afterEl) + 1, 0, ins);
  ins.parent = parent;
  deps.ownedRevisions.add(ins);
  return ins.attrs.find((a) => a.name === "w:id")!.raw;
}

/** the outermost hyperlink or smartTag around `run` with no visible text before
 *  `at`: text inserted at its leading edge sits before it, as at its trailing edge */
function leadingStructure(map: SpanMap, run: XElement, at: number): XElement | null {
  let outer: XElement | null = null;
  for (let cur = run.parent; cur && cur.name !== "w:p"; cur = cur.parent) {
    if (cur.name !== "w:hyperlink" && cur.name !== "w:smartTag") continue;
    const within = (r: XElement) => {
      for (let a = r.parent; a; a = a.parent) if (a === cur) return true;
      return false;
    };
    if (map.pieces.some((o) => o.end > o.start && o.end <= at && within(o.run))) break;
    outer = cur;
  }
  return outer;
}

/** point insertion at a text offset (word-diff insertions, empty-span replace).
 *  Like typing in Word, the text takes the formatting of the visible text just
 *  before it in the same parent, else just after it */
function applyPointInsertion(deps: ApplyDeps, para: XElement, at: number, text: string, author: string): string {
  // ensure no run straddles the insertion point, then place the ins between runs
  let map = buildSpanMap(para, acceptAllView, deps.styles, deps.fieldEntry(para));
  splitAt(map, at, () => deps.ids.take());
  map = buildSpanMap(para, acceptAllView, deps.styles, deps.fieldEntry(para));
  const pc = pieceAt(map, at);
  let parent: XElement;
  let index: number;
  const structure = pc && pc.start >= at ? leadingStructure(map, pc.run, at) : null;
  if (structure?.parent) {
    parent = structure.parent;
    index = parent.children.indexOf(structure);
  } else if (pc && pc.start >= at) {
    parent = pc.run.parent!;
    index = parent.children.indexOf(pc.run);
  } else if (pc) {
    parent = pc.run.parent!;
    index = parent.children.indexOf(pc.run) + 1;
  } else {
    parent = para;
    index = para.children.length;
  }
  // sibling climb: a point that lands inside a revision created by THIS call
  // must stay OUT of that revision — place the new ins beside it, never nested
  // (checkSpan already refuses points inside pre-existing revisions)
  const rev = pc ? map.runInfo.get(pc.run)?.rev : null;
  if (rev && rev.parent && deps.ownedRevisions.has(rev)) {
    parent = rev.parent;
    index = parent.children.indexOf(rev) + (pc && pc.start < at ? 1 : 0);
  }
  const sameParent = map.pieces.filter((o) => o.end > o.start && o.run.parent === parent && !o.inFieldResult);
  const formatRun = sameParent.filter((o) => o.end <= at).at(-1)?.run ?? sameParent.find((o) => o.start >= at)?.run ?? null;
  const ins = newRevisionEl("w:ins", deps.ids.take(), author, deps.date);
  const run = makeRun(deps.styles.insertedRunPr(formatRun, para), textToRunChildren(text));
  ins.children.push(run);
  run.parent = ins;
  parent.children.splice(index, 0, ins);
  ins.parent = parent;
  deps.ownedRevisions.add(ins);
  return ins.attrs.find((a) => a.name === "w:id")!.raw;
}

// ———————————————————————————————————————————————————————————— insert para

const BLOCK_SIBLINGS = new Set(["w:p", "w:tbl", "w:sdt", "w:customXml"]);

/** the next block-level sibling (paragraph, table, block sdt), or null */
function nextBlockSibling(e: XElement): XElement | null {
  const parent = e.parent;
  if (!parent) return null;
  for (let i = parent.children.indexOf(e) + 1; i < parent.children.length; i++) {
    const c = parent.children[i]!;
    if (c.kind === "el" && BLOCK_SIBLINGS.has(c.name)) return c;
  }
  return null;
}

/** element a chained insert hangs off: the containing table for scope "table", else the anchor paragraph */
function chainKeyOf(op: Extract<Op, { op: "insertParagraphAfter" }>, anchor: XElement): XElement {
  if (op.scope === "table") {
    let cur = anchor.parent;
    while (cur && cur.name !== "w:tbl") cur = cur.parent;
    if (cur) return cur;
  }
  return anchor;
}

function applyInsertParagraph(
  deps: ApplyDeps,
  op: Extract<Op, { op: "insertParagraphAfter" }>,
  chains: Map<XElement, XElement>,
  form: { followsParagraph: boolean },
): { insertedAddr: string; ids: string[]; element: XElement } {
  const anchor = deps.paraByAddr.get(op.addr)!;
  const formatSource = op.formatFrom === undefined ? null : deps.paraByAddr.get(op.formatFrom)!;
  const chainKey = chainKeyOf(op, anchor);
  const effectiveAfter = chains.get(chainKey) ?? chainKey;
  const author = op.author;

  // Word's container-end form, as typing at the end of a document makes it: no
  // paragraph follows the chain, so each insert tracks the mark of the
  // paragraph before it (the anchor's for the first) and keeps an untracked
  // mark plus a w:pPrChange recording the anchor's original properties. Any
  // subset of rejects then removes exactly the rejected paragraphs, and the
  // anchor's text always lands on a paragraph restored to its properties.
  // (After a table there is no anchor mark to track — the mark rides on the
  // new paragraph, whose removal Word cannot reject either.)
  const anchorMarkForm = !form.followsParagraph && chainKey.name === "w:p";

  const p = el("w:p");
  const ids: string[] = [];
  let markIns: XElement | null = null;
  if (anchorMarkForm) {
    const markId = deps.ids.take();
    const changeId = deps.ids.take();
    const runId = deps.ids.take();
    ids.push(markId, changeId, runId);

    const pPr = el("w:pPr");
    const change = newRevisionEl("w:pPrChange", changeId, author, deps.date);
    const saved = el("w:pPr");
    for (const c of portableProperties(anchor)) adopt(saved, c);
    if (formatSource) for (const c of portableProperties(formatSource)) adopt(pPr, c);
    change.children.push(saved);
    saved.parent = change;
    pPr.children.push(change);
    change.parent = pPr;
    p.children.push(pPr);
    pPr.parent = p;

    markIns = newRevisionEl("w:ins", markId, author, deps.date);
    addParagraphMarkIns(effectiveAfter, markIns);
    deps.ownedRevisions.add(markIns);
  } else {
    const markId = deps.ids.take();
    const runId = deps.ids.take();
    ids.push(markId, runId);
    const pPr = el("w:pPr");
    if (formatSource) for (const c of portableProperties(formatSource)) adopt(pPr, c);
    markIns = newRevisionEl("w:ins", markId, author, deps.date);
    const rPr = el("w:rPr");
    rPr.children.push(markIns);
    markIns.parent = rPr;
    pPr.children.push(rPr);
    rPr.parent = pPr;
    p.children.push(pPr);
    pPr.parent = p;
  }

  const runId = ids[ids.length - 1]!;
  const runIns = newRevisionEl("w:ins", runId, author, deps.date);
  const anchorRunPr = deps.styles.insertedRunPr(formatSource ? firstPlainRun(formatSource) : firstRun(anchor), p);
  const run = makeRun(anchorRunPr, textToRunChildren(op.text));
  runIns.children.push(run);
  run.parent = runIns;
  p.children.push(runIns);
  runIns.parent = p;
  deps.ownedRevisions.add(runIns);

  const parent = effectiveAfter.parent!;
  parent.children.splice(parent.children.indexOf(effectiveAfter) + 1, 0, p);
  p.parent = parent;
  chains.set(chainKey, p);

  const addr = deps.paraIds.take();
  setAttr(p, "w14:paraId", addr);
  return { insertedAddr: addr, ids, element: p };
}

/** copies of a paragraph's properties a new paragraph can carry: mark
 *  formatting, change records and the section break (which must not be
 *  duplicated onto the insert) stay put; a copied property's own revision
 *  record keeps its unique id only on the original */
function portableProperties(source: XElement): XNode[] {
  const holder = el("w:pPr");
  for (const c of kid(source, "w:pPr")?.children ?? []) {
    if (c.kind === "el" && (c.name === "w:rPr" || c.name === "w:pPrChange" || c.name === "w:sectPr")) continue;
    adopt(holder, cloneNode(c));
  }
  stripRevisionRecords(holder);
  return [...holder.children];
}

function adopt(parent: XElement, child: XNode): void {
  parent.children.push(child);
  if (child.kind === "el") child.parent = parent;
}

/**
 * Track an existing paragraph's mark as inserted: the w:ins goes into the
 * paragraph's w:pPr/w:rPr (creating either level when missing), first among
 * the rPr children per CT_ParaRPr order, and before w:sectPr/w:pPrChange in
 * the pPr per CT_PPr order.
 */
function addParagraphMarkIns(para: XElement, ins: XElement): void {
  let pPr = kid(para, "w:pPr");
  if (!pPr) {
    pPr = el("w:pPr");
    para.children.unshift(pPr);
    pPr.parent = para;
  }
  const rPr = kid(pPr, "w:rPr");
  if (rPr) {
    rPr.children.unshift(ins);
    ins.parent = rPr;
    return;
  }
  const fresh = el("w:rPr");
  fresh.children.push(ins);
  ins.parent = fresh;
  let at = pPr.children.length;
  while (at > 0) {
    const last = pPr.children[at - 1]!;
    if (last.kind === "el" && (last.name === "w:sectPr" || last.name === "w:pPrChange")) at--;
    else break;
  }
  pPr.children.splice(at, 0, fresh);
  fresh.parent = pPr;
}

function firstRun(p: XElement): XElement | null {
  for (const c of p.children) {
    if (c.kind === "el" && c.name === "w:r") return c;
  }
  return null;
}

/** the first run of ordinary text, through inline wrappers such as smartTag or
 *  customXml: link, deleted and moved-away text would carry formatting that
 *  misrepresents plain inserted text */
function firstPlainRun(p: XElement): XElement | null {
  const skip = new Set(["w:hyperlink", "w:del", "w:moveFrom", "w:pPr", "mc:AlternateContent", "m:oMath", "m:oMathPara"]);
  const visit = (node: XElement): XElement | null => {
    for (const c of node.children) {
      if (c.kind !== "el" || skip.has(c.name)) continue;
      if (c.name === "w:r") return c;
      const found = visit(c);
      if (found) return found;
    }
    return null;
  };
  return visit(p);
}

/**
 * Ensure no run straddles `pos`: split text pieces at char boundaries and
 * split runs at piece boundaries, so the covered pieces after both splits
 * belong to runs that are fully covered.
 */
export function splitAt(map: SpanMap, pos: number, takeId: () => string): void {
  const pc = pieceAt(map, pos);
  if (!pc) return;
  if (pc.kind === "text" && pc.start < pos && pos < pc.end) {
    splitRun(pc.run, pc.el, pos - pc.start, takeId);
    return;
  }
  if (pc.start === pos) splitBefore(pc.run, pc.el, takeId);
  else if (pc.end === pos) splitAfter(pc.run, pc.el, takeId);
}

/** split the run so `el` is the first content child of the right part */
function splitBefore(run: XElement, el_: XElement, takeId: () => string): void {
  const parent = run.parent;
  if (!parent) return;
  const idx = run.children.indexOf(el_);
  if (idx <= (runPrOf(run) ? 1 : 0)) return;
  const leftRun = cloneNode(run);
  leftRun.children = run.children.slice(0, idx);
  rightTo(run, parent, idx, leftRun, takeId);
}

/** split the run so `el` is the last content child of the left part */
function splitAfter(run: XElement, el_: XElement, takeId: () => string): void {
  const parent = run.parent;
  if (!parent) return;
  const idx = run.children.indexOf(el_);
  if (idx < 0 || idx >= run.children.length - 1) return;
  const leftRun = cloneNode(run);
  leftRun.children = run.children.slice(0, idx + 1);
  rightTo(run, parent, idx + 1, leftRun, takeId);
}

function rightTo(run: XElement, parent: XElement, from: number, leftRun: XElement, takeId: () => string): void {
  const rightRun = cloneNode(run);
  rightRun.children = run.children.slice(from);
  // the right half keeps a copy of the run's rPr (it is the covered/remaining content's formatting);
  // revision records inside the copy get fresh ids — the left half keeps the originals
  const head = runPrOf(run);
  if (head && from > 0) {
    const rPrClone = cloneNode(head);
    reidRevisionRecords(rPrClone, takeId);
    rightRun.children.unshift(rPrClone);
    rPrClone.parent = rightRun;
  }
  for (const c of leftRun.children) if (c.kind === "el") c.parent = leftRun;
  for (const c of rightRun.children) if (c.kind === "el") c.parent = rightRun;
  const idx = parent.children.indexOf(run);
  parent.children.splice(idx, 1, leftRun, rightRun);
  leftRun.parent = parent;
  rightRun.parent = parent;
}

/** the w:rPr child of a run, if its first child is one */
function runPrOf(run: XElement): XElement | null {
  const first = run.children[0];
  return first && first.kind === "el" && first.name === "w:rPr" ? first : null;
}

/** split the run whose text piece strictly contains pos (text-level split).
 *  `inner` is a Word-read offset inside the element. Each half carries the
 *  exact characters Word reads there, under xml:space="preserve": a preserved
 *  original keeps its raw bytes (read == raw up to newline/tab mapping), and a
 *  plain original contributes its Word-read slices — so edge whitespace Word
 *  drops cannot resurface at a cut, and internal whitespace cannot vanish. */
function splitRun(run: XElement, tEl: XElement, inner: number, takeId: () => string): void {
  const parent = run.parent;
  if (!parent) return;
  const idx = parent.children.indexOf(run);
  const source = attrVal(tEl, "xml:space") === "preserve" ? textContent(tEl) : wordText(tEl);
  const tIdx = run.children.indexOf(tEl);
  const rPr = runPrOf(run);

  const leftRun = cloneNode(run);
  leftRun.children = [...run.children.slice(0, tIdx), makeText(source.slice(0, inner))];
  const rightRun = cloneNode(run);
  const rPrCopy = rPr ? cloneNode(rPr) : null;
  if (rPrCopy) reidRevisionRecords(rPrCopy, takeId);
  rightRun.children = [...(rPrCopy ? [rPrCopy] : []), makeText(source.slice(inner)), ...run.children.slice(tIdx + 1)];
  for (const c of leftRun.children) if (c.kind === "el") c.parent = leftRun;
  for (const c of rightRun.children) if (c.kind === "el") c.parent = rightRun;

  parent.children.splice(idx, 1, leftRun, rightRun);
  leftRun.parent = parent;
  rightRun.parent = parent;
}

function makeText(text: string, preserve = true): XElement {
  const t = el("w:t", preserve ? { "xml:space": "preserve" } : {});
  t.children.push(textNode(text)); // textNode escapes on the way in
  return t;
}
