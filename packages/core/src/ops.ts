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
} from "@stitchdocx/package-io";
import { buildSpanMap, checkSpan, pieceAt, type SpanMap } from "./spanmap.js";
import { markRevision, paragraphText, textContent, type View } from "./projection.js";
import {
  cloneRunPr,
  makeRun,
  newRevisionEl,
  renameEl,
  RevisionIdAllocator,
  ParaIdAllocator,
  textToRunChildren,
} from "./revisions.js";
import { wordEdits } from "./worddiff.js";
import type { Addr, Op, Receipt } from "./types.js";

export interface ApplyDeps {
  paraByAddr: Map<Addr, XElement>;
  ids: RevisionIdAllocator;
  paraIds: ParaIdAllocator;
  date: string;
  /** revision elements created by THIS apply call — the only ones a later op
   *  may treat as its own (sibling placement); everything else is untouchable */
  ownedRevisions: Set<XElement>;
}

const acceptAllView: View = () => "accept";

interface Plan {
  op: Op;
  index: number;
  status: "applied" | "refused";
  reason?: Receipt["reason"];
  result?: { changeIds: string[]; insertedAddr?: string };
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
    if (op.op === "insertParagraphAfter") {
      if (!op.text || op.text.trim() === "") refuse(plan, "empty_text");
      else if (op.scope === "table" && !hasAncestor(para, "w:tbl")) refuse(plan, "addr_not_found");
      else if (!op.scope && nextBlockSibling(para)?.name !== "w:p" && markRevision(para)) {
        // at a container end the insert must track the ANCHOR's paragraph
        // mark (Word's form); an anchor whose mark is already tracked by any
        // author cannot carry ours, and Word could not reject the result
        refuse(plan, "paragraph_not_editable");
      }
      continue;
    }
    const map = getMap(maps, para);
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
    if (map.text.slice(op.start, op.end) !== op.expectText) {
      refuse(plan, "span_text_mismatch");
      continue;
    }
    if (op.op === "replace" && (!op.text || op.text.trim() === "")) {
      refuse(plan, "empty_text");
      continue;
    }
    const issue = checkSpan(map, op.start, op.end, priorApplied.get(op.addr) ?? [], ownedRevisions);
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
      changeIdsByIndex.set(plan.index, applySpanOp(deps, para, plan.op as SpanOp));
    }
  }
  const insertChains = new Map<XElement, XElement>();
  const insertPlans = plans.filter((p) => p.status === "applied" && p.op.op === "insertParagraphAfter");
  // decided up front against the pristine tree: the last insert per anchor
  // ends its chain, and whether a paragraph follows the chain decides the
  // markup form (inside the loop the tree is already mutating)
  const chainEndPlan = new Map<XElement, Plan>();
  const chainFollowsParagraph = new Map<XElement, boolean>();
  for (const plan of insertPlans) {
    const op = plan.op as Extract<Op, { op: "insertParagraphAfter" }>;
    const key = chainKeyOf(op, deps.paraByAddr.get(op.addr)!);
    chainEndPlan.set(key, plan);
    if (!chainFollowsParagraph.has(key)) chainFollowsParagraph.set(key, nextBlockSibling(key)?.name === "w:p");
  }
  for (const plan of insertPlans) {
    const op = plan.op as Extract<Op, { op: "insertParagraphAfter" }>;
    const key = chainKeyOf(op, deps.paraByAddr.get(op.addr)!);
    const { insertedAddr, ids, element } = applyInsertParagraph(deps, op, insertChains, {
      chainEnd: chainEndPlan.get(key) === plan,
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
        before: paragraphText(para, acceptAllView),
        after: paragraphText(deps.paraByAddr.get(insertedAddr) ?? para, acceptAllView),
      };
    }
    const rejectMine: View = (rev) => (rev.id !== null && changeIds.includes(rev.id) ? "reject" : "accept");
    return {
      ...base,
      changeIds,
      before: paragraphText(para, rejectMine),
      after: paragraphText(para, acceptAllView),
    };
  });
}

function refuse(plan: Plan, reason: NonNullable<Receipt["reason"]>): void {
  plan.status = "refused";
  plan.reason = reason;
}

function getMap(maps: Map<Addr, SpanMap>, para: XElement): SpanMap {
  const addr = paraAddr(para);
  let map = addr ? maps.get(addr) : undefined;
  if (!map) {
    map = buildSpanMap(para, acceptAllView);
    if (addr) maps.set(addr, map);
  }
  return map;
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

// —————————————————————————————————————————————————————————————— span ops

type SpanOp = Extract<Op, { op: "replace" }> | Extract<Op, { op: "delete" }>;

function applySpanOp(deps: ApplyDeps, para: XElement, op: SpanOp): string[] {
  if (op.op === "delete") return applyDeletion(deps, para, op.start, op.end, op.author).ids;
  if (op.granularity === "word") {
    const ids: string[] = [];
    const edits = wordEdits(op.expectText, op.text);
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
        ids.push(applyPointInsertion(deps, para, op.start + e.insAt, e.insText, op.author, null));
      }
    }
    return ids;
  }
  // span granularity: one deletion (possibly segmented around bookmarks) + one
  // insertion right after the last deleted segment
  const del = applyDeletion(deps, para, op.start, op.end, op.author);
  if (del.lastDelEl) {
    const insId = insertInsEl(deps, del.lastDelEl, op.text, op.author, del.firstRunPr);
    return [...del.ids, insId];
  }
  // pure insertion (empty span)
  return [applyPointInsertion(deps, para, op.start, op.text, op.author, null)];
}

interface DeletionResult {
  ids: string[];
  firstRunPr: XElement | null;
  lastDelEl: XElement | null;
}

function applyDeletion(deps: ApplyDeps, para: XElement, start: number, end: number, author: string): DeletionResult {
  if (end === start) return { ids: [], firstRunPr: null, lastDelEl: null };
  // rebuild the map after each split: splitting changes the tree (text is
  // preserved, so offsets stay valid), and stale piece references would
  // split already-split runs and corrupt the paragraph
  let map = buildSpanMap(para, acceptAllView);
  splitAt(map, start);
  map = buildSpanMap(para, acceptAllView);
  splitAt(map, end);
  map = buildSpanMap(para, acceptAllView);
  const covered = map.pieces.filter((pc) => pc.start >= start && pc.end <= end);
  const runs: XElement[] = [];
  for (const pc of covered) if (!runs.includes(pc.run)) runs.push(pc.run);
  if (runs.length === 0) return { ids: [], firstRunPr: null, lastDelEl: null };
  // defense in depth (checkSpan already gates this): refuse to mutate when the
  // covered runs do not share one parent — partial deletions are forbidden
  if (new Set(runs.map((r) => r.parent)).size > 1) {
    throw new Error("mutation safety: covered runs span multiple parents");
  }
  const firstRunPr = cloneRunPr(runs[0]!);

  // contiguous sibling segments (bookmark/comment markers between runs split
  // the deletion into several w:del elements; everything shares one parent —
  // the structure check guarantees it)
  const parent = runs[0]!.parent!;
  const coveredSet = new Set(runs);
  const segments: XElement[][] = [];
  let current: XElement[] = [];
  for (const c of parent.children) {
    if (c.kind === "el" && coveredSet.has(c)) current.push(c);
    else if (current.length) {
      segments.push(current);
      current = [];
    }
  }
  if (current.length) segments.push(current);

  const delEls: XElement[] = [];
  for (let i = segments.length - 1; i >= 0; i--) {
    const seg = segments[i]!;
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
  return { ids: delEls.map((d) => d.attrs.find((a) => a.name === "w:id")!.raw), firstRunPr, lastDelEl: delEls[delEls.length - 1]! };
}

function* descendantsOf(root: XElement): Generator<XElement> {
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

/** point insertion at a text offset (word-diff insertions, empty-span replace) */
function applyPointInsertion(deps: ApplyDeps, para: XElement, at: number, text: string, author: string, rPr: XElement | null): string {
  const ins = newRevisionEl("w:ins", deps.ids.take(), author, deps.date);
  const run = makeRun(rPr, textToRunChildren(text));
  ins.children.push(run);
  run.parent = ins;
  // ensure no run straddles the insertion point, then place the ins between runs
  let map = buildSpanMap(para, acceptAllView);
  splitAt(map, at);
  map = buildSpanMap(para, acceptAllView);
  const pc = pieceAt(map, at);
  let parent: XElement;
  let index: number;
  if (pc && pc.start >= at) {
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
  form: { chainEnd: boolean; followsParagraph: boolean },
): { insertedAddr: string; ids: string[]; element: XElement } {
  const anchor = deps.paraByAddr.get(op.addr)!;
  const chainKey = chainKeyOf(op, anchor);
  const effectiveAfter = chains.get(chainKey) ?? chainKey;
  const author = op.author;

  // Word's container-end form: no paragraph follows the chain, so the anchor's
  // own mark is the tracked insertion and the last new paragraph keeps an
  // untracked mark plus a w:pPrChange recording the anchor's original
  // properties. Rejecting the pair restores the anchor exactly. (After a
  // table there is no anchor mark to track — the mark rides on the new
  // paragraph, whose removal Word cannot reject either.)
  const anchorMarkForm = form.chainEnd && !form.followsParagraph && chainKey.name === "w:p";

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
    const anchorPPr = kid(anchor, "w:pPr");
    for (const c of anchorPPr?.children ?? []) {
      // the saved state carries the anchor's paragraph properties; mark
      // formatting and nested change records are not paragraph properties
      if (c.kind === "el" && (c.name === "w:rPr" || c.name === "w:pPrChange")) continue;
      const clone = cloneNode(c);
      saved.children.push(clone);
      if (clone.kind === "el") clone.parent = saved;
    }
    change.children.push(saved);
    saved.parent = change;
    pPr.children.push(change);
    change.parent = pPr;
    p.children.push(pPr);
    pPr.parent = p;

    markIns = newRevisionEl("w:ins", markId, author, deps.date);
    addParagraphMarkIns(anchor, markIns);
    deps.ownedRevisions.add(markIns);
  } else {
    const markId = deps.ids.take();
    const runId = deps.ids.take();
    ids.push(markId, runId);
    const pPr = el("w:pPr");
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
  const anchorRunPr = cloneRunPr(firstRun(anchor));
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

/**
 * Ensure no run straddles `pos`: split text pieces at char boundaries and
 * split runs at piece boundaries, so the covered pieces after both splits
 * belong to runs that are fully covered.
 */
function splitAt(map: SpanMap, pos: number): void {
  const pc = pieceAt(map, pos);
  if (!pc) return;
  if (pc.kind === "text" && pc.start < pos && pos < pc.end) {
    splitRun(pc.run, pc.el, pos - pc.start);
    return;
  }
  if (pc.start === pos) splitBefore(pc.run, pc.el);
  else if (pc.end === pos) splitAfter(pc.run, pc.el);
}

/** split the run so `el` is the first content child of the right part */
function splitBefore(run: XElement, el_: XElement): void {
  const parent = run.parent;
  if (!parent) return;
  const idx = run.children.indexOf(el_);
  if (idx <= 0) return;
  const leftRun = cloneNode(run);
  leftRun.children = run.children.slice(0, idx);
  rightTo(run, parent, idx, leftRun);
}

/** split the run so `el` is the last content child of the left part */
function splitAfter(run: XElement, el_: XElement): void {
  const parent = run.parent;
  if (!parent) return;
  const idx = run.children.indexOf(el_);
  if (idx < 0 || idx >= run.children.length - 1) return;
  const leftRun = cloneNode(run);
  leftRun.children = run.children.slice(0, idx + 1);
  rightTo(run, parent, idx + 1, leftRun);
}

function rightTo(run: XElement, parent: XElement, from: number, leftRun: XElement): void {
  const rightRun = cloneNode(run);
  rightRun.children = run.children.slice(from);
  // the right half keeps a copy of the run's rPr (it is the covered/remaining content's formatting)
  const head = runPrOf(run);
  if (head && from > 0) {
    const rPrClone = cloneNode(head);
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

/** split the run whose text piece strictly contains pos (text-level split) */
function splitRun(run: XElement, tEl: XElement, inner: number): void {
  const parent = run.parent;
  if (!parent) return;
  const idx = parent.children.indexOf(run);
  const decoded = textContent(tEl);
  const tIdx = run.children.indexOf(tEl);
  const rPr = runPrOf(run);

  const leftRun = cloneNode(run);
  leftRun.children = [...run.children.slice(0, tIdx), makeText(decoded.slice(0, inner))];
  const rightRun = cloneNode(run);
  rightRun.children = [...(rPr ? [cloneNode(rPr)] : []), makeText(decoded.slice(inner)), ...run.children.slice(tIdx + 1)];
  for (const c of leftRun.children) if (c.kind === "el") c.parent = leftRun;
  for (const c of rightRun.children) if (c.kind === "el") c.parent = rightRun;

  parent.children.splice(idx, 1, leftRun, rightRun);
  leftRun.parent = parent;
  rightRun.parent = parent;
}

function makeText(text: string): XElement {
  const t = el("w:t", { "xml:space": "preserve" });
  t.children.push(textNode(text)); // textNode escapes on the way in
  return t;
}
