// Span map: the paragraph's visible text tiled into char pieces, each tied to
// the run that produced it, with the ancestry metadata the refusal checks need
// (revision, hyperlink/smartTag structure, inline content control, field,
// locked content, tracked rows).
// Text coordinates are what Word reads (wordText), the same space the
// projection and the contract's offsets live in.
import { kid, kids, type XElement } from "@stitchdocx/package-io";
import { walkContent, wordText, symChar, type CharKind, type View } from "./projection.js";

export type LockKind = "alternate" | "math";

export interface SpanPiece {
  run: XElement;
  el: XElement;
  kind: CharKind;
  start: number;
  end: number;
  /** decoded text for "text" pieces; the character for "sym"; "" for other 1-char atomics */
  decoded: string;
  inFieldResult: boolean;
  /** content the engine lists but refuses to edit (AlternateContent branch, office math) */
  lock: LockKind | null;
}

export interface RunInfo {
  /** nearest enclosing revision element (any author), or null */
  rev: XElement | null;
  /** nearest enclosing w:hyperlink / w:smartTag, or null */
  structure: XElement | null;
  /** a w:sdt sits strictly between the run and its paragraph */
  inlineSdt: boolean;
  /** the run's row is itself a tracked insertion or deletion */
  trackedRow: boolean;
}

export interface SpanMap {
  text: string;
  pieces: SpanPiece[];
  runInfo: Map<XElement, RunInfo>;
}

export function buildSpanMap(p: XElement, view: View, fields: string[] = []): SpanMap {
  const pieces: SpanPiece[] = [];
  const runInfo = new Map<XElement, RunInfo>();
  let pos = 0;
  walkContent(p, view, {
    onCharEl: (el, kind, inFieldResult, locked) => {
      const run = el.parent;
      if (!run || run.kind !== "el") return;
      if (!runInfo.has(run)) runInfo.set(run, ancestry(run, p));
      const decoded = kind === "text" ? wordText(el) : kind === "sym" ? symChar(el) ?? "" : "";
      if (kind === "sym" && decoded === "") return;
      // a text piece may read empty (a lone unpreserved space); it occupies
      // no offset — only 1-char atomics (tab, break, hyphen) reserve one
      const len = decoded.length || (kind === "text" || kind === "sym" ? 0 : 1);
      pieces.push({
        run,
        el,
        kind,
        start: pos,
        end: pos + len,
        decoded,
        inFieldResult,
        lock: locked ? (el.name === "m:t" ? "math" : el.name === "w:sym" && lockedInside(p, el) ? "math" : "alternate") : null,
      });
      pos += len;
    },
  }, fields);
  return {
    text: pieces
      .map((pc) => (pc.kind === "text" || pc.kind === "sym" ? pc.decoded : pc.kind === "tab" ? "\t" : pc.kind === "break" ? "\n" : "-"))
      .join(""),
    pieces,
    runInfo,
  };
}

/** whether the w:sym sits inside office math (the run's ancestry reaches an
 *  m:oMath before the paragraph) */
function lockedInside(paragraph: XElement, el: XElement): boolean {
  let cur: XElement | null = el.parent;
  while (cur && cur !== paragraph) {
    if (cur.name === "m:oMath" || cur.name === "m:oMathPara") return true;
    cur = cur.parent;
  }
  return false;
}

function ancestry(run: XElement, paragraph: XElement): RunInfo {
  let rev: XElement | null = null;
  let structure: XElement | null = null;
  let inlineSdt = false;
  let cur: XElement | null = run.parent;
  while (cur && cur !== paragraph) {
    if (!rev && (cur.name === "w:ins" || cur.name === "w:del" || cur.name === "w:moveFrom" || cur.name === "w:moveTo")) rev = cur;
    if (!structure && (cur.name === "w:hyperlink" || cur.name === "w:smartTag")) structure = cur;
    if (cur.name === "w:sdt") inlineSdt = true;
    cur = cur.parent;
  }
  return { rev, structure, inlineSdt, trackedRow: trackedRowOf(paragraph) };
}

/** true when the paragraph's row carries a row-level tracked change */
function trackedRowOf(paragraph: XElement): boolean {
  let cur: XElement | null = paragraph.parent;
  while (cur) {
    if (cur.name === "w:tr") {
      const trPr = kid(cur, "w:trPr");
      return trPr !== undefined && kids(trPr).some((c) => c.name === "w:del" || c.name === "w:ins");
    }
    if (cur.name === "w:tbl") return false;
    cur = cur.parent;
  }
  return false;
}

export interface SpanIssue {
  reason:
    | "span_overlaps_revision"
    | "span_overlaps_prior_op"
    | "span_crosses_structure"
    | "span_in_field"
    | "span_in_content_control";
  detail?: string;
}

/**
 * Refusal checks over [start,end) in paragraph-text coordinates. Predicates
 * are collected over every covered piece first, then decided in the contract's
 * priority order (API.md refusal table): overlaps_revision, overlaps_prior_op,
 * crosses_structure, in_field, in_content_control, locked content — a
 * per-piece first-hit would misreport (e.g. a span crossing into a field
 * result inside a hyperlink must refuse crosses_structure, not in_field).
 */
export function checkSpan(map: SpanMap, start: number, end: number, priorOps: Array<{ start: number; end: number }>, ownedRevisions: ReadonlySet<XElement> = new Set()): SpanIssue | null {
  const covered = map.pieces.filter((pc) => pc.start < end && pc.end > start);
  let hasRev = false;
  let hasField = false;
  let hasSdt = false;
  let hasAlternate = false;
  let hasMath = false;
  let hasTrackedRow = false;
  const structures = new Set<XElement | null>();
  const parents = new Set<XElement | null>();
  for (const pc of covered) {
    const info = map.runInfo.get(pc.run)!;
    if (info.rev) hasRev = true;
    if (info.trackedRow) hasTrackedRow = true;
    if (pc.inFieldResult) hasField = true;
    if (info.inlineSdt) hasSdt = true;
    if (pc.lock === "alternate") hasAlternate = true;
    if (pc.lock === "math") hasMath = true;
    structures.add(info.structure);
    parents.add(pc.run.parent);
  }
  if (hasRev || hasTrackedRow) return { reason: "span_overlaps_revision" };
  for (const pr of priorOps) {
    if (start < pr.end && pr.start < end) return { reason: "span_overlaps_prior_op" };
  }
  // mutation safety: every covered run must share one parent (an unmodelled
  // inline wrapper such as w:customXml would otherwise be partially deleted).
  // Office math and AlternateContent are structure the engine lists but never rewrites.
  if (parents.size > 1 || structures.size > 1 || hasAlternate || hasMath) return { reason: "span_crosses_structure" };
  if (hasField) return { reason: "span_in_field" };
  if (hasSdt) return { reason: "span_in_content_control" };
  if (start === end) {
    // point insertion: the field/sdt protections still apply at the point, and
    // the surrounding run's revision is refused unless THIS call created it
    // (those get sibling placement instead)
    const pc = covered[0] ?? pieceAt(map, start);
    if (pc && start >= pc.start && start < pc.end) {
      const info = map.runInfo.get(pc.run)!;
      if ((info.rev && !ownedRevisions.has(info.rev)) || info.trackedRow) return { reason: "span_overlaps_revision" };
      if (pc.inFieldResult) return { reason: "span_in_field" };
      if (info.inlineSdt) return { reason: "span_in_content_control" };
      if (pc.lock) return { reason: "span_crosses_structure" };
    }
  }
  return null;
}

/** offset -> the piece covering it (or the next piece at/after it) */
export function pieceAt(map: SpanMap, offset: number): SpanPiece | null {
  for (const pc of map.pieces) {
    if (offset >= pc.start && offset < pc.end) return pc;
    if (pc.start >= offset) return pc;
  }
  return null;
}
