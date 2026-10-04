// Span map: the paragraph's visible text tiled into char pieces, each tied to
// the run that produced it, with the ancestry metadata the refusal checks need
// (revision, hyperlink/smartTag structure, inline content control, field).
// Text coordinates are what Word reads (wordText), the same space the
// projection and the contract's offsets live in.
import { type XElement } from "@stitchdocx/package-io";
import { walkContent, wordText, type CharKind, type View } from "./projection.js";

export interface SpanPiece {
  run: XElement;
  el: XElement;
  kind: CharKind;
  start: number;
  end: number;
  /** decoded text for "text" pieces; "" for 1-char atomics */
  decoded: string;
  inFieldResult: boolean;
}

export interface RunInfo {
  /** nearest enclosing revision element (any author), or null */
  rev: XElement | null;
  /** nearest enclosing w:hyperlink / w:smartTag, or null */
  structure: XElement | null;
  /** a w:sdt sits strictly between the run and its paragraph */
  inlineSdt: boolean;
}

export interface SpanMap {
  text: string;
  pieces: SpanPiece[];
  runInfo: Map<XElement, RunInfo>;
}

export function buildSpanMap(p: XElement, view: View): SpanMap {
  const pieces: SpanPiece[] = [];
  const runInfo = new Map<XElement, RunInfo>();
  let pos = 0;
  walkContent(p, view, {
    onCharEl: (el, kind, inFieldResult) => {
      const run = el.parent;
      if (!run || run.kind !== "el") return;
      if (!runInfo.has(run)) runInfo.set(run, ancestry(run, p));
      const decoded = kind === "text" ? wordText(el) : "";
      const len = kind === "text" ? decoded.length : 1;
      pieces.push({ run, el, kind, start: pos, end: pos + len, decoded, inFieldResult });
      pos += len;
    },
  });
  return { text: pieces.map((pc) => (pc.kind === "text" ? pc.decoded : pc.kind === "tab" ? "\t" : pc.kind === "break" ? "\n" : "-")).join(""), pieces, runInfo };
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
  return { rev, structure, inlineSdt };
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
 * crosses_structure, in_field, in_content_control — a per-piece first-hit
 * would misreport (e.g. a span crossing into a field result inside a
 * hyperlink must refuse crosses_structure, not in_field).
 */
export function checkSpan(map: SpanMap, start: number, end: number, priorOps: Array<{ start: number; end: number }>, ownedRevisions: ReadonlySet<XElement> = new Set()): SpanIssue | null {
  const covered = map.pieces.filter((pc) => pc.start < end && pc.end > start);
  let hasRev = false;
  let hasField = false;
  let hasSdt = false;
  const structures = new Set<XElement | null>();
  const parents = new Set<XElement | null>();
  for (const pc of covered) {
    const info = map.runInfo.get(pc.run)!;
    if (info.rev) hasRev = true;
    if (pc.inFieldResult) hasField = true;
    if (info.inlineSdt) hasSdt = true;
    structures.add(info.structure);
    parents.add(pc.run.parent);
  }
  if (hasRev) return { reason: "span_overlaps_revision" };
  for (const pr of priorOps) {
    if (start < pr.end && pr.start < end) return { reason: "span_overlaps_prior_op" };
  }
  // mutation safety: every covered run must share one parent (an unmodelled
  // inline wrapper such as w:customXml would otherwise be partially deleted)
  if (parents.size > 1 || structures.size > 1) return { reason: "span_crosses_structure" };
  if (hasField) return { reason: "span_in_field" };
  if (hasSdt) return { reason: "span_in_content_control" };
  if (start === end) {
    // point insertion: the field/sdt protections still apply at the point, and
    // the surrounding run's revision is refused unless THIS call created it
    // (those get sibling placement instead)
    const pc = covered[0] ?? pieceAt(map, start);
    if (pc && start >= pc.start && start < pc.end) {
      const info = map.runInfo.get(pc.run)!;
      if (info.rev && !ownedRevisions.has(info.rev)) return { reason: "span_overlaps_revision" };
      if (pc.inFieldResult) return { reason: "span_in_field" };
      if (info.inlineSdt) return { reason: "span_in_content_control" };
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
