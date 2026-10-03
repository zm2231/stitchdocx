// Span map: the paragraph's visible text tiled into char pieces, each tied to
// the run that produced it, with the ancestry metadata the refusal checks need
// (revision, hyperlink/smartTag structure, inline content control, field).
import { decodeEntities, type XElement } from "@stitchdocx/package-io";
import { walkContent, type CharKind, type View } from "./projection.js";

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
      const decoded = kind === "text" ? textOfCharEl(el) : "";
      const len = kind === "text" ? decoded.length : 1;
      pieces.push({ run, el, kind, start: pos, end: pos + len, decoded, inFieldResult });
      pos += len;
    },
  });
  return { text: pieces.map((pc) => (pc.kind === "text" ? pc.decoded : pc.kind === "tab" ? "\t" : pc.kind === "break" ? "\n" : "-")).join(""), pieces, runInfo };
}

function textOfCharEl(el: XElement): string {
  let out = "";
  for (const c of el.children) {
    if (c.kind === "text") out += decodeEntities(c.raw); // decoded: offsets live in text space
    else if (c.kind === "raw" && c.raw.startsWith("<![CDATA[")) out += c.raw.slice(9, -3);
  }
  return out;
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
 * Refusal checks over [start,end) in paragraph-text coordinates. Check order
 * is contract-relevant: a span that both overlaps a prior op and crosses a
 * structure is reported as span_overlaps_prior_op (S12/S17).
 */
export function checkSpan(map: SpanMap, start: number, end: number, priorOps: Array<{ start: number; end: number }>): SpanIssue | null {
  // 1. overlap with earlier ops in the same call (original coordinates)
  for (const pr of priorOps) {
    if (start < pr.end && pr.start < end) return { reason: "span_overlaps_prior_op" };
  }
  if (start === end) {
    // pure insertion at a point: refused only where even the surrounding run is untouchable
    const pc = pieceAt(map, start);
    if (pc) {
      const info = map.runInfo.get(pc.run)!;
      if (pc.inFieldResult) return { reason: "span_in_field" };
      if (info.inlineSdt) return { reason: "span_in_content_control" };
    }
    return null;
  }
  // 2. any covered char inside a tracked revision (client's or ours)
  let structure: XElement | null | undefined = undefined;
  for (const pc of map.pieces) {
    if (pc.end <= start || pc.start >= end) continue;
    const info = map.runInfo.get(pc.run)!;
    if (info.rev) return { reason: "span_overlaps_revision" };
    if (pc.inFieldResult) return { reason: "span_in_field" };
    if (info.inlineSdt) return { reason: "span_in_content_control" };
    if (structure === undefined) structure = info.structure;
    else if (structure !== info.structure) return { reason: "span_crosses_structure" };
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
