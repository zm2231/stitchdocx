// resolveSpan: hardened string anchoring. The caller names a paragraph (by
// address, by a unique containing substring, or by index) and quotes the exact
// span; the resolver returns exact offsets in that paragraph's projected text,
// or a named refusal with candidates. Never guesses.
import type { Paragraph } from "./types.js";

export type SpanTarget =
  | { addr: string; span: string; occurrence?: number }
  | { containing: string; nth?: number; span: string; occurrence?: number }
  | { index: number; span?: string; occurrence?: number };

export type SpanResolution =
  | { ok: true; addr: string; start: number; end: number }
  | {
      ok: false;
      reason: "addr_not_found" | "ambiguous_match" | "span_not_found" | "span_out_of_range" | "span_text_mismatch" | "paragraph_not_editable";
      candidates?: Array<{ addr: string; start: number; end: number }>;
    };

export function resolveSpan(projection: Paragraph[], target: SpanTarget): SpanResolution {
  let addr: string | null = null;
  if ("addr" in target) {
    addr = target.addr;
    if (!projection.some((p) => p.addr === addr)) return { ok: false, reason: "addr_not_found" };
  } else if ("containing" in target) {
    const hits = projection.filter((p) => p.text.includes(target.containing));
    const nth = target.nth ?? (hits.length === 1 ? 1 : null);
    if (nth === null || !hits[nth - 1]) {
      return { ok: false, reason: hits.length === 0 ? "addr_not_found" : "ambiguous_match", candidates: hits.map((h) => ({ addr: h.addr, start: 0, end: 0 })) };
    }
    addr = hits[nth - 1]!.addr;
  } else {
    const p = projection[target.index];
    if (!p) return { ok: false, reason: "addr_not_found" };
    addr = p.addr;
  }
  const para = projection.find((p) => p.addr === addr)!;
  // a record whose text spans a removed paragraph mark is a merged view: its
  // offsets cover more than the one paragraph an op may edit, so it must not
  // produce operations
  if (!para.editable) return { ok: false, reason: "paragraph_not_editable" };
  const span = target.span ?? "";
  if (!span) return { ok: true, addr: addr!, start: 0, end: 0 };
  const starts: number[] = [];
  for (let i = para.text.indexOf(span); i >= 0; i = para.text.indexOf(span, i + 1)) starts.push(i);
  const occ = target.occurrence ?? (starts.length === 1 ? 1 : null);
  if (occ === null || starts[occ - 1] === undefined) {
    return { ok: false, reason: starts.length === 0 ? "span_not_found" : "ambiguous_match", candidates: starts.map((s) => ({ addr: addr!, start: s, end: s + span.length })) };
  }
  const start = starts[occ - 1]!;
  return { ok: true, addr: addr!, start, end: start + span.length };
}
