// stitchdocx core — public engine API (see research/04-engine-api-adjudication.md)
export { openDocument, normalizeDocx, finalizeDocx, Document, isIsoDate } from "./document.js";
export { verifyOutput } from "./verify.js";
export { buildSpanMap, checkSpan } from "./spanmap.js";
export { resolveSpan } from "./resolve.js";
export type { SpanTarget, SpanResolution } from "./resolve.js";
export type {
  Addr,
  Paragraph,
  Op,
  RefusalReason,
  Receipt,
  Decision,
  DecisionResult,
} from "./types.js";
export type { View } from "./projection.js";
