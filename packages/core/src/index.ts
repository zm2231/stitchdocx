// stitchdocx core: the public engine API
export { openDocument, normalizeDocx, finalizeDocx, acceptAllDocx, Document, isIsoDate } from "./document.js";
export { verifyOutput } from "./verify.js";
export { DOCUMENT_LIMITS } from "./limits.js";
export { PackageLimitError } from "@stitchdocx/package-io";
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
