// Body traversal in document order: tables cell by cell (row-major), block
// content controls and customXml descended into. The paragraph registry is
// view-independent (addressing);
// the projection (views, mark merging) lives in projection.ts.
import { kid, kids, type XElement } from "@stitchdocx/package-io";

export interface BodyContext {
  cell: boolean;
  /** top-level table index with row/col — nested tables reuse the outer index */
  table?: { index: number; row: number; col: number };
  sdt: boolean;
}

export interface BodyWalk {
  /** every w:p element in the main body, reader order */
  paragraphs: XElement[];
  /** per-paragraph context, same order */
  contexts: BodyContext[];
  /** w:body element */
  body: XElement;
}

export function walkBody(documentRoot: XElement): BodyWalk {
  const body = kid(documentRoot, "w:body");
  if (!body) throw new Error("word/document.xml has no w:body");
  const paragraphs: XElement[] = [];
  const contexts: BodyContext[] = [];
  let tableIndex = -1;

  const visit = (node: XElement, ctx: BodyContext): void => {
    for (const c of kids(node)) {
      if (c.name === "w:p") {
        paragraphs.push(c);
        contexts.push(ctx);
      } else if (c.name === "w:tbl") {
        const t = ctx.table ? ctx.table : { index: ++tableIndex };
        kids(c, "w:tr").forEach((tr, r) => {
          kids(tr, "w:tc").forEach((tc, col) => visit(tc, { ...ctx, cell: true, table: { index: t.index, row: r, col } }));
        });
      } else if (c.name === "w:sdt") {
        const content = kid(c, "w:sdtContent");
        if (content) visit(content, { ...ctx, sdt: true });
      } else if (c.name === "w:customXml") {
        visit(c, ctx);
      }
    }
  };
  visit(body, { cell: false, sdt: false });
  return { paragraphs, contexts, body };
}

/** nearest ancestor satisfying pred, stopping at (and including) the element itself */
export function closest(el: XElement, pred: (e: XElement) => boolean, stopAt?: XElement): XElement | null {
  let cur: XElement | null = el;
  while (cur) {
    if (pred(cur)) return cur;
    if (stopAt && cur === stopAt) return null;
    cur = cur.parent;
  }
  return null;
}
