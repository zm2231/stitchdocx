// Body traversal in document order: tables cell by cell (row-major), block
// content controls and customXml descended into — including around rows and
// cells (Word wraps both in w:sdt/w:customXml). The paragraph registry is
// view-independent (addressing); the projection (views, mark merging) lives in
// projection.ts. Each paragraph's complex-field cursor state at entry is
// recorded so per-paragraph walks (span maps) agree with the whole-story walk.
import { kid, kids, type XElement } from "@stitchdocx/package-io";
import { stepFields } from "./fields.js";

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
  /** per-paragraph complex-field cursor state at entry, same order */
  fieldStates: Map<XElement, string[]>;
  /** w:body element */
  body: XElement;
}

export function walkBody(documentRoot: XElement): BodyWalk {
  const body = kid(documentRoot, "w:body");
  if (!body) throw new Error("word/document.xml has no w:body");
  const paragraphs: XElement[] = [];
  const contexts: BodyContext[] = [];
  const fieldStates = new Map<XElement, string[]>();
  const fields: string[] = [];
  let tableIndex = -1;

  const visit = (node: XElement, ctx: BodyContext): void => {
    for (const c of kids(node)) {
      if (c.name === "w:p") {
        fieldStates.set(c, [...fields]);
        advance(c);
        paragraphs.push(c);
        contexts.push(ctx);
      } else if (c.name === "w:tbl") {
        const t = ctx.table ? ctx.table : { index: ++tableIndex };
        let r = 0;
        for (const tr of rowsOf(c)) {
          let col = 0;
          for (const tc of cellsOf(tr)) {
            visit(tc, { ...ctx, cell: true, table: { index: t.index, row: r, col } });
            col++;
          }
          r++;
        }
      } else if (c.name === "w:sdt") {
        const content = kid(c, "w:sdtContent");
        if (content) visit(content, { ...ctx, sdt: true });
      } else if (c.name === "w:customXml") {
        visit(c, ctx);
      } else {
        advance(c);
      }
    }
  };

  /** advance the shared field cursor over a non-paragraph block */
  const advance = (node: XElement): void => {
    const walk = (n: XElement): void => {
      for (const c of kids(n)) {
        stepFields(fields, c);
        walk(c);
      }
    };
    walk(node);
  };

  visit(body, { cell: false, sdt: false });
  return { paragraphs, contexts, fieldStates, body };
}

/** the rows of a table: direct w:tr children plus rows wrapped in block
 *  content controls or customXml */
export function rowsOf(tbl: XElement): XElement[] {
  const out: XElement[] = [];
  const collect = (node: XElement): void => {
    for (const c of kids(node)) {
      if (c.name === "w:tr") out.push(c);
      else if (c.name === "w:sdt") {
        const content = kid(c, "w:sdtContent");
        if (content) collect(content);
      } else if (c.name === "w:customXml") collect(c);
    }
  };
  collect(tbl);
  return out;
}

/** the cells of a row: direct w:tc children plus cells wrapped in content
 *  controls or customXml */
export function cellsOf(tr: XElement): XElement[] {
  const out: XElement[] = [];
  const collect = (node: XElement): void => {
    for (const c of kids(node)) {
      if (c.name === "w:tc") out.push(c);
      else if (c.name === "w:sdt") {
        const content = kid(c, "w:sdtContent");
        if (content) collect(content);
      } else if (c.name === "w:customXml") collect(c);
    }
  };
  collect(tr);
  return out;
}

/** the row-level tracked change on a row (w:trPr/w:del or w:trPr/w:ins) */
export function rowRevision(tr: XElement): XElement | null {
  const trPr = kid(tr, "w:trPr");
  if (!trPr) return null;
  return kids(trPr).find((c) => c.name === "w:del" || c.name === "w:ins") ?? null;
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
