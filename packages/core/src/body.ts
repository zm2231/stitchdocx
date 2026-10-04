// Body traversal in document order: tables cell by cell (row-major), block
// content controls and customXml descended into — including around rows and
// cells (Word wraps both in w:sdt/w:customXml). The paragraph registry is
// view-independent (addressing); the projection (views, mark merging) lives in
// projection.ts. Each paragraph's complex-field cursor state at entry is
// recorded so per-paragraph walks (span maps) agree with the whole-story walk.
import { isPending, kid, kids, parseXml, pendingSource, readView, type XDoc, type XElement } from "@stitchdocx/package-io";
import { stepFields } from "./fields.js";
import { mainPartBudget } from "./limits.js";

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
  /** complex-field cursor state at entry, for each paragraph entered inside
   *  a field (any other paragraph enters with none) */
  fieldStates: Map<XElement, string[]>;
  /** w:body element */
  body: XElement;
}

/** the elements that give the main part its block structure; parseMainPart
 *  parses only these, and every other element keeps its content as source */
const BLOCK_STRUCTURE = new Set(["w:document", "w:body", "w:tbl", "w:tr", "w:tc", "w:sdt", "w:sdtContent", "w:customXml"]);

/** the main part with only its block structure parsed: paragraphs and
 *  property subtrees are lazy (package-io LazyElement) */
export function parseMainPart(source: string): { tree: XDoc; root: XElement } {
  const tree = parseXml(source, { lazy: (name) => !BLOCK_STRUCTURE.has(name), budget: mainPartBudget() });
  const root = kid(tree as unknown as XElement, "w:document");
  if (!root) throw new Error("the main document part has no w:document root");
  return { tree, root };
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
        if (fields.length > 0) fieldStates.set(c, [...fields]);
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

  /** advance the shared field cursor over a block; an unparsed paragraph is
   *  read through a transient view, and only when it holds a fldChar */
  const advance = (node: XElement): void => {
    const walk = (n: XElement): void => {
      for (const c of kids(n)) {
        stepFields(fields, c);
        if (!isPending(c)) walk(c);
        else if (pendingSource(c)!.includes("fldChar")) walk(readView(c));
      }
    };
    if (!isPending(node)) walk(node);
    else if (pendingSource(node)!.includes("fldChar")) walk(readView(node));
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
  return kids(readView(trPr)).find((c) => c.name === "w:del" || c.name === "w:ins") ?? null;
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
