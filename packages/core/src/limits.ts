// Memory bounds. Every XML part, the main part included, is bounded in bytes
// before it inflates (PACKAGE_LIMITS.maxXmlPartBytes). Every parse is bounded in nodes, because bytes do not bound
// a tree: a small part can hold a dense grid of tiny cells. The main part's
// block structure and every paragraph parsed in place draw on one budget per
// open document, and its paragraphs carry per-paragraph state besides, so its
// node limit is the tighter one; a transient read of a paragraph gets a budget
// of the same limit, and any other part one of PACKAGE_LIMITS.maxPartNodes. A
// document over a limit is refused as document_too_large, so the caller can
// route it elsewhere.
import { NodeBudget, PACKAGE_LIMITS, parseXml, type PackageReader, type XDoc } from "@stitchdocx/package-io";

export const DOCUMENT_LIMITS = {
  /** nodes of the main part's block structure plus its paragraphs parsed in place */
  maxMainPartNodes: 40_000,
};

/** the main part's budget, shared by its block structure and every paragraph
 *  parsed in place for one open document */
export function mainPartBudget(): NodeBudget {
  return new NodeBudget(DOCUMENT_LIMITS.maxMainPartNodes);
}

/** a fully parsed part (styles, numbering, settings, story parts), bounded */
export function parsePart(source: string): XDoc {
  return parseXml(source, { budget: new NodeBudget(PACKAGE_LIMITS.maxPartNodes) });
}

/** the main part's decoded source; refuses document_too_large over the limit */
export function mainPartSource(reader: PackageReader, mainPart: string): string {
  const source = reader.xml(mainPart);
  if (source === undefined) throw new Error(`${mainPart} missing`);
  return source;
}
