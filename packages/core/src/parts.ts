// WordprocessingML part enumeration: the main part plus every related story
// part (headers, footers, footnotes, endnotes, comments, glossary). Revision
// ids and paragraph ids must be seeded across ALL of them.
import { decodePart, kid, parseXml, relatedParts, type PackageReader, type XElement } from "@stitchdocx/package-io";
import { walkAll } from "./revisions.js";

const STORY_TYPES = ["header", "footer", "footnotes", "endnotes", "comments", "glossaryDocument"];

/** every WordprocessingML part of the document, main part first */
export function wordprocessingPartNames(reader: PackageReader, mainPart: string): string[] {
  const out = [mainPart];
  for (const t of STORY_TYPES) {
    for (const p of relatedParts(reader, mainPart, t)) {
      if (reader.has(p) && !out.includes(p)) out.push(p);
    }
  }
  return out;
}

/** parsed root of one part, or null when the part is missing or not w: XML */
export function partRoot(reader: PackageReader, name: string): XElement | null {
  const bytes = reader.part(name);
  if (!bytes) return null;
  const tree = parseXml(decodePart(bytes));
  const root = kid(tree as unknown as XElement);
  return root ?? null;
}

/** every w14:paraId attribute value used anywhere in the document's parts.
 *  `mainParagraphs` (the caller's parsed body paragraphs) avoids re-parsing
 *  the main part — the scan runs on every open. */
export function paraIdsAcrossParts(reader: PackageReader, mainPart: string, mainParagraphs?: XElement[]): string[] {
  const out: string[] = [];
  for (const name of wordprocessingPartNames(reader, mainPart)) {
    if (name === mainPart && mainParagraphs) {
      for (const p of mainParagraphs) {
        const id = p.attrs.find((a) => a.name === "w14:paraId");
        if (id) out.push(id.raw);
      }
      continue;
    }
    const root = partRoot(reader, name);
    if (!root) continue;
    for (const e of walkAll(root)) {
      if (e.name !== "w:p") continue;
      const id = e.attrs.find((a) => a.name === "w14:paraId");
      if (id) out.push(id.raw);
    }
  }
  return out;
}

/** parsed roots of every related (non-main) WordprocessingML part */
export function otherPartRoots(reader: PackageReader, mainPart: string): XElement[] {
  const out: XElement[] = [];
  for (const name of wordprocessingPartNames(reader, mainPart)) {
    if (name === mainPart) continue;
    const root = partRoot(reader, name);
    if (root) out.push(root);
  }
  return out;
}
