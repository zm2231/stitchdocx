// WordprocessingML part enumeration: the main part plus every related story
// part (headers, footers, footnotes, endnotes, comments, glossary). Revision
// ids and paragraph ids must be seeded across ALL of them. The seeds come from
// a tag scan of each part's source; no part is held parsed for them.
import { decodeAttrValue, decodePart, relatedParts, scanStartTags, type PackageReader } from "@stitchdocx/package-io";
import { REVISION_ID_NAMES } from "./revisions.js";

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

/** each part's source in turn, the main part's taken from the caller */
function* partSources(reader: PackageReader, mainPart: string, mainSource: string): Generator<string> {
  for (const name of wordprocessingPartNames(reader, mainPart)) {
    if (name === mainPart) {
      yield mainSource;
      continue;
    }
    const bytes = reader.part(name);
    if (bytes) yield decodePart(bytes);
  }
}

/** every w14:paraId attribute value on any w:p in any of the document's parts */
export function paraIdsAcrossParts(reader: PackageReader, mainPart: string, mainSource: string): string[] {
  const out: string[] = [];
  for (const source of partSources(reader, mainPart, mainSource)) {
    scanStartTags(source, (name) => name === "w:p", (_, attrs) => {
      const id = attrs.find((a) => a.name === "w14:paraId");
      if (id) out.push(id.raw);
    });
  }
  return out;
}

/** the largest w:id on any revision-type element in any of the document's parts */
export function maxRevisionIdAcrossParts(reader: PackageReader, mainPart: string, mainSource: string): number {
  let max = 0;
  for (const source of partSources(reader, mainPart, mainSource)) {
    scanStartTags(source, (name) => REVISION_ID_NAMES.has(name), (_, attrs) => {
      const id = attrs.find((a) => a.name === "w:id");
      const n = id ? Number(decodeAttrValue(id.raw)) : NaN;
      if (Number.isInteger(n) && n > max) max = n;
    });
  }
  return max;
}
