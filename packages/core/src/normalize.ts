// normalize: assign w14:paraId to every body paragraph lacking (or duplicating)
// one. Idempotent and deterministic; a document that needs nothing comes back
// byte-for-byte unchanged.
//
// No structural repair happens here: the parser is total (it opens anything
// well-formed, including an empty body-level w:sectPr that other runtimes
// choke on), so a rewrite would only break the untouched-bytes guarantee.
import { getAttr, kid, mainPartName, readPackage, relatedPart, setAttr, writePackage, encodePart, decodePart, type Entries, type XElement } from "@stitchdocx/package-io";
import { parseXml, serializeXml } from "@stitchdocx/package-io";
import { walkBody } from "./body.js";
import { ParaIdAllocator } from "./revisions.js";

export const W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";

/** the editing restriction Word enforces (word/settings.xml), or null. Only an
 *  enforced w:documentProtection with a known mode counts; the document is
 *  never modified here — callers surface the restriction to the user. */
export function readProtection(entries: Entries): string | null {
  const settingsPart = relatedPart(entries, mainPartName(entries), "settings") ?? "word/settings.xml";
  const settingsXml = entries.get(settingsPart);
  if (!settingsXml) return null;
  const tree = parseXml(decodePart(settingsXml));
  const settings = kid(tree as unknown as XElement, "w:settings");
  if (!settings) return null;
  const dp = kid(settings, "w:documentProtection");
  if (!dp) return null;
  const enforcement = getAttr(dp, "w:enforcement");
  if (enforcement === null || !["1", "true", "on"].includes(enforcement)) return null;
  const edit = getAttr(dp, "w:edit");
  return edit !== null && ["readOnly", "comments", "trackedChanges", "forms"].includes(edit) ? edit : null;
}

export function normalizeDocx(bytes: Uint8Array): { docx: Uint8Array; assigned: number; protection: string | null } {
  const entries: Entries = readPackage(bytes);
  const mainPart = mainPartName(entries);
  const docXml = entries.get(mainPart);
  if (!docXml) throw new Error(`${mainPart} missing`);
  const tree = parseXml(decodePart(docXml));
  const root = kid(tree as unknown as XElement, "w:document");
  if (!root) throw new Error("word/document.xml has no w:document root");
  const walk = walkBody(root);

  // paraIds: unique, deterministic, < 0x80000000. Duplicates are repaired:
  // the first occurrence keeps the id, later ones get fresh ids (document order).
  const alloc = new ParaIdAllocator(walk.paragraphs);
  let assigned = 0;
  const seen = new Set<string>();
  const reassign: XElement[] = [];
  for (const p of walk.paragraphs) {
    const id = getAttr(p, "w14:paraId");
    if (id === null) continue;
    const key = id.toUpperCase();
    if (seen.has(key)) reassign.push(p);
    else seen.add(key);
  }
  for (const p of walk.paragraphs) {
    if (getAttr(p, "w14:paraId") !== null && !reassign.includes(p)) continue;
    setAttr(p, "w14:paraId", alloc.take());
    assigned++;
  }
  if (assigned > 0) {
    const hasNs = root.attrs.some((a) => a.name === "xmlns:w14");
    if (!hasNs) setAttr(root, "xmlns:w14", W14_NS);
  }
  if (assigned > 0) {
    entries.set(mainPart, encodePart(serializeXml(tree)));
    return { docx: writePackage(entries), assigned, protection: readProtection(entries) };
  }
  return { docx: bytes, assigned: 0, protection: readProtection(entries) };
}
