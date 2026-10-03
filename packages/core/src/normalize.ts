// normalize: assign w14:paraId to every body paragraph lacking (or duplicating)
// one. Idempotent and deterministic; a document that needs nothing comes back
// byte-for-byte unchanged.
//
// No structural repair happens here: the parser is total (it opens anything
// well-formed, including an empty body-level w:sectPr that other runtimes
// choke on), so a rewrite would only break the untouched-bytes guarantee.
import { getAttr, kid, readPackage, setAttr, writePackage, encodePart, decodePart, type Entries, type XElement } from "@stitchdocx/package-io";
import { parseXml, serializeXml } from "@stitchdocx/package-io";
import { walkBody } from "./body.js";
import { ParaIdAllocator } from "./revisions.js";

export const W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";

export function normalizeDocx(bytes: Uint8Array): { docx: Uint8Array; assigned: number } {
  const entries: Entries = readPackage(bytes);
  const docXml = entries.get("word/document.xml");
  if (!docXml) throw new Error("word/document.xml missing");
  const tree = parseXml(decodePart(docXml));
  const root = kid(tree as unknown as XElement, "w:document");
  if (!root) throw new Error("word/document.xml has no w:document root");
  const walk = walkBody(root);

  // paraIds: unique, deterministic, < 0x80000000
  const alloc = new ParaIdAllocator(walk.paragraphs);
  let assigned = 0;
  for (const p of walk.paragraphs) {
    const id = getAttr(p, "w14:paraId");
    if (id !== null) continue;
    setAttr(p, "w14:paraId", alloc.take());
    assigned++;
  }
  if (assigned > 0) {
    const hasNs = root.attrs.some((a) => a.name === "xmlns:w14");
    if (!hasNs) setAttr(root, "xmlns:w14", W14_NS);
  }
  if (assigned > 0) {
    entries.set("word/document.xml", encodePart(serializeXml(tree)));
    return { docx: writePackage(entries), assigned };
  }
  return { docx: bytes, assigned: 0 };
}
