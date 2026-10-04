// normalize: assign w14:paraId to every body paragraph lacking (or duplicating)
// one. Idempotent and deterministic; a document that needs nothing comes back
// byte-for-byte unchanged.
//
// Ids are unique across EVERY WordprocessingML part (a comment or header
// paragraph's id counts), and ids that are not 8 hex digits below 0x80000000
// are reassigned wherever they appear. No structural repair happens here: the
// parser is total (it opens anything well-formed, including an empty
// body-level w:sectPr that other runtimes choke on), so a rewrite would only
// break the untouched-bytes guarantee.
import { decodePart, encodePart, getAttr, kid, mainPartName, openPackage, parseXml, relatedPart, serializeXml, setAttr, writePackage, type PackageReader, type XElement } from "@stitchdocx/package-io";
import { walkAll } from "./revisions.js";
import { walkBody } from "./body.js";
import { ParaIdAllocator } from "./revisions.js";
import { paraIdsAcrossParts, partRoot, wordprocessingPartNames } from "./parts.js";

export const W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";

/** the editing restriction Word enforces (word/settings.xml), or null. Only an
 *  enforced w:documentProtection with a known mode counts; the document is
 *  never modified here — callers surface the restriction to the user. */
export function readProtection(reader: PackageReader, mainPart: string): string | null {
  const settingsPart = relatedPart(reader, mainPart, "settings") ?? "word/settings.xml";
  const settingsXml = reader.part(settingsPart);
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

/** per-part paragraph id repair plan: which paragraphs take a fresh id */
function planPart(paragraphs: XElement[], globalSeen: Set<string>): XElement[] {
  const reassign: XElement[] = [];
  for (const p of paragraphs) {
    const id = getAttr(p, "w14:paraId");
    if (id === null) continue;
    if (!ParaIdAllocator.valid(id)) {
      reassign.push(p);
      continue;
    }
    const key = id.toUpperCase();
    if (globalSeen.has(key)) {
      reassign.push(p);
      continue;
    }
    globalSeen.add(key);
  }
  return reassign;
}

function partParagraphs(root: XElement): XElement[] {
  return [...walkAll(root)].filter((e) => e.name === "w:p");
}

/** declare the w14 namespace when a part gains paraId attributes */
function ensureW14Ns(rootEl: XElement | null): void {
  if (!rootEl) return;
  if (!rootEl.attrs.some((a) => a.name === "xmlns:w14")) setAttr(rootEl, "xmlns:w14", W14_NS);
}

export function normalizeDocx(bytes: Uint8Array): { docx: Uint8Array; assigned: number; protection: string | null } {
  const reader = openPackage(bytes);
  const mainPart = mainPartName(reader);
  const docXml = reader.part(mainPart);
  if (!docXml) throw new Error(`${mainPart} missing`);
  const tree = parseXml(decodePart(docXml));
  const root = kid(tree as unknown as XElement, "w:document");
  if (!root) throw new Error("word/document.xml has no w:document root");
  const walk = walkBody(root);

  // seed the allocator with every id used anywhere, then repair: duplicates
  // (first occurrence keeps the id) and ids that are not 8 hex digits below
  // 0x80000000 take fresh ids, in every part
  const alloc = new ParaIdAllocator(paraIdsAcrossParts(reader, mainPart, walk.paragraphs));
  const globalSeen = new Set<string>();
  const mainReassign = planPart(walk.paragraphs, globalSeen);
  // parse each other part once; plan and repair run on the same tree
  const otherTrees = new Map<string, XElement>();
  for (const part of wordprocessingPartNames(reader, mainPart).slice(1)) {
    const partRootEl = partRoot(reader, part);
    if (partRootEl) otherTrees.set(part, partRootEl);
  }
  const otherPlans = new Map<string, { root: XElement; reassign: XElement[] }>();
  for (const [part, partRootEl] of otherTrees) {
    const reassign = planPart(partParagraphs(partRootEl), globalSeen);
    if (reassign.length > 0) otherPlans.set(part, { root: partRootEl, reassign });
  }

  let assigned = 0;
  const reassign = (p: XElement): void => {
    setAttr(p, "w14:paraId", alloc.take());
    assigned++;
  };
  for (const p of mainReassign) reassign(p);
  let bodyNew = 0;
  for (const p of walk.paragraphs) {
    if (getAttr(p, "w14:paraId") !== null) continue;
    reassign(p);
    bodyNew++;
  }

  const changedParts = new Map<string, string>();
  if (mainReassign.length > 0 || bodyNew > 0) {
    ensureW14Ns(root);
    changedParts.set(mainPart, serializeXml(tree));
  }
  for (const [part, { root: partRootEl, reassign: list }] of otherPlans) {
    for (const p of list) reassign(p);
    ensureW14Ns(partRootEl);
    changedParts.set(part, serializeXml(partRootEl));
  }

  const protection = readProtection(reader, mainPart);
  if (changedParts.size === 0) return { docx: bytes, assigned: 0, protection };
  const entries = reader.readAll((name) => name === mainPart || changedParts.has(name));
  for (const [name, xml] of changedParts) entries.set(name, encodePart(xml));
  return { docx: writePackage(entries), assigned, protection };
}
