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
import { decodeAttrValue, decodePart, encodePart, getAttr, kid, mainPartName, openPackage, relatedPart, scanStartTags, serializeXml, setAttr, writePackage, type PackageReader, type XElement } from "@stitchdocx/package-io";
import { parseMainPart, walkBody } from "./body.js";
import { ParaIdAllocator } from "./revisions.js";
import { paraIdsAcrossParts, wordprocessingPartNames } from "./parts.js";
import { mainPartSource, parsePart } from "./limits.js";

export const W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";

/** the editing restriction Word enforces (word/settings.xml), or null. Only an
 *  enforced w:documentProtection with a known mode counts; the document is
 *  never modified here — callers surface the restriction to the user. */
export function readProtection(reader: PackageReader, mainPart: string): string | null {
  const settingsPart = relatedPart(reader, mainPart, "settings") ?? "word/settings.xml";
  const settingsXml = reader.xml(settingsPart);
  if (settingsXml === undefined) return null;
  const tree = parsePart(settingsXml);
  const settings = kid(tree as unknown as XElement, "w:settings");
  if (!settings) return null;
  const dp = kid(settings, "w:documentProtection");
  if (!dp) return null;
  const enforcement = getAttr(dp, "w:enforcement");
  if (enforcement === null || !["1", "true", "on"].includes(enforcement)) return null;
  const edit = getAttr(dp, "w:edit");
  return edit !== null && ["readOnly", "comments", "trackedChanges", "forms"].includes(edit) ? edit : null;
}

/** per-part paragraph id repair plan: given each paragraph's id in document
 *  order, the positions whose paragraph takes a fresh id */
function planPart(ids: (string | null)[], globalSeen: Set<string>): number[] {
  const reassign: number[] = [];
  ids.forEach((id, i) => {
    if (id === null) return;
    if (!ParaIdAllocator.valid(id)) {
      reassign.push(i);
      return;
    }
    const key = id.toUpperCase();
    if (globalSeen.has(key)) {
      reassign.push(i);
      return;
    }
    globalSeen.add(key);
  });
  return reassign;
}

/** each w:p's w14:paraId in a part's source, document order, from a tag scan */
function scannedParaIds(source: string): (string | null)[] {
  const ids: (string | null)[] = [];
  scanStartTags(source, (name) => name === "w:p", (_, attrs) => {
    const id = attrs.find((a) => a.name === "w14:paraId");
    ids.push(id ? decodeAttrValue(id.raw) : null);
  });
  return ids;
}

/** every w:p under `root`, document order */
function paragraphsInOrder(root: XElement, out: XElement[] = []): XElement[] {
  for (const c of root.children) {
    if (c.kind !== "el") continue;
    if (c.name === "w:p") out.push(c);
    paragraphsInOrder(c, out);
  }
  return out;
}

/** declare the w14 namespace when a part gains paraId attributes */
function ensureW14Ns(rootEl: XElement | null): void {
  if (!rootEl) return;
  if (!rootEl.attrs.some((a) => a.name === "xmlns:w14")) setAttr(rootEl, "xmlns:w14", W14_NS);
}

export function normalizeDocx(bytes: Uint8Array): { docx: Uint8Array; assigned: number; protection: string | null } {
  const reader = openPackage(bytes);
  const mainPart = mainPartName(reader);
  const source = mainPartSource(reader, mainPart);
  // paragraph content is never read here: ids sit on the w:p start tags
  const { tree, root } = parseMainPart(source);
  const walk = walkBody(root);

  // seed the allocator with every id used anywhere, then repair: duplicates
  // (first occurrence keeps the id) and ids that are not 8 hex digits below
  // 0x80000000 take fresh ids, in every part
  const alloc = new ParaIdAllocator(paraIdsAcrossParts(reader, mainPart, source));
  const globalSeen = new Set<string>();
  const mainReassign = planPart(walk.paragraphs.map((p) => getAttr(p, "w14:paraId")), globalSeen).map((i) => walk.paragraphs[i]!);
  // the other parts are planned from a tag scan; only a part that needs a
  // repair is parsed, one at a time, after the main part's ids are assigned
  const otherPlans = new Map<string, { count: number; positions: number[] }>();
  for (const part of wordprocessingPartNames(reader, mainPart).slice(1)) {
    const source = reader.xml(part);
    if (source === undefined) continue;
    const ids = scannedParaIds(source);
    const positions = planPart(ids, globalSeen);
    if (positions.length > 0) otherPlans.set(part, { count: ids.length, positions });
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
  for (const [part, { count, positions }] of otherPlans) {
    const tree = parsePart(reader.xml(part)!);
    const partRootEl = kid(tree as unknown as XElement);
    if (!partRootEl) throw new Error(`${part} has no root element`);
    const paragraphs = paragraphsInOrder(partRootEl);
    if (paragraphs.length !== count) throw new Error(`${part}: the scan found ${count} paragraphs, the parse ${paragraphs.length}`);
    for (const i of positions) reassign(paragraphs[i]!);
    ensureW14Ns(partRootEl);
    changedParts.set(part, serializeXml(tree));
  }

  const protection = readProtection(reader, mainPart);
  if (changedParts.size === 0) return { docx: bytes, assigned: 0, protection };
  const entries = reader.readAll();
  for (const [name, xml] of changedParts) entries.set(name, encodePart(xml));
  return { docx: writePackage(entries), assigned, protection };
}
