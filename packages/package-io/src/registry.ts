// Light part registry: [Content_Types].xml and relationship parts. The gate
// never adds or removes parts, so this exists to answer "what is this part"
// questions and to keep door open for later phases (comments, headers).
import { PACKAGE_LIMITS, xmlText, type EntryData } from "./zip.js";
import { decodeEntities, kid, kids, NodeBudget, parseXml, serializeXml, type XElement, type XDoc } from "./xml.js";

export interface ContentTypes {
  defaults: Map<string, string>; // extension -> content type
  overrides: Map<string, string>; // part name -> content type
}

export function parseContentTypes(xml: string): ContentTypes {
  const doc = parseXml(xml, { budget: new NodeBudget(PACKAGE_LIMITS.maxPartNodes) });
  const root = kid(doc as unknown as XElement, "Types");
  const defaults = new Map<string, string>();
  const overrides = new Map<string, string>();
  for (const d of kids(root, "Default")) {
    const ext = d.attrs.find((a) => a.name === "Extension");
    const ct = d.attrs.find((a) => a.name === "ContentType");
    if (ext && ct) defaults.set(decodeEntities(ext.raw), decodeEntities(ct.raw));
  }
  for (const o of kids(root, "Override")) {
    const pn = o.attrs.find((a) => a.name === "PartName");
    const ct = o.attrs.find((a) => a.name === "ContentType");
    if (pn && ct) overrides.set(decodeEntities(pn.raw), decodeEntities(ct.raw));
  }
  return { defaults, overrides };
}

export interface Relationship {
  id: string;
  type: string;
  target: string;
  targetMode?: string;
}

export function parseRels(xml: string): Relationship[] {
  const doc = parseXml(xml, { budget: new NodeBudget(PACKAGE_LIMITS.maxPartNodes) });
  const root = kid(doc as unknown as XElement, "Relationships");
  const out: Relationship[] = [];
  for (const r of kids(root, "Relationship")) {
    const get = (n: string) => {
      const a = r.attrs.find((x) => x.name === n);
      return a ? decodeEntities(a.raw) : undefined;
    };
    const id = get("Id");
    const type = get("Type");
    const target = get("Target");
    if (id && type && target) out.push({ id, type, target, targetMode: get("TargetMode") });
  }
  return out;
}

// Transitional and Strict OOXML name relationship types under different URIs.
const REL_BASES = [
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/",
  "http://purl.oclc.org/ooxml/officeDocument/relationships/",
];

/** The part a relationship of `type` points to from `source` ("" for the
 *  package), resolved per OPC: an absolute target is package-relative, any
 *  other is relative to the source's folder. Null when absent or external. */
export function relatedPart(entries: Parts, source: string, type: string): string | null {
  const slash = source.lastIndexOf("/");
  const dir = source.slice(0, slash + 1);
  const relsXml = entries.get(`${dir}_rels/${source.slice(slash + 1)}.rels`);
  if (!relsXml) return null;
  const rel = parseRels(xmlText(`${dir}_rels/${source.slice(slash + 1)}.rels`, relsXml)).find(
    (r) => r.targetMode !== "External" && REL_BASES.some((ns) => r.type === ns + type),
  );
  if (!rel) return null;
  const parts = (rel.target.startsWith("/") ? rel.target.slice(1) : dir + rel.target).split("/");
  const out: string[] = [];
  for (const x of parts) (x === ".." ? out.pop() : x !== "." && out.push(x));
  return out.join("/");
}

/** every part a relationship of `type` points to from `source`, in
 *  relationship order (a document can carry several headers/footers) */
export function relatedParts(entries: Parts, source: string, type: string): string[] {
  const slash = source.lastIndexOf("/");
  const dir = source.slice(0, slash + 1);
  const relsXml = entries.get(`${dir}_rels/${source.slice(slash + 1)}.rels`);
  if (!relsXml) return [];
  const out: string[] = [];
  for (const rel of parseRels(xmlText(`${dir}_rels/${source.slice(slash + 1)}.rels`, relsXml))) {
    if (r_targetModeExternal(rel) || !REL_BASES.some((ns) => rel.type === ns + type)) continue;
    const parts = (rel.target.startsWith("/") ? rel.target.slice(1) : dir + rel.target).split("/");
    const path: string[] = [];
    for (const x of parts) (x === ".." ? path.pop() : x !== "." && path.push(x));
    out.push(path.join("/"));
  }
  return out;
}

const r_targetModeExternal = (rel: Relationship): boolean => rel.targetMode === "External";

/** anything that can hand back part bytes by name (an Entries map or a
 *  PackageReader) */
export interface Parts {
  get(name: string): EntryData | undefined;
}

/** the main document part, named through the package's officeDocument
 *  relationship; the conventional name is the fallback */
export function mainPartName(entries: Parts): string {
  return relatedPart(entries, "", "officeDocument") ?? "word/document.xml";
}

export { parseXml, serializeXml };
export type { XElement, XDoc };
