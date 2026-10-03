// Light part registry: [Content_Types].xml and relationship parts. The gate
// never adds or removes parts, so this exists to answer "what is this part"
// questions and to keep door open for later phases (comments, headers).
import { decodeEntities, kid, kids, parseXml, serializeXml, type XElement, type XDoc } from "./xml.js";

export interface ContentTypes {
  defaults: Map<string, string>; // extension -> content type
  overrides: Map<string, string>; // part name -> content type
}

export function parseContentTypes(xml: string): ContentTypes {
  const doc = parseXml(xml);
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
  const doc = parseXml(xml);
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

export { parseXml, serializeXml };
export type { XElement, XDoc };
