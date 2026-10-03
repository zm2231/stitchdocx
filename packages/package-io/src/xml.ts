// Byte-faithful XML tree for OOXML parts.
//
// Design invariants (D1 decision, see research/):
// 1. serialize(parse(s)) === s for any well-formed XML part. Every node keeps
//    its raw source form (entities, attribute order, quoting, self-closing
//    form, whitespace inside tags), so untouched subtrees re-serialize
//    byte-identically by construction.
// 2. Namespace prefixes are matched literally as written ("w:p"). OOXML
//    writers (Word, LibreOffice, Google Docs, this repo's fixtures) use the
//    conventional prefixes; the harness reader does the same.
// 3. Text nodes store the raw source (entities intact). Decoding happens only
//    through decodedText()/getAttr(); newly created content is escaped on the
//    way in and stored in escaped form.

export interface XAttr {
  /** attribute name as written, e.g. "w14:paraId" */
  name: string;
  /** value as written between the quotes (entities intact, not decoded) */
  raw: string;
  /** quote character used in the source */
  q: '"' | "'";
}

export interface XElement {
  kind: "el";
  name: string;
  attrs: XAttr[];
  children: XNode[];
  parent: XElement | null;
  /** open-tag source between "<" and ">" (attributes as written); stale when attrsDirty */
  rawOpen: string;
  /** end-tag source between "</" and ">" as written; null until the matching end tag is seen */
  rawClose: string | null;
  /** the source used a self-closing tag; meaningful only while children is empty */
  selfClosing: boolean;
  attrsDirty: boolean;
}

export interface XText {
  kind: "text";
  /** character data exactly as in the source (entities intact) */
  raw: string;
}

/** comments, CDATA, processing instructions, doctype, BOM — emitted verbatim */
export interface XRaw {
  kind: "raw";
  raw: string;
}

export type XNode = XElement | XText | XRaw;

export interface XDoc {
  kind: "doc";
  children: XNode[];
}

export const isEl = (n: XNode): n is XElement => n.kind === "el";

// —————————————————————————————————————————————————————————————— parsing

export function parseXml(source: string): XDoc {
  const children: XNode[] = [];
  // element stack only tracks elements for parent links; nesting validity is
  // enforced by matching end tags to the innermost open element
  const stack: XElement[] = [];
  const push = (n: XNode): void => {
    if (stack.length) stack[stack.length - 1]!.children.push(n);
    else children.push(n);
  };
  let i = 0;
  const s = source;

  const fail = (msg: string): never => {
    const line = s.slice(0, i).split("\n").length;
    const col = i - (s.lastIndexOf("\n", i - 1) + 1) + 1;
    throw new Error(`xml parse at ${line}:${col}: ${msg}`);
  };

  if (s.startsWith("\uFEFF")) {
    push({ kind: "raw", raw: "\uFEFF" });
    i = 1;
  }

  while (i < s.length) {
    const lt = s.indexOf("<", i);
    if (lt < 0) {
      push({ kind: "text", raw: s.slice(i) });
      break;
    }
    if (lt > i) push({ kind: "text", raw: s.slice(i, lt) });
    i = lt;
    if (s.startsWith("<!--", i)) {
      const end = s.indexOf("-->", i + 4);
      if (end < 0) fail("unterminated comment");
      push({ kind: "raw", raw: s.slice(i, end + 3) });
      i = end + 3;
    } else if (s.startsWith("<![CDATA[", i)) {
      const end = s.indexOf("]]>", i + 9);
      if (end < 0) fail("unterminated CDATA");
      push({ kind: "raw", raw: s.slice(i, end + 3) });
      i = end + 3;
    } else if (s.startsWith("<!", i)) {
      // doctype, possibly with an internal subset in [ ... ]
      let j = i + 2;
      let depth = 0;
      while (j < s.length) {
        const c = s[j];
        if (c === "[") depth++;
        else if (c === "]") depth--;
        else if (c === ">" && depth <= 0) break;
        j++;
      }
      if (j >= s.length) fail("unterminated doctype");
      push({ kind: "raw", raw: s.slice(i, j + 1) });
      i = j + 1;
    } else if (s.startsWith("<?", i)) {
      const end = s.indexOf("?>", i + 2);
      if (end < 0) fail("unterminated processing instruction");
      push({ kind: "raw", raw: s.slice(i, end + 2) });
      i = end + 2;
    } else if (s.startsWith("</", i)) {
      const end = s.indexOf(">", i + 2);
      if (end < 0) fail("unterminated end tag");
      const name = s.slice(i + 2, end).trim();
      const open = stack.pop();
      if (!open) fail(`end tag </${name}> with no open element`);
      if (open.name !== name) fail(`end tag </${name}> does not match open <${open.name}>`);
      open.rawClose = s.slice(i + 2, end);
      i = end + 1;
    } else {
      // start tag
      const el = parseStartTag();
      push(el);
      if (!el.selfClosing) stack.push(el);
    }
  }
  if (stack.length) fail(`unclosed element <${stack[stack.length - 1]!.name}>`);
  return { kind: "doc", children };

  function parseStartTag(): XElement {
    // s[i] === "<"
    let j = i + 1;
    const nameStart = j;
    while (j < s.length && !/[\s/>]/.test(s[j]!)) j++;
    if (j === nameStart) fail("element name expected");
    const name = s.slice(nameStart, j);
    const attrsStart = j;
    let selfClosing = false;
    // scan to the end of the tag, respecting quoted attribute values
    let k = j;
    while (k < s.length) {
      const c = s[k]!;
      if (c === '"' || c === "'") {
        const end = s.indexOf(c, k + 1);
        if (end < 0) fail(`unterminated attribute value in <${name}>`);
        k = end + 1;
      } else if (c === ">") break;
      else k++;
    }
    if (k >= s.length) fail(`unterminated start tag <${name}>`);
    const tagBody = s.slice(attrsStart, k); // between name and ">", quotes included
    if (tagBody.trimEnd().endsWith("/")) selfClosing = true;
    let openStr = s.slice(i + 1, k); // exact "<...>" interior
    if (selfClosing) openStr = openStr.slice(0, openStr.lastIndexOf("/")); // "/" is re-emitted; the space in "<x />" is kept
    const el: XElement = {
      kind: "el",
      name,
      attrs: selfClosing ? parseAttrs(tagBody.slice(0, tagBody.trimEnd().length - 1)) : parseAttrs(tagBody),
      children: [],
      parent: null,
      rawOpen: openStr,
      rawClose: null,
      selfClosing,
      attrsDirty: false,
    };
    i = k + 1;
    return el;
  }

  function parseAttrs(body: string): XAttr[] {
    const attrs: XAttr[] = [];
    let k = 0;
    while (k < body.length) {
      while (k < body.length && /\s/.test(body[k]!)) k++;
      if (k >= body.length) break;
      const nameStart = k;
      while (k < body.length && !/[\s=]/.test(body[k]!)) k++;
      const name = body.slice(nameStart, k);
      if (!name) fail("attribute name expected");
      while (k < body.length && /\s/.test(body[k]!)) k++;
      if (body[k] !== "=") {
        // attribute without value (not valid OOXML, but keep it lossless)
        attrs.push({ name, raw: "", q: '"' });
        continue;
      }
      k++; // "="
      while (k < body.length && /\s/.test(body[k]!)) k++;
      const q = body[k] as '"' | "'" | undefined;
      if (q !== '"' && q !== "'") fail(`attribute ${name} value must be quoted`);
      const vStart = k + 1;
      const vEnd = body.indexOf(q, vStart);
      if (vEnd < 0) fail(`attribute ${name} value unterminated`);
      attrs.push({ name, raw: body.slice(vStart, vEnd), q });
      k = vEnd + 1;
    }
    return attrs;
  }
}

// ———————————————————————————————————————————————————————————— serializing

export function serializeXml(node: XDoc | XNode | XElement): string {
  if (node.kind === "doc") return node.children.map(serializeXml).join("");
  if (node.kind === "text") return node.raw;
  if (node.kind === "raw") return node.raw;
  const open = node.attrsDirty ? rebuildOpenTag(node) : node.rawOpen;
  if (node.selfClosing && node.children.length === 0) return `<${open}/>`;
  const inner = node.children.map(serializeXml).join("");
  return `<${open}>${inner}</${node.rawClose ?? node.name}>`;
}

function rebuildOpenTag(el: XElement): string {
  const attrs = el.attrs.map((a) => ` ${a.name}=${a.q}${a.raw}${a.q}`).join("");
  return `${el.name}${attrs}`;
}

// ———————————————————————————————————————————————————————————— navigation

export function kids(el: XElement | null | undefined, name?: string): XElement[] {
  if (!el) return [];
  return el.children.filter((c): c is XElement => c.kind === "el" && (name === undefined || c.name === name));
}

export function kid(el: XElement | null | undefined, name?: string): XElement | undefined {
  return kids(el, name)[0];
}

export function* walkEls(el: XElement): Generator<XElement> {
  yield el;
  for (const c of el.children) if (c.kind === "el") yield* walkEls(c);
}

/** index of a child within its parent, -1 if orphaned/root-level */
export function childIndex(el: XElement): number {
  if (!el.parent) return -1;
  return el.parent.children.indexOf(el);
}

// —————————————————————————————————————————————————————————— text & attrs

const NAMED: Record<string, string> = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" };

export function decodeEntities(raw: string): string {
  if (!raw.includes("&")) return raw;
  let out = "";
  for (let i = 0; i < raw.length; ) {
    const c = raw[i]!;
    if (c !== "&") {
      out += c;
      i++;
      continue;
    }
    const semi = raw.indexOf(";", i + 1);
    if (semi < 0) {
      out += c;
      i++;
      continue;
    }
    const body = raw.slice(i + 1, semi);
    if (body[0] === "#") {
      const code = body[1] === "x" || body[1] === "X" ? parseInt(body.slice(2), 16) : parseInt(body.slice(1), 10);
      if (Number.isNaN(code)) {
        out += c;
        i++;
      } else {
        out += String.fromCodePoint(code);
        i = semi + 1;
      }
    } else if (body in NAMED) {
      out += NAMED[body]!;
      i = semi + 1;
    } else {
      out += c;
      i++;
    }
  }
  return out;
}

/** escaped form stored inside a text node (quote-free context) */
export function encodeText(text: string): string {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/** escaped form for an attribute value; caller supplies the quote in use */
export function encodeAttrValue(text: string, q: '"' | "'"): string {
  let out = text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  if (q === '"') out = out.replace(/"/g, "&quot;");
  else out = out.replace(/'/g, "&apos;");
  return out;
}

export function getAttr(el: XElement, name: string): string | null {
  const a = el.attrs.find((a) => a.name === name);
  return a ? decodeEntities(a.raw) : null;
}

export function setAttr(el: XElement, name: string, value: string): void {
  const escaped = encodeAttrValue(value, '"');
  const existing = el.attrs.find((a) => a.name === name);
  if (existing) {
    existing.raw = escaped;
    existing.q = '"';
  } else {
    el.attrs.push({ name, raw: escaped, q: '"' });
  }
  el.attrsDirty = true;
}

export function removeAttr(el: XElement, name: string): void {
  const i = el.attrs.findIndex((a) => a.name === name);
  if (i >= 0) {
    el.attrs.splice(i, 1);
    el.attrsDirty = true;
  }
}

/** decoded character data of an element's direct text children */
export function textOf(el: XElement): string {
  let out = "";
  for (const c of el.children) if (c.kind === "text") out += decodeEntities(c.raw);
  return out;
}

// —————————————————————————————————————————————————————————— construction

/** create an element in canonical form (used for engine-authored nodes) */
export function el(name: string, attrs?: Record<string, string>, children?: XNode[]): XElement {
  const node: XElement = {
    kind: "el",
    name,
    attrs: [],
    children: children ?? [],
    parent: null,
    rawOpen: name,
    rawClose: name,
    selfClosing: false,
    attrsDirty: true,
  };
  if (attrs) for (const [k, v] of Object.entries(attrs)) setAttr(node, k, v);
  return node;
}

export function textNode(text: string): XText {
  return { kind: "text", raw: encodeText(text) };
}

/** element carrying raw inner source (used for verbatim grafts) */
export function rawNode(raw: string): XRaw {
  return { kind: "raw", raw };
}

export function setChildren(el: XElement, children: XNode[]): void {
  el.children = children;
  el.selfClosing = false;
  el.attrsDirty = el.attrsDirty; // rawOpen unchanged (attributes untouched)
  reassignParents(el);
}

export function appendChild(el: XElement, child: XNode): void {
  el.children.push(child);
  if (el.children.length === 1 && el.selfClosing) el.selfClosing = false;
  reassignParents(el);
}

export function insertAt(el: XElement, index: number, child: XNode): void {
  el.children.splice(index, 0, child);
  if (el.selfClosing) el.selfClosing = false;
  reassignParents(el);
}

export function removeChild(child: XNode): void {
  const p = child.kind === "el" ? child.parent : null;
  if (!p) return;
  const i = p.children.indexOf(child);
  if (i >= 0) p.children.splice(i, 1);
}

/** fix parent pointers of an element's direct children after structural edits */
export function reassignParents(el: XElement): void {
  for (const c of el.children) if (c.kind === "el") c.parent = el;
}

/** deep clone; clones are detached (parent null) and keep raw forms */
export function cloneNode<T extends XNode>(node: T): T {
  if (node.kind === "text" || node.kind === "raw") return { ...node };
  const copy: XElement = {
    kind: "el",
    name: node.name,
    attrs: node.attrs.map((a) => ({ ...a })),
    children: [],
    parent: null,
    rawOpen: node.rawOpen,
    rawClose: node.rawClose,
    selfClosing: node.selfClosing,
    attrsDirty: node.attrsDirty,
  };
  copy.children = node.children.map((c) => {
    const cc = cloneNode(c);
    if (cc.kind === "el") cc.parent = copy;
    return cc;
  });
  return copy as T;
}
