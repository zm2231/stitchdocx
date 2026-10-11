// Byte-faithful XML tree for OOXML parts.
//
// Design invariants:
// 1. serialize(parse(s)) === s for any well-formed XML part. Every node keeps
//    its raw source form (entities, attribute order, quoting, self-closing
//    form, whitespace inside tags), so untouched subtrees re-serialize
//    byte-identically by construction.
// 2. Namespace prefixes are matched literally as written ("w:p"). OOXML
//    writers (Word, LibreOffice, Google Docs, this repo's fixtures) use the
//    conventional prefixes; Word does the same.
// 3. Text nodes store the raw source (entities intact). Decoding happens only
//    through decodedText()/getAttr(); newly created content is escaped on the
//    way in and stored in escaped form.

import { PackageLimitError } from "./zip.js";

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

/** maximum element nesting the parser accepts */
export const MAX_XML_DEPTH = 512;

// —————————————————————————————————————————————————————————————— parsing

export interface ParseOptions {
  /** elements whose name this accepts keep their content as unparsed source
   *  until it is read (see LazyElement); serializing an unread one copies its
   *  source */
  lazy?: (name: string) => boolean;
  /** nesting depth of the element the source sits in, for MAX_XML_DEPTH */
  depth?: number;
  /** nodes the parse may build (a lazy element counts as one), shared with
   *  every lazy element it makes: reading one's content in place draws on the
   *  same budget, and a transient readView gets a fresh one of the same limit */
  budget?: NodeBudget;
}

/** a bound on the nodes parses may build; past it the parse stops with a
 *  document_too_large PackageLimitError */
export class NodeBudget {
  #used = 0;
  constructor(readonly limit: number) {}

  take(): void {
    if (++this.#used > this.limit) throw new PackageLimitError("document_too_large", `xml parse: more than ${this.limit} nodes`);
  }
}

export function parseXml(source: string, opts: ParseOptions = {}): XDoc {
  const children: XNode[] = [];
  const base = opts.depth ?? 0;
  const budget = opts.budget;
  // one copy of each name per parse: element and attribute names, and tags
  // written as the bare name, repeat on nearly every node
  const names = new Map<string, string>();
  const intern = (x: string): string => {
    const known = names.get(x);
    if (known !== undefined) return known;
    names.set(x, x);
    return x;
  };
  // element stack only tracks elements for parent links; nesting validity is
  // enforced by matching end tags to the innermost open element
  const stack: XElement[] = [];
  const push = (n: XNode): void => {
    budget?.take();
    if (n.kind === "el") {
      // bound the tree depth: hostile nesting would otherwise blow the stack
      // in every recursive walk (serialize, clone, resolve)
      if (base + stack.length >= MAX_XML_DEPTH) {
        i = s.length;
        throw new PackageLimitError("document_too_large", `xml parse: element nesting deeper than ${MAX_XML_DEPTH}`);
      }
      n.parent = stack.length ? stack[stack.length - 1]! : null;
    }
    if (stack.length) stack[stack.length - 1]!.children.push(n);
    else children.push(n);
  };
  let i = 0;
  const s = source;

  function fail(msg: string): never {
    const line = s.slice(0, i).split("\n").length;
    const col = i - (s.lastIndexOf("\n", i - 1) + 1) + 1;
    throw new Error(`xml parse at ${line}:${col}: ${msg}`);
  }

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
    const markup = skipMarkup(s, i, fail);
    if (markup > 0) {
      push({ kind: "raw", raw: s.slice(i, markup) });
      i = markup;
    } else if (s.startsWith("</", i)) {
      const end = s.indexOf(">", i + 2);
      if (end < 0) fail("unterminated end tag");
      const name = s.slice(i + 2, end).trim();
      const open = stack.pop();
      if (!open) fail(`end tag </${name}> with no open element`);
      if (open.name !== name) fail(`end tag </${name}> does not match open <${open.name}>`);
      open.rawClose = end - (i + 2) === open.name.length ? open.name : s.slice(i + 2, end);
      // a child list grown by push keeps spare capacity; keep only what is used
      open.children = open.children.slice();
      i = end + 1;
    } else {
      const tag = readStartTag(s, i, fail);
      const name = intern(tag.name);
      const head: XElement = {
        kind: "el",
        name,
        attrs: parseAttrs(tag.attrs, fail, intern),
        children: [],
        parent: null,
        rawOpen: tag.open === name ? name : tag.open,
        rawClose: null,
        selfClosing: tag.selfClosing,
        attrsDirty: false,
      };
      i = tag.end;
      if (!tag.selfClosing && opts.lazy?.(tag.name)) {
        const close = findClose(s, i, tag.name, fail);
        const lazy = new LazyElement(head, s, i, close.innerEnd, base + stack.length + 1, budget);
        lazy.rawClose = close.rawClose === name ? name : close.rawClose;
        push(lazy);
        i = close.end;
        continue;
      }
      push(head);
      if (!tag.selfClosing) stack.push(head);
    }
  }
  if (stack.length) fail(`unclosed element <${stack[stack.length - 1]!.name}>`);
  return { kind: "doc", children: children.slice() };
}

type Fail = (msg: string) => never;

/** end offset of a comment, CDATA section, doctype or processing instruction
 *  starting at `i`, or 0 when `i` starts a tag */
function skipMarkup(s: string, i: number, fail: Fail): number {
  if (s.startsWith("<!--", i)) {
    const end = s.indexOf("-->", i + 4);
    if (end < 0) fail("unterminated comment");
    return end + 3;
  }
  if (s.startsWith("<![CDATA[", i)) {
    const end = s.indexOf("]]>", i + 9);
    if (end < 0) fail("unterminated CDATA");
    return end + 3;
  }
  if (s.startsWith("<!", i)) {
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
    return j + 1;
  }
  if (s.startsWith("<?", i)) {
    const end = s.indexOf("?>", i + 2);
    if (end < 0) fail("unterminated processing instruction");
    return end + 2;
  }
  return 0;
}

const NAME_END = /[\s/>]/;

/** the start tag at `i` (s[i] === "<"): its name, the attribute source, the
 *  open-tag source as kept on XElement.rawOpen, and the offset after ">" */
function readStartTag(s: string, i: number, fail: Fail): { name: string; attrs: string; open: string; selfClosing: boolean; end: number } {
  let j = i + 1;
  const nameStart = j;
  while (j < s.length && !NAME_END.test(s[j]!)) j++;
  if (j === nameStart) fail("element name expected");
  const name = s.slice(nameStart, j);
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
  const tagBody = s.slice(j, k); // between name and ">", quotes included
  const selfClosing = tagBody.trimEnd().endsWith("/");
  let open = s.slice(i + 1, k); // exact "<...>" interior
  if (selfClosing) open = open.slice(0, open.lastIndexOf("/")); // "/" is re-emitted; the space in "<x />" is kept
  const attrs = selfClosing ? tagBody.slice(0, tagBody.trimEnd().length - 1) : tagBody;
  return { name, attrs, open, selfClosing, end: k + 1 };
}

/** the end tag matching the element of `name` whose content starts at `from` */
function findClose(s: string, from: number, name: string, fail: Fail): { innerEnd: number; rawClose: string; end: number } {
  let depth = 1;
  let j = from;
  while (true) {
    const lt = s.indexOf("<", j);
    if (lt < 0) fail(`unclosed element <${name}>`);
    const markup = skipMarkup(s, lt, fail);
    if (markup > 0) {
      j = markup;
    } else if (s.startsWith("</", lt)) {
      const end = s.indexOf(">", lt + 2);
      if (end < 0) fail("unterminated end tag");
      if (s.slice(lt + 2, end).trim() === name && --depth === 0) return { innerEnd: lt, rawClose: s.slice(lt + 2, end), end: end + 1 };
      j = end + 1;
    } else {
      const tag = readStartTag(s, lt, fail);
      if (tag.name === name && !tag.selfClosing) depth++;
      j = tag.end;
    }
  }
}

function parseAttrs(body: string, fail: Fail, intern: (name: string) => string = (x) => x): XAttr[] {
  const attrs: XAttr[] = [];
  let k = 0;
  while (k < body.length) {
    while (k < body.length && /\s/.test(body[k]!)) k++;
    if (k >= body.length) break;
    const nameStart = k;
    while (k < body.length && !/[\s=]/.test(body[k]!)) k++;
    const name = intern(body.slice(nameStart, k));
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
  return attrs.slice();
}

/**
 * Every start tag in `source` whose name `want` accepts, with its attributes,
 * in document order. Builds no tree: this is how ids are collected across a
 * whole part without holding it parsed.
 */
export function scanStartTags(source: string, want: (name: string) => boolean, visit: (name: string, attrs: XAttr[]) => void): void {
  const fail: Fail = (msg) => {
    throw new Error(`xml scan: ${msg}`);
  };
  let j = 0;
  while (true) {
    const lt = source.indexOf("<", j);
    if (lt < 0) return;
    const markup = skipMarkup(source, lt, fail);
    if (markup > 0) {
      j = markup;
    } else if (source.startsWith("</", lt)) {
      j = lt + 2;
    } else {
      const tag = readStartTag(source, lt, fail);
      if (want(tag.name)) visit(tag.name, parseAttrs(tag.attrs, fail));
      j = tag.end;
    }
  }
}

// ———————————————————————————————————————————————————————— lazy elements

/**
 * An element whose content stays source text until something reads
 * `children`; the first read parses it in place and it behaves like any other
 * element from then on. Until then it serializes as its source, byte for
 * byte. Read-only passes that must not keep a parse use readView().
 */
class LazyElement implements XElement {
  readonly kind = "el" as const;
  name: string;
  attrs: XAttr[];
  parent: XElement | null = null;
  rawOpen: string;
  rawClose: string | null = null;
  selfClosing = false;
  attrsDirty = false;
  #source: string | null;
  #from: number;
  #to: number;
  #depth: number;
  #children: XNode[] | null = null;
  #budget: NodeBudget | undefined;

  constructor(head: XElement, source: string, from: number, to: number, depth: number, budget: NodeBudget | undefined) {
    this.name = head.name;
    this.attrs = head.attrs;
    this.rawOpen = head.rawOpen;
    this.#source = source;
    this.#from = from;
    this.#to = to;
    this.#depth = depth;
    this.#budget = budget;
  }

  get children(): XNode[] {
    if (this.#children === null) {
      this.#children = this.parseContent(this, this.#budget);
      this.#source = null;
    }
    return this.#children;
  }

  set children(value: XNode[]) {
    this.#children = value;
    this.#source = null;
  }

  get budgetLimit(): number | undefined {
    return this.#budget?.limit;
  }

  /** the unparsed content source, or null once parsed */
  get pending(): string | null {
    return this.#children === null ? this.#source!.slice(this.#from, this.#to) : null;
  }

  /** a fresh parse of the content, its elements parented to `owner` */
  parseContent(owner: XElement, budget: NodeBudget | undefined): XNode[] {
    if (this.#children !== null) throw new Error("parseContent on a parsed element");
    const nodes = parseXml(this.#source!.slice(this.#from, this.#to), { depth: this.#depth, budget }).children;
    for (const n of nodes) if (n.kind === "el") n.parent = owner;
    return nodes;
  }
}

/** true while a lazy element's content is still unparsed source */
export function isPending(el: XElement): boolean {
  return el instanceof LazyElement && el.pending !== null;
}

/** a lazy element's unparsed content source, or null */
export function pendingSource(el: XElement): string | null {
  return el instanceof LazyElement ? el.pending : null;
}

/** parse a lazy element's content in place (a no-op once parsed) */
export function materialize(el: XElement): void {
  void el.children;
}

/** every lazy element under `root` whose content is still unparsed, in
 *  document order, without parsing any */
export function* pendingElements(root: XElement): Generator<XElement> {
  if (isPending(root)) {
    yield root;
    return;
  }
  for (const c of root.children) if (c.kind === "el") yield* pendingElements(c);
}

/**
 * The element for reading: itself once parsed, or a detached parse of its
 * content that is dropped after use, so a read-only pass over every lazy
 * element never holds more than one parsed at a time. The view shares the
 * element's attributes and parent; it must not be modified.
 */
export function readView(el: XElement): XElement {
  if (!(el instanceof LazyElement) || el.pending === null) return el;
  const view: XElement = {
    kind: "el",
    name: el.name,
    attrs: el.attrs,
    children: [],
    parent: el.parent,
    rawOpen: el.rawOpen,
    rawClose: el.rawClose,
    selfClosing: false,
    attrsDirty: el.attrsDirty,
  };
  view.children = el.parseContent(view, el.budgetLimit === undefined ? undefined : new NodeBudget(el.budgetLimit));
  return view;
}

// ———————————————————————————————————————————————————————————— serializing

export interface SerializeOptions {
  /** elements left out of the output, with everything inside them */
  omit?: (el: XElement) => boolean;
}

export function serializeXml(node: XDoc | XNode | XElement, opts: SerializeOptions = {}): string {
  if (node.kind === "doc") return node.children.map((c) => serializeXml(c, opts)).join("");
  if (node.kind === "text") return node.raw;
  if (node.kind === "raw") return node.raw;
  if (opts.omit?.(node)) return "";
  const open = node.attrsDirty ? rebuildOpenTag(node) : node.rawOpen;
  const pending = pendingSource(node);
  if (pending !== null) return `<${open}>${pending}</${node.rawClose ?? node.name}>`;
  if (node.selfClosing && node.children.length === 0) return `<${open}/>`;
  const inner = node.children.map((c) => serializeXml(c, opts)).join("");
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
  return escapeMarkup(normalizeNewlines(stripInvalidXmlChars(text)));
}

/** alias used by consumers that escape run text */
export const encodeTextSafe = encodeText;

/**
 * Strips characters XML 1.0 cannot represent (control chars, lone surrogates,
 * non-characters). Applied only on authored content — parsed raw text is
 * preserved verbatim.
 */
export function stripInvalidXmlChars(text: string): string {
  return text
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, "")
    .replace(/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g, "");
}

/** normalize CR/CRLF to LF (what a spec parser does to literal text newlines) */
export function normalizeNewlines(text: string): string {
  return text.includes("\r") ? text.replace(/\r\n?/g, "\n") : text;
}

function escapeMarkup(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/**
 * Decoded text as Word reads it: literal CR/CRLF in the source is normalized
 * to LF at parse time, while character references (&#13;) survive decoding as
 * CR. Input is raw source text.
 */
export function decodeText(raw: string): string {
  const normalized = raw.includes("\r") ? raw.replace(/\r\n?/g, "\n") : raw;
  return decodeEntities(normalized);
}

/** escaped form for an attribute value; caller supplies the quote in use.
 * Newlines/tabs/CR are written as character references: XML attribute
 * normalization would otherwise turn literal whitespace into spaces and the
 * value would not round-trip. */
export function encodeAttrValue(text: string, q: '"' | "'"): string {
  let out = escapeMarkup(stripInvalidXmlChars(text))
    .replace(/\n/g, "&#10;")
    .replace(/\t/g, "&#9;")
    .replace(/\r/g, "&#13;");
  if (q === '"') out = out.replace(/"/g, "&quot;");
  else out = out.replace(/'/g, "&apos;");
  return out;
}

export function getAttr(el: XElement, name: string): string | null {
  const a = el.attrs.find((a) => a.name === name);
  return a ? decodeAttrValue(a.raw) : null;
}

/**
 * Attribute-value normalization parity with a spec parser: literal CR/LF
 * collapse to a single space, then remaining literal tab/newline each become
 * one space (character references are exempt — decodeEntities runs after).
 */
export function decodeAttrValue(raw: string): string {
  const normalized = /[\t\n\r]/.test(raw) ? raw.replace(/\r\n/g, " ").replace(/[\t\n\r]/g, " ") : raw;
  return decodeEntities(normalized);
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
  for (const c of el.children) if (c.kind === "el" && !children.includes(c)) c.parent = null;
  for (const c of children) if (c.kind === "el" && c.parent && c.parent !== el) detachFromParent(c);
  el.children = children;
  el.selfClosing = false;
  reassignParents(el);
}

/** remove a node from its current parent's children (if attached) */
function detachFromParent(child: XNode): void {
  if (child.kind !== "el" || !child.parent) return;
  const i = child.parent.children.indexOf(child);
  if (i >= 0) child.parent.children.splice(i, 1);
}

export function appendChild(el: XElement, child: XNode): void {
  detachFromParent(child);
  el.children.push(child);
  if (el.children.length === 1 && el.selfClosing) el.selfClosing = false;
  reassignParents(el);
}

export function insertAt(el: XElement, index: number, child: XNode): void {
  detachFromParent(child);
  el.children.splice(index, 0, child);
  if (el.selfClosing) el.selfClosing = false;
  reassignParents(el);
}

export function removeChild(child: XNode): void {
  const p = child.kind === "el" ? child.parent : null;
  if (!p) return;
  const i = p.children.indexOf(child);
  if (i >= 0) p.children.splice(i, 1);
  if (child.kind === "el") child.parent = null; // detached: parent pointer must not dangle
}

/** fix parent pointers of an element's direct children after structural edits */
export function reassignParents(el: XElement): void {
  for (const c of el.children) if (c.kind === "el") c.parent = el;
}

/** deep clone; clones are detached (parent null) and keep raw forms */
export function cloneNode<T extends XNode>(node: T): T {
  if (node.kind === "text" || node.kind === "raw") return { ...node };
  const source = readView(node);
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
  copy.children = source.children.map((c) => {
    const cc = cloneNode(c);
    if (cc.kind === "el") cc.parent = copy;
    return cc;
  });
  return copy as T;
}
