// Input limits hold before anything inflates, ZIP64 archives included; an
// entry never inflates past its recorded size; entries compare in chunks; and
// the lazy parse keeps every source byte through views, clones and in-place
// parses.
import { deflateSync } from "fflate";
import { describe, expect, it } from "vitest";
import { entriesEqual, entryBytes, encodePart, openPackage, PACKAGE_LIMITS, PackageLimitError, RawEntry, writePackage, type Entries } from "./zip.js";
import { cloneNode, getAttr, isPending, kid, kids, materialize, MAX_XML_DEPTH, parseXml, readView, serializeXml, setAttr, textNode, type XElement } from "./xml.js";

const refused = (fn: () => unknown): string => {
  try {
    fn();
  } catch (e) {
    if (e instanceof PackageLimitError) return e.code;
    throw e;
  }
  throw new Error("expected a PackageLimitError");
};

/**
 * The same archive in ZIP64 form: every central header's sizes and offset
 * saturated and carried in a ZIP64 extra field, the entry count and directory
 * offset carried in a ZIP64 end record. `sizes` overrides recorded sizes.
 */
function toZip64(zip: Uint8Array, sizes: Record<string, number> = {}): Uint8Array {
  const v = new DataView(zip.buffer, zip.byteOffset, zip.byteLength);
  const eocd = zip.length - 22;
  const count = v.getUint16(eocd + 10, true);
  const cdStart = v.getUint32(eocd + 16, true);
  const out: number[] = [...zip.subarray(0, cdStart)];
  const u16 = (x: number) => [x & 0xff, (x >>> 8) & 0xff];
  const u32 = (x: number) => [x & 0xff, (x >>> 8) & 0xff, (x >>> 16) & 0xff, (x >>> 24) & 0xff];
  const u64 = (x: number) => [...u32(x % 2 ** 32), ...u32(Math.floor(x / 2 ** 32))];
  let p = cdStart;
  for (let n = 0; n < count; n++) {
    const nameLen = v.getUint16(p + 28, true);
    const name = new TextDecoder().decode(zip.subarray(p + 46, p + 46 + nameLen));
    const size = sizes[name] ?? v.getUint32(p + 24, true);
    const extra = [...u16(0x0001), ...u16(24), ...u64(size), ...u64(v.getUint32(p + 20, true)), ...u64(v.getUint32(p + 42, true))];
    const head = [...zip.subarray(p, p + 46)];
    head.splice(20, 4, ...u32(0xffffffff));
    head.splice(24, 4, ...u32(0xffffffff));
    head.splice(30, 2, ...u16(extra.length));
    head.splice(42, 4, ...u32(0xffffffff));
    out.push(...head, ...zip.subarray(p + 46, p + 46 + nameLen), ...extra);
    p += 46 + nameLen + v.getUint16(p + 30, true) + v.getUint16(p + 32, true);
  }
  const cdSize = out.length - cdStart;
  const eocd64 = out.length;
  out.push(...u32(0x06064b50), ...u64(44), ...u16(45), ...u16(45), ...u32(0), ...u32(0), ...u64(count), ...u64(count), ...u64(cdSize), ...u64(cdStart));
  out.push(...u32(0x07064b50), ...u32(0), ...u64(eocd64), ...u32(1));
  out.push(...u32(0x06054b50), ...u16(0), ...u16(0), ...u16(0xffff), ...u16(0xffff), ...u32(0xffffffff), ...u32(0xffffffff), ...u16(0));
  return new Uint8Array(out);
}

const sample = (): Entries =>
  new Map([
    ["[Content_Types].xml", encodePart('<?xml version="1.0"?><Types/>')],
    ["word/document.xml", encodePart("<w:document><w:body><w:p/></w:body></w:document>".repeat(50))],
  ]);

describe("archive limits", () => {
  it("a ZIP64 archive reads through its ZIP64 records, every entry byte-identical", () => {
    const plain = writePackage(sample());
    const reader = openPackage(toZip64(plain));
    expect(reader.names).toEqual([...sample().keys()]);
    for (const [name, bytes] of sample()) expect(entryBytes(reader.raw(name)!)).toEqual(bytes);
  });

  it("a ZIP64 archive whose entry records more than the limit is refused before anything inflates", () => {
    const zip = toZip64(writePackage(sample()), { "word/document.xml": PACKAGE_LIMITS.maxEntryUncompressed + 1 });
    expect(refused(() => openPackage(zip))).toBe("package_too_large");
  });

  it("ZIP64 records whose lengths overrun their enclosing record are unreadable", () => {
    const zip = toZip64(writePackage(sample()));
    const v = new DataView(zip.buffer);
    const eocd64 = Number(v.getBigUint64(zip.length - 22 - 20 + 8, true));
    const cd = Number(v.getBigUint64(eocd64 + 48, true));
    const nameLen = v.getUint16(cd + 28, true);
    const overlong = zip.slice();
    new DataView(overlong.buffer).setUint16(cd + 46 + nameLen + 2, 0xffff, true);
    expect(refused(() => openPackage(overlong))).toBe("unreadable_archive");
    const undersized = zip.slice();
    new DataView(undersized.buffer).setBigUint64(eocd64 + 4, 20n, true);
    expect(refused(() => openPackage(undersized))).toBe("unreadable_archive");
  });

  it("an entry that inflates past its recorded size is unreadable, not inflated in full", () => {
    const content = new Uint8Array(1_000_000).fill(65);
    const lying = new RawEntry(deflateSync(content), 8, 0, 1000);
    expect(refused(() => entryBytes(lying))).toBe("unreadable_archive");
    const short = new RawEntry(deflateSync(content), 8, 0, 2_000_000);
    expect(refused(() => entryBytes(short))).toBe("unreadable_archive");
  });

  it("an archive with no end record, or one that lies about its entry count, is unreadable", () => {
    expect(refused(() => openPackage(new Uint8Array(100)))).toBe("unreadable_archive");
    const zip = writePackage(sample());
    const v = new DataView(zip.buffer);
    v.setUint16(zip.length - 22 + 10, 5, true);
    expect(refused(() => openPackage(zip))).toBe("unreadable_archive");
  });

  it("an XML part over maxXmlPartBytes is refused as document_too_large before it inflates", () => {
    const reader = openPackage(writePackage(sample()));
    const limit = PACKAGE_LIMITS.maxXmlPartBytes;
    PACKAGE_LIMITS.maxXmlPartBytes = 100;
    try {
      expect(refused(() => reader.xml("word/document.xml"))).toBe("document_too_large");
      expect(reader.xml("[Content_Types].xml")).toBe('<?xml version="1.0"?><Types/>');
    } finally {
      PACKAGE_LIMITS.maxXmlPartBytes = limit;
    }
  });
});

describe("entriesEqual", () => {
  const big = new Uint8Array(3_000_000);
  for (let i = 0; i < big.length; i++) big[i] = (i * 7919) % 251;
  const raw = (bytes: Uint8Array, level: 1 | 9) => new RawEntry(deflateSync(bytes, { level }), 8, 0, bytes.length);

  it("the same content compressed differently is equal", () => {
    expect(entriesEqual(raw(big, 1), raw(big, 9))).toBe(true);
    expect(entriesEqual(raw(big, 1), big)).toBe(true);
  });

  it("one changed byte anywhere, or a different length, is not equal", () => {
    for (const at of [0, 1_500_000, big.length - 1]) {
      const changed = big.slice();
      changed[at] = changed[at]! ^ 1;
      expect(entriesEqual(raw(big, 1), raw(changed, 9))).toBe(false);
    }
    expect(entriesEqual(raw(big, 1), raw(big.subarray(1), 9))).toBe(false);
  });

  it("an entry that holds more than it records is unreadable", () => {
    const lying = new RawEntry(deflateSync(big), 8, 0, 1000);
    expect(refused(() => entriesEqual(lying, raw(big.subarray(0, 1000), 9)))).toBe("unreadable_archive");
  });
});

describe("xml depth", () => {
  it("nesting past MAX_XML_DEPTH is refused as document_too_large", () => {
    const deep = "<a>".repeat(MAX_XML_DEPTH + 1) + "</a>".repeat(MAX_XML_DEPTH + 1);
    expect(refused(() => parseXml(deep))).toBe("document_too_large");
    const doc = parseXml(`<r>${"<b>".repeat(MAX_XML_DEPTH)}${"</b>".repeat(MAX_XML_DEPTH)}</r>`, { lazy: (n) => n === "b" });
    expect(refused(() => materialize(kid(kid(doc as unknown as XElement)!)!))).toBe("document_too_large");
  });
});

describe("lazy parse fidelity", () => {
  // tag forms a rebuild would normalize: odd spacing, both quote styles,
  // entities, a spaced self-closing tag, an end tag with trailing space
  const SOURCE =
    `<?xml version="1.0"?>\n<w:body><w:p  w14:paraId = '0A' w:rsidR="00"><w:r><w:t xml:space="preserve"> a &amp; b </w:t ></w:r><w:br /><w:r><w:t>c</w:t></w:r></w:p>` +
    `<w:p w14:paraId="0B" w:rsidR="00"><w:r><w:t>d</w:t></w:r></w:p></w:body>`;
  const parse = () => {
    const doc = parseXml(SOURCE, { lazy: (n) => n === "w:p" });
    const body = kid(doc as unknown as XElement, "w:body")!;
    return { doc, paras: kids(body, "w:p") };
  };

  it("serializing unread, read in place, and partly read trees reproduces the source byte for byte", () => {
    const a = parse();
    expect(serializeXml(a.doc)).toBe(SOURCE);
    materialize(a.paras[0]!);
    expect(isPending(a.paras[0]!)).toBe(false);
    expect(isPending(a.paras[1]!)).toBe(true);
    expect(serializeXml(a.doc)).toBe(SOURCE);
    materialize(a.paras[1]!);
    expect(serializeXml(a.doc)).toBe(SOURCE);
  });

  it("a read view leaves the element unread and the tree unchanged, and its parse matches an in-place one", () => {
    const { doc, paras } = parse();
    const view = readView(paras[0]!);
    expect(isPending(paras[0]!)).toBe(true);
    expect(serializeXml(view)).toBe(serializeXml(paras[0]!));
    view.children.push(textNode("x"));
    expect(serializeXml(doc)).toBe(SOURCE);
    expect(view.parent).toBe(paras[0]!.parent);
  });

  it("a clone of an unread element is independent: editing it leaves the original and the source alone", () => {
    const { doc, paras } = parse();
    const copy = cloneNode(paras[0]!);
    expect(isPending(paras[0]!)).toBe(true);
    expect(serializeXml(copy)).toBe(serializeXml(paras[0]!));
    setAttr(copy, "w14:paraId", "0C");
    setAttr(kid(kid(copy, "w:r")!, "w:t")!, "xml:space", "default");
    kid(copy, "w:r")!.children.length = 0;
    expect(getAttr(paras[0]!, "w14:paraId")).toBe("0A");
    expect(serializeXml(doc)).toBe(SOURCE);
    expect(copy.parent).toBe(null);
  });

  it("elements parsed from identical tags share no attribute or child arrays", () => {
    const { doc, paras } = parse();
    const [p0, p1] = paras;
    const r0 = kids(p0!, "w:r");
    expect(r0[0]!.attrs).not.toBe(r0[1]!.attrs);
    expect(r0[0]!.children).not.toBe(r0[1]!.children);
    setAttr(p1!, "w:rsidR", "FF");
    expect(p0!.attrs.find((a) => a.name === "w:rsidR")!.raw).toBe("00");
    setAttr(r0[0]!, "w:rsidR", "01");
    expect(r0[1]!.attrs).toEqual([]);
    kid(r0[1]!, "w:t")!.children.push(textNode("!"));
    expect(serializeXml(r0[0]!)).toBe('<w:r w:rsidR="01"><w:t xml:space="preserve"> a &amp; b </w:t ></w:r>');
    expect(serializeXml(doc)).toContain("<w:t>c!</w:t>");
  });
});
