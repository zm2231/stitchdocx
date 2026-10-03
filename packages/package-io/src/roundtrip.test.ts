// M1: lossless package round-trip — re-saving an untouched package keeps
// every entry's content byte-identical, entry ORDER survives, and writing the
// same entries twice produces identical bytes (determinism).
import { readFileSync, readdirSync, statSync, existsSync } from "node:fs";
import { join } from "node:path";
import { homedir } from "node:os";
import { describe, it, expect } from "vitest";
import { zipSync } from "fflate";
import { decodePart, decodePart as dec, encodePart, readPackage, writePackage } from "./zip.js";
import { parseXml, serializeXml } from "./xml.js";

const HANDOFF = process.env.STITCH_HANDOFF ?? join(homedir(), "Downloads", "scope-docx-engine-handoff");

/** minimal stored (uncompressed) zip writer with explicit central-directory order */
function storedZip(entries: Array<[string, Uint8Array]>): Uint8Array {
  const chunks: Uint8Array[] = [];
  const central: Uint8Array[] = [];
  let offset = 0;
  const u16 = (v: number) => new Uint8Array([v & 0xff, (v >> 8) & 0xff]);
  const u32 = (v: number) => new Uint8Array([v & 0xff, (v >> 8) & 0xff, (v >> 16) & 0xff, (v >>> 24) & 0xff]);
  for (const [name, content] of entries) {
    const nameBytes = new TextEncoder().encode(name);
    const crc = crc32(content);
    const local = [...u32(0x04034b50), ...u16(20), ...u16(0), ...u16(0), ...u16(0), ...u16(0x21), ...u32(crc), ...u32(content.length), ...u32(content.length), ...u16(nameBytes.length), ...u16(0), ...nameBytes];
    chunks.push(new Uint8Array(local), content);
    const cd = [...u32(0x02014b50), ...u16(20), ...u16(20), ...u16(0), ...u16(0), ...u16(0), ...u16(0x21), ...u32(crc), ...u32(content.length), ...u32(content.length), ...u16(nameBytes.length), ...u16(0), ...u16(0), ...u16(0), ...u16(0), ...u32(0), ...u32(offset), ...nameBytes];
    central.push(new Uint8Array(cd));
    offset += local.length + content.length;
  }
  const cdStart = offset;
  let cdSize = 0;
  for (const c of central) cdSize += c.length;
  const eocd = new Uint8Array([...u32(0x06054b50), ...u16(0), ...u16(0), ...u16(entries.length), ...u16(entries.length), ...u32(cdSize), ...u32(cdStart), ...u16(0)]);
  const total = new Uint8Array(offset + cdSize + eocd.length);
  let at = 0;
  for (const c of [...chunks, ...central, eocd]) {
    total.set(c, at);
    at += c.length;
  }
  return total;
}

function crc32(data: Uint8Array): number {
  let c = ~0;
  for (let i = 0; i < data.length; i++) {
    c ^= data[i]!;
    for (let k = 0; k < 8; k++) c = (c >>> 1) ^ (0xedb88320 & -(c & 1));
  }
  return ~c >>> 0;
}

function* docxFiles(dir: string): Generator<string> {
  if (!existsSync(dir)) return;
  for (const name of readdirSync(dir)) {
    const p = join(dir, name);
    if (statSync(p).isDirectory()) yield* docxFiles(p);
    else if (name.endsWith(".docx")) yield p;
  }
}

describe("M1 lossless package round-trip", () => {
  const fixtures = [...docxFiles(join(HANDOFF, "fixtures"))];
  it("rezip keeps every entry byte-identical", () => {
    expect(fixtures.length).toBeGreaterThan(20);
    for (const docx of fixtures) {
      const entries = readPackage(readFileSync(docx));
      const out = readPackage(writePackage(entries));
      expect([...out.keys()].sort(), docx).toEqual([...entries.keys()].sort());
      for (const [name, bytes] of entries) {
        expect(Buffer.from(out.get(name)!).equals(Buffer.from(bytes)), `${docx} ${name}`).toBe(true);
      }
    }
  });

  it("document.xml through parse/serialize stays byte-identical (the edit path's untouched guarantee)", () => {
    for (const docx of fixtures) {
      const entries = readPackage(readFileSync(docx));
      const xml = dec(entries.get("word/document.xml")!);
      expect(serializeXml(parseXml(xml)), docx).toBe(xml);
    }
  });

  it("touched document.xml + untouched others: only document.xml content differs", () => {
    for (const docx of fixtures) {
      const entries = readPackage(readFileSync(docx));
      const edited = new Map(entries);
      const xml = decodePart(entries.get("word/document.xml")!);
      const tree = parseXml(xml);
      edited.set("word/document.xml", new TextEncoder().encode(serializeXml(tree)));
      const out = readPackage(writePackage(edited));
      for (const [name, bytes] of entries) {
        if (name === "word/document.xml") continue;
        expect(Buffer.from(out.get(name)!).equals(Buffer.from(bytes)), `${docx} ${name}`).toBe(true);
      }
      expect(new TextDecoder().decode(out.get("word/document.xml")!)).toBe(xml);
    }
  });

  it("entry ORDER survives a round-trip, including integer-like names (P2)", () => {
    // fflate's object API reorders integer-like keys before writing, so build a
    // raw stored-entry archive whose central directory order is 10, 2, word/ —
    // an Object-based read would return 2, 10, word/
    const raw = storedZip([
      ["10", encodePart("ten")],
      ["2", encodePart("two")],
      ["word/", new Uint8Array()],
    ]);
    const entries = readPackage(raw);
    expect([...entries.keys()]).toEqual(["10", "2", "word/"]);
    const out = readPackage(writePackage(entries));
    expect([...out.keys()]).toEqual(["10", "2", "word/"]);
  });

  it("UTF-8 entry names round-trip exactly (P6)", () => {
    const entries = new Map<string, Uint8Array>([
      ["word/é.xml", encodePart("café")],
      ["word/ok.xml", encodePart("plain")],
    ]);
    const out = readPackage(writePackage(entries));
    expect([...out.keys()]).toEqual(["word/é.xml", "word/ok.xml"]);
    expect(new TextDecoder().decode(out.get("word/é.xml")!)).toBe("café");
  });

  it("writePackage is byte-deterministic across independent invocations (P4)", () => {
    for (const docx of fixtures.slice(0, 5)) {
      const entries = readPackage(readFileSync(docx));
      expect(Buffer.from(writePackage(entries)).equals(Buffer.from(writePackage(entries))), docx).toBe(true);
    }
  });
});
