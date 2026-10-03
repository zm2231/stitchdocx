// M1: lossless package round-trip — re-saving an untouched package keeps
// every entry's content byte-identical, and the save path (parse → serialize
// word/document.xml → rezip) is byte-identical for the edited part too.
import { readFileSync, readdirSync, statSync, existsSync } from "node:fs";
import { join } from "node:path";
import { homedir } from "node:os";
import { describe, it, expect } from "vitest";
import { decodePart, decodePart as dec, readPackage, writePackage } from "./zip.js";
import { parseXml, serializeXml } from "./xml.js";

const HANDOFF = process.env.STITCH_HANDOFF ?? join(homedir(), "Downloads", "scope-docx-engine-handoff");

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
      // pretend an edit happened: re-serialize through the tree
      edited.set("word/document.xml", new TextEncoder().encode(serializeXml(tree)));
      const out = readPackage(writePackage(edited));
      for (const [name, bytes] of entries) {
        if (name === "word/document.xml") continue;
        expect(Buffer.from(out.get(name)!).equals(Buffer.from(bytes)), `${docx} ${name}`).toBe(true);
      }
      expect(new TextDecoder().decode(out.get("word/document.xml")!)).toBe(xml);
    }
  });
});
