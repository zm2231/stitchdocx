// Projection parity: our projectParagraphs must agree with the independent
// harness reader (imported here as an external oracle — test-only, never
// shipped in the engine) on every fixture, under every accept/reject view.
import { readFileSync, readdirSync, statSync, existsSync } from "node:fs";
import { join } from "node:path";
import { homedir } from "node:os";
import { describe, it, expect } from "vitest";
import { decodePart, readPackage, parseXml, kid } from "@stitchdocx/package-io";
import { walkBody } from "./body.js";
import { projectParagraphs, type View } from "./projection.js";

// external oracle (JS without types), loaded from its local path
const reader: {
  project: (b: Uint8Array, v: unknown) => Array<Record<string, unknown>>;
} = await import(join(homedir(), "Downloads", "scope-docx-engine-handoff", "harness", "ooxml.mjs"));

const HANDOFF = process.env.STITCH_HANDOFF ?? join(homedir(), "Downloads", "scope-docx-engine-handoff");

function* docxFiles(dir: string): Generator<string> {
  if (!existsSync(dir)) return;
  for (const name of readdirSync(dir)) {
    const p = join(dir, name);
    if (statSync(p).isDirectory()) yield* docxFiles(p);
    else if (name.endsWith(".docx")) yield p;
  }
}

const ours = (a: string | null) => (a ?? "").startsWith("Scope Optimizer");

const VIEWS: Array<{ name: string; mine: View; reader: unknown }> = [
  { name: "accept/accept", mine: (r) => (ours(r.author) ? "accept" : "accept"), reader: { ours: "accept", theirs: "accept" } },
  { name: "reject/accept", mine: (r) => (ours(r.author) ? "reject" : "accept"), reader: { ours: "reject", theirs: "accept" } },
  { name: "accept/reject", mine: (r) => (ours(r.author) ? "accept" : "reject"), reader: { ours: "accept", theirs: "reject" } },
  { name: "reject/reject", mine: () => "reject", reader: { ours: "reject", theirs: "reject" } },
];

function ourProject(bytes: Uint8Array, view: View) {
  const entries = readPackage(bytes);
  const tree = parseXml(decodePart(entries.get("word/document.xml")!));
  const root = kid(tree as any, "w:document")!;
  return projectParagraphs(walkBody(root), view);
}

const readerProject = (bytes: Uint8Array, v: unknown) =>
  (reader.project as (b: Uint8Array, v: unknown) => Array<Record<string, unknown>>)(bytes, v);

describe("M2 projection parity vs harness reader", () => {
  const fixtures = [...docxFiles(join(HANDOFF, "fixtures"))];
  it("fixtures present", () => expect(fixtures.length).toBeGreaterThan(20));

  for (const v of VIEWS) {
    it(`full projection matches under ${v.name} for every fixture`, () => {
      for (const docx of fixtures) {
        const bytes = new Uint8Array(readFileSync(docx));
        const mine = ourProject(bytes, v.mine);
        const theirs = readerProject(bytes, v.reader);
        expect(mine.length, `${docx} [${v.name}] count`).toBe(theirs.length);
        for (let i = 0; i < mine.length; i++) {
          const m = mine[i]!;
          const t = theirs[i]!;
          const label = `${docx} [${v.name}] #${i}`;
          expect(m.text, `${label} text`).toBe(t.text);
          expect(m.addr, `${label} addr`).toBe(t.paraId);
          expect(m.container, `${label} container`).toBe(t.container);
          expect(m.table, `${label} table`).toEqual(t.table);
          expect(m.style, `${label} style`).toBe(t.style);
          expect(m.numbered, `${label} numbered`).toBe(t.numbered);
        }
      }
    });
  }
});
