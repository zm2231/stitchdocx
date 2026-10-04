// Projection parity: our projectParagraphs must agree with the independent
// verification reader (imported here as an external oracle — test-only, never
// shipped in the engine) on every acceptance fixture, under every
// accept/reject view. Runs only when DOCX_ACCEPTANCE_DIR points at the
// acceptance folder.
import { readFileSync, readdirSync, statSync, existsSync } from "node:fs";
import { join } from "node:path";
import { describe, it, expect } from "vitest";
import { decodeAttrValue, decodePart, mainPartName, readPackage, parseXml, entryBytes } from "@stitchdocx/package-io";
import { openDocument } from "./document.js";
import { normalizeDocx } from "./normalize.js";
import { projectParagraphs, type View } from "./projection.js";
import type { BodyWalk } from "./body.js";

const ACCEPTANCE = process.env.DOCX_ACCEPTANCE_DIR ?? "";

// external oracle (JS without types), loaded from its local path
let reader: {
  project: (b: Uint8Array, v: unknown) => Array<Record<string, unknown>>;
};
if (ACCEPTANCE) {
  reader = await import(join(ACCEPTANCE, "harness", "ooxml.mjs"));
}

function* docxFiles(dir: string): Generator<string> {
  if (!existsSync(dir)) return;
  for (const name of readdirSync(dir)) {
    const p = join(dir, name);
    if (statSync(p).isDirectory()) yield* docxFiles(p);
    else if (name.endsWith(".docx")) yield p;
  }
}

/** every revision author present in a fixture — the "ours" side of a view is
 *  the document's own authors, so no author convention is assumed. Values are
 *  decoded like the core's attrVal, so a literal line break in an attribute
 *  normalizes the same way on both sides of the comparison. */
function revisionAuthors(bytes: Uint8Array): Set<string> {
  const entries = readPackage(bytes);
  const tree = parseXml(decodePart(entryBytes(entries.get(mainPartName(entries))!)));
  const authors = new Set<string>();
  const visit = (node: any): void => {
    for (const c of node.children) {
      if (c.kind !== "el") continue;
      if (c.name === "w:ins" || c.name === "w:del" || c.name === "w:moveFrom" || c.name === "w:moveTo" || c.name.endsWith("PrChange")) {
        const a = c.attrs.find((x: any) => x.name === "w:author");
        if (a) authors.add(decodeAttrValue(a.raw));
      }
      visit(c);
    }
  };
  visit(tree);
  return authors;
}

const VIEWS: Array<{ name: string; mine: "accept" | "reject"; reader: "accept" | "reject" }> = [
  { name: "accept/accept", mine: "accept", reader: "accept" },
  { name: "reject/accept", mine: "reject", reader: "accept" },
  { name: "accept/reject", mine: "accept", reader: "reject" },
  { name: "reject/reject", mine: "reject", reader: "reject" },
];

function ourProject(bytes: Uint8Array, view: View) {
  return openDocument(bytes).projectionRecords(view);
}

const readerProject = (bytes: Uint8Array, v: unknown) =>
  (reader.project as (b: Uint8Array, v: unknown) => Array<Record<string, unknown>>)(bytes, v);

const d = ACCEPTANCE ? describe : describe.skip;
d("M2 projection parity vs verification reader" + (ACCEPTANCE ? "" : " (skipped: DOCX_ACCEPTANCE_DIR unset)"), () => {
  const fixtures = [...docxFiles(join(ACCEPTANCE, "fixtures"))];
  it("fixtures present", () => expect(fixtures.length).toBeGreaterThan(20));

  for (const v of VIEWS) {
    it(`full projection matches under ${v.name} for every fixture`, () => {
      for (const docx of fixtures) {
        const bytes = new Uint8Array(readFileSync(docx));
        const authors = revisionAuthors(bytes);
        const mineView: View = (r) => (authors.has(r.author ?? "") ? v.mine : v.reader);
        const mine = ourProject(bytes, mineView);
        const theirs = readerProject(bytes, { decide: (a: string) => (authors.has(a) ? v.mine : v.reader) });
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

d("M2 projection restricted to some records" + (ACCEPTANCE ? "" : " (skipped: DOCX_ACCEPTANCE_DIR unset)"), () => {
  const fixtures = [...docxFiles(join(ACCEPTANCE, "fixtures"))];

  it("reads every record exactly as the full projection does, under every view", () => {
    for (const docx of fixtures) {
      const bytes = normalizeDocx(new Uint8Array(readFileSync(docx))).docx;
      const authors = revisionAuthors(bytes);
      for (const v of VIEWS) {
        const view: View = (r) => (authors.has(r.author ?? "") ? v.mine : v.reader);
        const doc = openDocument(bytes);
        const walk = (doc as unknown as { walk: BodyWalk }).walk;
        const full = projectParagraphs(walk, view, doc.styles);
        for (const rec of full) {
          if (!rec.addr) continue;
          const only = projectParagraphs(walk, view, doc.styles, new Set([rec.addr]));
          const label = `${docx} [${v.name}] ${rec.addr}`;
          expect(only.length, `${label} count`).toBe(1);
          const { node, merged, ...rest } = only[0]!;
          const { node: fullNode, merged: fullMerged, ...fullRest } = rec;
          expect(rest, label).toEqual(fullRest);
          expect(node, `${label} node`).toBe(fullNode);
          expect(merged, `${label} merged`).toEqual(fullMerged);
        }
      }
    }
  });
});
