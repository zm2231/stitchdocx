// D1 spike: byte-fidelity of the hand-written parser over every XML part of
// every acceptance fixture, contrasted with a saxes-based tree (the reader's
// parser), plus fuzz and timing. The decision: the hand-written parser is the
// engine's XML layer (saxes resolves entities and normalizes forms, so it
// cannot round-trip bytes). Runs only when DOCX_ACCEPTANCE_DIR points at the
// acceptance folder.
import { readFileSync, readdirSync, statSync, existsSync } from "node:fs";
import { join } from "node:path";
import { describe, it, expect } from "vitest";
import { SaxesParser } from "saxes";
import { decodePart, readPackage } from "./zip.js";
import { parseXml, serializeXml } from "./xml.js";

const ACCEPTANCE = process.env.DOCX_ACCEPTANCE_DIR ?? "";

function* docxFiles(dir: string): Generator<string> {
  if (!existsSync(dir)) return;
  for (const name of readdirSync(dir)) {
    const p = join(dir, name);
    if (statSync(p).isDirectory()) yield* docxFiles(p);
    else if (name.endsWith(".docx")) yield p;
  }
}

const fixtures = [...docxFiles(join(ACCEPTANCE, "fixtures"))];
const xmlParts = (docx: string): Array<[string, string]> => {
  const entries = readPackage(readFileSync(docx));
  return [...entries]
    .filter(([name]) => name.endsWith(".xml") || name.endsWith(".rels"))
    .map(([name, bytes]) => [name, decodePart(bytes)]);
};

// saxes contrast: the reader's parser + a plain serializer
function saxesRoundTrip(xml: string): string {
  const root = { name: "#doc", attrs: {} as Record<string, string>, children: [] as any[], parent: null as any };
  const stack = [root];
  const p = new SaxesParser({ xmlns: false, position: false });
  p.on("opentag", (t: any) => {
    const el = { name: t.name, attrs: { ...t.attributes }, children: [], parent: stack[stack.length - 1] };
    stack[stack.length - 1]!.children.push(el);
    if (!t.isSelfClosing) stack.push(el);
  });
  p.on("closetag", (t: any) => {
    if (!t.isSelfClosing) stack.pop();
  });
  p.on("text", (v: string) => stack[stack.length - 1]!.children.push({ text: v }));
  p.write(xml).close();
  const ser = (n: any): string => {
    if ("text" in n) return n.text.replace(/&/g, "&amp;").replace(/</g, "&lt;");
    const attrs = Object.entries(n.attrs).map(([k, v]) => ` ${k}="${String(v).replace(/"/g, "&quot;")}"`).join("");
    return `<${n.name}${attrs}>${n.children.map(ser).join("")}</${n.name}>`;
  };
  return root.children.map(ser).join("");
}

const d = ACCEPTANCE ? describe : describe.skip;
d("D1 parser spike" + (ACCEPTANCE ? "" : " (skipped: DOCX_ACCEPTANCE_DIR unset)"), () => {
  it("hand-written parser: serialize(parse(x)) === x for every XML part of every fixture", () => {
    expect(fixtures.length).toBeGreaterThan(20);
    let parts = 0;
    let bytes = 0;
    const bad: string[] = [];
    for (const docx of fixtures) {
      for (const [name, xml] of xmlParts(docx)) {
        parts++;
        bytes += xml.length;
        let out: string;
        try {
          out = serializeXml(parseXml(xml));
        } catch (e) {
          bad.push(`${docx} ${name}: parse threw ${(e as Error).message}`);
          continue;
        }
        if (out !== xml) bad.push(`${docx} ${name}: ${xml.length} -> ${out.length} bytes, first diff at ${firstDiff(xml, out)}`);
      }
    }
    console.log(`fidelity: ${parts - bad.length}/${parts} parts byte-identical (${(bytes / 1024).toFixed(0)} KiB)`);
    // pinned denominator: a regression in fixture discovery or part selection must fail loudly
    expect(parts).toBe(400);
    expect(bad).toEqual([]);
  });

  it("saxes contrast: records where entity/attr forms break byte fidelity", () => {
    let parts = 0;
    let identical = 0;
    for (const docx of fixtures) {
      for (const [, xml] of xmlParts(docx)) {
        parts++;
        if (saxesRoundTrip(xml) === xml) identical++;
      }
    }
    console.log(`saxes baseline: ${identical}/${parts} parts byte-identical (expected: lower; entities are decoded on parse)`);
    expect(parts).toBeGreaterThan(0);
  });

  it("timing: parse+serialize the largest fixture part", () => {
    let biggest = { docx: "", name: "", xml: "" };
    for (const docx of fixtures) {
      for (const [name, xml] of xmlParts(docx)) {
        if (xml.length > biggest.xml.length) biggest = { docx, name, xml };
      }
    }
    const t0 = performance.now();
    const n = 20;
    for (let i = 0; i < n; i++) serializeXml(parseXml(biggest.xml));
    const ms = (performance.now() - t0) / n;
    console.log(`largest part ${biggest.name} (${(biggest.xml.length / 1024).toFixed(0)} KiB, ${biggest.docx.split("/").pop()}): parse+serialize ${ms.toFixed(1)}ms`);
    expect(ms).toBeLessThan(250);
  });

  it("fuzz: identity on generated adversarial XML", () => {
    // every fragment is self-contained (no dangling open/close across picks)
    const bodyFrags = [
      `<w:p w14:paraId='0A1B2C3D' w14:textId="0A1B2C3E"><w:t xml:space="preserve">a &amp; b &lt;c&gt; &#8220;q&#8221; &#x1F600;</w:t></w:p>`,
      `<w:sectPr/>`,
      `<w:sectPr ></w:sectPr >`,
      `<w:p/>`,
      `<w:p w14:paraId="0A1B2C3F" />`,
      `<!-- comment with <tags> and -- not -->`,
      `<![CDATA[a < b && c > d]]>`,
      `<?mso-application progid="Word.Document"?>`,
      `text &gt; more &amp;amp; text`,
      `<empty></empty>`,
      `<a x='1' y = "2" z = "3"><b q="ends-in-slash/"></b></a>`,
    ];
    const prolog = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;
    const open = `<w:doc xmlns:w="urn:x" xmlns:w14="urn:y" mc:Ignorable="w14" xmlns:mc="urn:mc">`;
    for (let i = 0; i < 500; i++) {
      const body = Array.from({ length: 2 + (i % 7) }, (_, k) => bodyFrags[(i * 3 + k) % bodyFrags.length]!).join("");
      const doc = i % 2 ? prolog + open + body + `</w:doc>` : open + body + `</w:doc>`;
      expect(serializeXml(parseXml(doc)), doc).toBe(doc);
    }
  });
});

function firstDiff(a: string, b: string): number {
  const n = Math.min(a.length, b.length);
  for (let i = 0; i < n; i++) if (a[i] !== b[i]) return i;
  return n;
}
