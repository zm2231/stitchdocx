// Regression tests for insertParagraphAfter markup and resolution: an
// inserted paragraph at the end of a container must use Word's anchor-mark
// form so Reject All can remove it, and the engine's decide/verify must
// handle both that form and the mid-container form.
import { describe, expect, it } from "vitest";
import { encodePart, readPackage, writePackage, entryBytes } from "@stitchdocx/package-io";
import { normalizeDocx, openDocument, verifyOutput, type Op, type Receipt } from "./index.js";

const DATE = "2026-10-03T00:00:00Z";
const AUTHOR = "Reviewer A";

function docx(bodyChildrenXml: string): Uint8Array {
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"><w:body>${bodyChildrenXml}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;
  const entries = new Map<string, Uint8Array>([
    ["[Content_Types].xml", encodePart(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`)],
    ["word/document.xml", encodePart(xml)],
  ]);
  return writePackage(entries);
}

const P = (paraId: string, inner: string): string => `<w:p w14:paraId="${paraId}">${inner}</w:p>`;
const R = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

function runOn(bodyXml: string, ops: Op[]): { receipts: Receipt[]; out: Uint8Array } {
  const { docx: norm } = normalizeDocx(docx(bodyXml));
  const doc = openDocument(norm);
  const receipts = doc.apply(ops, { date: DATE });
  return { receipts, out: doc.save() };
}

function xmlOf(bytes: Uint8Array): string {
  return new TextDecoder().decode(entryBytes(readPackage(bytes).get("word/document.xml")!));
}

/** ids of every revision by `author` in the document part, in order */
const idsBy = (bytes: Uint8Array, author: string): string[] =>
  [...xmlOf(bytes).matchAll(/<w:(?:ins|del|moveFrom|moveTo|\w+PrChange)\b[^>]*>/g)]
    .filter((m) => m[0].includes(`w:author="${author}"`))
    .map((m) => /w:id="([^"]+)"/.exec(m[0])![1]!);

/** the source of one paragraph, from its paraId to its end tag */
function paraXml(xml: string, paraId: string): string {
  const at = xml.indexOf(`w14:paraId="${paraId}"`);
  expect(at).toBeGreaterThan(0);
  return xml.slice(at, xml.indexOf("</w:p>", at));
}

const insert = (addr: string, text: string): Op => ({ op: "insertParagraphAfter", ref: "ins", author: AUTHOR, addr, text });

describe("inserted paragraph at the end of a container", () => {
  // the anchor 10000000 is the LAST paragraph of the body
  const BODY_END = P("10000001", R("Earlier paragraph.")) + P("10000000", R("Final heading."));

  it("tracks the anchor's mark and records w:pPrChange, not its own mark", () => {
    const body = P("10000001", R("Earlier paragraph.")) + P("10000000", `<w:pPr><w:pStyle w:val="Heading1"/><w:rPr><w:b/></w:rPr></w:pPr>` + R("Final heading."));
    const { receipts, out } = runOn(body, [insert("10000000", "New clause.")]);
    expect(receipts[0]!.status).toBe("applied");
    const xml = xmlOf(out);
    // anchor mark tracked, first among the rPr children after its own props
    const anchor = paraXml(xml, "10000000");
    expect(anchor).toContain('<w:rPr><w:ins ');
    expect(anchor).toContain('<w:pStyle w:val="Heading1"/>');
    // the new paragraph carries a pPrChange whose saved pPr keeps the style
    // but drops the mark formatting, and no tracked mark of its own
    const inserted = paraXml(xml, receipts[0]!.insertedAddr!);
    expect(inserted).toContain('<w:pPrChange w:id="2" w:author="Reviewer A"');
    expect(inserted).toContain('<w:pPr><w:pStyle w:val="Heading1"/></w:pPr></w:pPrChange>');
    expect(inserted).not.toContain("<w:rPr>");
    // the receipt lists every revision the op wrote
    expect(receipts[0]!.changeIds).toHaveLength(3);
    expect(verifyOutput(normalizeDocx(docx(body)).docx, out, receipts).ok).toBe(true);
  });

  it("engine reject restores the anchor paragraph exactly, paraId included", () => {
    const { out } = runOn(BODY_END, [insert("10000000", "New clause.")]);
    const applied = openDocument(out);
    const insertedAddr = applied.projection().at(-1)!.addr;
    applied.decide([{ ref: "i", changeIds: idsBy(out, AUTHOR), decision: "reject" }]);
    const fin = applied.save();
    const paras = openDocument(fin).projection();
    expect(paras.map((p) => p.addr)).toEqual(["10000001", "10000000"]); // the anchor's own paraId survives
    expect(paras.map((p) => p.text)).toEqual(["Earlier paragraph.", "Final heading."]);
    const xml = xmlOf(fin);
    expect(xml).not.toContain(AUTHOR);
    expect(xml).not.toContain("w:pPrChange");
    expect(xml).not.toContain(`w14:paraId="${insertedAddr}"`);
  });

  it("engine accept leaves both paragraphs plain", () => {
    const { out } = runOn(BODY_END, [insert("10000000", "New clause.")]);
    const doc = openDocument(out);
    doc.decide([{ ref: "i", changeIds: idsBy(out, AUTHOR), decision: "accept" }]);
    const paras = openDocument(doc.save()).projection();
    expect(paras.map((p) => p.text)).toEqual(["Earlier paragraph.", "Final heading.", "New clause."]);
    const xml = xmlOf(doc.save());
    expect(xml).not.toContain(AUTHOR);
    expect(xml).not.toContain("w:pPrChange");
  });

  it("each insert in a chain tracks the mark before it, so any subset of decisions leaves exactly the kept paragraphs", () => {
    const body = P("10000001", R("Earlier paragraph.")) + P("10000000", `<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>` + R("Final heading."));
    const ops: Op[] = [
      { op: "insertParagraphAfter", ref: "a", author: "Reviewer A", addr: "10000000", text: "First." },
      { op: "insertParagraphAfter", ref: "b", author: "Reviewer B", addr: "10000000", text: "Second." },
    ];
    const { receipts, out } = runOn(body, ops);
    expect(receipts.map((r) => r.status)).toEqual(["applied", "applied"]);
    expect(receipts.map((r) => r.changeIds.length)).toEqual([3, 3]);
    expect(verifyOutput(normalizeDocx(docx(body)).docx, out, receipts).failures).toEqual([]);
    // the second insert owns the first's mark, not the anchor's
    expect(paraXml(xmlOf(out), receipts[0]!.insertedAddr!)).toContain('<w:rPr><w:ins w:id="' + receipts[1]!.changeIds[0] + '" w:author="Reviewer B"');
    for (const [a, b, want] of [
      ["accept", "accept", ["Earlier paragraph.", "Final heading.", "First.", "Second."]],
      ["accept", "reject", ["Earlier paragraph.", "Final heading.", "First."]],
      ["reject", "accept", ["Earlier paragraph.", "Final heading.", "Second."]],
      ["reject", "reject", ["Earlier paragraph.", "Final heading."]],
    ] as const) {
      const doc = openDocument(out);
      doc.decide([
        { ref: "a", changeIds: receipts[0]!.changeIds, decision: a },
        { ref: "b", changeIds: receipts[1]!.changeIds, decision: b },
      ]);
      const fin = doc.save();
      expect(openDocument(fin).projection().map((p) => p.text), `${a}/${b}`).toEqual(want);
      const xml = xmlOf(fin);
      expect(xml).not.toContain("Reviewer");
      // wherever the anchor's text lands, it has the anchor's properties
      const heading = xml.slice(xml.lastIndexOf("<w:p ", xml.indexOf("Final heading.")), xml.indexOf("Final heading."));
      expect(heading, `${a}/${b}`).toContain('<w:pStyle w:val="Heading1"/>');
    }
  });

  describe.each([
    { where: "body end", wrap: (anchor: string) => P("10000001", R("Earlier paragraph.")) + anchor, before: ["Earlier paragraph."], after: [] as string[] },
    { where: "mid body", wrap: (anchor: string) => anchor + P("10000009", R("Later paragraph.")), before: [] as string[], after: ["Later paragraph."] },
    { where: "cell end", wrap: (anchor: string) => `<w:tbl><w:tr><w:tc>${P("10000001", R("Cell start."))}${anchor}</w:tc></w:tr></w:tbl>` + P("10000009", R("After.")), before: ["Cell start."], after: ["After."] },
    { where: "content control end", wrap: (anchor: string) => `<w:sdt><w:sdtContent>${P("10000001", R("Control start."))}${anchor}</w:sdtContent></w:sdt>` + P("10000009", R("After.")), before: ["Control start."], after: ["After."] },
    { where: "before a table", wrap: (anchor: string) => anchor + `<w:tbl><w:tr><w:tc>${P("10000009", R("Cell."))}</w:tc></w:tr></w:tbl>`, before: [] as string[], after: ["Cell."] },
  ])("a chain at $where", ({ wrap, before, after }) => {
    it.each([1, 2, 3])("every accept/reject subset of %i inserts leaves exactly the kept paragraphs", (n) => {
      const body = wrap(P("10000000", `<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>` + R("Anchor.")));
      const texts = Array.from({ length: n }, (_, i) => `Insert ${i + 1}.`);
      const ops: Op[] = texts.map((text, i) => ({ op: "insertParagraphAfter", ref: `r${i}`, author: `Reviewer ${i}`, addr: "10000000", text }));
      const { receipts, out } = runOn(body, ops);
      expect(receipts.map((r) => r.status)).toEqual(texts.map(() => "applied"));
      expect(verifyOutput(normalizeDocx(docx(body)).docx, out, receipts).failures).toEqual([]);
      for (let mask = 0; mask < 1 << n; mask++) {
        const keep = texts.map((_, i) => (mask >> i) & 1 ? "accept" : "reject" as const);
        const doc = openDocument(out);
        doc.decide(receipts.map((r, i) => ({ ref: r.ref, changeIds: r.changeIds, decision: keep[i]! })));
        const fin = doc.save();
        const label = keep.join("/");
        expect(openDocument(fin).projection().map((p) => p.text), label)
          .toEqual([...before, "Anchor.", ...texts.filter((_, i) => keep[i] === "accept"), ...after]);
        const xml = xmlOf(fin);
        expect(xml, label).not.toContain("Reviewer");
        expect(xml.slice(xml.lastIndexOf("<w:p ", xml.indexOf("Anchor.")), xml.indexOf("Anchor.")), label).toContain('<w:pStyle w:val="Heading1"/>');
      }
    });
  });

  it("an insert directly before a table uses the anchor-mark form", () => {
    const body = P("10000001", R("Before the table.")) + `<w:tbl><w:tr><w:tc>${P("10000002", R("Cell."))}</w:tc></w:tr></w:tbl>`;
    const { receipts, out } = runOn(body, [insert("10000001", "Between.")]);
    expect(paraXml(xmlOf(out), "10000001")).toContain('<w:rPr><w:ins ');
    expect(xmlOf(out)).toContain("w:pPrChange");
    const doc = openDocument(out);
    doc.decide([{ ref: "i", changeIds: idsBy(out, AUTHOR), decision: "reject" }]);
    expect(openDocument(doc.save()).projection().map((p) => p.text)).toEqual(["Before the table.", "Cell."]);
  });

  it("refuses paragraph_not_editable when the anchor's mark is already tracked", () => {
    const body = P("10000001", R("Earlier paragraph.")) + P("10000000", `<w:pPr><w:rPr><w:ins w:id="7" w:author="Client Author" w:date="${DATE}"/></w:rPr></w:pPr>` + R("Tracked end."));
    const { receipts } = runOn(body, [insert("10000000", "New clause.")]);
    expect(receipts[0]!.status).toBe("refused");
    expect(receipts[0]!.reason).toBe("paragraph_not_editable");
  });

  it("reject never dissolves a pair whose next paragraph holds kept content", () => {
    // P1's tracked mark + P2's independently tracked property change by the
    // same author, with PLAIN (kept) runs in P2: the merge path must run, not
    // the dissolve — P2 survives with merged text and restored properties
    const body =
      P("10000001", `<w:pPr><w:rPr><w:ins w:id="7" w:author="Reviewer A" w:date="${DATE}"/></w:rPr></w:pPr>` + R("Anchor.")) +
      P("10000000", `<w:pPr><w:jc w:val="center"/><w:pPrChange w:id="8" w:author="Reviewer A" w:date="${DATE}"><w:pPr><w:jc w:val="right"/></w:pPr></w:pPrChange></w:pPr>` + R("Kept text."));
    const doc = openDocument(docx(body));
    doc.decide([{ ref: "p", changeIds: ["7", "8"], decision: "reject" }]);
    const out = doc.save();
    const paras = openDocument(out).projection();
    expect(paras.map((p) => p.text)).toEqual(["Anchor.Kept text."]);
    const xml = xmlOf(out);
    expect(xml).not.toContain("Reviewer A"); // both revisions resolved
    expect(xml).toContain('<w:jc w:val="right"/>'); // the saved props were restored
    expect(xml).not.toContain('<w:jc w:val="center"/>');
    expect(paraXml(xml, "10000000")).toBeTruthy(); // P2 survived on its own paraId
  });

  it("accepting with wording removes only the paired insert's pPrChange, never an unrelated one", () => {
    // positive control: a genuine container-end pair
    const body =
      P("10000001", `<w:pPr><w:rPr><w:ins w:id="7" w:author="Reviewer A" w:date="${DATE}"/></w:rPr></w:pPr>` + R("Anchor.")) +
      P("10000000", `<w:pPr><w:pPrChange w:id="8" w:author="Reviewer A" w:date="${DATE}"><w:pPr/></w:pPrChange></w:pPr><w:ins w:id="9" w:author="Reviewer A" w:date="${DATE}">` + R("Clause.") + `</w:ins>`);
    const doc = openDocument(docx(body));
    doc.decide([{ ref: "d", changeIds: ["7", "8", "9"], decision: "accept", text: "Clause." }]);
    const xml = xmlOf(doc.save());
    expect(xml).not.toContain("Reviewer A");
    expect(xml).not.toContain("w:pPrChange");
    expect(openDocument(doc.save()).projection().map((p) => p.text)).toEqual(["Anchor.", "Clause."]);

    // unrelated: same-author insertion-only edit and pPrChange in one paragraph
    const body2 =
      P("10000001", R("Plain paragraph.")) +
      P("10000000", `<w:pPr><w:jc w:val="center"/><w:pPrChange w:id="8" w:author="Reviewer A" w:date="${DATE}"><w:pPr><w:jc w:val="right"/></w:pPr></w:pPrChange></w:pPr>` + R("Kept ") + `<w:ins w:id="9" w:author="Reviewer A" w:date="${DATE}">` + R("words") + `</w:ins>` + R("."));
    const doc2 = openDocument(docx(body2));
    doc2.decide([{ ref: "d", changeIds: ["9"], decision: "accept", text: "wording", expectText: "" }]);
    const xml2 = xmlOf(doc2.save());
    expect(xml2).toContain("w:pPrChange"); // the property change is not ours to resolve
    expect(xml2).toContain('<w:jc w:val="center"/>');
    expect(openDocument(doc2.save()).projection().at(-1)!.text).toBe("Kept wording.");
  });
});

describe("mid-container inserts keep the previous form", () => {
  const MID = P("10000000", R("Anchor paragraph.")) + P("10000001", R("Next paragraph."));

  it("tracks the new paragraph's own mark and reject removes the paragraph", () => {
    const { receipts, out } = runOn(MID, [insert("10000000", "Inserted.")]);
    const xml = xmlOf(out);
    expect(paraXml(xml, receipts[0]!.insertedAddr!)).toContain("<w:pPr><w:rPr><w:ins ");
    expect(xml).not.toContain("w:pPrChange");
    const doc = openDocument(out);
    doc.decide([{ ref: "i", changeIds: idsBy(out, AUTHOR), decision: "reject" }]);
    const paras = openDocument(doc.save()).projection();
    expect(paras.map((p) => p.text)).toEqual(["Anchor paragraph.", "Next paragraph."]);
  });

  it("a mid-container insert's anchor stays byte-verified", () => {
    const body = P("10000000", `<w:pPr><w:jc w:val="left"/></w:pPr>` + R("Anchor paragraph.")) + P("10000001", R("Next paragraph."));
    const { docx: norm } = normalizeDocx(docx(body));
    const doc = openDocument(norm);
    const receipts = doc.apply([insert("10000000", "Inserted.")], { date: DATE });
    const out = doc.save();
    expect(verifyOutput(norm, out, receipts).ok).toBe(true);
    // the mid-container form never touches the anchor: a formatting tamper
    // on it must fail verification even though an op addressed the paragraph
    const entries = readPackage(out);
    const xml = new TextDecoder().decode(entryBytes(entries.get("word/document.xml")!));
    entries.set("word/document.xml", encodePart(xml.replace('<w:jc w:val="left"/>', '<w:jc w:val="center"/>')));
    expect(verifyOutput(norm, writePackage(entries), receipts).ok).toBe(false);
  });

  it("a span edit and an insert sharing one anchor verify consistently", () => {
    const body = P("10000000", `<w:pPr><w:jc w:val="left"/></w:pPr>` + R("Payments are due net-30.")) + P("10000001", R("Next paragraph."));
    const { docx: norm } = normalizeDocx(docx(body));
    const doc = openDocument(norm);
    const receipts = doc.apply([
      { op: "replace", ref: "r", author: AUTHOR, addr: "10000000", start: 15, end: 21, expectText: "net-30", text: "net-45" },
      insert("10000000", "Inserted."),
    ], { date: DATE });
    expect(verifyOutput(norm, doc.save(), receipts).ok).toBe(true);
  });

  it("a paragraph whose mark is already tracked can still be rejected by merge", () => {
    // pre-existing tracked insert (client redline): mark on the new paragraph, mid-container
    const body = P("10000000", R("Anchor.")) + P("10000002", `<w:pPr><w:rPr><w:ins w:id="7" w:author="Client Author" w:date="${DATE}"/></w:rPr></w:pPr><w:ins w:id="8" w:author="Client Author" w:date="${DATE}">` + R("Their clause.") + `</w:ins>`) + P("10000001", R("Next."));
    const doc = openDocument(docx(body));
    doc.decide([{ ref: "c", changeIds: ["7", "8"], decision: "reject" }]);
    expect(openDocument(doc.save()).projection().map((p) => p.text)).toEqual(["Anchor.", "Next."]);
  });
});

describe("verify catches formatting tamper on untouched paragraphs", () => {
  it("fails when an untouched paragraph's properties are altered", () => {
    const body = P("10000001", R("Payments are due net-30.")) + P("10000000", `<w:pPr><w:jc w:val="left"/></w:pPr>` + R("Untouched."));
    const { docx: norm } = normalizeDocx(docx(body));
    const doc = openDocument(norm);
    const receipts = doc.apply([{ op: "replace", ref: "f", author: AUTHOR, addr: "10000001", start: 15, end: 21, expectText: "net-30", text: "net-45" }], { date: DATE });
    const out = doc.save();
    expect(verifyOutput(norm, out, receipts).ok).toBe(true);
    const entries = readPackage(out);
    const xml = new TextDecoder().decode(entryBytes(entries.get("word/document.xml")!));
    entries.set("word/document.xml", encodePart(xml.replace('<w:jc w:val="left"/>', '<w:jc w:val="center"/>')));
    expect(verifyOutput(norm, writePackage(entries), receipts).ok).toBe(false);
  });
});
