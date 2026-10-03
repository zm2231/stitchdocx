// Regression tests for the real-document evaluation findings (D1): an
// inserted paragraph at the end of a container must use Word's anchor-mark
// form so Reject All can remove it, and the engine's own decide/verify must
// handle both that form and the mid-container form.
import { describe, expect, it } from "vitest";
import { encodePart, readPackage, writePackage } from "@stitchdocx/package-io";
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
  return new TextDecoder().decode(readPackage(bytes).get("word/document.xml")!);
}

/** the source of one paragraph, from its paraId to its end tag */
function paraXml(xml: string, paraId: string): string {
  const at = xml.indexOf(`w14:paraId="${paraId}"`);
  expect(at).toBeGreaterThan(0);
  return xml.slice(at, xml.indexOf("</w:p>", at));
}

const insert = (addr: string, text: string): Op => ({ op: "insertParagraphAfter", ref: "ins", author: AUTHOR, addr, text });

describe("D1: inserted paragraph at the end of a container", () => {
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
    applied.decide([{ author: AUTHOR, decision: "reject" }]);
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
    doc.decide([{ author: AUTHOR, decision: "accept" }]);
    const paras = openDocument(doc.save()).projection();
    expect(paras.map((p) => p.text)).toEqual(["Earlier paragraph.", "Final heading.", "New clause."]);
    const xml = xmlOf(doc.save());
    expect(xml).not.toContain(AUTHOR);
    expect(xml).not.toContain("w:pPrChange");
  });

  it("a chain of inserts tracks anchor and intermediates; only the last carries the pPrChange", () => {
    const { receipts, out } = runOn(BODY_END, [insert("10000000", "First."), insert("10000000", "Second.")]);
    expect(receipts.map((r) => r.status)).toEqual(["applied", "applied"]);
    expect(receipts[0]!.changeIds).toHaveLength(2);
    expect(receipts[1]!.changeIds).toHaveLength(3);
    expect(verifyOutput(normalizeDocx(docx(BODY_END)).docx, out, receipts).ok).toBe(true);
    const doc = openDocument(out);
    doc.decide([{ author: AUTHOR, decision: "reject" }]);
    const paras = openDocument(doc.save()).projection();
    expect(paras.map((p) => p.text)).toEqual(["Earlier paragraph.", "Final heading."]);
  });

  it("an insert directly before a table uses the anchor-mark form", () => {
    const body = P("10000001", R("Before the table.")) + `<w:tbl><w:tr><w:tc>${P("10000002", R("Cell."))}</w:tc></w:tr></w:tbl>`;
    const { receipts, out } = runOn(body, [insert("10000001", "Between.")]);
    expect(paraXml(xmlOf(out), "10000001")).toContain('<w:rPr><w:ins ');
    expect(xmlOf(out)).toContain("w:pPrChange");
    const doc = openDocument(out);
    doc.decide([{ author: AUTHOR, decision: "reject" }]);
    expect(openDocument(doc.save()).projection().map((p) => p.text)).toEqual(["Before the table.", "Cell."]);
  });

  it("refuses paragraph_not_editable when the anchor's mark is already tracked", () => {
    const body = P("10000001", R("Earlier paragraph.")) + P("10000000", `<w:pPr><w:rPr><w:ins w:id="7" w:author="Client Author" w:date="${DATE}"/></w:rPr></w:pPr>` + R("Tracked end."));
    const { receipts } = runOn(body, [insert("10000000", "New clause.")]);
    expect(receipts[0]!.status).toBe("refused");
    expect(receipts[0]!.reason).toBe("paragraph_not_editable");
  });
});

describe("D1: mid-container inserts keep the previous form", () => {
  const MID = P("10000000", R("Anchor paragraph.")) + P("10000001", R("Next paragraph."));

  it("tracks the new paragraph's own mark and reject removes the paragraph", () => {
    const { receipts, out } = runOn(MID, [insert("10000000", "Inserted.")]);
    const xml = xmlOf(out);
    expect(paraXml(xml, receipts[0]!.insertedAddr!)).toContain("<w:pPr><w:rPr><w:ins ");
    expect(xml).not.toContain("w:pPrChange");
    const doc = openDocument(out);
    doc.decide([{ author: AUTHOR, decision: "reject" }]);
    const paras = openDocument(doc.save()).projection();
    expect(paras.map((p) => p.text)).toEqual(["Anchor paragraph.", "Next paragraph."]);
  });

  it("a paragraph whose mark is already tracked can still be rejected by merge", () => {
    // pre-existing tracked insert (client redline): mark on the new paragraph, mid-container
    const body = P("10000000", R("Anchor.")) + P("10000002", `<w:pPr><w:rPr><w:ins w:id="7" w:author="Client Author" w:date="${DATE}"/></w:rPr></w:pPr><w:ins w:id="8" w:author="Client Author" w:date="${DATE}">` + R("Their clause.") + `</w:ins>`) + P("10000001", R("Next."));
    const doc = openDocument(docx(body));
    doc.decide([{ author: "Client Author", decision: "reject" }]);
    expect(openDocument(doc.save()).projection().map((p) => p.text)).toEqual(["Anchor.", "Next."]);
  });
});
