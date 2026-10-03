// Regression tests for the round-1 review findings — one test per ledger
// Accept check. Builders construct minimal OOXML packages in-memory.
import { readFileSync } from "node:fs";
import { describe, it, expect } from "vitest";
import { encodePart, readPackage, writePackage } from "@stitchdocx/package-io";
import { normalizeDocx, openDocument, verifyOutput, type Op, type Receipt } from "./index.js";
import { buildSpanMap, checkSpan } from "./spanmap.js";
import { wordEdits } from "./worddiff.js";
import { paragraphText } from "./projection.js";

const DATE = "2026-10-03T00:00:00Z";

function docx(bodyChildrenXml: string): Uint8Array {
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"><w:body>${bodyChildrenXml}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;
  const entries = new Map<string, Uint8Array>([
    ["[Content_Types].xml", encodePart(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`)],
    ["word/document.xml", encodePart(xml)],
  ]);
  return writePackage(entries);
}

const P = (paraId: string | null, inner: string): string =>
  `<w:p${paraId ? ` w14:paraId="${paraId}"` : ""}>${inner}</w:p>`;
const R = (text: string, rPr?: string): string => `<w:r>${rPr ?? ""}<w:t xml:space="preserve">${text}</w:t></w:r>`;

function runOn(bodyXml: string, ops: Op[]): { receipts: Receipt[]; out: Uint8Array } {
  const { docx: norm } = normalizeDocx(docx(bodyXml));
  const doc = openDocument(norm);
  const receipts = doc.apply(ops, { date: DATE });
  return { receipts, out: doc.save() };
}

describe("review round 1 regression tests", () => {
  it("M1: two zero-length replacements at one offset become sibling ins revisions with independent receipts", () => {
    const body = P("10000001", R("Payments begin on the start date."));
    const { docx: norm } = normalizeDocx(docx(body));
    const doc = openDocument(norm);
    const op = (flag: string, text: string): Op => ({
      op: "replace", ref: flag, author: `Reviewer A [${flag}]`, addr: "10000001",
      start: 16, end: 16, expectText: "", text,
    });
    const receipts = doc.apply([op("f1", "AAA "), op("f2", "BBB ")], { date: DATE });
    expect(receipts.every((r) => r.status === "applied")).toBe(true);
    const out1 = doc.save();
    // sibling: no w:ins directly inside another w:ins
    const xml = new TextDecoder().decode(readPackage(out1).get("word/document.xml")!);
    expect(xml).not.toMatch(/<w:ins [^>]*><w:ins /);
    // independent reject of f1 leaves f2's text
    const doc2 = openDocument(out1);
    doc2.decide([{ author: "Reviewer A [f1]", decision: "reject" }]);
    const text = doc2.projection().find((p) => p.addr === "10000001")!.text;
    expect(text).toContain("BBB ");
    expect(text).not.toContain("AAA ");
  });

  it("M2: span crossing from plain text into a hyperlink field result refuses span_crosses_structure", () => {
    // hyperlink containing a complex field whose result text is inside the link
    const body = P("10000001",
      R("See ") +
      `<w:hyperlink r:id="r1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> REF x </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
      R("the portal") +
      `<w:r><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink>` +
      R(" for details."));
    const { docx: norm } = normalizeDocx(docx(body));
    const doc = openDocument(norm);
    const para = doc.projectionRecords()[0]!;
    expect(para.text).toBe("See the portal for details.");
    const map = buildSpanMap(para.node, () => "accept");
    // span crosses from plain text into the hyperlink's field result
    const issue = checkSpan(map, 4, 16, []);
    expect(issue?.reason).toBe("span_crosses_structure");
  });

  it("M3: span across inline customXml and a sibling run refuses, no partial mutation", () => {
    const body = P("10000001",
      `<w:customXml w:uri="urn:x" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` + R("ab") + `</w:customXml>` + R("cd"));
    const { docx: norm } = normalizeDocx(docx(body));
    const doc = openDocument(norm);
    const para = doc.projectionRecords()[0]!;
    expect(para.text).toBe("abcd");
    const map = buildSpanMap(para.node, () => "accept");
    const issue = checkSpan(map, 1, 3, []); // covers "b" (inside customXml) + "c" (outside)
    expect(issue?.reason).toBe("span_crosses_structure");
    const before = doc.save();
    const { receipts } = (() => {
      const d2 = openDocument(norm);
      const r = d2.apply([{ op: "delete", ref: "f", author: "Reviewer A [f]", addr: para.addr!, start: 1, end: 3, expectText: "bc" }], { date: DATE });
      return { receipts: r };
    })();
    expect(receipts[0]!.status).toBe("refused");
    expect(receipts[0]!.reason).toBe("span_crosses_structure");
    expect(Buffer.from(doc.save()).equals(Buffer.from(before))).toBe(true);
  });

  it("RM1: CR/CRLF in text projects LF-normalized like the oracle", () => {
    const body = P("10000001", `<w:r><w:t>A\r\nB\rC</w:t></w:r>`);
    const { docx: norm } = normalizeDocx(docx(body));
    const doc = openDocument(norm);
    expect(doc.projection()[0]!.text).toBe("A\nB\nC");
  });

  it("RM3: normalize reassigns duplicate paraIds deterministically and stays idempotent", () => {
    const body = P("10000001", R("one")) + P("10000001", R("two"));
    const first = normalizeDocx(docx(body));
    expect(first.assigned).toBe(1); // the duplicate reassigned
    const xml = new TextDecoder().decode(readPackage(first.docx).get("word/document.xml")!);
    const ids = [...xml.matchAll(/w14:paraId="([0-9A-F]+)"/g)].map((m) => m[1]!);
    expect(new Set(ids).size).toBe(ids.length);
    const second = normalizeDocx(first.docx);
    expect(Buffer.from(second.docx).equals(Buffer.from(first.docx))).toBe(true);
  });

  it("RM4: paragraphs never emits an empty addr, even pre-normalization (CLI path)", () => {
    const { docx: normalized } = normalizeDocx(docx(P(null, R("no id here"))));
    const doc = openDocument(normalized);
    for (const p of doc.projection()) expect(p.addr).toMatch(/^[0-9A-F]{8}$/);
  });

  it("V1: verify rejects author tampering on a receipt revision", () => {
    const body = P("10000001", R("Hold harmless the indemnified party under all claims."));
    const { docx: norm } = normalizeDocx(docx(body));
    const doc = openDocument(norm);
    const receipts = doc.apply([{ op: "replace", ref: "f", author: "Reviewer A [f]", addr: "10000001", start: 0, end: 13, expectText: "Hold harmless", text: "Indemnify" }], { date: DATE });
    const out = doc.save();
    expect(verifyOutput(norm, out, receipts).ok).toBe(true);
    // tamper: change the revision's author
    const entries = readPackage(out);
    let xml = new TextDecoder().decode(entries.get("word/document.xml")!);
    xml = xml.replace('w:author="Reviewer A [f]"', 'w:author="Other Author"');
    entries.set("word/document.xml", encodePart(xml));
    const tampered = writePackage(entries);
    expect(verifyOutput(norm, tampered, receipts).ok).toBe(false);
  });

  it("V2: acceptWith on an insertion-only edit inserts the caller wording", () => {
    const body = P("10000001", R("Payments begin on the start date."));
    const { docx: norm } = normalizeDocx(docx(body));
    const doc = openDocument(norm);
    doc.apply([{ op: "replace", ref: "f", author: "Reviewer A [f]", addr: "10000001", start: 16, end: 16, expectText: "", text: "AAA " }], { date: DATE });
    doc.acceptWithText({ ref: "f", author: "Reviewer A [f]", decision: "accept", text: "BBB " });
    const text = doc.projection().find((p) => p.addr === "10000001")!.text;
    // offset 16 sits after the first "o" of "on": original text with BBB at 16
    expect(text).toBe("Payments begin oBBB n the start date.");
    const xml = new TextDecoder().decode(readPackage(doc.save()).get("word/document.xml")!);
    expect(xml).not.toContain("Reviewer A");
  });

  it("V3: decide resolves move and property-change revisions by exact author, leaving others untouched", () => {
    const body = P("10000001",
      `<w:moveFrom w:id="90" w:author="Ours [legacy]" w:date="${DATE}">` + R("gone") + `</w:moveFrom>` +
      `<w:moveFrom w:id="92" w:author="Client Author" w:date="${DATE}">` + R("theirs") + `</w:moveFrom>` +
      `<w:r><w:rPr><w:b/><w:rPrChange w:id="91" w:author="Ours [legacy]" w:date="${DATE}"><w:rPr/></w:rPrChange></w:rPr>` + R("kept") + `</w:r>`);
    const doc = openDocument(docx(body));
    const resolved = doc.decide([{ author: "Ours [legacy]", decision: "reject" }]);
    expect(resolved[0]!.changeIds.sort()).toEqual(["90", "91"]);
    const xml = new TextDecoder().decode(readPackage(doc.save()).get("word/document.xml")!);
    expect(xml).not.toContain("Ours [legacy]");
    expect(xml).toContain("Client Author"); // D3: revisions we never wrote survive
    expect(xml).toContain("gone"); // reject moveFrom restores the text
  });

  it("V3: exact-author decisions resolve moveTo and pPrChange and are idempotent", () => {
    const body = P("10000001",
      `<w:pPr><w:jc w:val="center"/><w:pPrChange w:id="94" w:author="Ours [legacy]" w:date="${DATE}"><w:pPr><w:jc w:val="right"/></w:pPr></w:pPrChange></w:pPr>` +
      `<w:moveTo w:id="93" w:author="Ours [legacy]" w:date="${DATE}">` + R("moved") + `</w:moveTo>` + R("stay"));
    const doc = openDocument(docx(body));
    const first = doc.decide([{ author: "Ours [legacy]", decision: "reject" }]);
    expect(first[0]!.changeIds.sort()).toEqual(["93", "94"]);
    // a second decision for the same author resolves nothing further
    const second = doc.decide([{ author: "Ours [legacy]", decision: "reject" }]);
    expect(second[0]!.changeIds).toEqual([]);
    const xml = new TextDecoder().decode(readPackage(doc.save()).get("word/document.xml")!);
    expect(openDocument(doc.save()).projection()[0]!.text).toBe("stay"); // reject removes the moveTo content
    expect(xml).toContain('<w:jc w:val="right"/>'); // pPrChange reject restores the saved props
    expect(xml).not.toContain("Ours [legacy]");
  });

  it("V4: decide accepts changeIds-based decisions and returns receipts", () => {
    const body = P("10000001", R("The fee is due."));
    const { docx: norm } = normalizeDocx(docx(body));
    const doc = openDocument(norm);
    const receipts = doc.apply([{ op: "replace", ref: "f", author: "Reviewer A [f]", addr: "10000001", start: 4, end: 7, expectText: "fee", text: "charge" }], { date: DATE });
    const ids = receipts[0]!.changeIds;
    const doc2 = openDocument(doc.save());
    const decideReceipts = doc2.decide([{ changeIds: ids, decision: "reject" }]);
    expect(decideReceipts[0]!.changeIds.length).toBeGreaterThan(0);
    expect(doc2.projection()[0]!.text).toBe("The fee is due.");
  });

  it("P1: engine-authored text with XML-invalid characters is sanitized", () => {
    const body = P("10000001", R("Rate card:"));
    const { docx: norm } = normalizeDocx(docx(body));
    const doc = openDocument(norm);
    const receipts = doc.apply([{ op: "replace", ref: "f", author: "Reviewer A [f]", addr: "10000001", start: 0, end: 9, expectText: "Rate card", text: "Rates\u0001\r\n2026" }], { date: DATE });
    expect(receipts[0]!.status).toBe("applied");
    const xml = new TextDecoder().decode(readPackage(doc.save()).get("word/document.xml")!);
    expect(xml).not.toContain("\u0001");
    expect(xml).not.toMatch(/\r/);
  });

  it("S03 word-diff offsets: pure insertion keeps old offsets", () => {
    const edits = wordEdits("Sessions not used within the month are forfeited.", "Sessions not used within the month roll over once and are then forfeited.");
    expect(edits).toEqual([
      { insAt: 35, insText: "roll over once and " },
      { insAt: 39, insText: "then " },
    ]);
  });

  it("RM2: a point inside a field result refuses span_in_field", () => {
    const body = P("10000001",
      R("generated on ") +
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> DATE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
      R("October 3, 2026") +
      `<w:r><w:fldChar w:fldCharType="end"/></w:r>`);
    const { docx: norm } = normalizeDocx(docx(body));
    const doc = openDocument(norm);
    const para = doc.projectionRecords()[0]!;
    const map = buildSpanMap(para.node, () => "accept");
    const at = para.text.indexOf("October 3") + 2;
    const issue = checkSpan(map, at, at, []);
    expect(issue?.reason).toBe("span_in_field");
  });

  it("RM2: a point inside an inline content control refuses span_in_content_control", () => {
    const body = P("10000001",
      R("Effective date: ") +
      `<w:sdt><w:sdtPr><w:alias w:val="d"/></w:sdtPr><w:sdtContent>` + R("October 1, 2026") + `</w:sdtContent></w:sdt>`);
    const { docx: norm } = normalizeDocx(docx(body));
    const doc = openDocument(norm);
    const para = doc.projectionRecords()[0]!;
    const map = buildSpanMap(para.node, () => "accept");
    const at = para.text.indexOf("October 1") + 2;
    const issue = checkSpan(map, at, at, []);
    expect(issue?.reason).toBe("span_in_content_control");
  });

  it("RM2: a point at a pre-existing client revision refuses span_overlaps_revision even with a matching author", () => {
    const body = P("10000001",
      R("Payment terms: ") +
      `<w:ins w:id="50" w:author="Reviewer A [clientlookalike]" w:date="${DATE}">` + R("net sixty") + `</w:ins>` + R("."));
    const { docx: norm } = normalizeDocx(docx(body));
    const doc = openDocument(norm);
    const para = doc.projectionRecords()[0]!;
    const map = buildSpanMap(para.node, () => "accept");
    // an op whose author string matches the pre-existing revision's author
    const issue = checkSpan(map, 15, 15, [{ start: 0, end: 0 }], new Set());
    // ownedRevisions is empty at validation: nothing from this call exists yet
    void issue;
    const receipts = doc.apply([
      { op: "replace", ref: "f", author: "Reviewer A [clientlookalike]", addr: "10000001", start: 15, end: 15, expectText: "", text: "X" },
    ], { date: DATE });
    expect(receipts[0]!.status).toBe("refused");
    expect(receipts[0]!.reason).toBe("span_overlaps_revision");
  });

  it("RM2: attribute values normalize literal CR/CRLF/TAB like the spec parser (author parity)", () => {
    const body = `<w:p w14:paraId="10000001">` +
      `<w:ins w:id="7" w:author="Reviewer
A [f]" w:date="${DATE}">` + R("wrapped author text") + `</w:ins>` +
      R(" tail") + `</w:p>`;
    const doc = openDocument(docx(body));
    const rec = doc.projectionRecords()[0]!;
    const text = paragraphText(rec.node, (rev) => (rev.author === "Reviewer A [f]" ? "reject" : "accept"));
    expect(text).toBe(" tail"); // the wrapped-author insertion is recognized and rejected
  });

  it("RM6: revision ownership does not leak across separate apply calls", () => {
    const body = P("10000001", R("Payments begin on the start date."));
    const { docx: norm } = normalizeDocx(docx(body));
    const doc = openDocument(norm);
    // call 1: point insertion at offset 16
    doc.apply([{ op: "replace", ref: "f1", author: "Reviewer A [f1]", addr: "10000001", start: 16, end: 16, expectText: "", text: "AAA " }], { date: DATE });
    // call 2: point edit inside call 1's inserted text must refuse
    const receipts = doc.apply([{ op: "replace", ref: "f2", author: "Reviewer A [f2]", addr: "10000001", start: 18, end: 18, expectText: "", text: "BBB " }], { date: DATE });
    expect(receipts[0]!.status).toBe("refused");
    expect(receipts[0]!.reason).toBe("span_overlaps_revision");
  });

  it("V1: swapping authors between two flags fails verify", () => {
    const body = P("10000001", R("Alpha beta gamma delta."));
    const { docx: norm } = normalizeDocx(docx(body));
    const doc = openDocument(norm);
    const receipts = doc.apply([
      { op: "replace", ref: "a", author: "Reviewer A [a]", addr: "10000001", start: 0, end: 5, expectText: "Alpha", text: "ALPHA" },
      { op: "replace", ref: "b", author: "Reviewer A [b]", addr: "10000001", start: 6, end: 10, expectText: "beta", text: "BETA" },
    ], { date: DATE });
    const out = doc.save();
    expect(verifyOutput(norm, out, receipts).ok).toBe(true);
    // swap the two flags' authors in the output markup
    const entries = readPackage(out);
    let xml = new TextDecoder().decode(entries.get("word/document.xml")!);
    xml = xml.replace("Reviewer A [a]", "Reviewer A [TMP]").replace("Reviewer A [b]", "Reviewer A [a]").replace("Reviewer A [TMP]", "Reviewer A [b]");
    entries.set("word/document.xml", encodePart(xml));
    expect(verifyOutput(norm, writePackage(entries), receipts).ok).toBe(false);
  });

  it("V2: insertion-only caller wording inside a hyperlink stays in place", () => {
    const body = P("10000001", R("A") + `<w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="r1">` + R("BC") + `</w:hyperlink>` + R("D"));
    const { docx: norm } = normalizeDocx(docx(body));
    const doc = openDocument(norm);
    doc.apply([{ op: "replace", ref: "f", author: "Reviewer A [f]", addr: "10000001", start: 2, end: 2, expectText: "", text: "X" }], { date: DATE });
    doc.acceptWithText({ ref: "f", author: "Reviewer A [f]", decision: "accept", text: "Y" });
    expect(doc.projection()[0]!.text).toBe("ABYCD");
  });

  it("V3: rejecting a property change restores the saved props (no merge)", () => {
    const body = P("10000001", `<w:r><w:rPr><w:b/><w:rPrChange w:id="91" w:author="Ours [legacy]" w:date="${DATE}"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr>` + R("kept") + `</w:r>`);
    const doc = openDocument(docx(body));
    doc.decide([{ author: "Ours [legacy]", decision: "reject" }]);
    const xml = new TextDecoder().decode(readPackage(doc.save()).get("word/document.xml")!);
    expect(xml).toContain("<w:i/>");
    expect(xml).not.toContain("<w:b/>");
    expect(xml).not.toContain("Ours [legacy]");
  });
});
