// Regression tests for tracked-change resolution, verification and
// wire-binding behavior. Builders construct minimal OOXML packages
// in-memory.
import { readFileSync } from "node:fs";
import { describe, it, expect } from "vitest";
import { encodePart, readPackage, writePackage, entryBytes } from "@stitchdocx/package-io";
import { normalizeDocx, openDocument, verifyOutput, type Op, type Receipt } from "./index.js";
import { buildSpanMap, checkSpan } from "./spanmap.js";
import { wordEdits } from "./worddiff.js";
import { paragraphText } from "./projection.js";
import { StyleSheet } from "./styles.js";

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

describe("tracked-change regression tests", () => {
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
    const xml = new TextDecoder().decode(entryBytes(readPackage(out1).get("word/document.xml")!));
    expect(xml).not.toMatch(/<w:ins [^>]*><w:ins /);
    // independent reject of f1 leaves f2's text
    const doc2 = openDocument(out1);
    doc2.decide([{ ref: "f1", changeIds: receipts[0]!.changeIds, decision: "reject" }]);
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
    const map = buildSpanMap(para.node, () => "accept", StyleSheet.empty);
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
    const map = buildSpanMap(para.node, () => "accept", StyleSheet.empty);
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

  it("CR/CRLF in text normalizes to LF at parse time and reads as a space, like Word", () => {
    const body = P("10000001", `<w:r><w:t>A\r\nB\rC</w:t></w:r>`);
    const { docx: norm } = normalizeDocx(docx(body));
    const doc = openDocument(norm);
    expect(doc.projection()[0]!.text).toBe("A B C");
  });

  it("w:t text reads the way Word reads it", () => {
    const textOf = (inner: string, preserve = false): string => {
      const attr = preserve ? ` xml:space="preserve"` : "";
      const doc = openDocument(normalizeDocx(docx(P("10000001", `<w:r><w:t${attr}>${inner}</w:t></w:r>`))).docx);
      return doc.projection()[0]!.text;
    };
    // edge whitespace is Word-dropped without xml:space="preserve", kept with it
    expect(textOf("  lead")).toBe("lead");
    expect(textOf("trail   ")).toBe("trail");
    expect(textOf("p   q  ", true)).toBe("p   q  ");
    // internal runs of spaces survive; newline and tab read as spaces
    expect(textOf("a   b")).toBe("a   b");
    expect(textOf("a\nb")).toBe("a b");
    expect(textOf("a\tb")).toBe("a b");
    expect(textOf("a\tb", true)).toBe("a b");
  });

  it("normalize reassigns duplicate paraIds deterministically and stays idempotent", () => {
    const body = P("10000001", R("one")) + P("10000001", R("two"));
    const first = normalizeDocx(docx(body));
    expect(first.assigned).toBe(1); // the duplicate reassigned
    const xml = new TextDecoder().decode(entryBytes(readPackage(first.docx).get("word/document.xml")!));
    const ids = [...xml.matchAll(/w14:paraId="([0-9A-F]+)"/g)].map((m) => m[1]!);
    expect(new Set(ids).size).toBe(ids.length);
    const second = normalizeDocx(first.docx);
    expect(Buffer.from(second.docx).equals(Buffer.from(first.docx))).toBe(true);
  });

  it("paragraphs never emits an empty addr, even pre-normalization (CLI path)", () => {
    const { docx: normalized } = normalizeDocx(docx(P(null, R("no id here"))));
    const doc = openDocument(normalized);
    for (const p of doc.projection()) expect(p.addr).toMatch(/^[0-9A-F]{8}$/);
  });

  it("verify rejects author tampering on a receipt revision", () => {
    const body = P("10000001", R("Hold harmless the indemnified party under all claims."));
    const { docx: norm } = normalizeDocx(docx(body));
    const doc = openDocument(norm);
    const receipts = doc.apply([{ op: "replace", ref: "f", author: "Reviewer A [f]", addr: "10000001", start: 0, end: 13, expectText: "Hold harmless", text: "Indemnify" }], { date: DATE });
    const out = doc.save();
    expect(verifyOutput(norm, out, receipts).ok).toBe(true);
    // tamper: change the revision's author
    const entries = readPackage(out);
    let xml = new TextDecoder().decode(entryBytes(entries.get("word/document.xml")!));
    xml = xml.replace('w:author="Reviewer A [f]"', 'w:author="Other Author"');
    entries.set("word/document.xml", encodePart(xml));
    const tampered = writePackage(entries);
    expect(verifyOutput(norm, tampered, receipts).ok).toBe(false);
  });

  it("acceptWith on an insertion-only edit inserts the caller wording", () => {
    const body = P("10000001", R("Payments begin on the start date."));
    const { docx: norm } = normalizeDocx(docx(body));
    const doc = openDocument(norm);
    const [r] = doc.apply([{ op: "replace", ref: "f", author: "Reviewer A [f]", addr: "10000001", start: 16, end: 16, expectText: "", text: "AAA " }], { date: DATE });
    doc.decide([{ ref: "f", changeIds: r!.changeIds, decision: "accept", text: "BBB ", expectText: "" }]);
    const text = doc.projection().find((p) => p.addr === "10000001")!.text;
    // offset 16 sits after the first "o" of "on": original text with BBB at 16
    expect(text).toBe("Payments begin oBBB n the start date.");
    const xml = new TextDecoder().decode(entryBytes(readPackage(doc.save()).get("word/document.xml")!));
    expect(xml).not.toContain("Reviewer A");
  });

  it("decide resolves move and property-change revisions by change id, leaving others untouched", () => {
    const body = P("10000001",
      `<w:moveFrom w:id="90" w:author="Reviewer A [m]" w:date="${DATE}">` + R("gone") + `</w:moveFrom>` +
      `<w:moveFrom w:id="92" w:author="Client Author" w:date="${DATE}">` + R("theirs") + `</w:moveFrom>` +
      `<w:r><w:rPr><w:b/><w:rPrChange w:id="91" w:author="Reviewer A [m]" w:date="${DATE}"><w:rPr/></w:rPrChange></w:rPr>` + R("kept") + `</w:r>`);
    const doc = openDocument(docx(body));
    const resolved = doc.decide([{ ref: "m", changeIds: ["90", "91"], decision: "reject" }]);
    expect(resolved[0]!.changeIds.sort()).toEqual(["90", "91"]);
    const xml = new TextDecoder().decode(entryBytes(readPackage(doc.save()).get("word/document.xml")!));
    expect(xml).not.toContain("Reviewer A [m]");
    expect(xml).toContain("Client Author"); // D3: revisions we never wrote survive
    expect(xml).toContain("gone"); // reject moveFrom restores the text
  });

  it("decisions resolve moveTo and pPrChange, and naming a resolved change fails", () => {
    const body = P("10000001",
      `<w:pPr><w:jc w:val="center"/><w:pPrChange w:id="94" w:author="Reviewer A [m]" w:date="${DATE}"><w:pPr><w:jc w:val="right"/></w:pPr></w:pPrChange></w:pPr>` +
      `<w:moveTo w:id="93" w:author="Reviewer A [m]" w:date="${DATE}">` + R("moved") + `</w:moveTo>` + R("stay"));
    const doc = openDocument(docx(body));
    const first = doc.decide([{ ref: "m", changeIds: ["93", "94"], decision: "reject" }]);
    expect(first[0]!.changeIds.sort()).toEqual(["93", "94"]);
    // the changes are gone, so a second decision naming them fails rather than resolving nothing
    expect(() => doc.decide([{ ref: "m", changeIds: ["93"], decision: "reject" }])).toThrow(/not a revision/);
    const xml = new TextDecoder().decode(entryBytes(readPackage(doc.save()).get("word/document.xml")!));
    expect(openDocument(doc.save()).projection()[0]!.text).toBe("stay"); // reject removes the moveTo content
    expect(xml).toContain('<w:jc w:val="right"/>'); // pPrChange reject restores the saved props
    expect(xml).not.toContain("Reviewer A [m]");
  });

  it("decide accepts changeIds-based decisions and returns receipts", () => {
    const body = P("10000001", R("The fee is due."));
    const { docx: norm } = normalizeDocx(docx(body));
    const doc = openDocument(norm);
    const receipts = doc.apply([{ op: "replace", ref: "f", author: "Reviewer A [f]", addr: "10000001", start: 4, end: 7, expectText: "fee", text: "charge" }], { date: DATE });
    const ids = receipts[0]!.changeIds;
    const doc2 = openDocument(doc.save());
    const decideReceipts = doc2.decide([{ ref: "f", changeIds: ids, decision: "reject" }]);
    expect(decideReceipts[0]!.changeIds.length).toBeGreaterThan(0);
    expect(doc2.projection()[0]!.text).toBe("The fee is due.");
  });

  it("engine-authored text with XML-invalid characters is sanitized", () => {
    const body = P("10000001", R("Rate card:"));
    const { docx: norm } = normalizeDocx(docx(body));
    const doc = openDocument(norm);
    const receipts = doc.apply([{ op: "replace", ref: "f", author: "Reviewer A [f]", addr: "10000001", start: 0, end: 9, expectText: "Rate card", text: "Rates\u0001\r\n2026" }], { date: DATE });
    expect(receipts[0]!.status).toBe("applied");
    const xml = new TextDecoder().decode(entryBytes(readPackage(doc.save()).get("word/document.xml")!));
    expect(xml).not.toContain("\u0001");
    expect(xml).not.toMatch(/\r/);
  });

  it("word-diff offsets: pure insertion keeps old offsets", () => {
    const edits = wordEdits("Sessions not used within the month are forfeited.", "Sessions not used within the month roll over once and are then forfeited.");
    expect(edits).toEqual([
      { insAt: 35, insText: "roll over once and " },
      { insAt: 39, insText: "then " },
    ]);
  });

  it("a point inside a field result refuses span_in_field", () => {
    const body = P("10000001",
      R("generated on ") +
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> DATE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
      R("October 3, 2026") +
      `<w:r><w:fldChar w:fldCharType="end"/></w:r>`);
    const { docx: norm } = normalizeDocx(docx(body));
    const doc = openDocument(norm);
    const para = doc.projectionRecords()[0]!;
    const map = buildSpanMap(para.node, () => "accept", StyleSheet.empty);
    const at = para.text.indexOf("October 3") + 2;
    const issue = checkSpan(map, at, at, []);
    expect(issue?.reason).toBe("span_in_field");
  });

  it("a point inside an inline content control refuses span_in_content_control", () => {
    const body = P("10000001",
      R("Effective date: ") +
      `<w:sdt><w:sdtPr><w:alias w:val="d"/></w:sdtPr><w:sdtContent>` + R("October 1, 2026") + `</w:sdtContent></w:sdt>`);
    const { docx: norm } = normalizeDocx(docx(body));
    const doc = openDocument(norm);
    const para = doc.projectionRecords()[0]!;
    const map = buildSpanMap(para.node, () => "accept", StyleSheet.empty);
    const at = para.text.indexOf("October 1") + 2;
    const issue = checkSpan(map, at, at, []);
    expect(issue?.reason).toBe("span_in_content_control");
  });

  it("a point at a pre-existing client revision refuses span_overlaps_revision even with a matching author", () => {
    const body = P("10000001",
      R("Payment terms: ") +
      `<w:ins w:id="50" w:author="Reviewer A [clientlookalike]" w:date="${DATE}">` + R("net sixty") + `</w:ins>` + R("."));
    const { docx: norm } = normalizeDocx(docx(body));
    const doc = openDocument(norm);
    const para = doc.projectionRecords()[0]!;
    const map = buildSpanMap(para.node, () => "accept", StyleSheet.empty);
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

  it("attribute values normalize literal CR/CRLF/TAB like the spec parser (author parity)", () => {
    const body = `<w:p w14:paraId="10000001">` +
      `<w:ins w:id="7" w:author="Reviewer
A [f]" w:date="${DATE}">` + R("wrapped author text") + `</w:ins>` +
      R(" tail") + `</w:p>`;
    const doc = openDocument(docx(body));
    const rec = doc.projectionRecords()[0]!;
    const text = paragraphText(rec.node, (rev) => (rev.author === "Reviewer A [f]" ? "reject" : "accept"), StyleSheet.empty);
    expect(text).toBe(" tail"); // the wrapped-author insertion is recognized and rejected
  });

  it("revision ownership does not leak across separate apply calls", () => {
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

  it("swapping authors between two flags fails verify", () => {
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
    let xml = new TextDecoder().decode(entryBytes(entries.get("word/document.xml")!));
    xml = xml.replace("Reviewer A [a]", "Reviewer A [TMP]").replace("Reviewer A [b]", "Reviewer A [a]").replace("Reviewer A [TMP]", "Reviewer A [b]");
    entries.set("word/document.xml", encodePart(xml));
    expect(verifyOutput(norm, writePackage(entries), receipts).ok).toBe(false);
  });

  it("insertion-only caller wording inside a hyperlink stays in place", () => {
    const body = P("10000001", R("A") + `<w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="r1">` + R("BC") + `</w:hyperlink>` + R("D"));
    const { docx: norm } = normalizeDocx(docx(body));
    const doc = openDocument(norm);
    const [r] = doc.apply([{ op: "replace", ref: "f", author: "Reviewer A [f]", addr: "10000001", start: 2, end: 2, expectText: "", text: "X" }], { date: DATE });
    doc.decide([{ ref: "f", changeIds: r!.changeIds, decision: "accept", text: "Y", expectText: "" }]);
    expect(doc.projection()[0]!.text).toBe("ABYCD");
  });

  describe("word granularity around a hyperlink", () => {
    const LINK = (text: string) => `<w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="r1">${R(text)}</w:hyperlink>`;
    const word = (expectText: string, text: string, granularity: "word" | "span" = "word"): Op => ({
      op: "replace", ref: "f", author: "Reviewer A [f]", addr: "10000001", start: 0, end: expectText.length, expectText, text, granularity,
    });
    const xmlOf = (out: Uint8Array) => new TextDecoder().decode(entryBytes(readPackage(out).get("word/document.xml")!));

    it("rewrites the text around an unchanged link", () => {
      const before = "Profile site.me • a@b.com";
      const body = P("10000001", R("Profile ") + LINK("site.me") + R(" • a@b.com"));
      const { docx: norm } = normalizeDocx(docx(body));
      const doc = openDocument(norm);
      const receipts = doc.apply([word(before, "Portfolio site.me • c@d.com")], { date: DATE });
      expect(receipts[0]!.status).toBe("applied");
      expect(doc.projection()[0]!.text).toBe("Portfolio site.me • c@d.com");
      const out = doc.save();
      expect(xmlOf(out)).toMatch(/<w:hyperlink [^>]*><w:r><w:t xml:space="preserve">site\.me<\/w:t><\/w:r><\/w:hyperlink>/);
      expect(verifyOutput(norm, out, receipts).ok).toBe(true);
    });

    it("replaces a whole span that reaches into a link, deleting the link text in place and adding the new text outside it", () => {
      const before = "Profile site.me • a@b.com";
      const after = "Portfolio site.me • c@d.com";
      const body = P("10000001", R("Profile ", "<w:rPr><w:b/></w:rPr>") + LINK("site.me") + R(" • a@b.com"));
      const { docx: norm } = normalizeDocx(docx(body));
      const doc = openDocument(norm);
      const receipts = doc.apply([word(before, after, "span")], { date: DATE });
      expect(receipts[0]!.status).toBe("applied");
      expect(receipts[0]!.after).toBe(after);
      const out = doc.save();
      const xml = xmlOf(out);
      expect(xml).toMatch(/<w:hyperlink [^>]*><w:del [^>]*><w:r><w:delText xml:space="preserve">site\.me<\/w:delText><\/w:r><\/w:del><\/w:hyperlink>/);
      expect(xml).toMatch(/<\/w:del><w:ins [^>]*><w:r><w:rPr><w:b\/><\/w:rPr><w:t xml:space="preserve">Portfolio site\.me • c@d\.com<\/w:t><\/w:r><\/w:ins><\/w:p>/);
      expect(verifyOutput(norm, out, receipts).ok).toBe(true);
      doc.decide([{ ref: "f", changeIds: receipts[0]!.changeIds, decision: "reject" }]);
      expect(doc.projection()[0]!.text).toBe(before);
    });

    it("rewrites a contact line whose link text changes, as a member's header edit needed", () => {
      const before = "www.linkedin.com/in/sample• sample@example.com • (555) 010-0199 ";
      const after = "Austin, TX | sample@example.com | (555) 010-0199 | linkedin.com/in/sample";
      const body = P("10000001", LINK("www.linkedin.com/in/sample") + R("• sample@example.com • (555) 010-0199 "));
      const { docx: norm } = normalizeDocx(docx(body));
      const doc = openDocument(norm);
      const receipts = doc.apply([word(before, after)], { date: DATE });
      expect(receipts[0]!.status).toBe("applied");
      expect(receipts[0]!.after).toBe(after);
      expect(verifyOutput(norm, doc.save(), receipts).ok).toBe(true);
      doc.decide([{ ref: "f", changeIds: receipts[0]!.changeIds, decision: "reject" }]);
      expect(doc.projection()[0]!.text).toBe(before);
    });

    it("deletes a whole line that holds a link", () => {
      const before = "site.me • a@b.com";
      const body = P("10000001", LINK("site.me") + R(" • a@b.com"));
      const { docx: norm } = normalizeDocx(docx(body));
      const doc = openDocument(norm);
      const receipts = doc.apply([{ op: "delete", ref: "f", author: "Reviewer A [f]", addr: "10000001", start: 0, end: before.length, expectText: before }], { date: DATE });
      expect(receipts[0]!.status).toBe("applied");
      expect(receipts[0]!.changeIds).toHaveLength(2);
      expect(doc.projection()[0]!.text).toBe("");
      expect(verifyOutput(norm, doc.save(), receipts).ok).toBe(true);
      doc.decide([{ ref: "f", changeIds: receipts[0]!.changeIds, decision: "reject" }]);
      expect(doc.projection()[0]!.text).toBe(before);
    });

    it("still refuses a span that reaches into a field result inside a link", () => {
      const NS = `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`;
      const before = "See the portal";
      const body = P("10000001", R("See ") + `<w:hyperlink ${NS} r:id="r1"><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> REF x </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${R("the portal")}<w:r><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink>`);
      const { receipts } = runOn(body, [word(before, "Read the docs", "span")]);
      expect(receipts[0]!.reason).toBe("span_crosses_structure");
    });

    it("removes link text and adds text before the link without moving it into the link", () => {
      const before = "www.site.me/in/x• a@b.com • (555) 010-0100 ";
      const after = "Austin, TX • a@b.com • (555) 010-0100 • site.me/in/x";
      const body = P("10000001", LINK("www.site.me/in/x") + R("• a@b.com • (555) 010-0100 "));
      const { docx: norm } = normalizeDocx(docx(body));
      const doc = openDocument(norm);
      const receipts = doc.apply([word(before, after)], { date: DATE });
      expect(receipts[0]!.status).toBe("applied");
      expect(receipts[0]!.before).toBe(before);
      expect(receipts[0]!.after).toBe(after);
      const out = doc.save();
      expect(xmlOf(out).match(/<w:hyperlink [\s\S]*?<\/w:hyperlink>/)![0]).not.toContain("<w:ins ");
      expect(verifyOutput(norm, out, receipts).ok).toBe(true);
    });

    it("places text inserted at the link's leading edge before the link", () => {
      const before = "site.me today";
      const body = P("10000001", LINK("site.me") + R(" today"));
      const { docx: norm } = normalizeDocx(docx(body));
      const doc = openDocument(norm);
      const receipts = doc.apply([word(before, "Visit site.me today")], { date: DATE });
      expect(receipts[0]!.status).toBe("applied");
      const out = doc.save();
      expect(xmlOf(out)).toMatch(/<w:ins [^>]*><w:r><w:t xml:space="preserve">Visit <\/w:t><\/w:r><\/w:ins><w:hyperlink /);
      expect(verifyOutput(norm, out, receipts).ok).toBe(true);
    });

    it("writes a rewritten line as one replacement", () => {
      const before = "Storyteller | Technology Enthusiast | Lifelong Learner";
      const { docx: norm } = normalizeDocx(docx(P("10000001", R(before))));
      const doc = openDocument(norm);
      const receipts = doc.apply([word(before, "Solutions Engineer turning technical depth into clear buyer value")], { date: DATE });
      expect(receipts[0]!.status).toBe("applied");
      expect(receipts[0]!.changeIds).toHaveLength(2);
      expect(verifyOutput(norm, doc.save(), receipts).ok).toBe(true);
    });

    it("falls back to word edits when one replacement would span two links", () => {
      const LINK2 = (text: string) => `<w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="r2">${R(text)}</w:hyperlink>`;
      const before = "LinkedIn | GitHub";
      const body = P("10000001", LINK("LinkedIn") + R(" | ") + LINK2("GitHub"));
      const { docx: norm } = normalizeDocx(docx(body));
      const doc = openDocument(norm);
      const receipts = doc.apply([word(before, "Profile | Code")], { date: DATE });
      expect(receipts[0]!.status).toBe("applied");
      expect(doc.projection()[0]!.text).toBe("Profile | Code");
      const out = doc.save();
      expect(xmlOf(out).match(/<w:hyperlink /g)).toHaveLength(2);
      expect(verifyOutput(norm, out, receipts).ok).toBe(true);
    });

    it("gives inserted words the formatting of the text before them, as typing in Word does", () => {
      const BOLD = "<w:rPr><w:b/></w:rPr>";
      const ITALIC = "<w:rPr><w:i/></w:rPr>";
      const dated = runOn(P("10000001", R("Acme | September 2023 - Present", BOLD)), [word("Acme | September 2023 - Present", "Acme | September 2023 - May 2025")]);
      expect(xmlOf(dated.out)).toMatch(/<w:ins [^>]*><w:r><w:rPr><w:b\/><\/w:rPr><w:t xml:space="preserve">May 2025<\/w:t>/);
      const tagline = "Storyteller | Technology Enthusiast | Lifelong Learner";
      const rewritten = runOn(P("10000001", R(tagline, ITALIC)), [word(tagline, "Solutions Engineer turning technical depth into clear buyer value")]);
      expect(xmlOf(rewritten.out)).toMatch(/<w:ins [^>]*><w:r><w:rPr><w:i\/><\/w:rPr><w:t xml:space="preserve">Solutions Engineer/);
    });

    it("gives replacement words the formatting of the words they replace, not of their neighbours", () => {
      const before = "A Present B";
      const body = P("10000001", R("A ", "<w:rPr><w:b/></w:rPr>") + R("Present", "<w:rPr><w:i/></w:rPr>") + R(" B", "<w:rPr><w:b/></w:rPr>"));
      const { docx: norm } = normalizeDocx(docx(body));
      const doc = openDocument(norm);
      const receipts = doc.apply([word(before, "A May 2025 B")], { date: DATE });
      const out = doc.save();
      expect(xmlOf(out)).toMatch(/<w:ins [^>]*><w:r><w:rPr><w:i\/><\/w:rPr><w:t xml:space="preserve">May 2025/);
      expect(verifyOutput(norm, out, receipts).ok).toBe(true);
      doc.decide([{ ref: "f", changeIds: receipts[0]!.changeIds, decision: "reject" }]);
      expect(doc.projection()[0]!.text).toBe(before);
    });

    it("formats a replacement that starts inside a run like that first replaced run", () => {
      const before = "A Present B";
      const BOLD = "<w:rPr><w:b/></w:rPr>";
      const body = P("10000001", R("A Pre", BOLD) + R("sent", "<w:rPr><w:i/></w:rPr>") + R(" B", BOLD));
      const { docx: norm } = normalizeDocx(docx(body));
      const doc = openDocument(norm);
      const receipts = doc.apply([word(before, "A May 2025 B")], { date: DATE });
      const out = doc.save();
      expect(xmlOf(out)).toMatch(/<w:ins [^>]*><w:r><w:rPr><w:b\/><\/w:rPr><w:t xml:space="preserve">May 2025/);
      expect(verifyOutput(norm, out, receipts).ok).toBe(true);
      doc.decide([{ ref: "f", changeIds: receipts[0]!.changeIds, decision: "reject" }]);
      expect(doc.projection()[0]!.text).toBe(before);
    });

    it("formats text added before a link like the plain text, and text added inside a link like the link", () => {
      const STYLED = (text: string) => `<w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="r1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>${text}</w:t></w:r></w:hyperlink>`;
      const body = P("10000001", STYLED("site.me") + R(" today", "<w:rPr><w:b/></w:rPr>"));
      const before = runOn(body, [word("site.me today", "Visit site.me today")]);
      expect(xmlOf(before.out)).toMatch(/<w:ins [^>]*><w:r><w:rPr><w:b\/><\/w:rPr><w:t xml:space="preserve">Visit <\/w:t><\/w:r><\/w:ins><w:hyperlink /);
      const inside = runOn(body, [word("site.me today", "site.new.me today")]);
      expect(inside.receipts[0]!.after).toBe("site.new.me today");
      expect(xmlOf(inside.out)).toMatch(/<w:hyperlink [^>]*>[^]*<w:ins [^>]*><w:r><w:rPr><w:rStyle w:val="Hyperlink"\/><\/w:rPr><w:t xml:space="preserve">[^<]*new[^<]*<\/w:t>[^]*<\/w:hyperlink>/);
    });

    it("edits a changed token that straddles the link boundary on both sides of it", () => {
      const before = "site.me• a@b.com";
      const { docx: norm } = normalizeDocx(docx(P("10000001", LINK("site.me") + R("• a@b.com"))));
      const doc = openDocument(norm);
      const receipts = doc.apply([word(before, "site.ca/ a@b.com")], { date: DATE });
      expect(receipts[0]!.status).toBe("applied");
      expect(doc.projection()[0]!.text).toBe("site.ca/ a@b.com");
      expect(verifyOutput(norm, doc.save(), receipts).ok).toBe(true);
    });
  });

  describe("text inserted at a link's leading edge", () => {
    const NS = `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`;
    const EMPTY = `<w:r><w:t/></w:r>`;
    const cases: Array<[string, string, RegExp]> = [
      ["a hyperlink that opens with an empty run", `<w:hyperlink ${NS} r:id="r1">${EMPTY}${R("site.me")}</w:hyperlink>`, /Visit <\/w:t><\/w:r><\/w:ins><w:hyperlink /],
      ["a smartTag that opens with an empty run", `<w:smartTag w:uri="u" w:element="e">${EMPTY}${R("site.me")}</w:smartTag>`, /Visit <\/w:t><\/w:r><\/w:ins><w:smartTag /],
      ["a smartTag nested at the start of a hyperlink", `<w:hyperlink ${NS} r:id="r1"><w:smartTag w:uri="u" w:element="e">${EMPTY}${R("site")}</w:smartTag>${R(".me")}</w:hyperlink>`, /Visit <\/w:t><\/w:r><\/w:ins><w:hyperlink /],
    ];
    for (const [name, structure, placed] of cases) {
      it(`places it before ${name}`, () => {
        const body = P("10000001", structure + R(" today"));
        const { docx: norm } = normalizeDocx(docx(body));
        const doc = openDocument(norm);
        const receipts = doc.apply([{ op: "replace", ref: "f", author: "Reviewer A [f]", addr: "10000001", start: 0, end: 0, expectText: "", text: "Visit " }], { date: DATE });
        expect(receipts[0]!.status).toBe("applied");
        const out = doc.save();
        expect(new TextDecoder().decode(entryBytes(readPackage(out).get("word/document.xml")!))).toMatch(placed);
        expect(verifyOutput(norm, out, receipts).ok).toBe(true);
      });
    }
  });

  describe("insertParagraphAfter with formatFrom", () => {
    const LIST = `<w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:rPr><w:b/></w:rPr></w:pPr>`;
    const insertedPPr = (out: Uint8Array, addr: string) => {
      const xml = new TextDecoder().decode(entryBytes(readPackage(out).get("word/document.xml")!));
      return xml.match(new RegExp(`<w:p w14:paraId="${addr}"[^>]*>(<w:pPr>[\\s\\S]*?</w:pPr>)?(?=<w:(?:ins|r|del)[ >])`))![1] ?? "";
    };
    for (const [form, body] of [
      ["before another paragraph", P("10000001", LIST + R("First")) + P("10000002", R("Next"))],
      ["at the end of the body", P("10000001", LIST + R("First"))],
    ] as const) {
      it(`gives the new paragraph the anchor's properties ${form}`, () => {
        const { docx: norm } = normalizeDocx(docx(body));
        const doc = openDocument(norm);
        const receipts = doc.apply([{ op: "insertParagraphAfter", ref: "f", author: "Reviewer A [f]", addr: "10000001", text: "Second", formatFrom: "10000001" }], { date: DATE });
        expect(receipts[0]!.status).toBe("applied");
        const out = doc.save();
        const pPr = insertedPPr(out, receipts[0]!.insertedAddr!);
        expect(pPr.replace(/<w:pPrChange[\s\S]*<\/w:pPrChange>/, "")).toContain(`<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>`);
        expect(pPr.replace(/<w:pPrChange[\s\S]*<\/w:pPrChange>/, "")).not.toContain("<w:b/>");
        expect(verifyOutput(norm, out, receipts).ok).toBe(true);
        const rejected = openDocument(out);
        rejected.decide([{ ref: "f", changeIds: receipts[0]!.changeIds, decision: "reject" }]);
        expect(rejected.projection().map((p) => p.text)).toEqual(body.includes("Next") ? ["First", "Next"] : ["First"]);
      });
    }

    for (const [form, after] of [["before another paragraph", P("10000002", R("Next"))], ["at the end of the body", ""]] as const) {
      it(`never repeats a revision id copied from the source's properties ${form}`, () => {
        const NUMBERED = `<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/><w:numberingChange w:id="7" w:author="Client" w:date="${DATE}" w:original="%1."/></w:numPr></w:pPr>`;
        const body = P("10000001", NUMBERED + R("First")) + after;
        const { docx: norm } = normalizeDocx(docx(body));
        const doc = openDocument(norm);
        const receipts = doc.apply([{ op: "insertParagraphAfter", ref: "f", author: "Reviewer A [f]", addr: "10000001", text: "Second", formatFrom: "10000001" }], { date: DATE });
        expect(receipts[0]!.status).toBe("applied");
        const out = doc.save();
        const xml = new TextDecoder().decode(entryBytes(readPackage(out).get("word/document.xml")!));
        const ids = [...xml.matchAll(/<w:[A-Za-z]+ w:id="(\d+)"/g)].map((m) => m[1]);
        expect(new Set(ids).size).toBe(ids.length);
        expect(insertedPPr(out, receipts[0]!.insertedAddr!)).toContain("<w:numId w:val=\"1\"/>");
        expect(verifyOutput(norm, out, receipts).ok).toBe(true);
        const rejected = openDocument(out);
        rejected.decide([{ ref: "f", changeIds: receipts[0]!.changeIds, decision: "reject" }]);
        expect(rejected.projection().map((p) => p.text)).toEqual(after ? ["First", "Next"] : ["First"]);
      });
    }

    it("takes properties and character formatting from another paragraph than the anchor", () => {
      const HEAD = `<w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Role</w:t></w:r>`;
      const body = P("10000001", HEAD) + P("10000002", LIST + R("Bullet")) + P("10000003", R("Next"));
      const { docx: norm } = normalizeDocx(docx(body));
      const doc = openDocument(norm);
      const receipts = doc.apply([{ op: "insertParagraphAfter", ref: "f", author: "Reviewer A [f]", addr: "10000002", text: "New role", formatFrom: "10000001" }], { date: DATE });
      expect(receipts[0]!.status).toBe("applied");
      const out = doc.save();
      const pPr = insertedPPr(out, receipts[0]!.insertedAddr!);
      expect(pPr).toContain(`<w:pStyle w:val="Heading2"/>`);
      expect(pPr).not.toContain("<w:numPr>");
      const xml = new TextDecoder().decode(entryBytes(readPackage(out).get("word/document.xml")!));
      expect(xml).toMatch(/<w:ins [^>]*><w:r><w:rPr><w:b\/><\/w:rPr><w:t xml:space="preserve">New role/);
      expect(verifyOutput(norm, out, receipts).ok).toBe(true);
    });

    const NS = `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`;
    const BOLD = `<w:r><w:rPr><w:b/></w:rPr><w:t>Role</w:t></w:r>`;
    for (const [name, source, copied] of [
      ["a smartTag", `<w:smartTag w:uri="u" w:element="e">${BOLD}</w:smartTag>`, true],
      ["a customXml wrapper", `<w:customXml w:element="e">${BOLD}</w:customXml>`, true],
      ["a hyperlink, whose link formatting stays with the link", `<w:hyperlink ${NS} r:id="r1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>site.me</w:t></w:r></w:hyperlink>${BOLD}`, true],
      ["only a hyperlink", `<w:hyperlink ${NS} r:id="r1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>site.me</w:t></w:r></w:hyperlink>`, false],
    ] as const) {
      it(`takes character formatting from the source's first plain run inside ${name}`, () => {
        const body = P("10000001", source) + P("10000002", R("Bullet")) + P("10000003", R("Next"));
        const { docx: norm } = normalizeDocx(docx(body));
        const doc = openDocument(norm);
        const receipts = doc.apply([{ op: "insertParagraphAfter", ref: "f", author: "Reviewer A [f]", addr: "10000002", text: "New role", formatFrom: "10000001" }], { date: DATE });
        expect(receipts[0]!.status).toBe("applied");
        const out = doc.save();
        const xml = new TextDecoder().decode(entryBytes(readPackage(out).get("word/document.xml")!));
        const inserted = xml.match(/<w:ins [^>]*><w:r>(<w:rPr>[\s\S]*?<\/w:rPr>)?<w:t xml:space="preserve">New role/)!;
        expect(inserted[1] ?? "").toBe(copied ? "<w:rPr><w:b/></w:rPr>" : "");
        expect(xml).not.toMatch(/<w:ins [^>]*><w:r><w:rPr><w:rStyle w:val="Hyperlink"\/>/);
        expect(verifyOutput(norm, out, receipts).ok).toBe(true);
      });
    }

    it("refuses a formatFrom that names a merged record, as an op address does", () => {
      const merged = P("10000001", `<w:pPr><w:rPr><w:del w:id="9" w:author="Client" w:date="${DATE}"/></w:rPr></w:pPr>` + R("Gone mark")) + P("10000002", R("Joined"));
      const doc = openDocument(normalizeDocx(docx(merged + P("10000003", R("Anchor")))).docx);
      const listed = doc.projection().find((p) => !p.editable);
      expect(listed).toBeDefined();
      const receipts = doc.apply([{ op: "insertParagraphAfter", ref: "f", author: "Reviewer A [f]", addr: "10000003", text: "New", formatFrom: listed!.addr }], { date: DATE });
      expect(receipts[0]).toMatchObject({ status: "refused", reason: "paragraph_not_editable" });
    });

    it("refuses a formatFrom that names no paragraph", () => {
      const body = P("10000001", LIST + R("First")) + P("10000002", R("Next"));
      const { receipts } = runOn(body, [{ op: "insertParagraphAfter", ref: "f", author: "Reviewer A [f]", addr: "10000001", text: "Second", formatFrom: "7FFFFFFE" }]);
      expect(receipts[0]).toMatchObject({ status: "refused", reason: "addr_not_found" });
    });

    it("leaves the new paragraph's properties empty without it", () => {
      const body = P("10000001", LIST + R("First")) + P("10000002", R("Next"));
      const { receipts, out } = runOn(body, [{ op: "insertParagraphAfter", ref: "f", author: "Reviewer A [f]", addr: "10000001", text: "Second" }]);
      expect(receipts[0]!.status).toBe("applied");
      expect(insertedPPr(out, receipts[0]!.insertedAddr!)).not.toContain("<w:numPr>");
    });
  });

  it("lists each revision with its paragraph's address", () => {
    const body = P("10000001", R("Alpha beta")) + P("10000002", R("Gamma"));
    const { docx: norm } = normalizeDocx(docx(body));
    const doc = openDocument(norm);
    const receipts = doc.apply([{ op: "replace", ref: "f", author: "Reviewer A [f]", addr: "10000002", start: 0, end: 5, expectText: "Gamma", text: "Delta" }], { date: DATE });
    const listed = openDocument(doc.save()).revisions();
    expect(listed.length).toBeGreaterThan(0);
    expect(listed.every((r) => r.addr === "10000002")).toBe(true);
    expect(listed.map((r) => r.id).sort()).toEqual([...receipts[0]!.changeIds].sort());
  });

  it("rejecting a property change restores the saved props (no merge)", () => {
    const body = P("10000001", `<w:r><w:rPr><w:b/><w:rPrChange w:id="91" w:author="Reviewer A [m]" w:date="${DATE}"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr>` + R("kept") + `</w:r>`);
    const doc = openDocument(docx(body));
    doc.decide([{ ref: "p", changeIds: ["91"], decision: "reject" }]);
    const xml = new TextDecoder().decode(entryBytes(readPackage(doc.save()).get("word/document.xml")!));
    expect(xml).toContain("<w:i/>");
    expect(xml).not.toContain("<w:b/>");
    expect(xml).not.toContain("Reviewer A [m]");
  });
});
