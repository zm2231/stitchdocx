// Regression tests for Word-parity behaviors over real-world document
// structures: revision ids across run splits, removed paragraph marks,
// plain-text content controls, Word's whitespace reading, text boxes, and
// editing restrictions. Builders construct minimal OOXML packages in-memory.
import { describe, expect, it } from "vitest";
import { decodePart, encodePart, readPackage, writePackage } from "@stitchdocx/package-io";
import { normalizeDocx, openDocument, verifyOutput, type Op } from "./index.js";

const DATE = "2026-10-03T00:00:00Z";
const AUTHOR = "Reviewer A";

function entriesOf(bodyChildrenXml: string, settingsXml?: string): Map<string, Uint8Array> {
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"><w:body>${bodyChildrenXml}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;
  const entries = new Map<string, Uint8Array>([
    ["[Content_Types].xml", encodePart(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`)],
    ["word/document.xml", encodePart(xml)],
  ]);
  if (settingsXml) entries.set("word/settings.xml", encodePart(settingsXml));
  return entries;
}

const docx = (body: string): Uint8Array => writePackage(entriesOf(body));

const P = (paraId: string, inner: string, pPr = ""): string => `<w:p w14:paraId="${paraId}">${pPr ? `<w:pPr>${pPr}</w:pPr>` : ""}${inner}</w:p>`;
const R = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const PLAIN = (text: string): string => `<w:r><w:t>${text}</w:t></w:r>`;
const DEL_MARK = (id: number): string => `<w:rPr><w:del w:id="${id}" w:author="Client Author" w:date="${DATE}"/></w:rPr>`;
const runApply = (bodyXml: string, ops: Op[]) => {
  const doc = openDocument(normalizeDocx(docx(bodyXml)).docx);
  const receipts = doc.apply(ops, { date: DATE });
  return { doc, receipts, out: doc.save() };
};
const xmlOf = (bytes: Uint8Array): string => new TextDecoder().decode(readPackage(bytes).get("word/document.xml")!);
const textsOf = (bytes: Uint8Array, view: "accept" | "reject" = "accept", author = "Client Author"): string[] =>
  openDocument(bytes).projectionRecords((r) => (r.author === author ? view : "accept")).map((p) => p.text);

describe("run splits keep revision ids unique", () => {
  it("a split across a run with a client formatting change re-ids the copies", () => {
    const body = P("10000001",
      `<w:r><w:rPr><w:i/><w:rPrChange w:id="17" w:author="Client Author" w:date="${DATE}"><w:rPr/></w:rPrChange></w:rPr>` + `<w:t xml:space="preserve">abcdefgh</w:t></w:r>`);
    const { receipts, out } = runOn(body);
    expect(receipts[0]!.status).toBe("applied");
    const xml = xmlOf(out);
    const changes = [...xml.matchAll(/<w:rPrChange w:id="(\d+)" w:author="([^"]*)"/g)].map((m) => ({ id: m[1]!, author: m[2]! }));
    // the original id survives exactly once; every copy is unique and keeps the author
    expect(changes.filter((c) => c.id === "17")).toHaveLength(1);
    expect(new Set(changes.map((c) => c.id)).size).toBe(changes.length);
    expect(changes.length).toBeGreaterThanOrEqual(2);
    expect(changes.every((c) => c.author === "Client Author")).toBe(true);
    // a NEW run never carries the client's record: the inserted run has no rPrChange
    const ins = xml.match(/<w:ins [^>]*>[\s\S]*?<\/w:ins>/g) ?? [];
    expect(ins.some((frag) => frag.includes("XY"))).toBe(true);
    expect(ins.some((frag) => frag.includes("w:rPrChange"))).toBe(false);
    expect(verifyOutput(normalizeDocx(docx(body)).docx, out, receipts).ok).toBe(true);
  });

  function runOn(body: string) {
    return runApply(body, [
      { op: "replace", ref: "f", author: AUTHOR, addr: "10000001", start: 2, end: 6, expectText: "cdef", text: "XY" },
    ]);
  }
});

describe("removed paragraph marks merge the way Word merges", () => {
  const delPara = (id: string, text: string): string => P(id, R(text), DEL_MARK(90));

  it("a removed mark directly before a table joins the first cell's first paragraph", () => {
    const body = delPara("10000001", "Deleted heading") + `<w:tbl><w:tr><w:tc>${P("10000002", R("Cell text."))}</w:tc></w:tr></w:tbl>`;
    const doc = openDocument(docx(body));
    const accepted = doc.projectionRecords();
    expect(accepted).toHaveLength(1);
    expect(accepted[0]!.text).toBe("Deleted headingCell text.");
    expect(accepted[0]!.container).toBe("cell");
    expect(textsOf(docx(body), "reject")).toEqual(["Deleted heading", "Cell text."]);
  });

  it("at the end of a cell an emptied deleted paragraph disappears; anything else stays", () => {
    const cell = (tail: string): string => `<w:tbl><w:tr><w:tc>${P("10000002", R("Cell text."))}${tail}</w:tc></w:tr></w:tbl>`;
    const emptyDel = openDocument(docx(cell(delPara("10000003", ""))));
    expect(emptyDel.projectionRecords().map((p) => p.text)).toEqual(["Cell text."]);
    const withText = openDocument(docx(cell(delPara("10000003", "leftover"))));
    expect(withText.projectionRecords().map((p) => p.text)).toEqual(["Cell text.", "leftover"]);
    // an inserted mark under reject stays even when empty
    const insTail = P("10000003", "", `<w:rPr><w:ins w:id="91" w:author="Client Author" w:date="${DATE}"/></w:rPr>`);
    expect(textsOf(docx(cell(insTail)), "reject")).toEqual(["Cell text.", ""]);
  });

  it("at the end of the body or of a content control the paragraph stays, empty or not", () => {
    const bodyDoc = openDocument(docx(P("10000001", R("Body.")) + delPara("10000002", "")));
    expect(bodyDoc.projectionRecords().map((p) => p.text)).toEqual(["Body.", ""]);
    const sdt = `<w:sdt><w:sdtPr><w:alias w:val="c"/></w:sdtPr><w:sdtContent>${P("10000004", R("In control."))}${delPara("10000005", "")}</w:sdtContent></w:sdt>`;
    const sdtDoc = openDocument(docx(P("10000001", R("Body.")) + sdt));
    expect(sdtDoc.projectionRecords().map((p) => p.text)).toEqual(["Body.", "In control.", ""]);
  });
});

describe("plain-text content controls hold one paragraph", () => {
  const control = (inner: string): string =>
    `<w:sdt><w:sdtPr><w:text/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

  it("inserting a paragraph into one is refused; editing its text still applies", () => {
    const body = P("10000001", R("Before. ")) + control(P("10000002", R("Control text."))) + P("10000003", R("After."));
    const ins = runApply(body, [{ op: "insertParagraphAfter", ref: "i", author: AUTHOR, addr: "10000002", text: "Second paragraph." }]);
    expect(ins.receipts[0]!.status).toBe("refused");
    expect(ins.receipts[0]!.reason).toBe("insert_in_plain_text_control");
    expect(verifyOutput(normalizeDocx(docx(body)).docx, ins.out, ins.receipts).ok).toBe(true);

    const rep = runApply(body, [
      { op: "replace", ref: "r", author: AUTHOR, addr: "10000002", start: 8, end: 12, expectText: "text", text: "value" },
    ]);
    expect(rep.receipts[0]!.status).toBe("applied");
    expect(verifyOutput(normalizeDocx(docx(body)).docx, rep.out, rep.receipts).ok).toBe(true);
  });
});

describe("text boxes are neither listed nor edited", () => {
  const BOX = (text: string): string =>
    `<w:mc:AlternateContent><w:mc:Choice Requires="wps"><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml"><v:textbox><w:txbxContent>${P("20000001", R(text))}</w:txbxContent></v:textbox></v:shape></w:pict></w:mc:Choice><w:mc:Fallback><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml"><v:textbox><w:txbxContent>${P("20000002", R(text))}</w:txbxContent></v:textbox></v:shape></w:pict></w:mc:Fallback></w:mc:AlternateContent>`;

  it("the anchoring paragraph lists only its own text, once", () => {
    const body = P("10000001", R("Visible ") + BOX("Box words") + R("tail"));
    const doc = openDocument(docx(body));
    expect(doc.projection().map((p) => p.text)).toEqual(["Visible tail"]);
  });

  it("an edit in the anchoring paragraph leaves every text box byte-identical", () => {
    const body = P("10000001", R("Visible ") + BOX("Box words") + R("tail"));
    const { receipts, out } = runApply(body, [
      { op: "replace", ref: "f", author: AUTHOR, addr: "10000001", start: 0, end: 7, expectText: "Visible", text: "Edited" },
    ]);
    expect(receipts[0]!.status).toBe("applied");
    expect(verifyOutput(normalizeDocx(docx(body)).docx, out, receipts).ok).toBe(true);
    const before = xmlOf(normalizeDocx(docx(body)).docx).match(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/g) ?? [];
    const after = xmlOf(out).match(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/g) ?? [];
    expect(after).toEqual(before);
  });
});

describe("splits keep Word's reading of untouched bytes stable", () => {
  it("an internal-space element splits without changing what Word drops or shows", () => {
    const body = P("10000001", PLAIN("a   b"));
    const { receipts, out } = runApply(body, [
      { op: "replace", ref: "f", author: AUTHOR, addr: "10000001", start: 0, end: 1, expectText: "a", text: "X" },
    ]);
    expect(receipts[0]!.status).toBe("applied");
    expect(textsOf(out, "accept", AUTHOR)).toEqual(["X   b"]);
    expect(textsOf(out, "reject", AUTHOR)).toEqual(["a   b"]);
    expect(verifyOutput(normalizeDocx(docx(body)).docx, out, receipts).ok).toBe(true);
  });

  it("a point insertion inside a trailing-space element keeps the drop", () => {
    const body = P("10000001", PLAIN("trail   "));
    const { receipts, out } = runApply(body, [
      { op: "replace", ref: "f", author: AUTHOR, addr: "10000001", start: 3, end: 3, expectText: "", text: "Z" },
    ]);
    expect(receipts[0]!.status).toBe("applied");
    expect(textsOf(out, "accept", AUTHOR)).toEqual(["traZil"]);
    expect(textsOf(out, "reject", AUTHOR)).toEqual(["trail"]);
  });

  it("a point insertion inside a leading-space element keeps the drop", () => {
    const body = P("10000001", PLAIN("  lead"));
    const { receipts, out } = runApply(body, [
      { op: "replace", ref: "f", author: AUTHOR, addr: "10000001", start: 2, end: 2, expectText: "", text: "Z" },
    ]);
    expect(receipts[0]!.status).toBe("applied");
    expect(textsOf(out, "accept", AUTHOR)).toEqual(["leZad"]);
    expect(textsOf(out, "reject", AUTHOR)).toEqual(["lead"]);
  });
});

describe("normalize reports an enforced editing restriction", () => {
  const settings = (dp: string): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${dp}</w:settings>`;

  it("an enforced mode is reported; anything else is null", () => {
    const forms = writePackage(entriesOf(P("10000001", R("Body.")), settings(`<w:documentProtection w:edit="forms" w:enforcement="true"/>`)));
    expect(normalizeDocx(forms).protection).toBe("forms");
    const tracked = writePackage(entriesOf(P("10000001", R("Body.")), settings(`<w:documentProtection w:edit="trackedChanges" w:enforcement="1"/>`)));
    expect(normalizeDocx(tracked).protection).toBe("trackedChanges");
    const off = writePackage(entriesOf(P("10000001", R("Body.")), settings(`<w:documentProtection w:edit="forms" w:enforcement="false"/>`)));
    expect(normalizeDocx(off).protection).toBeNull();
    const unknownMode = writePackage(entriesOf(P("10000001", R("Body.")), settings(`<w:documentProtection w:edit="whatever" w:enforcement="true"/>`)));
    expect(normalizeDocx(unknownMode).protection).toBeNull();
    const plain = docx(P("10000001", R("Body.")));
    expect(normalizeDocx(plain).protection).toBeNull();
  });

  it("normalize never adds or removes the restriction", () => {
    const locked = writePackage(entriesOf(P("10000001", R("Body.")), settings(`<w:documentProtection w:edit="forms" w:enforcement="true"/>`)));
    const { docx: out } = normalizeDocx(locked);
    expect(new TextDecoder().decode(readPackage(out).get("word/settings.xml")!))
      .toContain('<w:documentProtection w:edit="forms" w:enforcement="true"/>');
  });
});
