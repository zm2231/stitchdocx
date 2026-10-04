// Accepting every tracked change: the result holds no revision markup in any
// story part, and its body reads exactly as the accept projection of the input.
import { describe, it, expect } from "vitest";
import { encodePart, readPackage, writePackage, entryBytes } from "@stitchdocx/package-io";
import { acceptAllDocx, normalizeDocx, openDocument } from "./index.js";

const NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"`;
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const ANY_REVISION = /<w:(ins|del|moveFrom|moveTo|rPrChange|pPrChange|sectPrChange|tblPrChange|trPrChange|tcPrChange|tblGridChange|numberingChange|cellIns|cellDel|cellMerge|moveFromRangeStart|moveFromRangeEnd|moveToRangeStart|moveToRangeEnd|customXml(?:Ins|Del|MoveFrom|MoveTo)Range(?:Start|End))[\s/>]/;

let nextId = 100;
const rev = (name: string) => `<w:${name} w:id="${nextId++}" w:author="Client" w:date="2026-10-01T00:00:00Z"/>`;
const R = (t: string) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const DEL = (t: string) => `<w:del w:id="${nextId++}" w:author="Client" w:date="2026-10-01T00:00:00Z"><w:r><w:delText xml:space="preserve">${t}</w:delText></w:r></w:del>`;
const INS = (t: string) => `<w:ins w:id="${nextId++}" w:author="Client" w:date="2026-10-01T00:00:00Z">${R(t)}</w:ins>`;
/** `mark` names the revisions on the paragraph mark */
const P = (inner: string, ...mark: string[]) => `<w:p>${mark.length ? `<w:pPr><w:rPr>${mark.map(rev).join("")}</w:rPr></w:pPr>` : ""}${inner}</w:p>`;
const TC = (...paras: string[]) => `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${paras.join("")}</w:tc>`;
const TR = (cells: string[], rowRev?: string) => `<w:tr>${rowRev ? `<w:trPr>${rev(rowRev)}</w:trPr>` : ""}${cells.join("")}</w:tr>`;
const TBL = (...rows: string[]) => `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>${rows.join("")}</w:tbl>`;
const SDT = (...paras: string[]) => `<w:sdt><w:sdtPr/><w:sdtContent>${paras.join("")}</w:sdtContent></w:sdt>`;

function docx(body: string, parts: Record<string, { xml: string; type: string; rel: string }> = {}): Uint8Array {
  const rels = Object.entries(parts).map(([name, p], i) => `<Relationship Id="rIdX${i}" Type="${REL}/${p.rel}" Target="${name.replace("word/", "")}"/>`).join("");
  const overrides = Object.entries(parts).map(([name, p]) => `<Override PartName="/${name}" ContentType="${p.type}"/>`).join("");
  const entries = new Map<string, Uint8Array>([
    ["[Content_Types].xml", encodePart(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>${overrides}</Types>`)],
    ["_rels/.rels", encodePart(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`)],
    ["word/_rels/document.xml.rels", encodePart(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`)],
    ["word/document.xml", encodePart(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`)],
  ]);
  for (const [name, p] of Object.entries(parts)) entries.set(name, encodePart(p.xml));
  return normalizeDocx(writePackage(entries)).docx;
}

const texts = (bytes: Uint8Array, view?: "reject") => openDocument(bytes).projection(view ? () => "reject" : undefined).map((p) => p.text);

function partXml(bytes: Uint8Array, name: string): string {
  return new TextDecoder().decode(entryBytes(readPackage(bytes).get(name)!));
}

function acceptAll(bytes: Uint8Array): string[] {
  const out = acceptAllDocx(bytes);
  for (const name of readPackage(out).keys()) {
    if (/^word\/.*\.xml$/.test(name)) expect(partXml(out, name), name).not.toMatch(ANY_REVISION);
  }
  const got = texts(out);
  expect(got).toEqual(texts(bytes));
  return got;
}

describe("acceptAllDocx", () => {
  it("joins a paragraph whose mark was deleted to the next one, along a chain", () => {
    const body = P(R("One ") + DEL("gone"), "del") + P(R("two "), "del") + P(R("three")) + P(R("Last"));
    expect(acceptAll(docx(body))).toEqual(["One two three", "Last"]);
  });

  it("drops a deleted empty paragraph whose deleted mark joins an empty one", () => {
    const body = P(R("Before")) + P("", "del") + P(DEL("Food truck ??"), "del") + P("") + P(R("After"));
    expect(acceptAll(docx(body))).toEqual(["Before", "", "After"]);
  });

  it("joins a paragraph before a table to the first cell of the first row that stays", () => {
    const body = P(R("Lead "), "del") + TBL(TR([TC(P(R("gone row"))), TC(P(R("x")))], "del"), TR([TC(P(R("cell"))), TC(P(R("y")))])) + P(R("End"));
    expect(acceptAll(docx(body))).toEqual(["Lead cell", "y", "End"]);
  });

  it("joins a paragraph before a content control to its first paragraph, and keeps one at its end", () => {
    const body = P(R("Lead "), "del") + SDT(P(R("inside")), P(R("tail"), "del")) + P(R("End"));
    expect(acceptAll(docx(body))).toEqual(["Lead inside", "tail", "End"]);
  });

  it("at the end of a cell, drops an emptied paragraph and keeps the cell's only paragraph", () => {
    const body = TBL(TR([TC(P(R("kept")), P(DEL("old"), "del")), TC(P(DEL("solo"), "del"))])) + P(R("End"));
    const bytes = docx(body);
    expect(texts(bytes)).toEqual(["kept", "End"]);
    const out = acceptAllDocx(bytes);
    expect(partXml(out, "word/document.xml")).not.toMatch(ANY_REVISION);
    expect(texts(out)).toEqual(["kept", "", "End"]);
    expect(partXml(out, "word/document.xml")).toMatch(/<w:tc>(?:(?!<\/w:tc>).)*<w:p\b(?:(?!<\/w:tc>).)*<\/w:tc>\s*<\/w:tr>/);
  });

  it("keeps the last paragraph of the body", () => {
    expect(acceptAll(docx(P(R("Only")) + P(R("Final"), "del")))).toEqual(["Only", "Final"]);
  });

  it("removes a mark that was inserted and later deleted, under either view", () => {
    const bytes = docx(P(R("Head")) + P("", "ins", "del") + P(R("Body")));
    expect(texts(bytes)).toEqual(["Head", "Body"]);
    expect(texts(bytes, "reject")).toEqual(["Head", "Body"]);
    expect(acceptAll(bytes)).toEqual(["Head", "Body"]);
  });

  it("removes deleted rows, and a table left without rows", () => {
    const body = TBL(TR([TC(P(R("a"))), TC(P(R("b")))], "del"), TR([TC(P(R("c"))), TC(P(R("d")))], "ins")) + TBL(TR([TC(P(R("x"))), TC(P(R("y")))], "del")) + P(R("End"));
    const bytes = docx(body);
    expect(acceptAll(bytes)).toEqual(["c", "d", "End"]);
    expect(partXml(acceptAllDocx(bytes), "word/document.xml").match(/<w:tbl>/g)).toHaveLength(1);
  });

  it.each([
    ["a content control", (row: string) => `<w:sdt><w:sdtPr/><w:sdtContent>${row}</w:sdtContent></w:sdt>`],
    ["custom XML", (row: string) => `<w:customXml w:element="row">${row}</w:customXml>`],
  ])("removes a row wrapped in %s, and the table it empties", (_, wrap) => {
    const nested = (rowRev: string) => TBL(wrap(TR([TC(P(R("inner"))), TC(P(R("cell")))], rowRev)));
    const accepted = docx(TBL(TR([TC(P(R("outer")) + nested("del")), TC(P(R("x")))])) + TBL(wrap(TR([TC(P(R("gone"))), TC(P(R("row")))], "del"))) + P(R("End")));
    expect(acceptAll(accepted)).toEqual(["outer", "x", "End"]);
    const xml = partXml(acceptAllDocx(accepted), "word/document.xml");
    expect(xml.match(/<w:tbl>/g)).toHaveLength(1);
    expect(xml).not.toMatch(/<w:(sdt|customXml)\b/);
    expect(xml.match(/<\/w:tc>/g)!.length).toBe(xml.match(/<\/w:p><\/w:tc>/g)!.length);

    const rejected = docx(TBL(wrap(TR([TC(P(R("added"))), TC(P(R("row")))], "ins"))) + P(R("End")));
    const id = /<w:ins w:id="(\d+)"/.exec(partXml(rejected, "word/document.xml"))![1]!;
    const doc = openDocument(rejected);
    doc.decide([{ ref: "row", changeIds: [id], decision: "reject" }]);
    const out = partXml(doc.save(), "word/document.xml");
    expect(out).not.toMatch(/<w:(tbl|sdt|customXml)\b/);
    expect(texts(doc.save())).toEqual(["End"]);
  });

  it("clears custom XML revision ranges in the body and a header", () => {
    const W = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`;
    const ranges = ["Ins", "Del", "MoveFrom", "MoveTo"].map((k) => `<w:customXml${k}RangeStart w:id="${nextId++}" w:author="Client"/><w:customXml${k}RangeEnd w:id="${nextId++}"/>`).join("");
    const parts = { "word/header1.xml": { xml: `<w:hdr ${W}>${ranges}${P(R("Header"))}</w:hdr>`, type: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml", rel: "header" } };
    const bytes = docx(P(ranges + R("Body")) + ranges + P(R("End")), parts);
    expect(acceptAll(bytes)).toEqual(["Body", "End"]);
    expect(partXml(acceptAllDocx(bytes), "word/header1.xml")).not.toMatch(ANY_REVISION);
  });

  it("returns the input itself when nothing is tracked", () => {
    const bytes = docx(P(R("Clean")) + TBL(TR([TC(P(R("a"))), TC(P(R("b")))])));
    expect(acceptAllDocx(bytes)).toBe(bytes);
  });

  it("accepts in headers and footnotes, and clears cell, numbering and grid changes", () => {
    const W = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`;
    const parts = {
      "word/header1.xml": { xml: `<w:hdr ${W}>${P(R("Draft ") + INS("v2") + DEL("v1"))}</w:hdr>`, type: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml", rel: "header" },
      "word/footnotes.xml": { xml: `<w:footnotes ${W}><w:footnote w:id="1">${P(DEL("old note") + INS("new note"))}</w:footnote></w:footnotes>`, type: "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml", rel: "footnotes" },
    };
    const cell = (mark: string, text: string) => `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/>${mark}</w:tcPr>${P(R(text))}</w:tc>`;
    const grid = `<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:tblGridChange w:id="${nextId++}"><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid></w:tblGridChange></w:tblGrid>`;
    const table = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>${grid}<w:tr>${cell(rev("cellIns"), "added")}${cell(rev("cellDel"), "removed")}${cell(`<w:vMerge/><w:cellMerge w:id="${nextId++}" w:author="Client" w:vMerge="cont"/>`, "merged")}</w:tr></w:tbl>`;
    const numbered = `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/><w:numberingChange w:id="${nextId++}" w:author="Client" w:original="%1."/></w:numPr></w:pPr>${R("item")}</w:p>`;
    const bytes = docx(table + numbered + P(R("End")), parts);
    const out = acceptAllDocx(bytes);
    for (const name of ["word/document.xml", "word/header1.xml", "word/footnotes.xml"]) expect(partXml(out, name), name).not.toMatch(ANY_REVISION);
    expect(partXml(out, "word/header1.xml")).toContain("v2");
    expect(partXml(out, "word/header1.xml")).not.toContain("v1");
    expect(partXml(out, "word/footnotes.xml")).toContain("new note");
    expect(partXml(out, "word/footnotes.xml")).not.toContain("old note");
    expect(texts(out)).toEqual(["added", "merged", "item", "End"]);
  });

  it("decide merges on an accepted deleted mark, as the projection reads it", () => {
    const bytes = docx(P(R("Join "), "del") + P(R("me")));
    const doc = openDocument(bytes);
    const id = /<w:del w:id="(\d+)"/.exec(partXml(bytes, "word/document.xml"))![1]!;
    doc.decide([{ ref: "mark", changeIds: [id], decision: "accept" }]);
    expect(texts(doc.save())).toEqual(["Join me"]);
  });
});
