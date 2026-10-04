// Complex-field cursor state. A field's begin/separate/end fldChars may sit in
// different paragraphs (a TOC spans the whole document), so the instruction/
// result stack is threaded through the whole body walk — both the addressing
// walk (body.ts) and the text walker (projection.ts) advance the same shared
// cursor so text and refusal checks can never disagree.
import { decodeAttrValue, kid, kids, type XElement } from "@stitchdocx/package-io";

export function attrOf(el: XElement, name: string): string | null {
  const a = el.attrs.find((x) => x.name === name);
  return a ? decodeAttrValue(a.raw) : null;
}

/** advance the field cursor over one fldChar ("begin" pushes the instruction
 *  state, "separate" flips the innermost field to its result, "end" pops) */
export function stepFields(fields: string[], el: XElement): void {
  if (el.name !== "w:fldChar") return;
  const type = attrOf(el, "w:fldCharType");
  if (type === "begin") fields.push("instr");
  else if (type === "separate") {
    if (fields.length) fields[fields.length - 1] = "result";
  } else if (type === "end") fields.pop();
}

/** advance the field cursor across a whole subtree, in document order */
export function advanceFields(node: XElement, fields: string[]): void {
  for (const c of kids(node)) {
    stepFields(fields, c);
    advanceFields(c, fields);
  }
}
