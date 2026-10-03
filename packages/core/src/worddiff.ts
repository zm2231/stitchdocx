// Word-level granularity: diff old and new span text by tokens so a rewrite
// that keeps most wording strikes only the changed tokens.
//
// The diff must treat whitespace as significant (diffWordsWithSpace, not
// diffWords): every matched part has to be an exact substring of the old text
// at the offset the edit coordinates assume. diffWords ignores whitespace when
// matching and reports the NEW string's whitespace for matched parts, which
// makes the offset run past the end of the old text and silently drops a
// leading space from the insertion (e.g. "(Week 4)" -> "(Week 5) days" lost
// the space before "days").
import { diffWordsWithSpace } from "diff";

export interface WordEdit {
  /** deleted range in span-relative coordinates */
  delStart?: number;
  delEnd?: number;
  /** inserted text at span-relative offset */
  insAt?: number;
  insText?: string;
}

export function wordEdits(oldText: string, newText: string): WordEdit[] {
  const parts = diffWordsWithSpace(oldText, newText);
  const edits: WordEdit[] = [];
  let oldOff = 0; // advances only over parts that consume old text
  for (const part of parts) {
    if (part.added) {
      edits.push({ insAt: oldOff, insText: part.value });
    } else if (part.removed) {
      edits.push({ delStart: oldOff, delEnd: oldOff + part.value.length });
      oldOff += part.value.length;
    } else {
      oldOff += part.value.length;
    }
  }
  return edits;
}
