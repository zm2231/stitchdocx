// Word-level granularity: diff old and new span text by tokens so a rewrite
// that keeps most wording strikes only the changed tokens.
import { diffWords } from "diff";

export interface WordEdit {
  /** deleted range in span-relative coordinates */
  delStart?: number;
  delEnd?: number;
  /** inserted text at span-relative offset */
  insAt?: number;
  insText?: string;
}

export function wordEdits(oldText: string, newText: string): WordEdit[] {
  const parts = diffWords(oldText, newText);
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
