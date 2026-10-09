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

interface Region {
  changed: boolean;
  oldStart: number;
  oldEnd: number;
  newStart: number;
  newEnd: number;
  /** an unchanged region holding a letter or digit */
  wordy: boolean;
}

const WORD_CHAR = /[\p{L}\p{N}]/u;
const size = (r: Region): number => Math.max(r.oldEnd - r.oldStart, r.newEnd - r.newStart);

/**
 * wordEdits for a reader: an unchanged stretch between two changes is folded
 * into one change when it holds no letter or digit, or is strictly shorter
 * than both neighbouring changes, so a rewritten sentence reads as one
 * replacement rather than alternating word swaps around kept spaces and
 * separators. Folding only grows changes, so folding greedily from the left
 * reaches the same result as any order, in linear time.
 * The result keeps the wordEdits contract.
 */
export function readableWordEdits(oldText: string, newText: string): WordEdit[] {
  const stack: Region[] = [];
  let oldOff = 0;
  let newOff = 0;
  for (const part of diffWordsWithSpace(oldText, newText)) {
    const changed = Boolean(part.added || part.removed);
    const oldLen = part.added ? 0 : part.value.length;
    const newLen = part.removed ? 0 : part.value.length;
    const last = stack[stack.length - 1];
    if (last && last.changed === changed) {
      last.oldEnd += oldLen;
      last.newEnd += newLen;
      last.wordy ||= !changed && WORD_CHAR.test(part.value);
    } else {
      stack.push({ changed, oldStart: oldOff, oldEnd: oldOff + oldLen, newStart: newOff, newEnd: newOff + newLen, wordy: !changed && WORD_CHAR.test(part.value) });
    }
    oldOff += oldLen;
    newOff += newLen;
    if (!changed) continue;
    while (stack.length >= 3) {
      const after = stack[stack.length - 1]!;
      const kept = stack[stack.length - 2]!;
      const before = stack[stack.length - 3]!;
      if (kept.changed || !before.changed) break;
      if (kept.wordy && kept.oldEnd - kept.oldStart >= Math.min(size(before), size(after))) break;
      stack.length -= 2;
      before.oldEnd = after.oldEnd;
      before.newEnd = after.newEnd;
    }
  }
  const edits: WordEdit[] = [];
  for (const r of stack) {
    if (!r.changed) continue;
    if (r.oldEnd > r.oldStart) edits.push({ delStart: r.oldStart, delEnd: r.oldEnd });
    if (r.newEnd > r.newStart) edits.push({ insAt: r.oldEnd, insText: newText.slice(r.newStart, r.newEnd) });
  }
  return edits;
}
