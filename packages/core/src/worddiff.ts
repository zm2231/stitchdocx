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
  old: string;
  new: string;
}

const WORD_CHAR = /[\p{L}\p{N}]/u;
const size = (r: Region): number => Math.max(r.old.length, r.new.length);

/**
 * wordEdits for a reader: an unchanged stretch between two changes is folded
 * into one change when it holds no letter or digit, or is shorter than both
 * neighbouring changes, so a rewritten sentence reads as one replacement
 * rather than alternating word swaps around kept spaces and separators.
 * The result keeps the wordEdits contract.
 */
export function readableWordEdits(oldText: string, newText: string): WordEdit[] {
  let regions: Region[] = [];
  for (const part of diffWordsWithSpace(oldText, newText)) {
    const changed = Boolean(part.added || part.removed);
    const last = regions[regions.length - 1];
    const region = last && last.changed && changed ? last : null;
    const target = region ?? { changed, old: "", new: "" };
    if (!region) regions.push(target);
    if (!part.added) target.old += part.value;
    if (!part.removed) target.new += part.value;
  }
  for (let merged = true; merged; ) {
    merged = false;
    for (let i = 1; i + 1 < regions.length; i++) {
      const [before, kept, after] = [regions[i - 1]!, regions[i]!, regions[i + 1]!];
      if (kept.changed || !before.changed || !after.changed) continue;
      if (WORD_CHAR.test(kept.old) && kept.old.length > Math.min(size(before), size(after))) continue;
      regions.splice(i - 1, 3, { changed: true, old: before.old + kept.old + after.old, new: before.new + kept.new + after.new });
      merged = true;
      break;
    }
  }
  const edits: WordEdit[] = [];
  let oldOff = 0;
  for (const r of regions) {
    const start = oldOff;
    oldOff += r.old.length;
    if (!r.changed) continue;
    if (r.old) edits.push({ delStart: start, delEnd: oldOff });
    if (r.new) edits.push({ insAt: oldOff, insText: r.new });
  }
  return edits;
}

