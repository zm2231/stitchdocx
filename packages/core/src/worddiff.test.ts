// The word-diff contract: applying wordEdits(old, new) to old must reproduce
// new exactly, and every matched part must be an exact substring of the old
// text at the offset the edit coordinates assume. A diff that ignores
// whitespace when matching (jsdiff's diffWords) violates both and silently
// drops a leading space from insertions after the last old token.
import { describe, expect, it } from "vitest";
import { readableWordEdits, wordEdits, type WordEdit } from "./worddiff.js";

/** rebuild newText from oldText + edits; the slices taken between edit
 *  coordinates are exactly the matched parts, so this asserts both contract
 *  halves: matched parts are old substrings at the assumed offsets, and the
 *  edits reproduce newText */
function reconstruct(oldText: string, newText: string, diff: (o: string, n: string) => WordEdit[] = wordEdits): string {
  let out = "";
  let off = 0;
  for (const e of diff(oldText, newText)) {
    if (e.insAt !== undefined) {
      if (e.insAt < off) throw new Error(`insertion at ${e.insAt} precedes consumed offset ${off}`);
      out += oldText.slice(off, e.insAt);
      out += e.insText ?? "";
      off = e.insAt;
    }
    if (e.delStart !== undefined && e.delEnd !== undefined) {
      if (e.delStart < off) throw new Error(`deletion ${e.delStart}..${e.delEnd} precedes consumed offset ${off}`);
      out += oldText.slice(off, e.delStart);
      off = e.delEnd;
    }
  }
  return out + oldText.slice(off);
}

describe("wordEdits contract", () => {
  it("keeps the space when text is added after the last old token", () => {
    // the reported defect: "(Week 5)days" / "within ten (10)days"
    expect(reconstruct("(Week 4)", "(Week 5) days")).toBe("(Week 5) days");
    expect(reconstruct("(Week 4)", "within ten (10) days")).toBe("within ten (10) days");
    expect(reconstruct("(Week 4)", "ten days")).toBe("ten days");
    expect(reconstruct("(Week 4)", "a b")).toBe("a b");
  });

  it("strikes only changed tokens when wording is preserved", () => {
    const edits = wordEdits("net-30 days", "net-45 days");
    // tokens attach leading whitespace, so the deletion covers "30" and the
    // insertion "45" lands before the shared " days" — applying both
    // reproduces "net-45 days"
    expect(edits).toEqual([
      { delStart: 4, delEnd: 6 },
      { insAt: 6, insText: "45" },
    ]);
  });

  it("handles identical text, empty old and empty new", () => {
    expect(wordEdits("same", "same")).toEqual([]);
    expect(reconstruct("", "inserted")).toBe("inserted");
    expect(reconstruct("removed", "")).toBe("");
  });

  it.each([["wordEdits", wordEdits], ["readableWordEdits", readableWordEdits]] as const)("%s matches the contract on 10,000 random old/new pairs", { timeout: 60_000 }, (_name, diff) => {
    const TOKENS = [
      "alpha", "Bravo", "$12,500", "net-30", "Client", "deliverables", "within ten (10) days",
      "§4.2", "“quoted”", "é", "\u00A0", " ", "  ", "\t", "\n", "(", ")", ".", ",", "x", "Q5", "2026",
      "条款", "الشرط", "✓", "\u2009", "\u2007", "\u202F",
    ];
    let seed = 42;
    const rnd = (): number => ((seed = (seed * 1103515245 + 12345) % 2147483648) / 2147483648);
    const pick = (a: string[]): string => a[Math.floor(rnd() * a.length)]!;
    const gen = (): string => Array.from({ length: 1 + Math.floor(rnd() * 6) }, () => pick(TOKENS)).join("");
    for (let i = 0; i < 10_000; i++) {
      const oldText = gen();
      let newText = gen();
      if (rnd() < 0.5) {
        const at = Math.floor(rnd() * (oldText.length + 1));
        newText = oldText.slice(0, at) + pick(TOKENS) + oldText.slice(at);
      }
      let recon: string;
      try {
        recon = reconstruct(oldText, newText, diff);
      } catch (err) {
        throw new Error(`pair ${i}: ${JSON.stringify({ oldText, newText })}: ${String(err)}`);
      }
      if (recon !== newText) {
        throw new Error(`pair ${i}: reconstruction ${JSON.stringify(recon)} != ${JSON.stringify(newText)} (old ${JSON.stringify(oldText)})`);
      }
    }
  });
});

describe("readableWordEdits", () => {
  it("turns a rewritten sentence into one replacement", () => {
    const before = "Storyteller | Technology Enthusiast | Lifelong Learner";
    const after = "Solutions Engineer turning technical depth into clear buyer value";
    expect(readableWordEdits(before, after)).toEqual([
      { delStart: 0, delEnd: before.length },
      { insAt: before.length, insText: after },
    ]);
    expect(wordEdits(before, after).length).toBeGreaterThan(2);
  });

  it("keeps word-level edits when most of the wording stays", () => {
    const edits = readableWordEdits("Led value-based demos for enterprise buyers", "Led tailored demos for 40 enterprise buyers");
    expect(edits).toEqual([
      { delStart: 4, delEnd: 15 },
      { insAt: 15, insText: "tailored" },
      { insAt: 26, insText: "40 " },
    ]);
    expect(readableWordEdits("net-30 days", "net-45 days")).toEqual(wordEdits("net-30 days", "net-45 days"));
  });

  it("folds a separator-only gap between two changes", () => {
    expect(readableWordEdits("A | B", "C | D")).toEqual([{ delStart: 0, delEnd: 5 }, { insAt: 5, insText: "C | D" }]);
  });

  it("keeps a kept word that is as long as a neighbouring change", () => {
    expect(readableWordEdits("aaa x bbb", "AAA x BBB")).toEqual(wordEdits("aaa x bbb", "AAA x BBB"));
    expect(readableWordEdits("alpha x bravo", "ALPHA x BRAVO")).toEqual([{ delStart: 0, delEnd: 13 }, { insAt: 13, insText: "ALPHA x BRAVO" }]);
  });

  it("adds no more than linear work on top of the diff for a long alternating rewrite", () => {
    const oldText = Array.from({ length: 1000 }, (_, i) => `old${i} k`).join(" ") + " closing words";
    const newText = Array.from({ length: 1000 }, (_, i) => `NEW${i} k`).join(" ") + " CLOSING WORDS";
    const timed = (diff: () => WordEdit[]): [WordEdit[], number] => {
      const t0 = performance.now();
      const edits = diff();
      return [edits, performance.now() - t0];
    };
    const [, diffOnly] = timed(() => wordEdits(oldText, newText));
    const [edits, grouped] = timed(() => readableWordEdits(oldText, newText));
    expect(edits).toEqual([{ delStart: 0, delEnd: oldText.length }, { insAt: oldText.length, insText: newText }]);
    expect(grouped).toBeLessThan(diffOnly * 2 + 50);
  });

  it("turns a line changed in several places into one replacement when little of it stays", () => {
    const before = "www.linkedin.com/in/sample• sample@example.com • (555) 010-0199 ";
    const after = "Austin, TX | sample@example.com | (555) 010-0199 | linkedin.com/in/sample";
    expect(wordEdits(before, after).length).toBeGreaterThan(4);
    expect(readableWordEdits(before, after)).toEqual([{ delStart: 0, delEnd: before.length }, { insAt: before.length, insText: after }]);
  });

  it("keeps a few small edits word-level when most of the line stays", () => {
    const before = "Led 5 demos for 10 enterprise buyers in 3 regions every quarter";
    const after = "Led 8 demos for 12 enterprise buyers in 4 regions every quarter";
    expect(readableWordEdits(before, after)).toEqual(wordEdits(before, after));
    expect(readableWordEdits(before, after).filter((e) => e.insText !== undefined)).toHaveLength(3);
  });
});
