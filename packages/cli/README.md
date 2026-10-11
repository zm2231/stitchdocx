# @stitchdocx/cli

The command-line front end to the stitchdocx engine: five commands that read and write `.docx` files and print JSON. This page is the field reference. For what the engine is and a walkthrough, start with the [top-level README](../../README.md).

## Conventions

- Every command prints one JSON object on stdout and exits 0.
- A refused edit is a normal result, reported in the output with a `reason`. A non-zero exit means the engine itself failed.
- Every edit carries a `flagId` and every receipt echoes it back, so you can match results to your own records. The core library calls the same field `ref`. This CLI is the only part of stitchdocx that touches the filesystem.

| Command | Input | Output |
|---|---|---|
| [`normalize`](#normalize-indocx-outdocx) | a `.docx` | `{ "assigned", "protection" }` |
| [`paragraphs`](#paragraphs-indocx) | a `.docx` | `{ "paragraphs": [...] }` |
| [`apply`](#apply-indocx-opsjson-outdocx) | a `.docx` and `ops.json` | `{ "receipts": [...] }` |
| [`finalize`](#finalize-indocx-decisionsjson-outdocx) | a redlined `.docx` and `decisions.json` | `{ "results": [...] }` |
| [`verify`](#verify-originaldocx-outputdocx-receiptsjson) | the original, the output and `receipts.json` | `{ "ok", "failures" }` |

## `normalize <in.docx> <out.docx>`

Gives every paragraph without a `w14:paraId` a valid unique one, repairs duplicates, adds the `w14` namespace declaration if it's missing, and writes the result. Parts it doesn't need to touch stay byte-identical.

```json
{ "assigned": 3, "protection": null }
```

`protection` is the editing restriction the file enforces in `word/settings.xml` (`readOnly`, `comments`, `trackedChanges` or `forms`), or null. It is reported so you can ask the document's owner before going further, and it is never changed.

## `paragraphs <in.docx>`

Read-only. Lists every paragraph in reading order with its address, text, kind, style, list label and table position. The input is normalized in memory first, so every paragraph has an address even if the stored file has none.

```json
{ "paragraphs": [ { "addr": "10000008", "text": "The engagement is billed at $3,650 Fixed Fee, payable in advance.", "kind": "body", "editable": true } ] }
```

**Records you can't edit.** `editable` is `false`, with `readOnlyReason: "merged_removed_mark"`, for a record whose text runs across a removed paragraph mark. One example: a deleted heading merged into the first cell of the table after it. The text is what Word shows on reject, but its offsets span more than one physical paragraph, so skip it when building edits; `apply` refuses such spans with `paragraph_not_editable`. The underlying paragraphs keep their own addresses and are editable in their own coordinates.

**What text is listed and what can be targeted.**

| Text | Listed | Editable |
|---|---|---|
| Complex-field instructions | No | No |
| Field results, including `w:fldSimple` and everything after a table of contents' separator | Yes | Refused with `span_in_field` |
| AlternateContent branches and office math (`m:oMath`) | Yes | Refused with `span_crosses_structure` |
| Text in a hyperlink or smart tag, or a span that reaches into or out of one | Yes | Yes (see below) |
| A span reaching into any other wrapper (`w:customXml`, a field result, a content control) | Yes | Refused with `span_crosses_structure` |
| A span inside a table row the client tracked as inserted or deleted | Shown or dropped per Word's accept/reject rules | Refused with `span_overlaps_revision` |
| Text Word hides | No | No |

A span that crosses a link edge gets one `w:del` per parent, and a `replace`'s new text goes after the deleted text, outside the link, formatted like the first deleted run outside it.

Hidden text follows Word's rules: a run's own `w:vanish` decides; failing that, the paragraph and character style chains combine by exclusive or; failing that, `docDefaults` decides. `w:specVanish` and `w:webHidden` text prints in Word, so it is listed.

**Formatting of inserted words.** Insertions stay visible, with an explicit `w:vanish w:val="0"` wherever copied formatting would hide them. Words that replace deleted words take the formatting of the first replaced run in the same parent. Other inserted words take the formatting of the visible text just before them in the same parent (or just after, if nothing comes before), as typing in Word does. So a word added inside a link looks like the link, and one added beside it looks like the text around it.

## `apply <in.docx> <ops.json> <out.docx>`

```json
{
  "date": "2026-10-11T12:00:00Z",
  "ops": [
    { "op": "replace", "flagId": "fee", "author": "Reviewer", "addr": "10000008",
      "start": 28, "end": 34, "expectText": "$3,650", "text": "$4,200" }
  ]
}
```

`date` is required, must parse as ISO 8601, and is stamped on every revision. Every op needs `flagId`, a non-empty `author`, and `addr`.

| Op | Fields | Effect |
|---|---|---|
| `replace` | `addr`, `start`, `end`, `expectText`, `text`, `granularity`? | Tracks the deletion of `[start, end)` and the insertion of `text` after it. |
| `delete` | `addr`, `start`, `end`, `expectText` | Tracks the deletion of `[start, end)`. |
| `insertParagraphAfter` | `addr`, `text`, `scope`?, `formatFrom`? | Adds a new tracked paragraph after `addr`, or after the table containing it with `"scope": "table"`. |

Spans use the paragraph's original text coordinates and must match `expectText` exactly. Address only records listed with `editable: true`.

**`"granularity": "word"`** strikes only the words that change, grouped so a person can read them. A short unchanged word or bare punctuation between two changes is folded into them, and a span changed in three or more places that keeps less than two thirds of its wording becomes one replacement. If a grouped change would reach into a field or another fixed element, the finest word-level changes are used instead.

**`"formatFrom": "<addr>"`** gives the new paragraph that paragraph's properties (style, list, spacing) and the character formatting of its first run outside a link, as pressing Enter after it in Word does. A merged record refuses with `paragraph_not_editable`; an address that isn't listed refuses with `addr_not_found`.

The output is one receipt per op, in order:

```json
{ "receipts": [ { "index": 0, "flagId": "fee", "author": "Reviewer", "status": "applied", "addr": "10000008",
  "changeIds": ["1", "2"],
  "before": "The engagement is billed at $3,650 Fixed Fee, payable in advance.",
  "after": "The engagement is billed at $4,200 Fixed Fee, payable in advance." } ] }
```

`changeIds` are the `w:id`s the op wrote. Applied ops also carry `before` and `after` paragraph text, and an inserted paragraph's `insertedAddr`. A refused op carries `status: "refused"` and a `reason`:

| Reason | Meaning |
|---|---|
| `addr_not_found` | No listed paragraph has that address. |
| `span_out_of_range` | The offsets fall outside the paragraph, or would split a surrogate pair. |
| `span_text_mismatch` | The text at the offsets differs from `expectText`. |
| `empty_text` | An inserted paragraph's text is empty or only whitespace. |
| `span_overlaps_revision` | The span overlaps a revision already in the file. |
| `span_overlaps_prior_op` | The span overlaps an earlier op in the same batch. |
| `span_crosses_structure` | The span reaches into a wrapper the engine can't split (see the table above). |
| `span_in_field` | The span is inside a field result. |
| `span_in_content_control` | The span is inside an inline content control. |
| `paragraph_not_editable` | The address is a merged, display-only record. |
| `insert_in_plain_text_control` | A plain-text content control can hold only one paragraph. |

## `finalize <in.docx> <decisions.json> <out.docx>`

```json
{ "decisions": [
  { "flagId": "fee",  "changeIds": ["1", "2"], "decision": "accept" },
  { "flagId": "late", "changeIds": ["3", "4"], "decision": "reject" }
] }
```

`date` is optional. A decision resolves exactly the revisions its `changeIds` name (the receipt's), and every other revision is left as it was.

To accept with different wording, add `"text"`: the op's whole span then reads as that text, unchanged words included. For a `replace` or `delete`, the decision also carries the op's `expectText`, which locates the span: after rejecting the decision's revisions, it is the one occurrence of `expectText` that covers all of them.

The command fails and writes nothing when a decision names an id that isn't a revision in the document, when two decisions name the same id, or when the wording can't be placed. The output is:

```json
{ "results": [ { "index": 0, "flagId": "fee", "decision": "accept", "changeIds": ["1", "2"] } ] }
```

## `verify <original.docx> <output.docx> <receipts.json>`

Checks the output against the receipts `apply` produced (`{ "receipts": [...] }`) and returns `{ "ok": true, "failures": [] }` only if every one of these holds:

- Every change id is present and written by its receipt's author.
- Rejecting a receipt's revisions gives its `before`, and accepting them gives its `after`.
- Untouched paragraphs have identical text, and the main part outside paragraphs is identical once canonicalized.
- A touched paragraph carries nothing beyond its receipt's own revisions.
- An inserted paragraph holds nothing outside its revisions except its paragraph properties.
- No inserted text is hidden.
- Only the main part differs: every other package entry is byte-identical, with none added or removed.

Otherwise it returns `{ "ok": false, "failures": [...] }`.

## Documents that are refused whole

A document the engine can't hold within its memory bounds is refused before any work, and every command prints `{ "refused": { "reason", "detail" } }` and exits 0, so you can route the file somewhere else.

| Reason | Cause |
|---|---|
| `package_too_large` | An entry, or the whole archive, inflates past its limit. |
| `too_many_entries` | The archive holds more entries than the limit. |
| `duplicate_entry` | Two entries share a name. |
| `encrypted_entry` | An entry is encrypted. |
| `unreadable_archive` | An entry won't inflate, or uses a compression method the engine doesn't support. |
| `document_too_large` | The main part is over 8 MiB, or a parse would build more nodes than its budget. |

The node budgets are 40,000 for the main part's block structure plus the paragraphs an edit or decision parses, and 60,000 for any other part. They catch what a byte limit can't: a small part can still hold a dense grid of tiny cells. The limits are sized for a 128 MB Worker isolate. The heaviest document built to sit just inside all of them (a 7.9 MB main part with 19,000 table cells, and styles and numbering parts near 60,000 nodes each) runs normalize, apply and verify in one workerd request under an 87 MB V8 heap cap.
