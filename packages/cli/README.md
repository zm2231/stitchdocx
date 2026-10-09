# @stitchdocx/cli

Adapter CLI over the stitchdocx core. Five commands; each prints one JSON
object on stdout and exits 0. A non-zero exit means the engine itself failed —
per-op refusals are reported in the output, not as failures.

The core is a library: this CLI is the only place that touches the filesystem
and the only place the wire field `flagId` exists (the core's neutral name for
the same correlation field is `ref`; every op carries it and every receipt
echoes it).

## Whole-document refusals

A document the engine will not hold within its memory bounds is refused as a
whole: every command prints `{ "refused": { "reason", "detail" } }` and exits
0, so the caller can route the document elsewhere. `reason` is one of
`package_too_large` (an entry or the whole archive inflates past its limit),
`too_many_entries`, `duplicate_entry`, `encrypted_entry`,
`unreadable_archive`, or `document_too_large`: the main document part is over
8 MiB, or a parse would build more nodes than its budget allows (40,000 for
the main part's block structure plus the paragraphs an edit or decision
parses, 60,000 for any other part). Node budgets bound what bytes cannot: a
small part can hold a dense grid of tiny cells. The limits are sized for a
128 MB Worker isolate: the heaviest document built to sit just inside all of
them (a 7.9 MB main part of 19,000 table cells, with styles and numbering
parts near 60,000 nodes each) runs normalize, apply and verify in one workerd
request under an 87 MB V8 heap cap.

## Commands

### `normalize <in.docx> <out.docx>`

Assigns a valid unique `w14:paraId` to every paragraph that lacks one (and
repairs duplicates), adds the `w14` namespace declaration when missing, and
writes the result. Untouched parts stay byte-identical. `protection` is the
enforced editing restriction from `word/settings.xml` (`readOnly`,
`comments`, `trackedChanges` or `forms`), or null — surfaced so the caller
can ask the document owner before processing; it is never changed.

```json
{ "assigned": 3, "protection": null }
```

### `paragraphs <in.docx>`

Read-only projection: every paragraph in reader order with its address,
text, kind, style, numbering label and table position. The input is
normalized in memory first, so every paragraph carries an address without
rewriting the stored file.

```json
{ "paragraphs": [ { "addr": "00000001", "text": "…", "kind": "body", "editable": true } ] }
```

`editable` is `false` (with `readOnlyReason: "merged_removed_mark"`) for
display-only records whose text was carried across a removed paragraph mark —
for example a deleted heading merged into the first cell of the table that
followed it. Their text is what Word shows on reject, but their offsets span
more than one physical paragraph, so skip them when deriving operations; the
resolver and `apply` refuse such spans with `paragraph_not_editable`. The
underlying paragraphs keep their own addresses and are editable in their own
coordinates.

Text that Word does not render or will discard is never a valid edit target:
complex-field instructions are omitted from `text`, field results (including
`w:fldSimple` content and everything under a TOC's separate) are listed but
refused with `span_in_field`, and AlternateContent branches and office math
(`m:oMath`) are listed but refused with `span_crosses_structure`. Text Word hides is not
listed: a run's own `w:vanish` decides, else the paragraph and character
style chains combine by exclusive or, else `docDefaults` (`w:specVanish` and
`w:webHidden` text prints in Word and is listed). Engine insertions stay
visible, with an explicit `w:vanish w:val="0"` where copied formatting would
hide them. A table row the client tracked as inserted or deleted is
shown or dropped per Word's accept/reject rules, and spans inside a tracked
row refuse with `span_overlaps_revision`.

### `apply <in.docx> <ops.json> <out.docx>`

`ops.json` is `{ "date": "<ISO 8601 timestamp stamped on every revision>",
"ops": [...] }`; `date` is required and must parse as ISO 8601, and every op
needs a non-empty `author`. Ops (all require `flagId`, `author`, `addr`):

- `{ "op": "replace", "addr", "start", "end", "expectText", "text", "granularity"? }`
  — track-deletes `[start,end)` and track-inserts `text` after it. With
  `"granularity": "word"` only the changed tokens are struck.
- `{ "op": "delete", "addr", "start", "end", "expectText" }`
- `{ "op": "insertParagraphAfter", "addr", "text", "scope"?, "formatFrom"? }`
  — a new tracked-inserted paragraph after `addr` (or after the containing
  table with `"scope": "table"`). With `"formatFrom": "<addr>"` it takes that
  paragraph's properties (style, list, spacing) and the character formatting
  of its first run outside a link, as pressing Enter after it in Word does;
  a merged record refuses with `paragraph_not_editable` and an address that is
  not listed refuses with `addr_not_found`.

Spans are in the paragraph's original text coordinates and must match
`expectText` exactly — address only records listed with `editable: true`
(see `paragraphs`). Output is `{ "receipts": [...] }`; each receipt reports
`index`, `flagId`, `author`, `status` (`applied`/`refused`), `addr`,
`changeIds` (the `w:id`s the op wrote), and for applied ops `before`/`after`
paragraph text and `insertedAddr`.

Refusals carry a machine-readable `reason`: `addr_not_found`,
`span_out_of_range` (including an offset that would split a surrogate pair),
`span_text_mismatch`, `empty_text`,
`span_overlaps_revision`, `span_overlaps_prior_op`, `span_crosses_structure`,
`span_in_field`, `span_in_content_control`, `paragraph_not_editable`,
`insert_in_plain_text_control` (a plain-text content control may hold only
one paragraph).

### `finalize <in.docx> <decisions.json> <out.docx>`

`decisions.json` is `{ "date"?, "decisions": [{ "flagId", "changeIds",
"decision", "text"?, "expectText"? }] }`. A decision resolves exactly the
revisions its `changeIds` name (the receipt's); every other revision survives
untouched. `"text"` accepts with the member's wording: the op's whole span
reads as the given text, unchanged words included. For a replace or delete
the decision also carries the op's `expectText`, which locates that span:
after rejecting the decision's revisions, it is the one occurrence of
`expectText` covering all of them. The command fails, writing nothing, when
a decision names an id that is not a revision in the document, two decisions
name the same id, or the wording cannot be placed. Output is
`{ "results": [{ "index", "flagId", "decision", "changeIds" }] }`.

### `verify <original.docx> <output.docx> <receipts.json>`

Pure receipt check: `receipts.json` is `{ "receipts": [...] }` as produced by
`apply`. Fails ( `{ "ok": false, "failures": [...] }` ) unless the output is
the original plus exactly what the receipts claim — every changeId present
and authored by the receipt's author, rejecting one receipt's revisions
yields its `before`, accepting them yields its `after`, untouched paragraphs
byte-identical in text, every other package entry byte-identical with none
added or removed, the main part outside paragraphs canonically identical,
a touched paragraph carrying nothing beyond the receipt's own revisions, an
inserted paragraph holding nothing outside its revisions but its paragraph
properties, and no inserted text hidden — and only the main part differs.
