# Migration Audit prototype — client-side .xlsx formula extractor

Status: **working prototype, all automated tests passing (67/67).**

A fully client-side, dependency-free .xlsx formula extractor for
canispreadsheet.com. No CDNs, no npm, no libraries — the file never leaves the
browser (privacy is the selling point).

## Files

| File | Purpose |
|---|---|
| `audit.html` | UI page: file input + drag-drop, results tables. Loads `audit.js`. |
| `audit.js` | All parsing logic (pure functions, no DOM use). Exposes `globalThis.XlsxAudit` in browsers and `module.exports` in Node. |
| `test.mjs` | Node test suite (`node test.mjs`). 67 assertions. |
| `make_fixtures.py` | Regenerates the three fixture .xlsx files (`/home/jon/venv/bin/python make_fixtures.py`). |
| `basic-2sheet.xlsx` | 2 sheets, 10 formulas (VLOOKUP/SUMIFS/TEXTJOIN/IF/…), strings containing `&`, `<`, and embedded `""` quotes. |
| `shared-formula.xlsx` | Hand-crafted OOXML with two real shared-formula groups (see note below). |
| `array-formula.xlsx` | openpyxl `ArrayFormula` fixtures (`t="array"`, single-cell and `E1:E3` spill ref). |

## What works (verified by tests)

- **ZIP reading in pure JS**: scans backward for the End-Of-Central-Directory
  record, walks the central directory (name, method, sizes, local offset), then
  slices each needed entry's data past its local header. Method 8 (deflate) is
  inflated with native `DecompressionStream('deflate-raw')`; method 0 (stored)
  is passed through. Clear error if `DecompressionStream` is missing.
- **Sheet mapping**: `xl/workbook.xml` + `xl/_rels/workbook.xml.rels` are parsed
  to map sheet name → sheet XML path via `r:id` (never assumes `sheet1.xml`
  ordering). Relationship targets are resolved (relative, `/absolute`, `../`).
- **Formula extraction** from each worksheet's `<c>`/`<f>` elements:
  - normal formulas;
  - **shared formulas** (`t="shared"`): the master (text + `si` + `ref`) is
    recorded; members (only `si`) are attributed to their master **and their
    formula text is reconstructed by reference-shifting** — relative refs move
    by the row/col delta, `$` anchors are preserved, quoted strings and quoted
    sheet names are never touched, out-of-range shifts become `#REF!`.
    Tested cases include growing ranges (`SUM($A$1:A1)` → `SUM($A$1:A2)`),
    mixed anchors (`A$1+$B2`), `LOG10(` not being mistaken for a ref, and a
    sheet named `'My A1 Sheet'`;
  - **array formulas** (`t="array"`) with their spill `ref` preserved.
- **XML entities** (`&amp; &lt; &gt; &quot; &apos; &#NN; &#xNN;`) are decoded in
  formula text and attributes; formulas containing strings with `&` and doubled
  quotes round-trip exactly.
- **Function tokenizing** matches the site's checker: strip double-quoted
  string literals, match `([A-Z][A-Z0-9_.]*)\s*\(`, uppercase, dedupe per
  formula. (`_xlfn.XLOOKUP(` correctly yields `XLOOKUP`; function-like text
  inside strings is ignored.)
- **Data model**:
  `{sheets: [{name, formulaCount}], formulas: [{sheet, cell, formula, functions[]}], functionCounts: {FN: n}, totals}`
  — plus extra per-formula fields (`type`, `sharedRole`, `sharedMaster`, `ref`)
  and extra totals (`sharedMasters`, `sharedMembers`, `arrayFormulas`).
  `functionCounts` counts each function once per formula that uses it.
- **Errors**: legacy `.xls` (CFB magic `D0 CF 11 E0`) → clear "save as .xlsx"
  message; non-ZIP input → clear message; ZIP without `xl/workbook.xml` → clear
  message; ZIP64 → explicit unsupported message.
- **Smoke-tested on real workbooks**: our shipped products parse correctly
  (Freelance-Business-Hub-SAMPLE.xlsx: 8 sheets / 423 formulas / 17 unique
  functions; Debt-Payoff-Tracker-SAMPLE.xlsx: 5 sheets / 165 formulas).

## Test results (2026-08-23)

`node test.mjs` on Node v18.19.1: **67 passed, 0 failed.**

Honest caveat: Node 18's `DecompressionStream` does not accept the
`'deflate-raw'` format (added to Node in 20.12+/21.2+; all current browsers have
it). `test.mjs` detects this and installs a zlib-backed shim with the same
`{readable, writable}` shape, so everything except the literal native
`DecompressionStream('deflate-raw')` constructor call is exercised. The
browser-native path uses the identical
`Blob → stream → pipeThrough → Response` pipeline. **Manual browser check** (not
yet done in this run): open `audit.html` over `http://` (e.g.
`python3 -m http.server`), drop each fixture, confirm the same counts as the
tests (10 / 8 / 3 formulas). `file://` also works in Firefox/Chrome since
nothing is fetched except `audit.js`.

## openpyxl and shared formulas (investigated)

openpyxl **always writes one plain `<f>` element per cell** — it never emits
`t="shared"` groups, and on read it silently expands shared formulas. So a
"dragged-down" openpyxl file is indistinguishable from individually typed
formulas. To test real shared formulas, `make_fixtures.py` hand-crafts the
minimal OOXML package (Content_Types, rels, workbook, worksheet with
`<f t="shared" ref="B1:B5" si="0">…</f>` master + `<f t="shared" si="0"/>`
members) and zips it with Python's `zipfile`. The fixture is valid: openpyxl's
reader loads it without complaint.

## Known limitations

- **Legacy `.xls` (BIFF)**: unsupported; detected by magic bytes with a clear
  "save as .xlsx" message (also pre-filtered by extension in the UI).
- **`.xlsm`**: works read-only — macros live in `xl/vbaProject.bin`, which we
  simply never read; VBA-defined functions are reported like any other
  function token.
- **External links** (`[1]Sheet1!A1`): the bracketed workbook index is left
  as-is in the formula text; the external target name is not resolved.
- **Defined names / named ranges**: not resolved. `=MyRange*2` shows literally;
  a defined name followed by `(` (LAMBDA-style calls) would be tokenized as a
  function name.
- **ZIP64** (> 4 GB or > 65535 entries): rejected with a clear message.
  Encrypted workbooks (which are CFB containers, not ZIPs) hit the `.xls`/CFB
  detection and get the legacy-format message — acceptable but the wording
  could mention encryption.
- **XML parsing is regex-based** (no DOMParser, so it stays pure/Node-testable).
  Fine for real-world producer output (Excel, LibreOffice, Google export,
  openpyxl — all tested shapes) including namespace-prefixed tags
  (`<x:c>`), but not a validating parser; e.g. CDATA sections inside `<f>`
  (never produced by real writers) are not handled.
- **Data-table formulas** (`t="dataTable"`) are counted as plain formulas;
  their special semantics are not modeled.
- **Shared-formula orphan members** (member whose master never appeared —
  malformed file) are recorded with empty formula text and
  `sharedRole: "orphan-member"` rather than invented text.
- `DecompressionStream('deflate-raw')` needs Chrome/Edge 103+, Firefox 113+,
  Safari 16.4+. Older browsers get the explicit unsupported-browser message.

## Integration notes (for merging into canispreadsheet.com)

1. The site is static GitHub Pages with no build step — copy `audit.js` next to
   the page and add `<script src="audit.js"></script>`; everything is ES5-ish
   script code (no modules, no build required). It only defines
   `globalThis.XlsxAudit`.
2. Call `XlsxAudit.auditXlsx(arrayBuffer)` → Promise of the data model above.
   All errors are thrown as `Error` with user-presentable `.message` strings —
   render them directly in the error box.
3. To wire into the existing compatibility checker: feed
   `Object.keys(result.functionCounts)` into the checker's lookup — the
   tokenization (strip strings, `([A-Z][A-Z0-9_.]*)\s*\(`, uppercase, dedupe)
   was deliberately matched to it, so counts line up 1:1 with what the checker
   would extract from the same formula text.
4. The UI in `audit.html` (dropzone + totals + per-sheet + function table +
   collapsible raw list capped at 200) is self-contained inside two
   easily-lifted blocks: the `<style>` rules and the second inline `<script>`.
   Styling is system-font/minimal on purpose to match the docs-site look.
5. Keep the privacy line ("analyzed in your browser and never uploaded") — it
   is true: the page performs zero network requests with the file.
6. Regenerating fixtures requires openpyxl:
   `/home/jon/venv/bin/python make_fixtures.py`, then `node test.mjs`.
