# spreadsheet-compat

[![DOI](https://zenodo.org/badge/DOI/10.5281/zenodo.21974469.svg)](https://doi.org/10.5281/zenodo.21974469)

"caniuse.com for spreadsheets" — a database of *tested, real, executed*
function behavior across Excel, Google Sheets, and LibreOffice Calc.

Live site: <https://canispreadsheet.com/> &middot; Open dataset (CC BY 4.0): [`compat.csv`](https://canispreadsheet.com/data/compat.csv) / [`compat.json`](https://canispreadsheet.com/data/compat.json) &middot; [dataset page](https://canispreadsheet.com/data.html) &middot; archived on Zenodo with a citable DOI.

The entire value of this project is that every number in it was actually
computed by the engine in question, not inferred from documentation. Phase 1
builds the pipeline and proves out the hardest part of that promise: getting
LibreOffice to genuinely recalculate headless, with proof.

## Architecture

```
data/functions.json        Function inventory: every function name, its
                            category, and whether/where Excel, Google
                            Sheets, and LibreOffice Calc document it
                            (with source URLs).

data/tests/<FUNCTION>.json One file per function under test. Each file is a
                            list of test cases: {id, formula, setup_cells,
                            description, expected, expected_note,
                            check_range}. Formulas are written exactly as a
                            human would type them in the Excel UI — no
                            engine-specific storage quirks belong here.

harness/xlfn_map.py         Translates modern function names to the
                            "_xlfn."/"_xlfn._xlws." prefixed form Excel
                            requires inside the raw .xlsx XML (see below).

harness/corpus.py           Engine-agnostic shared machinery both runners
                            import: test loading, sheet-name sanitizing,
                            workbook building, read-back normalization and
                            expected-value comparison. Lives in one place so
                            an LO-vs-Sheets difference can never be a
                            harness artifact.

harness/run_lo.py           Engine runner for LibreOffice Calc: builds one
                            .xlsx from data/tests/*.json, forces a real
                            LibreOffice recalculation, reads back computed
                            values, writes results/libreoffice-24.2.json.

results/<engine>-<ver>.json Output of each engine runner: real, executed
                            values per test id, plus a canary block proving
                            recalculation actually happened.

scripts/gen_test_cases.py   Generator that produced the current
                            data/tests/*.json files (kept for reference /
                            as a pattern to follow when adding more
                            functions in bulk; hand-editing the JSON
                            directly is equally fine going forward).

harness/run_sheets.py       Engine runner for Google Sheets: builds chunked
                            .xlsx workbooks for Drive auto-conversion,
                            ingests the exported .xlsx readback, writes
                            results/google-sheets.json. See "Phase 2:
                            Google Sheets runner" below. Its
                            build-recipes / ingest-recipes / selftest-recipes
                            subcommands do the same for the RECIPE corpus.

data/recipes/<slug>.json   One file per how-to recipe: the task, the
                            per-app formulas, the explanation, a "verify"
                            block {setup_cells, setup_sheets?, formula,
                            expected, check_range?}, optional variants[]
                            carrying their own checks, and optional
                            extra_checks[] (checks belonging to the recipe
                            but to no variant). ANY check may carry
                            "engines": ["google_sheets"] / ["libreoffice"]
                            to scope it to one engine (absent = all) and an
                            "id" to pin its stable key. A second corpus,
                            independent of data/tests/.

harness/recipe_corpus.py    Engine-agnostic shared machinery for the RECIPE
                            corpus, the counterpart of harness/corpus.py:
                            recipe loading, check enumeration, variant setup
                            inheritance, read-back norm() and the
                            expected-vs-actual rule. Imported by BOTH
                            scripts/verify_recipes.py (LibreOffice) and
                            harness/run_sheets.py (Google Sheets).

scripts/verify_recipes.py   Engine runner for the RECIPE corpus in headless
                            LibreOffice -> results/recipes-verified.json.
                            The Google Sheets counterpart writes
                            results/recipes-verified-sheets.json.

(not built yet)
harness/run_excel.py         Excel engine runner (see Phase-2 notes below).
site/                        Static site generator consuming data/ + results/
                              -> deployed to GitHub Pages.
```

**Pipeline**: inventory (what exists) → tests (what behavior to check) →
engine runners (what actually happens) → results (raw truth) → site
(presentation layer, Phase 2+).

## Status (Phase 1 + Phase-2 corpus expansion)

- **Google Sheets is now EXECUTED.** `results/google-sheets.json` (2026-08-29)
  holds the whole corpus — 278 functions / 841 cases — run through the Drive
  import route (upload .xlsx → Sheets recalculation → .xlsx export readback),
  canary 841/841, `trusted: true`. Two engines are executed now: LibreOffice
  Calc (four pinned builds) and Google Sheets (one dated run — Sheets has no
  version to pin, so its `engine_version` is a DATE LABEL and must never be
  parsed as a version). **Microsoft Excel is still not executed** and its
  column remains documentation-only everywhere on the site.
  Five functions (BYCOL, BYROW, FILTER, MAKEARRAY, SORT) carry the verdict
  `inconclusive` because Google's importer did not map the OOXML storage
  prefix our Excel-authored workbooks use for them — see "Google Sheets
  execution caveats" in `DATASET_CARD.md`. They are NOT reported as
  unsupported.

- `data/functions.json`: 600 distinct function names inventoried from live
  official docs. Excel: 522 documented, Google Sheets: 515, LibreOffice: 469,
  documented in all three: 421. Sources actually fetched are listed in the
  `sources` array; one attempted source
  (`wiki.documentfoundation.org/List_of_Calc_Functions`) was blocked by an
  anti-bot wall and is recorded honestly as `fetched: false` — LibreOffice
  coverage instead comes from `help.libreoffice.org`'s category pages (18
  pages fetched), which gave full, real coverage anyway.
- `data/tests/`: 148 functions, 604 hand-authored test cases. Phase 1
  covered 31 functions (125 cases); the Phase-2 batch added 117 more
  workhorse/compat-interesting functions (479 cases) spanning math
  (CEILING/FLOOR + .MATH variants, MROUND, INT-vs-TRUNC...), statistics
  (STDEV/RANK/PERCENTILE families), text (TEXT format codes, FIND/SEARCH,
  TRIM/CLEAN, CHAR/CODE/UNICHAR/UNICODE...), date/time (WEEKDAY return
  types, YEARFRAC bases, WEEKNUM vs ISOWEEKNUM, DAYS360 US/EU...),
  lookup/reference (SUMIF/COUNTIF families, INDIRECT, OFFSET, HLOOKUP,
  LOOKUP, TRANSPOSE...), and information/logical (IS* family, TYPE,
  ERROR.TYPE, XOR...). Edge-case expectations cite official
  Microsoft/LibreOffice/Google doc URLs inline in each case's
  description/expected_note field.
- `harness/run_lo.py` executed against LibreOffice 24.2.7.2, results
  committed at `results/libreoffice-24.2.json`. **Recalculation is proven
  genuine** — see "How the LO runner forces recalculation" below. Of the
  604 cases: 513 matched their documented expectation, 86 diverged
  (preserved as `matched_expected: false` — divergences are the product,
  never "fixed" to match the engine), and 5 are intentionally
  non-deterministic existence probes (TODAY/RAND family). See "Phase-2
  headline quirks" below for the most interesting new divergences.

## How the LO runner forces recalculation (and how we know it's real)

The single biggest credibility risk for this whole project is silently
reporting *cached* or *stale* values as if they were freshly computed. Two
independent facts make `soffice --headless --convert-to xlsx` trustworthy
here:

1. **openpyxl never writes a cached `<v>` value for a formula cell** — only
   the formula string. There is nothing for LibreOffice to "fall back to."
   For a formula cell to show ANY value at all after the round trip,
   LibreOffice must have evaluated it from scratch.
2. We verify this on every run with two canaries written to every sheet:
   - `=1111+2222` (deterministic, arithmetic, impossible to have a
     pre-existing cached value) — must read back as exactly `3333`.
   - `=NOW()` on a dedicated `_meta` sheet — the file is converted twice,
     ~2 seconds apart, and the two `NOW()` values must differ. If
     LibreOffice were just echoing something static, they'd be identical.

   Both checks are in `results/libreoffice-24.2.json` under `"canary"`, and
   the run sets a top-level `"trusted": true/false` flag. **If `trusted` is
   ever `false`, treat every value in that results file as unverified.**
   From our Phase-1 run:
   ```json
   "canary": {
     "arithmetic_actual": 3333,
     "arithmetic_ok": true,
     "now_run_1": "2026-07-04 02:16:38.426000",
     "now_run_2": "2026-07-04 02:16:45.086000",
     "now_differs_across_runs": true
   }
   ```
   Every individual test sheet also carries its own copy of the arithmetic
   canary (`canary_ok_this_sheet`), so a single corrupted/unrecalculated
   sheet couldn't hide behind a passing global check.

## The `_xlfn.` / `_xlfn._xlws.` prefix gotcha (read this before adding tests)

The OOXML (.xlsx) file format froze its formula function list at Excel 2007.
Every function added since then (XLOOKUP, LET, LAMBDA, FILTER, SORT, UNIQUE,
SEQUENCE, TEXTSPLIT, TEXTBEFORE/AFTER, IFS, SWITCH, MAXIFS/MINIFS, TEXTJOIN,
CONCAT, IFNA, ARRAYTOTEXT, ...) has to be serialized into the raw XML with an
`_xlfn.` prefix (or the double `_xlfn._xlws.` prefix, for just `FILTER` and
`SORT`). Real Excel does this silently when it saves a file, and strips it
back off when displaying the formula bar. Libraries that write raw XML
(openpyxl, xlsxwriter) do **not** do this for you.

If you write `=XLOOKUP(...)` into an .xlsx with openpyxl and open it in
*any* engine — including real Excel — you get `#NAME?`, even though XLOOKUP
is fully supported. This is not a compatibility finding, it's an
openpyxl/xlsxwriter footgun, and getting it wrong would silently corrupt
every "unsupported" verdict in this database.

`harness/xlfn_map.py` handles this centrally: test-case JSON always stores
the natural, human-typed Excel formula; the engine runner translates it to
the correct storage form right before writing the .xlsx, based on a
`_XLFN_FUNCTIONS` / `_XLWS_FUNCTIONS` table sourced from XlsxWriter's public
"Working with Formulas" documentation (the de facto reference for this
quirk). We double-checked the *absence* of support for functions like LET
and XLOOKUP in LibreOffice 24.2 independently of this prefix question, by
driving LibreOffice's own native formula parser over PyUNO
(`createInstanceWithContext` + `Desktop.loadComponentFromURL` +
`cell.setFormula(...)`) — LO's own parser silently lower-cases and fails to
recognize `LET`, `XLOOKUP`, `FILTER`, `SORT`, `UNIQUE`, `SEQUENCE`,
`LAMBDA`, `TEXTBEFORE`, `TEXTAFTER`, `TEXTSPLIT`, and `ARRAYTOTEXT` as
function names at all, regardless of prefix — confirming these are genuine
support gaps in this LibreOffice version, not artifacts of our test
harness.

## Dynamic-array / spill results and legacy array-formula entry

A formula that's supposed to return a multi-cell array (e.g.
`INDEX(range,0,col)`, `FILTER`, `SORT`, `UNIQUE`, `SEQUENCE`) needs to be
written as a legacy Ctrl+Shift+Enter–style array formula
(`openpyxl.worksheet.formula.ArrayFormula`, with `ref` set to the full
output range) to spill correctly under LibreOffice's compatibility model.
We proved this empirically: `INDEX(A1:C3,0,2)` written as a **plain**
formula string returns `#VALUE!` in LibreOffice 24.2, but the identical
formula wrapped in `ArrayFormula(ref="A30:A32", ...)` correctly spills
`[2, 5, 8]`. `harness/run_lo.py` automatically wraps any test case that
declares a `check_range` in `ArrayFormula`. Test cases that only need a
single output cell don't set `check_range` and are written as plain
formulas.

## Phase 1 result summary — LibreOffice Calc 24.2.7.2

Of the 31 functions tested (125 total cases):

| Verdict | Functions |
|---|---|
| **Unsupported** (`#NAME?` / not recognized by LO's own parser) | XLOOKUP, LET, LAMBDA, FILTER, SORT, UNIQUE, SEQUENCE, TEXTSPLIT, TEXTBEFORE, TEXTAFTER, ARRAYTOTEXT (11) |
| **Supported, behaves as documented** | IFS, SWITCH, MAXIFS, MINIFS, SUMIFS, COUNTIFS, IFERROR, IFNA, MATCH, TEXTJOIN, CONCAT, ROUND, MOD, EDATE, NETWORKDAYS, RAND, RANDBETWEEN (17) |
| **Supported, with a discovered quirk vs. documented Excel behavior** | VLOOKUP, INDEX, DATEDIF (3) |

**Quirks found (exact formula → exact result):**

1. `=VLOOKUP("a",A1:B3,5,FALSE)` with a 2-column table → LibreOffice returns
   `#VALUE!`. Microsoft's docs say an out-of-range `col_index_num` should
   return `#REF!`. Real cross-engine divergence in error *code* (both agree
   it's an error).
2. `=DATEDIF(DATE(2024,1,10),DATE(2024,1,1),"D")` (end date before start
   date) → LibreOffice returns `#VALUE!`. Microsoft's docs say this should
   return `#NUM!`. Same pattern as above: both error, different code.
3. `=LAMBDA(x,x*2)(5)` (immediately-invoked LAMBDA) → LibreOffice returns
   `#VALUE!`, whereas `=LET(f,LAMBDA(x,x^2),f(4))` (LAMBDA bound via LET)
   returns `#NAME?`. LAMBDA is unsupported either way, but the exact error
   surfaced depends on the call shape — useful detail for anyone trying to
   distinguish "unsupported" from "syntax I got wrong" by error code alone.
4. `INDEX(range,0,col)` (whole-column spill) only spills correctly when
   entered as a legacy CSE array formula — see the array-formula section
   above. Not a bug, but a real trap for anyone building `.xlsx` files
   programmatically and expecting Excel-365-style implicit spilling.
5. `ROUND(1.005,2)` → LibreOffice correctly returns `1.01`, *not* the naive
   `1.0` a binary-float implementation would produce (the true IEEE-754
   double nearest 1.005 is ≈1.00499999999999989). Recorded as a passing
   "verified correct" case, not a quirk, but worth highlighting since this
   is exactly the kind of subtle numerical-fidelity question this database
   exists to answer.
6. `MOD(-7,3)` → `2` and `MOD(7,-3)` → `-2`: confirms LibreOffice follows
   the spreadsheet convention (result takes the sign of the divisor,
   `n - d*FLOOR(n/d)`) rather than C-style truncated-remainder semantics.
   Verified correct, not a quirk, but a common source of porting bugs.

All 45 `#NAME?` results plus the `#VALUE!`/other-error cases above are
recorded with matched engine version, formula (display + literal .xlsx
storage form), and full notes in `results/libreoffice-24.2.json`.

## Phase-2 headline quirks — LibreOffice Calc 24.2.7.2 vs documented Excel behavior

The full list is every `matched_expected: false` entry in
`results/libreoffice-24.2.json` (86 of them) and on the generated
`docs/quirks.html` page; these are the most interesting families:

1. **Booleans are numbers in LO.** `=ISNUMBER(TRUE)` → `TRUE` (Excel docs:
   FALSE — booleans are their own type); `=TYPE(TRUE)` → `1` (Excel: 4);
   `=COUNT()` over a range containing a boolean cell counts it (Excel:
   excluded). One consistent LO design decision that flips three functions'
   documented results.
2. **Error-CODE divergence family: LO surfaces `#VALUE!` where Microsoft
   documents `#NUM!` (or `#REF!`).** Reproduced across SQRT(-16), LN(0)/
   LN(-5), LOG(0)/LOG(-10), LOG10(0)/LOG10(-5), SMALL/LARGE with k out of
   range, PERCENTILE.INC/EXC and QUARTILE.INC out-of-range k/quart,
   WEEKDAY invalid return_type, YEARFRAC invalid basis, FLOOR mismatched
   signs, MODE with no duplicate (`#VALUE!` instead of documented `#N/A`),
   OFFSET off-sheet and HLOOKUP row_index out of range (`#VALUE!` instead
   of `#REF!`). Both engines agree these are errors; the *code* differs,
   which breaks error-code-sniffing formulas ported from Excel. (LO's
   internal Err:502 "invalid argument" maps to `#VALUE!` on xlsx export.)
3. **`=MROUND(5,-2)` → `6`.** Microsoft documents mixed-sign arguments as a
   hard `#NUM!` error; LibreOffice happily computes a value instead. A
   ported sheet relying on that error will silently produce numbers.
4. **`=POWER(-8,1/3)` → `-2`.** Excel documents/returns `#NUM!` for any
   negative base with non-integer exponent; LO computes the real odd root.
5. **Serial-number epoch offset below March 1900.** `=YEAR(1)` → `1899`,
   `=MONTH(1)` → `12`, `=DAY(1)` → `31`: LO maps serial 1 to Dec 31 1899,
   Excel maps it to Jan 1 1900 (a knock-on of Excel's fictitious
   Feb 29 1900). All dates from Mar 1 1900 onward agree.
6. **`=CHAR(0)` and `=UNICHAR(0)` return a NUL character** (stored in xlsx
   as the `_x0000_` escape) instead of Microsoft's documented `#VALUE!`.
7. **`=SUM(1,"2",3)` → `#VALUE!`.** Excel documents that numeric-looking
   text *literals* typed directly as arguments are coerced (result 6); LO
   refuses and errors even for direct literals.
8. **`=TRIM(CHAR(160)&"Hello"&CHAR(160))`**: LO's `CHAR(160)` does not
   produce a non-breaking space at all — the round-tripped result contains
   U+FFFD replacement characters, so the classic "TRIM doesn't strip
   nbsp" Excel behavior can't even be expressed the same way in LO.
9. **`=ERROR.TYPE(...)` on LO-internal errors**: for `OFFSET(A1,-1,0)` and
   `SQRT(-1)` inputs LO returns `#N/A` rather than the documented codes 4
   and 6 — consistent with quirk family 2 (the inner errors aren't the
   error codes Excel would produce, and LO's ERROR.TYPE doesn't map them).
10. **Where LO deliberately matches Excel:** the CEILING/FLOOR negative-
   number default-Mode divergence that LibreOffice's own documentation
   describes for ODF context does NOT appear via .xlsx —
   `=CEILING(-45.67,-2)` → `-46` and `=FLOOR(-45.67,-2)` → `-44`, exactly
   the Microsoft-documented defaults. LO applies Excel-compatible
   semantics when the formula arrives via an Excel-format file. Verified
   correct, not a quirk, but exactly the kind of context-dependent
   behavior this database exists to pin down.

Also verified as matching documentation (worth calling out because the
opposite is often assumed): `INT(-8.9)=-9` vs `TRUNC(-8.9)=-8`, XOR's
odd-count-of-TRUE rule, RANK.AVG tie-averaging (3.5), PERCENTILE.EXC's
exclusive k-bounds errors (as `#VALUE!`, see family 2), WEEKNUM
return_type 1 vs 2 divergence (10 vs 11 on Microsoft's own example date),
ISOWEEKNUM year-boundary behavior (Jan 1 2023 → ISO week 52 of 2022),
DAYS360 US-vs-European method (30 vs 29 on the same date pair), ISBLANK
FALSE on an `=""` formula result, COUNTBLANK counting that same cell as
blank, and TIME(27,0,0)=0.125 hour wrap-around.

## Phase 2: Google Sheets runner

`harness/run_sheets.py` executes the same corpus against **Google Sheets**.
There is no headless Sheets binary and the harness holds no Google
credentials, so the runner is split in half around one external step:

> Uploading a **formula-only** `.xlsx` to Google Drive with the
> Google-Sheets target MIME type auto-converts it into a real Google Sheet,
> and that conversion **recalculates every formula** with Google's own
> engine. Exporting that Sheet back out as `.xlsx` yields a workbook whose
> cells carry Google Sheets' computed cached values.

`build` emits the workbooks to upload; `ingest` reads the exported
workbooks back and writes `results/google-sheets.json` in **exactly** the
schema `results/libreoffice-*.json` uses. The Drive upload/export in the
middle is done by whatever has Drive access (the orchestrator).

### The orchestrator loop

```
# 1. Build the chunk workbooks + manifest (all 278 functions, ~40 per chunk)
python3 harness/run_sheets.py build
#    -> harness/sheets_chunks/chunk-01.xlsx ... chunk-07.xlsx
#    -> harness/sheets_chunks/manifest.json

# 2. For each chunk-NN.xlsx, via the Drive API / an MCP Drive tool:
#      a. upload it with contentMimeType = the .xlsx MIME type and the
#         Google-Sheets target MIME type, so Drive AUTO-CONVERTS it
#         (this is the step that recalculates every formula)
#      b. export that Sheet back out as .xlsx
#      c. save it as harness/sheets_exports/chunk-NN-export.xlsx
#         (the "chunk-NN" in the filename is how ingest identifies it)

# 3. Ingest — incremental, so run it per chunk as each export lands
python3 harness/run_sheets.py ingest \
    --export harness/sheets_exports/chunk-01-export.xlsx \
    --engine-label "Google Sheets (Drive import, 2026-08-29)"
#    -> results/google-sheets.json   (merged, never clobbered)
```

Use **xlsx** for the export, not CSV: a CSV export of a Google Sheet
returns **only the first sheet**, and every test case lives on its own
sheet.

### Why the workbooks are chunked

The `.xlsx` bytes travel through the orchestrator's context as base64
(~4/3 inflation), so the single 841-sheet ~430 KB workbook the LO runner
builds is not transportable. `build` splits the corpus **by function**
(never splitting one function's cases across two files, since results merge
per-function) into chunks of ~40 functions. Measured on the current
278-function corpus:

| chunk | functions | cases | bytes |
|-------|-----------|-------|-------|
| chunk-01 | 40 | 126 | 67,715 |
| chunk-02 | 40 | 107 | 59,763 |
| chunk-03 | 40 | 119 | 64,959 |
| chunk-04 | 40 | 137 | 73,798 |
| chunk-05 | 40 | 108 | 60,096 |
| chunk-06 | 40 | 123 | 67,277 |
| chunk-07 | 38 | 121 | 66,561 |
| **total** | **278** | **841** | **460,169** (~614 KB base64) |

(Byte counts vary by a byte or two between builds — zip metadata — so the
manifest records the actual `sha256` and `bytes` of the files it wrote.)

Every chunk is comfortably under the ~150 KB per-workbook budget. Sheet
names are the test ids, sanitized to ≤31 chars with `[ ] * ? / \ :`
stripped — valid OOXML *and* within Google Sheets' own rules (≤100 chars,
same forbidden characters), and `build` asserts this for every sheet
before writing.

### The canary, and how strong it actually is

Same pattern as the LO runner, because it rests on the same fact: openpyxl
**never** writes a cached `<v>` value for a formula cell, so the uploaded
chunk contains zero cached results.

- **Deterministic canary** — `=1111+2222` in `Z1` of *every* sheet. If
  Google had not recalculated, that cell would export back **blank**
  (`None`), because there was never a cached value to preserve. Reading
  back exactly `3333` proves Google computed it. `ingest` checks this on
  every sheet and reports the count.
- **Volatile canary** — `=NOW()` on the `_meta` sheet, recorded per chunk
  as corroboration.

**Honest limitation:** the LO runner converts the same file twice a few
seconds apart and shows `=NOW()` *differing between runs*. A single Drive
import gives one timestamp, so that cross-run comparison is not available
here — `canary.now_differs_across_runs` is `null` **by design**, and the
deterministic canary is the load-bearing proof. If any deterministic canary
fails, the whole run is `"trusted": false` and each affected case gets an
`UNTRUSTED_RECALC` note.

### The `_xlfn.` prefix: Sheets understands it (verified)

Google Sheets **does** honour the OOXML `_xlfn.` storage prefix on import —
verified empirically: `_xlfn.XLOOKUP(...)` imported and computed a real
value, it did *not* come back `#NAME?`. So the Sheets path uses
`harness/xlfn_map.py` **identically** to the LibreOffice path, with no
Sheets-specific special-casing. Writing the bare unprefixed name would be
the thing that breaks it.

### Honesty labelling — Sheets has no version

Google Sheets is a continuously-updated hosted product with no
user-visible version number and none exposed through Drive. So
`engine_version` in `results/google-sheets.json` is **not a version** — it
is a dated label recording *when* the corpus was executed:

```
"engine":          "google_sheets",
"engine_version":  "Google Sheets (Drive import, 2026-08-29)",
"recalc_method":   "Drive import + xlsx export readback",
```

A Sheets result means *"this is what Google Sheets did on that date"* and
nothing stronger. Anything presenting Sheets data must say so, and must not
imply a pinned version the way the LibreOffice columns legitimately can.
Merging an ingest into a file recorded under a **different** label is
refused unless `--allow-label-change` is passed (which records both in
`engine_version_history`), so two execution dates can never be silently
blended.

### Incremental ingestion merges, it does not overwrite

Every ingest is inherently a subset (some chunks out of N), so `ingest`
uses exactly the merge semantics `run_lo.py` uses for subset runs: only the
functions ingested this time are replaced, every other function's result is
preserved byte-for-byte, `generated_at` / `canary` / `recalc_method` are
refreshed to describe the most recent execution, `trusted` becomes the AND
of the previous flag and this run's (a merge can only downgrade trust), and
each merge appends to `subset_runs`. An export whose sheet list does not
match the manifest chunk it claims to be is **rejected**, so an
out-of-order upload can't be mapped against the wrong cells.

### Per-function execution dates (`executed_at`)

Because a subset run refreshes the file-level `generated_at`, that field
says when the **file** was last written, not when any given function was
executed. Every function block therefore carries its own date:

```json
"function_results": {
  "ABS":  { "executed_at": "2026-07-29", "ABS_positive_number": { ... } },
  "MIRR": { "executed_at": "2026-08-31", "MIRR_basic": { ... } }
}
```

Merges replace whole function blocks, so a re-executed function brings the
new date in and every untouched function keeps the date of the run that
produced it. Anything printing a per-function "last tested" date must read
`executed_at` and fall back to `generated_at` only for files written before
this key existed; `scripts/check_honesty.py` fails the build if a function
page's date does not match its `executed_at`. Consumers iterating a function
block's test cases must skip the metadata key — `harness/results_schema.py`
provides `function_cases()` for exactly that. The dates on existing files
were reconstructed from git history by
`scripts/backfill_executed_at.py` (one-off; its docstring documents how each
date was derived and what that derivation cannot know).

Error cells come back from Google's export as cached error strings
(`#NUM!`, `#N/A`, `#NAME?`, `#REF!`, `#DIV/0!`, `#VALUE!`) and are handled
exactly as the LO runner handles them, raw string kept verbatim. Sheets
adds one token Excel has no equivalent for: `#ERROR!`, its *parse-failure*
error — flagged with a note so it is never confused with `#NAME?`
(unknown function). Sheets has no `#CALC!`; it returns plain `#N/A` where
Excel 365 would say `#CALC!` (e.g. `FILTER` matching no rows).

### Proving the plumbing without Google

```
python3 harness/run_sheets.py selftest --only COUNT SUM MROUND XLOOKUP DATEDIF ISNUMBER
```

`selftest` builds the chunks, stands **LibreOffice** in for Drive
(`soffice --headless --convert-to xlsx`), and ingests the result as if it
were a Sheets export — proving the chunk → sheet → anchor-cell → test-id
mapping end-to-end. On the 6-function set it recovers values byte-identical
to `results/libreoffice-25.8.json` for all 28 cases.

The values it recovers are LibreOffice's, so two guards keep them out of the
published dataset: `site/build_site.py` discovers engines by globbing
`results/*.json` and matching the `engine` string, so selftest output (a)
**refuses to be written anywhere inside `results/`** — it goes to
`harness/sheets_selftest/plumbing-check.json` — and (b) carries
`"engine": "SELFTEST_libreoffice_via_sheets_pipeline"`, never
`"google_sheets"`. Either guard alone would prevent LibreOffice numbers
appearing in a Google Sheets column; both are enforced.

### Shared code: `harness/corpus.py`

Test loading, the sheet-name sanitizer, `build_workbook`, the anchor-cell
convention, read-back normalization and `compare_expected` live in
`harness/corpus.py`, imported by **both** runners. If the two runners
drifted apart even slightly, a reported LibreOffice-vs-Sheets "difference"
could be a harness artifact rather than a real engine difference. Each
runner owns only the engine-specific part: making the engine recalculate,
and proving that it did.

## Phase 2b: Google Sheets runner for the how-to RECIPE corpus

The 282 how-to recipes in `data/recipes/*.json` are a **second corpus**,
separate from the function corpus in `data/tests/*.json`. They have been
executed in headless LibreOffice since day one
(`scripts/verify_recipes.py` -> `results/recipes-verified.json`, which is
what the "Verified, not just documented" block on every how-to page
renders). `run_sheets.py` now gives them the same Drive-import treatment:

```
# 1. Build the chunk workbooks + manifest (all recipes, 60 per chunk)
python3 harness/run_sheets.py build-recipes
python3 harness/run_sheets.py build-recipes --chunk-size 40
python3 harness/run_sheets.py build-recipes --only how-to-use-xlookup add-days-to-a-date
python3 harness/run_sheets.py build-recipes --outdir harness/recipe_chunks
#    -> harness/recipe_chunks/chunk-01.xlsx ... chunk-05.xlsx
#    -> harness/recipe_chunks/manifest.json

# 2. Same Drive step as the function corpus: upload each chunk with the
#    Google-Sheets target MIME type so Drive auto-converts (and recalculates),
#    then export the Sheet back as .xlsx into
#    harness/recipe_exports/chunk-NN-export.xlsx

# 3. Ingest — incremental, run it per chunk as each export lands
python3 harness/run_sheets.py ingest-recipes \
    --export harness/recipe_exports/chunk-01-export.xlsx \
    --out results/recipes-verified-sheets.json

# Dry run of the whole loop with LibreOffice standing in for Drive. Every
# value must come back equal to results/recipes-verified.json or it exits 1.
python3 harness/run_sheets.py selftest-recipes
```

**One worksheet per CHECK.** A recipe contributes its main worked example
plus one check per `variants[].verify` entry, so 276 recipes become 291
sheets. Sheet names are `r<index>_<slug>` / `r<index>v<n>c<n>_<slug>`
truncated to 31 chars — the index prefix makes them unique and stable (a
`--only` build puts a check on the same sheet a full build would), the slug
tail is so a human opening the workbook in Drive can tell what they are
looking at. Each sheet carries that check's `setup_cells`, the formula at
the same anchor `verify_recipes.py` uses (`H1`, or the top-left cell of
`check_range` written as a real array formula), and the deterministic
`=1111+2222` canary in `Z1`. `=NOW()` sits on the `_meta` sheet, exactly as
for the function corpus.

**Plain function names by default.** Unlike the function corpus, recipe
chunks are written with formulas **exactly as authored** (`--plain-names`,
on unless you pass `--xlfn-names`). Google Sheets' xlsx importer maps bare
modern names but not the `_xlfn._xlws.FILTER/SORT` storage form, and a
recipe is by definition the formula a user types into the app. The cost is
recorded, not hidden: any check whose stored bytes differ from what the
LibreOffice reference run executed is flagged
`differs_from_lo_serialization` in the manifest, carries a note in the
results file, and is reported as **not comparable** (rather than as a pass
or a failure) by `selftest-recipes`.

**Engine-scoped checks, and results merged by key.** A check may name the
engines that should execute it:

```json
{ "id": "gs-alt",
  "label": "SUMPRODUCT does the whole OR in one pass",
  "engines": ["google_sheets"],
  "formula": "=SUMPRODUCT((A2:A6=\"East\")+(A2:A6=\"West\"),B2:B6)",
  "expected": 160 }
```

Absent (or empty) `engines` means **all** engines, so every check authored
before the field existed is untouched by it. `verify_recipes.py` runs as
`libreoffice` and `build-recipes` runs as `google_sheets` (override with
`--engine`; `selftest-recipes` forces `libreoffice`, because there
LibreOffice really is the engine standing in for Drive), and each filters
the corpus through `iter_checks(recipe, engine=...)`. That is what makes it
safe to publish a Sheets-only alternative formula: LibreOffice would run one
anyway — sometimes computing a different answer, sometimes failing to parse
it at all — and either outcome would drag the recipe's LibreOffice badge
down with it. A recipe's `verified` flag is therefore the AND over **that
engine's** checks only.

Every check also has a **stable key** — `main`, `v<vi>c<ci>`, `x<i>`, or an
explicit `"id"` — computed from the check's position in the *unfiltered*
JSON, so a filtered build keys its results exactly as a full build would.
Both results files store the key on every payload and every consumer merges
**by key, not by position**, so appending an engine-scoped check to a
variant cannot slide an older stored value onto a neighbouring formula.
Results files written before keys existed still load: their keys are derived
positionally by the same rule (`recipe_corpus.result_checks_by_key`).

On the site, a Sheets-only check renders in the Google Sheets column only —
labelled "Google Sheets alternative (executed `<date>`)" beside the value
Google returned — while the LibreOffice column reads `n/a (Sheets-only
formula)`, never a value and never a badge. Per-engine counts follow ("the N
further formulas ... executed" is counted separately for each engine). A
scoped check whose engine has not executed it yet is **hidden**, not
rendered as pending, so no page ever carries copy that goes stale the day
the run lands. `scripts/check_honesty.py` enforces the pairing: the number of
"Google Sheets alternative (executed …)" labels on a page must equal the
number of "n/a (Sheets-only formula)" cells, and such a label may only
appear on a page that carries real Sheets execution provenance.

Measured on the current 282-recipe corpus (`--chunk-size 60`):

| chunk | recipes | checks | bytes |
|-------|---------|--------|-------|
| chunk-01 | 60 | 60 | 36,734 |
| chunk-02 | 60 | 60 | 36,986 |
| chunk-03 | 60 | 60 | 37,554 |
| chunk-04 | 60 | 60 | 37,503 |
| chunk-05 | 36 | 51 | 34,885 |
| **total** | **276** | **291** | **183,662** (~245 KB base64) |

Every chunk is roughly a quarter of the ~150 KB per-workbook budget.

### Multi-sheet recipes are skipped, and listed

Six recipes need extra worksheets to exist (`setup_sheets`). `build-recipes`
does **not** build them; it lists them in the manifest under
`skipped_multi_sheet` (with the tab names each one wants) and prints them at
the end of the build. They stay LibreOffice-executed only.

That is the *simpler correct* option here, not a shortcut. The obvious
alternative — give every check its own copy of the tabs, prefixed to avoid
collisions, and rewrite the formula's sheet references — silently breaks the
very things these recipes test:

- `=INDIRECT(A1&"!B2")` builds the reference out of a **cell value**, so
  renaming a tab changes the answer unless the setup data is rewritten too —
  and one check exists specifically to assert `#REF!` for a tab that does
  *not* exist.
- `=SUM(Q1:Q3!A1)` is a 3-D reference over a **span of consecutive tabs**;
  it depends on sheet order as well as on names.
- `=$'Q1 Data'.B2` and `=Data.B2` are there to assert `#NAME?` — the
  LibreOffice-syntax forms Excel and Sheets reject. A rewriter that "fixed"
  those references would destroy the test.
- Tab names collide **within a single recipe**: one variant defines
  `Q1`/`Q2`/`Q3` as numbers and a check inside that same variant redefines
  them with different contents, so even one-workbook-per-recipe does not
  de-collide them.

A future version that uploads one workbook per **check** would cover all six
honestly with no rewriting at all, at the cost of ~34 more Drive round-trips.

### Shared code: `harness/recipe_corpus.py`

Which checks exist, how a variant check inherits `setup_cells` /
`setup_sheets`, the read-back `norm()`, and the expected-vs-actual rule live
in `harness/recipe_corpus.py`, imported by **both** `verify_recipes.py` and
`run_sheets.py` — the recipe-corpus counterpart to `harness/corpus.py`. The
site prints "LibreOffice returned X, Google Sheets returned Y" side by side,
so any drift in enumeration or comparison between the two paths would
manufacture a fake engine divergence. The move out of `verify_recipes.py`
was behaviour-preserving: a full 282-recipe / 325-check LibreOffice re-run
after the refactor reproduced `results/recipes-verified.json` byte for byte.
The later engine-scoping + stable-key change was proved the same way, one
notch looser: a full re-run reproduced every `verified` flag, every `actual`
and every `expected` in that file, the only difference being the `"key"`
field newly written onto each check.

### How the site renders a Sheets recipe result

`results/recipes-verified-sheets.json` is optional. Without it the how-to
pages are byte-identical to before. With it, and only when the file's
`trusted` flag is set (all per-sheet canaries read back 3333):

- each recipe that has a Sheets result gains a second executed column,
  **"Returned by Google Sheets (executed `<date>`)"**, beside the
  LibreOffice one, plus a header badge;
- a value Sheets returned that differs from LibreOffice's is shown **as
  Sheets returned it**, flagged `differs from LibreOffice`, and called out
  in prose — the disagreement is the interesting content, not an error to
  hide;
- that recipe's meta description becomes *"Recipe executed in LibreOffice
  Calc and Google Sheets; Excel per docs"*. A recipe **without** a Sheets
  result keeps the LibreOffice-only wording, so the skipped multi-sheet
  recipes never over-claim;
- the how-to index lede and the methodology page state how many recipes have
  been through Sheets, and why the rest have not;
- a check scoped to Google Sheets alone (`"engines": ["google_sheets"]`)
  renders in that column only — labelled **"Google Sheets alternative
  (executed `<date>`)"** beside the value Google returned — while the
  LibreOffice column reads `n/a (Sheets-only formula)`. It is hidden
  entirely until Sheets has actually executed it, and the "N further
  formulas executed" sentences are counted per engine.

`scripts/check_honesty.py` guards the new wording: it fails on an engine
list that ends "... and Excel" after an execution verb (the regression the
two-engine phrasing newly makes possible), on any single page that both
shows a Sheets result and still carries the LibreOffice-only disclaimer, and
on a Sheets-only row whose LibreOffice column shows anything other than
`n/a (Sheets-only formula)` (or that claims an alternative was executed on a
page with no Sheets provenance at all).

## Offline companion: per-cell recalc diff

`scripts/xlsx_recalc_diff.py` answers a narrower, sharper question than the
web audit does: **exactly which cells in *my* workbook will compute
differently in LibreOffice Calc than they did in Excel?**

```
python3 scripts/xlsx_recalc_diff.py BOOK.xlsx [--json out.json] [--md out.md]
                                    [--limit N] [--sheet NAME] [--quiet]
                                    [--include-volatile] [--keep-temp]
```

Exit codes: `0` every cell matches · `1` differences found · `2` untrusted
run, unusable input, or error. Dependencies: `openpyxl` and `soffice` on
`PATH` (or `$SOFFICE_BIN`). Single file, no network, nothing uploaded, your
workbook is never modified.

### Why this works

An `.xlsx` saved by Excel stores, for every formula cell, both the formula
(`<f>`) and **the value Excel last computed for it** (`<v>`, the cached
value). That cached value is real, executed Excel ground truth already
sitting on disk. Diff it against a forced LibreOffice recalculation of the
same formulas and you get per-cell truth with no Excel install and no
upload — which is why this is the offline companion to
[canispreadsheet.com/audit.html](https://canispreadsheet.com/audit.html),
whose client-side audit only reaches *function-level* verdicts.

### The trap it works around: LibreOffice may not recalculate

LibreOffice's "Recalculation on File Load" for Excel files does **not**
default to *always*. Run `soffice --convert-to xlsx` on an Excel-saved file
and it can copy Excel's cached values straight through — producing a
beautiful, totally fake, 100%-match report. This is not hypothetical; it is
what happens, and `scripts/test_xlsx_recalc_diff.py::test_stripping_is_necessary`
asserts it with a live LibreOffice run:

```
[evidence] unstripped conversion -> 1 ; stripped conversion -> 2
           (Excel's cached value was 1)
```

So the tool builds a **stripped copy** at the XML level: inside every
`xl/worksheets/sheet*.xml`, each `<v>…</v>` under a cell that has an `<f>`
is deleted, along with the now-meaningless `t="str"/"e"/"b"` result-type
attribute. Every other byte of every other part is copied verbatim — a
round trip through openpyxl would silently drop charts, drawings, pivot
caches and most styles. Cells inside a legacy array formula's `ref` range
are stripped too, since they hold results without carrying an `<f>`
themselves. The stripped copy lives in a temp dir and is converted with
`soffice --headless --convert-to xlsx` under an isolated
`-env:UserInstallation` profile. **Nothing is ever injected into your
file.**

Trust is then verified, not assumed: the report states what fraction of the
formula cells that had an Excel cached value came back from LibreOffice
with a value. If that fraction is ~0, LibreOffice evaluated nothing and the
entire run is marked `UNTRUSTED` (exit 2). Never trust a clean report
without the "Recalculation check … TRUSTED" line.

### What it reports

Console summary + optional `--json` / `--md`: per-category counts, the
functions appearing most often in differing cells, and the first N differing
cells with `Sheet!Addr`, formula, Excel value, LibreOffice value and
category. Categories are `match`, `volatile`, `numeric_mismatch`,
`text_mismatch`, `type_mismatch`, `value_vs_error`, `error_vs_value`,
`error_vs_error`, `missing_in_lo`, `no_excel_value`.

Comparison rules: numbers with a 1e-9 relative tolerance; dates/times
normalized to Excel serials on both sides; booleans compared as booleans
(a boolean-vs-number pair is `type_mismatch`, not a numeric one); strings
compared exactly after decoding OOXML `_xNNNN_` escapes (so LibreOffice's
`_x0000_` for `CHAR(0)` compares as a real NUL); error cells compared by
code, so `#NUM!` → `#VALUE!` surfaces as `error_vs_error` rather than
"an error either way". A formula returning `""` round-trips as a value-less
cell, so blank and empty-string are treated as equal — they are genuinely
indistinguishable at this layer.

Shared formulas are handled properly: the set of formula cells is built
from the raw XML, **not** from openpyxl's view, because openpyxl exposes the
formula text only on a shared block's master cell. Slaves are diffed too and
labelled with their master's address.

### Honest limits

- **Only LibreOffice is executed by this offline tool.** (The site's function
  verdicts are executed in Google Sheets too; this per-cell recalc diff is
  LibreOffice-only.) The Excel side is whatever Excel cached in the file. If the workbook was last saved by something other than
  Excel, you are not diffing against Excel.
- **`calcMode="manual"`** in `xl/workbook.xml` means Excel was not
  recalculating automatically, so those cached values may be stale. The tool
  warns; fix it by opening in Excel, pressing F9, and saving.
- **No cached values at all** (files written by openpyxl / xlsxwriter /
  pandas, or saved with calculation disabled) means there is nothing to diff
  against. The tool says the diff is meaningless and exits 2 rather than
  inventing a comparison. `site/audit-page/verdict-mix.xlsx` is exactly this
  case.
- **Volatile functions** (`NOW`, `TODAY`, `RAND`, `RANDBETWEEN`,
  `RANDARRAY`) always differ by definition. They get their own `volatile`
  category and do not count as mismatches; `--include-volatile` overrides
  that.
- A difference is a difference, not necessarily a *bug*: it can be a genuine
  engine divergence (see `docs/quirks.html`), an unsupported function, or a
  stale Excel cache. The tool tells you which cells to look at; it does not
  adjudicate.
- Cell parsing is regex-based over well-formed OOXML cell elements. It is
  deliberate (byte-preserving) but assumes cell payloads are not exotic.

### Tests

```
python3 scripts/test_xlsx_recalc_diff.py -v     # 23 tests, runs real soffice
```

Harness plumbing has its own end-to-end dry runs, both of which stand
LibreOffice in for the Drive round-trip and write to scratch files outside
`results/`:

```
python3 harness/run_sheets.py selftest --only COUNT SUM MROUND
python3 harness/run_sheets.py selftest-recipes   # exits 1 on any value that
                                                 # does not match
                                                 # results/recipes-verified.json
```

Since we cannot run Excel, the "Excel-saved" side of each fixture is
simulated honestly: build the workbook with openpyxl, then **inject** a
cached `<v>` at the XML level exactly where Excel would have written its own
result, using the value Microsoft documents. Covered: a matching value, a
diverging value (`=COUNT(A1,1)` with `A1=TRUE` — Excel documents 1 because
booleans reached through a *reference* are not counted; LibreOffice really
computes **2**, consistent with `COUNT_boolean_in_range_excluded` in
`results/libreoffice-25.8.json`), error-vs-value (`=CHAR(0)`: Excel `#VALUE!`
vs LibreOffice's real NUL character), a shared-formula range, the
manual-calc warning, the no-cached-values exit-2 path (including this repo's
own `verdict-mix.xlsx`), volatile classification, and a check that stripping
preserves every non-formula cell and every non-sheet zip part byte-for-byte.

## How to add a function

1. Add/verify the function's entry in `data/functions.json` (name,
   category, per-app `documented`/`url`). Only add a URL you actually
   fetched — never fabricate.
2. Create `data/tests/<FUNCTION>.json`:
   ```json
   {
     "function": "SOMEFUNC",
     "cases": [
       {
         "id": "SOMEFUNC_basic",
         "formula": "=SOMEFUNC(A1,B1)",
         "setup_cells": {"A1": 1, "B1": 2},
         "description": "What this case checks and why",
         "expected": 3,
         "expected_note": "optional: why this is the expected value, especially for edge cases"
       }
     ]
   }
   ```
   - Write formulas exactly as typed in the Excel UI — never add `_xlfn.`
     prefixes yourself; the runner does that.
   - If a case's result is a multi-cell spill/array, add
     `"check_range": "A30:C31"` (pick unused rows ≥20 or so to avoid
     colliding with setup_cells) and put the array-shaped value in
     `expected` (nested lists for 2-D).
   - Cover at least: normal use, an edge case (empty/blank input, wrong
     type, negative numbers), and — for anything with documented
     "special" error behavior — a case that provokes that error.
   - If a function is new to `harness/xlfn_map.py`'s target set (i.e., it
     was added to Excel after 2007), add it to `_XLFN_FUNCTIONS` or
     `_XLWS_FUNCTIONS` there. Check XlsxWriter's "Working with Formulas"
     docs if you're not sure which.
3. Run `python3 harness/run_lo.py SOMEFUNC` to test just that function (or
   omit args to run everything), and check `results/libreoffice-24.2.json`
   for the outcome and `"trusted"` flag.
4. Commit `data/tests/SOMEFUNC.json` and the updated results file together.

## Phase 2 notes (not built yet)

- **Google Sheets engine.** ✅ Built — see "Phase 2: Google Sheets runner"
  above. It takes the Drive-import route (upload a formula-only .xlsx,
  Drive auto-converts and recalculates, export back as .xlsx) rather than
  the Sheets API, which needs no service account and no per-cell API
  traffic. ✅ The full 7-chunk sweep ran on 2026-08-29 and
  `results/google-sheets.json` is wired into the site (verdicts, guides,
  `compat.json` `gv`/`gver`, checker, Migration Audit). Remaining: a re-run
  writing plain function names to resolve the five `inconclusive` verdicts.
- **Excel engine.** No good headless Linux path exists (no real Excel
  calculation engine on Linux). Two options, in order of preference:
  1. **Windows + Office Scripts / VBA automation**: a small Windows runner
     (real Windows VM, or Office Scripts via Excel Online/Power Automate)
     that opens the generated .xlsx, forces
     `Application.CalculateFullRebuild`, and reads back values — this is
     the only way to get *real, executed* Excel ground truth. This should
     be the priority for Phase 2 rather than substituting documentation.
  2. Until (1) exists, Excel's column in the compat matrix should be
     populated **only** from `data/functions.json`'s documented-existence
     data and Microsoft's published documented behavior (clearly labeled
     "per Microsoft documentation, not independently executed" — never
     presented with the same confidence as an executed LibreOffice/Sheets
     result).
- **Static site.** A generator (plain Python + Jinja2, or a static-site
  tool) that reads `data/functions.json` + all `results/*.json` and emits
  one page per function showing a compatibility matrix (supported /
  unsupported / quirky per engine) plus the exact formula and result for
  every test case, deployed to GitHub Pages. This is the actual
  "caniuse.com for spreadsheets" product surface — everything before this
  point is the data pipeline that makes it trustworthy.

## Known gaps / honesty notes

- LibreOffice function-inventory coverage in `data/functions.json` comes
  from `help.libreoffice.org` category pages, not the (blocked) wiki page;
  a handful of functions LibreOffice actually implements may be
  under-counted there. The **executed** results in
  `results/libreoffice-24.2.json` are the authoritative source of truth for
  actual LO 24.2 behavior — treat `data/functions.json`'s `documented` flags
  as "what the docs say," not "what's actually implemented." (We saw one
  concrete instance of this gap: MAXIFS computes correctly in LO 24.2 but
  wasn't found on the specific help page category we scraped.)
- LibreOffice and Google Sheets both have executed results in `results/`.
  The Sheets sweep ran on 2026-08-29 via `harness/run_sheets.py`; its
  `engine_version` is a dated import label, **not** a version — Sheets is a
  hosted product that changes under us, so a Sheets verdict is a dated
  observation rather than a release guarantee. `_version_tuple()` in
  `site/build_site.py` deliberately sorts any non-numeric label as `(0,)`
  so the label can never be compared as a version.
- The Sheets corpus was executed from the **Excel-authored** workbooks, so a
  handful of results describe the import path rather than Sheets. Those are
  classified `inconclusive` (never `unsupported`): see
  `sheets_case_inconclusive()` in `site/build_site.py` and the caveats in
  `DATASET_CARD.md`. A follow-up Sheets run writing plain, unprefixed
  function names is the fix.
- The **recipe** corpus (`results/recipes-verified.json`) is LibreOffice-only.
  How-to pages therefore carry Sheets-executed FUNCTION verdicts but a
  LibreOffice-only worked example, and their copy says exactly that.
- Excel columns in any compatibility matrix must not be populated with
  executed data until an Excel engine exists, per the quality bar for this
  project.
- `data/functions.json`'s per-function doc URLs point at the listing page
  each function was found on (the umbrella alphabetical/category page),
  not a dedicated per-function help article — no per-function URL was
  fabricated.
