# spreadsheet-compat

"caniuse.com for spreadsheets" — a database of *tested, real, executed*
function behavior across Excel, Google Sheets, and LibreOffice Calc.

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

(Phase 2, not built yet)
harness/run_sheets.py        Google Sheets engine runner via the Sheets API.
harness/run_excel.py         Excel engine runner (see Phase-2 notes below).
site/                        Static site generator consuming data/ + results/
                              -> deployed to GitHub Pages.
```

**Pipeline**: inventory (what exists) → tests (what behavior to check) →
engine runners (what actually happens) → results (raw truth) → site
(presentation layer, Phase 2+).

## Phase 1 status

- `data/functions.json`: 600 distinct function names inventoried from live
  official docs. Excel: 522 documented, Google Sheets: 515, LibreOffice: 469,
  documented in all three: 421. Sources actually fetched are listed in the
  `sources` array; one attempted source
  (`wiki.documentfoundation.org/List_of_Calc_Functions`) was blocked by an
  anti-bot wall and is recorded honestly as `fetched: false` — LibreOffice
  coverage instead comes from `help.libreoffice.org`'s category pages (18
  pages fetched), which gave full, real coverage anyway.
- `data/tests/`: 31 functions, 125 hand-authored test cases (see the Phase-1
  task's exact function list — a couple of "X/Y" pairs like TEXTBEFORE/AFTER
  were split into two full function files each, so 31 files instead of a
  literal 30).
- `harness/run_lo.py` executed against LibreOffice 24.2.7.2, results
  committed at `results/libreoffice-24.2.json`. **Recalculation is proven
  genuine** — see "How the LO runner forces recalculation" below.

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

- **Google Sheets engine.** Use the Sheets API
  (`spreadsheets.values.update` + `spreadsheets.get` with
  `valueRenderOption=UNFORMATTED_VALUE`, or batchUpdate) against a
  disposable spreadsheet: write each test case's setup cells + formula to
  its own sheet/tab (same one-sheet-per-case layout as the LO runner),
  then read back computed values. Google Sheets recalculates on write, so
  the "is this really recalculated" concern is much smaller than with
  LibreOffice, but the same canary pattern (deterministic + volatile) should
  still be applied for parity and to catch API quirks (e.g. stale reads
  from a cache layer). Needs a Google Cloud service account with Sheets API
  enabled; formulas may need locale-specific argument separators depending
  on the target spreadsheet's locale settings.
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
- Only LibreOffice has executed results so far. Excel and Google Sheets
  columns in any future compatibility matrix must not be populated until
  Phase 2 engines exist, per the quality bar for this project.
- `data/functions.json`'s per-function doc URLs point at the listing page
  each function was found on (the umbrella alphabetical/category page),
  not a dedicated per-function help article — no per-function URL was
  fabricated.
