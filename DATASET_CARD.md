# Spreadsheet formula compatibility dataset (Excel / Google Sheets / LibreOffice)

An open dataset recording, for ~600 spreadsheet functions, whether each is available in
**Microsoft Excel**, **Google Sheets**, and **LibreOffice Calc** — with the Google Sheets
and LibreOffice verdicts produced by *actually executing* each formula (headless LibreOffice,
and Google Sheets via a dated Drive import), not scraped from documentation. Excel is the one
engine we do not execute; its column is Microsoft's documentation, and it is the yardstick the
executed engines are measured against. As far as we know this is the only openly available
**executed** cross-application spreadsheet-compatibility dataset.

- **Homepage / source:** https://canispreadsheet.com/data.html
- **Methodology:** https://canispreadsheet.com/methodology.html
- **License:** Creative Commons Attribution 4.0 (CC BY 4.0) — free to use with attribution.
- **Files:** `compat.csv` (one row per function, headered) and `compat.json` (object keyed by function name). Both are regenerated from the site's live test results.

## Columns

| Column | Type | Meaning |
|---|---|---|
| `function` | string | Function name, uppercase (e.g. `VLOOKUP`). |
| `category` | string | Function category (e.g. "Lookup and reference"). |
| `in_excel` | bool | Documented as available in Microsoft Excel. |
| `in_google_sheets` | bool | Documented as available in Google Sheets. |
| `in_libreoffice` | bool | Documented as available in LibreOffice Calc. |
| `google_sheets_verdict` | string | Result of executing a real test case in Google Sheets: `supported`, `quirky`, `unsupported`, or `inconclusive`; empty if the function is not yet in the executed test set. **`inconclusive` is not a verdict** — see "Google Sheets execution caveats" and "Executed, but no verdict drawn" below. |
| `google_sheets_executed` | string | Label for the Google Sheets run that executed **this function**, e.g. `Google Sheets (Drive import, 2026-08-29)`. This is a **date, not a version**: Sheets is a rolling service with no release to pin, so never parse or compare it as a version string. Rows can carry different dates — a later run that re-executes part of the corpus re-dates only the functions it covered. Empty if not executed. |
| `libreoffice_verdict` | string | Result of executing a real test case in LibreOffice: `supported`, `quirky`, `unsupported`, or `inconclusive`; empty if the function is documented-only and not yet in the executed test set. **`inconclusive` is not a verdict** — see "Executed, but no verdict drawn" below. |
| `libreoffice_version_tested` | string | LibreOffice version the executed test ran on (e.g. `25.8.7.3`). |
| `libreoffice_newly_supported_in` | string | LibreOffice version in which the function first started working, when known (from testing across multiple versions); empty otherwise. |

## How it's produced

Excel availability comes from Microsoft's official function documentation; we do not run
Excel. LibreOffice verdicts come from writing each function's formula into a workbook,
converting it headlessly with LibreOffice (`soffice --headless`), and reading back the
recalculated result — with canary formulas proving the sheet actually recalculated. Tests are
run across several LibreOffice versions to capture when support was added.

Google Sheets verdicts come from the **same workbooks**: uploaded to Google Drive with
import-conversion on, opened as a Sheet (which recalculates every formula), exported back to
`.xlsx`, and read with the same reader. The run is dated rather than versioned. Full harness,
authored test cases, and per-version raw results:
https://github.com/xxaflabsxx/spreadsheet-compat

## Google Sheets execution caveats

The Sheets run reuses Excel-authored workbooks, and that round trip has one remaining artifact
that is **not** Google Sheets behaviour. Where a result is explained by it, the verdict is
`inconclusive` rather than a claim about the engine:

1. **Unmapped OOXML storage prefixes (resolved for the affected functions on 2026-08-29).**
   Excel stores post-2007 functions as `_xlfn.NAME`, `_xlfn._xlws.NAME` or `_xlpm.NAME`.
   Google's importer maps some but not all: `_xlfn.XLOOKUP`, `_xlfn.MAP` and `_xlfn.LAMBDA`
   evaluated fine even with the prefix, while `_xlfn._xlws.FILTER` and `_xlfn._xlws.SORT`
   returned `#NAME?` and BYROW/BYCOL/MAKEARRAY returned `#ERROR!` — despite Google documenting
   all five. A follow-up run on 2026-08-29 wrote the same corpus with plain, unprefixed function
   names, resolving all five: **BYCOL, BYROW, FILTER, MAKEARRAY, SORT** now carry real executed
   verdicts. The two runs are merged into one results file; per-function provenance (which run
   produced which function's verdict) is recorded in that file's `subset_runs` array. The
   general caveat still applies to any function or corpus not yet re-run this way: an `.xlsx`
   import can still fail to map a prefixed serialization, and such cases are still reported
   `inconclusive` rather than guessed at.
2. **Export readback.** Sheets' `.xlsx` export rounds floats to 10 significant digits
   (`PI()` exports as 3.141592654) and writes an empty cell for a blank/zero-length result.
   Where that is the only disagreement with the expected value (DEGREES(1), INDIRECT to a blank
   cell, TRANSPOSE of a blank cell) the difference is in the export, not the engine.

## Executed, but no verdict drawn

`inconclusive` has a second cause, and it applies to **either** executed engine. Some functions'
documented behaviour depends on something no flat test workbook can supply — an external data
server, a live fetch, another user's authorization, a plan entitlement — so the engine evaluates
the call and returns something that describes the **missing dependency** rather than the engine.
Scoring that as `quirky` would publish a defect the run never demonstrated, so those cases are
excluded from the verdict, the value is still published as it came back, and the reason is
printed beside it. As of the 2026-09-01 run there are 8 such cases across 8 functions:

| Function | Engine | What came back |
|---|---|---|
| `DDE` | LibreOffice | `#N/A` — a DDE link needs a running server application |
| `AI` | Google Sheets | the formula's own text, i.e. nothing evaluated it (the feature needs an eligible Workspace/AI plan) |
| `IMPORTDATA`, `IMPORTFEED`, `IMPORTHTML`, `IMPORTXML` | Google Sheets | `#REF!` — the corpus URL is deliberately inert and the workbook was uploaded, not authorised |
| `IMPORTRANGE` | Google Sheets | `#REF!` — nobody clicked **Allow Access**, which Google documents as required |
| `SPARKLINE` | Google Sheets | an empty cell — the function draws a chart *in* the cell, and a drawing has no cached value for an `.xlsx` export to carry |

Two neighbours in the same batch went the other way and keep ordinary `supported` verdicts,
because they genuinely evaluated in the engine that documents them: `GOOGLEFINANCE` returned a
live quote and `GOOGLETRANSLATE` returned a translation. Their **values** are still unasserted —
Google documents the quote as "delayed by up to 20 minutes" — so their cases are probes and the
verdict means "this engine evaluated the call and returned a value of the documented kind".

A `#NAME?` from Sheets on a formula with *no* storage prefix, or for a function Google does not
document (TEXTSPLIT, TAKE, DROP, AGGREGATE, ARRAYTOTEXT, FORECAST.ETS, GROUPBY, NUMBERVALUE,
PIVOTBY, SORTBY, TEXTBEFORE, TEXTAFTER), is a real `unsupported` verdict.

## Citation

> Can I Spreadsheet? — Open spreadsheet formula compatibility dataset. https://canispreadsheet.com/data.html (CC BY 4.0).

## Limitations

- Excel verdicts are documentation-based, not live-executed. Google Sheets and LibreOffice
  are both executed.
- Google Sheets has no version to pin: a Sheets verdict is a dated observation
  (`google_sheets_executed`), not a release guarantee. Google ships changes continuously.
- As of the 2026-08-29 plain-name re-run, no function carries `google_sheets_verdict =
  inconclusive` — the five that did (BYCOL, BYROW, FILTER, MAKEARRAY, SORT) were resolved by
  that run (see caveats above). The condition can recur for a function or corpus we haven't
  re-run with plain names yet; when it does, treat `inconclusive` as "no verdict" and fall back
  to `in_google_sheets`, never as "unsupported".
- The executed test set covers the most-used functions; documented-only functions have empty
  `libreoffice_verdict` / `google_sheets_verdict`.
- Corrections and additions welcome via the GitHub repository.
