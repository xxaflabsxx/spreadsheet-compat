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
| `google_sheets_verdict` | string | Result of executing a real test case in Google Sheets: `supported`, `quirky`, `unsupported`, or `inconclusive`; empty if the function is not yet in the executed test set. **`inconclusive` is not a verdict** — see "Google Sheets execution caveats" below. |
| `google_sheets_executed` | string | Label for the Google Sheets run, e.g. `Google Sheets (Drive import, 2026-08-29)`. This is a **date, not a version**: Sheets is a rolling service with no release to pin, so never parse or compare it as a version string. Empty if not executed. |
| `libreoffice_verdict` | string | Result of executing a real test case in LibreOffice (e.g. `supported`, `unsupported`); empty if the function is documented-only and not yet in the executed test set. |
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

The Sheets run reuses Excel-authored workbooks, and that round trip has two artifacts that are
**not** Google Sheets behaviour. Where a result is explained by either, the verdict is
`inconclusive` rather than a claim about the engine:

1. **Unmapped OOXML storage prefixes.** Excel stores post-2007 functions as `_xlfn.NAME`,
   `_xlfn._xlws.NAME` or `_xlpm.NAME`. Google's importer maps some but not all: `_xlfn.XLOOKUP`,
   `_xlfn.MAP` and `_xlfn.LAMBDA` evaluate fine, while `_xlfn._xlws.FILTER` and
   `_xlfn._xlws.SORT` return `#NAME?` and BYROW/BYCOL/MAKEARRAY return `#ERROR!` — all five of
   which Google documents. Affected: **BYCOL, BYROW, FILTER, MAKEARRAY, SORT**. A follow-up run
   writing plain, unprefixed function names will resolve them.
2. **Export readback.** Sheets' `.xlsx` export rounds floats to 10 significant digits
   (`PI()` exports as 3.141592654) and writes an empty cell for a blank/zero-length result.
   Where that is the only disagreement with the expected value (DEGREES(1), INDIRECT to a blank
   cell, TRANSPOSE of a blank cell) the difference is in the export, not the engine.

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
- Five functions carry `google_sheets_verdict = inconclusive` (see caveats above). Consumers
  should treat that as "no verdict" and fall back to `in_google_sheets`, never as "unsupported".
- The executed test set covers the most-used functions; documented-only functions have empty
  `libreoffice_verdict` / `google_sheets_verdict`.
- Corrections and additions welcome via the GitHub repository.
