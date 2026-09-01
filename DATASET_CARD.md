# Spreadsheet formula compatibility dataset (Excel / Excel for the web / Google Sheets / LibreOffice)

An open dataset recording, for the 600 spreadsheet functions in the catalog, whether each is available in
**Microsoft Excel**, **Google Sheets**, and **LibreOffice Calc** — with the Excel-for-the-web,
Google Sheets and LibreOffice verdicts produced by *actually executing* each formula
(headless LibreOffice; Google Sheets via a dated Drive import; Excel for the web via a dated
OneDrive upload and recalculation), not scraped from documentation. **Desktop** Excel is the one
engine we do not execute; its column is Microsoft's documentation, and it is the yardstick the
executed engines are measured against. As far as we know this is the only openly available
**executed** cross-application spreadsheet-compatibility dataset.

> **Excel for the web is not desktop Excel — read this before using `excel_web_verdict`.**
> Microsoft ships two implementations of the calculation engine. We execute the web one and
> we do not execute the desktop one, so an `excel_web_verdict` is evidence about Excel for
> the web and about nothing else. `in_excel` remains a *documentation* flag describing the
> desktop product. When the two disagree we cannot tell you why: it may be the web engine
> diverging from the desktop one, or the documentation being wrong about both. We have no
> desktop run to separate those, and nothing in this dataset should be read as if we did.

- **Homepage / source:** https://canispreadsheet.com/data.html
- **Methodology:** https://canispreadsheet.com/methodology.html
- **License:** Creative Commons Attribution 4.0 (CC BY 4.0) — free to use with attribution.
- **Files:** `compat.csv` (one row per function, headered — 600 rows) and `compat.json` (object keyed by function name). Both are regenerated from the site's live test results. `compat.csv` uses the column names in the table below; `compat.json` uses short keys for the same fields — `cat`, `x`, `g`, `l`, `xwv`, `xwver`, `gv`, `gver`, `lv`, `lver`, `lnew`, in that order. Note `xwv`/`xwver` deliberately do **not** live in the `x` namespace: `x` means desktop-documented and `xw` means web-executed, and reading them as one engine is the misuse this naming exists to prevent.

## Columns

| Column | Type | Meaning |
|---|---|---|
| `function` | string | Function name, uppercase (e.g. `VLOOKUP`). |
| `category` | string | Function category (e.g. "Lookup and reference"). |
| `in_excel` | bool | Documented as available in Microsoft Excel. This is the **desktop** product: Microsoft publishes one function reference and it describes desktop Excel. A documentation flag, never a measurement. There is deliberately no `in_excel_web` column — we have not extracted per-application availability from Microsoft's reference and will not invent it. |
| `in_google_sheets` | bool | Documented as available in Google Sheets. |
| `in_libreoffice` | bool | Documented as available in LibreOffice Calc. |
| `excel_web_verdict` | string | Result of executing a real test case in **Excel for the web**: `supported`, `quirky`, `unsupported`, or `inconclusive`; empty where we have no Excel-web run. **A different application from desktop Excel** — see the warning at the top. Empty covers two very different situations: the 14 functions no engine executes, and **7 functions (`LAMBDA`, `LET`, `ISOMITTED`, `MAP`, `MAKEARRAY`, `REDUCE`, `SCAN`) that Excel for the web could not be made to open at all** — a transport limit, never evidence of missing support. See "Excel for the web: 7 transport-unreachable functions" below. |
| `excel_web_executed` | string | Label for the Excel-for-the-web run that executed **this function**, e.g. `Excel for the web (recalc, 2026-09-01)`. Like `google_sheets_executed` this is a **date, not a version** — Excel for the web is a rolling service with no release to pin — so never parse or compare it as a version string. Empty if not executed. |
| `google_sheets_verdict` | string | Result of executing a real test case in Google Sheets: `supported`, `quirky`, `unsupported`, or `inconclusive`; empty for the 14 functions this corpus deliberately does not execute (see Limitations). **`inconclusive` is not a verdict** — see "Google Sheets execution caveats" and "Executed, but no verdict drawn" below. |
| `google_sheets_executed` | string | Label for the Google Sheets run that executed **this function**, e.g. `Google Sheets (Drive import, 2026-08-29)`. This is a **date, not a version**: Sheets is a rolling service with no release to pin, so never parse or compare it as a version string. Rows can carry different dates — a later run that re-executes part of the corpus re-dates only the functions it covered; three labels are in use today (2026-08-29 on 278 functions, 2026-08-31 on 241, 2026-09-01 on 67). Empty if not executed. |
| `libreoffice_verdict` | string | Result of executing a real test case in LibreOffice: `supported`, `quirky`, `unsupported`, or `inconclusive`; empty for the same 14 deliberately unexecuted functions. **`inconclusive` is not a verdict** — see "Executed, but no verdict drawn" below. |
| `libreoffice_version_tested` | string | LibreOffice version the published verdict was executed on — the newest of the four pinned builds, currently `25.8.7.3` on every executed row. |
| `libreoffice_newly_supported_in` | string | LibreOffice version in which the function first started working, when known (from testing across multiple versions); empty otherwise. |

## How it's produced

Desktop Excel availability comes from Microsoft's official function documentation; we do not
run desktop Excel. LibreOffice verdicts come from writing each function's formula into a workbook,
converting it headlessly with LibreOffice (`soffice --headless`), and reading back the
recalculated result — with canary formulas proving the sheet actually recalculated. Tests are
run across four pinned LibreOffice builds — 24.2.0.3, 24.8.7.2, 25.2.0.3 and 25.8.7.3 — to
capture when support was added; `libreoffice_verdict` reports the newest of them.

Google Sheets verdicts come from the **same workbooks**: uploaded to Google Drive with
import-conversion on, opened as a Sheet (which recalculates every formula), exported back to
`.xlsx`, and read with the same reader. The run is dated rather than versioned. Full harness,
authored test cases, and per-version raw results:
https://github.com/xxaflabsxx/spreadsheet-compat

Both engines currently hold 586 functions / 2334 executed cases each. In the raw results
files, every function block carries its own `executed_at` date (a run that re-executes part of
the corpus re-dates only what it covered — that is what `google_sheets_executed` is built
from), and every case carries the formula as displayed and as stored in the `.xlsx`, the value
or error read back, the expected value, `matched_expected`, and the per-sheet canary flag. A
case with `expected: null` is a probe: it asserts only that the engine evaluated the call, not
what it returned.

## Google Sheets execution caveats

The Sheets run reuses Excel-authored workbooks, and that round trip has two artifacts that are
**not** Google Sheets behaviour — the first resolved for every function it was ever shown to
affect, the second still live. Where a result is explained by either, it is held
`inconclusive` rather than turned into a claim about the engine:

1. **Unmapped OOXML storage prefixes (resolved for the affected functions on 2026-08-29).**
   Excel stores post-2007 functions as `_xlfn.NAME`, `_xlfn._xlws.NAME` or `_xlpm.NAME`.
   Google's importer maps some but not all: `_xlfn.XLOOKUP`, `_xlfn.MAP` and `_xlfn.LAMBDA`
   evaluated fine even with the prefix, while `_xlfn._xlws.FILTER` and `_xlfn._xlws.SORT`
   returned `#NAME?` and BYROW/BYCOL/MAKEARRAY returned `#ERROR!` — despite Google documenting
   all five. A follow-up run on 2026-08-29 re-executed the 37 functions whose formulas carry a
   storage prefix at all, this time written with plain, unprefixed names, resolving all five:
   **BYCOL, BYROW, FILTER, MAKEARRAY, SORT** now carry real executed verdicts. The two runs are
   merged into one results file; per-function provenance (which run produced which function's
   verdict) is recorded in that file's `subset_runs` array, and each case records which
   serialization produced it in its `serialization` field. The general caveat still applies to
   any function or corpus not yet re-run this way: 129 cases across 41 functions still hold the
   `_xlfn.`-prefixed formula the original sweep wrote — all 41 came back `supported`, so the
   prefix demonstrably imported for them — while every batch since 2026-08-29 has been written
   with plain names. An `.xlsx` import can still fail to map a prefixed serialization, and such
   cases are still reported `inconclusive` rather than guessed at.
2. **Export readback.** Sheets' `.xlsx` export rounds floats to 10 significant digits
   (`PI()` exports as 3.141592654) and writes an empty cell for a blank/zero-length result.
   Where that is the only disagreement with the expected value the difference is in the export,
   not the engine, and the case is held `inconclusive` rather than counted as a quirk. As of the
   2026-09-01 run that is 9 cases in 9 functions: DEGREES, INDIRECT, PRICE, PRICEDISC, PRICEMAT,
   RECEIVED, SORT, TRANSPOSE, YIELDDISC — float rounding in the bond-price family, and an empty
   cell where a blank was expected in the rest. None of the nine costs its function a verdict:
   each has other cases that carry it.

## Excel for the web execution and its caveats

The Excel-for-the-web run (`results/excel-web.json`, 2026-09-01) executed **579 of 586**
functions and about 2,200 cases. The corpus workbook is the *same* formula-only `.xlsx` the
LibreOffice and Sheets runs use: it is uploaded to OneDrive, Excel for the web recalculates it
on open, and `File > Create a Copy > Download a Copy` brings the workbook back for readback.
A deterministic arithmetic canary (`=1111+2222`) on every sheet proves the sheet actually
recalculated, and each returned package self-identifies in `docProps/app.xml` as
`Microsoft Excel Online` (AppVersion `16.0300`), which is in-band corroboration the Drive path
never offered.

**Readback fidelity is better than the Sheets path, and was verified at full-run scale.**
Across 10,278 numeric cells in 11 of the 15 readback packages: every value round-trips
losslessly, with up to **17 significant digits** (343 cells at ≥16) and scientific notation for
small magnitudes; zero parse failures and zero `repr` round-trip losses. No rounding or
normalization is applied to Excel-web numbers, and none is needed — unlike Google's export,
which rounds to 10 significant digits. Re-reading 1,647 cases directly out of the raw
worksheet XML and comparing them to the ingested values found **no discrepancies**. The error
vocabulary is exactly the six standard OOXML typed (`t="e"`) tokens observed —
`#DIV/0!`, `#N/A`, `#NAME?`, `#NUM!`, `#REF!`, `#VALUE!` — with no Google-style `#ERROR!`
token, and notably no `#CALC!` anywhere in the run.

Two case-level round-trip artifacts are recorded and kept out of every verdict:

1. **`INDIRECT` to a missing sheet became an external-workbook reference.** The uploaded
   workbook contains no `externalLinks` part; the downloaded one contains two, targeting
   `Nope` and `NotACell` — the literal text arguments of the corpus's `INDIRECT` calls. Excel
   for the web read `'Nope'!B2` as a reference into a *workbook* named `Nope` and could not
   resolve it, which is why the cell holds `0` rather than `#REF!`, `ISERROR` is `FALSE` and
   `ISREF` is `TRUE`. Whether the blank comes from the declined "trust externally linked
   workbooks" prompt or from an unresolvable external evaluating empty in this engine, the run
   cannot tell — and both are facts about the link, not about the function. Affects one case
   each in `INDIRECT`, `IFERROR`, `IFNA`, `ISERROR` and `ISREF`.
2. **LAMBDA-bearing formulas were silently deleted.** `BYROW`/`BYCOL` are the only corpus
   formulas embedding a `LAMBDA` without an `_xlpm.`-prefixed parameter, so their chunk opened
   where the others did not. It opened, but the formula did not survive: the uploaded workbook
   carries an array formula in `A30` of each sheet and the downloaded package has **no `A30`
   at all**, while the sheet's `Z1` canary is intact. A blank left by a removed formula is not
   a computed result.

## Excel for the web: 7 transport-unreachable functions

`LAMBDA`, `LET`, `ISOMITTED`, `MAP`, `MAKEARRAY`, `REDUCE` and `SCAN` carry an **empty**
`excel_web_verdict`, and it is important not to read that as "unsupported in Excel". Desktop
Excel implements every one of them; `LAMBDA` and `LET` are Microsoft's own headline additions
to the formula language.

What failed is the transport. Excel for the web's **file-open refuses any workbook whose
stored formulas carry the `_xlpm.`/`LAMBDA` storage serialization**, with "Couldn't Open the
Workbook" and no cell-level diagnostic. This was established by bisection, not assumed:

| Workbook | Result |
|---|---|
| chunk-08 (40 functions) | refused |
| the same chunk minus `LAMBDA`/`LET`/`ISOMITTED` (36 functions) | opened |
| `LAMBDA` + `LET` + `ISOMITTED` alone | refused |
| `MAP`/`MAKEARRAY`/`REDUCE`/`SCAN` class probe alone | refused |
| `LINEST` alone (the other suspect) | opened, ingested normally |

The irony is the finding, and it is a fact about the product rather than about the functions:
**Excel for the web will not open a file containing Excel's own LAMBDA-family storage
serialization.** The `_xlpm.` token exists precisely because Excel needed a way to write
LAMBDA parameters into OOXML; the web build rejects the file rather than reading them. These
seven are declared skips for the `excel_web` engine only — their LibreOffice and Google Sheets
verdicts are unaffected and are published as usual.

This mirrors the earlier LibreOffice alias-collapse decision in shape: there, LibreOffice's own
`.xlsx` export collapsed eleven names onto another function's token, so no OOXML token reached
the engine and the verdict was declined rather than published as "unsupported". Same judgement,
opposite layer — there the *writer* could not express the name, here the *reader* will not
accept it. In both cases the engine's actual capability is untouched and unmeasured.

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
PIVOTBY, SORTBY, TEXTBEFORE, TEXTAFTER), is a real `unsupported` verdict. As of the 2026-09-01
run every one of the 68 `unsupported` Sheets verdicts is of that kind: no function Google
documents is marked unsupported.

## Citation

> Can I Spreadsheet? — Open spreadsheet formula compatibility dataset. https://canispreadsheet.com/data.html (CC BY 4.0).

## Limitations

- Desktop Excel verdicts are documentation-based, not live-executed. Excel for the web,
  Google Sheets and LibreOffice are all executed.
- An Excel-for-the-web verdict is **not** a desktop Excel verdict, and a disagreement
  between `excel_web_verdict` and `in_excel`/the documented expected value is ambiguous:
  we have no desktop run and cannot say whether the web engine diverges or the
  documentation is wrong about both.
- Excel for the web has no version to pin either: its verdict is a dated observation
  (`excel_web_executed`), not a release guarantee.
- Google Sheets has no version to pin: a Sheets verdict is a dated observation
  (`google_sheets_executed`), not a release guarantee. Google ships changes continuously.
- As of the 2026-08-29 plain-name re-run, no function carries `google_sheets_verdict =
  inconclusive` **because of the importer** — the five that did (BYCOL, BYROW, FILTER,
  MAKEARRAY, SORT) were resolved by that run (see caveats above). The condition can recur for a
  function or corpus we haven't re-run with plain names yet. Eight functions do carry
  `inconclusive` for the other reason, "executed, but no verdict drawn": `AI`, `IMPORTDATA`,
  `IMPORTFEED`, `IMPORTHTML`, `IMPORTRANGE`, `IMPORTXML` and `SPARKLINE` in
  `google_sheets_verdict`, and `DDE` in `libreoffice_verdict`. Either way, treat `inconclusive`
  as "no verdict" and fall back to `in_excel` / `in_google_sheets` / `in_libreoffice`, never as
  "unsupported".
- 586 of the 600 catalog functions have executed cases in both engines. The 14 that do not are
  deliberate, published skips, not a backlog: `CALL` and `REGISTER.ID` (the argument names
  external code to load), `WEBSERVICE` (its documented behaviour is the network round trip, so
  the `#N/A` we would record describes the sandbox), and eleven LibreOffice legacy aliases
  (`CUMIPMT_ADD`, `CUMPRINC_ADD`, `EFFECT_ADD`, `NOMINAL_ADD`, `GCD_EXCEL2003`, `LCM_EXCEL2003`,
  `ISEVEN_ADD`, `ISODD_ADD`, `WEEKNUM_EXCEL2003`, `FORMULA`, `SKEWP`) whose name LibreOffice's
  own `.xlsx` export collapses onto another function's token, so no test can address them
  through this transport. Those 14 rows carry empty `libreoffice_verdict` /
  `google_sheets_verdict`; each one's reason is published at
  https://canispreadsheet.com/methodology.html.
- Corrections and additions welcome via the GitHub repository.
