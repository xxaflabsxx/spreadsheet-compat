# Plan: the 78 functions Excel does not document

Batch G (RTD..ZTEST) closed the Excel-documented set. Every remaining untested
function has `x == false` in `docs/data/compat.json`: **Excel does not document it
at all.** That changes what the corpus is allowed to say about them, so this plan
is written before any of them is executed. Nothing here has been run yet.

Counts are live as of batch G: `78` untested functions with `x == false`,
`47` Google-documented and `31` LibreOffice-documented, with no overlap.

## The groups

### A. Google-documented, deterministic offline — 38
The operator functions (`ADD`, `DIVIDE`, `EQ`, `GT`, `GTE`, `LT`, `LTE`, `MINUS`,
`MULTIPLY`, `NE`, `POW`, `UMINUS`, `UPLUS`, `UNARY_PERCENT`, `ISBETWEEN`), the
`TO_*` parsers, the array/text helpers (`ARRAY_CONSTRAIN`, `FLATTEN`, `JOIN`,
`SPLIT`, `SORTN`, `REGEXMATCH`, `COUNTUNIQUE`, `AVERAGE.WEIGHTED`,
`MARGINOFERROR`, `EPOCHTODATE`, `ISDATE`, `ISEMAIL`, `ISURL`, `ARRAYFORMULA`,
`QUERY`, `IMCOTH`, `IMLOG`, `IMTANH`).
**Execution:** normal corpus treatment. Derive independently, ROUND-wrap, assert.
**Authority:** Google's own function-list pages, cited by full URL *and* the date
read — Google publishes no version numbers, so a bare URL dates nothing.
`QUERY` additionally needs its Google Visualization API query-language page cited
separately; the function page does not define the language.

### B. Google-documented, service- or context-bound — 9
`AI`, `GOOGLEFINANCE`, `GOOGLETRANSLATE`, `IMPORTDATA`, `IMPORTFEED`,
`IMPORTHTML`, `IMPORTRANGE`, `IMPORTXML`, `SPARKLINE`.
**Execution: probe only, `expected: null`**, exactly as `COPILOT`, the `CUBE*`
family, `DETECTLANGUAGE`, and batch G's `RTD` / `STOCKHISTORY` / `TRANSLATE`.
Each note must say which dependency makes the value unassertable (live market
data, a translation service, a network fetch, another user's spreadsheet, a
rendered chart rather than a value) and must record what the multi-spelling
probe established. `SPARKLINE` returns a *drawing*, not a value: say so rather
than recording a blank as a result.

### C. LibreOffice-only — 31, in three sub-groups
- **C1. Legacy alias variants — 11.** `CONVERT_OOO`, `CUMIPMT_ADD`,
  `CUMPRINC_ADD`, `EFFECT_ADD`, `GCD_EXCEL2003`, `ISEVEN_ADD`, `ISODD_ADD`,
  `LCM_EXCEL2003`, `NOMINAL_ADD`, `WEEKNUM_EXCEL2003`, `WEEKNUM_OOO`.
  These are *storage-form* names for Excel-compatibility add-in functions, not
  separate behaviours. Every one must carry a case asserting it against its
  modern namesake in the same engine (`GCD_EXCEL2003` vs `GCD`, and so on) —
  that identity is the whole content of the group, and batch F's
  `NEGBINOMDIST`/`NEGBINOM.DIST` divergence proves aliases do drift.
  Expect a storage-form fight: probe all five spellings first, as batch G did.
- **C2. Ordinary LO-only functions — 16.** `CHISQDIST`, `CHISQINV`,
  `DAYSINMONTH`, `DAYSINYEAR`, `EASTERSUNDAY`, `ERRORTYPE`, `FORMULA`,
  `ISLEAPYEAR`, `MONTHS`, `RAWSUBTRACT`, `REGEX`, `ROT13`, `SKEWP`, `WEEKS`,
  `WEEKSINYEAR`, `YEARS`. Normal treatment. `SKEWP` gets a cross-check against
  batch G's `SKEW.P`; `RAWSUBTRACT` exists precisely to *not* do the
  floating-point tidying `-` does, so its test must be the one that shows that.
- **C3. Nondeterministic or context-bound — 4.** `CURRENT`, `DDE`, `RAND.NV`,
  `RANDBETWEEN.NV`. `DDE` is `RTD`'s twin (an external data server) and takes
  the `RTD` treatment verbatim. `RAND.NV`/`RANDBETWEEN.NV` have no assertable
  value but *do* have an assertable property — non-volatility — so assert the
  property (two references to one cell agree; the value survives a recalc) and
  never a number. `CURRENT` depends on position within its own formula.
- **Authority:** `help.libreoffice.org` module pages, cited by full URL, the date
  read, **and** the version of the help the page serves — LibreOffice versions
  its help and the corpus already pins four engine builds.

## Honesty rules that apply to all 78

1. **Never claim an Excel result.** These functions have no Excel column. No
   case note may say "matches Excel's documented behavior", and no title or
   description may reference Excel docs.
2. **Engine-scope every claim.** A group-A function cannot be executed in
   LibreOffice and a group-C function cannot be executed in Google Sheets. The
   corpus must publish one engine's result and say the other was not run —
   never imply it was.
3. **Cite by URL + date read** (Google has no versions; LibreOffice help does —
   record it).
4. Probes stay probes: `expected: null`, and the note says what makes the value
   unassertable.

## Rendering gap — this is a blocker, and it is real today

Zero functions currently have results in only one engine, so none of the
single-engine paths in `site/build_site.py` has ever rendered. Four problems,
all confirmed by reading the code and the generated HTML:

- **`site/build_site.py:711-713` — fabricated LibreOffice execution. BLOCKING.**
  `build_function_title_desc` does `verdict = le["verdict"]` / `total =
  len(le["cases"])` whenever `r["any_tested"]` is set, with no check that
  LibreOffice is the engine that ran. A Sheets-only executed function emits
  *"All 0 executed test cases pass in LibreOffice"* and renders
  `le['version']` as `None`. **Must be fixed before group A or B is executed.**
- **`site/build_site.py:1464` — the page lede** says "executed in ... LibreOffice
  Calc" whenever Sheets is untested, which is exactly backwards for a
  Sheets-only function.
- **`site/build_site.py:1500` — the Excel matrix cell** prints
  `No — documented only` for the Excel row *unconditionally*, ignoring `x`.
  This is already wrong on all 78 pages: `docs/functions/query.html` today reads
  `Documented: No` beside `Live-tested: No — documented only`. Needs an
  `e.documented` guard, not an `ek == 'excel'` one.
- **`site/build_site.py:808-812` and `:872`** hardcode "match Excel's docs" /
  "Excel/Sheets from official docs" in the quirky-verdict description even when
  `excel_doc` is false. The *title* already has an `excel_doc`-aware fallback at
  `:800-806`; the description does not.

**The mechanism to copy already exists.** `load_recipes` `_merge_check()`
(`site/build_site.py:2840-2888`) reads an `engines` scope off each recipe spec,
hides a row whose engine never ran it, and renders `n/a (Sheets-only formula)`
in the other column, with per-engine counts at `:2921-2929`. Function pages have
no equivalent — `build_records` and `FUNCTION_TMPL` have no scope concept at
all. Porting that idea to function records is the clean fix.

**`scripts/check_honesty.py` will not catch any of this.** Its
`FALSE_EXECUTION` guard (`:66-97`) only inspects Excel and "all three engines"
claims, so the fabricated *LibreOffice* execution above passes it silently. Add
a guard for it in the same change.

## Order of work

1. Fix the four rendering gaps and add the `check_honesty` guard. No execution.
2. Group C2 + C1 first — LibreOffice-only, and the harness already runs four
   LibreOffice builds, so it needs no new plumbing and exercises the LO-only
   render path.
3. Group A on Sheets, which exercises the Sheets-only render path.
4. Groups B and C3 last: probes only, no new machinery, nothing to get wrong
   once the precedents above are set.
