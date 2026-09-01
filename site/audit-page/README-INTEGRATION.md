# Migration Audit page (E2) — integration & deploy notes

Status: **built, 174/174 automated tests passing** (`node test.mjs`), on top of the
unchanged E1 parser (its own suite still passes 67/67). Not yet browser-smoke-tested —
do the 2-minute manual check below before announcing.

| Suite | Command | Assertions |
|---|---|---|
| Verdict engine + E2E | `node site/audit-page/test.mjs` | 174 |
| Adversarial extractors (audit + built checker) | `node site/audit-page/test-adversarial.mjs` | 170 |
| E1 parser prototype | `node site/audit-prototype/test.mjs` | 67 |
| Sitewide honesty guard (after a rebuild) | `python3 scripts/check_honesty.py` | exit 0 |

## Files

| File | Purpose |
|---|---|
| `audit.html` | The page. Standalone (inline CSS matching the site's house style, incl. nav + footer + print stylesheet). |
| `audit.js` | **Byte-for-byte copy of `../audit-prototype/audit.js`** (E1 parser). `test.mjs` fails if it ever drifts from the prototype. |
| `audit-verdicts.js` | Pure verdict engine: classification, report building, at-risk ordering, CSV export, license-key format check, function→divergence-guide lookup (`guidesForFunction`), and the per-release LibreOffice target rule (`compareVersions`, `resolveTargetVersion`, `targetLabel`). No DOM/network. |
| `audit-app.js` | DOM glue: dropzone, direction picker, target-LibreOffice-version picker (persisted in `localStorage['csps-audit-lo-version']`), rendering, free/paid tiers, Gumroad license verify, CSV download, print, optional guide-index fetch (`data/guides.json`, silent-fail). |
| `make_fixtures.py` | Generates `verdict-mix.xlsx` (needs openpyxl: `/home/jon/venv/bin/python make_fixtures.py`). |
| `verdict-mix.xlsx` | E2E fixture: safe functions + dataset-verified breakers per target (see docstring in `make_fixtures.py`). |
| `test.mjs` | 174 assertions: verdict-engine unit tests (incl. `guidesForFunction`, version comparison and the four LibreOffice target releases) + E2E through the real parser for 3 directions and 2 target versions. `node test.mjs`. |
| `test-adversarial.mjs` | 170 assertions: hostile formulas (quoted parens, structured refs, `_xlfn.` prefixes, LET/LAMBDA names, sheet names with parens, 8k-char formulas) through BOTH extractors — `extractFunctions()` here and the checker's `funcs()` read straight out of the built `docs/checker.html`. `node test-adversarial.mjs`. |

## Deploy steps

1. Copy these four files into `docs/` (flat, next to `checker.html`):
   `audit.html`, `audit.js`, `audit-verdicts.js`, `audit-app.js`.
   The page loads the dataset from `data/compat.json` **relative to the page**, which is
   already deployed at `docs/data/compat.json`. No build step needed. It also fetches
   `data/guides.json` (same directory, also already deployed by the site build) to show
   a "why?" link next to functions with a documented divergence guide — that fetch is
   an optional enhancement and fails silently, so the audit works fine without it.
   (Alternatively, fold `audit.html` into `build_site.py` as a template later; it
   deliberately mirrors the built `checker.html` head/nav/footer so either path works.)
2. Fill the two consts at the top of **`audit-app.js`**:
   - `PRODUCT_ID` — the Gumroad product id for the license product (used by
     `POST /v2/licenses/verify`). **Until you fill this, any correctly-formatted key
     unlocks (offline format check only)** — fine pre-launch, not after.
   - `GUMROAD_URL` — the public purchase URL for the Buy button
     (currently `https://aflabs.gumroad.com/l/FILL-ME-IN`).
   - Optional: `PRICE_NOW` / `PRICE_LATER` (currently `$19` launch / `$29`). When the
     launch discount ends, set both to `$29` (the strike-through hides itself only if
     you edit the HTML; simplest is to set `PRICE_NOW='$29'` and delete the
     `price-later` span in `audit.html`).
   - When creating the Gumroad product: **enable license keys**, single price $29 with
     a $19 launch discount code or launch price.
3. Add the nav link `<a href="audit.html">Migration&nbsp;Audit</a>` to `BASE_TMPL`'s
   nav in `build_site.py` (audit.html's own nav already includes it) and rebuild the
   site, otherwise the rest of the site won't link here.
4. Add `audit.html` to the sitemap generation in `build_site.py`.
5. Manual browser smoke test (not yet done): `python3 -m http.server` from `docs/`,
   open `audit.html`, drop `verdict-mix.xlsx`, and check against the tested numbers:
   12 formulas / 2 sheets / 11 unique functions; Excel→Sheets: 5 at-risk formulas,
   at-risk functions GROUPBY, AGGREGATE, BAHTTEXT (detail) + TEXTSPLIT (locked);
   Excel→LibreOffice: 7 at-risk formulas, free detail on GROUPBY/ARRAYFORMULA/GOOGLEFINANCE.
   Then switch the direction to Excel→LibreOffice and the version select to 24.2.0.3:
   TEXTSPLIT must flip to MISSING with a "works since 25.8.7.3" reason, the summary line
   must read `Target: LibreOffice Calc 24.2.0.3`, and the choice must survive a reload.
   Confirm in the network tab that the only requests are the page assets +
   `data/compat.json` + `data/guides.json` (and Gumroad only when a key is submitted).

## License verification — CORS findings (tested 2026-08-23)

`api.gumroad.com/v2/licenses/verify` **is CORS-open**: the OPTIONS preflight answers
`access-control-allow-origin: *`, `access-control-allow-methods: ... POST`,
`access-control-allow-headers: content-type`. A browser `fetch` POST with a
form-encoded body works cross-origin. So:

- **Primary path**: online verify at unlock time (counts one "use"), plus a silent
  re-verify on every page load with `increment_uses_count=false`. Refunded /
  chargebacked / disputed purchases are rejected.
- **Fallback** (network unreachable, ad-blocker, or `PRODUCT_ID` unset): offline
  format check (`XXXXXXXX-XXXXXXXX-XXXXXXXX-XXXXXXXX` hex) + localStorage persistence,
  with an honest "accepted by offline format check" notice in the UI. No server-side
  proxy is required given the open CORS, but if Gumroad ever closes it, the fallback
  keeps paying users unlocked and a tiny proxy would be the fix — noted in code
  comments in `audit-app.js`.
- Unlock state is stored in `localStorage['csps-audit-license']` per browser.
  Nothing is signed; a determined user can unlock DevTools-style. That's the accepted
  trade-off for a fully client-side product — the license is honesty-priced.

## Design decisions worth knowing

- **Per-release LibreOffice targets**: the direction picker's LibreOffice options are
  paired with a version select (`25.8` default / `25.2` / `24.8` / `24.2`, each labelled
  with the exact tested build). The rule, in `classifyFunction`:
  `lv === 'unsupported'` → MISSING for every target; otherwise `lnew` set and
  `compareVersions(target, lnew) < 0` → MISSING with an executed reason ("returns
  #NAME? in LibreOffice 24.2.0.3 (executed) — it works since 24.8.7.2"); otherwise the
  `lv` verdict, unchanged. `lnew` is the earliest release we tested a function *working*
  in, so this is an executed fact. Comparison is numeric per `major.minor.patch.build`,
  and an unrecognized version resolves to the latest tested build — never harsher than
  what we tested. **Default (25.8.7.3) verdicts and note wording are byte-identical to
  the pre-feature engine**; a test walks all 600 dataset entries × 2 sources to prove it.
  Honesty note: the dataset has no per-version presence table, so an `lnew: null`
  function is described as "supported in every LibreOffice release we tested
  (24.2.0.3 → 25.8.7.3)" rather than "executed in <that older build>". Quirks were
  measured in `lver` (25.8.7.3) and the note says so when an older target is picked.
  The same rule (same comparison, same four builds) is implemented in the checker's
  inline JS in `build_site.py` — `cmpVer()` / `migrate()` — with a `&v=` permalink param.
- **Verdict precedence (LibreOffice targets)**: executed `lv` always outranks the
  documentation flag — `supported`→OK, `quirky`→QUIRK ("recognized, but ≥1 executed
  case returned a different value/error than Excel"), `unsupported`→MISSING (#NAME?),
  `lv:null` falls back to the documented flag and the UI *labels the basis*
  ("execution-verified" vs "documentation-based") on every row. Desktop-Excel targets
  are documentation-based and say so; Google Sheets and Excel-for-the-web targets are
  execution-verified from `gv`/`xwv` (this line was already stale about Sheets before
  the web engine landed).
- **At-risk ordering / free top-3**: severity class first (MISSING before QUIRK), then
  usage count, then name. Rationale: in LibreOffice targets, ubiquitous quirk-flagged
  functions (SUM, VLOOKUP are genuine executed quirks) would out-count every hard
  breaker and the free tier would never show a single #NAME? case. To switch to pure
  usage-count ordering, drop the first comparison in `compareAtRisk`
  (`audit-verdicts.js`) — one line, tests pin the current behavior.
- **Unknown functions** are a separate tile + section, never counted as at-risk, with
  explicit "we do not guess" copy.
- **At-risk formula count** = formulas containing ≥1 MISSING or QUIRK function
  (worst-verdict-wins per formula).

## Honest limitations (also reflected in the page copy)

- **Function-level triage, not recalculation.** We match functions against a
  pre-computed dataset (executed LibreOffice, Google Sheets and Excel-for-the-web
  results; vendor docs for desktop Excel).
  We cannot promise specific numbers survive: argument-level differences within a
  supported function, cross-formula interactions, locale/separator issues, data types
  are out of scope.
- **Desktop Excel verdicts are documentation-based.** LibreOffice, Google Sheets and
  Excel for the web are all execution-verified, from `lv`, `gv` and `xwv` respectively.
  The UI labels the basis per row; don't remove that.
- **Excel for the web is not desktop Excel.** It is a separate application with its own
  calculation engine, and `xwv` measures only that one. There is no `xw` documentation
  flag, so a function with `xwv: null` returns UNKNOWN — it must never inherit `x`.
  Seven LAMBDA-family functions are null for a transport reason (Excel for the web
  refuses to open a workbook carrying their stored serialization), not for lack of
  support.
- **Per-release verdicts only cover the four builds we execute** (24.2.0.3, 24.8.7.2,
  25.2.0.3, 25.8.7.3). Picking "24.8" means "the 24.8.7.2 build we tested", not every
  24.8.x point release, and functions with no executed data (`lv: null`) do not vary by
  version at all — the page says so on the row.
- **Shared formulas**: member cells are reconstructed by reference-shifting the group
  master (E1 tested), but a malformed file with an orphan member yields an empty
  formula recorded as such, and data-table formulas are treated as plain.
- **Legacy `.xls` unsupported** (clear save-as-.xlsx message); `.xlsm` is read but
  VBA/macros are not analyzed; defined names called like functions show up as
  UNKNOWN; external-workbook references are not resolved; ZIP64 rejected.
- Needs `DecompressionStream('deflate-raw')`: Chrome/Edge 103+, Firefox 113+,
  Safari 16.4+ (older browsers get an explicit message).
- Rendering caps: 300 cells per function / 2000 formula rows on screen (noted inline);
  the CSV export always contains every row.

## Suggested metadata (already in the file)

- `<title>`: `Spreadsheet Migration Audit — which formulas break when you switch apps?`
- meta description: `Drop an .xlsx and see which formulas break or silently change in
  Google Sheets, LibreOffice, or Excel — matched against execution-verified
  compatibility data. The file never leaves your browser.`

## Sitewide honesty guard

`python3 scripts/check_honesty.py` — scans every built page in `docs/` and fails if any copy claims **desktop** Excel results were verified/tested/executed. LibreOffice, Google Sheets and Excel for the web are all executed, so claims about those are allowed when the results files back them (rule 1b checks that per function). It also fails an unqualified "we do not run Excel" (rule 1c — say "desktop Excel", since we DO run the web app) and any Excel-for-the-web value rendered under a heading naming Excel without "for the web" (rule 7). Run after every rebuild; negations and "documented in all three" are allowed.
