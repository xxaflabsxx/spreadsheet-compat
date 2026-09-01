#!/usr/bin/env python3
"""Honesty guard for the built site.

Ground truth (see site/build_site.py and results/*.json):

  * LibreOffice Calc     -- EXECUTED (four pinned builds, results/libreoffice-*.json)
  * Google Sheets        -- EXECUTED (one dated Drive-import run,
                            results/google-sheets.json, 2026-08-29)
  * Excel for the web    -- EXECUTED (one dated OneDrive-recalculation run,
                            results/excel-web.json, 2026-09-01). A SEPARATE
                            IMPLEMENTATION from the desktop product.
  * Microsoft Excel      -- NOT EXECUTED, and never will be by this harness.
    (DESKTOP)               Documentation only, and it is the yardstick every
                            executed engine is measured against.

THE DISTINCTION THIS FILE NOW EXISTS TO PROTECT: "Excel" alone means the
DESKTOP product, which we do not run. "Excel for the web" is a different
application with its own calculation engine, which we DO run. An executed
Excel-web value is evidence about the web engine and about nothing else; it is
never evidence about desktop Excel, and rendering it as though it were is the
single failure mode this integration could introduce. Because we have no
desktop run, a web-vs-documented mismatch is genuinely ambiguous -- it could be
the web engine diverging or the documentation being wrong about both -- and no
page may resolve it in either direction.

So this script fails (exit 1) on eight kinds of dishonesty:

  1. FALSE EXECUTION CLAIMS -- any page claiming DESKTOP Excel was verified,
     tested or executed. Truthful claims about executing Google Sheets,
     LibreOffice and/or Excel for the web are allowed, because they are true.
     Negations ("we do not run desktop Excel") are allowed.

     NARROWED for Excel for the web: "executed in Excel for the web" is now a
     TRUE sentence and must not be caught, while a bare "executed in Excel"
     stays fatal. The distinction is carried by negative lookaheads on the word
     "Excel" -- deliberately, so that adding the web engine could not be done
     by weakening the guard into "any sentence mentioning Excel and execution
     is fine".

     1c. STALE "WE DO NOT RUN EXCEL" COPY -- the inverse guard, and the reason
     it exists is that this sentence has just changed truth value. Until
     2026-09-01 "we do not run Excel" was the honest disclaimer on 32 guides.
     It is now ambiguous at best: we still do not run desktop Excel, but we DO
     run Excel for the web, and a page that says both without qualification
     contradicts itself. An unqualified "we do not run Excel" is therefore
     failed; "we do not run DESKTOP Excel" is the required form.

     1b. FABRICATED LIBREOFFICE / SHEETS EXECUTION -- the same lie about an
     engine we DO execute, on a function that engine never ran. Excel is
     always-false and so can be matched on wording alone; LibreOffice and
     Sheets are usually-true, so this half is derived FROM THE RESULTS FILES
     the way rule 5 is: a function page whose results carry no LibreOffice
     entry must not contain LibreOffice execution copy (a "Yes (<version>,
     <date>)" live-tested cell, or prose like "pass in LibreOffice"), and
     symmetrically for Google Sheets, and for Excel for the web (the "xw"
     slot: a page may claim an Excel-web live-tested cell only if
     results/excel-web.json actually has an entry for that function -- which
     is what keeps the seven LAMBDA-family transport skips from silently
     acquiring a web verdict). Until batch G every executed function
     had run in BOTH engines, so nothing exercised the single-engine paths in
     build_site.py -- and one of them fabricated "All 0 executed test cases
     pass in LibreOffice" for a Sheets-only function, with the version
     rendered as None. Rule 1's wording-only patterns could never catch that,
     because the sentence is true for 519 other functions.
  2. STALE "NOT YET EXECUTED" COPY -- any page still telling readers Google
     Sheets has not been executed. That was true until 2026-08-29 and is now a
     lie in the other direction. (Saying a specific CASE is "Not executed",
     or that a Sheets result is "inconclusive", is fine -- both are still
     true for individual cases; what must not survive is a blanket claim that
     Sheets has not been run.)
  3. MIS-ATTRIBUTED ENGINE-SCOPED RECIPE CHECKS -- a how-to page showing a
     "Google Sheets alternative (executed ...)" row without a real executed
     Sheets run behind it, or with a LibreOffice value where that row's
     LibreOffice column must read "n/a (Sheets-only formula)". A check may
     now be scoped to one engine (`"engines": ["google_sheets"]` in
     data/recipes/*.json), and the whole point of the scoping is that the
     other engine never ran the formula -- so the page must never show a
     number for it under the other engine's heading.
  5. FILE-LEVEL DATING OF PER-FUNCTION RESULTS -- a function page whose
     "Last tested" date is not the date that function was actually executed.
     Each results file records a per-function `executed_at`
     (harness/results_schema.py); the file-level `generated_at` is refreshed
     by every subset run, so dating pages from it silently re-dates the whole
     corpus every time five functions are re-run. Checked three ways per
     page: the "Last tested" line must equal the newest executed_at across
     the executed engines, and the support matrix's per-engine "Live-tested"
     cells must equal that engine's own executed_at for that function.
  4. SELF-CONTRADICTORY RECIPE COPY -- a how-to page that shows an executed
     Google Sheets result AND still carries the LibreOffice-only disclaimer.
     The how-to recipe corpus is executed in Sheets per-recipe (the
     multi-sheet recipes are skipped and stay LibreOffice-only), so the two
     wordings coexist ACROSS the site by design -- but never on one page.

  7. DESKTOP/WEB CONFLATION -- an Excel-for-the-web VALUE rendered under a
     heading that names Excel without saying "for the web". This is the guard
     that makes the distinction above enforceable rather than aspirational: it
     does not care what a page promises in prose, only whether web-executed
     data ever appears beneath a heading a reader would take to mean desktop
     Excel. The markers it keys on ("Yes (recalc, <date>)" and "OneDrive
     recalculation") are emitted by build_site.py for the excel_web engine and
     by nothing else.

  6. UNDECLARED EXCLUSIONS FROM A SHEETS VERDICT -- a recipe whose stored
     Google Sheets result carries `n_not_comparable > 0` has checks its
     Sheets badge does NOT cover: Google executed them against a workbook the
     LibreOffice reference run never saw (observed: the Drive importer
     renamed a data tab). The value is real and is still shown, so the page
     must SAY which checks the verdict leaves out and why. A page that
     silently drops them reads as full coverage, which is the quiet kind of
     dishonesty this file exists to catch.

Usage: python3 scripts/check_honesty.py [docs_dir]
"""
import re, sys, glob, os, html

ROOT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(__file__), "..", "docs")

# (1) Claims that an engine we do NOT execute was executed.
FALSE_EXECUTION = re.compile(
    r"(verified (?:in|across|on) all three|tested (?:in|across|on) all three|"
    r"executed (?:in|across|on) all three|"
    # NOT_WEB: the web build of Excel IS executed, so "executed in Excel for
    # the web" / "Excel Online" must survive while a bare "executed in Excel"
    # stays fatal. Attached to every pattern that names Excel as the object of
    # an execution verb.
    r"(?:verified|tested|executed|confirmed) (?:in|on) excel(?! *\(documented| for the web| online)|"
    r"verified formula for excel|"
    r"execution-verified compatibility data(?! *\((?:libreoffice|google sheets))|"
    r"every result (?:is )?verified(?! in (?:libreoffice|google sheets))|"
    r"machine-verified excel(?! for the web| online)|"
    # "in every version of Excel, Google Sheets and LibreOffice we test" — the
    # sneakiest form: it never uses the word "executed", but "we test" applied
    # to a list containing Excel claims exactly that. (Caught a real regression
    # in data/comparisons/*.json on 2026-08-29.)
    r"(?:every |all )?versions? of excel[^.]{0,80}?we (?:test|ran|run|execute)|"
    # NB: a bare "works in all Excel versions" is a documented-AVAILABILITY
    # claim, not a testing claim, so it is deliberately not matched here.
    r"we (?:tested|executed) (?:it )?in all excel versions|"
    r"we (?:test|execute|ran|run)[^.]{0,40}?\bin excel\b(?! for the web| online)|"
    # "executed in LibreOffice Calc, Google Sheets and Excel" -- the form the
    # two-engine recipe wording opens up. Once a sentence can honestly list
    # TWO executed engines, appending a third is a one-word regression none
    # of the patterns above would catch. Deliberately narrow: between the
    # verb and the trailing "and Excel" only ENGINE NAMES and separators may
    # appear, so this fires on a true engine list and not on the honest
    # "...executed in Sheets via Drive import (...), and Excel verdicts from
    # Microsoft's documentation", where the words in between are prose.
    r"(?:executed|verified|tested|ran|run) (?:in|on|across) "
    r"(?:(?:microsoft|google|libreoffice|calc|sheets|excel)[,&\s]+){1,4}"
    r"and (?:in )?(?:microsoft )?excel\b)",
    re.I,
)
NEGATION = re.compile(
    r"(not|never|haven't|have not|didn't|did not|can't|cannot|nor|rather than|"
    r"instead of|without|do not|don't)\W+(?:\w+\W+){0,4}$",
    re.I,
)

# (1c) "we do not run Excel" -- true about the desktop product, false as
# written now that Excel for the web is executed, and flatly self-contradictory
# on a page that also reports an Excel-web result. The required form names the
# product: "we do not run DESKTOP Excel". The lookahead lets the qualified
# forms through and nothing else, so this fires on exactly the copy the
# integration has to replace and stops firing when it has been.
STALE_EXCEL_DISCLAIMER = re.compile(
    r"\bwe (?:do|did|does)(?:n't| not) (?:run|execute)\s+"
    r"(?!desktop\b)(?:microsoft\s+)?excel\b(?! for the web| online)",
    re.I,
)

# (7) Desktop/web conflation. These two markers are emitted by
# site/build_site.py for the excel_web engine ONLY -- engine_tested_cell()
# renders "Yes (recalc, <date>)" and engine_exec_header() renders "... via
# OneDrive recalculation" -- so their presence in a section is proof that
# section carries web-executed data. A heading that says "Excel" without
# "for the web" (or "Online") above such data is the conflation.
XW_VALUE_MARKERS = re.compile(
    r"Yes \(recalc, \d{4}-\d{2}-\d{2}\)|OneDrive recalculation", re.I)
HEADING_RX = re.compile(r"<h[1-6][^>]*>(.*?)</h[1-6]>", re.S | re.I)
EXCEL_HEADING = re.compile(r"\bexcel\b", re.I)
WEB_QUALIFIER = re.compile(r"for the web|online", re.I)


def conflating_sections(raw):
    """Headings that name Excel without qualifying it, over Excel-web data.

    Splits the raw HTML at every heading and pairs each heading with the body
    that follows it, up to the next heading. Deliberately operates on the RAW
    html rather than stripped text, because the heading structure is the whole
    point of the check."""
    parts = list(HEADING_RX.finditer(raw))
    out = []
    for i, m in enumerate(parts):
        head = strip_tags(m.group(1)).strip()
        body = raw[m.end():parts[i + 1].start() if i + 1 < len(parts) else len(raw)]
        if not EXCEL_HEADING.search(head) or WEB_QUALIFIER.search(head):
            continue
        # The heading itself is searched too, not just the body. The most
        # likely way this rule ever fires is someone editing
        # engine_exec_header() so the Excel-web case table is titled "Excel
        # (executed ... via OneDrive recalculation)" -- there the giveaway
        # marker is IN the heading, and a body-only search sails past it.
        hit = XW_VALUE_MARKERS.search(m.group(0) + body)
        if hit:
            out.append((re.sub(r"\s+", " ", head)[:90], hit.group(0)))
    return out


# (2) Blanket "Google Sheets has not been executed" copy that is now stale.
# Matches "Google Sheets ... not (yet) executed/run/tested" within a short
# window, in either order, so the old phrasings ("Google Sheets is not yet run
# through our harness", "We have not yet executed this case in Google Sheets",
# "none has been executed in Sheets by us") are all caught.
STALE_SHEETS = [
    re.compile(
        # The verb may not be the tail of a hyphenated compound: the how-to
        # index renders per-recipe badges as "Sheets-executed", and a recipe
        # TITLE containing "not" ("How to sum where another column is not
        # blank") sitting between two such badges is not a claim about
        # anything -- "Sheets-executed ... not ... Sheets-executed" is three
        # unrelated tokens. A real stale claim reads "has not been executed",
        # never "-executed", so requiring a non-hyphen before the verb drops
        # the false positive without weakening the guard.
        r"(?:google )?sheets\b[^.]{0,80}?\b(?:not|never)\b[^.]{0,40}?"
        r"(?<!-)\b(?:yet )?(?:executed|run|tested|put through|been run)\b",
        re.I,
    ),
    # "we have not (yet) executed ... in Sheets". Anchored on an auxiliary verb
    # so the function name NOT ("the NOT function: executed in Google Sheets")
    # cannot masquerade as a negation.
    re.compile(
        r"\b(?:have|has|had|is|are|was|were|do|does|did|been|we)\s+not\b"
        r"[^.]{0,60}?\b(?:yet\s+)?(?:executed|run|tested)\b"
        r"[^.]{0,40}?\b(?:in|through|by|for)\s+(?:google\s+)?sheets\b",
        re.I,
    ),
    re.compile(
        r"\bnot\s+yet\s+(?:executed|run|tested)\b[^.]{0,40}?\b(?:google\s+)?sheets\b",
        re.I,
    ),
]
# Per-CASE honesty that is still true and must NOT trip the stale check.
STALE_ALLOW = re.compile(
    r"(inconclusive|no corpus case|not in our (?:executed )?(?:test )?set|"
    # Narrowed with the copy rewrite: the bare form is now caught by rule 1c,
    # so whitelisting it here would let a stale page hide behind it.
    r"we do not run desktop excel|desktop excel (?:is )?not (?:live-)?executed|"
    # A recipe page saying its own worked EXAMPLE was not run in Sheets is
    # true and must stay: recipes-verified.json is LibreOffice-only. Only the
    # blanket "Sheets has never been executed" claim is stale.
    r"recipe formulas|recipe example)",
    re.I,
)


def strip_tags(s):
    s = re.sub(r"<script.*?</script>|<style.*?</style>", " ", s, flags=re.S)
    return html.unescape(re.sub(r"<[^>]+>", " ", s))


# (3) Engine-scoped ("Sheets-only") recipe checks. Each such row renders the
# alternative's executed Sheets value AND, in the LibreOffice column, the
# literal "n/a (Sheets-only formula)" -- never a LibreOffice number, because
# scripts/verify_recipes.py skips the check entirely (see
# harness/recipe_corpus.check_in_scope). Two invariants follow, and both are
# cheap to check on the rendered text:
#
#   * the two markers appear the same number of times on a page: one n/a cell
#     per alternative row. Fewer n/a cells than alternative labels means a
#     LibreOffice value leaked into a row LibreOffice never executed.
#   * a page claiming "Google Sheets alternative (executed <date>)" must also
#     carry the page's real Sheets-execution provenance. build_site.py hides
#     an alternative row until an ingest has given it a value, so the label
#     without the provenance would be a fabricated execution claim.
SHEETS_ALT_LABEL = re.compile(r"Google Sheets alternative \(executed\b", re.I)
SHEETS_ONLY_NA = re.compile(r"n/a \(Sheets-only formula\)", re.I)

# (4) A recipe page cannot both show an executed Google Sheets result and
# still tell the reader its recipe was never run in Sheets. Both wordings are
# generated by site/build_site.py's RECIPE_TMPL, on mutually exclusive
# branches -- this catches a future edit that lets them render together.
RECIPE_SHEETS_EXECUTED = re.compile(
    r"(?:Returned by Google Sheets \(executed|Verified in Google Sheets \(|"
    r"We then ran the same formulas in\s+Google Sheets)", re.I)
RECIPE_SHEETS_DISCLAIMER = re.compile(
    r"recipe(?:'s|&rsquo;s)? (?:worked example|corpus)[^.]{0,80}?"
    r"(?:LibreOffice only|has not been through Sheets)", re.I)

# (5) Per-function "Last tested" dates must come from that function's own
# executed_at, not from the results file's generated_at. Cheap to verify: read
# the results files, recompute what each page's date must be, compare.
RESULTS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "results")
LAST_TESTED = re.compile(r"Last tested (\d{4}-\d{2}-\d{2})")
# Support-matrix "Live-tested" cells, per engine (site/build_site.py's
# engine_tested_cell): "Yes (25.8.7.3, 2026-07-29)" / "Yes (Drive import, …)".
LO_CELL = re.compile(r"Yes \((\d+(?:\.\d+)+), (\d{4}-\d{2}-\d{2})\)")
GS_CELL = re.compile(r"Yes \(Drive import, (\d{4}-\d{2}-\d{2})\)")
# Excel for the web has no version either, so its cell is dated prose too.
# Keeping it non-numeric is what keeps it out of LO_CELL's jaws above.
XW_CELL = re.compile(r"Yes \(recalc, (\d{4}-\d{2}-\d{2})\)")


def _version_tuple(v):
    if not re.match(r"^\s*\d+(\.\d+)*\s*$", str(v or "")):
        return (0,)
    return tuple(int("".join(c for c in tok if c.isdigit()) or 0)
                 for tok in str(v).split("."))


def expected_dates():
    """FUNCTION -> {"lo": date|None, "gs": date|None, "page": newest date}.

    Mirrors site/build_site.py: the NEWEST LibreOffice build wins the live
    verdict, Google Sheets is its own engine, and the page's "Last tested"
    line is the newest date the function was executed on in either."""
    import json as _json
    blobs, newest_lo = [], None
    for path in sorted(glob.glob(os.path.join(RESULTS_DIR, "*.json"))):
        try:
            d = _json.load(open(path, encoding="utf-8"))
        except Exception:
            continue
        engine = str(d.get("engine", "")).lower()
        if "libreoffice" in engine:
            if newest_lo is None or _version_tuple(d.get("engine_version")) >= \
                    _version_tuple(newest_lo.get("engine_version")):
                newest_lo = d
        # ORDER MATTERS, exactly as in build_site.engine_key_from_engine_name:
        # "excel_web" must be recognised before any bare "excel" test, or the
        # web run lands in the desktop slot. It was silently DROPPED here until
        # the web engine was published, which was safe only while nothing
        # rendered it -- the moment it reaches a page, a page's "Last tested"
        # line has to be the newest date across THREE engines, not two.
        elif "excel_web" in engine or "excel for the web" in engine:
            blobs.append(d)
        elif "sheet" in engine or "google" in engine:
            blobs.append(d)
    if newest_lo is not None:
        blobs.append(newest_lo)
    out = {}
    for d in blobs:
        _eng = str(d.get("engine", "")).lower()
        slot = ("lo" if "libreoffice" in _eng
                else "xw" if ("excel_web" in _eng or "excel for the web" in _eng)
                else "gs")
        fallback = (d.get("generated_at") or "")[:10]
        for fn, block in (d.get("function_results") or {}).items():
            date = (block.get("executed_at") if isinstance(block, dict) else None) or fallback
            e = out.setdefault(fn, {"lo": None, "gs": None, "xw": None, "page": ""})
            e[slot] = date
            if date and date > e["page"]:
                e["page"] = date
    return out


# (1b) Execution claims about an engine that did not run THIS function.
# High-signal patterns only: each must name the engine AND assert execution.
LO_EXEC_PROSE = re.compile(
    r"\b(?:executed|execution|ran|run|tested)\b[^.]{0,60}?\bin libreoffice\b"
    r"|\bpass(?:es|ed)?\b[^.]{0,40}?\bin libreoffice\b"
    r"|\blibreoffice\b[^.]{0,60}?\bexecuted (?:test )?cases?\b",
    re.I,
)
GS_EXEC_PROSE = re.compile(
    r"\b(?:executed|execution|ran|run|tested)\b[^.]{0,60}?\bin (?:google )?sheets\b"
    r"|\bpass(?:es|ed)?\b[^.]{0,40}?\bin (?:google )?sheets\b"
    r"|\b(?:google )?sheets\b[^.]{0,60}?\bexecuted (?:test )?cases?\b",
    re.I,
)


XW_EXEC_PROSE = re.compile(
    r"\b(?:executed|execution|ran|run|tested)\b[^.]{0,60}?\bin excel for the web\b"
    r"|\bpass(?:es|ed)?\b[^.]{0,40}?\bin excel for the web\b"
    r"|\bexcel for the web\b[^.]{0,60}?\bexecuted (?:test )?cases?\b",
    re.I,
)


# The regions of a function page that make claims about THAT function: the
# title, the meta description, the lede and the support matrix -- i.e. exactly
# what build_site.py renders from the function's own record. Deliberately NOT
# the whole page: the "Where <NAME> behaves differently" section and the
# related-recipe cards quote executed results for OTHER functions, and those
# citations are true. (Observed: functions/sortn.html legitimately quotes
# SORT's executed Sheets-vs-LibreOffice values, which a whole-page scan flags.)
OWN_CLAIM_REGIONS = [
    re.compile(r"<title>(.*?)</title>", re.S | re.I),
    re.compile(r'<meta name="description" content="(.*?)"', re.S | re.I),
    re.compile(r'<p class="lede">(.*?)</p>', re.S | re.I),
    re.compile(r'<table class="matrix">(.*?)</table>', re.S | re.I),
]


def own_claims(raw):
    """Raw HTML -> plain text of just this function's own claim regions."""
    return "\n".join(
        strip_tags(m.group(1))
        for rx in OWN_CLAIM_REGIONS for m in [rx.search(raw)] if m
    )


def executed_engines():
    """FUNCTION -> set of engine slots ("lo" / "gs") that actually ran it.

    Read from results/*.json, never from the page, so the guard cannot be
    satisfied by copy that merely looks right -- the same discipline rule 5
    and rule 6 use."""
    import json as _json
    out = {}
    for path in sorted(glob.glob(os.path.join(RESULTS_DIR, "*.json"))):
        try:
            d = _json.load(open(path, encoding="utf-8"))
        except Exception:
            continue
        engine = str(d.get("engine", "")).lower()
        if "libreoffice" in engine:
            slot = "lo"
        # Before any "excel" test, for the same reason as everywhere else.
        elif "excel_web" in engine or "excel for the web" in engine:
            slot = "xw"
        elif "sheet" in engine or "google" in engine:
            slot = "gs"
        else:
            continue
        for fn in (d.get("function_results") or {}):
            out.setdefault(fn, set()).add(slot)
    return out


# (6) Recipes whose Sheets verdict excludes checks must declare it on the page.
def recipes_with_exclusions():
    """{slug: (n_comparable, n_total, n_excluded)} from the Sheets recipe run.

    Read from the RESULTS file, not from the page, so the guard cannot be
    satisfied by copy that merely looks right: the numbers it demands are the
    ones the harness actually recorded."""
    import json as _json
    path = os.path.join(RESULTS_DIR, "recipes-verified-sheets.json")
    if not os.path.exists(path):
        return {}
    try:
        d = _json.load(open(path, encoding="utf-8"))
    except Exception:
        return {}
    if not d.get("trusted"):
        return {}          # an untrusted run is not published; nothing to declare
    out = {}
    for slug, rec in (d.get("recipes") or {}).items():
        n_excl = int(rec.get("n_not_comparable") or 0)
        n_total = int(rec.get("n_checks") or 0)
        if n_excl and n_total:
            out[slug] = (n_total - n_excl, n_total, n_excl)
    return out


RECIPE_EXCLUSIONS = recipes_with_exclusions()
undeclared = []

stale_excel = []
conflated = []
bad = []
fabricated = []
stale = []
misdated = []
contradictory = []
misattributed = []
files = glob.glob(os.path.join(ROOT, "**", "*.html"), recursive=True)
for f in files:
    _raw_page = open(f, encoding="utf-8", errors="replace").read()
    text = strip_tags(_raw_page)
    rel = os.path.relpath(f, ROOT)
    # (1c) the disclaimer that changed truth value on 2026-09-01.
    for m in STALE_EXCEL_DISCLAIMER.finditer(text):
        stale_excel.append(
            (rel, re.sub(r"\s+", " ",
                         text[max(0, m.start() - 70):m.end() + 70])))
    # (7) web values under a desktop-sounding heading. Runs on the RAW html
    # because it is the heading STRUCTURE that is being checked.
    for _head, _marker in conflating_sections(_raw_page):
        conflated.append(
            (rel, f"heading {_head!r} names Excel without saying \"for the web\", "
                  f"and the section under it renders Excel-for-the-web executed "
                  f"data ({_marker!r})"))
    slug = os.path.basename(f)[:-5]
    if rel.replace(os.sep, "/").startswith("how-to/") and slug in RECIPE_EXCLUSIONS:
        n_ok, n_all, n_excl = RECIPE_EXCLUSIONS[slug]
        want = (f"verified over {n_ok} of {n_all} checks \u2014 {n_excl} excluded "
                f"as not comparable")
        norm = re.sub(r"\s+", " ", text)
        if want.lower() not in norm.lower():
            undeclared.append(
                (rel, f"stored Sheets result excludes {n_excl} of {n_all} check(s) "
                      f"as not comparable, but the page does not say so "
                      f"(expected the line: {want!r})"))
    n_alt = len(SHEETS_ALT_LABEL.findall(text))
    n_na = len(SHEETS_ONLY_NA.findall(text))
    if n_alt != n_na:
        misattributed.append(
            (rel, f"{n_alt} 'Google Sheets alternative (executed …)' label(s) but "
                  f"{n_na} 'n/a (Sheets-only formula)' cell(s) -- every Sheets-only "
                  f"row must read n/a in the LibreOffice column"))
    elif n_alt and not RECIPE_SHEETS_EXECUTED.search(text):
        misattributed.append(
            (rel, "claims a 'Google Sheets alternative (executed …)' but carries no "
                  "Google Sheets execution provenance"))
    if RECIPE_SHEETS_EXECUTED.search(text) and RECIPE_SHEETS_DISCLAIMER.search(text):
        m = RECIPE_SHEETS_DISCLAIMER.search(text)
        contradictory.append(
            (rel, re.sub(r"\s+", " ", text[max(0, m.start() - 60):m.end() + 60])))
    for m in FALSE_EXECUTION.finditer(text):
        before = text[max(0, m.start() - 60):m.start()]
        if NEGATION.search(before):
            continue
        bad.append((rel, text[max(0, m.start() - 50):m.end() + 40].replace("\n", " ")))
    for rx in STALE_SHEETS:
        for m in rx.finditer(text):
            ctx = text[max(0, m.start() - 90):m.end() + 90].replace("\n", " ")
            if STALE_ALLOW.search(ctx):
                continue
            stale.append((rel, re.sub(r"\s+", " ", ctx)))

# (5) per-function "Last tested" dates
# (1b) fabricated LibreOffice / Sheets execution claims -- same loop, but over
# EVERY function page, including the ones no engine has executed (which is
# exactly where a fabricated claim would be invisible to rule 5).
_expected = expected_dates()
_ran = executed_engines()
for _page in sorted(glob.glob(os.path.join(ROOT, "functions", "*.html"))):
    _fn_slug = os.path.basename(_page)[:-5]
    _rel = os.path.relpath(_page, ROOT)
    _raw = open(_page, encoding="utf-8", errors="replace").read()
    _txt = own_claims(_raw)
    _slots = next((v for k, v in _ran.items() if k.lower() == _fn_slug), set())
    for _slot, _label, _cell_rx, _prose_rx in (
        ("lo", "LibreOffice", LO_CELL, LO_EXEC_PROSE),
        ("gs", "Google Sheets", GS_CELL, GS_EXEC_PROSE),
        # The seven LAMBDA-family functions Excel for the web could not open
        # are exactly what this slot protects: they have LibreOffice and Sheets
        # entries but no excel-web entry, so any web live-tested cell or web
        # execution prose on their pages is fabricated by definition.
        ("xw", "Excel for the web", XW_CELL, XW_EXEC_PROSE),
    ):
        if _slot in _slots:
            continue                      # that engine really did run it
        if _cell_rx.search(_txt):
            fabricated.append(
                (_rel, f"support matrix shows a {_label} live-tested cell, but no "
                       f"{_label} results file has an entry for this function"))
        m = _prose_rx.search(_txt)
        if m:
            fabricated.append(
                (_rel, f"claims {_label} execution but no {_label} results file has "
                       f"an entry for this function: "
                       f"…{re.sub(chr(92) + 's+', ' ', _txt[max(0, m.start() - 50):m.end() + 40])}…"))

for fn, want in sorted(_expected.items()):
    page = os.path.join(ROOT, "functions", fn.lower() + ".html")
    if not os.path.exists(page):
        continue
    rel_page = os.path.relpath(page, ROOT)
    raw = open(page, encoding="utf-8", errors="replace").read()
    m = LAST_TESTED.search(raw)
    if not m:
        misdated.append((rel_page,
                         f"executed on {want['page']} but the page shows no "
                         f"'Last tested' date"))
    elif m.group(1) != want["page"]:
        misdated.append((rel_page,
                         f"page says 'Last tested {m.group(1)}' but {fn} was last "
                         f"executed {want['page']} (results files' executed_at)"))
    for label, got, exp in (
        ("LibreOffice", [d for _v, d in LO_CELL.findall(raw)], want["lo"]),
        ("Google Sheets", GS_CELL.findall(raw), want["gs"]),
        ("Excel for the web", XW_CELL.findall(raw), want.get("xw")),
    ):
        if exp and got and got[0] != exp:
            misdated.append((rel_page,
                             f"{label} live-tested cell says {got[0]} but {fn} was "
                             f"executed on {exp} in that engine's results file"))

print(f"honesty check: {len(files)} pages")
print(f"  false execution claims (Excel / all three): {len(bad)}")
for f, ctx in bad[:40]:
    print(f"    {f}: …{ctx}…")
print(f"  fabricated LibreOffice/Sheets execution claims: {len(fabricated)}")
for f, ctx in fabricated[:40]:
    print(f"    {f}: {ctx}")
print(f"  stale 'Google Sheets not yet executed' copy: {len(stale)}")
for f, ctx in stale[:40]:
    print(f"    {f}: …{ctx}…")
print(f"  Sheets-only rows mis-attributed to LibreOffice: {len(misattributed)}")
for f, ctx in misattributed[:40]:
    print(f"    {f}: {ctx}")
print(f"  recipe pages both showing and denying a Sheets result: {len(contradictory)}")
for f, ctx in contradictory[:40]:
    print(f"    {f}: …{ctx}…")
print(f"  stale unqualified 'we do not run Excel' copy: {len(stale_excel)}")
for f, ctx in stale_excel[:40]:
    print(f"    {f}: \u2026{ctx}\u2026")
print(f"  Excel-web values under a desktop-Excel heading: {len(conflated)}")
for f, ctx in conflated[:40]:
    print(f"    {f}: {ctx}")
print(f"  recipe pages hiding checks their Sheets verdict excludes: {len(undeclared)}")
for f, ctx in undeclared[:10]:
    print(f"    {f}: {ctx}")
print(f"  function pages dated from the file instead of the function: {len(misdated)}"
      f"  ({len(_expected)} executed functions checked)")
for f, ctx in misdated[:40]:
    print(f"    {f}: {ctx}")
sys.exit(1 if (bad or fabricated or stale or contradictory or misattributed
               or misdated or undeclared or stale_excel or conflated) else 0)
