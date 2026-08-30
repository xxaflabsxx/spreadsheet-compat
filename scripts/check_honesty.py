#!/usr/bin/env python3
"""Honesty guard for the built site.

Ground truth (see site/build_site.py and results/*.json):

  * LibreOffice Calc  -- EXECUTED (four pinned builds, results/libreoffice-*.json)
  * Google Sheets     -- EXECUTED (one dated Drive-import run,
                         results/google-sheets.json, 2026-08-29)
  * Microsoft Excel   -- NOT EXECUTED. Documentation only, and it is the
                         yardstick the two executed engines are measured against.

So this script fails (exit 1) on four kinds of dishonesty:

  1. FALSE EXECUTION CLAIMS -- any page claiming Excel was verified, tested or
     executed, or claiming all three engines were. Truthful claims about
     executing Google Sheets and/or LibreOffice are allowed, because they are
     true. Negations ("we do not run Excel") are allowed.
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
  4. SELF-CONTRADICTORY RECIPE COPY -- a how-to page that shows an executed
     Google Sheets result AND still carries the LibreOffice-only disclaimer.
     The how-to recipe corpus is executed in Sheets per-recipe (the
     multi-sheet recipes are skipped and stay LibreOffice-only), so the two
     wordings coexist ACROSS the site by design -- but never on one page.

Usage: python3 scripts/check_honesty.py [docs_dir]
"""
import re, sys, glob, os, html

ROOT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(__file__), "..", "docs")

# (1) Claims that an engine we do NOT execute was executed.
FALSE_EXECUTION = re.compile(
    r"(verified (?:in|across|on) all three|tested (?:in|across|on) all three|"
    r"executed (?:in|across|on) all three|"
    r"(?:verified|tested|executed|confirmed) (?:in|on) excel(?! *\(documented)|"
    r"verified formula for excel|"
    r"execution-verified compatibility data(?! *\((?:libreoffice|google sheets))|"
    r"every result (?:is )?verified(?! in (?:libreoffice|google sheets))|"
    r"machine-verified excel|"
    # "in every version of Excel, Google Sheets and LibreOffice we test" — the
    # sneakiest form: it never uses the word "executed", but "we test" applied
    # to a list containing Excel claims exactly that. (Caught a real regression
    # in data/comparisons/*.json on 2026-08-29.)
    r"(?:every |all )?versions? of excel[^.]{0,80}?we (?:test|ran|run|execute)|"
    # NB: a bare "works in all Excel versions" is a documented-AVAILABILITY
    # claim, not a testing claim, so it is deliberately not matched here.
    r"we (?:tested|executed) (?:it )?in all excel versions|"
    r"we (?:test|execute|ran|run)[^.]{0,40}?\bin excel\b|"
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
    r"we do not run excel|excel (?:is )?not (?:live-)?executed|"
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

bad = []
stale = []
contradictory = []
misattributed = []
files = glob.glob(os.path.join(ROOT, "**", "*.html"), recursive=True)
for f in files:
    text = strip_tags(open(f, encoding="utf-8", errors="replace").read())
    rel = os.path.relpath(f, ROOT)
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

print(f"honesty check: {len(files)} pages")
print(f"  false execution claims (Excel / all three): {len(bad)}")
for f, ctx in bad[:40]:
    print(f"    {f}: …{ctx}…")
print(f"  stale 'Google Sheets not yet executed' copy: {len(stale)}")
for f, ctx in stale[:40]:
    print(f"    {f}: …{ctx}…")
print(f"  Sheets-only rows mis-attributed to LibreOffice: {len(misattributed)}")
for f, ctx in misattributed[:40]:
    print(f"    {f}: {ctx}")
print(f"  recipe pages both showing and denying a Sheets result: {len(contradictory)}")
for f, ctx in contradictory[:40]:
    print(f"    {f}: …{ctx}…")
sys.exit(1 if (bad or stale or contradictory or misattributed) else 0)
