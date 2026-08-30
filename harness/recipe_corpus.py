"""
Shared how-to RECIPE corpus machinery, used by every engine that executes it.

WHY THIS MODULE EXISTS
----------------------
The 282 how-to recipes in `data/recipes/*.json` are a second, separate corpus
from the FUNCTION corpus in `data/tests/*.json` (which `harness/corpus.py`
serves). Two runners now execute the recipes:

  * `scripts/verify_recipes.py`  -- headless LibreOffice, the original path
  * `harness/run_sheets.py build-recipes / ingest-recipes`
                                 -- Google Sheets, via the Drive-import
                                    round-trip

If those two enumerated the checks differently, resolved a variant's
inherited `setup_cells` differently, or compared a read-back value to
`expected` under even slightly different rules, then a reported
"LibreOffice and Google Sheets disagree" would be an artifact of the harness
rather than a real engine difference -- exactly the failure mode
`harness/corpus.py` exists to prevent for the function corpus.

So the corpus-shaped parts live here, once:

  * `load_recipe_files()`  -- which recipes exist, in which order
  * `iter_checks()`        -- what the checks ARE (main example + every
                              variant check) and how a variant check inherits
                              `setup_cells` / `setup_sheets`
  * `norm()`               -- read-back value normalization
  * `compare_check()`      -- expected-vs-actual verdict
  * `run_check()`          -- execute one check through a caller-supplied
                              engine `runner` callable

Nothing here knows how to make an engine calculate. `run_check()` takes the
engine as a callback, so LibreOffice's `soffice --convert-to` and Google
Sheets' Drive-import readback plug into identical surrounding logic.

REFACTOR PROVENANCE (behaviour-preserving)
------------------------------------------
`norm`, the check-resolution rules, and the comparison rule below were lifted
VERBATIM out of `scripts/verify_recipes.py`, which now imports them from here.
In particular `compare_check()` keeps that script's exact semantics, quirks
included:

  * scalars compare with plain `==` (no float tolerance -- deliberately
    unlike the function corpus, whose `expected` values include computed
    floats; recipe expectations are authored exact),
  * list expectations compare `[str(x) for x in actual] == [str(x) for x in
    expected]`, i.e. stringified, so `3` and `3.0` match after `norm()`,
  * a check dict's own `setup_cells` / `setup_sheets` key SHADOWS the
    variant's, and an explicit `null` in the JSON means "no setup", not
    "inherit" (`dict.get(key, default)`, not `dict.get(key) or default`).

Do not "clean up" any of those without re-running the full LibreOffice pass
and diffing `results/recipes-verified.json`.
"""
import glob
import json
import os

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
RECIPES_DIR = os.path.join(REPO_ROOT, "data", "recipes")

# Anchor cell for a check with no `check_range`. Verbatim from
# verify_recipes.py's run_case(); the Sheets builder must use the same one or
# ingest would read an empty cell.
SCALAR_ANCHOR = "H1"


def load_recipe_files(only=None):
    """Return [(slug, path, recipe_dict)] in the order the runners execute.

    Sorted by FILENAME (not by slug) because that is what verify_recipes.py
    has always done -- `sorted(glob.glob(...))`. Every recipe file in this
    corpus is named `<slug>.json`, so the two orders coincide today; sorting
    by filename keeps them coincident if that ever stops being true.
    """
    out = []
    for path in sorted(glob.glob(os.path.join(RECIPES_DIR, "*.json"))):
        with open(path) as f:
            recipe = json.load(f)
        if only and recipe["slug"] not in only:
            continue
        out.append((recipe["slug"], path, recipe))
    return out


def norm(v):
    """Normalize a read-back cell value. Verbatim from verify_recipes.py."""
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


def resolve_setup(check, default_setup=None, default_sheets=None):
    """Resolve one check's effective (setup_cells, setup_sheets).

    `dict.get(key, default)` -- NOT `or` -- so an explicit `"setup_cells":
    null` in a check means "this check has no setup cells", overriding the
    variant's, rather than silently inheriting them. verify_recipes.py has
    always done exactly this; recipe JSON depends on it (see
    reference-a-cell-on-another-sheet, whose INDIRECT variant overrides A1).
    """
    return (check.get("setup_cells", default_setup),
            check.get("setup_sheets", default_sheets))


def anchor_for(check_range):
    """Cell the formula is written to. Verbatim from verify_recipes.run_case."""
    return check_range.split(":")[0] if check_range else SCALAR_ANCHOR


def iter_checks(recipe):
    """Yield one dict per executable check, in verify_recipes.py's exact order:
    the recipe's main `verify` example first, then every variant's checks in
    order. A variant's `verify` may be a single dict or a list of them.

    Each yielded dict carries the RESOLVED setup, so a consumer never has to
    re-implement the inheritance rules:

        key            stable id, "main" or "v<vi>c<ci>"
        kind           "main" | "variant"
        variant_index  index into recipe["variants"], or None for the main one
        check_index    index within that variant's checks, or None
        heading        the variant's heading ("" for the main example)
        label          the check's own label ("" for the main example)
        formula, expected, check_range, setup_cells, setup_sheets
        anchor         cell the formula goes in
    """
    v = recipe["verify"]
    setup, sheets = resolve_setup(v)
    yield {
        "key": "main", "kind": "main",
        "variant_index": None, "check_index": None,
        "heading": "", "label": "",
        "formula": v["formula"], "expected": v["expected"],
        "check_range": v.get("check_range"),
        "setup_cells": setup, "setup_sheets": sheets,
        "anchor": anchor_for(v.get("check_range")),
    }
    for vi, var in enumerate(recipe.get("variants") or []):
        checks = var.get("verify") or []
        if isinstance(checks, dict):
            checks = [checks]
        for ci, c in enumerate(checks):
            setup, sheets = resolve_setup(
                c, var.get("setup_cells"), var.get("setup_sheets"))
            yield {
                "key": f"v{vi}c{ci}", "kind": "variant",
                "variant_index": vi, "check_index": ci,
                "heading": var.get("heading", ""), "label": c.get("label", ""),
                "formula": c["formula"], "expected": c["expected"],
                "check_range": c.get("check_range"),
                "setup_cells": setup, "setup_sheets": sheets,
                "anchor": anchor_for(c.get("check_range")),
            }


def uses_setup_sheets(recipe):
    """True if ANY of this recipe's checks needs extra worksheets to exist."""
    return any(bool(c["setup_sheets"]) for c in iter_checks(recipe))


def setup_sheet_names(recipe):
    """Every extra worksheet name this recipe's checks ask for, sorted."""
    names = set()
    for c in iter_checks(recipe):
        names.update((c["setup_sheets"] or {}).keys())
    return sorted(names)


def compare_check(expected, actual):
    """expected-vs-actual verdict. Verbatim from verify_recipes.py's check():

        ok = (actual==exp) if not isinstance(exp,list) else (
                 [str(x) for x in actual]==[str(x) for x in exp])
    """
    if isinstance(expected, list):
        # NOT wrapped in try/except here on purpose: in the original this
        # comparison sat INSIDE check()'s try, so a non-iterable `actual`
        # against a list `expected` raises TypeError and is recorded as
        # actual="ERR ...", ok=False -- not merely as a mismatch. run_check()
        # (and every other caller) must preserve that by catching around it.
        return [str(x) for x in actual] == [str(x) for x in expected]
    return actual == expected


def run_check(check, runner, default_setup=None, default_sheets=None):
    """Execute one check dict through `runner`; returns (actual, ok).

    `runner(setup_cells, formula, check_range, setup_sheets) -> actual`.
    Exactly verify_recipes.py's original `check()` body, with the LibreOffice
    call swapped for the callback -- including catching ANY exception from the
    runner and recording it as the string `"ERR <exception>"` with ok=False,
    so one broken recipe cannot abort a whole corpus run.
    """
    expected = check["expected"]
    cr = check.get("check_range")
    setup, sheets = resolve_setup(check, default_setup, default_sheets)
    try:
        actual = runner(setup, check["formula"], cr, sheets)
        ok = compare_check(expected, actual)
    except Exception as e:  # noqa: BLE001 -- deliberate: report, never abort
        actual = f"ERR {e}"
        ok = False
    return actual, bool(ok)
