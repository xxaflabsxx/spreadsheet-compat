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
                              variant check + every top-level `extra_checks`
                              entry), how a check inherits `setup_cells` /
                              `setup_sheets`, which ENGINES it is scoped to,
                              and its STABLE KEY
  * `norm()`               -- read-back value normalization
  * `compare_check()`      -- expected-vs-actual verdict
  * `run_check()`          -- execute one check through a caller-supplied
                              engine `runner` callable
  * `result_checks_by_key()` -- flatten a results-file record to {key: payload}

ENGINE SCOPING AND STABLE KEYS
------------------------------
Any check (the main `verify`, a variant check, or an `extra_checks` entry)
may carry an optional `"engines"` list naming the engines that should execute
it, e.g. `"engines": ["google_sheets"]`. ABSENT OR EMPTY MEANS ALL ENGINES,
so every pre-existing check keeps running everywhere -- the field is purely
additive. `iter_checks(recipe, engine=...)` drops the checks that are out of
scope for that engine, and each runner passes its own engine name, so a
Sheets-only alternative formula never reaches LibreOffice (which would
execute it, fail it, and flip the recipe's LibreOffice badge).

Every check also has a STABLE KEY, used as the identity under which a result
is stored and merged:

    main verify           "main"
    variant vi, check ci  "v<vi>c<ci>"   (vi/ci are positions in the JSON,
                                          counted BEFORE any engine filter,
                                          so a filtered run and a full run
                                          agree on every key)
    extra_checks[i]       "x<i>"
    any check with "id"   that id, verbatim

Keys must be unique within a recipe; `iter_checks()` raises if they are not.
Results files (results/recipes-verified.json, results/recipes-verified-sheets.json)
store the key on every check payload, and consumers merge BY KEY rather than
by position -- so appending an engine-scoped check to a variant cannot shift
an older result onto the wrong formula. Files written before keys existed
still load: `result_checks_by_key()` derives the same keys positionally.

`extra_checks` is a top-level list of checks that belong to the recipe but
not to any variant (the natural home for "here is the alternative formula
the other engine needs"). Each entry inherits the MAIN example's
`setup_cells` / `setup_sheets` unless it overrides them, under the same
`dict.get(key, default)` shadowing rule variants use.

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

# Engine names usable in a check's "engines" list. These are the runner
# identities, not display labels: scripts/verify_recipes.py runs as
# "libreoffice" and harness/run_sheets.py's recipe commands run as
# "google_sheets".
LIBREOFFICE = "libreoffice"
GOOGLE_SHEETS = "google_sheets"
ALL_ENGINES = (LIBREOFFICE, GOOGLE_SHEETS)


def check_engines(check):
    """The engines a check is scoped to, or None meaning ALL engines.

    An absent (or empty) "engines" key is the default and means every engine
    executes the check -- that is what makes the field backward compatible
    with every recipe authored before it existed.
    """
    engines = check.get("engines")
    if not engines:
        return None
    return [str(e) for e in engines]


def check_in_scope(check, engine):
    """Should `engine` execute this check? `engine=None` means "no filter"."""
    if engine is None:
        return True
    scoped = check_engines(check)
    return scoped is None or engine in scoped


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
    # Readback artifact: Google Sheets' xlsx export applies a date/time number
    # format to some numeric results, so openpyxl hands back a datetime for what
    # is really a plain serial number. Convert back with openpyxl's own epoch
    # logic (it honours the 1900 leap-year bug for serials < 61) so a numeric
    # expectation compares against the original number.
    import datetime as _dt
    if isinstance(v, (_dt.datetime, _dt.date)):
        from openpyxl.utils.datetime import to_excel as _to_excel
        v = _to_excel(v)
        if abs(v - round(v)) < 1e-9:
            v = int(round(v))
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


def check_key(check, default_key):
    """A check's stable key: its explicit "id" if it has one, else the
    positional default ("main", "v<vi>c<ci>", "x<i>"). An explicit id lets a
    check be moved or reordered without orphaning its stored result."""
    return str(check.get("id") or default_key)


def iter_checks(recipe, engine=None):
    """Yield one dict per executable check, in the runners' exact order: the
    recipe's main `verify` example first, then every variant's checks in
    order, then every top-level `extra_checks` entry. A variant's `verify`
    may be a single dict or a list of them.

    `engine` (e.g. "libreoffice", "google_sheets") drops the checks that are
    not scoped to it; the default None yields every check, which is what the
    corpus-shape helpers (`uses_setup_sheets`, `setup_sheet_names`) and the
    site's merge want. Positional keys are computed BEFORE the filter, so a
    filtered run keys its results exactly as a full run does.

    Each yielded dict carries the RESOLVED setup, so a consumer never has to
    re-implement the inheritance rules:

        key            stable id, "main" / "v<vi>c<ci>" / "x<i>" / explicit "id"
        kind           "main" | "variant" | "extra"
        engines        list of engines it is scoped to, or None = all engines
        variant_index  index into recipe["variants"], or None
        check_index    index within that variant's checks, or None
        heading        the variant's heading ("" for main/extra checks)
        label          the check's own label ("" for the main example)
        formula, expected, check_range, setup_cells, setup_sheets
        anchor         cell the formula goes in
    """
    seen = set()

    def _key(check, default):
        k = check_key(check, default)
        if k in seen:
            raise ValueError(
                f"{recipe.get('slug', '?')}: duplicate check key {k!r} -- keys "
                f"identify a check's stored result, so they must be unique")
        seen.add(k)
        return k

    v = recipe["verify"]
    setup, sheets = resolve_setup(v)
    main_key = _key(v, "main")
    if check_in_scope(v, engine):
        yield {
            "key": main_key, "kind": "main",
            "engines": check_engines(v),
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
            k = _key(c, f"v{vi}c{ci}")
            if not check_in_scope(c, engine):
                continue
            setup, sheets = resolve_setup(
                c, var.get("setup_cells"), var.get("setup_sheets"))
            yield {
                "key": k, "kind": "variant",
                "engines": check_engines(c),
                "variant_index": vi, "check_index": ci,
                "heading": var.get("heading", ""), "label": c.get("label", ""),
                "formula": c["formula"], "expected": c["expected"],
                "check_range": c.get("check_range"),
                "setup_cells": setup, "setup_sheets": sheets,
                "anchor": anchor_for(c.get("check_range")),
            }
    # Top-level extra checks: executable formulas that belong to the recipe
    # but to no variant. They inherit the MAIN example's setup by default, so
    # "the same sample data, a different formula" needs no duplication.
    main_setup, main_sheets = resolve_setup(v)
    for xi, c in enumerate(recipe.get("extra_checks") or []):
        k = _key(c, f"x{xi}")
        if not check_in_scope(c, engine):
            continue
        setup, sheets = resolve_setup(c, main_setup, main_sheets)
        yield {
            "key": k, "kind": "extra",
            "engines": check_engines(c),
            "variant_index": None, "check_index": xi,
            "heading": "", "label": c.get("label", ""),
            "formula": c["formula"], "expected": c["expected"],
            "check_range": c.get("check_range"),
            "setup_cells": setup, "setup_sheets": sheets,
            "anchor": anchor_for(c.get("check_range")),
        }


def result_checks_by_key(record):
    """Flatten ONE recipe's results-file record into {key: check_payload}.

    Backward compatible on purpose: results files written before checks had
    keys carry no "key" field, so the key is derived from the payload's
    POSITION using exactly the same rule iter_checks() uses ("main",
    "v<vi>c<ci>", "x<i>"). An old file therefore merges identically to a
    freshly written one -- which is what lets the site keep rendering a
    Sheets run from before this change.

    The main check's payload IS the record itself (that is how both results
    files have always stored it), so the returned mapping aliases it.
    """
    out = {}
    if not record:
        return out
    out[record.get("key") or "main"] = record
    for vi, var in enumerate(record.get("variants") or []):
        for ci, ch in enumerate(var.get("checks") or []):
            out[ch.get("key") or f"v{vi}c{ci}"] = ch
    for xi, ch in enumerate(record.get("extra_checks") or []):
        out[ch.get("key") or f"x{xi}"] = ch
    return out


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


def run_check(check, runner, default_setup=None, default_sheets=None,
              engine=None):
    """Execute one check dict through `runner`; returns (actual, ok).

    `runner(setup_cells, formula, check_range, setup_sheets) -> actual`.
    Exactly verify_recipes.py's original `check()` body, with the LibreOffice
    call swapped for the callback -- including catching ANY exception from the
    runner and recording it as the string `"ERR <exception>"` with ok=False,
    so one broken recipe cannot abort a whole corpus run.

    `engine` is the name of the engine `runner` speaks for. If the check is
    not scoped to it (see `check_in_scope`), the runner is NOT called and
    `(None, None)` comes back -- ok=None, distinguishable from a failure,
    meaning "this engine does not execute this check". Callers that enumerate
    through `iter_checks(recipe, engine=...)` have already been filtered and
    will never see it; the guard is here so a caller cannot execute an
    out-of-scope formula by going straight to run_check().
    """
    if not check_in_scope(check, engine):
        return None, None
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
