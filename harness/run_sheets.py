#!/usr/bin/env python3
"""
Google Sheets engine runner for the spreadsheet function-compatibility
harness.

WHAT THIS DOES (AND WHAT IT DELIBERATELY DOES NOT DO)
-----------------------------------------------------
There is no headless Google Sheets binary to drive, and this script has no
Google credentials of its own. What it does have is the one property that
makes Google Sheets executable at all from here:

  Uploading a formula-only .xlsx to Google Drive with the Google-Sheets
  target MIME type auto-converts it into a real Google Sheet, and that
  conversion RECALCULATES every formula with Google's own calculation
  engine. Exporting that Sheet back out as .xlsx produces a workbook whose
  cells carry Google Sheets' computed cached values.

So this runner is split in half around an external upload/download step
performed by whoever/whatever has Drive access (the orchestrator):

    build   -- emit N small .xlsx chunk workbooks + a manifest.json
               (upload these to Drive, let Drive convert them, export each
               back as .xlsx)
    ingest  -- read the exported .xlsx files back, map every cell to its
               test case via the manifest, and write
               results/google-sheets.json in exactly the schema
               results/libreoffice-*.json uses.

The full orchestrator loop is documented in README.md, "Phase 2: Google
Sheets runner".

WHY THE FILES ARE CHUNKED
-------------------------
The .xlsx bytes travel through the orchestrator's context window as base64
(~4/3 size inflation), so one 841-sheet ~430 KB workbook is not
transportable. `build` therefore splits the corpus by FUNCTION (never
splitting a function's cases across two files) into chunks of ~40 functions,
which measure ~60 KB each -- comfortably under the ~150 KB per-workbook
budget.

WHY THE RESULTS ARE STILL TRUSTWORTHY (THE CANARY, SAME AS THE LO RUNNER)
-------------------------------------------------------------------------
openpyxl NEVER writes a cached <v> value for a formula cell -- it writes
only the formula string. The uploaded chunk therefore contains zero cached
values. That gives us the same proof the LibreOffice runner relies on:

  (a) Deterministic canary: `=1111+2222` sits in Z1 of EVERY sheet. If
      Google Sheets had not recalculated, the exported cell would come back
      blank (None), because there was never a cached value to preserve.
      Reading back exactly 3333 proves the cell was computed by Google.
  (b) Volatile canary: `=NOW()` on the `_meta` sheet. Sheets evaluates it
      at import time, so a plausible fresh timestamp corroborates (a).
      NOTE the honest limitation: unlike the LO runner, which converts the
      same file twice a few seconds apart and shows =NOW() DIFFERING
      between runs, a single Drive import gives us one timestamp only. The
      cross-run volatile proof is therefore weaker here; the deterministic
      canary is the load-bearing one, and it is checked on every single
      sheet.

If any deterministic canary fails, the whole run is marked "trusted": false
and each affected case gets an "UNTRUSTED_RECALC" note.

WHAT WE CANNOT HONESTLY CLAIM
-----------------------------
Google Sheets has no user-visible version number and exposes none through
Drive. So `engine_version` here is NOT a version -- it is a dated label
("Google Sheets (Drive import, YYYY-MM-DD)") recording WHEN the corpus was
executed. Sheets is a continuously-updated hosted product; a result from
this file means "this is what Google Sheets did on that date", and nothing
stronger. Anything presenting Sheets results must say so.

THE _xlfn. PREFIX (VERIFIED, DO NOT "FIX")
------------------------------------------
Google Sheets DOES understand the OOXML `_xlfn.` storage prefix on import:
verified empirically -- `_xlfn.XLOOKUP(...)` imported and computed a real
value (1), it did not come back as #NAME?. So this runner uses the exact
same harness/xlfn_map.py translation as the LibreOffice path, with no
Sheets-specific special-casing. Writing the bare unprefixed name instead
would be the thing that breaks.

USAGE
-----
    # 1. build the chunk workbooks (all 278 functions, ~40 per chunk)
    python3 harness/run_sheets.py build
    python3 harness/run_sheets.py build --chunk-size 25
    python3 harness/run_sheets.py build --only XLOOKUP LET FILTER
    python3 harness/run_sheets.py build --outdir harness/sheets_chunks

    # 2. (external) upload each chunk to Drive as a Google Sheet, export it
    #    back as .xlsx into harness/sheets_exports/chunk-NN-export.xlsx

    # 3. ingest the exports (incremental: run it per chunk as they arrive)
    python3 harness/run_sheets.py ingest \
        --export harness/sheets_exports/chunk-01-export.xlsx \
        --engine-label "Google Sheets (Drive import, 2026-08-29)"

    # dry run of the whole pipeline with LibreOffice standing in for Drive,
    # writing to a scratch results file (never results/google-sheets.json):
    python3 harness/run_sheets.py selftest --only COUNT SUM MROUND

    # ---- the how-to RECIPE corpus (data/recipes/*.json) ----
    # Same three-step loop, one worksheet per recipe check, writing
    # results/recipes-verified-sheets.json alongside the LibreOffice-executed
    # results/recipes-verified.json.

    # 1. build (plain function names by default -- Sheets maps unprefixed
    #    names, and recipes are authored the way a user types them)
    python3 harness/run_sheets.py build-recipes
    python3 harness/run_sheets.py build-recipes --chunk-size 60
    python3 harness/run_sheets.py build-recipes --only how-to-use-xlookup add-days-to-a-date
    python3 harness/run_sheets.py build-recipes --outdir harness/recipe_chunks
    python3 harness/run_sheets.py build-recipes --xlfn-names   # opt out of plain names

    # 2. (external) upload each chunk to Drive, export back as
    #    harness/recipe_exports/chunk-NN-export.xlsx

    # 3. ingest (incremental, per chunk as exports land)
    python3 harness/run_sheets.py ingest-recipes \
        --export harness/recipe_exports/chunk-01-export.xlsx \
        --out results/recipes-verified-sheets.json

    # dry run of the recipe pipeline with LibreOffice standing in for Drive;
    # every value must match results/recipes-verified.json or it exits 1
    python3 harness/run_sheets.py selftest-recipes

INCREMENTAL INGESTION MERGES, IT DOES NOT OVERWRITE
---------------------------------------------------
Every ingest is inherently a subset (one or more chunks out of N), so if
--out already exists the new results are MERGED in with exactly the
semantics run_lo.py uses for subset runs: only the functions actually
ingested this time are replaced, every other function's previously executed
result is preserved byte-for-byte, generated_at / canary / recalc_method are
refreshed to describe the most recent execution, and `trusted` becomes the
AND of the previous file's flag and this run's -- a merge can only ever
downgrade trust, never launder an untrusted ingest into a trusted file. Each
merge appends to a top-level "subset_runs" list for auditability. Merging
into a file recorded under a DIFFERENT engine label is refused unless
--allow-label-change is passed (results executed on two different dates
against a silently-updated hosted product should not be blended without a
deliberate decision; when it is allowed, both labels are recorded).
"""
import argparse
import glob
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime, timezone

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import openpyxl  # noqa: E402

from openpyxl.worksheet.formula import ArrayFormula  # noqa: E402

from corpus import (  # noqa: E402
    RESULTS_DIR,
    CANARY_ANCHOR,
    CANARY_ARITH_EXPECTED,
    CANARY_ARITH_FORMULA,
    SHEET_NAME_MAX,
    SHEETS_EXTRA_ERROR_STRINGS,
    assert_sheets_safe_name,
    build_workbook,
    cell_addrs_in_range,
    compare_expected,
    flatten_cases,
    is_error_value,
    load_test_files,
    normalize_readback_value,
    sanitize_sheet_name,
)
# The RECIPE corpus (data/recipes/*.json) shares its enumeration, setup
# inheritance, normalization and comparison rules with
# scripts/verify_recipes.py -- see harness/recipe_corpus.py.
from recipe_corpus import (  # noqa: E402
    compare_check,
    iter_checks,
    load_recipe_files,
    norm,
    setup_sheet_names,
    uses_setup_sheets,
)
from xlfn_map import to_storage_formula_all  # noqa: E402

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
DEFAULT_CHUNK_DIR = os.path.join(REPO_ROOT, "harness", "sheets_chunks")
DEFAULT_EXPORT_DIR = os.path.join(REPO_ROOT, "harness", "sheets_exports")
DEFAULT_RESULTS_PATH = os.path.join(RESULTS_DIR, "google-sheets.json")
# Scratch output for `selftest`. Deliberately OUTSIDE results/: it holds
# LibreOffice values, and results/ is only ever real engine output.
SELFTEST_RESULTS_PATH = os.path.join(
    REPO_ROOT, "harness", "sheets_selftest", "plumbing-check.json")
MANIFEST_NAME = "manifest.json"

ENGINE_ID = "google_sheets"
RECALC_METHOD = "Drive import + xlsx export readback"

# `selftest` produces LibreOffice values through the Sheets plumbing. The site
# generator discovers engines by globbing results/*.json and matching on the
# "engine" string ("google" anywhere in it means Google Sheets), so selftest
# output MUST NOT carry ENGINE_ID and MUST NOT land in results/ -- either
# alone would be enough for the site to publish LibreOffice numbers in a
# Google Sheets column. Both are enforced in cmd_selftest().
SELFTEST_ENGINE_ID = "SELFTEST_libreoffice_via_sheets_pipeline"
SELFTEST_RECALC_METHOD = ("selftest plumbing check: soffice --convert-to xlsx "
                          "standing in for the Drive import -- NOT Google Sheets")

# Advisory ceiling per chunk workbook. The bytes travel as base64 through the
# orchestrator's context (~4/3 inflation), so this is about transportability,
# not about any Drive or Sheets limit.
SIZE_WARN_BYTES = 150 * 1024

META_SHEET = "_meta"
META_VOLATILE_CELL = "A1"
META_ARITH_CELL = "A2"


# --------------------------------------------------------------------------
# build
# --------------------------------------------------------------------------

def chunk_functions(cases_flat, chunk_size):
    """Group flat cases by function, then partition the FUNCTIONS (never a
    single function's cases) into chunks of at most `chunk_size` functions.

    Keeping every case of a function in one workbook matters: a function's
    results are merged into the results file wholesale, so a function split
    across two chunks could end up half-ingested and silently truncated."""
    by_fn = {}
    for c in cases_flat:
        by_fn.setdefault(c["function"], []).append(c)
    fns = sorted(by_fn)
    chunks = []
    for i in range(0, len(fns), chunk_size):
        group = fns[i:i + chunk_size]
        chunks.append((group, [c for fn in group for c in by_fn[fn]]))
    return chunks


def _sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for block in iter(lambda: f.read(65536), b""):
            h.update(block)
    return h.hexdigest()


def cmd_build(args):
    requested = set(args.only) if args.only else None
    test_files = load_test_files(requested)
    if not test_files:
        sys.exit("No test files matched.")
    if requested:
        missing = requested - {fn for fn, _, _ in test_files}
        if missing:
            sys.exit(f"No data/tests/*.json for: {', '.join(sorted(missing))}")

    plain_names = bool(getattr(args, "plain_names", False))
    manifest_note = getattr(args, "manifest_note", None)

    cases_flat = flatten_cases(test_files)
    chunks = chunk_functions(cases_flat, args.chunk_size)

    outdir = os.path.abspath(args.outdir)
    os.makedirs(outdir, exist_ok=True)
    # Stale chunk files from a previous, differently-sized build would be
    # silently ingestible against the new manifest, so clear them out.
    for stale in glob.glob(os.path.join(outdir, "chunk-*.xlsx")):
        os.remove(stale)

    width = max(2, len(str(len(chunks))))
    manifest_chunks = []
    for idx, (fn_group, chunk_cases) in enumerate(chunks, start=1):
        chunk_id = f"chunk-{idx:0{width}d}"
        wb, sheet_map = build_workbook(chunk_cases, plain_names=plain_names)

        # Every sheet name, including the meta sheet, must survive the Drive
        # import unmangled or the manifest's cell mapping silently breaks.
        for name in list(sheet_map.values()) + [META_SHEET]:
            assert_sheets_safe_name(name)

        path = os.path.join(outdir, f"{chunk_id}.xlsx")
        wb.save(path)

        case_entries = []
        for c in chunk_cases:
            # Must mirror build_workbook()'s own per-case choice exactly --
            # this is what actually got written into the sheet, and ingest
            # trusts it verbatim for both display and provenance.
            stored_formula = c["formula"] if plain_names else to_storage_formula_all(c["formula"])
            case_entries.append({
                "test_id": c["test_id"],
                "function": c["function"],
                "sheet": sheet_map[c["test_id"]],
                "anchor": c["anchor"],
                "check_range": c["check_range"],
                "description": c["description"],
                "formula_display": c["formula"],
                "formula_stored_xlsx": stored_formula,
                "serialization": "plain" if plain_names else "xlfn",
                "expected": c["expected"],
                "expected_note": c["expected_note"],
            })

        manifest_chunks.append({
            "chunk": chunk_id,
            "file": os.path.basename(path),
            "bytes": os.path.getsize(path),
            "sha256": _sha256(path),
            "n_functions": len(fn_group),
            "n_cases": len(chunk_cases),
            "functions": fn_group,
            "sheets": sorted(set(sheet_map.values())),
            "cases": case_entries,
        })

    manifest = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "chunk_size": args.chunk_size,
        "n_chunks": len(chunks),
        "n_functions": sum(c["n_functions"] for c in manifest_chunks),
        "n_cases": sum(c["n_cases"] for c in manifest_chunks),
        "subset_only": sorted(requested) if requested else None,
        "plain_names": plain_names,
        "plain_names_functions": sorted({fn for c in manifest_chunks for fn in c["functions"]}) if plain_names else None,
        "manifest_note": manifest_note,
        "canary": {
            "arithmetic_formula": CANARY_ARITH_FORMULA,
            "arithmetic_expected": CANARY_ARITH_EXPECTED,
            "arithmetic_cell_every_sheet": CANARY_ANCHOR,
            "meta_sheet": META_SHEET,
            "meta_volatile_cell": META_VOLATILE_CELL,
            "meta_volatile_formula": "=NOW()",
            "meta_arithmetic_cell": META_ARITH_CELL,
        },
        "chunks": manifest_chunks,
    }
    manifest_path = os.path.join(outdir, MANIFEST_NAME)
    with open(manifest_path, "w") as f:
        json.dump(manifest, f, indent=2, default=str)
        f.write("\n")

    mode = "PLAIN NAMES (no _xlfn./_xlfn._xlws. translation)" if plain_names else "xlfn-translated (default)"
    print(f"Built {len(chunks)} chunk(s) from {manifest['n_functions']} function(s), "
          f"{manifest['n_cases']} case(s) -> {outdir}  [{mode}]")
    if manifest_note:
        print(f"Manifest note: {manifest_note}")
    total = 0
    for c in manifest_chunks:
        total += c["bytes"]
        flag = "  <-- OVER SIZE BUDGET" if c["bytes"] > SIZE_WARN_BYTES else ""
        print(f"  {c['file']:>16}  {c['bytes']:>8,} bytes  "
              f"{c['n_functions']:>3} fn  {c['n_cases']:>4} cases{flag}")
    print(f"  {'TOTAL':>16}  {total:>8,} bytes "
          f"(base64 ~{int(total * 4 / 3):,} bytes)")
    print(f"Wrote {manifest_path}")
    return manifest


# --------------------------------------------------------------------------
# ingest
# --------------------------------------------------------------------------

def _chunk_id_from_path(path, manifest):
    """Figure out which manifest chunk an exported workbook corresponds to.

    Filename first (chunk-NN-export.xlsx), then verified against the chunk's
    sheet list so an export uploaded/downloaded out of order can never be
    silently ingested against the wrong cell map."""
    base = os.path.basename(path)
    ids = {c["chunk"] for c in manifest["chunks"]}
    m = re.search(r"(chunk-\d+)", base)
    if m and m.group(1) in ids:
        return m.group(1)
    raise SystemExit(
        f"Cannot tell which chunk {base!r} is. Name exports "
        f"'<chunk-id>-export.xlsx' using an id from the manifest "
        f"({', '.join(sorted(ids))})."
    )


def _boolean_text_artifact(expected, actual):
    """Sheets' .xlsx export can serialize a boolean result as the TEXT
    'TRUE'/'FALSE' rather than a boolean cell. That is an export-format
    artifact, not an engine disagreement, so we never silently coerce it --
    the raw value is recorded as-is and this note flags it for a human."""
    return (isinstance(expected, bool)
            and isinstance(actual, str)
            and actual.upper() in ("TRUE", "FALSE"))


def ingest_exports(export_paths, manifest, engine_label):
    """Read exported workbooks -> (function_results, canary, trusted, stats).

    Provenance: `manifest["plain_names"]` (set by `build --plain-names`)
    records whether these chunk workbooks had formulas written EXACTLY as
    authored in data/tests (no _xlfn./_xlfn._xlws. storage-form
    translation) rather than the default xlfn-translated serialization.
    Every ingested case gets a "serialization": "plain"/"xlfn" field so a
    later merge (or a human reading results/google-sheets.json) can tell
    which wire form produced a given result -- this matters for Google
    Sheets specifically because its xlsx importer maps plain `_xlfn.NAME`
    but NOT `_xlfn._xlws.FILTER/SORT` or LAMBDA-family functions, so a
    plain-names result for those functions supersedes an earlier
    xlfn-serialized one rather than merely re-confirming it.
    """
    plain_names = bool(manifest.get("plain_names", False))
    serialization = "plain" if plain_names else "xlfn"
    manifest_note = manifest.get("manifest_note")
    function_results = {}
    per_chunk = []
    sheet_canary_failures = []
    n_sheet_canaries = 0
    volatile_values = {}
    meta_arith = {}

    cases_by_chunk = {c["chunk"]: c for c in manifest["chunks"]}

    for path in export_paths:
        chunk_id = _chunk_id_from_path(path, manifest)
        chunk = cases_by_chunk[chunk_id]
        wb = openpyxl.load_workbook(path, data_only=True)

        expected_sheets = set(chunk["sheets"])
        present = set(wb.sheetnames)
        missing = expected_sheets - present
        if missing:
            raise SystemExit(
                f"{os.path.basename(path)} claims to be {chunk_id} but is missing "
                f"{len(missing)} of its sheets (e.g. {sorted(missing)[:3]}). "
                f"Wrong file for this chunk, or the manifest is stale -- rebuild."
            )

        if META_SHEET in present:
            meta = wb[META_SHEET]
            volatile_values[chunk_id] = meta[META_VOLATILE_CELL].value
            meta_arith[chunk_id] = meta[META_ARITH_CELL].value
        else:
            volatile_values[chunk_id] = None
            meta_arith[chunk_id] = None

        for case in chunk["cases"]:
            ws = wb[case["sheet"]]
            n_sheet_canaries += 1
            canary_val = ws[CANARY_ANCHOR].value
            canary_ok = canary_val == CANARY_ARITH_EXPECTED
            if not canary_ok:
                sheet_canary_failures.append(
                    {"chunk": chunk_id, "sheet": case["sheet"], "value": canary_val})

            raw_anchor = ws[case["anchor"]].value
            anchor_val = normalize_readback_value(raw_anchor)

            range_flat = None
            if case["check_range"]:
                grid = cell_addrs_in_range(case["check_range"])
                range_flat = [normalize_readback_value(ws[addr].value)
                              for row in grid for addr in row]

            # Google's export writes errors as cached error strings, exactly
            # like LibreOffice's does; the raw string is kept verbatim.
            error = anchor_val if is_error_value(
                anchor_val, SHEETS_EXTRA_ERROR_STRINGS) else None
            matched, mismatch_detail = compare_expected(
                case["expected"], anchor_val, range_flat)

            notes = []
            if not canary_ok:
                notes.append("UNTRUSTED_RECALC: per-sheet canary failed on this sheet")
            if case.get("expected_note"):
                notes.append(case["expected_note"])
            if mismatch_detail:
                notes.append(f"MISMATCH vs expected: {mismatch_detail}")
            if _boolean_text_artifact(case["expected"], anchor_val):
                notes.append("NOTE: read back as the TEXT 'TRUE'/'FALSE' rather "
                             "than a boolean cell -- likely an .xlsx export "
                             "serialization artifact, verify before reporting it "
                             "as an engine difference")
            if error == "#ERROR!":
                notes.append("NOTE: #ERROR! is Google Sheets' parse-failure error "
                             "(no Excel equivalent); it means Sheets could not "
                             "parse the formula, which is not the same as #NAME?")

            function_results.setdefault(case["function"], {})[case["test_id"]] = {
                "description": case["description"],
                "formula_display": case["formula_display"],
                "formula_stored_xlsx": case["formula_stored_xlsx"],
                "serialization": case.get("serialization", serialization),
                "value": anchor_val,
                "range_values": range_flat,
                "error": error,
                "expected": case["expected"],
                "matched_expected": matched,
                "canary_ok_this_sheet": canary_ok,
                "notes": "; ".join(notes) if notes else None,
            }

        per_chunk.append({
            "chunk": chunk_id,
            "export_file": os.path.basename(path),
            "n_functions": chunk["n_functions"],
            "n_cases": chunk["n_cases"],
        })

    arith_ok = (not sheet_canary_failures
                and n_sheet_canaries > 0
                and all(v == CANARY_ARITH_EXPECTED for v in meta_arith.values()))

    canary = {
        "arithmetic_formula": CANARY_ARITH_FORMULA,
        "arithmetic_expected": CANARY_ARITH_EXPECTED,
        "arithmetic_actual": (CANARY_ARITH_EXPECTED if arith_ok
                              else (sheet_canary_failures[0]["value"]
                                    if sheet_canary_failures else None)),
        "arithmetic_ok": arith_ok,
        "arithmetic_sheets_checked": n_sheet_canaries,
        "arithmetic_sheet_failures": sheet_canary_failures[:20],
        "meta_arithmetic_per_chunk": meta_arith,
        "volatile_formula": "=NOW()",
        "volatile_per_chunk": {k: str(v) for k, v in volatile_values.items()},
        "now_differs_across_runs": None,
        "method": (
            "openpyxl writes formulas with NO cached <v> value, so the uploaded "
            "chunk workbooks contain zero cached results. Google Drive's "
            "auto-conversion to a Google Sheet recalculates every formula with "
            "Google's engine; exporting that Sheet back to .xlsx carries those "
            "computed values out. The deterministic canary =1111+2222 in "
            f"{CANARY_ANCHOR} of EVERY sheet reading back exactly "
            f"{CANARY_ARITH_EXPECTED} therefore proves genuine computation -- "
            "without recalculation the cell would read back blank (None), since "
            "nothing was ever cached. The volatile =NOW() canary is recorded per "
            "chunk as corroboration, but UNLIKE the LibreOffice runner (which "
            "converts the same file twice and shows =NOW() differing) a single "
            "Drive import yields one timestamp, so no cross-run volatile "
            "comparison is possible: now_differs_across_runs is null by design, "
            "and the deterministic canary is the load-bearing proof here."
        ),
        "engine_label": engine_label,
    }

    stats = {"chunks": per_chunk,
             "n_functions": len(function_results),
             "n_cases": sum(len(v) for v in function_results.values()),
             "plain_names": plain_names,
             "serialization": serialization,
             "manifest_note": manifest_note}
    return function_results, canary, arith_ok, stats


def write_results(out_path, function_results, canary, trusted, engine_label,
                  allow_label_change=False, engine=ENGINE_ID,
                  recalc_method=RECALC_METHOD, serialization=None,
                  manifest_note=None):
    """Write (or incrementally merge into) a results file in the same schema
    results/libreoffice-*.json uses. See the module docstring."""
    generated_at = datetime.now(timezone.utc).isoformat()
    output = {
        "generated_at": generated_at,
        "engine": engine,
        "engine_version": engine_label,
        "recalc_method": recalc_method,
        "trusted": trusted,
        "canary": canary,
        "function_results": function_results,
    }

    if os.path.exists(out_path):
        with open(out_path) as f:
            prev = json.load(f)
        prev_label = prev.get("engine_version")
        if prev_label != engine_label and not allow_label_change:
            sys.exit(
                f"Refusing to merge: {out_path} records engine_version "
                f"{prev_label!r} but this ingest is labelled {engine_label!r}. "
                f"Google Sheets is a continuously-updated hosted product, so "
                f"results executed on different dates should not be blended "
                f"without a deliberate decision. Re-run with the same "
                f"--engine-label, or pass --allow-label-change to record both."
            )
        merged = dict(prev)
        merged_fr = dict(prev.get("function_results") or {})
        for fn, cases in function_results.items():
            merged_fr[fn] = cases  # whole function re-executed, so replace wholesale
        merged["function_results"] = merged_fr
        merged["generated_at"] = generated_at
        merged["engine"] = engine
        merged["engine_version"] = engine_label
        merged["recalc_method"] = recalc_method
        merged["canary"] = canary
        # A merge can only downgrade trust, never upgrade it.
        merged["trusted"] = bool(prev.get("trusted", False)) and trusted
        if prev_label != engine_label:
            merged.setdefault("engine_version_history", [])
            if prev_label not in merged["engine_version_history"]:
                merged["engine_version_history"].append(prev_label)
            merged["engine_version_history"].append(engine_label)
        merged.setdefault("subset_runs", []).append({
            "generated_at": generated_at,
            "functions": sorted(function_results.keys()),
            "trusted_this_run": trusted,
            "engine_label": engine_label,
            "superseded_generated_at": prev.get("generated_at"),
            # Provenance: which wire serialization produced this subset's
            # results. A "plain" run for a function supersedes an earlier
            # "xlfn" run for that SAME function (this merge already replaced
            # it above); it is not merely a re-confirmation, because Google
            # Sheets' xlsx importer treats the two serializations
            # differently for FILTER/SORT/LAMBDA-family functions. See
            # README.md "Phase 2: Google Sheets runner".
            "serialization": serialization,
            "manifest_note": manifest_note,
        })
        untouched = len(merged_fr) - len(function_results)
        print(f"Incremental ingest: merged {len(function_results)} function(s) into "
              f"{os.path.basename(out_path)}; {untouched} other function result(s) "
              f"preserved unchanged.")
        output = merged

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    with open(out_path, "w") as f:
        json.dump(output, f, indent=2, default=str)
        f.write("\n")
    return output


def summarize(function_results):
    n_name_error = n_other_error = n_ok = 0
    for cases in function_results.values():
        for r in cases.values():
            if r["error"] == "#NAME?":
                n_name_error += 1
            elif r["error"]:
                n_other_error += 1
            else:
                n_ok += 1
    return n_ok, n_name_error, n_other_error


def _load_manifest(args):
    path = args.manifest or os.path.join(os.path.abspath(args.chunkdir), MANIFEST_NAME)
    if not os.path.exists(path):
        sys.exit(f"Manifest not found: {path}. Run `run_sheets.py build` first.")
    with open(path) as f:
        return json.load(f)


def cmd_ingest(args):
    manifest = _load_manifest(args)
    for p in args.export:
        if not os.path.exists(p):
            sys.exit(f"Export not found: {p}")

    function_results, canary, trusted, stats = ingest_exports(
        args.export, manifest, args.engine_label)

    write_results(args.out, function_results, canary, trusted,
                  args.engine_label, args.allow_label_change,
                  serialization=stats["serialization"],
                  manifest_note=stats["manifest_note"])

    n_ok, n_name, n_other = summarize(function_results)
    print(f"Ingested {len(stats['chunks'])} chunk(s): "
          f"{stats['n_functions']} function(s), {stats['n_cases']} case(s). "
          f"[serialization={stats['serialization']}]")
    if stats["manifest_note"]:
        print(f"Manifest note: {stats['manifest_note']}")
    print(f"Deterministic canary OK on all {canary['arithmetic_sheets_checked']} "
          f"sheet(s): {canary['arithmetic_ok']}")
    if canary["arithmetic_sheet_failures"]:
        print(f"  CANARY FAILURES: {canary['arithmetic_sheet_failures']}")
    print(f"Volatile =NOW() per chunk: {canary['volatile_per_chunk']}")
    print(f"Trusted recalculation: {trusted}")
    print(f"Cases with a value (no error): {n_ok}")
    print(f"Cases returning #NAME? (unsupported function): {n_name}")
    print(f"Cases returning some other error: {n_other}")
    print(f"Wrote {args.out}")
    return function_results


# --------------------------------------------------------------------------
# selftest: prove the build -> export -> ingest plumbing end-to-end
# --------------------------------------------------------------------------

def cmd_selftest(args):
    """Run the whole pipeline with LibreOffice standing in for Google Drive.

    This proves the CELL MAPPING and the ingest plumbing, nothing about
    Google Sheets: the values recovered are LibreOffice's. It therefore
    writes to a scratch results file and never to results/google-sheets.json.
    """
    print("=== selftest: build ===")
    manifest = cmd_build(args)

    chunkdir = os.path.abspath(args.outdir)
    soffice = os.environ.get("SOFFICE_BIN", "soffice")

    with tempfile.TemporaryDirectory() as tmp:
        export_dir = os.path.join(tmp, "exports")
        os.makedirs(export_dir)
        exports = []
        print("\n=== selftest: simulated export (soffice --convert-to xlsx) ===")
        for chunk in manifest["chunks"]:
            src = os.path.join(chunkdir, chunk["file"])
            outd = os.path.join(tmp, chunk["chunk"])
            os.makedirs(outd, exist_ok=True)
            proc = subprocess.run(
                [soffice, "--headless", "--convert-to", "xlsx",
                 "--outdir", outd, src],
                capture_output=True, text=True, timeout=600)
            produced = os.path.join(outd, chunk["file"])
            if proc.returncode != 0 or not os.path.exists(produced):
                sys.exit(f"soffice conversion failed for {chunk['file']}: "
                         f"{proc.stderr.strip()}")
            dest = os.path.join(export_dir, f"{chunk['chunk']}-export.xlsx")
            shutil.copyfile(produced, dest)
            exports.append(dest)
            print(f"  {chunk['file']} -> {os.path.basename(dest)} "
                  f"({os.path.getsize(dest):,} bytes)")

        print("\n=== selftest: ingest ===")
        function_results, canary, trusted, stats = ingest_exports(
            exports, manifest, args.engine_label)

        out = os.path.abspath(args.out)
        # Hard stop: the site generator globs results/*.json and assigns an
        # engine from the "engine" string, so ANY selftest output inside
        # results/ risks LibreOffice values being published as Google Sheets.
        if os.path.commonpath([out, os.path.abspath(RESULTS_DIR)]) == \
                os.path.abspath(RESULTS_DIR):
            sys.exit(f"selftest refuses to write inside results/ ({out}). These "
                     f"are LibreOffice values produced through the Sheets "
                     f"plumbing, not Google Sheets values, and the site "
                     f"generator picks up every results/*.json by engine name.")
        if os.path.exists(out):
            os.remove(out)  # scratch file: always a fresh write, never a merge
        write_results(out, function_results, canary, trusted, args.engine_label,
                      engine=SELFTEST_ENGINE_ID,
                      recalc_method=SELFTEST_RECALC_METHOD,
                      serialization=stats["serialization"],
                      manifest_note=stats["manifest_note"])

    n_ok, n_name, n_other = summarize(function_results)
    print(f"\nChunks ingested       : {len(stats['chunks'])}")
    print(f"Functions / cases     : {stats['n_functions']} / {stats['n_cases']}")
    print(f"Per-sheet canaries    : {canary['arithmetic_sheets_checked']} checked, "
          f"ok={canary['arithmetic_ok']}, failures={canary['arithmetic_sheet_failures']}")
    print(f"Volatile =NOW()       : {canary['volatile_per_chunk']}")
    print(f"trusted               : {trusted}")
    print(f"value / #NAME? / other: {n_ok} / {n_name} / {n_other}")

    print("\n=== 3 sample round-tripped values ===")
    shown = 0
    for fn in sorted(function_results):
        for tid in sorted(function_results[fn]):
            r = function_results[fn][tid]
            print(f"  {tid}: {r['formula_display']}  ->  value={r['value']!r}  "
                  f"expected={r['expected']!r}  matched={r['matched_expected']}")
            shown += 1
            if shown >= 3:
                break
        if shown >= 3:
            break
    print(f"\nWrote scratch results -> {args.out} "
          f"(LibreOffice values; plumbing proof only)")


# ==========================================================================
# RECIPES: the second corpus (data/recipes/*.json), executed in Google Sheets
# ==========================================================================
# The how-to recipes have been executed in headless LibreOffice since day one
# (scripts/verify_recipes.py -> results/recipes-verified.json). These three
# subcommands give them the same Drive-import treatment the FUNCTION corpus
# gets above, writing results/recipes-verified-sheets.json.
#
# Everything about WHICH checks exist, how a variant check inherits its
# setup, how a read-back value is normalized and how it is judged against
# `expected` comes from harness/recipe_corpus.py -- the same module
# verify_recipes.py now imports. That is deliberate and load-bearing: the
# site prints "LibreOffice returned X, Google Sheets returned Y" side by
# side, so any difference in enumeration or comparison between the two paths
# would manufacture a fake engine divergence.
#
# SERIALIZATION: unlike the function corpus, recipe chunks are written with
# PLAIN function names by default (`--plain-names`, on unless you pass
# `--xlfn-names`). Google Sheets' xlsx importer maps bare modern names
# (FILTER, SORT, LAMBDA, TEXTJOIN, ...) but does NOT map the
# `_xlfn._xlws.FILTER` storage form, and recipes are authored the way a user
# would type them into Sheets. The trade-off is recorded per chunk in the
# manifest and per run in the results file, because it is the one thing that
# makes a Sheets recipe result not directly byte-comparable to the
# LibreOffice reference run for those functions.

DEFAULT_RECIPE_CHUNK_DIR = os.path.join(REPO_ROOT, "harness", "recipe_chunks")
DEFAULT_RECIPE_EXPORT_DIR = os.path.join(REPO_ROOT, "harness", "recipe_exports")
DEFAULT_RECIPE_RESULTS_PATH = os.path.join(RESULTS_DIR, "recipes-verified-sheets.json")
# Scratch output for `selftest-recipes`. Deliberately OUTSIDE results/, for
# exactly the reason cmd_selftest's is: these are LibreOffice values.
RECIPE_SELFTEST_DIR = os.path.join(REPO_ROOT, "harness", "recipe_selftest")
RECIPE_SELFTEST_RESULTS_PATH = os.path.join(RECIPE_SELFTEST_DIR, "plumbing-check.json")
RECIPE_SELFTEST_CHUNK_DIR = os.path.join(RECIPE_SELFTEST_DIR, "chunks")
LO_RECIPE_RESULTS_PATH = os.path.join(RESULTS_DIR, "recipes-verified.json")

RECIPE_RECALC_METHOD = "Drive import + xlsx export readback (how-to recipe corpus)"
RECIPE_SELFTEST_ENGINE_ID = "SELFTEST_libreoffice_via_sheets_pipeline"

# Six slugs that exercise a scalar result, a date-as-text result, an array
# result written as a real array formula over a check_range, a 1x1
# check_range, and the only single-sheet recipe that carries variants.
RECIPE_SELFTEST_SLUGS = [
    "abbreviate-large-numbers",       # plain scalar text result
    "add-days-to-a-date",             # TEXT()-wrapped date
    "case-sensitive-lookup",          # 1x1 check_range (array formula)
    "list-the-top-n-values",          # 3-cell spill via check_range
    "transpose-rows-to-columns",      # 3-cell horizontal spill
    "sum-with-multiple-criteria",     # 15 variant checks across 5 sections
]


def recipe_sheet_name(recipe_index, check_key, slug, used):
    """Worksheet name for one check: '<id>_<slug>', truncated to 31 chars.

    The id prefix (r007, r007v2c1) is what makes the name unique and is
    derived from the recipe's position in the FULL corpus, so a `--only`
    build lands its checks on exactly the sheets a full build would. The
    slug tail is purely so a human opening the workbook in Drive can tell
    what they are looking at; ingest reads the name from the manifest and
    never re-derives it.

    31 chars, not Sheets' 100: the transport is .xlsx in both directions and
    a >31-char sheet name is out-of-spec OOXML whatever opens the file (the
    same rule corpus.py applies to the function corpus).
    """
    prefix = f"r{recipe_index:03d}" if check_key == "main" else f"r{recipe_index:03d}{check_key}"
    room = SHEET_NAME_MAX - len(prefix) - 1
    name = f"{prefix}_{slug[:room]}" if room > 0 else prefix
    return assert_sheets_safe_name(sanitize_sheet_name(name, used))


def collect_recipe_checks(only=None):
    """Return (buildable, skipped, missing).

    buildable: [{slug, title, index, checks:[check dicts from iter_checks]}]
    skipped:   multi-sheet recipes, which v1 does not build (see below)
    missing:   requested slugs with no data/recipes/*.json

    WHY MULTI-SHEET RECIPES ARE SKIPPED (v1 decision, deliberate)
    -------------------------------------------------------------
    Six recipes need extra worksheets to exist (`setup_sheets`), and the
    natural fix -- give every check its own copy of those tabs, prefixed to
    avoid collisions, and rewrite the formula's sheet references to match --
    is NOT correct for this corpus. Their checks are precisely the ones a
    rename breaks:

      * `=INDIRECT(A1&"!B2")` builds the sheet reference out of a CELL VALUE,
        so renaming the tab silently changes the answer unless the setup data
        is rewritten too, and one of those checks exists specifically to
        assert a #REF! for a tab that does NOT exist.
      * `=SUM(Q1:Q3!A1)` is a 3-D reference over a SPAN of consecutive tabs;
        it depends on sheet ORDER as well as names.
      * `=$'Q1 Data'.B2` and `=Data.B2` are there to assert #NAME? -- the
        LibreOffice-syntax forms that Excel/Sheets reject. A rewriter that
        "fixed" those references would destroy the very thing under test.
      * Sheet names collide WITHIN a single recipe: one variant defines
        Q1/Q2/Q3 as numbers and a check inside it redefines Q1/Q2/Q3 with
        different contents, so even one-recipe-per-workbook does not
        de-collide them.

    Skipping is therefore the simpler CORRECT option: they stay
    LibreOffice-executed, the site keeps their existing LibreOffice-only
    wording, and the manifest lists them by name with their extra tabs so
    the omission is visible rather than silent. A v2 that runs one workbook
    per CHECK (not per recipe) would cover them honestly with no rewriting
    at all, at the cost of ~34 more Drive round-trips.
    """
    all_recipes = load_recipe_files()
    index_of = {slug: i + 1 for i, (slug, _, _) in enumerate(all_recipes)}
    wanted = set(only) if only else None
    missing = sorted(wanted - set(index_of)) if wanted else []

    buildable, skipped = [], []
    for slug, _path, recipe in all_recipes:
        if wanted and slug not in wanted:
            continue
        checks = list(iter_checks(recipe))
        entry = {
            "slug": slug,
            "title": recipe.get("title", ""),
            "index": index_of[slug],
            "checks": checks,
            "n_checks": len(checks),
        }
        if uses_setup_sheets(recipe):
            entry["setup_sheet_names"] = setup_sheet_names(recipe)
            skipped.append(entry)
        else:
            buildable.append(entry)
    return buildable, skipped, missing


def chunk_recipes(recipes, chunk_size):
    """Partition RECIPES (never one recipe's checks) into chunks.

    Same rule, same reason, as chunk_functions(): a recipe is merged into the
    results file wholesale and its `verified` flag is the AND over all of its
    checks, so a recipe split across two chunks could be written out as
    "verified" from a half-ingested export."""
    return [recipes[i:i + chunk_size] for i in range(0, len(recipes), chunk_size)]


def build_recipe_workbook(recipes, plain_names=True):
    """One worksheet per check. Returns (workbook, sheet_map) where
    sheet_map[(slug, check_key)] -> sheet name."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    used_names = set()
    sheet_map = {}

    for rec in recipes:
        for chk in rec["checks"]:
            name = recipe_sheet_name(rec["index"], chk["key"], rec["slug"], used_names)
            sheet_map[(rec["slug"], chk["key"])] = name
            ws = wb.create_sheet(name)

            for addr, val in (chk["setup_cells"] or {}).items():
                ws[addr] = val

            formula = chk["formula"] if plain_names else to_storage_formula_all(chk["formula"])
            if chk["check_range"]:
                # Same reasoning as corpus.build_workbook(): a result that is
                # meant to spill has to be written as a real array formula
                # over its full range, or an engine returns a scalar #VALUE!
                # for a function it fully supports.
                ws[chk["anchor"]] = ArrayFormula(chk["check_range"], formula)
            else:
                ws[chk["anchor"]] = formula

            ws[CANARY_ANCHOR] = CANARY_ARITH_FORMULA

    meta = wb.create_sheet(META_SHEET, 0)
    meta[META_VOLATILE_CELL] = "=NOW()"
    meta[META_ARITH_CELL] = CANARY_ARITH_FORMULA
    assert_sheets_safe_name(META_SHEET)
    return wb, sheet_map


def cmd_build_recipes(args):
    plain_names = bool(getattr(args, "plain_names", True))
    buildable, skipped, missing = collect_recipe_checks(args.only)
    if missing:
        sys.exit(f"No data/recipes/*.json for: {', '.join(missing)}")
    if not buildable:
        sys.exit("No buildable recipes matched (all matches were multi-sheet "
                 "recipes, which this command skips -- see --help).")

    outdir = os.path.abspath(args.outdir)
    os.makedirs(outdir, exist_ok=True)
    # Stale chunks from a previous, differently-sized build would still be
    # ingestible against the new manifest, so clear them first.
    for stale in glob.glob(os.path.join(outdir, "chunk-*.xlsx")):
        os.remove(stale)

    chunks = chunk_recipes(buildable, args.chunk_size)
    width = max(2, len(str(len(chunks))))
    manifest_chunks = []
    for idx, group in enumerate(chunks, start=1):
        chunk_id = f"chunk-{idx:0{width}d}"
        wb, sheet_map = build_recipe_workbook(group, plain_names=plain_names)
        path = os.path.join(outdir, f"{chunk_id}.xlsx")
        wb.save(path)

        recipe_entries = []
        for rec in group:
            checks = []
            for chk in rec["checks"]:
                stored = (chk["formula"] if plain_names
                          else to_storage_formula_all(chk["formula"]))
                checks.append({
                    "key": chk["key"],
                    "kind": chk["kind"],
                    "variant_index": chk["variant_index"],
                    "check_index": chk["check_index"],
                    "heading": chk["heading"],
                    "label": chk["label"],
                    "sheet": sheet_map[(rec["slug"], chk["key"])],
                    "anchor": chk["anchor"],
                    "check_range": chk["check_range"],
                    "formula_display": chk["formula"],
                    "formula_stored_xlsx": stored,
                    # Whether this exact check's bytes differ from the ones the
                    # LibreOffice reference run executed. Only true where the
                    # formula names a function xlfn_map rewrites.
                    "differs_from_lo_serialization":
                        stored != to_storage_formula_all(chk["formula"]),
                    "serialization": "plain" if plain_names else "xlfn",
                    "expected": chk["expected"],
                })
            recipe_entries.append({
                "slug": rec["slug"],
                "title": rec["title"],
                "index": rec["index"],
                "n_checks": len(checks),
                "n_variants": len({c["variant_index"] for c in checks
                                   if c["variant_index"] is not None}),
                "checks": checks,
            })

        manifest_chunks.append({
            "chunk": chunk_id,
            "file": os.path.basename(path),
            "bytes": os.path.getsize(path),
            "sha256": _sha256(path),
            "n_recipes": len(group),
            "n_checks": sum(r["n_checks"] for r in recipe_entries),
            "slugs": [r["slug"] for r in recipe_entries],
            "sheets": sorted(sheet_map.values()),
            "recipes": recipe_entries,
        })

    manifest = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "corpus": "recipes",
        "corpus_source": "data/recipes/*.json",
        "chunk_size": args.chunk_size,
        "n_chunks": len(chunks),
        "n_recipes": sum(c["n_recipes"] for c in manifest_chunks),
        "n_checks": sum(c["n_checks"] for c in manifest_chunks),
        "subset_only": sorted(set(args.only)) if args.only else None,
        "plain_names": plain_names,
        "serialization": "plain" if plain_names else "xlfn",
        "manifest_note": getattr(args, "manifest_note", None),
        "skipped_multi_sheet": [
            {"slug": r["slug"], "n_checks": r["n_checks"],
             "setup_sheets": r["setup_sheet_names"]}
            for r in skipped
        ],
        "skipped_multi_sheet_reason": (
            "v1 does not build recipes whose checks need extra worksheets. Their "
            "formulas resolve sheet names from cell VALUES (INDIRECT), span "
            "consecutive tabs (3-D refs), or deliberately assert #NAME?/#REF! for "
            "a wrong-syntax or missing tab, and a single recipe can define the "
            "same tab name twice with different contents -- so neither prefix-"
            "renaming nor one-workbook-per-recipe is correct. They remain "
            "LibreOffice-executed only (results/recipes-verified.json)."
        ),
        "canary": {
            "arithmetic_formula": CANARY_ARITH_FORMULA,
            "arithmetic_expected": CANARY_ARITH_EXPECTED,
            "arithmetic_cell_every_sheet": CANARY_ANCHOR,
            "meta_sheet": META_SHEET,
            "meta_volatile_cell": META_VOLATILE_CELL,
            "meta_volatile_formula": "=NOW()",
            "meta_arithmetic_cell": META_ARITH_CELL,
        },
        "chunks": manifest_chunks,
    }
    manifest_path = os.path.join(outdir, MANIFEST_NAME)
    with open(manifest_path, "w") as f:
        json.dump(manifest, f, indent=2, default=str)
        f.write("\n")

    mode = ("PLAIN NAMES (formulas exactly as authored -- Sheets maps unprefixed "
            "names)" if plain_names else "xlfn-translated")
    print(f"Built {len(chunks)} chunk(s) from {manifest['n_recipes']} recipe(s), "
          f"{manifest['n_checks']} check(s) -> {outdir}  [{mode}]")
    if manifest["manifest_note"]:
        print(f"Manifest note: {manifest['manifest_note']}")
    total = 0
    print(f"  {'file':>16}  {'bytes':>9}  {'recipes':>7}  {'checks':>6}")
    for c in manifest_chunks:
        total += c["bytes"]
        flag = "  <-- OVER SIZE BUDGET" if c["bytes"] > SIZE_WARN_BYTES else ""
        print(f"  {c['file']:>16}  {c['bytes']:>9,}  {c['n_recipes']:>7}  "
              f"{c['n_checks']:>6}{flag}")
    print(f"  {'TOTAL':>16}  {total:>9,} "
          f"(base64 ~{int(total * 4 / 3):,} bytes)")
    if skipped:
        print(f"\nSkipped {len(skipped)} multi-sheet recipe(s) "
              f"({sum(r['n_checks'] for r in skipped)} check(s)) -- these need extra "
              f"worksheets and stay LibreOffice-only; see manifest "
              f"'skipped_multi_sheet_reason':")
        for r in skipped:
            print(f"  {r['slug']:45} {r['n_checks']:>3} check(s)  "
                  f"tabs: {', '.join(r['setup_sheet_names'])}")
    print(f"\nWrote {manifest_path}")
    return manifest


# --------------------------------------------------------------------------
# ingest-recipes
# --------------------------------------------------------------------------

def _read_check_value(ws, entry):
    """Read one check's result back exactly the way verify_recipes.run_case
    returns it, so the two engines' `actual` values are directly comparable:

      * check_range  -> row-major list of norm()'d values with the Nones
                        dropped (an unfilled spill cell reads back as None)
      * otherwise    -> norm() of the single anchor cell

    norm() is recipe_corpus's, i.e. verify_recipes' own."""
    if entry["check_range"]:
        grid = cell_addrs_in_range(entry["check_range"])
        vals = [norm(ws[addr].value) for row in grid for addr in row]
        return [v for v in vals if v is not None]
    return norm(ws[entry["anchor"]].value)


def _readback_artifact_notes(expected, raw, actual):
    """Export-format artifacts worth flagging to a human, never silently
    coerced away. Same policy as the function-corpus ingest."""
    import datetime as _dt
    notes = []
    if isinstance(expected, bool) and isinstance(actual, str) \
            and actual.upper() in ("TRUE", "FALSE"):
        notes.append("NOTE: read back as the TEXT 'TRUE'/'FALSE' rather than a "
                     "boolean cell -- likely an .xlsx export serialization "
                     "artifact, verify before reporting it as an engine difference")
    if isinstance(raw, (_dt.datetime, _dt.date, _dt.time)):
        notes.append(f"NOTE: read back as a date/time-formatted cell ({raw!r}); "
                     f"the underlying value is the serial "
                     f"{normalize_readback_value(raw)!r}")
    if actual == "#ERROR!":
        notes.append("NOTE: #ERROR! is Google Sheets' parse-failure error (no Excel "
                     "equivalent); it means Sheets could not parse the formula, "
                     "which is not the same as #NAME?")
    return notes


def ingest_recipe_exports(export_paths, manifest, engine_label):
    """Read exported workbooks -> (recipe_results, canary, trusted, stats)."""
    plain_names = bool(manifest.get("plain_names", True))
    serialization = manifest.get("serialization", "plain" if plain_names else "xlfn")
    manifest_note = manifest.get("manifest_note")

    recipes_out = {}
    per_chunk = []
    sheet_canary_failures = []
    n_sheet_canaries = 0
    volatile_values = {}
    meta_arith = {}
    n_checks = 0

    chunks_by_id = {c["chunk"]: c for c in manifest["chunks"]}

    for path in export_paths:
        chunk_id = _chunk_id_from_path(path, manifest)
        chunk = chunks_by_id[chunk_id]
        wb = openpyxl.load_workbook(path, data_only=True)

        missing = set(chunk["sheets"]) - set(wb.sheetnames)
        if missing:
            raise SystemExit(
                f"{os.path.basename(path)} claims to be {chunk_id} but is missing "
                f"{len(missing)} of its sheets (e.g. {sorted(missing)[:3]}). "
                f"Wrong file for this chunk, or the manifest is stale -- rebuild."
            )

        if META_SHEET in wb.sheetnames:
            meta = wb[META_SHEET]
            volatile_values[chunk_id] = meta[META_VOLATILE_CELL].value
            meta_arith[chunk_id] = meta[META_ARITH_CELL].value
        else:
            volatile_values[chunk_id] = None
            meta_arith[chunk_id] = None

        for rec in chunk["recipes"]:
            main = None
            variants = {}
            recipe_ok = True
            for entry in rec["checks"]:
                ws = wb[entry["sheet"]]
                n_sheet_canaries += 1
                n_checks += 1
                canary_val = ws[CANARY_ANCHOR].value
                canary_ok = canary_val == CANARY_ARITH_EXPECTED
                if not canary_ok:
                    sheet_canary_failures.append(
                        {"chunk": chunk_id, "sheet": entry["sheet"],
                         "slug": rec["slug"], "value": canary_val})

                raw = (None if entry["check_range"]
                       else ws[entry["anchor"]].value)
                # The comparison is inside the try for the same reason it is in
                # recipe_corpus.run_check(): a list `expected` against a
                # non-iterable `actual` raises, and that is recorded as an ERR
                # result rather than crashing the ingest.
                try:
                    actual = _read_check_value(ws, entry)
                    ok = compare_check(entry["expected"], actual)
                except Exception as e:  # noqa: BLE001
                    actual = f"ERR {e}"
                    ok = False
                ok = bool(ok)

                notes = _readback_artifact_notes(entry["expected"], raw, actual)
                if not canary_ok:
                    notes.insert(0, "UNTRUSTED_RECALC: per-sheet canary failed "
                                    "on this sheet")
                if entry.get("differs_from_lo_serialization"):
                    notes.append(
                        "NOTE: written with the PLAIN function name; the "
                        "LibreOffice reference run executed the _xlfn. storage "
                        "form of this formula, so the two runs are not "
                        "byte-identical inputs")
                if not ok:
                    notes.append(f"MISMATCH vs expected: expected "
                                 f"{entry['expected']!r}, got {actual!r}")
                if not ok:
                    recipe_ok = False

                payload = {
                    "label": entry["label"],
                    "formula": entry["formula_display"],
                    "formula_stored_xlsx": entry["formula_stored_xlsx"],
                    "serialization": entry.get("serialization", serialization),
                    "expected": entry["expected"],
                    "actual": actual,
                    "verified": ok,
                    "sheet": entry["sheet"],
                    "canary_ok_this_sheet": canary_ok,
                    "notes": "; ".join(notes) if notes else None,
                }
                if entry["kind"] == "main":
                    main = payload
                else:
                    variants.setdefault(
                        entry["variant_index"],
                        {"heading": entry["heading"], "checks": []},
                    )["checks"].append(payload)

            if main is None:      # cannot happen for a manifest this tool wrote
                raise SystemExit(f"{rec['slug']}: manifest has no 'main' check")

            record = {
                "verified": recipe_ok,
                "engine": "Google Sheets",
                "engine_label": engine_label,
                "serialization": serialization,
                "formula": main["formula"],
                "formula_stored_xlsx": main["formula_stored_xlsx"],
                "expected": main["expected"],
                "actual": main["actual"],
                "main_verified": main["verified"],
                "sheet": main["sheet"],
                "canary_ok_this_sheet": main["canary_ok_this_sheet"],
                "notes": main["notes"],
            }
            if variants:
                record["variants"] = [variants[i] for i in sorted(variants)]
            recipes_out[rec["slug"]] = record

        per_chunk.append({
            "chunk": chunk_id,
            "export_file": os.path.basename(path),
            "n_recipes": chunk["n_recipes"],
            "n_checks": chunk["n_checks"],
        })

    arith_ok = (not sheet_canary_failures
                and n_sheet_canaries > 0
                and all(v == CANARY_ARITH_EXPECTED for v in meta_arith.values()))

    canary = {
        "arithmetic_formula": CANARY_ARITH_FORMULA,
        "arithmetic_expected": CANARY_ARITH_EXPECTED,
        "arithmetic_actual": (CANARY_ARITH_EXPECTED if arith_ok
                              else (sheet_canary_failures[0]["value"]
                                    if sheet_canary_failures else None)),
        "arithmetic_ok": arith_ok,
        "arithmetic_sheets_checked": n_sheet_canaries,
        "arithmetic_sheet_failures": sheet_canary_failures[:20],
        "meta_arithmetic_per_chunk": meta_arith,
        "volatile_formula": "=NOW()",
        "volatile_per_chunk": {k: str(v) for k, v in volatile_values.items()},
        "now_differs_across_runs": None,
        "method": (
            "openpyxl writes formulas with NO cached <v> value, so the uploaded "
            "chunk workbooks contain zero cached results. Google Drive's "
            "auto-conversion to a Google Sheet recalculates every formula with "
            "Google's engine; exporting that Sheet back to .xlsx carries those "
            "computed values out. The deterministic canary =1111+2222 in "
            f"{CANARY_ANCHOR} of EVERY check's sheet reading back exactly "
            f"{CANARY_ARITH_EXPECTED} therefore proves genuine computation -- "
            "without recalculation the cell would read back blank (None), since "
            "nothing was ever cached. The volatile =NOW() canary is recorded per "
            "chunk as corroboration, but a single Drive import yields one "
            "timestamp, so no cross-run volatile comparison is possible: "
            "now_differs_across_runs is null by design."
        ),
        "engine_label": engine_label,
    }

    stats = {"chunks": per_chunk,
             "n_recipes": len(recipes_out),
             "n_checks": n_checks,
             "serialization": serialization,
             "manifest_note": manifest_note}
    return recipes_out, canary, arith_ok, stats


def write_recipe_results(out_path, recipes, canary, trusted, engine_label,
                         allow_label_change=False, engine=ENGINE_ID,
                         recalc_method=RECIPE_RECALC_METHOD,
                         serialization=None, manifest_note=None):
    """Write (or incrementally merge into) results/recipes-verified-sheets.json.

    Merge semantics are the function corpus's, one level down: a recipe that
    was re-executed is replaced WHOLESALE (its variants come as one block),
    every other recipe is preserved byte-for-byte, and `trusted` is the AND of
    the previous file's flag and this run's, so a merge can only ever downgrade
    trust. Merging under a different engine label is refused without
    --allow-label-change, because Sheets is a continuously-updated hosted
    product and two dates' results should not be blended by accident."""
    generated_at = datetime.now(timezone.utc).isoformat()
    output = {
        "generated_at": generated_at,
        "corpus": "recipes",
        "engine": engine,
        "engine_label": engine_label,
        "recalc_method": recalc_method,
        "trusted": trusted,
        "canary": canary,
        "recipes": recipes,
    }

    if os.path.exists(out_path):
        with open(out_path) as f:
            prev = json.load(f)
        prev_label = prev.get("engine_label")
        if prev_label != engine_label and not allow_label_change:
            sys.exit(
                f"Refusing to merge: {out_path} records engine_label "
                f"{prev_label!r} but this ingest is labelled {engine_label!r}. "
                f"Re-run with the same --engine-label, or pass "
                f"--allow-label-change to record both."
            )
        merged = dict(prev)
        merged_recipes = dict(prev.get("recipes") or {})
        merged_recipes.update(recipes)
        merged["recipes"] = merged_recipes
        merged["generated_at"] = generated_at
        merged["corpus"] = "recipes"
        merged["engine"] = engine
        merged["engine_label"] = engine_label
        merged["recalc_method"] = recalc_method
        merged["canary"] = canary
        merged["trusted"] = bool(prev.get("trusted", False)) and trusted
        if prev_label != engine_label:
            merged.setdefault("engine_label_history", [])
            if prev_label not in merged["engine_label_history"]:
                merged["engine_label_history"].append(prev_label)
            merged["engine_label_history"].append(engine_label)
        merged.setdefault("subset_runs", []).append({
            "generated_at": generated_at,
            "recipes": sorted(recipes),
            "trusted_this_run": trusted,
            "engine_label": engine_label,
            "serialization": serialization,
            "manifest_note": manifest_note,
            "superseded_generated_at": prev.get("generated_at"),
        })
        untouched = len(merged_recipes) - len(recipes)
        print(f"Incremental ingest: merged {len(recipes)} recipe(s) into "
              f"{os.path.basename(out_path)}; {untouched} other recipe result(s) "
              f"preserved unchanged.")
        output = merged

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    with open(out_path, "w") as f:
        json.dump(output, f, indent=2, default=str)
        f.write("\n")
    return output


def _load_recipe_manifest(args):
    path = args.manifest or os.path.join(os.path.abspath(args.chunkdir), MANIFEST_NAME)
    if not os.path.exists(path):
        sys.exit(f"Manifest not found: {path}. Run `run_sheets.py build-recipes` first.")
    with open(path) as f:
        manifest = json.load(f)
    if manifest.get("corpus") != "recipes":
        sys.exit(f"{path} is not a recipe manifest (corpus="
                 f"{manifest.get('corpus', 'functions')!r}). Point --chunkdir at "
                 f"the directory `build-recipes` wrote.")
    return manifest


def cmd_ingest_recipes(args):
    manifest = _load_recipe_manifest(args)
    for p in args.export:
        if not os.path.exists(p):
            sys.exit(f"Export not found: {p}")

    recipes, canary, trusted, stats = ingest_recipe_exports(
        args.export, manifest, args.engine_label)

    write_recipe_results(args.out, recipes, canary, trusted, args.engine_label,
                         args.allow_label_change,
                         serialization=stats["serialization"],
                         manifest_note=stats["manifest_note"])

    n_ok = sum(1 for r in recipes.values() if r["verified"])
    print(f"Ingested {len(stats['chunks'])} chunk(s): {stats['n_recipes']} recipe(s), "
          f"{stats['n_checks']} check(s). [serialization={stats['serialization']}]")
    if stats["manifest_note"]:
        print(f"Manifest note: {stats['manifest_note']}")
    print(f"Deterministic canary OK on all {canary['arithmetic_sheets_checked']} "
          f"sheet(s): {canary['arithmetic_ok']}")
    if canary["arithmetic_sheet_failures"]:
        print(f"  CANARY FAILURES: {canary['arithmetic_sheet_failures']}")
    print(f"Volatile =NOW() per chunk: {canary['volatile_per_chunk']}")
    print(f"Trusted recalculation: {trusted}")
    print(f"Recipes where every check matched its expected result: "
          f"{n_ok} / {len(recipes)}")
    # Name the CHECK that differed, not just the recipe: a recipe is only
    # "verified" when its main example AND all of its variant checks match, so
    # printing the main example's value for a recipe that failed on a variant
    # would read as "got 60, want 60" and look like a harness bug.
    for slug, r in sorted(recipes.items()):
        if r["verified"]:
            continue
        if not r["main_verified"]:
            print(f"  DIFFERS  {slug} [main] {r['formula']}")
            print(f"           got={r['actual']!r} want={r['expected']!r}")
        for vi, var in enumerate(r.get("variants") or []):
            for ci, ch in enumerate(var.get("checks") or []):
                if ch["verified"]:
                    continue
                print(f"  DIFFERS  {slug} [v{vi}c{ci}] {ch['formula']}")
                print(f"           got={ch['actual']!r} want={ch['expected']!r}")
                if ch.get("notes"):
                    print(f"           {ch['notes']}")
    print(f"Wrote {args.out}")
    return recipes


# --------------------------------------------------------------------------
# selftest-recipes
# --------------------------------------------------------------------------

def cmd_selftest_recipes(args):
    """build-recipes -> LibreOffice standing in for Drive -> ingest-recipes,
    then compare every value against results/recipes-verified.json.

    This proves the CELL MAPPING (right formula on the right sheet, right
    anchor, right check_range read back into the right recipe/variant slot),
    nothing about Google Sheets: the values recovered are LibreOffice's, which
    is exactly why they can be checked against the LibreOffice reference run.
    It writes to a scratch file and never into results/."""
    print("=== selftest-recipes: build ===")
    manifest = cmd_build_recipes(args)

    chunkdir = os.path.abspath(args.outdir)
    soffice = os.environ.get("SOFFICE_BIN", "soffice")

    with tempfile.TemporaryDirectory() as tmp:
        export_dir = os.path.join(tmp, "exports")
        os.makedirs(export_dir)
        exports = []
        print("\n=== selftest-recipes: simulated export (soffice --convert-to xlsx) ===")
        for chunk in manifest["chunks"]:
            src = os.path.join(chunkdir, chunk["file"])
            outd = os.path.join(tmp, chunk["chunk"])
            os.makedirs(outd, exist_ok=True)
            proc = subprocess.run(
                [soffice, "--headless", "--convert-to", "xlsx", "--outdir", outd, src],
                capture_output=True, text=True, timeout=900)
            produced = os.path.join(outd, chunk["file"])
            if proc.returncode != 0 or not os.path.exists(produced):
                sys.exit(f"soffice conversion failed for {chunk['file']}: "
                         f"{proc.stderr.strip()}")
            dest = os.path.join(export_dir, f"{chunk['chunk']}-export.xlsx")
            shutil.copyfile(produced, dest)
            exports.append(dest)
            print(f"  {chunk['file']} -> {os.path.basename(dest)} "
                  f"({os.path.getsize(dest):,} bytes)")

        print("\n=== selftest-recipes: ingest ===")
        recipes, canary, trusted, stats = ingest_recipe_exports(
            exports, manifest, args.engine_label)

        out = os.path.abspath(args.out)
        if os.path.commonpath([out, os.path.abspath(RESULTS_DIR)]) == \
                os.path.abspath(RESULTS_DIR):
            sys.exit(f"selftest-recipes refuses to write inside results/ ({out}). "
                     f"These are LibreOffice values produced through the Sheets "
                     f"plumbing, not Google Sheets values.")
        if os.path.exists(out):
            os.remove(out)  # scratch file: always a fresh write, never a merge
        write_recipe_results(out, recipes, canary, trusted, args.engine_label,
                             allow_label_change=True,
                             engine=RECIPE_SELFTEST_ENGINE_ID,
                             recalc_method=(
                                 "selftest plumbing check: soffice --convert-to "
                                 "xlsx standing in for the Drive import -- NOT "
                                 "Google Sheets"),
                             serialization=stats["serialization"],
                             manifest_note=stats["manifest_note"])

    # ---- the actual proof: same values as the LibreOffice reference run ----
    if not os.path.exists(LO_RECIPE_RESULTS_PATH):
        sys.exit(f"{LO_RECIPE_RESULTS_PATH} not found -- run "
                 f"scripts/verify_recipes.py first; the selftest compares "
                 f"against it.")
    with open(LO_RECIPE_RESULTS_PATH) as f:
        reference = json.load(f).get("recipes", {})

    stored_by_key = {}
    for chunk in manifest["chunks"]:
        for rec in chunk["recipes"]:
            for entry in rec["checks"]:
                stored_by_key[(rec["slug"], entry["key"])] = entry

    def _ref_checks(slug):
        """Flatten the reference run into {key: {formula, actual}}."""
        ref = reference.get(slug)
        if ref is None:
            return None
        flat = {"main": {"formula": ref.get("formula"), "actual": ref.get("actual")}}
        for vi, var in enumerate(ref.get("variants") or []):
            for ci, ch in enumerate(var.get("checks") or []):
                flat[f"v{vi}c{ci}"] = {"formula": ch.get("formula"),
                                       "actual": ch.get("actual")}
        return flat

    print("\n=== selftest-recipes: values vs results/recipes-verified.json ===")
    n_match = n_diff = n_skipped = 0
    for slug in sorted(recipes):
        ref = _ref_checks(slug)
        if ref is None:
            print(f"  {slug}: NOT in the LibreOffice reference run -- cannot compare")
            n_diff += 1
            continue
        got = {"main": recipes[slug]["actual"]}
        for vi, var in enumerate(recipes[slug].get("variants") or []):
            for ci, ch in enumerate(var.get("checks") or []):
                got[f"v{vi}c{ci}"] = ch["actual"]
        for key in sorted(got, key=lambda k: (k != "main", k)):
            entry = stored_by_key[(slug, key)]
            mine, theirs = got[key], ref.get(key, {}).get("actual")
            # str() on both sides: the reference file was written with
            # json.dump(default=str), so a value that was a list on the way in
            # comes back as a list of the same reprs -- comparing reprs is the
            # apples-to-apples test.
            same = str(mine) == str(theirs)
            if entry.get("differs_from_lo_serialization"):
                # Not comparable, and honestly so: this check was written with
                # the plain function name while the reference run executed the
                # _xlfn. storage form, i.e. genuinely different input bytes.
                n_skipped += 1
                print(f"  n/c  {slug} {key:8} {entry['formula_display'][:52]:54} "
                      f"plain={mine!r} vs _xlfn-run={theirs!r} "
                      f"(different serialization, not comparable)")
                continue
            if same:
                n_match += 1
            else:
                n_diff += 1
                print(f"  DIFF {slug} {key:8} {entry['formula_display'][:52]:54} "
                      f"got={mine!r} reference={theirs!r}")

    print(f"\nChunks ingested        : {len(stats['chunks'])}")
    print(f"Recipes / checks       : {stats['n_recipes']} / {stats['n_checks']}")
    print(f"Per-sheet canaries     : {canary['arithmetic_sheets_checked']} checked, "
          f"ok={canary['arithmetic_ok']}")
    print(f"Volatile =NOW()        : {canary['volatile_per_chunk']}")
    print(f"trusted                : {trusted}")
    print(f"Values identical to LO : {n_match}")
    print(f"Values differing       : {n_diff}")
    print(f"Not comparable         : {n_skipped} (plain-name serialization)")
    print(f"\nWrote scratch results -> {args.out} "
          f"(LibreOffice values; plumbing proof only)")
    if n_diff:
        sys.exit(f"selftest-recipes FAILED: {n_diff} value(s) do not match the "
                 f"LibreOffice reference run -- the cell mapping is wrong.")
    print("selftest-recipes PASSED: every comparable value round-tripped to the "
          "same result the LibreOffice reference run recorded.")


# --------------------------------------------------------------------------

def main():
    today = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    default_label = f"Google Sheets (Drive import, {today})"

    ap = argparse.ArgumentParser(
        prog="run_sheets.py",
        description="Google Sheets engine runner: build chunked .xlsx workbooks "
                    "for Drive import, and ingest the exported .xlsx readback.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__.split("USAGE\n-----\n", 1)[1].split("\n\nINCREMENTAL")[0],
    )
    sub = ap.add_subparsers(dest="cmd", required=True)

    b = sub.add_parser("build", help="emit chunk-NN.xlsx workbooks + manifest.json")
    b.add_argument("--chunk-size", type=int, default=40,
                   help="max FUNCTIONS per chunk workbook (default 40)")
    b.add_argument("--only", nargs="+", metavar="FN",
                   help="build only these functions (default: all of data/tests)")
    b.add_argument("--outdir", default=DEFAULT_CHUNK_DIR,
                   help=f"output directory (default {DEFAULT_CHUNK_DIR})")
    b.add_argument("--plain-names", action="store_true",
                   help="write formulas EXACTLY as authored in data/tests -- no "
                        "_xlfn./_xlfn._xlws. storage-form translation at all. Use "
                        "for functions Google Sheets recognizes by their bare name "
                        "but whose xlfn/xlws-prefixed storage form its xlsx "
                        "importer does NOT map (verified: FILTER, SORT, and the "
                        "LAMBDA-family functions), which otherwise show a false "
                        "#NAME?/#ERROR! that looks like a real support gap. See "
                        "README.md 'Phase 2: Google Sheets runner'.")
    b.add_argument("--manifest-note", default=None, metavar="TEXT",
                   help="free-text note stored in manifest.json's top-level "
                        "'manifest_note' field and carried into ingest's results "
                        "provenance (subset_runs entries)")
    b.set_defaults(func=cmd_build)

    i = sub.add_parser("ingest", help="read exported .xlsx back into results JSON")
    i.add_argument("--export", nargs="+", required=True, metavar="XLSX",
                   help="exported workbook(s), named <chunk-id>-export.xlsx")
    i.add_argument("--chunkdir", default=DEFAULT_CHUNK_DIR,
                   help="directory holding manifest.json (default %(default)s)")
    i.add_argument("--manifest", help="explicit path to manifest.json")
    i.add_argument("--out", default=DEFAULT_RESULTS_PATH,
                   help=f"results file (default {DEFAULT_RESULTS_PATH})")
    i.add_argument("--engine-label", default=default_label,
                   help="honest engine_version label; Sheets has no version, so "
                        "this records the import DATE (default %(default)r)")
    i.add_argument("--allow-label-change", action="store_true",
                   help="permit merging into a file recorded under a different "
                        "engine label (records both in engine_version_history)")
    i.set_defaults(func=cmd_ingest)

    s = sub.add_parser("selftest",
                       help="dry run: build + LibreOffice-simulated export + "
                            "ingest, to a scratch results file")
    s.add_argument("--chunk-size", type=int, default=40)
    s.add_argument("--only", nargs="+", metavar="FN",
                   default=["COUNT", "SUM", "MROUND", "XLOOKUP", "DATEDIF", "ISNUMBER"])
    s.add_argument("--outdir", default=DEFAULT_CHUNK_DIR)
    s.add_argument("--out", default=SELFTEST_RESULTS_PATH,
                   help="scratch results path (never results/google-sheets.json -- "
                        "these are LibreOffice values, so they must not land in "
                        "results/ alongside real engine results)")
    s.add_argument("--engine-label",
                   default="SELFTEST (LibreOffice values via simulated export "
                           "- NOT Google Sheets)")
    s.add_argument("--allow-label-change", action="store_true", default=True,
                   help=argparse.SUPPRESS)
    s.add_argument("--plain-names", action="store_true",
                   help="dry-run the --plain-names build path too (see `build --help`)")
    s.add_argument("--manifest-note", default=None, metavar="TEXT", help=argparse.SUPPRESS)
    s.set_defaults(func=cmd_selftest)

    # ---- recipe corpus -----------------------------------------------------
    br = sub.add_parser(
        "build-recipes",
        help="emit chunk-NN.xlsx workbooks + manifest.json for the how-to "
             "RECIPE corpus (data/recipes/*.json)",
        description="One worksheet per recipe check (the main worked example, "
                    "plus every variant check), each carrying that check's "
                    "setup_cells, its formula at the same anchor "
                    "scripts/verify_recipes.py uses, and the deterministic "
                    "=1111+2222 canary in Z1. Recipes needing extra worksheets "
                    "(setup_sheets) are SKIPPED and listed -- see the manifest's "
                    "'skipped_multi_sheet_reason'.")
    br.add_argument("--chunk-size", type=int, default=60,
                    help="max RECIPES per chunk workbook (default %(default)s)")
    br.add_argument("--only", nargs="+", metavar="SLUG",
                    help="build only these recipe slugs (default: all of "
                         "data/recipes)")
    br.add_argument("--outdir", default=DEFAULT_RECIPE_CHUNK_DIR,
                    help="output directory (default %(default)s)")
    br.add_argument("--plain-names", dest="plain_names", action="store_true",
                    default=True,
                    help="write every formula EXACTLY as authored, with no "
                         "_xlfn./_xlfn._xlws. storage-form translation. THIS IS "
                         "THE DEFAULT for recipes: Google Sheets maps bare "
                         "modern names on xlsx import but does NOT map "
                         "_xlfn._xlws.FILTER/SORT or the LAMBDA family, and a "
                         "recipe is by definition the formula a user types.")
    br.add_argument("--xlfn-names", dest="plain_names", action="store_false",
                    help="opt out of --plain-names and write the OOXML storage "
                         "form instead (what the LibreOffice reference run "
                         "executes)")
    br.add_argument("--manifest-note", default=None, metavar="TEXT",
                    help="free-text note stored in the manifest and carried into "
                         "ingest provenance")
    br.set_defaults(func=cmd_build_recipes)

    ir = sub.add_parser("ingest-recipes",
                        help="read exported recipe .xlsx back into "
                             "results/recipes-verified-sheets.json")
    ir.add_argument("--export", nargs="+", required=True, metavar="XLSX",
                    help="exported workbook(s), named <chunk-id>-export.xlsx")
    ir.add_argument("--chunkdir", default=DEFAULT_RECIPE_CHUNK_DIR,
                    help="directory holding the recipe manifest.json "
                         "(default %(default)s)")
    ir.add_argument("--manifest", help="explicit path to manifest.json")
    ir.add_argument("--out", default=DEFAULT_RECIPE_RESULTS_PATH,
                    help="results file (default %(default)s)")
    ir.add_argument("--engine-label", default=default_label,
                    help="honest engine label; Sheets has no version, so this "
                         "records the import DATE (default %(default)r)")
    ir.add_argument("--allow-label-change", action="store_true",
                    help="permit merging into a file recorded under a different "
                         "engine label (records both in engine_label_history)")
    ir.set_defaults(func=cmd_ingest_recipes)

    sr = sub.add_parser("selftest-recipes",
                        help="dry run of the recipe pipeline: build + "
                             "LibreOffice-simulated export + ingest, then "
                             "compare every value to results/recipes-verified.json")
    sr.add_argument("--chunk-size", type=int, default=60)
    sr.add_argument("--only", nargs="+", metavar="SLUG",
                    default=list(RECIPE_SELFTEST_SLUGS))
    sr.add_argument("--outdir", default=RECIPE_SELFTEST_CHUNK_DIR)
    sr.add_argument("--out", default=RECIPE_SELFTEST_RESULTS_PATH,
                    help="scratch results path (never inside results/ -- these "
                         "are LibreOffice values)")
    sr.add_argument("--engine-label",
                    default="SELFTEST (LibreOffice values via simulated export "
                            "- NOT Google Sheets)")
    sr.add_argument("--plain-names", dest="plain_names", action="store_true",
                    default=True, help=argparse.SUPPRESS)
    sr.add_argument("--xlfn-names", dest="plain_names", action="store_false",
                    help="dry-run the --xlfn-names build path instead")
    sr.add_argument("--allow-label-change", action="store_true", default=True,
                    help=argparse.SUPPRESS)
    sr.add_argument("--manifest-note", default=None, metavar="TEXT",
                    help=argparse.SUPPRESS)
    sr.set_defaults(func=cmd_selftest_recipes)

    args = ap.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
