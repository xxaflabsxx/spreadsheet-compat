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

from corpus import (  # noqa: E402
    RESULTS_DIR,
    CANARY_ANCHOR,
    CANARY_ARITH_EXPECTED,
    CANARY_ARITH_FORMULA,
    SHEETS_EXTRA_ERROR_STRINGS,
    assert_sheets_safe_name,
    build_workbook,
    cell_addrs_in_range,
    compare_expected,
    flatten_cases,
    is_error_value,
    load_test_files,
    normalize_readback_value,
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

    args = ap.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
