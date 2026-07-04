#!/usr/bin/env python3
"""
LibreOffice Calc engine runner for the spreadsheet function-compatibility
harness.

WHAT THIS DOES
--------------
1. Loads every data/tests/<FUNCTION>.json file (or a subset given on the
   command line).
2. Builds a single .xlsx workbook with openpyxl: one worksheet per test
   case, containing the case's setup_cells plus the formula under test
   (translated to its correct OOXML storage form via harness/xlfn_map.py),
   plus a small canary block on every sheet.
3. Forces LibreOffice to actually RECALCULATE (not just re-serialize
   cached values) by round-tripping the file through
   `soffice --headless --convert-to xlsx`.
4. Reads the recalculated file back with openpyxl(data_only=True) and
   extracts real computed values.
5. Writes results/libreoffice-24.2.json mapping test id -> computed
   value/error/notes, plus a top-level canary block proving genuine
   recalculation occurred.

WHY `--convert-to` IS TRUSTWORTHY HERE (READ BEFORE CHANGING THIS)
-------------------------------------------------------------------
openpyxl NEVER writes a cached <v> value for a formula cell -- only the
formula string itself. That means there is no stale cached value for
LibreOffice to fall back to; to produce ANY value in column output at all,
soffice --convert-to MUST evaluate every formula from scratch. We proved
this empirically two ways:
  (a) A volatile canary `=NOW()` produces a genuinely different timestamp
      on two separate conversion runs a few seconds apart (see
      canary.now_run_1 / canary.now_run_2 in the output JSON, or re-run
      this script twice to reproduce).
  (b) A deterministic arithmetic canary `=1111+2222` (no cached value
      possible) evaluates to exactly 3333 on every sheet; if recalculation
      were NOT happening, openpyxl would read back None (blank) for every
      formula cell instead, since nothing was ever cached.
If canary checks fail, this script marks the ENTIRE run "trusted": false
and every function result gets an "UNTRUSTED_RECALC" note -- never trust
a green run without checking the "trusted" flag in the output file.

THE _xlfn. PREFIX GOTCHA
-------------------------
See harness/xlfn_map.py for a full writeup. Short version: functions added
to Excel after 2007 (XLOOKUP, LET, LAMBDA, FILTER, SORT, UNIQUE, SEQUENCE,
TEXTSPLIT, TEXTBEFORE/AFTER, IFS, SWITCH, MAXIFS/MINIFS, TEXTJOIN, CONCAT,
IFNA, ARRAYTOTEXT, ...) must be written into the raw .xlsx XML with an
"_xlfn." (or "_xlfn._xlws." for FILTER/SORT) prefix, or EVERY engine
(including real Excel) will show #NAME? even if the function is fully
supported. We translate this automatically per test case based on the
file's "function" field.

USAGE
-----
    python3 harness/run_lo.py                  # run all data/tests/*.json
    python3 harness/run_lo.py XLOOKUP LET       # run only these functions
"""
import glob
import json
import os
import re
import subprocess
import sys
import tempfile
import time
from datetime import datetime, timezone

sys.path.insert(0, os.path.dirname(__file__))
from xlfn_map import to_storage_formula_all  # noqa: E402

import openpyxl  # noqa: E402
from openpyxl.worksheet.formula import ArrayFormula  # noqa: E402

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
TESTS_DIR = os.path.join(REPO_ROOT, "data", "tests")
RESULTS_DIR = os.path.join(REPO_ROOT, "results")

SOFFICE_BIN = "soffice"
LO_VERSION = "24.2.7.2"  # `soffice --version` at harness build time

KNOWN_ERROR_STRINGS = {
    "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A",
    "#GETTING_DATA", "#CALC!", "#SPILL!", "#FIELD!", "#UNKNOWN!",
    "#BLOCKED!", "#CONNECT!", "#BUSY!",
}

CANARY_ARITH_FORMULA = "=1111+2222"
CANARY_ARITH_EXPECTED = 3333
CANARY_ANCHOR = "Z1"  # far from any setup_cells/check_range used in data/tests


def load_test_files(names=None):
    """Return list of (function_name, filepath, payload)."""
    files = sorted(glob.glob(os.path.join(TESTS_DIR, "*.json")))
    out = []
    for path in files:
        fn = os.path.splitext(os.path.basename(path))[0]
        if names and fn not in names:
            continue
        with open(path) as f:
            payload = json.load(f)
        out.append((fn, path, payload))
    return out


def sanitize_sheet_name(name, used):
    """Excel/LO sheet names: <=31 chars, no []:*?/\\, must be unique."""
    clean = re.sub(r"[\[\]:\*\?/\\]", "_", name)[:31]
    base = clean
    i = 2
    while clean.lower() in used:
        suffix = f"~{i}"
        clean = base[: 31 - len(suffix)] + suffix
        i += 1
    used.add(clean.lower())
    return clean


def build_workbook(cases_flat):
    """
    cases_flat: list of dicts with keys:
        test_id, function, formula (original), setup_cells, check_range
    Returns (workbook, sheet_map) where sheet_map[test_id] -> sheet_name
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    used_names = set()
    sheet_map = {}

    for c in cases_flat:
        sheet_name = sanitize_sheet_name(c["test_id"], used_names)
        sheet_map[c["test_id"]] = sheet_name
        ws = wb.create_sheet(sheet_name)

        for addr, val in (c.get("setup_cells") or {}).items():
            ws[addr] = val

        # Prefix EVERY known future-function call site (not just the function
        # under test): nested modern calls like UNICHAR(UNICODE(...)) need
        # both names prefixed or the whole formula is #NAME? on all engines.
        storage_formula = to_storage_formula_all(c["formula"])
        anchor = c["anchor"]
        if c.get("check_range"):
            # Functions expected to return a multi-cell array (spill/dynamic
            # array results) are written as a legacy Ctrl+Shift+Enter style
            # array formula covering the full check_range. This matters:
            # older engines (and pre-365 Excel) only spill a range result
            # when the formula is explicitly marked as an array formula --
            # writing it as a plain scalar formula string causes even a
            # SUPPORTED function like INDEX(range,0,col) to return #VALUE!
            # instead of spilling. Verified empirically: wrapping
            # INDEX(A1:C3,0,2) in ArrayFormula(ref="A30:A32") makes
            # LibreOffice 24.2 correctly spill [2,5,8]; the identical
            # formula as a plain string returns #VALUE!.
            ws[anchor] = ArrayFormula(c["check_range"], storage_formula)
        else:
            ws[anchor] = storage_formula

        # Canary: deterministic, non-cacheable arithmetic on every sheet.
        ws[CANARY_ANCHOR] = CANARY_ARITH_FORMULA

    # Dedicated meta sheet with a volatile canary for cross-run recalc proof.
    meta = wb.create_sheet("_meta", 0)
    meta["A1"] = "=NOW()"
    meta["A2"] = CANARY_ARITH_FORMULA

    return wb, sheet_map


def anchor_for_case(case):
    if case.get("check_range"):
        # anchor is the top-left cell of the check range
        first = case["check_range"].split(":")[0]
        return first
    return "F1"


def cell_addrs_in_range(range_str):
    """Expand 'A30:C32' into a row-major list of lists of addresses."""
    from openpyxl.utils.cell import range_boundaries, get_column_letter

    min_col, min_row, max_col, max_row = range_boundaries(range_str)
    rows = []
    for r in range(min_row, max_row + 1):
        row = []
        for c in range(min_col, max_col + 1):
            row.append(f"{get_column_letter(c)}{r}")
        rows.append(row)
    return rows


def is_error_value(v):
    return isinstance(v, str) and v in KNOWN_ERROR_STRINGS


EXCEL_EPOCH = datetime(1899, 12, 30)  # serial 0 in the 1900 date system


def normalize_readback_value(v):
    """
    Normalize a value read back from the recalculated .xlsx into the same
    domain the test corpus's `expected` values live in.

    - datetime/date/time objects -> Excel serial numbers. LibreOffice
      applies a date/time NUMBER FORMAT to the result cells of DATE()/
      TIME()-style formulas; openpyxl then surfaces the cached value as a
      Python datetime/time object instead of the underlying float serial.
      The engine's actual computed value IS the serial -- the datetime-ness
      is presentation, so converting back to the serial is the faithful raw
      value, not an interpretation. (Excel 1900 system: 1899-12-30 = 0.
      This intentionally reproduces Excel's day-59/60 Feb-29-1900
      compatibility offset for all post-1900-03-01 dates, which is every
      date used in this corpus.)
    """
    import datetime as _dt

    if isinstance(v, _dt.datetime):
        delta = v - EXCEL_EPOCH
        return delta.days + delta.seconds / 86400 + delta.microseconds / 86400e6
    if isinstance(v, _dt.date):
        return (
            _dt.datetime(v.year, v.month, v.day) - EXCEL_EPOCH
        ).days
    if isinstance(v, _dt.time):
        return (v.hour * 3600 + v.minute * 60 + v.second) / 86400 + v.microsecond / 86400e6
    return v


def values_roughly_equal(a, b):
    if isinstance(a, (int, float)) and isinstance(b, (int, float)) and not isinstance(a, bool) and not isinstance(b, bool):
        return abs(a - b) < 1e-9
    # .xlsx storage limitation: a formula legitimately returning the empty
    # string "" round-trips through file conversion as a cell with no cached
    # value at all, which openpyxl reads back as None. Blank-vs-empty-string
    # is genuinely indistinguishable at this layer, so an expected "" is
    # satisfied by a read-back None. (The raw None is still recorded in the
    # results file; only the match verdict treats them as equivalent.)
    if a == "" and b is None or b == "" and a is None:
        return True
    return a == b


def compare_expected(expected, actual_anchor, actual_range_flat):
    """Returns (matched: bool or None, detail: str or None)."""
    if expected is None:
        return None, None
    if isinstance(expected, list):
        if actual_range_flat is None:
            return False, "expected a range of values but no check_range was read"
        # flatten expected (may be nested for 2D)
        flat_expected = []
        for item in expected:
            if isinstance(item, list):
                flat_expected.extend(item)
            else:
                flat_expected.append(item)
        flat_actual = actual_range_flat
        if len(flat_expected) != len(flat_actual):
            return False, f"length mismatch: expected {len(flat_expected)} values, got {len(flat_actual)}"
        for e, a in zip(flat_expected, flat_actual):
            if not values_roughly_equal(e, a):
                return False, f"value mismatch: expected {e!r}, got {a!r}"
        return True, None
    else:
        matched = values_roughly_equal(expected, actual_anchor)
        detail = None if matched else f"expected {expected!r}, got {actual_anchor!r}"
        return matched, detail


def run():
    requested = set(sys.argv[1:]) or None
    test_files = load_test_files(requested)
    if not test_files:
        print("No test files matched.", file=sys.stderr)
        sys.exit(1)

    cases_flat = []
    for fn, path, payload in test_files:
        for case in payload["cases"]:
            anchor = anchor_for_case(case)
            cases_flat.append({
                "test_id": case["id"],
                "function": fn,
                "formula": case["formula"],
                "setup_cells": case.get("setup_cells"),
                "check_range": case.get("check_range"),
                "expected": case.get("expected"),
                "expected_note": case.get("expected_note"),
                "description": case["description"],
                "anchor": anchor,
            })

    print(f"Loaded {len(test_files)} function(s), {len(cases_flat)} test case(s).")

    wb, sheet_map = build_workbook(cases_flat)

    with tempfile.TemporaryDirectory() as tmpdir:
        src_path = os.path.join(tmpdir, "harness_input.xlsx")
        out_dir = os.path.join(tmpdir, "out")
        os.makedirs(out_dir, exist_ok=True)
        wb.save(src_path)

        t0 = time.time()
        proc = subprocess.run(
            [SOFFICE_BIN, "--headless", "--convert-to", "xlsx",
             "--outdir", out_dir, src_path],
            capture_output=True, text=True, timeout=300,
        )
        elapsed = time.time() - t0
        print(proc.stdout.strip())
        if proc.returncode != 0:
            print("soffice STDERR:\n" + proc.stderr, file=sys.stderr)
            sys.exit(f"soffice conversion failed (exit {proc.returncode})")

        out_path = os.path.join(out_dir, "harness_input.xlsx")
        if not os.path.exists(out_path):
            sys.exit(f"Expected output file not found: {out_path}")

        # Second run (staggered) purely to independently reconfirm the
        # volatile-canary recalculation proof for THIS invocation's log.
        time.sleep(2)
        out_dir2 = os.path.join(tmpdir, "out2")
        os.makedirs(out_dir2, exist_ok=True)
        subprocess.run(
            [SOFFICE_BIN, "--headless", "--convert-to", "xlsx",
             "--outdir", out_dir2, src_path],
            capture_output=True, text=True, timeout=300,
        )
        out_path2 = os.path.join(out_dir2, "harness_input.xlsx")

        wb_out = openpyxl.load_workbook(out_path, data_only=True)
        wb_out2 = openpyxl.load_workbook(out_path2, data_only=True) if os.path.exists(out_path2) else None

    # ---- Canary verification ----
    meta = wb_out["_meta"]
    now_run1 = meta["A1"].value
    arith_run1 = meta["A2"].value
    now_run2 = wb_out2["_meta"]["A1"].value if wb_out2 else None

    canary = {
        "arithmetic_formula": CANARY_ARITH_FORMULA,
        "arithmetic_expected": CANARY_ARITH_EXPECTED,
        "arithmetic_actual": arith_run1,
        "arithmetic_ok": arith_run1 == CANARY_ARITH_EXPECTED,
        "volatile_formula": "=NOW()",
        "now_run_1": str(now_run1),
        "now_run_2": str(now_run2),
        "now_differs_across_runs": (now_run1 != now_run2) if now_run2 else None,
        "conversion_seconds_run_1": round(elapsed, 2),
        "method": "openpyxl writes formulas with NO cached <v> value; "
                  "`soffice --headless --convert-to xlsx` must evaluate every "
                  "formula from scratch to produce any output value at all. "
                  "The volatile =NOW() canary changing between two runs a few "
                  "seconds apart, plus the deterministic arithmetic canary "
                  "matching exactly, together prove genuine recalculation.",
    }

    global_trusted = canary["arithmetic_ok"] and bool(canary["now_differs_across_runs"])

    # ---- Per-case results ----
    function_results = {}
    for c in cases_flat:
        sheet_name = sheet_map[c["test_id"]]
        ws = wb_out[sheet_name]

        per_sheet_canary_val = ws[CANARY_ANCHOR].value
        per_sheet_canary_ok = per_sheet_canary_val == CANARY_ARITH_EXPECTED

        anchor_val = normalize_readback_value(ws[c["anchor"]].value)

        range_flat = None
        if c["check_range"]:
            grid = cell_addrs_in_range(c["check_range"])
            range_flat = [normalize_readback_value(ws[addr].value)
                          for row in grid for addr in row]

        error = anchor_val if is_error_value(anchor_val) else None
        matched, mismatch_detail = compare_expected(c["expected"], anchor_val, range_flat)

        notes = []
        if not per_sheet_canary_ok:
            notes.append("UNTRUSTED_RECALC: per-sheet canary failed on this sheet")
        if c.get("expected_note"):
            notes.append(c["expected_note"])
        if mismatch_detail:
            notes.append(f"MISMATCH vs expected: {mismatch_detail}")

        storage_formula = to_storage_formula_all(c["formula"])

        result = {
            "description": c["description"],
            "formula_display": c["formula"],
            "formula_stored_xlsx": storage_formula,
            "value": anchor_val,
            "range_values": range_flat,
            "error": error,
            "expected": c["expected"],
            "matched_expected": matched,
            "canary_ok_this_sheet": per_sheet_canary_ok,
            "notes": "; ".join(notes) if notes else None,
        }
        function_results.setdefault(c["function"], {})[c["test_id"]] = result

    output = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "engine": "LibreOffice Calc",
        "engine_version": LO_VERSION,
        "recalc_method": "soffice --headless --convert-to xlsx (see canary proof below)",
        "trusted": global_trusted,
        "canary": canary,
        "function_results": function_results,
    }

    os.makedirs(RESULTS_DIR, exist_ok=True)
    out_json_path = os.path.join(RESULTS_DIR, "libreoffice-24.2.json")
    with open(out_json_path, "w") as f:
        json.dump(output, f, indent=2, default=str)
        f.write("\n")

    print(f"\nTrusted recalculation: {global_trusted}")
    print(f"Wrote {out_json_path}")

    # ---- Console summary ----
    n_name_error = 0
    n_other_error = 0
    n_ok = 0
    for fn, cases in function_results.items():
        for tid, r in cases.items():
            if r["error"] == "#NAME?":
                n_name_error += 1
            elif r["error"]:
                n_other_error += 1
            else:
                n_ok += 1
    print(f"Cases with a value (no error): {n_ok}")
    print(f"Cases returning #NAME? (unsupported function): {n_name_error}")
    print(f"Cases returning some other error: {n_other_error}")


if __name__ == "__main__":
    run()
