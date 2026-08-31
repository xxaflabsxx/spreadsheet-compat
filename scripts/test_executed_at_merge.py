#!/usr/bin/env python3
"""Unit tests for the per-function `executed_at` merge contract.

Fixture-only: no LibreOffice, no Google Sheets, no network. What is under
test is the schema/merge logic that keeps a SUBSET run from re-dating the
whole corpus -- the defect that made every function page read "Last tested
<today>" after a five-function re-run.

    python3 scripts/test_executed_at_merge.py
"""
import json
import os
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "harness"))

from results_schema import (  # noqa: E402
    EXECUTED_AT,
    function_cases,
    function_executed_at,
    stamp_executed_at,
)
import run_lo  # noqa: E402
import run_sheets  # noqa: E402

FAILS = []


def check(name, cond, detail=""):
    print(f"  {'ok  ' if cond else 'FAIL'}  {name}{'' if cond else '  -- ' + detail}")
    if not cond:
        FAILS.append(name)


def case(value=1):
    return {"description": "d", "formula_display": "=F()", "formula_stored_xlsx": "=F()",
            "value": value, "range_values": None, "error": None, "expected": value,
            "matched_expected": True, "canary_ok_this_sheet": True, "notes": None}


def existing_file():
    """A results file as an earlier full run left it."""
    return {
        "generated_at": "2026-07-29T13:50:05.261786+00:00",
        "engine": "LibreOffice Calc",
        "engine_version": "25.8.7.3",
        "recalc_method": "old method",
        "trusted": True,
        "canary": {"arithmetic_ok": True},
        "function_results": {
            "ABS": {EXECUTED_AT: "2026-07-29", "ABS_case": case(5)},
            "SUM": {EXECUTED_AT: "2026-07-29", "SUM_case": case(6)},
            "OLDSTYLE": {"OLD_case": case(7)},   # written before executed_at existed
        },
    }


print("results_schema helpers")
blk = {EXECUTED_AT: "2026-08-31", "F_case": case()}
check("function_cases() hides the metadata key", list(function_cases(blk)) == ["F_case"])
check("function_executed_at() reads the date", function_executed_at(blk) == "2026-08-31")
check("function_executed_at() falls back when absent",
      function_executed_at({"F_case": case()}, "2026-01-01") == "2026-01-01")
fresh = {"MIRR": {"MIRR_case": case()}, "VAR.P": {"VARP_case": case()}}
stamp_executed_at(fresh, "2026-08-31")
check("stamp_executed_at() stamps every block this run produced",
      all(b[EXECUTED_AT] == "2026-08-31" for b in fresh.values()))
check("stamp_executed_at() does not invent cases",
      all(len(function_cases(b)) == 1 for b in fresh.values()))

print("run_lo.py subset merge")
prev = existing_file()
subset = {"SUM": {"SUM_case": case(60)}, "MIRR": {"MIRR_case": case(9)}}
stamp_executed_at(subset, "2026-08-31")
output = {
    "generated_at": "2026-08-31T02:31:06.057472+00:00",
    "engine": "LibreOffice Calc",
    "engine_version": "25.8.7.3",
    "recalc_method": "new method",
    "trusted": True,
    "canary": {"arithmetic_ok": True},
    "function_results": subset,
}
merged = run_lo.merge_subset_run(prev, output, output["canary"], True)
fr = merged["function_results"]
check("untouched function keeps its own executed_at",
      fr["ABS"][EXECUTED_AT] == "2026-07-29", fr["ABS"].get(EXECUTED_AT))
check("untouched function's cases are preserved byte-for-byte",
      fr["ABS"] == prev["function_results"]["ABS"])
check("re-executed function is re-dated to this run",
      fr["SUM"][EXECUTED_AT] == "2026-08-31", fr["SUM"].get(EXECUTED_AT))
check("re-executed function carries this run's values",
      function_cases(fr["SUM"])["SUM_case"]["value"] == 60)
check("newly executed function is dated and merged in",
      fr["MIRR"][EXECUTED_AT] == "2026-08-31" and "MIRR" in fr)
check("pre-executed_at block is left alone rather than back-filled with today",
      EXECUTED_AT not in fr["OLDSTYLE"])
check("file-level generated_at still tracks the newest write",
      merged["generated_at"] == output["generated_at"])
check("subset_runs records the provenance",
      merged["subset_runs"][-1]["functions"] == ["MIRR", "SUM"]
      and merged["subset_runs"][-1]["superseded_generated_at"] == prev["generated_at"])
check("a merge can only downgrade trust",
      run_lo.merge_subset_run(prev, output, output["canary"], False)["trusted"] is False)

print("run_lo.py merge is idempotent for dates")
again = run_lo.merge_subset_run(merged, output, output["canary"], True)
check("re-merging the same subset leaves other functions' dates alone",
      again["function_results"]["ABS"][EXECUTED_AT] == "2026-07-29")

print("run_sheets.py ingest merge")
with tempfile.TemporaryDirectory() as td:
    path = os.path.join(td, "google-sheets.json")
    first = {
        "generated_at": "2026-08-29T16:29:17.599914+00:00",
        "engine": "google_sheets",
        "engine_version": "Google Sheets (Drive import, 2026-08-29)",
        "recalc_method": "Drive import",
        "trusted": True,
        "canary": {"arithmetic_ok": True},
        "function_results": {
            "ABS": {EXECUTED_AT: "2026-08-29", "ABS_case": case(5)},
            "SUM": {EXECUTED_AT: "2026-08-29", "SUM_case": case(6)},
        },
    }
    with open(path, "w") as f:
        json.dump(first, f, indent=2)
    out = run_sheets.write_results(
        path,
        {"SUM": {"SUM_case": case(60)}, "LENB": {"LENB_case": case(3)}},
        {"arithmetic_ok": True}, True,
        "Google Sheets (Drive import, 2026-08-29)",
    )
    fr = out["function_results"]
    today = out["generated_at"][:10]
    check("ingest stamps the functions it ingested",
          fr["SUM"][EXECUTED_AT] == today and fr["LENB"][EXECUTED_AT] == today)
    check("ingest preserves an untouched function's executed_at",
          fr["ABS"][EXECUTED_AT] == "2026-08-29", fr["ABS"].get(EXECUTED_AT))
    check("ingest preserves the untouched function's cases",
          fr["ABS"] == first["function_results"]["ABS"])
    check("summarize() ignores the metadata key",
          run_sheets.summarize(fr) == (3, 0, 0), str(run_sheets.summarize(fr)))

print()
if FAILS:
    print(f"{len(FAILS)} FAILED: {FAILS}")
    sys.exit(1)
print("all executed_at merge tests passed")
