#!/usr/bin/env python3
"""One-off backfill: stamp every existing function block in results/*.json
with the UTC date it was really executed, derived from git history.

WHY
---
Until harness/results_schema.py existed, a results file recorded only ONE
date: the top-level generated_at, refreshed by every write including subset
runs that re-executed a handful of functions. The site dated every function
page from it, so a five-function subset run silently re-dated ~300 function
pages to "Last tested <today>". This script recovers the true per-function
dates from the only honest record we have -- the git history of the results
files themselves -- and writes them into `executed_at` on each block.

HOW A DATE IS DERIVED (and what it can and cannot know)
------------------------------------------------------
For each results file we replay every commit that touched it, oldest first,
and classify each commit as an execution event or not:

  * generated_at unchanged from the previous commit  -> NOT a run. The blocks
    that changed were hand-edited metadata (e.g. commit ea6849e0 corrected
    DGET/MODE.SNGL `description`/`expected` text without re-executing
    anything: observed values and generated_at were untouched). Such a commit
    must not re-date anything.
  * the commit added subset_runs entries -> a SUBSET run: exactly the
    functions those entries name were executed, at the entry's own
    generated_at. (Later entries win: a function re-run twice is dated by the
    newer run.)
  * otherwise, if a function that was already in the file and is VOLATILE
    (RAND/RANDBETWEEN/RANDARRAY/NOW/TODAY) changed value, the whole file was
    rewritten by a full-corpus run -> every function in that commit's file
    was executed at generated_at. A volatile result cannot come back
    identical from a real run, so this is a reliable full-run fingerprint.
  * otherwise -> a partial run merged into the file before subset_runs
    existed (the Aug 3-6 "un-stub 3 pages" commits work this way): only the
    function blocks whose content changed were executed, at generated_at.

Known limit, deliberately conservative: a partial re-run whose results were
byte-identical to what was already stored is invisible in git, so such a
function keeps its earlier date. This can only ever UNDER-claim recency,
never over-claim it -- the opposite of the defect being fixed.

Renames/copies are NOT followed: results/google-sheets.json was created by
copying a LibreOffice results file, and results/libreoffice-24.2.json was
re-created from libreoffice-25.8.json, so pre-copy history describes a
different engine's executions and must not leak into these dates.

Usage:  python3 scripts/backfill_executed_at.py [--dry-run]
"""
import json
import os
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "harness"))
from results_schema import EXECUTED_AT, function_cases  # noqa: E402

FILES = [
    "results/libreoffice-24.2.json",
    "results/libreoffice-24.8.json",
    "results/libreoffice-25.2.json",
    "results/libreoffice-25.8.json",
    "results/google-sheets.json",
]

# A real run can never reproduce these byte-for-byte, so a commit in which an
# already-present one of them changed is a full-corpus rewrite.
VOLATILE = {"RAND", "RANDBETWEEN", "RANDARRAY", "NOW", "TODAY"}


def git(*args):
    p = subprocess.run(["git"] + list(args), cwd=ROOT,
                       capture_output=True, text=True)
    if p.returncode:
        return None
    return p.stdout


def blob_at(commit, path):
    out = git("show", f"{commit}:{path}")
    return json.loads(out) if out is not None else None


def block_key(block):
    """Comparable content of one function block, ignoring our own metadata."""
    return json.dumps(function_cases(block), sort_keys=True)


def derive(path, verbose=True):
    """function name -> (executed_at date, provenance string)."""
    commits = (git("log", "--reverse", "--format=%H", "--", path) or "").split()
    dates, why = {}, {}
    prev_fr, prev_ga, prev_subs = {}, None, set()
    for c in commits:
        d = blob_at(c, path)
        if d is None:           # path absent at this commit (pre-rename)
            continue
        fr = d.get("function_results", {})
        ga = d.get("generated_at") or ""
        subs = d.get("subset_runs") or []
        new_subs = [s for s in subs if s.get("generated_at") not in prev_subs]
        changed = {fn for fn, b in fr.items()
                   if block_key(b) != block_key(prev_fr.get(fn) or {})}

        if ga == prev_ga:
            # No new execution recorded -- metadata-only edit. Leave dates be.
            if verbose and changed:
                print(f"    {c[:8]} {ga[:10]} metadata-only edit, NOT dated: "
                      f"{sorted(changed)}")
        elif new_subs:
            for s in new_subs:                     # later entries win
                for fn in s.get("functions", []):
                    if fn in fr:
                        dates[fn] = (s.get("generated_at") or ga)[:10]
                        why[fn] = f"{c[:8]} subset run"
            uncovered = changed - {fn for s in new_subs
                                   for fn in s.get("functions", [])}
            for fn in uncovered:                   # shouldn't happen; be loud
                dates[fn] = ga[:10]
                why[fn] = f"{c[:8]} changed outside any subset_runs entry"
                if verbose:
                    print(f"    {c[:8]} WARNING: {fn} changed but is in no "
                          f"subset_runs entry; dated from generated_at")
        elif (VOLATILE & set(prev_fr)) & changed:
            for fn in fr:                          # full-corpus rewrite
                dates[fn] = ga[:10]
                why[fn] = f"{c[:8]} full run"
        else:
            for fn in changed:                     # pre-subset_runs partial
                dates[fn] = ga[:10]
                why[fn] = f"{c[:8]} partial run (merged, pre-subset_runs)"

        prev_fr, prev_ga = fr, ga
        prev_subs = {s.get("generated_at") for s in subs}
    return dates, why


def main():
    dry = "--dry-run" in sys.argv
    for path in FILES:
        full = os.path.join(ROOT, path)
        if not os.path.exists(full):
            continue
        print(f"== {path}")
        dates, why = derive(path)
        with open(full) as f:
            doc = json.load(f)
        fr = doc.get("function_results", {})
        missing = [fn for fn in fr if fn not in dates]
        if missing:
            print(f"    WARNING: no execution date derived for {missing} "
                  f"-- falling back to generated_at")
        for fn, block in fr.items():
            block[EXECUTED_AT] = dates.get(fn, (doc.get("generated_at") or "")[:10])
            # keep the date first for readability
            fr[fn] = {EXECUTED_AT: block.pop(EXECUTED_AT), **block}
        hist = {}
        for fn in fr:
            hist[fr[fn][EXECUTED_AT]] = hist.get(fr[fn][EXECUTED_AT], 0) + 1
        for d in sorted(hist):
            print(f"    {d}  {hist[d]:>4} function(s)")

        # Google Sheets: the file's own subset_runs carry dated engine labels
        # ("Google Sheets (Drive import, 2026-08-29)"). Cross-check them.
        if "google" in path:
            for s_run in doc.get("subset_runs", []):
                lbl = s_run.get("engine_label") or ""
                lbl_date = lbl.rstrip(")")[-10:]
                covered = [fn for fn in s_run.get("functions", []) if fn in fr]
                earlier = [fn for fn in covered if fr[fn][EXECUTED_AT] < lbl_date]
                later = [fn for fn in covered if fr[fn][EXECUTED_AT] > lbl_date]
                print(f"    label check {lbl!r}: {len(covered)} function(s), "
                      f"{len(covered) - len(earlier) - len(later)} agree, "
                      f"{len(later)} re-executed later (fine), "
                      f"{len(earlier)} DISAGREE")
                if earlier:
                    print(f"      DISAGREEMENT: {sorted(earlier)}")
            # Every function must also be >= the date of the first Sheets run.
            first = min(fr[fn][EXECUTED_AT] for fn in fr)
            print(f"    earliest Sheets executed_at: {first}")
        if not dry:
            with open(full, "w") as f:
                json.dump(doc, f, indent=2, default=str)
                f.write("\n")
            print(f"    wrote {path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
