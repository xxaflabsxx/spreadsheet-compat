"""Shared shape of the executed-results files (results/<engine>-*.json).

One results file is: top-level provenance (generated_at, engine,
engine_version, recalc_method, trusted, canary, subset_runs) plus

    "function_results": {
        "<FUNCTION>": {
            "executed_at": "YYYY-MM-DD",     <-- per-function metadata
            "<CASE_ID>": { ...one executed case... },
            ...
        },
        ...
    }

WHY A PER-FUNCTION DATE (READ BEFORE REMOVING IT)
-------------------------------------------------
Top-level `generated_at` is FILE-level provenance: it is refreshed by every
run that writes the file, including a SUBSET run that re-executed five
functions and merged them in. Dating every function page from it therefore
claims "Last tested <today>" for ~300 functions when only a handful were
actually executed that day. `executed_at` records, per function, the UTC
date of the run that produced THAT function's cases; merges preserve it for
functions they do not touch. Anything that prints a per-function "last
tested" date must read `executed_at` and fall back to `generated_at` only
when it is absent (files written before this key existed).

Because the per-function metadata lives alongside the case ids, every
consumer that iterates a function block's cases must go through
`function_cases()` (or otherwise skip FUNCTION_META_KEYS) — a raw
`block.items()` would hand back the date string as if it were a test case.
"""

EXECUTED_AT = "executed_at"

# Keys inside a function block that are metadata about the block, NOT cases.
FUNCTION_META_KEYS = frozenset({EXECUTED_AT})


def function_cases(block):
    """The executed cases of one function block: {case_id: case_dict}."""
    if not block:
        return {}
    return {k: v for k, v in block.items() if k not in FUNCTION_META_KEYS}


def function_executed_at(block, default=None):
    """UTC date (YYYY-MM-DD) that this function's cases were executed."""
    if not block:
        return default
    return block.get(EXECUTED_AT) or default


def stamp_executed_at(function_results, date_str):
    """Stamp every function block this run produced with its execution date.

    Called on the freshly-built function_results, BEFORE any merge into an
    existing file: the merge replaces whole function blocks, so re-executed
    functions carry this run's date in and every untouched function keeps
    the date already recorded for it.
    """
    for block in function_results.values():
        block[EXECUTED_AT] = date_str
    return function_results
