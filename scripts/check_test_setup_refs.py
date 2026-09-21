#!/usr/bin/env python3
"""Audit: every cell a test formula reads must actually be set up.

WHY THIS EXISTS
---------------
A test case in data/tests/*.json is a formula plus the `setup_cells` that
formula reads. If the formula names a cell the case never populated, the
engine dutifully reads a BLANK, computes something -- usually a wrong-looking
but perfectly deterministic number, sometimes an error -- and the harness
records a "divergence" that is really a typo in our own fixture. That is not
hypothetical: the alphabetical corpus batch E (IMLOG10..MMULT) authored five
cases whose formulas referenced rows that were never set, and all five would
have been published as LibreOffice divergences had an ad-hoc audit not caught
them before the run. This script is that audit, made reusable and mandatory:
run it BEFORE any LibreOffice execution.

WHAT COUNTS AS A VIOLATION
--------------------------
  * A single-cell reference (A2, $B$7) that is not a key of the case's
    setup_cells.
  * A range reference (A2:B3) with NO populated cell anywhere inside it.
  * Any sheet-qualified reference (Sheet2!A1, 'My Sheet'!A1:B2). The harness
    puts each case on its own sheet -- see harness/corpus.py build_workbook --
    so a cross-sheet reference can only ever resolve to nothing.

WHAT IS DELIBERATE, NOT A VIOLATION
-----------------------------------
Two things, and both have to be SAID somewhere the reader can see.

1. THE CASE SAYS IT IS TESTING AN EMPTY CELL. Plenty of cases read a blank on
   purpose -- that is the entire content of SUM_empty_range, COUNT_all_blank_
   range, ISNUMBER_blank, AVERAGE_empty_range_div0. Such a case is accepted
   when its own words say so: an emptiness word (empty / blank / unset / no
   data / not set) in the case id, the description or the expected_note. The
   test that documents itself as an empty-cell test is doing what it says; the
   one that quietly reads an unset cell is the bug this script exists to
   catch.

2. THE REFERENCE IS CONSUMED AS A REFERENCE, NOT AS A VALUE. ROW(A9),
   COLUMNS(A1:C4), CELL("row",B7), AREAS((B2:D4,E5)), OFFSET(A1,2,0),
   ISREF(...), FORMULATEXT(A1), GETPIVOTDATA("Sales",J20) all answer a
   question about the reference's SHAPE or ADDRESS. Populating those cells
   would not change a single expected value, so requiring it would be
   ceremony. A reference sitting inside such a call is therefore a warning,
   never a violation -- and only inside it: SUM(A1)+ROW(A9) still fails on A1.

WHAT IS ONLY A WARNING
----------------------
A range that is partially populated (SUM(A1:A10) over three set cells) is
normal and idiomatic, so it is reported with -v/--verbose only and never
fails the run.

WHAT THIS CANNOT SEE
--------------------
String literals are stripped before parsing, so a reference that only exists
inside a string -- INDIRECT("A1"), or a text argument that happens to look
like an address -- is not checked. That is the conservative direction: it can
miss a bad reference, it can never invent one.

Usage:
    python3 scripts/check_test_setup_refs.py               # audit all of data/tests
    python3 scripts/check_test_setup_refs.py MMULT PRICE   # audit named functions
    python3 scripts/check_test_setup_refs.py --only-batch  # (with names) fail only
                                                           # on the named functions,
                                                           # report the rest
    python3 scripts/check_test_setup_refs.py -v            # also show warnings

Exit status: 1 if any violation is found (in the failing scope), else 0.
"""
import argparse
import glob
import json
import os
import re
import sys

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
TESTS_DIR = os.path.join(REPO_ROOT, "data", "tests")

# A1-style reference, optionally $-anchored, optionally a range. The
# lookbehind rejects a preceding letter/digit/dot/underscore so neither a
# function name (IMLOG10) nor scientific notation (1E5) nor an already-dotted
# token can masquerade as a reference; the lookahead rejects a following "("
# so a name like LOG10(...) is read as a call, not as cell LOG10.
_REF = re.compile(
    r"(?<![A-Za-z0-9_.$!])"
    r"(\$?[A-Za-z]{1,3}\$?[0-9]{1,7})"
    r"(?::(\$?[A-Za-z]{1,3}\$?[0-9]{1,7}))?"
    r"(?![A-Za-z0-9_(])"
)
# Sheet-qualified reference: Sheet1!A1 or 'My Sheet'!A1:B2
_SHEET_REF = re.compile(r"(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!\$?[A-Za-z]{1,3}\$?[0-9]{1,7}")
_STRING = re.compile(r'"[^"]*"')
# Excel error literals contain no cell refs but do contain letters+punctuation.
_ERRLIT = re.compile(r"#[A-Z0-9_/]+[!?]")

# Words that make a case an admitted empty-cell test. Deliberately narrow:
# "omitted" and "missing" appear all over this corpus's notes in their ordinary
# documentation sense ("If num_chars is omitted, it is assumed to be 1"), so
# they are accepted only in the strong, ADDRESS-SPECIFIC form below, never as a
# blanket excuse for the whole case.
_EMPTY_WORDS = re.compile(
    r"\b(empty|blank|unset|unpopulated|not set|no data)\b", re.I
)
_EMPTY_WORDS_NEAR_ADDRESS = re.compile(
    r"\b(empty|blank|unset|unpopulated|not set|no data|no value|nothing|missing|omitted)\b",
    re.I,
)
_EXPLICIT_EMPTY = re.compile(
    r"\b(deliberately|intentionally|on purpose|purposely|left)\s+(?:\w+\s+){0,3}(empty|blank|unset)\b",
    re.I,
)
# Booleans and a few bare words the reference pattern can never match anyway,
# kept explicit so the intent is documented rather than accidental.
_NOT_A_REF = {"TRUE", "FALSE"}

# Functions that read a reference's shape or address rather than its contents,
# so an unset cell inside their argument list changes nothing. Kept as an
# explicit, auditable list: adding a name here is a claim that the function
# ignores cell VALUES, and that claim should be checkable against the
# function's documentation.
_REFERENCE_SHAPE_FUNCTIONS = {
    "ROW", "ROWS", "COLUMN", "COLUMNS", "AREAS", "CELL", "OFFSET", "ISREF",
    "FORMULATEXT", "ADDRESS", "GETPIVOTDATA", "SHEET", "SHEETS",
}


def col_to_num(col):
    n = 0
    for ch in col.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def split_addr(addr):
    a = addr.replace("$", "").upper()
    m = re.match(r"^([A-Z]{1,3})([0-9]{1,7})$", a)
    if not m:
        return None
    return m.group(1), int(m.group(2))


def normalize(addr):
    parts = split_addr(addr)
    return None if parts is None else f"{parts[0]}{parts[1]}"


def expand_range(a, b):
    pa, pb = split_addr(a), split_addr(b)
    if not pa or not pb:
        return []
    c1, r1 = col_to_num(pa[0]), pa[1]
    c2, r2 = col_to_num(pb[0]), pb[1]
    if c1 > c2:
        c1, c2 = c2, c1
    if r1 > r2:
        r1, r2 = r2, r1
    # A pathological range (whole column, huge block) is not worth expanding;
    # the caller only needs "is anything in here populated?".
    if (c2 - c1 + 1) * (r2 - r1 + 1) > 20000:
        return None
    out = []
    for c in range(c1, c2 + 1):
        letters = ""
        n = c
        while n:
            n, rem = divmod(n - 1, 26)
            letters = chr(65 + rem) + letters
        for r in range(r1, r2 + 1):
            out.append(f"{letters}{r}")
    return out


def reference_shape_spans(text):
    """Character spans of every _REFERENCE_SHAPE_FUNCTIONS call's arguments."""
    spans = []
    for m in re.finditer(r"(?<![A-Za-z0-9_.])([A-Za-z][A-Za-z0-9_.]*)\s*\(", text):
        if m.group(1).upper() not in _REFERENCE_SHAPE_FUNCTIONS:
            continue
        depth, i = 0, m.end() - 1
        while i < len(text):
            if text[i] == "(":
                depth += 1
            elif text[i] == ")":
                depth -= 1
                if depth == 0:
                    break
            i += 1
        spans.append((m.end(), i))
    return spans


def refs_in_formula(formula):
    """Return (single_refs, ranges, sheet_qualified) found in `formula`.

    Each single/range entry is (address..., shape_only) where shape_only says
    the reference sits inside a call that consumes references, not values.

    String literals and error literals are removed first: a reference that
    exists only inside a string is not a cell this formula reads (see the
    module docstring's "WHAT THIS CANNOT SEE").
    """
    text = _STRING.sub('""', formula or "")
    text = _ERRLIT.sub(" ", text)
    sheet_qualified = _SHEET_REF.findall(text)
    text = _SHEET_REF.sub(" ", text)
    spans = reference_shape_spans(text)
    singles, ranges = [], []
    for m in _REF.finditer(text):
        a, b = m.group(1), m.group(2)
        if a.replace("$", "").upper() in _NOT_A_REF:
            continue
        shape_only = any(lo <= m.start() < hi for lo, hi in spans)
        if b:
            ranges.append((normalize(a), normalize(b), shape_only))
        else:
            na = normalize(a)
            if na:
                singles.append((na, shape_only))
    return singles, ranges, sheet_qualified


def prose_allows_empty(prose, addr=None):
    """True when the case's own words say it is testing an empty cell.

    `prose` is the case id + description + expected_note. An emptiness word
    anywhere in it is enough: a case called SUM_empty_range whose description
    reads "an empty range" has told the reader exactly what it is doing. When
    `addr` is given, an address-specific sentence is also accepted, which is
    the stronger and more useful form for a note to use.
    """
    if not prose:
        return False
    if _EXPLICIT_EMPTY.search(prose) or _EMPTY_WORDS.search(prose):
        return True
    if addr:
        for sentence in re.split(r"(?<=[.;])\s+", prose):
            if re.search(r"(?<![A-Za-z0-9])\$?" + addr[0] + r"\$?" + addr[1:] + r"(?![A-Za-z0-9])",
                         sentence, re.I) and _EMPTY_WORDS_NEAR_ADDRESS.search(sentence):
                return True
    return False


def audit_case(fn, case):
    """Return (violations, warnings) for one case, each a printable string."""
    violations, warnings = [], []
    formula = case.get("formula", "")
    setup = {normalize(k): v for k, v in (case.get("setup_cells") or {}).items()
             if normalize(k)}
    cid = case.get("id", "<no id>")
    prose = " ".join(x for x in (cid.replace("_", " "), case.get("description"),
                                 case.get("expected_note")) if x)

    singles, ranges, sheet_refs = refs_in_formula(formula)

    for sref in sheet_refs:
        violations.append(
            f"{fn}/{cid}: sheet-qualified reference {sref!r} -- each case gets its own "
            f"sheet, so this can only resolve to a blank")

    for addr, shape_only in sorted(set(singles)):
        if addr in setup:
            continue
        if shape_only:
            warnings.append(
                f"{fn}/{cid}: {addr} is unset but is read as a reference, not a value")
            continue
        if prose_allows_empty(prose, addr):
            warnings.append(f"{fn}/{cid}: {addr} is unset and the case says so on purpose")
            continue
        violations.append(
            f"{fn}/{cid}: formula reads {addr}, which is not in setup_cells "
            f"(set it, or say in the note that {addr} is deliberately empty)")

    for a, b, shape_only in ranges:
        if not a or not b:
            continue
        cells = expand_range(a, b)
        label = f"{a}:{b}"
        if cells is None:
            warnings.append(f"{fn}/{cid}: range {label} is too large to expand; not checked")
            continue
        present = [c for c in cells if c in setup]
        if not present:
            if shape_only:
                warnings.append(
                    f"{fn}/{cid}: range {label} is unset but is read as a reference, "
                    f"not as values")
            elif prose_allows_empty(prose, a) or prose_allows_empty(prose, b):
                warnings.append(f"{fn}/{cid}: range {label} is entirely unset and the case says so")
            else:
                violations.append(
                    f"{fn}/{cid}: formula reads range {label}, and NOT ONE of its "
                    f"{len(cells)} cells is in setup_cells")
        elif len(present) != len(cells):
            missing = [c for c in cells if c not in setup]
            warnings.append(
                f"{fn}/{cid}: range {label} is partially populated "
                f"({len(present)}/{len(cells)} set; unset: {', '.join(missing[:8])}"
                f"{'...' if len(missing) > 8 else ''})")

    return violations, warnings


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("functions", nargs="*", help="function names to audit (default: all)")
    ap.add_argument("-v", "--verbose", action="store_true", help="show warnings too")
    ap.add_argument("--only-batch", action="store_true",
                    help="audit everything but fail only on the named functions "
                         "(pre-existing violations elsewhere are reported, not fatal)")
    args = ap.parse_args()

    names = set(args.functions) or None
    paths = sorted(glob.glob(os.path.join(TESTS_DIR, "*.json")))
    if not paths:
        print(f"No test files found under {TESTS_DIR}", file=sys.stderr)
        return 1

    scope = names if (names and not args.only_batch) else None
    fatal, other, warnings = [], [], []
    n_files = n_cases = 0
    for path in paths:
        fn = os.path.splitext(os.path.basename(path))[0]
        if scope is not None and fn not in scope:
            continue
        with open(path) as f:
            payload = json.load(f)
        n_files += 1
        for case in payload.get("cases", []):
            n_cases += 1
            v, w = audit_case(fn, case)
            warnings.extend(w)
            if names and fn not in names:
                other.extend(v)
            else:
                fatal.extend(v)

    print(f"check_test_setup_refs: {n_files} function file(s), {n_cases} case(s) audited")
    if args.verbose and warnings:
        print(f"\n{len(warnings)} warning(s) (not failures):")
        for w in warnings:
            print("  - " + w)
    elif warnings:
        print(f"({len(warnings)} warning(s) suppressed; re-run with -v)")

    if other:
        print(f"\n{len(other)} PRE-EXISTING violation(s) outside the named functions "
              f"(reported, not fatal):")
        for v in other:
            print("  ! " + v)

    if fatal:
        print(f"\n{len(fatal)} VIOLATION(S):")
        for v in fatal:
            print("  X " + v)
        return 1

    print("OK: every referenced cell is set up (or documented as deliberately empty).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
