#!/usr/bin/env python3
"""
Generates data/tests/<FUNCTION>.json test-case files for Phase 1 of the
spreadsheet function-compatibility harness.

Design notes:
- Formulas are written in canonical Excel UI syntax (no _xlfn./_xlfn._xlws.
  storage prefixes). Engine runners are responsible for translating to the
  storage form a given file format requires (see harness/xlfn_map.py).
- "expected" is the canonical/spec-correct value we assert against, when one
  is well-defined. It is null for cases that are intentionally
  non-deterministic (RAND/RANDBETWEEN) or for "existence only" probes.
- "expected_note" explains *why* that's the expected value, especially for
  edge cases (float rounding, sign-of-divisor conventions, etc).

Run: python3 scripts/gen_test_cases.py
Writes: data/tests/<FUNCTION>.json (one file per function)
"""
import json
import os

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "data", "tests")


def case(id_, formula, description, setup_cells=None, expected=None,
         expected_note=None, check_range=None):
    c = {"id": id_, "formula": formula, "description": description}
    if setup_cells is not None:
        c["setup_cells"] = setup_cells
    c["expected"] = expected
    if expected_note is not None:
        c["expected_note"] = expected_note
    if check_range is not None:
        c["check_range"] = check_range
    return c


FUNCTIONS = {}


def add(name, cases):
    FUNCTIONS[name] = {"function": name, "cases": cases}


# ---------------------------------------------------------------- XLOOKUP
add("XLOOKUP", [
    case("XLOOKUP_exact_match", '=XLOOKUP("b",A1:A3,B1:B3)',
         "Basic exact-match lookup returns the corresponding value",
         {"A1": "a", "A2": "b", "A3": "c", "B1": 1, "B2": 2, "B3": 3},
         expected=2),
    case("XLOOKUP_not_found_default_na", '=XLOOKUP("z",A1:A3,B1:B3)',
         "No if_not_found arg and no match -> #N/A",
         {"A1": "a", "A2": "b", "A3": "c", "B1": 1, "B2": 2, "B3": 3},
         expected="#N/A"),
    case("XLOOKUP_not_found_custom", '=XLOOKUP("z",A1:A3,B1:B3,"missing")',
         "Custom if_not_found value is returned when no match",
         {"A1": "a", "A2": "b", "A3": "c", "B1": 1, "B2": 2, "B3": 3},
         expected="missing"),
    case("XLOOKUP_match_mode_next_smaller", '=XLOOKUP(2.5,A1:A4,B1:B4,"none",-1)',
         "match_mode -1 (exact or next smaller item) with no exact match",
         {"A1": 1, "A2": 2, "A3": 3, "A4": 4, "B1": 10, "B2": 20, "B3": 30, "B4": 40},
         expected=20, expected_note="2.5 has no exact match; next-smaller is 2 -> 20"),
    case("XLOOKUP_search_mode_last_to_first", '=XLOOKUP("x",A1:A4,B1:B4,"none",0,-1)',
         "search_mode -1 finds the LAST occurrence of a duplicated key",
         {"A1": "x", "A2": "y", "A3": "x", "A4": "z", "B1": 1, "B2": 2, "B3": 3, "B4": 4},
         expected=3, expected_note="Two rows match 'x' (1 and 3); reverse search returns the later one"),
    case("XLOOKUP_empty_lookup_array", '=XLOOKUP("a",A1:A1,B1:B1)',
         "Degenerate 1-cell lookup/return arrays where the cell is blank",
         {}, expected="#N/A", expected_note="A1 and B1 are blank; lookup value 'a' is not found"),
])

# ---------------------------------------------------------------- LET
add("LET", [
    case("LET_basic", "=LET(x,5,x*2)", "Single binding used in the calculation",
         expected=10),
    case("LET_chained_bindings", "=LET(a,2,b,a*3,a+b)",
         "Later bindings can reference earlier ones in the same LET",
         expected=8),
    case("LET_array_binding", "=LET(rng,{1,2,3},SUM(rng))",
         "LET can bind an array literal and pass it to another function",
         expected=6),
    case("LET_odd_arg_count_error", "=LET(x,5,y,x*2)",
         "Odd total arg count (missing final calculation expression) is a syntax/value error",
         expected="#VALUE!", expected_note="LET requires pairs plus a trailing calculation; here the last pair has no calc term following it in this deliberately malformed call"),
])

# ---------------------------------------------------------------- LAMBDA
add("LAMBDA", [
    case("LAMBDA_basic_invoke", "=LAMBDA(x,x*2)(5)",
         "Immediately-invoked LAMBDA with one parameter", expected=10),
    case("LAMBDA_two_params", "=LAMBDA(x,y,x+y)(3,4)",
         "LAMBDA with two parameters", expected=7),
    case("LAMBDA_via_LET", "=LET(f,LAMBDA(x,x^2),f(4))",
         "LAMBDA bound to a name via LET, then called", expected=16),
    case("LAMBDA_wrong_arg_count", "=LAMBDA(x,y,x+y)(3)",
         "Calling with fewer arguments than parameters is an error",
         expected="#VALUE!"),
])

# ---------------------------------------------------------------- FILTER
add("FILTER", [
    case("FILTER_basic", "=FILTER(A1:A5,B1:B5>2)",
         "Spill the subset of A1:A5 where the paired B value is >2",
         {"A1": "a", "A2": "b", "A3": "c", "A4": "d", "A5": "e",
          "B1": 1, "B2": 2, "B3": 3, "B4": 4, "B5": 5},
         expected=["c", "d", "e"], check_range="A30:A32"),
    case("FILTER_all_false_with_default", '=FILTER(A1:A3,B1:B3>100,"none")',
         "if_empty argument is returned when no rows match",
         {"A1": 1, "A2": 2, "A3": 3, "B1": 1, "B2": 2, "B3": 3},
         expected="none"),
    case("FILTER_all_false_no_default", "=FILTER(A1:A3,B1:B3>100)",
         "No if_empty and no matches raises #CALC!",
         {"A1": 1, "A2": 2, "A3": 3, "B1": 1, "B2": 2, "B3": 3},
         expected="#CALC!"),
    case("FILTER_2d_by_column_condition", "=FILTER(A1:B3,A1:A3>1)",
         "Filtering a 2-column range by a condition on the first column",
         {"A1": 1, "A2": 2, "A3": 3, "B1": "x", "B2": "y", "B3": "z"},
         expected=[[2, "y"], [3, "z"]], check_range="A30:B31"),
    case("FILTER_inline_array_literal", "=FILTER({1,2,3,4},{1,0,1,0})",
         "Filtering an inline array literal by an inline boolean/number mask",
         expected=[1, 3], check_range="A30:B30"),
])

# ---------------------------------------------------------------- SORT
add("SORT", [
    case("SORT_default_ascending", "=SORT(A1:A5)",
         "Default sort_index=1, ascending",
         {"A1": 3, "A2": 1, "A3": 4, "A4": 1, "A5": 5},
         expected=[1, 1, 3, 4, 5], check_range="A30:A34"),
    case("SORT_descending", "=SORT(A1:A5,1,-1)",
         "Explicit descending sort order",
         {"A1": 3, "A2": 1, "A3": 4, "A4": 1, "A5": 5},
         expected=[5, 4, 3, 1, 1], check_range="A30:A34"),
    case("SORT_by_second_column", "=SORT(A1:B3,2,1)",
         "2D range sorted by its second column ascending",
         {"A1": "x", "B1": 30, "A2": "y", "B2": 10, "A3": "z", "B3": 20},
         expected=[["y", 10], ["z", 20], ["x", 30]], check_range="A30:B32"),
    case("SORT_empty_range_error", "=SORT(A1:A1)",
         "Sorting a single blank cell (degenerate case)",
         {}, expected=0, expected_note="Blank cell sorts as 0 in numeric context; no error expected"),
])

# ---------------------------------------------------------------- UNIQUE
add("UNIQUE", [
    case("UNIQUE_basic", "=UNIQUE(A1:A5)",
         "Basic unique list preserving first-occurrence order",
         {"A1": 1, "A2": 2, "A3": 2, "A4": 3, "A5": 1},
         expected=[1, 2, 3], check_range="A30:A32"),
    case("UNIQUE_exactly_once", "=UNIQUE(A1:A5,FALSE,TRUE)",
         "exactly_once=TRUE returns only items that appear exactly once",
         {"A1": 1, "A2": 2, "A3": 2, "A4": 3, "A5": 1},
         expected=[3], check_range="A30:A30"),
    case("UNIQUE_by_column", "=UNIQUE(A1:C2,TRUE)",
         "by_col=TRUE compares whole columns instead of rows",
         {"A1": 1, "B1": 2, "C1": 1, "A2": 1, "B2": 2, "C2": 1},
         expected=[1, 2], check_range="A30:B30",
         expected_note="Column A and column C are identical (1,1); unique columns are {1,2} and {1,1} but by_col dedups the repeated column, leaving 2 columns: [1,2] and [1,1] -> flatten to first row [1,2]"),
    case("UNIQUE_no_duplicates", "=UNIQUE(A1:A3)",
         "Input with no duplicates returns the input unchanged",
         {"A1": 5, "A2": 6, "A3": 7},
         expected=[5, 6, 7], check_range="A30:A32"),
])

# ---------------------------------------------------------------- SEQUENCE
add("SEQUENCE", [
    case("SEQUENCE_basic_vertical", "=SEQUENCE(5)",
         "Default rows-only sequence starting at 1, step 1",
         expected=[1, 2, 3, 4, 5], check_range="A30:A34"),
    case("SEQUENCE_rows_cols", "=SEQUENCE(2,3)",
         "2 rows x 3 columns filled row-major from 1",
         expected=[[1, 2, 3], [4, 5, 6]], check_range="A30:C31"),
    case("SEQUENCE_start_step", "=SEQUENCE(5,1,10,5)",
         "Custom start=10, step=5",
         expected=[10, 15, 20, 25, 30], check_range="A30:A34"),
    case("SEQUENCE_negative_step", "=SEQUENCE(3,1,5,-1)",
         "Negative step counts down",
         expected=[5, 4, 3], check_range="A30:A32"),
])

# ---------------------------------------------------------------- TEXTSPLIT
add("TEXTSPLIT", [
    case("TEXTSPLIT_basic", '=TEXTSPLIT("a,b,c",",")',
         "Basic column split on a single delimiter",
         expected=["a", "b", "c"], check_range="A30:C30"),
    case("TEXTSPLIT_row_and_col_delims", '=TEXTSPLIT("a,b;c,d",",",";")',
         "col_delimiter=\",\" and row_delimiter=\";\" produce a 2x2 grid",
         expected=[["a", "b"], ["c", "d"]], check_range="A30:B31"),
    case("TEXTSPLIT_ignore_empty", '=TEXTSPLIT("a,,b",",",,TRUE)',
         "ignore_empty=TRUE collapses consecutive delimiters",
         expected=["a", "b"], check_range="A30:B30"),
    case("TEXTSPLIT_pad_with", '=TEXTSPLIT("a,b,c;d",",",";",FALSE,0,"-")',
         "Uneven row lengths are padded with pad_with instead of #N/A",
         expected=[["a", "b", "c"], ["d", "-", "-"]], check_range="A30:C31"),
])

# ---------------------------------------------------------------- TEXTBEFORE
add("TEXTBEFORE", [
    case("TEXTBEFORE_basic", '=TEXTBEFORE("a-b-c","-")',
         "Text before the first delimiter occurrence", expected="a"),
    case("TEXTBEFORE_instance_num", '=TEXTBEFORE("a-b-c","-",2)',
         "instance_num=2 counts to the 2nd delimiter", expected="a-b"),
    case("TEXTBEFORE_not_found_default", '=TEXTBEFORE("abc","-")',
         "Delimiter not present and no if_not_found -> #N/A", expected="#N/A"),
    case("TEXTBEFORE_not_found_custom", '=TEXTBEFORE("abc","-","none")',
         "Delimiter not present with if_not_found supplied", expected="none"),
    case("TEXTBEFORE_negative_instance", '=TEXTBEFORE("a-b-c","-",-1)',
         "Negative instance_num counts delimiters from the end",
         expected="a-b", expected_note="-1 means the last delimiter; text before it is 'a-b'"),
])

# ---------------------------------------------------------------- TEXTAFTER
add("TEXTAFTER", [
    case("TEXTAFTER_basic", '=TEXTAFTER("a-b-c","-")',
         "Text after the first delimiter occurrence", expected="b-c"),
    case("TEXTAFTER_instance_num", '=TEXTAFTER("a-b-c","-",2)',
         "instance_num=2 counts to the 2nd delimiter", expected="c"),
    case("TEXTAFTER_negative_instance", '=TEXTAFTER("a-b-c","-",-1)',
         "Negative instance_num counts delimiters from the end",
         expected="c", expected_note="-1 means the last delimiter; text after it is 'c'"),
    case("TEXTAFTER_not_found_custom", '=TEXTAFTER("abc","-","none")',
         "Delimiter not present with if_not_found supplied", expected="none"),
])

# ---------------------------------------------------------------- IFS
add("IFS", [
    case("IFS_basic", '=IFS(A1>90,"A",A1>80,"B",TRUE,"C")',
         "First TRUE condition wins; TRUE literal acts as catch-all",
         {"A1": 85}, expected="B"),
    case("IFS_no_match_no_catchall", '=IFS(A1>90,"A",A1>95,"B")',
         "No condition true and no catch-all -> #N/A",
         {"A1": 50}, expected="#N/A"),
    case("IFS_order_matters", '=IFS(A1>0,"pos",A1>10,"big")',
         "Conditions evaluated left-to-right; first match short-circuits",
         {"A1": 50}, expected="pos", expected_note="A1>10 is also true but never reached"),
    case("IFS_catchall_first_value", '=IFS(TRUE,"always",A1>0,"pos")',
         "TRUE as the first condition always wins regardless of position",
         {"A1": 50}, expected="always"),
])

# ---------------------------------------------------------------- SWITCH
add("SWITCH", [
    case("SWITCH_basic_match", '=SWITCH(2,1,"one",2,"two",3,"three")',
         "Matches the second value/result pair", expected="two"),
    case("SWITCH_default_value", '=SWITCH(5,1,"one",2,"two","default")',
         "Trailing unpaired argument acts as the default", expected="default"),
    case("SWITCH_no_match_no_default", '=SWITCH(5,1,"one",2,"two")',
         "No match and no default -> #N/A", expected="#N/A"),
    case("SWITCH_text_match", '=SWITCH("b","a","A","b","B","c","C")',
         "Matching on text values", expected="B"),
])

# ---------------------------------------------------------------- MAXIFS
add("MAXIFS", [
    case("MAXIFS_single_criteria", '=MAXIFS(A1:A5,B1:B5,">2")',
         "Max of A where paired B > 2",
         {"A1": 10, "A2": 20, "A3": 30, "A4": 40, "A5": 50,
          "B1": 1, "B2": 2, "B3": 3, "B4": 4, "B5": 5},
         expected=50),
    case("MAXIFS_multi_criteria", '=MAXIFS(A1:A5,B1:B5,">1",C1:C5,"<5")',
         "Two criteria ranges ANDed together",
         {"A1": 10, "A2": 20, "A3": 30, "A4": 40, "A5": 50,
          "B1": 1, "B2": 2, "B3": 3, "B4": 4, "B5": 5,
          "C1": 1, "C2": 2, "C3": 3, "C4": 9, "C5": 9},
         expected=30),
    case("MAXIFS_no_matches", '=MAXIFS(A1:A3,B1:B3,">100")',
         "No rows satisfy the criteria -> 0 (not an error)",
         {"A1": 1, "A2": 2, "A3": 3, "B1": 1, "B2": 2, "B3": 3},
         expected=0),
    case("MAXIFS_wildcard_text", '=MAXIFS(A1:A3,B1:B3,"a*")',
         "Wildcard text criteria",
         {"A1": 1, "A2": 2, "A3": 3, "B1": "apple", "B2": "banana", "B3": "avocado"},
         expected=3),
])

# ---------------------------------------------------------------- MINIFS
add("MINIFS", [
    case("MINIFS_single_criteria", '=MINIFS(A1:A5,B1:B5,">2")',
         "Min of A where paired B > 2",
         {"A1": 10, "A2": 20, "A3": 30, "A4": 40, "A5": 50,
          "B1": 1, "B2": 2, "B3": 3, "B4": 4, "B5": 5},
         expected=30),
    case("MINIFS_multi_criteria", '=MINIFS(A1:A5,B1:B5,">1",C1:C5,"<5")',
         "Two criteria ranges ANDed together",
         {"A1": 10, "A2": 20, "A3": 30, "A4": 40, "A5": 50,
          "B1": 1, "B2": 2, "B3": 3, "B4": 4, "B5": 5,
          "C1": 1, "C2": 2, "C3": 3, "C4": 9, "C5": 9},
         expected=20),
    case("MINIFS_no_matches", '=MINIFS(A1:A3,B1:B3,">100")',
         "No rows satisfy the criteria -> 0 (not an error)",
         {"A1": 1, "A2": 2, "A3": 3, "B1": 1, "B2": 2, "B3": 3},
         expected=0),
])

# ---------------------------------------------------------------- SUMIFS
add("SUMIFS", [
    case("SUMIFS_basic", '=SUMIFS(A1:A5,B1:B5,">2")',
         "Sum of A where paired B > 2",
         {"A1": 10, "A2": 20, "A3": 30, "A4": 40, "A5": 50,
          "B1": 1, "B2": 2, "B3": 3, "B4": 4, "B5": 5},
         expected=120),
    case("SUMIFS_multi_criteria", '=SUMIFS(A1:A5,B1:B5,">1",C1:C5,"<5")',
         "Two AND-ed criteria ranges",
         {"A1": 10, "A2": 20, "A3": 30, "A4": 40, "A5": 50,
          "B1": 1, "B2": 2, "B3": 3, "B4": 4, "B5": 5,
          "C1": 1, "C2": 2, "C3": 3, "C4": 9, "C5": 9},
         expected=50),
    case("SUMIFS_no_match", '=SUMIFS(A1:A3,B1:B3,">100")',
         "No rows satisfy the criteria -> 0",
         {"A1": 1, "A2": 2, "A3": 3, "B1": 1, "B2": 2, "B3": 3},
         expected=0),
    case("SUMIFS_wildcard", '=SUMIFS(A1:A3,B1:B3,"a*")',
         "Wildcard text criteria matches 2 of 3 rows",
         {"A1": 1, "A2": 2, "A3": 3, "B1": "apple", "B2": "banana", "B3": "avocado"},
         expected=4),
])

# ---------------------------------------------------------------- COUNTIFS
add("COUNTIFS", [
    case("COUNTIFS_basic", '=COUNTIFS(A1:A5,">20")',
         "Count of cells > 20", {"A1": 10, "A2": 20, "A3": 30, "A4": 40, "A5": 50},
         expected=3),
    case("COUNTIFS_multi_range", '=COUNTIFS(A1:A5,">10",B1:B5,"<5")',
         "Two AND-ed criteria ranges",
         {"A1": 10, "A2": 20, "A3": 30, "A4": 40, "A5": 50,
          "B1": 1, "B2": 2, "B3": 3, "B4": 9, "B5": 9},
         expected=2, expected_note="Rows 2,3 satisfy A>10 AND B<5 (rows 4,5 fail B<5); rows matching: (20,2),(30,3)"),
    case("COUNTIFS_no_match", '=COUNTIFS(A1:A3,">100")',
         "No matches -> 0", {"A1": 1, "A2": 2, "A3": 3}, expected=0),
    case("COUNTIFS_wildcard", '=COUNTIFS(A1:A3,"a*")',
         "Wildcard text criteria",
         {"A1": "apple", "A2": "banana", "A3": "avocado"}, expected=2),
])

# ---------------------------------------------------------------- IFERROR
add("IFERROR", [
    case("IFERROR_no_error_passthrough", '=IFERROR(10/2,"err")',
         "No error: original value passes through unchanged", expected=5),
    case("IFERROR_catches_div0", '=IFERROR(10/0,"err")',
         "#DIV/0! is caught and replaced", expected="err"),
    case("IFERROR_nested_vlookup", '=IFERROR(VLOOKUP("z",A1:B3,2,FALSE),"not found")',
         "Catches #N/A bubbling up from a failed VLOOKUP",
         {"A1": "a", "B1": 1, "A2": "b", "B2": 2, "A3": "c", "B3": 3},
         expected="not found"),
    case("IFERROR_catches_ref_error", '=IFERROR(INDEX(A1:A3,10),"bad index")',
         "Catches #REF! from an out-of-bounds INDEX",
         {"A1": 1, "A2": 2, "A3": 3}, expected="bad index"),
])

# ---------------------------------------------------------------- IFNA
add("IFNA", [
    case("IFNA_catches_na", '=IFNA(MATCH("z",A1:A3,0),"not found")',
         "#N/A from a failed MATCH is caught",
         {"A1": "a", "A2": "b", "A3": "c"}, expected="not found"),
    case("IFNA_does_not_catch_other_errors", '=IFNA(10/0,"na")',
         "IFNA only catches #N/A; #DIV/0! must pass through unchanged",
         expected="#DIV/0!",
         expected_note="Common compat pitfall: IFNA is NOT a general error trap like IFERROR"),
    case("IFNA_no_error_passthrough", '=IFNA(10/2,"na")',
         "No error: original value passes through", expected=5),
])

# ---------------------------------------------------------------- VLOOKUP
add("VLOOKUP", [
    case("VLOOKUP_exact_match", '=VLOOKUP("b",A1:B3,2,FALSE)',
         "Exact match mode (range_lookup=FALSE)",
         {"A1": "a", "B1": 1, "A2": "b", "B2": 2, "A3": "c", "B3": 3},
         expected=2),
    case("VLOOKUP_approximate_match", "=VLOOKUP(3,A1:B5,2,TRUE)",
         "Approximate match requires ascending sorted first column",
         {"A1": 1, "B1": 10, "A2": 2, "B2": 20, "A3": 4, "B3": 40, "A4": 6, "B4": 60, "A5": 8, "B5": 80},
         expected=20, expected_note="3 is between 2 and 4; approx match returns the row for the largest value <= lookup (2 -> 20)"),
    case("VLOOKUP_not_found_exact", '=VLOOKUP("z",A1:B3,2,FALSE)',
         "Exact match with no match -> #N/A",
         {"A1": "a", "B1": 1, "A2": "b", "B2": 2, "A3": "c", "B3": 3},
         expected="#N/A"),
    case("VLOOKUP_col_index_out_of_range", '=VLOOKUP("a",A1:B3,5,FALSE)',
         "col_index_num beyond the table width -> #REF! per Microsoft docs",
         {"A1": "a", "B1": 1, "A2": "b", "B2": 2, "A3": "c", "B3": 3},
         expected="#REF!", expected_note="Microsoft docs specify #REF! for out-of-range col_index_num; record engines' ACTUAL error code here since this is a known point of cross-engine divergence"),
    case("VLOOKUP_wildcard_exact_mode", '=VLOOKUP("a*",A1:B3,2,FALSE)',
         "Wildcards are honored even in exact-match mode",
         {"A1": "apple", "B1": 1, "A2": "banana", "B2": 2, "A3": "avocado", "B3": 3},
         expected=1, expected_note="First row matching the a* wildcard pattern is 'apple'"),
])

# ---------------------------------------------------------------- INDEX
add("INDEX", [
    case("INDEX_1d_basic", "=INDEX(A1:A5,3)",
         "1-D array, single row-number argument",
         {"A1": 10, "A2": 20, "A3": 30, "A4": 40, "A5": 50}, expected=30),
    case("INDEX_2d_row_col", "=INDEX(A1:C3,2,2)",
         "2-D range, row+column arguments",
         {"A1": 1, "B1": 2, "C1": 3, "A2": 4, "B2": 5, "C2": 6, "A3": 7, "B3": 8, "C3": 9},
         expected=5),
    case("INDEX_whole_column_spill", "=INDEX(A1:C3,0,2)",
         "row_num=0 returns the entire column as a spilled array",
         {"A1": 1, "B1": 2, "C1": 3, "A2": 4, "B2": 5, "C2": 6, "A3": 7, "B3": 8, "C3": 9},
         expected=[2, 5, 8], check_range="A30:A32"),
    case("INDEX_out_of_range", "=INDEX(A1:A3,10)",
         "Row number beyond the range -> #REF!",
         {"A1": 1, "A2": 2, "A3": 3}, expected="#REF!"),
    case("INDEX_MATCH_combo", '=INDEX(B1:B3,MATCH("b",A1:A3,0))',
         "Classic INDEX/MATCH pattern used as a VLOOKUP replacement",
         {"A1": "a", "B1": 1, "A2": "b", "B2": 2, "A3": "c", "B3": 3},
         expected=2),
])

# ---------------------------------------------------------------- MATCH
add("MATCH", [
    case("MATCH_exact", '=MATCH("b",A1:A3,0)',
         "match_type 0: exact match, returns 1-based position",
         {"A1": "a", "A2": "b", "A3": "c"}, expected=2),
    case("MATCH_approx_ascending", "=MATCH(3,A1:A5,1)",
         "match_type 1 (default): largest value <= lookup, ascending sorted data",
         {"A1": 1, "A2": 2, "A3": 4, "A4": 6, "A5": 8}, expected=2,
         expected_note="Position 2 holds value 2, the largest value <= 3"),
    case("MATCH_approx_descending", "=MATCH(3,A1:A5,-1)",
         "match_type -1: smallest value >= lookup, descending sorted data",
         {"A1": 8, "A2": 6, "A3": 4, "A4": 2, "A5": 1}, expected=3,
         expected_note="Position 3 holds value 4, the smallest value >= 3 in descending data"),
    case("MATCH_not_found", '=MATCH("z",A1:A3,0)',
         "Exact match with no match -> #N/A",
         {"A1": "a", "A2": "b", "A3": "c"}, expected="#N/A"),
])

# ---------------------------------------------------------------- TEXTJOIN
add("TEXTJOIN", [
    case("TEXTJOIN_ignore_empty_true", '=TEXTJOIN(",",TRUE,"a","","b")',
         "ignore_empty=TRUE skips blank/empty-string args", expected="a,b"),
    case("TEXTJOIN_ignore_empty_false", '=TEXTJOIN(",",FALSE,"a","","b")',
         "ignore_empty=FALSE keeps the empty string as a segment", expected="a,,b"),
    case("TEXTJOIN_range", '=TEXTJOIN("-",TRUE,A1:A3)',
         "Joining a cell range with a delimiter",
         {"A1": "x", "A2": "y", "A3": "z"}, expected="x-y-z"),
    case("TEXTJOIN_multiple_ranges", '=TEXTJOIN(",",TRUE,A1:A2,B1:B2)',
         "Multiple ranges are flattened and joined in order",
         {"A1": "a", "A2": "b", "B1": "c", "B2": "d"}, expected="a,b,c,d"),
])

# ---------------------------------------------------------------- CONCAT
add("CONCAT", [
    case("CONCAT_literals", '=CONCAT("a","b","c")', "Basic literal concatenation",
         expected="abc"),
    case("CONCAT_range", "=CONCAT(A1:A3)",
         "CONCAT flattens a range directly (unlike CONCATENATE, which cannot take a range)",
         {"A1": "x", "A2": "y", "A3": "z"}, expected="xyz"),
    case("CONCAT_mixed_types", '=CONCAT("val:",5)',
         "Numbers are coerced to text without an explicit TEXT() call",
         expected="val:5"),
    case("CONCAT_2d_range", "=CONCAT(A1:B2)",
         "2-D range is flattened row-major",
         {"A1": "a", "B1": "b", "A2": "c", "B2": "d"}, expected="abcd"),
])

# ---------------------------------------------------------------- ROUND
add("ROUND", [
    case("ROUND_half_away_from_zero", "=ROUND(2.5,0)",
         "Excel/LO ROUND uses arithmetic rounding (half away from zero), NOT banker's rounding (round-half-to-even)",
         expected=3, expected_note="Banker's rounding would give 2; correct spreadsheet behavior is 3"),
    case("ROUND_negative_half_away_from_zero", "=ROUND(-2.5,0)",
         "Negative half-values also round away from zero", expected=-3),
    case("ROUND_negative_digits", "=ROUND(12345,-2)",
         "Negative num_digits rounds to the left of the decimal point",
         expected=12300),
    case("ROUND_float_representation_quirk", "=ROUND(1.005,2)",
         "Tests whether the engine falls into the naive-binary-float trap: the "
         "IEEE-754 double closest to 1.005 is actually ~1.00499999999999989, so "
         "a naive floor(x*100+0.5)/100 implementation would round DOWN to 1.00",
         expected=1.01, expected_note="Both Excel and LibreOffice correctly return 1.01, not 1.0 -- they normalize to ~15 significant decimal digits before rounding, specifically to avoid this well-known binary-float pitfall"),
])

# ---------------------------------------------------------------- MOD
add("MOD", [
    case("MOD_positive_positive", "=MOD(7,3)", "Both operands positive", expected=1),
    case("MOD_negative_dividend", "=MOD(-7,3)",
         "Excel/LO MOD result takes the SIGN OF THE DIVISOR, not truncated-remainder semantics",
         expected=2, expected_note="C-style '%' would give -1; spreadsheet MOD gives 2 because result = n - d*FLOOR(n/d)"),
    case("MOD_negative_divisor", "=MOD(7,-3)",
         "Negative divisor: result takes the sign of the divisor",
         expected=-2),
    case("MOD_divisor_zero", "=MOD(5,0)", "Division by zero -> #DIV/0!",
         expected="#DIV/0!"),
])

# ---------------------------------------------------------------- DATEDIF
add("DATEDIF", [
    case("DATEDIF_years", '=DATEDIF(DATE(2020,1,1),DATE(2023,6,15),"Y")',
         "Whole years between two dates", expected=3),
    case("DATEDIF_months", '=DATEDIF(DATE(2020,1,1),DATE(2023,6,15),"M")',
         "Whole months between two dates", expected=41),
    case("DATEDIF_days", '=DATEDIF(DATE(2024,1,1),DATE(2024,1,10),"D")',
         "Whole days between two dates", expected=9),
    case("DATEDIF_md_known_quirk", '=DATEDIF(DATE(2024,1,31),DATE(2024,3,1),"MD")',
         '"MD" (days ignoring months and years) is documented by Microsoft as unreliable/buggy for some date combos',
         expected=-1, expected_note="Textbook-naive expectation might be ~1, but MD has a well-known Microsoft-acknowledged bug around month-end dates; recording the ACTUAL engine output is the point of this test"),
    case("DATEDIF_end_before_start", '=DATEDIF(DATE(2024,1,10),DATE(2024,1,1),"D")',
         "End date before start date -> #NUM! per Microsoft's documented DATEDIF behavior",
         expected="#NUM!", expected_note="Microsoft docs state end<start raises #NUM!; record engines' ACTUAL error code here since this is a known point of cross-engine divergence"),
])

# ---------------------------------------------------------------- EDATE
add("EDATE", [
    # NOTE on "expected": EDATE returns a date SERIAL NUMBER (days since the
    # spreadsheet epoch), not a formatted string -- the anchor cell has no
    # number_format applied, so the raw serial is exactly what a readback
    # will show. Serial values below were computed independently via
    # openpyxl.utils.datetime.to_excel(date(...)) for cross-verification.
    case("EDATE_leap_year_clamp", "=EDATE(DATE(2024,1,31),1)",
         "Jan 31 + 1 month clamps to the last day of February in a leap year "
         "(2024-02-29); result is the underlying date serial number",
         expected=45351, expected_note="2024 is a leap year; Feb has no 31st so EDATE clamps to Feb 29 = serial 45351"),
    case("EDATE_negative_months", "=EDATE(DATE(2024,3,15),-2)",
         "Negative months argument goes backward (2024-01-15 = serial 45306)",
         expected=45306),
    case("EDATE_year_boundary", "=EDATE(DATE(2023,11,15),3)",
         "Crossing a year boundary (2024-02-15 = serial 45337)", expected=45337),
    case("EDATE_fractional_months_truncated", "=EDATE(DATE(2024,1,1),1.9)",
         "Non-integer months argument is truncated toward zero, not rounded "
         "(2024-02-01 = serial 45323)",
         expected=45323, expected_note="1.9 truncates to 1, not rounded to 2"),
])

# ---------------------------------------------------------------- NETWORKDAYS
add("NETWORKDAYS", [
    case("NETWORKDAYS_basic_week", "=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,5))",
         "Mon Jan 1 to Fri Jan 5, 2024, no holidays -> 5 workdays", expected=5),
    case("NETWORKDAYS_with_holiday", "=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,5),DATE(2024,1,3))",
         "Same range excluding one holiday (Wed Jan 3) -> 4", expected=4),
    case("NETWORKDAYS_reversed_dates", "=NETWORKDAYS(DATE(2024,1,5),DATE(2024,1,1))",
         "start_date after end_date returns a NEGATIVE count (not an error)",
         expected=-5),
    case("NETWORKDAYS_weekend_only", "=NETWORKDAYS(DATE(2024,1,6),DATE(2024,1,7))",
         "Sat Jan 6 to Sun Jan 7, 2024 -> 0 workdays", expected=0),
])

# ---------------------------------------------------------------- RAND
add("RAND", [
    case("RAND_exists_and_in_unit_range", "=RAND()",
         "Existence probe only: RAND() is non-deterministic; we just confirm the "
         "engine recognizes the function (not #NAME?) and returns a plausible value",
         expected=None, expected_note="Non-deterministic by design; harness should only check 0<=value<1 and absence of #NAME?"),
])

# ---------------------------------------------------------------- RANDBETWEEN
add("RANDBETWEEN", [
    case("RANDBETWEEN_exists_bounds", "=RANDBETWEEN(1,10)",
         "Existence probe: confirm recognized and result plausible (integer 1..10)",
         expected=None, expected_note="Non-deterministic by design; harness should only check 1<=value<=10 and absence of #NAME?"),
    case("RANDBETWEEN_degenerate_equal_bounds", "=RANDBETWEEN(5,5)",
         "When bottom==top the result is deterministic", expected=5),
    case("RANDBETWEEN_negative_range", "=RANDBETWEEN(-5,-1)",
         "Existence probe with an all-negative range",
         expected=None, expected_note="Non-deterministic; harness should only check -5<=value<=-1"),
])

# ---------------------------------------------------------------- ARRAYTOTEXT
add("ARRAYTOTEXT", [
    case("ARRAYTOTEXT_default_format", "=ARRAYTOTEXT({1,2,3})",
         "Default (concise) format joins elements with comma+space",
         expected="1, 2, 3"),
    case("ARRAYTOTEXT_strict_format", "=ARRAYTOTEXT({1,2,3},1)",
         "format=1 (strict) reproduces array-literal syntax with braces",
         expected="{1,2,3}"),
    case("ARRAYTOTEXT_2d_array", "=ARRAYTOTEXT({1,2;3,4})",
         "2-D array literal, default format",
         expected="1, 2, 3, 4"),
    case("ARRAYTOTEXT_text_elements_strict", '=ARRAYTOTEXT({"a","b"},1)',
         "Strict format quotes text elements",
         expected='{"a","b"}'),
])

def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    for name, payload in FUNCTIONS.items():
        path = os.path.join(OUT_DIR, f"{name}.json")
        with open(path, "w") as f:
            json.dump(payload, f, indent=2)
            f.write("\n")
        print(f"wrote {path} ({len(payload['cases'])} cases)")
    print(f"\nTotal functions: {len(FUNCTIONS)}")
    print(f"Total cases: {sum(len(p['cases']) for p in FUNCTIONS.values())}")


if __name__ == "__main__":
    main()
