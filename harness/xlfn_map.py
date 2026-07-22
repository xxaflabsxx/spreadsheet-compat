"""
Maps modern Excel function names to the storage-form token they must use
inside a raw .xlsx (OOXML) file.

WHY THIS EXISTS
----------------
The OOXML spreadsheet spec was frozen at the Excel 2007 function set. Any
function added in Excel 2010 or later ("future functions") is NOT stored in
a formula's XML as its plain name -- Excel silently prefixes it with
"_xlfn." (or "_xlfn._xlws." for a couple of dynamic-array functions) when it
serializes the file, and silently strips the prefix back off when it
displays the formula bar. Excel does this translation transparently, but
libraries that write raw OOXML XML (openpyxl, xlsxwriter, ...) do NOT do
this for you -- if you write "=XLOOKUP(...)" verbatim with openpyxl, Excel
*and* LibreOffice will both fail to recognize the function and show
#NAME?, even on engines that actually implement XLOOKUP. This is a
well-known openpyxl/xlsxwriter gotcha, not a real compatibility gap.

We verified this empirically against LibreOffice 24.2.7.2: e.g. writing
"=_xlfn.IFS(...)" or "=_xlfn.TEXTJOIN(...)" round-trips and evaluates
correctly, while the unprefixed form returns #NAME? even though LO DOES
support IFS/TEXTJOIN. Conversely, "=_xlfn.LET(...)" and
"=_xlfn.XLOOKUP(...)" (correct prefix) STILL return #NAME? in LO 24.2,
which we independently confirmed is a genuine support gap (not a prefix
bug) by driving LibreOffice's own native formula parser over UNO/PyUNO,
which also fails to recognize LET/XLOOKUP/FILTER/SORT/UNIQUE/SEQUENCE/
LAMBDA/TEXTBEFORE/TEXTAFTER/TEXTSPLIT/ARRAYTOTEXT as function names at all
(they get silently lower-cased and treated as unknown identifiers).

Test-case JSON files always store the NATURAL Excel-UI formula text (no
prefix) -- translation happens here, once, in the engine runner, so the
test corpus stays engine-agnostic and human-readable.

Source for the exact _xlfn. vs _xlfn._xlws. split: XlsxWriter's documented
"Working with Formulas" future-function table (the de facto public
reference for this OOXML quirk): only FILTER and SORT need the special
"_xlfn._xlws." double prefix; all other post-2007 functions use plain
"_xlfn.". A few Analysis-ToolPak-era functions absorbed into the Excel 2007
core (NETWORKDAYS, NETWORKDAYS.INTL, WORKDAY, WORKDAY.INTL, EDATE,
RANDBETWEEN, DATEDIF, etc.) need NO prefix at all, since they predate/are
part of the frozen OOXML spec.
"""

# Functions needing the double "_xlfn._xlws." prefix (dynamic-array
# functions that collided with pre-existing internal namespace names).
_XLWS_FUNCTIONS = {
    "FILTER",
    "SORT",
}

# Functions needing the plain "_xlfn." prefix (all other post-2007
# "future functions"). This list covers the Phase-1 function set plus a
# handful of common neighbors; extend as more functions are added.
_XLFN_FUNCTIONS = {
    "XLOOKUP", "XMATCH",
    "LET", "LAMBDA",
    # Lambda-helper functions (Excel 2022) -- all stored with the _xlfn. prefix.
    "MAP", "REDUCE", "SCAN", "BYROW", "BYCOL", "MAKEARRAY",
    "GROUPBY", "PIVOTBY",
    "UNIQUE", "SEQUENCE", "SORTBY", "RANDARRAY",
    "TEXTSPLIT", "TEXTBEFORE", "TEXTAFTER",
    "ARRAYTOTEXT", "VALUETOTEXT", "VSTACK", "HSTACK",
    "TOROW", "TOCOL", "WRAPROWS", "WRAPCOLS", "TAKE", "DROP",
    "EXPAND", "CHOOSEROWS", "CHOOSECOLS", "ISOMITTED",
    "IFS", "SWITCH", "MAXIFS", "MINIFS", "TEXTJOIN", "CONCAT",
    "IFNA", "NUMBERVALUE",
    # Excel 2010+ statistical ".INC"/".EQ"/".AVG"/".S"/".P" renames, plus a
    # handful of Excel 2013 additions, needed for the Phase-2 test batch.
    # Source: XlsxWriter's "Working with Formulas" future-function table.
    "STDEV.S", "STDEV.P", "VAR.S",
    "MODE.SNGL", "MODE.MULT",
    "RANK.EQ", "RANK.AVG",
    "PERCENTILE.INC", "PERCENTILE.EXC", "QUARTILE.INC",
    "UNICHAR", "UNICODE",
    "FORMULATEXT",
    "ISOWEEKNUM",
    "DAYS",
    "CEILING.MATH", "FLOOR.MATH",
    "XOR",  # Excel 2013 addition -- yes, even XOR needs the prefix
}


def to_storage_formula(formula: str, function_name: str) -> str:
    """
    Given a formula string as a human would type it in the Excel UI
    (e.g. '=XLOOKUP(1,A1:A3,B1:B3)') and the primary function name under
    test, return the formula with the correct _xlfn./_xlfn._xlws. prefix
    applied to that function's call sites, ready to write into an .xlsx
    via openpyxl.

    This only prefixes occurrences of `function_name` itself (word-boundary
    match, case-insensitive), not incidental other future-functions that
    might appear in setup formulas -- callers needing multiple prefixed
    functions in one formula should call this once per distinct function
    name that needs translating.
    """
    import re

    name = function_name.upper()
    if name in _XLWS_FUNCTIONS:
        prefix = "_xlfn._xlws."
    elif name in _XLFN_FUNCTIONS:
        prefix = "_xlfn."
    else:
        return formula  # no translation needed (legacy / pre-2007 function)

    pattern = re.compile(r"(?<![A-Za-z0-9_.])" + re.escape(name) + r"(?=\()", re.IGNORECASE)
    return pattern.sub(prefix + name, formula)


def storage_function_names():
    """Return the set of all function names known to need a storage prefix."""
    return set(_XLFN_FUNCTIONS) | set(_XLWS_FUNCTIONS)


def to_storage_formula_all(formula: str) -> str:
    """
    Like to_storage_formula, but prefixes EVERY known future-function call
    site in the formula, not just one named function. This matches what
    real Excel does when it serializes a formula: every post-2007 function
    in the expression gets its prefix, including nested calls (e.g.
    '=UNICHAR(UNICODE("x"))' must become
    '=_xlfn.UNICHAR(_xlfn.UNICODE("x"))' -- prefixing only the outer
    UNICHAR leaves the inner UNICODE unrecognized and the whole formula
    evaluates to #NAME? in every engine, which would be a harness artifact
    masquerading as an unsupported-function finding).

    Names are applied longest-first, and the word-boundary lookbehind
    (which rejects a preceding '.') prevents re-prefixing an
    already-prefixed call or matching a shorter name inside a longer
    dotted one (e.g. plain CEILING never matches 'CEILING.MATH(' because
    the lookahead requires an immediate '(').
    """
    import re

    out = formula
    for name in sorted(storage_function_names(), key=len, reverse=True):
        prefix = "_xlfn._xlws." if name in _XLWS_FUNCTIONS else "_xlfn."
        pattern = re.compile(
            r"(?<![A-Za-z0-9_.])" + re.escape(name) + r"(?=\()", re.IGNORECASE)
        out = pattern.sub(prefix + name, out)
    return out
