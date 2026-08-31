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
    "STDEV.S", "STDEV.P", "VAR.S", "VAR.P",
    "NORM.DIST", "NORM.INV", "NORM.S.DIST", "NORM.S.INV",
    "MODE.SNGL", "MODE.MULT",
    "FORECAST.LINEAR", "FORECAST.ETS", "FORECAST.ETS.CONFINT",
    "FORECAST.ETS.SEASONALITY", "FORECAST.ETS.STAT",
    "AGGREGATE",
    # Excel 2013 additions
    "ARABIC", "BASE", "DECIMAL", "SHEET", "SHEETS", "ISFORMULA",
    "BITAND", "BITOR", "BITXOR", "BITLSHIFT", "BITRSHIFT",
    "COMBINA", "PERMUTATIONA", "PDURATION", "RRI",
    "RANK.EQ", "RANK.AVG",
    "PERCENTILE.INC", "PERCENTILE.EXC", "QUARTILE.INC",
    "PERCENTRANK.INC", "PERCENTRANK.EXC", "QUARTILE.EXC",
    "CEILING.PRECISE", "FLOOR.PRECISE", "ISO.CEILING",
    "COVARIANCE.P", "COVARIANCE.S",
    "UNICHAR", "UNICODE",
    "FORMULATEXT",
    "ISOWEEKNUM",
    "DAYS",
    "CEILING.MATH", "FLOOR.MATH",
    "XOR",  # Excel 2013 addition -- yes, even XOR needs the prefix
    # Excel 2010 statistical renames and Excel 2013 trigonometric additions
    # needed for the alphabetical corpus batch (ACCRINT..CHIDIST). ACOT/ACOTH
    # are Excel 2013 additions and are on XlsxWriter's future-function table
    # even though their legacy neighbours ACOS/ASIN/ATAN are not; the BETA./
    # BINOM./CHISQ./CONFIDENCE. dotted names are the 2010 renames of BETADIST/
    # BETAINV/BINOMDIST/CHIDIST/CHIINV/CHITEST/CONFIDENCE, whose UNDOTTED
    # legacy forms predate 2007 and correctly need no prefix.
    "ACOT", "ACOTH",
    "BETA.DIST", "BETA.INV",
    # Excel 2013 trigonometric additions from the alphabetical corpus batch
    # CHIINV..CUBEVALUE. COT/COTH/CSC/CSCH are the direct siblings of the
    # ACOT/ACOTH already listed above and are on XlsxWriter's future-function
    # table for the same reason. Verified empirically on all four pinned
    # builds (24.2.0.3, 24.8.7.2, 25.2.0.3, 25.8.7.3): "=_xlfn.COT(30)"
    # returns -0.156119952161659 while the unprefixed "=COT(30)" returns
    # #NAME?, and likewise for COTH/CSC/CSCH -- LO DOES implement all four,
    # so without these entries the harness would have recorded four
    # supported functions as unsupported.
    "COT", "COTH", "CSC", "CSCH",
    # COPILOT (Excel 2025, Frontier/Insider only) is post-2007 and therefore
    # takes the prefix by the same rule. Recorded here for storage-form
    # correctness only: the probe run found #NAME? on all four LibreOffice
    # builds under the plain name, the _xlfn. form AND the COM.MICROSOFT.
    # form, so LO has no such function under any spelling.
    "COPILOT",
    # Excel 2010 statistical renames and Excel 2013 web functions from the
    # alphabetical corpus batch DBCS..FISHERINV. All eight dotted names below
    # are on XlsxWriter's future-function table (they are the 2010 renames of
    # ERF/ERFC/EXPONDIST/FDIST/FINV/FTEST), and ENCODEURL/FILTERXML are the
    # Excel 2013 web additions on the same table. Verified empirically on all
    # four pinned builds: the unprefixed spelling returns #NAME? while the
    # _xlfn. form evaluates -- e.g. "=_xlfn.F.INV.RT(0.01,6,4)" returns
    # 15.2068648611575 and "=_xlfn.ENCODEURL(...)" returns a percent-encoded
    # string, while "=F.INV.RT(...)" and "=ENCODEURL(...)" are #NAME?. Their
    # UNDOTTED legacy neighbours ERF, ERFC, EXPONDIST, FDIST, FINV (and the
    # never-renamed FISHER, FISHERINV, DELTA, EUROCONVERT, FINDB and the whole
    # D-database family) predate 2007 and correctly need no prefix -- all of
    # them were probed plain and evaluate.
    "ERF.PRECISE", "ERFC.PRECISE",
    "EXPON.DIST",
    "F.DIST", "F.DIST.RT", "F.INV", "F.INV.RT", "F.TEST",
    "ENCODEURL", "FILTERXML",
    # DETECTLANGUAGE (Excel for Microsoft 365 / Excel Mobile only) is post-2007
    # and therefore takes the prefix by the same rule. Recorded here for
    # storage-form correctness only: the probe run found #NAME? on all four
    # LibreOffice builds under the plain name, the _xlfn. form AND the
    # COM.MICROSOFT. add-in form, so LO has no such function under any
    # spelling. Same treatment as COPILOT above.
    "DETECTLANGUAGE",
    # Excel 2010 statistical renames, Excel 2013 additions and one Excel 365
    # addition from the alphabetical corpus batch D (FLOOR.PRECISE..IMLN). All
    # eleven names below are on the MS XLSX-extensions future-function list
    # (mirrored by XlsxWriter): _xlfn.GAMMA, _xlfn.GAMMA.DIST, _xlfn.GAMMA.INV,
    # _xlfn.GAMMALN.PRECISE, _xlfn.GAUSS, _xlfn.HYPGEOM.DIST, _xlfn.IMAGE,
    # _xlfn.IMCOSH, _xlfn.IMCOT, _xlfn.IMCSC, _xlfn.IMCSCH. Verified empirically
    # on all four pinned builds: "=_xlfn.GAMMA(2.5)" returns 1.32934038817914
    # and "=_xlfn.HYPGEOM.DIST(1,4,8,20,TRUE)" returns 0.465428276573787 while
    # the unprefixed spellings are #NAME?, so seven of these would have been
    # recorded as unsupported without an entry here.
    #
    # THE FOUR IM* ENTRIES GO THE OTHER WAY, and are deliberately kept anyway.
    # For IMCOSH, IMCOT, IMCSC and IMCSCH the PLAIN name evaluates in
    # LibreOffice and the _xlfn. form -- the one real Excel actually writes into
    # a .xlsx -- returns #NAME? on every build, as do COM.MICROSOFT.IMCOT,
    # ORG.OPENOFFICE.IMCOT and _xlfn.ORG.OPENOFFICE.IMCOT (five spellings
    # probed, four dead). That is the exact mirror image of the COT/COTH/CSC/
    # CSCH case above. This module's contract is "the storage-form token a
    # function must use inside a raw .xlsx", i.e. what Excel writes -- not
    # "whichever spelling makes a given engine answer" -- so the spec-faithful
    # token is what is recorded, and the consequence (an Excel-authored
    # workbook using IMCOT opens in LibreOffice as #NAME?, even though
    # LibreOffice can compute IMCOT under the bare name) is a real interop gap
    # this corpus should report rather than paper over. It is documented case
    # by case in data/tests/IMCOSH.json, IMCOT.json, IMCSC.json and IMCSCH.json.
    # Their pre-2007 Analysis-ToolPak neighbours IMABS, IMAGINARY, IMARGUMENT,
    # IMCONJUGATE, IMCOS, IMDIV, IMEXP and IMLN are absent from the MS list and
    # were probed plain and evaluate, so they correctly take no prefix -- as do
    # FTEST, GAMMADIST, GAMMAINV, GAMMALN, GESTEP, GETPIVOTDATA and HYPGEOMDIST.
    "GAMMA", "GAMMA.DIST", "GAMMA.INV", "GAMMALN.PRECISE", "GAUSS",
    "HYPGEOM.DIST",
    "IMAGE", "IMCOSH", "IMCOT", "IMCSC", "IMCSCH",
    "BINOM.DIST", "BINOM.DIST.RANGE", "BINOM.INV",
    "CHISQ.DIST", "CHISQ.DIST.RT", "CHISQ.INV", "CHISQ.INV.RT", "CHISQ.TEST",
    "CONFIDENCE.NORM", "CONFIDENCE.T",
    # BAHTTEXT is NOT on XlsxWriter's future-function table and predates 2007,
    # but it is nevertheless stored prefixed: LibreOffice's own OOXML function
    # table (sc/source/filter/oox/formulabase.cxx) tags BAHTTEXT
    # FuncFlags::MACROCALL, i.e. "stored as macro call in Excel (_xlfn.
    # prefix)". Verified empirically on all four pinned builds (24.2.0.3,
    # 24.8.7.2, 25.2.0.3, 25.8.7.3): "=_xlfn.BAHTTEXT(1234)" returns the exact
    # Thai string Microsoft documents, while the unprefixed "=BAHTTEXT(1234)"
    # returns #NAME? on every one of them. Without this entry the harness would
    # have recorded BAHTTEXT as unsupported in LibreOffice, which is false.
    "BAHTTEXT",
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
