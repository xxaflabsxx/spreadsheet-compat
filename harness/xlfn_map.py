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
    "CEILING.PRECISE", "FLOOR.PRECISE",
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
    # Excel 2010 renames and Excel 2013 additions from the alphabetical corpus
    # batch E (IMLOG10..MMULT). All six are on the MS XLSX-extensions
    # future-function list, [MS-XLSX] section 2.2.3 "Functions", which spells
    # them _xlfn.LOGNORM.DIST, _xlfn.LOGNORM.INV, _xlfn.IMSEC, _xlfn.IMSECH,
    # _xlfn.IMSINH and _xlfn.IMTAN. Verified empirically on all four pinned
    # builds: "=_xlfn.LOGNORM.DIST(4,3.5,1.2,TRUE)" returns 0.0390835557068005
    # and "=_xlfn.LOGNORM.INV(0.039084,3.5,1.2)" returns 4.00002521868064 while
    # the unprefixed spellings are #NAME?, so those two would have been recorded
    # as unsupported without an entry here.
    #
    # THE FOUR IM* ENTRIES GO THE OTHER WAY, exactly as batch D's IMCOSH/IMCOT/
    # IMCSC/IMCSCH do: the PLAIN name evaluates in LibreOffice and the _xlfn.
    # form -- the one real Excel writes into a .xlsx -- is #NAME? on every
    # build, as are COM.MICROSOFT.IMSEC, ORG.OPENOFFICE.IMSEC and
    # _xlfn.ORG.OPENOFFICE.IMSEC (five spellings probed, four dead). This
    # module records the storage-form token Excel writes, not whichever
    # spelling makes a given engine answer, so the spec-faithful token stays
    # and the resulting interop gap is reported case by case in
    # data/tests/IMSEC.json, IMSECH.json, IMSINH.json and IMTAN.json.
    "LOGNORM.DIST", "LOGNORM.INV",
    "IMSEC", "IMSECH", "IMSINH", "IMTAN",
    # ISO.CEILING WAS REMOVED FROM THIS SET BY BATCH E -- it was wrong, and it
    # would have published a false "unsupported" verdict. ISO.CEILING IS a
    # future function (it is absent from ISO/IEC 29500's predefined list), but
    # it is one of exactly FOUR future functions that MS-XLSX section 2.2.3
    # spells WITHOUT the prefix: ECMA.CEILING, ISO.CEILING, NETWORKDAYS.INTL
    # and WORKDAY.INTL, each listed there as a bare name while its 155
    # neighbours carry "_xlfn.". LibreOffice's own OOXML filter draws the same
    # line: sc/source/filter/oox/formulabase.cxx tags ISO.CEILING with plain
    # FuncFlags::MACROCALL ("stored as macro call in BIFF Excel"), i.e. the
    # prefix is added for the old binary format only, whereas CEILING.PRECISE,
    # LOGNORM.DIST/INV and IMSEC/IMSECH/IMSINH/IMTAN are tagged
    # MACROCALL_NEW (= MACROCALL | MACROCALL_FN), which is what adds the
    # prefix to the OOXML name as well. Confirmed empirically both ways: on
    # all four pinned builds "=ISO.CEILING(4.3)" returns 5 while
    # "=_xlfn.ISO.CEILING(4.3)" and "=COM.MICROSOFT.ISO.CEILING(4.3)" are
    # #NAME?, and round-tripping a workbook through LibreOffice's own xlsx
    # EXPORT writes the formula back out as the bare token ISO.CEILING.
    # (XlsxWriter's table, this module's usual mirror, simply drops the
    # ISO.CEILING row -- it keeps the other three unprefixed names at their
    # correct alphabetical slots -- which is how the wrong entry got here.
    # The same author's Perl and C ports both list it unprefixed.)
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
    # Excel 2010/2013 renames and additions, plus five Excel 365 functions, from
    # the alphabetical corpus batch F (MUNIT..RIGHTB). Every one of them was
    # probed empirically on all four pinned builds (24.2.0.3, 24.8.7.2,
    # 25.2.0.3, 25.8.7.3) in four spellings -- plain, _xlfn., COM.MICROSOFT.
    # and ORG.OPENOFFICE. -- before being written down here, and the four builds
    # agreed in every case.
    #
    # THE FIVE THAT LIBREOFFICE ANSWERS TO ONLY WHEN PREFIXED. Without an entry
    # here each of these would have been recorded as unsupported in LibreOffice,
    # which is false: "=_xlfn.MUNIT(2)" spills the unit matrix while "=MUNIT(2)"
    # is #NAME?, and likewise _xlfn.NEGBINOM.DIST(10,5,0.25,FALSE) returns
    # 0.0550486603751779, _xlfn.PHI(0.75) returns 0.301137432154804 and
    # _xlfn.POISSON.DIST(2,5,TRUE) returns 0.124652019483081, each against
    # #NAME? for the bare name. Their UNDOTTED compatibility neighbours go the
    # other way and correctly take no prefix: NEGBINOMDIST, NORMDIST, NORMINV,
    # NORMSDIST, NORMSINV and POISSON were all probed plain and evaluate.
    # (PERCENTRANK.EXC, PERCENTRANK.INC and PERMUTATIONA are in this batch too
    # and behave the same way, but they were already in the set above.)
    "MUNIT", "NEGBINOM.DIST", "PHI", "POISSON.DIST",
    # THE FOUR THAT LIBREOFFICE HAS UNDER NO SPELLING AT ALL. PERCENTOF (the
    # GROUPBY/PIVOTBY aggregation helper) and the REGEXEXTRACT / REGEXREPLACE /
    # REGEXTEST trio are Excel 365 additions and therefore take the prefix by
    # the same post-2007 rule; the probe found #NAME? on all four builds under
    # the plain name, the _xlfn. form, the COM.MICROSOFT. form AND the
    # ORG.OPENOFFICE. form, so the entry cannot change any verdict -- it is
    # recorded for storage-form correctness only, exactly as COPILOT and
    # DETECTLANGUAGE above are. LibreOffice does have regular expressions in a
    # worksheet function, but under a different name and a different flavour:
    # its own REGEX(Text; Expression [; [Replacement] [; Flags|Occurrence]])
    # documents ICU regular expressions, where Microsoft's three pages all say
    # "All regular expressions for this function ... use the PCRE2 'flavor' of
    # regex".
    "PERCENTOF", "REGEXEXTRACT", "REGEXREPLACE", "REGEXTEST",
    # Excel 2010 statistical renames, Excel 2013 additions and four Excel 365
    # functions, from the alphabetical corpus batch G (RTD..ZTEST) -- the batch
    # that closes the Excel-documented set. Every name below was probed
    # empirically on all four pinned builds (24.2.0.3, 24.8.7.2, 25.2.0.3,
    # 25.8.7.3) in FIVE spellings -- plain, _xlfn., COM.MICROSOFT.,
    # ORG.OPENOFFICE. and _xlfn.ORG.OPENOFFICE. -- before being written down,
    # and the four builds agreed on every one.
    #
    # THE ELEVEN THAT LIBREOFFICE ANSWERS TO ONLY WHEN PREFIXED. Without an
    # entry here each of these would have been published as unsupported in
    # LibreOffice, which is false -- the largest single verdict-changing block
    # this push has added. The bare name is #NAME? on every build while the
    # _xlfn. form evaluates: _xlfn.SEC(1) = 1.85081571768093, _xlfn.SECH(1) =
    # 0.648054273663886, _xlfn.SKEW.P(1,2,3,4,10) = 1.13841995766062,
    # _xlfn.T.DIST(60,1,TRUE) = 0.994695326367377, _xlfn.T.DIST.2T(1.96,60) =
    # 0.0546449297365292, _xlfn.T.DIST.RT(1.96,60) = 0.0273224648682646,
    # _xlfn.T.INV(0.75,2) = 0.816496580927726, _xlfn.T.INV.2T(0.546449,60) =
    # 0.606533075825754, _xlfn.T.TEST(...) = 0.196015784925282,
    # _xlfn.WEIBULL.DIST(105,20,100,TRUE) = 0.929581390069277 and
    # _xlfn.Z.TEST(...) = 0.0905741968513638. Their UNDOTTED legacy neighbours
    # go the other way and correctly take no prefix: SIN, SINH, TAN, TANH,
    # SQRTPI, SERIESSUM, SEARCHB, STEYX, SUMX2MY2, SUMX2PY2, SUMXMY2, TDIST,
    # TINV, TTEST, WEIBULL, ZTEST and the whole TBILL*/YIELD* family were all
    # probed plain and evaluate.
    "SEC", "SECH", "SKEW.P",
    "T.DIST", "T.DIST.2T", "T.DIST.RT", "T.INV", "T.INV.2T", "T.TEST",
    "WEIBULL.DIST", "Z.TEST",
    # THREE THAT LIBREOFFICE HAS UNDER NO SPELLING AT ALL, recorded for
    # storage-form correctness only, exactly as COPILOT, DETECTLANGUAGE and the
    # REGEX* trio above are: all five spellings are #NAME? on all four builds,
    # so these entries cannot change a verdict. (VALUETOTEXT is in the same
    # position and was already in the set above.)
    "STOCKHISTORY", "TRANSLATE", "TRIMRANGE",
    # WEBSERVICE IS DIFFERENT FROM ALL OF THE ABOVE AND THE DIFFERENCE IS WHY
    # BATCH G DECLINED TO EXECUTE IT. It is an Excel 2013 web function like its
    # neighbours ENCODEURL and FILTERXML, so it takes the prefix by the same
    # post-2007 rule, and there is NO data/tests/WEBSERVICE.json -- this entry
    # is unused by any test case. The probe is the reason it is recorded here
    # anyway: on all four builds "=_xlfn.WEBSERVICE(...)" does NOT return
    # #NAME?. It PARSES AND EVALUATES, returning #N/A for an unreachable host,
    # while the plain, COM.MICROSOFT., ORG.OPENOFFICE. and
    # _xlfn.ORG.OPENOFFICE. spellings are all #NAME?. So LibreOffice does
    # implement WEBSERVICE, under the prefixed spelling, and this module would
    # be wrong to omit it. What the corpus will not do is turn that #N/A into a
    # published verdict: WEBSERVICE fetches a URL, its documented behaviour IS
    # the network round trip, and the #N/A this harness sees describes the
    # sandbox's DNS rather than the engine -- classify_verdict() would read it
    # as "quirky", i.e. as a LibreOffice defect, which would be false. See the
    # batch G commit message and site/audit-page/make_fixtures.py.
    "WEBSERVICE",
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
