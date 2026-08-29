/*
 * AuditVerdicts — pure verdict engine for the canispreadsheet.com Migration
 * Audit. No DOM, no network, no state: every function here is a pure function
 * of its inputs, so the whole file is unit-testable in Node (see test.mjs).
 *
 * Inputs it works against:
 *   - the parse result of XlsxAudit.auditXlsx() (see audit.js, from the E1
 *     prototype — unchanged copy);
 *   - the site's verdict dataset docs/data/compat.json:
 *       { FUNC: { cat, x:bool, g:bool, l:bool, lv, lver, lnew } }
 *     where x/g   = documented in Excel / Google Sheets (documentation-based),
 *           l     = documented in LibreOffice Calc,
 *           lv    = EXECUTED LibreOffice verdict: "supported" | "quirky" |
 *                   "unsupported" | null (null = not in our executed set),
 *           lver  = LibreOffice version the executed verdict comes from,
 *           lnew  = earliest tested LO release where it works, or null
 *                   (null = it already worked in the oldest release we tested,
 *                   or it has no executed data at all).
 *
 * HONESTY CONTRACT (do not weaken when editing copy):
 *   - Excel and Google Sheets verdicts are documentation-based. Only
 *     LibreOffice verdicts with lv set are execution-based. The `basis` field
 *     says which, and the UI must surface it.
 *   - "quirky" means: the function IS recognized by LibreOffice, but at least
 *     one of our executed test cases returned a different value/error than
 *     Excel. It is not "broken" — it needs review, and we say exactly that.
 *   - Functions not in the dataset are UNKNOWN. We never guess.
 *   - Per-release LibreOffice verdicts come from `lnew` only. The dataset does
 *     NOT carry a per-version presence table, so we never claim "executed in
 *     <older release>" for a function whose lnew is null — we say it is
 *     supported in every release we tested, and name the range.
 *   - This is function-level triage against a pre-computed dataset. It does
 *     not recalculate the user's workbook in the target app.
 */
(function () {
  'use strict';

  var APP_NAMES = { x: 'Excel', g: 'Google Sheets', l: 'LibreOffice Calc' };

  /* ===================== LibreOffice target releases ===================== */

  // The LibreOffice builds we have actually executed the whole function set
  // in, oldest first. Every non-null `lnew` in compat.json is one of these.
  var LO_RELEASES = ['24.2.0.3', '24.8.7.2', '25.2.0.3', '25.8.7.3'];
  var LO_OLDEST = LO_RELEASES[0];
  var LO_LATEST = LO_RELEASES[LO_RELEASES.length - 1];

  // Short release series (what users actually say: "we're on 24.8") -> the
  // exact build we tested for that series. Distros ship a series, we test a
  // build; the UI labels both.
  var LO_SERIES = {
    '24.2': '24.2.0.3',
    '24.8': '24.8.7.2',
    '25.2': '25.2.0.3',
    '25.8': '25.8.7.3'
  };

  // Numeric major.minor.patch.build comparison. Missing segments count as 0
  // ('25.2' === '25.2.0.0'), non-numeric segments count as 0. Returns a
  // negative number, 0, or a positive number like every other comparator.
  function compareVersions(a, b) {
    var pa = String(a == null ? '' : a).split('.');
    var pb = String(b == null ? '' : b).split('.');
    var n = Math.max(pa.length, pb.length);
    for (var i = 0; i < n; i++) {
      var na = parseInt(pa[i], 10); if (isNaN(na)) na = 0;
      var nb = parseInt(pb[i], 10); if (isNaN(nb)) nb = 0;
      if (na !== nb) return na < nb ? -1 : 1;
    }
    return 0;
  }

  // Accepts a series ('25.2'), an exact tested build ('25.2.0.3'), or
  // null/garbage. Anything we did not test resolves to the LATEST tested
  // build, which is also the default: an unrecognized value can never
  // silently produce a harsher-than-tested verdict.
  function resolveTargetVersion(v) {
    if (!v) return LO_LATEST;
    var s = String(v).trim();
    if (LO_SERIES[s]) return LO_SERIES[s];
    for (var i = 0; i < LO_RELEASES.length; i++) {
      if (LO_RELEASES[i] === s) return s;
    }
    return LO_LATEST;
  }

  // "LibreOffice Calc 24.8.7.2" / "Google Sheets". Used by the summary line,
  // the CSV export and the print report.
  function targetLabel(target, targetVersion) {
    var name = APP_NAMES[target] || String(target);
    return target === 'l' ? name + ' ' + resolveTargetVersion(targetVersion) : name;
  }

  // Higher = worse. Used to pick a formula's overall verdict (worst function
  // wins) and to order the at-risk list.
  var SEVERITY = { missing: 3, quirk: 2, unknown: 1, ok: 0 };

  // How many at-risk functions get full per-cell detail in the free tier.
  var FREE_DETAIL_LIMIT = 3;

  // The 5 supported migration directions (source, target). Source changes
  // only the wording of "this exists in your current app" citations; the
  // verdict itself is purely a function of the TARGET app.
  var DIRECTIONS = [
    { id: 'x2g', source: 'x', target: 'g' },
    { id: 'x2l', source: 'x', target: 'l' },
    { id: 'g2x', source: 'g', target: 'x' },
    { id: 'g2l', source: 'g', target: 'l' },
    { id: 'l2x', source: 'l', target: 'x' }
  ];

  // How we know the function exists in the SOURCE app (for MISSING messages).
  // Returns a phrase or null if the dataset doesn't place it in the source.
  function sourcePresence(entry, source) {
    if (!entry) return null;
    if (source === 'l') {
      if (entry.lv === 'supported') return 'verified working in LibreOffice ' + entry.lver;
      if (entry.lv === 'quirky') return 'recognized (with quirks) in LibreOffice ' + entry.lver;
      if (entry.l) return 'documented for LibreOffice Calc';
      return null;
    }
    var flag = source === 'x' ? entry.x : entry.g;
    return flag ? 'documented in ' + APP_NAMES[source] : null;
  }

  /*
   * classifyFunction(fnName, entry, target, source, targetVersion)
   *   -> {verdict, basis, note}
   *   verdict: 'ok' | 'quirk' | 'missing' | 'unknown'
   *   basis:   'executed' | 'documented' | null (null only for 'unknown')
   *   note:    one honest sentence the UI can show verbatim.
   *
   * For LibreOffice targets the EXECUTED verdict (lv) always outranks the
   * documentation flag (l) — same precedence the site's checker uses.
   *
   * targetVersion (LibreOffice targets only; ignored otherwise) is the
   * LibreOffice release the user is migrating TO — a series ('24.8') or an
   * exact tested build ('24.8.7.2'). It defaults to the latest tested build,
   * for which the verdicts and their wording are byte-identical to what this
   * function returned before per-release targeting existed.
   *
   * PER-RELEASE RULE (the only version-dependent step):
   *   lv === 'unsupported'                     -> missing for EVERY target
   *   lnew set && targetVersion < lnew         -> missing (executed #NAME?)
   *   otherwise                                -> the lv verdict, unchanged
   * `lnew` is the earliest release we tested the function working in, so
   * "older than lnew" is an executed fact, not an inference.
   */
  function classifyFunction(fnName, entry, target, source, targetVersion) {
    // Not functions at all: Excel saves the dynamic-array spill operator (#)
    // as _xlfn.ANCHORARRAY(...) and the implicit-intersection operator (@)
    // as _xlfn.SINGLE(...) inside .xlsx files. Explain instead of reporting
    // them as unknown noise. We have not executed these operators, so the
    // verdict is honest at-risk-review, not a pass/fail claim.
    if (!entry && (fnName === 'ANCHORARRAY' || fnName === 'SINGLE')) {
      var opDesc = fnName === 'ANCHORARRAY'
        ? 'the spill operator # (e.g. =A1#)'
        : 'the implicit-intersection operator @ (e.g. =@A1:A10)';
      return {
        verdict: 'quirk',
        basis: 'documented',
        note: 'Not a real function: this is how Excel saves ' + opDesc + ' in the ' +
          '.xlsx file format. Behavior in ' + APP_NAMES[target] + ' depends on its ' +
          'dynamic-array support and version. We have not executed these operators ' +
          'yet, so review the affected cells manually.'
      };
    }
    if (!entry) {
      return {
        verdict: 'unknown',
        basis: null,
        note: 'Not in our dataset. This may be a newer function, an add-in or ' +
          'macro-defined function, or a named range called like a function. ' +
          'We do not guess: verify this one manually in ' + APP_NAMES[target] + '.'
      };
    }
    var src = sourcePresence(entry, source);
    var srcNote = src ? ' It is ' + src + '.' : '';

    if (target === 'l') {
      var tv = resolveTargetVersion(targetVersion);
      // The default target is the latest tested build; on that path every
      // note below is the exact string this engine produced before the
      // version selector existed (pinned by the existing fixtures).
      var latest = tv === LO_LATEST;
      var allTested = 'every LibreOffice release we tested (' + LO_OLDEST +
        ' \u2192 ' + LO_LATEST + ')';

      // Not recognized in the build we execute against => not recognized in
      // any older one either. Version-independent.
      if (entry.lv === 'unsupported') {
        return {
          verdict: 'missing', basis: 'executed',
          note: 'Executed in LibreOffice ' + entry.lver + ' and not recognized — it ' +
            'returns #NAME?.' + (latest ? '' : ' It is not recognized in ' + allTested +
            ', so LibreOffice ' + tv + ' will not run it either.') + srcNote
        };
      }

      // Per-release gate. lnew = the earliest release we tested it working
      // in; below that we executed it and got #NAME?.
      if (entry.lnew && compareVersions(tv, entry.lnew) < 0) {
        return {
          verdict: 'missing', basis: 'executed',
          note: 'Returns #NAME? in LibreOffice ' + tv + ' (executed) — it works since ' +
            entry.lnew + ', the earliest release we tested it working in. Upgrading ' +
            'LibreOffice to ' + entry.lnew + ' or later is the fix; otherwise replace it ' +
            'with a documented alternative.' + srcNote
        };
      }

      if (entry.lv === 'supported') {
        return {
          verdict: 'ok', basis: 'executed',
          note: latest
            ? 'Executed and verified in LibreOffice ' + entry.lver +
              (entry.lnew ? ' (works since ' + entry.lnew + ', the earliest release we tested it working in)' : '') + '.'
            : (entry.lnew
              ? 'Executed and verified in LibreOffice ' + entry.lnew + ' and every later ' +
                'release we tested, including your target ' + tv + '.'
              : 'Supported in ' + allTested + ', including your target ' + tv + '.')
        };
      }
      if (entry.lv === 'quirky') {
        return {
          verdict: 'quirk', basis: 'executed',
          note: 'Recognized by LibreOffice ' + entry.lver + ', but at least one of our ' +
            'executed test cases returned a different value or error than Excel. It will ' +
            'not error out as unknown — review the affected cells for silent differences.' +
            (latest ? '' : ' The quirk was measured in ' + entry.lver + '; the function is ' +
              'recognized in ' + allTested + ', but we have not re-measured the quirk in ' +
              tv + '.')
        };
      }
      // lv is null: no executed data; fall back to the documentation flag and
      // say so. Nothing here varies by release — we have no per-version data
      // for these functions and will not invent any.
      var noVersionData = latest ? '' : ' We have no executed per-release data for it, ' +
        'so this verdict does not change for your ' + tv + ' target.';
      if (entry.l) {
        return {
          verdict: 'ok', basis: 'documented',
          note: 'Documented for LibreOffice Calc, but not yet in our executed test set — ' +
            'this verdict is documentation-based, not execution-verified.' + noVersionData
        };
      }
      return {
        verdict: 'missing', basis: 'documented',
        note: 'Not documented for LibreOffice Calc (and not in our executed test set). ' +
          'Expect #NAME? after migration.' + srcNote + noVersionData
      };
    }

    // Excel / Google Sheets targets: documentation-based only, and we say so.
    var present = target === 'x' ? entry.x : entry.g;
    if (present) {
      return {
        verdict: 'ok', basis: 'documented',
        note: 'In ' + APP_NAMES[target] + '’s official function reference. ' +
          '(Excel/Sheets verdicts are documentation-based; we execution-verify LibreOffice.)'
      };
    }
    return {
      verdict: 'missing', basis: 'documented',
      note: 'Not in ' + APP_NAMES[target] + '’s documented function set — formulas ' +
        'using it typically fail with #NAME? after migration.' + srcNote
    };
  }

  // Ordering of the at-risk list, which also decides which functions get free
  // per-cell detail (the first FREE_DETAIL_LIMIT entries).
  // Severity class first (missing before quirk), then usage count, then name:
  // for LibreOffice targets ubiquitous quirk-flagged functions (SUM, VLOOKUP)
  // would otherwise out-count every genuine breaker and the free tier would
  // never show a single hard failure. To rank purely by usage count instead,
  // drop the first comparison.
  function compareAtRisk(a, b) {
    return (SEVERITY[b.verdict] - SEVERITY[a.verdict]) ||
      (b.count - a.count) ||
      (a.fn < b.fn ? -1 : a.fn > b.fn ? 1 : 0);
  }

  /*
   * buildReport(auditResult, db, source, target, targetVersion) -> report
   *
   * auditResult:   output of XlsxAudit.auditXlsx()
   * db:            parsed compat.json
   * targetVersion: LibreOffice release being migrated to (series or exact
   *                tested build); ignored for Excel/Sheets targets, defaults
   *                to the latest tested build.
   *
   * report = {
   *   source, target,
   *   targetVersion,   // resolved build string for LibreOffice targets, else null
   *   targetLabel,     // "LibreOffice Calc 24.8.7.2" / "Google Sheets"
   *   functionRows: [{fn, count, verdict, basis, note, cat, cells:[{sheet,cell,formula}]}],
   *   formulas:     [{sheet, cell, formula, functions, verdict}],   // verdict = worst of its functions
   *   atRiskFunctions:  functionRows with verdict missing|quirk, sorted by compareAtRisk,
   *   unknownFunctions: functionRows with verdict unknown, sorted by count desc,
   *   totals: {formulas, sheets, uniqueFunctions, atRiskFormulas, unknownFunctions}
   * }
   */
  function buildReport(auditResult, db, source, target, targetVersion) {
    var tv = target === 'l' ? resolveTargetVersion(targetVersion) : null;
    var byFn = {};
    var functionRows = Object.keys(auditResult.functionCounts).map(function (fn) {
      var entry = db[fn];
      var cls = classifyFunction(fn, entry, target, source, tv);
      var row = {
        fn: fn,
        count: auditResult.functionCounts[fn],
        verdict: cls.verdict,
        basis: cls.basis,
        note: cls.note,
        cat: entry ? entry.cat : null,
        cells: []
      };
      byFn[fn] = row;
      return row;
    });

    var formulas = auditResult.formulas.map(function (f) {
      var worst = 'ok';
      f.functions.forEach(function (fn) {
        var row = byFn[fn];
        row.cells.push({ sheet: f.sheet, cell: f.cell, formula: f.formula });
        if (SEVERITY[row.verdict] > SEVERITY[worst]) worst = row.verdict;
      });
      return {
        sheet: f.sheet, cell: f.cell, formula: f.formula,
        functions: f.functions, verdict: worst
      };
    });

    var atRisk = functionRows
      .filter(function (r) { return r.verdict === 'missing' || r.verdict === 'quirk'; })
      .sort(compareAtRisk);
    var unknown = functionRows
      .filter(function (r) { return r.verdict === 'unknown'; })
      .sort(function (a, b) {
        return (b.count - a.count) || (a.fn < b.fn ? -1 : a.fn > b.fn ? 1 : 0);
      });

    var atRiskFormulas = 0;
    formulas.forEach(function (f) {
      if (f.verdict === 'missing' || f.verdict === 'quirk') atRiskFormulas++;
    });

    return {
      source: source,
      target: target,
      targetVersion: tv,
      targetLabel: targetLabel(target, tv),
      functionRows: functionRows,
      formulas: formulas,
      atRiskFunctions: atRisk,
      unknownFunctions: unknown,
      totals: {
        formulas: auditResult.totals.formulas,
        sheets: auditResult.totals.sheets,
        uniqueFunctions: auditResult.totals.uniqueFunctions,
        atRiskFormulas: atRiskFormulas,
        unknownFunctions: unknown.length
      }
    };
  }

  /* ===================== CSV export (paid tier) ===================== */

  function csvField(v) {
    var s = String(v);
    return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  }

  // One row per formula: sheet,cell,formula,functions,verdict,target
  // The `target` column repeats the migration target on every row (including
  // the LibreOffice release the verdicts were computed for) so an exported
  // file is self-describing without a non-CSV preamble line.
  function reportToCsv(report) {
    var tgt = report.targetLabel || targetLabel(report.target, report.targetVersion);
    var lines = ['sheet,cell,formula,functions,verdict,target'];
    report.formulas.forEach(function (f) {
      lines.push([f.sheet, f.cell, '=' + f.formula, f.functions.join(' '), f.verdict, tgt]
        .map(csvField).join(','));
    });
    return lines.join('\r\n') + '\r\n';
  }

  /* ===================== license key format ===================== */

  // Gumroad license keys look like XXXXXXXX-XXXXXXXX-XXXXXXXX-XXXXXXXX (hex).
  // This is a FORMAT check only — it proves nothing about a purchase. It is
  // used as the offline fallback when the Gumroad verify API is unreachable
  // (the API allows cross-origin calls — access-control-allow-origin: * as of
  // 2026-08-23 — so the online check in audit-app.js is the primary path).
  var LICENSE_KEY_RE = /^[0-9A-Fa-f]{8}-[0-9A-Fa-f]{8}-[0-9A-Fa-f]{8}-[0-9A-Fa-f]{8}$/;

  function looksLikeLicenseKey(key) {
    return LICENSE_KEY_RE.test(String(key || '').trim());
  }

  /* ===================== guide links ===================== */

  // Map a function name -> the documented cross-engine divergence guides for
  // it, given a parsed docs/data/guides.json ({ FUNC: [{slug, title}] }).
  // Case-insensitive; missing function or missing/falsy guides map -> [].
  // Pure, no DOM/network — audit-app.js does the (silent-fail) fetch.
  function guidesForFunction(fnName, guides) {
    if (!guides) return [];
    var key = String(fnName || '').toUpperCase();
    return guides[key] || [];
  }

  /* ===================== exports ===================== */

  var AuditVerdicts = {
    APP_NAMES: APP_NAMES,
    SEVERITY: SEVERITY,
    DIRECTIONS: DIRECTIONS,
    FREE_DETAIL_LIMIT: FREE_DETAIL_LIMIT,
    LO_RELEASES: LO_RELEASES,
    LO_SERIES: LO_SERIES,
    LO_OLDEST: LO_OLDEST,
    LO_LATEST: LO_LATEST,
    compareVersions: compareVersions,
    resolveTargetVersion: resolveTargetVersion,
    targetLabel: targetLabel,
    sourcePresence: sourcePresence,
    classifyFunction: classifyFunction,
    compareAtRisk: compareAtRisk,
    buildReport: buildReport,
    csvField: csvField,
    reportToCsv: reportToCsv,
    looksLikeLicenseKey: looksLikeLicenseKey,
    guidesForFunction: guidesForFunction
  };

  if (typeof globalThis !== 'undefined') globalThis.AuditVerdicts = AuditVerdicts;
  if (typeof module !== 'undefined' && module.exports) module.exports = AuditVerdicts;
})();
