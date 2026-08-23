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
 *           lnew  = earliest tested LO release where it works, or null.
 *
 * HONESTY CONTRACT (do not weaken when editing copy):
 *   - Excel and Google Sheets verdicts are documentation-based. Only
 *     LibreOffice verdicts with lv set are execution-based. The `basis` field
 *     says which, and the UI must surface it.
 *   - "quirky" means: the function IS recognized by LibreOffice, but at least
 *     one of our executed test cases returned a different value/error than
 *     Excel. It is not "broken" — it needs review, and we say exactly that.
 *   - Functions not in the dataset are UNKNOWN. We never guess.
 *   - This is function-level triage against a pre-computed dataset. It does
 *     not recalculate the user's workbook in the target app.
 */
(function () {
  'use strict';

  var APP_NAMES = { x: 'Excel', g: 'Google Sheets', l: 'LibreOffice Calc' };

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
   * classifyFunction(fnName, entry, target, source) -> {verdict, basis, note}
   *   verdict: 'ok' | 'quirk' | 'missing' | 'unknown'
   *   basis:   'executed' | 'documented' | null (null only for 'unknown')
   *   note:    one honest sentence the UI can show verbatim.
   *
   * For LibreOffice targets the EXECUTED verdict (lv) always outranks the
   * documentation flag (l) — same precedence the site's checker uses.
   */
  function classifyFunction(fnName, entry, target, source) {
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
      if (entry.lv === 'supported') {
        return {
          verdict: 'ok', basis: 'executed',
          note: 'Executed and verified in LibreOffice ' + entry.lver +
            (entry.lnew ? ' (works since ' + entry.lnew + ', the earliest release we tested it working in)' : '') + '.'
        };
      }
      if (entry.lv === 'quirky') {
        return {
          verdict: 'quirk', basis: 'executed',
          note: 'Recognized by LibreOffice ' + entry.lver + ', but at least one of our ' +
            'executed test cases returned a different value or error than Excel. It will ' +
            'not error out as unknown — review the affected cells for silent differences.'
        };
      }
      if (entry.lv === 'unsupported') {
        return {
          verdict: 'missing', basis: 'executed',
          note: 'Executed in LibreOffice ' + entry.lver + ' and not recognized — it ' +
            'returns #NAME?.' + srcNote
        };
      }
      // lv is null: no executed data; fall back to the documentation flag and
      // say so.
      if (entry.l) {
        return {
          verdict: 'ok', basis: 'documented',
          note: 'Documented for LibreOffice Calc, but not yet in our executed test set — ' +
            'this verdict is documentation-based, not execution-verified.'
        };
      }
      return {
        verdict: 'missing', basis: 'documented',
        note: 'Not documented for LibreOffice Calc (and not in our executed test set). ' +
          'Expect #NAME? after migration.' + srcNote
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
   * buildReport(auditResult, db, source, target) -> report
   *
   * auditResult: output of XlsxAudit.auditXlsx()
   * db:          parsed compat.json
   *
   * report = {
   *   source, target,
   *   functionRows: [{fn, count, verdict, basis, note, cat, cells:[{sheet,cell,formula}]}],
   *   formulas:     [{sheet, cell, formula, functions, verdict}],   // verdict = worst of its functions
   *   atRiskFunctions:  functionRows with verdict missing|quirk, sorted by compareAtRisk,
   *   unknownFunctions: functionRows with verdict unknown, sorted by count desc,
   *   totals: {formulas, sheets, uniqueFunctions, atRiskFormulas, unknownFunctions}
   * }
   */
  function buildReport(auditResult, db, source, target) {
    var byFn = {};
    var functionRows = Object.keys(auditResult.functionCounts).map(function (fn) {
      var entry = db[fn];
      var cls = classifyFunction(fn, entry, target, source);
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

  // One row per formula: sheet,cell,formula,functions,verdict
  function reportToCsv(report) {
    var lines = ['sheet,cell,formula,functions,verdict'];
    report.formulas.forEach(function (f) {
      lines.push([f.sheet, f.cell, '=' + f.formula, f.functions.join(' '), f.verdict]
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

  /* ===================== exports ===================== */

  var AuditVerdicts = {
    APP_NAMES: APP_NAMES,
    SEVERITY: SEVERITY,
    DIRECTIONS: DIRECTIONS,
    FREE_DETAIL_LIMIT: FREE_DETAIL_LIMIT,
    sourcePresence: sourcePresence,
    classifyFunction: classifyFunction,
    compareAtRisk: compareAtRisk,
    buildReport: buildReport,
    csvField: csvField,
    reportToCsv: reportToCsv,
    looksLikeLicenseKey: looksLikeLicenseKey
  };

  if (typeof globalThis !== 'undefined') globalThis.AuditVerdicts = AuditVerdicts;
  if (typeof module !== 'undefined' && module.exports) module.exports = AuditVerdicts;
})();
