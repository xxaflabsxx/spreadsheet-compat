// Test suite for the Migration Audit page (E2): verdict engine unit tests +
// end-to-end fixture test through the real E1 parser.
// Usage: node test.mjs   (Node 18+; reuses E1's deflate-raw shim approach)
import fs from 'node:fs';
import path from 'node:path';
import url from 'node:url';
import { createRequire } from 'node:module';

const here = path.dirname(url.fileURLToPath(import.meta.url));
const require = createRequire(import.meta.url);

// Node 18 has DecompressionStream but not the 'deflate-raw' format (browsers
// and Node >= 20.12 / 21.2 do). Same zlib-backed shim as the E1 suite.
let usedShim = false;
try {
  new DecompressionStream('deflate-raw');
} catch {
  const zlib = await import('node:zlib');
  const stream = await import('node:stream');
  const Native = globalThis.DecompressionStream;
  globalThis.DecompressionStream = class DecompressionStream {
    constructor(format) {
      if (format !== 'deflate-raw') return new Native(format);
      const web = stream.Duplex.toWeb(zlib.createInflateRaw());
      this.readable = web.readable;
      this.writable = web.writable;
    }
  };
  usedShim = true;
}

const X = require(path.join(here, 'audit.js'));
const V = require(path.join(here, 'audit-verdicts.js'));
const DB = JSON.parse(fs.readFileSync(
  path.join(here, '..', '..', 'docs', 'data', 'compat.json'), 'utf8'));

let pass = 0, fail = 0;
function ok(cond, label) {
  if (cond) { pass++; console.log('  ok  ' + label); }
  else { fail++; console.log('  FAIL ' + label); }
}
function eq(actual, expected, label) {
  const good = JSON.stringify(actual) === JSON.stringify(expected);
  if (!good) console.log('       expected: ' + JSON.stringify(expected) +
    '\n       actual:   ' + JSON.stringify(actual));
  ok(good, label);
}

/* ---------- guard: the bundled parser is the tested E1 parser, unchanged ---------- */
console.log('guard: audit.js matches the E1 prototype byte-for-byte');
{
  const ours = fs.readFileSync(path.join(here, 'audit.js'));
  const e1 = fs.readFileSync(path.join(here, '..', 'audit-prototype', 'audit.js'));
  ok(ours.equals(e1), 'identical copy (re-copy from audit-prototype/ if this fails)');
}

/* ---------- guard: the doc-only fixture slot is still doc-only ---------- */
// verdict-mix.xlsx uses one function purely to exercise the DOCUMENTED-ONLY
// verdict basis (Excel-documented, LibreOffice-documented, never executed).
// The corpus executes functions alphabetically, so that function eventually
// acquires an lv and the four doc-only assertions below start failing for a
// reason that has nothing to do with the verdict engine. This guard names the
// cause up front: pick the next still-unexecuted x && l && !g function sorting
// between AGGREGATE and TEXTSPLIT, update this file and make_fixtures.py, and
// regenerate verdict-mix.xlsx.
console.log('guard: doc-only fixture function');
{
  const e = DB.WEBSERVICE;
  ok(e && e.x === true && e.g === false && e.l === true && e.lv === null,
    'WEBSERVICE is still Excel+LibreOffice documented and never executed ' +
    '(if this fails, re-point the doc-only fixture slot -- see make_fixtures.py)');
}

/* ---------- unit: classifyFunction — every branch, synthetic entries ---------- */
console.log('unit: classifyFunction branches');
{
  const c = V.classifyFunction;
  // unknown: not in dataset
  let r = c('WHATEVER', undefined, 'g', 'x');
  eq([r.verdict, r.basis], ['unknown', null], 'absent entry -> unknown, no basis');
  ok(/do not guess/i.test(r.note), 'unknown note says we do not guess');

  // Excel/Sheets targets are documentation-based
  // gv absent = no executed Sheets data for this function: documentation only.
  r = c('F', { x: true, g: true, l: false, gv: null, lv: null }, 'g', 'x');
  eq([r.verdict, r.basis], ['ok', 'documented'], 'documented in Sheets, not executed -> ok/documented');
  ok(/documentation-based/.test(r.note), 'un-executed Sheets ok note admits documentation basis');

  r = c('F', { x: true, g: false, l: false, gv: null, lv: null }, 'g', 'x');
  eq([r.verdict, r.basis], ['missing', 'documented'], 'g:false, no executed data -> missing/documented');
  ok(/documented in Excel/.test(r.note), 'missing note cites source-app presence (Excel)');

  // gv set = EXECUTED Google Sheets verdict, and it outranks the g flag exactly
  // as lv outranks l. Sheets has no version, so the note names the run DATE.
  const GVER = 'Google Sheets (Drive import, 2026-08-29)';
  r = c('F', { x: true, g: true, l: false, gv: 'supported', gver: GVER }, 'g', 'x');
  eq([r.verdict, r.basis], ['ok', 'executed'], 'gv supported -> ok/executed');
  ok(r.note.includes(GVER), 'executed Sheets ok note names the dated run');

  r = c('F', { x: true, g: true, l: false, gv: 'quirky', gver: GVER }, 'g', 'x');
  eq([r.verdict, r.basis], ['quirk', 'executed'], 'gv quirky -> quirk/executed');

  r = c('F', { x: true, g: true, l: false, gv: 'unsupported', gver: GVER }, 'g', 'x');
  eq([r.verdict, r.basis], ['missing', 'executed'], 'gv unsupported -> missing/executed even with g:true');
  ok(/#NAME\?/.test(r.note), 'executed Sheets missing note names the error it returned');

  // "inconclusive" is not a verdict: the .xlsx round trip, not Sheets, explains
  // the result, so we fall back to documentation and say so.
  r = c('F', { x: true, g: true, l: false, gv: 'inconclusive', gver: GVER }, 'g', 'x');
  eq([r.verdict, r.basis], ['ok', 'documented'], 'gv inconclusive + g:true -> ok/documented fallback');
  ok(/inconclusive/.test(r.note), 'inconclusive note says so');
  r = c('F', { x: true, g: false, l: false, gv: 'inconclusive', gver: GVER }, 'g', 'x');
  eq([r.verdict, r.basis], ['unknown', 'documented'], 'gv inconclusive + g:false -> unknown');

  // Sheets never takes a version: there is one dated run to compare against.
  eq(c('F', { x: true, g: true, gv: 'supported', gver: GVER }, 'g', 'x', '24.2').note,
     c('F', { x: true, g: true, gv: 'supported', gver: GVER }, 'g', 'x').note,
     'Sheets notes ignore targetVersion entirely');

  // Executed Sheets source presence, cited when the TARGET is missing it.
  r = c('F', { x: false, g: true, l: false, gv: 'supported', gver: GVER }, 'x', 'g');
  ok(/executed and verified in Google Sheets on /.test(r.note),
    'Sheets-source missing note cites the executed source presence');

  r = c('F', { x: false, g: true, l: false, lv: null }, 'x', 'g');
  eq([r.verdict, r.basis], ['missing', 'documented'], 'x:false -> missing for Excel target');
  ok(/documented in Google Sheets/.test(r.note), 'missing note cites source-app presence (Sheets)');

  // LibreOffice target: executed verdict outranks the documentation flag
  r = c('F', { x: true, g: true, l: false, lv: 'supported', lver: '25.8.7.3', lnew: '24.8.7.2' }, 'l', 'x');
  eq([r.verdict, r.basis], ['ok', 'executed'], 'lv supported -> ok/executed even with l:false');
  ok(/25\.8\.7\.3/.test(r.note) && /24\.8\.7\.2/.test(r.note), 'note cites lver and lnew');

  r = c('F', { x: true, g: true, l: true, lv: 'quirky', lver: '25.8.7.3' }, 'l', 'x');
  eq([r.verdict, r.basis], ['quirk', 'executed'], 'lv quirky -> quirk/executed');
  ok(/different value or error/.test(r.note), 'quirk note explains what quirky means');
  ok(!/broken/i.test(r.note), 'quirk note does not overclaim breakage');

  r = c('F', { x: true, g: false, l: false, lv: 'unsupported', lver: '25.8.7.3' }, 'l', 'x');
  eq([r.verdict, r.basis], ['missing', 'executed'], 'lv unsupported -> missing/executed');
  ok(/#NAME\?/.test(r.note), 'unsupported note cites #NAME?');

  r = c('F', { x: true, g: false, l: true, lv: null }, 'l', 'x');
  eq([r.verdict, r.basis], ['ok', 'documented'], 'l:true without executed data -> ok/documented');
  ok(/not yet in our executed test set/.test(r.note), 'doc-only LO note admits no execution');

  r = c('F', { x: false, g: true, l: false, lv: null }, 'l', 'g');
  eq([r.verdict, r.basis], ['missing', 'documented'], 'l:false, no executed data -> missing/documented');

  // source citation from an executed-LibreOffice source
  r = c('F', { x: false, g: false, l: true, lv: 'supported', lver: '25.8.7.3' }, 'x', 'l');
  ok(/verified working in LibreOffice 25\.8\.7\.3/.test(r.note),
    'LO-source missing note cites executed source presence');
}

/* ---------- unit: classifyFunction against the real dataset ---------- */
console.log('unit: classifyFunction on real compat.json entries');
{
  const c = V.classifyFunction;
  eq(c('AGGREGATE', DB.AGGREGATE, 'g', 'x').verdict, 'missing', 'AGGREGATE missing in Sheets');
  eq(c('AGGREGATE', DB.AGGREGATE, 'g', 'x').basis, 'executed', 'AGGREGATE Sheets basis is executed now');
  eq(c('TEXTSPLIT', DB.TEXTSPLIT, 'g', 'x').verdict, 'missing', 'TEXTSPLIT executed #NAME? in Sheets');
  eq(c('TEXTSPLIT', DB.TEXTSPLIT, 'g', 'x').basis, 'executed', 'TEXTSPLIT Sheets basis is executed');
  eq(c('MAP', DB.MAP, 'g', 'x'), {
    verdict: 'ok', basis: 'executed', note: c('MAP', DB.MAP, 'g', 'x').note
  }, 'MAP executes fine in Sheets even though LO returns #NAME?');
  eq(c('CONCAT', DB.CONCAT, 'g', 'x').verdict, 'quirk', 'CONCAT is an executed Sheets quirk (#N/A)');
  eq(c('CONCAT', DB.CONCAT, 'g', 'x').basis, 'executed', 'CONCAT Sheets quirk is execution-based');
  eq(c('AGGREGATE', DB.AGGREGATE, 'l', 'x').verdict, 'ok', 'AGGREGATE ok in LO (executed)');
  eq(c('AGGREGATE', DB.AGGREGATE, 'l', 'x').basis, 'executed', 'AGGREGATE LO basis executed');
  eq(c('GROUPBY', DB.GROUPBY, 'l', 'x').verdict, 'missing', 'GROUPBY missing in LO (executed #NAME?)');
  eq(c('FILTER', DB.FILTER, 'l', 'x').verdict, 'quirk', 'FILTER quirk in LO');
  // FILTER's Sheets run is quirky (a plain-name re-run confirmed it works but
  // returns #N/A instead of Excel's documented #CALC! on empty results, and
  // ignores if_empty), so it is now an executed quirk, not a doc fallback.
  eq(DB.FILTER.gv, 'quirky', 'FILTER Sheets verdict is quirky (executed, plain-name run)');
  eq(c('FILTER', DB.FILTER, 'g', 'x').verdict, 'quirk', 'FILTER quirk in Sheets (executed)');
  eq(c('FILTER', DB.FILTER, 'g', 'x').basis, 'executed', 'FILTER Sheets basis is executed, not documented');
  eq(c('SUM', DB.SUM, 'l', 'x').verdict, 'quirk', 'even SUM is an executed quirk in LO');
  eq(c('SUM', DB.SUM, 'g', 'x').verdict, 'ok', 'SUM ok in Sheets');
  eq(c('SUM', DB.SUM, 'g', 'x').basis, 'executed', 'SUM Sheets verdict is execution-based');
  eq(c('GOOGLEFINANCE', DB.GOOGLEFINANCE, 'x', 'g').verdict, 'missing', 'GOOGLEFINANCE missing in Excel');
  eq(c('WEBSERVICE', DB.WEBSERVICE, 'l', 'x'), {
    verdict: 'ok', basis: 'documented',
    note: c('WEBSERVICE', DB.WEBSERVICE, 'l', 'x').note
  }, 'WEBSERVICE LO doc-only -> ok/documented');
  eq(c('NOTAREALFUNCTION', DB.NOTAREALFUNCTION, 'g', 'x').verdict, 'unknown', 'absent fn unknown');
}

/* ---------- unit: at-risk ordering ---------- */
console.log('unit: compareAtRisk ordering');
{
  const rows = [
    { fn: 'QUIRKBIG', verdict: 'quirk', count: 99 },
    { fn: 'MISS2', verdict: 'missing', count: 2 },
    { fn: 'MISSB', verdict: 'missing', count: 1 },
    { fn: 'MISSA', verdict: 'missing', count: 1 },
    { fn: 'QUIRKSMALL', verdict: 'quirk', count: 1 }
  ].sort(V.compareAtRisk);
  eq(rows.map(r => r.fn), ['MISS2', 'MISSA', 'MISSB', 'QUIRKBIG', 'QUIRKSMALL'],
    'missing before quirk, then count desc, then name asc');
}

/* ---------- unit: CSV export ---------- */
console.log('unit: reportToCsv');
{
  eq(V.csvField('plain'), 'plain', 'csvField plain');
  eq(V.csvField('a,b'), '"a,b"', 'csvField comma quoted');
  eq(V.csvField('say "hi"'), '"say ""hi"""', 'csvField embedded quotes doubled');
  const audit = {
    functionCounts: { IF: 1 },
    formulas: [{ sheet: 'S,1', cell: 'A1', formula: 'IF(A1>2,"a,b","c")', functions: ['IF'] }],
    totals: { formulas: 1, sheets: 1, uniqueFunctions: 1 }
  };
  const rep = V.buildReport(audit, DB, 'x', 'g');
  const csv = V.reportToCsv(rep);
  const lines = csv.trim().split('\r\n');
  eq(lines[0], 'sheet,cell,formula,functions,verdict,target', 'CSV header');
  eq(lines[1], '"S,1",A1,"=IF(A1>2,""a,b"",""c"")",IF,ok,Google Sheets',
    'CSV row quoted correctly, target column stated');
  // LibreOffice targets carry the exact tested build in the target column, so
  // an exported CSV says which release the verdicts were computed for.
  const csvLo = V.reportToCsv(V.buildReport(audit, DB, 'x', 'l', '24.8'));
  eq(csvLo.trim().split('\r\n')[1], '"S,1",A1,"=IF(A1>2,""a,b"",""c"")",IF,ok,LibreOffice Calc 24.8.7.2',
    'CSV target column names the LibreOffice build');
}

/* ---------- unit: license key format ---------- */
console.log('unit: looksLikeLicenseKey');
{
  ok(V.looksLikeLicenseKey('ABCD1234-0F9E8D7C-12345678-DEADBEEF'), 'valid uppercase hex key');
  ok(V.looksLikeLicenseKey('abcd1234-0f9e8d7c-12345678-deadbeef'), 'lowercase accepted');
  ok(V.looksLikeLicenseKey('  ABCD1234-0F9E8D7C-12345678-DEADBEEF '), 'surrounding whitespace trimmed');
  ok(!V.looksLikeLicenseKey('ABCD1234-0F9E8D7C-12345678-DEADBEE'), 'short last group rejected');
  ok(!V.looksLikeLicenseKey('GHIJ1234-0F9E8D7C-12345678-DEADBEEF'), 'non-hex letters rejected');
  ok(!V.looksLikeLicenseKey('ABCD12340F9E8D7C12345678DEADBEEF'), 'missing dashes rejected');
  ok(!V.looksLikeLicenseKey(''), 'empty rejected');
  ok(!V.looksLikeLicenseKey(null), 'null rejected');
}

console.log('unit: guidesForFunction (function -> divergence-guide links)');
{
  const guides = {
    XLOOKUP: [{ slug: 'xlookup-vs-vlookup', title: 'XLOOKUP vs VLOOKUP' }],
    UNIQUE: [
      { slug: 'unique-a', title: 'UNIQUE quirk A' },
      { slug: 'unique-b', title: 'UNIQUE quirk B' },
    ],
  };
  eq(V.guidesForFunction('XLOOKUP', guides), guides.XLOOKUP, 'exact-case match returns its guides');
  eq(V.guidesForFunction('xlookup', guides), guides.XLOOKUP, 'lowercase function name is case-insensitive');
  eq(V.guidesForFunction('XlOoKuP', guides), guides.XLOOKUP, 'mixed-case function name is case-insensitive');
  eq(V.guidesForFunction('UNIQUE', guides), guides.UNIQUE, 'function with multiple guides returns all of them');
  eq(V.guidesForFunction('SUM', guides), [], 'function absent from the guides map returns []');
  eq(V.guidesForFunction('XLOOKUP', {}), [], 'empty guides map returns []');
  eq(V.guidesForFunction('XLOOKUP', null), [], 'null guides map returns [] (silent-fail-safe)');
  eq(V.guidesForFunction('XLOOKUP', undefined), [], 'undefined guides map returns []');
  eq(V.guidesForFunction('', guides), [], 'empty function name returns []');
}

/* ---------- end-to-end: verdict-mix.xlsx through parser + engine ---------- */
console.log('e2e: verdict-mix.xlsx');
const buf = fs.readFileSync(path.join(here, 'verdict-mix.xlsx'));
const audit = await X.auditXlsx(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength));
{
  eq(audit.sheets.map(s => s.name), ['Calc', 'Mix'], 'sheet names');
  eq(audit.totals.formulas, 12, 'total formulas');
  eq(audit.totals.uniqueFunctions, 11, 'unique functions');
  // note: bare SUM as GROUPBY's aggregation argument (no parenthesis) is
  // correctly NOT tokenized as a call — SUM counts only C1 and Mix!A4.
  for (const [fn, n] of [['SUM', 2], ['VLOOKUP', 1], ['AGGREGATE', 1], ['GROUPBY', 2],
    ['FILTER', 1], ['TEXTSPLIT', 1], ['WEBSERVICE', 1], ['GOOGLEFINANCE', 1],
    ['ARRAYFORMULA', 1], ['NOTAREALFUNCTION', 1], ['IF', 1]]) {
    eq(audit.functionCounts[fn], n, 'functionCounts.' + fn);
  }
}

function fnVerdicts(rep) {
  const out = {};
  rep.functionRows.forEach(r => { out[r.fn] = r.verdict; });
  return out;
}
function formulaVerdict(rep, sheet, cell) {
  const f = rep.formulas.find(f => f.sheet === sheet && f.cell === cell);
  return f && f.verdict;
}

console.log('e2e: Excel -> Google Sheets (default direction)');
{
  const rep = V.buildReport(audit, DB, 'x', 'g');
  eq(rep.totals, { formulas: 12, sheets: 2, uniqueFunctions: 11, atRiskFormulas: 6, unknownFunctions: 1 },
    'summary tiles');
  eq(fnVerdicts(rep), {
    SUM: 'ok', VLOOKUP: 'ok', AGGREGATE: 'missing', GROUPBY: 'missing', FILTER: 'quirk',
    TEXTSPLIT: 'missing', WEBSERVICE: 'missing', GOOGLEFINANCE: 'ok', ARRAYFORMULA: 'ok',
    NOTAREALFUNCTION: 'unknown', IF: 'ok'
  }, 'per-function verdicts');
  eq(rep.atRiskFunctions.map(r => r.fn), ['GROUPBY', 'AGGREGATE', 'TEXTSPLIT', 'WEBSERVICE', 'FILTER'],
    'at-risk order: count desc then name asc');
  eq(rep.atRiskFunctions.map(r => r.fn).slice(0, V.FREE_DETAIL_LIMIT),
    ['GROUPBY', 'AGGREGATE', 'TEXTSPLIT'], 'free tier details these 3; WEBSERVICE stays locked');
  eq(rep.unknownFunctions.map(r => r.fn), ['NOTAREALFUNCTION'], 'unknown list');
  // formula-level: worst function wins
  eq(formulaVerdict(rep, 'Mix', 'A4'), 'missing', 'IF+SUM+GROUPBY formula -> missing');
  eq(formulaVerdict(rep, 'Mix', 'A3'), 'unknown', 'unknown-function formula -> unknown');
  eq(formulaVerdict(rep, 'Mix', 'A5'), 'ok', 'function-free formula -> ok');
  eq(formulaVerdict(rep, 'Calc', 'C1'), 'ok', 'SUM ok for Sheets target');
  // per-function cell attribution
  const groupby = rep.atRiskFunctions[0];
  eq(groupby.cells.map(c => c.sheet + '!' + c.cell), ['Calc!C4', 'Mix!A4'], 'GROUPBY cell list');
  ok(rep.functionRows.every(r => r.verdict !== 'unknown' || r.basis === null),
    'unknown rows carry no basis');
}

console.log('e2e: Excel -> LibreOffice');
{
  const rep = V.buildReport(audit, DB, 'x', 'l');
  eq(rep.totals, { formulas: 12, sheets: 2, uniqueFunctions: 11, atRiskFormulas: 7, unknownFunctions: 1 },
    'summary tiles');
  eq(fnVerdicts(rep), {
    SUM: 'quirk', VLOOKUP: 'quirk', AGGREGATE: 'ok', GROUPBY: 'missing', FILTER: 'quirk',
    TEXTSPLIT: 'ok', WEBSERVICE: 'ok', GOOGLEFINANCE: 'missing', ARRAYFORMULA: 'missing',
    NOTAREALFUNCTION: 'unknown', IF: 'ok'
  }, 'per-function verdicts (executed lv wins over docs)');
  eq(rep.atRiskFunctions.map(r => r.fn),
    ['GROUPBY', 'ARRAYFORMULA', 'GOOGLEFINANCE', 'SUM', 'FILTER', 'VLOOKUP'],
    'missing ranked above quirks; free tier shows the 3 hard breakers');
  const basis = {};
  rep.functionRows.forEach(r => { basis[r.fn] = r.basis; });
  eq(basis.AGGREGATE, 'executed', 'AGGREGATE ok is execution-verified');
  eq(basis.WEBSERVICE, 'documented', 'WEBSERVICE ok is only documentation-based');
  eq(basis.GROUPBY, 'executed', 'GROUPBY missing is execution-verified');
  eq(basis.GOOGLEFINANCE, 'documented', 'GOOGLEFINANCE missing is documentation-based');
  eq(formulaVerdict(rep, 'Calc', 'C1'), 'quirk', 'SUM formula flagged quirk for LO target');
  eq(formulaVerdict(rep, 'Calc', 'C3'), 'ok', 'AGGREGATE formula ok for LO target');
  ok(/works since 25\.8\.7\.3/.test(rep.functionRows.find(r => r.fn === 'TEXTSPLIT').note),
    'TEXTSPLIT note carries the earliest-working LO release');
}

console.log('e2e: Google Sheets -> Excel');
{
  const rep = V.buildReport(audit, DB, 'g', 'x');
  eq(rep.totals.atRiskFormulas, 2, 'only the two Sheets-only formulas at risk');
  eq(rep.atRiskFunctions.map(r => r.fn), ['ARRAYFORMULA', 'GOOGLEFINANCE'],
    'Sheets-only functions missing in Excel (tie broken by name)');
  ok(rep.atRiskFunctions.every(r => /(documented in|executed and verified in) Google Sheets/.test(r.note)),
    'missing notes cite the Sheets source (documented or executed)');
  eq(fnVerdicts(rep).GROUPBY, 'ok', 'GROUPBY fine when moving TO Excel');
}

console.log('regression: _xlfn prefix + operator serializations (r/libreoffice bug report 2026-08-25)');
{
  const ex = X.extractFunctions;
  eq(ex('_xlfn.SINGLE(_xlfn.ANCHORARRAY(A1))').sort().join(','), 'ANCHORARRAY,SINGLE',
    'spill/@ serializations extracted with prefix stripped');
  eq(ex('_xlfn.XLOOKUP(A1,B:B,C:C)').join(','), 'XLOOKUP', '_xlfn.XLOOKUP normalizes to XLOOKUP');
  eq(ex('_xlfn._xlws.FILTER(A1:B9,A1:A9>1)').join(','), 'FILTER', '_xlfn._xlws.FILTER normalizes to FILTER');
  eq(ex('_XLFN.TEXTJOIN(",",TRUE,A1:A3)').join(','), 'TEXTJOIN', 'uppercase _XLFN. prefix stripped');
  eq(ex('SUM(A1:A3)').join(','), 'SUM', 'plain functions unaffected');

  const anchor = V.classifyFunction('ANCHORARRAY', null, 'l', 'x');
  eq(anchor.verdict, 'quirk', 'ANCHORARRAY classified quirk, not unknown');
  eq(anchor.basis, 'documented', 'ANCHORARRAY basis is documented');
  ok(/spill operator #/.test(anchor.note), 'ANCHORARRAY note explains the # operator');
  ok(/not executed these operators/.test(anchor.note), 'ANCHORARRAY note admits not executed');
  const single = V.classifyFunction('SINGLE', null, 'g', 'x');
  eq(single.verdict, 'quirk', 'SINGLE classified quirk');
  ok(/implicit-intersection operator @/.test(single.note), 'SINGLE note explains the @ operator');
  const unk = V.classifyFunction('MYMACRO', null, 'l', 'x');
  eq(unk.verdict, 'unknown', 'other unknown functions still report unknown');
}

/* ---------- unit: target LibreOffice version (per-release verdicts) ---------- */
console.log('unit: compareVersions / resolveTargetVersion / targetLabel');
{
  const cmp = V.compareVersions;
  ok(cmp('24.2.0.3', '24.8.7.2') < 0, '24.2.0.3 < 24.8.7.2');
  ok(cmp('24.8.7.2', '25.2.0.3') < 0, '24.8.7.2 < 25.2.0.3');
  ok(cmp('25.2.0.3', '25.8.7.3') < 0, '25.2.0.3 < 25.8.7.3');
  ok(cmp('25.8.7.3', '25.8.7.3') === 0, 'equal builds compare equal');
  ok(cmp('25.8.7.3', '24.2.0.3') > 0, 'newer build compares greater');
  ok(cmp('25.10.0.0', '25.9.0.0') > 0, 'minor compared numerically, not lexically (10 > 9)');
  ok(cmp('24.2', '24.2.0.3') < 0, 'missing segments count as 0');
  ok(cmp('24.2.0.0', '24.2') === 0, 'trailing zero segments are equal');
  ok(cmp('', '') === 0, 'empty strings compare equal, no throw');

  const rv = V.resolveTargetVersion;
  eq(rv('24.2'), '24.2.0.3', 'series 24.2 -> tested build');
  eq(rv('24.8'), '24.8.7.2', 'series 24.8 -> tested build');
  eq(rv('25.2'), '25.2.0.3', 'series 25.2 -> tested build');
  eq(rv('25.8'), '25.8.7.3', 'series 25.8 -> tested build');
  eq(rv('25.2.0.3'), '25.2.0.3', 'exact tested build passes through');
  eq(rv(undefined), V.LO_LATEST, 'undefined -> latest tested build (the default)');
  eq(rv(''), V.LO_LATEST, 'empty -> latest tested build');
  eq(rv('7.4.1.2'), V.LO_LATEST, 'untested build -> latest (never harsher than tested)');
  eq(V.LO_LATEST, '25.8.7.3', 'latest tested build is 25.8.7.3');
  eq(V.LO_OLDEST, '24.2.0.3', 'oldest tested build is 24.2.0.3');
  eq(V.LO_RELEASES, ['24.2.0.3', '24.8.7.2', '25.2.0.3', '25.8.7.3'], 'the four tested releases');

  eq(V.targetLabel('l', '24.8'), 'LibreOffice Calc 24.8.7.2', 'LO target label carries the build');
  eq(V.targetLabel('l'), 'LibreOffice Calc 25.8.7.3', 'LO target label defaults to latest');
  eq(V.targetLabel('g', '24.2'), 'Google Sheets', 'non-LO target label ignores the version');
  eq(V.targetLabel('x', '24.2'), 'Excel', 'Excel target label ignores the version');
}

console.log('unit: classifyFunction per target LibreOffice version');
{
  const c = V.classifyFunction;
  const at = (fn, v) => c(fn, DB[fn], 'l', 'x', v).verdict;

  // XLOOKUP: lnew 24.8.7.2 -> broken on 24.2 only.
  eq(DB.XLOOKUP.lnew, '24.8.7.2', 'dataset: XLOOKUP lnew is 24.8.7.2');
  eq(at('XLOOKUP', '24.2'), 'missing', 'XLOOKUP unsupported at target 24.2');
  eq(at('XLOOKUP', '24.8'), 'ok', 'XLOOKUP supported at target 24.8');
  eq(at('XLOOKUP', '25.2'), 'ok', 'XLOOKUP supported at target 25.2');
  eq(at('XLOOKUP', '25.8'), 'ok', 'XLOOKUP supported at target 25.8');
  const xl242 = c('XLOOKUP', DB.XLOOKUP, 'l', 'x', '24.2');
  eq(xl242.basis, 'executed', 'the version downgrade is an executed verdict');
  ok(/#NAME\? in LibreOffice 24\.2\.0\.3 \(executed\)/.test(xl242.note),
    'note says it returns #NAME? in the target build, executed');
  ok(/works since 24\.8\.7\.2/.test(xl242.note), 'note names the earliest working release');

  // VSTACK: lnew 25.8.7.3 -> broken on everything older.
  eq(DB.VSTACK.lnew, '25.8.7.3', 'dataset: VSTACK lnew is 25.8.7.3');
  eq(at('VSTACK', '24.2'), 'missing', 'VSTACK unsupported at target 24.2');
  eq(at('VSTACK', '24.8'), 'missing', 'VSTACK unsupported at target 24.8');
  eq(at('VSTACK', '25.2'), 'missing', 'VSTACK unsupported at target 25.2');
  eq(at('VSTACK', '25.8'), 'ok', 'VSTACK supported at target 25.8');

  // MAP: executed unsupported in the newest build -> unsupported everywhere.
  eq(DB.MAP.lv, 'unsupported', 'dataset: MAP is executed-unsupported');
  eq(['24.2', '24.8', '25.2', '25.8'].map(v => at('MAP', v)),
    ['missing', 'missing', 'missing', 'missing'], 'MAP missing for every target version');
  ok(/not recognized in every LibreOffice release we tested/.test(
    c('MAP', DB.MAP, 'l', 'x', '24.2').note), 'MAP note covers the whole tested range');

  // SUM: lnew null (worked in the oldest build we tested) -> same verdict everywhere.
  eq(DB.SUM.lnew, null, 'dataset: SUM lnew is null');
  eq(['24.2', '24.8', '25.2', '25.8'].map(v => at('SUM', v)),
    ['quirk', 'quirk', 'quirk', 'quirk'], 'SUM stays quirky for every target version');
  ok(/quirk was measured in 25\.8\.7\.3/.test(c('SUM', DB.SUM, 'l', 'x', '24.2').note),
    'quirk note says which build the quirk was measured in');

  // lnew null + supported: we claim the tested RANGE, never "executed in 24.2".
  const agg = c('AGGREGATE', DB.AGGREGATE, 'l', 'x', '24.2');
  eq(agg.verdict, 'ok', 'AGGREGATE ok at target 24.2');
  ok(/every LibreOffice release we tested \(24\.2\.0\.3 \u2192 25\.8\.7\.3\)/.test(agg.note),
    'lnew-null note claims the tested range, not a per-version execution');
  ok(!/Executed and verified in LibreOffice 24\.2\.0\.3/.test(agg.note),
    'lnew-null note never claims execution in a build we cannot cite');

  // Documentation-only rows do not pretend to vary by release.
  const bt = c('WEBSERVICE', DB.WEBSERVICE, 'l', 'x', '24.2');
  eq(bt.basis, 'documented', 'WEBSERVICE stays documentation-based at an older target');
  ok(/no executed per-release data/.test(bt.note), 'documented row admits it has no per-release data');

  // Excel / Sheets targets ignore the LibreOffice version entirely.
  eq(c('XLOOKUP', DB.XLOOKUP, 'g', 'x', '24.2').verdict, 'ok', 'Sheets target ignores LO version');
  eq(c('XLOOKUP', DB.XLOOKUP, 'x', 'g', '24.2').verdict, 'ok', 'Excel target ignores LO version');
}

console.log('unit: default target version is byte-identical to the pre-feature engine');
{
  const c = V.classifyFunction;
  // Every dataset entry, every LO direction: omitting targetVersion must give
  // exactly the same note (and verdict) as asking for the latest tested build,
  // and the latest-build wording is the wording the old fixtures pin.
  let drift = 0;
  for (const fn of Object.keys(DB)) {
    for (const src of ['x', 'g']) {
      const a = c(fn, DB[fn], 'l', src);
      const b = c(fn, DB[fn], 'l', src, '25.8.7.3');
      const d = c(fn, DB[fn], 'l', src, '25.8');
      if (a.verdict !== b.verdict || a.note !== b.note ||
          a.verdict !== d.verdict || a.note !== d.note) drift++;
    }
  }
  eq(drift, 0, 'no drift across all ' + Object.keys(DB).length + ' dataset entries x 2 sources');
  ok(/Executed and verified in LibreOffice 25\.8\.7\.3 \(works since 24\.8\.7\.2/.test(
    c('XLOOKUP', DB.XLOOKUP, 'l', 'x').note), 'default XLOOKUP note unchanged');
}

console.log('e2e: Excel -> LibreOffice at an older target version');
{
  const latest = V.buildReport(audit, DB, 'x', 'l');
  const old252 = V.buildReport(audit, DB, 'x', 'l', '25.2');
  eq(latest.targetVersion, '25.8.7.3', 'default report targets the latest tested build');
  eq(latest.targetLabel, 'LibreOffice Calc 25.8.7.3', 'default report label');
  eq(old252.targetVersion, '25.2.0.3', 'older report resolves the series to a tested build');
  eq(old252.targetLabel, 'LibreOffice Calc 25.2.0.3', 'older report label');
  eq(V.buildReport(audit, DB, 'x', 'g', '24.2').targetVersion, null,
    'non-LibreOffice reports carry no target version');
  eq(V.buildReport(audit, DB, 'x', 'g', '24.2').targetLabel, 'Google Sheets',
    'non-LibreOffice report label has no version');

  // TEXTSPLIT is the fixture's lnew=25.8.7.3 function: fine on 25.8, gone on 25.2.
  eq(fnVerdicts(latest).TEXTSPLIT, 'ok', 'TEXTSPLIT ok at the default target');
  eq(fnVerdicts(old252).TEXTSPLIT, 'missing', 'TEXTSPLIT missing at target 25.2');
  ok(old252.totals.atRiskFormulas > latest.totals.atRiskFormulas,
    'an older target flags at least one more at-risk formula');
  eq(old252.totals.atRiskFormulas, 8, 'Excel -> LO 25.2: 8 at-risk formulas');
  ok(old252.atRiskFunctions.some(r => r.fn === 'TEXTSPLIT'),
    'TEXTSPLIT joins the at-risk list at target 25.2');
  ok(!latest.atRiskFunctions.some(r => r.fn === 'TEXTSPLIT'),
    'TEXTSPLIT is not at risk at the default target');
  // Everything else in the fixture is unaffected by the version choice.
  const lv = fnVerdicts(latest), ov = fnVerdicts(old252);
  eq(Object.keys(lv).filter(fn => lv[fn] !== ov[fn]), ['TEXTSPLIT'],
    'only the lnew-gated function changes verdict between 25.8 and 25.2');
}

console.log('e2e: synthetic workbook, XLOOKUP across all four target versions');
{
  const synth = {
    functionCounts: { XLOOKUP: 1, SUM: 1 },
    formulas: [
      { sheet: 'S', cell: 'A1', formula: 'XLOOKUP(A2,B:B,C:C)', functions: ['XLOOKUP'] },
      { sheet: 'S', cell: 'A2', formula: 'SUM(B:B)', functions: ['SUM'] }
    ],
    totals: { formulas: 2, sheets: 1, uniqueFunctions: 2 }
  };
  const verdictOf = (v) => V.buildReport(synth, DB, 'x', 'l', v)
    .functionRows.find(r => r.fn === 'XLOOKUP').verdict;
  eq(['24.2', '24.8', '25.2', '25.8'].map(verdictOf),
    ['missing', 'ok', 'ok', 'ok'], 'XLOOKUP breaks only on the 24.2 target');
  eq(V.buildReport(synth, DB, 'x', 'l', '24.2').totals.atRiskFormulas, 2,
    '24.2: both formulas at risk (XLOOKUP missing + SUM quirk)');
  eq(V.buildReport(synth, DB, 'x', 'l', '25.8').totals.atRiskFormulas, 1,
    '25.8: only the SUM quirk remains');
}

console.log('\n' + pass + ' passed, ' + fail + ' failed' +
  (usedShim ? '  (deflate-raw via zlib shim: this Node lacks native deflate-raw ' +
    'DecompressionStream; browsers have it natively)' : ''));
process.exit(fail ? 1 : 0);
