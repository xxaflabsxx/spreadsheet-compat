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

/* ---------- unit: classifyFunction — every branch, synthetic entries ---------- */
console.log('unit: classifyFunction branches');
{
  const c = V.classifyFunction;
  // unknown: not in dataset
  let r = c('WHATEVER', undefined, 'g', 'x');
  eq([r.verdict, r.basis], ['unknown', null], 'absent entry -> unknown, no basis');
  ok(/do not guess/i.test(r.note), 'unknown note says we do not guess');

  // Excel/Sheets targets are documentation-based
  r = c('F', { x: true, g: true, l: false, lv: null }, 'g', 'x');
  eq([r.verdict, r.basis], ['ok', 'documented'], 'documented in Sheets -> ok/documented');
  ok(/documentation-based/.test(r.note), 'Sheets ok note admits documentation basis');

  r = c('F', { x: true, g: false, l: false, lv: null }, 'g', 'x');
  eq([r.verdict, r.basis], ['missing', 'documented'], 'g:false -> missing for Sheets target');
  ok(/documented in Excel/.test(r.note), 'missing note cites source-app presence (Excel)');

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
  eq(c('AGGREGATE', DB.AGGREGATE, 'l', 'x').verdict, 'ok', 'AGGREGATE ok in LO (executed)');
  eq(c('AGGREGATE', DB.AGGREGATE, 'l', 'x').basis, 'executed', 'AGGREGATE LO basis executed');
  eq(c('GROUPBY', DB.GROUPBY, 'l', 'x').verdict, 'missing', 'GROUPBY missing in LO (executed #NAME?)');
  eq(c('FILTER', DB.FILTER, 'l', 'x').verdict, 'quirk', 'FILTER quirk in LO');
  eq(c('FILTER', DB.FILTER, 'g', 'x').verdict, 'ok', 'FILTER ok in Sheets');
  eq(c('SUM', DB.SUM, 'l', 'x').verdict, 'quirk', 'even SUM is an executed quirk in LO');
  eq(c('SUM', DB.SUM, 'g', 'x').verdict, 'ok', 'SUM ok in Sheets');
  eq(c('GOOGLEFINANCE', DB.GOOGLEFINANCE, 'x', 'g').verdict, 'missing', 'GOOGLEFINANCE missing in Excel');
  eq(c('BAHTTEXT', DB.BAHTTEXT, 'l', 'x'), {
    verdict: 'ok', basis: 'documented',
    note: c('BAHTTEXT', DB.BAHTTEXT, 'l', 'x').note
  }, 'BAHTTEXT LO doc-only -> ok/documented');
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
  eq(lines[0], 'sheet,cell,formula,functions,verdict', 'CSV header');
  eq(lines[1], '"S,1",A1,"=IF(A1>2,""a,b"",""c"")",IF,ok', 'CSV row quoted correctly');
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
    ['FILTER', 1], ['TEXTSPLIT', 1], ['BAHTTEXT', 1], ['GOOGLEFINANCE', 1],
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
  eq(rep.totals, { formulas: 12, sheets: 2, uniqueFunctions: 11, atRiskFormulas: 5, unknownFunctions: 1 },
    'summary tiles');
  eq(fnVerdicts(rep), {
    SUM: 'ok', VLOOKUP: 'ok', AGGREGATE: 'missing', GROUPBY: 'missing', FILTER: 'ok',
    TEXTSPLIT: 'missing', BAHTTEXT: 'missing', GOOGLEFINANCE: 'ok', ARRAYFORMULA: 'ok',
    NOTAREALFUNCTION: 'unknown', IF: 'ok'
  }, 'per-function verdicts');
  eq(rep.atRiskFunctions.map(r => r.fn), ['GROUPBY', 'AGGREGATE', 'BAHTTEXT', 'TEXTSPLIT'],
    'at-risk order: count desc then name asc');
  eq(rep.atRiskFunctions.map(r => r.fn).slice(0, V.FREE_DETAIL_LIMIT),
    ['GROUPBY', 'AGGREGATE', 'BAHTTEXT'], 'free tier details these 3; TEXTSPLIT stays locked');
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
    TEXTSPLIT: 'ok', BAHTTEXT: 'ok', GOOGLEFINANCE: 'missing', ARRAYFORMULA: 'missing',
    NOTAREALFUNCTION: 'unknown', IF: 'ok'
  }, 'per-function verdicts (executed lv wins over docs)');
  eq(rep.atRiskFunctions.map(r => r.fn),
    ['GROUPBY', 'ARRAYFORMULA', 'GOOGLEFINANCE', 'SUM', 'FILTER', 'VLOOKUP'],
    'missing ranked above quirks; free tier shows the 3 hard breakers');
  const basis = {};
  rep.functionRows.forEach(r => { basis[r.fn] = r.basis; });
  eq(basis.AGGREGATE, 'executed', 'AGGREGATE ok is execution-verified');
  eq(basis.BAHTTEXT, 'documented', 'BAHTTEXT ok is only documentation-based');
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
  ok(rep.atRiskFunctions.every(r => /documented in Google Sheets/.test(r.note)),
    'missing notes cite the Sheets source');
  eq(fnVerdicts(rep).GROUPBY, 'ok', 'GROUPBY fine when moving TO Excel');
}

console.log('\n' + pass + ' passed, ' + fail + ' failed' +
  (usedShim ? '  (deflate-raw via zlib shim: this Node lacks native deflate-raw ' +
    'DecompressionStream; browsers have it natively)' : ''));
process.exit(fail ? 1 : 0);
