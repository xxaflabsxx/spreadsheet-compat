// Test suite for audit.js — runs in Node 18+ (DecompressionStream is global).
// Usage: node test.mjs
import fs from 'node:fs';
import path from 'node:path';
import url from 'node:url';
import { createRequire } from 'node:module';

const here = path.dirname(url.fileURLToPath(import.meta.url));
const require = createRequire(import.meta.url);

// Node 18 has DecompressionStream but not the 'deflate-raw' format (browsers and
// Node >= 20.12 / 21.2 do). If the native one rejects it, install a zlib-backed
// shim with the same {readable, writable} shape so the ZIP code path still runs.
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

const X = require(path.join(here, 'audit.js')); // plain script w/ module.exports

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

function loadFixture(name) {
  const buf = fs.readFileSync(path.join(here, name));
  // hand a real ArrayBuffer slice, exactly like file.arrayBuffer() in a browser
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
}

/* ---------- unit tests: pure helpers ---------- */
console.log('unit: helpers');
eq(X.decodeXmlEntities('a &amp; b &lt;x&gt; &quot;q&quot; &#65;&#x42;'),
  'a & b <x> "q" AB', 'decodeXmlEntities');
eq(X.colToNum('A'), 1, 'colToNum A');
eq(X.colToNum('AA'), 27, 'colToNum AA');
eq(X.numToCol(16384), 'XFD', 'numToCol XFD');
eq(X.parseCellRef('B7'), { col: 2, row: 7 }, 'parseCellRef');
eq(X.resolveZipPath('xl/', 'worksheets/sheet1.xml'), 'xl/worksheets/sheet1.xml', 'resolveZipPath rel');
eq(X.resolveZipPath('xl/', '/xl/worksheets/s.xml'), 'xl/worksheets/s.xml', 'resolveZipPath abs');
eq(X.resolveZipPath('xl/', '../custom/a.xml'), 'custom/a.xml', 'resolveZipPath dotdot');

console.log('unit: extractFunctions');
eq(X.extractFunctions('VLOOKUP(A1,$A$1:$B$5,2,FALSE)'), ['VLOOKUP'], 'single fn');
eq(X.extractFunctions('SUM(A1:A5)+MAX(B1:B5)'), ['SUM', 'MAX'], 'two fns');
eq(X.extractFunctions('IF(A1>2,"SUM(x)","MAX(")'), ['IF'],
  'fn-like text inside strings ignored');
eq(X.extractFunctions('IF(SUM(A1)>0,SUM(B1),0)'), ['IF', 'SUM'], 'dedupe within formula');
eq(X.extractFunctions('_xlfn.XLOOKUP(A1,B:B,C:C)'), ['XLOOKUP'], '_xlfn. prefix still finds fn');
eq(X.extractFunctions('A1+B2'), [], 'no functions');
eq(X.extractFunctions('TEXTJOIN(", ",TRUE,C1:C2)'), ['TEXTJOIN'], 'TEXTJOIN');

console.log('unit: shiftSharedFormula');
eq(X.shiftSharedFormula('A1*2+$A$1', 'B1', 'B3'), 'A3*2+$A$1', 'row shift, absolute kept');
eq(X.shiftSharedFormula('SUM($A$1:A1)', 'C1', 'C2'), 'SUM($A$1:A2)', 'growing range');
eq(X.shiftSharedFormula('A$1+$B2', 'D1', 'E5'), 'B$1+$B6', 'mixed anchors, col+row shift');
eq(X.shiftSharedFormula('IF(A1="A1",LOG10(A1),0)', 'B1', 'B2'),
  'IF(A2="A1",LOG10(A2),0)', 'string literal + LOG10 name untouched');
eq(X.shiftSharedFormula("'My A1 Sheet'!A1+1", 'B1', 'B2'),
  "'My A1 Sheet'!A2+1", 'quoted sheet name untouched');
eq(X.shiftSharedFormula('A1', 'B2', 'B1'), '#REF!', 'shift off the top -> #REF!');

console.log('unit: extractSheetFormulas on namespace-prefixed XML');
{
  const xml = '<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<x:sheetData><x:row r="1">' +
    '<x:c r="A1"><x:f>SUM(B1:B3)</x:f><x:v>6</x:v></x:c>' +
    '<x:c r="B1"><x:v>1</x:v></x:c>' +
    '</x:row></x:sheetData></x:worksheet>';
  const fs2 = X.extractSheetFormulas(xml, 'S');
  eq(fs2.length, 1, 'one formula found in <x:c>/<x:f>');
  eq(fs2[0] && fs2[0].cell, 'A1', 'prefixed cell ref');
  eq(fs2[0] && fs2[0].formula, 'SUM(B1:B3)', 'prefixed formula text');
}

/* ---------- ZIP-level error handling ---------- */
console.log('unit: error handling');
await X.auditXlsx(new TextEncoder().encode('this is not a zip file at all............').buffer)
  .then(() => ok(false, 'non-zip rejects'))
  .catch(e => ok(/Not a ZIP/i.test(e.message), 'non-zip rejects with clear message'));

const xlsMagic = new Uint8Array(64);
xlsMagic.set([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
await X.auditXlsx(xlsMagic.buffer)
  .then(() => ok(false, '.xls magic rejects'))
  .catch(e => ok(/legacy \.xls/i.test(e.message), '.xls magic rejects with legacy-format message'));

/* ---------- fixture: basic-2sheet.xlsx ---------- */
console.log('fixture: basic-2sheet.xlsx');
{
  const r = await X.auditXlsx(loadFixture('basic-2sheet.xlsx'));
  eq(r.sheets.map(s => s.name), ['Data', 'Summary'], 'sheet names in workbook order');
  eq(r.sheets.map(s => s.formulaCount), [5, 5], 'per-sheet formula counts');
  eq(r.totals.formulas, 10, 'total formulas');
  const byCell = {};
  for (const f of r.formulas) byCell[f.sheet + '!' + f.cell] = f;
  eq(byCell['Data!D1'].formula, 'VLOOKUP(A1,$A$1:$B$5,2,FALSE)', 'D1 text');
  eq(byCell['Data!D2'].formula, 'SUMIFS($B$1:$B$5,$A$1:$A$5,">2")',
    'D2 text (&gt; entity in criteria decoded)');
  eq(byCell['Data!D4'].formula, 'IF(A1>2,"a & b","x ""q"" y")',
    'D4 text (& entity and doubled quotes preserved)');
  eq(byCell['Summary!A5'].formula, 'IFERROR(VLOOKUP(99,Data!A1:B5,2,0),"missing & none")',
    'A5 text (cross-sheet ref, & in string)');
  for (const [fn, n] of [['VLOOKUP', 2], ['SUMIFS', 1], ['TEXTJOIN', 1], ['IF', 1],
    ['SUM', 3], ['MAX', 1], ['COUNT', 1], ['AVERAGE', 1], ['CONCATENATE', 1],
    ['ROUND', 1], ['IFERROR', 1]]) {
    eq(r.functionCounts[fn], n, 'functionCounts.' + fn);
  }
  eq(r.totals.uniqueFunctions, 11, 'unique function total');
  ok(r.formulas.every(f => f.type === 'normal'), 'all formulas typed normal');
}

/* ---------- fixture: shared-formula.xlsx ---------- */
console.log('fixture: shared-formula.xlsx');
{
  const r = await X.auditXlsx(loadFixture('shared-formula.xlsx'));
  eq(r.sheets, [{ name: 'SharedDemo', formulaCount: 8 }], 'sheet + count (2 groups: 5 + 3)');
  const byCell = {};
  for (const f of r.formulas) byCell[f.cell] = f;
  eq(byCell['B1'].formula, 'A1*2+$A$1', 'B1 master text');
  eq(byCell['B1'].sharedRole, 'master', 'B1 is master');
  eq(byCell['B1'].ref, 'B1:B5', 'B1 group ref');
  eq(byCell['B3'].formula, 'A3*2+$A$1', 'B3 member shifted (rel moves, $A$1 fixed)');
  eq(byCell['B3'].sharedMaster, 'B1', 'B3 attributes to master B1');
  eq(byCell['B5'].formula, 'A5*2+$A$1', 'B5 member shifted');
  eq(byCell['C2'].formula, 'SUM($A$1:A2)', 'C2 growing-range member');
  eq(byCell['C3'].formula, 'SUM($A$1:A3)', 'C3 growing-range member');
  eq(r.totals.sharedMasters, 2, 'shared masters counted');
  eq(r.totals.sharedMembers, 6, 'shared members counted');
  eq(r.functionCounts['SUM'], 3, 'SUM counted once per group cell (C1..C3)');
}

/* ---------- fixture: array-formula.xlsx ---------- */
console.log('fixture: array-formula.xlsx');
{
  const r = await X.auditXlsx(loadFixture('array-formula.xlsx'));
  eq(r.sheets, [{ name: 'Arrays', formulaCount: 3 }], 'sheet + count');
  const byCell = {};
  for (const f of r.formulas) byCell[f.cell] = f;
  eq(byCell['D1'].type, 'array', 'D1 is array formula');
  eq(byCell['D1'].formula, 'SUM(A1:A3*B1:B3)', 'D1 array text');
  eq(byCell['E1'].type, 'array', 'E1 is array formula');
  eq(byCell['E1'].ref, 'E1:E3', 'E1 array spill ref preserved');
  eq(byCell['E1'].formula, 'A1:A3*2', 'E1 array text');
  eq(byCell['F1'].type, 'normal', 'F1 normal formula alongside');
  eq(r.totals.arrayFormulas, 2, 'array formula total');
  eq(r.functionCounts['SUM'], 2, 'SUM count (D1 + F1)');
}

console.log('\n' + pass + ' passed, ' + fail + ' failed' +
  (usedShim ? '  (deflate-raw via zlib shim: this Node lacks native deflate-raw ' +
    'DecompressionStream; browsers have it natively)' : ''));
process.exit(fail ? 1 : 0);
