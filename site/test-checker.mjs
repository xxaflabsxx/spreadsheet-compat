// Unit tests for the compatibility checker's migration-report logic.
//
// The checker's JS is generated inline by site/build_site.py into
// docs/checker.html, so — exactly like site/audit-page/test-adversarial.mjs —
// this suite loads the code straight out of the BUILT page and therefore tests
// what actually ships. Run `python3 site/build_site.py` first.
//
// Scope: migrate() and its helpers (okIn / xwOk, xwWhy, xwDate, xwDateLabel,
// the quirk branch, the Excel-for-the-web unmeasured block). The extractor
// funcs() is covered by site/audit-page/test-adversarial.mjs, and the audit
// page's own verdict engine by site/audit-page/test.mjs.
//
// Usage: node site/test-checker.mjs   (Node 18+, no deps)

import fs from 'node:fs';
import path from 'node:path';
import url from 'node:url';

const here = path.dirname(url.fileURLToPath(import.meta.url));
const checkerPath = path.join(here, '..', 'docs', 'checker.html');
const html = fs.readFileSync(checkerPath, 'utf8');

/* ---------- load the shipped checker script with a stub DOM ---------- */
// Two <script> blocks matter: the generated-constants one (DATA_URL,
// XW_TRANSPORT_SKIPS, XW_RUN_DATE) and the {% raw %} logic block.
function scriptBefore(marker) {
  const i = html.indexOf(marker);
  if (i < 0) throw new Error('marker not found in docs/checker.html: ' + marker);
  const open = html.lastIndexOf('<script>', i);
  const close = html.indexOf('</script>', i);
  return html.slice(open + '<script>'.length, close);
}
const constSrc = scriptBefore('XW_TRANSPORT_SKIPS=');
const logicSrc = scriptBefore('function migrate(');

// Minimal DOM: the shipped script wires listeners and calls syncLovRow() at
// load. Nothing here touches the network — check() is never invoked.
const stubEl = () => ({ value: '', style: {}, addEventListener() {}, checked: false });
const doc = {
  getElementById: () => stubEl(),
  querySelector: () => null,          // -> target() === '' at load
  querySelectorAll: () => [],
};
const loc = { hash: '', origin: 'https://canispreadsheet.com', pathname: '/checker.html' };

// eslint-disable-next-line no-new-func
const api = new Function(
  'document', 'location', 'fetch', 'history',
  constSrc + '\n' + logicSrc +
  '\nreturn {migrate, xwDate, xwDateLabel, TGT_NAME, XW_TRANSPORT_SKIPS, XW_RUN_DATE};'
)(doc, loc, () => { throw new Error('no network in tests'); }, { replaceState() {} });

const { migrate, xwDate, xwDateLabel, XW_TRANSPORT_SKIPS } = api;

/* ---------------------------- harness ---------------------------- */
let pass = 0, fail = 0;
function ok(cond, label) {
  if (cond) { pass++; console.log('  ok  ' + label); }
  else { fail++; console.log('  FAIL ' + label); }
}
const strip = (h) => h.replace(/<[^>]+>/g, '')
  .replace(/&mdash;/g, '—').replace(/&rsquo;/g, '’').replace(/&rarr;/g, '→')
  .replace(/&middot;/g, '·').replace(/&amp;/g, '&')
  .replace(/\s+/g, ' ').trim();
function has(hay, needle, label) {
  const good = strip(hay).indexOf(needle) >= 0;
  if (!good) console.log('       looked for: ' + needle + '\n       in:         ' + strip(hay).slice(0, 400));
  ok(good, label);
}
function lacks(hay, needle, label) {
  const good = strip(hay).indexOf(needle) < 0;
  if (!good) console.log('       unexpectedly found: ' + needle);
  ok(good, label);
}

/* ------------------- fixture: shape of compat.json ------------------- */
const XWVER = 'Excel for the web (recalc, 2026-09-01)';
const DB = {
  // executed OK in the web engine
  XLOOKUP:  {cat:'Lookup', x:true,  g:true,  l:true,  gv:'supported', gver:'Google Sheets (Drive import, 2026-08-29)', xwv:'supported', xwver:XWVER, lv:'supported', lver:'25.8.7.3', lnew:'24.8.7.2'},
  // executed and NOT recognized in the web engine (#NAME?)
  ADD:      {cat:'Operator', x:false, g:true, l:false, gv:'supported', gver:'Google Sheets (Drive import, 2026-09-01)', xwv:'unsupported', xwver:XWVER, lv:'unsupported', lver:'25.8.7.3', lnew:null},
  // executed, ran, returned a value the docs do not describe
  HYPGEOMDIST:{cat:'Compatibility', x:true, g:true, l:true, gv:'quirky', gver:'Google Sheets (Drive import, 2026-08-31)', xwv:'quirky', xwver:XWVER, lv:'quirky', lver:'25.8.7.3', lnew:null},
  // transport skip: the web app refused to open the workbook
  LAMBDA:   {cat:'Logical', x:true,  g:true,  l:false, gv:'quirky', gver:'Google Sheets (Drive import, 2026-08-29)', xwv:null, xwver:null, lv:'quirky', lver:'25.8.7.3', lnew:null},
  // no web verdict, no transport reason, documented for desktop Excel
  NOTRUN:   {cat:'Math', x:true,  g:true,  l:true,  gv:'supported', gver:'Google Sheets (Drive import, 2026-08-29)', xwv:null, xwver:null, lv:'supported', lver:'25.8.7.3', lnew:null},
  // no web verdict AND not documented for desktop Excel either
  NODOC:    {cat:'Google', x:false, g:true,  l:false, gv:'supported', gver:'Google Sheets (Drive import, 2026-08-29)', xwv:null, xwver:null, lv:'unsupported', lver:'25.8.7.3', lnew:null},
  // web run drew no verdict
  INCONC:   {cat:'Math', x:true,  g:true,  l:true,  gv:'supported', gver:'Google Sheets (Drive import, 2026-08-29)', xwv:'inconclusive', xwver:XWVER, lv:'supported', lver:'25.8.7.3', lnew:null},
  // a second, later web run — proves the heading shows a SPAN, not one date
  LATER:    {cat:'Math', x:true,  g:true,  l:true,  gv:'supported', gver:'Google Sheets (Drive import, 2026-08-29)', xwv:'supported', xwver:'Excel for the web (recalc, 2026-09-03)', lv:'supported', lver:'25.8.7.3', lnew:null},
};
const GUIDES = { HYPGEOMDIST: [{slug:'hypgeom-divergence', title:'HYPGEOMDIST divergence'}] };

/* ===================================================================== */
console.log('unit: Excel-for-the-web target — verdict source');
{
  const h = migrate(['XLOOKUP'], DB, 'xw', {});
  has(h, 'Migration to Excel for the web (executed 2026-09-01)', 'heading names the web engine and its executed date');
  has(h, 'Target: Excel for the web, executed 2026-09-01.', 'Target: line carries the executed date label');
  has(h, 'recalculated on open in Excel Online', 'Target: line says how it was executed');
  has(h, 'separate implementation from desktop Excel, which we do not run', 'Target: line refuses the desktop conflation');
  has(h, 'All 1 recognized function was executed in Excel for the web', 'xwv supported -> all-clear, no blocker');
  lacks(h, 'attention:', 'xwv supported -> no blocker list');
  lacks(h, 'match documented behaviour', 'the web all-clear does NOT reuse the unqualified documented-behaviour sentence');
  has(h, 'every executed case returned the value Microsoft documents', 'the web all-clear says what was actually checked');
}
{
  const h = migrate(['ADD'], DB, 'xw', {});
  has(h, '1 function needs attention', 'xwv unsupported -> blocker');
  has(h, 'Executed in Excel for the web on 2026-09-01 and not recognized — it returns #NAME? there.', 'unsupported blocker cites the executed run');
  has(h, 'This is the web app, not desktop Excel.', 'unsupported blocker refuses the desktop conflation');
}
{
  const h = migrate(['HYPGEOMDIST'], DB, 'xw', GUIDES);
  has(h, '1 function runs in Excel for the web but returned a different value', 'xwv quirky -> quirk');
  has(h, 'we do not run desktop Excel, so we cannot tell you whether the web engine diverges from the desktop one or the documentation is wrong about both', 'quirk copy keeps the ambiguity honest');
  has(h, 'executed quirk in Excel for the web (2026-09-01)', 'quirk line dates the web run');
  has(h, 'Behaves differently across apps — read: HYPGEOMDIST divergence', 'guide link survives on the web target');
  lacks(h, 'attention:', 'a quirk is not a blocker');
}

console.log('unit: Excel-for-the-web target — transport skips');
{
  ok(XW_TRANSPORT_SKIPS.length === 7, 'seven LAMBDA-family transport skips ship in the page');
  ok(['LAMBDA','LET','ISOMITTED','MAP','MAKEARRAY','REDUCE','SCAN'].every(f => XW_TRANSPORT_SKIPS.indexOf(f) >= 0), 'the skip list is the LAMBDA family');
  const h = migrate(['LAMBDA'], DB, 'xw', {});
  has(h, '1 function needs attention', 'transport-skipped function is a blocker');
  has(h, 'could not be executed in Excel for the web: its file-open rejects workbooks carrying this serialization', 'transport blocker uses the transport wording');
  has(h, 'No verdict, and not a missing function', 'transport blocker does not claim unsupported');
  lacks(h, 'and not recognized', 'transport blocker never borrows the executed-unsupported sentence');
  // d.x is true for LAMBDA: the documented desktop flag must NOT rescue it.
  lacks(h, 'falls back to Microsoft', 'a transport skip gets no documentation fallback at all');
}

console.log('unit: Excel-for-the-web target — documented fallback (null xwv)');
{
  const h = migrate(['NOTRUN'], DB, 'xw', {});
  lacks(h, 'attention:', 'null xwv + documented desktop flag -> not a blocker');
  has(h, '0 of 1 recognized functions ran in Excel for the web', 'the all-clear counts measured vs unmeasured');
  has(h, 'so this is not an all-clear for them', 'the all-clear refuses to cover the unmeasured ones');
  has(h, '1 function could not be checked against a run of Excel for the web', 'unmeasured block is rendered');
  has(h, 'falls back to Microsoft’s documentation for desktop Excel — that is not a measurement of the web app', 'the fallback is labelled as desktop documentation');
  has(h, 'absent from our executed Excel-for-the-web set', 'unmeasured row says why');
}
{
  const h = migrate(['INCONC'], DB, 'xw', {});
  lacks(h, 'attention:', 'inconclusive xwv falls back to the documented desktop flag');
  has(h, 'our Excel-for-the-web run drew no verdict for it', 'inconclusive row says the run drew no verdict');
}
{
  const h = migrate(['NODOC'], DB, 'xw', {});
  has(h, '1 function needs attention', 'null xwv and no desktop documentation -> blocker');
  has(h, 'which documents the desktop product, not the web app', 'no-doc blocker names the documentation it checked');
}

console.log('unit: Excel-for-the-web target — dates');
{
  ok(xwDate(DB.XLOOKUP) === '2026-09-01', 'xwDate reads the per-function executed date out of xwver');
  ok(xwDate(DB.LAMBDA) === api.XW_RUN_DATE, 'xwDate falls back to the file-level run date when xwver is null');
  ok(xwDateLabel(['XLOOKUP'], DB) === 'executed 2026-09-01', 'one run date -> one date');
  ok(xwDateLabel(['XLOOKUP','LATER'], DB) === 'executed 2026-09-01–2026-09-03, each function on its own run date', 'two run dates -> a span');
  ok(xwDateLabel(['LAMBDA'], DB) === '', 'no measured function -> no date label at all (the heading must not imply a run)');
  const hn = migrate(['LAMBDA'], DB, 'xw', {});
  has(hn, 'Migration to Excel for the web', 'heading drops the parenthetical when nothing was measured');
  lacks(hn, 'Migration to Excel for the web (', 'heading carries no empty parenthetical');
  has(hn, 'nothing in this formula was measured there', 'Target: line says nothing was measured');
  has(hn, 'Our most recent Excel-web run was ' + api.XW_RUN_DATE, 'Target: line still names the file-level run date, labelled as such');
  const h = migrate(['XLOOKUP','LATER'], DB, 'xw', {});
  has(h, 'Migration to Excel for the web (executed 2026-09-01–2026-09-03, each function on its own run date)', 'heading shows the span, not a single date');
}

console.log('unit: mixed report + ordering');
{
  const h = migrate(['XLOOKUP','ADD','HYPGEOMDIST','LAMBDA','NOTRUN'], DB, 'xw', GUIDES);
  has(h, '2 functions need attention', 'ADD + LAMBDA are the two blockers');
  has(h, '1 function runs in Excel for the web', 'HYPGEOMDIST is the quirk');
  has(h, '1 function could not be checked against a run of Excel for the web', 'NOTRUN is the unmeasured one');
}

console.log('unit: the other targets are unchanged');
{
  const h = migrate(['XLOOKUP'], DB, 'g', {});
  has(h, 'All recognized functions work in Google Sheets and match documented behaviour', 'Sheets all-clear wording is untouched');
  has(h, 'Target: Google Sheets — executed by us via Drive import', 'Sheets Target: line names its basis');
  lacks(h, 'Excel for the web', 'no web copy leaks into a Sheets report');
}
{
  // LibreOffice defaults to the latest tested build (the stub select is empty).
  const h = migrate(['XLOOKUP'], DB, 'l', {});
  has(h, 'Migration to LibreOffice 25.8.7.3', 'LibreOffice heading keeps its version, not a date');
  has(h, 'All recognized functions work in LibreOffice 25.8.7.3 and match documented behaviour', 'LibreOffice all-clear wording is untouched');
  lacks(h, 'executed 2026-09-01', 'no web date leaks into a LibreOffice report');
}
{
  const h = migrate(['HYPGEOMDIST'], DB, 'x', {});
  has(h, 'documentation only', 'desktop Excel target says documentation only');
  has(h, 'not from a run of ours', 'desktop Excel target denies execution');
  lacks(h, 'runs in Excel (desktop) but', 'desktop Excel target never claims a value quirk');
  has(h, 'That is a documentation check, not a run', 'desktop all-clear is scoped to documentation');
}
{
  const h = migrate(['ADD'], DB, 'x', {});
  has(h, '1 function needs attention', 'desktop Excel still reports undocumented functions as blockers');
}

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
