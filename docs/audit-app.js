/*
 * Migration Audit — page logic (DOM only; all parsing lives in audit.js and
 * all verdict logic in audit-verdicts.js, both pure and unit-tested).
 *
 * PRIVACY INVARIANT: the user's file is parsed entirely in this browser tab.
 * The only network requests this page ever makes are
 *   1. fetching the site's own verdict dataset (DATA_URL, same-origin),
 *   2. fetching the site's own guide index (GUIDES_URL, same-origin — a
 *      static "does this function have a divergence writeup" lookup, no
 *      file data involved; failure is silent and never blocks the audit),
 *      and
 *   3. the OPTIONAL Gumroad license check — only when the user submits a
 *      license key, and the request contains ONLY the key, never file data.
 * Keep it that way.
 */
(function () {
  'use strict';

  /* ------------------- deploy-time constants (fill these) ------------------- */

  // Gumroad product id for the "Migration Audit full report" license
  // (Gumroad dashboard -> product -> the product_id used by /v2/licenses/verify).
  // Until this is filled in, license checks fall back to offline format
  // validation — see attemptUnlock() below.
  var PRODUCT_ID = 'hgrugXCijuM5xswSnBMUXw==';

  // Public Gumroad purchase URL for the "Buy license" button.
  var GUMROAD_URL = 'https://aflabs.gumroad.com/l/kluxmi';

  var PRICE_NOW = '$19';     // launch price
  var PRICE_LATER = '$29';   // regular price (shown struck through)

  var DATA_URL = 'data/compat.json';          // relative to the deployed page
  var GUIDES_URL = 'data/guides.json';        // relative to the deployed page
  var VERIFY_URL = 'https://api.gumroad.com/v2/licenses/verify';
  var LS_KEY = 'csps-audit-license';          // localStorage: {key, status}
  var LS_VER = 'csps-audit-lo-version';       // localStorage: target LO build string

  var MAX_CELLS_PER_FUNCTION = 300;  // rendering cap; CSV export is never capped
  var MAX_FORMULA_ROWS = 2000;       // rendering cap for the full formula table

  /* ------------------- state ------------------- */

  var db = null;            // compat.json
  var guides = null;        // guides.json, or {} once we know none is available
  var auditResult = null;   // XlsxAudit.auditXlsx output
  var report = null;        // AuditVerdicts.buildReport output
  var fileName = '';
  var unlocked = false;
  var unlockStatus = null;  // 'verified' | 'format-only'

  var $ = function (id) { return document.getElementById(id); };

  function esc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  /* ------------------- data ------------------- */

  function loadDb() {
    if (db) return Promise.resolve(db);
    return fetch(DATA_URL).then(function (r) {
      if (!r.ok) throw new Error('Could not load the verdict dataset (' + r.status + ').');
      return r.json();
    }).then(function (j) { db = j; return db; });
  }

  // Guide index is an optional enhancement, never a blocker: any failure
  // (network, 404, bad JSON) just leaves `guides` empty and the audit keeps
  // working exactly as before this feature existed.
  function loadGuides() {
    if (guides) return Promise.resolve(guides);
    return fetch(GUIDES_URL).then(function (r) {
      return r.ok ? r.json() : {};
    }).then(function (j) {
      guides = j || {};
      return guides;
    }).catch(function () {
      guides = {};
      return guides;
    });
  }

  // Small "why?" link to a function's divergence guide, or '' if it has
  // none. Public content — shown in free-tier rows too.
  function guideLinkHtml(fn) {
    var g = AuditVerdicts.guidesForFunction(fn, guides);
    if (!g.length) return '';
    return ' <a href="guides/' + esc(g[0].slug) + '.html" class="whylink" ' +
      'title="' + esc(g[0].title) + '">why?</a>';
  }

  function direction() {
    var v = $('direction').value; // e.g. "x2g"
    var d = null;
    AuditVerdicts.DIRECTIONS.forEach(function (dd) { if (dd.id === v) d = dd; });
    return d || AuditVerdicts.DIRECTIONS[0];
  }

  // Target LibreOffice release. Unknown/absent values fall back to the latest
  // tested build inside the engine, so this can never fail closed.
  function loVersion() {
    var el = $('loversion');
    return el ? el.value : AuditVerdicts.LO_LATEST;
  }

  // The version picker only means anything when the target IS LibreOffice.
  function syncVersionRow() {
    var row = $('loverrow');
    if (row) row.style.display = direction().target === 'l' ? 'block' : 'none';
  }

  function saveVersionPref() {
    try { localStorage.setItem(LS_VER, loVersion()); } catch (e) { /* private mode */ }
  }

  function restoreVersionPref() {
    var saved = null;
    try { saved = localStorage.getItem(LS_VER); } catch (e) { }
    if (!saved) return;
    var el = $('loversion');
    if (!el) return;
    // Only accept a value the select actually offers (a stale build string
    // from an older deploy must not silently select nothing).
    for (var i = 0; i < el.options.length; i++) {
      if (el.options[i].value === saved) { el.value = saved; return; }
    }
  }

  /* ------------------- file handling ------------------- */

  function showError(msg) {
    $('results').style.display = 'none';
    $('error').textContent = msg;
    $('error').style.display = 'block';
  }

  function handleFile(file) {
    if (!file) return;
    if (/\.xls$/i.test(file.name)) {
      showError('"' + file.name + '" is a legacy .xls file (old binary Excel format), ' +
        'which this tool cannot read. Open it in Excel and save as .xlsx, then try again.');
      return;
    }
    if (typeof DecompressionStream === 'undefined') {
      showError('Your browser does not support DecompressionStream, which this tool needs ' +
        'to read .xlsx files. Please use a current version of Chrome, Edge, Firefox or Safari.');
      return;
    }
    Promise.all([loadDb(), loadGuides(), file.arrayBuffer()])
      .then(function (pair) { return XlsxAudit.auditXlsx(pair[2]); })
      .then(function (r) {
        auditResult = r;
        fileName = file.name;
        rebuild();
      })
      .catch(function (e) { showError(e && e.message ? e.message : String(e)); });
  }

  function rebuild() {
    if (!auditResult || !db) return;
    var d = direction();
    report = AuditVerdicts.buildReport(auditResult, db, d.source, d.target, loVersion());
    render();
  }

  /* ------------------- rendering ------------------- */

  var VERDICT_BADGE = {
    missing: '<span class="badge badge-bad">Missing</span>',
    quirk: '<span class="badge badge-quirk">Quirk</span>',
    unknown: '<span class="badge badge-unknown">Unknown</span>',
    ok: '<span class="badge badge-good">OK</span>'
  };

  function basisTag(row) {
    if (row.basis === 'executed') return '<span class="basis">execution-verified</span>';
    if (row.basis === 'documented') return '<span class="basis">documentation-based</span>';
    return '<span class="basis">not in dataset</span>';
  }

  // Per-cell detail for one function, grouped by sheet.
  function detailHtml(row) {
    var bySheet = {};
    var order = [];
    var shown = row.cells.slice(0, MAX_CELLS_PER_FUNCTION);
    shown.forEach(function (c) {
      if (!bySheet[c.sheet]) { bySheet[c.sheet] = []; order.push(c.sheet); }
      bySheet[c.sheet].push(c);
    });
    var h = '';
    order.forEach(function (sheet) {
      h += '<p class="sheetname">' + esc(sheet) + '</p><ul class="celllist">';
      bySheet[sheet].forEach(function (c) {
        h += '<li><span class="cellref">' + esc(c.cell) + '</span> <code>=' +
          esc(c.formula) + '</code></li>';
      });
      h += '</ul>';
    });
    if (row.cells.length > shown.length) {
      h += '<p class="note">Showing the first ' + shown.length + ' of ' + row.cells.length +
        ' affected formulas here — the CSV export contains every row.</p>';
    }
    return h;
  }

  function fnRowHtml(row, withDetail, open) {
    var head = VERDICT_BADGE[row.verdict] + ' <strong class="fnname">' + esc(row.fn) +
      '</strong> <span class="fncount">' + row.count + ' formula' +
      (row.count === 1 ? '' : 's') + '</span> ' + basisTag(row) + guideLinkHtml(row.fn);
    var h = '<li class="fnrow v-' + row.verdict + '">';
    if (withDetail) {
      h += '<details' + (open ? ' open' : '') + '><summary>' + head +
        '</summary><p class="fnnote">' + esc(row.note) + '</p>' + detailHtml(row) + '</details>';
    } else {
      h += '<div class="lockedrow">' + head +
        ' <span class="lock">cell-level detail in the full report</span></div>' +
        '<p class="fnnote">' + esc(row.note) + '</p>';
    }
    return h + '</li>';
  }

  function render() {
    var r = report;
    $('error').style.display = 'none';
    $('filetitle').textContent = fileName;
    var d = direction();
    $('dirlabel').textContent = AuditVerdicts.APP_NAMES[d.source] + ' → ' +
      r.targetLabel;
    // Stated on-screen and in the print report: for LibreOffice this carries
    // the exact tested build the verdicts below were computed against.
    $('targetline').textContent = 'Target: ' + r.targetLabel +
      (d.target === 'l' ? ' (verdicts executed per tested release)' : '');

    // --- summary tiles ---
    var t = r.totals;
    var tiles = [
      [t.formulas, 'formulas scanned', ''],
      [t.sheets, 'sheets', ''],
      [t.uniqueFunctions, 'unique functions', ''],
      [t.atRiskFormulas, 'at-risk formulas', t.atRiskFormulas ? 'tile-bad' : 'tile-good'],
      [t.unknownFunctions, 'unknown functions', t.unknownFunctions ? 'tile-warn' : '']
    ];
    $('tiles').innerHTML = tiles.map(function (x) {
      return '<div class="stat-card ' + x[2] + '"><span class="num">' + x[0] +
        '</span><span class="label">' + x[1] + '</span></div>';
    }).join('');

    // --- at-risk functions ---
    var freeLimit = AuditVerdicts.FREE_DETAIL_LIMIT;
    var html = '';
    if (!r.atRiskFunctions.length) {
      html = '<p class="allclear">No at-risk functions found for this target. Every ' +
        'recognized function in this workbook is ' +
        (d.target === 'x' ? 'documented' : 'either execution-verified or documented') +
        ' in ' + AuditVerdicts.APP_NAMES[d.target] + '.' +
        (r.unknownFunctions.length ? ' (But see the unknown functions below.)' : '') + '</p>';
    } else {
      html = '<ul class="fnlist">';
      r.atRiskFunctions.forEach(function (row, i) {
        var free = unlocked || i < freeLimit;
        html += fnRowHtml(row, free, !unlocked && i < freeLimit);
      });
      html += '</ul>';
      if (!unlocked && r.atRiskFunctions.length > freeLimit) {
        html += '<p class="note">Free report: cell-level detail for the top ' + freeLimit +
          ' at-risk functions (most severe first, then most used). ' +
          (r.atRiskFunctions.length - freeLimit) + ' more at-risk function' +
          (r.atRiskFunctions.length - freeLimit === 1 ? ' is' : 's are') +
          ' listed above by name and count only — the full report unlocks every cell.</p>';
      }
    }
    $('atrisk').innerHTML = html;

    // --- unknown functions ---
    if (r.unknownFunctions.length) {
      var uh = '<ul class="fnlist">';
      r.unknownFunctions.forEach(function (row) {
        uh += fnRowHtml(row, unlocked, false);
      });
      uh += '</ul>';
      $('unknown').innerHTML = uh;
      $('unknown-section').style.display = 'block';
    } else {
      $('unknown-section').style.display = 'none';
    }

    // --- all functions table ---
    var rows = r.functionRows.slice().sort(function (a, b) {
      return (AuditVerdicts.SEVERITY[b.verdict] - AuditVerdicts.SEVERITY[a.verdict]) ||
        (b.count - a.count) || (a.fn < b.fn ? -1 : 1);
    });
    $('fntable-body').innerHTML = rows.map(function (row) {
      return '<tr><td class="fnname">' + esc(row.fn) + '</td><td>' + row.count +
        '</td><td>' + VERDICT_BADGE[row.verdict] + '</td><td>' + basisTag(row) +
        guideLinkHtml(row.fn) + '</td></tr>';
    }).join('');

    // --- full per-formula table (paid) ---
    if (unlocked) {
      var shown = r.formulas.slice(0, MAX_FORMULA_ROWS);
      $('formulas-body').innerHTML = shown.map(function (f) {
        return '<tr><td>' + esc(f.sheet) + '</td><td>' + esc(f.cell) + '</td><td><code>=' +
          esc(f.formula) + '</code></td><td>' + esc(f.functions.join(' ')) + '</td><td>' +
          VERDICT_BADGE[f.verdict] + '</td></tr>';
      }).join('');
      $('formulas-note').textContent = r.formulas.length > shown.length
        ? 'Showing the first ' + shown.length + ' of ' + r.formulas.length +
          ' formulas on screen — the CSV export contains every row.'
        : r.formulas.length + ' formulas.';
      $('formulas-section').style.display = 'block';
    } else {
      $('formulas-section').style.display = 'none';
    }

    renderUnlockBox();
    $('results').style.display = 'block';
  }

  function renderUnlockBox() {
    $('unlock-free').style.display = unlocked ? 'none' : 'block';
    $('unlock-paid').style.display = unlocked ? 'block' : 'none';
    if (unlocked) {
      $('unlock-status').textContent = unlockStatus === 'verified'
        ? 'License verified with Gumroad. Full report unlocked on this browser.'
        : 'License key accepted by offline format check (Gumroad could not be reached ' +
          'for verification). Full report unlocked on this browser; the key is ' +
          're-checked online when possible.';
    }
  }

  /* ------------------- license handling ------------------- */

  function licMsg(msg, good) {
    var el = $('licmsg');
    el.textContent = msg;
    el.className = good ? 'licmsg good' : 'licmsg bad';
    el.style.display = msg ? 'block' : 'none';
  }

  // POST to Gumroad's verify endpoint. Verified working cross-origin from a
  // browser: as of 2026-08-23 api.gumroad.com answers CORS preflight with
  // access-control-allow-origin: * and allows POST + content-type.
  // Resolves {ok:bool, message:string}; rejects only on network failure.
  function verifyOnline(key, countUse) {
    var body = 'product_id=' + encodeURIComponent(PRODUCT_ID) +
      '&license_key=' + encodeURIComponent(key) +
      (countUse ? '' : '&increment_uses_count=false');
    return fetch(VERIFY_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: body
    }).then(function (resp) {
      return resp.json().then(function (j) {
        if (j && j.success) {
          var p = j.purchase || {};
          if (p.refunded || p.chargebacked || p.disputed) {
            return { ok: false, message: 'This license’s purchase was refunded or charged back, so it is no longer valid.' };
          }
          return { ok: true, message: '' };
        }
        return { ok: false, message: (j && j.message) || 'Gumroad did not recognize this license key.' };
      });
    });
  }

  function setUnlocked(key, status) {
    unlocked = true;
    unlockStatus = status;
    try {
      localStorage.setItem(LS_KEY, JSON.stringify({ key: key, status: status }));
    } catch (e) { /* private mode: unlock is session-only */ }
    licMsg('', true);
    if (report) render(); else renderUnlockBox();
  }

  function attemptUnlock() {
    var key = $('lickey').value.trim();
    if (!AuditVerdicts.looksLikeLicenseKey(key)) {
      licMsg('That does not look like a Gumroad license key (format: ' +
        'XXXXXXXX-XXXXXXXX-XXXXXXXX-XXXXXXXX, from your Gumroad receipt).', false);
      return;
    }
    if (!PRODUCT_ID) {
      // Product not wired up yet: offline format validation only. Honest
      // limitation — without a server-side proxy or the PRODUCT_ID, we cannot
      // prove a purchase; the online path below is primary once PRODUCT_ID is
      // set (Gumroad's API is CORS-open, verified 2026-08-23).
      setUnlocked(key, 'format-only');
      return;
    }
    licMsg('Checking with Gumroad…', true);
    verifyOnline(key, true).then(function (res) {
      if (res.ok) setUnlocked(key, 'verified');
      else licMsg(res.message, false);
    }).catch(function () {
      // Network unreachable (offline, corporate proxy, ad-blocker): fall back
      // to the format check rather than locking out a paying user. Re-verified
      // online on every future page load.
      setUnlocked(key, 'format-only');
    });
  }

  function restoreLicense() {
    var saved = null;
    try { saved = JSON.parse(localStorage.getItem(LS_KEY) || 'null'); } catch (e) { }
    if (!saved || !AuditVerdicts.looksLikeLicenseKey(saved.key)) return;
    unlocked = true;
    unlockStatus = saved.status === 'verified' ? 'verified' : 'format-only';
    renderUnlockBox();
    if (PRODUCT_ID) {
      // Re-verify silently without consuming an activation.
      verifyOnline(saved.key, false).then(function (res) {
        if (res.ok) { setUnlocked(saved.key, 'verified'); }
        else {
          unlocked = false; unlockStatus = null;
          try { localStorage.removeItem(LS_KEY); } catch (e) { }
          if (report) render(); else renderUnlockBox();
          licMsg('Your saved license key failed Gumroad verification: ' + res.message, false);
        }
      }).catch(function () { /* offline: keep current unlock state */ });
    }
  }

  /* ------------------- exports (paid) ------------------- */

  function downloadCsv() {
    if (!report || !unlocked) return;
    var csv = AuditVerdicts.reportToCsv(report);
    var blob = new Blob([csv], { type: 'text/csv' });
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = (fileName.replace(/\.[^.]+$/, '') || 'workbook') + '-migration-audit.csv';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(a.href);
  }

  function printReport() {
    if (!report) return;
    // Open every rendered detail block so the printout is complete.
    document.querySelectorAll('#results details').forEach(function (el) { el.open = true; });
    window.print();
  }

  /* ------------------- wiring ------------------- */

  document.addEventListener('DOMContentLoaded', function () {
    var dropzone = $('dropzone');
    var fileinput = $('fileinput');

    dropzone.addEventListener('click', function () { fileinput.click(); });
    dropzone.addEventListener('keydown', function (e) {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fileinput.click(); }
    });
    fileinput.addEventListener('change', function () { handleFile(fileinput.files[0]); });
    ['dragenter', 'dragover'].forEach(function (ev) {
      dropzone.addEventListener(ev, function (e) {
        e.preventDefault(); dropzone.classList.add('drag');
      });
    });
    ['dragleave', 'drop'].forEach(function (ev) {
      dropzone.addEventListener(ev, function (e) {
        e.preventDefault(); dropzone.classList.remove('drag');
      });
    });
    dropzone.addEventListener('drop', function (e) {
      var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
      handleFile(f);
    });

    $('direction').addEventListener('change', function () {
      syncVersionRow();
      rebuild();
    });
    $('loversion').addEventListener('change', function () {
      saveVersionPref();
      rebuild();
    });
    $('licbtn').addEventListener('click', attemptUnlock);
    $('lickey').addEventListener('keydown', function (e) {
      if (e.key === 'Enter') attemptUnlock();
    });
    $('csvbtn').addEventListener('click', downloadCsv);
    $('printbtn').addEventListener('click', printReport);
    var buy = $('buybtn');
    buy.href = GUMROAD_URL;
    $('price-now').textContent = PRICE_NOW;
    $('price-later').textContent = PRICE_LATER;

    restoreVersionPref();
    syncVersionRow();
    restoreLicense();
    loadDb().catch(function () { /* surfaced on first file drop instead */ });
    loadGuides(); // optional enhancement; already silent-fail internally
  });
})();
