/*
 * XlsxAudit — dependency-free, fully client-side .xlsx formula extractor.
 * Part of the canispreadsheet.com "Migration Audit" prototype.
 *
 * No libraries. ZIP reading is implemented by hand (central directory walk),
 * decompression uses the browser-native DecompressionStream('deflate-raw').
 * Works in browsers and in Node 18+ (for tests).
 *
 * Public API (on globalThis.XlsxAudit / module.exports):
 *   auditXlsx(arrayBuffer) -> Promise<result>   (main entry point)
 *   plus the pure helper functions, exported for unit testing.
 */
(function () {
  'use strict';

  /* ===================== ZIP reading ===================== */

  var CEN_SIG = 0x02014b50; // central directory file header
  var LOC_SIG = 0x04034b50; // local file header

  // Scan backwards for the End-Of-Central-Directory signature (PK\x05\x06).
  // EOCD is 22 bytes + up to 65535 bytes of trailing comment.
  function findEocdOffset(u8) {
    var min = Math.max(0, u8.length - 22 - 65535);
    for (var i = u8.length - 22; i >= min; i--) {
      if (u8[i] === 0x50 && u8[i + 1] === 0x4b && u8[i + 2] === 0x05 && u8[i + 3] === 0x06) {
        return i;
      }
    }
    return -1;
  }

  // Parse the central directory into a list of entry records.
  function parseCentralDirectory(u8) {
    var eocd = findEocdOffset(u8);
    if (eocd < 0) {
      throw new Error('Not a ZIP archive (no end-of-central-directory record). ' +
        'Is this really a .xlsx file?');
    }
    var dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    var count = dv.getUint16(eocd + 10, true);      // total entries
    var cdOffset = dv.getUint32(eocd + 16, true);   // central directory start
    if (count === 0xffff || cdOffset === 0xffffffff) {
      throw new Error('ZIP64 archives are not supported by this prototype.');
    }
    var td = new TextDecoder('utf-8');
    var entries = [];
    var p = cdOffset;
    for (var i = 0; i < count; i++) {
      if (dv.getUint32(p, true) !== CEN_SIG) {
        throw new Error('Corrupt ZIP: bad central directory entry at offset ' + p + '.');
      }
      var method = dv.getUint16(p + 10, true);
      var compressedSize = dv.getUint32(p + 20, true);
      var uncompressedSize = dv.getUint32(p + 24, true);
      var nameLen = dv.getUint16(p + 28, true);
      var extraLen = dv.getUint16(p + 30, true);
      var commentLen = dv.getUint16(p + 32, true);
      var localOffset = dv.getUint32(p + 42, true);
      var name = td.decode(u8.subarray(p + 46, p + 46 + nameLen));
      entries.push({
        name: name,
        method: method,
        compressedSize: compressedSize,
        uncompressedSize: uncompressedSize,
        localOffset: localOffset
      });
      p += 46 + nameLen + extraLen + commentLen;
    }
    return entries;
  }

  function inflateRaw(compressedBytes) {
    if (typeof DecompressionStream === 'undefined') {
      throw new Error('Your browser does not support DecompressionStream, which this ' +
        'tool needs to read .xlsx files. Please use a current version of Chrome, ' +
        'Edge, Firefox or Safari.');
    }
    var ds = new DecompressionStream('deflate-raw');
    var stream = new Blob([compressedBytes]).stream().pipeThrough(ds);
    return new Response(stream).arrayBuffer().then(function (ab) {
      return new Uint8Array(ab);
    });
  }

  // Read + decompress one entry. Sizes come from the central directory (the
  // local header may use a data descriptor and carry zeros).
  function readZipEntry(u8, entry) {
    var dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    if (dv.getUint32(entry.localOffset, true) !== LOC_SIG) {
      return Promise.reject(new Error('Corrupt ZIP: bad local header for "' + entry.name + '".'));
    }
    var nameLen = dv.getUint16(entry.localOffset + 26, true);
    var extraLen = dv.getUint16(entry.localOffset + 28, true);
    var start = entry.localOffset + 30 + nameLen + extraLen;
    var comp = u8.subarray(start, start + entry.compressedSize);
    if (entry.method === 0) return Promise.resolve(comp);        // stored
    if (entry.method === 8) {                                    // deflate
      try { return inflateRaw(comp); } catch (e) { return Promise.reject(e); }
    }
    return Promise.reject(new Error('Unsupported ZIP compression method ' + entry.method +
      ' for "' + entry.name + '".'));
  }

  /* ===================== tiny XML helpers ===================== */
  // We deliberately avoid DOMParser so all parsing is pure JS and testable in
  // Node. Spreadsheet XML from real producers is machine-generated and regular
  // enough for targeted regex extraction of <sheet>, <Relationship>, <c>, <f>.

  function decodeXmlEntities(s) {
    if (s.indexOf('&') === -1) return s;
    return s.replace(/&(amp|lt|gt|quot|apos|#x?[0-9A-Fa-f]+);/g, function (m, name) {
      switch (name) {
        case 'amp': return '&';
        case 'lt': return '<';
        case 'gt': return '>';
        case 'quot': return '"';
        case 'apos': return "'";
        default:
          var code = name[1] === 'x' || name[1] === 'X'
            ? parseInt(name.slice(2), 16)
            : parseInt(name.slice(1), 10);
          return isNaN(code) ? m : String.fromCodePoint(code);
      }
    });
  }

  // Get an attribute value from the attribute portion of a single tag.
  function getAttr(tagText, attrName) {
    // attribute names may carry a namespace prefix; match exact name given
    var re = new RegExp('(?:^|\\s)' + attrName.replace('.', '\\.') +
      '\\s*=\\s*("([^"]*)"|\'([^\']*)\')');
    var m = re.exec(tagText);
    if (!m) return null;
    return decodeXmlEntities(m[2] !== undefined ? m[2] : m[3]);
  }

  /* ===================== workbook structure ===================== */

  // Join a relationship target ("worksheets/sheet1.xml", "/xl/worksheets/x.xml",
  // "../foo.xml") against base directory "xl/".
  function resolveZipPath(baseDir, target) {
    if (target.charAt(0) === '/') return target.slice(1);
    var parts = (baseDir + target).split('/');
    var out = [];
    for (var i = 0; i < parts.length; i++) {
      var p = parts[i];
      if (p === '' || p === '.') continue;
      if (p === '..') out.pop();
      else out.push(p);
    }
    return out.join('/');
  }

  // workbook.xml + workbook.xml.rels -> [{name, sheetId, rid, path}]
  // Never assume sheetN.xml ordering; always go through the rels file.
  function parseWorkbookSheets(workbookXml, relsXml) {
    var rels = {};
    var relRe = /<Relationship\b[^>]*>/g;
    var m;
    while ((m = relRe.exec(relsXml)) !== null) {
      var id = getAttr(m[0], 'Id');
      var target = getAttr(m[0], 'Target');
      var mode = getAttr(m[0], 'TargetMode');
      if (id && target && mode !== 'External') {
        rels[id] = resolveZipPath('xl/', target);
      }
    }
    var sheets = [];
    var sheetRe = /<(?:\w+:)?sheet\b[^>]*>/g;
    while ((m = sheetRe.exec(workbookXml)) !== null) {
      var name = getAttr(m[0], 'name');
      var rid = getAttr(m[0], 'r:id') || getAttr(m[0], 'id');
      if (name === null) continue;
      sheets.push({
        name: name,
        sheetId: getAttr(m[0], 'sheetId'),
        rid: rid,
        path: rid && rels[rid] ? rels[rid] : null
      });
    }
    return sheets;
  }

  /* ===================== cell references ===================== */

  function colToNum(letters) {
    var n = 0;
    for (var i = 0; i < letters.length; i++) {
      n = n * 26 + (letters.charCodeAt(i) - 64);
    }
    return n;
  }

  function numToCol(n) {
    var s = '';
    while (n > 0) {
      var r = (n - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  // "B7" -> {col: 2, row: 7}
  function parseCellRef(ref) {
    var m = /^\$?([A-Z]{1,3})\$?([0-9]+)$/.exec(ref);
    if (!m) return null;
    return { col: colToNum(m[1]), row: parseInt(m[2], 10) };
  }

  /* ===================== formula text utilities ===================== */

  // Split a formula into segments, marking string/sheet-name literals so
  // transformations never touch quoted content.
  // Double quotes delimit string literals ("" = escaped quote);
  // single quotes delimit sheet names ('' = escaped quote).
  function splitFormulaLiterals(formula) {
    var segs = [];
    var cur = '';
    var i = 0;
    while (i < formula.length) {
      var ch = formula[i];
      if (ch === '"' || ch === "'") {
        if (cur) { segs.push({ literal: false, text: cur }); cur = ''; }
        var q = ch;
        var lit = q;
        i++;
        while (i < formula.length) {
          if (formula[i] === q) {
            if (formula[i + 1] === q) { lit += q + q; i += 2; continue; }
            lit += q; i++;
            break;
          }
          lit += formula[i]; i++;
        }
        segs.push({ literal: true, text: lit });
      } else {
        cur += ch; i++;
      }
    }
    if (cur) segs.push({ literal: false, text: cur });
    return segs;
  }

  // Extract function names: strip double-quoted string literals first, then
  // match ([A-Z][A-Z0-9_.]*)\s*\( , uppercase, dedupe.
  // (This mirrors how the site's compatibility checker tokenizes formulas.)
  function extractFunctions(formula) {
    var stripped = formula.replace(/"(?:[^"]|"")*"/g, '""');
    var seen = [];
    var have = {};
    var re = /([A-Z][A-Z0-9_.]*)\s*\(/g;
    var m;
    while ((m = re.exec(stripped)) !== null) {
      var fn = m[1].toUpperCase();
      if (!have[fn]) { have[fn] = true; seen.push(fn); }
    }
    return seen;
  }

  // Shift the relative parts of every A1-style reference in a shared-formula
  // master by (deltaRow, deltaCol) to reconstruct a member cell's formula.
  // Absolute parts ($) are preserved; quoted strings/sheet names untouched.
  var REF_RE = /(?<![A-Za-z0-9_.])(\$?)([A-Z]{1,3})(\$?)([1-9][0-9]{0,6})(?![A-Za-z0-9_(])/g;

  function shiftSharedFormula(masterText, masterCell, memberCell) {
    var from = parseCellRef(masterCell);
    var to = parseCellRef(memberCell);
    if (!from || !to) return masterText;
    var dRow = to.row - from.row;
    var dCol = to.col - from.col;
    if (dRow === 0 && dCol === 0) return masterText;
    var segs = splitFormulaLiterals(masterText);
    var out = '';
    for (var i = 0; i < segs.length; i++) {
      if (segs[i].literal) { out += segs[i].text; continue; }
      out += segs[i].text.replace(REF_RE, function (all, dolCol, colL, dolRow, rowS) {
        var col = dolCol ? colToNum(colL) : colToNum(colL) + dCol;
        var row = dolRow ? parseInt(rowS, 10) : parseInt(rowS, 10) + dRow;
        if (col < 1 || col > 16384 || row < 1 || row > 1048576) return '#REF!';
        return dolCol + numToCol(col) + dolRow + row;
      });
    }
    return out;
  }

  /* ===================== worksheet formula extraction ===================== */

  // Walk every <c> element in a worksheet XML and pull out its <f> child.
  // Handles: plain formulas, shared formulas (master carries text+si+ref,
  // members carry only si -> we reconstruct by shifting the master), and
  // array formulas (t="array").
  // Returns [{cell, formula, type, si?, ref?, sharedMaster?}]
  var CELL_RE = /<(?:\w+:)?c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?c>)/g;
  var F_RE = /<(?:\w+:)?f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?f>)/;

  function extractSheetFormulas(sheetXml, sheetName) {
    var results = [];
    var sharedMasters = {}; // si -> {text, cell}
    var cm;
    CELL_RE.lastIndex = 0;
    while ((cm = CELL_RE.exec(sheetXml)) !== null) {
      var attrs = cm[1] || '';
      var inner = cm[2];
      if (inner === undefined || inner.indexOf('<') === -1) continue; // no children
      var fm = F_RE.exec(inner);
      if (!fm) continue;
      var cellRef = getAttr(attrs, 'r') || '?';
      var fAttrs = fm[1] || '';
      var fText = fm[2] !== undefined ? decodeXmlEntities(fm[2]) : '';
      var fType = getAttr(fAttrs, 't'); // null | "shared" | "array" | "dataTable"
      var si = getAttr(fAttrs, 'si');
      var ref = getAttr(fAttrs, 'ref');

      if (fType === 'shared') {
        if (fText !== '') {
          // master cell of the shared group
          sharedMasters[si] = { text: fText, cell: cellRef };
          results.push({ cell: cellRef, formula: fText, type: 'shared',
            sharedRole: 'master', si: si, ref: ref || null });
        } else {
          var master = sharedMasters[si];
          if (master) {
            results.push({
              cell: cellRef,
              formula: shiftSharedFormula(master.text, master.cell, cellRef),
              type: 'shared', sharedRole: 'member', si: si,
              sharedMaster: master.cell
            });
          } else {
            // master not seen (malformed file) — record it, don't invent text
            results.push({ cell: cellRef, formula: '', type: 'shared',
              sharedRole: 'orphan-member', si: si, sharedMaster: null });
          }
        }
      } else if (fText !== '') {
        results.push({ cell: cellRef, formula: fText,
          type: fType === 'array' ? 'array' : 'normal',
          ref: ref || null });
      }
    }
    void sheetName;
    return results;
  }

  /* ===================== main entry point ===================== */

  // arrayBuffer -> {sheets:[{name, formulaCount}],
  //                 formulas:[{sheet, cell, formula, functions, type, ...}],
  //                 functionCounts:{FN: count}, totals:{...}}
  function auditXlsx(arrayBuffer) {
    return Promise.resolve().then(function () {
      var u8 = new Uint8Array(arrayBuffer);
      // Legacy .xls (BIFF/CFB) magic: D0 CF 11 E0 A1 B1 1A E1
      if (u8.length >= 4 && u8[0] === 0xD0 && u8[1] === 0xCF && u8[2] === 0x11 && u8[3] === 0xE0) {
        throw new Error('This is a legacy .xls file (old binary Excel format), which this ' +
          'tool cannot read. Open it in Excel and save as .xlsx, then try again.');
      }
      var entries = parseCentralDirectory(u8);
      var byName = {};
      for (var i = 0; i < entries.length; i++) byName[entries[i].name] = entries[i];

      if (!byName['xl/workbook.xml']) {
        throw new Error('This ZIP does not contain xl/workbook.xml — it is not an Excel ' +
          '.xlsx/.xlsm workbook.');
      }
      var td = new TextDecoder('utf-8');
      function readText(name) {
        if (!byName[name]) return Promise.resolve(null);
        return readZipEntry(u8, byName[name]).then(function (bytes) {
          return td.decode(bytes);
        });
      }

      return Promise.all([readText('xl/workbook.xml'), readText('xl/_rels/workbook.xml.rels')])
        .then(function (pair) {
          var workbookXml = pair[0];
          var relsXml = pair[1] || '';
          var sheetDefs = parseWorkbookSheets(workbookXml, relsXml);
          if (sheetDefs.length === 0) {
            throw new Error('No sheets found in workbook.xml.');
          }
          return Promise.all(sheetDefs.map(function (def) {
            if (!def.path || !byName[def.path]) {
              // e.g. chartsheet or macro sheet target we don't read
              return Promise.resolve({ def: def, formulas: [] });
            }
            return readText(def.path).then(function (xml) {
              return { def: def, formulas: extractSheetFormulas(xml, def.name) };
            });
          }));
        })
        .then(function (perSheet) {
          var result = {
            sheets: [],
            formulas: [],
            functionCounts: {},
            totals: {
              sheets: perSheet.length,
              formulas: 0,
              uniqueFunctions: 0,
              sharedMasters: 0,
              sharedMembers: 0,
              arrayFormulas: 0
            }
          };
          perSheet.forEach(function (s) {
            result.sheets.push({ name: s.def.name, formulaCount: s.formulas.length });
            s.formulas.forEach(function (f) {
              var functions = extractFunctions(f.formula);
              var rec = {
                sheet: s.def.name,
                cell: f.cell,
                formula: f.formula,
                functions: functions,
                type: f.type
              };
              if (f.sharedRole) rec.sharedRole = f.sharedRole;
              if (f.sharedMaster !== undefined) rec.sharedMaster = f.sharedMaster;
              if (f.ref) rec.ref = f.ref;
              result.formulas.push(rec);
              result.totals.formulas++;
              if (f.type === 'array') result.totals.arrayFormulas++;
              if (f.sharedRole === 'master') result.totals.sharedMasters++;
              if (f.sharedRole === 'member' || f.sharedRole === 'orphan-member') {
                result.totals.sharedMembers++;
              }
              functions.forEach(function (fn) {
                result.functionCounts[fn] = (result.functionCounts[fn] || 0) + 1;
              });
            });
          });
          result.totals.uniqueFunctions = Object.keys(result.functionCounts).length;
          return result;
        });
    });
  }

  /* ===================== exports ===================== */

  var XlsxAudit = {
    auditXlsx: auditXlsx,
    // pure helpers, exported for tests
    findEocdOffset: findEocdOffset,
    parseCentralDirectory: parseCentralDirectory,
    readZipEntry: readZipEntry,
    decodeXmlEntities: decodeXmlEntities,
    getAttr: getAttr,
    resolveZipPath: resolveZipPath,
    parseWorkbookSheets: parseWorkbookSheets,
    extractSheetFormulas: extractSheetFormulas,
    extractFunctions: extractFunctions,
    shiftSharedFormula: shiftSharedFormula,
    splitFormulaLiterals: splitFormulaLiterals,
    colToNum: colToNum,
    numToCol: numToCol,
    parseCellRef: parseCellRef
  };

  if (typeof globalThis !== 'undefined') globalThis.XlsxAudit = XlsxAudit;
  if (typeof module !== 'undefined' && module.exports) module.exports = XlsxAudit;
})();
