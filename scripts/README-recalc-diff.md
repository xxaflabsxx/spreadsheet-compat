# `xlsx_recalc_diff.py` — quick reference

Offline, per-cell answer to: **which cells in my workbook compute
differently in LibreOffice Calc than they did in Excel?**

```bash
python3 scripts/xlsx_recalc_diff.py BOOK.xlsx
python3 scripts/xlsx_recalc_diff.py BOOK.xlsx --json report.json --md report.md --limit 50
SOFFICE_BIN=/opt/libreoffice25.8/program/soffice python3 scripts/xlsx_recalc_diff.py BOOK.xlsx
```

| Flag | Meaning |
|---|---|
| `--json OUT.json` | full machine-readable report (every cell, not just the first N) |
| `--md OUT.md` | Markdown report |
| `--limit N` | differing cells to print / tabulate (default 25) |
| `--sheet NAME` | restrict to one sheet; repeatable |
| `--include-volatile` | count `NOW`/`TODAY`/`RAND`/`RANDBETWEEN`/`RANDARRAY` diffs as real mismatches |
| `--keep-temp` | keep the stripped copy + LibreOffice output for debugging |
| `--quiet` | summary only, no per-cell listing |

| Exit | Meaning |
|---|---|
| `0` | every comparable cell matches |
| `1` | differences found |
| `2` | untrusted run, unusable input (no cached values / not an .xlsx), or error |

Requirements: `openpyxl`, and `soffice` on `PATH` or `$SOFFICE_BIN`. Single
file, no network, nothing uploaded, **your workbook is never modified** — the
tool works on a stripped throwaway copy in a temp dir.

## How to read the output

- **`Recalculation check: … TRUSTED`** — check this line first. It reports
  how many formula cells that carried an Excel cached value came back from
  LibreOffice with a value. If ~none did, LibreOffice evaluated nothing and
  the run is marked `UNTRUSTED` (exit 2); every "difference" would be an
  artifact.
- **Categories** — `match`, `volatile` (nondeterministic, expected to
  differ), `numeric_mismatch`, `text_mismatch`, `type_mismatch`,
  `value_vs_error`, `error_vs_value`, `error_vs_error` (same error family,
  different code — this breaks `ERROR.TYPE` and error-sniffing formulas),
  `missing_in_lo`, `no_excel_value`.
- A difference is a difference, **not necessarily a bug**: it may be a real
  engine divergence, an unsupported function, or a stale Excel cache.

## Gotchas

- **The file must have been saved by Excel.** Excel writes its computed
  result next to every formula; that cached result is the entire Excel side
  of this diff. Files written by openpyxl / xlsxwriter / pandas carry no
  cached values, so the tool reports that a diff is meaningless and exits 2.
- **`calcMode="manual"`** → Excel was not auto-recalculating and the cached
  values may be stale. The tool warns. Open in Excel, press F9, save.
- **LibreOffice does not recalculate xlsx files on load by default**, which
  is why the tool strips every cached `<v>` from formula cells (at the zip/XML
  level, preserving charts, styles and everything else byte-for-byte) before
  handing the copy to `soffice --headless --convert-to xlsx`.

Tests: `python3 scripts/test_xlsx_recalc_diff.py -v` (23 tests; the
end-to-end ones really execute LibreOffice).

Background and the full design writeup: "Offline companion: per-cell recalc
diff" in the repo [README](../README.md). Function-level verdicts for a whole
workbook, in-browser: <https://canispreadsheet.com/audit.html>; catalogue of
known engine divergences: <https://canispreadsheet.com/quirks.html>.
