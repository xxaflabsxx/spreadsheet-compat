# Probe: openpyxl-written XLOOKUP, plain vs `_xlfn.` spelling (2026-09-12)

`xlfn-probe.xlsx` was written by openpyxl 3.1.5 (no cached values):
A1="k1" B1=10 A2="k2" B2=20; D1 `=XLOOKUP("k2",A1:A2,B1:B2)` (stored `<f>XLOOKUP(...)</f>`),
D2 `=_xlfn.XLOOKUP("k2",A1:A2,B1:B2)` (stored `<f>_xlfn.XLOOKUP(...)</f>`), D3 `=VLOOKUP("k2",A1:B2,2,0)`.

Executed results (column D, rows 1-3):

| Engine | D1 plain XLOOKUP | D2 _xlfn.XLOOKUP | D3 VLOOKUP | Method |
|---|---|---|---|---|
| LibreOffice 24.2.0.3 | #NAME? | #NAME? | 20 | headless `--convert-to csv` (pinned AppImage), csv in this dir |
| LibreOffice 24.8.7.2 | #NAME? | 20 | 20 | same |
| LibreOffice 25.2.0.3 | #NAME? | 20 | 20 | same |
| LibreOffice 25.8.7.3 | #NAME? | 20 | 20 | same (system install) |
| Excel for the web (recalc, 2026-09-12) | #NAME? | 20 | 20 | OneDrive upload, opened in Excel Online, values read on screen (no download) |

Not executed: desktop Excel (never run by us); Google Sheets under the `_xlfn.` spelling.

Excel for the web details (2026-09-12): on open it showed an "Update Workbook for Compatibility" prompt ("This workbook requires updates to function optimally in Excel"; we chose Skip, so the file was not modified); the Workbook compatibility pane reported "1 cells with incompatible formulas"; with D1 selected the formula bar displayed `=@XLOOKUP("k2",A1:A2,B1:B2)` (unprefixed name parsed as an unknown legacy name, implicit-intersection `@` inserted), while D2 displayed `=XLOOKUP("k2",A1:A2,B1:B2)` with the `_xlfn.` prefix hidden.
