#!/home/jon/venv/bin/python
"""Generate the E2E verdict fixture for the Migration Audit page.

Run:  /home/jon/venv/bin/python make_fixtures.py

verdict-mix.xlsx — 2 sheets, 12 formulas mixing safe functions with real,
dataset-verified breakers so test.mjs can assert exact classifications per
migration target. Function choices (checked against docs/data/compat.json,
snapshot 2026-08-23):

  SUM, VLOOKUP, IF   x&g&l documented; lv=quirky for SUM/VLOOKUP, supported for IF
  AGGREGATE          Excel-only vs Sheets (g:false); LibreOffice lv=supported
  GROUPBY            Excel-only (g:false); LibreOffice lv=unsupported (executed #NAME?)
  ODDLYIELD          Excel-only (g:false); LibreOffice documented, lv=null (doc-only)
                     NB: this slot exists purely to exercise the DOCUMENTED-ONLY
                     verdict basis, so it must name a function the corpus has not
                     executed yet. It was BAHTTEXT until 2026-08-31, when the
                     alphabetical corpus push executed BAHTTEXT on all four
                     LibreOffice builds and lv stopped being null. The corpus
                     advances alphabetically, so when ODDLYIELD is executed in
                     turn, pick the next still-unexecuted x&&l&&!g function that
                     sorts between AGGREGATE and TEXTSPLIT (the at-risk ordering
                     assertions in test.mjs depend on that), update test.mjs, and
                     re-run this script.
  TEXTSPLIT          g:false; LibreOffice lv=supported, lnew=25.8.7.3
  FILTER             x&g documented; LibreOffice lv=quirky
  GOOGLEFINANCE      Sheets-only (x:false); l:false, lv=null
  ARRAYFORMULA       Sheets-only (x:false); l:false, lv=null
  NOTAREALFUNCTION   absent from the dataset -> UNKNOWN

The formulas are never executed by any test — only their stored text is
parsed — so writing Sheets-only functions into an .xlsx is fine.
"""
import os

from openpyxl import Workbook

HERE = os.path.dirname(os.path.abspath(__file__))


def verdict_mix():
    wb = Workbook()
    ws = wb.active
    ws.title = "Calc"
    for i in range(1, 6):
        ws.cell(row=i, column=1, value=i)          # A1:A5
        ws.cell(row=i, column=2, value=i * 10)     # B1:B5
    ws["C1"] = "=SUM(A1:A5)"
    ws["C2"] = "=VLOOKUP(A1,$A$1:$B$5,2,FALSE)"
    ws["C3"] = "=AGGREGATE(9,6,A1:A5)"
    ws["C4"] = "=GROUPBY(A1:A5,B1:B5,SUM)"
    ws["C5"] = "=FILTER(A1:A5,B1:B5>10)"
    ws["C6"] = '=TEXTSPLIT("a,b",",")'
    ws["C7"] = "=ODDLYIELD(A1,A2,A3,A4,A5,B1,B2)"

    ws2 = wb.create_sheet("Mix")
    ws2["A1"] = '=GOOGLEFINANCE("GOOG")'
    ws2["A2"] = "=ARRAYFORMULA(Calc!A1:A5*2)"
    ws2["A3"] = "=NOTAREALFUNCTION(A1)"
    ws2["A4"] = "=IF(SUM(A1)>0,GROUPBY(Calc!A1:A5,Calc!B1:B5,SUM),0)"
    ws2["A5"] = "=A1+A2"  # formula with no functions at all

    wb.save(os.path.join(HERE, "verdict-mix.xlsx"))


if __name__ == "__main__":
    verdict_mix()
    from openpyxl import load_workbook
    wb = load_workbook(os.path.join(HERE, "verdict-mix.xlsx"))
    print("verdict-mix.xlsx ->", wb.sheetnames)
    print("written to", HERE)
