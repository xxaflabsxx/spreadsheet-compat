#!/usr/bin/env python3
"""Verify each recipe's example formula by actually executing it in headless
LibreOffice (same convert-to recalc trick as the compat harness). Writes
results/recipes-verified.json: slug -> {verified, engine_version, actual, expected}."""
import json, glob, os, subprocess, tempfile, re
import openpyxl
from openpyxl.worksheet.formula import ArrayFormula

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SOFF = "soffice"
import sys
sys.path.insert(0, os.path.join(ROOT, "harness"))
from xlfn_map import to_storage_formula_all  # prefix modern funcs (_xlfn.) for OOXML

def lo_version():
    out = subprocess.run([SOFF,"--version"],capture_output=True,text=True,timeout=30).stdout
    for t in out.split():
        if t[:1].isdigit() and "." in t: return t
    return "unknown"

def norm(v):
    if isinstance(v,float) and v.is_integer(): return int(v)
    return v

def run_case(setup, formula, check_range):
    wb=openpyxl.Workbook(); ws=wb.active
    for a,val in (setup or {}).items(): ws[a]=val
    anchor = check_range.split(":")[0] if check_range else "H1"
    formula = to_storage_formula_all(formula)  # add _xlfn. prefixes for OOXML round-trip
    if check_range: ws[anchor]=ArrayFormula(check_range, formula)
    else: ws[anchor]=formula
    ws["Z1"]="=1+1"  # recalc canary
    d=tempfile.mkdtemp(); p=os.path.join(d,"in.xlsx"); wb.save(p)
    outd=os.path.join(d,"out"); os.makedirs(outd,exist_ok=True)
    subprocess.run([SOFF,"--headless","--convert-to","xlsx","--outdir",outd,p],
                   capture_output=True,timeout=120)
    wb2=openpyxl.load_workbook(os.path.join(outd,"in.xlsx"),data_only=True); ws2=wb2.active
    assert ws2["Z1"].value==2, "recalc canary failed"
    if check_range:
        mn=openpyxl.utils.cell.range_boundaries(check_range)
        vals=[norm(ws2.cell(row=r,column=c).value)
              for r in range(mn[1],mn[3]+1) for c in range(mn[0],mn[2]+1)]
        return [v for v in vals if v is not None]
    return norm(ws2[anchor].value)

ver=lo_version(); out={}
for f in sorted(glob.glob(os.path.join(ROOT,"data/recipes/*.json"))):
    r=json.load(open(f)); v=r["verify"]
    exp=v["expected"]; cr=v.get("check_range")
    try:
        actual=run_case(v.get("setup_cells"), v["formula"], cr)
        ok = (actual==exp) if not isinstance(exp,list) else ([str(x) for x in actual]==[str(x) for x in exp])
    except Exception as e:
        actual=f"ERR {e}"; ok=False
    out[r["slug"]]={"verified":bool(ok),"engine":"LibreOffice Calc","engine_version":ver,
                    "formula":v["formula"],"expected":exp,"actual":actual}
    print(f"  {'OK ' if ok else 'XX '} {r['slug']:42} got={actual} want={exp}")
json.dump({"generated_at_note":"stamped post-run","engine_version":ver,"recipes":out},
          open(os.path.join(ROOT,"results/recipes-verified.json"),"w"),indent=2,default=str)
print("engine:",ver,"| verified:",sum(1 for x in out.values() if x['verified']),"/",len(out))
