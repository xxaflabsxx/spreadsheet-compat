#!/usr/bin/env python3
"""Verify each recipe's example formula by actually executing it in headless
LibreOffice (same convert-to recalc trick as the compat harness). Writes
results/recipes-verified.json: slug -> {verified, engine_version, actual, expected}.

A recipe may also carry an optional "variants" list; each variant's "verify" is
either one check dict or a list of them (each {label?, formula, expected,
setup_cells?, check_range?}, falling back to the variant's own setup_cells).
Those are executed too and stored under the slug as
  "variants": [{"heading": ..., "checks": [{label, formula, expected, actual, verified}]}]

Usage:
  python3 scripts/verify_recipes.py              # all recipes (rewrites the file)
  python3 scripts/verify_recipes.py <slug> ...   # only those slugs, merged into
                                                 # the existing results file
"""
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

def run_case(setup, formula, check_range, setup_sheets=None):
    wb=openpyxl.Workbook(); ws=wb.active
    # Optional extra sheets so recipes can reference another tab (Sheet2!A1).
    # Written before the main sheet's data; the main formula lives on ws.
    for sheet_name, cells in (setup_sheets or {}).items():
        extra=wb.create_sheet(title=sheet_name)
        for a,val in (cells or {}).items(): extra[a]=val
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

def check(v, default_setup=None):
    """Execute one check dict; returns (actual, ok)."""
    exp=v["expected"]; cr=v.get("check_range")
    setup=v.get("setup_cells", default_setup)
    try:
        actual=run_case(setup, v["formula"], cr, v.get("setup_sheets"))
        ok = (actual==exp) if not isinstance(exp,list) else ([str(x) for x in actual]==[str(x) for x in exp])
    except Exception as e:
        actual=f"ERR {e}"; ok=False
    return actual, bool(ok)

RESULTS=os.path.join(ROOT,"results/recipes-verified.json")
only=set(sys.argv[1:])
ver=lo_version(); out={}
if only and os.path.exists(RESULTS):   # merge into existing results, don't drop the rest
    out=json.load(open(RESULTS)).get("recipes",{})
n_run=0
for f in sorted(glob.glob(os.path.join(ROOT,"data/recipes/*.json"))):
    r=json.load(open(f))
    if only and r["slug"] not in only: continue
    n_run+=1
    v=r["verify"]
    actual, ok = check(v)
    rec={"verified":ok,"engine":"LibreOffice Calc","engine_version":ver,
         "formula":v["formula"],"expected":v["expected"],"actual":actual}
    print(f"  {'OK ' if ok else 'XX '} {r['slug']:42} got={actual} want={v['expected']}")
    variants=[]
    for var in (r.get("variants") or []):
        checks = var.get("verify") or []
        if isinstance(checks, dict): checks=[checks]
        done=[]
        for c in checks:
            a, o = check(c, var.get("setup_cells"))
            if not o: ok=False
            done.append({"label":c.get("label",""),"formula":c["formula"],
                         "expected":c["expected"],"actual":a,"verified":o})
            print(f"    {'ok ' if o else 'XX '} {c['formula'][:64]:66} got={a} want={c['expected']}")
        variants.append({"heading":var.get("heading",""),"checks":done})
    if variants:
        rec["variants"]=variants
        rec["verified"]=ok   # recipe counts as verified only if every variant check passes too
    out[r["slug"]]=rec
if only and not n_run:
    print("no recipe matched:", ", ".join(sorted(only))); sys.exit(2)
json.dump({"generated_at_note":"stamped post-run","engine_version":ver,"recipes":out},
          open(RESULTS,"w"),indent=2,default=str)
print("engine:",ver,"| ran:",n_run,"| verified:",sum(1 for x in out.values() if x['verified']),"/",len(out))
