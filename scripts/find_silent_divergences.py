#!/usr/bin/env python3
"""List SILENT value divergences in the executed corpus: cases where Google
Sheets, LibreOffice 25.8 and Excel for the web all returned a non-error value
and the values disagree. These are the strongest guide candidates (no error to
warn the user). Marks whether a guide already covers the function.

Usage: python3 scripts/find_silent_divergences.py [--all]
  default: hide functions already covered by a guide; --all shows everything.
Kinds: LO-only (Sheets==Excel-web, LO differs), Sheets-only, Excel-web-only, all-differ.
"""
import json, glob, math, sys, collections
from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]
S = json.load(open(ROOT/'results/google-sheets.json'))['function_results']
L = json.load(open(ROOT/'results/libreoffice-25.8.json'))['function_results']
X = json.load(open(ROOT/'results/excel-web.json'))['function_results']
VOLATILE = {'RAND', 'RANDBETWEEN', 'RANDARRAY', 'NOW', 'TODAY', 'INFO', 'CELL'}
covered = set()
for f in glob.glob(str(ROOT/'site/seo-pages/*.json')):
    for fn in json.load(open(f)).get('functions', []):
        covered.add(fn.upper())

def norm(v):
    if isinstance(v, bool): return ('b', v)
    if isinstance(v, (int, float)): return ('n', v)
    return ('s', str(v).strip())

def same(a, b):
    a, b = norm(a), norm(b)
    if a[0] == 'n' and b[0] == 'n':
        return math.isclose(a[1], b[1], rel_tol=1e-9, abs_tol=1e-9)
    return a == b

rows = []
for fn, tests in S.items():
    if fn not in L or fn not in X or fn in VOLATILE: continue
    for tid, t in tests.items():
        if tid == 'executed_at' or not isinstance(t, dict): continue
        l, x = L[fn].get(tid), X[fn].get(tid)
        if not l or not x: continue
        if t.get('error') or l.get('error') or x.get('error'): continue
        vs, vl, vx = t.get('value'), l.get('value'), x.get('value')
        if vs is None or vl is None or vx is None: continue
        if any(isinstance(v, (list, dict)) for v in (vs, vl, vx)): continue
        sl, sx, lx = same(vs, vl), same(vs, vx), same(vl, vx)
        if sl and sx: continue
        kind = ('LO-only' if sx and not sl else 'Sheets-only' if lx and not sl
                else 'Excel-web-only' if sl and not sx else 'all-differ')
        rows.append((fn, tid, kind, t['formula_display'], vs, vl, vx))

show_all = '--all' in sys.argv
byfn = collections.defaultdict(list)
for r in rows: byfn[r[0]].append(r)
print(f'silent divergences: {len(rows)} cases in {len(byfn)} functions '
      f'({collections.Counter(r[2] for r in rows)}); guide-covered functions '
      f'{"shown" if show_all else "hidden"}')
for fn, rs in sorted(byfn.items(), key=lambda kv: -len(kv[1])):
    cov = fn in covered
    if cov and not show_all: continue
    print(f'{fn:16} n={len(rs):2} covered={cov} {dict(collections.Counter(r[2] for r in rs))}')
    for r in rs[:3]:
        print(f'    {r[3][:58]:58} S={str(r[4])[:20]:20} LO={str(r[5])[:20]:20} XW={str(r[6])[:20]}')
