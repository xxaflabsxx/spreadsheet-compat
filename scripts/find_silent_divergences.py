#!/usr/bin/env python3
"""List SILENT value divergences in the executed corpus: cases where Google
Sheets, LibreOffice 25.8 and Excel for the web all returned a non-error value
and the values disagree. These are the strongest guide candidates (no error to
warn the user). Marks whether a guide already covers the function.

The detection itself lives in site/build_site.py -- silent_divergences() --
because the published page docs/silent-divergences.html is generated from it.
This script imports that function rather than keeping a second copy, so the
count printed here and the count on the page cannot drift apart.

Usage: python3 scripts/find_silent_divergences.py [--all]
  default: hide functions already covered by a guide; --all shows everything.
Kinds: LO-only (Sheets==Excel-web, LO differs), Sheets-only, Excel-web-only, all-differ.
"""
import json, glob, sys, collections
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "site"))
from build_site import silent_divergences  # noqa: E402

covered = set()
for f in glob.glob(str(ROOT / 'site/seo-pages/*.json')):
    for fn in json.load(open(f)).get('functions', []):
        covered.add(fn.upper())

rows = [(r['function'], r['case_id'], r['kind'], r['formula'],
         r['sheets_value'], r['lo_value'], r['xw_value'])
        for r in silent_divergences()]

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
