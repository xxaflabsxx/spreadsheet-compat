#!/usr/bin/env python3
"""Honesty guard: only LibreOffice Calc results are executed on this site.
Fails (exit 1) if the built site claims Excel or Google Sheets were verified,
tested, or executed. Negations ("not executed in Excel") and documentation
statements ("documented in all three") are allowed.

Usage: python3 scripts/check_honesty.py [docs_dir]
"""
import re, sys, glob, os, html

ROOT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(__file__), "..", "docs")
AFFIRMATIVE = re.compile(
    r"(verified (?:in|across|on) all three|tested (?:in|across|on) all three|"
    r"(?:verified|tested|executed|confirmed) (?:in|on) (?:excel|google sheets)(?! *\(documented)|"
    r"verified formula for excel|execution-verified compatibility data(?! *\(libreoffice)|"
    r"every result (?:is )?verified(?! in libreoffice)|machine-verified (?:excel|google sheets))",
    re.I,
)
NEGATION = re.compile(r"(not|never|haven't|have not|didn't|did not|can't|cannot|nor|rather than|instead of|without)\W+(?:\w+\W+){0,4}$", re.I)

def strip_tags(s):
    s = re.sub(r"<script.*?</script>|<style.*?</style>", " ", s, flags=re.S)
    return html.unescape(re.sub(r"<[^>]+>", " ", s))

bad = []
files = glob.glob(os.path.join(ROOT, "**", "*.html"), recursive=True)
for f in files:
    text = strip_tags(open(f, encoding="utf-8", errors="replace").read())
    for m in AFFIRMATIVE.finditer(text):
        before = text[max(0, m.start() - 60):m.start()]
        if NEGATION.search(before):
            continue
        bad.append((os.path.relpath(f, ROOT), text[max(0, m.start() - 50):m.end() + 40].replace("\n", " ")))

print(f"honesty check: {len(files)} pages, {len(bad)} affirmative cross-app verification claims")
for f, ctx in bad[:40]:
    print(f"  {f}: …{ctx}…")
sys.exit(1 if bad else 0)
