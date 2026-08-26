#!/usr/bin/env python3
"""
Static site generator for "Can I Spreadsheet?" (working title) —
a caniuse.com-style compatibility database for spreadsheet functions
across Excel, Google Sheets, and LibreOffice Calc.

Reads:
  data/functions.json      function inventory (documented-in data)
  data/tests/<FUNC>.json   authored test cases per function
  results/<engine>-*.json  real, executed engine results

Writes (to docs/, served by GitHub Pages from the master branch's /docs dir):
  docs/index.html
  docs/quirks.html
  docs/functions/<name-lowercase>.html   (one per inventoried function)
  docs/sitemap.xml
  docs/robots.txt

Design constraints (see project brief): stdlib + jinja2 only, no external
CDNs, inline CSS/JS, mobile-first, readable with JS disabled, single script.

Every claim rendered about "executed" / "tested" behavior must trace back to
an actual results/*.json entry. Functions with no results file entry are
rendered as documentation-only inventory with an explicit "not yet
live-tested" badge — never implied to be tested.
"""

import json
import re
import shutil
from datetime import datetime, timezone
from pathlib import Path

from jinja2 import Environment, DictLoader, select_autoescape

# --------------------------------------------------------------------------
# Config — change branding/deployment details here, nowhere else.
# --------------------------------------------------------------------------

SITE_NAME = "Can I Spreadsheet?"
SITE_TAGLINE = "caniuse.com for spreadsheet functions"
BASE_URL = "https://canispreadsheet.com/"
ACCENT = "#4F46E5"
GITHUB_URL = "https://github.com/xxAFLabsxx/spreadsheet-compat"

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"
TESTS_DIR = DATA_DIR / "tests"
RESULTS_DIR = ROOT / "results"
OUT_DIR = ROOT / "docs"
SEO_PAGES_DIR = ROOT / "site" / "seo-pages"

ENGINE_ORDER = ["excel", "google_sheets", "libreoffice"]
ENGINE_LABELS = {
    "excel": "Excel",
    "google_sheets": "Google Sheets",
    "libreoffice": "LibreOffice Calc",
}

VERDICT_LABELS = {
    "supported": "Supported, behaves as documented",
    "quirky": "Quirk found",
    "unsupported": "Unsupported (not recognized)",
}
VERDICT_BADGE_CLASS = {
    "supported": "badge-good",
    "quirky": "badge-quirk",
    "unsupported": "badge-bad",
    None: "badge-unknown",
}
VERDICT_SHORT = {
    "supported": "Supported",
    "quirky": "Quirk",
    "unsupported": "Unsupported",
    None: "n/a",
}

ERROR_VALUES = {"#NAME?", "#REF!", "#VALUE!", "#NUM!", "#N/A", "#DIV/0!", "#NULL!", "#ERROR!"}


def engine_key_from_engine_name(name: str):
    n = (name or "").lower()
    if "libreoffice" in n:
        return "libreoffice"
    if "excel" in n:
        return "excel"
    if "google" in n or "sheets" in n:
        return "google_sheets"
    return None


def iso_date(iso_str: str) -> str:
    if not iso_str:
        return ""
    try:
        return datetime.fromisoformat(iso_str.replace("Z", "+00:00")).strftime("%Y-%m-%d")
    except ValueError:
        return iso_str[:10]


# --------------------------------------------------------------------------
# Data loading
# --------------------------------------------------------------------------

def load_functions():
    return json.loads((DATA_DIR / "functions.json").read_text())


def load_tests():
    """function name -> list of authored test case dicts (id, formula, ...)"""
    tests = {}
    for p in sorted(TESTS_DIR.glob("*.json")):
        d = json.loads(p.read_text())
        tests[d["function"]] = d["cases"]
    return tests


def load_results():
    """engine key -> raw results blob for that engine's results/*.json file.

    When several results files map to the same engine (e.g. multiple
    LibreOffice versions), the NEWEST version wins the live verdict, so the
    support matrix always reflects current-release behaviour.
    """
    out = {}
    for p in sorted(RESULTS_DIR.glob("*.json")):
        d = json.loads(p.read_text())
        key = engine_key_from_engine_name(d.get("engine", ""))
        if not key:
            continue
        prev = out.get(key)
        if prev is None or _version_tuple(d.get("engine_version")) >= _version_tuple(
            prev.get("engine_version")
        ):
            out[key] = d
    return out


def _version_tuple(v):
    """'25.8.7.3' -> (25, 8, 7, 3) for correct numeric version ordering."""
    if not v:
        return ()
    parts = []
    for tok in str(v).split("."):
        num = "".join(ch for ch in tok if ch.isdigit())
        parts.append(int(num) if num else 0)
    return tuple(parts)


def load_lo_versions():
    """Return every executed LibreOffice results blob, ascending by version:
    [(version_str, blob), ...]. Powers the caniuse-style version-range data
    (which functions gained support in which release)."""
    blobs = []
    for p in sorted(RESULTS_DIR.glob("libreoffice-*.json")):
        d = json.loads(p.read_text())
        if engine_key_from_engine_name(d.get("engine", "")) == "libreoffice":
            blobs.append((d.get("engine_version", ""), d))
    blobs.sort(key=lambda t: _version_tuple(t[0]))
    return blobs


# --------------------------------------------------------------------------
# Build per-function records
# --------------------------------------------------------------------------

def classify_verdict(case_results):
    """case_results: list of executed-result dicts (raw from results file) for
    one engine, for one function. Returns 'supported' | 'quirky' | 'unsupported'.
    """
    if not case_results:
        return None
    # Probe cases (expected == None, e.g. volatile NOW/RAND existence checks)
    # assert only "no error"; an error-free probe must not count against the
    # verdict, or fully-working volatile functions get mislabeled "quirky".
    def case_ok(c):
        if c.get("expected") is None and c.get("matched_expected") is None:
            return c.get("error") is None
        return bool(c.get("matched_expected"))
    if all(case_ok(c) for c in case_results):
        return "supported"
    if any(c.get("value") in ERROR_VALUES and c.get("value") == "#NAME?" for c in case_results):
        return "unsupported"
    # also treat range_values full of #NAME? as unsupported (spill formulas)
    if any(
        isinstance(c.get("range_values"), list)
        and c["range_values"]
        and all(v == "#NAME?" for v in c["range_values"])
        for c in case_results
    ):
        return "unsupported"
    return "quirky"


def build_records(functions_doc, tests_by_fn, results_by_engine, lo_versions=None):
    records = []
    all_quirks = []  # flattened, for the quirks page
    lo_versions = lo_versions or []

    for f in functions_doc["functions"]:
        name = f["name"]
        name_lower = name.lower()
        apps = f["apps"]
        authored_cases = tests_by_fn.get(name)  # list or None
        has_tests = authored_cases is not None

        engines = {}
        for ek in ENGINE_ORDER:
            app_info = apps.get(ek, {}) or {}
            res_blob = results_by_engine.get(ek)
            fn_results = (res_blob or {}).get("function_results", {}).get(name)

            entry = {
                "key": ek,
                "label": ENGINE_LABELS[ek],
                "documented": bool(app_info.get("documented")),
                "doc_url": app_info.get("url"),
                "tested": False,
                "verdict": None,
                "version": None,
                "generated_at": None,
                "cases": [],
            }

            if fn_results:
                merged_cases = []
                for c in authored_cases or []:
                    r = fn_results.get(c["id"])
                    if not r:
                        continue
                    merged_cases.append({**c, **r})
                verdict = classify_verdict(list(fn_results.values()))
                entry.update(
                    tested=True,
                    verdict=verdict,
                    version=res_blob.get("engine_version"),
                    generated_at=res_blob.get("generated_at"),
                    trusted=res_blob.get("trusted"),
                    cases=merged_cases,
                )
                for mc in merged_cases:
                    if mc.get("matched_expected") is False:
                        all_quirks.append(
                            {
                                "function": name,
                                "name_lower": name_lower,
                                "category": f["category"],
                                "engine_key": ek,
                                "engine_label": ENGINE_LABELS[ek],
                                "engine_version": res_blob.get("engine_version"),
                                "case": mc,
                            }
                        )

            # LibreOffice version history: run the SAME executed corpus under
            # each LibreOffice release we have results for, so we can show a
            # real, machine-verified "supported since version X" range rather
            # than only the current release's verdict.
            if ek == "libreoffice":
                history = []
                for vstr, blob in lo_versions:
                    vres = blob.get("function_results", {}).get(name)
                    if not vres:
                        continue
                    history.append(
                        {
                            "version": vstr,
                            "verdict": classify_verdict(list(vres.values())),
                            "generated_at": blob.get("generated_at"),
                        }
                    )
                entry["lo_history"] = history
                change = None
                if len(history) >= 2 and history[0]["verdict"] != history[-1]["verdict"]:
                    # Precise "supported since" = the FIRST tested release whose
                    # verdict is supported (with an earlier release that wasn't).
                    since = None
                    for h in history:
                        if h["verdict"] == "supported":
                            since = h["version"]
                            break
                    change = {
                        "from_version": history[0]["version"],
                        "from_verdict": history[0]["verdict"],
                        "to_version": history[-1]["version"],
                        "to_verdict": history[-1]["verdict"],
                        "since_version": since,
                        "newly_supported": (
                            history[0]["verdict"] == "unsupported"
                            and history[-1]["verdict"] == "supported"
                        ),
                    }
                entry["lo_change"] = change

            engines[ek] = entry

        any_tested = any(e["tested"] for e in engines.values())
        quirk_count = sum(
            1
            for e in engines.values()
            for c in e["cases"]
            if c.get("matched_expected") is False
        )
        tested_case_count = sum(len(e["cases"]) for e in engines.values())
        verdicts_present = [e["verdict"] for e in engines.values() if e["verdict"]]
        if "quirky" in verdicts_present:
            primary_verdict = "quirky"
        elif "unsupported" in verdicts_present:
            primary_verdict = "unsupported"
        elif "supported" in verdicts_present:
            primary_verdict = "supported"
        else:
            primary_verdict = None

        last_tested = None
        for e in engines.values():
            if e["generated_at"]:
                d = iso_date(e["generated_at"])
                if not last_tested or d > last_tested:
                    last_tested = d

        records.append(
            {
                "name": name,
                "name_lower": name_lower,
                "category": f["category"],
                "engines": engines,
                "has_tests": has_tests,
                "any_tested": any_tested,
                "quirk_count": quirk_count,
                "tested_case_count": tested_case_count,
                "primary_verdict": primary_verdict,
                "last_tested": last_tested,
            }
        )

    records.sort(key=lambda r: r["name"])
    return records, all_quirks


# --------------------------------------------------------------------------
# Templates (kept inline so the generator is a single self-contained script)
# --------------------------------------------------------------------------

CSS = """
:root {
  --accent: #4F46E5;
  --accent-dark: #3730A3;
  --bg: #ffffff;
  --bg-alt: #F8F8FC;
  --text: #1F2430;
  --text-muted: #5B6072;
  --border: #E4E4EE;
  --good: #0F7B4F;
  --good-bg: #E7F7EF;
  --bad: #B3261E;
  --bad-bg: #FDECEC;
  --quirk: #92600B;
  --quirk-bg: #FFF6E0;
  --unknown-bg: #EEEEF4;
  --unknown: #5B6072;
  font-size: 16px;
}
* { box-sizing: border-box; }
html, body { margin: 0; padding: 0; }
body {
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
  color: var(--text);
  background: var(--bg);
  line-height: 1.55;
}
.container { max-width: 960px; margin: 0 auto; padding: 0 1.25rem; }
a { color: var(--accent); text-decoration: none; }
a:hover { text-decoration: underline; }
code, .mono, .formula { font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, "Liberation Mono", monospace; }

header.site-header {
  border-bottom: 1px solid var(--border);
  padding: 1rem 0;
}
header.site-header .container {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 1rem;
  flex-wrap: wrap;
}
.brand { font-weight: 700; font-size: 1.15rem; color: var(--text); }
.brand span { color: var(--accent); }
nav.site-nav a { margin-left: 1.25rem; color: var(--text-muted); font-weight: 500; }
nav.site-nav a:first-child { margin-left: 0; }
nav.site-nav a:hover { color: var(--accent); }

main { padding: 2rem 0 4rem; }

.hero { padding: 1.5rem 0 2rem; }
.hero h1 { font-size: 2.4rem; margin: 0 0 0.5rem; }
h1 { font-size: 2.1rem; margin: 0.5rem 0 1rem; }
.quirk-h { font-size: 1.35rem; margin: 2rem 0 0.5rem; }
.quirk-h a { text-decoration: none; }
.hero p.tagline { color: var(--text-muted); font-size: 1.05rem; margin: 0 0 1.5rem; }

.search-box { position: relative; margin-bottom: 0.5rem; }
.search-box input[type="search"] {
  width: 100%;
  font-size: 1.05rem;
  padding: 0.85rem 1rem;
  border: 2px solid var(--border);
  border-radius: 10px;
  font-family: inherit;
}
.search-box input[type="search"]:focus {
  outline: none;
  border-color: var(--accent);
}
.search-hint { color: var(--text-muted); font-size: 0.8rem; margin: 0.4rem 0 1.5rem; }

.stats-grid {
  display: grid;
  grid-template-columns: repeat(2, 1fr);
  gap: 0.85rem;
  margin: 1.5rem 0 2rem;
}
@media (min-width: 640px) {
  .stats-grid { grid-template-columns: repeat(4, 1fr); }
}
.stat-card {
  background: var(--bg-alt);
  border: 1px solid var(--border);
  border-radius: 10px;
  padding: 1rem;
  text-align: center;
}
.stat-card .num { font-size: 1.6rem; font-weight: 700; color: var(--accent); display: block; }
.stat-card .label { font-size: 0.82rem; color: var(--text-muted); }

.methodology {
  background: var(--bg-alt);
  border: 1px solid var(--border); background: var(--bg-alt);
  border-radius: 6px;
  padding: 1rem 1.25rem;
  margin: 1.5rem 0 2rem;
  font-size: 0.95rem;
}
.methodology strong { color: var(--accent-dark); }

h2.section-title { font-size: 1.3rem; margin: 2.25rem 0 0.75rem; }

.top-list { list-style: none; padding: 0; margin: 0; display: grid; gap: 0.5rem; }
.top-list li {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 0.75rem;
  background: var(--bg-alt);
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 0.6rem 0.9rem;
}
.top-list .fname { font-weight: 600; }
.top-list .meta { color: var(--text-muted); font-size: 0.8rem; }

.badge {
  display: inline-block;
  font-size: 0.8rem;
  font-weight: 600;
  padding: 0.2rem 0.55rem;
  border-radius: 999px;
  white-space: nowrap;
}
.badge-good { background: var(--good-bg); color: var(--good); }
.badge-bad { background: var(--bad-bg); color: var(--bad); }
.badge-quirk { background: var(--quirk-bg); color: var(--quirk); }
.badge-unknown { background: var(--unknown-bg); color: var(--unknown); }

#fn-list { list-style: none; padding: 0; margin: 0; }
#fn-list li {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 0.75rem;
  padding: 0.55rem 0.2rem;
  border-bottom: 1px solid var(--border);
  flex-wrap: wrap;
}
#fn-list li a { font-weight: 600; font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; }
#fn-list li .cat { color: var(--text-muted); font-size: 0.82rem; margin-right: auto; padding-left: 0.75rem; }
#fn-count { color: var(--text-muted); font-size: 0.8rem; }
#fn-quick { display:flex; flex-wrap:wrap; gap:.4rem .6rem; margin:.4rem 0 0; font-size:.9rem; }
#fn-quick a { padding:.2rem .55rem; border:1px solid var(--border,#e5e7eb); border-radius:999px; text-decoration:none; }
#fn-quick span { color: var(--text-muted); }

table.matrix, table.cases {
  width: 100%;
  border-collapse: collapse;
  margin: 0.75rem 0 1.5rem;
  font-size: 0.92rem;
}
table.matrix caption, table.cases caption { text-align: left; caption-side: top; }
.table-scroll { overflow-x: auto; }
table.matrix th, table.matrix td, table.cases th, table.cases td {
  border: 1px solid var(--border);
  padding: 0.55rem 0.65rem;
  text-align: left;
  vertical-align: top;
}
table.matrix th, table.cases th { background: var(--bg-alt); font-size: 0.8rem; text-transform: uppercase; letter-spacing: 0.02em; color: var(--text-muted); }
table.cases td.formula, table.cases td.result { white-space: pre-wrap; }

.func-header { display: flex; align-items: center; gap: 0.75rem; flex-wrap: wrap; margin-bottom: 0.25rem; }
.func-header h1 { font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; margin: 0; font-size: 1.9rem; }
.category-tag { color: var(--text-muted); font-size: 0.9rem; margin: 0 0 1.25rem; }

.quirk-box {
  background: var(--quirk-bg);
  border: 1px solid #F1D48A;
  border: 1px solid var(--quirk); 
  border-radius: 8px;
  padding: 1rem 1.25rem;
  margin: 1.25rem 0;
}
.quirk-box h3 { margin: 0 0 0.6rem; color: var(--quirk); font-size: 1.05rem; }
.quirk-box ul { margin: 0; padding-left: 1.1rem; }
.quirk-box li { margin-bottom: 0.6rem; }

.not-live-tested {
  background: var(--unknown-bg);
  border: 1px solid var(--border);
  border-radius: 6px;
  padding: 0.9rem 1.15rem;
  margin: 1rem 0 1.5rem;
  font-size: 0.92rem;
  color: var(--text-muted);
}

.verdict-ok { color: var(--good); }
.verdict-bad { color: var(--bad); font-weight: 600; }

/* caniuse-style LibreOffice version-range callout + history table */
.newin-box {
  background: var(--good-bg);
  border: 1px solid var(--good);
  border-radius: 8px;
  padding: 1rem 1.25rem;
  margin: 1.25rem 0;
}
.newin-box strong { color: var(--good); }
.verhist { margin: 1rem 0 0.5rem; border-collapse: collapse; }
.verhist th, .verhist td {
  border: 1px solid var(--border);
  padding: 0.4rem 0.75rem;
  text-align: left;
  font-size: 0.92rem;
}
.verhist th { background: var(--unknown-bg); font-weight: 600; }
.ver-changed td { font-weight: 600; }

.quirks-list { list-style: none; padding: 0; margin: 0; }
.quirk-entry {
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 1rem 1.15rem;
  margin-bottom: 1rem;
}
.quirk-entry h3 { margin: 0 0 0.4rem; font-size: 1.05rem; }
.quirk-entry h3 a { font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; }
.quirk-grid {
  display: grid;
  grid-template-columns: 1fr;
  gap: 0.4rem 1rem;
  font-size: 0.9rem;
  margin-top: 0.5rem;
}
@media (min-width: 640px) {
  .quirk-grid { grid-template-columns: repeat(2, 1fr); }
}
.quirk-grid dt { color: var(--text-muted); font-size: 0.8rem; text-transform: uppercase; letter-spacing: 0.03em; }
.quirk-grid dd { margin: 0 0 0.5rem; }

.promo-card {
  display: flex; align-items: center; justify-content: space-between; gap: 1.5rem;
  flex-wrap: wrap;
  margin: 2.5rem 0 1rem;
  padding: 1.25rem 1.5rem;
  border: 1px solid var(--border);
  background: var(--bg-alt);
  border-radius: 12px;
}
.promo-title { font-weight: 700; font-size: 1.1rem; margin: 0 0 0.3rem; }
.promo-body { margin: 0; color: var(--text-muted); max-width: 46rem; }
.promo-btn {
  flex-shrink: 0;
  display: inline-block;
  background: var(--accent);
  color: #fff;
  font-weight: 600;
  padding: 0.65rem 1.2rem;
  border-radius: 8px;
  text-decoration: none;
  white-space: nowrap;
}
.promo-btn:hover { filter: brightness(1.1); }

footer.site-footer {
  border-top: 1px solid var(--border);
  padding: 1.5rem 0 3rem;
  color: var(--text-muted);
  font-size: 0.8rem;
}
footer.site-footer a { color: var(--text-muted); text-decoration: underline; }

.back-link { display: inline-block; margin-bottom: 1.25rem; font-size: 0.9rem; }
noscript p { background: var(--bg-alt); padding: 0.75rem 1rem; border-radius: 8px; }
"""

SEARCH_JS = """
(function () {
  var input = document.getElementById('fn-search');
  var list = document.getElementById('fn-list');
  var count = document.getElementById('fn-count');
  var quick = document.getElementById('fn-quick');
  if (!input || !list) return;
  var items = Array.prototype.slice.call(list.children);
  var visible = [];
  function apply() {
    var q = input.value.trim().toLowerCase();
    visible = [];
    items.forEach(function (li) {
      var match = !q || li.dataset.name.indexOf(q) !== -1 || li.dataset.cat.indexOf(q) !== -1;
      li.style.display = match ? '' : 'none';
      if (match) visible.push(li);
    });
    if (count) count.textContent = visible.length + ' of ' + items.length + ' functions' + (q ? ' match \u2014 press Enter to open the best match' : '');
    if (quick) {
      quick.innerHTML = '';
      if (q) {
        // Exact name first, then names starting with the query, then the rest.
        var ranked = visible.slice().sort(function (a, b) { return rank(a, q) - rank(b, q); }).slice(0, 8);
        ranked.forEach(function (li) {
          var a = li.querySelector('a');
          if (!a) return;
          var link = document.createElement('a');
          link.href = a.getAttribute('href');
          link.textContent = a.textContent;
          quick.appendChild(link);
        });
        if (!ranked.length) {
          var none = document.createElement('span');
          none.textContent = 'No function matches \u201c' + input.value.trim() + '\u201d.';
          quick.appendChild(none);
        }
      }
    }
  }
  function rank(li, q) {
    var n = li.dataset.name;
    if (n === q) return 0;
    if (n.indexOf(q) === 0) return 1;
    if (n.indexOf(q) !== -1) return 2;
    return 3;
  }
  input.addEventListener('input', apply);
  input.addEventListener('keydown', function (e) {
    if (e.key !== 'Enter') return;
    e.preventDefault();
    var q = input.value.trim().toLowerCase();
    if (!q) return;
    var best = null;
    if (visible.length) best = visible.slice().sort(function (a, b) { return rank(a, q) - rank(b, q); })[0];
    var a = best && best.querySelector('a');
    if (a && (best.dataset.name === q || visible.length === 1)) { window.location.href = a.getAttribute('href'); return; }
    if (visible.length) list.scrollIntoView({ behavior: 'smooth', block: 'start' });
  });
  apply();
})();
"""

BASE_TMPL = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{{ page_title }}</title>
<meta name="description" content="{{ meta_description }}">
{% if noindex %}<meta name="robots" content="noindex,follow">{% endif %}
<link rel="canonical" href="{{ canonical }}">
<meta property="og:title" content="{{ page_title }}">
<meta property="og:description" content="{{ meta_description }}">
<meta property="og:type" content="website">
<meta property="og:url" content="{{ canonical }}">
<meta property="og:image" content="https://canispreadsheet.com/og.png">
<meta property="og:site_name" content="Can I Spreadsheet?">
<meta name="twitter:card" content="summary_large_image">
<meta name="twitter:image" content="https://canispreadsheet.com/og.png">
{% if json_ld %}<script type="application/ld+json">{{ json_ld | safe }}</script>{% endif %}
<style>{{ css | safe }}</style>
</head>
<body>
<header class="site-header">
  <div class="container">
    <a class="brand" href="{{ rel }}index.html">{{ site_name_html|safe }}</a>
    <nav class="site-nav">
      <a href="{{ rel }}index.html">Functions</a>
      <a href="{{ rel }}how-to/">How-to</a>
      <a href="{{ rel }}compare/">Compare</a>
      <a href="{{ rel }}guides/">Guides</a>
      <a href="{{ rel }}checker.html">Checker</a>
      <a href="{{ rel }}audit.html">Migration&nbsp;Audit</a>
      <a href="{{ rel }}excel-vs-google-sheets.html">Excel&nbsp;vs&nbsp;Sheets</a>
      <a href="{{ rel }}libreoffice-version-support.html">LO&nbsp;versions</a>
      <a href="{{ rel }}quirks.html">Quirks</a>
      <a href="{{ github_url }}">GitHub</a>
    </nav>
  </div>
</header>
<main class="container">
{% block content %}{% endblock %}
<aside class="promo-card">
  <div>
    <p class="promo-title">Migrating a whole workbook?</p>
    <p class="promo-body">Load an .xlsx into the free Migration Audit and see which of
    <em>your</em> formulas break or silently change behavior in the target app.
    Fully client-side &mdash; the file never leaves your browser.</p>
  </div>
  <a class="promo-btn" href="{{ rel }}audit.html">Run the Migration Audit</a>
</aside>
<aside class="promo-card">
  <div>
    <p class="promo-title">Tired of debugging formulas?</p>
    <p class="promo-body">We make spreadsheet templates where the formulas are already built
    and tested: budgets, debt payoff, invoicing, and a complete freelance business hub
    for Excel &amp; Google Sheets &mdash; on
    <a href="https://www.etsy.com/shop/AFLabsStudio" rel="sponsored">Etsy</a> or
    <a href="https://aflabs.gumroad.com" rel="sponsored">Gumroad</a>.</p>
  </div>
  <a class="promo-btn" href="https://www.etsy.com/shop/AFLabsStudio" rel="sponsored">Browse templates on Etsy</a>
</aside>
</main>
<footer class="site-footer">
  <div class="container">
    <p>{{ site_name }}: every result on this site was executed by a real spreadsheet
    engine and recalculation-proven, never scraped from documentation alone.
    Functions without an executed-result badge are documentation-only inventory,
    clearly marked as not yet live-tested.</p>
    <p>Every result produced by executing real formulas &mdash; <a href="{{ rel }}methodology.html">how we verify</a>. Open <a href="{{ rel }}data.html">compatibility dataset</a> and test harness on <a href="{{ github_url }}">GitHub</a>.</p>
    <p>Built by AF Labs — spreadsheet templates on <a href="https://www.etsy.com/shop/AFLabsStudio" rel="sponsored">Etsy</a> and <a href="https://aflabs.gumroad.com" rel="sponsored">Gumroad</a> · <a href="{{ rel }}audit.html">Migration Audit</a>.</p>
  </div>
</footer>
</body>
</html>
"""

INDEX_TMPL = """{% extends "base.html" %}
{% block content %}
<section class="hero">
  <h1>{{ site_name }}</h1>
  <p class="tagline">{{ site_tagline }}. Search any function to see whether it's
  documented, tested, and how Excel, Google Sheets, and LibreOffice Calc actually
  behave.</p>

  <div class="search-box">
    <input type="search" id="fn-search" placeholder="Search a function, e.g. VLOOKUP, XLOOKUP, DATEDIF..." aria-label="Search functions">
  </div>
  <p class="search-hint" id="fn-count">{{ functions|length }} of {{ functions|length }} functions</p>
  <div id="fn-quick" aria-live="polite"></div>
  <noscript><p>Search needs JavaScript. Every function is still listed below and
  fully linked; use your browser's find-in-page instead.</p></noscript>
</section>

<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:1rem;margin:1.5rem 0">
  <a href="{{ rel }}checker.html" style="display:block;padding:1rem 1.1rem;border:1px solid var(--border,#e5e7eb);border-radius:10px;text-decoration:none;color:inherit">
    <strong style="display:block;margin-bottom:.35rem">&#128269; Formula compatibility checker</strong>
    <span style="color:var(--text-muted,#6b7280);font-size:.95rem">Paste any formula &mdash; instantly see if every function works in Excel, Google Sheets &amp; current LibreOffice.</span>
  </a>
  <a href="{{ rel }}how-to/" style="display:block;padding:1rem 1.1rem;border:1px solid var(--border,#e5e7eb);border-radius:10px;text-decoration:none;color:inherit">
    <strong style="display:block;margin-bottom:.35rem">&#128221; How-to recipes</strong>
    <span style="color:var(--text-muted,#6b7280);font-size:.95rem">Copy-paste formulas for common tasks &mdash; each one executed and verified in a real engine, not just documented.</span>
  </a>
  <a href="{{ rel }}quirks.html" style="display:block;padding:1rem 1.1rem;border:1px solid var(--border,#e5e7eb);border-radius:10px;text-decoration:none;color:inherit">
    <strong style="display:block;margin-bottom:.35rem">&#9888;&#65039; Quirks &amp; gotchas</strong>
    <span style="color:var(--text-muted,#6b7280);font-size:.95rem">Where the three apps disagree on the same formula &mdash; surprising differences caught by running them.</span>
  </a>
  <a href="{{ rel }}libreoffice-version-support.html" style="display:block;padding:1rem 1.1rem;border:1px solid var(--border,#e5e7eb);border-radius:10px;text-decoration:none;color:inherit">
    <strong style="display:block;margin-bottom:.35rem">&#128200; LibreOffice by version</strong>
    <span style="color:var(--text-muted,#6b7280);font-size:.95rem">Which functions each LibreOffice release supports &mdash; XLOOKUP, FILTER, SORT &amp; 15 more, tested across versions.</span>
  </a>
  <a href="{{ rel }}compare/" style="display:block;padding:1rem 1.1rem;border:1px solid var(--border,#e5e7eb);border-radius:10px;text-decoration:none;color:inherit">
    <strong style="display:block;margin-bottom:.35rem">&#9878;&#65039; Function comparisons</strong>
    <span style="color:var(--text-muted,#6b7280);font-size:.95rem">VLOOKUP vs XLOOKUP, SUMIF vs SUMIFS, IFERROR vs IFNA &mdash; which to use when, with real compatibility data.</span>
  </a>
</div>

{% if popular_functions %}
<h2 class="section-title">Popular functions</h2>
<p style="margin:-0.3rem 0 0.6rem;color:var(--text-muted,#6b7280);font-size:.95rem">The most-searched functions, each with executed cross-app results and version history:</p>
<p style="line-height:2.1">
{% for r in popular_functions %}<a href="{{ rel }}functions/{{ r.name_lower }}.html" style="display:inline-block;padding:.15rem .6rem;margin:0 .15rem .1rem 0;border:1px solid var(--border,#e5e7eb);border-radius:999px;text-decoration:none;font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;font-size:.9rem">{{ r.name }}</a>{% endfor %}
</p>
{% endif %}

<div class="stats-grid">
  <div class="stat-card"><span class="num">{{ stats.total_functions }}</span><span class="label">Functions inventoried</span></div>
  <div class="stat-card"><span class="num">{{ stats.engines_executed }}/{{ stats.engines_targeted }}</span><span class="label">Engines executed</span></div>
  <div class="stat-card"><span class="num">{{ stats.tested_case_count }}</span><span class="label">Test cases executed</span></div>
  <div class="stat-card"><span class="num">{{ stats.quirk_count }}</span><span class="label">Quirks discovered</span></div>
</div>

<div class="methodology">
  <strong>Methodology:</strong> every result badge on this site
  traces back to a formula that was actually written into a real workbook and
  recalculated by that engine, proven with deterministic and volatile canary
  formulas on every run (see the <a href="{{ github_url }}">test harness</a>).
  Nothing here is scraped from vendor docs and presented as tested. Functions
  we haven't run through an engine yet are labeled <span class="badge badge-unknown">not yet live-tested</span> and show inventory data only.
</div>

{% if top_functions %}
<h2 class="section-title">Most compatibility-interesting functions</h2>
<ul class="top-list">
  {% for r in top_functions %}
  <li>
    <span><a class="fname" href="{{ rel }}functions/{{ r.name_lower }}.html">{{ r.name }}</a>
    <span class="meta">{{ r.category }}</span></span>
    <span>
      <span class="badge {{ verdict_class[r.primary_verdict] }}">{{ verdict_label.get(r.primary_verdict, 'Unknown') }}</span>
      <span class="meta">{{ r.quirk_count }} quirk{{ 's' if r.quirk_count != 1 else '' }}</span>
    </span>
  </li>
  {% endfor %}
</ul>
{% endif %}

<h2 class="section-title">All functions</h2>
<ul id="fn-list">
  {% for f in functions %}
  <li data-name="{{ f.name_lower }}" data-cat="{{ f.category|lower }}">
    <a href="{{ rel }}functions/{{ f.name_lower }}.html">{{ f.name }}</a>
    <span class="cat">{{ f.category }}</span>
    {% if f.any_tested %}
      <span class="badge {{ verdict_class[f.primary_verdict] }}">{{ verdict_label.get(f.primary_verdict, 'Unknown') }}</span>
    {% else %}
      <span class="badge badge-unknown">not yet live-tested</span>
    {% endif %}
  </li>
  {% endfor %}
</ul>

<script>{{ search_js | safe }}</script>
{% endblock %}
"""

FUNCTION_TMPL = """{% extends "base.html" %}
{% block content %}
<a class="back-link" href="{{ rel }}index.html">&larr; All functions</a>
<div class="func-header">
  <h1>{{ r.name }}</h1>
  {% if r.any_tested %}
    <span class="badge {{ verdict_class[r.primary_verdict] }}">{{ verdict_label.get(r.primary_verdict, 'Unknown') }}</span>
  {% else %}
    <span class="badge badge-unknown">not yet live-tested</span>
  {% endif %}
</div>
<p class="category-tag">Category: {{ r.category }}{% if r.last_tested %} &middot; Last tested {{ r.last_tested }}{% endif %}</p>

{% if r.any_tested %}
<p class="lede">Real, executed compatibility results for the <strong>{{ r.name }}</strong> function across Microsoft Excel, Google Sheets, and LibreOffice Calc &mdash; verified by actually running it. Syntax and links to each vendor&rsquo;s official documentation are below.</p>
{% endif %}

{% set le = r.engines['libreoffice'] %}
{% set since = le.lo_change.since_version if le.lo_change else None %}
{% if le.lo_change and le.lo_change.newly_supported and since %}
<div class="newin-box">
  <strong>&#10003; Supported in LibreOffice since {{ since }}.</strong>
  We ran <code>{{ r.name }}</code> under every LibreOffice release we test
  ({{ le.lo_history|map(attribute='version')|join(', ') }}):
  it returned <code>#NAME?</code> (unrecognized) in {{ le.lo_change.from_version }} and first works in {{ since }}.
  If you need <strong>{{ r.name }}</strong> in LibreOffice Calc, use {{ since }} or newer.
</div>
{% endif %}

{% if not r.any_tested %}
<div class="not-live-tested">
  <strong>Not yet live-tested.</strong> No engine has executed real test cases
  for {{ r.name }} yet. The table below reflects only whether each vendor's
  official documentation lists this function; it is inventory data, not a
  tested result. Check back as the test corpus grows, or see the
  <a href="{{ github_url }}">project repo</a> to contribute a test file.
</div>
{% endif %}

<h2 class="section-title">Support matrix</h2>
<div class="table-scroll">
<table class="matrix">
<thead><tr><th>Engine</th><th>Documented</th><th>Live-tested</th><th>Verdict</th></tr></thead>
<tbody>
{% for ek in engine_order %}
{% set e = r.engines[ek] %}
<tr>
  <td>{{ e.label }}</td>
  <td>{% if e.doc_url %}<a href="{{ e.doc_url }}">Yes</a>{% elif e.documented %}Yes{% else %}No{% endif %}</td>
  <td>{% if e.tested %}Yes ({{ e.version }}, {{ e.generated_at|dateonly }}){% else %}Not yet{% endif %}</td>
  <td>
    {% if e.verdict %}
      <span class="badge {{ verdict_class[e.verdict] }}">{{ verdict_label[e.verdict] }}</span>
    {% else %}
      <span class="badge badge-unknown">n/a</span>
    {% endif %}
  </td>
</tr>
{% endfor %}
</tbody>
</table>
</div>

{% if le.lo_history and le.lo_history|length > 1 %}
<h2 class="section-title">LibreOffice version history</h2>
<p>We executed the same test cases under each LibreOffice release to show exactly when
{{ r.name }}&rsquo;s support changed &mdash; not documentation claims, real results.</p>
<div class="table-scroll">
<table class="verhist">
<thead><tr><th>LibreOffice version</th><th>Verdict</th><th>Tested</th></tr></thead>
<tbody>
{% for h in le.lo_history %}
<tr class="{% if not loop.first and h.verdict != le.lo_history[loop.index0 - 1].verdict %}ver-changed{% endif %}">
  <td>{{ h.version }}</td>
  <td><span class="badge {{ verdict_class[h.verdict] }}">{{ verdict_label[h.verdict] }}</span></td>
  <td>{{ h.generated_at|dateonly }}</td>
</tr>
{% endfor %}
</tbody>
</table>
</div>
{% endif %}

{% if le.tested and ((le.lo_change and le.lo_change.newly_supported and since) or le.verdict in ('unsupported', 'quirky')) %}
<h2 class="section-title">Why isn't {{ r.name }} working in LibreOffice?</h2>
{% if le.lo_change and le.lo_change.newly_supported and since %}
<p>If <code>{{ r.name }}</code> returns a <code>#NAME?</code> error in LibreOffice Calc, you are
almost certainly running a release older than <strong>{{ since }}</strong> &mdash; that is exactly
what our executed tests show: <code>#NAME?</code> in {{ le.lo_change.from_version }}, working from
{{ since }} onward. Check your version under <em>Help &rarr; About LibreOffice</em> and upgrade to
{{ since }} or newer; no setting or extension enables it in older releases. (Other causes of this
error: see the <a href="{{ rel }}spreadsheet-errors.html">error values guide</a>.)</p>
{% elif le.verdict == 'unsupported' %}
<p>LibreOffice Calc does not implement <code>{{ r.name }}</code> as of {{ le.version }} &mdash; in our
executed tests it returns a <code>#NAME?</code> (unrecognized function) error. This is not a typo or a
settings problem, and saving the file as .xlsx does not change it: the function simply isn&rsquo;t
available yet{% if le.documented %} despite appearing in some documentation{% endif %}.
{% if r.engines['excel'].documented %}The same formula does work in
{% if r.engines['google_sheets'].documented %}Excel and Google Sheets{% else %}Excel{% endif %}.{% endif %}
Watch the <a href="{{ rel }}libreoffice-version-support.html">LibreOffice version support page</a> &mdash;
we re-run every test on each new release, so it will flip to Supported here as soon as it lands.</p>
{% elif le.verdict == 'quirky' %}
<p><code>{{ r.name }}</code> exists in LibreOffice {{ le.version }}, but it is not a drop-in match for
Excel &mdash; our executed tests found real behavioral differences (detailed in the test results on this
page). If a formula that works in Excel or Google Sheets misbehaves in LibreOffice, compare your usage
against the failing cases above before assuming your data is wrong.</p>
{% endif %}
{% endif %}

{% if r.quirk_count > 0 %}
<div class="quirk-box">
  <h3>Discovered quirks</h3>
  <ul>
  {% for ek in engine_order %}
    {% for c in r.engines[ek].cases %}
      {% if c.matched_expected == false %}
      <li>
        <span class="formula">{{ c.formula_display or c.formula }}</span> on
        <strong>{{ r.engines[ek].label }}</strong> returned
        <span class="formula">{{ c.value|fmtval }}</span>, but the documented/expected
        result is <span class="formula">{{ c.expected|fmtval }}</span>.
        {% if c.notes %}{{ c.notes }}{% endif %}
      </li>
      {% endif %}
    {% endfor %}
  {% endfor %}
  </ul>
</div>
{% endif %}

{% if r.has_tests %}
<h2 class="section-title">Executed test cases</h2>
{% for ek in engine_order %}
{% set e = r.engines[ek] %}
{% if e.tested %}
<h3>{{ e.label }} {{ e.version }} <span class="category-tag">(tested {{ e.generated_at|dateonly }})</span></h3>
<div class="table-scroll">
<table class="cases">
<thead><tr><th>Formula</th><th>Description</th><th>Result</th><th>Expected</th><th>Verdict</th></tr></thead>
<tbody>
{% for c in e.cases %}
<tr>
  <td class="formula mono">{{ c.formula_display or c.formula }}</td>
  <td>{{ c.description }}</td>
  <td class="result mono">{{ (c.range_values if c.range_values else c.value)|fmtval }}</td>
  <td class="result mono">{{ c.expected|fmtval }}{% if c.expected_note %}<br><span class="category-tag">{{ c.expected_note }}</span>{% endif %}</td>
  <td>{% if c.matched_expected %}<span class="verdict-ok">Matched</span>{% elif c.matched_expected is none and c.expected is none %}{% if c.error %}<span class="verdict-bad">Error</span>{% else %}<span class="verdict-ok">Ran OK</span>{% endif %}{% else %}<span class="verdict-bad">Mismatch</span>{% endif %}</td>
</tr>
{% endfor %}
</tbody>
</table>
</div>
{% endif %}
{% endfor %}
{% else %}
<h2 class="section-title">Test cases</h2>
<p>No test cases have been authored for {{ r.name }} yet. This function's
entry above reflects documentation inventory only.</p>
{% endif %}

<h2 class="section-title">Docs &amp; syntax</h2>
<ul>
{% for ek in engine_order %}
{% set e = r.engines[ek] %}
{% if e.doc_url %}<li>{{ e.label }}: <a href="{{ e.doc_url }}">official documentation</a></li>{% endif %}
{% endfor %}
</ul>

{% if related_recipes %}
<h2 class="section-title">Related how-to recipes</h2>
<ul>
{% for rec in related_recipes %}<li><a href="{{ rel }}how-to/{{ rec.slug }}.html">{{ rec.title }}</a></li>{% endfor %}
</ul>
{% endif %}

{% if related_comparisons %}
<h2 class="section-title">Compared against other functions</h2>
<ul>
{% for c in related_comparisons %}<li><a href="{{ rel }}compare/{{ c.slug }}.html">{{ c.title }}</a></li>{% endfor %}
</ul>
{% endif %}
{% endblock %}
"""

QUIRKS_TMPL = """{% extends "base.html" %}
{% block content %}
<h1>Discovered quirks</h1>
<p class="tagline">Every case below is a real, executed formula whose result did
not match documented/expected behavior. This is the flagship content of
{{ site_name }}: cross-engine divergence that only shows up when you actually
run the formula.</p>
<p class="search-hint">{{ quirks|length }} quirks found across {{ quirk_fn_count }} functions.</p>

<ul class="quirks-list">
{% for q in quirks %}
<li class="quirk-entry">
  <h2 class="quirk-h"><a href="{{ rel }}functions/{{ q.name_lower }}.html">{{ q.function }}</a>
  <span class="badge badge-quirk">{{ q.engine_label }}</span></h2>
  <div class="formula mono">{{ q.case.formula_display or q.case.formula }}</div>
  <dl class="quirk-grid">
    <dt>Actual result</dt><dd class="mono">{{ (q.case.range_values if q.case.range_values else q.case.value)|fmtval }}</dd>
    <dt>Documented / expected</dt><dd class="mono">{{ q.case.expected|fmtval }}</dd>
    <dt>Engine</dt><dd>{{ q.engine_label }} {{ q.engine_version }}</dd>
    <dt>Category</dt><dd>{{ q.category }}</dd>
  </dl>
  {% if q.case.notes %}<p>{{ q.case.notes }}</p>{% endif %}
</li>
{% endfor %}
</ul>
{% if seo_guides %}
<h2 class="section-title" style="margin-top:2rem">Deep dives: formulas that behave differently</h2>
<ul>
{% for g in seo_guides %}<li><a href="{{ rel }}guides/{{ g.slug }}.html">{{ g.title }}</a></li>
{% endfor %}</ul>
{% endif %}
{% endblock %}
"""

SITEMAP_TMPL = """<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
{% for u in urls %}  <url><loc>{{ u.loc }}</loc><lastmod>{{ u.lastmod }}</lastmod></url>
{% endfor %}</urlset>
"""


def dateonly_filter(iso_str):
    return iso_date(iso_str) if iso_str else ""


def fmtval_filter(v):
    """Render a raw JSON value (scalar, error string, or possibly-nested list
    from a spill/array result) as a compact, readable literal."""
    if v is None:
        return ""
    if isinstance(v, (list, tuple)):
        return "{" + ", ".join(fmtval_filter(x) for x in v) + "}"
    return str(v)


RECIPE_INDEX_TMPL = """{% extends "base.html" %}
{% block content %}
<h1>Spreadsheet how-to recipes</h1>
<p class="lede">{{ recipes|length }} common spreadsheet tasks with copy-paste formulas for Microsoft Excel, Google Sheets, and LibreOffice Calc &mdash; each one <strong>executed and verified in a real engine</strong>, not just documented.</p>
<p>
{% for cat, items in grouped %}<a href="#{{ cat|lower|replace(' ','-')|replace('&','and') }}">{{ cat }}</a> ({{ items|length }}){% if not loop.last %} &middot; {% endif %}{% endfor %}
</p>
{% for cat, items in grouped %}
<h2 class="section-title" id="{{ cat|lower|replace(' ','-')|replace('&','and') }}">{{ cat }}</h2>
<ul class="recipe-list">
{% for r in items %}
<li><a href="{{ rel }}how-to/{{ r.slug }}.html">{{ r.title }}</a>{% if r.verified %} <span class="badge badge-good">verified</span>{% endif %}</li>
{% endfor %}
</ul>
{% endfor %}
{% endblock %}
"""


# Topical grouping for the how-to index. Order = display order; first match wins.
_RECIPE_CATEGORIES = [
    ("Dates & times", ("date", "day", "month", "week", "year", "quarter", "age",
                       "time", "hour", "minute", "second", "birthdate", "countdown")),
    ("Text & names", ("text", "name", "word", "character", "letter", "space",
                      "case", "concatenate", "combine-first", "split", "extract",
                      "initials", "line-break", "domain", "trim", "capitalize",
                      "leading-zeros", "combine-cells", "left-mid")),
    ("Lookups & filters", ("lookup", "match", "filter", "find-the", "find-values",
                           "duplicate", "unique", "blank-cells-from", "both-lists",
                           "all-matches", "transpose", "reverse", "list", "nth-largest",
                           "reference", "compare", "sort", "fill-blank", "position")),
    ("Counting & conditions", ("count", "if-", "-if", "contains", "rank",
                               "occurrences", "condition", "frequency")),
    ("Formatting & display", ("abbreviate", "fraction", "-dash", "plus-sign",
                              "scientific", "instead-of-zero")),
    ("Math, money & stats", ("sum", "average", "percentage", "round", "median",
                             "deviation", "interest", "loan", "cagr", "margin",
                             "tax", "discount", "price", "weighted", "random",
                             "interpolation", "multiply", "subtract", "convert",
                             "clamp", "mode", "frequent", "largest", "total",
                             "ratio", "z-score", "break-even", "absolute",
                             "normalize", "commission", "running", "difference",
                             "auto-number", "max-minus", "minimum",
                             "irr", "future-value", "investment",
                             "correlation", "percentile",
                             "coefficient", "variation", "geometric")),
]


def _recipe_category(slug):
    for cat, keys in _RECIPE_CATEGORIES:
        if any(k in slug for k in keys):
            return cat
    return "More tasks"


_FUNC_CALL_RE = re.compile(r"([A-Za-z_][A-Za-z0-9_.]*)\s*\(")


def _functions_used(recipe, by_name):
    """Extract the inventoried functions a recipe's formulas call, so each recipe
    can link to those function compatibility pages (contextual internal links that
    pass authority to the money pages, and let readers jump to full compat info)."""
    seen = {}
    for app in ("excel", "google_sheets", "libreoffice"):
        s = (recipe.get("solutions") or {}).get(app)
        if not s:
            continue
        # drop quoted string literals so text like "(sales)" can't match
        formula = re.sub(r'"[^"]*"', "", s.get("formula", ""))
        for tok in _FUNC_CALL_RE.findall(formula):
            name = tok.upper()
            if name in by_name and name not in seen:
                seen[name] = by_name[name]["name_lower"]
    return [{"name": n, "name_lower": nl} for n, nl in seen.items()]

RECIPE_TMPL = """{% extends "base.html" %}
{% block content %}
<a class="back-link" href="{{ rel }}how-to/">&larr; All how-to recipes</a>
<div class="func-header">
  <h1>{{ r.title }}</h1>
  {% if r.verified %}<span class="badge badge-good">&#10003; Verified in LibreOffice {{ r.engine_version }}</span>{% endif %}
</div>
<p class="lede">{{ r.task }}</p>

<h2 class="section-title">The formula</h2>
<div class="table-scroll">
<table class="matrix">
<thead><tr><th>App</th><th>Formula</th><th>Notes</th></tr></thead>
<tbody>
{% for app in app_order %}{% set s = r.solutions.get(app) %}{% if s %}
<tr><td>{{ app_labels[app] }}</td><td><code>{{ s.formula }}</code></td><td>{{ s.note }}</td></tr>
{% endif %}{% endfor %}
</tbody>
</table>
</div>

<h2 class="section-title">How it works</h2>
<p>{{ r.explanation }}</p>

{% if r.verified %}
<h2 class="section-title">Verified, not just documented</h2>
<p>We ran <code>{{ r.example_formula }}</code> in LibreOffice {{ r.engine_version }} (headless, with forced recalculation) and it returned <code>{{ r.example_actual }}</code> &mdash; exactly the expected result. Every formula here is confirmed by actually executing it.</p>
{% endif %}
{% if functions_used %}
<h2 class="section-title">Functions used</h2>
<p>{% for f in functions_used %}<a href="{{ rel }}functions/{{ f.name_lower }}.html">{{ f.name }}</a>{% if not loop.last %} &middot; {% endif %}{% endfor %} &mdash; see full Excel, Google Sheets &amp; LibreOffice compatibility for each.</p>
{% endif %}
{% if related_recipes %}
<h2 class="section-title">Related recipes</h2>
<ul>
{% for rec in related_recipes %}<li><a href="{{ rel }}how-to/{{ rec.slug }}.html">{{ rec.title }}</a></li>{% endfor %}
</ul>
{% endif %}
{% if related_comparisons %}
<h2 class="section-title">Related comparisons</h2>
<ul>
{% for c in related_comparisons %}<li><a href="{{ rel }}compare/{{ c.slug }}.html">{{ c.title }}</a></li>{% endfor %}
</ul>
{% endif %}
{% endblock %}
"""


COMPARISON_TMPL = """{% extends "base.html" %}
{% block content %}
<a class="back-link" href="{{ rel }}compare/">&larr; All comparisons</a>
<div class="func-header">
  <h1>{{ c.title }}</h1>
</div>
<p class="lede">{{ c.intro }}</p>

<h2 class="section-title">The differences at a glance</h2>
<div class="table-scroll">
<table class="matrix">
<thead><tr><th></th>{% for f in c.funcs %}<th><a href="{{ rel }}functions/{{ f|lower|replace('/','-') }}.html">{{ f }}</a></th>{% endfor %}</tr></thead>
<tbody>
{% for row in c.table %}
<tr><td><strong>{{ row.aspect }}</strong></td>{% for col in row.cols %}<td>{{ col }}</td>{% endfor %}</tr>
{% endfor %}
</tbody>
</table>
</div>

<h2 class="section-title">Which should you use?</h2>
<ul>
{% for w in c.when %}
<li><strong>{{ w.func }}</strong> &mdash; {{ w.use_when }}</li>
{% endfor %}
</ul>

<h2 class="section-title">Compatibility (from executed tests)</h2>
<p>{{ c.compat_note }}</p>

<h2 class="section-title">Example formulas</h2>
<div class="table-scroll">
<table class="matrix">
<tbody>
{% for ex in c.examples %}
<tr><td>{{ ex.label }}</td><td><code>{{ ex.formula }}</code></td></tr>
{% endfor %}
</tbody>
</table>
</div>

<p>Full per-version details on each function page:
{% for f in c.funcs %}<a href="{{ rel }}functions/{{ f|lower|replace('/','-') }}.html">{{ f }}</a>{% if not loop.last %} &middot; {% endif %}{% endfor %}.</p>
{% if c.see_also %}
<p>See also: {% for s in c.see_also %}<a href="{{ rel }}{{ s.href }}">{{ s.label }}</a>{% if not loop.last %} &middot; {% endif %}{% endfor %}.</p>
{% endif %}
{% if related_recipes %}
<h2 class="section-title">How-to recipes using these functions</h2>
<ul>
{% for rec in related_recipes %}<li><a href="{{ rel }}how-to/{{ rec.slug }}.html">{{ rec.title }}</a></li>{% endfor %}
</ul>
{% endif %}
{% endblock %}
"""


COMPARISON_INDEX_TMPL = """{% extends "base.html" %}
{% block content %}
<h1>Spreadsheet function comparisons</h1>
<p class="lede">Head-to-head guides for the functions people mix up &mdash; what actually differs, which to use when, and how support varies across Excel, Google Sheets, and LibreOffice (backed by executed tests, not documentation). For the app-level picture, see the <a href="{{ rel }}excel-vs-google-sheets.html">Excel vs Google Sheets formula guide</a>.</p>
<ul class="quirks-list">
{% for c in comparisons %}
<li class="quirk-entry">
  <h3><a href="{{ rel }}compare/{{ c.slug }}.html">{{ c.title }}</a></h3>
  <p style="margin:.3rem 0 0">{{ c.intro|truncate(160) }}</p>
</li>
{% endfor %}
</ul>
{% endblock %}
"""


ERRORS_TMPL = """{% extends "base.html" %}
{% block content %}
<h1>Spreadsheet error values, explained</h1>
<p class="lede">Every hash-error is the engine telling you something specific. Here&rsquo;s what each one means, the usual causes ranked by likelihood, and the fastest fix &mdash; for Excel, Google Sheets, and LibreOffice Calc.</p>

<h2 class="section-title" id="name">#NAME? &mdash; unrecognized name</h2>
<p>The engine doesn&rsquo;t know a function or name in the formula. In order of likelihood:</p>
<ul>
<li><strong>The function doesn&rsquo;t exist in this app or version.</strong> XLOOKUP in Excel 2019, QUERY anywhere outside Sheets, MAP in LibreOffice. Paste the formula into the <a href="{{ rel }}checker.html">compatibility checker</a> to see exactly which function fails where, and check the <a href="{{ rel }}libreoffice-version-support.html">LibreOffice version page</a> &mdash; upgrading often IS the fix (XLOOKUP needs LO 24.8+).</li>
<li><strong>A typo</strong> &mdash; =SUMIFF(...), =VLOOKUPP(...).</li>
<li><strong>Text without quotes</strong> &mdash; =IF(A2=yes,...) reads yes as a name; it needs \"yes\".</li>
<li><strong>Generated files missing the storage prefix.</strong> If a Python/library-generated .xlsx shows #NAME? on modern functions in every app, it&rsquo;s the OOXML <code>_xlfn.</code> prefix issue &mdash; explained in our <a href="{{ rel }}methodology.html">methodology</a>.</li>
</ul>

<h2 class="section-title" id="ref">#REF! &mdash; broken reference</h2>
<p>The formula points at cells that no longer exist: rows/columns deleted, a sheet removed, or a copied formula whose relative references walked off the edge of the grid. VLOOKUP with a column index bigger than its table is the classic. Undo is your friend; longer-term, INDEX/MATCH and whole-range references survive edits that hard-coded positions don&rsquo;t.</p>

<h2 class="section-title" id="value">#VALUE! &mdash; wrong type of input</h2>
<p>A function got text where it needed a number or date: arithmetic on cells containing text (often invisible &mdash; see the <a href="{{ rel }}compare/trim-vs-clean.html">TRIM vs CLEAN guide</a>), dates stored as text (<a href="{{ rel }}how-to/convert-text-to-date.html">convert them</a>), or FIND/SEARCH not finding its target. In legacy Excel it&rsquo;s also un-entered array formulas that needed Ctrl+Shift+Enter.</p>

<h2 class="section-title" id="div0">#DIV/0! &mdash; division by zero</h2>
<p>The denominator is zero or blank. Averages of empty ranges throw it too (AVERAGEIF with no matches). Guard the specific case &mdash; =IF(B2=0,\"\",A2/B2) &mdash; rather than blanket-wrapping in IFERROR, which also hides real bugs (<a href="{{ rel }}compare/iferror-vs-ifna.html">why that matters</a>).</p>

<h2 class="section-title" id="na">#N/A &mdash; not found (usually not an error)</h2>
<p>Lookup functions return it when the value isn&rsquo;t there &mdash; it&rsquo;s a legitimate answer, not breakage. Handle the expected miss with XLOOKUP&rsquo;s 4th argument or IFNA; investigate only when everything comes back #N/A (usually a type mismatch: numbers vs text-numbers, or stray spaces &mdash; the <a href="{{ rel }}compare/isblank-vs-empty-string.html">two-kinds-of-empty</a> and TRIM issues).</p>

<h2 class="section-title" id="spill">#SPILL! / blocked arrays</h2>
<p>A dynamic-array formula (FILTER, SORT, UNIQUE, SEQUENCE&hellip;) needs room to spill and something occupies the target cells. Clear the blocking cells &mdash; the error message highlights them in Excel. Google Sheets says #REF! with a &ldquo;result was not expanded&rdquo; note instead; LibreOffice reports its own error code. Legacy note: in pre-dynamic-array versions these functions don&rsquo;t exist at all (#NAME? instead) &mdash; see <a href="{{ rel }}libreoffice-version-support.html">which versions have them</a>.</p>

<h2 class="section-title" id="num">#NUM! &mdash; impossible number</h2>
<p>Math that can&rsquo;t produce a representable result: SQRT of a negative, IRR that doesn&rsquo;t converge (add a guess argument), dates before the epoch, or numbers beyond ~1E308. Usually the inputs are wrong, not the formula.</p>

<h2 class="section-title" id="tools">Tools for diagnosing</h2>
<ul>
<li><code>=ERROR.TYPE(A2)</code> returns a code per error kind (7 = #N/A) &mdash; see <a href="{{ rel }}compare/iserror-vs-iserr-vs-isna.html">the IS-function guide</a> for detector strategies.</li>
<li>The <a href="{{ rel }}checker.html">formula checker</a> answers &ldquo;is this #NAME? a version problem?&rdquo; instantly.</li>
<li>Our <a href="{{ rel }}quirks.html">quirks catalog</a> lists cases where an app returns a DIFFERENT error than Excel for the same formula &mdash; found by executing them.</li>
</ul>
{% endblock %}
"""


DATASET_TMPL = """{% extends "base.html" %}
{% block content %}
<h1>Open spreadsheet compatibility dataset</h1>
<p class="lede">The machine-verified data behind this site is free to use. It records, for {{ n_funcs }} spreadsheet functions, whether each works in Microsoft Excel and Google Sheets (from official documentation) and in LibreOffice Calc (from <a href="{{ rel }}methodology.html">actually executing the formula</a>, with per-version history). As far as we know it&rsquo;s the only openly available <em>executed</em> cross-application compatibility dataset.</p>

<h2 class="section-title">Download</h2>
<p><a href="{{ rel }}data/compat.json"><code>data/compat.json</code></a> &mdash; one JSON object, keyed by uppercase function name ({{ n_funcs }} entries, {{ kb }} KB).<br>
<a href="{{ rel }}data/compat.csv"><code>data/compat.csv</code></a> &mdash; the same data as a CSV (one row per function, headered columns) for spreadsheets and data tools.</p>
<p>The full test harness, authored test cases, and raw per-LibreOffice-version results are in the <a href="{{ github_url }}">GitHub repository</a>.</p>

<h2 class="section-title">Schema</h2>
<div class="table-scroll">
<table class="matrix">
<thead><tr><th>Field</th><th>Type</th><th>Meaning</th></tr></thead>
<tbody>
<tr><td><code>cat</code></td><td>string</td><td>Function category (e.g. &ldquo;Lookup and reference&rdquo;).</td></tr>
<tr><td><code>x</code></td><td>boolean</td><td>Documented in Microsoft Excel.</td></tr>
<tr><td><code>g</code></td><td>boolean</td><td>Documented in Google Sheets.</td></tr>
<tr><td><code>l</code></td><td>boolean</td><td>Documented in LibreOffice Calc.</td></tr>
<tr><td><code>lv</code></td><td>string / null</td><td>LibreOffice <strong>executed</strong> verdict: <code>supported</code>, <code>quirky</code>, <code>unsupported</code>, or null when not yet live-tested.</td></tr>
<tr><td><code>lver</code></td><td>string</td><td>LibreOffice version the verdict was produced on (e.g. <code>25.8.7.3</code>).</td></tr>
<tr><td><code>lnew</code></td><td>string / null</td><td>The LibreOffice version the function first became supported in, when known (else null).</td></tr>
</tbody>
</table>
</div>

<h2 class="section-title">Example</h2>
<div class="table-scroll">
<pre style="background:var(--bg-alt,#f6f8fa);padding:1rem;border-radius:8px;overflow:auto"><code>// fetch the dataset
const db = await (await fetch("https://canispreadsheet.com/data/compat.json")).json();

db["XLOOKUP"]
// {"cat":"Lookup and reference","x":true,"g":true,"l":true,
//  "lv":"supported","lver":"25.8.7.3","lnew":"24.8.7.2"}
//  -> documented in all three; executed as supported in LibreOffice,
//     first working in LibreOffice 24.8.</code></pre>
</div>

<h2 class="section-title">License</h2>
<p>The compatibility dataset is released under <a href="https://creativecommons.org/licenses/by/4.0/" rel="license">Creative Commons Attribution 4.0 (CC&nbsp;BY&nbsp;4.0)</a>. Use it freely, including commercially &mdash; just credit <strong>canispreadsheet.com</strong> with a link. If you build something with it, we&rsquo;d love to hear about it.</p>
<p style="font-size:.9em;color:var(--text-muted,#6b7280)">The data reflects executed tests on the LibreOffice versions noted and each vendor&rsquo;s published function documentation at the time of testing; it is provided as-is, without warranty. Corrections welcome via the <a href="{{ github_url }}">repository</a>.</p>
{% endblock %}
"""


EQUIV_TMPL = """{% extends "base.html" %}
{% block content %}
<h1>Excel &harr; Google Sheets function equivalents</h1>
<p class="lede">Moving a formula between Excel and Google Sheets? Most functions are <strong>identical</strong> &mdash; VLOOKUP, SUMIFS, INDEX/MATCH, IF, the whole everyday toolkit works the same name, same arguments. This page covers the ones that <em>don&rsquo;t</em>, in both directions, with the verified replacement to use instead.</p>
<p>Need it for one specific formula? Paste it into the <a href="{{ rel }}checker.html">compatibility checker</a> and pick a target app for an instant migration report.</p>

<h2 class="section-title">Only in Google Sheets &mdash; what to use in Excel</h2>
<div class="table-scroll">
<table class="matrix">
<thead><tr><th>Google Sheets</th><th>Excel equivalent</th></tr></thead>
<tbody>
{% for r in g_only %}
<tr><td>{% if r.exists %}<a href="{{ rel }}functions/{{ r.fn|lower }}.html"><code>{{ r.fn }}</code></a>{% else %}<code>{{ r.fn }}</code>{% endif %}</td><td>{{ r.note }}</td></tr>
{% endfor %}
</tbody>
</table>
</div>
<p style="font-size:.9em;color:var(--text-muted,#6b7280)">Full list: <a href="{{ rel }}sheets-functions-not-in-excel.html">all Google Sheets functions not in Excel</a>.</p>

<h2 class="section-title">Only in Excel &mdash; what to use in Google Sheets</h2>
<div class="table-scroll">
<table class="matrix">
<thead><tr><th>Excel</th><th>Google Sheets equivalent</th></tr></thead>
<tbody>
{% for r in x_only %}
<tr><td>{% if r.exists %}<a href="{{ rel }}functions/{{ r.fn|lower }}.html"><code>{{ r.fn }}</code></a>{% else %}<code>{{ r.fn }}</code>{% endif %}</td><td>{{ r.note }}</td></tr>
{% endfor %}
</tbody>
</table>
</div>
<p style="font-size:.9em;color:var(--text-muted,#6b7280)">Full list: <a href="{{ rel }}excel-functions-not-in-google-sheets.html">all Excel functions not in Google Sheets</a>.</p>

<h2 class="section-title">Same name, watch the difference</h2>
<div class="table-scroll">
<table class="matrix">
<thead><tr><th>Function</th><th>What differs between Excel and Sheets</th></tr></thead>
<tbody>
{% for r in gotcha %}
<tr><td><code>{{ r.fn }}</code></td><td>{{ r.note }}</td></tr>
{% endfor %}
</tbody>
</table>
</div>

<h2 class="section-title">A note on versions</h2>
<p>Some &ldquo;differences&rdquo; are really just version age. XLOOKUP, FILTER, SORT, LET, LAMBDA and the dynamic-array family need <strong>Excel 2021 or 365</strong> &mdash; current Google Sheets has had them for years, so a modern Sheet can be <em>more</em> compatible with these than an Excel 2019 install. Full support-by-version detail (including LibreOffice) is on each <a href="{{ rel }}index.html">function page</a>, and the app-level picture is in the <a href="{{ rel }}excel-vs-google-sheets.html">Excel vs Google Sheets guide</a>.</p>
{% endblock %}
"""


PILLAR_XVG_TMPL = """{% extends "base.html" %}
{% block content %}
<h1>Excel vs Google Sheets: the formula compatibility guide</h1>
<p class="lede">Most Excel-vs-Sheets comparisons argue about collaboration and price. This one covers the part that silently breaks when you switch: <strong>the formulas</strong> &mdash; which functions exist on each side, where the same function behaves differently, and how to keep a workbook portable.</p>

<h2 class="section-title">The numbers</h2>
<ul>
<li><strong>{{ n_both }}</strong> functions are documented for BOTH Excel and Google Sheets &mdash; the shared core where most everyday work lives.</li>
<li><strong>{{ n_xonly }}</strong> functions are Excel-only (<a href="{{ rel }}excel-functions-not-in-google-sheets.html">full list</a>): PIVOTBY, GROUPBY, the CUBE family, AGGREGATE&hellip; Importing an .xlsx that uses them leaves <code>#NAME?</code> cells in Sheets.</li>
<li><strong>{{ n_gonly }}</strong> functions are Sheets-only (<a href="{{ rel }}sheets-functions-not-in-excel.html">full list</a>): QUERY, ARRAYFORMULA, IMPORTRANGE, GOOGLEFINANCE, the REGEX family&hellip; These die on export to Excel.</li>
</ul>

<h2 class="section-title">Same function, different dialect</h2>
<p>The subtler traps are functions both apps HAVE but spell differently:</p>
<ul>
<li><strong>Array formulas.</strong> Sheets wraps per-row math in <code>ARRAYFORMULA(...)</code>; modern Excel just spills array expressions natively. Exporting a Sheet frequently degrades wrapped arithmetic to a single-cell result &mdash; details in <a href="{{ rel }}compare/arrayformula-vs-dynamic-arrays.html">ARRAYFORMULA vs dynamic arrays</a>.</li>
<li><strong>Filtering with SQL.</strong> Sheets&rsquo; <code>QUERY</code> has no Excel equivalent at all; the portable subset is FILTER/SORT/UNIQUE &mdash; see <a href="{{ rel }}compare/filter-vs-query.html">FILTER vs QUERY</a>.</li>
<li><strong>Regular expressions.</strong> REGEXMATCH/REGEXEXTRACT/REGEXREPLACE are Sheets-only; Excel&rsquo;s closest tools are SEARCH wildcards and TEXTBEFORE/TEXTAFTER.</li>
<li><strong>Version skew inside Excel itself.</strong> XLOOKUP, FILTER, LAMBDA and friends need Excel 2021+/365 &mdash; a current Sheet can be MORE compatible with modern functions than a 2019 Excel install.</li>
</ul>

<h2 class="section-title">Keeping a workbook portable</h2>
<ol>
<li>Stick to the shared core &mdash; paste any formula into the <a href="{{ rel }}checker.html">compatibility checker</a> and it flags every function per app.</li>
<li>Prefer FILTER/SORT/UNIQUE compositions over QUERY, and native ranges over ARRAYFORMULA, when a file might ever leave Sheets.</li>
<li>Avoid the newest Excel exclusives (GROUPBY, PIVOTBY) in files that co-workers will open in Sheets.</li>
<li>Test the round trip: export, reopen, and search for <code>#NAME?</code> and <code>#REF!</code>.</li>
</ol>

<h2 class="section-title">Where LibreOffice fits</h2>
<p>LibreOffice Calc reads both formats and &mdash; unlike either vendor&rsquo;s docs &mdash; we can actually EXECUTE formulas in it: {{ n_tested }} functions live-tested across {{ versions|join(', ') }}. Its modern-function support now tracks Excel closely (XLOOKUP since 24.8, the dynamic-array batch since 25.8 &mdash; see <a href="{{ rel }}libreoffice-version-support.html">support by version</a>), and several &ldquo;Excel-only&rdquo; functions like AGGREGATE work fine there, making it a viable escape hatch for files Sheets can&rsquo;t fully open.</p>

<p>Every claim here is backed by the underlying pages: <a href="{{ rel }}index.html">per-function verdicts</a>, <a href="{{ rel }}quirks.html">behavioral quirks</a>, the <a href="{{ rel }}excel-google-sheets-equivalents.html">function equivalents table</a>, and the <a href="{{ rel }}methodology.html">verification methodology</a>.</p>
{% endblock %}
"""


EXCLUSIVE_TMPL = """{% extends "base.html" %}
{% block content %}
<h1>{{ heading }}</h1>
<p class="lede">{{ lede }}</p>
<p>{{ items|length }} functions, from each vendor&rsquo;s official documentation{% if show_lo %} &mdash; with LibreOffice&rsquo;s status from our executed tests where available{% endif %}. Formulas using these <strong>break with <code>#NAME?</code></strong> when a file moves to the other app.</p>
<div class="table-scroll">
<table class="matrix">
<thead><tr><th>Function</th><th>Category</th>{% if show_lo %}<th>Also in LibreOffice?</th>{% endif %}</tr></thead>
<tbody>
{% for it in items %}
<tr>
  <td><a href="{{ rel }}functions/{{ it.name_lower }}.html">{{ it.name }}</a></td>
  <td>{{ it.category }}</td>
  {% if show_lo %}<td>{% if it.lo_verdict == 'supported' %}<span class="badge badge-good">Yes (tested)</span>{% elif it.lo_verdict == 'quirky' %}<span class="badge badge-quirk">Yes, with quirks</span>{% elif it.lo_verdict == 'unsupported' %}<span class="badge badge-bad">No (tested)</span>{% elif it.lo_doc %}<span class="badge badge-unknown">Documented</span>{% else %}<span class="badge badge-unknown">No</span>{% endif %}</td>{% endif %}
</tr>
{% endfor %}
</tbody>
</table>
</div>
<p>{{ outro }}</p>
<p>See the <a href="{{ rel }}excel-google-sheets-equivalents.html">Excel &harr; Google Sheets equivalents table</a> for the verified replacement to use for each. Part of the <a href="{{ rel }}excel-vs-google-sheets.html">Excel vs Google Sheets formula compatibility guide</a>.</p>
{% endblock %}
"""


METHODOLOGY_TMPL = """{% extends "base.html" %}
{% block content %}
<h1>How we verify spreadsheet compatibility</h1>
<p class="lede">Every LibreOffice verdict on this site comes from <strong>actually executing the formula</strong> in a real engine and checking what came back &mdash; not from reading documentation. This page explains the machinery, what &ldquo;verified&rdquo; means here, and the limits of the approach.</p>

<h2 class="section-title">The execution harness</h2>
<p>For each function we author test cases: a formula, any setup cells it needs, and the result Excel documents or produces for that input. The harness writes each case into a real <code>.xlsx</code> workbook with openpyxl, then runs <strong>headless LibreOffice Calc</strong> over it (<code>soffice --convert-to xlsx</code>), which forces a full recalculation. We reload the output and compare every result against the expected value.</p>
<p>Two guards make the results trustworthy:</p>
<ul>
<li><strong>Recalculation canaries.</strong> Every generated workbook contains sentinel formulas (deterministic arithmetic plus a volatile function) whose values prove the engine really recalculated rather than echoing stored results. A run that fails its canary is discarded, never published.</li>
<li><strong>The OOXML <code>_xlfn.</code> storage prefix.</strong> Functions added to Excel after 2007 are not stored in <code>.xlsx</code> files under their display names &mdash; Excel silently writes <code>_xlfn.XLOOKUP(...)</code> and strips the prefix for display. Libraries that write raw XML don&rsquo;t do this for you: write <code>=XLOOKUP(...)</code> verbatim with openpyxl and <em>every</em> engine, including ones that fully support XLOOKUP, shows <code>#NAME?</code>. The harness translates display formulas to correct storage form before writing, so a <code>#NAME?</code> in our results means the engine genuinely lacks the function &mdash; not a file-format artifact. (If you&rsquo;ve ever generated a spreadsheet from Python and hit an inexplicable <code>#NAME?</code> on a modern function, this prefix is almost certainly why.)</li>
</ul>

<h2 class="section-title">Version matrix</h2>
<p>The same corpus runs against multiple LibreOffice releases &mdash; currently {{ versions|join(', ') }} &mdash; which is how function pages can state a precise &ldquo;supported since&rdquo; release rather than a guess. Current corpus: <strong>{{ n_funcs }} functions live-tested</strong> across <strong>{{ n_cases }} executed cases</strong> per release.</p>

<h2 class="section-title">What the verdicts mean</h2>
<ul>
<li><strong>Supported</strong> &mdash; every executed case matched the Excel-canonical expected result (probe cases for volatile functions like NOW/RAND assert error-free execution and deterministic invariants instead of exact values).</li>
<li><strong>Quirk found</strong> &mdash; the function exists but at least one case returned a different value or error than Excel produces; the failing case is shown on the function&rsquo;s page.</li>
<li><strong>Unsupported</strong> &mdash; the engine returns <code>#NAME?</code> (unrecognized function) with the storage prefix correctly applied.</li>
</ul>

<h2 class="section-title">How-to recipes are verified too</h2>
<p>Every formula on a <a href="{{ rel }}how-to/">how-to recipe page</a> runs through the same pipeline before publishing: the exact formula shown is executed in LibreOffice {{ current_version }} with the sample data shown, and the page displays the value it actually returned.</p>

<h2 class="section-title">Honest limitations</h2>
<ul>
<li><strong>Excel and Google Sheets are not live-executed.</strong> Their columns reflect each vendor&rsquo;s official function documentation. We can&rsquo;t headlessly run those engines (yet); where our executed LibreOffice results reveal a difference against documented Excel behavior, that&rsquo;s labeled a quirk of LibreOffice, and disputed cases are re-checked by hand.</li>
<li><strong>Coverage is partial.</strong> {{ n_funcs }} of ~600 catalog functions have executed tests; untested functions say so explicitly rather than borrowing a verdict.</li>
<li><strong>A passing case is evidence, not proof.</strong> A function can match on our cases and still differ on inputs we haven&rsquo;t authored. When you find such an edge, please report it.</li>
</ul>

<h2 class="section-title">Reproduce or dispute a result</h2>
<p>The entire harness, test corpus, and raw per-version results are public in the <a href="{{ github_url }}">GitHub repository</a>. Every function page shows the exact formula, inputs, and returned value for each case, so any result can be reproduced in a few minutes &mdash; and if an engine disagrees with us on your machine, an issue with your version number is the fastest way to get it fixed.</p>
{% endblock %}
"""


CHECKER_TMPL = """{% extends "base.html" %}
{% block content %}
<h1>Spreadsheet formula compatibility checker</h1>
<p class="lede">Paste a formula and see whether every function works in Microsoft Excel, Google Sheets, and current LibreOffice Calc &mdash; based on real executed tests, not just documentation. Pick a target app for a <strong>migration report</strong> that flags what breaks and suggests verified alternatives.</p>
<textarea id="f" rows="3" style="width:100%;box-sizing:border-box;font-family:monospace;font-size:1rem;padding:.6rem" placeholder='=XLOOKUP("North", B2:B6, A2:A6)'></textarea>
<p><button id="btn" class="promo-btn" style="border:0;cursor:pointer">Check compatibility</button></p>
<p style="font-size:.9em;color:var(--text-muted,#6b7280)">Try:
<button data-ex='=XLOOKUP("North",B2:B9,A2:A9,"?")' style="cursor:pointer;margin:0 .2rem">XLOOKUP</button>
<button data-ex='=TEXTJOIN(", ",TRUE,UNIQUE(A2:A99))' style="cursor:pointer;margin:0 .2rem">TEXTJOIN+UNIQUE</button>
<button data-ex='=QUERY(A1:C99,"SELECT A, SUM(C) GROUP BY A",1)' style="cursor:pointer;margin:0 .2rem">QUERY</button>
<button data-ex='=MAP(A2:A9,LAMBDA(x,x*2))' style="cursor:pointer;margin:0 .2rem">MAP/LAMBDA</button>
</p>
<p style="font-size:.9em;color:var(--text-muted,#6b7280);margin:.3rem 0 0">Moving this formula to another app? Pick a target for a migration report:
<label style="margin-left:.4rem"><input type="radio" name="tgt" value="" checked> none</label>
<label style="margin-left:.5rem"><input type="radio" name="tgt" value="x"> Excel</label>
<label style="margin-left:.5rem"><input type="radio" name="tgt" value="g"> Google Sheets</label>
<label style="margin-left:.5rem"><input type="radio" name="tgt" value="l"> LibreOffice</label>
</p>
<div id="out"></div>
<p style="font-size:.9em;color:var(--text-muted,#6b7280)">Results are linkable &mdash; the URL updates with your formula, so you can share a check directly (e.g. in a forum answer).</p>
<script>const DATA_URL="{{ rel }}data/compat.json"; const FUNC_BASE="{{ rel }}functions/"; const CMP_BASE="{{ rel }}compare/";</script>
{% raw %}
<script>
let DB=null;
async function load(){ if(!DB){ DB=await (await fetch(DATA_URL)).json(); } return DB; }
// Function-name extraction. Kept behaviourally identical to the Migration
// Audit's extractFunctions() in site/audit-page/audit.js — the shared case
// list lives in site/audit-page/test-adversarial.mjs, which tests BOTH.
// Quoted content never counts: double quotes are string literals
// (="COUNT(A1)" -> nothing) and single quotes are sheet/workbook names
// (=SUM('My Data (2024)'!A1:A5) -> SUM, not DATA). A name counts only when
// the next non-space character is "(" and no identifier character (incl.
// non-ASCII letters) precedes it. Scanning is token-first rather than
// /NAME\\s*\\(/ because that form backtracks quadratically over a long run of
// letters — an 8k-char pasted formula was enough to stall it.
function splitLits(s){ const segs=[]; let cur='', i=0;
  while(i<s.length){ const ch=s[i];
    if(ch==='"'||ch==="'"){ if(cur){segs.push({lit:false,t:cur}); cur='';} const q=ch; i++;
      while(i<s.length){ if(s[i]===q){ if(s[i+1]===q){i+=2; continue;} i++; break; } i++; }
      segs.push({lit:true,t:''}); }
    else { cur+=ch; i++; } }
  if(cur) segs.push({lit:false,t:cur});
  return segs; }
const FN_TOKEN_RE=/[A-Z_][A-Z0-9_.]*/gi, FN_IDENT_RE=/[A-Za-z0-9_.\\u0080-\\uFFFF]/;
function funcs(s){ const set=new Set();
  for(const seg of splitLits(s===null||s===undefined?'':String(s))){ if(seg.lit) continue;
    const c=seg.t; let m; FN_TOKEN_RE.lastIndex=0;
    while((m=FN_TOKEN_RE.exec(c))!==null){
      if(m.index>0&&FN_IDENT_RE.test(c.charAt(m.index-1))) continue;
      let j=FN_TOKEN_RE.lastIndex; while(j<c.length&&/[ \\t\\r\\n]/.test(c.charAt(j))) j++;
      if(c.charAt(j)!=='(') continue;
      // Excel stores post-2007 functions as _xlfn.NAME / _xlfn._xlws.NAME.
      const fn=m[0].toUpperCase().replace(/^_?XLFN\\./,'').replace(/^_?XLWS\\./,'');
      if(fn) set.add(fn); } }
  return [...set]; }
function yn(ok){ return ok?'<span style="color:#0a7a2f">&#10003; yes</span>':'<span style="color:#c02020">&#10007; no</span>'; }
function lo(d){ const nw=d.lnew?' <span style="color:#0a7a2f;font-size:.85em">(new in '+d.lnew+')</span>':''; if(d.lv==='supported') return '<span style="color:#0a7a2f">&#10003; '+d.lver+'</span>'+nw; if(d.lv==='quirky') return '<span style="color:#b8860b">&#9888; quirk ('+d.lver+')</span>'; if(d.lv==='unsupported') return '<span style="color:#c02020">&#10007; not in '+d.lver+'</span>'; return d.l?'<span style="color:#888">documented</span>':'<span style="color:#c02020">&#10007; no</span>'; }
const TGT_NAME={x:'Excel',g:'Google Sheets',l:'LibreOffice'};
// Curated, verified portable alternatives (from the comparison pages) for
// functions with a genuine cross-app gap. cmp = comparison slug to link, if any.
const MIG={
 QUERY:{r:'Google Sheets only. Rebuild with FILTER (for SELECT/WHERE) plus SUMIF/UNIQUE/SORT for grouping and aggregation.',cmp:'filter-vs-query'},
 ARRAYFORMULA:{r:'Google Sheets only. In Excel 365 and LibreOffice 24.8+, array expressions spill natively — drop the wrapper.',cmp:'arrayformula-vs-dynamic-arrays'},
 REGEXMATCH:{r:'Google Sheets only. No regex functions in Excel/LibreOffice — use SEARCH/FIND with wildcards, or ISNUMBER(SEARCH(...)).'},
 REGEXEXTRACT:{r:'Google Sheets only. Use LEFT/MID/RIGHT with FIND, or TEXTBEFORE/TEXTAFTER (Excel 365 / LO 24.8+).'},
 REGEXREPLACE:{r:'Google Sheets only. Use SUBSTITUTE (by text) or nested SUBSTITUTE for multiple patterns.'},
 GOOGLEFINANCE:{r:'Google Sheets only — no Excel/LibreOffice equivalent for live market data.'},
 IMPORTRANGE:{r:'Google Sheets only. In Excel, link workbooks or use Power Query; LibreOffice uses external references.'},
 IMPORTHTML:{r:'Google Sheets only. In Excel use Power Query (Data → From Web).'},
 IMPORTXML:{r:'Google Sheets only. In Excel use Power Query.'},
 IMPORTDATA:{r:'Google Sheets only. In Excel use Power Query (From Text/CSV).'},
 GROUPBY:{r:'Excel 365 only — not in Google Sheets or LibreOffice yet. Rebuild with a PivotTable, or SUMIFS over UNIQUE keys.'},
 PIVOTBY:{r:'Excel 365 only. Use a PivotTable, or SUMIFS/COUNTIFS over UNIQUE row+column keys.'},
 MAP:{r:'Excel 365 and Google Sheets, but NOT LibreOffice (it has LAMBDA/LET, not the lambda-helpers). Fill a formula down instead.'},
 REDUCE:{r:'Excel 365 and Google Sheets only — not in LibreOffice. Use a running-total helper column.'},
 SCAN:{r:'Excel 365 and Google Sheets only — not in LibreOffice. Use a helper column of cumulative values.'},
 BYROW:{r:'Excel 365 and Google Sheets only — not in LibreOffice. Apply the per-row formula down a column.'},
 BYCOL:{r:'Excel 365 and Google Sheets only — not in LibreOffice.'},
 MAKEARRAY:{r:'Excel 365 and Google Sheets only — not in LibreOffice.'},
 XLOOKUP:{r:'Needs Excel 2021+/365, current Sheets, or LibreOffice 24.8+. In older Excel/LibreOffice use INDEX/MATCH.',cmp:'vlookup-vs-index-match'},
 XMATCH:{r:'Needs Excel 2021+/365, Sheets, or LibreOffice 24.8+. Older versions: use MATCH.'},
 TEXTSPLIT:{r:'Excel 365 & Sheets; LibreOffice 25.8+ only. Older: SUBSTITUTE/MID tricks or Text-to-Columns.'},
 TEXTBEFORE:{r:'Excel 365 & Sheets; LibreOffice 24.8+ (with quirks). Older: LEFT(A,FIND(delim,A)-1).',cmp:'textbefore-textafter-vs-left-mid-right'},
 TEXTAFTER:{r:'Excel 365 & Sheets; LibreOffice 24.8+ (with quirks). Older: MID(A,FIND(delim,A)+1,...).',cmp:'textbefore-textafter-vs-left-mid-right'},
 HSTACK:{r:'Excel 365 & Sheets; LibreOffice 25.8+ only.'},
 VSTACK:{r:'Excel 365 & Sheets; LibreOffice 25.8+ only. In Sheets you can also use {range1;range2}.'},
 TAKE:{r:'Excel 365 & Sheets; LibreOffice 25.8+ only. Older: INDEX ranges.'},
 DROP:{r:'Excel 365 & Sheets; LibreOffice 25.8+ only.'},
 LET:{r:'Excel 365, Sheets, LibreOffice 24.8+. Older versions: inline the repeated expressions.'},
 LAMBDA:{r:'Excel 365, Sheets, LibreOffice 24.8+ (direct-invoke has quirks in LO).'},
 SEQUENCE:{r:'Excel 365, Sheets, LibreOffice 24.8+. Older: ROW(INDIRECT(...)) tricks.'},
 FILTER:{r:'Excel 365, Sheets, LibreOffice 24.8+. Older Excel/LO: SMALL/IF array formulas (Ctrl+Shift+Enter).'},
 SORT:{r:'Excel 365, Sheets, LibreOffice 24.8+. Older: rank-and-INDEX or the menu sort.'},
 UNIQUE:{r:'Excel 365, Sheets, LibreOffice 24.8+. Older: Remove Duplicates, or COUNTIF-based formulas.',cmp:'unique-function-vs-remove-duplicates'}
};
function migrate(fs,db,tg){
  const okIn=(d)=> tg==='l' ? (d.lv?(d.lv!=='unsupported'):d.l) : (tg==='x'?d.x:d.g);
  let blockers=[];
  for(const fn of fs){ const d=db[fn]; if(!d) continue; if(!okIn(d)) blockers.push(fn); }
  let h='<h2 class="section-title" style="margin-top:1.4rem">Migration to '+TGT_NAME[tg]+'</h2>';
  if(!blockers.length){ h+='<p style="color:#0a7a2f;font-weight:600">All recognized functions work in '+TGT_NAME[tg]+'. This formula should port cleanly.</p>'; return h; }
  h+='<p style="font-weight:600;color:#c02020">'+blockers.length+' function'+(blockers.length>1?'s':'')+' need attention:</p><ul>';
  for(const fn of blockers){ const m=MIG[fn];
    h+='<li style="margin-bottom:.5rem"><a href="'+FUNC_BASE+fn.toLowerCase()+'.html"><strong>'+fn+'</strong></a> — '+(m?m.r:'not available in '+TGT_NAME[tg]+' (see its page for details).');
    if(m&&m.cmp) h+=' <a href="'+CMP_BASE+m.cmp+'.html" style="font-size:.9em">compare →</a>';
    h+='</li>';
  }
  h+='</ul><p style="font-size:.9em;color:#888">Alternatives are hand-verified from our comparison pages. This report flags what breaks; it does not auto-rewrite your formula.</p>';
  return h;
}
async function check(){
  const db=await load(); const fs=funcs(document.getElementById('f').value); const out=document.getElementById('out');
  if(!fs.length){ out.innerHTML='<p>No functions found. Try a formula like <code>=SUMIF(A:A,"x",B:B)</code>.</p>'; return; }
  let rows='', xAll=true,gAll=true,lAll=true, unknown=[];
  for(const fn of fs){ const d=db[fn]; if(!d){ unknown.push(fn); continue; }
    const lok=d.lv?(d.lv!=='unsupported'):d.l; xAll=xAll&&d.x; gAll=gAll&&d.g; lAll=lAll&&lok;
    rows+='<tr><td><a href="'+FUNC_BASE+fn.toLowerCase()+'.html">'+fn+'</a></td><td>'+yn(d.x)+'</td><td>'+yn(d.g)+'</td><td>'+lo(d)+'</td></tr>'; }
  const say=ok=>ok?'<span style="color:#0a7a2f">works</span>':'<span style="color:#c02020">has an unsupported function</span>';
  let html='<p style="font-weight:600;margin:1rem 0">Excel: '+say(xAll)+' &middot; Google Sheets: '+say(gAll)+' &middot; LibreOffice: '+say(lAll)+'</p>';
  html+='<div class="table-scroll"><table class="matrix"><thead><tr><th>Function</th><th>Excel</th><th>Google Sheets</th><th>LibreOffice</th></tr></thead><tbody>'+rows+'</tbody></table></div>';
  if(unknown.length) html+='<p style="color:#888">Not in our database (may be a name, cell range, or newer function): '+unknown.join(', ')+'</p>';
  const tg=target(); if(tg) html+=migrate(fs,db,tg);
  html+='<p style="font-size:.9em;color:#888">Shareable link: <a href="'+permalink()+'" style="word-break:break-all">'+permalink()+'</a></p>';
  out.innerHTML=html;
}
function target(){ const r=document.querySelector('input[name=tgt]:checked'); return r?r.value:''; }
function permalink(){ const tg=target(); return location.origin+location.pathname+'#f='+encodeURIComponent(document.getElementById('f').value)+(tg?'&t='+tg:''); }
function sync(){ history.replaceState(null,'',permalink()); }
function setAndCheck(v){ document.getElementById('f').value=v; sync(); check(); }
document.getElementById('btn').addEventListener('click',()=>{ sync(); check(); });
document.getElementById('f').addEventListener('keydown',e=>{ if((e.ctrlKey||e.metaKey)&&e.key==='Enter') check(); });
document.querySelectorAll('[data-ex]').forEach(b=>b.addEventListener('click',()=>setAndCheck(b.getAttribute('data-ex'))));
document.querySelectorAll('input[name=tgt]').forEach(r=>r.addEventListener('change',()=>{ if(document.getElementById('f').value.trim()){ sync(); check(); } }));
if(location.hash.startsWith('#f=')){ try{ const p=new URLSearchParams(location.hash.slice(1)); const v=p.get('f'); const t=p.get('t'); if(t){ const r=document.querySelector('input[name=tgt][value="'+t+'"]'); if(r) r.checked=true; } if(v){ document.getElementById('f').value=v; check(); } }catch(e){} }
</script>
{% endraw %}
{% endblock %}
"""

GUIDES_INDEX_TMPL = """{% extends "base.html" %}
{% block content %}
<h1>Formula behavior guides</h1>
<p class="lede">Executed-data writeups of specific cases where Excel, Google Sheets, and LibreOffice Calc give different results for the exact same formula &mdash; found by actually running the formula, not by comparing documentation pages. LibreOffice values shown are <strong>executed</strong> output from our test harness; Excel and Google Sheets values are each vendor&rsquo;s <strong>documented</strong> behavior unless a guide says otherwise. For the shorter, catalog-style version of these findings across every tested function, see the <a href="{{ rel }}quirks.html">quirks list</a>.</p>
<ul class="quirks-list">
{% for g in guides %}
<li class="quirk-entry">
  <h3><a href="{{ rel }}guides/{{ g.slug }}.html">{{ g.title }}</a></h3>
  <p style="margin:.3rem 0 0">{{ g.meta_description }}</p>
  {% if g.functions %}<p style="margin:.3rem 0 0;font-size:.9em;color:var(--text-muted,#6b7280)">Functions: {% for f in g.functions %}<code>{{ f }}</code>{% if not loop.last %}, {% endif %}{% endfor %}</p>{% endif %}
</li>
{% endfor %}
</ul>
{% endblock %}
"""

SEO_PAGE_TMPL = """{% extends "base.html" %}
{% block content %}
<a class="back-link" href="{{ rel }}quirks.html">&larr; All quirks &amp; gotchas</a>
<h1>{{ h1 }}</h1>
{{ body_html | safe }}
{% endblock %}"""

WHATSNEW_TMPL = """{% extends "base.html" %}
{% block content %}
<h1>LibreOffice Calc function support by version</h1>
<p class="lede">Which functions does each LibreOffice Calc release actually support? We ran the
same corpus of test formulas under LibreOffice {{ versions_tested|join(', ') }} and recorded the
real computed results &mdash; so this is machine-verified compatibility, not documentation claims.
<strong>{{ newly_supported|length }} functions</strong> that returned <code>#NAME?</code> in
{{ from_version }} now work in a later release &mdash; and below is the exact version each one
landed in.</p>

{% if newly_supported %}
<h2 class="section-title">Newly supported functions &mdash; and the release each landed in</h2>
<p>These functions were <strong>not recognized</strong> (returned <code>#NAME?</code>) in
LibreOffice {{ from_version }} but became fully supported in a later release. Most are modern
dynamic-array and lookup functions Excel and Google Sheets already had. The
<strong>Supported since</strong> column is the earliest release we tested where the function works.</p>
<div class="table-scroll">
<table class="matrix">
<thead><tr><th>Function</th><th>Category</th>{% for v in versions_tested %}<th>{{ v }}</th>{% endfor %}<th>Supported since</th></tr></thead>
<tbody>
{% for r in newly_supported %}
<tr>
  <td><a href="{{ rel }}functions/{{ r.name_lower }}.html">{{ r.name }}</a></td>
  <td>{{ r.category }}</td>
  {% for vd in r.verdicts %}<td><span class="badge {{ verdict_class[vd] }}">{{ verdict_short[vd] }}</span></td>{% endfor %}
  <td><strong>{{ r.since }}</strong></td>
</tr>
{% endfor %}
</tbody>
</table>
</div>
{% endif %}

{% if other_changes %}
<h2 class="section-title">Other support changes</h2>
<p>Functions whose behaviour changed across these releases in some other way &mdash; for example,
newly recognized but with an edge-case quirk rather than full support.</p>
<div class="table-scroll">
<table class="matrix">
<thead><tr><th>Function</th><th>Category</th>{% for v in versions_tested %}<th>{{ v }}</th>{% endfor %}</tr></thead>
<tbody>
{% for r in other_changes %}
<tr>
  <td><a href="{{ rel }}functions/{{ r.name_lower }}.html">{{ r.name }}</a></td>
  <td>{{ r.category }}</td>
  {% for vd in r.verdicts %}<td><span class="badge {{ verdict_class[vd] }}">{{ verdict_short[vd] }}</span></td>{% endfor %}
</tr>
{% endfor %}
</tbody>
</table>
</div>
{% endif %}

<h2 class="section-title">How we know</h2>
<p>For each LibreOffice release, we build a workbook of test formulas with no cached values,
force a full headless recalculation, and read back the computed results &mdash; with volatile
and arithmetic canaries proving the recalculation genuinely happened. The same method powers
every <a href="{{ rel }}index.html">function page</a> and the
<a href="{{ rel }}checker.html">formula checker</a>. Versions tested so far:
{{ versions_tested|join(', ') }}.</p>

<p>Related cross-app gaps:
<a href="{{ rel }}sheets-functions-not-in-excel.html">Google Sheets functions that don&rsquo;t exist in Excel</a> &middot;
<a href="{{ rel }}excel-functions-not-in-google-sheets.html">Excel functions that don&rsquo;t exist in Google Sheets</a> &middot;
<a href="{{ rel }}methodology.html">how we verify</a>.</p>
{% endblock %}
"""


def load_seo_pages():
    """Long-form 'behaves differently across engines' guides authored as one
    JSON file per page in site/seo-pages/. Each: slug, title, meta_description,
    h1, body_html (raw, pre-rendered HTML using the site's CSS classes)."""
    pages = []
    if SEO_PAGES_DIR.exists():
        for pth in sorted(SEO_PAGES_DIR.glob("*.json")):
            pages.append(json.loads(pth.read_text()))
    return pages


def load_recipes():
    recs = []
    verif = {}
    vpath = RESULTS_DIR / "recipes-verified.json"
    if vpath.exists():
        verif = json.loads(vpath.read_text()).get("recipes", {})
    rdir = DATA_DIR / "recipes"
    if not rdir.exists():
        return recs
    for p in sorted(rdir.glob("*.json")):
        d = json.loads(p.read_text())
        v = verif.get(d["slug"], {})
        act = v.get("actual", "")
        if isinstance(act, list):
            act = ", ".join(str(x) for x in act)
        d["verified"] = bool(v.get("verified"))
        d["engine_version"] = v.get("engine_version", "")
        d["example_formula"] = (d.get("verify") or {}).get("formula", "")
        d["example_actual"] = act
        recs.append(d)
    return recs


def load_comparisons():
    cdir = DATA_DIR / "comparisons"
    if not cdir.exists():
        return []
    return [json.loads(p.read_text()) for p in sorted(cdir.glob("*.json"))]


def breadcrumb_ld(items):
    """items: list of (name, absolute_url). Returns compact BreadcrumbList JSON-LD
    so deep pages show a breadcrumb trail in search results instead of a raw URL."""
    return json.dumps({
        "@context": "https://schema.org",
        "@type": "BreadcrumbList",
        "itemListElement": [
            {"@type": "ListItem", "position": i + 1, "name": n, "item": u}
            for i, (n, u) in enumerate(items)
        ],
    }, separators=(",", ":"))


def build_env():
    env = Environment(
        loader=DictLoader(
            {
                "base.html": BASE_TMPL,
                "index.html": INDEX_TMPL,
                "function.html": FUNCTION_TMPL,
                "quirks.html": QUIRKS_TMPL,
                "recipe.html": RECIPE_TMPL,
                "recipe_index.html": RECIPE_INDEX_TMPL,
                "comparison.html": COMPARISON_TMPL,
                "comparison_index.html": COMPARISON_INDEX_TMPL,
                "methodology.html": METHODOLOGY_TMPL,
                "exclusive.html": EXCLUSIVE_TMPL,
                "pillar_xvg.html": PILLAR_XVG_TMPL,
                "errors.html": ERRORS_TMPL,
                "equiv.html": EQUIV_TMPL,
                "dataset.html": DATASET_TMPL,
                "checker.html": CHECKER_TMPL,
                "whatsnew.html": WHATSNEW_TMPL,
                "seo_page.html": SEO_PAGE_TMPL,
                "guides_index.html": GUIDES_INDEX_TMPL,
                "sitemap.xml": SITEMAP_TMPL,
            }
        ),
        autoescape=select_autoescape(["html", "xml"]),
    )
    env.filters["dateonly"] = dateonly_filter
    env.filters["fmtval"] = fmtval_filter
    return env


# --------------------------------------------------------------------------
# Rendering
# --------------------------------------------------------------------------

def common_ctx(rel):
    return {
        "site_name": SITE_NAME,
        "site_name_html": f"{SITE_NAME.replace('Spreadsheet?', '<span>Spreadsheet?</span>')}",
        "site_tagline": SITE_TAGLINE,
        "github_url": GITHUB_URL,
        "css": CSS,
        "search_js": SEARCH_JS,
        "rel": rel,
        "engine_order": ENGINE_ORDER,
        "verdict_label": VERDICT_LABELS,
        "verdict_short": VERDICT_SHORT,
        "verdict_class": VERDICT_BADGE_CLASS,
    }


def copy_static_extras():
    """Files that must survive every rebuild (CNAME, search-engine verification)."""
    static_dir = ROOT / "site" / "static"
    if static_dir.exists():
        for f in static_dir.iterdir():
            if f.is_file():
                shutil.copy2(f, OUT_DIR / f.name)


def main():
    functions_doc = load_functions()
    tests_by_fn = load_tests()
    results_by_engine = load_results()
    lo_versions = load_lo_versions()

    records, quirks = build_records(
        functions_doc, tests_by_fn, results_by_engine, lo_versions
    )

    tested_functions = [r for r in records if r["any_tested"]]
    tested_case_count = sum(
        len(e["cases"]) for r in records for e in r["engines"].values() if e["tested"]
    )
    stats = {
        "total_functions": len(records),
        "engines_targeted": len(ENGINE_ORDER),
        "engines_executed": len(results_by_engine),
        "tested_functions": len(tested_functions),
        "tested_case_count": tested_case_count,
        "quirk_count": len(quirks),
    }

    top_functions = sorted(
        (r for r in records if r["quirk_count"] > 0),
        key=lambda r: (0 if r["primary_verdict"] == "quirky" else 1, -r["quirk_count"], r["name"]),
    )[:8]

    # Curated most-searched functions (Search Console impression leaders +
    # universally high-volume lookups/logic/text/date functions). Surfacing them
    # as prominent homepage links concentrates internal link equity on the pages
    # that already earn impressions, and gets visitors to common functions fast.
    # Only link functions that have a live-tested (indexable) page.
    _by_name = {r["name"]: r for r in records}
    # Ordered by GSC-proven demand: the functions actually pulling impressions
    # (offset, median, minifs, forecast.ets, mround, char, error.type, iserror,
    # numbervalue, abs, stdev — from Search Console, Aug 2026) go FIRST so the
    # homepage "Popular functions" block concentrates internal-link authority on
    # the pages Google already surfaces (all stuck ~page 3), then the
    # conventional lookup/logic staples follow.
    _POPULAR = [
        "OFFSET", "MEDIAN", "MINIFS", "FORECAST.ETS", "MROUND", "CHAR",
        "ERROR.TYPE", "ISERROR", "NUMBERVALUE", "ABS", "STDEV",
        "VLOOKUP", "XLOOKUP", "INDEX", "MATCH", "INDIRECT",
        "IF", "IFS", "IFERROR", "SUMIF", "SUMIFS", "COUNTIF", "COUNTIFS",
        "FILTER", "SORT", "UNIQUE", "SEQUENCE", "TEXTJOIN", "CONCATENATE",
        "TEXTSPLIT", "SUBSTITUTE", "LEFT", "MID", "RIGHT", "TEXT",
        "DATEDIF", "EOMONTH", "WEEKDAY", "NETWORKDAYS", "WORKDAY",
        "AVERAGEIF", "MAXIFS", "ROUND", "MOD", "PMT",
    ]
    popular_functions = [
        _by_name[n] for n in _POPULAR
        if n in _by_name and _by_name[n]["any_tested"]
    ]

    quirks.sort(key=lambda q: (q["function"], q["case"].get("id", "")))
    quirk_fn_count = len({q["function"] for q in quirks})

    if OUT_DIR.exists():
        shutil.rmtree(OUT_DIR)
    (OUT_DIR / "functions").mkdir(parents=True)

    env = build_env()
    build_date = iso_date(functions_doc.get("generated_at")) or datetime.now(timezone.utc).strftime("%Y-%m-%d")
    sitemap_urls = []

    # ---- Homepage ----
    ctx = common_ctx(rel="")
    ctx.update(
        page_title=f"{SITE_NAME} — Excel vs Google Sheets vs LibreOffice function compatibility",
        meta_description=(
            f"{stats['total_functions']} spreadsheet functions checked for real "
            f"compatibility across Excel, Google Sheets, and LibreOffice Calc. "
            f"{stats['quirk_count']} quirks found from {stats['tested_case_count']} "
            f"executed, recalculation-proven test cases."
        ),
        canonical=BASE_URL,
        functions=records,
        stats=stats,
        top_functions=top_functions,
        popular_functions=popular_functions,
        json_ld=json.dumps({
            "@context": "https://schema.org",
            "@type": "WebSite",
            "name": SITE_NAME,
            "url": BASE_URL,
            "description": (
                "Executed compatibility results for spreadsheet functions across "
                "Microsoft Excel, Google Sheets, and LibreOffice Calc."
            ),
            "potentialAction": {
                "@type": "SearchAction",
                "target": {"@type": "EntryPoint",
                           "urlTemplate": BASE_URL + "?q={search_term_string}"},
                "query-input": "required name=search_term_string",
            },
        }, separators=(",", ":")),
    )
    (OUT_DIR / "index.html").write_text(env.get_template("index.html").render(**ctx))
    sitemap_urls.append({"loc": BASE_URL, "lastmod": build_date})

    # ---- Quirks page ----
    latest_result_date = build_date
    for res in results_by_engine.values():
        d = iso_date(res.get("generated_at"))
        if d and d > latest_result_date:
            latest_result_date = d

    ctx = common_ctx(rel="")
    ctx.update(
        page_title=f"Spreadsheet function quirks — real Excel/Google Sheets/LibreOffice divergence | {SITE_NAME}",
        meta_description=(
            f"{stats['quirk_count']} real, executed spreadsheet function results that "
            f"diverge from documented behavior, found across {quirk_fn_count} functions."
        ),
        canonical=BASE_URL + "quirks.html",
        quirks=quirks,
        quirk_fn_count=quirk_fn_count,
        seo_guides=load_seo_pages(),
    )
    (OUT_DIR / "quirks.html").write_text(env.get_template("quirks.html").render(**ctx))
    sitemap_urls.append({"loc": BASE_URL + "quirks.html", "lastmod": latest_result_date})

    # ---- Function pages ----
    # Map each function -> how-to recipes that use it (internal linking).
    recipes_for_links = load_recipes()
    _fnre = re.compile(r"([A-Za-z][A-Za-z0-9_.]*)\s*\(")
    func_recipes = {}
    for rc in recipes_for_links:
        seen = set()
        for s in rc.get("solutions", {}).values():
            for m in _fnre.finditer(s.get("formula", "")):
                seen.add(m.group(1).upper())
        for fn in seen:
            func_recipes.setdefault(fn, []).append({"slug": rc["slug"], "title": rc["title"]})

    # Map each function -> comparison pages that feature it (internal linking).
    func_comparisons = {}
    for cp in load_comparisons():
        for fn in cp.get("funcs", []):
            func_comparisons.setdefault(fn.upper(), []).append(
                {"slug": cp["slug"], "title": cp["title"]}
            )

    func_tmpl = env.get_template("function.html")
    for r in records:
        page_date = r["last_tested"] or build_date
        if r["any_tested"]:
            title = f"{r['name']} function: Excel vs Google Sheets vs LibreOffice compatibility"
            desc = (
                f"Does {r['name']} work the same in Excel, Google Sheets, and "
                f"LibreOffice Calc? Real executed test results, syntax, and links to "
                f"each official doc for the {r['name']} function ({r['category']})."
            )
        else:
            title = f"{r['name']} function — is it in Excel, Google Sheets & LibreOffice?"
            desc = (
                f"{r['name']} ({r['category']}) documentation inventory: "
                f"is it documented for Excel, Google Sheets, and LibreOffice Calc? "
                f"Not yet live-tested by a real engine."
            )
        # Thin, untested stub pages (only a documentation-inventory table, no
        # executed data) are near-duplicates. On a young domain that already drew
        # a Google "alternate/duplicate" warning, they dilute crawl budget and
        # site-quality signal away from the ~455 pages that carry real executed
        # data. noindex them (keep crawlable/followable so their outbound links
        # still pass equity, and keep them reachable for direct navigation) and
        # drop them from the sitemap, which should list only indexable URLs.
        # Cleanly reversible: delete this branch to re-index them.
        stub = not r["any_tested"]
        ctx = common_ctx(rel="../")
        ctx.update(
            page_title=title,
            meta_description=desc,
            canonical=BASE_URL + f"functions/{r['name_lower']}.html",
            r=r,
            related_recipes=func_recipes.get(r["name"], []),
            related_comparisons=func_comparisons.get(r["name"], []),
            noindex=stub,
            json_ld=breadcrumb_ld([
                (SITE_NAME, BASE_URL),
                (f"{r['name']} function", BASE_URL + f"functions/{r['name_lower']}.html"),
            ]),
        )
        out_path = OUT_DIR / "functions" / f"{r['name_lower']}.html"
        out_path.write_text(func_tmpl.render(**ctx))
        if not stub:
            sitemap_urls.append(
                {"loc": BASE_URL + f"functions/{r['name_lower']}.html", "lastmod": page_date}
            )

    # ---- How-to recipe pages ----
    recipes = load_recipes()
    # Shared-function relatedness mesh: link each recipe to sibling recipes and
    # comparisons that use the same functions. Dense, relevant internal linking
    # keeps readers moving between pages and spreads authority across the site.
    _recipe_fns = {rc["slug"]: {f["name"] for f in _functions_used(rc, _by_name)}
                   for rc in recipes}
    _rec_by_slug = {rc["slug"]: rc for rc in recipes}
    _func_to_recipes = {}
    for rc in recipes:
        for fn in _recipe_fns[rc["slug"]]:
            _func_to_recipes.setdefault(fn, []).append(rc["slug"])
    _func_to_cmps = {}
    for cp in load_comparisons():
        for fn in cp.get("funcs", []):
            _func_to_cmps.setdefault(fn.upper(), []).append(cp)

    def _related_for(rc):
        my = _recipe_fns[rc["slug"]]
        if not my:
            return [], []
        scores = {}
        for fn in my:
            for slug in _func_to_recipes.get(fn, []):
                if slug != rc["slug"]:
                    scores[slug] = scores.get(slug, 0) + 1
        ranked = sorted(scores.items(), key=lambda kv: (-kv[1], kv[0]))[:5]
        rel_recipes = [_rec_by_slug[s] for s, _ in ranked]
        seen, rel_cmps = set(), []
        for fn in sorted(my):
            for cp in _func_to_cmps.get(fn, []):
                if cp["slug"] not in seen:
                    seen.add(cp["slug"])
                    rel_cmps.append(cp)
        return rel_recipes, rel_cmps[:3]

    def _recipes_for_cmp(cp):
        """Recipes whose formulas use the functions this comparison covers —
        the reverse link, giving comparison readers task-oriented next steps."""
        funcs = {f.upper() for f in cp.get("funcs", [])}
        scores = {}
        for fn in funcs:
            for slug in _func_to_recipes.get(fn, []):
                scores[slug] = scores.get(slug, 0) + 1
        ranked = sorted(scores.items(), key=lambda kv: (-kv[1], kv[0]))[:6]
        return [_rec_by_slug[s] for s, _ in ranked]

    if recipes:
        (OUT_DIR / "how-to").mkdir(parents=True, exist_ok=True)
        rctx = common_ctx(rel="../")
        rctx.update(
            page_title="Spreadsheet how-to recipes — verified formulas for Excel, Google Sheets & LibreOffice",
            meta_description=(
                "Copy-paste formulas for common spreadsheet tasks, each executed and "
                "verified in a real engine. Excel, Google Sheets, and LibreOffice Calc."
            ),
            canonical=BASE_URL + "how-to/",
            recipes=recipes,
            grouped=[
                (cat, [r for r in recipes if _recipe_category(r["slug"]) == cat])
                for cat in [c for c, _ in _RECIPE_CATEGORIES] + ["More tasks"]
                if any(_recipe_category(r["slug"]) == cat for r in recipes)
            ],
        )
        (OUT_DIR / "how-to" / "index.html").write_text(
            env.get_template("recipe_index.html").render(**rctx)
        )
        sitemap_urls.append({"loc": BASE_URL + "how-to/", "lastmod": build_date})
        for rc in recipes:
            kw = ", ".join(rc.get("keywords", [])[:3])
            _rr, _rc = _related_for(rc)
            cx = common_ctx(rel="../")
            cx.update(
                page_title=rc["title"],
                meta_description=(
                    f"{rc['task']} Verified formula for Excel, Google Sheets and "
                    f"LibreOffice Calc" + (f" ({kw})." if kw else ".")
                ),
                canonical=BASE_URL + f"how-to/{rc['slug']}.html",
                r=rc,
                app_order=ENGINE_ORDER,
                app_labels=ENGINE_LABELS,
                functions_used=_functions_used(rc, _by_name),
                related_recipes=_rr,
                related_comparisons=_rc,
                json_ld=breadcrumb_ld([
                    (SITE_NAME, BASE_URL),
                    ("How-to recipes", BASE_URL + "how-to/"),
                    (rc["title"], BASE_URL + f"how-to/{rc['slug']}.html"),
                ]),
            )
            (OUT_DIR / "how-to" / f"{rc['slug']}.html").write_text(
                env.get_template("recipe.html").render(**cx)
            )
            sitemap_urls.append(
                {"loc": BASE_URL + f"how-to/{rc['slug']}.html", "lastmod": build_date}
            )

    # ---- SEO guide pages (formulas that behave differently across engines) ----
    seo_pages = load_seo_pages()
    if seo_pages:
        (OUT_DIR / "guides").mkdir(parents=True, exist_ok=True)
        gictx = common_ctx(rel="../")
        gictx.update(
            page_title="Formula behavior guides — where Excel, Google Sheets & LibreOffice disagree",
            meta_description=(
                "Executed-data writeups of formulas that return different results in "
                "Excel, Google Sheets, and LibreOffice Calc for the same input. "
                "LibreOffice values are executed; Excel and Google Sheets are documented "
                "unless stated otherwise."
            ),
            canonical=BASE_URL + "guides/",
            guides=seo_pages,
        )
        (OUT_DIR / "guides" / "index.html").write_text(
            env.get_template("guides_index.html").render(**gictx)
        )
        sitemap_urls.append({"loc": BASE_URL + "guides/", "lastmod": latest_result_date})
        seo_tmpl = env.get_template("seo_page.html")
        for sp in seo_pages:
            gx = common_ctx(rel="../")
            gx.update(
                page_title=sp["title"],
                meta_description=sp["meta_description"],
                canonical=BASE_URL + f"guides/{sp['slug']}.html",
                h1=sp["h1"],
                body_html=sp["body_html"],
                noindex=False,
                json_ld=breadcrumb_ld([
                    (SITE_NAME, BASE_URL),
                    ("Quirks", BASE_URL + "quirks.html"),
                    (sp["title"], BASE_URL + f"guides/{sp['slug']}.html"),
                ]),
            )
            (OUT_DIR / "guides" / f"{sp['slug']}.html").write_text(seo_tmpl.render(**gx))
            sitemap_urls.append(
                {"loc": BASE_URL + f"guides/{sp['slug']}.html", "lastmod": latest_result_date}
            )

    # ---- Function comparison pages ----
    comparisons = load_comparisons()
    if comparisons:
        (OUT_DIR / "compare").mkdir(parents=True, exist_ok=True)
        ictx = common_ctx(rel="../")
        ictx.update(
            page_title="Spreadsheet function comparisons — VLOOKUP vs XLOOKUP and more",
            meta_description=(
                "Head-to-head guides for commonly confused spreadsheet functions: real "
                "differences, which to use when, and executed compatibility results for "
                "Excel, Google Sheets, and LibreOffice."
            ),
            canonical=BASE_URL + "compare/",
            comparisons=comparisons,
        )
        (OUT_DIR / "compare" / "index.html").write_text(
            env.get_template("comparison_index.html").render(**ictx)
        )
        sitemap_urls.append({"loc": BASE_URL + "compare/", "lastmod": build_date})
        for c in comparisons:
            cx = common_ctx(rel="../")
            cx.update(
                page_title=c["title"],
                meta_description=c["meta_desc"],
                canonical=BASE_URL + f"compare/{c['slug']}.html",
                c=c,
                related_recipes=_recipes_for_cmp(c),
                json_ld=breadcrumb_ld([
                    (SITE_NAME, BASE_URL),
                    ("Comparisons", BASE_URL + "compare/"),
                    (c["title"], BASE_URL + f"compare/{c['slug']}.html"),
                ]),
            )
            (OUT_DIR / "compare" / f"{c['slug']}.html").write_text(
                env.get_template("comparison.html").render(**cx)
            )
            sitemap_urls.append(
                {"loc": BASE_URL + f"compare/{c['slug']}.html", "lastmod": build_date}
            )

    # ---- Formula compatibility checker (client-side tool) ----
    compat_export = {}
    for r in records:
        e = r["engines"]
        lch = e["libreoffice"].get("lo_change")
        compat_export[r["name"]] = {
            "cat": r["category"],
            "x": bool(e["excel"]["documented"]),
            "g": bool(e["google_sheets"]["documented"]),
            "l": bool(e["libreoffice"]["documented"]),
            "lv": e["libreoffice"]["verdict"],
            "lver": e["libreoffice"]["version"],
            # newly supported: the exact release it started working in (else null)
            "lnew": lch["since_version"] if (lch and lch["newly_supported"]) else None,
        }
    (OUT_DIR / "data").mkdir(parents=True, exist_ok=True)
    (OUT_DIR / "data" / "compat.json").write_text(
        json.dumps(compat_export, separators=(",", ":"))
    )
    # CSV mirror of the dataset (easier to open in a spreadsheet / load into
    # data tools, and the format dataset registries expect).
    import csv as _csv
    import io as _io
    _buf = _io.StringIO()
    _w = _csv.writer(_buf)
    _w.writerow([
        "function", "category", "in_excel", "in_google_sheets", "in_libreoffice",
        "libreoffice_verdict", "libreoffice_version_tested",
        "libreoffice_newly_supported_in",
    ])
    for _name in sorted(compat_export):
        _v = compat_export[_name]
        _w.writerow([
            _name, _v["cat"], _v["x"], _v["g"], _v["l"],
            _v["lv"], _v["lver"], _v["lnew"] if _v["lnew"] is not None else "",
        ])
    (OUT_DIR / "data" / "compat.csv").write_text(_buf.getvalue())
    cctx = common_ctx(rel="")
    cctx.update(
        page_title="Spreadsheet formula compatibility checker — Excel, Google Sheets & LibreOffice",
        meta_description=(
            "Paste a formula and instantly see whether every function works in Excel, "
            "Google Sheets, and current LibreOffice Calc. Based on real executed tests."
        ),
        canonical=BASE_URL + "checker.html",
    )
    (OUT_DIR / "checker.html").write_text(env.get_template("checker.html").render(**cctx))
    sitemap_urls.append({"loc": BASE_URL + "checker.html", "lastmod": build_date})
    sitemap_urls.append({"loc": BASE_URL + "audit.html", "lastmod": build_date})

    # ---- App-exclusive function pages (data-driven from documented flags) ----
    def _excl_items(pred):
        out = []
        for r in records:
            e = r["engines"]
            if pred(e):
                out.append({
                    "name": r["name"], "name_lower": r["name_lower"],
                    "category": r["category"],
                    "lo_verdict": e["libreoffice"]["verdict"],
                    "lo_doc": e["libreoffice"]["documented"],
                })
        return sorted(out, key=lambda x: (x["category"], x["name"]))

    for slug, heading, lede, outro, items, show_lo in [
        (
            "sheets-functions-not-in-excel",
            "Google Sheets functions that don't exist in Excel",
            "QUERY, ARRAYFORMULA, IMPORTRANGE, GOOGLEFINANCE, the REGEX family… "
            "these work only in Google Sheets. Export the file to .xlsx and every "
            "cell using them fails.",
            "Portable alternatives exist for many of these — FILTER/SORT/UNIQUE for "
            "much of QUERY, native spilling instead of ARRAYFORMULA — see the "
            "comparison pages for migration patterns.",
            _excl_items(lambda e: e["google_sheets"]["documented"] and not e["excel"]["documented"]),
            True,
        ),
        (
            "excel-functions-not-in-google-sheets",
            "Excel functions that don't exist in Google Sheets",
            "PIVOTBY, GROUPBY, the CUBE family, AGGREGATE and more are documented "
            "for Excel but absent from Google Sheets — importing an .xlsx that uses "
            "them leaves broken formulas.",
            "The LibreOffice column shows that \"Excel-only\" is often really "
            "\"not-in-Sheets\": several of these run fine in LibreOffice Calc per "
            "our executed tests.",
            _excl_items(lambda e: e["excel"]["documented"] and not e["google_sheets"]["documented"]),
            True,
        ),
    ]:
        ectx = common_ctx(rel="")
        ectx.update(
            page_title=heading + " — full list",
            meta_description=lede + f" Full list of {len(items)} functions with categories.",
            canonical=BASE_URL + f"{slug}.html",
            heading=heading, lede=lede, outro=outro, items=items, show_lo=show_lo,
        )
        (OUT_DIR / f"{slug}.html").write_text(env.get_template("exclusive.html").render(**ectx))
        sitemap_urls.append({"loc": BASE_URL + f"{slug}.html", "lastmod": build_date})

    # ---- Error values reference page ----
    ectx2 = common_ctx(rel="")
    ectx2.update(
        page_title="#NAME?, #REF!, #VALUE!, #SPILL! — spreadsheet errors explained and fixed",
        meta_description=(
            "What every spreadsheet error value means — #NAME?, #REF!, #VALUE!, "
            "#DIV/0!, #N/A, #SPILL!, #NUM! — with likely causes ranked and the "
            "fastest fixes, for Excel, Google Sheets, and LibreOffice."
        ),
        canonical=BASE_URL + "spreadsheet-errors.html",
    )
    (OUT_DIR / "spreadsheet-errors.html").write_text(
        env.get_template("errors.html").render(**ectx2)
    )
    sitemap_urls.append({"loc": BASE_URL + "spreadsheet-errors.html", "lastmod": build_date})

    # ---- Open dataset documentation page ----
    _compat_path = OUT_DIR / "data" / "compat.json"
    _compat_obj = json.loads(_compat_path.read_text()) if _compat_path.exists() else {}
    dctx = common_ctx(rel="")
    dctx.update(
        page_title="Open spreadsheet function compatibility dataset (CC BY) — Excel, Sheets, LibreOffice",
        meta_description=(
            "Free, machine-verified dataset of spreadsheet function compatibility "
            "across Excel, Google Sheets, and LibreOffice Calc — executed results "
            "with per-version history, as JSON under CC BY 4.0. Schema and examples."
        ),
        canonical=BASE_URL + "data.html",
        n_funcs=len(_compat_obj),
        kb=max(1, round(_compat_path.stat().st_size / 1024)) if _compat_path.exists() else 0,
        json_ld=json.dumps({
            "@context": "https://schema.org",
            "@type": "Dataset",
            "name": "Spreadsheet function compatibility (Excel, Google Sheets, LibreOffice)",
            "description": (
                f"Machine-verified compatibility data for {len(_compat_obj)} spreadsheet "
                "functions across Microsoft Excel and Google Sheets (from official "
                "documentation) and LibreOffice Calc (from executed test results, with "
                "per-version history). The only openly available executed cross-application "
                "spreadsheet compatibility dataset."
            ),
            "url": BASE_URL + "data.html",
            "keywords": ["spreadsheet", "Excel", "Google Sheets", "LibreOffice",
                         "function compatibility", "formula compatibility"],
            "license": "https://creativecommons.org/licenses/by/4.0/",
            "isAccessibleForFree": True,
            "creator": {"@type": "Organization", "name": SITE_NAME, "url": BASE_URL},
            "distribution": [{
                "@type": "DataDownload",
                "encodingFormat": "application/json",
                "contentUrl": BASE_URL + "data/compat.json",
            }],
        }, separators=(",", ":")),
    )
    (OUT_DIR / "data.html").write_text(env.get_template("dataset.html").render(**dctx))
    sitemap_urls.append({"loc": BASE_URL + "data.html", "lastmod": build_date})

    # ---- Excel <-> Google Sheets equivalents reference ----
    equiv_g_only = [  # Sheets-only -> Excel equivalent
        {"fn": "QUERY", "note": "No single equivalent. Use FILTER for SELECT/WHERE; add SUMIFS/UNIQUE/SORT for grouping and aggregation. For heavy reshaping, Power Query."},
        {"fn": "ARRAYFORMULA", "note": "Not needed — Excel 365 spills array expressions natively. Just write =A2:A10*B2:B10."},
        {"fn": "REGEXMATCH", "note": "No regex in Excel. Use ISNUMBER(SEARCH(\"text\",A2)) for contains, or wildcards in COUNTIF/SUMIF."},
        {"fn": "REGEXEXTRACT", "note": "Use LEFT/MID/RIGHT with FIND, or TEXTBEFORE/TEXTAFTER (Excel 365)."},
        {"fn": "REGEXREPLACE", "note": "Use SUBSTITUTE (by text), nested for multiple patterns."},
        {"fn": "GOOGLEFINANCE", "note": "No native equivalent. Excel 365 has the STOCKHISTORY function and Stocks data type for some of this."},
        {"fn": "IMPORTRANGE", "note": "Link workbooks with external references, or pull data via Power Query."},
        {"fn": "IMPORTHTML", "note": "Power Query: Data → From Web imports HTML tables and lists."},
        {"fn": "IMPORTXML", "note": "Power Query (Data → From Web) for structured web data."},
        {"fn": "SPLIT", "note": "TEXTSPLIT (Excel 365), or the Text to Columns wizard."},
        {"fn": "JOIN", "note": "TEXTJOIN(delimiter, TRUE, range) — same idea, works in both."},
        {"fn": "FLATTEN", "note": "TOCOL(range) in Excel 365 flattens a range to one column."},
    ]
    equiv_x_only = [  # Excel-only -> Sheets equivalent
        {"fn": "GROUPBY", "note": "Not in Sheets yet. Use a pivot table, or SUMIFS/COUNTIFS over UNIQUE keys."},
        {"fn": "PIVOTBY", "note": "Use a pivot table, or QUERY with GROUP BY."},
        {"fn": "STOCKHISTORY", "note": "GOOGLEFINANCE(ticker, ...) pulls historical prices in Sheets."},
        {"fn": "AGGREGATE", "note": "No direct equivalent; SUBTOTAL covers the filter-aware subset, or FILTER out errors then aggregate."},
        {"fn": "TEXTBEFORE", "note": "Available in Sheets too; older sheets use LEFT(A2,FIND(delim,A2)-1)."},
        {"fn": "ARRAYTOTEXT", "note": "TEXTJOIN(\", \",TRUE,range) approximates it in Sheets."},
        {"fn": "IMAGE", "note": "Sheets has IMAGE(url) too — one of the few that matches."},
    ]
    equiv_gotcha = [  # same name, different behavior
        {"fn": "SORT", "note": "Direction argument differs: Excel uses 1/-1 (asc/desc); Google Sheets uses TRUE/FALSE."},
        {"fn": "WEEKDAY", "note": "Both support the return-type argument, but be explicit about it — the default (Sunday=1) trips up weekend logic in both."},
        {"fn": "TEXT", "note": "Format codes mostly match, but locale affects date/number separators — a comma-decimal locale reads codes differently."},
        {"fn": "CONCATENATE", "note": "Works in both, but Sheets also allows CONCATENATE of a whole range; Excel's version is cell-by-cell (use CONCAT/TEXTJOIN)."},
        {"fn": "FILTER", "note": "Same core behavior; the 'no matches' fallback argument exists in both (Excel 365), but error text differs."},
    ]
    _fn_names = {r["name"] for r in records}
    for _row in equiv_g_only + equiv_x_only:
        _row["exists"] = _row["fn"] in _fn_names
    eqctx = common_ctx(rel="")
    eqctx.update(
        page_title="Excel to Google Sheets function equivalents — what to use instead",
        meta_description=(
            "Google Sheets equivalents for Excel-only functions (GROUPBY, PIVOTBY, "
            "AGGREGATE) and Excel equivalents for Sheets-only ones (QUERY, "
            "ARRAYFORMULA, REGEXMATCH, IMPORTRANGE) — plus same-name gotchas. "
            "Verified from executed compatibility tests."
        ),
        canonical=BASE_URL + "excel-google-sheets-equivalents.html",
        g_only=equiv_g_only, x_only=equiv_x_only, gotcha=equiv_gotcha,
    )
    (OUT_DIR / "excel-google-sheets-equivalents.html").write_text(
        env.get_template("equiv.html").render(**eqctx)
    )
    sitemap_urls.append({"loc": BASE_URL + "excel-google-sheets-equivalents.html", "lastmod": build_date})

    # ---- Excel vs Google Sheets pillar page ----
    n_xonly = sum(1 for r in records if r["engines"]["excel"]["documented"] and not r["engines"]["google_sheets"]["documented"])
    n_gonly = sum(1 for r in records if r["engines"]["google_sheets"]["documented"] and not r["engines"]["excel"]["documented"])
    n_both = sum(1 for r in records if r["engines"]["excel"]["documented"] and r["engines"]["google_sheets"]["documented"])
    pctx = common_ctx(rel="")
    pctx.update(
        page_title="Excel vs Google Sheets: formula compatibility guide (what breaks and why)",
        meta_description=(
            f"Excel vs Google Sheets for formulas: {n_xonly} Excel-only functions, "
            f"{n_gonly} Sheets-only functions, dialect differences like ARRAYFORMULA "
            "vs spilling, and how to keep workbooks portable — with executed test data."
        ),
        canonical=BASE_URL + "excel-vs-google-sheets.html",
        n_xonly=n_xonly, n_gonly=n_gonly, n_both=n_both,
        n_tested=len(lo_versions[-1][1].get("function_results", {})) if lo_versions else 0,
        versions=[v for v, _ in lo_versions],
    )
    (OUT_DIR / "excel-vs-google-sheets.html").write_text(
        env.get_template("pillar_xvg.html").render(**pctx)
    )
    sitemap_urls.append({"loc": BASE_URL + "excel-vs-google-sheets.html", "lastmod": build_date})

    # ---- Methodology page ----
    if lo_versions:
        latest_blob = lo_versions[-1][1]
        fr = latest_blob.get("function_results", {})
        mctx = common_ctx(rel="")
        mctx.update(
            page_title="Methodology — how canispreadsheet verifies formulas by executing them",
            meta_description=(
                "How this site tests spreadsheet functions: headless LibreOffice "
                "execution, recalculation canaries, the OOXML _xlfn storage-prefix "
                "gotcha, a multi-release version matrix, and honest limitations."
            ),
            canonical=BASE_URL + "methodology.html",
            versions=[v for v, _ in lo_versions],
            current_version=lo_versions[-1][0],
            n_funcs=len(fr),
            n_cases=sum(len(v) for v in fr.values()),
        )
        (OUT_DIR / "methodology.html").write_text(
            env.get_template("methodology.html").render(**mctx)
        )
        sitemap_urls.append({"loc": BASE_URL + "methodology.html", "lastmod": build_date})

    # ---- LibreOffice version-support (caniuse-style) page ----
    lo_ver_list = [v for v, _ in lo_versions]
    if len(lo_ver_list) >= 2:
        from_v, to_v = lo_ver_list[0], lo_ver_list[-1]
        newly, other = [], []
        for r in records:
            ch = r["engines"]["libreoffice"].get("lo_change")
            if not ch:
                continue
            hm = {
                h["version"]: h["verdict"]
                for h in r["engines"]["libreoffice"].get("lo_history", [])
            }
            row = {
                "name": r["name"],
                "name_lower": r["name_lower"],
                "category": r["category"],
                "since": ch.get("since_version"),
                # per-version verdicts aligned to lo_ver_list (column order)
                "verdicts": [hm.get(v) for v in lo_ver_list],
            }
            (newly if ch["newly_supported"] else other).append(row)
        # newly-supported: group by the release they landed in, newest first
        newly.sort(key=lambda x: (_version_tuple(x["since"]), x["name"]))
        other.sort(key=lambda x: x["name"])
        wctx = common_ctx(rel="")
        wctx.update(
            page_title=(
                f"LibreOffice Calc function support by version — "
                f"what's new in {to_v} (XLOOKUP, FILTER, SORT, UNIQUE…)"
            ),
            meta_description=(
                f"Machine-verified LibreOffice Calc function compatibility by version: "
                f"{len(newly)} functions — including XLOOKUP, FILTER, SORT, UNIQUE, LET and "
                f"other dynamic-array functions — that returned #NAME? in LibreOffice {from_v} "
                f"now work in {to_v}. Real executed test results."
            ),
            canonical=BASE_URL + "libreoffice-version-support.html",
            from_version=from_v,
            to_version=to_v,
            newly_supported=newly,
            other_changes=other,
            versions_tested=lo_ver_list,
        )
        (OUT_DIR / "libreoffice-version-support.html").write_text(
            env.get_template("whatsnew.html").render(**wctx)
        )
        sitemap_urls.append(
            {"loc": BASE_URL + "libreoffice-version-support.html", "lastmod": latest_result_date}
        )

    # ---- sitemap.xml + robots.txt ----
    sitemap_xml = env.get_template("sitemap.xml").render(urls=sitemap_urls)
    (OUT_DIR / "sitemap.xml").write_text(sitemap_xml)

    # ---- Migration Audit page (static app maintained in site/audit-page/) ----
    # Copied at build time so rebuilds never drop it from docs/.
    import shutil as _shutil
    _audit_src = ROOT / "site" / "audit-page"
    for _f in ("audit.html", "audit.js", "audit-verdicts.js", "audit-app.js"):
        _shutil.copyfile(_audit_src / _f, OUT_DIR / _f)
    (OUT_DIR / "robots.txt").write_text(f"User-agent: *\nAllow: /\nSitemap: {BASE_URL}sitemap.xml\n")
    (OUT_DIR / ".nojekyll").write_text("")
    copy_static_extras()

    print(f"Built {len(records)} function pages.")
    print(f"Stats: {json.dumps(stats, indent=2)}")
    print(f"Top functions: {[r['name'] for r in top_functions]}")


if __name__ == "__main__":
    main()

