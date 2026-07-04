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
import shutil
from datetime import datetime, timezone
from pathlib import Path

from jinja2 import Environment, DictLoader, select_autoescape

# --------------------------------------------------------------------------
# Config — change branding/deployment details here, nowhere else.
# --------------------------------------------------------------------------

SITE_NAME = "Can I Spreadsheet?"
SITE_TAGLINE = "caniuse.com for spreadsheet functions"
BASE_URL = "https://xxaflabsxx.github.io/spreadsheet-compat/"
ACCENT = "#4F46E5"
GITHUB_URL = "https://github.com/xxAFLabsxx/spreadsheet-compat"

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"
TESTS_DIR = DATA_DIR / "tests"
RESULTS_DIR = ROOT / "results"
OUT_DIR = ROOT / "docs"

ENGINE_ORDER = ["excel", "google_sheets", "libreoffice"]
ENGINE_LABELS = {
    "excel": "Excel",
    "google_sheets": "Google Sheets",
    "libreoffice": "LibreOffice Calc",
}

VERDICT_LABELS = {
    "supported": "Supported, behaves as documented",
    "quirky": "Supported, with a discovered quirk",
    "unsupported": "Unsupported (not recognized)",
}
VERDICT_BADGE_CLASS = {
    "supported": "badge-good",
    "quirky": "badge-quirk",
    "unsupported": "badge-bad",
    None: "badge-unknown",
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
    """engine key -> raw results blob for that engine's results/*.json file"""
    out = {}
    for p in sorted(RESULTS_DIR.glob("*.json")):
        d = json.loads(p.read_text())
        key = engine_key_from_engine_name(d.get("engine", ""))
        if key:
            out[key] = d
    return out


# --------------------------------------------------------------------------
# Build per-function records
# --------------------------------------------------------------------------

def classify_verdict(case_results):
    """case_results: list of executed-result dicts (raw from results file) for
    one engine, for one function. Returns 'supported' | 'quirky' | 'unsupported'.
    """
    if not case_results:
        return None
    if all(c.get("matched_expected") for c in case_results):
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


def build_records(functions_doc, tests_by_fn, results_by_engine):
    records = []
    all_quirks = []  # flattened, for the quirks page

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
.hero h1 { font-size: 2rem; margin: 0 0 0.5rem; }
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
.search-hint { color: var(--text-muted); font-size: 0.85rem; margin: 0.4rem 0 1.5rem; }

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
  border-left: 4px solid var(--accent);
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
.top-list .meta { color: var(--text-muted); font-size: 0.85rem; }

.badge {
  display: inline-block;
  font-size: 0.72rem;
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: 0.03em;
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
#fn-count { color: var(--text-muted); font-size: 0.85rem; }

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
  border-left: 5px solid var(--quirk);
  border-radius: 8px;
  padding: 1rem 1.25rem;
  margin: 1.25rem 0;
}
.quirk-box h3 { margin: 0 0 0.6rem; color: var(--quirk); font-size: 1.05rem; }
.quirk-box ul { margin: 0; padding-left: 1.1rem; }
.quirk-box li { margin-bottom: 0.6rem; }

.not-live-tested {
  background: var(--unknown-bg);
  border-left: 4px solid var(--text-muted);
  border-radius: 6px;
  padding: 0.9rem 1.15rem;
  margin: 1rem 0 1.5rem;
  font-size: 0.92rem;
  color: var(--text-muted);
}

.verdict-ok { color: var(--good); }
.verdict-bad { color: var(--bad); font-weight: 600; }

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
.quirk-grid dt { color: var(--text-muted); font-size: 0.75rem; text-transform: uppercase; letter-spacing: 0.03em; }
.quirk-grid dd { margin: 0 0 0.5rem; }

footer.site-footer {
  border-top: 1px solid var(--border);
  padding: 1.5rem 0 3rem;
  color: var(--text-muted);
  font-size: 0.85rem;
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
  if (!input || !list) return;
  var items = Array.prototype.slice.call(list.children);
  function apply() {
    var q = input.value.trim().toLowerCase();
    var shown = 0;
    items.forEach(function (li) {
      var match = !q || li.dataset.name.indexOf(q) !== -1 || li.dataset.cat.indexOf(q) !== -1;
      li.style.display = match ? '' : 'none';
      if (match) shown++;
    });
    if (count) count.textContent = shown + ' of ' + items.length + ' functions';
  }
  input.addEventListener('input', apply);
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
<link rel="canonical" href="{{ canonical }}">
<meta property="og:title" content="{{ page_title }}">
<meta property="og:description" content="{{ meta_description }}">
<meta property="og:type" content="website">
<meta property="og:url" content="{{ canonical }}">
<meta name="twitter:card" content="summary">
<style>{{ css }}</style>
</head>
<body>
<header class="site-header">
  <div class="container">
    <a class="brand" href="{{ rel }}index.html">{{ site_name_html|safe }}</a>
    <nav class="site-nav">
      <a href="{{ rel }}index.html">Functions</a>
      <a href="{{ rel }}quirks.html">Quirks</a>
      <a href="{{ github_url }}">GitHub</a>
    </nav>
  </div>
</header>
<main class="container">
{% block content %}{% endblock %}
</main>
<footer class="site-footer">
  <div class="container">
    <p>{{ site_name }} — every result on this site was executed by a real spreadsheet
    engine and recalculation-proven, never scraped from documentation alone.
    Functions without an executed-result badge are documentation-only inventory,
    clearly marked as not yet live-tested.</p>
    <p>Data and test harness on <a href="{{ github_url }}">GitHub</a>.</p>
    <p class="footer-promo">Built by AF Labs. If this site saved you a debugging session,
    we also make <a href="https://aflabs.gumroad.com" rel="sponsored">spreadsheet templates
    that do the formulas for you</a> — budgets, debt payoff, invoicing, and a complete
    freelance business hub for Excel &amp; Google Sheets.</p>
  </div>
</footer>
</body>
</html>
"""

INDEX_TMPL = """{% extends "base.html" %}
{% block content %}
<section class="hero">
  <h1>{{ site_name }}</h1>
  <p class="tagline">{{ site_tagline }} — search any function to see whether it's
  documented, tested, and how Excel, Google Sheets, and LibreOffice Calc actually
  behave.</p>

  <div class="search-box">
    <input type="search" id="fn-search" placeholder="Search a function, e.g. VLOOKUP, XLOOKUP, DATEDIF..." aria-label="Search functions">
  </div>
  <p class="search-hint" id="fn-count">{{ functions|length }} of {{ functions|length }} functions</p>
  <noscript><p>Search needs JavaScript. Every function is still listed below and
  fully linked — use your browser's find-in-page instead.</p></noscript>
</section>

<div class="stats-grid">
  <div class="stat-card"><span class="num">{{ stats.total_functions }}</span><span class="label">Functions inventoried</span></div>
  <div class="stat-card"><span class="num">{{ stats.engines_executed }}/{{ stats.engines_targeted }}</span><span class="label">Engines executed</span></div>
  <div class="stat-card"><span class="num">{{ stats.tested_case_count }}</span><span class="label">Test cases executed</span></div>
  <div class="stat-card"><span class="num">{{ stats.quirk_count }}</span><span class="label">Quirks discovered</span></div>
</div>

<div class="methodology">
  <strong>Methodology, honestly stated:</strong> every result badge on this site
  traces back to a formula that was actually written into a real workbook and
  recalculated by that engine — proven with deterministic and volatile canary
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

<script>{{ search_js }}</script>
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

{% if not r.any_tested %}
<div class="not-live-tested">
  <strong>Not yet live-tested.</strong> No engine has executed real test cases
  for {{ r.name }} yet. The table below reflects only whether each vendor's
  official documentation lists this function — it is inventory data, not a
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
  <td>{% if c.matched_expected %}<span class="verdict-ok">Matched</span>{% else %}<span class="verdict-bad">Mismatch</span>{% endif %}</td>
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
{% endblock %}
"""

QUIRKS_TMPL = """{% extends "base.html" %}
{% block content %}
<h1>Discovered quirks</h1>
<p class="tagline">Every case below is a real, executed formula whose result did
not match documented/expected behavior. This is the flagship content of
{{ site_name }} — cross-engine divergence that only shows up when you actually
run the formula.</p>
<p class="search-hint">{{ quirks|length }} quirks found across {{ quirk_fn_count }} functions.</p>

<ul class="quirks-list">
{% for q in quirks %}
<li class="quirk-entry">
  <h3><a href="{{ rel }}functions/{{ q.name_lower }}.html">{{ q.function }}</a>
  <span class="badge badge-quirk">{{ q.engine_label }}</span></h3>
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


def build_env():
    env = Environment(
        loader=DictLoader(
            {
                "base.html": BASE_TMPL,
                "index.html": INDEX_TMPL,
                "function.html": FUNCTION_TMPL,
                "quirks.html": QUIRKS_TMPL,
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

    records, quirks = build_records(functions_doc, tests_by_fn, results_by_engine)

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
    )
    (OUT_DIR / "quirks.html").write_text(env.get_template("quirks.html").render(**ctx))
    sitemap_urls.append({"loc": BASE_URL + "quirks.html", "lastmod": latest_result_date})

    # ---- Function pages ----
    func_tmpl = env.get_template("function.html")
    for r in records:
        page_date = r["last_tested"] or build_date
        if r["any_tested"]:
            title = f"{r['name']} — Excel vs Google Sheets vs LibreOffice compatibility | {SITE_NAME}"
            verdict_bits = [
                f"{e['label']}: {VERDICT_LABELS[e['verdict']]}"
                for e in r["engines"].values()
                if e["verdict"]
            ]
            desc = (
                f"{r['name']} ({r['category']}) real compatibility results: "
                + "; ".join(verdict_bits)
                + f". {r['tested_case_count']} executed test cases, last tested {page_date}."
            )
        else:
            title = f"{r['name']} — Excel vs Google Sheets vs LibreOffice compatibility (not yet live-tested) | {SITE_NAME}"
            desc = (
                f"{r['name']} ({r['category']}) documentation inventory: "
                f"is it documented for Excel, Google Sheets, and LibreOffice Calc? "
                f"Not yet live-tested by a real engine."
            )
        ctx = common_ctx(rel="../")
        ctx.update(
            page_title=title,
            meta_description=desc,
            canonical=BASE_URL + f"functions/{r['name_lower']}.html",
            r=r,
        )
        out_path = OUT_DIR / "functions" / f"{r['name_lower']}.html"
        out_path.write_text(func_tmpl.render(**ctx))
        sitemap_urls.append(
            {"loc": BASE_URL + f"functions/{r['name_lower']}.html", "lastmod": page_date}
        )

    # ---- sitemap.xml + robots.txt ----
    sitemap_xml = env.get_template("sitemap.xml").render(urls=sitemap_urls)
    (OUT_DIR / "sitemap.xml").write_text(sitemap_xml)
    (OUT_DIR / "robots.txt").write_text(f"User-agent: *\nAllow: /\nSitemap: {BASE_URL}sitemap.xml\n")
    (OUT_DIR / ".nojekyll").write_text("")
    copy_static_extras()

    print(f"Built {len(records)} function pages.")
    print(f"Stats: {json.dumps(stats, indent=2)}")
    print(f"Top functions: {[r['name'] for r in top_functions]}")


if __name__ == "__main__":
    main()
