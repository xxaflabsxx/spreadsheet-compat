# Wiring `seo-pages/*.json` into the build

These pages are long-form "behaves differently across engines" guides. Each JSON file
here has `slug`, `title`, `meta_description`, `h1`, `body_html`. This doc describes the
exact edits to `site/build_site.py` to render them to `docs/guides/<slug>.html` and add
them to the sitemap. **Nothing here modifies `build_site.py` for you — apply the hunks
below.**

Key facts the code must honor (already baked into the JSON):

- Output path is `docs/guides/<slug>.html`, one level below the site root, so the render
  context must use **`rel="../"`** (same depth as `functions/` and `how-to/`).
- `body_html` already contains the two CTAs as `../checker.html` and `../audit.html`
  relative links, plus tables using the existing `table.cases` / `table-scroll` /
  `quirk-box` / `section-title` / `promo-card` CSS classes. It renders with `| safe`.
- These pages carry real executed data, so they should be **indexable** (`noindex=False`)
  and **belong in the sitemap** (unlike the thin stub function pages).

---

## Hunk 1 — a constant + a loader (near the other `load_*` helpers)

```diff
@@ ROOT = Path(__file__).resolve().parent.parent
 DATA_DIR = ROOT / "data"
 TESTS_DIR = DATA_DIR / "tests"
 RESULTS_DIR = ROOT / "results"
 OUT_DIR = ROOT / "docs"
+SEO_PAGES_DIR = ROOT / "site" / "seo-pages"
```

```diff
@@ def load_recipes():
+def load_seo_pages():
+    """Long-form 'behaves differently across engines' guides authored as one
+    JSON file per page in site/seo-pages/. Each: slug, title, meta_description,
+    h1, body_html (raw, pre-rendered HTML using the site's CSS classes)."""
+    pages = []
+    if SEO_PAGES_DIR.exists():
+        for p in sorted(SEO_PAGES_DIR.glob("*.json")):
+            pages.append(json.loads(p.read_text()))
+    return pages
+
+
 def load_recipes():
     recs = []
```

## Hunk 2 — a template, registered in `build_env`'s `DictLoader`

Define the template constant alongside the other `*_TMPL` strings (any location before
`build_env`), then register it. It links back to the Quirks hub and renders `body_html`
verbatim:

```python
SEO_PAGE_TMPL = """{% extends "base.html" %}
{% block content %}
<a class="back-link" href="{{ rel }}quirks.html">&larr; All quirks &amp; gotchas</a>
<h1>{{ h1 }}</h1>
{{ body_html | safe }}
{% endblock %}"""
```

```diff
@@ "dataset.html": DATASET_TMPL,
                 "checker.html": CHECKER_TMPL,
                 "whatsnew.html": WHATSNEW_TMPL,
+                "seo_page.html": SEO_PAGE_TMPL,
                 "sitemap.xml": SITEMAP_TMPL,
             }
```

## Hunk 3 — render loop in `main()` (right after the how-to block)

Insert immediately after the how-to recipe loop closes and before
`# ---- Function comparison pages ----`:

```diff
@@             sitemap_urls.append(
                 {"loc": BASE_URL + f"how-to/{rc['slug']}.html", "lastmod": build_date}
             )
+
+    # ---- SEO guide pages (formulas that behave differently across engines) ----
+    seo_pages = load_seo_pages()
+    if seo_pages:
+        (OUT_DIR / "guides").mkdir(parents=True, exist_ok=True)
+        seo_tmpl = env.get_template("seo_page.html")
+        for sp in seo_pages:
+            gx = common_ctx(rel="../")
+            gx.update(
+                page_title=sp["title"],
+                meta_description=sp["meta_description"],
+                canonical=BASE_URL + f"guides/{sp['slug']}.html",
+                h1=sp["h1"],
+                body_html=sp["body_html"],
+                noindex=False,
+                json_ld=breadcrumb_ld([
+                    (SITE_NAME, BASE_URL),
+                    ("Quirks", BASE_URL + "quirks.html"),
+                    (sp["title"], BASE_URL + f"guides/{sp['slug']}.html"),
+                ]),
+            )
+            (OUT_DIR / "guides" / f"{sp['slug']}.html").write_text(seo_tmpl.render(**gx))
+            sitemap_urls.append(
+                {"loc": BASE_URL + f"guides/{sp['slug']}.html", "lastmod": latest_result_date}
+            )

     # ---- Function comparison pages ----
     comparisons = load_comparisons()
```

`latest_result_date` is already computed earlier in `main()` (used for the quirks page) and
is in scope at this point; use it so `lastmod` reflects when the executed data was produced.
Fall back to `build_date` if you prefer a build-time stamp.

---

## Verified

Rendering `count-...json` through the real `build_env()` base template with the
`SEO_PAGE_TMPL` above produced valid HTML (16.4 KB) with the `cases` table, `table-scroll`,
`quirk-box`, `section-title`, and both `promo-btn` CTAs (`../checker.html`, `../audit.html`)
present and correctly resolved. No template variables are left unfilled.

## Optional follow-ups (not required)

- **Discovery links.** Add a small linked list of these guides to the bottom of
  `QUIRKS_TMPL` (they are the narrative long-form version of the raw quirk entries) and/or a
  card in the homepage grid, so they earn internal links rather than sitemap-only discovery.
- **Guides index.** If the set grows, add a `docs/guides/index.html` listing them and a
  `guides/` nav link, mirroring the `how-to/` index pattern.
