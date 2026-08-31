#!/usr/bin/env python3
"""Mobile smoke test: loads key pages at a 390px is_mobile viewport and fails on
non-200 status, horizontal overflow (>2px), or console/page errors.

Modes:
  --docs   (default) serve ./docs locally and test the built site pre-push
  --live   test https://canispreadsheet.com after a deploy

Page set: home, guides index, audit.html, every guide page, and a deterministic
sample of N function pages (--fn-sample, default 30; --all-functions for all).
Exit 0 = clean, 1 = failures (listed), 2 = setup error.
"""
import argparse, http.server, json, random, socketserver, threading
from pathlib import Path

def collect_urls(base, docs, fn_sample, all_functions):
    urls = [f"{base}/", f"{base}/guides/", f"{base}/audit.html"]
    guides = sorted(p.name for p in (docs / "guides").glob("*.html") if p.name != "index.html")
    urls += [f"{base}/guides/{g}" for g in guides]
    fns = sorted(p.name for p in (docs / "functions").glob("*.html"))
    if not all_functions and len(fns) > fn_sample:
        fns = sorted(random.Random(20260831).sample(fns, fn_sample))
    urls += [f"{base}/functions/{f}" for f in fns]
    return urls

def main():
    ap = argparse.ArgumentParser()
    mode = ap.add_mutually_exclusive_group()
    mode.add_argument("--live", action="store_true")
    mode.add_argument("--docs", action="store_true")
    ap.add_argument("--fn-sample", type=int, default=30)
    ap.add_argument("--all-functions", action="store_true")
    args = ap.parse_args()
    docs = Path(__file__).resolve().parent.parent / "docs"
    if not docs.is_dir():
        print("docs/ not found — run site/build_site.py first"); return 2
    server = None
    if args.live:
        base = "https://canispreadsheet.com"
    else:
        handler = lambda *a, **k: http.server.SimpleHTTPRequestHandler(*a, directory=str(docs), **k)
        server = socketserver.TCPServer(("127.0.0.1", 0), handler)
        threading.Thread(target=server.serve_forever, daemon=True).start()
        base = f"http://127.0.0.1:{server.server_address[1]}"
    urls = collect_urls(base, docs, args.fn_sample, args.all_functions)
    from playwright.sync_api import sync_playwright
    bad = []
    with sync_playwright() as p:
        browser = p.chromium.launch()
        ctx = browser.new_context(viewport={"width": 390, "height": 844}, is_mobile=True)
        page = ctx.new_page()
        errs = []
        page.on("console", lambda m: errs.append(m.text) if m.type == "error" else None)
        page.on("pageerror", lambda e: errs.append(str(e)))
        for u in urls:
            errs.clear()
            try:
                r = page.goto(u, timeout=30000, wait_until="load")
                m = page.evaluate("()=>({iw:window.innerWidth,sw:document.scrollingElement.scrollWidth,bw:document.body.scrollWidth})")
                over = max(m["sw"], m["bw"]) - m["iw"]
                if r.status != 200 or over > 2 or errs:
                    bad.append({"url": u, "status": r.status, "overflow": over, "errors": errs[:3]})
            except Exception as e:
                bad.append({"url": u, "status": "EXC", "overflow": 0, "errors": [str(e)[:150]]})
        browser.close()
    if server: server.shutdown()
    print(f"smoke_mobile: {len(urls)} pages at 390px, {len(bad)} failures")
    for b in bad: print("FAIL", json.dumps(b))
    return 1 if bad else 0

if __name__ == "__main__":
    raise SystemExit(main())
