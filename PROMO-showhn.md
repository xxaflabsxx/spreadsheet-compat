# Show HN draft — fire at a weekday US-morning window (Tue-Thu ~13:00-15:00 UTC)

Refreshed 2026-08-26 05:50 UTC from the live corpus (numbers verified against data/ + results/ at that time — re-verify before posting; every example claim below was checked against results/*.json on 8/26 — COUNT literal-vs-reference, MROUND mismatched signs, XLOOKUP/SORT since 24.8.7.2, HSTACK/TEXTSPLIT/TAKE since 25.8.7.3):
600 functions in inventory; 278 functions / 841 test cases executed in LibreOffice 25.8.7.3 (168 in 24.2.0.3, 168 in 24.8.7.2, 194 in 25.2.0.3) AND, since 2026-08-29, the same 278/841 executed in Google Sheets via Drive import (dated run — Sheets has no version to pin; 5 functions inconclusive on import serialization); 282 verified how-to recipes; 48 comparisons; 15 executed-data guides; Migration Audit live ($19 launch, free tier = full scan + top-3 detail); offline companion CLI scripts/xlsx_recalc_diff.py (23 tests).

## Title (pick one, <80 chars)
- Show HN: Caniuse for spreadsheet functions – Sheets and LibreOffice results actually executed
- Show HN: Which Excel formulas silently change in Google Sheets and LibreOffice – tested, not documented

## Submission
- URL: https://canispreadsheet.com
- Post as: fresh account (create at news.ycombinator.com/login just before posting; username "aflabs").
- Immediately after submitting, add the body below as a first comment.
- Do NOT cite Reddit engagement as social proof (one engaged commenter there was an LLM — Jon, 8/26). The quikee_LO
  bug report (ex-LO dev, reproduced, fixed) is fair to mention only if asked how the data gets corrected.

## Body / first comment
I kept hitting the same wall: a formula works in Excel but breaks in Google Sheets, or a "supported" function
returns #NAME? or a *different number* in LibreOffice. Vendor docs say a function exists; they don't tell you
that COUNT over a cell holding TRUE is 0 in Excel and 1 in Calc (a literal TRUE argument counts in both), or
that MROUND(5,-2) is #NUM! in Excel and silently 6 in Calc.

So I built canispreadsheet.com — a compatibility DB for ~600 spreadsheet functions across Excel, Google Sheets
and LibreOffice Calc. The twist: the Google Sheets and LibreOffice results are *executed*, not scraped. A
headless LibreOffice writes each formula into a real workbook, recalculates, and I read the output back — with
deterministic and volatile canary formulas in every run to prove recalculation actually happened. The same
workbooks go through Google Drive's .xlsx import, which makes Sheets recalculate them, and come back out as
.xlsx for readback. 278 functions / 841 cases have live-run results in both engines, executed against four
LibreOffice releases (24.2, 24.8, 25.2, 25.8), so pages show caniuse-style "supported since" versions — e.g.
XLOOKUP and SORT landed in 24.8; HSTACK/TEXTSPLIT/TAKE in 25.8. Sheets gets a date, not a version: it is a
rolling service with nothing to pin.

Things to try:
- /checker — paste any formula; it extracts every function and says whether it works in each app, and can
  produce a migration report for a target app with verified alternatives.
- /audit.html — drop a whole .xlsx; it's parsed in the browser (never uploaded) and every formula gets a
  verdict against the executed dataset. Free scan; the full per-formula report is a paid one-off ($19 right now),
  which is how I'm trying to fund the execution work.
- /guides — 15 "same formula, different result" writeups, each backed by executed cases (COUNT and booleans,
  SUM("2") coercion, ERROR.TYPE codes, CHAR(0) producing a NUL, ...).
- /how-to — 282 common tasks with a formula that was executed and verified, not just documented.
- /compare — 48 head-to-heads (VLOOKUP vs XLOOKUP, SUMIF vs SUMIFS...).
- /data.html — the whole dataset is open (CC BY), plus the harness on GitHub.

One thing I learned the hard way that may save someone else a week: `soffice --headless --convert-to xlsx`
on a file saved by Excel does NOT recalculate — it passes Excel's cached values straight through, so a naive
"diff Excel vs LibreOffice" script reports that everything matches while evaluating nothing. You have to strip
the cached <v> values from formula cells first. There's a small offline CLI in the repo
(scripts/xlsx_recalc_diff.py) that does exactly that diff for your own workbooks, per cell, nothing uploaded.

Honest limitations: Excel is the one engine I still can't run — its column is Microsoft's documented behavior,
and it's the yardstick the two executed engines are measured against. Every guide labels which column is
executed and which is documented. The Sheets run reuses Excel-authored .xlsx files, and Google's importer
doesn't map every OOXML storage prefix, so five functions (FILTER, SORT, BYROW, BYCOL, MAKEARRAY) come back
*inconclusive* rather than getting a verdict I can't stand behind — a re-run with plain function names is
queued. The how-to recipe examples are still LibreOffice-only. If you find a case where the engines disagree
with my data, that's the most useful feedback I can get.

Static site, no tracking, no ads.

## Notes
- Be around to answer comments: check the thread every tick after posting (HN JSON API is scriptable:
  https://hacker-news.firebaseio.com/v0/item/<id>.json — no login needed for reading).
- If asked "how do you verify Excel/Sheets?": Sheets IS executed (Drive import, dated 2026-08-29 — show
  /methodology.html#sheets-caveats for the inconclusive set). Excel is not, and won't be until a Windows
  runner exists; docs only. Point at the executed/documented labels on every guide.
- If asked about the paid audit: free tier is real and useful; the paid part is the full report + Team tier
  ($79) for consultants. Don't oversell.
- If asked who/what built it: it's an AF Labs project; be honest about heavy automation/AI assistance if the
  question is direct — do not invent a team or a backstory.
- After HN, r/excel / r/googlesheets only as participation (their rules forbid promo posts — verified 8/24).
