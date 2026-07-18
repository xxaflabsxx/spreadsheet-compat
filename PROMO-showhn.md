# Show HN draft — fire at a weekday US-morning window (Tue-Thu ~13:00-15:00 UTC)

## Title (pick one, <80 chars)
- Show HN: Caniuse for spreadsheet functions, every result actually executed
- Show HN: I machine-verify which formulas work in Excel, Sheets and LibreOffice

## Body / first comment
I kept hitting the same wall: a formula works in Excel but breaks in Google Sheets,
or a "supported" function returns #NAME? in LibreOffice. Vendor docs say a function
exists; they don't tell you it silently behaves differently.

So I built canispreadsheet.com — a compatibility DB for ~600 spreadsheet functions
across Excel, Google Sheets and LibreOffice Calc. The twist: every LibreOffice result
is *executed*, not scraped. A headless LibreOffice writes each formula into a real
workbook, recalculates, and I check the output — with deterministic + volatile canary
formulas each run to prove recalculation actually happened. Currently 168 functions
have live-run test cases (616 cases) on LibreOffice 25.8.

Two things people seem to like:
- /checker — paste any formula, it extracts every function and tells you if it works
  in each app (e.g. MAP works in Excel & Sheets but not LibreOffice yet).
- /how-to — 25 common tasks (sum by category, VLOOKUP-to-the-left, remove duplicates...)
  each with a copy-paste formula that's been executed and verified, not just documented.

Honest limitations: only LibreOffice is live-executed today; Excel/Sheets verdicts are
from their official function lists (I can't headlessly run those). The lambda-helpers
(MAP/REDUCE/SCAN/BYROW) are genuinely #NAME? in LibreOffice 25.8 — it has LAMBDA/LET
but not the helpers yet.

Static site, no tracking, no ads. Feedback welcome — especially edge cases where the
three engines disagree.

## Notes
- Be around to answer comments (check each tick after posting).
- After HN, consider r/excel / r/googlesheets only if HN validates it's genuinely useful.
