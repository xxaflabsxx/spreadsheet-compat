# canispreadsheet.com — distribution playbook

Status: site is comprehensive + fully on-page-optimized (215 tested funcs across 4 LO
versions, 206 verified recipes, 40 comparisons, checker+migration, error/pillar/
equivalents guides, open CC-BY dataset, structured data). The growth bottleneck is
AUTHORITY/backlinks, not content. This file makes distribution turnkey.

**GATE (important):** outward-facing posting from the AF Labs accounts is a brand-risk
decision reserved for Jon. Do NOT post to HN/Reddit/forums without his explicit
go-ahead. This file is prepared material, not a license to post. HN Show HN pitch:
see PROMO-showhn.md.

---

## Channel 1 — Show HN (best fit; needs an aged account)
Full pitch in PROMO-showhn.md. HN blocks Show HNs from brand-new accounts; the
`aflabs` account (created 2026-07-21) needs to age + ideally show some genuine
comment activity first. Best window: Tue–Thu ~13:00–15:00 UTC. Be present to answer.

## Channel 2 — Reddit (r/excel, r/googlesheets, r/spreadsheets, r/libreoffice)
Reddit hates self-promo. Rules vary; most require you to be an established participant
and forbid link-drops. The ONLY sustainable approach: genuinely help in comment
threads, linking the tool ONLY when it directly answers the question. A standalone
"I built this" post is allowed in some subs (r/excel has a "Show & Tell"/Discussion
flair; r/libreoffice is more tolerant) but only from an account with real history.

### Reddit "I made this" draft (use ONLY where the sub permits + from an aged account)
Title: I built a site that actually *runs* formulas to check Excel / Sheets /
LibreOffice compatibility (open dataset)

Body:
> I kept getting bitten by formulas that work in Excel but break in Google Sheets, or
> a "supported" function returning #NAME? in LibreOffice. So I made canispreadsheet.com.
> The part I care about: for LibreOffice, every result is produced by actually executing
> the formula in a headless LibreOffice and checking the output (with canary formulas
> proving it really recalculated) — not scraped from docs. It covers ~600 functions,
> 215 with live-run tests across four LibreOffice versions, plus a paste-a-formula
> checker with an Excel↔Sheets migration report, 206 verified how-to recipes, and the
> whole compatibility dataset is open under CC BY.
> Honest limitation: only LibreOffice is live-executed; Excel/Sheets verdicts are from
> official docs. Corrections very welcome — especially cases where the three disagree.
> No ads, no tracking.

## Channel 3 — Genuine help (highest-trust, lowest-risk, slowest)
Answer real questions on Reddit / r/excel / StackOverflow / MrExcel / spreadsheet
forums where the checker or a function/recipe page IS the best answer. Use the
shareable checker permalink so the reader sees the live result.

### Genuine-help reply template
> [Direct answer to their question first.] FWIW you can confirm this works in your app
> here — I ran it: https://canispreadsheet.com/checker.html#f=<url-encoded-formula>&t=g
> (that's a tool I built that actually executes formulas in LibreOffice; the target-app
> dropdown flags what breaks when you move between Excel/Sheets/LibreOffice).

Rules: only when genuinely the best answer; lead with the actual help; never copy-paste
the same link repeatedly; disclose it's your project. One good answer > ten link-drops.

## Channel 4 — Directory / resource listings (low effort, legit)
Submit the OPEN DATASET (not just the site) to places that list free datasets/tools:
- Google Dataset Search (already schema-tagged on /data.html — will get crawled)
- Awesome-lists on GitHub (awesome-spreadsheets, awesome-excel) via PR — genuinely
  useful addition, and GitHub PRs are a normal contribution, not spam.
- Free-tools / "no-signup tools" roundups where a submission form exists.

## Channel 5 — Answer-engine / LLM visibility
The open CC-BY dataset + clear methodology page position the site as a citable source.
Nothing to "post"; this compounds as the site is crawled. Keep the data accurate and
the licensing clear (done).

---

## Sequencing when greenlit
1. Show HN first (highest ceiling) once the account is credible — it's the one shot
   that can produce a spike + real backlinks.
2. In parallel, the GitHub awesome-list PR (Channel 4) — low-risk, legit, durable link.
3. Genuine-help (Channel 3) as an ongoing, low-volume habit — never spammy.
4. Reddit "I made this" only from an aged account, one sub at a time, honoring rules.

## Do-not
- No posting from brand-new accounts (filtered as spam, burns the domain's one launch).
- No repeated identical link-drops.
- No buying links / no PBNs / nothing that risks a manual action.
