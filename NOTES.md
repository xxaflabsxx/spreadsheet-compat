
## 2026-09-01 ~18:4x UTC — RESTART PREP (Jon adding a new MCP; s13 -> s14)
- IN FLIGHT AT RESTART: the Excel-web WRAP-UP + SITE-INTEGRATION agent (launched ~18:30 UTC) will be KILLED by the restart. Its task: (1) declare the 7 transport-unreachable fns (LAMBDA/LET/ISOMITTED/MAP/MAKEARRAY/REDUCE/SCAN) as excel_web skips w/ bisect evidence, (2) QA-classify INDIRECT->0 / TEXTBEFORE-TEXTAFTER if_not_found / BYCOL-BYROW blanks / ENCODEURL+FILTERXML+FORECAST.ETS web-absences, (3) full site integration per excel-web-site-plan.md (repo root): xwv/xwver compat keys, split Excel matrix row (desktop documented-only vs web executed 2026-09-01), honesty rules (narrow 1, extend 1b w/ xw slot, extend 5, NEW rule 7 anti-conflation w/ seeded proof), coverage section, ~24 build_site strings + 32 guides' "we do not run Excel" copy + 2 bespoke guide rewrites, DATASET_CARD/README. Verify: all guards + suites; list any verdict CHANGES (expect none); commit no push.
- ON RESUME: check the venture repo working tree FIRST — the killed agent may leave partial edits. Venture repo at kill-prep time: last commit "14468c71 Excel-web run: chunk-15 ingested — main run complete"; working tree then: DIRTY:  M DATASET_CARD.md
 M README.md
 M data/comparisons/average-vs-averagea.json
 M data/comparisons/char-vs-code-vs-unichar.json
 M data/comparisons/datedif-vs-yearfrac.json
 M data/comparisons/edate-vs-eomonth.json
 M data/comparisons/find-vs-search.json
 M data/comparisons/isblank-vs-empty-string.json. If dirty beyond that snapshot, the agent was mid-edit: git stash or reset to 14468c71-era HEAD, then RELAUNCH the agent with the paragraph above as its brief (all inputs durable: results/excel-web.json 579 fns COMMITTED, excel-web-site-plan.md, this entry).
- Also on resume (per runbook memory): re-register the cron from sprint/TICK-PROMPT.md (job 2724e0c8 dies with the session; prompt file is current, s13 queue with corpus-push COMPLETE + guides RESUMED), Gmail check, then gates: Sep 2 07:47 UTC ADS EARLY-STOP (1 view/0 clicks/$2.60 spent — decision + reasoning to log), Sep 3 marquee send (drafts ready: sprint/marquee-drafts-sep3.md) + GSC CTR read, Sep 4 ads EVAL, Sep 5 P3 Christmas + guide batch. Deferred idle-tick items: OneDrive file cleanup (~22 files, careful per-row deletes only), GSC re-requests for remaining unindexed guides.
- Everything else is committed + pushed: excel-web run data (14468c71), all NOTES entries, memories current (excel-web-execution-runbook has the full ops lessons incl. transport-unreachable class).
