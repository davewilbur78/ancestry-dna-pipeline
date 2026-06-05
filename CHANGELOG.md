# Changelog

## [2026-06-04] - v2.1 In-Browser Collection (Cowork fix)

### Context
First real Cowork run (Susan Beyer kit, 150 matches, 0 failures) succeeded but the
documented Step 2 did not run as written. browser_cookie3 cannot see the user's
Chrome from the sandboxed Linux VM. Full debrief: docs/PIPELINE_DEBRIEF_Cowork_run.md.

### Decisions
- In-browser fetch (Claude in Chrome, GET matchSharedDna with credentials:"include")
  is now the PRIMARY collection path. Works in both Cowork and local Claude Code.
- fetch_shared_dna.py + browser_cookie3 demoted to a local-Claude-Code-only fallback.
- Step 3 link collection: deterministic verified compare-URL construction is now the
  default (no per-profile scraping). Per-profile deep-link pass is optional, priority
  subset only. treeData/commonAncestors are POST-only/header-gated/SPA-cached -- noted.
- javascript_tool guardrails baked into the skill: no top-level await (poll a window
  flag instead), ~1KB output cap (read results in ~20-row chunks), browser and sandbox
  are separate filesystems (return data as tool output, not a shared CSV).
- Added an environment check at the top of the skill (Cowork vs local).

### Built / Produced
- CLAUDE.md v2.1: Step 2 and Step 3 rewritten per the above
- Plugin v0.3.0: dna-match-extractor SKILL.md v2.1 with the hardened in-browser collector
- docs/PIPELINE_DEBRIEF_Cowork_run.md (the Cowork run debrief, archived)

### Next
- Reinstall the rebuilt v0.3.0 plugin

## [2026-06-04] - v2.0 Kit-Agnostic Refactor

### Decisions
- Split the project into generic pipeline/methodology (CLAUDE.md) vs per-tester config (testers/)
- CLAUDE.md bumped to v2.0; added an Active Tester pointer and an "Adding a New Tester" flow
- Quadrant colors (PP green, PM blue, MP yellow, MM coral, Multiple purple) are a fixed project
  convention; surname-to-quadrant mapping now lives in the tester config, not in CLAUDE.md or skills
- Tester GUID is always a parameter -- never hardcoded anywhere

### Built / Produced
- testers/adrienne-peckler.md (archived Adrienne's full config + batch state + research priority)
- testers/_TEMPLATE.md (blank per-tester template for new kits)
- fetch_shared_dna.py parameterized: takes --kit-url/--tester-guid and --input-csv/--guids-file;
  extracts match GUIDs from the CSV URL column; retry-once and >10% failure warning built in
- Plugin v0.2.0: both SKILL.md files de-hardcoded (no more Adrienne surnames); they now read
  line names, colors, group mapping, and priorities from the active tester config
- README updated for multi-kit design

### Next
- Browser pass for Adrienne Batch 1 (tree + common ancestor URLs)
- Reinstall the rebuilt v0.2.0 plugin

## [2026-06-04] - Initial Pipeline Build: Adrienne Peckler Batch 1

### Decisions
- Use Unweighted cM (totalSharedCentimorgans) not TIMBER-adjusted for all Ashkenazi analysis
- AScM = Unweighted cM / Segments -- minimum threshold 12, longest segment minimum 20 cM
- Color tiers based on Longest Segment (not AScM -- distribution too compressed): 
  < 20 = red fail, 20-30 = light green, 30-50 = med green, 50+ = dark green (#70AD47 bold black)
- Line Assignment column in spreadsheet = working hypothesis; separate from Ancestry Groups tags
- Ancestry Groups = confirmed tagging in platform; Line Assignment = research working state
- 8-couple color tier system planned for Ancestry groups: Tier 1 (line placed), Tier 2 (spouse confirmed)
- Cool colors (blue/green) for paternal, warm (yellow/coral) for maternal
- Ancestry now supports 64 groups (not 24 as older guides state)
- GUID extracted from URL column; GUID column pushed to far right (col N)
- Family Tree and Common Ancestor columns need browser pass for actual hyperlinks (outstanding)
- Discovered Ancestry matchSharedDna API endpoint for bulk data collection

### Built / Produced
- Adrienne_Peckler_DNA_Matches_Batch1.xlsx (150 matches, fully enriched except hyperlinks)
- CLAUDE.md (this project brain)
- dna-match-extractor Cowork plugin (v0.1.0)
- ashkenazi-analyst skill

### Configuration Captured
- Tester: Adrienne Balsky Peckler, Kit 4FB3190E-37B7-401F-A8CF-431950413661
- Fan chart captured: all 4 grandparent lines and 8 great-grandparent couples identified
- Existing Ancestry groups: Balsky (PP), SINGER/SPRINGER (PM)
- Mendick (MP) and Weinberger/Danko (MM) groups not yet created in Ancestry

### Workflow Learned
- Genealogy Assistant CSV + Claude Code API script = fastest extraction path
- Browser character limit prevents reading large JS results back to Claude
- Solution: Claude Code writes results to CSV, user uploads to Claude chat
- For future batches: one Claude Code prompt collects all API data with no size limit
- Browser pass needed separately for tree URLs and common ancestor URLs

### Next
- Browser pass: collect tree URLs and CA URLs for Batch 1
- Build Batch 2 (next 150 or more matches) using the Cowork plugin
- Confirm anchor kit (Marvin Balsky or Harriette Mendick)
- Create Mendick and Weinberger/Danko groups in Ancestry
- Push this repo to GitHub and set up bootloader in Claude project
