# DNA Match Extractor Plugin

Cowork plugin for enriching AncestryDNA Genealogy Assistant CSV exports.

Standard and config-free: the same setup for every kit. Present a Genealogy Assistant
CSV and a kit URL, get the formatted workbook. Optional per-kit notes under `testers/`
only ever pre-fill the Line Assignment column.

## Skills

### dna-match-extractor
The main pipeline. Runs on a CSV export + Ancestry kit URL. Handles any
number of matches. Collects shared-DNA data via in-browser fetch (Claude in
Chrome), builds verified compare links, and produces the enriched Excel workbook.

Say: "process my matches", "run the DNA pipeline", "I have a new batch" --
or just upload a CSV and provide your kit URL.

### ashkenazi-analyst
Ashkenazi-specific analysis context. Loads endogamy-aware methodology,
line assignment logic, research priority ordering, and threshold guidance.
Auto-loads when working with Ashkenazi tester data.

## Requirements

- A logged-in Ancestry browser tab (for the primary in-browser collection path)
- Claude in Chrome connected
- For the optional local fallback only: Python with requests, browser-cookie3, openpyxl, pandas

## Version

0.3.2 -- standard config-free flow: CSV + kit URL in, workbook out, same for every kit.
Tester notes under testers/ are optional and only pre-fill Line Assignment. 2026-06-05
0.3.1 -- no default tester (fully kit-agnostic); FirstName_LastName output naming;
skill wording aligned to "no default." 2026-06-04
0.3.0 -- in-browser fetch is now the primary collection path (works in Cowork and
local Claude Code); browser_cookie3 demoted to a local-only fallback. 2026-06-04
0.2.0 -- kit-agnostic refactor, 2026-06-04
