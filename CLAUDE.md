---
Ancestry DNA Match Pipeline CLAUDE.md
Version: 2.2
Last updated: 2026-06-04 UTC
---

## Operating Model

This file is the single source of truth for the Ancestry DNA Match Pipeline project.
It lives at: https://raw.githubusercontent.com/davewilbur78/ancestry-dna-pipeline/main/CLAUDE.md

Fetch and read this file fully at the start of every session.
Then load the active tester config named under "Active Tester" below.
Never rely on memory from previous conversations.
Confirm the CLAUDE.md version, date, AND the active tester name out loud when loaded.

---

## What This Is

A repeatable, automated, kit-agnostic pipeline for enriching AncestryDNA match
exports with data that Ancestry does not include in its CSV downloads. The
pipeline takes a Genealogy Assistant CSV export and a kit URL as inputs, calls the
Ancestry API for missing DNA fields, collects tree and common ancestor hyperlinks
via browser automation, and produces a fully formatted Excel workbook optimized as
a working research tool.

The pipeline and methodology in this file are generic and apply to any tester.
Everything specific to one person -- kit ID, family lines, surname-to-quadrant
mapping, Ancestry group names, research priorities, and batch state -- lives in a
per-tester config file under `testers/`. To work a different kit, point the Active
Tester block below at a different config file.

Designed for Ashkenazi Jewish genetic genealogy research. All methodology follows
GPS standards and Ashkenazi-specific endogamy awareness.

---

## Active Tester

This project is kit-agnostic. The tester-specific configuration is loaded from a
separate file so the same pipeline can serve any kit.

ACTIVE TESTER: Jeannette Klein
CONFIG FILE: testers/jeannette-klein.md
CONFIG RAW URL: https://raw.githubusercontent.com/davewilbur78/ancestry-dna-pipeline/main/testers/jeannette-klein.md

After reading this CLAUDE.md, load the config file named above and read it fully.
It defines the active tester's kit ID, line anchors, surname-to-line mapping,
Ancestry groups, research priority order, and current batch state.

To switch kits: change the two lines above (ACTIVE TESTER and CONFIG FILE) to point
at the new tester's config, then load it.

To onboard a new kit: see "Adding a New Tester" below.

---

## Pipeline: What It Does

### Inputs
- Genealogy Assistant CSV export (any number of matches, any batch size)
- Tester's Ancestry kit URL (used to derive the tester GUID)

### Step 1: Parse and Extract
- Read the CSV, extract match GUIDs from the URL column
- Derive the tester GUID from the kit URL

### Step 2: API Data Collection

PRIMARY PATH (works in Cowork AND local Claude Code) -- in-browser fetch via Claude
in Chrome. Navigate the tab to ancestry.com (logged in), then call the GET endpoint
same-origin with credentials so the session cookie rides along; no cookie extraction.
  Endpoint (GET): https://www.ancestry.com/discoveryui-matches/parents/list/api/matchSharedDna/{TESTER_GUID}/{MATCH_GUID}
  Call shape: (async () => { const r = await fetch(url, {credentials:"include"}); return await r.json(); })()
  Note: wrap in async IIFE -- top-level await is rejected by the javascript_tool.
  Returns: totalSharedCentimorgans (unweighted), longestSharedSegment, numSharedSegments
- Batch 50 at a time using Promise.all; results returned as pipe-separated compact string
- The tester GUID is supplied as a parameter -- never hardcoded
- No hard limit on total matches

FALLBACK (local Claude Code ONLY) -- fetch_shared_dna.py with browser_cookie3 reads
Chrome cookies directly. This does NOT work in Cowork: the bash shell is a sandboxed
Linux VM with no access to the user's Chrome. Use only when running as local Claude Code.

### Step 3: Link Collection

DEFAULT -- construct verified compare URLs deterministically from the GUIDs (no
per-profile scraping; both patterns confirmed live on Ancestry):
  Profile/compare:        https://www.ancestry.com/dna/matches/{TESTER_GUID}/compare/{MATCH_GUID}
  Tree + ThruLines compare: https://www.ancestry.com/discoveryui-matches/compare/{TESTER_GUID}/with/{MATCH_GUID}

OPTIONAL deep-links enhancement -- a per-profile browser pass, run only over the
priority subset (matches passing longest >= 20 AND AScM >= 12), not all matches. The
bulk treeData/commonAncestors endpoints are POST-only, header-gated, and SPA-cached,
so replaying them is advanced and not required for a usable workbook.

### Step 4: Build Spreadsheet
- Column order: Match Name | Longest Segment | AScM | Unweighted cM | Segments |
  Weighted cM | Family Tree | Tree Size | Common Ancestor | Line Assignment |
  Groups | Notes | Match Side | GUID
- Match Name hyperlinks to match profile page
- Family Tree hyperlinks to actual tree
- Common Ancestor hyperlinks to ThruLines page
- AScM = IFERROR(D/E, "") -- live Excel formula, not hardcoded
- Color coding by tier (based on longest segment), generic for any tester:
    RED:         longest < 20 OR AScM < 12 (fails filter)
    LIGHT GREEN: longest 20-30 (passes, investigate)
    MED GREEN:   longest 30-50 (solid signal)
    DARK GREEN:  longest 50+ (high priority, bold black text on #70AD47)
- Line Assignment column color coded by grandparent quadrant. Quadrant colors are a
  fixed project convention (cool = paternal, warm = maternal); the surname behind
  each quadrant comes from the active tester config:
    PP (paternal-paternal):   #E2EFDA light green
    PM (paternal-maternal):   #BDD7EE light blue
    MP (maternal-paternal):   #FFEB9C light yellow
    MM (maternal-maternal):   #FCE4D6 light coral
    Multiple:                 #E2CEEF light purple

---

## AScM Filter and Thresholds

AScM = Unweighted cM / Number of Segments

FILTER: AScM >= 12 AND Longest Segment >= 20 cM
(Both conditions must be met to pass. Matches failing either are flagged red.)

TIER BASIS: Longest segment (not AScM -- distribution is too compressed for AScM tiers)
  Why: AScM in a typical Ashkenazi batch clusters tightly between 12-20.
  Longest segment has better spread and maps directly to traceability research.

KEY RESEARCH BASIS:
  Kitty Cooper: longest segment > 20 cM minimum for recent Ashkenazi relationship
  Jennifer Mendelsohn: no segment > 20 cM = likely untraceable
  Jennifer Mendelsohn: two segments > 30 cM = approx 3.3 generations to MRCA
  Adina Newman, Lara Diamond, Gil Bardige: focus on longest segment + AScM together
  Center for Jewish History: 100 cM total minimum threshold for 100% Ashkenazi testers

Note: Always use UNWEIGHTED cM. Never TIMBER-adjusted figures for Ashkenazi analysis.
Ancestry's TIMBER algorithm is calibrated for non-endogamous populations and
systematically understates Ashkenazi match values.

---

## DNA Groups System (Ancestry)

Ancestry supports up to 64 groups. Each match can hold multiple group tags.

TIER 1 (couple/line tag): lighter shade -- match placed on a line but spouse not confirmed
TIER 2 (individual confirmed): darker shade of same family -- specific spouse's line confirmed

Color assignment follows cool = paternal, warm = maternal convention:
  PP: green family
  PM: blue family
  MP: yellow/gold family
  MM: coral/red family

Dots display in Ancestry in a fixed palette order regardless of assignment sequence.
Plan group creation order accordingly. The specific group names in use for the
active tester are listed in that tester's config file.

---

## Methodology Rules (GPS + Ashkenazi)

- DNA evidence never stands alone -- always correlate with documentary evidence
- Platform relationship estimates are meaningless -- never use them in analysis
- Leeds Method does not apply to Ashkenazi data -- use spine/anchor methodology
- ThruLines and Theory of Family Relativity are hypotheses only, never conclusions
- "Common Ancestor" in Ancestry output = a strong clue, not a confirmed fact
- "Multiple" line assignment overrides all tiers -- held for separate investigation
- Anti-fabrication: never invent cM values, relationship conclusions, or sources

---

## Batch Management

Batches are numbered sequentially per tester. The current batch state for the active
tester (which batches exist, row counts, what is complete vs outstanding) is recorded
in that tester's config file under `testers/`, not here.

When adding batches: append to the same workbook as new sheets, OR create new
batch files and note them in the tester config. Do not overwrite prior batches.

---

## Adding a New Tester

1. Copy `testers/_TEMPLATE.md` to `testers/{firstname-lastname}.md`.
2. Fill in the new tester's kit ID, line anchors, surname-to-line mapping, known
   great-grandparent couples, Ancestry group names, and research priorities.
3. Update the Active Tester block in this CLAUDE.md to point at the new config file
   (both ACTIVE TESTER and CONFIG FILE / CONFIG RAW URL).
4. Provide the new kit's Genealogy Assistant CSV export and Ancestry kit URL.
5. Run the pipeline. The tester GUID is taken from the kit URL; nothing is hardcoded.

Prior testers' config files stay in `testers/` and remain reusable. Switching back
is just a matter of repointing the Active Tester block.

---

## Files in This Repo

```
ancestry-dna-pipeline/
├── CLAUDE.md                    -- this file, generic project brain
├── CHANGELOG.md                 -- session log
├── README.md                    -- human overview
├── fetch_shared_dna.py          -- parameterized API collection script (local fallback)
├── testers/
│   ├── _TEMPLATE.md             -- blank per-tester config template
│   ├── adrienne-peckler.md      -- Adrienne Balsky Peckler config
│   └── jeannette-klein.md       -- Jeannette Klein config (active)
├── docs/
│   ├── column-schema.md         -- full column spec with rationale
│   ├── threshold-research.md    -- AScM/longest segment research notes
│   └── PIPELINE_DEBRIEF_Cowork_run.md -- first Cowork run debrief (why Step 2 changed)
└── dna-match-extractor-plugin/  -- Cowork plugin
    ├── .claude-plugin/
    │   └── plugin.json
    ├── skills/
    │   ├── dna-match-extractor/
    │   │   └── SKILL.md
    │   └── ashkenazi-analyst/
    │       └── SKILL.md
    └── README.md
```

---

## Claude Code Conventions

Working directory: wherever the tester's files are stored locally.
Python dependencies: requests, browser-cookie3, openpyxl, pandas
Cookie source: Chrome (browser_cookie3 default) -- local fallback only; primary path is in-browser fetch
API rate limiting: 150ms delay minimum between requests, concurrent batches of 50
Tester GUID: always passed as a parameter to the script, never hardcoded
Output naming: {Tester_LastName}_DNA_Matches_Batch{N}.xlsx

---

## Session-Close Checklist

Before ending any productive session:
1. New decisions made? Update this file (or the tester config), bump version, commit.
2. New files produced? Commit them.
3. Write a CHANGELOG entry.
4. Note what's next in the active tester config under "What's Next."

## What To Work On Next Session

Generic project-level next steps live here; per-tester next steps live in each
tester's config file. For the current active tester, see the "What's Next" section
of `testers/jeannette-klein.md`.
