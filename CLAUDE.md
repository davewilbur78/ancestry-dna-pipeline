---
Ancestry DNA Match Pipeline CLAUDE.md
Version: 2.3
Last updated: 2026-06-05 UTC
---

## Operating Model

This file is the single source of truth for the Ancestry DNA Match Pipeline project.
It lives at: https://raw.githubusercontent.com/davewilbur78/ancestry-dna-pipeline/main/CLAUDE.md

Fetch and read this file fully at the start of every session.
Then load the tester config for the kit being worked (see "Active Tester" below);
if no kit is named, ask which one to load rather than assuming a default.
Never rely on memory from previous conversations.
Confirm the CLAUDE.md version and date out loud when loaded, plus the active tester
name once a kit is loaded (or note that no tester is loaded yet).

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
per-tester config file under `testers/`. To work a kit, name it at the start of the
session so its config loads; there is no default tester.

Designed for Ashkenazi Jewish genetic genealogy research. All methodology follows
GPS standards and Ashkenazi-specific endogamy awareness.

---

## Active Tester

This project is kit-agnostic. No tester is loaded by default. Each tester's
configuration lives in its own file under `testers/` so the same pipeline serves
any kit.

ACTIVE TESTER: (none set)

At the start of each session, after reading this file:
- If the user names a kit, load its config from `testers/{firstname-lastname}.md`
  and read it fully. It defines that tester's kit ID, line anchors, surname-to-line
  mapping, Ancestry groups, research priority order, and current batch state.
- If no config exists for that kit yet, create one from `testers/_TEMPLATE.md`
  (see "Adding a New Tester").
- If the user has not said which kit, ask which tester to load -- or list the
  configs present in `testers/`. Do not assume a default.

Existing tester configs in `testers/` are reusable; selecting one is just naming it.
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
  Note: wrap in an async IIFE -- the javascript_tool rejects top-level await. Run 50
  per batch with Promise.all; read results back compactly (pipe-separated, or ~20-row
  JSON chunks) to stay under the ~1KB output cap. Confirmed at 300-match scale.
  Returns: totalSharedCentimorgans (unweighted), longestSharedSegment, numSharedSegments
- Batch 50 at a time, 150ms between batches, retry once on failure
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
3. Name the kit at the start of the session so its config loads. (There is no
   default tester; CLAUDE.md's Active Tester is "(none set)".)
4. Provide the new kit's Genealogy Assistant CSV export and Ancestry kit URL.
5. Run the pipeline. The tester GUID is taken from the kit URL; nothing is hardcoded.

Prior testers' config files stay in `testers/` and remain reusable. Switching kits
is just a matter of naming a different one.

---

## Files in This Repo

```
ancestry-dna-pipeline/
├── CLAUDE.md                    -- this file, generic project brain
├── CHANGELOG.md                 -- session log
├── README.md                    -- human overview
├── fetch_shared_dna.py          -- parameterized API collection script
├── testers/
│   ├── _TEMPLATE.md             -- blank per-tester config template
│   ├── adrienne-peckler.md      -- tester config (Adrienne Balsky Peckler)
│   ├── jeannette-klein.md       -- tester config (Jeannette Klein)
│   └── cynthia-wilbur.md        -- tester config (Cynthia (Klein) Wilbur)
├── docs/
│   ├── column-schema.md         -- full column spec with rationale
│   ├── threshold-research.md    -- AScM/longest segment research notes
│   └── PIPELINE_DEBRIEF_Cowork_run.md -- first Cowork run debrief
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
Output naming: {FirstName}_{LastName}_DNA_Matches_Batch{N}.xlsx
  REQUIRED: always include BOTH first and last name. Surname-only filenames are
  forbidden -- a family shares one surname, so "Wilbur_..." is ambiguous across
  multiple testers. Example: Cynthia_Wilbur_DNA_Matches_Batch1.xlsx. Apply the same
  FirstName_LastName rule to any companion files (e.g. _dna_api_results_BatchN.csv).

---

## Session-Close Checklist

Before ending any productive session:
1. New decisions made? Update this file (or the tester config), bump version, commit.
2. New files produced? Commit them.
3. Write a CHANGELOG entry.
4. Note what's next in the active tester config under "What's Next."

## What To Work On Next Session

Generic project-level next steps live here; per-tester next steps live in each
tester's config file under `testers/`, in that tester's "What's Next" section.
