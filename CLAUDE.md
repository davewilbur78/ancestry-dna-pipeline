---
Ancestry DNA Match Pipeline CLAUDE.md
Version: 2.5
Last updated: 2026-06-14 UTC
---

## Operating Model

This file is the single source of truth for the Ancestry DNA Match Pipeline project.
It lives at: https://raw.githubusercontent.com/davewilbur78/ancestry-dna-pipeline/main/CLAUDE.md

Fetch and read this file fully at the start of every session.
To build a workbook you need only the kit's Genealogy Assistant CSV and its kit URL;
no per-tester config is required. (The files under `testers/` are optional notes.)
Never rely on memory from previous conversations.
Confirm the CLAUDE.md version and date out loud when loaded.

---

## What This Is

A repeatable, standard pipeline for enriching AncestryDNA match exports with data
that Ancestry does not include in its CSV downloads. The setup is the same for every
kit: give the pipeline a Genealogy Assistant CSV export and a kit URL, and it returns
a fully formatted Excel workbook -- DNA fields filled in from the Ancestry API,
verified compare links, color-coded by research priority.

That is the whole job. There is no per-kit configuration to load and no "active
tester" to set. Optional per-kit research notes (family lines, surname-to-branch
mapping, Ancestry groups, batch history) can live in a file under `testers/`, but
they are never required to produce a workbook.

Designed for Ashkenazi Jewish genetic genealogy research. All methodology follows
GPS standards and Ashkenazi-specific endogamy awareness.

---

## Standard Flow

1. User presents a Genealogy Assistant CSV export and the kit URL.
2. Run the pipeline (Steps 1-5 below). Same steps, same output schema, every kit.
3. Deliver the workbook.

The Family Line Assignment column (which branch a match sits on) is the only thing
that is ever kit-specific, and it is left blank by default. Fill it only when the
user supplies a surname-to-branch mapping -- inline, or from that kit's optional
notes file.

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

Run build_v2.1.py (or the current version) against the enriched input XLSX.

- Workbook structure: 4 tabs in order: Start Here | Priority Matches | Watch List | All Matches
- Column order (A-O): Match Name | Flag | Longest Segment | AScM | Unweighted cM |
  Segments | Weighted cM | Family Tree | Tree Size | Common Ancestor |
  Family Line Assignment | Groups | Notes | Match Side | GUID (Reference Anchor)
- Match Name: right-aligned, hyperlinked to profile comparison page, width 36
- Flag (col B): shows ⚠️ when a match has only 1 or 2 segments (scrutiny required)
- Family Tree: hyperlinked to tree compare view
- Common Ancestor: shows 👥 icon, hyperlinked to ThruLines page
- AScM = IFERROR(E/F, "") -- live Excel formula, not hardcoded
- Greyed columns (Weighted cM, Match Side, GUID): #707070 italic, visible on all tier backgrounds
- Color coding by tier (based on longest segment), the same for every kit:
    PINK:        does not meet one or both research criteria (longest < 20 OR AScM < 12)
    LIGHT GREEN: longest 20-29 cM (passes, worth investigating)
    MED GREEN:   longest 30-49 cM (solid signal)
    DARK GREEN:  longest 50+ cM (high priority, bold black text on #70AD47)
- Family Line Assignment column: left BLANK by default. Pre-fill it only when a
  surname-to-branch mapping is supplied (inline, or from the kit's optional notes).
  When filled, color by grandparent quadrant (fixed convention, cool = paternal,
  warm = maternal):
    PP (paternal-paternal):   #E2EFDA light green
    PM (paternal-maternal):   #BDD7EE light blue
    MP (maternal-paternal):   #FFEB9C light yellow
    MM (maternal-maternal):   #FCE4D6 light coral
    Multiple:                 #E2CEEF light purple

### Step 5: Deliver
- Save workbook using the auto-naming convention (see Output Naming below)
- Confirm tier breakdown: high / solid / investigate / watch / low counts

---

## AScM Filter and Thresholds

AScM = Unweighted cM / Number of Segments

FILTER: AScM >= 12 AND Longest Segment >= 20 cM
(Both conditions must be met to pass. Matches failing either are in the Watch List
or Low Priority category -- pink rows.)

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
This matters only when a user is actively tagging a kit's matches into groups; it has
no effect on the standard workbook build.

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

Workbooks are named per kit and per batch (see naming convention below). If you keep
an optional notes file for a kit, record its batch history there; otherwise just don't
overwrite a prior batch's workbook -- add the next batch as a new sheet or new file.

---

## Tester Configs (optional)

This pipeline is standard and uniform -- the same setup for every kit. Building a
workbook requires only two inputs: a Genealogy Assistant CSV export and the kit URL.
There is no "active tester" to set and no config to load first.

The files under `testers/` are optional research notes for kits where you want to
track family lines, a surname-to-branch (quadrant) mapping, Ancestry group names, or
batch history. Use one only if you want the Family Line Assignment column pre-filled,
or to record ongoing research on a specific kit. They never gate workbook creation.

To create a notes file: copy `testers/_TEMPLATE.md` to `testers/{firstname-lastname}.md`
and fill in what you know; leave the rest blank. Mention the kit by name in a session
if you want its notes loaded. Existing notes: adrienne-peckler.md, jeannette-klein.md,
cynthia-wilbur.md.

---

## Files in This Repo

```
ancestry-dna-pipeline/
├── CLAUDE.md                    -- this file, generic project brain
├── CHANGELOG.md                 -- session log
├── README.md                    -- human overview
├── fetch_shared_dna.py          -- parameterized API collection script (local fallback)
├── build_v2.1.py                -- workbook builder, current production version
├── testers/                     -- OPTIONAL per-kit research notes (never required)
│   ├── _TEMPLATE.md             -- blank notes template
│   ├── adrienne-peckler.md
│   ├── jeannette-klein.md
│   └── cynthia-wilbur.md
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

Working directory: wherever the kit's files are stored locally.
Python dependencies: requests, browser-cookie3, openpyxl, pandas
Cookie source: Chrome (browser_cookie3 default) -- local fallback only; primary path is in-browser fetch
API rate limiting: 150ms delay minimum between requests, concurrent batches of 50
Tester GUID: always passed as a parameter to the script, never hardcoded

### Build Script Versioning
The workbook builder is versioned: build_v{MAJOR.MINOR}.py
  MAJOR: significant structural changes (new tabs, schema changes)
  MINOR: design improvements, bug fixes, cosmetic changes
  Current production version: build_v2.1.py

### Output Naming
Format: {FirstName}_{LastName}_{N}matches_{YYYYMMDD}_v{SCRIPT_VERSION}.xlsx
  - FirstName and LastName: REQUIRED, both always included
  - N: total match count from the input file (all rows, not just priority matches)
  - YYYYMMDD: modification date of the enriched input XLSX (auto-derived by the script)
  - SCRIPT_VERSION: taken from SCRIPT_VERSION constant in the build script
  
  Example: Lesley_Sterling_1000matches_20260612_v2.1.xlsx

REQUIRED: always include BOTH first and last name. Surname-only filenames are
forbidden -- a family shares one surname, so "Wilbur_..." is ambiguous across
multiple kits. Apply the same FirstName_LastName rule to any companion files
(e.g. Susan_Beyer_dna_api_results_Batch1.csv).

### CONFIG Block (build script)
At the top of every build script, a clearly marked CONFIG block contains:
  TESTER_NAME    = "First Last"        # used in workbook header and output filename
  BATCH_NUM      = "Batch N"           # displayed in workbook subtitle
  SOURCE_FILE    = "/path/to/input.xlsx"
  OUTPUT_DIR     = "/path/to/output/folder"
  SCRIPT_VERSION = "2.1"               # drives output filename versioning

The extraction date and output filename are derived automatically; do not set them.

---

## Session-Close Checklist

Before ending any productive session:
1. New decisions about the pipeline itself? Update this file, bump version, commit.
2. New files produced? Commit them (especially the current build_v*.py).
3. Write a CHANGELOG entry.
4. If you keep optional notes for a kit, record what's next there.

## What To Work On Next Session

Generic project-level next steps live here. Optional per-kit next steps live in that
kit's notes file under `testers/`, if you keep one.

- GitHub MCP credentials need to be refreshed (token expired 2026-06-14; MCP returned
  "Bad credentials"). Reconnect GitHub in Settings > Connections before next commit.
