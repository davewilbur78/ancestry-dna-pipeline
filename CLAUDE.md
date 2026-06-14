---
Ancestry DNA Match Pipeline CLAUDE.md
Version: 2.7
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

TWO TRACKS -- use the one that fits the session context:

TRACK A -- build_workbook.py (CLI, recommended for Claude Code sessions)
  Takes the Genealogy Assistant CSV and the API results JSON directly.
  No enriched XLSX intermediate required. Handles all formatting internally.
  Usage:
    python build_workbook.py \
      --api-json  /tmp/{First}_{Last}_dna_api_results.json \
      --input-csv /path/to/GA_export.csv \
      --kit-url   "https://www.ancestry.com/dna/matches/GUID.../list" \
      --first-name {First} --last-name {Last} --batch N \
      --output-dir /path/to/output/

TRACK B -- build_v2.2.py (CONFIG block, takes enriched XLSX as input)
  Used when a prior step has already produced an enriched XLSX with hyperlinks.
  Edit the CONFIG block at the top before running. No argparse.

Both tracks produce identical output: the same 4-tab, 15-column workbook.

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
├── fetch_shared_dna.py          -- API collection script (local fallback, browser_cookie3)
├── build_v2.2.py                -- workbook builder, CONFIG-block version (takes enriched XLSX)
├── build_workbook.py            -- workbook builder, CLI version v3.0 (takes CSV + API JSON)
├── CLAUDE_CODE_PIPELINE_PROMPT.md -- full Claude Code pipeline prompt (paste into Claude Code)
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
Python dependencies: openpyxl (required); requests, browser-cookie3 (fallback path only)
Cookie source: Chrome (browser_cookie3 default) -- local fallback only; primary path is in-browser fetch
API rate limiting: 150ms delay minimum between requests, concurrent batches of 50
Tester GUID: always passed as a parameter to the script, never hardcoded

### Build Script Versions

TWO builders, same output format:

build_workbook.py v3.0 (CLI, preferred for Claude Code)
  - argparse interface: --api-json, --input-csv, --kit-url, --first-name, --last-name,
    --batch, --output-dir
  - Takes Genealogy Assistant CSV + API results JSON directly
  - No enriched XLSX intermediate needed
  - Self-contained: builds hyperlinks, applies all formatting internally

build_v2.2.py (CONFIG block, takes enriched XLSX)
  - Edit CONFIG block at top before running
  - TESTER_NAME, BATCH_NUM, SOURCE_FILE, OUTPUT_DIR, SCRIPT_VERSION = "2.2"
  - Used when workflow has already produced an enriched XLSX

Both produce the same 4-tab, 15-column workbook to the same spec.

### Output Naming
Format: {FirstName}_{LastName}_{N}matches_{YYYYMMDD}_v{SCRIPT_VERSION}.xlsx
  - FirstName and LastName: REQUIRED, both always included
  - N: total match count (all rows, not just priority matches)
  - YYYYMMDD: today's date (build_workbook.py) or enriched XLSX modification date (build_v2.2.py)
  - SCRIPT_VERSION: from SCRIPT_VERSION constant in the builder

  Example: Lesley_Sterling_1000matches_20260612_v2.2.xlsx

REQUIRED: always include BOTH first and last name. Surname-only filenames are
forbidden -- a family shares one surname, so "Wilbur_..." is ambiguous across
multiple kits. Apply the same FirstName_LastName rule to any companion files
(e.g. Susan_Beyer_dna_api_results_Batch1.csv).

### Source XLSX Column Schema (enriched pipeline output, for build_v2.2.py)
The build_v2.2.py script reads the enriched XLSX produced by Steps 1-3. Expected headers:
  Match Name | Longest Segment | AScM | Unweighted cM | Segments | Weighted cM |
  Family Tree | Tree Size | Common Ancestor | Line Assignment | Groups | Notes |
  Match Side | GUID

The build script (v2.2+) reads by header name, not column index, so minor column
order variations are handled automatically. The source sheet is auto-detected (no
hardcoded sheet name).

---

## Session-Close Checklist

Before ending any productive session:
1. New decisions about the pipeline itself? Update this file, bump version, commit.
2. New files produced? Commit them (especially build scripts and the pipeline prompt).
3. Write a CHANGELOG entry.
4. If you keep optional notes for a kit, record what's next there.

## What To Work On Next Session

Generic project-level next steps live here. Optional per-kit next steps live in that
kit's notes file under `testers/`, if you keep one.

- Pipeline infrastructure is complete: build_workbook.py v3.0 and
  CLAUDE_CODE_PIPELINE_PROMPT.md v5.0 are committed and ready to use.
- Re-run any kits that were built with the buggy build_v2.1.py and produced bad output.
  Use build_workbook.py (or build_v2.2.py with correct CONFIG) for those.
- Cynthia Wilbur Top1500 was already rebuilt with v2.2 on 2026-06-14 and is clean.
