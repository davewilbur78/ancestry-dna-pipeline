---
Ancestry DNA Match Pipeline CLAUDE.md
Version: 1.0
Last updated: 2026-06-04 UTC
---

## Operating Model

This file is the single source of truth for the Ancestry DNA Match Pipeline project.
It lives at: https://raw.githubusercontent.com/davewilbur78/ancestry-dna-pipeline/main/CLAUDE.md

Fetch and read this file fully at the start of every session.
Never rely on memory from previous conversations.
Confirm the version and date out loud when loaded.

---

## What This Is

A repeatable, automated pipeline for enriching AncestryDNA match exports with data
that Ancestry does not include in its CSV downloads. The pipeline takes a Genealogy
Assistant CSV export and a kit URL as inputs, calls the Ancestry API for missing DNA
fields, collects tree and common ancestor hyperlinks via browser automation, and
produces a fully formatted Excel workbook optimized as a working research tool.

Designed for Ashkenazi Jewish genetic genealogy research. All methodology follows
GPS standards and Ashkenazi-specific endogamy awareness. The primary researcher
is the owner of this repo.

---

## Current Tester

**Adrienne Balsky Peckler**
Kit ID: 4FB3190E-37B7-401F-A8CF-431950413661
Platform: AncestryDNA

### Configuration Block

TESTER: Adrienne Balsky Peckler
ANCHOR: [to be confirmed -- Marvin Balsky or Harriette Mendick when available]

LINE ANCHORS:
  PP - Paternal-Paternal (Green):  Balsky line
  PM - Paternal-Maternal (Blue):   Singer/Springer line
  MP - Maternal-Paternal (Yellow): Mendick line
  MM - Maternal-Maternal (Coral):  Weinberger/Danko line

GREAT-GRANDPARENT COUPLES (8-couple tier system, partially known):
  PP1: Isaac Balsky + Hhabena Balsky
  PP2: [Giberman/Lieberman couple -- Sarah Lieberman Giberman's parents]
  PM1: Samuel Singer + Minnie Jacobs
  PM2: Tzvi Dov Springer + unknown wife
  MP1: Harry Mendick + Lena Minster
  MP2: [to be determined]
  MM1: Adolph Weinberger + [unknown]
  MM2: Elsie Danko's family

PLATFORMS IN USE: AncestryDNA (primary). GEDmatch, FTDNA, MyHeritage to be confirmed.

ANCESTRY GROUPS ACTIVE:
  "Balsky" -- PP line (green)
  "SINGER/SPRINGER" -- PM line (blue)
  Mendick and Weinberger/Danko groups not yet created in Ancestry

---

## Pipeline: What It Does

### Inputs
- Genealogy Assistant CSV export (any number of matches, any batch size)
- Tester's Ancestry kit URL (used to derive the tester GUID)

### Step 1: Parse and Extract
- Read the CSV, extract match GUIDs from the URL column
- Derive the tester GUID from the kit URL

### Step 2: API Data Collection (Claude Code)
- Call Ancestry's matchSharedDna API for every GUID
  Endpoint: https://www.ancestry.com/discoveryui-matches/parents/list/api/matchSharedDna/{TESTER_GUID}/{MATCH_GUID}
  Returns: totalSharedCentimorgans (unweighted), longestSharedSegment, numSharedSegments, sharedCentimorgans (weighted)
- Uses browser_cookie3 to authenticate via Chrome session cookies
- Runs in concurrent batches (default 50 at a time) with 150ms delay between requests
- Handles errors gracefully, retries once on failure, logs any remaining failures
- No hard limit on total matches -- process in batches of any size

### Step 3: Browser Link Collection (Claude in Chrome)
- For each match with a linked tree: capture the direct tree URL
- For each match showing "Common Ancestor": capture the ThruLines/common ancestor URL
- Runs after API collection, same session

### Step 4: Build Spreadsheet
- Column order: Match Name | Longest Segment | AScM | Unweighted cM | Segments |
  Weighted cM | Family Tree | Tree Size | Common Ancestor | Line Assignment |
  Groups | Notes | Match Side | GUID
- Match Name hyperlinks to match profile page
- Family Tree hyperlinks to actual tree
- Common Ancestor hyperlinks to ThruLines page
- AScM = IFERROR(D/E, "") -- live Excel formula, not hardcoded
- Color coding by tier (based on longest segment):
    RED:         longest < 20 OR AScM < 12 (fails filter)
    LIGHT GREEN: longest 20-30 (passes, investigate)
    MED GREEN:   longest 30-50 (solid signal)
    DARK GREEN:  longest 50+ (high priority, bold black text on #70AD47)
- Line Assignment column color coded by grandparent line:
    PP (Balsky):           #E2EFDA light green
    PM (Singer/Springer):  #BDD7EE light blue
    MP (Mendick):          #FFEB9C light yellow
    MM (Weinberger/Danko): #FCE4D6 light coral
    Multiple:              #E2CEEF light purple

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
Plan group creation order accordingly.

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

Batches are numbered sequentially per tester. Current state:
  Batch 1: 150 matches (rows 1-150). Spreadsheet: Adrienne_Peckler_DNA_Matches_Batch1.xlsx
  Status: API data complete. Tree/CA hyperlinks outstanding.

When adding batches: append to the same workbook as new sheets, OR create new
batch files and note them here. Do not overwrite prior batches.

---

## Files in This Repo

```
ancestry-dna-pipeline/
├── CLAUDE.md                    -- this file, project brain
├── CHANGELOG.md                 -- session log
├── README.md                    -- human overview
├── docs/
│   ├── column-schema.md         -- full column spec with rationale
│   └── threshold-research.md    -- AScM/longest segment research notes
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
Cookie source: Chrome (browser_cookie3 default)
API rate limiting: 150ms delay minimum between requests, concurrent batches of 50
Output naming: {Tester_LastName}_DNA_Matches_Batch{N}.xlsx

---

## Session-Close Checklist

Before ending any productive session:
1. New decisions made? Update this file, bump version, commit.
2. New files produced? Commit them.
3. Write a CHANGELOG entry.
4. Note what's next under "What To Work On Next Session."

## What To Work On Next Session

- Browser pass for Adrienne Batch 1: collect tree URLs and common ancestor URLs
- Build Batch 2 (next 150 matches) using the Cowork plugin once it exists
- Confirm Anchor kit availability (Marvin Balsky or Harriette Mendick)
- Set up Mendick and Weinberger/Danko groups in Ancestry
