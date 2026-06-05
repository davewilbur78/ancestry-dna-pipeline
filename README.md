# Ancestry DNA Match Pipeline

A repeatable, kit-agnostic automated pipeline for enriching AncestryDNA match
exports with data not included in Genealogy Assistant CSV downloads. Designed for
Ashkenazi Jewish genetic genealogy research following GPS methodology.

## What It Does

Takes a Genealogy Assistant CSV export and an Ancestry kit URL. Collects the
missing DNA fields (unweighted cM, longest segment, segment count) by calling the
Ancestry API from inside your logged-in browser, builds verified compare links for
each match, and produces a formatted Excel workbook optimized as a daily research tool.

Works for any number of matches and any tester. Processes in batches with no hard limit.

## Kit-Agnostic Design

The pipeline and methodology are generic and live in `CLAUDE.md`. Everything
specific to one person -- kit ID, family lines, surname-to-quadrant mapping,
Ancestry group names, research priorities, and batch state -- lives in a per-tester
config file under `testers/`.

There is no default tester. Each session loads a kit by name from `testers/`, or
creates a new one from `testers/_TEMPLATE.md`. No kit is baked in.

## How This Repo Is Organized

```
ancestry-dna-pipeline/
├── CLAUDE.md          -- AI instruction file (generic). Load at the start of every session.
├── CHANGELOG.md       -- Session-by-session decision log.
├── README.md          -- This file.
├── fetch_shared_dna.py -- API collection script (local Claude Code fallback only).
├── testers/           -- Per-tester config files (_TEMPLATE.md plus one per kit).
├── docs/              -- Reference documents and run notes.
└── dna-match-extractor-plugin/  -- Cowork plugin for running the pipeline.
```

## Using With AI

**CLAUDE.md is the project brain.** At the start of any AI session, load CLAUDE.md
first, then load the tester config for whichever kit you name. Together they carry
the full context: generic pipeline, methodology, thresholds, plus that tester's
configuration and batch status.

Bootloader instruction for Claude projects:
```
At the start of every conversation, fetch and read this file fully:
https://raw.githubusercontent.com/davewilbur78/ancestry-dna-pipeline/main/CLAUDE.md
Then load the tester config for the kit being worked (or ask which one).
Read everything completely before responding. Never rely on memory.
```

## Data Collection

The primary collection path is in-browser fetch via Claude in Chrome: with a
logged-in Ancestry tab, the Ancestry API is called same-origin so the session
authenticates automatically. This works in both Cowork and local Claude Code.

A Python fallback (`fetch_shared_dna.py` with `browser-cookie3`) exists for local
Claude Code only; it reads Chrome cookies directly and does not work in Cowork,
where the shell is a sandboxed VM with no access to the browser. See
`docs/PIPELINE_DEBRIEF_Cowork_run.md` for the background.

## Adding a New Kit

1. Copy `testers/_TEMPLATE.md` to `testers/{firstname-lastname}.md` and fill it in.
2. Name the kit at the start of a session so its config loads.
3. Upload that kit's Genealogy Assistant CSV and provide its Ancestry kit URL.
4. Run the pipeline. The tester GUID comes from the kit URL; nothing is hardcoded.

## Setup Requirements

- A logged-in Ancestry browser tab + Claude in Chrome (primary path)
- For the optional local fallback only: Python 3.9+ with requests, browser-cookie3,
  openpyxl, pandas, and Chrome logged into Ancestry

## Methodology

Ashkenazi Jewish genetic genealogy under GPS standards. Endogamy-aware throughout.
Never uses Leeds Method or platform relationship estimates. AScM filter: minimum 12.
Longest segment minimum: 20 cM. Tiers based on longest segment per Kitty Cooper,
Jennifer Mendelsohn, Adina Newman, Lara Diamond, and Gil Bardige's research.
