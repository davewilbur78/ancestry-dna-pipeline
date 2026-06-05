# Ancestry DNA Match Pipeline

A repeatable, kit-agnostic automated pipeline for enriching AncestryDNA match
exports with data not included in Genealogy Assistant CSV downloads. Designed for
Ashkenazi Jewish genetic genealogy research following GPS methodology.

## What It Does

Takes a Genealogy Assistant CSV export and an Ancestry kit URL. Automatically
collects missing DNA fields (unweighted cM, longest segment, segment count) via
the Ancestry API, collects tree and common ancestor hyperlinks via browser
automation, and produces a formatted Excel workbook optimized as a daily research tool.

Works for any number of matches and any tester. Processes in batches with no hard limit.

## Kit-Agnostic Design

The pipeline and methodology are generic and live in `CLAUDE.md`. Everything
specific to one person -- kit ID, family lines, surname-to-quadrant mapping,
Ancestry group names, research priorities, and batch state -- lives in a per-tester
config file under `testers/`. To work a different kit, point the Active Tester
block in `CLAUDE.md` at a different config file. To add a kit, copy
`testers/_TEMPLATE.md`.

## Current Status

Active tester: Adrienne Balsky Peckler (`testers/adrienne-peckler.md`).
Batch 1: 150 matches completed (API data populated, hyperlinks outstanding).

## How This Repo Is Organized

```
ancestry-dna-pipeline/
├── CLAUDE.md          -- AI instruction file (generic). Load at the start of every session.
├── CHANGELOG.md       -- Session-by-session decision log.
├── README.md          -- This file.
├── fetch_shared_dna.py -- Parameterized API collection script (tester GUID is an argument).
├── testers/           -- Per-tester config files (_TEMPLATE.md plus one per kit).
├── docs/              -- Reference documents and research notes.
└── dna-match-extractor-plugin/  -- Cowork plugin for running the pipeline.
```

## Using With AI

**CLAUDE.md is the project brain.** At the start of any AI session, load CLAUDE.md
first, then load the active tester config it points to. Together they carry the full
context: generic pipeline, methodology, thresholds, plus the active tester's
configuration and batch status.

Bootloader instruction for Claude projects:
```
At the start of every conversation, fetch and read this file fully:
https://raw.githubusercontent.com/davewilbur78/ancestry-dna-pipeline/main/CLAUDE.md
Then load the active tester config it names under "Active Tester."
Read everything completely before responding. Never rely on memory.
```

## Adding a New Kit

1. Copy `testers/_TEMPLATE.md` to `testers/{firstname-lastname}.md` and fill it in.
2. Repoint the Active Tester block in `CLAUDE.md` at the new config.
3. Upload that kit's Genealogy Assistant CSV and provide its Ancestry kit URL.
4. Run the pipeline. The tester GUID comes from the kit URL; nothing is hardcoded.

## Setup Requirements

- Python 3.9+
- pip packages: requests, browser-cookie3, openpyxl, pandas
- Chrome browser open and logged into Ancestry
- Claude Code installed

## Methodology

Ashkenazi Jewish genetic genealogy under GPS standards. Endogamy-aware throughout.
Never uses Leeds Method or platform relationship estimates. AScM filter: minimum 12.
Longest segment minimum: 20 cM. Tiers based on longest segment per Kitty Cooper,
Jennifer Mendelsohn, Adina Newman, Lara Diamond, and Gil Bardige's research.
