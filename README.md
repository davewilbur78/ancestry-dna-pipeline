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

## Standard, Config-Free Flow

The pipeline is the same for every kit: present a Genealogy Assistant CSV and a kit
URL, get the formatted workbook. No per-kit configuration is required and there is no
"active tester" to set.

The files under `testers/` are OPTIONAL research notes -- used only when you want the
Line Assignment column pre-filled, or want to track a kit's family lines and batch
history. They never gate workbook creation.

## How This Repo Is Organized

```
ancestry-dna-pipeline/
├── CLAUDE.md          -- AI instruction file (generic). Load at the start of every session.
├── CHANGELOG.md       -- Session-by-session decision log.
├── README.md          -- This file.
├── fetch_shared_dna.py -- API collection script (local Claude Code fallback only).
├── testers/           -- OPTIONAL per-kit research notes (never required).
├── docs/              -- Reference documents and run notes.
└── dna-match-extractor-plugin/  -- Cowork plugin for running the pipeline.
```

## Using With AI

**CLAUDE.md is the project brain.** At the start of any AI session, load CLAUDE.md
first. It carries the full context: the standard pipeline, methodology, thresholds,
and naming rules. To build a workbook you then just provide a CSV and a kit URL.

Bootloader instruction for Claude projects:
```
At the start of every conversation, fetch and read this file fully:
https://raw.githubusercontent.com/davewilbur78/ancestry-dna-pipeline/main/CLAUDE.md
Read it completely before responding. Never rely on memory. To build a workbook you
need only a Genealogy Assistant CSV and the kit URL; no per-kit config is required.
```

## Data Collection

The primary collection path is in-browser fetch via Claude in Chrome: with a
logged-in Ancestry tab, the Ancestry API is called same-origin so the session
authenticates automatically. This works in both Cowork and local Claude Code.

A Python fallback (`fetch_shared_dna.py` with `browser-cookie3`) exists for local
Claude Code only; it reads Chrome cookies directly and does not work in Cowork,
where the shell is a sandboxed VM with no access to the browser. See
`docs/PIPELINE_DEBRIEF_Cowork_run.md` for the background.

## Optional: Per-Kit Notes

You never need this to build a workbook. Create a notes file only to pre-fill Line
Assignment or track a kit's research:

1. Copy `testers/_TEMPLATE.md` to `testers/{firstname-lastname}.md` and fill in what you know.
2. Mention the kit by name in a session if you want its notes loaded.

## Setup Requirements

- A logged-in Ancestry browser tab + Claude in Chrome (primary path)
- For the optional local fallback only: Python 3.9+ with requests, browser-cookie3,
  openpyxl, pandas, and Chrome logged into Ancestry

## Methodology

Ashkenazi Jewish genetic genealogy under GPS standards. Endogamy-aware throughout.
Never uses Leeds Method or platform relationship estimates. AScM filter: minimum 12.
Longest segment minimum: 20 cM. Tiers based on longest segment per Kitty Cooper,
Jennifer Mendelsohn, Adina Newman, Lara Diamond, and Gil Bardige's research.
