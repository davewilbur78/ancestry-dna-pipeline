# Ancestry DNA Match Pipeline

A repeatable automated pipeline for enriching AncestryDNA match exports with
data not included in Genealogy Assistant CSV downloads. Designed for Ashkenazi
Jewish genetic genealogy research following GPS methodology.

## What It Does

Takes a Genealogy Assistant CSV export and an Ancestry kit URL. Automatically
collects missing DNA fields (unweighted cM, longest segment, segment count) via
the Ancestry API, collects tree and common ancestor hyperlinks via browser
automation, and produces a formatted Excel workbook optimized as a daily research tool.

Works for any number of matches. Processes in batches with no hard limit.

## Current Status

Active research for Adrienne Balsky Peckler.
Batch 1: 150 matches completed (API data populated, hyperlinks outstanding).

## How This Repo Is Organized

```
ancestry-dna-pipeline/
├── CLAUDE.md          -- AI instruction file. Load this at the start of every session.
├── CHANGELOG.md       -- Session-by-session decision log.
├── README.md          -- This file.
├── docs/              -- Reference documents and research notes.
└── dna-match-extractor-plugin/  -- Cowork plugin for running the pipeline.
```

## Using With AI

**CLAUDE.md is the project brain.** At the start of any AI session working on this
project, load CLAUDE.md first. It contains the full context: tester configuration,
methodology rules, column schema, threshold research, and current batch status.

Bootloader instruction for Claude projects:
```
At the start of every conversation, fetch and read this file fully:
[PASTE RAW GITHUB URL FOR CLAUDE.md HERE]
Read it completely before responding to anything. Never rely on memory.
```

## Using the Cowork Plugin

Install `dna-match-extractor-plugin` in Cowork. Then simply:
1. Upload your Genealogy Assistant CSV
2. Provide your Ancestry kit URL
3. The plugin handles everything else

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
