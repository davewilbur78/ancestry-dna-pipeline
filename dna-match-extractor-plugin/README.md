# DNA Match Extractor Plugin

Cowork plugin for enriching AncestryDNA Genealogy Assistant CSV exports.

## Skills

### dna-match-extractor
The main pipeline. Runs on a CSV export + Ancestry kit URL. Handles any
number of matches. Calls the Ancestry API, collects browser links, builds
the enriched Excel workbook.

Say: "process my matches", "run the DNA pipeline", "I have a new batch" --
or just upload a CSV and provide your kit URL.

### ashkenazi-analyst
Ashkenazi-specific analysis context. Loads endogamy-aware methodology,
line assignment logic, research priority ordering, and threshold guidance.
Auto-loads when working with Ashkenazi tester data.

## Requirements

- Claude Code installed and accessible
- Chrome open and logged into Ancestry
- Python packages: requests, browser-cookie3, openpyxl, pandas

## Version

0.1.0 -- initial build, 2026-06-04
