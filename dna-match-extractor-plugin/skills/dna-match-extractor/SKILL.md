---
name: dna-match-extractor
description: >
  Full pipeline for enriching an AncestryDNA Genealogy Assistant CSV export with
  data not available in the export. Trigger when user says: "run the DNA pipeline",
  "process my matches", "enrich my CSV", "build the match spreadsheet", "extract
  DNA data", "I have a new batch of matches", or provides a Genealogy Assistant
  CSV and an Ancestry kit URL. Handles any number of matches -- processes in
  concurrent batches of 50. Calls the Ancestry API via Claude Code, collects tree
  and common ancestor links via browser, and builds the enriched Excel workbook.
license: CC-BY-NC-SA-4.0
metadata:
  version: "1.0"
  author: User + Claude collaboration
  base_skills: gra v8.5c, ashkenazi-genetic-genealogist
---

# DNA Match Extractor

Full enrichment pipeline for AncestryDNA Genealogy Assistant CSV exports.
Takes a CSV and kit URL. Produces a formatted Excel workbook with all missing
fields populated and color-coded by research priority.

**Never fabricate cM values, segment data, or URLs. If a field cannot be
collected, leave it blank and note it. Never invent data.**

---

## Inputs Required

Before starting, confirm you have:

1. **Genealogy Assistant CSV export** -- uploaded by user. Any number of rows.
2. **Ancestry kit URL** -- the URL of the tester's DNA match list page, e.g.:
   `https://www.ancestry.com/dna/matches/{TESTER_GUID}/list`
   The tester GUID is extracted from this URL automatically.

If either is missing, ask for it before proceeding.

---

## Step 1: Parse Inputs

Extract from the CSV:
- Match Name
- Shared cM (Weighted/TIMBER-adjusted) -- column labeled "Shared cM"
- Family Tree text
- Tree Size
- Common Ancestor text
- Groups
- Notes
- Match Side
- URL column -- extract match GUID from each URL using pattern:
  `.../compare/([A-F0-9-]{36})/...`
- Build ProfileURL: strip `/shared-matches` and query string from URL

Extract tester GUID from kit URL:
  Pattern: `/dna/matches/([A-F0-9-]{36})/`

Report: N matches loaded, M with trees, K with common ancestors.

---

## Step 2: API Data Collection (Claude Code)

Write and run a Python script via Claude Code.

**Script requirements:**
- Use `browser_cookie3` to get Chrome cookies for `ancestry.com`
- Call this endpoint for every match GUID:
  `https://www.ancestry.com/discoveryui-matches/parents/list/api/matchSharedDna/{TESTER_GUID}/{MATCH_GUID}`
- Response fields to capture:
  - `totalSharedCentimorgans` → Unweighted cM
  - `longestSharedSegment` → Longest Segment
  - `numSharedSegments` → Segments
- Process in concurrent batches of 50 using asyncio or threading
- 150ms delay between batches (not between individual requests within a batch)
- Retry failed requests once after 500ms
- Log progress: print after every 50 matches
- Save results to `dna_api_results.csv` with columns: guid, unweighted_cm, longest_segment, segments
- Print final summary: N succeeded, M failed (list failed GUIDs)
- Save to the same directory as the input CSV

**Error handling:**
- If browser_cookie3 fails: instruct user to ensure Chrome is open and logged into Ancestry
- If >10% of requests fail: pause and alert the user before continuing
- Never stop silently on failure

---

## Step 3: Browser Link Collection (Claude in Chrome)

After API collection, collect hyperlinks that require browser navigation.

**For each match with a linked tree (Family Tree column is not "No trees"):**
- Navigate to the match profile page (ProfileURL)
- Find the tree link in the Trees tab
- Capture the direct tree URL
- Store as: tree_url for that GUID

**For each match showing "Common Ancestor":**
- Navigate to the match profile page
- Find the Common Ancestor link
- Capture the URL Ancestry uses for that link
- Store as: ca_url for that GUID

Save both sets to `dna_browser_links.csv` with columns: guid, tree_url, ca_url

If browser automation is not available in this session, note which links are
outstanding and skip to Step 4. The spreadsheet will be built with empty
hyperlink slots that can be filled in a future pass.

---

## Step 4: Build Spreadsheet

Merge all data sources. Build the workbook using openpyxl.

### Column Order (do not change without updating CLAUDE.md)

| Col | Header          | Source             | Format     |
|-----|-----------------|-------------------|------------|
| A   | Match Name      | CSV + ProfileURL   | Hyperlink  |
| B   | Longest Segment | API               | Integer    |
| C   | AScM            | Formula =D/E      | 0.0        |
| D   | Unweighted cM   | API               | Integer    |
| E   | Segments        | API               | Integer    |
| F   | Weighted cM     | CSV (Shared cM)   | Integer    |
| G   | Family Tree     | CSV + browser     | Hyperlink  |
| H   | Tree Size       | CSV               | Integer    |
| I   | Common Ancestor | CSV + browser     | Hyperlink  |
| J   | Line Assignment | Derived + user    | Color fill |
| K   | Groups          | CSV               | Text       |
| L   | Notes           | CSV               | Text       |
| M   | Match Side      | CSV               | Text       |
| N   | GUID            | Extracted         | Text       |

### Hyperlinks
- Match Name: links to ProfileURL (match profile page, not shared-matches tab)
- Family Tree: links to tree_url if collected, otherwise text only
- Common Ancestor: links to ca_url if collected, otherwise text only

### AScM Formula
Use Excel formula: `=IFERROR(D{row}/E{row},"")` -- never hardcode calculated values.

### Color Tiers (row-level, based on Longest Segment)

| Tier       | Condition                          | Fill    | Font             |
|------------|------------------------------------|---------|------------------|
| Fail       | AScM < 12 OR Longest Segment < 20  | FFD7D7  | CC0000 (red)     |
| Light green| Longest 20-30                      | E2EFDA  | 000000 (black)   |
| Med green  | Longest 30-50                      | A9D18E  | 000000 (black)   |
| Dark green | Longest 50+                        | 70AD47  | 000000 bold      |

### Line Assignment Colors (column J only)

| Line                    | Fill    |
|-------------------------|---------|
| PP - Balsky             | E2EFDA  |
| PM - Singer/Springer    | BDD7EE  |
| MP - Mendick            | FFEB9C  |
| MM - Weinberger/Danko   | FCE4D6  |
| Multiple                | E2CEEF  |

Line Assignment is pre-filled from Groups column where known.
User fills in remaining assignments as research progresses.

### Header Style
- Dark teal header row (#2F4858), white bold Arial 10
- Freeze row 1
- Row height 16, header height 20
- Alternating light gray (#F5F5F5) on non-colored rows

### Output Naming
`{Tester_LastName}_DNA_Matches_Batch{N}.xlsx`

Run recalc.py to verify zero formula errors before delivering.

---

## Step 5: Deliver and Report

Present the completed spreadsheet file.

Report:
- Total matches processed
- Tier breakdown: X dark green / Y med green / Z light green / N red
- Tree links collected: X of Y
- Common ancestor links collected: X of Y
- Any failures or outstanding items

---

## Batch Management

This pipeline is open-ended. Any number of matches can be processed.

For large lists (300+): run API collection in a single script (no limit).
For browser link collection: page through in sessions of 50-100 if needed.
For the workbook: each batch is a separate sheet or file; document in CLAUDE.md.

When adding a new batch for the same tester:
- Load CLAUDE.md from GitHub first
- Note the batch number (increment from last)
- Do not overwrite prior batch files

---

## Methodology Notes

Read `/skills/ashkenazi-analyst/SKILL.md` for full Ashkenazi DNA analysis context.
Key reminders:
- Always use UNWEIGHTED cM. Never TIMBER-adjusted figures.
- Platform relationship estimates are meaningless. Strip them.
- "Common Ancestor" in Ancestry = a clue. Verify before concluding.
- Multiple-line matches are held for investigation. Never force-assign.
