# AncestryDNA Match Pipeline -- CLAUDE CODE VERSION
# Prompt version: 5.0  |  Built from CLAUDE.md v2.6 + build_workbook.py v3.0
#
# ── IMPORTANT: THIS IS THE CLAUDE CODE PROMPT ──────────────────────────────
# This file is for use in Claude Code (the CLI tool / local agent).
# The Cowork version of this pipeline lives in the plugin skill at:
#   dna-match-extractor-plugin/skills/dna-match-extractor/SKILL.md
# They do the same job. Key differences:
#   • Claude Code uses `mcp__Claude_in_Chrome__javascript_tool` with a tabId
#   • Cowork uses `javascript_tool` (no tabId needed)
#   • Claude Code can also fall back to fetch_shared_dna.py + browser_cookie3
#   • Cowork cannot -- its bash shell is a sandboxed VM with no Chrome access
# Do not mix up guidance from the two versions.
# ───────────────────────────────────────────────────────────────────────────
#
# HOW TO USE:
#   1. Fill in the two bracketed values in the INPUTS section.
#   2. Make sure ancestry.com is open and you are logged in in Chrome.
#   3. Paste everything from the "---" line below into Claude Code.
# ───────────────────────────────────────────────────────────────────────────

---

## Session Initialization -- do this first, before anything else

Fetch the project brain and read it completely:

```
https://raw.githubusercontent.com/davewilbur78/ancestry-dna-pipeline/main/CLAUDE.md?cb=<replace-with-any-random-6-digit-number>
```

Confirm the version and date out loud. Version must be 2.6 or higher. If you see
"Version 1.0" or a "Current Tester" block, the cache is stale -- refetch with a
different cache-busting value before continuing.

Also confirm you are running as **local Claude Code** (not Cowork). The in-browser
fetch path is primary in both environments; the browser_cookie3 fallback is
available here and not in Cowork.

---

## Inputs

- **CSV:** `[FULL PATH TO GENEALOGY ASSISTANT CSV]`
- **Kit URL:** `[ANCESTRY KIT URL -- e.g. https://www.ancestry.com/dna/matches/XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX/list]`

**Input validation:**
- The CSV must be a Genealogy Assistant export (`.csv` file, not `.xlsx`).
  If the user provides an `.xlsx` file, stop and ask for the correct CSV before
  proceeding. An existing processed workbook is not a valid input.
- The kit URL must contain a UUID. Extract it and confirm it out loud before
  doing anything else.

**Tester name:**
Derive the tester's first and last name from the CSV filename
(e.g. `Cynthia Wilbur - DNA Matches - ...` → first=`Cynthia`, last=`Wilbur`).
If the filename does not make the name clear, ask the user before proceeding.
Both names are required for output file naming -- the pipeline cannot default to
surname-only.

---

## Repo and Dependencies

Project repo: `https://github.com/davewilbur78/ancestry-dna-pipeline`

Clone if scripts are not on disk:
```bash
git clone https://github.com/davewilbur78/ancestry-dna-pipeline.git
```

`build_workbook.py` is also at the same location as the input CSV (it ships with
the Ancestry DNA Data Extractor folder). Use whichever copy is on disk.

Install dependencies:
```bash
pip install openpyxl browser-cookie3 requests
```

---

## Step 1: Parse the CSV

Write and run a Python script. Open with `encoding="utf-8-sig"` (handles UTF-8 BOM).

Derive tester GUID from kit URL:
```python
re.search(r'/dna/matches/([A-F0-9-]{36})/', kit_url, re.I).group(1).upper()
```

Extract match GUIDs from each row's URL column:
```python
re.search(r'compare/([A-F0-9-]{36})/', url, re.I).group(1).upper()
```

De-duplicate GUIDs. Exclude the tester GUID if it appears in the match list.

Write all match GUIDs to `/tmp/{FirstName}_{LastName}_guids.json`
(a plain JSON array of uppercase UUID strings).

**Pre-flight report -- print this before touching Chrome:**
```
Matches loaded:   N
Tester GUID:      XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX
With trees:       M
With common anc:  K
GUIDs file:       /tmp/{FirstName}_{LastName}_guids.json
Output will be:   {FirstName}_{LastName}_{N}matches_{YYYYMMDD}_v3.0.xlsx
```

Confirm with the user that ancestry.com is open and logged in before proceeding
to Step 2.

---

## Step 2: API Data Collection

### Step 2A: Write the GUID injection files (Python -- before touching Chrome)

**Never inline GUIDs into a Chrome tool parameter text.** For any batch over
~50 matches, inline embedding requires manually typing UUIDs -- which causes
fabrication of non-existent GUIDs. Always write GUIDs to JS files on disk, then
read those files and pass their content to the Chrome tool.

```python
import json

first, last = "{FirstName}", "{LastName}"
name = f"{first}_{last}"
guids = json.load(open(f"/tmp/{name}_guids.json"))

half = len(guids) // 2
chunk1, chunk2 = guids[:half], guids[half:]

with open(f"/tmp/{name}_inject1.js", "w") as f:
    f.write(f"window.__realGuids = {json.dumps(chunk1)}; window.__realGuids.length;")

with open(f"/tmp/{name}_inject2.js", "w") as f:
    f.write(
        f"window.__realGuids = (window.__realGuids || []).concat({json.dumps(chunk2)}); "
        f"window.__realGuids.length;"
    )

print(f"Chunk 1: {len(chunk1)} GUIDs  ->  /tmp/{name}_inject1.js")
print(f"Chunk 2: {len(chunk2)} GUIDs  ->  /tmp/{name}_inject2.js")
print(f"Total:   {len(guids)} GUIDs ready for injection")
```

Confirm both files are written and sizes look right before proceeding.

### Step 2B: Load the Chrome tools via ToolSearch

The Chrome JS tool is deferred and must be loaded before it can be called:
```
ToolSearch query: "select:mcp__Claude_in_Chrome__javascript_tool,mcp__Claude_in_Chrome__tabs_context_mcp"
```

Then call `mcp__Claude_in_Chrome__tabs_context_mcp` with `createIfEmpty: false`
to get the active tab list. Identify the tabId for the ancestry.com tab. You will
pass this tabId to every subsequent `mcp__Claude_in_Chrome__javascript_tool` call.

**Do not use `mcp__Control_Chrome__execute_javascript`.** That tool silently fails
with "Chrome not running" even when the tab is visible. The only correct tool is
`mcp__Claude_in_Chrome__javascript_tool`.

### Step 2C: Inject GUIDs into the browser window (two calls)

Read inject1.js from disk via bash:
```bash
cat /tmp/{FirstName}_{LastName}_inject1.js
```
Pass that content to `mcp__Claude_in_Chrome__javascript_tool` with the tabId from
Step 2B. The return value should be a number (chunk 1 length). If it returns
`undefined` or throws, confirm the tab is on ancestry.com and logged in, then retry.

Read inject2.js and execute it the same way. The return value should be the total
GUID count. Confirm the count matches before proceeding.

### Step 2D: Fire the API collector

Execute the following JS via `mcp__Claude_in_Chrome__javascript_tool`. Replace
`{TESTER_GUID}` with the actual tester GUID. This script reads from
`window.__realGuids` -- it contains no hardcoded GUIDs:

```js
window.__dnaStatus = "running";
window.__dnaResults = [];
(function(){
  const TESTER = "{TESTER_GUID}";
  const GUIDS  = window.__realGuids;
  const base   = "https://www.ancestry.com/discoveryui-matches/parents/list/api/matchSharedDna/";
  const out    = [];
  async function one(m, attempt){
    try {
      const r = await fetch(base + TESTER + "/" + m, {credentials:"include"});
      const d = await r.json();
      return {guid:m,
              unweighted_cm:   d.totalSharedCentimorgans,
              longest_segment: d.longestSharedSegment,
              segments:        d.numSharedSegments};
    } catch(e){
      if(attempt < 2){ await new Promise(s=>setTimeout(s,500)); return one(m,attempt+1); }
      return {guid:m, unweighted_cm:"", longest_segment:"", segments:""};
    }
  }
  (async () => {
    for(let i=0; i<GUIDS.length; i+=50){
      const chunk = GUIDS.slice(i,i+50);
      const res   = await Promise.all(chunk.map(m=>one(m,1)));
      out.push(...res);
      await new Promise(s=>setTimeout(s,150));
    }
    window.__dnaResults = out;
    window.__dnaStatus  = "done";
  })();
})();
"started";
```

The script returns `"started"` immediately. Results accumulate asynchronously.

### Step 2E: Poll for completion

```js
(window.__dnaStatus || "not_started") + " | collected: " + (window.__dnaResults ? window.__dnaResults.length : "undefined")
```

Poll every 10-15 seconds until status is `"done"` and count matches total GUIDs.

**If `window.__dnaResults` is `undefined` on first poll:** the page navigated or
the window context reset. Re-run Steps 2C and 2D. The inject files are still on
disk -- no need to regenerate them.

**If more than 10% of calls fail after retry:** pause and alert the user. The
session likely expired or the tab navigated away. Ask them to confirm, then retry
failed GUIDs only. Never stop silently. Never invent data for a failed GUID.

### Step 2F: Read results back in chunks (respect the ~1 KB output cap)

Never read all results in one call. Slice into ~20-row chunks:
```js
JSON.stringify(window.__dnaResults.slice(0, 20))
// then slice(20, 40), slice(40, 60), etc.
```

Issue multiple slice calls in one `browser_batch` to retrieve all chunks in a
single round trip. Accumulate into Python dicts keyed by GUID. Write to
`/tmp/{FirstName}_{LastName}_dna_api_results.json` before proceeding.

---

## Step 2 Fallback: fetch_shared_dna.py (no Chrome session only)

Does not work in Cowork. Use only when no Chrome session is available.

```bash
python fetch_shared_dna.py \
  --kit-url "PASTE KIT URL" \
  --input-csv "PASTE CSV PATH" \
  --output "/tmp/{FirstName}_{LastName}_dna_api_results.csv"
```

The script enforces 150 ms delay and single retry. Warns if failure rate > 10%.

The fallback produces a CSV. Convert it to JSON before running build_workbook.py:
```python
import csv, json
rows = list(csv.DictReader(open("/tmp/{FirstName}_{LastName}_dna_api_results.csv")))
records = [{"guid": r["guid"],
            "unweighted_cm": r.get("unweighted_cm",""),
            "longest_segment": r.get("longest_segment",""),
            "segments": r.get("segments","")} for r in rows]
with open("/tmp/{FirstName}_{LastName}_dna_api_results.json","w") as f:
    json.dump(records, f)
```

---

## Step 3: Build the Workbook

Run `build_workbook.py` with the CSV and API JSON. The script handles all
formatting, column schema (15 columns), 4 tabs, color tiers, hyperlinks, compare
links, ⚠ flags, AScM formulas, row heights, auto-filter, and Start Here content.

**Locate build_workbook.py** -- it should be in the same folder as the input CSV
(the Ancestry DNA Data Extractor directory). If it's not there, fetch it from the
repo: `https://raw.githubusercontent.com/davewilbur78/ancestry-dna-pipeline/main/build_workbook.py`

```bash
python build_workbook.py \
  --api-json  "/tmp/{FirstName}_{LastName}_dna_api_results.json" \
  --input-csv "[FULL PATH TO GENEALOGY ASSISTANT CSV]" \
  --kit-url   "[KIT URL]" \
  --first-name {FirstName} \
  --last-name  {LastName} \
  --batch      {N} \
  --output-dir "[DIRECTORY TO SAVE WORKBOOK]"
```

The script prints a tier breakdown and the full output path when done.

**Output naming** (automatic -- do not override):
```
{FirstName}_{LastName}_{TotalMatches}matches_{YYYYMMDD}_v3.0.xlsx
```

Example: `Cynthia_Wilbur_1500matches_20260614_v3.0.xlsx`

**What the script produces:**

- **Start Here** tab: methodology explanation, quick start guide, column guide,
  "Working Between This File and Ancestry" section, Gil Bardige YouTube link
- **Priority Matches** tab: matches passing both filter criteria
  (longest >= 20 cM AND AScM >= 12)
- **Watch List** tab: matches passing exactly one criterion
- **All Matches** tab: all matches with batch summary stats

**15-column schema** (A-O):
Match Name | ⚠ | Longest Segment | AScM | Unweighted cM | Segments |
Weighted cM | Family Tree | Tree Size | Common Ancestor |
Family Line Assignment | Groups | Notes | Match Side | GUID (Reference Anchor)

**Tier colors by longest segment:**
- Dark green (70AD47, bold black text): 50+ cM -- high priority
- Med green (A9D18E): 30-49 cM -- solid signal
- Light green (E2EFDA): 20-29 cM -- worth investigating
- Pink (FFD7D7): fails one or both criteria

---

## Step 4: Deliver and Report

Present the completed workbook file.

Report:
1. Total matches processed
2. Tier breakdown: X dark green / Y med green / Z light green / N watch / N low
3. Any GUIDs where API data could not be retrieved (list them)
4. Full path to saved file

---

## Methodology rules -- never violate

- **Unweighted cM only.** TIMBER understates Ashkenazi values.
- **AScM (col D) is always a live formula**, never hardcoded.
- **15 columns only.** The schema is fixed; build_workbook.py enforces it.
- **Platform relationship estimates are meaningless** under Ashkenazi endogamy.
- **ThruLines / Common Ancestor are hypotheses**, not conclusions.
- **Leeds Method does not apply** to Ashkenazi data.
- **Anti-fabrication:** never invent cM values, GUIDs, or URLs. Missing data =
  blank cell + note in report. GUIDs are always written to disk first and read
  from disk -- never typed inline into a Chrome tool call.
- **Both names always required** in output filenames. Surname-only is forbidden.

---

*CLAUDE CODE version -- Cowork version: dna-match-extractor SKILL.md*
*CLAUDE.md v2.6 (2026-06-14) | Prompt v5.0 | build_workbook.py v3.0*
