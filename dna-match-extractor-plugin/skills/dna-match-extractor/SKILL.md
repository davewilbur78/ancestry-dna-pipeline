---
name: dna-match-extractor
description: >
  Full pipeline for enriching an AncestryDNA Genealogy Assistant CSV export with
  data not available in the export. Trigger when user says: "run the DNA pipeline",
  "process my matches", "enrich my CSV", "build the match spreadsheet", "extract
  DNA data", "I have a new batch of matches", or provides a Genealogy Assistant
  CSV and an Ancestry kit URL. Handles any number of matches -- processes in
  batches of 50. Collects shared-DNA data via in-browser fetch (Claude in Chrome),
  builds verified compare links, and produces the enriched Excel workbook.
license: CC-BY-NC-SA-4.0
metadata:
  version: "2.1"
  author: User + Claude collaboration
  base_skills: gra v8.5c, ashkenazi-genetic-genealogist
---

# DNA Match Extractor

Full enrichment pipeline for AncestryDNA Genealogy Assistant CSV exports.
Takes a CSV and kit URL. Produces a formatted Excel workbook with all missing
fields populated and color-coded by research priority.

This skill is kit-agnostic. All tester-specific values -- kit ID, family surnames,
the surname-to-line (quadrant) mapping, Ancestry group names, and research
priorities -- come from the active tester config loaded via CLAUDE.md
(`testers/{tester}.md`). Never hardcode a tester or a surname here.

**Never fabricate cM values, segment data, or URLs. If a field cannot be
collected, leave it blank and note it. Never invent data.**

---

## Environment Check (do this first)

State which mode you are in and pick the matching collection path up front:

- **Cowork (Claude desktop app):** the bash shell is a sandboxed Linux VM with no
  access to the user's Chrome. `browser_cookie3` returns nothing. You MUST collect
  via in-browser fetch (Step 2 primary path). The browser and the sandbox are
  separate machines with separate filesystems -- data fetched in the browser comes
  back as tool-output text, not via a shared CSV on disk.
- **Local Claude Code:** the in-browser path still works and is preferred. The
  `browser_cookie3` Python fallback is available here if a browser session is not.

When in doubt, use the in-browser path -- it works in both.

---

## Inputs Required

Before starting, confirm you have:

1. **Genealogy Assistant CSV export** -- uploaded by user. Any number of rows.
2. **Ancestry kit URL** -- the URL of the tester's DNA match list page, e.g.:
   `https://www.ancestry.com/dna/matches/{TESTER_GUID}/list`
   The tester GUID is extracted from this URL automatically.
3. **Active tester config** -- confirm CLAUDE.md is loaded and the active tester
   config under `testers/` has been read. It supplies the line/surname mapping
   and group names used in Step 4.
4. **A logged-in Ancestry browser tab** (for the in-browser collection path).

If any of these is missing, ask for it before proceeding.

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

## Step 2: API Data Collection

### PRIMARY PATH -- in-browser fetch (Claude in Chrome). Works in Cowork AND local.

The match-shared-DNA endpoint is a GET and authenticates from the logged-in session
with `credentials:"include"`. No cookie extraction, no Python, no auth handling.

Navigate the Chrome tab to `https://www.ancestry.com` (must be logged in), then run a
collector via `javascript_tool`. Capture per GUID:
- `totalSharedCentimorgans` -> Unweighted cM
- `longestSharedSegment` -> Longest Segment
- `numSharedSegments` -> Segments

Endpoint:
`https://www.ancestry.com/discoveryui-matches/parents/list/api/matchSharedDna/{TESTER}/{MATCH}`

**Hardened collector pattern (respect the javascript_tool quirks below):**

```js
// Kick off without top-level await; store to window; poll separately.
window.__dnaStatus = "running";
window.__dnaResults = [];
(function(){
  const TESTER = "{TESTER_GUID}";
  const GUIDS = [ /* match GUIDs */ ];
  const base = "https://www.ancestry.com/discoveryui-matches/parents/list/api/matchSharedDna/";
  const out = [];
  async function one(m, attempt){
    try {
      const r = await fetch(base + TESTER + "/" + m, {credentials:"include"});
      const d = await r.json();
      return {guid:m, unweighted_cm:d.totalSharedCentimorgans,
              longest_segment:d.longestSharedSegment, segments:d.numSharedSegments};
    } catch(e){
      if(attempt < 2){ await new Promise(s=>setTimeout(s,500)); return one(m, attempt+1); }
      return {guid:m, unweighted_cm:"", longest_segment:"", segments:""};
    }
  }
  (async () => {
    for(let i=0;i<GUIDS.length;i+=50){
      const chunk = GUIDS.slice(i,i+50);
      const res = await Promise.all(chunk.map(m=>one(m,1)));
      out.push(...res);
      await new Promise(s=>setTimeout(s,150));
    }
    window.__dnaResults = out;
    window.__dnaStatus = "done";
  })();
})();
"started";
```

Then poll with a separate call: read `window.__dnaStatus`. When `"done"`, read results.

**Read results in chunks** to dodge the output cap (see quirks): pull 20 rows per call,
e.g. `JSON.stringify(window.__dnaResults.slice(0,20))`, then `slice(20,40)`, etc. A
single `browser_batch` of several `javascript_tool` calls can fetch all chunks at once.

**javascript_tool quirks (do not rediscover these):**
- Top-level `await` is unreliable and throws. Never use it. Fire async work without
  awaiting, store to a `window` variable, set a status flag, and poll.
- Output is truncated around ~1 KB. Slice large result sets into ~20-row chunks.
- The browser and the bash sandbox do not share a filesystem. Do not write a CSV in
  the browser and try to read it from bash -- return data as tool output.

**Error handling:**
- If >10% of requests fail: pause and alert the user (usually the tab is logged out
  or navigated off ancestry.com). Re-confirm the session and retry the failed GUIDs.
- Never stop silently on failure.

### FALLBACK -- `fetch_shared_dna.py` + browser_cookie3 (local Claude Code ONLY).

The repo script reads Chrome cookies directly and calls the same endpoint. It takes
the tester GUID as a parameter (`--kit-url` or `--tester-guid`) plus `--input-csv`
or `--guids-file`. **This does not work in Cowork** -- the sandbox cannot see the
user's Chrome. Use only when running as local Claude Code without a browser session.

---

## Step 3: Link Collection

### DEFAULT -- deterministic, verified URL construction (no per-profile scraping)

Build per-match links straight from the GUIDs. Both patterns are confirmed present
on live Ancestry; do not fabricate anything beyond them:

- Profile / compare:        `https://www.ancestry.com/dna/matches/{TESTER}/compare/{MATCH}`
- Tree + ThruLines compare: `https://www.ancestry.com/discoveryui-matches/compare/{TESTER}/with/{MATCH}`

These cover "view their tree" and "view common ancestor" for research triage without
loading any profile pages. The compare-with URL lands on Ancestry's comparison view
rather than a deep tree-person link -- sufficient for triage.

### OPTIONAL -- deep links (only if exact tree-person links are needed)

Run a per-profile browser pass over the **priority subset only** (matches passing
longest >= 20 AND AScM >= 12), not the whole batch. Note: the bulk `treeData` and
`commonAncestors` endpoints are POST-only, header-gated, and cached by the SPA;
calling them directly returns a 303 auth/CSRF redirect. Replaying them requires
hooking the live XHR before first paint and is advanced -- not required for a usable
workbook. Skip unless the user specifically asks for deep tree-person links.

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
| G   | Family Tree     | CSV + compare URL  | Hyperlink  |
| H   | Tree Size       | CSV               | Integer    |
| I   | Common Ancestor | CSV + compare URL  | Hyperlink  |
| J   | Line Assignment | Derived + user    | Color fill |
| K   | Groups          | CSV               | Text       |
| L   | Notes           | CSV               | Text       |
| M   | Match Side      | CSV               | Text       |
| N   | GUID            | Extracted         | Text       |

### Hyperlinks
- Match Name: links to ProfileURL (match profile page, not shared-matches tab)
- Family Tree: links to the compare URL if the match has a tree, otherwise text only
- Common Ancestor: links to the compare-with URL if Ancestry shows a common ancestor

### AScM Formula
Use Excel formula: `=IFERROR(D{row}/E{row},"")` -- never hardcode calculated values.

### Color Tiers (row-level, based on Longest Segment) -- generic for any tester

| Tier       | Condition                          | Fill    | Font             |
|------------|------------------------------------|---------|------------------|
| Fail       | AScM < 12 OR Longest Segment < 20  | FFD7D7  | CC0000 (red)     |
| Light green| Longest 20-30                      | E2EFDA  | 000000 (black)   |
| Med green  | Longest 30-50                      | A9D18E  | 000000 (black)   |
| Dark green | Longest 50+                        | 70AD47  | 000000 bold      |

### Line Assignment Colors (column J only) -- quadrant colors are fixed; surnames are per-tester

Colors are a fixed project convention (cool = paternal, warm = maternal). The
surname behind each quadrant, and the Groups-to-line mapping used to pre-fill this
column, come from the active tester config under `testers/`.

| Quadrant                 | Convention   | Fill    |
|--------------------------|--------------|---------|
| PP (paternal-paternal)   | green        | E2EFDA  |
| PM (paternal-maternal)   | blue         | BDD7EE  |
| MP (maternal-paternal)   | yellow/gold  | FFEB9C  |
| MM (maternal-maternal)   | coral        | FCE4D6  |
| Multiple                 | purple       | E2CEEF  |

Pre-fill Line Assignment from the Groups column using the active tester's
group-to-line mapping. Where no rule matches, leave blank for the user to fill in.
Never force-assign a match that maps to multiple lines -- mark it Multiple.

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
- Tree links built: X of Y
- Common ancestor links built: X of Y
- Any failures or outstanding items

---

## Batch Management

This pipeline is open-ended. Any number of matches can be processed.

For large lists (300+): collect in-browser in batches of 50, polling between.
For links: deterministic URL construction scales to any size with no extra requests.
For the workbook: each batch is a separate sheet or file; document in the tester config.

When adding a new batch for the same tester:
- Load CLAUDE.md and the active tester config from GitHub first
- Note the batch number (increment from the last recorded in the tester config)
- Do not overwrite prior batch files

When switching to a different tester:
- Repoint the Active Tester block in CLAUDE.md, then load that tester's config
- The tester GUID comes from that kit's URL; nothing in this skill is tester-specific

---

## Methodology Notes

Read `/skills/ashkenazi-analyst/SKILL.md` for full Ashkenazi DNA analysis context.
Key reminders:
- Always use UNWEIGHTED cM. Never TIMBER-adjusted figures.
- Platform relationship estimates are meaningless. Strip them.
- "Common Ancestor" in Ancestry = a clue. Verify before concluding.
- Multiple-line matches are held for investigation. Never force-assign.
