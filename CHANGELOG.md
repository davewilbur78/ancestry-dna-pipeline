# Changelog

## [2026-06-14] - v2.6 Robustness fix: build_v2.2.py

### Context
User reported "the kit built was a mess" after running a new kit. Root cause: the
build script committed to GitHub (build_v2.1.py) had three critical bugs that break
on any new Cowork session or Claude Code run.

### Bugs Fixed (build_v2.2.py)

**Hardcoded Cowork session paths**
- build_v2.1.py had SOURCE_FILE and OUTPUT_DIR pointing to
  `/sessions/happy-upbeat-ride/mnt/...` -- a path that only exists in one specific
  Cowork session. Every new session gets a different session ID, so the script fails
  immediately with FileNotFoundError.
- Fix: CONFIG block now has clearly labeled placeholder strings with inline comments
  explaining what to update. The extraction date and output filename are still
  auto-derived; only TESTER_NAME, BATCH_NUM, SOURCE_FILE, and OUTPUT_DIR need to be
  set per run.

**Hardcoded source sheet name**
- build_v2.1.py opened the source XLSX with `src["DNA Matches"]` -- but pipeline-
  produced source files use various sheet names ("Batch 1", "Top 1000 Batch 1",
  "Top 1000 - Batch 1", etc.). Any file without a sheet named exactly "DNA Matches"
  raised a KeyError and crashed.
- Fix: build_v2.2.py auto-detects the correct sheet by looking for "Match Name" in
  cell A1. No sheet name is hardcoded.

**Index-based column reading**
- build_v2.1.py read every source column by zero-based index (row[1], row[3], etc.).
  Any source file with a different column order or count produced silently wrong data.
- Fix: build_v2.2.py builds a header-name-to-index map from row 1 and reads every
  column by name (_hget(row, "Longest Segment"), etc.). Robust to column order
  variations. GUID column handled gracefully if absent.

**strftime cross-platform fix**
- `%-d` (Linux zero-strip) crashes on macOS when running build_v2.2.py as a local
  Python script (Claude Code or terminal). Fixed to platform-neutral str(dt.day).

### Built / Produced
- build_v2.2.py (current production workbook builder)
- CLAUDE.md v2.6 (this entry + source schema + CONFIG block documented; "current
  production version" updated to build_v2.2.py)

### Built / Produced (additional)
- Cynthia_Wilbur_1500matches_20260614_v2.2.xlsx -- rebuilt with fixed build_v2.2.py
  Tiers: Priority 114 (H13/S47/I54), Watch 191, Low 1195

### Note
All three files committed to GitHub on 2026-06-14 after GitHub MCP token was renewed.

### Next
- Re-run any other kits that produced bad output with the buggy build_v2.1.py

---

## [2026-06-14] - v2.5 Workbook Design Round 3 (build_v2.1.py)

### Context
Design session workshopping the Start Here / Priority Matches / Watch List / All Matches
workbook produced by the build script. Input: Lesley Sterling Batch 1 (1000 matches,
test case only -- her data does not belong in the repo). Build script moved from
internal version numbering (v5, v6) to the project's decimal versioning scheme.

### Decisions and Changes

**Workbook header (reverted and refined)**
- Row 1: TESTER_NAME (large, 20pt, navy) -- this was correct in v4 and accidentally
  changed to the tab name in the v5 build. Reverted.
- Row 2 subtitle: "{Tab Name}  ·  Ancestry data retrieved {DATE}". Date language
  now explicitly says "Ancestry data retrieved" to clarify what the date means.
  Match counts removed from subtitle.

**Output filename auto-naming (new)**
- Format: {FirstName}_{LastName}_{N}matches_{YYYYMMDD}_v{SCRIPT_VERSION}.xlsx
- N = total rows in input (all matches, not just priority).
- Date = file modification date of the enriched input XLSX (auto-derived; no manual entry).
- SCRIPT_VERSION = the constant at the top of the build script.
- Example: Lesley_Sterling_1000matches_20260612_v2.1.xlsx

**Build script versioning**
- Internal "v5, v6" numbering replaced with the project's decimal scheme.
- Script is now build_v2.1.py. Minor design improvements = v2.2, v2.3, etc.
- Output filename carries the script version so every workbook is traceable to
  the code that built it.

**Column rename: "Family Line Assignment"**
- "Line Assignment" renamed to "Family Line Assignment" throughout the entire workbook:
  COLS definition, column guide, section header, body text, Quick Start step references.
- Column width: 16 -> 22 to accommodate the longer label.

**Batch Summary redesign**
- Removed navy background from "BATCH SUMMARY" label (was visually fused with the header).
- Light grey fill (#F4F6FB) with navy text; thin navy box border around the entire block.
- Small gap (6px) between the navy banner and the stats block.

**Grey columns**
- Previous: D8D8D8 / BBBBBB / 999999 per tier -- invisible on medium green and most
  backgrounds.
- New: #707070 italic uniformly across all greyed columns and all tier backgrounds.
  Columns: Weighted cM, Match Side, GUID (Reference Anchor).

**Match Name column**
- Width: 28 -> 36.
- Alignment: right-aligned (user preference, matches name-column conventions in research tools).

**Font sizes**
- Column headers: 10 -> 11pt.
- Data rows: 11 -> 12pt.
- Body text (Start Here): 11 -> 12pt.
- Section headers (sh_navy): 12 -> 13pt.
- Row heights adjusted proportionally.

**Gil Bardige call-to-action (redesigned)**
- Old: small blue "WANT TO GO DEEPER?" header with a modest play button.
- New: large burgundy (#7B2D3E) banner row "▶ WATCH THIS PRESENTATION", 24pt bold white,
  height 60px, full hyperlink. Presentation title in a lighter burgundy row below.
  Description text in the same light burgundy fill.

**Quick Start: inline rich text for column references**
- Column name references in Quick Start steps (COMMON ANCESTOR, NOTES, FAMILY LINE
  ASSIGNMENT) now use openpyxl CellRichText / TextBlock / InlineFont to render as
  ALL CAPS BOLD UNDERLINED inline within the step text.

**Line Assignment call-to-action (Start Here)**
- Removed purple fill from the Family Line Assignment row in the column guide.
- Replaced with: light amber fill (#FFF5CC) + dashed amber border -- "action required"
  visual system. Dashed amber = something the workbook user must fill in themselves.
- Same dashed amber box applied to the action callout at the top of the Quadrant Colors
  section.

**"Two Numbers That Matter" section**
- Old: threshold descriptions in plain white body text.
- New: each threshold is shown as a color-coded row using the actual tier fills
  (dark green / medium green / light green / pink), so the user sees the color coding
  in context rather than just reading about it.

**Color coding language**
- "Fails filter" label removed from workbook and all documentation.
- Replaced with "Watch List · Low Priority" (consistent with the tab structure).

**Flag column (⚠️) description**
- Column guide now shows "B   ⚠️  Flag" as the label.
- Description leads with: "⚠️ appears when a match shares only 1 or 2 DNA segments."
  The symbol is shown in the guide so users know exactly what they are looking for.

**CONFIG block**
- Top of build script has a clearly labeled CONFIG block with all kit-specific variables.
- RETRIEVAL_DATE removed from CONFIG (now auto-derived from input file modification date).
- OUTPUT_FILE removed from CONFIG (now auto-built from tester name + count + date + version).

### Built / Produced
- build_v2.1.py (current production workbook builder)
- CLAUDE.md v2.5 (this entry + conventions updated)
- Lesley_Sterling_1000matches_20260612_v2.1.xlsx (test output -- not in repo)

### Note
GitHub MCP credentials expired mid-session. Files are ready to commit; reconnect
GitHub in Settings > Connections and run the commit manually or in the next session.

### Next
- Review the new workbook visually and note any remaining design issues.
- Next batch for a new tester: update CONFIG block (TESTER_NAME, BATCH_NUM, SOURCE_FILE,
  OUTPUT_DIR) -- everything else is automatic.

## [2026-06-14] - Cynthia (Klein) Wilbur Top1500 Batch 1 (1500 matches)

### Built
- `Cynthia_Wilbur_1500matches_20260614_v2.1.xlsx` -- full Top1500 workbook (v2.1 builder)
- `Cynthia_Wilbur_top1500_intermediate.xlsx` -- intermediate source file
- `make_intermediate_cynthia_top1500.py` -- script that builds intermediate from CSV + API JSON

### Tier breakdown
- DARK GREEN  (longest >=50 cM):  13
- MED GREEN   (longest 30-49):   47
- LIGHT GREEN (longest 20-29):   54
- RED (fails filter):           1386
- Total: 1500

### Notes
- New run extending beyond the prior Top1000 batch (built same session)
- API data: in-browser fetch, 1500/1500, 0 failures
- Match Name links: 1500/1500 (deterministic profile compare URLs)
- Family Tree links: constructed for all matches with Tree Size > 0
- Common Ancestor ThruLines links: constructed where GA export has CA text
- Line Assignment: blank (no quadrant mapping supplied)
- vs Top1000: added 9 med green, 17 light green, 474 red in positions 1001-1500; no new dark green
- Endogamy status not yet confirmed; high red% consistent with Ashkenazi pattern

---

## [2026-06-13] - Lee Klein Batch 2 (1251 matches, 40-90 cM sweet spot)

### Built
- `Lee_Klein_dna_api_results_Batch2.csv` -- 1251/1251 API calls succeeded (0 failures)
- `Lee_Klein_DNA_Matches_Batch2.xlsx` -- sweet spot filtered list (40-90 cM range)

### Tier breakdown
- RED (fails filter):          1145
- LIGHT GREEN (longest 20-30):   72
- MED GREEN   (longest 30-50):   33
- DARK GREEN  (longest 50+):      1
- Passing filter total:          106 / 1251

### Notes
- No testers/ notes file for Lee Klein; Line Assignment left blank
- High red% expected: the 40-90 cM GA filter captures total cM but many matches
  have short longest segments and/or high segment counts (AScM < 12), consistent
  with Ashkenazi endogamy pileup in this cM range
- 1 dark green match is high priority

---

## [2026-06-12] - Emily Jennewein Batch 1 (1000 matches)

### Built
- `Emily_Jennewein_dna_api_results_Batch1.csv` -- 1000/1000 API calls succeeded (0 failures)
- `Emily_Jennewein_DNA_Matches_Batch1.xlsx` -- full workbook delivered to Edra DNA project folder

### Tier breakdown
- RED (fails filter):          826
- LIGHT GREEN (longest 20-30): 101
- MED GREEN   (longest 30-50):  63
- DARK GREEN  (longest 50+):    10
- Passing filter total:         174 / 1000

### Notes
- No testers/ notes file for Emily Jennewein; Line Assignment left blank
- 10 dark green matches (longest 50+ cM) are high priority for investigation

---

## [2026-06-12] - Lesley Sterling Batch 1 (1000 matches)

### Built
- `Lesley_Sterling_dna_api_results_Batch1.csv` -- 1000/1000 API calls succeeded (0 failures)
- `Lesley_Sterling_DNA_Matches_Batch1.xlsx` -- full workbook delivered to Edra DNA project folder
- `fetch_shared_dna.py` updated: added concurrent batching (50/batch via ThreadPoolExecutor);
  150ms between batches; preserves original GUID order in output

### Tier breakdown
- RED (fails filter):          874
- LIGHT GREEN (longest 20-30):  64
- MED GREEN   (longest 30-50):  49
- DARK GREEN  (longest 50+):    13
- Passing filter total:         126 / 1000

### Notes
- No testers/ notes file for Lesley Sterling; Line Assignment left blank
- High red% (87%) is consistent with Ashkenazi endogamy; confirm tester's background
  before treating the 874 red rows as untraceable

---

## [2026-06-05] - v2.4 Standard config-free flow

### Decision
- The pipeline is now standard and uniform for every kit: present a Genealogy Assistant
  CSV + kit URL, get the formatted workbook. No per-tester config required, no "active
  tester" to set. Removed the config-loading step from the operating model and the skill.
- The only kit-specific element, the Line Assignment column, is left BLANK by default and
  pre-filled only when the user supplies a surname-to-branch mapping (inline or from notes).
- testers/ files are now explicitly OPTIONAL research notes (line mapping, groups, batch
  history). They never gate workbook creation. Kept: adrienne-peckler, jeannette-klein,
  cynthia-wilbur.

### Built / Produced
- CLAUDE.md v2.4: added "Standard Flow"; "Active Tester" replaced by "Tester Configs
  (optional)"; Step 4 Line Assignment now blank-by-default
- Plugin v0.3.2: dna-match-extractor SKILL.md v2.3 (config-free inputs; optional mapping)

### Next
- Reinstall the v0.3.2 plugin

## [2026-06-05] - v2.3 Reconcile parallel threads + no default tester

### Context
Two threads diverged after v2.1 and both reached "v2.2": one onboarded Jeannette
Klein (300 matches) and pushed to GitHub; another onboarded Cynthia (Klein) Wilbur
(300 matches) locally and adopted the FirstName_LastName naming rule. This entry
merges both lines and makes the project fully kit-agnostic. Nothing was dropped.

### Decisions
- No tester is loaded by default. CLAUDE.md "Active Tester" is now "(none set)". Each
  session loads a kit by name from testers/, creates one from _TEMPLATE.md, or asks
  which to load. Adrienne is no longer the default; she remains a selectable config.
- Kept all three tester configs: adrienne-peckler.md, jeannette-klein.md (was repo-only),
  cynthia-wilbur.md (was local-only). All now in the repo.
- Adopted the FirstName_LastName output-naming rule project-wide (CLAUDE.md + skill).
- Folded the Jeannette thread's field-confirmed Step 2 refinement into CLAUDE.md: the
  async-IIFE call shape (top-level await is rejected), Promise.all per 50-match batch,
  compact read-back to stay under the ~1KB output cap. Confirmed at 300-match scale.
- Genericized the last tester-specific bits: fetch_shared_dna.py uses a placeholder
  GUID in its examples; README has no named tester and the accurate in-browser method.

### Built / Produced
- CLAUDE.md v2.3 (no default tester, IIFE Step 2, all three configs listed)
- Plugin v0.3.1: dna-match-extractor SKILL.md v2.2 (FirstName_LastName naming; "no
  default" wording)
- README.md agnostic + accurate; testers/cynthia-wilbur.md added to repo

### Next
- Reinstall the v0.3.1 plugin (installed copy is still pre-fix)
- Per-tester next steps live in each tester's config under testers/

## [2026-06-04] - v2.2 Jeannette Klein Batch 1

### Context
Second full pipeline run. New tester: Jeannette Klein, kit E68B82EF-C277-43AA-91DF-21532961EFC6.
300 matches from Genealogy Assistant CSV export. Ran entirely in Cowork.

### What Was Done
- Parsed 300-match CSV, extracted all GUIDs (0 missing)
- API collection via in-browser fetch (Claude in Chrome, async IIFE pattern):
  6 batches of 50, Promise.all within each batch, 300/300 succeeded, 0 failures
- Built Klein_DNA_Matches_Batch1.xlsx with full column schema, color tiers,
  AScM formula, and Match Name hyperlinks
- Created testers/jeannette-klein.md config
- Switched CLAUDE.md Active Tester to Jeannette Klein (v2.2)

### Tier Breakdown (Batch 1)
  17 dark green (longest 50+ cM)
  57 med green (longest 30-50 cM)
  63 light green (longest 20-30 cM)
  163 red (fail filter: longest < 20 OR AScM < 12)

### Pattern Confirmed
The async IIFE call shape for javascript_tool is confirmed working:
  (async () => { const r = await fetch(url,{credentials:'include'}); const d=await r.json(); return ...; })()
Top-level await is rejected; IIFE resolves the Promise correctly.
50-match batches produce ~700 chars output -- well within the ~1KB cap.

### Decisions
- No new methodology decisions. Pipeline ran clean per v2.1 design.
- Line definitions for Jeannette Klein TBD (user to supply surnames/quadrant mapping)
- Tree/CA hyperlinks outstanding (optional browser pass when needed)

### Next (Jeannette Klein)
- User to supply line definitions (surnames for PP/PM/MP/MM)
- Once groups are created in Ancestry, add group -> line mapping to tester config
- Optional browser pass for tree URLs
- Build Batch 2 when ready

## [2026-06-04] - Naming convention: FirstName_LastName required

### Decision
- Output filenames MUST include both first and last name:
  {FirstName}_{LastName}_DNA_Matches_Batch{N}.xlsx. Surname-only names are forbidden --
  a tester shares a surname with relatives, so "Wilbur_..." or "Beyer_..." is ambiguous.
  Updated in CLAUDE.md (Conventions) and the dna-match-extractor SKILL.md (Output Naming).

### Renamed (existing files brought into compliance)
- Wilbur_DNA_Matches_Batch1.xlsx        -> Cynthia_Wilbur_DNA_Matches_Batch1.xlsx
- Klein_DNA_Matches_Batch1.xlsx         -> Cynthia_Wilbur_DNA_Matches_Batch1_SUPERSEDED.xlsx
    (same kit as the Wilbur file -- older duplicate built from the Klein maiden-name
     export; kept for reference, not the working file)
- Beyer_DNA_Matches_Batch1.xlsx         -> Susan_Beyer_DNA_Matches_Batch1.xlsx
- Beyer_dna_api_results_Batch1.csv      -> Susan_Beyer_dna_api_results_Batch1.csv
- Adrienne_Peckler_DNA_Matches_Batch1.xlsx already compliant (unchanged)

## [2026-06-04] - New Tester: Cynthia (Klein) Wilbur Batch 1 (300 matches)

### Context
First kit onboarded after the v2.1 in-browser fix. Full run end to end in Cowork
using the in-browser collection path (no browser_cookie3). 300 matches, 0 failures.

### Decisions
- Onboarded Cynthia (Klein) Wilbur as a new tester: testers/cynthia-wilbur.md.
- Confirmed the v2.1 in-browser fetch path works at 300-match scale: fire-and-forget
  collection into a window var + poll (sidesteps any tool timeout); results read back
  in ~50-row numeric chunks (sidesteps the ~1KB output cap).
- Step 3 links: treeData (bulk POST) carries only public/private/size flags -- no
  treeId or URL; commonAncestors (bulk POST) returned empty. So bulk endpoints cannot
  supply hyperlinks. Resolution used:
    * Tree links = deterministic compare /trees tab URL (profileURL + "/trees"),
      verified against a real rendered anchor. Applied to all 237 tree matches.
    * ThruLines / Common Ancestor links = captured per-match from compare pages for the
      15 common-ancestor matches (ThruLines URL is per-ancestor, not constructible).
- Endogamy caveat recorded: AScM >= 12 filter is Ashkenazi-calibrated; Cynthia's
  endogamy status is unconfirmed, so the 253 red rows may be over-flagged.

### Built / Produced
- Ancestry DNA Data extractor/Cynthia_Wilbur_DNA_Matches_Batch1.xlsx
  (300 matches; 300 name links, 237 tree links, 15 ThruLines links; AScM live formula;
   tiers 13 dark / 17 med / 17 light / 253 red; recalc verified 0 formula errors)
- testers/cynthia-wilbur.md (new tester config + batch state)

### Next
- Confirm endogamy status; capture fan chart to fill lines/surnames
- Create Ancestry groups, then auto-fill Line Assignment
- Optional deep-link pass beyond the priority 50; build Batch 2

## [2026-06-04] - v2.1 In-Browser Collection (Cowork fix)

### Context
First real Cowork run (Susan Beyer kit, 150 matches, 0 failures) succeeded but the
documented Step 2 did not run as written. browser_cookie3 cannot see the user's
Chrome from the sandboxed Linux VM. Full debrief: docs/PIPELINE_DEBRIEF_Cowork_run.md.

### Decisions
- In-browser fetch (Claude in Chrome, GET matchSharedDna with credentials:"include")
  is now the PRIMARY collection path. Works in both Cowork and local Claude Code.
- fetch_shared_dna.py + browser_cookie3 demoted to a local-Claude-Code-only fallback.
- Step 3 link collection: deterministic verified compare-URL construction is now the
  default (no per-profile scraping). Per-profile deep-link pass is optional, priority
  subset only. treeData/commonAncestors are POST-only/header-gated/SPA-cached -- noted.
- javascript_tool guardrails baked into the skill: no top-level await (poll a window
  flag instead), ~1KB output cap (read results in ~20-row chunks), browser and sandbox
  are separate filesystems (return data as tool output, not a shared CSV).
- Added an environment check at the top of the skill (Cowork vs local).

### Built / Produced
- CLAUDE.md v2.1: Step 2 and Step 3 rewritten per the above
- Plugin v0.3.0: dna-match-extractor SKILL.md v2.1 with the hardened in-browser collector
- docs/PIPELINE_DEBRIEF_Cowork_run.md (the Cowork run debrief, archived)

### Next
- Reinstall the rebuilt v0.3.0 plugin

## [2026-06-04] - v2.0 Kit-Agnostic Refactor

### Decisions
- Split the project into generic pipeline/methodology (CLAUDE.md) vs per-tester config (testers/)
- CLAUDE.md bumped to v2.0; added an Active Tester pointer and an "Adding a New Tester" flow
- Quadrant colors (PP green, PM blue, MP yellow, MM coral, Multiple purple) are a fixed project
  convention; surname-to-quadrant mapping now lives in the tester config, not in CLAUDE.md or skills
- Tester GUID is always a parameter -- never hardcoded anywhere

### Built / Produced
- testers/adrienne-peckler.md (archived Adrienne's full config + batch state + research priority)
- testers/_TEMPLATE.md (blank per-tester template for new kits)
- fetch_shared_dna.py parameterized: takes --kit-url/--tester-guid and --input-csv/--guids-file;
  extracts match GUIDs from the CSV URL column; retry-once and >10% failure warning built in
- Plugin v0.2.0: both SKILL.md files de-hardcoded (no more Adrienne surnames); they now read
  line names, colors, group mapping, and priorities from the active tester config
- README updated for multi-kit design

### Next
- Browser pass for Adrienne Batch 1 (tree + common ancestor URLs)
- Reinstall the rebuilt v0.2.0 plugin

## [2026-06-04] - Initial Pipeline Build: Adrienne Peckler Batch 1

### Decisions
- Use Unweighted cM (totalSharedCentimorgans) not TIMBER-adjusted for all Ashkenazi analysis
- AScM = Unweighted cM / Segments -- minimum threshold 12, longest segment minimum 20 cM
- Color tiers based on Longest Segment (not AScM -- distribution too compressed): 
  < 20 = red fail, 20-30 = light green, 30-50 = med green, 50+ = dark green (#70AD47 bold black)
- Line Assignment column in spreadsheet = working hypothesis; separate from Ancestry Groups tags
- Ancestry Groups = confirmed tagging in platform; Line Assignment = research working state
- 8-couple color tier system planned for Ancestry groups: Tier 1 (line placed), Tier 2 (spouse confirmed)
- Cool colors (blue/green) for paternal, warm (yellow/coral) for maternal
- Ancestry now supports 64 groups (not 24 as older guides state)
- GUID extracted from URL column; GUID column pushed to far right (col N)
- Family Tree and Common Ancestor columns need browser pass for actual hyperlinks (outstanding)
- Discovered Ancestry matchSharedDna API endpoint for bulk data collection

### Built / Produced
- Adrienne_Peckler_DNA_Matches_Batch1.xlsx (150 matches, fully enriched except hyperlinks)
- CLAUDE.md (this project brain)
- dna-match-extractor Cowork plugin (v0.1.0)
- ashkenazi-analyst skill

### Configuration Captured
- Tester: Adrienne Balsky Peckler, Kit 4FB3190E-37B7-401F-A8CF-431950413661
- Fan chart captured: all 4 grandparent lines and 8 great-grandparent couples identified
- Existing Ancestry groups: Balsky (PP), SINGER/SPRINGER (PM)
- Mendick (MP) and Weinberger/Danko (MM) groups not yet created in Ancestry

### Next
- Browser pass: collect tree URLs and CA URLs for Batch 1
- Build Batch 2 (next 150 or more matches) using the Cowork plugin
- Confirm anchor kit (Marvin Balsky or Harriette Mendick)
- Create Mendick and Weinberger/Danko groups in Ancestry
