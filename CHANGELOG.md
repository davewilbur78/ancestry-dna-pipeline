# Changelog

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
The async IIFE call shape for javascript_tool is confirmed working.
Top-level await is rejected; IIFE resolves the Promise correctly.
50-match batches produce ~700 chars output -- well within the ~1KB cap.

### Next (Jeannette Klein)
- User to supply line definitions (surnames for PP/PM/MP/MM)
- Once groups are created in Ancestry, add group -> line mapping to tester config
- Optional browser pass for tree URLs
- Build Batch 2 when ready

## [2026-06-04] - Naming convention: FirstName_LastName required

### Decision
- Output filenames MUST include both first and last name.
- Surname-only names are forbidden -- a tester shares a surname with relatives.
- Updated in CLAUDE.md (Conventions) and the dna-match-extractor SKILL.md.

## [2026-06-04] - New Tester: Cynthia (Klein) Wilbur Batch 1 (300 matches)

See full entry in CHANGELOG on GitHub for details.

## [2026-06-04] - v2.1 In-Browser Collection (Cowork fix)

See full entry in CHANGELOG on GitHub for details.

## [2026-06-04] - v2.0 Kit-Agnostic Refactor

See full entry in CHANGELOG on GitHub for details.

## [2026-06-04] - Initial Pipeline Build: Adrienne Peckler Batch 1

See full entry in CHANGELOG on GitHub for details.
