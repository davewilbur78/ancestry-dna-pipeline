---
Klein Super-Siblings Superkit — CLAUDE.md
Version: 1.1
Last updated: 2026-06-07 UTC
---

## Operating Model

This file is the single source of truth for the Klein Super-Siblings Superkit project.
It lives at: https://raw.githubusercontent.com/davewilbur78/ancestry-dna-pipeline/main/klein-superkit/CLAUDE.md

Fetch and read this file fully at the start of every session.
Never rely on memory from previous conversations.
Confirm the version and date out loud when loaded.

After loading this file, immediately read these two skills in order before
doing any analysis or interpretation:
  1. /mnt/skills/plugins/dna-match-extractor:ashkenazi-analyst/SKILL.md
  2. /mnt/skills/user/ashkenazi-genetic-genealogist/SKILL.md
These are mandatory, not optional. Do not begin any match interpretation,
prioritization, or line assignment without having read both. Confirm they
are loaded alongside the CLAUDE.md version confirmation.

This project lives in the `klein-superkit/` subfolder of the ancestry-dna-pipeline repo.
The root-level CLAUDE.md governs individual workbook construction (the pipeline tool).
This CLAUDE.md governs superkit comparison analysis. Both files work together but
cover different jobs. When building a workbook for a new comparison kit, load the
root CLAUDE.md. When doing superkit comparison analysis, load this file.

---

## What This Project Is

The Klein Super-Siblings Superkit project combines the AncestryDNA match exports of
four full siblings (children of Vernon Klein and Mildred Singer) into a single
analytical reference workbook -- the superkit. The superkit is used as a standing
reference against which individual match workbooks are compared, sorted by
grandparental line (PP/PM/MP/MM), and analyzed for genealogically significant connections.

This is not a pipeline project. It is an analytical project that uses the superkit as
a super-sibling: a proxy for nearly the full parental genome (~93-94% combined coverage
from four siblings) against which any external kit can be compared.

For background on the sibling superkit concept and the full build methodology, see:
`methodology/Klein_Sibling_Superkit_Methodology.md`

---

## NAMING CONVENTIONS (CRITICAL -- read before every session)

**DO NOT say "the Klein family."** This is ambiguous and will corrupt analysis.
Henry Klein (paternal grandfather) had four brothers. The Klein surname appears
across PP and PM lines AND in unrelated matches. Bruce Klein, Kenneth Klein, Saul Klein,
and Jeannette Klein all appear in the superkit -- they cannot all be assumed to connect
through the same Klein line.

**Correct terms for the four siblings:**
- **Klein super-siblings**
- Their father: Vernon Klein (1923-1976)
- Their mother: Mildred Singer (1927-2015)
- Sibling labels: G (Gerri), S (Susan), C (Cynthia), L (Lee)

**Correct terms for their workbook:**
- **Klein super-siblings superkit** or just **the superkit**
- Current file: Klein_Siblings_Superkit_1000_B1.xlsx

---

## The Klein Super-Siblings -- Four Siblings

All four are full siblings sharing the same four grandparents.
The fan chart in `family/fan-chart.md` applies equally to all four.

| Sibling | Ancestry Name | Maiden Name | Label | Tester GUID |
|---------|--------------|-------------|-------|-------------|
| Gerri   | Gerri Taylor | Klein | G | 6E56EBFB-76A8-4342-AABB-1F9AF8A1746C |
| Susan   | Susan Beyer  | Klein | S | 676A3661-6BE5-4901-B79E-A34CC178C785 |
| Cynthia | Cynthia Wilbur (Managed by Rachel Sanda) | Klein | C | 41814A70-DA1A-46D9-B7BF-9B5BBF12D87D |
| Lee     | Lee Klein    | Klein | L | 01CD5A24-FAA0-4A03-990B-49211928FA5E |

---

## Family Structure -- Four Grandparental Quadrants

| Quadrant | Color | Grandparent | Key Surnames |
|----------|-------|-------------|--------------|
| PP (Paternal-Paternal) | Green  | Henry Klein (1880-1941) | Klein, Ripner, Kornfeld |
| PM (Paternal-Maternal) | Blue   | Jeanette Schwartz (1900-1987) | Schwartz, Greenberg |
| MP (Maternal-Paternal) | Yellow | Jacob Singer (1895-1981) | Singer, Jacobs, Springer |
| MM (Maternal-Maternal) | Coral  | Lena Teitelbaum (1898-1957) | Teitelbaum, Halbfinger, Blasman, Pioro |

### PP -- Henry Klein line
Henry Klein's parents: Samuel Klein (1873-1949) + Katalin Ripner (1874-1923)
Samuel's parents: Ignatz Klein (1851-1906) + Hannah Kornfeld (1856-1925)
Katalin's parents: Jozsef Ripner (1832-1901) + [unknown]
Extended: Yakab Ripner (1796-1864) + Rosie Ripner (1796-); Moses Morris Moshe Kornfeld (1797-1895)
Collateral Klein line: Gustav Klein, Sarah Klein (1840-1902), Esther Klein, Miriam Klein, Aser Kline
NOTE: Multiple people named "Klein" in the superkit may connect through PP, but Klein is
also a widespread Ashkenazi surname. Never assume Klein = PP without evidence.

### PM -- Jeanette Schwartz line
Jeanette Schwartz's parents: Herman Schwartz (1873-1939) + Tillie Greenberg (1880-1965)
Herman's parents: Isaac Schwartz + Vita Schwartz
Tillie's parents: Kecell [Greenberg] + Sussa Sadie (~1901)

### MP -- Jacob Singer line
Jacob Singer's parents: Samuel Singer (1870-1933) + Minnie Jacobs (1871-1956)
Samuel's parents: Tzvi Dov Springer (-1908) + Unknown Mrs. Springer
Minnie's parents: *Avraham [unknown] + unknown (BRICK WALL -- only a first name known)
OPEN QUESTION: The Springer-to-Singer transition. Samuel Singer's father was Tzvi Dov
Springer. The surname change at immigration is unresolved and is an active research target.

### MM -- Lena Teitelbaum line
Lena's parents: Max Teitelbaum (1861-1905) + Gussie Golda Fannie Halbfinger (1875-1953)
Max's parents: Yitzchak Isaac Teitelbaum + Ledvar Leah
Gussie's parents: Wolf Halbfinger (1836-1869) + Sura Sarah Chaia Blasman (1835-1904)
Extended: Lewek Halbfinger; Maryem Pioro; Icek Iciek Icyk Isaac Blasman; Lala

---

## The Superkit -- Current State (Batch 1)

**File:** Klein_Siblings_Superkit_1000_B1.xlsx
**Build date:** 2026-06-07
**Source retrieval date:** 2026-06-05
**Batch:** 1 (top 1000 per sibling)

### Sheet Structure
| Sheet | Contents | Row Count |
|-------|----------|-----------|
| Superkit | Filter-passing matches. Primary research surface. | 337 matches |
| Cross-Kit Master | All unique matches, no filter. Per-sibling columns. | 3,264 matches |
| Gerri | Gerri's individual match data (pipeline schema) | 993 rows |
| Susan | Susan's individual match data | 993 rows |
| Cynthia | Cynthia's individual match data | 993 rows |
| Lee | Lee's individual match data | 993 rows |
| Stats | Summary counts, tier distribution, noise floor | -- |
| Legend | Color key, filter rules, methodology notes | -- |

Tab order note: Sibling sheets appear before Stats and Legend (build-order artifact).
Cosmetic only. Can be corrected with wb.move_sheet() on next rebuild.

### Current Counts and Distribution
- Total source rows (4 x 1000): 4,000
- Excluded (testers + immediate family): 28 rows
- Loaded into Cross-Kit Master: 3,264 unique GUIDs
- Superkit filter-passing (Longest >= 20 AND AScM >= 12): 337

**Tier distribution:**
- Dark Green (Max Longest >= 50 cM): 16 matches
- Medium Green (30-49 cM): 121 matches
- Light Green (20-29 cM): 200 matches

**Sibling count distribution:**
- 4 siblings (GSCL): 8 matches
- 3 siblings: 24 matches
- 2 siblings: 58 matches
- 1 sibling only: 247 matches (73% -- high single-sibling rate is expected and correct)

**Other:**
- Line Assignments filled: 0 (fresh build, no comparison work yet done)
- ThruLines populated: 11 matches

### Top Priority Matches (by Max Longest Segment)
| Match | Max Longest | Sib Count | Pattern | ThruLines | Notes |
|-------|-------------|-----------|---------|-----------|-------|
| Jeannette Klein | 186 cM | 4 | GSCL | No | High priority -- identity under investigation |
| H. N. (Managed by staterdave) | 129 cM | 4 | GSCL | No | High AScM (45.59) |
| Bruce Klein | 110 cM | 4 | GSCL | Yes | ThruLines active; Both sides tag |
| Kenneth Klein | 85 cM | 4 | GSCL | Yes | ThruLines active |
| Veronica Clark Feliz | 71 cM | 4 | GSCL | Yes | ThruLines active |
| Marcia Wall | 69 cM | 1 | S only | No | Susan only; high-value solo |
| Sandra Easley (peanut9142) | 63 cM | 3 | SCL | No | |
| Rhonda Gannon | 63 cM | 2 | SL | No | |
| Jenna Pisnoy | 60 cM | 4 | GSCL | No | |
| Burton Weiss | 60 cM | 3 | SCL | No | |

### ThruLines Matches (all 11)
Bruce Klein (110), Kenneth Klein (85), Veronica Clark Feliz (71), Arnold Schneider (57),
Adrienne Peckler (52), Claire Hertz (44), Saul Klein (38), Peri Sigman (35),
Patricia Rose (31), Troy60681 (30), Susan Novak Backer (29)

### Excluded GUIDs (from all analytical sheets)
| Person | Relationship | GUID |
|--------|-------------|------|
| Gerri Taylor | Sibling tester | 6E56EBFB-76A8-4342-AABB-1F9AF8A1746C |
| Susan Beyer | Sibling tester | 676A3661-6BE5-4901-B79E-A34CC178C785 |
| Cynthia Wilbur | Sibling tester | 41814A70-DA1A-46D9-B7BF-9B5BBF12D87D |
| Lee Klein | Sibling tester | 01CD5A24-FAA0-4A03-990B-49211928FA5E |
| David Wilbur | Cynthia's spouse | E4936FEE-6266-4579-A6A5-501A2D41B365 |
| Rebekah Wilbur | Cynthia's child | AE5CE724-01ED-443B-8E47-CFDBDEE6A8A8 |
| Rachel Sanda | Cynthia's child | 7E652513-2FB4-4FF4-8E23-066D8C5D4514 |
| Harry Nelson | Susan's grandson (name pattern) | [GUID not yet captured] |
| Jacob Sanda | Cynthia's grandson (name pattern) | [GUID not yet captured] |
| Troy60681 | Adrienne Peckler's son -- exclude from all Adrienne comparison work. His AP-side figure (268 cM Longest, 3448 unwtd cM) is immediate-family scale and is not a comparable shared match. | 21DA164D-DC18-4507-A753-61944E401D2F |

---

## The Individual Workbooks -- Schema and Role

Each sibling's individual workbook is the source of record for that sibling's match data.
It is produced by the pipeline (root CLAUDE.md governs construction). The sibling
reference sheets in the superkit (Gerri, Susan, Cynthia, Lee) are convenience copies
regenerated from these source files at build time. When individual workbooks are updated,
the superkit is rebuilt from them.

Comparison workbooks for external kits (known relatives, research group members) use
the IDENTICAL pipeline schema. An external kit is just another person's match export,
processed by the same pipeline, producing the same column layout.

### Individual Workbook Column Schema
| Column | Notes |
|--------|-------|
| Match Name | Hyperlinked to compare page (tester-dependent URL) |
| Longest Segment | Longest shared segment in cM -- PRIMARY METRIC |
| AScM | Ashkenazi Segment Metric = Unweighted cM / Segments (live formula in individual workbooks; recomputed fresh in superkit build) |
| Unweighted cM | Total shared DNA, unweighted. Always use this. |
| Segments | Number of shared segments |
| Weighted cM | TIMBER-adjusted figure. Recorded but NEVER used in analysis. |
| Family Tree | Hyperlinked to match's tree (match-level URL, same for all siblings) |
| Tree Size | Number of people in match's tree |
| Common Ancestor | Hyperlinked to ThruLines page if populated |
| Line Assignment | Blank by default; filled manually as evidence accumulates |
| Groups | Ancestry group tags assigned to this match |
| Notes | Researcher notes |
| Match Side | Ancestry's paternal/maternal/both (tree-derived; only reliable if match has a linked tree) |
| GUID | Match's Ancestry kit identifier -- primary join key across all workbooks |

### How Hyperlinks Work
All hyperlinks are constructed from GUIDs, not scraped from the source files.
URL patterns:
- Match Name: https://www.ancestry.com/dna/matches/{TESTER_GUID}/compare/{MATCH_GUID}
- Family Tree: https://www.ancestry.com/discoveryui-matches/compare/{TESTER_GUID}/with/{MATCH_GUID}
- Common Ancestor: same URL pattern as Family Tree

The openpyxl hyperlink trap: loading with read_only=True and values_only=True silently
discards embedded hyperlink objects. Always construct URLs from GUIDs rather than
reading them from source files. See methodology doc Appendix C.1 for full detail.

---

## The Superkit Architecture -- How the Three Layers Relate

**Layer 1: Individual workbooks** (source of record)
One per tester. Pipeline output. Authoritative data for that sibling's or comparison
kit's matches. Never modified by comparison work.

**Layer 2: The Klein super-siblings superkit** (the reference)
Derived from the four sibling workbooks. One row per unique GUID across all siblings.
Treats the four siblings as a single super-sibling. The Superkit sheet (filter-passing) is
the primary research surface. The Cross-Kit Master is the full analytical layer.
Comparison work does not modify the superkit except for manually entered Line Assignments.

**Layer 3: Comparison workbooks** (analytical products)
One per external kit pairing. Cross-references the external kit against the superkit
by GUID. Produces the intersection (shared matches) for line-sorting evidence.
Results feed back into the superkit only through the researcher's manual Line Assignment entries.

### The Super-Sibling Concept
The four siblings together cover ~93-94% of both parents' combined genome. Any real
recent match is likely to appear in at least one sibling's top-N. The Sibling Pattern
column (GSCL encoding, dots for absent) provides a phasing proxy: matches appearing
consistently in the same subset of siblings likely connect through the same ancestral
line, because those siblings inherited the relevant genomic region.

This makes the superkit more powerful than any single kit for:
1. Distinguishing endogamy background from real recent matches (low cM Range + all 4 siblings = background signal)
2. Finding matches too specific to land in one sibling's top-N but present in at least one
3. Observing the distribution pattern of a comparison kit's shared matches across the Klein super-siblings

---

## Comparison Architecture -- How a Comparison Run Works

### Step 1: Build the external kit's workbook
Run the pipeline (root CLAUDE.md) on the external kit's Genealogy Assistant CSV export.
Apply all standard filters (Longest >= 20, AScM >= 12). Record the retrieval date.

### Step 2: Cross-reference against the superkit
Join external kit and superkit by GUID. Identify:
- Intersection (shared matches above filter threshold in both)
- Superkit-only matches (not found in external kit above threshold)

### Step 3: Produce the comparison workbook
Name: KSS_B1_vs_{RelativeName}.xlsx
Sheets: Shared Matches | Superkit Only | [RelativeName] Kit | Notes

### Comparison Workbook -- Required Output Standards
All comparison workbooks must include hyperlinks. A workbook without hyperlinks
is data, not a tool. This is non-negotiable.

Hyperlink construction rules:
- Match Name column: https://www.ancestry.com/dna/matches/{TESTER_GUID}/compare/{MATCH_GUID}
- Family Tree column: https://www.ancestry.com/discoveryui-matches/compare/{TESTER_GUID}/with/{MATCH_GUID}
- Common Ancestor column: same pattern as Family Tree

GUIDs are always present in every sheet and are the sole construction source for
all hyperlinks. Never attempt to read hyperlinks from source files -- openpyxl
read_only mode silently discards embedded hyperlink objects. Always build from GUIDs.

For comparison workbooks, the relevant TESTER_GUID depends on which sibling's
perspective the compare URL should use. Use the superkit's primary tester (Gerri,
GUID: 6E56EBFB-76A8-4342-AABB-1F9AF8A1746C) as the default tester anchor for
shared match and superkit-only sheets. For the external kit's own sheet, that kit's
GUID is the tester anchor.

### Step 4: Interpret the intersection
If the external kit's relationship is known:
- Paternal half-sibling: every shared match is almost certainly paternal
- Known paternal first cousin: shared matches narrow to the specific grandparent pair
- Known maternal relative: symmetric -- shared matches are maternal

If the relationship is unknown:
- Examine which Sibling Pattern subsets the shared matches fall into
- Check if any shared matches already have Line Assignments in the superkit
- Existing line assignments in shared matches are a hypothesis about the unknown person's line

### Step 5: Write back conclusions
Confirmed line assignments are entered manually into the superkit's Line Assignment column.
The Evidence Basis column records the source of the assignment (e.g., "PatHalfSib confirmed 2026-06").
The superkit is the conclusion layer. The comparison workbooks are the evidence layer.

### Temporal Currency Warning
The superkit reflects the Ancestry database as of 2026-06-05. When an external kit was
retrieved at a meaningfully different date, note this explicitly in the comparison workbook's
Notes sheet. A match absent from the superkit may simply be a newer tester.

---

## Known Comparison Kits

| Person | Relationship to Klein Super-Siblings | Relevant Line | Workbook Status |
|--------|--------------------------------------|---------------|-----------------|
| Paternal half-sibling | Half-sib (paternal) | Primary P/M sieve | Pending |
| Adrienne Balsky Peckler | 2nd cousin (confirmed) | Singer/Springer = MP | KSS_B1_vs_Adrienne_Peckler.xlsx built 2026-06-07; rebuild needed (hyperlinks missing) |
| Paternal 1st cousin(s) | 1st cousin | PP or PM | Pending |
| Maternal 2nd cousin | 2nd cousin | MP or MM | Pending |
| Shtetl research group | Unknown (suspected maternal) | TBD | Pending |

**Note on Adrienne Peckler:** Her pipeline configuration (root CLAUDE.md,
testers section) shows PM line as Samuel Singer + Minnie Jacobs + Tzvi Dov Springer.
These are Jacob Singer's parents -- the Klein super-siblings' maternal grandfather's parents.
Adrienne's connection runs through the Singer/Springer line (MP for the Klein super-siblings,
PM for Adrienne). She appears in the superkit with ThruLines at 52 cM Max Longest.
As a confirmed 2nd cousin, she is not expected to show overlapping matches with all four
siblings. The intersection of 11 shared matches is consistent with this relationship.
Troy60681 (her son, GUID: 21DA164D-DC18-4507-A753-61944E401D2F) must be excluded from
all Adrienne comparison work -- see Excluded GUIDs table.

**Adrienne comparison priorities (from first comparison run, 2026-06-07):**
- Arnold Schneider: ThruLines match in SK (57 cM, GSCL) + 37 cM for Adrienne. Only match
  with a documented common ancestor hypothesis. Highest priority for documentary follow-up.
- Rachelle Holden + S.V. (giggerus): both show GS·L pattern in SK and appear in Adrienne's
  filtered set. Same 3-sibling subset appearing in two independent matches is worth tracking.
- Surnames alone (e.g. Jacobs) are never sufficient to assign a line. Do not flag surname
  matches without corroborating DNA or documentary evidence.

---

## Ashkenazi Methodology -- Non-Negotiable Rules

1. Always use UNWEIGHTED total cM. Never TIMBER-adjusted (weighted) cM.
2. Platform relationship estimates are meaningless. Never reference them in analysis or output.
3. Leeds Method does not apply. Use spine/anchor methodology for line sorting.
4. ThruLines and Theory of Family Relativity are hypotheses only, never conclusions.
5. AScM filter: Longest >= 20 cM AND AScM >= 12. Both must be met. Applied at ingestion.
6. "Triangulation" = three individuals sharing an identical chromosomal segment.
   Do not use this term for cross-platform appearance. That is "cross-pool corroboration."
7. Pile-up regions on certain chromosomes produce elevated Ashkenazi background.
   Flag; do not treat as confirming a specific common ancestor without support.
   Chromosomal coordinates are required to identify a pile-up. Similar longest segment
   values across multiple matches are not evidence of a pile-up without coordinates.
8. DNA evidence never stands alone. Correlate with documentary evidence (GPS standard).
9. Anti-fabrication: never invent cM values, segment data, relationship conclusions, or sources.
10. GPS methodology governs all proof standards.
11. Surnames are search leads, not evidence. A match sharing a surname with a known
    ancestor line is not a candidate for that line without corroborating DNA or documentary
    evidence. Never flag a match on surname alone.

### Key Research Authorities
- Kitty Cooper: longest segment > 20 cM minimum for recent traceable Ashkenazi connection
- Jennifer Mendelsohn: no segment > 20 cM = likely untraceable; two segments > 30 cM ~3.3 generations to MRCA
- Adina Newman, Lara Diamond, Gil Bardige: AScM + longest segment framework
- Blaine Bettinger: Shared cM Project (apply with wider ranges under endogamy)

---

## File Naming Conventions

| File type | Pattern | Example |
|-----------|---------|---------|
| Superkit | Klein_Siblings_Superkit_{N}_B{B}.xlsx | Klein_Siblings_Superkit_1000_B1.xlsx |
| Individual sibling workbook | {First}_{Last}_DNA_Matches_Top{N}_Batch{B}.xlsx | Gerri_Klein_Taylor_DNA_Matches_Top1000_Batch1.xlsx |
| Comparison workbook | KSS_B{B}_vs_{RelativeName}.xlsx | KSS_B1_vs_Adrienne_Peckler.xlsx |

---

## Rebuild Triggers

Rebuild the superkit when:
- Any sibling workbook is updated with a new batch
- Export size is expanded (e.g., 1000 to 2000 per sibling)
- More than ~6 months have elapsed since the last retrieval

When rebuilding, retain the prior version. Do not overwrite. Increment batch number.

---

## Repo Location

https://github.com/davewilbur78/ancestry-dna-pipeline/tree/main/klein-superkit/

---

## What To Work On Next Session

1. Rebuild KSS_B1_vs_Adrienne_Peckler.xlsx with hyperlinks (omitted in first build)
2. Investigate Arnold Schneider -- ThruLines match in SK + 37 cM for Adrienne;
   seek documentary corroboration for ThruLines hypothesis before drawing conclusions
3. Track Rachelle Holden and S.V. (giggerus) -- GS·L pattern in SK, both appear in
   Adrienne filtered set; candidate MP line matches pending corroboration
4. Begin Line Assignment for the 11 ThruLines matches (strongest evidence available)
5. Analyze Jeannette Klein (186 cM, GSCL) and H.N. (129 cM, GSCL) -- both high priority
6. Identify and obtain paternal half-sibling kit (primary P/M sieve)
7. Commit methodology document to `methodology/` subfolder

---

## Session-Close Checklist

1. New decisions or direction changes? Update this file, bump version, commit.
2. New files produced? Commit them.
3. Write a CHANGELOG entry in klein-superkit/CHANGELOG.md.
4. Update "What To Work On Next Session" before closing.
