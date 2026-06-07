# Klein Super-Siblings Superkit

**Subfolder of:** [ancestry-dna-pipeline](https://github.com/davewilbur78/ancestry-dna-pipeline)

A multi-sibling DNA reference workbook and comparison analysis project for
Ashkenazi Jewish genetic genealogy research, following GPS methodology and
Ashkenazi-specific endogamy-aware standards.

---

## What This Is

The four children of Vernon Klein (1923-1976) and Mildred Singer (1927-2015)
have each tested on AncestryDNA. Their combined match data has been merged into
a single analytical reference called the **Klein super-siblings superkit**.

Rather than working from one sibling's match list at a time, the superkit treats
the four siblings together as a single super-sibling -- providing approximately
93-94% combined coverage of both parents' genomes. This makes it far more powerful
as a reference than any individual kit.

The project uses the superkit as a standing reference against which individual
workbooks for known relatives and research group members are compared. Comparison
runs produce line-sorting evidence: when a known paternal relative's matches
intersect with the superkit, those intersecting matches are almost certainly paternal.
Over time, comparison runs fill in the Line Assignment column of the superkit,
resolving which grandparental line each match connects through.

---

## The Four Siblings (Klein Super-Siblings)

| Sibling | Ancestry Name | Label |
|---------|--------------|-------|
| Gerri   | Gerri Taylor | G |
| Susan   | Susan Beyer  | S |
| Cynthia | Cynthia Wilbur (Managed by Rachel Sanda) | C |
| Lee     | Lee Klein    | L |

All four are full siblings. All four grandparents are common to all four siblings.
Their fan chart is documented in `family/fan-chart.md`.

---

## Current Status

- **Superkit:** Klein_Siblings_Superkit_1000_B1.xlsx (Batch 1, built 2026-06-07)
- **Matches in superkit:** 337 filter-passing; 3,264 total unique GUIDs
- **Line assignments filled:** 0 (fresh build, comparison work not yet begun)
- **ThruLines populated:** 11 matches
- **Comparison workbooks built:** 0 (pending)

---

## How This Repo Is Organized

```
klein-superkit/
├── CLAUDE.md          -- AI instruction file. Load this at the start of every AI session.
├── CHANGELOG.md       -- Session-by-session decision log.
├── README.md          -- This file.
├── family/
│   ├── fan-chart.md   -- Four-quadrant grandparental breakdown (PP/PM/MP/MM).
│   └── known-matches.md -- Known comparison kits and relationship notes (to be added).
├── methodology/
│   └── Klein_Sibling_Superkit_Methodology.md  -- Full build and analysis methodology (to be added).
├── docs/
│   └── workbook-guide.md -- Explanation of workbook architecture for researchers.
└── workbooks/
    └── superkit-state.md -- Current batch state (to be added; actual .xlsx files not in repo).
```

Note: The actual .xlsx workbook files are not stored in this repo. This repo holds
the project documentation, methodology, and AI instruction files. Workbooks are
stored locally by the researcher.

---

## Using With AI

**CLAUDE.md is the project brain.** At the start of any AI session working on this
project, load this file:

```
At the start of every conversation, fetch and read this file fully:
https://raw.githubusercontent.com/davewilbur78/ancestry-dna-pipeline/main/klein-superkit/CLAUDE.md
Read it completely before responding. Never rely on memory.
```

For workbook construction (building individual pipeline workbooks), also load the
root-level CLAUDE.md.

---

## Naming Note

Do not refer to these four siblings as "the Klein family." Henry Klein (paternal
grandfather) had four brothers. The Klein surname appears across multiple
grandparental lines and in unrelated matches. Always say **Klein super-siblings**
when referring specifically to these four siblings.

---

## Methodology

All work follows GPS standards and Ashkenazi-specific endogamy-aware methodology.
Never uses Leeds Method or platform relationship estimates. AScM filter (minimum 12)
and longest segment threshold (minimum 20 cM) applied throughout. Always uses
unweighted cM -- never Ancestry's TIMBER-adjusted figures.

See `methodology/Klein_Sibling_Superkit_Methodology.md` for the complete specification.

---

## Related

The root `ancestry-dna-pipeline/` project governs individual workbook construction
(the pipeline that produces the .xlsx files used as inputs here).
The `testers/` folder in the root project contains per-kit research notes for
individual testers including comparison kit candidates.
