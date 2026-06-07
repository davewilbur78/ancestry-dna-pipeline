# Workbook Architecture Guide

**Project:** Klein Super-Siblings Superkit
**Audience:** Researchers working with the Klein super-siblings superkit and associated comparison workbooks

---

## Overview: Three Tiers of Workbooks

All analysis in this project uses three tiers of workbooks. Understanding what
each tier is for -- and what it is NOT for -- is essential.

```
Tier 1: Individual workbooks          (source of record)
Tier 2: Klein super-siblings superkit (standing reference)
Tier 3: Comparison workbooks          (analytical products)
```

---

## Tier 1: Individual Workbooks

**What they are:** The direct output of the DNA match enrichment pipeline. One
workbook per tester. These are the source files that the superkit is built from.

**The four sibling source files:**
- Gerri_Klein_Taylor_DNA_Matches_Top1000_Batch1.xlsx
- Susan_Klein_Beyer_DNA_Matches_Top1000_Batch1.xlsx
- Cynthia_KleinWilbur_DNA_Matches_Top1000_Batch1.xlsx
- Lee_Klein_DNA_Matches_Batch1_top1000.xlsx

When a known relative or research group member has their workbook built by the
pipeline, that workbook is also a Tier 1 workbook. It is NOT a sibling workbook,
but it is structurally identical.

### Column Schema (same for every Tier 1 workbook)

| Column | What it is | Key notes |
|--------|-----------|-----------|
| Match Name | Hyperlinked to compare page | URL is tester-dependent |
| Longest Segment | Longest shared segment in cM | **Primary metric for Ashkenazi analysis** |
| AScM | Unweighted cM / Segments | Live formula in source; recomputed fresh in superkit build |
| Unweighted cM | Total shared DNA, unweighted | **Always use this. Never weighted.** |
| Segments | Number of shared segments | |
| Weighted cM | TIMBER-adjusted figure | **Recorded but NEVER used in analysis** |
| Family Tree | Hyperlinked to match's tree | Same URL regardless of which sibling's workbook |
| Tree Size | Number of people in tree | |
| Common Ancestor | Hyperlinked to ThruLines | Only populated when Ancestry shows a hint |
| Line Assignment | Blank by default | Filled manually as evidence accumulates |
| Groups | Ancestry group tags | Encodes researcher's current knowledge |
| Notes | Researcher notes | |
| Match Side | Ancestry's Paternal/Maternal/Both | Only reliable when match has a linked tree |
| GUID | Match's Ancestry kit identifier | **Primary join key across all workbooks** |

### Color Coding (in individual workbooks and sibling reference sheets)

| Color | Condition | Meaning |
|-------|-----------|---------|
| Dark Green (#70AD47) bold | Longest >= 50 cM | High priority. Strong recent signal. |
| Medium Green (#A9D18E) | Longest 30-49 cM | Solid signal. Worth investigating. |
| Light Green (#E2EFDA) | Longest 20-29 cM | Passes filter. Watch for endogamy noise at low values. |
| Red | Longest < 20 OR AScM < 12 | Below filter. Not included in Superkit. Do not analyze in standard workflow. |

### The AScM Filter

Before any match is included in the Superkit, it must pass both conditions:
- **Longest Segment >= 20 cM** (the minimum for a traceable recent Ashkenazi connection)
- **AScM >= 12** (Unweighted cM / Number of Segments; screens out background endogamy noise)

Both conditions must be met. A match with a 25 cM longest segment but AScM of 8 fails.
A match with AScM of 15 but longest segment of 18 cM also fails.

Matches failing the filter are colored red and excluded from the Superkit sheet.
They remain accessible in the Cross-Kit Master with their full per-sibling data.

---

## Tier 2: The Klein Super-Siblings Superkit Workbook

**What it is:** A derived analytical reference. Built from the four individual
sibling workbooks. One row per unique GUID across all four siblings.

**What it is NOT:** The source of record. The individual workbooks remain authoritative.
When individual workbooks are updated (new batches, new exports), the superkit is
rebuilt from the updated sources.

**File naming:** Klein_Siblings_Superkit_{N}_B{B}.xlsx

### Sheet by Sheet

**Sheet 1: Superkit (primary research surface)**

One row per unique match GUID that passes the filter threshold in at least one
sibling's workbook. Sorted by Max Longest Segment descending.

Key columns unique to this sheet:

| Column | What it is |
|--------|-----------|
| Max Longest Seg | Highest Longest Segment across all siblings. Primary sort metric. |
| Sib Count | How many siblings share this match above filter threshold. |
| Sib Pattern | Which siblings share this match. G=Gerri S=Susan C=Cynthia L=Lee. Dot = absent. |
| cM Range | Max unweighted cM minus min unweighted cM across siblings (absent = 0). |
| Best AScM / Best Unwtd cM / Best Segments | Values from the "best-evidence sibling" -- the one with the highest Longest Segment for this match. |
| Line Assignment | Blank at build. Filled manually from comparison work. |
| Evidence Basis | Free text. Records how and when Line Assignment was determined. |

**Reading Sib Pattern:** `G.S.` means Gerri and Susan share this match above filter;
Cynthia and Lee do not. `GSCL` means all four. `.S..` means Susan only.
The pattern is sortable -- matches with the same pattern cluster together when sorted,
revealing groups of matches that the Klein super-siblings share with the same external individuals.

**Interpreting priority:**
- 1-2 siblings, Max Longest >= 30 cM: HIGH PRIORITY. Specific shared segment -- likely a real traceable recent connection.
- 3-4 siblings, Max Longest >= 30 cM: HIGH CONFIDENCE. Strong signal across most of the family.
- 4 siblings, Max Longest 20-25, low cM Range: LIKELY NOISE. Uniformly small values in all siblings = endogamy background.
- 1 sibling, Max Longest 20-22: MARGINAL. Real but specific, or barely above noise floor.

**Sheet 2: Cross-Kit Master (full analytical layer)**

All 3,264 unique GUIDs with no filter applied. For each sibling, there is a column
group showing their individual Longest Segment, Unweighted cM, Segments, AScM, and
a compare link to that sibling's view of the match. Use this sheet when you need to
see a specific match's complete per-sibling profile, or when investigating a match
that did not pass the Superkit threshold.

**Sheets 3-6: Gerri, Susan, Cynthia, Lee (reference copies)**

Convenience copies of each sibling's match data. Regenerated fresh during the superkit
build -- NOT copied from source files. Hyperlinks are constructed from GUIDs. AScM is
recomputed fresh. Same color coding as individual workbooks.

These are for reference during comparison sessions where having all data in one file
is useful. The source individual workbook files remain authoritative.

**Sheet 7: Stats**

Summary counts, tier distribution, noise floor estimate, source file names, retrieval
date, and build date. Consult this sheet first when opening the workbook after a break
to orient yourself on the current batch state.

**Sheet 8: Legend**

Color key, filter threshold definitions, AScM explanation, Sibling Pattern notation key,
snapshot warning, excluded GUID list, and methodology notes. Written so anyone opening
the workbook cold can orient themselves without prior context.

---

## Tier 3: Comparison Workbooks

**What they are:** Analytical products produced when an external kit (a known relative,
a research group member, an unknown match) is compared against the superkit.

**File naming:** KSS_B{B}_vs_{RelativeName}.xlsx

**Sheet structure (every comparison workbook):**

| Sheet | Contents |
|-------|----------|
| Shared Matches | Intersection: matches appearing in BOTH the superkit (above filter) and the external kit (above filter). Primary research surface. |
| Superkit Only | Superkit matches not found in the external kit above threshold. Useful for understanding what the external kit does not share. |
| [RelativeName] Kit | Clean copy of the external kit's data for reference during the session. |
| Notes | Observations, data currency notes, and research conclusions specific to this pairing. |

### Why the Intersection Is Powerful

**Paternal half-sibling comparison:** A paternal half-sibling shares ONLY paternal DNA.
Every match in the intersection is almost certainly paternal (either PP or PM). Every
superkit match absent from the intersection is almost certainly maternal (either MP or MM).
This is a near-complete maternal/paternal sieve in one comparison run.

**First cousin (known line):** A known first cousin shares one grandparent pair with
the Klein super-siblings. Shared matches narrow to the specific quadrant of that grandparent pair.
If the cousin is known to be on the PP (Henry Klein) line, their intersection with the
superkit resolves PP vs. all other lines for every shared match.

**Second and third cousins:** Corroborating signal only. Under Ashkenazi endogamy,
second and third cousin comparisons are noisier. Use as corroboration for assignments
already established, not as primary evidence for new assignments.

**Unknown match:** Examine the intersection for matches that already have Line Assignments
in the superkit. Clusters of already-assigned matches in the intersection are a hypothesis
about which line the unknown person connects through.

---

## The GUID -- Why It Matters

The GUID is the Ancestry kit identifier for each match. It is the only reliable join
key across all workbooks. Display names are unreliable because:
- The same person can appear under different names across sibling kits (managed accounts,
  name changes, married vs. maiden names)
- Ancestry account names change over time

Always join on GUID. Display name is for human reading only.
GUIDs must be normalized to uppercase and stripped of whitespace before any join operation.

---

## What the Comparison Workflow Looks Like in Practice

1. Build the external kit's individual workbook using the pipeline (root CLAUDE.md).
2. Load both the superkit and the external kit workbook.
3. Join on GUID to find the intersection.
4. For each shared match in the intersection, record both the superkit data (sibling
   pattern, Max Longest, existing Line Assignment) and the external kit data (how
   large the match is to the external person).
5. If the external kit's line is known, tentatively assign the shared matches to that line.
6. Record the assignment and evidence basis in the comparison workbook Notes sheet.
7. Enter confirmed assignments back into the superkit's Line Assignment and Evidence
   Basis columns (manual step -- researcher review always required).

The comparison workbooks accumulate. The superkit grows more resolved with each run.
Matches move from blank (unassigned) to one of PP/PM/MP/MM/Multiple. The Reference
Panel of confirmed individuals with known positions grows.

---

## Snapshot Warning

The superkit is a snapshot of the Ancestry database as of the source retrieval date
(Batch 1: 2026-06-05). It does not update automatically. A match absent from the
superkit may not have tested yet at retrieval time. Note the temporal gap explicitly
in comparison workbook Notes when external kits were retrieved at a different date.
Rebuild when approximately 6 months have elapsed or when new individual workbooks are available.

---

## GPS Standards Apply Throughout

DNA evidence never stands alone. Every line assignment conclusion must be correlated
with documentary evidence before it meets GPS proof standard. ThruLines suggestions
are hypotheses, not conclusions. Confidence language: Proved, Probable, Possible, Not Proved, Disproved.
