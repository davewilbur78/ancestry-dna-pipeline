# Tester Config: [Full Name]

Loaded by CLAUDE.md when ACTIVE TESTER = [Full Name].
This file holds everything specific to this kit. The pipeline and methodology are
generic and live in CLAUDE.md.

To start a new tester: copy this file to `testers/{firstname-lastname}.md`, fill in
every bracketed field, then point the Active Tester block in CLAUDE.md at it.

---

## Identity

TESTER: [Full Name]
KIT ID: [tester GUID from the Ancestry kit URL]
PLATFORM: AncestryDNA
ANCHOR: [closest known anchor kit, or "to be confirmed"]

---

## Line Anchors (surname -> quadrant -> color)

Quadrant colors are the fixed project convention (cool = paternal, warm = maternal).
Fill in the surname behind each quadrant for this tester.

  PP - Paternal-Paternal - [surname] line - green  (#E2EFDA)
  PM - Paternal-Maternal - [surname] line - blue   (#BDD7EE)
  MP - Maternal-Paternal - [surname] line - yellow (#FFEB9C)
  MM - Maternal-Maternal - [surname] line - coral  (#FCE4D6)
  Multiple                                 - purple (#E2CEEF)

---

## Group -> Line Mapping (for spreadsheet auto-fill)

  "[group tag]" in Groups -> PP - [surname]
  "[group tag]"           -> PM - [surname]
  "[group tag]"           -> MP - [surname]
  "[group tag]"           -> MM - [surname]
  Multiple of the above   -> Multiple (hold for investigation)
  None of the above       -> blank (unassigned)

Never force-assign a match that appears under multiple lines. Hold as Multiple.

---

## Great-Grandparent Couples (8-couple tier system)

  PP1: [couple]
  PP2: [couple]
  PM1: [couple]
  PM2: [couple]
  MP1: [couple]
  MP2: [couple]
  MM1: [couple]
  MM2: [couple]

---

## Platforms In Use

AncestryDNA (primary). [Add GEDmatch / FTDNA / MyHeritage if applicable.]

---

## Ancestry Groups Active

  [list group names already created in Ancestry, with their line]

---

## Research Priority Order

1. Dark green rows (longest segment 50+ cM) -- highest priority
2. Medium green rows (longest segment 30-50 cM)
3. Known line members
4. Matches with Common Ancestor predicted by Ancestry
5. Light green rows (longest segment 20-30 cM)
6. Red rows -- set aside unless specific reason to investigate

---

## Batch State

  Batch 1: [N] matches. Spreadsheet: [filename]
  Status: [API data status]. [hyperlink status].

---

## What's Next (this tester)

- [next steps]
