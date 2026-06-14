#!/usr/bin/env python3
"""
build_v2.1.py — Design Round 3

Changes from v5:
  - Header: tester name back to row 1; row 2 = tab name + "Ancestry data retrieved DATE"
  - Batch Summary: no navy header; grey-banded box with dashed navy border, gap from header
  - Grey data columns: #707070 italic — visible on ALL tier backgrounds
  - Match Name width 28→36, RIGHT-aligned per user preference
  - Font sizes up everywhere: col headers 11, data rows 12, body text 12, section headers 13
  - "Line Assignment" → "Family Line Assignment" throughout (width 22)
  - Gil Bardige section: large burgundy banner, real call-to-action
  - Quick Start: inline rich text — column name refs in ALL CAPS BOLD UNDERLINE
  - Line Assignment call-to-action: dashed amber box system (amber=action, solid=info)
  - CONFIG block at top for easy kit switching
  - All data row heights adjusted for larger font
"""

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.cell.rich_text import TextBlock, CellRichText
from openpyxl.cell.text import InlineFont

# ══════════════════════════════════════════════════════════════════════════════
# KIT CONFIGURATION — update these for each new tester
# ══════════════════════════════════════════════════════════════════════════════
TESTER_NAME    = "Lesley Sterling"
BATCH_NUM      = "Batch 1"
SOURCE_FILE    = "/sessions/happy-upbeat-ride/mnt/uploads/Lesley_Sterling_DNA_Matches_Batch1.xlsx"
OUTPUT_DIR     = "/sessions/happy-upbeat-ride/mnt/Ancestry DNA Data extractor"
SCRIPT_VERSION = "2.1"
# OUTPUT_FILE is built automatically from tester name, match count, date, and version.
# No need to set it here.
# ══════════════════════════════════════════════════════════════════════════════

# ── PALETTE ───────────────────────────────────────────────────────────────────
NAVY       = "1F3864"
BLUE_SEC   = "2E75B6"
DARK_GRN   = "375623"
GIL_BG     = "D6E4F0"
GIL_FG     = "1F3864"
CALLOUT_BG = "FFF2CC"
CALLOUT_FG = "7F6000"
LBLUE_TEXT = "BDD7EE"
BURGUNDY   = "7B2D3E"
BURG_LIGHT = "F5E6EA"
BURG_DARK  = "4A1020"

# Fills
F_DARK    = PatternFill("solid", fgColor="70AD47")
F_MED     = PatternFill("solid", fgColor="A9D18E")
F_LIGHT   = PatternFill("solid", fgColor="E2EFDA")
F_LOW     = PatternFill("solid", fgColor="FFD7D7")
F_WHITE   = PatternFill("solid", fgColor="FFFFFF")
F_NAVY    = PatternFill("solid", fgColor=NAVY)
F_CALLOUT = PatternFill("solid", fgColor=CALLOUT_BG)
F_GIL     = PatternFill("solid", fgColor=GIL_BG)
F_BLUE    = PatternFill("solid", fgColor=BLUE_SEC)
F_BURG    = PatternFill("solid", fgColor=BURGUNDY)
F_BURGL   = PatternFill("solid", fgColor=BURG_LIGHT)
F_STATS   = PatternFill("solid", fgColor="F4F6FB")   # very light blue-grey for stats header
F_ACTION  = PatternFill("solid", fgColor="FFF5CC")   # light amber — action required bg

TIER_FILL = {"high": F_DARK, "solid": F_MED, "investigate": F_LIGHT,
             "watch": F_LOW, "low": F_LOW}

BLACK   = "000000"
WHITE   = "FFFFFF"
GREY_D  = "595959"   # column guide display only
LINK_C  = "1155CC"
NAVY_C  = NAVY
ACTION_COLOR = "9C6500"   # amber text for action callouts
ACTION_FILL  = "FFF5CC"   # light amber bg for action callouts

# Border helpers
DASH_SIDE   = Side(style='dashed', color=ACTION_COLOR)
THIN_SIDE   = Side(style='thin',   color=NAVY)
NO_SIDE     = Side(style=None)

def box_border(style='thin', color=NAVY):
    s = Side(style=style, color=color)
    return s

def apply_outer_box(ws, r1, r2, c1, c2, style='thin', color=NAVY):
    """Apply box border (outer edges only) to a cell range."""
    s = Side(style=style, color=color)
    n = Side(style=None)
    for ri in range(r1, r2+1):
        for ci in range(c1, c2+1):
            cell = ws.cell(row=ri, column=ci)
            t = s if ri == r1 else n
            b = s if ri == r2 else n
            l = s if ci == c1 else n
            r = s if ci == c2 else n
            cell.border = Border(top=t, bottom=b, left=l, right=r)

YT_LINK = "https://www.youtube.com/watch?v=0hGonlNLinM"

# ── COLUMN SCHEMA ─────────────────────────────────────────────────────────────
COLS = [
    ("match_name",        "Match Name",                36, False),  # widened 28→36
    ("flag",              "⚠",                          7, False),
    ("longest",           "Longest Segment",            15, False),
    ("ascm",              "AScM",                       9, False),
    ("unweighted_cm",     "Unweighted cM",              14, False),
    ("segments",          "Segments",                   10, False),
    ("weighted_cm",       "Weighted cM",                13, True),
    ("family_tree",       "Family Tree",                20, False),
    ("tree_size",         "Tree Size",                  10, False),
    ("common_ancestor",   "Common Ancestor",            22, False),
    ("line_assignment",   "Family Line Assignment",     22, False),  # renamed; widened
    ("groups",            "Groups",                     22, False),
    ("notes",             "Notes",                      40, False),
    ("match_side",        "Match Side",                 18, True),
    ("guid",              "GUID (Reference Anchor)",    38, True),
]
NCOLS   = len(COLS)
UWT_COL = get_column_letter(next(i+1 for i,(k,_,_,_) in enumerate(COLS) if k=="unweighted_cm"))
SEG_COL = get_column_letter(next(i+1 for i,(k,_,_,_) in enumerate(COLS) if k=="segments"))

# ── AUTO-DERIVE EXTRACTION DATE AND OUTPUT PATH ───────────────────────────────
import os, datetime

_mod_ts = os.path.getmtime(SOURCE_FILE)
_mod_dt = datetime.datetime.fromtimestamp(_mod_ts)
RETRIEVAL_DATE = _mod_dt.strftime("%B %-d, %Y")       # e.g. "June 12, 2026"
_date_compact  = _mod_dt.strftime("%Y%m%d")            # e.g. "20260612"
# OUTPUT filename built after n_total is known (see bottom of LOAD SOURCE block)

# ── LOAD SOURCE ───────────────────────────────────────────────────────────────
src    = openpyxl.load_workbook(SOURCE_FILE)
src_ws = src["DNA Matches"]

rows_raw = []
for row in src_ws.iter_rows(min_row=2):
    rd = {
        "match_name":      row[0].value,
        "match_url":       row[0].hyperlink.target if row[0].hyperlink else None,
        "longest":         row[1].value,
        "unweighted_cm":   row[3].value,
        "segments":        row[4].value,
        "weighted_cm":     row[5].value,
        "family_tree":     row[6].value,
        "family_tree_url": row[6].hyperlink.target if row[6].hyperlink else None,
        "tree_size":       row[7].value,
        "ca_url":          row[8].hyperlink.target if row[8].hyperlink else None,
        "line_assignment": row[9].value,
        "groups":          row[10].value,
        "notes":           row[11].value,
        "match_side":      row[12].value,
        "guid":            row[13].value,
    }
    uwt = rd["unweighted_cm"] or 0
    seg = rd["segments"] or 0
    av  = (uwt / seg) if seg > 0 else 0.0
    lng = rd["longest"] or 0
    pl, pa = lng >= 20, av >= 12
    if pl and pa:
        rd["tier"] = "high" if lng >= 50 else ("solid" if lng >= 30 else "investigate")
    elif pl or pa:
        rd["tier"] = "watch"
    else:
        rd["tier"] = "low"
    rd["flag"]    = "⚠️" if (rd["segments"] and rd["segments"] <= 2) else ""
    rd["asc_val"] = av
    rows_raw.append(rd)

def _sk(x): return x["longest"] or 0
rows_all      = sorted(rows_raw, key=_sk, reverse=True)
rows_priority = [rd for rd in rows_all if rd["tier"] in ("high","solid","investigate")]
rows_watch    = sorted([rd for rd in rows_raw if rd["tier"]=="watch"], key=_sk, reverse=True)

n_total = len(rows_raw)

# Build the unique output filename now that we have the match count
_first, _last = TESTER_NAME.split(" ", 1)
OUTPUT_FILE = os.path.join(
    OUTPUT_DIR,
    f"{_first}_{_last}_{n_total}matches_{_date_compact}_v{SCRIPT_VERSION}.xlsx"
)
n_high  = sum(1 for rd in rows_raw if rd["tier"]=="high")
n_solid = sum(1 for rd in rows_raw if rd["tier"]=="solid")
n_inv   = sum(1 for rd in rows_raw if rd["tier"]=="investigate")
n_green = n_high + n_solid + n_inv
n_watch = sum(1 for rd in rows_raw if rd["tier"]=="watch")
n_low   = sum(1 for rd in rows_raw if rd["tier"]=="low")

# ── GLOBAL HELPERS ────────────────────────────────────────────────────────────

def set_col_widths(ws):
    for i, (_, _, w, _) in enumerate(COLS, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

def navy_cell(ws, row, col, value, size=13, h=None, bold=True, span=NCOLS):
    c = ws.cell(row=row, column=col, value=value)
    c.fill = F_NAVY
    c.font = Font(name="Calibri", bold=bold, size=size, color=WHITE)
    c.alignment = Alignment(vertical="center")
    if span > 1:
        ws.merge_cells(start_row=row, start_column=col,
                       end_row=row, end_column=col+span-1)
    if h:
        ws.row_dimensions[row].height = h
    return c

def write_col_headers(ws, hr):
    """ALL CAPS, font 11, height 44, wrap_text."""
    ws.row_dimensions[hr].height = 44.0
    for i, (_, label, _, _) in enumerate(COLS, 1):
        c = ws.cell(row=hr, column=i, value=label.upper())
        c.fill = F_NAVY
        c.font = Font(name="Calibri", bold=True, size=11, color=WHITE)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

def banner(ws, subtitle, row=1):
    """Row 1 = tester name. Row 2 = tab context + retrieval date."""
    c = ws.cell(row=row, column=1, value=TESTER_NAME)
    c.fill = F_NAVY
    c.font = Font(name="Calibri", bold=True, size=20, color=WHITE)
    c.alignment = Alignment(vertical="center")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=NCOLS)
    ws.row_dimensions[row].height = 40.0
    c2 = ws.cell(row=row+1, column=1, value=subtitle)
    c2.fill = F_NAVY
    c2.font = Font(name="Calibri", size=11, color=LBLUE_TEXT)
    c2.alignment = Alignment(vertical="center")
    ws.merge_cells(start_row=row+1, start_column=1, end_row=row+1, end_column=NCOLS)
    ws.row_dimensions[row+1].height = 22.0
    return row + 2

def write_data_rows(ws, rows, start_row):
    for rd in rows:
        fill = TIER_FILL[rd["tier"]]
        dark = (rd["tier"] == "high")
        notes_val = rd.get("notes") or ""
        ws.row_dimensions[start_row].height = 40.0 if notes_val else 20.0

        for ci, (key, _, _, greyed) in enumerate(COLS, 1):
            c = ws.cell(row=start_row, column=ci)
            c.fill = fill
            c.alignment = Alignment(vertical="center")

            # Grey: #707070 italic — visible on ALL backgrounds, clearly secondary
            gc = Font(name="Calibri", size=12, color="707070", italic=True) if greyed else None

            std  = Font(name="Calibri", size=12, color=BLACK, bold=dark)
            link = Font(name="Calibri", size=12, color=(BLACK if dark else LINK_C),
                        underline="single", bold=dark)

            if key == "match_name":
                c.value = rd["match_name"]
                c.alignment = Alignment(horizontal="right", vertical="center")
                if rd["match_url"]:
                    c.hyperlink = rd["match_url"]
                    c.font = link
                else:
                    c.font = std

            elif key == "flag":
                c.value = rd["flag"]
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.font = (Font(name="Calibri", size=13, bold=True, color="FF0000")
                          if rd["flag"] else Font(name="Calibri", size=12, color=BLACK))

            elif key == "longest":
                c.value = rd["longest"]
                c.number_format = "0"
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.font = std

            elif key == "ascm":
                c.value = f"=IFERROR({UWT_COL}{start_row}/{SEG_COL}{start_row},\"\")"
                c.number_format = "0.0"
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.font = std

            elif key == "unweighted_cm":
                c.value = rd["unweighted_cm"]
                c.number_format = "0"
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.font = std

            elif key == "segments":
                c.value = rd["segments"]
                c.number_format = "0"
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.font = std

            elif key == "weighted_cm":
                c.value = rd["weighted_cm"]
                c.number_format = "0"
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.font = gc

            elif key == "family_tree":
                if rd["family_tree_url"]:
                    c.value = rd["family_tree"] or "View"
                    c.hyperlink = rd["family_tree_url"]
                    c.font = link
                elif rd["family_tree"]:
                    c.value = rd["family_tree"]
                    c.font = std
                c.alignment = Alignment(horizontal="center", vertical="center")

            elif key == "tree_size":
                c.value = rd["tree_size"]
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.font = std

            elif key == "common_ancestor":
                if rd["ca_url"]:
                    c.value = "\U0001f465"
                    c.hyperlink = rd["ca_url"]
                    c.font = Font(name="Calibri", size=14,
                                  color=(BLACK if dark else LINK_C),
                                  underline="single", bold=dark)
                else:
                    c.value = ""
                    c.font = std
                c.alignment = Alignment(horizontal="center", vertical="center")

            elif key == "line_assignment":
                c.value = rd["line_assignment"]
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.font = std

            elif key == "groups":
                c.value = rd["groups"]
                c.font = std

            elif key == "notes":
                c.value = rd["notes"]
                c.font = std
                c.alignment = Alignment(wrap_text=True, vertical="top")

            elif key == "match_side":
                c.value = rd["match_side"]
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.font = gc

            elif key == "guid":
                c.value = rd["guid"]
                c.font = gc

        start_row += 1
    return start_row


# ══════════════════════════════════════════════════════════════════════════════
# BUILD WORKBOOK
# ══════════════════════════════════════════════════════════════════════════════
wb = Workbook()
wb.remove(wb.active)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1: START HERE
# ══════════════════════════════════════════════════════════════════════════════
ws_s = wb.create_sheet("Start Here")
ws_s.column_dimensions["A"].width = 5
ws_s.column_dimensions["B"].width = 18
ws_s.column_dimensions["C"].width = 90
ws_s.column_dimensions["D"].width = 87

r = 1

def sh_white_row():
    """Ensure all 4 cols get explicit white — prevents fill bleed."""
    for ci in range(1, 5):
        ws_s.cell(row=r, column=ci).fill = F_WHITE

def sh_navy(val, size=13, h=28, span=4):
    global r
    c = ws_s.cell(row=r, column=1, value=val)
    c.fill = F_NAVY
    c.font = Font(name="Calibri", bold=True, size=size, color=WHITE)
    c.alignment = Alignment(vertical="center")
    if span > 1:
        ws_s.merge_cells(start_row=r, start_column=1, end_row=r, end_column=span)
    ws_s.row_dimensions[r].height = h
    r += 1

def sh_body(val, h=46, color=BLACK, fill=None, size=12, italic=False, bold=False,
            col=1, span=4):
    global r
    row_fill = fill or F_WHITE
    c = ws_s.cell(row=r, column=col, value=val)
    c.font = Font(name="Calibri", size=size, color=color, italic=italic, bold=bold)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    c.fill = row_fill
    if span > 1:
        ws_s.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col+span-1)
    ws_s.row_dimensions[r].height = h
    r += 1
    return c

def sh_blank(h=5):
    global r
    sh_white_row()
    ws_s.row_dimensions[r].height = h
    r += 1

def sh_lbl_body(lbl, txt, h=26, lbl_fill=None, lbl_color=NAVY_C,
                lbl_bold=True, lbl_size=14, txt_size=12, txt_color=NAVY_C):
    global r
    ca = ws_s.cell(row=r, column=1, value=lbl)
    ca.fill = lbl_fill or F_WHITE
    ca.font = Font(name="Calibri", bold=lbl_bold, size=lbl_size, color=lbl_color)
    ca.alignment = Alignment(horizontal="center", vertical="center")
    ws_s.cell(row=r, column=2).fill = F_WHITE
    cc = ws_s.cell(row=r, column=3, value=txt)
    cc.fill = F_WHITE
    cc.font = Font(name="Calibri", size=txt_size, color=txt_color)
    cc.alignment = Alignment(wrap_text=True, vertical="center")
    ws_s.cell(row=r, column=4).fill = F_WHITE
    ws_s.row_dimensions[r].height = h
    r += 1

def sh_lbl_body_rich(lbl, rich_val, h=26, txt_size=12):
    """Variant of sh_lbl_body that accepts a CellRichText value for inline formatting."""
    global r
    ca = ws_s.cell(row=r, column=1, value=lbl)
    ca.fill = F_WHITE
    ca.font = Font(name="Calibri", bold=True, size=14, color=NAVY_C)
    ca.alignment = Alignment(horizontal="center", vertical="center")
    ws_s.cell(row=r, column=2).fill = F_WHITE
    cc = ws_s.cell(row=r, column=3)
    cc.value = rich_val
    cc.fill = F_WHITE
    cc.alignment = Alignment(wrap_text=True, vertical="center")
    ws_s.cell(row=r, column=4).fill = F_WHITE
    ws_s.row_dimensions[r].height = h
    r += 1

# Rich text helpers for Quick Start column references
def _rt(text_parts):
    """Build a CellRichText. text_parts = list of (text, bold, underline, caps)."""
    blocks = []
    for txt, bold, uline, caps in text_parts:
        display = txt.upper() if caps else txt
        ifont = InlineFont(
            rFont="Calibri", sz=12,
            color=NAVY_C,
            b=bold, u=("single" if uline else None)
        )
        blocks.append(TextBlock(ifont, display))
    return CellRichText(*blocks)

def _col(name):
    """Shortcut: return rich-text span for a column name (CAPS BOLD UNDERLINE)."""
    return (name, True, True, True)

def _txt(t):
    """Shortcut: return rich-text span for plain body text."""
    return (t, False, False, False)

# ── Identity Banner ──
c = ws_s.cell(row=r, column=1, value=TESTER_NAME)
c.fill = F_NAVY
c.font = Font(name="Calibri", bold=True, size=20, color=WHITE)
c.alignment = Alignment(vertical="center")
ws_s.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
ws_s.row_dimensions[r].height = 40.0; r += 1

c = ws_s.cell(row=r, column=1,
    value=f"DNA Match Analysis  ·  {BATCH_NUM}  ·  Ancestry data retrieved {RETRIEVAL_DATE}")
c.fill = F_NAVY
c.font = Font(name="Calibri", size=11, color=LBLUE_TEXT)
c.alignment = Alignment(vertical="center")
ws_s.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
ws_s.row_dimensions[r].height = 22.0; r += 1
sh_blank(h=8)

# ── Intro ──
sh_body(
    "When you open your AncestryDNA match list, you see hundreds of names with no clear way "
    "to know where to start. This workbook changes that. It pulls the segment data Ancestry "
    "keeps off the main screen, calculates a quality score called AScM, and uses those two "
    "numbers to sort your matches by what is actually worth your time. "
    "The Priority Matches tab is where you begin.",
    h=66, color=BLACK
)
sh_body(
    "Important: this workbook is a snapshot. It does not sync with Ancestry in either "
    "direction. Changes you make here stay here only. To record your research in both places, "
    "enter it in both — see Working Between This File and Ancestry below.",
    h=46, color=CALLOUT_FG, fill=F_CALLOUT
)
sh_blank(h=14)

# ── Quick Start ──
sh_navy("QUICK START")

# Step 1 - plain text
sh_lbl_body(1, "Open the Priority Matches tab.", h=26)
sh_blank(h=5)

# Step 2 - plain text
sh_lbl_body(2,
    "Start at the top — the dark green rows are your highest-priority matches: "
    "a longest segment of 50 cM or more.", h=26)
sh_blank(h=5)

# Step 3 - plain text
sh_lbl_body(3, "Click a match name to open their Ancestry profile comparison page.", h=26)
sh_blank(h=5)

# Step 4 - rich text: COMMON ANCESTOR in caps bold underline
sh_lbl_body_rich(4,
    _rt([
        _txt("Check the "),
        _col("Common Ancestor"),
        _txt(" column. If you see the \U0001f465 icon, "
             "click it to open Ancestry's ThruLines view for that match.")
    ]), h=26)
sh_blank(h=5)

# Step 5 - rich text: NOTES and FAMILY LINE ASSIGNMENT in caps bold underline
sh_lbl_body_rich(5,
    _rt([
        _txt("If you find a connection, write it in the "),
        _col("Notes"),
        _txt(" column. Fill in "),
        _col("Family Line Assignment"),
        _txt(" when you know which branch the match belongs to.")
    ]), h=26)
sh_blank(h=5)

# Step 6 - plain text
sh_lbl_body(6, "Work through all dark green rows, then medium green, then light green.", h=26)
sh_blank(h=5)

# Step 7 - plain text
sh_lbl_body(7,
    "If you see someone you already know anywhere in the list — even in the "
    "pink rows — add a note with the relationship. Known relatives can appear "
    "at any color tier.", h=46)
sh_blank(h=5)

c = ws_s.cell(row=r, column=1,
    value="The explanations below are here when you want to understand the why. "
          "You do not need to read them to get started.")
c.fill = F_WHITE
c.font = Font(name="Calibri", size=12, color="808080")
c.alignment = Alignment(wrap_text=True, vertical="top")
ws_s.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
ws_s.row_dimensions[r].height = 25.0; r += 1
sh_blank(h=14)

# ── TIMBER ──
sh_navy("A NOTE ON THE TWO CM COLUMNS: WEIGHTED VS. UNWEIGHTED")
sh_body(
    "Ancestry displays a Weighted cM figure that has been processed by an algorithm called "
    "TIMBER. TIMBER was designed to identify DNA segments that appear across many thousands "
    "of people from the same ancestral population — pile-ups — and reduce their weight, "
    "on the theory that shared DNA reflecting broad population membership is not useful "
    "for tracing a specific ancestor.",
    h=46, color=BLACK
)
sh_blank(h=5)
sh_body(
    "For Ashkenazi Jewish testers, TIMBER may over-correct. Because Ashkenazi Jews married "
    "within the community for many generations, what TIMBER identifies as a pile-up may still "
    "point to a genuine genealogical connection. This workbook uses Unweighted cM throughout. "
    "The Weighted cM column is greyed out and kept only for reference.",
    h=66, color=BLACK
)
sh_blank(h=14)

# ── Two Numbers ──
sh_navy("THE TWO NUMBERS THAT MATTER")

# --- helper: colored tier row for "two numbers" section ---
def two_num_tier_row(label, desc, fill, fc, h=22):
    global r
    ca = ws_s.cell(row=r, column=1, value=label)
    ca.fill = fill
    ca.font = Font(name="Calibri", bold=True, size=11, color=fc)
    ca.alignment = Alignment(horizontal="center", vertical="center")
    ws_s.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    cc = ws_s.cell(row=r, column=3, value=desc)
    cc.fill = F_WHITE
    cc.font = Font(name="Calibri", size=12, color=NAVY_C)
    cc.alignment = Alignment(wrap_text=True, vertical="center")
    ws_s.cell(row=r, column=4).fill = F_WHITE
    ws_s.row_dimensions[r].height = h
    r += 1

# --- LONGEST SEGMENT (CM) ---
c = ws_s.cell(row=r, column=1, value="LONGEST SEGMENT (CM)")
c.fill = PatternFill("solid", fgColor=DARK_GRN)
c.font = Font(name="Calibri", bold=True, size=12, color=WHITE)
ws_s.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
ws_s.row_dimensions[r].height = 22.0; r += 1

sh_body(
    "A single long segment is the strongest signal of a recent shared ancestor. In Ashkenazi "
    "DNA research, where centuries of marrying within the community means everyone shares some "
    "DNA with everyone, a long segment is hard to fake. Short scattered segments are common "
    "background noise; a long one points to something specific. "
    "The minimum threshold for this workbook is 20 cM.",
    h=66, color=BLACK
)
sh_blank(h=5)

# Tier rows — showing the actual colors used in the data tabs
two_num_tier_row("50+ cM", "High priority. Start here.", F_DARK, WHITE)
two_num_tier_row("30-49 cM", "Solid genealogical signal.", F_MED, BLACK)
two_num_tier_row("20-29 cM", "Worth investigating.", F_LIGHT, "375623")
two_num_tier_row("Under 20 cM", "Does not meet the minimum criterion.", F_LOW, "833C00")
sh_blank(h=5)

sh_body(
    "Real example: two matches, both 40 cM total, both with a longest segment of 22 cM. "
    "Match 1 has 2 segments (AScM = 20): research priority. Match 2 has 10 segments "
    "(AScM = 4): almost certainly untraceable background sharing. AScM is the number "
    "that tells them apart.",
    h=46, color=CALLOUT_FG, fill=F_CALLOUT
)
sh_blank(h=14)

# --- ASCM ---
c = ws_s.cell(row=r, column=1, value="ASCM  (UNWEIGHTED cM / SEGMENTS)")
c.fill = PatternFill("solid", fgColor=DARK_GRN)
c.font = Font(name="Calibri", bold=True, size=12, color=WHITE)
ws_s.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
ws_s.row_dimensions[r].height = 22.0; r += 1

sh_body(
    "AScM is Ancestry Segment cM: the average size of each shared segment "
    "(Unweighted cM divided by Segments). This workbook calculates it automatically. "
    "A higher AScM means the shared DNA is concentrated into fewer, larger segments — "
    "a stronger signal pointing to a real genealogical connection rather than broad "
    "Ashkenazi background sharing. Threshold: 12 to pass.",
    h=80, color=BLACK
)
sh_blank(h=5)

two_num_tier_row("20+", "Unusually strong segment quality.", F_DARK, WHITE)
two_num_tier_row("12-20", "Typical range for passing Ashkenazi matches.", F_MED, BLACK)
two_num_tier_row("12-20", "Light green tier also passes at AScM 12+.", F_LIGHT, "375623")
two_num_tier_row("Under 12", "Does not meet the AScM criterion.", F_LOW, "833C00")
sh_blank(h=14)
sh_blank(h=14)

# ── Color Coding ──
sh_navy("COLOR CODING AT A GLANCE")
ccodes = [
    (F_DARK,  WHITE,    "Longest segment 50+ cM",
     "High priority. Both filter conditions passed comfortably. Start your research here."),
    (F_MED,   BLACK,    "Longest segment 30-49 cM",
     "Solid genealogical signal. Passes the filter; investigate after dark green."),
    (F_LIGHT, "375623", "Longest segment 20-29 cM",
     "Passes the filter. Worth investigating, especially if a tree or ThruLines hint is present."),
    (F_LOW,   "833C00", "Watch List  ·  Low Priority",
     "Does not meet one or both research criteria. Likely difficult to trace. Kept for reference in the Watch List and All Matches tabs."),
]
for fill_cc, fc, label, desc in ccodes:
    ca = ws_s.cell(row=r, column=1, value=label)
    ca.fill = fill_cc
    ca.font = Font(name="Calibri", bold=True, size=11, color=fc)
    ca.alignment = Alignment(horizontal="center", vertical="center")
    ws_s.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    cc_cell = ws_s.cell(row=r, column=3, value=desc)
    cc_cell.fill = F_WHITE
    cc_cell.font = Font(name="Calibri", size=12, color=NAVY_C)
    cc_cell.alignment = Alignment(wrap_text=True, vertical="center")
    ws_s.cell(row=r, column=4).fill = F_WHITE
    ws_s.row_dimensions[r].height = 26.0; r += 1
sh_blank(h=14)

# ── Column Guide ──
sh_navy("COLUMN GUIDE")
col_guide = [
    ("70AD47", WHITE,      "A   Match Name",
     "The match's display name on Ancestry, linked to their profile comparison page. Click to open a side-by-side view.",
     False),
    ("FFE699", CALLOUT_FG, "B   ⚠️  Flag",
     "⚠️ appears when a match shares only 1 or 2 DNA segments. A single long segment can sometimes be an artifact rather than a genuine genealogical connection. Worth extra scrutiny before drawing conclusions.",
     False),
    ("70AD47", WHITE,      "C   Longest Segment",
     "The longest single DNA segment you share, in centiMorgans. The primary filter criterion. Not adjusted by TIMBER.",
     False),
    ("70AD47", WHITE,      "D   AScM",
     "Average centiMorgans per segment (Unweighted cM / Segments). Live formula, recalculates automatically. Threshold: 12 to pass the filter.",
     False),
    ("A9D18E", BLACK,      "E   Unweighted cM",
     "Total shared DNA before TIMBER adjustment. The figure used in all calculations in this workbook.",
     False),
    ("E2EFDA", BLACK,      "F   Segments",
     "Number of distinct DNA segments you share. Used in the AScM formula.",
     False),
    ("DCDCDC", GREY_D,     "G   Weighted cM",
     "Ancestry's TIMBER-adjusted figure. Greyed out; included for reference only. Not used in any calculation here.",
     False),
    ("2E75B6", WHITE,      "H   Family Tree",
     "Linked to the Ancestry compare view showing both trees and any ThruLines suggestions.",
     False),
    ("9DC3E6", BLACK,      "I   Tree Size",
     "Number of people in the match's linked tree. A larger tree improves the odds of finding a documented connection.",
     False),
    ("2E75B6", WHITE,      "J   Common Ancestor",
     "Shows the persons icon when Ancestry has a ThruLines suggestion. Click the icon to open the ThruLines page. This is a hypothesis, not a confirmed relationship.",
     False),
    (ACTION_FILL, ACTION_COLOR, "K   Family Line Assignment",
     "Which branch of the family this match belongs to: PP, PM, MP, MM, or Multiple. This column is yours to fill in as you research each match.",
     True),   # <-- action_required=True → dashed border
    ("4472C4", WHITE,      "L   Groups",
     "Ancestry color-dot group tags assigned to this match.",
     False),
    ("D9E1F2", BLACK,      "M   Notes",
     "Notes from Ancestry's notes field, plus any you add directly in this column.",
     False),
    ("DCDCDC", GREY_D,     "N   Match Side",
     "Ancestry's paternal/maternal/unassigned call. Greyed out — unreliable for fully Ashkenazi testers.",
     False),
    ("F2F2F2", GREY_D,     "O   GUID (Reference Anchor)",
     "Ancestry's internal unique identifier for each match. The backbone of all hyperlinks. You will never need to read it.",
     False),
]
for hex_fill, fc, label, desc, action_req in col_guide:
    row_r = r
    ca = ws_s.cell(row=r, column=1, value=label)
    ca.fill = PatternFill("solid", fgColor=hex_fill)
    ca.font = Font(name="Calibri", bold=True, size=11, color=fc)
    ca.alignment = Alignment(horizontal="center", vertical="center")
    ws_s.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    cc_cell = ws_s.cell(row=r, column=3, value=desc)
    cc_cell.fill = F_WHITE
    cc_cell.font = Font(name="Calibri", size=12, color=NAVY_C)
    cc_cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws_s.cell(row=r, column=4).fill = F_WHITE
    h_val = 46.0 if len(desc) > 80 else 26.0
    ws_s.row_dimensions[r].height = h_val
    r += 1
    if action_req:
        # Dashed amber border around the entire row (cols 1-4) = "action required" signal
        apply_outer_box(ws_s, row_r, row_r, 1, 4, style='dashed', color=ACTION_COLOR)
sh_blank(h=14)

# ── Family Line Assignment ──
sh_navy("FAMILY LINE ASSIGNMENT: QUADRANT COLORS")

# Action-required callout (dashed amber box = "your work here")
action_row = r
sh_body(
    "Your work — As you identify which branch of the family a match belongs to, "
    "enter the code in the Family Line Assignment column (column K). "
    "This is the column you fill in yourself; nothing here is pre-populated.",
    h=52, color=ACTION_COLOR, fill=PatternFill("solid", fgColor=ACTION_FILL)
)
apply_outer_box(ws_s, action_row, action_row, 1, 4, style='dashed', color=ACTION_COLOR)
sh_blank(h=5)

sh_body(
    "When you confirm which branch a match belongs to, enter the code below. "
    "Cool colors (green, blue) = paternal lines. Warm colors (yellow, coral) = maternal lines.",
    h=46, color=NAVY_C
)
sh_blank(h=5)
quads = [
    ("PP   Paternal-Paternal", "E2EFDA", "375623", "Your father's father's side"),
    ("PM   Paternal-Maternal", "BDD7EE", "1F3864", "Your father's mother's side"),
    ("MP   Maternal-Paternal", "FFEB9C", "7F6000", "Your mother's father's side"),
    ("MM   Maternal-Maternal", "FCE4D6", "833C00", "Your mother's mother's side"),
    ("Multiple",               "E2CEEF", "3D1152", "Spans more than one branch; investigate separately"),
]
for label, bg, fc, desc in quads:
    ca = ws_s.cell(row=r, column=1, value=label)
    ca.fill = PatternFill("solid", fgColor=bg)
    ca.font = Font(name="Calibri", bold=True, size=11, color=fc)
    ca.alignment = Alignment(horizontal="center", vertical="center")
    ws_s.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    cc_cell = ws_s.cell(row=r, column=3, value=desc)
    cc_cell.fill = F_WHITE
    cc_cell.font = Font(name="Calibri", size=12, color=NAVY_C)
    cc_cell.alignment = Alignment(vertical="center")
    ws_s.cell(row=r, column=4).fill = F_WHITE
    ws_s.row_dimensions[r].height = 22.0; r += 1
sh_blank(h=14)

# ── Working Between File and Ancestry ──
sh_navy("WORKING BETWEEN THIS FILE AND ANCESTRY")
sh_body(
    "This workbook and Ancestry do not sync. They are separate tools that complement each other.\n\n"
    "Work here: analyze your matches, assign family lines, make sense of the numbers.\n\n"
    "Record there: when you have worked out a family line assignment, go into Ancestry and tag "
    "that match with the appropriate color group. When you add a note here, copy it to "
    "Ancestry's note field on that match's profile.\n\n"
    "The workbook is your scratch pad. Ancestry is your record of assignment.",
    h=105, color=NAVY_C
)
sh_blank(h=14)

# ── Gil Bardige — redesigned for impact ──
sh_blank(h=6)

# Big burgundy banner — the call to action
c = ws_s.cell(row=r, column=1)
c.value = "▶ WATCH THIS PRESENTATION"
c.hyperlink = YT_LINK
c.fill = F_BURG
c.font = Font(name="Calibri", bold=True, size=24, color=WHITE)
c.alignment = Alignment(horizontal="center", vertical="center")
ws_s.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
ws_s.row_dimensions[r].height = 60.0; r += 1

# Presentation title row
c = ws_s.cell(row=r, column=1,
    value='"Help! I Got My DNA Results and I\'m Confused, Part 1.25"   —   Gil Bardige')
c.hyperlink = YT_LINK
c.fill = F_BURGL
c.font = Font(name="Calibri", bold=True, size=13, color=BURGUNDY)
c.alignment = Alignment(horizontal="center", vertical="center")
ws_s.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
ws_s.row_dimensions[r].height = 32.0; r += 1

# Description
c = ws_s.cell(row=r, column=1,
    value="Gil Bardige teaches DNA analysis for Jewish genealogy researchers. "
          "His videos cover how to read your match list, why AScM matters, and "
          "how to work effectively under Ashkenazi endogamy. The methodology behind "
          "the filter and tier system in this workbook is built on his work.")
c.fill = F_BURGL
c.font = Font(name="Calibri", size=12, color=BURG_DARK)
c.alignment = Alignment(wrap_text=True, vertical="top")
ws_s.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
ws_s.row_dimensions[r].height = 52.0; r += 1
sh_blank(h=14)

# ── A Few Things Worth Knowing ──
sh_navy("A FEW THINGS WORTH KNOWING")
bullets = [
    ("You may see people you know in the pink rows. The tiers are a triage tool for unknown "
     "matches, not a judgment about every match in your list. A known cousin sharing 14 cM "
     "across 5 segments is still your cousin. If you recognize someone, add a note.", 46),
    ("DNA evidence never stands alone. A match in a color tier tells you where to look, "
     "not what you have found. Always correlate with documentary evidence before drawing conclusions.", 46),
    ("ThruLines and Common Ancestor suggestions are hypotheses, not conclusions. Ancestry builds "
     "them from trees, which contain errors. Treat them as research leads.", 46),
    ("Ancestry's relationship estimates (2nd cousin, 3rd cousin, etc.) are not reliable for "
     "Ashkenazi testers. Work from the numbers in this workbook instead.", 46),
    ("The pink rows are not a dead end. They may become useful as trees grow or when a match "
     "adds documentation.", 26),
    ("Match Side is greyed out because it is unreliable for fully Ashkenazi testers. "
     "Ancestry cannot cleanly phase endogamous DNA. More useful for testers with mixed heritage.", 46),
]
for txt, h in bullets:
    ca = ws_s.cell(row=r, column=1, value="•")
    ca.fill = F_WHITE
    ca.font = Font(name="Calibri", size=14, color=NAVY_C)
    ca.alignment = Alignment(horizontal="center", vertical="top")
    ws_s.cell(row=r, column=2).fill = F_WHITE
    cc_cell = ws_s.cell(row=r, column=3, value=txt)
    cc_cell.fill = F_WHITE
    cc_cell.font = Font(name="Calibri", size=12, color=NAVY_C)
    cc_cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws_s.cell(row=r, column=4).fill = F_WHITE
    ws_s.row_dimensions[r].height = h; r += 1
    sh_blank(h=5)

ws_s.freeze_panes = "A3"


# ══════════════════════════════════════════════════════════════════════════════
# DATA TAB BUILDER
# ══════════════════════════════════════════════════════════════════════════════

def build_data_tab(ws, rows, tab_label, watch_note=False, stats=False):
    set_col_widths(ws)
    subtitle = f"{tab_label}  ·  Ancestry data retrieved {RETRIEVAL_DATE}"
    next_r = banner(ws, subtitle)

    if stats:
        # Small gap so stats block doesn't visually attach to navy header
        ws.row_dimensions[next_r].height = 6.0; next_r += 1

        stats_start = next_r

        # "BATCH SUMMARY" header — light grey fill, navy text, no navy bg
        for ci in range(1, NCOLS+1):
            ws.cell(row=next_r, column=ci).fill = F_STATS
        c = ws.cell(row=next_r, column=1, value="BATCH SUMMARY")
        c.fill = F_STATS
        c.font = Font(name="Calibri", bold=True, size=12, color=NAVY)
        c.alignment = Alignment(vertical="center", indent=1)
        ws.merge_cells(start_row=next_r, start_column=1, end_row=next_r, end_column=NCOLS)
        ws.row_dimensions[next_r].height = 26.0; next_r += 1

        # Small internal gap
        for ci in range(1, NCOLS+1):
            ws.cell(row=next_r, column=ci).fill = F_WHITE
        ws.row_dimensions[next_r].height = 4.0; next_r += 1

        stat_rows = [
            (None,    "Total matches retrieved",          n_total, None),
            (None,    None,                               None,    None),
            (F_DARK,  "High priority  (50+ cM)",          n_high,  WHITE),
            (F_MED,   "Solid signal  (30-49 cM)",         n_solid, BLACK),
            (F_LIGHT, "Worth investigating  (20-29 cM)",  n_inv,   "375623"),
            (None,    "All priority matches",             n_green, NAVY_C),
            (None,    None,                               None,    None),
            (F_LOW,   "Watch List",                       n_watch, "833C00"),
            (None,    "Low priority",                     n_low,   NAVY_C),
        ]
        for sfill, label, val, fc in stat_rows:
            ws.row_dimensions[next_r].height = 20.0
            if label is None:
                ws.row_dimensions[next_r].height = 5.0
                next_r += 1
                continue
            row_fill = sfill if sfill else F_WHITE
            bold_it  = (sfill is not None)
            for ci in range(1, NCOLS+1):
                ws.cell(row=next_r, column=ci).fill = row_fill
            ca = ws.cell(row=next_r, column=1, value=label)
            ca.fill = row_fill
            ca.font = Font(name="Calibri", size=12, bold=bold_it, color=fc or NAVY_C)
            ca.alignment = Alignment(vertical="center", indent=2)
            ws.merge_cells(start_row=next_r, start_column=1, end_row=next_r, end_column=3)
            cv = ws.cell(row=next_r, column=5, value=val)
            cv.fill = row_fill
            cv.font = Font(name="Calibri", size=12, bold=bold_it, color=fc or NAVY_C)
            cv.number_format = "#,##0"
            cv.alignment = Alignment(horizontal="right", vertical="center")
            next_r += 1

        stats_end = next_r - 1
        # Dashed navy box around the entire stats block
        apply_outer_box(ws, stats_start, stats_end, 1, NCOLS, style='thin', color=NAVY)

        ws.row_dimensions[next_r].height = 8.0; next_r += 1

    elif watch_note:
        note = (
            "These matches came close but did not meet both research criteria. "
            "A match lands here when it passes one condition — either a longest "
            "segment of 20 cM or more, or an AScM of 12 or more — but not both. "
            "Not a priority now, but worth keeping an eye on as trees grow."
        )
        c = ws.cell(row=next_r, column=1, value=note)
        c.fill = F_CALLOUT
        c.font = Font(name="Calibri", size=12, color=CALLOUT_FG)
        c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.merge_cells(start_row=next_r, start_column=1, end_row=next_r, end_column=NCOLS)
        ws.row_dimensions[next_r].height = 60.0
        next_r += 1
        ws.row_dimensions[next_r].height = 5.0; next_r += 1

    hdr_row = next_r
    write_col_headers(ws, hdr_row)
    next_r += 1

    data_start = next_r
    last_row = write_data_rows(ws, rows, data_start) - 1

    ws.freeze_panes = ws.cell(row=data_start, column=1)

    if last_row >= data_start:
        ws.auto_filter.ref = f"A{hdr_row}:{get_column_letter(NCOLS)}{last_row}"


# ══════════════════════════════════════════════════════════════════════════════
# BUILD TABS
# ══════════════════════════════════════════════════════════════════════════════

ws_p = wb.create_sheet("Priority Matches")
build_data_tab(ws_p, rows_priority, "Priority Matches")

ws_w = wb.create_sheet("Watch List")
build_data_tab(ws_w, rows_watch, "Watch List", watch_note=True)

ws_a = wb.create_sheet("All Matches")
build_data_tab(ws_a, rows_all, "All Matches", stats=True)

wb.save(OUTPUT_FILE)
print(f"Saved. green={n_green} (h={n_high} s={n_solid} i={n_inv}) watch={n_watch} low={n_low}")
