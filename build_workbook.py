#!/usr/bin/env python3
"""
build_workbook.py  --  AncestryDNA Match Pipeline workbook builder
Version: 3.0
Date: 2026-06-14

Builds the 4-tab gold-standard workbook (Start Here | Priority Matches |
Watch List | All Matches) directly from a Genealogy Assistant CSV export
and an API results JSON file.  Formatting matches build_v2.2.py exactly.

Usage:
    python build_workbook.py \\
        --api-json   /tmp/Cynthia_Wilbur_api_results.json \\
        --input-csv  "/path/to/GA_export.csv" \\
        --kit-url    "https://www.ancestry.com/dna/matches/GUID.../list" \\
        --first-name Cynthia \\
        --last-name  Wilbur \\
        --batch      1 \\
        [--output-dir /path/to/output]

API JSON:  list of {guid, unweighted_cm, longest_segment, segments}
Output:    {FirstName}_{LastName}_{N}matches_{YYYYMMDD}_v{VERSION}.xlsx

Dependencies:  pip install openpyxl
"""

import argparse
import csv
import json
import os
import re
import sys
from collections import Counter
from datetime import date

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.cell.rich_text import TextBlock, CellRichText
    from openpyxl.cell.text import InlineFont
    _HAS_RICHTEXT = True
except ImportError:
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
        _HAS_RICHTEXT = False
    except ImportError:
        sys.exit("ERROR: openpyxl is required.  Run: pip install openpyxl")

SCRIPT_VERSION = "3.0"

# ─── PALETTE (matches build_v2.2.py exactly) ─────────────────────────────────
NAVY       = "1F3864"
DARK_GRN   = "375623"
CALLOUT_BG = "FFF2CC"
CALLOUT_FG = "7F6000"
LBLUE_TEXT = "BDD7EE"
BURGUNDY   = "7B2D3E"
BURG_LIGHT = "F5E6EA"
BURG_DARK  = "4A1020"
ACTION_COLOR = "9C6500"
ACTION_FILL  = "FFF5CC"

BLACK   = "000000"
WHITE   = "FFFFFF"
GREY_D  = "595959"
LINK_C  = "1155CC"
GREY_I  = "707070"    # italic greyed columns

F_DARK    = PatternFill("solid", fgColor="70AD47")
F_MED     = PatternFill("solid", fgColor="A9D18E")
F_LIGHT   = PatternFill("solid", fgColor="E2EFDA")
F_LOW     = PatternFill("solid", fgColor="FFD7D7")
F_WHITE   = PatternFill("solid", fgColor=WHITE)
F_NAVY    = PatternFill("solid", fgColor=NAVY)
F_CALLOUT = PatternFill("solid", fgColor=CALLOUT_BG)
F_BURG    = PatternFill("solid", fgColor=BURGUNDY)
F_BURGL   = PatternFill("solid", fgColor=BURG_LIGHT)
F_STATS   = PatternFill("solid", fgColor="F4F6FB")
F_ACTION  = PatternFill("solid", fgColor=ACTION_FILL)

TIER_FILL = {"high": F_DARK, "solid": F_MED, "investigate": F_LIGHT,
             "watch": F_LOW, "low": F_LOW, "nodata": F_WHITE}

QUAD_FILL = {
    "PP":       PatternFill("solid", fgColor="E2EFDA"),
    "PM":       PatternFill("solid", fgColor="BDD7EE"),
    "MP":       PatternFill("solid", fgColor="FFEB9C"),
    "MM":       PatternFill("solid", fgColor="FCE4D6"),
    "Multiple": PatternFill("solid", fgColor="E2CEEF"),
}

YT_LINK = "https://www.youtube.com/watch?v=0hGonlNLinM"

# ─── COLUMN SCHEMA ────────────────────────────────────────────────────────────
# (key, header, width, greyed)
COLS = [
    ("match_name",      "Match Name",               36, False),
    ("flag",            "⚠",                          7, False),
    ("longest",         "Longest Segment",           15, False),
    ("ascm",            "AScM",                       9, False),
    ("unweighted_cm",   "Unweighted cM",             14, False),
    ("segments",        "Segments",                  10, False),
    ("weighted_cm",     "Weighted cM",               13, True),
    ("family_tree",     "Family Tree",               20, False),
    ("tree_size",       "Tree Size",                 10, False),
    ("common_ancestor", "Common Ancestor",           22, False),
    ("line_assignment", "Family Line Assignment",    22, False),
    ("groups",          "Groups",                    22, False),
    ("notes",           "Notes",                     40, False),
    ("match_side",      "Match Side",                18, True),
    ("guid",            "GUID (Reference Anchor)",   38, True),
]
NCOLS   = len(COLS)
UWT_COL = get_column_letter(next(i+1 for i,(k,_,_,_) in enumerate(COLS) if k=="unweighted_cm"))
SEG_COL = get_column_letter(next(i+1 for i,(k,_,_,_) in enumerate(COLS) if k=="segments"))

GUID_RE = re.compile(
    r'compare/([0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12})/', re.I)
KIT_RE  = re.compile(
    r'/dna/matches/([0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12})/', re.I)

# ─── URL BUILDERS ─────────────────────────────────────────────────────────────
def url_profile(tester, match):
    return f"https://www.ancestry.com/dna/matches/{tester}/compare/{match}"

def url_thrulines(tester, match):
    return f"https://www.ancestry.com/discoveryui-matches/compare/{tester}/with/{match}"

# ─── HELPERS ──────────────────────────────────────────────────────────────────
def safe_int(val):
    try:
        v = int(float(str(val).strip()))
        return v if v >= 0 else None
    except (ValueError, TypeError):
        return None

def outer_box(ws, r1, r2, c1, c2, style='thin', color=NAVY):
    s, n = Side(style=style, color=color), Side(style=None)
    for ri in range(r1, r2+1):
        for ci in range(c1, c2+1):
            cell = ws.cell(row=ri, column=ci)
            cell.border = Border(
                top=s    if ri==r1 else n,
                bottom=s if ri==r2 else n,
                left=s   if ci==c1 else n,
                right=s  if ci==c2 else n,
            )

# ─── DATA LOADING ─────────────────────────────────────────────────────────────
def load_csv(csv_path):
    rows = []
    with open(csv_path, encoding='utf-8-sig', newline='') as f:
        reader = csv.DictReader(f)
        headers = reader.fieldnames or []
        url_col      = next((h for h in headers if 'url' in h.lower()), None)
        ca_col       = next((h for h in headers
                             if 'common' in h.lower() and 'ancestor' in h.lower()), None)
        tree_col     = next((h for h in headers
                             if 'tree' in h.lower() and 'size' not in h.lower()), None)
        tree_sz_col  = next((h for h in headers
                             if 'tree' in h.lower() and 'size' in h.lower()), None)
        wcm_col      = next((h for h in headers
                             if 'shared' in h.lower() and 'cm' in h.lower()), None) or \
                       next((h for h in headers if 'weighted' in h.lower()), None)
        for row in reader:
            raw_url = (row.get(url_col) or '') if url_col else ''
            m = GUID_RE.search(raw_url)
            guid = m.group(1).upper() if m else ''
            name = ''
            for k in ('Match Name', 'Name', 'match_name'):
                if row.get(k):
                    name = row[k].strip(); break
            if not name and headers:
                name = row.get(headers[0], '').strip()
            rows.append({
                'name':       name,
                'weighted':   safe_int(row.get(wcm_col) if wcm_col else None),
                'tree':       (row.get(tree_col) or '').strip() if tree_col else '',
                'tree_size':  safe_int(row.get(tree_sz_col) if tree_sz_col else None),
                'common_anc': (row.get(ca_col) or '').strip() if ca_col else '',
                'groups':     (row.get('Groups') or row.get('Group', '')).strip(),
                'notes':      (row.get('Notes') or '').strip(),
                'match_side': (row.get('Match Side') or row.get('Side', '')).strip(),
                'guid':       guid,
            })
    return rows

def load_api_json(json_path):
    with open(json_path) as f:
        raw = json.load(f)
    result = {}
    if isinstance(raw, list):
        for item in raw:
            g = str(item.get('guid', '')).upper()
            if g: result[g] = item
    elif isinstance(raw, dict):
        for g, item in raw.items():
            result[g.upper()] = item
    return result

def merge_data(csv_rows, api_data, tester_guid):
    merged, seen = [], set()
    for m in csv_rows:
        guid = m['guid']
        if not guid or guid in seen or guid == tester_guid:
            continue
        seen.add(guid)
        api = api_data.get(guid, {})
        lng  = safe_int(api.get('longest_segment')  or api.get('longestSharedSegment'))
        uw   = safe_int(api.get('unweighted_cm')    or api.get('totalSharedCentimorgans'))
        segs = safe_int(api.get('segments')         or api.get('numSharedSegments'))

        if lng is None:
            tier = 'nodata'
        else:
            passes_lng  = lng >= 20
            passes_ascm = (uw / segs >= 12) if segs and segs > 0 else False
            if passes_lng and passes_ascm:
                tier = 'high' if lng >= 50 else ('solid' if lng >= 30 else 'investigate')
            elif passes_lng or passes_ascm:
                tier = 'watch'
            else:
                tier = 'low'

        merged.append({
            **m,
            'longest':     lng,
            'unweighted':  uw,
            'segments':    segs,
            'tier':        tier,
            'flag':        '⚠️' if segs and 0 < segs <= 2 else '',
            'has_api':     bool(api),
        })
    merged.sort(key=lambda x: (x['longest'] or 0, x['unweighted'] or 0), reverse=True)
    return merged

# ─── LAYOUT HELPERS ───────────────────────────────────────────────────────────
def set_col_widths(ws):
    for i, (_, _, w, _) in enumerate(COLS, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

def write_col_headers(ws, hr):
    ws.row_dimensions[hr].height = 44.0
    for i, (_, label, _, _) in enumerate(COLS, 1):
        c = ws.cell(row=hr, column=i, value=label.upper())
        c.fill = F_NAVY
        c.font = Font(name="Calibri", bold=True, size=11, color=WHITE)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

def banner(ws, subtitle, row=1):
    """Navy banner (rows 1-2), returns next available row."""
    c = ws.cell(row=row, column=1, value=ws._tester_name)
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

# ─── DATA ROW WRITER ─────────────────────────────────────────────────────────
def write_data_rows(ws, rows, start_row, tester_guid):
    for rd in rows:
        fill = TIER_FILL[rd['tier']]
        dark = (rd['tier'] == 'high')
        guid = rd['guid']
        has_notes = bool(rd.get('notes') or '')
        ws.row_dimensions[start_row].height = 40.0 if has_notes else 20.0

        for ci, (key, _, _, greyed) in enumerate(COLS, 1):
            c = ws.cell(row=start_row, column=ci)
            c.fill = fill
            c.alignment = Alignment(vertical="center")

            gc   = Font(name="Calibri", size=12, color=GREY_I, italic=True)
            std  = Font(name="Calibri", size=12, color=BLACK, bold=dark)
            link = Font(name="Calibri", size=12,
                        color=(BLACK if dark else LINK_C), underline="single", bold=dark)

            if key == "match_name":
                c.value     = rd['name']
                c.alignment = Alignment(horizontal="right", vertical="center")
                if guid:
                    c.hyperlink = url_profile(tester_guid, guid); c.font = link
                else:
                    c.font = std

            elif key == "flag":
                c.value     = rd['flag']
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.font = (Font(name="Calibri", size=13, bold=True, color="FF0000")
                          if rd['flag'] else Font(name="Calibri", size=12, color=BLACK))

            elif key == "longest":
                c.value = rd['longest']; c.number_format = "0"
                c.alignment = Alignment(horizontal="center", vertical="center"); c.font = std

            elif key == "ascm":
                c.value = f"=IFERROR({UWT_COL}{start_row}/{SEG_COL}{start_row},\"\")"
                c.number_format = "0.0"
                c.alignment = Alignment(horizontal="center", vertical="center"); c.font = std

            elif key == "unweighted_cm":
                c.value = rd['unweighted']; c.number_format = "0"
                c.alignment = Alignment(horizontal="center", vertical="center"); c.font = std

            elif key == "segments":
                c.value = rd['segments']; c.number_format = "0"
                c.alignment = Alignment(horizontal="center", vertical="center"); c.font = std

            elif key == "weighted_cm":
                c.value = rd['weighted']; c.number_format = "0"
                c.alignment = Alignment(horizontal="center", vertical="center"); c.font = gc

            elif key == "family_tree":
                tree_text = rd.get('tree') or ''
                if tree_text and guid:
                    c.value = tree_text
                    c.hyperlink = url_thrulines(tester_guid, guid); c.font = link
                elif tree_text:
                    c.value = tree_text; c.font = std
                c.alignment = Alignment(horizontal="center", vertical="center")

            elif key == "tree_size":
                c.value = rd.get('tree_size')
                c.alignment = Alignment(horizontal="center", vertical="center"); c.font = std

            elif key == "common_ancestor":
                ca = rd.get('common_anc') or ''
                if ca and guid:
                    c.value = '\U0001f465'
                    c.hyperlink = url_thrulines(tester_guid, guid)
                    c.font = Font(name="Calibri", size=14,
                                  color=(BLACK if dark else LINK_C),
                                  underline="single", bold=dark)
                else:
                    c.value = ''; c.font = std
                c.alignment = Alignment(horizontal="center", vertical="center")

            elif key == "line_assignment":
                c.value = None
                c.alignment = Alignment(horizontal="center", vertical="center"); c.font = std

            elif key == "groups":
                c.value = rd.get('groups') or None; c.font = std

            elif key == "notes":
                c.value = rd.get('notes') or None; c.font = std
                c.alignment = Alignment(wrap_text=True, vertical="top")

            elif key == "match_side":
                c.value = rd.get('match_side') or None
                c.alignment = Alignment(horizontal="center", vertical="center"); c.font = gc

            elif key == "guid":
                c.value = guid or None; c.font = gc

        start_row += 1
    return start_row

# ─── DATA TAB BUILDER ────────────────────────────────────────────────────────
def build_data_tab(ws, rows, tab_label, retrieval_date, tester_guid,
                   watch_note=False, stats=False,
                   n_high=0, n_solid=0, n_inv=0, n_green=0, n_watch=0, n_low=0, n_total=0):
    ws._tester_name = ws.parent._tester_name
    set_col_widths(ws)
    next_r = banner(ws, f"{tab_label}  ·  Ancestry data retrieved {retrieval_date}")

    if stats:
        ws.row_dimensions[next_r].height = 6.0; next_r += 1
        stats_start = next_r
        for ci in range(1, NCOLS+1):
            ws.cell(row=next_r, column=ci).fill = F_STATS
        c = ws.cell(row=next_r, column=1, value="BATCH SUMMARY")
        c.fill = F_STATS
        c.font = Font(name="Calibri", bold=True, size=12, color=NAVY)
        c.alignment = Alignment(vertical="center", indent=1)
        ws.merge_cells(start_row=next_r, start_column=1, end_row=next_r, end_column=NCOLS)
        ws.row_dimensions[next_r].height = 26.0; next_r += 1
        for ci in range(1, NCOLS+1):
            ws.cell(row=next_r, column=ci).fill = F_WHITE
        ws.row_dimensions[next_r].height = 4.0; next_r += 1
        for sfill, label, val, fc in [
            (None,    "Total matches retrieved",          n_total, None),
            (None,    None,                               None,    None),
            (F_DARK,  "High priority  (50+ cM)",          n_high,  WHITE),
            (F_MED,   "Solid signal  (30-49 cM)",         n_solid, BLACK),
            (F_LIGHT, "Worth investigating  (20-29 cM)",  n_inv,   "375623"),
            (None,    "All priority matches",             n_green, NAVY),
            (None,    None,                               None,    None),
            (F_LOW,   "Watch List",                       n_watch, "833C00"),
            (None,    "Low priority",                     n_low,   NAVY),
        ]:
            ws.row_dimensions[next_r].height = 20.0
            if label is None:
                ws.row_dimensions[next_r].height = 5.0; next_r += 1; continue
            row_fill = sfill if sfill else F_WHITE
            for ci in range(1, NCOLS+1):
                ws.cell(row=next_r, column=ci).fill = row_fill
            ca = ws.cell(row=next_r, column=1, value=label)
            ca.fill = row_fill
            ca.font = Font(name="Calibri", size=12, bold=(sfill is not None), color=fc or NAVY)
            ca.alignment = Alignment(vertical="center", indent=2)
            ws.merge_cells(start_row=next_r, start_column=1, end_row=next_r, end_column=3)
            cv = ws.cell(row=next_r, column=5, value=val)
            cv.fill = row_fill
            cv.font = Font(name="Calibri", size=12, bold=(sfill is not None), color=fc or NAVY)
            cv.number_format = "#,##0"
            cv.alignment = Alignment(horizontal="right", vertical="center")
            next_r += 1
        outer_box(ws, stats_start, next_r-1, 1, NCOLS, style='thin', color=NAVY)
        ws.row_dimensions[next_r].height = 8.0; next_r += 1

    elif watch_note:
        c = ws.cell(row=next_r, column=1,
            value=("These matches came close but did not meet both research criteria. "
                   "A match lands here when it passes one condition — either a longest "
                   "segment of 20 cM or more, or an AScM of 12 or more — but not both. "
                   "Not a priority now, but worth keeping an eye on as trees grow."))
        c.fill = F_CALLOUT
        c.font = Font(name="Calibri", size=12, color=CALLOUT_FG)
        c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.merge_cells(start_row=next_r, start_column=1, end_row=next_r, end_column=NCOLS)
        ws.row_dimensions[next_r].height = 60.0; next_r += 1
        ws.row_dimensions[next_r].height = 5.0; next_r += 1

    hdr_row = next_r
    write_col_headers(ws, hdr_row); next_r += 1
    data_start = next_r
    last_row = write_data_rows(ws, rows, data_start, tester_guid) - 1
    ws.freeze_panes = ws.cell(row=data_start, column=1)
    if last_row >= data_start:
        ws.auto_filter.ref = f"A{hdr_row}:{get_column_letter(NCOLS)}{last_row}"

# ─── START HERE ───────────────────────────────────────────────────────────────
def build_start_here(wb, tester_name, retrieval_date, batch):
    ws_s = wb.create_sheet("Start Here")
    ws_s.column_dimensions["A"].width = 5
    ws_s.column_dimensions["B"].width = 18
    ws_s.column_dimensions["C"].width = 90
    ws_s.column_dimensions["D"].width = 87
    r = 1

    def sh_white_row():
        for ci in range(1, 5):
            ws_s.cell(row=r, column=ci).fill = F_WHITE

    def sh_navy(val, size=13, h=28, span=4):
        nonlocal r
        c = ws_s.cell(row=r, column=1, value=val)
        c.fill = F_NAVY; c.font = Font(name="Calibri", bold=True, size=size, color=WHITE)
        c.alignment = Alignment(vertical="center")
        if span > 1:
            ws_s.merge_cells(start_row=r, start_column=1, end_row=r, end_column=span)
        ws_s.row_dimensions[r].height = h; r += 1

    def sh_body(val, h=46, color=BLACK, fill=None, span=4):
        nonlocal r
        c = ws_s.cell(row=r, column=1, value=val)
        c.font = Font(name="Calibri", size=12, color=color)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        c.fill = fill or F_WHITE
        if span > 1:
            ws_s.merge_cells(start_row=r, start_column=1, end_row=r, end_column=span)
        ws_s.row_dimensions[r].height = h; r += 1
        return c

    def sh_blank(h=5):
        nonlocal r
        sh_white_row()
        ws_s.row_dimensions[r].height = h; r += 1

    def sh_lbl_body(lbl, txt, h=26, lbl_color=NAVY, txt_color=NAVY):
        nonlocal r
        ca = ws_s.cell(row=r, column=1, value=lbl)
        ca.fill = F_WHITE
        ca.font = Font(name="Calibri", bold=True, size=14, color=lbl_color)
        ca.alignment = Alignment(horizontal="center", vertical="center")
        ws_s.cell(row=r, column=2).fill = F_WHITE
        cc = ws_s.cell(row=r, column=3, value=txt)
        cc.fill = F_WHITE
        cc.font = Font(name="Calibri", size=12, color=txt_color)
        cc.alignment = Alignment(wrap_text=True, vertical="center")
        ws_s.cell(row=r, column=4).fill = F_WHITE
        ws_s.row_dimensions[r].height = h; r += 1

    def tier_row(fill_cc, fc, label, desc, h=26):
        nonlocal r
        ca = ws_s.cell(row=r, column=1, value=label)
        ca.fill = fill_cc; ca.font = Font(name="Calibri", bold=True, size=11, color=fc)
        ca.alignment = Alignment(horizontal="center", vertical="center")
        ws_s.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        cc = ws_s.cell(row=r, column=3, value=desc)
        cc.fill = F_WHITE; cc.font = Font(name="Calibri", size=12, color=NAVY)
        cc.alignment = Alignment(wrap_text=True, vertical="center")
        ws_s.cell(row=r, column=4).fill = F_WHITE
        ws_s.row_dimensions[r].height = h; r += 1

    # Banner
    c = ws_s.cell(row=r, column=1, value=tester_name)
    c.fill = F_NAVY; c.font = Font(name="Calibri", bold=True, size=20, color=WHITE)
    c.alignment = Alignment(vertical="center")
    ws_s.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    ws_s.row_dimensions[r].height = 40.0; r += 1
    c = ws_s.cell(row=r, column=1,
        value=f"DNA Match Analysis  ·  Batch {batch}  ·  Ancestry data retrieved {retrieval_date}")
    c.fill = F_NAVY; c.font = Font(name="Calibri", size=11, color=LBLUE_TEXT)
    c.alignment = Alignment(vertical="center")
    ws_s.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    ws_s.row_dimensions[r].height = 22.0; r += 1
    sh_blank(h=8)

    sh_body("When you open your AncestryDNA match list, you see hundreds of names with no clear way "
            "to know where to start. This workbook changes that. It pulls the segment data Ancestry "
            "keeps off the main screen, calculates a quality score called AScM, and uses those two "
            "numbers to sort your matches by what is actually worth your time. "
            "The Priority Matches tab is where you begin.", h=66)
    sh_body("Important: this workbook is a snapshot. It does not sync with Ancestry in either "
            "direction. Changes you make here stay here only. To record your research in both "
            "places, enter it in both — see Working Between This File and Ancestry below.",
            h=46, color=CALLOUT_FG, fill=F_CALLOUT)
    sh_blank(h=14)

    sh_navy("QUICK START")
    sh_lbl_body(1, "Open the Priority Matches tab.", h=26)
    sh_blank(h=5)
    sh_lbl_body(2, "Start at the top — the dark green rows are your highest-priority matches: "
                   "a longest segment of 50 cM or more.", h=26)
    sh_blank(h=5)
    sh_lbl_body(3, "Click a match name to open their Ancestry profile comparison page.", h=26)
    sh_blank(h=5)

    if _HAS_RICHTEXT:
        def _rt(parts):
            blocks = []
            for txt, bold, uline, caps in parts:
                display = txt.upper() if caps else txt
                ifont = InlineFont(rFont="Calibri", sz=12, color=NAVY,
                                   b=bold, u=("single" if uline else None))
                blocks.append(TextBlock(ifont, display))
            return CellRichText(*blocks)
        def _col(n): return (n, True, True, True)
        def _txt(t): return (t, False, False, False)

        def lbl_rich(num, parts, h=26):
            nonlocal r
            ca = ws_s.cell(row=r, column=1, value=num)
            ca.fill = F_WHITE
            ca.font = Font(name="Calibri", bold=True, size=14, color=NAVY)
            ca.alignment = Alignment(horizontal="center", vertical="center")
            ws_s.cell(row=r, column=2).fill = F_WHITE
            cc = ws_s.cell(row=r, column=3)
            cc.fill = F_WHITE; cc.value = _rt(parts)
            cc.alignment = Alignment(wrap_text=True, vertical="center")
            ws_s.cell(row=r, column=4).fill = F_WHITE
            ws_s.row_dimensions[r].height = h; r += 1

        lbl_rich(4, [_txt("Check the "), _col("Common Ancestor"),
                     _txt(" column. If you see the \U0001f465 icon, click it to open "
                          "Ancestry's ThruLines view.")], h=26)
        sh_blank(h=5)
        lbl_rich(5, [_txt("If you find a connection, write it in the "), _col("Notes"),
                     _txt(" column. Fill in "), _col("Family Line Assignment"),
                     _txt(" when you know which branch the match belongs to.")], h=26)
    else:
        sh_lbl_body(4, "Check the COMMON ANCESTOR column. If you see the \U0001f465 icon, "
                       "click it to open Ancestry's ThruLines view.", h=26)
        sh_blank(h=5)
        sh_lbl_body(5, "If you find a connection, write it in the NOTES column. Fill in "
                       "FAMILY LINE ASSIGNMENT when you know which branch the match belongs to.", h=26)

    sh_blank(h=5)
    sh_lbl_body(6, "Work through all dark green rows, then medium green, then light green.", h=26)
    sh_blank(h=5)
    sh_lbl_body(7, "If you see someone you already know anywhere in the list — even in the "
                   "pink rows — add a note with the relationship. Known relatives can appear "
                   "at any color tier.", h=46)
    sh_blank(h=5)
    c = ws_s.cell(row=r, column=1,
        value="The explanations below are here when you want to understand the why. "
              "You do not need to read them to get started.")
    c.fill = F_WHITE; c.font = Font(name="Calibri", size=12, color="808080")
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws_s.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    ws_s.row_dimensions[r].height = 25.0; r += 1
    sh_blank(h=14)

    sh_navy("A NOTE ON THE TWO CM COLUMNS: WEIGHTED VS. UNWEIGHTED")
    sh_body("Ancestry displays a Weighted cM figure processed by an algorithm called TIMBER. "
            "TIMBER was designed to reduce the weight of DNA segments that appear across many "
            "thousands of people from the same ancestral population. For Ashkenazi Jewish testers, "
            "TIMBER may over-correct — what it flags as a pile-up may still point to a genuine "
            "genealogical connection. This workbook uses Unweighted cM throughout. "
            "The Weighted cM column is greyed out and kept only for reference.", h=86)
    sh_blank(h=14)

    sh_navy("THE TWO NUMBERS THAT MATTER")
    c = ws_s.cell(row=r, column=1, value="LONGEST SEGMENT (CM)")
    c.fill = PatternFill("solid", fgColor=DARK_GRN)
    c.font = Font(name="Calibri", bold=True, size=12, color=WHITE)
    ws_s.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    ws_s.row_dimensions[r].height = 22.0; r += 1
    sh_body("A single long segment is the strongest signal of a recent shared ancestor. "
            "In Ashkenazi DNA research, a long segment is hard to fake. Short scattered segments "
            "are common background noise; a long one points to something specific. "
            "The minimum threshold for this workbook is 20 cM.", h=66)
    sh_blank(h=5)
    tier_row(F_DARK,  WHITE,     "50+ cM",      "High priority. Start here.")
    tier_row(F_MED,   BLACK,     "30-49 cM",    "Solid genealogical signal.")
    tier_row(F_LIGHT, "375623",  "20-29 cM",    "Worth investigating.")
    tier_row(F_LOW,   "833C00",  "Under 20 cM", "Does not meet the minimum criterion.")
    sh_blank(h=5)
    sh_body("Real example: two matches, both 40 cM total, both with a longest segment of 22 cM. "
            "Match 1 has 2 segments (AScM = 20): research priority. Match 2 has 10 segments "
            "(AScM = 4): almost certainly untraceable background sharing. AScM tells them apart.",
            h=46, color=CALLOUT_FG, fill=F_CALLOUT)
    sh_blank(h=14)

    c = ws_s.cell(row=r, column=1, value="ASCM  (UNWEIGHTED cM / SEGMENTS)")
    c.fill = PatternFill("solid", fgColor=DARK_GRN)
    c.font = Font(name="Calibri", bold=True, size=12, color=WHITE)
    ws_s.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    ws_s.row_dimensions[r].height = 22.0; r += 1
    sh_body("AScM is the average size of each shared segment (Unweighted cM divided by Segments). "
            "This workbook calculates it automatically. A higher AScM means the shared DNA is "
            "concentrated into fewer, larger segments — a stronger signal of a real genealogical "
            "connection. Threshold: 12 to pass.", h=80)
    sh_blank(h=5)
    tier_row(F_DARK,  WHITE,    "20+",       "Unusually strong segment quality.")
    tier_row(F_MED,   BLACK,    "12-20",     "Typical range for passing Ashkenazi matches.")
    tier_row(F_LIGHT, "375623", "12-20",     "Light green tier also passes at AScM 12+.")
    tier_row(F_LOW,   "833C00", "Under 12",  "Does not meet the AScM criterion.")
    sh_blank(h=14)

    sh_navy("COLOR CODING AT A GLANCE")
    for fill_cc, fc, label, desc in [
        (F_DARK,  WHITE,     "Longest segment 50+ cM",
         "High priority. Both filter conditions passed comfortably. Start your research here."),
        (F_MED,   BLACK,     "Longest segment 30-49 cM",
         "Solid genealogical signal. Passes the filter; investigate after dark green."),
        (F_LIGHT, "375623",  "Longest segment 20-29 cM",
         "Passes the filter. Worth investigating, especially if a tree or ThruLines hint is present."),
        (F_LOW,   "833C00",  "Watch List  ·  Low Priority",
         "Does not meet one or both research criteria. Kept for reference in the Watch List and All Matches tabs."),
    ]:
        ca = ws_s.cell(row=r, column=1, value=label)
        ca.fill = fill_cc; ca.font = Font(name="Calibri", bold=True, size=11, color=fc)
        ca.alignment = Alignment(horizontal="center", vertical="center")
        ws_s.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        cc = ws_s.cell(row=r, column=3, value=desc)
        cc.fill = F_WHITE; cc.font = Font(name="Calibri", size=12, color=NAVY)
        cc.alignment = Alignment(wrap_text=True, vertical="center")
        ws_s.cell(row=r, column=4).fill = F_WHITE
        ws_s.row_dimensions[r].height = 26.0; r += 1
    sh_blank(h=14)

    sh_navy("COLUMN GUIDE")
    col_guide = [
        ("70AD47", WHITE,       "A   Match Name",
         "The match's display name on Ancestry, linked to their profile comparison page.", False),
        ("FFE699", CALLOUT_FG,  "B   ⚠️  Flag",
         "⚠️ appears when a match shares only 1 or 2 DNA segments. A single long segment "
         "can sometimes be an artifact. Worth extra scrutiny before drawing conclusions.", False),
        ("70AD47", WHITE,       "C   Longest Segment",
         "The longest single DNA segment you share, in centiMorgans. Primary filter criterion. Not adjusted by TIMBER.", False),
        ("70AD47", WHITE,       "D   AScM",
         "Average centiMorgans per segment (Unweighted cM / Segments). Live formula. Threshold: 12 to pass.", False),
        ("A9D18E", BLACK,       "E   Unweighted cM",
         "Total shared DNA before TIMBER adjustment. The figure used in all calculations.", False),
        ("E2EFDA", BLACK,       "F   Segments",
         "Number of distinct DNA segments you share. Used in the AScM formula.", False),
        ("DCDCDC", GREY_D,      "G   Weighted cM",
         "Ancestry's TIMBER-adjusted figure. Greyed out; reference only. Not used in any calculation.", False),
        ("2E75B6", WHITE,       "H   Family Tree",
         "Linked to the Ancestry compare view showing both trees and any ThruLines suggestions.", False),
        ("9DC3E6", BLACK,       "I   Tree Size",
         "Number of people in the match's linked tree. A larger tree improves odds of finding a connection.", False),
        ("2E75B6", WHITE,       "J   Common Ancestor",
         "Shows the \U0001f465 icon when Ancestry has a ThruLines suggestion. Click to open ThruLines. Hypothesis only.", False),
        (ACTION_FILL, ACTION_COLOR, "K   Family Line Assignment",
         "Which branch of the family this match belongs to: PP, PM, MP, MM, or Multiple. Your column to fill in.", True),
        ("4472C4", WHITE,       "L   Groups",
         "Ancestry color-dot group tags assigned to this match.", False),
        ("D9E1F2", BLACK,       "M   Notes",
         "Notes from Ancestry's notes field, plus any you add directly in this column.", False),
        ("DCDCDC", GREY_D,      "N   Match Side",
         "Ancestry's paternal/maternal/unassigned call. Greyed out — unreliable for fully Ashkenazi testers.", False),
        ("F2F2F2", GREY_D,      "O   GUID (Reference Anchor)",
         "Ancestry's internal unique identifier. The backbone of all hyperlinks. You will never need to read it.", False),
    ]
    for hex_fill, fc, label, desc, action_req in col_guide:
        row_r = r
        ca = ws_s.cell(row=r, column=1, value=label)
        ca.fill = PatternFill("solid", fgColor=hex_fill)
        ca.font = Font(name="Calibri", bold=True, size=11, color=fc)
        ca.alignment = Alignment(horizontal="center", vertical="center")
        ws_s.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        cc = ws_s.cell(row=r, column=3, value=desc)
        cc.fill = F_WHITE; cc.font = Font(name="Calibri", size=12, color=NAVY)
        cc.alignment = Alignment(wrap_text=True, vertical="top")
        ws_s.cell(row=r, column=4).fill = F_WHITE
        ws_s.row_dimensions[r].height = 46.0 if len(desc) > 80 else 26.0
        r += 1
        if action_req:
            outer_box(ws_s, row_r, row_r, 1, 4, style='dashed', color=ACTION_COLOR)
    sh_blank(h=14)

    sh_navy("FAMILY LINE ASSIGNMENT: QUADRANT COLORS")
    action_row = r
    sh_body("Your work — As you identify which branch of the family a match belongs to, "
            "enter the code in the Family Line Assignment column (column K). "
            "This is the column you fill in yourself; nothing here is pre-populated.",
            h=52, color=ACTION_COLOR, fill=F_ACTION)
    outer_box(ws_s, action_row, action_row, 1, 4, style='dashed', color=ACTION_COLOR)
    sh_blank(h=5)
    sh_body("Cool colors (green, blue) = paternal lines. Warm colors (yellow, coral) = maternal lines.", h=26)
    sh_blank(h=5)
    for label, bg, fc, desc in [
        ("PP   Paternal-Paternal", "E2EFDA", "375623", "Your father's father's side"),
        ("PM   Paternal-Maternal", "BDD7EE", "1F3864", "Your father's mother's side"),
        ("MP   Maternal-Paternal", "FFEB9C", "7F6000", "Your mother's father's side"),
        ("MM   Maternal-Maternal", "FCE4D6", "833C00", "Your mother's mother's side"),
        ("Multiple",               "E2CEEF", "3D1152", "Spans more than one branch; investigate separately"),
    ]:
        ca = ws_s.cell(row=r, column=1, value=label)
        ca.fill = PatternFill("solid", fgColor=bg)
        ca.font = Font(name="Calibri", bold=True, size=11, color=fc)
        ca.alignment = Alignment(horizontal="center", vertical="center")
        ws_s.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        cc = ws_s.cell(row=r, column=3, value=desc)
        cc.fill = F_WHITE; cc.font = Font(name="Calibri", size=12, color=NAVY)
        cc.alignment = Alignment(vertical="center")
        ws_s.cell(row=r, column=4).fill = F_WHITE
        ws_s.row_dimensions[r].height = 22.0; r += 1
    sh_blank(h=14)

    sh_navy("WORKING BETWEEN THIS FILE AND ANCESTRY")
    sh_body("This workbook and Ancestry do not sync. They are separate tools that complement each other.\n\n"
            "Work here: analyze your matches, assign family lines, make sense of the numbers.\n\n"
            "Record there: when you have worked out a family line assignment, go into Ancestry and tag "
            "that match with the appropriate color group. When you add a note here, copy it to "
            "Ancestry's note field on that match's profile.\n\n"
            "The workbook is your scratch pad. Ancestry is your record of assignment.", h=105)
    sh_blank(h=14)

    sh_blank(h=6)
    c = ws_s.cell(row=r, column=1, value="▶ WATCH THIS PRESENTATION")
    c.hyperlink = YT_LINK; c.fill = F_BURG
    c.font = Font(name="Calibri", bold=True, size=24, color=WHITE)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws_s.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    ws_s.row_dimensions[r].height = 60.0; r += 1
    c = ws_s.cell(row=r, column=1,
        value='"Help! I Got My DNA Results and I’m Confused, Part 1.25"   —   Gil Bardige')
    c.hyperlink = YT_LINK; c.fill = F_BURGL
    c.font = Font(name="Calibri", bold=True, size=13, color=BURGUNDY)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws_s.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    ws_s.row_dimensions[r].height = 32.0; r += 1
    c = ws_s.cell(row=r, column=1,
        value="Gil Bardige teaches DNA analysis for Jewish genealogy researchers. "
              "His videos cover how to read your match list, why AScM matters, and "
              "how to work effectively under Ashkenazi endogamy. The methodology behind "
              "the filter and tier system in this workbook is built on his work.")
    c.fill = F_BURGL; c.font = Font(name="Calibri", size=12, color=BURG_DARK)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws_s.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    ws_s.row_dimensions[r].height = 52.0; r += 1
    sh_blank(h=14)

    sh_navy("A FEW THINGS WORTH KNOWING")
    for txt, h in [
        ("You may see people you know in the pink rows. The tiers are a triage tool for unknown "
         "matches, not a judgment about every match in your list. If you recognize someone, add a note.", 46),
        ("DNA evidence never stands alone. A match in a color tier tells you where to look, "
         "not what you have found. Always correlate with documentary evidence.", 46),
        ("ThruLines and Common Ancestor suggestions are hypotheses, not conclusions. "
         "Ancestry builds them from trees, which contain errors. Treat them as research leads.", 46),
        ("Ancestry's relationship estimates are not reliable for Ashkenazi testers. "
         "Work from the numbers in this workbook instead.", 26),
        ("The pink rows are not a dead end. They may become useful as trees grow.", 26),
        ("Match Side is greyed out because it is unreliable for fully Ashkenazi testers.", 26),
    ]:
        ca = ws_s.cell(row=r, column=1, value="•")
        ca.fill = F_WHITE; ca.font = Font(name="Calibri", size=14, color=NAVY)
        ca.alignment = Alignment(horizontal="center", vertical="top")
        ws_s.cell(row=r, column=2).fill = F_WHITE
        cc = ws_s.cell(row=r, column=3, value=txt)
        cc.fill = F_WHITE; cc.font = Font(name="Calibri", size=12, color=NAVY)
        cc.alignment = Alignment(wrap_text=True, vertical="top")
        ws_s.cell(row=r, column=4).fill = F_WHITE
        ws_s.row_dimensions[r].height = h; r += 1
        sh_blank(h=5)

    ws_s.freeze_panes = "A3"
    return ws_s

# ─── MAIN ─────────────────────────────────────────────────────────────────────
def main():
    p = argparse.ArgumentParser(
        description="Build AncestryDNA match workbook (4-tab, build_v2.2.py format).")
    p.add_argument("--api-json",       required=True,
                   help="Path to API results JSON (list of {guid, unweighted_cm, longest_segment, segments})")
    p.add_argument("--input-csv",      required=True,
                   help="Path to Genealogy Assistant CSV export")
    p.add_argument("--kit-url",        required=True,
                   help="Ancestry kit URL (tester GUID derived from it)")
    p.add_argument("--first-name",     required=True, help="Tester first name")
    p.add_argument("--last-name",      required=True, help="Tester last name")
    p.add_argument("--batch",          default=1, type=int, help="Batch number (default: 1)")
    p.add_argument("--output-dir",     default=None,
                   help="Output directory (default: same as input CSV)")
    args = p.parse_args()

    if not os.path.isfile(args.api_json):
        sys.exit(f"ERROR: API JSON not found: {args.api_json}")
    if not os.path.isfile(args.input_csv):
        sys.exit(f"ERROR: CSV not found: {args.input_csv}")
    if args.input_csv.lower().endswith(".xlsx"):
        sys.exit("ERROR: --input-csv must be a .csv file, not .xlsx")

    m = KIT_RE.search(args.kit_url)
    if not m:
        sys.exit("ERROR: could not find a GUID in --kit-url")
    tester_guid = m.group(1).upper()

    today           = date.today()
    retrieval_date  = f"{today.strftime('%B')} {today.day}, {today.year}"
    date_compact    = today.strftime("%Y%m%d")
    tester_name     = f"{args.first_name} {args.last_name}"

    print(f"Tester:       {tester_name}")
    print(f"Tester GUID:  {tester_guid}")
    print(f"Batch:        {args.batch}")
    print(f"Date:         {retrieval_date}")

    print("\nLoading CSV...")
    csv_rows = load_csv(args.input_csv)
    print(f"  {len(csv_rows)} rows")

    print("Loading API results...")
    api_data = load_api_json(args.api_json)
    print(f"  {len(api_data)} records")

    print("Merging and sorting...")
    matches = merge_data(csv_rows, api_data, tester_guid)

    cnt = Counter(m["tier"] for m in matches)
    n_high  = cnt["high"]
    n_solid = cnt["solid"]
    n_inv   = cnt["investigate"]
    n_green = n_high + n_solid + n_inv
    n_watch = cnt["watch"]
    n_low   = cnt["low"]
    n_total = len(matches)

    print(f"\n  High priority  (50+ cM):   {n_high}")
    print(f"  Solid signal   (30-49 cM): {n_solid}")
    print(f"  Investigate    (20-29 cM): {n_inv}")
    print(f"  Watch List:               {n_watch}")
    print(f"  Low priority:             {n_low}")
    print(f"  Total:                    {n_total}")

    no_data = [m["guid"] for m in matches if not m["has_api"]]
    if no_data:
        print(f"\n  WARNING: {len(no_data)} matches have no API data")
        for g in no_data[:10]:
            print(f"    {g}")
        if len(no_data) > 10:
            print(f"    ... and {len(no_data)-10} more")

    rows_priority = [rd for rd in matches if rd["tier"] in ("high","solid","investigate")]
    rows_watch    = [rd for rd in matches if rd["tier"] == "watch"]

    print("\nBuilding workbook...")
    wb = Workbook()
    wb.remove(wb.active)
    wb._tester_name = tester_name

    build_start_here(wb, tester_name, retrieval_date, args.batch)

    ws_p = wb.create_sheet("Priority Matches")
    ws_p._tester_name = tester_name
    build_data_tab(ws_p, rows_priority, "Priority Matches", retrieval_date, tester_guid)

    ws_w = wb.create_sheet("Watch List")
    ws_w._tester_name = tester_name
    build_data_tab(ws_w, rows_watch, "Watch List", retrieval_date, tester_guid, watch_note=True)

    ws_a = wb.create_sheet("All Matches")
    ws_a._tester_name = tester_name
    build_data_tab(ws_a, matches, "All Matches", retrieval_date, tester_guid, stats=True,
                   n_high=n_high, n_solid=n_solid, n_inv=n_inv, n_green=n_green,
                   n_watch=n_watch, n_low=n_low, n_total=n_total)

    out_dir = args.output_dir or os.path.dirname(os.path.abspath(args.input_csv))
    os.makedirs(out_dir, exist_ok=True)
    filename   = f"{args.first_name}_{args.last_name}_{n_total}matches_{date_compact}_v{SCRIPT_VERSION}.xlsx"
    output_path = os.path.join(out_dir, filename)
    wb.save(output_path)

    print(f"\nSaved: {output_path}")
    print(f"Tiers: Priority {n_green} (H{n_high}/S{n_solid}/I{n_inv}) | Watch {n_watch} | Low {n_low}")

if __name__ == "__main__":
    main()
