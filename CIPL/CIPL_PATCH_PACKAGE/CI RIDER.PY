\
# CI RIDER.PY
# Patched: remove merges, match original row/col geometry, header grey fill, minimal borders, and totals band.
# Optimized: Uses excel_helpers for cached styles and faster border operations

from __future__ import annotations

import sys
from pathlib import Path
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# Add root directory to path for excel_helpers
ROOT_DIR = Path(__file__).resolve().parents[2]  # CONVERT directory
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from excel_helpers import (
    get_font,
    get_alignment,
    apply_border_outline_fast,
    apply_inner_grid_fast,
    resolve_merged_addr,
)


BLUE_THICK = Side(style="thick", color="000000")
GRAY_THIN = Side(style="thin", color="808080")
BLACK_THIN = Side(style="thin", color="000000")
BLACK_THICK = Side(style="thick", color="000000")

FILL_GRAY = PatternFill("solid", fgColor="F2F2F2")

FONT_TITLE = Font(name="Tahoma", size=15, bold=True)
FONT_SECTION = Font(name="Tahoma", size=9, bold=True)
FONT_B = Font(name="Tahoma", size=9, bold=True)
FONT_N = Font(name="Tahoma", size=9, bold=False)

ALIGN_L = Alignment(horizontal="left", vertical="center")
ALIGN_C = Alignment(horizontal="center", vertical="center")
ALIGN_R = Alignment(horizontal="right", vertical="center")
ALIGN_C_WRAP = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_L_WRAP_TOP = Alignment(horizontal="left", vertical="top", wrap_text=True)


def _set_border(cell, *, left=None, right=None, top=None, bottom=None):
    b = cell.border if cell.border else Border()
    cell.border = Border(
        left=left if left is not None else b.left,
        right=right if right is not None else b.right,
        top=top if top is not None else b.top,
        bottom=bottom if bottom is not None else b.bottom,
        diagonal=b.diagonal,
        diagonal_direction=b.diagonal_direction,
        outline=b.outline,
        vertical=b.vertical,
        horizontal=b.horizontal,
    )


def _outline(ws, r1, c1, r2, c2, side):
    """Optimized: Uses apply_border_outline_fast for edge-only processing."""
    apply_border_outline_fast(ws, r1, c1, r2, c2, side)


def _fill(ws, r1, c1, r2, c2, fill):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(r, c).fill = fill


def _center_across(ws, row, c1, c2, text, *, font=None):
    """Optimized center across using cached alignment."""
    align_center_continuous = get_alignment(horizontal="centerContinuous", vertical="center")
    for c in range(c1, c2 + 1):
        cell = ws.cell(row=row, column=c)
        cell.alignment = align_center_continuous
        if font is not None:
            cell.font = font
        if c == c1:
            cell.value = text
        else:
            cell.value = None


def apply_page_setup(ws):
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = 9  # A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_view.showGridLines = False


def _parse_pkg_count(items):
    # Try to parse from "OF 5"
    for it in items:
        of_txt = str(it.get("of_text", "")).strip().upper()
        if of_txt.startswith("OF"):
            parts = of_txt.replace("OF", "").strip().split()
            if parts:
                try:
                    return int(parts[0])
                except Exception:
                    pass
    return len(items) if items else 0


def populate_sheet(ws, data: dict):
    # Columns A..L
    COL_WIDTHS = [7.5, 7.5, 7.5, 37.0, 1.0, 1.0, 13.0, 9.0, 9.0, 9.0, 9.0, 9.0]
    for i, w in enumerate(COL_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Rows (dynamic based on item count)
    items = data.get("items", [])
    item_start = 8
    item_rows = max(len(items), 1)
    item_end = item_start + item_rows - 1
    totals_start = item_end + 1
    totals_end = totals_start + 2
    bottom_row = totals_end + 1

    # Fixed row heights (match screenshot)
    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 30
    ws.row_dimensions[3].height = 18
    ws.row_dimensions[4].height = 50
    ws.row_dimensions[5].height = 7.5
    ws.row_dimensions[6].height = 28
    ws.row_dimensions[7].height = 7.5
    for r in range(item_start, item_end + 1):
        ws.row_dimensions[r].height = 95
    for r in range(totals_start, totals_end + 1):
        ws.row_dimensions[r].height = 15
    ws.row_dimensions[bottom_row].height = 24

    ws.print_area = f"A1:L{bottom_row}"

    # Outer border
    _outline(ws, 1, 1, bottom_row, 12, GRAY_THIN)

    # Page no
    ws.cell(1, 12).value = data.get("meta", {}).get("page_no", "")
    ws.cell(1, 12).font = FONT_N
    ws.cell(1, 12).alignment = ALIGN_R

    # Title
    _center_across(ws, 2, 1, 12, data.get("title", "COMMERCIAL INVOICE ATTACHED RIDER"), font=FONT_TITLE)

    # Header row (grey)

    # Header group boxes (no merges)
    # A-B Marks
    _outline(ws, 4, 1, 4, 2, BLACK_THIN)
    _center_across(ws, 4, 1, 2, "Marks & No. of Pkgs", font=FONT_B)

    # C Item No
    _outline(ws, 4, 3, 4, 3, BLACK_THIN)
    ws.cell(4, 3).value = "Item No."
    ws.cell(4, 3).font = FONT_B
    ws.cell(4, 3).alignment = ALIGN_C_WRAP

    # D-F DESCRIPTION
    _outline(ws, 4, 4, 4, 6, BLACK_THIN)
    _center_across(ws, 4, 4, 6, "DESCRIPTION", font=FONT_B)

    # G HS CODE
    _outline(ws, 4, 7, 4, 7, BLACK_THIN)
    ws.cell(4, 7).value = "HS CODE"
    ws.cell(4, 7).font = FONT_B
    ws.cell(4, 7).alignment = ALIGN_C_WRAP

    # H ORIGIN
    _outline(ws, 4, 8, 4, 8, BLACK_THIN)
    ws.cell(4, 8).value = "ORIGIN"
    ws.cell(4, 8).font = FONT_B
    ws.cell(4, 8).alignment = ALIGN_C_WRAP

    # I Q'TY
    _outline(ws, 4, 9, 4, 9, BLACK_THIN)
    ws.cell(4, 9).value = "Q'TY"
    ws.cell(4, 9).font = FONT_B
    ws.cell(4, 9).alignment = ALIGN_C_WRAP

    # J UNIT
    _outline(ws, 4, 10, 4, 10, BLACK_THIN)
    ws.cell(4, 10).value = "UNIT"
    ws.cell(4, 10).font = FONT_B
    ws.cell(4, 10).alignment = ALIGN_C_WRAP

    # K Unit Price
    _outline(ws, 4, 11, 4, 11, BLACK_THIN)
    ws.cell(4, 11).value = "Unit Price\n(USD)"
    ws.cell(4, 11).font = FONT_B
    ws.cell(4, 11).alignment = ALIGN_C_WRAP

    # L Total Price
    _outline(ws, 4, 12, 4, 12, BLACK_THIN)
    ws.cell(4, 12).value = "Total Price\n(USD)"
    ws.cell(4, 12).font = FONT_B
    ws.cell(4, 12).alignment = ALIGN_C_WRAP

    # Section title
    _center_across(ws, 6, 1, 12, data.get("section_title", ""), font=FONT_SECTION)

    # Items
    if not items:
        items = [{
            "marks_no": 1,
            "of_text": "OF 1",
            "item_no": "",
            "description": "",
            "hs_code": "",
            "origin": "",
            "qty": "",
            "unit": "",
            "unit_price": "",
            "total_price": "",
        }]

    for idx, it in enumerate(items):
        r = item_start + idx
        ws.cell(r, 1).value = it.get("marks_no", "")
        ws.cell(r, 1).font = FONT_N
        ws.cell(r, 1).alignment = ALIGN_C

        ws.cell(r, 2).value = it.get("of_text", "")
        ws.cell(r, 2).font = FONT_N
        ws.cell(r, 2).alignment = ALIGN_C

        ws.cell(r, 3).value = it.get("item_no", "")
        ws.cell(r, 3).font = FONT_N
        ws.cell(r, 3).alignment = ALIGN_C

        ws.cell(r, 4).value = it.get("description", "")
        ws.cell(r, 4).font = FONT_N
        ws.cell(r, 4).alignment = ALIGN_L_WRAP_TOP

        ws.cell(r, 7).value = it.get("hs_code", "")
        ws.cell(r, 7).font = FONT_N
        ws.cell(r, 7).alignment = ALIGN_C

        ws.cell(r, 8).value = it.get("origin", "")
        ws.cell(r, 8).font = FONT_N
        ws.cell(r, 8).alignment = ALIGN_C

        ws.cell(r, 9).value = it.get("qty", "")
        ws.cell(r, 9).font = FONT_N
        ws.cell(r, 9).alignment = ALIGN_R

        ws.cell(r, 10).value = it.get("unit", "")
        ws.cell(r, 10).font = FONT_N
        ws.cell(r, 10).alignment = ALIGN_C

        ws.cell(r, 11).value = it.get("unit_price", "")
        ws.cell(r, 11).font = FONT_N
        ws.cell(r, 11).alignment = ALIGN_R

        ws.cell(r, 12).value = it.get("total_price", "")
        ws.cell(r, 12).font = FONT_N
        ws.cell(r, 12).alignment = ALIGN_R

    # Totals band (grey)

    totals = data.get("totals", {})
    labels = ["TOTAL FOB CHARGE", "FREIGHT", "INSURANCE"]
    values = [totals.get("total_fob", ""), totals.get("freight", ""), totals.get("insurance", "")]
    for i in range(3):
        rr = totals_start + i
        ws.cell(rr, 8).value = labels[i]
        ws.cell(rr, 8).font = FONT_N
        ws.cell(rr, 8).alignment = ALIGN_L
        ws.cell(rr, 12).value = values[i]
        ws.cell(rr, 12).font = FONT_N
        ws.cell(rr, 12).alignment = ALIGN_R

    # Bottom summary row
    _set_border(ws.cell(bottom_row, 1), top=BLACK_THICK)
    for c in range(2, 13):
        _set_border(ws.cell(bottom_row, c), top=BLACK_THICK)

    pkg_count = _parse_pkg_count(items)
    ws.cell(bottom_row, 1).value = "TOTAL"
    ws.cell(bottom_row, 1).font = FONT_B
    ws.cell(bottom_row, 1).alignment = ALIGN_L

    ws.cell(bottom_row, 2).value = pkg_count if pkg_count else ""
    ws.cell(bottom_row, 2).font = FONT_N
    ws.cell(bottom_row, 2).alignment = ALIGN_C

    ws.cell(bottom_row, 4).value = data.get("bottom_terms", "")
    ws.cell(bottom_row, 4).font = FONT_B
    ws.cell(bottom_row, 4).alignment = ALIGN_L

    # Total qty, unit, amount
    ws.cell(bottom_row, 9).value = data.get("total_qty", "")
    ws.cell(bottom_row, 9).font = FONT_N
    ws.cell(bottom_row, 9).alignment = ALIGN_R

    ws.cell(bottom_row, 10).value = data.get("bottom_unit", "")
    ws.cell(bottom_row, 10).font = FONT_N
    ws.cell(bottom_row, 10).alignment = ALIGN_C

    ws.cell(bottom_row, 12).value = totals.get("total_fob", "")
    ws.cell(bottom_row, 12).font = FONT_N
    ws.cell(bottom_row, 12).alignment = ALIGN_R


# Backward-compatible API expected by make_cipl_set.py
def set_col_widths(ws):
    widths = [7.5, 7.5, 7.5, 37.0, 1.0, 1.0, 13.0, 9.0, 9.0, 9.0, 9.0, 9.0]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

def build_sheet(ws, data: dict):
    populate_sheet(ws, data)
