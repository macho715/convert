\
# PACKING LIST ATTACHED RIDER.PY
# Patched: remove merges, match original row/col geometry, minimal borders, and correct total row.
# Optimized: Uses excel_helpers for cached styles and faster border operations

from __future__ import annotations

import sys
from pathlib import Path
from openpyxl.styles import Font, Alignment, Border, Side
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


def set_page_setup(ws):
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = 9  # A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_view.showGridLines = False


def _parse_pkg_count(items):
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
    # Columns A..R
    COL_WIDTHS = [7.5, 7.5, 7.5, 7.5, 7.5, 51.0, 2.0, 13.0, 9.0, 9.0, 9.0, 9.0, 5.5, 1.0, 5.5, 1.0, 5.5, 9.0]
    for i, w in enumerate(COL_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    items = data.get("items", [])
    item_start = 8
    item_rows = max(len(items), 1)
    item_end = item_start + item_rows - 1
    blank_row = item_end + 1
    total_row = blank_row + 1

    # Row heights (match screenshot)
    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 30
    ws.row_dimensions[3].height = 18
    ws.row_dimensions[4].height = 50
    ws.row_dimensions[5].height = 7.5
    ws.row_dimensions[6].height = 28
    ws.row_dimensions[7].height = 7.5
    for r in range(item_start, item_end + 1):
        ws.row_dimensions[r].height = 95
    ws.row_dimensions[blank_row].height = 15
    ws.row_dimensions[total_row].height = 24

    ws.print_area = f"A1:R{total_row}"

    # Outer border
    _outline(ws, 1, 1, total_row, 18, GRAY_THIN)

    # Page no
    ws.cell(1, 18).value = data.get("meta", {}).get("page_no", "")
    ws.cell(1, 18).font = FONT_N
    ws.cell(1, 18).alignment = ALIGN_R

    # Title
    _center_across(ws, 2, 1, 18, data.get("title", "PACKING LIST ATTACHED RIDER"), font=FONT_TITLE)

    # Header row (row4) — boxed, no fill
    # Group outlines
    def box(c1, c2, text):
        _outline(ws, 4, c1, 4, c2, BLACK_THIN)
        _center_across(ws, 4, c1, c2, text, font=FONT_B)

    box(1, 2, "Marks & No. of Pkgs")
    box(3, 3, "Item No.")
    box(4, 4, "Item\nLocation")
    box(5, 5, "Packing\nStyle")
    box(6, 7, "Description of Goods")
    box(8, 8, "HS CODE")
    box(9, 9, "Q'ty")
    box(10, 10, "Unit")
    box(11, 11, "Net\nWeight\n(kgs)")
    box(12, 12, "Gross\nWeight\n(kgs)")
    box(13, 17, "Dimension\nL x W x H (CM)")
    box(18, 18, "Volume\n(CBM)")

    # Section title
    _center_across(ws, 6, 1, 18, data.get("section_title", ""), font=FONT_SECTION)

    # Items
    if not items:
        items = [{
            "marks_no": 1,
            "of_text": "OF 1",
            "item_no": "",
            "item_location": "",
            "packing_style": "",
            "description": "",
            "hs_code": "",
            "qty": "",
            "unit": "",
            "net_weight": "",
            "gross_weight": "",
            "dim_l": "",
            "dim_w": "",
            "dim_h": "",
            "volume_cbm": "",
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

        ws.cell(r, 4).value = it.get("item_location", "")
        ws.cell(r, 4).font = FONT_N
        ws.cell(r, 4).alignment = ALIGN_C

        ws.cell(r, 5).value = it.get("packing_style", "")
        ws.cell(r, 5).font = FONT_N
        ws.cell(r, 5).alignment = ALIGN_C

        ws.cell(r, 6).value = it.get("description", "")
        ws.cell(r, 6).font = FONT_N
        ws.cell(r, 6).alignment = ALIGN_L_WRAP_TOP
        # spacer col G left empty intentionally

        ws.cell(r, 8).value = it.get("hs_code", "")
        ws.cell(r, 8).font = FONT_N
        ws.cell(r, 8).alignment = ALIGN_C

        ws.cell(r, 9).value = it.get("qty", "")
        ws.cell(r, 9).font = FONT_N
        ws.cell(r, 9).alignment = ALIGN_R

        ws.cell(r, 10).value = it.get("unit", "")
        ws.cell(r, 10).font = FONT_N
        ws.cell(r, 10).alignment = ALIGN_C

        ws.cell(r, 11).value = it.get("net_weight", "")
        ws.cell(r, 11).font = FONT_N
        ws.cell(r, 11).alignment = ALIGN_R

        ws.cell(r, 12).value = it.get("gross_weight", "")
        ws.cell(r, 12).font = FONT_N
        ws.cell(r, 12).alignment = ALIGN_R

        # Dimensions across M-Q
        ws.cell(r, 13).value = it.get("dim_l", "")
        ws.cell(r, 13).font = FONT_N
        ws.cell(r, 13).alignment = ALIGN_R
        ws.cell(r, 14).value = "x"
        ws.cell(r, 14).font = FONT_N
        ws.cell(r, 14).alignment = ALIGN_C
        ws.cell(r, 15).value = it.get("dim_w", "")
        ws.cell(r, 15).font = FONT_N
        ws.cell(r, 15).alignment = ALIGN_R
        ws.cell(r, 16).value = "x"
        ws.cell(r, 16).font = FONT_N
        ws.cell(r, 16).alignment = ALIGN_C
        ws.cell(r, 17).value = it.get("dim_h", "")
        ws.cell(r, 17).font = FONT_N
        ws.cell(r, 17).alignment = ALIGN_R

        ws.cell(r, 18).value = it.get("volume_cbm", "")
        ws.cell(r, 18).font = FONT_N
        ws.cell(r, 18).alignment = ALIGN_R

    # Total row (thick top border)
    for c in range(1, 19):
        _set_border(ws.cell(total_row, c), top=BLACK_THICK)

    pkg_count = _parse_pkg_count(items)

    ws.cell(total_row, 1).value = "TOTAL"
    ws.cell(total_row, 1).font = FONT_B
    ws.cell(total_row, 1).alignment = ALIGN_L

    ws.cell(total_row, 2).value = pkg_count if pkg_count else ""
    ws.cell(total_row, 2).font = FONT_N
    ws.cell(total_row, 2).alignment = ALIGN_C

    ws.cell(total_row, 9).value = data.get("total_qty", "")
    ws.cell(total_row, 9).font = FONT_N
    ws.cell(total_row, 9).alignment = ALIGN_R

    ws.cell(total_row, 10).value = data.get("bottom_unit", "") or (items[0].get("unit") if items else "")
    ws.cell(total_row, 10).font = FONT_N
    ws.cell(total_row, 10).alignment = ALIGN_C

    ws.cell(total_row, 11).value = data.get("total_net", "")
    ws.cell(total_row, 11).font = FONT_N
    ws.cell(total_row, 11).alignment = ALIGN_R

    ws.cell(total_row, 12).value = data.get("total_gross", "")
    ws.cell(total_row, 12).font = FONT_N
    ws.cell(total_row, 12).alignment = ALIGN_R

    ws.cell(total_row, 18).value = data.get("total_volume", "")
    ws.cell(total_row, 18).font = FONT_N
    ws.cell(total_row, 18).alignment = ALIGN_R


# Backward-compatible API expected by make_cipl_set.py
def build_rider(ws, data: dict):
    populate_sheet(ws, data)
