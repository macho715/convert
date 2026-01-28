#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PACKING LIST ATTACHED RIDER (Page 2/2) — Fixed Layout Excel Generator

Goal
- 항차가 바뀌어도 "기본 정보 + 라인아이템"만 바꾸면 동일 포맷(폭/행높이/헤더/테두리)의 Rider(2/2) Excel 생성

Dependencies
  pip install openpyxl

Run
  python packing_list_rider_p2.py --out PACKING_LIST_RIDER_P2.xlsx

Notes (per your spec)
- Columns A:R widths (exact)
- Row heights sequence (18,30,18,50,7.5,28,7.5,95*items,15,24)
- Title centered: PACKING LIST ATTACHED RIDER
- Page No top-right: PAGE NO.: 2 OF 2
- Header columns (no merge in table body)
- Blue outer frame + black table borders/grid
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
from typing import List, Dict, Any, Optional

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins


# =============================================================================
# 0) TEMPLATE FIXED SETTINGS (from user)
# =============================================================================

# Column widths A:R (exact)
COL_WIDTHS = {
    "A": 7.5,
    "B": 7.5,
    "C": 7.5,
    "D": 7.5,
    "E": 7.5,
    "F": 51.0,
    "G": 2.0,   # spacer between Description and HS
    "H": 13.0,  # HS CODE
    "I": 9.0,   # Q'ty
    "J": 9.0,   # Unit
    "K": 9.0,   # Net Weight
    "L": 9.0,   # Gross Weight
    "M": 5.5,   # Dimension L
    "N": 1.0,   # 'x'
    "O": 5.5,   # Dimension W
    "P": 1.0,   # 'x'
    "Q": 5.5,   # Dimension H
    "R": 9.0,   # Volume
}

COLS = list("ABCDEFGHIJKLMNOPQR")  # A..R

# Layout rows (fixed positions in this template)
ROW_PAGE_NO = 1     # height 18
ROW_TITLE   = 2     # height 30
ROW_GAP_1   = 3     # height 18
ROW_HEADER  = 4     # height 50
ROW_GAP_2   = 5     # height 7.5
ROW_SECTION = 6     # height 28
ROW_GAP_3   = 7     # height 7.5
ROW_ITEMS_START = 8 # each item row height 95
# after items:
# ROW_GAP_4 = items_end + 1 (height 15)
# ROW_TOTAL = items_end + 2 (height 24)

ROW_HEIGHTS_FIXED = {
    ROW_PAGE_NO: 18.0,
    ROW_TITLE: 30.0,
    ROW_GAP_1: 18.0,
    ROW_HEADER: 50.0,
    ROW_GAP_2: 7.5,
    ROW_SECTION: 28.0,
    ROW_GAP_3: 7.5,
}

ITEM_ROW_HEIGHT = 95.0
GAP_AFTER_ITEMS_HEIGHT = 15.0
TOTAL_ROW_HEIGHT = 24.0

MIN_COL = 1
MAX_COL = 18  # A..R


# =============================================================================
# 1) DATA MODEL (항차별 입력)
# =============================================================================
@dataclass
class RiderItem:
    marks_no: int
    of_text: str
    item_no: str
    item_location: str
    packing_style: str
    description: str
    hs_code: str
    qty: float
    unit: str
    net_kg: float
    gross_kg: float
    dim_l_cm: float
    dim_w_cm: float
    dim_h_cm: float
    volume_cbm: float


DEFAULT_DATA: Dict[str, Any] = {
    "page_no": "PAGE NO.: 2 OF 2",
    "title": "PACKING LIST ATTACHED RIDER",
    "section_title": "STEEL CONDUIT (14TH)",
    "items": [
        {
            "marks_no": 1, "of_text": "OF 5", "item_no": "4.1", "item_location": "AGI", "packing_style": "SKID",
            "description": (
                'RIGID STEEL CONDUIT  W/ ONE(1) COUPLING\n'
                '- HOT DIPPED GALVANIZED ( H.D.G. ) 65㎛ as per BS729\n'
                '- STANDARD LENGTH : 3.048M\n'
                '- PIPE THREADED : "NPT" TYPE\n'
                '- IEC 60423, NEMA FB.I, OR EQUIVALENT\n'
                '28mm (1")\n'
                '(1 EA = 3 M, 408 EA)'
            ),
            "hs_code": "7303.00.1090", "qty": 1224, "unit": "M",
            "net_kg": 2800.00, "gross_kg": 3000.00,
            "dim_l_cm": 330, "dim_w_cm": 110, "dim_h_cm": 105,
            "volume_cbm": 3.812
        },
        {
            "marks_no": 2, "of_text": "OF 5", "item_no": "4.1", "item_location": "AGI", "packing_style": "SKID",
            "description": (
                'RIGID STEEL CONDUIT  W/ ONE(1) COUPLING\n'
                '- HOT DIPPED GALVANIZED ( H.D.G. ) 65㎛ as per BS729\n'
                '- STANDARD LENGTH : 3.048M\n'
                '- PIPE THREADED : "NPT" TYPE\n'
                '- IEC 60423, NEMA FB.I, OR EQUIVALENT\n'
                '28mm (1")\n'
                '(1 EA = 3 M, 288 EA)'
            ),
            "hs_code": "7303.00.1090", "qty": 864, "unit": "M",
            "net_kg": 1950.00, "gross_kg": 2150.00,
            "dim_l_cm": 330, "dim_w_cm": 110, "dim_h_cm": 105,
            "volume_cbm": 3.812
        },
        {
            "marks_no": 3, "of_text": "OF 5", "item_no": "4.1", "item_location": "DAS", "packing_style": "SKID",
            "description": (
                'RIGID STEEL CONDUIT  W/ ONE(1) COUPLING\n'
                '- HOT DIPPED GALVANIZED ( H.D.G. ) 65㎛ as per BS729\n'
                '- STANDARD LENGTH : 3.048M\n'
                '- PIPE THREADED : "NPT" TYPE\n'
                '- IEC 60423, NEMA FB.I, OR EQUIVALENT\n'
                '28mm (1")\n'
                '(1 EA = 3 M, 408 EA)'
            ),
            "hs_code": "7303.00.1090", "qty": 1224, "unit": "M",
            "net_kg": 2800.00, "gross_kg": 3000.00,
            "dim_l_cm": 330, "dim_w_cm": 110, "dim_h_cm": 105,
            "volume_cbm": 3.812
        },
        {
            "marks_no": 4, "of_text": "OF 5", "item_no": "4.1", "item_location": "DAS", "packing_style": "SKID",
            "description": (
                'RIGID STEEL CONDUIT  W/ ONE(1) COUPLING\n'
                '- HOT DIPPED GALVANIZED ( H.D.G. ) 65㎛ as per BS729\n'
                '- STANDARD LENGTH : 3.048M\n'
                '- PIPE THREADED : "NPT" TYPE\n'
                '- IEC 60423, NEMA FB.I, OR EQUIVALENT\n'
                '28mm (1")\n'
                '(1 EA = 3 M, 288 EA)'
            ),
            "hs_code": "7303.00.1090", "qty": 864, "unit": "M",
            "net_kg": 1950.00, "gross_kg": 2150.00,
            "dim_l_cm": 330, "dim_w_cm": 110, "dim_h_cm": 105,
            "volume_cbm": 3.812
        },
        {
            "marks_no": 5, "of_text": "OF 5", "item_no": "4.1", "item_location": "DAS", "packing_style": "SKID",
            "description": (
                'RIGID STEEL CONDUIT  W/ ONE(1) COUPLING\n'
                '- HOT DIPPED GALVANIZED ( H.D.G. ) 65㎛ as per BS729\n'
                '- STANDARD LENGTH : 3.048M\n'
                '- PIPE THREADED : "NPT" TYPE\n'
                '- IEC 60423, NEMA FB.I, OR EQUIVALENT\n'
                '54mm (2")\n'
                '(1 EA = 3 M, 120 EA)'
            ),
            "hs_code": "7303.00.1090", "qty": 360, "unit": "M",
            "net_kg": 1830.00, "gross_kg": 2030.00,
            "dim_l_cm": 330, "dim_w_cm": 110, "dim_h_cm": 110,
            "volume_cbm": 3.993
        },
    ],
}


# =============================================================================
# 2) STYLE HELPERS
# =============================================================================
def _side(style: str, color_hex: str) -> Side:
    return Side(style=style, color=color_hex)

def set_page_setup(ws):
    # Rider page is landscape
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_margins = PageMargins(left=0.30, right=0.30, top=0.40, bottom=0.40, header=0.00, footer=0.00)
    ws.sheet_view.showGridLines = False
    ws.print_options.gridLines = False

def set_dimensions(ws, last_row: int):
    # Columns
    for c in COLS:
        ws.column_dimensions[c].width = float(COL_WIDTHS[c])
    # Rows
    for r in range(1, last_row + 1):
        if r in ROW_HEIGHTS_FIXED:
            ws.row_dimensions[r].height = float(ROW_HEIGHTS_FIXED[r])

def set_cell(ws, addr: str, value: Any, *,
             bold: bool = False,
             size: int = 10,
             h: str = "left",
             v: str = "center",
             wrap: bool = False):
    c = ws[addr]
    c.value = value
    c.font = Font(name="Calibri", size=size, bold=bold)
    c.alignment = Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def apply_outline(ws, r1: int, c1: int, r2: int, c2: int, *, color="000000", thick=False):
    s = _side("thick" if thick else "thin", color)
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            top = cell.border.top
            bottom = cell.border.bottom
            left = cell.border.left
            right = cell.border.right
            if r == r1: top = s
            if r == r2: bottom = s
            if c == c1: left = s
            if c == c2: right = s
            cell.border = Border(top=top, bottom=bottom, left=left, right=right)

def apply_grid(ws, r1: int, c1: int, r2: int, c2: int, *, color="000000"):
    s = _side("thin", color)
    grid = Border(top=s, bottom=s, left=s, right=s)
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(row=r, column=c).border = grid


# =============================================================================
# 3) BUILD SHEET
# =============================================================================
def parse_items(payload: Dict[str, Any]) -> List[RiderItem]:
    out: List[RiderItem] = []
    for x in payload.get("items", []):
        out.append(RiderItem(
            marks_no=int(x["marks_no"]),
            of_text=str(x.get("of_text", "")),
            item_no=str(x.get("item_no", "")),
            item_location=str(x.get("item_location", "")),
            packing_style=str(x.get("packing_style", "")),
            description=str(x.get("description", "")),
            hs_code=str(x.get("hs_code", "")),
            qty=float(x.get("qty", 0)),
            unit=str(x.get("unit", "")),
            net_kg=float(x.get("net_kg", 0)),
            gross_kg=float(x.get("gross_kg", 0)),
            dim_l_cm=float(x.get("dim_l_cm", 0)),
            dim_w_cm=float(x.get("dim_w_cm", 0)),
            dim_h_cm=float(x.get("dim_h_cm", 0)),
            volume_cbm=float(x.get("volume_cbm", 0)),
        ))
    return out

def build_rider(ws, payload: Dict[str, Any]) -> int:
    title = payload.get("title", "PACKING LIST ATTACHED RIDER")
    page_no = payload.get("page_no", "PAGE NO.: 2 OF 2")
    section = payload.get("section_title", "STEEL CONDUIT (14TH)")
    items = parse_items(payload)

    # Compute bottom rows
    items_end = ROW_ITEMS_START + len(items) - 1
    row_gap4 = items_end + 1
    row_total = items_end + 2
    last_row = row_total

    # Base row heights
    set_dimensions(ws, last_row)
    for r in range(ROW_ITEMS_START, items_end + 1):
        ws.row_dimensions[r].height = ITEM_ROW_HEIGHT
    ws.row_dimensions[row_gap4].height = GAP_AFTER_ITEMS_HEIGHT
    ws.row_dimensions[row_total].height = TOTAL_ROW_HEIGHT

    # Fonts / Align
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=False)
    align_center_wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left_top_wrap = Alignment(horizontal="left", vertical="top", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center")

    # 1) Page no (top-right, R1)
    set_cell(ws, "R1", page_no, size=10, h="right")

    # 2) Title (Row 2, across A:R)
    ws.merge_cells(start_row=ROW_TITLE, start_column=1, end_row=ROW_TITLE, end_column=18)
    set_cell(ws, f"A{ROW_TITLE}", title, bold=True, size=18, h="center")

    # 3) Header row (Row 4)
    # Columns:
    # A Marks No. | B OF | C Item No. | D Item Location | E Packing Style | F Description | G spacer | H HS | I Qty | J Unit
    # K Net | L Gross | M L | N x | O W | P x | Q H | R Volume
    hdr = ROW_HEADER
    set_cell(ws, f"A{hdr}", "Marks & No. of Pkgs", bold=True, size=10, h="center", wrap=True)
    set_cell(ws, f"C{hdr}", "Item No.", bold=True, size=10, h="center", wrap=True)
    set_cell(ws, f"D{hdr}", "Item\nLocation", bold=True, size=10, h="center", wrap=True)
    set_cell(ws, f"E{hdr}", "Packing\nStyle", bold=True, size=10, h="center", wrap=True)
    set_cell(ws, f"F{hdr}", "Description of Goods", bold=True, size=10, h="center", wrap=True)
    set_cell(ws, f"H{hdr}", "HS CODE", bold=True, size=10, h="center", wrap=True)
    set_cell(ws, f"I{hdr}", "Q'ty", bold=True, size=10, h="center", wrap=True)
    set_cell(ws, f"J{hdr}", "Unit", bold=True, size=10, h="center", wrap=True)
    set_cell(ws, f"K{hdr}", "Net\nWeight\n(kgs)", bold=True, size=10, h="center", wrap=True)
    set_cell(ws, f"L{hdr}", "Gross\nWeight\n(kgs)", bold=True, size=10, h="center", wrap=True)
    set_cell(ws, f"M{hdr}", "Dimension\nL x W x H (CM)", bold=True, size=10, h="center", wrap=True)
    set_cell(ws, f"R{hdr}", "Volume\n(CBM)", bold=True, size=10, h="center", wrap=True)

    # Make header alignment consistent (apply wrap+center to A..R)
    for col in range(1, 19):
        cell = ws.cell(row=hdr, column=col)
        if cell.value is None:
            cell.value = ""  # keep grid visible
        cell.font = Font(name="Calibri", size=10, bold=True) if col not in (2, 7, 14, 16) else Font(name="Calibri", size=10, bold=False)
        cell.alignment = align_center_wrap

    # Header borders: thin grid + thick bottom line
    apply_grid(ws, hdr, 1, hdr, 18, color="000000")
    apply_outline(ws, hdr, 1, hdr, 18, color="000000", thick=False)

    # 4) Section title (Row 6, centered across A:R)
    ws.merge_cells(start_row=ROW_SECTION, start_column=1, end_row=ROW_SECTION, end_column=18)
    set_cell(ws, f"A{ROW_SECTION}", section, bold=True, size=12, h="center")

    # 5) Items
    r = ROW_ITEMS_START
    for it in items:
        # Left columns
        set_cell(ws, f"A{r}", it.marks_no, h="center")
        set_cell(ws, f"B{r}", it.of_text, h="center")
        set_cell(ws, f"C{r}", it.item_no, h="center")
        set_cell(ws, f"D{r}", it.item_location, h="center")
        set_cell(ws, f"E{r}", it.packing_style, h="center")

        # Description
        ws[f"F{r}"].value = it.description
        ws[f"F{r}"].font = Font(name="Calibri", size=10, bold=False)
        ws[f"F{r}"].alignment = align_left_top_wrap

        # Spacer G left blank
        set_cell(ws, f"G{r}", "")

        # HS / Qty / Unit
        set_cell(ws, f"H{r}", it.hs_code, h="center")
        ws[f"I{r}"].value = it.qty
        ws[f"I{r}"].number_format = "#,##0"
        ws[f"I{r}"].alignment = align_right
        ws[f"I{r}"].font = Font(name="Calibri", size=10)

        set_cell(ws, f"J{r}", it.unit, h="center")

        # Weights
        ws[f"K{r}"].value = it.net_kg
        ws[f"K{r}"].number_format = "#,##0.00"
        ws[f"K{r}"].alignment = align_right

        ws[f"L{r}"].value = it.gross_kg
        ws[f"L{r}"].number_format = "#,##0.00"
        ws[f"L{r}"].alignment = align_right

        # Dimension in M,N,O,P,Q as "L x W x H"
        ws[f"M{r}"].value = it.dim_l_cm
        ws[f"M{r}"].number_format = "#,##0"
        ws[f"M{r}"].alignment = align_right

        set_cell(ws, f"N{r}", "x", h="center")
        ws[f"O{r}"].value = it.dim_w_cm
        ws[f"O{r}"].number_format = "#,##0"
        ws[f"O{r}"].alignment = align_right

        set_cell(ws, f"P{r}", "x", h="center")
        ws[f"Q{r}"].value = it.dim_h_cm
        ws[f"Q{r}"].number_format = "#,##0"
        ws[f"Q{r}"].alignment = align_right

        # Volume
        ws[f"R{r}"].value = it.volume_cbm
        ws[f"R{r}"].number_format = "0.000"
        ws[f"R{r}"].alignment = align_right

        # Fonts for the row
        for col in range(1, 19):
            cell = ws.cell(row=r, column=col)
            if col != 6:  # Description already styled
                cell.font = Font(name="Calibri", size=10, bold=False)
            # keep vertical center except description
            if col != 6:
                if cell.alignment is None or cell.alignment == Alignment():
                    cell.alignment = align_center

        r += 1

    # Body table borders: grid for items range
    if items:
        apply_grid(ws, ROW_ITEMS_START, 1, items_end, 18, color="000000")
        apply_outline(ws, ROW_ITEMS_START, 1, items_end, 18, color="000000", thick=False)

    # 6) Total row (Row total)
    total_pkgs = len(items)
    total_qty = sum(x.qty for x in items)
    total_net = sum(x.net_kg for x in items)
    total_gross = sum(x.gross_kg for x in items)
    total_volume = sum(x.volume_cbm for x in items)

    set_cell(ws, f"A{row_total}", "TOTAL", bold=True, h="left")
    set_cell(ws, f"B{row_total}", total_pkgs, bold=True, h="center")

    ws[f"I{row_total}"].value = total_qty
    ws[f"I{row_total}"].number_format = "#,##0"
    ws[f"I{row_total}"].font = Font(name="Calibri", size=10, bold=True)
    ws[f"I{row_total}"].alignment = align_right

    set_cell(ws, f"J{row_total}", "M", bold=True, h="center")

    ws[f"K{row_total}"].value = total_net
    ws[f"K{row_total}"].number_format = "#,##0.00"
    ws[f"K{row_total}"].font = Font(name="Calibri", size=10, bold=True)
    ws[f"K{row_total}"].alignment = align_right

    ws[f"L{row_total}"].value = total_gross
    ws[f"L{row_total}"].number_format = "#,##0.00"
    ws[f"L{row_total}"].font = Font(name="Calibri", size=10, bold=True)
    ws[f"L{row_total}"].alignment = align_right

    # Put total volume at R, formatted to 3 decimals (image shows 19.239)
    ws[f"R{row_total}"].value = total_volume
    ws[f"R{row_total}"].number_format = "0.000"
    ws[f"R{row_total}"].font = Font(name="Calibri", size=10, bold=True)
    ws[f"R{row_total}"].alignment = align_right

    # Total row borders: thick top/bottom
    thick = _side("thick", "000000")
    for c in range(1, 19):
        cell = ws.cell(row=row_total, column=c)
        cell.border = Border(
            top=thick,
            bottom=thick,
            left=cell.border.left,
            right=cell.border.right
        )

    # 7) Blue outer frame (A1:R{last_row})
    apply_outline(ws, 1, 1, last_row, 18, color="0000D0", thick=True)

    # Freeze panes below header
    ws.freeze_panes = ws[f"A{ROW_ITEMS_START}"]

    return last_row


# =============================================================================
# 4) Workbook creator
# =============================================================================
def create_xlsx(out_path: str, payload: Optional[Dict[str, Any]] = None) -> None:
    if payload is None:
        payload = DEFAULT_DATA

    wb = Workbook()
    ws = wb.active
    ws.title = "PL_Rider_P2"

    set_page_setup(ws)

    # Column widths
    for c in COLS:
        ws.column_dimensions[c].width = float(COL_WIDTHS[c])

    last_row = build_rider(ws, payload)

    # Ensure any rows not set get default height (prevents Excel weirdness)
    for r in range(1, last_row + 1):
        if ws.row_dimensions[r].height is None:
            ws.row_dimensions[r].height = 15.0

    wb.save(out_path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True, help="Output XLSX path")
    args = ap.parse_args()
    create_xlsx(args.out, DEFAULT_DATA)
    print(f"Saved: {args.out}")


if __name__ == "__main__":
    main()
