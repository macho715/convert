#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CIPL Excel Generator (Fixed Rider Layout)
- Input: JSON (header + line items)
- Output: .xlsx (same format every voyage)

Install:
  pip install openpyxl

Run:
  python cipl_generator.py --in cipl_input.json --out CIPL_OUTPUT.xlsx
"""

from __future__ import annotations

import argparse
import json
from dataclasses import dataclass
from typing import Any, Dict, List, Optional

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins


# ----------------------------
# Config (template fixed)
# ----------------------------
COL_WIDTHS = [7, 7, 7, 47, 3, 3, 14, 10, 10, 10, 13, 13]  # A..L

# Rows (fixed to coordinate table)
ROW_PAGE = 1
ROW_TITLE = 2
ROW_HEADER = 3
ROW_GAP_1 = 4
ROW_SECTION = 5
ROW_GAP_2 = 6
ROW_ITEMS_START = 7
ROW_GAP_3 = 12
ROW_SUMMARY_START = 13
ROW_TOTAL = 17

ROW_HEIGHTS = {
    ROW_PAGE: 18.0,
    ROW_TITLE: 30.0,
    ROW_HEADER: 35.0,
    ROW_GAP_1: 7.0,
    ROW_SECTION: 26.0,
    ROW_GAP_2: 7.0,
    ROW_GAP_3: 7.0,
    ROW_SUMMARY_START: 18.0,
    ROW_SUMMARY_START + 1: 18.0,
    ROW_SUMMARY_START + 2: 18.0,
    ROW_SUMMARY_START + 3: 18.0,
    ROW_TOTAL: 18.0,
}
ITEM_ROW_HEIGHT = 95.0


@dataclass
class LineItem:
    marks_no: int
    of_text: str
    item_no: str
    description: str
    hs_code: str
    origin: str
    qty: float
    unit: str
    unit_price: float
    total_price: float


# ----------------------------
# Helpers
# ----------------------------
def _side(style: str, color_hex: str) -> Side:
    return Side(style=style, color=color_hex)

def set_col_widths(ws) -> None:
    for i, w in enumerate(COL_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = float(w)

def set_page_setup(ws) -> None:
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_margins = PageMargins(left=0.30, right=0.30, top=0.50, bottom=0.50, header=0.00, footer=0.00)
    ws.sheet_view.showGridLines = False
    ws.print_options.gridLines = False

def apply_border_range(ws, r1: int, r2: int, c1: int, c2: int,
                       outline: Optional[Border] = None,
                       inside: Optional[Border] = None) -> None:
    """
    outline: only outer edges
    inside: inner grid lines
    """
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            b = cell.border

            top, bottom, left, right = b.top, b.bottom, b.left, b.right

            if inside:
                if r > r1: top = inside.top
                if r < r2: bottom = inside.bottom
                if c > c1: left = inside.left
                if c < c2: right = inside.right

            if outline:
                if r == r1: top = outline.top
                if r == r2: bottom = outline.bottom
                if c == c1: left = outline.left
                if c == c2: right = outline.right

            cell.border = Border(top=top, bottom=bottom, left=left, right=right)

def load_input(path: str) -> Dict[str, Any]:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def parse_items(data: Dict[str, Any]) -> List[LineItem]:
    items = []
    for x in data.get("items", []):
        items.append(LineItem(
            marks_no=int(x["marks_no"]),
            of_text=str(x.get("of_text", "")),
            item_no=str(x.get("item_no", "")),
            description=str(x.get("description", "")),
            hs_code=str(x.get("hs_code", "")),
            origin=str(x.get("origin", "")),
            qty=float(x.get("qty", 0)),
            unit=str(x.get("unit", "")),
            unit_price=float(x.get("unit_price", 0)),
            total_price=float(x.get("total_price", 0)),
        ))
    return items


# ----------------------------
# Builder
# ----------------------------
def build_sheet(ws, payload: Dict[str, Any]) -> None:
    meta = payload.get("meta", {})
    section = payload.get("section_title", "STEEL CONDUIT (14TH)")
    title = payload.get("title", "COMMERCIAL INVOICE ATTACHED RIDER")
    page_no = meta.get("page_no", "PAGE NO.: 2 OF 2")

    totals = payload.get("totals", {})
    total_fob = float(totals.get("total_fob", 0.0))
    freight = totals.get("freight", "")
    insurance = totals.get("insurance", "")

    items = parse_items(payload)

    # Fonts/Align
    font_title = Font(name="Calibri", size=18, bold=True)
    font_hdr = Font(name="Calibri", size=11, bold=True)
    font_body = Font(name="Calibri", size=10)
    font_section = Font(name="Calibri", size=12, bold=True)

    center = Alignment(horizontal="center", vertical="center")
    center_wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)
    right = Alignment(horizontal="right", vertical="center")
    left_top_wrap = Alignment(horizontal="left", vertical="top", wrap_text=True)

    for row_num, height in ROW_HEIGHTS.items():
        ws.row_dimensions[row_num].height = height
    for row_num in range(ROW_ITEMS_START, ROW_GAP_3):
        ws.row_dimensions[row_num].height = ITEM_ROW_HEIGHT

    # Page No (L1)
    ws[f"L{ROW_PAGE}"].value = page_no
    ws[f"L{ROW_PAGE}"].alignment = Alignment(horizontal="right", vertical="center")
    ws[f"L{ROW_PAGE}"].font = Font(name="Calibri", size=10)

    # Title (A2:L2)
    ws.merge_cells(start_row=ROW_TITLE, start_column=1, end_row=ROW_TITLE, end_column=12)
    ws[f"A{ROW_TITLE}"].value = title
    ws[f"A{ROW_TITLE}"].font = font_title
    ws[f"A{ROW_TITLE}"].alignment = center

    # Header row merges: A:B, D:F
    ws.merge_cells(start_row=ROW_HEADER, start_column=1, end_row=ROW_HEADER, end_column=2)  # A:B
    ws.merge_cells(start_row=ROW_HEADER, start_column=4, end_row=ROW_HEADER, end_column=6)  # D:F

    ws[f"A{ROW_HEADER}"].value = "Marks & No. of Pkgs"
    ws[f"C{ROW_HEADER}"].value = "Item No."
    ws[f"D{ROW_HEADER}"].value = "DESCRIPTION"
    ws[f"G{ROW_HEADER}"].value = "HS CODE"
    ws[f"H{ROW_HEADER}"].value = "ORIGIN"
    ws[f"I{ROW_HEADER}"].value = "Q'TY"
    ws[f"J{ROW_HEADER}"].value = "UNIT"
    ws[f"K{ROW_HEADER}"].value = "Unit Price\n(USD)"
    ws[f"L{ROW_HEADER}"].value = "Total Price\n(USD)"

    for addr in [f"A{ROW_HEADER}", f"C{ROW_HEADER}", f"D{ROW_HEADER}", f"G{ROW_HEADER}", f"H{ROW_HEADER}",
                 f"I{ROW_HEADER}", f"J{ROW_HEADER}", f"K{ROW_HEADER}", f"L{ROW_HEADER}"]:
        ws[addr].font = font_hdr
        ws[addr].alignment = center_wrap if addr in (f"K{ROW_HEADER}", f"L{ROW_HEADER}") else center

    # Section row
    ws.merge_cells(start_row=ROW_SECTION, start_column=1, end_row=ROW_SECTION, end_column=12)
    ws[f"A{ROW_SECTION}"].value = section
    ws[f"A{ROW_SECTION}"].font = font_section
    ws[f"A{ROW_SECTION}"].alignment = center

    # Line items
    max_items = ROW_GAP_3 - ROW_ITEMS_START
    r = ROW_ITEMS_START
    for it in items[:max_items]:
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=6)  # D:F

        ws[f"A{r}"].value = it.marks_no
        ws[f"B{r}"].value = it.of_text
        ws[f"C{r}"].value = it.item_no
        ws[f"D{r}"].value = it.description
        ws[f"G{r}"].value = it.hs_code
        ws[f"H{r}"].value = it.origin
        ws[f"I{r}"].value = it.qty
        ws[f"J{r}"].value = it.unit
        ws[f"K{r}"].value = it.unit_price
        ws[f"L{r}"].value = it.total_price

        # Formats
        ws[f"A{r}"].alignment = center
        ws[f"B{r}"].alignment = center
        ws[f"C{r}"].alignment = center
        ws[f"D{r}"].alignment = left_top_wrap
        ws[f"G{r}"].alignment = center
        ws[f"H{r}"].alignment = center
        ws[f"I{r}"].alignment = right
        ws[f"J{r}"].alignment = center
        ws[f"K{r}"].alignment = right
        ws[f"L{r}"].alignment = right

        ws[f"I{r}"].number_format = "#,##0"
        ws[f"K{r}"].number_format = "#,##0.00"
        ws[f"L{r}"].number_format = "#,##0.00"

        for col in "ABCDEFGHIJKL":
            ws[f"{col}{r}"].font = font_body

        r += 1

    last_item_row = ROW_GAP_3 - 1

    # Summary block at H (NO MERGE)
    sum_r = ROW_SUMMARY_START
    ws[f"H{sum_r}"].value = "TOTAL FOB CHARGE"
    ws[f"H{sum_r+1}"].value = "FREIGHT"
    ws[f"H{sum_r+2}"].value = "INSURANCE"
    ws[f"H{sum_r}"].alignment = right
    ws[f"H{sum_r+1}"].alignment = right
    ws[f"H{sum_r+2}"].alignment = right
    ws[f"L{sum_r}"].value = total_fob
    ws[f"L{sum_r}"].number_format = "#,##0.00"
    ws[f"L{sum_r}"].alignment = right

    # Optional: write freight/insurance values if provided (still no-merge)
    if isinstance(freight, (int, float)) and freight != "":
        ws[f"L{sum_r+1}"].value = float(freight)
        ws[f"L{sum_r+1}"].number_format = "#,##0.00"
        ws[f"L{sum_r+1}"].alignment = right
    if isinstance(insurance, (int, float)) and insurance != "":
        ws[f"L{sum_r+2}"].value = float(insurance)
        ws[f"L{sum_r+2}"].number_format = "#,##0.00"
        ws[f"L{sum_r+2}"].alignment = right

    # Bottom TOTAL row (optional)
    total_row = ROW_TOTAL
    ws[f"A{total_row}"].value = "TOTAL"
    ws[f"B{total_row}"].value = len(items)

    qty_total = sum(float(x.qty) for x in items)
    ws[f"I{total_row}"].value = qty_total
    ws[f"I{total_row}"].number_format = "#,##0"
    ws[f"J{total_row}"].value = payload.get("bottom_unit", "M")
    ws[f"L{total_row}"].value = total_fob
    ws[f"L{total_row}"].number_format = "#,##0.00"

    for addr in [f"A{total_row}", f"B{total_row}", f"D{total_row}", f"I{total_row}", f"J{total_row}", f"L{total_row}"]:
        ws[addr].font = Font(name="Calibri", size=10, bold=True)
    ws[f"A{total_row}"].alignment = center
    ws[f"B{total_row}"].alignment = center
    ws[f"D{total_row}"].alignment = center
    ws[f"I{total_row}"].alignment = right
    ws[f"J{total_row}"].alignment = center
    ws[f"L{total_row}"].alignment = right

    # Borders (layered)
    blue_thick = _side("thick", "0000D0")
    black_thick = _side("thick", "000000")
    black_thin = _side("thin", "000000")

    # Blue page frame A1:L{total_row}
    outline_blue = Border(top=blue_thick, bottom=blue_thick, left=blue_thick, right=blue_thick)
    apply_border_range(ws, 1, total_row, 1, 12, outline=outline_blue, inside=None)

    # Main table outline+grid: header -> last item row
    outline_black = Border(top=black_thick, bottom=black_thick, left=black_thick, right=black_thick)
    inside_black = Border(top=black_thin, bottom=black_thin, left=black_thin, right=black_thin)
    apply_border_range(ws, ROW_HEADER, last_item_row, 1, 12, outline=outline_black, inside=inside_black)

    # Header thick top/bottom
    for c in range(1, 13):
        cell = ws.cell(row=ROW_HEADER, column=c)
        cell.border = Border(top=black_thick, bottom=black_thick, left=cell.border.left, right=cell.border.right)

    # TOTAL row thick top/bottom
    for c in range(1, 13):
        cell = ws.cell(row=total_row, column=c)
        cell.border = Border(top=black_thick, bottom=black_thick, left=cell.border.left, right=cell.border.right)


def generate_xlsx(input_path: str, output_path: str) -> None:
    payload = load_input(input_path)

    wb = Workbook()
    ws = wb.active
    ws.title = payload.get("sheet_name", "Rider_P2")

    set_page_setup(ws)
    set_col_widths(ws)
    build_sheet(ws, payload)

    wb.save(output_path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="inp", required=True, help="Input JSON path")
    ap.add_argument("--out", required=True, help="Output XLSX path")
    args = ap.parse_args()
    generate_xlsx(args.inp, args.out)
    print(f"Saved: {args.out}")


if __name__ == "__main__":
    main()
