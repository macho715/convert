#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CIPL / COMMERCIAL INVOICE (Page 1/1) — Fixed Layout Excel Generator
- 목적: 항차가 바뀌어도 "기본 정보만 입력"하면 동일 포맷의 Commercial Invoice(CIPL 1/1)를 자동 생성
- 본 코드: 사용자가 제공한 (1) 컬럼폭 A:J, (2) 행높이(20/15), (3) 셀 텍스트 좌표(행/열) 를 그대로 반영
- 출력: .xlsx (인쇄 설정 + 파란 외곽 프레임 + 주요 박스 테두리 포함)

Install:
  pip install openpyxl

Run:
  python cipl_invoice_p1.py --out COMMERCIAL_INVOICE_P1.xlsx

(선택) 입력을 JSON으로 바꾸고 싶으면:
  - 아래 DEFAULT_DATA의 값을 항차별로 교체하거나
  - load_json()을 사용해 외부 파일로 주입하면 됩니다.
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
from typing import Dict, List, Tuple, Optional

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins


# =============================================================================
# 0) TEMPLATE FIXED SETTINGS (from user)
# =============================================================================

# Column widths for A:J
COLS = ["A","B","C","D","E","F","G","H","I","J"]
COL_WIDTHS = {
    "A": 2.00,
    "B": 16.50,
    "C": 2.00,
    "D": 25.00,
    "E": 2.00,
    "F": 12.00,
    "G": 12.00,
    "H": 6.00,
    "I": 6.00,
    "J": 7.50,
}

# Row heights: as provided (20 for rows 1-4, 15 for rows 5-54)
ROW_HEIGHT_20 = {1, 2, 3, 4}
ROW_HEIGHT_DEFAULT = 15.00

# Used range
MIN_ROW, MAX_ROW = 1, 54
MIN_COL, MAX_COL = 1, 10  # A..J


# =============================================================================
# 1) DATA (항차별로 여기만 바꾸면 됨)
# =============================================================================

DEFAULT_DATA = {
    "page_no": "PAGE NO.: 1 OF 1",
    "title": "COMMERCIAL INVOICE",

    "shipper_name": "Samsung C&T Corporation",
    "shipper_addr1": "26, SANGIL-RO 6-GIL, GANGDONG-GU, SEOUL,",
    "shipper_addr2": "REPUBLIC OF KOREA (05288)",

    "consignee_name1": "Abu Dhabi Offshore Power Transmission Company",
    "consignee_name2": "Limited LLC",
    "consignee_addr1": "1301 & 1304 Al Wahda City, Commercial City Tower",
    "consignee_addr2": "Level-13, hazza Bin Zayed Street, Abu Dhabi, UAE",
    "consignee_addr3": "P.O. Box No: 108708",

    "notify_name": "DSV SOLUTIONS PJSC",
    "notify_addr1": "M19 Mussafah 2nd round about after Al Jabber Mussafah",
    "notify_addr2": "Abu Dhabi, UAE. P.O.Box 93971",

    "pol": "BUSAN",
    "pol_country": "KOREA",
    "pod": "KHALIFA PORT, ABU DHABI",
    "pod_country": "U.A.E.",

    "carrier": "HMM RAON 0022W",
    "sailing_on": "26-Dec-25",

    "invoice_no": "HVDC-ADOPT-SCT-0159",
    "invoice_date": "26-Dec-25",

    "hs_code": "7303.00.1090",
    "project_no": "AD164",
    "po_no": "5000802593",

    "mfg_name": "SANGDONG INDUSTRIES CO., LTD.",
    "mfg_addr1": "54, NEUNGHEODAE-RO 595BEON-GIL, NAMDONG-GU, INCHEON",
    "mfg_addr2": "REPUBLIC OF KOREA",
    "country_of_origin": "KOREA",

    "box_port_discharge": "ABU DHABI, U.A.E",
    "box_shipper": "SAMSUNG C&T CORPORATION",
    "box_consignee": "Abu Dhabi Offshore Power Transmission Company Limited LLC",
    "box_project_no": "AD164",
    "box_project_name": "Independent Subsea HVDC System Project (Project Lightning), UAE",
    "box_po_no": "5000802593",
    "box_item_desc": "STEEL CONDUIT (14TH)",
    "box_package_no": "5 PACKAGES",
    "box_net_wt": "11,330.00 KGS",
    "box_gross_wt": "12,330.00 KGS",
    "box_dimension": "19.239 CBM",
    "box_origin": "KOREA",

    "sum_item": "STEEL CONDUIT (14TH)",
    "sum_qty": "5 PKGS",
    "sum_fob": "USD 26,483.04",
    "sum_freight": "USD 0.00",
    "sum_insurance": "USD 0.00",
    "sum_cif": "USD 26,483.04",
}


# =============================================================================
# 2) STYLE HELPERS
# =============================================================================
def _side(style: str, color_hex: str) -> Side:
    return Side(style=style, color=color_hex)

def set_page_setup(ws):
    # Image looks portrait
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_margins = PageMargins(left=0.30, right=0.30, top=0.40, bottom=0.40, header=0.00, footer=0.00)
    ws.sheet_view.showGridLines = False
    ws.print_options.gridLines = False

def set_dimensions(ws):
    # columns
    for c in COLS:
        ws.column_dimensions[c].width = float(COL_WIDTHS[c])
    # rows
    for r in range(MIN_ROW, MAX_ROW + 1):
        if r in ROW_HEIGHT_20:
            ws.row_dimensions[r].height = 20.00
        else:
            ws.row_dimensions[r].height = ROW_HEIGHT_DEFAULT

def set_cell(ws, cell: str, value: str, *,
             bold: bool = False,
             size: int = 10,
             h: str = "left",
             v: str = "center",
             wrap: bool = False):
    ws[cell].value = value
    ws[cell].font = Font(name="Calibri", size=size, bold=bold)
    ws[cell].alignment = Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def apply_border_outline(ws, r1, c1, r2, c2, *, color="000000", thick=False):
    s = _side("thick" if thick else "thin", color)
    outline = Border(top=s, bottom=s, left=s, right=s)
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            top = cell.border.top
            bottom = cell.border.bottom
            left = cell.border.left
            right = cell.border.right
            if r == r1: top = outline.top
            if r == r2: bottom = outline.bottom
            if c == c1: left = outline.left
            if c == c2: right = outline.right
            cell.border = Border(top=top, bottom=bottom, left=left, right=right)

def apply_border_grid(ws, r1, c1, r2, c2, *, color="000000"):
    s = _side("thin", color)
    grid = Border(top=s, bottom=s, left=s, right=s)
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(row=r, column=c).border = grid

def merge(ws, r1, c1, r2, c2):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)


# =============================================================================
# 3) BUILD SHEET (좌표는 사용자가 제공한 표를 그대로 반영)
# =============================================================================
def build_commercial_invoice(ws, d: Dict[str, str]) -> None:
    # Base fonts
    title_font_size = 18

    # --- Outer blue frame (A1:J54)
    apply_border_outline(ws, MIN_ROW, MIN_COL, MAX_ROW, MAX_COL, color="0000D0", thick=True)

    # --- Title row (Row 3): COMMERCIAL INVOICE centered across B..I (approx)
    # Use merge for visual match (allowed; user only forbade merge for summary in Rider page)
    merge(ws, 3, 2, 3, 9)  # B3:I3
    set_cell(ws, "B3", d["title"], bold=True, size=title_font_size, h="center")

    # Page no (Row 4, right side)
    set_cell(ws, "J4", d["page_no"], bold=False, size=10, h="right")

    # --- Left block (Shipper/Consignee/Notify) and Right block (Invoice/Remarks)
    # Draw grid for the top form area (Row 5..24, Col A..J)
    apply_border_grid(ws, 5, 1, 24, 10, color="000000")
    apply_border_outline(ws, 5, 1, 24, 10, color="000000", thick=False)

    # Create a vertical split: Left A..E, Right F..J (visual like image)
    # Make the split line slightly thicker
    apply_border_outline(ws, 5, 6, 24, 6, color="000000", thick=True)  # thick line at F column edge

    # Labels + values (as user table)
    set_cell(ws, "B5", "①", bold=True)
    set_cell(ws, "C5", "Shipper / Exporter", bold=True)
    set_cell(ws, "B6", d["shipper_name"], bold=True)
    set_cell(ws, "B7", d["shipper_addr1"])
    set_cell(ws, "B8", d["shipper_addr2"])

    set_cell(ws, "B9", "②", bold=True)
    set_cell(ws, "C9", "Consignee", bold=True)
    set_cell(ws, "B10", d["consignee_name1"], bold=True)
    set_cell(ws, "B11", d["consignee_name2"], bold=True)
    set_cell(ws, "B12", d["consignee_addr1"])
    set_cell(ws, "B13", d["consignee_addr2"])
    set_cell(ws, "B14", d["consignee_addr3"])

    set_cell(ws, "B15", "③", bold=True)
    set_cell(ws, "C15", "Notify party", bold=True)
    set_cell(ws, "B16", d["notify_name"], bold=True)
    set_cell(ws, "B17", d["notify_addr1"])
    set_cell(ws, "B18", d["notify_addr2"])

    # Ports
    set_cell(ws, "B19", "④", bold=True)
    set_cell(ws, "C19", "Port of loading", bold=True)
    set_cell(ws, "E19", "⑤", bold=True, h="center")
    set_cell(ws, "F19", "Port of Discharge", bold=True)

    set_cell(ws, "B20", d["pol"], bold=False)
    set_cell(ws, "D20", d["pod"], bold=False)
    set_cell(ws, "B21", d["pol_country"], bold=False)
    set_cell(ws, "D21", d["pod_country"], bold=False)

    # Carrier / Sailing
    set_cell(ws, "B22", "⑥", bold=True)
    set_cell(ws, "C22", "Carrier", bold=True)
    set_cell(ws, "E22", "⑦", bold=True, h="center")
    set_cell(ws, "F22", "Sailing on or about", bold=True)

    set_cell(ws, "B24", d["carrier"])
    set_cell(ws, "D24", d["sailing_on"])

    # Right block: invoice no/date and remarks
    set_cell(ws, "F5", "⑧", bold=True)
    set_cell(ws, "G5", "No. & date of invoice", bold=True)
    set_cell(ws, "F6", d["invoice_no"], bold=True)
    set_cell(ws, "I6", "DATE :", bold=True, h="right")
    set_cell(ws, "J6", d["invoice_date"], bold=False, h="left")

    set_cell(ws, "F8", "⑨", bold=True)
    set_cell(ws, "G8", "Remarks", bold=True)

    set_cell(ws, "F9", "-", bold=True)
    set_cell(ws, "G9", f"HS CODE : {d['hs_code']}")
    set_cell(ws, "F11", "-", bold=True)
    set_cell(ws, "G11", f"PROJECT NO : {d['project_no']}")
    set_cell(ws, "F13", "-", bold=True)
    set_cell(ws, "G13", f"PO NO : {d['po_no']}")

    set_cell(ws, "F15", "-", bold=True)
    set_cell(ws, "G15", "MANUFACTURER'S NAME & ADDRESS :", bold=False)

    set_cell(ws, "G16", d["mfg_name"], bold=True)
    set_cell(ws, "G17", d["mfg_addr1"])
    set_cell(ws, "G18", d["mfg_addr2"])

    set_cell(ws, "F20", "-", bold=True)
    set_cell(ws, "G20", f"COUNTRY OF ORIGIN : {d['country_of_origin']}")

    # Row 25 labels
    set_cell(ws, "B25", "⑩", bold=True)
    set_cell(ws, "C25", "Mark and numbers", bold=True)
    set_cell(ws, "F25", "⑪", bold=True)
    set_cell(ws, "G25", "Description of Goods", bold=True)
    set_cell(ws, "C26", "of PKGS", bold=True)

    # Add a thick line under row 25 area (like section separation)
    apply_border_outline(ws, 25, 1, 25, 10, color="000000", thick=True)

    # --- Middle big project box (Rows 31..42)
    # Outer box
    apply_border_outline(ws, 31, 2, 42, 9, color="000000", thick=True)

    # Title inside box: PROJECT LIGHTING (HVDC)
    merge(ws, 29, 3, 29, 8)  # C29:H29 (as image; you may adjust)
    set_cell(ws, "C29", "PROJECT LIGHTING (HVDC)", bold=True, size=12, h="center")

    # Key-value lines inside box (use B..H like your row list)
    kv = [
        ("PORT OF DISCHARGE", d["box_port_discharge"]),
        ("SHIPPER", d["box_shipper"]),
        ("CONSIGNEE", d["box_consignee"]),
        ("PROJECT NO.", d["box_project_no"]),
        ("PROJECT NAME", d["box_project_name"]),
        ("P/O NO.", d["box_po_no"]),
        ("ITEM DESCRIPTION", d["box_item_desc"]),
        ("PACKAGE NO.", d["box_package_no"]),
        ("NET WEIGHT", d["box_net_wt"]),
        ("GROSS WEIGHT", d["box_gross_wt"]),
        ("DIMENSION", d["box_dimension"]),
        ("COUNTRY OF ORIGIN", d["box_origin"]),
    ]
    start_r = 31
    for i, (k, v) in enumerate(kv):
        rr = start_r + i
        set_cell(ws, f"B{rr}", k, bold=False)
        set_cell(ws, f"D{rr}", ":", bold=False, h="center")
        set_cell(ws, f"E{rr}", v, bold=False)

    # --- Bottom summary (Rows 44..49)
    set_cell(ws, "D44", "ITEM", bold=True, h="right")
    set_cell(ws, "E44", ":", bold=True, h="center")
    set_cell(ws, "F44", d["sum_item"], bold=False)

    set_cell(ws, "D45", "QUANTITY", bold=True, h="right")
    set_cell(ws, "E45", ":", bold=True, h="center")
    set_cell(ws, "F45", d["sum_qty"], bold=False)

    set_cell(ws, "D46", "TOTAL FOB VALUE", bold=True, h="right")
    set_cell(ws, "E46", ":", bold=True, h="center")
    set_cell(ws, "F46", d["sum_fob"], bold=False)

    set_cell(ws, "D47", "FREIGHT", bold=True, h="right")
    set_cell(ws, "E47", ":", bold=True, h="center")
    set_cell(ws, "F47", d["sum_freight"], bold=False)

    set_cell(ws, "D48", "INSURANCE", bold=True, h="right")
    set_cell(ws, "E48", ":", bold=True, h="center")
    set_cell(ws, "F48", d["sum_insurance"], bold=False)

    set_cell(ws, "D49", "TOTAL CIF  VALUE", bold=True, h="right")
    set_cell(ws, "E49", ":", bold=True, h="center")
    set_cell(ws, "F49", d["sum_cif"], bold=False)

    # A thin line above TOTAL CIF row (like image)
    apply_border_outline(ws, 49, 4, 49, 9, color="000000", thick=False)

    # Signature block
    set_cell(ws, "E51", "Signed by  _____________________________________", bold=False, h="left")
    set_cell(ws, "F52", "Project Manager of Project Lightning", bold=False)
    set_cell(ws, "F53", "Samsung C&T Corporation", bold=False)

    # General alignment tweaks: center the circled numbers column B
    for rr in range(5, 26):
        if ws[f"B{rr}"].value in ("①","②","③","④","⑥","⑩"):
            ws[f"B{rr}"].alignment = Alignment(horizontal="center", vertical="center")

    # Optional: freeze panes below header
    ws.freeze_panes = "A5"


def create_invoice_xlsx(out_path: str, data: Dict[str, str]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Commercial_Invoice_P1"

    set_page_setup(ws)
    set_dimensions(ws)

    build_commercial_invoice(ws, data)

    wb.save(out_path)


# =============================================================================
# 4) CLI
# =============================================================================
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True, help="Output XLSX path")
    args = ap.parse_args()

    create_invoice_xlsx(args.out, DEFAULT_DATA)
    print(f"Saved: {args.out}")


if __name__ == "__main__":
    main()
