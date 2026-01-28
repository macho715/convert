#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PACKING LIST (Page 1/1) — Fixed Layout Excel Generator
- 목적: 항차가 바뀌어도 "기본 정보만 입력"하면 동일 포맷의 PACKING LIST를 자동 생성
- 본 코드: 사용자가 제공한 (1) 컬럼폭 A:K, (2) 행높이(20/15), (3) 셀 텍스트 좌표(행/열) 기반

Dependencies:
  pip install openpyxl

Run:
  python packing_list_p1.py --out PACKING_LIST_P1.xlsx
"""

from __future__ import annotations

import argparse
from pathlib import Path
import sys
from typing import Dict, Optional

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.worksheet.page import PageMargins

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from excel_helpers import apply_border_outline_fast, apply_inner_grid_fast, get_alignment, get_font


# =============================================================================
# 0) TEMPLATE FIXED SETTINGS (from user)
# =============================================================================
# Columns: A..K
COL_WIDTHS = {
    "A": 2.00,
    "B": 18.00,
    "C": 2.00,
    "D": 24.00,
    "E": 2.00,
    "F": 11.00,
    "G": 11.00,
    "H": 5.50,
    "I": 5.50,
    "J": 3.50,
    "K": 7.50,
}

MIN_ROW, MAX_ROW = 1, 53
MIN_COL, MAX_COL = 1, 11  # A..K

ROWS_20 = {1, 2, 3, 4}
ROW_H_DEFAULT = 15.00


# =============================================================================
# 1) DATA (항차별로 여기만 교체)
# =============================================================================
DEFAULT_DATA = {
    "page_no": "PAGE NO.: 1 OF 1",
    "title": "PACKING LIST",

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

    # Project box
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
    "box_measure": "19.239 CBM",
    "box_origin": "KOREA",

    # Line summary (bottom)
    "line_goods": "STEEL CONDUIT (14TH)",
    "line_qty": "5",
    "line_qty_unit": "PKGS",

    "total_pkg_line": "TOTAL 5 PACKAGES",
    "total_net": "11,330.00 KGS",
    "total_gross": "12,330.00 KGS",
    "total_measure": "19.239 CBM",
    "details_note": "(DETAILS ARE AS PER ATTACHED SHEETS)",
}


# =============================================================================
# 2) STYLE HELPERS
# =============================================================================
def _side(style: str, color_hex: str) -> Side:
    return Side(style=style, color=color_hex)

def set_page_setup(ws):
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_margins = PageMargins(left=0.30, right=0.30, top=0.40, bottom=0.40, header=0.00, footer=0.00)
    ws.sheet_view.showGridLines = False
    ws.print_options.gridLines = False

def set_dimensions(ws):
    for col, w in COL_WIDTHS.items():
        ws.column_dimensions[col].width = float(w)
    for r in range(MIN_ROW, MAX_ROW + 1):
        ws.row_dimensions[r].height = 20.00 if r in ROWS_20 else ROW_H_DEFAULT

def set_cell(ws, cell: str, value: str, *,
             bold: bool = False,
             size: int = 10,
             h: str = "left",
             v: str = "center",
             wrap: bool = False):
    c = ws[cell]
    c.value = value
    c.font = get_font(name="Calibri", size=size, bold=bold)
    c.alignment = get_alignment(horizontal=h, vertical=v, wrap_text=wrap)

def apply_border_outline(ws, r1, c1, r2, c2, *, color="000000", thick=False):
    side = _side("thick" if thick else "thin", color)
    apply_border_outline_fast(ws, r1, c1, r2, c2, side)

def apply_border_grid(ws, r1, c1, r2, c2, *, color="000000"):
    side = _side("thin", color)
    apply_inner_grid_fast(ws, r1, c1, r2, c2, side)


# =============================================================================
# 3) BUILD SHEET (좌표: 사용자 제공 표 기반)
# =============================================================================
def build_packing_list(ws, d: Dict[str, str]) -> None:
    # Outer blue frame
    apply_border_outline(ws, MIN_ROW, MIN_COL, MAX_ROW, MAX_COL, color="0000D0", thick=True)

    # Title row (Row 3): PACKING LIST centered across B..J
    ws.merge_cells("B3:J3")
    set_cell(ws, "B3", d["title"], bold=True, size=18, h="center")

    # Page no (Row 4, right side)
    set_cell(ws, "K4", d["page_no"], size=10, h="right")

    # Top form grid (Row 5..23, Col A..K)
    apply_border_grid(ws, 5, 1, 23, 11, color="000000")
    apply_border_outline(ws, 5, 1, 23, 11, color="000000", thick=False)

    # Make a strong vertical split between left (A..E) and right (F..K)
    # Split line at F column edge
    apply_border_outline(ws, 5, 6, 23, 6, color="000000", thick=True)

    # Left block contents
    set_cell(ws, "B5", "①", bold=True, h="center")
    set_cell(ws, "C5", "Shipper / Exporter", bold=True)

    set_cell(ws, "B6", d["shipper_name"], bold=True)
    set_cell(ws, "B7", d["shipper_addr1"])
    set_cell(ws, "B8", d["shipper_addr2"])

    set_cell(ws, "B9", "②", bold=True, h="center")
    set_cell(ws, "C9", "Consignee", bold=True)

    set_cell(ws, "B10", d["consignee_name1"], bold=True)
    set_cell(ws, "B11", d["consignee_name2"], bold=True)
    set_cell(ws, "B12", d["consignee_addr1"])
    set_cell(ws, "B13", d["consignee_addr2"])
    set_cell(ws, "B14", d["consignee_addr3"])

    set_cell(ws, "B15", "③", bold=True, h="center")
    set_cell(ws, "C15", "Notify party", bold=True)

    set_cell(ws, "B16", d["notify_name"], bold=True)
    set_cell(ws, "B17", d["notify_addr1"])
    set_cell(ws, "B18", d["notify_addr2"])

    # Ports (Row 19..21)
    set_cell(ws, "B19", "④", bold=True, h="center")
    set_cell(ws, "C19", "Port of loading", bold=True)
    set_cell(ws, "E19", "⑤", bold=True, h="center")
    set_cell(ws, "F19", "Port of Discharge", bold=True)

    set_cell(ws, "B20", d["pol"])
    set_cell(ws, "D20", d["pod"])
    set_cell(ws, "B21", d["pol_country"])
    set_cell(ws, "D21", d["pod_country"])

    # Carrier / Sailing (Row 22..23)
    set_cell(ws, "B22", "⑥", bold=True, h="center")
    set_cell(ws, "C22", "Carrier", bold=True)
    set_cell(ws, "E22", "⑦", bold=True, h="center")
    set_cell(ws, "F22", "Sailing on or about", bold=True)

    set_cell(ws, "B23", d["carrier"])
    set_cell(ws, "D23", d["sailing_on"])

    # Right block: invoice + remarks
    set_cell(ws, "F5", "⑧", bold=True)
    set_cell(ws, "G5", "No. & date of invoice", bold=True)

    set_cell(ws, "F6", d["invoice_no"], bold=True)
    set_cell(ws, "J6", "DATE :", bold=True, h="right")
    set_cell(ws, "K6", d["invoice_date"], h="left")

    set_cell(ws, "F8", "⑨", bold=True)
    set_cell(ws, "G8", "Remarks", bold=True)

    set_cell(ws, "F9", "-", bold=True, h="center")
    set_cell(ws, "G9", f"HS CODE : {d['hs_code']}")

    set_cell(ws, "F11", "-", bold=True, h="center")
    set_cell(ws, "G11", f"PROJECT NO : {d['project_no']}")

    set_cell(ws, "F13", "-", bold=True, h="center")
    set_cell(ws, "G13", f"PO NO : {d['po_no']}")

    set_cell(ws, "F15", "-", bold=True, h="center")
    set_cell(ws, "G15", "MANUFACTURER'S NAME & ADDRESS :")

    set_cell(ws, "G16", d["mfg_name"], bold=True)
    set_cell(ws, "G17", d["mfg_addr1"])
    set_cell(ws, "G18", d["mfg_addr2"])

    set_cell(ws, "F20", "-", bold=True, h="center")
    set_cell(ws, "G20", f"COUNTRY OF ORIGIN : {d['country_of_origin']}")

    # Header row for packing list table (Row 24..25)
    # Row 24: ⑩, ⑪, ⑫, ⑬, ⑭, ⑮ + headings
    set_cell(ws, "B24", "⑩", bold=True, h="center")
    set_cell(ws, "C24", "Mark and numbers", bold=True)

    set_cell(ws, "F24", "⑪", bold=True, h="center")
    set_cell(ws, "G24", "Description of Goods", bold=True)

    set_cell(ws, "H24", "⑫", bold=True, h="center")
    set_cell(ws, "I24", "Quantity", bold=True, h="center")

    set_cell(ws, "J24", "⑬  Net-", bold=True, h="center")
    set_cell(ws, "J25", "Weight", bold=True, h="center")

    set_cell(ws, "K24", "⑭  Gross-", bold=True, h="center")
    set_cell(ws, "K25", "Weight", bold=True, h="center")

    # Measurement header is split in the image; here we place it across J/K area minimally
    # If you want a dedicated measurement column, you can allocate in your own spec.
    # User table indicates Measure -ment on the far right; we place it in K24/K25 already used.
    # So we will instead place "⑮ Measure -ment" in I24 (small), consistent with provided lines:
    # (to keep strict "A..K only" and avoid new columns)
    set_cell(ws, "I25", "⑮ Measure\n-ment", bold=True, h="center", wrap=True)

    # Thick separator under row 25
    apply_border_outline(ws, 25, 1, 25, 11, color="000000", thick=True)

    # Rows 26..29 intentionally left blank per coordinate table.
    # Project box (like invoice): rows 30..41, columns B..I (fits A..K layout)
    # Outer box
    apply_border_outline(ws, 30, 2, 41, 9, color="000000", thick=True)

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
        ("DIMENSION", d["box_measure"]),
        ("COUNTRY OF ORIGIN", d["box_origin"]),
    ]
    start_r = 30
    for i, (k, v) in enumerate(kv):
        rr = start_r + i
        set_cell(ws, f"B{rr}", k)
        set_cell(ws, f"D{rr}", ":", h="center")
        set_cell(ws, f"E{rr}", v)

    # Bottom line item (Row 43)
    # "STEEL CONDUIT (14TH)" + qty + unit (align to the right-side qty columns)
    set_cell(ws, "C43", d["line_goods"], bold=False)
    set_cell(ws, "H43", d["line_qty"], h="center")
    set_cell(ws, "I43", d["line_qty_unit"], h="center")

    # Totals row (Row 46) + note (Row 47)
    set_cell(ws, "C46", d["total_pkg_line"], bold=False)
    set_cell(ws, "J46", d["total_net"], h="center")
    set_cell(ws, "K46", d["total_gross"], h="center")
    # Measure (place in I46 due to A..K constraint)
    set_cell(ws, "I46", d["total_measure"], h="center")

    set_cell(ws, "F47", d["details_note"], h="center")

    # Signature (Row 51..53)
    set_cell(ws, "E51", "Signed by  _____________________________________", h="left")
    set_cell(ws, "F52", "Project Manager of Project Lightning")
    set_cell(ws, "F53", "Samsung C&T Corporation")


def create_packing_list_xlsx(out_path: str, data: Optional[Dict[str, str]] = None) -> None:
    if data is None:
        data = DEFAULT_DATA

    wb = Workbook()
    ws = wb.active
    ws.title = "Packing_List_P1"

    set_page_setup(ws)
    set_dimensions(ws)
    build_packing_list(ws, data)

    wb.save(out_path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True, help="Output XLSX path")
    args = ap.parse_args()
    create_packing_list_xlsx(args.out, DEFAULT_DATA)
    print(f"Saved: {args.out}")


if __name__ == "__main__":
    main()
