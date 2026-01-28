# -*- coding: utf-8 -*-
"""
commons → 4페이지 data_dict 생성 Patch (Single Source of Truth)

목적
- 항차마다 바뀌는 commons 1회 입력으로
  1) Commercial Invoice (P1)
  2) Commercial Invoice Rider (P2)
  3) Packing List (P1)
  4) Packing List Rider (P2)
  에 들어갈 data_dict를 자동 생성한다.

전제(당신이 업로드한 기존 스크립트와 호환)
- COMMERCIAL INVOICE.PY  : build_commercial_invoice(ws, data_dict)
- PACKING LIST.PY        : build_packing_list(ws, data_dict)
- CI RIDER.PY            : build_sheet(ws, payload)   # ci_rider_p2 dict
- PACKING LIST ATTACHED RIDER.PY : build_rider(ws, payload)  # pl_rider_p2 dict

입력 payload (권장)
{
  "commons": {...},
  "static_parties": {...},
  "ci_rider_items": [...],  # 선택
  "pl_rider_items": [...]   # 선택
}

가정(문서 일치 우선)
- CBM/Volume은 3 decimals 유지(예: 19.239)
- USD/Weight는 2 decimals 유지
"""

from __future__ import annotations

from copy import deepcopy
from typing import Any, Dict, List, Tuple


# ---------------------------
# Format helpers
# ---------------------------
def _fmt_usd(x: float) -> str:
    return f"USD {x:,.2f}"

def _fmt_kg(x: float) -> str:
    return f"{x:,.2f} KGS"

def _fmt_cbm(x: float, decimals: int = 3) -> str:
    return f"{x:.{decimals}f} CBM"

def _fmt_pkgs(n: int) -> str:
    return f"{n} PKGS"

def _require(d: Dict[str, Any], key: str) -> Any:
    if key not in d:
        raise KeyError(f"Missing required key in commons/static_parties: {key}")
    return d[key]

def _default_if_missing_item_fields(items: List[Dict[str, Any]], commons: Dict[str, Any]) -> List[Dict[str, Any]]:
    """
    Rider line items에 hs_code / item_description 등이 누락된 경우 commons로 보정.
    """
    out = []
    for it in items:
        x = deepcopy(it)
        if "hs_code" not in x or x["hs_code"] in (None, ""):
            x["hs_code"] = commons.get("hs_code", "")
        out.append(x)
    return out


# ---------------------------
# Core patch: commons -> 4 dicts
# ---------------------------
def make_4page_data_dicts(payload: Dict[str, Any], *, cbm_decimals: int = 3) -> Dict[str, Dict[str, Any]]:
    """
    returns:
    {
      "ci_p1": {...},
      "ci_rider_p2": {...},
      "pl_p1": {...},
      "pl_rider_p2": {...}
    }
    """
    commons = _require(payload, "commons")
    parties = _require(payload, "static_parties")

    # Optional rider items
    ci_rider_items = payload.get("ci_rider_items", [])
    pl_rider_items = payload.get("pl_rider_items", [])

    # ---- commons required fields (you listed)
    pol = _require(commons, "pol")
    pod = _require(commons, "pod")
    carrier = _require(commons, "carrier")
    sailing_on = _require(commons, "sailing_on")
    invoice_no = _require(commons, "invoice_no")
    invoice_date = _require(commons, "invoice_date")
    hs_code = _require(commons, "hs_code")

    mfg_name = _require(commons, "manufacturer_name")
    mfg_a1 = _require(commons, "manufacturer_address_1")
    mfg_a2 = _require(commons, "manufacturer_address_2")

    item_desc = _require(commons, "item_description")
    package_no = _require(commons, "package_no")
    net_kg = float(_require(commons, "net_weight_kg"))
    gross_kg = float(_require(commons, "gross_weight_kg"))
    dim_cbm = float(_require(commons, "dimension_cbm"))

    item = _require(commons, "item")
    qty_pkgs = int(_require(commons, "quantity_pkgs"))

    freight = float(_require(commons, "freight_usd"))
    insurance = float(_require(commons, "insurance_usd"))
    total_cif = float(_require(commons, "total_cif_usd"))

    # ---- parties / static
    shipper_name = _require(parties, "shipper_name")
    shipper_addr1 = _require(parties, "shipper_addr1")
    shipper_addr2 = _require(parties, "shipper_addr2")

    consignee_name1 = _require(parties, "consignee_name1")
    consignee_name2 = _require(parties, "consignee_name2")
    consignee_addr1 = _require(parties, "consignee_addr1")
    consignee_addr2 = _require(parties, "consignee_addr2")
    consignee_addr3 = _require(parties, "consignee_addr3")

    notify_name = _require(parties, "notify_name")
    notify_addr1 = _require(parties, "notify_addr1")
    notify_addr2 = _require(parties, "notify_addr2")

    project_no = _require(parties, "project_no")
    project_name = _require(parties, "project_name")
    po_no = _require(parties, "po_no")
    country_of_origin = _require(parties, "country_of_origin")

    # 문서 표기: POD의 국가(U.A.E.)는 parties에 없으면 그대로 두거나 별도 필드로 확장 가능
    pod_country = parties.get("pod_country", "U.A.E.")
    pol_country = parties.get("pol_country", "KOREA")

    # ---- CI P1 dict (COMMERCIAL INVOICE.PY 기대 키)
    ci_p1 = {
        "page_no": "PAGE NO.: 1 OF 1",
        "title": "COMMERCIAL INVOICE",

        # Shipper / Consignee / Notify
        "shipper_name": shipper_name,
        "shipper_addr1": shipper_addr1,
        "shipper_addr2": shipper_addr2,

        "consignee_name1": consignee_name1,
        "consignee_name2": consignee_name2,
        "consignee_addr1": consignee_addr1,
        "consignee_addr2": consignee_addr2,
        "consignee_addr3": consignee_addr3,

        "notify_name": notify_name,
        "notify_addr1": notify_addr1,
        "notify_addr2": notify_addr2,

        # POL/POD + Carrier/Sailing
        "pol": pol,
        "pol_country": pol_country,
        "pod": pod,
        "pod_country": pod_country,
        "carrier": carrier,
        "sailing_on": sailing_on,

        # Invoice / Remarks
        "invoice_no": invoice_no,
        "invoice_date": invoice_date,
        "hs_code": hs_code,
        "project_no": project_no,
        "po_no": po_no,

        "mfg_name": mfg_name,
        "mfg_addr1": mfg_a1,
        "mfg_addr2": mfg_a2,
        "country_of_origin": country_of_origin,

        # Project box
        "box_port_discharge": f"{pod_country and 'ABU DHABI, ' or ''}{pod_country}".strip(),  # safe default
        "box_shipper": shipper_name.upper(),
        "box_consignee": f"{consignee_name1} {consignee_name2}".strip(),
        "box_project_no": project_no,
        "box_project_name": project_name,
        "box_po_no": po_no,
        "box_item_desc": item_desc,
        "box_package_no": package_no,
        "box_net_wt": _fmt_kg(net_kg),
        "box_gross_wt": _fmt_kg(gross_kg),
        "box_dimension": _fmt_cbm(dim_cbm, decimals=cbm_decimals),
        "box_origin": country_of_origin,

        # Bottom summary
        "sum_item": item,
        "sum_qty": _fmt_pkgs(qty_pkgs),
        "sum_fob": _fmt_usd(total_cif),          # NOTE: CI P1에는 FOB/CIF 둘 다 쓰는데, 원본 예시에서 FOB=CIF(운임/보험 0) 패턴.
        "sum_freight": _fmt_usd(freight),
        "sum_insurance": _fmt_usd(insurance),
        "sum_cif": _fmt_usd(total_cif),
    }

    # ---- PL P1 dict (PACKING LIST.PY 기대 키)
    pl_p1 = {
        "page_no": "PAGE NO.: 1 OF 1",
        "title": "PACKING LIST",

        # parties
        "shipper_name": shipper_name,
        "shipper_addr1": shipper_addr1,
        "shipper_addr2": shipper_addr2,

        "consignee_name1": consignee_name1,
        "consignee_name2": consignee_name2,
        "consignee_addr1": consignee_addr1,
        "consignee_addr2": consignee_addr2,
        "consignee_addr3": consignee_addr3,

        "notify_name": notify_name,
        "notify_addr1": notify_addr1,
        "notify_addr2": notify_addr2,

        # POL/POD + Carrier/Sailing
        "pol": pol,
        "pol_country": pol_country,
        "pod": pod,
        "pod_country": pod_country,
        "carrier": carrier,
        "sailing_on": sailing_on,

        # invoice/remarks
        "invoice_no": invoice_no,
        "invoice_date": invoice_date,
        "hs_code": hs_code,
        "project_no": project_no,
        "po_no": po_no,

        "mfg_name": mfg_name,
        "mfg_addr1": mfg_a1,
        "mfg_addr2": mfg_a2,
        "country_of_origin": country_of_origin,

        # project box
        "box_port_discharge": f"ABU DHABI, {pod_country}".strip(),
        "box_shipper": shipper_name.upper(),
        "box_consignee": f"{consignee_name1} {consignee_name2}".strip(),
        "box_project_no": project_no,
        "box_project_name": project_name,
        "box_po_no": po_no,
        "box_item_desc": item_desc,
        "box_package_no": package_no,
        "box_net_wt": _fmt_kg(net_kg),
        "box_gross_wt": _fmt_kg(gross_kg),
        "box_measure": _fmt_cbm(dim_cbm, decimals=cbm_decimals),
        "box_origin": country_of_origin,

        # bottom lines
        "line_goods": item_desc,
        "line_qty": str(qty_pkgs),
        "line_qty_unit": "PKGS",

        "total_pkg_line": f"TOTAL {qty_pkgs} PACKAGES",
        "total_net": _fmt_kg(net_kg),
        "total_gross": _fmt_kg(gross_kg),
        "total_measure": _fmt_cbm(dim_cbm, decimals=cbm_decimals),
        "details_note": "(DETAILS ARE AS PER ATTACHED SHEETS)",
    }

    # ---- CI Rider P2 dict (CI RIDER.PY 기대 키)
    # CI Rider는 라인아이템(가격표)이 핵심이므로 items는 외부 입력을 사용.
    # section_title은 commons.item_description로 고정 전파.
    ci_rider_p2 = {
        "sheet_name": "CI_Rider_P2",
        "title": "COMMERCIAL INVOICE ATTACHED RIDER",
        "section_title": item_desc,
        "bottom_terms": "CIF ABU DHABI PORT",
        "bottom_unit": "M",
        "meta": {"page_no": "PAGE NO.: 2 OF 2"},
        "items": _default_if_missing_item_fields(ci_rider_items, commons),
        "totals": {
            # CI Rider 예시는 TOTAL FOB CHARGE만 표기. 운임/보험은 공란 가능.
            "total_fob": float(total_cif),  # 운임/보험 0인 경우 FOB=CIF 패턴. 필요 시 별도 fob_usd를 commons에 추가 권장.
            "freight": "" if freight == 0.0 else float(freight),
            "insurance": "" if insurance == 0.0 else float(insurance),
        },
    }

    # ---- PL Rider P2 dict (PACKING LIST ATTACHED RIDER.PY 기대 키)
    pl_rider_p2 = {
        "page_no": "PAGE NO.: 2 OF 2",
        "title": "PACKING LIST ATTACHED RIDER",
        "section_title": item_desc,
        "items": _default_if_missing_item_fields(pl_rider_items, commons),
    }

    return {
        "ci_p1": ci_p1,
        "ci_rider_p2": ci_rider_p2,
        "pl_p1": pl_p1,
        "pl_rider_p2": pl_rider_p2,
    }


# ---------------------------
# Example usage in your unified runner
# ---------------------------
if __name__ == "__main__":
    # Example minimal payload (replace with your voyage_input.json loaded dict)
    payload_example = {
        "commons": {
            "pol": "BUSAN",
            "pod": "KHALIFA PORT, ABU DHABI",
            "carrier": "HMM RAON 0022W",
            "sailing_on": "26-Dec-25",
            "invoice_no": "HVDC-ADOPT-SCT-0159",
            "invoice_date": "26-Dec-25",
            "hs_code": "7303.00.1090",
            "manufacturer_name": "SANGDONG INDUSTRIES CO., LTD.",
            "manufacturer_address_1": "54, NEUNGHEODAE-RO 595BEON-GIL, NAMDONG-GU, INCHEON",
            "manufacturer_address_2": "REPUBLIC OF KOREA",
            "item_description": "STEEL CONDUIT (14TH)",
            "package_no": "5 PACKAGES",
            "net_weight_kg": 11330.00,
            "gross_weight_kg": 12330.00,
            "dimension_cbm": 19.239,
            "item": "STEEL CONDUIT (14TH)",
            "quantity_pkgs": 5,
            "freight_usd": 0.00,
            "insurance_usd": 0.00,
            "total_cif_usd": 26483.04
        },
        "static_parties": {
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
            "project_no": "AD164",
            "project_name": "Independent Subsea HVDC System Project (Project Lightning), UAE",
            "po_no": "5000802593",
            "country_of_origin": "KOREA",
            "pol_country": "KOREA",
            "pod_country": "U.A.E."
        },
        "ci_rider_items": [],
        "pl_rider_items": [],
    }

    out = make_4page_data_dicts(payload_example, cbm_decimals=3)
    # out["ci_p1"], out["pl_p1"], out["ci_rider_p2"], out["pl_rider_p2"] ready to pass into your builders
    print("Generated keys:", list(out.keys()))
