# -*- coding: utf-8 -*-
"""CIPL commons mapper (Single Source of Truth)

- Input supports:
  (A) commons + static_parties (+ optional ci_rider_items / pl_rider_items)
  (B) already-expanded 4-page dicts (ci_p1, ci_rider_p2, pl_p1, pl_rider_p2) -> pass-through

Core:
- make_4page_data_dicts(payload, cbm_decimals=3) -> dicts for all 4 pages
- If rider items are empty, auto-generate minimal 1-line rider items from commons.
"""

from __future__ import annotations

from copy import deepcopy
from typing import Any, Dict, List


def _require(d: Dict[str, Any], key: str) -> Any:
    if key not in d:
        raise KeyError(f"Missing required key: {key}")
    return d[key]


def _fmt_usd(x: float) -> str:
    return f"USD {x:,.2f}"


def _fmt_kg(x: float) -> str:
    return f"{x:,.2f} KGS"


def _fmt_cbm(x: float, decimals: int = 3) -> str:
    return f"{x:.{decimals}f} CBM"


def _fmt_pkgs(n: int) -> str:
    return f"{n} PKGS"


def _safe_float(x: Any, default: float = 0.0) -> float:
    try:
        if x is None or x == "":
            return default
        return float(x)
    except Exception:
        return default


def _safe_int(x: Any, default: int = 0) -> int:
    try:
        if x is None or x == "":
            return default
        return int(float(x))
    except Exception:
        return default


def _infer_fob(total_cif: float, freight: float, insurance: float) -> float:
    # FOB = CIF - freight - insurance (non-negative guard)
    fob = round(total_cif - freight - insurance, 2)
    return fob if fob >= 0 else 0.0


def _infer_ci_unit_price(commons: Dict[str, Any]) -> float:
    # Priority: explicit unit price -> else total_cif / ci_qty (or quantity_pkgs fallback)
    if "unit_price_usd_per_unit" in commons and commons["unit_price_usd_per_unit"] not in (None, ""):
        return round(float(commons["unit_price_usd_per_unit"]), 2)

    total = _safe_float(_require(commons, "total_cif_usd"), 0.0)
    qty = _safe_float(commons.get("ci_qty", commons.get("quantity_pkgs", 0)), 0.0)
    if qty == 0:
        return 0.0
    return round(total / qty, 2)


def _default_if_missing_items(items: List[Dict[str, Any]], commons: Dict[str, Any]) -> List[Dict[str, Any]]:
    out = []
    for it in items:
        x = deepcopy(it)
        if not x.get("hs_code"):
            x["hs_code"] = commons.get("hs_code", "")
        if "origin" in x and (x.get("origin") in (None, "")):
            x["origin"] = commons.get("country_of_origin", "KOREA")
        out.append(x)
    return out


def _auto_ci_rider_items(commons: Dict[str, Any]) -> List[Dict[str, Any]]:
    item_desc = str(_require(commons, "item_description"))
    hs = str(_require(commons, "hs_code"))
    origin = str(commons.get("country_of_origin", "KOREA"))

    qty = _safe_float(commons.get("ci_qty", commons.get("quantity_pkgs", 0)), 0.0)
    unit = str(commons.get("ci_unit", "M"))

    desc = str(commons.get("ci_description", item_desc))

    unit_price = _infer_ci_unit_price(commons)
    total_price = round(unit_price * qty, 2)

    # If OF text isn't meaningful for CI rider, still provide a stable default.
    of_text = str(commons.get("ci_of_text", "")) or f"OF {_safe_int(commons.get('quantity_pkgs', 1), 1)}"

    return [{
        "marks_no": _safe_int(commons.get("ci_marks_no", 1), 1),
        "of_text": of_text,
        "item_no": str(commons.get("ci_item_no", "4.1")),
        "description": desc,
        "hs_code": hs,
        "origin": origin,
        "qty": qty,
        "unit": unit,
        "unit_price": unit_price,
        "total_price": total_price,
    }]


def _auto_pl_rider_items(commons: Dict[str, Any]) -> List[Dict[str, Any]]:
    item_desc = str(_require(commons, "item_description"))
    hs = str(_require(commons, "hs_code"))

    qty = _safe_float(commons.get("pl_qty", commons.get("quantity_pkgs", 0)), 0.0)
    unit = str(commons.get("pl_unit", "PKGS"))

    net_kg = _safe_float(_require(commons, "net_weight_kg"), 0.0)
    gross_kg = _safe_float(_require(commons, "gross_weight_kg"), 0.0)
    vol = _safe_float(_require(commons, "dimension_cbm"), 0.0)

    # Optional dimensions; if missing, keep 0 (script requirement: numeric)
    dim_l = _safe_float(commons.get("dim_l_cm", 0), 0.0)
    dim_w = _safe_float(commons.get("dim_w_cm", 0), 0.0)
    dim_h = _safe_float(commons.get("dim_h_cm", 0), 0.0)

    desc = str(commons.get("pl_description", item_desc))

    of_text = str(commons.get("pl_of_text", "")) or f"OF {_safe_int(commons.get('quantity_pkgs', 1), 1)}"

    return [{
        "marks_no": _safe_int(commons.get("pl_marks_no", 1), 1),
        "of_text": of_text,
        "item_no": str(commons.get("pl_item_no", "4.1")),
        "item_location": str(commons.get("item_location", "AGI")),
        "packing_style": str(commons.get("packing_style", "SKID")),
        "description": desc,
        "hs_code": hs,
        "qty": qty,
        "unit": unit,
        "net_kg": net_kg,
        "gross_kg": gross_kg,
        "dim_l_cm": dim_l,
        "dim_w_cm": dim_w,
        "dim_h_cm": dim_h,
        "volume_cbm": vol,
    }]


def _ensure_rider_items(items: List[Dict[str, Any]], commons: Dict[str, Any], kind: str) -> List[Dict[str, Any]]:
    if items and len(items) > 0:
        return _default_if_missing_items(items, commons)
    return _auto_ci_rider_items(commons) if kind == "ci" else _auto_pl_rider_items(commons)


def make_4page_data_dicts(payload: Dict[str, Any], *, cbm_decimals: int = 3) -> Dict[str, Dict[str, Any]]:
    """Main entry.
    If payload already has 4-page dicts, returns them (pass-through).
    Else expects commons + static_parties and builds dicts for all pages.
    """

    if all(k in payload for k in ("ci_p1", "ci_rider_p2", "pl_p1", "pl_rider_p2")):
        return {
            "ci_p1": payload["ci_p1"],
            "ci_rider_p2": payload["ci_rider_p2"],
            "pl_p1": payload["pl_p1"],
            "pl_rider_p2": payload["pl_rider_p2"],
        }

    commons = _require(payload, "commons")
    parties = _require(payload, "static_parties")

    # required commons
    pol = str(_require(commons, "pol"))
    pod = str(_require(commons, "pod"))
    carrier = str(_require(commons, "carrier"))
    sailing_on = str(_require(commons, "sailing_on"))
    invoice_no = str(_require(commons, "invoice_no"))
    invoice_date = str(_require(commons, "invoice_date"))
    hs_code = str(_require(commons, "hs_code"))

    mfg_name = str(_require(commons, "manufacturer_name"))
    mfg_a1 = str(_require(commons, "manufacturer_address_1"))
    mfg_a2 = str(_require(commons, "manufacturer_address_2"))

    item_desc = str(_require(commons, "item_description"))
    package_no = str(_require(commons, "package_no"))
    net_kg = _safe_float(_require(commons, "net_weight_kg"), 0.0)
    gross_kg = _safe_float(_require(commons, "gross_weight_kg"), 0.0)
    dim_cbm = _safe_float(_require(commons, "dimension_cbm"), 0.0)

    item = str(_require(commons, "item"))
    qty_pkgs = _safe_int(_require(commons, "quantity_pkgs"), 0)

    freight = _safe_float(_require(commons, "freight_usd"), 0.0)
    insurance = _safe_float(_require(commons, "insurance_usd"), 0.0)
    total_cif = _safe_float(_require(commons, "total_cif_usd"), 0.0)

    fob = _infer_fob(total_cif, freight, insurance)

    # required parties
    shipper_name = str(_require(parties, "shipper_name"))
    shipper_addr1 = str(_require(parties, "shipper_addr1"))
    shipper_addr2 = str(_require(parties, "shipper_addr2"))

    consignee_name1 = str(_require(parties, "consignee_name1"))
    consignee_name2 = str(_require(parties, "consignee_name2"))
    consignee_addr1 = str(_require(parties, "consignee_addr1"))
    consignee_addr2 = str(_require(parties, "consignee_addr2"))
    consignee_addr3 = str(_require(parties, "consignee_addr3"))

    notify_name = str(_require(parties, "notify_name"))
    notify_addr1 = str(_require(parties, "notify_addr1"))
    notify_addr2 = str(_require(parties, "notify_addr2"))

    project_no = str(_require(parties, "project_no"))
    project_name = str(_require(parties, "project_name"))
    po_no = str(_require(parties, "po_no"))
    country_of_origin = str(_require(parties, "country_of_origin"))

    pol_country = str(parties.get("pol_country", "KOREA"))
    pod_country = str(parties.get("pod_country", "U.A.E."))

    ci_p1 = {
        "page_no": "PAGE NO.: 1 OF 1",
        "title": "COMMERCIAL INVOICE",  # keep original spelling in template

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

        "pol": pol,
        "pol_country": pol_country,
        "pod": pod,
        "pod_country": pod_country,

        "carrier": carrier,
        "sailing_on": sailing_on,

        "invoice_no": invoice_no,
        "invoice_date": invoice_date,

        "hs_code": hs_code,
        "project_no": project_no,
        "po_no": po_no,

        "mfg_name": mfg_name,
        "mfg_addr1": mfg_a1,
        "mfg_addr2": mfg_a2,

        "country_of_origin": country_of_origin,

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
        "box_dimension": _fmt_cbm(dim_cbm, decimals=cbm_decimals),
        "box_origin": country_of_origin,

        "sum_item": item,
        "sum_qty": _fmt_pkgs(qty_pkgs),
        "sum_fob": _fmt_usd(fob),
        "sum_freight": _fmt_usd(freight),
        "sum_insurance": _fmt_usd(insurance),
        "sum_cif": _fmt_usd(total_cif),
    }

    pl_p1 = {
        "page_no": "PAGE NO.: 1 OF 1",
        "title": "PACKING LIST",  # keep template

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

        "pol": pol,
        "pol_country": pol_country,
        "pod": pod,
        "pod_country": pod_country,

        "carrier": carrier,
        "sailing_on": sailing_on,

        "invoice_no": invoice_no,
        "invoice_date": invoice_date,

        "hs_code": hs_code,
        "project_no": project_no,
        "po_no": po_no,

        "mfg_name": mfg_name,
        "mfg_addr1": mfg_a1,
        "mfg_addr2": mfg_a2,

        "country_of_origin": country_of_origin,

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

        "line_goods": item_desc,
        "line_qty": str(qty_pkgs),
        "line_qty_unit": "PKGS",
        "total_pkg_line": f"TOTAL {qty_pkgs} PACKAGES",
        "total_net": _fmt_kg(net_kg),
        "total_gross": _fmt_kg(gross_kg),
        "total_measure": _fmt_cbm(dim_cbm, decimals=cbm_decimals),
        "details_note": "(DETAILS ARE AS PER ATTACHED SHEETS)",
    }

    ci_rider_items = _ensure_rider_items(payload.get("ci_rider_items", []), commons, "ci")
    pl_rider_items = _ensure_rider_items(payload.get("pl_rider_items", []), commons, "pl")

    ci_total_items = round(sum(_safe_float(x.get("total_price", 0.0), 0.0) for x in ci_rider_items), 2)
    total_fob_charge = ci_total_items if ci_total_items > 0 else fob

    ci_rider_p2 = {
        "sheet_name": "CI_Rider_P2",
        "title": "COMMERCIAL INVOICE ATTACHED RIDER",
        "section_title": item_desc,
        "bottom_terms": "CIF ABU DHABI PORT",
        "bottom_unit": str(commons.get("ci_unit", "M")),
        "meta": {"page_no": "PAGE NO.: 2 OF 2"},
        "items": ci_rider_items,
        "totals": {
            "total_fob": float(total_fob_charge),
            "freight": "" if freight == 0.0 else float(freight),
            "insurance": "" if insurance == 0.0 else float(insurance),
        },
    }

    pl_rider_p2 = {
        "page_no": "PAGE NO.: 2 OF 2",
        "title": "PACKING LIST ATTACHED RIDER",
        "section_title": item_desc,
        "items": pl_rider_items,
    }

    return {
        "ci_p1": ci_p1,
        "ci_rider_p2": ci_rider_p2,
        "pl_p1": pl_p1,
        "pl_rider_p2": pl_rider_p2,
    }
