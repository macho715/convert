#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Outlook HVDC Analyzer — v2.0 (Python-first · Tidy-First ready)

목적
- Outlook PST 스캔 산출물(엑셀)을 읽어 HVDC 문맥(케이스/사이트/LPO/단계)을 자동 추출·분류
- 결과를 다중 시트 엑셀로 저장(analysis, summary_by_stage, summary_by_site)
- 구조/행위 분리 원칙(Tidy First): parsing 규칙/정규식/룰셋은 별도 상수로 분리

입력
- XLSX/XLSM 파일(기본: '전체_이메일' 시트). 컬럼명은 유연 매칭(헤더 노멀라이즈)
  최소 권장 컬럼: Subject, Body, From, To, Date

사용
  python outlook_hvdc_analyzer_v2.py --input OUTLOOK_202510.xlsx --sheet "전체_이메일" --output report.xlsx
옵션
  --stage-rules rules.json|yaml (키워드→단계 맵)
  --site-alias  site_alias.json|yaml (사이트 별칭→정규명)
"""

from __future__ import annotations

import argparse
import sys
import json
import re
import logging
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple

import pandas as pd

try:
    import yaml  # optional
except Exception:  # pragma: no cover
    yaml = None

# -------------------------
# Logging
# -------------------------
LOG = logging.getLogger("hvdc.outlook.analyzer")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)

# -------------------------
# Header normalization
# -------------------------

HEADER_ALIASES = {
    # normalized : {possible variants...}
    "subject": {"subject", "제목", "subj", "메일제목"},
    "body": {"body", "본문", "내용", "메일내용"},
    "from": {"from", "sender", "발신자", "보낸사람"},
    "to": {"to", "수신자", "받는사람"},
    "cc": {"cc", "참조"},
    "date": {"date", "sent", "날짜", "보낸날짜", "sent on"},
    "attachments": {"attachments", "attachment", "첨부", "첨부파일"},
}

def normalize_header(h: str) -> str:
    s = str(h).strip().lower()
    s = re.sub(r"[\s\u3000]+", " ", s)     # spaces incl. full-width
    s = s.replace("_", " ").replace("-", " ")
    s = re.sub(r"\s+", " ", s)
    # alias map
    for norm, variants in HEADER_ALIASES.items():
        if s in variants:
            return norm
    return s

def normalize_headers(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [normalize_header(c) for c in df.columns]
    return df

# -------------------------
# Rules & Regex
# -------------------------

# HVDC Case patterns
CASE_REGEXES = [
    # HVDC-ADOPT-XXX-XXXX (strict)
    r"(HVDC[-\s]?ADOPT[-\s]?[A-Z0-9]+[-\s]?[A-Z0-9-]+)",
    # Parentheses short form e.g., (HE-1234) -> HVDC-ADOPT-HE-1234
    r"\(([A-Z]{2,5})-([A-Z0-9-]{2,})\)",
    # Loose form: ADOPT HE 1234-56
    r"ADOPT[\s:_-]*([A-Z]{2,5})[\s:_-]*([A-Z0-9-]{2,})",
]

SITE_ALIASES_DEFAULT = {
    "DAS": {"das"},
    "AGI": {"agi", "al ghallan", "al ghallan island", "ghallan", "al ghallan is.", "al ghallan isl"},
    "MIR": {"mir"},
    "MIRFA": {"mirfa"},
    "GHALLAN": {"ghallan", "al ghallan", "al ghallan island"},
}

# LPO / PO regex
LPO_REGEXES = [
    r"\bLPO[:\s\-]*([A-Z0-9\-\/]{4,})\b",
    r"\bP\.?O\.?[:\s\-]*([A-Z0-9\-\/]{4,})\b",
]

# Stage rules (keyword → stage)
STAGE_RULES_DEFAULT = {
    "procurement": [
        "rfq", "quotation", "quote", "proforma", "pro-forma", "pr ", "po ", "lpo", "supplier",
    ],
    "shipping": [
        "booking", "awb", "mawb", "hawb", "bl ", "b/l", "vessel", "eta", "etd", "sail", "voyage",
    ],
    "customs": [
        "boe", "import code", "customs", "vat", "duty", "declaration", "clearance", "inspection",
        "fanr", "moiat", "coo", "cert of origin", "gate pass",
    ],
    "logistics": [
        "warehouse", "delivery", "pickup", "trailer", "berth", "jetty", "lolo", "roro", "stow",
        "grn", "wms", "putaway", "picking",
    ],
    "installation": [
        "install", "erection", "fit up", "site work", "site installation",
    ],
    "testing": [
        "test ", "fat", "sat", "load test", "pressure test", "stability",
    ],
    "certification": [
        "certificate", "approval", "permit", "ms", "method statement", "coe", "sow",
    ],
}

def load_rules(path: Optional[str]) -> Dict[str, List[str]]:
    if not path:
        return STAGE_RULES_DEFAULT
    p = Path(path)
    if not p.exists():
        LOG.warning("Stage rules file not found: %s", path)
        return STAGE_RULES_DEFAULT
    if p.suffix.lower() in {".yaml", ".yml"} and yaml:
        return yaml.safe_load(p.read_text())
    return json.loads(p.read_text())

def load_site_alias(path: Optional[str]) -> Dict[str, set]:
    base = {k: set(v) for k, v in SITE_ALIASES_DEFAULT.items()}
    if not path:
        return base
    p = Path(path)
    if not p.exists():
        LOG.warning("Site alias file not found: %s", path)
        return base
    if p.suffix.lower() in {".yaml", ".yml"} and yaml:
        extra = yaml.safe_load(p.read_text())
    else:
        extra = json.loads(p.read_text())
    # merge
    for k, vals in extra.items():
        base.setdefault(k, set()).update({str(v).lower() for v in vals})
    return base

# -------------------------
# Extractors
# -------------------------

def _extract_cases(text: str) -> List[str]:
    if not text:
        return []
    t = str(text)
    out = []
    for rx in CASE_REGEXES:
        for m in re.finditer(rx, t, re.IGNORECASE):
            if m.lastindex and m.lastindex >= 2 and "(" not in m.group(0):
                vendor = m.group(1).upper()
                num = m.group(2).upper()
                out.append(f"HVDC-ADOPT-{vendor}-{num}")
            else:
                out.append(re.sub(r"\s+", "-", m.group(1).upper()))
    # normalize HVDC-ADOPT spacing variants
    out = [re.sub(r"HVDC[-\s]?ADOPT[-\s]?", "HVDC-ADOPT-", s, flags=re.IGNORECASE) for s in out]
    # dedup preserve order
    seen = set()
    dedup = []
    for s in out:
        if s not in seen:
            seen.add(s)
            dedup.append(s)
    return dedup

def _extract_sites(text: str, alias_map: Dict[str, set]) -> List[str]:
    if not text:
        return []
    t = str(text).lower()
    hits = []
    for site, aliases in alias_map.items():
        for a in aliases | {site.lower()}:
            if re.search(rf"\b{re.escape(a)}\b", t):
                hits.append(site)
                break
    # prefer deterministic order by project priority
    prio = ["AGI", "DAS", "MIR", "MIRFA", "GHALLAN"]
    ordered = [s for s in prio if s in hits] + [s for s in hits if s not in prio]
    return list(dict.fromkeys(ordered))

def _extract_lpos(text: str) -> List[str]:
    if not text:
        return []
    t = str(text)
    out = []
    for rx in LPO_REGEXES:
        out += [m.group(1).upper() for m in re.finditer(rx, t, re.IGNORECASE)]
    return list(dict.fromkeys(out))

def _classify_stage(subject: str, body: str, rules: Dict[str, List[str]]) -> Tuple[str, List[str]]:
    text = f"{subject or ''}\n{body or ''}".lower()
    hits = []
    chosen = "uncategorized"
    for stage, kws in rules.items():
        stage_hits = [kw for kw in kws if kw in text]
        if stage_hits and chosen == "uncategorized":
            chosen = stage
        hits.extend([f"{stage}:{kw}" for kw in stage_hits])
    return chosen, sorted(list(dict.fromkeys(hits)))

# -------------------------
# Core
# -------------------------

def auto_sheet_name(xl: pd.ExcelFile) -> str:
    """우선순위: '전체_이메일' → 'emails' → 첫 시트"""
    names = [s.lower() for s in xl.sheet_names]
    if "전체_이메일" in xl.sheet_names:
        return "전체_이메일"
    if "emails" in names:
        return xl.sheet_names[names.index("emails")]
    return xl.sheet_names[0]

def load_dataframe(path: str, sheet: Optional[str]) -> pd.DataFrame:
    xl = pd.ExcelFile(path, engine="openpyxl")
    if not sheet:
        sheet = auto_sheet_name(xl)
    df = pd.read_excel(xl, sheet_name=sheet, engine="openpyxl")
    df = normalize_headers(df)
    return df

def pick_col(df: pd.DataFrame, *candidates: str) -> Optional[str]:
    cols = set(df.columns)
    for c in candidates:
        if c in cols:
            return c
    return None

def analyze(df: pd.DataFrame, rules: Dict[str, List[str]], site_alias: Dict[str, set]) -> pd.DataFrame:
    # choose useful columns
    subj_col = pick_col(df, "subject")
    body_col = pick_col(df, "body")
    from_col = pick_col(df, "from")
    date_col = pick_col(df, "date")

    if not subj_col:
        raise ValueError("필수 헤더(Subject)를 찾을 수 없습니다. 헤더 별칭을 확인하세요.")

    out_rows = []
    for _, row in df.iterrows():
        subj = str(row.get(subj_col, "") or "")
        body = str(row.get(body_col, "") or "")
        sender = str(row.get(from_col, "") or "")
        date = row.get(date_col)

        cases = list(dict.fromkeys(_extract_cases(subj) + _extract_cases(body)))
        sites = list(dict.fromkeys(_extract_sites(subj, site_alias) + _extract_sites(body, site_alias)))
        lpos  = list(dict.fromkeys(_extract_lpos(subj) + _extract_lpos(body)))
        stage, stage_hits = _classify_stage(subj, body, rules)

        out_rows.append({
            "date": date,
            "from": sender,
            "subject": subj,
            "hvdc_cases": "; ".join(cases),
            "primary_case": cases[0] if cases else "",
            "sites": "; ".join(sites),
            "primary_site": sites[0] if sites else "",
            "lpo_numbers": "; ".join(lpos),
            "stage": stage,
            "stage_hits": "; ".join(stage_hits),
        })
    return pd.DataFrame(out_rows)

def summaries(df: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    s1 = (df.groupby("stage")["subject"].count()
            .sort_values(ascending=False)
            .rename("count")
            .reset_index())
    s2 = (df.groupby("primary_site")["subject"].count()
            .sort_values(ascending=False)
            .rename("count")
            .reset_index())
    return {"summary_by_stage": s1, "summary_by_site": s2}

def save_excel(analysis: pd.DataFrame, sums: Dict[str, pd.DataFrame], out_path: Path) -> Path:
    out_path = out_path.with_suffix(".xlsx")
    with pd.ExcelWriter(out_path, engine="xlsxwriter") as w:
        analysis.to_excel(w, sheet_name="analysis", index=False)
        for name, df in sums.items():
            df.to_excel(w, sheet_name=name, index=False)
    return out_path

# -------------------------
# CLI
# -------------------------

def build_argparser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(description="Outlook HVDC Analyzer v2.0")
    ap.add_argument("--input", "-i", required=True, help="Outlook PST scan Excel path (.xlsx/.xlsm)")
    ap.add_argument("--sheet", "-s", default=None, help="Sheet name (default: auto)")
    ap.add_argument("--output", "-o", default=None, help="Output excel path (default: OUTLOOK_HVDC_ANALYSIS_YYYYMMDD.xlsx)")
    ap.add_argument("--stage-rules", default=None, help="JSON/YAML stage rules mapping")
    ap.add_argument("--site-alias", default=None, help="JSON/YAML site alias mapping")
    return ap

def main(argv: Optional[List[str]] = None) -> int:
    args = build_argparser().parse_args(argv)

    in_path = Path(args.input)
    if not in_path.exists():
        LOG.error("입력 파일을 찾을 수 없습니다: %s", in_path)
        return 2

    try:
        rules = load_rules(args.stage_rules)
        sites = load_site_alias(args.site_alias)
        df_raw = load_dataframe(str(in_path), args.sheet)
        df_ana = analyze(df_raw, rules, sites)
        sums = summaries(df_ana)

        if args.output:
            out_path = Path(args.output)
        else:
            ts = datetime.now().strftime("%Y%m%d_%H%M")
            out_path = in_path.parent / f"OUTLOOK_HVDC_ANALYSIS_{ts}.xlsx"

        out_file = save_excel(df_ana, sums, out_path)
        LOG.info("완료: %s", out_file)
        return 0
    except Exception as e:
        LOG.exception("실패: %s", e)
        return 1

if __name__ == "__main__":
    raise SystemExit(main())
