
from __future__ import annotations
from typing import List
import re

LPO_RE = re.compile(r"\b(?:LPO|PO)\s*[-:]?\s*(\d{5,12})\b", re.IGNORECASE)

def extract_lpos(text: str) -> List[str]:
    vals = [m.group(1) for m in LPO_RE.finditer(text)]
    norm = [f"PO-{v}" for v in vals]
    seen = set()
    out: list[str] = []
    for v in norm:
        if v not in seen:
            seen.add(v)
            out.append(v)
    return out
