
from __future__ import annotations
from typing import List
import re

PHASE_RE = re.compile(r"\b(?:PHASE|PH|STAGE)[- ]?(\d{1,2})\b", re.IGNORECASE)

def extract_phases(text: str) -> List[str]:
    return [f"PHASE-{m.group(1)}" for m in PHASE_RE.finditer(text)]
