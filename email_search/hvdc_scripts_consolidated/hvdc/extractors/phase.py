from __future__ import annotations
from typing import List
import re

# 예시: PHASE-1, PH1, STAGE-2 등(실제 규칙은 이후 확장)
PHASE_RE = re.compile(r"\b(?:PHASE|PH|STAGE)[- ]?(\d{1,2})\b", re.IGNORECASE)

def extract_phases(text: str) -> List[str]:
    return [f"PHASE-{m.group(1)}" for m in PHASE_RE.finditer(text)]
