
from __future__ import annotations
from typing import List
import re

SITE_TOKENS = ["AGI", "DAS", "MIRFA", "SHU", "ZAK"]

def extract_sites(text: str) -> List[str]:
    found = []
    upper = text.upper()
    for tok in SITE_TOKENS:
        if re.search(rf"\b{tok}\b", upper):
            found.append(tok)
    return sorted(set(found), key=found.index)
