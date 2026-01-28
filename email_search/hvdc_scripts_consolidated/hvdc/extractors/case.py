from __future__ import annotations
from typing import List
from ..core.regex import COMPILED
from ..core.typing import CaseHit

def extract_cases(text: str) -> List[CaseHit]:
    """제목/본문/폴더명에서 케이스류 식별자를 추출(중복 제거, 안정 정렬)."""
    hits: list[CaseHit] = []

    def add(kind: str, value: str, span: tuple[int, int] | None):
        hits.append({"value": value, "kind": kind, "span": span})

    # HVDC ADOPT
    for m in COMPILED["HVDC_ADOPT"].finditer(text):
        add("HVDC_ADOPT", m.group(0), m.span())

    # 일반 HVDC
    for m in COMPILED["HVDC_GENERIC"].finditer(text):
        val = m.group(0)
        # ADOPT와 중복되는 경우는 스킵
        if not val.upper().startswith("HVDC-ADOPT-"):
            add("HVDC_GENERIC", val, m.span())

    # PRL
    for m in COMPILED["PRL"].finditer(text):
        add("PRL", m.group(0), m.span())

    # JPTW / GRM (페어)
    for m in COMPILED["JPTW_GRM"].finditer(text):
        add("JPTW", f"JPTW-{m.group(1)}", m.span(1))
        add("GRM",  f"GRM-{m.group(2)}", m.span(2))

    # 중복 제거(값+kind 기준) & 입력 순서 안정
    seen = set()
    dedup: list[CaseHit] = []
    for h in hits:
        key = (h["kind"].upper(), h["value"].upper())
        if key in seen:
            continue
        seen.add(key)
        dedup.append(h)

    return dedup
