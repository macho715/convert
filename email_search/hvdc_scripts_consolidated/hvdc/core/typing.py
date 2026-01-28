from __future__ import annotations
from typing import TypedDict, NotRequired, List

class CaseHit(TypedDict):
    value: str
    kind: str        # "HVDC_ADOPT" | "HVDC_GENERIC" | "PRL" | "JPTW" | "GRM"
    span: NotRequired[tuple[int, int]]

class ParseResult(TypedDict):
    cases: List[CaseHit]
    sites: List[str]
    lpos: List[str]
    phases: List[str]
